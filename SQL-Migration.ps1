#Requires -Version 5.1
# =============================================================================
# SQL-Migration.ps1
# Hauptskript mit WinForms-GUI fuer SQL Server Migrationen
#
# Betriebsmodi (Option C):
#   - Beim Start: Rollenauswahl (Quelle / Ziel / Automatisch)
#   - Automatisch: Zustandsdatei im Exchange-Pfad vorhanden -> Ziel-Modus
#                  sonst -> Quell-Modus
#   - GUI zeigt nur die relevante Seite (Quelle ODER Ziel)
#
# Szenario-Erkennung (Option B):
#   - Direct:   Zielserver TCP-erreichbar -> versuche UNC direkt;
#               bei Fehler automatisch lokal+Copy
#   - TwoPhase: Zielserver nicht erreichbar -> nur Phase1 (Quelle)
#               oder nur Phase2 (Ziel)
# =============================================================================

[CmdletBinding()]
param(
    [string]$ConfigFile = "$PSScriptRoot\config\migration.config.json",
    # Rollenvorgabe per Parameter moeglich (uebersteuert Dialog)
    [ValidateSet('','Source','Target','Auto')]
    [string]$Role = ''
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ---------------------------------------------------------------------------
# Assemblies laden
# ---------------------------------------------------------------------------
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# ---------------------------------------------------------------------------
# Module laden
# ---------------------------------------------------------------------------
$modulePath = Join-Path $PSScriptRoot 'modules'
@(
    'Write-MigrationLog',
    'Connect-SqlServer',
    'Get-SqlObjects',
    'Invoke-MigrationState',
    'Invoke-Migration'
) | ForEach-Object {
    $mp = Join-Path $modulePath "$_.psm1"
    if (Test-Path $mp) {
        Import-Module $mp -Force -DisableNameChecking -ErrorAction Stop
    } else {
        throw "Modul nicht gefunden: $mp"
    }
}

# dbaTools pruefen
if (-not (Get-Module -ListAvailable -Name dbaTools)) {
    [System.Windows.Forms.MessageBox]::Show(
        "dbaTools ist nicht installiert.`nBitte installieren: Install-Module dbaTools",
        'Fehlende Abhaengigkeit', 'OK', 'Error') | Out-Null
    exit 1
}
Import-Module dbaTools -ErrorAction Stop

# ---------------------------------------------------------------------------
# Konfiguration laden
# ---------------------------------------------------------------------------
$Config = @{
    DefaultExchangePath     = "\\exchange-server\SQLMigration\Backups"
    DefaultLocalBackupPath  = "F:\Daten\SQL\Backup"
    DefaultLogPath          = "C:\SQLMigration\Logs"
    LogFilePrefix           = 'SQL-Migration'
    ConnectionTimeout       = 30
    BackupCompression       = $true
    VerifyBackup            = $true
    CopyOnlyBackup          = $true
    DefaultMigrationMethod  = 'BackupRestore'
    UncAccessTestTimeoutSec = 10
    StateFileName           = '_migration_state.json'
    TrustServerCertificate  = $true
}
if (Test-Path $ConfigFile) {
    try {
        $jsonCfg = Get-Content $ConfigFile -Raw | ConvertFrom-Json
        foreach ($prop in $jsonCfg.PSObject.Properties) {
            $Config[$prop.Name] = $prop.Value
        }
    } catch { <# Defaults beibehalten #> }
}

# Logverzeichnis anlegen
if (-not (Test-Path $Config.DefaultLogPath)) {
    New-Item -ItemType Directory -Path $Config.DefaultLogPath -Force | Out-Null
}
$null = Initialize-MigrationLog -LogDirectory $Config.DefaultLogPath `
                                -Prefix $Config.LogFilePrefix

# ---------------------------------------------------------------------------
# Globale Zustandsvariablen
# ---------------------------------------------------------------------------
$script:ActiveServer   = $null    # Verbundener Server (Quelle oder Ziel)
$script:ActiveAuth     = 'Windows'
$script:ExchangePath   = $Config.DefaultExchangePath
$script:LocalBackupPath= $Config.DefaultLocalBackupPath
$script:ActiveRole     = ''       # 'Source' | 'Target'
$script:Scenario       = ''       # 'Direct' | 'TwoPhase'
$script:LoadedState    = $null    # Zustandsdatei (im Ziel-Modus)

# ===========================================================================
# FARBEN & SCHRIFTEN
# ===========================================================================
$clrBg        = [System.Drawing.Color]::FromArgb(24, 26, 32)
$clrPanel     = [System.Drawing.Color]::FromArgb(33, 37, 43)
$clrBorder    = [System.Drawing.Color]::FromArgb(52, 58, 70)
$clrAccent    = [System.Drawing.Color]::FromArgb(0, 120, 215)
$clrAccentGrn = [System.Drawing.Color]::FromArgb(40, 167, 69)
$clrAccentAmb = [System.Drawing.Color]::FromArgb(255, 140, 0)
$clrText      = [System.Drawing.Color]::FromArgb(220, 225, 235)
$clrSubText   = [System.Drawing.Color]::FromArgb(140, 150, 165)
$clrInput     = [System.Drawing.Color]::FromArgb(40, 44, 52)
$clrHeader    = [System.Drawing.Color]::FromArgb(15, 17, 22)

$fntTitle   = New-Object System.Drawing.Font('Consolas', 12, [System.Drawing.FontStyle]::Bold)
$fntNormal  = New-Object System.Drawing.Font('Segoe UI',  9)
$fntBold    = New-Object System.Drawing.Font('Segoe UI',  9, [System.Drawing.FontStyle]::Bold)
$fntSmall   = New-Object System.Drawing.Font('Segoe UI',  8)
$fntMono    = New-Object System.Drawing.Font('Consolas',  8)

# ===========================================================================
# HILFS-DIALOG: Rollenauswahl beim Start (Option C)
# ===========================================================================
function Show-RoleDialog {
    param([string]$ExchangePath, [string]$StateFileName)

    $stateFile   = Join-Path $ExchangePath $StateFileName
    $stateExists = Test-Path $stateFile

    $dlg              = New-Object System.Windows.Forms.Form
    $dlg.Text         = 'SQL Migration - Rollenauswahl'
    $dlg.Size         = New-Object System.Drawing.Size(480, 320)
    $dlg.StartPosition= 'CenterScreen'
    $dlg.BackColor    = $clrBg
    $dlg.ForeColor    = $clrText
    $dlg.Font         = $fntNormal
    $dlg.FormBorderStyle = 'FixedDialog'
    $dlg.MaximizeBox  = $false

    $lbl           = New-Object System.Windows.Forms.Label
    $lbl.Text      = 'Welche Rolle uebernimmt dieses Script auf diesem Rechner?'
    $lbl.Font      = $fntBold
    $lbl.ForeColor = $clrText
    $lbl.AutoSize  = $false
    $lbl.Size      = New-Object System.Drawing.Size(440, 40)
    $lbl.Location  = New-Object System.Drawing.Point(16, 16)
    $dlg.Controls.Add($lbl)

    $lblState          = New-Object System.Windows.Forms.Label
    $lblState.Font     = $fntSmall
    $lblState.AutoSize = $false
    $lblState.Size     = New-Object System.Drawing.Size(440, 20)
    $lblState.Location = New-Object System.Drawing.Point(16, 58)
    if ($stateExists) {
        $lblState.Text     = '[OK] Zustandsdatei gefunden im Exchange-Pfad - Ziel-Modus empfohlen'
        $lblState.ForeColor= [System.Drawing.Color]::FromArgb(40,167,69)
    } else {
        $lblState.Text     = '[ ] Keine Zustandsdatei gefunden - Quell-Modus empfohlen'
        $lblState.ForeColor= $clrSubText
    }
    $dlg.Controls.Add($lblState)

    # -----------------------------------------------------------------------
    # Buttons ohne Closures:
    #   - DialogResult = OK signalisiert dass ein Button gedrueckt wurde
    #   - Tag des Forms traegt den gewaehlten Wert
    # -----------------------------------------------------------------------
    $btnSource = New-Object System.Windows.Forms.Button
    $btnSource.Size     = New-Object System.Drawing.Size(430, 46)
    $btnSource.Location = New-Object System.Drawing.Point(16, 88)
    $btnSource.BackColor= [System.Drawing.Color]::FromArgb(40,44,52)
    $btnSource.ForeColor= $clrText
    $btnSource.FlatStyle= 'Flat'
    $btnSource.FlatAppearance.BorderColor = $clrAccent
    $btnSource.FlatAppearance.BorderSize  = 2
    $btnSource.Text     = ">>  QUELL-Server (Phase 1)`r`nBackup / Detach + Copy auf Exchange-Pfad. Zustandsdatei wird erstellt."
    $btnSource.TextAlign= 'MiddleLeft'
    $btnSource.Padding  = New-Object System.Windows.Forms.Padding(10,0,0,0)
    $btnSource.DialogResult = [System.Windows.Forms.DialogResult]::Yes   # Yes = Source
    $dlg.Controls.Add($btnSource)

    $btnTarget = New-Object System.Windows.Forms.Button
    $btnTarget.Size     = New-Object System.Drawing.Size(430, 46)
    $btnTarget.Location = New-Object System.Drawing.Point(16, 144)
    $btnTarget.BackColor= [System.Drawing.Color]::FromArgb(40,44,52)
    $btnTarget.ForeColor= $clrText
    $btnTarget.FlatStyle= 'Flat'
    $btnTarget.FlatAppearance.BorderColor = $clrAccentGrn
    $btnTarget.FlatAppearance.BorderSize  = 2
    $btnTarget.Text     = "<<  ZIEL-Server (Phase 2)`r`nCopy vom Exchange-Pfad + Restore / Attach. Liest Zustandsdatei."
    $btnTarget.TextAlign= 'MiddleLeft'
    $btnTarget.Padding  = New-Object System.Windows.Forms.Padding(10,0,0,0)
    $btnTarget.DialogResult = [System.Windows.Forms.DialogResult]::No    # No = Target
    $dlg.Controls.Add($btnTarget)

    $btnAuto = New-Object System.Windows.Forms.Button
    $btnAuto.Size     = New-Object System.Drawing.Size(430, 46)
    $btnAuto.Location = New-Object System.Drawing.Point(16, 200)
    $btnAuto.BackColor= [System.Drawing.Color]::FromArgb(40,44,52)
    $btnAuto.ForeColor= $clrText
    $btnAuto.FlatStyle= 'Flat'
    $btnAuto.FlatAppearance.BorderColor = $clrAccentAmb
    $btnAuto.FlatAppearance.BorderSize  = 2
    $btnAuto.Text     = "**  AUTOMATISCH erkennen`r`nZustandsdatei vorhanden -> Ziel; sonst -> Quelle."
    $btnAuto.TextAlign= 'MiddleLeft'
    $btnAuto.Padding  = New-Object System.Windows.Forms.Padding(10,0,0,0)
    $btnAuto.DialogResult = [System.Windows.Forms.DialogResult]::Retry   # Retry = Auto
    $dlg.Controls.Add($btnAuto)

    $dr = $dlg.ShowDialog()
    $dlg.Dispose()

    switch ($dr) {
        ([System.Windows.Forms.DialogResult]::Yes)   { return 'Source' }
        ([System.Windows.Forms.DialogResult]::No)    { return 'Target' }
        ([System.Windows.Forms.DialogResult]::Retry) { return 'Auto'   }
        default                                       { return ''       }
    }
}

# ===========================================================================
# STATUSLEISTE (thread-safe)
# ===========================================================================
function Update-StatusBar {
    param([string]$Text, [string]$Color = 'Info')
    $c = switch ($Color) {
        'Error'   { [System.Drawing.Color]::FromArgb(220,53,69)  }
        'Success' { [System.Drawing.Color]::FromArgb(40,167,69)  }
        'Warn'    { [System.Drawing.Color]::FromArgb(255,193,7)  }
        default   { [System.Drawing.Color]::FromArgb(23,162,184) }
    }
    if ($script:statusLabel -and $script:statusLabel.IsHandleCreated) {
        $script:statusLabel.Invoke([Action]{
            $script:statusLabel.Text      = $Text
            $script:statusLabel.BackColor = $c
        })
    }
}

# ===========================================================================
# HILFS: Scrollbarer Fehler-Dialog (zeigt komplette Exception-Kette)
# ===========================================================================
function Show-ErrorDialog {
    param(
        [string]$Title   = 'Fehler',
        [string]$Message
    )
    $dlgErr              = New-Object System.Windows.Forms.Form
    $dlgErr.Text         = $Title
    $dlgErr.Size         = New-Object System.Drawing.Size(700, 450)
    $dlgErr.StartPosition= 'CenterScreen'
    $dlgErr.BackColor    = $clrBg
    $dlgErr.ForeColor    = $clrText
    $dlgErr.Font         = $fntSmall
    $dlgErr.FormBorderStyle = 'Sizable'

    $lblErrHint          = New-Object System.Windows.Forms.Label
    $lblErrHint.Text     = 'Vollstaendige Fehlermeldung (kann kopiert werden):'
    $lblErrHint.Font     = $fntBold
    $lblErrHint.ForeColor= [System.Drawing.Color]::FromArgb(220,53,69)
    $lblErrHint.AutoSize = $true
    $lblErrHint.Location = New-Object System.Drawing.Point(10, 10)
    $dlgErr.Controls.Add($lblErrHint)

    $txtErr              = New-Object System.Windows.Forms.TextBox
    $txtErr.Multiline    = $true
    $txtErr.ScrollBars   = 'Both'
    $txtErr.ReadOnly     = $true
    $txtErr.WordWrap     = $false
    $txtErr.Font         = $fntMono
    $txtErr.BackColor    = $clrInput
    $txtErr.ForeColor    = [System.Drawing.Color]::FromArgb(255,120,120)
    $txtErr.BorderStyle  = 'None'
    $txtErr.Text         = $Message
    $txtErr.Size         = New-Object System.Drawing.Size(670, 340)
    $txtErr.Location     = New-Object System.Drawing.Point(10, 35)
    $dlgErr.Controls.Add($txtErr)

    $btnClose            = New-Object System.Windows.Forms.Button
    $btnClose.Text       = 'Schliessen'
    $btnClose.Font       = $fntBold
    $btnClose.Size       = New-Object System.Drawing.Size(120, 28)
    $btnClose.Location   = New-Object System.Drawing.Point(10, 385)
    $btnClose.BackColor  = [System.Drawing.Color]::FromArgb(60,63,70)
    $btnClose.ForeColor  = $clrText
    $btnClose.FlatStyle  = 'Flat'
    $btnClose.FlatAppearance.BorderSize = 0
    $btnClose.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $dlgErr.Controls.Add($btnClose)

    $dlgErr.ShowDialog() | Out-Null
    $dlgErr.Dispose()
}

# Hilfsfunktion: vollstaendige Exception-Kette als Text
function Get-ExceptionDetail {
    param([System.Exception]$Ex)
    $lines = [System.Collections.Generic.List[string]]::new()
    $depth = 0
    $current = $Ex
    while ($current -ne $null) {
        $indent = '  ' * $depth
        $lines.Add("${indent}[$($current.GetType().Name)]")
        $lines.Add("${indent}$($current.Message)")
        if ($current.StackTrace) {
            $st = ($current.StackTrace -split "`n") | Select-Object -First 5
            foreach ($line in $st) {
                $lines.Add("${indent}  $($line.Trim())")
            }
        }
        $lines.Add('')
        $current = $current.InnerException
        $depth++
    }
    return $lines -join "`r`n"
}

# ===========================================================================
# HILFS: ListView befuellen
# ===========================================================================
function Set-ListViewData {
    param(
        [System.Windows.Forms.ListView]$ListView,
        [object[]]$Data,
        [string[]]$Properties
    )
    $ListView.BeginUpdate()
    $ListView.Items.Clear()
    if ($Data) {
        foreach ($row in $Data) {
            $vals = $Properties | ForEach-Object {
                $v = $row.$_; if ($null -eq $v) { '' } else { $v.ToString() }
            }
            $item = New-Object System.Windows.Forms.ListViewItem($vals[0])
            $item.Checked = $true   # standardmaessig alle markiert
            for ($i = 1; $i -lt $vals.Count; $i++) {
                $null = $item.SubItems.Add($vals[$i])
            }
            $ListView.Items.Add($item) | Out-Null
        }
    }
    $ListView.EndUpdate()
    # Zaehler-Label aktualisieren (Tag = Label-Referenz)
    if ($ListView.Tag -and $ListView.Tag -is [System.Windows.Forms.Label]) {
        $ListView.Tag.Text = "$($ListView.CheckedItems.Count) / $($ListView.Items.Count) ausgewaehlt"
    }
}

# ===========================================================================
# HINTERGRUND-JOB (Runspace, GUI friert nicht ein)
# ===========================================================================
# Job-Kontexte: Hashtable mit eindeutigem Schluessel pro Job.
# Benoetigt weil PS 5.1 ISE lokale Variablen in Timer-Add_Tick-Handlern
# nicht zuverlaessig bereitstellt. $script:-Scope ist immer erreichbar.
$script:_JobContexts = @{}
$script:_JobCounter  = 0

function Invoke-BackgroundJob {
    param(
        [Parameter(Mandatory)][scriptblock]$ScriptBlock,
        [object[]]$ArgumentList = @(),
        [System.Action[object]]$OnComplete,
        [System.Action[string]]$OnError
    )

    # Eindeutiger Schluessel fuer diesen Job
    $script:_JobCounter++
    $jobKey = "Job_$($script:_JobCounter)"

    $pool = [runspacefactory]::CreateRunspacePool(1, 1)
    $pool.ApartmentState = 'STA'
    $pool.Open()

    $ps = [powershell]::Create()
    $ps.RunspacePool = $pool

    $initScript = [scriptblock]::Create(@"
Set-StrictMode -Off
`$ErrorActionPreference = 'Continue'
Import-Module dbaTools -ErrorAction SilentlyContinue
Import-Module '$modulePath\Write-MigrationLog.psm1'     -Force
Import-Module '$modulePath\Connect-SqlServer.psm1'       -Force
Import-Module '$modulePath\Get-SqlObjects.psm1'          -Force
Import-Module '$modulePath\Invoke-MigrationState.psm1'   -Force
Import-Module '$modulePath\Invoke-Migration.psm1'        -Force
"@)
    $null = $ps.AddScript($initScript).Invoke()
    $ps.Commands.Clear()

    $null = $ps.AddScript($ScriptBlock)
    foreach ($a in $ArgumentList) { $null = $ps.AddArgument($a) }

    $handle = $ps.BeginInvoke()

    # Kontext im script:-Scope speichern - vom Timer-Handler abrufbar
    $script:_JobContexts[$jobKey] = @{
        Ps         = $ps
        Pool       = $pool
        Handle     = $handle
        OnComplete = $OnComplete
        OnError    = $OnError
    }

    $timer          = New-Object System.Windows.Forms.Timer
    $timer.Interval = 250
    $timer.Tag      = $jobKey   # Schluessel am Timer-Objekt selbst speichern

    $timer.Add_Tick({
        # $this = der Timer der gerade feuert (PS 5.1 stellt $this in Event-Handlern bereit)
        $key = $this.Tag
        $ctx = $script:_JobContexts[$key]
        if (-not $ctx) { $this.Stop(); $this.Dispose(); return }

        if ($ctx.Handle.IsCompleted) {
            $this.Stop()
            $this.Dispose()
            $script:_JobContexts.Remove($key)
            try {
                $result = $ctx.Ps.EndInvoke($ctx.Handle)
                # Streams.Error zusaetzlich pruefen (non-terminating errors)
                $streamErrors = @($ctx.Ps.Streams.Error)
                if ($ctx.Ps.HadErrors -and $streamErrors.Count -gt 0) {
                    $lines = [System.Collections.Generic.List[string]]::new()
                    foreach ($se in $streamErrors) {
                        $lines.Add($se.ToString())
                        if ($se.Exception -and $se.Exception.InnerException) {
                            $lines.Add('  Inner: ' + $se.Exception.InnerException.Message)
                        }
                    }
                    $errMsg = $lines -join "`r`n"
                    if ($ctx.OnError) { $ctx.OnError.Invoke($errMsg) }
                } else {
                    if ($ctx.OnComplete) { $ctx.OnComplete.Invoke($result) }
                }
            } catch {
                # EndInvoke wirft bei terminating errors im Runspace
                $detail = Get-ExceptionDetail -Ex $_.Exception
                # Streams nochmals pruefen fuer zusaetzlichen Kontext
                try {
                    $streamDetail = $ctx.Ps.Streams.Error |
                                    ForEach-Object { $_.ToString() } | Out-String
                    if ($streamDetail.Trim()) {
                        $detail = "=== Runspace Streams.Error ===`r`n$streamDetail`r`n`r`n=== Exception ===`r`n$detail"
                    }
                } catch { }
                if ($ctx.OnError) { $ctx.OnError.Invoke($detail) }
            } finally {
                try { $ctx.Ps.Dispose()  } catch { }
                try { $ctx.Pool.Close()  } catch { }
                try { $ctx.Pool.Dispose()} catch { }
            }
        }
    })
    $timer.Start()
}

# ===========================================================================
# ROLLENAUSWAHL
# ===========================================================================
$chosenRole = $Role

if (-not $chosenRole -or $chosenRole -eq 'Auto') {
    if (-not $chosenRole) {
        # Exchange-Pfad fuer Dialog brauchen wir schon hier
        $tmpExPath = $Config.DefaultExchangePath
        $chosenRole = Show-RoleDialog -ExchangePath $tmpExPath `
                                      -StateFileName $Config.StateFileName
        if (-not $chosenRole) { exit 0 }  # Abbruch
    }
    if ($chosenRole -eq 'Auto') {
        $stateFile = Join-Path $Config.DefaultExchangePath $Config.StateFileName
        $chosenRole = if (Test-Path $stateFile) { 'Target' } else { 'Source' }
        Write-MigrationLog -Level 'INFO' -Category 'ROLE' `
            -Message "Automatische Rolle: $chosenRole"
    }
}
$script:ActiveRole = $chosenRole

# Im Ziel-Modus: Zustandsdatei laden
if ($script:ActiveRole -eq 'Target') {
    $script:LoadedState = Read-MigrationState `
        -ExchangePath $Config.DefaultExchangePath `
        -StateFileName $Config.StateFileName
}

# ===========================================================================
# HAUPTFENSTER
# ===========================================================================
$roleLabel = if ($script:ActiveRole -eq 'Source') { 'QUELL-SERVER' } else { 'ZIEL-SERVER' }
$roleColor = if ($script:ActiveRole -eq 'Source') { $clrAccent } else { $clrAccentGrn }

$yearSpan = "2025-$((Get-Date).ToString('yy'))"
$form                  = New-Object System.Windows.Forms.Form
$form.Text             = "SQL Server Migration Tool  v1.1  -  $roleLabel   |   powershelldba.de - Janke (c) $yearSpan"
$form.Size             = New-Object System.Drawing.Size(1600, 900)   # 16:9
$form.MinimumSize      = New-Object System.Drawing.Size(1100, 650)
$form.StartPosition    = 'CenterScreen'
$form.BackColor        = $clrBg
$form.ForeColor        = $clrText
$form.Font             = $fntNormal

# ---------------------------------------------------------------------------
# TITELLEISTE
# ---------------------------------------------------------------------------
$pnlTitle           = New-Object System.Windows.Forms.Panel
$pnlTitle.Dock      = 'Top'
$pnlTitle.Height    = 50
$pnlTitle.BackColor = $clrHeader
$form.Controls.Add($pnlTitle)

$lblTitle           = New-Object System.Windows.Forms.Label
$lblTitle.Text      = ">>  SQL SERVER MIGRATION  -  $roleLabel"
$lblTitle.Font      = $fntTitle
$lblTitle.ForeColor = $roleColor
$lblTitle.AutoSize  = $true
$lblTitle.Location  = New-Object System.Drawing.Point(15, 13)
$pnlTitle.Controls.Add($lblTitle)

$lblVersion           = New-Object System.Windows.Forms.Label
$lblVersion.Text      = 'PS ' + $PSVersionTable.PSVersion.ToString() + '  |  dbaTools'
$lblVersion.Font      = $fntSmall
$lblVersion.ForeColor = $clrSubText
$lblVersion.AutoSize  = $true
$lblVersion.Location  = New-Object System.Drawing.Point(560, 17)
$pnlTitle.Controls.Add($lblVersion)

$lblLogPath           = New-Object System.Windows.Forms.Label
$lblLogPath.Font      = $fntSmall
$lblLogPath.ForeColor = $clrSubText
$lblLogPath.AutoSize  = $false
$lblLogPath.TextAlign = 'MiddleRight'
$lblLogPath.Size      = New-Object System.Drawing.Size(300, 20)
$lblLogPath.Location  = New-Object System.Drawing.Point(570, 15)
$lblLogPath.Text      = 'Log: ' + (Get-MigrationLogPath)
$pnlTitle.Controls.Add($lblLogPath)

# ---------------------------------------------------------------------------
# STATUSLEISTE
# ---------------------------------------------------------------------------
$pnlStatus           = New-Object System.Windows.Forms.Panel
$pnlStatus.Dock      = 'Bottom'
$pnlStatus.Height    = 28
$pnlStatus.BackColor = $clrHeader
$form.Controls.Add($pnlStatus)

$script:statusLabel          = New-Object System.Windows.Forms.Label
$script:statusLabel.Dock     = 'Fill'
$script:statusLabel.Font     = $fntSmall
$script:statusLabel.ForeColor= $clrText
$script:statusLabel.BackColor= [System.Drawing.Color]::FromArgb(23,162,184)
$script:statusLabel.TextAlign= 'MiddleLeft'
$script:statusLabel.Padding  = New-Object System.Windows.Forms.Padding(10,0,0,0)
$script:statusLabel.Text     = "  Bereit. Bitte $roleLabel verbinden."
$pnlStatus.Controls.Add($script:statusLabel)

# ---------------------------------------------------------------------------
# HAUPT-PANEL (einspaltig - nur eine Rolle)
# ---------------------------------------------------------------------------
$pnlMain           = New-Object System.Windows.Forms.Panel
$pnlMain.Dock      = 'Fill'
$pnlMain.BackColor = $clrBg
$form.Controls.Add($pnlMain)
$form.Controls.SetChildIndex($pnlMain, 0)

# --- Kopfzeile Server-Panel ---
$pnlHead            = New-Object System.Windows.Forms.Panel
$pnlHead.Dock       = 'Top'
$pnlHead.Height     = 32
$pnlHead.BackColor  = $clrHeader
$pnlMain.Controls.Add($pnlHead)

$lblCaption         = New-Object System.Windows.Forms.Label
$lblCaption.Text    = $roleLabel
$lblCaption.Font    = $fntBold
$lblCaption.ForeColor = $roleColor
$lblCaption.AutoSize  = $true
$lblCaption.Location  = New-Object System.Drawing.Point(10,7)
$pnlHead.Controls.Add($lblCaption)

$lblConnState        = New-Object System.Windows.Forms.Label
$lblConnState.Text   = '* Nicht verbunden'
$lblConnState.Font   = $fntSmall
$lblConnState.ForeColor = [System.Drawing.Color]::FromArgb(220,53,69)
$lblConnState.AutoSize  = $true
$lblConnState.Location  = New-Object System.Drawing.Point(160,9)
$pnlHead.Controls.Add($lblConnState)

# --- Verbindungs-Panel ---
$pnlConn            = New-Object System.Windows.Forms.Panel
$pnlConn.Dock       = 'Top'
$pnlConn.Height     = 115
$pnlConn.BackColor  = $clrPanel
$pnlConn.Padding    = New-Object System.Windows.Forms.Padding(8,6,8,6)
$pnlMain.Controls.Add($pnlConn)

$lblSrv             = New-Object System.Windows.Forms.Label
$lblSrv.Text        = 'Server\Instanz:'
$lblSrv.Font        = $fntSmall
$lblSrv.ForeColor   = $clrSubText
$lblSrv.AutoSize    = $true
$lblSrv.Location    = New-Object System.Drawing.Point(8, 10)
$pnlConn.Controls.Add($lblSrv)

$txtServer          = New-Object System.Windows.Forms.TextBox
$txtServer.Font     = $fntMono
$txtServer.BackColor= $clrInput
$txtServer.ForeColor= $clrText
$txtServer.BorderStyle = 'FixedSingle'
$txtServer.Size     = New-Object System.Drawing.Size(250, 22)
$txtServer.Location = New-Object System.Drawing.Point(115, 7)
# Standard: aktuellen Maschinennamen vorbelegen (lokale Instanz)
$txtServer.Text = $env:COMPUTERNAME
# Im Ziel-Modus: Zielserver aus Zustandsdatei vorbelegen (ueberschreibt Default)
if ($script:ActiveRole -eq 'Target' -and $script:LoadedState -and $script:LoadedState.TargetServer) {
    $txtServer.Text = $script:LoadedState.TargetServer
}
$pnlConn.Controls.Add($txtServer)

$lblAuth            = New-Object System.Windows.Forms.Label
$lblAuth.Text       = 'Auth:'
$lblAuth.Font       = $fntSmall
$lblAuth.ForeColor  = $clrSubText
$lblAuth.AutoSize   = $true
$lblAuth.Location   = New-Object System.Drawing.Point(380, 10)
$pnlConn.Controls.Add($lblAuth)

$cmbAuth            = New-Object System.Windows.Forms.ComboBox
$cmbAuth.Font       = $fntSmall
$cmbAuth.BackColor  = $clrInput
$cmbAuth.ForeColor  = $clrText
$cmbAuth.FlatStyle  = 'Flat'
$cmbAuth.DropDownStyle = 'DropDownList'
$cmbAuth.Size       = New-Object System.Drawing.Size(110, 22)
$cmbAuth.Location   = New-Object System.Drawing.Point(415, 7)
$cmbAuth.Items.AddRange(@('Windows', 'SQL-Login')) | Out-Null
$cmbAuth.SelectedIndex = 0
$pnlConn.Controls.Add($cmbAuth)

$lblUser = New-Object System.Windows.Forms.Label
$lblUser.Text = 'Benutzer:'; $lblUser.Font = $fntSmall
$lblUser.ForeColor = $clrSubText; $lblUser.AutoSize = $true
$lblUser.Location = New-Object System.Drawing.Point(8, 40); $lblUser.Visible = $false
$pnlConn.Controls.Add($lblUser)

$txtUser = New-Object System.Windows.Forms.TextBox
$txtUser.Font = $fntMono; $txtUser.BackColor = $clrInput; $txtUser.ForeColor = $clrText
$txtUser.BorderStyle = 'FixedSingle'; $txtUser.Size = New-Object System.Drawing.Size(140, 22)
$txtUser.Location = New-Object System.Drawing.Point(75, 37); $txtUser.Visible = $false
$pnlConn.Controls.Add($txtUser)

$lblPwd = New-Object System.Windows.Forms.Label
$lblPwd.Text = 'Passwort:'; $lblPwd.Font = $fntSmall
$lblPwd.ForeColor = $clrSubText; $lblPwd.AutoSize = $true
$lblPwd.Location = New-Object System.Drawing.Point(230, 40); $lblPwd.Visible = $false
$pnlConn.Controls.Add($lblPwd)

$txtPwd = New-Object System.Windows.Forms.TextBox
$txtPwd.Font = $fntMono; $txtPwd.BackColor = $clrInput; $txtPwd.ForeColor = $clrText
$txtPwd.BorderStyle = 'FixedSingle'; $txtPwd.PasswordChar = '*'
$txtPwd.Size = New-Object System.Drawing.Size(140, 22)
$txtPwd.Location = New-Object System.Drawing.Point(295, 37); $txtPwd.Visible = $false
$pnlConn.Controls.Add($txtPwd)

$cmbAuth.Add_SelectedIndexChanged({
    $isSql = ($cmbAuth.SelectedItem -eq 'SQL-Login')
    $lblUser.Visible = $isSql; $txtUser.Visible = $isSql
    $lblPwd.Visible  = $isSql; $txtPwd.Visible  = $isSql
    $pnlConn.Height  = if ($isSql) { 155 } else { 140 }
})

$chkTrust = New-Object System.Windows.Forms.CheckBox
$chkTrust.Text     = 'TrustServerCertificate (SQL 2022 / selbstsigniertes Zertifikat)'
$chkTrust.Font     = $fntSmall
$chkTrust.ForeColor= [System.Drawing.Color]::FromArgb(255,193,7)
$chkTrust.AutoSize = $true
$chkTrust.Checked  = [bool]$Config.TrustServerCertificate
$chkTrust.Location = New-Object System.Drawing.Point(8, 75)
$pnlConn.Controls.Add($chkTrust)

$btnConnect = New-Object System.Windows.Forms.Button
$btnConnect.Text = 'Verbinden'; $btnConnect.Font = $fntBold
$btnConnect.Size = New-Object System.Drawing.Size(90, 26)
$btnConnect.Location = New-Object System.Drawing.Point(8, 100)
$btnConnect.BackColor = $roleColor; $btnConnect.ForeColor = [System.Drawing.Color]::White
$btnConnect.FlatStyle = 'Flat'; $btnConnect.FlatAppearance.BorderSize = 0
$pnlConn.Controls.Add($btnConnect)

$btnDisconnect = New-Object System.Windows.Forms.Button
$btnDisconnect.Text = 'Trennen'; $btnDisconnect.Font = $fntSmall
$btnDisconnect.Size = New-Object System.Drawing.Size(70, 26)
$btnDisconnect.Location = New-Object System.Drawing.Point(108, 100)
$btnDisconnect.BackColor = [System.Drawing.Color]::FromArgb(60,63,70)
$btnDisconnect.ForeColor = $clrText; $btnDisconnect.FlatStyle = 'Flat'
$btnDisconnect.FlatAppearance.BorderSize = 0; $btnDisconnect.Enabled = $false
$pnlConn.Controls.Add($btnDisconnect)

$lblInfo = New-Object System.Windows.Forms.Label
$lblInfo.Font = $fntSmall; $lblInfo.ForeColor = $clrSubText
$lblInfo.AutoSize = $false; $lblInfo.Size = New-Object System.Drawing.Size(350, 18)
$lblInfo.Location = New-Object System.Drawing.Point(190, 107)
$pnlConn.Controls.Add($lblInfo)

# pnlConn Standardhoehe anpassen
$pnlConn.Height = 140

# --- Globale Auswahl-Leiste (alle Tabs) ---
$pnlGlobalSel           = New-Object System.Windows.Forms.Panel
$pnlGlobalSel.Dock      = 'Top'
$pnlGlobalSel.Height    = 28
$pnlGlobalSel.BackColor = $clrHeader
$pnlMain.Controls.Add($pnlGlobalSel)

$lblGlobal              = New-Object System.Windows.Forms.Label
$lblGlobal.Text         = 'Alle Tabs:'
$lblGlobal.Font         = $fntSmall
$lblGlobal.ForeColor    = $clrSubText
$lblGlobal.AutoSize     = $true
$lblGlobal.Location     = New-Object System.Drawing.Point(6, 7)
$pnlGlobalSel.Controls.Add($lblGlobal)

$btnGlobalAll           = New-Object System.Windows.Forms.Button
$btnGlobalAll.Text      = 'Alle markieren'
$btnGlobalAll.Font      = $fntSmall
$btnGlobalAll.Size      = New-Object System.Drawing.Size(110, 22)
$btnGlobalAll.Location  = New-Object System.Drawing.Point(68, 3)
$btnGlobalAll.BackColor = [System.Drawing.Color]::FromArgb(40,167,69)
$btnGlobalAll.ForeColor = [System.Drawing.Color]::White
$btnGlobalAll.FlatStyle = 'Flat'
$btnGlobalAll.FlatAppearance.BorderSize = 0
$pnlGlobalSel.Controls.Add($btnGlobalAll)

$btnGlobalNone          = New-Object System.Windows.Forms.Button
$btnGlobalNone.Text     = 'Alle abwaehlen'
$btnGlobalNone.Font     = $fntSmall
$btnGlobalNone.Size     = New-Object System.Drawing.Size(110, 22)
$btnGlobalNone.Location = New-Object System.Drawing.Point(182, 3)
$btnGlobalNone.BackColor= [System.Drawing.Color]::FromArgb(108,117,125)
$btnGlobalNone.ForeColor= [System.Drawing.Color]::White
$btnGlobalNone.FlatStyle= 'Flat'
$btnGlobalNone.FlatAppearance.BorderSize = 0
$pnlGlobalSel.Controls.Add($btnGlobalNone)

# Globale Buttons: $script:_AllListViews wird nach Tab-Erstellung gesetzt
$btnGlobalAll.Add_Click({
    foreach ($lv_ in $script:_AllListViews) {
        $lv_.BeginUpdate()
        foreach ($item in $lv_.Items) { $item.Checked = $true }
        $lv_.EndUpdate()
    }
})
$btnGlobalNone.Add_Click({
    foreach ($lv_ in $script:_AllListViews) {
        $lv_.BeginUpdate()
        foreach ($item in $lv_.Items) { $item.Checked = $false }
        $lv_.EndUpdate()
    }
})

# --- Tab-Control ---
$tabs           = New-Object System.Windows.Forms.TabControl
$tabs.Dock      = 'Fill'
$tabs.BackColor = $clrBg
$tabs.Font      = $fntSmall
# Owner-Draw: Tab-Reiter im Dark-Theme einfaerben (Standard-Tabs ignorieren BackColor)
$tabs.DrawMode  = 'OwnerDrawFixed'
$tabs.SizeMode  = 'Fixed'
$tabs.ItemSize  = New-Object System.Drawing.Size(120, 26)
$tabs.Add_DrawItem({
    param($sender, $e)
    $tc       = $sender
    $page     = $tc.TabPages[$e.Index]
    $selected = ($e.Index -eq $tc.SelectedIndex)
    $bg = if ($selected) { $roleColor }  else { $clrHeader }
    $fg = if ($selected) { [System.Drawing.Color]::White } else { $clrSubText }
    $rect  = $e.Bounds
    $brush = New-Object System.Drawing.SolidBrush($bg)
    $e.Graphics.FillRectangle($brush, $rect)
    $sf = New-Object System.Drawing.StringFormat
    $sf.Alignment     = [System.Drawing.StringAlignment]::Center
    $sf.LineAlignment = [System.Drawing.StringAlignment]::Center
    $tb = New-Object System.Drawing.SolidBrush($fg)
    $e.Graphics.DrawString($page.Text, $tc.Font, $tb, ([System.Drawing.RectangleF]$rect), $sf)
    $brush.Dispose(); $tb.Dispose(); $sf.Dispose()
})
# Tab-Reiter auf die volle Breite ziehen, damit kein heller (themed) Reststreifen
# rechts neben den Tabs sichtbar bleibt. Bei Groessenaenderung neu berechnen.
$script:_fitTabsBusy = $false
$script:FitTabWidth = {
    # ACHTUNG: Das Setzen von ItemSize loest selbst SizeChanged aus.
    # Ohne Schutz entsteht eine Endlosrekursion (Haenger beim Layout).
    if ($script:_fitTabsBusy) { return }
    if ($tabs.TabCount -le 0) { return }
    $w = [int](($tabs.ClientSize.Width - 6) / $tabs.TabCount)
    if ($w -lt 70) { $w = 70 }
    if ($tabs.ItemSize.Width -eq $w) { return }   # bereits korrekt -> nichts tun
    $script:_fitTabsBusy = $true
    try   { $tabs.ItemSize = New-Object System.Drawing.Size($w, 26) }
    catch { }
    finally { $script:_fitTabsBusy = $false }
}
$tabs.Add_SizeChanged({ & $script:FitTabWidth })
$pnlMain.Controls.Add($tabs)

# WinForms Dock-Reihenfolge: Fill muss Index 0 sein.
# Top-Controls danach in umgekehrter visueller Reihenfolge
# (zuletzt eingefuegtes Top-Control erscheint oben).
# Gewuenschte visuelle Reihenfolge von oben nach unten:
#   pnlHead (Servertitel)
#   pnlConn (Verbindung)
#   pnlGlobalSel (Alle/Keine)
#   tabs (Fill - Rest)
$pnlMain.Controls.SetChildIndex($tabs,         0)  # Fill: immer Index 0
$pnlMain.Controls.SetChildIndex($pnlGlobalSel, 1)  # unterste Top-Leiste
$pnlMain.Controls.SetChildIndex($pnlConn,      2)  # darueber
$pnlMain.Controls.SetChildIndex($pnlHead,      3)  # ganz oben

function New-ObjectTab {
    param([string]$Title, [string[][]]$Columns)
    $tp = New-Object System.Windows.Forms.TabPage
    $tp.Text = $Title; $tp.BackColor = $clrBg; $tp.ForeColor = $clrText
    $tp.Padding = New-Object System.Windows.Forms.Padding(0)

    # --- Button-Leiste oben im Tab ---
    $pnlBtn            = New-Object System.Windows.Forms.Panel
    $pnlBtn.Dock       = 'Top'
    $pnlBtn.Height     = 26
    $pnlBtn.BackColor  = $clrHeader

    $btnAll            = New-Object System.Windows.Forms.Button
    $btnAll.Text       = 'Alle'
    $btnAll.Font       = $fntSmall
    $btnAll.Size       = New-Object System.Drawing.Size(55, 22)
    $btnAll.Location   = New-Object System.Drawing.Point(2, 2)
    $btnAll.BackColor  = [System.Drawing.Color]::FromArgb(40,167,69)
    $btnAll.ForeColor  = [System.Drawing.Color]::White
    $btnAll.FlatStyle  = 'Flat'
    $btnAll.FlatAppearance.BorderSize = 0
    $pnlBtn.Controls.Add($btnAll)

    $btnNone           = New-Object System.Windows.Forms.Button
    $btnNone.Text      = 'Keine'
    $btnNone.Font      = $fntSmall
    $btnNone.Size      = New-Object System.Drawing.Size(55, 22)
    $btnNone.Location  = New-Object System.Drawing.Point(60, 2)
    $btnNone.BackColor = [System.Drawing.Color]::FromArgb(108,117,125)
    $btnNone.ForeColor = [System.Drawing.Color]::White
    $btnNone.FlatStyle = 'Flat'
    $btnNone.FlatAppearance.BorderSize = 0
    $pnlBtn.Controls.Add($btnNone)

    $lblCount          = New-Object System.Windows.Forms.Label
    $lblCount.Text     = ''
    $lblCount.Font     = $fntSmall
    $lblCount.ForeColor= $clrSubText
    $lblCount.AutoSize = $true
    $lblCount.Location = New-Object System.Drawing.Point(120, 5)
    $pnlBtn.Controls.Add($lblCount)

    $tp.Controls.Add($pnlBtn)

    # --- ListView ---
    $lv = New-Object System.Windows.Forms.ListView
    $lv.Dock = 'Fill'; $lv.View = 'Details'; $lv.FullRowSelect = $true
    $lv.GridLines = $true; $lv.CheckBoxes = $true
    $lv.BackColor = $clrInput; $lv.ForeColor = $clrText
    $lv.BorderStyle = 'None'; $lv.Font = $fntSmall; $lv.MultiSelect = $true
    foreach ($col in $Columns) {
        $c = New-Object System.Windows.Forms.ColumnHeader
        $c.Text = $col[0]; $c.Width = [int]$col[1]
        $lv.Columns.Add($c) | Out-Null
    }

    # Tag = ListView-Referenz fuer Event-Handler (kein Closure noetig)
    $btnAll.Tag  = $lv
    $btnNone.Tag = $lv

    $btnAll.Add_Click({
        $lv_ = $this.Tag
        $lv_.BeginUpdate()
        foreach ($item in $lv_.Items) { $item.Checked = $true }
        $lv_.EndUpdate()
    })
    $btnNone.Add_Click({
        $lv_ = $this.Tag
        $lv_.BeginUpdate()
        foreach ($item in $lv_.Items) { $item.Checked = $false }
        $lv_.EndUpdate()
    })

    # Zaehler aktualisieren wenn Checked-Status sich aendert
    $lv.Tag = $lblCount
    $lv.Add_ItemChecked({
        $lbl_ = $this.Tag
        if ($lbl_) {
            $checked_ = ($this.CheckedItems.Count)
            $total_   = ($this.Items.Count)
            $lbl_.Text = "$checked_ / $total_ ausgewaehlt"
        }
    })

    $tp.Controls.Add($lv)
    # Dock-Reihenfolge: Fill (ListView) muss Index 0 sein, sonst ueberdeckt das
    # obere Button-Panel die Spaltenkoepfe der ListView.
    $tp.Controls.SetChildIndex($lv, 0)
    $tp.Controls.SetChildIndex($pnlBtn, 1)
    $tabs.TabPages.Add($tp)
    return $lv
}

$lvDbs    = New-ObjectTab 'Datenbanken' @(
                @('Name','160'),@('Status','70'),@('RecoveryModel','80'),
                @('SizeGB','65'),@('Compatibility','90'),@('Owner','110'))
$lvLogins = New-ObjectTab 'Logins' @(
                @('Name','180'),@('Typ','90'),@('Disabled','60'),
                @('Locked','55'),@('System','55'),@('Erstellt','120'))
$lvUsers  = New-ObjectTab 'DB-User' @(
                @('Datenbank','140'),@('User','140'),@('LoginTyp','90'),
                @('Login','130'),@('System','55'))
$lvLS     = New-ObjectTab 'Linked Server' @(
                @('Name','160'),@('Provider','110'),@('DataSource','160'),@('Product','100'))
$lvJobs   = New-ObjectTab 'Agent Jobs' @(
                @('Name','200'),@('Kategorie','110'),@('Owner','110'),
                @('Aktiv','50'),@('LetzterLauf','130'))
$lvCreds  = New-ObjectTab 'Credentials' @(
                @('Name','160'),@('Identity','160'),@('Erstellt','120'))
$lvProx   = New-ObjectTab 'Proxies' @(
                @('Name','160'),@('Credential','140'),@('Aktiv','50'),@('Beschreibung','200'))

# Liste aller ListViews fuer globale Alle/Keine Buttons
$script:_AllListViews = @($lvDbs,$lvLogins,$lvUsers,$lvLS,$lvJobs,$lvCreds,$lvProx)

# Tab-Breiten initial auf volle Breite ziehen (nach dem Anlegen aller Tabs)
& $script:FitTabWidth

# ===========================================================================
# VERBINDEN-LOGIK
# ===========================================================================
$btnConnect.Add_Click({
    # Alle Werte sofort in $script: sichern - lokale Variablen sind
    # in OnComplete/OnError Callbacks nicht verfuegbar (PS 5.1 ISE)
    $script:_ConnSrvName   = $txtServer.Text.Trim()
    $script:_ConnAuthMode  = $cmbAuth.SelectedItem
    $script:_ConnSqlUser   = $txtUser.Text.Trim()
    $script:_ConnTrust     = $chkTrust.Checked
    $script:_ConnTimeout   = $Config.ConnectionTimeout
    $script:_ConnPwdSS     = if ($txtPwd.Text) {
        $ss = New-Object System.Security.SecureString
        $txtPwd.Text.ToCharArray() | ForEach-Object { $ss.AppendChar($_) }; $ss
    } else { $null }

    if (-not $script:_ConnSrvName) {
        [System.Windows.Forms.MessageBox]::Show(
            'Bitte Servernamen eingeben.','Hinweis','OK','Warning') | Out-Null
        return
    }
    $btnConnect.Enabled = $false; $btnConnect.Text = '...'
    $lblConnState.Text = '* Verbinde...'
    $lblConnState.ForeColor = [System.Drawing.Color]::FromArgb(255,193,7)
    Update-StatusBar "Verbinde mit $($script:_ConnSrvName) ..."

    Invoke-BackgroundJob -ScriptBlock {
        param($srv, $auth, $user, $pwdSS, $timeout, $trust)
        $p = @{ ServerInstance=$srv; AuthMode=$auth; ConnectTimeout=$timeout
                TrustServerCertificate=$trust }
        if ($auth -eq 'SQL-Login') { $p['SqlUser']=$user; $p['SqlPassword']=$pwdSS }
        New-SqlConnection @p
    } -ArgumentList @(
        $script:_ConnSrvName, $script:_ConnAuthMode, $script:_ConnSqlUser,
        $script:_ConnPwdSS,   $script:_ConnTimeout,  $script:_ConnTrust
    ) `
      -OnComplete {
        param($result)
        $conn = if ($result -is [array]) { $result[0] } else { $result }
        $script:ActiveServer = $conn
        $script:ActiveAuth   = $script:_ConnAuthMode

        $lblInfo.Invoke([Action]{ $lblInfo.Text = $conn.VersionString })
        $lblConnState.Invoke([Action]{
            $lblConnState.Text     = "* Verbunden: $($script:_ConnSrvName)"
            $lblConnState.ForeColor= [System.Drawing.Color]::FromArgb(40,167,69)
        })
        $btnConnect.Invoke([Action]{
            $btnConnect.Enabled = $true; $btnConnect.Text = 'Neu laden'
        })
        $btnDisconnect.Invoke([Action]{ $btnDisconnect.Enabled = $true })
        Update-StatusBar "Verbunden mit $($script:_ConnSrvName). Lade Objekte..." 'Info'

        $dbs    = Get-SqlDatabases    -Server $conn -ExcludeSystem
        $logins = Get-SqlLogins       -Server $conn
        $users  = Get-SqlDbUsers      -Server $conn
        $ls     = Get-SqlLinkedServers -Server $conn
        $jobs   = Get-SqlAgentJobs    -Server $conn
        $creds  = Get-SqlCredentials  -Server $conn
        $prox   = Get-SqlProxies      -Server $conn

        $lvDbs.Invoke([Action]{
            Set-ListViewData $lvDbs $dbs @('Name','Status','RecoveryModel','SizeGB','Compatibility','Owner')
        })
        $lvLogins.Invoke([Action]{
            Set-ListViewData $lvLogins $logins @('Name','LoginType','IsDisabled','IsLocked','IsSystem','CreateDate')
        })
        $lvUsers.Invoke([Action]{
            Set-ListViewData $lvUsers $users @('Database','Name','LoginType','Login','IsSystemObject')
        })
        $lvLS.Invoke([Action]{
            Set-ListViewData $lvLS $ls @('Name','ProviderName','DataSource','ProductName')
        })
        $lvJobs.Invoke([Action]{
            Set-ListViewData $lvJobs $jobs @('Name','Category','OwnerLoginName','IsEnabled','LastRunDate')
        })
        $lvCreds.Invoke([Action]{
            Set-ListViewData $lvCreds $creds @('Name','Identity','CreateDate')
        })
        $lvProx.Invoke([Action]{
            Set-ListViewData $lvProx $prox @('Name','CredentialName','IsEnabled','Description')
        })
        Update-StatusBar "Objekte geladen von $($script:_ConnSrvName)" 'Success'

    } -OnError {
        param($err)
        $btnConnect.Invoke([Action]{
            $btnConnect.Enabled = $true; $btnConnect.Text = 'Verbinden'
        })
        $lblConnState.Invoke([Action]{
            $lblConnState.Text = '* Fehler'
            $lblConnState.ForeColor = [System.Drawing.Color]::FromArgb(220,53,69)
        })
        Update-StatusBar "Verbindungsfehler - Details im Fehler-Dialog" 'Error'
        Show-ErrorDialog -Title 'Verbindungsfehler' -Message $err
    }
})

$btnDisconnect.Add_Click({
    $script:ActiveServer = $null
    $lblConnState.Text = '* Nicht verbunden'
    $lblConnState.ForeColor = [System.Drawing.Color]::FromArgb(220,53,69)
    $btnDisconnect.Enabled = $false; $btnConnect.Text = 'Verbinden'
    foreach ($lv in @($lvDbs,$lvLogins,$lvUsers,$lvLS,$lvJobs,$lvCreds,$lvProx)) {
        $lv.Items.Clear()
    }
    $lblInfo.Text = ''
    Update-StatusBar 'Verbindung getrennt.'
})

# ===========================================================================
# MIGRATIONS-STEUERUNG (unteres Panel)
# ===========================================================================
$pnlMigration           = New-Object System.Windows.Forms.Panel
$pnlMigration.Dock      = 'Bottom'
$pnlMigration.Height    = 230
$pnlMigration.BackColor = $clrPanel
$pnlMigration.Padding   = New-Object System.Windows.Forms.Padding(10,6,10,6)
$form.Controls.Add($pnlMigration)
$form.Controls.SetChildIndex($pnlMigration, 1)

$pnlDivider = New-Object System.Windows.Forms.Panel
$pnlDivider.Dock = 'Top'; $pnlDivider.Height = 2; $pnlDivider.BackColor = $roleColor
$pnlMigration.Controls.Add($pnlDivider)

$lblMigTitle = New-Object System.Windows.Forms.Label
$lblMigTitle.Text = if ($script:ActiveRole -eq 'Source') {
    '**  PHASE 1 - QUELLE: Backup / Detach + Copy'
} else {
    '**  PHASE 2 - ZIEL: Copy + Restore / Attach'
}
$lblMigTitle.Font = $fntBold; $lblMigTitle.ForeColor = $roleColor
$lblMigTitle.AutoSize = $true; $lblMigTitle.Location = New-Object System.Drawing.Point(10,12)
$pnlMigration.Controls.Add($lblMigTitle)

# --- Methode ---
$lblMethod = New-Object System.Windows.Forms.Label
$lblMethod.Text = 'Methode:'; $lblMethod.Font = $fntSmall
$lblMethod.ForeColor = $clrSubText; $lblMethod.AutoSize = $true
$lblMethod.Location = New-Object System.Drawing.Point(10,42)
$pnlMigration.Controls.Add($lblMethod)

$cmbMethod = New-Object System.Windows.Forms.ComboBox
$cmbMethod.Font = $fntNormal; $cmbMethod.BackColor = $clrInput
$cmbMethod.ForeColor = $clrText; $cmbMethod.FlatStyle = 'Flat'
$cmbMethod.DropDownStyle = 'DropDownList'
$cmbMethod.Size = New-Object System.Drawing.Size(160, 24)
$cmbMethod.Location = New-Object System.Drawing.Point(80,39)
$cmbMethod.Items.AddRange(@('Backup / Restore','Detach / Attach')) | Out-Null
$cmbMethod.SelectedIndex = 0
# Im Ziel-Modus: Methode aus Zustandsdatei vorbelegen und sperren
if ($script:ActiveRole -eq 'Target' -and $script:LoadedState) {
    $idx = if ($script:LoadedState.Method -eq 'DetachAttach') { 1 } else { 0 }
    $cmbMethod.SelectedIndex = $idx
    $cmbMethod.Enabled = $false
}
$pnlMigration.Controls.Add($cmbMethod)

# --- Exchange-Pfad ---
$lblBpth = New-Object System.Windows.Forms.Label
$lblBpth.Text = 'Exchange-Pfad:'; $lblBpth.Font = $fntSmall
$lblBpth.ForeColor = $clrSubText; $lblBpth.AutoSize = $true
$lblBpth.Location = New-Object System.Drawing.Point(255,42)
$pnlMigration.Controls.Add($lblBpth)

$txtExPath = New-Object System.Windows.Forms.TextBox
$txtExPath.Font = $fntMono; $txtExPath.BackColor = $clrInput
$txtExPath.ForeColor = $clrText; $txtExPath.BorderStyle = 'FixedSingle'
$txtExPath.Size = New-Object System.Drawing.Size(310,22)
$txtExPath.Location = New-Object System.Drawing.Point(355,39)
$txtExPath.Text = $Config.DefaultExchangePath
$pnlMigration.Controls.Add($txtExPath)

$btnBrowseEx = New-Object System.Windows.Forms.Button
$btnBrowseEx.Text = '...'; $btnBrowseEx.Font = $fntBold
$btnBrowseEx.Size = New-Object System.Drawing.Size(30,22)
$btnBrowseEx.Location = New-Object System.Drawing.Point(673,39)
$btnBrowseEx.BackColor = [System.Drawing.Color]::FromArgb(60,63,70)
$btnBrowseEx.ForeColor = $clrText; $btnBrowseEx.FlatStyle = 'Flat'
$btnBrowseEx.FlatAppearance.BorderSize = 0
$btnBrowseEx.Add_Click({
    $fbd = New-Object System.Windows.Forms.FolderBrowserDialog
    $fbd.Description = 'Exchange-Verzeichnis waehlen'
    $fbd.SelectedPath = $txtExPath.Text
    if ($fbd.ShowDialog() -eq 'OK') {
        $txtExPath.Text = $fbd.SelectedPath
        $script:ExchangePath = $fbd.SelectedPath
    }
})
$pnlMigration.Controls.Add($btnBrowseEx)

# --- Lokaler Backup-Pfad ---
$lblLPath = New-Object System.Windows.Forms.Label
$lblLPath.Text = 'Lokaler Backup-Pfad:'; $lblLPath.Font = $fntSmall
$lblLPath.ForeColor = $clrSubText; $lblLPath.AutoSize = $true
$lblLPath.Location = New-Object System.Drawing.Point(10,72)
$pnlMigration.Controls.Add($lblLPath)

$txtLPath = New-Object System.Windows.Forms.TextBox
$txtLPath.Font = $fntMono; $txtLPath.BackColor = $clrInput
$txtLPath.ForeColor = $clrText; $txtLPath.BorderStyle = 'FixedSingle'
$txtLPath.Size = New-Object System.Drawing.Size(310,22)
$txtLPath.Location = New-Object System.Drawing.Point(145,69)
$txtLPath.Text = $Config.DefaultLocalBackupPath
$pnlMigration.Controls.Add($txtLPath)

$btnBrowseL = New-Object System.Windows.Forms.Button
$btnBrowseL.Text = '...'; $btnBrowseL.Font = $fntBold
$btnBrowseL.Size = New-Object System.Drawing.Size(30,22)
$btnBrowseL.Location = New-Object System.Drawing.Point(463,69)
$btnBrowseL.BackColor = [System.Drawing.Color]::FromArgb(60,63,70)
$btnBrowseL.ForeColor = $clrText; $btnBrowseL.FlatStyle = 'Flat'
$btnBrowseL.FlatAppearance.BorderSize = 0
$btnBrowseL.Add_Click({
    $fbd = New-Object System.Windows.Forms.FolderBrowserDialog
    $fbd.Description = 'Lokalen Backup-Pfad waehlen'
    $fbd.SelectedPath = $txtLPath.Text
    if ($fbd.ShowDialog() -eq 'OK') {
        $txtLPath.Text = $fbd.SelectedPath
        $script:LocalBackupPath = $fbd.SelectedPath
    }
})
$pnlMigration.Controls.Add($btnBrowseL)

$lblLPathHint = New-Object System.Windows.Forms.Label
$lblLPathHint.Text = '(Fallback wenn Dienstkonto keinen UNC-Zugriff hat)'
$lblLPathHint.Font = $fntSmall; $lblLPathHint.ForeColor = $clrSubText
$lblLPathHint.AutoSize = $true
$lblLPathHint.Location = New-Object System.Drawing.Point(500,73)
$pnlMigration.Controls.Add($lblLPathHint)

# --- Optionen (nur Quell-Modus) ---
$chkWhatIf = New-Object System.Windows.Forms.CheckBox
$chkWhatIf.Text = 'WhatIf (nur simulieren)'; $chkWhatIf.Font = $fntSmall
$chkWhatIf.ForeColor = [System.Drawing.Color]::FromArgb(255,193,7)
$chkWhatIf.AutoSize = $true
$chkWhatIf.Location = New-Object System.Drawing.Point(10,102)
$pnlMigration.Controls.Add($chkWhatIf)

$chkReattach = New-Object System.Windows.Forms.CheckBox
$chkReattach.Text = 'Quelle Re-Attach nach Detach'; $chkReattach.Font = $fntSmall
$chkReattach.ForeColor = $clrText; $chkReattach.AutoSize = $true; $chkReattach.Checked = $true
$chkReattach.Location = New-Object System.Drawing.Point(200,102); $chkReattach.Visible = $false
$pnlMigration.Controls.Add($chkReattach)

$cmbMethod.Add_SelectedIndexChanged({
    $isDetach = ($cmbMethod.SelectedIndex -eq 1)
    $chkReattach.Visible = ($isDetach -and $script:ActiveRole -eq 'Source')
})

# --- Objekt-Auswahl (nur Quell-Modus sinnvoll; im Ziel aus Zustandsdatei) ---
$lblMigObj = New-Object System.Windows.Forms.Label
$lblMigObj.Text = 'Migrieren:'; $lblMigObj.Font = $fntSmall
$lblMigObj.ForeColor = $clrSubText; $lblMigObj.AutoSize = $true
$lblMigObj.Location = New-Object System.Drawing.Point(10,130)
$pnlMigration.Controls.Add($lblMigObj)

$migObjects = [ordered]@{
    'Datenbanken'    = $true
    'Logins'         = $true
    'Linked Server'  = $true
    'Agent Jobs'     = $true
    'Credentials'    = $true
    'Proxies'        = $true
}
$script:MigChkBoxes = @{}
$xPos = 90
foreach ($kv in $migObjects.GetEnumerator()) {
    $chk = New-Object System.Windows.Forms.CheckBox
    $chk.Text = $kv.Key; $chk.Checked = $kv.Value
    $chk.Font = $fntSmall; $chk.ForeColor = $clrText
    $chk.AutoSize = $true
    $chk.Location = New-Object System.Drawing.Point($xPos, 128)
    # Im Ziel-Modus: aus Zustandsdatei vorbelegen und sperren
    if ($script:ActiveRole -eq 'Target' -and $script:LoadedState -and
        $script:LoadedState.MigrationObjects) {
        $mo = $script:LoadedState.MigrationObjects
        $prop = $mo.PSObject.Properties[$kv.Key]
        if ($prop) { $chk.Checked = [bool]$prop.Value }
        $chk.Enabled = $false
    }
    $pnlMigration.Controls.Add($chk)
    $script:MigChkBoxes[$kv.Key] = $chk
    $xPos += 115
}

# --- Progress ---
$progress = New-Object System.Windows.Forms.ProgressBar
$progress.Style = 'Continuous'
$progress.Size = New-Object System.Drawing.Size(600,16)
$progress.Location = New-Object System.Drawing.Point(10,158)
$progress.BackColor = $clrInput; $progress.ForeColor = $roleColor
$pnlMigration.Controls.Add($progress)

$lblProgress = New-Object System.Windows.Forms.Label
$lblProgress.Text = ''; $lblProgress.Font = $fntSmall
$lblProgress.ForeColor = $clrSubText; $lblProgress.AutoSize = $true
$lblProgress.Location = New-Object System.Drawing.Point(620,161)
$pnlMigration.Controls.Add($lblProgress)

# --- Buttons ---
$btnActionText = if ($script:ActiveRole -eq 'Source') {
    '>>  PHASE 1 STARTEN  (Backup / Detach + Copy)'
} else {
    '>>  PHASE 2 STARTEN  (Restore / Attach)'
}

$btnMigrate = New-Object System.Windows.Forms.Button
$btnMigrate.Text = $btnActionText; $btnMigrate.Font = $fntBold
$btnMigrate.Size = New-Object System.Drawing.Size(310,34)
$btnMigrate.Location = New-Object System.Drawing.Point(10,183)
$btnMigrate.BackColor = $clrAccentGrn; $btnMigrate.ForeColor = [System.Drawing.Color]::White
$btnMigrate.FlatStyle = 'Flat'; $btnMigrate.FlatAppearance.BorderSize = 0
$pnlMigration.Controls.Add($btnMigrate)

$btnOpenLog = New-Object System.Windows.Forms.Button
$btnOpenLog.Text = '[Log]'; $btnOpenLog.Font = $fntSmall
$btnOpenLog.Size = New-Object System.Drawing.Size(80,34)
$btnOpenLog.Location = New-Object System.Drawing.Point(330,183)
$btnOpenLog.BackColor = [System.Drawing.Color]::FromArgb(60,63,70)
$btnOpenLog.ForeColor = $clrText; $btnOpenLog.FlatStyle = 'Flat'
$btnOpenLog.FlatAppearance.BorderSize = 0
$btnOpenLog.Add_Click({
    $lp = Get-MigrationLogPath
    if ($lp -and (Test-Path $lp)) { Start-Process notepad.exe -ArgumentList $lp }
    else { [System.Windows.Forms.MessageBox]::Show('Logdatei nicht gefunden.','Hinweis','OK','Info') | Out-Null }
})
$pnlMigration.Controls.Add($btnOpenLog)

$btnOpenCsv = New-Object System.Windows.Forms.Button
$btnOpenCsv.Text = '[CSV]'; $btnOpenCsv.Font = $fntSmall
$btnOpenCsv.Size = New-Object System.Drawing.Size(80,34)
$btnOpenCsv.Location = New-Object System.Drawing.Point(420,183)
$btnOpenCsv.BackColor = [System.Drawing.Color]::FromArgb(60,63,70)
$btnOpenCsv.ForeColor = $clrText; $btnOpenCsv.FlatStyle = 'Flat'
$btnOpenCsv.FlatAppearance.BorderSize = 0
$btnOpenCsv.Add_Click({
    $cp = Get-MigrationCsvPath
    if ($cp -and (Test-Path $cp)) { Start-Process $cp }
    else { [System.Windows.Forms.MessageBox]::Show('CSV nicht gefunden.','Hinweis','OK','Info') | Out-Null }
})
$pnlMigration.Controls.Add($btnOpenCsv)

# ===========================================================================
# MIGRATIONS-KLICK-HANDLER
# ===========================================================================
$btnMigrate.Add_Click({

    if (-not $script:ActiveServer) {
        [System.Windows.Forms.MessageBox]::Show(
            "Bitte zuerst $roleLabel verbinden.",'Hinweis','OK','Warning') | Out-Null
        return
    }

    $exPath  = $txtExPath.Text.Trim()
    $lPath   = $txtLPath.Text.Trim()
    $method  = if ($cmbMethod.SelectedIndex -eq 1) { 'DetachAttach' } else { 'BackupRestore' }
    $whatif  = $chkWhatIf.Checked
    $reattach= $chkReattach.Checked

    if (-not $exPath) {
        [System.Windows.Forms.MessageBox]::Show(
            'Bitte Exchange-Pfad angeben.','Hinweis','OK','Warning') | Out-Null
        return
    }

    # UNC-Zugriff testen (Admin-Credentials)
    Update-StatusBar "Pruefe Exchange-Pfad ..." 'Info'
    if (-not (Test-UncAccess -Path $exPath -TimeoutSec $Config.UncAccessTestTimeoutSec)) {
        [System.Windows.Forms.MessageBox]::Show(
            "Exchange-Pfad nicht erreichbar:`n$exPath`n`nBitte Pfad pruefen.",
            'Pfad nicht erreichbar','OK','Error') | Out-Null
        Update-StatusBar "Exchange-Pfad nicht erreichbar." 'Error'
        return
    }

    # Ausgewaehlte DBs
    $selDbs = @($lvDbs.CheckedItems | ForEach-Object { $_.Text })
    if ($selDbs.Count -eq 0 -and $script:MigChkBoxes['Datenbanken'].Checked) {
        [System.Windows.Forms.MessageBox]::Show(
            'Bitte mindestens eine Datenbank auswaehlen.','Hinweis','OK','Warning') | Out-Null
        return
    }

    $migChks = @{}
    foreach ($kv in $script:MigChkBoxes.GetEnumerator()) {
        $migChks[$kv.Key] = $kv.Value.Checked
    }

    # Szenario ermitteln (nur Quell-Modus; Ziel-Modus ist immer Phase2)
    $scenario = 'TwoPhase'
    if ($script:ActiveRole -eq 'Source') {
        $tgtSrvName = if ($script:LoadedState) { $script:LoadedState.TargetServer } else { '' }
        if ($tgtSrvName) {
            $scenario = Get-MigrationScenario `
                -TargetServerInstance $tgtSrvName `
                -ExchangePath $exPath `
                -UncTimeout $Config.UncAccessTestTimeoutSec
        }
        # Kein Zielserver bekannt -> immer TwoPhase
    }
    $script:Scenario = $scenario

    # Zusammenfassung
    $phaseInfo = if ($script:ActiveRole -eq 'Source') {
        if ($scenario -eq 'Direct') {
            "Szenario:    Direct (Zielserver erreichbar, UNC-Fallback aktiv)`n"
        } else {
            "Szenario:    TwoPhase (nur Phase 1 - Quelle)`n"
        }
    } else {
        "Szenario:    Phase 2 - Ziel`n"
    }

    $summary  = $phaseInfo
    $summary += "Methode:     $method`n"
    $summary += "WhatIf:      $whatif`n"
    $summary += "Exchange:    $exPath`n"
    $summary += "Lokal:       $lPath`n"
    if ($selDbs.Count -gt 0) {
        $summary += "Datenbanken: $($selDbs -join ', ')`n"
    }
    $summary += "`nObjekte:`n"
    foreach ($kv in $migChks.GetEnumerator()) {
        if ($kv.Value) { $summary += "  [x] $($kv.Key)`n" }
    }

    $resp = [System.Windows.Forms.MessageBox]::Show(
        "Migration starten?`n`n$summary",
        'Bestaetigung','YesNo','Question')
    if ($resp -ne 'Yes') { return }

    $btnMigrate.Enabled = $false
    $progress.Value = 0; $lblProgress.Text = 'Starte...'

    $srcSrv      = $script:ActiveServer
    $activeRole  = $script:ActiveRole
    $loadedState = $script:LoadedState
    $cfgRef      = $Config

    Invoke-BackgroundJob -ScriptBlock {
        param($srv, $role, $dbs, $exPath, $lPath, $method, $detach,
              $reatt, $whatif, $chks, $scenario, $loadedState, $cfg)

        $done = 0

        # ================================================================
        # QUELL-MODUS (Phase 1)
        # ================================================================
        if ($role -eq 'Source') {

            $backupFiles = @()
            $dbFileMap   = @{}

            if ($chks['Datenbanken'] -and $dbs.Count -gt 0) {

                $migMode = if ($scenario -eq 'Direct') { 'Full' } else { 'Phase1' }

                if ($method -eq 'BackupRestore') {
                    $fallbackCb = [System.Action[string]]{
                        param($msg)
                        Update-StatusBar "Fallback: Backup lokal - Exchange-Pfad fuer Dienstkonto nicht erreichbar" 'Warn'
                        [System.Windows.Forms.MessageBox]::Show(
                            $msg,
                            'Berechtigungs-Fallback aktiv',
                            [System.Windows.Forms.MessageBoxButtons]::OK,
                            [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
                    }
                    $res = Invoke-DatabaseMigrationBackupRestore `
                        -SourceServer      $srv `
                        -TargetServer      $srv `
                        -Databases         $dbs `
                        -ExchangePath      $exPath `
                        -LocalBackupPath   $lPath `
                        -CopyOnly:($cfg.CopyOnlyBackup) `
                        -WithCompression:($cfg.BackupCompression) `
                        -VerifyBackup:($cfg.VerifyBackup) `
                        -WhatIf:$whatif `
                        -Mode              $migMode `
                        -FallbackCallback  $fallbackCb
                    $backupFiles = @($res | Where-Object { $_.BackupFile } |
                                     ForEach-Object { $_.BackupFile })
                }
                elseif ($method -eq 'DetachAttach') {
                    $res = Invoke-DatabaseMigrationDetachAttach `
                        -SourceServer      $srv `
                        -TargetServer      $srv `
                        -Databases         $dbs `
                        -ExchangePath      $exPath `
                        -ReattachOnSource:$reatt `
                        -WhatIf:$whatif `
                        -Mode              $migMode
                    if ($res._FileMap) { $dbFileMap = $res._FileMap }
                }
                $done++
            }

            # Logins/Jobs/etc. nur im Direct-Modus (Zielserver erreichbar)
            # Im TwoPhase-Modus werden diese in Phase 2 migriert
            if ($scenario -eq 'Direct') {
                # Im Direct-Modus brauchen wir den Zielserver -> nicht verfuegbar
                # hier -> diese Objekte koennen nur migriert werden wenn Ziel verbunden
                Write-MigrationLog -Level 'WARN' -Category 'MIGRATE' `
                    -Message "Direct-Modus: Login/Job-Migration erfordert Zielserver-Verbindung" `
                    -Detail "Bitte Script auf Zielserver ausfuehren fuer vollstaendige Migration"
            }

            # Logins: SQL-Logins erkennen (-> Ziel braucht Mixed Mode) und als Skript
            # exportieren (Phase 2 legt sie per Invoke-DbaQuery an - domaenenuebergreifend).
            $sqlLoginsPresent = $false
            $loginScriptFile  = ''
            if ($chks['Logins']) {
                try { $sqlLoginsPresent = Test-SourceHasSqlLogins -SourceServer $srv } catch { }
                try {
                    $sf = Export-MigrationLogins -SourceServer $srv -ExchangePath $exPath -WhatIf:$whatif
                    if ($sf) { $loginScriptFile = $sf }
                } catch { }
            }

            # Objekt-Skripte fuer den Umweg exportieren (Jobs/LS/Cred/Proxy)
            $objectScripts = @{}
            if ($chks['Credentials'])   { try { $x = Export-MigrationCredentials  -SourceServer $srv -ExchangePath $exPath -WhatIf:$whatif; if ($x) { $objectScripts['Credentials']   = $x } } catch { } }
            if ($chks['Proxies'])       { try { $x = Export-MigrationProxies       -SourceServer $srv -ExchangePath $exPath -WhatIf:$whatif; if ($x) { $objectScripts['Proxies']       = $x } } catch { } }
            if ($chks['Linked Server']) { try { $x = Export-MigrationLinkedServers -SourceServer $srv -ExchangePath $exPath -WhatIf:$whatif; if ($x) { $objectScripts['LinkedServers'] = $x } } catch { } }
            if ($chks['Agent Jobs'])    { try { $x = Export-MigrationAgentJobs     -SourceServer $srv -ExchangePath $exPath -WhatIf:$whatif; if ($x) { $objectScripts['Jobs']          = $x } } catch { } }

            # Zustandsdatei schreiben (auch im Direct-Modus als Protokoll)
            if (-not $whatif) {
                $tgtSrv = if ($loadedState) { $loadedState.TargetServer } else { '' }
                Write-MigrationState `
                    -ExchangePath      $exPath `
                    -StateFileName     $cfg.StateFileName `
                    -SourceServer      $srv.Name `
                    -TargetServer      $tgtSrv `
                    -Method            $method `
                    -Databases         $dbs `
                    -MigrationObjects  $chks `
                    -BackupFiles       $backupFiles `
                    -DbFileMap         $dbFileMap `
                    -ReattachOnSource  $reatt `
                    -WhatIf            $whatif `
                    -LocalBackupPath   $lPath `
                    -SqlLoginsPresent  $sqlLoginsPresent `
                    -LoginScriptFile   $loginScriptFile `
                    -ObjectScripts     $objectScripts | Out-Null
            }

            Write-MigrationLog -Level 'INFO' -Category 'PHASE1' `
                -Message "Phase 1 abgeschlossen."
        }

        # ================================================================
        # ZIEL-MODUS (Phase 2)
        # ================================================================
        elseif ($role -eq 'Target') {

            if (-not $loadedState) {
                throw "Keine Zustandsdatei gefunden. Bitte zuerst Phase 1 auf dem Quellserver ausfuehren."
            }

            $srcDbs    = @($loadedState.Databases)
            $fileMap   = @{}
            if ($loadedState.DbFileMap) {
                $loadedState.DbFileMap.PSObject.Properties |
                    ForEach-Object { $fileMap[$_.Name] = $_.Value }
            }

            if ($chks['Datenbanken'] -and $srcDbs.Count -gt 0) {

                if ($method -eq 'BackupRestore') {
                    $bkFiles = if ($loadedState.BackupFiles) {
                        @($loadedState.BackupFiles)
                    } else { @() }

                    $fallbackCb2 = [System.Action[string]]{
                        param($msg)
                        Update-StatusBar "Fallback: Backup lokal kopiert - Exchange-Pfad fuer Dienstkonto nicht erreichbar" 'Warn'
                        [System.Windows.Forms.MessageBox]::Show(
                            $msg,
                            'Berechtigungs-Fallback aktiv',
                            [System.Windows.Forms.MessageBoxButtons]::OK,
                            [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
                    }
                    $res = Invoke-DatabaseMigrationBackupRestore `
                        -SourceServer    $srv `
                        -TargetServer    $srv `
                        -Databases       $srcDbs `
                        -ExchangePath    $exPath `
                        -LocalBackupPath $lPath `
                        -WhatIf:$whatif `
                        -Mode            'Phase2' `
                        -FallbackCallback $fallbackCb2
                }
                elseif ($method -eq 'DetachAttach') {
                    $res = Invoke-DatabaseMigrationDetachAttach `
                        -SourceServer $srv `
                        -TargetServer $srv `
                        -Databases    $srcDbs `
                        -ExchangePath $exPath `
                        -WhatIf:$whatif `
                        -Mode         'Phase2' `
                        -DbFileMap    $fileMap
                }
                $done++
            }

            # ----------------------------------------------------------------
            # Post-Restore-Bereinigung (Ziel)
            # ----------------------------------------------------------------
            if ($chks['Datenbanken'] -and $srcDbs.Count -gt 0) {
                # Verwaiste DB-User reparieren + DB-Owner auf sa setzen
                Invoke-PostRestoreCleanup -TargetServer $srv -Databases $srcDbs -WhatIf:$whatif
            }
            if ($chks['Logins']) {
                # 1. Mixed Mode aktivieren (+ Dienst-Neustart), falls SQL-Logins
                #    transferiert werden und das Ziel nur Windows-Auth nutzt.
                #    Rueckgabe = ggf. neu aufgebaute Verbindung nach dem Neustart.
                $sqlLoginsPresent = $false
                $loginScriptFile  = ''
                if ($loadedState) {
                    if ($null -ne $loadedState.PSObject.Properties['SqlLoginsPresent']) {
                        $sqlLoginsPresent = [bool]$loadedState.SqlLoginsPresent
                    }
                    if ($null -ne $loadedState.PSObject.Properties['LoginScriptFile']) {
                        $loginScriptFile = [string]$loadedState.LoginScriptFile
                    }
                }
                $srv = Enable-MixedModeIfNeeded -TargetServer $srv -SqlLoginsPresent $sqlLoginsPresent -WhatIf:$whatif

                # 2. Passwort-Policy VOR dem Login-Import abschalten
                Set-NamedPbmPolicyState -TargetServer $srv -PolicyName 'New_Password_Policy' -Enabled $false -WhatIf:$whatif

                # 3. Logins aus dem Skript anlegen (Export-DbaLogin -> Invoke-DbaQuery)
                if ($loginScriptFile) {
                    Import-MigrationLogins -TargetServer $srv -ScriptFile $loginScriptFile -WhatIf:$whatif
                } else {
                    Write-MigrationLog -Level 'WARN' -Category 'LOGIN-IMPORT' `
                        -Message "Kein Login-Skript in der Zustandsdatei - Login-Import uebersprungen"
                }

                # 4. Passwort-Policy abschliessend wieder einschalten
                Set-NamedPbmPolicyState -TargetServer $srv -PolicyName 'New_Password_Policy' -Enabled $true -WhatIf:$whatif

                # 5. Verwaiste AD-Logins (geloeschte Domaenenkonten) entfernen
                Remove-DeadAdLogin -TargetServer $srv -WhatIf:$whatif
            }

            # Objekte aus Skripten anlegen (Umweg). Reihenfolge wegen Abhaengigkeiten:
            # Credentials -> Proxies -> Linked Server -> Agent Jobs.
            $objScripts = $null
            if ($loadedState -and ($null -ne $loadedState.PSObject.Properties['ObjectScripts'])) {
                $objScripts = $loadedState.ObjectScripts
            }
            if ($objScripts) {
                if ($chks['Credentials']   -and $objScripts.Credentials)   { Import-MigrationScriptFile -TargetServer $srv -ScriptFile $objScripts.Credentials   -Category 'CRED-IMPORT'  -WhatIf:$whatif }
                if ($chks['Proxies']       -and $objScripts.Proxies)       { Import-MigrationScriptFile -TargetServer $srv -ScriptFile $objScripts.Proxies       -Category 'PROXY-IMPORT' -WhatIf:$whatif }
                if ($chks['Linked Server'] -and $objScripts.LinkedServers) { Import-MigrationScriptFile -TargetServer $srv -ScriptFile $objScripts.LinkedServers -Category 'LS-IMPORT'    -WhatIf:$whatif }
                if ($chks['Agent Jobs']    -and $objScripts.Jobs)          { Import-MigrationScriptFile -TargetServer $srv -ScriptFile $objScripts.Jobs          -Category 'JOB-IMPORT'   -WhatIf:$whatif }
            }

            # Phase2 in Zustandsdatei vermerken
            if (-not $whatif) {
                Complete-MigrationState -ExchangePath $exPath `
                    -StateFileName $cfg.StateFileName
            }

            Write-MigrationLog -Level 'INFO' -Category 'PHASE2' `
                -Message "Phase 2 abgeschlossen. Lokale Backup-Dateien verbleiben in: $lPath" `
                -Detail "Bitte nach Pruefung manuell bereinigen."
        }

        return $done

    } -ArgumentList @(
        $srcSrv, $activeRole, $selDbs, $exPath, $lPath, $method,
        ($cmbMethod.SelectedIndex -eq 1), $reattach, $whatif, $migChks,
        $script:Scenario, $loadedState, $cfgRef
    ) -OnComplete {
        param($r)
        $btnMigrate.Invoke([Action]{ $btnMigrate.Enabled = $true })
        $progress.Invoke([Action]{ $progress.Value = 100 })
        $lblProgress.Invoke([Action]{ $lblProgress.Text = 'Abgeschlossen!' })

        $hinweis = if ($script:ActiveRole -eq 'Source' -and $script:Scenario -eq 'TwoPhase') {
            "`n`nPhase 1 abgeschlossen.`nBitte Script nun auf dem ZIEL-Server ausfuehren."
        } elseif ($script:ActiveRole -eq 'Source') {
            "`n`nDirect-Modus: Datenbank-Transfer abgeschlossen."
        } else {
            "`n`nPhase 2 abgeschlossen. Bitte Zieldatenbanken pruefen."
        }

        $lPathCurrent = $txtLPath.Text
        $lokHinweis = "`n`n[!] Lokale Backup-Dateien in '$lPathCurrent' wurden NICHT geloescht.`nBitte nach Pruefung manuell bereinigen."

        Update-StatusBar 'Abgeschlossen. Log pruefen.' 'Success'
        [System.Windows.Forms.MessageBox]::Show(
            "Migration abgeschlossen!$hinweis$lokHinweis`n`nLog: $(Get-MigrationLogPath)",
            'Fertig','OK','Information') | Out-Null

    } -OnError {
        param($err)
        $btnMigrate.Invoke([Action]{ $btnMigrate.Enabled = $true })
        $lblProgress.Invoke([Action]{ $lblProgress.Text = 'FEHLER' })
        Update-StatusBar "Fehler - Details im Fehler-Dialog" 'Error'
        Show-ErrorDialog -Title 'Migrationsfehler' -Message $err
    }
})

# ===========================================================================
# FORM ANZEIGEN
# ===========================================================================
Write-MigrationLog -Level 'INFO' -Category 'GUI' `
    -Message "GUI gestartet" -Detail "Rolle: $($script:ActiveRole)"
# Tab-Breiten an die finale Fenstergroesse anpassen, sobald das Formular sichtbar ist
$form.Add_Shown({ & $script:FitTabWidth })
[System.Windows.Forms.Application]::Run($form)
Write-MigrationLog -Level 'INFO' -Category 'GUI' -Message 'GUI beendet'
