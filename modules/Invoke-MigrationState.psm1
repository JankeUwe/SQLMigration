# =============================================================================
# Modul: Invoke-MigrationState.psm1
# Zweck: Zustandsdatei fuer zweistufige Migration (Quelle -> Ziel)
#        sowie UNC-Zugriffstest fuer automatische Strategiewahl
# =============================================================================

# ---------------------------------------------------------------------------
# UNC-Zugriffstest (unter den Credentials des ausfuehrenden Admins)
# ---------------------------------------------------------------------------
function Test-UncAccess {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [int]$TimeoutSec = 10
    )

    # Schneller Erreichbarkeitstest per Job mit Timeout
    $job = Start-Job -ScriptBlock {
        param($p)
        try {
            $null = [System.IO.Directory]::GetFileSystemEntries($p) | Select-Object -First 1
            return $true
        } catch {
            return $false
        }
    } -ArgumentList $Path

    $done = Wait-Job $job -Timeout $TimeoutSec
    if (-not $done) {
        Stop-Job  $job
        Remove-Job $job -Force
        return $false
    }
    $result = Receive-Job $job
    Remove-Job $job -Force
    return ($result -eq $true)
}

# ---------------------------------------------------------------------------
# Erreichbarkeit des Zielservers testen (TCP Port 1433 / named instance)
# ---------------------------------------------------------------------------
function Test-ServerReachable {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ServerInstance,
        [int]$TimeoutMs = 3000
    )
    # Instanzname abschneiden
    $host_ = $ServerInstance.Split('\')[0].Split(',')[0].Trim()
    try {
        $tcp = New-Object System.Net.Sockets.TcpClient
        $ar  = $tcp.BeginConnect($host_, 1433, $null, $null)
        $ok  = $ar.AsyncWaitHandle.WaitOne($TimeoutMs, $false)
        $tcp.Close()
        return $ok
    } catch {
        return $false
    }
}

# ---------------------------------------------------------------------------
# Szenario ermitteln
#   Returns: 'Direct' | 'TwoPhase'
#   Direct    = beide Server erreichbar UND UNC-Zugriff (Dienstkonto) moeglich
#               => wird weiter unten beim Backup-Versuch getestet
#   TwoPhase  = Zielserver nicht sichtbar ODER Cross-Domain
# ---------------------------------------------------------------------------
function Get-MigrationScenario {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$TargetServerInstance,
        [Parameter(Mandatory)][string]$ExchangePath,
        [int]$TimeoutMs  = 3000,
        [int]$UncTimeout = 10
    )

    $reachable = Test-ServerReachable -ServerInstance $TargetServerInstance `
                                      -TimeoutMs $TimeoutMs
    if (-not $reachable) {
        Write-MigrationLog -Level 'INFO' -Category 'SCENARIO' `
            -Message "Zielserver nicht erreichbar -> TwoPhase" `
            -Detail $TargetServerInstance
        return 'TwoPhase'
    }

    # UNC unter Admin-Credentials testbar, aber Dienstkonto-Zugriff
    # wird erst beim echten Backup getestet (Option B)
    Write-MigrationLog -Level 'INFO' -Category 'SCENARIO' `
        -Message "Zielserver erreichbar -> pruefe Direct/TwoPhase beim Backup" `
        -Detail $TargetServerInstance
    return 'Direct'
}

# ---------------------------------------------------------------------------
# Zustandsdatei schreiben (nach Phase 1 / Quelle)
# ---------------------------------------------------------------------------
function Write-MigrationState {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ExchangePath,
        [Parameter(Mandatory)][string]$StateFileName,
        [Parameter(Mandatory)][string]$SourceServer,
        [Parameter(Mandatory)][string]$TargetServer,
        [Parameter(Mandatory)][string]$Method,          # BackupRestore | DetachAttach
        [Parameter(Mandatory)][string[]]$Databases,
        [hashtable]$MigrationObjects,                   # welche Objekte migrieren
        # Backup/Restore spezifisch
        [string[]]$BackupFiles      = @(),
        # Detach/Attach spezifisch
        [hashtable]$DbFileMap       = @{},              # DbName -> @(Dateien)
        [bool]$ReattachOnSource     = $true,
        [bool]$WhatIf               = $false,
        [string]$LocalBackupPath    = ''
    )

    $state = [ordered]@{
        SchemaVersion    = '1.1'
        CreatedAt        = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
        CreatedBy        = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
        SourceServer     = $SourceServer
        TargetServer     = $TargetServer
        Method           = $Method
        Databases        = $Databases
        MigrationObjects = $MigrationObjects
        BackupFiles      = $BackupFiles
        DbFileMap        = $DbFileMap
        ReattachOnSource = $ReattachOnSource
        WhatIf           = $WhatIf
        LocalBackupPath  = $LocalBackupPath
        Phase1CompletedAt = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
        Phase2CompletedAt = ''
    }

    $stateFile = Join-Path $ExchangePath $StateFileName
    $state | ConvertTo-Json -Depth 10 | Out-File -FilePath $stateFile -Encoding UTF8 -Force

    Write-MigrationLog -Level 'SUCCESS' -Category 'STATE' `
        -Message "Zustandsdatei geschrieben" -Detail $stateFile
    return $stateFile
}

# ---------------------------------------------------------------------------
# Zustandsdatei lesen
# ---------------------------------------------------------------------------
function Read-MigrationState {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ExchangePath,
        [Parameter(Mandatory)][string]$StateFileName
    )

    $stateFile = Join-Path $ExchangePath $StateFileName
    if (-not (Test-Path $stateFile)) {
        return $null
    }

    try {
        $raw   = Get-Content $stateFile -Raw -Encoding UTF8
        $state = $raw | ConvertFrom-Json
        Write-MigrationLog -Level 'INFO' -Category 'STATE' `
            -Message "Zustandsdatei gelesen" `
            -Detail "Erstellt: $($state.CreatedAt) | Quelle: $($state.SourceServer)"
        return $state
    } catch {
        Write-MigrationLog -Level 'ERROR' -Category 'STATE' `
            -Message "Zustandsdatei unlesbar" -Detail $_.Exception.Message
        return $null
    }
}

# ---------------------------------------------------------------------------
# Phase2-Abschluss in Zustandsdatei vermerken
# ---------------------------------------------------------------------------
function Complete-MigrationState {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ExchangePath,
        [Parameter(Mandatory)][string]$StateFileName
    )

    $stateFile = Join-Path $ExchangePath $StateFileName
    if (-not (Test-Path $stateFile)) { return }

    try {
        $raw   = Get-Content $stateFile -Raw -Encoding UTF8
        $state = $raw | ConvertFrom-Json
        # PSObject -> Hashtable fuer Bearbeitung
        $ht    = @{}
        $state.PSObject.Properties | ForEach-Object { $ht[$_.Name] = $_.Value }
        $ht['Phase2CompletedAt'] = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
        $ht | ConvertTo-Json -Depth 10 | Out-File -FilePath $stateFile -Encoding UTF8 -Force
        Write-MigrationLog -Level 'SUCCESS' -Category 'STATE' `
            -Message "Phase 2 in Zustandsdatei vermerkt"
    } catch {
        Write-MigrationLog -Level 'WARN' -Category 'STATE' `
            -Message "Phase2-Vermerk fehlgeschlagen" -Detail $_.Exception.Message
    }
}

Export-ModuleMember -Function Test-UncAccess,
                               Test-ServerReachable,
                               Get-MigrationScenario,
                               Write-MigrationState,
                               Read-MigrationState,
                               Complete-MigrationState
