# =============================================================================
# Modul: Invoke-Migration.psm1
# Zweck: Migrationslogik fuer alle SQL-Objekte
#
# Standard-Strategie (Tool laeuft IMMER als Admin):
#   Backup/Restore:
#     Phase 1: Backup in das lokale Standard-Backup-Verzeichnis des Servers
#              (das SQL-Dienstkonto kann dort immer schreiben), danach robocopy
#              der .bak auf den Exchange-Pfad (laeuft im Admin-Kontext -> kein
#              Dienstkonto-/UNC-Problem).
#     Phase 2: robocopy der .bak vom Exchange-Pfad in das lokale Backup-Verzeichnis
#              des Zielservers, danach Restore von dort.
#   Detach/Attach:
#     Phase 1: Detach, robocopy der Datenbankdateien auf den Exchange-Pfad.
#     Phase 2: robocopy der Dateien in die Standard-DATA/LOG-Verzeichnisse des
#              Zielservers, danach Attach von dort.
#
# Post-Restore (Ziel):
#   - Verwaiste DB-User reparieren (Repair-DbaDbOrphanUser)
#   - DB-Owner auf sa setzen (sa per SID 0x01 ermittelt -> auch bei umbenanntem sa)
#   - Verwaiste AD-Logins (geloeschte Domaenenkonten) entfernen
# =============================================================================

# ---------------------------------------------------------------------------
# Interner Helfer: lokalen Backup-Pfad sicherstellen
# ---------------------------------------------------------------------------
function Ensure-LocalBackupPath {
    param([string]$Path)
    if (-not (Test-Path $Path)) {
        New-Item -ItemType Directory -Path $Path -Force -ErrorAction Stop | Out-Null
        Write-MigrationLog -Level 'INFO' -Category 'PATH' `
            -Message "Lokales Backup-Verzeichnis angelegt" -Detail $Path
    }
}

# ---------------------------------------------------------------------------
# Interner Helfer: Standard-Backup-Verzeichnis des Servers ermitteln
# (Server-Einstellung; Fallback auf den konfigurierten lokalen Pfad)
# ---------------------------------------------------------------------------
function Get-ServerBackupDirectory {
    param($Server, [string]$Fallback)
    try {
        $p = Get-DbaDefaultPath -SqlInstance $Server -ErrorAction Stop
        if ($p -and $p.Backup) { return $p.Backup }
    }
    catch {
        Write-MigrationLog -Level 'WARN' -Category 'PATH' `
            -Message "Server-Backup-Verzeichnis nicht ermittelbar, nutze Fallback" `
            -Detail $_.Exception.Message
    }
    return $Fallback
}

# ---------------------------------------------------------------------------
# Interner Helfer: eine Datei robust per robocopy kopieren (Admin-Kontext).
# robocopy ExitCodes < 8 = Erfolg, >= 8 = Fehler.
# ---------------------------------------------------------------------------
function Copy-FileRobocopy {
    param(
        [Parameter(Mandatory)][string]$SourceFile,
        [Parameter(Mandatory)][string]$DestDir
    )
    $srcDir = Split-Path $SourceFile -Parent
    $name   = Split-Path $SourceFile -Leaf
    if (-not (Test-Path $DestDir)) {
        New-Item -ItemType Directory -Path $DestDir -Force -ErrorAction Stop | Out-Null
    }
    $null = robocopy $srcDir $DestDir $name /R:2 /W:3 /NP /NJH /NJS /NDL /COPY:DAT
    $code = $LASTEXITCODE
    if ($code -ge 8) {
        throw "robocopy fehlgeschlagen (ExitCode $code): '$name' -> '$DestDir'"
    }
    $dest = Join-Path $DestDir $name
    if (-not (Test-Path $dest)) {
        throw "robocopy: Zieldatei fehlt nach Kopie: $dest"
    }
    return $dest
}

# ---------------------------------------------------------------------------
# DATENBANKEN - Backup / Restore (Standard: lokal + robocopy)
# ---------------------------------------------------------------------------
function Invoke-DatabaseMigrationBackupRestore {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$SourceServer,
        [Parameter(Mandatory)]$TargetServer,
        [Parameter(Mandatory)][string[]]$Databases,
        [Parameter(Mandatory)][string]$ExchangePath,
        [string]$LocalBackupPath  = 'F:\Daten\SQL\Backup',
        [switch]$CopyOnly,
        [switch]$WithCompression,
        [switch]$VerifyBackup,
        [switch]$WhatIf,
        [ValidateSet('Full','Phase1','Phase2')]
        [string]$Mode = 'Full',
        [System.Action[string]]$ProgressCallback,
        # Beibehalten fuer Signatur-Kompatibilitaet (GUI); im Standardpfad nicht mehr noetig.
        [System.Action[string]]$FallbackCallback
    )

    $results = [System.Collections.Generic.List[PSObject]]::new()

    foreach ($dbName in $Databases) {
        $stepResult = [PSCustomObject]@{
            Database        = $dbName
            Method          = 'BackupRestore'
            BackupOK        = $false
            CopyOK          = $false
            RestoreOK       = $false
            BackupFile      = ''
            UsedLocalBackup = $true
            Error           = ''
            Duration        = 0
        }
        $sw = [System.Diagnostics.Stopwatch]::StartNew()

        try {
            # ----------------------------------------------------------------
            # PHASE 1 / FULL: Backup lokal -> robocopy auf Exchange
            # ----------------------------------------------------------------
            if ($Mode -in 'Full','Phase1') {

                if ($ProgressCallback) { $ProgressCallback.Invoke("Backup: $dbName") }
                Write-MigrationLog -Level 'STEP' -Category 'BACKUP' -Message "Starte Backup: $dbName"

                if (-not $WhatIf) {
                    $localDir = Get-ServerBackupDirectory -Server $SourceServer -Fallback $LocalBackupPath
                    Ensure-LocalBackupPath -Path $localDir
                    Write-MigrationLog -Level 'INFO' -Category 'BACKUP' `
                        -Message "Backup in lokales Server-Verzeichnis" -Detail $localDir

                    $bpResult = Backup-DbaDatabase `
                        -SqlInstance      $SourceServer `
                        -Database         $dbName `
                        -BackupDirectory  $localDir `
                        -CompressBackup   $WithCompression.IsPresent `
                        -CopyOnly         $CopyOnly.IsPresent `
                        -ErrorAction      Stop
                    $localFile = $bpResult.BackupPath
                    $stepResult.BackupOK = $true
                    Write-MigrationLog -Level 'SUCCESS' -Category 'BACKUP' `
                        -Message "Backup lokal OK: $dbName" -Detail $localFile

                    if ($VerifyBackup) {
                        Write-MigrationLog -Level 'INFO' -Category 'BACKUP' -Message "Verifiziere Backup: $dbName"
                        $null = Test-DbaLastBackup -SqlInstance $SourceServer -Database $dbName -ErrorAction SilentlyContinue
                    }

                    # robocopy auf Exchange-Pfad (Admin-Kontext -> kein Dienstkonto-Problem)
                    if ($ProgressCallback) { $ProgressCallback.Invoke("Kopiere (robocopy): $dbName") }
                    $destFile = Copy-FileRobocopy -SourceFile $localFile -DestDir $ExchangePath
                    $stepResult.BackupFile = $destFile
                    $stepResult.CopyOK     = $true
                    Write-MigrationLog -Level 'SUCCESS' -Category 'COPY' `
                        -Message "Backup per robocopy auf Exchange-Pfad" -Detail $destFile
                    Write-MigrationLog -Level 'INFO' -Category 'BACKUP' `
                        -Message "Lokale Sicherung verbleibt im Server-Backup-Verzeichnis" -Detail $localFile
                }
                else {
                    $stepResult.BackupOK   = $true
                    $stepResult.CopyOK     = $true
                    $stepResult.BackupFile = Join-Path $ExchangePath "$dbName.bak"
                    Write-MigrationLog -Level 'INFO' -Category 'BACKUP' -Message "[WHATIF] Backup lokal + robocopy: $dbName"
                }
            }

            # ----------------------------------------------------------------
            # PHASE 2 / FULL: robocopy Exchange -> lokal, dann Restore
            # ----------------------------------------------------------------
            if ($Mode -in 'Full','Phase2') {

                if ($ProgressCallback) { $ProgressCallback.Invoke("Restore: $dbName") }
                Write-MigrationLog -Level 'STEP' -Category 'RESTORE' `
                    -Message "Starte Restore: $dbName auf $($TargetServer.Name)"

                if (-not $WhatIf) {
                    # Quelle der Backup-Datei
                    $srcFile = if ($stepResult.BackupFile) {
                        $stepResult.BackupFile
                    } else {
                        Get-ChildItem $ExchangePath -Filter "$dbName*.bak" -ErrorAction Stop |
                            Sort-Object LastWriteTime -Descending |
                            Select-Object -First 1 -ExpandProperty FullName
                    }
                    if (-not $srcFile) { throw "Keine Backup-Datei fuer '$dbName' im Exchange-Pfad gefunden." }

                    # robocopy in lokales Backup-Verzeichnis des Zielservers
                    $localDir  = Get-ServerBackupDirectory -Server $TargetServer -Fallback $LocalBackupPath
                    Ensure-LocalBackupPath -Path $localDir
                    if ($ProgressCallback) { $ProgressCallback.Invoke("Kopiere (robocopy): $dbName") }
                    $localCopy = Copy-FileRobocopy -SourceFile $srcFile -DestDir $localDir
                    Write-MigrationLog -Level 'SUCCESS' -Category 'COPY' `
                        -Message "Backup per robocopy lokal bereitgestellt" -Detail $localCopy

                    $null = Restore-DbaDatabase `
                        -SqlInstance  $TargetServer `
                        -Path         $localCopy `
                        -DatabaseName $dbName `
                        -WithReplace `
                        -ErrorAction  Stop
                    $stepResult.RestoreOK = $true
                    Write-MigrationLog -Level 'SUCCESS' -Category 'RESTORE' -Message "Restore OK: $dbName"
                }
                else {
                    $stepResult.RestoreOK = $true
                    Write-MigrationLog -Level 'INFO' -Category 'RESTORE' -Message "[WHATIF] robocopy + Restore: $dbName"
                }
            }
        }
        catch {
            $stepResult.Error = $_.Exception.Message
            Write-MigrationLog -Level 'ERROR' -Category 'DB-MIGRATION' `
                -Message "Fehler bei $dbName" -Detail $_.Exception.Message
        }
        finally {
            $sw.Stop()
            $stepResult.Duration = [math]::Round($sw.Elapsed.TotalSeconds, 1)
            $results.Add($stepResult)
        }
    }
    return $results
}

# ---------------------------------------------------------------------------
# DATENBANKEN - Detach / Attach (Standard: robocopy; Attach von lokal)
# ---------------------------------------------------------------------------
function Invoke-DatabaseMigrationDetachAttach {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$SourceServer,
        [Parameter(Mandatory)]$TargetServer,
        [Parameter(Mandatory)][string[]]$Databases,
        [Parameter(Mandatory)][string]$ExchangePath,
        [switch]$ReattachOnSource,
        [switch]$WhatIf,
        [ValidateSet('Full','Phase1','Phase2')]
        [string]$Mode = 'Full',
        [hashtable]$DbFileMap = @{},
        [System.Action[string]]$ProgressCallback
    )

    $results  = [System.Collections.Generic.List[PSObject]]::new()
    $fileMap  = [ordered]@{}

    foreach ($dbName in $Databases) {
        $stepResult = [PSCustomObject]@{
            Database    = $dbName
            Method      = 'DetachAttach'
            DetachOK    = $false
            CopyOK      = $false
            AttachOK    = $false
            ReattachOK  = $false
            DestFiles   = @()
            Error       = ''
            Duration    = 0
        }
        $sw = [System.Diagnostics.Stopwatch]::StartNew()

        try {
            # ----------------------------------------------------------------
            # PHASE 1 / FULL: Detach + robocopy auf Exchange-Pfad
            # ----------------------------------------------------------------
            if ($Mode -in 'Full','Phase1') {

                $db      = Get-DbaDatabase -SqlInstance $SourceServer -Database $dbName -ErrorAction Stop
                $dbFiles = $db.FileGroups | ForEach-Object { $_.Files } | Select-Object -ExpandProperty FileName
                $logFiles= $db.LogFiles | Select-Object -ExpandProperty FileName
                $allFiles= @($dbFiles) + @($logFiles)

                if ($ProgressCallback) { $ProgressCallback.Invoke("Detach: $dbName") }
                Write-MigrationLog -Level 'STEP' -Category 'DETACH' -Message "Detach: $dbName von $($SourceServer.Name)"

                if (-not $WhatIf) {
                    Detach-DbaDatabase -SqlInstance $SourceServer -Database $dbName -Force -ErrorAction Stop | Out-Null
                    $stepResult.DetachOK = $true
                    Write-MigrationLog -Level 'SUCCESS' -Category 'DETACH' -Message "Detach OK: $dbName"
                } else {
                    $stepResult.DetachOK = $true
                    Write-MigrationLog -Level 'INFO' -Category 'DETACH' -Message "[WHATIF] Detach: $dbName"
                }

                if ($ProgressCallback) { $ProgressCallback.Invoke("Kopiere (robocopy): $dbName") }
                $destFiles = [System.Collections.Generic.List[string]]::new()

                if (-not $WhatIf) {
                    foreach ($f in $allFiles) {
                        $dest = Copy-FileRobocopy -SourceFile $f -DestDir $ExchangePath
                        $destFiles.Add($dest)
                        Write-MigrationLog -Level 'INFO' -Category 'COPY' `
                            -Message "robocopy: $(Split-Path $f -Leaf)" -Detail $dest
                    }
                    $stepResult.CopyOK    = $true
                    $stepResult.DestFiles = $destFiles.ToArray()
                    $fileMap[$dbName]     = $destFiles.ToArray()
                } else {
                    $stepResult.CopyOK = $true
                    Write-MigrationLog -Level 'INFO' -Category 'COPY' -Message "[WHATIF] robocopy: $dbName -> $ExchangePath"
                }

                if ($ReattachOnSource -and -not $WhatIf -and $stepResult.DetachOK) {
                    if ($ProgressCallback) { $ProgressCallback.Invoke("Re-Attach Quelle: $dbName") }
                    try {
                        $reattachQuery = Build-AttachQuery -DbName $dbName -Files $allFiles
                        Invoke-DbaQuery -SqlInstance $SourceServer -Query $reattachQuery -ErrorAction Stop
                        $stepResult.ReattachOK = $true
                        Write-MigrationLog -Level 'SUCCESS' -Category 'REATTACH' -Message "Re-Attach Quelle OK: $dbName"
                    }
                    catch {
                        Write-MigrationLog -Level 'WARN' -Category 'REATTACH' `
                            -Message "Re-Attach Quelle fehlgeschlagen: $dbName" -Detail $_.Exception.Message
                    }
                }
            }

            # ----------------------------------------------------------------
            # PHASE 2 / FULL: robocopy in DATA/LOG des Ziels, dann Attach
            # ----------------------------------------------------------------
            if ($Mode -in 'Full','Phase2') {

                if ($ProgressCallback) { $ProgressCallback.Invoke("Attach: $dbName") }
                Write-MigrationLog -Level 'STEP' -Category 'ATTACH' -Message "Attach: $dbName auf $($TargetServer.Name)"

                # Quell-Dateien (auf Exchange) ermitteln
                $exFiles = if ($stepResult.DestFiles -and $stepResult.DestFiles.Count -gt 0) {
                    $stepResult.DestFiles
                } elseif ($DbFileMap.ContainsKey($dbName)) {
                    @($DbFileMap[$dbName] | ForEach-Object { $_.ToString() })
                } else {
                    @(Get-ChildItem $ExchangePath |
                        Where-Object { $_.BaseName -like "$dbName*" -and $_.Extension -match '\.(mdf|ndf|ldf)$' } |
                        Select-Object -ExpandProperty FullName)
                }

                if (-not $WhatIf -and $exFiles.Count -gt 0) {
                    # Ziel-Standardpfade ermitteln
                    $dp       = Get-DbaDefaultPath -SqlInstance $TargetServer -ErrorAction Stop
                    $dataDir  = $dp.Data
                    $logDir   = if ($dp.Log) { $dp.Log } else { $dp.Data }

                    $localFiles = [System.Collections.Generic.List[string]]::new()
                    foreach ($f in $exFiles) {
                        $targetDir = if ($f -match '\.ldf$') { $logDir } else { $dataDir }
                        $lf = Copy-FileRobocopy -SourceFile $f -DestDir $targetDir
                        $localFiles.Add($lf)
                        Write-MigrationLog -Level 'INFO' -Category 'COPY' `
                            -Message "robocopy ans Ziel: $(Split-Path $f -Leaf)" -Detail $lf
                    }

                    $attachQuery = Build-AttachQuery -DbName $dbName -Files $localFiles.ToArray()
                    Invoke-DbaQuery -SqlInstance $TargetServer -Query $attachQuery -ErrorAction Stop
                    $stepResult.AttachOK  = $true
                    $stepResult.DestFiles = $localFiles.ToArray()
                    Write-MigrationLog -Level 'SUCCESS' -Category 'ATTACH' -Message "Attach OK: $dbName"
                } else {
                    $stepResult.AttachOK = $true
                    Write-MigrationLog -Level 'INFO' -Category 'ATTACH' -Message "[WHATIF] robocopy + Attach: $dbName"
                }
            }
        }
        catch {
            $stepResult.Error = $_.Exception.Message
            Write-MigrationLog -Level 'ERROR' -Category 'DETACH-ATTACH' `
                -Message "Fehler bei $dbName" -Detail $_.Exception.Message
        }
        finally {
            $sw.Stop()
            $stepResult.Duration = [math]::Round($sw.Elapsed.TotalSeconds, 1)
            $results.Add($stepResult)
        }
    }

    $results | Add-Member -NotePropertyName '_FileMap' -NotePropertyValue $fileMap -Force
    return $results
}

# ---------------------------------------------------------------------------
# Interner Helfer: CREATE DATABASE ... FOR ATTACH Query bauen
# ---------------------------------------------------------------------------
function Build-AttachQuery {
    param(
        [string]$DbName,
        [string[]]$Files
    )
    $dataFiles = $Files | Where-Object { $_ -match '\.(mdf|ndf)$' }
    $logFiles  = $Files | Where-Object { $_ -match '\.ldf$' }

    $q  = "CREATE DATABASE [$DbName] ON "
    $q += ($dataFiles | ForEach-Object { "(FILENAME = N'$_')" }) -join ','
    if ($logFiles) {
        $q += " LOG ON "
        $q += ($logFiles | ForEach-Object { "(FILENAME = N'$_')" }) -join ','
    }
    $q += " FOR ATTACH"
    return $q
}

# ---------------------------------------------------------------------------
# POST-RESTORE (Ziel): verwaiste User reparieren + DB-Owner auf sa
# ---------------------------------------------------------------------------
function Invoke-PostRestoreCleanup {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$TargetServer,
        [Parameter(Mandatory)][string[]]$Databases,
        [switch]$WhatIf
    )

    # sa-Login robust per SID 0x01 ermitteln (funktioniert auch bei umbenanntem sa)
    $saName = $null
    try {
        $saRow = Invoke-DbaQuery -SqlInstance $TargetServer `
            -Query "SELECT name FROM sys.server_principals WHERE sid = 0x01" -ErrorAction Stop
        if ($saRow) { $saName = $saRow.name }
    }
    catch {
        Write-MigrationLog -Level 'WARN' -Category 'OWNER' `
            -Message "sa-Login konnte nicht ermittelt werden" -Detail $_.Exception.Message
    }

    foreach ($db in $Databases) {

        # 1. Verwaiste DB-User reparieren
        Write-MigrationLog -Level 'STEP' -Category 'ORPHAN' -Message "Repariere verwaiste DB-User: $db"
        if (-not $WhatIf) {
            try {
                $rep = Repair-DbaDbOrphanUser -SqlInstance $TargetServer -Database $db -ErrorAction Stop
                $cnt = @($rep).Count
                Write-MigrationLog -Level 'SUCCESS' -Category 'ORPHAN' `
                    -Message "Verwaiste User bearbeitet: $db" -Detail "Anzahl: $cnt"
            }
            catch {
                Write-MigrationLog -Level 'WARN' -Category 'ORPHAN' `
                    -Message "Reparatur verwaister User fehlgeschlagen: $db" -Detail $_.Exception.Message
            }
        } else {
            Write-MigrationLog -Level 'INFO' -Category 'ORPHAN' -Message "[WHATIF] Verwaiste User reparieren: $db"
        }

        # 2. DB-Owner auf sa setzen
        if ($saName) {
            Write-MigrationLog -Level 'STEP' -Category 'OWNER' -Message "Setze DB-Owner '$saName': $db"
            if (-not $WhatIf) {
                try {
                    $null = Set-DbaDbOwner -SqlInstance $TargetServer -Database $db -TargetLogin $saName -ErrorAction Stop
                    Write-MigrationLog -Level 'SUCCESS' -Category 'OWNER' -Message "DB-Owner gesetzt ($saName): $db"
                }
                catch {
                    Write-MigrationLog -Level 'WARN' -Category 'OWNER' `
                        -Message "DB-Owner konnte nicht gesetzt werden: $db" -Detail $_.Exception.Message
                }
            } else {
                Write-MigrationLog -Level 'INFO' -Category 'OWNER' -Message "[WHATIF] DB-Owner '$saName' setzen: $db"
            }
        }
    }
}

# ---------------------------------------------------------------------------
# VERWAISTE AD-LOGINS entfernen (geloeschte Domaenenkonten)
# Sicherheit: nur Windows-Logins, nur Domaenen-SIDs (S-1-5-21-...),
#             keine System-/sysadmin-Logins; Loeschen NUR wenn AD den SID
#             positiv NICHT aufloesen kann (IdentityNotMappedException).
# ---------------------------------------------------------------------------
function Remove-DeadAdLogin {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$TargetServer,
        [switch]$WhatIf
    )

    Write-MigrationLog -Level 'STEP' -Category 'AD-CLEAN' -Message "Suche verwaiste AD-Logins (geloeschte Konten)"

    $sysadmins = @()
    try {
        $sysadmins = @(Get-DbaServerRoleMember -SqlInstance $TargetServer -ServerRole sysadmin -ErrorAction Stop |
                       Select-Object -ExpandProperty Name)
    } catch { }

    try {
        $logins = Get-DbaLogin -SqlInstance $TargetServer -ErrorAction Stop |
                  Where-Object { $_.LoginType -in 'WindowsUser','WindowsGroup' -and -not $_.IsSystemObject }
    }
    catch {
        Write-MigrationLog -Level 'ERROR' -Category 'AD-CLEAN' `
            -Message "Logins konnten nicht gelesen werden" -Detail $_.Exception.Message
        return
    }

    $removed = 0
    foreach ($l in $logins) {
        if ($l.Name -in $sysadmins) { continue }

        # SID aus dem Login (byte[]) -> SecurityIdentifier
        $sid = $null
        try { $sid = New-Object System.Security.Principal.SecurityIdentifier(([byte[]]$l.Sid), 0) }
        catch { continue }

        # Nur Domaenenkonten (lokale/builtin ueberspringen)
        if ($sid.Value -notmatch '^S-1-5-21-') { continue }

        # AD-Existenz pruefen
        $missing = $false
        try {
            $null = $sid.Translate([System.Security.Principal.NTAccount])
        }
        catch [System.Security.Principal.IdentityNotMappedException] {
            $missing = $true
        }
        catch {
            # Transienter Fehler (DC nicht erreichbar o.ae.) -> NICHT loeschen
            Write-MigrationLog -Level 'WARN' -Category 'AD-CLEAN' `
                -Message "AD-Pruefung uebersprungen (unklar): $($l.Name)" -Detail $_.Exception.Message
            continue
        }
        if (-not $missing) { continue }

        Write-MigrationLog -Level 'WARN' -Category 'AD-CLEAN' `
            -Message "Verwaistes AD-Login (Konto geloescht): $($l.Name)"
        if (-not $WhatIf) {
            try {
                Remove-DbaLogin -SqlInstance $TargetServer -Login $l.Name -Force -ErrorAction Stop | Out-Null
                $removed++
                Write-MigrationLog -Level 'SUCCESS' -Category 'AD-CLEAN' -Message "Login entfernt: $($l.Name)"
            }
            catch {
                Write-MigrationLog -Level 'WARN' -Category 'AD-CLEAN' `
                    -Message "Login konnte nicht entfernt werden: $($l.Name)" -Detail $_.Exception.Message
            }
        } else {
            Write-MigrationLog -Level 'INFO' -Category 'AD-CLEAN' -Message "[WHATIF] Login wuerde entfernt: $($l.Name)"
        }
    }
    Write-MigrationLog -Level 'INFO' -Category 'AD-CLEAN' -Message "AD-Login-Bereinigung abgeschlossen" -Detail "Entfernt: $removed"
}

# ---------------------------------------------------------------------------
# LOGINS
# ---------------------------------------------------------------------------
function Invoke-LoginMigration {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$SourceServer,
        [Parameter(Mandatory)]$TargetServer,
        [string[]]$Logins,
        [switch]$SyncSids,
        [switch]$WhatIf
    )

    Write-MigrationLog -Level 'STEP' -Category 'LOGIN' -Message "Starte Login-Migration"
    try {
        $params = @{ Source = $SourceServer; Destination = $TargetServer; ErrorAction = 'Stop' }
        if ($Logins)   { $params['Login']    = $Logins }
        if ($SyncSids) { $params['SyncSids'] = $true }
        if ($WhatIf)   { $params['WhatIf']   = $true }

        $result = Copy-DbaLogin @params
        $ok  = ($result | Where-Object { $_.Status -eq 'Successful' }).Count
        $err = ($result | Where-Object { $_.Status -ne 'Successful' }).Count
        Write-MigrationLog -Level 'SUCCESS' -Category 'LOGIN' `
            -Message "Login-Migration abgeschlossen" -Detail "OK: $ok | Fehler: $err"
        return $result
    }
    catch {
        Write-MigrationLog -Level 'ERROR' -Category 'LOGIN' `
            -Message "Login-Migration fehlgeschlagen" -Detail $_.Exception.Message
        throw
    }
}

# ---------------------------------------------------------------------------
# LINKED SERVER
# ---------------------------------------------------------------------------
function Invoke-LinkedServerMigration {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$SourceServer,
        [Parameter(Mandatory)]$TargetServer,
        [string[]]$LinkedServers,
        [switch]$WhatIf
    )

    Write-MigrationLog -Level 'STEP' -Category 'LINKED-SRV' -Message "Starte Linked-Server-Migration"
    try {
        $params = @{ Source = $SourceServer; Destination = $TargetServer; ErrorAction = 'Stop' }
        if ($LinkedServers) { $params['LinkedServer'] = $LinkedServers }
        if ($WhatIf)        { $params['WhatIf'] = $true }

        $result = Copy-DbaLinkedServer @params
        $ok = ($result | Where-Object { $_.Status -eq 'Successful' }).Count
        Write-MigrationLog -Level 'SUCCESS' -Category 'LINKED-SRV' `
            -Message "Linked-Server-Migration abgeschlossen" -Detail "OK: $ok"
        return $result
    }
    catch {
        Write-MigrationLog -Level 'ERROR' -Category 'LINKED-SRV' `
            -Message "Linked-Server-Migration fehlgeschlagen" -Detail $_.Exception.Message
        throw
    }
}

# ---------------------------------------------------------------------------
# AGENT JOBS
# ---------------------------------------------------------------------------
function Invoke-AgentJobMigration {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$SourceServer,
        [Parameter(Mandatory)]$TargetServer,
        [string[]]$Jobs,
        [switch]$WhatIf
    )

    Write-MigrationLog -Level 'STEP' -Category 'AGENT-JOB' -Message "Starte Agent-Job-Migration"
    try {
        $params = @{ Source = $SourceServer; Destination = $TargetServer; ErrorAction = 'Stop' }
        if ($Jobs)   { $params['Job']    = $Jobs }
        if ($WhatIf) { $params['WhatIf'] = $true }

        $result = Copy-DbaAgentJob @params
        $ok = ($result | Where-Object { $_.Status -eq 'Successful' }).Count
        Write-MigrationLog -Level 'SUCCESS' -Category 'AGENT-JOB' `
            -Message "Agent-Job-Migration abgeschlossen" -Detail "OK: $ok"
        return $result
    }
    catch {
        Write-MigrationLog -Level 'ERROR' -Category 'AGENT-JOB' `
            -Message "Agent-Job-Migration fehlgeschlagen" -Detail $_.Exception.Message
        throw
    }
}

# ---------------------------------------------------------------------------
# CREDENTIALS
# ---------------------------------------------------------------------------
function Invoke-CredentialMigration {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$SourceServer,
        [Parameter(Mandatory)]$TargetServer,
        [switch]$WhatIf
    )

    Write-MigrationLog -Level 'STEP' -Category 'CREDENTIAL' -Message "Starte Credential-Migration"
    try {
        $params = @{ Source = $SourceServer; Destination = $TargetServer; ErrorAction = 'Stop' }
        if ($WhatIf) { $params['WhatIf'] = $true }
        $result = Copy-DbaCredential @params
        Write-MigrationLog -Level 'SUCCESS' -Category 'CREDENTIAL' -Message "Credential-Migration abgeschlossen"
        return $result
    }
    catch {
        Write-MigrationLog -Level 'ERROR' -Category 'CREDENTIAL' `
            -Message "Credential-Migration fehlgeschlagen" -Detail $_.Exception.Message
        throw
    }
}

# ---------------------------------------------------------------------------
# PROXIES
# ---------------------------------------------------------------------------
function Invoke-ProxyMigration {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$SourceServer,
        [Parameter(Mandatory)]$TargetServer,
        [switch]$WhatIf
    )

    Write-MigrationLog -Level 'STEP' -Category 'PROXY' -Message "Starte Proxy-Migration"
    try {
        $params = @{ Source = $SourceServer; Destination = $TargetServer; ErrorAction = 'Stop' }
        if ($WhatIf) { $params['WhatIf'] = $true }
        $result = Copy-DbaAgentProxy @params
        Write-MigrationLog -Level 'SUCCESS' -Category 'PROXY' -Message "Proxy-Migration abgeschlossen"
        return $result
    }
    catch {
        Write-MigrationLog -Level 'ERROR' -Category 'PROXY' `
            -Message "Proxy-Migration fehlgeschlagen" -Detail $_.Exception.Message
        throw
    }
}

# ---------------------------------------------------------------------------
# AUTH-VORBEREITUNG ZIEL: SQL-Logins erkennen + Mixed Mode + PBM-Policy
# ---------------------------------------------------------------------------

# Prueft, ob auf der Quelle (echte) SQL-Logins existieren, die transferiert werden.
function Test-SourceHasSqlLogins {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$SourceServer,
        [string[]]$Logins
    )
    try {
        $sql = Get-DbaLogin -SqlInstance $SourceServer -ErrorAction Stop |
               Where-Object { $_.LoginType -eq 'SqlLogin' -and -not $_.IsSystemObject }
        if ($Logins) { $sql = $sql | Where-Object { $_.Name -in $Logins } }
        $cnt = @($sql).Count
        Write-MigrationLog -Level 'INFO' -Category 'LOGIN-CHECK' -Message "SQL-Logins auf Quelle gefunden: $cnt"
        return ($cnt -gt 0)
    }
    catch {
        Write-MigrationLog -Level 'WARN' -Category 'LOGIN-CHECK' `
            -Message "SQL-Login-Pruefung fehlgeschlagen" -Detail $_.Exception.Message
        return $false
    }
}

# Schaltet das Ziel auf Mixed Mode um (+ Dienst-Neustart), wenn SQL-Logins
# transferiert werden und das Ziel aktuell nur Windows-Authentifizierung nutzt.
# Gibt die (ggf. nach Neustart neu aufgebaute) Serververbindung zurueck.
function Enable-MixedModeIfNeeded {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$TargetServer,
        [bool]$SqlLoginsPresent,
        [switch]$WhatIf
    )

    if (-not $SqlLoginsPresent) {
        Write-MigrationLog -Level 'INFO' -Category 'AUTHMODE' `
            -Message "Keine SQL-Logins zu transferieren - Auth-Modus bleibt unveraendert"
        return $TargetServer
    }

    try {
        $mode = $TargetServer.Settings.LoginMode.ToString()
        Write-MigrationLog -Level 'INFO' -Category 'AUTHMODE' -Message "Aktueller Ziel-Auth-Modus: $mode"

        # 'Integrated' = nur Windows-Authentifizierung
        if ($mode -ne 'Integrated') {
            Write-MigrationLog -Level 'INFO' -Category 'AUTHMODE' `
                -Message "Mixed Mode bereits aktiv - kein Umschalten noetig"
            return $TargetServer
        }

        if ($WhatIf) {
            Write-MigrationLog -Level 'INFO' -Category 'AUTHMODE' `
                -Message "[WHATIF] Wuerde auf Mixed Mode umstellen und SQL-Dienst neu starten"
            return $TargetServer
        }

        # 1. Auth-Modus umstellen (wirkt erst nach Neustart)
        $TargetServer.Settings.LoginMode = [Microsoft.SqlServer.Management.Smo.ServerLoginMode]::Mixed
        $TargetServer.Alter()
        Write-MigrationLog -Level 'SUCCESS' -Category 'AUTHMODE' `
            -Message "Auth-Modus auf Mixed gesetzt - Neustart erforderlich"

        # 2. SQL-Dienst neu starten (damit Mixed Mode aktiv wird)
        $full = $TargetServer.Name
        $comp = ($full -split '\\')[0]
        $inst = if ($full -match '\\') { ($full -split '\\')[1] } else { 'MSSQLSERVER' }
        Write-MigrationLog -Level 'STEP' -Category 'AUTHMODE' -Message "Starte SQL-Dienst neu" -Detail "$comp / $inst"
        $null = Restart-DbaService -ComputerName $comp -InstanceName $inst -Type Engine -Force -ErrorAction Stop
        Write-MigrationLog -Level 'SUCCESS' -Category 'AUTHMODE' -Message "SQL-Dienst neu gestartet"
        Start-Sleep -Seconds 5

        # 3. Verbindung neu aufbauen (alte ist durch den Neustart ungueltig)
        $new = Connect-DbaInstance -SqlInstance $full -TrustServerCertificate -ErrorAction Stop
        Write-MigrationLog -Level 'SUCCESS' -Category 'AUTHMODE' -Message "Verbindung nach Neustart wiederhergestellt"
        return $new
    }
    catch {
        Write-MigrationLog -Level 'ERROR' -Category 'AUTHMODE' `
            -Message "Mixed-Mode-Umstellung fehlgeschlagen" -Detail $_.Exception.Message
        return $TargetServer
    }
}

# Deaktiviert eine Policy-Based-Management-Policy (z.B. 'New_Password_Policy')
# vor dem Login-Import, falls vorhanden und aktiv.
function Disable-NamedPbmPolicy {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$TargetServer,
        [Parameter(Mandatory)][string]$PolicyName,
        [switch]$WhatIf
    )
    try {
        $pol = Get-DbaPbmPolicy -SqlInstance $TargetServer -ErrorAction Stop |
               Where-Object { $_.Name -eq $PolicyName } | Select-Object -First 1
        if (-not $pol) {
            Write-MigrationLog -Level 'INFO' -Category 'PBM' -Message "Policy '$PolicyName' nicht vorhanden - nichts zu tun"
            return
        }
        if (-not $pol.Enabled) {
            Write-MigrationLog -Level 'INFO' -Category 'PBM' -Message "Policy '$PolicyName' bereits deaktiviert"
            return
        }
        if ($WhatIf) {
            Write-MigrationLog -Level 'INFO' -Category 'PBM' -Message "[WHATIF] Policy '$PolicyName' wuerde deaktiviert"
            return
        }
        $pol.Enabled = $false
        $pol.Alter()
        Write-MigrationLog -Level 'SUCCESS' -Category 'PBM' -Message "Policy deaktiviert (vor Login-Import): $PolicyName"
    }
    catch {
        Write-MigrationLog -Level 'WARN' -Category 'PBM' `
            -Message "Policy '$PolicyName' konnte nicht deaktiviert werden" -Detail $_.Exception.Message
    }
}

Export-ModuleMember -Function Invoke-DatabaseMigrationBackupRestore,
                               Invoke-DatabaseMigrationDetachAttach,
                               Invoke-PostRestoreCleanup,
                               Remove-DeadAdLogin,
                               Test-SourceHasSqlLogins,
                               Enable-MixedModeIfNeeded,
                               Disable-NamedPbmPolicy,
                               Invoke-LoginMigration,
                               Invoke-LinkedServerMigration,
                               Invoke-AgentJobMigration,
                               Invoke-CredentialMigration,
                               Invoke-ProxyMigration,
                               Build-AttachQuery
