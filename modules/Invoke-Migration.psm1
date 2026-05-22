# =============================================================================
# Modul: Invoke-Migration.psm1
# Zweck: Migrationslogik fuer alle SQL-Objekte
#
# Strategie Backup/Restore:
#   1. Versuche direkt auf UNC-Pfad (Dienstkonto muss Zugriff haben)
#   2. Bei Fehler: Backup lokal, dann PowerShell-Copy auf Exchange-Pfad
#      (laeuft unter Admin-Credentials -> kein Dienstkonto-Problem)
#
# Strategie Detach/Attach:
#   Copy-Item laeuft immer unter Admin-Credentials -> kein Problem
# =============================================================================

# ---------------------------------------------------------------------------
# Interner Hilfer: lokalen Backup-Pfad sicherstellen
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
# DATENBANKEN - Backup / Restore
# Strategie (Option B):
#   Versuch 1: Backup direkt in ExchangePath (Dienstkonto braucht Zugriff)
#   Versuch 2: Backup lokal -> Copy-Item nach ExchangePath (Admin-Credentials)
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
        # Im TwoPhase-Modus: nur Backup+Copy (Phase 1) oder nur Restore (Phase 2)
        [ValidateSet('Full','Phase1','Phase2')]
        [string]$Mode = 'Full',
        [System.Action[string]]$ProgressCallback,
        # Callback wenn Fallback auf lokales Backup greift (GUI-Hinweis)
        [System.Action[string]]$FallbackCallback
    )

    $results = [System.Collections.Generic.List[PSObject]]::new()

    foreach ($dbName in $Databases) {
        $stepResult = [PSCustomObject]@{
            Database     = $dbName
            Method       = 'BackupRestore'
            BackupOK     = $false
            CopyOK       = $false
            RestoreOK    = $false
            BackupFile   = ''
            UsedLocalBackup = $false
            Error        = ''
            Duration     = 0
        }
        $sw = [System.Diagnostics.Stopwatch]::StartNew()

        try {
            # ----------------------------------------------------------------
            # PHASE 1 / FULL: Backup
            # ----------------------------------------------------------------
            if ($Mode -in 'Full','Phase1') {

                if ($ProgressCallback) { $ProgressCallback.Invoke("Backup: $dbName") }
                Write-MigrationLog -Level 'STEP' -Category 'BACKUP' `
                    -Message "Starte Backup: $dbName"

                if (-not $WhatIf) {

                    # --- Versuch 1: direkt in Exchange-Pfad ---
                    $backupOK = $false
                    try {
                        Write-MigrationLog -Level 'INFO' -Category 'BACKUP' `
                            -Message "Versuch 1: Backup direkt nach Exchange-Pfad" `
                            -Detail $ExchangePath

                        $bpResult = Backup-DbaDatabase `
                            -SqlInstance      $SourceServer `
                            -Database         $dbName `
                            -BackupDirectory  $ExchangePath `
                            -CompressBackup   $WithCompression.IsPresent `
                            -CopyOnly         $CopyOnly.IsPresent `
                            -ErrorAction      Stop

                        $stepResult.BackupFile = $bpResult.BackupPath
                        $stepResult.BackupOK   = $true
                        $stepResult.CopyOK     = $true   # kein separater Copy noetig
                        $backupOK = $true

                        Write-MigrationLog -Level 'SUCCESS' -Category 'BACKUP' `
                            -Message "Backup direkt OK: $dbName" `
                            -Detail $bpResult.BackupPath
                    }
                    catch {
                        Write-MigrationLog -Level 'WARN' -Category 'BACKUP' `
                            -Message "Versuch 1 fehlgeschlagen (Dienstkonto ohne UNC-Zugriff?)" `
                            -Detail $_.Exception.Message
                    }

                    # --- Versuch 2: lokal + Copy ---
                    if (-not $backupOK) {
                        Write-MigrationLog -Level 'WARN' -Category 'BACKUP' `
                            -Message "Fallback: Backup lokal, dann Copy (Dienstkonto hat keinen UNC-Zugriff)" `
                            -Detail $LocalBackupPath

                        # GUI-Hinweis ausloesen
                        if ($FallbackCallback) {
                            $FallbackCallback.Invoke(
                                "HINWEIS: Dienstkonto hat keinen Zugriff auf Exchange-Pfad.`n" +
                                "Backup wird lokal erstellt ($LocalBackupPath)`nund danach unter Admin-Credentials kopiert.")
                        }

                        Ensure-LocalBackupPath -Path $LocalBackupPath

                        $bpResult = Backup-DbaDatabase `
                            -SqlInstance      $SourceServer `
                            -Database         $dbName `
                            -BackupDirectory  $LocalBackupPath `
                            -CompressBackup   $WithCompression.IsPresent `
                            -CopyOnly         $CopyOnly.IsPresent `
                            -ErrorAction      Stop

                        $localFile = $bpResult.BackupPath
                        $stepResult.BackupOK        = $true
                        $stepResult.UsedLocalBackup = $true

                        Write-MigrationLog -Level 'SUCCESS' -Category 'BACKUP' `
                            -Message "Backup lokal OK: $dbName" -Detail $localFile

                        # Copy unter Admin-Credentials
                        if ($ProgressCallback) { $ProgressCallback.Invoke("Kopiere: $dbName") }
                        $destFile = Join-Path $ExchangePath (Split-Path $localFile -Leaf)
                        Copy-Item -Path $localFile -Destination $destFile -Force -ErrorAction Stop

                        $stepResult.BackupFile = $destFile
                        $stepResult.CopyOK     = $true

                        Write-MigrationLog -Level 'SUCCESS' -Category 'COPY' `
                            -Message "Backup kopiert nach Exchange-Pfad" -Detail $destFile

                        # Lokale Datei nach erfolgreichem Copy aufraumen
                        try {
                            Remove-Item -Path $localFile -Force -ErrorAction Stop
                            Write-MigrationLog -Level 'INFO' -Category 'CLEANUP' `
                                -Message "Lokale Backup-Datei bereinigt" -Detail $localFile
                        }
                        catch {
                            Write-MigrationLog -Level 'WARN' -Category 'CLEANUP' `
                                -Message "Lokale Backup-Datei konnte nicht geloescht werden (manuell pruefen)" `
                                -Detail $localFile
                        }
                    }

                    if ($VerifyBackup) {
                        Write-MigrationLog -Level 'INFO' -Category 'BACKUP' `
                            -Message "Verifiziere Backup: $dbName"
                        $null = Test-DbaLastBackup -SqlInstance $SourceServer `
                            -Database $dbName -ErrorAction SilentlyContinue
                    }

                } else {
                    $stepResult.BackupOK = $true
                    $stepResult.CopyOK   = $true
                    $stepResult.BackupFile = Join-Path $ExchangePath "$dbName.bak"
                    Write-MigrationLog -Level 'INFO' -Category 'BACKUP' `
                        -Message "[WHATIF] Backup wuerde ausgefuehrt: $dbName"
                }
            }

            # ----------------------------------------------------------------
            # PHASE 2 / FULL: Restore
            # ----------------------------------------------------------------
            if ($Mode -in 'Full','Phase2') {

                if ($ProgressCallback) { $ProgressCallback.Invoke("Restore: $dbName") }
                Write-MigrationLog -Level 'STEP' -Category 'RESTORE' `
                    -Message "Starte Restore: $dbName auf $($TargetServer.Name)"

                if (-not $WhatIf) {

                    # Im Phase2-Modus kommt BackupFile aus dem Parameter
                    $restorePath = if ($stepResult.BackupFile) {
                        $stepResult.BackupFile
                    } else {
                        $ExchangePath   # dbaTools sucht passende .bak
                    }

                    # --- Versuch 1: Restore direkt vom Exchange-Pfad ---
                    $restoreOK = $false
                    try {
                        Write-MigrationLog -Level 'INFO' -Category 'RESTORE' `
                            -Message "Versuch 1: Restore direkt vom Exchange-Pfad" `
                            -Detail $restorePath

                        $null = Restore-DbaDatabase `
                            -SqlInstance  $TargetServer `
                            -Path         $restorePath `
                            -DatabaseName $dbName `
                            -WithReplace  `
                            -ErrorAction  Stop

                        $stepResult.RestoreOK = $true
                        $restoreOK = $true
                        Write-MigrationLog -Level 'SUCCESS' -Category 'RESTORE' `
                            -Message "Restore direkt OK: $dbName"
                    }
                    catch {
                        Write-MigrationLog -Level 'WARN' -Category 'RESTORE' `
                            -Message "Versuch 1 fehlgeschlagen (Dienstkonto ohne UNC-Zugriff?)" `
                            -Detail $_.Exception.Message
                    }

                    # --- Versuch 2: Copy lokal + Restore lokal ---
                    if (-not $restoreOK) {
                        Write-MigrationLog -Level 'WARN' -Category 'RESTORE' `
                            -Message "Fallback: Backup lokal kopieren, dann Restore (Dienstkonto hat keinen UNC-Zugriff)" `
                            -Detail $LocalBackupPath

                        # GUI-Hinweis ausloesen
                        if ($FallbackCallback) {
                            $FallbackCallback.Invoke(
                                "HINWEIS: Dienstkonto hat keinen Zugriff auf Exchange-Pfad.`n" +
                                "Backup-Datei wird lokal kopiert ($LocalBackupPath)`nund danach der Restore ausgefuehrt.")
                        }

                        Ensure-LocalBackupPath -Path $LocalBackupPath

                        $srcFile = if ($stepResult.BackupFile) {
                            $stepResult.BackupFile
                        } else {
                            # Suche passende .bak im Exchange-Pfad
                            Get-ChildItem $ExchangePath -Filter "$dbName*.bak" |
                                Sort-Object LastWriteTime -Descending |
                                Select-Object -First 1 -ExpandProperty FullName
                        }

                        $localDest = Join-Path $LocalBackupPath (Split-Path $srcFile -Leaf)
                        Copy-Item -Path $srcFile -Destination $localDest -Force -ErrorAction Stop

                        Write-MigrationLog -Level 'SUCCESS' -Category 'COPY' `
                            -Message "Backup lokal kopiert fuer Restore" -Detail $localDest

                        $null = Restore-DbaDatabase `
                            -SqlInstance  $TargetServer `
                            -Path         $localDest `
                            -DatabaseName $dbName `
                            -WithReplace  `
                            -ErrorAction  Stop

                        $stepResult.RestoreOK       = $true
                        $stepResult.UsedLocalBackup = $true
                        Write-MigrationLog -Level 'SUCCESS' -Category 'RESTORE' `
                            -Message "Restore lokal OK: $dbName"

                        # Lokale Kopie nach erfolgreichem Restore aufraumen
                        try {
                            Remove-Item -Path $localDest -Force -ErrorAction Stop
                            Write-MigrationLog -Level 'INFO' -Category 'CLEANUP' `
                                -Message "Lokale Restore-Kopie bereinigt" -Detail $localDest
                        }
                        catch {
                            Write-MigrationLog -Level 'WARN' -Category 'CLEANUP' `
                                -Message "Lokale Restore-Kopie konnte nicht geloescht werden (manuell pruefen)" `
                                -Detail $localDest
                        }
                    }

                } else {
                    $stepResult.RestoreOK = $true
                    Write-MigrationLog -Level 'INFO' -Category 'RESTORE' `
                        -Message "[WHATIF] Restore wuerde ausgefuehrt: $dbName"
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
# DATENBANKEN - Detach / Attach
# Copy-Item laeuft unter Admin-Credentials -> kein Dienstkonto-Problem
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
        # Phase2: DbFileMap aus Zustandsdatei { DbName -> @(Zielpfade) }
        [hashtable]$DbFileMap = @{},
        [System.Action[string]]$ProgressCallback
    )

    $results  = [System.Collections.Generic.List[PSObject]]::new()
    $fileMap  = [ordered]@{}   # wird fuer Zustandsdatei zurueckgegeben

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
            # PHASE 1 / FULL: Detach + Copy auf Exchange-Pfad
            # ----------------------------------------------------------------
            if ($Mode -in 'Full','Phase1') {

                # Dateipfade vom Quell-Server ermitteln
                $db      = Get-DbaDatabase -SqlInstance $SourceServer -Database $dbName -ErrorAction Stop
                $dbFiles = $db.FileGroups | ForEach-Object { $_.Files } |
                           Select-Object -ExpandProperty FileName
                $logFiles= $db.LogFiles | Select-Object -ExpandProperty FileName
                $allFiles= @($dbFiles) + @($logFiles)

                if ($ProgressCallback) { $ProgressCallback.Invoke("Detach: $dbName") }
                Write-MigrationLog -Level 'STEP' -Category 'DETACH' `
                    -Message "Detach: $dbName von $($SourceServer.Name)"

                if (-not $WhatIf) {
                    Detach-DbaDatabase -SqlInstance $SourceServer `
                        -Database $dbName -Force -ErrorAction Stop | Out-Null
                    $stepResult.DetachOK = $true
                    Write-MigrationLog -Level 'SUCCESS' -Category 'DETACH' `
                        -Message "Detach OK: $dbName"
                } else {
                    $stepResult.DetachOK = $true
                    Write-MigrationLog -Level 'INFO' -Category 'DETACH' `
                        -Message "[WHATIF] Detach: $dbName"
                }

                # Copy auf Exchange-Pfad (Admin-Credentials)
                if ($ProgressCallback) { $ProgressCallback.Invoke("Kopiere: $dbName") }
                $destFiles = [System.Collections.Generic.List[string]]::new()

                if (-not $WhatIf) {
                    foreach ($f in $allFiles) {
                        $dest = Join-Path $ExchangePath (Split-Path $f -Leaf)
                        Copy-Item -Path $f -Destination $dest -Force -ErrorAction Stop
                        $destFiles.Add($dest)
                        Write-MigrationLog -Level 'INFO' -Category 'COPY' `
                            -Message "Kopiert: $(Split-Path $f -Leaf)" -Detail $dest
                    }
                    $stepResult.CopyOK   = $true
                    $stepResult.DestFiles = $destFiles.ToArray()
                    $fileMap[$dbName]    = $destFiles.ToArray()
                } else {
                    $stepResult.CopyOK = $true
                    Write-MigrationLog -Level 'INFO' -Category 'COPY' `
                        -Message "[WHATIF] Kopieren: $dbName nach $ExchangePath"
                }

                # Re-Attach auf Quelle
                if ($ReattachOnSource -and -not $WhatIf -and $stepResult.DetachOK) {
                    if ($ProgressCallback) { $ProgressCallback.Invoke("Re-Attach Quelle: $dbName") }
                    try {
                        $reattachQuery = Build-AttachQuery -DbName $dbName -Files $allFiles
                        Invoke-DbaQuery -SqlInstance $SourceServer `
                            -Query $reattachQuery -ErrorAction Stop
                        $stepResult.ReattachOK = $true
                        Write-MigrationLog -Level 'SUCCESS' -Category 'REATTACH' `
                            -Message "Re-Attach Quelle OK: $dbName"
                    }
                    catch {
                        Write-MigrationLog -Level 'WARN' -Category 'REATTACH' `
                            -Message "Re-Attach Quelle fehlgeschlagen: $dbName" `
                            -Detail $_.Exception.Message
                    }
                }
            }

            # ----------------------------------------------------------------
            # PHASE 2 / FULL: Attach auf Ziel
            # ----------------------------------------------------------------
            if ($Mode -in 'Full','Phase2') {

                if ($ProgressCallback) { $ProgressCallback.Invoke("Attach: $dbName") }
                Write-MigrationLog -Level 'STEP' -Category 'ATTACH' `
                    -Message "Attach: $dbName auf $($TargetServer.Name)"

                # Dateien: aus Phase1-Ergebnis oder aus DbFileMap (Zustandsdatei)
                $filesToAttach = if ($stepResult.DestFiles -and $stepResult.DestFiles.Count -gt 0) {
                    $stepResult.DestFiles
                } elseif ($DbFileMap.ContainsKey($dbName)) {
                    # Zustandsdatei liefert PSCustomObject-Array -> in strings wandeln
                    @($DbFileMap[$dbName] | ForEach-Object { $_.ToString() })
                } else {
                    # Fallback: Exchange-Pfad nach passenden Dateien durchsuchen
                    @(Get-ChildItem $ExchangePath |
                        Where-Object { $_.BaseName -like "$dbName*" } |
                        Select-Object -ExpandProperty FullName)
                }

                if (-not $WhatIf -and $filesToAttach.Count -gt 0) {
                    $attachQuery = Build-AttachQuery -DbName $dbName -Files $filesToAttach
                    Invoke-DbaQuery -SqlInstance $TargetServer `
                        -Query $attachQuery -ErrorAction Stop
                    $stepResult.AttachOK = $true
                    Write-MigrationLog -Level 'SUCCESS' -Category 'ATTACH' `
                        -Message "Attach OK: $dbName"
                } else {
                    $stepResult.AttachOK = $true
                    Write-MigrationLog -Level 'INFO' -Category 'ATTACH' `
                        -Message "[WHATIF] Attach: $dbName"
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

    # FileMap am Ergebnis anhaengen (fuer Zustandsdatei)
    $results | Add-Member -NotePropertyName '_FileMap' -NotePropertyValue $fileMap -Force
    return $results
}

# ---------------------------------------------------------------------------
# Interner Hilfer: CREATE DATABASE ... FOR ATTACH Query bauen
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

    Write-MigrationLog -Level 'STEP' -Category 'LOGIN' `
        -Message "Starte Login-Migration"
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

    Write-MigrationLog -Level 'STEP' -Category 'LINKED-SRV' `
        -Message "Starte Linked-Server-Migration"
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

    Write-MigrationLog -Level 'STEP' -Category 'AGENT-JOB' `
        -Message "Starte Agent-Job-Migration"
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

    Write-MigrationLog -Level 'STEP' -Category 'CREDENTIAL' `
        -Message "Starte Credential-Migration"
    try {
        $params = @{ Source = $SourceServer; Destination = $TargetServer; ErrorAction = 'Stop' }
        if ($WhatIf) { $params['WhatIf'] = $true }
        $result = Copy-DbaCredential @params
        Write-MigrationLog -Level 'SUCCESS' -Category 'CREDENTIAL' `
            -Message "Credential-Migration abgeschlossen"
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

    Write-MigrationLog -Level 'STEP' -Category 'PROXY' `
        -Message "Starte Proxy-Migration"
    try {
        $params = @{ Source = $SourceServer; Destination = $TargetServer; ErrorAction = 'Stop' }
        if ($WhatIf) { $params['WhatIf'] = $true }
        $result = Copy-DbaAgentProxy @params
        Write-MigrationLog -Level 'SUCCESS' -Category 'PROXY' `
            -Message "Proxy-Migration abgeschlossen"
        return $result
    }
    catch {
        Write-MigrationLog -Level 'ERROR' -Category 'PROXY' `
            -Message "Proxy-Migration fehlgeschlagen" -Detail $_.Exception.Message
        throw
    }
}

Export-ModuleMember -Function Invoke-DatabaseMigrationBackupRestore,
                               Invoke-DatabaseMigrationDetachAttach,
                               Invoke-LoginMigration,
                               Invoke-LinkedServerMigration,
                               Invoke-AgentJobMigration,
                               Invoke-CredentialMigration,
                               Invoke-ProxyMigration,
                               Build-AttachQuery
