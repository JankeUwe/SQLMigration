# SQLMigration Changelog

All notable changes to this project will be documented in this file.

## [1.3.0.0] - 2026-10-30

### Features
- **Final step after the database migration** (Direct and Detour phase 2)
  - New `Invoke-MigrationFinalize`: runs `Invoke-sqmDatabaseStandardization` (orphaned users, users without login, compatibility level, TARGET_RECOVERY_TIME, owner sa) on the migrated databases, now AFTER logins/objects so orphans can actually be re-mapped
  - If the target is the primary of an Always On availability group: `Add-sqmDatabaseToAG` adds the databases (exactly one AG is picked automatically; with several AGs the step is skipped with a hint)
  - New `Get-SqmSqlToolSource` checks for sqmSQLTool (loaded, installed, or source next to the tool); the start confirmation dialog and status bar warn up front if it is missing and the final step will run
  - Without sqmSQLTool: fallback to the previous cleanup (`Invoke-PostRestoreCleanup`), AG step skipped with a warning

### GUI
- Title bar now uses the role colour (blue = source, green = target); log path no longer overlaps the version label
- Quellserver and Zielserver are two separate rows, always source on top; the connected server's row carries the Auth fields, the counterpart row carries "Verbindung pruefen" and Verfahren
- TrustServerCertificate checkbox aligned with the server text boxes; Exchange and local backup path boxes aligned in one column; grid rows white; unselected tabs light grey
- Source and target server each have their own login (Windows or SQL-Login) and their own Verbinden button on the server row; the counterpart's Verbinden checks reachability and login, and the Direct path connects to it with those credentials
- Logins: system logins (sa, ##...##, NT SERVICE / NT AUTHORITY incl. German NT-AUTORITAET, BUILTIN) are no longer listed and are excluded from the Direct (Copy-DbaLogin) and detour (Export-DbaLogin) transfer
- Final step passes the target's SQL-Login on to sqmSQLTool; Test-SqlConnection now really connects (Test-DbaConnection reported False for working SQL logins)
- Direct procedure: title and start button no longer say "Phase 1" (now "Migration: Backup + Restore"); the button shows "Migration fertig" after the run
- Live backup/restore progress in the GUI: percent, database and remaining minutes of running BACKUP/RESTORE statements on source and (Direct) target, polled every ~2 s from sys.dm_exec_requests, shown next to the progress bar
- Copy progress: while a backup file is copied to the target, the GUI shows "[Kopie] file NN % (x / y GB, MB/s, remaining minutes)"; the percentage comes from robocopy's own output (the destination file size is useless, robocopy preallocates it)

## [1.2.0.0] - 2026-06-21

### Documentation
- **Script variant documentation** (commit 556dfc6)
  - Handbook: TwoPhase migrates ALL object types via script (not just Direct)
  - Secrets TODO notes and troubleshooting corrected; documentation version 1.2
  - README: Credentials/Proxies + Database users added, SSIS (not supported) removed
  - Migration procedures (Direct/Detour) and automatic post-processing documented; version 1.2

### Features
- **Direct path with admin-share UNC** (commit e5a74fc)
  - Direct (one-run) path now correctly addresses target via admin-share UNC
  - New `Invoke-DirectDatabaseTransfer`: backup/detach locally on source, robocopy files to target's admin share (\\target\D$\...), restore/attach using target's LOCAL paths
  - New `ConvertTo-AdminShareUnc` function for cross-machine path conversion
  - Handbook notes updated

- **Auto-detect connection mode and set target backup path** (commit effcdd9 + 4814a5d)
  - Connection panel now includes counterpart server field and "Verbindung prüfen" (Test-ServerReachable) button
  - Procedure selector: Auto/Direct/Detour (Auto picks Direct when counterpart is reachable, otherwise two-phase detour)
  - Auto-set Exchange path to target backup directory via admin-share UNC when counterpart is found
  - New `Invoke-DirectMigration`: single-pass migration when both servers visible from source
    - Databases via robust backup/robocopy/restore (Mode Full)
    - Logins/Jobs/Linked Servers/Credentials/Proxies via Copy-Dba* functions
    - Automatic Mixed Mode enablement, policy handling, and post-restore cleanup
  - Handbook updated with procedure details and counterpart field documentation

- **Script-based migration for Agent Jobs, Linked Servers, Credentials, Proxies** (commit c67fb11)
  - Cross-domain compatible: export to script on source, import on target (no simultaneous connection required)
  - New `Export-MigrationAgentJobs` / `Export-MigrationProxies` (via Export-DbaScript)
  - New `Export-MigrationLinkedServers` / `Export-MigrationCredentials` (decrypt secrets on source)
  - New generic `Import-MigrationScriptFile` (batch split on GO, resilient per batch)
  - Undecryptable secrets recorded in `migration_secrets_TODO.txt` for manual follow-up
  - State file tracks ObjectScripts paths; Phase 1 exports, Phase 2 imports in dependency order
  - Old "Direct only" warning removed

- **Login migration via Export-DbaLogin; PBM policy management** (commit 3064493)
  - Two-phase capable: avoids requirement for simultaneous visibility of both servers
  - Phase 1 (source): `Export-MigrationLogins` writes CREATE LOGIN script (SID + hashed password) to `migration_logins.sql`
  - Phase 2 (target): `Import-MigrationLogins` runs script batch-by-batch via Invoke-DbaQuery
  - New `Set-NamedPbmPolicyState` (replaces Disable-NamedPbmPolicy); handles New_Password_Policy
  - Phase 2 order: Mixed Mode → disable policy → import logins → re-enable policy → remove dead AD logins
  - Logins removed from "Direct-only" warning; handbook updated

- **Target auto-prep: Mixed Mode, PBM policy, SQL login detection** (commit 39a06d4)
  - New `Test-SourceHasSqlLogins`: detects non-system SQL logins; result stored in state file
  - New `Enable-MixedModeIfNeeded`: switches target from Windows-only to Mixed auth via SMO, restarts SQL engine
  - New `Disable-NamedPbmPolicy`: disables New_Password_Policy before login import if present
  - Both functions are WhatIf-aware
  - Wired into Phase 2 (target) under Logins object; handbook updated

- **Backup/Restore and Detach/Attach with robocopy** (commit 3cec03f)
  - Default strategy changed to local-first (avoids SQL service account UNC permission issues)
  - Phase 1: backup to server's default directory, robocopy to Exchange path
  - Phase 2: robocopy from Exchange path to target's local backup directory, restore from local
  - Detach/Attach also uses robocopy; Phase 2 copies to target's DATA/LOG directories, attaches from there
  - New `Invoke-PostRestoreCleanup` (Phase 2): repairs orphaned DB users, sets owner to sa (by SID 0x01)
  - New `Remove-DeadAdLogin` (Phase 2): removes Windows logins whose domain account was deleted (SID-only, safety checks)
  - No more direct-UNC attempt / permission fallback dialog
  - Handbook updated with strategy, automatic post-restore steps, and troubleshooting

### Bug Fixes
- **Fix GUI hang on startup** (commit 63d645a)
  - TabControl.ItemSize setter raised SizeChanged event, causing infinite recursion in FitTabWidth handler
  - Fixed with re-entrancy guard; form now completes quickly
  - Suppress "unapproved verbs" warning on startup by importing migration modules with `-DisableNameChecking`

- **Fix GUI ListView headers hidden by button panel** (commit 491e3e9)
  - TabPage dock order: ListView (Fill) must be child index 0 to prevent overlap
  - Set explicitly to show column headers
  - TabControl strip now full-width and dark themed (owner-draw tab items stretched to full width)

- **Fix Get-SqlDbUsers AddRange type error** (commit 62bd699)
  - `$result.AddRange()` failed on Object[] (not IEnumerable[PSObject])
  - Cast to [PSObject[]] to support both single user and array

- **Fix broken migration calls (inline comments on backtick-continuation)** (commit ae1017e)
  - Inline comments on backtick-continuation lines caused parameter parsing failures
  - All Invoke-DatabaseMigrationBackupRestore/DetachAttach calls fixed
  - Every migration previously failed at runtime

### UI/UX Improvements
- **GUI enhancements** (commit ae1017e)
  - Window aspect ratio: 1600x900 (minimum 1100x650)
  - Server\Instance prefilled with current machine name on start
  - Title shows "powershelldba.de - Janke (c) 2025-YY" (year dynamic)
  - TabControl headers owner-drawn to match dark theme (selected = role colour)
  - Added comprehensive Migration-Handbuch.md (admin runbook and checklists)

### Documentation Additions
- **Added Migration-Handbuch.md** (commit ae1017e)
  - Detailed admin runbook and operation checklists
  - Troubleshooting and procedure documentation

- **SEO: Documentation links in README** (commit e3f44c9)
  - Added documentation links to README for better discoverability

### Technical Details
- **Start-SQLMigration.cmd wrapper** (commit 91f25c6)
  - Copies all files (including modules/, config/) to C:\ProgramData\SQLMigration
  - Starts SQL-Migration.ps1 as Administrator via UAC
  - Optional role parameter: Source / Target (passed to PS1)
  - Without parameter: role selection via dialog

- **Auto-cleanup of backup files and fallback callback** (commit aeb3c05)
  - Invoke-Migration.psm1: FallbackCallback parameter in Invoke-DatabaseMigrationBackupRestore
  - Local backup file automatically deleted after successful Copy/Restore
  - Fallback-Log level upgraded from INFO to WARN
  - SQL-Migration.ps1: FallbackCallback wired into Phase 1 and Phase 2
  - Fallback triggers MessageBox warning + status bar turns yellow
  - Old manual cleanup instructions removed

## [1.1.0.0] - 2026-05-18

### Initial Release
- **MIT License added** (commit 69c9b95)

### Features
- **PowerShell WinForms-based SQL Server migration tool** (commit e5fba4d)
  - Two-phase migration modes: Direct and TwoPhase (for separated networks)
  - Multiple migration methods supported
  - JSON-based state management for multi-phase coordination
  - Supports migration of:
    - Databases (backup/restore or detach/attach)
    - Logins
    - Agent Jobs
    - Linked Servers
    - Credentials
    - Proxies
  - Developed by dtcSoftware (Uwe Janke)

---

## Format Notes

- Dates follow ISO 8601 (YYYY-MM-DD)
- Version numbers follow [Semantic Versioning](https://semver.org/)
- Commit references shown in parentheses for traceability
