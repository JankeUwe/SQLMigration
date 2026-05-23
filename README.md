# SQLMigration

PowerShell WinForms-Tool zur zweiphasigen SQL Server Datenbankmigrationen — entwickelt von [dtcSoftware](https://www.powershelldba.de) (Uwe Janke).

## Übersicht

`SQLMigration` ist eine grafische PowerShell-Anwendung (WinForms) für SQL Server Migrationen in komplexen Netzwerkumgebungen. Das Tool unterstützt direkte Migrationen (Quell- und Zielserver erreichbar) sowie zweiphasige Migrationen über eine Zustandsdatei — für Umgebungen ohne direkte Netzwerkverbindung zwischen den Servern.

**Version:** 1.1 | **Getestet auf:** Windows Server 2022 / SQL Server 2016–2022

## Features

- **WinForms GUI**: Rollenauswahl (Quelle / Ziel / Automatisch) beim Start, zeigt nur die jeweils relevante Seite
- **Zweiphasige Migration**: Phase 1 auf dem Quellserver (Backup/Export), Phase 2 auf dem Zielserver (Restore/Import) — State-Übergabe per JSON-Datei über Exchange-Pfad
- **Automatische Szenario-Erkennung**: Zielserver TCP-erreichbar → Direct-Modus; nicht erreichbar → TwoPhase-Modus
- **Mehrere Migrationsmethoden**:
  - Backup / Restore
  - Detach / Attach
- **Migrierbare Objekte**: Datenbanken, Logins, SQL Agent Jobs, Linked Server, SSIS-Pakete
- **State Management**: JSON-basierte Zustandsdatei für getrennte Netzwerke
- **WhatIf-Modus**: Vollständige Simulation ohne tatsächliche Änderungen
- **Strukturiertes Logging**: Alle Schritte werden protokolliert
- **dbaTools-Integration**: Verbindungsmanagement und SQL-Operationen

## Voraussetzungen

| Anforderung | Mindestversion | Hinweis |
|-------------|---------------|---------|
| Windows Server | 2016 | Quell- und Zielserver |
| SQL Server | 2016 | Quelle: ab 2016 — Ziel: 2016 bis 2025 |
| PowerShell | 5.1 | |
| dbaTools | 2.0 | |

## Verwendung

```powershell
# Interaktiv starten (Rollenauswahl per Dialog)
.\SQL-Migration.ps1

# Rolle per Parameter vorgeben
.\SQL-Migration.ps1 -Role Source
.\SQL-Migration.ps1 -Role Target
.\SQL-Migration.ps1 -Role Auto

# Konfigurationsdatei angeben
.\SQL-Migration.ps1 -ConfigFile "C:\Migration\migration.config.json"
```

### Betriebsmodi

| Modus | Beschreibung |
|-------|-------------|
| `Source` | Nur Phase 1 — Backup/Export auf Quellserver |
| `Target` | Nur Phase 2 — Restore/Import auf Zielserver |
| `Auto` | Erkennt automatisch: Zustandsdatei vorhanden → Target, sonst → Source |

## Projektstruktur

```
SQLMigration/
├── SQL-Migration.ps1           # Hauptskript (WinForms GUI + Steuerlogik)
├── migration.config.json       # Konfiguration (Pfade, Optionen)
└── modules/
    ├── Connect-SqlServer.psm1  # Verbindungsmanagement (Windows/SQL-Auth)
    ├── Get-SqlObjects.psm1     # Inventarisierung migrierbarer Objekte
    ├── Invoke-Migration.psm1   # Migrationsdurchführung (Backup/Restore, Detach/Attach)
    ├── Invoke-MigrationState.psm1  # State-Management (JSON-Zustandsdatei)
    └── Write-MigrationLog.psm1 # Logging
```

## Konfiguration

Die `migration.config.json` steuert:
- Exchange-Pfad für zweiphasige Migrationen
- Standard-Migrationsmethode
- Logging-Pfad
- WhatIf-Modus

## Mehr Informationen

- **Dokumentation:** [powershelldba.de/tools/sql-migrations-automation](https://www.powershelldba.de/tools/sql-migrations-automation)
- **Website:** [www.powershelldba.de](https://www.powershelldba.de)
- **PowerShell Gallery:** [powershellgallery.com/profiles/JankeUwe](https://www.powershellgallery.com/profiles/JankeUwe)
- **Alle Projekte:** [github.com/JankeUwe](https://github.com/JankeUwe)
- Entwickler: Uwe Janke, Senior SQL Server DBA | dtcSoftware
