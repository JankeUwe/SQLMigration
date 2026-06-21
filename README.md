# SQLMigration

PowerShell WinForms-Tool zur zweiphasigen SQL Server Datenbankmigrationen — entwickelt von [dtcSoftware](https://www.powershelldba.de) (Uwe Janke).

## Übersicht

`SQLMigration` ist eine grafische PowerShell-Anwendung (WinForms) für SQL Server Migrationen in komplexen Netzwerkumgebungen. Das Tool unterstützt direkte Migrationen (Quell- und Zielserver erreichbar) sowie zweiphasige Migrationen über eine Zustandsdatei — für Umgebungen ohne direkte Netzwerkverbindung zwischen den Servern.

**Version:** 1.2 | **Getestet auf:** Windows Server 2022 / SQL Server 2016–2022

## Features

- **WinForms GUI**: Rollenauswahl (Quelle / Ziel / Automatisch) beim Start, zeigt nur die jeweils relevante Seite. Objektauswahl je Typ über Reiter (Buttons *Alle/Keine* je Reiter bzw. *Alle Tabs*).
- **Zwei Verfahren mit automatischer Erkennung** (TCP-Erreichbarkeit der Gegenstelle; per *Verfahren* `Auto`/`Direkt`/`Umweg` übersteuerbar):
  - **Direkt** – Quelle und Ziel gleichzeitig erreichbar → **ein Durchlauf** vom Quellserver; Objekte direkt via `Copy-Dba*`. Transport: lokales Backup → **robocopy in die Admin-Freigabe des Ziels** (UNC), Restore über den lokalen Zielpfad.
  - **Umweg (TwoPhase)** – Quelle/Ziel getrennt (z. B. verschiedene Domänen): Phase 1 Backup + **Skript-Export** aller Objekte, Phase 2 Restore + **Skript-Import** (`Invoke-DbaQuery`). State-Übergabe per JSON-Zustandsdatei über den Exchange-Pfad.
- **Migrationsmethoden**: **Backup / Restore** (Standard, DB bleibt online) oder **Detach / Attach** (optional *Re-Attach* auf der Quelle).
- **Migrierbare Objekte**: Datenbanken, **Logins**, DB-User, **Linked Server**, **Agent Jobs**, **Credentials**, **Proxies** — in **beiden** Verfahren (Direct via `Copy-Dba*`, TwoPhase via Skript-Export/-Import).
- **Domänenübergreifend**: Logins via `Export-DbaLogin` → `Invoke-DbaQuery` (inkl. SID + Passwort-Hash); Jobs/LS/Cred/Proxy via `Export-Dba*` / `Export-DbaScript`. **Secrets-TODO**: nicht entschlüsselbare Credential-/Linked-Server-Passwörter werden zum manuellen Nachtragen aufgelistet.
- **Automatische Nachbearbeitung am Ziel**: verwaiste DB-User reparieren, DB-Owner → `sa` (per SID `0x01`), verwaiste **AD-Logins** entfernen, bei SQL-Logins **Mixed Mode** aktivieren (+ SQL-Dienst-Neustart) und Policy `New_Password_Policy` aus/ein.
- **WhatIf-Modus**: vollständige Simulation ohne tatsächliche Änderungen.
- **State Management**: JSON-basierte Zustandsdatei für getrennte Netzwerke.
- **Strukturiertes Logging**: alle Schritte protokolliert (Log + CSV in der GUI).
- **dbatools-Integration**: Verbindungsmanagement und SQL-Operationen.

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
