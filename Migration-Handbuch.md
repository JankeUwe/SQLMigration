# SQL Server Migration – Arbeitshandbuch & Checkliste

Praxisleitfaden für Administratoren zum Tool **SQL-Migration.ps1** (WinForms-GUI auf dbatools-Basis).
Stand: v1.1 · powershelldba.de – Janke

---

## 1. Überblick

Das Tool migriert Datenbanken und Server-Objekte (Logins, DB-User, Linked Server, Agent-Jobs,
Credentials, Proxies) von einem **Quell-** auf einen **Ziel-SQL-Server**.

Es kennt zwei Betriebsarten, die beim Start abgefragt werden (oder per `-Role` vorgegeben):

| Rolle | Bedeutung |
|-------|-----------|
| **Quelle (Phase 1)** | Backup bzw. Detach + Kopieren der Dateien auf den Exchange-Pfad. Schreibt die Zustandsdatei. |
| **Ziel (Phase 2)** | Kopieren vom Exchange-Pfad + Restore bzw. Attach. Liest die Zustandsdatei. |
| **Automatisch** | Zustandsdatei im Exchange-Pfad vorhanden → Ziel, sonst → Quelle. |

**Szenarien:**

- **Direct** – Zielserver ist vom Quellserver erreichbar (ein Durchlauf möglich).
- **TwoPhase** – Quelle und Ziel sind getrennt: erst Phase 1 auf der Quelle, dann Phase 2 auf dem Ziel,
  Datenträgeraustausch über den Exchange-Pfad (UNC-Share).

> **Wichtig:** **Datenbanken** und **Logins** werden in beiden Modi migriert – Logins über ein
> exportiertes Skript (`Export-DbaLogin` → `Invoke-DbaQuery`), das auch bei getrennten Domänen
> funktioniert. **Agent-Jobs, Linked Server, Credentials und Proxies** benötigen weiterhin den
> **Direct-Modus** (gleichzeitige Verbindung zu beiden Servern, `Copy-Dba*`) und sind im TwoPhase-Modus
> manuell bzw. via Direct nachzuziehen.

---

## 2. Voraussetzungen

- [ ] **Windows PowerShell 5.1** (oder höher) auf Quell- und Zielserver.
- [ ] **dbatools** installiert (`Install-Module dbatools`).
- [ ] Ausführung als **Administrator** (Dateizugriff, Dienste, ggf. Detach/Attach).
- [ ] Auf beiden SQL-Instanzen **sysadmin**-Rechte des ausführenden Kontos.
- [ ] **Exchange-Pfad** (UNC-Share) von beiden Servern erreichbar; ausreichend Speicherplatz für die Backups.
- [ ] **Lokaler Backup-Pfad** als Fallback vorhanden (wenn das SQL-Dienstkonto keinen UNC-Zugriff hat).
- [ ] Netzwerk/Firewall: SQL-Port (Standard 1433) im Direct-Modus, SMB (445) für den Exchange-Pfad.
- [ ] Bei SQL 2022 / selbstsignierten Zertifikaten: **TrustServerCertificate** aktiviert lassen.

**Konfiguration** (`config\migration.config.json`, Standardwerte):

| Schlüssel | Standard |
|-----------|----------|
| `DefaultExchangePath` | `\\exchange-server\SQLMigration\Backups` |
| `DefaultLocalBackupPath` | `F:\Daten\SQL\Backup` |
| `DefaultLogPath` | `C:\SQLMigration\Logs` |
| `StateFileName` | `_migration_state.json` |
| `BackupCompression` / `VerifyBackup` / `CopyOnlyBackup` | `true` |
| `DefaultMigrationMethod` | `BackupRestore` |

---

## 3. Vor der Migration – Pre-Check (Quelle)

- [ ] Wartungsfenster abgestimmt, Anwendungen informiert/gestoppt.
- [ ] Aktuelles, **unabhängiges** Voll-Backup aller betroffenen DBs vorhanden (Sicherheitsnetz).
- [ ] Liste der zu migrierenden Datenbanken und Objekte festgelegt.
- [ ] Zielserver: Version ≥ Quellversion, ausreichend Plattenplatz, Pfade (DATA/LOG) vorbereitet.
- [ ] Kompatibilitätsgrad / Collation-Anforderungen geklärt.
- [ ] Exchange-Pfad leer bzw. keine alte `_migration_state.json` von einem früheren Lauf vorhanden.
- [ ] **Testlauf mit WhatIf** eingeplant.

---

## 4. Ablauf – Phase 1 (Quell-Server)

1. `Start-SQLMigration.cmd` als Administrator starten → Rolle **Quelle** wählen.
2. **Server\Instanz** ist mit dem aktuellen Rechner vorbelegt – prüfen/anpassen, **Auth** wählen, **Verbinden**.
3. Objekte werden geladen; in den Reitern (Datenbanken, Logins, …) die zu migrierenden Einträge ankreuzen
   (Buttons *Alle/Keine* je Reiter bzw. *Alle Tabs*).
4. Unten **Methode** wählen:
   - **Backup / Restore** (empfohlen, Standard) – DB bleibt online.
   - **Detach / Attach** – DB wird offline genommen; optional *Re-Attach nach Detach* auf der Quelle.
5. **Exchange-Pfad** und **Lokalen Backup-Pfad** prüfen.
6. Optional **WhatIf** aktivieren für einen Trockenlauf.
7. **PHASE 1 STARTEN** → Zusammenfassung bestätigen.
8. Nach Abschluss: Erfolgsmeldung beachten. Es wurde eine **Zustandsdatei** im Exchange-Pfad erzeugt.

> **Standard-Strategie (kein Dienstkonto-/UNC-Problem):** Das Backup wird zuerst in das
> **lokale Standard-Backup-Verzeichnis des Servers** geschrieben (dort hat das SQL-Dienstkonto
> immer Schreibrechte) und anschließend per **robocopy** auf den Exchange-Pfad kopiert – der
> Kopiervorgang läuft im **Admin-Kontext**, daher ist kein UNC-Zugriff des Dienstkontos nötig.
> Die lokale Sicherung verbleibt als Kopie im Server-Backup-Verzeichnis.

---

## 5. Ablauf – Phase 2 (Ziel-Server)

1. `Start-SQLMigration.cmd` als Administrator starten → Rolle **Ziel** (oder *Automatisch*).
2. Das Tool liest die **Zustandsdatei**; Methode und Objektauswahl sind vorbelegt und gesperrt.
3. **Zielserver** verbinden (vorbelegt mit dem aktuellen Rechner).
4. **Exchange-/Lokalen Pfad** prüfen.
5. **PHASE 2 STARTEN** → robocopy vom Exchange-Pfad in das lokale Verzeichnis des Ziels, dann
   Restore/Attach der Datenbanken.
6. **Automatische Nachbearbeitung am Ziel** (läuft direkt nach dem Restore/Attach, respektiert WhatIf):
   - **Login-Migration (Objekttyp *Logins*)** – domänenübergreifend/zweistufig tauglich:
     - In **Phase 1** werden die Logins per `Export-DbaLogin` als **CREATE-LOGIN-Skript**
       (inkl. SID + gehashtem Passwort) auf den Exchange-Pfad exportiert (`migration_logins.sql`).
       Das ersetzt `Copy-DbaLogin`, das beide Server gleichzeitig sehen müsste.
     - In **Phase 2** wird das Skript **batchweise** (an `GO`) per `Invoke-DbaQuery` ausgeführt;
       einzelne Batches dürfen scheitern (z. B. Windows-Logins fremder Domänen) ohne den Rest zu stoppen.
     - **Vor** dem Import: Erkennt das Tool **SQL-Logins** und steht das Ziel auf **nur Windows-Auth**,
       wird **Mixed Mode** aktiviert und der **SQL-Dienst neu gestartet** (Verbindung wird automatisch
       neu aufgebaut – bestehende Verbindungen zum Ziel brechen kurz ab). Die Policy
       **`New_Password_Policy`** wird – falls vorhanden – **deaktiviert**.
     - **Nach** dem Import wird `New_Password_Policy` wieder **aktiviert**.
   - Verwaiste DB-User werden repariert (`Repair-DbaDbOrphanUser`).
   - DB-Owner wird auf **sa** gesetzt (per SID `0x01` ermittelt – funktioniert auch bei umbenanntem sa).
   - Verwaiste **AD-Logins** (gelöschte Domänenkonten) werden entfernt – nur wenn der Objekttyp
     *Logins* angehakt ist. Sicherheits­regeln: nur Windows-Logins mit Domänen-SID, **keine**
     System-/sysadmin-Logins, Löschung nur wenn AD den SID **positiv nicht** auflösen kann.
7. Nach Abschluss: Zieldatenbanken prüfen (siehe Post-Check).

---

## 6. Post-Migration – Abnahme (Ziel)

- [ ] Alle erwarteten Datenbanken vorhanden, Status **ONLINE**.
- [ ] `DBCC CHECKDB` ohne Fehler.
- [ ] Recovery-Modell und Kompatibilitätsgrad korrekt.
- [ ] **Owner** der Datenbanken = **sa** (wird automatisch in Phase 2 gesetzt) – stichprobenartig prüfen.
- [ ] **Verwaiste Benutzer** repariert (automatisch in Phase 2) – stichprobenartig prüfen.
- [ ] **Verwaiste AD-Logins** entfernt (automatisch in Phase 2, wenn *Logins* angehakt) – Log prüfen.
- [ ] Logins inkl. SIDs vorhanden (per Skript angelegt) – Windows-Logins fremder Domänen ggf. im Log prüfen.
- [ ] Bei SQL-Logins: Ziel steht auf **Mixed Mode** (wurde bei Bedarf automatisch umgestellt + Dienst neu gestartet).
- [ ] Policy `New_Password_Policy` wieder **aktiv** (wird nach dem Import automatisch reaktiviert) – stichprobenartig prüfen.
- [ ] Agent-Jobs / Linked Server / Credentials / Proxies vorhanden und lauffähig (ggf. Direct-Migration).
- [ ] Anwendungs-Connectionstrings auf den neuen Server umgestellt.
- [ ] Funktionstest der Anwendung.
- [ ] Quelldatenbanken erst nach erfolgreicher Abnahme deaktivieren/abhängen.

---

## 7. Logs, Nachweise & Aufräumen

- **Log:** `C:\SQLMigration\Logs\SQL-Migration_*.log` – Button **[Log]** in der GUI.
- **CSV-Protokoll:** Button **[CSV]**.
- **Lokale Backup-Dateien** werden **nicht** automatisch gelöscht – nach Abnahme manuell bereinigen.
- Zustandsdatei `_migration_state.json` nach Abschluss aus dem Exchange-Pfad entfernen, bevor eine
  neue Migration startet (sonst wird fälschlich der Ziel-Modus erkannt).

---

## 8. Rollback

- **Vor Restore/Attach** (Ziel): Es wurde noch nichts verändert → einfach abbrechen.
- **Backup/Restore-Methode:** Die Quelle bleibt unverändert online → Rollback = Anwendung auf Quelle belassen.
- **Detach/Attach-Methode:** Quelle wurde abgehängt. Mit *Re-Attach* (Option) oder manuellem `Attach`
  der ursprünglichen Dateien zurücksetzen.
- Im Zweifel: das vor der Migration erstellte **unabhängige Voll-Backup** zurückspielen.

---

## 9. Troubleshooting

| Symptom | Ursache / Lösung |
|---------|------------------|
| „dbaTools ist nicht installiert" | `Install-Module dbatools` ausführen. |
| Verbindungsfehler / Zertifikatfehler | **TrustServerCertificate** aktivieren; Server/Instanz, Port, Firewall prüfen. |
| „Exchange-Pfad nicht erreichbar" | UNC-Pfad und Berechtigungen prüfen; ggf. lokalen Backup-Pfad nutzen. |
| Backup wird lokal erstellt | **So gewollt** (Standard): Backup lokal → robocopy auf den Share im Admin-Kontext. Die lokale Kopie verbleibt im Server-Backup-Verzeichnis. |
| robocopy-Fehler (ExitCode ≥ 8) | Ziel-/Quellpfad bzw. Berechtigungen des **ausführenden Admins** prüfen; Plattenplatz prüfen. |
| Phase 2 findet keine Zustandsdatei | Phase 1 wurde nicht (erfolgreich) ausgeführt oder Exchange-Pfad weicht ab. |
| Verwaiste Benutzer nach Restore | Werden in Phase 2 automatisch repariert; bei Bedarf erneut `Repair-DbaDbOrphanUser` ausführen. |
| AD-Login fälschlich behalten/entfernt | Bereinigung löscht nur bei **positiv** nicht auflösbarem Domänen-SID; bei DC-Störung wird übersprungen (Log: „AD-Pruefung uebersprungen"). |
| Logins fehlen nach TwoPhase | Login-Skript prüfen (`migration_logins.sql` im Exchange-Pfad); einzelne Batches können scheitern (Log „LOGIN-IMPORT"). Windows-Logins fremder Domänen lassen sich nicht anlegen. |
| Jobs/LS/Cred/Proxy fehlen nach TwoPhase | Erwartet – nur im **Direct-Modus** migrierbar; manuell nachziehen. |
| Detaillierte Fehlermeldung | Der **Fehler-Dialog** zeigt die komplette Exception-Kette (kopierbar) + Logdatei. |

---

## 10. Schnell-Checkliste (zum Abhaken)

**Quelle:** ☐ Backup-Sicherheitsnetz ☐ Verbinden ☐ Objekte wählen ☐ Methode/Pfade ☐ WhatIf-Test
☐ Phase 1 starten ☐ Zustandsdatei + Dateien am Exchange-Pfad

**Ziel:** ☐ Verbinden ☐ Pfade prüfen ☐ Phase 2 starten ☐ Restore/Attach ok

**Abnahme:** ☐ DBs online ☐ CHECKDB ☐ Owner ☐ Orphans ☐ Logins/Jobs ☐ Connectionstrings ☐ App-Test
☐ Aufräumen (lokale Backups, Zustandsdatei)
