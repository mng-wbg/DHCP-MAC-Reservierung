# DHCP MAC-Reservierung

Startprojekt fuer Windows Server 2022 Standard mit PowerShell 5.1, WinForms und DHCPServer-Modul.

## Struktur

```text
DhcpMacReservierung
├─ Start-DhcpMacReservierung.ps1
├─ Settings.psd1
├─ services
│  ├─ CsvMacListService.ps1
│  └─ DhcpReservationService.ps1
└─ views
   └─ DhcpMacReservierungForm.ps1
```

## Start

PowerShell als Administrator starten:

```powershell
Set-ExecutionPolicy RemoteSigned -Scope Process
.\Start-DhcpMacReservierung.ps1
```

## Wichtig

- `Settings.psd1` anpassen.
- DHCPServer-Modul muss installiert/verfuegbar sein.
- Das Tool schreibt die CSV vor und nach der Reservierung.
- Vor dem Speichern wird automatisch ein Backup der CSV erstellt.
- Der bekannte Fehler bei `Get-DhcpServerv4Lease -ScopeId ... -IPAddress ...` ist hier bereits vermieden.
