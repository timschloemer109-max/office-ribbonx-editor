# Office imageMso catalog

Die Datei `OfficeImageMso-Catalog.csv` enthaelt eine durchsuchbare Liste von
Office Fluent UI Control Identifiers, die als `imageMso`-Namen fuer RibbonX
hilfreich sind.

Quelle ist Microsofts offizielles GitHub-Repository:

https://github.com/OfficeDev/office-fluent-ui-command-identifiers

Verwendet wird standardmaessig:

`Microsoft 365/Current Channel`

## Spalten

- `ImageMso`: Name fuer `imageMso="..."`
- `Application`: Office-Anwendung oder Outlook-Kontextdatei aus der Quelle
- `ControlType`: Typ des Controls, z. B. `button`, `toggleButton`, `menu`
- `TabSet`, `Tab`, `GroupOrContextMenu`: Ribbon- oder Kontextmenue-Position
- `ParentControl`, `SecondaryParentControl`: uebergeordnete Controls, falls vorhanden
- `PolicyId`: Policy-ID aus Microsofts Liste
- `Description`: kompakter Suchtext aus Anwendung, Typ und Ribbon-Kontext

Die Microsoft-Dateien enthalten keine freien Icon-Beschreibungen oder
Tooltip-Texte. Die Beschreibung in der CSV ist deshalb aus den vorhandenen
Kontextspalten zusammengesetzt.

## Aktualisieren

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\Update-OfficeImageMsoCatalog.ps1
```

Optional kann ein anderer Microsoft-365-Kanal angegeben werden:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\Update-OfficeImageMsoCatalog.ps1 `
  -ChannelPath "Microsoft 365/Semi-Annual Enterprise Channel"
```

## Suchen

Die CSV laesst sich in Excel, VS Code oder per PowerShell filtern:

```powershell
Import-Csv .\docs\OfficeImageMso-Catalog.csv |
  Where-Object { $_.ImageMso -like '*Save*' -or $_.Description -like '*Clipboard*' } |
  Select-Object -First 20 ImageMso,Application,ControlType,Description
```
