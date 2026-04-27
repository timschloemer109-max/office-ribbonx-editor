# CustomUI14 per PowerShell aktualisieren

`scripts/Update-CustomUI14.ps1` aktualisiert den Office-2010-RibbonX-Part
`/customUI/customUI14.xml` in einer `.xlsm`- oder `.xlam`-Datei. Es wird nur Windows
PowerShell 5.1 benoetigt; Office RibbonX Editor und .NET-SDK sind fuer die
Ausfuehrung nicht erforderlich.

## Aufruf

```powershell
.\scripts\Update-CustomUI14.ps1 `
  -WorkbookPath "C:\Pfad\Datei.xlsm" `
  -DefinitionPath "C:\Pfad\ribbon-def.txt"
```

Standardmaessig wird neben der Eingabedatei eine neue Datei mit dem Suffix
`.ribbon.<ext>` geschrieben (z. B. `.ribbon.xlsm` oder `.ribbon.xlam`). Das Original bleibt unveraendert.

Falls die lokale Execution Policy direkte Skriptaufrufe blockiert, kann der
Aufruf explizit ueber PowerShell gestartet werden:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\Update-CustomUI14.ps1 `
  -WorkbookPath "C:\Pfad\Datei.xlsm" `
  -DefinitionPath "C:\Pfad\ribbon-def.txt"
```

```powershell
.\scripts\Update-CustomUI14.ps1 `
  -WorkbookPath "C:\Pfad\Datei.xlsm" `
  -DefinitionPath "C:\Pfad\ribbon-def.txt" `
  -OutputPath "C:\Pfad\Datei.mit-ribbon.xlsm"
```

Mit `-ForceInPlace` wird die angegebene Arbeitsmappe direkt geaendert:

```powershell
.\scripts\Update-CustomUI14.ps1 `
  -WorkbookPath "C:\Pfad\Datei.xlsm" `
  -DefinitionPath "C:\Pfad\ribbon-def.txt" `
  -ForceInPlace
```

## Definitionsdatei

Die Datei ist eine einfache, Semikolon-getrennte Textdatei. Leere Zeilen sind
erlaubt. Metadaten stehen vor dem Tabellenkopf.

```text
#tab_id=tabCustom
#tab_label=Meine Tools
#group_id=grpMain
#group_label=Aktionen
#insert_after_mso=TabHome

id;label;size;icon;onAction
btnExport;Export;large;FileSaveAs;ExportMakro
btnSync;Synchronisieren;normal;RefreshAll;SyncMakro
```

Pflicht-Metadaten:

- `tab_id`
- `tab_label`
- `group_id`
- `group_label`

Optionale Metadaten:

- `insert_after_mso`

Tabellenfelder:

- `id`: Pflicht, eindeutig innerhalb der Definitionsdatei
- `label`: Pflicht
- `size`: Pflicht, nur `normal` oder `large`
- `icon`: Pflicht, wird als `imageMso` geschrieben
- `onAction`: Pflicht, Name des VBA-Callbacks

Eine durchsuchbare Uebersicht aktueller Microsoft-365-`imageMso`-Namen liegt
in `docs/OfficeImageMso-Catalog.csv`; Hinweise zum Aktualisieren stehen in
`docs/OfficeImageMso-Catalog.md`.

Doppelte `id`-Werte sind ein Fehler. Das Skript bricht ab, bevor die Datei
geaendert wird.

## Verhalten

- Fehlt `/customUI/customUI14.xml`, wird der Part erstellt.
- Eine Package-Relationship auf `customUI/customUI14.xml` wird bei Bedarf
  erstellt.
- Ein vorhandenes `/customUI/customUI.xml` bleibt unangetastet.
- Existiert der konfigurierte Tab, die Gruppe oder ein Button bereits, werden
  die Attribute aktualisiert.
- Buttons in der konfigurierten Gruppe werden mit der Definitionsdatei
  synchronisiert: IDs, die nicht mehr in `ribbon-def.txt` enthalten sind,
  werden aus der Gruppe entfernt.
- Makros werden nicht erzeugt. `onAction` referenziert nur vorhandene oder
  spaeter anzulegende VBA-Prozeduren.

## Pruefen

Der optionale Helfer zeigt den vorhandenen CustomUI14-Part und die Buttons an:

```powershell
.\scripts\Test-InspectCustomUI14.ps1 -WorkbookPath "C:\Pfad\Datei.ribbon.xlsm"
```

## Typische Fehler

Wenn die Datei in Excel geoeffnet ist, kann das Paket nicht geschrieben werden.
Excel schliessen und den Befehl erneut ausfuehren.

Bei ungueltiger Definition meldet das Skript die Zeilennummer, zum Beispiel bei
fehlenden Feldern, `size=big` oder doppelten IDs.
