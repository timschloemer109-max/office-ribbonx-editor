# CustomUI14-Update-Paket

Dieses Paket aktualisiert `customUI/customUI14.xml` in Office-Makrodateien ohne Office RibbonX Editor. Auf dem Ziel-PC wird nur Windows PowerShell 5.1 benoetigt.

## Empfohlenes Vorlagenformat

Fuehrendes Format ist ab jetzt `customUI14.txt` bzw. `customUI14.xml`.

Warum:

- es ist das echte RibbonX-XML und kann 1:1 in Office-Dateien eingespielt werden
- es funktioniert gleich fuer Excel und Word
- mehrere Tabs/Gruppen, Screentips und Spezialattribute bleiben erhalten
- vorhandene CustomUI aus einer Datei kann direkt als Vorlage weitergepflegt werden

`ribbon-def.txt` bleibt als einfache Tabellen-Variante fuer kleine Buttonlisten erhalten, ist aber nicht mehr das Standardformat fuer DOF-Projekte.

## Dateien

- `Apply-CustomUI14Interactive.ps1`: interaktive Datei-Auswahl fuer Office-Datei und `customUI14.txt`.
- `CustomUI14-App.cmd`: startet eine kleine Fenster-Anwendung fuer Auswahl und Ausfuehrung.
- `CustomUI14-App.ps1`: Fenster-Anwendung mit Datei-Auswahl, Ausgabepfad und Log.
- `Set-CustomUI14Xml.ps1`: spielt eine komplette `customUI14.txt`/`.xml` in eine Office-Datei ein.
- `customUI14-template.txt`: kleine XML-Vorlage fuer neue Ribbons.
- `Apply-RibbonUpdate.ps1`: alter Wrapper fuer `ribbon-def.txt`.
- `Update-CustomUI14.ps1`: altes Buttonlisten-Update-Skript.
- `Test-InspectCustomUI14.ps1`: zeigt den vorhandenen CustomUI14-Part und Buttons an.
- `ribbon-def.txt`: einfache Button-Definition im Tabellenformat.
- `customUI14-original.txt`: Original-XML als Referenz.

## Standard: interaktive Auswahl

Am bequemsten:

```powershell
.\CustomUI14-App.cmd
```

Oder direkt per PowerShell:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Apply-CustomUI14Interactive.ps1
```

Danach:

1. Office-Datei auswaehlen (`.xlsm`, `.xlam`, `.docm`, `.dotm`, `.pptm`, `.ppam`)
2. `customUI14.txt` oder `customUI14.xml` auswaehlen
3. Ausgabedatei auswaehlen

## Standard: neue Datei per Parameter erzeugen

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Set-CustomUI14Xml.ps1 `
  -OfficePath "C:\Pfad\Datei.dotm" `
  -CustomUiPath "C:\Pfad\customUI14.txt"
```

Das erzeugt neben der Originaldatei automatisch `Datei.customui.dotm`.

## Direkt in der Originaldatei aendern

Office vorher schliessen.

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Set-CustomUI14Xml.ps1 `
  -OfficePath "C:\Pfad\Datei.dotm" `
  -CustomUiPath "C:\Pfad\customUI14.txt" `
  -ForceInPlace
```

## Expliziten Ausgabe-Pfad verwenden

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Set-CustomUI14Xml.ps1 `
  -OfficePath "C:\Pfad\Datei.dotm" `
  -CustomUiPath "C:\Pfad\customUI14.txt" `
  -OutputPath "C:\Pfad\Datei.neu.dotm"
```

## Legacy: `ribbon-def.txt`

Die alte Tabellenvariante bleibt nutzbar, wenn nur eine einfache Gruppe von Buttons aktualisiert werden soll.

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Apply-RibbonUpdate.ps1 `
  -WorkbookPath "C:\Pfad\Datei.xlsm"
```

In `ribbon-def.txt` die gewuenschte Zeile anpassen. Format:

```text
id;label;size;icon;onAction
```

Beispiel:

```text
btnSetzeBlattschutz;setze Blattschutz;large;HighImportance;Setze_Blattschutz_Ribbon
```

Wichtig: Die `id` gleich lassen, wenn ein vorhandener Button aktualisiert werden soll. Eine neue `id` erzeugt einen neuen Button.

## Pruefen

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Test-InspectCustomUI14.ps1 `
  -WorkbookPath "C:\Pfad\Datei.customui.dotm"
```

## Hinweis zu alten Screentips

Die alte `ribbon-def.txt`-Variante setzt nur `id`, `label`, `size`, `imageMso` und `onAction`. Fuer alles Weitere bitte `customUI14.txt` verwenden.
