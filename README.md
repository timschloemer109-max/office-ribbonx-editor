# Office RibbonX Editor - Fork

Dieses Repository ist unser Fork des Office RibbonX Editors. Grundlage ist das
Originalprojekt von Fernando Andreu:

- Original: https://github.com/fernandreu/office-ribbonx-editor
- Dieser Fork: https://github.com/timschloemer109-max/office-ribbonx-editor
- Lizenz: MIT, siehe [LICENSE](LICENSE)

Der Fork muss nicht dieselbe README wie das Original behalten. Wichtig ist, dass
die Herkunft und die Lizenz erhalten bleiben. Diese README beschreibt daher vor
allem unsere Anpassungen und wo die relevanten Dateien liegen.

## Was ist Office RibbonX Editor?

Office RibbonX Editor ist ein Windows-Tool zum Bearbeiten der Custom-UI-Teile in
Office-Dateien wie `.xlsm`, `.xlam`, `.xltm`, `.docm`, `.dotm`, `.pptm` und
`.ppam`. Die Anwendung oeffnet Office-Dateien als Open-XML-Pakete, zeigt die
enthaltenen `customUI.xml`- und `customUI14.xml`-Teile an und schreibt
Aenderungen wieder in die Datei.

Das Originalprojekt enthaelt die WPF-Anwendung, den XML-Editor, Schema-Pruefung,
Samples, Ressourcen, Uebersetzungen und Build-/Test-Projekte.

## Unsere Aenderungen

In diesem Fork wurden vor allem Werkzeuge rund um `customUI/customUI14.xml`
ergaenzt und bestehende Speicherablaeufe stabilisiert.

- PowerShell-Werkzeuge zum Aktualisieren, Pruefen und Entfernen von
  `customUI14.xml` in Office-Makrodateien
- Standalone-Paket unter `customui14-update-package`, das ohne die komplette
  WPF-Anwendung benutzt werden kann
- kleines Fenster-Tool fuer interaktive Datei-Auswahl, Einspielen, Loeschen und
  Wahl zwischen In-Place-Ueberschreiben und neuer Ausgabedatei
- Unterstuetzung fuer weitere Makrodateitypen, unter anderem `.xlam` und `.xltm`
- vollstaendige `customUI14.txt`/`.xml`-Vorlage als bevorzugtes Format fuer
  neue Projekte
- alte `ribbon-def.txt`-Tabellenvariante bleibt als einfache Legacy-Variante
  erhalten
- Logging und Nachpruefung, ob `customUI/customUI14.xml` nach dem Schreiben
  wirklich im Zielpaket vorhanden ist
- Fixes fuer Save-As, Reload-on-Save und echtes Ueberschreiben, damit geaenderte
  Custom-UI-Inhalte nicht versehentlich durch Originalinhalte ersetzt werden

## Wo liegt was?

| Pfad | Inhalt |
| --- | --- |
| `customui14-update-package/` | Standalone-Variante zum Einspielen oder Entfernen von `customUI14.xml` ohne Office RibbonX Editor |
| `customui14-update-package/README.md` | genaue Anleitung fuer das Standalone-Paket |
| `customui14-update-package/CustomUI14-App.cmd` | bequemer Starter fuer die kleine Fenster-Anwendung |
| `customui14-update-package/CustomUI14-App.ps1` | PowerShell-Fenster-Anwendung fuer Auswahl, Log und Ausfuehrung |
| `customui14-update-package/Set-CustomUI14Xml.ps1` | schreibt eine komplette `customUI14.txt`/`.xml` in eine Office-Datei |
| `customui14-update-package/Remove-CustomUIXml.ps1` | entfernt `customUI.xml`, `customUI14.xml` oder beide Varianten |
| `customui14-update-package/Apply-CustomUI14Interactive.ps1` | interaktive Konsolenvariante |
| `customui14-update-package/customUI14-template.txt` | kleine XML-Vorlage fuer neue Ribbons |
| `customui14-update-package/ribbon-def.txt` | Legacy-Buttonliste im Tabellenformat |
| `src/OfficeRibbonXEditor/` | WPF-Hauptanwendung |
| `src/OfficeRibbonXEditor.Common/` | gemeinsame Open-XML-/Office-Paketlogik |
| `src/OfficeRibbonXEditor.CommandLine/` | Kommandozeilenprojekt des Editors |
| `scripts/` | Hilfsskripte und aeltere CustomUI14-Werkzeuge |
| `docs/` | Zusatzdokumentation, ImageMso-Katalog und PowerShell-Notizen |
| `tests/` | Unit-, Functional-, Integration- und UI-Tests |
| `build/` | Build-, Installer- und Pipeline-Dateien |

## Standalone: `customui14-update-package`

Der Ordner `customui14-update-package` enthaelt die Standalone-Variante. Dieser
Ordner kann auf einen Ziel-PC kopiert werden, ohne die komplette Anwendung zu
installieren oder zu bauen. Auf dem Ziel-PC wird nur Windows PowerShell 5.1
benoetigt.

Empfohlener Start:

```powershell
.\CustomUI14-App.cmd
```

Typischer Ablauf:

1. Office-Datei auswaehlen, zum Beispiel `.xlsm`, `.xlam`, `.xltm`, `.docm`,
   `.dotm`, `.pptm` oder `.ppam`
2. Modus waehlen: CustomUI einspielen oder CustomUI loeschen
3. Beim Einspielen eine `customUI14.txt` oder `customUI14.xml` auswaehlen
4. Speichermodus waehlen: Originaldatei ueberschreiben oder neue Ausgabedatei
5. Log pruefen

Details stehen in [customui14-update-package/README.md](customui14-update-package/README.md).

## Empfohlenes CustomUI-Format

Fuer neue Projekte ist `customUI14.txt` bzw. `customUI14.xml` das bevorzugte
Format. Es enthaelt das echte RibbonX-XML und kann direkt in die Office-Datei
eingespielt werden.

`ribbon-def.txt` bleibt nur fuer einfache Buttonlisten erhalten. Sobald Tabs,
Gruppen, Screentips, Spezialattribute oder mehrere Ribbon-Bereiche gebraucht
werden, sollte `customUI14.txt` verwendet werden.

## Bauen und Testen

Die WPF-Anwendung nutzt .NET 10. Wenn das lokale `.dotnet`-Verzeichnis vorhanden
ist, kann die Solution so gebaut oder getestet werden:

```powershell
.\.dotnet\dotnet.exe build .\OfficeRibbonXEditor.slnx
.\.dotnet\dotnet.exe test .\OfficeRibbonXEditor.slnx
```

Das Standalone-Paket unter `customui14-update-package` braucht zum Ausfuehren
kein .NET SDK.

## Upstream-Dokumentation

Ausfuehrliche Informationen zur urspruenglichen Anwendung, Releases,
Uebersetzungen, Code Signing und den allgemeinen RibbonX-Editor-Funktionen
stehen im Originalprojekt:

https://github.com/fernandreu/office-ribbonx-editor

Nuetzliche RibbonX-Quellen:

- https://www.rondebruin.nl/win/s2/win001.htm
- https://github.com/OfficeDev/office-fluent-ui-command-identifiers
