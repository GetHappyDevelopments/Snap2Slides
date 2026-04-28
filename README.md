# Image2PPTSlicer

Kleine Dialoganwendung, die ein Screenshot-/Rasterbild mit PowerPoint-Folien in eine editierbare `.pptx`-Datei umwandelt.

## Start

```powershell
python app.py
```

Oder unter Windows per Doppelklick auf `start.bat`.

## Installation

Die notwendigen Pakete sind in `requirements.txt` aufgefuehrt:

```powershell
python -m pip install --user -r requirements.txt
```

## Was erzeugt wird

- erkannte Einzel-Folien werden zu separaten PowerPoint-Folien
- OCR-Texte werden als editierbare Textfelder eingefuegt
- erkannte visuelle Flaechen wie Fotos, Logos, Icons oder farbige Boxen werden als einzelne Bildobjekte eingefuegt

Die Rekonstruktion ist OCR-basiert. Je sauberer und hoeher aufgeloest die Vorlage ist, desto besser werden Textpositionen und Texterkennung.

## Bedienung

- Bild auswaehlen, Zielpfad pruefen und `PPTX erzeugen` starten
- waehrend der Umwandlung zeigt die Statusleiste den aktuellen Schritt
- nach erfolgreichem Export kann die erzeugte PowerPoint direkt ueber `Oeffnen` gestartet werden
- Tastaturkuerzel: `Strg+O` fuer Bildauswahl, `Strg+S` fuer Zielpfad, `Enter` zum Starten

Das Programm nutzt `assets/app_icon.ico` als Windows-/Taskbar-Icon.
