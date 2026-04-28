# Image2PPTSlicer

Kleine Dialoganwendung, die ein Screenshot-/Rasterbild mit PowerPoint-Folien in eine editierbare `.pptx`-Datei umwandelt.

## Start

```powershell
python launcher.py
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
- erkannte visuelle Flaechen wie Fotos, Logos, Icons, Diagramme oder farbige Boxen werden als einzelne Bildobjekte eingefuegt
- abgegrenzte Bildbereiche werden als eigene PowerPoint-Objekte exportiert; es wird kein Screenshot als Folien-Hintergrund verwendet
- grosse zusammenhaengende Bildbereiche bleiben ein einzelnes Objekt und werden nicht in kleine Fragmente zerlegt
- erkannter Text wird aus den Bildobjekten entfernt und danach als editierbares Textfeld neu eingefuegt
- zusammengehoerige OCR-Zeilen werden zu Textbloecken gruppiert; Bullet-Zeilen werden als editierbare Listenzeilen rekonstruiert
- Text wird pixelgenauer aus Bildobjekten entfernt, auch bei hellem Text auf dunkleren Bildbereichen

Die Rekonstruktion ist OCR-basiert. Je sauberer und hoeher aufgeloest die Vorlage ist, desto besser werden Textpositionen und Texterkennung.

## Bedienung

- Bild auswaehlen, Zielpfad pruefen und `PPTX erzeugen` starten
- beim Start zeigt Snap2Slides mindestens 3 Sekunden lang den Splash-Screen
- falls Abhaengigkeiten fehlen, installiert der Launcher sie waehrend der Splash-Screen sichtbar bleibt
- waehrend der Umwandlung zeigt die Statusleiste den aktuellen Schritt
- nach erfolgreichem Export kann die erzeugte PowerPoint direkt ueber `Oeffnen` gestartet werden
- Tastaturkuerzel: `Strg+O` fuer Bildauswahl, `Strg+S` fuer Zielpfad, `Enter` zum Starten

Das Programm nutzt `assets/app_icon.ico` als Windows-/Taskbar-Icon.
Der Splash-Screen liegt unter `assets/splash.png`.
