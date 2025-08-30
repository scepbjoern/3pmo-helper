# StudentQuiz Table Extractor

Extrahiert aus einer in HTML eingefügten StudentQuiz-Tabelle die Felder questionname, creatorname, difficultylevel, rate und comments und ermöglicht den Export als Excel (XLSX) oder CSV.

## 🎥 Video Tutorial
[▶️ How to use, install, and build the app with Windsurf AI](https://youtu.be/ICW7BqLz2B8)

## Features
- Einfügen von HTML im Textfeld
- Extraktion pro Tabellenzeile
- Vorschau in einer Tabelle
- Export als Excel (.xlsx) via SheetJS oder CSV-Fallback
- Dunkles, aufgeräumtes UI

## Ordnerstruktur
- `index.html` – UI und Script-Einbindung
- `script.js` – Parsing, Extraktion, Vorschau, Export
- `styles.css` – Styles
- `vendor/xlsx.full.min.js` – lokale SheetJS-Bibliothek (für XLSX-Export)

## Quick Start
- Öffne `index.html` im Browser (Doppelklick), oder
- Lokaler Server (optional):
  - Python:
    ```powershell
    python -m http.server 5500
    ```
    Dann: http://127.0.0.1:5500
  - Node (eine Option):
    ```powershell
    npx http-server -p 5500
    ```
- Klicke „Beispiel einfügen“ zum Testen.

## Nutzung
1. HTML, das eine StudentQuiz-Tabelle enthält, in das Textfeld einfügen.
2. „Extrahieren & Vorschau“ klicken.
3. Optional: „Als Excel herunterladen“ (oder CSV-Fallback).

## Extraktion im Detail
- `questionname`: aus `td.questionname` (Label-Inhalt; Fallback: reiner Text)
- `creatorname`: aus `td.creatorname` (Datums-Span wird entfernt)
- `difficultylevel`: nur aus `data-difficultylevel` (kein Titel-Fallback)
- `rate`: nur aus `data-rate` (kein Titel-Fallback)
- `comments`: numerischer Inhalt in `td.comment` (`n.a.` → 0)

Hinweis: Wenn `data-difficultylevel`/`data-rate` fehlen, bleiben die Felder leer.

## Excel-Export (SheetJS)
Für den XLSX-Export wird SheetJS benötigt:

- Datei: `vendor/xlsx.full.min.js`
- Mögliche Download-Quellen:
  - https://cdn.jsdelivr.net/npm/xlsx@0.19.3/dist/xlsx.full.min.js
  - https://unpkg.com/xlsx@0.19.3/dist/xlsx.full.min.js
  - https://raw.githubusercontent.com/SheetJS/sheetjs/v0.19.3/dist/xlsx.full.min.js

Falls SheetJS nicht geladen ist, erzeugt die App automatisch eine CSV-Datei.

## Grenzen
- Funktioniert zuverlässig, solange die HTML-Struktur (Klassen/Attribute) der StudentQuiz-Tabelle beibehalten wird.
- Lokalisierung: „n.a.“ wird als 0 gezählt; andere Lokalisierungen können Anpassungen erfordern.

## Datenschutz
Alle Daten bleiben lokal im Browser. Es findet keine Serververarbeitung statt.
