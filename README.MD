# 3PMo Helper

Web‑App für das 3PMo‑Modul (Prozessmodellierung). Sie bündelt aktuell zwei Werkzeuge:

- StudentQuiz‑Extraktor (HTML → Tabelle → Excel)
- Studierenden‑Themenzuteilung (Excel‑Upload → balancierte Zuteilung → Excel)

Die App ist vollständig clientseitig, benötigt keinen Server und kann einfach durch Öffnen der `index.html` genutzt werden.

## 🎥 Video‑Tutorials
- ▶️ How to use, install, and build the app with Windsurf AI: https://youtu.be/ICW7BqLz2B8
- ▶️ 3PMo Helper – Studierenden‑Themenzuteilung: HIER LINK EINFÜGEN

## Features

### StudentQuiz‑Extraktor
- HTML in ein Textfeld einfügen
- Extraktion pro Tabellenzeile der Felder: `questionname`, `creatorname`, `difficultylevel`, `rate`, `comments`
- Vorschau in einer Tabelle
- Export als Excel (.xlsx) via SheetJS oder CSV‑Fallback

Details zur Extraktion:
- `questionname`: aus `td.questionname` (Label‑Inhalt; Fallback: Text)
- `creatorname`: aus `td.creatorname` (Datums‑Span wird entfernt)
- `difficultylevel`: nur aus `data-difficultylevel`
- `rate`: nur aus `data-rate`
- `comments`: Zahl aus `td.comment` (`n.a.` → 0)

Export: `3PMo_Helper_StudentQuiz_Extrakt.xlsx` (bzw. CSV‑Fallback)

### Studierenden‑Themenzuteilung
- Excel‑Upload mit den Spalten „Kuerzel“, „Klasse“ (weitere Spalten werden ignoriert)
- Eingabefelder: Testnummer (1–12), Blöcke (Semikolon‑getrennt)
- Blöcke‑Format: `Semesterwoche_Folienbereich`, z. B. `1_8-12; 1_13-17; 2_2-6`
- Zuteilungslogik (klassenweise, balanciert):
  - Jede/r Studierende erhält 1 Block zum Erstellen der Frage
  - Zusätzlich 2 verschiedene Blöcke zum Beantworten (beide ≠ Erstell‑Block)
  - Gleichmäßige Verteilung über alle Blöcke innerhalb der Klasse
- Zusammenfassung vor der Vorschau:
  - Gesamtsicht je Block (Anzahl „Erstellen“/„Beantworten“)
  - Pro Klasse dieselbe Übersicht
  - Mit Total‑Zeile
- Export als Excel mit AutoFilter und breiten Spalten C/D
  - Dateiname: `3PMo_Zuteilung_TestNN.xlsx` (NN zweistellig)

## Ordnerstruktur
- `index.html` – UI mit Tabs (Extraktor/Zuteilung) und Script‑Einbindung
- `script.js` – Parsing, Extraktion, Zuteilung, Vorschau, Exporte
- `styles.css` – Styles (Tabs, Tabellen, Summary)
- `vendor/xlsx.full.min.js` – lokale SheetJS‑Bibliothek (für XLSX‑Export)

## Quick Start
- `index.html` im Browser öffnen (Doppelklick)
- Optional lokaler Server:
  - Python:
    ```powershell
    python -m http.server 5500
    ```
    → http://127.0.0.1:5500
  - Node (eine Option):
    ```powershell
    npx http-server -p 5500
    ```

## Nutzung

### StudentQuiz‑Extraktor
1. HTML der StudentQuiz‑Tabelle ins Textfeld einfügen
2. „Extrahieren & Vorschau“ klicken
3. „Als Excel herunterladen“ (oder CSV‑Fallback)

### Studierenden‑Themenzuteilung
1. Excel „Kuerzel/Klasse“ hochladen (erstes Tabellenblatt wird gelesen)
2. Testnummer (1–12) eintragen
3. Blöcke als Semikolon‑Liste eingeben, z. B. `1_8-12; 1_13-17; 2_2-6`
4. „Zuteilen & Vorschau“ klicken: Vorschau + Zusammenfassungen
5. „Als Excel herunterladen“ speichert `3PMo_Zuteilung_TestNN.xlsx`
6. Optional: „Template herunterladen“ erzeugt ein Excel mit den Kopfzeilen „Kuerzel“, „Klasse“

## Excel‑Export (SheetJS)
Die App nutzt SheetJS (`vendor/xlsx.full.min.js`). Falls nicht vorhanden, wird automatisch eine CSV erzeugt.

Mögliche Download‑Quellen für SheetJS:
- https://cdn.jsdelivr.net/npm/xlsx@0.19.3/dist/xlsx.full.min.js
- https://unpkg.com/xlsx@0.19.3/dist/xlsx.full.min.js
- https://raw.githubusercontent.com/SheetJS/sheetjs/v0.19.3/dist/xlsx.full.min.js

## Grenzen
- Extraktor: funktioniert, solange die HTML‑Struktur (Klassen/Attribute) der StudentQuiz‑Tabelle beibehalten wird
- `n.a.` wird als 0 gezählt; andere Lokalisierungen können Anpassungen erfordern

## Datenschutz
Alle Daten bleiben lokal im Browser. Es findet keine Server‑Verarbeitung statt.
