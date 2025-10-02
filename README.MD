# 3PMo Helper

Web‑App für das 3PMo‑Modul (Prozessmodellierung). Sie bündelt aktuell zwei Werkzeuge:

- StudentQuiz‑Bewertung (inkl. Extraktion aus StudentQuiz- und Rangliste‑Reiter)
- Studierenden‑Themenzuteilung (Excel‑Upload → balancierte Zuteilung → Excel)

Die App ist vollständig clientseitig, benötigt keinen Server und kann einfach durch Öffnen der `index.html` genutzt werden.

## 🎥 Video‑Tutorials
- ▶️ How to use, install, and build the app with Windsurf AI: https://youtu.be/ICW7BqLz2B8
- ▶️ 3PMo Helper – Studierenden‑Themenzuteilung: https://youtu.be/E9MqgW1tyAA

## Features

<!-- Ehemalige 'StudentQuiz‑Extraktor' Sektion entfernt; Extraktion ist Teil von 'StudentQuiz‑Bewertung'. -->

### StudentQuiz‑Bewertung
- Kombination der extrahierten Daten aus Bereich 2 (StudentQuiz) und Bereich 3 (Rangliste)
- Automatische Bewertung je Studierende/n gemäß Gewichtung 50/25/25 (siehe Abschnitt „Hinweise zur Bewertung“ in der App)
- Zusätzliche Abzüge: falscher Frageblock (−20), mehr falsche als richtige Antworten (−10)
- Markierung „Manuell zu bewerten“, wenn definierte Kriterien zutreffen (z. B. Rating < 4, wenig Kommentare, etc.)
- Direkte Bearbeitung in der Tabelle: `Bew. man.` und `Begründung` sind editierbar; Zeilenumbrüche in Begründungen werden unterstützt
- „Nur manuell zu Bewertende anzeigen“ Filter und Sortierfunktionen pro Spalte
- Excel‑Export mit Formatierung (Gitternetzlinien via Zellrahmen, Textumbruch in Spalte Begründung, vertikale Ausrichtung oben)
- Export filtert Einträge mit `Kürzel = '-'` automatisch heraus und sortiert nach Kürzel
- Extraktion in Bereich 2: Felder `questionname`, `creatorname`, `difficultylevel` (aus `data-difficultylevel`), `rate` (aus `data-rate`), `comments` (`n.a.` → 0)
- Extraktion in Bereich 3: Felder `student_name`, `published_question_points`, `rating_points`, `correct_answers_points`, `false_answers_points`
- Optional: Export der reinen Extraktionsdaten (Bereich 2/3) als Excel, CSV‑Fallback

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
- `index.html` – UI mit Tabs (Bewertung/Zuteilung) und Script‑Einbindung
- `script.js` – Parsing, Extraktion, Zuteilung, Vorschau, Exporte
- `styles.css` – Styles (Tabs, Tabellen, Summary)
- `vendor/xlsx.bundle.js` – eingebundene SheetJS‑Bibliothek (für XLSX‑Export)
- `vendor/xlsx.full.min.js` – alternative/zusätzliche Distribution von SheetJS

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

<!-- Ehemalige 'StudentQuiz‑Extraktor' Nutzungssektion entfernt; die Extraktion erfolgt in den Bereichen 2 und 3 des Bewertungstabs. -->

### StudentQuiz‑Bewertung
1. In Bereich 1 einen Test anlegen oder laden (Name angeben)
2. Optional: „Studierenden‑Hilfstabelle“ hochladen (global gespeichert)
3. Optional: „Zuteilungstabelle für aktuellen Test“ laden (testspezifisch gespeichert)
4. Bereich 2: StudentQuiz‑HTML einfügen und extrahieren
5. Bereich 3: Ranglisten‑HTML einfügen und extrahieren
6. Bereich 4: „Bewertungen generieren“ klicken, anschließend bei Bedarf `Bew. man.` und `Begründung` bearbeiten
7. Optional: „Nur manuell zu Bewertende anzeigen“ aktivieren
8. „Bewertungen als Excel herunterladen“

### Studierenden‑Themenzuteilung
1. Excel „Kuerzel/Klasse“ hochladen (erstes Tabellenblatt wird gelesen)
2. Testnummer (1–12) eintragen
3. Blöcke als Semikolon‑Liste eingeben, z. B. `1_8-12; 1_13-17; 2_2-6`
4. „Zuteilen & Vorschau“ klicken: Vorschau + Zusammenfassungen
5. „Als Excel herunterladen“ speichert `3PMo_Zuteilung_TestNN.xlsx`
6. Optional: „Template herunterladen“ erzeugt ein Excel mit den Kopfzeilen „Kuerzel“, „Klasse“

## Excel‑Export (SheetJS)
Die App nutzt SheetJS (in dieser App per `vendor/xlsx.bundle.js` eingebunden). Falls XLSX nicht verfügbar ist, wird automatisch eine CSV erzeugt.

Formatierung im Bewertungs‑Export:
- Gitternetzlinien (dünne Zellrahmen) für Kopf‑ und Datenzeilen
- Spalte „Begründung“: kleinerer Font, Textumbruch aktiv, vertikale Ausrichtung oben
- Alle Zellen: vertikale Ausrichtung oben
- Zeilen mit `Kürzel = '-'` werden nicht exportiert

Mögliche Download‑Quellen für SheetJS:
- https://cdn.jsdelivr.net/npm/xlsx@0.19.3/dist/xlsx.full.min.js
- https://unpkg.com/xlsx@0.19.3/dist/xlsx.full.min.js
- https://raw.githubusercontent.com/SheetJS/sheetjs/v0.19.3/dist/xlsx.full.min.js

## Grenzen
- Extraktion: funktioniert, solange die HTML‑Struktur (Klassen/Attribute) der StudentQuiz‑Tabelle beibehalten wird
- `n.a.` wird als 0 gezählt; andere Lokalisierungen können Anpassungen erfordern

## Datenschutz
Alle Daten bleiben lokal im Browser (LocalStorage). Es findet keine Server‑Verarbeitung statt.

Persistenz (Überblick):
- Pro Test gespeichert: eingegebene HTML‑Extrakte (StudentQuiz/Rangliste), manuelle Bewertungen/Begründungen, Markierung „manuell zu bewerten“, Zuteilungstabelle
- Global gespeichert: „Studierenden‑Hilfstabelle“, Tabellenhöhe‑Einstellung für Bereich 4
- „Test löschen“ entfernt alle Daten des aktuellen Tests (die globale Hilfstabelle bleibt erhalten)
– Export/Import (JSON) enthält: manuelle Noten, Begründungen, „manuell zu bewerten“-Markierung
