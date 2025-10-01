# Automatische Vorauswahl für manuelle Bewertung

## Übersicht
Die App wählt automatisch Studierende aus, die manuell überprüft werden sollten, basierend auf definierten Kriterien.

## Auswahlkriterien

### 1. Mehr als 1 Frage erstellt
- **Feld**: `question_count`
- **Bedingung**: `> 1`
- **Begründung**: Zu viele Fragen könnten auf niedrige Qualität hinweisen

### 2. Durchschnittsbewertung unter 4
- **Feld**: `avg_rate` (Ø Bewertung)
- **Bedingung**: `> 0 und < 4`
- **Begründung**: Schlechte Community-Bewertung der Fragen

### 3. Geringe Antwortpunkte
- **Feld**: `total_answers_points` (Σ Punkte Antworten)
- **Bedingung**: `> 0 und < 5`
- **Begründung**: Wenig Engagement bei Beantwortung

### 4. Wenige Kommentare
- **Feld**: `total_comments` (Σ Kommentare)
- **Bedingung**: `< 3`
- **Begründung**: Geringe Interaktion mit anderen Studierenden

### 5. Mehr falsche als richtige Antworten
- **Felder**: `false_answers_points` > `correct_answers_points`
- **Bedingung**: Negative Antwortbilanz
- **Begründung**: Schlechtes Verständnis des Stoffes

## Visuelle Markierungen

### Rot markiert (`.review-required`)
- **Studentenname**: Rot, wenn **irgendein** Kriterium erfüllt ist
- **Betroffene Werte**: Nur die Werte, die das Kriterium erfüllen

### Beispiel
**Student**: Max Muster (rot, da Kriterium erfüllt)
- **Anzahl Fragen**: 3 (rot, da > 1)
- **Ø Bewertung**: 3.2 (rot, da < 4)
- **Σ Punkte Antworten**: 4.5 (rot, da > 0 und < 5)
- **Σ Kommentare**: 2 (rot, da < 3)
- **Punkte richtige Antworten**: 5 (rot, da falsche > richtige)
- **Punkte falsche Antworten**: 8 (rot, da falsche > richtige)

## Interaktion

### Checkbox automatisch aktiviert
- Studierende, die **mindestens ein** Kriterium erfüllen, werden automatisch mit Checkbox markiert
- **Kann manuell angepasst werden**: Benutzer kann Checkbox deaktivieren oder weitere aktivieren

### Filter-Buttons
- **"Nur manuell zu Bewertende anzeigen"**: Zeigt nur ausgewählte Studierende
- **"Alle anzeigen"**: Zeigt alle Studierenden

## Technische Details

### Funktion: `autoSelectManualReview()`
```javascript
// Wird nach Bewertungsgenerierung aufgerufen
// Iteriert über combinedGradesData
// Setzt requiresManualReview = true
// Speichert reviewFlags = ['field1', 'field2', ...]
```

### Flag-Speicherung
```javascript
student.reviewFlags = [
  'question_count',      // Wenn > 1
  'avg_rate',            // Wenn < 4
  'total_answers_points', // Wenn 0-5
  'total_comments',      // Wenn < 3
  'false_answers_points', // Wenn > correct
  'correct_answers_points' // Wenn < false
];
```

### CSS-Klasse
```css
.review-required {
  color: #ef4444; /* Rot */
  font-weight: bold;
}
```

## Workflow

1. **Bewertungen generieren** → Automatische Kriterienprüfung
2. **Tabelle wird gerendert** → Markierungen in rot sichtbar
3. **Checkboxen sind aktiviert** → Für betroffene Studierende
4. **Benutzer passt an** → Kann Checkboxen ändern
5. **Manuelle Bewertungen eingeben** → In rot markierten Zeilen
6. **Excel exportieren** → Mit manuellen Bewertungen

## Vorteile

✅ **Zeitersparnis**: Keine manuelle Suche nach problematischen Studierenden  
✅ **Konsistenz**: Gleiche Kriterien für alle Tests  
✅ **Transparenz**: Benutzer sieht sofort, warum jemand markiert wurde  
✅ **Flexibilität**: Benutzer kann Auswahl anpassen  
✅ **Visuelle Klarheit**: Rote Markierungen fallen sofort auf  

## Anpassungsmöglichkeiten

Die Kriterien können in `autoSelectManualReview()` leicht angepasst werden:
- Schwellenwerte ändern (z.B. `< 4` → `< 3.5`)
- Neue Kriterien hinzufügen
- Gewichtung implementieren
