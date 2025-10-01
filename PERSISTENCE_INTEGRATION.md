# Persistence Integration Guide

## Erforderliche Änderungen in script.js

### 1. Element-Referenzen hinzufügen (ca. Zeile 33)

Nach `gradesTableBody: $('#gradesTable tbody'),` hinzufügen:

```javascript
    // Test management UI
    testNameInput: $('#testNameInput'),
    btnNewTest: $('#btnNewTest'),
    btnLoadTest: $('#btnLoadTest'),
    btnSaveTest: $('#btnSaveTest'),
    btnImportTest: $('#btnImportTest'),
    testFileInput: $('#testFileInput'),
    testStatus: $('#testStatus')
```

### 2. Globale Variable für aktuellen Test hinzufügen (ca. Zeile 48)

Nach `let combinedGradesData = [];` hinzufügen:

```javascript
let currentTestName = null;
```

### 3. Test-Status-Funktion hinzufügen (ca. Zeile 68)

Nach `function setGradesStatus(msg)` hinzufügen:

```javascript
  function setTestStatus(msg) {
    if (els.testStatus) els.testStatus.textContent = msg || '';
  }
```

### 4. renderCombinedGradesTable() anpassen (ca. Zeile 613-616)

Ersetze die Zeilen für `manual_grade` und `justification`:

```javascript
          } else if (cell.key === 'manual_grade') {
            td.innerHTML = `<div class="editable-cell" contenteditable="true" data-student="${escapeHtml(r.student_name)}" data-field="manual_grade" data-placeholder="Klicken zum Bearbeiten">${cell.val || ''}</div>`;
          } else if (cell.key === 'justification') {
            td.innerHTML = `<div class="editable-cell" contenteditable="true" data-student="${escapeHtml(r.student_name)}" data-field="justification" data-placeholder="Klicken zum Bearbeiten">${cell.val || ''}</div>`;
```

### 5. Event-Handler für editierbare Zellen hinzufügen (nach renderCombinedGradesTable)

```javascript
  function attachEditableListeners() {
    const editableCells = document.querySelectorAll('.editable-cell');
    editableCells.forEach(cell => {
      cell.addEventListener('blur', handleCellEdit);
      cell.addEventListener('keydown', (e) => {
        if (e.key === 'Enter' && !e.shiftKey) {
          e.preventDefault();
          cell.blur();
        }
      });
    });
  }

  function handleCellEdit(e) {
    const cell = e.target;
    const studentName = cell.dataset.student;
    const field = cell.dataset.field;
    const value = cell.textContent.trim();
    
    if (!currentTestName) {
      setGradesStatus('Bitte zuerst einen Test anlegen oder laden.');
      return;
    }
    
    // Update in combinedGradesData
    const student = combinedGradesData.find(s => s.student_name === studentName);
    if (student) {
      student[field] = value;
      saveCurrentTestData();
      setGradesStatus(`Änderung für ${studentName} gespeichert.`);
    }
  }
```

### 6. Test-Verwaltungs-Funktionen hinzufügen (vor Event-Bindings)

```javascript
  function createNewTest() {
    const testName = els.testNameInput && els.testNameInput.value.trim();
    if (!testName) {
      setTestStatus('Bitte einen Testnamen eingeben.');
      return;
    }
    currentTestName = testName;
    setCurrentTestName(testName);
    setTestStatus(`Test "${testName}" erstellt.`);
    els.testNameInput.value = testName;
  }

  function loadExistingTest() {
    const testName = els.testNameInput && els.testNameInput.value.trim();
    if (!testName) {
      setTestStatus('Bitte einen Testnamen eingeben.');
      return;
    }
    
    const data = loadTestData(testName);
    if (!data) {
      setTestStatus(`Test "${testName}" nicht gefunden.`);
      return;
    }
    
    currentTestName = testName;
    applyManualGradesFromStorage(data);
    setTestStatus(`Test "${testName}" geladen.`);
  }

  function applyManualGradesFromStorage(data) {
    if (!data.manualGrades) return;
    
    // Apply to combinedGradesData
    combinedGradesData.forEach(row => {
      const stored = data.manualGrades[row.student_name];
      if (stored) {
        row.manual_grade = stored.manual_grade || null;
        row.justification = stored.justification || '';
      }
    });
    
    // Re-render table
    if (combinedGradesData.length > 0) {
      renderCombinedGradesTable(combinedGradesData);
      attachEditableListeners();
    }
  }

  function saveCurrentTestData() {
    if (!currentTestName) return;
    
    // Extract only manual grades
    const manualGrades = {};
    combinedGradesData.forEach(row => {
      if (row.manual_grade || row.justification) {
        manualGrades[row.student_name] = {
          manual_grade: row.manual_grade || null,
          justification: row.justification || ''
        };
      }
    });
    
    saveTestData(currentTestName, { manualGrades });
  }

  function exportTest() {
    if (!currentTestName) {
      setTestStatus('Bitte zuerst einen Test anlegen oder laden.');
      return;
    }
    
    const manualGrades = {};
    combinedGradesData.forEach(row => {
      if (row.manual_grade || row.justification) {
        manualGrades[row.student_name] = {
          manual_grade: row.manual_grade || null,
          justification: row.justification || ''
        };
      }
    });
    
    exportTestDataAsJSON(currentTestName, { manualGrades });
    setTestStatus(`Test "${currentTestName}" exportiert.`);
  }

  async function importTest() {
    if (!els.testFileInput) return;
    els.testFileInput.click();
  }

  async function handleTestFileImport(e) {
    const file = e.target.files[0];
    if (!file) return;
    
    try {
      const imported = await importTestDataFromJSON(file);
      currentTestName = imported.testName;
      setCurrentTestName(imported.testName);
      els.testNameInput.value = imported.testName;
      
      // Save to LocalStorage
      saveTestData(imported.testName, imported.data);
      
      // Apply to current data
      applyManualGradesFromStorage(imported.data);
      
      setTestStatus(`Test "${imported.testName}" importiert.`);
    } catch (err) {
      setTestStatus('Fehler beim Importieren: ' + err.message);
    }
    
    // Reset file input
    e.target.value = '';
  }
```

### 7. Event-Bindings hinzufügen (ca. Zeile 1015)

Nach den Grades-Events hinzufügen:

```javascript
  // Test management events
  if (els.btnNewTest) els.btnNewTest.addEventListener('click', createNewTest);
  if (els.btnLoadTest) els.btnLoadTest.addEventListener('click', loadExistingTest);
  if (els.btnSaveTest) els.btnSaveTest.addEventListener('click', exportTest);
  if (els.btnImportTest) els.btnImportTest.addEventListener('click', importTest);
  if (els.testFileInput) els.testFileInput.addEventListener('change', handleTestFileImport);
```

### 8. renderCombinedGradesTable() erweitern (am Ende der Funktion)

Nach `els.gradesTableBody.appendChild(fragment);` hinzufügen:

```javascript
    // Attach listeners for editable cells
    attachEditableListeners();
    
    // Load manual grades if test is active
    if (currentTestName) {
      const data = loadTestData(currentTestName);
      if (data) {
        applyManualGradesFromStorage(data);
      }
    }
```

### 9. Beim App-Start aktuellen Test laden (am Ende von script.js)

```javascript
  // Load current test on startup
  const lastTest = getCurrentTestName();
  if (lastTest && els.testNameInput) {
    els.testNameInput.value = lastTest;
    currentTestName = lastTest;
    setTestStatus(`Aktiver Test: "${lastTest}"`);
  }
```

## Zusammenfassung

Die Implementierung ermöglicht:
- ✅ Automatisches Speichern von manuellen Bewertungen im Browser (LocalStorage)
- ✅ Editierbare Tabellenzellen für manuelle Bewertung und Begründung
- ✅ Test-Verwaltung (Neu, Laden, Exportieren, Importieren)
- ✅ JSON-Export für Backups und Sharing
- ✅ Mehrere Tests parallel verwaltbar
