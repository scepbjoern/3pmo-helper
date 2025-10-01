(function(){
  const $ = (sel, root=document) => root.querySelector(sel);
  const $$ = (sel, root=document) => Array.from(root.querySelectorAll(sel));

  const els = {
    input: $('#htmlInput'),
    btnParse: $('#btnParse'),
    btnClear: $('#btnClear'),
    btnDownload: $('#btnDownload'),
    status: $('#status'),
    tableBody: $('#previewTable tbody'),

    // Tabs & Pages
    tabSQ: $('#tabSQ'),
    tabAssign: $('#tabAssign'),
    pageSQ: $('#page-sq'),
    pageAssign: $('#page-assign'),

    // Ranking UI (Bereich 3)
    inputRanking: $('#htmlInputRanking'),
    btnParseRanking: $('#btnParseRanking'),
    btnClearRanking: $('#btnClearRanking'),
    btnDownloadRanking: $('#btnDownloadRanking'),
    statusRanking: $('#statusRanking'),
    rankingTableBody: $('#previewTableRanking tbody'),
    // Grades UI (Bereich 4)
    btnGenerateGrades: $('#btnGenerateGrades'),
    btnDownloadGrades: $('#btnDownloadGrades'),
    statusGrades: $('#statusGrades'),
    gradesTableBody: $('#gradesTable tbody'),
    gradeFilterButtons: $('#gradeFilterButtons'),
    btnFilterManual: $('#btnFilterManual'),
    btnShowAll: $('#btnShowAll'),
    tableHeightControl: $('#tableHeightControl'),
    tableHeightSlider: $('#tableHeightSlider'),
    tableHeightValue: $('#tableHeightValue'),
    gradesTableWrapper: $('#gradesTableWrapper'),
    
    // Helper tables UI
    studentHelperFile: $('#studentHelperFile'),
    btnUploadStudentHelper: $('#btnUploadStudentHelper'),
    btnClearStudentHelper: $('#btnClearStudentHelper'),
    statusStudentHelper: $('#statusStudentHelper'),
    assignmentFile: $('#assignmentFile'),
    btnUploadAssignment: $('#btnUploadAssignment'),
    btnClearAssignment: $('#btnClearAssignment'),
    statusAssignment: $('#statusAssignment'),

    // Test management UI
    testNameInput: $('#testNameInput'),
    btnNewTest: $('#btnNewTest'),
    btnLoadTest: $('#btnLoadTest'),
    btnSaveTest: $('#btnSaveTest'),
    btnImportTest: $('#btnImportTest'),
    testFileInput: $('#testFileInput'),
    testStatus: $('#testStatus'),

    // Assignment UI
    studentsFile: $('#studentsFile'),
    testNumber: $('#testNumber'),
    blocksInput: $('#blocksInput'),
    btnTemplate: $('#btnTemplate'),
    btnAssign: $('#btnAssign'),
    btnDownloadAssign: $('#btnDownloadAssign'),
    assignStatus: $('#assignStatus'),
    assignTableBody: $('#assignPreviewTable tbody')
  };

  let extractedRows = [];
  let rankingRows = [];
  let combinedGradesData = []; // Combined data from Bereich 2 & 3
  let currentTestName = null;
  let saveDebounceTimer = null; // Debounce timer for auto-save
  let currentSortColumn = 'student_name'; // Default sort column
  let currentSortOrder = 'asc'; // Default sort order
  let isFilterActive = false; // Track if manual review filter is active
  let studentHelperData = []; // [{kuerzel, fullname}] - Global helper table
  let assignmentData = []; // Current test assignment data
  // Data for assignment page
  let studentsList = []; // [{Kuerzel, Klasse}]
  let assignRows = [];   // rows for preview/export
  let currentTestNumber = null;

  function setStatus(msg) {
    els.status.textContent = msg || '';
  }

  function setAssignStatus(msg) {
    if (els.assignStatus) els.assignStatus.textContent = msg || '';
  }

  function setRankingStatus(msg) {
    if (els.statusRanking) els.statusRanking.textContent = msg || '';
  }

  function setGradesStatus(msg) {
    if (els.statusGrades) els.statusGrades.textContent = msg || '';
  }

  function setTestStatus(msg) {
    if (els.testStatus) els.testStatus.textContent = msg || '';
  }

  function clearAll() {
    els.input.value = '';
    extractedRows = [];
    renderTable([]);
    setStatus('Eingabe und Ergebnis geleert.');
  }

  function normalizeNumber(str) {
    if (str == null) return '';
    const s = String(str).trim().replace(/\u00a0/g, ' ').replace(',', '.');
    // keep digits and single dot
    const m = s.match(/-?[0-9]+(?:\.[0-9]+)?/);
    return m ? m[0] : '';
  }

  function numberOrNull(str) {
    const n = normalizeNumber(str);
    return n === '' ? null : parseFloat(n);
  }

  function extractWithRegex(innerHTML, patterns) {
    for (const re of patterns) {
      const m = innerHTML.match(re);
      if (m && m[1] != null) return m[1].trim();
    }
    return '';
  }

  function stripTags(html) {
    return html.replace(/<[^>]*>/g, '').replace(/\s+/g, ' ').trim();
  }

  function removeParentheses(name) {
    // Remove everything from first opening parenthesis onwards
    const idx = name.indexOf('(');
    return idx >= 0 ? name.substring(0, idx).trim() : name.trim();
  }

  function updateGradesButtonState() {
    // Enable "Bewertungen generieren" if we have data from Bereich 2 or 3
    if (els.btnGenerateGrades) {
      els.btnGenerateGrades.disabled = !(extractedRows.length > 0 || rankingRows.length > 0);
    }
  }

  function parseHTML() {
    const html = els.input.value;
    if (!html || !html.trim()) {
      setStatus('Bitte HTML einfügen.');
      return;
    }

    let doc;
    try {
      const parser = new DOMParser();
      doc = parser.parseFromString(html, 'text/html');
    } catch (e) {
      setStatus('Fehler beim Parsen des HTML.');
      return;
    }

    const rows = $$('table tbody tr', doc);
    const trList = rows.length ? rows : $$('tr', doc); // Fallback

    const results = [];
    trList.forEach(tr => {
      // Extract target tds by class
      const tdQ = $('td.questionname', tr);
      const tdCreator = $('td.creatorname', tr);
      const tdDiff = $('td.difficultylevel', tr);
      const tdRate = $('td.rates', tr);
      const tdComment = $('td.comment', tr);
      if (!tdQ && !tdCreator && !tdDiff && !tdRate && !tdComment) return;

      const qHTML = tdQ ? tdQ.innerHTML : '';
      const creatorHTML = tdCreator ? tdCreator.innerHTML : '';
      const diffHTML = tdDiff ? tdDiff.innerHTML : '';
      const rateHTML = tdRate ? tdRate.innerHTML : '';
      const commentHTML = tdComment ? tdComment.innerHTML : '';

      // questionname via regex on label, fallback to text
      let questionname = extractWithRegex(qHTML, [
        /<label[^>]*>([\s\S]*?)<\/label>/i
      ]);
      if (!questionname) questionname = stripTags(qHTML);

      // creatorname: remove date span, strip tags, cleanup
      let creatorname = creatorHTML
        .replace(/<span[^>]*class="[^"]*date[^"]*"[^>]*>[\s\S]*?<\/span>/gi, ' ');
      creatorname = stripTags(creatorname);
      creatorname = removeParentheses(creatorname);
      // Sometimes there is only an empty span; ensure empty becomes ''
      if (/^n\.a\.$/i.test(creatorname)) creatorname = '';

      // difficultylevel: only from data-difficultylevel (no title fallback)
      let difficultylevel = extractWithRegex(diffHTML, [
        /data-difficultylevel="([^"]*)"/i
      ]);
      difficultylevel = normalizeNumber(difficultylevel || '');

      // rate: only from data-rate (no title fallback)
      let rate = extractWithRegex(rateHTML, [
        /data-rate="([^"]*)"/i
      ]);
      rate = normalizeNumber(rate || '');

      // comments: number inside the comment cell, n.a. -> 0
      let comments = 0;
      if (/n\.a\./i.test(commentHTML)) {
        comments = 0;
      } else {
        const m = commentHTML.match(/>\s*(\d+)\s*</);
        comments = m ? parseInt(m[1], 10) : 0;
      }

      const row = {
        questionname: questionname || '',
        creatorname: creatorname || '',
        difficultylevel: difficultylevel || '',
        rate: rate || '',
        comments
      };

      // Only add if at least questionname present
      if (row.questionname || row.creatorname || row.difficultylevel || row.rate || row.comments) {
        results.push(row);
      }
    });

    extractedRows = results;
    renderTable(results);
    setStatus(results.length ? `${results.length} Zeilen extrahiert.` : 'Keine passenden Daten gefunden.');
    updateGradesButtonState();
    
    // Clear input to improve performance
    els.input.value = '';
  }

  function renderTable(rows) {
    els.tableBody.innerHTML = '';
    const fragment = document.createDocumentFragment();
    rows.forEach(r => {
      const tr = document.createElement('tr');
      tr.innerHTML = `
        <td>${escapeHtml(r.questionname)}</td>
        <td>${escapeHtml(r.creatorname)}</td>
        <td>${escapeHtml(r.difficultylevel)}</td>
        <td>${escapeHtml(r.rate)}</td>
        <td>${escapeHtml(r.comments)}</td>
      `;
      fragment.appendChild(tr);
    });
    els.tableBody.appendChild(fragment);
  }

  function escapeHtml(text) {
    return String(text ?? '')
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#039;');
  }

  function toWorksheetData(rows) {
    return rows.map(r => ({
      questionname: r.questionname,
      creatorname: r.creatorname,
      difficultylevel: r.difficultylevel,
      rate: r.rate,
      comments: r.comments
    }));
  }

  function downloadExcel() {
    if (!extractedRows.length) {
      setStatus('Keine Daten zum Exportieren. Bitte zuerst extrahieren.');
      return;
    }

    const data = toWorksheetData(extractedRows);

    if (window.XLSX && XLSX.utils && XLSX.writeFile) {
      const ws = XLSX.utils.json_to_sheet(data);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, 'Daten');
      XLSX.writeFile(wb, '3PMo_Helper_StudentQuiz_Extrakt.xlsx');
      setStatus(`Excel exportiert (${data.length} Zeilen).`);
    } else {
      // CSV Fallback
      const csv = toCSV(data);
      const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = '3PMo_Helper_StudentQuiz_Extrakt.csv';
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
      setStatus(`CSV exportiert (${data.length} Zeilen).`);
    }
  }

  function toCSV(arr) {
    if (!arr.length) return '';
    const headers = Object.keys(arr[0]);
    const escape = v => '"' + String(v ?? '').replace(/"/g, '""') + '"';
    const lines = [headers.map(escape).join(',')];
    for (const row of arr) {
      lines.push(headers.map(h => escape(row[h])).join(','));
    }
    return lines.join('\r\n');
  }

  // --- Bereich 3: Rangliste-Extraktion ---
  function clearRanking() {
    if (els.inputRanking) els.inputRanking.value = '';
    rankingRows = [];
    renderRankingTable([]);
    setRankingStatus('Eingabe und Ergebnis geleert.');
  }

  function parseRankingHTML() {
    const html = els.inputRanking && els.inputRanking.value;
    if (!html || !String(html).trim()) { setRankingStatus('Bitte HTML einfügen.'); return; }

    let doc;
    try {
      const parser = new DOMParser();
      doc = parser.parseFromString(html, 'text/html');
    } catch (e) {
      setRankingStatus('Fehler beim Parsen des HTML.');
      return;
    }

    const rows = $$('table tbody tr', doc);
    const trList = rows.length ? rows : $$('tr', doc);

    const results = [];
    trList.forEach(tr => {
      const tdName = $('td.cell.c1', tr);
      const tdPub = $('td.cell.c3', tr);
      const tdStars = $('td.cell.c5', tr);
      const tdCorrect = $('td.cell.c6', tr);
      const tdWrong = $('td.cell.c7', tr);

      if (!tdName && !tdPub && !tdStars && !tdCorrect && !tdWrong) return;

      const nameHTML = tdName ? tdName.innerHTML : '';
      let name = stripTags(nameHTML);
      name = removeParentheses(name);

      const published = numberOrNull(tdPub ? tdPub.textContent : '');
      const rating = numberOrNull(tdStars ? tdStars.textContent : '');
      const correct = numberOrNull(tdCorrect ? tdCorrect.textContent : '');
      const wrong = numberOrNull(tdWrong ? tdWrong.textContent : '');

      const row = {
        student_name: name || '',
        published_question_points: published ?? null,
        rating_points: rating ?? null,
        correct_answers_points: correct ?? null,
        false_answers_points: wrong ?? null
      };

      if (row.student_name || row.published_question_points != null || row.rating_points != null || row.correct_answers_points != null || row.false_answers_points != null) {
        results.push(row);
      }
    });

    rankingRows = results;
    renderRankingTable(results);
    setRankingStatus(results.length ? `${results.length} Zeilen extrahiert.` : 'Keine passenden Daten gefunden.');
    updateGradesButtonState();
    
    // Clear input to improve performance
    if (els.inputRanking) els.inputRanking.value = '';
  }

  function renderRankingTable(rows) {
    if (!els.rankingTableBody) return;
    els.rankingTableBody.innerHTML = '';
    const fragment = document.createDocumentFragment();
    rows.forEach(r => {
      const tr = document.createElement('tr');
      tr.innerHTML = `
        <td>${escapeHtml(r.student_name)}</td>
        <td>${r.published_question_points ?? ''}</td>
        <td>${r.rating_points ?? ''}</td>
        <td>${r.correct_answers_points ?? ''}</td>
        <td>${r.false_answers_points ?? ''}</td>
      `;
      fragment.appendChild(tr);
    });
    els.rankingTableBody.appendChild(fragment);
  }

  function toRankingWorksheetData(rows) {
    return rows.map(r => ({
      student_name: r.student_name,
      published_question_points: r.published_question_points,
      rating_points: r.rating_points,
      correct_answers_points: r.correct_answers_points,
      false_answers_points: r.false_answers_points
    }));
  }

  function downloadRankingExcel() {
    if (!Array.isArray(rankingRows) || !rankingRows.length) { setRankingStatus('Keine Daten zum Exportieren. Bitte zuerst extrahieren.'); return; }
    const data = toRankingWorksheetData(rankingRows);
    if (window.XLSX && XLSX.utils && XLSX.writeFile) {
      const ws = XLSX.utils.json_to_sheet(data);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, 'Rangliste');
      XLSX.writeFile(wb, '3PMo_Helper_Rangliste_Extrakt.xlsx');
      setRankingStatus(`Excel exportiert (${data.length} Zeilen).`);
    } else {
      const csv = toCSV(data);
      const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = '3PMo_Helper_Rangliste_Extrakt.csv';
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
      setRankingStatus(`CSV exportiert (${data.length} Zeilen).`);
    }
  }

  // getSampleRankingHTML removed - now loaded via fetch from samples/ranking-sample.html
  function getSampleRankingHTML_DEPRECATED() {
    return `
<table class="generaltable rankingtable">
  <thead>
    <tr>
      <th class="header c0">Rang</th>
      <th class="header c1">Vollständiger Name</th>
      <th class="header c2">Total Punkte</th>
      <th class="header c3">Punkte für veröffentlichte Fragen</th>
      <th class="header c4">Punkte für bestätigte Fragen</th>
      <th class="header c5">Punkte für erhaltene Sterne</th>
      <th class="header c6">Punkte für richtige Antworten beim letzten Versuch</th>
      <th class="header c7">Punkte für falsche Antworten beim letzten Versuch</th>
      <th class="header c8 lastcol">Persönlicher Fortschritt</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <td class="cell c0">1</td>
      <td class="cell c1"><a href="#">Max Muster</a></td>
      <td class="cell c2">28.5</td>
      <td class="cell c3">1</td>
      <td class="cell c4">0</td>
      <td class="cell c5">4.5</td>
      <td class="cell c6">13</td>
      <td class="cell c7">10</td>
      <td class="cell c8 lastcol">13 %</td>
    </tr>
  </tbody>
</table>`;
  }

  // --- Bereich 4: Bewertungen (kombinierte Tabelle) ---
  function normalizeStudentName(name) {
    // Normalize student names for matching (trim, lowercase, remove extra spaces)
    return String(name || '').trim().toLowerCase().replace(/\s+/g, ' ');
  }

  function generateCombinedGrades() {
    if (!extractedRows.length && !rankingRows.length) {
      setGradesStatus('Bitte zuerst Daten in Bereich 2 und/oder 3 extrahieren.');
      return;
    }

    // Build map: normalized name -> data from Bereich 2 (StudentQuiz)
    const sqMap = new Map();
    for (const row of extractedRows) {
      const name = normalizeStudentName(row.creatorname);
      if (!name) continue;
      if (!sqMap.has(name)) {
        sqMap.set(name, []);
      }
      sqMap.get(name).push(row);
    }

    // Build map: normalized name -> data from Bereich 3 (Rangliste)
    const rankMap = new Map();
    for (const row of rankingRows) {
      const name = normalizeStudentName(row.student_name);
      if (!name) continue;
      rankMap.set(name, row);
    }

    // Merge: union of all student names
    const allNames = new Set([...sqMap.keys(), ...rankMap.keys()]);
    const combined = [];

    for (const normName of allNames) {
      const sqData = sqMap.get(normName) || [];
      const rankData = rankMap.get(normName);

      // Use original name from ranking if available, else from SQ
      let displayName = '';
      if (rankData && rankData.student_name) {
        displayName = rankData.student_name;
      } else if (sqData.length && sqData[0].creatorname) {
        displayName = sqData[0].creatorname;
      } else {
        displayName = normName; // fallback
      }

      // Aggregate SQ data: count questions, avg difficulty, avg rate, sum comments
      let questionCount = sqData.length;
      let avgDifficulty = null;
      let avgRate = null;
      let totalComments = 0;

      if (sqData.length) {
        const difficulties = sqData.map(r => parseFloat(r.difficultylevel)).filter(v => !isNaN(v));
        const rates = sqData.map(r => parseFloat(r.rate)).filter(v => !isNaN(v));
        if (difficulties.length) avgDifficulty = (difficulties.reduce((a,b)=>a+b,0) / difficulties.length).toFixed(2);
        if (rates.length) avgRate = (rates.reduce((a,b)=>a+b,0) / rates.length).toFixed(2);
        totalComments = sqData.reduce((sum, r) => sum + (parseInt(r.comments,10) || 0), 0);
      }

      // Calculate sum of correct + false answers
      let totalAnswersPoints = null;
      const correctPts = rankData ? rankData.correct_answers_points : null;
      const falsePts = rankData ? rankData.false_answers_points : null;
      if (correctPts != null && falsePts != null) {
        totalAnswersPoints = correctPts + falsePts;
      } else if (correctPts != null) {
        totalAnswersPoints = correctPts;
      } else if (falsePts != null) {
        totalAnswersPoints = falsePts;
      }

      const row = {
        student_name: displayName,
        // From Bereich 2 (StudentQuiz)
        question_name: sqRow ? sqRow.questionname : null,
        question_count: questionCount > 0 ? questionCount : null,
        avg_difficultylevel: avgDifficulty,
        avg_rate: avgRate,
        total_comments: totalComments > 0 ? totalComments : null,
        // From Bereich 3 (Rangliste)
        published_question_points: rankData ? rankData.published_question_points : null,
        rating_points: rankData ? rankData.rating_points : null,
        correct_answers_points: rankData ? rankData.correct_answers_points : null,
        false_answers_points: rankData ? rankData.false_answers_points : null,
        total_answers_points: totalAnswersPoints,
        automatic_grade: null,
        manual_grade: null,
        justification: '',
        wrong_block: '',
        expected_responders: ''
      };

      // Calculate automatic grade
      row.automatic_grade = calculateAutomaticGrade(row);

      combined.push(row);
    }

    // Sort by student name
    combined.sort((a, b) => String(a.student_name).localeCompare(String(b.student_name), 'de-CH', { sensitivity: 'base' }));

    combinedGradesData = combined;
    
    // Load saved manual grades BEFORE rendering
    if (currentTestName) {
      const data = loadTestData(currentTestName);
      if (data) {
        applyManualGradesFromStorage(data);
      }
    }
    
    // Auto-select students for manual review based on criteria
    autoSelectManualReview();
    
    // Validate assignments (wrong block, expected responders)
    validateAssignments();
    
    renderCombinedGradesTable(combined);
    setGradesStatus(`${combined.length} Studierende kombiniert.`);

    // Reset filter status
    isFilterActive = false;

    // Enable download button and show filter buttons
    if (els.btnDownloadGrades) els.btnDownloadGrades.disabled = false;
    if (els.gradeFilterButtons) els.gradeFilterButtons.style.display = '';
    if (els.tableHeightControl) els.tableHeightControl.style.display = '';

    // Clear and collapse Bereich 2 & 3 to save browser resources
    if (els.tableBody) els.tableBody.innerHTML = '';
    if (els.rankingTableBody) els.rankingTableBody.innerHTML = '';
    
    // Clear the data arrays to free memory
    extractedRows = [];
    rankingRows = [];
    
    // Collapse the details elements
    const bereich2 = document.querySelector('details:has(#htmlInput)');
    const bereich3 = document.querySelector('details:has(#htmlInputRanking)');
    if (bereich2) bereich2.open = false;
    if (bereich3) bereich3.open = false;
  }

  function autoSelectManualReview() {
    if (!combinedGradesData || !combinedGradesData.length) return;
    
    combinedGradesData.forEach(student => {
      // Skip if already has manual grade from storage
      if (student.manual_grade) return;
      
      const flags = [];
      
      // Helper function to check if value is valid (not null, undefined, or empty string)
      const isValid = (val) => val != null && val !== '';
      
      // Criterion 1: More than 1 question (only if value exists)
      if (isValid(student.question_count)) {
        const questionCount = parseFloat(student.question_count);
        if (!isNaN(questionCount) && questionCount > 1) {
          flags.push('question_count');
        }
      }
      
      // Criterion 2: Average rating < 4 (only if value exists and > 0)
      if (isValid(student.avg_rate)) {
        const avgRate = parseFloat(student.avg_rate);
        if (!isNaN(avgRate) && avgRate > 0 && avgRate < 4) {
          flags.push('avg_rate');
        }
      }
      
      // Criterion 3: Total answers points > 0 but < 5 (only if value exists)
      if (isValid(student.total_answers_points)) {
        const totalAnswersPoints = parseFloat(student.total_answers_points);
        if (!isNaN(totalAnswersPoints) && totalAnswersPoints > 0 && totalAnswersPoints < 5) {
          flags.push('total_answers_points');
        }
      }
      
      // Criterion 4: Total comments < 3 (only if value exists)
      if (isValid(student.total_comments)) {
        const totalComments = parseFloat(student.total_comments);
        if (!isNaN(totalComments) && totalComments >= 0 && totalComments < 3) {
          flags.push('total_comments');
        }
      }
      
      // Criterion 5: False answers > correct answers (only if both values exist)
      if (isValid(student.correct_answers_points) && isValid(student.false_answers_points)) {
        const correctPoints = parseFloat(student.correct_answers_points);
        const falsePoints = parseFloat(student.false_answers_points);
        if (!isNaN(correctPoints) && !isNaN(falsePoints) && falsePoints > correctPoints) {
          flags.push('false_answers_points');
          flags.push('correct_answers_points');
        }
      }
      
      // Mark for manual review if any criterion is met
      if (flags.length > 0) {
        student.requiresManualReview = true;
        student.reviewFlags = flags; // Store which criteria triggered the review
      } else {
        student.reviewFlags = [];
      }
    });
  }

  function renderCombinedGradesTable(rows) {
    if (!els.gradesTableBody) return;
    els.gradesTableBody.innerHTML = '';
    const fragment = document.createDocumentFragment();

    rows.forEach(r => {
      const tr = document.createElement('tr');
      tr.dataset.studentName = r.student_name;
      if (r.requiresManualReview) tr.dataset.manualReview = 'true';

      // Checkbox cell
      const tdCheckbox = document.createElement('td');
      const checkbox = document.createElement('input');
      checkbox.type = 'checkbox';
      checkbox.className = 'manual-review-checkbox';
      checkbox.dataset.student = r.student_name;
      checkbox.checked = r.requiresManualReview || false;
      // Event listener will be attached via delegation, not per checkbox
      tdCheckbox.appendChild(checkbox);
      tr.appendChild(tdCheckbox);

      const cells = [
        { val: r.student_name, key: 'student_name' },
        { val: r.automatic_grade, key: 'automatic_grade' },
        { val: r.manual_grade, key: 'manual_grade' },
        { val: r.justification, key: 'justification' },
        { val: r.question_count, key: 'question_count' },
        { val: r.rating_points, key: 'rating_points' },
        { val: r.published_question_points, key: 'published_question_points' },
        { val: r.total_answers_points, key: 'total_answers_points' },
        { val: r.total_comments, key: 'total_comments' },
        { val: r.avg_rate, key: 'avg_rate' },
        { val: r.avg_difficultylevel, key: 'avg_difficultylevel' },
        { val: r.correct_answers_points, key: 'correct_answers_points' },
        { val: r.false_answers_points, key: 'false_answers_points' },
        { val: r.wrong_block, key: 'wrong_block' },
        { val: r.expected_responders, key: 'expected_responders' }
      ];

      const reviewFlags = r.reviewFlags || [];
      
      cells.forEach(cell => {
        const td = document.createElement('td');
        
        // Check if this field is flagged for review
        const isFlagged = reviewFlags.includes(cell.key);
        
        // Handle editable cells first (manual_grade and justification)
        if (cell.key === 'manual_grade') {
          td.innerHTML = `<div class="editable-cell" contenteditable="true" data-student="${escapeHtml(r.student_name)}" data-field="manual_grade" data-placeholder="Klicken zum Bearbeiten">${cell.val || ''}</div>`;
        } else if (cell.key === 'justification') {
          td.innerHTML = `<div class="editable-cell" contenteditable="true" data-student="${escapeHtml(r.student_name)}" data-field="justification" data-placeholder="Klicken zum Bearbeiten">${cell.val || ''}</div>`;
        } else if (cell.val == null || cell.val === '') {
          td.innerHTML = '<span class="missing">MISSING</span>';
        } else {
          // Student name - mark red if requires manual review
          if (cell.key === 'student_name' && r.requiresManualReview) {
            td.innerHTML = `<span class="review-required">${escapeHtml(cell.val)}</span>`;
          }
          // Color coding for question_count
          else if (cell.key === 'question_count') {
            const count = parseInt(cell.val, 10);
            if (isFlagged) {
              td.innerHTML = `<span class="review-required">${cell.val}</span>`;
            } else if (count > 1) {
              td.innerHTML = `<span class="highlight-violet">${cell.val}</span>`;
            } else if (count === 0) {
              td.innerHTML = `<span class="highlight-red">${cell.val}</span>`;
            } else {
              td.textContent = cell.val;
            }
          } else if (cell.key === 'automatic_grade' && cell.val != null) {
            td.innerHTML = `<strong>${cell.val}%</strong>`;
          }
          // Wrong block - show YES in red
          else if (cell.key === 'wrong_block' && cell.val === 'YES') {
            td.innerHTML = `<span class="review-required">YES</span>`;
          }
          // Expected responders - preserve line breaks
          else if (cell.key === 'expected_responders' && cell.val) {
            td.innerHTML = escapeHtml(String(cell.val)).replace(/\n/g, '<br>');
            td.style.whiteSpace = 'pre-wrap';
          }
          // Mark flagged fields in red
          else if (isFlagged) {
            td.innerHTML = `<span class="review-required">${escapeHtml(String(cell.val))}</span>`;
          }
          else {
            td.textContent = cell.val;
          }
        }
        tr.appendChild(td);
      });

      fragment.appendChild(tr);
    });
    
    els.gradesTableBody.appendChild(fragment);
    
    // Attach listeners for editable cells
    attachEditableListeners();
    
    // Attach delegated event listener for checkboxes (only once)
    if (!els.gradesTableBody.dataset.checkboxListenerAttached) {
      els.gradesTableBody.addEventListener('change', (e) => {
        if (e.target.classList.contains('manual-review-checkbox')) {
          handleManualReviewCheckbox(e);
        }
      });
      els.gradesTableBody.dataset.checkboxListenerAttached = 'true';
    }
  }

  function downloadCombinedGradesExcel() {
    if (!Array.isArray(combinedGradesData) || !combinedGradesData.length) {
      setGradesStatus('Keine Daten zum Exportieren. Bitte zuerst Bewertungen generieren.');
      return;
    }

    const data = combinedGradesData.map(r => ({
      student_name: r.student_name,
      automatic_grade: r.automatic_grade ?? 'MISSING',
      manual_grade: r.manual_grade ?? '',
      justification: r.justification ?? '',
      question_count: r.question_count ?? 'MISSING',
      rating_points: r.rating_points ?? 'MISSING',
      published_question_points: r.published_question_points ?? 'MISSING',
      total_answers_points: r.total_answers_points ?? 'MISSING',
      total_comments: r.total_comments ?? 'MISSING',
      avg_rate: r.avg_rate ?? 'MISSING',
      avg_difficultylevel: r.avg_difficultylevel ?? 'MISSING',
      correct_answers_points: r.correct_answers_points ?? 'MISSING',
      false_answers_points: r.false_answers_points ?? 'MISSING'
    }));

    if (window.XLSX && XLSX.utils && XLSX.writeFile) {
      const ws = XLSX.utils.json_to_sheet(data);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, 'Bewertungen');
      XLSX.writeFile(wb, '3PMo_Helper_Bewertungen.xlsx');
      setGradesStatus(`Excel exportiert (${data.length} Zeilen).`);
    } else {
      const csv = toCSV(data);
      const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = '3PMo_Helper_Bewertungen.csv';
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
      setGradesStatus(`CSV exportiert (${data.length} Zeilen).`);
    }
  }

  // --- Editable cells & persistence ---
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
    let value = cell.textContent.trim();
    
    if (!currentTestName) {
      setGradesStatus('Bitte zuerst einen Test anlegen oder laden.');
      return;
    }
    
    // If manual_grade field: format as percentage if it's a number
    if (field === 'manual_grade' && value) {
      // Remove existing % if present
      value = value.replace('%', '').trim();
      // Check if it's a valid number
      const numValue = parseFloat(value);
      if (!isNaN(numValue)) {
        value = `${numValue}%`;
        cell.textContent = value;
      }
    }
    
    // Update in combinedGradesData
    const student = combinedGradesData.find(s => s.student_name === studentName);
    if (student) {
      student[field] = value;
      
      // Debounced save - wait 500ms after last edit
      clearTimeout(saveDebounceTimer);
      saveDebounceTimer = setTimeout(() => {
        saveCurrentTestData();
        setGradesStatus(`Änderung für ${studentName} gespeichert.`);
      }, 500);
    }
  }

  // --- Manual review checkbox ---
  function handleManualReviewCheckbox(e) {
    const checkbox = e.target;
    const studentName = checkbox.dataset.student;
    const student = combinedGradesData.find(s => s.student_name === studentName);
    if (student) {
      student.requiresManualReview = checkbox.checked;
      const tr = checkbox.closest('tr');
      if (tr) {
        if (checkbox.checked) {
          tr.dataset.manualReview = 'true';
        } else {
          delete tr.dataset.manualReview;
        }
      }
      saveCurrentTestData();
    }
  }

  function filterManualReviewOnly() {
    const rows = els.gradesTableBody && els.gradesTableBody.querySelectorAll('tr');
    if (!rows) return;
    let visibleCount = 0;
    rows.forEach(tr => {
      if (tr.dataset.manualReview === 'true') {
        tr.style.display = '';
        visibleCount++;
      } else {
        tr.style.display = 'none';
      }
    });
    isFilterActive = true;
    setGradesStatus(`${visibleCount} manuell zu bewertende Studierende angezeigt.`);
  }

  function showAllGrades() {
    const rows = els.gradesTableBody && els.gradesTableBody.querySelectorAll('tr');
    if (!rows) return;
    rows.forEach(tr => {
      tr.style.display = '';
    });
    isFilterActive = false;
    setGradesStatus(`Alle ${rows.length} Studierenden angezeigt.`);
  }

  // --- Table sorting ---
  function sortGradesTable(column) {
    if (!combinedGradesData || !combinedGradesData.length) return;
    
    // Toggle sort order if clicking same column
    if (currentSortColumn === column) {
      currentSortOrder = currentSortOrder === 'asc' ? 'desc' : 'asc';
    } else {
      currentSortColumn = column;
      currentSortOrder = 'asc';
    }
    
    // Sort the data
    combinedGradesData.sort((a, b) => {
      let valA = a[column];
      let valB = b[column];
      
      // Handle null/undefined
      if (valA == null) valA = '';
      if (valB == null) valB = '';
      
      // Remove % for manual_grade comparison
      if (column === 'manual_grade') {
        valA = String(valA).replace('%', '').trim();
        valB = String(valB).replace('%', '').trim();
        valA = valA ? parseFloat(valA) : -Infinity;
        valB = valB ? parseFloat(valB) : -Infinity;
      }
      
      // Try numeric comparison first
      const numA = parseFloat(valA);
      const numB = parseFloat(valB);
      if (!isNaN(numA) && !isNaN(numB)) {
        return currentSortOrder === 'asc' ? numA - numB : numB - numA;
      }
      
      // String comparison
      const strA = String(valA).toLowerCase();
      const strB = String(valB).toLowerCase();
      if (currentSortOrder === 'asc') {
        return strA.localeCompare(strB, 'de-CH');
      } else {
        return strB.localeCompare(strA, 'de-CH');
      }
    });
    
    // Update header indicators
    updateSortIndicators();
    
    // Re-render table
    renderCombinedGradesTable(combinedGradesData);
    
    // Re-apply filter if it was active
    if (isFilterActive) {
      filterManualReviewOnly();
    }
  }
  
  function updateSortIndicators() {
    const headers = document.querySelectorAll('#gradesTable th.sortable');
    headers.forEach(th => {
      const col = th.dataset.sort;
      if (col === currentSortColumn) {
        th.dataset.order = currentSortOrder;
        const indicator = currentSortOrder === 'asc' ? ' ▲' : ' ▼';
        th.textContent = th.textContent.replace(/ [▲▼]$/, '') + indicator;
      } else {
        delete th.dataset.order;
        th.textContent = th.textContent.replace(/ [▲▼]$/, '');
      }
    });
  }

  // --- Test management ---
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
        row.requiresManualReview = stored.requiresManualReview || false;
      }
    });
    
    // NOTE: Do NOT re-render table here to avoid infinite recursion
    // The table will be rendered by the caller (generateCombinedGrades or loadExistingTest)
  }

  function saveCurrentTestData() {
    if (!currentTestName) return;
    
    // Extract only manual grades
    const manualGrades = {};
    combinedGradesData.forEach(row => {
      if (row.manual_grade || row.justification || row.requiresManualReview) {
        manualGrades[row.student_name] = {
          manual_grade: row.manual_grade || null,
          justification: row.justification || '',
          requiresManualReview: row.requiresManualReview || false
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

  // --- Tabs ---
  function showPage(which) {
    const isSQ = which === 'sq';
    if (els.tabSQ && els.tabAssign) {
      els.tabSQ.classList.toggle('active', isSQ);
      els.tabAssign.classList.toggle('active', !isSQ);
    }
    if (els.pageSQ && els.pageAssign) {
      els.pageSQ.classList.toggle('active', isSQ);
      els.pageAssign.classList.toggle('active', !isSQ);
    }
  }

  // --- Assignment: Blocks parsing ---
  function parseBlocks(input) {
    const items = String(input).split(';').map(s => s.trim()).filter(Boolean);
    const re = /^(\d+)_([0-9]+-[0-9]+)$/;
    const blocks = items.map((it) => {
      const m = it.match(re);
      if (!m) throw new Error(`Ungültiger Block: "${it}"`);
      const week = parseInt(m[1], 10);
      const range = m[2];
      return { week, range };
    });
    return blocks;
  }

  function formatBlockLabel(block) {
    const sw = String(block.week).padStart(2, '0');
    return `SW${sw}F${block.range}`;
  }

  // --- Assignment: read students Excel ---
  function readStudentsFromExcel(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = (e) => {
        try {
          const data = new Uint8Array(e.target.result);
          const wb = XLSX.read(data, { type: 'array' });
          const sheetName = wb.SheetNames[0];
          const ws = wb.Sheets[sheetName];
          const rows = XLSX.utils.sheet_to_json(ws, { defval: '' });
          const students = rows.map(r => ({
            Kuerzel: (r['Kuerzel'] ?? '').toString().trim(),
            Klasse: (r['Klasse'] ?? '').toString().trim()
          })).filter(r => r.Kuerzel && r.Klasse);
          resolve(students);
        } catch (err) {
          reject(err);
        }
      };
      reader.onerror = () => reject(new Error('Datei konnte nicht gelesen werden.'));
      reader.readAsArrayBuffer(file);
    });
  }

  function shuffle(arr) {
    const a = arr.slice();
    for (let i = a.length - 1; i > 0; i--) {
      const j = Math.floor(Math.random() * (i + 1));
      [a[i], a[j]] = [a[j], a[i]];
    }
    return a;
  }

  function pickMinIndex(counts, allowed, order) {
    let bestIdx = null;
    let bestVal = Infinity;
    for (const i of order) {
      if (!allowed.includes(i)) continue;
      const v = counts[i];
      if (v < bestVal) {
        bestVal = v;
        bestIdx = i;
      }
    }
    if (bestIdx == null) {
      for (const i of allowed) {
        const v = counts[i];
        if (v < bestVal) {
          bestVal = v;
          bestIdx = i;
        }
      }
    }
    return bestIdx;
  }

  function computeBalancedAssignment(students, blocks) {
    const byClass = new Map();
    for (const s of students) {
      const k = s.Klasse;
      if (!byClass.has(k)) byClass.set(k, []);
      byClass.get(k).push(s);
    }

    const results = [];
    const nBlocks = blocks.length;
    const allIdx = [...Array(nBlocks).keys()];

    for (const [klass, studs] of byClass.entries()) {
      const studsShuffled = shuffle(studs);
      const createCount = Array(nBlocks).fill(0);
      const answerCount = Array(nBlocks).fill(0);
      const blockOrder = shuffle(allIdx);
      const createIdxForStudent = new Map();

      for (const s of studsShuffled) {
        const idx = pickMinIndex(createCount, allIdx, blockOrder);
        createCount[idx]++;
        createIdxForStudent.set(s, idx);
      }

      for (const s of studsShuffled) {
        const cIdx = createIdxForStudent.get(s);
        const allowed1 = allIdx.filter(i => i !== cIdx);
        const a1 = pickMinIndex(answerCount, allowed1, blockOrder);
        answerCount[a1]++;

        const allowed2 = allowed1.filter(i => i !== a1);
        const a2 = pickMinIndex(answerCount, allowed2, blockOrder);
        answerCount[a2]++;

        const createLabel = formatBlockLabel(blocks[cIdx]);
        const ans1Label = formatBlockLabel(blocks[a1]);
        const ans2Label = formatBlockLabel(blocks[a2]);
        results.push({
          'Kürzel': s.Kuerzel,
          'Klasse': klass,
          'Frage erstellen in Block': createLabel,
          'Fragen beantworten in Blöcken': `${ans1Label} & ${ans2Label}`
        });
      }
    }

    return results;
  }

  function renderAssignTable(rows) {
    if (!els.assignTableBody) return;
    els.assignTableBody.innerHTML = '';
    for (const r of rows) {
      const tr = document.createElement('tr');
      tr.innerHTML = `
        <td>${escapeHtml(r['Kürzel'])}</td>
        <td>${escapeHtml(r['Klasse'])}</td>
        <td>${escapeHtml(r['Frage erstellen in Block'])}</td>
        <td>${escapeHtml(r['Fragen beantworten in Blöcken'])}</td>
      `;
      els.assignTableBody.appendChild(tr);
    }
  }

  function downloadStudentsTemplate() {
    if (!(window.XLSX && XLSX.utils && XLSX.writeFile)) return;
    const ws = XLSX.utils.aoa_to_sheet([[ 'Kuerzel', 'Klasse' ]]);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Studierende');
    XLSX.writeFile(wb, '3PMo_Helper_Studierenden_Template.xlsx');
  }

  async function handleAssignClick() {
    setAssignStatus('');
    if (els.assignTableBody) els.assignTableBody.innerHTML = '';
    const summaryRoot = document.getElementById('assignSummary');
    if (summaryRoot) summaryRoot.innerHTML = '';

    const file = els.studentsFile && els.studentsFile.files && els.studentsFile.files[0];
    if (!file) { setAssignStatus('Bitte Excel-Datei mit Studierenden auswählen.'); return; }

    const testNumber = parseInt(els.testNumber && els.testNumber.value, 10);
    if (!(testNumber >= 1 && testNumber <= 12)) { setAssignStatus('Bitte eine gültige Testnummer zwischen 1 und 12 eingeben.'); return; }

    const blocksStr = String(els.blocksInput && els.blocksInput.value || '').trim();
    if (!blocksStr) { setAssignStatus('Bitte Blöcke eingeben (z. B. 1_8-12; 1_13-17; 2_2-6).'); return; }

    let blocks;
    try {
      blocks = parseBlocks(blocksStr);
      if (blocks.length < 3) { setAssignStatus('Es werden mindestens 3 Blöcke benötigt.'); return; }
    } catch (e) {
      setAssignStatus(e.message || 'Ungültiges Blöcke-Format.');
      return;
    }

    setAssignStatus('Lade Studierende...');
    try {
      studentsList = await readStudentsFromExcel(file);
    } catch (e) {
      setAssignStatus('Fehler beim Lesen der Excel-Datei.');
      return;
    }
    if (!studentsList.length) { setAssignStatus('Keine Studierenden gefunden.'); return; }

    setAssignStatus('Zuteilung wird berechnet...');
    assignRows = computeBalancedAssignment(studentsList, blocks);
    assignRows.sort((a, b) => String(a['Kürzel']).localeCompare(String(b['Kürzel']), 'de-CH', { sensitivity: 'base' }));
    currentTestNumber = testNumber;

    renderAssignTable(assignRows);
    setAssignStatus(`Zuteilung erstellt (${assignRows.length} Einträge).`);
    renderAssignSummary(assignRows, blocks);
  }

  function downloadAssignmentExcel() {
    if (!Array.isArray(assignRows) || !assignRows.length) { setAssignStatus('Keine Zuteilung zum Exportieren. Bitte zuerst Zuteilung erstellen.'); return; }
    const wb = XLSX.utils.book_new();
    const headers = ['Kürzel','Klasse','Frage erstellen in Block','Fragen beantworten in Blöcken'];
    const ws = XLSX.utils.json_to_sheet(assignRows, { header: headers });
    applyAssignmentWorksheetFormatting(ws, assignRows);
    XLSX.utils.book_append_sheet(wb, ws, 'Zuteilung');
    const fileName = (currentTestNumber != null) ? `3PMo_Zuteilung_Test${String(parseInt(currentTestNumber,10)).padStart(2,'0')}.xlsx` : '3PMo_Zuteilung.xlsx';
    XLSX.writeFile(wb, fileName);
  }

  function renderAssignSummary(rows, blocks) {
    const root = document.getElementById('assignSummary');
    if (!root) return;
    root.innerHTML = '';

    const blockLabels = blocks.map(b => formatBlockLabel(b));

    const buildEmptyCounts = () => {
      const obj = {};
      for (const lbl of blockLabels) obj[lbl] = { create: 0, answer: 0 };
      return obj;
    };

    // Overall counts
    const overall = buildEmptyCounts();
    // Per-class counts
    const classSet = new Set();
    const perClass = new Map(); // class -> counts

    for (const r of rows) {
      const klass = r['Klasse'];
      classSet.add(klass);
      if (!perClass.has(klass)) perClass.set(klass, buildEmptyCounts());
      const cLbl = r['Frage erstellen in Block'];
      if (overall[cLbl]) overall[cLbl].create++;
      if (perClass.get(klass)[cLbl]) perClass.get(klass)[cLbl].create++;

      const answers = String(r['Fragen beantworten in Blöcken'] || '').split('&').map(s => s.trim()).filter(Boolean);
      for (const aLbl of answers) {
        if (overall[aLbl]) overall[aLbl].answer++;
        if (perClass.get(klass)[aLbl]) perClass.get(klass)[aLbl].answer++;
      }
    }

    const makeTable = (titleText, counts) => {
      const section = document.createElement('section');
      section.className = 'summary-section';

      const h = document.createElement('h3');
      h.textContent = titleText;
      section.appendChild(h);

      const wrapper = document.createElement('div');
      wrapper.className = 'table-wrapper';

      const table = document.createElement('table');
      table.className = 'summary-table';
      table.innerHTML = `
        <thead>
          <tr>
            <th>Block</th>
            <th>Frage erstellen</th>
            <th>Fragen beantworten</th>
          </tr>
        </thead>
        <tbody></tbody>
      `;

      const tbody = table.querySelector('tbody');
      let totalCreate = 0;
      let totalAnswer = 0;
      for (const lbl of blockLabels) {
        const { create, answer } = counts[lbl] || { create: 0, answer: 0 };
        totalCreate += create;
        totalAnswer += answer;
        const tr = document.createElement('tr');
        tr.innerHTML = `
          <td>${escapeHtml(lbl)}</td>
          <td>${create}</td>
          <td>${answer}</td>
        `;
        tbody.appendChild(tr);
      }
      const trTotal = document.createElement('tr');
      trTotal.innerHTML = `
        <td><strong>Total</strong></td>
        <td><strong>${totalCreate}</strong></td>
        <td><strong>${totalAnswer}</strong></td>
      `;
      tbody.appendChild(trTotal);

      wrapper.appendChild(table);
      section.appendChild(wrapper);
      return section;
    };

    // Overall
    root.appendChild(makeTable('Zusammenfassung (gesamt)', overall));

    // Per class
    const classes = Array.from(classSet).sort((a, b) => String(a).localeCompare(String(b), 'de-CH', { sensitivity: 'base' }));
    for (const klass of classes) {
      root.appendChild(makeTable(`Zusammenfassung – Klasse ${klass}`, perClass.get(klass)));
    }
  }

  function applyAssignmentWorksheetFormatting(ws, rows) {
    try {
      const headers = ['Kürzel','Klasse','Frage erstellen in Block','Fragen beantworten in Blöcken'];
      // Column widths driven by content lengths; ensure C/D are wide enough
      const cols = headers.map((h, idx) => {
        let maxLen = h.length;
        for (const r of rows) {
          const v = (r[h] == null ? '' : String(r[h]));
          if (v.length > maxLen) maxLen = v.length;
        }
        // Minimum width for columns C and D
        if (idx === 2) maxLen = Math.max(maxLen, 20);
        if (idx === 3) maxLen = Math.max(maxLen, 35);
        return { wch: Math.min(100, maxLen + 2) };
      });
      ws['!cols'] = cols;

      // AutoFilter over header row
      const rangeRef = XLSX.utils.encode_range({ s: { c: 0, r: 0 }, e: { c: 3, r: rows.length } });
      ws['!autofilter'] = { ref: rangeRef };

      // Bold header row (best-effort; some XLSX builds may ignore styles)
      const headerCells = ['A1', 'B1', 'C1', 'D1'];
      for (const addr of headerCells) {
        if (ws[addr]) {
          ws[addr].s = ws[addr].s || {};
          ws[addr].s.font = Object.assign({}, ws[addr].s.font, { bold: true });
        }
      }
    } catch (e) {
      // non-fatal; formatting is optional
    }
  }

  // getSampleHTML removed - now loaded via fetch from samples/studentquiz-sample.html
  function getSampleHTML_DEPRECATED() {
    return `
<table id="categoryquestions" class="question-bank-table generaltable" data-defaultsort="{&quot;mod_studentquiz__question__bank__anonym_creator_name_column-timecreated&quot;:3,&quot;mod_studentquiz__question__bank__question_name_column&quot;:4}"><thead id="yui_3_18_1_1_1756465703414_173"><tr class="qbank-column-list" id="yui_3_18_1_1_1756465703414_172"><th class="header align-top checkbox" scope="col" data-pluginname="core_question\local\bank\checkbox_column" data-name="Alle auswählen" data-columnid="core_question\local\bank\checkbox_column-checkbox_column" style="width: 30px;" id="yui_3_18_1_1_1756465703414_171"> <div class="header-container" id="yui_3_18_1_1_1756465703414_170"> <div class="header-text" id="yui_3_18_1_1_1756465703414_169"> <span class="me-1" title="Fragen für Sammelaktion auswählen" id="yui_3_18_1_1_1756465703414_168"><div class="form-check" id="yui_3_18_1_1_1756465703414_167"> <input id="qbheadercheckbox" name="qbheadercheckbox" type="checkbox" class="" value="1" aria-labelledby="qbheadercheckbox-label" data-action="toggle" data-toggle="master" data-togglegroup="qbank" data-toggle-selectall="Alle auswählen" data-toggle-deselectall="Nichts auswählen"> <label id="qbheadercheckbox-label" for="qbheadercheckbox" class="form-check-label d-block pe-2 accesshide">Nichts auswählen</label> </div></span> </div> </div> </th><th class="header align-top qtype" scope="col" data-pluginname="qbank_viewquestiontype__question_type_column" data-name="T" data-columnid="qbank_viewquestiontype\question_type_column-question_type_column" style="width: 45px;"> <div class="header-container"> <div class="header-text"> </div> </div> <div class="sorters"> <a href="https://moodle.zhaw.ch/mod/studentquiz/view.php?cmid=1723409&amp;cat=519228%2C2302532&amp;id=1723409&amp;group=163617&amp;sortdata%5Bqbank_viewquestiontype__question_type_column%5D=4&amp;sortdata%5Bmod_studentquiz__question__bank__anonym_creator_name_column-timecreated%5D=3&amp;sortdata%5Bmod_studentquiz__question__bank__question_name_column%5D=4" data-sortname="qbank_viewquestiontype__question_type_column" data-sortorder="4" title="Sortierung nach Fragetyp aufsteigend"> T </a> </div> </th><th class="header align-top state" scope="col" data-pluginname="mod_studentquiz__question__bank__state_column" data-name="S" data-columnid="mod_studentquiz\question\bank\state_column-state_column" style="width: 120px;"> <div class="header-container"> <div class="header-text"> </div> </div> <div class="sorters"> <a href="https://moodle.zhaw.ch/mod/studentquiz/view.php?cmid=1723409&amp;cat=519228%2C2302532&amp;id=1723409&amp;group=163617&amp;sortdata%5Bmod_studentquiz__question__bank__state_column%5D=4&amp;sortdata%5Bmod_studentquiz__question__bank__anonym_creator_name_column-timecreated%5D=3&amp;sortdata%5Bmod_studentquiz__question__bank__question_name_column%5D=4" data-sortname="mod_studentquiz__question__bank__state_column" data-sortorder="4" title="Sortierung nach Status aufsteigend"> S </a> </div> </th><th class="header align-top state_pin" scope="col" data-pluginname="mod_studentquiz__question__bank__state_pin_column" data-name="" data-columnid="mod_studentquiz\question\bank\state_pin_column-state_pin_column" style="width: 120px;"> <div class="header-container"> <div class="header-text"> <span class="me-1"></span> </div> </div> </th><th class="header align-top questionname" scope="col" data-pluginname="mod_studentquiz__question__bank__question_name_column" data-name="Frage" data-columnid="mod_studentquiz\question\bank\question_name_column-question_name_column" style="width: 250px;"> <div class="header-container"> <div class="header-text"> </div> </div> <div class="sorters"> </div> </th><th class="header align-top pe-3 editmenu" scope="col" data-pluginname="mod_studentquiz__question__bank__sq_edit_menu_column" data-name="Aktionen" data-columnid="mod_studentquiz\question\bank\sq_edit_menu_column-sq_edit_menu_column" style="width: 120px;"> <div class="header-container"> <div class="header-text"> <span class="me-1">Aktionen</span> </div> </div> </th><th class="header align-top pe-3 questionversionnumber" scope="col" data-pluginname="qbank_history__version_number_column" data-name="Version" data-columnid="qbank_history\version_number_column-version_number_column" style="width: 120px;"> <div class="header-container"> <div class="header-text"> </div> </div> <div class="sorters"> </div> </th><th class="header align-top creatorname" scope="col" data-pluginname="mod_studentquiz__question__bank__anonym_creator_name_column" data-name="Erstellt von" data-columnid="mod_studentquiz\question\bank\anonym_creator_name_column-anonym_creator_name_column" style="width: 120px;"> <div class="header-container"> <div class="header-text"> <div class="title me-1">Erstellt von</div> </div> </div> <div class="sorters"> / / </div> </th><th class="header align-top tags" scope="col" data-pluginname="mod_studentquiz__question__bank__tag_column" data-name="Tags" data-columnid="mod_studentquiz\question\bank\tag_column-tag_column" style="width: 120px;"> <div class="header-container"> <div class="header-text"> </div> </div> <div class="sorters"> </div> </th><th class="header align-top attempts" scope="col" data-pluginname="mod_studentquiz__question__bank__attempts_column" data-name="Meine Versuche" data-columnid="mod_studentquiz\question\bank\attempts_column-attempts_column" style="width: 120px;"> <div class="header-container"> <div class="header-text"> <div class="title me-1">Meine Versuche</div> </div> </div> <div class="sorters"> / </div> </th><th class="header align-top difficultylevel" scope="col" data-pluginname="mod_studentquiz__question__bank__difficulty_level_column" data-name="Schwierigkeit" data-columnid="mod_studentquiz\question\bank\\difficulty_level_column-difficulty_level_column" style="width: 120px;"> <div class="header-container"> <div class="header-text"> <div class="title me-1">Schwierigkeit</div> </div> </div> <div class="sorters"> / </div> </th><th class="header align-top rates" scope="col" data-pluginname="mod_studentquiz__question__bank__rate_column" data-name="Bewertung" data-columnid="mod_studentquiz\question\bank\rate_column-rate_column" style="width: 120px;"> <div class="header-container"> <div class="header-text"> <div class="title me-1">Bewertung</div> </div> </div> <div class="sorters"> / </div> </th><th class="header align-top comment" scope="col" data-pluginname="mod_studentquiz__question__bank__comment_column" data-name="Kommentare" data-columnid="mod_studentquiz\question\bank\comment_column-comment_column" style="width: 120px;"> <div class="header-container"> <div class="header-text"> </div> </div> <div class="sorters"> </div> </th></tr></thead><tbody><tr class="r0"><td class="checkbox" data-columnid="core_question\local\bank\checkbox_column-checkbox_column"><input id="checkq20458800" name="q20458800" type="checkbox" value="1" data-action="toggle" data-toggle="slave" data-togglegroup="qbank"> <label for="checkq20458800" class="accesshide">Auswahl</label></td><td class="qtype" data-columnid="qbank_viewquestiontype\question_type_column-question_type_column"><img class="icon " title="Kprim (ETH)" alt="Kprim (ETH)" src="https://moodle.zhaw.ch/theme/image.php/boost_union/qtype_kprime/1756457962/icon"></td><td class="state" data-columnid="mod_studentquiz\question\bank\state_column-state_column"></td><td class="state_pin" data-columnid="mod_studentquiz\question\bank\state_pin_column-state_pin_column"></td><td class="questionname" data-columnid="mod_studentquiz\question\bank\question_name_column-question_name_column"><label for="checkq20458800">Frage 01, einfach</label></td><td class="pe-3 editmenu" data-columnid="mod_studentquiz\question\bank\sq_edit_menu_column-sq_edit_menu_column"><div class="action-menu moodle-actionmenu" id="action-menu-1" data-enhance="moodle-core-actionmenu"> <div class="menubar d-flex " id="action-menu-1-menubar"> <div class="action-menu-trigger"> <div class="dropdown"> <div class="dropdown-menu menu dropdown-menu-right" id="action-menu-1-menu" data-rel="menu-content" aria-labelledby="action-menu-toggle-1" role="menu"> </div> </div> </div> </div> </div></td><td class="pe-3 questionversionnumber" data-columnid="qbank_history\version_number_column-version_number_column">v1</td><td class="creatorname" data-columnid="mod_studentquiz\question\bank\anonym_creator_name_column-anonym_creator_name_column"><span></span><br><span class="date">25. August 2025, 14:29</span></td><td class="tags" data-columnid="mod_studentquiz\question\bank\tag_column-tag_column">n.a.</td><td class="attempts" data-columnid="mod_studentquiz\question\bank\attempts_column-attempts_column"><span class="pratice_info" tabindex="0" aria-label="Diese Frage wurde noch nie versucht.">n.a.&nbsp;|&nbsp;n.a.</span></td><td class="difficultylevel" data-columnid="mod_studentquiz\question\bank\\difficulty_level_column-difficulty_level_column"><span class="mod_studentquiz_difficulty" data-mydifficulty="0" title=""></span></td><td class="rates" data-columnid="mod_studentquiz\question\bank\rate_column-rate_column"><span class="mod_studentquiz_ratingbar" title="">n.a.</span></td><td class="comment" data-columnid="mod_studentquiz\question\bank\comment_column-comment_column"><span class="public-comment badge badge-secondary" title="Anzahl an öffentlichen Kommentaren. Ein blauer Hintergrund bedeutet, dass Sie mindest einen ungelesenen Kommentar haben."> n.a. <span class="sr-only">Öffentlich Kommentare</span> </span></td></tr><tr class="r1"><td class="checkbox" data-columnid="core_question\local\bank\checkbox_column-checkbox_column"><input id="checkq20451333" name="q20451333" type="checkbox" value="1" data-action="toggle" data-toggle="slave" data-togglegroup="qbank"> <label for="checkq20451333" class="accesshide">Auswahl</label></td><td class="qtype" data-columnid="qbank_viewquestiontype\question_type_column-question_type_column"><img class="icon " title="Kprim (ETH)" alt="Kprim (ETH)" src="https://moodle.zhaw.ch/theme/image.php/boost_union/qtype_kprime/1756457962/icon"></td><td class="state" data-columnid="mod_studentquiz\question\bank\state_column-state_column"></td><td class="state_pin" data-columnid="mod_studentquiz\question\bank\state_pin_column-state_pin_column"></td><td class="questionname" data-columnid="mod_studentquiz\question\bank\question_name_column-question_name_column"><label for="checkq20451333">Gitarren im Musikbusiness</label></td><td class="pe-3 editmenu" data-columnid="mod_studentquiz\question\bank\sq_edit_menu_column-sq_edit_menu_column"><div class="action-menu moodle-actionmenu" id="action-menu-2" data-enhance="moodle-core-actionmenu"> <div class="menubar d-flex " id="action-menu-2-menubar"> <div class="action-menu-trigger"> <div class="dropdown"> <div class="dropdown-menu menu dropdown-menu-right" id="action-menu-2-menu" data-rel="menu-content" aria-labelledby="action-menu-toggle-2" role="menu"> </div> </div> </div> </div> </div></td><td class="pe-3 questionversionnumber" data-columnid="qbank_history\version_number_column-version_number_column">v1</td><td class="creatorname" data-columnid="mod_studentquiz\question\bank\anonym_creator_name_column-anonym_creator_name_column"><span></span><br><span class="date">21. August 2025, 10:42</span></td><td class="tags" data-columnid="mod_studentquiz\question\bank\tag_column-tag_column">n.a.</td><td class="attempts" data-columnid="mod_studentquiz\question\bank\attempts_column-attempts_column"><span class="pratice_info" tabindex="0" aria-label="Beim letzten Versuch falsch">1&nbsp;|&nbsp;✗</span></td><td class="difficultylevel" data-columnid="mod_studentquiz\question\bank\\difficulty_level_column-difficulty_level_column"><span class="mod_studentquiz_difficulty" data-difficultylevel="0.78" data-mydifficulty="1.00" title="Community Schwierigkeit: 78% , Meine Schwierigkeit: 100%"></span></td><td class="rates" data-columnid="mod_studentquiz\question\bank\rate_column-rate_column"><span class="mod_studentquiz_ratingbar" data-rate="4.33" data-myrate="4" title="Community Bewertung: 4.33 , Meine Bewertung: 4"></span></td><td class="comment" data-columnid="mod_studentquiz\question\bank\comment_column-comment_column"><span class="public-comment badge badge-secondary" title="Anzahl an öffentlichen Kommentaren. Ein blauer Hintergrund bedeutet, dass Sie mindest einen ungelesenen Kommentar haben."> 2 <span class="sr-only">Öffentlich Kommentare</span> </span></td></tr><tr class="r0"><td class="checkbox" data-columnid="core_question\local\bank\checkbox_column-checkbox_column"><input id="checkq20450640" name="q20450640" type="checkbox" value="1" data-action="toggle" data-toggle="slave" data-togglegroup="qbank"> <label for="checkq20450640" class="accesshide">Auswahl</label></td><td class="qtype" data-columnid="qbank_viewquestiontype\question_type_column-question_type_column"><img class="icon " title="Kprim (ETH)" alt="Kprim (ETH)" src="https://moodle.zhaw.ch/theme/image.php/boost_union/qtype_kprime/1756457962/icon"></td><td class="state" data-columnid="mod_studentquiz\question\bank\state_column-state_column"></td><td class="state_pin" data-columnid="mod_studentquiz\question\bank\state_pin_column-state_pin_column"></td><td class="questionname" data-columnid="mod_studentquiz\question\bank\question_name_column-question_name_column"><label for="checkq20450640">IKEA</label></td><td class="pe-3 editmenu" data-columnid="mod_studentquiz\question\bank\sq_edit_menu_column-sq_edit_menu_column"><div class="action-menu moodle-actionmenu" id="action-menu-3" data-enhance="moodle-core-actionmenu"> <div class="menubar d-flex " id="action-menu-3-menubar"> <div class="action-menu-trigger"> <div class="dropdown"> <div class="dropdown-menu menu dropdown-menu-right" id="action-menu-3-menu" data-rel="menu-content" aria-labelledby="action-menu-toggle-3" role="menu"> </div> </div> </div> </div> </div></td><td class="pe-3 questionversionnumber" data-columnid="qbank_history\version_number_column-version_number_column">v1</td><td class="creatorname" data-columnid="mod_studentquiz\question\bank\anonym_creator_name_column-anonym_creator_name_column"><span></span><br><span class="date">20. August 2025, 15:20</span></td><td class="tags" data-columnid="mod_studentquiz\question\bank\tag_column-tag_column">n.a.</td><td class="attempts" data-columnid="mod_studentquiz\question\bank\attempts_column-attempts_column"><span class="pratice_info" tabindex="0" aria-label="Diese Frage wurde noch nie versucht.">n.a.&nbsp;|&nbsp;n.a.</span></td><td class="difficultylevel" data-columnid="mod_studentquiz\question\bank\\difficulty_level_column-difficulty_level_column"><span class="mod_studentquiz_difficulty" data-difficultylevel="1.00" data-mydifficulty="0" title="Community Schwierigkeit: 100% , Meine Schwierigkeit: n.a."></span></td><td class="rates" data-columnid="mod_studentquiz\question\bank\rate_column-rate_column"><span class="mod_studentquiz_ratingbar" data-rate="2.50" title="Community Bewertung: 2.5 , Meine Bewertung: n.a."></span></td><td class="comment" data-columnid="mod_studentquiz\question\bank\comment_column-comment_column"><span class="public-comment badge badge-primary" title="Anzahl an öffentlichen Kommentaren. Ein blauer Hintergrund bedeutet, dass Sie mindest einen ungelesenen Kommentar haben."> 1 <span class="sr-only">Öffentlich Kommentar(inklusive ungelesener)</span> </span></td></tr><tr class="r1"><td class="checkbox" data-columnid="core_question\local\bank\checkbox_column-checkbox_column"><input id="checkq20450639" name="q20450639" type="checkbox" value="1" data-action="toggle" data-toggle="slave" data-togglegroup="qbank"> <label for="checkq20450639" class="accesshide">Auswahl</label></td><td class="qtype" data-columnid="qbank_viewquestiontype\question_type_column-question_type_column"><img class="icon " title="Kprim (ETH)" alt="Kprim (ETH)" src="https://moodle.zhaw.ch/theme/image.php/boost_union/qtype_kprime/1756457962/icon"></td><td class="state" data-columnid="mod_studentquiz\question\bank\state_column-state_column"></td><td class="state_pin" data-columnid="mod_studentquiz\question\bank\state_pin_column-state_pin_column"></td><td class="questionname" data-columnid="mod_studentquiz\question\bank\question_name_column-question_name_column"><label for="checkq20450639">Laubfrösche</label></td><td class="pe-3 editmenu" data-columnid="mod_studentquiz\question\bank\sq_edit_menu_column-sq_edit_menu_column"><div class="action-menu moodle-actionmenu" id="action-menu-4" data-enhance="moodle-core-actionmenu"> <div class="menubar d-flex " id="action-menu-4-menubar"> <div class="action-menu-trigger"> <div class="dropdown"> <div class="dropdown-menu menu dropdown-menu-right" id="action-menu-4-menu" data-rel="menu-content" aria-labelledby="action-menu-toggle-4" role="menu"> </div> </div> </div> </div> </div></td><td class="pe-3 questionversionnumber" data-columnid="qbank_history\version_number_column-version_number_column">v1</td><td class="creatorname" data-columnid="mod_studentquiz\question\bank\anonym_creator_name_column-anonym_creator_name_column"><span></span><br><span class="date">20. August 2025, 15:15</span></td><td class="tags" data-columnid="mod_studentquiz\question\bank\tag_column-tag_column">n.a.</td><td class="attempts" data-columnid="mod_studentquiz\question\bank\attempts_column-attempts_column"><span class="pratice_info" tabindex="0" aria-label="Diese Frage wurde noch nie versucht.">n.a.&nbsp;|&nbsp;n.a.</span></td><td class="difficultylevel" data-columnid="mod_studentquiz\question\bank\\difficulty_level_column-difficulty_level_column"><span class="mod_studentquiz_difficulty" data-difficultylevel="0.50" data-mydifficulty="0" title="Community Schwierigkeit: 50% , Meine Schwierigkeit: n.a."></span></td><td class="rates" data-columnid="mod_studentquiz\question\bank\rate_column-rate_column"><span class="mod_studentquiz_ratingbar" data-rate="3.50" title="Community Bewertung: 3.5 , Meine Bewertung: n.a."></span></td><td class="comment" data-columnid="mod_studentquiz\question\bank\comment_column-comment_column"><span class="public-comment badge badge-primary" title="Anzahl an öffentlichen Kommentaren. Ein blauer Hintergrund bedeutet, dass Sie mindest einen ungelesenen Kommentar haben."> 1 <span class="sr-only">Öffentlich Kommentar(inklusive ungelesener)</span> </span></td></tr><tr class="r0"><td class="checkbox" data-columnid="core_question\local\bank\checkbox_column-checkbox_column"><input id="checkq20450638" name="q20450638" type="checkbox" value="1" data-action="toggle" data-toggle="slave" data-togglegroup="qbank"> <label for="checkq20450638" class="accesshide">Auswahl</label></td><td class="qtype" data-columnid="qbank_viewquestiontype\question_type_column-question_type_column"><img class="icon " title="Kprim (ETH)" alt="Kprim (ETH)" src="https://moodle.zhaw.ch/theme/image.php/boost_union/qtype_kprime/1756457962/icon"></td><td class="state" data-columnid="mod_studentquiz\question\bank\state_column-state_column"></td><td class="state_pin" data-columnid="mod_studentquiz\question\bank\state_pin_column-state_pin_column"></td><td class="questionname" data-columnid="mod_studentquiz\question\bank\question_name_column-question_name_column"><label for="checkq20450638">Wertschöpfungsanalyse</label></td><td class="pe-3 editmenu" data-columnid="mod_studentquiz\question\bank\sq_edit_menu_column-sq_edit_menu_column"><div class="action-menu moodle-actionmenu" id="action-menu-5" data-enhance="moodle-core-actionmenu"> <div class="menubar d-flex " id="action-menu-5-menubar"> <div class="action-menu-trigger"> <div class="dropdown"> <div class="dropdown-menu menu dropdown-menu-right" id="action-menu-5-menu" data-rel="menu-content" aria-labelledby="action-menu-toggle-5" role="menu"> </div> </div> </div> </div> </div></td><td class="pe-3 questionversionnumber" data-columnid="qbank_history\version_number_column-version_number_column">v1</td><td class="creatorname" data-columnid="mod_studentquiz\question\bank\anonym_creator_name_column-anonym_creator_name_column"><span></span><br><span class="date">20. August 2025, 15:10</span></td><td class="tags" data-columnid="mod_studentquiz\question\bank\tag_column-tag_column">n.a.</td><td class="attempts" data-columnid="mod_studentquiz\question\bank\attempts_column-attempts_column"><span class="pratice_info" tabindex="0" aria-label="Diese Frage wurde noch nie versucht.">n.a.&nbsp;|&nbsp;n.a.</span></td><td class="difficultylevel" data-columnid="mod_studentquiz\question\bank\\difficulty_level_column-difficulty_level_column"><span class="mod_studentquiz_difficulty" data-mydifficulty="0" title=""></span></td><td class="rates" data-columnid="mod_studentquiz\question\bank\rate_column-rate_column"><span class="mod_studentquiz_ratingbar" title="">n.a.</span></td><td class="comment" data-columnid="mod_studentquiz\question\bank\comment_column-comment_column"><span class="public-comment badge badge-secondary" title="Anzahl an öffentlichen Kommentaren. Ein blauer Hintergrund bedeutet, dass Sie mindest einen ungelesenen Kommentar haben."> n.a. <span class="sr-only">Öffentlich Kommentare</span> </span></td></tr></tbody></table>`;
  }

  // --- Helper Tables Functions ---
  
  async function uploadStudentHelperTable() {
    const file = els.studentHelperFile && els.studentHelperFile.files && els.studentHelperFile.files[0];
    if (!file) {
      els.statusStudentHelper.textContent = 'Bitte Datei auswählen.';
      return;
    }
    
    els.statusStudentHelper.textContent = 'Lade...';
    
    try {
      const data = await readExcelFile(file);
      if (!data || data.length === 0) {
        els.statusStudentHelper.textContent = 'Fehler: Keine Daten gefunden.';
        return;
      }
      
      // Expected columns: kuerzel, "Vorname Nachname"
      studentHelperData = data.map(row => ({
        kuerzel: String(row.kuerzel || row.Kuerzel || '').trim().toLowerCase(),
        fullname: String(row['Vorname Nachname'] || row.fullname || row.name || '').trim()
      })).filter(r => r.kuerzel && r.fullname);
      
      if (studentHelperData.length === 0) {
        els.statusStudentHelper.textContent = 'Fehler: Spalten "kuerzel" und "Vorname Nachname" nicht gefunden.';
        return;
      }
      
      // Save to localStorage
      localStorage.setItem('3pmo_student_helper', JSON.stringify(studentHelperData));
      
      els.statusStudentHelper.textContent = `✓ ${studentHelperData.length} Studierende geladen.`;
      els.btnClearStudentHelper.style.display = '';
      els.btnUploadStudentHelper.textContent = 'Bestehende Angaben ersetzen';
    } catch (e) {
      els.statusStudentHelper.textContent = 'Fehler beim Laden der Datei.';
      console.error(e);
    }
  }
  
  function clearStudentHelperTable() {
    studentHelperData = [];
    localStorage.removeItem('3pmo_student_helper');
    els.statusStudentHelper.textContent = 'Hilfstabelle wurde gelöscht.';
    els.btnClearStudentHelper.style.display = 'none';
    els.btnUploadStudentHelper.textContent = 'Hochladen';
    if (els.studentHelperFile) els.studentHelperFile.value = '';
  }
  
  async function uploadAssignmentTable() {
    if (!currentTestName) {
      els.statusAssignment.textContent = 'Fehler: Kein Test ausgewählt. Bitte erst Test erstellen/laden.';
      return;
    }
    
    const file = els.assignmentFile && els.assignmentFile.files && els.assignmentFile.files[0];
    if (!file) {
      els.statusAssignment.textContent = 'Bitte Datei auswählen.';
      return;
    }
    
    els.statusAssignment.textContent = 'Lade...';
    
    try {
      const data = await readExcelFile(file);
      if (!data || data.length === 0) {
        els.statusAssignment.textContent = 'Fehler: Keine Daten gefunden.';
        return;
      }
      
      // Expected columns: Kürzel, Frage erstellen in Block, Fragen beantworten in Blöcken
      assignmentData = data.map(row => ({
        kuerzel: String(row['Kürzel'] || row.kuerzel || '').trim().toLowerCase(),
        createBlock: String(row['Frage erstellen in Block'] || '').trim(),
        answerBlocks: String(row['Fragen beantworten in Blöcken'] || '').trim()
      })).filter(r => r.kuerzel && r.createBlock);
      
      if (assignmentData.length === 0) {
        els.statusAssignment.textContent = 'Fehler: Erwartete Spalten nicht gefunden.';
        return;
      }
      
      // Save to localStorage with test name
      const key = `3pmo_assignment_${currentTestName}`;
      localStorage.setItem(key, JSON.stringify(assignmentData));
      
      els.statusAssignment.textContent = `✓ Zuteilung für Test "${currentTestName}" geladen (${assignmentData.length} Einträge).`;
      els.btnClearAssignment.style.display = '';
      els.btnUploadAssignment.textContent = 'Zuteilung ersetzen';
      
      // Trigger re-validation if grades are already generated
      if (combinedGradesData && combinedGradesData.length > 0) {
        validateAssignments();
        renderCombinedGradesTable(combinedGradesData);
        if (isFilterActive) filterManualReviewOnly();
      }
    } catch (e) {
      els.statusAssignment.textContent = 'Fehler beim Laden der Datei.';
      console.error(e);
    }
  }
  
  function clearAssignmentTable() {
    if (!currentTestName) return;
    assignmentData = [];
    const key = `3pmo_assignment_${currentTestName}`;
    localStorage.removeItem(key);
    els.statusAssignment.textContent = 'Zuteilungstabelle wurde gelöscht.';
    els.btnClearAssignment.style.display = 'none';
    els.btnUploadAssignment.textContent = 'Hochladen';
    if (els.assignmentFile) els.assignmentFile.value = '';
    
    // Trigger re-validation
    if (combinedGradesData && combinedGradesData.length > 0) {
      validateAssignments();
      renderCombinedGradesTable(combinedGradesData);
      if (isFilterActive) filterManualReviewOnly();
    }
  }
  
  function readExcelFile(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = (e) => {
        try {
          const data = new Uint8Array(e.target.result);
          const workbook = XLSX.read(data, { type: 'array' });
          const sheetName = workbook.SheetNames[0];
          const worksheet = workbook.Sheets[sheetName];
          const json = XLSX.utils.sheet_to_json(worksheet);
          resolve(json);
        } catch (err) {
          reject(err);
        }
      };
      reader.onerror = reject;
      reader.readAsArrayBuffer(file);
    });
  }
  
  function loadHelperTablesFromStorage() {
    // Load student helper table
    const helperStr = localStorage.getItem('3pmo_student_helper');
    if (helperStr) {
      try {
        studentHelperData = JSON.parse(helperStr);
        if (studentHelperData && studentHelperData.length > 0) {
          els.statusStudentHelper.textContent = `✓ ${studentHelperData.length} Studierende geladen (aus LocalStorage).`;
          els.btnClearStudentHelper.style.display = '';
          els.btnUploadStudentHelper.textContent = 'Bestehende Angaben ersetzen';
        }
      } catch (e) {
        console.error('Error loading student helper data:', e);
      }
    } else {
      els.statusStudentHelper.textContent = '⚠️ Hilfstabelle fehlt - Datei muss hochgeladen werden.';
    }
    
    // Load assignment table for current test
    if (currentTestName) {
      const key = `3pmo_assignment_${currentTestName}`;
      const assignStr = localStorage.getItem(key);
      if (assignStr) {
        try {
          assignmentData = JSON.parse(assignStr);
          if (assignmentData && assignmentData.length > 0) {
            els.statusAssignment.textContent = `✓ Zuteilung für Test "${currentTestName}" geladen (${assignmentData.length} Einträge).`;
            els.btnClearAssignment.style.display = '';
            els.btnUploadAssignment.textContent = 'Zuteilung ersetzen';
          }
        } catch (e) {
          console.error('Error loading assignment data:', e);
        }
      } else {
        els.statusAssignment.textContent = `⚠️ Zuteilung für Test "${currentTestName}" fehlt - Datei muss hochgeladen werden.`;
      }
    }
  }

  // --- Validation Functions ---

  function validateAssignments() {
    if (!assignmentData || assignmentData.length === 0) {
      // No assignment data, clear validation fields
      combinedGradesData.forEach(student => {
        student.wrong_block = '';
        student.expected_responders = '';
      });
      return;
    }
    
    combinedGradesData.forEach(student => {
      const kuerzel = getKuerzelFromName(student.student_name);
      const assignment = assignmentData.find(a => a.kuerzel === kuerzel);
      
      if (!assignment) {
        student.wrong_block = '';
        student.expected_responders = '';
        return;
      }
      
      // Check if question block matches
      const questionName = student.question_name || '';
      const expectedBlock = assignment.createBlock;
      const actualBlock = extractBlockPrefix(questionName);
      
      if (actualBlock && expectedBlock && actualBlock !== expectedBlock) {
        student.wrong_block = 'YES';
        student.requiresManualReview = true;
        if (!student.reviewFlags) student.reviewFlags = [];
        if (!student.reviewFlags.includes('wrong_block')) {
          student.reviewFlags.push('wrong_block');
        }
        // Add to justification
        const msg = `Falsch: Frage in Block "${actualBlock}" statt "${expectedBlock}".`;
        if (!student.justification || !student.justification.includes(msg)) {
          student.justification = student.justification ? (student.justification + ' ' + msg) : msg;
        }
      } else {
        student.wrong_block = '';
      }
      
      // Find expected responders if comments < 3
      const comments = parseFloat(student.total_comments) || 0;
      if (comments < 3 && expectedBlock) {
        const responders = findExpectedResponders(expectedBlock);
        student.expected_responders = responders.join('\n');
      } else {
        student.expected_responders = '';
      }
    });
  }

  function getKuerzelFromName(fullname) {
    if (!studentHelperData || studentHelperData.length === 0) return '';
    const normalized = fullname.trim().toLowerCase();
    const match = studentHelperData.find(s => s.fullname.toLowerCase() === normalized);
    return match ? match.kuerzel : '';
  }

  function extractBlockPrefix(questionName) {
    // Extract prefix like "SW01F13-17" from question name
    // Pattern: starts with letters/numbers followed by F and numbers-numbers
    const match = questionName.match(/^([A-Z0-9]+F\d+-\d+)/i);
    return match ? match[1] : '';
  }

  function findExpectedResponders(blockPrefix) {
    if (!assignmentData || assignmentData.length === 0) return [];
    if (!studentHelperData || studentHelperData.length === 0) return [];
    
    const responders = [];
    assignmentData.forEach(assignment => {
      if (assignment.answerBlocks.includes(blockPrefix)) {
        const kuerzel = assignment.kuerzel;
        const student = studentHelperData.find(s => s.kuerzel === kuerzel);
        if (student) {
          responders.push(student.fullname);
        }
      }
    });
    return responders;
  }


  // Event bindings
  if (els.btnParse) els.btnParse.addEventListener('click', parseHTML);
  if (els.btnClear) els.btnClear.addEventListener('click', clearAll);
  if (els.btnDownload) els.btnDownload.addEventListener('click', downloadExcel);

  // Ranking events
  if (els.btnParseRanking) els.btnParseRanking.addEventListener('click', parseRankingHTML);
  if (els.btnClearRanking) els.btnClearRanking.addEventListener('click', clearRanking);
  if (els.btnDownloadRanking) els.btnDownloadRanking.addEventListener('click', downloadRankingExcel);

  // Grades events (Bereich 4)
  if (els.btnGenerateGrades) els.btnGenerateGrades.addEventListener('click', generateCombinedGrades);
  if (els.btnDownloadGrades) els.btnDownloadGrades.addEventListener('click', downloadCombinedGradesExcel);
  if (els.btnFilterManual) els.btnFilterManual.addEventListener('click', filterManualReviewOnly);
  if (els.btnShowAll) els.btnShowAll.addEventListener('click', showAllGrades);
  
  // Sortable table headers
  document.querySelectorAll('#gradesTable th.sortable').forEach(th => {
    th.addEventListener('click', () => {
      const column = th.dataset.sort;
      if (column) sortGradesTable(column);
    });
  });
  
  // Table height slider
  if (els.tableHeightSlider && els.gradesTableWrapper && els.tableHeightValue) {
    els.tableHeightSlider.addEventListener('input', (e) => {
      const height = e.target.value;
      els.gradesTableWrapper.style.maxHeight = height + 'px';
      els.tableHeightValue.textContent = height;
      // Save preference to localStorage
      localStorage.setItem('3pmo_table_height', height);
    });
    
    // Load saved height preference
    const savedHeight = localStorage.getItem('3pmo_table_height');
    if (savedHeight) {
      els.tableHeightSlider.value = savedHeight;
      els.gradesTableWrapper.style.maxHeight = savedHeight + 'px';
      els.tableHeightValue.textContent = savedHeight;
    }
  }

  // Test management events
  if (els.btnNewTest) els.btnNewTest.addEventListener('click', createNewTest);
  if (els.btnLoadTest) els.btnLoadTest.addEventListener('click', loadExistingTest);
  if (els.btnSaveTest) els.btnSaveTest.addEventListener('click', exportTest);
  if (els.btnImportTest) els.btnImportTest.addEventListener('click', importTest);
  if (els.testFileInput) els.testFileInput.addEventListener('change', handleTestFileImport);

  // Tabs
  if (els.tabSQ && els.tabAssign) {
    els.tabSQ.addEventListener('click', () => showPage('sq'));
    els.tabAssign.addEventListener('click', () => showPage('assign'));
  }

  // Assignment actions
  if (els.btnTemplate) {
    els.btnTemplate.addEventListener('click', downloadStudentsTemplate);
  }
  if (els.btnAssign) {
    els.btnAssign.addEventListener('click', handleAssignClick);
  }
  if (els.btnDownloadAssign) {
    els.btnDownloadAssign.addEventListener('click', downloadAssignmentExcel);
  
  // Helper tables events
  if (els.btnUploadStudentHelper) els.btnUploadStudentHelper.addEventListener('click', uploadStudentHelperTable);
  if (els.btnClearStudentHelper) els.btnClearStudentHelper.addEventListener('click', clearStudentHelperTable);
  if (els.btnUploadAssignment) els.btnUploadAssignment.addEventListener('click', uploadAssignmentTable);
  if (els.btnClearAssignment) els.btnClearAssignment.addEventListener('click', clearAssignmentTable);
  }

  // Load current test on startup
  const lastTest = getCurrentTestName();
  if (lastTest && els.testNameInput) {
    els.testNameInput.value = lastTest;
    currentTestName = lastTest;
    setTestStatus(`Aktiver Test: "${lastTest}"`);
  }
  
  // Load helper tables from storage
  loadHelperTablesFromStorage();
})();
