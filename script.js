(function(){
  const $ = (sel, root=document) => root.querySelector(sel);
  const $$ = (sel, root=document) => Array.from(root.querySelectorAll(sel));

  const els = {
    input: $('#htmlInput'),
    btnSample: $('#btnSample'),
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
    btnSampleRanking: $('#btnSampleRanking'),
    btnParseRanking: $('#btnParseRanking'),
    btnClearRanking: $('#btnClearRanking'),
    btnDownloadRanking: $('#btnDownloadRanking'),
    statusRanking: $('#statusRanking'),
    rankingTableBody: $('#previewTableRanking tbody'),

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

  function insertSample() {
    els.input.value = getSampleHTML().trim();
    setStatus('Beispiel eingefügt.');
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
  }

  function renderTable(rows) {
    els.tableBody.innerHTML = '';
    rows.forEach(r => {
      const tr = document.createElement('tr');
      tr.innerHTML = `
        <td>${escapeHtml(r.questionname)}</td>
        <td>${escapeHtml(r.creatorname)}</td>
        <td>${escapeHtml(r.difficultylevel)}</td>
        <td>${escapeHtml(r.rate)}</td>
        <td>${r.comments != null ? r.comments : ''}</td>
      `;
      els.tableBody.appendChild(tr);
    });
  }

  function escapeHtml(s) {
    return String(s ?? '')
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
  function insertSampleRanking() {
    if (!els.inputRanking) return;
    els.inputRanking.value = getSampleRankingHTML().trim();
    setRankingStatus('Beispiel eingefügt.');
  }

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
      const name = stripTags(nameHTML);

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
  }

  function renderRankingTable(rows) {
    if (!els.rankingTableBody) return;
    els.rankingTableBody.innerHTML = '';
    rows.forEach(r => {
      const tr = document.createElement('tr');
      tr.innerHTML = `
        <td>${escapeHtml(r.student_name)}</td>
        <td>${r.published_question_points ?? ''}</td>
        <td>${r.rating_points ?? ''}</td>
        <td>${r.correct_answers_points ?? ''}</td>
        <td>${r.false_answers_points ?? ''}</td>
      `;
      els.rankingTableBody.appendChild(tr);
    });
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

  function getSampleRankingHTML() {
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

  function getSampleHTML() {
    return `
<table id="categoryquestions" class="question-bank-table generaltable" data-defaultsort="{&quot;mod_studentquiz__question__bank__anonym_creator_name_column-timecreated&quot;:3,&quot;mod_studentquiz__question__bank__question_name_column&quot;:4}"><thead id="yui_3_18_1_1_1756465703414_173"><tr class="qbank-column-list" id="yui_3_18_1_1_1756465703414_172"><th class="header align-top checkbox" scope="col" data-pluginname="core_question\local\bank\checkbox_column" data-name="Alle auswählen" data-columnid="core_question\local\bank\checkbox_column-checkbox_column" style="width: 30px;" id="yui_3_18_1_1_1756465703414_171"> <div class="header-container" id="yui_3_18_1_1_1756465703414_170"> <div class="header-text" id="yui_3_18_1_1_1756465703414_169"> <span class="me-1" title="Fragen für Sammelaktion auswählen" id="yui_3_18_1_1_1756465703414_168"><div class="form-check" id="yui_3_18_1_1_1756465703414_167"> <input id="qbheadercheckbox" name="qbheadercheckbox" type="checkbox" class="" value="1" aria-labelledby="qbheadercheckbox-label" data-action="toggle" data-toggle="master" data-togglegroup="qbank" data-toggle-selectall="Alle auswählen" data-toggle-deselectall="Nichts auswählen"> <label id="qbheadercheckbox-label" for="qbheadercheckbox" class="form-check-label d-block pe-2 accesshide">Nichts auswählen</label> </div></span> </div> </div> </th><th class="header align-top qtype" scope="col" data-pluginname="qbank_viewquestiontype__question_type_column" data-name="T" data-columnid="qbank_viewquestiontype\question_type_column-question_type_column" style="width: 45px;"> <div class="header-container"> <div class="header-text"> </div> </div> <div class="sorters"> <a href="https://moodle.zhaw.ch/mod/studentquiz/view.php?cmid=1723409&amp;cat=519228%2C2302532&amp;id=1723409&amp;group=163617&amp;sortdata%5Bqbank_viewquestiontype__question_type_column%5D=4&amp;sortdata%5Bmod_studentquiz__question__bank__anonym_creator_name_column-timecreated%5D=3&amp;sortdata%5Bmod_studentquiz__question__bank__question_name_column%5D=4" data-sortname="qbank_viewquestiontype__question_type_column" data-sortorder="4" title="Sortierung nach Fragetyp aufsteigend"> T </a> </div> </th><th class="header align-top state" scope="col" data-pluginname="mod_studentquiz__question__bank__state_column" data-name="S" data-columnid="mod_studentquiz\question\bank\state_column-state_column" style="width: 120px;"> <div class="header-container"> <div class="header-text"> </div> </div> <div class="sorters"> <a href="https://moodle.zhaw.ch/mod/studentquiz/view.php?cmid=1723409&amp;cat=519228%2C2302532&amp;id=1723409&amp;group=163617&amp;sortdata%5Bmod_studentquiz__question__bank__state_column%5D=4&amp;sortdata%5Bmod_studentquiz__question__bank__anonym_creator_name_column-timecreated%5D=3&amp;sortdata%5Bmod_studentquiz__question__bank__question_name_column%5D=4" data-sortname="mod_studentquiz__question__bank__state_column" data-sortorder="4" title="Sortierung nach Status aufsteigend"> S </a> </div> </th><th class="header align-top state_pin" scope="col" data-pluginname="mod_studentquiz__question__bank__state_pin_column" data-name="" data-columnid="mod_studentquiz\question\bank\state_pin_column-state_pin_column" style="width: 120px;"> <div class="header-container"> <div class="header-text"> <span class="me-1"></span> </div> </div> </th><th class="header align-top questionname" scope="col" data-pluginname="mod_studentquiz__question__bank__question_name_column" data-name="Frage" data-columnid="mod_studentquiz\question\bank\question_name_column-question_name_column" style="width: 250px;"> <div class="header-container"> <div class="header-text"> </div> </div> <div class="sorters"> </div> </th><th class="header align-top pe-3 editmenu" scope="col" data-pluginname="mod_studentquiz__question__bank__sq_edit_menu_column" data-name="Aktionen" data-columnid="mod_studentquiz\question\bank\sq_edit_menu_column-sq_edit_menu_column" style="width: 120px;"> <div class="header-container"> <div class="header-text"> <span class="me-1">Aktionen</span> </div> </div> </th><th class="header align-top pe-3 questionversionnumber" scope="col" data-pluginname="qbank_history__version_number_column" data-name="Version" data-columnid="qbank_history\version_number_column-version_number_column" style="width: 120px;"> <div class="header-container"> <div class="header-text"> </div> </div> <div class="sorters"> </div> </th><th class="header align-top creatorname" scope="col" data-pluginname="mod_studentquiz__question__bank__anonym_creator_name_column" data-name="Erstellt von" data-columnid="mod_studentquiz\question\bank\anonym_creator_name_column-anonym_creator_name_column" style="width: 120px;"> <div class="header-container"> <div class="header-text"> <div class="title me-1">Erstellt von</div> </div> </div> <div class="sorters"> / / </div> </th><th class="header align-top tags" scope="col" data-pluginname="mod_studentquiz__question__bank__tag_column" data-name="Tags" data-columnid="mod_studentquiz\question\bank\tag_column-tag_column" style="width: 120px;"> <div class="header-container"> <div class="header-text"> </div> </div> <div class="sorters"> </div> </th><th class="header align-top attempts" scope="col" data-pluginname="mod_studentquiz__question__bank__attempts_column" data-name="Meine Versuche" data-columnid="mod_studentquiz\question\bank\attempts_column-attempts_column" style="width: 120px;"> <div class="header-container"> <div class="header-text"> <div class="title me-1">Meine Versuche</div> </div> </div> <div class="sorters"> / </div> </th><th class="header align-top difficultylevel" scope="col" data-pluginname="mod_studentquiz__question__bank__difficulty_level_column" data-name="Schwierigkeit" data-columnid="mod_studentquiz\question\bank\\difficulty_level_column-difficulty_level_column" style="width: 120px;"> <div class="header-container"> <div class="header-text"> <div class="title me-1">Schwierigkeit</div> </div> </div> <div class="sorters"> / </div> </th><th class="header align-top rates" scope="col" data-pluginname="mod_studentquiz__question__bank__rate_column" data-name="Bewertung" data-columnid="mod_studentquiz\question\bank\rate_column-rate_column" style="width: 120px;"> <div class="header-container"> <div class="header-text"> <div class="title me-1">Bewertung</div> </div> </div> <div class="sorters"> / </div> </th><th class="header align-top comment" scope="col" data-pluginname="mod_studentquiz__question__bank__comment_column" data-name="Kommentare" data-columnid="mod_studentquiz\question\bank\comment_column-comment_column" style="width: 120px;"> <div class="header-container"> <div class="header-text"> </div> </div> <div class="sorters"> </div> </th></tr></thead><tbody><tr class="r0"><td class="checkbox" data-columnid="core_question\local\bank\checkbox_column-checkbox_column"><input id="checkq20458800" name="q20458800" type="checkbox" value="1" data-action="toggle" data-toggle="slave" data-togglegroup="qbank"> <label for="checkq20458800" class="accesshide">Auswahl</label></td><td class="qtype" data-columnid="qbank_viewquestiontype\question_type_column-question_type_column"><img class="icon " title="Kprim (ETH)" alt="Kprim (ETH)" src="https://moodle.zhaw.ch/theme/image.php/boost_union/qtype_kprime/1756457962/icon"></td><td class="state" data-columnid="mod_studentquiz\question\bank\state_column-state_column"></td><td class="state_pin" data-columnid="mod_studentquiz\question\bank\state_pin_column-state_pin_column"></td><td class="questionname" data-columnid="mod_studentquiz\question\bank\question_name_column-question_name_column"><label for="checkq20458800">Frage 01, einfach</label></td><td class="pe-3 editmenu" data-columnid="mod_studentquiz\question\bank\sq_edit_menu_column-sq_edit_menu_column"><div class="action-menu moodle-actionmenu" id="action-menu-1" data-enhance="moodle-core-actionmenu"> <div class="menubar d-flex " id="action-menu-1-menubar"> <div class="action-menu-trigger"> <div class="dropdown"> <div class="dropdown-menu menu dropdown-menu-right" id="action-menu-1-menu" data-rel="menu-content" aria-labelledby="action-menu-toggle-1" role="menu"> </div> </div> </div> </div> </div></td><td class="pe-3 questionversionnumber" data-columnid="qbank_history\version_number_column-version_number_column">v1</td><td class="creatorname" data-columnid="mod_studentquiz\question\bank\anonym_creator_name_column-anonym_creator_name_column"><span></span><br><span class="date">25. August 2025, 14:29</span></td><td class="tags" data-columnid="mod_studentquiz\question\bank\tag_column-tag_column">n.a.</td><td class="attempts" data-columnid="mod_studentquiz\question\bank\attempts_column-attempts_column"><span class="pratice_info" tabindex="0" aria-label="Diese Frage wurde noch nie versucht.">n.a.&nbsp;|&nbsp;n.a.</span></td><td class="difficultylevel" data-columnid="mod_studentquiz\question\bank\\difficulty_level_column-difficulty_level_column"><span class="mod_studentquiz_difficulty" data-mydifficulty="0" title=""></span></td><td class="rates" data-columnid="mod_studentquiz\question\bank\rate_column-rate_column"><span class="mod_studentquiz_ratingbar" title="">n.a.</span></td><td class="comment" data-columnid="mod_studentquiz\question\bank\comment_column-comment_column"><span class="public-comment badge badge-secondary" title="Anzahl an öffentlichen Kommentaren. Ein blauer Hintergrund bedeutet, dass Sie mindest einen ungelesenen Kommentar haben."> n.a. <span class="sr-only">Öffentlich Kommentare</span> </span></td></tr><tr class="r1"><td class="checkbox" data-columnid="core_question\local\bank\checkbox_column-checkbox_column"><input id="checkq20451333" name="q20451333" type="checkbox" value="1" data-action="toggle" data-toggle="slave" data-togglegroup="qbank"> <label for="checkq20451333" class="accesshide">Auswahl</label></td><td class="qtype" data-columnid="qbank_viewquestiontype\question_type_column-question_type_column"><img class="icon " title="Kprim (ETH)" alt="Kprim (ETH)" src="https://moodle.zhaw.ch/theme/image.php/boost_union/qtype_kprime/1756457962/icon"></td><td class="state" data-columnid="mod_studentquiz\question\bank\state_column-state_column"></td><td class="state_pin" data-columnid="mod_studentquiz\question\bank\state_pin_column-state_pin_column"></td><td class="questionname" data-columnid="mod_studentquiz\question\bank\question_name_column-question_name_column"><label for="checkq20451333">Gitarren im Musikbusiness</label></td><td class="pe-3 editmenu" data-columnid="mod_studentquiz\question\bank\sq_edit_menu_column-sq_edit_menu_column"><div class="action-menu moodle-actionmenu" id="action-menu-2" data-enhance="moodle-core-actionmenu"> <div class="menubar d-flex " id="action-menu-2-menubar"> <div class="action-menu-trigger"> <div class="dropdown"> <div class="dropdown-menu menu dropdown-menu-right" id="action-menu-2-menu" data-rel="menu-content" aria-labelledby="action-menu-toggle-2" role="menu"> </div> </div> </div> </div> </div></td><td class="pe-3 questionversionnumber" data-columnid="qbank_history\version_number_column-version_number_column">v1</td><td class="creatorname" data-columnid="mod_studentquiz\question\bank\anonym_creator_name_column-anonym_creator_name_column"><span></span><br><span class="date">21. August 2025, 10:42</span></td><td class="tags" data-columnid="mod_studentquiz\question\bank\tag_column-tag_column">n.a.</td><td class="attempts" data-columnid="mod_studentquiz\question\bank\attempts_column-attempts_column"><span class="pratice_info" tabindex="0" aria-label="Beim letzten Versuch falsch">1&nbsp;|&nbsp;✗</span></td><td class="difficultylevel" data-columnid="mod_studentquiz\question\bank\\difficulty_level_column-difficulty_level_column"><span class="mod_studentquiz_difficulty" data-difficultylevel="0.78" data-mydifficulty="1.00" title="Community Schwierigkeit: 78% , Meine Schwierigkeit: 100%"></span></td><td class="rates" data-columnid="mod_studentquiz\question\bank\rate_column-rate_column"><span class="mod_studentquiz_ratingbar" data-rate="4.33" data-myrate="4" title="Community Bewertung: 4.33 , Meine Bewertung: 4"></span></td><td class="comment" data-columnid="mod_studentquiz\question\bank\comment_column-comment_column"><span class="public-comment badge badge-secondary" title="Anzahl an öffentlichen Kommentaren. Ein blauer Hintergrund bedeutet, dass Sie mindest einen ungelesenen Kommentar haben."> 2 <span class="sr-only">Öffentlich Kommentare</span> </span></td></tr><tr class="r0"><td class="checkbox" data-columnid="core_question\local\bank\checkbox_column-checkbox_column"><input id="checkq20450640" name="q20450640" type="checkbox" value="1" data-action="toggle" data-toggle="slave" data-togglegroup="qbank"> <label for="checkq20450640" class="accesshide">Auswahl</label></td><td class="qtype" data-columnid="qbank_viewquestiontype\question_type_column-question_type_column"><img class="icon " title="Kprim (ETH)" alt="Kprim (ETH)" src="https://moodle.zhaw.ch/theme/image.php/boost_union/qtype_kprime/1756457962/icon"></td><td class="state" data-columnid="mod_studentquiz\question\bank\state_column-state_column"></td><td class="state_pin" data-columnid="mod_studentquiz\question\bank\state_pin_column-state_pin_column"></td><td class="questionname" data-columnid="mod_studentquiz\question\bank\question_name_column-question_name_column"><label for="checkq20450640">IKEA</label></td><td class="pe-3 editmenu" data-columnid="mod_studentquiz\question\bank\sq_edit_menu_column-sq_edit_menu_column"><div class="action-menu moodle-actionmenu" id="action-menu-3" data-enhance="moodle-core-actionmenu"> <div class="menubar d-flex " id="action-menu-3-menubar"> <div class="action-menu-trigger"> <div class="dropdown"> <div class="dropdown-menu menu dropdown-menu-right" id="action-menu-3-menu" data-rel="menu-content" aria-labelledby="action-menu-toggle-3" role="menu"> </div> </div> </div> </div> </div></td><td class="pe-3 questionversionnumber" data-columnid="qbank_history\version_number_column-version_number_column">v1</td><td class="creatorname" data-columnid="mod_studentquiz\question\bank\anonym_creator_name_column-anonym_creator_name_column"><span></span><br><span class="date">20. August 2025, 15:20</span></td><td class="tags" data-columnid="mod_studentquiz\question\bank\tag_column-tag_column">n.a.</td><td class="attempts" data-columnid="mod_studentquiz\question\bank\attempts_column-attempts_column"><span class="pratice_info" tabindex="0" aria-label="Diese Frage wurde noch nie versucht.">n.a.&nbsp;|&nbsp;n.a.</span></td><td class="difficultylevel" data-columnid="mod_studentquiz\question\bank\\difficulty_level_column-difficulty_level_column"><span class="mod_studentquiz_difficulty" data-difficultylevel="1.00" data-mydifficulty="0" title="Community Schwierigkeit: 100% , Meine Schwierigkeit: n.a."></span></td><td class="rates" data-columnid="mod_studentquiz\question\bank\rate_column-rate_column"><span class="mod_studentquiz_ratingbar" data-rate="2.50" title="Community Bewertung: 2.5 , Meine Bewertung: n.a."></span></td><td class="comment" data-columnid="mod_studentquiz\question\bank\comment_column-comment_column"><span class="public-comment badge badge-primary" title="Anzahl an öffentlichen Kommentaren. Ein blauer Hintergrund bedeutet, dass Sie mindest einen ungelesenen Kommentar haben."> 1 <span class="sr-only">Öffentlich Kommentar(inklusive ungelesener)</span> </span></td></tr><tr class="r1"><td class="checkbox" data-columnid="core_question\local\bank\checkbox_column-checkbox_column"><input id="checkq20450639" name="q20450639" type="checkbox" value="1" data-action="toggle" data-toggle="slave" data-togglegroup="qbank"> <label for="checkq20450639" class="accesshide">Auswahl</label></td><td class="qtype" data-columnid="qbank_viewquestiontype\question_type_column-question_type_column"><img class="icon " title="Kprim (ETH)" alt="Kprim (ETH)" src="https://moodle.zhaw.ch/theme/image.php/boost_union/qtype_kprime/1756457962/icon"></td><td class="state" data-columnid="mod_studentquiz\question\bank\state_column-state_column"></td><td class="state_pin" data-columnid="mod_studentquiz\question\bank\state_pin_column-state_pin_column"></td><td class="questionname" data-columnid="mod_studentquiz\question\bank\question_name_column-question_name_column"><label for="checkq20450639">Laubfrösche</label></td><td class="pe-3 editmenu" data-columnid="mod_studentquiz\question\bank\sq_edit_menu_column-sq_edit_menu_column"><div class="action-menu moodle-actionmenu" id="action-menu-4" data-enhance="moodle-core-actionmenu"> <div class="menubar d-flex " id="action-menu-4-menubar"> <div class="action-menu-trigger"> <div class="dropdown"> <div class="dropdown-menu menu dropdown-menu-right" id="action-menu-4-menu" data-rel="menu-content" aria-labelledby="action-menu-toggle-4" role="menu"> </div> </div> </div> </div> </div></td><td class="pe-3 questionversionnumber" data-columnid="qbank_history\version_number_column-version_number_column">v1</td><td class="creatorname" data-columnid="mod_studentquiz\question\bank\anonym_creator_name_column-anonym_creator_name_column"><span></span><br><span class="date">20. August 2025, 15:15</span></td><td class="tags" data-columnid="mod_studentquiz\question\bank\tag_column-tag_column">n.a.</td><td class="attempts" data-columnid="mod_studentquiz\question\bank\attempts_column-attempts_column"><span class="pratice_info" tabindex="0" aria-label="Diese Frage wurde noch nie versucht.">n.a.&nbsp;|&nbsp;n.a.</span></td><td class="difficultylevel" data-columnid="mod_studentquiz\question\bank\\difficulty_level_column-difficulty_level_column"><span class="mod_studentquiz_difficulty" data-difficultylevel="0.50" data-mydifficulty="0" title="Community Schwierigkeit: 50% , Meine Schwierigkeit: n.a."></span></td><td class="rates" data-columnid="mod_studentquiz\question\bank\rate_column-rate_column"><span class="mod_studentquiz_ratingbar" data-rate="3.50" title="Community Bewertung: 3.5 , Meine Bewertung: n.a."></span></td><td class="comment" data-columnid="mod_studentquiz\question\bank\comment_column-comment_column"><span class="public-comment badge badge-primary" title="Anzahl an öffentlichen Kommentaren. Ein blauer Hintergrund bedeutet, dass Sie mindest einen ungelesenen Kommentar haben."> 1 <span class="sr-only">Öffentlich Kommentar(inklusive ungelesener)</span> </span></td></tr><tr class="r0"><td class="checkbox" data-columnid="core_question\local\bank\checkbox_column-checkbox_column"><input id="checkq20450638" name="q20450638" type="checkbox" value="1" data-action="toggle" data-toggle="slave" data-togglegroup="qbank"> <label for="checkq20450638" class="accesshide">Auswahl</label></td><td class="qtype" data-columnid="qbank_viewquestiontype\question_type_column-question_type_column"><img class="icon " title="Kprim (ETH)" alt="Kprim (ETH)" src="https://moodle.zhaw.ch/theme/image.php/boost_union/qtype_kprime/1756457962/icon"></td><td class="state" data-columnid="mod_studentquiz\question\bank\state_column-state_column"></td><td class="state_pin" data-columnid="mod_studentquiz\question\bank\state_pin_column-state_pin_column"></td><td class="questionname" data-columnid="mod_studentquiz\question\bank\question_name_column-question_name_column"><label for="checkq20450638">Wertschöpfungsanalyse</label></td><td class="pe-3 editmenu" data-columnid="mod_studentquiz\question\bank\sq_edit_menu_column-sq_edit_menu_column"><div class="action-menu moodle-actionmenu" id="action-menu-5" data-enhance="moodle-core-actionmenu"> <div class="menubar d-flex " id="action-menu-5-menubar"> <div class="action-menu-trigger"> <div class="dropdown"> <div class="dropdown-menu menu dropdown-menu-right" id="action-menu-5-menu" data-rel="menu-content" aria-labelledby="action-menu-toggle-5" role="menu"> </div> </div> </div> </div> </div></td><td class="pe-3 questionversionnumber" data-columnid="qbank_history\version_number_column-version_number_column">v1</td><td class="creatorname" data-columnid="mod_studentquiz\question\bank\anonym_creator_name_column-anonym_creator_name_column"><span></span><br><span class="date">20. August 2025, 15:10</span></td><td class="tags" data-columnid="mod_studentquiz\question\bank\tag_column-tag_column">n.a.</td><td class="attempts" data-columnid="mod_studentquiz\question\bank\attempts_column-attempts_column"><span class="pratice_info" tabindex="0" aria-label="Diese Frage wurde noch nie versucht.">n.a.&nbsp;|&nbsp;n.a.</span></td><td class="difficultylevel" data-columnid="mod_studentquiz\question\bank\\difficulty_level_column-difficulty_level_column"><span class="mod_studentquiz_difficulty" data-mydifficulty="0" title=""></span></td><td class="rates" data-columnid="mod_studentquiz\question\bank\rate_column-rate_column"><span class="mod_studentquiz_ratingbar" title="">n.a.</span></td><td class="comment" data-columnid="mod_studentquiz\question\bank\comment_column-comment_column"><span class="public-comment badge badge-secondary" title="Anzahl an öffentlichen Kommentaren. Ein blauer Hintergrund bedeutet, dass Sie mindest einen ungelesenen Kommentar haben."> n.a. <span class="sr-only">Öffentlich Kommentare</span> </span></td></tr></tbody></table>`;
  }

  // Event bindings
  if (els.btnSample) els.btnSample.addEventListener('click', insertSample);
  if (els.btnParse) els.btnParse.addEventListener('click', parseHTML);
  if (els.btnClear) els.btnClear.addEventListener('click', clearAll);
  if (els.btnDownload) els.btnDownload.addEventListener('click', downloadExcel);

  // Ranking events
  if (els.btnSampleRanking) els.btnSampleRanking.addEventListener('click', insertSampleRanking);
  if (els.btnParseRanking) els.btnParseRanking.addEventListener('click', parseRankingHTML);
  if (els.btnClearRanking) els.btnClearRanking.addEventListener('click', clearRanking);
  if (els.btnDownloadRanking) els.btnDownloadRanking.addEventListener('click', downloadRankingExcel);

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
  }
})();
