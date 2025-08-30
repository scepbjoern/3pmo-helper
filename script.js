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
    tableBody: $('#previewTable tbody')
  };

  let extractedRows = [];

  function setStatus(msg) {
    els.status.textContent = msg || '';
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
      XLSX.writeFile(wb, 'studentquiz_extrakt.xlsx');
      setStatus(`Excel exportiert (${data.length} Zeilen).`);
    } else {
      // CSV Fallback
      const csv = toCSV(data);
      const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = 'studentquiz_extrakt.csv';
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

  function getSampleHTML() {
    return `
<table id="categoryquestions" class="question-bank-table generaltable" data-defaultsort="{&quot;mod_studentquiz__question__bank__anonym_creator_name_column-timecreated&quot;:3,&quot;mod_studentquiz__question__bank__question_name_column&quot;:4}"><thead id="yui_3_18_1_1_1756465703414_173"><tr class="qbank-column-list" id="yui_3_18_1_1_1756465703414_172"><th class="header align-top checkbox" scope="col" data-pluginname="core_question\local\bank\checkbox_column" data-name="Alle auswählen" data-columnid="core_question\local\bank\checkbox_column-checkbox_column" style="width: 30px;" id="yui_3_18_1_1_1756465703414_171"> <div class="header-container" id="yui_3_18_1_1_1756465703414_170"> <div class="header-text" id="yui_3_18_1_1_1756465703414_169"> <span class="me-1" title="Fragen für Sammelaktion auswählen" id="yui_3_18_1_1_1756465703414_168"><div class="form-check" id="yui_3_18_1_1_1756465703414_167"> <input id="qbheadercheckbox" name="qbheadercheckbox" type="checkbox" class="" value="1" aria-labelledby="qbheadercheckbox-label" data-action="toggle" data-toggle="master" data-togglegroup="qbank" data-toggle-selectall="Alle auswählen" data-toggle-deselectall="Nichts auswählen"> <label id="qbheadercheckbox-label" for="qbheadercheckbox" class="form-check-label d-block pe-2 accesshide">Nichts auswählen</label> </div></span> </div> </div> </th><th class="header align-top qtype" scope="col" data-pluginname="qbank_viewquestiontype__question_type_column" data-name="T" data-columnid="qbank_viewquestiontype\question_type_column-question_type_column" style="width: 45px;"> <div class="header-container"> <div class="header-text"> </div> </div> <div class="sorters"> <a href="https://moodle.zhaw.ch/mod/studentquiz/view.php?cmid=1723409&amp;cat=519228%2C2302532&amp;id=1723409&amp;group=163617&amp;sortdata%5Bqbank_viewquestiontype__question_type_column%5D=4&amp;sortdata%5Bmod_studentquiz__question__bank__anonym_creator_name_column-timecreated%5D=3&amp;sortdata%5Bmod_studentquiz__question__bank__question_name_column%5D=4" data-sortname="qbank_viewquestiontype__question_type_column" data-sortorder="4" title="Sortierung nach Fragetyp aufsteigend"> T </a> </div> </th><th class="header align-top state" scope="col" data-pluginname="mod_studentquiz__question__bank__state_column" data-name="S" data-columnid="mod_studentquiz\question\bank\state_column-state_column" style="width: 120px;"> <div class="header-container"> <div class="header-text"> </div> </div> <div class="sorters"> <a href="https://moodle.zhaw.ch/mod/studentquiz/view.php?cmid=1723409&amp;cat=519228%2C2302532&amp;id=1723409&amp;group=163617&amp;sortdata%5Bmod_studentquiz__question__bank__state_column%5D=4&amp;sortdata%5Bmod_studentquiz__question__bank__anonym_creator_name_column-timecreated%5D=3&amp;sortdata%5Bmod_studentquiz__question__bank__question_name_column%5D=4" data-sortname="mod_studentquiz__question__bank__state_column" data-sortorder="4" title="Sortierung nach Status aufsteigend"> S </a> </div> </th><th class="header align-top state_pin" scope="col" data-pluginname="mod_studentquiz__question__bank__state_pin_column" data-name="" data-columnid="mod_studentquiz\question\bank\state_pin_column-state_pin_column" style="width: 120px;"> <div class="header-container"> <div class="header-text"> <span class="me-1"></span> </div> </div> </th><th class="header align-top questionname" scope="col" data-pluginname="mod_studentquiz__question__bank__question_name_column" data-name="Frage" data-columnid="mod_studentquiz\question\bank\question_name_column-question_name_column" style="width: 250px;"> <div class="header-container"> <div class="header-text"> </div> </div> <div class="sorters"> </div> </th><th class="header align-top pe-3 editmenu" scope="col" data-pluginname="mod_studentquiz__question__bank__sq_edit_menu_column" data-name="Aktionen" data-columnid="mod_studentquiz\question\bank\sq_edit_menu_column-sq_edit_menu_column" style="width: 120px;"> <div class="header-container"> <div class="header-text"> <span class="me-1">Aktionen</span> </div> </div> </th><th class="header align-top pe-3 questionversionnumber" scope="col" data-pluginname="qbank_history__version_number_column" data-name="Version" data-columnid="qbank_history\version_number_column-version_number_column" style="width: 120px;"> <div class="header-container"> <div class="header-text"> </div> </div> <div class="sorters"> </div> </th><th class="header align-top creatorname" scope="col" data-pluginname="mod_studentquiz__question__bank__anonym_creator_name_column" data-name="Erstellt von" data-columnid="mod_studentquiz\question\bank\anonym_creator_name_column-anonym_creator_name_column" style="width: 120px;"> <div class="header-container"> <div class="header-text"> <div class="title me-1">Erstellt von</div> </div> </div> <div class="sorters"> / / </div> </th><th class="header align-top tags" scope="col" data-pluginname="mod_studentquiz__question__bank__tag_column" data-name="Tags" data-columnid="mod_studentquiz\question\bank\tag_column-tag_column" style="width: 120px;"> <div class="header-container"> <div class="header-text"> </div> </div> <div class="sorters"> </div> </th><th class="header align-top attempts" scope="col" data-pluginname="mod_studentquiz__question__bank__attempts_column" data-name="Meine Versuche" data-columnid="mod_studentquiz\question\bank\attempts_column-attempts_column" style="width: 120px;"> <div class="header-container"> <div class="header-text"> <div class="title me-1">Meine Versuche</div> </div> </div> <div class="sorters"> / </div> </th><th class="header align-top difficultylevel" scope="col" data-pluginname="mod_studentquiz__question__bank__difficulty_level_column" data-name="Schwierigkeit" data-columnid="mod_studentquiz\question\bank\\difficulty_level_column-difficulty_level_column" style="width: 120px;"> <div class="header-container"> <div class="header-text"> <div class="title me-1">Schwierigkeit</div> </div> </div> <div class="sorters"> / </div> </th><th class="header align-top rates" scope="col" data-pluginname="mod_studentquiz__question__bank__rate_column" data-name="Bewertung" data-columnid="mod_studentquiz\question\bank\rate_column-rate_column" style="width: 120px;"> <div class="header-container"> <div class="header-text"> <div class="title me-1">Bewertung</div> </div> </div> <div class="sorters"> / </div> </th><th class="header align-top comment" scope="col" data-pluginname="mod_studentquiz__question__bank__comment_column" data-name="Kommentare" data-columnid="mod_studentquiz\question\bank\comment_column-comment_column" style="width: 120px;"> <div class="header-container"> <div class="header-text"> </div> </div> <div class="sorters"> </div> </th></tr></thead><tbody><tr class="r0"><td class="checkbox" data-columnid="core_question\local\bank\checkbox_column-checkbox_column"><input id="checkq20458800" name="q20458800" type="checkbox" value="1" data-action="toggle" data-toggle="slave" data-togglegroup="qbank"> <label for="checkq20458800" class="accesshide">Auswahl</label></td><td class="qtype" data-columnid="qbank_viewquestiontype\question_type_column-question_type_column"><img class="icon " title="Kprim (ETH)" alt="Kprim (ETH)" src="https://moodle.zhaw.ch/theme/image.php/boost_union/qtype_kprime/1756457962/icon"></td><td class="state" data-columnid="mod_studentquiz\question\bank\state_column-state_column"></td><td class="state_pin" data-columnid="mod_studentquiz\question\bank\state_pin_column-state_pin_column"></td><td class="questionname" data-columnid="mod_studentquiz\question\bank\question_name_column-question_name_column"><label for="checkq20458800">Frage 01, einfach</label></td><td class="pe-3 editmenu" data-columnid="mod_studentquiz\question\bank\sq_edit_menu_column-sq_edit_menu_column"><div class="action-menu moodle-actionmenu" id="action-menu-1" data-enhance="moodle-core-actionmenu"> <div class="menubar d-flex " id="action-menu-1-menubar"> <div class="action-menu-trigger"> <div class="dropdown"> <div class="dropdown-menu menu dropdown-menu-right" id="action-menu-1-menu" data-rel="menu-content" aria-labelledby="action-menu-toggle-1" role="menu"> </div> </div> </div> </div> </div></td><td class="pe-3 questionversionnumber" data-columnid="qbank_history\version_number_column-version_number_column">v1</td><td class="creatorname" data-columnid="mod_studentquiz\question\bank\anonym_creator_name_column-anonym_creator_name_column"><span></span><br><span class="date">25. August 2025, 14:29</span></td><td class="tags" data-columnid="mod_studentquiz\question\bank\tag_column-tag_column">n.a.</td><td class="attempts" data-columnid="mod_studentquiz\question\bank\attempts_column-attempts_column"><span class="pratice_info" tabindex="0" aria-label="Diese Frage wurde noch nie versucht.">n.a.&nbsp;|&nbsp;n.a.</span></td><td class="difficultylevel" data-columnid="mod_studentquiz\question\bank\\difficulty_level_column-difficulty_level_column"><span class="mod_studentquiz_difficulty" data-mydifficulty="0" title=""></span></td><td class="rates" data-columnid="mod_studentquiz\question\bank\rate_column-rate_column"><span class="mod_studentquiz_ratingbar" title="">n.a.</span></td><td class="comment" data-columnid="mod_studentquiz\question\bank\comment_column-comment_column"><span class="public-comment badge badge-secondary" title="Anzahl an öffentlichen Kommentaren. Ein blauer Hintergrund bedeutet, dass Sie mindest einen ungelesenen Kommentar haben."> n.a. <span class="sr-only">Öffentlich Kommentare</span> </span></td></tr><tr class="r1"><td class="checkbox" data-columnid="core_question\local\bank\checkbox_column-checkbox_column"><input id="checkq20451333" name="q20451333" type="checkbox" value="1" data-action="toggle" data-toggle="slave" data-togglegroup="qbank"> <label for="checkq20451333" class="accesshide">Auswahl</label></td><td class="qtype" data-columnid="qbank_viewquestiontype\question_type_column-question_type_column"><img class="icon " title="Kprim (ETH)" alt="Kprim (ETH)" src="https://moodle.zhaw.ch/theme/image.php/boost_union/qtype_kprime/1756457962/icon"></td><td class="state" data-columnid="mod_studentquiz\question\bank\state_column-state_column"></td><td class="state_pin" data-columnid="mod_studentquiz\question\bank\state_pin_column-state_pin_column"></td><td class="questionname" data-columnid="mod_studentquiz\question\bank\question_name_column-question_name_column"><label for="checkq20451333">Gitarren im Musikbusiness</label></td><td class="pe-3 editmenu" data-columnid="mod_studentquiz\question\bank\sq_edit_menu_column-sq_edit_menu_column"><div class="action-menu moodle-actionmenu" id="action-menu-2" data-enhance="moodle-core-actionmenu"> <div class="menubar d-flex " id="action-menu-2-menubar"> <div class="action-menu-trigger"> <div class="dropdown"> <div class="dropdown-menu menu dropdown-menu-right" id="action-menu-2-menu" data-rel="menu-content" aria-labelledby="action-menu-toggle-2" role="menu"> </div> </div> </div> </div> </div></td><td class="pe-3 questionversionnumber" data-columnid="qbank_history\version_number_column-version_number_column">v1</td><td class="creatorname" data-columnid="mod_studentquiz\question\bank\anonym_creator_name_column-anonym_creator_name_column"><span></span><br><span class="date">21. August 2025, 10:42</span></td><td class="tags" data-columnid="mod_studentquiz\question\bank\tag_column-tag_column">n.a.</td><td class="attempts" data-columnid="mod_studentquiz\question\bank\attempts_column-attempts_column"><span class="pratice_info" tabindex="0" aria-label="Beim letzten Versuch falsch">1&nbsp;|&nbsp;✗</span></td><td class="difficultylevel" data-columnid="mod_studentquiz\question\bank\\difficulty_level_column-difficulty_level_column"><span class="mod_studentquiz_difficulty" data-difficultylevel="0.78" data-mydifficulty="1.00" title="Community Schwierigkeit: 78% , Meine Schwierigkeit: 100%"></span></td><td class="rates" data-columnid="mod_studentquiz\question\bank\rate_column-rate_column"><span class="mod_studentquiz_ratingbar" data-rate="4.33" data-myrate="4" title="Community Bewertung: 4.33 , Meine Bewertung: 4"></span></td><td class="comment" data-columnid="mod_studentquiz\question\bank\comment_column-comment_column"><span class="public-comment badge badge-secondary" title="Anzahl an öffentlichen Kommentaren. Ein blauer Hintergrund bedeutet, dass Sie mindest einen ungelesenen Kommentar haben."> 2 <span class="sr-only">Öffentlich Kommentare</span> </span></td></tr><tr class="r0"><td class="checkbox" data-columnid="core_question\local\bank\checkbox_column-checkbox_column"><input id="checkq20450640" name="q20450640" type="checkbox" value="1" data-action="toggle" data-toggle="slave" data-togglegroup="qbank"> <label for="checkq20450640" class="accesshide">Auswahl</label></td><td class="qtype" data-columnid="qbank_viewquestiontype\question_type_column-question_type_column"><img class="icon " title="Kprim (ETH)" alt="Kprim (ETH)" src="https://moodle.zhaw.ch/theme/image.php/boost_union/qtype_kprime/1756457962/icon"></td><td class="state" data-columnid="mod_studentquiz\question\bank\state_column-state_column"></td><td class="state_pin" data-columnid="mod_studentquiz\question\bank\state_pin_column-state_pin_column"></td><td class="questionname" data-columnid="mod_studentquiz\question\bank\question_name_column-question_name_column"><label for="checkq20450640">IKEA</label></td><td class="pe-3 editmenu" data-columnid="mod_studentquiz\question\bank\sq_edit_menu_column-sq_edit_menu_column"><div class="action-menu moodle-actionmenu" id="action-menu-3" data-enhance="moodle-core-actionmenu"> <div class="menubar d-flex " id="action-menu-3-menubar"> <div class="action-menu-trigger"> <div class="dropdown"> <div class="dropdown-menu menu dropdown-menu-right" id="action-menu-3-menu" data-rel="menu-content" aria-labelledby="action-menu-toggle-3" role="menu"> </div> </div> </div> </div> </div></td><td class="pe-3 questionversionnumber" data-columnid="qbank_history\version_number_column-version_number_column">v1</td><td class="creatorname" data-columnid="mod_studentquiz\question\bank\anonym_creator_name_column-anonym_creator_name_column"><span></span><br><span class="date">20. August 2025, 15:20</span></td><td class="tags" data-columnid="mod_studentquiz\question\bank\tag_column-tag_column">n.a.</td><td class="attempts" data-columnid="mod_studentquiz\question\bank\attempts_column-attempts_column"><span class="pratice_info" tabindex="0" aria-label="Diese Frage wurde noch nie versucht.">n.a.&nbsp;|&nbsp;n.a.</span></td><td class="difficultylevel" data-columnid="mod_studentquiz\question\bank\\difficulty_level_column-difficulty_level_column"><span class="mod_studentquiz_difficulty" data-difficultylevel="1.00" data-mydifficulty="0" title="Community Schwierigkeit: 100% , Meine Schwierigkeit: n.a."></span></td><td class="rates" data-columnid="mod_studentquiz\question\bank\rate_column-rate_column"><span class="mod_studentquiz_ratingbar" data-rate="2.50" title="Community Bewertung: 2.5 , Meine Bewertung: n.a."></span></td><td class="comment" data-columnid="mod_studentquiz\question\bank\comment_column-comment_column"><span class="public-comment badge badge-primary" title="Anzahl an öffentlichen Kommentaren. Ein blauer Hintergrund bedeutet, dass Sie mindest einen ungelesenen Kommentar haben."> 1 <span class="sr-only">Öffentlich Kommentar(inklusive ungelesener)</span> </span></td></tr><tr class="r1"><td class="checkbox" data-columnid="core_question\local\bank\checkbox_column-checkbox_column"><input id="checkq20450639" name="q20450639" type="checkbox" value="1" data-action="toggle" data-toggle="slave" data-togglegroup="qbank"> <label for="checkq20450639" class="accesshide">Auswahl</label></td><td class="qtype" data-columnid="qbank_viewquestiontype\question_type_column-question_type_column"><img class="icon " title="Kprim (ETH)" alt="Kprim (ETH)" src="https://moodle.zhaw.ch/theme/image.php/boost_union/qtype_kprime/1756457962/icon"></td><td class="state" data-columnid="mod_studentquiz\question\bank\state_column-state_column"></td><td class="state_pin" data-columnid="mod_studentquiz\question\bank\state_pin_column-state_pin_column"></td><td class="questionname" data-columnid="mod_studentquiz\question\bank\question_name_column-question_name_column"><label for="checkq20450639">Laubfrösche</label></td><td class="pe-3 editmenu" data-columnid="mod_studentquiz\question\bank\sq_edit_menu_column-sq_edit_menu_column"><div class="action-menu moodle-actionmenu" id="action-menu-4" data-enhance="moodle-core-actionmenu"> <div class="menubar d-flex " id="action-menu-4-menubar"> <div class="action-menu-trigger"> <div class="dropdown"> <div class="dropdown-menu menu dropdown-menu-right" id="action-menu-4-menu" data-rel="menu-content" aria-labelledby="action-menu-toggle-4" role="menu"> </div> </div> </div> </div> </div></td><td class="pe-3 questionversionnumber" data-columnid="qbank_history\version_number_column-version_number_column">v1</td><td class="creatorname" data-columnid="mod_studentquiz\question\bank\anonym_creator_name_column-anonym_creator_name_column"><span></span><br><span class="date">20. August 2025, 15:15</span></td><td class="tags" data-columnid="mod_studentquiz\question\bank\tag_column-tag_column">n.a.</td><td class="attempts" data-columnid="mod_studentquiz\question\bank\attempts_column-attempts_column"><span class="pratice_info" tabindex="0" aria-label="Diese Frage wurde noch nie versucht.">n.a.&nbsp;|&nbsp;n.a.</span></td><td class="difficultylevel" data-columnid="mod_studentquiz\question\bank\\difficulty_level_column-difficulty_level_column"><span class="mod_studentquiz_difficulty" data-difficultylevel="0.50" data-mydifficulty="0" title="Community Schwierigkeit: 50% , Meine Schwierigkeit: n.a."></span></td><td class="rates" data-columnid="mod_studentquiz\question\bank\rate_column-rate_column"><span class="mod_studentquiz_ratingbar" data-rate="3.50" title="Community Bewertung: 3.5 , Meine Bewertung: n.a."></span></td><td class="comment" data-columnid="mod_studentquiz\question\bank\comment_column-comment_column"><span class="public-comment badge badge-primary" title="Anzahl an öffentlichen Kommentaren. Ein blauer Hintergrund bedeutet, dass Sie mindest einen ungelesenen Kommentar haben."> 1 <span class="sr-only">Öffentlich Kommentar(inklusive ungelesener)</span> </span></td></tr><tr class="r0"><td class="checkbox" data-columnid="core_question\local\bank\checkbox_column-checkbox_column"><input id="checkq20450638" name="q20450638" type="checkbox" value="1" data-action="toggle" data-toggle="slave" data-togglegroup="qbank"> <label for="checkq20450638" class="accesshide">Auswahl</label></td><td class="qtype" data-columnid="qbank_viewquestiontype\question_type_column-question_type_column"><img class="icon " title="Kprim (ETH)" alt="Kprim (ETH)" src="https://moodle.zhaw.ch/theme/image.php/boost_union/qtype_kprime/1756457962/icon"></td><td class="state" data-columnid="mod_studentquiz\question\bank\state_column-state_column"></td><td class="state_pin" data-columnid="mod_studentquiz\question\bank\state_pin_column-state_pin_column"></td><td class="questionname" data-columnid="mod_studentquiz\question\bank\question_name_column-question_name_column"><label for="checkq20450638">Wertschöpfungsanalyse</label></td><td class="pe-3 editmenu" data-columnid="mod_studentquiz\question\bank\sq_edit_menu_column-sq_edit_menu_column"><div class="action-menu moodle-actionmenu" id="action-menu-5" data-enhance="moodle-core-actionmenu"> <div class="menubar d-flex " id="action-menu-5-menubar"> <div class="action-menu-trigger"> <div class="dropdown"> <div class="dropdown-menu menu dropdown-menu-right" id="action-menu-5-menu" data-rel="menu-content" aria-labelledby="action-menu-toggle-5" role="menu"> </div> </div> </div> </div> </div></td><td class="pe-3 questionversionnumber" data-columnid="qbank_history\version_number_column-version_number_column">v1</td><td class="creatorname" data-columnid="mod_studentquiz\question\bank\anonym_creator_name_column-anonym_creator_name_column"><span></span><br><span class="date">20. August 2025, 15:10</span></td><td class="tags" data-columnid="mod_studentquiz\question\bank\tag_column-tag_column">n.a.</td><td class="attempts" data-columnid="mod_studentquiz\question\bank\attempts_column-attempts_column"><span class="pratice_info" tabindex="0" aria-label="Diese Frage wurde noch nie versucht.">n.a.&nbsp;|&nbsp;n.a.</span></td><td class="difficultylevel" data-columnid="mod_studentquiz\question\bank\\difficulty_level_column-difficulty_level_column"><span class="mod_studentquiz_difficulty" data-mydifficulty="0" title=""></span></td><td class="rates" data-columnid="mod_studentquiz\question\bank\rate_column-rate_column"><span class="mod_studentquiz_ratingbar" title="">n.a.</span></td><td class="comment" data-columnid="mod_studentquiz\question\bank\comment_column-comment_column"><span class="public-comment badge badge-secondary" title="Anzahl an öffentlichen Kommentaren. Ein blauer Hintergrund bedeutet, dass Sie mindest einen ungelesenen Kommentar haben."> n.a. <span class="sr-only">Öffentlich Kommentare</span> </span></td></tr></tbody></table>`;
  }

  // Event bindings
  els.btnSample.addEventListener('click', insertSample);
  els.btnParse.addEventListener('click', parseHTML);
  els.btnClear.addEventListener('click', clearAll);
  els.btnDownload.addEventListener('click', downloadExcel);
})();
