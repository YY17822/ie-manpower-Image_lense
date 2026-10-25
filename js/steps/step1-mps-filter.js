/*
 * 第 1 步：篩選 MPS 報表
 *   1. 上傳 MPS 的 Excel
 *   2. 選取條件，兩種方式二選一：
 *        a. A 行等於多少（原本的方式；A 行可在「報表版面不一樣時」改成別的行）
 *        b. 指定某一行，等於或不等於某些值
 *           「不等於」時可勾「排除這一行空白的列」（預設勾選），把分類之間的「x」「TTL (…)」
 *           小計列拿掉；被排除的列會逐列列出，取消勾選的那列會保留（不記住，因為每個月列號不同）
 *      符合的列才留下（從第 5 列開始比對）
 *   3. 輸入要保留哪些行
 *   4. 匯出新的 Excel，第 1–4 列表頭（含合併儲存格）原樣保留
 *
 * 做完之後把結果放在 IEApp.state.mpsFiltered，第 2 步直接拿來用。
 *
 * 名詞：行 = 直的（A、B、C…），列 = 橫的（1、2、3…）。
 */
(function (global) {
  'use strict';

  var IEApp = global.IEApp;
  var XIO = global.ExcelIO;
  var ui = IEApp.ui;
  var h = IEApp.h;

  var STORE_KEY = 'step1';
  var PREVIEW_ROWS = 8;     // 預覽顯示幾列資料
  var MAX_CHIPS = 24;

  // mode：'a' = A 行等於（原本的方式）、'col' = 指定某一行等於／不等於
  // fCol、fOp（'eq' 等於、'ne' 不等於）、fValues：指定某一行的條件
  // fSkipBlank：「不等於」時排除那一行空白的列
  var DEFAULTS = { values: '', cols: '', startRow: 5, selCol: 'A', mode: 'a', fCol: '', fOp: 'eq', fValues: '', fSkipBlank: true };
  var MAX_BLANK_LIST = 200;  // 空白列清單最多列幾列

  // 這一步的狀態（切到別的步驟再回來時還在）
  // keepBlank：空白列清單裡使用者取消勾選（要保留）的列號 { 52: true }；換檔案、換行、換起始列時清空
  var S = { fileName: '', workbook: null, ws: null, info: null, infoKey: '', result: null, keepBlank: {}, blankKey: '' };
  var settings = loadSettings();
  var els = {};
  var timer = null;

  /* ---------- 記住上次輸入的條件（每個月做的事情差不多） ---------- */
  function loadSettings() {
    var s = {};
    try { s = IEApp.settings.read(STORE_KEY) || {}; } catch (e) { s = {}; }
    var out = {};
    Object.keys(DEFAULTS).forEach(function (k) { out[k] = s[k] === undefined ? DEFAULTS[k] : s[k]; });
    return out;
  }
  function saveSettings() {
    try { IEApp.settings.write(STORE_KEY, settings); } catch (e) { /* 存不了就算了 */ }
  }

  /* ---------- 讀取目前的設定值 ---------- */
  function startRow() {
    var n = parseInt(XIO.normalizeText(settings.startRow), 10);
    return n >= 1 ? n : DEFAULTS.startRow;
  }
  function selCol() {
    var p = XIO.parseColumns(settings.selCol);
    return p.cols.length === 1 && !p.bad.length ? p.cols[0] : 1;
  }
  function byColumn() { return settings.mode === 'col'; }
  function exclude() { return byColumn() && settings.fOp === 'ne'; }
  function skipBlank() { return exclude() && settings.fSkipBlank !== false; }
  // 指定某一行的方式：行號看不懂或沒填時回傳 0
  function fCol() {
    var p = XIO.parseColumns(settings.fCol);
    return p.cols.length === 1 && !p.bad.length ? p.cols[0] : 0;
  }
  // 目前用來篩選的行
  function activeCol() { return byColumn() ? fCol() : selCol(); }
  function valuesKey() { return byColumn() ? 'fValues' : 'values'; }
  function wantedValues() {
    return XIO.normalizeText(settings[valuesKey()]).split(/[,、;]+/)
      .map(function (v) { return v.trim(); })
      .filter(function (v) { return v !== ''; });
  }
  function outputName() {
    return S.fileName.replace(/\.[^.]+$/, '') + '_篩選.xlsx';
  }

  /* ---------- 畫面 ---------- */
  function render(panel) {
    els = {};

    els.drop = ui.fileDrop({
      id: 'f-file',
      title: '選擇 MPS 的 Excel 檔',
      sub: '或把檔案拖進這裡。支援 .xlsx',
      onFile: openFile
    });
    els.sheetSelect = h('select', { id: 'f-sheet', 'aria-label': '工作表' });
    els.sheetRow = h('div', { class: 'field-row', hidden: true },
      h('label', { for: 'f-sheet', text: '工作表' }), els.sheetSelect);
    els.fileMsg = h('p', { class: 'msg is-error', role: 'alert', hidden: true });

    els.selLabel = h('span', { text: 'A' });
    els.values = h('input', {
      id: 'f-values', type: 'text', class: 'in in-short', autocomplete: 'off',
      inputmode: 'decimal', placeholder: '例如 1', value: settings.values
    });
    els.chips = h('div', { class: 'chips' });

    // 選取條件的方式（二選一）
    els.modeA = h('input', { type: 'radio', name: 'f-mode', id: 'f-mode-a', value: 'a' });
    els.modeCol = h('input', { type: 'radio', name: 'f-mode', id: 'f-mode-col', value: 'col' });
    els.fCol = h('input', {
      id: 'f-fcol', type: 'text', class: 'in in-tiny', autocomplete: 'off', spellcheck: 'false',
      placeholder: '例如 F', value: settings.fCol, 'aria-label': '要篩選的行'
    });
    els.fOp = h('select', { id: 'f-fop', 'aria-label': '等於或不等於' },
      h('option', { value: 'eq', text: '等於' }), h('option', { value: 'ne', text: '不等於' }));
    els.fOp.value = settings.fOp === 'ne' ? 'ne' : 'eq';
    els.fValues = h('input', {
      id: 'f-fvalues', type: 'text', class: 'in in-short', autocomplete: 'off',
      placeholder: '例如 EOL, 停產', value: settings.fValues, 'aria-label': '值'
    });
    els.rowA = h('div', { class: 'field-row mode-body' },
      h('label', { for: 'f-values' }, els.selLabel, ' 行等於'), els.values);
    els.rowCol = h('div', { class: 'field-row mode-body' },
      els.fCol, h('span', { text: '行' }), els.fOp, els.fValues);
    els.skipBlank = h('input', { type: 'checkbox', id: 'f-skipblank' });
    els.skipBlank.checked = settings.fSkipBlank !== false;
    els.skipRow = h('div', { class: 'field-row mode-body skip-blank' },
      h('label', { class: 'check-label', for: 'f-skipblank' }, els.skipBlank, h('span', { text: '排除這一行空白的列' })));
    els.blankList = h('div', { class: 'blank-list mode-body' });

    els.cols = h('input', {
      id: 'f-cols', type: 'text', class: 'in in-wide', autocomplete: 'off', spellcheck: 'false',
      placeholder: '例如 B, C, U, Z, AE', value: settings.cols
    });
    els.srcPreview = h('div', { class: 'sheet-area' });

    els.startRow = h('input', {
      id: 'f-start', type: 'text', class: 'in in-tiny', inputmode: 'numeric', value: String(settings.startRow)
    });
    els.selCol = h('input', {
      id: 'f-selcol', type: 'text', class: 'in in-tiny', autocomplete: 'off', value: String(settings.selCol)
    });

    els.summary = h('p', { class: 'summary', 'aria-live': 'polite' });
    els.outPreview = h('div', { class: 'sheet-area' });
    els.exportBtn = h('button', { type: 'button', class: 'btn-primary', disabled: true, text: '匯出 Excel', onclick: onExport });
    els.nextBtn = h('button', {
      type: 'button', class: 'btn-secondary', disabled: true, text: '下一步：匯總',
      onclick: function () { IEApp.go('mps-totals'); }
    });
    els.exportNote = h('span', { class: 'export-note', 'aria-live': 'polite' });

    panel.appendChild(h('div', { class: 'step-head' },
      h('h1', { text: '篩選 MPS 報表' }),
      h('p', { text: '依選取條件留下要的列（A 行等於指定數值，或指定某一行等於／不等於某些值），以及你要的行，匯出成新的 Excel。表頭（第 1–4 列）會原樣保留。' })));

    panel.appendChild(h('ol', { class: 'tasks' },
      h('li', { class: 'task' },
        h('h2', { text: '上傳 MPS 報表' }),
        els.drop.el, els.sheetRow, els.fileMsg),
      h('li', { class: 'task' },
        h('h2', { text: '選取條件' }),
        h('fieldset', { class: 'mode-group' },
          h('legend', { class: 'visually-hidden', text: '選取條件的方式' }),
          h('div', { class: 'mode-option' },
            h('label', { class: 'mode-label', for: 'f-mode-a' }, els.modeA, h('span', { text: '依 A 行（原本的方式）' })),
            els.rowA),
          h('div', { class: 'mode-option' },
            h('label', { class: 'mode-label', for: 'f-mode-col' }, els.modeCol, h('span', { text: '指定某一行，等於或不等於某些值' })),
            els.rowCol, els.skipRow, els.blankList)),
        h('p', { class: 'hint', text: '兩種方式二選一。要選多個值時用逗號分開，例如 1, 3。選「不等於」時，可以勾「排除這一行空白的列」，把分類之間的 x、TTL 小計列拿掉。' }),
        els.chips),
      h('li', { class: 'task' },
        h('h2', { text: '要保留的行' }),
        h('div', { class: 'field-row' },
          h('label', { for: 'f-cols', text: '行號' }), els.cols),
        h('p', { class: 'hint', text: '連續的行可以寫成 U:Z。上傳後也可以直接點下面表格最上方的行號來選。' }),
        els.srcPreview,
        h('details', { class: 'advanced' },
          h('summary', { text: '報表版面不一樣時' }),
          h('div', { class: 'field-row' },
            h('label', { for: 'f-start', text: '資料從第' }), els.startRow, h('span', { text: '列開始（之前的列都當作表頭保留）' })),
          h('div', { class: 'field-row' },
            h('label', { for: 'f-selcol', text: '選取條件在' }), els.selCol, h('span', { text: '行' })))),
      h('li', { class: 'task' },
        h('h2', { text: '匯出' }),
        els.summary, els.outPreview,
        h('div', { class: 'actions' }, els.exportBtn, els.nextBtn, els.exportNote))));

    els.sheetSelect.addEventListener('change', function () {
      pickSheet(parseInt(els.sheetSelect.value, 10));
    });
    bindInput(els.values, 'values');
    bindInput(els.fCol, 'fCol');
    bindInput(els.fValues, 'fValues');
    els.fOp.addEventListener('change', function () { settings.fOp = els.fOp.value; recompute(); });
    els.skipBlank.addEventListener('change', function () { settings.fSkipBlank = els.skipBlank.checked; setMode('col'); recompute(); });
    [els.modeA, els.modeCol].forEach(function (r) {
      r.addEventListener('change', function () { if (r.checked) { settings.mode = r.value; recompute(); } });
    });
    // 點輸入框時自動切到那一種方式
    els.values.addEventListener('focus', function () { setMode('a'); });
    [els.fCol, els.fOp, els.fValues].forEach(function (el) { el.addEventListener('focus', function () { setMode('col'); }); });
    bindInput(els.cols, 'cols');
    bindInput(els.startRow, 'startRow');
    bindInput(els.selCol, 'selCol');

    if (S.workbook) showFile();
    recompute();
  }

  function setMode(mode) {
    if (settings.mode === mode) return;
    settings.mode = mode;
    recompute();
  }

  function bindInput(input, key) {
    input.addEventListener('input', function () {
      settings[key] = input.value;
      clearTimeout(timer);
      timer = setTimeout(recompute, 200);
    });
  }

  /* ---------- 開檔 ---------- */
  function openFile(file) {
    els.fileMsg.hidden = true;
    els.drop.setBusy(file.name);
    ui.readWorkbook(file)
      .then(function (res) {
        S.workbook = res.workbook;
        S.fileName = file.name;
        S.ws = null;
        showFile();
        pickSheet(0);
      })
      .catch(function (err) {
        els.fileMsg.textContent = err && err.message ? err.message : String(err);
        els.fileMsg.hidden = false;
        if (S.workbook) showFile(); else els.drop.setIdle();
      });
  }

  function showFile() {
    els.drop.setFile(S.fileName);
    var sheets = S.workbook.worksheets;
    els.sheetSelect.textContent = '';
    sheets.forEach(function (ws, i) {
      els.sheetSelect.appendChild(h('option', { value: String(i), text: ws.name }));
    });
    els.sheetRow.hidden = sheets.length < 2;
    if (S.ws) els.sheetSelect.value = String(sheets.indexOf(S.ws));
  }

  function pickSheet(index) {
    S.ws = S.workbook.worksheets[index];
    S.infoKey = '';
    els.sheetSelect.value = String(index);
    els.exportNote.textContent = '';
    recompute();
  }

  /* ---------- 重新計算並更新畫面 ---------- */
  function recompute() {
    saveSettings();
    S.result = null;
    IEApp.state.mpsFiltered = null;
    els.exportBtn.disabled = true;
    els.nextBtn.disabled = true;
    els.selLabel.textContent = XIO.numToCol(selCol());
    els.modeA.checked = !byColumn();
    els.modeCol.checked = byColumn();
    els.rowA.classList.toggle('is-off', byColumn());
    els.rowCol.classList.toggle('is-off', !byColumn());
    els.skipRow.hidden = !exclude();
    els.skipRow.classList.toggle('is-off', !byColumn());
    els.skipBlank.checked = settings.fSkipBlank !== false;
    els.blankList.textContent = '';

    if (!S.ws) {
      els.chips.textContent = '';
      els.srcPreview.textContent = '';
      els.outPreview.textContent = '';
      setSummary('先上傳 MPS 報表。');
      return;
    }

    var sr = startRow();
    var sc = activeCol();
    if (!sc) {
      els.chips.textContent = '';
      els.outPreview.textContent = '';
      var ext0 = XIO.extent(S.ws);
      ui.renderSheet(els.srcPreview, S.ws, {
        headerEnd: sr - 1, lastRow: ext0.lastRow, colCount: Math.max(ext0.colCount, 1), rows: PREVIEW_ROWS,
        keep: keepMap(), onToggle: toggleColumn, caption: '原始報表（前 ' + PREVIEW_ROWS + ' 列資料）。'
      });
      setSummary(XIO.normalizeText(settings.fCol) ? '看不懂「' + settings.fCol + '」。要篩選的行請輸入一個英文字母，例如 F。' : '輸入要篩選的是哪一行。',
        !!XIO.normalizeText(settings.fCol));
      return;
    }
    var key = sr + ':' + sc;
    if (S.infoKey !== key) {
      S.info = XIO.inspect(S.ws, { dataStartRow: sr, selectorCol: sc });
      S.infoKey = key;
    }
    var info = S.info;
    var values = wantedValues();
    var wanted = {};
    values.forEach(function (v) { wanted[XIO.matchKey(v)] = true; });
    var parsed = XIO.parseColumns(settings.cols);
    var keep = keepMap();

    renderChips(info, wanted);
    ui.renderSheet(els.srcPreview, S.ws, {
      headerEnd: sr - 1,
      lastRow: info.lastRow,
      colCount: Math.max(info.colCount, 1),
      rows: PREVIEW_ROWS,
      keep: keep,
      onToggle: toggleColumn,
      selectorCol: sc,
      wanted: values.length ? wanted : null,
      exclude: exclude(),
      caption: '原始報表（前 ' + PREVIEW_ROWS + ' 列資料）。共 ' + info.dataRows.toLocaleString('en-US') +
        ' 列資料、' + info.colCount + ' 行（到 ' + XIO.numToCol(info.colCount || 1) + ' 行）。'
    });

    els.outPreview.textContent = '';
    var colName = XIO.numToCol(sc);

    if (!info.dataRows) {
      setSummary('第 ' + sr + ' 列之後沒有資料。請確認上傳的檔案和工作表。', true);
      return;
    }
    var opWord = exclude() ? '不等於' : '等於';
    if (!values.length) { setSummary('輸入 ' + colName + ' 行要' + opWord + '的值。'); return; }
    if (parsed.bad.length) {
      setSummary('看不懂「' + parsed.bad.join('、') + '」。行號請用英文字母，例如 B 或 U:Z。', true);
      return;
    }
    if (!parsed.cols.length) { setSummary('輸入要保留的行。'); return; }

    // 這組條件存起來：匯出、以及第 2 步要重新產生活頁簿時都用同一組
    var ws = S.ws;
    var blanks = [];
    var keepBlank = {};
    if (skipBlank()) {
      var bkey = S.fileName + '|' + (ws.name || '') + '|' + sr + '|' + sc;
      if (S.blankKey !== bkey) { S.blankKey = bkey; S.keepBlank = {}; }
      blanks = XIO.blankRows(ws, { dataStartRow: sr, selectorCol: sc, lastRow: info.lastRow, keepCols: parsed.cols });
      blanks.forEach(function (b) { if (S.keepBlank[b.row]) keepBlank[b.row] = true; });
      renderBlankList(blanks, colName);
    }
    var opts = {
      dataStartRow: sr, selectorCol: sc, values: values, keepCols: parsed.cols, lastRow: info.lastRow,
      exclude: exclude(), skipBlank: skipBlank(), keepBlankRows: keepBlank
    };
    var result = XIO.filter(ws, opts);
    var skipped = blanks.length - Object.keys(keepBlank).length;

    if (!result.matchedRows) {
      var seen = info.values.slice(0, 8).map(function (v) { return v.label; }).join('、');
      setSummary(colName + ' 行沒有' + opWord + ' ' + values.join('、') + ' 的列。' +
        (seen ? '這份報表的 ' + colName + ' 行有：' + seen + '。' : colName + ' 行從第 ' + sr + ' 列起都是空白。'), true);
      return;
    }

    S.result = result;
    var text = colName + ' 行' + opWord + ' ' + values.join('、') + '：符合 ' + result.matchedRows.toLocaleString('en-US') + ' 列（共 ' +
      result.totalRows.toLocaleString('en-US') + ' 列資料），保留 ' + parsed.cols.length + ' 行。';
    if (skipBlank() && blanks.length) {
      text += ' 另外排除 ' + colName + ' 行空白的 ' + skipped + ' 列' + (skipped < blanks.length ? '（保留 ' + (blanks.length - skipped) + ' 列）' : '') + '。';
    }
    var beyond = parsed.cols.filter(function (c) { return c > info.colCount; });
    if (beyond.length) text += ' 注意：' + XIO.formatColumns(beyond) + ' 行在報表裡沒有資料。';
    setSummary(text, false, beyond.length > 0);

    ui.renderSheet(els.outPreview, result.worksheet, {
      headerEnd: result.headerEnd,
      lastRow: result.lastRow,
      colCount: result.colCount,
      rows: PREVIEW_ROWS,
      caption: '匯出後的樣子（前 ' + PREVIEW_ROWS + ' 列資料）。'
    });
    els.exportBtn.disabled = false;
    els.nextBtn.disabled = false;

    // 給第 2 步用。build() 每次都產生一份全新的活頁簿，後面的步驟可以放心往裡面加工作表。
    IEApp.state.mpsFiltered = {
      sourceName: S.fileName,
      outName: outputName(),
      sheetName: result.worksheet.name,
      matchedRows: result.matchedRows,
      headerEnd: result.headerEnd,
      build: function () { return XIO.filter(ws, opts).workbook; }
    };
  }

  /* ---------- 「不等於」時，選取行空白的列清單（打勾＝排除，取消勾選＝保留） ---------- */
  function renderBlankList(blanks, colName) {
    els.blankList.textContent = '';
    if (!blanks.length) {
      els.blankList.appendChild(h('p', { class: 'blank-none', text: colName + ' 行沒有空白的列。' }));
      return;
    }
    var kept = blanks.filter(function (b) { return S.keepBlank[b.row]; }).length;
    var head = h('div', { class: 'blank-head' },
      h('span', { text: colName + ' 行空白的列：' + blanks.length + ' 列，打勾的會排除' + (kept ? '（已取消 ' + kept + ' 列，會保留）' : '') }),
      h('button', { type: 'button', class: 'link-btn', text: '全部排除', onclick: function () { S.keepBlank = {}; recompute(); } }),
      h('button', { type: 'button', class: 'link-btn', text: '全部保留', onclick: function () {
        blanks.forEach(function (b) { S.keepBlank[b.row] = true; }); recompute();
      } }));
    var list = h('ul', { class: 'blank-items' });
    blanks.slice(0, MAX_BLANK_LIST).forEach(function (b) {
      var id = 'f-blank-' + b.row;
      var box = h('input', { type: 'checkbox', id: id });
      box.checked = !S.keepBlank[b.row];
      box.addEventListener('change', function () {
        if (box.checked) delete S.keepBlank[b.row]; else S.keepBlank[b.row] = true;
        var scroller = els.blankList.querySelector('.blank-items');
        var top = scroller ? scroller.scrollTop : 0;
        recompute();
        scroller = els.blankList.querySelector('.blank-items');
        if (scroller) scroller.scrollTop = top;
        var again = document.getElementById(id);
        if (again) again.focus({ preventScroll: true });
      });
      list.appendChild(h('li', { class: box.checked ? '' : 'is-kept' },
        h('label', { for: id }, box,
          h('span', { class: 'blank-row', text: '第 ' + b.row + ' 列' }),
          h('span', { class: 'blank-text', text: b.text || '（保留的行以外有資料）' }))));
    });
    if (blanks.length > MAX_BLANK_LIST) {
      list.appendChild(h('li', { class: 'blank-more', text: '另有 ' + (blanks.length - MAX_BLANK_LIST) + ' 列沒有列出（一樣會排除）' }));
    }
    els.blankList.appendChild(head);
    els.blankList.appendChild(list);
  }

  function keepMap() {
    var keep = {};
    XIO.parseColumns(settings.cols).cols.forEach(function (c) { keep[c] = true; });
    return keep;
  }

  function setSummary(text, isError, isWarn) {
    els.summary.textContent = text;
    els.summary.className = 'summary' + (isError ? ' is-error' : '') + (isWarn ? ' is-warn' : '');
  }

  /* ---------- 選取行出現過的值（點一下就填入） ---------- */
  function renderChips(info, wanted) {
    els.chips.textContent = '';
    if (!info.values.length && !info.blankCount) return;
    var col = XIO.numToCol(activeCol());
    els.chips.appendChild(h('span', { class: 'chips-label', text: '報表裡 ' + col + ' 行的值' + (exclude() ? '（點選的是不要的）' : '') }));
    info.values.slice(0, MAX_CHIPS).forEach(function (v) {
      var on = !!wanted[v.key];
      els.chips.appendChild(h('button', {
        type: 'button', class: 'chip', 'aria-pressed': on ? 'true' : 'false',
        onclick: function () { toggleValue(v); }
      }, h('span', { class: 'chip-value', text: v.label }),
         h('span', { class: 'chip-count', text: v.count.toLocaleString('en-US') + ' 列' })));
    });
    if (info.values.length > MAX_CHIPS) {
      els.chips.appendChild(h('span', { class: 'chips-more', text: '另有 ' + (info.values.length - MAX_CHIPS) + ' 種' }));
    }
    if (info.blankCount) {
      els.chips.appendChild(h('span', { class: 'chips-more', text: '空白 ' + info.blankCount.toLocaleString('en-US') + ' 列' }));
    }
  }

  function toggleValue(v) {
    var list = wantedValues();
    var kept = list.filter(function (x) { return XIO.matchKey(x) !== v.key; });
    if (kept.length === list.length) kept.push(v.label);
    var k = valuesKey();
    settings[k] = kept.join(', ');
    (k === 'fValues' ? els.fValues : els.values).value = settings[k];
    recompute();
  }

  function toggleColumn(c) {
    var cols = XIO.parseColumns(settings.cols).cols;
    var i = cols.indexOf(c);
    if (i >= 0) cols.splice(i, 1); else cols.push(c);
    cols.sort(function (a, b) { return a - b; });
    settings.cols = XIO.formatColumns(cols);
    els.cols.value = settings.cols;
    var scroller = els.srcPreview.querySelector('.sheet-scroll');
    var left = scroller ? scroller.scrollLeft : 0;
    recompute();
    scroller = els.srcPreview.querySelector('.sheet-scroll');
    if (scroller) scroller.scrollLeft = left;
    var btn = els.srcPreview.querySelector('button[data-col="' + c + '"]');
    if (btn) btn.focus({ preventScroll: true });
  }

  /* ---------- 匯出 ---------- */
  function onExport() {
    var ready = IEApp.state.mpsFiltered;
    if (!S.result || !ready) return;
    var rows = ready.matchedRows;
    var name = ready.outName;
    els.exportBtn.disabled = true;
    els.exportNote.textContent = '產生檔案中…';
    XIO.toBlob(ready.build()).then(function (blob) {
      ui.download(blob, name);
      els.exportNote.textContent = '已匯出 ' + name + '（' + rows.toLocaleString('en-US') + ' 列），在瀏覽器的下載資料夾。';
    }).catch(function (err) {
      els.exportNote.textContent = '匯出失敗：' + (err && err.message ? err.message : err);
    }).then(function () {
      els.exportBtn.disabled = false;
    });
  }

  IEApp.registerStep({ id: 'mps-filter', render: render });
})(window);
