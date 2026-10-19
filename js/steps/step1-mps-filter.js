/*
 * 第 1 步：篩選 MPS 報表
 *   1. 上傳 MPS 的 Excel
 *   2. 輸入「A 行等於多少」——符合的列才留下（從第 5 列開始比對）
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

  var DEFAULTS = { values: '', cols: '', startRow: 5, selCol: 'A' };

  // 這一步的狀態（切到別的步驟再回來時還在）
  var S = { fileName: '', workbook: null, ws: null, info: null, infoKey: '', result: null };
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
  function wantedValues() {
    return XIO.normalizeText(settings.values).split(/[,、;]+/)
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
      h('p', { text: '留下 A 行等於指定數值的列，以及你要的行，匯出成新的 Excel。表頭（第 1–4 列）會原樣保留。' })));

    panel.appendChild(h('ol', { class: 'tasks' },
      h('li', { class: 'task' },
        h('h2', { text: '上傳 MPS 報表' }),
        els.drop.el, els.sheetRow, els.fileMsg),
      h('li', { class: 'task' },
        h('h2', { text: '選取條件' }),
        h('div', { class: 'field-row' },
          h('label', { for: 'f-values' }, els.selLabel, ' 行等於'), els.values),
        h('p', { class: 'hint', text: '要選多個數值時用逗號分開，例如 1, 3。' }),
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
    bindInput(els.cols, 'cols');
    bindInput(els.startRow, 'startRow');
    bindInput(els.selCol, 'selCol');

    if (S.workbook) showFile();
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

    if (!S.ws) {
      els.chips.textContent = '';
      els.srcPreview.textContent = '';
      els.outPreview.textContent = '';
      setSummary('先上傳 MPS 報表。');
      return;
    }

    var sr = startRow();
    var sc = selCol();
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
    var keep = {};
    parsed.cols.forEach(function (c) { keep[c] = true; });

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
      caption: '原始報表（前 ' + PREVIEW_ROWS + ' 列資料）。共 ' + info.dataRows.toLocaleString('en-US') +
        ' 列資料、' + info.colCount + ' 行（到 ' + XIO.numToCol(info.colCount || 1) + ' 行）。'
    });

    els.outPreview.textContent = '';
    var colName = XIO.numToCol(sc);

    if (!info.dataRows) {
      setSummary('第 ' + sr + ' 列之後沒有資料。請確認上傳的檔案和工作表。', true);
      return;
    }
    if (!values.length) { setSummary('輸入 ' + colName + ' 行要等於的數值。'); return; }
    if (parsed.bad.length) {
      setSummary('看不懂「' + parsed.bad.join('、') + '」。行號請用英文字母，例如 B 或 U:Z。', true);
      return;
    }
    if (!parsed.cols.length) { setSummary('輸入要保留的行。'); return; }

    // 這組條件存起來：匯出、以及第 2 步要重新產生活頁簿時都用同一組
    var ws = S.ws;
    var opts = {
      dataStartRow: sr, selectorCol: sc, values: values, keepCols: parsed.cols, lastRow: info.lastRow
    };
    var result = XIO.filter(ws, opts);

    if (!result.matchedRows) {
      var seen = info.values.slice(0, 8).map(function (v) { return v.label; }).join('、');
      setSummary(colName + ' 行沒有等於 ' + values.join('、') + ' 的列。' +
        (seen ? '這份報表的 ' + colName + ' 行有：' + seen + '。' : colName + ' 行從第 ' + sr + ' 列起都是空白。'), true);
      return;
    }

    S.result = result;
    var text = '符合 ' + result.matchedRows.toLocaleString('en-US') + ' 列（共 ' +
      result.totalRows.toLocaleString('en-US') + ' 列資料），保留 ' + parsed.cols.length + ' 行。';
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

  function setSummary(text, isError, isWarn) {
    els.summary.textContent = text;
    els.summary.className = 'summary' + (isError ? ' is-error' : '') + (isWarn ? ' is-warn' : '');
  }

  /* ---------- 選取行出現過的值（點一下就填入） ---------- */
  function renderChips(info, wanted) {
    els.chips.textContent = '';
    if (!info.values.length && !info.blankCount) return;
    var col = XIO.numToCol(selCol());
    els.chips.appendChild(h('span', { class: 'chips-label', text: '報表裡 ' + col + ' 行的值' }));
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
    settings.values = kept.join(', ');
    els.values.value = settings.values;
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
