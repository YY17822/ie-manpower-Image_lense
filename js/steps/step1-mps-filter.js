/*
 * 第 1 步：篩選 MPS 報表
 *   1. 上傳 MPS 的 Excel
 *   2. 選取條件，三種方式選一種：
 *        a. A 行等於多少（原本的方式；A 行可在「報表版面不一樣時」改成別的行）
 *        b. 指定某一行，等於或不等於某些值
 *           「不等於」時可勾「排除這一行空白的列」（預設勾選），把分類之間的「x」「TTL (…)」
 *           小計列拿掉；被排除的列會逐列列出，取消勾選的那列會保留（不記住，因為每個月列號不同）
 *        c. 指定某一行，模糊比對：預設三個輸入框（可再加），儲存格含有任一框的字就算符合
 *           （英文大小寫不分）；含有的列要「選取」或「排除」二選一；那一行空白的列一律不留。
 *           下面逐列列出會留下的列，方便檢查
 *      符合的列才留下（從第 5 列開始比對）
 *   3. 輸入要保留哪些行
 *   4. 挑出 RMA：從保留的行選一行，模糊比對（含有）或等於任一個值（預設三個輸入框，可再加）。
 *      符合的列逐列列出、預設打勾；取消勾選的列留在一般 MPS（不記住，因為每個月列號不同）。
 *      RMA 的列移到一般 MPS 最後一列之後第 8 列起：空一列、三列留給第 2 步的匯總（Total／Top 3／Non-Top 3）、
 *      再空三列，然後是一列「RMA」標題（跨保留的行合併置中），下面是 RMA 的列。沒有 RMA 時一樣留標題。
 *   5. 匯出新的 Excel，第 1–4 列表頭（含合併儲存格）原樣保留
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
  // mode 'fz' = 指定某一行模糊比對：zCol 哪一行、zOp（'eq' 含有的列選取／'ne' 含有的列排除）、zValues 各輸入框的值（陣列）
  var DEFAULTS = {
    values: '', cols: '', startRow: 5, selCol: 'A', mode: 'a', fCol: '', fOp: 'eq', fValues: '', fSkipBlank: true,
    zCol: '', zOp: 'eq', zValues: ['', '', ''],
    rmaCol: '', rmaOp: 'fz', rmaValues: ['', '', '']    // 挑出 RMA：哪一行、'fz' 模糊比對／'eq' 等於、各輸入框的值
  };
  var MAX_BLANK_LIST = 200;  // 空白列清單最多列幾列
  var MAX_FZ_LIST = 300;     // 模糊比對預覽最多列幾列

  // 這一步的狀態（切到別的步驟再回來時還在）
  // keepBlank：空白列清單裡使用者取消勾選（要保留）的列號 { 52: true }；換檔案、換行、換起始列時清空
  // rmaOff：挑出 RMA 的清單裡取消勾選（留在一般 MPS）的列號 { 52: true }；換檔案、換工作表、換起始列時清空
  var S = { fileName: '', workbook: null, ws: null, info: null, infoKey: '', result: null, keepBlank: {}, blankKey: '', rmaOff: {}, rmaKey: '' };
  var settings = loadSettings();
  var els = {};
  var timer = null;

  /* ---------- 記住上次輸入的條件（每個月做的事情差不多） ---------- */
  function loadSettings() {
    var s = {};
    try { s = IEApp.settings.read(STORE_KEY) || {}; } catch (e) { s = {}; }
    var out = {};
    Object.keys(DEFAULTS).forEach(function (k) { out[k] = s[k] === undefined ? DEFAULTS[k] : s[k]; });
    out.zValues = Array.isArray(out.zValues) && out.zValues.length ? out.zValues.map(function (v) { return String(v == null ? '' : v); }) : ['', '', ''];
    out.rmaValues = Array.isArray(out.rmaValues) && out.rmaValues.length ? out.rmaValues.map(function (v) { return String(v == null ? '' : v); }) : ['', '', ''];
    out.rmaOp = out.rmaOp === 'eq' ? 'eq' : 'fz';
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
  function byFuzzy() { return settings.mode === 'fz'; }
  function exclude() { return byFuzzy() ? settings.zOp === 'ne' : byColumn() && settings.fOp === 'ne'; }
  function skipBlank() { return byColumn() && exclude() && settings.fSkipBlank !== false; }
  // 模糊比對的行：看不懂或沒填時回傳 0
  function zCol() {
    var p = XIO.parseColumns(settings.zCol);
    return p.cols.length === 1 && !p.bad.length ? p.cols[0] : 0;
  }
  // 指定某一行的方式：行號看不懂或沒填時回傳 0
  function fCol() {
    var p = XIO.parseColumns(settings.fCol);
    return p.cols.length === 1 && !p.bad.length ? p.cols[0] : 0;
  }
  // 目前用來篩選的行
  function activeCol() { return byFuzzy() ? zCol() : byColumn() ? fCol() : selCol(); }
  function valuesKey() { return byColumn() ? 'fValues' : 'values'; }
  function wantedValues() {
    if (byFuzzy()) {
      return settings.zValues.map(function (v) { return XIO.normalizeText(v); })
        .filter(function (v) { return v !== ''; });
    }
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

    // 第三種：指定某一行，模糊比對
    els.modeFz = h('input', { type: 'radio', name: 'f-mode', id: 'f-mode-fz', value: 'fz' });
    els.zCol = h('input', {
      id: 'f-zcol', type: 'text', class: 'in in-tiny', autocomplete: 'off', spellcheck: 'false',
      placeholder: '例如 C', value: settings.zCol, 'aria-label': '要模糊比對的行'
    });
    // 含有任一個值的列：選取（留下）或排除（不要）
    els.zOpSel = h('input', { type: 'radio', name: 'f-zop', id: 'f-zop-eq', value: 'eq' });
    els.zOpExc = h('input', { type: 'radio', name: 'f-zop', id: 'f-zop-ne', value: 'ne' });
    els.zOpBox = h('span', { class: 'fz-op', role: 'radiogroup', 'aria-label': '含有的列要選取還是排除' },
      h('label', { class: 'check-label', for: 'f-zop-eq' }, els.zOpSel, h('span', { text: '選取' })),
      h('label', { class: 'check-label', for: 'f-zop-ne' }, els.zOpExc, h('span', { text: '排除' })));
    els.rowFz = h('div', { class: 'field-row mode-body' },
      els.zCol, h('span', { text: '行含有下面任一個值的列：' }), els.zOpBox);
    els.zBox = h('div', { class: 'fz-values mode-body' });
    els.zAdd = h('button', { type: 'button', class: 'link-btn fz-add', text: '＋ 再加一個輸入框', onclick: addZInput });
    els.zAddRow = h('div', { class: 'mode-body' }, els.zAdd);
    els.zPreview = h('div', { class: 'fz-preview mode-body', 'aria-live': 'polite' });
    buildZInputs();

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

    // 挑出 RMA
    els.rmaCol = h('select', { id: 'f-rmacol' });
    els.rmaOp = h('select', { id: 'f-rmaop', 'aria-label': '比對方式' },
      h('option', { value: 'fz', text: '模糊比對（含有）' }), h('option', { value: 'eq', text: '等於' }));
    els.rmaOp.value = settings.rmaOp;
    els.rmaBox = h('div', { class: 'fz-values' });
    els.rmaAdd = h('button', { type: 'button', class: 'link-btn fz-add', text: '＋ 再加一個輸入框', onclick: addRmaInput });
    els.rmaChips = h('div', { class: 'chips' });
    els.rmaList = h('div', { class: 'fz-preview', 'aria-live': 'polite' });
    buildRmaInputs();

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
      h('p', { text: '依選取條件留下要的列（A 行等於指定數值、指定某一行等於／不等於某些值，或指定某一行模糊比對），以及你要的行，再挑出哪些列是 RMA，匯出成新的 Excel。表頭（第 1–4 列）會原樣保留。' })));

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
            els.rowCol, els.skipRow, els.blankList),
          h('div', { class: 'mode-option' },
            h('label', { class: 'mode-label', for: 'f-mode-fz' }, els.modeFz, h('span', { text: '指定某一行，模糊比對（含有某些字）' })),
            els.rowFz, els.zBox, els.zAddRow, els.zPreview)),
        h('p', { class: 'hint', text: '三種方式選一種。前兩種要選多個值時用逗號分開，例如 1, 3。選「不等於」時，可以勾「排除這一行空白的列」，把分類之間的 x、TTL 小計列拿掉。模糊比對：儲存格含有任一個輸入框的字就算符合（英文大小寫不分），符合的列可選「選取」（留下）或「排除」（不要）；那一行空白的列不會留下。' }),
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
        h('h2', { text: '挑出 RMA' }),
        h('p', { class: 'hint', text: '從保留的行裡挑出哪些列是 RMA。這些列會移到一般 MPS 的 Total 匯總下面（再空三列），上面加一列「RMA」標題；CT.、需求工時和一般 MPS 一樣算。這個月沒有 RMA 時，輸入框留空白就好（一樣會留一個空的 RMA 區塊，合計是 0）。' }),
        h('div', { class: 'field-row' },
          h('label', { for: 'f-rmacol', text: '看哪一行' }), els.rmaCol, els.rmaOp, h('span', { text: '下面任一個值' })),
        els.rmaBox, h('div', null, els.rmaAdd), els.rmaChips, els.rmaList),
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
    bindInput(els.zCol, 'zCol');
    [els.zOpSel, els.zOpExc].forEach(function (r) {
      r.addEventListener('change', function () { if (r.checked) { settings.zOp = r.value; settings.mode = 'fz'; recompute(); } });
    });
    [els.zCol, els.zOpSel, els.zOpExc].forEach(function (el) { el.addEventListener('focus', function () { setMode('fz'); }); });
    [els.modeA, els.modeCol, els.modeFz].forEach(function (r) {
      r.addEventListener('change', function () { if (r.checked) { settings.mode = r.value; recompute(); } });
    });
    // 點輸入框時自動切到那一種方式
    els.values.addEventListener('focus', function () { setMode('a'); });
    [els.fCol, els.fOp, els.fValues].forEach(function (el) { el.addEventListener('focus', function () { setMode('col'); }); });
    els.rmaCol.addEventListener('change', function () { settings.rmaCol = els.rmaCol.value; recompute(); });
    els.rmaOp.addEventListener('change', function () { settings.rmaOp = els.rmaOp.value === 'eq' ? 'eq' : 'fz'; recompute(); });
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

  /* ---------- 挑出 RMA 的輸入框（預設三個，可以再加、可以刪） ---------- */
  function buildRmaInputs(focusIndex) {
    els.rmaBox.textContent = '';
    var list = settings.rmaValues;
    list.forEach(function (val, i) {
      var id = 'f-rmaval-' + i;
      var input = h('input', {
        id: id, type: 'text', class: 'in in-short', autocomplete: 'off', spellcheck: 'false',
        placeholder: i === 0 ? '例如 RMA' : '', value: val
      });
      input.addEventListener('input', function () {
        settings.rmaValues[i] = input.value;
        clearTimeout(timer);
        timer = setTimeout(recompute, 200);
      });
      var row = h('div', { class: 'fz-value' },
        h('label', { class: 'fz-label', for: id, text: '值 ' + (i + 1) }), input);
      if (list.length > 1) {
        row.appendChild(h('button', {
          type: 'button', class: 'link-btn', text: '刪除', 'aria-label': '刪除 RMA 的值 ' + (i + 1),
          onclick: function () { removeRmaInput(i); }
        }));
      }
      els.rmaBox.appendChild(row);
    });
    if (focusIndex !== undefined) {
      var el = document.getElementById('f-rmaval-' + focusIndex);
      if (el) el.focus();
    }
  }
  function addRmaInput() {
    settings.rmaValues.push('');
    buildRmaInputs(settings.rmaValues.length - 1);
    recompute();
  }
  function removeRmaInput(i) {
    settings.rmaValues.splice(i, 1);
    if (!settings.rmaValues.length) settings.rmaValues.push('');
    buildRmaInputs(Math.min(i, settings.rmaValues.length - 1));
    recompute();
  }
  function rmaValues() {
    return settings.rmaValues.map(function (v) { return XIO.normalizeText(v); }).filter(function (v) { return v !== ''; });
  }
  // 「看哪一行」的選單：只列保留的行（行號＋表頭文字）
  function fillRmaCols(cols) {
    var labels = S.ws ? XIO.columnLabels(S.ws, startRow() > 2 ? 2 : 1, Math.max(startRow() - 1, 1), Math.max.apply(null, cols.concat([1]))) : [];
    var chosen = XIO.normalizeText(settings.rmaCol).toUpperCase();
    els.rmaCol.textContent = '';
    els.rmaCol.appendChild(h('option', { value: '', text: '請選擇' }));
    cols.forEach(function (c) {
      var letter = XIO.numToCol(c);
      var label = labels[c] ? letter + '　' + labels[c] : letter;
      if (label.length > 42) label = label.slice(0, 41) + '…';
      els.rmaCol.appendChild(h('option', { value: letter, text: label }));
    });
    var ok = cols.some(function (c) { return XIO.numToCol(c) === chosen; });
    els.rmaCol.value = ok ? chosen : '';
    return ok ? XIO.colToNum(chosen) : 0;
  }
  function rmaMessage(text) {
    els.rmaChips.textContent = '';
    els.rmaList.textContent = '';
    els.rmaList.appendChild(h('p', { class: 'blank-none', text: text }));
  }

  /*
   * 挑出 RMA 的值按鈕和逐列清單（打勾＝RMA，取消勾選＝留在一般 MPS）
   *   kept   符合選取條件的列號；col 看哪一行；picks 符合的列
   */
  function renderRma(kept, col, picks, values) {
    els.rmaChips.textContent = '';
    els.rmaList.textContent = '';
    var colName = XIO.numToCol(col);
    var seen = XIO.columnValues(S.ws, kept, col);
    if (seen.length) {
      var have = {};
      values.forEach(function (v) { have[XIO.fuzzyText(v)] = true; });
      els.rmaChips.appendChild(h('span', { class: 'chips-label', text: '篩選出來的列裡 ' + colName + ' 行的值' }));
      seen.slice(0, MAX_CHIPS).forEach(function (v) {
        els.rmaChips.appendChild(h('button', {
          type: 'button', class: 'chip', 'aria-pressed': have[v.key] ? 'true' : 'false',
          onclick: function () { toggleRmaValue(v); }
        }, h('span', { class: 'chip-value', text: v.label }),
           h('span', { class: 'chip-count', text: v.count.toLocaleString('en-US') + ' 列' })));
      });
      if (seen.length > MAX_CHIPS) els.rmaChips.appendChild(h('span', { class: 'chips-more', text: '另有 ' + (seen.length - MAX_CHIPS) + ' 種' }));
    }
    if (!values.length) {
      els.rmaList.appendChild(h('p', { class: 'blank-none', text: '沒有輸入值：這個月不分 RMA，匯出時留一個空的 RMA 區塊。' }));
      return;
    }
    if (!picks.length) {
      els.rmaList.appendChild(h('p', { class: 'blank-none', text: colName + ' 行沒有符合的列。RMA 區塊會是空的（RMA Total、RMA總需求工時都是 0）。' }));
      return;
    }
    var off = picks.filter(function (p) { return S.rmaOff[p.row]; }).length;
    els.rmaList.appendChild(h('div', { class: 'blank-head' },
      h('span', { class: 'fz-count', text: 'RMA：' + (picks.length - off).toLocaleString('en-US') + ' 列' }),
      h('span', { text: off ? '（取消勾選 ' + off + ' 列，留在一般 MPS）' : '（打勾的會移到 RMA 區塊）' }),
      h('button', { type: 'button', class: 'link-btn', text: '全部勾選', onclick: function () { S.rmaOff = {}; recompute(); } }),
      h('button', { type: 'button', class: 'link-btn', text: '全部取消', onclick: function () {
        picks.forEach(function (p) { S.rmaOff[p.row] = true; }); recompute();
      } })));
    var terms = settings.rmaOp === 'eq' ? [] : XIO.fuzzyTerms(values);
    var list = h('ul', { class: 'blank-items' });
    picks.slice(0, MAX_FZ_LIST).forEach(function (p) {
      var id = 'f-rma-' + p.row;
      var box = h('input', { type: 'checkbox', id: id });
      box.checked = !S.rmaOff[p.row];
      box.addEventListener('change', function () {
        if (box.checked) delete S.rmaOff[p.row]; else S.rmaOff[p.row] = true;
        var scroller = els.rmaList.querySelector('.blank-items');
        var top = scroller ? scroller.scrollTop : 0;
        recompute();
        scroller = els.rmaList.querySelector('.blank-items');
        if (scroller) scroller.scrollTop = top;
        var again = document.getElementById(id);
        if (again) again.focus({ preventScroll: true });
      });
      list.appendChild(h('li', { class: box.checked ? '' : 'is-kept' },
        h('label', { for: id }, box,
          h('span', { class: 'blank-row', text: '第 ' + p.row + ' 列' }),
          markText(p.value, terms),
          h('span', { class: 'blank-text fz-ctx', text: p.text }))));
    });
    if (picks.length > MAX_FZ_LIST) {
      list.appendChild(h('li', { class: 'blank-more', text: '另有 ' + (picks.length - MAX_FZ_LIST) + ' 列沒有列出（一樣會移到 RMA）' }));
    }
    els.rmaList.appendChild(list);
  }
  // 點值的按鈕：已經在輸入框裡就清掉，不在就填進第一個空的輸入框（沒有空的就加一個）
  function toggleRmaValue(v) {
    var vs = settings.rmaValues;
    var at = -1, empty = -1;
    vs.forEach(function (x, i) {
      if (at < 0 && XIO.normalizeText(x) !== '' && XIO.fuzzyText(x) === v.key) at = i;
      if (empty < 0 && XIO.normalizeText(x) === '') empty = i;
    });
    if (at >= 0) vs[at] = '';
    else if (empty >= 0) vs[empty] = v.label;
    else vs.push(v.label);
    buildRmaInputs();
    recompute();
  }

  /* ---------- 模糊比對的輸入框（預設三個，可以再加、可以刪） ---------- */
  function buildZInputs(focusIndex) {
    els.zBox.textContent = '';
    var list = settings.zValues;
    list.forEach(function (val, i) {
      var id = 'f-zval-' + i;
      var input = h('input', {
        id: id, type: 'text', class: 'in in-short', autocomplete: 'off', spellcheck: 'false',
        placeholder: i === 0 ? '例如 Lens' : '', value: val
      });
      input.addEventListener('input', function () {
        settings.zValues[i] = input.value;
        clearTimeout(timer);
        timer = setTimeout(recompute, 200);
      });
      input.addEventListener('focus', function () { setMode('fz'); });
      var row = h('div', { class: 'fz-value' },
        h('label', { class: 'fz-label', for: id, text: '值 ' + (i + 1) }), input);
      if (list.length > 1) {
        row.appendChild(h('button', {
          type: 'button', class: 'link-btn', text: '刪除', 'aria-label': '刪除值 ' + (i + 1),
          onclick: function () { removeZInput(i); }
        }));
      }
      els.zBox.appendChild(row);
    });
    if (focusIndex !== undefined) {
      var el = document.getElementById('f-zval-' + focusIndex);
      if (el) el.focus();
    }
  }
  function addZInput() {
    settings.zValues.push('');
    settings.mode = 'fz';
    buildZInputs(settings.zValues.length - 1);
    recompute();
  }
  function removeZInput(i) {
    settings.zValues.splice(i, 1);
    if (!settings.zValues.length) settings.zValues.push('');
    settings.mode = 'fz';
    buildZInputs(Math.min(i, settings.zValues.length - 1));
    recompute();
  }

  /* ---------- 模糊比對的預覽：逐列列出會留下的列 ---------- */
  function renderFzPreview(sr, sc, lastRow, values, colName) {
    els.zPreview.textContent = '';
    var res = XIO.fuzzyRows(S.ws, { dataStartRow: sr, selectorCol: sc, lastRow: lastRow, values: values, exclude: exclude() });
    var terms = XIO.fuzzyTerms(values);
    var note = exclude()
      ? '（' + colName + ' 行含有任一值的 ' + res.hit.toLocaleString('en-US') + ' 列排除'
      : '（';
    note += (exclude() ? '；' : '') + colName + ' 行空白的 ' + res.blank.toLocaleString('en-US') + ' 列不留）';
    els.zPreview.appendChild(h('div', { class: 'blank-head' },
      h('span', { class: 'fz-count', text: '篩選出來的列：' + res.rows.length.toLocaleString('en-US') + ' 列' }),
      h('span', { text: note })));
    if (!res.rows.length) {
      els.zPreview.appendChild(h('p', { class: 'blank-none', text: '沒有符合的列。' }));
      return;
    }
    var list = h('ul', { class: 'blank-items fz-items' });
    res.rows.slice(0, MAX_FZ_LIST).forEach(function (it) {
      list.appendChild(h('li', { class: 'fz-item' },
        h('span', { class: 'blank-row', text: '第 ' + it.row + ' 列' }),
        markText(it.value, terms),
        h('span', { class: 'blank-text fz-ctx', text: it.text })));
    });
    if (res.rows.length > MAX_FZ_LIST) {
      list.appendChild(h('li', { class: 'blank-more', text: '另有 ' + (res.rows.length - MAX_FZ_LIST).toLocaleString('en-US') + ' 列沒有列出（一樣會留下）' }));
    }
    els.zPreview.appendChild(list);
  }
  // 把含有的字用螢光筆標出來
  function markText(text, terms) {
    var box = h('span', { class: 'fz-val' });
    var low = text.toLowerCase();
    var pos = 0;
    while (pos < text.length) {
      var best = -1, len = 0;
      terms.forEach(function (t) {
        var j = low.indexOf(t, pos);
        if (j >= 0 && (best < 0 || j < best || (j === best && t.length > len))) { best = j; len = t.length; }
      });
      if (best < 0 || len === 0) { box.appendChild(document.createTextNode(text.slice(pos))); break; }
      if (best > pos) box.appendChild(document.createTextNode(text.slice(pos, best)));
      box.appendChild(h('mark', { text: text.slice(best, best + len) }));
      pos = best + len;
    }
    return box;
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
    els.modeA.checked = !byColumn() && !byFuzzy();
    els.modeCol.checked = byColumn();
    els.modeFz.checked = byFuzzy();
    els.rowA.classList.toggle('is-off', byColumn() || byFuzzy());
    els.rowCol.classList.toggle('is-off', !byColumn());
    els.skipRow.hidden = !(byColumn() && exclude());
    els.skipRow.classList.toggle('is-off', !byColumn());
    els.skipBlank.checked = settings.fSkipBlank !== false;
    els.blankList.textContent = '';
    els.zOpSel.checked = settings.zOp !== 'ne';
    els.zOpExc.checked = settings.zOp === 'ne';
    [els.rowFz, els.zBox, els.zAddRow].forEach(function (el) { el.classList.toggle('is-off', !byFuzzy()); });
    els.zPreview.textContent = '';
    fillRmaCols(XIO.parseColumns(settings.cols).cols);
    rmaMessage('先完成上面的選取條件和要保留的行。');

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
      var colText = byFuzzy() ? settings.zCol : settings.fCol;
      setSummary(XIO.normalizeText(colText) ? '看不懂「' + colText + '」。要篩選的行請輸入一個英文字母，例如 F。' : '輸入要篩選的是哪一行。',
        !!XIO.normalizeText(colText));
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
      wanted: !byFuzzy() && values.length ? wanted : null,
      fuzzyTerms: byFuzzy() && values.length ? XIO.fuzzyTerms(values) : null,
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
    var opWord = byFuzzy() ? (exclude() ? '不含' : '含有') : (exclude() ? '不等於' : '等於');
    if (!values.length) { setSummary('輸入 ' + colName + ' 行要' + opWord + '的值。'); return; }
    if (byFuzzy()) renderFzPreview(sr, sc, info.lastRow, values, colName);
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
      exclude: exclude(), skipBlank: skipBlank(), keepBlankRows: keepBlank, fuzzy: byFuzzy(),
      rmaBlock: true, rmaRows: {}
    };

    // 挑出 RMA：在符合選取條件的列裡找，打勾的列放到 RMA 區塊
    var rkey = S.fileName + '|' + (ws.name || '') + '|' + sr;
    if (S.rmaKey !== rkey) { S.rmaKey = rkey; S.rmaOff = {}; }
    var rmaCol = fillRmaCols(parsed.cols);
    var kept = XIO.filterRows(ws, opts);
    if (!rmaCol) {
      rmaMessage(XIO.normalizeText(settings.rmaCol) ? settings.rmaCol + ' 行不在保留的行裡，請重新選擇。沒有選的話，匯出時留一個空的 RMA 區塊。'
        : '選擇要看哪一行（例如料號或說明）。沒有選的話，匯出時留一個空的 RMA 區塊。');
    } else {
      var rv = rmaValues();
      var picks = XIO.pickRows(ws, kept, rmaCol, settings.rmaOp, rv, parsed.cols);
      picks.forEach(function (p) { if (!S.rmaOff[p.row]) opts.rmaRows[p.row] = true; });
      renderRma(kept, rmaCol, picks, rv);
    }
    var result = XIO.filter(ws, opts);
    var skipped = blanks.length - Object.keys(keepBlank).length;

    if (!result.matchedRows) {
      var seen = info.values.slice(0, 8).map(function (v) { return v.label; }).join('、');
      setSummary(colName + ' 行沒有' + opWord + ' ' + values.join('、') + ' 的列。' +
        (seen ? '這份報表的 ' + colName + ' 行有：' + seen + '。' : colName + ' 行從第 ' + sr + ' 列起都是空白。'), true);
      return;
    }

    S.result = result;
    var text = colName + ' 行' + opWord + ' ' + values.join('、') + (byFuzzy() ? ' 任一個' : '') + '：符合 ' + result.matchedRows.toLocaleString('en-US') + ' 列（共 ' +
      result.totalRows.toLocaleString('en-US') + ' 列資料），保留 ' + parsed.cols.length + ' 行。';
    if (skipBlank() && blanks.length) {
      text += ' 另外排除 ' + colName + ' 行空白的 ' + skipped + ' 列' + (skipped < blanks.length ? '（保留 ' + (blanks.length - skipped) + ' 列）' : '') + '。';
    }
    if (result.rma) {
      text += ' 其中 RMA ' + result.rma.rows.toLocaleString('en-US') + ' 列，放在匯出檔第 ' + result.rma.titleRow + ' 列的「RMA」標題下面' +
        '（一般 MPS 到第 ' + (result.rma.titleRow - XIO.RMA_GAP) + ' 列）。';
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
      mainRows: result.mainRows,
      rmaRows: result.rma ? result.rma.rows : 0,
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
    if (byFuzzy()) {
      // 模糊比對：已經在輸入框裡就清掉，不在就填進第一個空的輸入框（沒有空的就加一個）
      var zs = settings.zValues;
      var at = -1, empty = -1;
      zs.forEach(function (x, i) {
        if (at < 0 && XIO.normalizeText(x) !== '' && XIO.matchKey(x) === v.key) at = i;
        if (empty < 0 && XIO.normalizeText(x) === '') empty = i;
      });
      if (at >= 0) zs[at] = '';
      else if (empty >= 0) zs[empty] = v.label;
      else zs.push(v.label);
      buildZInputs();
      recompute();
      return;
    }
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
