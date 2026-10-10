/*
 * 第 1 步：篩選 MPS 報表
 *   1. 上傳 MPS 的 Excel
 *   2. 輸入「A 行等於多少」——符合的列才留下（從第 5 列開始比對）
 *   3. 輸入要保留哪些行
 *   4. 匯出新的 Excel，第 1–4 列表頭（含合併儲存格）原樣保留
 *
 * 名詞：行 = 直的（A、B、C…），列 = 橫的（1、2、3…）。
 */
(function (global) {
  'use strict';

  var IEApp = global.IEApp;
  var XIO = global.ExcelIO;
  var h = IEApp.h;

  var STORE_KEY = 'ie-manpower.step1';
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
    try { s = JSON.parse(localStorage.getItem(STORE_KEY) || '{}') || {}; } catch (e) { s = {}; }
    var out = {};
    Object.keys(DEFAULTS).forEach(function (k) { out[k] = s[k] === undefined ? DEFAULTS[k] : s[k]; });
    return out;
  }
  function saveSettings() {
    try { localStorage.setItem(STORE_KEY, JSON.stringify(settings)); } catch (e) { /* 存不了就算了 */ }
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

  /* ---------- 畫面 ---------- */
  function render(panel) {
    els = {};

    els.fileInput = h('input', { type: 'file', accept: '.xlsx,.xlsm', class: 'visually-hidden', id: 'f-file' });
    els.dropMain = h('span', { class: 'drop-main', text: '選擇 MPS 的 Excel 檔' });
    els.dropSub = h('span', { class: 'drop-sub', text: '或把檔案拖進這裡。支援 .xlsx' });
    els.drop = h('label', { class: 'drop', for: 'f-file' }, els.fileInput, els.dropMain, els.dropSub);
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
    els.exportNote = h('span', { class: 'export-note', 'aria-live': 'polite' });

    panel.appendChild(h('div', { class: 'step-head' },
      h('h1', { text: '篩選 MPS 報表' }),
      h('p', { text: '留下 A 行等於指定數值的列，以及你要的行，匯出成新的 Excel。表頭（第 1–4 列）會原樣保留。' })));

    panel.appendChild(h('ol', { class: 'tasks' },
      h('li', { class: 'task' },
        h('h2', { text: '上傳 MPS 報表' }),
        els.drop, els.sheetRow, els.fileMsg),
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
        h('div', { class: 'actions' }, els.exportBtn, els.exportNote))));

    // 事件
    els.fileInput.addEventListener('change', function () {
      if (els.fileInput.files[0]) openFile(els.fileInput.files[0]);
      els.fileInput.value = '';
    });
    ['dragenter', 'dragover'].forEach(function (ev) {
      els.drop.addEventListener(ev, function (e) { e.preventDefault(); els.drop.classList.add('is-over'); });
    });
    ['dragleave', 'drop'].forEach(function (ev) {
      els.drop.addEventListener(ev, function (e) { e.preventDefault(); els.drop.classList.remove('is-over'); });
    });
    els.drop.addEventListener('drop', function (e) {
      var f = e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files[0];
      if (f) openFile(f);
    });
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
    els.dropMain.textContent = '讀取中…';
    els.dropSub.textContent = file.name;
    file.arrayBuffer()
      .then(function (buf) {
        return IEApp.loadExcelJS().then(function () { return XIO.load(buf); });
      })
      .then(function (wb) {
        S.workbook = wb;
        S.fileName = file.name;
        S.ws = null;
        showFile();
        pickSheet(0);
      })
      .catch(function (err) {
        els.fileMsg.textContent = err && err.message ? err.message : String(err);
        els.fileMsg.hidden = false;
        if (S.workbook) showFile();
        else {
          els.dropMain.textContent = '選擇 MPS 的 Excel 檔';
          els.dropSub.textContent = '或把檔案拖進這裡。支援 .xlsx';
        }
      });
  }

  function showFile() {
    els.drop.classList.add('has-file');
    els.dropMain.textContent = S.fileName;
    els.dropSub.textContent = '點這裡換一個檔案';
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
    els.exportBtn.disabled = true;
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
    renderSheet(els.srcPreview, S.ws, {
      headerEnd: sr - 1,
      lastRow: info.lastRow,
      colCount: Math.max(info.colCount, 1),
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

    var result = XIO.filter(S.ws, {
      dataStartRow: sr, selectorCol: sc, values: values, keepCols: parsed.cols, lastRow: info.lastRow
    });

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

    renderSheet(els.outPreview, result.worksheet, {
      headerEnd: result.headerEnd,
      lastRow: result.lastRow,
      colCount: result.colCount,
      caption: '匯出後的樣子（前 ' + PREVIEW_ROWS + ' 列資料）。'
    });
    els.exportBtn.disabled = false;

    // 給下一步用
    IEApp.state.mpsFiltered = {
      fileName: S.fileName, sheetName: S.ws.name, values: values,
      keepCols: parsed.cols, headerEnd: result.headerEnd, worksheet: result.worksheet
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

  /* ---------- 試算表預覽 ---------- */
  /*
   * o.headerEnd   表頭最後一列
   * o.lastRow     最後一列資料
   * o.colCount    要畫幾行
   * o.keep        { 行號: true }，要標示的行（可省略）
   * o.onToggle    點行號時呼叫（可省略；有給的話行號會變成按鈕）
   * o.selectorCol / o.wanted  用來把不符合條件的列變淡（可省略）
   */
  function renderSheet(container, ws, o) {
    container.textContent = '';
    var lastShown = Math.min(o.lastRow, o.headerEnd + PREVIEW_ROWS);
    var colCount = o.colCount;
    var keep = o.keep || {};

    // 合併儲存格 -> colspan / rowspan
    var span = {};
    var covered = {};
    XIO.listMerges(ws).forEach(function (m) {
      if (m.top > lastShown || m.left > colCount) return;
      var bottom = Math.min(m.bottom, lastShown);
      var right = Math.min(m.right, colCount);
      span[m.top + ':' + m.left] = { rows: bottom - m.top + 1, cols: right - m.left + 1 };
      for (var r = m.top; r <= bottom; r++) {
        for (var c = m.left; c <= right; c++) {
          if (r !== m.top || c !== m.left) covered[r + ':' + c] = true;
        }
      }
    });

    var headRow = h('tr', null, h('th', { class: 'corner', scope: 'col' }));
    for (var c = 1; c <= colCount; c++) {
      var letter = XIO.numToCol(c);
      var on = !!keep[c];
      var label = o.onToggle
        ? h('button', {
            type: 'button', 'data-col': String(c), 'aria-pressed': on ? 'true' : 'false',
            title: (on ? '取消保留 ' : '保留 ') + letter + ' 行',
            onclick: (function (col) { return function () { o.onToggle(col); }; })(c)
          }, letter)
        : letter;
      headRow.appendChild(h('th', { class: 'colhead' + (on ? ' is-kept' : ''), scope: 'col' }, label));
    }

    var body = h('tbody');
    for (var r = 1; r <= lastShown; r++) {
      var row = ws.findRow(r);
      var isHeader = r <= o.headerEnd;
      var out = false;
      if (!isHeader && o.wanted) {
        var k = XIO.matchKey(valueOf(row && row.findCell(o.selectorCol)));
        out = !(k !== '' && o.wanted[k]);
      }
      var tr = h('tr', { class: (isHeader ? 'is-header' : '') + (out ? ' is-out' : '') },
        h('th', { class: 'rowhead', scope: 'row', text: String(r) }));
      for (c = 1; c <= colCount; c++) {
        if (covered[r + ':' + c]) continue;
        var cell = row ? row.findCell(c) : null;
        var sp = span[r + ':' + c];
        // 合併儲存格：全部的行都保留才整格畫螢光，只保留一部分就用淡色
        var width = sp ? sp.cols : 1;
        var keptCount = 0;
        for (var cc = c; cc < c + width; cc++) if (keep[cc]) keptCount++;
        var raw = XIO.plainValue(valueOf(cell));
        var cls = (keptCount === width ? 'is-kept' : keptCount ? 'is-part' : '') +
          (typeof raw === 'number' ? ' is-num' : '') + (sp ? ' is-merged' : '') + (width > 6 ? ' is-wide' : '');
        var td = h('td', { class: cls.trim() || null, text: cell ? XIO.displayText(cell) : '' });
        if (sp && sp.cols > 1) td.colSpan = sp.cols;
        if (sp && sp.rows > 1) td.rowSpan = sp.rows;
        tr.appendChild(td);
      }
      body.appendChild(tr);
    }

    var more = o.lastRow - lastShown;
    container.appendChild(h('p', { class: 'sheet-caption', text: o.caption + (more > 0 ? ' 還有 ' + more.toLocaleString('en-US') + ' 列沒顯示。' : '') }));
    container.appendChild(h('div', { class: 'sheet-scroll', tabindex: '0', role: 'group', 'aria-label': '試算表預覽' },
      h('table', { class: 'sheet' }, h('thead', null, headRow), body)));
  }

  function valueOf(cell) {
    if (!cell) return null;
    return cell.isMerged ? cell.master.value : cell.value;
  }

  /* ---------- 匯出 ---------- */
  function onExport() {
    if (!S.result || !S.ws) return;
    var info = S.info;
    var parsed = XIO.parseColumns(settings.cols);
    els.exportBtn.disabled = true;
    els.exportNote.textContent = '產生檔案中…';
    // 重新產生一份乾淨的活頁簿再寫出
    var fresh = XIO.filter(S.ws, {
      dataStartRow: startRow(), selectorCol: selCol(), values: wantedValues(),
      keepCols: parsed.cols, lastRow: info.lastRow
    });
    var name = S.fileName.replace(/\.[^.]+$/, '') + '_篩選.xlsx';
    XIO.toBlob(fresh.workbook).then(function (blob) {
      var url = URL.createObjectURL(blob);
      var a = h('a', { href: url, download: name });
      document.body.appendChild(a);
      a.click();
      a.remove();
      setTimeout(function () { URL.revokeObjectURL(url); }, 10000);
      els.exportNote.textContent = '已匯出 ' + name + '（' + fresh.matchedRows.toLocaleString('en-US') + ' 列），在瀏覽器的下載資料夾。';
    }).catch(function (err) {
      els.exportNote.textContent = '匯出失敗：' + (err && err.message ? err.message : err);
    }).then(function () {
      els.exportBtn.disabled = false;
    });
  }

  IEApp.registerStep({ id: 'mps-filter', render: render });
})(window);
