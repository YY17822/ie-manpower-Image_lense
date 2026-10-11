/*
 * 第 2 步：匯總
 *   算出 MPS 各月的預計總產量（Total）、前三大料號的合計（Top 3）、其餘的合計（Non-Top 3）。
 *   表格放在篩選後的 MPS 資料下面：最後一列資料之後空一列開始，共三列，對齊各月的行；
 *   文字標籤在第一個月左邊那一行。每格都是公式。
 *
 *   當月（第一個月）比較特別：數量要先加上 WIP。總產量、前三大、標示都是用「WIP + 當月產量」來算。
 *   以 WIP 在 D 行、當月在 E 行、資料在第 5 到 74 列為例：
 *
 *                 當月                                          其他月份
 *     Total       =SUM(D5:D74)+SUM(E5:E74)                     =SUM(F5:F74)
 *     Top 3       =SUMPRODUCT(LARGE(D5:D74+E5:E74,{1,2,3}))    =LARGE(F5:F74,1)+LARGE(F5:F74,2)+LARGE(F5:F74,3)
 *     Non-Top 3   =E76-E77                                     =F76-F77
 *
 *   各月前三大的儲存格用 Excel 的條件式格式標成淺橘色（數值是 0 的不算）。
 *   Total 那一列是淺綠色底（#B5E48C）；整塊外框粗線、裡面細格線。
 *
 *   WIP 和六個月各在哪一行由使用者指定，和第 5 步（計算需求工時）共用同一組設定，
 *   在其中一步改了，另一步也會跟著變，避免兩邊對不起來。
 *
 *   來源有兩種：
 *     a. 剛才第 1 步篩選好的結果（不用先匯出，直接沿用）
 *     b. 上傳之前第 1 步匯出的檔案
 *
 * 做完之後把結果放在 IEApp.state.mpsTotals，下一步直接拿來用。
 */
(function (global) {
  'use strict';

  var IEApp = global.IEApp;
  var XIO = global.ExcelIO;
  var ui = IEApp.ui;
  var h = IEApp.h;

  var STORE_KEY = 'ie-manpower.totals';
  var STEP1_KEY = 'ie-manpower.step1';
  var REF_SHEET = 'P003';
  var LABELS = { total: 'Total', top: 'Top 3', rest: 'Non-Top 3' };
  var SUFFIX = '_匯總';
  var TOP_N = 3;
  var KEY_COL = 1;              // 料號在第 1 行
  var WIP_DEFAULT = 4;          // 第 4 行是當月 WIP
  var MONTH_FIRST = 5;          // 第 5 行起是月份
  var MONTH_COUNT = 6;          // 當月 + 往後五個月

  var S = {
    upload: null,         // 上傳的第 1 步檔案 { name, buffer }
    preferUpload: false,
    work: null            // 拿來預覽、統計的活頁簿 { src, loading, error, wb, main, mainExt }
  };
  var settings = { startRow: '', wipCol: '', monthCols: '' };
  var els = {};
  var timer = null;

  function loadSettings() {
    var s = {};
    try { s = JSON.parse(localStorage.getItem(STORE_KEY) || '{}') || {}; } catch (e) { s = {}; }
    var shared = ui.columns.get();          // WIP、月份的行和第 5 步共用
    settings = {
      startRow: s.startRow === undefined ? '' : s.startRow,
      wipCol: shared.wipCol,
      monthCols: shared.monthCols
    };
  }
  function saveSettings() {
    try { localStorage.setItem(STORE_KEY, JSON.stringify({ startRow: settings.startRow })); } catch (e) { /* 存不了就算了 */ }
    ui.columns.set({ wipCol: settings.wipCol, monthCols: settings.monthCols });
  }

  function isRefName(name) { return String(name).toLowerCase() === REF_SHEET.toLowerCase(); }

  function activeSource() {
    var prev = IEApp.state.mpsFiltered;
    if (S.upload && (S.preferUpload || !prev)) return { kind: 'file', key: S.upload, name: S.upload.name };
    if (prev) return { kind: 'prev', key: prev, name: prev.outName, prev: prev };
    return null;
  }

  function buildFrom(src) {
    return src.kind === 'prev' ? Promise.resolve(src.prev.build()) : XIO.load(src.key.buffer.slice(0));
  }

  function pickMain(wb, src) {
    var others = wb.worksheets.filter(function (w) { return !isRefName(w.name); });
    var wanted = src.kind === 'prev' ? src.prev.sheetName : null;
    return others.filter(function (w) { return w.name === wanted; })[0] || others[0] || null;
  }

  // MPS 的資料從第幾列開始：接著第 1 步做下來的話沿用第 1 步的設定
  function chainStartRow() {
    var src = activeSource();
    return src && src.kind === 'prev' ? src.prev.headerEnd + 1 : 0;
  }
  function startRow() {
    var chain = chainStartRow();
    if (chain) return chain;
    var n = parseInt(XIO.normalizeText(settings.startRow), 10);
    if (n >= 1) return n;
    try {
      var s1 = JSON.parse(localStorage.getItem(STEP1_KEY) || '{}') || {};
      n = parseInt(XIO.normalizeText(s1.startRow === undefined ? '' : s1.startRow), 10);
    } catch (e) { n = 0; }
    return n >= 1 ? n : 5;
  }

  function colOf(letter, max) {
    var p = XIO.parseColumns(letter || '');
    return p.cols.length === 1 && !p.bad.length && p.cols[0] <= max ? p.cols[0] : 0;
  }

  /* ---------- 畫面 ---------- */
  function render(panel) {
    els = {};
    loadSettings();

    els.srcInfo = h('p', { class: 'source-line', 'aria-live': 'polite' });
    els.drop = ui.fileDrop({
      id: 'ft-base',
      title: '上傳第 1 步匯出的 Excel',
      sub: '或把檔案拖進這裡。支援 .xlsx',
      onFile: openUpload
    });
    els.switchBtn = h('button', { type: 'button', class: 'btn-link', hidden: true, onclick: switchSource });
    els.backBtn = h('button', {
      type: 'button', class: 'btn-link', hidden: true, text: '回到第 1 步篩選 MPS',
      onclick: function () { IEApp.go('mps-filter'); }
    });
    els.srcMsg = h('p', { class: 'msg is-error', role: 'alert', hidden: true });

    els.wipCol = h('select', { id: 'ft-wip', class: 'in-select' });
    els.monthCols = h('input', {
      id: 'ft-months', type: 'text', class: 'in in-select', autocomplete: 'off', spellcheck: 'false',
      placeholder: '例如 E:J'
    });
    els.mainPreview = h('div', { class: 'sheet-area' });
    els.startRow = h('input', { id: 'ft-start', type: 'text', class: 'in in-tiny', inputmode: 'numeric' });
    els.startNote = h('span');

    els.summary = h('p', { class: 'summary', 'aria-live': 'polite' });
    els.notes = h('ul', { class: 'notes' });
    els.result = h('div', { class: 'sheet-area' });
    els.tops = h('div', { class: 'sheet-area' });
    els.exportBtn = h('button', { type: 'button', class: 'btn-primary', disabled: true, text: '匯出 Excel', onclick: onExport });
    els.nextBtn = h('button', {
      type: 'button', class: 'btn-secondary', disabled: true, text: '下一步：匯入 P003',
      onclick: function () { IEApp.go('p003-import'); }
    });
    els.exportNote = h('span', { class: 'export-note', 'aria-live': 'polite' });

    panel.appendChild(h('div', { class: 'step-head' },
      h('h1', { text: '匯總' }),
      h('p', { text: '在篩選後的 MPS 下面加上各月的預計總產量（Total）、前三大料號的合計（Top 3），以及其餘的合計（Non-Top 3）。當月會先加上 WIP 再算。每格都是公式，各月前三大的儲存格會標成淺橘色。' })));

    panel.appendChild(h('ol', { class: 'tasks' },
      h('li', { class: 'task' },
        h('h2', { text: '第 1 步的 Excel' }),
        els.srcInfo, els.drop.el,
        h('div', { class: 'link-row' }, els.switchBtn, els.backBtn),
        els.srcMsg),
      h('li', { class: 'task' },
        h('h2', { text: 'WIP 和各月產量在哪幾行' }),
        h('div', { class: 'field-grid' },
          h('label', { for: 'ft-wip', text: '當月 WIP 在' }), els.wipCol,
          h('label', { for: 'ft-months', text: '各月產量在' }), els.monthCols),
        h('p', { class: 'hint', text: '各月產量的第一行是當月，會先加上 WIP 再算；連續的行可以寫成 E:J，也可以點下面表格最上方的行號來選。這裡的設定和第 5 步（計算需求工時）共用。' }),
        els.mainPreview,
        h('details', { class: 'advanced' },
          h('summary', { text: '報表版面不一樣時' }),
          h('div', { class: 'field-row' },
            h('label', { for: 'ft-start', text: 'MPS 的資料從第' }), els.startRow, els.startNote))),
      h('li', { class: 'task' },
        h('h2', { text: '匯出' }),
        els.summary, els.notes, els.result, els.tops,
        h('div', { class: 'actions' }, els.exportBtn, els.nextBtn, els.exportNote))));

    els.wipCol.addEventListener('change', function () {
      settings.wipCol = els.wipCol.value;
      saveSettings();
      els.exportNote.textContent = '';
      refresh();
    });
    els.monthCols.addEventListener('input', function () {
      settings.monthCols = els.monthCols.value;
      saveSettings();
      els.exportNote.textContent = '';
      clearTimeout(timer);
      timer = setTimeout(function () { refresh(true); }, 250);
    });
    els.startRow.addEventListener('change', function () {
      settings.startRow = els.startRow.value;
      saveSettings();
      refresh();
    });

    refresh();
  }

  /* ---------- 來源 ---------- */
  function openUpload(file) {
    els.srcMsg.hidden = true;
    els.drop.setBusy(file.name);
    ui.readWorkbook(file)
      .then(function (res) {
        var others = res.workbook.worksheets.filter(function (w) { return !isRefName(w.name); });
        if (!others.length) throw new Error('這個檔案沒有 MPS 工作表。請上傳第 1 步匯出的檔案。');
        S.upload = { name: file.name, buffer: res.buffer };
        S.preferUpload = true;
        els.exportNote.textContent = '';
        refresh();
      })
      .catch(function (err) {
        els.srcMsg.textContent = err && err.message ? err.message : String(err);
        els.srcMsg.hidden = false;
        refresh();
      });
  }

  function switchSource() {
    S.preferUpload = !S.preferUpload;
    els.exportNote.textContent = '';
    refresh();
  }

  function ensureWork(src) {
    if (S.work && S.work.src === src.key) return;
    var work = { src: src.key, loading: true };
    S.work = work;
    buildFrom(src).then(function (wb) {
      if (S.work !== work) return;
      var main = pickMain(wb, src);
      if (!main) throw new Error('找不到 MPS 工作表。');
      work.wb = wb;
      work.main = main;
      work.mainExt = XIO.extent(main);
      work.loading = false;
      refresh();
    }).catch(function (err) {
      if (S.work !== work) return;
      work.loading = false;
      work.error = err && err.message ? err.message : String(err);
      refresh();
    });
  }

  function fillSelect(select, labels, colCount, chosen) {
    select.textContent = '';
    select.appendChild(h('option', { value: '', text: '請選擇' }));
    for (var c = 1; c <= colCount; c++) {
      var letter = XIO.numToCol(c);
      var label = labels[c] ? letter + '　' + labels[c] : letter;
      if (label.length > 42) label = label.slice(0, 41) + '…';
      select.appendChild(h('option', { value: letter, text: label }));
    }
    select.value = chosen ? XIO.numToCol(chosen) : '';
  }

  // 連續的行寫成 E:J，其餘用逗號分開
  function compact(cols) {
    if (!cols.length) return '';
    var contiguous = cols.every(function (c, i) { return i === 0 || c === cols[i - 1] + 1; });
    if (contiguous && cols.length > 1) return XIO.numToCol(cols[0]) + ':' + XIO.numToCol(cols[cols.length - 1]);
    return XIO.formatColumns(cols);
  }

  function toggleMonth(c) {
    var text = XIO.normalizeText(settings.monthCols) ? settings.monthCols : els.monthCols.value;
    var cols = XIO.parseColumns(text).cols;
    var i = cols.indexOf(c);
    if (i >= 0) cols.splice(i, 1); else cols.push(c);
    cols.sort(function (a, b) { return a - b; });
    settings.monthCols = compact(cols);
    saveSettings();
    els.exportNote.textContent = '';
    var scroller = els.mainPreview.querySelector('.sheet-scroll');
    var left = scroller ? scroller.scrollLeft : 0;
    refresh();
    scroller = els.mainPreview.querySelector('.sheet-scroll');
    if (scroller) scroller.scrollLeft = left;
    var btn = els.mainPreview.querySelector('button[data-col="' + c + '"]');
    if (btn) btn.focus({ preventScroll: true });
  }

  /* ---------- 更新畫面 ---------- */
  // keepMonthText：使用者正在打字時不要改寫輸入框的內容
  function refresh(keepMonthText) {
    var src = activeSource();
    var prev = IEApp.state.mpsFiltered;
    IEApp.state.mpsTotals = null;
    els.exportBtn.disabled = true;
    els.nextBtn.disabled = true;
    els.notes.textContent = '';
    els.result.textContent = '';
    els.tops.textContent = '';

    if (!src) {
      els.srcInfo.textContent = '還沒有第 1 步的結果。回到第 1 步篩選 MPS，或上傳之前匯出的檔案。';
      els.drop.setIdle();
    } else if (src.kind === 'prev') {
      els.srcInfo.textContent = '沿用第 1 步剛篩選好的結果：' + prev.outName + '（' +
        prev.matchedRows.toLocaleString('en-US') + ' 列資料）。不用先匯出；要改用之前匯出的檔案，再從下面上傳。';
      if (S.upload) els.drop.setFile(S.upload.name, '已上傳，目前沒有使用。點這裡換一個檔案');
      else els.drop.setIdle();
    } else {
      els.srcInfo.textContent = '使用上傳的檔案。';
      els.drop.setFile(S.upload.name);
    }
    els.backBtn.hidden = !!src;
    els.switchBtn.hidden = !(S.upload && prev);
    if (src) els.switchBtn.textContent = src.kind === 'file' ? '改用第 1 步剛篩選好的結果' : '改用上傳的檔案';

    var chain = chainStartRow();
    els.startRow.value = String(startRow());
    els.startRow.disabled = !!chain;
    els.startNote.textContent = chain ? '列開始（沿用第 1 步的設定）' : '列開始';

    function clear() {
      els.wipCol.textContent = '';
      els.wipCol.appendChild(h('option', { value: '', text: '請選擇' }));
      els.mainPreview.textContent = '';
    }
    if (!src) { clear(); setSummary('先準備好第 1 步的 Excel。'); return; }

    ensureWork(src);
    var work = S.work;
    if (work.loading) { clear(); setSummary('準備資料中…'); return; }
    if (work.error) { clear(); setSummary(work.error, true); return; }

    var sr = startRow();
    var headerEnd = sr - 1;
    var cols = Math.max(work.mainExt.colCount, 1);
    var labels = XIO.columnLabels(work.main, headerEnd >= 2 ? 2 : 1, Math.max(headerEnd, 1), cols);

    // WIP、月份：先用記住的，沒有就用預設的位置（第 4 行、第 5 到 10 行）
    var wipCol = colOf(settings.wipCol, cols) || (WIP_DEFAULT <= cols ? WIP_DEFAULT : 0);
    var parsed = XIO.parseColumns(settings.monthCols);
    var monthCols = parsed.cols;
    if (!XIO.normalizeText(settings.monthCols)) {
      monthCols = [];
      for (var c = MONTH_FIRST; c < MONTH_FIRST + MONTH_COUNT && c <= cols; c++) monthCols.push(c);
    }

    fillSelect(els.wipCol, labels, cols, wipCol);
    if (!keepMonthText) els.monthCols.value = XIO.normalizeText(settings.monthCols) ? settings.monthCols : compact(monthCols);

    var keep = {};
    monthCols.forEach(function (mc) { keep[mc] = true; });
    ui.renderSheet(els.mainPreview, work.main, {
      headerEnd: headerEnd, lastRow: work.mainExt.lastRow, colCount: cols, rows: 5, keep: keep,
      onToggle: toggleMonth,
      caption: work.main.name + ' 工作表（前 5 列資料）。標起來的是各月產量。'
    });

    if (parsed.bad.length) {
      setSummary('看不懂「' + parsed.bad.join('、') + '」。行號請用英文字母，例如 E:J。', true);
      return;
    }
    if (!wipCol) { setSummary('選擇當月 WIP 在哪一行。'); return; }
    if (!monthCols.length) { setSummary('輸入各月產量在哪幾行。'); return; }
    var clash = monthCols.filter(function (mc) { return mc === wipCol || mc === KEY_COL || mc > cols; });
    if (clash.length) {
      setSummary(XIO.formatColumns(clash) + ' 行不能當作各月產量（和 WIP 或料號重複，或超出報表範圍）。', true);
      return;
    }

    var opts = {
      dataStartRow: sr, keyCol: KEY_COL, wipCol: wipCol, monthCols: monthCols, topN: TOP_N, labels: LABELS
    };
    var plan = XIO.planMpsTotals(work.main, opts);
    if (!plan.found) {
      setSummary('第 ' + sr + ' 列之後沒有料號。請確認上傳的檔案。', true);
      return;
    }

    var firstL = XIO.numToCol(monthCols[0]);
    var lastL = XIO.numToCol(monthCols[monthCols.length - 1]);
    var hasError = plan.months.some(function (m) { return !!m.error; });
    setSummary('匯總會填在 ' + work.main.name + ' 工作表的第 ' + plan.totalRow + ' 到 ' + plan.restRow +
      ' 列（最後一列資料是第 ' + plan.lastRow + ' 列，空一列），' + firstL + ' 到 ' + lastL + ' 行，共 ' +
      monthCols.length + ' 個月；文字標籤在 ' + XIO.numToCol(plan.labelCol) + ' 行。', false, hasError);

    if (hasError) {
      var bad = plan.months.filter(function (m) { return !!m.error; }).map(function (m) { return XIO.numToCol(m.col); });
      els.notes.appendChild(h('li', { class: 'is-warn' },
        bad.join('、') + ' 行有不是數字的格子，Excel 裡這幾個月的匯總會顯示錯誤。請先檢查 MPS 的數量。'));
    }
    if (monthCols.length !== MONTH_COUNT) {
      els.notes.appendChild(h('li', { class: 'is-warn' },
        '目前選了 ' + monthCols.length + ' 個月份，通常是 ' + MONTH_COUNT + ' 個（當月加往後五個月）。'));
    }

    renderResult(plan, labels, wipCol);
    renderTops(plan, labels);
    els.exportBtn.disabled = false;
    els.nextBtn.disabled = false;

    // 給下一步用（欄位和第 1 步的結果一樣，下一步可以直接接著做）
    var rows = plan.lastRow - plan.firstRow + 1;
    IEApp.state.mpsTotals = {
      sourceName: src.name,
      outName: outputName(src),
      sheetName: work.main.name,
      matchedRows: src.kind === 'prev' ? src.prev.matchedRows : rows,
      headerEnd: headerEnd,
      wipCol: wipCol,
      monthCols: monthCols,
      build: function () { return buildWorkbook(src, opts); }
    };
  }

  function setSummary(text, isError, isWarn) {
    els.summary.textContent = text;
    els.summary.className = 'summary' + (isError ? ' is-error' : '') + (isWarn ? ' is-warn' : '');
  }

  function showNumber(v) {
    return Math.round(v).toLocaleString('en-US');
  }

  function monthHead(m, labels, i) {
    var text = XIO.numToCol(m.col) + (labels[m.col] ? '　' + labels[m.col] : '') + (i === 0 ? '（加 WIP）' : '');
    return h('th', { class: 'colhead', scope: 'col', text: text });
  }

  // 三個匯總列算出來的數字，以及當月和下一個月填入的公式
  function renderResult(plan, labels, wipCol) {
    var head = h('tr', null,
      h('th', { class: 'colhead', scope: 'col', text: '列' }),
      h('th', { class: 'colhead', scope: 'col', text: '項目' }));
    plan.months.forEach(function (m, i) { head.appendChild(monthHead(m, labels, i)); });

    var body = h('tbody');
    [['total', plan.totalRow], ['top', plan.topRow], ['rest', plan.restRow]].forEach(function (pair) {
      var tr = h('tr', null,
        h('th', { class: 'rowhead', scope: 'row', text: String(pair[1]) }),
        h('td', { text: LABELS[pair[0]] }));
      plan.months.forEach(function (m) {
        var cls = 'is-num' + (pair[0] === 'total' ? ' is-total' : pair[0] === 'top' ? ' is-top' : '');
        tr.appendChild(h('td', { class: cls, text: m.error ? m.error : showNumber(m[pair[0]]) }));
      });
      body.appendChild(tr);
    });

    els.result.appendChild(h('p', { class: 'sheet-caption', text: '匯總結果。' }));
    els.result.appendChild(h('div', { class: 'sheet-scroll', tabindex: '0', role: 'group', 'aria-label': '匯總結果' },
      h('table', { class: 'sheet is-result' }, h('thead', null, head), body)));

    var W = XIO.numToCol(wipCol);
    var a = plan.firstRow, b = plan.lastRow;
    var n = Math.min(TOP_N, b - a + 1);
    var ks = [];
    for (var k = 1; k <= n; k++) ks.push(k);
    var L = XIO.numToCol(plan.months[0].col);
    var text = '當月（' + L + ' 行）的公式：' +
      L + plan.totalRow + ' =SUM(' + W + a + ':' + W + b + ')+SUM(' + L + a + ':' + L + b + ')　' +
      L + plan.topRow + ' =SUMPRODUCT(LARGE(' + W + a + ':' + W + b + '+' + L + a + ':' + L + b + ',{' + ks.join(',') + '}))　' +
      L + plan.restRow + ' =' + L + plan.totalRow + '-' + L + plan.topRow;
    els.result.appendChild(h('p', { class: 'hint', text: text }));
    if (plan.months.length > 1) {
      var M = XIO.numToCol(plan.months[1].col);
      var range = M + a + ':' + M + b;
      els.result.appendChild(h('p', { class: 'hint', text: '其他月份（以 ' + M + ' 行為例）：' +
        M + plan.totalRow + ' =SUM(' + range + ')　' +
        M + plan.topRow + ' =' + ks.map(function (x) { return 'LARGE(' + range + ',' + x + ')'; }).join('+') + '　' +
        M + plan.restRow + ' =' + M + plan.totalRow + '-' + M + plan.topRow }));
    }
  }

  // 每個月前三大是哪幾個料號（這些儲存格會標成淺橘色）
  function renderTops(plan, labels) {
    var head = h('tr', null, h('th', { class: 'colhead', scope: 'col', text: '名次' }));
    plan.months.forEach(function (m, i) { head.appendChild(monthHead(m, labels, i)); });
    var body = h('tbody');
    var depth = 0;
    plan.months.forEach(function (m) { depth = Math.max(depth, m.topRows.length); });
    for (var i = 0; i < depth; i++) {
      var tr = h('tr', null, h('th', { class: 'rowhead', scope: 'row', text: String(i + 1) }));
      plan.months.forEach(function (m) {
        var it = m.topRows[i];
        tr.appendChild(h('td', { class: 'is-top' },
          it ? h('span', { class: 'top-name', text: it.label }) : '',
          it ? h('span', { class: 'top-value', text: showNumber(it.value) + '（第 ' + it.row + ' 列）' }) : ''));
      });
      body.appendChild(tr);
    }
    els.tops.appendChild(h('p', { class: 'sheet-caption', text: '各月前三大的料號（數值是 0 的不算；當月是 WIP 加當月產量）。這些儲存格在 Excel 裡會標成淺橘色。' }));
    els.tops.appendChild(h('div', { class: 'sheet-scroll', tabindex: '0', role: 'group', 'aria-label': '各月前三大' },
      h('table', { class: 'sheet is-result is-tops' }, h('thead', null, head), body)));
  }

  /* ---------- 產生活頁簿 ---------- */
  function buildWorkbook(src, opts) {
    return buildFrom(src).then(function (wb) {
      XIO.addMpsTotals(pickMain(wb, src), opts);
      return wb;
    });
  }

  function outputName(src) {
    var stem = src.name.replace(/\.[^.]+$/, '');
    if (stem.slice(-SUFFIX.length) === SUFFIX) stem = stem.slice(0, -SUFFIX.length);
    return stem + SUFFIX + '.xlsx';
  }

  /* ---------- 匯出 ---------- */
  function onExport() {
    var ready = IEApp.state.mpsTotals;
    if (!ready) return;
    els.exportBtn.disabled = true;
    els.exportNote.textContent = '產生檔案中…';
    ready.build()
      .then(function (wb) { return XIO.toBlob(wb); })
      .then(function (blob) {
        ui.download(blob, ready.outName);
        els.exportNote.textContent = '已匯出 ' + ready.outName + '，在瀏覽器的下載資料夾。';
      })
      .catch(function (err) {
        els.exportNote.textContent = '匯出失敗：' + (err && err.message ? err.message : err);
      })
      .then(function () {
        els.exportBtn.disabled = false;
      });
  }

  IEApp.registerStep({ id: 'mps-totals', render: render });
})(window);
