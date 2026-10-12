/*
 * 第 9 步：總表
 *   在活頁簿裡新增「總表」工作表。目前先建立表格的樣式、格式、外框格線和位置，
 *   各格的內容（公式）之後再一項一項補上。
 *
 *   版面（B 到 J 行、第 2 到 24 列）：
 *     第 2 列      法定加班上限46hr/人（紅底白字）
 *     第 3 列      表頭：B、C 合併畫斜線；D 到 I 是六個月；J 是備註（深藍底白字）
 *     第 4 列      MPS (Production)
 *     第 5–14 列   需求工時、工作天數、工作時間、加班時間、綜合效率、出勤率
 *     第 15–18 列  需求人力（On line、Off line_直接、Off line_間接、小計）
 *     第 19–22 列  現有人力（早+夜）（同上，小計(on+off)）
 *     第 23 列     人力需求差異(現有-需求)
 *     第 24 列     Loading Rate%(含加班)
 *
 *   月份表頭用公式參照 MPS 工作表「需求工時」的月份（第 5 步輸入的），例如 ='MPS'!N3。
 *
 *   來源有兩種：
 *     a. 剛才第 8 步做好的結果（不用先匯出，直接沿用）
 *     b. 上傳之前第 8 步匯出的檔案
 *
 * 做完之後把結果放在 IEApp.state.withReport。
 */
(function (global) {
  'use strict';

  var IEApp = global.IEApp;
  var XIO = global.ExcelIO;
  var ui = IEApp.ui;
  var h = IEApp.h;

  var SHEET_NAME = XIO.REPORT_SHEET;
  var STEP1_KEY = 'ie-manpower.step1';
  var TITLE = '需求工時';
  var SUFFIX = '_總表';
  var EXTRA = ['P003', '人力', SHEET_NAME];

  var S = {
    upload: null,         // 上傳的第 8 步檔案 { name, buffer, sheets }
    preferUpload: false,
    work: null            // 拿來預覽的活頁簿 { src, loading, error, wb, main }
  };
  var els = {};

  function sameName(a, b) { return String(a).toLowerCase() === String(b).toLowerCase(); }
  function isExtra(name) { return EXTRA.some(function (n) { return sameName(n, name); }); }

  function activeSource() {
    var prev = IEApp.state.withManpower;
    if (S.upload && (S.preferUpload || !prev)) return { kind: 'file', key: S.upload, name: S.upload.name };
    if (prev) return { kind: 'prev', key: prev, name: prev.outName, prev: prev };
    return null;
  }

  function buildFrom(src) {
    return src.kind === 'prev' ? Promise.resolve(src.prev.build()) : XIO.load(src.key.buffer.slice(0));
  }

  function pickMain(wb, src) {
    var others = wb.worksheets.filter(function (w) { return !isExtra(w.name); });
    var wanted = src.kind === 'prev' ? src.prev.mainSheet : null;
    return others.filter(function (w) { return w.name === wanted; })[0] || others[0] || null;
  }

  // MPS 的資料從第幾列開始：沿用前面步驟的；上傳的檔案用第 1 步記住的設定（預設第 5 列）
  function dataStartRow(src) {
    if (src.kind === 'prev' && src.prev.dataStartRow) return src.prev.dataStartRow;
    var n = 0;
    try {
      var s1 = JSON.parse(localStorage.getItem(STEP1_KEY) || '{}') || {};
      n = parseInt(XIO.normalizeText(s1.startRow === undefined ? '' : s1.startRow), 10);
    } catch (e) { n = 0; }
    return n >= 1 ? n : 5;
  }

  // 月份表頭要參照的儲存格（MPS 工作表「需求工時」那一塊的月份）
  function monthRefs(main, src) {
    var start = dataStartRow(src);
    var keyCol = src.kind === 'prev' && src.prev.keyCol ? src.prev.keyCol : 1;
    var plan = XIO.planSummary(main, { dataStartRow: start, keyCol: keyCol, title: TITLE });
    if (!plan.found) return null;
    var headRow = start - 1 >= 3 ? 3 : start - 1;
    if (headRow < 1) return null;
    var refs = [];
    for (var i = 0; i < Math.min(plan.width, XIO.REPORT.months); i++) {
      var c = plan.startCol + i;
      var cell = main.findCell(headRow, c);
      refs.push({ sheet: main.name, col: c, row: headRow, text: cell ? XIO.displayText(cell) : '' });
    }
    return refs;
  }

  /* ---------- 畫面 ---------- */
  function render(panel) {
    els = {};

    els.srcInfo = h('p', { class: 'source-line', 'aria-live': 'polite' });
    els.drop = ui.fileDrop({
      id: 'f9-base',
      title: '上傳第 8 步匯出的 Excel',
      sub: '或把檔案拖進這裡。支援 .xlsx',
      onFile: openUpload
    });
    els.switchBtn = h('button', { type: 'button', class: 'btn-link', hidden: true, onclick: switchSource });
    els.backBtn = h('button', {
      type: 'button', class: 'btn-link', hidden: true, text: '回到第 8 步匯入人力報表',
      onclick: function () { IEApp.go('manpower'); }
    });
    els.srcMsg = h('p', { class: 'msg is-error', role: 'alert', hidden: true });

    els.summary = h('p', { class: 'summary', 'aria-live': 'polite' });
    els.notes = h('ul', { class: 'notes' });
    els.preview = h('div', { class: 'sheet-area' });
    els.exportBtn = h('button', { type: 'button', class: 'btn-primary', disabled: true, text: '匯出 Excel', onclick: onExport });
    els.exportNote = h('span', { class: 'export-note', 'aria-live': 'polite' });

    panel.appendChild(h('div', { class: 'step-head' },
      h('h1', { text: '總表' }),
      h('p', { text: '在同一個檔案裡新增「總表」工作表。這一版先建立表格的樣式、格式和外框格線，各格的內容之後再補上。' })));

    panel.appendChild(h('ol', { class: 'tasks' },
      h('li', { class: 'task' },
        h('h2', { text: '第 8 步的 Excel' }),
        els.srcInfo, els.drop.el,
        h('div', { class: 'link-row' }, els.switchBtn, els.backBtn),
        els.srcMsg),
      h('li', { class: 'task' },
        h('h2', { text: '總表' }),
        els.summary, els.notes, els.preview,
        h('div', { class: 'actions' }, els.exportBtn, els.exportNote))));

    refresh();
  }

  /* ---------- 來源 ---------- */
  function openUpload(file) {
    els.srcMsg.hidden = true;
    els.drop.setBusy(file.name);
    ui.readWorkbook(file)
      .then(function (res) {
        var names = res.workbook.worksheets.map(function (w) { return w.name; });
        if (!names.some(function (n) { return !isExtra(n); })) {
          throw new Error('這個檔案沒有 MPS 工作表。請上傳第 8 步匯出的檔案。');
        }
        S.upload = { name: file.name, buffer: res.buffer, sheets: names };
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

  // 先產生一份活頁簿，加上總表，給畫面預覽
  function ensureWork(src) {
    if (S.work && S.work.src === src.key) return;
    var work = { src: src.key, loading: true };
    S.work = work;
    buildFrom(src).then(function (wb) {
      if (S.work !== work) return;
      var main = pickMain(wb, src);
      if (!main) throw new Error('找不到 MPS 工作表。');
      var refs = monthRefs(main, src);
      work.refs = refs;
      work.report = XIO.addReportSheet(wb, { monthRefs: refs || [] });
      work.wb = wb;
      work.main = main;
      work.loading = false;
      refresh();
    }).catch(function (err) {
      if (S.work !== work) return;
      work.loading = false;
      work.error = err && err.message ? err.message : String(err);
      refresh();
    });
  }

  /* ---------- 更新畫面 ---------- */
  function refresh() {
    var src = activeSource();
    var prev = IEApp.state.withManpower;
    IEApp.state.withReport = null;
    els.exportBtn.disabled = true;
    els.notes.textContent = '';
    els.preview.textContent = '';

    if (!src) {
      els.srcInfo.textContent = '還沒有第 8 步的結果。回到第 8 步，或上傳之前匯出的檔案。';
      els.drop.setIdle();
    } else if (src.kind === 'prev') {
      els.srcInfo.textContent = '沿用第 8 步剛做好的結果：' + prev.outName + '。不用先匯出；要改用之前匯出的檔案，再從下面上傳。';
      if (S.upload) els.drop.setFile(S.upload.name, '已上傳，目前沒有使用。點這裡換一個檔案');
      else els.drop.setIdle();
    } else {
      els.srcInfo.textContent = '使用上傳的檔案，裡面的工作表：' + S.upload.sheets.join('、') + '。';
      els.drop.setFile(S.upload.name);
    }
    els.backBtn.hidden = !!src;
    els.switchBtn.hidden = !(S.upload && prev);
    if (src) els.switchBtn.textContent = src.kind === 'file' ? '改用第 8 步剛做好的結果' : '改用上傳的檔案';

    if (!src) { setSummary('先準備好第 8 步的 Excel。'); return; }
    ensureWork(src);
    if (S.work.loading) { setSummary('準備資料中…'); return; }
    if (S.work.error) { setSummary(S.work.error, true); return; }

    var work = S.work;
    var ws = work.report.worksheet;
    var names = work.wb.worksheets.map(function (w) { return w.name; });
    var replacing = src.kind === 'file' && S.upload.sheets.some(function (n) { return sameName(n, SHEET_NAME); });
    var text = '匯出的檔案有 ' + names.length + ' 個工作表：' + names.join('、') + '。「' + SHEET_NAME +
      '」目前只有表格的樣式和框線，內容之後再補。';
    if (replacing) text += ' 檔案裡原本的「' + SHEET_NAME + '」工作表會重新建立。';
    setSummary(text, false, replacing);

    if (work.refs && work.refs.length) {
      var r0 = work.refs[0];
      els.notes.appendChild(h('li', null, '月份表頭參照 ' + work.main.name + ' 工作表「需求工時」的月份，例如 D3 ＝ ' +
        "='" + r0.sheet + "'!" + XIO.numToCol(r0.col) + r0.row + '（' + work.refs.map(function (r) { return r.text || '空白'; }).join('、') + '）。'));
    } else {
      els.notes.appendChild(h('li', { class: 'is-warn' },
        work.main.name + ' 工作表裡找不到「需求工時」的月份，總表的月份表頭先留空白。'));
    }

    ui.renderSheet(els.preview, ws, {
      headerEnd: XIO.REPORT.headRow,
      lastRow: XIO.REPORT.lastRow,
      colCount: XIO.REPORT.noteCol,
      rows: XIO.REPORT.lastRow,
      caption: '總表的版面（顏色、框線在 Excel 裡才看得到）。'
    });

    els.exportBtn.disabled = false;
    var refs = work.refs || [];
    IEApp.state.withReport = {
      outName: outputName(src),
      mainSheet: work.main.name,
      reportSheet: SHEET_NAME,
      reportRows: work.report.rows,
      monthCol: work.report.monthCol,
      monthCount: work.report.months,
      build: function () { return buildWorkbook(src, refs); }
    };
  }

  function setSummary(text, isError, isWarn) {
    els.summary.textContent = text;
    els.summary.className = 'summary' + (isError ? ' is-error' : '') + (isWarn ? ' is-warn' : '');
  }

  /* ---------- 產生活頁簿 ---------- */
  function buildWorkbook(src, refs) {
    return buildFrom(src).then(function (wb) {
      XIO.addReportSheet(wb, { monthRefs: refs });
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
    var ready = IEApp.state.withReport;
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

  IEApp.registerStep({ id: 'report', render: render });
})(window);
