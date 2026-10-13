/*
 * 第 10 步：RMA人力需求
 *   在活頁簿裡新增「RMA人力需求」工作表。目前先建立表格的樣式、格式、外框格線和位置，
 *   各格的內容之後再補上。
 *
 *   版面（B 到 J 行、第 3 到 13 列）：
 *     第 3 列      B:C 法定加班上限46hr/人（紅底白字置中）；J 是 Cut-off 日期和版次
 *     第 4 列      表頭：B、C 合併畫斜線；D 到 I 是六個月（參照總表的月份）；J 是備註（深藍底白字）
 *     第 5 列      MPS (PCS)
 *     第 6–10 列   需求工時、平日／假日工作天數、工作時間、加班時間（C 行是單位）
 *     第 11–12 列  綜合效率%、出勤率%
 *     第 13 列     需求人力 On line（淺綠底）
 *
 *   來源有兩種：
 *     a. 剛才第 9 步做好的結果（不用先匯出，直接沿用）
 *     b. 上傳之前第 9 步匯出的檔案
 *
 * 做完之後把結果放在 IEApp.state.withRmaManpower。
 */
(function (global) {
  'use strict';

  var IEApp = global.IEApp;
  var XIO = global.ExcelIO;
  var ui = IEApp.ui;
  var h = IEApp.h;

  var SHEET_NAME = XIO.RMA_SHEET;
  var REPORT_SHEET = XIO.REPORT_SHEET;
  var SUFFIX = '_RMA人力';
  var STORE_KEY = 'ie-manpower.rma-manpower';

  var S = {
    upload: null,         // 上傳的第 9 步檔案 { name, buffer, sheets }
    preferUpload: false,
    work: null,           // 拿來預覽的活頁簿 { src, loading, error, wb, refs }
    cutDate: today(),     // Cut-off 日期（每次打開預設今天）
    version: loadVersion()
  };
  var els = {};
  var timer = null;

  function today() {
    var d = new Date();
    return d.getFullYear() + '/' + (d.getMonth() + 1) + '/' + d.getDate();
  }
  function loadVersion() {
    try { return (JSON.parse(localStorage.getItem(STORE_KEY) || '{}') || {}).version || 'A'; } catch (e) { return 'A'; }
  }
  function saveVersion() {
    try { localStorage.setItem(STORE_KEY, JSON.stringify({ version: S.version })); } catch (e) { /* 存不了就算了 */ }
  }
  function cutoffText() {
    return 'Cut-off: ' + S.cutDate.trim() + 'Ver:' + S.version.trim();
  }

  function sameName(a, b) { return String(a).toLowerCase() === String(b).toLowerCase(); }

  function activeSource() {
    var prev = IEApp.state.withReport;
    if (S.upload && (S.preferUpload || !prev)) return { kind: 'file', key: S.upload, name: S.upload.name };
    if (prev) return { kind: 'prev', key: prev, name: prev.outName, prev: prev };
    return null;
  }

  function buildFrom(src) {
    return src.kind === 'prev' ? Promise.resolve(src.prev.build()) : XIO.load(src.key.buffer.slice(0));
  }

  // 月份表頭：參照總表第一張表的月份（總表 D3 到 I3）
  function monthRefs(wb) {
    var report = wb.worksheets.filter(function (w) { return sameName(w.name, REPORT_SHEET); })[0];
    if (!report) return null;
    var R = XIO.REPORT;
    var refs = [];
    for (var i = 0; i < R.months; i++) {
      var cell = report.findCell(R.headRow, R.monthCol + i);
      refs.push({ sheet: report.name, col: R.monthCol + i, row: R.headRow, text: cell ? XIO.displayText(cell) : '' });
    }
    return refs;
  }

  /* ---------- 畫面 ---------- */
  function render(panel) {
    els = {};

    els.srcInfo = h('p', { class: 'source-line', 'aria-live': 'polite' });
    els.drop = ui.fileDrop({
      id: 'f10-base',
      title: '上傳第 9 步匯出的 Excel',
      sub: '或把檔案拖進這裡。支援 .xlsx',
      onFile: openUpload
    });
    els.switchBtn = h('button', { type: 'button', class: 'btn-link', hidden: true, onclick: switchSource });
    els.backBtn = h('button', {
      type: 'button', class: 'btn-link', hidden: true, text: '回到第 9 步總表',
      onclick: function () { IEApp.go('report'); }
    });
    els.srcMsg = h('p', { class: 'msg is-error', role: 'alert', hidden: true });

    els.cutDate = h('input', { id: 'f10-cutoff', type: 'text', class: 'in in-short', autocomplete: 'off', value: S.cutDate });
    els.version = h('input', { id: 'f10-ver', type: 'text', class: 'in in-tiny', autocomplete: 'off', value: S.version });
    els.cutDate.addEventListener('input', function () { S.cutDate = els.cutDate.value; onType(); });
    els.version.addEventListener('input', function () { S.version = els.version.value; saveVersion(); onType(); });

    els.summary = h('p', { class: 'summary', 'aria-live': 'polite' });
    els.notes = h('ul', { class: 'notes' });
    els.preview = h('div', { class: 'sheet-area' });
    els.exportBtn = h('button', { type: 'button', class: 'btn-primary', disabled: true, text: '匯出 Excel', onclick: onExport });
    els.exportNote = h('span', { class: 'export-note', 'aria-live': 'polite' });

    panel.appendChild(h('div', { class: 'step-head' },
      h('h1', { text: 'RMA人力需求' }),
      h('p', { text: '在同一個檔案裡新增「RMA人力需求」工作表。這一版先建立表格的樣式、格式和外框格線，各格的內容之後再補上。' })));

    panel.appendChild(h('ol', { class: 'tasks' },
      h('li', { class: 'task' },
        h('h2', { text: '第 9 步的 Excel' }),
        els.srcInfo, els.drop.el,
        h('div', { class: 'link-row' }, els.switchBtn, els.backBtn),
        els.srcMsg),
      h('li', { class: 'task' },
        h('h2', { text: '右上角的 Cut-off' }),
        h('div', { class: 'field-row' },
          h('label', { for: 'f10-cutoff', text: 'Cut-off 日期' }), els.cutDate,
          h('label', { for: 'f10-ver', text: '版次 Ver' }), els.version),
        h('p', { class: 'hint', text: '日期預設今天，版次會記住上次輸入的值。' })),
      h('li', { class: 'task' },
        h('h2', { text: 'RMA人力需求' }),
        els.summary, els.notes, els.preview,
        h('div', { class: 'actions' }, els.exportBtn, els.exportNote))));

    refresh();
  }

  function onType() {
    els.exportNote.textContent = '';
    clearTimeout(timer);
    timer = setTimeout(refresh, 250);
  }

  /* ---------- 來源 ---------- */
  function openUpload(file) {
    els.srcMsg.hidden = true;
    els.drop.setBusy(file.name);
    ui.readWorkbook(file)
      .then(function (res) {
        var names = res.workbook.worksheets.map(function (w) { return w.name; });
        if (!names.some(function (n) { return sameName(n, REPORT_SHEET); })) {
          throw new Error('這個檔案沒有「' + REPORT_SHEET + '」工作表。請上傳第 9 步匯出的檔案。');
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

  function ensureWork(src) {
    if (S.work && S.work.src === src.key) return;
    var work = { src: src.key, loading: true };
    S.work = work;
    buildFrom(src).then(function (wb) {
      if (S.work !== work) return;
      work.wb = wb;
      work.refs = monthRefs(wb);
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
    var prev = IEApp.state.withReport;
    IEApp.state.withRmaManpower = null;
    els.exportBtn.disabled = true;
    els.notes.textContent = '';
    els.preview.textContent = '';

    if (!src) {
      els.srcInfo.textContent = '還沒有第 9 步的結果。回到第 9 步，或上傳之前匯出的檔案。';
      els.drop.setIdle();
    } else if (src.kind === 'prev') {
      els.srcInfo.textContent = '沿用第 9 步剛做好的結果：' + prev.outName + '。不用先匯出；要改用之前匯出的檔案，再從下面上傳。';
      if (S.upload) els.drop.setFile(S.upload.name, '已上傳，目前沒有使用。點這裡換一個檔案');
      else els.drop.setIdle();
    } else {
      els.srcInfo.textContent = '使用上傳的檔案，裡面的工作表：' + S.upload.sheets.join('、') + '。';
      els.drop.setFile(S.upload.name);
    }
    els.backBtn.hidden = !!src;
    els.switchBtn.hidden = !(S.upload && prev);
    if (src) els.switchBtn.textContent = src.kind === 'file' ? '改用第 9 步剛做好的結果' : '改用上傳的檔案';

    if (!src) { setSummary('先準備好第 9 步的 Excel。'); return; }
    ensureWork(src);
    if (S.work.loading) { setSummary('準備資料中…'); return; }
    if (S.work.error) { setSummary(S.work.error, true); return; }

    var work = S.work;
    var cutoff = cutoffText();
    var report = XIO.addRmaManpowerSheet(work.wb, { cutoff: cutoff, monthRefs: work.refs || [] });
    var names = work.wb.worksheets.map(function (w) { return w.name; });
    var replacing = src.kind === 'file' && S.upload.sheets.some(function (n) { return sameName(n, SHEET_NAME); });
    var text = '匯出的檔案有 ' + names.length + ' 個工作表：' + names.join('、') + '。「' + SHEET_NAME +
      '」目前只有表格的樣式和框線，內容之後再補。';
    if (replacing) text += ' 檔案裡原本的「' + SHEET_NAME + '」工作表會重新建立。';
    setSummary(text, false, replacing);

    if (work.refs) {
      els.notes.appendChild(h('li', null, '月份表頭參照總表的月份，例如 D4 ＝ =\'' + REPORT_SHEET + "'!D3（" +
        work.refs.map(function (r) { return r.text || '空白'; }).join('、') + '）。'));
    } else {
      els.notes.appendChild(h('li', { class: 'is-warn' }, '檔案裡沒有「' + REPORT_SHEET + '」工作表，月份表頭先留空白。'));
    }
    els.notes.appendChild(h('li', null, '右上角寫「' + cutoff + '」。'));

    ui.renderSheet(els.preview, report.worksheet, {
      headerEnd: XIO.RMA_LAYOUT.headRow,
      lastRow: XIO.RMA_LAYOUT.lastRow,
      colCount: XIO.RMA_LAYOUT.noteCol,
      rows: XIO.RMA_LAYOUT.lastRow,
      caption: 'RMA人力需求的版面（顏色、框線在 Excel 裡才看得到）。'
    });

    els.exportBtn.disabled = false;
    var refs = work.refs || [];
    IEApp.state.withRmaManpower = {
      outName: outputName(src),
      sheet: SHEET_NAME,
      rows: report.rows,
      build: function () { return buildWorkbook(src, refs, cutoff); }
    };
  }

  function setSummary(text, isError, isWarn) {
    els.summary.textContent = text;
    els.summary.className = 'summary' + (isError ? ' is-error' : '') + (isWarn ? ' is-warn' : '');
  }

  /* ---------- 產生活頁簿 ---------- */
  function buildWorkbook(src, refs, cutoff) {
    return buildFrom(src).then(function (wb) {
      XIO.addRmaManpowerSheet(wb, { cutoff: cutoff, monthRefs: refs });
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
    var ready = IEApp.state.withRmaManpower;
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

  IEApp.registerStep({ id: 'rma-manpower', render: render });
})(window);
