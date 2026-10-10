/*
 * excel-io.js — 所有 Excel 讀寫與排版都集中在這裡（使用 ExcelJS）。
 * 各步驟只呼叫 ExcelIO，不直接碰 ExcelJS，之後要調整匯出格式只需要改這個檔。
 *
 * 名詞：行 = 直的（A、B、C…），列 = 橫的（1、2、3…）。
 */
(function (global) {
  'use strict';

  /* ---------- 行號（字母）與數字互轉 ---------- */

  function colToNum(letters) {
    var n = 0;
    for (var i = 0; i < letters.length; i++) {
      n = n * 26 + (letters.charCodeAt(i) - 64);
    }
    return n;
  }

  function numToCol(n) {
    var s = '';
    while (n > 0) {
      var m = (n - 1) % 26;
      s = String.fromCharCode(65 + m) + s;
      n = Math.floor((n - 1) / 26);
    }
    return s;
  }

  // 全形轉半形、去頭尾空白（使用者常用全形輸入 Ａ、１）
  function normalizeText(s) {
    return String(s).normalize('NFKC').trim();
  }

  /*
   * 解析「B, C, U:Z, AE」這類輸入，回傳 { cols: [2,3,21,...], bad: ['看不懂的字'] }
   * 接受逗號、頓號、空白、分號分隔；範圍可用 : 或 - 或 ~
   */
  function parseColumns(text) {
    var cols = {};
    var bad = [];
    var tokens = normalizeText(text || '').toUpperCase().split(/[\s,、;]+/).filter(Boolean);
    tokens.forEach(function (tok) {
      var m = /^([A-Z]{1,3})(?:[:\-~]([A-Z]{1,3}))?$/.exec(tok);
      if (!m) { bad.push(tok); return; }
      var a = colToNum(m[1]);
      var b = m[2] ? colToNum(m[2]) : a;
      if (a > 16384 || b > 16384) { bad.push(tok); return; }
      if (a > b) { var t = a; a = b; b = t; }
      for (var c = a; c <= b; c++) cols[c] = true;
    });
    var list = Object.keys(cols).map(Number).sort(function (x, y) { return x - y; });
    return { cols: list, bad: bad };
  }

  function formatColumns(cols) {
    return cols.map(numToCol).join(', ');
  }

  /* ---------- 儲存格的值 ---------- */

  // 把 ExcelJS 的各種值型態（公式、RichText、超連結…）化成單純的值
  function plainValue(v) {
    if (v === null || v === undefined) return null;
    if (v instanceof Date) return v;
    if (typeof v !== 'object') return v;
    if (v.formula !== undefined || v.sharedFormula !== undefined) return plainValue(v.result);
    if (Array.isArray(v.richText)) {
      return v.richText.map(function (t) { return t.text; }).join('');
    }
    if (v.error !== undefined) return v.error;
    if (v.text !== undefined) return plainValue(v.text);
    if (v.result !== undefined) return plainValue(v.result);
    return String(v);
  }

  // 合併儲存格裡不是左上角的格子，值要去找左上角那一格
  function cellRawValue(cell) {
    if (!cell) return null;
    return cell.isMerged ? cell.master.value : cell.value;
  }

  // 比對用的鍵：1、"1"、"１"、"1.0" 都視為相同
  function matchKey(v) {
    v = plainValue(v);
    if (v === null || v === undefined) return '';
    if (v instanceof Date) return v.toISOString();
    if (typeof v === 'number') return String(v);
    var s = normalizeText(v);
    if (s !== '' && !isNaN(Number(s))) return String(Number(s));
    return s;
  }

  function pad2(n) { return (n < 10 ? '0' : '') + n; }

  // 預覽畫面顯示用的文字
  function displayText(cell) {
    var v = plainValue(cellRawValue(cell));
    if (v === null || v === undefined) return '';
    if (v instanceof Date) {
      var d = v.getUTCFullYear() + '/' + (v.getUTCMonth() + 1) + '/' + v.getUTCDate();
      if (v.getUTCHours() || v.getUTCMinutes()) {
        d += ' ' + pad2(v.getUTCHours()) + ':' + pad2(v.getUTCMinutes());
      }
      return d;
    }
    if (typeof v === 'number') {
      var fmt = (cell && cell.numFmt) || '';
      if (fmt.indexOf('%') >= 0) return (Math.round(v * 10000) / 100) + '%';
      var rounded = Math.round(v * 10000) / 10000;
      if (fmt.indexOf('#,##') >= 0) return rounded.toLocaleString('en-US');
      return String(rounded);
    }
    return String(v);
  }

  // 匯出時寫入的值：公式換成計算結果（篩掉列與行之後，原公式的參照會錯位）
  function exportValue(cell) {
    var v = cellRawValue(cell);
    if (v && typeof v === 'object' && !(v instanceof Date) &&
        (v.formula !== undefined || v.sharedFormula !== undefined)) {
      return v.result === undefined ? null : v.result;
    }
    return v === undefined ? null : v;
  }

  /* ---------- 讀檔 ---------- */

  function ioError(code, message) {
    var e = new Error(message);
    e.code = code;
    return e;
  }

  function load(arrayBuffer) {
    var head = new Uint8Array(arrayBuffer.slice(0, 4));
    if (head[0] === 0xD0 && head[1] === 0xCF) {
      return Promise.reject(ioError('OLD_XLS',
        '這個檔案是舊版 .xls 格式，或有密碼保護。請在 Excel 用「另存新檔」存成 .xlsx（並取消密碼）後再上傳。'));
    }
    if (!(head[0] === 0x50 && head[1] === 0x4B)) {
      return Promise.reject(ioError('NOT_XLSX',
        '這不是 Excel 活頁簿（.xlsx）。請確認上傳的是 MPS 報表的 .xlsx 檔。'));
    }
    var wb = new global.ExcelJS.Workbook();
    return wb.xlsx.load(arrayBuffer).then(function () {
      if (!wb.worksheets.length) throw ioError('NO_SHEET', '這個活頁簿裡沒有工作表。');
      return wb;
    }, function (err) {
      throw ioError('READ_FAIL', '讀不了這個 Excel 檔：' + (err && err.message ? err.message : err));
    });
  }

  /* ---------- 合併儲存格 ---------- */

  function listMerges(ws) {
    var out = [];
    var merges = ws._merges || {};
    Object.keys(merges).forEach(function (key) {
      var m = merges[key];
      if (m) out.push({ top: m.top, left: m.left, bottom: m.bottom, right: m.right });
    });
    return out;
  }

  /* ---------- 檢視工作表內容 ---------- */

  /*
   * 回傳：
   *   lastRow     最後一列有資料的列號
   *   colCount    最右邊有資料的行號
   *   dataRows    資料列數（從 dataStartRow 起）
   *   values      選取行出現過的值 [{ key, label, count }]，依出現順序
   *   blankCount  選取行是空白的列數
   */
  function inspect(ws, opts) {
    var dataStartRow = opts.dataStartRow;
    var selectorCol = opts.selectorCol;
    var lastRow = 0;
    var colCount = 0;
    var counts = {};
    var order = [];
    var nonBlank = 0;

    ws.eachRow({ includeEmpty: false }, function (row, r) {
      var rowHasValue = false;
      row.eachCell({ includeEmpty: false }, function (cell, c) {
        var v = cellRawValue(cell);
        if (v === null || v === undefined || v === '') return;
        rowHasValue = true;
        if (c > colCount) colCount = c;
      });
      if (!rowHasValue) return;
      if (r > lastRow) lastRow = r;
      if (r < dataStartRow) return;
      var cell = row.findCell(selectorCol);
      var key = matchKey(cellRawValue(cell));
      if (key === '') return;
      nonBlank++;
      if (!counts[key]) {
        counts[key] = { key: key, label: displayText(cell), count: 0 };
        order.push(key);
      }
      counts[key].count++;
    });

    var dataRows = Math.max(0, lastRow - dataStartRow + 1);
    return {
      lastRow: lastRow,
      colCount: colCount,
      dataRows: dataRows,
      values: order.map(function (k) { return counts[k]; }),
      blankCount: dataRows - nonBlank
    };
  }

  /* ---------- 篩選並產生新的活頁簿 ---------- */

  function cleanObject(o) {
    var out = {};
    Object.keys(o).forEach(function (k) { if (o[k] !== undefined) out[k] = o[k]; });
    return out;
  }

  /*
   * opts:
   *   dataStartRow  資料從第幾列開始（之前的列是表頭，全部保留）
   *   selectorCol   選取條件所在的行（數字，A=1）
   *   values        要選取的值（字串陣列）
   *   keepCols      要保留的行（數字陣列，已排序）
   *   lastRow       最後一列資料的列號
   *
   * 表頭列的樣式、合併儲存格、行寬、列高都會照原樣帶到新檔；
   * 合併儲存格如果只保留了其中幾行，會縮成只合併那幾行。
   */
  function filter(ws, opts) {
    var ExcelJS = global.ExcelJS;
    var dataStartRow = opts.dataStartRow;
    var headerEnd = dataStartRow - 1;
    var keepCols = opts.keepCols;
    var wanted = {};
    opts.values.forEach(function (v) { wanted[matchKey(v)] = true; });

    var newIndex = {}; // 原行號 -> 新行號
    keepCols.forEach(function (c, i) { newIndex[c] = i + 1; });

    var out = new ExcelJS.Workbook();
    out.creator = 'IE 人力需求小程式';
    out.created = new Date();
    var props = ws.properties || {};
    var ows = out.addWorksheet(ws.name || 'MPS', {
      properties: cleanObject({
        defaultRowHeight: props.defaultRowHeight,
        defaultColWidth: props.defaultColWidth
      })
    });

    // 行寬
    keepCols.forEach(function (c, i) {
      var w = ws.getColumn(c).width;
      if (w) ows.getColumn(i + 1).width = w;
    });

    // 表頭列：每一格先照抄（值只有左上角那格有，其餘合併格只抄樣式）
    var r, i, src, dst, srcRow, dstRow;
    for (r = 1; r <= headerEnd; r++) {
      srcRow = ws.findRow(r);
      if (!srcRow) continue;
      dstRow = ows.getRow(r);
      if (srcRow.height) dstRow.height = srcRow.height;
      for (i = 0; i < keepCols.length; i++) {
        src = srcRow.findCell(keepCols[i]);
        if (!src) continue;
        dst = dstRow.getCell(i + 1);
        dst.style = src.style;
        if (!src.isMerged || src.master === src) dst.value = exportValue(src);
      }
    }

    // 表頭的合併儲存格：對應到新的行號後重新合併
    listMerges(ws).forEach(function (m) {
      if (m.top > headerEnd) return;
      var bottom = Math.min(m.bottom, headerEnd);
      var kept = keepCols.filter(function (c) { return c >= m.left && c <= m.right; });
      if (!kept.length) return;
      var n1 = newIndex[kept[0]];
      var n2 = newIndex[kept[kept.length - 1]];
      var master = ws.getCell(m.top, m.left);
      var target = ows.getCell(m.top, n1);
      target.value = exportValue(master);
      if (kept[0] !== m.left) target.style = master.style;
      if (n2 > n1 || bottom > m.top) {
        ows.mergeCellsWithoutStyle(m.top, n1, bottom, n2);
      }
    });

    // 第 1 列的大標題如果落在沒保留的行，補到新檔的左上角，避免標題不見
    srcRow = headerEnd >= 1 ? ws.findRow(1) : null;
    if (srcRow) {
      var titleCell = null;
      srcRow.eachCell({ includeEmpty: false }, function (cell) {
        if (titleCell || (cell.isMerged && cell.master !== cell)) return;
        var tv = cell.value;
        if (tv !== null && tv !== undefined && tv !== '') titleCell = cell;
      });
      var titleKept = false;
      for (i = 1; i <= keepCols.length; i++) {
        var cv = cellRawValue(ows.getCell(1, i));
        if (cv !== null && cv !== undefined && cv !== '') { titleKept = true; break; }
      }
      if (titleCell && !titleKept) {
        var a1 = ows.getCell(1, 1);
        a1.value = exportValue(titleCell);
        a1.style = titleCell.style;
      }
    }

    // 資料列：選取行的值符合才保留
    var matched = 0;
    var total = 0;
    var outRow = dataStartRow;
    for (r = dataStartRow; r <= opts.lastRow; r++) {
      total++;
      srcRow = ws.findRow(r);
      if (!srcRow) continue;
      var key = matchKey(cellRawValue(srcRow.findCell(opts.selectorCol)));
      if (key === '' || !wanted[key]) continue;
      matched++;
      dstRow = ows.getRow(outRow++);
      if (srcRow.height) dstRow.height = srcRow.height;
      for (i = 0; i < keepCols.length; i++) {
        src = srcRow.findCell(keepCols[i]);
        if (!src) continue;
        dst = dstRow.getCell(i + 1);
        dst.style = src.style;
        var val = exportValue(src);
        if (val !== null) dst.value = val;
      }
    }

    // 凍結窗格、顯示比例
    var v = (ws.views && ws.views[0]) || {};
    var view = cleanObject({ showGridLines: v.showGridLines, zoomScale: v.zoomScale });
    if (v.state === 'frozen') {
      view.state = 'frozen';
      view.xSplit = keepCols.filter(function (c) { return c <= (v.xSplit || 0); }).length;
      view.ySplit = v.ySplit || 0;
    }
    if (Object.keys(view).length) ows.views = [view];

    // 列印設定：紙張方向照原檔，每頁重複表頭
    var ps = ws.pageSetup || {};
    Object.assign(ows.pageSetup, cleanObject({
      orientation: ps.orientation,
      paperSize: ps.paperSize,
      fitToPage: ps.fitToPage,
      fitToWidth: ps.fitToWidth,
      fitToHeight: ps.fitToHeight,
      printTitlesRow: headerEnd >= 1 ? '1:' + headerEnd : undefined
    }));

    return {
      workbook: out,
      worksheet: ows,
      matchedRows: matched,
      totalRows: total,
      headerEnd: headerEnd,
      lastRow: outRow - 1,
      colCount: keepCols.length
    };
  }

  function toBlob(workbook) {
    return workbook.xlsx.writeBuffer().then(function (buf) {
      return new Blob([buf], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
      });
    });
  }

  var ExcelIO = {
    colToNum: colToNum,
    numToCol: numToCol,
    normalizeText: normalizeText,
    parseColumns: parseColumns,
    formatColumns: formatColumns,
    plainValue: plainValue,
    matchKey: matchKey,
    displayText: displayText,
    load: load,
    listMerges: listMerges,
    inspect: inspect,
    filter: filter,
    toBlob: toBlob
  };

  global.ExcelIO = ExcelIO;
  if (typeof module !== 'undefined' && module.exports) module.exports = ExcelIO;
})(typeof window !== 'undefined' ? window : globalThis);
