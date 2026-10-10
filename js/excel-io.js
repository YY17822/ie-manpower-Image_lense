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

  /*
   * 讀檔前的修正：字型旗標。
   *
   * ExcelJS 4.x 讀字型設定時，只要看到 <b>、<i>、<strike> 這些標籤就當成「有」，
   * 不管後面是不是寫著 val="0"（＝沒有）。SAP 這類系統匯出的報表會把「沒有粗體、
   * 沒有刪除線」明寫成 <b val="0"/> <strike val="0"/>，結果整張表被讀成
   * 粗體＋斜體＋刪除線，匯出後就全部被畫上橫線。
   *
   * 做法：.xlsx 其實是一個 zip 壓縮檔。交給 ExcelJS 之前，先把裡面的樣式檔
   * （styles.xml）和文字檔（sharedStrings.xml）解開，拿掉這些「明寫為沒有」的標籤，
   * 再包回去。解壓縮用下面自己寫的 inflateRaw()，不靠瀏覽器的功能，所以 Chrome、
   * Edge、Safari 新舊版本的結果都一樣。任何一步不順利就原封不動交給 ExcelJS。
   */
  var OFF_FLAGS = /<(b|i|strike|outline|shadow|condense|extend)\s+val\s*=\s*(["'])(?:0|false)\2\s*(?:\/>|>\s*<\/\1>)/g;
  var FIX_ENTRIES = { 'xl/styles.xml': true, 'xl/sharedStrings.xml': true };

  /* 解開 zip 裡用 deflate 壓縮的資料。src：壓縮資料；size：解開後的大小。 */
  var LEN_BASE = [3, 4, 5, 6, 7, 8, 9, 10, 11, 13, 15, 17, 19, 23, 27, 31, 35, 43, 51, 59, 67, 83, 99, 115, 131, 163, 195, 227, 258];
  var LEN_EXTRA = [0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 2, 2, 2, 2, 3, 3, 3, 3, 4, 4, 4, 4, 5, 5, 5, 5, 0];
  var DIST_BASE = [1, 2, 3, 4, 5, 7, 9, 13, 17, 25, 33, 49, 65, 97, 129, 193, 257, 385, 513, 769, 1025, 1537, 2049, 3073, 4097, 6145, 8193, 12289, 16385, 24577];
  var DIST_EXTRA = [0, 0, 0, 0, 1, 1, 2, 2, 3, 3, 4, 4, 5, 5, 6, 6, 7, 7, 8, 8, 9, 9, 10, 10, 11, 11, 12, 12, 13, 13];
  var CODE_ORDER = [16, 17, 18, 0, 8, 7, 9, 6, 10, 5, 11, 4, 12, 3, 13, 2, 14, 1, 15];

  function huffman(lengths, start, n) {
    var count = new Uint16Array(16);
    var symbol = new Uint16Array(n);
    var offs = new Uint16Array(16);
    var i;
    for (i = 0; i < n; i++) count[lengths[start + i]]++;
    for (i = 1; i < 15; i++) offs[i + 1] = offs[i] + count[i];
    for (i = 0; i < n; i++) {
      var len = lengths[start + i];
      if (len) symbol[offs[len]++] = i;
    }
    return { count: count, symbol: symbol };
  }

  function inflateRaw(src, size) {
    var out = new Uint8Array(size);
    var op = 0, ip = 0, bitBuf = 0, bitCnt = 0;

    function bits(n) {
      while (bitCnt < n) {
        if (ip >= src.length) throw new Error('inflate: 資料不完整');
        bitBuf |= src[ip++] << bitCnt;
        bitCnt += 8;
      }
      var v = bitBuf & ((1 << n) - 1);
      bitBuf >>>= n;
      bitCnt -= n;
      return v;
    }
    function decode(h) {
      var code = 0, first = 0, index = 0;
      for (var len = 1; len <= 15; len++) {
        code |= bits(1);
        var cnt = h.count[len];
        if (code - cnt < first) return h.symbol[index + (code - first)];
        index += cnt;
        first = (first + cnt) << 1;
        code <<= 1;
      }
      throw new Error('inflate: 編碼錯誤');
    }
    function block(lenCode, distCode) {
      for (;;) {
        var sym = decode(lenCode);
        if (sym < 256) {
          if (op >= size) throw new Error('inflate: 超出大小');
          out[op++] = sym;
        } else if (sym === 256) {
          return;
        } else {
          sym -= 257;
          if (sym >= 29) throw new Error('inflate: 長度錯誤');
          var len = LEN_BASE[sym] + bits(LEN_EXTRA[sym]);
          var ds = decode(distCode);
          if (ds >= 30) throw new Error('inflate: 距離錯誤');
          var dist = DIST_BASE[ds] + bits(DIST_EXTRA[ds]);
          if (dist > op || op + len > size) throw new Error('inflate: 距離錯誤');
          for (var k = 0; k < len; k++, op++) out[op] = out[op - dist];
        }
      }
    }

    var last, lengths, i;
    do {
      last = bits(1);
      var type = bits(2);
      if (type === 0) {                      // 沒壓縮的區塊
        bitBuf = 0; bitCnt = 0;
        if (ip + 4 > src.length) throw new Error('inflate: 資料不完整');
        var n = src[ip] | (src[ip + 1] << 8);
        ip += 4;
        if (ip + n > src.length || op + n > size) throw new Error('inflate: 資料不完整');
        out.set(src.subarray(ip, ip + n), op);
        ip += n; op += n;
      } else if (type === 1) {               // 固定編碼
        lengths = new Uint8Array(288 + 30);
        for (i = 0; i < 144; i++) lengths[i] = 8;
        for (; i < 256; i++) lengths[i] = 9;
        for (; i < 280; i++) lengths[i] = 7;
        for (; i < 288; i++) lengths[i] = 8;
        for (; i < 318; i++) lengths[i] = 5;
        block(huffman(lengths, 0, 288), huffman(lengths, 288, 30));
      } else if (type === 2) {               // 動態編碼
        var nlen = bits(5) + 257, ndist = bits(5) + 1, ncode = bits(4) + 4;
        if (nlen > 286 || ndist > 30) throw new Error('inflate: 編碼表錯誤');
        lengths = new Uint8Array(320);
        for (i = 0; i < ncode; i++) lengths[CODE_ORDER[i]] = bits(3);
        var codeLen = huffman(lengths, 0, 19);
        lengths = new Uint8Array(320);
        i = 0;
        while (i < nlen + ndist) {
          var sym = decode(codeLen);
          if (sym < 16) {
            lengths[i++] = sym;
          } else {
            var prev = 0, rep;
            if (sym === 16) {
              if (i === 0) throw new Error('inflate: 編碼表錯誤');
              prev = lengths[i - 1];
              rep = 3 + bits(2);
            } else if (sym === 17) {
              rep = 3 + bits(3);
            } else {
              rep = 11 + bits(7);
            }
            if (i + rep > nlen + ndist) throw new Error('inflate: 編碼表錯誤');
            while (rep--) lengths[i++] = prev;
          }
        }
        block(huffman(lengths, 0, nlen), huffman(lengths, nlen, ndist));
      } else {
        throw new Error('inflate: 區塊類型錯誤');
      }
    } while (!last);
    if (op !== size) throw new Error('inflate: 大小不符');
    return out;
  }

  var crcTable = null;
  function crc32(bytes) {
    if (!crcTable) {
      crcTable = new Uint32Array(256);
      for (var n = 0; n < 256; n++) {
        var c = n;
        for (var k = 0; k < 8; k++) c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1);
        crcTable[n] = c >>> 0;
      }
    }
    var crc = 0xFFFFFFFF;
    for (var i = 0; i < bytes.length; i++) crc = crcTable[(crc ^ bytes[i]) & 0xFF] ^ (crc >>> 8);
    return (crc ^ 0xFFFFFFFF) >>> 0;
  }

  // 讀出 zip 裡每個檔案的資訊與（仍是壓縮狀態的）內容。看不懂的格式回傳 null。
  function readZip(buffer) {
    var dv = new DataView(buffer);
    var size = buffer.byteLength;
    var eocd = -1;
    for (var i = size - 22; i >= 0 && i >= size - 22 - 65535; i--) {
      if (dv.getUint32(i, true) === 0x06054b50) { eocd = i; break; }
    }
    if (eocd < 0) return null;
    var count = dv.getUint16(eocd + 10, true);
    var p = dv.getUint32(eocd + 16, true);
    if (count === 0xFFFF || p === 0xFFFFFFFF) return null;       // zip64，不處理
    var decoder = new TextDecoder();
    var entries = [];
    for (var k = 0; k < count; k++) {
      if (dv.getUint32(p, true) !== 0x02014b50) return null;
      var e = {
        flags: dv.getUint16(p + 8, true),
        method: dv.getUint16(p + 10, true),
        time: dv.getUint16(p + 12, true),
        date: dv.getUint16(p + 14, true),
        crc: dv.getUint32(p + 16, true),
        csize: dv.getUint32(p + 20, true),
        usize: dv.getUint32(p + 24, true),
        intAttr: dv.getUint16(p + 36, true),
        extAttr: dv.getUint32(p + 38, true)
      };
      var nameLen = dv.getUint16(p + 28, true);
      var extraLen = dv.getUint16(p + 30, true);
      var commentLen = dv.getUint16(p + 32, true);
      var local = dv.getUint32(p + 42, true);
      if (e.flags & 1) return null;                              // 有加密，不處理
      if (e.csize === 0xFFFFFFFF || e.usize === 0xFFFFFFFF || local === 0xFFFFFFFF) return null;
      e.nameBytes = new Uint8Array(buffer, p + 46, nameLen);
      e.name = decoder.decode(e.nameBytes);
      if (dv.getUint32(local, true) !== 0x04034b50) return null;
      var dataStart = local + 30 + dv.getUint16(local + 26, true) + dv.getUint16(local + 28, true);
      if (dataStart + e.csize > size) return null;
      e.data = new Uint8Array(buffer, dataStart, e.csize);
      entries.push(e);
      p += 46 + nameLen + extraLen + commentLen;
    }
    return entries;
  }

  // 把檔案清單重新包成 zip
  function writeZip(entries) {
    var total = 22;
    entries.forEach(function (e) { total += 30 + 46 + 2 * e.nameBytes.length + e.data.length; });
    var out = new Uint8Array(total);
    var dv = new DataView(out.buffer);
    var p = 0;
    entries.forEach(function (e) {
      e.offset = p;
      e.outFlags = e.flags & (e.method === 8 ? 0x0806 : 0x0800);  // 留下壓縮等級與 UTF-8 檔名旗標
      dv.setUint32(p, 0x04034b50, true);
      dv.setUint16(p + 4, 20, true);
      dv.setUint16(p + 6, e.outFlags, true);
      dv.setUint16(p + 8, e.method, true);
      dv.setUint16(p + 10, e.time, true);
      dv.setUint16(p + 12, e.date, true);
      dv.setUint32(p + 14, e.crc, true);
      dv.setUint32(p + 18, e.data.length, true);
      dv.setUint32(p + 22, e.usize, true);
      dv.setUint16(p + 26, e.nameBytes.length, true);
      out.set(e.nameBytes, p + 30);
      out.set(e.data, p + 30 + e.nameBytes.length);
      p += 30 + e.nameBytes.length + e.data.length;
    });
    var cdStart = p;
    entries.forEach(function (e) {
      dv.setUint32(p, 0x02014b50, true);
      dv.setUint16(p + 4, 20, true);
      dv.setUint16(p + 6, 20, true);
      dv.setUint16(p + 8, e.outFlags, true);
      dv.setUint16(p + 10, e.method, true);
      dv.setUint16(p + 12, e.time, true);
      dv.setUint16(p + 14, e.date, true);
      dv.setUint32(p + 16, e.crc, true);
      dv.setUint32(p + 20, e.data.length, true);
      dv.setUint32(p + 24, e.usize, true);
      dv.setUint16(p + 28, e.nameBytes.length, true);
      dv.setUint16(p + 36, e.intAttr, true);
      dv.setUint32(p + 38, e.extAttr, true);
      dv.setUint32(p + 42, e.offset, true);
      out.set(e.nameBytes, p + 46);
      p += 46 + e.nameBytes.length;
    });
    dv.setUint32(p, 0x06054b50, true);
    dv.setUint16(p + 8, entries.length, true);
    dv.setUint16(p + 10, entries.length, true);
    dv.setUint32(p + 12, p - cdStart, true);
    dv.setUint32(p + 16, cdStart, true);
    return out.buffer;
  }

  // 回傳修正過的檔案內容（ArrayBuffer）；不需要修或修不了就回傳原本的
  function fixFontFlags(buffer) {
    try {
      var entries = readZip(buffer);
      if (!entries) return buffer;
      var changed = false;
      entries.forEach(function (e) {
        if (!FIX_ENTRIES[e.name] || (e.method !== 8 && e.method !== 0)) return;
        var bytes = e.method === 8 ? inflateRaw(e.data, e.usize) : e.data;
        if (crc32(bytes) !== e.crc) return;                       // 解出來的內容不對，不處理
        if (bytes[0] === 0xFF || bytes[0] === 0xFE) return;       // 不是 UTF-8，不處理
        var text = new TextDecoder('utf-8').decode(bytes);
        var fixed = text.replace(OFF_FLAGS, '');
        if (fixed === text) return;
        var plain = new TextEncoder().encode(fixed);
        e.data = plain;                                           // 改過的檔案不壓縮，直接存
        e.method = 0;
        e.crc = crc32(plain);
        e.usize = plain.length;
        changed = true;
      });
      return changed ? writeZip(entries) : buffer;
    } catch (err) {
      return buffer;
    }
  }

  function load(arrayBuffer) {
    var head = new Uint8Array(arrayBuffer.slice(0, 4));
    if (head[0] === 0xD0 && head[1] === 0xCF) {
      return Promise.reject(ioError('OLD_XLS',
        '這個檔案是舊版 .xls 格式，或有密碼保護。請在 Excel 用「另存新檔」存成 .xlsx（並取消密碼）後再上傳。'));
    }
    if (!(head[0] === 0x50 && head[1] === 0x4B)) {
      return Promise.reject(ioError('NOT_XLSX',
        '這不是 Excel 活頁簿（.xlsx）。請確認上傳的是 .xlsx 檔。'));
    }
    var wb = new global.ExcelJS.Workbook();
    return wb.xlsx.load(fixFontFlags(arrayBuffer)).then(function () {
      if (!wb.worksheets.length) throw ioError('NO_SHEET', '這個活頁簿裡沒有工作表。');
      return wb;
    }, function (err) {
      if (err && err.code) throw err;
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

  /* ---------- 工作表的範圍 ---------- */

  // 逐列、逐格走訪「實際存在」的列與儲存格（包含只有格式、沒有值的格子）
  function eachStoredRow(ws, fn) {
    var rows = ws._rows;
    if (Array.isArray(rows)) {
      for (var i = 0; i < rows.length; i++) if (rows[i]) fn(rows[i], i + 1);
    } else {
      ws.eachRow({ includeEmpty: false }, fn);
    }
  }
  function eachStoredCell(row, fn) {
    var cells = row._cells;
    if (Array.isArray(cells)) {
      for (var i = 0; i < cells.length; i++) if (cells[i]) fn(cells[i], i + 1);
    } else {
      row.eachCell({ includeEmpty: false }, fn);
    }
  }

  // 回傳 { lastRow, colCount }：最後一列、最右一行「有值」的位置
  function extent(ws) {
    var lastRow = 0;
    var colCount = 0;
    eachStoredRow(ws, function (row, r) {
      var has = false;
      eachStoredCell(row, function (cell, c) {
        var v = cellRawValue(cell);
        if (v === null || v === undefined || v === '') return;
        has = true;
        if (c > colCount) colCount = c;
      });
      if (has && r > lastRow) lastRow = r;
    });
    return { lastRow: lastRow, colCount: colCount };
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

  /* ---------- 把整張工作表複製到另一個活頁簿 ---------- */

  /*
   * 把 src 工作表原樣複製進 dstWorkbook，取名為 name。
   * 值、格式、合併儲存格、行寬、列高、隱藏的行列、凍結窗格、自動篩選都會帶過去；
   * 公式換成計算結果（跨檔案之後原公式的參照會失效）。
   * dstWorkbook 裡如果已經有同名的工作表，會先移除再放新的。
   * 回傳 { worksheet, lastRow, colCount }
   */
  function copySheet(src, dstWorkbook, name) {
    // Excel 的工作表名稱不分大小寫，p003 和 P003 算同名
    dstWorkbook.worksheets
      .filter(function (w) { return String(w.name).toLowerCase() === String(name).toLowerCase(); })
      .forEach(function (w) { dstWorkbook.removeWorksheet(w.id); });

    var props = src.properties || {};
    var ws = dstWorkbook.addWorksheet(name, {
      properties: cleanObject({
        defaultRowHeight: props.defaultRowHeight,
        defaultColWidth: props.defaultColWidth
      })
    });

    var ext = extent(src);
    var maxCol = ext.colCount;

    eachStoredRow(src, function (row, r) {
      var nrow = null;
      eachStoredCell(row, function (cell, c) {
        if (!nrow) nrow = ws.getRow(r);
        var ncell = nrow.getCell(c);
        ncell.style = cell.style;
        if (c > maxCol) maxCol = c;
        if (cell.isMerged && cell.master !== cell) return;
        var v = exportValue(cell);
        if (v !== null && v !== '') ncell.value = v;
      });
      if (row.height || row.hidden) {
        if (!nrow) nrow = ws.getRow(r);
        if (row.height) nrow.height = row.height;
        if (row.hidden) nrow.hidden = true;
      }
    });

    for (var c = 1; c <= maxCol; c++) {
      var col = src.getColumn(c);
      if (col.width) ws.getColumn(c).width = col.width;
      if (col.hidden) ws.getColumn(c).hidden = true;
    }

    listMerges(src).forEach(function (m) {
      ws.mergeCellsWithoutStyle(m.top, m.left, m.bottom, m.right);
    });

    var v = (src.views && src.views[0]) || {};
    var view = cleanObject({ showGridLines: v.showGridLines, zoomScale: v.zoomScale });
    if (v.state === 'frozen') {
      view.state = 'frozen';
      view.xSplit = v.xSplit || 0;
      view.ySplit = v.ySplit || 0;
    }
    if (Object.keys(view).length) ws.views = [view];

    if (src.autoFilter) ws.autoFilter = src.autoFilter;

    var ps = src.pageSetup || {};
    Object.assign(ws.pageSetup, cleanObject({
      orientation: ps.orientation,
      paperSize: ps.paperSize,
      fitToPage: ps.fitToPage,
      fitToWidth: ps.fitToWidth,
      fitToHeight: ps.fitToHeight,
      printTitlesRow: ps.printTitlesRow
    }));

    return { worksheet: ws, lastRow: ext.lastRow, colCount: ext.colCount };
  }

  /* ---------- 表頭文字 ---------- */

  // 猜表頭在哪一列：前幾列當中，有值的格子最多的那一列（大標題通常只有一兩格有值）
  function guessHeaderRow(ws, maxRows) {
    var best = 1, bestCount = -1;
    for (var r = 1; r <= (maxRows || 10); r++) {
      var row = ws.findRow(r);
      if (!row) continue;
      var n = 0;
      eachStoredCell(row, function (cell) {
        if (cell.isMerged && cell.master !== cell) return;
        var v = cell.value;
        if (v !== null && v !== undefined && v !== '') n++;
      });
      if (n > bestCount) { bestCount = n; best = r; }
    }
    return best;
  }

  // 每一行的表頭文字（fromRow 到 toRow 之間，由上到下串起來；合併儲存格用左上角的字）
  // 回傳陣列，labels[行號] = '2027/01 月合計'
  function columnLabels(ws, fromRow, toRow, colCount) {
    var labels = [];
    for (var c = 1; c <= colCount; c++) {
      var parts = [];
      for (var r = fromRow; r <= toRow; r++) {
        var row = ws.findRow(r);
        var cell = row ? row.findCell(c) : null;
        var t = cell ? normalizeText(displayText(cell)).replace(/\s+/g, ' ') : '';
        if (t && parts.indexOf(t) < 0) parts.push(t);
      }
      labels[c] = parts.join(' ');
    }
    return labels;
  }

  /* ---------- 查表：用料號到另一張工作表找數值 ---------- */

  // 料號比對用的鍵：去頭尾空白、全形轉半形、不分大小寫。
  // 數字和文字的 123 視為相同，但 0012 和 12 不同（料號前面的 0 有意義）。
  function lookupKey(v) {
    v = plainValue(v);
    if (v === null || v === undefined) return '';
    if (v instanceof Date) return v.toISOString();
    return normalizeText(v).toUpperCase();
  }

  // 把參照工作表的某一行做成索引：{ 鍵: [列號, 列號…] }（由上到下）
  function buildIndex(ref, keyCol) {
    var index = {};
    eachStoredRow(ref, function (row, r) {
      var k = lookupKey(cellRawValue(row.findCell(keyCol)));
      if (k === '') return;
      if (!index[k]) index[k] = [];
      index[k].push(r);
    });
    return index;
  }

  function sheetRef(name) {
    return "'" + String(name).replace(/'/g, "''") + "'!";
  }

  // 在表頭列裡找文字是 header 的那一行（重做這一步時，覆蓋原本那一行而不是再加一行）
  function findHeaderColumn(ws, headerEnd, header) {
    var want = lookupKey(header);
    var found = 0;
    for (var r = 1; r <= headerEnd && !found; r++) {
      var row = ws.findRow(r);
      if (!row) continue;
      eachStoredCell(row, function (cell, c) {
        if (found || (cell.isMerged && cell.master !== cell)) return;
        if (lookupKey(cell.value) === want) found = c;
      });
    }
    return found;
  }

  // 把儲存格的值當成數字：數字照用、文字的 "100" 也算；空白回傳 null；不是數字回傳 NaN
  function toNumber(v) {
    v = plainValue(v);
    if (typeof v === 'number') return v;
    if (v === null || v === undefined) return null;
    if (v instanceof Date) return NaN;
    var t = normalizeText(v).replace(/,/g, '');
    if (t === '') return null;
    var n = Number(t);
    return isNaN(n) ? NaN : n;
  }

  // 取出 ref 第 r 列要帶回去的數值。有給 o.refDivCol 時是「取值行 ÷ 除數行」。
  // 回傳 { value, error, numFmt, top, bottom }；error 是 Excel 會顯示的錯誤（例如 #DIV/0!）
  function lookupValue(ref, r, o) {
    var row = ref.findRow(r);
    var topCell = row ? row.findCell(o.refValueCol) : null;
    if (!o.refDivCol) {
      return { value: plainValue(cellRawValue(topCell)), error: null, numFmt: topCell ? topCell.numFmt : undefined };
    }
    var bottomCell = row ? row.findCell(o.refDivCol) : null;
    var top = toNumber(cellRawValue(topCell));
    var bottom = toNumber(cellRawValue(bottomCell));
    var out = {
      value: null, error: null, numFmt: 'General',
      top: topCell ? displayText(topCell) : '', bottom: bottomCell ? displayText(bottomCell) : ''
    };
    if (top !== null && isNaN(top)) out.error = '#VALUE!';
    else if (bottom !== null && isNaN(bottom)) out.error = '#VALUE!';
    else if (bottom === null || bottom === 0) out.error = '#DIV/0!';
    else out.value = (top === null ? 0 : top) / bottom;
    return out;
  }

  function sameResult(a, b) {
    if (a.error || b.error) return a.error === b.error;
    if (typeof a.value === 'number' && typeof b.value === 'number') {
      return Math.abs(a.value - b.value) <= 1e-12 * Math.max(1, Math.abs(a.value), Math.abs(b.value));
    }
    return lookupKey(a.value) === lookupKey(b.value);
  }

  /*
   * 先算出查表結果，不動工作表（給畫面預覽、統計用）。
   *
   * o.dataStartRow  main 的資料從第幾列開始
   * o.keyCol        main 拿來對照的行（放料號的那一行）
   * o.refKeyCol     ref 要在哪一行找料號
   * o.refValueCol   ref 找到之後取哪一行的數值（例如 Labor）
   * o.refDivCol     （可省略）要除以 ref 的哪一行（例如 Base Qty）；有給的話帶回的是 取值行 ÷ 這一行
   * o.header        新增那一行的表頭文字，例如 'CT.'
   * o.index         （可省略）buildIndex(ref, o.refKeyCol) 的結果，重複計算時可以沿用
   *
   * 同一個料號在 ref 出現很多列時，取最上面那一列（和 VLOOKUP 一樣）。
   *
   * 回傳：
   *   targetCol   要填入的行（原本最後一行之後空一行；已經有同名表頭就用那一行）
   *   isNew       是不是新增的行
   *   rows        [{ row, key, label, refRow, value, error, top, bottom }]  refRow = 0 表示找不到
   *   found / missing / blank   找到、找不到、料號空白的列數
   *   errors      找到了、但算不出數值的列數（除數是 0 或空白、或不是數字）
   *   multi       找到的料號裡，有幾個在 ref 出現不只一列
   *   multiDiff   其中有幾個「各列算出來的數值不一樣」
   */
  function planLookup(main, ref, o) {
    var headerEnd = o.dataStartRow - 1;
    var ext = extent(main);
    var existing = findHeaderColumn(main, headerEnd, o.header);
    var lastCol = ext.colCount;
    var targetCol = existing || lastCol + 2;
    var index = o.index || buildIndex(ref, o.refKeyCol);

    var rows = [];
    var found = 0, missing = 0, blank = 0, errors = 0;
    var multiKeys = {}, diffKeys = {};
    for (var r = o.dataStartRow; r <= ext.lastRow; r++) {
      var row = main.findRow(r);
      var cell = row ? row.findCell(o.keyCol) : null;
      var key = lookupKey(cellRawValue(cell));
      if (key === '') { blank++; continue; }
      var hits = index[key];
      var item = { row: r, key: key, label: displayText(cell), refRow: 0, value: null, error: null };
      if (hits && hits.length) {
        found++;
        var first = lookupValue(ref, hits[0], o);
        item.refRow = hits[0];
        item.value = first.value;
        item.error = first.error;
        item.numFmt = first.numFmt;
        item.top = first.top;
        item.bottom = first.bottom;
        if (first.error) errors++;
        if (hits.length > 1 && !multiKeys[key]) {
          multiKeys[key] = true;
          for (var i = 1; i < hits.length; i++) {
            if (!sameResult(first, lookupValue(ref, hits[i], o))) { diffKeys[key] = true; break; }
          }
        }
        item.multi = hits.length;
        item.differs = !!diffKeys[key];
      } else {
        missing++;
      }
      rows.push(item);
    }

    return {
      targetCol: targetCol,
      isNew: !existing,
      lastCol: existing ? existing : lastCol,
      rows: rows,
      found: found,
      missing: missing,
      blank: blank,
      errors: errors,
      multi: Object.keys(multiKeys).length,
      multiDiff: Object.keys(diffKeys).length
    };
  }

  // 某一列要填入的公式（不含開頭的 =），例如 'P003'!M5 或 'P003'!M5/'P003'!K5
  function lookupFormula(refName, r, o) {
    var prefix = sheetRef(refName);
    var f = prefix + numToCol(o.refValueCol) + r;
    if (o.refDivCol) f += '/' + prefix + numToCol(o.refDivCol) + r;
    return f;
  }

  /*
   * 把查表結果寫進 main 工作表：每一列填入公式（='P003'!M5 或 ='P003'!M5/'P003'!K5），
   * 點儲存格就看得到數值是從 ref 的哪一格來的。找不到的料號留空白並標上底色。
   * 參數和 planLookup 一樣，回傳的也是同一份結果。
   */
  var MISSING_FILL = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFD9D9' } };
  function isMissingFill(fill) {
    return !!(fill && fill.fgColor && fill.fgColor.argb === MISSING_FILL.fgColor.argb);
  }

  function addLookupColumn(main, ref, o) {
    var plan = planLookup(main, ref, o);
    var headerEnd = o.dataStartRow - 1;
    var col = plan.targetCol;

    // 表頭：樣式照抄對照那一行（料號那一行）的表頭，大標題那一列不動
    if (headerEnd >= 1) {
      var top = headerEnd >= 2 ? 2 : 1;
      if (plan.isNew) {
        for (var r = top; r <= headerEnd; r++) {
          var src = main.findCell(r, o.keyCol);
          if (src) main.getCell(r, col).style = src.style;
        }
        if (headerEnd > top) main.mergeCellsWithoutStyle(top, col, headerEnd, col);
      }
      var headCell = main.getCell(top, col);
      (headCell.isMerged ? headCell.master : headCell).value = o.header;
    }

    // 資料：樣式照抄原本最後一行（重做時沿用這一行原本的樣式），數字格式改用 ref 那一格的
    var done = {};
    plan.rows.forEach(function (item) {
      var cell = main.getCell(item.row, col);
      var base = main.findCell(item.row, plan.isNew ? plan.lastCol : col);
      var style = Object.assign({}, base ? base.style : {});
      if (isMissingFill(style.fill)) delete style.fill;
      done[item.row] = true;
      if (item.refRow) {
        style.numFmt = item.numFmt || 'General';
        cell.style = style;
        cell.value = {
          formula: lookupFormula(ref.name, item.refRow, o),
          result: item.error ? { error: item.error }
            : (item.value === null || item.value === '' ? 0 : item.value)
        };
      } else {
        style.fill = MISSING_FILL;
        cell.style = style;
        cell.value = null;
      }
    });

    // 重做時，把這一行裡不再對得到的舊內容清掉
    if (!plan.isNew) {
      eachStoredRow(main, function (row, r) {
        if (r < o.dataStartRow || done[r]) return;
        var old = row.findCell(col);
        if (old) old.value = null;
      });
    }

    if (plan.isNew) main.getColumn(col).width = 11;

    // 請 Excel 開檔時重新計算一次公式
    var wb = main.workbook;
    if (wb) wb.calcProperties = Object.assign({}, wb.calcProperties, { fullCalcOnLoad: true });

    return plan;
  }

  function toBlob(workbook) {
    return workbook.xlsx.writeBuffer().then(function (buf) {
      return new Blob([buf], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
      });
    });
  }

  var ExcelIO = {
    VERSION: 5,           // 程式版本，顯示在畫面右上角；每次改版加 1（index.html 的 ?v= 也要改）
    colToNum: colToNum,
    numToCol: numToCol,
    normalizeText: normalizeText,
    parseColumns: parseColumns,
    formatColumns: formatColumns,
    plainValue: plainValue,
    matchKey: matchKey,
    displayText: displayText,
    load: load,
    fixFontFlags: fixFontFlags,
    inflateRaw: inflateRaw,
    listMerges: listMerges,
    inspect: inspect,
    extent: extent,
    filter: filter,
    copySheet: copySheet,
    guessHeaderRow: guessHeaderRow,
    columnLabels: columnLabels,
    lookupKey: lookupKey,
    buildIndex: buildIndex,
    planLookup: planLookup,
    lookupFormula: lookupFormula,
    addLookupColumn: addLookupColumn,
    toBlob: toBlob
  };

  global.ExcelIO = ExcelIO;
  if (typeof module !== 'undefined' && module.exports) module.exports = ExcelIO;
})(typeof window !== 'undefined' ? window : globalThis);
