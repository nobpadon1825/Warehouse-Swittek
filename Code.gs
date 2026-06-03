// =====================================================
//  Google Apps Script — Warehouse System Backend
//  วางโค้ดทั้งหมดนี้ใน Apps Script Editor แล้ว Deploy
// =====================================================

var SPREADSHEET_ID = '1wUsxLhjcLV8D3U5v3W8AWgHIrlODcJGGVx7AD5EQMVU'; // deployment: AKfycbx6XxPfkPcDI9kNgHz07oABWNadlmrFzIqeNuZL9XKYoFVctFW5iaPrYhu7MWum7QU

var DATA_KEYS = [
  'wh_stocks', 'wh_damaged', 'wh_transactions', 'wh_installs',
  'wh_demo_loans', 'wh_categories', 'wh_buyers',
  'wh_login_history', 'wh_extra_accounts', 'wh_avatars'
];

// ---- GET: โหลดข้อมูลทั้งหมด (รองรับ JSONP callback) ----
function doGet(e) {
  var callback = e.parameter.callback;
  var result = getAllData();
  var json = JSON.stringify(result);
  if (callback) {
    return ContentService.createTextOutput(callback + '(' + json + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

// ---- POST: บันทึกข้อมูล / อัพโหลดรูปไป Google Drive ----
function doPost(e) {
  try {
    var d = JSON.parse(e.postData.contents);

    // ---------- บันทึกข้อมูลทั้งหมด ----------
    if (d.action === 'saveAll') {
      saveAllData(d.payload);
      return ContentService.createTextOutput(JSON.stringify({ ok: true }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // ---------- อัพโหลดรูปไป Google Drive ----------
    if (d.action === 'uploadImage') {
      var folder = getOrCreateImageFolder();
      var base64 = d.imageData;
      // ตัด prefix "data:image/jpeg;base64," ออก
      if (base64.indexOf(',') !== -1) base64 = base64.split(',')[1];
      var blob = Utilities.newBlob(
        Utilities.base64Decode(base64),
        d.mimeType || 'image/jpeg',
        d.filename || ('warehouse_img_' + Date.now() + '.jpg')
      );
      var file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      var url = 'https://lh3.googleusercontent.com/d/' + file.getId();
      return ContentService.createTextOutput(JSON.stringify({ ok: true, url: url }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // ---------- ลบรูปใน Google Drive ----------
    if (d.action === 'deleteImages') {
      var ids = d.fileIds || [];
      ids.forEach(function(fid) {
        try { DriveApp.getFileById(fid).setTrashed(true); } catch(e2) {}
      });
      return ContentService.createTextOutput(JSON.stringify({ ok: true }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: 'unknown action' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ---- หา / สร้าง folder "Warehouse Images" ใน Google Drive ----
function getOrCreateImageFolder() {
  var folders = DriveApp.getFoldersByName('Warehouse Images');
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder('Warehouse Images');
}

var CACHE_KEY = 'wh_data_';
var CACHE_TTL = 300; // วินาที (5 นาที)

function _pad2(n) { return n < 10 ? '0' + n : String(n); }

function _parseSheetValues(allValues) {
  if (!allValues || allValues.length < 2) return [];
  var headers = allValues[0];
  return allValues.slice(1).map(function(row) {
    var obj = {};
    headers.forEach(function(h, i) {
      var v = row[i];
      // Google Sheets Date → YYYY-MM-DD ตาม timezone ของ Spreadsheet
      if (v instanceof Date) {
        obj[h] = Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        return;
      }
      if (typeof v === 'string') {
        var t = v.trim();
        if ((t.startsWith('{') && t.endsWith('}')) ||
            (t.startsWith('[') && t.endsWith(']'))) {
          try { obj[h] = JSON.parse(t); return; } catch(e2) {}
        }
      }
      obj[h] = v;
    });
    return obj;
  });
}

// ---- อ่านข้อมูลจากทุก sheet — batch API calls ----
function getAllData() {
  var cache = CacheService.getScriptCache();

  // ดึง cache ทั้งหมดในคราวเดียว
  var cacheKeys = DATA_KEYS.map(function(k){ return CACHE_KEY + k; });
  var cached    = cache.getAll(cacheKeys);
  var result    = {};
  var missing   = [];

  DATA_KEYS.forEach(function(key) {
    var hit = cached[CACHE_KEY + key];
    if (hit) {
      try { result[key] = JSON.parse(hit); return; } catch(e) {}
    }
    missing.push(key);
  });

  if (missing.length === 0) return result;

  // อ่านเฉพาะ sheet ที่ยังไม่มี cache — เปิด SS 1 ครั้ง + getSheets() 1 ครั้ง
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheetMap = {};
  ss.getSheets().forEach(function(s){ sheetMap[s.getName()] = s; });

  var toCache = {};
  missing.forEach(function(key) {
    var sheet = sheetMap[key];
    if (!sheet) { result[key] = []; return; }
    // getDataRange() = 1 API call แทน 4 calls เดิม
    var vals = sheet.getDataRange().getValues();
    result[key] = _parseSheetValues(vals);
    var str = JSON.stringify(result[key]);
    if (str.length < 100000) toCache[CACHE_KEY + key] = str;
  });

  if (Object.keys(toCache).length > 0) {
    try { cache.putAll(toCache, CACHE_TTL); } catch(e) {}
  }
  return result;
}

// ---- เขียนข้อมูลลง Sheets (columnar: header row + data rows) ----
function saveAllData(payload) {
  var keys = Object.keys(payload).filter(function(k){ return DATA_KEYS.indexOf(k) !== -1; });
  // invalidate เฉพาะ key ที่กำลังเขียน
  try { CacheService.getScriptCache().removeAll(keys.map(function(k){ return CACHE_KEY + k; })); } catch(e) {}
  var ss   = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheetMap = {};
  ss.getSheets().forEach(function(s){ sheetMap[s.getName()] = s; });
  var lock = LockService.getScriptLock();
  try {
    lock.tryLock(5000);
    keys.forEach(function(key) {
      var sheet = sheetMap[key] || ss.insertSheet(key);
      var arr = payload[key];
      sheet.clearContents();
      if (!arr || arr.length === 0) return;
      var headers = Object.keys(arr[0]);
      var rows = [headers];
      arr.forEach(function(obj) {
        rows.push(headers.map(function(h) {
          var v = obj[h];
          if (v === null || v === undefined) return '';
          if (typeof v === 'object') return JSON.stringify(v);
          return v;
        }));
      });
      sheet.getRange(1, 1, rows.length, headers.length).setValues(rows);
    });
  } finally {
    lock.releaseLock();
  }
}

// ---- แปลง old JSON-blob format → columnar (รันครั้งเดียว) ----
function migrateOldToNew() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var migrated = 0, skipped = 0;
  DATA_KEYS.forEach(function(key) {
    var sheet = ss.getSheetByName(key);
    if (!sheet) { skipped++; return; }
    var lastCol = sheet.getLastColumn();
    if (lastCol === 0) { skipped++; return; }
    var firstCell = sheet.getRange(1, 1).getValue();
    if (typeof firstCell !== 'string' || !firstCell.trim().startsWith('[')) { skipped++; return; }
    var values  = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    var combined = values.join('');
    var arr;
    try { arr = JSON.parse(combined); } catch(e) { skipped++; return; }
    if (!Array.isArray(arr) || arr.length === 0) { skipped++; return; }
    var headers = Object.keys(arr[0]);
    sheet.clearContents();
    var rows = [headers];
    arr.forEach(function(obj) {
      rows.push(headers.map(function(h) {
        var v = obj[h];
        if (v === null || v === undefined) return '';
        if (typeof v === 'object') return JSON.stringify(v);
        return v;
      }));
    });
    sheet.getRange(1, 1, rows.length, headers.length).setValues(rows);
    migrated++;
  });
  Logger.log('Migrated: ' + migrated + ', Skipped: ' + skipped);
}

// ---- แก้ wh_categories / wh_buyers ที่ headers เป็นตัวเลข (character-split bug) ----
function fixStringArraySheets(ss) {
  var toFix = ['wh_categories', 'wh_buyers'];
  toFix.forEach(function(key) {
    var sheet = ss.getSheetByName(key);
    if (!sheet) return;
    var data = sheet.getDataRange().getValues();
    if (data.length === 0) return;
    var firstHeader = data[0][0];
    if (firstHeader === 0 || firstHeader === '0') {
      var names = [];
      for (var i = 1; i < data.length; i++) {
        var chars = data[i].filter(function(c){ return c !== '' && c !== null; });
        if (chars.length > 0) names.push(chars.join(''));
      }
      sheet.clearContents();
      if (names.length > 0) {
        var rows = [['name']];
        names.forEach(function(n){ rows.push([n]); });
        sheet.getRange(1, 1, rows.length, 1).setValues(rows);
      }
    }
  });
}

// ---- จัดรูปแบบทุก Sheet ให้สวยงาม (รัน fixAndFormatAllSheets จาก Editor) ----
function fixAndFormatAllSheets() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  fixStringArraySheets(ss);
  DATA_KEYS.forEach(function(key) {
    var sheet = ss.getSheetByName(key);
    if (sheet) formatSheet(sheet);
  });
}

function formatSheet(sheet) {
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 1 || lastCol < 1) return;

  // Header row — blue
  var headerRange = sheet.getRange(1, 1, 1, lastCol);
  headerRange.setBackground('#1565C0');
  headerRange.setFontColor('#FFFFFF');
  headerRange.setFontWeight('bold');
  headerRange.setFontSize(10);
  headerRange.setHorizontalAlignment('center');
  headerRange.setVerticalAlignment('middle');
  sheet.setRowHeight(1, 32);

  // ล้าง per-row backgrounds เดิม
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, lastCol).setBackground(null);
  }

  // Row Banding — ครอบคลุมทุก row ที่จะเพิ่มในอนาคตด้วย
  sheet.getBandings().forEach(function(b){ b.remove(); });
  var maxRows = sheet.getMaxRows();
  if (maxRows > 1) {
    var banding = sheet.getRange(2, 1, maxRows - 1, lastCol).applyRowBanding();
    banding.setFirstRowColor('#FFFFFF');
    banding.setSecondRowColor('#E3F2FD');
    banding.setHeaderRowColor(null);
    banding.setFooterRowColor(null);
  }

  // Borders
  var allData = sheet.getRange(1, 1, lastRow, lastCol);
  allData.setBorder(true, true, true, true, true, true,
    '#90CAF9', SpreadsheetApp.BorderStyle.SOLID);
  allData.setBorder(true, true, true, true, null, null,
    '#1565C0', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  // Auto-resize columns (min 80px)
  for (var c = 1; c <= lastCol; c++) {
    sheet.autoResizeColumn(c);
    if (sheet.getColumnWidth(c) < 80) sheet.setColumnWidth(c, 80);
  }

  // Freeze header row
  sheet.setFrozenRows(1);

  // Tab color ตาม sheet name
  var tabColors = {
    'wh_stocks':        '#1565C0',
    'wh_damaged':       '#B71C1C',
    'wh_transactions':  '#1B5E20',
    'wh_installs':      '#4A148C',
    'wh_demo_loans':    '#E65100',
    'wh_categories':    '#006064',
    'wh_buyers':        '#880E4F',
    'wh_login_history': '#37474F',
    'wh_extra_accounts':'#3E2723'
  };
  var color = tabColors[sheet.getName()];
  if (color) sheet.setTabColor(color);
}
