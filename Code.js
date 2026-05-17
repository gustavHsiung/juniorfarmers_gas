// =============================================
// 農產品登記表單 — Google Apps Script v2
// 新增：提交ID、查詢（多週次）、覆蓋修改
// =============================================

var SHEET_NAME = '農場菜單';
var HEADERS = ['提交ID', '時間戳記', '登記人', '農場', '週次',
               '品名', '數量', '單位', '基本進貨價',
               '批價', '批價門檻', '末端建議售價', '品項備注', '本批備註'];

function getOrCreateSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    var hr = sheet.getRange(1, 1, 1, HEADERS.length);
    hr.setFontWeight('bold');
    hr.setBackground('#27500A');
    hr.setFontColor('#ffffff');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function rowToObj(row) {
  var obj = {};
  HEADERS.forEach(function(h, i) { obj[h] = row[i]; });
  return obj;
}

function serializeRow(row) {
  var r = {};
  Object.keys(row).forEach(function(k) {
    r[k] = row[k] instanceof Date
      ? Utilities.formatDate(row[k], 'Asia/Taipei', 'yyyy/MM/dd HH:mm')
      : String(row[k]);
  });
  return r;
}

// ── POST ────────────────────────────────────
function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    if (data.action === 'update') return handleUpdate(data);
    return handleInsert(data);
  } catch (err) {
    return jsonResponse({ status: 'error', message: err.toString() });
  }
}

function handleInsert(data) {
  var sheet = getOrCreateSheet();
  var rows = data.rows;
  if (!rows || !Array.isArray(rows) || rows.length === 0)
    return jsonResponse({ status: 'error', message: '沒有資料' });
  rows.forEach(function(row) {
    sheet.appendRow(HEADERS.map(function(h) {
      return row[h] !== undefined ? row[h] : '';
    }));
  });
  return jsonResponse({ status: 'success', inserted: rows.length });
}

function handleUpdate(data) {
  var sheet = getOrCreateSheet();
  var submissionId = data.submissionId;
  var rows = data.rows;
  if (!submissionId)
    return jsonResponse({ status: 'error', message: '缺少提交ID' });

  // 由下往上刪，避免列號位移
  var lastRow = sheet.getLastRow();
  for (var i = lastRow; i >= 2; i--) {
    if (sheet.getRange(i, 1).getValue() === submissionId)
      sheet.deleteRow(i);
  }

  rows.forEach(function(row) {
    sheet.appendRow(HEADERS.map(function(h) {
      return row[h] !== undefined ? row[h] : '';
    }));
  });
  return jsonResponse({ status: 'success', updated: rows.length });
}

// ── GET ─────────────────────────────────────
function doGet(e) {
  var params = (e && e.parameter) || {};
  if (params.action === 'query') return handleQuery(params);
  if (params.action === 'weeks') return handleWeeks();
  return jsonResponse({ status: 'ok', message: '農產品登記 API 正常運作中 ✅' });
}

// 回傳工作表中所有不重複週次（供前端顯示選擇按鈕）
function handleWeeks() {
  var sheet = getOrCreateSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return jsonResponse({ status: 'success', weeks: [] });

  var weekCol = sheet.getRange(2, 5, lastRow - 1, 1).getValues(); // 欄5 = 週次
  var seen = {};
  weekCol.forEach(function(r) {
    var w = String(r[0]).trim();
    if (w) seen[w] = true;
  });

  // 排序：最新在前
  var weeks = Object.keys(seen).sort().reverse();
  return jsonResponse({ status: 'success', weeks: weeks });
}

// 查詢：weeks 為逗號分隔多週次，例如 "2025-05-第1週,2025-05-第2週"
function handleQuery(params) {
  var sheet = getOrCreateSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return jsonResponse({ status: 'success', data: [] });

  var weeksParam = (params.weeks || '').trim();
  var weekSet = {};
  if (weeksParam) {
    weeksParam.split(',').forEach(function(w) {
      var t = w.trim();
      if (t) weekSet[t] = true;
    });
  }
  var filterByWeek = Object.keys(weekSet).length > 0;

  var allData = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();
  var results = allData
    .map(rowToObj)
    .filter(function(row) {
      if (filterByWeek && !weekSet[String(row['週次'])]) return false;
      return String(row['品名']).trim() !== '';
    })
    .map(serializeRow);

  return jsonResponse({ status: 'success', data: results });
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
