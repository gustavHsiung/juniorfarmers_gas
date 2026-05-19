// =============================================
// 農產品登記表單 — Google Apps Script v2
// 新增：提交ID、查詢（多週次）、覆蓋修改
// =============================================

var SHEET_NAME = '農場菜單';
var HEADERS = ['提交ID', '時間戳記', '登記人', '農場', '週次',
               '品名', '數量', '單位', '基本進貨價',
               '批價', '批價門檻', '末端建議售價', '品項備注', '本批備註'];

var ORDER_SHEET_NAME = '料理人下單';
var ORDER_HEADERS = ['時間戳記', '登記人', '店家/料理人', '週次',
                     '品名', '數量', '單位', '單價（元）',
                     '交易總價', '送貨地點', '送貨時間', '付款狀態', '備註'];

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
    if (data.action === 'update')    return handleUpdate(data);
    if (data.action === 'updateRow') return handleUpdateRow(data);
    if (data.action === 'order')     return handleOrderInsert(data);
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

  // 為這批提交統一產生一個 ID
  var submissionId = Utilities.getUuid();

  rows.forEach(function(row, i) {
    sheet.appendRow(HEADERS.map(function(h) {
      if (h === '提交ID') return submissionId + '-' + i;          // ← 補上這行
      return row[h] !== undefined ? row[h] : '';
    }));
  });

  return jsonResponse({ status: 'success', inserted: rows.length, submissionId: submissionId });
}

function handleUpdateRow(data) {
  var sheet = getOrCreateSheet();
  var rowId = data['提交ID'];
  if (!rowId) return jsonResponse({ status: 'error', message: '缺少提交ID' });

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return jsonResponse({ status: 'error', message: '找不到對應的資料列' });

  // 一次讀完，在記憶體裡找目標列
  var allValues = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();
  var rowIndex = -1;
  for (var i = 0; i < allValues.length; i++) {
    if (String(allValues[i][0]) === rowId) { rowIndex = i; break; }
  }
  if (rowIndex === -1) return jsonResponse({ status: 'error', message: '找不到對應的資料列' });

  // 更新欄位，一次寫回整列
  var updated = allValues[rowIndex].slice();
  var row = data.row || {};
  HEADERS.forEach(function(h, colIdx) {
    if (h !== '提交ID' && row[h] !== undefined) updated[colIdx] = row[h];
  });
  sheet.getRange(rowIndex + 2, 1, 1, HEADERS.length).setValues([updated]);

  return jsonResponse({ status: 'success', updated: 1 });
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
    var cellVal = String(sheet.getRange(i, 1).getValue());
    if (cellVal === submissionId || cellVal.startsWith(submissionId + '-'))
      sheet.deleteRow(i);
  }

  rows.forEach(function(row) {
    sheet.appendRow(HEADERS.map(function(h) {
      return row[h] !== undefined ? row[h] : '';
    }));
  });
  return jsonResponse({ status: 'success', updated: rows.length });
}

function handleOrderInsert(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(ORDER_SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(ORDER_SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(ORDER_HEADERS);
    var hr = sheet.getRange(1, 1, 1, ORDER_HEADERS.length);
    hr.setFontWeight('bold');
    hr.setBackground('#27500A');
    hr.setFontColor('#ffffff');
    sheet.setFrozenRows(1);
  }

  var rows = data.rows;
  if (!rows || !Array.isArray(rows) || rows.length === 0)
    return jsonResponse({ status: 'error', message: '沒有資料' });

  rows.forEach(function(row) {
    sheet.appendRow(ORDER_HEADERS.map(function(h) {
      // 前端欄位名稱對應
      if (h === '店家/料理人') return row['店家料理人'] || '';
      return row[h] !== undefined ? row[h] : '';
    }));
  });

  return jsonResponse({ status: 'success', inserted: rows.length });
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
