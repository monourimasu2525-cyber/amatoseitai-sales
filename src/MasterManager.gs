// ========== MasterManager.gs ==========
// 売上マスタ管理

var MASTER_SHEET_NAME = '売上マスタ';

function getMasterItems() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(MASTER_SHEET_NAME);
  if (!sheet) {
    initMasterSheet();
    sheet = ss.getSheetByName(MASTER_SHEET_NAME);
  }
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var data = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
  var result = [];
  data.forEach(function(row, i) {
    if (row[3] === true || row[3] === 'TRUE') {
      result.push({
        rowIndex: i + 2,
        type: row[0],
        amount: Number(row[1]) || 0,
        description: row[2] || ''
      });
    }
  });
  return result;
}

function addMasterItem(type, amount, description) {
  if (!type || !amount) {
    return { success: false, message: '種別と金額は必須です' };
  }
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(MASTER_SHEET_NAME);
    if (!sheet) {
      initMasterSheet();
      sheet = ss.getSheetByName(MASTER_SHEET_NAME);
    }
    sheet.appendRow([type, Number(amount), description || '', true]);
    return { success: true, message: type + ' を追加しました' };
  } catch (err) {
    return { success: false, message: 'エラー: ' + err.message };
  }
}

function updateMasterItem(rowIndex, type, amount, description) {
  if (!rowIndex || !type || !amount) {
    return { success: false, message: '行番号・種別・金額は必須です' };
  }
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(MASTER_SHEET_NAME);
    if (!sheet) return { success: false, message: 'マスタシートが存在しません' };
    sheet.getRange(rowIndex, 1, 1, 4).setValues([[type, Number(amount), description || '', true]]);
    return { success: true, message: type + ' を更新しました' };
  } catch (err) {
    return { success: false, message: 'エラー: ' + err.message };
  }
}

function deleteMasterItem(rowIndex) {
  if (!rowIndex) {
    return { success: false, message: '行番号は必須です' };
  }
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(MASTER_SHEET_NAME);
    if (!sheet) return { success: false, message: 'マスタシートが存在しません' };
    // 有効フラグをFALSEに
    sheet.getRange(rowIndex, 4).setValue(false);
    return { success: true, message: 'マスタを無効化しました' };
  } catch (err) {
    return { success: false, message: 'エラー: ' + err.message };
  }
}

function initMasterSheet() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var existing = ss.getSheetByName(MASTER_SHEET_NAME);
    if (existing) return { success: true, message: 'マスタシートは既に存在します' };

    var sheet = ss.insertSheet(MASTER_SHEET_NAME);

    // ヘッダー行
    sheet.getRange(1, 1, 1, 4).setValues([['種別', '金額', '説明', '有効']]);
    var headerRange = sheet.getRange(1, 1, 1, 4);
    headerRange.setBackground('#1a237e');
    headerRange.setFontColor('#ffffff');
    headerRange.setFontWeight('bold');
    headerRange.setFontSize(11);
    headerRange.setHorizontalAlignment('center');

    // 初期データ
    sheet.appendRow(['新規', 3270, '新規施術', true]);
    sheet.appendRow(['常連', 5500, '常連施術', true]);

    // データ行の書式
    sheet.getRange(2, 1, 2, 4).setBackground('#e8eaf6');

    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, 4);

    return { success: true, message: 'マスタシートを初期化しました' };
  } catch (err) {
    return { success: false, message: 'マスタシート初期化エラー: ' + err.message };
  }
}
