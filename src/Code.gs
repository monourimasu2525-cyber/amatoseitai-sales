var SPREADSHEET_ID = '17bAyQngDEjoDgqSLLUU5p45HWXomF09bLf_h6FySsjs';
var BACKUP_FOLDER_NAME = 'あまと整体院_売上バックアップ';

// ========== SalesManager ==========

class SalesManager {
  constructor(spreadsheetId) {
    this.ss = SpreadsheetApp.openById(spreadsheetId);
    this.dataSheet = this.ss.getSheetByName('売上データ') || this.ss.getSheets()[0];
  }

  addSale(type, amount) {
    if (type !== '新規' && type !== '常連') {
      return { success: false, message: '種別は「新規」または「常連」です' };
    }
    if (typeof amount !== 'number' || amount < 0) {
      return { success: false, message: '金額は正の数値です' };
    }
    try {
      const now = new Date();
      this.dataSheet.appendRow([now, now, type, amount, 'WebAPI']);
      return { success: true, message: type + ' ¥' + amount + ' を登録しました', timestamp: now.toISOString(), type: type, amount: amount };
    } catch (err) {
      return { success: false, message: 'エラー: ' + err.message };
    }
  }

  getTodayStats() {
    const today = new Date();
    const year = today.getFullYear();
    const month = today.getMonth() + 1;
    const day = today.getDate();
    const lastRow = this.dataSheet.getLastRow();
    if (lastRow < 2) {
      return { date: year + '年' + month + '月' + day + '日', shinkiCount: 0, jorenCount: 0, totalCount: 0, shinkiSales: 0, jorenSales: 0, totalSales: 0 };
    }
    const allData = this.dataSheet.getRange(2, 1, lastRow - 1, 5).getValues();
    let shinkiCount = 0, shinkiSales = 0, jorenCount = 0, jorenSales = 0;
    allData.forEach(function(row) {
      if (!row[0]) return;
      const date = new Date(row[0]);
      if (date.getFullYear() === year && date.getMonth() + 1 === month && date.getDate() === day) {
        const amount = Number(row[3]) || 0;
        if (row[2] === '新規') { shinkiCount++; shinkiSales += amount; }
        else if (row[2] === '常連') { jorenCount++; jorenSales += amount; }
      }
    });
    return { date: year + '年' + month + '月' + day + '日', shinkiCount: shinkiCount, jorenCount: jorenCount, totalCount: shinkiCount + jorenCount, shinkiSales: shinkiSales, jorenSales: jorenSales, totalSales: shinkiSales + jorenSales };
  }

  // 月間集計（先月比・前年度比用）
  getMonthStats(year, month) {
    const lastRow = this.dataSheet.getLastRow();
    if (lastRow < 2) return { shinkiCount: 0, jorenCount: 0, totalCount: 0, shinkiSales: 0, jorenSales: 0, totalSales: 0 };
    const allData = this.dataSheet.getRange(2, 1, lastRow - 1, 5).getValues();
    let shinkiCount = 0, shinkiSales = 0, jorenCount = 0, jorenSales = 0;
    allData.forEach(function(row) {
      if (!row[0]) return;
      const date = new Date(row[0]);
      if (date.getFullYear() === year && date.getMonth() + 1 === month) {
        const amount = Number(row[3]) || 0;
        if (row[2] === '新規') { shinkiCount++; shinkiSales += amount; }
        else if (row[2] === '常連') { jorenCount++; jorenSales += amount; }
      }
    });
    return { shinkiCount: shinkiCount, jorenCount: jorenCount, totalCount: shinkiCount + jorenCount, shinkiSales: shinkiSales, jorenSales: jorenSales, totalSales: shinkiSales + jorenSales };
  }

  // 履歴取得（直近N日分）
  getRecentHistory(days) {
    const lastRow = this.dataSheet.getLastRow();
    if (lastRow < 2) return [];
    const allData = this.dataSheet.getRange(2, 1, lastRow - 1, 5).getValues();
    const cutoff = new Date();
    cutoff.setDate(cutoff.getDate() - days);
    const result = [];
    allData.forEach(function(row) {
      if (!row[0]) return;
      const date = new Date(row[0]);
      if (date >= cutoff) {
        result.push({
          date: date.getFullYear() + '/' + (date.getMonth()+1) + '/' + date.getDate(),
          time: date.getHours() + ':' + String(date.getMinutes()).padStart(2,'0'),
          type: row[2],
          amount: Number(row[3]) || 0
        });
      }
    });
    return result.reverse();
  }

  // CSV出力用データ取得
  getCsvData(year, month) {
    const lastRow = this.dataSheet.getLastRow();
    if (lastRow < 2) return [];
    const allData = this.dataSheet.getRange(2, 1, lastRow - 1, 5).getValues();
    const result = [['日付', '時刻', '種別', '金額', '入力方法']];
    allData.forEach(function(row) {
      if (!row[0]) return;
      const date = new Date(row[0]);
      if (!year || (date.getFullYear() === year && date.getMonth() + 1 === month)) {
        result.push([
          date.getFullYear() + '/' + (date.getMonth()+1) + '/' + date.getDate(),
          date.getHours() + ':' + String(date.getMinutes()).padStart(2,'0'),
          row[2],
          Number(row[3]) || 0,
          row[4] || 'WebAPI'
        ]);
      }
    });
    return result;
  }
}

// ========== バックアップ ==========

function getOrCreateBackupFolder() {
  const folders = DriveApp.getFoldersByName(BACKUP_FOLDER_NAME);
  if (folders.hasNext()) {
    return folders.next();
  }
  return DriveApp.createFolder(BACKUP_FOLDER_NAME);
}

function runBackup() {
  try {
    const folder = getOrCreateBackupFolder();
    const now = new Date();
    const label = now.getFullYear() + '-'
      + String(now.getMonth()+1).padStart(2,'0') + '-'
      + String(now.getDate()).padStart(2,'0')
      + '_' + String(now.getHours()).padStart(2,'0')
      + String(now.getMinutes()).padStart(2,'0');
    const fileName = 'あまと整体院_売上データ_' + label;
    const original = DriveApp.getFileById(SPREADSHEET_ID);
    const copy = original.makeCopy(fileName, folder);
    return { success: true, message: 'バックアップ完了: ' + fileName, fileId: copy.getId() };
  } catch (err) {
    return { success: false, message: 'バックアップエラー: ' + err.message };
  }
}

// 自動バックアップ用（トリガーから呼ばれる）
function dailyAutoBackup() {
  runBackup();
}

// トリガーセットアップ（初回1回だけ手動で実行）
function setupDailyBackupTrigger() {
  // 既存トリガー削除
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'dailyAutoBackup') {
      ScriptApp.deleteTrigger(t);
    }
  });
  // 毎日深夜2時に実行
  ScriptApp.newTrigger('dailyAutoBackup')
    .timeBased()
    .everyDays(1)
    .atHour(2)
    .create();
}

// ========== 集計シート書き込み ==========

function updateSummarySheet() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var summarySheet = ss.getSheetByName('集計');
    if (!summarySheet) {
      summarySheet = ss.insertSheet('集計');
    }

    // シート全体をクリアしてから書き直す
    summarySheet.clearContents();
    summarySheet.clearFormats();

    var now = new Date();
    var headers = ['年月', '新規件数', '新規売上', '常連件数', '常連売上', '合計件数', '合計売上'];
    var rows = [headers];

    for (var i = 11; i >= 0; i--) {
      var d = new Date(now.getFullYear(), now.getMonth() - i, 1);
      var y = d.getFullYear();
      var m = d.getMonth() + 1;
      var manager = new SalesManager(SPREADSHEET_ID);
      var stats = manager.getMonthStats(y, m);
      rows.push([
        y + '年' + m + '月',
        stats.shinkiCount,
        stats.shinkiSales,
        stats.jorenCount,
        stats.jorenSales,
        stats.totalCount,
        stats.totalSales
      ]);
    }

    var range = summarySheet.getRange(1, 1, rows.length, headers.length);
    range.setValues(rows);

    // ヘッダー行の書式設定
    var headerRange = summarySheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#1a237e');
    headerRange.setFontColor('#ffffff');
    headerRange.setFontWeight('bold');

    // データ行の交互背景色
    for (var r = 2; r <= rows.length; r++) {
      var rowRange = summarySheet.getRange(r, 1, 1, headers.length);
      rowRange.setBackground(r % 2 === 0 ? '#e8eaf6' : '#ffffff');
    }

    // 数値列を通貨フォーマット
    summarySheet.getRange(2, 3, rows.length - 1, 1).setNumberFormat('¥#,##0');
    summarySheet.getRange(2, 5, rows.length - 1, 1).setNumberFormat('¥#,##0');
    summarySheet.getRange(2, 7, rows.length - 1, 1).setNumberFormat('¥#,##0');

    summarySheet.autoResizeColumns(1, headers.length);

    return { success: true, message: '集計シートを更新しました（過去12ヶ月）' };
  } catch (err) {
    return { success: false, message: '集計シート更新エラー: ' + err.message };
  }
}

// ========== 経理シート生成 ==========

function generateAccountingSheet(year, month) {
  try {
    year = parseInt(year);
    month = parseInt(month);
    if (!year || !month || month < 1 || month > 12) {
      return { success: false, message: '年月が不正です' };
    }

    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheetName = '経理_' + year + '_' + String(month).padStart(2, '0');

    // 既存シートがあれば削除して再作成
    var existing = ss.getSheetByName(sheetName);
    if (existing) {
      ss.deleteSheet(existing);
    }
    var accSheet = ss.insertSheet(sheetName);

    // ヘッダー
    var headers = ['日付', '新規件数', '新規売上', '常連件数', '常連売上', '日計'];
    accSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    var headerRange = accSheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#1a237e');
    headerRange.setFontColor('#ffffff');
    headerRange.setFontWeight('bold');

    // 売上データシートからデータ取得
    var dataSheet = ss.getSheetByName('売上データ') || ss.getSheets()[0];
    var lastRow = dataSheet.getLastRow();
    var allData = lastRow >= 2 ? dataSheet.getRange(2, 1, lastRow - 1, 5).getValues() : [];

    // 日別集計
    var daysInMonth = new Date(year, month, 0).getDate();
    var dayMap = {};
    for (var d = 1; d <= daysInMonth; d++) {
      dayMap[d] = { shinkiCount: 0, shinkiSales: 0, jorenCount: 0, jorenSales: 0 };
    }

    allData.forEach(function(row) {
      if (!row[0]) return;
      var date = new Date(row[0]);
      if (date.getFullYear() === year && date.getMonth() + 1 === month) {
        var day = date.getDate();
        var amount = Number(row[3]) || 0;
        if (row[2] === '新規') {
          dayMap[day].shinkiCount++;
          dayMap[day].shinkiSales += amount;
        } else if (row[2] === '常連') {
          dayMap[day].jorenCount++;
          dayMap[day].jorenSales += amount;
        }
      }
    });

    // 行データ作成
    var dataRows = [];
    var totalShinki = 0, totalShinkiSales = 0, totalJoren = 0, totalJorenSales = 0;
    for (var day = 1; day <= daysInMonth; day++) {
      var dk = dayMap[day];
      var dailyTotal = dk.shinkiSales + dk.jorenSales;
      totalShinki += dk.shinkiCount;
      totalShinkiSales += dk.shinkiSales;
      totalJoren += dk.jorenCount;
      totalJorenSales += dk.jorenSales;
      dataRows.push([
        year + '/' + String(month).padStart(2, '0') + '/' + String(day).padStart(2, '0'),
        dk.shinkiCount,
        dk.shinkiSales,
        dk.jorenCount,
        dk.jorenSales,
        dailyTotal
      ]);
    }

    // 月合計行
    dataRows.push([
      '【月合計】',
      totalShinki,
      totalShinkiSales,
      totalJoren,
      totalJorenSales,
      totalShinkiSales + totalJorenSales
    ]);

    var dataRange = accSheet.getRange(2, 1, dataRows.length, headers.length);
    dataRange.setValues(dataRows);

    // 月合計行の書式
    var totalRowRange = accSheet.getRange(daysInMonth + 2, 1, 1, headers.length);
    totalRowRange.setBackground('#e8eaf6');
    totalRowRange.setFontWeight('bold');

    accSheet.autoResizeColumns(1, headers.length);

    return { success: true, message: sheetName + ' を生成しました（' + daysInMonth + '日分）' };
  } catch (err) {
    return { success: false, message: '経理シート生成エラー: ' + err.message };
  }
}

// ========== WebAPI ==========

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const manager = new SalesManager(SPREADSHEET_ID);
    let result;

    if (data.action === 'addSale') {
      result = manager.addSale(data.type, data.amount);
    } else if (data.action === 'backup') {
      result = runBackup();
    } else if (data.action === 'updateSummary') {
      result = updateSummarySheet();
    } else if (data.action === 'generateAccounting') {
      result = generateAccountingSheet(data.year, data.month);
    } else {
      result = { success: false, message: '不明なアクション: ' + data.action };
    }

    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  try {
    const action = e.parameter.action || 'getTodayStats';
    const manager = new SalesManager(SPREADSHEET_ID);
    let result;

    if (action === 'getSheets') {
      const sheets = manager.ss.getSheets().map(function(s) { return s.getName(); });
      result = { sheets: sheets, dataSheetName: manager.dataSheet ? manager.dataSheet.getName() : null };
    } else if (action === 'getTodayStats') {
      result = manager.getTodayStats();
    } else if (action === 'getMonthStats') {
      const year = parseInt(e.parameter.year) || new Date().getFullYear();
      const month = parseInt(e.parameter.month) || new Date().getMonth() + 1;
      result = manager.getMonthStats(year, month);
    } else if (action === 'getRecentHistory') {
      const days = parseInt(e.parameter.days) || 30;
      result = { records: manager.getRecentHistory(days) };
    } else if (action === 'getCsv') {
      const year = e.parameter.year ? parseInt(e.parameter.year) : null;
      const month = e.parameter.month ? parseInt(e.parameter.month) : null;
      const rows = manager.getCsvData(year, month);
      const csv = rows.map(function(r) { return r.join(','); }).join('\n');
      return ContentService
        .createTextOutput(csv)
        .setMimeType(ContentService.MimeType.CSV);
    } else if (action === 'updateSummary') {
      result = updateSummarySheet();
    } else if (action === 'generateAccounting') {
      const year = parseInt(e.parameter.year) || new Date().getFullYear();
      const month = parseInt(e.parameter.month) || new Date().getMonth() + 1;
      result = generateAccountingSheet(year, month);
    } else {
      result = { success: false, message: '不明なアクション' };
    }

    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ========== スプレッドシートメニュー ==========

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('売上管理')
    .addItem('新規売上を記録', 'addNewCustomer')
    .addItem('常連売上を記録', 'addRegularCustomer')
    .addSeparator()
    .addItem('今すぐバックアップ', 'manualBackupFromSheet')
    .addItem('自動バックアップ設定', 'setupDailyBackupTrigger')
    .addToUi();
}

function addNewCustomer() {
  const manager = new SalesManager(SPREADSHEET_ID);
  manager.addSale('新規', 3270);
  SpreadsheetApp.getUi().alert('新規 ¥3,270 を登録しました！');
}

function addRegularCustomer() {
  const manager = new SalesManager(SPREADSHEET_ID);
  manager.addSale('常連', 5500);
  SpreadsheetApp.getUi().alert('常連 ¥5,500 を登録しました！');
}

function manualBackupFromSheet() {
  const result = runBackup();
  SpreadsheetApp.getUi().alert(result.message);
}
