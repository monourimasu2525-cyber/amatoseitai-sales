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
