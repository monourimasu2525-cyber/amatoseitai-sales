// ========== Code.gs ==========
// メインエントリ・doGet/doPost/onOpen/メニュー関数

var SPREADSHEET_ID = '17bAyQngDEjoDgqSLLUU5p45HWXomF09bLf_h6FySsjs';
var BACKUP_FOLDER_NAME = 'あまと整体院_売上バックアップ';

// ========== WebAPI: POST ==========

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const manager = new SalesManager(SPREADSHEET_ID);
    let result;

    switch (data.action) {
      case 'addSale':
        result = manager.addSale(data.type, data.amount);
        break;
      case 'editSale':
        result = manager.editSale(data.rowIndex, data.type, data.amount);
        break;
      case 'deleteSale':
        result = manager.deleteSale(data.rowIndex);
        break;
      case 'addMaster':
        result = addMasterItem(data.type, data.amount, data.description);
        break;
      case 'updateMaster':
        result = updateMasterItem(data.rowIndex, data.type, data.amount, data.description);
        break;
      case 'deleteMaster':
        result = deleteMasterItem(data.rowIndex);
        break;
      case 'backup':
        result = runBackup();
        break;
      case 'updateSummary':
        result = updateSummarySheet();
        break;
      case 'generateAccounting':
        result = generateAccountingSheet(data.year, data.month);
        break;
      case 'createDashboard':
        result = createDashboard();
        break;
      case 'formatAllSheets':
        formatSalesDataSheet();
        formatMasterSheet();
        result = { success: true, message: 'フォーマット完了' };
        break;
      default:
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

// ========== WebAPI: GET ==========

function doGet(e) {
  try {
    const action = e.parameter.action || 'getTodayStats';
    const manager = new SalesManager(SPREADSHEET_ID);
    let result;

    switch (action) {
      case 'getSheets': {
        const sheets = manager.ss.getSheets().map(function(s) { return s.getName(); });
        result = { sheets: sheets, dataSheetName: manager.dataSheet ? manager.dataSheet.getName() : null };
        break;
      }
      case 'getTodayStats':
        result = manager.getTodayStats();
        break;
      case 'getMonthStats': {
        const year = parseInt(e.parameter.year) || new Date().getFullYear();
        const month = parseInt(e.parameter.month) || new Date().getMonth() + 1;
        result = manager.getMonthStats(year, month);
        break;
      }
      case 'getRecentHistory': {
        const days = parseInt(e.parameter.days) || 30;
        result = { records: manager.getRecentHistory(days) };
        break;
      }
      case 'getCsv': {
        const year = e.parameter.year ? parseInt(e.parameter.year) : null;
        const month = e.parameter.month ? parseInt(e.parameter.month) : null;
        const rows = manager.getCsvData(year, month);
        const csv = rows.map(function(r) { return r.join(','); }).join('\n');
        return ContentService
          .createTextOutput(csv)
          .setMimeType(ContentService.MimeType.CSV);
      }
      case 'getMaster':
        result = { items: getMasterItems() };
        break;
      case 'updateSummary':
        result = updateSummarySheet();
        break;
      case 'generateAccounting': {
        const year = parseInt(e.parameter.year) || new Date().getFullYear();
        const month = parseInt(e.parameter.month) || new Date().getMonth() + 1;
        result = generateAccountingSheet(year, month);
        break;
      }
      case 'createDashboard':
        result = createDashboard();
        break;
      default:
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
    .addItem('🆕 新規売上を記録', 'addNewCustomer')
    .addItem('👥 常連売上を記録', 'addRegularCustomer')
    .addSeparator()
    .addItem('📊 ダッシュボードを更新', 'updateDashboard')
    .addItem('📈 集計シートを更新', 'updateSummaryFromMenu')
    .addItem('📑 経理シートを生成（今月）', 'generateThisMonthAccounting')
    .addSeparator()
    .addItem('⚙️ マスタシートを初期化', 'initMasterFromMenu')
    .addItem('✨ 全シートをフォーマット', 'formatAllSheets')
    .addSeparator()
    .addItem('💾 今すぐバックアップ', 'manualBackupFromSheet')
    .addItem('⏰ 自動バックアップ設定', 'setupDailyBackupTrigger')
    .addToUi();
}

function addNewCustomer() {
  const manager = new SalesManager(SPREADSHEET_ID);
  manager.addSale('新規', 3270);
  try { SpreadsheetApp.getUi().alert('新規 ¥3,270 を登録しました！'); } catch(e) { Logger.log('新規 ¥3,270 を登録しました！'); }
}

function addRegularCustomer() {
  const manager = new SalesManager(SPREADSHEET_ID);
  manager.addSale('常連', 5500);
  try { SpreadsheetApp.getUi().alert('常連 ¥5,500 を登録しました！'); } catch(e) { Logger.log('常連 ¥5,500 を登録しました！'); }
}

function manualBackupFromSheet() {
  const result = runBackup();
  try { SpreadsheetApp.getUi().alert(result.message); } catch(e) { Logger.log(result.message); }
}

function updateDashboard() {
  const result = createDashboard();
  try { SpreadsheetApp.getUi().alert(result.message); } catch(e) { Logger.log(result.message); }
}

function updateSummaryFromMenu() {
  const result = updateSummarySheet();
  try { SpreadsheetApp.getUi().alert(result.message); } catch(e) { Logger.log(result.message); }
}

function generateThisMonthAccounting() {
  const now = new Date();
  const year = now.getFullYear();
  const month = now.getMonth() + 1;
  const result = generateAccountingSheet(year, month);
  try { SpreadsheetApp.getUi().alert(result.message); } catch(e) { Logger.log(result.message); }
}

function initMasterFromMenu() {
  const result = initMasterSheet();
  try { SpreadsheetApp.getUi().alert(result.message); } catch(e) { Logger.log(result.message); }
}
