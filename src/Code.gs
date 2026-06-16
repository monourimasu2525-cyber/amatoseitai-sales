// ========== Code.gs ==========
var SPREADSHEET_ID = '17bAyQngDEjoDgqSLLUU5p45HWXomF09bLf_h6FySsjs';

function doGet(e) {
  try {
    var action = e.parameter.action || 'getTodayStats';
    var manager = new SalesManager(SPREADSHEET_ID);
    var result;
    switch (action) {
      case 'initData': {
        var now = new Date();
        var curYear = now.getFullYear(), curMonth = now.getMonth() + 1;
        var prevMonth = curMonth === 1 ? 12 : curMonth - 1;
        var prevYear  = curMonth === 1 ? curYear - 1 : curYear;
        result = {
          master:     getMasterItems(),
          todayStats: manager.getTodayStats(),
          thisMonth:  manager.getMonthStats(curYear, curMonth),
          prevMonth:  manager.getMonthStats(prevYear, prevMonth),
          history:    { records: manager.getRecentHistory(30) }
        };
        break;
      }
      case 'getTodayStats':
        result = manager.getTodayStats(); break;
      case 'getMonthStats':
        result = manager.getMonthStats(
          parseInt(e.parameter.year)  || new Date().getFullYear(),
          parseInt(e.parameter.month) || new Date().getMonth() + 1
        ); break;
      case 'getRecentHistory':
        result = { records: manager.getRecentHistory(parseInt(e.parameter.days) || 30) }; break;
      case 'getMaster':
        result = { items: getMasterItems() }; break;
      case 'getCsv': {
        var rows = manager.getCsvData(
          e.parameter.year  ? parseInt(e.parameter.year)  : null,
          e.parameter.month ? parseInt(e.parameter.month) : null
        );
        return ContentService.createTextOutput(rows.map(function(r){return r.join(',');}).join('\n'))
          .setMimeType(ContentService.MimeType.CSV);
      }
      case 'addSale':
        result = manager.addSale(e.parameter.type, parseFloat(e.parameter.amount)); break;
      case 'editSale':
        result = manager.editSale(
          parseInt(e.parameter.rowIndex), e.parameter.type, parseFloat(e.parameter.amount)
        ); break;
      case 'deleteSale':
        result = manager.deleteSale(parseInt(e.parameter.rowIndex)); break;
      case 'addMaster':
        result = addMasterItem(e.parameter.type, parseFloat(e.parameter.amount), e.parameter.description || ''); break;
      case 'updateMaster':
        result = updateMasterItem(
          parseInt(e.parameter.rowIndex), e.parameter.type,
          parseFloat(e.parameter.amount), e.parameter.description || ''
        ); break;
      case 'deleteMaster':
        result = deleteMasterItem(parseInt(e.parameter.rowIndex)); break;
      case 'generateAccounting':
        result = generateAccountingSheet(
          parseInt(e.parameter.year)  || new Date().getFullYear(),
          parseInt(e.parameter.month) || new Date().getMonth() + 1
        ); break;
      case 'updateSummary':
        result = updateSummarySheet(); break;
      case 'createDashboard':
        result = createDashboard(); break;
      default:
        result = { success: false, message: '不明なアクション: ' + action };
    }
    return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('売上管理')
    .addItem('新規売上を記録', 'addNewCustomer')
    .addItem('常連売上を記録', 'addRegularCustomer')
    .addSeparator()
    .addItem('ダッシュボードを更新', 'updateDashboard')
    .addItem('集計シートを更新', 'updateSummaryFromMenu')
    .addItem('経理シートを生成（今月）', 'generateThisMonthAccounting')
    .addSeparator()
    .addItem('マスタシートを初期化', 'initMasterFromMenu')
    .addItem('全シートをフォーマット', 'formatAllSheets')
    .addToUi();
}

function addNewCustomer() {
  new SalesManager(SPREADSHEET_ID).addSale('新規', 3270);
  try { SpreadsheetApp.getUi().alert('新規 ¥3,270 を登録しました！'); } catch(e) {}
}
function addRegularCustomer() {
  new SalesManager(SPREADSHEET_ID).addSale('常連', 5500);
  try { SpreadsheetApp.getUi().alert('常連 ¥5,500 を登録しました！'); } catch(e) {}
}
function updateDashboard() { var r = createDashboard(); try { SpreadsheetApp.getUi().alert(r.message); } catch(e) {} }
function updateSummaryFromMenu() { var r = updateSummarySheet(); try { SpreadsheetApp.getUi().alert(r.message); } catch(e) {} }
function generateThisMonthAccounting() {
  var now = new Date();
  var r = generateAccountingSheet(now.getFullYear(), now.getMonth() + 1);
  try { SpreadsheetApp.getUi().alert(r.message); } catch(e) {}
}
function initMasterFromMenu() { var r = initMasterSheet(); try { SpreadsheetApp.getUi().alert(r.message); } catch(e) {} }
