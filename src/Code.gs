 
var SPREADSHEET_ID = '17bAyQngDEjoDgqSLLUU5p45HWXomF09bLf_h6FySsjs';
var CALENDAR_ID = 'primary';

class SalesManager {
  constructor(spreadsheetId) {
    this.ss = SpreadsheetApp.openById(spreadsheetId);
    this.dataSheet = this.ss.getSheetByName('売上データ');
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
      return { success: true, message: `${type} ¥${amount} を登録しました`, timestamp: now.toISOString(), type: type, amount: amount };
    } catch (err) {
      return { success: false, message: `エラー: ${err.message}` };
    }
  }
  
  getTodayStats() {
    const today = new Date();
    const year = today.getFullYear();
    const month = today.getMonth() + 1;
    const day = today.getDate();
    const lastRow = this.dataSheet.getLastRow();
    if (lastRow < 2) {
      return { date: `${year}年${month}月${day}日`, shinkiCount: 0, jorenCount: 0, totalCount: 0, shinkiSales: 0, jorenSales: 0, totalSales: 0 };
    }
    const allData = this.dataSheet.getRange(2, 1, lastRow - 1, 5).getValues();
    let shinkiCount = 0, shinkiSales = 0, jorenCount = 0, jorenSales = 0;
    allData.forEach(row => {
      if (!row[0]) return;
      const date = new Date(row[0]);
      if (date.getFullYear() === year && date.getMonth() + 1 === month && date.getDate() === day) {
        const amount = Number(row[3]) || 0;
        if (row[2] === '新規') { shinkiCount++; shinkiSales += amount; }
        else if (row[2] === '常連') { jorenCount++; jorenSales += amount; }
      }
    });
    return { date: `${year}年${month}月${day}日`, shinkiCount, jorenCount, totalCount: shinkiCount + jorenCount, shinkiSales, jorenSales, totalSales: shinkiSales + jorenSales };
  }
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const manager = new SalesManager(SPREADSHEET_ID);
    const result = manager.addSale(data.type, data.amount);
    const response = ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
    
    // ✅ CORS対応
    response.addHeader("Access-Control-Allow-Origin", "*");
    response.addHeader("Access-Control-Allow-Methods", "GET, POST, OPTIONS");
    response.addHeader("Access-Control-Allow-Headers", "Content-Type");
    
    return response;
  } catch (err) {
    const response = ContentService.createTextOutput(JSON.stringify({ success: false, message: err.message })).setMimeType(ContentService.MimeType.JSON);
    response.addHeader("Access-Control-Allow-Origin", "*");
    return response;
  }
}

function doGet(e) {
  try {
    const action = e.parameter.action;
    const manager = new SalesManager(SPREADSHEET_ID);
    const result = manager.getTodayStats();
    const response = ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
    
    // ✅ CORS対応
    response.addHeader("Access-Control-Allow-Origin", "*");
    response.addHeader("Access-Control-Allow-Methods", "GET, POST, OPTIONS");
    response.addHeader("Access-Control-Allow-Headers", "Content-Type");
    
    return response;
  } catch (err) {
    const response = ContentService.createTextOutput(JSON.stringify({ success: false, message: err.message })).setMimeType(ContentService.MimeType.JSON);
    response.addHeader("Access-Control-Allow-Origin", "*");
    return response;
  }
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('売上管理')
    .addItem('新規売上を記録', 'addNewCustomer')
    .addItem('常連売上を記録', 'addRegularCustomer')
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
