// ========== SheetFormatter.gs ==========
// スプレッドシートUI整形

function formatSalesDataSheet() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('売上データ');
  if (!sheet) return;

  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastCol < 1) return;

  var headerRange = sheet.getRange(1, 1, 1, lastCol);
  headerRange.setBackground('#1a237e');
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');
  headerRange.setFontSize(11);
  headerRange.setHorizontalAlignment('center');
  headerRange.setVerticalAlignment('middle');

  sheet.setFrozenRows(1);

  if (lastRow >= 2) {
    for (var r = 2; r <= lastRow; r++) {
      var rowRange = sheet.getRange(r, 1, 1, lastCol);
      rowRange.setBackground(r % 2 === 0 ? '#e8eaf6' : '#ffffff');
      rowRange.setFontColor('#212121');
    }
    sheet.getRange(2, 1, lastRow - 1, 1).setNumberFormat('yyyy/MM/dd HH:mm');
    if (lastCol >= 2) {
      sheet.getRange(2, 2, lastRow - 1, 1).setNumberFormat('yyyy/MM/dd HH:mm');
    }
    if (lastCol >= 4) {
      sheet.getRange(2, 4, lastRow - 1, 1).setNumberFormat('¥#,##0');
    }
  }

  if (lastRow >= 1) {
    var allRange = sheet.getRange(1, 1, Math.max(lastRow, 1), lastCol);
    allRange.setBorder(true, true, true, true, true, true, '#9e9e9e', SpreadsheetApp.BorderStyle.SOLID);
  }

  sheet.autoResizeColumns(1, lastCol);
}

function formatMasterSheet() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('売上マスタ');
  if (!sheet) return;

  var lastRow = sheet.getLastRow();
  var lastCol = 4;

  if (lastRow >= 1) {
    var headerRange = sheet.getRange(1, 1, 1, lastCol);
    headerRange.setBackground('#1a237e');
    headerRange.setFontColor('#ffffff');
    headerRange.setFontWeight('bold');
    headerRange.setFontSize(11);
    headerRange.setHorizontalAlignment('center');
    sheet.setFrozenRows(1);
  }

  if (lastRow >= 2) {
    for (var r = 2; r <= lastRow; r++) {
      var rowRange = sheet.getRange(r, 1, 1, lastCol);
      rowRange.setBackground(r % 2 === 0 ? '#e8eaf6' : '#ffffff');
    }
    sheet.getRange(2, 2, lastRow - 1, 1).setNumberFormat('¥#,##0');
  }

  sheet.autoResizeColumns(1, lastCol);
}

function formatAllSheets() {
  try {
    formatSalesDataSheet();
    formatMasterSheet();
    updateSummarySheet();
    createDashboard();

    try {
      SpreadsheetApp.getUi().alert('✅ 全シートのフォーマットが完了しました！\n\n・売上データシート：ヘッダー強化・交互背景・書式設定\n・売上マスタシート：フォーマット\n・集計シート：グラフ追加・合計行\n・ダッシュボード：今月サマリー更新');
    } catch(e) {
      Logger.log('✅ 全シートのフォーマットが完了しました！');
    }
  } catch (err) {
    try {
      SpreadsheetApp.getUi().alert('❌ フォーマットエラー: ' + err.message);
    } catch(e) {
      Logger.log('❌ フォーマットエラー: ' + err.message);
    }
  }
}
