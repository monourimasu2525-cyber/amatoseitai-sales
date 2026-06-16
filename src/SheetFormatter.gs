// ========== SheetFormatter.gs ==========

function formatSalesDataSheet() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('売上データ');
  if (!sheet) return;
  var lastRow = sheet.getLastRow(), lastCol = sheet.getLastColumn();
  if (lastCol < 1) return;

  sheet.getRange(1, 1, 1, lastCol)
    .setBackground('#1a237e').setFontColor('#ffffff').setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sheet.setFrozenRows(1);

  if (lastRow >= 2) {
    for (var r = 2; r <= lastRow; r++) {
      sheet.getRange(r, 1, 1, lastCol).setBackground(r % 2 === 0 ? '#e8eaf6' : '#ffffff');
    }
    sheet.getRange(2, 1, lastRow - 1, 1).setNumberFormat('yyyy/MM/dd HH:mm');
    if (lastCol >= 2) sheet.getRange(2, 2, lastRow - 1, 1).setNumberFormat('yyyy/MM/dd HH:mm');
    if (lastCol >= 4) sheet.getRange(2, 4, lastRow - 1, 1).setNumberFormat('¥#,##0');
  }
  sheet.getRange(1, 1, Math.max(lastRow, 1), lastCol)
    .setBorder(true, true, true, true, true, true, '#9e9e9e', SpreadsheetApp.BorderStyle.SOLID);
  sheet.autoResizeColumns(1, lastCol);
}

function formatMasterSheet() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('売上マスタ');
  if (!sheet) return;
  var lastRow = sheet.getLastRow();

  if (lastRow >= 1) {
    sheet.getRange(1, 1, 1, 4)
      .setBackground('#1a237e').setFontColor('#ffffff').setFontWeight('bold')
      .setHorizontalAlignment('center');
    sheet.setFrozenRows(1);
  }
  if (lastRow >= 2) {
    for (var r = 2; r <= lastRow; r++) {
      sheet.getRange(r, 1, 1, 4).setBackground(r % 2 === 0 ? '#e8eaf6' : '#ffffff');
    }
    sheet.getRange(2, 2, lastRow - 1, 1).setNumberFormat('¥#,##0');
  }
  sheet.autoResizeColumns(1, 4);
}

function formatAllSheets() {
  try {
    formatSalesDataSheet();
    formatMasterSheet();
    updateSummarySheet();
    createDashboard();
    try { SpreadsheetApp.getUi().alert('✅ 全シートのフォーマットが完了しました！'); } catch(e) {}
  } catch (err) {
    try { SpreadsheetApp.getUi().alert('❌ エラー: ' + err.message); } catch(e) {}
  }
}
