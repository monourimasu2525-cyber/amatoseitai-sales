// ========== Accounting.gs ==========

function generateAccountingSheet(year, month) {
  try {
    year = parseInt(year); month = parseInt(month);
    if (!year || !month || month < 1 || month > 12) return { success: false, message: '年月が不正です' };

    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheetName = '経理_' + year + '_' + String(month).padStart(2, '0');
    var existing = ss.getSheetByName(sheetName);
    if (existing) ss.deleteSheet(existing);
    var accSheet = ss.insertSheet(sheetName);

    accSheet.getRange(1, 1, 1, 6).merge()
      .setValue('📑 あまと整体院 — ' + year + '年' + month + '月 経理レポート')
      .setBackground('#1a237e').setFontColor('#ffffff').setFontWeight('bold')
      .setFontSize(14).setHorizontalAlignment('center').setVerticalAlignment('middle');
    accSheet.setRowHeight(1, 44);

    var headers = ['日付', '新規件数', '新規売上', '常連件数', '常連売上', '日計'];
    accSheet.getRange(2, 1, 1, headers.length).setValues([headers])
      .setBackground('#3949ab').setFontColor('#ffffff').setFontWeight('bold')
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
    accSheet.setRowHeight(2, 32);
    accSheet.setFrozenRows(2);

    var dataSheet = ss.getSheetByName('売上データ') || ss.getSheets()[0];
    var lastRow = dataSheet.getLastRow();
    var allData = lastRow >= 2 ? dataSheet.getRange(2, 1, lastRow - 1, 5).getValues() : [];

    var daysInMonth = new Date(year, month, 0).getDate();
    var dayMap = {};
    for (var d = 1; d <= daysInMonth; d++) dayMap[d] = { shinkiCount: 0, shinkiSales: 0, jorenCount: 0, jorenSales: 0 };

    allData.forEach(function(row) {
      if (!row[0]) return;
      var dt = new Date(row[0]);
      if (dt.getFullYear() === year && dt.getMonth() + 1 === month) {
        var day = dt.getDate(), amt = Number(row[3]) || 0;
        if (row[2] === '新規')      { dayMap[day].shinkiCount++; dayMap[day].shinkiSales += amt; }
        else if (row[2] === '常連') { dayMap[day].jorenCount++;  dayMap[day].jorenSales  += amt; }
      }
    });

    var dataRows = [], totals = [0, 0, 0, 0];
    for (var day = 1; day <= daysInMonth; day++) {
      var dk = dayMap[day];
      totals[0] += dk.shinkiCount; totals[1] += dk.shinkiSales;
      totals[2] += dk.jorenCount;  totals[3] += dk.jorenSales;
      dataRows.push([year + '/' + String(month).padStart(2,'0') + '/' + String(day).padStart(2,'0'),
        dk.shinkiCount, dk.shinkiSales, dk.jorenCount, dk.jorenSales, dk.shinkiSales + dk.jorenSales]);
    }
    dataRows.push(['【月合計】', totals[0], totals[1], totals[2], totals[3], totals[1] + totals[3]]);

    accSheet.getRange(3, 1, dataRows.length, headers.length).setValues(dataRows);
    for (var r = 0; r < daysInMonth; r++) {
      accSheet.getRange(3 + r, 1, 1, headers.length).setBackground(r % 2 === 0 ? '#ffffff' : '#e8eaf6');
    }
    accSheet.getRange(3 + daysInMonth, 1, 1, headers.length)
      .setBackground('#1a237e').setFontColor('#ffffff').setFontWeight('bold');

    accSheet.getRange(3, 3, dataRows.length, 1).setNumberFormat('¥#,##0');
    accSheet.getRange(3, 5, dataRows.length, 1).setNumberFormat('¥#,##0');
    accSheet.getRange(3, 6, dataRows.length, 1).setNumberFormat('¥#,##0');
    accSheet.getRange(1, 1, 2 + dataRows.length, headers.length)
      .setBorder(true, true, true, true, true, true, '#9e9e9e', SpreadsheetApp.BorderStyle.SOLID);
    accSheet.autoResizeColumns(1, headers.length);

    return { success: true, message: sheetName + ' を生成しました（' + daysInMonth + '日分）',
      sheetName, spreadsheetUrl: 'https://docs.google.com/spreadsheets/d/' + SPREADSHEET_ID };
  } catch (err) { return { success: false, message: err.message }; }
}
