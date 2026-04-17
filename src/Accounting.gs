// ========== Accounting.gs ==========
// 経理シート生成

function generateAccountingSheet(year, month) {
  try {
    year = parseInt(year);
    month = parseInt(month);
    if (!year || !month || month < 1 || month > 12) {
      return { success: false, message: '年月が不正です' };
    }

    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheetName = '経理_' + year + '_' + String(month).padStart(2, '0');

    var existing = ss.getSheetByName(sheetName);
    if (existing) {
      ss.deleteSheet(existing);
    }
    var accSheet = ss.insertSheet(sheetName);

    // タイトル行
    var titleRange = accSheet.getRange(1, 1, 1, 6);
    titleRange.merge();
    titleRange.setValue('📑 あまと整体院 — ' + year + '年' + month + '月 経理レポート');
    titleRange.setBackground('#1a237e');
    titleRange.setFontColor('#ffffff');
    titleRange.setFontWeight('bold');
    titleRange.setFontSize(14);
    titleRange.setHorizontalAlignment('center');
    titleRange.setVerticalAlignment('middle');
    accSheet.setRowHeight(1, 44);

    var headers = ['日付', '新規件数', '新規売上', '常連件数', '常連売上', '日計'];
    accSheet.getRange(2, 1, 1, headers.length).setValues([headers]);
    var headerRange = accSheet.getRange(2, 1, 1, headers.length);
    headerRange.setBackground('#3949ab');
    headerRange.setFontColor('#ffffff');
    headerRange.setFontWeight('bold');
    headerRange.setFontSize(11);
    headerRange.setHorizontalAlignment('center');
    headerRange.setVerticalAlignment('middle');
    accSheet.setRowHeight(2, 32);
    accSheet.setFrozenRows(2);

    var dataSheet = ss.getSheetByName('売上データ') || ss.getSheets()[0];
    var lastRow = dataSheet.getLastRow();
    var allData = lastRow >= 2 ? dataSheet.getRange(2, 1, lastRow - 1, 5).getValues() : [];

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

    dataRows.push([
      '【月合計】',
      totalShinki,
      totalShinkiSales,
      totalJoren,
      totalJorenSales,
      totalShinkiSales + totalJorenSales
    ]);

    var dataRange = accSheet.getRange(3, 1, dataRows.length, headers.length);
    dataRange.setValues(dataRows);

    for (var r = 0; r < daysInMonth; r++) {
      var rowRange = accSheet.getRange(3 + r, 1, 1, headers.length);
      rowRange.setBackground(r % 2 === 0 ? '#ffffff' : '#e8eaf6');
      rowRange.setFontColor('#212121');
    }

    var totalRowRange = accSheet.getRange(3 + daysInMonth, 1, 1, headers.length);
    totalRowRange.setBackground('#1a237e');
    totalRowRange.setFontColor('#ffffff');
    totalRowRange.setFontWeight('bold');
    totalRowRange.setFontSize(11);

    accSheet.getRange(3, 3, dataRows.length, 1).setNumberFormat('¥#,##0');
    accSheet.getRange(3, 5, dataRows.length, 1).setNumberFormat('¥#,##0');
    accSheet.getRange(3, 6, dataRows.length, 1).setNumberFormat('¥#,##0');

    accSheet.getRange(1, 1, 2 + dataRows.length, headers.length)
      .setBorder(true, true, true, true, true, true, '#9e9e9e', SpreadsheetApp.BorderStyle.SOLID);

    accSheet.autoResizeColumns(1, headers.length);

    return {
      success: true,
      message: sheetName + ' を生成しました（' + daysInMonth + '日分）',
      sheetName: sheetName,
      spreadsheetId: SPREADSHEET_ID,
      spreadsheetUrl: 'https://docs.google.com/spreadsheets/d/' + SPREADSHEET_ID
    };
  } catch (err) {
    return { success: false, message: '経理シート生成エラー: ' + err.message };
  }
}
