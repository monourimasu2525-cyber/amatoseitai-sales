// ========== Summary.gs ==========

function updateSummarySheet() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var summarySheet = ss.getSheetByName('集計') || ss.insertSheet('集計');
    summarySheet.getCharts().forEach(function(c) { summarySheet.removeChart(c); });
    summarySheet.clearContents();
    summarySheet.clearFormats();

    var now = new Date();
    var headers = ['年月', '新規件数', '新規売上', '常連件数', '常連売上', '合計件数', '合計売上'];
    var rows = [headers];
    var totals = [0, 0, 0, 0, 0, 0];

    for (var i = 11; i >= 0; i--) {
      var d = new Date(now.getFullYear(), now.getMonth() - i, 1);
      var y = d.getFullYear(), m = d.getMonth() + 1;
      var stats = new SalesManager(SPREADSHEET_ID).getMonthStats(y, m);
      totals[0] += stats.shinkiCount; totals[1] += stats.shinkiSales;
      totals[2] += stats.jorenCount;  totals[3] += stats.jorenSales;
      totals[4] += stats.totalCount;  totals[5] += stats.totalSales;
      rows.push([y + '年' + m + '月', stats.shinkiCount, stats.shinkiSales, stats.jorenCount, stats.jorenSales, stats.totalCount, stats.totalSales]);
    }
    rows.push(['【12ヶ月合計】'].concat(totals));

    summarySheet.getRange(1, 1, rows.length, headers.length).setValues(rows);

    summarySheet.getRange(1, 1, 1, headers.length)
      .setBackground('#1a237e').setFontColor('#ffffff').setFontWeight('bold')
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
    summarySheet.setRowHeight(1, 36);
    summarySheet.setFrozenRows(1);

    for (var r = 2; r <= rows.length - 1; r++) {
      summarySheet.getRange(r, 1, 1, headers.length).setBackground(r % 2 === 0 ? '#e8eaf6' : '#ffffff');
    }
    summarySheet.getRange(rows.length, 1, 1, headers.length)
      .setBackground('#1a237e').setFontColor('#ffffff').setFontWeight('bold');

    summarySheet.getRange(2, 3, rows.length - 1, 1).setNumberFormat('¥#,##0');
    summarySheet.getRange(2, 5, rows.length - 1, 1).setNumberFormat('¥#,##0');
    summarySheet.getRange(2, 7, rows.length - 1, 1).setNumberFormat('¥#,##0');
    summarySheet.getRange(1, 1, rows.length, headers.length)
      .setBorder(true, true, true, true, true, true, '#9e9e9e', SpreadsheetApp.BorderStyle.SOLID);
    summarySheet.autoResizeColumns(1, headers.length);

    var chartBuilder = summarySheet.newChart()
      .setChartType(Charts.ChartType.COLUMN)
      .addRange(summarySheet.getRange(1, 1, rows.length - 1, 1))
      .addRange(summarySheet.getRange(1, 7, rows.length - 1, 1))
      .setPosition(rows.length + 2, 1, 0, 0)
      .setOption('title', '月別合計売上（過去12ヶ月）')
      .setOption('colors', ['#3949ab'])
      .setOption('vAxis', { format: '¥#,##0' })
      .setOption('width', 700).setOption('height', 400)
      .setOption('legend', { position: 'none' });
    summarySheet.insertChart(chartBuilder.build());

    return { success: true, message: '集計シートを更新しました' };
  } catch (err) { return { success: false, message: err.message }; }
}

function createDashboard() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var dashboard = ss.getSheetByName('ダッシュボード');
    if (!dashboard) dashboard = ss.insertSheet('ダッシュボード', 0);
    else { dashboard.clearContents(); dashboard.clearFormats(); }

    var now = new Date();
    var year = now.getFullYear(), month = now.getMonth() + 1;
    var manager = new SalesManager(SPREADSHEET_ID);
    var stats = manager.getMonthStats(year, month);
    var today = manager.getTodayStats();

    dashboard.getRange(1, 1, 50, 10).setBackground('#f5f5f5');

    var mainTitle = dashboard.getRange(1, 1, 2, 8);
    mainTitle.merge().setValue('🏥  あまと整体院  売上ダッシュボード')
      .setBackground('#1a237e').setFontColor('#ffffff').setFontWeight('bold')
      .setFontSize(22).setHorizontalAlignment('center').setVerticalAlignment('middle');
    dashboard.setRowHeight(1, 40); dashboard.setRowHeight(2, 40);

    var dateStr = year + '年' + month + '月' + now.getDate() + '日  ' +
      String(now.getHours()).padStart(2,'0') + ':' + String(now.getMinutes()).padStart(2,'0') + ' 更新';
    dashboard.getRange(3, 1, 1, 8).merge().setValue('最終更新：' + dateStr)
      .setBackground('#283593').setFontColor('#c5cae9').setFontSize(10)
      .setHorizontalAlignment('right').setVerticalAlignment('middle');
    dashboard.setRowHeight(3, 24);

    dashboard.setRowHeight(4, 16);
    dashboard.getRange(5, 1, 1, 8).merge()
      .setValue('📊  ' + year + '年' + month + '月  月次サマリー')
      .setBackground('#e8eaf6').setFontColor('#1a237e').setFontWeight('bold')
      .setFontSize(14).setHorizontalAlignment('left').setVerticalAlignment('middle');
    dashboard.setRowHeight(5, 36);
    dashboard.setRowHeight(6, 12);

    _drawCard(dashboard, 7,  2, '🆕  新規', stats.shinkiCount + ' 件', '¥' + stats.shinkiSales.toLocaleString(), '#e3f2fd', '#1565c0');
    _drawCard(dashboard, 7,  6, '👥  常連', stats.jorenCount  + ' 件', '¥' + stats.jorenSales.toLocaleString(),  '#e8f5e9', '#2e7d32');
    dashboard.setRowHeight(12, 12);
    _drawCard(dashboard, 13, 2, '💰  今月合計', stats.totalCount + ' 件', '¥' + stats.totalSales.toLocaleString(), '#fff8e1', '#f57f17');
    _drawCard(dashboard, 13, 6, '📅  本日',     today.totalCount + ' 件', '¥' + today.totalSales.toLocaleString(), '#fce4ec', '#880e4f');

    [1,2,3,4,5,6,7,8].forEach(function(c, i) {
      var widths = [20, 180, 120, 150, 20, 180, 120, 150];
      dashboard.setColumnWidth(i + 1, widths[i]);
    });

    return { success: true, message: 'ダッシュボードを更新しました' };
  } catch (err) { return { success: false, message: err.message }; }
}

function _drawCard(sheet, startRow, startCol, title, count, amount, bgColor, accentColor) {
  var cardRange = sheet.getRange(startRow, startCol, 5, 3);
  cardRange.setBackground(bgColor)
    .setBorder(true, true, true, true, false, false, accentColor, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  sheet.getRange(startRow, startCol, 1, 3).merge().setValue(title)
    .setBackground(accentColor).setFontColor('#ffffff').setFontWeight('bold')
    .setFontSize(13).setHorizontalAlignment('center').setVerticalAlignment('middle');
  sheet.setRowHeight(startRow, 36);

  sheet.getRange(startRow + 1, startCol, 2, 3).merge().setValue(count)
    .setFontColor(accentColor).setFontWeight('bold').setFontSize(28)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sheet.setRowHeight(startRow + 1, 44); sheet.setRowHeight(startRow + 2, 44);

  sheet.getRange(startRow + 3, startCol, 2, 3).merge().setValue(amount)
    .setFontColor('#212121').setFontWeight('bold').setFontSize(22)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sheet.setRowHeight(startRow + 3, 40); sheet.setRowHeight(startRow + 4, 40);
}
