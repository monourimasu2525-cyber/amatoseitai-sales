// ========== Summary.gs ==========
// 集計・ダッシュボード

function updateSummarySheet() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var summarySheet = ss.getSheetByName('集計');
    if (!summarySheet) {
      summarySheet = ss.insertSheet('集計');
    }

    // 既存グラフを全削除
    var charts = summarySheet.getCharts();
    charts.forEach(function(chart) {
      summarySheet.removeChart(chart);
    });

    summarySheet.clearContents();
    summarySheet.clearFormats();

    var now = new Date();
    var headers = ['年月', '新規件数', '新規売上', '常連件数', '常連売上', '合計件数', '合計売上'];
    var rows = [headers];

    var totalShinki = 0, totalShinkiSales = 0, totalJoren = 0, totalJorenSales = 0, totalCount = 0, totalSales = 0;

    for (var i = 11; i >= 0; i--) {
      var d = new Date(now.getFullYear(), now.getMonth() - i, 1);
      var y = d.getFullYear();
      var m = d.getMonth() + 1;
      var manager = new SalesManager(SPREADSHEET_ID);
      var stats = manager.getMonthStats(y, m);
      totalShinki += stats.shinkiCount;
      totalShinkiSales += stats.shinkiSales;
      totalJoren += stats.jorenCount;
      totalJorenSales += stats.jorenSales;
      totalCount += stats.totalCount;
      totalSales += stats.totalSales;
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

    rows.push(['【12ヶ月合計】', totalShinki, totalShinkiSales, totalJoren, totalJorenSales, totalCount, totalSales]);

    var range = summarySheet.getRange(1, 1, rows.length, headers.length);
    range.setValues(rows);

    // ヘッダー行書式
    var headerRange = summarySheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#1a237e');
    headerRange.setFontColor('#ffffff');
    headerRange.setFontWeight('bold');
    headerRange.setFontSize(11);
    headerRange.setHorizontalAlignment('center');
    headerRange.setVerticalAlignment('middle');
    summarySheet.setRowHeight(1, 36);
    summarySheet.setFrozenRows(1);

    // データ行の交互背景色
    for (var r = 2; r <= rows.length - 1; r++) {
      var rowRange = summarySheet.getRange(r, 1, 1, headers.length);
      rowRange.setBackground(r % 2 === 0 ? '#e8eaf6' : '#ffffff');
      rowRange.setFontColor('#212121');
    }

    // 合計行書式
    var totalRowRange = summarySheet.getRange(rows.length, 1, 1, headers.length);
    totalRowRange.setBackground('#1a237e');
    totalRowRange.setFontColor('#ffffff');
    totalRowRange.setFontWeight('bold');
    totalRowRange.setFontSize(11);

    // 数値列を通貨フォーマット
    summarySheet.getRange(2, 3, rows.length - 1, 1).setNumberFormat('¥#,##0');
    summarySheet.getRange(2, 5, rows.length - 1, 1).setNumberFormat('¥#,##0');
    summarySheet.getRange(2, 7, rows.length - 1, 1).setNumberFormat('¥#,##0');

    // 枠線
    summarySheet.getRange(1, 1, rows.length, headers.length)
      .setBorder(true, true, true, true, true, true, '#9e9e9e', SpreadsheetApp.BorderStyle.SOLID);

    summarySheet.autoResizeColumns(1, headers.length);

    // 棒グラフ（月別合計売上）
    var dataRowCount = rows.length - 1;
    var chartRange = summarySheet.getRange(1, 1, dataRowCount, 1);
    var chartRange2 = summarySheet.getRange(1, 7, dataRowCount, 1);

    var chartBuilder = summarySheet.newChart()
      .setChartType(Charts.ChartType.COLUMN)
      .addRange(chartRange)
      .addRange(chartRange2)
      .setPosition(rows.length + 2, 1, 0, 0)
      .setOption('title', '月別合計売上（過去12ヶ月）')
      .setOption('titleTextStyle', { fontSize: 14, bold: true, color: '#1a237e' })
      .setOption('hAxis', { title: '年月', titleTextStyle: { color: '#424242', bold: true } })
      .setOption('vAxis', { title: '合計売上（円）', titleTextStyle: { color: '#424242', bold: true }, format: '¥#,##0' })
      .setOption('colors', ['#3949ab'])
      .setOption('backgroundColor', '#fafafa')
      .setOption('legend', { position: 'none' })
      .setOption('width', 700)
      .setOption('height', 400);

    summarySheet.insertChart(chartBuilder.build());

    return { success: true, message: '集計シートを更新しました（過去12ヶ月 + グラフ）' };
  } catch (err) {
    return { success: false, message: '集計シート更新エラー: ' + err.message };
  }
}

function createDashboard() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var dashboard = ss.getSheetByName('ダッシュボード');
    if (!dashboard) {
      dashboard = ss.insertSheet('ダッシュボード', 0);
    } else {
      dashboard.clearContents();
      dashboard.clearFormats();
    }

    var now = new Date();
    var year = now.getFullYear();
    var month = now.getMonth() + 1;
    var manager = new SalesManager(SPREADSHEET_ID);
    var stats = manager.getMonthStats(year, month);
    var todayStats = manager.getTodayStats();

    dashboard.getRange(1, 1, 50, 10).setBackground('#f5f5f5');

    // タイトル
    var mainTitle = dashboard.getRange(1, 1, 2, 8);
    mainTitle.merge();
    mainTitle.setValue('🏥  あまと整体院  売上ダッシュボード');
    mainTitle.setBackground('#1a237e');
    mainTitle.setFontColor('#ffffff');
    mainTitle.setFontWeight('bold');
    mainTitle.setFontSize(22);
    mainTitle.setHorizontalAlignment('center');
    mainTitle.setVerticalAlignment('middle');
    dashboard.setRowHeight(1, 40);
    dashboard.setRowHeight(2, 40);

    // 更新日時
    var updateTime = dashboard.getRange(3, 1, 1, 8);
    updateTime.merge();
    var dateStr = year + '年' + month + '月' + now.getDate() + '日  ' +
      String(now.getHours()).padStart(2,'0') + ':' + String(now.getMinutes()).padStart(2,'0') + ' 更新';
    updateTime.setValue('最終更新：' + dateStr);
    updateTime.setBackground('#283593');
    updateTime.setFontColor('#c5cae9');
    updateTime.setFontSize(10);
    updateTime.setHorizontalAlignment('right');
    updateTime.setVerticalAlignment('middle');
    dashboard.setRowHeight(3, 24);

    // 月次サマリー
    dashboard.setRowHeight(4, 16);
    var sectionTitle = dashboard.getRange(5, 1, 1, 8);
    sectionTitle.merge();
    sectionTitle.setValue('📊  ' + year + '年' + month + '月  月次サマリー');
    sectionTitle.setBackground('#e8eaf6');
    sectionTitle.setFontColor('#1a237e');
    sectionTitle.setFontWeight('bold');
    sectionTitle.setFontSize(14);
    sectionTitle.setHorizontalAlignment('left');
    sectionTitle.setVerticalAlignment('middle');
    sectionTitle.setBorder(true, true, true, true, false, false, '#9fa8da', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    dashboard.setRowHeight(5, 36);

    dashboard.setRowHeight(6, 12);
    _drawCard(dashboard, 7, 2, '🆕  新規', stats.shinkiCount + ' 件', '¥' + stats.shinkiSales.toLocaleString(), '#e3f2fd', '#1565c0');
    _drawCard(dashboard, 7, 6, '👥  常連', stats.jorenCount + ' 件', '¥' + stats.jorenSales.toLocaleString(), '#e8f5e9', '#2e7d32');

    dashboard.setRowHeight(12, 12);
    _drawCard(dashboard, 13, 2, '💰  今月合計', stats.totalCount + ' 件', '¥' + stats.totalSales.toLocaleString(), '#fff8e1', '#f57f17');
    _drawCard(dashboard, 13, 6, '📅  本日', todayStats.totalCount + ' 件', '¥' + todayStats.totalSales.toLocaleString(), '#fce4ec', '#880e4f');

    dashboard.setRowHeight(17, 12);

    // 今日詳細
    var todaySection = dashboard.getRange(18, 1, 1, 8);
    todaySection.merge();
    todaySection.setValue('📅  本日詳細（' + todayStats.date + '）');
    todaySection.setBackground('#e8eaf6');
    todaySection.setFontColor('#1a237e');
    todaySection.setFontWeight('bold');
    todaySection.setFontSize(13);
    todaySection.setHorizontalAlignment('left');
    todaySection.setVerticalAlignment('middle');
    todaySection.setBorder(true, true, true, true, false, false, '#9fa8da', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    dashboard.setRowHeight(18, 32);

    var todayHeader = dashboard.getRange(19, 2, 1, 3);
    todayHeader.setValues([['種別', '件数', '売上']]);
    todayHeader.setBackground('#3949ab');
    todayHeader.setFontColor('#ffffff');
    todayHeader.setFontWeight('bold');
    todayHeader.setFontSize(11);
    todayHeader.setHorizontalAlignment('center');

    var todayRows = [
      ['🆕 新規', todayStats.shinkiCount, todayStats.shinkiSales],
      ['👥 常連', todayStats.jorenCount, todayStats.jorenSales],
      ['💰 合計', todayStats.totalCount, todayStats.totalSales]
    ];
    var todayDataRange = dashboard.getRange(20, 2, 3, 3);
    todayDataRange.setValues(todayRows);
    todayDataRange.setBackground('#ffffff');
    todayDataRange.setFontSize(12);
    todayDataRange.setBorder(true, true, true, true, true, true, '#9e9e9e', SpreadsheetApp.BorderStyle.SOLID);

    var todayTotalRow = dashboard.getRange(22, 2, 1, 3);
    todayTotalRow.setBackground('#fff9c4');
    todayTotalRow.setFontWeight('bold');
    todayTotalRow.setFontSize(13);

    dashboard.getRange(19, 4, 4, 1).setNumberFormat('¥#,##0');

    dashboard.setColumnWidth(1, 20);
    dashboard.setColumnWidth(2, 180);
    dashboard.setColumnWidth(3, 120);
    dashboard.setColumnWidth(4, 150);
    dashboard.setColumnWidth(5, 20);
    dashboard.setColumnWidth(6, 180);
    dashboard.setColumnWidth(7, 120);
    dashboard.setColumnWidth(8, 150);

    return { success: true, message: 'ダッシュボードを作成・更新しました（' + year + '年' + month + '月）' };
  } catch (err) {
    return { success: false, message: 'ダッシュボード作成エラー: ' + err.message };
  }
}

function _drawCard(sheet, startRow, startCol, title, count, amount, bgColor, accentColor) {
  var cardRows = 5;
  var cardRange = sheet.getRange(startRow, startCol, cardRows, 3);
  cardRange.setBackground(bgColor);
  cardRange.setBorder(true, true, true, true, false, false, accentColor, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  var titleCell = sheet.getRange(startRow, startCol, 1, 3);
  titleCell.merge();
  titleCell.setValue(title);
  titleCell.setBackground(accentColor);
  titleCell.setFontColor('#ffffff');
  titleCell.setFontWeight('bold');
  titleCell.setFontSize(13);
  titleCell.setHorizontalAlignment('center');
  titleCell.setVerticalAlignment('middle');
  sheet.setRowHeight(startRow, 36);

  var countCell = sheet.getRange(startRow + 1, startCol, 2, 3);
  countCell.merge();
  countCell.setValue(count);
  countCell.setBackground(bgColor);
  countCell.setFontColor(accentColor);
  countCell.setFontWeight('bold');
  countCell.setFontSize(28);
  countCell.setHorizontalAlignment('center');
  countCell.setVerticalAlignment('middle');
  sheet.setRowHeight(startRow + 1, 44);
  sheet.setRowHeight(startRow + 2, 44);

  var amountCell = sheet.getRange(startRow + 3, startCol, 2, 3);
  amountCell.merge();
  amountCell.setValue(amount);
  amountCell.setBackground(bgColor);
  amountCell.setFontColor('#212121');
  amountCell.setFontWeight('bold');
  amountCell.setFontSize(22);
  amountCell.setHorizontalAlignment('center');
  amountCell.setVerticalAlignment('middle');
  sheet.setRowHeight(startRow + 3, 40);
  sheet.setRowHeight(startRow + 4, 40);
}
