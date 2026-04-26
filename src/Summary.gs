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

    // ✅ バグ修正: データ検証・コンテンツ・フォーマット・余分な列を全クリア
    summarySheet.clearContents();
    summarySheet.clearFormats();
    summarySheet.clearNotes();
    summarySheet.getRange(1, 1, summarySheet.getMaxRows(), summarySheet.getMaxColumns()).clearDataValidations();

    // 余分な列を削除（H列以降）
    var maxCols = summarySheet.getMaxColumns();
    if (maxCols > 9) {
      summarySheet.deleteColumns(10, maxCols - 9);
    }

    var now = new Date();
    var manager = new SalesManager(SPREADSHEET_ID);

    // ===== タイトルエリア =====
    var titleRange = summarySheet.getRange(1, 1, 1, 9);
    titleRange.merge();
    titleRange.setValue('📊 あまと整体院 売上集計レポート（過去12ヶ月）');
    titleRange.setBackground('#1a237e');
    titleRange.setFontColor('#ffffff');
    titleRange.setFontWeight('bold');
    titleRange.setFontSize(14);
    titleRange.setHorizontalAlignment('center');
    titleRange.setVerticalAlignment('middle');
    summarySheet.setRowHeight(1, 44);

    // 更新日時
    var updateRange = summarySheet.getRange(2, 1, 1, 9);
    updateRange.merge();
    var dateStr = now.getFullYear() + '年' + (now.getMonth()+1) + '月' + now.getDate() + '日 ' +
      String(now.getHours()).padStart(2,'0') + ':' + String(now.getMinutes()).padStart(2,'0') + ' 更新';
    updateRange.setValue('最終更新：' + dateStr);
    updateRange.setBackground('#283593');
    updateRange.setFontColor('#c5cae9');
    updateRange.setFontSize(10);
    updateRange.setHorizontalAlignment('right');
    summarySheet.setRowHeight(2, 22);

    // 空白行
    summarySheet.setRowHeight(3, 12);

    // ===== ヘッダー行（4行目） =====
    var headers = ['年月', '新規件数', '新規売上', '常連件数', '常連売上', '合計件数', '合計売上', '客単価', '前月比'];
    var headerRange = summarySheet.getRange(4, 1, 1, headers.length);
    headerRange.setValues([headers]);
    headerRange.setBackground('#3949ab');
    headerRange.setFontColor('#ffffff');
    headerRange.setFontWeight('bold');
    headerRange.setFontSize(11);
    headerRange.setHorizontalAlignment('center');
    headerRange.setVerticalAlignment('middle');
    summarySheet.setRowHeight(4, 36);
    summarySheet.setFrozenRows(4);

    // ===== データ収集（過去12ヶ月） =====
    var monthData = [];
    var totalShinki = 0, totalShinkiSales = 0, totalJoren = 0, totalJorenSales = 0;
    var totalCount = 0, totalSales = 0;

    for (var i = 11; i >= 0; i--) {
      var d = new Date(now.getFullYear(), now.getMonth() - i, 1);
      var y = d.getFullYear();
      var m = d.getMonth() + 1;
      var stats = manager.getMonthStats(y, m);
      monthData.push({
        label: y + '/' + String(m).padStart(2,'0'),
        shinkiCount: stats.shinkiCount,
        shinkiSales: stats.shinkiSales,
        jorenCount: stats.jorenCount,
        jorenSales: stats.jorenSales,
        totalCount: stats.totalCount,
        totalSales: stats.totalSales
      });
      totalShinki += stats.shinkiCount;
      totalShinkiSales += stats.shinkiSales;
      totalJoren += stats.jorenCount;
      totalJorenSales += stats.jorenSales;
      totalCount += stats.totalCount;
      totalSales += stats.totalSales;
    }

    // ===== データ行書き込み（5行目〜） =====
    var rows = [];
    for (var j = 0; j < monthData.length; j++) {
      var md = monthData[j];
      var tanka = md.totalCount > 0 ? Math.round(md.totalSales / md.totalCount) : 0;
      var zengetsu = j > 0 ? monthData[j-1].totalSales : 0;
      var zengetsuHi = '';
      if (j > 0 && zengetsu > 0) {
        var ratio = ((md.totalSales - zengetsu) / zengetsu * 100).toFixed(1);
        zengetsuHi = (ratio >= 0 ? '+' : '') + ratio + '%';
      } else if (j === 0) {
        zengetsuHi = '-';
      } else {
        zengetsuHi = md.totalSales > 0 ? '新規' : '-';
      }
      rows.push([
        md.label,
        md.shinkiCount,
        md.shinkiSales,
        md.jorenCount,
        md.jorenSales,
        md.totalCount,
        md.totalSales,
        tanka,
        zengetsuHi
      ]);
    }

    // 合計行
    var avgTanka = totalCount > 0 ? Math.round(totalSales / totalCount) : 0;
    rows.push(['【12ヶ月合計】', totalShinki, totalShinkiSales, totalJoren, totalJorenSales, totalCount, totalSales, avgTanka, '']);

    var dataStartRow = 5;
    var dataRange = summarySheet.getRange(dataStartRow, 1, rows.length, headers.length);
    dataRange.setValues(rows);

    // ===== データ行のフォーマット =====
    for (var r = 0; r < rows.length - 1; r++) {
      var rowRange = summarySheet.getRange(dataStartRow + r, 1, 1, headers.length);
      var bgColor = r % 2 === 0 ? '#ffffff' : '#e8eaf6';
      rowRange.setBackground(bgColor);
      rowRange.setFontColor('#212121');
      rowRange.setFontSize(11);
      rowRange.setVerticalAlignment('middle');
      summarySheet.setRowHeight(dataStartRow + r, 28);

      // 前月比の色付け
      var zenHi = rows[r][8];
      var zenCell = summarySheet.getRange(dataStartRow + r, 9);
      if (typeof zenHi === 'string' && zenHi.startsWith('+')) {
        zenCell.setFontColor('#2e7d32');
        zenCell.setFontWeight('bold');
      } else if (typeof zenHi === 'string' && zenHi.startsWith('-') && zenHi !== '-') {
        zenCell.setFontColor('#c62828');
        zenCell.setFontWeight('bold');
      }
    }

    // 合計行フォーマット
    var totalRowRange = summarySheet.getRange(dataStartRow + rows.length - 1, 1, 1, headers.length);
    totalRowRange.setBackground('#1a237e');
    totalRowRange.setFontColor('#ffffff');
    totalRowRange.setFontWeight('bold');
    totalRowRange.setFontSize(11);
    summarySheet.setRowHeight(dataStartRow + rows.length - 1, 32);

    // 数値列フォーマット（通貨）
    var numRows = rows.length;
    summarySheet.getRange(dataStartRow, 3, numRows, 1).setNumberFormat('¥#,##0');
    summarySheet.getRange(dataStartRow, 5, numRows, 1).setNumberFormat('¥#,##0');
    summarySheet.getRange(dataStartRow, 7, numRows, 1).setNumberFormat('¥#,##0');
    summarySheet.getRange(dataStartRow, 8, numRows, 1).setNumberFormat('¥#,##0');

    // 件数列は中央揃え
    summarySheet.getRange(dataStartRow, 2, numRows, 1).setHorizontalAlignment('center');
    summarySheet.getRange(dataStartRow, 4, numRows, 1).setHorizontalAlignment('center');
    summarySheet.getRange(dataStartRow, 6, numRows, 1).setHorizontalAlignment('center');
    summarySheet.getRange(dataStartRow, 9, numRows, 1).setHorizontalAlignment('center');

    // 枠線
    summarySheet.getRange(4, 1, rows.length + 1, headers.length)
      .setBorder(true, true, true, true, true, true, '#9e9e9e', SpreadsheetApp.BorderStyle.SOLID);

    // 列幅調整
    summarySheet.setColumnWidth(1, 90);
    summarySheet.setColumnWidth(2, 80);
    summarySheet.setColumnWidth(3, 100);
    summarySheet.setColumnWidth(4, 80);
    summarySheet.setColumnWidth(5, 100);
    summarySheet.setColumnWidth(6, 80);
    summarySheet.setColumnWidth(7, 110);
    summarySheet.setColumnWidth(8, 90);
    summarySheet.setColumnWidth(9, 80);

    // ===== グラフ（✅ Y軸スケール修正） =====
    var chartDataRows = monthData.length + 1; // ヘッダー + 12ヶ月
    var labelRange = summarySheet.getRange(4, 1, chartDataRows, 1);
    var salesRange = summarySheet.getRange(4, 7, chartDataRows, 1);

    // 最大売上を取得してY軸スケールを設定
    var maxSales = 0;
    for (var k = 0; k < monthData.length; k++) {
      if (monthData[k].totalSales > maxSales) maxSales = monthData[k].totalSales;
    }
    var yMax = maxSales > 0 ? Math.ceil(maxSales * 1.2 / 10000) * 10000 : 100000;

    var chartBuilder = summarySheet.newChart()
      .setChartType(Charts.ChartType.COLUMN)
      .addRange(labelRange)
      .addRange(salesRange)
      .setPosition(dataStartRow + rows.length + 2, 1, 0, 0)
      .setOption('title', '月別合計売上（過去12ヶ月）')
      .setOption('titleTextStyle', { fontSize: 14, bold: true, color: '#1a237e' })
      .setOption('hAxis', {
        title: '年月',
        titleTextStyle: { color: '#424242', bold: true },
        slantedText: true,
        slantedTextAngle: 30
      })
      .setOption('vAxis', {
        title: '合計売上（円）',
        titleTextStyle: { color: '#424242', bold: true },
        format: '¥#,##0',
        minValue: 0,
        viewWindow: { min: 0, max: yMax }
      })
      .setOption('colors', ['#3949ab'])
      .setOption('backgroundColor', '#fafafa')
      .setOption('legend', { position: 'none' })
      .setOption('width', 750)
      .setOption('height', 420);

    summarySheet.insertChart(chartBuilder.build());

    return { success: true, message: '集計シートを更新しました（バグ修正版）' };
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
