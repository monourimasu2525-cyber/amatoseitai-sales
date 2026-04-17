var SPREADSHEET_ID = '17bAyQngDEjoDgqSLLUU5p45HWXomF09bLf_h6FySsjs';
var BACKUP_FOLDER_NAME = 'あまと整体院_売上バックアップ';

// ========== SalesManager ==========

class SalesManager {
  constructor(spreadsheetId) {
    this.ss = SpreadsheetApp.openById(spreadsheetId);
    this.dataSheet = this.ss.getSheetByName('売上データ') || this.ss.getSheets()[0];
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
      return { success: true, message: type + ' ¥' + amount + ' を登録しました', timestamp: now.toISOString(), type: type, amount: amount };
    } catch (err) {
      return { success: false, message: 'エラー: ' + err.message };
    }
  }

  getTodayStats() {
    const today = new Date();
    const year = today.getFullYear();
    const month = today.getMonth() + 1;
    const day = today.getDate();
    const lastRow = this.dataSheet.getLastRow();
    if (lastRow < 2) {
      return { date: year + '年' + month + '月' + day + '日', shinkiCount: 0, jorenCount: 0, totalCount: 0, shinkiSales: 0, jorenSales: 0, totalSales: 0 };
    }
    const allData = this.dataSheet.getRange(2, 1, lastRow - 1, 5).getValues();
    let shinkiCount = 0, shinkiSales = 0, jorenCount = 0, jorenSales = 0;
    allData.forEach(function(row) {
      if (!row[0]) return;
      const date = new Date(row[0]);
      if (date.getFullYear() === year && date.getMonth() + 1 === month && date.getDate() === day) {
        const amount = Number(row[3]) || 0;
        if (row[2] === '新規') { shinkiCount++; shinkiSales += amount; }
        else if (row[2] === '常連') { jorenCount++; jorenSales += amount; }
      }
    });
    return { date: year + '年' + month + '月' + day + '日', shinkiCount: shinkiCount, jorenCount: jorenCount, totalCount: shinkiCount + jorenCount, shinkiSales: shinkiSales, jorenSales: jorenSales, totalSales: shinkiSales + jorenSales };
  }

  // 月間集計（先月比・前年度比用）
  getMonthStats(year, month) {
    const lastRow = this.dataSheet.getLastRow();
    if (lastRow < 2) return { shinkiCount: 0, jorenCount: 0, totalCount: 0, shinkiSales: 0, jorenSales: 0, totalSales: 0 };
    const allData = this.dataSheet.getRange(2, 1, lastRow - 1, 5).getValues();
    let shinkiCount = 0, shinkiSales = 0, jorenCount = 0, jorenSales = 0;
    allData.forEach(function(row) {
      if (!row[0]) return;
      const date = new Date(row[0]);
      if (date.getFullYear() === year && date.getMonth() + 1 === month) {
        const amount = Number(row[3]) || 0;
        if (row[2] === '新規') { shinkiCount++; shinkiSales += amount; }
        else if (row[2] === '常連') { jorenCount++; jorenSales += amount; }
      }
    });
    return { shinkiCount: shinkiCount, jorenCount: jorenCount, totalCount: shinkiCount + jorenCount, shinkiSales: shinkiSales, jorenSales: jorenSales, totalSales: shinkiSales + jorenSales };
  }

  // 履歴取得（直近N日分）
  getRecentHistory(days) {
    const lastRow = this.dataSheet.getLastRow();
    if (lastRow < 2) return [];
    const allData = this.dataSheet.getRange(2, 1, lastRow - 1, 5).getValues();
    const cutoff = new Date();
    cutoff.setDate(cutoff.getDate() - days);
    const result = [];
    allData.forEach(function(row) {
      if (!row[0]) return;
      const date = new Date(row[0]);
      if (date >= cutoff) {
        result.push({
          date: date.getFullYear() + '/' + (date.getMonth()+1) + '/' + date.getDate(),
          time: date.getHours() + ':' + String(date.getMinutes()).padStart(2,'0'),
          type: row[2],
          amount: Number(row[3]) || 0
        });
      }
    });
    return result.reverse();
  }

  // CSV出力用データ取得
  getCsvData(year, month) {
    const lastRow = this.dataSheet.getLastRow();
    if (lastRow < 2) return [];
    const allData = this.dataSheet.getRange(2, 1, lastRow - 1, 5).getValues();
    const result = [['日付', '時刻', '種別', '金額', '入力方法']];
    allData.forEach(function(row) {
      if (!row[0]) return;
      const date = new Date(row[0]);
      if (!year || (date.getFullYear() === year && date.getMonth() + 1 === month)) {
        result.push([
          date.getFullYear() + '/' + (date.getMonth()+1) + '/' + date.getDate(),
          date.getHours() + ':' + String(date.getMinutes()).padStart(2,'0'),
          row[2],
          Number(row[3]) || 0,
          row[4] || 'WebAPI'
        ]);
      }
    });
    return result;
  }
}

// ========== バックアップ ==========

function getOrCreateBackupFolder() {
  const folders = DriveApp.getFoldersByName(BACKUP_FOLDER_NAME);
  if (folders.hasNext()) {
    return folders.next();
  }
  return DriveApp.createFolder(BACKUP_FOLDER_NAME);
}

function runBackup() {
  try {
    const folder = getOrCreateBackupFolder();
    const now = new Date();
    const label = now.getFullYear() + '-'
      + String(now.getMonth()+1).padStart(2,'0') + '-'
      + String(now.getDate()).padStart(2,'0')
      + '_' + String(now.getHours()).padStart(2,'0')
      + String(now.getMinutes()).padStart(2,'0');
    const fileName = 'あまと整体院_売上データ_' + label;
    const original = DriveApp.getFileById(SPREADSHEET_ID);
    const copy = original.makeCopy(fileName, folder);
    return { success: true, message: 'バックアップ完了: ' + fileName, fileId: copy.getId() };
  } catch (err) {
    return { success: false, message: 'バックアップエラー: ' + err.message };
  }
}

// 自動バックアップ用（トリガーから呼ばれる）
function dailyAutoBackup() {
  runBackup();
}

// トリガーセットアップ（初回1回だけ手動で実行）
function setupDailyBackupTrigger() {
  // 既存トリガー削除
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'dailyAutoBackup') {
      ScriptApp.deleteTrigger(t);
    }
  });
  // 毎日深夜2時に実行
  ScriptApp.newTrigger('dailyAutoBackup')
    .timeBased()
    .everyDays(1)
    .atHour(2)
    .create();
}

// ========== 売上データシートのUI強化 ==========

function formatSalesDataSheet() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('売上データ');
  if (!sheet) return;

  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastCol < 1) return;

  // ヘッダー行（1行目）のフォーマット
  var headerRange = sheet.getRange(1, 1, 1, lastCol);
  headerRange.setBackground('#1a237e');
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');
  headerRange.setFontSize(11);
  headerRange.setHorizontalAlignment('center');
  headerRange.setVerticalAlignment('middle');

  // 1行目を固定
  sheet.setFrozenRows(1);

  // データ行の交互背景色
  if (lastRow >= 2) {
    for (var r = 2; r <= lastRow; r++) {
      var rowRange = sheet.getRange(r, 1, 1, lastCol);
      rowRange.setBackground(r % 2 === 0 ? '#e8eaf6' : '#ffffff');
      rowRange.setFontColor('#212121');
    }

    // 日付列フォーマット（1列目・2列目）
    sheet.getRange(2, 1, lastRow - 1, 1).setNumberFormat('yyyy/MM/dd HH:mm');
    if (lastCol >= 2) {
      sheet.getRange(2, 2, lastRow - 1, 1).setNumberFormat('yyyy/MM/dd HH:mm');
    }

    // 金額列フォーマット（4列目）
    if (lastCol >= 4) {
      sheet.getRange(2, 4, lastRow - 1, 1).setNumberFormat('¥#,##0');
    }
  }

  // 全セルに枠線
  if (lastRow >= 1) {
    var allRange = sheet.getRange(1, 1, Math.max(lastRow, 1), lastCol);
    allRange.setBorder(true, true, true, true, true, true, '#9e9e9e', SpreadsheetApp.BorderStyle.SOLID);
  }

  // 列幅自動調整
  sheet.autoResizeColumns(1, lastCol);
}

// ========== 入力シートのUI強化 ==========

function formatInputSheet() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('入力');
  if (!sheet) return;

  sheet.clearFormats();

  // シート全体の背景をリセット
  sheet.getRange(1, 1, 50, 10).setBackground('#fafafa');

  // ========== タイトルエリア（行1） ==========
  var titleRange = sheet.getRange(1, 1, 1, 6);
  titleRange.merge();
  titleRange.setValue('🏥 あまと整体院 — 売上入力フォーム');
  titleRange.setBackground('#1a237e');
  titleRange.setFontColor('#ffffff');
  titleRange.setFontWeight('bold');
  titleRange.setFontSize(16);
  titleRange.setHorizontalAlignment('center');
  titleRange.setVerticalAlignment('middle');
  sheet.setRowHeight(1, 50);

  // ========== サブタイトル（行2） ==========
  var subRange = sheet.getRange(2, 1, 1, 6);
  subRange.merge();
  subRange.setValue('施術ごとに種別・金額を入力して「売上を登録する」ボタンを押してください');
  subRange.setBackground('#3949ab');
  subRange.setFontColor('#e8eaf6');
  subRange.setFontSize(10);
  subRange.setHorizontalAlignment('center');
  subRange.setVerticalAlignment('middle');
  sheet.setRowHeight(2, 30);

  // ========== セクション：種別 ==========
  sheet.setRowHeight(3, 10); // 余白行

  var label1 = sheet.getRange(4, 1);
  label1.setValue('📋 種別');
  label1.setBackground('#e8eaf6');
  label1.setFontWeight('bold');
  label1.setFontSize(11);
  label1.setFontColor('#1a237e');
  label1.setBorder(true, true, true, true, false, false, '#9fa8da', SpreadsheetApp.BorderStyle.SOLID);

  var desc1 = sheet.getRange(4, 2);
  desc1.setValue('「新規」または「常連」を選択してください');
  desc1.setBackground('#e8eaf6');
  desc1.setFontColor('#616161');
  desc1.setFontSize(10);
  desc1.setBorder(true, true, true, true, false, false, '#9fa8da', SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange(4, 2, 1, 4).merge();
  sheet.setRowHeight(4, 30);

  var input1 = sheet.getRange(5, 2);
  input1.setBackground('#fff9c4');
  input1.setFontSize(12);
  input1.setBorder(true, true, true, true, false, false, '#f9a825', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sheet.getRange(5, 2, 1, 2).merge();
  sheet.setRowHeight(5, 35);

  // DataValidation for 種別
  var typeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['新規', '常連'], true)
    .setAllowInvalid(false)
    .build();
  input1.setDataValidation(typeRule);

  // ========== セクション：金額 ==========
  sheet.setRowHeight(6, 10); // 余白行

  var label2 = sheet.getRange(7, 1);
  label2.setValue('💰 金額');
  label2.setBackground('#e8eaf6');
  label2.setFontWeight('bold');
  label2.setFontSize(11);
  label2.setFontColor('#1a237e');
  label2.setBorder(true, true, true, true, false, false, '#9fa8da', SpreadsheetApp.BorderStyle.SOLID);

  var desc2 = sheet.getRange(7, 2);
  desc2.setValue('施術金額（税込）を数字で入力してください　例：3270');
  desc2.setBackground('#e8eaf6');
  desc2.setFontColor('#616161');
  desc2.setFontSize(10);
  desc2.setBorder(true, true, true, true, false, false, '#9fa8da', SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange(7, 2, 1, 4).merge();
  sheet.setRowHeight(7, 30);

  var input2 = sheet.getRange(8, 2);
  input2.setBackground('#fff9c4');
  input2.setFontSize(12);
  input2.setNumberFormat('¥#,##0');
  input2.setBorder(true, true, true, true, false, false, '#f9a825', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sheet.getRange(8, 2, 1, 2).merge();
  sheet.setRowHeight(8, 35);

  // ========== ボタン風セル（行10-11） ==========
  sheet.setRowHeight(9, 15); // 余白

  var btn1 = sheet.getRange(10, 2, 1, 2);
  btn1.merge();
  btn1.setValue('✅  新規 ¥3,270 を登録する');
  btn1.setBackground('#1565c0');
  btn1.setFontColor('#ffffff');
  btn1.setFontWeight('bold');
  btn1.setFontSize(12);
  btn1.setHorizontalAlignment('center');
  btn1.setVerticalAlignment('middle');
  btn1.setBorder(true, true, true, true, false, false, '#0d47a1', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sheet.setRowHeight(10, 40);

  sheet.setRowHeight(11, 8); // 余白

  var btn2 = sheet.getRange(12, 2, 1, 2);
  btn2.merge();
  btn2.setValue('✅  常連 ¥5,500 を登録する');
  btn2.setBackground('#1565c0');
  btn2.setFontColor('#ffffff');
  btn2.setFontWeight('bold');
  btn2.setFontSize(12);
  btn2.setHorizontalAlignment('center');
  btn2.setVerticalAlignment('middle');
  btn2.setBorder(true, true, true, true, false, false, '#0d47a1', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sheet.setRowHeight(12, 40);

  // ========== 注記 ==========
  sheet.setRowHeight(13, 10);
  var note = sheet.getRange(14, 1, 1, 6);
  note.merge();
  note.setValue('💡 ヒント：上のメニュー「売上管理」からも登録できます。売上データは「売上データ」シートに自動保存されます。');
  note.setBackground('#f3f4f6');
  note.setFontColor('#757575');
  note.setFontSize(9);
  note.setHorizontalAlignment('left');
  note.setVerticalAlignment('middle');
  note.setBorder(true, true, true, true, false, false, '#e0e0e0', SpreadsheetApp.BorderStyle.SOLID);
  sheet.setRowHeight(14, 32);

  // 列幅調整
  sheet.setColumnWidth(1, 100);
  sheet.setColumnWidth(2, 160);
  sheet.setColumnWidth(3, 160);
}

// ========== 集計シートのUI強化（グラフ追加版） ==========

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

    // シート全体をクリアしてから書き直す
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

    // 合計行
    rows.push(['【12ヶ月合計】', totalShinki, totalShinkiSales, totalJoren, totalJorenSales, totalCount, totalSales]);

    var range = summarySheet.getRange(1, 1, rows.length, headers.length);
    range.setValues(rows);

    // ヘッダー行の書式設定
    var headerRange = summarySheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#1a237e');
    headerRange.setFontColor('#ffffff');
    headerRange.setFontWeight('bold');
    headerRange.setFontSize(11);
    headerRange.setHorizontalAlignment('center');
    headerRange.setVerticalAlignment('middle');
    summarySheet.setRowHeight(1, 36);

    // 1行目を固定
    summarySheet.setFrozenRows(1);

    // データ行の交互背景色
    for (var r = 2; r <= rows.length - 1; r++) {
      var rowRange = summarySheet.getRange(r, 1, 1, headers.length);
      rowRange.setBackground(r % 2 === 0 ? '#e8eaf6' : '#ffffff');
      rowRange.setFontColor('#212121');
    }

    // 合計行の書式
    var totalRowRange = summarySheet.getRange(rows.length, 1, 1, headers.length);
    totalRowRange.setBackground('#1a237e');
    totalRowRange.setFontColor('#ffffff');
    totalRowRange.setFontWeight('bold');
    totalRowRange.setFontSize(11);

    // 数値列を通貨フォーマット
    summarySheet.getRange(2, 3, rows.length - 1, 1).setNumberFormat('¥#,##0');
    summarySheet.getRange(2, 5, rows.length - 1, 1).setNumberFormat('¥#,##0');
    summarySheet.getRange(2, 7, rows.length - 1, 1).setNumberFormat('¥#,##0');

    // 全セルに枠線
    summarySheet.getRange(1, 1, rows.length, headers.length)
      .setBorder(true, true, true, true, true, true, '#9e9e9e', SpreadsheetApp.BorderStyle.SOLID);

    summarySheet.autoResizeColumns(1, headers.length);

    // ========== 棒グラフ（月別合計売上）を埋め込む ==========
    // グラフ用データ範囲：年月（A列）と合計売上（G列）、データ行のみ（合計行除く）
    var dataRowCount = rows.length - 1; // 合計行を除く
    var chartRange = summarySheet.getRange(1, 1, dataRowCount, 1); // 年月
    var chartRange2 = summarySheet.getRange(1, 7, dataRowCount, 1); // 合計売上

    var chartBuilder = summarySheet.newChart()
      .setChartType(Charts.ChartType.COLUMN)
      .addRange(chartRange)
      .addRange(chartRange2)
      .setPosition(rows.length + 2, 1, 0, 0)
      .setOption('title', '月別合計売上（過去12ヶ月）')
      .setOption('titleTextStyle', { fontSize: 14, bold: true, color: '#1a237e' })
      .setOption('hAxis', {
        title: '年月',
        titleTextStyle: { color: '#424242', bold: true }
      })
      .setOption('vAxis', {
        title: '合計売上（円）',
        titleTextStyle: { color: '#424242', bold: true },
        format: '¥#,##0'
      })
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

// ========== 経理シート生成 ==========

function generateAccountingSheet(year, month) {
  try {
    year = parseInt(year);
    month = parseInt(month);
    if (!year || !month || month < 1 || month > 12) {
      return { success: false, message: '年月が不正です' };
    }

    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheetName = '経理_' + year + '_' + String(month).padStart(2, '0');

    // 既存シートがあれば削除して再作成
    var existing = ss.getSheetByName(sheetName);
    if (existing) {
      ss.deleteSheet(existing);
    }
    var accSheet = ss.insertSheet(sheetName);

    // ========== タイトル行 ==========
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

    // ヘッダー
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

    // 1・2行目を固定
    accSheet.setFrozenRows(2);

    // 売上データシートからデータ取得
    var dataSheet = ss.getSheetByName('売上データ') || ss.getSheets()[0];
    var lastRow = dataSheet.getLastRow();
    var allData = lastRow >= 2 ? dataSheet.getRange(2, 1, lastRow - 1, 5).getValues() : [];

    // 日別集計
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

    // 行データ作成
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

    // 月合計行
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

    // データ行の交互背景色
    for (var r = 0; r < daysInMonth; r++) {
      var rowRange = accSheet.getRange(3 + r, 1, 1, headers.length);
      rowRange.setBackground(r % 2 === 0 ? '#ffffff' : '#e8eaf6');
      rowRange.setFontColor('#212121');
    }

    // 月合計行の書式
    var totalRowRange = accSheet.getRange(3 + daysInMonth, 1, 1, headers.length);
    totalRowRange.setBackground('#1a237e');
    totalRowRange.setFontColor('#ffffff');
    totalRowRange.setFontWeight('bold');
    totalRowRange.setFontSize(11);

    // 金額列フォーマット
    accSheet.getRange(3, 3, dataRows.length, 1).setNumberFormat('¥#,##0');
    accSheet.getRange(3, 5, dataRows.length, 1).setNumberFormat('¥#,##0');
    accSheet.getRange(3, 6, dataRows.length, 1).setNumberFormat('¥#,##0');

    // 全セルに枠線
    accSheet.getRange(1, 1, 2 + dataRows.length, headers.length)
      .setBorder(true, true, true, true, true, true, '#9e9e9e', SpreadsheetApp.BorderStyle.SOLID);

    accSheet.autoResizeColumns(1, headers.length);

    return { success: true, message: sheetName + ' を生成しました（' + daysInMonth + '日分）' };
  } catch (err) {
    return { success: false, message: '経理シート生成エラー: ' + err.message };
  }
}

// ========== ダッシュボードシート新規作成 ==========

function createDashboard() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var dashboard = ss.getSheetByName('ダッシュボード');
    if (!dashboard) {
      dashboard = ss.insertSheet('ダッシュボード', 0); // 先頭に挿入
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

    // シート全体の背景
    dashboard.getRange(1, 1, 50, 10).setBackground('#f5f5f5');

    // ========== メインタイトル（行1-2） ==========
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

    // ========== 更新日時（行3） ==========
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

    // ========== 月次サマリーセクションタイトル（行5） ==========
    dashboard.setRowHeight(4, 16); // 余白
    var sectionTitle = dashboard.getRange(5, 1, 1, 8);
    sectionTitle.merge();
    sectionTitle.setValue('📊  ' + year + '年' + month + '月  月次サマリー');
    sectionTitle.setBackground('#e8eaf6');
    sectionTitle.setFontColor('#1a237e');
    sectionTitle.setFontWeight('bold');
    sectionTitle.setFontSize(14);
    sectionTitle.setHorizontalAlignment('left');
    sectionTitle.setVerticalAlignment('middle');
    sectionTitle.setBorder(false, false, false, false, false, false);
    sectionTitle.setBorder(true, true, true, true, false, false, '#9fa8da', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    dashboard.setRowHeight(5, 36);

    // ========== カード行1：新規 / 常連 ==========
    dashboard.setRowHeight(6, 12); // 余白

    // --- 新規カード（B7:D10） ---
    _drawCard(dashboard, 7, 2, '🆕  新規', stats.shinkiCount + ' 件', '¥' + stats.shinkiSales.toLocaleString(), '#e3f2fd', '#1565c0');

    // --- 常連カード（F7:H10） ---
    _drawCard(dashboard, 7, 6, '👥  常連', stats.jorenCount + ' 件', '¥' + stats.jorenSales.toLocaleString(), '#e8f5e9', '#2e7d32');

    // ========== カード行2：合計 / 今日 ==========
    dashboard.setRowHeight(12, 12); // 余白

    // --- 今月合計カード（B13:D16） ---
    _drawCard(dashboard, 13, 2, '💰  今月合計', stats.totalCount + ' 件', '¥' + stats.totalSales.toLocaleString(), '#fff8e1', '#f57f17');

    // --- 今日の売上カード（F13:H16） ---
    _drawCard(dashboard, 13, 6, '📅  本日', todayStats.totalCount + ' 件', '¥' + todayStats.totalSales.toLocaleString(), '#fce4ec', '#880e4f');

    dashboard.setRowHeight(17, 12); // 余白後

    // ========== 今日詳細セクション（行18-) ==========
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

    // 今日の新規詳細
    var todayData = [
      ['', '種別', '件数', '売上'],
      ['', '🆕 新規', todayStats.shinkiCount, todayStats.shinkiSales],
      ['', '👥 常連', todayStats.jorenCount, todayStats.jorenSales],
      ['', '💰 合計', todayStats.totalCount, todayStats.totalSales]
    ];

    var todayRange = dashboard.getRange(19, 1, 4, 8);
    // ヘッダー
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

    // 合計行を強調
    var todayTotalRow = dashboard.getRange(22, 2, 1, 3);
    todayTotalRow.setBackground('#fff9c4');
    todayTotalRow.setFontWeight('bold');
    todayTotalRow.setFontSize(13);

    // 金額列フォーマット
    dashboard.getRange(19, 4, 4, 1).setNumberFormat('¥#,##0');

    // 列幅設定
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

// カード描画ヘルパー関数
function _drawCard(sheet, startRow, startCol, title, count, amount, bgColor, accentColor) {
  var cardRows = 5;

  // カード全体
  var cardRange = sheet.getRange(startRow, startCol, cardRows, 3);
  cardRange.setBackground(bgColor);
  cardRange.setBorder(true, true, true, true, false, false, accentColor, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  // タイトル行
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

  // 件数行
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

  // 金額行
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

// ========== 全シートのフォーマットをまとめて実行 ==========

function formatAllSheets() {
  try {
    // 売上データシートのUI強化
    formatSalesDataSheet();

    // 入力シートのUI強化
    formatInputSheet();

    // 集計シートのUI強化（グラフ埋め込み）
    updateSummarySheet();

    // ダッシュボードシートを作成・更新
    createDashboard();

    SpreadsheetApp.getUi().alert('✅ 全シートのフォーマットが完了しました！\n\n・売上データシート：ヘッダー強化・交互背景・書式設定\n・入力シート：フォームUI強化\n・集計シート：グラフ追加・合計行\n・ダッシュボード：今月サマリー更新');
  } catch (err) {
    SpreadsheetApp.getUi().alert('❌ フォーマットエラー: ' + err.message);
  }
}

// ========== WebAPI ==========

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const manager = new SalesManager(SPREADSHEET_ID);
    let result;

    if (data.action === 'addSale') {
      result = manager.addSale(data.type, data.amount);
    } else if (data.action === 'backup') {
      result = runBackup();
    } else if (data.action === 'updateSummary') {
      result = updateSummarySheet();
    } else if (data.action === 'generateAccounting') {
      result = generateAccountingSheet(data.year, data.month);
    } else if (data.action === 'createDashboard') {
      result = createDashboard();
    } else if (data.action === 'formatAllSheets') {
      formatSalesDataSheet();
      formatInputSheet();
      result = { success: true, message: 'フォーマット完了' };
    } else {
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

function doGet(e) {
  try {
    const action = e.parameter.action || 'getTodayStats';
    const manager = new SalesManager(SPREADSHEET_ID);
    let result;

    if (action === 'getSheets') {
      const sheets = manager.ss.getSheets().map(function(s) { return s.getName(); });
      result = { sheets: sheets, dataSheetName: manager.dataSheet ? manager.dataSheet.getName() : null };
    } else if (action === 'getTodayStats') {
      result = manager.getTodayStats();
    } else if (action === 'getMonthStats') {
      const year = parseInt(e.parameter.year) || new Date().getFullYear();
      const month = parseInt(e.parameter.month) || new Date().getMonth() + 1;
      result = manager.getMonthStats(year, month);
    } else if (action === 'getRecentHistory') {
      const days = parseInt(e.parameter.days) || 30;
      result = { records: manager.getRecentHistory(days) };
    } else if (action === 'getCsv') {
      const year = e.parameter.year ? parseInt(e.parameter.year) : null;
      const month = e.parameter.month ? parseInt(e.parameter.month) : null;
      const rows = manager.getCsvData(year, month);
      const csv = rows.map(function(r) { return r.join(','); }).join('\n');
      return ContentService
        .createTextOutput(csv)
        .setMimeType(ContentService.MimeType.CSV);
    } else if (action === 'updateSummary') {
      result = updateSummarySheet();
    } else if (action === 'generateAccounting') {
      const year = parseInt(e.parameter.year) || new Date().getFullYear();
      const month = parseInt(e.parameter.month) || new Date().getMonth() + 1;
      result = generateAccountingSheet(year, month);
    } else if (action === 'createDashboard') {
      result = createDashboard();
    } else {
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
    .addItem('新規売上を記録', 'addNewCustomer')
    .addItem('常連売上を記録', 'addRegularCustomer')
    .addSeparator()
    .addItem('📊 ダッシュボードを更新', 'updateDashboard')
    .addItem('📈 集計シートを更新', 'updateSummaryFromMenu')
    .addItem('📑 経理シートを生成（今月）', 'generateThisMonthAccounting')
    .addSeparator()
    .addItem('✨ 全シートをフォーマット', 'formatAllSheets')
    .addSeparator()
    .addItem('今すぐバックアップ', 'manualBackupFromSheet')
    .addItem('自動バックアップ設定', 'setupDailyBackupTrigger')
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

function manualBackupFromSheet() {
  const result = runBackup();
  SpreadsheetApp.getUi().alert(result.message);
}

// メニューから呼ばれるラッパー関数
function updateDashboard() {
  const result = createDashboard();
  SpreadsheetApp.getUi().alert(result.message);
}

function updateSummaryFromMenu() {
  const result = updateSummarySheet();
  SpreadsheetApp.getUi().alert(result.message);
}

function generateThisMonthAccounting() {
  const now = new Date();
  const year = now.getFullYear();
  const month = now.getMonth() + 1;
  const result = generateAccountingSheet(year, month);
  SpreadsheetApp.getUi().alert(result.message);
}
