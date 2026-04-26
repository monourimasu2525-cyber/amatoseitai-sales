var SPREADSHEET_ID = '17bAyQngDEjoDgqSLLUU5p45HWXomF09bLf_h6FySsjs';
var CALENDAR_ID = 'primary';

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('売上管理')
    .addItem('シートを再セットアップ', 'setupSheets')
    .addSeparator()
    .addItem('集計シートを更新', 'refreshSummary')
    .addItem('売上入力を更新', 'refreshSalesInput')
    .addSeparator()
    .addItem('最新1件を削除', 'deleteLastSale')
    .addItem('売上データを修正', 'editSaleEntry')
    .addSeparator()
    .addItem('カレンダー連携を今すぐ実行', 'syncCalendarToSales')
    .addItem('カレンダー自動連携トリガー設定', 'setupCalendarTrigger')
    .addToUi();
}
// ============================================
// メインセットアップ
// ============================================
function setupSheets() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  var sheetNames = ['入力', '売上入力', '売上データ', '集計'];

  // 既存シートを確認、不足分を追加
  sheetNames.forEach(function(name) {
    if (!ss.getSheetByName(name)) {
      ss.insertSheet(name);
    }
  });

  // 順序を整える
  sheetNames.forEach(function(name, index) {
    var sheet = ss.getSheetByName(name);
    ss.setActiveSheet(sheet);
    ss.moveActiveSheet(index + 1);
  });

  setup入力Sheet(ss);
  setup売上入力Sheet(ss);
  setup売上データSheet(ss);
  setup集計Sheet(ss);

  SpreadsheetApp.getActive().toast('セットアップ完了！', '完了', 3);
}

// ============================================
// 入力シート（ボタン専用）
// ============================================
function setup入力Sheet(ss) {
  var sheet = ss.getSheetByName('入力');
  sheet.clear();
  sheet.setColumnWidth(1, 40);
  sheet.setColumnWidth(2, 220);
  sheet.setColumnWidth(3, 220);
  sheet.setColumnWidth(4, 40);
  sheet.setRowHeight(1, 30);
  sheet.setRowHeight(2, 30);
  sheet.setRowHeight(3, 120);
  sheet.setRowHeight(4, 30);

  sheet.getRange('B2').setValue('ボタンを押して売上を登録').setFontSize(14).setFontWeight('bold');
  sheet.getRange('B2:C2').setBackground('#F5F5F5');
}

// ============================================
// 売上入力シート（サマリー + 日別内訳）
// ============================================
function setup売上入力Sheet(ss) {
  if (!ss) ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('売上入力');
  sheet.clear();

  // ---- 列幅設定 ----
  sheet.setColumnWidth(1, 30);   // A: マージン
  sheet.setColumnWidth(2, 100);  // B: ラベル
  sheet.setColumnWidth(3, 110);  // C: 新規
  sheet.setColumnWidth(4, 110);  // D: 常連
  sheet.setColumnWidth(5, 110);  // E: 合計
  sheet.setColumnWidth(6, 30);   // F: スペース
  sheet.setColumnWidth(7, 60);   // G: No.
  sheet.setColumnWidth(8, 170);  // H: 日時
  sheet.setColumnWidth(9, 70);   // I: 種別
  sheet.setColumnWidth(10, 100); // J: 金額
  sheet.setColumnWidth(11, 110); // K: 入力方法

  // ---- 左側：本日・今月サマリー ----
  // タイトル
  sheet.getRange('B1:E1').merge().setValue('あまと整体院 売上管理').setFontSize(14).setFontWeight('bold')
    .setBackground('#1a237e').setFontColor('#ffffff').setHorizontalAlignment('center');

  // 今日の日付（数式）
  sheet.getRange('B2').setValue('集計日:').setFontWeight('bold').setHorizontalAlignment('right');
  sheet.getRange('C2:E2').merge().setFormula('=TEXT(TODAY(),"yyyy年m月d日（aaa）")').setFontWeight('bold').setFontSize(11);

  // ---- 本日の実績 ----
  sheet.getRange('B4:E4').merge().setValue('【本日の実績】').setFontWeight('bold').setFontSize(11)
    .setBackground('#E8EAF6').setFontColor('#1A237E');

  var hdrs = ['項目', '新規', '常連', '合計'];
  sheet.getRange('B5:E5').setValues([hdrs])
    .setBackground('#283593').setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');

  sheet.getRange('B6').setValue('件数').setBackground('#E8EAF6').setFontWeight('bold');
  sheet.getRange('B7').setValue('売上').setBackground('#E8EAF6').setFontWeight('bold');

  sheet.getRange('C6').setFormula('=COUNTIFS(売上データ!B:B,">="&TODAY(),売上データ!B:B,"<"&TODAY()+1,売上データ!C:C,"新規")');
  sheet.getRange('D6').setFormula('=COUNTIFS(売上データ!B:B,">="&TODAY(),売上データ!B:B,"<"&TODAY()+1,売上データ!C:C,"常連")');
  sheet.getRange('E6').setFormula('=C6+D6');
  sheet.getRange('C7').setFormula('=COUNTIFS(売上データ!B:B,">="&TODAY(),売上データ!B:B,"<"&TODAY()+1,売上データ!C:C,"新規")*3270');
  sheet.getRange('D7').setFormula('=COUNTIFS(売上データ!B:B,">="&TODAY(),売上データ!B:B,"<"&TODAY()+1,売上データ!C:C,"常連")*5500');
  sheet.getRange('E7').setFormula('=C7+D7');
  sheet.getRange('C6:E6').setHorizontalAlignment('center');
  sheet.getRange('C7:E7').setNumberFormat('¥#,##0').setHorizontalAlignment('center');
  sheet.getRange('E6:E7').setBackground('#FFF9C4');
  sheet.getRange('E7').setFontWeight('bold').setFontColor('#C62828');

  // ---- 今月の実績 ----
  sheet.getRange('B9:E9').merge().setValue('【今月の実績】').setFontWeight('bold').setFontSize(11)
    .setBackground('#E8F5E9').setFontColor('#1B5E20');

  sheet.getRange('B10:E10').setValues([hdrs])
    .setBackground('#2E7D32').setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');

  sheet.getRange('B11').setValue('件数').setBackground('#E8F5E9').setFontWeight('bold');
  sheet.getRange('B12').setValue('売上').setBackground('#E8F5E9').setFontWeight('bold');

  sheet.getRange('C11').setFormula('=COUNTIFS(売上データ!B:B,">="&DATE(YEAR(TODAY()),MONTH(TODAY()),1),売上データ!B:B,"<"&DATE(YEAR(TODAY()),MONTH(TODAY())+1,1),売上データ!C:C,"新規")');
  sheet.getRange('D11').setFormula('=COUNTIFS(売上データ!B:B,">="&DATE(YEAR(TODAY()),MONTH(TODAY()),1),売上データ!B:B,"<"&DATE(YEAR(TODAY()),MONTH(TODAY())+1,1),売上データ!C:C,"常連")');
  sheet.getRange('E11').setFormula('=C11+D11');
  sheet.getRange('C12').setFormula('=COUNTIFS(売上データ!B:B,">="&DATE(YEAR(TODAY()),MONTH(TODAY()),1),売上データ!B:B,"<"&DATE(YEAR(TODAY()),MONTH(TODAY())+1,1),売上データ!C:C,"新規")*3270');
  sheet.getRange('D12').setFormula('=COUNTIFS(売上データ!B:B,">="&DATE(YEAR(TODAY()),MONTH(TODAY()),1),売上データ!B:B,"<"&DATE(YEAR(TODAY()),MONTH(TODAY())+1,1),売上データ!C:C,"常連")*5500');
  sheet.getRange('E12').setFormula('=C12+D12');
  sheet.getRange('C11:E11').setHorizontalAlignment('center');
  sheet.getRange('C12:E12').setNumberFormat('¥#,##0').setHorizontalAlignment('center');
  sheet.getRange('E11:E12').setBackground('#DCEDC8');
  sheet.getRange('E12').setFontWeight('bold').setFontColor('#1B5E20');

  // ---- 今週の実績 ----
  sheet.getRange('B15:E15').merge().setValue('【今週の実績】').setFontWeight('bold').setFontSize(11)
    .setBackground('#E3F2FD').setFontColor('#0D47A1');

  sheet.getRange('B16:E16').setValues([hdrs])
    .setBackground('#1565C0').setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');

  sheet.getRange('B17').setValue('件数').setBackground('#E3F2FD').setFontWeight('bold');
  sheet.getRange('B18').setValue('売上').setBackground('#E3F2FD').setFontWeight('bold');

  // 今週の集計はrefreshで動的に計算するため初期値空欄
  sheet.getRange('C17:E18').setBackground('#F0F8FF');
  sheet.getRange('C18:E18').setNumberFormat('\u00A5#,##0').setHorizontalAlignment('center');
  sheet.getRange('E17:E18').setBackground('#B3D9FF');
  sheet.getRange('E18').setFontWeight('bold').setFontColor('#0D47A1');

  // ---- 前月比 / 前年比 ----
  sheet.getRange('B20:E20').merge().setValue('【前月比・前年比】').setFontWeight('bold').setFontSize(11)
    .setBackground('#FFF3E0').setFontColor('#E65100');

  sheet.getRange('B21:E21').setValues([['項目', '前月', '今月', '前月比']])
    .setBackground('#E65100').setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');

  sheet.getRange('B22').setValue('件数').setBackground('#FFF8F0').setFontWeight('bold');
  sheet.getRange('B23').setValue('売上').setBackground('#FFF8F0').setFontWeight('bold');

  sheet.getRange('C22:E23').setBackground('#FFF8F0');
  sheet.getRange('C23:D23').setNumberFormat('\u00A5#,##0').setHorizontalAlignment('center');
  sheet.getRange('E22:E23').setNumberFormat('0.0%').setHorizontalAlignment('center');
  sheet.getRange('E23').setFontWeight('bold').setFontColor('#E65100');

  sheet.getRange('B25:E25').merge().setValue('【前年比】').setFontWeight('bold').setFontSize(11)
    .setBackground('#F3E5F5').setFontColor('#4A148C');

  sheet.getRange('B26:E26').setValues([['項目', '前年同月', '今月', '前年比']])
    .setBackground('#4A148C').setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');

  sheet.getRange('B27').setValue('件数').setBackground('#F9F0FF').setFontWeight('bold');
  sheet.getRange('B28').setValue('売上').setBackground('#F9F0FF').setFontWeight('bold');

  sheet.getRange('C27:E28').setBackground('#F9F0FF');
  sheet.getRange('C28:D28').setNumberFormat('\u00A5#,##0').setHorizontalAlignment('center');
  sheet.getRange('E27:E28').setNumberFormat('0.0%').setHorizontalAlignment('center');
  sheet.getRange('E28').setFontWeight('bold').setFontColor('#4A148C');

  // ---- 注記 ----
  sheet.getRange('B14:E14').merge().setValue('※ 売上登録は「入力」シートのボタンから行ってください')
    .setFontColor('#9E9E9E').setFontSize(9).setHorizontalAlignment('center');

  // ---- 右側：日別内訳テーブル ----
  // タイトル
  sheet.getRange('G1:L1').merge().setValue('日別売上内訳').setFontSize(14).setFontWeight('bold')
    .setBackground('#4A148C').setFontColor('#ffffff').setHorizontalAlignment('center');

// 年・月・日 選択セル
  var today = new Date();
    sheet.getRange('G2').setValue('年:').setFontWeight('bold').setHorizontalAlignment('right');
      sheet.getRange('H2').setValue(today.getFullYear()).setBackground('#FFF9C4').setFontWeight('bold').setHorizontalAlignment('center').setNumberFormat('0');
        sheet.getRange('I2').setValue('月:').setFontWeight('bold').setHorizontalAlignment('right');
          sheet.getRange('J2').setValue(today.getMonth() + 1).setBackground('#FFF9C4').setFontWeight('bold').setHorizontalAlignment('center').setNumberFormat('0');
            sheet.getRange('K2').setValue('日:').setFontWeight('bold').setHorizontalAlignment('right');
              sheet.getRange('L2').setValue(today.getDate()).setBackground('#FFF9C4').setFontWeight('bold').setHorizontalAlignment('center').setNumberFormat('0');
                sheet.setColumnWidth(12, 60); // L列：日入力
                  
                    // バリデーション不要（数値直接入力）

  // 内訳テーブルヘッダー
  sheet.getRange('G3:K3').setValues([['No.', '日時', '種別', '金額', '入力方法']])
    .setBackground('#6A1B9A').setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');

  // 内訳データ行（最大30件分プレースホルダー）
  for (var i = 0; i < 30; i++) {
    var r = 4 + i;
    var bg = (i % 2 === 0) ? '#F3E5F5' : '#FAFAFA';
    sheet.getRange(r, 7, 1, 5).setBackground(bg);
    sheet.getRange(r, 10).setNumberFormat('¥#,##0');
  }

  // 合計行
  sheet.getRange('G34:K34').setBackground('#CE93D8');
  sheet.getRange('G34').setValue('合計').setFontWeight('bold').setHorizontalAlignment('center')
    .setBackground('#CE93D8');
  sheet.getRange('J34').setFontWeight('bold').setNumberFormat('¥#,##0').setBackground('#CE93D8');

  // 初回データ読み込み
  refreshSalesInputInternal(ss, sheet);
}

// ============================================
// 売上入力シートの日別内訳を更新
// ============================================
function refreshSalesInput() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('売上入力');
  refreshSalesInputInternal(ss, sheet);
  SpreadsheetApp.getActive().toast('内訳を更新しました', '完了', 2);
}

function refreshSalesInputInternal(ss, sheet) {
  if (!ss) ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  if (!sheet) sheet = ss.getSheetByName('売上入力');
  var dataSheet = ss.getSheetByName('売上データ');
  if (!sheet || !dataSheet) return;

  // 表示日を取得
var today2 = new Date();
  var y = parseInt(sheet.getRange('H2').getValue()) || today2.getFullYear();
    var moInput = parseInt(sheet.getRange('J2').getValue());
      var mo = (moInput >= 1 && moInput <= 12 ? moInput : today2.getMonth() + 1) - 1;
        var d = parseInt(sheet.getRange('L2').getValue()) || today2.getDate();

  // 売上データを全取得（A=日付, B=時間, C=種別, D=金額, E=入力方法, F=メモ）
  var lastRow = dataSheet.getLastRow();
  var allData = [];
  if (lastRow >= 2) {
    allData = dataSheet.getRange(2, 1, lastRow - 1, 6).getValues();
  }

  // 対象日のデータを抽出
  // A列が日付（yyyy/MM/dd形式）なので日付で比較
  var dayData = allData.filter(function(row) {
    if (!row[0]) return false;
    var dt = new Date(row[0]);
    return dt.getFullYear() === y && dt.getMonth() === mo && dt.getDate() === d;
  });

  // 時刻（B列）でソート
  dayData.sort(function(a, b) {
    var ta = String(a[1] || '');
    var tb = String(b[1] || '');
    return ta.localeCompare(tb);
  });

  // 既存データ行をクリア（行4〜33）
  sheet.getRange(4, 7, 30, 5).clearContent();
  for (var ci = 0; ci < 30; ci++) {
    var cr = 4 + ci;
    var cbg = (ci % 2 === 0) ? '#F3E5F5' : '#FAFAFA';
    sheet.getRange(cr, 7, 1, 5).setBackground(cbg);
    sheet.getRange(cr, 10).setNumberFormat('¥#,##0');
  }
  sheet.getRange(34, 7, 1, 5).clearContent();

  // データを書き込み
  // 列: G=No(連番), H=日時, I=種別, J=金額, K=入力方法
  var totalSales = 0;
  dayData.forEach(function(row, i) {
    if (i >= 30) return;
    var r = 4 + i;
    var bg = (i % 2 === 0) ? '#F3E5F5' : '#FAFAFA';

    // 日付文字列を作成（A列はDateTime型）
    var dtDate = new Date(row[0]);
    var dateStr = (dtDate.getMonth()+1) + '/' + dtDate.getDate();
    // B列もDateTimeオブジェクトなので時刻を取り出す
    var timeStr = '';
    if (row[1] && row[1] instanceof Date) {
      timeStr = ('0' + row[1].getHours()).slice(-2) + ':' + ('0' + row[1].getMinutes()).slice(-2);
    } else if (row[1] && String(row[1]).match(/\d+:\d+/)) {
      timeStr = String(row[1]).substring(0, 5);
    } else {
      timeStr = ('0' + dtDate.getHours()).slice(-2) + ':' + ('0' + dtDate.getMinutes()).slice(-2);
    }
    var displayStr = dateStr + ' ' + timeStr;

    sheet.getRange(r, 7).setValue(i + 1).setBackground(bg).setHorizontalAlignment('center').setNumberFormat('0');
    sheet.getRange(r, 8).setValue(displayStr).setBackground(bg).setHorizontalAlignment('center');
    sheet.getRange(r, 9).setValue(row[2]).setBackground(bg).setHorizontalAlignment('center'); // 種別
    sheet.getRange(r, 10).setValue(Number(row[3])).setBackground(bg).setNumberFormat('¥#,##0').setHorizontalAlignment('right'); // 金額
    sheet.getRange(r, 11).setValue(row[4] || 'ボタン').setBackground(bg).setHorizontalAlignment('center'); // 入力方法
    totalSales += Number(row[3]);
  });

  // データなしの場合
  if (dayData.length === 0) {
    sheet.getRange(4, 7, 1, 5).setBackground('#F3E5F5');
    sheet.getRange(4, 8).setValue('この日の売上データはありません').setFontColor('#9E9E9E')
      .setHorizontalAlignment('center').setFontStyle('italic');
  }

  // 合計行
  sheet.getRange(34, 7).setValue('合計').setFontWeight('bold').setHorizontalAlignment('center').setBackground('#CE93D8');
  sheet.getRange(34, 8).setValue(dayData.length + '件').setFontWeight('bold').setHorizontalAlignment('center').setBackground('#CE93D8');
  sheet.getRange(34, 9).setValue('').setBackground('#CE93D8');
  sheet.getRange(34, 10).setValue(totalSales).setFontWeight('bold').setNumberFormat('¥#,##0').setHorizontalAlignment('right').setBackground('#CE93D8');
  sheet.getRange(34, 11).setValue('').setBackground('#CE93D8');

  // G1:K1 タイトルを更新
  var weekdays = ['日', '月', '火', '水', '木', '金', '土'];
  var wday = weekdays[new Date(y, mo, d).getDay()];
  sheet.getRange('G1:L1').setValue(
    y + '年' + (mo+1) + '月' + d + '日（' + wday + '）の売上内訳'
  ).setFontSize(13).setFontWeight('bold')
    .setBackground('#4A148C').setFontColor('#ffffff').setHorizontalAlignment('center');

  // ---- 今週の実績を計算 ----
  var todayDate = new Date(y, mo, d);
  var dayOfWeek = todayDate.getDay(); // 0=日, 1=月...6=土
  // 週の開始(月曜)と終了(日曜)を計算
  var daysFromMonday = (dayOfWeek === 0) ? 6 : dayOfWeek - 1;
  var weekStart = new Date(todayDate);
  weekStart.setDate(weekStart.getDate() - daysFromMonday);
  weekStart.setHours(0,0,0,0);
  var weekEnd = new Date(weekStart);
  weekEnd.setDate(weekEnd.getDate() + 6);
  weekEnd.setHours(23,59,59,999);

  var weekNew = 0, weekReg = 0, weekNewSales = 0, weekRegSales = 0;
  allData.forEach(function(row) {
    if (!row[0]) return;
    var dt2 = new Date(row[0]);
    if (dt2 >= weekStart && dt2 <= weekEnd) {
      var amt = Number(row[3]) || 0;
      if (String(row[4]) === '新規') { weekNew++; weekNewSales += amt; }
      else if (String(row[4]) === '常連') { weekReg++; weekRegSales += amt; }
    }
  });
  sheet.getRange('C17').setValue(weekNew).setHorizontalAlignment('center');
  sheet.getRange('D17').setValue(weekReg).setHorizontalAlignment('center');
  sheet.getRange('E17').setValue(weekNew + weekReg).setHorizontalAlignment('center');
  sheet.getRange('C18').setValue(weekNewSales).setNumberFormat('\u00A5#,##0').setHorizontalAlignment('center');
  sheet.getRange('D18').setValue(weekRegSales).setNumberFormat('\u00A5#,##0').setHorizontalAlignment('center');
  sheet.getRange('E18').setValue(weekNewSales + weekRegSales).setNumberFormat('\u00A5#,##0').setHorizontalAlignment('center');

  // ---- 前月比を計算 ----
  var prevMo = mo - 1;
  var prevY = y;
  if (prevMo < 0) { prevMo = 11; prevY = y - 1; }
  var prevMonthStart = new Date(prevY, prevMo, 1);
  var prevMonthEnd = new Date(prevY, prevMo + 1, 0, 23, 59, 59);
  var thisMonthStart = new Date(y, mo, 1);
  var thisMonthEnd = new Date(y, mo + 1, 0, 23, 59, 59);

  var prevMonthCount = 0, prevMonthSales = 0;
  var thisMonthCount = 0, thisMonthSales = 0;
  allData.forEach(function(row) {
    if (!row[0]) return;
    var dt3 = new Date(row[0]);
    var amt = Number(row[3]) || 0;
    if (dt3 >= prevMonthStart && dt3 <= prevMonthEnd) { prevMonthCount++; prevMonthSales += amt; }
    if (dt3 >= thisMonthStart && dt3 <= thisMonthEnd) { thisMonthCount++; thisMonthSales += amt; }
  });

  sheet.getRange('C22').setValue(prevMonthCount).setHorizontalAlignment('center');
  sheet.getRange('D22').setValue(thisMonthCount).setHorizontalAlignment('center');
  sheet.getRange('E22').setValue(prevMonthCount > 0 ? thisMonthCount / prevMonthCount : 0)
    .setNumberFormat('0.0%').setHorizontalAlignment('center')
    .setFontColor(thisMonthCount >= prevMonthCount ? '#388E3C' : '#D32F2F')
    .setFontWeight('bold');
  sheet.getRange('C23').setValue(prevMonthSales).setNumberFormat('\u00A5#,##0').setHorizontalAlignment('center');
  sheet.getRange('D23').setValue(thisMonthSales).setNumberFormat('\u00A5#,##0').setHorizontalAlignment('center');
  sheet.getRange('E23').setValue(prevMonthSales > 0 ? thisMonthSales / prevMonthSales : 0)
    .setNumberFormat('0.0%').setHorizontalAlignment('center')
    .setFontColor(thisMonthSales >= prevMonthSales ? '#388E3C' : '#D32F2F')
    .setFontWeight('bold');

  // ---- 前年比を計算 ----
  var lastY = y - 1;
  var lastYearMonthStart = new Date(lastY, mo, 1);
  var lastYearMonthEnd = new Date(lastY, mo + 1, 0, 23, 59, 59);

  var lastYearCount = 0, lastYearSales = 0;
  allData.forEach(function(row) {
    if (!row[0]) return;
    var dt4 = new Date(row[0]);
    var amt = Number(row[3]) || 0;
    if (dt4 >= lastYearMonthStart && dt4 <= lastYearMonthEnd) { lastYearCount++; lastYearSales += amt; }
  });

  sheet.getRange('C27').setValue(lastYearCount).setHorizontalAlignment('center');
  sheet.getRange('D27').setValue(thisMonthCount).setHorizontalAlignment('center');
  sheet.getRange('E27').setValue(lastYearCount > 0 ? thisMonthCount / lastYearCount : 0)
    .setNumberFormat('0.0%').setHorizontalAlignment('center')
    .setFontColor(thisMonthCount >= lastYearCount ? '#388E3C' : '#D32F2F')
    .setFontWeight('bold');
  sheet.getRange('C28').setValue(lastYearSales).setNumberFormat('\u00A5#,##0').setHorizontalAlignment('center');
  sheet.getRange('D28').setValue(thisMonthSales).setNumberFormat('\u00A5#,##0').setHorizontalAlignment('center');
  sheet.getRange('E28').setValue(lastYearSales > 0 ? thisMonthSales / lastYearSales : 0)
    .setNumberFormat('0.0%').setHorizontalAlignment('center')
    .setFontColor(thisMonthSales >= lastYearSales ? '#388E3C' : '#D32F2F')
    .setFontWeight('bold');
}

// ============================================
// 売上データシート
// ============================================
function setup売上データSheet(ss) {
  var sheet = ss.getSheetByName('売上データ');
  var existingData = [];
  if (sheet.getLastRow() > 1) {
    existingData = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  }
  sheet.clear();

  var headers = ['年', '月', '日', '時間', '種別', '金額', '入力方法'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setBackground('#37474F').setFontColor('#ffffff').setFontWeight('bold');

    sheet.setColumnWidth(1, 60);
      sheet.setColumnWidth(2, 50);
        sheet.setColumnWidth(3, 50);
          sheet.setColumnWidth(4, 80);
            sheet.setColumnWidth(5, 80);
              sheet.setColumnWidth(6, 100);
                sheet.setColumnWidth(7, 120);

  if (existingData.length > 0) {
    sheet.getRange(2, 1, existingData.length, existingData[0].length).setValues(existingData);
  }

  sheet.getRange('A1:G1').setHorizontalAlignment('center');
}

// ============================================
// 集計シート（月別カレンダービュー）
// ============================================
function setup集計Sheet(ss) {
  if (!ss) ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('集計');
  sheet.clear();

  var now = new Date();

  // ---- 列幅設定 ----
  sheet.setColumnWidth(1, 30);   // A: マージン
  sheet.setColumnWidth(2, 60);   // B: 日
  sheet.setColumnWidth(3, 60);   // C: 曜日
  sheet.setColumnWidth(4, 70);   // D: 新規件数
  sheet.setColumnWidth(5, 70);   // E: 常連件数
  sheet.setColumnWidth(6, 90);   // F: 売上
  sheet.setColumnWidth(7, 30);   // G: スペース
  sheet.setColumnWidth(8, 80);   // H: 月
  sheet.setColumnWidth(9, 70);   // I: 件数
  sheet.setColumnWidth(10, 90);  // J: 売上

  // ---- タイトル ----
  sheet.getRange('B1:F1').merge().setValue('あまと整体院 売上集計').setFontSize(16).setFontWeight('bold')
    .setBackground('#1a237e').setFontColor('#ffffff').setHorizontalAlignment('center');

  // ---- 年・月セレクター ----
  sheet.getRange('B3').setValue('年:').setFontWeight('bold').setHorizontalAlignment('right');
  sheet.getRange('C3').setValue(now.getFullYear()).setBackground('#FFF9C4').setFontWeight('bold').setHorizontalAlignment('center');
  sheet.getRange('D3').setValue('月:').setFontWeight('bold').setHorizontalAlignment('right');
  sheet.getRange('E3').setValue(now.getMonth() + 1).setBackground('#FFF9C4').setFontWeight('bold').setHorizontalAlignment('center');

  // 年のバリデーション
  var yearRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['2024','2025','2026','2027','2028'], true).build();
  sheet.getRange('C3').setDataValidation(yearRule);

  // 月のバリデーション
  var monthRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['1','2','3','4','5','6','7','8','9','10','11','12'], true).build();
  sheet.getRange('E3').setDataValidation(monthRule);

  // ---- 左テーブル: 月別日別ビューのヘッダー ----
  sheet.getRange('B5:F5').setValues([['日', '曜日', '新規', '常連', '売上']])
    .setBackground('#1565C0').setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');

  // ---- 右テーブル: 年間月別サマリーのヘッダー ----
  sheet.getRange('H3:J3').merge().setValue('年間月別サマリー')
    .setBackground('#2E7D32').setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');
  sheet.getRange('H4:J4').setValues([['月', '件数', '売上']])
    .setBackground('#388E3C').setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');

  // 右テーブル: 12ヶ月分（H5:J16）
  for (var m = 1; m <= 12; m++) {
    var row = 4 + m;
    var bg = (m % 2 === 0) ? '#E8F5E9' : '#F1F8E9';
    sheet.getRange(row, 8).setValue(m + '月').setBackground(bg).setHorizontalAlignment('center');
    // 件数・売上は refreshSummary で入力
    sheet.getRange(row, 9).setBackground(bg).setHorizontalAlignment('center');
    sheet.getRange(row, 10).setBackground(bg).setNumberFormat('¥#,##0').setHorizontalAlignment('right');
  }
  // K列: 月選択チェックボックス
  sheet.setColumnWidth(11, 40);
  sheet.getRange('K4').setValue('選択').setBackground('#388E3C').setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');
  var currentMonth = new Date().getMonth() + 1;
  for (var mk = 1; mk <= 12; mk++) {
    var ck = sheet.getRange(4 + mk, 11);
    var bgk = (mk % 2 === 0) ? '#E8F5E9' : '#F1F8E9';
    ck.insertCheckboxes().setBackground(bgk);
    if (mk === currentMonth) ck.setValue(true);
  }

  // 年間合計行
  sheet.getRange('H17:J17').setBackground('#A5D6A7');
  sheet.getRange('H17').setValue('年間合計').setFontWeight('bold').setHorizontalAlignment('center');
  sheet.getRange('I17').setFontWeight('bold').setHorizontalAlignment('center');
  sheet.getRange('J17').setFontWeight('bold').setNumberFormat('¥#,##0').setHorizontalAlignment('right');

  // 左テーブル: 最大31日分の行を事前作成（行6〜36）
  for (var d = 1; d <= 31; d++) {
    var r = 5 + d;
    var rowBg = (d % 2 === 0) ? '#E3F2FD' : '#FAFAFA';
    sheet.getRange(r, 2).setBackground(rowBg).setHorizontalAlignment('center').setFontWeight('bold');   // 日
    sheet.getRange(r, 3).setBackground(rowBg).setHorizontalAlignment('center');   // 曜日
    sheet.getRange(r, 4).setBackground(rowBg).setHorizontalAlignment('center');   // 新規
    sheet.getRange(r, 5).setBackground(rowBg).setHorizontalAlignment('center');   // 常連
    sheet.getRange(r, 6).setBackground(rowBg).setNumberFormat('¥#,##0').setHorizontalAlignment('right'); // 売上
  }
  // 合計行（行38）
  sheet.getRange(38, 2, 1, 5).setBackground('#BBDEFB');
  sheet.getRange(38, 2).setValue('合計').setFontWeight('bold').setHorizontalAlignment('center');
  sheet.getRange(38, 6).setNumberFormat('¥#,##0').setFontWeight('bold');

  // ---- 下部: 月内明細テーブル ----
  sheet.getRange('B40:F40').merge().setValue('月内明細').setFontSize(12).setFontWeight('bold')
    .setBackground('#6A1B9A').setFontColor('#ffffff').setHorizontalAlignment('center');
  sheet.getRange('B41:F41').setValues([['日時', '種別', '金額', '入力方法', 'No.']])
    .setBackground('#7B1FA2').setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');

  // 初回データ投入
  refreshSummaryInternal(ss);
}

// ============================================
// 集計シートを更新（メニューから呼び出し可）
// ============================================
function refreshSummary() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  refreshSummaryInternal(ss);
  SpreadsheetApp.getActive().toast('集計を更新しました', '完了', 3);
}

function refreshSummaryInternal(ss, overrideMonth) {
  if (!ss) ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var summarySheet = ss.getSheetByName('集計');
  var dataSheet = ss.getSheetByName('売上データ');
  if (!summarySheet || !dataSheet) return;

  // 年・月を取得
  var year = parseInt(summarySheet.getRange('C3').getValue()) || new Date().getFullYear();
  var month = (typeof overrideMonth !== 'undefined' && overrideMonth > 0) 
    ? overrideMonth 
    : (parseInt(summarySheet.getRange('E3').getValue()) || (new Date().getMonth() + 1));
  // E3にも書き込む（月セルを同期）
  if (typeof overrideMonth !== 'undefined' && overrideMonth > 0) {
    summarySheet.getRange('E3').setValue(month);
  }

  // 売上データを全取得
  var lastRow = dataSheet.getLastRow();
  var allData = [];
  if (lastRow >= 2) {
    allData = dataSheet.getRange(2, 1, lastRow - 1, 5).getValues();
  }

  // 対象月の日数
  var daysInMonth = new Date(year, month, 0).getDate();
  var weekdays = ['日', '月', '火', '水', '木', '金', '土'];

  // ---- 左テーブル: 日別集計 ----
  // まず31行分をクリア
  for (var d = 1; d <= 31; d++) {
    var r = 5 + d;
    if (d <= daysInMonth) {
      var date = new Date(year, month - 1, d);
      var wday = weekdays[date.getDay()];
      var wdayColor = date.getDay() === 0 ? '#C62828' : (date.getDay() === 6 ? '#1565C0' : '#000000');
      var rowBg = (d % 2 === 0) ? '#E3F2FD' : '#FAFAFA';

      // 今日ハイライト
      var today = new Date();
      if (date.toDateString() === today.toDateString()) rowBg = '#FFF9C4';

      // 当日のデータを集計
      var shinki = 0, joren = 0, sales = 0;
      allData.forEach(function(row) {
        if (!row[1]) return;
        var d2 = new Date(row[1]);
        if (d2.getFullYear() === year && d2.getMonth() + 1 === month && d2.getDate() === d) {
          if (row[2] === '新規') { shinki++; sales += Number(row[3]); }
          else if (row[2] === '常連') { joren++; sales += Number(row[3]); }
        }
      });

      summarySheet.getRange(r, 2).setValue(d).setNumberFormat('0').setBackground(rowBg).setFontWeight('bold').setHorizontalAlignment('center').setFontColor(wdayColor);
      summarySheet.getRange(r, 3).setValue(wday).setBackground(rowBg).setFontColor(wdayColor).setHorizontalAlignment('center');
      summarySheet.getRange(r, 4).setValue(shinki || '').setBackground(rowBg).setHorizontalAlignment('center');
      summarySheet.getRange(r, 5).setValue(joren || '').setBackground(rowBg).setHorizontalAlignment('center');
      summarySheet.getRange(r, 6).setValue(sales || '').setBackground(rowBg).setNumberFormat('¥#,##0').setHorizontalAlignment('right');
    } else {
      // 月の日数を超える行はグレーアウト
      summarySheet.getRange(r, 2, 1, 5).clearContent().setBackground('#EEEEEE');
    }
  }

  // 合計行（行38）
  var totalShinki = 0, totalJoren = 0, totalSales = 0;
  allData.forEach(function(row) {
    if (!row[1]) return;
    var d2 = new Date(row[1]);
    if (d2.getFullYear() === year && d2.getMonth() + 1 === month) {
      if (row[2] === '新規') { totalShinki++; totalSales += Number(row[3]); }
      else if (row[2] === '常連') { totalJoren++; totalSales += Number(row[3]); }
    }
  });
  summarySheet.getRange(38, 2).setValue('合計').setFontWeight('bold').setBackground('#BBDEFB').setHorizontalAlignment('center');
  summarySheet.getRange(38, 3).setValue('').setBackground('#BBDEFB');
  summarySheet.getRange(38, 4).setValue(totalShinki || 0).setFontWeight('bold').setBackground('#BBDEFB').setHorizontalAlignment('center');
  summarySheet.getRange(38, 5).setValue(totalJoren || 0).setFontWeight('bold').setBackground('#BBDEFB').setHorizontalAlignment('center');
  summarySheet.getRange(38, 6).setValue(totalSales).setFontWeight('bold').setBackground('#BBDEFB').setNumberFormat('¥#,##0').setHorizontalAlignment('right');

  // ---- 右テーブル: 年間月別サマリー ----
  var totalYear = 0, totalYearCount = 0;
  for (var m = 1; m <= 12; m++) {
    var mShinki = 0, mJoren = 0, mSales = 0;
    allData.forEach(function(row) {
      if (!row[1]) return;
      var d3 = new Date(row[1]);
      if (d3.getFullYear() === year && d3.getMonth() + 1 === m) {
        if (row[2] === '新規') { mShinki++; mSales += Number(row[3]); }
        else if (row[2] === '常連') { mJoren++; mSales += Number(row[3]); }
      }
    });
    var mRow = 4 + m;
    var bg = (m % 2 === 0) ? '#E8F5E9' : '#F1F8E9';
    // 選択中の月はハイライト
    if (m === month) bg = '#C8E6C9';
    summarySheet.getRange(mRow, 9).setValue(mShinki + mJoren || 0).setBackground(bg).setHorizontalAlignment('center');
    summarySheet.getRange(mRow, 10).setValue(mSales).setBackground(bg).setNumberFormat('¥#,##0').setHorizontalAlignment('right');
    summarySheet.getRange(mRow, 8).setBackground(bg);
    totalYearCount += mShinki + mJoren;
    totalYear += mSales;
  }
  summarySheet.getRange(17, 9).setValue(totalYearCount).setFontWeight('bold').setBackground('#A5D6A7').setHorizontalAlignment('center');
  summarySheet.getRange(17, 10).setValue(totalYear).setFontWeight('bold').setBackground('#A5D6A7').setNumberFormat('¥#,##0').setHorizontalAlignment('right');

  // ---- 下部: 月内明細 ----
  // まず既存明細をクリア
  var detailStartRow = 42;
  var existingDetailRows = summarySheet.getLastRow() - detailStartRow + 1;
  if (existingDetailRows > 0) {
    summarySheet.getRange(detailStartRow, 2, Math.max(existingDetailRows, 50), 5).clearContent().setBackground('#FFFFFF');
  }

  // 対象月のデータを日時順で抽出
  var monthData = allData.filter(function(row) {
    if (!row[1]) return false;
    var d4 = new Date(row[1]);
    return d4.getFullYear() === year && d4.getMonth() + 1 === month;
  }).sort(function(a, b) { return new Date(a[1]) - new Date(b[1]); });

  monthData.forEach(function(row, i) {
    var r = detailStartRow + i;
    var bg = (i % 2 === 0) ? '#F3E5F5' : '#FAFAFA';
    var dt = new Date(row[1]);
    var dtStr = (dt.getMonth()+1) + '/' + dt.getDate() + ' ' + 
                ('0'+dt.getHours()).slice(-2) + ':' + ('0'+dt.getMinutes()).slice(-2);
    summarySheet.getRange(r, 2).setValue(dtStr).setBackground(bg).setHorizontalAlignment('center');
    summarySheet.getRange(r, 3).setValue(row[2]).setBackground(bg).setHorizontalAlignment('center');
    summarySheet.getRange(r, 4).setValue(row[3]).setBackground(bg).setNumberFormat('¥#,##0').setHorizontalAlignment('right');
    summarySheet.getRange(r, 5).setValue(row[4] || 'ボタン').setBackground(bg).setHorizontalAlignment('center');
    summarySheet.getRange(r, 6).setValue(Number(row[0])).setBackground(bg).setHorizontalAlignment('center').setNumberFormat('0');
  });
}

// ============================================
// onEdit: 年・月変更で自動更新 / 売上入力シートの日付変更
// ============================================
function onEdit(e) {
  var sheet = e.source.getActiveSheet();
  var cell = e.range;

  // ---- 集計シート ----
  if (sheet.getName() === '集計') {
    // C3（年）またはE3（月）が変更されたら更新
    if ((cell.getRow() === 3 && cell.getColumn() === 3) ||
        (cell.getRow() === 3 && cell.getColumn() === 5)) {
      refreshSummaryInternal(e.source);
      return;
    }
    // K列（列11）の行5〜16: チェックボックスで月を切り替え
    if (cell.getColumn() === 11 && cell.getRow() >= 5 && cell.getRow() <= 16) {
      var clickedMonth = cell.getRow() - 4;
      var ss = e.source;
      var summarySheet = ss.getSheetByName('集計');
      for (var r = 5; r <= 16; r++) {
        summarySheet.getRange(r, 11).setValue(r === cell.getRow() ? true : false);
      }
      refreshSummaryInternal(ss, clickedMonth);
    }
    return;
  }

  // ---- 売上入力シート ----
  if (sheet.getName() === '売上入力') {
    // H2（表示日）が変更されたら内訳を更新
    if (cell.getRow() === 2 && cell.getColumn() === 8) {
      refreshSalesInputInternal(e.source, sheet);
    }
  }
}

// ============================================
// 月クリックで左テーブルを切り替え
// ============================================
function onSelectionChange(e) {
  try {
    var sheet = e.source.getActiveSheet();
    if (sheet.getName() !== '集計') return;
    var cell = e.range;
    // H列（列8）の行5〜16が月リスト（1月〜12月）
    if (cell.getColumn() === 8 && cell.getRow() >= 5 && cell.getRow() <= 16) {
      var clickedMonth = cell.getRow() - 4; // 行5=1月, 行6=2月, ...行16=12月
      // 月を直接渡して更新（E3への書き込みも内部でする）
      refreshSummaryInternal(e.source, clickedMonth);
    }
  } catch(err) {
    Logger.log('onSelectionChange error: ' + err);
  }
}

// ============================================
// ボタン関数
// ============================================
function addNewCustomer() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('売上データ');
  var now = new Date();
  var lastRow = sheet.getLastRow();
  
    var newRow = [now, now, '新規', 3270, 'ボタン'];
  sheet.appendRow(newRow);
    sheet.getRange(lastRow + 1, 1).setNumberFormat('yyyy/MM/dd HH:mm');
      sheet.getRange(lastRow + 1, 2).setNumberFormat('yyyy/MM/dd HH:mm');
        sheet.getRange(lastRow + 1, 4).setNumberFormat('¥#,##0');

  var month = now.getMonth() + 1;
  var day = now.getDate();
  var hour = ('0' + now.getHours()).slice(-2);
  var min = ('0' + now.getMinutes()).slice(-2);
    SpreadsheetApp.getUi().alert('新規　' + month + '月' + day + '日 ' + hour + ':' + min + '　¥3,270 を登録しました！');
  refreshSummaryInternal(ss);
  var salesInputSheet2 = ss.getSheetByName('売上入力');
  if (salesInputSheet2) refreshSalesInputInternal(ss, salesInputSheet2);
}

function addRegularCustomer() {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
      var sheet = ss.getSheetByName('売上データ');
        var now = new Date();
          var lastRow = sheet.getLastRow();

            var newRow = [now, now, '常連', 5500, 'ボタン'];
              sheet.appendRow(newRow);
                sheet.getRange(lastRow + 1, 1).setNumberFormat('yyyy/MM/dd HH:mm');
                  sheet.getRange(lastRow + 1, 2).setNumberFormat('yyyy/MM/dd HH:mm');
                    sheet.getRange(lastRow + 1, 4).setNumberFormat('¥#,##0');

                      var month = now.getMonth() + 1;
                        var day = now.getDate();
                          var hour = ('0' + now.getHours()).slice(-2);
                            var min = ('0' + now.getMinutes()).slice(-2);
                              SpreadsheetApp.getUi().alert('常連　' + month + '月' + day + '日 ' + hour + ':' + min + '　¥5,500 を登録しました！');
                                refreshSummaryInternal(ss);
                                  var salesInputSheet2 = ss.getSheetByName('売上入力');
                                    if (salesInputSheet2) refreshSalesInputInternal(ss, salesInputSheet2);
}


// ============================================
// 売上データ削除・修正機能
// ============================================

// 最新1件を削除
function deleteLastSale() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var dataSheet = ss.getSheetByName('売上データ');
  var lastRow = dataSheet.getLastRow();

  if (lastRow <= 1) {
    SpreadsheetApp.getUi().alert('削除できるデータがありません。');
    return;
  }

  // 最後の行のデータを取得して確認ダイアログ表示
  var lastData = dataSheet.getRange(lastRow, 1, 1, 7).getValues()[0];
  var dateVal = lastData[0] ? new Date(lastData[0]) : null;
  var dateStr = dateVal ? (dateVal.getMonth()+1) + '月' + dateVal.getDate() + '日 ' +
    ('0' + dateVal.getHours()).slice(-2) + ':' + ('0' + dateVal.getMinutes()).slice(-2) : '日時不明';
  var typeStr = lastData[2] || '';
  var amtStr = lastData[3] ? '¥' + Number(lastData[3]).toLocaleString() : '¥0';

  var ui = SpreadsheetApp.getUi();
  var response = ui.alert(
    '最新データを削除',
    '以下のデータを削除しますか？\n\n' +
    '日時: ' + dateStr + '\n' +
    '種別: ' + typeStr + '\n' +
    '金額: ' + amtStr + '\n\n' +
    '※この操作は元に戻せません',
    ui.ButtonSet.YES_NO
  );

  if (response === ui.Button.YES) {
    dataSheet.deleteRow(lastRow);
    refreshSummaryInternal(ss);
    var salesInputSheet = ss.getSheetByName('売上入力');
    if (salesInputSheet) refreshSalesInputInternal(ss, salesInputSheet);
    SpreadsheetApp.getActive().toast('削除しました: ' + typeStr + ' ' + amtStr, '完了', 3);
  }
}

// 売上データを修正（ダイアログで選択・編集）
function editSaleEntry() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var dataSheet = ss.getSheetByName('売上データ');
  var lastRow = dataSheet.getLastRow();

  if (lastRow <= 1) {
    SpreadsheetApp.getUi().alert('修正できるデータがありません。');
    return;
  }

  // 最新10件を取得してリスト表示
  var startRow = Math.max(2, lastRow - 9);
  var rowCount = lastRow - startRow + 1;
  var recentData = dataSheet.getRange(startRow, 1, rowCount, 5).getValues();

  // 選択肢リストを作成（新しい順）
  var listItems = [];
  for (var i = recentData.length - 1; i >= 0; i--) {
    var row = recentData[i];
    var dt = row[0] ? new Date(row[0]) : null;
    var dStr = dt ? (dt.getMonth()+1) + '/' + dt.getDate() + ' ' +
      ('0' + dt.getHours()).slice(-2) + ':' + ('0' + dt.getMinutes()).slice(-2) : '不明';
    listItems.push((recentData.length - i) + ': ' + dStr + ' ' + (row[2]||'') + ' ¥' + Number(row[3]||0).toLocaleString());
  }

  var ui = SpreadsheetApp.getUi();
  var listStr = listItems.join('\n');

  var numResponse = ui.prompt(
    '売上データ修正',
    '修正する行の番号を入力してください（新しい順）：\n\n' + listStr,
    ui.ButtonSet.OK_CANCEL
  );

  if (numResponse.getSelectedButton() !== ui.Button.OK) return;
  var selectedNum = parseInt(numResponse.getResponseText());
  if (isNaN(selectedNum) || selectedNum < 1 || selectedNum > recentData.length) {
    ui.alert('有効な番号を入力してください。');
    return;
  }

  // 選択された行のデータ（新しい順なので逆算）
  var selectedIdx = recentData.length - selectedNum;
  var targetRow = startRow + selectedIdx;
  var targetData = recentData[selectedIdx];

  // 種別の修正
  var typeResponse = ui.prompt(
    '種別を修正',
    '現在の種別: ' + (targetData[2]||'') + '\n\n新しい種別を入力 (新規 / 常連):',
    ui.ButtonSet.OK_CANCEL
  );
  if (typeResponse.getSelectedButton() !== ui.Button.OK) return;
  var newType = typeResponse.getResponseText().trim();
  if (newType !== '新規' && newType !== '常連') {
    ui.alert('"新規" または "常連" と入力してください。');
    return;
  }

  // 金額の修正
  var amtResponse = ui.prompt(
    '金額を修正',
    '現在の金額: ¥' + Number(targetData[3]||0).toLocaleString() + '\n\n新しい金額を入力（数字のみ）:',
    ui.ButtonSet.OK_CANCEL
  );
  if (amtResponse.getSelectedButton() !== ui.Button.OK) return;
  var newAmt = parseInt(amtResponse.getResponseText().replace(/[^0-9]/g, ''));
  if (isNaN(newAmt) || newAmt < 0) {
    ui.alert('正しい金額を入力してください。');
    return;
  }

  // データ更新
  dataSheet.getRange(targetRow, 3).setValue(newType);
  dataSheet.getRange(targetRow, 4).setValue(newAmt);

  refreshSummaryInternal(ss);
  var salesInputSheet = ss.getSheetByName('売上入力');
  if (salesInputSheet) refreshSalesInputInternal(ss, salesInputSheet);

  SpreadsheetApp.getActive().toast('修正しました: ' + newType + ' ¥' + newAmt.toLocaleString(), '完了', 3);
}


// ============================================
// カレンダー連携
// ============================================
function syncCalendarToSales() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var dataSheet = ss.getSheetByName('売上データ');

  var calendar = CalendarApp.getCalendarById(CALENDAR_ID);
  if (!calendar) {
    SpreadsheetApp.getUi().alert('カレンダーが見つかりません。CALENDAR_IDを確認してください。');
    return;
  }

  // 過去30日分を対象
  var startDate = new Date();
  startDate.setDate(startDate.getDate() - 30);
  var endDate = new Date();

  var events = calendar.getEvents(startDate, endDate);

  // 既存データのタイムスタンプセットを作成（重複防止）
  var existingTimes = new Set();
  var lastRow = dataSheet.getLastRow();
  if (lastRow >= 2) {
    var existingData = dataSheet.getRange(2, 2, lastRow - 1, 1).getValues();
    existingData.forEach(function(row) {
      if (row[0]) existingTimes.add(new Date(row[0]).getTime());
    });
  }

  var addedCount = 0;

  events.forEach(function(event) {
    var title = event.getTitle();

    // キャンセル/先延ばし/リスケはスキップ
    if (title.match(/キャンセル|先延ばし|リスケ/)) return;

    var startTime = event.getStartTime();

    // 重複チェック
    if (existingTimes.has(startTime.getTime())) return;

    // 種別・金額の判定
    var kind, amount;
    if (title.match(/新規/)) {
      kind = '新規'; amount = 3270;
    } else {
      kind = '常連'; amount = 5500;
    }

    var currentLastRow = dataSheet.getLastRow();
    var no = currentLastRow;
    dataSheet.appendRow([no, startTime, kind, amount, 'カレンダー']);
    dataSheet.getRange(currentLastRow + 1, 2).setNumberFormat('yyyy/MM/dd HH:mm');
    dataSheet.getRange(currentLastRow + 1, 4).setNumberFormat('¥#,##0');

    existingTimes.add(startTime.getTime());
    addedCount++;
  });

  refreshSummaryInternal(ss);
  SpreadsheetApp.getUi().alert(addedCount + '件のカレンダー予定を売上に追加しました。');
}

function setupCalendarTrigger() {
  // 既存トリガーを削除
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'syncCalendarToSales') {
      ScriptApp.deleteTrigger(t);
    }
  });
  // 1時間ごとに実行
  ScriptApp.newTrigger('syncCalendarToSales').timeBased().everyHours(1).create();
  SpreadsheetApp.getUi().alert('カレンダー自動連携トリガーを設定しました（1時間ごと）。');
}

// ============================================
// チェックボックスを集計シートに追加
// ============================================
function addCheckboxesToSummary() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('集計');
  if (!sheet) { SpreadsheetApp.getUi().alert('集計シートが見つかりません'); return; }
  
  sheet.setColumnWidth(11, 40);
  sheet.getRange('K4').setValue('選択').setBackground('#388E3C').setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');
  
  var currentMonth = parseInt(sheet.getRange('E3').getValue()) || (new Date().getMonth() + 1);
  
  for (var m = 1; m <= 12; m++) {
    var row = 4 + m;
    var bg = (m % 2 === 0) ? '#E8F5E9' : '#F1F8E9';
    var ck = sheet.getRange(row, 11);
    ck.insertCheckboxes();
    ck.setBackground(bg);
    ck.setValue(m === currentMonth);
  }
  SpreadsheetApp.getActive().toast('チェックボックスを追加しました！', '完了', 3);
}

// ============================================
// onSelectionChangeトリガーをインストール
// ============================================
function setupSelectionTrigger() {
  // 既存のonSelectionChangeトリガーを削除
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'onSelectionChange') {
      ScriptApp.deleteTrigger(t);
    }
  });
  // インストール型トリガーとして登録
  ScriptApp.newTrigger('onSelectionChange')
    .forSpreadsheet(SPREADSHEET_ID)
    .onSelectionChange()
    .create();
  SpreadsheetApp.getUi().alert('選択変更トリガーを設定しました！月リストをクリックすると切り替わります。');
}

// ============================================
// ボタン位置調整（手動実行）
// ============================================
function positionButtons() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('入力');
  var drawings = sheet.getDrawings();
  SpreadsheetApp.getUi().alert('ボタン数: ' + drawings.length);
  if (drawings.length >= 2) {
    drawings[0].setPosition(3, 2, 10, 10);
    drawings[1].setPosition(3, 3, 10, 10);
  }
  SpreadsheetApp.getUi().alert('ボタンの位置を調整しました！');
}