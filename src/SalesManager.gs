// ========== SalesManager.gs ==========

class SalesManager {
  constructor(spreadsheetId) {
    this.ss = SpreadsheetApp.openById(spreadsheetId);
    this.dataSheet = this.ss.getSheetByName('売上データ') || this.ss.getSheets()[0];
  }

  addSale(type, amount) {
    if (!type) return { success: false, message: '種別は必須です' };
    if (typeof amount !== 'number' || amount < 0) return { success: false, message: '金額は正の数値です' };
    try {
      var now = new Date();
      this.dataSheet.appendRow([now, now, type, amount, 'WebAPI']);
      return { success: true, message: type + ' ¥' + amount + ' を登録しました', timestamp: now.toISOString(), type: type, amount: amount };
    } catch (err) { return { success: false, message: 'エラー: ' + err.message }; }
  }

  editSale(rowIndex, type, amount) {
    if (!rowIndex || !type) return { success: false, message: '行番号と種別は必須です' };
    if (typeof amount !== 'number' || amount < 0) return { success: false, message: '金額は正の数値です' };
    try {
      var now = new Date();
      this.dataSheet.getRange(rowIndex, 2).setValue(now);
      this.dataSheet.getRange(rowIndex, 3).setValue(type);
      this.dataSheet.getRange(rowIndex, 4).setValue(amount);
      this.dataSheet.getRange(rowIndex, 5).setValue('WebAPI（修正）');
      return { success: true, message: '行' + rowIndex + ' を修正しました' };
    } catch (err) { return { success: false, message: 'エラー: ' + err.message }; }
  }

  deleteSale(rowIndex) {
    if (!rowIndex) return { success: false, message: '行番号は必須です' };
    try {
      this.dataSheet.deleteRow(rowIndex);
      return { success: true, message: '行' + rowIndex + ' を削除しました' };
    } catch (err) { return { success: false, message: 'エラー: ' + err.message }; }
  }

  getTodayStats() {
    var today = new Date();
    var year = today.getFullYear(), month = today.getMonth() + 1, day = today.getDate();
    var lastRow = this.dataSheet.getLastRow();
    if (lastRow < 2) return { date: year+'年'+month+'月'+day+'日', shinkiCount:0, jorenCount:0, totalCount:0, shinkiSales:0, jorenSales:0, otherCount:0, otherSales:0, totalSales:0 };
    var allData = this.dataSheet.getRange(2, 1, lastRow - 1, 5).getValues();
    var shinkiCount=0, shinkiSales=0, jorenCount=0, jorenSales=0, otherCount=0, otherSales=0;
    allData.forEach(function(row) {
      if (!row[0]) return;
      var d = new Date(row[0]);
      if (d.getFullYear()===year && d.getMonth()+1===month && d.getDate()===day) {
        var amt = Number(row[3])||0;
        if (row[2]==='新規') { shinkiCount++; shinkiSales+=amt; }
        else if (row[2]==='常連') { jorenCount++; jorenSales+=amt; }
        else { otherCount++; otherSales+=amt; }
      }
    });
    return { date:year+'年'+month+'月'+day+'日', shinkiCount, jorenCount, totalCount:shinkiCount+jorenCount+otherCount, shinkiSales, jorenSales, otherCount, otherSales, totalSales:shinkiSales+jorenSales+otherSales };
  }

  getMonthStats(year, month) {
    var lastRow = this.dataSheet.getLastRow();
    if (lastRow < 2) return { shinkiCount:0, jorenCount:0, totalCount:0, shinkiSales:0, jorenSales:0, otherCount:0, otherSales:0, totalSales:0 };
    var allData = this.dataSheet.getRange(2, 1, lastRow - 1, 5).getValues();
    var shinkiCount=0, shinkiSales=0, jorenCount=0, jorenSales=0, otherCount=0, otherSales=0;
    allData.forEach(function(row) {
      if (!row[0]) return;
      var d = new Date(row[0]);
      if (d.getFullYear()===year && d.getMonth()+1===month) {
        var amt = Number(row[3])||0;
        if (row[2]==='新規') { shinkiCount++; shinkiSales+=amt; }
        else if (row[2]==='常連') { jorenCount++; jorenSales+=amt; }
        else { otherCount++; otherSales+=amt; }
      }
    });
    return { shinkiCount, jorenCount, totalCount:shinkiCount+jorenCount+otherCount, shinkiSales, jorenSales, otherCount, otherSales, totalSales:shinkiSales+jorenSales+otherSales };
  }

  getRecentHistory(days) {
    var lastRow = this.dataSheet.getLastRow();
    if (lastRow < 2) return [];
    var allData = this.dataSheet.getRange(2, 1, lastRow - 1, 5).getValues();
    var cutoff = new Date(); cutoff.setDate(cutoff.getDate() - days);
    var result = [];
    allData.forEach(function(row, i) {
      if (!row[0]) return;
      var d = new Date(row[0]);
      if (d >= cutoff) {
        result.push({
          rowIndex: i + 2,
          date: d.getFullYear()+'/'+(d.getMonth()+1)+'/'+d.getDate(),
          time: String(d.getHours()).padStart(2,'0')+':'+String(d.getMinutes()).padStart(2,'0'),
          type: row[2],
          amount: Number(row[3])||0
        });
      }
    });
    return result.reverse();
  }

  getCsvData(year, month) {
    var lastRow = this.dataSheet.getLastRow();
    if (lastRow < 2) return [];
    var allData = this.dataSheet.getRange(2, 1, lastRow - 1, 5).getValues();
    var result = [['日付','時刻','種別','金額','入力方法']];
    allData.forEach(function(row) {
      if (!row[0]) return;
      var d = new Date(row[0]);
      if (!year || (d.getFullYear()===year && d.getMonth()+1===month)) {
        result.push([
          d.getFullYear()+'/'+(d.getMonth()+1)+'/'+d.getDate(),
          String(d.getHours()).padStart(2,'0')+':'+String(d.getMinutes()).padStart(2,'0'),
          row[2], Number(row[3])||0, row[4]||'WebAPI'
        ]);
      }
    });
    return result;
  }
}
