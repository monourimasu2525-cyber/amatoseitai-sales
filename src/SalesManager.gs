// ========== SalesManager.gs ==========
// 売上CRUD操作

class SalesManager {
  constructor(spreadsheetId) {
    this.ss = SpreadsheetApp.openById(spreadsheetId);
    this.dataSheet = this.ss.getSheetByName('売上データ') || this.ss.getSheets()[0];
  }

  // 売上追加
  addSale(type, amount) {
    if (!type) {
      return { success: false, message: '種別は必須です' };
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

  // 売上修正
  editSale(rowIndex, type, amount) {
    if (!rowIndex || !type) {
      return { success: false, message: '行番号と種別は必須です' };
    }
    if (typeof amount !== 'number' || amount < 0) {
      return { success: false, message: '金額は正の数値です' };
    }
    try {
      const now = new Date();
      // 既存行の種別・金額・更新日時を更新
      this.dataSheet.getRange(rowIndex, 2).setValue(now);
      this.dataSheet.getRange(rowIndex, 3).setValue(type);
      this.dataSheet.getRange(rowIndex, 4).setValue(amount);
      this.dataSheet.getRange(rowIndex, 5).setValue('WebAPI（修正）');
      return { success: true, message: '行' + rowIndex + ' を修正しました: ' + type + ' ¥' + amount };
    } catch (err) {
      return { success: false, message: 'エラー: ' + err.message };
    }
  }

  // 売上削除（行削除）
  deleteSale(rowIndex) {
    if (!rowIndex) {
      return { success: false, message: '行番号は必須です' };
    }
    try {
      this.dataSheet.deleteRow(rowIndex);
      return { success: true, message: '行' + rowIndex + ' を削除しました' };
    } catch (err) {
      return { success: false, message: 'エラー: ' + err.message };
    }
  }

  // 今日の集計
  getTodayStats() {
    const today = new Date();
    const year = today.getFullYear();
    const month = today.getMonth() + 1;
    const day = today.getDate();
    const lastRow = this.dataSheet.getLastRow();
    if (lastRow < 2) {
      return { date: year + '年' + month + '月' + day + '日', shinkiCount: 0, jorenCount: 0, totalCount: 0, shinkiSales: 0, jorenSales: 0, totalSales: 0, otherSales: 0, otherCount: 0 };
    }
    const allData = this.dataSheet.getRange(2, 1, lastRow - 1, 5).getValues();
    let shinkiCount = 0, shinkiSales = 0, jorenCount = 0, jorenSales = 0, otherCount = 0, otherSales = 0;
    allData.forEach(function(row) {
      if (!row[0]) return;
      const date = new Date(row[0]);
      if (date.getFullYear() === year && date.getMonth() + 1 === month && date.getDate() === day) {
        const amount = Number(row[3]) || 0;
        if (row[2] === '新規') { shinkiCount++; shinkiSales += amount; }
        else if (row[2] === '常連') { jorenCount++; jorenSales += amount; }
        else { otherCount++; otherSales += amount; }
      }
    });
    return {
      date: year + '年' + month + '月' + day + '日',
      shinkiCount: shinkiCount,
      jorenCount: jorenCount,
      totalCount: shinkiCount + jorenCount + otherCount,
      shinkiSales: shinkiSales,
      jorenSales: jorenSales,
      otherCount: otherCount,
      otherSales: otherSales,
      totalSales: shinkiSales + jorenSales + otherSales
    };
  }

  // 月間集計
  getMonthStats(year, month) {
    const lastRow = this.dataSheet.getLastRow();
    if (lastRow < 2) return { shinkiCount: 0, jorenCount: 0, totalCount: 0, shinkiSales: 0, jorenSales: 0, totalSales: 0, otherCount: 0, otherSales: 0 };
    const allData = this.dataSheet.getRange(2, 1, lastRow - 1, 5).getValues();
    let shinkiCount = 0, shinkiSales = 0, jorenCount = 0, jorenSales = 0, otherCount = 0, otherSales = 0;
    allData.forEach(function(row) {
      if (!row[0]) return;
      const date = new Date(row[0]);
      if (date.getFullYear() === year && date.getMonth() + 1 === month) {
        const amount = Number(row[3]) || 0;
        if (row[2] === '新規') { shinkiCount++; shinkiSales += amount; }
        else if (row[2] === '常連') { jorenCount++; jorenSales += amount; }
        else { otherCount++; otherSales += amount; }
      }
    });
    return {
      shinkiCount: shinkiCount,
      jorenCount: jorenCount,
      totalCount: shinkiCount + jorenCount + otherCount,
      shinkiSales: shinkiSales,
      jorenSales: jorenSales,
      otherCount: otherCount,
      otherSales: otherSales,
      totalSales: shinkiSales + jorenSales + otherSales
    };
  }

  // 直近履歴（rowIndexも含める）
  getRecentHistory(days) {
    const lastRow = this.dataSheet.getLastRow();
    if (lastRow < 2) return [];
    const allData = this.dataSheet.getRange(2, 1, lastRow - 1, 5).getValues();
    const cutoff = new Date();
    cutoff.setDate(cutoff.getDate() - days);
    const result = [];
    allData.forEach(function(row, i) {
      if (!row[0]) return;
      const date = new Date(row[0]);
      if (date >= cutoff) {
        result.push({
          rowIndex: i + 2,
          date: date.getFullYear() + '/' + (date.getMonth()+1) + '/' + date.getDate(),
          time: date.getHours() + ':' + String(date.getMinutes()).padStart(2,'0'),
          type: row[2],
          amount: Number(row[3]) || 0
        });
      }
    });
    return result.reverse();
  }

  // CSV用データ取得
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
