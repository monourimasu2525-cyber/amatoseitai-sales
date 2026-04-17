// ========== SalesManager.gs ==========
// 売上CRUD操作
// シート列構成: A=年 B=月 C=日 D=時間 E=種別 F=金額 G=入力方法

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
      const y = now.getFullYear();
      const m = now.getMonth() + 1;
      const d = now.getDate();
      const hhmm = Utilities.formatDate(now, 'Asia/Tokyo', 'HH:mm');
      this.dataSheet.appendRow([y, m, d, hhmm, type, amount, 'WebApp']);
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
      const hhmm = Utilities.formatDate(now, 'Asia/Tokyo', 'HH:mm');
      this.dataSheet.getRange(rowIndex, 4).setValue(hhmm);           // D=時間
      this.dataSheet.getRange(rowIndex, 5).setValue(type);           // E=種別
      this.dataSheet.getRange(rowIndex, 6).setValue(amount);         // F=金額
      this.dataSheet.getRange(rowIndex, 7).setValue('WebApp（修正）'); // G=入力方法
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
    const year  = today.getFullYear();
    const month = today.getMonth() + 1;
    const day   = today.getDate();
    const lastRow = this.dataSheet.getLastRow();
    if (lastRow < 2) {
      return { date: year+'年'+month+'月'+day+'日', shinkiCount:0, jorenCount:0, totalCount:0, shinkiSales:0, jorenSales:0, totalSales:0, otherCount:0, otherSales:0 };
    }
    // A=年 B=月 C=日 D=時間 E=種別 F=金額 → 6列取得
    const allData = this.dataSheet.getRange(2, 1, lastRow - 1, 6).getValues();
    let shinkiCount=0, shinkiSales=0, jorenCount=0, jorenSales=0, otherCount=0, otherSales=0;
    allData.forEach(function(row) {
      const y = Number(row[0]), m = Number(row[1]), d = Number(row[2]);
      if (!y) return;
      if (y === year && m === month && d === day) {
        const amount = Number(row[5]) || 0;
        const type   = row[4];
        if (type === '新規')      { shinkiCount++; shinkiSales += amount; }
        else if (type === '常連') { jorenCount++;  jorenSales  += amount; }
        else                      { otherCount++;  otherSales  += amount; }
      }
    });
    return {
      date: year+'年'+month+'月'+day+'日',
      shinkiCount, jorenCount, otherCount,
      totalCount:  shinkiCount + jorenCount + otherCount,
      shinkiSales, jorenSales, otherSales,
      totalSales:  shinkiSales + jorenSales + otherSales
    };
  }

  // 月間集計
  getMonthStats(year, month) {
    const lastRow = this.dataSheet.getLastRow();
    if (lastRow < 2) return { shinkiCount:0, jorenCount:0, totalCount:0, shinkiSales:0, jorenSales:0, totalSales:0, otherCount:0, otherSales:0 };
    const allData = this.dataSheet.getRange(2, 1, lastRow - 1, 6).getValues();
    let shinkiCount=0, shinkiSales=0, jorenCount=0, jorenSales=0, otherCount=0, otherSales=0;
    allData.forEach(function(row) {
      const y = Number(row[0]), m = Number(row[1]);
      if (!y) return;
      if (y === year && m === month) {
        const amount = Number(row[5]) || 0;
        const type   = row[4];
        if (type === '新規')      { shinkiCount++; shinkiSales += amount; }
        else if (type === '常連') { jorenCount++;  jorenSales  += amount; }
        else                      { otherCount++;  otherSales  += amount; }
      }
    });
    return {
      shinkiCount, jorenCount, otherCount,
      totalCount:  shinkiCount + jorenCount + otherCount,
      shinkiSales, jorenSales, otherSales,
      totalSales:  shinkiSales + jorenSales + otherSales,
      year, month
    };
  }

  // 直近履歴
  getRecentHistory(days) {
    const lastRow = this.dataSheet.getLastRow();
    if (lastRow < 2) return [];
    const allData = this.dataSheet.getRange(2, 1, lastRow - 1, 6).getValues();
    const today = new Date();
    const cutoffY = new Date(); cutoffY.setDate(today.getDate() - days);
    const result = [];
    allData.forEach(function(row, i) {
      const y = Number(row[0]), m = Number(row[1]), d = Number(row[2]);
      if (!y) return;
      const date = new Date(y, m-1, d);
      if (date >= cutoffY) {
        result.push({
          rowIndex: i + 2,
          date: y + '/' + m + '/' + d,
          time: String(row[3]) || '00:00',
          type: row[4],
          amount: Number(row[5]) || 0
        });
      }
    });
    return result.reverse();
  }

  // CSV出力
  getCsvData() {
    const lastRow = this.dataSheet.getLastRow();
    if (lastRow < 2) return '';
    const header = ['年','月','日','時間','種別','金額','入力方法'];
    const data = this.dataSheet.getRange(2, 1, lastRow - 1, 7).getValues();
    const rows = [header, ...data];
    return rows.map(r => r.join(',')).join('\n');
  }
}
