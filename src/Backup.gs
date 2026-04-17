// ========== Backup.gs ==========
// バックアップ機能

function getOrCreateBackupFolder() {
  var folders = DriveApp.getFoldersByName(BACKUP_FOLDER_NAME);
  if (folders.hasNext()) {
    return folders.next();
  }
  return DriveApp.createFolder(BACKUP_FOLDER_NAME);
}

function runBackup() {
  try {
    var folder = getOrCreateBackupFolder();
    var now = new Date();
    var label = now.getFullYear() + '-'
      + String(now.getMonth()+1).padStart(2,'0') + '-'
      + String(now.getDate()).padStart(2,'0')
      + '_' + String(now.getHours()).padStart(2,'0')
      + String(now.getMinutes()).padStart(2,'0');
    var fileName = 'あまと整体院_売上データ_' + label;
    var original = DriveApp.getFileById(SPREADSHEET_ID);
    var copy = original.makeCopy(fileName, folder);
    return { success: true, message: 'バックアップ完了: ' + fileName, fileId: copy.getId() };
  } catch (err) {
    return { success: false, message: 'バックアップエラー: ' + err.message };
  }
}

// 自動バックアップ（トリガーから呼ばれる）
function dailyAutoBackup() {
  runBackup();
}

// トリガーセットアップ（初回1回だけ手動で実行）
function setupDailyBackupTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'dailyAutoBackup') {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger('dailyAutoBackup')
    .timeBased()
    .everyDays(1)
    .atHour(2)
    .create();
  Logger.log('自動バックアップトリガーを設定しました（毎日深夜2時）');
}
