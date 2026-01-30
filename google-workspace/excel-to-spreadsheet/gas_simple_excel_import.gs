/**
 * ===== 設定エリア（ここを編集してください） =====
 */
const EXCEL_URL = 'https://docs.google.com/spreadsheets/d/SPREADSHEET_ID_PLACEHOLDER/edit';  // GoogleスプレッドシートのURL

// 複数ファイルの場合は配列で指定
// const EXCEL_URLS = [
//   'https://drive.google.com/file/d/FILE_ID_1/view',
//   'https://drive.google.com/file/d/FILE_ID_2/view',
// ];

/**
 * メイン実行関数（手動実行 or トリガー実行）
 */
function importExcel() {
  try {
    console.log('Excel取り込み開始...');
    
    const fileId = extractFileIdFromUrl(EXCEL_URL);
    const result = processExcelFile(fileId);
    
    console.log(`取り込み完了: ${result.sheetCount}シート`);
    return result;
    
  } catch (error) {
    console.error('エラー:', error);
    throw error;
  }
}

/**
 * URLからファイルIDを抽出
 */
function extractFileIdFromUrl(url) {
  const patterns = [
    /\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/,  // Googleスプレッドシート用
    /\/d\/([a-zA-Z0-9-_]+)/,
    /id=([a-zA-Z0-9-_]+)/,
    /\/file\/d\/([a-zA-Z0-9-_]+)/,
  ];
  
  for (const pattern of patterns) {
    const match = url.match(pattern);
    if (match) return match[1];
  }
  
  // URLがそのままファイルIDの場合
  if (/^[a-zA-Z0-9-_]+$/.test(url)) {
    return url;
  }
  
  throw new Error('有効なURLではありません');
}

/**
 * スプレッドシート/Excelファイルを処理（値のみ取得）
 */
function processExcelFile(fileId) {
  try {
    // ファイルの存在とアクセス権限を確認
    const file = DriveApp.getFileById(fileId);
    console.log('ファイル名:', file.getName());
    console.log('MIME Type:', file.getMimeType());
    
    const mimeType = file.getMimeType();
    
    // Googleスプレッドシートの場合は直接処理
    if (mimeType === 'application/vnd.google-apps.spreadsheet') {
      console.log('Googleスプレッドシートを処理中...');
      return processGoogleSpreadsheet(fileId);
    }
    
    // Excelファイルかチェック
    const isExcel = mimeType.includes('excel') || 
                    mimeType.includes('openxmlformats');
    
    if (!isExcel) {
      throw new Error(`対応していないファイル形式です。MIME Type: ${mimeType}`);
    }
    
  } catch (error) {
    throw new Error(`ファイルアクセスエラー: ${error.toString()}`);
  }
  
  // DriveAppを使ってExcelファイルを取得
  console.log('Excelファイルを変換中...');
  const excelFile = DriveApp.getFileById(fileId);
  const blob = excelFile.getBlob();
  
  // BlobからGoogleスプレッドシートを作成
  const resource = {
    title: 'TEMP_' + new Date().getTime(),
    mimeType: 'application/vnd.google-apps.spreadsheet'
  };
  
  const tempFile = Drive.Files.insert(resource, blob, {
    convert: true  // Excelをスプレッドシートに変換
  });
  
  const destSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const importedSheets = [];
  
  try {
    const tempSpreadsheet = SpreadsheetApp.openById(tempFile.id);
    const sourceSheets = tempSpreadsheet.getSheets();
    
    sourceSheets.forEach(sourceSheet => {
      const sheetName = sourceSheet.getName();
      const values = sourceSheet.getDataRange().getValues();
      
      // 空シートはスキップ
      if (values.length === 0) return;
      
      // 取り込み先シート（同名があれば上書き、なければ新規作成）
      let destSheet = destSpreadsheet.getSheetByName(sheetName);
      if (!destSheet) {
        destSheet = destSpreadsheet.insertSheet(sheetName);
      } else {
        destSheet.clear();  // 既存内容をクリア
      }
      
      // 値を貼り付け
      destSheet.getRange(1, 1, values.length, values[0].length).setValues(values);
      importedSheets.push(sheetName);
    });
    
    return {
      sheetCount: importedSheets.length,
      sheets: importedSheets
    };
    
  } finally {
    // 一時ファイル削除
    try {
      Drive.Files.remove(tempFile.id);
    } catch (e) {
      console.log('一時ファイル削除エラー（無視）:', e);
    }
  }
}

/**
 * 毎日定時実行用のトリガー設定（1回だけ実行）
 */
function setupDailyTrigger() {
  // 既存トリガー削除
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === 'importExcel') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  // 毎日午前9時に実行
  ScriptApp.newTrigger('importExcel')
    .timeBased()
    .atHour(9)
    .everyDays(1)
    .create();
  
  console.log('毎日午前9時に実行するよう設定しました');
}

/**
 * 1時間ごと実行用のトリガー設定
 */
function setupHourlyTrigger() {
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === 'importExcel') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  ScriptApp.newTrigger('importExcel')
    .timeBased()
    .everyHours(1)
    .create();
  
  console.log('1時間ごとに実行するよう設定しました');
}

/**
 * Googleスプレッドシートを直接処理（値のみ取得）
 */
function processGoogleSpreadsheet(fileId) {
  const sourceSpreadsheet = SpreadsheetApp.openById(fileId);
  const destSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const importedSheets = [];
  
  const sourceSheets = sourceSpreadsheet.getSheets();
  
  sourceSheets.forEach(sourceSheet => {
    const sheetName = sourceSheet.getName();
    const values = sourceSheet.getDataRange().getValues();
    
    // 空シートはスキップ
    if (values.length === 0) return;
    
    // 取り込み先シート（同名があれば上書き、なければ新規作成）
    let destSheet = destSpreadsheet.getSheetByName(sheetName);
    if (!destSheet) {
      destSheet = destSpreadsheet.insertSheet(sheetName);
    } else {
      destSheet.clear();  // 既存内容をクリア
    }
    
    // 値を貼り付け
    destSheet.getRange(1, 1, values.length, values[0].length).setValues(values);
    importedSheets.push(sheetName);
  });
  
  return {
    sheetCount: importedSheets.length,
    sheets: importedSheets
  };
}