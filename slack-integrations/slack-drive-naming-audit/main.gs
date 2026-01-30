/**
 * メイン実行関数と共通ユーティリティ
 * 全体の制御と共通関数
 */

// =================================================================
// メイン監査実行関数
// =================================================================
function mainAudit() {
  const startTime = new Date();
  console.log('========================================');
  console.log('命名ルール監査を開始します...');
  console.log(`実行時刻: ${startTime.toISOString()}`);
  console.log('========================================');
  
  try {
    // スプレッドシートIDの取得または設定
    const spreadsheetId = getSpreadsheetId();
    if (!spreadsheetId) {
      throw new Error('スプレッドシートIDが設定されていません');
    }
    
    // 1. Slackチャンネルの取得と判定
    console.log('\n[1/4] Slackチャンネルを処理中...');
    fetchSlackChannels();
    
    // 2. Google Driveデータの取得と判定
    console.log('\n[2/4] Google Driveデータを処理中...');
    fetchGoogleDriveData();
    
    // 3. 違反ログの更新
    console.log('\n[3/4] 違反ログを更新中...');
    updateViolationsLog();
    
    // 4. レポート生成と送信
    console.log('\n[4/4] レポートを生成・送信中...');
    generateAndSendReport();
    
    // 実行時間の記録
    const endTime = new Date();
    const executionTime = (endTime - startTime) / 1000;
    
    // 実行結果を更新
    updateExecutionTime(executionTime);
    
    console.log('\n========================================');
    console.log('命名ルール監査が正常に完了しました');
    console.log(`実行時間: ${executionTime}秒`);
    console.log('========================================');
    
  } catch (error) {
    console.error('\nエラーが発生しました:', error);
    
    // エラー情報を記録
    recordError(error);
    
    // エラー通知
    sendErrorNotification(error);
    
    throw error;
  }
}

// =================================================================
// 手動実行用関数（テスト用）
// =================================================================
function testSlackOnly() {
  console.log('Slackチャンネルのみテスト');
  fetchSlackChannels();
  console.log('テスト完了');
}

function testDriveOnly() {
  console.log('Google Driveのみテスト');
  fetchGoogleDriveData();
  console.log('テスト完了');
}

function testReportOnly() {
  console.log('レポート生成のみテスト');
  updateViolationsLog();
  generateAndSendReport();
  console.log('テスト完了');
}

// =================================================================
// トリガー設定関数
// =================================================================
function setupTriggers() {
  // 既存のトリガーを削除
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'mainAudit') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  // 新しいトリガーを設定（毎日朝8:30）
  ScriptApp.newTrigger('mainAudit')
    .timeBased()
    .atHour(8)
    .nearMinute(30)
    .everyDays(1)
    .create();
  
  console.log('トリガーを設定しました: 毎日 8:30 に実行');
}

// =================================================================
// 共通ユーティリティ関数
// =================================================================
function getSpreadsheetId() {
  // プロパティから取得
  let spreadsheetId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  
  // プロパティにない場合は、現在のスプレッドシートから取得
  if (!spreadsheetId) {
    try {
      const activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      if (activeSpreadsheet) {
        spreadsheetId = activeSpreadsheet.getId();
        // プロパティに保存
        PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', spreadsheetId);
      }
    } catch (e) {
      // スタンドアロンスクリプトの場合はアクティブスプレッドシートがない
      console.log('アクティブスプレッドシートが見つかりません');
    }
  }
  
  return spreadsheetId;
}

function getConfig() {
  const spreadsheet = SpreadsheetApp.openById(getSpreadsheetId());
  const configSheet = spreadsheet.getSheetByName('Config');
  
  if (!configSheet) {
    throw new Error('Configシートが見つかりません');
  }
  
  const configData = configSheet.getRange(2, 1, 10, 2).getValues();
  const config = {};
  
  const configMap = {
    '実行モード': 'executionMode',
    'Slack対象': 'slackTarget',
    'アーカイブチャンネル含む': 'includeArchived',
    'Drive対象': 'driveTarget',
    '最大フォルダ深度': 'maxFolderDepth',
    '通知先メール': 'notificationEmail',
    '除外パスRegex': 'excludePathRegex',
    '実行時間制限(秒)': 'executionTimeLimit',
    'デバッグモード': 'debugMode'
  };
  
  for (const row of configData) {
    const key = configMap[row[0]];
    if (key) {
      config[key] = row[1];
    }
  }
  
  return config;
}

function isExecutionTimeLimitNear() {
  const config = getConfig();
  const limit = parseInt(config.executionTimeLimit) || 300; // デフォルト5分
  const startTime = PropertiesService.getScriptProperties().getProperty('EXECUTION_START_TIME');
  
  if (!startTime) {
    // 開始時刻を記録
    PropertiesService.getScriptProperties().setProperty('EXECUTION_START_TIME', new Date().getTime().toString());
    return false;
  }
  
  const elapsed = (new Date().getTime() - parseInt(startTime)) / 1000;
  return elapsed > limit;
}

function saveCheckpoint(data) {
  PropertiesService.getScriptProperties().setProperty('CHECKPOINT_DATA', JSON.stringify(data));
}

function getCheckpoint() {
  const data = PropertiesService.getScriptProperties().getProperty('CHECKPOINT_DATA');
  return data ? JSON.parse(data) : {};
}

function clearCheckpoint() {
  PropertiesService.getScriptProperties().deleteProperty('CHECKPOINT_DATA');
  PropertiesService.getScriptProperties().deleteProperty('EXECUTION_START_TIME');
}

// =================================================================
// エラーハンドリング
// =================================================================
function recordError(error) {
  const spreadsheet = SpreadsheetApp.openById(getSpreadsheetId());
  const sheet = spreadsheet.getSheetByName('Report_LastRun');
  
  if (sheet) {
    sheet.getRange(2, 2).setValue(new Date());
    sheet.getRange(3, 2).setValue('ERROR');
    sheet.getRange(10, 2).setValue(error.toString());
  }
}

function sendErrorNotification(error) {
  const config = getConfig();
  const recipients = config.notificationEmail;
  
  if (!recipients || recipients.trim() === '') {
    return;
  }
  
  const subject = '[命名監査] エラー通知';
  const body = `
命名ルール監査システムでエラーが発生しました。

エラー内容:
${error.toString()}

スタックトレース:
${error.stack || 'N/A'}

実行時刻: ${new Date().toISOString()}

スプレッドシート:
${SpreadsheetApp.openById(getSpreadsheetId()).getUrl()}
`;
  
  try {
    MailApp.sendEmail({
      to: recipients,
      subject: subject,
      body: body,
      name: 'Naming Audit System'
    });
  } catch (e) {
    console.error('エラー通知の送信に失敗:', e);
  }
}

function updateExecutionTime(executionTime) {
  const spreadsheet = SpreadsheetApp.openById(getSpreadsheetId());
  const sheet = spreadsheet.getSheetByName('Report_LastRun');
  
  if (sheet) {
    sheet.getRange(11, 2).setValue(executionTime);
  }
}

// =================================================================
// 初回セットアップ用メニュー
// =================================================================
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('⚙️ 命名監査システム')
    .addItem('🚀 初期セットアップ実行', 'setupSpreadsheetAndSystem')
    .addSeparator()
    .addItem('🔄 手動実行 (全体)', 'mainAudit')
    .addItem('💬 Slackのみテスト', 'testSlackOnly')
    .addItem('📁 Driveのみテスト', 'testDriveOnly')
    .addItem('📧 レポートのみテスト', 'testReportOnly')
    .addSeparator()
    .addItem('⏰ トリガー設定', 'setupTriggers')
    .addToUi();
}