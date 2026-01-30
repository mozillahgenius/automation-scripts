/**
 * 初期設定ガイド用スクリプト
 * システムの初期設定を簡単に行うための関数群
 */

/**
 * 初期設定を行う関数
 * この関数を最初に実行してください
 */
function initialSetup() {
  console.log('=== 初期設定開始 ===');
  
  try {
    // 1. 設定スプレッドシートの作成
    console.log('1. 設定スプレッドシートを作成中...');
    const spreadsheet = createConfigSpreadsheet();
    console.log(`✓ スプレッドシート作成完了: ${spreadsheet.getUrl()}`);
    // Script Properties に ID を登録
    if (typeof setConfigSpreadsheetId === 'function') {
      setConfigSpreadsheetId(spreadsheet.getId());
      console.log('✓ CONFIG_SPREADSHEET_ID を Script Properties に登録しました');
    }
    
    // 2. デフォルト設定の書き込み
    console.log('2. デフォルト設定を書き込み中...');
    setupDefaultConfig(spreadsheet);
    console.log('✓ デフォルト設定の書き込み完了');
    
    // 3. 必要なスコープの確認
    console.log('3. 必要な権限を確認中...');
    checkRequiredScopes();
    console.log('✓ 権限確認完了');
    
    // 4. テスト実行
    console.log('4. システムテストを実行中...');
    testBasicFunctions();
    console.log('✓ システムテスト完了');
    
    console.log('\n=== 初期設定完了 ===');
    console.log('次のステップ:');
    console.log('1. スプレッドシートのConfigシートを開いて設定を確認・編集してください');
    console.log(`   URL: ${spreadsheet.getUrl()}`);
    console.log('2. 特にGCP_PROJECT_IDは必ず設定してください');
    console.log('3. 設定後、testSystem()関数を実行して動作確認してください');
    console.log('4. 問題なければsetupTrigger()でトリガーを設定してください');
    
    return spreadsheet.getUrl();
    
  } catch (error) {
    console.error('初期設定エラー:', error);
    throw error;
  }
}

/**
 * 設定スプレッドシートを作成（既存の場合はそれを使用）
 */
function createConfigSpreadsheet() {
  const files = DriveApp.getFilesByName(CONFIG_SPREADSHEET_NAME);
  
  if (files.hasNext()) {
    console.log('既存のスプレッドシートを使用します');
    return SpreadsheetApp.open(files.next());
  }
  
  // 新規作成
  const spreadsheet = SpreadsheetApp.create(CONFIG_SPREADSHEET_NAME);
  
  // 必要なシートを作成
  const configSheet = spreadsheet.getActiveSheet();
  configSheet.setName(CONFIG_SHEET_NAME);
  
  const logSheet = spreadsheet.insertSheet(LOG_SHEET_NAME);
  
  return spreadsheet;
}

/**
 * デフォルト設定を書き込む
 */
function setupDefaultConfig(spreadsheet) {
  const sheet = spreadsheet.getSheetByName(CONFIG_SHEET_NAME);
  
  // ヘッダーと初期値を設定
  const configData = [
    ['Setting', 'Value', 'Description'],
    ['CalendarId', '', 'カレンダーID（空の場合はデフォルトカレンダー使用）'],
    ['GCP_PROJECT_ID', 'your-project-id', '【必須】GCPプロジェクトID'],
    ['MinutesPrompt', '# 役割\nあなたは、プロの書記です。\n\n# 指示\n以下の会議トランスクリプトを読み、下記のフォーマットに従って簡潔で分かりやすい日本語の議事録を作成してください。\n\n# フォーマット\n## 1. 会議の概要\n- **会議名**: (会議のタイトルを記載)\n- **日時**: (YYYY/MM/DD HH:MM)\n- **参加者**: (参加者名を「、」で区切って列挙)\n\n## 2. 決定事項\n- (決定したことを箇条書きで記載)\n- (複数ある場合は全て記載)\n\n## 3. アクションアイテム\n- **担当者名**: (タスク内容を具体的に記載) (期限: YYYY/MM/DD)\n- (複数ある場合は全て記載)\n\n## 4. 議論内容の要約\n- (主要な議題や議論の流れを箇条書きで要約)\n\n# 制約\n- トランスクリプトに存在しない情報は記載しないでください。\n- 各項目で記載することがない場合は、「特になし」と記載してください。\n- 参加者名は敬称を省略してください。\n\n# 会議トランスクリプト\n---\n{transcript_text}', '議事録生成用プロンプト'],
    ['AdditionalRecipients', '', '追加のメール送信先（カンマ区切り）'],
    ['CheckHours', '1', '何時間前までの会議をチェックするか']
  ];
  
  // データを書き込み
  sheet.getRange(1, 1, configData.length, 3).setValues(configData);
  
  // 書式設定
  sheet.getRange('1:1').setFontWeight('bold');
  sheet.setColumnWidth(1, 200);
  sheet.setColumnWidth(2, 400);
  sheet.setColumnWidth(3, 300);
  
  // ログシートの設定
  const logSheet = spreadsheet.getSheetByName(LOG_SHEET_NAME);
  logSheet.getRange('A1:E1').setValues([['Event ID', 'Event Title', 'Processed At', 'Status', 'Error Message']]);
  logSheet.getRange('1:1').setFontWeight('bold');
  logSheet.setColumnWidth(1, 150);
  logSheet.setColumnWidth(2, 300);
  logSheet.setColumnWidth(3, 150);
  logSheet.setColumnWidth(4, 100);
  logSheet.setColumnWidth(5, 300);
}

/**
 * 必要なスコープを確認
 */
function checkRequiredScopes() {
  const requiredScopes = [
    'https://www.googleapis.com/auth/calendar.readonly',
    'https://www.googleapis.com/auth/gmail.send',
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive',
    'https://www.googleapis.com/auth/script.external_request'
  ];
  
  console.log('必要なスコープ:');
  requiredScopes.forEach(scope => {
    console.log(`  - ${scope}`);
  });
  
  console.log('\n注意: Meet APIへのアクセスには追加の設定が必要です');
  console.log('GCPコンソールでMeet APIを有効化してください');
}

/**
 * 基本機能のテスト
 */
function testBasicFunctions() {
  // カレンダーアクセステスト
  try {
    const calendar = CalendarApp.getDefaultCalendar();
    console.log(`✓ カレンダーアクセス成功: ${calendar.getName()}`);
  } catch (error) {
    console.error('✗ カレンダーアクセス失敗:', error.message);
  }
  
  // スプレッドシートアクセステスト
  try {
    const spreadsheet = getConfigSpreadsheet();
    console.log(`✓ スプレッドシートアクセス成功: ${spreadsheet.getName()}`);
  } catch (error) {
    console.error('✗ スプレッドシートアクセス失敗:', error.message);
  }
  
  // Gmail アクセステスト
  try {
    const email = Session.getActiveUser().getEmail();
    console.log(`✓ Gmailアクセス成功: ${email}`);
  } catch (error) {
    console.error('✗ Gmailアクセス失敗:', error.message);
  }
}

/**
 * トリガーを設定する関数
 */
function setupTrigger() {
  // 既存のトリガーを削除
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'main') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  // 新しいトリガーを作成（1時間ごと）
  ScriptApp.newTrigger('main')
    .timeBased()
    .everyHours(1)
    .create();
  
  console.log('✓ トリガー設定完了: 1時間ごとに実行されます');
}

/**
 * トリガーを削除する関数
 */
function removeTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  let count = 0;
  
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'main') {
      ScriptApp.deleteTrigger(trigger);
      count++;
    }
  });
  
  console.log(`✓ ${count}個のトリガーを削除しました`);
}

/**
 * 設定を確認する関数
 */
function checkCurrentConfig() {
  console.log('=== 現在の設定 ===');
  
  try {
    const config = getConfig();
    
    console.log('CalendarId:', config.calendarId || 'デフォルトカレンダー使用');
    console.log('CheckHours:', config.checkHours);
    console.log('AdditionalRecipients:', config.additionalRecipients || 'なし');
    console.log('MinutesPrompt:', config.minutesPrompt ? '設定済み' : '未設定');
    
    // GCPプロジェクトID確認
    if (GCP_PROJECT_ID === 'your-project-id') {
      console.warn('⚠️ GCP_PROJECT_IDが設定されていません！');
      console.warn('   Code.gsの定数を編集するか、環境変数を設定してください');
    } else {
      console.log('GCP_PROJECT_ID:', GCP_PROJECT_ID);
    }
    
    // トリガー確認
    const triggers = ScriptApp.getProjectTriggers();
    const mainTriggers = triggers.filter(t => t.getHandlerFunction() === 'main');
    
    if (mainTriggers.length > 0) {
      console.log('トリガー: 設定済み');
    } else {
      console.log('トリガー: 未設定');
    }
    
  } catch (error) {
    console.error('設定の読み込みエラー:', error);
  }
}