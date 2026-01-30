/**
 * Slack/Google Drive 命名ルール監査システム - 初期設定
 * スプレッドシートとGASプロジェクトの初期セットアップを自動化
 */

// =================================================================
// 初期設定実行関数
// =================================================================
function setupSpreadsheetAndSystem() {
  try {
    // 新規スプレッドシート作成
    const spreadsheet = createNewSpreadsheet();
    
    // 各シートを作成
    setupConfigSheet(spreadsheet);
    setupRulesSheet(spreadsheet);
    setupWhitelistSheet(spreadsheet);
    setupSlackChannelsSheet(spreadsheet);
    setupDriveSharedDrivesSheet(spreadsheet);
    setupDriveFoldersSheet(spreadsheet);
    setupViolationsLogSheet(spreadsheet);
    setupReportLastRunSheet(spreadsheet);
    
    // GASプロパティストアの設定
    setupScriptProperties();
    
    // Drive APIを有効化するための指示を表示
    showSetupInstructions(spreadsheet);
    
    return spreadsheet.getUrl();
  } catch (error) {
    console.error('Setup error:', error);
    throw error;
  }
}

// =================================================================
// スプレッドシート作成
// =================================================================
function createNewSpreadsheet() {
  const spreadsheetName = `[命名監査] Slack/Drive Naming Audit - ${new Date().toISOString().split('T')[0]}`;
  const spreadsheet = SpreadsheetApp.create(spreadsheetName);
  
  // デフォルトシートを削除
  const defaultSheet = spreadsheet.getSheets()[0];
  if (spreadsheet.getSheets().length > 1) {
    spreadsheet.deleteSheet(defaultSheet);
  }
  
  return spreadsheet;
}

// =================================================================
// Config シート作成
// =================================================================
function setupConfigSheet(spreadsheet) {
  const sheet = spreadsheet.insertSheet('Config');
  
  const headers = [
    ['設定項目', '値', '説明'],
    ['実行モード', 'FULL', 'FULL: 全取得, INCREMENTAL: 差分のみ'],
    ['Slack対象', 'PUBLIC', 'PUBLIC, PRIVATE, BOTH, NONE'],
    ['アーカイブチャンネル含む', 'FALSE', 'TRUE/FALSE'],
    ['Drive対象', 'SHARED_DRIVES', 'SHARED_DRIVES, MY_DRIVE, BOTH'],
    ['最大フォルダ深度', '10', '取得する最大階層数'],
    ['通知先メール', '', 'カンマ区切りで複数指定可'],
    ['除外パスRegex', '^(Archive|Backup|Old).*', '除外するフォルダパターン'],
    ['実行時間制限(秒)', '300', 'GAS実行時間制限（最大360）'],
    ['デバッグモード', 'FALSE', 'TRUE: 詳細ログ出力']
  ];
  
  sheet.getRange(1, 1, headers.length, 3).setValues(headers);
  
  // ヘッダー装飾
  sheet.getRange('A1:C1').setBackground('#4285F4').setFontColor('white').setFontWeight('bold');
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, 3);
}

// =================================================================
// Rules シート作成
// =================================================================
function setupRulesSheet(spreadsheet) {
  const sheet = spreadsheet.insertSheet('Rules');
  
  const headers = [
    ['RuleName', 'Target', 'Regex', 'Severity', 'Priority', 'Description', 'Enabled']
  ];
  
  const sampleRules = [
    ['Slack-ProjectChannel', 'Slack', '^(prj|proj|project)-[a-z0-9-]{3,40}$', 'ERROR', 1, 'プロジェクトチャンネルは prj- で始まる必要があります', 'TRUE'],
    ['Slack-DepartmentChannel', 'Slack', '^(dep|dept|team)-[a-z0-9-]{3,40}$', 'WARN', 2, '部署チャンネルは dep- または team- で始まる必要があります', 'TRUE'],
    ['Slack-GeneralFormat', 'Slack', '^[a-z][a-z0-9-_]{2,80}$', 'WARN', 10, 'チャンネル名は小文字英数字とハイフンのみ', 'TRUE'],
    ['Drive-ProjectDrive', 'Drive', '^PRJ-[0-9]{4}-[A-Za-z0-9 _()-]{3,50}$', 'ERROR', 1, '共有ドライブはPRJ-YYYY-形式', 'TRUE'],
    ['Drive-DepartmentDrive', 'Drive', '^DEPT-[A-Z]{2,10}-[A-Za-z0-9 _()-]{3,50}$', 'WARN', 2, '部署ドライブはDEPT-形式', 'TRUE'],
    ['DriveFolder-DatePrefix', 'DriveFolder', '^[0-9]{8}_.*', 'WARN', 1, '日付プレフィックス形式（YYYYMMDD_）', 'TRUE'],
    ['DriveFolder-CategoryPrefix', 'DriveFolder', '^[A-Z]{2,6}_.*', 'WARN', 2, 'カテゴリプレフィックス形式', 'TRUE'],
    ['DriveFolder-General', 'DriveFolder', '^[A-Za-z0-9][A-Za-z0-9 _()-]{2,100}$', 'WARN', 10, '一般的な命名規則', 'TRUE']
  ];
  
  const data = [...headers, ...sampleRules];
  sheet.getRange(1, 1, data.length, 7).setValues(data);
  
  // ヘッダー装飾
  sheet.getRange('A1:G1').setBackground('#34A853').setFontColor('white').setFontWeight('bold');
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, 7);
}

// =================================================================
// Whitelist シート作成
// =================================================================
function setupWhitelistSheet(spreadsheet) {
  const sheet = spreadsheet.insertSheet('Whitelist');
  
  const headers = [
    ['Type', 'Pattern', 'IsRegex', 'Reason', 'ExpiryDate', 'AddedBy', 'AddedDate']
  ];
  
  const sampleWhitelist = [
    ['Slack', 'general', 'FALSE', 'デフォルトチャンネル', '', 'System', new Date().toISOString()],
    ['Slack', 'random', 'FALSE', 'デフォルトチャンネル', '', 'System', new Date().toISOString()],
    ['Drive', 'マイドライブ', 'FALSE', 'システムフォルダ', '', 'System', new Date().toISOString()],
    ['DriveFolder', '^\\.(config|settings|cache)$', 'TRUE', '隠しフォルダ', '', 'System', new Date().toISOString()]
  ];
  
  const data = [...headers, ...sampleWhitelist];
  sheet.getRange(1, 1, data.length, 7).setValues(data);
  
  // ヘッダー装飾
  sheet.getRange('A1:G1').setBackground('#FBBC04').setFontColor('black').setFontWeight('bold');
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, 7);
}

// =================================================================
// Slack_Channels シート作成
// =================================================================
function setupSlackChannelsSheet(spreadsheet) {
  const sheet = spreadsheet.insertSheet('Slack_Channels');
  
  const headers = [
    ['ChannelID', 'ChannelName', 'IsPrivate', 'IsArchived', 'MemberCount', 
     'Created', 'LastChecked', 'Violation', 'ViolationType', 'ViolationMessage', 'MatchedRule']
  ];
  
  sheet.getRange(1, 1, 1, headers[0].length).setValues(headers);
  sheet.getRange('A1:K1').setBackground('#4285F4').setFontColor('white').setFontWeight('bold');
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(2);
}

// =================================================================
// Drive_SharedDrives シート作成
// =================================================================
function setupDriveSharedDrivesSheet(spreadsheet) {
  const sheet = spreadsheet.insertSheet('Drive_SharedDrives');
  
  const headers = [
    ['DriveID', 'DriveName', 'Created', 'LastChecked', 
     'Violation', 'ViolationType', 'ViolationMessage', 'MatchedRule']
  ];
  
  sheet.getRange(1, 1, 1, headers[0].length).setValues(headers);
  sheet.getRange('A1:H1').setBackground('#34A853').setFontColor('white').setFontWeight('bold');
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(2);
}

// =================================================================
// Drive_Folders シート作成
// =================================================================
function setupDriveFoldersSheet(spreadsheet) {
  const sheet = spreadsheet.insertSheet('Drive_Folders');
  
  const headers = [
    ['FolderID', 'FolderName', 'ParentID', 'DriveID', 'FullPath', 'Depth',
     'Created', 'Modified', 'LastChecked', 
     'Violation', 'ViolationType', 'ViolationMessage', 'MatchedRule']
  ];
  
  sheet.getRange(1, 1, 1, headers[0].length).setValues(headers);
  sheet.getRange('A1:M1').setBackground('#34A853').setFontColor('white').setFontWeight('bold');
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(2);
}

// =================================================================
// Violations_Log シート作成
// =================================================================
function setupViolationsLogSheet(spreadsheet) {
  const sheet = spreadsheet.insertSheet('Violations_Log');
  
  const headers = [
    ['ViolationID', 'Type', 'ResourceID', 'ResourceName', 'FullPath',
     'ViolationType', 'ViolationMessage', 'MatchedRule', 'Severity',
     'FirstDetected', 'LastConfirmed', 'Status', 'ResolvedDate', 'Notes']
  ];
  
  sheet.getRange(1, 1, 1, headers[0].length).setValues(headers);
  sheet.getRange('A1:N1').setBackground('#EA4335').setFontColor('white').setFontWeight('bold');
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(4);
}

// =================================================================
// Report_LastRun シート作成
// =================================================================
function setupReportLastRunSheet(spreadsheet) {
  const sheet = spreadsheet.insertSheet('Report_LastRun');
  
  const headers = [
    ['実行項目', '値'],
    ['最終実行日時', ''],
    ['実行ステータス', ''],
    ['Slackチャンネル取得数', '0'],
    ['Drive取得数', '0'],
    ['フォルダ取得数', '0'],
    ['新規違反検出数', '0'],
    ['継続違反数', '0'],
    ['解決済み違反数', '0'],
    ['エラー内容', ''],
    ['実行時間(秒)', '0']
  ];
  
  sheet.getRange(1, 1, headers.length, 2).setValues(headers);
  sheet.getRange('A1:B1').setBackground('#9E9E9E').setFontColor('white').setFontWeight('bold');
  sheet.autoResizeColumns(1, 2);
}

// =================================================================
// GASプロパティストア設定
// =================================================================
function setupScriptProperties() {
  const scriptProperties = PropertiesService.getScriptProperties();
  
  // 初期プロパティ設定
  scriptProperties.setProperties({
    'SLACK_BOT_TOKEN': '',  // 後で設定が必要
    'SPREADSHEET_ID': '',    // 実行時に自動設定
    'EXECUTION_MODE': 'MANUAL',
    'LAST_EXECUTION': '',
    'CHECKPOINT_DATA': '{}'
  });
}

// =================================================================
// セットアップ完了後の指示表示
// =================================================================
function showSetupInstructions(spreadsheet) {
  const message = `
========================================
セットアップが完了しました！
========================================

スプレッドシートURL:
${spreadsheet.getUrl()}

【次の手順】

1. Slack Bot Token の設定:
   - https://api.slack.com/apps でアプリを作成
   - OAuth & Permissions で以下のスコープを追加:
     * channels:read
     * groups:read (プライベートチャンネル用)
     * users:read
   - Bot User OAuth Token をコピー
   - GASエディタでプロパティに設定

2. Google Drive API の有効化:
   - GASエディタで「サービス」をクリック
   - "Drive API" を追加

3. Config シートの設定:
   - 通知先メールアドレスを入力
   - その他の設定を必要に応じて調整

4. 実行権限の付与:
   - mainAudit() 関数を一度手動実行
   - 権限要求ダイアログで承認

5. トリガーの設定:
   - エディタで「トリガー」をクリック
   - mainAudit を毎日実行に設定（推奨: 朝8:30）

========================================
  `;
  
  console.log(message);
  
  // スプレッドシートにも説明シートを追加
  const instructionSheet = spreadsheet.insertSheet('📋 Setup Instructions');
  instructionSheet.getRange('A1').setValue(message);
  instructionSheet.getRange('A1').setWrap(true);
  instructionSheet.setColumnWidth(1, 600);
}

// =================================================================
// 手動実行用：スプレッドシートを開く
// =================================================================
function openSpreadsheet() {
  const url = setupSpreadsheetAndSystem();
  const html = `<script>window.open('${url}', '_blank');google.script.host.close();</script>`;
  const userInterface = HtmlService.createHtmlOutput(html);
  SpreadsheetApp.getUi().showModalDialog(userInterface, 'Opening Spreadsheet...');
}