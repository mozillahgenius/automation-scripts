/**
 * Slack/Google Drive 命名ルール監査システム - 統合版
 * 全機能を1つのファイルに統合
 */

// =====================================================================
// メイン実行関数
// =====================================================================
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

    // 1. Slackチャンネルの取得と判定（エラーでも続行）
    console.log('\n[1/4] Slackチャンネルを処理中...');
    try {
      fetchSlackChannels();
    } catch (slackError) {
      console.error('Slack処理エラー（続行）:', slackError);
    }

    // 2. Google Driveデータの取得と判定（エラーでも続行）
    console.log('\n[2/4] Google Driveデータを処理中...');
    try {
      fetchGoogleDriveData();
    } catch (driveError) {
      console.error('Drive処理エラー（続行）:', driveError);
    }

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

// =====================================================================
// 初期設定関数
// =====================================================================
function setupSpreadsheetAndSystem() {
  try {
    // 現在のスプレッドシートを使用（新規作成しない）
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

    if (!spreadsheet) {
      throw new Error('スプレッドシートが開かれていません。GASエディタをスプレッドシートから開いてください。');
    }

    // 既存シートをクリア（必要に応じて）
    clearExistingSheets(spreadsheet);

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

function clearExistingSheets(spreadsheet) {
  const sheetsToKeep = ['シート1']; // デフォルトシート名
  const sheetsToCreate = [
    'Config', 'Rules', 'Whitelist', 'Slack_Channels',
    'Drive_SharedDrives', 'Drive_Folders', 'Violations_Log', 'Report_LastRun'
  ];

  // 既存のシートを確認して削除
  const existingSheets = spreadsheet.getSheets();
  existingSheets.forEach(sheet => {
    const sheetName = sheet.getName();
    if (sheetsToCreate.includes(sheetName)) {
      spreadsheet.deleteSheet(sheet);
    }
  });
}

function setupConfigSheet(spreadsheet) {
  const sheet = spreadsheet.insertSheet('Config');

  const headers = [
    ['設定項目', '値', '説明'],
    ['実行モード', 'FULL', 'FULL: 全取得, INCREMENTAL: 差分のみ'],
    ['Slack対象', 'PUBLIC', 'PUBLIC, PRIVATE, BOTH, NONE'],
    ['アーカイブチャンネル含む', 'FALSE', 'TRUE/FALSE'],
    ['Drive対象', 'BOTH', 'SHARED_DRIVES, MY_DRIVE, BOTH'],
    ['最大フォルダ深度', '99', '取得する最大階層数（99=無制限）'],
    ['通知先メール', '', 'カンマ区切りで複数指定可'],
    ['除外パスRegex', '^(Archive|Backup|Old|Trash|ゴミ箱).*', '除外するフォルダパターン'],
    ['実行時間制限(秒)', '300', 'GAS実行時間制限（最大360）'],
    ['デバッグモード', 'FALSE', 'TRUE: 詳細ログ出力']
  ];

  sheet.getRange(1, 1, headers.length, 3).setValues(headers);

  // ヘッダー装飾
  sheet.getRange('A1:C1').setBackground('#4285F4').setFontColor('white').setFontWeight('bold');
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, 3);
}

function setupRulesSheet(spreadsheet) {
  const sheet = spreadsheet.insertSheet('Rules');

  const headers = [
    ['RuleName', 'Target', 'Regex', 'Severity', 'Priority', 'Description', 'Enabled']
  ];

  const sampleRules = [
    // Slack命名ルール（チャンネル番号システム）
    ['Slack-DefaultChannel', 'Slack', '^0[0-9]{2}_.*$', 'WARN', 1, 'デフォルトチャンネル（000-099）', 'TRUE'],
    ['Slack-OrgChannel', 'Slack', '^1[0-9]{2}_org_[a-z0-9-ぁ-んァ-ヶー一-龯]+$', 'ERROR', 2, '組織系チャンネル（100-199）は形式: 番号_org_名前', 'TRUE'],
    ['Slack-ProjectChannel', 'Slack', '^[2-8][0-9]{2}_pj_[a-z0-9-ぁ-んァ-ヶー一-龯]+$', 'ERROR', 3, 'プロジェクト系（200-899）は形式: 番号_pj_名前', 'TRUE'],
    ['Slack-PrivateChannel', 'Slack', '^9[0-9]{2}_.*$', 'WARN', 4, 'プライベートチャンネル（900-999）', 'TRUE'],
    // 旧形式（互換性のため残す）
    ['Slack-LegacyProject', 'Slack', '^(prj|proj|project)-[a-z0-9-]{3,40}$', 'WARN', 10, '旧形式: プロジェクトチャンネル', 'FALSE'],
    ['Slack-LegacyDepartment', 'Slack', '^(dep|dept|team)-[a-z0-9-]{3,40}$', 'WARN', 11, '旧形式: 部署チャンネル', 'FALSE'],
    ['Slack-GeneralFormat', 'Slack', '^[a-z][a-z0-9-_]{2,80}$', 'WARN', 20, 'チャンネル名は小文字英数字とハイフンのみ', 'FALSE'],
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

function setupWhitelistSheet(spreadsheet) {
  const sheet = spreadsheet.insertSheet('Whitelist');

  const headers = [
    ['Type', 'Pattern', 'IsRegex', 'Reason', 'ExpiryDate', 'AddedBy', 'AddedDate']
  ];

  const sampleWhitelist = [
    ['Slack', 'general', 'FALSE', 'デフォルトチャンネル', '', 'System', new Date().toISOString()],
    ['Slack', 'random', 'FALSE', 'デフォルトチャンネル', '', 'System', new Date().toISOString()],
    // 既存の番号なしチャンネル（移行期間用）
    ['Slack', '^(支社|営業|管理|経営|開発|広報|人事|総務|経理|法務)', 'TRUE', '既存の部署チャンネル（移行期間）', '', 'System', new Date().toISOString()],
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

function showSetupInstructions(spreadsheet) {
  const message = `
========================================
セットアップが完了しました！
========================================

スプレッドシートURL:
${spreadsheet.getUrl()}

【次の手順】

1. Slack Bot Token の設定（重要）:
   a) Slack アプリの作成:
      - https://api.slack.com/apps で「Create New App」
      - 「From scratch」を選択
      - App Name: 任意（例: Naming Audit Bot）
      - Workspace: 対象のSlackワークスペースを選択

   b) 権限の設定:
      - 左メニューの「OAuth & Permissions」をクリック
      - 「Scopes」セクションの「Bot Token Scopes」に以下を追加:
        * channels:read（公開チャンネル読み取り）
        * groups:read（プライベートチャンネル読み取り）
        * users:read（ユーザー情報読み取り）
      - 上部の「Install to Workspace」をクリック
      - 許可画面で「許可する」をクリック

   c) トークンの取得と設定:
      - 「OAuth Tokens for Your Workspace」セクションの
        「Bot User OAuth Token」をコピー（xoxb-で始まる文字列）

   d) GASに設定:
      - GASエディタで「プロジェクトの設定」（歯車アイコン）
      - 「スクリプト プロパティ」セクションの「プロパティを追加」
      - プロパティ名: SLACK_BOT_TOKEN
      - 値: コピーしたBot Token（xoxb-...）
      - 「スクリプト プロパティを保存」をクリック

2. Google Drive API の有効化:
   - GASエディタで「サービス」をクリック
   - "Drive API" を追加して有効化

3. Config シートの設定:
   - 通知先メールアドレスを入力（カンマ区切りで複数可）
   - その他の設定を必要に応じて調整

4. 実行権限の付与:
   - mainAudit() 関数を一度手動実行
   - 権限要求ダイアログで承認

5. トリガーの設定:
   - エディタで「トリガー」（時計アイコン）をクリック
   - mainAudit を毎日実行に設定（推奨: 朝8:30）

【Slack命名ルールについて】
チャンネル番号体系を使用しています:
- 000-099: デフォルトチャンネル
- 100-199: 組織系（形式: 番号_org_チャンネル名）
- 200-899: プロジェクト系（形式: 番号_pj_チャンネル名）
- 900-999: プライベートチャンネル

========================================
  `;

  console.log(message);

  // 既存の説明シートがあれば削除
  const existingSheet = spreadsheet.getSheetByName('📋 Setup Instructions');
  if (existingSheet) {
    spreadsheet.deleteSheet(existingSheet);
  }

  // スプレッドシートに説明シートを追加
  const instructionSheet = spreadsheet.insertSheet('📋 Setup Instructions');
  instructionSheet.getRange('A1').setValue(message);
  instructionSheet.getRange('A1').setWrap(true);
  instructionSheet.setColumnWidth(1, 600);
}

// =====================================================================
// Slack API 関連関数
// =====================================================================
function fetchSlackChannels() {
  const startTime = new Date();
  console.log('Slackチャンネル取得を開始...');

  try {
    // 設定読み込み
    const config = getConfig();
    const token = PropertiesService.getScriptProperties().getProperty('SLACK_BOT_TOKEN');

    if (!token) {
      throw new Error('Slack Bot Tokenが設定されていません');
    }

    // Slack対象がNONEの場合はスキップ
    if (config.slackTarget === 'NONE') {
      console.log('Slackチャンネル取得はスキップされました（Config設定）');
      return [];
    }

    // チャンネル一覧を取得
    const channels = getAllSlackChannels(token, config);

    // スプレッドシートに保存
    saveSlackChannelsToSheet(channels);

    const executionTime = (new Date() - startTime) / 1000;
    console.log(`Slackチャンネル取得完了: ${channels.length}件 (${executionTime}秒)`);

    return channels;

  } catch (error) {
    console.error('Slackチャンネル取得エラー:', error);
    throw error;
  }
}

function getAllSlackChannels(token, config) {
  const channels = [];
  let cursor = '';
  const limit = 200; // 一度に取得する最大数

  // APIパラメータ設定
  const types = [];
  if (config.slackTarget === 'PUBLIC' || config.slackTarget === 'BOTH') {
    types.push('public_channel');
  }
  if (config.slackTarget === 'PRIVATE' || config.slackTarget === 'BOTH') {
    types.push('private_channel');
  }

  const excludeArchived = config.includeArchived !== 'TRUE';

  do {
    const url = 'https://slack.com/api/conversations.list';
    const params = {
      'types': types.join(','),
      'exclude_archived': excludeArchived,
      'limit': limit
    };

    if (cursor) {
      params.cursor = cursor;
    }

    const response = callSlackAPI(url, params, token);

    if (!response.ok) {
      throw new Error(`Slack APIエラー: ${response.error}`);
    }

    // チャンネル情報を追加
    if (response.channels && response.channels.length > 0) {
      channels.push(...response.channels);
    }

    // ページネーション
    cursor = response.response_metadata && response.response_metadata.next_cursor
      ? response.response_metadata.next_cursor : '';

  } while (cursor);

  return channels;
}

function callSlackAPI(url, params, token) {
  const options = {
    'method': 'get',
    'headers': {
      'Authorization': `Bearer ${token}`,
      'Content-Type': 'application/json'
    },
    'muteHttpExceptions': true
  };

  // URLパラメータを構築
  const queryString = Object.keys(params)
    .map(key => `${key}=${encodeURIComponent(params[key])}`)
    .join('&');

  const fullUrl = `${url}?${queryString}`;

  try {
    const response = UrlFetchApp.fetch(fullUrl, options);
    const responseCode = response.getResponseCode();

    // レート制限処理
    if (responseCode === 429) {
      const retryAfter = response.getHeaders()['Retry-After'] || 60;
      console.log(`Rate limited. Waiting ${retryAfter} seconds...`);
      Utilities.sleep(retryAfter * 1000);
      return callSlackAPI(url, params, token); // リトライ
    }

    return JSON.parse(response.getContentText());

  } catch (error) {
    console.error('Slack API呼び出しエラー:', error);
    throw error;
  }
}

function saveSlackChannelsToSheet(channels) {
  const spreadsheet = SpreadsheetApp.openById(getSpreadsheetId());
  const sheet = spreadsheet.getSheetByName('Slack_Channels');

  if (!sheet) {
    throw new Error('Slack_Channels シートが見つかりません');
  }

  // 既存データをクリア（ヘッダー以外）
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clear();
  }

  // データを整形
  const now = new Date();
  const data = channels.map(channel => [
    channel.id,
    channel.name,
    channel.is_private ? 'TRUE' : 'FALSE',
    channel.is_archived ? 'TRUE' : 'FALSE',
    channel.num_members || 0,
    channel.created ? new Date(channel.created * 1000) : '',
    now,
    '', // Violation（後で判定）
    '', // ViolationType
    '', // ViolationMessage
    ''  // MatchedRule
  ]);

  if (data.length > 0) {
    sheet.getRange(2, 1, data.length, data[0].length).setValues(data);
  }

  // 違反判定を実行
  validateSlackChannels();
}

function validateSlackChannels() {
  const spreadsheet = SpreadsheetApp.openById(getSpreadsheetId());
  const channelsSheet = spreadsheet.getSheetByName('Slack_Channels');
  const rulesSheet = spreadsheet.getSheetByName('Rules');
  const whitelistSheet = spreadsheet.getSheetByName('Whitelist');

  if (!channelsSheet || !rulesSheet || !whitelistSheet) {
    throw new Error('必要なシートが見つかりません');
  }

  // ルールとホワイトリストを取得
  const rules = getSlackRules(rulesSheet);
  const whitelist = getSlackWhitelist(whitelistSheet);

  // チャンネルデータを取得
  const lastRow = channelsSheet.getLastRow();
  if (lastRow <= 1) return;

  const channelsData = channelsSheet.getRange(2, 1, lastRow - 1, 11).getValues();
  const validationResults = [];

  // 各チャンネルを判定
  for (let i = 0; i < channelsData.length; i++) {
    const channelName = channelsData[i][1]; // ChannelName列
    const result = validateName(channelName, rules, whitelist, 'Slack');

    validationResults.push([
      result.violation ? 'TRUE' : 'FALSE',
      result.violationType || '',
      result.violationMessage || '',
      result.matchedRule || ''
    ]);
  }

  // 結果をシートに書き込み
  if (validationResults.length > 0) {
    channelsSheet.getRange(2, 8, validationResults.length, 4).setValues(validationResults);
  }
}

function getSlackRules(rulesSheet) {
  const lastRow = rulesSheet.getLastRow();
  if (lastRow <= 1) return [];

  const rulesData = rulesSheet.getRange(2, 1, lastRow - 1, 7).getValues();

  return rulesData
    .filter(row => row[1] === 'Slack' && row[6] === 'TRUE') // Target=Slack, Enabled=TRUE
    .map(row => ({
      ruleName: row[0],
      regex: row[2],
      severity: row[3],
      priority: row[4],
      description: row[5]
    }))
    .sort((a, b) => a.priority - b.priority); // 優先度順にソート
}

function getSlackWhitelist(whitelistSheet) {
  const lastRow = whitelistSheet.getLastRow();
  if (lastRow <= 1) return [];

  const whitelistData = whitelistSheet.getRange(2, 1, lastRow - 1, 5).getValues();
  const now = new Date();

  return whitelistData
    .filter(row => {
      // Type=Slackで、期限がないか期限内のもの
      if (row[0] !== 'Slack') return false;
      if (row[4] && new Date(row[4]) < now) return false; // 期限切れ
      return true;
    })
    .map(row => ({
      pattern: row[1],
      isRegex: row[2] === 'TRUE',
      reason: row[3]
    }));
}

// =====================================================================
// Google Drive API 関連関数
// =====================================================================
function fetchGoogleDriveData() {
  const startTime = new Date();
  console.log('Google Driveデータ取得を開始...');

  try {
    // 設定読み込み
    const config = getConfig();

    // 共有ドライブ一覧を取得
    const sharedDrives = fetchSharedDrives(config);

    // 各ドライブのフォルダを取得
    const allFolders = fetchAllFolders(sharedDrives, config);

    const executionTime = (new Date() - startTime) / 1000;
    console.log(`Driveデータ取得完了: ドライブ${sharedDrives.length}件, フォルダ${allFolders.length}件 (${executionTime}秒)`);

    return {
      sharedDrives: sharedDrives,
      folders: allFolders
    };

  } catch (error) {
    console.error('Driveデータ取得エラー:', error);
    throw error;
  }
}

function fetchSharedDrives(config) {
  console.log('共有ドライブ一覧を取得中...');

  if (config.driveTarget === 'MY_DRIVE') {
    // マイドライブのみの場合
    return [];
  }

  const sharedDrives = [];

  try {
    // Drive API v2では共有ドライブ（旧チームドライブ）は teamdrives を使用
    let pageToken = null;

    do {
      const params = {
        maxResults: 100
      };

      if (pageToken) {
        params.pageToken = pageToken;
      }

      const response = Drive.Teamdrives.list(params);

      if (response.items && response.items.length > 0) {
        // itemsをsharedDrives形式に変換
        response.items.forEach(drive => {
          sharedDrives.push({
            id: drive.id,
            name: drive.name,
            createdDate: drive.createdDate || new Date()
          });
        });
      }

      pageToken = response.nextPageToken;

    } while (pageToken);

  } catch (error) {
    console.error('共有ドライブ一覧取得エラー:', error);
    console.log('共有ドライブの取得をスキップします');
    // エラーでも続行（空配列を返す）
    return [];
  }

  // スプレッドシートに保存
  saveSharedDrivesToSheet(sharedDrives);

  return sharedDrives;
}

function fetchAllFolders(sharedDrives, config) {
  console.log('全フォルダを取得中...');

  const allFolders = [];
  const maxDepth = parseInt(config.maxFolderDepth) || 10;
  const excludeRegex = config.excludePathRegex ? new RegExp(config.excludePathRegex) : null;

  // 共有ドライブのフォルダを取得
  if (config.driveTarget === 'SHARED_DRIVES' || config.driveTarget === 'BOTH') {
    for (const drive of sharedDrives) {
      console.log(`共有ドライブ "${drive.name}" の全階層フォルダを取得中...`);
      try {
        const driveFolders = fetchAllFoldersRecursive(drive.id, drive.name, 'teamDrive', maxDepth, excludeRegex);
        allFolders.push(...driveFolders);
      } catch (error) {
        console.error(`共有ドライブ "${drive.name}" のフォルダ取得エラー:`, error);
      }

      // 実行時間制限チェック
      if (isExecutionTimeLimitNear()) {
        console.log('実行時間制限が近づいたため、処理を中断');
        saveCheckpoint({ lastProcessedDrive: drive.id });
        break;
      }
    }
  }

  // マイドライブのフォルダを取得（共有ドライブが取得できない場合も実行）
  if (config.driveTarget === 'MY_DRIVE' || config.driveTarget === 'BOTH' || sharedDrives.length === 0) {
    console.log('マイドライブの全階層フォルダを取得中...');
    try {
      const myDriveFolders = fetchAllFoldersRecursive('root', 'マイドライブ', 'default', maxDepth, excludeRegex);
      allFolders.push(...myDriveFolders);
    } catch (error) {
      console.error('マイドライブのフォルダ取得エラー:', error);
    }
  }

  // スプレッドシートに保存
  saveFoldersToSheet(allFolders);

  return allFolders;
}

// 新しい実装：すべての階層のフォルダを効率的に取得
function fetchAllFoldersRecursive(rootId, rootName, corporaType, maxDepth, excludeRegex) {
  const allFolders = [];
  const folderMap = new Map(); // ID -> フォルダ情報のマップ
  let totalFolders = 0;

  // まず、すべてのフォルダを一括取得（階層関係なし）
  console.log(`${rootName} からすべてのフォルダを検索中...`);

  let pageToken = null;
  const query = "mimeType='application/vnd.google-apps.folder' and trashed=false";

  do {
    try {
      const options = {
        q: query,
        maxResults: 1000, // 一度により多く取得
        fields: 'items(id,title,parents,createdDate,modifiedDate),nextPageToken'
      };

      if (pageToken) {
        options.pageToken = pageToken;
      }

      // corporaTypeに応じて設定
      if (corporaType === 'teamDrive') {
        options.corpora = 'teamDrive';
        options.teamDriveId = rootId;
        options.supportsTeamDrives = true;
        options.includeTeamDriveItems = true;
      } else if (corporaType === 'default') {
        // マイドライブの場合
        options.corpora = 'default';
      }

      const response = Drive.Files.list(options);

      if (response.items && response.items.length > 0) {
        response.items.forEach(file => {
          const folderInfo = {
            id: file.id,
            name: file.title,
            parentId: file.parents && file.parents.length > 0 ? file.parents[0].id : null,
            driveId: rootId,
            driveName: rootName,
            createdTime: file.createdDate,
            modifiedTime: file.modifiedDate,
            fullPath: null,
            depth: null
          };
          folderMap.set(file.id, folderInfo);
          totalFolders++;
        });
      }

      pageToken = response.nextPageToken;

    } catch (error) {
      console.error(`フォルダ一覧取得エラー (${rootName}):`, error);
      throw error;
    }
  } while (pageToken);

  console.log(`${rootName} で ${totalFolders} 個のフォルダを検出`);

  // フルパスと階層深度を構築
  if (totalFolders > 0) {
    buildAllFolderPaths(folderMap, rootId, rootName, maxDepth);

    // 結果を配列に変換
    for (const [id, folder] of folderMap) {
      if (folder.fullPath) {
        // 除外パスチェック
        if (!excludeRegex || !excludeRegex.test(folder.fullPath)) {
          // 最大深度チェック
          if (!maxDepth || folder.depth <= maxDepth) {
            allFolders.push(folder);
          }
        }
      }
    }
  }

  console.log(`${rootName} で ${allFolders.length} 個の有効なフォルダを処理`);
  return allFolders;
}

// 旧関数は削除または名前変更
function fetchFoldersFromDrive_OLD(driveId, driveName, maxDepth, excludeRegex) {
  console.log(`ドライブ "${driveName}" のフォルダを取得中...`);

  const folders = [];
  const folderMap = {}; // ID -> フォルダ情報のマッピング
  let pageToken = null;

  // クエリ構築
  const query = "mimeType='application/vnd.google-apps.folder' and trashed=false";

  do {
    try {
      const options = {
        q: query,
        maxResults: 100
      };

      if (pageToken) {
        options.pageToken = pageToken;
      }

      // 共有ドライブの場合
      if (driveId !== 'root') {
        options.corpora = 'teamDrive';
        options.teamDriveId = driveId;
        options.supportsTeamDrives = true;
        options.includeTeamDriveItems = true;
      } else {
        // マイドライブの場合（明示的に設定）
        options.corpora = 'default';
      }

      const response = Drive.Files.list(options);

      if (response.items && response.items.length > 0) {
        for (const file of response.items) {
          folderMap[file.id] = {
            id: file.id,
            name: file.title, // v2 APIではnameではなくtitle
            parentId: file.parents && file.parents.length > 0 ? file.parents[0].id : null,
            driveId: driveId,
            driveName: driveName,
            createdTime: file.createdDate,
            modifiedTime: file.modifiedDate,
            fullPath: null, // 後で構築
            depth: null     // 後で計算
          };
        }
      }

      pageToken = response.nextPageToken;

    } catch (error) {
      console.error(`フォルダ取得エラー (${driveName}):`, error);
      throw error;
    }
  } while (pageToken);

  // フルパスと階層深度を構築
  buildFolderPaths(folderMap, driveId, driveName, maxDepth, excludeRegex);

  // マップから配列に変換
  for (const folderId in folderMap) {
    const folder = folderMap[folderId];
    if (folder.fullPath && folder.depth <= maxDepth) {
      // 除外パスチェック
      if (!excludeRegex || !excludeRegex.test(folder.fullPath)) {
        folders.push(folder);
      }
    }
  }

  return folders;
}

// 新しい実装：Mapベースのパス構築
function buildAllFolderPaths(folderMap, rootId, rootName, maxDepth) {
  // ルートレベルのフォルダから開始
  for (const [folderId, folder] of folderMap) {
    if (!folder.fullPath) {
      buildFolderPathRecursive(folder, folderMap, rootId, rootName, [], maxDepth);
    }
  }
}

// 再帰的にパスを構築（循環参照対策付き）
function buildFolderPathRecursive(folder, folderMap, rootId, rootName, visitedIds, maxDepth) {
  // 既にパスが構築済み
  if (folder.fullPath) {
    return folder.fullPath;
  }

  // 循環参照チェック
  if (visitedIds.includes(folder.id)) {
    folder.fullPath = `/${rootName}/[循環参照]/${folder.name}`;
    folder.depth = 999; // 無効な深さ
    return folder.fullPath;
  }

  // ルートレベルのフォルダ（親がないか、親がルート）
  if (!folder.parentId || folder.parentId === rootId) {
    folder.fullPath = `/${rootName}/${folder.name}`;
    folder.depth = 1;
    return folder.fullPath;
  }

  // 親フォルダを探す
  const parentFolder = folderMap.get(folder.parentId);
  if (!parentFolder) {
    // 親フォルダが見つからない（権限不足など）
    folder.fullPath = `/${rootName}/[親フォルダ不明]/${folder.name}`;
    folder.depth = 2;
    return folder.fullPath;
  }

  // 親フォルダのパスを再帰的に構築
  const newVisitedIds = [...visitedIds, folder.id];
  const parentPath = buildFolderPathRecursive(parentFolder, folderMap, rootId, rootName, newVisitedIds, maxDepth);

  if (parentPath) {
    folder.fullPath = `${parentPath}/${folder.name}`;
    folder.depth = parentFolder.depth + 1;
    return folder.fullPath;
  }

  return null;
}

// 旧関数も互換性のため残す
function buildFolderPaths(folderMap, rootId, rootName, maxDepth) {
  // 各フォルダのフルパスを再帰的に構築
  for (const folderId in folderMap) {
    if (!folderMap[folderId].fullPath) {
      buildSingleFolderPath(folderMap[folderId], folderMap, rootId, rootName, 0, maxDepth);
    }
  }
}

function buildSingleFolderPath(folder, folderMap, rootId, rootName, currentDepth, maxDepth) {
  if (currentDepth > maxDepth) {
    return null;
  }

  // 既にパスが構築されている場合
  if (folder.fullPath) {
    return folder.fullPath;
  }

  // ルートレベルのフォルダ
  if (!folder.parentId || folder.parentId === rootId) {
    folder.fullPath = `/${rootName}/${folder.name}`;
    folder.depth = 1;
    return folder.fullPath;
  }

  // 親フォルダのパスを取得
  const parentFolder = folderMap[folder.parentId];
  if (!parentFolder) {
    // 親フォルダが見つからない場合（アクセス権限など）
    folder.fullPath = `/${rootName}/[Unknown]/${folder.name}`;
    folder.depth = 2;
    return folder.fullPath;
  }

  // 親フォルダのパスを再帰的に構築
  const parentPath = buildSingleFolderPath(parentFolder, folderMap, rootId, rootName, currentDepth + 1, maxDepth);

  if (parentPath) {
    folder.fullPath = `${parentPath}/${folder.name}`;
    folder.depth = parentFolder.depth + 1;
    return folder.fullPath;
  }

  return null;
}

function saveSharedDrivesToSheet(sharedDrives) {
  const spreadsheet = SpreadsheetApp.openById(getSpreadsheetId());
  const sheet = spreadsheet.getSheetByName('Drive_SharedDrives');

  if (!sheet) {
    throw new Error('Drive_SharedDrives シートが見つかりません');
  }

  // 既存データをクリア
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clear();
  }

  // データ整形
  const now = new Date();
  const data = sharedDrives.map(drive => [
    drive.id,
    drive.name,
    drive.createdDate ? new Date(drive.createdDate) : '',
    now,
    '', // Violation
    '', // ViolationType
    '', // ViolationMessage
    ''  // MatchedRule
  ]);

  if (data.length > 0) {
    sheet.getRange(2, 1, data.length, data[0].length).setValues(data);
  }

  // 違反判定を実行
  validateSharedDrives();
}

function saveFoldersToSheet(folders) {
  const spreadsheet = SpreadsheetApp.openById(getSpreadsheetId());
  const sheet = spreadsheet.getSheetByName('Drive_Folders');

  if (!sheet) {
    throw new Error('Drive_Folders シートが見つかりません');
  }

  // 既存データをクリア
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clear();
  }

  // データ整形
  const now = new Date();
  const data = folders.map(folder => [
    folder.id,
    folder.name,
    folder.parentId || '',
    folder.driveId,
    folder.fullPath,
    folder.depth,
    folder.createdTime ? new Date(folder.createdTime) : '',
    folder.modifiedTime ? new Date(folder.modifiedTime) : '',
    now,
    '', // Violation
    '', // ViolationType
    '', // ViolationMessage
    ''  // MatchedRule
  ]);

  if (data.length > 0) {
    sheet.getRange(2, 1, data.length, data[0].length).setValues(data);
  }

  // 違反判定を実行
  validateDriveFolders();
}

function validateSharedDrives() {
  const spreadsheet = SpreadsheetApp.openById(getSpreadsheetId());
  const drivesSheet = spreadsheet.getSheetByName('Drive_SharedDrives');
  const rulesSheet = spreadsheet.getSheetByName('Rules');
  const whitelistSheet = spreadsheet.getSheetByName('Whitelist');

  if (!drivesSheet || !rulesSheet || !whitelistSheet) {
    throw new Error('必要なシートが見つかりません');
  }

  // ルールとホワイトリストを取得
  const rules = getDriveRules(rulesSheet);
  const whitelist = getDriveWhitelist(whitelistSheet);

  // ドライブデータを取得
  const lastRow = drivesSheet.getLastRow();
  if (lastRow <= 1) return;

  const drivesData = drivesSheet.getRange(2, 1, lastRow - 1, 8).getValues();
  const validationResults = [];

  // 各ドライブを判定
  for (let i = 0; i < drivesData.length; i++) {
    const driveName = drivesData[i][1]; // DriveName列
    const result = validateName(driveName, rules, whitelist, 'Drive');

    validationResults.push([
      result.violation ? 'TRUE' : 'FALSE',
      result.violationType || '',
      result.violationMessage || '',
      result.matchedRule || ''
    ]);
  }

  // 結果をシートに書き込み
  if (validationResults.length > 0) {
    drivesSheet.getRange(2, 5, validationResults.length, 4).setValues(validationResults);
  }
}

function validateDriveFolders() {
  const spreadsheet = SpreadsheetApp.openById(getSpreadsheetId());
  const foldersSheet = spreadsheet.getSheetByName('Drive_Folders');
  const rulesSheet = spreadsheet.getSheetByName('Rules');
  const whitelistSheet = spreadsheet.getSheetByName('Whitelist');

  if (!foldersSheet || !rulesSheet || !whitelistSheet) {
    throw new Error('必要なシートが見つかりません');
  }

  // ルールとホワイトリストを取得
  const rules = getDriveFolderRules(rulesSheet);
  const whitelist = getDriveFolderWhitelist(whitelistSheet);

  // フォルダデータを取得
  const lastRow = foldersSheet.getLastRow();
  if (lastRow <= 1) return;

  const foldersData = foldersSheet.getRange(2, 1, lastRow - 1, 13).getValues();
  const validationResults = [];

  // 各フォルダを判定
  for (let i = 0; i < foldersData.length; i++) {
    const folderName = foldersData[i][1]; // FolderName列
    const result = validateName(folderName, rules, whitelist, 'DriveFolder');

    validationResults.push([
      result.violation ? 'TRUE' : 'FALSE',
      result.violationType || '',
      result.violationMessage || '',
      result.matchedRule || ''
    ]);
  }

  // 結果をシートに書き込み
  if (validationResults.length > 0) {
    foldersSheet.getRange(2, 10, validationResults.length, 4).setValues(validationResults);
  }
}

function getDriveRules(rulesSheet) {
  const lastRow = rulesSheet.getLastRow();
  if (lastRow <= 1) return [];

  const rulesData = rulesSheet.getRange(2, 1, lastRow - 1, 7).getValues();

  return rulesData
    .filter(row => row[1] === 'Drive' && row[6] === 'TRUE')
    .map(row => ({
      ruleName: row[0],
      regex: row[2],
      severity: row[3],
      priority: row[4],
      description: row[5]
    }))
    .sort((a, b) => a.priority - b.priority);
}

function getDriveFolderRules(rulesSheet) {
  const lastRow = rulesSheet.getLastRow();
  if (lastRow <= 1) return [];

  const rulesData = rulesSheet.getRange(2, 1, lastRow - 1, 7).getValues();

  return rulesData
    .filter(row => row[1] === 'DriveFolder' && row[6] === 'TRUE')
    .map(row => ({
      ruleName: row[0],
      regex: row[2],
      severity: row[3],
      priority: row[4],
      description: row[5]
    }))
    .sort((a, b) => a.priority - b.priority);
}

function getDriveWhitelist(whitelistSheet) {
  const lastRow = whitelistSheet.getLastRow();
  if (lastRow <= 1) return [];

  const whitelistData = whitelistSheet.getRange(2, 1, lastRow - 1, 5).getValues();
  const now = new Date();

  return whitelistData
    .filter(row => {
      if (row[0] !== 'Drive') return false;
      if (row[4] && new Date(row[4]) < now) return false;
      return true;
    })
    .map(row => ({
      pattern: row[1],
      isRegex: row[2] === 'TRUE',
      reason: row[3]
    }));
}

function getDriveFolderWhitelist(whitelistSheet) {
  const lastRow = whitelistSheet.getLastRow();
  if (lastRow <= 1) return [];

  const whitelistData = whitelistSheet.getRange(2, 1, lastRow - 1, 5).getValues();
  const now = new Date();

  return whitelistData
    .filter(row => {
      if (row[0] !== 'DriveFolder') return false;
      if (row[4] && new Date(row[4]) < now) return false;
      return true;
    })
    .map(row => ({
      pattern: row[1],
      isRegex: row[2] === 'TRUE',
      reason: row[3]
    }));
}

// =====================================================================
// 命名ルール判定ロジック
// =====================================================================
function validateName(name, rules, whitelist, targetType) {
  // ホワイトリストチェック
  if (isWhitelisted(name, whitelist)) {
    return {
      violation: false,
      violationType: null,
      violationMessage: null,
      matchedRule: 'Whitelisted'
    };
  }

  // ルールチェック（優先度順）
  for (const rule of rules) {
    try {
      const regex = new RegExp(rule.regex);
      if (regex.test(name)) {
        // ルールにマッチ（正常）
        return {
          violation: false,
          violationType: null,
          violationMessage: null,
          matchedRule: rule.ruleName
        };
      }
    } catch (error) {
      console.error(`不正な正規表現: ${rule.ruleName} - ${rule.regex}`, error);
    }
  }

  // どのルールにもマッチしない（違反）
  const defaultMessage = getDefaultViolationMessage(targetType);
  return {
    violation: true,
    violationType: 'RULE_VIOLATION',
    violationMessage: defaultMessage,
    matchedRule: 'None'
  };
}

function isWhitelisted(name, whitelist) {
  for (const item of whitelist) {
    if (item.isRegex) {
      // 正規表現パターン
      try {
        const regex = new RegExp(item.pattern);
        if (regex.test(name)) {
          return true;
        }
      } catch (error) {
        console.error(`ホワイトリストの正規表現エラー: ${item.pattern}`, error);
      }
    } else {
      // 完全一致
      if (name === item.pattern) {
        return true;
      }
    }
  }
  return false;
}

function getDefaultViolationMessage(targetType) {
  const messages = {
    'Slack': 'チャンネル名が命名ルールに準拠していません',
    'Drive': '共有ドライブ名が命名ルールに準拠していません',
    'DriveFolder': 'フォルダ名が命名ルールに準拠していません'
  };
  return messages[targetType] || '命名ルール違反';
}

// =====================================================================
// 違反ログ管理
// =====================================================================
function updateViolationsLog() {
  const spreadsheet = SpreadsheetApp.openById(getSpreadsheetId());
  const violationsSheet = spreadsheet.getSheetByName('Violations_Log');

  if (!violationsSheet) {
    throw new Error('Violations_Log シートが見つかりません');
  }

  // 既存の違反ログを取得
  const existingViolations = getExistingViolations(violationsSheet);

  // 現在の違反を収集
  const currentViolations = collectCurrentViolations(spreadsheet);

  // 違反ステータスを更新
  const updatedViolations = mergeViolations(existingViolations, currentViolations);

  // シートを更新
  saveViolationsToSheet(violationsSheet, updatedViolations);

  return updatedViolations;
}

function getExistingViolations(violationsSheet) {
  const lastRow = violationsSheet.getLastRow();
  if (lastRow <= 1) return new Map();

  const data = violationsSheet.getRange(2, 1, lastRow - 1, 14).getValues();
  const violations = new Map();

  for (const row of data) {
    const key = `${row[1]}_${row[2]}`; // Type_ResourceID
    violations.set(key, {
      violationId: row[0],
      type: row[1],
      resourceId: row[2],
      resourceName: row[3],
      fullPath: row[4],
      violationType: row[5],
      violationMessage: row[6],
      matchedRule: row[7],
      severity: row[8],
      firstDetected: row[9],
      lastConfirmed: row[10],
      status: row[11],
      resolvedDate: row[12],
      notes: row[13]
    });
  }

  return violations;
}

function collectCurrentViolations(spreadsheet) {
  const violations = [];
  const now = new Date();

  console.log('=== 現在の違反を収集開始 ===');

  // Slackチャンネルの違反
  const slackSheet = spreadsheet.getSheetByName('Slack_Channels');
  if (slackSheet && slackSheet.getLastRow() > 1) {
    const slackData = slackSheet.getRange(2, 1, slackSheet.getLastRow() - 1, 11).getValues();
    console.log(`Slackチャンネル数: ${slackData.length}`);
    for (let i = 0; i < slackData.length; i++) {
      const row = slackData[i];
      // TRUEは文字列ではなくブール値として比較
      if (row[7] === true || row[7] === 'TRUE') { // Violation列
        console.log(`Slack違反検出: 行${i+2}, チャンネル名=${row[1]}`);
        violations.push({
          type: 'Slack',
          resourceId: row[0],
          resourceName: row[1],
          fullPath: row[1],
          violationType: row[8],
          violationMessage: row[9],
          matchedRule: row[10],
          severity: determineSeverity(row[8]),
          timestamp: now
        });
      }
    }
  }

  // 共有ドライブの違反
  const drivesSheet = spreadsheet.getSheetByName('Drive_SharedDrives');
  if (drivesSheet && drivesSheet.getLastRow() > 1) {
    const drivesData = drivesSheet.getRange(2, 1, drivesSheet.getLastRow() - 1, 8).getValues();
    console.log(`共有ドライブ数: ${drivesData.length}`);
    for (let i = 0; i < drivesData.length; i++) {
      const row = drivesData[i];
      // TRUEは文字列ではなくブール値として比較
      if (row[4] === true || row[4] === 'TRUE') { // Violation列
        console.log(`Drive違反検出: 行${i+2}, ドライブ名=${row[1]}`);
        violations.push({
          type: 'Drive',
          resourceId: row[0],
          resourceName: row[1],
          fullPath: row[1],
          violationType: row[5],
          violationMessage: row[6],
          matchedRule: row[7],
          severity: determineSeverity(row[5]),
          timestamp: now
        });
      }
    }
  }

  // フォルダの違反
  const foldersSheet = spreadsheet.getSheetByName('Drive_Folders');
  if (foldersSheet && foldersSheet.getLastRow() > 1) {
    const foldersData = foldersSheet.getRange(2, 1, foldersSheet.getLastRow() - 1, 13).getValues();
    console.log(`フォルダ数: ${foldersData.length}`);
    for (let i = 0; i < foldersData.length; i++) {
      const row = foldersData[i];
      // TRUEは文字列ではなくブール値として比較
      if (row[9] === true || row[9] === 'TRUE') { // Violation列
        console.log(`Folder違反検出: 行${i+2}, フォルダ名=${row[1]}, パス=${row[4]}`);
        violations.push({
          type: 'DriveFolder',
          resourceId: row[0],
          resourceName: row[1],
          fullPath: row[4],
          violationType: row[10],
          violationMessage: row[11],
          matchedRule: row[12],
          severity: determineSeverity(row[10]),
          timestamp: now
        });
      }
    }
  }

  console.log(`=== 違反収集完了: 総数=${violations.length} ===`);
  return violations;
}

function mergeViolations(existingViolations, currentViolations) {
  const mergedViolations = [];
  const currentViolationKeys = new Set();
  const now = new Date();

  // 現在の違反を処理
  for (const violation of currentViolations) {
    const key = `${violation.type}_${violation.resourceId}`;
    currentViolationKeys.add(key);

    if (existingViolations.has(key)) {
      // 既存の違反（ONGOING）
      const existing = existingViolations.get(key);
      mergedViolations.push({
        ...existing,
        lastConfirmed: now,
        status: 'ONGOING',
        violationType: violation.violationType,
        violationMessage: violation.violationMessage,
        severity: violation.severity
      });
    } else {
      // 新規違反（NEW）
      mergedViolations.push({
        violationId: generateViolationId(),
        ...violation,
        firstDetected: now,
        lastConfirmed: now,
        status: 'NEW',
        resolvedDate: '',
        notes: ''
      });
    }
  }

  // 解決された違反を処理
  for (const [key, existing] of existingViolations) {
    if (!currentViolationKeys.has(key) && existing.status !== 'RESOLVED') {
      mergedViolations.push({
        ...existing,
        status: 'RESOLVED',
        resolvedDate: now,
        notes: '自動解決確認'
      });
    } else if (existing.status === 'RESOLVED') {
      // 解決済みの違反はそのまま保持
      mergedViolations.push(existing);
    }
  }

  return mergedViolations;
}

function saveViolationsToSheet(violationsSheet, violations) {
  // 既存データをクリア
  if (violationsSheet.getLastRow() > 1) {
    violationsSheet.getRange(2, 1, violationsSheet.getLastRow() - 1, violationsSheet.getLastColumn()).clear();
  }

  if (violations.length === 0) return;

  // データを整形
  const data = violations.map(v => [
    v.violationId,
    v.type,
    v.resourceId,
    v.resourceName,
    v.fullPath,
    v.violationType,
    v.violationMessage,
    v.matchedRule,
    v.severity,
    v.firstDetected,
    v.lastConfirmed,
    v.status,
    v.resolvedDate || '',
    v.notes || ''
  ]);

  // シートに書き込み
  violationsSheet.getRange(2, 1, data.length, data[0].length).setValues(data);

  // ステータスで色分け
  for (let i = 0; i < data.length; i++) {
    const statusCell = violationsSheet.getRange(i + 2, 12);
    const status = data[i][11];

    if (status === 'NEW') {
      statusCell.setBackground('#FFE0E0'); // 赤系
    } else if (status === 'ONGOING') {
      statusCell.setBackground('#FFF0E0'); // オレンジ系
    } else if (status === 'RESOLVED') {
      statusCell.setBackground('#E0FFE0'); // 緑系
    }
  }
}

function generateViolationId() {
  return `VIO-${new Date().getTime()}-${Math.random().toString(36).substr(2, 5).toUpperCase()}`;
}

function determineSeverity(violationType) {
  // ルールからSeverityを判定（将来的に拡張可能）
  return violationType === 'RULE_VIOLATION' ? 'ERROR' : 'WARN';
}

function getViolationsSummary() {
  const spreadsheet = SpreadsheetApp.openById(getSpreadsheetId());
  const violationsSheet = spreadsheet.getSheetByName('Violations_Log');

  console.log('=== getViolationsSummary デバッグ開始 ===');

  if (!violationsSheet) {
    console.log('Violations_Logシートが見つかりません');
    return {
      total: 0,
      new: 0,
      ongoing: 0,
      resolved: 0,
      byType: {},
      bySeverity: {}
    };
  }

  const lastRow = violationsSheet.getLastRow();
  console.log('Violations_Log 最終行:', lastRow);

  if (lastRow <= 1) {
    console.log('データがありません');
    return {
      total: 0,
      new: 0,
      ongoing: 0,
      resolved: 0,
      byType: {},
      bySeverity: {}
    };
  }

  const data = violationsSheet.getRange(2, 1, lastRow - 1, 14).getValues();
  console.log('取得データ行数:', data.length);

  const summary = {
    total: 0,
    new: 0,
    ongoing: 0,
    resolved: 0,
    byType: {
      Slack: 0,
      Drive: 0,
      DriveFolder: 0
    },
    bySeverity: {
      ERROR: 0,
      WARN: 0
    }
  };

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const status = row[11]; // L列: Status
    const type = row[1];    // B列: Type
    const severity = row[8]; // I列: Severity

    // 空行をスキップ
    if (!row[0] && !row[1] && !row[11]) {
      console.log(`行${i + 2}: 空行をスキップ`);
      continue;
    }

    console.log(`行${i + 2}: Status=${status}, Type=${type}, Severity=${severity}`);

    // NEWまたはONGOINGの違反をカウント
    if (status === 'NEW' || status === 'ONGOING') {
      summary.total++;

      if (status === 'NEW') {
        summary.new++;
      } else if (status === 'ONGOING') {
        summary.ongoing++;
      }

      // タイプ別カウント
      if (type === 'Slack' || type === 'Drive' || type === 'DriveFolder') {
        summary.byType[type]++;
      }

      // 重要度別カウント
      if (severity === 'ERROR' || severity === 'WARN') {
        summary.bySeverity[severity]++;
      }
    } else if (status === 'RESOLVED') {
      summary.resolved++;
    }
  }

  console.log('=== サマリー結果 ===');
  console.log('総違反数:', summary.total);
  console.log('新規:', summary.new);
  console.log('継続:', summary.ongoing);
  console.log('解決済み:', summary.resolved);
  console.log('タイプ別:', JSON.stringify(summary.byType));
  console.log('重要度別:', JSON.stringify(summary.bySeverity));
  console.log('=== getViolationsSummary デバッグ終了 ===');

  return summary;
}

// =====================================================================
// レポート生成とメール送信
// =====================================================================
function generateAndSendReport() {
  console.log('レポート生成を開始...');

  try {
    // 違反情報の集計
    const summary = getViolationsSummary();

    // 詳細データの取得
    const violations = getDetailedViolations();

    // HTMLレポートの作成
    const htmlReport = createHtmlReport(summary, violations);

    // メール送信
    sendReportEmail(htmlReport, summary);

    // 最終実行情報を更新
    updateLastRunInfo(summary);

    console.log('レポート生成完了');

    return true;

  } catch (error) {
    console.error('レポート生成エラー:', error);
    throw error;
  }
}

function getDetailedViolations() {
  const spreadsheet = SpreadsheetApp.openById(getSpreadsheetId());
  const violationsSheet = spreadsheet.getSheetByName('Violations_Log');

  if (!violationsSheet || violationsSheet.getLastRow() <= 1) {
    return [];
  }

  const data = violationsSheet.getRange(2, 1, violationsSheet.getLastRow() - 1, 14).getValues();
  const violations = [];

  for (const row of data) {
    // アクティブな違反のみ（NEWまたはONGOING）
    if (row[11] === 'NEW' || row[11] === 'ONGOING') {
      violations.push({
        violationId: row[0],
        type: row[1],
        resourceId: row[2],
        resourceName: row[3],
        fullPath: row[4],
        violationType: row[5],
        violationMessage: row[6],
        matchedRule: row[7],
        severity: row[8],
        firstDetected: row[9],
        lastConfirmed: row[10],
        status: row[11],
        daysSinceDetection: getDaysSince(row[9])
      });
    }
  }

  // ステータスと重要度でソート
  violations.sort((a, b) => {
    const statusOrder = { 'NEW': 0, 'ONGOING': 1 };
    const severityOrder = { 'ERROR': 0, 'WARN': 1 };

    if (a.status !== b.status) {
      return statusOrder[a.status] - statusOrder[b.status];
    }
    if (a.severity !== b.severity) {
      return severityOrder[a.severity] - severityOrder[b.severity];
    }
    return b.daysSinceDetection - a.daysSinceDetection;
  });

  return violations;
}

function createHtmlReport(summary, violations) {
  const today = new Date().toLocaleDateString('ja-JP');
  const spreadsheetUrl = SpreadsheetApp.openById(getSpreadsheetId()).getUrl();

  let html = `
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <style>
    body {
      font-family: Arial, sans-serif;
      line-height: 1.6;
      color: #333;
      max-width: 900px;
      margin: 0 auto;
      padding: 20px;
    }
    h1 {
      color: #2c3e50;
      border-bottom: 3px solid #3498db;
      padding-bottom: 10px;
    }
    h2 {
      color: #34495e;
      margin-top: 30px;
      border-left: 4px solid #3498db;
      padding-left: 10px;
    }
    .summary-grid {
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
      gap: 15px;
      margin: 20px 0;
    }
    .summary-card {
      background: #f8f9fa;
      border-radius: 8px;
      padding: 15px;
      border-left: 4px solid #3498db;
    }
    .summary-card.error {
      border-left-color: #e74c3c;
      background: #fff5f5;
    }
    .summary-card.warning {
      border-left-color: #f39c12;
      background: #fffaf0;
    }
    .summary-card.new {
      border-left-color: #e74c3c;
      background: #fff5f5;
    }
    .metric-value {
      font-size: 24px;
      font-weight: bold;
      color: #2c3e50;
    }
    .metric-label {
      font-size: 12px;
      color: #7f8c8d;
      text-transform: uppercase;
    }
    table {
      border-collapse: collapse;
      width: 100%;
      margin-top: 20px;
      box-shadow: 0 2px 4px rgba(0,0,0,0.1);
    }
    th {
      background-color: #34495e;
      color: white;
      padding: 12px;
      text-align: left;
      font-size: 14px;
    }
    td {
      padding: 10px;
      border-bottom: 1px solid #ecf0f1;
      font-size: 13px;
    }
    tr:hover {
      background-color: #f5f6fa;
    }
    .status-new {
      background-color: #ff6b6b;
      color: white;
      padding: 3px 8px;
      border-radius: 12px;
      font-size: 11px;
      font-weight: bold;
    }
    .status-ongoing {
      background-color: #feca57;
      color: #333;
      padding: 3px 8px;
      border-radius: 12px;
      font-size: 11px;
      font-weight: bold;
    }
    .severity-error {
      color: #e74c3c;
      font-weight: bold;
    }
    .severity-warn {
      color: #f39c12;
      font-weight: bold;
    }
    .footer {
      margin-top: 40px;
      padding-top: 20px;
      border-top: 1px solid #ecf0f1;
      font-size: 12px;
      color: #7f8c8d;
    }
    .cta-button {
      display: inline-block;
      background: #3498db;
      color: white;
      padding: 10px 20px;
      text-decoration: none;
      border-radius: 5px;
      margin-top: 20px;
    }
    .cta-button:hover {
      background: #2980b9;
    }
  </style>
</head>
<body>
  <h1>📋 命名ルール監査レポート</h1>
  <p>実行日: ${today}</p>

  <h2>📊 サマリー</h2>
  <div class="summary-grid">
    <div class="summary-card">
      <div class="metric-label">総違反数</div>
      <div class="metric-value">${summary.total}</div>
    </div>
    <div class="summary-card new">
      <div class="metric-label">新規違反</div>
      <div class="metric-value">${summary.new}</div>
    </div>
    <div class="summary-card warning">
      <div class="metric-label">継続違反</div>
      <div class="metric-value">${summary.ongoing}</div>
    </div>
    <div class="summary-card">
      <div class="metric-label">解決済み</div>
      <div class="metric-value">${summary.resolved}</div>
    </div>
  </div>

  <h3>📊 タイプ別違反</h3>
  <div class="summary-grid">
    <div class="summary-card">
      <div class="metric-label">Slack</div>
      <div class="metric-value">${summary.byType.Slack || 0}</div>
    </div>
    <div class="summary-card">
      <div class="metric-label">Drive</div>
      <div class="metric-value">${summary.byType.Drive || 0}</div>
    </div>
    <div class="summary-card">
      <div class="metric-label">Folder</div>
      <div class="metric-value">${summary.byType.DriveFolder || 0}</div>
    </div>
  </div>

  <h3>⚠️ 重要度別</h3>
  <div class="summary-grid">
    <div class="summary-card error">
      <div class="metric-label">ERROR</div>
      <div class="metric-value">${summary.bySeverity.ERROR || 0}</div>
    </div>
    <div class="summary-card warning">
      <div class="metric-label">WARN</div>
      <div class="metric-value">${summary.bySeverity.WARN || 0}</div>
    </div>
  </div>
`;

  // 違反詳細テーブル
  if (violations.length > 0) {
    html += `
  <h2>📝 違反詳細</h2>
  <table>
    <thead>
      <tr>
        <th>ステータス</th>
        <th>タイプ</th>
        <th>名称</th>
        <th>フルパス</th>
        <th>重要度</th>
        <th>違反内容</th>
        <th>初回検出</th>
        <th>経過日数</th>
      </tr>
    </thead>
    <tbody>
`;

    // 最大100件まで表示
    const displayViolations = violations.slice(0, 100);

    for (const violation of displayViolations) {
      const statusClass = violation.status.toLowerCase();
      const severityClass = violation.severity.toLowerCase();
      const firstDetectedDate = new Date(violation.firstDetected).toLocaleDateString('ja-JP');

      html += `
      <tr>
        <td><span class="status-${statusClass}">${violation.status}</span></td>
        <td>${violation.type}</td>
        <td><strong>${escapeHtml(violation.resourceName)}</strong></td>
        <td>${escapeHtml(violation.fullPath)}</td>
        <td><span class="severity-${severityClass}">${violation.severity}</span></td>
        <td>${escapeHtml(violation.violationMessage)}</td>
        <td>${firstDetectedDate}</td>
        <td>${violation.daysSinceDetection}日</td>
      </tr>
`;
    }

    html += `
    </tbody>
  </table>
`;

    if (violations.length > 100) {
      html += `
  <p><em>※ 表示されているのは先頭100件です。全件を確認するにはスプレッドシートをご覧ください。</em></p>
`;
    }
  } else {
    html += `
  <h2>✅ 違反なし</h2>
  <p>現在、アクティブな命名ルール違反はありません。</p>
`;
  }

  // フッター
  html += `
  <div class="footer">
    <a href="${spreadsheetUrl}" class="cta-button">📋 詳細をスプレッドシートで確認</a>
    <p>
      このレポートは自動生成されました。<br>
      問い合わせ: IT管理部門<br>
      次回実行: 翌営業日 8:30
    </p>
  </div>
</body>
</html>
`;

  return html;
}

function sendReportEmail(htmlReport, summary) {
  const config = getConfig();
  const recipients = config.notificationEmail;

  if (!recipients || recipients.trim() === '') {
    console.log('通知先メールが設定されていないため、メール送信をスキップ');
    return;
  }

  const today = new Date().toLocaleDateString('ja-JP');
  const subject = `[命名監査] ${today} Slack/Drive/Folder 違反 ${summary.total}件`;

  // テキスト版も作成
  const textBody = createTextReport(summary);

  // メール送信
  MailApp.sendEmail({
    to: recipients,
    subject: subject,
    body: textBody,
    htmlBody: htmlReport,
    name: 'Naming Audit System'
  });

  console.log(`レポートを送信しました: ${recipients}`);
}

function createTextReport(summary) {
  const today = new Date().toLocaleDateString('ja-JP');
  const spreadsheetUrl = SpreadsheetApp.openById(getSpreadsheetId()).getUrl();

  let text = `
命名ルール監査レポート
========================
実行日: ${today}

[サマリー]
- 総違反数: ${summary.total}件
- 新規違反: ${summary.new}件
- 継続違反: ${summary.ongoing}件
- 解決済み: ${summary.resolved}件

[タイプ別]
- Slack: ${summary.byType.Slack || 0}件
- Drive: ${summary.byType.Drive || 0}件
- Folder: ${summary.byType.DriveFolder || 0}件

[重要度別]
- ERROR: ${summary.bySeverity.ERROR || 0}件
- WARN: ${summary.bySeverity.WARN || 0}件

詳細は以下のスプレッドシートをご確認ください:
${spreadsheetUrl}

---
このレポートは自動生成されました。
`;

  return text;
}

function updateLastRunInfo(summary) {
  const spreadsheet = SpreadsheetApp.openById(getSpreadsheetId());
  const sheet = spreadsheet.getSheetByName('Report_LastRun');

  if (!sheet) {
    console.error('Report_LastRun シートが見つかりません');
    return;
  }

  const now = new Date();
  const data = [
    ['最終実行日時', now],
    ['実行ステータス', 'SUCCESS'],
    ['Slackチャンネル取得数', getSheetRowCount('Slack_Channels')],
    ['Drive取得数', getSheetRowCount('Drive_SharedDrives')],
    ['フォルダ取得数', getSheetRowCount('Drive_Folders')],
    ['新規違反検出数', summary.new],
    ['継続違反数', summary.ongoing],
    ['解決済み違反数', summary.resolved],
    ['エラー内容', ''],
    ['実行時間(秒)', ''] // 実行時間はメイン関数で計測
  ];

  sheet.getRange(2, 1, data.length, 2).setValues(data);
}

// =====================================================================
// ユーティリティ関数
// =====================================================================
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

function getDaysSince(date) {
  if (!date) return 0;
  const now = new Date();
  const then = new Date(date);
  const diffTime = Math.abs(now - then);
  const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
  return diffDays;
}

function escapeHtml(text) {
  const map = {
    '&': '&amp;',
    '<': '&lt;',
    '>': '&gt;',
    '"': '&quot;',
    "'": '&#039;'
  };
  return String(text).replace(/[&<>"']/g, m => map[m]);
}

function getSheetRowCount(sheetName) {
  try {
    const spreadsheet = SpreadsheetApp.openById(getSpreadsheetId());
    const sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) return 0;
    const lastRow = sheet.getLastRow();
    return lastRow > 1 ? lastRow - 1 : 0; // ヘッダーを除く
  } catch (error) {
    return 0;
  }
}

// =====================================================================
// トリガー設定と手動実行用関数
// =====================================================================
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

// この関数は削除（新規スプレッドシート作成しないため）
// function openSpreadsheet() { ... }

// =====================================================================
// スプレッドシートメニュー
// =====================================================================
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('⚙️ 命名監査システム')
    .addItem('🚀 初期セットアップ実行', 'setupSpreadsheetAndSystem')
    .addSeparator()
    .addItem('🔑 Slack Tokenテスト', 'testSlackToken')
    .addItem('🔄 手動実行 (全体)', 'mainAudit')
    .addItem('💬 Slackのみテスト', 'testSlackOnly')
    .addItem('📁 Driveのみテスト', 'testDriveOnly')
    .addItem('📧 レポートのみテスト', 'testReportOnly')
    .addSeparator()
    .addItem('⏰ トリガー設定', 'setupTriggers')
    .addToUi();
}

// Slack Token テスト関数
function testSlackToken() {
  const token = PropertiesService.getScriptProperties().getProperty('SLACK_BOT_TOKEN');

  if (!token) {
    SpreadsheetApp.getUi().alert('エラー', 'Slack Bot Tokenが設定されていません。\n\nプロジェクトの設定 → スクリプト プロパティで設定してください。', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  try {
    // テスト用APIコール
    const url = 'https://slack.com/api/auth.test';
    const options = {
      'method': 'get',
      'headers': {
        'Authorization': `Bearer ${token}`
      },
      'muteHttpExceptions': true
    };

    const response = UrlFetchApp.fetch(url, options);
    const data = JSON.parse(response.getContentText());

    if (data.ok) {
      SpreadsheetApp.getUi().alert('成功', `Slack接続成功！\n\nチーム: ${data.team}\nユーザー: ${data.user}`, SpreadsheetApp.getUi().ButtonSet.OK);
    } else {
      SpreadsheetApp.getUi().alert('エラー', `Slack接続失敗: ${data.error}`, SpreadsheetApp.getUi().ButtonSet.OK);
    }
  } catch (error) {
    SpreadsheetApp.getUi().alert('エラー', `エラーが発生しました: ${error.toString()}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}