/**
 * Slack API 連携機能
 * Slackチャンネル情報の取得とスプレッドシートへの保存
 */

// =================================================================
// Slack API メイン関数
// =================================================================
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

// =================================================================
// Slack API 呼び出し
// =================================================================
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

// =================================================================
// Slack API ユーティリティ
// =================================================================
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

// =================================================================
// スプレッドシートへの保存
// =================================================================
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

// =================================================================
// チャンネル名違反判定
// =================================================================
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

// =================================================================
// Slack用ルール取得
// =================================================================
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

// =================================================================
// Slack用ホワイトリスト取得
// =================================================================
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