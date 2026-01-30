/**
 * ========================================
 * Slack ログ収集 最適化版 - 大量データ対応
 * ========================================
 */

// 既存の設定値をそのまま使用
const SLACK_BOT_TOKEN = 'YOUR_SLACK_BOT_TOKEN';
const GOOGLE_DOC_ID = '1dkxrY8mtC28bWyDtxm0NVDohlESzNwqHJqq4PQFimqY';
const LOG_SHEET_NAME = 'slack_log';
const LAST_TS_SHEET_NAME = 'slack_channel_last_ts';
const MANUAL_SHEET_NAME = 'business_manual';
const FAQ_SHEET_NAME = 'faq_list';
const OPENAI_API_KEY = 'YOUR_OPENAI_API_KEY';
const NOTIFICATION_EMAIL = 'your-email@example.com';

// パフォーマンス設定
const BATCH_SIZE = 500; // 一度に処理するメッセージ数
const MAX_EXECUTION_TIME = 5 * 60 * 1000; // 5分（GASの制限は6分だが余裕を持つ）
const CACHE_EXPIRATION = 60 * 60; // キャッシュ有効期限（秒）

/**
 * 実行時間管理クラス
 */
class ExecutionTimer {
  constructor(maxTime = MAX_EXECUTION_TIME) {
    this.startTime = Date.now();
    this.maxTime = maxTime;
  }
  
  isTimeUp() {
    return Date.now() - this.startTime > this.maxTime;
  }
  
  getElapsedTime() {
    return Date.now() - this.startTime;
  }
  
  getRemainingTime() {
    return Math.max(0, this.maxTime - this.getElapsedTime());
  }
}

/**
 * バッチ処理対応のデータ書き込みクラス
 */
class BatchWriter {
  constructor(sheet, batchSize = BATCH_SIZE) {
    this.sheet = sheet;
    this.batchSize = batchSize;
    this.buffer = [];
  }
  
  add(row) {
    this.buffer.push(row);
    if (this.buffer.length >= this.batchSize) {
      this.flush();
    }
  }
  
  flush() {
    if (this.buffer.length === 0) return;
    
    const lastRow = this.sheet.getLastRow();
    const range = this.sheet.getRange(lastRow + 1, 1, this.buffer.length, this.buffer[0].length);
    range.setValues(this.buffer);
    
    console.log(`バッチ書き込み: ${this.buffer.length}行を追加`);
    this.buffer = [];
  }
}

/**
 * 処理状態を保存・復元するクラス
 */
class ProcessState {
  constructor() {
    this.scriptProperties = PropertiesService.getScriptProperties();
  }
  
  save(state) {
    this.scriptProperties.setProperty('processing_state', JSON.stringify(state));
  }
  
  load() {
    const stateStr = this.scriptProperties.getProperty('processing_state');
    return stateStr ? JSON.parse(stateStr) : null;
  }
  
  clear() {
    this.scriptProperties.deleteProperty('processing_state');
  }
  
  isProcessing() {
    return this.load() !== null;
  }
}

/**
 * キャッシュ管理クラス
 */
class CacheManager {
  constructor() {
    this.cache = CacheService.getScriptCache();
  }
  
  get(key) {
    const value = this.cache.get(key);
    return value ? JSON.parse(value) : null;
  }
  
  set(key, value, expirationInSeconds = CACHE_EXPIRATION) {
    this.cache.put(key, JSON.stringify(value), expirationInSeconds);
  }
  
  remove(key) {
    this.cache.remove(key);
  }
}

/**
 * 最適化されたメインのログ取得関数
 */
function fetchAndAppendAllChannelsOptimized() {
  const timer = new ExecutionTimer();
  const state = new ProcessState();
  const cache = new CacheManager();
  
  // 前回の処理が中断されていた場合は再開
  let processData = state.load();
  if (!processData) {
    // 新規処理開始
    const channels = getSlackChannelsWithCache(cache);
    processData = {
      channels: channels,
      currentChannelIndex: 0,
      totalProcessed: 0,
      startTime: new Date().toISOString()
    };
  }
  
  console.log(`処理開始: チャンネル ${processData.currentChannelIndex + 1}/${processData.channels.length} から`);
  
  const sheet = getOrCreateLogSheet();
  const lastTsSheet = getOrCreateLastTsSheet();
  const batchWriter = new BatchWriter(sheet);
  const userCache = {};
  
  try {
    // チャンネルごとに処理
    while (processData.currentChannelIndex < processData.channels.length) {
      if (timer.isTimeUp()) {
        // 時間切れの場合、状態を保存して終了
        batchWriter.flush();
        state.save(processData);
        console.log(`時間切れ: ${processData.totalProcessed}件処理済み。次回自動再開します。`);
        
        // 次回の実行をトリガーで予約
        scheduleNextRun();
        return;
      }
      
      const channel = processData.channels[processData.currentChannelIndex];
      const processedCount = processChannelOptimized(
        channel, 
        sheet, 
        lastTsSheet, 
        batchWriter, 
        userCache, 
        timer,
        cache
      );
      
      processData.totalProcessed += processedCount;
      processData.currentChannelIndex++;
      
      console.log(`チャンネル ${channel.name} 処理完了: ${processedCount}件`);
    }
    
    // すべて完了
    batchWriter.flush();
    state.clear();
    
    console.log(`全処理完了: 合計 ${processData.totalProcessed}件のメッセージを処理`);
    
    // サマリーを更新
    updateSummarySheet(processData);
    
  } catch (error) {
    console.error('エラーが発生しました:', error);
    batchWriter.flush();
    state.save(processData);
    throw error;
  }
}

/**
 * チャンネルの処理（最適化版）
 */
function processChannelOptimized(channel, sheet, lastTsSheet, batchWriter, userCache, timer, cache) {
  const lastFetchedTs = getLastFetchedTs(lastTsSheet, channel.id);
  let processedCount = 0;
  let hasMore = true;
  let cursor = null;
  let maxTs = lastFetchedTs;
  
  while (hasMore && !timer.isTimeUp()) {
    // ページネーション対応でメッセージを取得
    const result = getChannelMessagesPaginated(channel.id, lastFetchedTs, cursor);
    if (!result || !result.messages) break;
    
    // メッセージをバッチ処理
    const messages = result.messages.reverse();
    for (const msg of messages) {
      if (timer.isTimeUp()) break;
      
      if (!msg.text) continue;
      
      const userName = getUserNameCached(msg.user, userCache, cache);
      const fullText = processMessageText(msg);
      const date = new Date(Number(msg.ts.split('.')[0]) * 1000);
      const threadTs = msg.thread_ts || msg.ts;
      const messageId = `${channel.id}_${msg.ts}`;
      
      batchWriter.add([
        channel.id,
        channel.name,
        msg.ts,
        threadTs,
        userName,
        fullText,
        date,
        messageId
      ]);
      
      processedCount++;
      maxTs = Math.max(maxTs, msg.ts);
    }
    
    hasMore = result.has_more;
    cursor = result.response_metadata?.next_cursor;
  }
  
  // 最終タイムスタンプを更新
  if (maxTs !== lastFetchedTs) {
    updateLastFetchedTs(lastTsSheet, channel.id, maxTs);
  }
  
  return processedCount;
}

/**
 * ページネーション対応のメッセージ取得
 */
function getChannelMessagesPaginated(channelId, oldest, cursor) {
  let url = `https://slack.com/api/conversations.history?channel=${channelId}&limit=200`;
  
  if (oldest && oldest !== '0') {
    url += `&oldest=${oldest}`;
  }
  
  if (cursor) {
    url += `&cursor=${cursor}`;
  }
  
  const options = {
    method: 'get',
    headers: { 
      'Authorization': 'Bearer ' + SLACK_BOT_TOKEN,
      'Content-Type': 'application/json'
    },
    muteHttpExceptions: true
  };
  
  try {
    const response = UrlFetchApp.fetch(url, options);
    const data = JSON.parse(response.getContentText());
    
    if (!data.ok) {
      console.error(`API エラー: ${data.error}`);
      return null;
    }
    
    return data;
  } catch (error) {
    console.error(`通信エラー: ${error}`);
    return null;
  }
}

/**
 * キャッシュ対応のチャンネル一覧取得
 */
function getSlackChannelsWithCache(cache) {
  const cacheKey = 'slack_channels';
  let channels = cache.get(cacheKey);
  
  if (channels) {
    console.log('キャッシュからチャンネル一覧を取得');
    return channels;
  }
  
  channels = [];
  let cursor = null;
  
  do {
    let url = 'https://slack.com/api/conversations.list?limit=200&types=public_channel,private_channel';
    if (cursor) url += `&cursor=${cursor}`;
    
    const options = {
      method: 'get',
      headers: { 'Authorization': 'Bearer ' + SLACK_BOT_TOKEN },
      muteHttpExceptions: true
    };
    
    const response = UrlFetchApp.fetch(url, options);
    const data = JSON.parse(response.getContentText());
    
    if (data.ok && data.channels) {
      channels = channels.concat(data.channels.filter(ch => ch.is_member));
      cursor = data.response_metadata?.next_cursor;
    } else {
      break;
    }
  } while (cursor);
  
  cache.set(cacheKey, channels);
  console.log(`${channels.length}個のチャンネルを取得`);
  
  return channels;
}

/**
 * キャッシュ対応のユーザー名取得
 */
function getUserNameCached(userId, localCache, globalCache) {
  if (!userId) return 'Unknown';
  
  // ローカルキャッシュを確認
  if (localCache[userId]) {
    return localCache[userId];
  }
  
  // グローバルキャッシュを確認
  const cacheKey = `user_${userId}`;
  let userName = globalCache.get(cacheKey);
  
  if (userName) {
    localCache[userId] = userName;
    return userName;
  }
  
  // APIから取得
  const url = `https://slack.com/api/users.info?user=${userId}`;
  const options = {
    method: 'get',
    headers: { 'Authorization': 'Bearer ' + SLACK_BOT_TOKEN },
    muteHttpExceptions: true
  };
  
  try {
    const response = UrlFetchApp.fetch(url, options);
    const data = JSON.parse(response.getContentText());
    
    if (data.ok && data.user) {
      userName = data.user.real_name || data.user.name || 'Unknown';
    } else {
      userName = 'Unknown';
    }
  } catch (error) {
    userName = 'Unknown';
  }
  
  // キャッシュに保存
  localCache[userId] = userName;
  globalCache.set(cacheKey, userName);
  
  return userName;
}

/**
 * メッセージテキストの処理
 */
function processMessageText(msg) {
  let fullText = msg.text || '';
  
  // リプライがある場合は追加
  if (msg.reply_count && msg.reply_count > 0) {
    fullText += ` [${msg.reply_count}件の返信]`;
  }
  
  // 添付ファイルがある場合は追加
  if (msg.files && msg.files.length > 0) {
    const fileNames = msg.files.map(f => f.name || 'ファイル').join(', ');
    fullText += ` [添付: ${fileNames}]`;
  }
  
  return fullText;
}

/**
 * 次回実行をスケジュール
 */
function scheduleNextRun() {
  // 既存のトリガーを削除
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'fetchAndAppendAllChannelsOptimized') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  // 1分後に再実行
  ScriptApp.newTrigger('fetchAndAppendAllChannelsOptimized')
    .timeBased()
    .after(1 * 60 * 1000)
    .create();
    
  console.log('1分後に処理を再開します');
}

/**
 * 定期実行用のトリガー設定
 */
function setupOptimizedTriggers() {
  // 既存のトリガーをすべて削除
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => ScriptApp.deleteTrigger(trigger));
  
  // 1時間ごとの定期実行を設定
  ScriptApp.newTrigger('fetchAndAppendAllChannelsOptimized')
    .timeBased()
    .everyHours(1)
    .create();
    
  console.log('最適化版の定期実行トリガーを設定しました（1時間ごと）');
}

/**
 * 手動実行用：処理状態をリセット
 */
function resetProcessingState() {
  const state = new ProcessState();
  state.clear();
  
  // 残っているトリガーも削除
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'fetchAndAppendAllChannelsOptimized') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  console.log('処理状態をリセットしました');
}

/**
 * サマリーシートの更新
 */
function updateSummarySheet(processData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('processing_summary');
  
  if (!sheet) {
    sheet = ss.insertSheet('processing_summary');
    sheet.getRange(1, 1, 1, 5).setValues([['処理日時', '処理チャンネル数', '処理メッセージ数', '処理時間', 'ステータス']]);
  }
  
  const lastRow = sheet.getLastRow();
  const endTime = new Date();
  const startTime = new Date(processData.startTime);
  const processingTime = Math.round((endTime - startTime) / 1000) + '秒';
  
  sheet.getRange(lastRow + 1, 1, 1, 5).setValues([[
    endTime,
    processData.channels.length,
    processData.totalProcessed,
    processingTime,
    '完了'
  ]]);
}

// 既存の関数との互換性維持
function getOrCreateLogSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(LOG_SHEET_NAME);
  
  if (!sheet) {
    sheet = ss.insertSheet(LOG_SHEET_NAME);
    const headers = ['channel_id', 'channel_name', 'timestamp', 'thread_ts', 'user_name', 'message', 'date', 'message_id'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  
  return sheet;
}

function getOrCreateLastTsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(LAST_TS_SHEET_NAME);
  
  if (!sheet) {
    sheet = ss.insertSheet(LAST_TS_SHEET_NAME);
    const headers = ['channel_id', 'last_ts', 'last_updated'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  
  return sheet;
}

function getLastFetchedTs(sheet, channelId) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === channelId) {
      return data[i][1];
    }
  }
  return '0';
}

function updateLastFetchedTs(sheet, channelId, ts) {
  const data = sheet.getDataRange().getValues();
  const now = new Date();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === channelId) {
      sheet.getRange(i + 1, 2, 1, 2).setValues([[ts, now]]);
      return;
    }
  }
  
  // 新規追加
  const lastRow = sheet.getLastRow();
  sheet.getRange(lastRow + 1, 1, 1, 3).setValues([[channelId, ts, now]]);
}

/**
 * カスタムメニューの追加
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  ui.createMenu('📋 Slack ログツール（最適化版）')
    .addItem('▶️ ログ取得開始（大量データ対応）', 'fetchAndAppendAllChannelsOptimized')
    .addItem('⏸️ 処理状態をリセット', 'resetProcessingState')
    .addItem('⚙️ 定期実行を設定（1時間ごと）', 'setupOptimizedTriggers')
    .addItem('📊 処理状況を確認', 'showProcessingStatus')
    .addToUi();
}

/**
 * 処理状況の表示
 */
function showProcessingStatus() {
  const state = new ProcessState();
  const processData = state.load();
  
  if (processData) {
    const ui = SpreadsheetApp.getUi();
    ui.alert(
      '処理状況',
      `現在処理中です：\n` +
      `- 進捗: ${processData.currentChannelIndex}/${processData.channels.length} チャンネル\n` +
      `- 処理済みメッセージ: ${processData.totalProcessed}件\n` +
      `- 開始時刻: ${processData.startTime}`,
      ui.ButtonSet.OK
    );
  } else {
    SpreadsheetApp.getUi().alert('現在処理中のタスクはありません');
  }
}