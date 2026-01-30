/**
 * Slack AI Bot - 統合版（改良版・デバッグ機能付き）
 * 
 * このファイルは以下のモジュールを統合したものです：
 * - 設定管理とユーティリティ
 * - Slack Bot基本クラス
 * - メインエントリポイント
 * - FAQ検索とDrive検索機能
 * - 自然言語処理
 * - 文字列処理ユーティリティ
 * - ログ機能
 * - ファイル処理機能（PDF、Word、Googleドキュメント）
 * - ドキュメント編集・修正案作成機能
 * - デバッグ機能（Slack投稿問題の診断）
 */

// ===========================
// デバッグ設定
// ===========================
const DEBUG_MODE = true; // デバッグモードの有効/無効
const DEBUG_SHEET_NAME = 'debug_log'; // デバッグ用のシート名

/**
 * デバッグログを記録
 */
function debugLog(category, message, data = null) {
  if (!DEBUG_MODE) return;
  
  console.log(`[${category}] ${message}`, data);
  Logger.log(`[${category}] ${message} ${data ? JSON.stringify(data) : ''}`);
  
  // スプレッドシートにも記録
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let debugSheet;
    try {
      debugSheet = ss.getSheetByName(DEBUG_SHEET_NAME);
    } catch (e) {
      debugSheet = ss.insertSheet(DEBUG_SHEET_NAME);
      debugSheet.appendRow(['Timestamp', 'Category', 'Message', 'Data']);
    }
    debugSheet.appendRow([
      new Date(),
      category,
      message,
      data ? JSON.stringify(data) : ''
    ]);
  } catch (e) {
    Logger.log('Debug sheet error: ' + e.toString());
  }
}

// ===========================
// 設定管理とユーティリティ
// ===========================

/**
 * 環境変数取得（エラーチェック付き）
 */
function Settings() {
  try {
    const env = PropertiesService.getScriptProperties().getProperties();
    
    // 必須プロパティのチェック
    const required = ['SLACK_TOKEN', 'OPEN_AI_TOKEN'];
    const missing = required.filter(key => !env[key]);
    
    if (missing.length > 0) {
      debugLog('Settings', 'Missing required properties', missing);
      throw new Error(`Missing required properties: ${missing.join(', ')}`);
    }
    
    debugLog('Settings', 'Properties loaded successfully', Object.keys(env));
    return env;
  } catch (e) {
    debugLog('Settings', 'Error loading properties', e.toString());
    throw e;
  }
}

/**
 * ログ出力
 */
function Log(title, text) {
  Logger.log(title, text);
  debugLog('Log', title, text);
}

// ===========================
// ログシート管理
// ===========================

var SheetLog = {
  log: function(message) {
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      if (!ss) {
        debugLog('SheetLog', 'No active spreadsheet');
        return;
      }
      
      const logSheet = ss.getSheetByName('log') || ss.insertSheet('log');
      const now = new Date();    
      logSheet.appendRow([now, message]);
      debugLog('SheetLog', 'Message logged', message);
    } catch(e) {
      debugLog('SheetLog', 'Error', e.toString());
    }
  }
}

// ===========================
// 文字列処理ユーティリティ
// ===========================

function katakanaToHiragana(text) {
  return text.replace(/[\u30a1-\u30f6]/g, function(match) {
    var chr = match.charCodeAt(0) - 0x60;
    return String.fromCharCode(chr);
  });
}

function toHalfWidth(str) {
  str = str.replace(/[Ａ-Ｚａ-ｚ０-９]/g, function(s) {
    return String.fromCharCode(s.charCodeAt(0) - 0xFEE0);
  });
  return str;
}

// ===========================
// Slack Bot クラス（改良版）
// ===========================

class SlackBot {
  constructor(e) {
    debugLog('SlackBot', 'Constructor called', e);
    this.requestEvent = e;
    this.postData = null;
    this.slackEvent = null;
    this.responseData = this.init();
    this.verification();
  }

  responseJsonData(json) {
    debugLog('SlackBot', 'Response JSON', json);
    return ContentService.createTextOutput(JSON.stringify(json)).setMimeType(ContentService.MimeType.JSON);
  }

  init() {
    const e = this.requestEvent;
    debugLog('SlackBot', 'Init started', { hasPostData: !!e?.postData });
    
    if (!e?.postData) {
      const error = { error: 'postData is missing or undefined.', request: JSON.stringify(e, null, "  ") };
      debugLog('SlackBot', 'No postData', error);
      return error;
    }
    
    // postDataの内容を解析
    try {
      const contents = e.postData.contents;
      debugLog('SlackBot', 'PostData contents', contents);
      
      this.postData = JSON.parse(contents);
      debugLog('SlackBot', 'Parsed postData', this.postData);
      
      // URL verification check
      if (this.postData.type === 'url_verification') {
        debugLog('SlackBot', 'URL verification detected');
        return null;
      }
      
      // Event callback check
      if (this.postData.type === 'event_callback') {
        this.slackEvent = this.postData;
        debugLog('SlackBot', 'Event callback detected', this.slackEvent);
        return null;
      }
      
      // Unknown type
      const error = { error: 'Unknown event type', type: this.postData.type };
      debugLog('SlackBot', 'Unknown event type', error);
      return error;
      
    } catch (error) {
      const err = { error: 'Invalid JSON format in postData contents.', details: error.toString() };
      debugLog('SlackBot', 'JSON parse error', err);
      return err;
    }
  }

  verification() {
    if (!this.postData || this.responseData) return null;
    
    if (this.postData.type === 'url_verification') {
      this.responseData = { "challenge": this.postData.challenge };
      debugLog('SlackBot', 'URL verification response', this.responseData);
      return this.responseData;
    }
    return null;
  }

  hasCache(key) {
    if (!key) return true;
    const cache = CacheService.getScriptCache();
    const cached = cache.get(key);
    if (cached) {
      debugLog('SlackBot', 'Cache hit', key);
      return true;
    }
    cache.put(key, 'true', 30 * 60);
    debugLog('SlackBot', 'Cache miss, stored', key);
    return false;
  }

  handleEvent(type, callback = () => { }) {
    debugLog('SlackBot', 'HandleEvent', { type, hasEvent: !!this.slackEvent });
    
    if (!this.slackEvent || this.responseData) return null;
    
    const event = this.slackEvent?.event;
    if (!event || event.type !== type) {
      debugLog('SlackBot', 'Event type mismatch', { expected: type, actual: event?.type });
      return null;
    }
    
    const callbackResponse = callback({ event });
    if (callbackResponse) {
      this.responseData = callbackResponse;
      debugLog('SlackBot', 'Callback response set', callbackResponse);
    }
    return callbackResponse;
  }

  handleBase(type, targetType, callback = () => {}) {
    return this.handleEvent(type, ({ event }) => {
      debugLog('SlackBot', 'HandleBase', { type, targetType, event });
      
      const { text: message, channel, thread_ts: threadTs, ts, client_msg_id, bot_id, app_id } = event;
      
      if (bot_id || app_id) {
        debugLog('SlackBot', 'Bot message ignored', { bot_id, app_id });
        return null;
      }
      
      if (event.type !== targetType) {
        debugLog('SlackBot', 'Type mismatch in handleBase', { expected: targetType, actual: event.type });
        return null;
      }
      
      const cacheKey = `${channel}:${client_msg_id}`;
      if (this.hasCache(cacheKey)) {
        debugLog('SlackBot', 'Duplicate message ignored', cacheKey);
        return null;
      }
      
      return callback ? callback({ message, channel, threadTs: threadTs ?? ts, event }) : null;
    });
  }

  handleMessageEventBase(callback) { 
    return this.handleBase("message", "message", callback); 
  }
  
  handleMentionEventBase(callback) { 
    debugLog('SlackBot', 'HandleMentionEventBase called');
    return this.handleBase("app_mention", "app_mention", callback); 
  }
  
  handleReactionEventBase(callback) { 
    return this.handleBase("reaction_added", "reaction_added", callback); 
  }

  response() {
    debugLog('SlackBot', 'Final response', this.responseData);
    
    if (this.responseData) {
      return this.responseJsonData(this.responseData);
    }
    
    // 成功レスポンスを返す（Slack 3秒タイムアウト対策）
    return ContentService.createTextOutput('');
  }
}

// ===========================
// Slack API 関連機能（改良版）
// ===========================

/**
 * チャンネル情報取得
 */
function getChannelInfo(channelId) {
  debugLog('API', 'Getting channel info', channelId);
  
  const url = 'https://slack.com/api/conversations.info';
  const config = Settings();
  
  if (!config?.SLACK_TOKEN) {
    debugLog('API', 'No Slack token');
    return null;
  }
  
  const payload = {
    token: config.SLACK_TOKEN,
    channel: channelId,
  };
  
  const options = {
    method: 'post',
    payload,
    muteHttpExceptions: true
  };
  
  try {
    const response = UrlFetchApp.fetch(url, options);
    const data = JSON.parse(response.getContentText());
    
    if (!data.ok) {
      debugLog('API', 'Channel info error', data.error);
      return null;
    }
    
    debugLog('API', 'Channel info success', data.channel?.name);
    return data.channel;
  } catch (e) {
    debugLog('API', 'Channel info exception', e.toString());
    return null;
  }
}

/**
 * スレッドメッセージ取得
 */
function getThreadMessages(channelId, threadTs) {
  debugLog('API', 'Getting thread messages', { channelId, threadTs });
  
  const url = 'https://slack.com/api/conversations.replies';
  const config = Settings();
  
  if (!config?.SLACK_TOKEN) {
    debugLog('API', 'No Slack token');
    return [];
  }
  
  const payload = {
    token: config.SLACK_TOKEN,
    channel: channelId,
    ts: threadTs,
  };
  
  const options = {
    method: 'get',
    payload,
    muteHttpExceptions: true
  };
  
  try {
    const response = UrlFetchApp.fetch(url, options);
    const data = JSON.parse(response.getContentText());
    
    if (!data.ok) {
      debugLog('API', 'Thread messages error', data.error);
      return [];
    }
    
    debugLog('API', 'Thread messages success', data.messages?.length);
    return data.messages || [];
  } catch (e) {
    debugLog('API', 'Thread messages exception', e.toString());
    return [];
  }
}

/**
 * メッセージ送信（改良版）
 */
function postMessage(message, channel, threadTs = null) {
  debugLog('API', 'Posting message', { channel, threadTs, messageLength: message?.length });
  
  const url = 'https://slack.com/api/chat.postMessage';
  const config = Settings();
  
  if (!config?.SLACK_TOKEN) {
    debugLog('API', 'No Slack token for posting');
    return false;
  }
  
  const payload = {
    token: config.SLACK_TOKEN,
    channel: channel,
    text: message,
    unfurl_links: true,
    ...(threadTs ? { thread_ts: threadTs } : {}),
  };
  
  const options = {
    method: 'post',
    payload,
    muteHttpExceptions: true
  };
  
  try {
    const response = UrlFetchApp.fetch(url, options);
    const data = JSON.parse(response.getContentText());
    
    if (!data.ok) {
      debugLog('API', 'Message post error', { error: data.error, response: data });
      return false;
    }
    
    debugLog('API', 'Message posted successfully', data.ts);
    return true;
  } catch (e) {
    debugLog('API', 'Message post exception', e.toString());
    return false;
  }
}

// ===========================
// 自然言語処理 (Google Natural Language API)
// ===========================

function gNL(textdata) {
  var apiKey = PropertiesService.getScriptProperties().getProperty('GOOGLE_NL_API');
  if (!apiKey) {
    debugLog('NLP', 'No Google NL API key');
    return null;
  }
  
  var url = "https://language.googleapis.com/v1/documents:analyzeSyntax?key=" + apiKey;
  var payload = {
    document: {
      type: "PLAIN_TEXT",
      content: textdata
    },
    encodingType: "UTF8"
  };  
  var options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  
  try {
    var response = UrlFetchApp.fetch(url, options);
    return JSON.parse(response.getContentText());  
  } catch(e) {
    debugLog('NLP', 'Error', e.toString());
    return null;
  }
}

function filterGNL(gNLobj, tags) {
  if (!gNLobj) return [];
  var words = gNLobj.tokens
    .filter(token => tags.includes(token.partOfSpeech.tag))
    .map(token => token.text.content); 
  return words;
}

// ===========================
// FAQ検索機能
// ===========================

function getFaqRole(question) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
      debugLog('FAQ', 'No active spreadsheet');
      return null;
    }
    
    const faqSheet = ss.getSheetByName('faq');
    if (!faqSheet) {
      debugLog('FAQ', 'No FAQ sheet found');
      return null;
    }
    
    const morpths = filterGNL(gNL(question), ['NOUN', 'NUM', 'NUMBER']);
    let words = [];
    
    for (let i = 0; i < morpths.length; i++) {
      let d = katakanaToHiragana(
        toHalfWidth(morpths[i]).toLowerCase().replace(',', '')
      );
      if (d.indexOf('-') !== -1) {
        const arr = morpths[i].split('-');
        for (let n = 0; n < arr.length; n++) {
          words.push(
            katakanaToHiragana(
              toHalfWidth(arr[n]).toLowerCase().replace(',', '')
            )
          );
        }
        continue;
      }
      words.push(d);
    }

    const faqs = faqSheet
      .getRange('A:B')
      .getValues()
      .filter((row) => !row.every((cell) => cell.toString().trim() === ''));
      
    let sfaqs = [], result = [];
    for (let i = 1; i < faqs.length; i++) {
      sfaqs[i] = faqs[i].map((cell) =>
        katakanaToHiragana(toHalfWidth(cell.toString()).toLowerCase().replace(',', ''))
      );
    }
    
    for (let i = 1; i < sfaqs.length; i++) {
      if (sfaqs[i].some((faq) => words.some((w) => faq.includes(w)))) {
        if (result.length === 0) result.push(faqs[0]);
        result.push(faqs[i]);
      }
    }
    
    if (!result.length) {
      debugLog('FAQ', 'No matching FAQ found');
      return null;
    }
    
    debugLog('FAQ', 'Found matching FAQs', result.length);
    return {
      role: 'system',
      content: '今から記載するJSON形式のFAQを踏まえて回答を望む(FAQの回答とは言わない)' + JSON.stringify(result),
    };
  } catch (e) {
    debugLog('FAQ', 'Error', e.toString());
    return null;
  }
}

function mergeRoleAndThread(optionRole, threadMessages) {
  for (let i = 0; i < threadMessages.length; i++) {
    optionRole.push({
      role: threadMessages[i].hasOwnProperty('app_id') ? 'assistant' : 'user',
      content: threadMessages[i].text || '',
    });
  }
}

// ===========================
// Drive検索機能
// ===========================

/**
 * FAQシートの A列キーワード、C列チェックを元に、
 * ドライブ一覧シート A列のすべてのフォルダIDを検索対象として
 * 指定キーワードの含まれるファイル本文／行を抜き出し、
 * D列に結果をリッチテキストで書き出します。
 */
function updateFaqDriveResults() {
  const ss             = SpreadsheetApp.getActive();
  const faqSheet       = ss.getSheetByName('faq');
  const driveListSheet = ss.getSheetByName('ドライブ一覧');
  const lastFaqRow     = faqSheet.getLastRow();
  const lastDriveRow   = driveListSheet.getLastRow();
  if (lastFaqRow < 2 || lastDriveRow < 2) return;

  const folderIds = driveListSheet
    .getRange(`A2:A${lastDriveRow}`)
    .getValues()
    .flat()
    .filter(id => id);

  const faqData = faqSheet.getRange(`A2:C${lastFaqRow}`).getValues();

  faqData.forEach((row, i) => {
    const [ keyword, /*manualAnswer*/, doSearch ] = row;
    const rowNum = i + 2;
    const resultCell = faqSheet.getRange(rowNum, 4);

    if (doSearch === true) {
      let allResults = [];
      folderIds.forEach(folderId => {
        const res = searchDriveLinkReturn(keyword, folderId);
        allResults = allResults.concat(res);
      });
      if (allResults.length) {
        cellSetLink(resultCell, allResults);
      } else {
        resultCell.setValue('該当ファイルが見つかりませんでした');
      }
    } else {
      resultCell.clearContent();
    }
  });
}

/**
 * Drive API で指定フォルダ内を全文検索
 */
function searchDriveLinkReturn(keyword, folderId) {
  const ret = [];
  const baseUrl = 'https://www.googleapis.com/drive/v3/files';
  const params = {
    q: `\'${folderId}\' in parents and trashed = false and fullText contains '${keyword}'`,
    corpora: 'allDrives',
    includeItemsFromAllDrives: true,
    supportsAllDrives: true,
    fields: 'files(id,name,mimeType,webViewLink)'
  };
  const query = Object.entries(params)
    .map(([k,v]) => `${encodeURIComponent(k)}=${encodeURIComponent(v)}`)
    .join('&');
  const url = `${baseUrl}?${query}`;

  const response = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }
  });
  const files = JSON.parse(response.getContentText()).files || [];

  files.forEach(file => {
    let snippets = [];
    try {
      if (file.mimeType === 'application/vnd.google-apps.document') {
        snippets = extractSnippetFromDoc(file.id, keyword);
      } else if (file.mimeType === 'application/vnd.google-apps.spreadsheet') {
        snippets = extractSnippetFromSheet(file.id, keyword);
      } else if (file.mimeType === 'application/pdf') {
        const blob      = DriveApp.getFileById(file.id).getBlob();
        const tmpFile   = DriveApp.createFile(blob).setName('temp');
        const resource  = { title: 'temp-doc', mimeType: MimeType.GOOGLE_DOCS };
        const converted = Drive.Files.insert(resource, tmpFile.getBlob());
        snippets        = extractSnippetFromDoc(converted.id, keyword);
        DriveApp.getFileById(converted.id).setTrashed(true);
        tmpFile.setTrashed(true);
      }
    } catch (e) {
      Logger.log(`処理エラー (${file.name}): ${e}`);
    }
    if (snippets.length > 0) {
      ret.push({ file: file, snippets: snippets });
    }
  });

  return ret;
}

/**
 * Googleドキュメントからキーワードを含む段落を抽出
 */
function extractSnippetFromDoc(docId, keyword) {
  const paras = DocumentApp.openById(docId).getBody().getParagraphs();
  return paras
    .map(p => p.getText().trim())
    .filter(t => t.includes(keyword));
}

/**
 * スプレッドシートからキーワードを含む行を抽出
 */
function extractSnippetFromSheet(sheetId, keyword) {
  const rows = SpreadsheetApp.openById(sheetId)
    .getSheets()
    .flatMap(sh => sh.getDataRange().getValues());
  return rows
    .filter(r => r.some(c => c.toString().includes(keyword)))
    .map(r => r.join('\t'));
}

/**
 * 結果をリッチテキストでセルに書き込む
 */
function cellSetLink(range, data) {
  const maxLen = 5000;
  let text     = '';
  const links  = [];

  data.forEach(item => {
    const nameBlock    = item.file.name + '\n';
    const snippetBlock = item.snippets
      .slice(0, 2)
      .map(s => s.replace(/\t/g,' ').replace(/\n/g,' ').trim())
      .join('\n') + '\n';

    let block = nameBlock + snippetBlock;
    if (text.length + block.length > maxLen) {
      block = block.substring(0, maxLen - text.length);
    }

    const start = text.length;
    text += block;
    const end   = start + nameBlock.length;
    if (end <= maxLen) {
      links.push({ start, end, url: item.file.webViewLink });
    }
  });

  const builder = SpreadsheetApp.newRichTextValue().setText(text);
  links.forEach(l => builder.setLinkUrl(l.start, l.end, l.url));
  range.setRichTextValue(builder.build());
}

// ===========================
// AI レスポンス機能（改良版）
// ===========================

function chatGPTResponse(message, { optionRole = [], temperature }) {
  debugLog('AI', 'ChatGPT request', { messageLength: message?.length, roles: optionRole.length });
  
  const config = Settings();
  if (!config?.OPEN_AI_TOKEN) {
    debugLog('AI', 'No OpenAI token');
    return 'OpenAI APIトークンが設定されていません。';
  }
  
  const apiKey = config.OPEN_AI_TOKEN;
  const url = 'https://api.openai.com/v1/chat/completions';
  const payload = {
    model: 'gpt-4',
    messages: [...optionRole, { role: 'user', content: message }],
    temperature: temperature ?? 1,
  };
  
  const options = {
    method: 'post',
    muteHttpExceptions: true,
    headers: {
      'Authorization': 'Bearer ' + apiKey,
      'Content-Type': 'application/json',
    },
    payload: JSON.stringify(payload),
  };
  
  try {
    const res = UrlFetchApp.fetch(url, options);
    const result = JSON.parse(res.getContentText());
    
    if (result.error) {
      debugLog('AI', 'ChatGPT error', result.error);
      return `AIエラー: ${result.error.message}`;
    }
    
    const content = result?.choices?.[0]?.message?.content || '';
    debugLog('AI', 'ChatGPT success', { responseLength: content.length });
    return content;
  } catch (e) {
    debugLog('AI', 'ChatGPT exception', e.toString());
    return `AI処理エラー: ${e.toString()}`;
  }
}

function geminiResponse(text, optionRole) {
  // Gemini APIの実装
  debugLog('AI', 'Gemini not implemented');
  return '';
}

// ===========================
// メインエントリポイント（改良版）
// ===========================

/**
 * テスト用エンドポイント
 */
function doGet(e) {
  debugLog('Main', 'GET request received', e.parameter);
  
  return ContentService
    .createTextOutput(JSON.stringify({
      status: 'ok',
      message: 'Slack Bot is running',
      timestamp: new Date().toISOString()
    }))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * メインエントリポイント（改良版）
 */
function doPost(e) {
  debugLog('Main', 'POST request received');
  
  try {
    // リクエストの詳細をログ
    debugLog('Main', 'Request details', {
      contentLength: e.contentLength,
      queryString: e.queryString,
      contextPath: e.contextPath,
      hasPostData: !!e.postData
    });
    
    if (!e.postData) {
      debugLog('Main', 'No postData');
      return ContentService.createTextOutput('No data');
    }
    
    const params = JSON.parse(e.postData.contents);
    debugLog('Main', 'Parsed params', { type: params.type, event_type: params.event?.type });
    
    // URL検証の処理
    if (params.type === 'url_verification') {
      debugLog('Main', 'URL verification');
      return ContentService.createTextOutput(params.challenge);
    }
    
    // イベントコールバックの処理
    if (params.type !== 'event_callback') {
      debugLog('Main', 'Not event_callback', params.type);
      return ContentService.createTextOutput('');
    }
    
    // 重複受信防止
    const cache = CacheService.getScriptCache();
    if (params.event && params.event.client_msg_id) {
      if (cache.get(params.event.client_msg_id) === 'done') {
        debugLog('Main', 'Duplicate message');
        return ContentService.createTextOutput('');
      }
      cache.put(params.event.client_msg_id, 'done', 600);
    }
    
    const Bot = new SlackBot(e);
    
    // メンションイベント用ハンドラ
    const run = ({ event, message, channel, threadTs }) => {
      debugLog('Main', 'Run handler', { channel, message, threadTs });
      
      try {
        // チャンネル説明を取得
        const channelInfo = getChannelInfo(channel);
        const channelDescription = channelInfo?.purpose?.value || 'チャンネル説明がありません。';
        
        // ベースロール定義
        const baseRole = [
          { role: 'system', content: 'You are a helpful assistant.' },
          { role: 'system', content: 'Please answer in Japanese.' },
          { role: 'system', content: 'Be concise and clear.' }
        ];
        
        // メンション部分を除去
        let text = message || '';
        const mentionMatch = text.match(/<@[A-Z0-9]+>/);
        if (mentionMatch) {
          text = text.replace(mentionMatch[0], '').trim();
        }
        
        debugLog('Main', 'Processed text', text);
        
        // 簡易コマンド対応
        if (text === 'Hello' || text === 'hello') {
          postMessage('こんにちは！ご用件をお聞かせください。', channel, threadTs || event.ts);
          return;
        }
        
        if (text === 'help') {
          postMessage('このボットはAIアシスタントです。質問や依頼をメンションと共に送信してください。', channel, threadTs || event.ts);
          return;
        }
        
        // AI レスポンス取得
        const optionRole = [...baseRole];
        
        // FAQから追加ロールを取得
        const faq = getFaqRole(text);
        if (faq) optionRole.push(faq);
        
        // スレッド履歴をマージ
        const threadMessages = event.thread_ts ? getThreadMessages(channel, event.thread_ts) : [];
        mergeRoleAndThread(optionRole, threadMessages);
        
        let responseText = '';
        if (Settings().GEMINI_TOKEN) {
          responseText = geminiResponse(text, optionRole);
        } else {
          responseText = chatGPTResponse(text, { optionRole });
        }
        
        if (!responseText) {
          responseText = '申し訳ありません。応答の生成に失敗しました。';
        }
        
        // スレッドIDを決定
        const replyThread = threadTs || event.ts;
        
        // メッセージ送信
        const success = postMessage(responseText, channel, replyThread);
        debugLog('Main', 'Message post result', success);
        
      } catch (error) {
        debugLog('Main', 'Handler error', error.toString());
        postMessage('エラーが発生しました: ' + error.toString(), channel, event.ts);
      }
    };
    
    // メンションされたときのみ処理
    Bot.handleMentionEventBase(run);
    
    return Bot.response();
    
  } catch (error) {
    debugLog('Main', 'Fatal error', error.toString());
    return ContentService.createTextOutput('');
  }
}

// ===========================
// デバッグ用テスト関数
// ===========================

/**
 * 設定のテスト
 */
function testSettings() {
  try {
    const settings = Settings();
    console.log('Settings test passed:', settings);
    return true;
  } catch (e) {
    console.log('Settings test failed:', e.toString());
    return false;
  }
}

/**
 * Slack API接続テスト
 */
function testSlackConnection() {
  const config = Settings();
  const url = 'https://slack.com/api/auth.test';
  
  const options = {
    method: 'post',
    payload: {
      token: config.SLACK_TOKEN
    },
    muteHttpExceptions: true
  };
  
  try {
    const response = UrlFetchApp.fetch(url, options);
    const data = JSON.parse(response.getContentText());
    
    console.log('Slack connection test:', data);
    
    if (data.ok) {
      console.log('Connected as:', data.user);
      console.log('Team:', data.team);
      return true;
    } else {
      console.log('Connection failed:', data.error);
      return false;
    }
  } catch (e) {
    console.log('Connection test error:', e.toString());
    return false;
  }
}

/**
 * テストメッセージの送信
 */
function testPostMessage() {
  const testChannel = 'C1234567890'; // テスト用チャンネルIDに変更してください
  const testMessage = 'テストメッセージ: ' + new Date().toISOString();
  
  const result = postMessage(testMessage, testChannel);
  console.log('Test message result:', result);
  return result;
}

// ===========================
// 設定情報
// ===========================

/**
 * 必要なスクリプトプロパティ:
 * - SLACK_TOKEN: Slack Bot Token
 * - OPEN_AI_TOKEN: OpenAI API Key
 * - GEMINI_TOKEN: Gemini API Key (オプション)
 * - GOOGLE_NL_API: Google Natural Language API Key (オプション)
 * 
 * デバッグ手順:
 * 1. testSettings() を実行して設定を確認
 * 2. testSlackConnection() を実行してSlack接続を確認
 * 3. testPostMessage() を実行してメッセージ送信を確認
 * 4. デバッグログは debug_log シートで確認
 */