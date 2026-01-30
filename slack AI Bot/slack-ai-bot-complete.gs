/**
 * Slack AI Bot - 統合版（完全版・デバッグ機能付き）
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
 */

// ===========================
// デバッグ設定
// ===========================
const DEBUG_MODE = true; // デバッグモードの有効/無効
const DEBUG_SHEET_NAME = 'debug_log'; // デバッグ用のシート名
const SPREADSHEET_NAME = 'Slack Bot Data'; // スプレッドシート名

/**
 * スプレッドシートの取得または作成
 */
function getOrCreateSpreadsheet() {
  try {
    // スクリプトプロパティから既存のIDを確認
    const savedId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
    if (savedId) {
      try {
        const ss = SpreadsheetApp.openById(savedId);
        console.log('Found spreadsheet by saved ID: ' + ss.getUrl());
        return ss;
      } catch (e) {
        console.log('Saved spreadsheet ID is invalid, searching by name...');
      }
    }
    
    // 名前で既存のスプレッドシートを探す
    const files = DriveApp.getFilesByName(SPREADSHEET_NAME);
    
    if (files.hasNext()) {
      const file = files.next();
      const ss = SpreadsheetApp.openById(file.getId());
      console.log('Found existing spreadsheet by name: ' + ss.getUrl());
      // IDを保存
      PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', ss.getId());
      return ss;
    }
    
    // 存在しない場合は新規作成
    console.log('Creating new spreadsheet: ' + SPREADSHEET_NAME);
    
    // 新規作成して即座にIDで開き直す（より確実）
    const newSS = SpreadsheetApp.create(SPREADSHEET_NAME);
    const ssId = newSS.getId();
    
    // 作成を確実にするため少し待機
    Utilities.sleep(2000);
    
    // IDで開き直す
    const ss = SpreadsheetApp.openById(ssId);
    
    if (ss) {
      console.log('Spreadsheet created successfully');
      console.log('ID: ' + ss.getId());
      console.log('URL: ' + ss.getUrl());
      
      // スプレッドシートIDを保存
      PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', ssId);
      
      // シートを初期化
      initializeSheets(ss);
      
      return ss;
    } else {
      throw new Error('Failed to create spreadsheet');
    }
    
  } catch (e) {
    console.log('Error in getOrCreateSpreadsheet: ' + e.toString());
    console.log('Attempting fallback method...');
    
    // フォールバック: 手動でスプレッドシートを作成
    try {
      const ss = createSpreadsheetManually();
      return ss;
    } catch (e2) {
      console.log('Fallback also failed: ' + e2.toString());
      throw e2;
    }
  }
}

/**
 * 手動でスプレッドシートを作成（フォールバック）
 */
function createSpreadsheetManually() {
  console.log('Using manual spreadsheet creation method...');
  
  // 新しいスプレッドシートを作成
  const ss = SpreadsheetApp.create(SPREADSHEET_NAME + ' ' + new Date().getTime());
  
  if (!ss) {
    throw new Error('Cannot create spreadsheet');
  }
  
  const ssId = ss.getId();
  PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', ssId);
  
  // デフォルトシートの名前を変更
  const sheets = ss.getSheets();
  if (sheets.length > 0) {
    sheets[0].setName('log');
  }
  
  // 必要なシートを追加
  try {
    const faqSheet = ss.insertSheet('faq');
    faqSheet.appendRow(['キーワード', '回答', '検索フラグ', 'Drive検索結果']);
    
    const driveSheet = ss.insertSheet('ドライブ一覧');
    driveSheet.appendRow(['フォルダID', 'フォルダ名', '説明']);
    
    const debugSheet = ss.insertSheet('debug_log');
    debugSheet.appendRow(['Timestamp', 'Category', 'Message', 'Data']);
  } catch (e) {
    console.log('Error adding sheets: ' + e.toString());
  }
  
  console.log('Manual creation successful: ' + ss.getUrl());
  return ss;
}

/**
 * 必要なシートの初期化
 */
function initializeSheets(ss) {
  if (!ss) {
    console.log('Error: Spreadsheet object is undefined');
    return;
  }
  
  try {
    // デフォルトシートの名前を変更
    const sheets = ss.getSheets();
    if (sheets && sheets.length > 0) {
      sheets[0].setName('log');
      // ログシートのヘッダー設定
      const logSheet = sheets[0];
      if (logSheet.getLastRow() === 0) {
        logSheet.appendRow(['Timestamp', 'Message']);
        logSheet.getRange('1:1').setFontWeight('bold');
        logSheet.setFrozenRows(1);
      }
    }
    
    // FAQシートの作成（既存の場合はスキップ）
    let faqSheet = ss.getSheetByName('faq');
    if (!faqSheet) {
      faqSheet = ss.insertSheet('faq');
      faqSheet.appendRow(['キーワード', '回答', '検索フラグ', 'Drive検索結果']);
      faqSheet.getRange('1:1').setFontWeight('bold');
      faqSheet.setFrozenRows(1);
    }
    
    // ドライブ一覧シートの作成（既存の場合はスキップ）
    let driveSheet = ss.getSheetByName('ドライブ一覧');
    if (!driveSheet) {
      driveSheet = ss.insertSheet('ドライブ一覧');
      driveSheet.appendRow(['フォルダID', 'フォルダ名', '説明']);
      driveSheet.getRange('1:1').setFontWeight('bold');
      driveSheet.setFrozenRows(1);
    }
    
    // デバッグログシートの作成（既存の場合はスキップ）
    let debugSheet = ss.getSheetByName(DEBUG_SHEET_NAME);
    if (!debugSheet) {
      debugSheet = ss.insertSheet(DEBUG_SHEET_NAME);
      debugSheet.appendRow(['Timestamp', 'Category', 'Message', 'Data']);
      debugSheet.getRange('1:1').setFontWeight('bold');
      debugSheet.setFrozenRows(1);
    }
    
    console.log('Initialized sheets: log, faq, ドライブ一覧, ' + DEBUG_SHEET_NAME);
  } catch (e) {
    console.log('Error initializing sheets: ' + e.toString());
  }
}

/**
 * アクティブなスプレッドシートを取得
 */
function getActiveSpreadsheet() {
  try {
    // まずスクリプトプロパティからIDを取得
    const spreadsheetId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
    if (spreadsheetId) {
      try {
        return SpreadsheetApp.openById(spreadsheetId);
      } catch (e) {
        console.log('Could not open spreadsheet by ID, creating new one');
      }
    }
    
    // IDがない場合は名前で検索または作成
    return getOrCreateSpreadsheet();
  } catch (e) {
    console.log('Error getting spreadsheet: ' + e.toString());
    return getOrCreateSpreadsheet();
  }
}

/**
 * デバッグログを記録
 */
function debugLog(category, message, data = null) {
  if (!DEBUG_MODE) return;
  
  console.log(`[${category}] ${message}`, data);
  Logger.log(`[${category}] ${message} ${data ? JSON.stringify(data) : ''}`);
  
  // スプレッドシートにも記録
  try {
    const ss = getActiveSpreadsheet();
    if (!ss) {
      Logger.log('Could not get spreadsheet for debug logging');
      return;
    }
    
    let debugSheet = ss.getSheetByName(DEBUG_SHEET_NAME);
    if (!debugSheet) {
      debugSheet = ss.insertSheet(DEBUG_SHEET_NAME);
      debugSheet.appendRow(['Timestamp', 'Category', 'Message', 'Data']);
      debugSheet.getRange('1:1').setFontWeight('bold');
      debugSheet.setFrozenRows(1);
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
      const ss = getActiveSpreadsheet();
      if (!ss) {
        console.log('SheetLog: No spreadsheet available');
        return;
      }
      
      let logSheet = ss.getSheetByName('log');
      if (!logSheet) {
        logSheet = ss.insertSheet('log');
        logSheet.appendRow(['Timestamp', 'Message']);
        logSheet.getRange('1:1').setFontWeight('bold');
        logSheet.setFrozenRows(1);
      }
      
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
    // カタカナの文字コードからひらがなの文字コードへ変換
    var chr = match.charCodeAt(0) - 0x60;
    return String.fromCharCode(chr);
  });
}

function toHalfWidth(str) {
  // 全角英数字を半角に変換
  str = str.replace(/[Ａ-Ｚａ-ｚ０-９]/g, function(s) {
    return String.fromCharCode(s.charCodeAt(0) - 0xFEE0);
  });
  return str;
}

// ===========================
// Slack Bot クラス
// ===========================

class SlackBot {
  constructor(e) {
    this.requestEvent = e;
    this.postData = null;
    this.slackEvent = null;
    this.responseData = this.init();
    this.verification();
  }

  responseJsonData(json) {
    return ContentService.createTextOutput(JSON.stringify(json)).setMimeType(ContentService.MimeType.JSON);
  }

  init() {
    const e = this.requestEvent;
    if (!e?.postData) return { error: 'postData is missing or undefined.', request: JSON.stringify(e, null, "  ") };
    this.postData = e.postData;
    if (!this.postData?.type) return { error: 'postData type is missing or undefined.', request: JSON.stringify(this.postData, null, "  ") };
    try { var event = JSON.parse(this.postData.contents); }
    catch (error) {
      event = e.parameter?.command && e.parameter?.text ? { event: { type: "command", event: { ...e.parameter } } } : { error: 'Invalid JSON format in postData contents.', request: this.postData };
    }
    this.slackEvent = event;
    return event?.event ? null : { error: 'Slack event is missing or undefined.', request: JSON.stringify(e, null, "  ") };
  }

  verification() {
    //SheetLog.log(JSON.stringify(this.postData));
    if (!this.postData || this.responseData) return null;
    if (this.postData.type !== 'url_verification') return null;
    this.responseData = { "challenge": this.postData.challenge };
    return this.responseData;
  }

  hasCache(key) {
    if (!key) return true;
    const cache = CacheService.getScriptCache();
    const cached = cache.get(key);
    if (cached) return true;
    cache.put(key, true, 30 * 60);
    return false;
  }

  handleEvent(type, callback = () => { }) {
    if (!this.slackEvent || this.responseData || this.slackEvent?.event?.type !== type) return null;
    const callbackResponse = callback({ event: this.slackEvent.event });
    if (!callbackResponse) return null;
    this.responseData = callbackResponse;
    return callbackResponse;
  }

  handleBase(type, targetType, callback = () => {}) {
    return this.handleEvent(type, ({ event }) => {
      const { text: message, channel, thread_ts: threadTs, ts, client_msg_id, bot_id, app_id } = event;
      if (bot_id || app_id) return null;
      if (event.type !== targetType || this.hasCache(`${channel}:${client_msg_id}`)) return null;
      return callback ? callback({ message, channel, threadTs: threadTs ?? ts, event }) : null;
    });
  }

  handleMessageEventBase(callback) { 
    return this.handleBase("message", "message", callback); 
  }
  
  handleMentionEventBase(callback) { 
    return this.handleBase("app_mention", "app_mention", callback); 
  }
  
  handleReactionEventBase(callback) { 
    return this.handleBase("reaction_added", "reaction_added", callback); 
  }

  response() {
    Logger.log(this.responseData);
    return this.responseData && this.responseJsonData(this.responseData);
  }
}

// ===========================
// Slack API 関連機能
// ===========================

/**
 * チャンネル情報取得
 */
function getChannelInfo(channelId) {
  const url = 'https://slack.com/api/conversations.info';
  const config = Settings();
  if (!config?.SLACK_TOKEN) return;
  const payload = {
    token: config.SLACK_TOKEN,
    channel: channelId,
  };
  const options = {
    method: 'post',
    payload,
  };
  const response = UrlFetchApp.fetch(url, options);
  const data = JSON.parse(response.getContentText());
  return data.channel;
}

/**
 * スレッドメッセージ取得
 */
function getThreadMessages(channelId, threadTs) {
  const url = 'https://slack.com/api/conversations.replies';
  const config = Settings();
  if (!config?.SLACK_TOKEN) return [];
  const payload = {
    token: config.SLACK_TOKEN,
    channel: channelId,
    ts: threadTs,
  };
  const options = {
    method: 'get',
    payload,
  };
  const response = UrlFetchApp.fetch(url, options);
  const data = JSON.parse(response.getContentText());
  return data.messages || [];
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

/**
 * Google Natural Language API alnalyzeSyntax
 */
function gNL(textdata) {
  var apiKey = ScriptProperties.getProperty('GOOGLE_NL_API');  // ここに取得したAPIキーを入れる
  //形態素解析（品詞取得） = analyzeSyntax
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
    payload: JSON.stringify(payload)
  };
  
  var response = UrlFetchApp.fetch(url, options);
  //SheetLog.log('NL:' + response.getContentText());
  //Logger.log(response.getContentText());
  try {
    return JSON.parse(response.getContentText());  
  } catch(e) {
    Logger.log(response.getContentText());
    Logger.log(e);
    return null;
  }
}

/**
 * Google Natural Language API の戻り値より必要なものを抽出する
 * 品詞の場合は tagsの欄に ['NOUN','NUM','NUMBER']
 * https://cloud.google.com/natural-language/docs/morphology?hl=ja
 */
function filterGNL(gNLobj, tags) {
  if (!gNLobj) return [];
  var words = gNLobj.tokens
    .filter(token => tags.includes(token.partOfSpeech.tag)) // 配列内の品詞と一致するものを抽出
    .map(token => token.text.content); 
  return words;
}

// ===========================
// FAQ検索機能
// ===========================

/**
 * FAQロールを取得
 */
function getFaqRole(question) {
  try {
    const ss = getActiveSpreadsheet();
    if (!ss) {
      debugLog('FAQ', 'No spreadsheet available');
      return null;
    }
    
    let faqSheet = ss.getSheetByName('faq');
    if (!faqSheet) {
      debugLog('FAQ', 'FAQ sheet not found, creating');
      faqSheet = ss.insertSheet('faq');
      faqSheet.appendRow(['キーワード', '回答', '検索フラグ', 'Drive検索結果']);
      faqSheet.getRange('1:1').setFontWeight('bold');
      faqSheet.setFrozenRows(1);
      return null; // 新規作成した場合はデータがない
    }
    
    const morpths = filterGNL(gNL(question), ['NOUN', 'NUM', 'NUMBER']);
    let words = [];
    for (let i = 0; i < morpths.length; i++) {
      let d = katakanaToHiragana(
        toHalfWidth(morpths[i]).toLowerCase().replace(',', '')
      );
      if (d.indexOf('-')) {
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
        katakanaToHiragana(toHalfWidth(cell).toLowerCase().replace(',', ''))
      );
    }
    for (let i = 1; i < sfaqs.length; i++) {
      if (sfaqs[i].some((faq) => words.some((w) => faq.includes(w)))) {
        if (result.length === 0) result.push(faqs[0]);
        result.push(faqs[i]);
      }
    }
    if (!result.length) return null;
    return {
      role: 'system',
      content:
        '今から記載するJSON形式のFAQを踏まえて回答を望む(FAQの回答とは言わない)' +
        JSON.stringify(result),
    };
  } catch (e) {
    return null;
  }
}

/**
 * スレッド履歴とロールをマージ
 */
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
 *
 * B列の手動回答はそのまま残し、FAQシートの構造は変更しません。
 */
function updateFaqDriveResults() {
  const ss = getActiveSpreadsheet();
  if (!ss) {
    debugLog('Drive', 'No spreadsheet available');
    return;
  }
  
  const faqSheet = ss.getSheetByName('faq');
  const driveListSheet = ss.getSheetByName('ドライブ一覧');
  
  if (!faqSheet || !driveListSheet) {
    debugLog('Drive', 'Required sheets not found');
    return;
  }
  
  const lastFaqRow = faqSheet.getLastRow();
  const lastDriveRow = driveListSheet.getLastRow();
  if (lastFaqRow < 2 || lastDriveRow < 2) return;

  // ドライブ一覧シート A2:A に書かれたフォルダIDを取得
  const folderIds = driveListSheet
    .getRange(`A2:A${lastDriveRow}`)
    .getValues()
    .flat()
    .filter(id => id);

  // FAQシートの A列キーワード、C列検索フラグを取得
  const faqData = faqSheet.getRange(`A2:C${lastFaqRow}`).getValues();

  faqData.forEach((row, i) => {
    const [ keyword, /*manualAnswer*/, doSearch ] = row;
    const rowNum = i + 2;
    const resultCell = faqSheet.getRange(rowNum, 4); // D列

    if (doSearch === true) {
      // C列が TRUE の場合のみ、全フォルダを検索して結果をまとめる
      let allResults = [];
      folderIds.forEach(folderId => {
        const res = searchDriveLinkReturn(keyword, folderId);
        allResults = allResults.concat(res);
      });
      // 抜き出した結果を D列に書き込む
      if (allResults.length) {
        cellSetLink(resultCell, allResults);
      } else {
        resultCell.setValue('該当ファイルが見つかりませんでした');
      }
    } else {
      // C列が FALSE の場合は D列をクリア
      resultCell.clearContent();
    }
  });
}

/**
 * Drive API で指定フォルダ内を全文検索し、
 * ファイルごとに段落／行を抜き出して配列で返す
 * @returns Array<{file: Object, snippets: string[]}>
 */
function searchDriveLinkReturn(keyword, folderId) {
  const ret = [];
  const baseUrl = 'https://www.googleapis.com/drive/v3/files';
  const params = {
    q:                          `\'${folderId}\' in parents and trashed = false and fullText contains '${keyword}'`,
    corpora:                    'allDrives',
    includeItemsFromAllDrives:  true,
    supportsAllDrives:          true,
    fields:                     'files(id,name,mimeType,webViewLink)'
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
        // PDF を Docs に変換して抜き出す
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
 * スプレッドシートからキーワードを含む行を抽出 (タブ区切り)
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
 * 結果をリッチテキスト (リンク付き) でセルに書き込む
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
// ファイル処理機能
// ===========================

/**
 * Slackイベントからファイル情報を取得して処理
 */
function handleFileShared(event, channel) {
  const files = event.files || [];
  if (files.length === 0) return null;
  
  const results = [];
  
  files.forEach(file => {
    try {
      const fileContent = downloadAndProcessFile(file);
      if (fileContent) {
        results.push({
          name: file.name,
          type: file.mimetype,
          content: fileContent
        });
      }
    } catch (e) {
      Logger.log(`ファイル処理エラー: ${file.name} - ${e.toString()}`);
      results.push({
        name: file.name,
        type: file.mimetype,
        error: `ファイル処理中にエラーが発生しました: ${e.toString()}`
      });
    }
  });
  
  return results;
}

/**
 * Slackからファイルをダウンロードして処理
 */
function downloadAndProcessFile(file) {
  const config = Settings();
  if (!config?.SLACK_TOKEN) return null;
  
  // Slack APIを使用してファイルをダウンロード
  const downloadUrl = file.url_private_download || file.url_private;
  
  const options = {
    method: 'get',
    headers: {
      'Authorization': 'Bearer ' + config.SLACK_TOKEN
    },
    muteHttpExceptions: true
  };
  
  const response = UrlFetchApp.fetch(downloadUrl, options);
  const blob = response.getBlob();
  
  // ファイルタイプに応じて処理
  const mimeType = file.mimetype;
  
  if (mimeType === 'application/pdf') {
    return processPDF(blob);
  } else if (mimeType === 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' || 
             mimeType === 'application/msword') {
    return processWord(blob);
  } else if (mimeType === 'text/plain') {
    return blob.getDataAsString();
  } else if (file.name && file.name.includes('docs.google.com')) {
    // GoogleドキュメントのURLの場合
    return processGoogleDoc(file.url_private);
  } else {
    return `ファイルタイプ ${mimeType} はサポートされていません。`;
  }
}

/**
 * PDFファイルを処理
 */
function processPDF(blob) {
  try {
    // PDFをGoogle Docsに変換して読み取る
    const resource = {
      title: 'temp-pdf-' + Utilities.getUuid(),
      mimeType: MimeType.GOOGLE_DOCS
    };
    
    // Drive APIを使用してPDFをGoogle Docsに変換
    const file = Drive.Files.insert(resource, blob, {
      ocr: true,
      ocrLanguage: 'ja'
    });
    
    // 変換されたドキュメントからテキストを抽出
    const doc = DocumentApp.openById(file.id);
    const text = doc.getBody().getText();
    
    // 一時ファイルを削除
    DriveApp.getFileById(file.id).setTrashed(true);
    
    return text;
  } catch (e) {
    Logger.log('PDF処理エラー: ' + e.toString());
    throw new Error('PDFの読み取りに失敗しました: ' + e.toString());
  }
}

/**
 * Wordファイルを処理
 */
function processWord(blob) {
  try {
    // WordファイルをGoogle Docsに変換
    const resource = {
      title: 'temp-word-' + Utilities.getUuid(),
      mimeType: MimeType.GOOGLE_DOCS
    };
    
    // Drive APIを使用してWordをGoogle Docsに変換
    const file = Drive.Files.insert(resource, blob);
    
    // 変換されたドキュメントからテキストを抽出
    const doc = DocumentApp.openById(file.id);
    const text = doc.getBody().getText();
    
    // 一時ファイルを削除
    DriveApp.getFileById(file.id).setTrashed(true);
    
    return text;
  } catch (e) {
    Logger.log('Word処理エラー: ' + e.toString());
    throw new Error('Wordファイルの読み取りに失敗しました: ' + e.toString());
  }
}

/**
 * GoogleドキュメントのURLから内容を取得
 */
function processGoogleDoc(url) {
  try {
    // URLからドキュメントIDを抽出
    const docIdMatch = url.match(/\/d\/([a-zA-Z0-9-_]+)/);
    if (!docIdMatch) {
      throw new Error('Google ドキュメントのIDが見つかりません');
    }
    
    const docId = docIdMatch[1];
    
    // ドキュメントを開いてテキストを取得
    const doc = DocumentApp.openById(docId);
    const text = doc.getBody().getText();
    
    return text;
  } catch (e) {
    Logger.log('Google Docs処理エラー: ' + e.toString());
    throw new Error('Google ドキュメントの読み取りに失敗しました: ' + e.toString());
  }
}

/**
 * ファイル付きメッセージのハンドラ
 */
function handleMessageWithFiles(event) {
  const { text, channel, thread_ts, ts, files } = event;
  
  if (!files || files.length === 0) {
    return null;
  }
  
  // ファイルを処理
  const fileResults = handleFileShared(event, channel);
  
  // ファイル内容を含めて応答を生成
  if (fileResults && fileResults.length > 0) {
    const fileContents = fileResults.map(r => r.content || r.error).join('\n\n');
    
    // ファイル内容を含めたコンテキストを作成
    const context = {
      userMessage: text || 'ファイルが添付されました',
      fileContents: fileContents,
      fileInfo: fileResults.map(r => ({
        name: r.name,
        type: r.type
      }))
    };
    
    return context;
  }
  
  return null;
}

// ===========================
// ドキュメント編集・修正案作成機能
// ===========================

/**
 * ドキュメントを処理して修正案を作成
 */
function processDocumentForReview(file, userMessage, channel) {
  try {
    let docId = null;
    let originalFileName = '';
    
    // ファイルタイプに応じて処理
    if (file.mimetype === 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' || 
        file.mimetype === 'application/msword') {
      // Wordファイルの場合
      const result = convertWordToGoogleDoc(file);
      docId = result.docId;
      originalFileName = file.name;
    } else if (file.url_private && file.url_private.includes('docs.google.com')) {
      // GoogleドキュメントのURLの場合
      const originalDocId = extractDocIdFromUrl(file.url_private);
      docId = copyGoogleDoc(originalDocId);
      originalFileName = file.name || 'Google Document';
    } else {
      throw new Error('このファイルタイプはドキュメント編集に対応していません');
    }
    
    if (!docId) {
      throw new Error('ドキュメントの処理に失敗しました');
    }
    
    // ドキュメントの内容を取得
    const doc = DocumentApp.openById(docId);
    const body = doc.getBody();
    const originalText = body.getText();
    
    // ドキュメント名を設定
    const timestamp = Utilities.formatDate(new Date(), 'JST', 'yyyy-MM-dd_HH:mm');
    doc.setName(`【修正版】${originalFileName}_${timestamp}`);
    
    // AIに修正案を依頼
    const suggestions = getAISuggestions(originalText, userMessage);
    
    // 修正案をドキュメントに適用
    applyAISuggestionsToDocument(doc, suggestions);
    
    // ドキュメントに編集履歴のヘッダーを追加
    addReviewHeader(doc, originalFileName, userMessage);
    
    // 共有設定（元のドキュメントと同じ設定をコピー）
    const originalDocId = file.url_private && file.url_private.includes('docs.google.com') 
      ? extractDocIdFromUrl(file.url_private) 
      : null;
    setDocumentSharing(docId, originalDocId);
    
    // ドキュメントのURLを取得
    const docUrl = doc.getUrl();
    
    return {
      success: true,
      url: docUrl,
      docId: docId,
      fileName: doc.getName(),
      originalFileName: originalFileName,
      suggestionsCount: suggestions.length
    };
    
  } catch (error) {
    Logger.log('ドキュメント処理エラー: ' + error.toString());
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * WordファイルをGoogle Documentに変換
 */
function convertWordToGoogleDoc(file) {
  const config = Settings();
  if (!config?.SLACK_TOKEN) throw new Error('Slack Tokenが設定されていません');
  
  // Slackからファイルをダウンロード
  const downloadUrl = file.url_private_download || file.url_private;
  const options = {
    method: 'get',
    headers: {
      'Authorization': 'Bearer ' + config.SLACK_TOKEN
    },
    muteHttpExceptions: true
  };
  
  const response = UrlFetchApp.fetch(downloadUrl, options);
  const blob = response.getBlob();
  
  // Google Documentに変換
  const resource = {
    title: file.name.replace(/\.(docx?|doc)$/i, ''),
    mimeType: MimeType.GOOGLE_DOCS
  };
  
  const convertedFile = Drive.Files.insert(resource, blob);
  
  return {
    docId: convertedFile.id,
    fileName: convertedFile.title
  };
}

/**
 * Google DocumentのURLからドキュメントIDを抽出
 */
function extractDocIdFromUrl(url) {
  const patterns = [
    /\/document\/d\/([a-zA-Z0-9-_]+)/,
    /id=([a-zA-Z0-9-_]+)/,
    /\/([a-zA-Z0-9-_]+)\/edit/
  ];
  
  for (const pattern of patterns) {
    const match = url.match(pattern);
    if (match && match[1]) {
      return match[1];
    }
  }
  
  throw new Error('Google DocumentのIDを抽出できませんでした');
}

/**
 * Google Documentをコピー
 */
function copyGoogleDoc(originalDocId) {
  try {
    const originalDoc = DriveApp.getFileById(originalDocId);
    const copy = originalDoc.makeCopy();
    return copy.getId();
  } catch (error) {
    throw new Error('ドキュメントのコピーに失敗しました: ' + error.toString());
  }
}

/**
 * AIから修正案を取得
 */
function getAISuggestions(documentText, userRequest) {
  const config = Settings();
  
  // プロンプトを構築
  const systemPrompt = `あなたは優秀な文書校正者です。以下の文書を読んで、ユーザーのリクエストに基づいて修正案を提供してください。
修正案は以下のJSON形式で返してください：
[
  {
    "type": "correction" | "comment" | "addition",
    "originalText": "修正対象の元のテキスト（最大50文字）",
    "suggestion": "修正案または追加するテキスト",
    "comment": "修正理由やコメント",
    "position": "before" | "after" | "replace"
  }
]

重要な注意事項：
- originalTextは必ず元の文書に存在する一意のテキストを指定してください
- 同じテキストが複数ある場合は、より長い文脈を含めて一意になるようにしてください
- commentには修正の理由を簡潔に記載してください
- 修正案は具体的で実装可能なものにしてください`;

  const userPrompt = `以下の文書を確認して、「${userRequest || '全般的な改善提案をお願いします'}」という観点で修正案を提供してください。

文書内容：
${documentText}`;

  let suggestions = [];
  
  if (config?.OPEN_AI_TOKEN) {
    const response = chatGPTResponse(userPrompt, {
      optionRole: [{ role: 'system', content: systemPrompt }]
    });
    
    try {
      // レスポンスからJSON部分を抽出
      const jsonMatch = response.match(/\[[\s\S]*\]/);
      if (jsonMatch) {
        suggestions = JSON.parse(jsonMatch[0]);
      }
    } catch (e) {
      Logger.log('JSON解析エラー: ' + e.toString());
      // フォールバック: 基本的な修正案を作成
      suggestions = [{
        type: 'comment',
        originalText: documentText.substring(0, 50),
        suggestion: '',
        comment: 'AIによる詳細な分析が必要です: ' + response.substring(0, 200),
        position: 'after'
      }];
    }
  }
  
  return suggestions;
}

/**
 * AIの修正案をドキュメントに適用
 */
function applyAISuggestionsToDocument(doc, suggestions) {
  const body = doc.getBody();
  
  // 修正案を適用（逆順で処理して位置ずれを防ぐ）
  suggestions.reverse().forEach(suggestion => {
    try {
      const searchResult = body.findText(suggestion.originalText);
      
      if (searchResult) {
        const element = searchResult.getElement();
        const text = element.asText();
        const startOffset = searchResult.getStartOffset();
        const endOffset = searchResult.getEndOffsetInclusive();
        
        // 修正タイプに応じて処理
        switch (suggestion.type) {
          case 'correction':
            // 取り消し線を追加して修正案を併記
            text.setStrikethrough(startOffset, endOffset, true);
            text.setForegroundColor(startOffset, endOffset, '#FF0000');
            
            // 修正案を追加
            const correctionText = ` [修正案: ${suggestion.suggestion}]`;
            text.insertText(endOffset + 1, correctionText);
            text.setForegroundColor(endOffset + 1, endOffset + correctionText.length, '#0000FF');
            text.setBold(endOffset + 1, endOffset + correctionText.length, true);
            
            // コメントを追加
            if (suggestion.comment) {
              const commentText = ` (理由: ${suggestion.comment})`;
              text.insertText(endOffset + correctionText.length + 1, commentText);
              text.setForegroundColor(endOffset + correctionText.length + 1, 
                                     endOffset + correctionText.length + commentText.length, '#666666');
              text.setItalic(endOffset + correctionText.length + 1, 
                            endOffset + correctionText.length + commentText.length, true);
            }
            break;
            
          case 'comment':
            // コメントを黄色ハイライトで追加
            text.setBackgroundColor(startOffset, endOffset, '#FFFF00');
            const commentOnlyText = ` [コメント: ${suggestion.comment}]`;
            text.insertText(endOffset + 1, commentOnlyText);
            text.setForegroundColor(endOffset + 1, endOffset + commentOnlyText.length, '#008000');
            text.setItalic(endOffset + 1, endOffset + commentOnlyText.length, true);
            break;
            
          case 'addition':
            // 追加提案を緑色で表示
            const additionText = ` [追加提案: ${suggestion.suggestion}]`;
            text.insertText(endOffset + 1, additionText);
            text.setForegroundColor(endOffset + 1, endOffset + additionText.length, '#008000');
            text.setBold(endOffset + 1, endOffset + additionText.length, true);
            text.setBackgroundColor(endOffset + 1, endOffset + additionText.length, '#E8F5E9');
            break;
        }
      } else {
        Logger.log(`テキストが見つかりません: ${suggestion.originalText}`);
      }
    } catch (e) {
      Logger.log(`修正案の適用エラー: ${e.toString()}`);
    }
  });
}

/**
 * ドキュメントにレビューヘッダーを追加
 */
function addReviewHeader(doc, originalFileName, userRequest) {
  const body = doc.getBody();
  const timestamp = Utilities.formatDate(new Date(), 'JST', 'yyyy/MM/dd HH:mm:ss');
  
  // ヘッダーセクションを作成
  const headerParagraph = body.insertParagraph(0, '━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  headerParagraph.setHeading(DocumentApp.ParagraphHeading.NORMAL);
  
  const titleParagraph = body.insertParagraph(1, '📝 AI文書レビュー結果');
  titleParagraph.setHeading(DocumentApp.ParagraphHeading.HEADING1);
  titleParagraph.setBold(true);
  
  body.insertParagraph(2, `元ファイル: ${originalFileName}`);
  body.insertParagraph(3, `レビュー日時: ${timestamp}`);
  body.insertParagraph(4, `レビュー要求: ${userRequest || '全般的な改善'}`);
  
  const legendParagraph = body.insertParagraph(5, '\n凡例:');
  legendParagraph.setBold(true);
  
  body.insertParagraph(6, '• 赤字取り消し線: 修正が必要な箇所');
  body.insertParagraph(7, '• 青字: 修正案');
  body.insertParagraph(8, '• 黄色ハイライト: コメント箇所');
  body.insertParagraph(9, '• 緑字: 追加提案');
  
  body.insertParagraph(10, '━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n');
}

/**
 * ドキュメントの共有設定（元のドキュメントと同じ設定をコピー）
 */
function setDocumentSharing(docId, originalDocId) {
  try {
    const file = DriveApp.getFileById(docId);
    
    // 元のドキュメントが存在する場合は、その共有設定をコピー
    if (originalDocId) {
      try {
        const originalFile = DriveApp.getFileById(originalDocId);
        
        // 元のファイルの共有設定を取得
        const originalAccess = originalFile.getSharingAccess();
        const originalPermission = originalFile.getSharingPermission();
        
        // 同じ共有設定を適用
        file.setSharing(originalAccess, originalPermission);
        
        // 元のファイルのエディターをコピー
        const editors = originalFile.getEditors();
        editors.forEach(editor => {
          try {
            file.addEditor(editor.getEmail());
          } catch (e) {
            Logger.log(`エディター追加エラー: ${editor.getEmail()}`);
          }
        });
        
        // 元のファイルのビューアーをコピー
        const viewers = originalFile.getViewers();
        viewers.forEach(viewer => {
          try {
            file.addViewer(viewer.getEmail());
          } catch (e) {
            Logger.log(`ビューアー追加エラー: ${viewer.getEmail()}`);
          }
        });
        
        Logger.log(`元のドキュメントの共有設定をコピーしました: ${docId}`);
      } catch (e) {
        Logger.log('元のドキュメントの共有設定取得エラー: ' + e.toString());
        // フォールバック: デフォルトの共有設定を適用
        setDefaultSharing(file);
      }
    } else {
      // 元のドキュメントがない場合（Wordファイルの場合など）はデフォルト設定
      setDefaultSharing(file);
    }
    
  } catch (error) {
    Logger.log('共有設定エラー: ' + error.toString());
    throw new Error('ドキュメントの共有設定に失敗しました');
  }
}

/**
 * デフォルトの共有設定を適用
 */
function setDefaultSharing(file) {
  // デフォルト: 組織内のユーザーがリンクで編集可能
  try {
    // まず組織内での共有を試みる
    file.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.EDIT);
    Logger.log('組織内共有設定を適用しました');
  } catch (e) {
    // 組織設定が使えない場合は、リンクを知っている人が編集可能に
    try {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.EDIT);
      Logger.log('リンク共有設定を適用しました');
    } catch (e2) {
      Logger.log('共有設定の適用に失敗しました: ' + e2.toString());
    }
  }
}

/**
 * Slackメッセージ用のフォーマット
 */
function formatDocumentReviewResult(result) {
  if (!result.success) {
    return `❌ ドキュメントの処理中にエラーが発生しました:\n${result.error}`;
  }
  
  const message = `✅ **ドキュメントレビューが完了しました！**

📄 **元ファイル:** ${result.originalFileName}
📝 **修正版ファイル:** ${result.fileName}
💡 **修正提案数:** ${result.suggestionsCount}件

🔗 **編集可能なリンク:**
${result.url}

💬 このリンクから直接ドキュメントを編集できます。
   元のドキュメントと同じ共有設定が適用されています。

📌 修正案の見方:
• 赤字取り消し線: 修正が必要な箇所
• 青字: 修正案
• 黄色ハイライト: コメント箇所
• 緑字: 追加提案`;
  
  return message;
}

// ===========================
// AI レスポンス機能
// ===========================

/**
 * OpenAI Chat Completion 呼び出し（改良版）
 */
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

/**
 * Gemini API レスポンス (実装が必要な場合は追加)
 */
function geminiResponse(text, optionRole) {
  // Gemini API の実装をここに追加
  // 現在のコードには実装がないため、空の関数として定義
  return '';
}

// ===========================
// イベントハンドラ
// ===========================

/**
 * スプレッドシート編集時のトリガー
 */
function onEditTrigger(e) {
  //SheetLog.log("onEditTrigger");
  var sheet = e.source.getActiveSheet();
  var range = e.range;
  if(sheet.getName() === 'faq'){
    //SheetLog.log("onEditTrigger: row="+range.getRow()+" col="+range.getColumn());
    // チェックボックスC列
    if (range.getColumn() === 3){
      if(range.getValue() === true && sheet.getRange(range.getRow(),1)) {
        var keyword = sheet.getRange(range.getRow(), 1).getValue(); // キーワード取得
        var result = searchDriveLink(keyword,sheet.getRange(range.getRow(), 4));
        //SheetLog.log("onEditTrigger:"+JSON.stringify(result));
        range.setValue(false);
      }
    }
  }
}

// ===========================
// メインエントリポイント
// ===========================

/**
 * メインエントリポイント
 */
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
        // ファイルが添付されている場合の処理
        let fileContext = null;
        let documentReviewResult = null;
    
    // ドキュメントレビューのキーワードをチェック
    const isDocumentReview = message && (
      message.includes('修正') || 
      message.includes('レビュー') || 
      message.includes('校正') ||
      message.includes('チェック') ||
      message.includes('改善') ||
      message.includes('コメント')
    );
    
    if (event.files && event.files.length > 0) {
      const file = event.files[0]; // 最初のファイルを処理
      
      // Word/Googleドキュメントの場合でレビューが要求されている場合
      if (isDocumentReview && (
        file.mimetype === 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' ||
        file.mimetype === 'application/msword' ||
        (file.url_private && file.url_private.includes('docs.google.com'))
      )) {
        // ドキュメントレビューモードで処理
        documentReviewResult = processDocumentForReview(file, message, channel);
        
        if (documentReviewResult && documentReviewResult.success) {
          // レビュー結果をSlackに送信
          const reviewMessage = formatDocumentReviewResult(documentReviewResult);
          postMessage(reviewMessage, channel, threadTs || event.ts);
          return; // ここで処理終了
        }
      } else {
        // 通常のファイル読み取り処理
        fileContext = handleMessageWithFiles(event);
      }
    }
    
    // チャンネル説明を取得
    const channelInfo = getChannelInfo(channel);
    const channelDescription =
      channelInfo?.purpose?.value ||
      'チャンネル説明がありません。以下の標準プロンプトを使用します。';

    // ベースロール定義
    const baseRole = [
      { role: 'system', content: 'You are a helpful assistant.' },
      {
        role: 'system',
        content:
          "When asked about the latest, be sure to give the Perplexity search URL link of 'https://www.perplexity.ai/?q='",
      },
      {
        role: 'system',
        content: 'When creating graphs, diagrams, or charts, use Mermaid syntax wrapped in ```mermaid code blocks. Slack will automatically render these as visual diagrams. Example formats: flowchart, sequence diagram, gantt chart, pie chart, etc.'
      },
      {
        role: 'system',
        content:
          'If you have any questions, please ask clearly. Also, never answer by guessing.',
      },
      { role: 'system', content: 'Please answer as concisely as possible.' },
      {
        role: 'system',
        content:
          'If you have a reference source, please list the URL in list form at the end.',
      },
    ];

    // メンション部分を除去して本文を取得
    let text = message;
    const mentionStart = text.indexOf('<@');
    const mentionEnd = text.indexOf('>');
    if (mentionStart === 0 && mentionEnd !== -1) {
      text = text.substring(mentionEnd + 1).trim();
    }

    // 簡易コマンド対応
    if (text === 'Hello') {
      postMessage('Hi there!', channel, event.ts);
      return;
    } else if (text === 'help') {
      postMessage(
        'このボットはSlackでChatGPTに質問を投げられるボットです！',
        channel,
        event.ts
      );
      return;
    }

    // system プロンプトを追加
    const optionRole = [...baseRole];
    optionRole.push(
      {
        role: 'system',
        content: 'Please explain the response results in Japanese.',
      },
      {
        role: 'system',
        content:
          '今から説明するslackチャンネルとしてふさわしい回答を望む。付与する情報を前提として回答してください。FAQのスプレッドシートでの処理が可能な場合にはそのFAQの内容を踏まえて回答しつつ、スプレッドシートのリンク(https://docs.google.com/spreadsheets/d/1MKMjUp2F3r71-VCsT4wVfZo1G6IjEpxobQ7fJKtRPlA/edit?usp=sharing)を掲載して。もしもFAQのシートを参照する必要がなければ、FAQのシートには言及しないで。またFAQのスプレッドシートや参考情報のURLはなるべく自然な形で伝えるようにして。' +
          channelDescription,
      }
    );

    // FAQから追加ロールを取得
    const faq = getFaqRole(text);
    if (faq) optionRole.push(faq);

    // ファイルコンテキストがある場合は追加
    if (fileContext) {
      optionRole.push({
        role: 'system',
        content: `ユーザーが添付したファイルの内容:\n${fileContext.fileContents}\n\nこの内容を考慮して回答してください。`
      });
      // ユーザーメッセージにファイル情報を追加
      text = `${text || ''}\n\n[添付ファイル: ${fileContext.fileInfo.map(f => f.name).join(', ')}]`;
    }

    // スレッド履歴をマージ
    const threadMessages = event.thread_ts
      ? getThreadMessages(channel, event.thread_ts)
      : [];
    mergeRoleAndThread(optionRole, threadMessages);

    // AI レスポンス取得
    let responseText = '';
    if (Settings().GEMINI_TOKEN) {
      responseText = geminiResponse(text, optionRole);
    } else {
      responseText = chatGPTResponse(text, { optionRole });
    }

    // スレッドIDを決定（既存スレッド or 新規スレッドを起こす）
    const replyThread = threadTs || event.ts;

        // メッセージ送信（スレッド内）
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
// 設定方法の説明
// ===========================

/**
 * テスト関数
 */

/**
 * シンプルなスプレッドシート作成テスト
 */
function testCreateSpreadsheet() {
  console.log('=== スプレッドシート作成テスト ===');
  
  try {
    // シンプルにスプレッドシートを作成
    console.log('1. スプレッドシートを作成中...');
    const testName = 'Test Spreadsheet ' + new Date().getTime();
    const ss = SpreadsheetApp.create(testName);
    
    if (ss) {
      console.log('✅ 作成成功');
      console.log('   Name: ' + ss.getName());
      console.log('   ID: ' + ss.getId());
      console.log('   URL: ' + ss.getUrl());
      
      // シートの確認
      console.log('\n2. シートを確認中...');
      const sheets = ss.getSheets();
      console.log('   シート数: ' + sheets.length);
      
      if (sheets.length > 0) {
        console.log('   デフォルトシート名: ' + sheets[0].getName());
        
        // シート名を変更してみる
        sheets[0].setName('test_sheet');
        console.log('   シート名変更後: ' + sheets[0].getName());
      }
      
      // データを追加してみる
      console.log('\n3. データを追加中...');
      sheets[0].appendRow(['Test', 'Data', new Date()]);
      console.log('✅ データ追加成功');
      
      // クリーンアップ（テストファイルを削除）
      console.log('\n4. テストファイルを削除中...');
      DriveApp.getFileById(ss.getId()).setTrashed(true);
      console.log('✅ 削除完了');
      
      console.log('\nテスト完了: スプレッドシートの作成と操作が正常に動作します');
      return true;
    } else {
      console.log('❌ スプレッドシートが作成されませんでした');
      return false;
    }
    
  } catch (e) {
    console.log('❌ エラー: ' + e.toString());
    console.log('\nエラーの詳細:');
    console.log(e.stack);
    return false;
  }
}

/**
 * スプレッドシートのアクセス権限テスト
 */
function testSpreadsheetAccess() {
  console.log('=== スプレッドシートアクセステスト ===');
  
  try {
    // Driveへのアクセステスト
    console.log('1. Google Driveへのアクセスを確認中...');
    const files = DriveApp.getFilesByName('test');
    console.log('✅ DriveアクセスOK');
    
    // スプレッドシートサービスの確認
    console.log('\n2. SpreadsheetAppの動作を確認中...');
    const testSS = SpreadsheetApp.create('Access Test ' + new Date().getTime());
    
    if (testSS) {
      console.log('✅ SpreadsheetApp.create() 動作確認');
      console.log('   Type: ' + typeof testSS);
      console.log('   Class: ' + testSS.toString());
      
      // メソッドの存在確認
      console.log('\n3. メソッドの存在を確認中...');
      console.log('   getId: ' + (typeof testSS.getId === 'function' ? '✅' : '❌'));
      console.log('   getName: ' + (typeof testSS.getName === 'function' ? '✅' : '❌'));
      console.log('   getUrl: ' + (typeof testSS.getUrl === 'function' ? '✅' : '❌'));
      console.log('   getSheets: ' + (typeof testSS.getSheets === 'function' ? '✅' : '❌'));
      
      // クリーンアップ
      DriveApp.getFileById(testSS.getId()).setTrashed(true);
      console.log('\n✅ テスト完了: すべてのアクセス権限が正常です');
      return true;
    } else {
      console.log('❌ SpreadsheetApp.create() がnullまたはundefinedを返しました');
      return false;
    }
    
  } catch (e) {
    console.log('❌ エラー: ' + e.toString());
    console.log('\nエラーの詳細:');
    console.log(e.stack);
    console.log('\n可能な原因:');
    console.log('- Google Drive APIが無効');
    console.log('- スプレッドシート作成の権限が不足');
    console.log('- プロジェクトの設定に問題');
    return false;
  }
}

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
  try {
    const config = Settings();
    if (!config.SLACK_TOKEN) {
      console.log('SLACK_TOKENが設定されていません');
      return false;
    }
    
    const url = 'https://slack.com/api/auth.test';
    
    const options = {
      method: 'post',
      payload: {
        token: config.SLACK_TOKEN
      },
      muteHttpExceptions: true
    };
    
    const response = UrlFetchApp.fetch(url, options);
    const data = JSON.parse(response.getContentText());
    
    if (data.ok) {
      console.log('  ユーザー: ' + data.user);
      console.log('  チーム: ' + data.team);
      return true;
    } else {
      console.log('  エラー: ' + data.error);
      return false;
    }
  } catch (e) {
    console.log('  テストエラー: ' + e.toString());
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
// 設定方法の説明
// ===========================

/**
 * 初期設定関数（初回実行時に使用）
 */
function initializeBot() {
  console.log('========================================');
  console.log('Slack Bot 初期設定を開始します...');
  console.log('========================================');
  
  // まずスプレッドシートのアクセスをテスト
  console.log('\n0. システムチェック中...');
  const accessTest = testSpreadsheetAccess();
  if (!accessTest) {
    console.log('❌ システムチェック失敗');
    console.log('Google Drive APIや権限設定を確認してください');
    return null;
  }
  
  let ss = null;
  
  try {
    // スプレッドシートの作成
    console.log('\n1. スプレッドシートを確認中...');
    ss = getOrCreateSpreadsheet();
    
    if (ss) {
      console.log('✅ スプレッドシート準備完了');
      console.log('   URL: ' + ss.getUrl());
    } else {
      console.log('❌ スプレッドシートの作成に失敗しました');
      return null;
    }
  } catch (e) {
    console.log('❌ スプレッドシートの作成中にエラー: ' + e.toString());
    return null;
  }
  
  // 必須プロパティの確認
  console.log('\n2. スクリプトプロパティを確認中...');
  const config = PropertiesService.getScriptProperties().getProperties();
  const required = ['SLACK_TOKEN', 'OPEN_AI_TOKEN'];
  const missing = required.filter(key => !config[key]);
  
  if (missing.length > 0) {
    console.log('⚠️  必要なプロパティが設定されていません: ' + missing.join(', '));
    console.log('\n以下のプロパティを設定してください:');
    console.log('プロジェクトの設定 → スクリプト プロパティ');
    missing.forEach(prop => {
      console.log('  - ' + prop);
    });
  } else {
    console.log('✅ すべての必須プロパティが設定されています');
  }
  
  // オプションプロパティの確認
  const optional = ['GEMINI_TOKEN', 'GOOGLE_NL_API'];
  const missingOptional = optional.filter(key => !config[key]);
  if (missingOptional.length > 0) {
    console.log('\nオプションプロパティ（未設定）:');
    missingOptional.forEach(prop => {
      console.log('  - ' + prop);
    });
  }
  
  // テスト接続
  if (config.SLACK_TOKEN) {
    console.log('\n3. Slack接続をテスト中...');
    try {
      const testResult = testSlackConnection();
      if (testResult) {
        console.log('✅ Slack接続成功');
      } else {
        console.log('❌ Slack接続失敗');
      }
    } catch (e) {
      console.log('❌ Slack接続テスト中にエラー: ' + e.toString());
    }
  }
  
  console.log('\n========================================');
  console.log('初期設定完了！');
  console.log('========================================');
  
  if (ss) {
    console.log('\n📄 スプレッドシートURL:');
    console.log(ss.getUrl());
  }
  
  console.log('\n📝 次のステップ:');
  console.log('1. デプロイ → 新しいデプロイ → ウェブアプリ');
  console.log('2. Web App URLをコピー');
  console.log('3. Slack AppのEvent SubscriptionsにURLを設定');
  console.log('4. SlackチャンネルでBotをテスト');
  
  return {
    spreadsheetUrl: ss ? ss.getUrl() : null,
    spreadsheetId: ss ? ss.getId() : null,
    hasRequiredProps: missing.length === 0
  };
}

/**
 * Slack Bot作成
 * https://api.slack.com/apps
 * 
 * Botの作成手順
 * https://blog.da-vinci-studio.com/entry/2022/09/13/101530
 * 
 * Bot更新時に必要な処理（デプロイ更新）
 * https://ryjkmr.com/gas-web-app-deploy-new-same-url/
 * 
 * 利用開始までの流れ
 * 1. SlackにBotを設定したいチャンネルを作成しchannel_idをコピー
 * 2. channel_idをコードに貼り付け
 * 3. OpenAI API Keyを発行しプロジェクト設定のスクリプトプロパティに入力
 * 4. GASのデプロイでウェブアプリを作成し、自分で実行、全員が利用可能で共有し発行されたウェブアプリURLをコピー
 * 5. SlackAppを以下のManifestにウェブアプリURLを貼り付けしSlackBotを作成
 * 6. SlackAppのOAuthのBot Tokenをコピー
 * 7. Bot Tokenをプロジェクト設定のスクリプトプロパティに入力
 * 8. 作成したBotをSlackにインストール
 * 9. チャンネルにBotを追加
 * 10. Botの動作確認を行う
 * 
 * 必要なAPIの有効化:
 * - Drive API (v2) - GASエディタで「サービス」→「+」から追加
 * 
 * Slack Manifest (JSON形式):
 */

const SLACK_MANIFEST_JSON = {
  "display_information": {
    "name": "ChatGPT",
    "description": "AI Assistant Bot with Document Review",
    "background_color": "#2eb886"
  },
  "features": {
    "bot_user": {
      "display_name": "ChatGPT",
      "always_online": false
    }
  },
  "oauth_config": {
    "scopes": {
      "bot": [
        "app_mentions:read",
        "channels:history",
        "channels:read",
        "chat:write",
        "chat:write.public",
        "files:read",
        "files:write",
        "groups:read",
        "reactions:read",
        "remote_files:read",
        "remote_files:share",
        "users:read"
      ]
    }
  },
  "settings": {
    "event_subscriptions": {
      "request_url": "{GAS_WEB_APP_URL}",  // ここにGASのウェブアプリURLを設定
      "bot_events": [
        "message.channels",
        "app_mention",
        "file_shared",
        "file_change",
        "file_deleted"
      ]
    },
    "org_deploy_enabled": false,
    "socket_mode_enabled": false,
    "token_rotation_enabled": false
  }
};

/**
 * Slack Manifestを取得する関数
 * @param {string} webAppUrl - GASのウェブアプリURL
 * @returns {object} Slack Manifest JSON
 */
function getSlackManifest(webAppUrl) {
  const manifest = JSON.parse(JSON.stringify(SLACK_MANIFEST_JSON));
  manifest.settings.event_subscriptions.request_url = webAppUrl;
  return manifest;
}

/**
 * Slack ManifestをYAML形式で出力する関数
 * SlackアプリのManifest EditorにYAML形式で貼り付ける場合に使用
 */
function getSlackManifestYAML(webAppUrl) {
  const yaml = `display_information:
  name: ChatGPT
  description: AI Assistant Bot with Document Review
  background_color: "#2eb886"
features:
  bot_user:
    display_name: ChatGPT
    always_online: false
oauth_config:
  scopes:
    bot:
      - app_mentions:read
      - channels:history
      - channels:read
      - chat:write
      - chat:write.public
      - files:read
      - files:write
      - groups:read
      - reactions:read
      - remote_files:read
      - remote_files:share
      - users:read
settings:
  event_subscriptions:
    request_url: ${webAppUrl}
    bot_events:
      - message.channels
      - app_mention
      - file_shared
      - file_change
      - file_deleted
  org_deploy_enabled: false
  socket_mode_enabled: false
  token_rotation_enabled: false`;
  
  return yaml;
}

/**
 * 必要なスクリプトプロパティ:
 * - SLACK_TOKEN: Slack Bot Token
 * - OPEN_AI_TOKEN: OpenAI API Key
 * - GEMINI_TOKEN: Gemini API Key (オプション)
 * - GOOGLE_NL_API: Google Natural Language API Key
 */