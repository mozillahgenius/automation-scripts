/**
 * Slack AI Bot - 統合版
 * 
 * このファイルは以下のモジュールを統合したものです：
 * - 設定管理とユーティリティ
 * - Slack Bot基本クラス
 * - メインエントリポイント
 * - FAQ検索とDrive検索機能
 * - 自然言語処理
 * - 文字列処理ユーティリティ
 * - ログ機能
 */

// ===========================
// 設定管理とユーティリティ
// ===========================

/**
 * 環境変数取得
 */
function Settings() {
  const env = ScriptProperties.getProperties();
  return {
    ...env,
  };
}

/**
 * ログ出力
 */
function Log(title, text) {
  Logger.log(title, text);
}

// ===========================
// ログシート管理
// ===========================

var SheetLog = {
  log: function(message) {
    const as = SpreadsheetApp.getActiveSpreadsheet();
    try {
      const logSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('log');
      var now = new Date();    
      logSheet.appendRow([now, message]);
    } catch(e) {
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
 * メッセージ送信
 */
function postMessage(message, channel, threadTs = null) {
  const url = 'https://slack.com/api/chat.postMessage';
  const config = Settings();
  if (!config?.SLACK_TOKEN) return;
  const payload = {
    token: config.SLACK_TOKEN,
    channel: channel,
    unfurl_links: true,
    text: message,
    ...(threadTs ? { thread_ts: threadTs } : {}),
  };
  UrlFetchApp.fetch(url, { method: 'post', payload });
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

    const faqs = SpreadsheetApp.getActive()
      .getSheetByName('faq')
      .getRange('A:B')
      .getValues()
      .filter((row) => !row.every((cell) => cell.trim() === ''));
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
  const ss             = SpreadsheetApp.getActive();
  const faqSheet       = ss.getSheetByName('faq');
  const driveListSheet = ss.getSheetByName('ドライブ一覧');
  const lastFaqRow     = faqSheet.getLastRow();
  const lastDriveRow   = driveListSheet.getLastRow();
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
// AI レスポンス機能
// ===========================

/**
 * OpenAI Chat Completion 呼び出し
 */
function chatGPTResponse(message, { optionRole = [], temperature }) {
  const config = Settings();
  if (!config?.OPEN_AI_TOKEN) return '';
  const apiKey = config.OPEN_AI_TOKEN;
  const url = 'https://api.openai.com/v1/chat/completions';
  const payload = {
    model: 'o3',
    messages: [...optionRole, { role: 'user', content: message }],
    temperature: temperature ?? 1,
  };
  const options = {
    method: 'post',
    muteHttpExceptions: true,
    headers: {
      Authorization: 'Bearer ' + apiKey,
      'Content-Type': 'application/json',
    },
    payload: JSON.stringify(payload),
  };
  const res = UrlFetchApp.fetch(url, options);
  const result = JSON.parse(res.getContentText());
  return result?.choices?.[0]?.message?.content || '';
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
function doPost(e) {
  const params = JSON.parse(e.postData.contents);

  // 重複受信防止（3秒超過リトライ対策）
  const cache = CacheService.getScriptCache();
  if (
    params.hasOwnProperty('event') &&
    params.event.hasOwnProperty('client_msg_id')
  ) {
    if (cache.get(params.event.client_msg_id) === 'done') {
      return ContentService.createTextOutput();
    } else {
      cache.put(params.event.client_msg_id, 'done', 600);
    }
  }

  const Bot = new SlackBot(e);

  // メンションイベント用ハンドラ
  const run = ({ event, message, channel, threadTs }) => {
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
      { role: 'system', content: 'Graphs should be output utilizing quickchart.io' },
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
    postMessage(responseText, channel, replyThread);
  };

  // メンションされたときのみ処理
  Bot.handleMentionEventBase(run);

  return Bot.response();
}

// ===========================
// 設定方法の説明
// ===========================

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
 * Slack Manifest:
 * 
 * display_information:
 *   name: ChatGPT
 * features:
 *   bot_user:
 *     display_name: ChatGPT
 *     always_online: false
 * oauth_config:
 *   scopes:
 *     bot:
 *       - app_mentions:read
 *       - channels:history
 *       - chat:write
 *       - chat:write.public
 *       - groups:read
 *       - reactions:read
 *       - users:read
 *       - channels:read
 * settings:
 *   event_subscriptions:
 *     request_url: {GAS ウェブアプリURL}
 *     bot_events:
 *       - message.channels
 *   org_deploy_enabled: false
 *   socket_mode_enabled: false
 *   token_rotation_enabled: false
 * 
 * 必要なスクリプトプロパティ:
 * - SLACK_TOKEN: Slack Bot Token
 * - OPEN_AI_TOKEN: OpenAI API Key
 * - GEMINI_TOKEN: Gemini API Key (オプション)
 * - GOOGLE_NL_API: Google Natural Language API Key
 */