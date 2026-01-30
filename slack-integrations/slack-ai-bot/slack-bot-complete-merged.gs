/**
 * Slack AI Bot - 完全統合版
 * Version: 2.0
 * 
 * このファイルにはすべての機能が統合されています：
 * - 初期セットアップ機能
 * - Slack Bot本体
 * - ファイル処理（PDF、Word、Googleドキュメント）
 * - ドキュメント編集機能
 * - FAQ検索
 * - Drive検索
 * - デバッグ機能
 */

// =====================================
// セクション1: 初期セットアップ機能
// =====================================

/**
 * 【最初に実行】ステップ1: スプレッドシートIDを手動で設定
 * 
 * 1. Google Driveで新しいスプレッドシートを作成
 * 2. URLからIDをコピー（/d/と/editの間の文字列）
 * 3. 下記のSPREADSHEET_IDに貼り付け
 * 4. この関数を実行
 */
function setupStep1_SetSpreadsheetId() {
  // ★★★ ここにスプレッドシートIDを入力してください ★★★
  const SPREADSHEET_ID = 'YOUR_SPREADSHEET_ID_HERE';
  
  if (SPREADSHEET_ID === 'YOUR_SPREADSHEET_ID_HERE') {
    console.log('❌ エラー: SPREADSHEET_IDを設定してください');
    console.log('\n手順:');
    console.log('1. Google Driveで新しいスプレッドシートを作成');
    console.log('2. スプレッドシートを開く');
    console.log('3. URLから以下の部分をコピー:');
    console.log('   https://docs.google.com/spreadsheets/d/【ここの部分】/edit');
    console.log('4. コピーしたIDを上記のSPREADSHEET_IDに貼り付け');
    console.log('5. この関数を再度実行');
    return;
  }
  
  PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', SPREADSHEET_ID);
  console.log('✅ スプレッドシートIDを保存しました: ' + SPREADSHEET_ID);
  
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    console.log('✅ スプレッドシートに接続成功');
    console.log('   名前: ' + ss.getName());
    console.log('   URL: ' + ss.getUrl());
    console.log('\n次のステップ: setupStep2_InitializeSheets() を実行してください');
  } catch (e) {
    console.log('❌ スプレッドシートを開けません: ' + e.toString());
    console.log('IDが正しいか、アクセス権限があるか確認してください');
  }
}

/**
 * ステップ2: シートを初期化
 */
function setupStep2_InitializeSheets() {
  const SPREADSHEET_ID = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  
  if (!SPREADSHEET_ID) {
    console.log('❌ エラー: 先にsetupStep1_SetSpreadsheetId()を実行してください');
    return;
  }
  
  try {
    console.log('スプレッドシートを開いています...');
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    console.log('✅ スプレッドシート接続成功: ' + ss.getName());
    
    const sheets = ss.getSheets();
    console.log('\n現在のシート数: ' + sheets.length);
    
    // logシートの作成または確認
    console.log('\n1. logシートを設定中...');
    let logSheet = ss.getSheetByName('log');
    if (!logSheet) {
      if (sheets.length > 0 && sheets[0].getName() === 'シート1') {
        sheets[0].setName('log');
        logSheet = sheets[0];
        console.log('   デフォルトシートをlogに変更');
      } else {
        logSheet = ss.insertSheet('log');
        console.log('   logシートを作成');
      }
    } else {
      console.log('   logシートは既に存在');
    }
    
    if (logSheet.getLastRow() === 0) {
      logSheet.appendRow(['Timestamp', 'Message']);
      logSheet.getRange('1:1').setFontWeight('bold');
      logSheet.setFrozenRows(1);
      console.log('   ヘッダーを追加');
    }
    
    // faqシートの作成
    console.log('\n2. faqシートを設定中...');
    let faqSheet = ss.getSheetByName('faq');
    if (!faqSheet) {
      faqSheet = ss.insertSheet('faq');
      faqSheet.appendRow(['キーワード', '回答', '検索フラグ', 'Drive検索結果']);
      faqSheet.getRange('1:1').setFontWeight('bold');
      faqSheet.setFrozenRows(1);
      faqSheet.setColumnWidth(1, 150);
      faqSheet.setColumnWidth(2, 400);
      faqSheet.setColumnWidth(3, 100);
      faqSheet.setColumnWidth(4, 400);
      console.log('   faqシートを作成');
    } else {
      console.log('   faqシートは既に存在');
    }
    
    // ドライブ一覧シートの作成
    console.log('\n3. ドライブ一覧シートを設定中...');
    let driveSheet = ss.getSheetByName('ドライブ一覧');
    if (!driveSheet) {
      driveSheet = ss.insertSheet('ドライブ一覧');
      driveSheet.appendRow(['フォルダID', 'フォルダ名', '説明']);
      driveSheet.getRange('1:1').setFontWeight('bold');
      driveSheet.setFrozenRows(1);
      driveSheet.setColumnWidth(1, 300);
      driveSheet.setColumnWidth(2, 200);
      driveSheet.setColumnWidth(3, 300);
      console.log('   ドライブ一覧シートを作成');
    } else {
      console.log('   ドライブ一覧シートは既に存在');
    }
    
    // debug_logシートの作成
    console.log('\n4. debug_logシートを設定中...');
    let debugSheet = ss.getSheetByName('debug_log');
    if (!debugSheet) {
      debugSheet = ss.insertSheet('debug_log');
      debugSheet.appendRow(['Timestamp', 'Category', 'Message', 'Data']);
      debugSheet.getRange('1:1').setFontWeight('bold');
      debugSheet.setFrozenRows(1);
      debugSheet.setColumnWidth(1, 150);
      debugSheet.setColumnWidth(2, 100);
      debugSheet.setColumnWidth(3, 300);
      debugSheet.setColumnWidth(4, 400);
      console.log('   debug_logシートを作成');
    } else {
      console.log('   debug_logシートは既に存在');
    }
    
    console.log('\n========================================');
    console.log('✅ スプレッドシートの初期化完了！');
    console.log('========================================');
    console.log('\nスプレッドシートURL:');
    console.log(ss.getUrl());
    console.log('\n次のステップ: setupStep3_SetAPIKeys() を実行してください');
    
  } catch (e) {
    console.log('❌ エラー: ' + e.toString());
    console.log('\nエラーの詳細:');
    console.log(e.stack);
  }
}

/**
 * ステップ3: APIキーを設定
 */
function setupStep3_SetAPIKeys() {
  console.log('========================================');
  console.log('APIキーの設定');
  console.log('========================================');
  
  const props = PropertiesService.getScriptProperties().getProperties();
  
  console.log('\n現在の設定:');
  console.log('✅ SPREADSHEET_ID: ' + (props.SPREADSHEET_ID ? '設定済み' : '未設定'));
  console.log((props.SLACK_TOKEN ? '✅' : '❌') + ' SLACK_TOKEN: ' + (props.SLACK_TOKEN ? '設定済み' : '未設定'));
  console.log((props.OPEN_AI_TOKEN ? '✅' : '❌') + ' OPEN_AI_TOKEN: ' + (props.OPEN_AI_TOKEN ? '設定済み' : '未設定'));
  console.log('   GEMINI_TOKEN: ' + (props.GEMINI_TOKEN ? '設定済み（オプション）' : '未設定（オプション）'));
  console.log('   GOOGLE_NL_API: ' + (props.GOOGLE_NL_API ? '設定済み（オプション）' : '未設定（オプション）'));
  
  if (!props.SLACK_TOKEN || !props.OPEN_AI_TOKEN) {
    console.log('\n⚠️ 必要なAPIキーが設定されていません');
    console.log('\n設定方法:');
    console.log('1. プロジェクトの設定 → スクリプト プロパティ');
    console.log('2. 「プロパティを追加」をクリック');
    console.log('3. 以下のプロパティを追加:');
    console.log('   - SLACK_TOKEN: Slack Bot Token');
    console.log('   - OPEN_AI_TOKEN: OpenAI API Key');
    console.log('4. 保存');
    console.log('5. この関数を再度実行');
  } else {
    console.log('\n✅ 必要なAPIキーはすべて設定されています');
    console.log('\n次のステップ: setupStep4_TestConnection() を実行してください');
  }
}

/**
 * ステップ4: 接続テスト
 */
function setupStep4_TestConnection() {
  console.log('========================================');
  console.log('接続テスト');
  console.log('========================================');
  
  // スプレッドシート接続テスト
  console.log('\n1. スプレッドシート接続テスト...');
  const ssId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (!ssId) {
    console.log('❌ SPREADSHEET_IDが設定されていません');
    return;
  }
  
  try {
    const ss = SpreadsheetApp.openById(ssId);
    console.log('✅ スプレッドシート接続成功');
    console.log('   ' + ss.getUrl());
    
    const debugSheet = ss.getSheetByName('debug_log');
    if (debugSheet) {
      debugSheet.appendRow([new Date(), 'Test', 'Connection test', 'Success']);
      console.log('✅ テストデータ書き込み成功');
    }
  } catch (e) {
    console.log('❌ スプレッドシート接続失敗: ' + e.toString());
    return;
  }
  
  // Slack接続テスト
  console.log('\n2. Slack API接続テスト...');
  const slackToken = PropertiesService.getScriptProperties().getProperty('SLACK_TOKEN');
  if (!slackToken) {
    console.log('⚠️ SLACK_TOKENが設定されていません（スキップ）');
  } else {
    try {
      const url = 'https://slack.com/api/auth.test';
      const response = UrlFetchApp.fetch(url, {
        method: 'post',
        payload: { token: slackToken },
        muteHttpExceptions: true
      });
      const data = JSON.parse(response.getContentText());
      
      if (data.ok) {
        console.log('✅ Slack接続成功');
        console.log('   ユーザー: ' + data.user);
        console.log('   チーム: ' + data.team);
      } else {
        console.log('❌ Slack接続失敗: ' + data.error);
      }
    } catch (e) {
      console.log('❌ Slackテストエラー: ' + e.toString());
    }
  }
  
  // OpenAI接続テスト
  console.log('\n3. OpenAI API接続テスト...');
  const openAIToken = PropertiesService.getScriptProperties().getProperty('OPEN_AI_TOKEN');
  if (!openAIToken) {
    console.log('⚠️ OPEN_AI_TOKENが設定されていません（スキップ）');
  } else {
    console.log('✅ OPEN_AI_TOKEN設定確認（実際の接続テストは省略）');
  }
  
  console.log('\n========================================');
  console.log('セットアップ完了！');
  console.log('========================================');
  console.log('\n次のステップ:');
  console.log('1. デプロイ → 新しいデプロイ');
  console.log('2. 種類: ウェブアプリ');
  console.log('3. Web App URLをコピー');
  console.log('4. Slack Appの設定でURLを登録');
}

/**
 * クイックセットアップ（既存のスプレッドシートがある場合）
 */
function quickSetupWithExistingSpreadsheet(spreadsheetId) {
  console.log('クイックセットアップを開始...\n');
  
  PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', spreadsheetId);
  console.log('✅ スプレッドシートID設定: ' + spreadsheetId);
  
  setupStep2_InitializeSheets();
  setupStep3_SetAPIKeys();
  setupStep4_TestConnection();
}

// =====================================
// セクション2: グローバル設定と定数
// =====================================

const CONFIG = {
  MAX_MESSAGE_LENGTH: 3000,
  OPENAI_MODEL: 'gpt-4o',
  GEMINI_MODEL: 'gemini-1.5-pro',
  DEBUG_MODE: true,
  SPREADSHEET_NAME: 'Slack Bot Data',
  SHEET_NAMES: {
    LOG: 'log',
    FAQ: 'faq',
    DRIVE_LIST: 'ドライブ一覧',
    DEBUG: 'debug_log'
  }
};

// =====================================
// セクション3: ユーティリティ関数
// =====================================

/**
 * デバッグログ出力
 */
function debugLog(category, message, data = null) {
  if (!CONFIG.DEBUG_MODE) return;
  
  try {
    const ss = getActiveSpreadsheet();
    if (!ss) {
      console.log(`Debug: ${category} - ${message}`, data);
      return;
    }
    
    const debugSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.DEBUG);
    if (debugSheet) {
      debugSheet.appendRow([
        new Date(),
        category,
        message,
        data ? JSON.stringify(data) : ''
      ]);
    }
  } catch (e) {
    console.log('Debug log error:', e);
  }
}

/**
 * スプレッドシートを取得（自動作成対応）
 */
function getActiveSpreadsheet() {
  const ssId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  
  if (!ssId) {
    debugLog('Error', 'SPREADSHEET_ID not found in properties');
    throw new Error('SPREADSHEET_ID not configured. Run setupStep1_SetSpreadsheetId() first');
  }
  
  try {
    return SpreadsheetApp.openById(ssId);
  } catch (e) {
    debugLog('Error', 'Failed to open spreadsheet', e.toString());
    throw new Error('Cannot open spreadsheet. Check SPREADSHEET_ID');
  }
}

/**
 * APIキーを取得
 */
function getApiKey(keyName) {
  const key = PropertiesService.getScriptProperties().getProperty(keyName);
  if (!key) {
    throw new Error(`${keyName} not found in Script Properties`);
  }
  return key;
}

/**
 * 文字列処理：コードブロックの除去
 */
function removeCodeBlocks(text) {
  return text.replace(/```[\s\S]*?```/g, '');
}

/**
 * 文字列処理：メッセージの短縮
 */
function truncateMessage(message, maxLength = CONFIG.MAX_MESSAGE_LENGTH) {
  if (message.length <= maxLength) return message;
  return message.substring(0, maxLength - 3) + '...';
}

/**
 * エラーメッセージのフォーマット
 */
function formatError(error) {
  return `エラーが発生しました: ${error.toString()}\n詳細はdebug_logシートを確認してください。`;
}

// =====================================
// セクション4: Slack Bot本体
// =====================================

/**
 * SlackBotクラス
 */
class SlackBot {
  constructor() {
    this.token = getApiKey('SLACK_TOKEN');
    this.apiUrl = 'https://slack.com/api/';
  }
  
  /**
   * メッセージ送信
   */
  sendMessage(channel, text, thread_ts = null) {
    const payload = {
      token: this.token,
      channel: channel,
      text: truncateMessage(text),
      thread_ts: thread_ts
    };
    
    const response = UrlFetchApp.fetch(this.apiUrl + 'chat.postMessage', {
      method: 'post',
      payload: payload,
      muteHttpExceptions: true
    });
    
    const result = JSON.parse(response.getContentText());
    debugLog('Slack', 'Message sent', result);
    
    if (!result.ok) {
      throw new Error('Slack API error: ' + result.error);
    }
    
    return result;
  }
  
  /**
   * ファイル情報取得
   */
  getFileInfo(fileId) {
    const response = UrlFetchApp.fetch(this.apiUrl + 'files.info', {
      method: 'post',
      payload: {
        token: this.token,
        file: fileId
      },
      muteHttpExceptions: true
    });
    
    const result = JSON.parse(response.getContentText());
    if (!result.ok) {
      throw new Error('Failed to get file info: ' + result.error);
    }
    
    return result.file;
  }
  
  /**
   * ファイルダウンロード
   */
  downloadFile(url) {
    const response = UrlFetchApp.fetch(url, {
      headers: {
        'Authorization': 'Bearer ' + this.token
      },
      muteHttpExceptions: true
    });
    
    return response.getBlob();
  }
  
  /**
   * リアクション追加
   */
  addReaction(channel, timestamp, emoji) {
    UrlFetchApp.fetch(this.apiUrl + 'reactions.add', {
      method: 'post',
      payload: {
        token: this.token,
        channel: channel,
        timestamp: timestamp,
        name: emoji
      },
      muteHttpExceptions: true
    });
  }
}

// =====================================
// セクション5: AI API統合
// =====================================

/**
 * OpenAI API呼び出し
 */
function callOpenAI(messages, model = CONFIG.OPENAI_MODEL) {
  const apiKey = getApiKey('OPEN_AI_TOKEN');
  
  const payload = {
    model: model,
    messages: messages,
    temperature: 0.7,
    max_tokens: 2000
  };
  
  const response = UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions', {
    method: 'post',
    headers: {
      'Authorization': 'Bearer ' + apiKey,
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
  
  const result = JSON.parse(response.getContentText());
  
  if (result.error) {
    throw new Error('OpenAI API error: ' + result.error.message);
  }
  
  return result.choices[0].message.content;
}

/**
 * Gemini API呼び出し
 */
function callGemini(prompt) {
  try {
    const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_TOKEN');
    if (!apiKey) return null;
    
    const url = `https://generativelanguage.googleapis.com/v1beta/models/${CONFIG.GEMINI_MODEL}:generateContent?key=${apiKey}`;
    
    const payload = {
      contents: [{
        parts: [{
          text: prompt
        }]
      }]
    };
    
    const response = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    
    const result = JSON.parse(response.getContentText());
    
    if (result.error) {
      debugLog('Gemini', 'API error', result.error);
      return null;
    }
    
    return result.candidates[0].content.parts[0].text;
  } catch (e) {
    debugLog('Gemini', 'Error calling API', e.toString());
    return null;
  }
}

// =====================================
// セクション6: ファイル処理機能
// =====================================

/**
 * ファイルハンドラークラス
 */
class FileHandler {
  constructor(slackBot) {
    this.slackBot = slackBot;
  }
  
  /**
   * ファイル処理のメインメソッド
   */
  async processFile(fileId, channel, thread_ts) {
    try {
      debugLog('FileHandler', 'Processing file', fileId);
      
      // ファイル情報を取得
      const fileInfo = this.slackBot.getFileInfo(fileId);
      const fileName = fileInfo.name;
      const mimeType = fileInfo.mimetype;
      
      debugLog('FileHandler', 'File info', {name: fileName, type: mimeType});
      
      // ファイルタイプによって処理を分岐
      let content = '';
      
      if (mimeType === 'application/pdf') {
        content = await this.processPDF(fileInfo);
      } else if (mimeType.includes('word') || mimeType.includes('document')) {
        content = await this.processWord(fileInfo);
      } else if (fileName.includes('docs.google.com')) {
        content = await this.processGoogleDoc(fileInfo);
      } else {
        content = await this.processTextFile(fileInfo);
      }
      
      debugLog('FileHandler', 'Content extracted', content.substring(0, 100));
      
      return content;
      
    } catch (e) {
      debugLog('FileHandler', 'Error processing file', e.toString());
      throw e;
    }
  }
  
  /**
   * PDF処理
   */
  processPDF(fileInfo) {
    try {
      const blob = this.slackBot.downloadFile(fileInfo.url_private);
      
      // Google Drive APIを使用してOCR
      const driveFile = Drive.Files.insert({
        title: fileInfo.name + '_temp',
        mimeType: 'application/pdf'
      }, blob, {
        ocr: true,
        ocrLanguage: 'ja'
      });
      
      // テキストを抽出
      const doc = DocumentApp.openById(driveFile.id);
      const text = doc.getBody().getText();
      
      // 一時ファイルを削除
      Drive.Files.remove(driveFile.id);
      
      return text;
    } catch (e) {
      debugLog('FileHandler', 'PDF processing error', e.toString());
      throw new Error('PDF処理エラー: ' + e.toString());
    }
  }
  
  /**
   * Word文書処理
   */
  processWord(fileInfo) {
    try {
      const blob = this.slackBot.downloadFile(fileInfo.url_private);
      
      // Google Driveにアップロードして変換
      const driveFile = Drive.Files.insert({
        title: fileInfo.name + '_temp',
        mimeType: 'application/vnd.google-apps.document',
        convert: true
      }, blob);
      
      // Google Documentとして開く
      const doc = DocumentApp.openById(driveFile.id);
      const text = doc.getBody().getText();
      
      // 一時ファイルを削除
      Drive.Files.remove(driveFile.id);
      
      return text;
    } catch (e) {
      debugLog('FileHandler', 'Word processing error', e.toString());
      throw new Error('Word文書処理エラー: ' + e.toString());
    }
  }
  
  /**
   * Google Document処理
   */
  processGoogleDoc(fileInfo) {
    try {
      // URLからドキュメントIDを抽出
      const url = fileInfo.url_private || fileInfo.external_url;
      const docId = this.extractGoogleDocId(url);
      
      if (!docId) {
        throw new Error('Google DocumentのIDを取得できません');
      }
      
      const doc = DocumentApp.openById(docId);
      return doc.getBody().getText();
      
    } catch (e) {
      debugLog('FileHandler', 'Google Doc processing error', e.toString());
      throw new Error('Google Document処理エラー: ' + e.toString());
    }
  }
  
  /**
   * テキストファイル処理
   */
  processTextFile(fileInfo) {
    try {
      const blob = this.slackBot.downloadFile(fileInfo.url_private);
      return blob.getDataAsString();
    } catch (e) {
      debugLog('FileHandler', 'Text file processing error', e.toString());
      throw new Error('テキストファイル処理エラー: ' + e.toString());
    }
  }
  
  /**
   * Google Document IDの抽出
   */
  extractGoogleDocId(url) {
    const patterns = [
      /\/document\/d\/([a-zA-Z0-9-_]+)/,
      /id=([a-zA-Z0-9-_]+)/,
      /\/([a-zA-Z0-9-_]+)$/
    ];
    
    for (const pattern of patterns) {
      const match = url.match(pattern);
      if (match) return match[1];
    }
    
    return null;
  }
}

// =====================================
// セクション7: ドキュメント編集機能
// =====================================

/**
 * ドキュメントエディタークラス
 */
class DocumentEditor {
  constructor() {
    this.apiKey = getApiKey('OPEN_AI_TOKEN');
  }
  
  /**
   * ドキュメントレビューと編集
   */
  reviewAndEditDocument(docId, reviewType = 'general') {
    try {
      debugLog('DocumentEditor', 'Starting review', {docId, reviewType});
      
      // 元のドキュメントを開く
      const originalDoc = DocumentApp.openById(docId);
      const originalBody = originalDoc.getBody();
      const originalText = originalBody.getText();
      
      // 元の共有設定を取得
      const originalFile = DriveApp.getFileById(docId);
      const originalSharing = this.getShareSettings(originalFile);
      
      // ドキュメントのコピーを作成
      const copyName = originalDoc.getName() + ' - レビュー版 ' + new Date().toLocaleString('ja-JP');
      const copyFile = originalFile.makeCopy(copyName);
      const copyDoc = DocumentApp.openById(copyFile.getId());
      const copyBody = copyDoc.getBody();
      
      // AIによるレビューを取得
      const review = this.getAIReview(originalText, reviewType);
      
      // レビュー結果をドキュメントに反映
      this.applyReviewToDocument(copyBody, review);
      
      // 共有設定を復元
      this.restoreShareSettings(copyFile, originalSharing);
      
      debugLog('DocumentEditor', 'Review completed', copyFile.getId());
      
      return {
        url: copyDoc.getUrl(),
        docId: copyFile.getId(),
        name: copyName
      };
      
    } catch (e) {
      debugLog('DocumentEditor', 'Error in review', e.toString());
      throw e;
    }
  }
  
  /**
   * AIレビューの取得
   */
  getAIReview(text, reviewType) {
    const prompts = {
      general: '以下の文章をレビューし、改善点と修正案を提供してください。',
      grammar: '以下の文章の文法エラーを指摘し、修正案を提供してください。',
      clarity: '以下の文章の明確性を改善する提案をしてください。',
      professional: '以下の文章をよりプロフェッショナルにする提案をしてください。'
    };
    
    const prompt = prompts[reviewType] || prompts.general;
    
    const messages = [
      {
        role: 'system',
        content: 'あなたは文書レビューの専門家です。修正が必要な箇所を特定し、具体的な修正案を提供してください。JSON形式で回答してください。'
      },
      {
        role: 'user',
        content: `${prompt}\n\n文章:\n${text}\n\nJSON形式で以下の構造で回答してください:\n{\n  "sections": [\n    {\n      "original": "元の文章の一部",\n      "suggestion": "修正案",\n      "comment": "修正理由"\n    }\n  ],\n  "overall_feedback": "全体的なフィードバック"\n}`
      }
    ];
    
    const response = callOpenAI(messages);
    
    try {
      return JSON.parse(response);
    } catch (e) {
      debugLog('DocumentEditor', 'Failed to parse AI response', response);
      return {
        sections: [],
        overall_feedback: response
      };
    }
  }
  
  /**
   * レビュー結果をドキュメントに適用
   */
  applyReviewToDocument(body, review) {
    // 全体的なフィードバックを先頭に追加
    if (review.overall_feedback) {
      const feedbackPara = body.insertParagraph(0, '【全体的なフィードバック】');
      feedbackPara.setHeading(DocumentApp.ParagraphHeading.HEADING1);
      feedbackPara.setForegroundColor('#1a73e8');
      
      body.insertParagraph(1, review.overall_feedback)
        .setForegroundColor('#1a73e8')
        .setItalic(true);
      
      body.insertParagraph(2, '');
      body.insertHorizontalRule(3);
      body.insertParagraph(4, '');
    }
    
    // 各セクションの修正提案を追加
    if (review.sections && review.sections.length > 0) {
      const text = body.getText();
      
      review.sections.forEach(section => {
        if (!section.original) return;
        
        const searchResult = body.findText(section.original);
        if (searchResult) {
          const element = searchResult.getElement();
          const startOffset = searchResult.getStartOffset();
          
          // 修正案をハイライト付きで挿入
          if (element.getType() === DocumentApp.ElementType.TEXT) {
            const textElement = element.asText();
            
            // 元のテキストに背景色を追加
            textElement.setBackgroundColor(startOffset, searchResult.getEndOffsetInclusive(), '#fff3cd');
            
            // コメントボックス風の修正案を追加
            const parent = element.getParent();
            const parentIndex = body.getChildIndex(parent);
            
            // 修正案のボックスを作成
            const suggestionBox = body.insertParagraph(parentIndex + 1, '');
            suggestionBox.appendText('💡 修正案: ').setBold(true).setForegroundColor('#0d6efd');
            suggestionBox.appendText(section.suggestion).setForegroundColor('#0d6efd');
            
            if (section.comment) {
              suggestionBox.appendText('\n📝 理由: ').setBold(true).setForegroundColor('#6c757d');
              suggestionBox.appendText(section.comment).setForegroundColor('#6c757d').setItalic(true);
            }
            
            // ボックススタイルの設定
            suggestionBox.setIndentFirstLine(20);
            suggestionBox.setLeftIndent(20);
            suggestionBox.setSpacingAfter(10);
          }
        }
      });
    }
  }
  
  /**
   * 共有設定の取得
   */
  getShareSettings(file) {
    try {
      const access = file.getSharingAccess();
      const permission = file.getSharingPermission();
      const editors = file.getEditors().map(user => user.getEmail());
      const viewers = file.getViewers().map(user => user.getEmail());
      
      return {
        access: access,
        permission: permission,
        editors: editors,
        viewers: viewers
      };
    } catch (e) {
      debugLog('DocumentEditor', 'Error getting share settings', e.toString());
      return null;
    }
  }
  
  /**
   * 共有設定の復元
   */
  restoreShareSettings(file, settings) {
    if (!settings) return;
    
    try {
      // アクセスレベルの設定
      file.setSharing(settings.access, settings.permission);
      
      // 編集者の追加
      settings.editors.forEach(email => {
        try {
          file.addEditor(email);
        } catch (e) {
          debugLog('DocumentEditor', 'Failed to add editor', email);
        }
      });
      
      // 閲覧者の追加
      settings.viewers.forEach(email => {
        try {
          file.addViewer(email);
        } catch (e) {
          debugLog('DocumentEditor', 'Failed to add viewer', email);
        }
      });
      
    } catch (e) {
      debugLog('DocumentEditor', 'Error restoring share settings', e.toString());
    }
  }
}

// =====================================
// セクション8: FAQ・Drive検索機能
// =====================================

/**
 * FAQ検索
 */
function searchFAQ(query) {
  try {
    const ss = getActiveSpreadsheet();
    const faqSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.FAQ);
    
    if (!faqSheet || faqSheet.getLastRow() <= 1) {
      return null;
    }
    
    const data = faqSheet.getRange(2, 1, faqSheet.getLastRow() - 1, 4).getValues();
    const lowerQuery = query.toLowerCase();
    
    for (const row of data) {
      const keyword = row[0].toString().toLowerCase();
      if (lowerQuery.includes(keyword)) {
        return {
          answer: row[1],
          searchDrive: row[2] === true || row[2] === 'TRUE',
          driveResults: row[3]
        };
      }
    }
    
    return null;
  } catch (e) {
    debugLog('FAQ', 'Search error', e.toString());
    return null;
  }
}

/**
 * Drive検索
 */
function searchDrive(query) {
  try {
    const ss = getActiveSpreadsheet();
    const driveSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.DRIVE_LIST);
    
    if (!driveSheet || driveSheet.getLastRow() <= 1) {
      return [];
    }
    
    const folderIds = driveSheet.getRange(2, 1, driveSheet.getLastRow() - 1, 1).getValues();
    const results = [];
    
    for (const [folderId] of folderIds) {
      if (!folderId) continue;
      
      try {
        const folder = DriveApp.getFolderById(folderId);
        const files = folder.searchFiles(`fullText contains "${query}"`);
        
        while (files.hasNext()) {
          const file = files.next();
          results.push({
            name: file.getName(),
            url: file.getUrl(),
            lastModified: file.getLastUpdated()
          });
        }
      } catch (e) {
        debugLog('Drive', 'Folder search error', {folderId, error: e.toString()});
      }
    }
    
    return results;
  } catch (e) {
    debugLog('Drive', 'Search error', e.toString());
    return [];
  }
}

// =====================================
// セクション9: Natural Language API
// =====================================

/**
 * Google Natural Language API呼び出し
 */
function analyzeTextWithNL(text) {
  try {
    const apiKey = PropertiesService.getScriptProperties().getProperty('GOOGLE_NL_API');
    if (!apiKey) return null;
    
    const url = `https://language.googleapis.com/v1/documents:analyzeSentiment?key=${apiKey}`;
    
    const payload = {
      document: {
        type: 'PLAIN_TEXT',
        content: text,
        language: 'ja'
      }
    };
    
    const response = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    
    const result = JSON.parse(response.getContentText());
    
    if (result.error) {
      debugLog('NL API', 'Error', result.error);
      return null;
    }
    
    return {
      sentiment: result.documentSentiment,
      language: result.language
    };
    
  } catch (e) {
    debugLog('NL API', 'Error calling API', e.toString());
    return null;
  }
}

// =====================================
// セクション10: メインハンドラー
// =====================================

/**
 * Webhookエンドポイント（SlackからのPOSTリクエスト）
 */
function doPost(e) {
  try {
    debugLog('Main', 'Request received', e.postData.contents);
    
    const params = JSON.parse(e.postData.contents);
    
    // URL Verification
    if (params.type === 'url_verification') {
      return ContentService.createTextOutput(params.challenge);
    }
    
    // Event処理
    if (params.event) {
      handleSlackEvent(params.event);
    }
    
    return ContentService.createTextOutput('OK');
    
  } catch (error) {
    debugLog('Main', 'Error in doPost', error.toString());
    return ContentService.createTextOutput('Error: ' + error.toString());
  }
}

/**
 * Slackイベントハンドラー
 */
function handleSlackEvent(event) {
  try {
    const bot = new SlackBot();
    
    // ボット自身のメッセージは無視
    if (event.bot_id) return;
    
    debugLog('Event', 'Processing', event);
    
    switch (event.type) {
      case 'message':
      case 'app_mention':
        handleMessage(event, bot);
        break;
        
      case 'file_shared':
        handleFileShared(event, bot);
        break;
        
      default:
        debugLog('Event', 'Unknown event type', event.type);
    }
    
  } catch (e) {
    debugLog('Event', 'Handler error', e.toString());
  }
}

/**
 * メッセージハンドラー
 */
function handleMessage(event, bot) {
  try {
    const message = event.text || '';
    const channel = event.channel;
    const thread_ts = event.thread_ts || event.ts;
    
    // メンションを除去
    const cleanMessage = message.replace(/<@[A-Z0-9]+>/g, '').trim();
    
    if (!cleanMessage) return;
    
    // リアクションを追加
    bot.addReaction(channel, event.ts, 'thinking_face');
    
    // FAQ検索
    const faqResult = searchFAQ(cleanMessage);
    
    let response = '';
    
    if (faqResult) {
      response = faqResult.answer;
      
      // Drive検索も実行する場合
      if (faqResult.searchDrive) {
        const driveResults = searchDrive(cleanMessage);
        if (driveResults.length > 0) {
          response += '\n\n📁 関連ファイル:\n';
          driveResults.slice(0, 5).forEach(file => {
            response += `• <${file.url}|${file.name}>\n`;
          });
        }
      }
    } else {
      // AIに質問
      const messages = [
        {
          role: 'system',
          content: 'あなたは親切なアシスタントです。質問に簡潔に答えてください。Mermaid形式でグラフを作成する場合は、```mermaid と ``` で囲んでください。'
        },
        {
          role: 'user',
          content: cleanMessage
        }
      ];
      
      response = callOpenAI(messages);
      
      // Mermaidグラフの処理
      response = processMermaidGraphs(response);
    }
    
    // 返信
    bot.sendMessage(channel, response, thread_ts);
    
    // ログ記録
    logMessage(cleanMessage, response);
    
  } catch (e) {
    debugLog('Message', 'Handler error', e.toString());
    bot.sendMessage(event.channel, formatError(e), event.thread_ts || event.ts);
  }
}

/**
 * ファイル共有ハンドラー
 */
function handleFileShared(event, bot) {
  try {
    const fileHandler = new FileHandler(bot);
    const documentEditor = new DocumentEditor();
    
    const fileId = event.file_id;
    const channel = event.channel_id;
    
    bot.addReaction(channel, event.ts, 'eyes');
    
    // ファイル情報を取得
    const fileInfo = bot.getFileInfo(fileId);
    const fileName = fileInfo.name;
    
    // ファイル内容を処理
    const content = fileHandler.processFile(fileId, channel, event.ts);
    
    // ドキュメントレビューのキーワードチェック
    const reviewKeywords = ['修正', 'レビュー', 'チェック', '確認', '添削'];
    const shouldReview = reviewKeywords.some(keyword => 
      event.text && event.text.includes(keyword)
    );
    
    if (shouldReview && content) {
      // Google Docに変換してレビュー
      const tempDoc = DocumentApp.create('Temp_' + fileName);
      tempDoc.getBody().setText(content);
      const docId = tempDoc.getId();
      
      const reviewResult = documentEditor.reviewAndEditDocument(docId, 'general');
      
      bot.sendMessage(
        channel,
        `📝 ドキュメントをレビューしました\n` +
        `修正版: ${reviewResult.url}\n` +
        `このドキュメントは元の共有設定と同じ権限で共有されています。`,
        event.ts
      );
      
      // 一時ドキュメントを削除
      DriveApp.getFileById(docId).setTrashed(true);
      
    } else if (content) {
      // 内容のサマリーを生成
      const summary = generateSummary(content);
      bot.sendMessage(
        channel,
        `📄 ファイル「${fileName}」を読み取りました\n\n${summary}`,
        event.ts
      );
    }
    
  } catch (e) {
    debugLog('File', 'Handler error', e.toString());
    bot.sendMessage(event.channel_id, formatError(e), event.ts);
  }
}

/**
 * Mermaidグラフの処理
 */
function processMermaidGraphs(text) {
  const mermaidPattern = /```mermaid\n([\s\S]*?)```/g;
  
  return text.replace(mermaidPattern, (match, graphCode) => {
    return `\n[Mermaidグラフ]\n\`\`\`\n${graphCode}\`\`\`\n(Mermaid Live Editorで表示: https://mermaid.live/)`;
  });
}

/**
 * サマリー生成
 */
function generateSummary(text) {
  try {
    const maxLength = 500;
    const truncatedText = text.length > maxLength ? 
      text.substring(0, maxLength) + '...' : text;
    
    const messages = [
      {
        role: 'system',
        content: '以下のテキストを3-5行で要約してください。'
      },
      {
        role: 'user',
        content: truncatedText
      }
    ];
    
    return callOpenAI(messages);
    
  } catch (e) {
    debugLog('Summary', 'Generation error', e.toString());
    return text.substring(0, 200) + '...';
  }
}

/**
 * メッセージログ記録
 */
function logMessage(input, output) {
  try {
    const ss = getActiveSpreadsheet();
    const logSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.LOG);
    
    if (logSheet) {
      logSheet.appendRow([
        new Date(),
        `Q: ${input}\nA: ${output}`
      ]);
    }
  } catch (e) {
    debugLog('Log', 'Error writing log', e.toString());
  }
}

// =====================================
// セクション11: テスト関数
// =====================================

/**
 * 設定テスト
 */
function testSettings() {
  console.log('========================================');
  console.log('設定テスト');
  console.log('========================================\n');
  
  const props = PropertiesService.getScriptProperties().getProperties();
  
  console.log('スクリプトプロパティ:');
  for (const key in props) {
    const value = props[key];
    const display = key.includes('TOKEN') || key.includes('API') ? 
      '***' + value.substring(value.length - 4) : value;
    console.log(`  ${key}: ${display}`);
  }
  
  try {
    const ss = getActiveSpreadsheet();
    console.log('\nスプレッドシート: ✅ 接続成功');
    console.log('  URL: ' + ss.getUrl());
    
    const sheets = ss.getSheets();
    console.log('\nシート一覧:');
    sheets.forEach(sheet => {
      console.log(`  - ${sheet.getName()} (${sheet.getLastRow()} 行)`);
    });
    
  } catch (e) {
    console.log('\nスプレッドシート: ❌ エラー');
    console.log('  ' + e.toString());
  }
}

/**
 * Slack接続テスト
 */
function testSlackConnection() {
  console.log('Slack接続テスト...\n');
  
  try {
    const bot = new SlackBot();
    const url = 'https://slack.com/api/auth.test';
    
    const response = UrlFetchApp.fetch(url, {
      method: 'post',
      payload: { token: bot.token },
      muteHttpExceptions: true
    });
    
    const data = JSON.parse(response.getContentText());
    
    if (data.ok) {
      console.log('✅ Slack接続成功');
      console.log('  Bot: ' + data.user);
      console.log('  Team: ' + data.team);
      console.log('  Team ID: ' + data.team_id);
    } else {
      console.log('❌ Slack接続失敗');
      console.log('  Error: ' + data.error);
    }
    
  } catch (e) {
    console.log('❌ エラー: ' + e.toString());
  }
}

/**
 * OpenAI接続テスト
 */
function testOpenAI() {
  console.log('OpenAI APIテスト...\n');
  
  try {
    const response = callOpenAI([
      { role: 'system', content: 'You are a test bot.' },
      { role: 'user', content: 'Say "Hello, World!" in Japanese.' }
    ]);
    
    console.log('✅ OpenAI API接続成功');
    console.log('  Response: ' + response);
    
  } catch (e) {
    console.log('❌ OpenAI APIエラー: ' + e.toString());
  }
}

/**
 * 最新ログ確認
 */
function checkRecentLogs() {
  try {
    const ss = getActiveSpreadsheet();
    const debugSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.DEBUG);
    
    if (!debugSheet || debugSheet.getLastRow() <= 1) {
      console.log('ログがありません');
      return;
    }
    
    const lastRow = debugSheet.getLastRow();
    const numRows = Math.min(10, lastRow - 1);
    const startRow = Math.max(2, lastRow - numRows + 1);
    
    const logs = debugSheet.getRange(startRow, 1, numRows, 4).getValues();
    
    console.log('最新のログ（最大10件）:\n');
    logs.forEach(log => {
      console.log(`[${log[0]}] ${log[1]}: ${log[2]}`);
      if (log[3]) console.log(`  Data: ${log[3]}`);
    });
    
  } catch (e) {
    console.log('ログ確認エラー: ' + e.toString());
  }
}

// =====================================
// セクション12: 初期化関数（メイン）
// =====================================

/**
 * 完全初期化（新規インストール用）
 */
function initializeBot() {
  console.log('========================================');
  console.log('Slack AI Bot 初期化');
  console.log('========================================\n');
  
  console.log('⚠️ 注意: この関数は自動でスプレッドシートIDを作成しません');
  console.log('手動でのセットアップが必要です。\n');
  console.log('セットアップ手順:');
  console.log('1. setupStep1_SetSpreadsheetId() を実行');
  console.log('2. setupStep2_InitializeSheets() を実行');
  console.log('3. setupStep3_SetAPIKeys() を実行');
  console.log('4. setupStep4_TestConnection() を実行\n');
  
  console.log('詳細な手順はSETUP.mdを参照してください。');
}

/**
 * appsscript.json設定（参考用）
 */
function getAppsScriptJson() {
  return {
    "timeZone": "Asia/Tokyo",
    "dependencies": {
      "enabledAdvancedServices": [
        {
          "userSymbol": "Drive",
          "version": "v2",
          "serviceId": "drive"
        }
      ]
    },
    "exceptionLogging": "STACKDRIVER",
    "runtimeVersion": "V8",
    "webapp": {
      "executeAs": "USER_DEPLOYING",
      "access": "ANYONE_ANONYMOUS"
    }
  };
}

// =====================================
// 完了メッセージ
// =====================================
console.log('Slack AI Bot スクリプトが読み込まれました');
console.log('初めての方は initializeBot() を実行してセットアップ手順を確認してください');