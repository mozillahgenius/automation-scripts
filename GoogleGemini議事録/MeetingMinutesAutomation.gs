/**
 * 会議議事録自動化・ナレッジ資産化システム
 * Google Apps Script
 *
 * 作成日: 2026-01-19
 * 更新日: 2026-01-19
 * 作成者: Hodaka / IntelligentBeast
 *
 * 【複数ユーザー対応版】
 * - 共有フォルダ方式で全ユーザーのトランスクリプトを処理
 * - スプレッドシートで設定・ログ管理
 */

// ============================================
// 基本設定（スプレッドシートから読み込み）
// ============================================
function getConfig() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName('設定');

  if (!configSheet) {
    throw new Error('設定シートが見つかりません。setupSpreadsheet() を実行してください。');
  }

  const data = configSheet.getRange('A2:B20').getValues();
  const config = {};

  for (const row of data) {
    if (row[0] && row[1]) {
      config[row[0]] = row[1];
    }
  }

  // 必須項目チェック
  const required = ['TRANSCRIPT_FOLDER_ID', 'OUTPUT_FOLDER_ID', 'GEMINI_API_KEY'];
  for (const key of required) {
    if (!config[key]) {
      throw new Error(`設定シートに ${key} が設定されていません。`);
    }
  }

  // 通知メールはカンマ区切りで複数対応
  if (config['NOTIFICATION_EMAILS']) {
    config['NOTIFICATION_EMAILS'] = config['NOTIFICATION_EMAILS'].split(',').map(e => e.trim());
  } else {
    config['NOTIFICATION_EMAILS'] = [];
  }

  // デフォルト値設定
  config['GEMINI_MODEL'] = config['GEMINI_MODEL'] || 'gemini-1.5-pro';
  config['GEMINI_ENDPOINT'] = 'https://generativelanguage.googleapis.com/v1beta/models/';

  return config;
}

// ============================================
// スプレッドシート初期セットアップ
// ============================================
function setupSpreadsheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. 設定シート作成
  let configSheet = ss.getSheetByName('設定');
  if (!configSheet) {
    configSheet = ss.insertSheet('設定');
  }
  setupConfigSheet(configSheet);

  // 2. 処理ログシート作成
  let logSheet = ss.getSheetByName('処理ログ');
  if (!logSheet) {
    logSheet = ss.insertSheet('処理ログ');
  }
  setupLogSheet(logSheet);

  // 3. エラーログシート作成
  let errorSheet = ss.getSheetByName('エラーログ');
  if (!errorSheet) {
    errorSheet = ss.insertSheet('エラーログ');
  }
  setupErrorSheet(errorSheet);

  // 4. ユーザー管理シート作成
  let userSheet = ss.getSheetByName('ユーザー管理');
  if (!userSheet) {
    userSheet = ss.insertSheet('ユーザー管理');
  }
  setupUserSheet(userSheet);

  // 5. 統計シート作成
  let statsSheet = ss.getSheetByName('統計');
  if (!statsSheet) {
    statsSheet = ss.insertSheet('統計');
  }
  setupStatsSheet(statsSheet);

  Logger.log('=== スプレッドシートのセットアップが完了しました ===');
  Logger.log('次のステップ:');
  Logger.log('1. 「設定」シートの各項目を入力してください');
  Logger.log('2. 共有フォルダを作成し、TRANSCRIPT_FOLDER_ID を設定');
  Logger.log('3. testGeminiConnection() で接続テスト');
  Logger.log('4. createTimeTrigger() でトリガー作成');

  // UIにも表示
  SpreadsheetApp.getUi().alert(
    'セットアップ完了',
    '各シートが作成されました。\n\n' +
    '次のステップ:\n' +
    '1. 「設定」シートの各項目を入力\n' +
    '2. 共有フォルダを作成しIDを設定\n' +
    '3. testGeminiConnection() で接続テスト\n' +
    '4. createTimeTrigger() でトリガー作成',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function setupConfigSheet(sheet) {
  sheet.clear();

  // ヘッダー
  sheet.getRange('A1:B1').setValues([['設定項目', '値']]);
  sheet.getRange('A1:B1')
    .setBackground('#4285f4')
    .setFontColor('#ffffff')
    .setFontWeight('bold');

  // 設定項目
  const configItems = [
    ['TRANSCRIPT_FOLDER_ID', '', 'トランスクリプト保存先の共有フォルダID'],
    ['OUTPUT_FOLDER_ID', '', '議事録出力先フォルダID'],
    ['MASTER_DOC_ID', '', 'マスタードキュメントID（集約用、空欄可）'],
    ['GEMINI_API_KEY', '', 'Gemini API キー'],
    ['GEMINI_MODEL', 'gemini-1.5-pro', '使用するGeminiモデル'],
    ['NOTIFICATION_EMAILS', '', '通知先メール（カンマ区切りで複数可）'],
    ['ENABLE_MASTER_DOC', 'TRUE', 'マスタードキュメント集約を有効化'],
    ['ENABLE_EMAIL_NOTIFICATION', 'TRUE', 'メール通知を有効化'],
    ['POLLING_INTERVAL_MINUTES', '15', 'ポーリング間隔（分）'],
  ];

  for (let i = 0; i < configItems.length; i++) {
    sheet.getRange(i + 2, 1).setValue(configItems[i][0]);
    sheet.getRange(i + 2, 2).setValue(configItems[i][1]);
    sheet.getRange(i + 2, 3).setValue(configItems[i][2]).setFontColor('#666666');
  }

  // 列幅調整
  sheet.setColumnWidth(1, 250);
  sheet.setColumnWidth(2, 400);
  sheet.setColumnWidth(3, 350);

  // 説明ヘッダー
  sheet.getRange('C1').setValue('説明').setBackground('#f0f0f0').setFontWeight('bold');
}

function setupLogSheet(sheet) {
  sheet.clear();

  // ヘッダー
  const headers = ['処理日時', 'ファイル名', 'アップロードユーザー', '議事録URL', 'ステータス', '処理時間(秒)'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setBackground('#34a853')
    .setFontColor('#ffffff')
    .setFontWeight('bold');

  // 列幅調整
  sheet.setColumnWidth(1, 180);
  sheet.setColumnWidth(2, 250);
  sheet.setColumnWidth(3, 200);
  sheet.setColumnWidth(4, 350);
  sheet.setColumnWidth(5, 100);
  sheet.setColumnWidth(6, 120);

  // フィルター設定
  sheet.getRange(1, 1, 1, headers.length).createFilter();
}

function setupErrorSheet(sheet) {
  sheet.clear();

  // ヘッダー
  const headers = ['発生日時', 'ファイル名', 'エラー内容', '対応状況'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setBackground('#ea4335')
    .setFontColor('#ffffff')
    .setFontWeight('bold');

  // 列幅調整
  sheet.setColumnWidth(1, 180);
  sheet.setColumnWidth(2, 250);
  sheet.setColumnWidth(3, 500);
  sheet.setColumnWidth(4, 150);
}

function setupUserSheet(sheet) {
  sheet.clear();

  // ヘッダー
  const headers = ['メールアドレス', '名前', '登録日', '処理件数', '最終処理日', '個別通知'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setBackground('#fbbc04')
    .setFontColor('#000000')
    .setFontWeight('bold');

  // 列幅調整
  sheet.setColumnWidth(1, 250);
  sheet.setColumnWidth(2, 150);
  sheet.setColumnWidth(3, 120);
  sheet.setColumnWidth(4, 100);
  sheet.setColumnWidth(5, 150);
  sheet.setColumnWidth(6, 100);

  // 説明行
  sheet.getRange(2, 1, 1, headers.length).setValues([[
    '（自動登録）', '（手動入力可）', '（自動）', '（自動）', '（自動）', 'TRUE/FALSE'
  ]]);
  sheet.getRange(2, 1, 1, headers.length).setFontColor('#999999').setFontStyle('italic');
}

function setupStatsSheet(sheet) {
  sheet.clear();

  // タイトル
  sheet.getRange('A1').setValue('📊 議事録自動化システム 統計ダッシュボード');
  sheet.getRange('A1').setFontSize(16).setFontWeight('bold');

  // 統計項目
  const stats = [
    ['', ''],
    ['総処理件数', '=COUNTA(\'処理ログ\'!A:A)-1'],
    ['今月の処理件数', '=COUNTIFS(\'処理ログ\'!A:A,">="&DATE(YEAR(TODAY()),MONTH(TODAY()),1),\'処理ログ\'!A:A,"<="&EOMONTH(TODAY(),0))'],
    ['今週の処理件数', '=COUNTIFS(\'処理ログ\'!A:A,">="&(TODAY()-WEEKDAY(TODAY(),2)+1),\'処理ログ\'!A:A,"<="&TODAY())'],
    ['エラー件数', '=COUNTA(\'エラーログ\'!A:A)-1'],
    ['登録ユーザー数', '=COUNTA(\'ユーザー管理\'!A:A)-2'],
    ['平均処理時間(秒)', '=IFERROR(AVERAGE(\'処理ログ\'!F:F),0)'],
  ];

  sheet.getRange(2, 1, stats.length, 2).setValues(stats);

  // スタイル
  sheet.getRange('A3:A8').setFontWeight('bold');
  sheet.setColumnWidth(1, 200);
  sheet.setColumnWidth(2, 150);

  // 値セルの書式
  sheet.getRange('B3:B8').setNumberFormat('#,##0');
}

// ============================================
// メイン処理：新規トランスクリプト検知・処理
// ============================================
function processNewTranscripts() {
  const config = getConfig();
  const folder = DriveApp.getFolderById(config['TRANSCRIPT_FOLDER_ID']);
  const files = folder.getFiles();
  const processedIds = getProcessedFileIds();

  let processedCount = 0;

  while (files.hasNext()) {
    const file = files.next();
    const fileId = file.getId();
    const fileName = file.getName();

    // 処理済みファイルはスキップ
    if (processedIds.includes(fileId)) {
      continue;
    }

    // トランスクリプトファイルのみ処理
    if (isTranscriptFile(file)) {
      const startTime = new Date();

      try {
        Logger.log(`処理開始: ${fileName}`);

        // ファイルオーナー（アップロードユーザー）取得
        const owner = getFileOwner(file);

        // トランスクリプト取得
        const transcript = extractTranscript(file);

        if (!transcript || transcript.trim() === '') {
          Logger.log(`トランスクリプトが空です: ${fileName}`);
          markAsProcessed(fileId);
          continue;
        }

        // Gemini で構造化議事録生成
        const structuredMinutes = generateStructuredMinutes(transcript, fileName, config);

        // Google Docs として保存
        const docUrl = saveAsGoogleDoc(structuredMinutes, fileName, config);

        // 処理時間計算
        const processingTime = Math.round((new Date() - startTime) / 1000);

        // 処理ログ記録
        logProcessing(fileName, owner, docUrl, '成功', processingTime);

        // ユーザー情報更新
        updateUserStats(owner);

        // メール通知
        if (config['ENABLE_EMAIL_NOTIFICATION'] === 'TRUE' || config['ENABLE_EMAIL_NOTIFICATION'] === true) {
          sendNotificationEmail(structuredMinutes, docUrl, fileName, owner, config);
        }

        // マスタードキュメントへ追記
        if ((config['ENABLE_MASTER_DOC'] === 'TRUE' || config['ENABLE_MASTER_DOC'] === true) && config['MASTER_DOC_ID']) {
          appendToMasterDoc(structuredMinutes, fileName, config);
        }

        // 処理済みとしてマーク
        markAsProcessed(fileId);

        Logger.log(`処理完了: ${fileName} (${processingTime}秒)`);
        processedCount++;

      } catch (error) {
        Logger.log(`エラー発生 (${fileName}): ${error.message}`);
        logError(fileName, error.message);
        sendErrorNotification(fileName, error.message, config);
      }
    }
  }

  if (processedCount > 0) {
    Logger.log(`今回の処理件数: ${processedCount}`);
  }
}

// ============================================
// ファイルオーナー取得
// ============================================
function getFileOwner(file) {
  try {
    const owner = file.getOwner();
    if (owner) {
      return owner.getEmail();
    }
  } catch (e) {
    // 共有ドライブの場合はオーナー取得できない場合がある
  }

  // 最終更新者で代替
  try {
    const lastUpdater = file.getLastUpdated();
    return '不明';
  } catch (e) {
    return '不明';
  }
}

// ============================================
// ログ記録
// ============================================
function logProcessing(fileName, owner, docUrl, status, processingTime) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName('処理ログ');

  if (!logSheet) return;

  const now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
  logSheet.appendRow([now, fileName, owner, docUrl, status, processingTime]);
}

function logError(fileName, errorMessage) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const errorSheet = ss.getSheetByName('エラーログ');

  if (!errorSheet) return;

  const now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
  errorSheet.appendRow([now, fileName, errorMessage, '未対応']);
}

// ============================================
// ユーザー統計更新
// ============================================
function updateUserStats(email) {
  if (!email || email === '不明') return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const userSheet = ss.getSheetByName('ユーザー管理');

  if (!userSheet) return;

  const data = userSheet.getDataRange().getValues();
  const now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');

  // 既存ユーザーを探す
  for (let i = 2; i < data.length; i++) {
    if (data[i][0] === email) {
      // 処理件数を+1、最終処理日を更新
      const currentCount = data[i][3] || 0;
      userSheet.getRange(i + 1, 4).setValue(currentCount + 1);
      userSheet.getRange(i + 1, 5).setValue(now);
      return;
    }
  }

  // 新規ユーザー登録
  const newRow = data.length + 1;
  userSheet.getRange(newRow, 1, 1, 6).setValues([[
    email,
    '',
    Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd'),
    1,
    now,
    'TRUE'
  ]]);
}

// ============================================
// トランスクリプト抽出
// ============================================
function isTranscriptFile(file) {
  const mimeType = file.getMimeType();
  const name = file.getName().toLowerCase();

  return (
    name.endsWith('.vtt') ||
    name.endsWith('.srt') ||
    mimeType === MimeType.GOOGLE_DOCS ||
    (mimeType === MimeType.PLAIN_TEXT && name.includes('transcript'))
  );
}

function extractTranscript(file) {
  const mimeType = file.getMimeType();

  if (mimeType === MimeType.GOOGLE_DOCS) {
    const doc = DocumentApp.openById(file.getId());
    return doc.getBody().getText();
  }

  // VTT/SRT/テキストファイル
  const content = file.getBlob().getDataAsString();
  return parseVttContent(content);
}

function parseVttContent(vttContent) {
  const lines = vttContent.split('\n');
  const textLines = [];

  for (const line of lines) {
    if (line.includes('-->') ||
        line.startsWith('WEBVTT') ||
        line.trim() === '' ||
        /^\d+$/.test(line.trim())) {
      continue;
    }
    textLines.push(line.trim());
  }

  return textLines.join(' ');
}

// ============================================
// Gemini API 連携
// ============================================
function generateStructuredMinutes(transcript, fileName, config) {
  const prompt = buildPrompt(transcript, fileName);

  const url = `${config['GEMINI_ENDPOINT']}${config['GEMINI_MODEL']}:generateContent?key=${config['GEMINI_API_KEY']}`;

  const payload = {
    contents: [{
      parts: [{
        text: prompt
      }]
    }],
    generationConfig: {
      temperature: 0.3,
      maxOutputTokens: 8192
    }
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);
  const responseCode = response.getResponseCode();

  if (responseCode !== 200) {
    throw new Error(`Gemini API エラー: ${responseCode} - ${response.getContentText()}`);
  }

  const result = JSON.parse(response.getContentText());

  if (!result.candidates || !result.candidates[0] || !result.candidates[0].content) {
    throw new Error('Gemini API から有効なレスポンスがありません');
  }

  return result.candidates[0].content.parts[0].text;
}

function buildPrompt(transcript, fileName) {
  const today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm');

  return `
あなたは優秀な議事録作成アシスタントです。
以下の会議トランスクリプトを分析し、構造化された議事録を作成してください。

## 入力情報
- ファイル名: ${fileName}
- 処理日時: ${today}

## トランスクリプト
${transcript}

## 出力形式（必ずこの構造で出力）

# 議事録

## 1. 会議概要
- **日時**: （トランスクリプトから推測、不明なら「要確認」）
- **参加者**: （発言者名を列挙）
- **会議目的**: （議論内容から推測）

## 2. 決定事項
（各決定事項について、背景と結論を明記）
- **決定1**:
  - 背景:
  - 結論:

## 3. ネクストアクション（ToDo）
| タスク | 担当者 | 期限 |
|--------|--------|------|
| タスク内容 | 担当者名 | 期限日 |

## 4. 議論の要点
（主要な議論について、賛否と最終結論を整理）
- **議題1**:
  - 賛成意見:
  - 反対意見:
  - 結論:

## 5. 保留・検討事項
（未決定で持ち越しになった事項）
-

## 6. 要約（3行）
（会議全体の要約を3行で）

---

注意事項:
- 固有名詞は正確に
- 推測部分は「（推測）」と明記
- 不明な箇所は「要確認」と記載
- 日本語で出力
`;
}

// ============================================
// Google Docs 保存
// ============================================
function saveAsGoogleDoc(content, originalFileName, config) {
  const today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd');
  const meetingName = extractMeetingName(originalFileName);
  const docTitle = `${today}_${meetingName}_議事録`;

  const doc = DocumentApp.create(docTitle);
  const body = doc.getBody();

  applyContentToDoc(body, content);

  doc.saveAndClose();

  // 指定フォルダに移動
  const file = DriveApp.getFileById(doc.getId());
  const outputFolder = DriveApp.getFolderById(config['OUTPUT_FOLDER_ID']);
  file.moveTo(outputFolder);

  return doc.getUrl();
}

function extractMeetingName(fileName) {
  let name = fileName
    .replace(/\.(vtt|srt|txt)$/i, '')
    .replace(/transcript/i, '')
    .replace(/[_-]/g, ' ')
    .trim();

  return name || '会議';
}

function applyContentToDoc(body, content) {
  const lines = content.split('\n');

  for (const line of lines) {
    if (line.startsWith('# ')) {
      const para = body.appendParagraph(line.substring(2));
      para.setHeading(DocumentApp.ParagraphHeading.HEADING1);
    } else if (line.startsWith('## ')) {
      const para = body.appendParagraph(line.substring(3));
      para.setHeading(DocumentApp.ParagraphHeading.HEADING2);
    } else if (line.startsWith('### ')) {
      const para = body.appendParagraph(line.substring(4));
      para.setHeading(DocumentApp.ParagraphHeading.HEADING3);
    } else if (line.startsWith('|')) {
      body.appendParagraph(line);
    } else if (line.startsWith('- ')) {
      const listItem = body.appendListItem(line.substring(2));
      listItem.setGlyphType(DocumentApp.GlyphType.BULLET);
    } else if (line.trim() === '---') {
      body.appendHorizontalRule();
    } else if (line.trim() !== '') {
      body.appendParagraph(line);
    }
  }
}

// ============================================
// メール通知
// ============================================
function sendNotificationEmail(content, docUrl, fileName, owner, config) {
  const subject = `【議事録生成完了】${extractMeetingName(fileName)}`;

  const summaryMatch = content.match(/## 6\. 要約[\s\S]*?(?=---|$)/);
  const summary = summaryMatch ? summaryMatch[0] : '要約を取得できませんでした';

  const htmlBody = `
    <h2>議事録が自動生成されました</h2>
    <p><strong>元ファイル:</strong> ${fileName}</p>
    <p><strong>アップロード者:</strong> ${owner}</p>
    <p><strong>議事録URL:</strong> <a href="${docUrl}">${docUrl}</a></p>
    <hr>
    <h3>要約</h3>
    <pre>${summary}</pre>
    <hr>
    <p><small>このメールは会議議事録自動化システムから送信されました。</small></p>
  `;

  // 設定シートの通知先に送信
  const emails = config['NOTIFICATION_EMAILS'] || [];
  for (const email of emails) {
    if (email) {
      GmailApp.sendEmail(email, subject, summary, { htmlBody: htmlBody });
    }
  }

  // アップロードユーザーにも通知（個別通知が有効な場合）
  if (owner && owner !== '不明' && shouldNotifyUser(owner)) {
    if (!emails.includes(owner)) {
      GmailApp.sendEmail(owner, subject, summary, { htmlBody: htmlBody });
    }
  }
}

function shouldNotifyUser(email) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const userSheet = ss.getSheetByName('ユーザー管理');

  if (!userSheet) return true;

  const data = userSheet.getDataRange().getValues();
  for (let i = 2; i < data.length; i++) {
    if (data[i][0] === email) {
      return data[i][5] === true || data[i][5] === 'TRUE';
    }
  }

  return true; // デフォルトは通知する
}

function sendErrorNotification(fileName, errorMessage, config) {
  const subject = `【議事録生成エラー】${fileName}`;
  const body = `
議事録の自動生成中にエラーが発生しました。

ファイル名: ${fileName}
エラー内容: ${errorMessage}

手動での確認をお願いします。
  `;

  const emails = config['NOTIFICATION_EMAILS'] || [];
  for (const email of emails) {
    if (email) {
      GmailApp.sendEmail(email, subject, body);
    }
  }
}

// ============================================
// マスタードキュメント集約
// ============================================
function appendToMasterDoc(content, fileName, config) {
  if (!config['MASTER_DOC_ID']) return;

  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(30000);

    const masterDoc = DocumentApp.openById(config['MASTER_DOC_ID']);
    const body = masterDoc.getBody();

    const today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm');
    const meetingName = extractMeetingName(fileName);

    body.appendHorizontalRule();

    const header = body.appendParagraph(`📅 ${today} - ${meetingName}`);
    header.setHeading(DocumentApp.ParagraphHeading.HEADING1);

    applyContentToDoc(body, content);

    body.appendParagraph('');

    masterDoc.saveAndClose();

  } finally {
    lock.releaseLock();
  }
}

// ============================================
// 処理済みファイル管理
// ============================================
function getProcessedFileIds() {
  const props = PropertiesService.getScriptProperties();
  const stored = props.getProperty('processedFileIds');
  return stored ? JSON.parse(stored) : [];
}

function markAsProcessed(fileId) {
  const props = PropertiesService.getScriptProperties();
  const processedIds = getProcessedFileIds();

  if (!processedIds.includes(fileId)) {
    processedIds.push(fileId);
    props.setProperty('processedFileIds', JSON.stringify(processedIds));
  }
}

function clearProcessedFiles() {
  const props = PropertiesService.getScriptProperties();
  props.deleteProperty('processedFileIds');
  Logger.log('処理済みファイルリストをクリアしました');

  SpreadsheetApp.getUi().alert('処理済みファイルリストをクリアしました。');
}

// ============================================
// トリガー設定
// ============================================
function createTimeTrigger() {
  // 既存のトリガーを削除
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === 'processNewTranscripts') {
      ScriptApp.deleteTrigger(trigger);
    }
  }

  // 設定からインターバル取得（デフォルト15分）
  let interval = 15;
  try {
    const config = getConfig();
    interval = parseInt(config['POLLING_INTERVAL_MINUTES']) || 15;
  } catch (e) {
    // 設定シートがない場合はデフォルト値を使用
  }

  // トリガー作成
  ScriptApp.newTrigger('processNewTranscripts')
    .timeBased()
    .everyMinutes(interval)
    .create();

  Logger.log(`${interval}分間隔のトリガーを作成しました`);
  SpreadsheetApp.getUi().alert(`${interval}分間隔のトリガーを作成しました。`);
}

function deleteAllTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    ScriptApp.deleteTrigger(trigger);
  }
  Logger.log('すべてのトリガーを削除しました');
  SpreadsheetApp.getUi().alert('すべてのトリガーを削除しました。');
}

// ============================================
// カスタムメニュー
// ============================================
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🎙️ 議事録自動化')
    .addItem('📋 初期セットアップ', 'setupSpreadsheet')
    .addSeparator()
    .addItem('▶️ 手動実行（今すぐ処理）', 'processNewTranscripts')
    .addItem('🔗 Gemini接続テスト', 'testGeminiConnection')
    .addSeparator()
    .addItem('⏰ トリガー作成', 'createTimeTrigger')
    .addItem('🗑️ トリガー削除', 'deleteAllTriggers')
    .addSeparator()
    .addItem('🔄 処理済みリストをクリア', 'clearProcessedFiles')
    .addToUi();
}

// ============================================
// テスト用関数
// ============================================
function testGeminiConnection() {
  try {
    const config = getConfig();
    const testPrompt = 'こんにちは、接続テストです。「接続成功」と返答してください。';

    const url = `${config['GEMINI_ENDPOINT']}${config['GEMINI_MODEL']}:generateContent?key=${config['GEMINI_API_KEY']}`;

    const payload = {
      contents: [{
        parts: [{ text: testPrompt }]
      }]
    };

    const options = {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();

    if (responseCode === 200) {
      const result = JSON.parse(response.getContentText());
      const responseText = result.candidates[0].content.parts[0].text;
      Logger.log(`接続成功: ${responseText}`);
      SpreadsheetApp.getUi().alert(`✅ Gemini API 接続成功\n\nレスポンス: ${responseText}`);
    } else {
      Logger.log(`接続失敗: ${responseCode}`);
      SpreadsheetApp.getUi().alert(`❌ Gemini API 接続失敗\n\nステータス: ${responseCode}\n${response.getContentText()}`);
    }
  } catch (error) {
    Logger.log(`エラー: ${error.message}`);
    SpreadsheetApp.getUi().alert(`❌ エラー\n\n${error.message}`);
  }
}

function testWithSampleTranscript() {
  const sampleTranscript = `
    田中: それでは、今週のプロジェクト進捗会議を始めます。
    鈴木: はい、まず開発の進捗ですが、API実装が80%完了しました。
    田中: 素晴らしいですね。残りの20%はいつ頃完了予定ですか？
    鈴木: 来週金曜日には完了できる見込みです。
    山田: デザインチームからですが、UIの最終調整が必要です。
    田中: 了解しました。では、来週金曜までにAPI完了、UIは山田さんが来週水曜までに調整、ということで決定しましょう。
    全員: 了解しました。
  `;

  try {
    const config = getConfig();
    const result = generateStructuredMinutes(sampleTranscript, 'テスト会議', config);
    Logger.log(result);
    SpreadsheetApp.getUi().alert('✅ サンプルテスト完了\n\n結果はログを確認してください。');
  } catch (error) {
    SpreadsheetApp.getUi().alert(`❌ エラー\n\n${error.message}`);
  }
}
