/**
 * Google Apps Script を使用した Google Meet 議事録自動生成とメール報告システム
 * 
 * 主な機能:
 * - Google Calendar から会議イベントを取得
 * - Google Meet API でトランスクリプトを取得
 * - Vertex AI Gemini API で議事録を生成
 * - Gmail で議事録をメール送信
 * - Google Sheets で設定とログを管理
 */

// 設定用スプレッドシートID（実際の値に変更してください）
const CONFIG_SPREADSHEET_ID = '1xFBWhUiJBsD_iatVcQnOjiiI1dZSazAo0y3jRDG0UrM';
const CONFIG_SHEET_NAME = '設定';
const LOG_SHEET_NAME = '実行ログ';
// 初期設定ガイドで使用するスプレッドシート名
const CONFIG_SPREADSHEET_NAME = '議事録自動化設定';

// GCP設定（実際の値に変更してください）
const GCP_PROJECT_ID = 'company-gas';
const GCP_REGION = 'asia-northeast1';
const GEMINI_MODEL = 'gemini-1.5-pro-002';

/**
 * 設定スプレッドシートIDの取得（Script Properties優先、未設定時はエラー）
 */
function getConfigSpreadsheetId() {
  const idFromProperties = PropertiesService.getScriptProperties().getProperty('CONFIG_SPREADSHEET_ID');
  const resolvedId = idFromProperties || CONFIG_SPREADSHEET_ID;
  if (!resolvedId || resolvedId === 'YOUR_CONFIG_SPREADSHEET_ID') {
    throw new Error('CONFIG_SPREADSHEET_ID が未設定です。以下のいずれかで設定してください:\n1) コード中の定数 CONFIG_SPREADSHEET_ID を実スプレッドシートIDに変更\n2) setConfigSpreadsheetId("<スプレッドシートID>") を一度実行して保存（Script Properties）');
  }
  return resolvedId;
}

/**
 * 設定スプレッドシートオブジェクトの取得（存在/権限チェック付き）
 */
function getConfigSpreadsheet() {
  const spreadsheetId = getConfigSpreadsheetId();
  try {
    return SpreadsheetApp.openById(spreadsheetId);
  } catch (error) {
    throw new Error(`設定スプレッドシートへのアクセスに失敗しました。ID: ${spreadsheetId}\n権限(閲覧可以上)とIDの正しさを確認してください。\n詳細: ${error}`);
  }
}

/**
 * メイン処理関数（トリガーから呼び出される）
 */
function main() {
  try {
    console.log('議事録自動化処理を開始します');
    
    // 設定を読み込み
    const config = getConfig();
    console.log('設定を読み込みました:', config);
    
    // 直近の会議イベントを取得
    const events = getRecentEvents(config.calendarId, config.checkHours);
    console.log(`${events.length}件の会議イベントを取得しました`);
    
    // 各イベントを処理
    for (const event of events) {
      try {
        // イベントの有効性を確認
        if (!event) {
          console.error('無効なイベントオブジェクト');
          continue;
        }
        
        let eventId, eventTitle;
        try {
          eventId = event.getId();
          eventTitle = event.getTitle() || '無題の会議';
        } catch (error) {
          console.error('イベント情報の取得に失敗:', error);
          continue;
        }
        
        // 既に処理済みかチェック
        if (isProcessed(eventId)) {
          console.log(`イベント ${eventId} は処理済みのためスキップします`);
          continue;
        }
        
        console.log(`イベント ${eventTitle} を処理開始`);
        
        // トランスクリプトを取得
        const transcript = getTranscriptByEvent(event);
        if (!transcript) {
          logExecution(eventId, 'ERROR', 'トランスクリプトが見つかりません', eventTitle);
          continue;
        }
        
        // 議事録を生成
        const minutes = generateMinutes(config.minutesPrompt, transcript);
        if (!minutes) {
          logExecution(eventId, 'ERROR', '議事録の生成に失敗しました', eventTitle);
          continue;
        }
        
        // メール宛先を収集
        const recipients = collectRecipients(event, config.additionalRecipients);
        
        // メールを送信
        const subject = `【議事録】${eventTitle}`;
        sendMinutesEmail(recipients, subject, minutes);
        
        // 処理完了をログに記録
        logExecution(eventId, 'SUCCESS', '', eventTitle);
        console.log(`イベント ${eventTitle} の処理が完了しました`);
        
      } catch (error) {
        const errorTitle = eventTitle || 'unknown';
        const errorId = eventId || 'unknown';
        console.error(`イベント ${errorTitle} の処理中にエラーが発生:`, error);
        logExecution(errorId, 'ERROR', error.toString(), errorTitle);
      }
    }
    
    console.log('議事録自動化処理が完了しました');
    
  } catch (error) {
    console.error('メイン処理でエラーが発生:', error);
    // 管理者への通知メール（オプション）
    sendErrorNotification(error);
  }
}

/**
 * 設定シートから設定値を取得
 */
function getConfig() {
  const sheet = getConfigSpreadsheet().getSheetByName(CONFIG_SHEET_NAME);
  if (!sheet) {
    throw new Error(`設定シート "${CONFIG_SHEET_NAME}" が見つかりません。initializeConfigSheet() を実行して初期化してください。`);
  }
  const values = sheet.getRange('A:B').getValues();
  
  const config = {};
  for (const [key, value] of values) {
    if (key && value) {
      config[key] = value;
    }
  }
  
  // 必要な設定項目をチェック
  const requiredKeys = ['CalendarId', 'MinutesPrompt', 'CheckHours'];
  for (const key of requiredKeys) {
    if (!config[key]) {
      throw new Error(`必要な設定項目 "${key}" が見つかりません`);
    }
  }
  
  return {
    calendarId: config.CalendarId,
    minutesPrompt: config.MinutesPrompt,
    additionalRecipients: config.AdditionalRecipients || '',
    checkHours: parseInt(config.CheckHours) || 1
  };
}

/**
 * 直近の会議イベントを取得
 */
function getRecentEvents(calendarId, checkHours) {
  const now = new Date();
  const startTime = new Date(now.getTime() - (checkHours * 60 * 60 * 1000));

  // カレンダーの取得（エラーハンドリング付き）
  let calendar;
  try {
    if (!calendarId || calendarId === '' || calendarId === 'your-calendar-id@group.calendar.google.com') {
      console.log('デフォルトカレンダーを使用します');
      calendar = CalendarApp.getDefaultCalendar();
    } else {
      console.log(`指定されたカレンダーIDを使用: ${calendarId}`);
      calendar = CalendarApp.getCalendarById(calendarId);
    }

    // カレンダーが取得できない場合はデフォルトを試行
    if (!calendar) {
      console.error(`カレンダーが見つかりません: ${calendarId}`);
      console.log('デフォルトカレンダーを試行します');
      calendar = CalendarApp.getDefaultCalendar();
    }

    // それでも取得できない場合は空配列を返す
    if (!calendar) {
      console.error('カレンダーにアクセスできません');
      return [];
    }

  } catch (error) {
    console.error('カレンダー取得エラー:', error);
    // エラーの場合もデフォルトカレンダーを試行
    try {
      calendar = CalendarApp.getDefaultCalendar();
      if (!calendar) {
        return [];
      }
    } catch (e) {
      console.error('デフォルトカレンダーも取得できません:', e);
      return [];
    }
  }

  // イベントの取得
  try {
    const events = calendar.getEvents(startTime, now);
    console.log(`取得したイベント数: ${events.length}`);
    
    // 各イベントの詳細をログ出力（デバッグ用）
    events.forEach((event, index) => {
      console.log(`イベント${index + 1}: ${event.getTitle()}`);
      const description = event.getDescription() || '';
      console.log(`  説明文字数: ${description.length}`);
      if (description.length > 0) {
        console.log(`  説明（最初の200文字）: ${description.substring(0, 200)}`);
      }
      
      // Hangouts/Meet情報を別途チェック
      try {
        const conferenceData = event.getConferenceData();
        if (conferenceData) {
          console.log('  カンファレンスデータあり');
        }
      } catch (e) {
        // Conference dataへのアクセスができない場合は無視
      }
    });
    
    // Google Meet リンクがあるイベントのみフィルタ（拡張パターン）
    const meetEvents = events.filter(event => {
      // MeetDetection.gsの高度な検出機能を使用
      try {
        const meetInfo = detectMeetInfo(event);
        if (meetInfo.hasMeet) {
          console.log(`  ✓ Meetイベントとして認識: ${event.getTitle()}`);
          console.log(`    URL: ${meetInfo.meetUrl}`);
          console.log(`    検出方法: ${meetInfo.detectionMethod}`);
          return true;
        }
      } catch (error) {
        console.log(`  Meet情報検出エラー: ${error.message}`);
      }
      
      // フォールバック: 従来の説明欄チェック
      const description = event.getDescription() || '';
      const hasMeetInDescription = 
        description.includes('meet.google.com') || 
        description.includes('Google Meet') ||
        description.includes('ビデオ会議');
      
      if (hasMeetInDescription) {
        console.log(`  ✓ Meetイベントとして認識（説明欄）: ${event.getTitle()}`);
        return true;
      }
      
      return false;
    });
    
    console.log(`Meetイベント数: ${meetEvents.length}`);
    return meetEvents;
  } catch (error) {
    console.error('イベント取得エラー:', error);
    return [];
  }
}

/**
 * イベントが既に処理済みかチェック
 */
function isProcessed(eventId) {
  const sheet = getConfigSpreadsheet().getSheetByName(LOG_SHEET_NAME);
  if (!sheet) {
    return false;
  }
  const values = sheet.getRange('A:A').getValues();
  
  for (const [id] of values) {
    if (id === eventId) {
      return true;
    }
  }
  return false;
}

/**
 * イベントからトランスクリプトを取得（リトライ機能付き）
 */
function getTranscriptByEvent(event) {
  // イベントオブジェクトの検証
  if (!event) {
    console.error('イベントオブジェクトが無効です');
    return null;
  }
  
  // MeetDetection.gsの高度な検出機能を使用
  let meetingCode = null;
  
  try {
    const meetInfo = detectMeetInfo(event);
    if (meetInfo.hasMeet && meetInfo.meetCode) {
      meetingCode = meetInfo.meetCode;
      console.log(`Meet情報検出成功 (${meetInfo.detectionMethod}): ${meetingCode}`);
    }
  } catch (error) {
    console.log('Meet情報の高度な検出に失敗、従来の方法を試行:', error.message);
  }
  
  // フォールバック: 従来の説明欄からの抽出
  if (!meetingCode) {
    let description;
    try {
      description = event.getDescription() || '';
    } catch (error) {
      console.error('イベントの説明を取得できません:', error);
      return null;
    }
    
    // 複数のMeet URLパターンに対応
    const meetPatterns = [
      /https:\/\/meet\.google\.com\/([a-z]{3}-[a-z]{4}-[a-z]{3})/i,  // xxx-xxxx-xxx形式
      /https:\/\/meet\.google\.com\/([a-z0-9\-]+)/i,                // その他の形式
      /meet\.google\.com\/([a-z]{3}-[a-z]{4}-[a-z]{3})/i,            // httpsなし
      /meet\.google\.com\/([a-z0-9\-]+)/i                            // httpsなし、その他
    ];
    
    for (const pattern of meetPatterns) {
      const match = description.match(pattern);
      if (match && match[1]) {
        meetingCode = match[1];
        break;
      }
    }
  }
  
  if (!meetingCode) {
    console.log('Meet URLが見つかりません');
    return null;
  }
  
  console.log(`会議コード: ${meetingCode}`);
  
  // 指数バックオフでリトライ
  const maxRetries = 3;
  for (let i = 0; i < maxRetries; i++) {
    try {
      const transcript = fetchTranscriptFromMeet(meetingCode);
      if (transcript) {
        return transcript;
      }
    } catch (error) {
      console.warn(`トランスクリプト取得試行 ${i + 1} 失敗:`, error);
    }
    
    // 次の試行まで待機（指数バックオフ）
    if (i < maxRetries - 1) {
      const waitTime = Math.pow(2, i) * 1000; // 1秒、2秒、4秒...
      Utilities.sleep(waitTime);
    }
  }
  
  return null;
}

/**
 * Meet APIからトランスクリプトを取得
 */
function fetchTranscriptFromMeet(meetingCode) {
  // 会議コードの検証とクリーニング
  if (!meetingCode || meetingCode === 'undefined') {
    console.error('無効な会議コード:', meetingCode);
    return null;
  }
  
  // 会議コードから余分な文字を削除（ハイフンは保持）
  meetingCode = meetingCode.trim();
  console.log('クリーンな会議コード:', meetingCode);
  
  const baseUrl = 'https://meet.googleapis.com/v2';
  const accessToken = ScriptApp.getOAuthToken();
  
  const options = {
    method: 'GET',
    headers: {
      'Authorization': `Bearer ${accessToken}`,
      'Accept': 'application/json'
    },
    muteHttpExceptions: true
  };
  
  try {
    // 1. 会議レコードを取得（フィルタ形式を修正）
    // 注: Meet API v2 のフィルタ形式は特殊な場合があります
    const recordsUrl = `${baseUrl}/conferenceRecords?filter=meeting_code%3D%22${encodeURIComponent(meetingCode)}%22`;
    console.log('Meet API URL (エンコード済み):', recordsUrl);
    
    const recordsResponse = UrlFetchApp.fetch(recordsUrl, options);
    
    const responseCode = recordsResponse.getResponseCode();
    const responseText = recordsResponse.getContentText();
    
    console.log(`Meet API レスポンスコード: ${responseCode}`);
    
    if (responseCode === 403) {
      console.error('Meet API アクセス権限エラー。必要なスコープ:');
      console.error('- https://www.googleapis.com/auth/meetings.space.readonly');
      console.error('- または Google Workspace の管理者権限');
      console.error('詳細:', responseText);
      return null;
    }
    
    if (responseCode === 404) {
      console.log('会議レコードが見つかりません（会議が終了していないか、録画が有効でない可能性）');
      return null;
    }
    
    if (responseCode !== 200) {
      console.error(`Meet API エラー (${responseCode}):`, responseText);
      return null;
    }
    
    const recordsData = JSON.parse(responseText);
    if (!recordsData.conferenceRecords || recordsData.conferenceRecords.length === 0) {
      console.log('会議レコードが見つかりません（会議終了後しばらく時間がかかる場合があります）');
      return null;
    }
  
  const recordName = recordsData.conferenceRecords[0].name;
  
  // 2. トランスクリプトを取得
  const transcriptsUrl = `${baseUrl}/${recordName}/transcripts`;
  const transcriptsResponse = UrlFetchApp.fetch(transcriptsUrl, options);
  
  if (transcriptsResponse.getResponseCode() !== 200) {
    throw new Error(`トランスクリプト取得エラー: ${transcriptsResponse.getContentText()}`);
  }
  
  const transcriptsData = JSON.parse(transcriptsResponse.getContentText());
  if (!transcriptsData.transcripts || transcriptsData.transcripts.length === 0) {
    console.log('トランスクリプトが見つかりません');
    return null;
  }
  
  const transcriptName = transcriptsData.transcripts[0].name;
  
  // 3. トランスクリプトエントリを取得
  const entriesUrl = `${baseUrl}/${transcriptName}/entries`;
  const entriesResponse = UrlFetchApp.fetch(entriesUrl, options);
  
  if (entriesResponse.getResponseCode() !== 200) {
    throw new Error(`トランスクリプトエントリ取得エラー: ${entriesResponse.getContentText()}`);
  }
  
  const entriesData = JSON.parse(entriesResponse.getContentText());
  if (!entriesData.entries || entriesData.entries.length === 0) {
    console.log('トランスクリプトエントリが見つかりません');
    return null;
  }
  
    // トランスクリプトテキストを結合
    let transcriptText = '';
    for (const entry of entriesData.entries) {
      const participant = entry.participant || '不明';
      const text = entry.text || '';
      transcriptText += `${participant}: ${text}\n`;
    }
    
    return transcriptText;
    
  } catch (error) {
    console.error('Meet API 呼び出しエラー:', error.toString());
    
    // より詳細なエラー情報を提供
    if (error.toString().includes('Invalid argument')) {
      console.error('URLフォーマットエラーの可能性があります');
      console.error('会議コード:', meetingCode);
    }
    
    return null;
  }
}

/**
 * Vertex AI Gemini APIで議事録を生成
 */
function generateMinutes(promptTemplate, transcript) {
  const url = `https://${GCP_REGION}-aiplatform.googleapis.com/v1/projects/${GCP_PROJECT_ID}/locations/${GCP_REGION}/publishers/google/models/${GEMINI_MODEL}:generateContent`;
  
  // プロンプトにトランスクリプトを埋め込み
  const prompt = promptTemplate.replace('{transcript_text}', transcript);
  
  const payload = {
    contents: [{
      role: 'user',
      parts: [{ text: prompt }]
    }],
    generationConfig: {
      temperature: 0.1,
      topP: 0.8,
      maxOutputTokens: 8192
    }
  };
  
  const options = {
    method: 'POST',
    contentType: 'application/json',
    headers: {
      'Authorization': `Bearer ${ScriptApp.getOAuthToken()}`
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  
  const response = UrlFetchApp.fetch(url, options);
  const responseCode = response.getResponseCode();
  
  if (responseCode !== 200) {
    throw new Error(`Vertex AI APIエラー (${responseCode}): ${response.getContentText()}`);
  }
  
  const data = JSON.parse(response.getContentText());
  
  if (!data.candidates || data.candidates.length === 0) {
    throw new Error('Vertex AIから応答が得られませんでした');
  }
  
  const generatedText = data.candidates[0].content?.parts?.[0]?.text;
  if (!generatedText) {
    throw new Error('生成されたテキストが空です');
  }
  
  return generatedText;
}

/**
 * メール宛先を収集
 */
function collectRecipients(event, additionalRecipients) {
  const recipients = new Set();
  
  // カレンダーイベントの参加者を追加
  const guests = event.getGuestList();
  for (const guest of guests) {
    const email = guest.getEmail();
    if (email) {
      recipients.add(email);
    }
  }
  
  // 追加宛先を追加
  if (additionalRecipients) {
    const additionalEmails = additionalRecipients.split(',').map(email => email.trim());
    for (const email of additionalEmails) {
      if (email) {
        recipients.add(email);
      }
    }
  }
  
  return Array.from(recipients);
}

/**
 * 議事録メールを送信
 */
function sendMinutesEmail(recipients, subject, body) {
  if (recipients.length === 0) {
    console.warn('メール宛先が設定されていません');
    return;
  }
  
  const recipientString = recipients.join(',');
  
  try {
    GmailApp.sendEmail(recipientString, subject, body, {
      htmlBody: body.replace(/\n/g, '<br>'),
      name: '議事録自動化システム'
    });
    
    console.log(`メールを送信しました: ${recipientString}`);
  } catch (error) {
    throw new Error(`メール送信エラー: ${error.toString()}`);
  }
}

/**
 * 実行ログを記録
 */
function logExecution(eventId, status, errorMessage, meetingTitle) {
  const spreadsheet = getConfigSpreadsheet();
  let sheet = spreadsheet.getSheetByName(LOG_SHEET_NAME);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(LOG_SHEET_NAME);
    sheet.getRange(1, 1, 1, 5).setValues([[
      'CalendarEventID', 'ProcessedAt', 'Status', 'ErrorMessage', 'MeetingTitle'
    ]]);
  }
  
  const now = new Date();
  const timestamp = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  
  sheet.appendRow([eventId, timestamp, status, errorMessage || '', meetingTitle || '']);
}

/**
 * エラー通知メールを送信（管理者向け）
 */
function sendErrorNotification(error) {
  try {
    const adminEmail = 'admin@example.com'; // 実際の管理者メールアドレスに変更
    const subject = '【エラー】議事録自動化システム';
    const body = `議事録自動化システムでエラーが発生しました。\n\nエラー内容:\n${error.toString()}\n\n時刻: ${new Date()}`;
    
    GmailApp.sendEmail(adminEmail, subject, body);
  } catch (notificationError) {
    console.error('エラー通知の送信に失敗:', notificationError);
  }
}

/**
 * 手動でトリガーを設定する関数
 */
function setupTriggers() {
  try {
    // 既存のトリガーを削除
    const triggers = ScriptApp.getProjectTriggers();
    for (const trigger of triggers) {
      if (trigger.getHandlerFunction() === 'main') {
        ScriptApp.deleteTrigger(trigger);
      }
    }
    
    // 毎時実行のトリガーを作成
    ScriptApp.newTrigger('main')
      .timeBased()
      .everyHours(1)
      .create();
    
    console.log('トリガーを設定しました（毎時実行）');
  } catch (error) {
    console.error('トリガー設定エラー:', error);
    console.log('権限エラーの場合は、appsscript.jsonを保存してから再度認証してください');
    throw error;
  }
}

/**
 * 設定シートの初期化（初回セットアップ用）
 */
function initializeConfigSheet() {
  const spreadsheet = getConfigSpreadsheet();
  
  // 設定シートを作成
  let configSheet = spreadsheet.getSheetByName(CONFIG_SHEET_NAME);
  if (!configSheet) {
    configSheet = spreadsheet.insertSheet(CONFIG_SHEET_NAME);
  }
  
  // 設定データを設定
  const configData = [
    ['CalendarId', 'your-calendar-id@group.calendar.google.com'],
    ['AdditionalRecipients', 'manager@example.com, team@example.com'],
    ['CheckHours', '1'],
    ['MinutesPrompt', `# 役割
あなたは、プロの書記です。

# 指示
以下の会議トランスクリプトを読み、下記のフォーマットに従って簡潔で分かりやすい日本語の議事録を作成してください。

# フォーマット
## 1. 会議の概要
- **会議名**: (会議のタイトルを記載)
- **日時**: (YYYY/MM/DD HH:MM)
- **参加者**: (参加者名を「、」で区切って列挙)

## 2. 決定事項
- (決定したことを箇条書きで記載)
- (複数ある場合は全て記載)

## 3. アクションアイテム
- **担当者名**: (タスク内容を具体的に記載) (期限: YYYY/MM/DD)
- (複数ある場合は全て記載)

## 4. 議論内容の要約
- (主要な議題や議論の流れを箇条書きで要約)

# 制約
- トランスクリプトに存在しない情報は記載しないでください。
- 各項目で記載することがない場合は、「特になし」と記載してください。
- 参加者名は敬称を省略してください。

# 会議トランスクリプト
---
{transcript_text}`]
  ];
  
  configSheet.clear();
  configSheet.getRange(1, 1, configData.length, 2).setValues(configData);
  
  // 実行ログシートを作成
  let logSheet = spreadsheet.getSheetByName(LOG_SHEET_NAME);
  if (!logSheet) {
    logSheet = spreadsheet.insertSheet(LOG_SHEET_NAME);
  }
  
  // ログシートのヘッダーを設定
  const logHeaders = [['CalendarEventID', 'ProcessedAt', 'Status', 'ErrorMessage', 'MeetingTitle']];
  logSheet.clear();
  logSheet.getRange(1, 1, 1, 5).setValues(logHeaders);
  
  console.log('設定シートとログシートを初期化しました');
}

/**
 * Script Properties に CONFIG_SPREADSHEET_ID を設定するユーティリティ
 */
function setConfigSpreadsheetId(spreadsheetId) {
  if (!spreadsheetId || typeof spreadsheetId !== 'string') {
    throw new Error('有効なスプレッドシートIDを指定してください');
  }
  PropertiesService.getScriptProperties().setProperty('CONFIG_SPREADSHEET_ID', spreadsheetId);
  console.log(`Script Properties に CONFIG_SPREADSHEET_ID を設定しました: ${spreadsheetId}`);
}
