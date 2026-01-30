// 基本定数
const CONFIG_SHEET = 'Config';
const INBOX_SHEET = 'Inbox';
const SPEC_SHEET = '業務記述書';
const FLOW_SHEET = 'フロー';
const VISUAL_SHEET = 'ビジュアルフロー';
const ACTIVITY_LOG_SHEET = 'ActivityLog';

// スプレッドシート関連ユーティリティ
function ss() {
  return SpreadsheetApp.getActive();
}

function file() {
  return DriveApp.getFileById(ss().getId());
}

// Config管理
function getConfig(key) {
  const sh = ss().getSheetByName(CONFIG_SHEET);
  if (!sh) return null;
  
  const values = sh.getRange(1, 1, sh.getLastRow(), 2).getValues();
  const configMap = new Map(values.map(r => [String(r[0]).trim(), String(r[1]).trim()]));
  return configMap.get(key);
}

function setConfig(key, value) {
  const sh = ss().getSheetByName(CONFIG_SHEET);
  if (!sh) return;
  
  const values = sh.getRange(1, 1, sh.getLastRow(), 2).getValues();
  let found = false;
  
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]).trim() === key) {
      sh.getRange(i + 1, 2).setValue(value);
      found = true;
      break;
    }
  }
  
  if (!found) {
    sh.appendRow([key, value]);
  }
}

// 初期Config設定
function initializeConfig() {
  const sh = ss().getSheetByName(CONFIG_SHEET) || ss().insertSheet(CONFIG_SHEET);
  
  const defaultConfigs = [
    ['PROCESSING_QUERY', 'label:inbox is:unread subject:"[WORK-REQ]"'],
    ['DEFAULT_TO_EMAIL', ''],
    ['OPENAI_MODEL', 'gpt-4o-mini'],
    ['ORG_PROFILE_JSON', '{"listing":"上場区分","industry":"業種","jurisdictions":["JP"],"policies":["内部統制準拠"]}'],
    ['SHARE_ANYONE_WITH_LINK', 'TRUE'],
    ['FLOW_SHEET_NAME', 'フロー'],
    ['VISUAL_SHEET_NAME', 'ビジュアルフロー'],
    ['LEGAL_JURISDICTIONS', 'JP, global'],
    ['MAX_RETRY_COUNT', '3'],
    ['RETRY_DELAY_MS', '2000']
  ];
  
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, defaultConfigs.length, 2).setValues(defaultConfigs);
  }
}

// 共有設定
function shareSheetAnyWithLink() {
  try {
    file().setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.EDIT);
    logActivity('SHARE', 'Set ANYONE_WITH_LINK permissions');
    return true;
  } catch (e) {
    logActivity('SHARE_ERROR', `Failed to set ANYONE_WITH_LINK: ${e.toString()}`);
    return false;
  }
}

function addEditor(email) {
  try {
    file().addEditor(email);
    logActivity('SHARE', `Added editor: ${email}`);
    return true;
  } catch (e) {
    logActivity('SHARE_ERROR', `Failed to add editor ${email}: ${e.toString()}`);
    return false;
  }
}

// メールアドレス抽出
function extractEmail(fromHeader) {
  const match = fromHeader.match(/<([^>]+)>/);
  return match ? match[1] : fromHeader.replace(/.*\s/, '').trim();
}

// HTMLをテキストに変換
function htmlToText(html) {
  return html
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<\/p>/gi, '\n\n')
    .replace(/<[^>]+>/g, ' ')
    .replace(/&nbsp;/g, ' ')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&amp;/g, '&')
    .replace(/\s+/g, ' ')
    .trim();
}

// 処理済みチェック
function isProcessed(messageId) {
  const sh = ss().getSheetByName(INBOX_SHEET);
  if (!sh || sh.getLastRow() === 0) return false;
  
  const values = sh.getRange(2, 3, sh.getLastRow() - 1, 1).getValues();
  return values.some(row => row[0] === messageId);
}

function markProcessed(messageId) {
  const sh = ss().getSheetByName(INBOX_SHEET);
  if (!sh) return;
  
  const values = sh.getRange(2, 3, sh.getLastRow() - 1, 1).getValues();
  for (let i = 0; i < values.length; i++) {
    if (values[i][0] === messageId) {
      sh.getRange(i + 2, 7).setValue('PROCESSED');
      sh.getRange(i + 2, 8).setValue(new Date());
      break;
    }
  }
}

// Inboxログ記録
function logInbox(messageId, threadId, from, subject, summary, status) {
  const sh = ss().getSheetByName(INBOX_SHEET) || createInboxSheet();
  sh.appendRow([
    new Date(),
    threadId,
    messageId,
    from,
    subject,
    summary,
    status,
    status === 'PROCESSED' ? new Date() : '',
    ''
  ]);
}

function createInboxSheet() {
  const sh = ss().insertSheet(INBOX_SHEET);
  sh.getRange(1, 1, 1, 9).setValues([[
    '受信日時', 'ThreadId', 'MessageId', 'From', 'Subject', 
    '要約', 'ステータス', '処理日時', 'エラー'
  ]]);
  sh.getRange(1, 1, 1, 9).setFontWeight('bold');
  return sh;
}

// エラーログ記録
function logError(messageId, error) {
  const sh = ss().getSheetByName(INBOX_SHEET);
  if (!sh) return;
  
  const values = sh.getRange(2, 3, sh.getLastRow() - 1, 1).getValues();
  for (let i = 0; i < values.length; i++) {
    if (values[i][0] === messageId) {
      sh.getRange(i + 2, 7).setValue('ERROR');
      sh.getRange(i + 2, 9).setValue(error.toString());
      break;
    }
  }
  
  logActivity('ERROR', `MessageId: ${messageId}, Error: ${error.toString()}`);
}

// アクティビティログ
function logActivity(type, details) {
  const sh = ss().getSheetByName(ACTIVITY_LOG_SHEET) || createActivityLogSheet();
  let userEmail;
  try {
    userEmail = Session.getActiveUser().getEmail();
  } catch (e) {
    userEmail = 'Unknown User';
  }
  sh.appendRow([
    new Date(),
    type,
    details,
    userEmail
  ]);
}

function createActivityLogSheet() {
  const sh = ss().insertSheet(ACTIVITY_LOG_SHEET);
  sh.getRange(1, 1, 1, 4).setValues([[
    'タイムスタンプ', 'タイプ', '詳細', 'ユーザー'
  ]]);
  sh.getRange(1, 1, 1, 4).setFontWeight('bold');
  sh.hideSheet();
  return sh;
}

// リトライ処理
function retryWithBackoff(func, maxRetries = 3) {
  const configRetries = parseInt(getConfig('MAX_RETRY_COUNT') || '3');
  const retryDelay = parseInt(getConfig('RETRY_DELAY_MS') || '2000');
  maxRetries = configRetries;
  
  for (let i = 0; i < maxRetries; i++) {
    try {
      return func();
    } catch (e) {
      if (i === maxRetries - 1) throw e;
      Utilities.sleep(retryDelay * Math.pow(2, i));
    }
  }
}