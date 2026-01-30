/**
 * 文書管理システム - メインコード
 * スプレッドシートとGASで完結する文書管理システム
 */

// シート名定義
const SHEET_NAME = 'DocRegistry';

// 列定義（0始まりのインデックス）
const COLUMNS = {
  DOC_KEY: 0,
  REV: 1,
  TITLE: 2,
  STAGE: 3,
  DUE_DATE: 4,
  PROJECT_STATUS: 5,
  LAST_SENT_BY: 6,
  GOOGLE_DOC_URL: 7,
  WORD_FILE_URL: 8,
  CREATED_AT: 9,
  LAST_EDITED_AT: 10,
  OWNER_EMAIL: 11,
  LAST_EDITOR: 12,
  LAST_MAIL_URL: 13,
  MAIL_TREE: 14
};

// ステージ定義
const STAGES = {
  DRAFT: 'DRAFT',
  FOR_REVIEW: 'FOR-REVIEW',
  APPROVED: 'APPROVED',
  ARCHIVED: 'ARCHIVED'
};

// プロジェクトステータス定義
const PROJECT_STATUS = {
  OPEN: 'OPEN',
  IN_PROGRESS: 'IN-PROGRESS',
  CLOSED: 'CLOSED',
  DELAYED: 'DELAYED'
};

// 送信者定義
const SENDER_TYPE = {
  SELF: 'SELF',
  PARTNER: 'PARTNER'
};

/**
 * スプレッドシートを初期化
 */
function initializeSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(SHEET_NAME);
  
  if (!sheet) {
    sheet = spreadsheet.insertSheet(SHEET_NAME);
  }
  
  // ヘッダー行を設定
  const headers = [
    'DocKey',
    'Rev',
    'Title',
    'Stage',
    'DueDate',
    'ProjectStatus',
    'LastSentBy',
    'GoogleDocURL',
    'WordFileURL',
    'CreatedAt',
    'LastEditedAt',
    'OwnerEmail',
    'LastEditor',
    'LastMailURL',
    'MailTree'
  ];
  
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#f0f0f0');
  
  // 列幅を調整
  sheet.setColumnWidth(COLUMNS.DOC_KEY + 1, 100);
  sheet.setColumnWidth(COLUMNS.REV + 1, 60);
  sheet.setColumnWidth(COLUMNS.TITLE + 1, 250);
  sheet.setColumnWidth(COLUMNS.STAGE + 1, 120);
  sheet.setColumnWidth(COLUMNS.DUE_DATE + 1, 100);
  sheet.setColumnWidth(COLUMNS.PROJECT_STATUS + 1, 120);
  sheet.setColumnWidth(COLUMNS.LAST_SENT_BY + 1, 100);
  sheet.setColumnWidth(COLUMNS.GOOGLE_DOC_URL + 1, 200);
  sheet.setColumnWidth(COLUMNS.WORD_FILE_URL + 1, 200);
  
  // データ検証を設定
  setupDataValidation(sheet);
  
  return sheet;
}

/**
 * データ検証を設定
 */
function setupDataValidation(sheet) {
  const lastRow = sheet.getLastRow() || 100;
  
  // Stage列のバリデーション
  const stageRange = sheet.getRange(2, COLUMNS.STAGE + 1, lastRow, 1);
  const stageRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(Object.values(STAGES))
    .setAllowInvalid(false)
    .build();
  stageRange.setDataValidation(stageRule);
  
  // ProjectStatus列のバリデーション
  const statusRange = sheet.getRange(2, COLUMNS.PROJECT_STATUS + 1, lastRow, 1);
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(Object.values(PROJECT_STATUS))
    .setAllowInvalid(false)
    .build();
  statusRange.setDataValidation(statusRule);
  
  // LastSentBy列のバリデーション
  const senderRange = sheet.getRange(2, COLUMNS.LAST_SENT_BY + 1, lastRow, 1);
  const senderRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(Object.values(SENDER_TYPE))
    .setAllowInvalid(false)
    .build();
  senderRange.setDataValidation(senderRule);
}

/**
 * 新規文書を追加
 */
function addDocument(docData) {
  const sheet = getOrCreateSheet();
  const lastRow = sheet.getLastRow();
  
  // デフォルト値を設定
  const now = new Date();
  const rowData = [
    docData.docKey || generateDocKey(),
    docData.rev || 'r1',
    docData.title || '',
    docData.stage || STAGES.DRAFT,
    docData.dueDate || '',
    docData.projectStatus || PROJECT_STATUS.OPEN,
    docData.lastSentBy || SENDER_TYPE.SELF,
    docData.googleDocUrl || '',
    docData.wordFileUrl || '',
    formatDate(now),
    formatDate(now),
    Session.getActiveUser().getEmail(),
    Session.getActiveUser().getEmail(),
    '',
    ''
  ];
  
  sheet.getRange(lastRow + 1, 1, 1, rowData.length).setValues([rowData]);
  
  return {
    success: true,
    docKey: rowData[COLUMNS.DOC_KEY],
    row: lastRow + 1
  };
}

/**
 * 文書を更新
 */
function updateDocument(docKey, updates) {
  const sheet = getOrCreateSheet();
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][COLUMNS.DOC_KEY] === docKey) {
      const now = new Date();
      
      // 更新可能なフィールドを処理
      if (updates.rev !== undefined) data[i][COLUMNS.REV] = updates.rev;
      if (updates.title !== undefined) data[i][COLUMNS.TITLE] = updates.title;
      if (updates.stage !== undefined) data[i][COLUMNS.STAGE] = updates.stage;
      if (updates.dueDate !== undefined) data[i][COLUMNS.DUE_DATE] = updates.dueDate;
      if (updates.projectStatus !== undefined) data[i][COLUMNS.PROJECT_STATUS] = updates.projectStatus;
      if (updates.lastSentBy !== undefined) data[i][COLUMNS.LAST_SENT_BY] = updates.lastSentBy;
      if (updates.googleDocUrl !== undefined) data[i][COLUMNS.GOOGLE_DOC_URL] = updates.googleDocUrl;
      if (updates.wordFileUrl !== undefined) data[i][COLUMNS.WORD_FILE_URL] = updates.wordFileUrl;
      if (updates.lastMailUrl !== undefined) data[i][COLUMNS.LAST_MAIL_URL] = updates.lastMailUrl;
      if (updates.mailTree !== undefined) data[i][COLUMNS.MAIL_TREE] = updates.mailTree;
      
      // 更新情報を記録
      data[i][COLUMNS.LAST_EDITED_AT] = formatDate(now);
      data[i][COLUMNS.LAST_EDITOR] = Session.getActiveUser().getEmail();
      
      // データを書き戻し
      sheet.getRange(i + 1, 1, 1, data[i].length).setValues([data[i]]);
      
      return {
        success: true,
        docKey: docKey,
        row: i + 1
      };
    }
  }
  
  return {
    success: false,
    error: 'Document not found'
  };
}

/**
 * 文書を検索
 */
function searchDocuments(criteria) {
  const sheet = getOrCreateSheet();
  const data = sheet.getDataRange().getValues();
  const results = [];
  
  for (let i = 1; i < data.length; i++) {
    let match = true;
    
    if (criteria.docKey && data[i][COLUMNS.DOC_KEY] !== criteria.docKey) match = false;
    if (criteria.stage && data[i][COLUMNS.STAGE] !== criteria.stage) match = false;
    if (criteria.projectStatus && data[i][COLUMNS.PROJECT_STATUS] !== criteria.projectStatus) match = false;
    if (criteria.lastSentBy && data[i][COLUMNS.LAST_SENT_BY] !== criteria.lastSentBy) match = false;
    
    if (match) {
      results.push({
        row: i + 1,
        docKey: data[i][COLUMNS.DOC_KEY],
        rev: data[i][COLUMNS.REV],
        title: data[i][COLUMNS.TITLE],
        stage: data[i][COLUMNS.STAGE],
        dueDate: data[i][COLUMNS.DUE_DATE],
        projectStatus: data[i][COLUMNS.PROJECT_STATUS],
        lastSentBy: data[i][COLUMNS.LAST_SENT_BY],
        googleDocUrl: data[i][COLUMNS.GOOGLE_DOC_URL],
        wordFileUrl: data[i][COLUMNS.WORD_FILE_URL],
        createdAt: data[i][COLUMNS.CREATED_AT],
        lastEditedAt: data[i][COLUMNS.LAST_EDITED_AT],
        ownerEmail: data[i][COLUMNS.OWNER_EMAIL],
        lastEditor: data[i][COLUMNS.LAST_EDITOR],
        lastMailUrl: data[i][COLUMNS.LAST_MAIL_URL],
        mailTree: data[i][COLUMNS.MAIL_TREE]
      });
    }
  }
  
  return results;
}

/**
 * 期限切れの文書を取得
 */
function getOverdueDocuments() {
  const sheet = getOrCreateSheet();
  const data = sheet.getDataRange().getValues();
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const results = [];
  
  for (let i = 1; i < data.length; i++) {
    const dueDate = data[i][COLUMNS.DUE_DATE];
    if (dueDate && new Date(dueDate) < today) {
      if (data[i][COLUMNS.PROJECT_STATUS] !== PROJECT_STATUS.CLOSED) {
        results.push({
          row: i + 1,
          docKey: data[i][COLUMNS.DOC_KEY],
          title: data[i][COLUMNS.TITLE],
          dueDate: dueDate,
          projectStatus: data[i][COLUMNS.PROJECT_STATUS],
          daysPastDue: Math.floor((today - new Date(dueDate)) / (1000 * 60 * 60 * 24))
        });
      }
    }
  }
  
  return results;
}

/**
 * シートを取得または作成
 */
function getOrCreateSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(SHEET_NAME);
  
  if (!sheet) {
    return initializeSheet();
  }
  
  return sheet;
}

/**
 * DocKeyを生成
 */
function generateDocKey() {
  const timestamp = new Date().getTime();
  const random = Math.floor(Math.random() * 1000);
  return `DOC${timestamp}${random}`;
}

/**
 * 日付をフォーマット
 */
function formatDate(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
}