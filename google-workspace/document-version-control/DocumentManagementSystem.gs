/**
 * 文書管理システム - 統合版
 * スプレッドシートとGASで完結する文書管理システム（Slack通知機能付き）
 * 
 * 使い方：
 * 1. このコードを新規Googleスプレッドシートのスクリプトエディタに貼り付け
 * 2. スプレッドシートを開き直すとメニューが表示される
 * 3. メニューから「初期設定」を実行
 */

// ========================================
// 定数定義
// ========================================

const SHEET_NAMES = {
  REGISTRY: 'DocRegistry',
  CONFIG: 'Config'
};

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
  MAIL_TREE: 14,
  FOLDER_ID: 15,
  FOLDER_URL: 16
};

// 必須項目の定義
const REQUIRED_COLUMNS = {
  DOC_KEY: true,
  TITLE: true,
  STAGE: true,
  PROJECT_STATUS: true,
  CREATED_AT: true,
  OWNER_EMAIL: true
};

const STAGES = {
  DRAFT: 'DRAFT',
  FOR_REVIEW: 'FOR-REVIEW',
  APPROVED: 'APPROVED',
  ARCHIVED: 'ARCHIVED'
};

const PROJECT_STATUS = {
  OPEN: 'OPEN',
  IN_PROGRESS: 'IN-PROGRESS',
  CLOSED: 'CLOSED',
  DELAYED: 'DELAYED'
};

const SENDER_TYPE = {
  SELF: 'SELF',
  PARTNER: 'PARTNER'
};

// ========================================
// メニューとUI
// ========================================

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  ui.createMenu('📄 文書管理システム')
    .addItem('📖 システム概要', 'showSystemOverview')
    .addItem('📚 使い方ガイド', 'showQuickGuide')
    .addSeparator()
    .addItem('🚀 初期設定', 'initializeSystem')
    .addSeparator()
    .addSubMenu(ui.createMenu('📋 文書操作')
      .addItem('新規文書を追加', 'showAddDocumentUI')
      .addItem('文書を検索', 'showSearchUI')
      .addItem('選択行を編集', 'editSelectedRow'))
    .addSeparator()
    .addSubMenu(ui.createMenu('🔄 文書同期')
      .addItem('選択行の文書を同期', 'syncSelectedDocument')
      .addItem('全文書を同期', 'syncAllDocumentsUI')
      .addItem('差分チェック', 'checkSelectedDocumentDifference')
      .addItem('添付ファイルをDriveに保存', 'saveSelectedAttachments')
      .addItem('親フォルダを設定', 'setDriveFolder'))
    .addSeparator()
    .addSubMenu(ui.createMenu('📧 メール連携')
      .addItem('選択行のメール情報を更新', 'updateSelectedRowEmailInfo')
      .addItem('全文書のメール情報を更新', 'updateAllEmailInfo'))
    .addSeparator()
    .addSubMenu(ui.createMenu('🔔 Slack通知')
      .addItem('Slack設定を確認', 'showSlackConfig')
      .addItem('テスト通知を送信', 'testSlackNotification')
      .addItem('期限切れ通知を送信', 'notifyOverdueDocuments')
      .addItem('週次サマリーを送信', 'sendWeeklySummary'))
    .addSeparator()
    .addSubMenu(ui.createMenu('📊 レポート')
      .addItem('期限切れ文書一覧', 'showOverdueReport')
      .addItem('ステータス別集計', 'showStatusReport'))
    .addSeparator()
    .addSubMenu(ui.createMenu('⏰ 自動化')
      .addItem('定期実行を設定', 'setupTriggers')
      .addItem('定期実行を停止', 'removeTriggers'))
    .addSeparator()
    .addItem('ℹ️ ヘルプ', 'showHelp')
    .addToUi();
  
  // 初回起動時のウェルカムメッセージ
  checkFirstRun();
}

// ========================================
// 初期設定
// ========================================

function initializeSystem() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    '初期設定',
    'システムの初期設定を行います。既存のデータは保持されます。続行しますか？',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) {
    return;
  }
  
  // 設定シートを作成
  initializeConfigSheet();
  
  // メインシートを作成
  initializeMainSheet();
  
  ui.alert('初期設定が完了しました。\n\nSlack通知を使用する場合は、Configシートで設定を行ってください。');
}

function initializeConfigSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(SHEET_NAMES.CONFIG);
  
  if (!sheet) {
    sheet = spreadsheet.insertSheet(SHEET_NAMES.CONFIG);
  }
  
  // 設定項目を設定
  const configData = [
    ['設定項目', '値', '説明'],
    ['PARENT_FOLDER_ID', '', '全プロジェクトの親フォルダID（空欄でマイドライブ直下）'],
    ['OPENAI_API_KEY', '', 'OpenAI APIキー（差分検出用）'],
    ['SLACK_WEBHOOK_URL', '', 'Slack Incoming Webhook URL'],
    ['SLACK_CHANNEL', '#general', '通知先チャンネル'],
    ['SLACK_USERNAME', '文書管理システム', 'Bot表示名'],
    ['SLACK_ICON_EMOJI', ':page_facing_up:', 'Botアイコン絵文字'],
    ['AUTO_EMAIL_UPDATE', 'TRUE', 'メール情報の自動更新（TRUE/FALSE）'],
    ['NOTIFY_ON_ADD', 'TRUE', '文書追加時に通知（TRUE/FALSE）'],
    ['NOTIFY_ON_STATUS_CHANGE', 'TRUE', 'ステータス変更時に通知（TRUE/FALSE）'],
    ['OVERDUE_CHECK_HOUR', '9', '期限切れチェック実行時刻（0-23）'],
    ['WEEKLY_SUMMARY_DAY', 'MONDAY', '週次サマリー送信曜日'],
    ['WEEKLY_SUMMARY_HOUR', '10', '週次サマリー送信時刻（0-23）']
  ];
  
  sheet.getRange(1, 1, configData.length, 3).setValues(configData);
  
  // ヘッダー行の書式設定
  const headerRange = sheet.getRange(1, 1, 1, 3);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#4285f4');
  headerRange.setFontColor('#ffffff');
  
  // 列幅調整
  sheet.setColumnWidth(1, 250);
  sheet.setColumnWidth(2, 300);
  sheet.setColumnWidth(3, 400);
  
  // 保護設定の説明を追加
  sheet.getRange(configData.length + 2, 1).setValue('※ 設定変更後は保存の必要はありません。自動的に反映されます。');
}

function initializeMainSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(SHEET_NAMES.REGISTRY);
  
  if (!sheet) {
    sheet = spreadsheet.insertSheet(SHEET_NAMES.REGISTRY);
  }
  
  // ヘッダー行を設定（必須項目には*を付ける）
  const headers = [
    'DocKey *',
    'Rev',
    'Title *',
    'Stage *',
    'DueDate',
    'ProjectStatus *',
    'LastSentBy',
    'GoogleDocURL',
    'WordFileURL',
    'CreatedAt *',
    'LastEditedAt',
    'OwnerEmail *',
    'LastEditor',
    'LastMailURL',
    'MailTree',
    'FolderID',
    'FolderURL'
  ];
  
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#f0f0f0');
  
  // 必須項目のヘッダーを赤色で強調
  for (let i = 0; i < headers.length; i++) {
    if (headers[i].includes('*')) {
      sheet.getRange(1, i + 1).setFontColor('#ff0000');
    }
  }
  
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
  sheet.setColumnWidth(COLUMNS.MAIL_TREE + 1, 300);
  sheet.setColumnWidth(COLUMNS.FOLDER_ID + 1, 150);
  sheet.setColumnWidth(COLUMNS.FOLDER_URL + 1, 200);
  
  // データ検証を設定
  setupDataValidation(sheet);
  
  // 条件付き書式を設定
  setupConditionalFormatting(sheet);
}

function setupDataValidation(sheet) {
  const lastRow = Math.max(sheet.getLastRow(), 100);
  
  // Stage列
  const stageRange = sheet.getRange(2, COLUMNS.STAGE + 1, lastRow, 1);
  const stageRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(Object.values(STAGES))
    .setAllowInvalid(false)
    .build();
  stageRange.setDataValidation(stageRule);
  
  // ProjectStatus列
  const statusRange = sheet.getRange(2, COLUMNS.PROJECT_STATUS + 1, lastRow, 1);
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(Object.values(PROJECT_STATUS))
    .setAllowInvalid(false)
    .build();
  statusRange.setDataValidation(statusRule);
  
  // LastSentBy列
  const senderRange = sheet.getRange(2, COLUMNS.LAST_SENT_BY + 1, lastRow, 1);
  const senderRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(Object.values(SENDER_TYPE))
    .setAllowInvalid(false)
    .build();
  senderRange.setDataValidation(senderRule);
}

function setupConditionalFormatting(sheet) {
  const lastRow = Math.max(sheet.getLastRow(), 100);
  
  // 期限切れを赤色でハイライト
  const dueDateRange = sheet.getRange(2, COLUMNS.DUE_DATE + 1, lastRow, 1);
  const today = new Date();
  const rule = SpreadsheetApp.newConditionalFormatRule()
    .whenDateBefore(today)
    .setBackground('#ffcccc')
    .setRanges([dueDateRange])
    .build();
  
  const rules = sheet.getConditionalFormatRules();
  rules.push(rule);
  sheet.setConditionalFormatRules(rules);
}

// ========================================
// 設定管理
// ========================================

function getConfig(key) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.CONFIG);
  if (!sheet) return null;
  
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === key) {
      return data[i][1];
    }
  }
  return null;
}

function setConfig(key, value) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.CONFIG);
  if (!sheet) return false;
  
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === key) {
      sheet.getRange(i + 1, 2).setValue(value);
      return true;
    }
  }
  return false;
}

// ========================================
// 文書管理機能
// ========================================

function addDocument(docData) {
  const sheet = getOrCreateSheet(SHEET_NAMES.REGISTRY);
  const lastRow = sheet.getLastRow();
  
  const now = new Date();
  const docKey = docData.docKey || generateDocKey();
  
  // 必須項目のチェック
  if (!docData.title) {
    SpreadsheetApp.getUi().alert('エラー: タイトルは必須項目です');
    return { success: false, error: 'タイトルは必須項目です' };
  }
  
  // DocKey用のフォルダを作成
  const folderInfo = createProjectFolder(docKey, docData.title);
  
  const rowData = [
    docKey,
    docData.rev || 'r1',
    docData.title,
    docData.stage || STAGES.DRAFT,
    docData.dueDate || '',
    docData.projectStatus || PROJECT_STATUS.OPEN,
    docData.lastSentBy || '',
    docData.googleDocUrl || '',
    docData.wordFileUrl || '',
    formatDate(now),
    formatDate(now),
    Session.getActiveUser().getEmail(),
    Session.getActiveUser().getEmail(),
    '',
    '',
    folderInfo.folderId,
    folderInfo.folderUrl
  ];
  
  sheet.getRange(lastRow + 1, 1, 1, rowData.length).setValues([rowData]);
  
  // Slack通知
  if (getConfig('NOTIFY_ON_ADD') === 'TRUE') {
    notifyNewDocument({
      docKey: rowData[COLUMNS.DOC_KEY],
      title: rowData[COLUMNS.TITLE],
      stage: rowData[COLUMNS.STAGE],
      dueDate: rowData[COLUMNS.DUE_DATE]
    });
  }
  
  return {
    success: true,
    docKey: rowData[COLUMNS.DOC_KEY],
    row: lastRow + 1,
    folderId: folderInfo.folderId,
    folderUrl: folderInfo.folderUrl
  };
}

/**
 * プロジェクト用フォルダを作成
 */
function createProjectFolder(docKey, title) {
  // 親フォルダを取得（全体設定から）
  let parentFolder;
  const parentFolderId = getConfig('PARENT_FOLDER_ID');
  
  if (parentFolderId) {
    try {
      parentFolder = DriveApp.getFolderById(parentFolderId);
    } catch (e) {
      console.error('親フォルダIDが無効です:', e);
      parentFolder = DriveApp.getRootFolder();
    }
  } else {
    parentFolder = DriveApp.getRootFolder();
  }
  
  // DocKey用のプロジェクトフォルダを作成
  const folderName = `${docKey}_${title.replace(/[\/\\:*?"<>|]/g, '_')}`;
  let projectFolder;
  
  // 既存のフォルダを検索
  const folders = parentFolder.getFoldersByName(folderName);
  if (folders.hasNext()) {
    projectFolder = folders.next();
  } else {
    projectFolder = parentFolder.createFolder(folderName);
    
    // サブフォルダを作成
    projectFolder.createFolder('WordFiles');
    projectFolder.createFolder('GoogleDocs');
    projectFolder.createFolder('DifferenceReports');
    projectFolder.createFolder('Attachments');
  }
  
  return {
    folderId: projectFolder.getId(),
    folderUrl: projectFolder.getUrl()
  };
}

function updateDocument(docKey, updates) {
  const sheet = getOrCreateSheet(SHEET_NAMES.REGISTRY);
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][COLUMNS.DOC_KEY] === docKey) {
      const oldStage = data[i][COLUMNS.STAGE];
      const oldStatus = data[i][COLUMNS.PROJECT_STATUS];
      const now = new Date();
      
      // 更新
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
      
      data[i][COLUMNS.LAST_EDITED_AT] = formatDate(now);
      data[i][COLUMNS.LAST_EDITOR] = Session.getActiveUser().getEmail();
      
      sheet.getRange(i + 1, 1, 1, data[i].length).setValues([data[i]]);
      
      // Slack通知
      if (getConfig('NOTIFY_ON_STATUS_CHANGE') === 'TRUE') {
        if (updates.stage && oldStage !== updates.stage) {
          notifyStatusChange(docKey, oldStage, updates.stage, data[i][COLUMNS.TITLE]);
        }
        if (updates.projectStatus === PROJECT_STATUS.CLOSED && oldStatus !== PROJECT_STATUS.CLOSED) {
          notifyProjectCompletion(docKey, data[i][COLUMNS.TITLE]);
        }
      }
      
      return { success: true, docKey: docKey, row: i + 1 };
    }
  }
  
  return { success: false, error: 'Document not found' };
}

function searchDocuments(criteria) {
  const sheet = getOrCreateSheet(SHEET_NAMES.REGISTRY);
  const data = sheet.getDataRange().getValues();
  const results = [];
  
  for (let i = 1; i < data.length; i++) {
    let match = true;
    
    if (criteria.docKey && !data[i][COLUMNS.DOC_KEY].toString().toLowerCase().includes(criteria.docKey.toLowerCase())) match = false;
    if (criteria.title && !data[i][COLUMNS.TITLE].toString().toLowerCase().includes(criteria.title.toLowerCase())) match = false;
    if (criteria.stage && data[i][COLUMNS.STAGE] !== criteria.stage) match = false;
    if (criteria.projectStatus && data[i][COLUMNS.PROJECT_STATUS] !== criteria.projectStatus) match = false;
    
    if (match) {
      results.push({
        row: i + 1,
        docKey: data[i][COLUMNS.DOC_KEY],
        rev: data[i][COLUMNS.REV],
        title: data[i][COLUMNS.TITLE],
        stage: data[i][COLUMNS.STAGE],
        dueDate: data[i][COLUMNS.DUE_DATE],
        projectStatus: data[i][COLUMNS.PROJECT_STATUS],
        lastSentBy: data[i][COLUMNS.LAST_SENT_BY]
      });
    }
  }
  
  return results;
}

function getOverdueDocuments() {
  const sheet = getOrCreateSheet(SHEET_NAMES.REGISTRY);
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

// ========================================
// Gmail連携とドキュメント管理
// ========================================

function searchRelatedEmails(docKey, title) {
  const queries = [];
  
  if (docKey) queries.push(`"${docKey}"`);
  if (title) queries.push(`"${title}"`);
  
  if (queries.length === 0) return [];
  
  const searchQuery = queries.join(' OR ');
  const threads = GmailApp.search(searchQuery, 0, 50);
  const emails = [];
  
  threads.forEach(thread => {
    const messages = thread.getMessages();
    messages.forEach(message => {
      const attachments = message.getAttachments();
      emails.push({
        id: message.getId(),
        threadId: thread.getId(),
        subject: message.getSubject(),
        from: message.getFrom(),
        to: message.getTo(),
        date: message.getDate(),
        body: message.getPlainBody().substring(0, 200),
        hasAttachments: attachments.length > 0,
        attachments: attachments.map(att => ({
          name: att.getName(),
          type: att.getContentType(),
          size: att.getSize()
        })),
        messageUrl: `https://mail.google.com/mail/u/0/#inbox/${message.getId()}`
      });
    });
  });
  
  return emails;
}

/**
 * メール添付ファイルをDriveに保存
 * 重要: 添付ファイル名にDocKeyが含まれている場合のみ保存
 * これにより、関係のないファイルが誤って保存されることを防ぐ
 */
function saveAttachmentsToDrive(docKey) {
  // DocKeyでメールを検索（メールの件名か本文にDocKeyが含まれるもの）
  const emails = searchRelatedEmails(docKey, '');
  const savedFiles = [];
  
  // DocKey専用フォルダを取得
  const projectFolder = getProjectFolder(docKey);
  
  if (!projectFolder) {
    console.error('プロジェクトフォルダが見つかりません:', docKey);
    return savedFiles;
  }
  
  // Attachmentsサブフォルダを取得
  let attachmentsFolder;
  const attachmentsFolders = projectFolder.getFoldersByName('Attachments');
  if (attachmentsFolders.hasNext()) {
    attachmentsFolder = attachmentsFolders.next();
  } else {
    attachmentsFolder = projectFolder.createFolder('Attachments');
  }
  
  emails.forEach(email => {
    if (email.hasAttachments) {
      const threads = GmailApp.search(`rfc822msgid:${email.id}`, 0, 1);
      if (threads.length > 0) {
        const messages = threads[0].getMessages();
        messages.forEach(message => {
          if (message.getId() === email.id) {
            const attachments = message.getAttachments();
            attachments.forEach(attachment => {
              const attachmentName = attachment.getName();
              
              // 添付ファイル名にDocKeyが含まれているかチェック
              // DocKeyがファイル名に含まれている場合のみ保存
              if (!attachmentName.toUpperCase().includes(docKey.toUpperCase())) {
                console.log(`スキップ: ファイル名にDocKey "${docKey}" が含まれていません: ${attachmentName}`);
                return; // このファイルはスキップ
              }
              
              // サポートする文書形式をチェック
              const isWordDoc = attachment.getContentType().includes('word') || 
                  attachmentName.endsWith('.docx') || 
                  attachmentName.endsWith('.doc');
              
              const isPDF = attachment.getContentType().includes('pdf') || 
                  attachmentName.endsWith('.pdf');
              
              const isExcel = attachment.getContentType().includes('excel') || 
                  attachment.getContentType().includes('spreadsheet') ||
                  attachmentName.endsWith('.xlsx') || 
                  attachmentName.endsWith('.xls');
              
              // サポートする形式の場合のみ保存
              if (isWordDoc || isPDF || isExcel) {
                
                // ファイル名を生成（統一命名規則）
                const emailDate = email.date;
                const dateStr = formatFileDateTime(emailDate);
                const stage = getCurrentStage(docKey);
                const rev = getCurrentRevision(docKey);
                const extension = attachmentName.split('.').pop();
                
                // DocKey_YYYYMMDD_HHMM_バージョン_ステージ.拡張子
                const fileName = `${docKey}_${dateStr}_${rev}_${stage}.${extension}`;
                
                const existingFiles = attachmentsFolder.getFilesByName(fileName);
                
                if (!existingFiles.hasNext()) {
                  const blob = attachment.copyBlob();
                  const file = attachmentsFolder.createFile(blob);
                  file.setName(fileName);
                  
                  // ファイルのプロパティにメタデータを保存
                  file.setDescription(JSON.stringify({
                    originalName: attachment.getName(),
                    receivedDate: emailDate.toISOString(),
                    emailFrom: email.from,
                    emailSubject: email.subject,
                    docKey: docKey,
                    stage: stage,
                    revision: rev
                  }));
                  
                  // Wordファイルの場合はWordFilesフォルダにもコピー
                  if (attachment.getContentType().includes('word') || 
                      attachment.getName().endsWith('.docx') || 
                      attachment.getName().endsWith('.doc')) {
                    const wordFolder = getProjectSubFolder(projectFolder, 'WordFiles');
                    if (wordFolder) {
                      const wordCopy = file.makeCopy(fileName, wordFolder);
                      
                      // 最新のWordファイルとして記録
                      updateDocument(docKey, {
                        wordFileUrl: wordCopy.getUrl(),
                        lastMailUrl: email.messageUrl
                      });
                    }
                  }
                  
                  savedFiles.push({
                    name: fileName,
                    url: file.getUrl(),
                    id: file.getId(),
                    receivedDate: emailDate,
                    originalName: attachment.getName()
                  });
                }
              }
            });
          }
        });
      }
    }
  });
  
  return savedFiles;
}

/**
 * WordファイルからGoogle Documentを作成
 */
function createGoogleDocFromWord(wordFileId, docKey, title) {
  try {
    const wordFile = DriveApp.getFileById(wordFileId);
    const blob = wordFile.getBlob();
    
    // Google Documentとして変換
    const resource = {
      title: `${docKey}_${title}`,
      mimeType: MimeType.GOOGLE_DOCS
    };
    
    const file = Drive.Files.insert(resource, blob, {
      convert: true
    });
    
    return {
      id: file.id,
      url: file.alternateLink || `https://docs.google.com/document/d/${file.id}/edit`
    };
  } catch (e) {
    console.error('Error creating Google Doc from Word:', e);
    return null;
  }
}

/**
 * 文書の自動同期と差分チェック
 */
function syncDocuments(docKey) {
  const sheet = getOrCreateSheet(SHEET_NAMES.REGISTRY);
  const data = sheet.getDataRange().getValues();
  let docRow = -1;
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][COLUMNS.DOC_KEY] === docKey) {
      docRow = i;
      break;
    }
  }
  
  if (docRow === -1) return { success: false, error: 'Document not found' };
  
  const updates = {};
  
  // 1. メール添付ファイルをDriveに保存
  const savedFiles = saveAttachmentsToDrive(docKey);
  
  if (savedFiles.length > 0) {
    // 最新のWordファイルを取得
    const latestWordFile = savedFiles[savedFiles.length - 1];
    updates.wordFileUrl = latestWordFile.url;
    
    // 2. Google Documentを作成または更新
    const googleDocInfo = createGoogleDocFromWord(
      latestWordFile.id,
      docKey,
      data[docRow][COLUMNS.TITLE]
    );
    
    if (googleDocInfo) {
      updates.googleDocUrl = googleDocInfo.url;
    }
  }
  
  // 3. 既存のGoogle DocとWordの差分をチェック
  const diffStatus = checkDocumentDifference(docKey);
  if (diffStatus) {
    // 差分がある場合はSlack通知
    if (diffStatus.hasDifference && getConfig('NOTIFY_ON_STATUS_CHANGE') === 'TRUE') {
      notifyDocumentDifference(docKey, data[docRow][COLUMNS.TITLE], diffStatus);
    }
  }
  
  // 4. 更新を反映
  if (Object.keys(updates).length > 0) {
    updateDocument(docKey, updates);
  }
  
  return {
    success: true,
    savedFiles: savedFiles.length,
    updates: updates,
    diffStatus: diffStatus
  };
}

/**
 * Google DocとWordファイルの内容差分チェック
 */
function checkDocumentDifference(docKey) {
  const sheet = getOrCreateSheet(SHEET_NAMES.REGISTRY);
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][COLUMNS.DOC_KEY] === docKey) {
      const googleDocUrl = data[i][COLUMNS.GOOGLE_DOC_URL];
      const wordFileUrl = data[i][COLUMNS.WORD_FILE_URL];
      const title = data[i][COLUMNS.TITLE];
      
      if (!googleDocUrl || !wordFileUrl) {
        return { hasDifference: false, reason: 'Missing document URLs' };
      }
      
      try {
        // ファイルのメタデータを取得
        const googleDocId = googleDocUrl.match(/\/d\/([a-zA-Z0-9-_]+)/)[1];
        const wordFileId = wordFileUrl.match(/\/d\/([a-zA-Z0-9-_]+)/)[1];
        
        const googleDocFile = DriveApp.getFileById(googleDocId);
        const wordFile = DriveApp.getFileById(wordFileId);
        
        const googleDocModified = googleDocFile.getLastUpdated();
        const wordFileModified = wordFile.getLastUpdated();
        
        // Google Docのテキストを取得
        const googleDocText = getGoogleDocText(googleDocId);
        
        // Wordファイルをテキストに変換
        const wordText = getWordFileText(wordFileId);
        
        // OpenAI APIで差分を検出
        const diffAnalysis = analyzeDocumentDifference(googleDocText, wordText, docKey);
        
        const diffResult = {
          hasDifference: diffAnalysis.hasDifference,
          differences: diffAnalysis.differences,
          summary: diffAnalysis.summary,
          docKey: docKey,
          title: title,
          googleDocUrl: googleDocUrl,
          wordFileUrl: wordFileUrl,
          metadata: {
            googleDocModified: googleDocModified,
            wordFileModified: wordFileModified,
            googleDocSize: googleDocFile.getSize(),
            wordFileSize: wordFile.getSize(),
            timeDifference: Math.abs(googleDocModified - wordFileModified)
          }
        };
        
        // 差分レポートを保存
        if (diffAnalysis.hasDifference) {
          saveDifferenceReport(docKey, title, diffResult);
        }
        
        return diffResult;
      } catch (e) {
        console.error('Error checking document difference:', e);
        return { hasDifference: false, error: e.toString() };
      }
    }
  }
  
  return null;
}

/**
 * Google Documentのテキストを取得
 */
function getGoogleDocText(docId) {
  try {
    const doc = DocumentApp.openById(docId);
    return doc.getBody().getText();
  } catch (e) {
    console.error('Error getting Google Doc text:', e);
    return '';
  }
}

/**
 * Wordファイルをテキストに変換
 */
function getWordFileText(fileId) {
  try {
    const file = DriveApp.getFileById(fileId);
    const blob = file.getBlob();
    
    // WordファイルをGoogle Docに一時変換してテキストを取得
    const resource = {
      title: 'temp_conversion',
      mimeType: MimeType.GOOGLE_DOCS
    };
    
    const convertedFile = Drive.Files.insert(resource, blob, {
      convert: true
    });
    
    const tempDoc = DocumentApp.openById(convertedFile.id);
    const text = tempDoc.getBody().getText();
    
    // 一時ファイルを削除
    DriveApp.getFileById(convertedFile.id).setTrashed(true);
    
    return text;
  } catch (e) {
    console.error('Error converting Word file to text:', e);
    return '';
  }
}

/**
 * 文書テキストからバージョン情報やステータスなどのメタデータを除去
 */
function cleanDocumentText(text) {
  if (!text) return '';
  
  let cleanText = text;
  
  // バージョン番号のパターンを除去（r1, r2, rev1, version1, v1.0など）
  cleanText = cleanText.replace(/\b[rR]ev(ision)?\s*\d+/g, '');
  cleanText = cleanText.replace(/\b[vV]ersion\s*\d+(\.\d+)*/g, '');
  cleanText = cleanText.replace(/\b[vV]\d+(\.\d+)*/g, '');
  cleanText = cleanText.replace(/\b[rR]\d+/g, '');
  
  // ステータス情報を除去（DRAFT, APPROVED, FOR-REVIEW, FINAL等）
  cleanText = cleanText.replace(/\b(DRAFT|APPROVED|FOR[-_]?REVIEW|FINAL|ARCHIVED|PENDING|WIP)\b/gi, '');
  
  // 日付パターンを正規化（ヘッダーやフッターの日付は除去）
  // ただし本文中の重要な日付は残す
  cleanText = cleanText.replace(/^.*\d{4}[-\/]\d{1,2}[-\/]\d{1,2}.*$/gm, (match) => {
    // 行頭や行末にある日付（ヘッダー/フッター）は除去
    if (match.match(/^[\s\-_]*\d{4}[-\/]\d{1,2}[-\/]\d{1,2}[\s\-_]*$/)) {
      return '';
    }
    return match;
  });
  
  // DocKey付きのファイル名パターンを除去
  cleanText = cleanText.replace(/[A-Z]+\d+_\d{8}_\d{4}_[^_\s]+_[^_\s]+/g, '');
  
  // ページ番号を除去
  cleanText = cleanText.replace(/^\s*[-\d]+\s*$/gm, '');
  cleanText = cleanText.replace(/\bpage\s+\d+\b/gi, '');
  cleanText = cleanText.replace(/\b\d+\s*\/\s*\d+\b/g, ''); // 1/5 のようなページ表記
  
  // 複数の空白行を単一の改行に変換
  cleanText = cleanText.replace(/\n\s*\n\s*\n+/g, '\n\n');
  
  // 行頭・行末の空白を除去
  cleanText = cleanText.split('\n').map(line => line.trim()).join('\n');
  
  // 最初と最後の空白行を除去
  cleanText = cleanText.trim();
  
  return cleanText;
}

/**
 * OpenAI APIを使用して文書の差分を分析
 */
function analyzeDocumentDifference(googleDocText, wordText, docKey) {
  // テキストから余計な情報を除去（バージョン情報、日付、ステータスなど）
  const cleanGoogleText = cleanDocumentText(googleDocText);
  const cleanWordText = cleanDocumentText(wordText);
  
  const apiKey = getConfig('OPENAI_API_KEY');
  
  if (!apiKey) {
    // APIキーがない場合は簡易比較
    const hasDifference = cleanGoogleText !== cleanWordText;
    return {
      hasDifference: hasDifference,
      differences: hasDifference ? ['テキストに差分があります（詳細分析にはOpenAI APIキーが必要です）'] : [],
      summary: hasDifference ? '文書間に差分が検出されました' : '文書は同一です'
    };
  }
  
  const prompt = `
以下の2つの文書の本文内容を比較し、実質的な差分のみを分析してください。
以下は無視してください：
- ファイル名やヘッダー/フッターの違い
- 日付形式の違い（内容が同じ場合）
- バージョン番号やステータス表記
- 改行やスペースの違い
- ページ番号

重要な内容の差分のみを報告してください。

【Google Document】
${cleanGoogleText.substring(0, 3000)}

【Word Document】
${cleanWordText.substring(0, 3000)}

以下の形式でJSONで回答してください：
{
  "hasDifference": true/false,
  "differences": ["内容の差分1", "内容の差分2", ...],
  "summary": "実質的な差分の要約"
}
`;
  
  try {
    const response = UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions', {
      method: 'post',
      headers: {
        'Authorization': `Bearer ${apiKey}`,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify({
        model: 'gpt-5',
        messages: [{
          role: 'user',
          content: prompt
        }],
        response_format: { type: 'json_object' },
        temperature: 0.3,
        max_tokens: 1000
      })
    });
    
    const result = JSON.parse(response.getContentText());
    const analysis = JSON.parse(result.choices[0].message.content);
    
    return analysis;
  } catch (e) {
    console.error('OpenAI API error:', e);
    
    // APIエラーの場合は簡易比較にフォールバック
    const hasDifference = googleDocText.trim() !== wordText.trim();
    return {
      hasDifference: hasDifference,
      differences: hasDifference ? ['APIエラー: 簡易比較で差分が検出されました'] : [],
      summary: hasDifference ? '文書間に差分が検出されました' : '文書は同一です'
    };
  }
}

/**
 * 差分レポートを保存
 */
function saveDifferenceReport(docKey, title, diffResult) {
  // DocKey専用フォルダを取得
  const projectFolder = getProjectFolder(docKey);
  
  if (!projectFolder) {
    console.error('プロジェクトフォルダが見つかりません:', docKey);
    return null;
  }
  
  // DifferenceReportsサブフォルダを取得
  const reportFolder = getProjectSubFolder(projectFolder, 'DifferenceReports');
  
  // レポート内容を作成
  const now = new Date();
  let reportContent = `文書差分分析レポート
================================================================
生成日時: ${formatDateTime(now)}
分析方法: ${diffResult.differences ? 'AI内容分析' : '簡易比較'}

【文書情報】
DocKey: ${docKey}
タイトル: ${title}

【比較対象】
Google Document: ${diffResult.googleDocUrl || 'N/A'}
Word File: ${diffResult.wordFileUrl || 'N/A'}

【分析結果】
差分検出: ${diffResult.hasDifference ? '★ 差分あり' : '✓ 差分なし'}

【差分サマリー】
${diffResult.summary || '差分なし'}

【検出された差分詳細】
${diffResult.differences && diffResult.differences.length > 0 ? 
  diffResult.differences.map((diff, index) => `${index + 1}. ${diff}`).join('\n') : 
  '差分は検出されませんでした'}

【推奨アクション】
${diffResult.hasDifference ? 
  '1. 上記の差分を確認し、どちらが最新版か判断してください\n' +
  '2. 必要に応じて文書を同期してください\n' +
  '3. 最新版をマスターとして確定してください\n' +
  '4. ステージとバージョンを更新してください' : 
  '• 文書は同一内容です\n' +
  '• 追加のアクションは不要です'}

【メタデータ】
レポートID: ${Utilities.getUuid()}
分析実行者: ${Session.getActiveUser().getEmail()}

================================================================
このレポートは文書管理システムによって自動生成されました
バージョン: 1.0`;
  
  // レポートファイルを作成
  const fileName = `DiffReport_${docKey}_${formatDate(now)}_${now.getHours()}${String(now.getMinutes()).padStart(2, '0')}.txt`;
  const blob = Utilities.newBlob(reportContent, 'text/plain', fileName);
  const file = reportFolder.createFile(blob);
  
  console.log(`差分レポートを保存: ${file.getUrl()}`);
  
  return {
    fileName: fileName,
    fileUrl: file.getUrl(),
    folderId: reportFolder.getId()
  };
}

/**
 * 全文書の同期
 */
function syncAllDocuments() {
  const sheet = getOrCreateSheet(SHEET_NAMES.REGISTRY);
  const data = sheet.getDataRange().getValues();
  let syncedCount = 0;
  const results = [];
  
  for (let i = 1; i < data.length; i++) {
    const docKey = data[i][COLUMNS.DOC_KEY];
    if (docKey) {
      const result = syncDocuments(docKey);
      if (result.success) {
        syncedCount++;
        results.push({
          docKey: docKey,
          savedFiles: result.savedFiles,
          hasDifference: result.diffStatus?.hasDifference
        });
      }
    }
  }
  
  return {
    success: true,
    syncedCount: syncedCount,
    results: results
  };
}

function getLatestEmailInfo(docKey, title) {
  const emails = searchRelatedEmails(docKey, title);
  
  if (emails.length === 0) return null;
  
  emails.sort((a, b) => b.date - a.date);
  const latestEmail = emails[0];
  
  const myEmail = Session.getActiveUser().getEmail();
  const lastSentBy = latestEmail.from.includes(myEmail) ? SENDER_TYPE.SELF : SENDER_TYPE.PARTNER;
  
  return {
    lastMailUrl: latestEmail.messageUrl,
    lastSentBy: lastSentBy,
    lastMailDate: latestEmail.date,
    subject: latestEmail.subject,
    from: latestEmail.from,
    to: latestEmail.to
  };
}

function buildMailTree(docKey, title) {
  const emails = searchRelatedEmails(docKey, title);
  
  if (emails.length === 0) return '';
  
  const threads = {};
  emails.forEach(email => {
    if (!threads[email.threadId]) {
      threads[email.threadId] = [];
    }
    threads[email.threadId].push(email);
  });
  
  let tree = [];
  Object.keys(threads).forEach(threadId => {
    const threadEmails = threads[threadId];
    threadEmails.sort((a, b) => a.date - b.date);
    
    threadEmails.forEach((email, index) => {
      const indent = '  '.repeat(index);
      const dateStr = formatDateTime(email.date);
      tree.push(`${indent}[${dateStr}] ${email.from} → ${email.subject}`);
    });
  });
  
  return tree.join('\n');
}

function updateEmailInfoForDocument(docKey) {
  const sheet = getOrCreateSheet(SHEET_NAMES.REGISTRY);
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][COLUMNS.DOC_KEY] === docKey) {
      const emailInfo = getLatestEmailInfo(docKey, data[i][COLUMNS.TITLE]);
      
      if (emailInfo) {
        const mailTree = buildMailTree(docKey, data[i][COLUMNS.TITLE]);
        
        return updateDocument(docKey, {
          lastSentBy: emailInfo.lastSentBy,
          lastMailUrl: emailInfo.lastMailUrl,
          mailTree: mailTree
        });
      }
      break;
    }
  }
  
  return { success: false, error: 'No related emails found' };
}

function updateAllEmailInfo() {
  const sheet = getOrCreateSheet(SHEET_NAMES.REGISTRY);
  const data = sheet.getDataRange().getValues();
  let updatedCount = 0;
  
  for (let i = 1; i < data.length; i++) {
    const docKey = data[i][COLUMNS.DOC_KEY];
    const title = data[i][COLUMNS.TITLE];
    
    if (docKey) {
      const emailInfo = getLatestEmailInfo(docKey, title);
      
      if (emailInfo) {
        const mailTree = buildMailTree(docKey, title);
        
        data[i][COLUMNS.LAST_SENT_BY] = emailInfo.lastSentBy;
        data[i][COLUMNS.LAST_MAIL_URL] = emailInfo.lastMailUrl;
        data[i][COLUMNS.MAIL_TREE] = mailTree;
        
        updatedCount++;
      }
    }
  }
  
  if (updatedCount > 0) {
    sheet.getDataRange().setValues(data);
  }
  
  SpreadsheetApp.getUi().alert(`${updatedCount}件の文書のメール情報を更新しました`);
  
  return { success: true, updatedCount: updatedCount };
}

function updateSelectedRowEmailInfo() {
  const sheet = SpreadsheetApp.getActiveSheet();
  if (sheet.getName() !== SHEET_NAMES.REGISTRY) {
    SpreadsheetApp.getUi().alert('DocRegistryシートで実行してください');
    return;
  }
  
  const activeRange = sheet.getActiveRange();
  const row = activeRange.getRow();
  
  if (row <= 1) {
    SpreadsheetApp.getUi().alert('データ行を選択してください');
    return;
  }
  
  const docKey = sheet.getRange(row, COLUMNS.DOC_KEY + 1).getValue();
  
  if (!docKey) {
    SpreadsheetApp.getUi().alert('DocKeyが見つかりません');
    return;
  }
  
  const result = updateEmailInfoForDocument(docKey);
  
  if (result.success) {
    SpreadsheetApp.getUi().alert('メール情報を更新しました');
  } else {
    SpreadsheetApp.getUi().alert('更新に失敗しました: ' + result.error);
  }
}

// ========================================
// Slack通知
// ========================================

function sendToSlack(message, attachments = []) {
  const webhookUrl = getConfig('SLACK_WEBHOOK_URL');
  
  if (!webhookUrl) {
    console.error('Slack Webhook URLが設定されていません');
    return false;
  }
  
  const payload = {
    channel: getConfig('SLACK_CHANNEL') || '#general',
    username: getConfig('SLACK_USERNAME') || '文書管理システム',
    icon_emoji: getConfig('SLACK_ICON_EMOJI') || ':page_facing_up:',
    text: message,
    attachments: attachments
  };
  
  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  
  try {
    const response = UrlFetchApp.fetch(webhookUrl, options);
    return response.getResponseCode() === 200;
  } catch (e) {
    console.error('Slack送信エラー:', e);
    return false;
  }
}

function notifyNewDocument(docData) {
  const message = '新しい文書が追加されました';
  
  const attachment = {
    fallback: message,
    color: 'good',
    title: '新規文書',
    fields: [
      { title: 'DocKey', value: docData.docKey, short: true },
      { title: 'タイトル', value: docData.title, short: true },
      { title: 'ステージ', value: docData.stage, short: true },
      { title: '期限', value: docData.dueDate || '未設定', short: true },
      { title: '作成者', value: Session.getActiveUser().getEmail(), short: true }
    ],
    footer: '文書管理システム',
    ts: Math.floor(Date.now() / 1000)
  };
  
  return sendToSlack(message, [attachment]);
}

function notifyStatusChange(docKey, oldStatus, newStatus, docTitle) {
  const message = `文書のステータスが変更されました`;
  
  const statusColors = {
    'DRAFT': '#808080',
    'FOR-REVIEW': '#FFA500',
    'APPROVED': '#008000',
    'ARCHIVED': '#4B0082'
  };
  
  const attachment = {
    fallback: message,
    color: statusColors[newStatus] || 'warning',
    title: 'ステータス変更',
    fields: [
      { title: 'DocKey', value: docKey, short: true },
      { title: 'タイトル', value: docTitle, short: false },
      { title: '変更前', value: oldStatus, short: true },
      { title: '変更後', value: newStatus, short: true },
      { title: '変更者', value: Session.getActiveUser().getEmail(), short: true }
    ],
    footer: '文書管理システム',
    ts: Math.floor(Date.now() / 1000)
  };
  
  return sendToSlack(message, [attachment]);
}

function notifyOverdueDocuments() {
  const overdueDocuments = getOverdueDocuments();
  
  if (overdueDocuments.length === 0) {
    SpreadsheetApp.getUi().alert('期限切れの文書はありません');
    return;
  }
  
  const message = `⚠️ ${overdueDocuments.length}件の文書が期限切れです`;
  
  const attachments = overdueDocuments.slice(0, 5).map(doc => ({
    fallback: `${doc.docKey}: ${doc.title} - ${doc.daysPastDue}日超過`,
    color: doc.daysPastDue > 7 ? 'danger' : 'warning',
    title: doc.title,
    fields: [
      { title: 'DocKey', value: doc.docKey, short: true },
      { title: '期限', value: doc.dueDate, short: true },
      { title: '超過日数', value: `${doc.daysPastDue}日`, short: true },
      { title: 'ステータス', value: doc.projectStatus, short: true }
    ]
  }));
  
  if (overdueDocuments.length > 5) {
    attachments.push({
      fallback: `他${overdueDocuments.length - 5}件`,
      color: '#808080',
      text: `他${overdueDocuments.length - 5}件の期限切れ文書があります`
    });
  }
  
  const result = sendToSlack(message, attachments);
  
  if (result) {
    SpreadsheetApp.getUi().alert('期限切れ通知を送信しました');
  } else {
    SpreadsheetApp.getUi().alert('通知の送信に失敗しました');
  }
}

function notifyProjectCompletion(docKey, docTitle) {
  const message = '✅ プロジェクトが完了しました';
  
  const attachment = {
    fallback: message,
    color: 'good',
    title: 'プロジェクト完了',
    fields: [
      { title: 'DocKey', value: docKey, short: true },
      { title: 'タイトル', value: docTitle, short: false },
      { title: '完了日', value: formatDate(new Date()), short: true },
      { title: '完了者', value: Session.getActiveUser().getEmail(), short: true }
    ],
    footer: '文書管理システム',
    ts: Math.floor(Date.now() / 1000)
  };
  
  return sendToSlack(message, [attachment]);
}

/**
 * 文書の差分を通知
 */
function notifyDocumentDifference(docKey, docTitle, diffStatus) {
  const message = '⚠️ 文書に差分が検出されました';
  
  const googleModified = formatDateTime(new Date(diffStatus.googleDocModified));
  const wordModified = formatDateTime(new Date(diffStatus.wordFileModified));
  const timeDiffMinutes = Math.floor(diffStatus.timeDifference / 60000);
  
  const attachment = {
    fallback: message,
    color: 'warning',
    title: '文書差分検出',
    fields: [
      { title: 'DocKey', value: docKey, short: true },
      { title: 'タイトル', value: docTitle, short: false },
      { title: 'Google Doc更新', value: googleModified, short: true },
      { title: 'Word更新', value: wordModified, short: true },
      { title: '差分', value: `${timeDiffMinutes}分`, short: true },
      { title: '推奨アクション', value: '文書を確認して同期してください', short: false }
    ],
    footer: '文書管理システム',
    ts: Math.floor(Date.now() / 1000)
  };
  
  return sendToSlack(message, [attachment]);
}

/**
 * 文書同期完了を通知
 */
function notifyDocumentSync(docKey, docTitle, syncResult) {
  const message = '🔄 文書の同期が完了しました';
  
  const attachment = {
    fallback: message,
    color: 'good',
    title: '文書同期完了',
    fields: [
      { title: 'DocKey', value: docKey, short: true },
      { title: 'タイトル', value: docTitle, short: false },
      { title: '保存ファイル数', value: syncResult.savedFiles.toString(), short: true },
      { title: '更新項目', value: Object.keys(syncResult.updates).join(', ') || 'なし', short: true }
    ],
    footer: '文書管理システム',
    ts: Math.floor(Date.now() / 1000)
  };
  
  if (syncResult.diffStatus?.hasDifference) {
    attachment.fields.push({
      title: '⚠️ 注意',
      value: 'Google DocとWordファイルに差分があります',
      short: false
    });
  }
  
  return sendToSlack(message, [attachment]);
}

function sendWeeklySummary() {
  const sheet = getOrCreateSheet(SHEET_NAMES.REGISTRY);
  const data = sheet.getDataRange().getValues();
  
  const stats = {
    total: data.length - 1,
    draft: 0,
    forReview: 0,
    approved: 0,
    archived: 0,
    open: 0,
    inProgress: 0,
    closed: 0,
    delayed: 0,
    overdue: 0
  };
  
  const today = new Date();
  
  for (let i = 1; i < data.length; i++) {
    switch (data[i][COLUMNS.STAGE]) {
      case STAGES.DRAFT: stats.draft++; break;
      case STAGES.FOR_REVIEW: stats.forReview++; break;
      case STAGES.APPROVED: stats.approved++; break;
      case STAGES.ARCHIVED: stats.archived++; break;
    }
    
    switch (data[i][COLUMNS.PROJECT_STATUS]) {
      case PROJECT_STATUS.OPEN: stats.open++; break;
      case PROJECT_STATUS.IN_PROGRESS: stats.inProgress++; break;
      case PROJECT_STATUS.CLOSED: stats.closed++; break;
      case PROJECT_STATUS.DELAYED: stats.delayed++; break;
    }
    
    const dueDate = data[i][COLUMNS.DUE_DATE];
    if (dueDate && new Date(dueDate) < today) {
      stats.overdue++;
    }
  }
  
  const message = '📊 週次文書管理レポート';
  
  const attachment = {
    fallback: message,
    color: '#36a64f',
    title: '週次サマリー',
    pretext: `${formatDate(today)} 時点の文書管理状況`,
    fields: [
      { title: '総文書数', value: stats.total.toString(), short: true },
      { title: '期限切れ', value: stats.overdue > 0 ? `⚠️ ${stats.overdue}件` : '0件', short: true },
      { title: '文書ステージ', value: `下書き: ${stats.draft}\nレビュー中: ${stats.forReview}\n承認済み: ${stats.approved}\nアーカイブ: ${stats.archived}`, short: true },
      { title: 'プロジェクト状況', value: `オープン: ${stats.open}\n進行中: ${stats.inProgress}\n完了: ${stats.closed}\n遅延: ${stats.delayed}`, short: true }
    ],
    footer: '文書管理システム',
    ts: Math.floor(Date.now() / 1000)
  };
  
  const result = sendToSlack(message, [attachment]);
  
  if (result) {
    SpreadsheetApp.getUi().alert('週次サマリーを送信しました');
  } else {
    SpreadsheetApp.getUi().alert('送信に失敗しました');
  }
}

function testSlackNotification() {
  const message = '🔔 Slack連携テスト';
  
  const attachment = {
    fallback: message,
    color: 'good',
    title: 'テスト通知',
    text: 'Slack通知機能が正常に動作しています',
    fields: [
      { title: 'テスト実行者', value: Session.getActiveUser().getEmail(), short: true },
      { title: '実行時刻', value: formatDateTime(new Date()), short: true }
    ],
    footer: '文書管理システム',
    ts: Math.floor(Date.now() / 1000)
  };
  
  const result = sendToSlack(message, [attachment]);
  
  if (result) {
    SpreadsheetApp.getUi().alert('Slackテスト通知を送信しました');
  } else {
    SpreadsheetApp.getUi().alert('Slack通知の送信に失敗しました。Configシートで設定を確認してください。');
  }
}

function showSlackConfig() {
  const webhookUrl = getConfig('SLACK_WEBHOOK_URL');
  const channel = getConfig('SLACK_CHANNEL');
  const username = getConfig('SLACK_USERNAME');
  
  const message = `現在のSlack設定：\n\nWebhook URL: ${webhookUrl || '未設定'}\nチャンネル: ${channel}\nユーザー名: ${username}\n\n設定を変更する場合は、Configシートを編集してください。`;
  
  SpreadsheetApp.getUi().alert(message);
}

// ========================================
// 文書同期UI機能
// ========================================

function syncSelectedDocument() {
  const sheet = SpreadsheetApp.getActiveSheet();
  if (sheet.getName() !== SHEET_NAMES.REGISTRY) {
    SpreadsheetApp.getUi().alert('DocRegistryシートで実行してください');
    return;
  }
  
  const activeRange = sheet.getActiveRange();
  const row = activeRange.getRow();
  
  if (row <= 1) {
    SpreadsheetApp.getUi().alert('データ行を選択してください');
    return;
  }
  
  const docKey = sheet.getRange(row, COLUMNS.DOC_KEY + 1).getValue();
  const title = sheet.getRange(row, COLUMNS.TITLE + 1).getValue();
  
  if (!docKey) {
    SpreadsheetApp.getUi().alert('DocKeyが見つかりません');
    return;
  }
  
  SpreadsheetApp.getUi().alert('同期を開始します...');
  
  const result = syncDocuments(docKey);
  
  if (result.success) {
    let message = `文書の同期が完了しました\n\n`;
    message += `保存ファイル数: ${result.savedFiles}\n`;
    
    if (result.updates.googleDocUrl) {
      message += `Google Docを作成/更新しました\n`;
    }
    if (result.updates.wordFileUrl) {
      message += `Wordファイルを保存しました\n`;
    }
    if (result.diffStatus?.hasDifference) {
      message += `\n⚠️ 警告: Google DocとWordファイルに差分があります`;
    }
    
    SpreadsheetApp.getUi().alert(message);
    
    // Slack通知
    if (getConfig('NOTIFY_ON_STATUS_CHANGE') === 'TRUE') {
      notifyDocumentSync(docKey, title, result);
    }
  } else {
    SpreadsheetApp.getUi().alert('同期に失敗しました: ' + result.error);
  }
}

function syncAllDocumentsUI() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    '全文書の同期',
    'すべての文書を同期します。時間がかかる場合があります。続行しますか？',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) {
    return;
  }
  
  ui.alert('同期を開始します...');
  
  const result = syncAllDocuments();
  
  if (result.success) {
    let message = `同期が完了しました\n\n`;
    message += `同期文書数: ${result.syncedCount}\n\n`;
    
    if (result.results.length > 0) {
      message += '詳細:\n';
      result.results.slice(0, 5).forEach(r => {
        message += `${r.docKey}: ${r.savedFiles}ファイル保存`;
        if (r.hasDifference) {
          message += ' (差分あり)';
        }
        message += '\n';
      });
      
      if (result.results.length > 5) {
        message += `...他${result.results.length - 5}件`;
      }
    }
    
    ui.alert(message);
  } else {
    ui.alert('同期に失敗しました');
  }
}

function checkSelectedDocumentDifference() {
  const sheet = SpreadsheetApp.getActiveSheet();
  if (sheet.getName() !== SHEET_NAMES.REGISTRY) {
    SpreadsheetApp.getUi().alert('DocRegistryシートで実行してください');
    return;
  }
  
  const activeRange = sheet.getActiveRange();
  const row = activeRange.getRow();
  
  if (row <= 1) {
    SpreadsheetApp.getUi().alert('データ行を選択してください');
    return;
  }
  
  const docKey = sheet.getRange(row, COLUMNS.DOC_KEY + 1).getValue();
  const title = sheet.getRange(row, COLUMNS.TITLE + 1).getValue();
  
  if (!docKey) {
    SpreadsheetApp.getUi().alert('DocKeyが見つかりません');
    return;
  }
  
  const diffStatus = checkDocumentDifference(docKey);
  
  if (!diffStatus) {
    SpreadsheetApp.getUi().alert('差分チェックできませんでした');
    return;
  }
  
  if (diffStatus.error) {
    SpreadsheetApp.getUi().alert('エラー: ' + diffStatus.error);
    return;
  }
  
  if (!diffStatus.hasDifference) {
    SpreadsheetApp.getUi().alert('Google DocとWordファイルは同期しています');
  } else {
    const googleModified = formatDateTime(new Date(diffStatus.googleDocModified));
    const wordModified = formatDateTime(new Date(diffStatus.wordFileModified));
    const timeDiffMinutes = Math.floor(diffStatus.timeDifference / 60000);
    
    let message = '⚠️ 文書に差分があります\n\n';
    message += `Google Doc更新: ${googleModified}\n`;
    message += `Word更新: ${wordModified}\n`;
    message += `差分: ${timeDiffMinutes}分\n\n`;
    message += '文書を確認して同期してください';
    
    SpreadsheetApp.getUi().alert(message);
    
    // Slack通知
    if (getConfig('NOTIFY_ON_STATUS_CHANGE') === 'TRUE') {
      notifyDocumentDifference(docKey, title, diffStatus);
    }
  }
}

function saveSelectedAttachments() {
  const sheet = SpreadsheetApp.getActiveSheet();
  if (sheet.getName() !== SHEET_NAMES.REGISTRY) {
    SpreadsheetApp.getUi().alert('DocRegistryシートで実行してください');
    return;
  }
  
  const activeRange = sheet.getActiveRange();
  const row = activeRange.getRow();
  
  if (row <= 1) {
    SpreadsheetApp.getUi().alert('データ行を選択してください');
    return;
  }
  
  const docKey = sheet.getRange(row, COLUMNS.DOC_KEY + 1).getValue();
  
  if (!docKey) {
    SpreadsheetApp.getUi().alert('DocKeyが見つかりません');
    return;
  }
  
  SpreadsheetApp.getUi().alert('添付ファイルの保存を開始します...');
  
  const savedFiles = saveAttachmentsToDrive(docKey);
  
  if (savedFiles.length > 0) {
    let message = `${savedFiles.length}個のファイルをDriveに保存しました\n\n`;
    savedFiles.forEach(file => {
      message += `• ${file.name}\n`;
    });
    
    SpreadsheetApp.getUi().alert(message);
  } else {
    SpreadsheetApp.getUi().alert('保存するファイルが見つかりませんでした');
  }
}

/**
 * 親フォルダを設定
 */
function setDriveFolder() {
  const ui = SpreadsheetApp.getUi();
  
  // 現在の設定を取得
  const currentFolderId = getConfig('PARENT_FOLDER_ID') || '';
  
  let currentInfo = '【親フォルダの設定】\n\n';
  currentInfo += 'すべてのプロジェクトフォルダが作成される親フォルダを設定します。\n\n';
  currentInfo += '現在の設定:\n';
  if (currentFolderId) {
    try {
      const folder = DriveApp.getFolderById(currentFolderId);
      currentInfo += `フォルダ名: ${folder.getName()}\n`;
      currentInfo += `フォルダID: ${currentFolderId}\n`;
    } catch (e) {
      currentInfo += `フォルダID: ${currentFolderId}（アクセス不可）\n`;
    }
  } else {
    currentInfo += 'マイドライブ直下（デフォルト）\n';
  }
  
  currentInfo += '\n新しい親フォルダのURLまたはIDを入力してください。\n';
  currentInfo += '空欄にするとマイドライブ直下に作成されます。\n\n';
  currentInfo += '例:\n';
  currentInfo += '• https://drive.google.com/drive/folders/xxxxx\n';
  currentInfo += '• xxxxx（フォルダIDのみ）';
  
  const response = ui.prompt(
    '親フォルダの設定',
    currentInfo,
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  
  const input = response.getResponseText().trim();
  
  if (!input) {
    // 空欄の場合はクリア
    setConfig('PARENT_FOLDER_ID', '');
    ui.alert('親フォルダの設定をクリアしました。マイドライブ直下に作成されます。');
    return;
  }
  
  // URLからフォルダIDを抽出
  let folderId = input;
  
  if (input.includes('drive.google.com')) {
    const match = input.match(/folders\/([a-zA-Z0-9-_]+)/);
    if (match) {
      folderId = match[1];
    } else {
      ui.alert('エラー: 無効なDrive URLです。');
      return;
    }
  }
  
  // フォルダの存在確認
  try {
    const folder = DriveApp.getFolderById(folderId);
    const folderName = folder.getName();
    
    // 設定を保存
    setConfig('PARENT_FOLDER_ID', folderId);
    
    ui.alert(
      '設定完了',
      `親フォルダを設定しました。\n\nフォルダ名: ${folderName}\nフォルダID: ${folderId}\n\n今後作成されるプロジェクトフォルダはこのフォルダ内に作成されます。`,
      ui.ButtonSet.OK
    );
  } catch (e) {
    ui.alert(
      'エラー',
      '指定されたフォルダにアクセスできません。\nフォルダIDまたはURLを確認してください。',
      ui.ButtonSet.OK
    );
  }
}

// ========================================
// UI機能
// ========================================

function showAddDocumentUI() {
  const ui = SpreadsheetApp.getUi();
  
  const docKeyResponse = ui.prompt('新規文書追加', 'DocKey（空欄で自動生成）:', ui.ButtonSet.OK_CANCEL);
  if (docKeyResponse.getSelectedButton() !== ui.Button.OK) return;
  
  const titleResponse = ui.prompt('新規文書追加', 'タイトル:', ui.ButtonSet.OK_CANCEL);
  if (titleResponse.getSelectedButton() !== ui.Button.OK) return;
  
  const dueDateResponse = ui.prompt('新規文書追加', '期限（YYYY-MM-DD形式、空欄可）:', ui.ButtonSet.OK_CANCEL);
  if (dueDateResponse.getSelectedButton() !== ui.Button.OK) return;
  
  const result = addDocument({
    docKey: docKeyResponse.getResponseText() || null,
    title: titleResponse.getResponseText(),
    dueDate: dueDateResponse.getResponseText() || null
  });
  
  if (result.success) {
    ui.alert(`文書を追加しました\nDocKey: ${result.docKey}\n行: ${result.row}`);
  }
}

function showSearchUI() {
  const ui = SpreadsheetApp.getUi();
  
  const searchResponse = ui.prompt('文書検索', 'キーワード（DocKeyまたはタイトル）:', ui.ButtonSet.OK_CANCEL);
  if (searchResponse.getSelectedButton() !== ui.Button.OK) return;
  
  const keyword = searchResponse.getResponseText();
  const results = searchDocuments({ docKey: keyword, title: keyword });
  
  if (results.length === 0) {
    ui.alert('該当する文書が見つかりませんでした');
    return;
  }
  
  let message = `${results.length}件の文書が見つかりました：\n\n`;
  results.forEach(doc => {
    message += `行${doc.row}: ${doc.docKey} - ${doc.title}\n`;
    message += `  ステージ: ${doc.stage}, ステータス: ${doc.projectStatus}\n\n`;
  });
  
  ui.alert(message);
}

function editSelectedRow() {
  const sheet = SpreadsheetApp.getActiveSheet();
  if (sheet.getName() !== SHEET_NAMES.REGISTRY) {
    SpreadsheetApp.getUi().alert('DocRegistryシートで実行してください');
    return;
  }
  
  const activeRange = sheet.getActiveRange();
  const row = activeRange.getRow();
  
  if (row <= 1) {
    SpreadsheetApp.getUi().alert('データ行を選択してください');
    return;
  }
  
  const docKey = sheet.getRange(row, COLUMNS.DOC_KEY + 1).getValue();
  const currentTitle = sheet.getRange(row, COLUMNS.TITLE + 1).getValue();
  const currentStage = sheet.getRange(row, COLUMNS.STAGE + 1).getValue();
  const currentStatus = sheet.getRange(row, COLUMNS.PROJECT_STATUS + 1).getValue();
  
  const ui = SpreadsheetApp.getUi();
  
  const titleResponse = ui.prompt('文書編集', `タイトル（現在: ${currentTitle}）:`, ui.ButtonSet.OK_CANCEL);
  if (titleResponse.getSelectedButton() !== ui.Button.OK) return;
  
  const stageResponse = ui.prompt('文書編集', `ステージ（${Object.values(STAGES).join(', ')}）\n現在: ${currentStage}:`, ui.ButtonSet.OK_CANCEL);
  if (stageResponse.getSelectedButton() !== ui.Button.OK) return;
  
  const statusResponse = ui.prompt('文書編集', `プロジェクトステータス（${Object.values(PROJECT_STATUS).join(', ')}）\n現在: ${currentStatus}:`, ui.ButtonSet.OK_CANCEL);
  if (statusResponse.getSelectedButton() !== ui.Button.OK) return;
  
  const updates = {};
  if (titleResponse.getResponseText()) updates.title = titleResponse.getResponseText();
  if (stageResponse.getResponseText()) updates.stage = stageResponse.getResponseText();
  if (statusResponse.getResponseText()) updates.projectStatus = statusResponse.getResponseText();
  
  const result = updateDocument(docKey, updates);
  
  if (result.success) {
    ui.alert('文書を更新しました');
  } else {
    ui.alert('更新に失敗しました: ' + result.error);
  }
}

function showOverdueReport() {
  const overdueDocuments = getOverdueDocuments();
  
  if (overdueDocuments.length === 0) {
    SpreadsheetApp.getUi().alert('期限切れの文書はありません');
    return;
  }
  
  let report = `期限切れ文書一覧（${overdueDocuments.length}件）\n\n`;
  overdueDocuments.forEach(doc => {
    report += `${doc.docKey}: ${doc.title}\n`;
    report += `  期限: ${doc.dueDate} (${doc.daysPastDue}日超過)\n`;
    report += `  ステータス: ${doc.projectStatus}\n\n`;
  });
  
  SpreadsheetApp.getUi().alert(report);
}

function showStatusReport() {
  const sheet = getOrCreateSheet(SHEET_NAMES.REGISTRY);
  const data = sheet.getDataRange().getValues();
  
  const stageCount = {};
  const statusCount = {};
  
  for (let i = 1; i < data.length; i++) {
    const stage = data[i][COLUMNS.STAGE];
    const status = data[i][COLUMNS.PROJECT_STATUS];
    
    stageCount[stage] = (stageCount[stage] || 0) + 1;
    statusCount[status] = (statusCount[status] || 0) + 1;
  }
  
  let report = 'ステータス別集計\n\n';
  report += '【文書ステージ】\n';
  Object.keys(stageCount).forEach(stage => {
    report += `  ${stage}: ${stageCount[stage]}件\n`;
  });
  
  report += '\n【プロジェクトステータス】\n';
  Object.keys(statusCount).forEach(status => {
    report += `  ${status}: ${statusCount[status]}件\n`;
  });
  
  SpreadsheetApp.getUi().alert(report);
}

// ========================================
// 自動化・トリガー
// ========================================

function setupTriggers() {
  removeTriggers();
  
  // 期限切れチェック
  const overdueHour = parseInt(getConfig('OVERDUE_CHECK_HOUR') || '9');
  ScriptApp.newTrigger('notifyOverdueDocuments')
    .timeBased()
    .everyDays(1)
    .atHour(overdueHour)
    .create();
  
  // 週次サマリー
  const summaryDay = getConfig('WEEKLY_SUMMARY_DAY') || 'MONDAY';
  const summaryHour = parseInt(getConfig('WEEKLY_SUMMARY_HOUR') || '10');
  const weekDay = ScriptApp.WeekDay[summaryDay];
  
  if (weekDay) {
    ScriptApp.newTrigger('sendWeeklySummary')
      .timeBased()
      .onWeekDay(weekDay)
      .atHour(summaryHour)
      .create();
  }
  
  // メール情報の自動更新
  if (getConfig('AUTO_EMAIL_UPDATE') === 'TRUE') {
    ScriptApp.newTrigger('scheduledEmailUpdate')
      .timeBased()
      .everyHours(1)
      .create();
  }
  
  SpreadsheetApp.getUi().alert('定期実行トリガーを設定しました');
}

function removeTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    ScriptApp.deleteTrigger(trigger);
  });
}

function scheduledEmailUpdate() {
  if (getConfig('AUTO_EMAIL_UPDATE') !== 'TRUE') return;
  
  const sheet = getOrCreateSheet(SHEET_NAMES.REGISTRY);
  const data = sheet.getDataRange().getValues();
  let updatedCount = 0;
  
  for (let i = 1; i < data.length; i++) {
    const docKey = data[i][COLUMNS.DOC_KEY];
    const title = data[i][COLUMNS.TITLE];
    
    if (docKey) {
      const emailInfo = getLatestEmailInfo(docKey, title);
      
      if (emailInfo) {
        const mailTree = buildMailTree(docKey, title);
        
        data[i][COLUMNS.LAST_SENT_BY] = emailInfo.lastSentBy;
        data[i][COLUMNS.LAST_MAIL_URL] = emailInfo.lastMailUrl;
        data[i][COLUMNS.MAIL_TREE] = mailTree;
        
        updatedCount++;
      }
    }
  }
  
  if (updatedCount > 0) {
    sheet.getDataRange().setValues(data);
  }
  
  console.log(`Updated email info for ${updatedCount} documents`);
}

// ========================================
// ヘルプとガイド
// ========================================

function showSystemOverview() {
  const overview = `
📄 文書管理システム v1.0

【システム概要】
このシステムは、GoogleスプレッドシートとGoogle Apps Script（GAS）を使用した
統合型文書管理システムです。

【主要機能】
━━━━━━━━━━━━━━━━━━━━━
🎯 文書管理
  • 文書の一元管理とバージョン管理
  • DocKeyによる一意の識別
  • 改版番号（Rev）による版管理
  • 4段階のステージ管理
  • プロジェクト全体のステータス管理

📧 Gmail連携
  • 文書に関連するメールの自動検索
  • 送受信履歴の自動記録
  • メールツリーの可視化
  • 最新送信者の自動判定

🔔 Slack通知
  • リアルタイム通知機能
  • 期限切れアラート
  • ステータス変更通知
  • 週次レポート自動送信

📊 レポート機能
  • 期限切れ文書の一覧表示
  • ステータス別の集計
  • プロジェクト進捗の可視化

⏰ 自動化
  • 定期的なメール情報更新
  • 期限チェックの自動実行
  • 定期レポートの自動送信
━━━━━━━━━━━━━━━━━━━━━

【データ構造】
• DocRegistry: メインの文書管理シート
• Config: システム設定シート

【特徴】
✅ コード不要で全設定が可能
✅ スプレッドシート上で完結
✅ Gmail・Slackとの連携
✅ 自動化による業務効率化
✅ 直感的なメニュー操作

バージョン: 1.0
作成日: 2025-09-13`;
  
  SpreadsheetApp.getUi().alert('システム概要', overview, SpreadsheetApp.getUi().ButtonSet.OK);
}

function showQuickGuide() {
  const guide = `
📚 クイックスタートガイド

【初回セットアップ】
━━━━━━━━━━━━━━━━━━━━━
1️⃣ 初期設定
   メニュー → 「🚀 初期設定」をクリック
   → 必要なシートが自動作成されます

2️⃣ Slack連携（任意）
   Configシートを開き、SLACK_WEBHOOK_URLを設定
   → メニューから「テスト通知」で動作確認

3️⃣ 自動化設定（任意）
   メニュー → 「⏰ 自動化」→「定期実行を設定」
   → 定期チェックが自動設定されます
━━━━━━━━━━━━━━━━━━━━━

【基本的な使い方】
━━━━━━━━━━━━━━━━━━━━━
📝 新規文書の追加
   1. メニュー → 「📋 文書操作」→「新規文書を追加」
   2. DocKey、タイトル、期限を入力
   3. 自動でDocRegistryシートに追加

🔍 文書の検索
   1. メニュー → 「📋 文書操作」→「文書を検索」
   2. キーワードを入力
   3. 該当する文書が表示

✏️ 文書の編集
   1. DocRegistryシートで編集したい行を選択
   2. メニュー → 「📋 文書操作」→「選択行を編集」
   3. 更新内容を入力

📧 メール情報の更新
   • 選択行のみ：該当行を選択してメニューから実行
   • 全文書：メニューから「全文書のメール情報を更新」
━━━━━━━━━━━━━━━━━━━━━

【ステージの流れ】
DRAFT（下書き）
  ↓
FOR-REVIEW（レビュー中）
  ↓
APPROVED（承認済み）
  ↓
ARCHIVED（アーカイブ）

【プロジェクトステータス】
• OPEN: 新規案件
• IN-PROGRESS: 進行中
• CLOSED: 完了
• DELAYED: 遅延

【💡 便利な使い方】
• 期限が近い文書は自動で赤色表示
• メール履歴はMailTree列で確認可能
• Slack通知で重要な変更を見逃さない
• 週次サマリーで全体状況を把握

詳細は「ℹ️ ヘルプ」をご覧ください`;
  
  SpreadsheetApp.getUi().alert('使い方ガイド', guide, SpreadsheetApp.getUi().ButtonSet.OK);
}

function checkFirstRun() {
  const properties = PropertiesService.getUserProperties();
  const hasRunBefore = properties.getProperty('HAS_RUN_BEFORE');
  
  if (!hasRunBefore) {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      '🎉 文書管理システムへようこそ！',
      '初めてご利用いただきありがとうございます。\n\n' +
      'このシステムは文書の管理、Gmail連携、Slack通知などの機能を提供します。\n\n' +
      '【次のステップ】\n' +
      '1. 「システム概要」で機能を確認\n' +
      '2. 「使い方ガイド」でクイックスタート\n' +
      '3. 「初期設定」でシステムをセットアップ\n\n' +
      '今すぐ初期設定を実行しますか？',
      ui.ButtonSet.YES_NO
    );
    
    if (response === ui.Button.YES) {
      initializeSystem();
    } else {
      ui.alert(
        '準備ができたら、メニューから「🚀 初期設定」を実行してください。\n\n' +
        'メニューの「📖 システム概要」と「📚 使い方ガイド」もご確認ください。'
      );
    }
    
    properties.setProperty('HAS_RUN_BEFORE', 'true');
  }
}

function showHelp() {
  const help = `
文書管理システム ヘルプ

【初期設定】
1. メニューから「初期設定」を実行
2. Configシートが作成されるので、必要に応じて設定を編集
3. Slack通知を使う場合は、SLACK_WEBHOOK_URLを設定

【基本機能】
• 文書の登録・更新・検索
• 改版管理（Rev）
• ステージ管理（下書き→レビュー→承認→アーカイブ）
• プロジェクトステータス管理
• 期限管理

【Gmail連携】
• 関連メールの自動検索
• 送受信履歴の記録
• メールツリーの構築

【Slack通知】
• 文書追加・更新通知
• 期限切れアラート
• 週次サマリー
• プロジェクト完了通知

【自動化】
• 定期的なメール情報更新
• 自動期限チェック
• 定期レポート送信

【DocKey命名規則例】
• IR: 投資家向け文書
• BOD: 取締役会関連
• FIN: 財務関連
• LEG: 法務関連
※自由に設定可能です

【設定項目】
Configシートで各種設定を変更できます：
• Slack通知設定
• 自動更新設定
• 通知タイミング設定

【トラブルシューティング】
• メニューが表示されない → ページを更新（F5）
• Slack通知が届かない → Webhook URLを確認
• Gmail連携が動作しない → 権限を確認

【サポート】
問題が発生した場合は、エラーメッセージと
実行ログを確認してください。`;
  
  SpreadsheetApp.getUi().alert(help);
}

// ========================================
// ユーティリティ関数
// ========================================

/**
 * ファイル名用の日時フォーマット（YYYYMMDD_HHMM）
 */
function formatFileDateTime(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  const hours = String(date.getHours()).padStart(2, '0');
  const minutes = String(date.getMinutes()).padStart(2, '0');
  return `${year}${month}${day}_${hours}${minutes}`;
}

/**
 * 現在のステージを取得
 */
function getCurrentStage(docKey) {
  const sheet = getOrCreateSheet(SHEET_NAMES.REGISTRY);
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][COLUMNS.DOC_KEY] === docKey) {
      return data[i][COLUMNS.STAGE] || STAGES.DRAFT;
    }
  }
  
  return STAGES.DRAFT;
}

/**
 * 現在のリビジョンを取得
 */
function getCurrentRevision(docKey) {
  const sheet = getOrCreateSheet(SHEET_NAMES.REGISTRY);
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][COLUMNS.DOC_KEY] === docKey) {
      return data[i][COLUMNS.REV] || 'r1';
    }
  }
  
  return 'r1';
}

/**
 * プロジェクトフォルダを取得
 */
function getProjectFolder(docKey) {
  const sheet = getOrCreateSheet(SHEET_NAMES.REGISTRY);
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][COLUMNS.DOC_KEY] === docKey) {
      const folderId = data[i][COLUMNS.FOLDER_ID];
      if (folderId) {
        try {
          return DriveApp.getFolderById(folderId);
        } catch (e) {
          console.error('フォルダIDが無効です:', e);
          return null;
        }
      }
      break;
    }
  }
  
  return null;
}

/**
 * プロジェクトのサブフォルダを取得
 */
function getProjectSubFolder(projectFolder, subFolderName) {
  if (!projectFolder) return null;
  
  const folders = projectFolder.getFoldersByName(subFolderName);
  if (folders.hasNext()) {
    return folders.next();
  } else {
    return projectFolder.createFolder(subFolderName);
  }
}

/**
 * DocKeyのフォルダ情報を更新
 */
function updateProjectFolder(docKey, folderId, folderUrl) {
  const sheet = getOrCreateSheet(SHEET_NAMES.REGISTRY);
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][COLUMNS.DOC_KEY] === docKey) {
      sheet.getRange(i + 1, COLUMNS.FOLDER_ID + 1).setValue(folderId);
      sheet.getRange(i + 1, COLUMNS.FOLDER_URL + 1).setValue(folderUrl);
      return true;
    }
  }
  
  return false;
}

function getOrCreateSheet(sheetName) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(sheetName);
  
  if (!sheet) {
    if (sheetName === SHEET_NAMES.REGISTRY) {
      initializeMainSheet();
      sheet = spreadsheet.getSheetByName(sheetName);
    } else if (sheetName === SHEET_NAMES.CONFIG) {
      initializeConfigSheet();
      sheet = spreadsheet.getSheetByName(sheetName);
    }
  }
  
  return sheet;
}

function generateDocKey() {
  const timestamp = new Date().getTime();
  const random = Math.floor(Math.random() * 1000);
  return `DOC${timestamp}${random}`;
}

function formatDate(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
}

function formatDateTime(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  const hours = String(date.getHours()).padStart(2, '0');
  const minutes = String(date.getMinutes()).padStart(2, '0');
  return `${year}-${month}-${day} ${hours}:${minutes}`;
}
