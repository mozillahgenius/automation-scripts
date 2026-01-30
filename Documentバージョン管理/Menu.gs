/**
 * メニューとユーザーインターフェース
 * スプレッドシートのカスタムメニューとダイアログ
 */

/**
 * スプレッドシート開いた時にメニューを作成
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  ui.createMenu('📄 文書管理')
    .addItem('📋 シートを初期化', 'initializeSheet')
    .addSeparator()
    .addSubMenu(ui.createMenu('➕ 文書操作')
      .addItem('新規文書を追加', 'showAddDocumentDialog')
      .addItem('文書を検索', 'showSearchDialog')
      .addItem('文書を更新', 'showUpdateDialog'))
    .addSeparator()
    .addSubMenu(ui.createMenu('📧 メール連携')
      .addItem('選択行のメール情報を更新', 'updateSelectedRowEmailInfo')
      .addItem('全文書のメール情報を更新', 'updateAllEmailInfo')
      .addItem('関連メールを検索', 'showEmailSearchDialog'))
    .addSeparator()
    .addSubMenu(ui.createMenu('🔔 Slack通知')
      .addItem('Slack設定', 'showSlackConfigDialog')
      .addItem('テスト通知を送信', 'testSlackNotification')
      .addItem('期限切れ通知を送信', 'notifyOverdueDocuments')
      .addItem('週次サマリーを送信', 'sendWeeklySummary'))
    .addSeparator()
    .addSubMenu(ui.createMenu('📊 レポート')
      .addItem('期限切れ文書一覧', 'showOverdueReport')
      .addItem('ステータス別集計', 'showStatusReport')
      .addItem('プロジェクト進捗', 'showProjectReport'))
    .addSeparator()
    .addSubMenu(ui.createMenu('⏰ 自動化設定')
      .addItem('定期実行トリガーを設定', 'setupTriggers')
      .addItem('トリガーを削除', 'removeTriggers'))
    .addSeparator()
    .addItem('ℹ️ ヘルプ', 'showHelp')
    .addToUi();
}

/**
 * 新規文書追加ダイアログを表示
 */
function showAddDocumentDialog() {
  const html = HtmlService.createHtmlOutputFromFile('AddDocumentDialog')
    .setWidth(500)
    .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, '新規文書を追加');
}

/**
 * 文書検索ダイアログを表示
 */
function showSearchDialog() {
  const html = HtmlService.createHtmlOutputFromFile('SearchDialog')
    .setWidth(600)
    .setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(html, '文書を検索');
}

/**
 * 文書更新ダイアログを表示
 */
function showUpdateDialog() {
  const html = HtmlService.createHtmlOutputFromFile('UpdateDialog')
    .setWidth(500)
    .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, '文書を更新');
}

/**
 * メール検索ダイアログを表示
 */
function showEmailSearchDialog() {
  const html = HtmlService.createHtmlOutputFromFile('EmailSearchDialog')
    .setWidth(600)
    .setHeight(500);
  SpreadsheetApp.getUi().showModalDialog(html, '関連メールを検索');
}

/**
 * Slack設定ダイアログを表示
 */
function showSlackConfigDialog() {
  const html = HtmlService.createHtmlOutputFromFile('SlackConfigDialog')
    .setWidth(500)
    .setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(html, 'Slack設定');
}

/**
 * 選択行のメール情報を更新
 */
function updateSelectedRowEmailInfo() {
  const sheet = SpreadsheetApp.getActiveSheet();
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

/**
 * 期限切れレポートを表示
 */
function showOverdueReport() {
  const overdueDocuments = getOverdueDocuments();
  
  if (overdueDocuments.length === 0) {
    SpreadsheetApp.getUi().alert('期限切れの文書はありません');
    return;
  }
  
  let report = '期限切れ文書一覧\n\n';
  overdueDocuments.forEach(doc => {
    report += `${doc.docKey}: ${doc.title}\n`;
    report += `  期限: ${doc.dueDate} (${doc.daysPastDue}日超過)\n`;
    report += `  ステータス: ${doc.projectStatus}\n\n`;
  });
  
  SpreadsheetApp.getUi().alert(report);
}

/**
 * ステータス別集計を表示
 */
function showStatusReport() {
  const sheet = getOrCreateSheet();
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

/**
 * プロジェクト進捗レポートを表示
 */
function showProjectReport() {
  const sheet = getOrCreateSheet();
  const data = sheet.getDataRange().getValues();
  
  const projects = {};
  
  for (let i = 1; i < data.length; i++) {
    const docKey = data[i][COLUMNS.DOC_KEY];
    const title = data[i][COLUMNS.TITLE];
    const stage = data[i][COLUMNS.STAGE];
    const status = data[i][COLUMNS.PROJECT_STATUS];
    const dueDate = data[i][COLUMNS.DUE_DATE];
    const lastSentBy = data[i][COLUMNS.LAST_SENT_BY];
    
    projects[docKey] = {
      title: title,
      stage: stage,
      status: status,
      dueDate: dueDate,
      lastSentBy: lastSentBy
    };
  }
  
  let report = 'プロジェクト進捗レポート\n\n';
  
  // 進行中のプロジェクト
  report += '【進行中】\n';
  Object.keys(projects).forEach(key => {
    if (projects[key].status === PROJECT_STATUS.IN_PROGRESS) {
      report += `${key}: ${projects[key].title}\n`;
      report += `  ステージ: ${projects[key].stage}\n`;
      report += `  期限: ${projects[key].dueDate || '未設定'}\n`;
      report += `  最終送信: ${projects[key].lastSentBy || '未送信'}\n\n`;
    }
  });
  
  // 遅延中のプロジェクト
  report += '【遅延中】\n';
  Object.keys(projects).forEach(key => {
    if (projects[key].status === PROJECT_STATUS.DELAYED) {
      report += `${key}: ${projects[key].title}\n`;
      report += `  ステージ: ${projects[key].stage}\n`;
      report += `  期限: ${projects[key].dueDate || '未設定'}\n\n`;
    }
  });
  
  SpreadsheetApp.getUi().alert(report);
}

/**
 * 定期実行トリガーを設定
 */
function setupTriggers() {
  // 既存のトリガーを削除
  removeTriggers();
  
  // 毎日朝9時に期限切れ通知
  ScriptApp.newTrigger('notifyOverdueDocuments')
    .timeBased()
    .everyDays(1)
    .atHour(9)
    .create();
  
  // 毎週月曜日朝10時に週次サマリー
  ScriptApp.newTrigger('sendWeeklySummary')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(10)
    .create();
  
  // 1時間ごとにメール情報を更新
  ScriptApp.newTrigger('scheduledEmailUpdate')
    .timeBased()
    .everyHours(1)
    .create();
  
  SpreadsheetApp.getUi().alert('定期実行トリガーを設定しました');
}

/**
 * トリガーを削除
 */
function removeTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    ScriptApp.deleteTrigger(trigger);
  });
}

/**
 * ヘルプを表示
 */
function showHelp() {
  const help = `
文書管理システム ヘルプ

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

【使い方】
1. まず「シートを初期化」を実行
2. Slack通知を使う場合は「Slack設定」でWebhook URLを設定
3. 「新規文書を追加」で文書を登録
4. 定期実行が必要な場合は「定期実行トリガーを設定」を実行

【DocKey命名規則】
• IR: 投資家向け文書
• BOD: 取締役会関連
• FIN: 財務関連
• LEG: 法務関連
• その他自由に設定可能

お問い合わせ: システム管理者まで`;
  
  SpreadsheetApp.getUi().alert(help);
}

/**
 * フォームから文書を追加（HTMLダイアログから呼び出し）
 */
function addDocumentFromForm(formData) {
  const result = addDocument(formData);
  
  if (result.success && formData.notifySlack) {
    notifyNewDocument(formData);
  }
  
  return result;
}

/**
 * フォームから文書を更新（HTMLダイアログから呼び出し）
 */
function updateDocumentFromForm(docKey, formData) {
  const oldDoc = searchDocuments({ docKey: docKey })[0];
  const result = updateDocument(docKey, formData);
  
  if (result.success && formData.notifySlack) {
    if (oldDoc && oldDoc.stage !== formData.stage) {
      notifyStatusChange(docKey, oldDoc.stage, formData.stage, formData.title);
    }
    if (formData.projectStatus === PROJECT_STATUS.CLOSED) {
      notifyProjectCompletion(docKey, formData.title);
    }
  }
  
  return result;
}

/**
 * 現在のSlack設定を取得（HTMLダイアログから呼び出し）
 */
function getCurrentSlackConfig() {
  return getSlackConfig();
}

/**
 * Slack設定を保存（HTMLダイアログから呼び出し）
 */
function saveSlackConfig(config) {
  setSlackConfig(config.webhookUrl, config.channel, config.username, config.iconEmoji);
  return { success: true };
}