/**
 * スプレッドシートを開いた時にメニューを作成
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('訴訟管理')
    .addItem('🔧 初期セットアップ', 'setupExistingSheet')
    .addSeparator()
    .addSubMenu(ui.createMenu('📝 案件管理')
      .addItem('新規案件を登録', 'showAddCaseDialog')
      .addItem('案件を更新', 'showUpdateCaseDialog')
      .addItem('進行状況を追加', 'showAddTimelineDialog'))
    .addSeparator()
    .addSubMenu(ui.createMenu('📊 レポート')
      .addItem('週次レポートを生成', 'generateAndShowWeeklyReport')
      .addItem('期日チェック', 'checkAndShowUpcomingDeadlines')
      .addItem('未更新案件チェック', 'checkAndShowStaleCases'))
    .addSeparator()
    .addSubMenu(ui.createMenu('🔔 リマインド')
      .addItem('手動リマインド送信', 'manualReminderTest')
      .addItem('リマインドトリガー設定', 'setupReminderTriggers')
      .addItem('Slack通知テスト', 'testSlackNotification'))
    .addSeparator()
    .addSubMenu(ui.createMenu('ℹ️ ヘルプ')
      .addItem('セットアップガイド', 'showSetupGuide')
      .addItem('Slack連携ガイド', 'showSlackGuide'))
    .addToUi();
}

/**
 * セットアップガイド（実行してください）
 */
function setupGuide() {
  const guide = `
=== 訴訟管理システム セットアップガイド ===

1. CONFIG設定を更新してください：
   - SPREADSHEET_ID: スプレッドシートのIDを入力
   - SHEET_NAME: メインシート名を入力
   - REMINDER_EMAIL: リマインド送信先メールアドレスを入力
   
2. オプション設定：
   - DOCUMENT_SHEET_NAME: 関連書類管理用シート名
   - TIMELINE_SHEET_NAME: 進行状況管理用シート名
   - SLACK_WEBHOOK_URL: Slack通知用WebhookURL

3. セットアップを実行：
   setupExistingSheet() を実行

4. リマインド機能：
   - 毎日午前9時: 期日チェック
   - 毎週月曜日午前10時: 週次レポート
   - 30日未更新案件の通知

5. 手動テスト：
   manualReminderTest() でリマインドをテスト実行

6. Slackテスト関数：
   - testSlackReminderOnly() : 期日リマインドのみ
   - testSlackStaleOnly() : 未更新案件リマインドのみ  
   - testSlackWeeklyOnly() : 週次レポートのみ
   - testSlackNotification() : 全ての通知をテスト
  `;
  
  Logger.log(guide);
  return guide;
}

// 訴訟管理システム - Google Apps Script (既存シート使用版)
// メインスクリプト (Code.gs)

// 設定 - 既存のスプレッドシートとシートを指定
const CONFIG = {
  SPREADSHEET_ID: 'SPREADSHEET_ID_PLACEHOLDER', // ここにスプレッドシートIDを入力
  SHEET_NAME: '訴訟管理',         // ここにシート名を入力
  DOCUMENT_SHEET_NAME: '関連書類', // 関連書類シート名
  TIMELINE_SHEET_NAME: '進行状況', // 進行状況シート名
  REMINDER_EMAIL: 'admin@example.com',    // リマインド送信先メール
  SLACK_WEBHOOK_URL: 'https://hooks.slack.com/services/YOUR/WEBHOOK/URL',                       // Slack通知用（オプション）
};

// 列のインデックス定義
const COLUMNS = {
  ID: 0,
  CASE_NUMBER: 1,
  CASE_NAME: 2,
  CASE_TYPE: 3,
  PLAINTIFF: 4,
  DEFENDANT: 5,
  COURT: 6,
  LAWYER: 7,
  STATUS: 8,
  FILING_DATE: 9,
  NEXT_HEARING: 10,
  AMOUNT: 11,
  DESCRIPTION: 12,
  RESPONSIBLE_PERSON: 13,
  CREATED_DATE: 14,
  UPDATED_DATE: 15
};

/**
 * 既存シートのセットアップ
 */
function setupExistingSheet() {
  try {
    // 設定チェック
    if (CONFIG.SPREADSHEET_ID === 'YOUR_SPREADSHEET_ID_HERE') {
      throw new Error('CONFIG.SPREADSHEET_IDを実際のスプレッドシートIDに変更してください');
    }
    
    const spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    
    // メインシートのセットアップ
    setupMainSheet(spreadsheet);
    
    // 関連シートのセットアップ（存在する場合のみ）
    if (CONFIG.DOCUMENT_SHEET_NAME && CONFIG.DOCUMENT_SHEET_NAME !== 'YOUR_DOCUMENT_SHEET_NAME_HERE') {
      setupDocumentSheet(spreadsheet);
    }
    
    if (CONFIG.TIMELINE_SHEET_NAME && CONFIG.TIMELINE_SHEET_NAME !== 'YOUR_TIMELINE_SHEET_NAME_HERE') {
      setupTimelineSheet(spreadsheet);
    }
    
    // リマインド用トリガーの設定
    setupReminderTriggers();
    
    Logger.log('既存シートのセットアップが完了しました');
    return '既存シートのセットアップが完了しました';
  } catch (error) {
    Logger.log('セットアップエラー: ' + error.toString());
    throw error;
  }
}

/**
 * メインシートのセットアップ（既存シートにヘッダー追加）
 */
function setupMainSheet(spreadsheet) {
  const sheet = spreadsheet.getSheetByName(CONFIG.SHEET_NAME);
  
  if (!sheet) {
    throw new Error(`シート "${CONFIG.SHEET_NAME}" が見つかりません`);
  }
  
  // ヘッダーの確認と設定
  const headers = [
    'ID', '事件番号', '事件名', '事件種別', '原告', '被告', 
    '裁判所', '担当弁護士', 'ステータス', '提訴日', 
    '次回期日', '訴額', '概要', '担当者', '作成日', '更新日'
  ];
  
  // 既存のヘッダーをチェック
  const existingHeaders = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  const needsHeader = existingHeaders.every(cell => cell === '');
  
  if (needsHeader) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    // ヘッダーの書式設定
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#4285f4');
    headerRange.setFontColor('white');
    headerRange.setFontWeight('bold');
    
    // 列幅の調整
    sheet.setColumnWidth(1, 60);   // ID
    sheet.setColumnWidth(2, 150);  // 事件番号
    sheet.setColumnWidth(3, 200);  // 事件名
    sheet.setColumnWidth(4, 100);  // 事件種別
    sheet.setColumnWidth(12, 300); // 概要
    
    Logger.log('ヘッダーを追加しました');
  } else {
    Logger.log('既存のヘッダーを使用します');
  }
  
  return sheet;
}

/**
 * 関連書類シートのセットアップ
 */
function setupDocumentSheet(spreadsheet) {
  let sheet = spreadsheet.getSheetByName(CONFIG.DOCUMENT_SHEET_NAME);
  
  if (!sheet) {
    Logger.log(`関連書類シート "${CONFIG.DOCUMENT_SHEET_NAME}" が見つかりません`);
    return null;
  }
  
  const headers = ['文書ID', '事件ID', '文書名', '文書種別', 'ファイルURL', '作成日', '備考'];
  
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#34a853');
    headerRange.setFontColor('white');
    headerRange.setFontWeight('bold');
  }
  
  return sheet;
}

/**
 * 進行状況シートのセットアップ
 */
function setupTimelineSheet(spreadsheet) {
  let sheet = spreadsheet.getSheetByName(CONFIG.TIMELINE_SHEET_NAME);
  
  if (!sheet) {
    Logger.log(`進行状況シート "${CONFIG.TIMELINE_SHEET_NAME}" が見つかりません`);
    return null;
  }
  
  const headers = ['記録ID', '事件ID', '日付', '内容', '担当者', '次回アクション', '期限'];
  
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#ff9900');
    headerRange.setFontColor('white');
    headerRange.setFontWeight('bold');
  }
  
  return sheet;
}

/**
 * リマインド用トリガーの設定
 */
function setupReminderTriggers() {
  // 既存のトリガーを削除
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'dailyReminderCheck' || 
        trigger.getHandlerFunction() === 'weeklyReminderCheck') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  // 毎日のリマインドチェック（午前9時）
  ScriptApp.newTrigger('dailyReminderCheck')
    .timeBased()
    .everyDays(1)
    .atHour(9)
    .create();
  
  // 週次リマインド（月曜日午前10時）
  ScriptApp.newTrigger('weeklyReminderCheck')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(10)
    .create();
  
  Logger.log('リマインドトリガーを設定しました');
}

/**
 * 新しい訴訟案件を追加
 */
function addLitigationCase(caseData) {
  try {
    const spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = spreadsheet.getSheetByName(CONFIG.SHEET_NAME);
    
    // 新しいIDを生成
    const newId = generateNewId(sheet);
    const currentDate = new Date();
    
    // データの準備
    const rowData = [
      newId,
      caseData.caseNumber || '',
      caseData.caseName || '',
      caseData.caseType || '',
      caseData.plaintiff || '',
      caseData.defendant || '',
      caseData.court || '',
      caseData.lawyer || '',
      caseData.status || '係属中',
      caseData.filingDate || '',
      caseData.nextHearing || '',
      caseData.amount || '',
      caseData.description || '',
      caseData.responsiblePerson || '',
      currentDate,
      currentDate
    ];
    
    // データの追加
    sheet.appendRow(rowData);
    
    // 進行状況に初期記録を追加
    addTimelineEntry(newId, '事件登録', '新規事件として登録されました', caseData.responsiblePerson || '');
    
    Logger.log('新しい訴訟案件を追加しました: ID=' + newId);
    return newId;
  } catch (error) {
    Logger.log('案件追加エラー: ' + error.toString());
    throw error;
  }
}

/**
 * 毎日のリマインドチェック
 */
function dailyReminderCheck() {
  try {
    const urgentCases = checkUpcomingDeadlines(3); // 3日以内
    const soonCases = checkUpcomingDeadlines(7);   // 7日以内
    
    if (urgentCases.length > 0 || soonCases.length > 0) {
      sendEmailReminder(urgentCases, soonCases);
      
      if (CONFIG.SLACK_WEBHOOK_URL) {
        sendSlackReminder(urgentCases, soonCases);
      }
    }
    
    // 長期未更新案件のチェック
    const staleCases = checkStaleCases(30); // 30日間未更新
    if (staleCases.length > 0) {
      sendStaleReminder(staleCases);
      
      if (CONFIG.SLACK_WEBHOOK_URL) {
        sendSlackStaleReminder(staleCases);
      }
    }
    
    Logger.log('毎日のリマインドチェック完了');
  } catch (error) {
    Logger.log('リマインドチェックエラー: ' + error.toString());
  }
}

/**
 * 週次リマインドチェック
 */
function weeklyReminderCheck() {
  try {
    const weeklyReport = generateWeeklyReport();
    sendWeeklyReport(weeklyReport);
    
    if (CONFIG.SLACK_WEBHOOK_URL) {
      sendSlackWeeklyReport(weeklyReport);
    }
    
    Logger.log('週次リマインド送信完了');
  } catch (error) {
    Logger.log('週次リマインドエラー: ' + error.toString());
  }
}

/**
 * メールリマインドの送信
 */
function sendEmailReminder(urgentCases, soonCases) {
  let subject = '【訴訟管理】期日リマインド';
  let body = '訴訟管理システムからのリマインドです。\n\n';
  
  if (urgentCases.length > 0) {
    subject = '【緊急】' + subject;
    body += '🚨 緊急：3日以内に期日がある案件 🚨\n';
    body += '='.repeat(40) + '\n';
    
    urgentCases.forEach(case_ => {
      const daysLeft = Math.ceil((case_.nextHearing - new Date()) / (1000 * 60 * 60 * 24));
      body += `📋 ${case_.caseName}\n`;
      body += `   事件番号: ${case_.caseNumber || 'なし'}\n`;
      body += `   次回期日: ${case_.nextHearing.toLocaleDateString('ja-JP')} (${daysLeft}日後)\n`;
      body += `   担当者: ${case_.responsiblePerson || 'なし'}\n`;
      body += `   裁判所: ${case_.court || 'なし'}\n\n`;
    });
  }
  
  if (soonCases.length > 0) {
    body += '⚠️ 7日以内に期日がある案件 ⚠️\n';
    body += '=' .repeat(40) + '\n';
    
    soonCases.forEach(case_ => {
      if (!urgentCases.some(urgent => urgent.id === case_.id)) {
        const daysLeft = Math.ceil((case_.nextHearing - new Date()) / (1000 * 60 * 60 * 24));
        body += `📋 ${case_.caseName}\n`;
        body += `   次回期日: ${case_.nextHearing.toLocaleDateString('ja-JP')} (${daysLeft}日後)\n`;
        body += `   担当者: ${case_.responsiblePerson || 'なし'}\n\n`;
      }
    });
  }
  
  body += '\n詳細は訴訟管理シートをご確認ください。\n';
  body += `シートURL: https://docs.google.com/spreadsheets/d/${CONFIG.SPREADSHEET_ID}\n`;
  
  try {
    MailApp.sendEmail({
      to: CONFIG.REMINDER_EMAIL,
      subject: subject,
      body: body
    });
    
    Logger.log('リマインドメールを送信しました');
  } catch (error) {
    Logger.log('メール送信エラー: ' + error.toString());
  }
}

/**
 * Slackリマインドの送信（リッチフォーマット版）
 */
function sendSlackReminder(urgentCases, soonCases) {
  if (!CONFIG.SLACK_WEBHOOK_URL) return;
  
  const attachments = [];
  
  // 緊急案件のアタッチメント
  if (urgentCases.length > 0) {
    let urgentText = '';
    urgentCases.forEach(case_ => {
      const daysLeft = Math.ceil((case_.nextHearing - new Date()) / (1000 * 60 * 60 * 24));
      urgentText += `📋 *${case_.caseName}*\n`;
      urgentText += `　🏛️ ${case_.court || '裁判所未設定'}\n`;
      urgentText += `　📅 ${case_.nextHearing.toLocaleDateString('ja-JP')} (*${daysLeft}日後*)\n`;
      urgentText += `　👤 ${case_.responsiblePerson || '担当者未設定'}\n\n`;
    });
    
    attachments.push({
      "color": "#ff0000",
      "title": "🚨 緊急：3日以内の期日",
      "text": urgentText,
      "footer": "訴訟管理システム",
      "ts": Math.floor(Date.now() / 1000)
    });
  }
  
  // 注意案件のアタッチメント
  if (soonCases.length > 0) {
    let soonText = '';
    soonCases.forEach(case_ => {
      if (!urgentCases.some(urgent => urgent.id === case_.id)) {
        const daysLeft = Math.ceil((case_.nextHearing - new Date()) / (1000 * 60 * 60 * 24));
        soonText += `📋 *${case_.caseName}*\n`;
        soonText += `　📅 ${case_.nextHearing.toLocaleDateString('ja-JP')} (${daysLeft}日後)\n`;
        soonText += `　👤 ${case_.responsiblePerson || '担当者未設定'}\n\n`;
      }
    });
    
    if (soonText) {
      attachments.push({
        "color": "#ff9900",
        "title": "⚠️ 7日以内の期日",
        "text": soonText,
        "footer": "訴訟管理システム"
      });
    }
  }
  
  // スプレッドシートへのリンクボタン
  attachments.push({
    "color": "#36a64f",
    "title": "📊 詳細情報",
    "text": "スプレッドシートで詳細を確認",
    "actions": [
      {
        "type": "button",
        "text": "スプレッドシートを開く",
        "url": `https://docs.google.com/spreadsheets/d/${CONFIG.SPREADSHEET_ID}`
      }
    ]
  });
  
  const payload = {
    'text': '🔔 *訴訟管理リマインド*',
    'username': '訴訟管理システム',
    'icon_emoji': ':scales:',
    'attachments': attachments
  };
  
  try {
    UrlFetchApp.fetch(CONFIG.SLACK_WEBHOOK_URL, {
      'method': 'POST',
      'contentType': 'application/json',
      'payload': JSON.stringify(payload)
    });
    
    Logger.log('Slackリマインドを送信しました');
  } catch (error) {
    Logger.log('Slack送信エラー: ' + error.toString());
  }
}

/**
 * Slack週次レポートの送信
 */
function sendSlackWeeklyReport(report) {
  if (!CONFIG.SLACK_WEBHOOK_URL) return;
  
  const lines = report.split('\n');
  let formattedReport = '';
  
  lines.forEach(line => {
    if (line.includes('週次レポート')) {
      formattedReport += `📊 *${line.replace('===', '').trim()}*\n`;
    } else if (line.includes('全体状況')) {
      formattedReport += `\n📈 *${line}*\n`;
    } else if (line.includes('ステータス別内訳')) {
      formattedReport += `\n📋 *${line}*\n`;
    } else if (line.includes('今後2週間の期日')) {
      formattedReport += `\n⏰ *${line}*\n`;
    } else if (line.trim() && !line.includes('期間:') && !line.includes('シートURL:')) {
      formattedReport += `${line}\n`;
    }
  });
  
  const attachments = [
    {
      "color": "#36a64f",
      "text": formattedReport,
      "footer": "訴訟管理システム 週次レポート",
      "ts": Math.floor(Date.now() / 1000),
      "actions": [
        {
          "type": "button",
          "text": "スプレッドシートを開く",
          "url": `https://docs.google.com/spreadsheets/d/${CONFIG.SPREADSHEET_ID}`
        }
      ]
    }
  ];
  
  const payload = {
    'text': '📊 *訴訟管理 週次レポート*',
    'username': '訴訟管理システム',
    'icon_emoji': ':chart_with_upwards_trend:',
    'attachments': attachments
  };
  
  try {
    UrlFetchApp.fetch(CONFIG.SLACK_WEBHOOK_URL, {
      'method': 'POST',
      'contentType': 'application/json',
      'payload': JSON.stringify(payload)
    });
    
    Logger.log('Slack週次レポートを送信しました');
  } catch (error) {
    Logger.log('Slack週次レポート送信エラー: ' + error.toString());
  }
}

/**
 * Slack未更新案件リマインドの送信
 */
function sendSlackStaleReminder(staleCases) {
  if (!CONFIG.SLACK_WEBHOOK_URL || staleCases.length === 0) return;
  
  let staleText = '';
  staleCases.forEach(case_ => {
    const daysSinceUpdate = Math.ceil((new Date() - case_.updatedDate) / (1000 * 60 * 60 * 24));
    staleText += `📋 *${case_.caseName}*\n`;
    staleText += `　📅 最終更新: ${case_.updatedDate.toLocaleDateString('ja-JP')} (*${daysSinceUpdate}日前*)\n`;
    staleText += `　👤 ${case_.responsiblePerson || '担当者未設定'}\n\n`;
  });
  
  const attachments = [
    {
      "color": "#ffcc00",
      "title": "📝 長期未更新案件",
      "text": staleText,
      "footer": "進行状況の更新をお願いします",
      "actions": [
        {
          "type": "button",
          "text": "スプレッドシートを開く",
          "url": `https://docs.google.com/spreadsheets/d/${CONFIG.SPREADSHEET_ID}`
        }
      ]
    }
  ];
  
  const payload = {
    'text': '⏰ *長期未更新案件のお知らせ*',
    'username': '訴訟管理システム',
    'icon_emoji': ':hourglass_flowing_sand:',
    'attachments': attachments
  };
  
  try {
    UrlFetchApp.fetch(CONFIG.SLACK_WEBHOOK_URL, {
      'method': 'POST',
      'contentType': 'application/json',
      'payload': JSON.stringify(payload)
    });
    
    Logger.log('Slack未更新リマインドを送信しました');
  } catch (error) {
    Logger.log('Slack未更新リマインド送信エラー: ' + error.toString());
  }
}

/**
 * 長期未更新案件のチェック
 */
function checkStaleCases(days = 30) {
  try {
    const spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = spreadsheet.getSheetByName(CONFIG.SHEET_NAME);
    
    const data = sheet.getDataRange().getValues();
    const staleCases = [];
    const checkDate = new Date();
    checkDate.setDate(checkDate.getDate() - days);
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const updatedDate = row[COLUMNS.UPDATED_DATE];
      const status = row[COLUMNS.STATUS];
      
      if (updatedDate && updatedDate instanceof Date && 
          updatedDate < checkDate && 
          status === '係属中') {
        staleCases.push({
          id: row[COLUMNS.ID],
          caseName: row[COLUMNS.CASE_NAME],
          caseNumber: row[COLUMNS.CASE_NUMBER],
          updatedDate: updatedDate,
          responsiblePerson: row[COLUMNS.RESPONSIBLE_PERSON]
        });
      }
    }
    
    return staleCases;
  } catch (error) {
    Logger.log('未更新案件チェックエラー: ' + error.toString());
    return [];
  }
}

/**
 * 長期未更新リマインドの送信
 */
function sendStaleReminder(staleCases) {
  const subject = '【訴訟管理】長期未更新案件の確認';
  let body = '以下の案件が30日以上更新されていません。\n';
  body += '進行状況の確認をお願いします。\n\n';
  
  staleCases.forEach(case_ => {
    const daysSinceUpdate = Math.ceil((new Date() - case_.updatedDate) / (1000 * 60 * 60 * 24));
    body += `📋 ${case_.caseName}\n`;
    body += `   事件番号: ${case_.caseNumber || 'なし'}\n`;
    body += `   最終更新: ${case_.updatedDate.toLocaleDateString('ja-JP')} (${daysSinceUpdate}日前)\n`;
    body += `   担当者: ${case_.responsiblePerson || 'なし'}\n\n`;
  });
  
  body += `\nシートURL: https://docs.google.com/spreadsheets/d/${CONFIG.SPREADSHEET_ID}\n`;
  
  try {
    MailApp.sendEmail({
      to: CONFIG.REMINDER_EMAIL,
      subject: subject,
      body: body
    });
    
    Logger.log('未更新リマインドメールを送信しました');
  } catch (error) {
    Logger.log('未更新リマインドメール送信エラー: ' + error.toString());
  }
}

/**
 * 週次レポートの生成
 */
function generateWeeklyReport() {
  const spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = spreadsheet.getSheetByName(CONFIG.SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  
  const totalCases = data.length - 1;
  const statusCounts = {};
  const upcomingCases = checkUpcomingDeadlines(14); // 2週間以内
  
  for (let i = 1; i < data.length; i++) {
    const status = data[i][COLUMNS.STATUS] || '不明';
    statusCounts[status] = (statusCounts[status] || 0) + 1;
  }
  
  let report = `=== 訴訟管理 週次レポート ===\n`;
  report += `期間: ${new Date().toLocaleDateString('ja-JP')}\n\n`;
  report += `📊 全体状況\n`;
  report += `総案件数: ${totalCases}件\n\n`;
  
  report += `📈 ステータス別内訳\n`;
  Object.keys(statusCounts).forEach(status => {
    report += `  ${status}: ${statusCounts[status]}件\n`;
  });
  
  if (upcomingCases.length > 0) {
    report += `\n⏰ 今後2週間の期日 (${upcomingCases.length}件)\n`;
    upcomingCases.forEach(case_ => {
      const daysLeft = Math.ceil((case_.nextHearing - new Date()) / (1000 * 60 * 60 * 24));
      report += `  ${case_.nextHearing.toLocaleDateString('ja-JP')} (${daysLeft}日後): ${case_.caseName}\n`;
    });
  }
  
  return report;
}

/**
 * 週次レポートの送信
 */
function sendWeeklyReport(report) {
  const subject = '【訴訟管理】週次レポート';
  const body = report + `\n\nシートURL: https://docs.google.com/spreadsheets/d/${CONFIG.SPREADSHEET_ID}\n`;
  
  try {
    MailApp.sendEmail({
      to: CONFIG.REMINDER_EMAIL,
      subject: subject,
      body: body
    });
    
    Logger.log('週次レポートを送信しました');
  } catch (error) {
    Logger.log('週次レポート送信エラー: ' + error.toString());
  }
}

/**
 * 期限が近い案件をチェック
 */
function checkUpcomingDeadlines(daysAhead = 7) {
  try {
    const spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = spreadsheet.getSheetByName(CONFIG.SHEET_NAME);
    
    const data = sheet.getDataRange().getValues();
    const upcomingCases = [];
    const checkDate = new Date();
    checkDate.setDate(checkDate.getDate() + daysAhead);
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const nextHearing = row[COLUMNS.NEXT_HEARING];
      
      if (nextHearing && nextHearing instanceof Date) {
        if (nextHearing <= checkDate && nextHearing >= new Date()) {
          upcomingCases.push({
            id: row[COLUMNS.ID],
            caseName: row[COLUMNS.CASE_NAME],
            caseNumber: row[COLUMNS.CASE_NUMBER],
            nextHearing: nextHearing,
            responsiblePerson: row[COLUMNS.RESPONSIBLE_PERSON],
            court: row[COLUMNS.COURT]
          });
        }
      }
    }
    
    return upcomingCases;
  } catch (error) {
    Logger.log('期限チェックエラー: ' + error.toString());
    return [];
  }
}

/**
 * 進行状況を追加
 */
function addTimelineEntry(caseId, title, content, responsiblePerson, nextAction = '', deadline = '') {
  try {
    if (!CONFIG.TIMELINE_SHEET_NAME || CONFIG.TIMELINE_SHEET_NAME === 'YOUR_TIMELINE_SHEET_NAME_HERE') {
      Logger.log('進行状況シートが設定されていません');
      return;
    }
    
    const spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = spreadsheet.getSheetByName(CONFIG.TIMELINE_SHEET_NAME);
    
    if (!sheet) {
      Logger.log('進行状況シートが見つかりません');
      return;
    }
    
    const newId = generateNewId(sheet);
    const currentDate = new Date();
    
    const rowData = [
      newId,
      caseId,
      currentDate,
      title + ': ' + content,
      responsiblePerson,
      nextAction,
      deadline
    ];
    
    sheet.appendRow(rowData);
    Logger.log('進行状況を追加しました: 事件ID=' + caseId);
  } catch (error) {
    Logger.log('進行状況追加エラー: ' + error.toString());
  }
}

/**
 * 手動リマインド実行（テスト用）
 */
function manualReminderTest() {
  Logger.log('手動リマインドテストを実行します');
  dailyReminderCheck();
}

// ヘルパー関数

function generateNewId(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return 1;
  
  const lastId = sheet.getRange(lastRow, 1).getValue();
  return (typeof lastId === 'number') ? lastId + 1 : lastRow;
}

/**
 * Slack設定ガイド
 */
function slackSetupGuide() {
  const guide = `
=== Slack連携設定ガイド ===

1. Slack Webhook URLの取得方法：
   ① Slackワークスペースで "Apps" → "Incoming Webhooks" を検索
   ② "Add to Slack" をクリック
   ③ 通知を送信したいチャンネルを選択
   ④ 生成されたWebhook URLをコピー

2. GASでの設定：
   CONFIG.SLACK_WEBHOOK_URL に取得したURLを設定
   例: 'https://hooks.slack.com/services/YOUR/WEBHOOK/URL'

3. Slack通知の特徴：
   🚨 緊急リマインド（3日以内）- 赤色
   ⚠️  注意リマインド（7日以内）- オレンジ色  
   📊 週次レポート - 緑色
   ⏰ 未更新案件 - 黄色

4. 含まれる情報：
   - 案件名、裁判所、期日、担当者
   - スプレッドシートへの直リンクボタン
   - リッチフォーマットで見やすい表示

5. テスト送信：
   testSlackNotification() でテスト通知を送信可能

=== カスタマイズオプション ===

- チャンネル別通知（複数Webhook設定）
- @here や @channel でのメンション
- 案件の重要度による色分け
- カスタム絵文字の使用
  `;
  
  Logger.log(guide);
  return guide;
}

/**
 * Slack通知のテスト送信
 */
function testSlackNotification() {
  if (!CONFIG.SLACK_WEBHOOK_URL) {
    Logger.log('SLACK_WEBHOOK_URLが設定されていません');
    return;
  }
  
  // テスト用の緊急案件
  const testUrgentCase = [{
    id: 999,
    caseName: 'テスト緊急案件',
    caseNumber: 'テスト番号001',
    nextHearing: new Date(Date.now() + 2 * 24 * 60 * 60 * 1000), // 2日後
    responsiblePerson: 'テスト担当者',
    court: 'テスト裁判所'
  }];
  
  // テスト用の注意案件
  const testSoonCase = [{
    id: 998,
    caseName: 'テスト注意案件',
    caseNumber: 'テスト番号002',
    nextHearing: new Date(Date.now() + 5 * 24 * 60 * 60 * 1000), // 5日後
    responsiblePerson: 'テスト担当者2',
    court: 'テスト地方裁判所'
  }];
  
  Logger.log('Slackテスト通知を送信します...');
  sendSlackReminder(testUrgentCase, testSoonCase);
  
  // 少し待ってから未更新案件のテストを送信
  Utilities.sleep(2000); // 2秒待機（GAS用のsleep関数）
  
  // 未更新案件のテスト
  const testStaleCase = [{
    id: 997,
    caseName: 'テスト未更新案件',
    caseNumber: 'テスト番号003',
    updatedDate: new Date(Date.now() - 35 * 24 * 60 * 60 * 1000), // 35日前
    responsiblePerson: 'テスト担当者3'
  }];
  
  sendSlackStaleReminder(testStaleCase);
  
  Logger.log('Slackテスト通知送信完了');
}

/**
 * 個別のSlackテスト関数
 */
function testSlackReminderOnly() {
  if (!CONFIG.SLACK_WEBHOOK_URL) {
    Logger.log('SLACK_WEBHOOK_URLが設定されていません');
    return;
  }
  
  const testUrgentCase = [{
    id: 999,
    caseName: 'テスト緊急案件',
    caseNumber: 'テスト番号001',
    nextHearing: new Date(Date.now() + 2 * 24 * 60 * 60 * 1000),
    responsiblePerson: 'テスト担当者',
    court: 'テスト裁判所'
  }];
  
  const testSoonCase = [{
    id: 998,
    caseName: 'テスト注意案件',
    caseNumber: 'テスト番号002',
    nextHearing: new Date(Date.now() + 5 * 24 * 60 * 60 * 1000),
    responsiblePerson: 'テスト担当者2',
    court: 'テスト地方裁判所'
  }];
  
  sendSlackReminder(testUrgentCase, testSoonCase);
  Logger.log('期日リマインドテスト送信完了');
}

function testSlackStaleOnly() {
  if (!CONFIG.SLACK_WEBHOOK_URL) {
    Logger.log('SLACK_WEBHOOK_URLが設定されていません');
    return;
  }
  
  const testStaleCase = [{
    id: 997,
    caseName: 'テスト未更新案件',
    caseNumber: 'テスト番号003',
    updatedDate: new Date(Date.now() - 35 * 24 * 60 * 60 * 1000),
    responsiblePerson: 'テスト担当者3'
  }];
  
  sendSlackStaleReminder(testStaleCase);
  Logger.log('未更新案件リマインドテスト送信完了');
}

function testSlackWeeklyOnly() {
  if (!CONFIG.SLACK_WEBHOOK_URL) {
    Logger.log('SLACK_WEBHOOK_URLが設定されていません');
    return;
  }
  
  const testReport = `=== 訴訟管理 週次レポート ===
期間: ${new Date().toLocaleDateString('ja-JP')}

📊 全体状況
総案件数: 15件

📈 ステータス別内訳
  係属中: 12件
  和解: 2件
  判決: 1件

⏰ 今後2週間の期日 (3件)
  2025/07/05 (6日後): テスト契約違反訴訟
  2025/07/10 (11日後): テスト損害賠償請求事件
  2025/07/15 (16日後): テスト知的財産権侵害事件`;
  
  sendSlackWeeklyReport(testReport);
  Logger.log('週次レポートテスト送信完了');
}

/**
 * 高度なSlack通知設定
 */
function setupAdvancedSlackNotifications() {
  // 複数チャンネルへの通知設定例
  const ADVANCED_SLACK_CONFIG = {
    URGENT_WEBHOOK: '', // 緊急用チャンネル
    GENERAL_WEBHOOK: '', // 一般通知用チャンネル
    REPORT_WEBHOOK: '', // レポート用チャンネル
    MENTION_USERS: [], // メンション対象ユーザーID
    CUSTOM_EMOJIS: {
      urgent: ':rotating_light:',
      warning: ':warning:',
      report: ':bar_chart:',
      stale: ':clock:'
    }
  };
  
  // 使用例をログに出力
  Logger.log('高度なSlack設定例:');
  Logger.log(JSON.stringify(ADVANCED_SLACK_CONFIG, null, 2));
  
  return ADVANCED_SLACK_CONFIG;
}

/**
 * 使用例とテスト関数
 */
function testSystem() {
  // セットアップガイドの表示
  setupGuide();

  // テストデータの追加例
  const testCase = {
    caseNumber: '令和7年（ワ）第1号',
    caseName: 'テスト損害賠償請求事件',
    caseType: '民事訴訟',
    plaintiff: '株式会社テスト',
    defendant: '山田太郎',
    court: '東京地方裁判所',
    lawyer: '弁護士 田中一郎',
    status: '係属中',
    filingDate: new Date('2025-01-15'),
    nextHearing: new Date('2025-07-05'), // 近い日付でテスト
    amount: '10,000,000',
    description: 'テスト用の訴訟案件です',
    responsiblePerson: '法務部 佐藤'
  };

  // 注意：実際にテストデータを追加する場合は下記のコメントアウトを外してください
  // const caseId = addLitigationCase(testCase);
  // Logger.log('テストケースID: ' + caseId);
}

// ===== UI/ダイアログ関連の関数 =====

/**
 * 新規案件登録ダイアログを表示
 */
function showAddCaseDialog() {
  const html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: Arial, sans-serif; padding: 20px; }
      .form-group { margin-bottom: 15px; }
      label { display: block; margin-bottom: 5px; font-weight: bold; }
      input, select, textarea { width: 100%; padding: 5px; }
      textarea { height: 80px; }
      button { background: #4285f4; color: white; padding: 10px 20px; border: none; cursor: pointer; margin-right: 10px; }
      button:hover { background: #357ae8; }
      .required { color: red; }
    </style>
    <h2>新規案件登録</h2>
    <div class="form-group">
      <label>事件番号 <span class="required">*</span></label>
      <input type="text" id="caseNumber" placeholder="例: 令和7年（ワ）第1号">
    </div>
    <div class="form-group">
      <label>事件名 <span class="required">*</span></label>
      <input type="text" id="caseName" placeholder="例: 損害賠償請求事件">
    </div>
    <div class="form-group">
      <label>事件種別</label>
      <select id="caseType">
        <option value="民事訴訟">民事訴訟</option>
        <option value="刑事訴訟">刑事訴訟</option>
        <option value="行政訴訟">行政訴訟</option>
        <option value="労働審判">労働審判</option>
        <option value="その他">その他</option>
      </select>
    </div>
    <div class="form-group">
      <label>原告</label>
      <input type="text" id="plaintiff">
    </div>
    <div class="form-group">
      <label>被告</label>
      <input type="text" id="defendant">
    </div>
    <div class="form-group">
      <label>裁判所</label>
      <input type="text" id="court" placeholder="例: 東京地方裁判所">
    </div>
    <div class="form-group">
      <label>担当弁護士</label>
      <input type="text" id="lawyer">
    </div>
    <div class="form-group">
      <label>ステータス</label>
      <select id="status">
        <option value="係属中">係属中</option>
        <option value="和解">和解</option>
        <option value="判決">判決</option>
        <option value="取下げ">取下げ</option>
        <option value="その他">その他</option>
      </select>
    </div>
    <div class="form-group">
      <label>提訴日</label>
      <input type="date" id="filingDate">
    </div>
    <div class="form-group">
      <label>次回期日</label>
      <input type="date" id="nextHearing">
    </div>
    <div class="form-group">
      <label>訴額</label>
      <input type="text" id="amount" placeholder="例: 10,000,000">
    </div>
    <div class="form-group">
      <label>概要</label>
      <textarea id="description"></textarea>
    </div>
    <div class="form-group">
      <label>担当者</label>
      <input type="text" id="responsiblePerson">
    </div>
    <button onclick="addCase()">登録</button>
    <button onclick="google.script.host.close()">キャンセル</button>

    <script>
      function addCase() {
        const caseData = {
          caseNumber: document.getElementById('caseNumber').value,
          caseName: document.getElementById('caseName').value,
          caseType: document.getElementById('caseType').value,
          plaintiff: document.getElementById('plaintiff').value,
          defendant: document.getElementById('defendant').value,
          court: document.getElementById('court').value,
          lawyer: document.getElementById('lawyer').value,
          status: document.getElementById('status').value,
          filingDate: document.getElementById('filingDate').value,
          nextHearing: document.getElementById('nextHearing').value,
          amount: document.getElementById('amount').value,
          description: document.getElementById('description').value,
          responsiblePerson: document.getElementById('responsiblePerson').value
        };

        if (!caseData.caseNumber || !caseData.caseName) {
          alert('事件番号と事件名は必須です');
          return;
        }

        google.script.run
          .withSuccessHandler(function(result) {
            alert('案件を登録しました (ID: ' + result + ')');
            google.script.host.close();
          })
          .withFailureHandler(function(error) {
            alert('エラー: ' + error);
          })
          .addLitigationCase(caseData);
      }
    </script>
  `)
  .setWidth(500)
  .setHeight(700);

  SpreadsheetApp.getUi().showModalDialog(html, '新規案件登録');
}

/**
 * 案件更新ダイアログを表示
 */
function showUpdateCaseDialog() {
  const spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = spreadsheet.getSheetByName(CONFIG.SHEET_NAME);
  const data = sheet.getDataRange().getValues();

  let options = '';
  for (let i = 1; i < data.length; i++) {
    const id = data[i][COLUMNS.ID];
    const name = data[i][COLUMNS.CASE_NAME];
    const number = data[i][COLUMNS.CASE_NUMBER];
    options += `<option value="${id}">${id}: ${name} (${number})</option>`;
  }

  const html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: Arial, sans-serif; padding: 20px; }
      .form-group { margin-bottom: 15px; }
      label { display: block; margin-bottom: 5px; font-weight: bold; }
      input, select, textarea { width: 100%; padding: 5px; }
      textarea { height: 80px; }
      button { background: #4285f4; color: white; padding: 10px 20px; border: none; cursor: pointer; margin-right: 10px; }
      button:hover { background: #357ae8; }
    </style>
    <h2>案件更新</h2>
    <div class="form-group">
      <label>更新する案件を選択</label>
      <select id="caseId" onchange="loadCaseData()">
        <option value="">選択してください</option>
        ${options}
      </select>
    </div>
    <div id="updateForm" style="display:none;">
      <div class="form-group">
        <label>ステータス</label>
        <select id="status">
          <option value="係属中">係属中</option>
          <option value="和解">和解</option>
          <option value="判決">判決</option>
          <option value="取下げ">取下げ</option>
          <option value="その他">その他</option>
        </select>
      </div>
      <div class="form-group">
        <label>次回期日</label>
        <input type="date" id="nextHearing">
      </div>
      <div class="form-group">
        <label>更新内容メモ</label>
        <textarea id="updateNote" placeholder="更新内容を記載"></textarea>
      </div>
      <button onclick="updateCase()">更新</button>
      <button onclick="google.script.host.close()">キャンセル</button>
    </div>

    <script>
      function loadCaseData() {
        const caseId = document.getElementById('caseId').value;
        if (caseId) {
          document.getElementById('updateForm').style.display = 'block';
        } else {
          document.getElementById('updateForm').style.display = 'none';
        }
      }

      function updateCase() {
        const caseId = document.getElementById('caseId').value;
        const status = document.getElementById('status').value;
        const nextHearing = document.getElementById('nextHearing').value;
        const updateNote = document.getElementById('updateNote').value;

        if (!caseId) {
          alert('案件を選択してください');
          return;
        }

        google.script.run
          .withSuccessHandler(function() {
            alert('案件を更新しました');
            google.script.host.close();
          })
          .withFailureHandler(function(error) {
            alert('エラー: ' + error);
          })
          .updateLitigationCase(caseId, status, nextHearing, updateNote);
      }
    </script>
  `)
  .setWidth(500)
  .setHeight(450);

  SpreadsheetApp.getUi().showModalDialog(html, '案件更新');
}

/**
 * 進行状況追加ダイアログを表示
 */
function showAddTimelineDialog() {
  const spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = spreadsheet.getSheetByName(CONFIG.SHEET_NAME);
  const data = sheet.getDataRange().getValues();

  let options = '';
  for (let i = 1; i < data.length; i++) {
    const id = data[i][COLUMNS.ID];
    const name = data[i][COLUMNS.CASE_NAME];
    options += `<option value="${id}">${id}: ${name}</option>`;
  }

  const html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: Arial, sans-serif; padding: 20px; }
      .form-group { margin-bottom: 15px; }
      label { display: block; margin-bottom: 5px; font-weight: bold; }
      input, select, textarea { width: 100%; padding: 5px; }
      textarea { height: 80px; }
      button { background: #4285f4; color: white; padding: 10px 20px; border: none; cursor: pointer; margin-right: 10px; }
      button:hover { background: #357ae8; }
    </style>
    <h2>進行状況追加</h2>
    <div class="form-group">
      <label>案件を選択</label>
      <select id="caseId">
        <option value="">選択してください</option>
        ${options}
      </select>
    </div>
    <div class="form-group">
      <label>タイトル</label>
      <input type="text" id="title" placeholder="例: 口頭弁論">
    </div>
    <div class="form-group">
      <label>内容</label>
      <textarea id="content" placeholder="詳細な内容"></textarea>
    </div>
    <div class="form-group">
      <label>担当者</label>
      <input type="text" id="responsiblePerson">
    </div>
    <div class="form-group">
      <label>次回アクション</label>
      <input type="text" id="nextAction" placeholder="例: 準備書面提出">
    </div>
    <div class="form-group">
      <label>期限</label>
      <input type="date" id="deadline">
    </div>
    <button onclick="addTimeline()">追加</button>
    <button onclick="google.script.host.close()">キャンセル</button>

    <script>
      function addTimeline() {
        const caseId = document.getElementById('caseId').value;
        const title = document.getElementById('title').value;
        const content = document.getElementById('content').value;
        const responsiblePerson = document.getElementById('responsiblePerson').value;
        const nextAction = document.getElementById('nextAction').value;
        const deadline = document.getElementById('deadline').value;

        if (!caseId || !title) {
          alert('案件とタイトルは必須です');
          return;
        }

        google.script.run
          .withSuccessHandler(function() {
            alert('進行状況を追加しました');
            google.script.host.close();
          })
          .withFailureHandler(function(error) {
            alert('エラー: ' + error);
          })
          .addTimelineEntry(caseId, title, content, responsiblePerson, nextAction, deadline);
      }
    </script>
  `)
  .setWidth(500)
  .setHeight(550);

  SpreadsheetApp.getUi().showModalDialog(html, '進行状況追加');
}

/**
 * 案件更新処理
 */
function updateLitigationCase(caseId, status, nextHearing, updateNote) {
  try {
    const spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = spreadsheet.getSheetByName(CONFIG.SHEET_NAME);
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (data[i][COLUMNS.ID] == caseId) {
        // ステータス更新
        if (status) {
          sheet.getRange(i + 1, COLUMNS.STATUS + 1).setValue(status);
        }
        // 次回期日更新
        if (nextHearing) {
          sheet.getRange(i + 1, COLUMNS.NEXT_HEARING + 1).setValue(new Date(nextHearing));
        }
        // 更新日時
        sheet.getRange(i + 1, COLUMNS.UPDATED_DATE + 1).setValue(new Date());

        // 進行状況に記録
        if (updateNote) {
          addTimelineEntry(caseId, '案件更新', updateNote, Session.getActiveUser().getEmail());
        }

        Logger.log('案件を更新しました: ID=' + caseId);
        return true;
      }
    }

    throw new Error('指定されたIDの案件が見つかりません');
  } catch (error) {
    Logger.log('案件更新エラー: ' + error.toString());
    throw error;
  }
}

/**
 * 週次レポートを生成して表示
 */
function generateAndShowWeeklyReport() {
  const report = generateWeeklyReport();
  const ui = SpreadsheetApp.getUi();
  ui.alert('週次レポート', report, ui.ButtonSet.OK);
}

/**
 * 期日チェックして表示
 */
function checkAndShowUpcomingDeadlines() {
  const urgentCases = checkUpcomingDeadlines(3);
  const soonCases = checkUpcomingDeadlines(7);

  let message = '📅 期日チェック結果\n\n';

  if (urgentCases.length > 0) {
    message += '🚨 3日以内の期日:\n';
    urgentCases.forEach(case_ => {
      const daysLeft = Math.ceil((case_.nextHearing - new Date()) / (1000 * 60 * 60 * 24));
      message += `・${case_.caseName} (${daysLeft}日後)\n`;
    });
    message += '\n';
  }

  if (soonCases.length > 0) {
    message += '⚠️ 7日以内の期日:\n';
    soonCases.forEach(case_ => {
      if (!urgentCases.some(urgent => urgent.id === case_.id)) {
        const daysLeft = Math.ceil((case_.nextHearing - new Date()) / (1000 * 60 * 60 * 24));
        message += `・${case_.caseName} (${daysLeft}日後)\n`;
      }
    });
  }

  if (urgentCases.length === 0 && soonCases.length === 0) {
    message += '期日が近い案件はありません。';
  }

  const ui = SpreadsheetApp.getUi();
  ui.alert('期日チェック', message, ui.ButtonSet.OK);
}

/**
 * 未更新案件チェックして表示
 */
function checkAndShowStaleCases() {
  const staleCases = checkStaleCases(30);

  let message = '📝 未更新案件チェック結果\n\n';

  if (staleCases.length > 0) {
    message += '30日以上更新されていない案件:\n\n';
    staleCases.forEach(case_ => {
      const daysSinceUpdate = Math.ceil((new Date() - case_.updatedDate) / (1000 * 60 * 60 * 24));
      message += `・${case_.caseName}\n`;
      message += `  最終更新: ${daysSinceUpdate}日前\n`;
      message += `  担当者: ${case_.responsiblePerson || '未設定'}\n\n`;
    });
  } else {
    message += '長期未更新の案件はありません。';
  }

  const ui = SpreadsheetApp.getUi();
  ui.alert('未更新案件チェック', message, ui.ButtonSet.OK);
}

/**
 * セットアップガイドを表示
 */
function showSetupGuide() {
  const guide = setupGuide();
  const ui = SpreadsheetApp.getUi();
  ui.alert('セットアップガイド', guide, ui.ButtonSet.OK);
}

/**
 * Slack連携ガイドを表示
 */
function showSlackGuide() {
  const guide = slackSetupGuide();
  const ui = SpreadsheetApp.getUi();
  ui.alert('Slack連携ガイド', guide, ui.ButtonSet.OK);
}