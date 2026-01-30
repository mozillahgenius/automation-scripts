/**
 * 手動実行用：設定テスト
 */
function testConfiguration() {
  console.log('設定テストを実行します...');
  console.log('CONFIG:', CONFIG);
  
  // スプレッドシートアクセステスト
  try {
    const sheet = prepareSpreadsheet();
    console.log('✅ スプレッドシートアクセス: OK');
  } catch (error) {
    console.error('❌ スプレッドシートアクセス: エラー', error);
  }
  
  // Gmail API テスト
  try {
    const threads = GmailApp.search('in:inbox', 0, 1);
    console.log('✅ Gmail API: OK');
  } catch (error) {
    console.error('❌ Gmail API: エラー', error);
  }
  
  console.log('設定テスト完了');
}/**
 * 改良版Gmail未処理メール管理システム
 * 機能追加：エラーハンドリング、フィルタリング、優先度判定、Slack通知対応
 */

// 設定情報
const CONFIG = {
  SHEET_NAME: 'Report',
  EMAIL_RECIPIENT: 'admin@example-company.test', // 通知先メールアドレス
  SLACK_WEBHOOK_URL: '', // Slack Webhook URL（任意）
  EXCLUDE_SENDERS: ['noreply@', 'no-reply@', 'newsletter@'], // 除外したい送信者
  PRIORITY_KEYWORDS: ['緊急', '至急', '重要', 'urgent', 'important'], // 優先度判定キーワード
  MAX_EMAILS_TO_PROCESS: 50 // 処理する最大メール数
};

/**
 * メイン処理：未アーカイブメールのレポート作成と通知
 */
function reportUnarchivedEmails() {
  try {
    console.log('未処理メールレポート処理を開始します...');
    
    // Gmail検索クエリ：受信トレイの未読・既読両方を含む
    const searchQuery = 'in:inbox -is:draft';
    const threads = GmailApp.search(searchQuery, 0, CONFIG.MAX_EMAILS_TO_PROCESS);
    
    console.log(`検索結果: ${threads.length}件のスレッドが見つかりました`);
    
    // スプレッドシートの準備
    const sheet = prepareSpreadsheet();
    
    // メールデータの処理
    const emailData = processEmailThreads(threads);
    
    // スプレッドシートに書き込み
    writeToSpreadsheet(sheet, emailData);
    
    // 通知送信
    sendNotifications(emailData);
    
    console.log('処理が正常に完了しました');
    
  } catch (error) {
    console.error('エラーが発生しました:', error);
    sendErrorNotification(error);
  }
}

/**
 * スプレッドシートの準備
 */
function prepareSpreadsheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(CONFIG.SHEET_NAME);
  
  // シートが存在しない場合は作成
  if (!sheet) {
    sheet = spreadsheet.insertSheet(CONFIG.SHEET_NAME);
    console.log(`新しいシート "${CONFIG.SHEET_NAME}" を作成しました`);
  }
  
  // ヘッダーの設定
  const headers = ['送信者', '件名', '受信日時', 'メールリンク', '優先度', '日数経過'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // 前回のデータをクリア（ヘッダーを除く）
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, headers.length).clearContent();
  }
  
  return sheet;
}

/**
 * メールスレッドの処理
 */
function processEmailThreads(threads) {
  const emailData = [];
  const now = new Date();
  
  threads.forEach((thread, index) => {
    try {
      const messages = thread.getMessages();
      const lastMessage = messages[messages.length - 1];
      const sender = lastMessage.getFrom();
      
      // 除外対象の送信者をスキップ
      if (shouldExcludeSender(sender)) {
        return;
      }
      
      const subject = thread.getFirstMessageSubject();
      const date = lastMessage.getDate();
      const url = `https://mail.google.com/mail/u/0/#inbox/${thread.getId()}`;
      
      // 優先度判定
      const priority = determinePriority(subject, sender);
      
      // 経過日数計算
      const daysPassed = Math.floor((now - date) / (1000 * 60 * 60 * 24));
      
      emailData.push([sender, subject, date, url, priority, daysPassed]);
      
    } catch (error) {
      console.error(`スレッド ${index} の処理中にエラー:`, error);
    }
  });
  
  // 優先度と日数でソート（優先度高→日数多い順）
  emailData.sort((a, b) => {
    const priorityOrder = { '高': 0, '中': 1, '低': 2 };
    const priorityA = priorityOrder[a[4]] || 2;
    const priorityB = priorityOrder[b[4]] || 2;
    
    if (priorityA !== priorityB) {
      return priorityA - priorityB;
    }
    return b[5] - a[5]; // 日数の降順
  });
  
  return emailData;
}

/**
 * 送信者の除外判定
 */
function shouldExcludeSender(sender) {
  return CONFIG.EXCLUDE_SENDERS.some(exclude => 
    sender.toLowerCase().includes(exclude.toLowerCase())
  );
}

/**
 * 優先度の判定
 */
function determinePriority(subject, sender) {
  const text = (subject + ' ' + sender).toLowerCase();
  
  if (CONFIG.PRIORITY_KEYWORDS.some(keyword => 
    text.includes(keyword.toLowerCase())
  )) {
    return '高';
  }
  
  // その他の優先度判定ロジック
  if (sender.includes('@important-client.com')) {
    return '高';
  }
  
  if (text.includes('確認') || text.includes('承認')) {
    return '中';
  }
  
  return '低';
}

/**
 * スプレッドシートへの書き込み
 */
function writeToSpreadsheet(sheet, emailData) {
  if (emailData.length === 0) {
    console.log('書き込むデータがありません');
    return;
  }
  
  // データを書き込み
  sheet.getRange(2, 1, emailData.length, 6).setValues(emailData);
  
  // フォーマット設定
  formatSpreadsheet(sheet, emailData.length);
  
  console.log(`${emailData.length}件のメールデータを書き込みました`);
}

/**
 * スプレッドシートのフォーマット設定
 */
function formatSpreadsheet(sheet, dataLength) {
  // ヘッダーのフォーマット
  const headerRange = sheet.getRange(1, 1, 1, 6);
  headerRange.setBackground('#4285f4');
  headerRange.setFontColor('white');
  headerRange.setFontWeight('bold');
  
  if (dataLength > 0) {
    // 優先度列の条件付きフォーマット
    const priorityRange = sheet.getRange(2, 5, dataLength, 1);
    
    // 高優先度：赤背景
    const highPriorityRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('高')
      .setBackground('#ff9999')
      .setRanges([priorityRange])
      .build();
    
    // 中優先度：黄背景
    const medPriorityRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('中')
      .setBackground('#ffff99')
      .setRanges([priorityRange])
      .build();
    
    // 日数経過列の条件付きフォーマット（時間経過度による色分け）
    const daysRange = sheet.getRange(2, 6, dataLength, 1);
    
    // 14日以上：濃い赤
    const criticalDaysRule = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThanOrEqualTo(14)
      .setBackground('#cc0000')
      .setFontColor('white')
      .setRanges([daysRange])
      .build();
    
    // 7-13日：赤
    const urgentDaysRule = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberBetween(7, 13)
      .setBackground('#ff4444')
      .setFontColor('white')
      .setRanges([daysRange])
      .build();
    
    // 3-6日：オレンジ
    const warningDaysRule = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberBetween(3, 6)
      .setBackground('#ff9900')
      .setFontColor('white')
      .setRanges([daysRange])
      .build();
    
    // 1-2日：黄色
    const cautionDaysRule = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberBetween(1, 2)
      .setBackground('#ffcc00')
      .setRanges([daysRange])
      .build();
    
    // 0日（当日）：薄い緑
    const todayRule = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberEqualTo(0)
      .setBackground('#ccffcc')
      .setRanges([daysRange])
      .build();
    
    // 全ての条件付きフォーマットルールを適用
    const rules = [
      highPriorityRule, 
      medPriorityRule, 
      criticalDaysRule, 
      urgentDaysRule, 
      warningDaysRule, 
      cautionDaysRule, 
      todayRule
    ];
    sheet.setConditionalFormatRules(rules);
    
    // 行全体の条件付きフォーマット（最優先項目の強調）
    const dataRange = sheet.getRange(2, 1, dataLength, 6);
    
    // 高優先度かつ7日以上経過：行全体を赤背景
    const criticalRowRule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND($E2="高",$F2>=7)')
      .setBackground('#ffcccc')
      .setRanges([dataRange])
      .build();
    
    // 既存のルールに追加
    rules.push(criticalRowRule);
    sheet.setConditionalFormatRules(rules);
  }
  
  // 列幅の自動調整
  sheet.autoResizeColumns(1, 6);
  
  // 日数経過列のヘッダーに説明を追加
  const daysHeader = sheet.getRange(1, 6);
  daysHeader.setNote('色分け:\n🟢 0日(当日)\n🟡 1-2日\n🟠 3-6日\n🔴 7-13日\n⚫ 14日以上');
}

/**
 * 通知の送信
 */
function sendNotifications(emailData) {
  const totalCount = emailData.length;
  const highPriorityCount = emailData.filter(row => row[4] === '高').length;
  const oldEmailsCount = emailData.filter(row => row[5] > 3).length; // 3日以上経過
  
  // メール通知
  sendEmailNotification(totalCount, highPriorityCount, oldEmailsCount, emailData);
  
  // Slack通知（Webhook URLが設定されている場合）
  if (CONFIG.SLACK_WEBHOOK_URL) {
    sendSlackNotification(totalCount, highPriorityCount, oldEmailsCount, emailData);
  }
}

/**
 * メール通知の送信
 */
function sendEmailNotification(totalCount, highPriorityCount, oldEmailsCount, emailData) {
  if (totalCount === 0) {
    console.log('未処理メールがないため、通知をスキップします');
    return;
  }
  
  const spreadsheetUrl = SpreadsheetApp.getActiveSpreadsheet().getUrl();
  const subject = `【要対応】未処理メール通知 (${totalCount}件)`;
  
  // メール詳細リストの作成
  const emailDetails = createEmailDetailsList(emailData);
  
  const body = `未処理メールの状況をお知らせします。

概要:
• 総件数: ${totalCount}件
• 高優先度: ${highPriorityCount}件
• 3日以上経過: ${oldEmailsCount}件

未処理メール一覧:
${emailDetails}

詳細レポート:
${spreadsheetUrl}

対応が必要なメールがあります。早めの確認をお願いします。

---
このメールは自動送信されています。`;
  
  try {
    GmailApp.sendEmail(CONFIG.EMAIL_RECIPIENT, subject, body);
    console.log('メール通知を送信しました');
  } catch (error) {
    console.error('メール送信エラー:', error);
  }
}

/**
 * メール詳細リストの作成
 */
function createEmailDetailsList(emailData) {
  if (emailData.length === 0) {
    return '未処理メールはありません。';
  }
  
  let details = '';
  
  // 最大10件まで表示（長くなりすぎないよう制限）
  const displayData = emailData.slice(0, 10);
  
  displayData.forEach((row, index) => {
    const [sender, subject, date, url, priority, daysPassed] = row;
    const urgencyIcon = getUrgencyIcon(daysPassed, priority);
    const shortSender = extractSenderName(sender);
    const truncatedSubject = subject.length > 50 ? subject.substring(0, 50) + '...' : subject;
    
    details += `${index + 1}. [${priority}優先度] ${truncatedSubject}\n`;
    details += `   送信者: ${shortSender} | ${daysPassed}日前 ${urgencyIcon}\n`;
    details += `   リンク: ${url}\n\n`;
  });
  
  if (emailData.length > 10) {
    details += `※ 他に${emailData.length - 10}件の未処理メールがあります。\n`;
    details += `詳細はスプレッドシートをご確認ください。\n`;
  }
  
  return details;
}

/**
 * 緊急度アイコンの取得
 */
function getUrgencyIcon(daysPassed, priority) {
  if (priority === '高') {
    return daysPassed >= 7 ? '【超緊急】' : daysPassed >= 3 ? '【緊急】' : '【要注意】';
  } else if (priority === '中') {
    return daysPassed >= 7 ? '【緊急】' : daysPassed >= 5 ? '【要注意】' : '【確認】';
  } else {
    return daysPassed >= 14 ? '【要注意】' : daysPassed >= 7 ? '【確認】' : '【通常】';
  }
}

/**
 * 送信者名の抽出（メールアドレスから名前部分を取得）
 */
function extractSenderName(sender) {
  // "Name <email@domain.com>" 形式から名前を抽出
  const nameMatch = sender.match(/^([^<]+)\s*</);
  if (nameMatch) {
    return nameMatch[1].trim();
  }
  
  // メールアドレスのみの場合、@マークより前を取得
  const emailMatch = sender.match(/([^@\s]+)@/);
  if (emailMatch) {
    return emailMatch[1];
  }
  
  return sender.length > 30 ? sender.substring(0, 30) + '...' : sender;
}

/**
 * Slack通知の送信
 */
function sendSlackNotification(totalCount, highPriorityCount, oldEmailsCount, emailData) {
  const spreadsheetUrl = SpreadsheetApp.getActiveSpreadsheet().getUrl();
  
  // 緊急度の高い上位5件のメールを表示
  const topUrgentEmails = emailData.slice(0, 5).map((row, index) => {
    const [sender, subject, date, url, priority, daysPassed] = row;
    const urgencyIcon = getUrgencyIcon(daysPassed, priority);
    const shortSender = extractSenderName(sender);
    const truncatedSubject = subject.length > 40 ? subject.substring(0, 40) + '...' : subject;
    
    return `${index + 1}. [${priority}] ${truncatedSubject}\n送信者: ${shortSender} (${daysPassed}日前) ${urgencyIcon}`;
  }).join('\n\n');
  
  const payload = {
    text: `未処理メール通知`,
    attachments: [{
      color: highPriorityCount > 0 ? 'danger' : (oldEmailsCount > 0 ? 'warning' : 'good'),
      fields: [
        { title: '総件数', value: `${totalCount}件`, short: true },
        { title: '高優先度', value: `${highPriorityCount}件`, short: true },
        { title: '3日以上経過', value: `${oldEmailsCount}件`, short: true }
      ],
      text: totalCount > 0 ? `要注意メール TOP5:\n${topUrgentEmails}` : '未処理メールはありません！',
      actions: [{
        type: 'button',
        text: '詳細レポートを見る',
        url: spreadsheetUrl
      }]
    }]
  };
  
  try {
    const response = UrlFetchApp.fetch(CONFIG.SLACK_WEBHOOK_URL, {
      method: 'POST',
      contentType: 'application/json',
      payload: JSON.stringify(payload)
    });
    
    if (response.getResponseCode() === 200) {
      console.log('Slack通知を送信しました');
    } else {
      console.error('Slack通知送信失敗:', response.getResponseCode());
    }
  } catch (error) {
    console.error('Slack通知エラー:', error);
  }
}

/**
 * エラー通知の送信
 */
function sendErrorNotification(error) {
  const subject = '【エラー】Gmail未処理メール管理システム';
  const body = `
システムでエラーが発生しました。

エラー詳細:
${error.toString()}

スタックトレース:
${error.stack}

確認をお願いします。
  `.trim();
  
  try {
    GmailApp.sendEmail(CONFIG.EMAIL_RECIPIENT, subject, body);
  } catch (e) {
    console.error('エラー通知の送信に失敗:', e);
  }
}

/**
 * 複数時間でのトリガー設定（手動実行用）
 * 10時から20時まで1時間おきにトリガーを設定
 */
function setupHourlyTriggers() {
  // 既存のトリガーを削除
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'reportUnarchivedEmails') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  // 10時から20時まで1時間おきにトリガーを設定
  for (let hour = 10; hour <= 20; hour++) {
    ScriptApp.newTrigger('reportUnarchivedEmails')
      .timeBased()
      .everyDays(1)
      .atHour(hour)
      .create();
    
    console.log(`${hour}時のトリガーを設定しました`);
  }
  
  console.log('すべてのトリガー設定が完了しました（10時-20時、1時間おき）');
}

/**
 * トリガーの確認（手動実行用）
 */
function checkTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  const emailTriggers = triggers.filter(trigger => 
    trigger.getHandlerFunction() === 'reportUnarchivedEmails'
  );
  
  console.log(`設定済みトリガー数: ${emailTriggers.length}`);
  emailTriggers.forEach((trigger, index) => {
    const eventType = trigger.getEventType();
    if (eventType === ScriptApp.EventType.CLOCK) {
      console.log(`トリガー${index + 1}: 毎日実行`);
    }
  });
}