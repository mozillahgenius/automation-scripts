// ===== 設定項目 =====
const CONFIG = {
  // Slack Webhook URL
  SLACK_WEBHOOK_URL: PropertiesService.getScriptProperties().getProperty('SLACK_WEBHOOK_URL'),
  
  // スプレッドシートのID
  SPREADSHEET_ID: '1civv-zYpplWxg2Uu3esQaSuWO0wTGC2fRFPTqI1QY5Q',
    
  // シート名の定義
  SHEETS: {
    SETTINGS: '設定',
    FILE_LIST: 'ファイルリスト'
  },
  
  // Slack通知チャンネル
  SLACK_CHANNEL: '#clientwork-document',
  
  // 定期実行の間隔（分）
  CHECK_INTERVAL_MINUTES: 30,
  
  // 催促メッセージの設定
  REMINDER_SETTINGS: {
    // 催促メッセージを送信する時間（24時間形式）
    LUNCH_TIME: 12,     // 昼12時
    EVENING_TIME: 17,   // 夕方17時
    
    // 催促メッセージのチャンネル（通知チャンネルと同じまたは別チャンネル）
    REMINDER_CHANNEL: '#clientwork-document',
    
    // 土日は催促を送らない場合はtrue
    SKIP_WEEKENDS: true,
    
    // 祝日は催促を送らない場合はtrue（簡易実装）
    SKIP_HOLIDAYS: false
  }
};

// 催促メッセージのテンプレート
const REMINDER_MESSAGES = {
  LUNCH: {
    text: "📋 **お疲れ様です！ファイルアップロードのお時間です** 📋",
    attachments: [{
      color: '#36a64f',
      title: '🍱 お昼の定期リマインダー',
      text: 'ShareDriveへのファイルアップロードはお済みですか？\n午前中に作成した資料やデータがあれば、適切なフォルダにアップロードをお願いします。',
      fields: [
        {
          title: '📁 アップロード先',
          value: 'ShareDrive内の適切なプロジェクトフォルダ',
          short: true
        },
        {
          title: '⏰ 推奨タイミング',
          value: '作業完了後、すぐにアップロード',
          short: true
        }
      ],
      footer: 'ShareDrive管理Bot',
      footer_icon: 'https://platform.slack-edge.com/img/default_application_icon.png'
    }]
  },
  
  EVENING: {
    text: "📋 **本日もお疲れ様でした！最終確認のお時間です** 📋",
    attachments: [{
      color: '#ff9900',
      title: '🌅 夕方の定期リマインダー',
      text: '本日作成した資料やデータのShareDriveへのアップロードはお済みですか？\n明日のスムーズな業務のため、今日の成果物を適切にアップロードしましょう。',
      fields: [
        {
          title: '✅ 確認事項',
          value: '・今日作成したファイル\n・更新したドキュメント\n・会議資料など',
          short: true
        },
        {
          title: '📂 整理のポイント',
          value: '・適切なフォルダに分類\n・ファイル名を分かりやすく\n・バージョン管理を意識',
          short: true
        }
      ],
      footer: 'ShareDrive管理Bot - 今日も1日お疲れ様でした！',
      footer_icon: 'https://platform.slack-edge.com/img/default_application_icon.png'
    }]
  }
};

// ===== メイン処理 =====

/**
 * 初期設定とトリガーの作成
 */
function setup() {
  // スプレッドシートの初期化
  initializeSpreadsheet();
  
  // ファイル監視の定期実行トリガーの設定
  if ([1, 5, 10, 15, 30].includes(CONFIG.CHECK_INTERVAL_MINUTES)) {
    ScriptApp.newTrigger('checkAndNotify')
      .timeBased()
      .everyMinutes(CONFIG.CHECK_INTERVAL_MINUTES)
      .create();
  } else {
    const hours = Math.max(1, Math.round(CONFIG.CHECK_INTERVAL_MINUTES / 60));
    ScriptApp.newTrigger('checkAndNotify')
      .timeBased()
      .everyHours(hours)
      .create();
    console.log(`ファイル監視: ${hours}時間ごとに実行するよう設定しました`);
  }
  
  // 催促メッセージの定期実行トリガーの設定
  setupReminderTriggers();
  
  console.log('セットアップが完了しました');
}

/**
 * 催促メッセージのトリガーを設定
 */
function setupReminderTriggers() {
  // 既存の催促トリガーを削除
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'sendReminderMessage') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  // 昼の催促トリガー（毎日12時）
  ScriptApp.newTrigger('sendReminderMessage')
    .timeBased()
    .everyDays(1)
    .atHour(CONFIG.REMINDER_SETTINGS.LUNCH_TIME)
    .create();
  
  // 夕方の催促トリガー（毎日17時）
  ScriptApp.newTrigger('sendReminderMessage')
    .timeBased()
    .everyDays(1)
    .atHour(CONFIG.REMINDER_SETTINGS.EVENING_TIME)
    .create();
  
  console.log(`催促メッセージトリガーを設定しました: ${CONFIG.REMINDER_SETTINGS.LUNCH_TIME}時と${CONFIG.REMINDER_SETTINGS.EVENING_TIME}時`);
}

/**
 * 催促メッセージを送信
 */
function sendReminderMessage() {
  const now = new Date();
  const currentHour = now.getHours();
  const currentDay = now.getDay(); // 0=日曜, 6=土曜
  
  // 土日スキップの確認
  if (CONFIG.REMINDER_SETTINGS.SKIP_WEEKENDS && (currentDay === 0 || currentDay === 6)) {
    console.log('土日のため催促メッセージをスキップしました');
    return;
  }
  
  // 祝日スキップの確認（簡易実装）
  if (CONFIG.REMINDER_SETTINGS.SKIP_HOLIDAYS && isHoliday(now)) {
    console.log('祝日のため催促メッセージをスキップしました');
    return;
  }
  
  let messageTemplate;
  let messageType;
  
  // 時間に応じてメッセージを選択
  if (currentHour === CONFIG.REMINDER_SETTINGS.LUNCH_TIME) {
    messageTemplate = REMINDER_MESSAGES.LUNCH;
    messageType = 'lunch';
  } else if (currentHour === CONFIG.REMINDER_SETTINGS.EVENING_TIME) {
    messageTemplate = REMINDER_MESSAGES.EVENING;
    messageType = 'evening';
  } else {
    console.log('催促メッセージの送信時間ではありません');
    return;
  }
  
  // ShareDriveの状況を取得して追加情報を付加
  const activeDrives = getActiveDrives();
  const driveInfo = getDriveStatistics(activeDrives);
  
  // メッセージをカスタマイズ
  const customizedMessage = customizeReminderMessage(messageTemplate, driveInfo, messageType);
  
  // Slackに送信
  sendSlackMessage(customizedMessage, CONFIG.REMINDER_SETTINGS.REMINDER_CHANNEL);
  
  console.log(`${messageType}の催促メッセージを送信しました`);
}

/**
 * 催促メッセージをカスタマイズ
 */
function customizeReminderMessage(template, driveInfo, messageType) {
  const message = JSON.parse(JSON.stringify(template)); // ディープコピー
  
  // 統計情報を追加
  if (driveInfo.totalFiles > 0) {
    message.attachments[0].fields.push({
      title: '📊 現在の状況',
      value: `監視中ドライブ: ${driveInfo.driveCount}個\n総ファイル数: ${driveInfo.totalFiles}個\n最終更新: ${driveInfo.lastUpdate || '未確認'}`,
      short: false
    });
  }
  
  // メッセージタイプに応じた追加情報
  if (messageType === 'evening') {
    const todayFiles = getTodayUploadedFiles();
    if (todayFiles.length > 0) {
      message.attachments[0].fields.push({
        title: '🎉 本日アップロードされたファイル',
        value: `${todayFiles.length}個のファイルが本日アップロードされました`,
        short: true
      });
    }
  }
  
  return message;
}

/**
 * ShareDriveの統計情報を取得
 */
function getDriveStatistics(activeDrives) {
  let totalFiles = 0;
  let lastUpdate = null;
  
  try {
    const sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.SHEETS.FILE_LIST);
    const lastRow = sheet.getLastRow();
    
    if (lastRow > 1) {
      totalFiles = lastRow - 1; // ヘッダー行を除く
      
      // 最新のファイル更新日を取得
      const lastUpdateData = sheet.getRange(lastRow, 6).getValue();
      if (lastUpdateData) {
        lastUpdate = new Date(lastUpdateData).toLocaleString('ja-JP');
      }
    }
  } catch (error) {
    console.error('統計情報取得エラー:', error);
  }
  
  return {
    driveCount: activeDrives.length,
    totalFiles: totalFiles,
    lastUpdate: lastUpdate
  };
}

/**
 * 本日アップロードされたファイルを取得
 */
function getTodayUploadedFiles() {
  const today = new Date();
  const todayString = today.toDateString();
  const todayFiles = [];
  
  try {
    const sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.SHEETS.FILE_LIST);
    const lastRow = sheet.getLastRow();
    
    if (lastRow > 1) {
      const data = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
      
      data.forEach(row => {
        const uploadDate = new Date(row[3]); // アップロード日のカラム
        if (uploadDate.toDateString() === todayString) {
          todayFiles.push({
            driveName: row[0],
            fileName: row[1],
            uploadDate: uploadDate
          });
        }
      });
    }
  } catch (error) {
    console.error('本日ファイル取得エラー:', error);
  }
  
  return todayFiles;
}

/**
 * 簡易祝日判定（日本の主要祝日のみ）
 */
function isHoliday(date) {
  if (!CONFIG.REMINDER_SETTINGS.SKIP_HOLIDAYS) return false;
  
  const month = date.getMonth() + 1;
  const day = date.getDate();
  
  // 固定祝日の簡易チェック
  const fixedHolidays = [
    '1-1',   // 元日
    '2-11',  // 建国記念の日
    '4-29',  // 昭和の日
    '5-3',   // 憲法記念日
    '5-4',   // みどりの日
    '5-5',   // こどもの日
    '8-11',  // 山の日
    '11-3',  // 文化の日
    '11-23', // 勤労感謝の日
    '12-23'  // 天皇誕生日
  ];
  
  const dateString = `${month}-${day}`;
  return fixedHolidays.includes(dateString);
}

/**
 * Slackメッセージ送信（共通関数）
 */
function sendSlackMessage(messageData, channel = null) {
  const targetChannel = channel || CONFIG.SLACK_CHANNEL;
  
  const message = {
    channel: targetChannel,
    ...messageData
  };
  
  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(message)
  };
  
  try {
    const response = UrlFetchApp.fetch(CONFIG.SLACK_WEBHOOK_URL, options);
    console.log('Slackメッセージ送信成功:', response.getResponseCode());
  } catch (error) {
    console.error('Slackメッセージ送信エラー:', error);
    throw error;
  }
}

// ===== 手動実行用関数 =====

/**
 * 手動実行用：昼の催促メッセージをテスト送信
 */
function sendTestLunchReminder() {
  const activeDrives = getActiveDrives();
  const driveInfo = getDriveStatistics(activeDrives);
  const message = customizeReminderMessage(REMINDER_MESSAGES.LUNCH, driveInfo, 'lunch');
  
  sendSlackMessage(message);
  console.log('昼の催促メッセージ（テスト）を送信しました');
}

/**
 * 手動実行用：夕方の催促メッセージをテスト送信
 */
function sendTestEveningReminder() {
  const activeDrives = getActiveDrives();
  const driveInfo = getDriveStatistics(activeDrives);
  const message = customizeReminderMessage(REMINDER_MESSAGES.EVENING, driveInfo, 'evening');
  
  sendSlackMessage(message);
  console.log('夕方の催促メッセージ（テスト）を送信しました');
}

/**
 * 手動実行用：催促メッセージ設定の確認
 */
function checkReminderSettings() {
  console.log('=== 催促メッセージ設定 ===');
  console.log('昼の送信時間:', CONFIG.REMINDER_SETTINGS.LUNCH_TIME + '時');
  console.log('夕方の送信時間:', CONFIG.REMINDER_SETTINGS.EVENING_TIME + '時');
  console.log('送信チャンネル:', CONFIG.REMINDER_SETTINGS.REMINDER_CHANNEL);
  console.log('土日スキップ:', CONFIG.REMINDER_SETTINGS.SKIP_WEEKENDS ? 'あり' : 'なし');
  console.log('祝日スキップ:', CONFIG.REMINDER_SETTINGS.SKIP_HOLIDAYS ? 'あり' : 'なし');
  
  // 現在のトリガー状況
  const triggers = ScriptApp.getProjectTriggers();
  const reminderTriggers = triggers.filter(t => t.getHandlerFunction() === 'sendReminderMessage');
  console.log('設定済み催促トリガー数:', reminderTriggers.length);
  
  reminderTriggers.forEach((trigger, index) => {
    console.log(`トリガー${index + 1}:`, trigger.getTriggerSource(), trigger.getEventType());
  });
}

/**
 * 手動実行用：催促トリガーのリセット
 */
function resetReminderTriggers() {
  setupReminderTriggers();
  console.log('催促メッセージトリガーをリセットしました');
}

// ===== 既存の関数（元のコードから継承） =====

/**
 * スプレッドシートの初期化
 */
function initializeSpreadsheet() {
  const spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  
  // 設定シートの作成
  let settingsSheet = spreadsheet.getSheetByName(CONFIG.SHEETS.SETTINGS);
  if (!settingsSheet) {
    settingsSheet = spreadsheet.insertSheet(CONFIG.SHEETS.SETTINGS);
    
    // ヘッダー行の設定
    const headers = ['有効/無効', '共有ドライブ名', '共有ドライブID', '最終チェック日時', 'ステータス'];
    settingsSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    settingsSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    
    // 列幅の調整
    settingsSheet.setColumnWidth(1, 100);  // 有効/無効
    settingsSheet.setColumnWidth(2, 300);  // 共有ドライブ名
    settingsSheet.setColumnWidth(3, 350);  // 共有ドライブID
    settingsSheet.setColumnWidth(4, 200);  // 最終チェック日時
    settingsSheet.setColumnWidth(5, 200);  // ステータス
    
    // サンプル行を追加
    settingsSheet.getRange(2, 1, 1, 5).setValues([
      ['TRUE', 'サンプル共有ドライブ', 'YOUR_DRIVE_ID_HERE', '', '設定してください']
    ]);
    
    // チェックボックスの設定
    settingsSheet.getRange(2, 1, 100, 1).insertCheckboxes();
  }
  
  // ファイルリストシートの作成
  let fileListSheet = spreadsheet.getSheetByName(CONFIG.SHEETS.FILE_LIST);
  if (!fileListSheet) {
    fileListSheet = spreadsheet.insertSheet(CONFIG.SHEETS.FILE_LIST);
    
    // ヘッダー行の設定
    const headers = ['共有ドライブ名', 'ファイル名', 'URL', 'アップロード日', 'ファイルID', '最終更新日'];
    fileListSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    fileListSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    
    // 列幅の調整
    fileListSheet.setColumnWidth(1, 200); // 共有ドライブ名
    fileListSheet.setColumnWidth(2, 300); // ファイル名
    fileListSheet.setColumnWidth(3, 400); // URL
    fileListSheet.setColumnWidth(4, 150); // アップロード日
    fileListSheet.setColumnWidth(5, 250); // ファイルID
    fileListSheet.setColumnWidth(6, 150); // 最終更新日
  }
}

/**
 * 有効な共有ドライブの設定を取得
 */
function getActiveDrives() {
  const sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.SHEETS.SETTINGS);
  const lastRow = sheet.getLastRow();
  
  if (lastRow <= 1) return []; // ヘッダー行のみの場合
  
  const data = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
  const activeDrives = [];
  
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === true && data[i][2]) { // 有効かつIDが設定されている
      activeDrives.push({
        name: data[i][1] || `共有ドライブ${i + 1}`,
        id: data[i][2],
        row: i + 2 // スプレッドシート上の行番号
      });
    }
  }
  
  return activeDrives;
}

/**
 * 定期実行される処理
 */
function checkAndNotify() {
  const activeDrives = getActiveDrives();
  
  if (activeDrives.length === 0) {
    console.log('有効な共有ドライブが設定されていません');
    return;
  }
  
  const allNewFiles = [];
  const settingsSheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.SHEETS.SETTINGS);
  
  // 各共有ドライブをチェック
  for (const drive of activeDrives) {
    try {
      const newFiles = checkNewFilesInDrive(drive);
      
      if (newFiles.length > 0) {
        allNewFiles.push({
          driveName: drive.name,
          files: newFiles
        });
      }
      
      // 最終チェック日時とステータスを更新
      settingsSheet.getRange(drive.row, 4).setValue(new Date().toLocaleString('ja-JP'));
      settingsSheet.getRange(drive.row, 5).setValue('正常');
      
    } catch (error) {
      console.error(`${drive.name}のチェックでエラー:`, error);
      settingsSheet.getRange(drive.row, 4).setValue(new Date().toLocaleString('ja-JP'));
      settingsSheet.getRange(drive.row, 5).setValue(`エラー: ${error.toString()}`);
    }
  }
  
  // 新しいファイルがあればSlackに通知
  if (allNewFiles.length > 0) {
    notifyToSlack(allNewFiles);
  }
  
  // 全ドライブのファイルリストを更新
  updateAllFileLists(activeDrives);
}

/**
 * 特定の共有ドライブで新しいファイルをチェック
 */
function checkNewFilesInDrive(drive) {
  const sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.SHEETS.FILE_LIST);
  const lastRow = sheet.getLastRow();
  
  // この共有ドライブの既存のファイルIDを取得
  const existingFileIds = [];
  if (lastRow > 1) {
    const allData = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
    allData.forEach(row => {
      if (row[0] === drive.name && row[5]) {
        existingFileIds.push(row[5]);
      }
    });
  }
  
  // 過去1時間以内に作成されたファイルを検索
  const oneHourAgo = new Date();
  oneHourAgo.setHours(oneHourAgo.getHours() - 1);
  
  // Drive API v2ではcreatedDateを使用
  const query = `modifiedDate > '${oneHourAgo.toISOString()}' and trashed = false and mimeType != 'application/vnd.google-apps.folder'`;
  const response = Drive.Files.list({
    q: query,
    driveId: drive.id,
    corpora: 'drive',
    includeItemsFromAllDrives: true,
    supportsAllDrives: true,
    fields: 'items(id,title,createdDate,alternateLink)',
    maxResults: 100
  });
  
  const files = response.items || [];
  
  // 新しいファイルのみをフィルタリング（createdDateでチェック）
  return files.filter(file => {
    const createdDate = new Date(file.createdDate);
    return createdDate > oneHourAgo && !existingFileIds.includes(file.id);
  });
}

/**
 * Slackに通知（既存のファイルアップロード通知）
 */
function notifyToSlack(driveFilesList) {
  let attachments = [];
  
  driveFilesList.forEach(driveInfo => {
    driveInfo.files.forEach(file => {
      attachments.push({
        color: 'good',
        title: `📁 ${driveInfo.driveName}`,
        fields: [
          {
            title: 'ファイル名',
            value: file.title || file.name,
            short: true
          },
          {
            title: 'アップロード日時',
            value: new Date(file.createdDate || file.createdTime).toLocaleString('ja-JP'),
            short: true
          },
          {
            title: 'URL',
            value: file.alternateLink || file.webViewLink,
            short: false
          }
        ]
      });
    });
  });
  
  const message = {
    channel: CONFIG.SLACK_CHANNEL,
    text: '新しいファイルがShareDriveにアップロードされました\n\n:warning: ShareDrive内の適切なフォルダにデータを格納してください。',
    attachments: attachments
  };
  
  sendSlackMessage(message);
}

/**
 * 全ドライブのファイルリストを更新
 */
function updateAllFileLists(activeDrives) {
  const sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.SHEETS.FILE_LIST);
  
  // 既存のデータをクリア（ヘッダー行は残す）
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).clear();
  }
  
  const allData = [];
  
  // 各共有ドライブのファイルを取得
  for (const drive of activeDrives) {
    try {
      const files = getAllFilesInDrive(drive);
      
      for (const file of files) {
        allData.push([
          drive.name,
          file.title,
          file.alternateLink,
          new Date(file.createdDate).toLocaleString('ja-JP'),
          file.id,
          new Date(file.modifiedDate).toLocaleString('ja-JP')
        ]);
      }
    } catch (error) {
      console.error(`${drive.name}のファイルリスト更新でエラー:`, error);
    }
  }
  
  // データを書き込み
  if (allData.length > 0) {
    sheet.getRange(2, 1, allData.length, 6).setValues(allData);
  }
}

/**
 * 特定の共有ドライブ内の全ファイルを取得
 */
function getAllFilesInDrive(drive) {
  const allFiles = [];
  let pageToken = null;
  
  do {
    const response = Drive.Files.list({
      q: "trashed = false and mimeType != 'application/vnd.google-apps.folder'",
      driveId: drive.id,
      corpora: 'drive',
      includeItemsFromAllDrives: true,
      supportsAllDrives: true,
      fields: 'nextPageToken,items(id,title,createdDate,modifiedDate,alternateLink)',
      maxResults: 100,
      pageToken: pageToken
    });
    
    if (response.items) {
      allFiles.push(...response.items);
    }
    
    pageToken = response.nextPageToken;
  } while (pageToken);
  
  return allFiles;
}