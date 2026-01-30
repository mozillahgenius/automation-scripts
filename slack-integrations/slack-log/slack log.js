/**
 * ========================================
 * Slack ログ収集 & 業務マニュアル自動生成ツール
 * ========================================
 * 
 * 【Slack アプリのセットアップガイド】
 * 
 * 1. Slack App の作成
 *    a. https://api.slack.com/apps にアクセス
 *    b. "Create New App" → "From scratch" を選択
 *    c. App Name（例: "Log Collector"）とワークスペースを選択
 * 
 * 2. Bot Token Scopes の設定（OAuth & Permissions）
 *    必要な権限を追加:
 *    - channels:history     （パブリックチャンネルの履歴読み取り）
 *    - channels:read        （パブリックチャンネル情報の読み取り）
 *    - groups:history       （プライベートチャンネルの履歴読み取り）
 *    - groups:read          （プライベートチャンネル情報の読み取り）
 *    - users:read           （ユーザー情報の読み取り）
 *    - users:read.email     （ユーザーのメールアドレス読み取り）※任意
 * 
 * 3. アプリのインストール
 *    a. "Install to Workspace" をクリック
 *    b. 権限を確認して "Allow"
 *    c. Bot User OAuth Token をコピー（xoxb-で始まる文字列）
 * 
 * 4. チャンネルへの招待
 *    - ログを取得したい各チャンネルでコマンド入力: /invite @[your-app-name]
 *    - プライベートチャンネルは必ず招待が必要
 * 
 * 【Google スプレッドシート & ドキュメントの設定】
 * 
 * 1. Google スプレッドシート
 *    a. 新規スプレッドシートを作成
 *    b. URLから SPREADSHEET_ID をコピー
 *       https://docs.google.com/spreadsheets/d/[SPREADSHEET_ID]/edit
 * 
 * 2. Google ドキュメント
 *    a. 新規ドキュメントを作成
 *    b. URLから DOCUMENT_ID をコピー
 *       https://docs.google.com/document/d/[DOCUMENT_ID]/edit
 * 
 * 3. Google Apps Script の設定
 *    a. スプレッドシートで「拡張機能」→「Apps Script」
 *    b. このコードを貼り付け
 *    c. 定期実行トリガーを設定（例: 1時間ごと）
 * 
 * 【OpenAI API の設定】（業務マニュアル自動生成用）
 * 
 * 1. https://platform.openai.com/api-keys でAPIキーを作成
 * 2. 使用量制限の設定を推奨
 * 
 * 【Gmail 送信設定】
 * 
 * 1. Google Apps Script で MailApp の権限を許可
 * 2. 初回実行時に認証が必要
 * 
 * ========================================
 */

// 必須設定項目（*** を実際の値に置き換えてください）
const SLACK_BOT_TOKEN = 'YOUR_SLACK_BOT_TOKEN'; // Slack Bot User OAuth Token (xoxb-で始まる)
const GOOGLE_DOC_ID = '1dkxrY8mtC28bWyDtxm0NVDohlESzNwqHJqq4PQFimqY'; // Google ドキュメントのID
const LOG_SHEET_NAME = 'slack_log'; // スプレッドシートのシート名
const LAST_TS_SHEET_NAME = 'slack_channel_last_ts'; // 追加
const MANUAL_SHEET_NAME = 'business_manual'; // 業務マニュアル用シート
const FAQ_SHEET_NAME = 'faq_list'; // FAQ用シート

// オプション設定項目（使用する場合は *** を実際の値に置き換え）
const OPENAI_API_KEY = 'YOUR_OPENAI_API_KEY'; // OpenAI APIキー（業務マニュアル生成・要約作成用）
const NOTIFICATION_EMAIL = 'your-email@example.com'; // 日次要約送信先メールアドレス

/**
 * スプレッドシートを開いた時に実行される関数
 * カスタムメニューを追加
 */
function onOpen() {
  // Googleドキュメントを初期化（マニュアル専用）
  initializeGoogleDoc();
  
  const ui = SpreadsheetApp.getUi();
  
  ui.createMenu('📋 Slack ログツール')
    // メイン機能
    .addItem('▶️ Slackログを取得', 'fetchAndAppendAllChannels')
    .addSeparator()
    
    // 業務マニュアル・FAQ生成
    .addSubMenu(ui.createMenu('📚 マニュアル・FAQ生成')
      .addItem('✨ 改良版：独立タスクで生成', 'generateManualAndFAQImproved')
      .addSeparator()
      .addItem('🤖 自動判別で生成（旧版）', 'generateManualAndFAQ')
      .addSeparator()
      .addItem('📖 マニュアルのみ生成', 'manualGenerateBusinessManual')
      .addItem('❓ FAQのみ生成', 'manualGenerateFAQ')
      .addSeparator()
      .addItem('📅 過去7日間から生成', 'generateManualForPeriod')
      .addItem('📢 特定チャンネルから生成...', 'showChannelSelectionDialog')
      .addItem('🔍 期間を指定して生成...', 'showPeriodSelectionDialog'))
    
    // レポート送信
    .addSubMenu(ui.createMenu('📧 レポート送信')
      .addItem('日次要約メールを送信', 'manualSendDailySummary')
      .addItem('週次レポートを送信', 'sendWeeklySummary')
      .addItem('月次レポートを送信', 'sendMonthlySummary'))
    
    // セットアップ＆管理
    .addSeparator()
    .addSubMenu(ui.createMenu('⚙️ セットアップ')
      .addItem('初期セットアップウィザード', 'fullSetupWizard')
      .addItem('スプレッドシート初期化', 'setupSpreadsheet')
      .addItem('設定チェック', 'checkSetup')
      .addItem('トリガー設定', 'setupTriggers'))
    
    // データ管理
    .addSubMenu(ui.createMenu('🗂️ データ管理')
      .addItem('データをバックアップ', 'backupData')
      .addItem('全シートをリセット', 'resetAllSheetsWithConfirmation')
      .addItem('重複データをクリーンアップ', 'cleanupDuplicates')
      .addItem('古いデータをアーカイブ', 'archiveOldData'))
    
    // ヘルプ
    .addSeparator()
    .addItem('❓ ヘルプ・使い方', 'showHelp')
    .addItem('ℹ️ バージョン情報', 'showAbout')
    .addSeparator()
    .addItem('📄 業務マニュアルドキュメントを開く', 'openManualDocument')
    .addToUi();
    
  // サイドバーに状態を表示（オプション）
  showStatusSidebar();
}

/**
 * 特定チャンネル選択ダイアログ
 */
function showChannelSelectionDialog() {
  const html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: Arial, sans-serif; padding: 20px; }
      input { width: 100%; padding: 10px; margin: 10px 0; }
      button { background: #4a86e8; color: white; padding: 10px 20px; border: none; cursor: pointer; }
      button:hover { background: #3b6ec6; }
    </style>
    <h3>チャンネルを指定</h3>
    <input type="text" id="channelName" placeholder="例: general">
    <button onclick="generate()">生成</button>
    <script>
      function generate() {
        const channel = document.getElementById('channelName').value;
        if (channel) {
          google.script.run.generateManualForChannel(channel);
          google.script.host.close();
        }
      }
    </script>
  `).setWidth(300).setHeight(200);
  
  SpreadsheetApp.getUi().showModalDialog(html, 'チャンネル選択');
}

/**
 * 期間選択ダイアログ
 */
function showPeriodSelectionDialog() {
  const html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: Arial, sans-serif; padding: 20px; }
      input { width: 100%; padding: 10px; margin: 10px 0; }
      button { background: #4a86e8; color: white; padding: 10px 20px; border: none; cursor: pointer; }
      button:hover { background: #3b6ec6; }
    </style>
    <h3>期間を指定</h3>
    <label>開始日:</label>
    <input type="date" id="startDate">
    <label>終了日:</label>
    <input type="date" id="endDate">
    <button onclick="generate()">生成</button>
    <script>
      function generate() {
        const start = document.getElementById('startDate').value;
        const end = document.getElementById('endDate').value;
        if (start && end) {
          google.script.run.generateManualForPeriod(new Date(start), new Date(end));
          google.script.host.close();
        }
      }
    </script>
  `).setWidth(300).setHeight(300);
  
  SpreadsheetApp.getUi().showModalDialog(html, '期間選択');
}

/**
 * ステータスサイドバーを表示
 */
function showStatusSidebar() {
  try {
    const sheet = getOrCreateLogSheet();
    const lastRow = sheet.getLastRow();
    const totalMessages = lastRow > 1 ? lastRow - 1 : 0;
    
    const manualSheet = getOrCreateManualSheet();
    const totalManuals = manualSheet.getLastRow() > 1 ? manualSheet.getLastRow() - 1 : 0;
    
    const html = HtmlService.createHtmlOutput(`
      <style>
        body { font-family: Arial, sans-serif; padding: 20px; }
        .stat { background: #f8f9fa; padding: 15px; margin: 10px 0; border-radius: 5px; }
        .stat-number { font-size: 24px; font-weight: bold; color: #4a86e8; }
        .stat-label { color: #666; font-size: 12px; margin-top: 5px; }
        button { background: #4a86e8; color: white; padding: 10px; width: 100%; border: none; cursor: pointer; margin: 5px 0; }
        button:hover { background: #3b6ec6; }
      </style>
      <h2>📊 システム状態</h2>
      
      <div class="stat">
        <div class="stat-number">${totalMessages}</div>
        <div class="stat-label">総メッセージ数</div>
      </div>
      
      <div class="stat">
        <div class="stat-number">${totalManuals}</div>
        <div class="stat-label">業務マニュアル数</div>
      </div>
      
      <div class="stat">
        <div class="stat-number">${getThreadCount()}</div>
        <div class="stat-label">スレッド数</div>
      </div>
      
      <h3>クイックアクション</h3>
      <button onclick="google.script.run.fetchAndAppendAllChannels()">ログ取得</button>
      <button onclick="google.script.run.manualGenerateBusinessManual()">マニュアル生成</button>
      <button onclick="google.script.run.manualSendDailySummary()">日次レポート送信</button>
      
      <script>
        // 5秒ごとに統計を更新
        setInterval(() => {
          google.script.run.withSuccessHandler(updateStats).getStatistics();
        }, 5000);
        
        function updateStats(stats) {
          // 統計情報を更新
        }
      </script>
    `).setTitle('Slack ログツール');
    
    SpreadsheetApp.getUi().showSidebar(html);
  } catch (error) {
    console.log('サイドバー表示エラー:', error);
  }
}

/**
 * 週次サマリーを送信
 */
function sendWeeklySummary() {
  const endDate = new Date();
  const startDate = new Date(Date.now() - 7 * 24 * 60 * 60 * 1000);
  
  const sheet = getOrCreateLogSheet();
  const data = sheet.getDataRange().getValues();
  
  const weekMessages = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][6] instanceof Date) {
      const msgDate = data[i][6];
      if (msgDate >= startDate && msgDate <= endDate) {
        weekMessages.push({
          channel: data[i][1],
          user: data[i][4],
          text: data[i][5],
          ts: data[i][2],
          threadTs: data[i][3]
        });
      }
    }
  }
  
  if (weekMessages.length > 0) {
    const manualInfo = weekMessages.length >= 10 ? generateBusinessManual(weekMessages) : null;
    sendDailySummaryEmail(weekMessages, manualInfo);
    SpreadsheetApp.getUi().alert('週次レポートを送信しました');
  } else {
    SpreadsheetApp.getUi().alert('今週のメッセージがありません');
  }
}

/**
 * 月次サマリーを送信
 */
function sendMonthlySummary() {
  const endDate = new Date();
  const startDate = new Date(endDate.getFullYear(), endDate.getMonth(), 1);
  
  const sheet = getOrCreateLogSheet();
  const data = sheet.getDataRange().getValues();
  
  const monthMessages = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][6] instanceof Date) {
      const msgDate = data[i][6];
      if (msgDate >= startDate && msgDate <= endDate) {
        monthMessages.push({
          channel: data[i][1],
          user: data[i][4],
          text: data[i][5],
          ts: data[i][2],
          threadTs: data[i][3]
        });
      }
    }
  }
  
  if (monthMessages.length > 0) {
    const manualInfo = monthMessages.length >= 20 ? generateBusinessManual(monthMessages) : null;
    sendDailySummaryEmail(monthMessages, manualInfo);
    SpreadsheetApp.getUi().alert('月次レポートを送信しました');
  } else {
    SpreadsheetApp.getUi().alert('今月のメッセージがありません');
  }
}

/**
 * 重複データをクリーンアップ
 */
function cleanupDuplicates() {
  const sheet = getOrCreateLogSheet();
  const data = sheet.getDataRange().getValues();
  const uniqueIds = new Set();
  const rowsToDelete = [];
  
  for (let i = data.length - 1; i >= 1; i--) {
    const messageId = data[i][7];
    if (uniqueIds.has(messageId)) {
      rowsToDelete.push(i + 1);
    } else {
      uniqueIds.add(messageId);
    }
  }
  
  // 重複行を削除
  rowsToDelete.forEach(row => {
    sheet.deleteRow(row);
  });
  
  SpreadsheetApp.getUi().alert(`${rowsToDelete.length}件の重複データを削除しました`);
}

/**
 * 古いデータをアーカイブ
 */
function archiveOldData() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.alert(
    '確認',
    '90日以上前のデータをアーカイブしますか？',
    ui.ButtonSet.YES_NO
  );
  
  if (result !== ui.Button.YES) return;
  
  const sheet = getOrCreateLogSheet();
  const data = sheet.getDataRange().getValues();
  const archiveSheet = getOrCreateArchiveSheet();
  const cutoffDate = new Date(Date.now() - 90 * 24 * 60 * 60 * 1000);
  
  let archivedCount = 0;
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][6] instanceof Date && data[i][6] < cutoffDate) {
      archiveSheet.appendRow(data[i]);
      sheet.deleteRow(i + 1);
      archivedCount++;
    }
  }
  
  ui.alert(`${archivedCount}件のデータをアーカイブしました`);
}

/**
 * アーカイブシートを取得または作成
 */
function getOrCreateArchiveSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('archive');
  if (!sheet) {
    sheet = ss.insertSheet('archive');
    const headers = ['channel_id', 'channel_name', 'timestamp', 'thread_ts', 'user_name', 'message', 'date', 'message_id'];
    sheet.appendRow(headers);
  }
  return sheet;
}

/**
 * 確認付きリセット
 */
function resetAllSheetsWithConfirmation() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.alert(
    '⚠️ 警告',
    'すべてのデータをリセットしますか？この操作は取り消せません。',
    ui.ButtonSet.YES_NO
  );
  
  if (result === ui.Button.YES) {
    backupData();
    resetAllSheets();
    ui.alert('データをリセットしました（バックアップを作成済み）');
  }
}

/**
 * ヘルプを表示
 */
function showHelp() {
  const html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: Arial, sans-serif; padding: 20px; line-height: 1.6; }
      h2 { color: #4a86e8; }
      h3 { color: #666; }
      code { background: #f5f5f5; padding: 2px 5px; }
    </style>
    <h2>Slack ログツール - 使い方</h2>
    
    <h3>基本操作</h3>
    <ul>
      <li><strong>Slackログを取得</strong>: 最新のメッセージを取得します</li>
      <li><strong>業務マニュアル生成</strong>: AIが会話から業務マニュアルを作成</li>
      <li><strong>レポート送信</strong>: 日次・週次・月次のサマリーをメール送信</li>
    </ul>
    
    <h3>初期設定</h3>
    <ol>
      <li>「セットアップ」→「初期セットアップウィザード」を実行</li>
      <li>Slack Bot Tokenを設定</li>
      <li>OpenAI APIキーを設定（オプション）</li>
      <li>メールアドレスを設定</li>
    </ol>
    
    <h3>定期実行</h3>
    <p>「セットアップ」→「トリガー設定」で自動実行を設定できます</p>
    
    <h3>トラブルシューティング</h3>
    <ul>
      <li>エラーが出る場合は「設定チェック」を実行</li>
      <li>重複データは「データ管理」→「重複データをクリーンアップ」</li>
    </ul>
  `).setWidth(500).setHeight(600);
  
  SpreadsheetApp.getUi().showModalDialog(html, 'ヘルプ');
}

/**
 * バージョン情報を表示
 */
function showAbout() {
  const ui = SpreadsheetApp.getUi();
  ui.alert(
    'Slack ログ収集ツール',
    'バージョン: 2.1.0\\n' +
    '\\n機能:\\n' +
    '• Slackメッセージの自動収集\\n' +
    '• AIによる業務マニュアル生成\\n' +
    '• エグゼクティブサマリー作成\\n' +
    '• HTMLメールレポート\\n' +
    '• スレッド統合処理\\n' +
    '• Googleドキュメントへのマニュアル記録\\n' +
    '\\n© 2024 COMPANY_X',
    ui.ButtonSet.OK
  );
}

// 業務マニュアルドキュメントを開く
function openManualDocument() {
  const url = `https://docs.google.com/document/d/${GOOGLE_DOC_ID}/edit`;
  const html = HtmlService.createHtmlOutput(`
    <script>
      window.open('${url}', '_blank');
      google.script.host.close();
    </script>
  `);
  SpreadsheetApp.getUi().showModalDialog(html, 'ドキュメントを開いています...');
}

/**
 * 統計情報を取得（サイドバー更新用）
 */
function getStatistics() {
  const sheet = getOrCreateLogSheet();
  const manualSheet = getOrCreateManualSheet();
  
  return {
    totalMessages: sheet.getLastRow() - 1,
    totalManuals: manualSheet.getLastRow() - 1,
    lastUpdate: new Date().toLocaleString()
  };
}

// スレッド数を取得
function getThreadCount() {
  const sheet = getOrCreateLogSheet();
  const data = sheet.getDataRange().getValues();
  const threads = new Set();
  
  for (let i = 1; i < data.length; i++) {
    const threadTs = data[i][3];
    if (threadTs && data[i][5] && data[i][5].includes('[スレッド開始]')) {
      threads.add(`${data[i][0]}_${threadTs}`);
    }
  }
  
  return threads.size;
}

function fetchAndAppendAllChannels() {
  const channels = getSlackChannels();
  const sheet = getOrCreateLogSheet();
  const lastTsSheet = getOrCreateLastTsSheet();
  const userCache = {};
  const todayMessages = [];
  
  console.log(`処理開始: ${channels.length}個のチャンネル`);

  channels.forEach((channel, index) => {
    console.log(`処理中: ${index + 1}/${channels.length} - ${channel.name}`);
    
    // API呼び出し間隔を空ける（レート制限対策）
    if (index > 0) {
      Utilities.sleep(1000); // 1秒待機
    }
    
    const lastFetchedTs = getLastFetchedTs(lastTsSheet, channel.id);
    const messages = getChannelMessages(channel.id, lastFetchedTs);
    if (!messages) return;

    let maxTs = lastFetchedTs;

    messages.reverse().forEach(msg => {
      if (!msg.text) return;
      
      // スレッドの返信はスキップ（親メッセージでまとめて処理）
      if (msg.thread_ts && msg.thread_ts !== msg.ts) {
        return; // スレッド返信は後でまとめて処理
      }
      
      // ユニークIDを生成（チャンネルID + タイムスタンプ）
      const messageId = `${channel.id}_${msg.ts}`;
      
      // 重複チェック（メッセージIDで確認）
      if (isMessageExists(sheet, messageId)) {
        return; // 既に処理済み
      }
      
      const realName = getRealName(msg.user, userCache, msg);
      let fullText = replaceMentionsWithRealNames(msg.text, userCache);
      
      // スレッドがある場合は、すべての返信を取得してまとめる
      if (msg.thread_ts && msg.reply_count > 0) {
        const threadReplies = getThreadReplies(channel.id, msg.thread_ts, userCache);
        if (threadReplies && threadReplies.length > 0) {
          fullText = formatThreadConversation(realName, fullText, threadReplies);
        }
      }

      // Googleドキュメントへの通常ログ記録はスキップ（マニュアルのみ記録）
      // appendToGoogleDoc(channel.name, msg, realName, fullText);
      const date = new Date(Number(msg.ts.split('.')[0]) * 1000);
      const threadTs = msg.thread_ts || msg.ts;
      sheet.appendRow([channel.id, channel.name, msg.ts, threadTs, realName, fullText, date, messageId]);
      
      // 今日のメッセージを収集
      todayMessages.push({
        channel: channel.name,
        user: realName,
        text: fullText,
        ts: msg.ts,
        threadTs: threadTs
      });
      
      if (Number(msg.ts) > Number(maxTs)) {
        maxTs = msg.ts;
      }
    });

    setLastFetchedTs(lastTsSheet, channel.id, maxTs);
  });
  
  // 業務マニュアル生成とメール送信
  if (todayMessages.length > 0) {
    console.log(`本日のメッセージ: ${todayMessages.length}件`);
    
    // メッセージがあればコンテンツ生成を試みる
    let manualInfo = null;
    let faqInfo = null;
    
    if (todayMessages.length >= 1) { // 1件以上で処理
      console.log('コンテンツ生成を開始...');
      try {
        const contentResult = generateContentWithAI(todayMessages);
        if (contentResult) {
          console.log(`コンテンツ生成成功: マニュアル${contentResult.manualCount}件, FAQ${contentResult.faqCount}件`);
          
          // マニュアル情報を取得
          if (contentResult.manualCount > 0) {
            manualInfo = {
              count: contentResult.manualCount,
              manuals: getLatestManualsFromSheet(contentResult.manualCount)
            };
          }
          
          // FAQ情報を取得
          if (contentResult.faqCount > 0) {
            faqInfo = {
              count: contentResult.faqCount,
              faqs: getFAQsFromSheet(contentResult.faqCount)
            };
          }
        } else {
          console.log('コンテンツ生成結果はnull');
        }
      } catch (error) {
        console.error('コンテンツ生成中にエラー:', error);
      }
    } else {
      console.log('メッセージがありません');
    }
    
    sendDailySummaryEmail(todayMessages, manualInfo, faqInfo);
  } else {
    console.log('本日のメッセージがありません');
  }
}

// FAQシートから最新のFAQを取得
function getFAQsFromSheet(limit = 10) {
  const sheet = getOrCreateFAQSheet();
  const data = sheet.getDataRange().getValues();
  const faqs = [];
  
  // 最新の順に取得（ヘッダーを除く）
  for (let i = data.length - 1; i >= 1 && faqs.length < limit; i--) {
    if (data[i][1] && data[i][2]) { // 質問と回答が存在
      faqs.push({
        question: data[i][1],
        answer: data[i][2],
        category: data[i][3] || 'その他',
        tags: data[i][4] || ''
      });
    }
  }
  
  return faqs;
}

// マニュアルシートから最新のマニュアルを取得
function getLatestManualsFromSheet(limit = 10) {
  const sheet = getOrCreateManualSheet();
  const data = sheet.getDataRange().getValues();
  const manuals = [];
  
  // 最新の順に取得（ヘッダーを除く）
  for (let i = data.length - 1; i >= 1 && manuals.length < limit; i--) {
    if (data[i][1] && data[i][2]) { // カテゴリとタイトルが存在
      manuals.push({
        category: data[i][1],
        title: data[i][2],
        content: data[i][3] || ''
      });
    }
  }
  
  return manuals;
}

// Botがメンバーのチャンネル一覧
// Slackチャンネル一覧を取得 - レート制限対応版
function getSlackChannels(retryCount = 0) {
  const url = 'https://slack.com/api/conversations.list?types=public_channel,private_channel&limit=1000';
  const options = {
    method: 'get',
    headers: { Authorization: 'Bearer ' + SLACK_BOT_TOKEN },
    muteHttpExceptions: true
  };
  
  try {
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    
    // レート制限（429）の場合
    if (responseCode === 429) {
      if (retryCount < 3) {
        const headers = response.getHeaders();
        const retryAfter = parseInt(headers['Retry-After'] || headers['retry-after'] || '60');
        console.log(`レート制限検出（チャンネル一覧）。${retryAfter}秒待機後にリトライ (${retryCount + 1}/3)`);
        Utilities.sleep(Math.min(retryAfter * 1000, 120000));
        return getSlackChannels(retryCount + 1);
      } else {
        throw new Error('レート制限: 最大リトライ回数に到達');
      }
    }
    
    const data = JSON.parse(response.getContentText());
    
    if (!data.ok) {
      if (data.error === 'ratelimited' && retryCount < 3) {
        console.log(`レート制限エラー検出。60秒待機後にリトライ (${retryCount + 1}/3)`);
        Utilities.sleep(60000);
        return getSlackChannels(retryCount + 1);
      }
      throw new Error(`Failed to get channels: ${data.error}`);
    }
    
    // ボットが参加しているチャンネルのみフィルタリング
    return data.channels.filter(channel => channel.is_member);
    
  } catch (error) {
    console.error(`チャンネル一覧取得エラー: ${error.toString()}`);
    
    // エラーの場合もリトライ
    if (retryCount < 2) {
      console.log(`エラー発生。30秒後にリトライ (${retryCount + 1}/3)`);
      Utilities.sleep(30000);
      return getSlackChannels(retryCount + 1);
    }
    
    throw error;
  }
}

// 指定チャンネルの新着メッセージ（oldest以降だけ）- レート制限対応版
function getChannelMessages(channelId, oldest, retryCount = 0) {
  let url = `https://slack.com/api/conversations.history?channel=${channelId}&limit=100`;
  if (oldest && oldest !== '0') {
    url += `&oldest=${oldest}`;
  }
  const options = {
    method: 'get',
    headers: { Authorization: 'Bearer ' + SLACK_BOT_TOKEN },
    muteHttpExceptions: true  // エラーレスポンスを取得できるようにする
  };
  
  try {
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    
    // レート制限（429）の場合
    if (responseCode === 429) {
      if (retryCount < 3) {
        // Retry-Afterヘッダーから待機時間を取得（デフォルト60秒）
        const headers = response.getHeaders();
        const retryAfter = parseInt(headers['Retry-After'] || headers['retry-after'] || '60');
        
        console.log(`レート制限検出 (チャンネル: ${channelId})。${retryAfter}秒待機後にリトライ (${retryCount + 1}/3)`);
        
        // 待機（最大120秒まで）
        const waitTime = Math.min(retryAfter * 1000, 120000);
        Utilities.sleep(waitTime);
        
        // リトライ
        return getChannelMessages(channelId, oldest, retryCount + 1);
      } else {
        console.error(`レート制限: 最大リトライ回数に到達 (チャンネル: ${channelId})`);
        return null;
      }
    }
    
    // その他のエラーコード
    if (responseCode !== 200) {
      console.error(`APIエラー: ステータスコード ${responseCode}`);
      return null;
    }
    
    const data = JSON.parse(response.getContentText());
    
    if (!data.ok) {
      // レート制限エラーの別パターン
      if (data.error === 'ratelimited') {
        if (retryCount < 3) {
          console.log(`レート制限エラー検出。60秒待機後にリトライ (${retryCount + 1}/3)`);
          Utilities.sleep(60000);
          return getChannelMessages(channelId, oldest, retryCount + 1);
        }
      }
      console.error(`APIエラー: ${data.error}`);
      return null;
    }
    
    return data.messages;
    
  } catch (error) {
    console.error(`チャンネルメッセージ取得エラー: ${error.toString()}`);
    
    // ネットワークエラーなどの場合もリトライ
    if (retryCount < 2) {
      console.log(`エラー発生。30秒後にリトライ (${retryCount + 1}/3)`);
      Utilities.sleep(30000);
      return getChannelMessages(channelId, oldest, retryCount + 1);
    }
    
    return null;
  }
}

// Googleドキュメントの初期化（マニュアル専用）
function initializeGoogleDoc() {
  try {
    const doc = DocumentApp.openById(GOOGLE_DOC_ID);
    const body = doc.getBody();
    
    // ドキュメントが空の場合のみタイトルを設定
    if (body.getNumChildren() === 0) {
      const title = body.appendParagraph('📖 業務マニュアル集');
      title.setHeading(DocumentApp.ParagraphHeading.HEADING1);
      title.setFontSize(24);
      title.setBold(true);
      title.setForegroundColor('#1a73e8');
      title.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
      
      const subtitle = body.appendParagraph('Slackコミュニケーションから生成された業務マニュアルとFAQ');
      subtitle.setFontSize(12);
      subtitle.setForegroundColor('#666666');
      subtitle.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
      subtitle.setSpacingAfter(20);
      
      body.appendHorizontalRule();
      
      const toc = body.appendParagraph('📋 目次');
      toc.setHeading(DocumentApp.ParagraphHeading.HEADING2);
      toc.setFontSize(18);
      toc.setForegroundColor('#0066cc');
      toc.setSpacingBefore(20);
      
      body.appendParagraph('※ このドキュメントはAIによって自動生成された業務マニュアルを記録しています。')
        .setFontSize(10)
        .setForegroundColor('#999999')
        .setSpacingAfter(20);
      
      body.appendHorizontalRule();
      
      doc.saveAndClose();
      console.log('Googleドキュメントを初期化しました');
    }
  } catch (error) {
    console.error('Google Doc初期化エラー:', error);
  }
}

// Googleドキュメントに業務マニュアルを追記
function appendManualToGoogleDoc(category, title, content) {
  try {
    const doc = DocumentApp.openById(GOOGLE_DOC_ID);
    const body = doc.getBody();
    
    // セクション番号を取得
    const manualNumber = getNextManualNumber();
    
    // マニュアルヘッダー
    const header = body.appendParagraph(`📢 マニュアル #${manualNumber}`);
    header.setHeading(DocumentApp.ParagraphHeading.HEADING2);
    header.setFontSize(16);
    header.setForegroundColor('#0066cc');
    header.setBold(true);
    header.setSpacingBefore(30);
    
    // メタ情報ボックス
    const metaTable = body.appendTable();
    metaTable.setBorderWidth(0);
    
    const row1 = metaTable.appendTableRow();
    row1.appendTableCell('📅 作成日時').setBackgroundColor('#f8f9fa').setBold(true);
    row1.appendTableCell(new Date().toLocaleString());
    
    const row2 = metaTable.appendTableRow();
    row2.appendTableCell('🏷️ カテゴリ').setBackgroundColor('#f8f9fa').setBold(true);
    row2.appendTableCell(category);
    
    const row3 = metaTable.appendTableRow();
    row3.appendTableCell('🎯 ステータス').setBackgroundColor('#f8f9fa').setBold(true);
    row3.appendTableCell('新規作成').setForegroundColor('#34a853');
    
    // タイトル
    const titlePara = body.appendParagraph(title);
    titlePara.setHeading(DocumentApp.ParagraphHeading.HEADING3);
    titlePara.setFontSize(18);
    titlePara.setBold(true);
    titlePara.setForegroundColor('#1a73e8');
    titlePara.setSpacingBefore(15);
    titlePara.setSpacingAfter(10);
    
    // 内容セクション
    const contentHeader = body.appendParagraph('📝 内容');
    contentHeader.setFontSize(12);
    contentHeader.setBold(true);
    contentHeader.setForegroundColor('#5f6368');
    contentHeader.setSpacingBefore(10);
    
    // 内容をパラグラフごとに処理
    const contentLines = content.split('\n');
    contentLines.forEach(line => {
      if (line.trim()) {
        const para = body.appendParagraph(line);
        para.setFontSize(11);
        para.setLineSpacing(1.5);
        
        // バレットポイントの処理
        if (line.trim().startsWith('-') || line.trim().startsWith('•')) {
          para.setIndentFirstLine(20);
          para.setIndentStart(20);
        }
        
        // 番号付きリストの処理
        if (/^\d+\./.test(line.trim())) {
          para.setIndentFirstLine(20);
          para.setIndentStart(20);
        }
      }
    });
    
    // セクション終了マーカー
    const endMarker = body.appendParagraph('--- マニュアル終了 ---');
    endMarker.setFontSize(10);
    endMarker.setForegroundColor('#999999');
    endMarker.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    endMarker.setSpacingBefore(20);
    endMarker.setSpacingAfter(10);
    
    // 区切り線
    body.appendHorizontalRule();
    
    // 改ページ（オプション）
    if (manualNumber % 5 === 0) {
      body.appendPageBreak();
    }
    
    doc.saveAndClose();
    console.log(`業務マニュアル #${manualNumber}「${title}」をGoogleドキュメントに追加しました`);
  } catch (error) {
    console.error('Google Docマニュアル追記エラー:', error);
  }
}

// 次のマニュアル番号を取得
function getNextManualNumber() {
  try {
    const props = PropertiesService.getScriptProperties();
    let manualCount = parseInt(props.getProperty('MANUAL_COUNT') || '0');
    manualCount++;
    props.setProperty('MANUAL_COUNT', manualCount.toString());
    return manualCount;
  } catch (error) {
    console.error('マニュアル番号取得エラー:', error);
    return 1;
  }
}

// 本文内のメンション<@Uxxxxxx>を実名に置換
function replaceMentionsWithRealNames(text, userCache) {
  if (!text) return '';
  return text.replace(/<@([A-Z0-9]+)>/g, function(match, userId) {
    let name = userCache[userId];
    if (!name) {
      name = getRealName(userId, userCache);
      userCache[userId] = name;
    }
    return '@' + name;
  });
}

// スプレッドシートのログ用シート取得
function getOrCreateLogSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(LOG_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(LOG_SHEET_NAME);
    // ヘッダー行を設定（thread_tsを追加）
    const headers = ['channel_id', 'channel_name', 'timestamp', 'thread_ts', 'user_name', 'message', 'date', 'message_id'];
    sheet.appendRow(headers);
    
    // ヘッダーのスタイル設定
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#4a86e8');
    headerRange.setFontColor('#ffffff');
    headerRange.setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// チャンネルごとに最新tsを保存するシート取得
function getOrCreateLastTsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(LAST_TS_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(LAST_TS_SHEET_NAME);
    const headers = ['channel_id', 'last_ts', 'last_updated'];
    sheet.appendRow(headers);
    
    // ヘッダーのスタイル設定
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#4a86e8');
    headerRange.setFontColor('#ffffff');
    headerRange.setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// 最新ts取得
function getLastFetchedTs(sheet, channelId) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === channelId) {
      return data[i][1];
    }
  }
  return '0';
}

// 最新ts保存
function setLastFetchedTs(sheet, channelId, ts) {
  const data = sheet.getDataRange().getValues();
  const now = new Date();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === channelId) {
      sheet.getRange(i + 1, 2).setValue(ts);
      sheet.getRange(i + 1, 3).setValue(now);
      return;
    }
  }
  sheet.appendRow([channelId, ts, now]);
}

// ユーザー「実名」（real_name）を取得（キャッシュ付き、Bot名等もカバー）
function getRealName(userId, cache, msg = null, retryCount = 0) {
  // BotやWebhookなどはmsg.usernameが入っている場合も
  if (msg && !userId && msg.username) {
    return msg.username;
  }
  if (!userId) return '';
  if (cache[userId]) return cache[userId];
  
  const url = `https://slack.com/api/users.info?user=${userId}`;
  const options = {
    method: 'get',
    headers: { Authorization: 'Bearer ' + SLACK_BOT_TOKEN },
    muteHttpExceptions: true
  };
  
  try {
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    
    // レート制限の場合
    if (responseCode === 429) {
      if (retryCount < 2) {
        // ユーザー情報取得は重要度が低いので短めのリトライ
        console.log(`ユーザー情報取得でレート制限。5秒後にリトライ (${retryCount + 1}/2)`);
        Utilities.sleep(5000);
        return getRealName(userId, cache, msg, retryCount + 1);
      } else {
        console.log(`ユーザー情報取得失敗: ${userId}`);
        cache[userId] = userId; // IDをそのまま使用
        return userId;
      }
    }
    
    const data = JSON.parse(response.getContentText());
    
    if (!data.ok) {
      if (data.error === 'ratelimited' && retryCount < 2) {
        console.log(`ユーザー情報取得でレート制限エラー。5秒後にリトライ`);
        Utilities.sleep(5000);
        return getRealName(userId, cache, msg, retryCount + 1);
      }
      cache[userId] = userId;
      return userId;
    }
    
    const profile = data.user.profile;
    // real_name優先、なければdisplay_name、なければuser.name、なければuserId
    const realName = profile.real_name 
                  || profile.display_name 
                  || data.user.name 
                  || userId;
    cache[userId] = realName;
    return realName;
    
  } catch (error) {
    console.log(`ユーザー情報取得エラー (${userId}):`, error.toString());
    cache[userId] = userId;
    return userId;
  }
}

// メッセージの重複チェック（メッセージIDで確認）
function isMessageExists(sheet, messageId) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][7] === messageId) { // message_idカラムでチェック
      return true;
    }
  }
  return false;
}

// スレッドの返信を取得（まとめて返す）
function getThreadReplies(channelId, threadTs, userCache) {
  try {
    const url = `https://slack.com/api/conversations.replies?channel=${channelId}&ts=${threadTs}&limit=100`;
    const options = {
      method: 'get',
      headers: { Authorization: 'Bearer ' + SLACK_BOT_TOKEN }
    };
    const response = UrlFetchApp.fetch(url, options);
    const data = JSON.parse(response.getContentText());
    
    if (!data.ok || !data.messages) return [];
    
    // 最初のメッセージは親メッセージなのでスキップ
    const replies = [];
    data.messages.slice(1).forEach(reply => {
      if (!reply.text) return;
      
      const realName = getRealName(reply.user, userCache, reply);
      const textWithRealNames = replaceMentionsWithRealNames(reply.text, userCache);
      const time = new Date(Number(reply.ts.split('.')[0]) * 1000).toLocaleTimeString();
      
      replies.push({
        user: realName,
        text: textWithRealNames,
        time: time
      });
    });
    
    return replies;
  } catch (error) {
    console.error('スレッド返信取得エラー:', error);
    return [];
  }
}

// スレッドの会話をフォーマット
function formatThreadConversation(originalUser, originalText, replies) {
  let formattedText = `[スレッド開始]\n`;
  formattedText += `${originalUser}: ${originalText}\n`;
  
  if (replies.length > 0) {
    formattedText += `\n--- 返信 (${replies.length}件) ---\n`;
    replies.forEach(reply => {
      formattedText += `  └ ${reply.user} (${reply.time}): ${reply.text}\n`;
    });
  }
  
  formattedText += `[スレッド終了]`;
  return formattedText;
}

// チャンネル名を取得（キャッシュ付き）
const channelNameCache = {};
function getChannelName(channelId) {
  if (channelNameCache[channelId]) {
    return channelNameCache[channelId];
  }
  
  const sheet = getOrCreateLogSheet();
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === channelId && data[i][1]) {
      channelNameCache[channelId] = data[i][1];
      return data[i][1];
    }
  }
  return channelId; // 名前が見つからない場合はIDを返す
}

// 業務マニュアル用シート取得
function getOrCreateManualSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(MANUAL_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(MANUAL_SHEET_NAME);
    const headers = ['作成日時', 'カテゴリ', 'タイトル', '内容', '元のチャンネル', '関連メッセージ', 'ステータス'];
    sheet.appendRow(headers);
    
    // ヘッダーのスタイル設定
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#34a853');
    headerRange.setFontColor('#ffffff');
    headerRange.setFontWeight('bold');
    sheet.setFrozenRows(1);
    
    // 列幅の調整
    sheet.setColumnWidth(1, 150); // 作成日時
    sheet.setColumnWidth(2, 100); // カテゴリ
    sheet.setColumnWidth(3, 200); // タイトル
    sheet.setColumnWidth(4, 400); // 内容
    sheet.setColumnWidth(5, 120); // 元のチャンネル
    sheet.setColumnWidth(6, 300); // 関連メッセージ
    sheet.setColumnWidth(7, 80);  // ステータス
  }
  return sheet;
}

/**
 * OpenAI API呼び出し共通関数（リトライロジック付き）
 */
function callOpenAIWithRetry(prompt, temperature = 0.3, maxTokens = 2000, retryCount = 0, systemPrompt = null) {
  if (!OPENAI_API_KEY || OPENAI_API_KEY === '***') {
    console.error('OpenAI APIキーが設定されていません');
    return { success: false, error: 'APIキー未設定' };
  }
  
  const url = 'https://api.openai.com/v1/responses';
  const sys = systemPrompt || 'あなたは優秀なアシスタントです。';
  const input = `SYSTEM:\n${sys}\n\nUSER:\n${prompt}`;
  
  const payload = {
    model: 'gpt-5',
    input: input,
    temperature: temperature,
    // Responses API uses max_output_tokens. Keep max_tokens for compatibility.
    max_output_tokens: maxTokens
  };
  
  const options = {
    method: 'post',
    headers: {
      'Authorization': 'Bearer ' + OPENAI_API_KEY,
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  
  try {
    console.log(`OpenAI API呼び出し中... (試行 ${retryCount + 1}/3)`);
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    
    // 成功
    if (responseCode === 200) {
      const data = JSON.parse(response.getContentText());
      const content = extractTextFromOpenAIResponse(data);
      if (content) {
        return { success: true, content };
      }
    }
    
    // レート制限
    if (responseCode === 429) {
      if (retryCount < 2) {
        console.log('OpenAI APIレート制限。30秒後にリトライ');
        Utilities.sleep(30000);
        return callOpenAIWithRetry(prompt, temperature, maxTokens, retryCount + 1, systemPrompt);
      }
    }
    
    // その他のエラー
    console.error(`OpenAI APIエラー: ステータスコード ${responseCode}`);
    console.error(`レスポンス: ${response.getContentText()}`);
    
    // リトライ可能なエラーの場合
    if (responseCode >= 500 && retryCount < 2) {
      console.log(`サーバーエラー。10秒後にリトライ`);
      Utilities.sleep(10000);
      return callOpenAIWithRetry(prompt, temperature, maxTokens, retryCount + 1, systemPrompt);
    }
    
    return { success: false, error: `APIエラー: ${responseCode}` };
    
  } catch (error) {
    console.error('OpenAI API呼び出しエラー:', error.toString());
    
    // ネットワークエラーの場合のリトライ
    if (retryCount < 2) {
      // Address unavailableエラーの場合は長めに待つ
      if (error.toString().includes('Address unavailable')) {
        console.log('接続エラー。60秒後にリトライ');
        Utilities.sleep(60000);
      } else {
        console.log('エラー発生。20秒後にリトライ');
        Utilities.sleep(20000);
      }
      return callOpenAIWithRetry(prompt, temperature, maxTokens, retryCount + 1, systemPrompt);
    }
    
    return { success: false, error: error.toString() };
  }
}

/**
 * OpenAI Responses APIのレスポンスからテキストを抽出（後方互換付き）
 */
function extractTextFromOpenAIResponse(data) {
  try {
    if (!data) return '';
    // Responses API convenience field
    if (typeof data.output_text === 'string' && data.output_text.trim()) {
      return data.output_text;
    }
    // Responses API structured output
    if (Array.isArray(data.output)) {
      // output -> array of content blocks; try to join text parts
      const parts = [];
      data.output.forEach(block => {
        if (block && Array.isArray(block.content)) {
          block.content.forEach(c => {
            if (c && typeof c.text === 'string') parts.push(c.text);
            if (c && c.type === 'text' && c.text && c.text.value) parts.push(c.text.value);
          });
        }
      });
      const joined = parts.join('\n').trim();
      if (joined) return joined;
    }
    // Backward compatibility: Chat Completions
    if (data.choices && data.choices[0]) {
      const choice = data.choices[0];
      if (choice.message && typeof choice.message.content === 'string') {
        return choice.message.content;
      }
      if (typeof choice.text === 'string') return choice.text;
    }
  } catch (e) {
    console.error('レスポンス解析エラー:', e);
  }
  return '';
}

/**
 * メッセージをトピック別に分類する（改良版）
 */
function classifyMessagesByTopic(messages) {
  const topics = [];
  let currentTopic = [];
  let lastThreadTs = null;
  
  for (const msg of messages) {
    const threadTs = msg.thread_ts || msg.ts;
    
    // 新しいスレッドまたは時間差が大きい場合は新トピック
    if (lastThreadTs !== threadTs || 
        (lastThreadTs && Math.abs(parseFloat(msg.ts) - parseFloat(lastThreadTs)) > 3600)) {
      
      if (currentTopic.length > 0) {
        topics.push([...currentTopic]);
        currentTopic = [];
      }
    }
    
    currentTopic.push(msg);
    lastThreadTs = threadTs;
  }
  
  // 最後のトピックを追加
  if (currentTopic.length > 0) {
    topics.push(currentTopic);
  }
  
  console.log(`${messages.length}件のメッセージを${topics.length}個のトピックに分類`);
  return topics;
}

/**
 * メッセージのコンテキストを分析（改良版）
 */
function analyzeMessageContext(messages) {
  const keywords = new Set();
  const participants = new Set();
  let hasQuestion = false;
  let hasDecision = false;
  let hasInstruction = false;
  let hasTroubleshooting = false;
  
  for (const msg of messages) {
    // 参加者を記録
    if (msg.user) participants.add(msg.user);
    
    // メッセージの特徴を分析
    const text = (msg.text || '').toLowerCase();
    
    // 質問パターン
    if (text.match(/[?？]|どう|なぜ|いつ|どこ|誰|何|how|what|when|where|why|who/)) {
      hasQuestion = true;
    }
    
    // 決定パターン
    if (text.match(/決定|決まり|確定|承認|approved|decided|confirmed/)) {
      hasDecision = true;
    }
    
    // 指示パターン
    if (text.match(/してください|お願い|やって|実行|please|execute|run/)) {
      hasInstruction = true;
    }
    
    // トラブルシューティングパターン
    if (text.match(/エラー|失敗|問題|修正|解決|error|failed|issue|fix|solve/)) {
      hasTroubleshooting = true;
    }
    
    // キーワード抽出（簡易版）
    const words = text.split(/[\s　,、。.!！?？]+/);
    words.forEach(word => {
      if (word.length > 3 && !word.match(/^(です|ます|した|して|これ|それ|あれ|this|that|have|will|would)$/)) {
        keywords.add(word);
      }
    });
  }
  
  return {
    messageCount: messages.length,
    participantCount: participants.size,
    participants: Array.from(participants),
    keywords: Array.from(keywords).slice(0, 10),
    hasQuestion,
    hasDecision,
    hasInstruction,
    hasTroubleshooting,
    estimatedType: determineDocumentType(hasQuestion, hasDecision, hasInstruction, hasTroubleshooting, messages)
  };
}

/**
 * ドキュメントタイプを判定（改良版）
 */
function determineDocumentType(hasQuestion, hasDecision, hasInstruction, hasTroubleshooting, messages) {
  if (hasTroubleshooting) return 'TROUBLESHOOTING';
  if (hasQuestion && messages.length > 1) return 'FAQ';
  if (hasDecision) return 'DECISION';
  if (hasInstruction) return 'PROCEDURE';
  return 'INFORMATION';
}

// OpenAI APIを使用して業務マニュアルを生成（既存版）
function generateBusinessManual(messages, retryCount = 0) {
  console.log(`=== 業務マニュアル生成開始: ${messages.length}件のメッセージ (リトライ: ${retryCount}) ===`);
  
  if (!OPENAI_API_KEY || OPENAI_API_KEY === '***') {
    console.error('OpenAI APIキーが設定されていません');
    return null;
  }
  
  // メッセージ数の制限を撤廃（以前は2件以上が条件だった）
  if (!messages || messages.length === 0) {
    console.log('メッセージが空です');
    return null;
  }
  
  const sheet = getOrCreateManualSheet();
  
  // スレッドをまとめて処理
  const conversationText = formatMessagesForAI(messages);
  
  // テキストの最小文字数制限を緩和（50→20文字）
  if (!conversationText || conversationText.length < 20) {
    console.log(`会話テキストが短すぎます: ${conversationText?.length || 0}文字`);
    console.log(`テキスト内容: ${conversationText}`);
    return null;
  }
  
  console.log(`会話テキスト長: ${conversationText.length}文字`);
  console.log(`テキストサンプル: ${conversationText.substring(0, 200)}...`);
  
  const prompt = `以下のSlackでのやり取りから、業務マニュアルやFAQとして有用な情報を抽出してください。
メッセージが少なくても、推測や一般化を行って、必ず最低1つ以上のマニュアル項目を生成してください。

抽出する際のポイント：
1. 業務プロセスやワークフローに関する議論
2. 意思決定事項や承認プロセス
3. ツールやシステムの使用方法
4. 問題解決やエラー対応の手順
5. チーム内のルールやガイドライン
6. プロジェクトの進め方や管理方法
7. コミュニケーションのプロトコル
8. セキュリティやコンプライアンス関連
9. 日常的な質問や情報共有
10. 会議や打ち合わせの内容

重要：
- 会話が断片的でも、文脈から推測して有用な情報を抽出してください
- 簡単な挨拶や雑談からも、チームのコミュニケーション文化として文書化してください
- 必ず最低1つ以上の項目を生成してください

出力形式（必ず以下の形式を厳守してください）：
カテゴリ: [業務プロセス/システム操作/トラブルシューティング/意思決定/ベストプラクティス/コミュニケーション/その他]
タイトル: [具体的で明確なタイトル]
内容: [詳細な内容。可能な限り以下の要素を含める]
  - 目的・背景
  - 前提条件・必要なリソース
  - 具体的な手順（ステップバイステップ）
  - 注意事項・リスク
  - 期待される結果・成果物
  - 関連資料・参考情報
  - 担当者・責任範囲
  - フォローアップ・レビュープロセス
---

やり取り内容：
${conversationText}`;
  
  try {
    const systemPrompt = `あなたは企業のナレッジマネジメント専門家であり、ISO 9001、PMBOK、ITIL、Six Sigmaなどの国際標準に精通しています。
コミュニケーションログから以下を抽出・体系化してください：
1. 暗黙知を形式知に変換
2. 業務プロセスをBPMNに準拠した形で整理
3. RACIマトリクスに基づく役割分担の明確化
4. KPI/SLAの設定可能な指標の特定
5. リスク管理と緩和策の提案
6. コンプライアンスと監査証跡の確保
7. 継続的改善(PDCA/DMAIC)の観点

出力は実務で即座に使用可能なレベルの詳細度と専門性を保ち、必要に応じて業界標準やベストプラクティスを参照してください。`;
    
    console.log('OpenAI APIを呼び出し中...');
    const openAIResult = callOpenAIWithRetry(prompt, 0.3, 4096, 0, systemPrompt);
    
    if (!openAIResult || !openAIResult.success) {
      console.error('API呼び出し失敗:', openAIResult?.error);
      return null;
    }
    
    const manualContent = openAIResult.content;
    console.log(`生成されたコンテンツ: ${manualContent.substring(0, 200)}...`);
    
    // マニュアル内容を複数の項目に分割
    let items = manualContent.split('---').filter(item => item.trim());
    console.log(`分割された項目数: ${items.length}`);
    
    // 分割できなかった場合の処理を改善
    if (items.length === 0) {
      console.log('「---」による分割ができませんでした。全体を１つの項目として処理します');
      // カテゴリ、タイトル、内容を含んでいるか確認
      if (manualContent.includes('カテゴリ:') && manualContent.includes('タイトル:')) {
        items = [manualContent];
      } else {
        // フォーマットが不正な場合、デフォルト値で項目を作成
        console.log('期待されるフォーマットではありません。デフォルト値で項目を作成します');
        const defaultItem = `カテゴリ: その他\nタイトル: Slack会話からの抽出情報\n内容: ${manualContent}`;
        items = [defaultItem];
      }
    }
    
    const channels = [...new Set(messages.map(m => m.channel))].join(', ');
    const timestamp = new Date().toLocaleString();
    let generatedCount = 0;
    
    items.forEach((item, index) => {
      console.log(`項目 ${index + 1}/${items.length} を処理中...`);
      const lines = item.trim().split('\n');
      let category = '';
      let title = '';
      let content = [];
      
      lines.forEach(line => {
        const trimmedLine = line.trim();
        
        // カテゴリの検出（大文字小文字、スペースに柔軟に対応）
        if (trimmedLine.match(/^カテゴリ[:：]/i)) {
          category = trimmedLine.replace(/^カテゴリ[:：]/i, '').trim();
          // デフォルトカテゴリの設定
          if (!category) category = 'その他';
        } 
        // タイトルの検出
        else if (trimmedLine.match(/^タイトル[:：]/i)) {
          title = trimmedLine.replace(/^タイトル[:：]/i, '').trim();
          // デフォルトタイトルの設定
          if (!title) title = '無題のマニュアル';
        } 
        // 内容の検出
        else if (trimmedLine.match(/^内容[:：]/i)) {
          const contentLine = trimmedLine.replace(/^内容[:：]/i, '').trim();
          if (contentLine) content.push(contentLine);
        } 
        // その他の行は内容として追加
        else if (trimmedLine && (category || title)) {
          content.push(trimmedLine);
        }
      });
      
      // デフォルト値の設定
      if (!category) category = 'その他';
      if (!title) title = `マニュアル項目 ${index + 1}`;
      if (content.length === 0 && item.trim().length > 20) {
        // 内容が空でも、項目自体にテキストがあれば内容として使用
        content.push(item.trim());
      }
      
      // 最後のエントリを保存（条件を緩和）
      if (category || title || content.length > 0) {
        // デフォルト値の再確認
        if (!category) category = 'その他';
        if (!title) title = `マニュアル項目 ${index + 1}`;
        if (content.length === 0) content.push('内容が抽出されませんでした');
        
        console.log(`保存: カテゴリ="${category}", タイトル="${title}"`);
        console.log(`内容プレビュー: ${content.join(' ').substring(0, 100)}...`);
        
        sheet.appendRow([
          timestamp,
          category,
          title,
          content.join('\n'),
          channels,
          conversationText.substring(0, 500), // 元メッセージの最初の500文字
          '新規' // ステータス
        ]);
        
        // Googleドキュメントにも追記
        try {
          appendManualToGoogleDoc(category, title, content.join('\n'));
        } catch (docError) {
          console.error('Google Docへの追記エラー:', docError);
        }
        
        generatedCount++;
      } else {
        console.log(`スキップ: カテゴリ="${category}", タイトル="${title}", コンテンツ数=${content.length}`);
      }
    });
    
    console.log(`業務マニュアル ${generatedCount}/${items.length} 件を生成しました`);
    
    // 生成されたマニュアル情報を返す
    if (generatedCount > 0) {
      return {
        count: generatedCount,
        manuals: items.slice(0, generatedCount).map(item => {
          const lines = item.trim().split('\n');
          let category = '', title = '';
          lines.forEach(line => {
            if (line.startsWith('カテゴリ:')) category = line.replace('カテゴリ:', '').trim();
            if (line.startsWith('タイトル:')) title = line.replace('タイトル:', '').trim();
          });
          return { category, title };
        })
      };
    } else {
      console.log('マニュアルが生成されませんでした');
      return null;
    }
  } catch (error) {
    console.error('業務マニュアル生成エラー:', error.toString());
    console.error('スタックトレース:', error.stack);
    
    // エラーの種類を判別
    if (error.toString().includes('429') || error.toString().includes('rate limit')) {
      console.error('APIレート制限に達しました。時間をおいて再試行してください');
      
      // レート制限の場合、少し待ってリトライ（最大3回）
      if (retryCount < 3) {
        console.log(`${5 * (retryCount + 1)}秒後にリトライします...`);
        Utilities.sleep(5000 * (retryCount + 1)); // 5秒、10秒、15秒と増加
        return generateBusinessManual(messages, retryCount + 1);
      }
    } else if (error.toString().includes('401') || error.toString().includes('Unauthorized')) {
      console.error('APIキーが無効です。OPENAI_API_KEYを確認してください');
    } else if (error.toString().includes('timeout')) {
      console.error('APIリクエストがタイムアウトしました');
      
      // タイムアウトの場合もリトライ
      if (retryCount < 2) {
        console.log('10秒後にリトライします...');
        Utilities.sleep(10000);
        return generateBusinessManual(messages, retryCount + 1);
      }
    }
    
    return null;
  }
}

// 日次要約をメール送信（業務マニュアル・FAQ情報を含む）
function sendDailySummaryEmail(messages, manualInfo = null, faqInfo = null) {
  if (!NOTIFICATION_EMAIL || NOTIFICATION_EMAIL === '***') {
    console.log('送信先メールアドレスが設定されていません');
    return;
  }
  
  // スレッドを考慮してメッセージをグループ化
  const { channelMessages, threadCount } = groupMessagesByThread(messages);
  
  // メッセージ数のチェック
  if (messages.length === 0) {
    console.log('送信するメッセージがありません');
    return;
  }
  
  // 要約を生成（OpenAI APIを使用する場合）
  let summary = '';
  
  if (OPENAI_API_KEY && OPENAI_API_KEY !== '***' && messages.length >= 3) { // 最低3件以上のメッセージがある場合のみ要約生成
    try {
      const conversationText = formatMessagesForAI(messages);
      
      // メッセージが少ない場合は簡潔なプロンプト
      const prompt = messages.length < 10 
        ? `以下のSlackメッセージを簡潔に要約してください。\n\n${conversationText}`
        : `以下のSlackでのやり取りについて、エグゼクティブサマリーを作成してください。

以下の構成で詳細にまとめてください：

1. 【エグゼクティブサマリー】
   - 本日のハイライト
   - 重要な意思決定事項
   - クリティカルな課題

2. 【意思決定事項とアクションアイテム】
   - 決定事項（意思決定者、内容、理由）
   - アクションアイテム（担当者、期限、優先度）
   - フォローアップ事項

3. 【プロジェクト・タスクの進捗】
   - 完了したタスク
   - 進行中のタスクと進捗率
   - ブロッカーとリスク
   - 次のマイルストーン

4. 【チームコミュニケーション分析】
   - コミュニケーションパターン
   - コラボレーションの質
   - 改善点と提案

5. 【ナレッジキャプチャー】
   - 新しく得られた知見
   - ベストプラクティス
   - 教訓と学び

6. 【KPIとメトリクス】
   - パフォーマンス指標
   - 目標達成状況
   - 改善が必要な領域

7. 【推奨事項】
   - 短期的な推奨
   - 中長期的な提案
   - プロセス改善の機会

やり取り内容：
${conversationText}`;
      
      // OpenAI API呼び出し（リトライロジック付き）
      const systemPrompt = `あなたは経営コンサルタントであり、ビジネスインテリジェンスとデータ分析の専門家です。
コミュニケーションログから戦略的インサイトを抽出し、以下の観点で分析してください：
- ビジネスインパクトとROI
- ステークホルダーへの影響
- リスクと機会のSWOT分析
- 競争優位性と市場ポジショニング
- オペレーショナルエクセレンス
- チェンジマネジメントの観点
定量的なデータと定性的な洞察をバランスよく含め、実行可能な提言を提供してください。`;
      
      const openAIResult = callOpenAIWithRetry(prompt, 0.4, 4096, 0, systemPrompt);
      
      if (openAIResult && openAIResult.success) {
        const content = openAIResult.content;
        
        // 空のレスポンスやエラーメッセージをチェック
        if (content && !content.includes('[空白]') && !content.includes('具体的なやり取り内容が提供されれば')) {
          summary = content.replace(/\n/g, '<br>');
        } else {
          // フォールバックメッセージ
          summary = generateSimpleSummary(messages);
        }
      } else {
        // API呼び出し失敗時のフォールバック
        console.log('OpenAI API呼び出し失敗。簡易要約を使用');
        summary = generateSimpleSummary(messages);
      }
    } catch (error) {
      console.error('要約生成エラー:', error);
      summary = generateSimpleSummary(messages);
    }
  } else if (messages.length > 0) {
    summary = generateSimpleSummary(messages);
  }
  
  // HTMLメール本文を作成
  let htmlBody = `
    <!DOCTYPE html>
    <html>
    <head>
      <style>
        body { font-family: 'Helvetica Neue', Arial, sans-serif; line-height: 1.6; color: #333; }
        .container { max-width: 800px; margin: 0 auto; padding: 20px; }
        .header { background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; padding: 30px; border-radius: 10px; margin-bottom: 30px; }
        .header h1 { margin: 0; font-size: 28px; }
        .date { opacity: 0.9; margin-top: 10px; }
        .section { background: #f8f9fa; border-left: 4px solid #667eea; padding: 20px; margin: 20px 0; border-radius: 5px; }
        .section h2 { color: #667eea; margin-top: 0; }
        .channel-box { background: white; padding: 15px; margin: 15px 0; border-radius: 5px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
        .channel-name { font-weight: bold; color: #667eea; margin-bottom: 10px; font-size: 16px; }
        .message { padding: 5px 0; border-bottom: 1px solid #eee; }
        .message:last-child { border-bottom: none; }
        .user { font-weight: bold; color: #555; }
        .manual-section { background: #e8f5e9; border-left: 4px solid #4caf50; }
        .manual-item { background: white; padding: 10px; margin: 10px 0; border-radius: 5px; }
        .footer { text-align: center; color: #999; font-size: 12px; margin-top: 40px; padding-top: 20px; border-top: 1px solid #eee; }
        .stats { display: flex; justify-content: space-around; margin: 20px 0; }
        .stat-box { text-align: center; padding: 15px; background: white; border-radius: 5px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
        .stat-number { font-size: 24px; font-weight: bold; color: #667eea; }
        .stat-label { color: #999; font-size: 12px; margin-top: 5px; }
      </style>
    </head>
    <body>
      <div class="container">
        <div class="header">
          <h1>📨 Slack 日次レポート</h1>
          <div class="date">📅 ${new Date().toLocaleDateString('ja-JP', { year: 'numeric', month: 'long', day: 'numeric', weekday: 'long' })}</div>
        </div>
        
        <div class="stats">
          <div class="stat-box">
            <div class="stat-number">${messages.length}</div>
            <div class="stat-label">総メッセージ数</div>
          </div>
          <div class="stat-box">
            <div class="stat-number">${Object.keys(channelMessages).length}</div>
            <div class="stat-label">アクティブチャンネル</div>
          </div>
          ${manualInfo ? `
          <div class="stat-box">
            <div class="stat-number">${manualInfo.count}</div>
            <div class="stat-label">生成マニュアル</div>
          </div>
          ` : ''}
        </div>
        
        <div class="section">
          <h2>📝 今日の要約</h2>
          <div>${summary}</div>
        </div>
        
        ${manualInfo && manualInfo.count > 0 ? `
        <div class="section manual-section">
          <h2>📖 生成された業務マニュアル</h2>
          ${manualInfo.manuals.map(m => `
            <div class="manual-item">
              <strong>カテゴリ:</strong> ${m.category}<br>
              <strong>タイトル:</strong> ${m.title}
            </div>
          `).join('')}
          <div style="margin-top: 15px;">
            <a href="${getSpreadsheetUrl()}#gid=${getSheetId(MANUAL_SHEET_NAME)}" style="color: #4caf50; text-decoration: none;">
              🔗 スプレッドシートで詳細を確認
            </a>
          </div>
        </div>
        ` : ''}
        
        ${faqInfo && faqInfo.count > 0 ? `
        <div class="section faq-section">
          <h2>❓ 生成されたFAQ</h2>
          <div style="background: #f8f9fa; padding: 15px; border-radius: 8px;">
            ${faqInfo.faqs.slice(0, 5).map(faq => `
              <div style="margin-bottom: 15px; padding-bottom: 15px; border-bottom: 1px solid #e0e0e0;">
                <div style="color: #1a73e8; font-weight: bold; margin: 5px 0;">Q: ${faq.question}</div>
                <div style="color: #5f6368; margin: 5px 0 5px 20px;">A: ${faq.answer}</div>
                <div style="font-size: 11px; color: #999; margin: 5px 0 0 20px;">🎯 ${faq.category} ${faq.tags ? '| 🏿️ ' + faq.tags : ''}</div>
              </div>
            `).join('')}
            ${faqInfo.count > 5 ? `
              <div style="text-align: center; color: #666; margin-top: 10px;">
                ...他${faqInfo.count - 5}件のFAQがあります
              </div>
            ` : ''}
          </div>
          <div style="margin-top: 15px;">
            <a href="${getSpreadsheetUrl()}#gid=${getSheetId(FAQ_SHEET_NAME)}" style="color: #ea4335; text-decoration: none;">
              🔗 すべてのFAQをスプレッドシートで確認
            </a>
          </div>
        </div>
        ` : ''}
        
        <div class="section">
          <h2>💬 チャンネル別詳細</h2>
          ${Object.keys(channelMessages).map(channel => `
            <div class="channel-box">
              <div class="channel-name">#${channel} (${channelMessages[channel].length}件)</div>
              ${channelMessages[channel].slice(0, 3).map(msg => `
                <div class="message">
                  <span class="user">${msg.user}:</span> ${msg.text.substring(0, 100)}${msg.text.length > 100 ? '...' : ''}
                </div>
              `).join('')}
              ${channelMessages[channel].length > 3 ? `
                <div style="color: #999; font-size: 12px; margin-top: 10px;">
                  ...他 ${channelMessages[channel].length - 3} 件のメッセージ
                </div>
              ` : ''}
            </div>
          `).join('')}
        </div>
        
        <div class="footer">
          <p>🤖 このレポートはSlackログ収集ツールによって自動生成されました</p>
          <p>
            <a href="${getSpreadsheetUrl()}" style="color: #667eea;">📊 スプレッドシート</a> | 
            <a href="https://docs.google.com/document/d/${GOOGLE_DOC_ID}" style="color: #667eea;">📄 Googleドキュメント</a>
          </p>
        </div>
      </div>
    </body>
    </html>
  `;
  
  // プレーンテキスト版も作成
  let plainBody = `Slack日次レポート - ${new Date().toLocaleDateString()}\n\n`;
  plainBody += `総メッセージ数: ${messages.length}\n`;
  plainBody += `アクティブチャンネル: ${Object.keys(channelMessages).length}\n`;
  if (manualInfo) {
    plainBody += `生成マニュアル: ${manualInfo.count}\n`;
  }
  plainBody += `\n詳細はHTML版をご覧ください。`;
  
  // メールを送信
  try {
    MailApp.sendEmail({
      to: NOTIFICATION_EMAIL,
      subject: `📨 Slack日次レポート - ${new Date().toLocaleDateString()} ${manualInfo && manualInfo.count > 0 ? `[マニュアル${manualInfo.count}件生成]` : ''}`,
      body: plainBody,
      htmlBody: htmlBody
    });
    console.log('日次要約メールを送信しました');
  } catch (error) {
    console.error('メール送信エラー:', error);
  }
}

// スプレッドシートのURLを取得
function getSpreadsheetUrl() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getUrl();
}

// シートIDを取得
function getSheetId(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  return sheet ? sheet.getSheetId() : '';
}

// シンプルな要約を生成（フォールバック用）
function generateSimpleSummary(messages) {
  const channelMessages = {};
  messages.forEach(msg => {
    if (!channelMessages[msg.channel]) {
      channelMessages[msg.channel] = [];
    }
    channelMessages[msg.channel].push(msg);
  });
  
  let summary = '<h3>📊 Slack活動レポート</h3>';
  summary += '<div style="background: #f8f9fa; padding: 15px; border-radius: 5px; margin: 10px 0;">';
  summary += `<p><strong>📨 総メッセージ数:</strong> ${messages.length}件</p>`;
  summary += `<p><strong>💬 アクティブチャンネル:</strong> ${Object.keys(channelMessages).length}チャンネル</p>`;
  
  // チャンネル別の内訳
  summary += '<h4>チャンネル別活動:</h4><ul>';
  Object.keys(channelMessages).forEach(channel => {
    const msgs = channelMessages[channel];
    const users = [...new Set(msgs.map(m => m.user))];
    summary += `<li><strong>#${channel}</strong>: ${msgs.length}件のメッセージ (参加者: ${users.length}名)</li>`;
  });
  summary += '</ul>';
  
  // アクティブユーザー
  const userMessages = {};
  messages.forEach(msg => {
    if (!userMessages[msg.user]) userMessages[msg.user] = 0;
    userMessages[msg.user]++;
  });
  
  const topUsers = Object.entries(userMessages)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 5);
  
  if (topUsers.length > 0) {
    summary += '<h4>🏆 アクティブユーザー TOP5:</h4><ol>';
    topUsers.forEach(([user, count]) => {
      summary += `<li>${user}: ${count}件</li>`;
    });
    summary += '</ol>';
  }
  
  summary += '</div>';
  
  // スレッド情報
  const threadMessages = messages.filter(m => m.threadTs && m.threadTs !== m.ts);
  if (threadMessages.length > 0) {
    summary += `<p><strong>🧵 スレッド返信:</strong> ${threadMessages.length}件</p>`;
  }
  
  return summary;
}

// メッセージをAI用にフォーマット（スレッドを考慮）
function formatMessagesForAI(messages) {
  // スレッドごとにグループ化
  const threads = {};
  const standaloneMessages = [];
  
  messages.forEach(msg => {
    if (msg.threadTs) {
      const threadKey = `${msg.channel}_${msg.threadTs}`;
      if (!threads[threadKey]) {
        threads[threadKey] = {
          channel: msg.channel,
          messages: []
        };
      }
      threads[threadKey].messages.push(msg);
    } else {
      standaloneMessages.push(msg);
    }
  });
  
  let formattedText = '';
  
  // スレッドメッセージをフォーマット
  Object.values(threads).forEach(thread => {
    formattedText += `\n[スレッド会話 - #${thread.channel}]\n`;
    thread.messages.forEach(msg => {
      // スレッド内のメッセージを解析
      if (msg.text.includes('[スレッド開始]')) {
        // 既にフォーマット済みのスレッド
        formattedText += msg.text + '\n';
      } else {
        formattedText += `${msg.user}: ${msg.text}\n`;
      }
    });
    formattedText += '\n';
  });
  
  // スタンドアロンメッセージをフォーマット
  if (standaloneMessages.length > 0) {
    formattedText += '\n[独立メッセージ]\n';
    standaloneMessages.forEach(msg => {
      formattedText += `[${msg.channel}] ${msg.user}: ${msg.text}\n`;
    });
  }
  
  return formattedText;
}

// スレッドごとにメッセージをグループ化
function groupMessagesByThread(messages) {
  const channelMessages = {};
  const threads = new Set();
  
  messages.forEach(msg => {
    if (!channelMessages[msg.channel]) {
      channelMessages[msg.channel] = [];
    }
    channelMessages[msg.channel].push(msg);
    
    if (msg.threadTs) {
      threads.add(`${msg.channel}_${msg.threadTs}`);
    }
  });
  
  return {
    channelMessages: channelMessages,
    threadCount: threads.size
  };
}

// 手動で業務マニュアルを生成（デバッグ用）
function manualGenerateBusinessManual() {
  const ui = SpreadsheetApp.getUi();
  console.log('=== 手動業務マニュアル生成開始 ===');
  
  try {
    // OpenAI APIキーのチェック
    if (!OPENAI_API_KEY || OPENAI_API_KEY === '***') {
      ui.alert('エラー', 'OpenAI APIキーが設定されていません。\nconst OPENAI_API_KEY を設定してください。', ui.ButtonSet.OK);
      return;
    }
    
    const sheet = getOrCreateLogSheet();
    const data = sheet.getDataRange().getValues();
    
    // ヘッダー行を除外してすべてのメッセージを取得
    if (data.length <= 1) {
      ui.alert('データなし', 'slack_logシートにデータがありません。\nまず「Slackログ取得」を実行してください。', ui.ButtonSet.OK);
      return;
    }
    
    // 最新のメッセージから取得（最大100件）
    const maxMessages = Math.min(data.length - 1, 100);
    const startRow = Math.max(1, data.length - maxMessages);
    const allMessages = [];
    
    for (let i = startRow; i < data.length; i++) {
      if (data[i][1] && data[i][4] && data[i][5]) { // channel, user, textが存在する場合
        allMessages.push({
          channel: data[i][1],
          user: data[i][4],
          text: data[i][5],
          ts: data[i][2],
          threadTs: data[i][3]
        });
      }
    }
    
    if (allMessages.length === 0) {
      ui.alert('データなし', '処理可能なメッセージがありません。', ui.ButtonSet.OK);
      return;
    }
    
    // プログレス表示
    const progressHtml = HtmlService.createHtmlOutput(
      `<p>業務マニュアルを生成中です...</p>
       <p>${allMessages.length}件のメッセージを処理しています。</p>
       <p>この処理には30秒〜1分程度かかる場合があります。</p>`
    ).setWidth(400).setHeight(150);
    ui.showModalDialog(progressHtml, '処理中');
    
    console.log(`${allMessages.length}件のメッセージから業務マニュアルを生成します`);
    const result = generateBusinessManual(allMessages);
    
    // 結果を表示
    if (result && result.count > 0) {
      const successMessage = `業務マニュアルを${result.count}件生成しました！\n\n` +
        result.manuals.map((m, i) => `${i+1}. [${m.category}] ${m.title}`).join('\n');
      ui.alert('成功', successMessage, ui.ButtonSet.OK);
    } else {
      // 詳細なエラー情報を取得
      const logs = console.getLog ? console.getLog() : 'ログ取得不可';
      ui.alert('生成失敗', 
        'マニュアルの生成に失敗しました。\n\n' +
        '考えられる原因:\n' +
        '1. メッセージ内容が業務マニュアル化に適していない\n' +
        '2. OpenAI APIの応答が期待と異なる\n' +
        '3. ネットワークエラー\n\n' +
        '詳細はApps Scriptエディタのログを確認してください。',
        ui.ButtonSet.OK);
    }
  } catch (error) {
    console.error('マニュアル生成エラー:', error.toString());
    console.error('スタックトレース:', error.stack);
    ui.alert('エラー', `エラーが発生しました:\n${error.toString()}`, ui.ButtonSet.OK);
  }
}

// 期間を指定して業務マニュアルを生成
function generateManualForPeriod(startDate, endDate) {
  const sheet = getOrCreateLogSheet();
  const data = sheet.getDataRange().getValues();
  
  if (!startDate) startDate = new Date(Date.now() - 7 * 24 * 60 * 60 * 1000); // デフォルト: 過去7日間
  if (!endDate) endDate = new Date();
  
  startDate.setHours(0, 0, 0, 0);
  endDate.setHours(23, 59, 59, 999);
  
  const periodMessages = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][6] instanceof Date) { // dateカラムが存在する場合
      const msgDate = data[i][6];
      if (msgDate >= startDate && msgDate <= endDate) {
        periodMessages.push({
          channel: data[i][1],
          user: data[i][4],
          text: data[i][5],
          ts: data[i][2],
          threadTs: data[i][3]
        });
      }
    }
  }
  
  if (periodMessages.length > 0) {
    console.log(`${startDate.toLocaleDateString()}から${endDate.toLocaleDateString()}までの${periodMessages.length}件のメッセージから業務マニュアルを生成します`);
    generateBusinessManual(periodMessages);
  } else {
    console.log('指定期間内にメッセージがありません');
  }
}

// チャンネルを指定して業務マニュアルを生成
function generateManualForChannel(channelName) {
  const sheet = getOrCreateLogSheet();
  const data = sheet.getDataRange().getValues();
  
  const channelMessages = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === channelName) {
      channelMessages.push({
        channel: data[i][1],
        user: data[i][4],
        text: data[i][5],
        ts: data[i][2],
        threadTs: data[i][3]
      });
    }
  }
  
  if (channelMessages.length > 0) {
    console.log(`チャンネル「${channelName}」の${channelMessages.length}件のメッセージから業務マニュアルを生成します`);
    generateBusinessManual(channelMessages);
  } else {
    console.log(`チャンネル「${channelName}」のメッセージが見つかりません`);
  }
}

// 手動で日次要約メールを送信（デバッグ用）
function manualSendDailySummary() {
  const sheet = getOrCreateLogSheet();
  const data = sheet.getDataRange().getValues();
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  
  const todayMessages = [];
  for (let i = 1; i < data.length; i++) {
    const ts = Number(data[i][2].split('.')[0]) * 1000;
    const msgDate = new Date(ts);
    msgDate.setHours(0, 0, 0, 0);
    
    if (msgDate.getTime() === today.getTime()) {
      todayMessages.push({
        channel: data[i][1],
        user: data[i][4],
        text: data[i][5],
        ts: data[i][2],
        threadTs: data[i][3]
      });
    }
  }
  
  if (todayMessages.length > 0) {
    sendDailySummaryEmail(todayMessages);
  }
}

/**
 * 初期設定チェック関数
 * 設定が正しく行われているか確認します
 */
function checkSetup() {
  console.log('=== セットアップチェック開始 ===');
  
  // Slack Token チェック
  if (!SLACK_BOT_TOKEN || SLACK_BOT_TOKEN === '***') {
    console.error('❌ Slack Bot Token が設定されていません');
    console.log('→ https://api.slack.com/apps でアプリを作成し、Bot User OAuth Token を設定してください');
  } else {
    console.log('✅ Slack Bot Token: 設定済み');
    // トークンの有効性をテスト
    try {
      const testUrl = 'https://slack.com/api/auth.test';
      const options = {
        method: 'get',
        headers: { Authorization: 'Bearer ' + SLACK_BOT_TOKEN }
      };
      const response = UrlFetchApp.fetch(testUrl, options);
      const data = JSON.parse(response.getContentText());
      if (data.ok) {
        console.log(`✅ Slack接続成功: Team=${data.team}, User=${data.user}`);
      } else {
        console.error('❌ Slack Token が無効です:', data.error);
      }
    } catch (error) {
      console.error('❌ Slack接続エラー:', error);
    }
  }
  
  // Google Doc ID チェック
  if (!GOOGLE_DOC_ID || GOOGLE_DOC_ID === '***') {
    console.error('❌ Google Doc ID が設定されていません');
    console.log('→ Google ドキュメントを作成し、URLからIDをコピーしてください');
  } else {
    console.log('✅ Google Doc ID: 設定済み');
    try {
      const doc = DocumentApp.openById(GOOGLE_DOC_ID);
      console.log(`✅ Google Doc アクセス成功: ${doc.getName()}`);
    } catch (error) {
      console.error('❌ Google Doc アクセスエラー:', error);
      console.log('→ ドキュメントの共有設定を確認してください');
    }
  }
  
  // OpenAI API Key チェック（オプション）
  if (!OPENAI_API_KEY || OPENAI_API_KEY === '***') {
    console.log('⚠️ OpenAI API Key: 未設定（業務マニュアル生成機能は利用不可）');
  } else {
    console.log('✅ OpenAI API Key: 設定済み');
  }
  
  // Email チェック（オプション）
  if (!NOTIFICATION_EMAIL || NOTIFICATION_EMAIL === '***') {
    console.log('⚠️ 通知メールアドレス: 未設定（日次要約メール送信機能は利用不可）');
  } else {
    console.log(`✅ 通知メールアドレス: ${NOTIFICATION_EMAIL}`);
  }
  
  // スプレッドシートの確認
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    console.log(`✅ スプレッドシート: ${ss.getName()}`);
    
    // 必要なシートの存在確認
    const sheets = ss.getSheets().map(s => s.getName());
    console.log('  現在のシート:', sheets.join(', '));
    
    // 各シートの存在確認
    const requiredSheets = [LOG_SHEET_NAME, LAST_TS_SHEET_NAME, MANUAL_SHEET_NAME, FAQ_SHEET_NAME];
    requiredSheets.forEach(sheetName => {
      if (sheets.includes(sheetName)) {
        console.log(`  ✅ ${sheetName}: 存在`);
      } else {
        console.log(`  ⚠️ ${sheetName}: 未作成（setupSpreadsheet()を実行してください）`);
      }
    });
  } catch (error) {
    console.error('❌ スプレッドシートエラー:', error);
  }
  
  console.log('\n=== チャンネル取得テスト ===');
  try {
    const channels = getSlackChannels();
    console.log(`✅ 取得可能なチャンネル数: ${channels.length}`);
    if (channels.length > 0) {
      console.log('  最初の5チャンネル:');
      channels.slice(0, 5).forEach(ch => {
        console.log(`    - #${ch.name} (${ch.is_private ? 'private' : 'public'})`);
      });
    } else {
      console.log('⚠️ チャンネルが見つかりません。Botをチャンネルに招待してください');
      console.log('→ Slackで /invite @[your-bot-name] を実行');
    }
  } catch (error) {
    console.error('❌ チャンネル取得エラー:', error);
  }
  
  console.log('\n=== セットアップチェック完了 ===');
  console.log('\n【次のステップ】');
  console.log('1. エラーがある場合は修正してください');
  console.log('2. fetchAndAppendAllChannels() を実行してログ収集開始');
  console.log('3. トリガー設定で定期実行を設定（例: 1時間ごと）');
}

/**
 * 初回セットアップヘルパー
 * 必要な設定値を対話的に設定できます
 */
function setupWizard() {
  const ui = SpreadsheetApp.getUi();
  
  ui.alert('Slack ログ収集ツール セットアップ', 
    'このウィザードでは必要な設定を順番に行います。\n\n' +
    '準備するもの:\n' +
    '1. Slack Bot Token\n' +
    '2. Google Doc ID\n' +
    '3. OpenAI API Key（オプション）\n' +
    '4. 通知用メールアドレス（オプション）', 
    ui.ButtonSet.OK);
  
  // Slack Token 入力
  const tokenResult = ui.prompt('Slack Bot Token', 
    'Slack Bot User OAuth Token を入力してください (xoxb-で始まる文字列):', 
    ui.ButtonSet.OK_CANCEL);
  
  if (tokenResult.getSelectedButton() === ui.Button.OK) {
    const token = tokenResult.getResponseText();
    // ここで実際にはスクリプトプロパティに保存することを推奨
    PropertiesService.getScriptProperties().setProperty('SLACK_BOT_TOKEN', token);
    ui.alert('✅ Slack Token を保存しました');
  }
  
  // 続けて他の設定も同様に...
  
  ui.alert('セットアップ完了', 
    'checkSetup() 関数を実行して設定を確認してください', 
    ui.ButtonSet.OK);
}

/**
 * トリガー自動設定関数
 * 定期実行トリガーを自動で設定します
 */
function setupTriggers() {
  // 既存のトリガーを削除
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'fetchAndAppendAllChannels') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  // 新しいトリガーを作成（1時間ごと）
  ScriptApp.newTrigger('fetchAndAppendAllChannels')
    .timeBased()
    .everyHours(1)
    .create();
  
  console.log('✅ 定期実行トリガーを設定しました（1時間ごと）');
  
  // 日次要約メール送信トリガー（毎日午前9時）
  if (NOTIFICATION_EMAIL && NOTIFICATION_EMAIL !== '***') {
    ScriptApp.newTrigger('manualSendDailySummary')
      .timeBased()
      .atHour(9)
      .everyDays(1)
      .create();
    console.log('✅ 日次要約メール送信トリガーを設定しました（毎日9:00）');
  }
}

/**
 * スプレッドシートの初期設定
 * すべての必要なシートを作成し、フォーマットを設定します
 */
function setupSpreadsheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  console.log('=== スプレッドシートセットアップ開始 ===');
  
  // 1. ログシートの作成
  const logSheet = getOrCreateLogSheet();
  console.log('✅ ログシート作成/確認完了:', LOG_SHEET_NAME);
  
  // 2. タイムスタンプシートの作成
  const tsSheet = getOrCreateLastTsSheet();
  console.log('✅ タイムスタンプシート作成/確認完了:', LAST_TS_SHEET_NAME);
  
  // 3. 業務マニュアルシートの作成
  const manualSheet = getOrCreateManualSheet();
  console.log('✅ 業務マニュアルシート作成/確認完了:', MANUAL_SHEET_NAME);
  
  // 4. FAQシートの作成
  const faqSheet = getOrCreateFAQSheet();
  console.log('✅ FAQシート作成/確認完了:', FAQ_SHEET_NAME);
  
  // 5. サマリーシートの作成（ダッシュボード）
  let summarySheet = ss.getSheetByName('summary');
  if (!summarySheet) {
    summarySheet = ss.insertSheet('summary');
    summarySheet.appendRow(['Slack ログ収集ツール ダッシュボード']);
    summarySheet.appendRow(['']);
    summarySheet.appendRow(['統計情報']);
    summarySheet.appendRow(['総メッセージ数:', '=COUNTA(' + LOG_SHEET_NAME + '!E:E)-1']);
    summarySheet.appendRow(['総チャンネル数:', '=COUNTA(UNIQUE(' + LOG_SHEET_NAME + '!B:B))-1']);
    summarySheet.appendRow(['業務マニュアル数:', '=COUNTA(' + MANUAL_SHEET_NAME + '!A:A)-1']);
    summarySheet.appendRow(['FAQ数:', '=COUNTA(' + FAQ_SHEET_NAME + '!A:A)-1']);
    summarySheet.appendRow(['最終更新:', '=MAX(' + LAST_TS_SHEET_NAME + '!C:C)']);
    
    // ダッシュボードのスタイル設定
    summarySheet.getRange(1, 1).setFontSize(18).setFontWeight('bold');
    summarySheet.getRange(3, 1).setFontSize(14).setFontWeight('bold').setBackground('#f0f0f0');
    summarySheet.setColumnWidth(1, 200);
    summarySheet.setColumnWidth(2, 300);
  }
  console.log('✅ サマリーシート作成/確認完了: summary');
  
  // 6. シートの並び替え
  summarySheet.activate();
  ss.moveActiveSheet(1);
  
  console.log('\n=== スプレッドシートセットアップ完了 ===');
  console.log('作成されたシート:');
  ss.getSheets().forEach((sheet, index) => {
    console.log(`  ${index + 1}. ${sheet.getName()}`);
  });
  
  return {
    logSheet: logSheet,
    tsSheet: tsSheet,
    manualSheet: manualSheet,
    faqSheet: faqSheet,
    summarySheet: summarySheet
  };
}

/**
 * データのバックアップ
 * 既存データを別シートにバックアップします
 */
function backupData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const timestamp = new Date().toISOString().replace(/[:.]/g, '-').substring(0, 19);
  
  console.log('=== データバックアップ開始 ===');
  
  // ログシートのバックアップ
  const logSheet = ss.getSheetByName(LOG_SHEET_NAME);
  if (logSheet && logSheet.getLastRow() > 1) {
    const backupSheet = logSheet.copyTo(ss);
    backupSheet.setName(`backup_${LOG_SHEET_NAME}_${timestamp}`);
    console.log(`✅ ${LOG_SHEET_NAME} をバックアップしました`);
  }
  
  // 業務マニュアルシートのバックアップ
  const manualSheet = ss.getSheetByName(MANUAL_SHEET_NAME);
  if (manualSheet && manualSheet.getLastRow() > 1) {
    const backupSheet = manualSheet.copyTo(ss);
    backupSheet.setName(`backup_${MANUAL_SHEET_NAME}_${timestamp}`);
    console.log(`✅ ${MANUAL_SHEET_NAME} をバックアップしました`);
  }
  
  console.log('=== データバックアップ完了 ===');
}

/**
 * 完全セットアップウィザード
 * 初回利用時にすべての設定を行います
 */
function fullSetupWizard() {
  const ui = SpreadsheetApp.getUi();
  
  const result = ui.alert(
    '完全セットアップウィザード',
    'このウィザードでは以下の設定を行います:\n\n' +
    '1. スプレッドシートの初期設定\n' +
    '2. 必要なシートの作成\n' +
    '3. 設定値の確認\n' +
    '4. 接続テスト\n\n' +
    '続行しますか？',
    ui.ButtonSet.YES_NO
  );
  
  if (result !== ui.Button.YES) {
    ui.alert('セットアップをキャンセルしました');
    return;
  }
  
  // スプレッドシートセットアップ
  ui.alert('ステップ 1/4', 'スプレッドシートを初期化しています...', ui.ButtonSet.OK);
  setupSpreadsheet();
  
  // 設定確認
  ui.alert('ステップ 2/4', '設定値を確認しています...', ui.ButtonSet.OK);
  checkSetup();
  
  // 接続テスト
  ui.alert('ステップ 3/4', 'Slack接続をテストしています...', ui.ButtonSet.OK);
  try {
    const channels = getSlackChannels();
    ui.alert('接続成功', `${channels.length}個のチャンネルが見つかりました`, ui.ButtonSet.OK);
  } catch (error) {
    ui.alert('接続エラー', 'Slack接続に失敗しました。設定を確認してください。', ui.ButtonSet.OK);
  }
  
  // トリガー設定
  const triggerResult = ui.alert(
    'ステップ 4/4',
    '定期実行トリガーを設定しますか？',
    ui.ButtonSet.YES_NO
  );
  
  if (triggerResult === ui.Button.YES) {
    setupTriggers();
    ui.alert('✅ トリガー設定完了', '1時間ごとにログを自動収集します', ui.ButtonSet.OK);
  }
  
  ui.alert(
    'セットアップ完了',
    '初期設定が完了しました！\n\n' +
    '次のステップ:\n' +
    '1. fetchAndAppendAllChannels() を実行してテスト\n' +
    '2. 問題がなければ定期実行を開始',
    ui.ButtonSet.OK
  );
}

/**
 * シートのリセット（開発/テスト用）
 * 指定したシートのデータをクリアします
 */
function resetSheet(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    console.log(`シート ${sheetName} が見つかりません`);
    return;
  }
  
  const ui = SpreadsheetApp.getUi();
  const result = ui.alert(
    '確認',
    `${sheetName} のデータをすべて削除しますか？`,
    ui.ButtonSet.YES_NO
  );
  
  if (result === ui.Button.YES) {
    // ヘッダー行を残してデータをクリア
    if (sheet.getLastRow() > 1) {
      sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clear();
      console.log(`✅ ${sheetName} のデータをクリアしました`);
    }
  }
}

/**
 * すべてのシートをリセット（開発/テスト用）
 */
function resetAllSheets() {
  backupData(); // まずバックアップ
  resetSheet(LOG_SHEET_NAME);
  resetSheet(LAST_TS_SHEET_NAME);
  resetSheet(MANUAL_SHEET_NAME);
  resetSheet(FAQ_SHEET_NAME);
  console.log('✅ すべてのシートをリセットしました');
}

// FAQ用シート取得
function getOrCreateFAQSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(FAQ_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(FAQ_SHEET_NAME);
    const headers = ['作成日時', '質問', '回答', 'カテゴリ', 'タグ', '元のチャンネル', '関連メッセージ', 'ステータス'];
    sheet.appendRow(headers);
    
    // ヘッダーのスタイル設定
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#4285f4');
    headerRange.setFontColor('#ffffff');
    headerRange.setFontWeight('bold');
    sheet.setFrozenRows(1);
    
    // 列幅の調整
    sheet.setColumnWidth(1, 150); // 作成日時
    sheet.setColumnWidth(2, 300); // 質問
    sheet.setColumnWidth(3, 400); // 回答
    sheet.setColumnWidth(4, 100); // カテゴリ
    sheet.setColumnWidth(5, 150); // タグ
    sheet.setColumnWidth(6, 120); // 元のチャンネル
    sheet.setColumnWidth(7, 300); // 関連メッセージ
    sheet.setColumnWidth(8, 80);  // ステータス
  }
  return sheet;
}

// マニュアルとFAQを自動判別して生成
function generateManualAndFAQ() {
  const ui = SpreadsheetApp.getUi();
  console.log('=== マニュアル・FAQ自動判別生成開始 ===');
  
  try {
    // OpenAI APIキーのチェック
    if (!OPENAI_API_KEY || OPENAI_API_KEY === '***') {
      ui.alert('エラー', 'OpenAI APIキーが設定されていません。\nconst OPENAI_API_KEY を設定してください。', ui.ButtonSet.OK);
      return;
    }
    
    const sheet = getOrCreateLogSheet();
    const data = sheet.getDataRange().getValues();
    
    if (data.length <= 1) {
      ui.alert('データなし', 'slack_logシートにデータがありません。', ui.ButtonSet.OK);
      return;
    }
    
    // 最新のメッセージから取得（最大100件）
    const maxMessages = Math.min(data.length - 1, 100);
    const startRow = Math.max(1, data.length - maxMessages);
    const allMessages = [];
    
    for (let i = startRow; i < data.length; i++) {
      if (data[i][1] && data[i][4] && data[i][5]) {
        allMessages.push({
          channel: data[i][1],
          user: data[i][4],
          text: data[i][5],
          ts: data[i][2],
          threadTs: data[i][3]
        });
      }
    }
    
    if (allMessages.length === 0) {
      ui.alert('データなし', '処理可能なメッセージがありません。', ui.ButtonSet.OK);
      return;
    }
    
    // プログレス表示
    const progressHtml = HtmlService.createHtmlOutput(
      `<p>メッセージを分析してマニュアルとFAQを自動判別中...</p>
       <p>${allMessages.length}件のメッセージを処理しています。</p>
       <p>この処理には30秒〜1分程度かかる場合があります。</p>`
    ).setWidth(400).setHeight(150);
    ui.showModalDialog(progressHtml, '処理中');
    
    console.log(`${allMessages.length}件のメッセージからマニュアルとFAQを自動判別して生成します`);
    const result = generateContentWithAI(allMessages);
    
    // 結果を表示
    if (result) {
      const successMessage = `生成完了！\n\n` +
        `マニュアル: ${result.manualCount}件\n` +
        `FAQ: ${result.faqCount}件\n\n` +
        `詳細は各シートを確認してください。`;
      ui.alert('成功', successMessage, ui.ButtonSet.OK);
    } else {
      ui.alert('生成失敗', 'コンテンツの生成に失敗しました。ログを確認してください。', ui.ButtonSet.OK);
    }
  } catch (error) {
    console.error('マニュアル・FAQ生成エラー:', error.toString());
    ui.alert('エラー', `エラーが発生しました:\n${error.toString()}`, ui.ButtonSet.OK);
  }
}

// AIを使用してコンテンツを分類して生成
function generateContentWithAI(messages, retryCount = 0) {
  console.log(`=== AIコンテンツ生成開始: ${messages.length}件のメッセージ ===`);
  
  if (!OPENAI_API_KEY || OPENAI_API_KEY === '***') {
    console.error('OpenAI APIキーが設定されていません');
    return null;
  }
  
  const conversationText = formatMessagesForAI(messages);
  
  if (!conversationText || conversationText.length < 20) {
    console.log(`会話テキストが短すぎます: ${conversationText?.length || 0}文字`);
    return null;
  }
  
  const prompt = `以下のSlackでのやり取りを分析し、業務マニュアルとFAQに分類して抽出してください。

分類基準：
【業務マニュアル】
- 体系的なプロセスや手順
- 詳細な作業ステップが必要な内容
- 意思決定や承認プロセス
- プロジェクト管理方法
- セキュリティ・コンプライアンス関連

【FAQ】
- 簡単な質問と回答
- トラブルシューティング
- 日常的な問い合わせ
- ツールの使い方
- 用語の説明
- 簡潔な情報共有

重要：
- 情報量が少ないものはFAQとして扱う
- どちらにも分類できない場合はFAQとして扱う
- 必ず最低1つ以上の項目を生成する

出力形式：
=== MANUAL START ===
カテゴリ: [カテゴリ名]
タイトル: [タイトル]
内容: [詳細な内容]
---
=== MANUAL END ===

=== FAQ START ===
質問: [質問文]
回答: [回答文]
カテゴリ: [カテゴリ]
タグ: [関連タグ、カンマ区切り]
---
=== FAQ END ===

やり取り内容：
${conversationText}`;
  
  try {
    const url = 'https://api.openai.com/v1/responses';
    const options = {
      method: 'post',
      headers: {
        'Authorization': 'Bearer ' + OPENAI_API_KEY,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify({
        model: 'gpt-5',
        input: `SYSTEM:\nあなたは企業の知識管理専門家です。会話から業務マニュアルとFAQを適切に分類して抽出してください。\n\nUSER:\n${prompt}`,
        temperature: 0.3,
        max_output_tokens: 4096
      })
    };
    
    console.log('OpenAI APIを呼び出し中...');
    const response = UrlFetchApp.fetch(url, options);
    const data = JSON.parse(response.getContentText());
    const content = extractTextFromOpenAIResponse(data);
    if (!content) {
      console.error('APIレスポンスが不正:', data);
      return null;
    }
    console.log(`生成されたコンテンツ: ${content.substring(0, 200)}...`);
    
    // マニュアルとFAQを分離して処理
    const manualCount = processManuals(content, messages);
    const faqCount = processFAQs(content, messages);
    
    return {
      manualCount: manualCount,
      faqCount: faqCount
    };
    
  } catch (error) {
    console.error('コンテンツ生成エラー:', error.toString());
    
    // リトライ処理
    if (error.toString().includes('429') || error.toString().includes('rate limit')) {
      if (retryCount < 3) {
        console.log(`${5 * (retryCount + 1)}秒後にリトライします...`);
        Utilities.sleep(5000 * (retryCount + 1));
        return generateContentWithAI(messages, retryCount + 1);
      }
    }
    
    return null;
  }
}

// マニュアルを処理
function processManuals(content, messages) {
  const sheet = getOrCreateManualSheet();
  const channels = [...new Set(messages.map(m => m.channel))].join(', ');
  const timestamp = new Date().toLocaleString();
  
  // マニュアル部分を抽出
  const manualRegex = /=== MANUAL START ===[\s\S]*?=== MANUAL END ===/g;
  const manualMatches = content.match(manualRegex) || [];
  
  let count = 0;
  manualMatches.forEach(manual => {
    const categoryMatch = manual.match(/カテゴリ[:：]\s*(.+)/);
    const titleMatch = manual.match(/タイトル[:：]\s*(.+)/);
    const contentMatch = manual.match(/内容[:：]\s*([\s\S]+?)(?=---|=== MANUAL END ===)/);
    
    const category = categoryMatch ? categoryMatch[1].trim() : 'その他';
    const title = titleMatch ? titleMatch[1].trim() : '無題のマニュアル';
    const manualContent = contentMatch ? contentMatch[1].trim() : '';
    
    if (title && manualContent) {
      sheet.appendRow([
        timestamp,
        category,
        title,
        manualContent,
        channels,
        messages[0].text.substring(0, 200),
        '新規'
      ]);
      
      // Googleドキュメントにも追記
      try {
        appendManualToGoogleDoc(category, title, manualContent);
      } catch (docError) {
        console.error('Google Docへの追記エラー:', docError);
      }
      
      count++;
    }
  });
  
  console.log(`マニュアル${count}件を保存しました`);
  return count;
}

// FAQを処理
function processFAQs(content, messages) {
  const sheet = getOrCreateFAQSheet();
  const channels = [...new Set(messages.map(m => m.channel))].join(', ');
  const timestamp = new Date().toLocaleString();
  
  // FAQ部分を抽出
  const faqRegex = /=== FAQ START ===[\s\S]*?=== FAQ END ===/g;
  const faqMatches = content.match(faqRegex) || [];
  
  let count = 0;
  faqMatches.forEach(faq => {
    const questionMatch = faq.match(/質問[:：]\s*(.+)/);
    const answerMatch = faq.match(/回答[:：]\s*([\s\S]+?)(?=\nカテゴリ|\nタグ|---|=== FAQ END ===)/);
    const categoryMatch = faq.match(/カテゴリ[:：]\s*(.+)/);
    const tagMatch = faq.match(/タグ[:：]\s*(.+)/);
    
    const question = questionMatch ? questionMatch[1].trim() : '';
    const answer = answerMatch ? answerMatch[1].trim() : '';
    const category = categoryMatch ? categoryMatch[1].trim() : 'その他';
    const tags = tagMatch ? tagMatch[1].trim() : '';
    
    if (question && answer) {
      sheet.appendRow([
        timestamp,
        question,
        answer,
        category,
        tags,
        channels,
        messages[0].text.substring(0, 200),
        '新規'
      ]);
      
      // GoogleドキュメントにもFAQを追記
      try {
        appendFAQToGoogleDoc(question, answer, category, tags);
      } catch (docError) {
        console.error('Google DocへのFAQ追記エラー:', docError);
      }
      
      count++;
    }
  });
  
  console.log(`FAQ${count}件を保存しました`);
  return count;
}

/**
 * ========================================
 * 改良版：小さなタスク単位での生成機能
 * ========================================
 */

/**
 * 改良版：業務マニュアル生成（トピック別）
 */
function generateBusinessManualImproved(messages) {
  console.log(`=== 改良版マニュアル生成開始: ${messages.length}件のメッセージ ===`);
  
  if (!OPENAI_API_KEY || OPENAI_API_KEY === '***') {
    console.error('OpenAI APIキーが設定されていません');
    return null;
  }
  
  const sheet = getOrCreateManualSheet();
  const results = [];
  
  // メッセージをトピック別に分類
  const topics = classifyMessagesByTopic(messages);
  console.log(`${topics.length}個の独立したトピックを検出`);
  
  // 各トピックを個別に処理
  for (let i = 0; i < topics.length; i++) {
    const topicMessages = topics[i];
    const context = analyzeMessageContext(topicMessages);
    
    console.log(`トピック ${i + 1}/${topics.length}: ${context.messageCount}件のメッセージ, タイプ: ${context.estimatedType}`);
    
    // メッセージが少なすぎる場合はスキップ
    if (topicMessages.length < 1) continue;
    
    // トピックごとにマニュアルを生成
    const manualItem = generateSingleManualItem(topicMessages, context);
    if (manualItem) {
      results.push(manualItem);
      saveManualToSheetImproved(sheet, manualItem, topicMessages);
    }
    
    // API制限対策のため少し待機
    Utilities.sleep(500);
  }
  
  console.log(`生成完了: ${results.length}件のマニュアル項目`);
  return results;
}

/**
 * 単一のマニュアル項目を生成（改良版）
 */
function generateSingleManualItem(messages, context) {
  const conversationText = formatMessagesForAI(messages);
  
  // コンテキストに応じたプロンプトを生成
  const prompt = createContextAwarePrompt(conversationText, context);
  
  try {
    const url = 'https://api.openai.com/v1/responses';
    const response = UrlFetchApp.fetch(url, {
      method: 'post',
      headers: {
        'Authorization': 'Bearer ' + OPENAI_API_KEY,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify({
        model: 'gpt-5',
        input: `SYSTEM:\nあなたは業務文書作成の専門家です。与えられた会話から、独立した1つの明確なタスクや手順を抽出してください。\n複数の異なるタスクを無理に1つにまとめないでください。最も重要な1つのポイントに焦点を当ててください。\n\nUSER:\n${prompt}`,
        temperature: 0.3,
        max_output_tokens: 2000
      })
    });
    
    const data = JSON.parse(response.getContentText());
    const content = extractTextFromOpenAIResponse(data);
    
    return parseManualContentImproved(content, context);
    
  } catch (error) {
    console.error('マニュアル生成エラー:', error);
    return null;
  }
}

/**
 * コンテキストに応じたプロンプトを生成（改良版）
 */
function createContextAwarePrompt(conversationText, context) {
  let promptType = '';
  
  switch (context.estimatedType) {
    case 'TROUBLESHOOTING':
      promptType = `
この会話から、具体的な問題と解決方法を1つ抽出してください。
複数の問題がある場合は、最も重要な1つに絞ってください。

出力形式：
カテゴリ: トラブルシューティング
タイトル: [具体的な問題]
問題の症状: [具体的な症状]
原因: [判明した原因]
解決手順:
1. [手順1]
2. [手順2]
...
確認方法: [解決を確認する方法]
予防策: [再発防止策]`;
      break;
      
    case 'FAQ':
      promptType = `
この会話から、最も重要な質問と回答を1つ抽出してください。
複数の質問がある場合は、それぞれ独立して扱い、ここでは1つだけ出力してください。

出力形式：
カテゴリ: FAQ
質問: [明確な質問文]
回答: [簡潔で正確な回答]
補足情報: [必要に応じて]
関連事項: [関連する他の情報]`;
      break;
      
    case 'DECISION':
      promptType = `
この会話から、行われた意思決定を1つ抽出してください。
複数の決定がある場合は、最も重要な1つに絞ってください。

出力形式：
カテゴリ: 意思決定記録
タイトル: [決定事項]
背景: [決定に至った背景]
決定内容: [具体的な決定内容]
理由: [決定の根拠]
実行事項: [必要なアクション]
責任者: [担当者または部門]
期限: [実施期限]`;
      break;
      
    case 'PROCEDURE':
      promptType = `
この会話から、具体的な作業手順を1つ抽出してください。
複数の手順がある場合は、最も完結した1つのタスクに絞ってください。

出力形式：
カテゴリ: 作業手順
タイトル: [作業名]
目的: [この作業の目的]
前提条件: [必要な準備や条件]
手順:
1. [手順1]
2. [手順2]
...
確認事項: [完了確認の方法]
注意点: [気をつけるべきこと]`;
      break;
      
    default:
      promptType = `
この会話から、業務に有用な情報を1つ抽出してください。
複数のトピックがある場合は、最も重要な1つに絞ってください。

出力形式：
カテゴリ: [適切なカテゴリ]
タイトル: [内容を表す明確なタイトル]
内容: [詳細な説明]
ポイント: [重要な点]
関連情報: [参考になる情報]`;
  }
  
  return `${promptType}

会話内容：
${conversationText}

注意事項：
- 1つの独立したトピックとして完結させてください
- 無関係な複数のタスクを混ぜないでください
- 具体的で実用的な内容にしてください
- 推測や一般化は最小限にしてください`;
}

/**
 * マニュアルコンテンツをパース（改良版）
 */
function parseManualContentImproved(content, context) {
  const lines = content.split('\n');
  const manual = {
    category: '',
    title: '',
    content: '',
    keywords: context.keywords.join(', '),
    participants: context.participants.join(', '),
    messageCount: context.messageCount
  };
  
  let currentSection = '';
  
  for (const line of lines) {
    if (line.startsWith('カテゴリ:')) {
      manual.category = line.replace('カテゴリ:', '').trim();
    } else if (line.startsWith('タイトル:')) {
      manual.title = line.replace('タイトル:', '').trim();
    } else if (line.startsWith('質問:')) {
      manual.title = '【FAQ】' + line.replace('質問:', '').trim();
      currentSection = 'content';
    } else if (currentSection || (!manual.category && !manual.title)) {
      manual.content += line + '\n';
    }
  }
  
  // 内容が空でなければ返す
  if (manual.title && manual.content) {
    manual.content = manual.content.trim();
    return manual;
  }
  
  return null;
}

/**
 * マニュアルをシートに保存（改良版）
 */
function saveManualToSheetImproved(sheet, manual, originalMessages) {
  const timestamp = new Date();
  const channelName = originalMessages[0]?.channel || '';
  const messageIds = originalMessages.map(m => `${m.channel}_${m.ts}`).join(', ');
  
  // 既存のシート構造に合わせて保存
  sheet.appendRow([
    timestamp.toLocaleString(),
    manual.category || 'その他',
    manual.title,
    manual.content,
    channelName,
    messageIds.substring(0, 500), // 関連メッセージIDを短縮
    'アクティブ'
  ]);
  
  console.log(`保存: ${manual.title}`);
}

/**
 * 改良版FAQ生成
 */
function generateFAQImproved(messages) {
  console.log(`=== 改良版FAQ生成開始: ${messages.length}件のメッセージ ===`);
  
  if (!OPENAI_API_KEY || OPENAI_API_KEY === '***') {
    console.error('OpenAI APIキーが設定されていません');
    return null;
  }
  
  const sheet = getOrCreateFAQSheet();
  const results = [];
  
  // Q&Aパターンを検出
  const qaPairs = detectQAPairs(messages);
  console.log(`${qaPairs.length}個のQ&Aペアを検出`);
  
  // 各Q&Aペアを個別に処理
  for (const qaPair of qaPairs) {
    const faqItem = generateSingleFAQ(qaPair);
    if (faqItem) {
      results.push(faqItem);
      saveFAQToSheetImproved(sheet, faqItem, qaPair.messages);
    }
    
    // API制限対策
    Utilities.sleep(500);
  }
  
  console.log(`生成完了: ${results.length}件のFAQ`);
  return results;
}

/**
 * Q&Aペアを検出（改良版）
 */
function detectQAPairs(messages) {
  const pairs = [];
  
  for (let i = 0; i < messages.length; i++) {
    const msg = messages[i];
    const text = (msg.text || '').toLowerCase();
    
    // 質問パターンを検出
    if (text.match(/[?？]|どう|なぜ|いつ|どこ|誰|何/)) {
      // 次の数メッセージを回答候補として収集
      const relatedMessages = [msg];
      const threadTs = msg.threadTs || msg.ts;
      
      for (let j = i + 1; j < Math.min(i + 10, messages.length); j++) {
        const nextMsg = messages[j];
        
        // 同じスレッドまたは直後のメッセージ
        if (nextMsg.threadTs === threadTs || 
            (Math.abs(parseFloat(nextMsg.ts) - parseFloat(msg.ts)) < 300)) {
          relatedMessages.push(nextMsg);
        } else {
          break;
        }
      }
      
      // 回答が含まれていそうな場合のみ追加
      if (relatedMessages.length > 1) {
        pairs.push({
          question: msg,
          messages: relatedMessages
        });
        
        // 処理済みメッセージをスキップ
        i += relatedMessages.length - 1;
      }
    }
  }
  
  return pairs;
}

/**
 * 単一のFAQを生成（改良版）
 */
function generateSingleFAQ(qaPair) {
  const conversationText = formatMessagesForAI(qaPair.messages);
  
  const prompt = `以下の会話から、1つの明確な質問と回答を抽出してください。

出力形式：
質問: [ユーザーの質問を明確に]
回答: [簡潔で分かりやすい回答]
カテゴリ: [適切なカテゴリ]
タグ: [関連キーワード、カンマ区切り]
補足: [必要に応じて追加情報]

会話内容：
${conversationText}

注意：
- 質問と回答は1対1で明確にしてください
- 複数の質問を混ぜないでください
- 回答は実用的で具体的にしてください`;
  
  try {
    const url = 'https://api.openai.com/v1/responses';
    const response = UrlFetchApp.fetch(url, {
      method: 'post',
      headers: {
        'Authorization': 'Bearer ' + OPENAI_API_KEY,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify({
        model: 'gpt-5',
        input: `SYSTEM:\nFAQ作成の専門家として、明確で有用なQ&Aを作成してください。\n\nUSER:\n${prompt}`,
        temperature: 0.3,
        max_output_tokens: 1000
      })
    });
    
    const data = JSON.parse(response.getContentText());
    const content = extractTextFromOpenAIResponse(data);
    
    return parseFAQContentImproved(content);
    
  } catch (error) {
    console.error('FAQ生成エラー:', error);
    return null;
  }
}

/**
 * FAQコンテンツをパース（改良版）
 */
function parseFAQContentImproved(content) {
  const lines = content.split('\n');
  const faq = {
    question: '',
    answer: '',
    category: '',
    tags: '',
    supplement: ''
  };
  
  for (const line of lines) {
    if (line.startsWith('質問:')) {
      faq.question = line.replace('質問:', '').trim();
    } else if (line.startsWith('回答:')) {
      faq.answer = line.replace('回答:', '').trim();
    } else if (line.startsWith('カテゴリ:')) {
      faq.category = line.replace('カテゴリ:', '').trim();
    } else if (line.startsWith('タグ:')) {
      faq.tags = line.replace('タグ:', '').trim();
    } else if (line.startsWith('補足:')) {
      faq.supplement = line.replace('補足:', '').trim();
    }
  }
  
  // 質問と回答があれば返す
  if (faq.question && faq.answer) {
    return faq;
  }
  
  return null;
}

/**
 * FAQをシートに保存（改良版）
 */
function saveFAQToSheetImproved(sheet, faq, originalMessages) {
  const timestamp = new Date();
  const channelName = originalMessages[0]?.channel || '';
  const messageIds = originalMessages.map(m => `${m.channel}_${m.ts}`).join(', ');
  
  const fullAnswer = faq.answer + (faq.supplement ? '\n\n補足: ' + faq.supplement : '');
  
  sheet.appendRow([
    timestamp.toLocaleString(),
    faq.question,
    fullAnswer,
    faq.category || 'その他',
    faq.tags || '',
    channelName,
    messageIds.substring(0, 500), // 関連メッセージIDを短縮
    'アクティブ'
  ]);
  
  console.log(`FAQ保存: ${faq.question.substring(0, 50)}...`);
}

/**
 * 改良版：マニュアルとFAQを自動生成（エントリーポイント）
 */
function generateManualAndFAQImproved() {
  const ui = SpreadsheetApp.getUi();
  console.log('=== 改良版：マニュアル・FAQ生成開始 ===');
  
  try {
    // OpenAI APIキーのチェック
    if (!OPENAI_API_KEY || OPENAI_API_KEY === '***') {
      ui.alert('エラー', 'OpenAI APIキーが設定されていません。', ui.ButtonSet.OK);
      return;
    }
    
    const sheet = getOrCreateLogSheet();
    const data = sheet.getDataRange().getValues();
    
    if (data.length <= 1) {
      ui.alert('データなし', 'slack_logシートにデータがありません。', ui.ButtonSet.OK);
      return;
    }
    
    // 過去24時間のメッセージを取得
    const now = new Date();
    const yesterday = new Date(now.getTime() - 24 * 60 * 60 * 1000);
    const messages = [];
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const date = row[6]; // date列
      
      if (date instanceof Date && date >= yesterday) {
        messages.push({
          channel: row[1],
          user: row[4],
          text: row[5],
          ts: row[2],
          threadTs: row[3],
          date: row[6]
        });
      }
    }
    
    console.log(`過去24時間のメッセージ: ${messages.length}件`);
    
    if (messages.length === 0) {
      ui.alert('情報', '過去24時間にメッセージがありません。', ui.ButtonSet.OK);
      return;
    }
    
    // プログレス表示
    const progressHtml = HtmlService.createHtmlOutput(
      `<p>メッセージをトピック別に分析中...</p>
       <p>${messages.length}件のメッセージを処理しています。</p>
       <p>各トピックを独立した文書として生成します。</p>`
    ).setWidth(400).setHeight(150);
    ui.showModalDialog(progressHtml, '処理中');
    
    // チャンネルごとに処理
    const channelMap = {};
    messages.forEach(msg => {
      if (!channelMap[msg.channel]) {
        channelMap[msg.channel] = [];
      }
      channelMap[msg.channel].push(msg);
    });
    
    let totalManuals = 0;
    let totalFAQs = 0;
    
    for (const [channelName, channelMessages] of Object.entries(channelMap)) {
      console.log(`\nチャンネル: ${channelName} (${channelMessages.length}件)`);
      
      // マニュアル生成
      const manuals = generateBusinessManualImproved(channelMessages);
      if (manuals) totalManuals += manuals.length;
      
      // FAQ生成
      const faqs = generateFAQImproved(channelMessages);
      if (faqs) totalFAQs += faqs.length;
    }
    
    ui.alert(
      '生成完了',
      `改良版で生成されました！\n\n` +
      `マニュアル: ${totalManuals}件\n` +
      `FAQ: ${totalFAQs}件\n\n` +
      `各項目は独立したタスクとして生成されました。`,
      ui.ButtonSet.OK
    );
    
  } catch (error) {
    console.error('改良版生成エラー:', error.toString());
    ui.alert('エラー', `エラーが発生しました:\n${error.toString()}`, ui.ButtonSet.OK);
  }
}

// GoogleドキュメントにFAQを追記
function appendFAQToGoogleDoc(question, answer, category, tags) {
  try {
    const doc = DocumentApp.openById(GOOGLE_DOC_ID);
    const body = doc.getBody();
    
    // FAQセクションを追加
    const faqNumber = getNextFAQNumber();
    
    // FAQヘッダー
    const header = body.appendParagraph(`❓ FAQ #${faqNumber}`);
    header.setHeading(DocumentApp.ParagraphHeading.HEADING2);
    header.setFontSize(16);
    header.setForegroundColor('#ea4335');
    header.setBold(true);
    header.setSpacingBefore(30);
    
    // メタ情報
    const metaInfo = body.appendParagraph(`🎯 カテゴリ: ${category} | 🏿️ タグ: ${tags || 'なし'}`);
    metaInfo.setFontSize(10);
    metaInfo.setForegroundColor('#5f6368');
    metaInfo.setSpacingAfter(10);
    
    // 質問
    const qPara = body.appendParagraph('Q: ' + question);
    qPara.setFontSize(14);
    qPara.setBold(true);
    qPara.setForegroundColor('#1a73e8');
    qPara.setSpacingAfter(10);
    
    // 回答
    const aPara = body.appendParagraph('A: ' + answer);
    aPara.setFontSize(12);
    aPara.setLineSpacing(1.5);
    aPara.setIndentFirstLine(20);
    aPara.setSpacingAfter(20);
    
    // 区切り線
    body.appendHorizontalRule();
    
    doc.saveAndClose();
    console.log(`FAQ #${faqNumber}「${question}」をGoogleドキュメントに追加しました`);
  } catch (error) {
    console.error('Google Doc FAQ追記エラー:', error);
  }
}

// 次のFAQ番号を取得
function getNextFAQNumber() {
  try {
    const props = PropertiesService.getScriptProperties();
    let faqCount = parseInt(props.getProperty('FAQ_COUNT') || '0');
    faqCount++;
    props.setProperty('FAQ_COUNT', faqCount.toString());
    return faqCount;
  } catch (error) {
    console.error('FAQ番号取得エラー:', error);
    return 1;
  }
}

// FAQのみを生成
function manualGenerateFAQ() {
  const ui = SpreadsheetApp.getUi();
  console.log('=== FAQ生成開始 ===');
  
  try {
    if (!OPENAI_API_KEY || OPENAI_API_KEY === '***') {
      ui.alert('エラー', 'OpenAI APIキーが設定されていません。', ui.ButtonSet.OK);
      return;
    }
    
    const sheet = getOrCreateLogSheet();
    const data = sheet.getDataRange().getValues();
    
    if (data.length <= 1) {
      ui.alert('データなし', 'slack_logシートにデータがありません。', ui.ButtonSet.OK);
      return;
    }
    
    // 最新のメッセージから取得
    const maxMessages = Math.min(data.length - 1, 100);
    const startRow = Math.max(1, data.length - maxMessages);
    const allMessages = [];
    
    for (let i = startRow; i < data.length; i++) {
      if (data[i][1] && data[i][4] && data[i][5]) {
        allMessages.push({
          channel: data[i][1],
          user: data[i][4],
          text: data[i][5],
          ts: data[i][2],
          threadTs: data[i][3]
        });
      }
    }
    
    if (allMessages.length === 0) {
      ui.alert('データなし', '処理可能なメッセージがありません。', ui.ButtonSet.OK);
      return;
    }
    
    console.log(`${allMessages.length}件のメッセージからFAQを生成します`);
    const result = generateFAQOnly(allMessages);
    
    if (result && result.count > 0) {
      ui.alert('成功', `FAQを${result.count}件生成しました！`, ui.ButtonSet.OK);
    } else {
      ui.alert('生成失敗', 'FAQの生成に失敗しました。', ui.ButtonSet.OK);
    }
  } catch (error) {
    console.error('FAQ生成エラー:', error.toString());
    ui.alert('エラー', `エラーが発生しました:\n${error.toString()}`, ui.ButtonSet.OK);
  }
}

// FAQのみを生成する関数
function generateFAQOnly(messages) {
  const conversationText = formatMessagesForAI(messages);
  
  const prompt = `以下のSlackでのやり取りから、FAQ（よくある質問と回答）を抽出してください。
短い会話や簡単な質問も積極的にFAQ化してください。

出力形式：
=== FAQ START ===
質問: [質問文]
回答: [簡潔でわかりやすい回答]
カテゴリ: [カテゴリ]
タグ: [関連タグ、カンマ区切り]
---
=== FAQ END ===

やり取り内容：
${conversationText}`;
  
  try {
    const url = 'https://api.openai.com/v1/responses';
    const response = UrlFetchApp.fetch(url, {
      method: 'post',
      headers: {
        'Authorization': 'Bearer ' + OPENAI_API_KEY,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify({
        model: 'gpt-5',
        input: `SYSTEM:\nFAQを作成する専門家です。会話から有用な質問と回答を抽出してください。\n\nUSER:\n${prompt}`,
        temperature: 0.3,
        max_output_tokens: 4096
      })
    });
    
    const data = JSON.parse(response.getContentText());
    const content = extractTextFromOpenAIResponse(data);
    
    const faqCount = processFAQs(content, messages);
    return { count: faqCount };
    
  } catch (error) {
    console.error('FAQ生成エラー:', error);
    return null;
  }
}
