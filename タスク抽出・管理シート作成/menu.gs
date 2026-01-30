// カスタムメニューとセットアップ機能

// スプレッドシート開いた時の処理
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  ui.createMenu('📋 タスク管理システム')
    .addSubMenu(ui.createMenu('⚙️ システム')
      .addItem('🚀 初回セットアップ', 'setupSystem')
      .addItem('🔧 設定を開く', 'openConfigSheet')
      .addSeparator()
      .addItem('🔑 APIキーを設定', 'setApiKey')
      .addItem('⏰ トリガーを設定', 'setupTriggers')
      .addItem('🗑️ トリガーを削除', 'deleteTriggers'))
    .addSubMenu(ui.createMenu('📧 メール')
      .addItem('✉️ 業務メール作成', 'showEmailComposer')
      .addItem('📥 新着メール処理を今すぐ実行', 'processNewEmailsManually')
      .addSeparator()
      .addItem('🏷️ 処理済みラベルを作成', 'createProcessedLabel'))
    .addSubMenu(ui.createMenu('📊 フロー')
      .addItem('🎨 ビジュアルフロー生成', 'generateVisualFlow')
      .addItem('📝 サンプルデータ作成', 'createSampleFlowData')
      .addSeparator()
      .addItem('🔄 フローシートをリセット', 'resetFlowSheet')
      .addItem('🧪 新エンジンテスト', 'testNewDataEngine'))
    .addSubMenu(ui.createMenu('📈 レポート')
      .addItem('📊 処理統計を表示', 'showProcessingStats')
      .addItem('📋 アクティビティログを表示', 'showActivityLog'))
    .addSeparator()
    .addItem('❓ ヘルプ', 'showHelp')
    .addItem('ℹ️ バージョン情報', 'showAbout')
    .addToUi();
    
  // 初回起動チェック
  checkFirstRun();
}

// 初回セットアップ
function setupSystem() {
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.alert(
    '初回セットアップ',
    '以下の処理を実行します：\n\n' +
    '1. 必要なシートの作成\n' +
    '2. 初期設定の配置\n' +
    '3. APIキーの設定確認\n' +
    '4. タイマートリガーの設定\n\n' +
    '続行しますか？',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) {
    return;
  }
  
  try {
    // プログレス表示
    const progressHtml = HtmlService.createHtmlOutput(getProgressHtml())
      .setWidth(400)
      .setHeight(200);
    ui.showModalDialog(progressHtml, 'セットアップ中...');
    
    // 1. シート作成
    createRequiredSheets();
    
    // 2. 初期設定
    initializeConfig();
    
    // 3. APIキー確認
    const apiKey = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
    if (!apiKey) {
      ui.alert(
        '⚠️ APIキー未設定',
        'OpenAI APIキーが設定されていません。\n' +
        'メニューから「システム > APIキーを設定」を選択して設定してください。',
        ui.ButtonSet.OK
      );
    }
    
    // 4. トリガー設定
    setupTriggers();
    
    // 完了メッセージ
    ui.alert(
      '✅ セットアップ完了',
      'システムのセットアップが完了しました。\n\n' +
      '次のステップ：\n' +
      '1. Config シートで設定を確認\n' +
      '2. APIキーを設定（未設定の場合）\n' +
      '3. テストメールを送信して動作確認',
      ui.ButtonSet.OK
    );
    
  } catch (e) {
    ui.alert(
      '❌ セットアップエラー',
      'セットアップ中にエラーが発生しました：\n' + e.toString(),
      ui.ButtonSet.OK
    );
    logActivity('SETUP_ERROR', e.toString());
  }
}

// 必要なシートの作成
function createRequiredSheets() {
  const requiredSheets = [
    CONFIG_SHEET,
    INBOX_SHEET,
    SPEC_SHEET,
    FLOW_SHEET,
    VISUAL_SHEET,
    ACTIVITY_LOG_SHEET
  ];
  
  requiredSheets.forEach(sheetName => {
    if (!ss().getSheetByName(sheetName)) {
      if (sheetName === CONFIG_SHEET) {
        initializeConfig();
      } else if (sheetName === INBOX_SHEET) {
        createInboxSheet();
      } else if (sheetName === SPEC_SHEET) {
        createWorkSpecSheet();
      } else if (sheetName === FLOW_SHEET) {
        createFlowSheet(sheetName);
      } else if (sheetName === ACTIVITY_LOG_SHEET) {
        createActivityLogSheet();
      } else {
        ss().insertSheet(sheetName);
      }
    }
  });
  
  logActivity('SETUP', 'Required sheets created');
}

// APIキー設定
function setApiKey() {
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.prompt(
    'OpenAI APIキー設定',
    'OpenAI APIキーを入力してください：\n' +
    '（キーは安全に保存されます）',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() === ui.Button.OK) {
    const apiKey = response.getResponseText().trim();
    
    if (apiKey) {
      PropertiesService.getScriptProperties().setProperty('OPENAI_API_KEY', apiKey);
      ui.alert('✅ APIキーを設定しました。');
      logActivity('API_KEY', 'API key configured');
    } else {
      ui.alert('⚠️ APIキーが入力されていません。');
    }
  }
}

// トリガー設定
function setupTriggers() {
  // 既存のトリガーを削除
  deleteTriggers();
  
  // 時間ベーストリガーを作成（5分ごと）
  ScriptApp.newTrigger('processNewEmails')
    .timeBased()
    .everyMinutes(5)
    .create();
    
  logActivity('TRIGGER', 'Time-based trigger created (every 5 minutes)');
}

// トリガー削除
function deleteTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'processNewEmails') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  logActivity('TRIGGER', 'Existing triggers deleted');
}

// 手動でメール処理実行
function processNewEmailsManually() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    ui.alert('📥 処理中...', 'メールを処理しています。しばらくお待ちください。', ui.ButtonSet.OK);
    processNewEmails();
    ui.alert('✅ 完了', 'メール処理が完了しました。', ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('❌ エラー', 'メール処理中にエラーが発生しました：\n' + e.toString(), ui.ButtonSet.OK);
  }
}

// Config シートを開く
function openConfigSheet() {
  const sheet = ss().getSheetByName(CONFIG_SHEET);
  if (sheet) {
    ss().setActiveSheet(sheet);
  } else {
    SpreadsheetApp.getUi().alert('Config シートが見つかりません。');
  }
}

// 処理済みラベル作成
function createProcessedLabel() {
  try {
    let label = GmailApp.getUserLabelByName('PROCESSED');
    if (!label) {
      label = GmailApp.createLabel('PROCESSED');
      SpreadsheetApp.getUi().alert('✅ PROCESSEDラベルを作成しました。');
    } else {
      SpreadsheetApp.getUi().alert('ℹ️ PROCESSEDラベルは既に存在します。');
    }
  } catch (e) {
    SpreadsheetApp.getUi().alert('❌ エラー：' + e.toString());
  }
}

// フローシートリセット
function resetFlowSheet() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    '確認',
    'フローシートのデータをすべて削除しますか？',
    ui.ButtonSet.YES_NO
  );
  
  if (response === ui.Button.YES) {
    const sheet = ss().getSheetByName(FLOW_SHEET);
    if (sheet) {
      sheet.clear();
      const headers = FLOW_HEADERS;
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#e8f5e9');
      ui.alert('✅ フローシートをリセットしました。');
    }
  }
}

// 処理統計表示
function showProcessingStats() {
  const inboxSheet = ss().getSheetByName(INBOX_SHEET);
  if (!inboxSheet || inboxSheet.getLastRow() <= 1) {
    SpreadsheetApp.getUi().alert('処理データがありません。');
    return;
  }
  
  const data = inboxSheet.getRange(2, 7, inboxSheet.getLastRow() - 1, 1).getValues();
  const stats = {
    total: data.length,
    processed: data.filter(row => row[0] === 'PROCESSED').length,
    error: data.filter(row => row[0] === 'ERROR').length,
    new: data.filter(row => row[0] === 'NEW').length
  };
  
  const message = `📊 処理統計\n\n` +
    `合計: ${stats.total} 件\n` +
    `処理済み: ${stats.processed} 件\n` +
    `エラー: ${stats.error} 件\n` +
    `未処理: ${stats.new} 件`;
    
  SpreadsheetApp.getUi().alert(message);
}

// アクティビティログ表示
function showActivityLog() {
  const sheet = ss().getSheetByName(ACTIVITY_LOG_SHEET);
  if (sheet) {
    sheet.showSheet();
    ss().setActiveSheet(sheet);
  } else {
    SpreadsheetApp.getUi().alert('アクティビティログがありません。');
  }
}

// ヘルプ表示
function showHelp() {
  const helpText = `📋 タスク管理システム - ヘルプ\n\n` +
    `【基本的な使い方】\n` +
    `1. 初回セットアップを実行\n` +
    `2. OpenAI APIキーを設定\n` +
    `3. Config シートで設定を調整\n` +
    `4. メール受信または送信で業務記述書を自動生成\n\n` +
    `【メール処理】\n` +
    `- 件名に [WORK-REQ] を含むメールを自動処理\n` +
    `- 5分ごとに自動チェック（変更可能）\n` +
    `- 処理結果は送信者にメール通知\n\n` +
    `【トラブルシューティング】\n` +
    `- エラーが発生した場合はInboxシートを確認\n` +
    `- APIキーが正しく設定されているか確認\n` +
    `- アクティビティログで詳細を確認`;
    
  SpreadsheetApp.getUi().alert(helpText);
}

// バージョン情報表示
function showAbout() {
  const about = `📋 タスク管理システム\n\n` +
    `バージョン: 1.0.0\n` +
    `作成日: 2024\n` +
    `説明: メールから業務記述書とタスクフローを自動生成\n\n` +
    `機能:\n` +
    `- OpenAI GPTによる業務記述書生成\n` +
    `- ビジュアルフロー自動描画\n` +
    `- Gmail連携による自動処理\n` +
    `- 上場企業レベルの品質管理`;
    
  SpreadsheetApp.getUi().alert(about);
}

// 初回起動チェック
function checkFirstRun() {
  const isFirstRun = PropertiesService.getDocumentProperties().getProperty('FIRST_RUN_COMPLETE');
  
  if (!isFirstRun) {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      '👋 ようこそ！',
      'タスク管理システムへようこそ！\n\n' +
      '初回セットアップを実行しますか？',
      ui.ButtonSet.YES_NO
    );
    
    if (response === ui.Button.YES) {
      setupSystem();
    }
    
    PropertiesService.getDocumentProperties().setProperty('FIRST_RUN_COMPLETE', 'true');
  }
}

// プログレス表示用HTML
function getProgressHtml() {
  return `
    <!DOCTYPE html>
    <html>
      <head>
        <style>
          body {
            font-family: Arial, sans-serif;
            text-align: center;
            padding: 20px;
          }
          .spinner {
            border: 4px solid #f3f3f3;
            border-top: 4px solid #3498db;
            border-radius: 50%;
            width: 50px;
            height: 50px;
            animation: spin 1s linear infinite;
            margin: 20px auto;
          }
          @keyframes spin {
            0% { transform: rotate(0deg); }
            100% { transform: rotate(360deg); }
          }
        </style>
      </head>
      <body>
        <h3>セットアップ中...</h3>
        <div class="spinner"></div>
        <p>しばらくお待ちください</p>
      </body>
    </html>
  `;
}

// 新データエンジンのテスト
function testNewDataEngine() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    // テストデータを作成
    const testFlowRows = [
      {
        '工程': 'テスト工程1',
        '実施タイミング': '第1期',
        '部署': 'テスト部署',
        '担当役割': 'テスト担当',
        '作業内容': 'テスト作業内容',
        '条件分岐': 'なし',
        '利用ツール': 'テストツール',
        'URLリンク': '',
        '備考': 'テスト備考3'  // 末尾に数字を含むテストケース
      }
    ];
    
    console.log('新データエンジンテスト開始');
    
    // 新しいエンジンをテスト
    if (typeof writeFlowRowsImproved === 'function') {
      writeFlowRowsImproved(testFlowRows);
      ui.alert('✅ テスト成功', '新しいデータ処理エンジンが正常に動作しています。', ui.ButtonSet.OK);
    } else {
      ui.alert('⚠️ 警告', '新しいデータ処理エンジンが見つかりません。レガシー関数を使用します。', ui.ButtonSet.OK);
      writeFlowRows(testFlowRows);
    }
    
  } catch (error) {
    console.error('新データエンジンテストエラー:', error);
    ui.alert('❌ テスト失敗', `エラー: ${error.message}`, ui.ButtonSet.OK);
  }
}