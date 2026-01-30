/**
 * スプレッドシート初期設定スクリプト
 * 
 * このファイルを単独で実行して、スプレッドシートを設定します
 */

/**
 * ステップ1: スプレッドシートIDを手動で設定
 * 
 * 1. Google Driveで新しいスプレッドシートを作成
 * 2. URLからIDをコピー（/d/と/editの間の文字列）
 * 3. 下記のSPREADSHEET_IDに貼り付け
 */
function setupStep1_SetSpreadsheetId() {
  // ここにスプレッドシートIDを入力してください
  const SPREADSHEET_ID = 'YOUR_SPREADSHEET_ID_HERE';
  
  if (SPREADSHEET_ID === 'YOUR_SPREADSHEET_ID_HERE') {
    console.log('❌ エラー: SPREADSHEET_IDを設定してください');
    console.log('\n手順:');
    console.log('1. Google Driveで新しいスプレッドシートを作成');
    console.log('2. スプレッドシートを開く');
    console.log('3. URLから以下の部分をコピー:');
    console.log('   https://docs.google.com/spreadsheets/d/【ここの部分】/edit');
    console.log('4. コピーしたIDを上記のSPREADSHEET_IDに貼り付け');
    console.log('5. この関数を再度実行');
    return;
  }
  
  // スクリプトプロパティに保存
  PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', SPREADSHEET_ID);
  console.log('✅ スプレッドシートIDを保存しました: ' + SPREADSHEET_ID);
  
  // スプレッドシートを開いてみる
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    console.log('✅ スプレッドシートに接続成功');
    console.log('   名前: ' + ss.getName());
    console.log('   URL: ' + ss.getUrl());
    console.log('\n次のステップ: setupStep2_InitializeSheets() を実行してください');
  } catch (e) {
    console.log('❌ スプレッドシートを開けません: ' + e.toString());
    console.log('IDが正しいか、アクセス権限があるか確認してください');
  }
}

/**
 * ステップ2: シートを初期化
 */
function setupStep2_InitializeSheets() {
  const SPREADSHEET_ID = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  
  if (!SPREADSHEET_ID) {
    console.log('❌ エラー: 先にsetupStep1_SetSpreadsheetId()を実行してください');
    return;
  }
  
  try {
    console.log('スプレッドシートを開いています...');
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    console.log('✅ スプレッドシート接続成功: ' + ss.getName());
    
    // 既存のシートを確認
    const sheets = ss.getSheets();
    console.log('\n現在のシート数: ' + sheets.length);
    
    // logシートの作成または確認
    console.log('\n1. logシートを設定中...');
    let logSheet = ss.getSheetByName('log');
    if (!logSheet) {
      if (sheets.length > 0 && sheets[0].getName() === 'シート1') {
        // デフォルトシートをlogに変更
        sheets[0].setName('log');
        logSheet = sheets[0];
        console.log('   デフォルトシートをlogに変更');
      } else {
        logSheet = ss.insertSheet('log');
        console.log('   logシートを作成');
      }
    } else {
      console.log('   logシートは既に存在');
    }
    
    // ヘッダーを設定
    if (logSheet.getLastRow() === 0) {
      logSheet.appendRow(['Timestamp', 'Message']);
      logSheet.getRange('1:1').setFontWeight('bold');
      logSheet.setFrozenRows(1);
      console.log('   ヘッダーを追加');
    }
    
    // faqシートの作成
    console.log('\n2. faqシートを設定中...');
    let faqSheet = ss.getSheetByName('faq');
    if (!faqSheet) {
      faqSheet = ss.insertSheet('faq');
      faqSheet.appendRow(['キーワード', '回答', '検索フラグ', 'Drive検索結果']);
      faqSheet.getRange('1:1').setFontWeight('bold');
      faqSheet.setFrozenRows(1);
      // 列幅を調整
      faqSheet.setColumnWidth(1, 150);
      faqSheet.setColumnWidth(2, 400);
      faqSheet.setColumnWidth(3, 100);
      faqSheet.setColumnWidth(4, 400);
      console.log('   faqシートを作成');
    } else {
      console.log('   faqシートは既に存在');
    }
    
    // ドライブ一覧シートの作成
    console.log('\n3. ドライブ一覧シートを設定中...');
    let driveSheet = ss.getSheetByName('ドライブ一覧');
    if (!driveSheet) {
      driveSheet = ss.insertSheet('ドライブ一覧');
      driveSheet.appendRow(['フォルダID', 'フォルダ名', '説明']);
      driveSheet.getRange('1:1').setFontWeight('bold');
      driveSheet.setFrozenRows(1);
      // 列幅を調整
      driveSheet.setColumnWidth(1, 300);
      driveSheet.setColumnWidth(2, 200);
      driveSheet.setColumnWidth(3, 300);
      console.log('   ドライブ一覧シートを作成');
    } else {
      console.log('   ドライブ一覧シートは既に存在');
    }
    
    // debug_logシートの作成
    console.log('\n4. debug_logシートを設定中...');
    let debugSheet = ss.getSheetByName('debug_log');
    if (!debugSheet) {
      debugSheet = ss.insertSheet('debug_log');
      debugSheet.appendRow(['Timestamp', 'Category', 'Message', 'Data']);
      debugSheet.getRange('1:1').setFontWeight('bold');
      debugSheet.setFrozenRows(1);
      // 列幅を調整
      debugSheet.setColumnWidth(1, 150);
      debugSheet.setColumnWidth(2, 100);
      debugSheet.setColumnWidth(3, 300);
      debugSheet.setColumnWidth(4, 400);
      console.log('   debug_logシートを作成');
    } else {
      console.log('   debug_logシートは既に存在');
    }
    
    console.log('\n========================================');
    console.log('✅ スプレッドシートの初期化完了！');
    console.log('========================================');
    console.log('\nスプレッドシートURL:');
    console.log(ss.getUrl());
    console.log('\n次のステップ: setupStep3_SetAPIKeys() を実行してください');
    
  } catch (e) {
    console.log('❌ エラー: ' + e.toString());
    console.log('\nエラーの詳細:');
    console.log(e.stack);
  }
}

/**
 * ステップ3: APIキーを設定
 */
function setupStep3_SetAPIKeys() {
  console.log('========================================');
  console.log('APIキーの設定');
  console.log('========================================');
  
  // 現在の設定を確認
  const props = PropertiesService.getScriptProperties().getProperties();
  
  console.log('\n現在の設定:');
  console.log('✅ SPREADSHEET_ID: ' + (props.SPREADSHEET_ID ? '設定済み' : '未設定'));
  console.log((props.SLACK_TOKEN ? '✅' : '❌') + ' SLACK_TOKEN: ' + (props.SLACK_TOKEN ? '設定済み' : '未設定'));
  console.log((props.OPEN_AI_TOKEN ? '✅' : '❌') + ' OPEN_AI_TOKEN: ' + (props.OPEN_AI_TOKEN ? '設定済み' : '未設定'));
  console.log('   GEMINI_TOKEN: ' + (props.GEMINI_TOKEN ? '設定済み（オプション）' : '未設定（オプション）'));
  console.log('   GOOGLE_NL_API: ' + (props.GOOGLE_NL_API ? '設定済み（オプション）' : '未設定（オプション）'));
  
  if (!props.SLACK_TOKEN || !props.OPEN_AI_TOKEN) {
    console.log('\n⚠️ 必要なAPIキーが設定されていません');
    console.log('\n設定方法:');
    console.log('1. プロジェクトの設定 → スクリプト プロパティ');
    console.log('2. 「プロパティを追加」をクリック');
    console.log('3. 以下のプロパティを追加:');
    console.log('   - SLACK_TOKEN: Slack Bot Token');
    console.log('   - OPEN_AI_TOKEN: OpenAI API Key');
    console.log('4. 保存');
    console.log('5. この関数を再度実行');
  } else {
    console.log('\n✅ 必要なAPIキーはすべて設定されています');
    console.log('\n次のステップ: setupStep4_TestConnection() を実行してください');
  }
}

/**
 * ステップ4: 接続テスト
 */
function setupStep4_TestConnection() {
  console.log('========================================');
  console.log('接続テスト');
  console.log('========================================');
  
  // スプレッドシート接続テスト
  console.log('\n1. スプレッドシート接続テスト...');
  const ssId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (!ssId) {
    console.log('❌ SPREADSHEET_IDが設定されていません');
    return;
  }
  
  try {
    const ss = SpreadsheetApp.openById(ssId);
    console.log('✅ スプレッドシート接続成功');
    console.log('   ' + ss.getUrl());
    
    // テストデータを書き込み
    const debugSheet = ss.getSheetByName('debug_log');
    if (debugSheet) {
      debugSheet.appendRow([new Date(), 'Test', 'Connection test', 'Success']);
      console.log('✅ テストデータ書き込み成功');
    }
  } catch (e) {
    console.log('❌ スプレッドシート接続失敗: ' + e.toString());
    return;
  }
  
  // Slack接続テスト
  console.log('\n2. Slack API接続テスト...');
  const slackToken = PropertiesService.getScriptProperties().getProperty('SLACK_TOKEN');
  if (!slackToken) {
    console.log('⚠️ SLACK_TOKENが設定されていません（スキップ）');
  } else {
    try {
      const url = 'https://slack.com/api/auth.test';
      const response = UrlFetchApp.fetch(url, {
        method: 'post',
        payload: { token: slackToken },
        muteHttpExceptions: true
      });
      const data = JSON.parse(response.getContentText());
      
      if (data.ok) {
        console.log('✅ Slack接続成功');
        console.log('   ユーザー: ' + data.user);
        console.log('   チーム: ' + data.team);
      } else {
        console.log('❌ Slack接続失敗: ' + data.error);
      }
    } catch (e) {
      console.log('❌ Slackテストエラー: ' + e.toString());
    }
  }
  
  // OpenAI接続テスト
  console.log('\n3. OpenAI API接続テスト...');
  const openAIToken = PropertiesService.getScriptProperties().getProperty('OPEN_AI_TOKEN');
  if (!openAIToken) {
    console.log('⚠️ OPEN_AI_TOKENが設定されていません（スキップ）');
  } else {
    console.log('✅ OPEN_AI_TOKEN設定確認（実際の接続テストは省略）');
  }
  
  console.log('\n========================================');
  console.log('セットアップ完了！');
  console.log('========================================');
  console.log('\n次のステップ:');
  console.log('1. デプロイ → 新しいデプロイ');
  console.log('2. 種類: ウェブアプリ');
  console.log('3. Web App URLをコピー');
  console.log('4. Slack Appの設定でURLを登録');
}

/**
 * すべてを一度に実行（既にスプレッドシートがある場合）
 */
function quickSetupWithExistingSpreadsheet(spreadsheetId) {
  console.log('クイックセットアップを開始...\n');
  
  // スプレッドシートIDを設定
  PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', spreadsheetId);
  console.log('✅ スプレッドシートID設定: ' + spreadsheetId);
  
  // シートを初期化
  setupStep2_InitializeSheets();
  
  // APIキーを確認
  setupStep3_SetAPIKeys();
  
  // 接続テスト
  setupStep4_TestConnection();
}