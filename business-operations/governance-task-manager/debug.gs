// ================================================================================
// デバッグ機能
// ================================================================================

// デバッグモードの設定
const DEBUG_MODE = true; // デバッグモードのON/OFF

// デバッグログ出力
function debugLog(functionName, message, data = null) {
  if (!DEBUG_MODE) return;
  
  const timestamp = new Date().toISOString();
  const logMessage = `[${timestamp}] [${functionName}] ${message}`;
  
  console.log(logMessage);
  
  if (data !== null) {
    console.log('Data:', JSON.stringify(data, null, 2));
  }
  
  // デバッグシートにも記録
  logToDebugSheet(functionName, message, data);
}

// デバッグシートへの記録
function logToDebugSheet(functionName, message, data) {
  try {
    const debugSheetName = 'DebugLog';
    let debugSheet = ss().getSheetByName(debugSheetName);
    
    if (!debugSheet) {
      debugSheet = createDebugSheet();
    }
    
    // 古いログを削除（1000行を超えたら古いものから削除）
    if (debugSheet.getLastRow() > 1000) {
      debugSheet.deleteRows(2, 100);
    }
    
    debugSheet.appendRow([
      new Date(),
      functionName,
      message,
      data ? JSON.stringify(data, null, 2) : '',
      Session.getActiveUser().getEmail()
    ]);
  } catch (e) {
    console.error('Failed to log to debug sheet:', e);
  }
}

// デバッグシート作成
function createDebugSheet() {
  const sh = ss().insertSheet('DebugLog');
  sh.getRange(1, 1, 1, 5).setValues([[
    'タイムスタンプ', '関数名', 'メッセージ', 'データ', 'ユーザー'
  ]]);
  sh.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#ffeb3b');
  sh.setFrozenRows(1);
  
  // 列幅調整
  sh.setColumnWidth(1, 150); // タイムスタンプ
  sh.setColumnWidth(2, 150); // 関数名
  sh.setColumnWidth(3, 300); // メッセージ
  sh.setColumnWidth(4, 400); // データ
  sh.setColumnWidth(5, 150); // ユーザー
  
  return sh;
}

// ================================================================================
// テスト機能
// ================================================================================

// テストメニュー追加（onOpenに追加）
function addDebugMenu() {
  const ui = SpreadsheetApp.getUi();
  
  ui.createMenu('🔧 デバッグ')
    .addItem('📝 デバッグログを表示', 'showDebugLog')
    .addItem('🗑️ デバッグログをクリア', 'clearDebugLog')
    .addSeparator()
    .addItem('🧪 接続テスト', 'testConnections')
    .addItem('📧 メール送信テスト', 'testEmailSend')
    .addItem('🤖 OpenAI APIテスト', 'testOpenAI')
    .addItem('📊 サンプルデータでフロー生成テスト', 'testFlowGeneration')
    .addSeparator()
    .addItem('🔍 現在の設定を表示', 'showCurrentConfig')
    .addItem('⚠️ エラーシミュレーション', 'simulateError')
    .addToUi();
}

// デバッグログ表示
function showDebugLog() {
  const sheet = ss().getSheetByName('DebugLog');
  if (sheet) {
    sheet.showSheet();
    ss().setActiveSheet(sheet);
  } else {
    SpreadsheetApp.getUi().alert('デバッグログがありません。');
  }
}

// デバッグログクリア
function clearDebugLog() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    '確認',
    'デバッグログをすべて削除しますか？',
    ui.ButtonSet.YES_NO
  );
  
  if (response === ui.Button.YES) {
    const sheet = ss().getSheetByName('DebugLog');
    if (sheet) {
      const lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        sheet.deleteRows(2, lastRow - 1);
      }
      ui.alert('デバッグログをクリアしました。');
    }
  }
}

// 接続テスト
function testConnections() {
  const ui = SpreadsheetApp.getUi();
  const results = [];
  
  debugLog('testConnections', 'Starting connection tests');
  
  // 1. スプレッドシート接続
  try {
    const testSheet = ss();
    results.push('✅ スプレッドシート接続: OK');
    debugLog('testConnections', 'Spreadsheet connection successful');
  } catch (e) {
    results.push('❌ スプレッドシート接続: ' + e.toString());
    debugLog('testConnections', 'Spreadsheet connection failed', e.toString());
  }
  
  // 2. Drive API接続
  try {
    const testFile = file();
    results.push('✅ Drive API接続: OK');
    debugLog('testConnections', 'Drive API connection successful');
  } catch (e) {
    results.push('❌ Drive API接続: ' + e.toString());
    debugLog('testConnections', 'Drive API connection failed', e.toString());
  }
  
  // 3. Gmail API接続
  try {
    const testLabel = GmailApp.getUserLabelByName('TEST_LABEL_DELETE_ME');
    results.push('✅ Gmail API接続: OK');
    debugLog('testConnections', 'Gmail API connection successful');
  } catch (e) {
    results.push('❌ Gmail API接続: ' + e.toString());
    debugLog('testConnections', 'Gmail API connection failed', e.toString());
  }
  
  // 4. Properties Service接続
  try {
    PropertiesService.getScriptProperties().getProperty('TEST_PROP');
    results.push('✅ Properties Service接続: OK');
    debugLog('testConnections', 'Properties Service connection successful');
  } catch (e) {
    results.push('❌ Properties Service接続: ' + e.toString());
    debugLog('testConnections', 'Properties Service connection failed', e.toString());
  }
  
  // 5. OpenAI APIキー確認
  try {
    const apiKey = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
    if (apiKey) {
      results.push('✅ OpenAI APIキー: 設定済み');
      debugLog('testConnections', 'OpenAI API key is set');
    } else {
      results.push('⚠️ OpenAI APIキー: 未設定');
      debugLog('testConnections', 'OpenAI API key is not set');
    }
  } catch (e) {
    results.push('❌ OpenAI APIキー確認: ' + e.toString());
    debugLog('testConnections', 'OpenAI API key check failed', e.toString());
  }
  
  ui.alert('接続テスト結果', results.join('\n'), ui.ButtonSet.OK);
}

// メール送信テスト
function testEmailSend() {
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.prompt(
    'メール送信テスト',
    'テストメールを送信する宛先を入力してください：',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() === ui.Button.OK) {
    const testEmail = response.getResponseText().trim();
    
    if (!testEmail) {
      ui.alert('メールアドレスが入力されていません。');
      return;
    }
    
    debugLog('testEmailSend', `Sending test email to: ${testEmail}`);
    
    try {
      GmailApp.sendEmail(
        testEmail,
        '[TEST] タスク管理システム - テストメール',
        'これはテストメールです。\n\nシステムが正常に動作しています。',
        {
          htmlBody: `
            <div style="font-family: Arial, sans-serif; padding: 20px;">
              <h2>テストメール</h2>
              <p>これはタスク管理システムからのテストメールです。</p>
              <p>システムが正常に動作しています。</p>
              <hr>
              <p style="color: #666; font-size: 12px;">
                送信日時: ${new Date().toLocaleString('ja-JP')}
              </p>
            </div>
          `,
          name: 'タスク管理システム（テスト）'
        }
      );
      
      ui.alert('✅ テストメールを送信しました。');
      debugLog('testEmailSend', 'Test email sent successfully');
    } catch (e) {
      ui.alert('❌ エラー', 'メール送信に失敗しました：\n' + e.toString(), ui.ButtonSet.OK);
      debugLog('testEmailSend', 'Test email failed', e.toString());
    }
  }
}

// OpenAI APIテスト
function testOpenAI() {
  const ui = SpreadsheetApp.getUi();
  
  debugLog('testOpenAI', 'Starting OpenAI API test');
  
  try {
    const apiKey = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
    
    if (!apiKey) {
      ui.alert('⚠️ OpenAI APIキーが設定されていません。');
      debugLog('testOpenAI', 'API key not set');
      return;
    }
    
    // シンプルなテストリクエスト
    const testPayload = {
      model: 'gpt-3.5-turbo',
      messages: [
        { role: 'system', content: 'You are a helpful assistant.' },
        { role: 'user', content: 'Reply with "OK" if you receive this message.' }
      ],
      max_tokens: 10
    };
    
    debugLog('testOpenAI', 'Sending test request', testPayload);
    
    const response = UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions', {
      method: 'post',
      headers: {
        'Authorization': `Bearer ${apiKey}`,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(testPayload),
      muteHttpExceptions: true
    });
    
    const status = response.getResponseCode();
    const responseData = JSON.parse(response.getContentText());
    
    debugLog('testOpenAI', `Response status: ${status}`, responseData);
    
    if (status === 200) {
      const reply = responseData.choices[0].message.content;
      ui.alert('✅ OpenAI API接続成功', `応答: ${reply}`, ui.ButtonSet.OK);
    } else {
      ui.alert('❌ OpenAI APIエラー', `ステータス: ${status}\n${responseData.error?.message || 'Unknown error'}`, ui.ButtonSet.OK);
    }
    
  } catch (e) {
    ui.alert('❌ エラー', 'OpenAI APIテストに失敗しました：\n' + e.toString(), ui.ButtonSet.OK);
    debugLog('testOpenAI', 'Test failed', e.toString());
  }
}

// フロー生成テスト
function testFlowGeneration() {
  const ui = SpreadsheetApp.getUi();
  
  debugLog('testFlowGeneration', 'Starting flow generation test');
  
  try {
    // サンプルデータ作成
    createSampleFlowData();
    
    // ビジュアルフロー生成
    generateVisualFlow();
    
    ui.alert('✅ フロー生成テスト完了', 'サンプルデータでフローを生成しました。', ui.ButtonSet.OK);
    debugLog('testFlowGeneration', 'Flow generation test completed');
    
  } catch (e) {
    ui.alert('❌ エラー', 'フロー生成テストに失敗しました：\n' + e.toString(), ui.ButtonSet.OK);
    debugLog('testFlowGeneration', 'Flow generation test failed', e.toString());
  }
}

// 現在の設定表示
function showCurrentConfig() {
  const ui = SpreadsheetApp.getUi();
  const configs = [];
  
  debugLog('showCurrentConfig', 'Retrieving current configuration');
  
  // Config シートから設定を取得
  const configSheet = ss().getSheetByName(CONFIG_SHEET);
  if (configSheet && configSheet.getLastRow() > 0) {
    const values = configSheet.getRange(1, 1, configSheet.getLastRow(), 2).getValues();
    
    configs.push('【Config設定】');
    values.forEach(row => {
      if (row[0]) {
        // APIキーなどは一部マスク
        let value = String(row[1]);
        if (row[0].toString().includes('API_KEY') && value) {
          value = value.substring(0, 10) + '...' + '(masked)';
        }
        configs.push(`${row[0]}: ${value}`);
      }
    });
  }
  
  // トリガー情報
  configs.push('\n【トリガー設定】');
  const triggers = ScriptApp.getProjectTriggers();
  if (triggers.length > 0) {
    triggers.forEach(trigger => {
      configs.push(`- ${trigger.getHandlerFunction()} (${trigger.getEventType()})`);
    });
  } else {
    configs.push('トリガーなし');
  }
  
  // シート情報
  configs.push('\n【シート情報】');
  const sheets = ss().getSheets();
  sheets.forEach(sheet => {
    configs.push(`- ${sheet.getName()} (${sheet.getLastRow()}行)`);
  });
  
  const message = configs.join('\n');
  
  // ダイアログで表示
  const html = HtmlService.createHtmlOutput(`<pre>${message}</pre>`)
    .setWidth(600)
    .setHeight(400);
  ui.showModalDialog(html, '現在の設定');
  
  debugLog('showCurrentConfig', 'Configuration displayed');
}

// エラーシミュレーション
function simulateError() {
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.alert(
    'エラーシミュレーション',
    'テスト用のエラーを発生させます。続行しますか？',
    ui.ButtonSet.YES_NO
  );
  
  if (response === ui.Button.YES) {
    debugLog('simulateError', 'Simulating error');
    
    try {
      // 意図的にエラーを発生させる
      throw new Error('これはテスト用のエラーです。システムは正常です。');
    } catch (e) {
      // エラーログに記録
      logError('TEST_MESSAGE_ID', e);
      logActivity('ERROR_SIMULATION', e.toString());
      debugLog('simulateError', 'Error simulated', e.toString());
      
      ui.alert('エラーシミュレーション完了', 'エラーログとアクティビティログに記録しました。', ui.ButtonSet.OK);
    }
  }
}

// ================================================================================
// デバッグ用ラッパー関数（既存の関数をラップ）
// ================================================================================

// processNewEmailsのデバッグラッパー
function processNewEmailsDebug() {
  debugLog('processNewEmailsDebug', 'Starting email processing with debug');
  
  try {
    const query = getConfig('PROCESSING_QUERY') || 'label:inbox is:unread';
    debugLog('processNewEmailsDebug', `Query: ${query}`);
    
    const threads = GmailApp.search(query);
    debugLog('processNewEmailsDebug', `Found ${threads.length} threads`);
    
    if (threads.length === 0) {
      debugLog('processNewEmailsDebug', 'No new emails found');
      return;
    }
    
    threads.forEach((thread, index) => {
      debugLog('processNewEmailsDebug', `Processing thread ${index + 1}/${threads.length}`);
      processThread(thread);
    });
    
    debugLog('processNewEmailsDebug', 'Email processing completed');
  } catch (e) {
    debugLog('processNewEmailsDebug', 'Email processing failed', e.toString());
    throw e;
  }
}

// OpenAI呼び出しのデバッグラッパー
function callOpenAIDebug(mailBody, orgProfileJson) {
  debugLog('callOpenAIDebug', 'Calling OpenAI API', {
    bodyLength: mailBody.length,
    orgProfile: orgProfileJson
  });
  
  try {
    const result = callOpenAI(mailBody, orgProfileJson);
    debugLog('callOpenAIDebug', 'OpenAI API call successful', {
      title: result.work_spec?.title,
      flowRowsCount: result.flow_rows?.length
    });
    return result;
  } catch (e) {
    debugLog('callOpenAIDebug', 'OpenAI API call failed', e.toString());
    throw e;
  }
}

// ================================================================================
// 初期化時にデバッグメニューを追加
// ================================================================================

// onOpenの拡張（既存のonOpenに追加）
function onOpenWithDebug() {
  onOpen(); // 既存のメニューを追加
  
  if (DEBUG_MODE) {
    addDebugMenu(); // デバッグメニューを追加
    debugLog('onOpenWithDebug', 'Debug menu added');
  }
}