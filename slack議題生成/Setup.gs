// ==========================================
// セットアップ関数
// ==========================================

/**
 * 初回セットアップを実行
 * この関数を最初に実行してください
 */
function runInitialSetup() {
  const ui = SpreadsheetApp.getUi();
  
  ui.alert(
    '🎯 セットアップを開始します',
    'これから以下の設定を行います：\n\n' +
    '1. スプレッドシートの初期化\n' +
    '2. 必要なシートの作成\n' +
    '3. API認証情報の設定\n' +
    '4. 初期データの投入\n\n' +
    '続行しますか？',
    ui.ButtonSet.OK_CANCEL
  );
  
  try {
    // Step 1: スプレッドシートIDを保存
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const spreadsheetId = spreadsheet.getId();
    PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', spreadsheetId);
    
    // Step 2: 必要なシートを作成
    initializeAllSheets();
    
    // Step 3: 設定ダイアログを表示
    showSetupWizard();
    
    ui.alert(
      '✅ セットアップ準備完了',
      'スプレッドシートの初期化が完了しました。\n\n' +
      '次に表示される設定画面で、Slack APIとOpenAI APIの情報を入力してください。',
      ui.ButtonSet.OK
    );
    
  } catch (error) {
    ui.alert(
      '❌ エラー',
      'セットアップ中にエラーが発生しました：\n' + error.toString(),
      ui.ButtonSet.OK
    );
    throw error;
  }
}

/**
 * 全シートを初期化
 */
function initializeAllSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 既存のシートをクリア（最初のシートは残す）
  const existingSheets = ss.getSheets();
  for (let i = existingSheets.length - 1; i > 0; i--) {
    ss.deleteSheet(existingSheets[i]);
  }
  
  // 各シートを作成
  createConfigSheet(ss);
  createSyncStateSheet(ss);
  createMessagesSheet(ss);
  createCategoriesSheet(ss);
  createChecklistsSheet(ss);
  createTemplatesSheet(ss);
  createDraftsSheet(ss);
  createLogsSheet(ss);
  createSlackLogSheet(ss);
  createBusinessManualSheet(ss);
  createFAQListSheet(ss);
  createDailyReportSheet(ss);
  
  // 最初のデフォルトシートを削除
  try {
    ss.deleteSheet(existingSheets[0]);
  } catch (e) {
    // 削除できない場合は無視
  }
}

/**
 * セットアップウィザードを表示
 */
function showSetupWizard() {
  const html = HtmlService.createHtmlOutputFromFile('setup')
    .setWidth(700)
    .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, '🔧 API設定ウィザード');
}

/**
 * Slack接続テスト
 */
function testSlackConnection() {
  const token = PropertiesService.getScriptProperties().getProperty('SLACK_BOT_TOKEN');
  
  if (!token) {
    return { success: false, message: 'Slack Bot Tokenが設定されていません' };
  }
  
  try {
    const url = 'https://slack.com/api/auth.test';
    const response = UrlFetchApp.fetch(url, {
      method: 'post',
      headers: {
        'Authorization': `Bearer ${token}`,
        'Content-Type': 'application/json'
      },
      muteHttpExceptions: true
    });
    
    const data = JSON.parse(response.getContentText());
    
    if (data.ok) {
      return {
        success: true,
        message: `✅ 接続成功！\nボット名: ${data.user}\nチーム: ${data.team}`
      };
    } else {
      return {
        success: false,
        message: `❌ 接続失敗: ${data.error}`
      };
    }
  } catch (error) {
    return {
      success: false,
      message: `❌ エラー: ${error.toString()}`
    };
  }
}

/**
 * OpenAI接続テスト
 */
function testOpenAIConnection() {
  const apiKey = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
  
  if (!apiKey) {
    return { success: false, message: 'OpenAI API Keyが設定されていません' };
  }
  
  try {
    const url = 'https://api.openai.com/v1/responses';
    const response = UrlFetchApp.fetch(url, {
      method: 'post',
      headers: {
        'Authorization': `Bearer ${apiKey}`,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify({
        model: 'gpt-5',
        input: 'SYSTEM:\nYou are a helpful assistant.\n\nUSER:\nhealth check',
        max_output_tokens: 5
      }),
      muteHttpExceptions: true
    });
    
    const data = JSON.parse(response.getContentText());
    const ok = (typeof data.output_text === 'string') || (Array.isArray(data.output)) || data.choices;
    if (ok) {
      return {
        success: true,
        message: '✅ OpenAI API接続成功！'
      };
    } else if (data.error) {
      return {
        success: false,
        message: `❌ 接続失敗: ${data.error.message}`
      };
    }
  } catch (error) {
    return {
      success: false,
      message: `❌ エラー: ${error.toString()}`
    };
  }
}

/**
 * 設定を保存して検証
 */
function saveAndValidateSettings(settings) {
  const results = {
    saved: false,
    slackTest: null,
    openaiTest: null,
    errors: []
  };
  
  try {
    // 設定を保存
    const scriptProperties = PropertiesService.getScriptProperties();
    
    if (settings.SLACK_BOT_TOKEN) {
      scriptProperties.setProperty('SLACK_BOT_TOKEN', settings.SLACK_BOT_TOKEN);
    }
    if (settings.OPENAI_API_KEY) {
      scriptProperties.setProperty('OPENAI_API_KEY', settings.OPENAI_API_KEY);
    }
    if (settings.REPORT_EMAIL) {
      scriptProperties.setProperty('REPORT_EMAIL', settings.REPORT_EMAIL);
    }
    if (settings.SLACK_SIGNING_SECRET) {
      scriptProperties.setProperty('SLACK_SIGNING_SECRET', settings.SLACK_SIGNING_SECRET);
    }
    
    results.saved = true;
    
    // 接続テスト
    if (settings.SLACK_BOT_TOKEN) {
      results.slackTest = testSlackConnection();
    }
    if (settings.OPENAI_API_KEY) {
      results.openaiTest = testOpenAIConnection();
    }
    
  } catch (error) {
    results.errors.push(error.toString());
  }
  
  return results;
}

/**
 * サンプルデータを投入
 */
function insertSampleData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  try {
    // Categoriesシートにサンプルカテゴリを追加
    const categoriesSheet = ss.getSheetByName('Categories');
    if (categoriesSheet && categoriesSheet.getLastRow() === 1) {
      const sampleCategories = [
        ['開示事項', '適時開示が必要な事項', '金融商品取引法、取引所規則に基づく開示が必要な事項', '決算,業績予想,配当,買収,合併,提携,新株発行', '高'],
        ['取締役会決議事項', '取締役会での決議が必要な事項', '会社法および定款で定められた取締役会決議事項', '重要な財産,借入,投資,組織変更,人事,規程改定', '高'],
        ['監査等委員会決議事項', '監査等委員会での決議が必要な事項', '監査等委員会の職務に関する事項', '監査計画,監査報告,会計監査人,内部統制', '中'],
        ['株主総会決議事項', '株主総会での決議が必要な事項', '会社法および定款で定められた株主総会決議事項', '定款変更,取締役選任,剰余金配当,資本政策', '高'],
        ['プロジェクト進捗', 'プロジェクトの進捗報告', '各プロジェクトの状況報告と課題共有', 'プロジェクト,進捗,課題,リスク,スケジュール', '中'],
        ['緊急対応事項', '緊急対応が必要な事項', '即座の判断や対応が必要な事項', '緊急,至急,ASAP,本日中,重要', '高']
      ];
      
      categoriesSheet.getRange(2, 1, sampleCategories.length, 5).setValues(sampleCategories);
    }
    
    // Checklistsシートにサンプルチェックリストを追加
    const checklistsSheet = ss.getSheetByName('Checklists');
    if (checklistsSheet && checklistsSheet.getLastRow() === 1) {
      const sampleChecklists = [
        ['取締役会決議事項', '決議事項の明確化', '議案の内容を具体的に記載', '必須', '第○号議案：○○について'],
        ['取締役会決議事項', '出席者の確認', '定足数の充足を確認', '必須', '出席取締役○名（定足数○名）'],
        ['取締役会決議事項', '決議結果', '賛成・反対・棄権の記録', '必須', '全員一致で承認/賛成○名、反対○名'],
        ['開示事項', '開示時期', '適時開示のタイミング', '必須', '決議後速やかに開示'],
        ['開示事項', '開示内容', '開示する情報の範囲', '必須', '決議内容、理由、今後の見通し'],
        ['監査等委員会決議事項', '監査計画', '年度監査計画の策定', '必須', '○○年度監査計画について'],
        ['株主総会決議事項', '招集通知', '株主総会招集通知の発送', '必須', '総会の○週間前までに発送'],
        ['プロジェクト進捗', '進捗率', 'プロジェクトの進捗状況', '必須', '全体の○○%完了'],
        ['プロジェクト進捗', '課題', '現在の課題と対策', '任意', '課題：○○、対策：○○'],
        ['緊急対応事項', '影響範囲', '影響を受ける範囲の特定', '必須', '影響：○○部門、○○システム'],
        ['緊急対応事項', '対応策', '具体的な対応方法', '必須', '即座に○○を実施']
      ];
      
      checklistsSheet.getRange(2, 1, sampleChecklists.length, 5).setValues(sampleChecklists);
    }
    
    // Configシートに初期設定を追加
    const configSheet = ss.getSheetByName('Config');
    if (configSheet) {
      const configData = [
        ['company', '', '会社名を入力してください'],
        ['targetChannels', '', 'Slack監視対象チャンネルID（カンマ区切り）'],
        ['notifySlackChannel', '', '通知先SlackチャンネルID'],
        ['notifyEmails', PropertiesService.getScriptProperties().getProperty('REPORT_EMAIL') || '', '通知先メールアドレス（カンマ区切り）'],
        ['openaiModel', 'gpt-5', 'OpenAIモデル名（要約・分類用）'],
        ['openaiModelDraft', 'gpt-5', 'OpenAIモデル名（ドラフト生成用）'],
        ['classificationThreshold', '0.6', '該当判定しきい値（0-1）'],
        ['syncIntervalMinutes', '30', 'Slack同期間隔（分）'],
        ['analysisIntervalHours', '1', 'AI分析実行間隔（時間）'],
        ['notificationHours', '9,15', '通知時刻（カンマ区切り）']
      ];
      
      // 既存のヘッダー行をスキップして設定を更新
      configSheet.getRange(2, 1, configData.length, 3).setValues(configData);
    }
    
    ui.alert('✅ サンプルデータの投入完了', 
             'カテゴリとチェックリストのサンプルデータを追加しました。\n必要に応じて編集してください。', 
             ui.ButtonSet.OK);
    
  } catch (error) {
    ui.alert('❌ エラー', 'サンプルデータの投入中にエラーが発生しました：\n' + error.toString(), ui.ButtonSet.OK);
  }
}

/**
 * 権限チェック関数
 */
function checkRequiredPermissions() {
  const results = {
    spreadsheet: false,
    urlFetch: false,
    mail: false,
    script: false,
    errors: []
  };
  
  try {
    // スプレッドシートアクセス権限
    SpreadsheetApp.getActiveSpreadsheet();
    results.spreadsheet = true;
  } catch (e) {
    results.errors.push('スプレッドシート権限: ' + e.toString());
  }
  
  try {
    // 外部URL取得権限（ダミーリクエスト）
    UrlFetchApp.getRequest('https://www.google.com');
    results.urlFetch = true;
  } catch (e) {
    results.errors.push('外部API権限: ' + e.toString());
  }
  
  try {
    // メール送信権限
    MailApp.getRemainingDailyQuota();
    results.mail = true;
  } catch (e) {
    results.errors.push('メール送信権限: ' + e.toString());
  }
  
  try {
    // スクリプトプロパティ権限
    PropertiesService.getScriptProperties();
    results.script = true;
  } catch (e) {
    results.errors.push('スクリプトプロパティ権限: ' + e.toString());
  }
  
  return results;
}

// ========= 個別シート作成関数 =========

function createConfigSheet(ss) {
  const sheet = ss.insertSheet('Config');
  const headers = ['設定項目', '値', '説明'];
  sheet.getRange(1, 1, 1, 3).setValues([headers]);
  sheet.getRange(1, 1, 1, 3)
    .setBackground('#4285F4')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold');
  sheet.setColumnWidth(1, 200);
  sheet.setColumnWidth(2, 300);
  sheet.setColumnWidth(3, 400);
  sheet.setFrozenRows(1);
}

function createSyncStateSheet(ss) {
  const sheet = ss.insertSheet('SyncState');
  const headers = ['channel_id', 'last_sync_ts', 'last_sync_datetime', 'message_count', 'status'];
  sheet.getRange(1, 1, 1, 5).setValues([headers]);
  sheet.getRange(1, 1, 1, 5)
    .setBackground('#34A853')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold');
  sheet.setFrozenRows(1);
}

function createMessagesSheet(ss) {
  const sheet = ss.insertSheet('Messages');
  const headers = [
    'id', 'channel_id', 'message_ts', 'thread_ts', 'text_raw',
    'user_name', 'summary_json', 'classification_json', 'match_flag',
    'human_judgement', 'permalink', 'checklist_proposed', 'agenda_selected',
    'draft_doc_url', 'created_at', 'updated_at'
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setBackground('#EA4335')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold');
  
  // データ検証の設定
  const validationRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['必要', '不要', '保留'], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, 10, 1000, 1).setDataValidation(validationRule);
  
  // チェックボックスの設定
  sheet.getRange(2, 9, 1000, 1).insertCheckboxes();
  
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(1);
}

function createCategoriesSheet(ss) {
  const sheet = ss.insertSheet('Categories');
  const headers = ['カテゴリ名', '説明', '判定基準', 'キーワード（カンマ区切り）', '重要度'];
  sheet.getRange(1, 1, 1, 5).setValues([headers]);
  sheet.getRange(1, 1, 1, 5)
    .setBackground('#FBBC04')
    .setFontColor('#000000')
    .setFontWeight('bold');
  
  // データ検証の設定
  const validationRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['高', '中', '低'], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, 5, 100, 1).setDataValidation(validationRule);
  
  sheet.setFrozenRows(1);
}

function createChecklistsSheet(ss) {
  const sheet = ss.insertSheet('Checklists');
  const headers = ['カテゴリ', 'チェック項目', '説明', '必須/任意', 'テンプレート文'];
  sheet.getRange(1, 1, 1, 5).setValues([headers]);
  sheet.getRange(1, 1, 1, 5)
    .setBackground('#9333EA')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold');
  
  // データ検証の設定
  const validationRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['必須', '任意'], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, 4, 100, 1).setDataValidation(validationRule);
  
  sheet.setFrozenRows(1);
}

function createTemplatesSheet(ss) {
  const sheet = ss.insertSheet('Templates');
  const headers = ['テンプレート名', 'カテゴリ', 'Google Doc ID', 'プレースホルダー', '最終更新日'];
  sheet.getRange(1, 1, 1, 5).setValues([headers]);
  sheet.getRange(1, 1, 1, 5)
    .setBackground('#00ACC1')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold');
  sheet.setFrozenRows(1);
}

function createDraftsSheet(ss) {
  const sheet = ss.insertSheet('Drafts');
  const headers = [
    'message_id', 'category', 'doc_url', 'created_at', 'created_by',
    'status', 'reviewed_by', 'reviewed_at', 'notes'
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setBackground('#FF6D00')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold');
  
  // データ検証の設定
  const validationRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['下書き', 'レビュー中', '承認済み', '却下'], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, 6, 100, 1).setDataValidation(validationRule);
  
  sheet.setFrozenRows(1);
}

function createLogsSheet(ss) {
  const sheet = ss.insertSheet('Logs');
  const headers = ['timestamp', 'level', 'message', 'details'];
  sheet.getRange(1, 1, 1, 4).setValues([headers]);
  sheet.getRange(1, 1, 1, 4)
    .setBackground('#616161')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold');
  
  // 条件付き書式（エラーレベルを赤色に）
  const range = sheet.getRange(2, 2, 1000, 1);
  const rule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('ERROR')
    .setBackground('#FFCDD2')
    .setFontColor('#B71C1C')
    .setRanges([range])
    .build();
  
  const rules = sheet.getConditionalFormatRules();
  rules.push(rule);
  sheet.setConditionalFormatRules(rules);
  
  sheet.setFrozenRows(1);
}

function createSlackLogSheet(ss) {
  const sheet = ss.insertSheet('slack_log');
  const headers = [
    'channel_id', 'channel_name', 'ts', 'thread_ts', 
    'user_name', 'message', 'date', 'reactions', 'files'
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setBackground('#4A154B')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold');
  sheet.setFrozenRows(1);
}

function createBusinessManualSheet(ss) {
  const sheet = ss.insertSheet('business_manual');
  const headers = [
    '作成日時', 'カテゴリ', 'タイトル', '内容', 
    '元のチャンネル', '関連メッセージ', 'ステータス',
    '参加者', 'キーワード', '重要度'
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setBackground('#0F9D58')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold');
  sheet.setFrozenRows(1);
}

function createFAQListSheet(ss) {
  const sheet = ss.insertSheet('faq_list');
  const headers = [
    '作成日時', '質問', '回答', 'カテゴリ', 'タグ', 
    '元のチャンネル', '関連メッセージ', 'ステータス'
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setBackground('#1A73E8')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold');
  sheet.setFrozenRows(1);
}

function createDailyReportSheet(ss) {
  const sheet = ss.insertSheet('daily_report');
  const headers = ['日付', 'タイトル', 'レポート内容', '生成時刻'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setBackground('#673AB7')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold');
  sheet.setFrozenRows(1);
}
