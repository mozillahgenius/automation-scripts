/**
 * 標的型メール訓練システム - シート初期化スクリプト
 * スプレッドシートの初期構造を作成
 */

/**
 * スプレッドシートを初期化
 */
function initializeSpreadsheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 既存のシートを削除（初回実行時）
  const sheets = ss.getSheets();
  if (sheets.length === 1 && sheets[0].getName() === 'シート1') {
    sheets[0].setName('Config');
  }

  // 各シートを作成
  createConfigSheet(ss);
  createTargetsSheet(ss);
  createCampaignsSheet(ss);
  createMailsSheet(ss);
  createClicksSheet(ss);
  createResultsSheet(ss);

  // スクリプトプロパティの初期設定
  initializeScriptProperties();

  SpreadsheetApp.getUi().alert('スプレッドシートの初期化が完了しました。');
}

/**
 * Configシート作成
 */
function createConfigSheet(ss) {
  let sheet = ss.getSheetByName('Config');
  if (!sheet) {
    sheet = ss.insertSheet('Config');
  }

  sheet.clear();

  const headers = ['Key', 'Value', '備考'];
  const data = [
    ['PPLX_API_KEY', '', 'Perplexity APIキー（Script Propertiesでも設定可）'],
    ['COMPANY_URLS', 'https://example.com,https://example.com/news', '自社/関連/採用ページURL（カンマ区切り）'],
    ['SENDER_ALIAS', 'alerts@training.example.com', '許諾済みエイリアス'],
    ['SENDER_NAME_WEAK_SUS', 'exarnple HR', 'わずかな不審性（rn/m混同など）'],
    ['REPLY_TO', 'no-reply@training.example.com', 'Reply-To設定'],
    ['LANDING_PAGE_URL', '', 'GAS WebApp URL（デプロイ後に設定）'],
    ['EXPLAIN_LP_URL', 'https://intra.example.com/security/phishing-lesson', '教育ページURL'],
    ['RATE_LIMIT_PER_MIN', '80', '1分あたりの送信上限'],
    ['PIXEL_TRACKING', 'FALSE', '開封率トラッキング（TRUE/FALSE）'],
    ['TOKEN_LENGTH', '24', 'URLトークン長'],
    ['MAIL_DIFFICULTY_LEVELS', 'Easy,Std,Hard', '難易度レベル'],
    ['AUTO_EXPLAIN_ENABLED', 'TRUE', '自動解説メール送信（TRUE/FALSE）']
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#f0f0f0');
  sheet.getRange(2, 1, data.length, 3).setValues(data);
  sheet.autoResizeColumns(1, 3);

  // 条件付き書式（重要項目をハイライト）
  const range = sheet.getRange('A2:A13');
  const rule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('PPLX_API_KEY')
    .setBackground('#fff3cd')
    .setRanges([range])
    .build();
  sheet.setConditionalFormatRules([rule]);
}

/**
 * Targetsシート作成
 */
function createTargetsSheet(ss) {
  let sheet = ss.getSheetByName('Targets');
  if (!sheet) {
    sheet = ss.insertSheet('Targets');
  }

  sheet.clear();

  const headers = ['enabled', 'email', 'name', 'dept', 'title', 'uid', 'manager_email'];
  const sampleData = [
    ['TRUE', 'tanaka@example.com', '田中太郎', '営業部', '課長', 'U001', 'suzuki@example.com'],
    ['TRUE', 'yamada@example.com', '山田花子', '人事部', '主任', 'U002', 'sato@example.com'],
    ['FALSE', 'test@example.com', 'テストユーザー', 'IT部', 'エンジニア', 'U999', '']
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#f0f0f0');
  sheet.getRange(2, 1, sampleData.length, headers.length).setValues(sampleData);
  sheet.autoResizeColumns(1, headers.length);

  // データ検証（enabled列）
  const enabledRange = sheet.getRange('A2:A1000');
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['TRUE', 'FALSE'], true)
    .setAllowInvalid(false)
    .build();
  enabledRange.setDataValidation(rule);
}

/**
 * Campaignsシート作成
 */
function createCampaignsSheet(ss) {
  let sheet = ss.getSheetByName('Campaigns');
  if (!sheet) {
    sheet = ss.insertSheet('Campaigns');
  }

  sheet.clear();

  const headers = ['campaign_id', 'name', 'scheduled_at', 'difficulty', 'persona_hint', 'prompt_template', 'suspicious_flags', 'status'];
  const sampleData = [
    [
      'C001',
      '2025Q1_標的型訓練',
      '2025/01/15 10:00',
      'Easy',
      '新入社員向け',
      '会社の最新ニュースを基に、部署{dept}の{title}向けにメールを作成。緊急性を演出。',
      'homoglyph_from,time_pressure',
      'draft'
    ]
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#f0f0f0');
  sheet.getRange(2, 1, sampleData.length, headers.length).setValues(sampleData);
  sheet.autoResizeColumns(1, headers.length);

  // データ検証
  const difficultyRange = sheet.getRange('D2:D1000');
  const difficultyRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Easy', 'Std', 'Hard'], true)
    .build();
  difficultyRange.setDataValidation(difficultyRule);

  const statusRange = sheet.getRange('H2:H1000');
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['draft', 'ready', 'sent', 'completed'], true)
    .build();
  statusRange.setDataValidation(statusRule);
}

/**
 * Mailsシート作成
 */
function createMailsSheet(ss) {
  let sheet = ss.getSheetByName('Mails');
  if (!sheet) {
    sheet = ss.insertSheet('Mails');
  }

  sheet.clear();

  const headers = ['campaign_id', 'uid', 'email', 'token', 'subject', 'body_html', 'body_text', 'send_status', 'sent_at', 'explained'];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#f0f0f0');
  sheet.autoResizeColumns(1, headers.length);
}

/**
 * Clicksシート作成
 */
function createClicksSheet(ss) {
  let sheet = ss.getSheetByName('Clicks');
  if (!sheet) {
    sheet = ss.insertSheet('Clicks');
  }

  sheet.clear();

  const headers = ['ts', 'campaign_id', 'uid', 'email', 'token', 'ua', 'ip_hash', 'referer'];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#f0f0f0');
  sheet.autoResizeColumns(1, headers.length);
}

/**
 * Resultsシート作成
 */
function createResultsSheet(ss) {
  let sheet = ss.getSheetByName('Results');
  if (!sheet) {
    sheet = ss.insertSheet('Results');
  }

  sheet.clear();

  const headers = ['campaign_id', 'campaign_name', '送信数', 'クリック数', 'クリック率', '部署別クリック率', '役職別クリック率', '更新日時'];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#f0f0f0');
  sheet.autoResizeColumns(1, headers.length);
}

/**
 * スクリプトプロパティ初期化
 */
function initializeScriptProperties() {
  const properties = PropertiesService.getScriptProperties();

  // 既存のプロパティがない場合のみ設定
  if (!properties.getProperty('INITIALIZED')) {
    properties.setProperties({
      'INITIALIZED': 'true',
      'PPLX_API_KEY': '', // 実際のAPIキーは後で設定
      'VERSION': '1.0.0'
    });
  }
}

/**
 * 設定値を取得するヘルパー関数
 */
function getConfig(key) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName('Config');

  if (!configSheet) {
    throw new Error('Configシートが見つかりません');
  }

  const data = configSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === key) {
      return data[i][1];
    }
  }

  // スクリプトプロパティからも探す
  const scriptValue = PropertiesService.getScriptProperties().getProperty(key);
  if (scriptValue) {
    return scriptValue;
  }

  return null;
}

/**
 * 設定値を保存するヘルパー関数
 */
function setConfig(key, value) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName('Config');

  if (!configSheet) {
    throw new Error('Configシートが見つかりません');
  }

  const data = configSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === key) {
      configSheet.getRange(i + 1, 2).setValue(value);
      return;
    }
  }

  // 新規行として追加
  const lastRow = configSheet.getLastRow();
  configSheet.getRange(lastRow + 1, 1, 1, 3).setValues([[key, value, '']]);
}