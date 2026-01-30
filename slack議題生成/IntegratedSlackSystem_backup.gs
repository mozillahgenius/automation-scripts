// ==========================================
// 統合版：Slack議題生成＆メッセージ分析システム
// ==========================================

// ========= 設定値 =========
const SPREADSHEET_ID = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID') || '';
const SLACK_BOT_TOKEN = PropertiesService.getScriptProperties().getProperty('SLACK_BOT_TOKEN') || '';
const OPENAI_API_KEY = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY') || '';
const REPORT_EMAIL = PropertiesService.getScriptProperties().getProperty('REPORT_EMAIL') || '';

// パフォーマンス設定
const FETCH_THREAD_REPLIES = false;  // スレッド返信の取得を無効化（エラー回避とパフォーマンス向上）
const MAX_MESSAGES_PER_CHANNEL = 50;  // チャンネルごとの最大取得メッセージ数（パフォーマンス向上のため削減）
const ENABLE_USER_INFO_FETCH = true;  // ユーザー情報の取得を有効化（実名表示のため）
const BATCH_SIZE = 100;  // スプレッドシートへの一括書き込みサイズ

// ========= Slack Bot Token設定ガイド =========
// 
// 【重要】Bot User OAuth Token（xoxb-で始まる）を使用してください
// User OAuth Token（xoxp-）では正常に動作しません
//
// 必須スコープ（Bot Token Scopes）：
// - channels:history     : パブリックチャンネルのメッセージ履歴を読む
// - channels:read        : パブリックチャンネルの基本情報を取得
// - chat:write          : メッセージを投稿する
// - users:read          : ユーザー情報を取得（実名表示のため必須）
// - groups:history      : プライベートチャンネルのメッセージ履歴を読む（重要）
// - groups:read         : プライベートチャンネルの情報を取得（重要）
//
// 追加で推奨されるスコープ：
// - users:read.email     : ユーザーのメールアドレスを取得（完全な情報取得用）
// - im:read             : ダイレクトメッセージを読む
// - mpim:read           : グループダイレクトメッセージを読む
//
// 【重要】プライベートチャンネルへのアクセスには groups:read と groups:history が必須です
//
// 【重要】ユーザーID（@USERID）を実名表示するためには users:read が必須です
//
// 設定手順：
// 1. https://api.slack.com/apps でSlack Appを作成
// 2. OAuth & Permissions ページで上記のBot Token Scopesを追加
//    ※ 必ず users:read を含めてください（実名表示のため）
// 3. 「Install to Workspace」をクリック
//    ※ 権限を追加・変更した場合は必ず「Reinstall to Workspace」を実行
// 4. Bot User OAuth Token (xoxb-...)をコピー
// 5. GASのスクリプトプロパティにSLACK_BOT_TOKENとして設定
// 6. Botを対象チャンネルに招待（/invite @bot-name）
//
// トラブルシューティング：
// - 「invalid_arguments」エラー: Botがチャンネルメンバーではない
// - 「channel_not_found」エラー: チャンネルIDが間違っているか、プライベートチャンネルでBotが未招待
// - 「not_in_channel」エラー: プライベートチャンネルにBotが招待されていない（/invite @bot-name が必要）
// - 「missing_scope」エラー: 必要な権限が不足している（Slack Appで権限追加後、再インストールが必要）
// - ユーザーIDが実名に変換されない: users:read 権限が不足、またはENABLE_USER_INFO_FETCHがfalse
// - プライベートチャンネルが取得できない: 
//   1. Botをチャンネルに招待していない（/invite @Kushim Slack Governance）
//   2. 権限追加後に再インストールしていない
//   3. conversations.listでtypes='private_channel'を指定していない
// - testUserInfoFetch()関数でユーザー情報取得をテストできます
// - diagnoseBotPermissions()関数で詳細診断を実行できます
// - diagnosePrivateChannels()関数でプライベートチャンネルアクセスを詳細診断
// =========================================================

// シート名定義
const SHEETS = {
  CONFIG: 'Config',
  SYNC_STATE: 'SyncState',
  MESSAGES: 'Messages',
  CATEGORIES: 'Categories',
  CHECKLISTS: 'Checklists',
  TEMPLATES: 'Templates',
  DRAFTS: 'Drafts',
  LOGS: 'Logs',
  SLACK_LOG: 'slack_log',
  BUSINESS_MANUAL: 'business_manual',
  FAQ_LIST: 'faq_list',
  DAILY_REPORT: 'daily_report'
};

// ========= slack_logシート作成 =========
function createSlackLogSheet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  return createSlackLogSheetInSpreadsheet(ss);
}

// スプレッドシート指定版のslack_logシート作成
function createSlackLogSheetInSpreadsheet(ss) {
  if (!ss) {
    ss = SpreadsheetApp.getActiveSpreadsheet();
  }
  
  let sheet = ss.getSheetByName(SHEETS.SLACK_LOG);
  
  if (sheet) {
    console.log('slack_logシートは既に存在します');
    return sheet;
  }
  
  sheet = ss.insertSheet(SHEETS.SLACK_LOG);
  const headers = [
    'channel_id', 'channel_name', 'ts', 'thread_ts', 
    'user_name', 'message', 'date', 'reactions', 'files'
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setBackground('#4285F4')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold');
  sheet.setFrozenRows(1);
  
  // 列幅の調整
  sheet.setColumnWidth(1, 120); // channel_id
  sheet.setColumnWidth(2, 150); // channel_name
  sheet.setColumnWidth(3, 120); // ts
  sheet.setColumnWidth(4, 120); // thread_ts
  sheet.setColumnWidth(5, 120); // user_name
  sheet.setColumnWidth(6, 400); // message
  sheet.setColumnWidth(7, 120); // date
  sheet.setColumnWidth(8, 150); // reactions
  sheet.setColumnWidth(9, 150); // files
  
  console.log('slack_logシートを作成しました');
  return sheet;
}

// business_manualシート作成
function createBusinessManualSheet(ss) {
  if (!ss) {
    ss = SpreadsheetApp.getActiveSpreadsheet();
  }
  
  let sheet = ss.getSheetByName(SHEETS.BUSINESS_MANUAL);
  
  if (sheet) {
    console.log('business_manualシートは既に存在します');
    return sheet;
  }
  
  sheet = ss.insertSheet(SHEETS.BUSINESS_MANUAL);
  const headers = [
    '作成日時', 'タイトル', 'カテゴリ', '手順', '詳細説明', 
    '必要なツール', '注意事項', '参考リンク', 'ステータス'
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setBackground('#34A853')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold');
  sheet.setFrozenRows(1);
  
  // 列幅の調整
  sheet.setColumnWidths(1, headers.length, 150);
  sheet.setColumnWidth(4, 300); // 手順
  sheet.setColumnWidth(5, 400); // 詳細説明
  
  console.log('business_manualシートを作成しました');
  return sheet;
}

// faq_listシート作成
function createFAQListSheet(ss) {
  if (!ss) {
    ss = SpreadsheetApp.getActiveSpreadsheet();
  }
  
  let sheet = ss.getSheetByName(SHEETS.FAQ_LIST);
  
  if (sheet) {
    console.log('faq_listシートは既に存在します');
    return sheet;
  }
  
  sheet = ss.insertSheet(SHEETS.FAQ_LIST);
  const headers = [
    '作成日時', '質問', '回答', 'カテゴリ', 'タグ', 
    '元のチャンネル', '関連メッセージ', 'ステータス'
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setBackground('#FBBC04')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold');
  sheet.setFrozenRows(1);
  
  // 列幅の調整
  sheet.setColumnWidths(1, headers.length, 150);
  sheet.setColumnWidth(2, 300); // 質問
  sheet.setColumnWidth(3, 400); // 回答
  
  console.log('faq_listシートを作成しました');
  return sheet;
}

// daily_reportシート作成
function createDailyReportSheet(ss) {
  if (!ss) {
    ss = SpreadsheetApp.getActiveSpreadsheet();
  }
  
  let sheet = ss.getSheetByName(SHEETS.DAILY_REPORT);
  
  if (sheet) {
    console.log('daily_reportシートは既に存在します');
    return sheet;
  }
  
  sheet = ss.insertSheet(SHEETS.DAILY_REPORT);
  const headers = [
    '報告日', 'チャンネル数', 'メッセージ数', '重要議題数', 
    'アクション数', '送信先', 'ステータス', '詳細'
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setBackground('#EA4335')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold');
  sheet.setFrozenRows(1);
  
  // 列幅の調整
  sheet.setColumnWidths(1, headers.length, 120);
  sheet.setColumnWidth(8, 400); // 詳細
  
  console.log('daily_reportシートを作成しました');
  return sheet;
}

// ========= 初期セットアップ =========
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🤖 統合議題生成システム')
    .addItem('⚙️ 初期設定', 'showSetupDialog')
    .addSeparator()
    .addItem('📋 業務フロー生成＆通知', 'getMessagesAsAppWithWorkflow')
    .addSeparator()
    .addSubMenu(ui.createMenu('📥 アプリ統合機能')
      .addItem('🤖 メッセージ取得', 'getMessagesAsApp')
      .addItem('🚀 取得＆分析', 'getMessagesAsAppAndAnalyze')
      .addItem('🏛️ ガバナンス分析', 'getMessagesAsAppWithGovernance'))
    .addSeparator()
    .addSubMenu(ui.createMenu('🔄 同期・分析')
      .addItem('🔄 手動同期（Slack→シート）', 'syncSlackMessages')
      .addItem('🌐 Slack全チャンネル同期', 'syncAllSlackChannels')
      .addItem('🔒 全チャンネル同期（UI安全版）', 'syncAllChannelsSafe')
      .addItem('🔍 プライベートチャンネル存在確認', 'checkPrivateChannelsExist')
      .addItem('🤖 AI分析実行', 'runAIAnalysis')
      .addItem('⚡ メッセージ取得＆分析（簡易版）', 'fetchAndAnalyzeSlackMessages')
      .addItem('✅ Botが参加済みチャンネルから取得＆分析', 'fetchAndAnalyzeFromJoinedChannels')
      .addItem('📊 過去ログ分析', 'analyzeHistoricalMessages'))
    .addSeparator()
    .addSubMenu(ui.createMenu('📝 文書生成')
      .addItem('🎯 Slack議題抽出＆メール送信', 'analyzeSlackAndSendReport')
      .addItem('📝 選択行でドラフト生成', 'generateDraftForSelected')
      .addItem('📊 複数議案の一括議事録作成', 'generateBatchMinutes')
      .addItem('📚 マニュアル・FAQ生成', 'generateManualAndFAQFromMessages'))
    .addSeparator()
    .addItem('➕ Botをチャンネルに追加', 'joinBotToChannel')
    .addItem('🔐 プライベートチャンネル自動検出＆招待リスト', 'generateInviteList')
    .addItem('🔍 チャンネル名からID検索', 'getChannelIdByName')
    .addSeparator()
    .addItem('📧 日次レポート送信', 'sendDailyReport')
    .addItem('📢 通知テスト', 'testNotification')
    .addItem('👥 ユーザー情報を更新', 'refreshUserInfo')
    .addItem('🔍 プライベートチャンネル診断', 'debugPrivateChannelsComplete')
    .addItem('🔐 プライベートチャンネルアクセステスト', 'testPrivateChannelAccess')
    .addItem('🔄 プライベートチャンネル同期', 'syncPrivateChannels')
    .addSeparator()
    .addSubMenu(ui.createMenu('📋 シート管理')
      .addItem('✅ 全シートの存在確認', 'checkAllSheets')
      .addSeparator()
      .addItem('⚙️ Configシート作成', 'createConfigSheetOnly')
      .addItem('🔄 SyncStateシート作成', 'createSyncStateSheetOnly')
      .addItem('💬 Messagesシート作成', 'createMessagesSheetOnly')
      .addItem('🏷️ Categoriesシート作成', 'createCategoriesSheetOnly')
      .addItem('✅ Checklistsシート作成', 'createChecklistsSheetOnly')
      .addItem('📝 Templatesシート作成', 'createTemplatesSheetOnly')
      .addItem('📄 Draftsシート作成', 'createDraftsSheetOnly')
      .addItem('📊 Logsシート作成', 'createLogsSheetOnly')
      .addItem('📝 slack_logシート作成', 'createSlackLogSheet')
      .addSeparator()
      .addItem('🗄️ 全シート初期化', 'initializeSpreadsheet'))
    .addSubMenu(ui.createMenu('📄 テンプレート管理')
      .addItem('📋 テンプレート初期化', 'initializeTemplates')
      .addItem('📝 テンプレート作成', 'createTemplateDocuments')
      .addItem('📑 テンプレート一覧表示', 'showTemplateList'))
    .addSubMenu(ui.createMenu('🔧 診断・テスト')
      .addItem('🏥 Slackメッセージ取得診断', 'testSlackMessageRetrieval')
      .addItem('📡 チャンネルアクセス診断', 'diagnoseChannelAccess')
      .addItem('🚨 チャンネルエラー診断', 'diagnoseChannelNotFoundError')
      .addItem('🔑 Bot権限診断', 'diagnoseBotPermissions')
      .addItem('🔬 詳細Bot診断', 'detailedBotDiagnostics')
      .addItem('📥 簡易メッセージ取得', 'getSlackMessagesSimple')
      .addItem('🛡️ 安全なメッセージ取得', 'getSlackMessagesSafe'))
    .addSubMenu(ui.createMenu('🤖 AIモデル設定')
      .addItem('📊 現在のモデル確認', 'showCurrentModel')
      .addSeparator()
      .addItem('🚀 o3に切り替え（推奨）', 'useO3')
      .addItem('⚡ GPT-4oに切り替え', 'useGPT4o')
      )
    .addSubMenu(ui.createMenu('🔧 デバッグ・診断')
      .addItem('⚙️ 設定内容確認', 'checkConfigSettings')
      .addSeparator()
      .addItem('🔌 Slack API接続テスト', 'testSlackConnection')
      .addItem('🐛 Slack APIデバッグ情報', 'debugSlackAPI')
      .addItem('📡 Slack API呼び出しテスト', 'testSlackAPICall')
      .addItem('🔍 チャンネルアクセステスト', 'testChannelAccess'))
    .addSeparator()
    .addItem('⏰ 自動実行タイマー設定', 'installTriggers')
    .addItem('🗑️ タイマー削除', 'removeTriggers')
    .addItem('🔍 タイマー確認', 'checkTriggers')
    .addItem('🧹 重複タイマー削除', 'removeDuplicateTriggers')
    .addToUi();
}

// ========= プロパティ管理 =========
function showSetupDialog() {
  const html = HtmlService.createHtmlOutputFromFile('setup')
    .setWidth(600)
    .setHeight(500);
  SpreadsheetApp.getUi().showModalDialog(html, '初期設定');
}

function saveSettings(settings) {
  const scriptProperties = PropertiesService.getScriptProperties();
  Object.keys(settings).forEach(key => {
    if (settings[key]) {
      scriptProperties.setProperty(key, settings[key]);
    }
  });
  return '設定を保存しました';
}

function getSettings() {
  return PropertiesService.getScriptProperties().getProperties();
}

// ========= トリガー管理 =========
function installTriggers() {
  removeTriggers(); // 既存のトリガーを削除
  
  // 30分ごとにSlack同期
  ScriptApp.newTrigger('syncAllSlackChannels')
    .timeBased()
    .everyMinutes(30)
    .create();
  
  // 1時間ごとにAI分析
  ScriptApp.newTrigger('runAIAnalysis')
    .timeBased()
    .everyHours(1)
    .create();
  
  // 毎日9時に日次レポート
  ScriptApp.newTrigger('sendDailyReport')
    .timeBased()
    .atHour(9)
    .everyDays(1)
    .create();
  
  // 毎日15時に議題候補通知
  ScriptApp.newTrigger('sendDailyNotification')
    .timeBased()
    .atHour(15)
    .everyDays(1)
    .create();
  
  SpreadsheetApp.getUi().alert('自動実行タイマーを設定しました');
}

function removeTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => ScriptApp.deleteTrigger(trigger));
}

// トリガーの状態を確認
function checkTriggers() {
  const ui = SpreadsheetApp.getUi();
  const triggers = ScriptApp.getProjectTriggers();
  
  if (triggers.length === 0) {
    ui.alert('確認', 'トリガーは設定されていません。', ui.ButtonSet.OK);
    return;
  }
  
  let triggerInfo = '=== 現在のトリガー設定 ===\n\n';
  const triggerCount = {};
  
  triggers.forEach((trigger, index) => {
    const handlerFunction = trigger.getHandlerFunction();
    const triggerSource = trigger.getTriggerSource();
    const eventType = trigger.getEventType();
    
    // 同じ関数のトリガーをカウント
    triggerCount[handlerFunction] = (triggerCount[handlerFunction] || 0) + 1;
    
    triggerInfo += `${index + 1}. ${handlerFunction}\n`;
    triggerInfo += `   タイプ: ${triggerSource}\n`;
    
    if (triggerSource === ScriptApp.TriggerSource.CLOCK) {
      // 時間ベースのトリガーの詳細を取得
      triggerInfo += `   イベント: ${eventType}\n`;
    }
    triggerInfo += '\n';
  });
  
  // 重複チェック
  let duplicates = [];
  Object.entries(triggerCount).forEach(([func, count]) => {
    if (count > 1) {
      duplicates.push(`${func}: ${count}個`);
    }
  });
  
  if (duplicates.length > 0) {
    triggerInfo += '⚠️ 重複しているトリガー:\n';
    triggerInfo += duplicates.join('\n');
  }
  
  ui.alert('トリガー確認', triggerInfo, ui.ButtonSet.OK);
}

// ===== ユーティリティ関数 =====

// ユーザーID→表示名
function getSlackUserName(userId) {
  const token = PropertiesService.getScriptProperties().getProperty('SLACK_BOT_TOKEN');
  if (!userId) return "";
  try {
    const res  = UrlFetchApp.fetch(
      `https://slack.com/api/users.info?user=${ userId }`,
      { method: 'get', headers: { Authorization: 'Bearer ' + token } }
    );
    const json = JSON.parse(res.getContentText());
    if (json.ok && json.user && json.user.profile) {
      return json.user.profile.display_name || json.user.real_name || userId;
    }
  } catch (e) {
    Logger.log("getSlackUserName error: " + e);
  }
  return userId;
}

// <@U12345>→@表示名
function convertMentionsToDisplayName(text) {
  if (!text) return "";
  return text.replace(/<@([A-Z0-9]+)>/g, (_, uid) => '@' + getSlackUserName(uid));
}

// チャンネルID→チャンネル名
function getSlackChannelName(channelId) {
  const token = PropertiesService.getScriptProperties().getProperty('SLACK_BOT_TOKEN');
  try {
    const res  = UrlFetchApp.fetch(
      `https://slack.com/api/conversations.info?channel=${ channelId }`,
      { method: 'get', headers: { Authorization: 'Bearer ' + token } }
    );
    const json = JSON.parse(res.getContentText());
    if (json.ok && json.channel && json.channel.name) {
      return json.channel.name;
    }
  } catch (e) {
    Logger.log("getSlackChannelName error: " + e);
  }
  return channelId;
}

// ========= 包括的AI分析（強化版） =========
function performEnhancedAIAnalysis(messages) {
  const messageText = messages.map(m => {
    const timestamp = m.ts ? new Date(m.ts * 1000).toLocaleString('ja-JP') : '';
    const user = m.user || 'unknown';
    const text = m.text || '';
    return `[${timestamp}] ${user}: ${text}`;
  }).join('\n');
  
  const prompt = `
以下のSlackメッセージを分析し、純粋なJSON形式のみで結果を返してください。
重要：マークダウンのコードブロック（\`\`\`json）は使用せず、JSONデータのみを直接返してください。

【分析項目】
1. カテゴリ分類（予算、契約、人事、開発、マーケティング、法務、財務、その他）
2. 議題・論点の抽出（最大5つ、優先順位付き）
   - 日常的な業務連絡や単なる情報共有は議題として抽出しない
   - 決定や検討が必要な事項のみを議題とする
3. 重要度判定（HIGH/MEDIUM/LOW）とその理由
   - HIGH: 経営に重大な影響がある、法的リスクがある、多額の金銭が関わる場合のみ
   - MEDIUM: 部門横断的な調整が必要、中程度の予算が関わる場合
   - LOW: それ以外の通常業務
4. 具体的なアクションアイテム（実際に行動が必要なもののみ）
5. 関係者のリスト
6. 緊急度の評価
7. 決定事項（明確に決定された事項のみ）
8. リスク要因（実際にリスクがある場合のみ）

メッセージ:
${messageText.substring(0, 2000)}

以下の構造のJSONを返してください（コードブロックなしで純粋なJSONのみ）:
{
  "categories": ["カテゴリ名"],
  "topics": [{"title": "議題", "description": "説明", "priority": 1}],
  "priority": "MEDIUM",
  "priorityReason": "理由",
  "actionItems": [{"task": "タスク", "owner": "担当", "deadline": "期限"}],
  "stakeholders": ["関係者"],
  "urgency": "normal",
  "deadline": "",
  "decisions": ["決定事項"],
  "risks": [{"risk": "リスク", "impact": "medium", "mitigation": "対策"}],
  "resources": {"human": [], "financial": "", "time": ""},
  "kpis": [],
  "summary": "要約"
}`;

  try {
    console.log('AI分析開始: o3モデルを使用');
    let response = callOpenAIAPI(prompt, 'gpt-5');
    
    console.log('OpenAI APIレスポンス受信');
    console.log('Response type:', typeof response);
    console.log('Response length:', response ? response.length : 'null/undefined');
    
    // レスポンスが空の場合のチェック
    if (!response || (typeof response === 'string' && response.trim() === '')) {
      console.error('AI分析: 空のレスポンス');
      console.error('Response value:', response);
      throw new Error('Empty response from AI');
    }
    
    // マークダウンコードブロックを除去
    response = response.replace(/^```json\s*/i, '').replace(/```\s*$/i, '').trim();
    
    // JSONパースを試行
    let result;
    try {
      result = JSON.parse(response);
    } catch (parseError) {
      console.error('AI分析: JSONパースエラー');
      console.error('Parse error:', parseError.toString());
      console.error('Response (first 500 chars):', response.substring(0, 500));
      console.error('Response length:', response.length);
      
      // レスポンスの最初と最後の文字をチェック
      if (response.length > 0) {
        console.error('First char:', response.charAt(0), 'Code:', response.charCodeAt(0));
        console.error('Last char:', response.charAt(response.length - 1), 'Code:', response.charCodeAt(response.length - 1));
      }
      
      throw parseError;
    }
    
    // 必須フィールドのデフォルト値を設定
    result.categories = result.categories || ['その他'];
    result.topics = result.topics || [];
    result.priority = result.priority || 'MEDIUM';
    result.actionItems = result.actionItems || [];
    result.stakeholders = result.stakeholders || [];
    result.urgency = result.urgency || 'normal';
    result.decisions = result.decisions || [];
    result.risks = result.risks || [];
    result.resources = result.resources || { human: [], financial: '', time: '' };
    result.kpis = result.kpis || [];
    result.summary = result.summary || 'メッセージを分析しました';
    
    // ユーザーIDを実名に変換
    return convertAnalysisUserIds(result);
    
  } catch (error) {
    console.error('AI分析エラー:', error);
    if (typeof response !== 'undefined') {
      console.error('Response type:', typeof response);
      console.error('Response (first 200 chars):', String(response).substring(0, 200));
    }
    
    // エラー時のデフォルト分析結果
    const defaultResult = {
      categories: ['その他'],
      topics: [{
        title: 'メッセージ分析',
        description: `${messages.length}件のメッセージを処理`,
        priority: 1
      }],
      priority: 'MEDIUM',
      priorityReason: '自動判定',
      actionItems: [],
      stakeholders: [],
      urgency: 'normal',
      deadline: '',
      decisions: [],
      risks: [],
      resources: { human: [], financial: '', time: '' },
      kpis: [],
      summary: `${messages.length}件のメッセージを処理しました。AI分析に問題がありました。`,
      error: error.toString()
    };
    // エラー時も念のため変換を適用
    return convertAnalysisUserIds(defaultResult);
  }
}

// ========= 包括的ガバナンス・コンプライアンスチェック =========
function performComprehensiveGovernanceCheck(messages, analysisResult) {
  const checkResult = {
    requiresApproval: false,
    approvalLevel: '', // 部長/取締役会/株主総会
    requiresDisclosure: false,
    disclosureType: '', // 適時開示/決算開示/任意開示
    requiresMeetingMinutes: false,
    meetingType: '', // 取締役会/監査等委員会/株主総会
    requiresExpertConsultation: false,
    requiredExperts: [],
    riskLevel: 'LOW', // HIGH/MEDIUM/LOW
    complianceGaps: [],
    auditPoints: [],
    controlNumber: generateControlNumber(),
    internalControlIssues: [],
    requiresAction: false
  };
  
  // キーワードベースのチェック（より厳密な判定）
  const approvalKeywords = {
    '部長承認': ['1000万円以上の予算', '重要な購買', '部長決裁', '管理職採用'],
    '取締役会': ['M&A', '億円以上の契約', '組織再編', '重要な規程改定', '大型投資'],
    '株主総会': ['定款変更', '取締役選任', '剰余金配当', '増資決議', '減資決議']
  };
  
  const disclosureKeywords = {
    '適時開示': ['業績予想修正', '重要事実発生', '決算短信公表', '重要な業務提携'],
    '決算開示': ['四半期決算', '通期決算', '決算短信'],
    '任意開示': ['中期経営計画発表', 'IR説明会']
  };
  
  const expertKeywords = {
    '弁護士': ['訴訟', '重大な法的問題', 'コンプライアンス違反', '重要契約書レビュー'],
    '会計士': ['会計監査', '財務諸表', '重要な会計処理', '税務調査'],
    '税理士': ['税務申告', '重要な節税対策', '税務調査対応'],
    '社労士': ['労基署対応', '重要な就業規則改定', '労務トラブル'],
    '弁理士': ['特許出願', '商標登録', '知財紛争']
  };
  
  const combinedText = JSON.stringify(analysisResult) + ' ' + messages.map(m => m.text || '').join(' ');
  
  // 承認レベル判定
  Object.entries(approvalKeywords).forEach(([level, keywords]) => {
    if (keywords.some(keyword => combinedText.includes(keyword))) {
      checkResult.requiresApproval = true;
      checkResult.approvalLevel = level;
      checkResult.requiresMeetingMinutes = level !== '部長承認';
      if (level === '取締役会') checkResult.meetingType = '取締役会';
      if (level === '株主総会') checkResult.meetingType = '株主総会';
    }
  });
  
  // 開示要件チェック
  Object.entries(disclosureKeywords).forEach(([type, keywords]) => {
    if (keywords.some(keyword => combinedText.includes(keyword))) {
      checkResult.requiresDisclosure = true;
      checkResult.disclosureType = type;
      checkResult.riskLevel = 'HIGH';
    }
  });
  
  // 専門家相談要件
  Object.entries(expertKeywords).forEach(([expert, keywords]) => {
    if (keywords.some(keyword => combinedText.includes(keyword))) {
      checkResult.requiresExpertConsultation = true;
      if (!checkResult.requiredExperts.includes(expert)) {
        checkResult.requiredExperts.push(expert);
      }
    }
  });
  
  // リスクレベル評価
  if (analysisResult.priority === 'HIGH') {
    checkResult.riskLevel = checkResult.riskLevel === 'LOW' ? 'MEDIUM' : checkResult.riskLevel;
    checkResult.auditPoints.push('高重要度案件のため監査対象');
  }
  
  if (analysisResult.risks && analysisResult.risks.some(r => r.impact === 'high')) {
    checkResult.riskLevel = 'HIGH';
    checkResult.auditPoints.push('高リスク要因が識別されたため重点監査対象');
  }
  
  // 内部統制チェック
  if (combinedText.includes('統制') || combinedText.includes('ルール違反')) {
    checkResult.internalControlIssues.push('内部統制の見直しが必要');
  }
  
  // コンプライアンスギャップ
  if (combinedText.includes('法令') || combinedText.includes('規制')) {
    checkResult.complianceGaps.push('法令遵守状況の確認が必要');
  }
  
  // アクション要否
  checkResult.requiresAction = checkResult.requiresApproval || 
                               checkResult.requiresDisclosure || 
                               checkResult.riskLevel === 'HIGH' ||
                               checkResult.requiresExpertConsultation;
  
  return checkResult;
}

// ========= 議題の分類（重要議題とトピック） =========
function classifyTopics(analysisResult, governanceCheck) {
  const classified = {
    importantAgendas: [],  // 重要議題（決議事項、開示事項等）
    generalTopics: []      // 一般的なトピック
  };
  
  if (!analysisResult.topics || !Array.isArray(analysisResult.topics)) {
    return classified;
  }
  
  // 重要議題のキーワード（より厳密な判定）
  const importantPatterns = [
    // 取締役会決議事項（より具体的なパターン）
    { pattern: /M&A|企業買収|事業譲渡|合併/i, category: '取締役会' },
    { pattern: /億円以上の.*契約|重要な資産.*譲渡/i, category: '取締役会' },
    { pattern: /組織再編|会社分割|持株会社/i, category: '取締役会' },
    { pattern: /定款.*変更|取締役.*選任|役員.*解任/i, category: '株主総会' },
    
    // 監査等委員会決議事項（重大な事項のみ）
    { pattern: /不正.*発覚|重大.*コンプライアンス.*違反|法令違反/i, category: '監査等委員会' },
    { pattern: /内部統制.*重大.*不備|監査.*指摘事項/i, category: '監査等委員会' },
    
    // 株主総会決議事項（明確な決議事項）
    { pattern: /配当.*決議|剰余金.*配当/i, category: '株主総会' },
    { pattern: /増資.*決議|減資.*決議|自己株式.*取得/i, category: '株主総会' },
    
    // 東証開示事項（適時開示が必要な事項）
    { pattern: /業績予想.*修正|決算短信|適時開示.*必要/i, category: '東証開示' },
    { pattern: /重要事実.*発生|インサイダー/i, category: '東証開示' },
    
    // 財務局照会事項（報告書関連）
    { pattern: /有価証券報告書.*提出|内部統制報告書.*作成/i, category: '財務局' }
  ];
  
  analysisResult.topics.forEach(topic => {
    const topicText = topic.title + ' ' + (topic.description || '');
    let isImportant = false;
    let category = '';
    
    // パターンマッチングによるチェック
    for (const patternObj of importantPatterns) {
      if (patternObj.pattern.test(topicText)) {
        isImportant = true;
        category = patternObj.category;
        break;
      }
    }
    
    // ガバナンスチェック結果も考慮（ただし、より厳密に）
    if (!isImportant && governanceCheck) {
      // 取締役会・株主総会レベルの承認が必要な場合のみ
      if (governanceCheck.requiresApproval && 
          (governanceCheck.approvalLevel === '取締役会' || 
           governanceCheck.approvalLevel === '株主総会')) {
        isImportant = true;
        category = governanceCheck.approvalLevel;
      }
      // 適時開示が必要な場合
      if (governanceCheck.requiresDisclosure && 
          governanceCheck.disclosureType === '適時開示') {
        isImportant = true;
        category = '東証開示';
      }
    }
    
    // 注意：優先度だけでは重要議題にしない（削除）
    
    if (isImportant) {
      classified.importantAgendas.push({
        ...topic,
        category: category || '重要議題'
      });
    } else {
      classified.generalTopics.push(topic);
    }
  });
  
  return classified;
}

// 重複トリガーを削除
function removeDuplicateTriggers() {
  const ui = SpreadsheetApp.getUi();
  const triggers = ScriptApp.getProjectTriggers();
  
  const triggerMap = new Map();
  let removedCount = 0;
  
  triggers.forEach(trigger => {
    const handlerFunction = trigger.getHandlerFunction();
    const key = `${handlerFunction}_${trigger.getTriggerSource()}_${trigger.getEventType()}`;
    
    if (triggerMap.has(key)) {
      // 重複トリガーを削除
      ScriptApp.deleteTrigger(trigger);
      removedCount++;
    } else {
      triggerMap.set(key, trigger);
    }
  });
  
  if (removedCount > 0) {
    ui.alert('完了', `${removedCount}個の重複トリガーを削除しました。`, ui.ButtonSet.OK);
  } else {
    ui.alert('確認', '重複トリガーは見つかりませんでした。', ui.ButtonSet.OK);
  }
}

// ========= Slack API 連携 =========
function slackAPI(method, params = {}) {
  // メソッド名のバリデーション
  if (!method || method === 'undefined' || typeof method !== 'string') {
    // 直接実行の可能性が高い
    const stack = new Error().stack;
    if (stack.includes('__GS_INTERNAL_top_function_call__') && !method) {
      console.error('slackAPI関数は直接実行できません。');
      console.log('使用例: slackAPI("conversations.list", { types: "public_channel" })');
      console.log('利用可能なメソッド: conversations.list, conversations.history, users.info, chat.postMessage など');
      throw new Error('slackAPI関数は直接実行できません。他の関数から呼び出してください。');
    }
    
    const errorMsg = `無効なAPIメソッド: ${method}`;
    logError('Slack API', errorMsg);
    console.error('スタックトレース:', stack);
    throw new Error(errorMsg);
  }
  
  // Bot Tokenの確認
  if (!SLACK_BOT_TOKEN || SLACK_BOT_TOKEN === '') {
    const errorMsg = 'Slack Bot Tokenが設定されていません。スクリプトプロパティでSLACK_BOT_TOKENを設定してください。';
    console.error(errorMsg);
    throw new Error(errorMsg);
  }
  
  // Bot Tokenの形式確認
  if (!SLACK_BOT_TOKEN.startsWith('xoxb-')) {
    console.warn('警告: Bot TokenがUser Token (xoxp-)の可能性があります。Bot Token (xoxb-)を使用することを推奨します。');
  }
  
  // 動作確認済みのコードと同じ方法でAPI呼び出し
  // conversations.listは GET メソッドでも動作する（プライベートチャンネル取得に重要）
  if (method === 'conversations.list' && params && params.types) {
    const queryParams = Object.keys(params).map(key => 
      `${key}=${encodeURIComponent(params[key])}`
    ).join('&');
    
    const getUrl = `https://slack.com/api/${method}?${queryParams}`;
    const getOptions = {
      method: 'get',
      headers: {
        'Authorization': `Bearer ${SLACK_BOT_TOKEN}`
      },
      muteHttpExceptions: true
    };
    
    try {
      console.log(`Slack API呼び出し (GET): ${method} - types=${params.types}`);
      const response = UrlFetchApp.fetch(getUrl, getOptions);
      const result = JSON.parse(response.getContentText());
      
      if (result.ok && result.channels) {
        const privateChannels = result.channels.filter(ch => ch.is_private === true);
        if (privateChannels.length > 0) {
          console.log(`✅ プライベートチャンネル検出: ${privateChannels.length}個`);
        }
      }
      
      return result;
    } catch (error) {
      console.error(`Slack API Error (${method}):`, error.toString());
      throw error;
    }
  }
  
  // 通常のPOSTメソッド
  const url = `https://slack.com/api/${method}`;
  const options = {
    method: 'post',
    headers: {
      'Authorization': `Bearer ${SLACK_BOT_TOKEN}`,
      'Content-Type': 'application/json; charset=utf-8'
    },
    payload: JSON.stringify(params),
    muteHttpExceptions: true
  };
  
  try {
    console.log(`Slack API呼び出し: ${method}`);
    logInfo(`Slack API Call: ${method}`, JSON.stringify(params).substring(0, 200));
    
    const response = UrlFetchApp.fetch(url, options);
    const responseText = response.getContentText();
    
    // レスポンスが空の場合のエラーハンドリング
    if (!responseText) {
      throw new Error('Slack APIから空のレスポンスが返されました');
    }
    
    const data = JSON.parse(responseText);
    
    if (!data.ok) {
      // 詳細なエラー情報をログに記録
      console.error(`Slack APIエラー詳細:`);
      console.error(`  メソッド: ${method}`);
      console.error(`  エラー: ${data.error}`);
      console.error(`  パラメータ: ${JSON.stringify(params)}`);
      
      logError('Slack API', `Method: ${method}, Error: ${data.error}, Params: ${JSON.stringify(params)}`);
      
      // よくあるエラーの原因を提示
      if (data.error === 'unknown_method') {
        throw new Error(`Slack API Error: ${data.error} - メソッド「${method}」は存在しません。APIメソッド名を確認してください。`);
      } else if (data.error === 'not_authed') {
        throw new Error(`Slack API Error: ${data.error} - Bot Tokenが設定されていないか無効です。`);
      } else if (data.error === 'invalid_auth') {
        throw new Error(`Slack API Error: ${data.error} - Bot Tokenが無効です。正しいトークンを設定してください。`);
      } else if (data.error === 'channel_not_found') {
        throw new Error(`Slack API Error: ${data.error} - チャンネルID「${params.channel}」が見つかりません。`);
      } else if (data.error === 'not_in_channel') {
        throw new Error(`Slack API Error: ${data.error} - Botがチャンネル「${params.channel}」のメンバーではありません。/invite @bot-name でBotを招待してください。`);
      } else if (data.error === 'invalid_arguments') {
        throw new Error(`Slack API Error: ${data.error} - 無効な引数です。Botがチャンネルのメンバーでない可能性があります。`);
      } else if (data.error === 'missing_scope') {
        throw new Error(`Slack API Error: ${data.error} - 必要なスコープが不足しています。Bot Token Scopesを確認してください。`);
      } else {
        throw new Error(`Slack API Error: ${data.error}`);
      }
    }
    
    return data;
  } catch (error) {
    logError('Slack API', `Method: ${method}, Error: ${error.toString()}`);
    throw error;
  }
}

// ========= 個別チャンネルSlackメッセージ同期 =========
function syncSlackMessages() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const configSheet = ss.getSheetByName(SHEETS.CONFIG);
  const syncSheet = getOrCreateSheet(ss, SHEETS.SYNC_STATE, ['channel_id', 'last_sync_ts', 'last_sync_datetime', 'message_count', 'status']);
  const messagesSheet = getOrCreateSheet(ss, SHEETS.MESSAGES, [
    'id', 'channel_id', 'message_ts', 'thread_ts', 'text_raw', 'user_name',
    'summary_json', 'classification_json', 'match_flag', 'human_judgement',
    'permalink', 'checklist_proposed', 'agenda_selected', 'draft_doc_url',
    'created_at', 'updated_at'
  ]);
  
  // 設定からチャンネルIDを取得
  const config = getConfigData(configSheet);
  const channelIds = config.targetChannels || [];
  
  if (channelIds.length === 0) {
    logInfo('監視対象チャンネルが設定されていません');
    ui.alert('エラー', '監視対象チャンネルが設定されていません。\nConfigシートを確認してください。', ui.ButtonSet.OK);
    return;
  }
  
  let totalMessageCount = 0;
  let syncResults = [];
  
  channelIds.forEach(channelId => {
    try {
      console.log(`チャンネル ${channelId} の同期を開始...`);
      
      // チャンネル情報を取得して確認
      const channelInfo = getChannelInfo(channelId);
      if (!channelInfo) {
        throw new Error(`チャンネル情報を取得できません。Botがチャンネルに参加していない可能性があります。`);
      }
      
      console.log(`チャンネル名: ${channelInfo.name}`);
      
      // 最終同期タイムスタンプを取得
      const lastTs = getLastSyncTime(syncSheet, channelId);
      console.log(`最終同期タイムスタンプ: ${lastTs || 'なし（初回同期）'}`);
      
      // メッセージ履歴を取得
      const messages = fetchChannelHistory(channelId, lastTs);
      console.log(`取得したメッセージ数: ${messages.length}`);
      
      // メッセージをスプレッドシートに保存
      let savedCount = 0;
      const startTime = new Date().getTime();
      const maxExecutionTime = 5 * 60 * 1000; // 5分のタイムアウト
      
      for (let i = 0; i < messages.length; i++) {
        // タイムアウトチェック
        if (new Date().getTime() - startTime > maxExecutionTime) {
          console.log('実行時間制限に達しました。処理を中断します。');
          break;
        }
        
        const message = messages[i];
        console.log(`メッセージ ${i + 1}/${messages.length} を処理中`);
        
        saveMessage(messagesSheet, channelId, message);
        savedCount++;
        
        // スレッドがある場合は返信も取得（スキップ可能）
        if (message.thread_ts && message.reply_count > 0 && message.reply_count < 10) {
          // 返信が10件未満の場合のみ取得（パフォーマンス対策）
          const replies = fetchThreadReplies(channelId, message.thread_ts);
          replies.forEach(reply => {
            saveMessage(messagesSheet, channelId, reply);
            savedCount++;
          });
        } else if (message.reply_count >= 10) {
          console.log(`スレッド返信が多い(${message.reply_count}件)ため、スキップ`);
        }
      }
      
      // 最終同期時刻を更新
      // 注意: fetchChannelHistoryは古い順に並べ替えているので、最後が最新
      if (messages.length > 0) {
        const latestMessage = messages[messages.length - 1];
        console.log(`最新メッセージのタイムスタンプ: ${latestMessage.ts} - ${new Date(parseFloat(latestMessage.ts) * 1000).toLocaleString('ja-JP')}`);
        updateLastSyncTime(syncSheet, channelId, latestMessage.ts);
      } else {
        // メッセージがない場合も同期時刻を更新（現在時刻を使用）
        const now = new Date();
        const nowTs = Math.floor(now.getTime() / 1000).toString();
        console.log(`新規メッセージなし。現在時刻で更新: ${nowTs}`);
        updateLastSyncTime(syncSheet, channelId, nowTs);
      }
      
      totalMessageCount += savedCount;
      syncResults.push(`${channelInfo.name} (${channelId}): ${messages.length}件の新規メッセージ`);
      logInfo(`チャンネル ${channelId}: ${messages.length}件のメッセージを同期`);
      
    } catch (error) {
      const errorMsg = `チャンネル ${channelId}: エラー - ${error.toString()}`;
      syncResults.push(errorMsg);
      logError(`Channel ${channelId} sync`, error.toString());
    }
  });
  
  // 結果を表示
  const resultMessage = syncResults.join('\n');
  ui.alert(
    '同期完了', 
    `Slackメッセージの同期が完了しました。\n\n${resultMessage}\n\n合計: ${totalMessageCount}件のメッセージを保存`, 
    ui.ButtonSet.OK
  );
  
  logInfo('Slack個別チャンネル同期完了');
}

// ========= 全チャンネルSlackメッセージ同期 =========
function syncAllSlackChannels() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const configSheet = ss.getSheetByName(SHEETS.CONFIG);
  const syncSheet = ss.getSheetByName(SHEETS.SYNC_STATE);
  const messagesSheet = ss.getSheetByName(SHEETS.MESSAGES);
  
  // slack_logシートを明示的に作成または取得
  console.log('slack_logシートを確認中...');
  const slackLogSheet = getOrCreateSheet(ss, SHEETS.SLACK_LOG, [
    'channel_id', 'channel_name', 'ts', 'thread_ts', 
    'user_name', 'message', 'date', 'reactions', 'files'
  ]);
  
  if (!slackLogSheet) {
    console.error('slack_logシートの作成に失敗しました');
    return;
  }
  console.log(`slack_logシートの準備完了: ${slackLogSheet.getName()}`);
  
  // 設定からチャンネルIDを取得
  const config = getConfigData(configSheet);
  const channelIds = config.targetChannels || [];
  
  // Botが参加しているすべてのチャンネルを取得
  const joinedChannels = getAllJoinedChannels();
  console.log(`Botが参加しているチャンネル数: ${joinedChannels.length}`);
  
  // 設定されたチャンネルIDとBotが参加しているチャンネルを統合
  const joinedChannelIds = joinedChannels.map(ch => ch.id);
  const channelsToSync = [...new Set([...channelIds, ...joinedChannelIds])];
  
  // 実際にBotが参加しているチャンネルのみをフィルタ
  const validChannels = channelsToSync.filter(channelId => 
    joinedChannelIds.includes(channelId)
  );
  
  console.log(`同期対象チャンネル: ${validChannels.length}個`);
  validChannels.forEach((id, index) => {
    const channel = joinedChannels.find(ch => ch.id === id);
    console.log(`  ${index + 1}. ${channel?.name || id} (${id})`);
  });
  
  // 通知チャンネルを除外
  const notifyChannel = config.notifySlackChannel;
  const filteredChannels = validChannels.filter(ch => ch !== notifyChannel);
  
  if (notifyChannel && validChannels.length !== filteredChannels.length) {
    console.log(`通知チャンネル ${notifyChannel} を同期対象から除外しました`);
  }
  
  let totalMessageCount = 0;
  
  filteredChannels.forEach(channelId => {
    try {
      // チャンネル情報を取得
      const channelInfo = getChannelInfo(channelId);
      const channelName = channelInfo?.name || channelId;
      
      // 最終同期タイムスタンプを取得
      const lastTs = getLastSyncTime(syncSheet, channelId);
      console.log(`最終同期タイムスタンプ: ${lastTs}`);
      
      // タイムスタンプの検証とメッセージ取得
      let messages;
      if (lastTs && lastTs !== '0' && isNaN(parseFloat(lastTs))) {
        console.warn(`無効なタイムスタンプをリセット: ${lastTs}`);
        updateLastSyncTime(syncSheet, channelId, '0');
        messages = fetchChannelHistoryWithDetails(channelId, '0');
      } else {
        // メッセージ履歴を取得
        messages = fetchChannelHistoryWithDetails(channelId, lastTs);
      }
      
      // バッチ処理のためのデータ準備
      const messageBatch = [];
      const slackLogBatch = [];
      
      messages.forEach(message => {
        // メインメッセージをバッチに追加
        messageBatch.push(prepareMessageRow(channelId, message));
        slackLogBatch.push(prepareSlackLogRow(channelId, channelName, message));
        
        // スレッド返信の取得（設定で有効化されている場合のみ）
        if (FETCH_THREAD_REPLIES && message.thread_ts && message.reply_count > 0 && message.reply_count <= 3) {
          // プライベートチャンネルでもBotが参加していれば返信を取得
            // 既にスレッド返信が保存されているか確認
            const threadId = `${channelId}_${message.thread_ts}_1`; // 最初の返信のID
            const existingData = messagesSheet.getDataRange().getValues();
            const threadExists = existingData.some(row => row[0] && row[0].toString().startsWith(`${channelId}_${message.thread_ts}_`));
            
            if (threadExists) {
              console.log(`スレッド返信は既に取得済み: ${message.thread_ts}`);
            } else {
              console.log(`スレッド返信を取得: ${message.thread_ts} (${message.reply_count}件)`);
              try {
                const replies = fetchThreadReplies(channelId, message.thread_ts);
                if (replies && replies.length > 0) {
                  replies.slice(0, 3).forEach(reply => {
                    messageBatch.push(prepareMessageRow(channelId, reply));
                    slackLogBatch.push(prepareSlackLogRow(channelId, channelName, reply));
                  });
                  console.log(`スレッド返信${replies.length}件を追加`);
                }
              } catch (error) {
                // スレッド返信取得エラーは警告レベルに留める（処理は継続）
                console.warn(`スレッド返信取得をスキップ (${message.thread_ts}): ${error.toString()}`);
            }
          }
        }
      });
      
      // バッチで一括保存
      if (messageBatch.length > 0) {
        console.log(`${messageBatch.length}件のメッセージをバッチ保存`);
        saveMessagesBatch(messagesSheet, messageBatch);
        saveSlackLogBatch(slackLogSheet, slackLogBatch);
      }
      
      totalMessageCount += messages.length;
      
      // 最終同期時刻を更新
      if (messages.length > 0) {
        // Slack APIは新しい順で返すので、最初のメッセージが最新
        const latestTs = messages[0].ts;
        console.log(`最新メッセージのタイムスタンプ: ${latestTs}`);
        updateLastSyncTime(syncSheet, channelId, latestTs);
      } else {
        // 新規メッセージがない場合も現在時刻で更新（重複チェックを避けるため）
        const nowTs = Math.floor(Date.now() / 1000).toString();
        console.log(`新規メッセージなし。現在時刻で更新: ${nowTs}`);
        updateLastSyncTime(syncSheet, channelId, nowTs);
      }
      
      logInfo(`Channel ${channelName}: ${messages.length}件のメッセージを同期`);
      
    } catch (error) {
      logError(`Channel ${channelId} sync`, error.toString());
    }
  });
  
  logInfo(`Slack同期完了: 全${totalMessageCount}件のメッセージ`);
}

// ========= チャンネル情報取得 =========
function getAllPublicChannels() {
  try {
    const response = slackAPI('conversations.list', {
      types: 'public_channel',
      limit: 200,
      exclude_archived: true
    });
    
    return response.channels || [];
  } catch (error) {
    logError('Get channels', error.toString());
    return [];
  }
}

// ========= Botが参加しているすべてのチャンネルを取得 =========
function getAllJoinedChannels() {
  try {
    console.log('Botが参加しているチャンネルを取得中...');
    
    const joinedChannels = [];
    
    // 1. まずBotのユーザーIDを取得
    const authInfo = slackAPI('auth.test', {});
    const botUserId = authInfo.user_id;
    console.log(`Bot User ID: ${botUserId}`);
    
    // 2. 推奨: users.conversations でBotが参加しているチャンネルのみ取得
    try {
      let cursorUser = '';
      do {
        const paramsUser = {
          types: 'public_channel,private_channel',
          limit: 1000,
          exclude_archived: true,
          user: botUserId
        };
        if (cursorUser) paramsUser.cursor = cursorUser;
        
        console.log('users.conversations APIを呼び出し中...');
        const userConv = slackAPI('users.conversations', paramsUser);
        if (userConv.ok && userConv.channels) {
          joinedChannels.push(...userConv.channels);
          cursorUser = userConv.response_metadata?.next_cursor || '';
                } else {
          console.warn('users.conversations失敗、conversations.listにフォールバック');
          throw new Error(userConv?.error || 'users.conversations failed');
        }
      } while (cursorUser);
      
      // チャンネル名でソート
      joinedChannels.sort((a, b) => (a.name || '').localeCompare(b.name || ''));
      console.log(`取得したチャンネル総数: ${joinedChannels.length}`);
      return joinedChannels;
    } catch (userConvError) {
      console.warn(`users.conversationsでの取得に失敗: ${userConvError.toString()}`);
      console.warn('conversations.list にフォールバックします');
    }
    
    // 3. フォールバック: conversations.list から is_member=true のみ採用
    let cursor = '';
    do {
      const params = {
        types: 'public_channel,private_channel',
        limit: 1000,
        exclude_archived: true
      };
      if (cursor) params.cursor = cursor;
      
      console.log('conversations.list APIを呼び出し中...(fallback)');
      const response = slackAPI('conversations.list', params);
      if (response.ok && response.channels) {
        response.channels.forEach(channel => {
          if (channel.is_member) joinedChannels.push(channel);
        });
        cursor = response.response_metadata?.next_cursor || '';
      } else {
        console.error('チャンネル取得エラー:', response.error);
        break;
      }
    } while (cursor);
    
    // チャンネル名でソート
    joinedChannels.sort((a, b) => (a.name || '').localeCompare(b.name || ''));
    console.log(`取得したチャンネル総数: ${joinedChannels.length}`);
    return joinedChannels;
  } catch (error) {
    console.error('チャンネル一覧取得エラー:', error);
    return [];
  }
}

// チャンネル情報キャッシュのグローバル変数
let channelInfoCache = {};
let cacheExpiry = 0;

function getChannelInfo(channelId) {
  try {
    // キャッシュを確認
    const now = Date.now();
    if (channelInfoCache[channelId] && cacheExpiry > now) {
      console.log(`キャッシュからチャンネル情報を取得: ${channelId}`);
      return channelInfoCache[channelId];
    }
    
    // Rate limit対策: API呼び出し前に少し待機
    Utilities.sleep(200); // 200ms待機
    
    // まずconversations.listから情報を取得を試みる（エラーが少ない）
    try {
      const listResponse = slackAPI('conversations.list', {
        types: 'public_channel,private_channel',
        limit: 1000
      });
      
      // 全チャンネルをキャッシュに保存（5分間）
      if (listResponse.channels) {
        listResponse.channels.forEach(ch => {
          channelInfoCache[ch.id] = ch;
        });
        cacheExpiry = now + 5 * 60 * 1000; // 5分後に期限切れ
      }
      
      const channel = listResponse.channels?.find(ch => ch.id === channelId);
      if (channel) {
        return channel;
      }
    } catch (listError) {
      console.warn(`conversations.list失敗: ${listError.toString()}`);
    }
    
    // conversations.listで見つからない場合のみconversations.infoを試す
    // ただし、invalid_argumentsエラーが多いので、エラーは警告レベルで処理
    try {
      const response = slackAPI('conversations.info', {
        channel: channelId
      });
      if (response.channel) {
        channelInfoCache[channelId] = response.channel;
        return response.channel;
      }
    } catch (infoError) {
      // conversations.infoのエラーは予想されるので警告レベルで記録
      console.warn(`チャンネル情報取得スキップ (${channelId}): conversations.info失敗`);
    }
    
    // どちらも失敗した場合はnullを返す
    return null;
  } catch (error) {
    // Rate limitエラーの場合は待機してリトライ
    if (error.toString().includes('ratelimited')) {
      console.log('Slack API rate limitに達しました。1秒待機してリトライ...');
      Utilities.sleep(1000);
      return getChannelInfo(channelId);
    }
    
    console.error(`チャンネル情報取得エラー: ${error.toString()}`);
    return null;
  }
}

// ========= メール重複チェック機能 =========

// メール重複チェック
function isDuplicateEmail(emailType, subject, content) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  
  // 送信履歴シートを取得または作成
  let sentLogSheet = ss.getSheetByName('SentEmailLog');
  if (!sentLogSheet) {
    sentLogSheet = ss.insertSheet('SentEmailLog');
    sentLogSheet.getRange(1, 1, 1, 4).setValues([['送信日時', 'タイプ', 'タイトル', 'ハッシュ']]);
  }
  
  // コンテンツのハッシュを生成
  const contentHash = Utilities.computeDigest(
    Utilities.DigestAlgorithm.MD5,
    JSON.stringify(content) + subject
  ).map(byte => (byte < 0 ? byte + 256 : byte).toString(16).padStart(2, '0')).join('');
  
  // 過去2時間以内に同じ内容のメールを送信していないかチェック
  const twoHoursAgo = new Date(new Date().getTime() - 2 * 60 * 60 * 1000);
  const sentLogs = sentLogSheet.getDataRange().getValues();
  
  for (let i = 1; i < sentLogs.length; i++) {
    const sentDate = sentLogs[i][0];
    const sentType = sentLogs[i][1];
    const sentHash = sentLogs[i][3];
    
    if (sentDate instanceof Date && sentDate > twoHoursAgo && 
        sentType === emailType && sentHash === contentHash) {
      return true; // 重複あり
    }
  }
  
  return false; // 重複なし
}

// メール送信履歴を記録
function recordEmailSent(emailType, subject, content) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sentLogSheet = ss.getSheetByName('SentEmailLog');
  
  if (!sentLogSheet) {
    sentLogSheet = ss.insertSheet('SentEmailLog');
    sentLogSheet.getRange(1, 1, 1, 4).setValues([['送信日時', 'タイプ', 'タイトル', 'ハッシュ']]);
  }
  
  // コンテンツのハッシュを生成
  const contentHash = Utilities.computeDigest(
    Utilities.DigestAlgorithm.MD5,
    JSON.stringify(content) + subject
  ).map(byte => (byte < 0 ? byte + 256 : byte).toString(16).padStart(2, '0')).join('');
  
  // 送信履歴を記録
  sentLogSheet.appendRow([new Date(), emailType, subject, contentHash]);
  
  // 古い送信履歴を削除（30日以上前のものを削除）
  const thirtyDaysAgo = new Date(new Date().getTime() - 30 * 24 * 60 * 60 * 1000);
  const dataRange = sentLogSheet.getDataRange();
  const values = dataRange.getValues();
  
  for (let i = values.length - 1; i >= 1; i--) {
    if (values[i][0] instanceof Date && values[i][0] < thirtyDaysAgo) {
      sentLogSheet.deleteRow(i + 1);
    }
  }
}

// ========= 詳細なメッセージ履歴取得 =========
function fetchChannelHistory(channelId, oldestOrDays = '0') {
  // パラメータのバリデーション
  if (!channelId) {
    logError('fetchChannelHistory', 'チャンネルIDが指定されていません');
    return [];
  }
  
  // oldestOrDaysが数値の場合は日数として扱い、タイムスタンプに変換
  let oldest = oldestOrDays;
  if (typeof oldestOrDays === 'number' && oldestOrDays > 0) {
    const date = new Date();
    date.setDate(date.getDate() - oldestOrDays);
    oldest = Math.floor(date.getTime() / 1000).toString();
  } else if (!oldestOrDays || oldestOrDays === '0' || oldestOrDays === 0) {
    oldest = '0';
  } else {
    // 文字列の場合、有効な数値か確認
    oldest = String(oldestOrDays);
    if (isNaN(parseFloat(oldest))) {
      console.warn(`無効なタイムスタンプ: ${oldestOrDays}、0を使用します`);
      oldest = '0';
    }
  }
  
  console.log(`fetchChannelHistory: channelId=${channelId}, oldest=${oldest}`);
  
  try {
    // 設定値の確認
    const limit = MAX_MESSAGES_PER_CHANNEL || 100;
    console.log(`メッセージ取得上限: ${limit}`);
    
    const response = slackAPI('conversations.history', {
      channel: channelId,
      oldest: oldest,
      limit: limit,
      inclusive: false
    });
    
    // レスポンスの検証
    if (!response || !response.ok) {
      throw new Error(`Slack APIレスポンスエラー: ${response?.error || 'unknown error'}`);
    }
    
    const messages = response.messages || [];
    console.log(`取得したメッセージ数: ${messages.length}`);
    
    // メッセージが空の場合の処理
    if (messages.length === 0) {
      console.log('新しいメッセージはありません');
      return [];
    }
    
    // メッセージを古い順に並べ替え（Slack APIは新しい順で返す）
    messages.reverse();
    
    return messages;
  } catch (error) {
    logError('fetchChannelHistory', `チャンネル ${channelId} のメッセージ取得エラー: ${error.toString()}`);
    
    // エラーの詳細を出力
    if (error.toString().includes('not_in_channel')) {
      console.error('Botがチャンネルに参加していません。Botをチャンネルに招待してください。');
    } else if (error.toString().includes('invalid_auth')) {
      console.error('認証エラー: Bot Tokenが無効です。');
    } else if (error.toString().includes('channel_not_found')) {
      console.error('チャンネルが見つかりません: ' + channelId);
    }
    
    return [];
  }
}

function fetchChannelHistoryWithDetails(channelId, oldest = '0') {
  // oldestパラメータの検証と正規化
  let validOldest = oldest;
  if (!oldest || oldest === '0' || oldest === 0) {
    validOldest = '0';
  } else {
    validOldest = String(oldest);
    if (isNaN(parseFloat(validOldest))) {
      console.warn(`無効なタイムスタンプ: ${oldest}、0を使用します`);
      validOldest = '0';
    }
  }
  
  console.log(`fetchChannelHistoryWithDetails: channelId=${channelId}, oldest=${validOldest}`);
  
  try {
    const response = slackAPI('conversations.history', {
      channel: channelId,
      oldest: validOldest,
      limit: MAX_MESSAGES_PER_CHANNEL || 100,
      inclusive: false  // oldestのタイムスタンプのメッセージ自体は含めない
    });
    
    if (!response.ok) {
      console.error(`チャンネル履歴取得エラー: ${response.error}`);
      
      // エラーの種類に応じた対処
      if (response.error === 'channel_not_found') {
        console.error(`チャンネル ${channelId} が見つかりません`);
      } else if (response.error === 'not_in_channel') {
        console.error(`Botはチャンネル ${channelId} のメンバーではありません`);
      } else if (response.error === 'missing_scope') {
        console.error(`権限不足: チャンネル ${channelId} の履歴を取得するには適切な権限が必要です`);
      }
      
      return [];
    }
    
    const messages = response.messages || [];
    
    // 各メッセージの詳細情報を追加（最適化版）
    return messages.map(message => {
      // ユーザー情報取得をスキップ（パフォーマンス向上）
      const userName = ENABLE_USER_INFO_FETCH ? getUserInfo(message.user).name : (message.user || 'unknown');
      
      // リアクション情報を整形（軽量化）
      const reactions = message.reactions && message.reactions.length > 0 ? 
        `${message.reactions.length} reactions` : '';
      
      // ファイル情報を整形（軽量化）
      const files = message.files && message.files.length > 0 ? 
        `${message.files.length} files` : '';
      
      return {
        ...message,
        user_name: userName,
        reactions: reactions,
        files: files
      };
    });
  } catch (error) {
    console.error(`fetchChannelHistoryWithDetails エラー: ${error.toString()}`);
    return [];
  }
}

function fetchThreadReplies(channelId, threadTs) {
  try {
    // Rate limit対策
    Utilities.sleep(100);
    
    const response = slackAPI('conversations.replies', {
      channel: channelId,
      ts: threadTs,
      limit: 10  // 最大10件に制限
    });
    
    const messages = response.messages || [];
    
    // 最初のメッセージ（親）をスキップ
    if (messages.length <= 1) {
      return [];
    }
    
    return messages.slice(1).map(message => {
      return {
        ...message,
        user_name: message.user || 'unknown'
      };
    });
  } catch (error) {
    // エラーは呼び出し元で処理するため、ここでは再スローのみ
    throw error;
  }
}

// ========= バッチ保存用ヘルパー関数 =========
function prepareMessageRow(channelId, message) {
  const messageId = `${channelId}_${message.ts}`;
  const permalink = `https://slack.com/archives/${channelId}/p${message.ts.replace('.', '')}`;
  const messageDate = new Date(Number(message.ts.split('.')[0]) * 1000);
  
  // ユーザー情報を取得
  const userInfo = getUserInfo(message.user);
  const userName = userInfo.real_name || userInfo.name || message.user || 'unknown';
  const userEmail = userInfo.email || '';
  
  return [
    messageId,  // id
    channelId,  // channel_id
    message.ts, // message_ts
    message.thread_ts || '', // thread_ts
    message.text || '', // text_raw
    `${userName}${userEmail ? ' (' + userEmail + ')' : ''}`, // user_name (名前とメール)
    '', // summary_json
    '', // classification_json
    '', // match_flag
    '', // human_judgement
    permalink, // permalink
    '', // checklist_proposed
    '', // agenda_selected
    '', // draft_doc_url
    messageDate.toISOString(), // timestamp
    message.reactions || '', // reactions
    message.files || '' // files
  ];
}

function prepareSlackLogRow(channelId, channelName, message) {
  const messageDate = new Date(Number(message.ts.split('.')[0]) * 1000);
  
  // ユーザー情報を取得
  const userInfo = getUserInfo(message.user);
  const userName = userInfo.real_name || userInfo.name || message.user || 'unknown';
  const userEmail = userInfo.email || '';
  
  return [
    channelId,
    channelName,
    message.ts,
    message.thread_ts || '',
    `${userName}${userEmail ? ' (' + userEmail + ')' : ''}`, // ユーザー名とメール
    message.text || '',
    messageDate,
    message.reactions || '',
    message.files || ''
  ];
}

function saveMessagesBatch(sheet, rows) {
  if (rows.length === 0) return;
  
  // 既存データの最終行を取得
  const lastRow = sheet.getLastRow();
  
  // 重複チェック用に既存のメッセージIDを取得
  const existingIds = new Set();
  if (lastRow > 1) {
    const existingData = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    existingData.forEach(row => existingIds.add(row[0]));
  }
  
  // 重複を除いた新規メッセージのみをフィルタ
  const newRows = rows.filter(row => {
    const messageId = row[0];
    if (existingIds.has(messageId)) {
      console.log(`重複メッセージをスキップ: ${messageId}`);
      return false;
    }
    return true;
  });
  
  // バッチで一括挿入
  if (newRows.length > 0) {
    sheet.getRange(lastRow + 1, 1, newRows.length, newRows[0].length).setValues(newRows);
    console.log(`Messagesシートに${newRows.length}件を保存（${rows.length - newRows.length}件の重複をスキップ）`);
  } else {
    console.log(`すべてのメッセージ（${rows.length}件）は既に保存済みです`);
  }
}

function saveSlackLogBatch(sheet, rows) {
  if (rows.length === 0) {
    console.log('保存するSlackログがありません');
    return;
  }
  
  if (!sheet) {
    console.error('slack_logシートが存在しません');
    return;
  }
  
  console.log(`slack_logシートへの保存を開始: ${rows.length}件`);
  
  // 既存データの最終行を取得
  const lastRow = sheet.getLastRow();
  
  // 重複チェック用に既存のタイムスタンプを取得
  const existingTs = new Set();
  if (lastRow > 1) {
    try {
      const existingData = sheet.getRange(2, 3, lastRow - 1, 1).getValues();
      existingData.forEach(row => {
        if (row[0]) existingTs.add(row[0].toString());
      });
      console.log(`既存のログ数: ${existingTs.size}件`);
    } catch (e) {
      console.warn('既存ログの読み込みエラー:', e.toString());
    }
  }
  
  // 重複を除いた新規メッセージのみをフィルタ
  const newRows = rows.filter(row => {
    const ts = row[2];
    if (existingTs.has(ts)) {
      console.log(`重複ログをスキップ: ${ts}`);
      return false;
    }
    return true;
  });
  
  // バッチで一括挿入
  if (newRows.length > 0) {
    try {
      sheet.getRange(lastRow + 1, 1, newRows.length, newRows[0].length).setValues(newRows);
      console.log(`✅ slack_logシートに${newRows.length}件を保存（${rows.length - newRows.length}件の重複をスキップ）`);
    } catch (e) {
      console.error('slack_logシートへの保存エラー:', e.toString());
    }
  } else {
    console.log(`すべてのログ（${rows.length}件）は既に保存済みです`);
  }
}

// ========= slack_logシートへの保存 =========
function saveMessageToSlackLog(sheet, channelId, channelName, message) {
  const messageDate = new Date(Number(message.ts.split('.')[0]) * 1000);
  
  // 重複チェック
  const existingData = sheet.getDataRange().getValues();
  for (let i = 1; i < existingData.length; i++) {
    if (existingData[i][0] === channelId && existingData[i][2] === message.ts) {
      return; // 既に存在
    }
  }
  
  const lastRow = sheet.getLastRow();
  sheet.getRange(lastRow + 1, 1, 1, 9).setValues([[
    channelId,
    channelName,
    message.ts,
    message.thread_ts || '',
    message.user_name || '',
    message.text || '',
    messageDate,
    message.reactions || '',
    message.files || ''
  ]]);
}

// ========= 過去ログ分析 =========
function analyzeHistoricalMessages() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const slackLogSheet = ss.getSheetByName(SHEETS.SLACK_LOG);
  
  if (!slackLogSheet) {
    SpreadsheetApp.getUi().alert('slack_logシートが見つかりません');
    return;
  }
  
  // 過去7日間のメッセージを取得
  const now = new Date();
  const weekAgo = new Date(now.getTime() - 7 * 24 * 60 * 60 * 1000);
  
  const data = slackLogSheet.getDataRange().getValues();
  const messages = [];
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const date = row[6]; // date列
    
    if (date instanceof Date && date >= weekAgo) {
      messages.push({
        channel_id: row[0],
        channel_name: row[1],
        ts: row[2],
        thread_ts: row[3],
        user_name: row[4],
        message: row[5],
        date: row[6],
        reactions: row[7],
        files: row[8]
      });
    }
  }
  
  logInfo(`過去7日間のメッセージ: ${messages.length}件を分析開始`);
  
  // チャンネルごとに分析
  const channelGroups = groupMessagesByChannel(messages);
  const analysisResults = [];
  
  for (const [channelName, channelMessages] of Object.entries(channelGroups)) {
    const analysis = analyzeChannelMessages(channelName, channelMessages);
    analysisResults.push(analysis);
  }
  
  // 分析結果を保存
  saveAnalysisResults(ss, analysisResults);
  
  SpreadsheetApp.getUi().alert(`分析完了: ${messages.length}件のメッセージを処理しました`);
}

// ========= メッセージのグループ化と分析 =========
function groupMessagesByChannel(messages) {
  const groups = {};
  
  messages.forEach(msg => {
    const channel = msg.channel_name || msg.channel_id;
    if (!groups[channel]) {
      groups[channel] = [];
    }
    groups[channel].push(msg);
  });
  
  return groups;
}

function analyzeChannelMessages(channelName, messages) {
  // トピック分類
  const topics = classifyMessagesByTopic(messages);
  
  // 統計情報
  const stats = {
    totalMessages: messages.length,
    topicCount: topics.length,
    participants: [...new Set(messages.map(m => m.user_name))],
    mostActiveUser: getMostActiveUser(messages),
    peakHours: getPeakActivityHours(messages)
  };
  
  // 重要なトピックを特定
  const importantTopics = topics.filter(topic => {
    const context = analyzeMessageContext(topic);
    return context.hasDecision || context.hasInstruction || 
           (context.hasQuestion && topic.length > 3);
  });
  
  return {
    channel: channelName,
    stats: stats,
    importantTopics: importantTopics.length,
    topics: topics.map(t => ({
      messageCount: t.length,
      context: analyzeMessageContext(t)
    }))
  };
}

// ========= メッセージ分析ヘルパー関数 =========
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
  
  return topics;
}

function analyzeMessageContext(messages) {
  const keywords = new Set();
  const participants = new Set();
  let hasQuestion = false;
  let hasDecision = false;
  let hasInstruction = false;
  let hasTroubleshooting = false;
  
  for (const msg of messages) {
    // 参加者を記録
    if (msg.user_name) participants.add(msg.user_name);
    
    // メッセージの特徴を分析
    const text = (msg.message || msg.text || '').toLowerCase();
    
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
    
    // キーワード抽出
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
    estimatedType: determineDocumentType(hasQuestion, hasDecision, hasInstruction, hasTroubleshooting)
  };
}

function determineDocumentType(hasQuestion, hasDecision, hasInstruction, hasTroubleshooting) {
  if (hasTroubleshooting) return 'TROUBLESHOOTING';
  if (hasQuestion) return 'FAQ';
  if (hasDecision) return 'DECISION';
  if (hasInstruction) return 'PROCEDURE';
  return 'INFORMATION';
}

function getMostActiveUser(messages) {
  const userCounts = {};
  
  messages.forEach(msg => {
    const user = msg.user_name || 'Unknown';
    userCounts[user] = (userCounts[user] || 0) + 1;
  });
  
  const sorted = Object.entries(userCounts).sort((a, b) => b[1] - a[1]);
  return sorted[0] ? sorted[0][0] : null;
}

function getPeakActivityHours(messages) {
  const hourCounts = {};
  
  messages.forEach(msg => {
    if (msg.date instanceof Date) {
      const hour = msg.date.getHours();
      hourCounts[hour] = (hourCounts[hour] || 0) + 1;
    }
  });
  
  const sorted = Object.entries(hourCounts).sort((a, b) => b[1] - a[1]);
  return sorted.slice(0, 3).map(([hour, count]) => `${hour}時`);
}

// ========= マニュアル・FAQ生成 =========
function generateManualAndFAQFromMessages() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const slackLogSheet = ss.getSheetByName(SHEETS.SLACK_LOG);
  
  if (!slackLogSheet) {
    SpreadsheetApp.getUi().alert('slack_logシートが見つかりません');
    return;
  }
  
  // 過去24時間のメッセージを取得
  const now = new Date();
  const yesterday = new Date(now.getTime() - 24 * 60 * 60 * 1000);
  
  const data = slackLogSheet.getDataRange().getValues();
  const messages = [];
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const date = row[6]; // date列
    
    if (date instanceof Date && date >= yesterday) {
      messages.push({
        channel_id: row[0],
        channel_name: row[1],
        ts: row[2],
        thread_ts: row[3],
        user_name: row[4],
        message: row[5],
        date: row[6]
      });
    }
  }
  
  logInfo(`過去24時間のメッセージ: ${messages.length}件`);
  
  if (messages.length === 0) {
    SpreadsheetApp.getUi().alert('過去24時間にメッセージがありません');
    return;
  }
  
  // トピック別に分類
  const topics = classifyMessagesByTopic(messages);
  
  let totalManuals = 0;
  let totalFAQs = 0;
  
  // 各トピックを処理
  topics.forEach(topicMessages => {
    const context = analyzeMessageContext(topicMessages);
    
    if (context.estimatedType === 'FAQ' || context.hasQuestion) {
      const faq = generateFAQFromTopic(topicMessages, context);
      if (faq) {
        saveFAQToSheet(ss, faq, topicMessages);
        totalFAQs++;
      }
    }
    
    if (context.estimatedType !== 'FAQ' && topicMessages.length > 2) {
      const manual = generateManualFromTopic(topicMessages, context);
      if (manual) {
        saveManualToSheet(ss, manual, topicMessages);
        totalManuals++;
      }
    }
  });
  
  SpreadsheetApp.getUi().alert(
    `生成完了\nマニュアル: ${totalManuals}件\nFAQ: ${totalFAQs}件`
  );
}

// ========= OpenAI API 連携（議題生成用） =========
function callOpenAIForAgenda(messages, model = 'gpt-5', responseFormat = null) {
  const url = 'https://api.openai.com/v1/responses';
  if (!OPENAI_API_KEY) {
    console.error('OpenAI APIキーが設定されていません');
    throw new Error('OpenAI APIキーが設定されていません');
  }
  const combinedInput = messages.map(m => {
    const content = (typeof m.content === 'string') ? m.content : JSON.stringify(m.content);
    return `${String(m.role || 'user').toUpperCase()}: ${content}`;
  }).join("\n\n");

  const payload = {
    model: model,
    input: combinedInput,
    max_output_tokens: 2000
  };
  if (responseFormat && responseFormat.type === 'json_object') {
    payload.text = { format: { type: 'json_schema', name: 'agenda', schema: { type: 'object', properties: {}, additionalProperties: true }, strict: false } };
  }

  const options = {
    method: 'post',
    headers: {
      'Authorization': `Bearer ${OPENAI_API_KEY}`,
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const responseText = response.getContentText();
    const data = JSON.parse(responseText);
    if (data.error) throw new Error(`OpenAI API Error: ${data.error.message}`);
    const content = extractTextFromOpenAIResponse_A_(data);
    if (responseFormat && responseFormat.type === 'json_object') {
      try { return JSON.parse(content); } catch (e) { return { summary: '要約を生成できませんでした', categories: [], error: 'JSONパースエラー' }; }
    }
    return content;
  } catch (error) {
    logError('OpenAI API', error.toString());
    throw error;
  }
}

// ========= AI分析処理 =========
function runAIAnalysis() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const messagesSheet = ss.getSheetByName(SHEETS.MESSAGES);
  
  if (!messagesSheet) {
    ui.alert('エラー', 'Messagesシートが見つかりません。先にメッセージを同期してください。', ui.ButtonSet.OK);
    return;
  }
  
  // 未分析のメッセージを取得
  const lastRow = messagesSheet.getLastRow();
  if (lastRow <= 1) {
    ui.alert('情報', '分析対象のメッセージがありません。', ui.ButtonSet.OK);
    return;
  }
  
  // summary_jsonが空のメッセージを探す
  const data = messagesSheet.getRange(2, 1, lastRow - 1, 16).getValues();
  const unanalyzedMessages = [];
  
  data.forEach((row, index) => {
    if (!row[6] || row[6] === '') { // summary_json列が空
      unanalyzedMessages.push({
        rowIndex: index + 2,
        messageId: row[0],
        channelId: row[1],
        text: row[4]
      });
    }
  });
  
  if (unanalyzedMessages.length === 0) {
    ui.alert('情報', '全てのメッセージは既に分析済みです。', ui.ButtonSet.OK);
    return;
  }
  
  ui.alert('分析開始', `${unanalyzedMessages.length}件のメッセージを分析します。`, ui.ButtonSet.OK);
  
  let analyzedCount = 0;
  let errorCount = 0;
  
  // バッチ処理で分析
  const batchSize = 5;
  for (let i = 0; i < unanalyzedMessages.length; i += batchSize) {
    const batch = unanalyzedMessages.slice(i, Math.min(i + batchSize, unanalyzedMessages.length));
    
    batch.forEach(msg => {
      try {
        if (!msg.text || msg.text.trim() === '') {
          console.log(`メッセージ ${msg.messageId} はテキストがありません`);
          return;
        }
        
        // メッセージを分析
        const summary = summarizeMessage(msg.text);
        
        // 結果を保存
        messagesSheet.getRange(msg.rowIndex, 7).setValue(JSON.stringify(summary)); // summary_json
        messagesSheet.getRange(msg.rowIndex, 16).setValue(new Date()); // updated_at
        
        analyzedCount++;
        console.log(`メッセージ ${msg.messageId} を分析完了`);
        
      } catch (error) {
        console.error(`メッセージ ${msg.messageId} の分析エラー:`, error);
        errorCount++;
      }
    });
    
    // API制限対策
    if (i + batchSize < unanalyzedMessages.length) {
      Utilities.sleep(1000); // 1秒待機
    }
  }
  
  const resultMessage = `
分析完了！

✅ 成功: ${analyzedCount}件
❌ エラー: ${errorCount}件

Messagesシートの分析結果を確認してください。
  `;
  
  ui.alert('分析結果', resultMessage, ui.ButtonSet.OK);
  logInfo(`AI分析完了: ${analyzedCount}件のメッセージを処理`);
}

function summarizeMessage(text) {
  const prompt = `
以下のSlackメッセージを要約してください。
必ず以下のJSON形式で出力してください。JSON以外のテキストは含めないでください：

{
  "summary": "要約（100文字以内）",
  "decisions": ["決定事項1", "決定事項2"],
  "action_items": [
    {
      "owner": "担当者名",
      "task": "タスク内容",
      "due": "期限"
    }
  ],
  "people": ["関係者1", "関係者2"],
  "dates": ["言及された日付1", "言及された日付2"]
}

メッセージ:
${text}

重要: 必ず有効なJSON形式で出力してください。`;

  try {
    const response = callOpenAIForAgenda([
      { role: 'system', content: 'あなたは議事録作成アシスタントです。必ず有効なJSON形式で応答してください。' },
      { role: 'user', content: prompt }
    ], 'gpt-5', { type: 'json_object' });
    
    // responseがすでにJSONオブジェクトの場合はそのまま返す
    if (typeof response === 'object') {
      return response;
    }
    
    // 文字列の場合はパース
    return JSON.parse(response);
  } catch (error) {
    console.error('要約生成エラー:', error.toString());
    // エラー時はデフォルト値を返す
    return {
      summary: text.substring(0, 100),
      decisions: [],
      action_items: [],
      people: [],
      dates: []
    };
  }
}

function classifyMessage(summary, categories) {
  const categoriesText = categories.map(c => 
    `- ${c.name}: ${c.description}\n  判定基準: ${c.criteria}`
  ).join('\n');
  
  const prompt = `
以下の要約を各カテゴリに分類し、該当度をスコア（0-1）で評価してください。
必ずJSON配列形式で出力してください。JSON以外のテキストは含めないでください。

カテゴリ:
${categoriesText}

要約:
${JSON.stringify(summary)}

出力形式（JSON配列）:
[
  {
    "category": "カテゴリ名",
    "score": 0.8,
    "rationale": "判定理由",
    "key_quotes": ["関連する引用"]
  }
]

重要: 必ず有効なJSON配列形式で出力してください。`;

  try {
    const response = callOpenAIForAgenda([
      { role: 'system', content: 'あなたはガバナンス判定アシスタントです。必ず有効なJSON配列形式で応答してください。' },
      { role: 'user', content: prompt }
    ], 'gpt-5', { type: 'json_object' });
    
    // responseがすでに配列の場合はそのまま返す
    if (Array.isArray(response)) {
      return response;
    }
    
    // オブジェクトでcategoriesプロパティがある場合
    if (typeof response === 'object' && response.categories) {
      return response.categories;
    }
    
    // 文字列の場合はパース
    if (typeof response === 'string') {
      return JSON.parse(response);
    }
    
    // その他の場合は空配列を返す
    return [];
  } catch (error) {
    console.error('分類エラー:', error.toString());
    // エラー時は空配列を返す
    return [];
  }
}

// ========= FAQ生成 =========
function generateFAQFromTopic(messages, context) {
  const conversationText = formatMessagesForAI(messages);
  
  const prompt = `
以下の会話から、1つの明確な質問と回答を抽出してください。

出力形式（JSON）:
{
  "question": "ユーザーの質問を明確に",
  "answer": "簡潔で分かりやすい回答",
  "category": "適切なカテゴリ",
  "tags": "関連キーワード（カンマ区切り）",
  "supplement": "必要に応じて追加情報"
}

会話内容：
${conversationText}

注意：
- 質問と回答は1対1で明確にしてください
- 回答は実用的で具体的にしてください
`;
  
  try {
    const response = callOpenAIForAgenda([
      { role: 'system', content: 'FAQ作成の専門家として、明確で有用なQ&Aを作成してください。' },
      { role: 'user', content: prompt }
    ], 'gpt-5', { type: 'json_object' });
    
    return JSON.parse(response);
  } catch (error) {
    logError('FAQ生成エラー', error.toString());
    return null;
  }
}

// ========= マニュアル生成 =========
function generateManualFromTopic(messages, context) {
  const conversationText = formatMessagesForAI(messages);
  
  let promptType = '';
  switch (context.estimatedType) {
    case 'TROUBLESHOOTING':
      promptType = 'トラブルシューティング手順';
      break;
    case 'DECISION':
      promptType = '意思決定記録';
      break;
    case 'PROCEDURE':
      promptType = '作業手順書';
      break;
    default:
      promptType = '業務情報';
  }
  
  const prompt = `
以下の会話から、${promptType}を作成してください。

出力形式（JSON）:
{
  "category": "${promptType}",
  "title": "内容を表す明確なタイトル",
  "content": "詳細な内容（手順がある場合は番号付きリスト）",
  "keywords": "関連キーワード（カンマ区切り）",
  "importance": "high/medium/low"
}

会話内容：
${conversationText}

注意：
- 1つの独立したトピックとして完結させてください
- 具体的で実用的な内容にしてください
`;
  
  try {
    const response = callOpenAIForAgenda([
      { role: 'system', content: '業務文書作成の専門家として、実用的な文書を作成してください。' },
      { role: 'user', content: prompt }
    ], 'gpt-5', { type: 'json_object' });
    
    return JSON.parse(response);
  } catch (error) {
    logError('マニュアル生成エラー', error.toString());
    return null;
  }
}

// ========= Slackメッセージから議題抽出＆業務フロー生成・メール送信 =========
function analyzeSlackAndSendReport() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const slackLogSheet = ss.getSheetByName(SHEETS.SLACK_LOG);
  const messagesSheet = ss.getSheetByName(SHEETS.MESSAGES);
  
  // 1. Slackから最新メッセージを同期
  syncAllSlackChannels();
  
  // 2. 過去24時間のSlackメッセージを取得
  const now = new Date();
  const yesterday = new Date(now.getTime() - 24 * 60 * 60 * 1000);
  const recentMessages = getRecentSlackMessages(slackLogSheet, yesterday, now);
  
  if (recentMessages.length === 0) {
    logInfo('分析対象のSlackメッセージがありません');
    return;
  }
  
  // 3. メッセージから議題・論点を抽出
  const agendaItems = extractAgendasFromSlackMessages(recentMessages);
  
  // 4. 業務フローチャートを生成
  const flowchart = generateBusinessFlowchart(agendaItems);
  
  // 5. レポート生成
  const report = {
    date: now,
    messageCount: recentMessages.length,
    agendaItems: agendaItems,
    flowchart: flowchart,
    summary: summarizeAgendaItems(agendaItems)
  };
  
  // 6. 新規スプレッドシートにレポートをエクスポート
  const newSpreadsheetUrl = exportReportToNewSpreadsheet(report);
  report.spreadsheetUrl = newSpreadsheetUrl;
  
  // 7. メール送信（業務フローチャート付き）
  if (REPORT_EMAIL) {
    sendAgendaReportWithFlowchart(report);
  }
  
  // 8. マスタースプレッドシートに記録
  recordAgendaAnalysis(ss, report);
  
  logInfo(`Slack議題分析完了: ${agendaItems.length}件の議題を抽出`);
}

// 日次レポート送信のエイリアス（既存のトリガーとの互換性のため）
function sendDailyReport() {
  analyzeSlackAndSendReport();
}

function collectMessageStats(sheet, startDate, endDate) {
  if (!sheet) return null;
  
  const data = sheet.getDataRange().getValues();
  const stats = {
    totalMessages: 0,
    channels: new Set(),
    users: new Set(),
    topChannels: {},
    topUsers: {},
    hourlyDistribution: {}
  };
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const date = row[6]; // date列
    
    if (date instanceof Date && date >= startDate && date <= endDate) {
      stats.totalMessages++;
      stats.channels.add(row[1]); // channel_name
      stats.users.add(row[4]); // user_name
      
      // チャンネル別カウント
      const channel = row[1];
      stats.topChannels[channel] = (stats.topChannels[channel] || 0) + 1;
      
      // ユーザー別カウント
      const user = row[4];
      stats.topUsers[user] = (stats.topUsers[user] || 0) + 1;
      
      // 時間帯別
      const hour = date.getHours();
      stats.hourlyDistribution[hour] = (stats.hourlyDistribution[hour] || 0) + 1;
    }
  }
  
  // Top 5を抽出
  stats.topChannels = Object.entries(stats.topChannels)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 5);
  
  stats.topUsers = Object.entries(stats.topUsers)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 5);
  
  stats.channelCount = stats.channels.size;
  stats.userCount = stats.users.size;
  
  return stats;
}

function getSheetStats(sheet, startDate, endDate) {
  const data = sheet.getDataRange().getValues();
  let count = 0;
  
  for (let i = 1; i < data.length; i++) {
    const date = data[i][0]; // 最初の列が日時と仮定
    if (date instanceof Date && date >= startDate && date <= endDate) {
      count++;
    }
  }
  
  return { newItems: count };
}

function generateDailyReport(data) {
  const reportDate = Utilities.formatDate(data.date, 'JST', 'yyyy年MM月dd日');
  
  let report = {
    title: `Slack活動日次レポート - ${reportDate}`,
    sections: []
  };
  
  // メッセージ統計セクション
  if (data.messageStats) {
    const stats = data.messageStats;
    report.sections.push({
      title: '📊 メッセージ統計',
      content: `
総メッセージ数: ${stats.totalMessages}件
アクティブチャンネル: ${stats.channelCount}
アクティブユーザー: ${stats.userCount}

【最も活発なチャンネル】
${stats.topChannels.map(([ch, count]) => `• ${ch}: ${count}件`).join('\n')}

【最も活発なユーザー】
${stats.topUsers.map(([user, count]) => `• ${user}: ${count}件`).join('\n')}
      `.trim()
    });
  }
  
  // 議題候補セクション
  if (data.agendaCandidates && data.agendaCandidates.length > 0) {
    report.sections.push({
      title: '📋 議題候補',
      content: `
本日${data.agendaCandidates.length}件の議題候補が検出されました。

${data.agendaCandidates.slice(0, 5).map((match, i) => 
  `${i + 1}. ${match.summary.summary}\n   カテゴリ: ${match.topCategory} (スコア: ${match.topScore})`
).join('\n\n')}
      `.trim()
    });
  }
  
  // ドキュメント生成セクション
  if (data.manualStats || data.faqStats) {
    let docContent = '';
    if (data.manualStats) {
      docContent += `マニュアル: ${data.manualStats.newItems}件\n`;
    }
    if (data.faqStats) {
      docContent += `FAQ: ${data.faqStats.newItems}件\n`;
    }
    
    report.sections.push({
      title: '📚 ドキュメント生成',
      content: docContent.trim()
    });
  }
  
  return report;
}

function sendReportEmail(report) {
  const htmlBody = `
<!DOCTYPE html>
<html>
<head>
  <style>
    body { font-family: 'Helvetica Neue', Arial, sans-serif; line-height: 1.6; color: #333; }
    .container { max-width: 600px; margin: 0 auto; padding: 20px; }
    h1 { color: #2c3e50; border-bottom: 2px solid #3498db; padding-bottom: 10px; }
    h2 { color: #34495e; margin-top: 30px; }
    .section { background: #f8f9fa; padding: 15px; margin: 15px 0; border-radius: 5px; }
    .stats { display: grid; grid-template-columns: repeat(3, 1fr); gap: 10px; margin: 15px 0; }
    .stat-card { background: white; padding: 10px; border-radius: 3px; text-align: center; }
    .stat-number { font-size: 24px; font-weight: bold; color: #3498db; }
    .stat-label { font-size: 12px; color: #7f8c8d; }
    pre { white-space: pre-wrap; word-wrap: break-word; }
    .footer { margin-top: 40px; padding-top: 20px; border-top: 1px solid #e0e0e0; font-size: 12px; color: #7f8c8d; }
  </style>
</head>
<body>
  <div class="container">
    <h1>${report.title}</h1>
    
    ${report.sections.map(section => `
      <div class="section">
        <h2>${section.title}</h2>
        <pre>${section.content}</pre>
      </div>
    `).join('')}
    
    <div class="footer">
      <p>このレポートは自動生成されました。</p>
      <p><a href="https://docs.google.com/spreadsheets/d/${SPREADSHEET_ID}">スプレッドシートで詳細を確認</a></p>
    </div>
  </div>
</body>
</html>
`;
  
  // 重複送信チェック
  if (isDuplicateEmail('general_report', report.title, report)) {
    logInfo('同じ内容のレポートが既に送信されています。スキップします。');
    return;
  }
  
  MailApp.sendEmail({
    to: REPORT_EMAIL,
    subject: report.title,
    htmlBody: htmlBody
  });
  
  // 送信履歴を記録
  recordEmailSent('general_report', report.title, report);
}

function sendSlackReport(channel, report) {
  const blocks = [
    {
      type: 'header',
      text: {
        type: 'plain_text',
        text: report.title
      }
    }
  ];
  
  report.sections.forEach(section => {
    blocks.push({
      type: 'section',
      text: {
        type: 'mrkdwn',
        text: `*${section.title}*\n\`\`\`${section.content}\`\`\``
      }
    });
    blocks.push({ type: 'divider' });
  });
  
  blocks.push({
    type: 'actions',
    elements: [
      {
        type: 'button',
        text: {
          type: 'plain_text',
          text: 'スプレッドシートで詳細を確認'
        },
        url: `https://docs.google.com/spreadsheets/d/${SPREADSHEET_ID}`
      }
    ]
  });
  
  slackAPI('chat.postMessage', {
    channel: channel,
    blocks: blocks
  });
}

function saveReportToSheet(ss, report) {
  const reportSheet = getOrCreateSheet(ss, SHEETS.DAILY_REPORT, [
    '日付', 'タイトル', 'レポート内容', '生成時刻'
  ]);
  
  const reportContent = report.sections.map(s => 
    `${s.title}\n${s.content}`
  ).join('\n\n');
  
  const lastRow = reportSheet.getLastRow();
  reportSheet.getRange(lastRow + 1, 1, 1, 4).setValues([[
    new Date(),
    report.title,
    reportContent,
    new Date()
  ]]);
}

// ========= スプレッドシート初期化 =========
function initializeSpreadsheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  // 確認ダイアログ
  const response = ui.alert(
    '初期化確認',
    '既存のシートをすべて削除して、新しいシートを作成します。\n続行しますか？',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) {
    return;
  }
  
  // 既存のシートをクリア（最初のシートは残す）
  const sheets = ss.getSheets();
  for (let i = sheets.length - 1; i > 0; i--) {
    ss.deleteSheet(sheets[i]);
  }
  
  // すべての必要なシートを作成
  console.log('シート作成開始...');
  
  // 基本設定シート
  createConfigSheet(ss);
  
  // Slack関連シート
  createSyncStateSheet(ss);
  createMessagesSheet(ss);
  createSlackLogSheetInSpreadsheet(ss);  // slack_logシート
  
  // 分析・分類シート
  createCategoriesSheet(ss);
  createChecklistsSheet(ss);
  
  // ドキュメント関連シート
  createTemplatesSheet(ss);
  createDraftsSheet(ss);
  
  // ログシート
  createLogsSheet(ss);
  
  // 追加のシート（必要に応じて）
  createBusinessManualSheet(ss);  // business_manual
  createFAQListSheet(ss);         // faq_list
  createDailyReportSheet(ss);     // daily_report
  
  // 最初のデフォルトシートを削除
  try {
    ss.deleteSheet(sheets[0]);
  } catch (e) {
    // 削除できない場合は無視
  }
  
  ui.alert('初期化完了', 'スプレッドシートの初期化が完了しました。\n\n作成されたシート:\n- Config\n- SyncState\n- Messages\n- slack_log\n- Categories\n- Checklists\n- Templates\n- Drafts\n- Logs\n- business_manual\n- faq_list\n- daily_report', ui.ButtonSet.OK);
}

// ========= シート作成関数（initializeSpreadsheet用） =========

// Config シート
function createConfigSheet(ss) {
  // スプレッドシートオブジェクトの確認
  if (!ss) {
    ss = SpreadsheetApp.getActiveSpreadsheet();
  }
  
  // 既存のシートがあるか確認
  let sheet = ss.getSheetByName('Config');
  if (sheet) {
    console.log('Configシートは既に存在します');
    return sheet;
  }
  
  sheet = ss.insertSheet('Config');
  
  const headers = ['設定項目', '値', '説明'];
  const data = [
    ['company', '', '会社名'],
    ['targetChannels', '', 'Slack監視対象チャンネルID（カンマ区切り）'],
    ['notifySlackChannel', '', '通知先SlackチャンネルID'],
    ['notifyEmails', '', '通知先メールアドレス（カンマ区切り）'],
    ['openaiModel', 'gpt-5', 'OpenAIモデル名（要約・分類用）'],
    ['openaiModelDraft', 'gpt-5', 'OpenAIモデル名（ドラフト生成用）'],
    ['OPENAI_MODEL', 'gpt-5', 'メイン処理用OpenAIモデル名'],
    ['classificationThreshold', '0.6', '該当判定しきい値（0-1）'],
    ['syncIntervalMinutes', '5', 'Slack同期間隔（分）'],
    ['analysisIntervalHours', '1', 'AI分析実行間隔（時間）'],
    ['notificationHours', '9,15', '通知時刻（カンマ区切り）']
  ];
  
  sheet.getRange(1, 1, 1, 3).setValues([headers]);
  sheet.getRange(2, 1, data.length, 3).setValues(data);
  
  // ヘッダーのフォーマット
  sheet.getRange(1, 1, 1, 3)
    .setBackground('#4285F4')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold');
  
  // 列幅の調整
  sheet.setColumnWidth(1, 200);
  sheet.setColumnWidth(2, 300);
  sheet.setColumnWidth(3, 400);
  
  // フリーズ
  sheet.setFrozenRows(1);
}

// SyncState シート
function createSyncStateSheet(ss) {
  // スプレッドシートオブジェクトの確認
  if (!ss) {
    ss = SpreadsheetApp.getActiveSpreadsheet();
  }
  
  // 既存のシートがあるか確認
  let sheet = ss.getSheetByName('SyncState');
  if (sheet) {
    console.log('SyncStateシートは既に存在します');
    return sheet;
  }
  
  sheet = ss.insertSheet('SyncState');
  
  const headers = ['channel_id', 'last_sync_ts', 'last_sync_datetime', 'message_count', 'status'];
  
  sheet.getRange(1, 1, 1, 5).setValues([headers]);
  
  // ヘッダーのフォーマット
  sheet.getRange(1, 1, 1, 5)
    .setBackground('#34A853')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold');
  
  // 列幅の調整
  sheet.setColumnWidth(1, 150);
  sheet.setColumnWidth(2, 150);
  sheet.setColumnWidth(3, 180);
  sheet.setColumnWidth(4, 120);
  sheet.setColumnWidth(5, 100);
  
  sheet.setFrozenRows(1);
}

// Messages シート
function createMessagesSheet(ss) {
  // スプレッドシートオブジェクトの確認
  if (!ss) {
    ss = SpreadsheetApp.getActiveSpreadsheet();
  }
  
  // 既存のシートがあるか確認
  let sheet = ss.getSheetByName('Messages');
  if (sheet) {
    console.log('Messagesシートは既に存在します');
    return sheet;
  }
  
  sheet = ss.insertSheet('Messages');
  
  const headers = [
    'id',
    'channel_id',
    'message_ts',
    'thread_ts',
    'text_raw',
    'user_name',
    'summary_json',
    'classification_json',
    'match_flag',
    'human_judgement',
    'permalink',
    'checklist_proposed',
    'agenda_selected',
    'draft_doc_url',
    'created_at',
    'updated_at'
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // ヘッダーのフォーマット
  sheet.getRange(1, 1, 1, headers.length)
    .setBackground('#EA4335')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold');
  
  // 列幅の調整
  sheet.setColumnWidth(1, 200); // id
  sheet.setColumnWidth(2, 120); // channel_id
  sheet.setColumnWidth(3, 150); // message_ts
  sheet.setColumnWidth(4, 150); // thread_ts
  sheet.setColumnWidth(5, 400); // text_raw
  sheet.setColumnWidth(6, 150); // user_name
  sheet.setColumnWidth(7, 300); // summary_json
  sheet.setColumnWidth(8, 300); // classification_json
  sheet.setColumnWidth(9, 100); // match_flag
  sheet.setColumnWidth(10, 100); // human_judgement
  sheet.setColumnWidth(11, 300); // permalink
  
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(1);
}

// Categories シート
function createCategoriesSheet(ss) {
  // スプレッドシートオブジェクトの確認
  if (!ss) {
    ss = SpreadsheetApp.getActiveSpreadsheet();
  }
  
  // 既存のシートがあるか確認
  let sheet = ss.getSheetByName('Categories');
  if (sheet) {
    console.log('Categoriesシートは既に存在します');
    return sheet;
  }
  
  sheet = ss.insertSheet('Categories');
  
  const headers = ['カテゴリID', 'カテゴリ名', '説明', 'キーワード', 'アクティブ'];
  const data = [
    ['board_meeting', '取締役会', '取締役会で議論すべき事項', '承認,決議,報告,取締役', 'TRUE'],
    ['shareholder_meeting', '株主総会', '株主総会で決議が必要な事項', '株主,定款,配当,資本', 'TRUE'],
    ['investment', '投資・M&A', '投資案件やM&A関連', '投資,買収,出資,DD', 'TRUE'],
    ['compliance', 'コンプライアンス', '法務・コンプライアンス関連', '法律,規制,契約,リスク', 'TRUE'],
    ['finance', '財務・経理', '財務・経理に関する事項', '予算,決算,資金,財務', 'TRUE'],
    ['hr', '人事・労務', '人事・労務に関する事項', '採用,人事,給与,労務', 'TRUE'],
    ['strategy', '経営戦略', '経営戦略・事業計画', '戦略,計画,方針,目標', 'TRUE'],
    ['urgent', '緊急対応', '緊急で対応が必要な事項', '緊急,至急,ASAP,急ぎ', 'TRUE']
  ];
  
  sheet.getRange(1, 1, 1, 5).setValues([headers]);
  sheet.getRange(2, 1, data.length, 5).setValues(data);
  
  // ヘッダーのフォーマット
  sheet.getRange(1, 1, 1, 5)
    .setBackground('#FBBC04')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold');
  
  // 列幅の調整
  sheet.setColumnWidth(1, 150);
  sheet.setColumnWidth(2, 150);
  sheet.setColumnWidth(3, 300);
  sheet.setColumnWidth(4, 300);
  sheet.setColumnWidth(5, 80);
  
  sheet.setFrozenRows(1);
}

// Checklists シート
function createChecklistsSheet(ss) {
  // スプレッドシートオブジェクトの確認
  if (!ss) {
    ss = SpreadsheetApp.getActiveSpreadsheet();
  }
  
  // 既存のシートがあるか確認
  let sheet = ss.getSheetByName('Checklists');
  if (sheet) {
    console.log('Checklistsシートは既に存在します');
    return sheet;
  }
  
  sheet = ss.insertSheet('Checklists');
  
  const headers = ['ID', 'カテゴリ', 'チェック項目', '重要度', '必須'];
  const data = [
    ['CHK001', 'board_meeting', '議事録の作成と承認', '高', 'TRUE'],
    ['CHK002', 'board_meeting', '決議事項の明確化', '高', 'TRUE'],
    ['CHK003', 'board_meeting', '利益相反の確認', '高', 'TRUE'],
    ['CHK004', 'shareholder_meeting', '招集通知の送付', '高', 'TRUE'],
    ['CHK005', 'shareholder_meeting', '委任状の回収', '中', 'FALSE'],
    ['CHK006', 'investment', 'デューデリジェンスの実施', '高', 'TRUE'],
    ['CHK007', 'investment', '投資委員会での承認', '高', 'TRUE'],
    ['CHK008', 'compliance', '法的リスクの評価', '高', 'TRUE'],
    ['CHK009', 'compliance', '関連法規の確認', '高', 'TRUE'],
    ['CHK010', 'finance', '予算との整合性確認', '中', 'TRUE']
  ];
  
  sheet.getRange(1, 1, 1, 5).setValues([headers]);
  sheet.getRange(2, 1, data.length, 5).setValues(data);
  
  // ヘッダーのフォーマット
  sheet.getRange(1, 1, 1, 5)
    .setBackground('#9C27B0')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold');
  
  // 列幅の調整
  sheet.setColumnWidth(1, 80);
  sheet.setColumnWidth(2, 150);
  sheet.setColumnWidth(3, 300);
  sheet.setColumnWidth(4, 80);
  sheet.setColumnWidth(5, 80);
  
  sheet.setFrozenRows(1);
}

// Templates シート
function createTemplatesSheet(ss) {
  // スプレッドシートオブジェクトの確認
  if (!ss) {
    ss = SpreadsheetApp.getActiveSpreadsheet();
  }
  
  // 既存のシートがあるか確認
  let sheet = ss.getSheetByName('Templates');
  if (sheet) {
    console.log('Templatesシートは既に存在します');
    return sheet;
  }
  
  sheet = ss.insertSheet('Templates');
  
  const headers = ['テンプレートID', 'テンプレート名', 'カテゴリ', 'ドキュメントURL', '説明', 'プレースホルダー', '最終更新'];
  
  sheet.getRange(1, 1, 1, 7).setValues([headers]);
  
  // ヘッダーのフォーマット
  sheet.getRange(1, 1, 1, 7)
    .setBackground('#00ACC1')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold');
  
  // 列幅の調整
  sheet.setColumnWidth(1, 120);
  sheet.setColumnWidth(2, 200);
  sheet.setColumnWidth(3, 120);
  sheet.setColumnWidth(4, 300);
  sheet.setColumnWidth(5, 300);
  sheet.setColumnWidth(6, 200);
  sheet.setColumnWidth(7, 150);
  
  sheet.setFrozenRows(1);
}

// Drafts シート
function createDraftsSheet(ss) {
  // スプレッドシートオブジェクトの確認
  if (!ss) {
    ss = SpreadsheetApp.getActiveSpreadsheet();
  }
  
  // 既存のシートがあるか確認
  let sheet = ss.getSheetByName('Drafts');
  if (sheet) {
    console.log('Draftsシートは既に存在します');
    return sheet;
  }
  
  sheet = ss.insertSheet('Drafts');
  
  const headers = [
    'draft_id',
    'message_id',
    'category',
    'draft_type',
    'title',
    'content',
    'doc_url',
    'status',
    'created_at',
    'updated_at'
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // ヘッダーのフォーマット
  sheet.getRange(1, 1, 1, headers.length)
    .setBackground('#795548')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold');
  
  // 列幅の調整
  sheet.setColumnWidth(1, 150); // draft_id
  sheet.setColumnWidth(2, 200); // message_id
  sheet.setColumnWidth(3, 120); // category
  sheet.setColumnWidth(4, 120); // draft_type
  sheet.setColumnWidth(5, 300); // title
  sheet.setColumnWidth(6, 500); // content
  sheet.setColumnWidth(7, 300); // doc_url
  sheet.setColumnWidth(8, 100); // status
  
  sheet.setFrozenRows(1);
}

// Logs シート
function createLogsSheet(ss) {
  // スプレッドシートオブジェクトの確認
  if (!ss) {
    ss = SpreadsheetApp.getActiveSpreadsheet();
  }
  
  // 既存のシートがあるか確認
  let sheet = ss.getSheetByName('Logs');
  if (sheet) {
    console.log('Logsシートは既に存在します');
    return sheet;
  }
  
  sheet = ss.insertSheet('Logs');
  
  const headers = ['timestamp', 'level', 'message', 'details'];
  
  sheet.getRange(1, 1, 1, 4).setValues([headers]);
  
  // ヘッダーのフォーマット
  sheet.getRange(1, 1, 1, 4)
    .setBackground('#607D8B')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold');
  
  // 列幅の調整
  sheet.setColumnWidth(1, 180);
  sheet.setColumnWidth(2, 80);
  sheet.setColumnWidth(3, 400);
  sheet.setColumnWidth(4, 400);
  
  sheet.setFrozenRows(1);
}

// ========= シート個別作成関数 =========

// Configシートを個別に作成
function createConfigSheetOnly() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const existingSheet = ss.getSheetByName('Config');
  
  if (existingSheet) {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      'シート作成確認',
      'Configシートは既に存在します。削除して再作成しますか？',
      ui.ButtonSet.YES_NO
    );
    
    if (response === ui.Button.YES) {
      ss.deleteSheet(existingSheet);
    } else {
      return;
    }
  }
  
  createConfigSheet(ss);
  SpreadsheetApp.getUi().alert('Configシートを作成しました');
}

// SyncStateシートを個別に作成
function createSyncStateSheetOnly() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const existingSheet = ss.getSheetByName('SyncState');
  
  if (existingSheet) {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      'シート作成確認',
      'SyncStateシートは既に存在します。削除して再作成しますか？',
      ui.ButtonSet.YES_NO
    );
    
    if (response === ui.Button.YES) {
      ss.deleteSheet(existingSheet);
    } else {
      return;
    }
  }
  
  createSyncStateSheet(ss);
  SpreadsheetApp.getUi().alert('SyncStateシートを作成しました');
}

// Messagesシートを個別に作成
function createMessagesSheetOnly() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const existingSheet = ss.getSheetByName('Messages');
  
  if (existingSheet) {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      'シート作成確認',
      'Messagesシートは既に存在します。削除して再作成しますか？',
      ui.ButtonSet.YES_NO
    );
    
    if (response === ui.Button.YES) {
      ss.deleteSheet(existingSheet);
    } else {
      return;
    }
  }
  
  createMessagesSheet(ss);
  SpreadsheetApp.getUi().alert('Messagesシートを作成しました');
}

// Categoriesシートを個別に作成
function createCategoriesSheetOnly() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const existingSheet = ss.getSheetByName('Categories');
  
  if (existingSheet) {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      'シート作成確認',
      'Categoriesシートは既に存在します。削除して再作成しますか？',
      ui.ButtonSet.YES_NO
    );
    
    if (response === ui.Button.YES) {
      ss.deleteSheet(existingSheet);
    } else {
      return;
    }
  }
  
  createCategoriesSheet(ss);
  SpreadsheetApp.getUi().alert('Categoriesシートを作成しました');
}

// Checklistsシートを個別に作成
function createChecklistsSheetOnly() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const existingSheet = ss.getSheetByName('Checklists');
  
  if (existingSheet) {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      'シート作成確認',
      'Checklistsシートは既に存在します。削除して再作成しますか？',
      ui.ButtonSet.YES_NO
    );
    
    if (response === ui.Button.YES) {
      ss.deleteSheet(existingSheet);
    } else {
      return;
    }
  }
  
  createChecklistsSheet(ss);
  SpreadsheetApp.getUi().alert('Checklistsシートを作成しました');
}

// Templatesシートを個別に作成
function createTemplatesSheetOnly() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const existingSheet = ss.getSheetByName('Templates');
  
  if (existingSheet) {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      'シート作成確認',
      'Templatesシートは既に存在します。削除して再作成しますか？',
      ui.ButtonSet.YES_NO
    );
    
    if (response === ui.Button.YES) {
      ss.deleteSheet(existingSheet);
    } else {
      return;
    }
  }
  
  createTemplatesSheet(ss);
  SpreadsheetApp.getUi().alert('Templatesシートを作成しました');
}

// Draftsシートを個別に作成
function createDraftsSheetOnly() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const existingSheet = ss.getSheetByName('Drafts');
  
  if (existingSheet) {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      'シート作成確認',
      'Draftsシートは既に存在します。削除して再作成しますか？',
      ui.ButtonSet.YES_NO
    );
    
    if (response === ui.Button.YES) {
      ss.deleteSheet(existingSheet);
    } else {
      return;
    }
  }
  
  createDraftsSheet(ss);
  SpreadsheetApp.getUi().alert('Draftsシートを作成しました');
}

// Logsシートを個別に作成
function createLogsSheetOnly() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const existingSheet = ss.getSheetByName('Logs');
  
  if (existingSheet) {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      'シート作成確認',
      'Logsシートは既に存在します。削除して再作成しますか？',
      ui.ButtonSet.YES_NO
    );
    
    if (response === ui.Button.YES) {
      ss.deleteSheet(existingSheet);
    } else {
      return;
    }
  }
  
  createLogsSheet(ss);
  SpreadsheetApp.getUi().alert('Logsシートを作成しました');
}

// 全シートの存在確認
function checkAllSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const requiredSheets = ['Config', 'SyncState', 'Messages', 'Categories', 
                         'Checklists', 'Templates', 'Drafts', 'Logs'];
  
  const missingSheets = [];
  const existingSheets = [];
  
  requiredSheets.forEach(sheetName => {
    if (ss.getSheetByName(sheetName)) {
      existingSheets.push(sheetName);
    } else {
      missingSheets.push(sheetName);
    }
  });
  
  let message = '📊 シート存在確認\n\n';
  
  if (existingSheets.length > 0) {
    message += '✅ 存在するシート:\n' + existingSheets.join(', ') + '\n\n';
  }
  
  if (missingSheets.length > 0) {
    message += '❌ 存在しないシート:\n' + missingSheets.join(', ');
  } else {
    message += '全ての必須シートが存在します。';
  }
  
  SpreadsheetApp.getUi().alert(message);
}

// ========= UIエラー対策 =========
/**
 * UIが利用可能かチェック
 * @returns {boolean} UIが利用可能な場合true
 */
function isUiAvailable() {
  try {
    SpreadsheetApp.getUi();
    return true;
  } catch (e) {
    return false;
  }
}

/**
 * 安全にUIアラートを表示（UIが使えない場合はログ出力）
 * @param {string} title - アラートのタイトル
 * @param {string} message - アラートのメッセージ
 * @param {ButtonSet} buttons - ボタンセット（省略可）
 */
function showAlertSafely(title, message, buttons) {
  if (isUiAvailable()) {
    const ui = SpreadsheetApp.getUi();
    if (buttons) {
      ui.alert(title, message, buttons);
    } else {
      ui.alert(title, message, ui.ButtonSet.OK);
    }
  } else {
    // UIが使えない場合はログに出力
    console.log(`[ALERT] ${title}: ${message}`);
  }
}

// ========= ユーティリティ関数 =========
function getOrCreateSheet(spreadsheet, sheetName, headers) {
  let sheet = spreadsheet.getSheetByName(sheetName);
  
  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
    if (headers && headers.length > 0) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length)
        .setBackground('#4285F4')
        .setFontColor('#FFFFFF')
        .setFontWeight('bold');
      sheet.setFrozenRows(1);
    }
  }
  
  return sheet;
}

function formatMessagesForAI(messages) {
  return messages.map(msg => {
    const time = msg.date instanceof Date ? 
      Utilities.formatDate(msg.date, 'JST', 'HH:mm') : 
      new Date(Number(msg.ts.split('.')[0]) * 1000).toLocaleTimeString('ja-JP');
    return `[${time}] ${msg.user_name || 'Unknown'}: ${msg.message || msg.text || ''}`;
  }).join('\n');
}

function saveFAQToSheet(ss, faq, originalMessages) {
  const sheet = getOrCreateSheet(ss, SHEETS.FAQ_LIST, [
    '作成日時', '質問', '回答', 'カテゴリ', 'タグ', 
    '元のチャンネル', '関連メッセージ', 'ステータス'
  ]);
  
  const timestamp = new Date();
  const channelName = originalMessages[0]?.channel_name || '';
  const messageIds = originalMessages.map(m => `${m.channel_id}_${m.ts}`).join(', ');
  
  const fullAnswer = faq.answer + (faq.supplement ? '\n\n補足: ' + faq.supplement : '');
  
  sheet.appendRow([
    timestamp,
    faq.question,
    fullAnswer,
    faq.category || 'その他',
    faq.tags || '',
    channelName,
    messageIds,
    'アクティブ'
  ]);
}

function saveManualToSheet(ss, manual, originalMessages) {
  const sheet = getOrCreateSheet(ss, SHEETS.BUSINESS_MANUAL, [
    '作成日時', 'カテゴリ', 'タイトル', '内容', 
    '元のチャンネル', '関連メッセージ', 'ステータス',
    '参加者', 'キーワード', '重要度'
  ]);
  
  const timestamp = new Date();
  const channelName = originalMessages[0]?.channel_name || '';
  const messageIds = originalMessages.map(m => `${m.channel_id}_${m.ts}`).join(', ');
  const participants = [...new Set(originalMessages.map(m => m.user_name).filter(Boolean))].join(', ');
  
  sheet.appendRow([
    timestamp,
    manual.category || 'その他',
    manual.title,
    manual.content,
    channelName,
    messageIds,
    'アクティブ',
    participants,
    manual.keywords || '',
    manual.importance || 'medium'
  ]);
}

function saveAnalysisResults(ss, results) {
  const sheet = getOrCreateSheet(ss, 'analysis_results', [
    '分析日時', 'チャンネル', '総メッセージ数', 'トピック数', 
    '重要トピック数', '最もアクティブなユーザー', 'ピーク時間帯'
  ]);
  
  const timestamp = new Date();
  
  results.forEach(result => {
    sheet.appendRow([
      timestamp,
      result.channel,
      result.stats.totalMessages,
      result.stats.topicCount,
      result.importantTopics,
      result.stats.mostActiveUser,
      result.stats.peakHours.join(', ')
    ]);
  });
}

// 既存の関数（互換性維持）
function getConfigData(sheet) {
  if (!sheet) return {};
  
  const data = sheet.getDataRange().getValues();
  const config = {};
  
  // ヘッダー行（1行目）をスキップして、2行目から処理
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[0]) {
      config[row[0]] = row[1];
    }
  }
  
  // 配列形式のデータを処理
  if (config.targetChannels) {
    config.targetChannels = config.targetChannels.split(',').map(s => s.trim()).filter(s => s);
  }
  if (config.notifyEmails) {
    config.notifyEmails = config.notifyEmails.split(',').map(s => s.trim()).filter(s => s);
  }
  
  // デバッグログ
  logInfo('Config読み込み', `targetChannels: ${config.targetChannels ? config.targetChannels.join(', ') : 'なし'}`);
  
  return config;
}

function getCategoriesData(sheet) {
  if (!sheet) return [];
  
  const data = sheet.getDataRange().getValues();
  const categories = [];
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) {
      categories.push({
        name: data[i][0],
        description: data[i][1],
        criteria: data[i][2],
        keywords: data[i][3] ? data[i][3].split(',').map(s => s.trim()) : []
      });
    }
  }
  
  return categories;
}

function getLastSyncTime(sheet, channelId) {
  if (!sheet) return '0';
  
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === channelId) {
      const timestamp = data[i][1];
      // タイムスタンプが有効な値か確認
      if (timestamp && !isNaN(timestamp)) {
        // 文字列として返す（Slack APIは文字列を期待）
        return String(timestamp);
      }
      return '0';
    }
  }
  
  return '0';
}

function updateLastSyncTime(sheet, channelId, timestamp) {
  if (!sheet) {
    console.error('SyncStateシートが存在しません');
    return;
  }
  
  console.log(`updateLastSyncTime: channelId=${channelId}, timestamp=${timestamp}, date=${new Date(parseFloat(timestamp) * 1000).toLocaleString('ja-JP')}`);
  
  const data = sheet.getDataRange().getValues();
  let found = false;
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === channelId) {
      // タイムスタンプと日時の両方を更新
      sheet.getRange(i + 1, 2).setValue(timestamp);
      sheet.getRange(i + 1, 3).setValue(new Date());
      sheet.getRange(i + 1, 4).setValue(1); // message_count（初期値）
      sheet.getRange(i + 1, 5).setValue('synced'); // status
      found = true;
      console.log(`既存のチャンネル ${channelId} の同期時刻を更新: ${timestamp}`);
      break;
    }
  }
  
  if (!found) {
    const lastRow = sheet.getLastRow();
    // 新規行を追加: channel_id, last_sync_ts, last_sync_datetime, message_count, status
    sheet.getRange(lastRow + 1, 1, 1, 5).setValues([[
      channelId, 
      timestamp, 
      new Date(),
      1,  // message_count初期値
      'synced'  // status
    ]]);
    console.log(`新規チャンネル ${channelId} の同期情報を追加: ${timestamp}`);
  }
}

function saveMessage(sheet, channelId, message) {
  if (!sheet) {
    console.error('シートが存在しません');
    return;
  }
  
  const messageId = `${channelId}:${message.ts}`;
  
  console.log(`メッセージ保存中: ${messageId}`);
  
  // 既存のメッセージをチェック（パフォーマンス改善版）
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    const existingIds = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
    if (existingIds.includes(messageId)) {
      console.log(`既存メッセージをスキップ: ${messageId}`);
      return;
    }
  }
  
  // パーマリンクを生成（API呼び出しを避ける）
  const permalink = `https://slack.com/archives/${channelId}/p${message.ts.replace('.', '')}`;
  
  // ユーザー名を取得
  const userName = message.user_name || getUserInfo(message.user).name || message.user || 'unknown';
  
  // 現在時刻を取得
  const now = new Date();
  
  console.log(`新規メッセージを行${lastRow + 1}に追加`);
  
  try {
    // シートのヘッダーに合わせてカラム数を調整（16カラム）
    sheet.getRange(lastRow + 1, 1, 1, 16).setValues([[
      messageId,                      // id
      channelId,                       // channel_id
      message.ts,                      // message_ts
      message.thread_ts || '',         // thread_ts
      message.text || '',              // text_raw
      userName,                        // user_name
      '',                              // summary_json
      '',                              // classification_json
      '',                              // match_flag
      '',                              // human_judgement
      permalink,                       // permalink
      '',                              // checklist_proposed
      '',                              // agenda_selected
      '',                              // draft_doc_url
      now,                             // created_at
      now                              // updated_at
    ]]);
    console.log(`メッセージ保存完了: ${messageId}`);
  } catch (error) {
    console.error(`メッセージ保存エラー: ${error.toString()}`);
    console.error(`詳細: チャンネルID=${channelId}, タイムスタンプ=${message.ts}`);
    throw error;
  }
}

// ユーザー情報キャッシュ
let userInfoCache = {};
let userCacheExpiry = 0;

// すべてのユーザー情報を一括で取得してキャッシュ
function loadAllUsers() {
  try {
    console.log('全ユーザー情報を取得中...');
    const response = slackAPI('users.list', {
      limit: 1000
    });
    
    if (response.members) {
      response.members.forEach(user => {
        userInfoCache[user.id] = {
          id: user.id,
          name: user.real_name || user.name || user.id,
          real_name: user.real_name || user.name || user.id,
          email: user.profile?.email || '',
          display_name: user.profile?.display_name || user.name || user.id,
          avatar: user.profile?.image_48 || '',
          is_bot: user.is_bot || false
        };
      });
      
      // キャッシュを30分間保持
      userCacheExpiry = Date.now() + 30 * 60 * 1000;
      console.log(`${response.members.length}人のユーザー情報をキャッシュしました`);
      return true;
    }
  } catch (error) {
    console.error('ユーザーリスト取得エラー:', error);
    return false;
  }
}

function getUserInfo(userId) {
  if (!userId) return { name: 'Unknown', email: '', real_name: 'Unknown' };
  
  // キャッシュが期限切れの場合は再読み込み
  if (Date.now() > userCacheExpiry) {
    loadAllUsers();
  }
  
  // キャッシュをチェック
  if (userInfoCache[userId]) {
    return userInfoCache[userId];
  }
  
  // キャッシュにない場合は個別に取得を試みる
  try {
    // users.info APIを使用してユーザー情報を取得
    const response = slackAPI('users.info', {
      user: userId
    });
    
    if (response.user) {
      const userInfo = {
        id: userId,
        name: response.user.real_name || response.user.name || userId,
        real_name: response.user.real_name || response.user.name || userId,
        email: response.user.profile?.email || '',
        display_name: response.user.profile?.display_name || response.user.name || userId,
        avatar: response.user.profile?.image_48 || '',
        is_bot: response.user.is_bot || false
      };
      userInfoCache[userId] = userInfo;
      return userInfo;
    }
  } catch (error) {
    console.warn(`ユーザー情報取得失敗 (${userId}): ${error.toString()}`);
  }
  
  // エラー時のフォールバック
  const fallback = { 
    id: userId, 
    name: userId, 
    real_name: userId,
    email: '',
    display_name: userId,
    avatar: '',
    is_bot: false
  };
  userInfoCache[userId] = fallback;
  return fallback;
}

// ========= SlackユーザーIDを実名に変換 =========
function convertSlackUserIdsToNames(text) {
  if (!text || typeof text !== 'string') return text;
  
  // <@USERID>形式のメンションを検出して実名に変換
  return text.replace(/<@([A-Z0-9]+)>/g, (match, userId) => {
    const userInfo = getUserInfo(userId);
    return userInfo.real_name || userInfo.name || match;
  });
}

// ========= 分析結果内のユーザーIDを実名に変換 =========
function convertAnalysisUserIds(analysisResult) {
  if (!analysisResult) return analysisResult;
  
  // actionItemsのownerを変換
  if (analysisResult.actionItems && Array.isArray(analysisResult.actionItems)) {
    analysisResult.actionItems = analysisResult.actionItems.map(item => {
      if (typeof item === 'object' && item.owner) {
        item.owner = convertSlackUserIdsToNames(item.owner);
      }
      return item;
    });
  }
  
  // stakeholdersを変換
  if (analysisResult.stakeholders && Array.isArray(analysisResult.stakeholders)) {
    analysisResult.stakeholders = analysisResult.stakeholders.map(stakeholder => {
      return convertSlackUserIdsToNames(stakeholder);
    });
  }
  
  // topicsの説明文内のユーザーIDも変換
  if (analysisResult.topics && Array.isArray(analysisResult.topics)) {
    analysisResult.topics = analysisResult.topics.map(topic => {
      if (typeof topic === 'object' && topic.description) {
        topic.description = convertSlackUserIdsToNames(topic.description);
      }
      return topic;
    });
  }
  
  // summaryのユーザーIDも変換
  if (analysisResult.summary) {
    analysisResult.summary = convertSlackUserIdsToNames(analysisResult.summary);
  }
  
  return analysisResult;
}

function getMessagePermalink(channelId, messageTs) {
  try {
    const response = slackAPI('chat.getPermalink', {
      channel: channelId,
      message_ts: messageTs
    });
    return response.permalink;
  } catch (error) {
    return '';
  }
}

function getUnanalyzedMessages(sheet) {
  if (!sheet) return [];
  
  const data = sheet.getDataRange().getValues();
  const unanalyzed = [];
  
  for (let i = 1; i < data.length; i++) {
    if (!data[i][6]) { // summary_json列が空
      unanalyzed.push(data[i]);
    }
  }
  
  return unanalyzed;
}

function getMatchedMessages(sheet) {
  if (!sheet) return [];
  
  const data = sheet.getDataRange().getValues();
  const matches = [];
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][8] === true && !data[i][9]) { // match_flag=true かつ human_judgement未設定
      try {
        const summary = JSON.parse(data[i][6] || '{}');
        const classification = JSON.parse(data[i][7] || '[]');
        
        const topCategory = classification.reduce((prev, current) => 
          (prev.score > current.score) ? prev : current, { score: 0 }
        );
        
        matches.push({
          id: data[i][0],
          summary: summary,
          classification: classification,
          topCategory: topCategory.category,
          topScore: topCategory.score,
          permalink: data[i][10]
        });
      } catch (e) {
        // JSONパースエラーは無視
      }
    }
  }
  
  return matches;
}

function updateAnalysisResult(sheet, messageId, result) {
  if (!sheet) return;
  
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === messageId) {
      sheet.getRange(i + 1, 7).setValue(JSON.stringify(result.summary));
      sheet.getRange(i + 1, 8).setValue(JSON.stringify(result.classification));
      sheet.getRange(i + 1, 9).setValue(result.matchFlag);
      break;
    }
  }
}

// ========= 通知処理（既存機能） =========
function sendDailyNotification() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const messagesSheet = ss.getSheetByName(SHEETS.MESSAGES);
  const configSheet = ss.getSheetByName(SHEETS.CONFIG);
  
  const config = getConfigData(configSheet);
  const matches = getMatchedMessages(messagesSheet);
  
  if (matches.length === 0) {
    logInfo('通知対象がありません');
    return;
  }
  
  // Slack通知
  if (config.notifySlackChannel) {
    sendSlackNotification(config.notifySlackChannel, matches);
  }
  
  // メール通知
  if (config.notifyEmails && config.notifyEmails.length > 0) {
    sendEmailNotification(config.notifyEmails, matches);
  }
  
  logInfo(`${matches.length}件の該当案件を通知しました`);
}

function sendSlackNotification(channel, matches) {
  const blocks = [
    {
      type: 'header',
      text: {
        type: 'plain_text',
        text: '📋 議題候補の通知'
      }
    },
    {
      type: 'section',
      text: {
        type: 'mrkdwn',
        text: `*${matches.length}件の該当案件があります*`
      }
    }
  ];
  
  matches.forEach((match, index) => {
    blocks.push({
      type: 'divider'
    });
    blocks.push({
      type: 'section',
      text: {
        type: 'mrkdwn',
        text: `*${index + 1}. ${match.summary.summary}*\n` +
               `カテゴリ: ${match.topCategory}\n` +
               `スコア: ${match.topScore}\n` +
               `<${match.permalink}|元のメッセージを見る>`
      }
    });
  });
  
  blocks.push({
    type: 'actions',
    elements: [
      {
        type: 'button',
        text: {
          type: 'plain_text',
          text: 'スプレッドシートで確認'
        },
        url: `https://docs.google.com/spreadsheets/d/${SPREADSHEET_ID}`
      }
    ]
  });
  
  slackAPI('chat.postMessage', {
    channel: channel,
    blocks: blocks
  });
}

function sendEmailNotification(emails, matches) {
  const subject = `【議題候補】${matches.length}件の該当案件`;
  
  // 重複送信チェック
  if (isDuplicateEmail('daily_notification', subject, matches)) {
    logInfo('同じ内容の議題候補通知が既に送信されています。スキップします。');
    return;
  }
  
  const htmlBody = `
<!DOCTYPE html>
<html>
<head>
  <style>
    body { font-family: Arial, sans-serif; }
    h2 { color: #333; }
    .match { margin: 20px 0; padding: 15px; border: 1px solid #ddd; border-radius: 5px; }
    .category { color: #0066cc; font-weight: bold; }
    .score { color: #666; }
    .link { margin-top: 10px; }
  </style>
</head>
<body>
  <h2>📋 議題候補の通知</h2>
  <p>${matches.length}件の該当案件があります。</p>
  
  ${matches.map((match, index) => `
    <div class="match">
      <h3>${index + 1}. ${match.summary.summary}</h3>
      <p class="category">カテゴリ: ${match.topCategory}</p>
      <p class="score">スコア: ${match.topScore}</p>
      <div class="link">
        <a href="${match.permalink}">Slackで元のメッセージを見る</a>
      </div>
    </div>
  `).join('')}
  
  <hr>
  <p><a href="https://docs.google.com/spreadsheets/d/${SPREADSHEET_ID}">スプレッドシートで詳細を確認</a></p>
</body>
</html>
`;
  
  MailApp.sendEmail({
    to: emails.join(','),
    subject: subject,
    htmlBody: htmlBody
  });
  
  // 送信履歴を記録
  recordEmailSent('daily_notification', subject, matches);
}

// ========= ドキュメント生成（既存機能） =========
function generateDraftForSelected() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const selection = sheet.getActiveRange();
  
  if (sheet.getName() !== SHEETS.MESSAGES) {
    SpreadsheetApp.getUi().alert('Messagesシートで実行してください');
    return;
  }
  
  const rows = [];
  for (let i = selection.getRow(); i <= selection.getLastRow(); i++) {
    rows.push(sheet.getRange(i, 1, 1, sheet.getLastColumn()).getValues()[0]);
  }
  
  generateDocuments(rows);
}

function generateDocuments(rows) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const templatesSheet = ss.getSheetByName(SHEETS.TEMPLATES);
  const draftsSheet = ss.getSheetByName(SHEETS.DRAFTS);
  
  rows.forEach(row => {
    const messageId = row[0];
    const summary = JSON.parse(row[6] || '{}');
    const classification = JSON.parse(row[7] || '[]');
    const humanJudgement = row[9];
    
    if (humanJudgement !== '必要') {
      return;
    }
    
    // 最も高いスコアのカテゴリを特定
    const topCategory = classification.reduce((prev, current) => 
      (prev.score > current.score) ? prev : current
    );
    
    // ドラフト生成
    const draft = generateDraft(topCategory.category, summary);
    
    // Googleドキュメントに出力
    const docUrl = createDocument(topCategory.category, draft);
    
    // 結果を保存
    saveDraftRecord(draftsSheet, messageId, topCategory.category, docUrl);
  });
  
  SpreadsheetApp.getUi().alert('ドラフト生成完了');
}

function generateDraft(category, summary) {
  const prompt = `
以下の要約を基に、${category}の議事録案を作成してください。

要約:
${JSON.stringify(summary, null, 2)}

以下の形式で出力してください：
- 議題名
- 開催日時
- 参加者
- 議事内容
- 決議事項
- 今後の対応
`;

  const response = callOpenAIForAgenda([
    { role: 'system', content: '法務文書作成の専門家として議事録案を作成してください。' },
    { role: 'user', content: prompt }
  ], 'gpt-5');
  
  return response;
}

function createDocument(category, content) {
  const doc = DocumentApp.create(`${category}_議事録案_${new Date().toISOString()}`);
  const body = doc.getBody();
  
  body.setText(content);
  
  // スタイル設定
  const title = body.getParagraphs()[0];
  title.setHeading(DocumentApp.ParagraphHeading.HEADING1);
  
  doc.saveAndClose();
  
  return doc.getUrl();
}

function saveDraftRecord(sheet, messageId, category, docUrl) {
  if (!sheet) return;
  
  const lastRow = sheet.getLastRow();
  sheet.getRange(lastRow + 1, 1, 1, 5).setValues([[
    messageId,
    category,
    docUrl,
    new Date(),
    Session.getActiveUser().getEmail()
  ]]);
}

// ========= ログ管理 =========
function logInfo(message) {
  log('INFO', message);
}

function logError(context, error) {
  log('ERROR', `${context}: ${error}`);
}

function log(level, message) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const logsSheet = ss.getSheetByName(SHEETS.LOGS);
  
  if (!logsSheet) return;
  
  const lastRow = logsSheet.getLastRow();
  logsSheet.getRange(lastRow + 1, 1, 1, 3).setValues([[
    new Date(),
    level,
    message
  ]]);
}

// ========= Slack議題抽出関数 =========
// 最近のSlackメッセージを取得
function getRecentSlackMessages(slackLogSheet, startDate, endDate) {
  const data = slackLogSheet.getDataRange().getValues();
  const headers = data[0];
  const dateIndex = headers.indexOf('date');
  const messageIndex = headers.indexOf('message');
  const channelIndex = headers.indexOf('channel_name');
  const userIndex = headers.indexOf('user_name');
  
  if (dateIndex === -1 || messageIndex === -1) {
    return [];
  }
  
  const messages = [];
  for (let i = 1; i < data.length; i++) {
    const messageDate = new Date(data[i][dateIndex]);
    if (messageDate >= startDate && messageDate <= endDate) {
      messages.push({
        date: messageDate,
        message: data[i][messageIndex],
        channel: data[i][channelIndex] || 'unknown',
        user: data[i][userIndex] || 'unknown',
        raw: data[i]
      });
    }
  }
  
  return messages;
}

// Slackメッセージから議題・論点を抽出
function extractAgendasFromSlackMessages(messages) {
  const agendaItems = [];
  
  // メッセージをチャンネルごとにグループ化
  const channelGroups = {};
  messages.forEach(msg => {
    if (!channelGroups[msg.channel]) {
      channelGroups[msg.channel] = [];
    }
    channelGroups[msg.channel].push(msg);
  });
  
  // 各チャンネルのメッセージをAIで分析
  Object.keys(channelGroups).forEach(channel => {
    const channelMessages = channelGroups[channel];
    const concatenatedMessages = channelMessages
      .map(m => `[${m.user}] ${m.message}`)
      .join('\n');
    
    // OpenAI APIで議題抽出
    try {
      const analysis = analyzeMessagesWithAI(channel, concatenatedMessages);
      if (analysis && analysis.agendas) {
        analysis.agendas.forEach(agenda => {
          agendaItems.push({
            channel: channel,
            title: agenda.title,
            description: agenda.description,
            priority: agenda.priority || 'medium',
            participants: agenda.participants || [],
            keywords: agenda.keywords || [],
            sourceMessages: channelMessages.slice(0, 3) // 最初の3メッセージを保持
          });
        });
      }
    } catch (e) {
      console.error(`チャンネル ${channel} の分析エラー:`, e);
    }
  });
  
  return agendaItems;
}

// AIを使用してメッセージを分析
function analyzeMessagesWithAI(channel, messages) {
  if (!OPENAI_API_KEY) {
    console.error('OpenAI APIキーが設定されていません');
    return null;
  }
  
  const prompt = `
以下はSlackチャンネル「${channel}」での会話です。
この会話から重要な議題・論点を抽出してください。

会話内容:
${messages}

以下のJSON形式で議題を抽出してください:
{
  "agendas": [
    {
      "title": "議題のタイトル",
      "description": "議題の詳細説明",
      "priority": "high/medium/low",
      "participants": ["関係者のリスト"],
      "keywords": ["関連キーワード"]
    }
  ]
}

重要な議題がない場合は空の配列を返してください。
`;
  
  try {
    // APIペイロードを構築
    const apiPayload = {
        model: 'gpt-5',
        messages: [
          {
            role: 'system',
            content: '日本語のSlackメッセージから重要な議題を抽出する専門家として動作してください。'
          },
          {
            role: 'user',
            content: prompt
          }
      ]
    };
    
    // o3モデルの特別な処理（この関数では使用しないが、将来の拡張のため）
    const modelName = 'gpt-5';  // この関数では固定
    if (false) {
      // o3はtemperature=1のみサポート（デフォルト値なので設定不要）
      apiPayload.max_completion_tokens = 1000;
    } else {
      apiPayload.temperature = 0.6;  // より柔軟な判定のために温度を上げる
      apiPayload.max_tokens = 1000;
    }
    
    const response = UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions', {
      method: 'post',
      headers: {
        'Authorization': `Bearer ${OPENAI_API_KEY}`,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(apiPayload)
    });
    
    const result = JSON.parse(response.getContentText());
    const content = result.choices[0].message.content;
    return JSON.parse(content);
  } catch (e) {
    console.error('OpenAI API呼び出しエラー:', e);
    return null;
  }
}

// 業務フローチャートを生成
function generateBusinessFlowchart(agendaItems) {
  if (agendaItems.length === 0) {
    return null;
  }
  
  // 議題を優先度順にソート
  const sortedItems = agendaItems.sort((a, b) => {
    const priorityOrder = { high: 0, medium: 1, low: 2 };
    return priorityOrder[a.priority] - priorityOrder[b.priority];
  });
  
  // フローチャートのデータ構造を作成
  const flowchart = {
    title: '議題処理フロー',
    nodes: [],
    edges: []
  };
  
  // 開始ノード
  flowchart.nodes.push({
    id: 'start',
    label: '開始',
    type: 'start'
  });
  
  // 各議題をノードとして追加
  sortedItems.forEach((item, index) => {
    const nodeId = `agenda_${index}`;
    flowchart.nodes.push({
      id: nodeId,
      label: item.title,
      type: 'process',
      priority: item.priority,
      channel: item.channel
    });
    
    // エッジを追加
    if (index === 0) {
      flowchart.edges.push({
        from: 'start',
        to: nodeId,
        label: '最優先'
      });
    } else {
      flowchart.edges.push({
        from: `agenda_${index - 1}`,
        to: nodeId,
        label: '次の議題'
      });
    }
  });
  
  // 終了ノード
  flowchart.nodes.push({
    id: 'end',
    label: '完了',
    type: 'end'
  });
  
  if (sortedItems.length > 0) {
    flowchart.edges.push({
      from: `agenda_${sortedItems.length - 1}`,
      to: 'end',
      label: '終了'
    });
  }
  
  return flowchart;
}

// 議題項目を要約
function summarizeAgendaItems(agendaItems) {
  if (agendaItems.length === 0) {
    return '議題なし';
  }
  
  const highPriority = agendaItems.filter(item => item.priority === 'high').length;
  const mediumPriority = agendaItems.filter(item => item.priority === 'medium').length;
  const lowPriority = agendaItems.filter(item => item.priority === 'low').length;
  
  const channels = [...new Set(agendaItems.map(item => item.channel))];
  
  return `
議題総数: ${agendaItems.length}件
優先度内訳:
  - 高: ${highPriority}件
  - 中: ${mediumPriority}件
  - 低: ${lowPriority}件
関連チャンネル: ${channels.join(', ')}
`.trim();
}

// 議題レポートをメール送信（業務フローチャート付き）
function sendAgendaReportWithFlowchart(report) {
  const subject = `[Slackガバナンスレポート] ${Utilities.formatDate(report.date, 'JST', 'yyyy年MM月dd日')}`;
  
  // 重複送信チェック
  if (isDuplicateEmail('agenda_report', subject, report)) {
    logInfo('同じ内容の議題レポートが既に送信されています。スキップします。');
    return;
  }
  
  // HTMLメール本文を作成
  const htmlBody = createAgendaReportHtml(report);
  
  // プレーンテキスト版
  const plainBody = createAgendaReportPlainText(report);
  
  // メール送信
  GmailApp.sendEmail(REPORT_EMAIL, subject, plainBody, {
    htmlBody: htmlBody,
    name: 'Slackガバナンスシステム'
  });
  
  // 送信履歴を記録
  recordEmailSent('agenda_report', subject, report);
  
  logInfo(`議題レポートをメール送信: ${REPORT_EMAIL}`);
}

// HTMLレポートを作成
function createAgendaReportHtml(report) {
  const dateStr = Utilities.formatDate(report.date, 'JST', 'yyyy年MM月dd日');
  const spreadsheetUrl = report.spreadsheetUrl || `https://docs.google.com/spreadsheets/d/${SPREADSHEET_ID}/edit`;
  
  let html = `
<!DOCTYPE html>
<html>
<head>
  <style>
    body { font-family: 'Helvetica Neue', Arial, sans-serif; line-height: 1.6; color: #333; }
    .header { background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; padding: 30px; border-radius: 10px; margin-bottom: 30px; }
    .header h1 { margin: 0; font-size: 28px; }
    .summary { background: #f7f9fc; padding: 20px; border-radius: 8px; margin-bottom: 25px; }
    .spreadsheet-link { 
      display: inline-block; 
      background: #4285f4; 
      color: white; 
      padding: 12px 24px; 
      border-radius: 6px; 
      text-decoration: none; 
      font-weight: bold; 
      margin: 15px 0;
      box-shadow: 0 2px 4px rgba(0,0,0,0.2);
    }
    .spreadsheet-link:hover { background: #3367d6; }
    .action-items { 
      background: #fff3cd; 
      border: 2px solid #ffc107; 
      border-radius: 8px; 
      padding: 20px; 
      margin: 20px 0;
    }
    .action-items h3 { 
      color: #856404; 
      margin-top: 0; 
      display: flex; 
      align-items: center; 
    }
    .action-items ul { 
      margin: 10px 0; 
      padding-left: 25px;
    }
    .action-items li { 
      margin: 8px 0; 
      font-weight: 500;
    }
    .agenda-item { background: white; border: 1px solid #e1e4e8; border-radius: 8px; padding: 20px; margin-bottom: 15px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
    .priority-high { border-left: 4px solid #ff4757; }
    .priority-medium { border-left: 4px solid #ffa502; }
    .priority-low { border-left: 4px solid #5f27cd; }
    .flowchart { 
      background: #e8f5e9; 
      border: 2px solid #4caf50; 
      border-radius: 8px; 
      padding: 20px; 
      margin: 20px 0; 
    }
    .flowchart h2 { color: #2e7d32; }
    .flowchart-node { display: inline-block; padding: 10px 20px; margin: 10px; border-radius: 5px; background: white; border: 2px solid #4a5568; font-weight: 500; }
    .footer { margin-top: 30px; padding: 20px; background: #f7f9fc; border-radius: 8px; text-align: center; color: #666; }
  </style>
</head>
<body>
  <div class="header">
    <h1>📋 Slackガバナンスレポート</h1>
    <p style="margin: 10px 0 0 0; opacity: 0.9;">生成日時: ${dateStr}</p>
    <p style="margin: 5px 0 0 0; opacity: 0.9;">スプレッドシート: <a href="${spreadsheetUrl}" style="color: white; text-decoration: underline;" target="_blank">${spreadsheetUrl}</a></p>
  </div>
  
  <div class="summary">
    <h2>📊 サマリー</h2>
    <p>${report.summary.replace(/\n/g, '<br>')}</p>
    <p>分析対象メッセージ数: ${report.messageCount}件</p>
  </div>
`;
  
  // アクションアイテムを上部に配置
  const actionItems = [];
  if (report.agendaItems && report.agendaItems.length > 0) {
    report.agendaItems.forEach(item => {
      if (item.actionRequired || item.priority === 'high') {
        actionItems.push({
          title: item.title,
          priority: item.priority,
          channel: item.channel
        });
      }
    });
  }
  
  if (actionItems.length > 0) {
      html += `
    <div class="action-items">
      <h3>⚡ 緊急対応が必要なアクションアイテム</h3>
      <ul>
    `;
    actionItems.forEach(action => {
      const priorityIcon = action.priority === 'high' ? '🔴' : '🟡';
      html += `<li>${priorityIcon} ${action.title} (#${action.channel})</li>`;
    });
    html += `
      </ul>
    </div>
    `;
  }
  
  // 業務フローチャートを先に配置
  if (report.flowchart) {
    html += `
    <div class="flowchart">
      <h2>📈 業務処理フロー</h2>
      <p style="font-weight: bold; margin-bottom: 15px;">推奨される処理順序：</p>
      <div style="text-align: center; padding: 20px; background: white; border-radius: 8px;">
    `;
    
    report.flowchart.nodes.forEach((node, idx) => {
      let nodeStyle = '';
      if (node.type === 'start') nodeStyle = 'background: #48bb78; color: white;';
      else if (node.type === 'end') nodeStyle = 'background: #f56565; color: white;';
      else if (node.priority === 'high') nodeStyle = 'background: #ffebee; border-color: #ff4757; color: #c62828;';
      else if (node.priority === 'medium') nodeStyle = 'background: #fff8e1; border-color: #ffa502; color: #f57c00;';
      else if (node.priority === 'low') nodeStyle = 'background: #f3f4f6; border-color: #5f27cd;';
      
      html += `<div class="flowchart-node" style="${nodeStyle}">${node.label}</div>`;
      if (node.type !== 'end' && idx < report.flowchart.nodes.length - 1) {
        html += ' <span style="font-size: 20px; color: #4caf50; font-weight: bold;">→</span> ';
      }
    });
    
    html += `
      </div>
    </div>
    `;
  }
  
  // 議題一覧
  if (report.agendaItems && report.agendaItems.length > 0) {
    html += '<h2>🎯 抽出された議題詳細</h2>';
    report.agendaItems.forEach((item, index) => {
      html += `
      <div class="agenda-item priority-${item.priority}">
        <h3>${index + 1}. ${item.title}</h3>
        <p><strong>説明:</strong> ${item.description}</p>
        <p><strong>優先度:</strong> ${item.priority.toUpperCase()}</p>
        <p><strong>チャンネル:</strong> #${item.channel}</p>
        ${item.participants.length > 0 ? `<p><strong>関係者:</strong> ${item.participants.join(', ')}</p>` : ''}
        ${item.keywords.length > 0 ? `<p><strong>キーワード:</strong> ${item.keywords.join(', ')}</p>` : ''}
      </div>
      `;
    });
  }
  
  html += `
  <div class="footer">
    <p>このレポートは自動生成されました</p>
    <p style="font-size: 12px; margin-top: 10px;">
      <a href="${spreadsheetUrl}" style="color: #4285f4; text-decoration: none;">📊 スプレッドシートで詳細を確認</a>
    </p>
    <p style="font-size: 12px; margin-top: 5px;">Slackガバナンスシステム v1.0</p>
  </div>
</body>
</html>
  `;
  
  return html;
}

// プレーンテキストレポートを作成
function createAgendaReportPlainText(report) {
  const dateStr = Utilities.formatDate(report.date, 'JST', 'yyyy年MM月dd日');
  let text = `Slackガバナンスレポート - ${dateStr}\n`;
  text += '=' .repeat(50) + '\n\n';
  
  text += '【サマリー】\n';
  text += report.summary + '\n';
  text += `分析対象メッセージ数: ${report.messageCount}件\n\n`;
  
  if (report.agendaItems && report.agendaItems.length > 0) {
    text += '【抽出された議題】\n';
    report.agendaItems.forEach((item, index) => {
      text += `\n${index + 1}. ${item.title}\n`;
      text += `   説明: ${item.description}\n`;
      text += `   優先度: ${item.priority.toUpperCase()}\n`;
      text += `   チャンネル: #${item.channel}\n`;
      if (item.participants.length > 0) {
        text += `   関係者: ${item.participants.join(', ')}\n`;
      }
      if (item.keywords.length > 0) {
        text += `   キーワード: ${item.keywords.join(', ')}\n`;
      }
    });
  }
  
  if (report.flowchart) {
    text += '\n【業務フロー】\n';
    text += '処理順序: ';
    text += report.flowchart.nodes.map(node => node.label).join(' → ');
    text += '\n';
  }
  
  return text;
}

// スプレッドシートに議題分析結果を記録
function recordAgendaAnalysis(ss, report) {
  const sheetName = 'AgendaAnalysis';
  let sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.getRange(1, 1, 1, 8).setValues([[
      '分析日時', '対象メッセージ数', '議題数', '高優先度', '中優先度', '低優先度', '関連チャンネル', 'メール送信'
    ]]);
    sheet.getRange(1, 1, 1, 8).setFontWeight('bold');
  }
  
  const highCount = report.agendaItems.filter(item => item.priority === 'high').length;
  const mediumCount = report.agendaItems.filter(item => item.priority === 'medium').length;
  const lowCount = report.agendaItems.filter(item => item.priority === 'low').length;
  const channels = [...new Set(report.agendaItems.map(item => item.channel))].join(', ');
  
  sheet.appendRow([
    report.date,
    report.messageCount,
    report.agendaItems.length,
    highCount,
    mediumCount,
    lowCount,
    channels,
    REPORT_EMAIL ? 'YES' : 'NO'
  ]);
}

// 新規スプレッドシートを作成してレポートをエクスポート
function exportReportToNewSpreadsheet(report) {
  const dateStr = Utilities.formatDate(report.date, 'JST', 'yyyy_MM_dd_HHmm');
  const spreadsheetName = `Slackガバナンスレポート_${dateStr}`;
  
  // 新規スプレッドシートを作成
  const newSpreadsheet = SpreadsheetApp.create(spreadsheetName);
  const newSpreadsheetId = newSpreadsheet.getId();
  const newSpreadsheetUrl = newSpreadsheet.getUrl();
  
  logInfo(`新規スプレッドシート作成: ${newSpreadsheetUrl}`);
  
  // 1. サマリーシートを作成
  const summarySheet = newSpreadsheet.getActiveSheet();
  summarySheet.setName('サマリー');
  
  // サマリー情報を書き込み
  const summaryData = [
    ['Slackガバナンスレポート', ''],
    ['生成日時', Utilities.formatDate(report.date, 'JST', 'yyyy年MM月dd日 HH:mm')],
    ['分析対象メッセージ数', report.messageCount],
    ['抽出された議題数', report.agendaItems.length],
    ['', ''],
    ['優先度別内訳', ''],
    ['高優先度', report.agendaItems.filter(item => item.priority === 'high').length],
    ['中優先度', report.agendaItems.filter(item => item.priority === 'medium').length],
    ['低優先度', report.agendaItems.filter(item => item.priority === 'low').length],
    ['', ''],
    ['サマリー', ''],
    [report.summary.replace(/\n/g, '\n'), '']
  ];
  
  summarySheet.getRange(1, 1, summaryData.length, 2).setValues(summaryData);
  summarySheet.getRange(1, 1).setFontSize(16).setFontWeight('bold');
  summarySheet.getRange(6, 1).setFontWeight('bold');
  summarySheet.getRange(11, 1).setFontWeight('bold');
  summarySheet.getRange(12, 1).setWrap(true);
  summarySheet.setColumnWidth(1, 200);
  summarySheet.setColumnWidth(2, 400);
  
  // 2. 議題詳細シートを作成
  if (report.agendaItems && report.agendaItems.length > 0) {
    const agendaSheet = newSpreadsheet.insertSheet('議題詳細');
    
    // ヘッダーを設定
    const headers = ['優先度', 'タイトル', '詳細', 'アクション', 'チャンネル', '関連メッセージ'];
    agendaSheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    
    // 議題データを書き込み
    const agendaData = report.agendaItems.map(item => [
      item.priority || '',
      item.title || '',
      item.description || '',
      item.actionRequired || '',
      item.channel || '',
      item.messages ? item.messages.join('\n') : ''
    ]);
    
    if (agendaData.length > 0) {
      agendaSheet.getRange(2, 1, agendaData.length, headers.length).setValues(agendaData);
      agendaSheet.getRange(2, 6, agendaData.length, 1).setWrap(true);
      
      // 優先度に応じて色を設定
      for (let i = 0; i < agendaData.length; i++) {
        const priority = agendaData[i][0];
        const rowRange = agendaSheet.getRange(i + 2, 1, 1, headers.length);
        
        if (priority === 'high') {
          rowRange.setBackground('#ffebee');
        } else if (priority === 'medium') {
          rowRange.setBackground('#fff3e0');
        } else if (priority === 'low') {
          rowRange.setBackground('#e8f5e9');
        }
      }
      
      // 列幅を調整
      agendaSheet.setColumnWidth(1, 80);
      agendaSheet.setColumnWidth(2, 250);
      agendaSheet.setColumnWidth(3, 350);
      agendaSheet.setColumnWidth(4, 250);
      agendaSheet.setColumnWidth(5, 150);
      agendaSheet.setColumnWidth(6, 300);
    }
  }
  
  // 3. 業務フローシートを作成（フローチャートがある場合）
  if (report.flowchart) {
    const flowSheet = newSpreadsheet.insertSheet('業務フロー');
    flowSheet.getRange(1, 1).setValue('業務フローチャート').setFontSize(16).setFontWeight('bold');
    flowSheet.getRange(3, 1).setValue(report.flowchart).setWrap(true);
    flowSheet.setColumnWidth(1, 800);
  }
  
  // 4. ビジュアルフローシートを作成
  try {
    const visualFlowSheet = newSpreadsheet.insertSheet('ビジュアルフロー');
    generateVisualFlowFromAgenda(visualFlowSheet, report.agendaItems);
    logInfo('ビジュアルフローシートを作成しました');
  } catch (error) {
    logError('ビジュアルフローシート作成エラー', error.toString());
    // エラーが発生しても処理を継続
  }
  
  // 全員編集可能に設定
  try {
    newSpreadsheet.addEditor(Session.getActiveUser().getEmail());
  } catch (e) {
    logInfo('スプレッドシートの編集権限設定をスキップ: ' + e.toString());
  }
  
  return newSpreadsheetUrl;
}

// ========= フロービジュアライザー機能 =========

// 高度なカラーパレット定義
const ADVANCED_COLORS = {
  // ヘッダー系
  MAIN_HEADER: '#2C3E50',
  SUB_HEADER: '#34495E',
  SECTION_HEADER: '#5D6D7E',
  
  // プロセス系
  START_END: '#27AE60',
  PROCESS: '#3498DB',
  DECISION: '#F39C12',
  SUBPROCESS: '#9B59B6',
  
  // 背景系
  TIMELINE_BG: '#ECF0F1',
  DEPT_BG: '#E8F5E9',
  EMPTY_BG: '#FAFAFA',
  HIGHLIGHT_BG: '#FFF9C4',
  
  // ツール系
  TOOL_BG: '#E3F2FD',
  DATASOURCE_BG: '#FFF3E0',
  
  // ステータス系
  SUCCESS: '#4CAF50',
  WARNING: '#FF9800',
  ERROR: '#F44336',
  INFO: '#2196F3'
};

// 部署別カラーパレット
const DEPT_COLOR_PALETTE = [
  '#E3F2FD', '#FCE4EC', '#F3E5F5', '#EDE7F6', '#E8EAF6',
  '#E1F5FE', '#E0F2F1', '#E8F5E9', '#F9FBE7', '#FFF8E1',
  '#FFF3E0', '#FBE9E7', '#EFEBE9', '#FAFAFA', '#ECEFF1'
];

// ツール別アイコンとカラー
const TOOL_ICONS = {
  'Word': { icon: '📝', color: '#2B579A' },
  'Excel': { icon: '📊', color: '#217346' },
  'PowerPoint': { icon: '📰', color: '#D24726' },
  'PPT': { icon: '📰', color: '#D24726' },
  'Teams': { icon: '👥', color: '#5B5FC7' },
  'Outlook': { icon: '📧', color: '#0078D4' },
  'Gmail': { icon: '📨', color: '#EA4335' },
  'Slack': { icon: '💬', color: '#4A154B' },
  'GitHub': { icon: '🐙', color: '#24292E' },
  'Jira': { icon: '📋', color: '#0052CC' },
  'Notion': { icon: '📓', color: '#000000' },
  'Google Drive': { icon: '☁️', color: '#4285F4' },
  'Zoom': { icon: '📹', color: '#2D8CFF' },
  'メール': { icon: '✉️', color: '#EA4335' },
  'ブラウザ': { icon: '🌐', color: '#4CAF50' },
  'システム': { icon: '⚙️', color: '#607D8B' },
  'データベース': { icon: '🗄️', color: '#FF6F00' }
};

// 議題からビジュアルフローを生成
function generateVisualFlowFromAgenda(visualSheet, agendaItems) {
  if (!agendaItems || agendaItems.length === 0) {
    logInfo('議題データがないためビジュアルフロー作成をスキップ');
    return;
  }
  
  // ビジュアルフローシートをクリア
  visualSheet.clear();
  visualSheet.clearFormats();
  
  // フローデータを議題から構築
  const flowData = convertAgendaToFlowData(agendaItems);
  
  // 高度なフローチャートを描画
  drawAdvancedFlowChart(visualSheet, flowData);
  
  logInfo('議題ベースのビジュアルフローを生成しました');
}

// 議題データをフローデータに変換
function convertAgendaToFlowData(agendaItems) {
  const flowData = {
    departments: {},
    departmentList: [],
    timings: [],
    tools: new Set(),
    datasources: {},
    processName: "議題フロー",
    statistics: {
      totalTasks: 0,
      totalDepartments: 0,
      totalTools: 0,
      decisionPoints: 0
    }
  };
  
  // 優先度別にグループ化
  const priorities = ['high', 'medium', 'low'];
  const priorityLabels = {
    'high': '高優先度議題',
    'medium': '中優先度議題',
    'low': '低優先度議題'
  };
  
  priorities.forEach((priority, index) => {
    const items = agendaItems.filter(item => item.priority === priority);
    if (items.length === 0) return;
    
    const timing = priorityLabels[priority];
    flowData.timings.push(timing);
    
    items.forEach(item => {
      // チャンネルを部署として扱う
      const dept = item.channel || 'その他';
      
      if (!flowData.departments[dept]) {
        flowData.departments[dept] = {};
        if (!flowData.departmentList.includes(dept)) {
          flowData.departmentList.push(dept);
        }
      }
      
      if (!flowData.departments[dept][timing]) {
        flowData.departments[dept][timing] = [];
      }
      
      flowData.departments[dept][timing].push({
        task: item.title || '',
        role: item.participants ? item.participants.join(', ') : '',
        condition: item.actionRequired ? 'アクション必要' : '',
        tool: 'Slack',
        url: '',
        note: item.description || ''
      });
      
      flowData.statistics.totalTasks++;
      if (item.actionRequired) {
        flowData.statistics.decisionPoints++;
      }
      
      flowData.tools.add('Slack');
    });
  });
  
  flowData.statistics.totalDepartments = flowData.departmentList.length;
  flowData.statistics.totalTools = flowData.tools.size;
  
  return flowData;
}

// 高度なフローチャート描画（議題用にカスタマイズ）
function drawAdvancedFlowChart(sheet, flowData) {
  let currentRow = 1;
  const maxCols = Math.max(flowData.departmentList.length + 2, 10);
  
  // タイトル行
  const flowTitle = flowData.processName || "業務フロー";
  sheet.getRange(currentRow, 1, 1, maxCols).merge();
  const titleCell = sheet.getRange(currentRow, 1);
  titleCell.setValue(flowTitle + "（" + new Date().getFullYear() + "年" + (new Date().getMonth() + 1) + "月版）");
  titleCell.setBackground(ADVANCED_COLORS.MAIN_HEADER);
  titleCell.setFontColor("#FFFFFF");
  titleCell.setFontSize(18);
  titleCell.setFontWeight("bold");
  titleCell.setHorizontalAlignment("center");
  titleCell.setVerticalAlignment("middle");
  sheet.setRowHeight(currentRow, 50);
  currentRow++;
  
  // 統計情報行
  sheet.getRange(currentRow, 1, 1, maxCols).merge();
  const statsCell = sheet.getRange(currentRow, 1);
  statsCell.setValue(`📊 総議題: ${flowData.statistics.totalTasks} | 📍 チャンネル: ${flowData.statistics.totalDepartments} | ⚡ アクション必要: ${flowData.statistics.decisionPoints}`);
  statsCell.setBackground(ADVANCED_COLORS.SUB_HEADER);
  statsCell.setFontColor("#FFFFFF");
  statsCell.setHorizontalAlignment("center");
  sheet.setRowHeight(currentRow, 35);
  currentRow++;
  
  // ヘッダー行
  drawAdvancedHeaderRow(sheet, currentRow, flowData.departmentList, false);
  currentRow++;
  
  // 開始行
  drawAdvancedStartRow(sheet, currentRow, maxCols);
  currentRow++;
  
  // 各タイミング（優先度）の行
  flowData.timings.forEach((timing, index) => {
    drawAdvancedTimingRow(sheet, currentRow, timing, flowData, index);
    currentRow++;
  });
  
  // 終了行
  drawAdvancedEndRow(sheet, currentRow, flowData.departmentList.length, false);
  currentRow++;
  
  // 凡例行
  drawAdvancedLegendRow(sheet, currentRow, maxCols);
  
  // 列幅の調整
  adjustAdvancedColumnWidths(sheet, flowData.departmentList.length, false);
  
  // 罫線の設定
  applyAdvancedBorders(sheet, currentRow, maxCols);
}

// ヘッダー行の描画
function drawAdvancedHeaderRow(sheet, row, departments, hasDataSource) {
  // 優先度列
  sheet.getRange(row, 1).setValue("優先度");
  sheet.getRange(row, 1).setBackground(ADVANCED_COLORS.SECTION_HEADER);
  sheet.getRange(row, 1).setFontColor("#FFFFFF");
  sheet.getRange(row, 1).setFontWeight("bold");
  
  // チャンネル列
  departments.forEach((dept, index) => {
    const col = index + 2;
    const cell = sheet.getRange(row, col);
    cell.setValue('#' + dept);
    cell.setBackground(DEPT_COLOR_PALETTE[index % DEPT_COLOR_PALETTE.length]);
    cell.setFontWeight("bold");
    cell.setWrap(true);
    cell.setHorizontalAlignment("center");
    cell.setVerticalAlignment("middle");
  });
  
  sheet.setRowHeight(row, 50);
}

// 開始行の描画
function drawAdvancedStartRow(sheet, row, maxCols) {
  sheet.getRange(row, 1, 1, maxCols).merge();
  const cell = sheet.getRange(row, 1);
  cell.setValue("🚀 【議題分析開始】");
  cell.setBackground(ADVANCED_COLORS.START_END);
  cell.setFontColor("#FFFFFF");
  cell.setFontWeight("bold");
  cell.setFontSize(14);
  cell.setHorizontalAlignment("center");
  cell.setBorder(true, true, true, true, false, false, "#228B22", SpreadsheetApp.BorderStyle.SOLID_THICK);
  sheet.setRowHeight(row, 40);
}

// タイミング行の描画
function drawAdvancedTimingRow(sheet, row, timing, flowData, timingIndex) {
  // タイミング（優先度）列
  const timingCell = sheet.getRange(row, 1);
  timingCell.setValue(timing);
  
  // 優先度に応じて背景色を設定
  if (timing.includes('高優先度')) {
    timingCell.setBackground('#ffcdd2');
  } else if (timing.includes('中優先度')) {
    timingCell.setBackground('#fff9c4');
  } else {
    timingCell.setBackground('#c8e6c9');
  }
  timingCell.setFontWeight("bold");
  timingCell.setWrap(true);
  
  // 各チャンネルの議題
  flowData.departmentList.forEach((dept, deptIndex) => {
    const col = deptIndex + 2;
    const tasks = flowData.departments[dept][timing];
    
    if (tasks && tasks.length > 0) {
      const cell = sheet.getRange(row, col);
      
      // 複数の議題をまとめて表示
      let content = tasks.map((task, index) => {
        let taskText = `${index + 1}. ${task.task}`;
        if (task.role) {
          taskText += `\n   👥 ${task.role}`;
        }
        if (task.note) {
          taskText += `\n   📝 ${task.note.substring(0, 50)}...`;
        }
        return taskText;
      }).join('\n\n');
      
      cell.setValue(content);
      cell.setWrap(true);
      cell.setHorizontalAlignment("left");
      cell.setVerticalAlignment("top");
      
      // アクション必要な議題は強調
      const hasAction = tasks.some(task => task.condition === 'アクション必要');
      if (hasAction) {
        cell.setBackground(ADVANCED_COLORS.DECISION);
        cell.setBorder(true, true, true, true, false, false, "#FF8C00", SpreadsheetApp.BorderStyle.SOLID_THICK);
        cell.setFontWeight("bold");
      } else {
        cell.setBackground(ADVANCED_COLORS.PROCESS);
        cell.setBorder(true, true, true, true, false, false, "#4682B4", SpreadsheetApp.BorderStyle.SOLID_THICK);
      }
    } else {
      // 空のセル
      const cell = sheet.getRange(row, col);
      cell.setBackground(ADVANCED_COLORS.EMPTY_BG);
    }
  });
  
  // 行の高さを議題数に応じて調整
  const maxTasks = Math.max(...flowData.departmentList.map(dept => 
    (flowData.departments[dept][timing] || []).length
  ));
  sheet.setRowHeight(row, Math.max(90, 60 * maxTasks));
}

// 終了行の描画
function drawAdvancedEndRow(sheet, row, deptCount, hasDataSource) {
  const mergeCols = deptCount + 1;
  sheet.getRange(row, 1, 1, mergeCols).merge();
  const cell = sheet.getRange(row, 1);
  cell.setValue("✅ 【議題分析完了】");
  cell.setBackground(ADVANCED_COLORS.START_END);
  cell.setFontColor("#FFFFFF");
  cell.setFontSize(14);
  cell.setFontWeight("bold");
  cell.setHorizontalAlignment("center");
  cell.setBorder(true, true, true, true, false, false, "#228B22", SpreadsheetApp.BorderStyle.SOLID_THICK);
  sheet.setRowHeight(row, 50);
}

// 凡例行の描画
function drawAdvancedLegendRow(sheet, row, maxCols) {
  sheet.getRange(row, 1, 1, maxCols).merge();
  const legendCell = sheet.getRange(row, 1);
  legendCell.setValue("【凡例】 📦 議題　⚡ アクション必要　👥 関係者　📝 詳細");
  legendCell.setBackground(ADVANCED_COLORS.TIMELINE_BG);
  legendCell.setFontWeight("bold");
  legendCell.setHorizontalAlignment("left");
  legendCell.setVerticalAlignment("middle");
  sheet.setRowHeight(row, 40);
}

// 列幅の調整
function adjustAdvancedColumnWidths(sheet, deptCount, hasDataSource) {
  sheet.setColumnWidth(1, 150); // 優先度列
  for (let i = 2; i <= deptCount + 1; i++) {
    sheet.setColumnWidth(i, 300); // チャンネル列（議題が多いため広めに）
  }
}

// 罫線の設定
function applyAdvancedBorders(sheet, lastRow, lastCol) {
  const range = sheet.getRange(1, 1, lastRow, lastCol);
  range.setBorder(true, true, true, true, true, true, "#d0d0d0", SpreadsheetApp.BorderStyle.SOLID);
}

// ========= デバッグ・設定確認関数 =========

/**
 * Config設定内容を確認
 */
function checkConfigSettings() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const configSheet = ss.getSheetByName(SHEETS.CONFIG);
  
  if (!configSheet) {
    ui.alert('エラー', 'Configシートが見つかりません。', ui.ButtonSet.OK);
    return;
  }
  
  // まず生のデータを確認
  const rawData = configSheet.getDataRange().getValues();
  let rawMessage = '=== Configシートの生データ ===\n';
  for (let i = 0; i < Math.min(rawData.length, 15); i++) {
    rawMessage += `行${i+1}: [${rawData[i][0]}] = [${rawData[i][1]}]\n`;
  }
  console.log(rawMessage);
  
  const config = getConfigData(configSheet);
  
  let message = '=== 現在の設定内容 ===\n\n';
  message += `会社名: ${config.company || '未設定'}\n`;
  message += `監視対象チャンネル: ${config.targetChannels ? config.targetChannels.join(', ') : '未設定'}\n`;
  message += `通知先Slackチャンネル: ${config.notifySlackChannel || '未設定'}\n`;
  message += `通知先メール: ${config.notifyEmails ? config.notifyEmails.join(', ') : '未設定'}\n`;
  message += `AIモデル（要約・分類）: ${config.openaiModel || '未設定'}\n`;
  message += `AIモデル（ドラフト生成）: ${config.openaiModelDraft || '未設定'}\n`;
  message += `AIモデル（メイン処理）: ${config.OPENAI_MODEL || '未設定'}\n`;
  message += `判定しきい値: ${config.classificationThreshold || '未設定'}\n`;
  
  // チャンネルIDの形式チェック
  if (config.targetChannels && config.targetChannels.length > 0) {
    message += '\n--- チャンネルID形式チェック ---\n';
    config.targetChannels.forEach(channelId => {
      if (channelId.startsWith('C') || channelId.startsWith('G')) {
        message += `✅ ${channelId}: 正しい形式\n`;
      } else {
        message += `❌ ${channelId}: 不正な形式（CまたはGで始まる必要があります）\n`;
      }
    });
  } else {
    message += '\n⚠️ 監視対象チャンネルが設定されていません\n';
    message += '設定シートの「targetChannels」行にチャンネルIDを入力してください\n';
    message += '（複数の場合はカンマ区切り）\n';
  }
  
  ui.alert('設定確認', message, ui.ButtonSet.OK);
}

// ========= テスト関数 =========
function testSystem() {
  SpreadsheetApp.getUi().alert('統合システムが正常に動作しています');
}




/*
================================================================================
                    タスク抽出・管理シート作成 - 統合ファイル
================================================================================

【目次 - Table of Contents】
1. config.gs - 基本設定とユーティリティ機能
2. flow_visualizer.gs - フロービジュアライザー機能
3. gmail_inbound.gs - Gmail受信処理機能
4. gmail_outbound.gs - Gmail送信処理機能
5. menu.gs - カスタムメニューとセットアップ機能
6. openai_client.gs - OpenAI API連携機能
7. parser_and_writer.gs - データ解析・書き込み処理機能

作成日: 2025-08-16
説明: Google Apps Scriptによるタスク管理システムの全ソースコード

================================================================================
*/

// ================================================================================
// 1. config.gs - 基本設定とユーティリティ機能
// ================================================================================

// 基本定数
const CONFIG_SHEET = 'Config';
const INBOX_SHEET = 'Inbox';
const SPEC_SHEET = '業務記述書';
const FLOW_SHEET = 'フロー';
const VISUAL_SHEET = 'ビジュアルフロー';

// CSV行を解析する関数
function parseCSVLine(line) {
  const result = [];
  let current = '';
  let inQuotes = false;
  
  for (let i = 0; i < line.length; i++) {
    const char = line[i];
    const nextChar = line[i + 1];
    
    if (char === '"') {
      if (inQuotes && nextChar === '"') {
        // エスケープされた引用符
        current += '"';
        i++; // 次の引用符をスキップ
      } else {
        // 引用符の開始/終了
        inQuotes = !inQuotes;
      }
    } else if (char === ',' && !inQuotes) {
      // フィールドの区切り
      result.push(current.trim());
      current = '';
    } else {
      current += char;
    }
  }
  
  // 最後のフィールドを追加
  result.push(current.trim());
  
  return result;
}
const ACTIVITY_LOG_SHEET = 'ActivityLog';

// スプレッドシート関連ユーティリティ
function ss() {
  return SpreadsheetApp.getActive();
}

function file() {
  return DriveApp.getFileById(ss().getId());
}

// Config管理
function getConfig(key) {
  const sh = ss().getSheetByName(CONFIG_SHEET);
  if (!sh) return null;
  
  const values = sh.getRange(1, 1, sh.getLastRow(), 2).getValues();
  const configMap = new Map(values.map(r => [String(r[0]).trim(), String(r[1]).trim()]));
  return configMap.get(key);
}

function setConfig(key, value) {
  const sh = ss().getSheetByName(CONFIG_SHEET);
  if (!sh) return;
  
  const values = sh.getRange(1, 1, sh.getLastRow(), 2).getValues();
  let found = false;
  
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]).trim() === key) {
      sh.getRange(i + 1, 2).setValue(value);
      found = true;
      break;
    }
  }
  
  if (!found) {
    sh.appendRow([key, value]);
  }
}

// 初期Config設定
function initializeConfig() {
  const sh = ss().getSheetByName(CONFIG_SHEET) || ss().insertSheet(CONFIG_SHEET);
  
  const defaultConfigs = [
    ['PROCESSING_QUERY', 'subject:[task] is:unread'],
    ['DEFAULT_TO_EMAIL', ''],
    ['OPENAI_MODEL', 'gpt-5'],  // gpt-5をデフォルトで使用
    ['ORG_PROFILE_JSON', '{"listing":"上場区分","industry":"業種","jurisdictions":["JP"],"policies":["内部統制準拠"]}'],
    ['SHARE_ANYONE_WITH_LINK', 'TRUE'],
    ['FLOW_SHEET_NAME', 'フロー'],
    ['VISUAL_SHEET_NAME', 'ビジュアルフロー'],
    ['LEGAL_JURISDICTIONS', 'JP, global'],
    ['MAX_RETRY_COUNT', '3'],
    ['RETRY_DELAY_MS', '2000']
  ];
  
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, defaultConfigs.length, 2).setValues(defaultConfigs);
  }
}

// 共有設定
function shareSheetAnyWithLink() {
  try {
    file().setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.EDIT);
    logActivity('SHARE', 'Set ANYONE_WITH_LINK permissions');
    return true;
  } catch (e) {
    logActivity('SHARE_ERROR', `Failed to set ANYONE_WITH_LINK: ${e.toString()}`);
    return false;
  }
}

function addEditor(email) {
  try {
    file().addEditor(email);
    logActivity('SHARE', `Added editor: ${email}`);
    return true;
  } catch (e) {
    logActivity('SHARE_ERROR', `Failed to add editor ${email}: ${e.toString()}`);
    return false;
  }
}

// メールアドレス抽出
function extractEmail(fromHeader) {
  const match = fromHeader.match(/<([^>]+)>/);
  return match ? match[1] : fromHeader.replace(/.*\s/, '').trim();
}

// HTMLをテキストに変換
function htmlToText(html) {
  return html
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<\/p>/gi, '\n\n')
    .replace(/<[^>]+>/g, ' ')
    .replace(/&nbsp;/g, ' ')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&amp;/g, '&')
    .replace(/\s+/g, ' ')
    .trim();
}

// メール署名を除去
function removeEmailSignature(text) {
  if (!text) return text;
  
  // 一般的な署名の開始パターン
  const signaturePatterns = [
    /--\s*\n/,  // -- で始まる署名
    /—\s*\n/,   // — で始まる署名
    /＿+\s*\n/, // アンダースコアの連続
    /━+\s*\n/,  // 罫線
    /※この.*$/s, // ※この...で始まる免責事項
    /\n\n(敬具|よろしくお願い|Best regards|Regards|Sincerely|Thanks)[\s\S]*$/i,
    /\n\n-{3,}[\s\S]*$/, // 3つ以上のハイフン
    /\n\n_{3,}[\s\S]*$/, // 3つ以上のアンダースコア
    /\n\n={3,}[\s\S]*$/, // 3つ以上のイコール
    /\n\n\*{3,}[\s\S]*$/  // 3つ以上のアスタリスク
  ];
  
  let cleanedText = text;
  
  // 各パターンでマッチする最初の位置を探す
  let earliestIndex = text.length;
  for (const pattern of signaturePatterns) {
    const match = text.match(pattern);
    if (match && match.index < earliestIndex) {
      earliestIndex = match.index;
    }
  }
  
  // 署名部分を除去
  if (earliestIndex < text.length) {
    cleanedText = text.substring(0, earliestIndex).trim();
  }
  
  // 連絡先情報のパターンも除去（メールアドレス、電話番号が連続する部分）
  const contactPattern = /(\n.*[@].*\n.*[0-9-()]+.*\n)/g;
  cleanedText = cleanedText.replace(contactPattern, '\n');
  
  return cleanedText.trim();
}

// 処理済みチェック
function isProcessed(messageId) {
  const sh = ss().getSheetByName(INBOX_SHEET);
  if (!sh || sh.getLastRow() <= 1) return false;
  
  const lastRow = sh.getLastRow();
  const dataRows = Math.max(1, lastRow - 1);
  const values = sh.getRange(2, 3, dataRows, 1).getValues();
  return values.some(row => row[0] === messageId);
}

function markProcessed(messageId) {
  const sh = ss().getSheetByName(INBOX_SHEET);
  if (!sh || sh.getLastRow() <= 1) return;
  
  const lastRow = sh.getLastRow();
  const dataRows = Math.max(1, lastRow - 1);
  const values = sh.getRange(2, 3, dataRows, 1).getValues();
  for (let i = 0; i < values.length; i++) {
    if (values[i][0] === messageId) {
      sh.getRange(i + 2, 7).setValue('PROCESSED');
      sh.getRange(i + 2, 8).setValue(new Date());
      break;
    }
  }
}

// Inboxログ記録
function logInbox(messageId, threadId, from, subject, summary, status) {
  const sh = ss().getSheetByName(INBOX_SHEET) || createInboxSheet();
  sh.appendRow([
    new Date(),
    threadId,
    messageId,
    from,
    subject,
    summary,
    status,
    status === 'PROCESSED' ? new Date() : '',
    ''
  ]);
}

function createInboxSheet() {
  const sh = ss().insertSheet(INBOX_SHEET);
  sh.getRange(1, 1, 1, 9).setValues([[
    '受信日時', 'ThreadId', 'MessageId', 'From', 'Subject', 
    '要約', 'ステータス', '処理日時', 'エラー'
  ]]);
  sh.getRange(1, 1, 1, 9).setFontWeight('bold');
  return sh;
}

// エラーログ記録
function logError(messageId, error) {
  const sh = ss().getSheetByName(INBOX_SHEET);
  if (!sh || sh.getLastRow() <= 1) return;
  
  const lastRow = sh.getLastRow();
  const dataRows = Math.max(1, lastRow - 1);
  const values = sh.getRange(2, 3, dataRows, 1).getValues();
  for (let i = 0; i < values.length; i++) {
    if (values[i][0] === messageId) {
      sh.getRange(i + 2, 7).setValue('ERROR');
      sh.getRange(i + 2, 9).setValue(error.toString());
      break;
    }
  }
  
  logActivity('ERROR', `MessageId: ${messageId}, Error: ${error.toString()}`);
}

// アクティビティログ
function logActivity(type, details) {
  const sh = ss().getSheetByName(ACTIVITY_LOG_SHEET) || createActivityLogSheet();
  let userEmail;
  try {
    userEmail = Session.getActiveUser().getEmail();
  } catch (e) {
    userEmail = 'Unknown User';
  }
  sh.appendRow([
    new Date(),
    type,
    details,
    userEmail
  ]);
}

function createActivityLogSheet() {
  // 既存のシートがないか再確認
  let sh = ss().getSheetByName(ACTIVITY_LOG_SHEET);
  if (sh) {
    console.log('ActivityLogシートは既に存在します');
    return sh;
  }
  
  try {
    sh = ss().insertSheet(ACTIVITY_LOG_SHEET);
    sh.getRange(1, 1, 1, 4).setValues([[
      'タイムスタンプ', 'タイプ', '詳細', 'ユーザー'
    ]]);
    sh.getRange(1, 1, 1, 4).setFontWeight('bold');
    sh.hideSheet();
    console.log('ActivityLogシートを作成しました');
  } catch (e) {
    console.error('ActivityLogシート作成エラー:', e.toString());
    // エラーが発生した場合は、既存のシートを探す
    const sheets = ss().getSheets();
    for (const sheet of sheets) {
      if (sheet.getName().toLowerCase() === 'activitylog') {
        console.log('ActivityLogシートが別の形式で存在しています');
        return sheet;
      }
    }
  }
  
  return sh;
}

// リトライ処理
function retryWithBackoff(func, maxRetries = 3) {
  const configRetries = parseInt(getConfig('MAX_RETRY_COUNT') || '3');
  const retryDelay = parseInt(getConfig('RETRY_DELAY_MS') || '2000');
  maxRetries = configRetries;
  
  for (let i = 0; i < maxRetries; i++) {
    try {
      return func();
    } catch (e) {
      if (i === maxRetries - 1) throw e;
      Utilities.sleep(retryDelay * Math.pow(2, i));
    }
  }
}



// 高度なカラーパレット定義
const ADVANCED_COLORS = {
  // ヘッダー系
  MAIN_HEADER: '#2C3E50',
  SUB_HEADER: '#34495E',
  SECTION_HEADER: '#5D6D7E',
  
  // プロセス系
  START_END: '#27AE60',
  PROCESS: '#3498DB',
  DECISION: '#F39C12',
  SUBPROCESS: '#9B59B6',
  
  // 背景系
  TIMELINE_BG: '#ECF0F1',
  DEPT_BG: '#E8F5E9',
  EMPTY_BG: '#FAFAFA',
  HIGHLIGHT_BG: '#FFF9C4',
  
  // ツール系
  TOOL_BG: '#E3F2FD',
  DATASOURCE_BG: '#FFF3E0',
  
  // ステータス系
  SUCCESS: '#4CAF50',
  WARNING: '#FF9800',
  ERROR: '#F44336',
  INFO: '#2196F3'
};

// 部署別カラーパレット
const DEPT_COLOR_PALETTE = [
  '#E3F2FD', '#FCE4EC', '#F3E5F5', '#EDE7F6', '#E8EAF6',
  '#E1F5FE', '#E0F2F1', '#E8F5E9', '#F9FBE7', '#FFF8E1',
  '#FFF3E0', '#FBE9E7', '#EFEBE9', '#FAFAFA', '#ECEFF1'
];

// ツール別アイコンとカラー
const TOOL_ICONS = {
  'Word': { icon: '📝', color: '#2B579A' },
  'Excel': { icon: '📊', color: '#217346' },
  'PowerPoint': { icon: '📰', color: '#D24726' },
  'PPT': { icon: '📰', color: '#D24726' },
  'Teams': { icon: '👥', color: '#5B5FC7' },
  'Outlook': { icon: '📧', color: '#0078D4' },
  'Gmail': { icon: '📨', color: '#EA4335' },
  'Slack': { icon: '💬', color: '#4A154B' },
  'GitHub': { icon: '🐙', color: '#24292E' },
  'Jira': { icon: '📋', color: '#0052CC' },
  'Notion': { icon: '📓', color: '#000000' },
  'Google Drive': { icon: '☁️', color: '#4285F4' },
  'Zoom': { icon: '📹', color: '#2D8CFF' },
  'メール': { icon: '✉️', color: '#EA4335' },
  'ブラウザ': { icon: '🌐', color: '#4CAF50' },
  'システム': { icon: '⚙️', color: '#607D8B' },
  'データベース': { icon: '🗄️', color: '#FF6F00' }
};

// フロービジュアライザー - ビジュアルフロー生成機能

function generateVisualFlow() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const flowSheet = ss.getSheetByName(getConfig('FLOW_SHEET_NAME') || 'フロー');
  const visualSheet = ss.getSheetByName(getConfig('VISUAL_SHEET_NAME') || 'ビジュアルフロー') || 
                      ss.insertSheet(getConfig('VISUAL_SHEET_NAME') || 'ビジュアルフロー');
  
  if (!flowSheet) {
    console.error('フローシートが見つかりません。');
    return;
  }
  
  // ビジュアルフローシートをクリア
  visualSheet.clear();
  visualSheet.clearFormats();
  
  // フローデータを取得
  const flowData = flowSheet.getDataRange().getValues();
  if (flowData.length <= 1) {
    console.error('フローデータがありません。');
    return;
  }
  
  const headers = flowData[0];
  const rows = flowData.slice(1).filter(row => row[0]); // 工程が入力されている行のみ
  
  if (rows.length === 0) {
    console.error('有効なフローデータがありません。');
    return;
  }
  
  // 部署リストを作成
  const departments = [...new Set(rows.map(row => row[2]).filter(d => d))];
  
  // ビジュアルフローのレイアウト設定
  const startRow = 3;
  const startCol = 2;
  const boxWidth = 3;
  const boxHeight = 3;
  const horizontalGap = 1;
  const verticalGap = 1;
  
  // タイトル設定
  visualSheet.getRange(1, 1).setValue('業務フロー図');
  visualSheet.getRange(1, 1).setFontSize(16).setFontWeight('bold');
  
  // 部署別のレーン作成
  let currentCol = startCol;
  const deptColumns = {};
  
  departments.forEach((dept, index) => {
    const deptCol = currentCol + index * (boxWidth + horizontalGap);
    deptColumns[dept] = deptCol;
    
    // 部署名を表示
    visualSheet.getRange(startRow - 1, deptCol, 1, boxWidth).merge();
    visualSheet.getRange(startRow - 1, deptCol).setValue(dept);
    visualSheet.getRange(startRow - 1, deptCol).setBackground('#e8eaf6')
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setBorder(true, true, true, true, true, false);
  });
  
  // フローボックスの描画
  let currentRow = startRow + 1;
  const processedSteps = [];
  
  rows.forEach((row, rowIndex) => {
    const step = row[0]; // 工程
    const timing = row[1]; // 実施タイミング
    const dept = row[2]; // 部署
    const role = row[3]; // 担当役割
    const task = row[4]; // 作業内容
    const condition = row[5]; // 条件分岐
    const tool = row[6]; // 利用ツール
    const url = row[7]; // URLリンク
    const note = row[8]; // 備考
    
    const col = deptColumns[dept] || startCol;
    const currentRowPos = currentRow;
    
    // ボックスのセル範囲を取得
    const boxRange = visualSheet.getRange(currentRowPos, col, boxHeight, boxWidth);
    boxRange.merge();
    
    // ボックスの内容設定
    let boxContent = `【${step}】\n${task}`;
    if (role) boxContent += `\n(${role})`;
    if (tool) boxContent += `\n[${tool}]`;
    
    boxRange.setValue(boxContent);
    
    // ボックスのスタイル設定
    if (condition) {
      // 条件分岐は菱形風に黄色背景
      boxRange.setBackground('#fff9c4')
        .setBorder(true, true, true, true, false, false, '#ff9800', SpreadsheetApp.BorderStyle.SOLID_THICK);
    } else if (rowIndex === 0) {
      // 開始は緑背景
      boxRange.setBackground('#c8e6c9')
        .setBorder(true, true, true, true, false, false, '#4caf50', SpreadsheetApp.BorderStyle.SOLID_THICK);
    } else if (rowIndex === rows.length - 1) {
      // 終了は赤背景
      boxRange.setBackground('#ffcdd2')
        .setBorder(true, true, true, true, false, false, '#f44336', SpreadsheetApp.BorderStyle.SOLID_THICK);
    } else {
      // 通常処理は青背景
      boxRange.setBackground('#e3f2fd')
        .setBorder(true, true, true, true, false, false, '#2196f3', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    }
    
    boxRange.setWrap(true)
      .setVerticalAlignment('middle')
      .setHorizontalAlignment('center')
      .setFontSize(10);
    
    // URLリンクがある場合
    if (url) {
      const linkRange = visualSheet.getRange(row + boxHeight, col, 1, boxWidth);
      linkRange.merge();
      linkRange.setValue('📎 リンク');
      linkRange.setFormula(`=HYPERLINK("${url}", "📎 リンク")`);
      linkRange.setFontSize(9).setFontColor('#1a73e8');
    }
    
    // 備考がある場合
    if (note) {
      const noteRange = visualSheet.getRange(row, col + boxWidth + 1);
      noteRange.setValue(`💡 ${note}`);
      noteRange.setFontSize(9).setFontColor('#666').setWrap(true);
    }
    
    // 矢印の描画（次のステップがある場合）
    if (rowIndex < rows.length - 1) {
      const nextDept = rows[rowIndex + 1][2];
      const nextCol = deptColumns[nextDept] || startCol;
      
      if (col === nextCol) {
        // 同じ部署内での移動（下向き矢印）
        const arrowRange = visualSheet.getRange(row + boxHeight, col + Math.floor(boxWidth / 2));
        arrowRange.setValue('↓');
        arrowRange.setFontSize(16).setHorizontalAlignment('center');
      } else {
        // 異なる部署への移動（横向き矢印）
        const direction = nextCol > col ? '→' : '←';
        const arrowCol = col < nextCol ? col + boxWidth : col - 1;
        const arrowRange = visualSheet.getRange(row + Math.floor(boxHeight / 2), arrowCol);
        arrowRange.setValue(direction);
        arrowRange.setFontSize(16).setHorizontalAlignment('center');
      }
    }
    
    processedSteps.push({
      step: step,
      row: row,
      col: col,
      dept: dept,
      condition: condition
    });
    
    currentRow += boxHeight + verticalGap + 1;
  });
  
  // 凡例の追加
  const legendRow = currentRow + 3;
  visualSheet.getRange(legendRow, startCol).setValue('【凡例】');
  visualSheet.getRange(legendRow, startCol).setFontWeight('bold');
  
  const legends = [
    { color: '#c8e6c9', text: '開始', border: '#4caf50' },
    { color: '#e3f2fd', text: '通常処理', border: '#2196f3' },
    { color: '#fff9c4', text: '条件分岐', border: '#ff9800' },
    { color: '#ffcdd2', text: '終了', border: '#f44336' }
  ];
  
  legends.forEach((legend, index) => {
    const legendCol = startCol + index * 3;
    const legendRange = visualSheet.getRange(legendRow + 1, legendCol, 1, 2);
    legendRange.merge();
    legendRange.setValue(legend.text);
    legendRange.setBackground(legend.color)
      .setBorder(true, true, true, true, false, false, legend.border, SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
      .setHorizontalAlignment('center');
  });
  
  // 列幅と行高の調整
  for (let i = 1; i <= visualSheet.getMaxColumns(); i++) {
    visualSheet.setColumnWidth(i, 120);
  }
  
  for (let i = startRow; i <= currentRow; i++) {
    visualSheet.setRowHeight(i, 60);
  }
  
  // シート全体の書式設定
  visualSheet.getRange(1, 1, visualSheet.getMaxRows(), visualSheet.getMaxColumns())
    .setFontFamily('Noto Sans JP');
  
  logActivity('VISUAL_FLOW', 'Visual flow generated successfully');
  
  console.log('ビジュアルフローを生成しました。');
}

// サンプルデータ作成（開発・テスト用）
function createSampleFlowData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const flowSheet = ss.getSheetByName('フロー') || ss.insertSheet('フロー');
  
  // シートをクリア
  flowSheet.clear();
  
  // ヘッダー設定
  const headers = [
    '工程', '実施タイミング', '部署', '担当役割', '作業内容', 
    '条件分岐', '利用ツール', 'URLリンク', '備考'
  ];
  
  flowSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  flowSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#e8f5e9');
  
  // サンプルデータ
  const sampleData = [
    ['要件定義', 'Day 1-5', '企画部', 'プロジェクトマネージャー', '業務要件のヒアリングと整理', '', 'Teams, Miro', 'https://example.com/requirements', '関係者全員参加必須'],
    ['承認判断', 'Day 6', '経営企画部', '部長', '要件の承認可否を判断', '承認/差戻し', '', '', '予算上限確認'],
    ['基本設計', 'Day 7-15', 'IT部', 'システムアーキテクト', 'システム構成の設計', '', 'draw.io, Confluence', 'https://example.com/design', ''],
    ['詳細設計', 'Day 16-25', 'IT部', '開発リード', '機能仕様の詳細化', '', 'GitHub, Figma', '', 'UI/UXチームと連携'],
    ['開発', 'Day 26-50', '開発部', '開発チーム', 'コーディングと単体テスト', '', 'VS Code, Git', 'https://github.com/example', 'アジャイル開発'],
    ['品質チェック', 'Day 51-55', '品質管理部', 'QAエンジニア', 'テスト実施と不具合修正', '合格/再テスト', 'Selenium, JIRA', '', ''],
    ['リリース準備', 'Day 56-58', 'IT部', 'インフラチーム', '本番環境へのデプロイ準備', '', 'Jenkins, Docker', '', ''],
    ['本番リリース', 'Day 59', 'IT部', 'リリースマネージャー', '本番環境への展開', '', 'Kubernetes', '', '夜間作業'],
    ['運用引継ぎ', 'Day 60', '運用部', '運用チーム', '運用手順書の確認と引継ぎ', '', 'ServiceNow', 'https://example.com/operations', '24時間体制確立']
  ];
  
  flowSheet.getRange(2, 1, sampleData.length, headers.length).setValues(sampleData);
  
  // 列幅調整
  flowSheet.setColumnWidth(1, 100); // 工程
  flowSheet.setColumnWidth(2, 120); // 実施タイミング
  flowSheet.setColumnWidth(3, 100); // 部署
  flowSheet.setColumnWidth(4, 150); // 担当役割
  flowSheet.setColumnWidth(5, 250); // 作業内容
  flowSheet.setColumnWidth(6, 150); // 条件分岐
  flowSheet.setColumnWidth(7, 120); // 利用ツール
  flowSheet.setColumnWidth(8, 200); // URLリンク
  flowSheet.setColumnWidth(9, 200); // 備考
  
  console.log('サンプルフローデータを作成しました。');
}

// 高度なビジュアルフロー生成（業務フローチャート図作成.js参考版）
function generateAdvancedVisualFlow() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(FLOW_SHEET);
    if (!sheet) {
      console.error('フローシートが見つかりません。');
      return;
    }
    
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) {
      console.error('データがありません。');
      return;
    }
    
    // ビジュアルシートの準備
    const visualSheet = getOrCreateSheet(VISUAL_SHEET);
    visualSheet.clear();
    visualSheet.clearFormats();
    
    // データの解析
    const flowData = parseAdvancedFlowData(data);
    
    // フローチャートの描画
    drawAdvancedFlowChart(visualSheet, flowData);
    
    // 業務サマリーシート作成
    createBusinessSummarySheet(flowData);
    
    console.log('高度なビジュアルフローが生成されました。');
    
  } catch (error) {
    console.error('エラー:', error);
    console.error('フロー生成中にエラーが発生しました:', error);
  }
}

// 高度なフローデータ解析
function parseAdvancedFlowData(data) {
  const headers = data[0];
  const columnIndex = {};
  headers.forEach((header, index) => {
    columnIndex[header] = index;
  });
  
  const flowData = {
    departments: {},
    departmentList: [],
    timings: [],
    tools: new Set(),
    datasources: {},
    processName: "",
    statistics: {
      totalTasks: 0,
      totalDepartments: 0,
      totalTools: 0,
      decisionPoints: 0
    }
  };
  
  // プロセス名の取得
  if (data.length > 1 && data[1][columnIndex["工程"]]) {
    flowData.processName = data[1][columnIndex["工程"]];
  }
  
  // データの整理
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[columnIndex["工程"]] || row[columnIndex["工程"]] === "") continue;
    
    const dept = row[columnIndex["部署"]] || "その他";
    const timing = row[columnIndex["実施タイミング"]] || "";
    const tool = row[columnIndex["利用ツール"]] || "";
    const url = row[columnIndex["URLリンク"]] || "";
    const condition = row[columnIndex["条件分岐"]] || "";
    
    // 部署の初期化
    if (!flowData.departments[dept]) {
      flowData.departments[dept] = {};
      if (!flowData.departmentList.includes(dept)) {
        flowData.departmentList.push(dept);
      }
    }
    
    // タイミングの追加
    if (timing && !flowData.timings.includes(timing)) {
      flowData.timings.push(timing);
    }
    
    // タスクの追加
    if (!flowData.departments[dept][timing]) {
      flowData.departments[dept][timing] = [];
    }
    
    flowData.departments[dept][timing].push({
      task: row[columnIndex["作業内容"]] || "",
      role: row[columnIndex["担当役割"]] || "",
      condition: condition,
      tool: tool,
      url: url,
      note: row[columnIndex["備考"]] || ""
    });
    
    // 統計更新
    flowData.statistics.totalTasks++;
    if (condition && condition !== "-") {
      flowData.statistics.decisionPoints++;
    }
    
    // ツールの収集
    if (tool && tool !== "-") {
      const tools = tool.split(/[／、,]/);
      tools.forEach(t => {
        const trimmedTool = t.trim();
        if (trimmedTool) {
          flowData.tools.add(trimmedTool);
        }
      });
    }
    
    // データソース管理
    if (url && url !== "-") {
      if (!flowData.datasources[timing]) {
        flowData.datasources[timing] = [];
      }
      if (!flowData.datasources[timing].includes(url)) {
        flowData.datasources[timing].push(url);
      }
    }
  }
  
  flowData.statistics.totalDepartments = flowData.departmentList.length;
  flowData.statistics.totalTools = flowData.tools.size;
  
  return flowData;
}

// 高度なフローチャート描画
function drawAdvancedFlowChart(sheet, flowData) {
  let currentRow = 1;
  const maxCols = Math.max(flowData.departmentList.length + 2, 10);
  
  // タイトル行
  const flowTitle = flowData.processName || "業務フロー";
  sheet.getRange(currentRow, 1, 1, maxCols).merge();
  const titleCell = sheet.getRange(currentRow, 1);
  titleCell.setValue(flowTitle + "（" + new Date().getFullYear() + "年" + (new Date().getMonth() + 1) + "月版）");
  titleCell.setBackground(ADVANCED_COLORS.MAIN_HEADER);
  titleCell.setFontColor("#FFFFFF");
  titleCell.setFontSize(18);
  titleCell.setFontWeight("bold");
  titleCell.setHorizontalAlignment("center");
  titleCell.setVerticalAlignment("middle");
  sheet.setRowHeight(currentRow, 50);
  currentRow++;
  
  // 統計情報行
  sheet.getRange(currentRow, 1, 1, maxCols).merge();
  const statsCell = sheet.getRange(currentRow, 1);
  statsCell.setValue(`📊 総タスク: ${flowData.statistics.totalTasks} | 👥 部署: ${flowData.statistics.totalDepartments} | 🔧 ツール: ${flowData.statistics.totalTools} | ⚡ 判断: ${flowData.statistics.decisionPoints}`);
  statsCell.setBackground(ADVANCED_COLORS.SUB_HEADER);
  statsCell.setFontColor("#FFFFFF");
  statsCell.setHorizontalAlignment("center");
  sheet.setRowHeight(currentRow, 35);
  currentRow++;
  
  // ツール行
  if (flowData.tools.size > 0) {
    drawAdvancedToolsRow(sheet, currentRow, flowData.tools, maxCols);
    currentRow++;
  }
  
  // ヘッダー行
  drawAdvancedHeaderRow(sheet, currentRow, flowData.departmentList, Object.keys(flowData.datasources).length > 0);
  currentRow++;
  
  // 開始行
  drawAdvancedStartRow(sheet, currentRow, maxCols);
  currentRow++;
  
  // 各タイミングの行
  flowData.timings.forEach((timing, index) => {
    drawAdvancedTimingRow(sheet, currentRow, timing, flowData, index);
    currentRow++;
  });
  
  // 終了行
  drawAdvancedEndRow(sheet, currentRow, flowData.departmentList.length, Object.keys(flowData.datasources).length > 0);
  currentRow++;
  
  // 凡例行
  drawAdvancedLegendRow(sheet, currentRow, maxCols);
  
  // 列幅の調整
  adjustAdvancedColumnWidths(sheet, flowData.departmentList.length, Object.keys(flowData.datasources).length > 0);
  
  // 罫線の設定
  applyAdvancedBorders(sheet, currentRow, maxCols);
}

// ツール行の描画（高度版）
function drawAdvancedToolsRow(sheet, row, tools, maxCols) {
  sheet.getRange(row, 1).setValue("使用ツール");
  sheet.getRange(row, 1).setBackground(ADVANCED_COLORS.SECTION_HEADER);
  sheet.getRange(row, 1).setFontColor("#FFFFFF");
  sheet.getRange(row, 1).setFontWeight("bold");
  
  sheet.getRange(row, 2, 1, maxCols - 1).merge();
  const toolCell = sheet.getRange(row, 2);
  toolCell.setBackground(ADVANCED_COLORS.TOOL_BG);
  
  let toolText = "";
  tools.forEach(tool => {
    const toolInfo = TOOL_ICONS[tool] || { icon: '🔧', color: '#666666' };
    toolText += ` ${toolInfo.icon} ${tool} `;
  });
  
  toolCell.setValue(toolText);
  toolCell.setHorizontalAlignment("left");
  sheet.setRowHeight(row, 35);
}

// ヘッダー行の描画（高度版）
function drawAdvancedHeaderRow(sheet, row, departments, hasDataSource) {
  // 日程列
  sheet.getRange(row, 1).setValue("タイミング");
  sheet.getRange(row, 1).setBackground(ADVANCED_COLORS.SECTION_HEADER);
  sheet.getRange(row, 1).setFontColor("#FFFFFF");
  sheet.getRange(row, 1).setFontWeight("bold");
  
  // 部署列
  departments.forEach((dept, index) => {
    const col = index + 2;
    const cell = sheet.getRange(row, col);
    cell.setValue(dept);
    cell.setBackground(DEPT_COLOR_PALETTE[index % DEPT_COLOR_PALETTE.length]);
    cell.setFontWeight("bold");
    cell.setWrap(true);
    cell.setHorizontalAlignment("center");
    cell.setVerticalAlignment("middle");
  });
  
  // データソース列
  if (hasDataSource) {
    const dataCol = departments.length + 2;
    sheet.getRange(row, dataCol).setValue("📚 関連資料");
    sheet.getRange(row, dataCol).setBackground(ADVANCED_COLORS.DATASOURCE_BG);
    sheet.getRange(row, dataCol).setFontWeight("bold");
  }
  
  sheet.setRowHeight(row, 50);
}

// 開始行の描画（高度版）
function drawAdvancedStartRow(sheet, row, maxCols) {
  sheet.getRange(row, 1, 1, maxCols).merge();
  const cell = sheet.getRange(row, 1);
  cell.setValue("🚀 【プロセス開始】");
  cell.setBackground(ADVANCED_COLORS.START_END);
  cell.setFontColor("#FFFFFF");
  cell.setFontWeight("bold");
  cell.setFontSize(14);
  cell.setHorizontalAlignment("center");
  cell.setBorder(true, true, true, true, false, false, "#228B22", SpreadsheetApp.BorderStyle.SOLID_THICK);
  sheet.setRowHeight(row, 40);
}

// タイミング行の描画（高度版）
function drawAdvancedTimingRow(sheet, row, timing, flowData, timingIndex) {
  // タイミング列
  const timingCell = sheet.getRange(row, 1);
  timingCell.setValue(timing);
  timingCell.setBackground(ADVANCED_COLORS.TIMELINE_BG);
  timingCell.setFontWeight("bold");
  timingCell.setWrap(true);
  
  // 各部署のタスク
  flowData.departmentList.forEach((dept, deptIndex) => {
    const col = deptIndex + 2;
    const tasks = flowData.departments[dept][timing];
    
    if (tasks && tasks.length > 0) {
      const task = tasks[0];
      const cell = sheet.getRange(row, col);
      
      // タスク内容の設定
      let content = task.task;
      if (task.role && task.role !== "-") {
        content = "【" + task.role + "】\n" + content;
      }
      if (task.tool && task.tool !== "-") {
        const toolInfo = TOOL_ICONS[task.tool.split(/[／、,]/)[0].trim()];
        if (toolInfo) {
          content += "\n" + toolInfo.icon + " " + task.tool;
        } else {
          content += "\n🔧 " + task.tool;
        }
      }
      
      cell.setValue(content);
      cell.setWrap(true);
      cell.setHorizontalAlignment("center");
      cell.setVerticalAlignment("middle");
      
      // スタイルの設定
      if (task.condition && task.condition !== "-") {
        // 判断ボックス
        cell.setBackground(ADVANCED_COLORS.DECISION);
        cell.setBorder(true, true, true, true, false, false, "#FF8C00", SpreadsheetApp.BorderStyle.SOLID_THICK);
        cell.setFontWeight("bold");
        
        const noteContent = "⚡ 条件分岐: " + task.condition + 
                          (task.note ? "\n📝 備考: " + task.note : "");
        cell.setNote(noteContent);
      } else {
        // プロセスボックス
        cell.setBackground(ADVANCED_COLORS.PROCESS);
        cell.setBorder(true, true, true, true, false, false, "#4682B4", SpreadsheetApp.BorderStyle.SOLID_THICK);
        
        if (task.note) {
          cell.setNote("📝 備考: " + task.note);
        }
      }
      
      // 矢印の追加
      if (timingIndex < flowData.timings.length - 1) {
        const nextTiming = flowData.timings[timingIndex + 1];
        if (flowData.departments[dept][nextTiming]) {
          addAdvancedArrowToCell(cell, "↓");
        }
      }
    } else {
      // 空のセル
      const cell = sheet.getRange(row, col);
      cell.setBackground(ADVANCED_COLORS.EMPTY_BG);
    }
  });
  
  // データソース
  const hasDataSource = Object.keys(flowData.datasources).length > 0;
  if (hasDataSource) {
    const dataCol = flowData.departmentList.length + 2;
    const dataCell = sheet.getRange(row, dataCol);
    
    if (flowData.datasources[timing] && flowData.datasources[timing].length > 0) {
      const urls = flowData.datasources[timing].join("\n");
      dataCell.setValue("📎 " + urls);
      dataCell.setBackground(ADVANCED_COLORS.DATASOURCE_BG);
      dataCell.setBorder(true, true, true, true, false, false, "#2196F3", SpreadsheetApp.BorderStyle.DASHED);
      dataCell.setWrap(true);
      dataCell.setHorizontalAlignment("center");
      dataCell.setVerticalAlignment("middle");
    } else {
      dataCell.setValue("");
      dataCell.setBackground(ADVANCED_COLORS.EMPTY_BG);
    }
  }
  
  sheet.setRowHeight(row, 90);
}

// 終了行の描画（高度版）
function drawAdvancedEndRow(sheet, row, deptCount, hasDataSource) {
  const mergeCols = hasDataSource ? deptCount + 2 : deptCount + 1;
  sheet.getRange(row, 1, 1, mergeCols).merge();
  const cell = sheet.getRange(row, 1);
  cell.setValue("✅ 【プロセス完了】");
  cell.setBackground(ADVANCED_COLORS.START_END);
  cell.setFontColor("#FFFFFF");
  cell.setFontSize(14);
  cell.setFontWeight("bold");
  cell.setHorizontalAlignment("center");
  cell.setBorder(true, true, true, true, false, false, "#228B22", SpreadsheetApp.BorderStyle.SOLID_THICK);
  sheet.setRowHeight(row, 50);
}

// 凡例行の描画（高度版）
function drawAdvancedLegendRow(sheet, row, maxCols) {
  sheet.getRange(row, 1, 1, maxCols).merge();
  const legendCell = sheet.getRange(row, 1);
  legendCell.setValue("【凡例】 📦 処理・作業　⚡ 判断・分岐　→ 処理の流れ　📎 関連資料　📝 備考（セルの注記に詳細）");
  legendCell.setBackground(ADVANCED_COLORS.TIMELINE_BG);
  legendCell.setFontWeight("bold");
  legendCell.setHorizontalAlignment("left");
  legendCell.setVerticalAlignment("middle");
  sheet.setRowHeight(row, 40);
}

// セルに矢印を追加（高度版）
function addAdvancedArrowToCell(cell, arrow) {
  const currentValue = cell.getValue();
  const richText = SpreadsheetApp.newRichTextValue()
    .setText(currentValue + "\n" + arrow)
    .setTextStyle(currentValue.length + 1, currentValue.length + arrow.length + 1, 
      SpreadsheetApp.newTextStyle()
        .setForegroundColor(ADVANCED_COLORS.INFO)
        .setFontSize(16)
        .setBold(true)
        .build())
    .build();
  cell.setRichTextValue(richText);
}

// 列幅の調整（高度版）
function adjustAdvancedColumnWidths(sheet, deptCount, hasDataSource) {
  sheet.setColumnWidth(1, 150); // タイムライン列
  for (let i = 2; i <= deptCount + 1; i++) {
    sheet.setColumnWidth(i, 200); // 部署列
  }
  if (hasDataSource) {
    sheet.setColumnWidth(deptCount + 2, 150); // データソース列
  }
}

// 罫線の設定（高度版）
function applyAdvancedBorders(sheet, lastRow, lastCol) {
  const range = sheet.getRange(1, 1, lastRow, lastCol);
  range.setBorder(true, true, true, true, true, true, "#d0d0d0", SpreadsheetApp.BorderStyle.SOLID);
}

// 業務サマリーシート作成
function createBusinessSummarySheet(flowData) {
  const summarySheet = getOrCreateSheet('業務サマリー');
  summarySheet.clear();
  
  let row = 1;
  
  // タイトル
  summarySheet.getRange(row, 1, 1, 6).merge();
  const titleCell = summarySheet.getRange(row, 1);
  titleCell.setValue('業務プロセスサマリー');
  titleCell.setBackground(ADVANCED_COLORS.MAIN_HEADER);
  titleCell.setFontColor('#FFFFFF');
  titleCell.setFontSize(18);
  titleCell.setFontWeight('bold');
  titleCell.setHorizontalAlignment('center');
  summarySheet.setRowHeight(row, 50);
  row += 2;
  
  // 基本情報
  summarySheet.getRange(row, 1).setValue('プロセス名');
  summarySheet.getRange(row, 2, 1, 5).merge();
  summarySheet.getRange(row, 2).setValue(flowData.processName || '未設定');
  summarySheet.getRange(row, 1).setBackground(ADVANCED_COLORS.TIMELINE_BG);
  summarySheet.getRange(row, 1).setFontWeight('bold');
  row++;
  
  summarySheet.getRange(row, 1).setValue('総タスク数');
  summarySheet.getRange(row, 2, 1, 5).merge();
  summarySheet.getRange(row, 2).setValue(flowData.statistics.totalTasks);
  summarySheet.getRange(row, 1).setBackground(ADVANCED_COLORS.TIMELINE_BG);
  summarySheet.getRange(row, 1).setFontWeight('bold');
  row++;
  
  summarySheet.getRange(row, 1).setValue('関連部署数');
  summarySheet.getRange(row, 2, 1, 5).merge();
  summarySheet.getRange(row, 2).setValue(flowData.statistics.totalDepartments);
  summarySheet.getRange(row, 1).setBackground(ADVANCED_COLORS.TIMELINE_BG);
  summarySheet.getRange(row, 1).setFontWeight('bold');
  row++;
  
  summarySheet.getRange(row, 1).setValue('使用ツール数');
  summarySheet.getRange(row, 2, 1, 5).merge();
  summarySheet.getRange(row, 2).setValue(flowData.statistics.totalTools);
  summarySheet.getRange(row, 1).setBackground(ADVANCED_COLORS.TIMELINE_BG);
  summarySheet.getRange(row, 1).setFontWeight('bold');
  row++;
  
  summarySheet.getRange(row, 1).setValue('判断ポイント数');
  summarySheet.getRange(row, 2, 1, 5).merge();
  summarySheet.getRange(row, 2).setValue(flowData.statistics.decisionPoints);
  summarySheet.getRange(row, 1).setBackground(ADVANCED_COLORS.TIMELINE_BG);
  summarySheet.getRange(row, 1).setFontWeight('bold');
  row += 2;
  
  // 部署別タスク
  summarySheet.getRange(row, 1, 1, 6).merge();
  summarySheet.getRange(row, 1).setValue('👥 部署別タスク分布');
  summarySheet.getRange(row, 1).setBackground(ADVANCED_COLORS.SUB_HEADER);
  summarySheet.getRange(row, 1).setFontColor('#FFFFFF');
  summarySheet.getRange(row, 1).setFontWeight('bold');
  row++;
  
  flowData.departmentList.forEach(dept => {
    let taskCount = 0;
    Object.values(flowData.departments[dept]).forEach(timingTasks => {
      taskCount += timingTasks.length;
    });
    
    summarySheet.getRange(row, 1).setValue(dept);
    summarySheet.getRange(row, 2, 1, 5).merge();
    summarySheet.getRange(row, 2).setValue(`${taskCount} タスク`);
    summarySheet.getRange(row, 1).setBackground(DEPT_COLOR_PALETTE[flowData.departmentList.indexOf(dept) % DEPT_COLOR_PALETTE.length]);
    summarySheet.getRange(row, 1).setFontWeight('bold');
    row++;
  });
  
  // 書式調整
  summarySheet.autoResizeColumns(1, 6);
}

// シート取得または作成（単純版）
function getOrCreateSheetSimple(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }
  
  return sheet;
}

// 業務サマリーのみ作成
function createBusinessSummaryOnly() {
  try {
    const flowSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(FLOW_SHEET);
    if (!flowSheet) {
      console.error('フローシートが見つかりません');
      return;
    }
    
    const data = flowSheet.getDataRange().getValues();
    if (data.length < 2) {
      console.error('データがありません');
      return;
    }
    
    const flowData = parseAdvancedFlowData(data);
    createBusinessSummarySheet(flowData);
    
    console.log('業務サマリーを作成しました');
  } catch (error) {
    console.error('エラー:', error);
  }
}

// ================================================================================
// 3. gmail_inbound.gs - Gmail受信処理機能
// ================================================================================

// Gmail受信処理

// 新着メール処理（メイン関数）
function processNewEmails() {
  const query = getConfig('PROCESSING_QUERY') || 'subject:[task] is:unread';
  logActivity('PROCESS_START', `Processing emails with query: ${query}`);
  console.log('検索クエリ:', query);
  
  try {
    const threads = GmailApp.search(query);
    console.log('検索結果:', threads.length, '件のスレッドが見つかりました');
    
    if (threads.length === 0) {
      logActivity('PROCESS_INFO', 'No new emails found');
      
      // デバッグ用: 全ての未読メールを確認
      const allUnread = GmailApp.search('is:unread', 0, 5);
      console.log('未読メール総数:', allUnread.length);
      if (allUnread.length > 0) {
        console.log('未読メールの件名リスト:');
        allUnread.forEach(thread => {
          const firstMessage = thread.getMessages()[0];
          const subject = firstMessage.getSubject();
          console.log('  - "' + subject + '"');
          if (subject.toLowerCase().includes('task')) {
            console.log('    → このメールは"task"を含んでいます');
          }
        });
      }
      return;
    }
    
    threads.forEach(thread => {
      processThread(thread);
    });
    
    logActivity('PROCESS_END', `Processed ${threads.length} threads`);
  } catch (e) {
    logActivity('PROCESS_ERROR', e.toString());
    throw e;
  }
}

// スレッド処理
function processThread(thread) {
  const messages = thread.getMessages();
  
  messages.forEach(msg => {
    try {
      processMessage(msg, thread);
    } catch (e) {
      logActivity('MESSAGE_ERROR', `Failed to process message: ${e.toString()}`);
    }
  });
}

// メッセージ処理
function processMessage(msg, thread) {
  const messageId = msg.getId();
  
  // 処理済みチェック
  if (isProcessed(messageId)) {
    logActivity('SKIP', `Message ${messageId} already processed`);
    return;
  }
  
  // メール情報抽出
  const from = extractEmail(msg.getFrom());
  const subject = msg.getSubject();
  const htmlBody = msg.getBody();
  let plainBody = msg.getPlainBody() || htmlToText(htmlBody);
  const receivedDate = msg.getDate();
  
  // 署名部分を除去
  plainBody = removeEmailSignature(plainBody);
  
  // 件名から[task]を除去して、本文と結合
  const cleanSubject = subject.replace(/\[task\]/gi, '').trim();
  const combinedContent = `【件名】${cleanSubject}\n\n【本文】\n${plainBody}`;
  
  // Inboxにログ記録
  logInbox(messageId, thread.getId(), from, subject, plainBody.substring(0, 200), 'NEW');
  
  try {
    // OpenAI呼び出し（件名と本文を結合したものを送信）
    const orgProfile = getConfig('ORG_PROFILE_JSON') || '{}';
    const result = callOpenAI(combinedContent, orgProfile);
    
    // 検証
    validateOpenAIResponse(result);
    
    // ガバナンスチェックを自動実行
    console.log('=== ガバナンスチェック開始 ===');
    const governanceCheck = performComprehensiveGovernanceCheck(result.work_spec, result.flow_rows);
    console.log('ガバナンススコア:', governanceCheck.overallScore);
    console.log('開示要件:', governanceCheck.disclosureRequirements.length, '件');
    console.log('要専門家相談:', governanceCheck.advisorConsultations ? governanceCheck.advisorConsultations.length : 0, '件');
    
    // 新規スプレッドシートを作成（メールごとに独立）
    const newSpreadsheet = createIndependentSpreadsheetWithGovernance(cleanSubject, result, governanceCheck);
    
    // 共有設定（URLを知っている人は誰でも編集可能）
    setPublicEditAccess(newSpreadsheet);
    
    // 返信メール送信（新規スプレッドシートのURLを送信）
    sendNotificationEmail(from, result.work_spec, newSpreadsheet.getUrl());
    
    // 処理済みマーク
    markProcessed(messageId);
    labelThreadProcessed(thread);
    
    logActivity('PROCESS_SUCCESS', `Successfully processed message ${messageId}`);
  } catch (e) {
    logError(messageId, e);
    
    // エラーの詳細情報を取得
    let errorDetails = '';
    if (e.message) {
      errorDetails = e.message;
    } else if (e.toString) {
      errorDetails = e.toString();
    } else {
      errorDetails = String(e);
    }
    
    // スタックトレースがある場合は追加
    if (e.stack) {
      console.error('エラースタックトレース:', e.stack);
    }
    
    // エラー通知メール送信（詳細情報付き）
    sendErrorNotificationEmail(from, subject, errorDetails);
    
    // 元のエラーを再スロー（ただしint変換エラーは特別処理）
    if (errorDetails.includes('Cannot convert') && errorDetails.includes('to int')) {
      console.error('int変換エラーを検出。データ形式の問題の可能性があります。');
      // int変換エラーの場合は処理を続行しない
      return;
    }
    
    throw e;
  }
}

// メールごとに独立したスプレッドシートを作成
function createIndependentSpreadsheet(subject, result) {
  // 後方互換性のため、ガバナンスチェックなしで作成
  return createIndependentSpreadsheetWithGovernance(subject, result, null);
}

// ガバナンスチェック結果を含むスプレッドシートを作成
function createIndependentSpreadsheetWithGovernance(subject, result, governanceCheck) {
  console.log('=== 新規スプレッドシート作成開始（ガバナンス機能付き） ===');
  
  // スプレッドシート名を生成（日時を含む）
  const now = new Date();
  const dateStr = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyyMMdd_HHmmss');
  const spreadsheetName = `【業務記述書】${subject}_${dateStr}`;
  
  // 新規スプレッドシートを作成
  const newSpreadsheet = SpreadsheetApp.create(spreadsheetName);
  console.log('新規スプレッドシート作成:', newSpreadsheet.getUrl());
  
  // 1. デフォルトのシートを取得して業務サマリシートを作成
  const defaultSheet = newSpreadsheet.getSheets()[0];
  defaultSheet.setName('業務サマリ');
  console.log('業務サマリシート作成開始');
  createSummarySheetWithGovernance(defaultSheet, result.work_spec, governanceCheck);
  console.log('業務サマリシート作成完了');
  
  // スプレッドシートを明示的に保存
  SpreadsheetApp.flush();
  
  // 2. 業務記述書シートを作成
  console.log('業務記述書シート作成開始');
  const specSheet = newSpreadsheet.insertSheet('業務記述書');
  writeWorkSpecToSheet(specSheet, result.work_spec);
  console.log('業務記述書シート作成完了');
  
  // スプレッドシートを明示的に保存
  SpreadsheetApp.flush();
  
  // 3. フローシートを作成（ガバナンス情報付き）
  console.log('フローシート作成開始');
  const flowSheet = newSpreadsheet.insertSheet('フロー');
  const cleanedFlowRows = cleanFlowRowsData(result.flow_rows);
  writeFlowToSheetWithGovernance(flowSheet, cleanedFlowRows, governanceCheck);
  console.log('フローシート作成完了');
  
  // スプレッドシートを明示的に保存
  SpreadsheetApp.flush();
  
  // 4. ガバナンスレポートシートを作成
  if (governanceCheck) {
    console.log('ガバナンスレポート作成開始');
    try {
      const govSheet = newSpreadsheet.insertSheet('ガバナンス・コンプライアンス');
      createGovernanceReportSheet(govSheet, governanceCheck);
      console.log('ガバナンスレポート作成完了');
    } catch (error) {
      console.error('ガバナンスレポート作成エラー:', error);
    }
    SpreadsheetApp.flush();
  }
  
  // 5. 外部専門家相談チェックリストシートを作成
  if (governanceCheck && governanceCheck.advisorConsultations && governanceCheck.advisorConsultations.length > 0) {
    console.log('専門家相談チェックリスト作成開始');
    try {
      const consultSheet = newSpreadsheet.insertSheet('専門家相談チェックリスト');
      createConsultationChecklistSheet(consultSheet, governanceCheck.advisorConsultations);
      console.log('専門家相談チェックリスト作成完了');
    } catch (error) {
      console.error('専門家相談チェックリスト作成エラー:', error);
    }
    SpreadsheetApp.flush();
  }
  
  // 6. ビジュアルフローシートを作成（最後に作成して確実に実行）
  console.log('ビジュアルフローシート作成開始');
  try {
    const visualSheet = newSpreadsheet.insertSheet('ビジュアルフロー');
    createVisualFlowInSheet(visualSheet, flowSheet);
    console.log('ビジュアルフローシート作成完了');
  } catch (error) {
    console.error('ビジュアルフロー作成エラー:', error);
    // エラーが発生しても処理を継続
  }
  
  // 最終保存
  SpreadsheetApp.flush();
  
  console.log('=== 新規スプレッドシート作成完了 ===');
  return newSpreadsheet;
}

// 業務サマリシートを作成
function createSummarySheet(sheet, workSpec) {
  // 後方互換性のため、ガバナンスチェックなしで作成
  createSummarySheetWithGovernance(sheet, workSpec, null);
}

// ガバナンス情報を含む業務サマリシートを作成
function createSummarySheetWithGovernance(sheet, workSpec, governanceCheck) {
  // タイトル
  sheet.getRange('A1').setValue('業務サマリ');
  sheet.getRange('A1').setFontSize(18).setFontWeight('bold').setBackground('#4285F4').setFontColor('#FFFFFF');
  sheet.getRange('A1:D1').merge();
  
  // 基本情報
  const summaryData = [
    ['項目', '内容'],
    ['タイトル', workSpec.title || ''],
    ['概要', workSpec.summary || ''],
    ['目的', workSpec.purpose || ''],
    ['対象範囲', workSpec.scope || ''],
    ['前提条件', formatArray(workSpec.prerequisites) || ''],
    ['成果物', formatArray(workSpec.deliverables) || ''],
    ['関係者', formatArray(workSpec.stakeholders) || ''],
    ['作成日時', new Date()],
    ['最終更新', new Date()]
  ];
  
  sheet.getRange(3, 1, summaryData.length, 2).setValues(summaryData);
  sheet.getRange(3, 1, summaryData.length, 1).setFontWeight('bold').setBackground('#F0F0F0');
  sheet.getRange(3, 2, summaryData.length, 1).setWrap(true);
  
  // 列幅調整
  sheet.setColumnWidth(1, 150);
  sheet.setColumnWidth(2, 500);
  
  // 罫線
  sheet.getRange(3, 1, summaryData.length, 2).setBorder(true, true, true, true, true, true);
}

// 業務記述書をシートに書き込み
function writeWorkSpecToSheet(sheet, workSpec) {
  // ヘッダー
  const headers = ['項目', '内容', '詳細', '備考'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#E8F5E9');
  
  // データ
  const specData = [
    ['タイトル', workSpec.title || '', '', ''],
    ['概要', workSpec.summary || '', '', ''],
    ['目的', workSpec.purpose || '', '', ''],
    ['対象範囲', workSpec.scope || '', '', ''],
    ['前提条件', formatArray(workSpec.prerequisites) || '', '', ''],
    ['必要なリソース', formatArray(workSpec.resources) || '', '', ''],
    ['成果物', formatArray(workSpec.deliverables) || '', '', ''],
    ['関係者', formatArray(workSpec.stakeholders) || '', '', ''],
    ['承認プロセス', workSpec.approval_process || '', '', ''],
    ['リスクと対策', formatRisks(workSpec.risks) || '', '', ''],
    ['期限・頻度', formatTimeline(workSpec.timeline) || '', '', ''],
    ['KPI/成功基準', formatArray(workSpec.kpis) || '', '', '']
  ];
  
  sheet.getRange(2, 1, specData.length, headers.length).setValues(specData);
  sheet.getRange(2, 1, specData.length, 4).setWrap(true);
  
  // 列幅調整
  sheet.setColumnWidth(1, 150);
  sheet.setColumnWidth(2, 400);
  sheet.setColumnWidth(3, 300);
  sheet.setColumnWidth(4, 200);
  
  // 罫線
  sheet.getRange(1, 1, specData.length + 1, headers.length).setBorder(true, true, true, true, true, true);
}

// フローをシートに書き込み
function writeFlowToSheet(sheet, flowRows) {
  // 後方互換性のため、ガバナンスチェックなしで書き込み
  writeFlowToSheetWithGovernance(sheet, flowRows, null);
}

// ガバナンス情報を含むフローをシートに書き込み
function writeFlowToSheetWithGovernance(sheet, flowRows, governanceCheck) {
  const headers = FLOW_HEADERS; // 定数を使用して一貫性を保つ
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#E8F5E9');
  sheet.setFrozenRows(1);
  
  // フローデータを処理して書き込み
  const processedData = [];
  
  if (Array.isArray(flowRows)) {
    for (let i = 0; i < flowRows.length; i++) {
      const row = flowRows[i];
      
      if (typeof row === 'object' && row !== null) {
        const workContent = row['作業内容'] || '';
        const actions = splitIntoActions(workContent);
        const processName = row['工程'] || '';
        const timing = row['実施タイミング'] || '';
        const dept = row['部署'] || '';
        const condition = row['条件分岐'] || '';
        
        for (let j = 0; j < actions.length; j++) {
          const rowArray = [];
          
          for (const header of headers) {
            let value = '';
            
            if (header === '作業内容') {
              value = actions[j];
            } else if (header === '法令・規制') {
              value = checkLegalRegulations(processName, actions[j], timing, dept);
            } else if (header === '内部統制') {
              value = checkInternalControl(processName, actions[j], condition, dept);
            } else if (header === 'コンプライアンス留意点') {
              value = j === 0 ? generateComplianceNotes(processName, actions[j], timing, dept, condition) : '';
            } else if (j === 0) {
              value = row[header] || '';
            } else {
              if (header === '工程' || header === '実施タイミング' || header === '部署' || header === '担当役割') {
                value = row[header] || '';
              } else {
                value = '';
              }
            }
            
            rowArray.push(value);
          }
          processedData.push(rowArray);
        }
      }
    }
  }
  
  if (processedData.length > 0) {
    sheet.getRange(2, 1, processedData.length, headers.length).setValues(processedData);
    sheet.getRange(2, 1, processedData.length, headers.length).setWrap(true);
  }
  
  // 列幅調整
  sheet.setColumnWidth(1, 120); // 工程
  sheet.setColumnWidth(2, 150); // 実施タイミング
  sheet.setColumnWidth(3, 120); // 部署
  sheet.setColumnWidth(4, 150); // 担当役割
  sheet.setColumnWidth(5, 300); // 作業内容
  sheet.setColumnWidth(6, 150); // 条件分岐
  sheet.setColumnWidth(7, 150); // 利用ツール
  sheet.setColumnWidth(8, 200); // URLリンク
  sheet.setColumnWidth(9, 200); // 備考
  sheet.setColumnWidth(10, 250); // 法令・規制
  sheet.setColumnWidth(11, 250); // 内部統制
  sheet.setColumnWidth(12, 300); // コンプライアンス留意点
  
  // 罫線
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(1, 1, lastRow, headers.length).setBorder(true, true, true, true, true, true);
  }
}

// ビジュアルフローをシートに作成
function createVisualFlowInSheet(visualSheet, flowSheet) {
  const data = flowSheet.getDataRange().getValues();
  if (data.length < 2) {
    console.log('フローデータがないためビジュアルフロー作成をスキップ');
    return;
  }
  
  const flowData = parseFlowDataForVisual(data);
  drawVisualFlowChart(visualSheet, flowData);
}

// URLを知っている人は誰でも編集可能な共有設定
function setPublicEditAccess(spreadsheet) {
  try {
    const file = DriveApp.getFileById(spreadsheet.getId());
    
    // リンクを知っている全員が編集可能に設定
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.EDIT);
    
    console.log('共有設定完了: URLを知っている人は誰でも編集可能');
    return true;
  } catch (error) {
    console.error('共有設定エラー:', error);
    return false;
  }
}

// 共有設定処理（旧関数、互換性のため残す）
function handleSharing(senderEmail) {
  let shareSuccess = false;
  
  // ANYONE_WITH_LINKの設定を試行
  if (String(getConfig('SHARE_ANYONE_WITH_LINK')).toUpperCase() === 'TRUE') {
    shareSuccess = shareSheetAnyWithLink();
  }
  
  // 送信者を編集者として追加
  const editorSuccess = addEditor(senderEmail);
  
  return shareSuccess && editorSuccess;
}

// 成功通知メール送信
function sendNotificationEmail(to, workSpec, sheetUrl) {
  const subject = `[業務記述書完成] ${workSpec.title}`;
  const plainBody = buildPlainTextNotification(workSpec, sheetUrl);
  const htmlBody = buildHtmlNotification(workSpec, sheetUrl);
  
  GmailApp.sendEmail(to, subject, plainBody, {
    htmlBody: htmlBody,
    name: 'タスク管理システム'
  });
  
  logActivity('EMAIL_SENT', `Notification sent to ${to}`);
}

// プレーンテキスト通知作成
function buildPlainTextNotification(workSpec, sheetUrl) {
  return `業務記述書の作成が完了しました。

タイトル: ${workSpec.title}
概要: ${workSpec.summary}

スプレッドシートURL: ${sheetUrl}

このスプレッドシートでは以下の内容を確認・編集できます：
- 業務記述書（詳細仕様）
- タスクフロー表
- ビジュアルフロー図

【重要な注意事項】
- 本書面は自動生成されたものです。最終的な判断は専門家にご確認ください。
- 法令・規制に関する記載は参考情報であり、法的助言ではありません。
- スプレッドシートは編集可能です。必要に応じて内容を更新してください。

---
タスク管理システム by Google Apps Script`;
}

// HTML通知作成（UTF-8エンコーディング対応）
function buildHtmlNotification(workSpec, sheetUrl) {
  return `
    <!DOCTYPE html>
    <html>
    <head>
      <meta charset="UTF-8">
      <meta name="viewport" content="width=device-width, initial-scale=1.0">
    </head>
    <body>
    <div style="font-family: 'Noto Sans JP', Arial, sans-serif; max-width: 600px; margin: 0 auto;">
      <div style="background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; padding: 20px; text-align: center; border-radius: 10px 10px 0 0;">
        <h1 style="margin: 0; font-size: 24px;">業務記述書が完成しました</h1>
      </div>
      
      <div style="padding: 20px; background-color: #f8f9fa; border: 1px solid #e9ecef; border-top: none;">
        <h2 style="color: #495057; margin-top: 0;">${workSpec.title}</h2>
        <p style="font-size: 16px; line-height: 1.5; color: #6c757d;">${workSpec.summary}</p>
        
        <div style="background-color: white; padding: 15px; border-radius: 8px; margin: 20px 0; border-left: 4px solid #81C784;">
          <h3 style="margin-top: 0; color: #66BB6A;">スプレッドシート（編集可能）</h3>
          <p style="margin-bottom: 10px;">以下のリンクから業務記述書とタスク管理シートにアクセス・編集できます：</p>
          <a href="${sheetUrl}" style="display: inline-block; background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; padding: 12px 24px; text-decoration: none; border-radius: 4px; font-weight: bold;">スプレッドシートを開く</a>
        </div>
        
        ${workSpec.timeline && workSpec.timeline.length > 0 ? `
        <div style="background-color: white; padding: 15px; border-radius: 8px; margin: 20px 0; border-left: 4px solid #FFD54F;">
          <h3 style="margin-top: 0; color: #FFB74D;">主要マイルストーン</h3>
          <ul style="margin: 0; padding-left: 20px;">
            ${workSpec.timeline.map(phase => `
              <li style="margin-bottom: 8px;">
                <strong>${phase.phase}</strong> (${phase.duration_hint})
                ${phase.milestones && phase.milestones.length > 0 ? 
                  `<ul style="margin-top: 5px;">${phase.milestones.map(milestone => 
                    `<li style="color: #6c757d;">${milestone}</li>`
                  ).join('')}</ul>` 
                  : ''}
              </li>
            `).join('')}
          </ul>
        </div>
        ` : ''}
        
        <div style="background-color: white; padding: 15px; border-radius: 8px; margin: 20px 0; border-left: 4px solid #EF9A9A;">
          <h3 style="margin-top: 0; color: #E57373;">重要な注意事項</h3>
          <ul style="margin: 0; padding-left: 20px; color: #6c757d;">
            <li>本書面は自動生成されたものです。最終的な判断は専門家にご確認ください。</li>
            <li>法令・規制に関する記載は参考情報であり、法的助言ではありません。</li>
            <li>スプレッドシートは編集可能です。必要に応じて内容を更新してください。</li>
          </ul>
        </div>
        
        <div style="text-align: center; margin-top: 30px; padding-top: 20px; border-top: 1px solid #dee2e6;">
          <p style="color: #6c757d; font-size: 14px; margin: 0;">
            このメールは自動送信されています。<br>
            タスク管理システム by Google Apps Script
          </p>
        </div>
      </div>
    </div>
    </body>
    </html>
  `;
}

// エラー通知メール送信
function sendErrorNotificationEmail(to, originalSubject, errorMessage) {
  const subject = `[処理エラー] ${originalSubject}`;
  const body = `業務記述書の作成中にエラーが発生しました。

元の件名: ${originalSubject}
エラー内容: ${errorMessage}

お手数ですが、システム管理者にお問い合わせください。

---
タスク管理システム by Google Apps Script`;
  
  try {
    GmailApp.sendEmail(to, subject, body);
    logActivity('ERROR_EMAIL_SENT', `Error notification sent to ${to}`);
  } catch (e) {
    logActivity('ERROR_EMAIL_FAILED', `Failed to send error notification: ${e.toString()}`);
  }
}

// スレッドに処理済みラベルを付与
function labelThreadProcessed(thread) {
  try {
    // 既存のラベルを取得または作成
    let label = GmailApp.getUserLabelByName('PROCESSED');
    if (!label) {
      label = GmailApp.createLabel('PROCESSED');
    }
    
    thread.addLabel(label);
    thread.markRead();
    
    logActivity('LABEL', `Added PROCESSED label to thread ${thread.getId()}`);
  } catch (e) {
    logActivity('LABEL_ERROR', `Failed to label thread: ${e.toString()}`);
  }
}

// ================================================================================
// 4. gmail_outbound.gs - Gmail送信処理機能
// ================================================================================

// Gmail送信処理（任意の業務メール送信機能）

// サイドバーUI表示
function showEmailComposer() {
  const html = HtmlService.createHtmlOutput(getEmailComposerHtml())
    .setTitle('業務メール作成')
    .setWidth(400);
  SpreadsheetApp.getUi().showSidebar(html);
}

// メール送信（サイドバーから呼び出し）
function sendBusinessEmail(to, subject, body) {
  try {
    // 入力検証
    if (!to || !subject || !body) {
      throw new Error('宛先、件名、本文はすべて必須です。');
    }
    
    // メールアドレス検証
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    if (!emailRegex.test(to)) {
      throw new Error('有効なメールアドレスを入力してください。');
    }
    
    // デフォルトの件名プレフィックスを追加
    const prefixedSubject = subject.startsWith('[task]') ? subject : `[task] ${subject}`;
    
    // HTML形式のメール本文作成
    const htmlBody = createBusinessEmailHtml(body);
    
    // メール送信
    GmailApp.sendEmail(to, prefixedSubject, body, {
      htmlBody: htmlBody,
      name: 'タスク管理システム'
    });
    
    // ログ記録
    logActivity('OUTBOUND_EMAIL', `Sent to: ${to}, Subject: ${prefixedSubject}`);
    
    return {
      success: true,
      message: 'メールを送信しました。'
    };
    
  } catch (e) {
    logActivity('OUTBOUND_ERROR', e.toString());
    return {
      success: false,
      message: `エラー: ${e.toString()}`
    };
  }
}

// ビジネスメールHTML作成
function createBusinessEmailHtml(body) {
  // 改行をHTMLのbrタグに変換
  const htmlBody = body.replace(/\n/g, '<br>');
  
  return `
    <div style="font-family: 'Noto Sans JP', Arial, sans-serif; max-width: 600px; margin: 0 auto;">
      <div style="background-color: #f8f9fa; padding: 20px; border: 1px solid #dee2e6; border-radius: 8px;">
        <div style="background-color: white; padding: 20px; border-radius: 4px;">
          <p style="font-size: 16px; line-height: 1.6; color: #333; margin: 0;">
            ${htmlBody}
          </p>
        </div>
        
        <div style="margin-top: 20px; padding-top: 20px; border-top: 1px solid #dee2e6;">
          <p style="color: #6c757d; font-size: 14px; margin: 0;">
            このメールはタスク管理システムから送信されています。<br>
            業務内容に基づいて自動的に業務記述書とタスクフローが生成されます。
          </p>
        </div>
      </div>
    </div>
  `;
}

// サイドバーHTML取得
function getEmailComposerHtml() {
  const defaultTo = getConfig('DEFAULT_TO_EMAIL') || '';
  
  return `
    <!DOCTYPE html>
    <html>
      <head>
        <base target="_top">
        <style>
          body {
            font-family: 'Noto Sans JP', Arial, sans-serif;
            padding: 15px;
            margin: 0;
          }
          .form-group {
            margin-bottom: 15px;
          }
          label {
            display: block;
            margin-bottom: 5px;
            font-weight: bold;
            color: #333;
          }
          input, textarea {
            width: 100%;
            padding: 8px;
            border: 1px solid #ddd;
            border-radius: 4px;
            font-size: 14px;
            box-sizing: border-box;
          }
          textarea {
            resize: vertical;
            min-height: 150px;
          }
          button {
            width: 100%;
            padding: 10px;
            border: none;
            border-radius: 4px;
            font-size: 16px;
            font-weight: bold;
            cursor: pointer;
            transition: background-color 0.3s;
          }
          .btn-primary {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            margin-bottom: 10px;
          }
          .btn-primary:hover {
            opacity: 0.9;
          }
          .btn-secondary {
            background-color: #90A4AE;
            color: white;
          }
          .btn-secondary:hover {
            background-color: #78909C;
          }
          .loading {
            display: none;
            text-align: center;
            padding: 20px;
          }
          .spinner {
            border: 3px solid #f3f3f3;
            border-top: 3px solid #667eea;
            border-radius: 50%;
            width: 40px;
            height: 40px;
            animation: spin 1s linear infinite;
            margin: 0 auto;
          }
          @keyframes spin {
            0% { transform: rotate(0deg); }
            100% { transform: rotate(360deg); }
          }
          .message {
            padding: 10px;
            border-radius: 4px;
            margin-bottom: 15px;
            display: none;
          }
          .message.success {
            background-color: #E8F5E9;
            color: #155724;
            border: 1px solid #c3e6cb;
          }
          .message.error {
            background-color: #FFEBEE;
            color: #721c24;
            border: 1px solid #f5c6cb;
          }
          .info {
            background-color: #e3f2fd;
            border: 1px solid #90caf9;
            border-radius: 4px;
            padding: 10px;
            margin-bottom: 15px;
            color: #1565c0;
            font-size: 13px;
          }
        </style>
      </head>
      <body>
        <h2 style="color: #333; margin-top: 0;">業務メール作成</h2>
        
        <div class="info">
          ℹ️ このフォームから送信されたメールは自動的に処理され、業務記述書とタスクフローが生成されます。
        </div>
        
        <div id="message" class="message"></div>
        
        <form id="emailForm">
          <div class="form-group">
            <label for="to">宛先メールアドレス *</label>
            <input type="email" id="to" name="to" value="${defaultTo}" required placeholder="example@example.com">
          </div>
          
          <div class="form-group">
            <label for="subject">件名 *</label>
            <input type="text" id="subject" name="subject" required placeholder="業務依頼のタイトル">
            <small style="color: #666; font-size: 12px;">※ [task] プレフィックスは自動付与されます</small>
          </div>
          
          <div class="form-group">
            <label for="body">業務内容 *</label>
            <textarea id="body" name="body" required placeholder="実施したい業務の詳細を記入してください。&#10;&#10;例：&#10;- 業務の目的&#10;- 必要な成果物&#10;- 期限&#10;- 関係者&#10;- その他要件"></textarea>
          </div>
          
          <button type="submit" class="btn-primary">送信</button>
          <button type="button" class="btn-secondary" onclick="clearForm()">クリア</button>
        </form>
        
        <div id="loading" class="loading">
          <div class="spinner"></div>
          <p>送信中...</p>
        </div>
        
        <script>
          document.getElementById('emailForm').addEventListener('submit', function(e) {
            e.preventDefault();
            
            const to = document.getElementById('to').value;
            const subject = document.getElementById('subject').value;
            const body = document.getElementById('body').value;
            
            // フォームを非表示、ローディング表示
            document.getElementById('emailForm').style.display = 'none';
            document.getElementById('loading').style.display = 'block';
            document.getElementById('message').style.display = 'none';
            
            // GASの関数を呼び出し
            google.script.run
              .withSuccessHandler(function(result) {
                document.getElementById('loading').style.display = 'none';
                document.getElementById('emailForm').style.display = 'block';
                
                const messageDiv = document.getElementById('message');
                messageDiv.className = result.success ? 'message success' : 'message error';
                messageDiv.textContent = result.message;
                messageDiv.style.display = 'block';
                
                if (result.success) {
                  // 成功時はフォームをクリア
                  clearForm();
                  
                  // 3秒後にメッセージを非表示
                  setTimeout(function() {
                    messageDiv.style.display = 'none';
                  }, 3000);
                }
              })
              .withFailureHandler(function(error) {
                document.getElementById('loading').style.display = 'none';
                document.getElementById('emailForm').style.display = 'block';
                
                const messageDiv = document.getElementById('message');
                messageDiv.className = 'message error';
                messageDiv.textContent = 'エラーが発生しました: ' + error.toString();
                messageDiv.style.display = 'block';
              })
              .sendBusinessEmail(to, subject, body);
          });
          
          function clearForm() {
            document.getElementById('subject').value = '';
            document.getElementById('body').value = '';
            // 宛先はデフォルト値があれば保持
          }
        </script>
      </body>
    </html>
  `;
}

// ================================================================================
// 5. menu.gs - カスタムメニューとセットアップ機能
// ================================================================================

// カスタムメニューとセットアップ機能

// スプレッドシート開いた時の処理（重複のため削除）
// この関数は既に29行目で定義されています
/*
function onOpenDuplicate() {
  // この関数は重複のためコメントアウト
  // メイン機能は29行目のonOpen()関数を使用してください
}
*/

// 初回起動チェック用の関数
function checkFirstRun() {
  // 初回起動チェック処理をここに記載
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
      console.log('✅ APIキーを設定しました。');
      logActivity('API_KEY', 'API key configured');
    } else {
      console.warn('⚠️ APIキーが入力されていません。');
    }
  }
}

// モデル選択ダイアログ
function selectOpenAIModel() {
  const ui = SpreadsheetApp.getUi();
  const currentModel = getConfig('OPENAI_MODEL') || 'gpt-5';
  
  const htmlContent = `
    <div style="font-family: Arial, sans-serif; padding: 20px;">
      <h3>OpenAIモデル選択</h3>
      <p>現在のモデル: <strong>${currentModel}</strong></p>
      <br>
      <label style="background-color: #e8f5e9; padding: 5px; border-radius: 5px; display:block; margin-bottom:8px;">
        <input type="radio" name="model" value="gpt-5" ${currentModel === 'gpt-5' ? 'checked' : ''}>
        <strong>GPT-5</strong> 🌟 (推奨) - 高精度推論、Responses API最適化、構造化出力に強い
      </label>
      <label style="background-color: #eef3fb; padding: 5px; border-radius: 5px; display:block;">
        <input type="radio" name="model" value="gpt-4o" ${currentModel === 'gpt-4o' ? 'checked' : ''}>
        <strong>GPT-4o</strong> 🚀 - 高度な推論、深層分析
      </label><br>
      <hr>
      <p style="color: #666; font-size: 12px;">
        ※ GPT-4.1シリーズは2024年6月までの知識を持ち、1Mトークンのコンテキストに対応
      </p>
      <br>
      <button onclick="google.script.run.updateOpenAIModel(document.querySelector('input[name=model]:checked').value); google.script.host.close();">
        保存
      </button>
      <button onclick="google.script.host.close();">キャンセル</button>
    </div>
  `;
  
  const html = HtmlService.createHtmlOutput(htmlContent)
    .setWidth(450)
    .setHeight(420);
  
  ui.showModalDialog(html, 'AIモデル選択');
}

// モデル更新
function updateOpenAIModel(model) {
  setConfig('OPENAI_MODEL', model);
  const ui = SpreadsheetApp.getUi();
  
  // モデル別のメッセージ
  let message = `AIモデルを ${model} に変更しました。`;
  if (model === 'gpt-5') {
    message += '\n\n🌟 GPT-5を使用します。Responses API最適化で構造化出力が安定します。';
  } else if (model === 'gpt-4o') {
    message += '\n\n🚀 GPT-4oを使用します。高度な推論と深層分析が可能です。';
  }
  
  ui.alert('設定完了', message, ui.ButtonSet.OK);
  logActivity('MODEL_CHANGED', `OpenAI model changed to: ${model}`);
}

// GPT-4への自動アップグレード（既存ユーザー向け）
function upgradeToGPT4() {
  const ui = SpreadsheetApp.getUi();
  setConfig('OPENAI_MODEL', 'gpt-5');
  ui.alert('アップグレード完了', 'GPT-5モデルに切り替えました。Responses APIに最適化されています。', ui.ButtonSet.OK);
  logActivity('MODEL_UPGRADE', 'Upgraded to GPT-5');
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
  try {
    console.log('📥 手動メール処理開始');
    processNewEmails();
    console.log('✅ メール処理完了');
  } catch (e) {
    console.error('❌ メール処理エラー:', e.toString());
    throw e; // エラーをコンソールに表示
  }
}

// Config シートを開く
function openConfigSheet() {
  const sheet = ss().getSheetByName(CONFIG_SHEET);
  if (sheet) {
    ss().setActiveSheet(sheet);
  } else {
    console.error('Config シートが見つかりません。');
  }
}

// 処理済みラベル作成
function createProcessedLabel() {
  try {
    let label = GmailApp.getUserLabelByName('PROCESSED');
    if (!label) {
      label = GmailApp.createLabel('PROCESSED');
      console.log('✅ PROCESSEDラベルを作成しました。');
    } else {
      console.log('ℹ️ PROCESSEDラベルは既に存在します。');
    }
  } catch (e) {
    console.error('❌ エラー：', e);
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
      console.log('✅ フローシートをリセットしました。');
    }
  }
}

// 処理統計表示
function showProcessingStats() {
  const inboxSheet = ss().getSheetByName(INBOX_SHEET);
  if (!inboxSheet || inboxSheet.getLastRow() <= 1) {
    console.warn('処理データがありません。');
    return;
  }
  
  const lastRow = inboxSheet.getLastRow();
  const dataRows = Math.max(1, lastRow - 1);
  const data = inboxSheet.getRange(2, 7, dataRows, 1).getValues();
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
  let sheet = ss().getSheetByName(ACTIVITY_LOG_SHEET);
  
  // シートが存在しない場合は作成
  if (!sheet) {
    console.log('ActivityLogシートが存在しないため作成します');
    sheet = createActivityLogSheet();
  }
  
  if (sheet) {
    // 隠されている場合は表示
    if (sheet.isSheetHidden()) {
      sheet.showSheet();
    }
    ss().setActiveSheet(sheet);
    const ui = SpreadsheetApp.getUi();
    ui.alert('アクティビティログ', 'アクティビティログを表示しました。', ui.ButtonSet.OK);
  } else {
    console.error('アクティビティログシートの作成に失敗しました');
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
    `- 件名に [task] を含むメールを自動処理\n` +
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

// ================================================================================
// 6. openai_client.gs - OpenAI API連携機能
// ================================================================================

// OpenAI API設定
const OPENAI_URL_CHAT = 'https://api.openai.com/v1/chat/completions';
const OPENAI_URL_RESPONSES = 'https://api.openai.com/v1/responses'; // 将来的な拡張用

// JSON Schema定義
function buildWorkSpecSchema() {
  return {
    type: 'object',
    additionalProperties: false,
    properties: {
      work_spec: {
        type: 'object',
        additionalProperties: false,
        properties: {
          title: { type: 'string' },
          summary: { type: 'string' },
          scope: { type: 'string' },
          deliverables: { type: 'array', items: { type: 'string' } },
          org_structure: { type: 'array', items: { type: 'string' } },
          raci: {
            type: 'array',
            items: {
              type: 'object',
              additionalProperties: false,
              properties: {
                role: { type: 'string' },
                dept: { type: 'string' },
                R: { type: 'boolean' },
                A: { type: 'boolean' },
                C: { type: 'boolean' },
                I: { type: 'boolean' }
              },
              required: ['role', 'dept', 'R', 'A', 'C', 'I']
            }
          },
          timeline: {
            type: 'array',
            items: {
              type: 'object',
              additionalProperties: false,
              properties: {
                phase: { type: 'string' },
                duration_hint: { type: 'string' },
                milestones: { type: 'array', items: { type: 'string' } },
                dependencies: { type: 'array', items: { type: 'string' } }
              },
              required: ['phase', 'duration_hint', 'milestones', 'dependencies']
            }
          },
          requirements_constraints: { type: 'array', items: { type: 'string' } },
          risks_mitigations: { type: 'array', items: { type: 'string' } },
          pro_considerations: { type: 'array', items: { type: 'string' } },
          kpi_sla: { type: 'array', items: { type: 'string' } },
          approvals: { type: 'array', items: { type: 'string' } },
          security_privacy_controls: { type: 'array', items: { type: 'string' } },
          legal_regulations: {
            type: 'array',
            items: {
              type: 'object',
              additionalProperties: false,
              properties: {
                name: { type: 'string' },
                scope: { type: 'string' },
                note: { type: 'string' }
              },
              required: ['name', 'scope', 'note']
            }
          },
          references: { type: 'array', items: { type: 'string' } },
          assumptions: { type: 'array', items: { type: 'string' } }
        },
        required: [
          'title', 'summary', 'scope', 'deliverables', 'org_structure',
          'raci', 'timeline', 'requirements_constraints', 'risks_mitigations',
          'pro_considerations', 'kpi_sla', 'approvals', 'security_privacy_controls',
          'legal_regulations', 'references', 'assumptions'
        ]
      },
      flow_rows: {
        type: 'array',
        items: {
          type: 'object',
          additionalProperties: false,
          properties: {
            '工程': { type: 'string' },
            '実施タイミング': { type: 'string' },
            '部署': { type: 'string' },
            '担当役割': { type: 'string' },
            '作業内容': { type: 'string' },
            '条件分岐': { type: 'string' },
            '利用ツール': { type: 'string' },
            'URLリンク': { type: 'string' },
            '備考': { type: 'string' }
          },
          required: ['工程', '実施タイミング', '部署', '担当役割', '作業内容', '条件分岐', '利用ツール', 'URLリンク', '備考']
        }
      }
    },
    required: ['work_spec', 'flow_rows']
  };
}

// システムプロンプト構築
function buildSystemPrompt() {
  return `あなたは日本の上場企業（東証プライム市場）において、プロジェクトマネジメント、法務、内部統制、リスク管理の実務経験を20年以上持つ専門家です。

【あなたの専門性】
- 金融商品取引法、会社法、J-SOX法に精通
- 内部統制報告制度の構築・運用経験豊富
- ISO9001/27001、プライバシーマーク認証取得支援経験
- 監査法人対応、コーポレートガバナンス・コード対応の実績多数
- PMBOK、COBIT、COSOフレームワークの実装経験

【作成方針】
MECE（Mutually Exclusive, Collectively Exhaustive）の原則に基づき、以下の観点で業務を詳細に分解してください：

1. 法令・規制の具体的対応
   - 金融商品取引法（開示規制、内部統制報告制度）
   - 会社法（取締役会規程、監査役監査基準）
   - 個人情報保護法（プライバシーポリシー、同意取得）
   - 下請法、独占禁止法、労働基準法など関連法令
   - 業界特有の規制（金融業：銀行法、製造業：PL法など）
   - 各法令の具体的な条文番号まで特定
   - 違反時の罰則規定と影響範囲を明記

2. 内部統制・リスク管理の詳細設計
   - 3点セット（業務記述書、フローチャート、RCM）の作成
   - キーコントロールの特定と評価手続き
   - IT全般統制（ITGC）とIT業務処理統制（ITAC）
   - 職務分離（SoD）マトリックスの設計
   - 不正のトライアングル理論に基づく予防的統制
   - リスクアセスメント（発生可能性×影響度）
   - BCP/DR計画との連携

3. 具体的なアクション分解（MECE原則）
   - WBSレベル3以上の詳細度で作業を分解
   - 各タスクを15分～2時間単位の作業に細分化
   - 前工程・後工程の依存関係を明確化
   - クリティカルパスの特定
   - バッファ時間の設定（リスク対応）
   - チェックポイント、承認ポイントの明示
   - エスカレーションパスの定義

4. 実務的な具体例の提示
   - 使用する具体的な文書テンプレート名
   - 参照すべき社内規程・マニュアル名
   - 利用システム・ツールの具体名（SAP、Salesforce等）
   - 承認フロー（稟議システムの承認ルート）
   - 監査証跡の取得方法
   - KPIの計算式と測定頻度

5. ステークホルダー管理
   - RACIマトリックス（Responsible/Accountable/Consulted/Informed）
   - コミュニケーション計画（頻度、手段、参加者）
   - 報告書フォーマット（取締役会、監査役会向け）
   - 外部機関対応（監査法人、規制当局、証券取引所）

6. 上場企業特有の考慮事項（上場企業の場合に適用）
   - 東京証券取引所との関係：適時開示、コーポレートガバナンス報告書の提出、株式事務など
   - 金融庁・財務局との関係：有価証券報告書、内部統制報告書の提出、検査対応
   - 開示に関する観点：適時開示規則、インサイダー取引防止、IR活動の実施
   - これらの観点を業務フローに組み込み、必要なタスクとして明示
   - 入力に「株主総会」や「取締役会」などのキーワードが含まれる場合、開示義務（適時開示、法定開示）に関連するかを評価し、関連する場合、開示手続き、内部統制、監査対応などのタスクを追加

【重要な指示】
- 抽象的な表現を避け、実行可能な具体的アクションを記載
- 「検討する」→「〇〇の基準に基づき△△を評価し、□□を決定する」
- 「確認する」→「〇〇チェックリストの全項目が基準値を満たすことを確認する」
- 「管理する」→「〇〇管理台帳に記録し、週次で△△指標をモニタリングする」
- すべての法的記載には「※最終的には顧問弁護士・専門家による確認が必要」を付記

出力は日本語で、JSON Schema準拠。法的助言の代替ではないことを明記。`;
}

// ユーザープロンプト構築
function buildUserPrompt(mailBody, orgProfileJson) {
  const orgProfile = orgProfileJson ? JSON.parse(orgProfileJson) : {};
  
  return `以下の業務内容をMECE原則に基づき、実行可能な詳細タスクに分解して業務記述書とフロー表を作成してください。

【業務内容】
${mailBody}

【組織プロフィール】
- 上場区分: ${orgProfile.listing || '東証プライム'}
- 業種: ${orgProfile.industry || '製造業'}
- 対象地域: ${(orgProfile.jurisdictions || ['JP']).join(', ')}
- 社内基準: ${(orgProfile.policies || ['J-SOX対応', 'ISO27001認証']).join(', ')}
- 従業員数: ${orgProfile.employees || '1000名以上'}
- 売上規模: ${orgProfile.revenue || '100億円以上'}

【詳細化の要求水準】

1. タスクの粒度
   - 各タスクは最大2時間で完了可能な単位に分解
   - 具体的な成果物・アウトプットを明記
   - 判断基準・チェックポイントを数値化
   例：「確認する」→「〇〇チェックリスト25項目中23項目以上が基準値80%を超えることを確認」

2. 法令・規制の具体的記載
   - 法令名と該当条文を明記（例：金融商品取引法第24条）
   - 違反時のペナルティを記載（例：5年以下の懲役又は500万円以下の罰金）
   - 監督官庁への届出期限（例：変更後2週間以内に関東財務局へ届出）
   - 業界ガイドライン（例：日本証券業協会自主規制規則第〇条）

3. 内部統制の具体的設計
   - 予防的統制：承認権限規程（例：100万円以上は部長決裁）
   - 発見的統制：月次照合作業（例：売掛金残高と補助簿の照合）
   - IT統制：アクセスログの定期レビュー（例：特権IDの使用記録を週次確認）
   - 証跡保存：7年間の文書保存（電子帳簿保存法準拠）

4. リスク対応の詳細
   - リスクシナリオ：具体的な事象と発生確率（H/M/L）
   - 影響額：定量評価（売上の〇%相当）
   - 対応策：予防策、発生時対応、復旧計画
   - モニタリング指標：KRI（Key Risk Indicator）の設定

5. 実務ツール・システム
   - ERP：SAP S/4HANA（モジュール名まで特定）
   - ワークフロー：ServiceNow（申請フォームID）
   - 文書管理：SharePoint（フォルダ構成）
   - コミュニケーション：Teams（チャネル名）

【flow_rows作成の詳細要求】
各行は以下の粒度で記載：

- 工程：WBSレベル2（例：「1.2 要件定義フェーズ」）
- 実施タイミング：具体的な日付・期間（例：「2024年4月1日～4月15日（10営業日）」）
- 部署：正式部署名と人数（例：「経理部決算チーム（5名）」）
- 担当役割：RACI形式（例：「R:主任、A:課長、C:部長、I:監査役」）
- 作業内容：5W1H形式の具体的記述
  例：「売掛金年齢表を作成し、90日超の債権リスト（想定20件）を抽出。
       各債権について営業担当者へヒアリング（1件30分）を実施し、
       回収可能性を5段階評価。評価結果を貸倒引当金算定表に反映。」
- 条件分岐：判断基準を数値化（例：「売上高1000万円以上の場合は役員承認ルートへ」）
- 利用ツール：バージョンまで特定（例：「Excel 2021 売掛金管理テンプレートv3.2」）
- URLリンク：具体的な参照先（例：「社内ポータル/規程集/与信管理規程」）
- 備考：注意事項、過去の失敗事例、改善提案など
- 法令・規制：該当する具体的な法令・条文・ガイドライン
- 内部統制：統制活動の種類とコントロール番号（例：「予防的統制 CC-AR-001」）
- コンプライアンス留意点：過去の違反事例、監査指摘事項、業界のベストプラクティス

【最低限含めるべきタスク数】
- メインプロセス：最低15タスク
- 各タスクのサブタスク：3～5個
- チェックポイント：各フェーズに2箇所以上
- 承認ポイント：重要な意思決定箇所すべて

重要：抽象的な表現は使用禁止。すべて測定可能・実行可能な具体的記述にすること。`;
}

// OpenAI API呼び出し
function callOpenAI(mailBody, orgProfileJson) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
  if (!apiKey) {
    throw new Error('OpenAI APIキーが設定されていません。スクリプトプロパティに OPENAI_API_KEY を設定してください。');
  }
  
  const modelName = getConfig('OPENAI_MODEL') || 'gpt-5';
  const schema = buildWorkSpecSchema();
  
  // 特殊モデルの判定（現在は使用しない）
  const useResponsesEndpoint = false;
  
  if (useResponsesEndpoint) {
    return callOpenAIResponses(mailBody, orgProfileJson, apiKey, modelName, schema);
  }
  
  // 通常のチャットモデル（gpt-5）
  const messages = [
    { role: 'system', content: buildSystemPrompt() },
    { role: 'user', content: buildUserPrompt(mailBody, orgProfileJson) }
  ];
  
  // モデルによってresponse_formatを調整
  const supportsJsonSchema = true;
  const supportsJsonObject = false;
  const noResponseFormat = false;
  
  const payload = {
    model: modelName,
    messages: messages,
    seed: 42  // 再現性のためのシード値
  };
  
  payload.temperature = 0.1;  // より一貫した出力のために温度を下げる
  payload.max_tokens = 6000;
  
  if (supportsJsonSchema) {
    // json_schemaをサポートするモデルの場合
    payload.response_format = {
      type: 'json_schema',
      json_schema: {
        name: 'WorkSpecSchema',
        schema: schema,
        strict: true  // strictモードを有効化してデータ品質を向上
      }
    };
  } else if (supportsJsonObject) {
    // json_objectタイプをサポートするモデルの場合
    payload.response_format = { type: 'json_object' };
    
    // スキーマ情報をプロンプトに追加
    const schemaInstruction = `\n\n重要: 以下のJSONスキーマに厳密に従って出力してください：\n${JSON.stringify(schema, null, 2)}`;
    messages[messages.length - 1].content += schemaInstruction;
  } else {
    // response_formatをサポートしないモデルの場合
    // プロンプトでJSON出力を明示的に指示
    const enhancedSystemPrompt = messages[0].content + '\n\n重要: 必ず有効なJSONフォーマットで出力してください。マークダウンやコードブロック（```json```）は使用せず、純粋なJSONのみを出力してください。';
    messages[0].content = enhancedSystemPrompt;
    
    const schemaInstruction = `\n\n出力は以下のJSONスキーマに厳密に従ってください。追加のテキストや説明は一切含めないでください：\n${JSON.stringify(schema, null, 2)}`;
    messages[messages.length - 1].content += schemaInstruction;
  }
  
  logActivity('OPENAI_CALL', `Calling OpenAI Chat API with model: ${modelName}`);
  
  const response = retryWithBackoff(() => {
    const res = UrlFetchApp.fetch(OPENAI_URL_CHAT, {
      method: 'post',
      headers: {
        'Authorization': `Bearer ${apiKey}`,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    
    const status = res.getResponseCode();
    if (status === 404) {
      const errorBody = res.getContentText();
      if (errorBody.includes('v1/responses') || errorBody.includes('This model is only supported')) {
        // このモデルはv1/responsesエンドポイントを使用する必要がある
        logActivity('ENDPOINT_SWITCH', `Model ${modelName} requires v1/responses endpoint, switching...`);
        return callOpenAIResponses(mailBody, orgProfileJson, apiKey, modelName, schema);
      }
    }
    
    if (status >= 300) {
      const errorBody = res.getContentText();
      logActivity('OPENAI_ERROR', `Status: ${status}, Body: ${errorBody}`);
      throw new Error(`OpenAI API error ${status}: ${errorBody}`);
    }
    
    return res;
  });
  
  const responseData = JSON.parse(response.getContentText());
  const content = responseData.choices[0].message.content;
  
  logActivity('OPENAI_SUCCESS', 'Successfully received response from OpenAI');
  
  return JSON.parse(content);
}

// v1/responsesエンドポイント用のAPI呼び出し（gpt-5用）
function callOpenAIResponses(mailBody, orgProfileJson, apiKey, modelName, schema) {
  const requestId = `req_${new Date().getTime()}_${Math.random().toString(36).substr(2, 9)}`;
  logActivity('OPENAI_CALL', `Calling OpenAI Responses API with model: ${modelName}, Request ID: ${requestId}`);

  // システムプロンプトとユーザープロンプトを結合
  const systemPrompt = buildSystemPrompt();
  const userPrompt = buildUserPrompt(mailBody, orgProfileJson);

  // JSON出力を明示的に指示
  const enhancedPrompt = `${systemPrompt}\n\n${userPrompt}\n\n重要: 必ず有効なJSONフォーマットで出力してください。出力は以下のJSONスキーマに厳密に従ってください。追加のテキストや説明は一切含めないでください。\nJSON Schema: ${JSON.stringify(schema, null, 2)}`;

  // gpt-5はv1/responsesエンドポイントを使用
  let url = OPENAI_URL_RESPONSES;
  let payload = {
    model: modelName,
    messages: [
      {
        role: 'system',
        content: systemPrompt
      },
      {
        role: 'user',
        content: userPrompt
      }
    ],
    text: {
      format: 'json_schema',
      json_schema: {
        name: 'work_spec_response',
        strict: true,
        schema: schema
      }
    }
  };
  
  // gpt-5共通設定
  payload.temperature = 0.3;
  payload.max_output_tokens = 8000;

  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'Authorization': `Bearer ${apiKey}`,
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const response = retryWithBackoff(() => {
    const res = UrlFetchApp.fetch(url, options);
    const status = res.getResponseCode();
    
    if (status === 200) {
      return res;
    }
    
    if (status === 429) {
      const retryAfter = res.getHeaders()['Retry-After'] || '60';
      logActivity('OPENAI_RATE_LIMIT', `Rate limited. Retry after ${retryAfter} seconds`);
      throw new Error(`Rate limit exceeded. Retry after ${retryAfter} seconds`);
    }
    
    if (status >= 500) {
      logActivity('OPENAI_SERVER_ERROR', `Server error: ${status}`);
      throw new Error(`OpenAI server error: ${status}`);
    }
    
    if (status === 404) {
      // モデルが利用できない場合はgpt-5にフォールバック
      const errorBody = res.getContentText();
      logActivity('OPENAI_MODEL_ERROR', `Model ${effectiveModel} not available: ${errorBody}`);
      
      // gpt-5にフォールバック
      setConfig('OPENAI_MODEL', 'gpt-5');
      logActivity('MODEL_FALLBACK', 'Falling back to gpt-5');
      return callOpenAI(mailBody, orgProfileJson);  // 再帰的に呼び出し
    }
    
    const errorBody = res.getContentText();
    const errorDetails = JSON.stringify(payload, null, 2);
    logActivity('OPENAI_ERROR', `Status: ${status}, Body: ${errorBody}`);
    logActivity('OPENAI_PAYLOAD', `Failed payload: ${errorDetails}`);
    
    // より詳細なエラーメッセージ
    console.error('OpenAI API Error:', {
      status: status,
      error: errorBody,
      model: effectiveModel,
      endpoint: url
    });
    
    throw new Error(`OpenAI API error ${status}: ${errorBody}`);
  });
  
  // フォールバックの場合の型チェック
  if (response && typeof response.getContentText !== 'function') {
    // 既に解析済みのJSONの場合
    logActivity('FALLBACK_RESPONSE', 'Using fallback response');
    return response;
  }
  
  const responseData = JSON.parse(response.getContentText());
  
  logActivity('OPENAI_SUCCESS', 'Successfully received response from OpenAI Responses API');
  
  // JSONとしてパースを試みる
  try {
    return JSON.parse(responseData.choices[0].message.content);
  } catch (e) {
    // 既にオブジェクトの場合はそのまま返す
    if (typeof responseData.choices[0].message.content === 'object') {
      return responseData.choices[0].message.content;
    }
    throw new Error('Failed to parse OpenAI response: ' + e.toString());
  }
}

// JSON検証
function validateOpenAIResponse(data) {
  if (!data || typeof data !== 'object') {
    throw new Error('Invalid response format');
  }
  
  if (!data.work_spec || typeof data.work_spec !== 'object') {
    throw new Error('Missing or invalid work_spec');
  }
  
  if (!data.flow_rows || !Array.isArray(data.flow_rows)) {
    throw new Error('Missing or invalid flow_rows');
  }
  
  const ws = data.work_spec;
  const required = ['title', 'summary'];
  for (const field of required) {
    if (!ws[field]) {
      throw new Error(`Missing required field in work_spec: ${field}`);
    }
  }
  
  // flow_rowsの検証をより柔軟に
  for (let i = 0; i < data.flow_rows.length; i++) {
    const row = data.flow_rows[i];
    const requiredFlow = ['工程', '実施タイミング', '部署', '担当役割', '作業内容'];
    
    // 空値の自動補完
    if (!row['工程'] || row['工程'].trim() === '') {
      row['工程'] = `フェーズ${i + 1}`;
    }
    if (!row['実施タイミング'] || row['実施タイミング'].trim() === '') {
      row['実施タイミング'] = `第${i + 1}期`;
    }
    if (!row['部署'] || row['部署'].trim() === '') {
      row['部署'] = '経営企画部';
    }
    if (!row['担当役割'] || row['担当役割'].trim() === '') {
      row['担当役割'] = '担当者';
    }
    if (!row['作業内容'] || row['作業内容'].trim() === '') {
      row['作業内容'] = 'タスク実施';
    }
    
    // オプションフィールドのデフォルト値
    if (!row['条件分岐']) row['条件分岐'] = 'なし';
    if (!row['利用ツール']) row['利用ツール'] = '手動作業';
    if (!row['URLリンク']) row['URLリンク'] = 'なし';
    if (!row['備考']) row['備考'] = '特になし';
  }
  
  return true;
}

// ================================================================================
// 7. parser_and_writer.gs - データ解析・書き込み処理機能
// ================================================================================

// Config修正関数
function fixProcessingQuery() {
  const ui = SpreadsheetApp.getUi();
  const sh = ss().getSheetByName(CONFIG_SHEET);
  
  if (!sh) {
    ui.alert('エラー', 'Configシートが見つかりません。初回セットアップを実行してください。', ui.ButtonSet.OK);
    return;
  }
  
  // 現在の値を確認
  const currentQuery = getConfig('PROCESSING_QUERY');
  console.log('現在の検索クエリ:', currentQuery);
  
  // 古い値が含まれていたら修正
  if (currentQuery && (currentQuery.includes('WORK-REQ') || currentQuery.includes('label:inbox'))) {
    setConfig('PROCESSING_QUERY', 'subject:[task] is:unread');
    console.log('検索クエリを修正しました: subject:[task] is:unread');
    ui.alert('修正完了', '検索クエリを [task] 用に修正しました。\n新しいクエリ: subject:[task] is:unread', ui.ButtonSet.OK);
  } else if (!currentQuery) {
    setConfig('PROCESSING_QUERY', 'subject:[task] is:unread');
    console.log('検索クエリを設定しました: subject:[task] is:unread');
    ui.alert('設定完了', '検索クエリを設定しました。\nクエリ: subject:[task] is:unread', ui.ButtonSet.OK);
  } else {
    console.log('検索クエリは既に正しく設定されています');
    ui.alert('確認', `現在の検索クエリ:\n${currentQuery}\n\n既に正しく設定されています。`, ui.ButtonSet.OK);
  }
}

// メール検索テスト関数
function testEmailSearch() {
  console.log('===== メール検索テスト開始 =====');
  
  // 先に設定を修正
  fixProcessingQuery();
  
  // 1. 現在の設定を確認
  const currentQuery = getConfig('PROCESSING_QUERY');
  console.log('修正後の検索クエリ:', currentQuery);
  
  // 2. 様々なパターンで検索をテスト
  const testQueries = [
    'subject:[task]',
    'subject:"[task]"',
    'subject:task',
    '[task]',
    'is:unread subject:[task]',
    'is:unread subject:task',
    'is:unread',
    'in:anywhere [task]',
    'in:anywhere subject:task'
  ];
  
  console.log('\n各検索パターンの結果:');
  testQueries.forEach(query => {
    try {
      const threads = GmailApp.search(query, 0, 5);
      console.log(`  "${query}": ${threads.length}件`);
      
      if (threads.length > 0 && query.includes('task')) {
        // 最初のメールの件名を表示
        const firstMessage = threads[0].getMessages()[0];
        console.log(`    例: "${firstMessage.getSubject()}"`);
      }
    } catch (e) {
      console.log(`  "${query}": エラー - ${e.toString()}`);
    }
  });
  
  // 3. 未読メール全体から[task]を含むものを探す
  console.log('\n未読メールから[task]を含むものを検索:');
  const allUnread = GmailApp.search('is:unread', 0, 20);
  console.log(`未読メール総数: ${allUnread.length}件`);
  
  let taskCount = 0;
  allUnread.forEach(thread => {
    const firstMessage = thread.getMessages()[0];
    const subject = firstMessage.getSubject();
    
    // 様々なパターンでマッチング
    if (subject.includes('[task]') || 
        subject.toLowerCase().includes('[task]') ||
        subject.includes('task') ||
        subject.includes('【task】')) {
      taskCount++;
      console.log(`  ✓ "${subject}"`);
    }
  });
  
  console.log(`\n[task]関連メール: ${taskCount}件`);
  
  // 4. 推奨クエリの提案
  console.log('\n推奨される検索クエリ:');
  if (taskCount > 0) {
    console.log('  ・ "subject:task is:unread" - より広範囲にマッチ');
    console.log('  ・ "is:unread [task]" - 件名と本文から検索');
  }
  
  console.log('\n===== テスト完了 =====');
  
  // 結果をUIに表示
  const ui = SpreadsheetApp.getUi();
  ui.alert('検索テスト結果', 
    `現在のクエリ: ${currentQuery}\n` +
    `未読メール: ${allUnread.length}件\n` +
    `[task]関連: ${taskCount}件\n\n` +
    '詳細はログを確認してください', 
    ui.ButtonSet.OK);
}

// データ解析・書き込み処理

// フローシートのヘッダー定義（法令・内部統制の観点を追加）
const FLOW_HEADERS = [
  '工程', 
  '実施タイミング', 
  '部署', 
  '担当役割', 
  '作業内容', 
  '条件分岐', 
  '利用ツール', 
  'URLリンク', 
  '備考',
  '法令・規制',
  '内部統制',
  'コンプライアンス留意点'
];

// 業務記述書の書き込み
function writeWorkSpec(workSpec) {
  const sh = ss().getSheetByName(SPEC_SHEET) || createWorkSpecSheet();
  
  // IDを生成
  const id = Utilities.getUuid();
  const timestamp = new Date();
  
  // データ整形
  const rowData = [
    id,
    timestamp,
    workSpec.title || '',
    workSpec.summary || '',
    workSpec.scope || '',
    formatArray(workSpec.deliverables),
    formatArray(workSpec.org_structure),
    formatRaci(workSpec.raci),
    formatTimeline(workSpec.timeline),
    formatArray(workSpec.requirements_constraints),
    formatArray(workSpec.risks_mitigations),
    formatArray(workSpec.pro_considerations),
    formatArray(workSpec.kpi_sla),
    formatArray(workSpec.approvals),
    formatArray(workSpec.security_privacy_controls),
    formatLegalRegulations(workSpec.legal_regulations),
    formatArray(workSpec.references),
    formatArray(workSpec.assumptions)
  ];
  
  // データ書き込み
  sh.appendRow(rowData);
  
  // 書式設定
  const lastRow = sh.getLastRow();
  sh.getRange(lastRow, 1, 1, rowData.length).setWrap(true);
  sh.getRange(lastRow, 3).setFontWeight('bold'); // タイトルを太字
  
  logActivity('WRITE_SPEC', `Written work spec: ${workSpec.title}`);
}

// 業務記述書シート作成
function createWorkSpecSheet() {
  const sh = ss().insertSheet(SPEC_SHEET);
  
  const headers = [
    'ID',
    '作成日時',
    'タイトル',
    '概要',
    'スコープ',
    '成果物',
    '体制',
    'RACI',
    'スケジュール',
    '要件・制約',
    'リスク・対策',
    'プロ水準留意事項',
    'KPI/SLA',
    '承認フロー',
    'セキュリティ/個情保/内部統制',
    '法令・規制',
    '参考URL',
    '仮定条件'
  ];
  
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  sh.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  sh.getRange(1, 1, 1, headers.length).setBackground('#f0f0f0');
  sh.setFrozenRows(1);
  
  // 列幅調整
  sh.setColumnWidth(1, 150); // ID
  sh.setColumnWidth(2, 120); // 作成日時
  sh.setColumnWidth(3, 200); // タイトル
  sh.setColumnWidth(4, 300); // 概要
  
  return sh;
}

// フロー行の書き込み（レガシー関数 - 改善版にリダイレクト）
function writeFlowRows(flowRows) {
  // 新しい安全な実装を使用
  return writeFlowRowsSafe(flowRows);
}

// フローシート作成
function createFlowSheet(sheetName) {
  const sh = ss().insertSheet(sheetName);
  
  // ヘッダー行を設定
  sh.getRange(1, 1, 1, FLOW_HEADERS.length).setValues([FLOW_HEADERS]);
  sh.getRange(1, 1, 1, FLOW_HEADERS.length).setFontWeight('bold');
  sh.getRange(1, 1, 1, FLOW_HEADERS.length).setBackground('#f0f0f0');
  sh.setFrozenRows(1);
  
  // 列幅を調整（法令・内部統制の列を追加）
  sh.setColumnWidth(1, 120); // 工程
  sh.setColumnWidth(2, 150); // 実施タイミング
  sh.setColumnWidth(3, 120); // 部署
  sh.setColumnWidth(4, 150); // 担当役割
  sh.setColumnWidth(5, 300); // 作業内容
  sh.setColumnWidth(6, 150); // 条件分岐
  sh.setColumnWidth(7, 150); // 利用ツール
  sh.setColumnWidth(8, 200); // URLリンク
  sh.setColumnWidth(9, 200); // 備考
  sh.setColumnWidth(10, 250); // 法令・規制
  sh.setColumnWidth(11, 250); // 内部統制
  sh.setColumnWidth(12, 300); // コンプライアンス留意点
  
  return sh;
}

// フローシート作成
function createFlowSheet(sheetName) {
  const sh = ss().insertSheet(sheetName);
  
  sh.getRange(1, 1, 1, FLOW_HEADERS.length).setValues([FLOW_HEADERS]);
  sh.getRange(1, 1, 1, FLOW_HEADERS.length).setFontWeight('bold');
  sh.getRange(1, 1, 1, FLOW_HEADERS.length).setBackground('#e8f5e9');
  sh.setFrozenRows(1);
  
  // 列幅調整（法令・内部統制の列を追加）
  sh.setColumnWidth(1, 100); // 工程
  sh.setColumnWidth(2, 120); // 実施タイミング
  sh.setColumnWidth(3, 100); // 部署
  sh.setColumnWidth(4, 100); // 担当役割
  sh.setColumnWidth(5, 250); // 作業内容
  sh.setColumnWidth(6, 150); // 条件分岐
  sh.setColumnWidth(7, 120); // 利用ツール
  sh.setColumnWidth(8, 150); // URLリンク
  sh.setColumnWidth(9, 200); // 備考
  sh.setColumnWidth(10, 200); // 法令・規制
  sh.setColumnWidth(11, 200); // 内部統制
  sh.setColumnWidth(12, 250); // コンプライアンス留意点
  
  return sh;
}

// ビジュアルフロー生成関数
function generateVisualFlow() {
  try {
    console.log('=== ビジュアルフロー生成開始 ===');
    
    const flowSheetName = getConfig('FLOW_SHEET_NAME') || FLOW_SHEET;
    const sheet = ss().getSheetByName(flowSheetName);
    
    if (!sheet) {
      console.error('フローシートが見つかりません');
      return;
    }
    
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) {
      console.error('データがありません');
      return;
    }
    
    // ビジュアルシートの準備
    const visualSheetName = 'ビジュアルフロー';
    let visualSheet = ss().getSheetByName(visualSheetName);
    if (!visualSheet) {
      visualSheet = ss().insertSheet(visualSheetName);
    }
    visualSheet.clear();
    
    // データの解析
    const flowData = parseFlowDataForVisual(data);
    
    // フローチャートの描画
    drawVisualFlowChart(visualSheet, flowData);
    
    console.log('=== ビジュアルフロー生成完了 ===');
    
  } catch (error) {
    console.error('ビジュアルフロー生成エラー:', error);
  }
}

// ビジュアル用フローデータの解析
function parseFlowDataForVisual(data) {
  const headers = data[0];
  const columnIndex = {};
  headers.forEach((header, index) => {
    columnIndex[header] = index;
  });
  
  const flowData = {
    departments: {},
    departmentList: [],
    timings: [],
    tools: new Set(),
    datasources: {},
    processName: ''
  };
  
  // プロセス名の取得
  if (data.length > 1 && data[1][columnIndex['工程']]) {
    flowData.processName = data[1][columnIndex['工程']];
  }
  
  // データの整理
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[columnIndex['工程']] || row[columnIndex['工程']] === '') continue;
    
    const dept = row[columnIndex['部署']] || 'その他';
    const timing = row[columnIndex['実施タイミング']] || '';
    const tool = row[columnIndex['利用ツール']] || '';
    const url = row[columnIndex['URLリンク']] || '';
    
    // 部署の初期化
    if (!flowData.departments[dept]) {
      flowData.departments[dept] = {};
      if (!flowData.departmentList.includes(dept)) {
        flowData.departmentList.push(dept);
      }
    }
    
    // タイミングの追加
    if (timing && !flowData.timings.includes(timing)) {
      flowData.timings.push(timing);
    }
    
    // タスクの追加
    if (!flowData.departments[dept][timing]) {
      flowData.departments[dept][timing] = [];
    }
    
    flowData.departments[dept][timing].push({
      task: row[columnIndex['作業内容']] || '',
      role: row[columnIndex['担当役割']] || '',
      condition: row[columnIndex['条件分岐']] || '',
      tool: tool,
      url: url,
      note: row[columnIndex['備考']] || '',
      legal: row[columnIndex['法令・規制']] || '',
      control: row[columnIndex['内部統制']] || '',
      compliance: row[columnIndex['コンプライアンス留意点']] || ''
    });
    
    // ツールの収集
    if (tool && tool !== '-' && tool !== 'なし') {
      const tools = tool.split(/[／、,]/);
      tools.forEach(t => {
        const trimmedTool = t.trim();
        if (trimmedTool) {
          flowData.tools.add(trimmedTool);
        }
      });
    }
    
    // URLをタイミングごとに管理
    if (url && url !== '-' && url !== 'なし') {
      if (!flowData.datasources[timing]) {
        flowData.datasources[timing] = [];
      }
      if (!flowData.datasources[timing].includes(url)) {
        flowData.datasources[timing].push(url);
      }
    }
  }
  
  return flowData;
}

// 配列データのフォーマット
function formatArray(arr) {
  if (!arr || !Array.isArray(arr)) return '';
  return arr.filter(item => item).join('\n');
}

// リスクデータのフォーマット
function formatRisks(risks) {
  if (!risks || !Array.isArray(risks)) return '';
  
  return risks.map(risk => {
    if (typeof risk === 'string') {
      return risk;
    } else if (typeof risk === 'object' && risk !== null) {
      const parts = [];
      if (risk.risk) parts.push(`リスク: ${risk.risk}`);
      if (risk.mitigation) parts.push(`対策: ${risk.mitigation}`);
      if (risk.probability) parts.push(`確率: ${risk.probability}`);
      if (risk.impact) parts.push(`影響: ${risk.impact}`);
      return parts.join(' / ');
    }
    return '';
  }).filter(item => item).join('\n');
}

// ビジュアルフローチャートの描画
function drawVisualFlowChart(sheet, flowData) {
  // カラーパレット
  const COLORS = {
    HEADER: '#4A5568',
    TIMELINE: '#F7FAFC',
    PROCESS: '#87CEEB',
    DECISION: '#FFD700',
    START_END: '#90EE90',
    EMPTY: '#FAFAFA',
    DATASOURCE: '#E3F2FD'
  };
  
  let currentRow = 1;
  
  // タイトル行
  const flowTitle = flowData.processName || '業務フロー';
  const maxCols = Math.max(flowData.departmentList.length + 2, 10);
  sheet.getRange(currentRow, 1, 1, maxCols).merge();
  const titleCell = sheet.getRange(currentRow, 1);
  titleCell.setValue(flowTitle);
  titleCell.setBackground(COLORS.HEADER);
  titleCell.setFontColor('#FFFFFF');
  titleCell.setFontSize(16);
  titleCell.setFontWeight('bold');
  titleCell.setHorizontalAlignment('center');
  titleCell.setVerticalAlignment('middle');
  sheet.setRowHeight(currentRow, 50);
  currentRow++;
  
  // ヘッダー行（タイムライン + 部署）
  sheet.getRange(currentRow, 1).setValue('タイミング');
  sheet.getRange(currentRow, 1).setBackground(COLORS.TIMELINE);
  sheet.getRange(currentRow, 1).setFontWeight('bold');
  
  flowData.departmentList.forEach((dept, index) => {
    sheet.getRange(currentRow, index + 2).setValue(dept);
    sheet.getRange(currentRow, index + 2).setBackground('#E8EAF6');
    sheet.getRange(currentRow, index + 2).setFontWeight('bold');
    sheet.getRange(currentRow, index + 2).setHorizontalAlignment('center');
  });
  
  if (Object.keys(flowData.datasources).length > 0) {
    const dataCol = flowData.departmentList.length + 2;
    sheet.getRange(currentRow, dataCol).setValue('関連資料');
    sheet.getRange(currentRow, dataCol).setBackground('#E3F2FD');
    sheet.getRange(currentRow, dataCol).setFontWeight('bold');
  }
  
  sheet.setRowHeight(currentRow, 40);
  currentRow++;
  
  // 開始行
  sheet.getRange(currentRow, 1).setValue('【開始】');
  sheet.getRange(currentRow, 1).setBackground(COLORS.START_END);
  sheet.getRange(currentRow, 1).setFontWeight('bold');
  currentRow++;
  
  // 各タイミングの行
  flowData.timings.forEach((timing) => {
    sheet.getRange(currentRow, 1).setValue(timing);
    sheet.getRange(currentRow, 1).setBackground(COLORS.TIMELINE);
    sheet.getRange(currentRow, 1).setFontWeight('bold');
    sheet.getRange(currentRow, 1).setVerticalAlignment('middle');
    
    // 各部署のタスク
    flowData.departmentList.forEach((dept, deptIndex) => {
      const col = deptIndex + 2;
      if (flowData.departments[dept] && flowData.departments[dept][timing]) {
        const tasks = flowData.departments[dept][timing];
        const taskTexts = tasks.map(t => {
          let text = t.task;
          if (t.role) text = `[${t.role}] ${text}`;
          if (t.condition && t.condition !== 'なし') text = `◆ ${text}`;
          return text;
        });
        
        const cell = sheet.getRange(currentRow, col);
        cell.setValue(taskTexts.join('\n'));
        
        // 条件分岐があるかチェック
        const hasCondition = tasks.some(t => t.condition && t.condition !== 'なし');
        cell.setBackground(hasCondition ? COLORS.DECISION : COLORS.PROCESS);
        
        cell.setWrap(true);
        cell.setVerticalAlignment('top');
        cell.setBorder(true, true, true, true, false, false);
        
        // ツール情報と法令・内部統制情報をノートに追加
        const noteItems = [];
        const tools = tasks.map(t => t.tool).filter(t => t && t !== 'なし').join(', ');
        if (tools) {
          noteItems.push(`使用ツール: ${tools}`);
        }
        
        const legals = tasks.map(t => t.legal).filter(t => t && t !== 'なし');
        if (legals.length > 0) {
          noteItems.push(`法令・規制: ${[...new Set(legals)].join(', ')}`);
        }
        
        const controls = tasks.map(t => t.control).filter(t => t && t !== 'なし');
        if (controls.length > 0) {
          noteItems.push(`内部統制: ${[...new Set(controls)].join(', ')}`);
        }
        
        const compliances = tasks.map(t => t.compliance).filter(t => t && t !== '特になし');
        if (compliances.length > 0) {
          noteItems.push(`留意点: ${compliances[0]}`); // 最初の留意点のみ表示
        }
        
        if (noteItems.length > 0) {
          cell.setNote(noteItems.join('\n\n'));
        }
      } else {
        sheet.getRange(currentRow, col).setBackground(COLORS.EMPTY);
      }
    });
    
    // データソース列
    if (Object.keys(flowData.datasources).length > 0) {
      const dataCol = flowData.departmentList.length + 2;
      if (flowData.datasources[timing] && flowData.datasources[timing].length > 0) {
        const urls = flowData.datasources[timing].join('\n');
        const dataCell = sheet.getRange(currentRow, dataCol);
        dataCell.setValue('📄 ' + urls);
        dataCell.setBackground(COLORS.DATASOURCE);
        dataCell.setWrap(true);
      }
    }
    
    sheet.setRowHeight(currentRow, 90);
    currentRow++;
  });
  
  // 終了行
  sheet.getRange(currentRow, 1).setValue('【完了】');
  sheet.getRange(currentRow, 1).setBackground(COLORS.START_END);
  sheet.getRange(currentRow, 1).setFontWeight('bold');
  
  const mergeCols = Object.keys(flowData.datasources).length > 0 ? 
                     flowData.departmentList.length + 1 : 
                     flowData.departmentList.length;
  sheet.getRange(currentRow, 2, 1, mergeCols).merge();
  const msgCell = sheet.getRange(currentRow, 2);
  msgCell.setValue('✅ プロセス完了');
  msgCell.setBackground('#E8F5E9');
  msgCell.setFontSize(14);
  msgCell.setFontWeight('bold');
  msgCell.setHorizontalAlignment('center');
  sheet.setRowHeight(currentRow, 50);
  currentRow++;
  
  // 凡例行
  sheet.getRange(currentRow, 1, 1, maxCols).merge();
  const legendCell = sheet.getRange(currentRow, 1);
  legendCell.setValue('凡例： □ 処理・作業　◆ 判断・分岐　📄 関連資料　※セルのノートに法令・内部統制・コンプライアンス情報があります');
  legendCell.setBackground(COLORS.TIMELINE);
  legendCell.setFontWeight('bold');
  sheet.setRowHeight(currentRow, 40);
  
  // 列幅の調整
  sheet.setColumnWidth(1, 150); // タイムライン列
  for (let i = 2; i <= flowData.departmentList.length + 1; i++) {
    sheet.setColumnWidth(i, 200); // 部署列
  }
  if (Object.keys(flowData.datasources).length > 0) {
    sheet.setColumnWidth(flowData.departmentList.length + 2, 150); // データソース列
  }
  
  // 全体に罫線を設定
  const range = sheet.getRange(1, 1, currentRow, maxCols);
  range.setBorder(true, true, true, true, true, true, '#d0d0d0', SpreadsheetApp.BorderStyle.SOLID);
}

// RACIマトリクスのフォーマット
function formatRaci(raciArray) {
  if (!raciArray || !Array.isArray(raciArray)) return '';
  
  return raciArray.map(item => {
    const roles = [];
    if (item.R) roles.push('R');
    if (item.A) roles.push('A');
    if (item.C) roles.push('C');
    if (item.I) roles.push('I');
    
    return `${item.dept || ''} - ${item.role || ''}: ${roles.join('')}`;
  }).join('\n');
}

// タイムラインのフォーマット
function formatTimeline(timeline) {
  if (!timeline || !Array.isArray(timeline)) return '';
  
  return timeline.map(phase => {
    let result = `【${phase.phase}】 ${phase.duration_hint || ''}`;
    
    if (phase.milestones && phase.milestones.length > 0) {
      result += '\nマイルストーン:\n' + phase.milestones.map(m => `  ・${m}`).join('\n');
    }
    
    if (phase.dependencies && phase.dependencies.length > 0) {
      result += '\n依存関係:\n' + phase.dependencies.map(d => `  ・${d}`).join('\n');
    }
    
    return result;
  }).join('\n\n');
}

// 法令・規制のフォーマット
function formatLegalRegulations(regulations) {
  if (!regulations || !Array.isArray(regulations)) return '';
  
  const formatted = regulations.map(reg => {
    let result = reg.name || '';
    if (reg.scope) result += `（${reg.scope}）`;
    if (reg.note) result += `: ${reg.note}`;
    return result;
  }).join('\n');
  
  // 法的助言の免責事項を追加
  return formatted + '\n\n※ 上記は参考情報です。最終的な判断は法務・専門家にご確認ください。法的助言ではありません。';
}

// 生データ保存（エラー時のフォールバック）
function saveRawData(data, error) {
  const sheetName = '業務記述書（Raw）';
  let sh = ss().getSheetByName(sheetName);
  
  if (!sh) {
    sh = ss().insertSheet(sheetName);
    sh.getRange(1, 1, 1, 4).setValues([['タイムスタンプ', 'エラー', 'データタイプ', '生データ']]);
    sh.getRange(1, 1, 1, 4).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  
  sh.appendRow([
    new Date(),
    error.toString(),
    typeof data,
    JSON.stringify(data, null, 2)
  ]);
  
  logActivity('SAVE_RAW', 'Saved raw data due to error');
}

// データ検証とクリーニング
function sanitizeData(data) {
  if (!data || typeof data !== 'object') return data;
  
  // 再帰的にオブジェクトをクリーニング
  const cleaned = {};
  for (const key in data) {
    if (data.hasOwnProperty(key)) {
      const value = data[key];
      
      if (value === null || value === undefined) {
        cleaned[key] = '';
      } else if (Array.isArray(value)) {
        cleaned[key] = value.map(item => 
          typeof item === 'object' ? sanitizeData(item) : item
        );
      } else if (typeof value === 'object') {
        cleaned[key] = sanitizeData(value);
      } else {
        cleaned[key] = value;
      }
    }
  }
  
  return cleaned;
}

// 個人情報マスキング（オプション）
function maskSensitiveInfo(text) {
  if (!text || typeof text !== 'string') return text;
  
  // メールアドレスのマスキング
  text = text.replace(/([a-zA-Z0-9._%+-]+)@([a-zA-Z0-9.-]+\.[a-zA-Z]{2,})/g, 
    (match, p1, p2) => p1.substring(0, 2) + '***@' + p2);
  
  // 電話番号のマスキング（日本の形式）
  text = text.replace(/(\d{2,4})-(\d{2,4})-(\d{4})/g, '$1-****-****');
  text = text.replace(/0\d{1,4}-\d{1,4}-\d{4}/g, '0**-****-****');
  
  // 郵便番号のマスキング
  text = text.replace(/〒?\d{3}-\d{4}/g, '〒***-****');
  
  return text;
}

// ================================================================================
// 9. data_processor.gs - 根本的に改善されたデータ処理エンジン
// ================================================================================

// データ処理の根本的改善
// 新しいデータ処理エンジン

// データ正規化クラス
class DataNormalizer {
  constructor() {
    this.cleaningPatterns = [
      // 末尾数字パターン
      { pattern: /\d+$/, replacement: '', condition: (str) => str.length > 1 },
      // 特定文字列後の数字
      { pattern: /特になし\d+$/, replacement: '特になし', condition: () => true },
      { pattern: /なし\d+$/, replacement: 'なし', condition: () => true },
      // 一般的な末尾数字（文字の後に数字）
      { pattern: /([^\d])\d+$/, replacement: '$1', condition: (str) => str.length > 2 }
    ];
  }

  // 文字列のクリーニング
  cleanString(value) {
    if (!value || typeof value !== 'string') {
      return String(value || '').trim();
    }

    let cleaned = value.trim();
    const original = cleaned;

    for (const rule of this.cleaningPatterns) {
      if (rule.condition(cleaned) && rule.pattern.test(cleaned)) {
        const newValue = cleaned.replace(rule.pattern, rule.replacement).trim();
        if (newValue.length > 0 && newValue !== cleaned) {
          console.log(`データクリーニング: "${cleaned}" -> "${newValue}"`);
          cleaned = newValue;
          break; // 最初にマッチしたルールのみ適用
        }
      }
    }

    return cleaned;
  }

  // フロー行データの正規化
  normalizeFlowRow(row, index) {
    const normalizedRow = {};
    const requiredFields = ['工程', '実施タイミング', '部署', '担当役割', '作業内容'];
    const optionalFields = ['条件分岐', '利用ツール', 'URLリンク', '備考'];

    // 必須フィールドの処理
    for (const field of requiredFields) {
      let value = row[field];
      
      if (!value || String(value).trim() === '') {
        // デフォルト値を設定
        switch (field) {
          case '工程':
            value = `フェーズ${index + 1}`;
            break;
          case '実施タイミング':
            value = `第${index + 1}期`;
            break;
          case '部署':
            value = '経営企画部';
            break;
          case '担当役割':
            value = '担当者';
            break;
          case '作業内容':
            value = 'タスク実施';
            break;
        }
        console.log(`デフォルト値設定: ${field} = "${value}"`);
      }

      normalizedRow[field] = this.cleanString(value);
    }

    // オプションフィールドの処理
    for (const field of optionalFields) {
      let value = row[field] || '';
      if (field === '条件分岐' && !value) {
        value = 'なし';
      }
      normalizedRow[field] = this.cleanString(value);
    }

    return normalizedRow;
  }

  // 2次元配列への変換（厳密な型保証）
  convertToSpreadsheetArray(flowRows, headers) {
    if (!Array.isArray(flowRows) || !Array.isArray(headers)) {
      throw new Error('Invalid input: flowRows and headers must be arrays');
    }

    const result = [];
    
    for (let i = 0; i < flowRows.length; i++) {
      const row = flowRows[i];
      const arrayRow = [];

      for (const header of headers) {
        const value = row[header] || '';
        arrayRow.push(String(value)); // 明示的に文字列変換
      }

      // 配列の長さを確認
      if (arrayRow.length !== headers.length) {
        throw new Error(`Row ${i + 1}: Expected ${headers.length} columns, got ${arrayRow.length}`);
      }

      result.push(arrayRow);
    }

    console.log(`変換完了: ${result.length}行 x ${headers.length}列の2次元配列`);
    return result;
  }
}

// 安全なスプレッドシート書き込みクラス
class SafeSpreadsheetWriter {
  constructor(sheet, headers) {
    this.sheet = sheet;
    this.headers = headers;
    this.normalizer = new DataNormalizer();
  }

  // データの書き込み（複数の安全策を実装）
  writeData(flowRows) {
    if (!flowRows || flowRows.length === 0) {
      console.log('書き込むデータがありません');
      return;
    }

    try {
      // Step 1: データの正規化
      const normalizedRows = flowRows.map((row, index) => 
        this.normalizer.normalizeFlowRow(row, index)
      );

      // Step 2: 2次元配列への変換
      const spreadsheetArray = this.normalizer.convertToSpreadsheetArray(normalizedRows, this.headers);

      // Step 3: 既存データのクリア
      this.clearExistingData();

      // Step 4: 安全な書き込み
      this.performSafeWrite(spreadsheetArray);

      // Step 5: 書式設定
      this.applyFormatting(spreadsheetArray.length);

      console.log(`データ書き込み完了: ${spreadsheetArray.length}行`);

    } catch (error) {
      console.error('データ書き込みエラー:', error);
      throw new Error(`データ書き込み失敗: ${error.message}`);
    }
  }

  // 既存データのクリア
  clearExistingData() {
    const lastRow = this.sheet.getLastRow();
    if (lastRow > 1) {
      const dataRows = Math.max(1, lastRow - 1);
      this.sheet.getRange(2, 1, dataRows, this.headers.length).clearContent();
    }
  }

  // 安全な書き込み実行
  performSafeWrite(data) {
    try {
      // データの検証
      if (!Array.isArray(data) || data.length === 0) {
        console.log('書き込むデータが無効または空です');
        return;
      }
      
      // データ構造のログ出力
      console.log(`書き込みデータ: ${data.length}行 x ${this.headers.length}列`);
      console.log('最初の行サンプル:', data[0]);
      
      // 数値パラメータの検証
      const startRow = 2;
      const startCol = 1;
      const numRows = Number(data.length);
      const numCols = Number(this.headers.length);
      
      if (isNaN(numRows) || isNaN(numCols)) {
        throw new Error(`無効な行数または列数: rows=${numRows}, cols=${numCols}`);
      }
      
      // 一括書き込みを試行
      this.sheet.getRange(startRow, startCol, numRows, numCols).setValues(data);
      console.log('一括書き込み成功');
    } catch (error) {
      console.warn('一括書き込み失敗、行ごと書き込みに切り替え:', error.message);
      console.error('エラー詳細:', error);
      this.writeRowByRow(data);
    }
  }

  // 行ごとの書き込み
  writeRowByRow(data) {
    for (let i = 0; i < data.length; i++) {
      try {
        const rowNum = Number(i + 2);
        const colStart = 1;
        const numRows = 1;
        const numCols = Number(this.headers.length);
        
        if (!Array.isArray(data[i])) {
          console.error(`Row ${i + 1} が配列ではありません:`, data[i]);
          continue;
        }
        
        this.sheet.getRange(rowNum, colStart, numRows, numCols).setValues([data[i]]);
        console.log(`Row ${i + 1} 書き込み成功`);
      } catch (error) {
        console.error(`Row ${i + 1} 書き込みエラー:`, error.message);
        console.error(`問題のデータ:`, data[i]);
        // セルごとの書き込みにフォールバック
        this.writeCellByCell(i + 2, data[i]);
      }
    }
  }

  // セルごとの書き込み
  writeCellByCell(rowIndex, rowData) {
    for (let j = 0; j < rowData.length; j++) {
      try {
        this.sheet.getRange(rowIndex, j + 1).setValue(rowData[j]);
      } catch (error) {
        console.error(`Cell (${rowIndex}, ${j + 1}) 書き込みエラー:`, error.message);
        this.sheet.getRange(rowIndex, j + 1).setValue('エラー');
      }
    }
  }

  // 書式設定
  applyFormatting(rowCount) {
    try {
      // テキストの折り返し
      this.sheet.getRange(2, 1, rowCount, this.headers.length).setWrap(true);
      
      // 工程列を太字
      this.sheet.getRange(2, 1, rowCount, 1).setFontWeight('bold');
      
      // 条件分岐がある行の背景色設定
      for (let i = 0; i < rowCount; i++) {
        const conditionValue = this.sheet.getRange(i + 2, 6).getValue(); // 条件分岐列
        if (conditionValue && conditionValue !== 'なし' && conditionValue !== '') {
          this.sheet.getRange(i + 2, 1, 1, this.headers.length).setBackground('#fff3cd');
        }
      }
    } catch (error) {
      console.warn('書式設定エラー:', error.message);
    }
  }
}

// 改善されたフロー行書き込み関数（新しい安全な実装にリダイレクト）
function writeFlowRowsImproved(flowRows) {
  return writeFlowRowsSafe(flowRows);
}

// 法令・規制をチェックする関数（ガバナンス強化版）
function checkLegalRegulations(processName, workContent, timing, dept) {
  const regulations = [];
  
  // 外部専門家相談の必要性を判定
  const advisorsNeeded = determineRequiredAdvisors(processName + ' ' + workContent);
  if (advisorsNeeded.length > 0) {
    regulations.push('【要専門家相談】' + advisorsNeeded.map(a => a.type).join('、'));
  }
  
  // 開示要件チェック
  const disclosureCheck = checkDisclosureRequirement(processName + ' ' + workContent);
  if (disclosureCheck.requiresDisclosure) {
    regulations.push('【要開示】' + disclosureCheck.disclosureType.join('、'));
  }
  
  // 株主総会関連
  if (processName.includes('株主総会') || workContent.includes('株主総会')) {
    regulations.push('会社法（第295条〜第325条）');
    if (timing.includes('6月')) {
      regulations.push('定時株主総会（会社法第296条）');
    }
    if (workContent.includes('招集通知')) {
      regulations.push('招集通知期限（会社法第299条：2週間前）');
    }
    if (workContent.includes('議決権')) {
      regulations.push('議決権行使（会社法第308条〜第313条）');
    }
  }
  
  // 決算・開示関連
  if (processName.includes('決算') || workContent.includes('決算') || workContent.includes('開示')) {
    regulations.push('金融商品取引法');
    if (workContent.includes('四半期')) {
      regulations.push('四半期報告書（金商法第24条の4の7）45日以内');
    }
    if (workContent.includes('有価証券報告書')) {
      regulations.push('有価証券報告書（金商法第24条）3ヶ月以内');
    }
    if (workContent.includes('内部統制報告書')) {
      regulations.push('内部統制報告書（金商法第24条の4の4）');
    }
  }
  
  // 取締役会関連（ガバナンス強化）
  if (workContent.includes('取締役会') || dept.includes('取締役')) {
    regulations.push('会社法第362条（取締役会の権限）');
    if (workContent.includes('議事録')) {
      regulations.push('会社法第369条（取締役会議事録）');
    }
    // 重要事項の判定
    if (workContent.includes('重要') || workContent.includes('決議')) {
      regulations.push('【重要決議】東証への適時開示検討');
      regulations.push('【専門家相談】法律事務所への事前確認推奨');
    }
  }
  
  // 監査関連
  if (workContent.includes('監査') || dept.includes('監査')) {
    if (workContent.includes('会計監査')) {
      regulations.push('会社法第436条（計算書類の監査）');
      regulations.push('金商法第193条の2（監査証明）');
    }
    if (workContent.includes('内部監査')) {
      regulations.push('J-SOX（金商法第24条の4の4）');
    }
  }
  
  // 個人情報保護
  if (workContent.includes('個人情報') || workContent.includes('顧客情報')) {
    regulations.push('個人情報保護法');
    regulations.push('GDPR（EU居住者データを扱う場合）');
  }
  
  // インサイダー取引規制
  if (workContent.includes('重要事実') || workContent.includes('適時開示')) {
    regulations.push('金商法第166条（インサイダー取引規制）');
    regulations.push('東証適時開示規則');
  }
  
  // 労働関連
  if (dept.includes('人事') || workContent.includes('労働') || workContent.includes('雇用')) {
    regulations.push('労働基準法');
    if (workContent.includes('36協定')) {
      regulations.push('労基法第36条（時間外労働）');
    }
  }
  
  return regulations.length > 0 ? regulations.join('、') : 'なし';
}

// 内部統制の観点をチェックする関数
function checkInternalControl(processName, workContent, condition, dept) {
  const controls = [];
  
  // 職務分離
  if (condition && condition !== 'なし') {
    controls.push('職務分離の原則');
  }
  
  // 承認権限
  if (workContent.includes('承認') || workContent.includes('決裁')) {
    controls.push('承認権限規程の遵守');
    if (workContent.includes('金額')) {
      controls.push('金額基準による承認権限の設定');
    }
  }
  
  // 文書化
  if (workContent.includes('記録') || workContent.includes('議事録') || workContent.includes('文書')) {
    controls.push('文書化（Documentation）');
    controls.push('監査証跡の保持');
  }
  
  // IT統制
  if (workContent.includes('システム') || workContent.includes('データ')) {
    controls.push('IT全般統制（ITGC）');
    if (workContent.includes('アクセス')) {
      controls.push('アクセス権限管理');
    }
    if (workContent.includes('バックアップ')) {
      controls.push('データバックアップ体制');
    }
  }
  
  // 財務報告
  if (processName.includes('決算') || workContent.includes('財務') || workContent.includes('会計')) {
    controls.push('財務報告に係る内部統制（J-SOX）');
    if (workContent.includes('仕訳')) {
      controls.push('仕訳承認プロセス');
    }
  }
  
  // リスク評価
  if (workContent.includes('リスク') || workContent.includes('評価')) {
    controls.push('リスク評価と対応');
    controls.push('COSOフレームワーク準拠');
  }
  
  // モニタリング
  if (workContent.includes('確認') || workContent.includes('検証') || workContent.includes('レビュー')) {
    controls.push('独立的モニタリング');
    controls.push('予防的統制');
  }
  
  // 相互牽制
  if (dept.includes('経理') || dept.includes('財務')) {
    controls.push('相互牽制体制');
    if (workContent.includes('出納') || workContent.includes('支払')) {
      controls.push('出納業務の分離');
    }
  }
  
  return controls.length > 0 ? controls.join('、') : 'なし';
}

// コンプライアンス留意点を生成する関数（ガバナンス強化版）
function generateComplianceNotes(processName, workContent, timing, dept, condition) {
  const notes = [];
  
  // 外部専門家相談チェックリスト生成
  const taskDescription = `${processName} - ${workContent}`;
  const requiredAdvisors = determineRequiredAdvisors(taskDescription);
  if (requiredAdvisors.length > 0) {
    const checklist = generateConsultationChecklist(taskDescription, requiredAdvisors);
    notes.push('【最優先】外部専門家への事前相談実施');
    checklist.consultationSteps.forEach(step => {
      if (step.phase.includes('専門家')) {
        notes.push(`- ${step.phase}: ${step.timeline}`);
      }
    });
  }
  
  // 株主総会特有の留意点
  if (processName.includes('株主総会')) {
    notes.push('【重要】招集通知は法定期限（2週間前）を厳守');
    if (workContent.includes('議決権')) {
      notes.push('議決権行使書の管理を徹底（改ざん防止）');
    }
    if (workContent.includes('質問')) {
      notes.push('想定問答集の事前準備と法務確認');
    }
  }
  
  // 開示関連
  if (workContent.includes('開示') || workContent.includes('IR')) {
    notes.push('【開示】東証への事前相談を検討');
    notes.push('公平開示の原則を遵守（フェア・ディスクロージャー）');
    if (workContent.includes('業績')) {
      notes.push('業績予想の修正は速やかに開示（軽微基準の確認）');
    }
  }
  
  // 決算関連
  if (processName.includes('決算') || workContent.includes('決算')) {
    notes.push('【決算】会計監査人との事前協議を実施');
    notes.push('重要な会計上の見積りは文書化');
    if (timing.includes('四半期')) {
      notes.push('四半期レビュー対応（監査より簡易だが重要）');
    }
  }
  
  // インサイダー情報管理
  if (workContent.includes('重要事実') || workContent.includes('未公表')) {
    notes.push('【インサイダー】情報管理を徹底（need to know原則）');
    notes.push('役職員の自社株売買は事前申請制');
  }
  
  // データ保護
  if (workContent.includes('個人情報') || workContent.includes('データ')) {
    notes.push('【個人情報】取得時に利用目的を明示');
    notes.push('第三者提供には本人同意が必要');
    if (workContent.includes('削除') || workContent.includes('廃棄')) {
      notes.push('データ削除は復元不可能な方法で実施');
    }
  }
  
  // 契約関連
  if (workContent.includes('契約') || workContent.includes('締結')) {
    notes.push('【契約】法務部門の事前レビュー必須');
    notes.push('利益相反取引は取締役会承認が必要');
  }
  
  // 監査対応
  if (workContent.includes('監査')) {
    notes.push('【監査】監査調書は7年間保存');
    notes.push('監査人の独立性を阻害する行為は禁止');
  }
  
  // リスク管理全般
  if (condition && condition !== 'なし') {
    notes.push('【統制】判断基準を明文化し、恣意性を排除');
    notes.push('例外処理は必ず上位者の承認を取得');
  }
  
  // タイミングに関する留意点
  if (timing.includes('期限') || timing.includes('以内')) {
    notes.push('【期限】法定期限がある場合は余裕を持ったスケジュール設定');
  }
  
  return notes.length > 0 ? notes.join('\n') : '特になし';
}

// 作業内容を個別のアクションに分割する関数
function splitIntoActions(workContent) {
  if (!workContent || typeof workContent !== 'string') {
    return [''];
  }
  
  // 複数の区切り文字で分割（句読点、改行、「・」など）
  const separators = [
    '。',      // 句点
    '\n',      // 改行
    '・',      // 中黒
    '、その後', // 順序を示す表現
    '、次に',   // 順序を示す表現
    '→',       // 矢印
    '①', '②', '③', '④', '⑤', // 番号付きリスト
    '1.', '2.', '3.', '4.', '5.', // 番号付きリスト
    '；'        // セミコロン
  ];
  
  let actions = [workContent];
  
  // 各区切り文字で分割を試みる
  for (const separator of separators) {
    let tempActions = [];
    for (const action of actions) {
      if (action.includes(separator)) {
        const parts = action.split(separator);
        tempActions.push(...parts);
      } else {
        tempActions.push(action);
      }
    }
    actions = tempActions;
  }
  
  // 「および」「また」「さらに」で始まる部分も分割
  let finalActions = [];
  for (const action of actions) {
    if (action.match(/^(および|また|さらに|そして)/)) {
      // 接続詞で始まる場合は独立したアクションとして扱う
      finalActions.push(action);
    } else if (action.includes('および') || action.includes('また')) {
      // 文中に接続詞がある場合も分割を検討
      const subParts = action.split(/(?=および|また)/);
      finalActions.push(...subParts);
    } else {
      finalActions.push(action);
    }
  }
  
  // 空白のみのアクションを除去し、トリミング
  finalActions = finalActions
    .map(action => action.trim())
    .filter(action => action.length > 0);
  
  // アクションが空の場合は元のテキストを返す
  if (finalActions.length === 0) {
    return [workContent];
  }
  
  // 各アクションに連番を付ける（オプション）
  const numbered = finalActions.map((action, index) => {
    // すでに番号が付いている場合はそのまま
    if (action.match(/^[①-⑩\d+\.]/)) {
      return action;
    }
    // 短いアクション（10文字以下）の場合は番号を付けない
    if (action.length <= 10) {
      return action;
    }
    // それ以外は番号を付ける
    return `${index + 1}. ${action}`;
  });
  
  console.log(`アクション分割結果: ${numbered.length}個`);
  numbered.forEach((action, i) => {
    console.log(`  アクション${i + 1}: ${action.substring(0, 50)}${action.length > 50 ? '...' : ''}`);
  });
  
  return numbered;
}

// flow_rowsデータのクリーニング関数
function cleanFlowRowsData(flowRows) {
  console.log('flow_rowsクリーニング開始');
  
  if (!flowRows) {
    console.log('flow_rowsがnullまたはundefined');
    return [];
  }
  
  // 配列でない場合は配列に変換
  if (!Array.isArray(flowRows)) {
    console.log('flow_rowsが配列ではないため変換');
    flowRows = [flowRows];
  }
  
  const cleaned = flowRows.map((row, index) => {
    console.log(`行${index + 1}のクリーニング開始`);
    
    // オブジェクトの場合はそのまま処理
    if (typeof row === 'object' && row !== null && !Array.isArray(row)) {
      const cleanedRow = {};
      for (const key in row) {
        let value = row[key];
        
        // 値をクリーニング
        if (typeof value === 'string') {
          // 末尾の数字を削除（「特になし3」→「特になし」）
          value = value.replace(/(\D+)\d+$/, '$1').trim();
          // 「なし」の末尾の数字も削除
          value = value.replace(/^(なし)\d+$/, '$1');
          value = value.replace(/^(特になし)\d+$/, '$1');
        }
        
        cleanedRow[key] = value || '';
      }
      console.log(`行${index + 1}クリーニング完了（オブジェクト）`);
      return cleanedRow;
    }
    
    // 文字列の場合
    if (typeof row === 'string') {
      console.log(`行${index + 1}は文字列: ${row.substring(0, 100)}`);
      // カンマ区切りで分割
      const parts = row.split(',').map(part => {
        let cleaned = part.trim();
        // 末尾の数字を削除
        cleaned = cleaned.replace(/(\D+)\d+$/, '$1').trim();
        cleaned = cleaned.replace(/^(なし)\d+$/, '$1');
        cleaned = cleaned.replace(/^(特になし)\d+$/, '$1');
        return cleaned;
      });
      
      // ヘッダーに基づいてオブジェクトを作成
      const headers = ['工程', '実施タイミング', '部署', '担当役割', '作業内容', '条件分岐', '利用ツール', 'URLリンク', '備考'];
      const cleanedRow = {};
      headers.forEach((header, i) => {
        cleanedRow[header] = parts[i] || '';
      });
      console.log(`行${index + 1}クリーニング完了（文字列→オブジェクト）`);
      return cleanedRow;
    }
    
    // 配列の場合
    if (Array.isArray(row)) {
      console.log(`行${index + 1}は配列`);
      const headers = ['工程', '実施タイミング', '部署', '担当役割', '作業内容', '条件分岐', '利用ツール', 'URLリンク', '備考'];
      const cleanedRow = {};
      headers.forEach((header, i) => {
        let value = row[i] || '';
        if (typeof value === 'string') {
          // 末尾の数字を削除
          value = value.replace(/(\D+)\d+$/, '$1').trim();
          value = value.replace(/^(なし)\d+$/, '$1');
          value = value.replace(/^(特になし)\d+$/, '$1');
        }
        cleanedRow[header] = value;
      });
      console.log(`行${index + 1}クリーニング完了（配列→オブジェクト）`);
      return cleanedRow;
    }
    
    console.warn(`行${index + 1}は未対応の型: ${typeof row}`);
    return null;
  }).filter(row => row !== null);
  
  console.log(`クリーニング完了: ${cleaned.length}行`);
  return cleaned;
}

// デバッグ用関数：データ型と内容を詳細に出力
function debugDataStructure(data, label = 'データ') {
  console.log('\n========== デバッグ情報開始 ==========');
  console.log(`【${label}】`);
  console.log('データ型:', typeof data);
  console.log('null/undefined?:', data === null || data === undefined);
  console.log('配列?:', Array.isArray(data));
  
  if (data === null || data === undefined) {
    console.log('データは null または undefined です');
    console.log('========== デバッグ情報終了 ==========\n');
    return;
  }
  
  if (Array.isArray(data)) {
    console.log('配列の長さ:', data.length);
    console.log('最初の3要素の詳細:');
    for (let i = 0; i < Math.min(3, data.length); i++) {
      console.log(`  [${i}] 型: ${typeof data[i]}`);
      if (typeof data[i] === 'string') {
        console.log(`      値: "${data[i].substring(0, 100)}${data[i].length > 100 ? '...' : ''}"`);
        console.log(`      長さ: ${data[i].length}文字`);
        console.log(`      カンマの数: ${(data[i].match(/,/g) || []).length}`);
      } else if (Array.isArray(data[i])) {
        console.log(`      配列長: ${data[i].length}`);
        console.log(`      内容: [${data[i].slice(0, 3).map(v => typeof v).join(', ')}...]`);
      } else if (typeof data[i] === 'object' && data[i] !== null) {
        console.log(`      キー: ${Object.keys(data[i]).slice(0, 5).join(', ')}`);
      } else {
        console.log(`      値: ${data[i]}`);
      }
    }
  } else if (typeof data === 'string') {
    console.log('文字列の長さ:', data.length);
    console.log('最初の200文字:', data.substring(0, 200) + (data.length > 200 ? '...' : ''));
    console.log('改行の数:', (data.match(/\n/g) || []).length);
    console.log('カンマの数:', (data.match(/,/g) || []).length);
    console.log('最初の行:', data.split('\n')[0]);
  } else if (typeof data === 'object') {
    const keys = Object.keys(data);
    console.log('オブジェクトのキー数:', keys.length);
    console.log('最初の10個のキー:', keys.slice(0, 10).join(', '));
    console.log('最初の3つのキーと値:');
    for (let i = 0; i < Math.min(3, keys.length); i++) {
      const key = keys[i];
      const value = data[key];
      console.log(`  ${key}: (${typeof value}) ${String(value).substring(0, 50)}${String(value).length > 50 ? '...' : ''}`);
    }
  } else {
    console.log('その他の型のデータ:', data);
  }
  
  console.log('========== デバッグ情報終了 ==========\n');
}

// 完全に安全な新しいフロー行書き込み関数（1アクション1セル形式）
function writeFlowRowsSafe(flowRows) {
  const sheetName = getConfig('FLOW_SHEET_NAME') || FLOW_SHEET;
  const sheet = ss().getSheetByName(sheetName) || createFlowSheet(sheetName);
  const headers = FLOW_HEADERS; // 定数を使用（法令・規制等を含む）

  console.log('=== 安全なフロー行書き込み開始（1アクション1セル形式） ===');
  
  // 詳細なデバッグ情報を出力
  debugDataStructure(flowRows, '入力データ (flowRows)');
  
  // データを安全に処理（1アクション1セル形式）
  let processedData = [];
  
  try {
    // flowRowsが配列かどうかチェック
    if (!flowRows) {
      console.log('データがnullまたはundefined');
      return;
    }
    
    if (Array.isArray(flowRows)) {
      console.log(`配列として受信: ${flowRows.length}個の要素`);
      
      // 各要素を安全に処理（作業内容を分割）
      for (let i = 0; i < flowRows.length; i++) {
        const row = flowRows[i];
        console.log(`\n--- 行${i + 1}の処理開始 ---`);
        debugDataStructure(row, `行${i + 1}`);
        
        if (typeof row === 'object' && row !== null) {
          // 作業内容を分割して複数行に展開（1アクション1セル）
          const workContent = row['作業内容'] || '';
          const actions = splitIntoActions(workContent);
          
          console.log(`作業内容を${actions.length}個のアクションに分割`);
          
          // 各アクションごとに行を作成（法令・内部統制の観点を追加）
          for (let j = 0; j < actions.length; j++) {
            const rowArray = [];
            const processName = row['工程'] || '';
            const timing = row['実施タイミング'] || '';
            const dept = row['部署'] || '';
            const condition = row['条件分岐'] || '';
            
            for (const header of headers) {
              let value = '';
              
              if (header === '作業内容') {
                // 作業内容は分割されたアクション
                value = actions[j];
              } else if (header === '法令・規制') {
                // 法令・規制を自動判定
                value = checkLegalRegulations(processName, actions[j], timing, dept);
              } else if (header === '内部統制') {
                // 内部統制の観点を自動判定
                value = checkInternalControl(processName, actions[j], condition, dept);
              } else if (header === 'コンプライアンス留意点') {
                // コンプライアンス留意点を自動生成
                value = j === 0 ? generateComplianceNotes(processName, actions[j], timing, dept, condition) : '';
              } else if (j === 0) {
                // 最初のアクションの場合は全ての情報を含める
                value = row[header] || '';
              } else {
                // 2番目以降のアクションは作業内容以外を空にするか、継続する情報のみ
                if (header === '工程' || header === '実施タイミング' || header === '部署' || header === '担当役割') {
                  value = row[header] || '';
                } else {
                  value = '';
                }
              }
              
              // 末尾の不要な数字を削除
              const cleanValue = String(value).replace(/特になし\d+$/, '特になし').replace(/なし\d+$/, 'なし');
              rowArray.push(cleanValue);
            }
            processedData.push(rowArray);
          }
        } else if (typeof row === 'string') {
          // 文字列の場合、カンマで分割して配列に変換
          console.log(`行${i + 1}は文字列です。解析を試みます`);
          console.log('文字列の内容（最初の100文字）:', row.substring(0, 100));
          const parts = row.split(',').map(part => part.trim());
          console.log('分割後の要素数:', parts.length);
          console.log('分割結果:', parts);
          const rowArray = [];
          for (let j = 0; j < headers.length; j++) {
            const value = parts[j] || '';
            const cleanValue = String(value).replace(/特になし\d+$/, '特になし').replace(/なし\d+$/, 'なし');
            rowArray.push(cleanValue);
          }
          processedData.push(rowArray);
        } else if (Array.isArray(row)) {
          // 既に配列の場合
          const rowArray = [];
          for (let j = 0; j < headers.length; j++) {
            const value = row[j] || '';
            const cleanValue = String(value).replace(/特になし\d+$/, '特になし').replace(/なし\d+$/, 'なし');
            rowArray.push(cleanValue);
          }
          processedData.push(rowArray);
        }
      }
    } else if (typeof flowRows === 'string') {
      // 全体が文字列の場合、行ごとに分割してから処理
      console.log('全体が文字列として受信');
      console.log('文字列の長さ:', flowRows.length);
      console.log('最初の200文字:', flowRows.substring(0, 200));
      const lines = flowRows.split('\n').filter(line => line.trim());
      console.log('行数:', lines.length);
      if (lines.length > 0) {
        console.log('最初の行:', lines[0]);
      }
      for (let i = 0; i < lines.length; i++) {
        const parts = lines[i].split(',').map(part => part.trim());
        const rowArray = [];
        for (let j = 0; j < headers.length; j++) {
          const value = parts[j] || '';
          const cleanValue = String(value).replace(/特になし\d+$/, '特になし').replace(/なし\d+$/, 'なし');
          rowArray.push(cleanValue);
        }
        processedData.push(rowArray);
      }
    } else if (typeof flowRows === 'object') {
      // 単一のオブジェクトの場合
      console.log('単一のオブジェクトとして受信');
      const rowArray = [];
      for (const header of headers) {
        const value = flowRows[header] || '';
        const cleanValue = String(value).replace(/特になし\d+$/, '特になし').replace(/なし\d+$/, 'なし');
        rowArray.push(cleanValue);
      }
      processedData.push(rowArray);
    } else {
      console.error('サポートされていないデータ型:', typeof flowRows);
      return;
    }
    
    // 処理済みデータのデバッグ情報
    console.log('\n=== 処理済みデータの確認 ===');
    console.log('処理済み行数:', processedData.length);
    if (processedData.length > 0) {
      console.log('最初の行のデータ:', processedData[0]);
    }
    
    // データが存在しない場合は終了
    if (processedData.length === 0) {
      console.log('処理可能なデータがありません');
      return;
    }
    
    // 既存データをクリア
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      const clearRows = lastRow - 1;
      const clearCols = headers.length;
      console.log(`既存データをクリア: ${clearRows}行 x ${clearCols}列`);
      sheet.getRange(2, 1, clearRows, clearCols).clearContent();
    }
    
    // セルごとに安全に書き込み（エラーを完全に回避）
    console.log(`書き込み開始: ${processedData.length}行`);
    for (let i = 0; i < processedData.length; i++) {
      for (let j = 0; j < headers.length; j++) {
        try {
          const cellValue = processedData[i][j] || '';
          // セルごとに個別に書き込み（最も安全）
          sheet.getRange(i + 2, j + 1).setValue(cellValue);
        } catch (cellError) {
          console.error(`セル(${i + 2}, ${j + 1})書き込みエラー:`, cellError.message);
          console.error('エラーの詳細:', cellError);
          console.error('書き込もうとした値:', processedData[i][j]);
          console.error('値の型:', typeof processedData[i][j]);
          // エラーが発生してもデフォルト値を設定
          try {
            sheet.getRange(i + 2, j + 1).setValue('');
          } catch (e) {
            // それでも失敗したら無視
          }
        }
      }
      console.log(`行${i + 1}書き込み完了`);
    }
    
    // 書式設定（エラーが発生しても続行）
    try {
      sheet.getRange(2, 1, processedData.length, headers.length).setWrap(true);
      sheet.getRange(2, 1, processedData.length, 1).setFontWeight('bold');
    } catch (e) {
      console.warn('書式設定エラー:', e.message);
    }
    
    console.log('=== フロー行書き込み完了 ===');
    logActivity('WRITE_FLOW_SAFE', `Successfully written ${processedData.length} flow rows`);
    
  } catch (error) {
    console.error('致命的なエラー:', error.message);
    console.error('スタックトレース:', error.stack);
    // エラーが発生してもアプリケーションは続行
  }
}

// ================================================================================
// 8. governance_functions.gs - ガバナンス・コンプライアンス機能
// ================================================================================

/**
 * ガバナンス機能設定
 */
const GOVERNANCE_CONFIG = {
  enableDisclosureCheck: true,
  enableAdvisorConsultation: true,
  enableMECEClassification: true,
  autoGenerateTimeline: true,
  strictComplianceMode: true
};

// 開示判定マスターデータ
const DISCLOSURE_TRIGGERS = {
  TIMELY_DISCLOSURE_DECISIONS: {
    '株式発行': {
      criteria: ['新株発行', '増資', '第三者割当', '公募', '株主割当'],
      threshold: '発行済株式総数の10%以上',
      timeline: '決議後直ちに',
      authority: '取締役会決議',
      documents: ['有価証券届出書', '適時開示資料'],
      regulations: ['金商法第4条', '東証適時開示規則第2条']
    },
    '資本政策': {
      criteria: ['自己株式取得', '資本金減少', '株式分割', '株式併合'],
      threshold: '資本金の10%以上',
      timeline: '決議後直ちに',
      authority: '取締役会決議（一部株主総会）',
      documents: ['適時開示資料', '臨時報告書'],
      regulations: ['会社法第156条', '東証適時開示規則']
    },
    'M&A': {
      criteria: ['合併', '会社分割', '株式交換', '株式移転', '事業譲渡'],
      threshold: '純資産の30%以上',
      timeline: '基本合意時及び決議後',
      authority: '取締役会決議及び株主総会特別決議',
      documents: ['適時開示資料', '臨時報告書', '公開買付届出書'],
      regulations: ['会社法第783条', '金商法第27条の3']
    },
    '業務提携': {
      criteria: ['資本提携', '業務提携', '技術提携'],
      threshold: '売上高の10%以上の影響',
      timeline: '契約締結後直ちに',
      authority: '取締役会決議',
      documents: ['適時開示資料'],
      regulations: ['東証適時開示規則第2条']
    }
  },
  TIMELY_DISCLOSURE_EVENTS: {
    '災害・事故': {
      criteria: ['火災', '爆発', '自然災害', '事故'],
      threshold: '純資産の3%以上の損害',
      timeline: '発生後直ちに',
      authority: '代表取締役',
      documents: ['適時開示資料'],
      regulations: ['東証適時開示規則第2条']
    },
    '訴訟': {
      criteria: ['訴訟提起', '仲裁申立', '調停申立'],
      threshold: '純資産の15%以上の請求',
      timeline: '提起後直ちに',
      authority: '法務部門',
      documents: ['適時開示資料'],
      regulations: ['東証適時開示規則']
    }
  },
  FINANCIAL_DISCLOSURE: {
    '決算短信': {
      criteria: ['四半期決算', '通期決算'],
      timeline: '決算後45日以内（推奨30日）',
      authority: '取締役会承認',
      documents: ['決算短信', '四半期報告書'],
      regulations: ['東証決算短信作成要領']
    },
    '業績予想修正': {
      criteria: ['売上高', '営業利益', '経常利益', '純利益', '配当'],
      threshold: '10%以上の乖離',
      timeline: '判明後直ちに',
      authority: '取締役会決議',
      documents: ['適時開示資料'],
      regulations: ['東証適時開示規則']
    }
  },
  STATUTORY_DISCLOSURE: {
    '有価証券報告書': {
      timeline: '事業年度終了後3か月以内',
      authority: '代表取締役',
      edinet: true,
      regulations: ['金商法第24条']
    },
    '四半期報告書': {
      timeline: '四半期終了後45日以内',
      authority: '代表取締役',
      edinet: true,
      regulations: ['金商法第24条の4の7']
    },
    '臨時報告書': {
      triggers: ['主要株主異動', '代表取締役異動', '監査人異動'],
      timeline: '発生後遅滞なく',
      authority: '代表取締役',
      edinet: true,
      regulations: ['金商法第24条の5']
    },
    '内部統制報告書': {
      timeline: '有価証券報告書と同時',
      authority: '代表取締役',
      edinet: true,
      regulations: ['金商法第24条の4の4']
    }
  }
};

// 外部専門家マスターデータ
const EXTERNAL_ADVISORS = {
  '法律事務所': {
    specialties: ['M&A・企業再編', 'コーポレートガバナンス', '株主総会対応', '取締役会運営', '契約法務', 'コンプライアンス'],
    consultationTiming: '重要な法的判断が必要な段階の初期',
    deliverables: ['リーガルオピニオン', '契約書レビュー', 'デューデリジェンス報告書'],
    urgencyLevels: { 'CRITICAL': '即日対応', 'HIGH': '2-3営業日', 'MEDIUM': '1週間', 'LOW': '2週間' }
  },
  '監査法人': {
    specialties: ['会計監査', '内部統制監査（J-SOX）', '四半期レビュー', 'M&Aデューデリジェンス'],
    consultationTiming: '決算前・重要な会計処理の変更前',
    deliverables: ['監査報告書', '内部統制監査報告書', 'マネジメントレター'],
    urgencyLevels: { 'CRITICAL': '即日対応', 'HIGH': '3-5営業日', 'MEDIUM': '2週間', 'LOW': '1か月' }
  },
  '税理士事務所': {
    specialties: ['税務申告', '税務調査対応', '移転価格', '国際税務', '組織再編税制'],
    consultationTiming: '税務判断が必要な取引の実行前',
    deliverables: ['税務意見書', '税務デューデリジェンス報告書', 'タックスストラクチャリング提案'],
    urgencyLevels: { 'CRITICAL': '即日対応', 'HIGH': '2-3営業日', 'MEDIUM': '1週間', 'LOW': '2週間' }
  },
  '司法書士事務所': {
    specialties: ['商業登記', '不動産登記', '役員変更登記', '定款変更', '増資・減資登記'],
    consultationTiming: '登記が必要な決議の前',
    deliverables: ['登記申請書', '定款', '議事録作成支援'],
    urgencyLevels: { 'CRITICAL': '当日対応', 'HIGH': '1-2営業日', 'MEDIUM': '3-5営業日', 'LOW': '1週間' }
  },
  '社会保険労務士事務所': {
    specialties: ['就業規則作成・変更', '労働契約', '労使協定', '社会保険手続き'],
    consultationTiming: '人事労務施策の実施前',
    deliverables: ['就業規則', '労使協定書', '労務監査報告書'],
    urgencyLevels: { 'CRITICAL': '即日対応', 'HIGH': '2-3営業日', 'MEDIUM': '1週間', 'LOW': '2週間' }
  }
};

/**
 * タスクが開示対象かを判定
 */
function checkDisclosureRequirement(task) {
  const result = {
    requiresDisclosure: false,
    disclosureType: [],
    timeline: [],
    authorities: [],
    documents: [],
    regulations: [],
    notes: []
  };

  const taskLower = task.toLowerCase();
  
  // 株主総会関連
  if (taskLower.includes('株主総会')) {
    result.requiresDisclosure = true;
    result.disclosureType.push('株主総会関連開示');
    
    if (taskLower.includes('定時')) {
      result.timeline.push('招集通知: 総会2週間前');
      result.timeline.push('招集通知Web開示: 発送前');
      result.documents.push('招集通知', '事業報告', '計算書類');
      result.regulations.push('会社法第299条');
    }
    
    if (taskLower.includes('臨時')) {
      result.timeline.push('適時開示: 招集決定後直ちに');
      result.documents.push('適時開示資料', '招集通知');
      result.regulations.push('東証適時開示規則');
    }
    
    result.authorities.push('取締役会', '代表取締役');
    result.notes.push('TDnetでの開示と自社Webサイトでの公表を並行実施');
  }

  // 取締役会関連
  if (taskLower.includes('取締役会')) {
    const boardItems = ['決算', '配当', '自己株式', '役員', '組織再編', '業務提携'];
    for (const item of boardItems) {
      if (taskLower.includes(item)) {
        result.requiresDisclosure = true;
        result.disclosureType.push('取締役会決議事項');
        result.timeline.push('決議後直ちに');
        result.authorities.push('取締役会');
        result.documents.push('適時開示資料');
        result.regulations.push('東証適時開示規則第2条');
        break;
      }
    }
  }

  // M&A・組織再編
  if (taskLower.includes('合併') || taskLower.includes('買収') || 
      taskLower.includes('m&a') || taskLower.includes('事業譲渡')) {
    result.requiresDisclosure = true;
    result.disclosureType.push('組織再編・M&A');
    result.timeline.push('基本合意時', '最終契約時', '効力発生時');
    result.authorities.push('取締役会', '株主総会（特別決議）');
    result.documents.push('適時開示資料', '臨時報告書', '公開買付届出書');
    result.regulations.push('金商法第27条の3', '会社法第783条');
    result.notes.push('財務アドバイザー・法務アドバイザーとの連携必須');
  }

  // 決算・業績関連
  if (taskLower.includes('決算') || taskLower.includes('業績')) {
    result.requiresDisclosure = true;
    result.disclosureType.push('決算開示');
    
    if (taskLower.includes('四半期')) {
      result.timeline.push('四半期終了後45日以内');
      result.documents.push('四半期決算短信', '四半期報告書');
      result.regulations.push('金商法第24条の4の7');
    } else if (taskLower.includes('通期') || taskLower.includes('年度')) {
      result.timeline.push('期末後45日以内（決算短信）', '期末後3か月以内（有価証券報告書）');
      result.documents.push('決算短信', '有価証券報告書');
      result.regulations.push('金商法第24条');
    }
    
    if (taskLower.includes('修正') || taskLower.includes('予想')) {
      result.timeline.push('判明後直ちに（業績予想修正）');
      result.notes.push('売上高・利益が10%以上乖離する場合は開示必須');
    }
    
    result.authorities.push('取締役会', '監査役会', '会計監査人');
  }

  return result;
}

/**
 * タスクに対して必要な外部専門家を判定
 */
function determineRequiredAdvisors(task, context = {}) {
  const requiredAdvisors = [];
  const taskLower = task.toLowerCase();
  
  // 必須相談パターンの定義
  const mandatoryPatterns = [
    {
      keywords: ['株主総会', '株主', '総会'],
      advisors: ['法律事務所', '司法書士事務所'],
      reason: '株主総会の適法な運営と手続きの確認',
      checkpoints: ['招集手続きの適法性確認', '議案の適法性確認', '決議要件の確認', '議事録作成要領の確認']
    },
    {
      keywords: ['取締役会', '取締役', '役員', '執行役'],
      advisors: ['法律事務所', '司法書士事務所'],
      reason: '取締役会運営の適法性と役員変更登記',
      checkpoints: ['決議事項の適法性確認', '利益相反取引の確認', '特別利害関係の確認', '登記手続きの確認']
    },
    {
      keywords: ['M&A', '買収', '合併', '事業譲渡', '会社分割'],
      advisors: ['法律事務所', '監査法人', '税理士事務所'],
      reason: 'M&A取引の法務・財務・税務面での総合的検証',
      checkpoints: ['ストラクチャーの検討', 'デューデリジェンスの実施', '価格の妥当性検証', '契約条件の交渉']
    },
    {
      keywords: ['決算', '財務諸表', '有価証券報告書', '四半期報告書'],
      advisors: ['監査法人', '税理士事務所'],
      reason: '適正な財務報告と税務申告',
      checkpoints: ['会計処理の妥当性確認', '開示内容の適切性確認', '内部統制の有効性評価', '税務リスクの確認']
    },
    {
      keywords: ['増資', '減資', '自己株式', '新株', '社債'],
      advisors: ['法律事務所', '司法書士事務所'],
      reason: '資本政策の適法性と実行可能性の確認',
      checkpoints: ['発行条件の妥当性', '既存株主への影響分析', '開示書類の作成', '登記手続きの準備']
    },
    {
      keywords: ['労働', '雇用', '解雇', '就業規則', 'ハラスメント'],
      advisors: ['社会保険労務士事務所', '法律事務所'],
      reason: '労働法令遵守と労使紛争の予防',
      checkpoints: ['労働法令の遵守確認', '就業規則の整備', '労使協定の締結', '紛争リスクの評価']
    },
    {
      keywords: ['契約', '締結', '変更', '解除'],
      advisors: ['法律事務所'],
      reason: '契約リスクの評価と条件交渉',
      checkpoints: ['契約条件の妥当性確認', 'リスク条項の確認', '責任範囲の明確化', '紛争解決条項の確認']
    },
    {
      keywords: ['コンプライアンス', '違反', '不正', '内部統制'],
      advisors: ['法律事務所', '監査法人'],
      reason: 'コンプライアンス体制の強化と違反防止',
      checkpoints: ['現状のリスク評価', '改善策の立案', 'モニタリング体制の構築', '教育研修の実施']
    },
    {
      keywords: ['訴訟', '紛争', '係争', '調停', '仲裁'],
      advisors: ['法律事務所'],
      reason: '法的紛争の適切な解決',
      checkpoints: ['勝訴可能性の評価', '和解条件の検討', '証拠の収集・保全', '訴訟戦略の立案']
    },
    {
      keywords: ['個人情報', 'プライバシー', 'GDPR', '情報漏洩'],
      advisors: ['法律事務所'],
      reason: '個人情報保護法令の遵守',
      checkpoints: ['現行体制の評価', '規程・手順の整備', 'セキュリティ対策の確認', 'インシデント対応体制の構築']
    }
  ];

  // パターンマッチング
  mandatoryPatterns.forEach(pattern => {
    const hasKeyword = pattern.keywords.some(keyword => taskLower.includes(keyword));
    if (hasKeyword) {
      pattern.advisors.forEach(advisor => {
        requiredAdvisors.push({
          type: advisor,
          reason: pattern.reason,
          priority: 'MANDATORY',
          checkpoints: pattern.checkpoints,
          timing: EXTERNAL_ADVISORS[advisor].consultationTiming
        });
      });
    }
  });

  // 金額基準での判定
  if (context.amount) {
    const amount = parseInt(context.amount);
    if (amount > 100000000) { // 1億円以上
      requiredAdvisors.push({
        type: '法律事務所',
        reason: '高額取引のため法的リスク評価が必要',
        priority: 'HIGH',
        checkpoints: ['契約条件の精査', 'リスク分析', '交渉戦略']
      });
    }
  }

  return requiredAdvisors;
}

/**
 * 外部専門家相談チェックリスト生成
 */
function generateConsultationChecklist(task, advisors) {
  const checklist = {
    task: task,
    consultationSteps: [],
    documentationRequired: [],
    timeline: [],
    budgetConsiderations: []
  };

  // ステップ1: 事前準備
  checklist.consultationSteps.push({
    step: 1,
    phase: '事前準備',
    actions: [
      '相談事項の明確化と論点整理',
      '関連資料の収集と整理',
      '社内での事前検討と方針案の作成',
      '予算の確保と決裁取得'
    ],
    timeline: 'T-14日',
    responsible: '担当部門'
  });

  // ステップ2: 専門家選定
  checklist.consultationSteps.push({
    step: 2,
    phase: '専門家選定',
    actions: [
      '複数の専門家候補のリストアップ',
      '見積もり取得と比較検討',
      '利益相反チェック',
      '秘密保持契約（NDA）の締結'
    ],
    timeline: 'T-10日',
    responsible: '法務部・総務部'
  });

  // ステップ3: 各専門家への相談実施
  advisors.forEach((advisor, index) => {
    checklist.consultationSteps.push({
      step: 3 + index,
      phase: `${advisor.type}への相談`,
      actions: [
        '初回ミーティングの実施',
        '詳細情報の提供と質疑応答',
        '中間報告の受領とフィードバック',
        '最終意見書・報告書の受領'
      ],
      timeline: `T-${7 - index}日`,
      responsible: `担当部門・${advisor.type}`,
      deliverables: EXTERNAL_ADVISORS[advisor.type].deliverables,
      checkpoints: advisor.checkpoints || []
    });
  });

  // ステップ4: 社内検討
  checklist.consultationSteps.push({
    step: 3 + advisors.length,
    phase: '社内検討・意思決定',
    actions: [
      '専門家意見の社内共有と検討',
      'リスク評価と対応策の決定',
      '実行計画の策定',
      '必要な社内承認の取得'
    ],
    timeline: 'T-2日',
    responsible: '経営陣・関連部門'
  });

  // 必要書類リスト
  checklist.documentationRequired = [
    '相談依頼書',
    '背景説明資料',
    '関連契約書・規程類',
    '財務データ（必要に応じて）',
    '過去の類似案件資料',
    '社内検討資料',
    '取締役会・経営会議資料'
  ];

  return checklist;
}

/**
 * MECEなタスク分類体系
 */
const TASK_CLASSIFICATION_MECE = {
  'ガバナンス・コンプライアンス': {
    '株主総会運営': {
      tasks: ['定時株主総会の準備・開催', '臨時株主総会の準備・開催', '株主総会招集通知の作成・送付'],
      disclosure: true,
      priority: 'HIGH'
    },
    '取締役会運営': {
      tasks: ['取締役会の開催・運営', '取締役会議事録の作成', '取締役会規程の管理'],
      disclosure: true,
      priority: 'HIGH'
    },
    '監査対応': {
      tasks: ['監査役監査への対応', '内部監査への対応', '会計監査人監査への対応'],
      disclosure: false,
      priority: 'HIGH'
    },
    'コンプライアンス管理': {
      tasks: ['コンプライアンス違反の防止・発見', '内部通報制度の運営', 'コンプライアンス研修の実施'],
      disclosure: false,
      priority: 'MEDIUM'
    }
  },
  '情報開示・IR': {
    '適時開示': {
      tasks: ['決定事実の開示', '発生事実の開示', '決算情報の開示', '業績予想修正の開示'],
      disclosure: true,
      priority: 'CRITICAL'
    },
    '法定開示': {
      tasks: ['有価証券報告書の作成・提出', '四半期報告書の作成・提出', '臨時報告書の作成・提出'],
      disclosure: true,
      priority: 'CRITICAL'
    },
    'IR活動': {
      tasks: ['決算説明会の開催', 'アナリスト・機関投資家対応', '個人投資家向け説明会'],
      disclosure: false,
      priority: 'HIGH'
    }
  },
  '内部統制・リスク管理': {
    '内部統制システム': {
      tasks: ['J-SOX対応', '内部統制の整備・運用', '内部統制の評価', '内部統制報告書の作成'],
      disclosure: true,
      priority: 'HIGH'
    },
    'リスク管理': {
      tasks: ['リスクアセスメント', 'リスク対応策の策定', 'BCP（事業継続計画）の策定・更新'],
      disclosure: false,
      priority: 'HIGH'
    }
  },
  '経営管理': {
    '経営企画': {
      tasks: ['中期経営計画の策定', '年度事業計画の策定', '予算策定・管理', 'KPI管理'],
      disclosure: false,
      priority: 'HIGH'
    },
    '組織管理': {
      tasks: ['組織変更・改編', '規程・規則の制定・改廃', '権限委譲・決裁権限の管理'],
      disclosure: false,
      priority: 'MEDIUM'
    }
  }
};

/**
 * タスクをMECE分類に振り分け
 */
function classifyTaskMECE(task) {
  const classification = {
    level1: null,
    level2: null,
    level3: null,
    requiresDisclosure: false,
    priority: 'LOW',
    relatedTasks: []
  };

  const taskLower = task.toLowerCase();

  // 各分類を検査
  for (const [l1Key, l1Value] of Object.entries(TASK_CLASSIFICATION_MECE)) {
    for (const [l2Key, l2Value] of Object.entries(l1Value)) {
      for (const l3Task of l2Value.tasks) {
        if (taskLower.includes(l3Task.toLowerCase()) || 
            l3Task.toLowerCase().includes(taskLower)) {
          classification.level1 = l1Key;
          classification.level2 = l2Key;
          classification.level3 = l3Task;
          classification.requiresDisclosure = l2Value.disclosure;
          classification.priority = l2Value.priority;
          classification.relatedTasks = l2Value.tasks.filter(t => t !== l3Task);
          return classification;
        }
      }
    }
  }

  // マッチしない場合はキーワードベースで推定
  if (taskLower.includes('開示') || taskLower.includes('報告書')) {
    classification.level1 = '情報開示・IR';
    classification.requiresDisclosure = true;
    classification.priority = 'HIGH';
  } else if (taskLower.includes('監査') || taskLower.includes('統制')) {
    classification.level1 = '内部統制・リスク管理';
    classification.priority = 'HIGH';
  } else if (taskLower.includes('取締役') || taskLower.includes('株主')) {
    classification.level1 = 'ガバナンス・コンプライアンス';
    classification.requiresDisclosure = true;
    classification.priority = 'HIGH';
  }

  return classification;
}

/**
 * 統合的なガバナンスチェック
 */
function performComprehensiveGovernanceCheck(workSpec, flowData) {
  const governanceReport = {
    overallScore: 0,
    disclosureRequirements: [],
    advisorConsultations: [],
    complianceGaps: [],
    recommendations: [],
    timeline: [],
    riskAssessment: []
  };

  // 1. 業務仕様書からガバナンス要素を抽出
  if (workSpec) {
    const specText = JSON.stringify(workSpec).toLowerCase();
    const disclosureCheck = checkDisclosureRequirement(specText);
    if (disclosureCheck.requiresDisclosure) {
      governanceReport.disclosureRequirements.push({
        type: disclosureCheck.disclosureType.join(', '),
        timeline: disclosureCheck.timeline,
        documents: disclosureCheck.documents,
        regulations: disclosureCheck.regulations
      });
    }
  }

  // 2. フローデータから承認プロセスを分析
  if (flowData && Array.isArray(flowData)) {
    const approvalSteps = flowData.filter(row => 
      row['作業内容'] && (
        row['作業内容'].includes('承認') ||
        row['作業内容'].includes('決裁') ||
        row['作業内容'].includes('決議')
      )
    );

    // 承認階層の適切性を評価
    const requiredApprovers = new Set();
    approvalSteps.forEach(step => {
      if (step['担当役割']) {
        requiredApprovers.add(step['担当役割']);
      }
    });

    // 必要な承認者が不足していないかチェック
    const essentialApprovers = ['取締役会', '代表取締役', '監査役'];
    essentialApprovers.forEach(approver => {
      if (!Array.from(requiredApprovers).some(r => r.includes(approver))) {
        if (governanceReport.disclosureRequirements.length > 0) {
          governanceReport.complianceGaps.push(
            `重要な承認者「${approver}」が承認フローに含まれていません`
          );
        }
      }
    });
  }

  // 3. リスク評価
  const risks = [
    {
      category: '開示遅延リスク',
      probability: governanceReport.disclosureRequirements.length > 2 ? 'HIGH' : 'MEDIUM',
      impact: 'HIGH',
      mitigation: 'IR部門との事前調整、開示チェックリストの活用'
    },
    {
      category: 'コンプライアンス違反リスク',
      probability: governanceReport.complianceGaps.length > 0 ? 'HIGH' : 'LOW',
      impact: 'CRITICAL',
      mitigation: '法務部門による事前レビュー、コンプライアンスチェックの実施'
    }
  ];
  governanceReport.riskAssessment = risks;

  // 4. 推奨事項の生成
  if (governanceReport.disclosureRequirements.length > 0) {
    governanceReport.recommendations.push(
      '東証への事前相談を検討してください（複雑な開示案件の場合）'
    );
    governanceReport.recommendations.push(
      'IR部門と法務部門の早期巻き込みを推奨します'
    );
  }

  if (governanceReport.complianceGaps.length > 0) {
    governanceReport.recommendations.push(
      '承認フローの見直しと必要な承認者の追加を検討してください'
    );
  }

  // 5. フローデータから専門家相談要件を抽出
  if (flowData && Array.isArray(flowData)) {
    flowData.forEach((row, index) => {
      const taskDescription = `${row['工程'] || ''} ${row['作業内容'] || ''}`;
      const advisors = determineRequiredAdvisors(taskDescription);
      
      if (advisors.length > 0) {
        const checklist = generateConsultationChecklist(taskDescription, advisors);
        governanceReport.advisorConsultations.push({
          taskId: index + 1,
          task: taskDescription,
          advisors: advisors,
          checklist: checklist
        });
      }
    });
  }

  // 6. スコアリング（100点満点）
  let score = 100;
  score -= governanceReport.complianceGaps.length * 10;
  score -= governanceReport.riskAssessment.filter(r => r.probability === 'HIGH').length * 5;
  score = Math.max(0, score);
  governanceReport.overallScore = score;

  return governanceReport;
}

// ガバナンスレポートシートを作成
function createGovernanceReportSheet(sheet, governanceCheck) {
  let row = 1;

  // タイトル
  sheet.getRange(row, 1, 1, 8).merge();
  sheet.getRange(row, 1).setValue('ガバナンス・コンプライアンスチェックレポート');
  sheet.getRange(row, 1).setFontSize(16).setFontWeight('bold');
  sheet.getRange(row, 1).setBackground('#1a73e8').setFontColor('#ffffff');
  row += 2;

  // サマリー
  sheet.getRange(row, 1).setValue('【サマリー】');
  sheet.getRange(row, 1).setFontWeight('bold').setBackground('#e8f0fe');
  row++;
  
  const summaryData = [
    ['ガバナンススコア', governanceCheck.overallScore + '/100点'],
    ['開示要件', governanceCheck.disclosureRequirements.length + '件'],
    ['外部専門家相談', (governanceCheck.advisorConsultations ? governanceCheck.advisorConsultations.length : 0) + '件'],
    ['コンプライアンスギャップ', governanceCheck.complianceGaps.length + '件']
  ];
  
  sheet.getRange(row, 1, summaryData.length, 2).setValues(summaryData);
  sheet.getRange(row, 1, summaryData.length, 1).setFontWeight('bold');
  row += summaryData.length + 2;

  // 開示要件
  if (governanceCheck.disclosureRequirements.length > 0) {
    sheet.getRange(row, 1).setValue('【開示要件】');
    sheet.getRange(row, 1).setFontWeight('bold').setBackground('#fce8b2');
    row++;
    
    const disclosureHeaders = ['No.', '開示種別', '期限', '必要書類', '関連法規'];
    sheet.getRange(row, 1, 1, disclosureHeaders.length).setValues([disclosureHeaders]);
    sheet.getRange(row, 1, 1, disclosureHeaders.length).setFontWeight('bold');
    row++;
    
    governanceCheck.disclosureRequirements.forEach((req, index) => {
      const rowData = [
        index + 1,
        req.type || '',
        Array.isArray(req.timeline) ? req.timeline.join(', ') : '',
        Array.isArray(req.documents) ? req.documents.join(', ') : '',
        Array.isArray(req.regulations) ? req.regulations.join(', ') : ''
      ];
      sheet.getRange(row, 1, 1, rowData.length).setValues([rowData]);
      row++;
    });
    row++;
  }

  // 外部専門家相談
  if (governanceCheck.advisorConsultations && governanceCheck.advisorConsultations.length > 0) {
    sheet.getRange(row, 1).setValue('【外部専門家相談が必要なタスク】');
    sheet.getRange(row, 1).setFontWeight('bold').setBackground('#d9ead3');
    row++;
    
    governanceCheck.advisorConsultations.forEach((consultation, index) => {
      sheet.getRange(row, 1).setValue(`${index + 1}. ${consultation.task}`);
      sheet.getRange(row, 1).setFontWeight('bold');
      row++;
      
      consultation.advisors.forEach(advisor => {
        sheet.getRange(row, 2).setValue(`・${advisor.type}: ${advisor.reason}`);
        row++;
      });
      row++;
    });
  }

  // 推奨事項
  if (governanceCheck.recommendations.length > 0) {
    sheet.getRange(row, 1).setValue('【推奨事項】');
    sheet.getRange(row, 1).setFontWeight('bold').setBackground('#f4cccc');
    row++;
    
    governanceCheck.recommendations.forEach(rec => {
      sheet.getRange(row, 1).setValue(`・${rec}`);
      row++;
    });
    row++;
  }

  // リスク評価
  if (governanceCheck.riskAssessment && governanceCheck.riskAssessment.length > 0) {
    sheet.getRange(row, 1).setValue('【リスク評価】');
    sheet.getRange(row, 1).setFontWeight('bold').setBackground('#ffe599');
    row++;
    
    const riskHeaders = ['リスクカテゴリ', '発生確率', '影響度', '対策'];
    sheet.getRange(row, 1, 1, riskHeaders.length).setValues([riskHeaders]);
    sheet.getRange(row, 1, 1, riskHeaders.length).setFontWeight('bold');
    row++;
    
    governanceCheck.riskAssessment.forEach(risk => {
      const rowData = [
        risk.category,
        risk.probability,
        risk.impact,
        risk.mitigation
      ];
      sheet.getRange(row, 1, 1, rowData.length).setValues([rowData]);
      row++;
    });
  }

  // 書式調整
  sheet.autoResizeColumns(1, 8);
}

// 専門家相談チェックリストシートを作成
function createConsultationChecklistSheet(sheet, consultations) {
  let row = 1;

  // タイトル
  sheet.getRange(row, 1, 1, 6).merge();
  sheet.getRange(row, 1).setValue('外部専門家相談チェックリスト');
  sheet.getRange(row, 1).setFontSize(16).setFontWeight('bold');
  sheet.getRange(row, 1).setBackground('#34a853').setFontColor('#ffffff');
  row += 2;

  consultations.forEach((consultation, consultIndex) => {
    // タスクタイトル
    sheet.getRange(row, 1, 1, 6).merge();
    sheet.getRange(row, 1).setValue(`【タスク${consultIndex + 1}】 ${consultation.task}`);
    sheet.getRange(row, 1).setFontWeight('bold').setBackground('#e8f5e9');
    row++;

    // 必要な専門家
    sheet.getRange(row, 1).setValue('必要な専門家:');
    sheet.getRange(row, 1).setFontWeight('bold');
    row++;
    
    consultation.advisors.forEach(advisor => {
      sheet.getRange(row, 2).setValue(`${advisor.type}`);
      sheet.getRange(row, 3).setValue(`理由: ${advisor.reason}`);
      sheet.getRange(row, 4).setValue(`優先度: ${advisor.priority}`);
      row++;
    });
    row++;

    // 相談ステップ
    if (consultation.checklist && consultation.checklist.consultationSteps) {
      sheet.getRange(row, 1).setValue('相談ステップ:');
      sheet.getRange(row, 1).setFontWeight('bold');
      row++;
      
      const stepHeaders = ['ステップ', 'フェーズ', 'アクション', 'タイミング', '担当'];
      sheet.getRange(row, 1, 1, stepHeaders.length).setValues([stepHeaders]);
      sheet.getRange(row, 1, 1, stepHeaders.length).setBackground('#f0f0f0').setFontWeight('bold');
      row++;
      
      consultation.checklist.consultationSteps.forEach(step => {
        const stepData = [
          step.step,
          step.phase,
          step.actions.join('\n'),
          step.timeline,
          step.responsible
        ];
        sheet.getRange(row, 1, 1, stepData.length).setValues([stepData]);
        sheet.getRange(row, 3).setWrap(true);
        row++;
      });
      row += 2;
    }
  });

  // 必要書類リスト
  sheet.getRange(row, 1).setValue('【準備が必要な書類】');
  sheet.getRange(row, 1).setFontWeight('bold').setBackground('#fce8b2');
  row++;
  
  const documents = [
    '相談依頼書',
    '背景説明資料',
    '関連契約書・規程類',
    '財務データ（必要に応じて）',
    '過去の類似案件資料',
    '社内検討資料',
    '取締役会・経営会議資料'
  ];
  
  documents.forEach(doc => {
    sheet.getRange(row, 1).setValue(`□ ${doc}`);
    row++;
  });

  // 書式調整
  sheet.autoResizeColumns(1, 6);
}

// ガバナンス情報を追加したフローシートを作成
function writeFlowToSheetWithGovernance(sheet, flowRows, governanceCheck) {
  const headers = FLOW_HEADERS; // 定数を使用して一貫性を保つ
  
  // ガバナンス情報を追加したヘッダー
  const enhancedHeaders = [...headers, '開示要件', '要専門家相談', '優先度'];
  
  sheet.getRange(1, 1, 1, enhancedHeaders.length).setValues([enhancedHeaders]);
  sheet.getRange(1, 1, 1, enhancedHeaders.length).setFontWeight('bold').setBackground('#E8F5E9');
  sheet.setFrozenRows(1);
  
  // フローデータを処理して書き込み
  const processedData = [];
  
  if (Array.isArray(flowRows)) {
    for (let i = 0; i < flowRows.length; i++) {
      const row = flowRows[i];
      
      if (typeof row === 'object' && row !== null) {
        const workContent = row['作業内容'] || '';
        const actions = splitIntoActions(workContent);
        const processName = row['工程'] || '';
        const timing = row['実施タイミング'] || '';
        const dept = row['部署'] || '';
        const condition = row['条件分岐'] || '';
        
        // ガバナンス情報を取得
        const taskDescription = `${processName} ${workContent}`;
        const disclosureCheck = checkDisclosureRequirement(taskDescription);
        const advisors = determineRequiredAdvisors(taskDescription);
        const classification = classifyTaskMECE(taskDescription);
        
        for (let j = 0; j < actions.length; j++) {
          const rowArray = [];
          
          for (const header of headers) {
            let value = '';
            
            if (header === '作業内容') {
              value = actions[j];
            } else if (header === '法令・規制') {
              value = checkLegalRegulations(processName, actions[j], timing, dept);
            } else if (header === '内部統制の観点') {
              value = checkInternalControl(processName, actions[j], condition, dept);
            } else if (header === 'コンプライアンス留意点') {
              value = generateComplianceNotes(processName, actions[j], timing, dept, condition);
            } else if (row.hasOwnProperty(header)) {
              value = j === 0 ? row[header] : '';
            } else {
              value = '';
            }
            
            rowArray.push(value || '');
          }
          
          // ガバナンス情報を追加
          rowArray.push(disclosureCheck.requiresDisclosure ? '要開示' : '');
          rowArray.push(advisors.length > 0 ? advisors.map(a => a.type).join(', ') : '');
          rowArray.push(classification.priority || '');
          
          processedData.push(rowArray);
        }
      }
    }
    
    // データを書き込み
    if (processedData.length > 0) {
      sheet.getRange(2, 1, processedData.length, enhancedHeaders.length).setValues(processedData);
      sheet.getRange(2, 1, processedData.length, enhancedHeaders.length).setWrap(true);
      
      // 優先度による色分け
      for (let i = 0; i < processedData.length; i++) {
        const priority = processedData[i][enhancedHeaders.length - 1];
        let bgColor = '#ffffff';
        
        switch(priority) {
          case 'CRITICAL': bgColor = '#f4cccc'; break;
          case 'HIGH': bgColor = '#fce5cd'; break;
          case 'MEDIUM': bgColor = '#fff2cc'; break;
          case 'LOW': bgColor = '#d9ead3'; break;
        }
        
        if (priority) {
          sheet.getRange(i + 2, enhancedHeaders.length).setBackground(bgColor);
        }
      }
    }
  }
  
  console.log('ガバナンス情報付きフロー書き込み完了');
}

// ========= モデル設定関数 =========

/**
 * OpenAIモデルを設定
 */
function setOpenAIModel(modelName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName('Config');
  
  if (!configSheet) {
    throw new Error('Configシートが見つかりません');
  }
  
  // OPENAI_MODEL行を探す
  const data = configSheet.getDataRange().getValues();
  let modelRowIndex = -1;
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === 'OPENAI_MODEL') {
      modelRowIndex = i + 1;
      break;
    }
  }
  
  if (modelRowIndex === -1) {
    // 新規追加
    const lastRow = configSheet.getLastRow();
    configSheet.getRange(lastRow + 1, 1, 1, 3).setValues([['OPENAI_MODEL', modelName, 'OpenAIモデル名']]);
  } else {
    // 既存更新
    configSheet.getRange(modelRowIndex, 2).setValue(modelName);
  }
  
  // スクリプトプロパティにも保存
  PropertiesService.getScriptProperties().setProperty('OPENAI_MODEL', modelName);
  
  SpreadsheetApp.getUi().alert(`OpenAIモデルを「${modelName}」に設定しました。`);
}


/**
 * GPT-4oモデルに切り替え
 */
function useGPT4o() {
  setOpenAIModel('gpt-5');
}

/**
 * OpenAIモデルをo3に切り替え
 */
function useO3() {
  setOpenAIModel('gpt-5');
  const ui = SpreadsheetApp.getUi();
  ui.alert('成功', 'OpenAIモデルを gpt-5 に切り替えました', ui.ButtonSet.OK);
}

/**
 * OpenAIモデルをgpt-5に切り替え
 */
function useO3DeepResearch() {
  setOpenAIModel('gpt-5');
  const ui = SpreadsheetApp.getUi();
  ui.alert('成功', 'OpenAIモデルを gpt-5 に切り替えました', ui.ButtonSet.OK);
}

/**
 * 現在のモデルを表示
 */
function showCurrentModel() {
  const currentModel = getConfig('OPENAI_MODEL') || 'gpt-5';
  const ui = SpreadsheetApp.getUi();
  
  let message = `現在のOpenAIモデル: ${currentModel}\n\n`;
  
  // モデルの特徴を説明
  const modelFeatures = {
    'gpt-5': '次世代モデル - 高精度な推論と構造化出力、Responses API最適化',
    'gpt-4o': '従来の高性能モデル - 高度な推論、深層分析、マルチモーダル対応'
  };
  
  if (modelFeatures[currentModel]) {
    message += `特徴: ${modelFeatures[currentModel]}`;
  }
  
  ui.alert('モデル情報', message);
}

/**
 * Slack API接続をテスト
 */
function testSlackConnection() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    // トークンの確認
    if (!SLACK_BOT_TOKEN) {
      ui.alert('エラー', 'Slack BOTトークンが設定されていません。\nスクリプトプロパティにSLACK_BOT_TOKENを設定してください。', ui.ButtonSet.OK);
      return;
    }
    
    // トークンの形式確認（xoxb-で始まるはず）
    if (!SLACK_BOT_TOKEN.startsWith('xoxb-')) {
      ui.alert('警告', 'Slack BOTトークンの形式が正しくない可能性があります。\nBotトークンは通常「xoxb-」で始まります。', ui.ButtonSet.OK);
    }
    
    // API接続テスト
    const response = slackAPI('auth.test', {});
    
    if (response.ok) {
      ui.alert(
        '接続成功',
        `Slack API接続テスト成功！\n\n` +
        `Bot名: ${response.user || 'N/A'}\n` +
        `チーム: ${response.team || 'N/A'}\n` +
        `URL: ${response.url || 'N/A'}`,
        ui.ButtonSet.OK
      );
    }
  } catch (error) {
    ui.alert('エラー', `Slack API接続エラー:\n${error.toString()}`, ui.ButtonSet.OK);
  }
}

/**
 * Slack API デバッグ情報を表示
 */
function debugSlackAPI() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const configSheet = ss.getSheetByName(SHEETS.CONFIG);
  const config = getConfigData(configSheet);
  
  let debugInfo = '=== Slack API デバッグ情報 ===\n\n';
  
  // トークン情報（一部マスク）
  if (SLACK_BOT_TOKEN) {
    const maskedToken = SLACK_BOT_TOKEN.substring(0, 10) + '...' + SLACK_BOT_TOKEN.substring(SLACK_BOT_TOKEN.length - 4);
    debugInfo += `Bot Token: ${maskedToken}\n`;
    debugInfo += `Token形式: ${SLACK_BOT_TOKEN.startsWith('xoxb-') ? '✅ OK (Bot Token)' : '❌ 異常'}\n`;
  } else {
    debugInfo += 'Bot Token: ❌ 未設定\n';
  }
  
  debugInfo += '\n--- 設定情報 ---\n';
  debugInfo += `監視対象チャンネル: ${config.targetChannels ? config.targetChannels.join(', ') : '未設定'}\n`;
  debugInfo += `通知先Slackチャンネル: ${config.notifySlackChannel || '未設定'}\n`;
  
  // 最近のエラーログを取得
  const logsSheet = ss.getSheetByName(SHEETS.LOGS);
  if (logsSheet && logsSheet.getLastRow() > 1) {
    debugInfo += '\n--- 最近のエラー ---\n';
    const recentLogs = logsSheet.getRange(Math.max(2, logsSheet.getLastRow() - 4), 1, 5, 3).getValues();
    recentLogs.forEach(log => {
      if (log[1] === 'ERROR') {
        debugInfo += `${log[0]}: ${log[2]}\n`;
      }
    });
  }
  
  ui.alert('デバッグ情報', debugInfo, ui.ButtonSet.OK);
}

// ========= テスト関数 =========

/**
 * Slack API動作テスト（slackAPI関数を正しく呼び出すサンプル）
 */
function testSlackAPICall() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    // 正しいslackAPI呼び出しの例
    console.log('Slack API接続テストを開始します...');
    
    // auth.testで接続確認
    const authResponse = slackAPI('auth.test', {});
    if (authResponse.ok) {
      console.log('✅ 認証成功:', authResponse);
      ui.alert('成功', `Slack接続成功\nBot名: ${authResponse.bot_name}\nチーム: ${authResponse.team}`, ui.ButtonSet.OK);
    }
    
    // チャンネル一覧取得の例
    console.log('\nチャンネル一覧を取得中...');
    const channelsResponse = slackAPI('conversations.list', {
      types: 'public_channel',
      limit: 10
    });
    
    if (channelsResponse.ok) {
      const channels = channelsResponse.channels.map(ch => ch.name).join(', ');
      console.log('✅ チャンネル取得成功:', channels);
    }
    
  } catch (error) {
    console.error('❌ エラー:', error.toString());
    ui.alert('エラー', error.toString(), ui.ButtonSet.OK);
  }
}

/**
 * チャンネルアクセステスト（詳細診断版）
 */
function testChannelAccess() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const configSheet = ss.getSheetByName(SHEETS.CONFIG);
  const config = getConfigData(configSheet);
  
  if (!config.targetChannels || config.targetChannels.length === 0) {
    ui.alert('エラー', '監視対象チャンネルが設定されていません。', ui.ButtonSet.OK);
    return;
  }
  
  let results = [];
  
  // Bot情報を取得
  const botInfo = checkBotPermissions();
  console.log(`Bot名: ${botInfo.name}`);
  
  config.targetChannels.forEach(channelId => {
    try {
      console.log(`\nチャンネル ${channelId} を診断中...`);
      
      // チャンネル情報を取得
      let channelInfo = null;
      let channelName = 'Unknown';
      let isPrivate = false;
      
      try {
        const response = slackAPI('conversations.info', {
          channel: channelId
        });
        
        if (response.ok) {
          channelInfo = response.channel;
          channelName = channelInfo.name;
          isPrivate = channelInfo.is_private;
          console.log(`✅ チャンネル情報取得: ${channelName} (${isPrivate ? 'プライベート' : 'パブリック'})`);
        }
      } catch (infoError) {
        console.log(`❌ チャンネル情報取得失敗: ${infoError.toString()}`);
        
        // invalid_argumentsエラーの場合
        if (infoError.toString().includes('invalid_arguments')) {
          // チャンネルリストから検索
          try {
            const listResponse = slackAPI('conversations.list', {
              types: 'public_channel',
              limit: 1000
            });
            
            const foundChannel = listResponse.channels?.find(ch => ch.id === channelId);
            if (foundChannel) {
              results.push({
                channelId: channelId,
                channelName: foundChannel.name,
                status: '❌ アクセス不可',
                details: `Botはメンバーではありません。\nSlackで「/invite @${botInfo.name}」を実行してください。`
              });
            } else {
              results.push({
                channelId: channelId,
                channelName: 'プライベートチャンネル',
                status: '❌ アクセス不可',
                details: `プライベートチャンネルへのアクセス権限がありません。\nSlackで「/invite @${botInfo.name}」を実行してください。`
              });
            }
          } catch (listError) {
            console.error(`チャンネルリスト取得エラー: ${listError.toString()}`);
          }
          return;
        }
        
        throw infoError;
      }
      
      // メンバーシップ確認
      const isMember = channelInfo?.is_member || false;
      
      if (isMember) {
        // メッセージ履歴を取得（最新1件）
        try {
          const history = slackAPI('conversations.history', {
            channel: channelId,
            limit: 1
          });
          
          results.push({
            channelId: channelId,
            channelName: channelName,
            status: '✅ アクセス可能',
            isMember: true,
            messageCount: history.messages ? history.messages.length : 0,
            details: `正常にアクセスできます。\nタイプ: ${isPrivate ? 'プライベート' : 'パブリック'}チャンネル`
          });
        } catch (historyError) {
          results.push({
            channelId: channelId,
            channelName: channelName,
            status: '⚠️ 制限付きアクセス',
            isMember: true,
            details: `メンバーですが履歴取得不可: ${historyError.toString()}`
          });
        }
      } else {
        results.push({
          channelId: channelId,
          channelName: channelName,
          status: '❌ アクセス不可',
          isMember: false,
          details: `Botはメンバーではありません。\nSlackで「/invite @${botInfo.name}」を実行してください。`
        });
      }
    } catch (error) {
      console.error(`エラー詳細: ${error.toString()}`);
      
      let errorDetails = 'チャンネル情報を取得できません。';
      
      if (error.toString().includes('channel_not_found')) {
        errorDetails = 'チャンネルが見つかりません。IDを確認してください。';
      } else if (error.toString().includes('invalid_auth')) {
        errorDetails = 'Bot Tokenが無効です。';
      } else if (error.toString().includes('not_in_channel')) {
        errorDetails = `Botをチャンネルに招待してください。`;
      }
      
      results.push({
        channelId: channelId,
        status: '❌ エラー',
        error: error.toString(),
        details: errorDetails
      });
    }
  });
  
  // 結果を表示
  let message = `=== チャンネルアクセス診断 ===\n\n`;
  message += `Bot名: ${botInfo.name}\n\n`;
  
  results.forEach(result => {
    message += `【${result.channelName || result.channelId}】\n`;
    message += `状態: ${result.status}\n`;
    message += `${result.details}\n`;
    if (result.error) {
      message += `エラー: ${result.error}\n`;
    }
    message += '\n';
  });
  
  // 推奨アクション
  const needsInvite = results.filter(r => r.status.includes('アクセス不可'));
  if (needsInvite.length > 0) {
    message += `\n=== 必要なアクション ===\n`;
    message += `${needsInvite.length}個のチャンネルへの招待が必要です。\n`;
    needsInvite.forEach(r => {
      message += `• #${r.channelName || r.channelId}\n`;
    });
  }
  
  ui.alert('チャンネルアクセス診断', message, ui.ButtonSet.OK);
}

// ========= Slackメッセージ取得診断 =========
function testSlackMessageRetrieval() {
  const ui = SpreadsheetApp.getUi();
  console.log('=== Slackメッセージ取得診断開始 ===');
  
  // 1. Bot Token確認
  console.log('1. Bot Token確認...');
  if (!SLACK_BOT_TOKEN) {
    ui.alert('エラー', 'SLACK_BOT_TOKENが設定されていません。\nスクリプトプロパティを確認してください。', ui.ButtonSet.OK);
    return;
  }
  
  if (!SLACK_BOT_TOKEN.startsWith('xoxb-')) {
    console.warn('警告: Bot TokenがUser Token (xoxp-)の可能性があります。Bot Token (xoxb-)を使用してください。');
  }
  
  console.log('Bot Token形式: ' + SLACK_BOT_TOKEN.substring(0, 10) + '...');
  
  // 2. Bot情報取得
  console.log('2. Bot情報取得...');
  let botInfo;
  try {
    const authTest = slackAPI('auth.test', {});
    botInfo = {
      user_id: authTest.user_id,
      team: authTest.team,
      url: authTest.url
    };
    console.log(`Bot ID: ${botInfo.user_id}`);
    console.log(`Team: ${botInfo.team}`);
    console.log(`URL: ${botInfo.url}`);
  } catch (error) {
    ui.alert('認証エラー', `Bot Tokenが無効です。\n${error.toString()}`, ui.ButtonSet.OK);
    return;
  }
  
  // 3. チャンネル一覧取得テスト
  console.log('3. チャンネル一覧取得テスト...');
  let channels = [];
  try {
    const response = slackAPI('conversations.list', {
      types: 'public_channel,private_channel',
      limit: 10
    });
    channels = response.channels || [];
    console.log(`取得したチャンネル数: ${channels.length}`);
  } catch (error) {
    console.error('チャンネル一覧取得エラー:', error.toString());
  }
  
  // 4. テストチャンネル選択
  if (channels.length === 0) {
    ui.alert('エラー', 'アクセス可能なチャンネルがありません。', ui.ButtonSet.OK);
    return;
  }
  
  const testChannel = channels[0];
  console.log(`テストチャンネル: #${testChannel.name} (${testChannel.id})`);
  
  // 5. メッセージ取得テスト
  console.log('5. メッセージ取得テスト...');
  let messages = [];
  try {
    const response = slackAPI('conversations.history', {
      channel: testChannel.id,
      limit: 5
    });
    messages = response.messages || [];
    console.log(`取得したメッセージ数: ${messages.length}`);
  } catch (error) {
    console.error('メッセージ取得エラー:', error.toString());
    
    if (error.toString().includes('not_in_channel')) {
      ui.alert('エラー', `Botがチャンネル #${testChannel.name} のメンバーではありません。\nSlackで /invite @bot-name を実行してください。`, ui.ButtonSet.OK);
      return;
    }
  }
  
  // 6. スプレッドシート保存テスト
  console.log('6. スプレッドシート保存テスト...');
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let testSheet = ss.getSheetByName('TestMessages');
  
  if (!testSheet) {
    testSheet = ss.insertSheet('TestMessages');
    const headers = ['タイムスタンプ', 'ユーザー', 'メッセージ', '日時'];
    testSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  
  if (messages.length > 0) {
    const messageData = messages.map(msg => [
      msg.ts,
      msg.user || 'unknown',
      msg.text || '',
      new Date(parseFloat(msg.ts) * 1000)
    ]);
    
    testSheet.getRange(2, 1, messageData.length, 4).setValues(messageData);
    console.log('メッセージをスプレッドシートに保存しました');
  }
  
  // 結果表示
  const result = `
=== Slackメッセージ取得診断結果 ===

✅ Bot Token: 設定済み
✅ Bot認証: 成功
✅ Bot ID: ${botInfo.user_id}
✅ Team: ${botInfo.team}

📊 チャンネル情報:
- 取得可能なチャンネル数: ${channels.length}
- テストチャンネル: #${testChannel.name}
- 取得したメッセージ数: ${messages.length}

${messages.length > 0 ? '✅ メッセージ取得: 成功' : '⚠️ メッセージ取得: メッセージが見つかりません'}

診断完了！
TestMessagesシートで取得したメッセージを確認できます。
  `;
  
  ui.alert('診断結果', result, ui.ButtonSet.OK);
  console.log('=== 診断完了 ===');
}

// ========= メッセージ取得＆分析統合関数 =========
function fetchAndAnalyzeSlackMessages() {
  const ui = SpreadsheetApp.getUi();
  
  // チャンネルIDを入力
  const response = ui.prompt(
    'Slackメッセージ取得＆分析',
    'チャンネルID（例: C01234567）を入力してください:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() !== ui.Button.OK) return;
  
  const channelId = response.getResponseText();
  
  if (!channelId || !channelId.startsWith('C')) {
    ui.alert('エラー', 'チャンネルIDは「C」で始まる必要があります。', ui.ButtonSet.OK);
    return;
  }
  
  try {
    // 1. Bot情報取得
    console.log('Bot情報を取得中...');
    let botInfo;
    try {
      const authTest = slackAPI('auth.test', {});
      botInfo = {
        user_id: authTest.user_id,
        team: authTest.team,
        bot_id: authTest.bot_id
      };
    } catch (error) {
      ui.alert('認証エラー', 'Bot Tokenが無効です。スクリプトプロパティを確認してください。', ui.ButtonSet.OK);
      return;
    }
    
    // 2. チャンネル情報とメンバーシップ確認
    console.log('チャンネル情報を取得中...');
    let channelName = channelId;
    let isMember = false;
    
    try {
      const channelInfo = slackAPI('conversations.info', { channel: channelId });
      channelName = channelInfo.channel?.name || channelId;
      isMember = channelInfo.channel?.is_member || false;
      
      console.log(`チャンネル: #${channelName}, Botメンバー: ${isMember}`);
      
      if (!isMember) {
        // Botがメンバーでない場合のエラーメッセージ
        const errorMessage = `
Botがチャンネル #${channelName} のメンバーではありません。

【解決方法】
1. Slackでチャンネル #${channelName} を開く
2. 以下のコマンドを入力:
   /invite @${botInfo.user_id}
   
または、チャンネル設定から手動でBotを追加してください。

Bot ID: ${botInfo.user_id}
        `;
        ui.alert('チャンネルアクセスエラー', errorMessage, ui.ButtonSet.OK);
        return;
      }
    } catch (error) {
      if (error.toString().includes('channel_not_found')) {
        ui.alert('エラー', `チャンネル ${channelId} が見つかりません。IDを確認してください。`, ui.ButtonSet.OK);
        return;
      }
      // チャンネル情報取得に失敗しても続行を試みる
      console.warn('チャンネル情報取得エラー:', error);
    }
    
    // 3. メッセージ取得
    ui.alert('処理開始', `チャンネル #${channelName} からメッセージを取得中...`, ui.ButtonSet.OK);
    
    const history = slackAPI('conversations.history', {
      channel: channelId,
      limit: 50
    });
    
    const messages = history.messages || [];
    console.log(`取得したメッセージ数: ${messages.length}`);
    
    if (messages.length === 0) {
      ui.alert('情報', 'メッセージが見つかりませんでした。', ui.ButtonSet.OK);
      return;
    }
    
    // 2. スプレッドシートに保存
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const timestamp = new Date().getTime();
    const rawSheetName = `Raw_${channelName}_${timestamp}`;
    const rawSheet = ss.insertSheet(rawSheetName);
    
    // ヘッダー設定
    const headers = ['タイムスタンプ', 'ユーザーID', 'メッセージ', '日時', 'スレッドTS', '返信数'];
    rawSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    rawSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    
    // メッセージデータを整形
    const messageData = messages.map(msg => [
      msg.ts,
      msg.user || '',
      msg.text || '',
      new Date(parseFloat(msg.ts) * 1000),
      msg.thread_ts || '',
      msg.reply_count || 0
    ]);
    
    rawSheet.getRange(2, 1, messageData.length, headers.length).setValues(messageData);
    
    // 3. メッセージ分析
    ui.alert('分析開始', 'AIでメッセージを分析中...', ui.ButtonSet.OK);
    
    const analysisResults = [];
    let agendaItems = [];
    
    // メッセージをバッチで分析
    const batchSize = 10;
    for (let i = 0; i < messages.length; i += batchSize) {
      const batch = messages.slice(i, Math.min(i + batchSize, messages.length));
      const batchText = batch.map(msg => `[${new Date(parseFloat(msg.ts) * 1000).toLocaleString('ja-JP')}] ${msg.text}`).join('\\n\\n');
      
      try {
        const analysis = analyzeMessageBatch(batchText);
        analysisResults.push(analysis);
        
        // 議題を抽出
        if (analysis.agenda_items && analysis.agenda_items.length > 0) {
          agendaItems = agendaItems.concat(analysis.agenda_items);
        }
      } catch (error) {
        console.error(`バッチ ${i/batchSize + 1} の分析エラー:`, error);
      }
    }
    
    // 4. 分析結果を新しいシートに保存
    const analysisSheetName = `Analysis_${channelName}_${timestamp}`;
    const analysisSheet = ss.insertSheet(analysisSheetName);
    
    const analysisHeaders = ['カテゴリ', '重要度', '議題', '概要', '関係者', 'アクションアイテム', '期限'];
    analysisSheet.getRange(1, 1, 1, analysisHeaders.length).setValues([analysisHeaders]);
    analysisSheet.getRange(1, 1, 1, analysisHeaders.length).setFontWeight('bold');
    
    if (agendaItems.length > 0) {
      const agendaData = agendaItems.map(item => [
        item.category || '未分類',
        item.priority || '中',
        item.title || '',
        item.summary || '',
        item.people ? item.people.join(', ') : '',
        item.action_items ? item.action_items.join(', ') : '',
        item.deadline || ''
      ]);
      
      analysisSheet.getRange(2, 1, agendaData.length, analysisHeaders.length).setValues(agendaData);
    }
    
    // 5. 結果表示
    const resultMessage = `
処理完了！

📊 取得結果:
- チャンネル: #${channelName}
- メッセージ数: ${messages.length}件
- 抽出された議題: ${agendaItems.length}件

📁 作成されたシート:
- 生データ: ${rawSheetName}
- 分析結果: ${analysisSheetName}

議題が${agendaItems.length}件抽出されました。
詳細は「${analysisSheetName}」シートをご確認ください。
    `;
    
    ui.alert('完了', resultMessage, ui.ButtonSet.OK);
    
  } catch (error) {
    console.error('エラー詳細:', error);
    ui.alert('エラー', `処理中にエラーが発生しました:\\n${error.toString()}`, ui.ButtonSet.OK);
  }
}

// メッセージバッチを分析
function analyzeMessageBatch(messagesText) {
  const prompt = `
以下のSlackメッセージから、重要な議題、決定事項、アクションアイテムを抽出してください。
以下のJSON形式で出力してください：

{
  "summary": "全体の要約",
  "agenda_items": [
    {
      "category": "カテゴリ（経営/開発/営業/その他）",
      "priority": "重要度（高/中/低）",
      "title": "議題タイトル",
      "summary": "議題の概要",
      "people": ["関係者1", "関係者2"],
      "action_items": ["アクション1", "アクション2"],
      "deadline": "期限（あれば）"
    }
  ],
  "decisions": ["決定事項1", "決定事項2"],
  "next_steps": ["次のステップ1", "次のステップ2"]
}

メッセージ:
${messagesText}

重要: 議題として抽出すべきものがない場合は、agenda_itemsを空配列[]にしてください。`;

  try {
    // OpenAI APIキーの確認
    if (!OPENAI_API_KEY) {
      console.error('OpenAI APIキーが設定されていません');
      return { agenda_items: [] };
    }
    
    const url = 'https://api.openai.com/v1/chat/completions';
    const options = {
      method: 'post',
      headers: {
        'Authorization': `Bearer ${OPENAI_API_KEY}`,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify({
        model: 'gpt-5',
        messages: [
          { role: 'system', content: 'あなたは議事録作成と議題抽出の専門家です。重要な議題を見逃さないようにしてください。' },
          { role: 'user', content: prompt }
        ],
        temperature: 0.3,
        response_format: { type: 'json_object' }
      }),
      muteHttpExceptions: true
    };
    
    const response = UrlFetchApp.fetch(url, options);
    const data = JSON.parse(response.getContentText());
    
    if (data.choices && data.choices[0]) {
      return JSON.parse(data.choices[0].message.content);
    }
    
    return { agenda_items: [] };
    
  } catch (error) {
    console.error('分析エラー:', error);
    return { agenda_items: [] };
  }
}

// ========= 安全なメッセージ取得（conversations.infoを使わない） =========
function getSlackMessagesSafe() {
  const ui = SpreadsheetApp.getUi();
  
  // チャンネルIDを入力
  const response = ui.prompt(
    'Slackメッセージ取得（安全版）',
    'チャンネルID（例: C09BW2EEVAR）を入力してください:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() !== ui.Button.OK) return;
  
  const channelId = response.getResponseText().trim();
  
  if (!channelId || !channelId.startsWith('C')) {
    ui.alert('エラー', 'チャンネルIDは「C」で始まる必要があります。', ui.ButtonSet.OK);
    return;
  }
  
  try {
    console.log(`チャンネル ${channelId} の処理開始...`);
    
    // 方法1: 直接conversations.historyを試す（conversations.infoをスキップ）
    let channelName = channelId;
    let messages = [];
    
    try {
      console.log('conversations.historyでメッセージ取得を試行...');
      const history = slackAPI('conversations.history', {
        channel: channelId,
        limit: 50
      });
      
      messages = history.messages || [];
      console.log(`✅ メッセージ取得成功: ${messages.length}件`);
      
      // チャンネル名を取得する（オプション - conversations.listから）
      try {
        const listResult = slackAPI('conversations.list', {
          limit: 1000,
          types: 'public_channel,private_channel'
        });
        const channel = listResult.channels?.find(ch => ch.id === channelId);
        if (channel) {
          channelName = channel.name;
          console.log(`チャンネル名: #${channelName}`);
          console.log(`Botメンバー: ${channel.is_member ? 'はい' : 'いいえ'}`);
        }
      } catch (listError) {
        console.log('チャンネル名取得失敗（処理は継続）');
      }
      
    } catch (historyError) {
      console.error('conversations.history失敗:', historyError);
      
      // エラーメッセージを分析
      const errorStr = historyError.toString();
      let helpMessage = '';
      
      if (errorStr.includes('not_in_channel') || errorStr.includes('invalid_arguments')) {
        helpMessage = `
Botがチャンネルのメンバーではない可能性があります。

【解決方法】
1. Slackでチャンネルを開く
2. 以下のいずれかを実行:
   - /invite @[bot-name]
   - チャンネル設定 → インテグレーション → アプリを追加

【チャンネルID】
${channelId}
        `;
      } else if (errorStr.includes('channel_not_found')) {
        helpMessage = 'チャンネルが見つかりません。IDを確認してください。';
      } else if (errorStr.includes('missing_scope')) {
        helpMessage = `
必要な権限が不足しています。

【必要なBot Token Scopes】
- channels:history
- channels:read
- groups:history (プライベートチャンネル用)
- groups:read (プライベートチャンネル用)

Slack App設定で権限を追加してください。
        `;
      } else {
        helpMessage = errorStr;
      }
      
      ui.alert('エラー', helpMessage, ui.ButtonSet.OK);
      return;
    }
    
    if (messages.length === 0) {
      ui.alert('情報', 'メッセージが見つかりませんでした。', ui.ButtonSet.OK);
      return;
    }
    
    // スプレッドシートに保存
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const timestamp = new Date().getTime();
    const sheetName = `Slack_${channelName}_${timestamp}`;
    const sheet = ss.insertSheet(sheetName);
    
    // ヘッダー設定
    const headers = ['タイムスタンプ', 'ユーザーID', 'メッセージ', '日時', 'スレッドTS', '返信数'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    
    // メッセージデータを整形
    const messageData = messages.map(msg => [
      msg.ts,
      msg.user || '',
      msg.text || '',
      new Date(parseFloat(msg.ts) * 1000),
      msg.thread_ts || '',
      msg.reply_count || 0
    ]);
    
    // データを書き込み
    sheet.getRange(2, 1, messageData.length, headers.length).setValues(messageData);
    
    // 列幅を調整
    sheet.autoResizeColumns(1, headers.length);
    
    ui.alert(
      '取得完了',
      `${messages.length}件のメッセージを取得しました。\nシート「${sheetName}」に保存されました。`,
      ui.ButtonSet.OK
    );
    
  } catch (error) {
    console.error('エラー詳細:', error);
    ui.alert('エラー', `処理中にエラーが発生しました:\n${error.toString()}`, ui.ButtonSet.OK);
  }
}

// ========= 簡易メッセージ取得関数 =========
function getSlackMessagesSimple() {
  const ui = SpreadsheetApp.getUi();
  
  // チャンネルIDを入力
  const response = ui.prompt(
    'チャンネル指定',
    'チャンネルID（例: C01234567）を入力してください:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() !== ui.Button.OK) return;
  
  const channelId = response.getResponseText();
  
  if (!channelId || !channelId.startsWith('C')) {
    ui.alert('エラー', 'チャンネルIDは「C」で始まる必要があります。', ui.ButtonSet.OK);
    return;
  }
  
  try {
    // チャンネル情報取得（conversations.listを使用）
    let channelName = channelId;
    try {
      const listResult = slackAPI('conversations.list', {
        limit: 1000,
        types: 'public_channel,private_channel'
      });
      const channel = listResult.channels?.find(ch => ch.id === channelId);
      if (channel) {
        channelName = channel.name;
      }
    } catch (error) {
      console.log('チャンネル名取得失敗（処理は継続）');
    }
    
    console.log(`チャンネル: #${channelName} (${channelId})`);
    
    // メッセージ取得
    const history = slackAPI('conversations.history', {
      channel: channelId,
      limit: 100
    });
    
    const messages = history.messages || [];
    console.log(`取得したメッセージ数: ${messages.length}`);
    
    if (messages.length === 0) {
      ui.alert('情報', 'メッセージが見つかりませんでした。', ui.ButtonSet.OK);
      return;
    }
    
    // スプレッドシートに保存
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheetName = `Slack_${channelName}_${new Date().getTime()}`;
    const sheet = ss.insertSheet(sheetName);
    
    // ヘッダー設定
    const headers = [
      'タイムスタンプ',
      'ユーザーID',
      'メッセージ',
      '日時',
      'スレッドTS',
      '返信数'
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    
    // メッセージデータを整形
    const messageData = messages.map(msg => [
      msg.ts,
      msg.user || '',
      msg.text || '',
      new Date(parseFloat(msg.ts) * 1000),
      msg.thread_ts || '',
      msg.reply_count || 0
    ]);
    
    // データを書き込み
    sheet.getRange(2, 1, messageData.length, headers.length).setValues(messageData);
    
    // 列幅を調整
    sheet.autoResizeColumns(1, headers.length);
    
    ui.alert(
      '取得完了',
      `${messages.length}件のメッセージを取得しました。\nシート「${sheetName}」に保存されました。`,
      ui.ButtonSet.OK
    );
    
  } catch (error) {
    console.error('エラー詳細:', error);
    
    let errorMessage = error.toString();
    
    if (error.toString().includes('not_in_channel')) {
      errorMessage = 'Botがチャンネルのメンバーではありません。\nSlackで /invite @bot-name を実行してください。';
    } else if (error.toString().includes('channel_not_found')) {
      errorMessage = 'チャンネルが見つかりません。IDを確認してください。';
    } else if (error.toString().includes('invalid_auth')) {
      errorMessage = 'Bot Tokenが無効です。スクリプトプロパティを確認してください。';
    }
    
    ui.alert('エラー', errorMessage, ui.ButtonSet.OK);
  }
}

// ========= Botをチャンネルに追加する関数 =========
function joinBotToChannel() {
  const ui = SpreadsheetApp.getUi();
  
  // チャンネルIDを入力
  const response = ui.prompt(
    'Botをチャンネルに追加',
    'チャンネルID（例: C09BW2EEVAR）を入力してください:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() !== ui.Button.OK) return;
  
  const channelId = response.getResponseText().trim();
  
  if (!channelId || !channelId.startsWith('C')) {
    ui.alert('エラー', 'チャンネルIDは「C」で始まる必要があります。', ui.ButtonSet.OK);
    return;
  }
  
  try {
    // Bot情報を取得
    const authTest = slackAPI('auth.test', {});
    const botUserId = authTest.user_id;
    
    console.log(`Bot User ID: ${botUserId}`);
    console.log(`チャンネルID: ${channelId}`);
    
    // conversations.joinを使用してBotをチャンネルに参加させる
    try {
      const joinResult = slackAPI('conversations.join', {
        channel: channelId
      });
      
      if (joinResult.ok) {
        ui.alert('成功', `Botをチャンネルに追加しました。\nチャンネル: ${joinResult.channel?.name || channelId}`, ui.ButtonSet.OK);
      }
    } catch (joinError) {
      console.error('conversations.join エラー:', joinError);
      
      // エラーの詳細分析
      const errorStr = joinError.toString();
      
      if (errorStr.includes('method_not_supported_for_channel_type')) {
        // プライベートチャンネルの場合、招待が必要
        const inviteMessage = `
このチャンネルはプライベートチャンネルです。
Botを追加するには、Slackで以下の手順を実行してください：

1. チャンネルを開く
2. チャンネル名をクリック → 「インテグレーション」タブ
3. 「アプリを追加」をクリック
4. お使いのBotアプリを選択

または、チャンネルで以下のコマンドを実行：
/invite @${botUserId}
        `;
        ui.alert('プライベートチャンネル', inviteMessage, ui.ButtonSet.OK);
      } else if (errorStr.includes('already_in_channel')) {
        ui.alert('情報', 'Botは既にこのチャンネルのメンバーです。', ui.ButtonSet.OK);
      } else if (errorStr.includes('is_archived')) {
        ui.alert('エラー', 'このチャンネルはアーカイブされています。', ui.ButtonSet.OK);
      } else {
        ui.alert('エラー', `チャンネル参加エラー: ${errorStr}`, ui.ButtonSet.OK);
      }
    }
    
  } catch (error) {
    ui.alert('エラー', `エラーが発生しました: ${error.toString()}`, ui.ButtonSet.OK);
  }
}

// ========= 統合連続ワークフロー（メイン関数） =========
function executeIntegratedContinuousWorkflow(channelId = null, autoMode = false) {
  const startTime = new Date();
  const workflowLog = [];
  let workflowResult = {
    success: false,
    steps: [],
    messages: [],
    analysis: null,
    governance: null,
    spreadsheetUrl: null,
    notifications: [],
    errors: []
  };
  
  try {
    // ステップ1: Slackからメッセージ取得
    workflowLog.push(createLogEntry('START', 'ワークフロー開始', `チャンネル: ${channelId || 'ALL'}`));
    
    const messages = channelId 
      ? fetchChannelMessagesWithRetry(channelId)
      : fetchAllJoinedChannelsMessages();
    
    if (!messages || messages.length === 0) {
      // メッセージが取得できない場合は警告として記録
      console.warn('メッセージが取得できませんでした');
      workflowLog.push(createLogEntry('WARNING', 'メッセージなし', 'チャンネルアクセス権限を確認してください'));
      
      // デモデータで続行
      const demoMessages = [{
        text: 'サンプル: 新しいプロジェクトの予算承認が必要です',
        user: 'demo_user',
        ts: Date.now() / 1000
      }];
      
      workflowResult.messages = demoMessages;
      workflowResult.analysis = {
        categories: ['予算'],
        topics: [{title: 'デモ分析', description: 'メッセージ取得失敗のためデモデータ使用', priority: 1}],
        priority: 'LOW',
        priorityReason: 'デモデータ',
        actionItems: [],
        stakeholders: [],
        urgency: 'normal',
        deadline: '',
        decisions: [],
        risks: [],
        resources: { human: [], financial: '', time: '' },
        kpis: [],
        summary: 'メッセージ取得失敗。チャンネルアクセス権限を確認してください。'
      };
      workflowResult.governance = {
        requiresApproval: false,
        requiresDisclosure: false,
        requiresAction: false,
        riskLevel: 'LOW',
        controlNumber: generateControlNumber()
      };
      
      // 通知送信
      const notificationResult = sendEnhancedNotifications(workflowResult);
      workflowResult.notifications = notificationResult;
      workflowLog.push(createLogEntry('INFO', 'デモモード', 'メッセージ取得失敗のためデモデータで処理'));
      
      recordComprehensiveWorkflowLogs(workflowLog, workflowResult);
      workflowResult.success = true;
      
      if (!autoMode) {
        const ui = SpreadsheetApp.getUi();
        ui.alert('処理完了', 'メッセージが取得できませんでした。\n\n【必要な設定】\n1. Slackアプリをチャンネルに追加\n2. 必要なスコープ:\n - channels:history\n - channels:read\n - groups:history (プライベートチャンネル用)\n - groups:read (プライベートチャンネル用)', ui.ButtonSet.OK);
      }
      
      return workflowResult;
    }
    
    workflowResult.messages = messages;
    workflowLog.push(createLogEntry('SUCCESS', 'メッセージ取得完了', `${messages.length}件`));
    
    // ステップ2: 包括的AI分析
    const analysisResult = performEnhancedAIAnalysis(messages);
    workflowResult.analysis = analysisResult;
    workflowLog.push(createLogEntry('SUCCESS', 'AI分析完了', `重要度: ${analysisResult.priority}`));
    
    // ステップ3: ガバナンス・コンプライアンスチェック
    const governanceResult = performComprehensiveGovernanceCheck(messages, analysisResult);
    workflowResult.governance = governanceResult;
    workflowLog.push(createLogEntry('SUCCESS', 'ガバナンスチェック完了', `リスク: ${governanceResult.riskLevel}`));
    
    // ステップ4: 重要度判定による分岐処理
    if (analysisResult.priority === 'HIGH' || governanceResult.requiresAction) {
      // ステップ5: 新規スプレッドシート作成（重要案件用）
      const newSpreadsheet = createDetailedWorkflowSpreadsheet(analysisResult, governanceResult, messages);
      workflowResult.spreadsheetUrl = newSpreadsheet.getUrl();
      workflowLog.push(createLogEntry('SUCCESS', 'スプレッドシート作成', workflowResult.spreadsheetUrl));
      
      // ステップ6: 業務フロー・文書生成
      generateComprehensiveWorkflowDocuments(newSpreadsheet, analysisResult, governanceResult);
      workflowLog.push(createLogEntry('SUCCESS', '業務フロー生成完了', ''));
      
      // ステップ7: 議事録案作成（必要な場合）
      if (governanceResult.requiresMeetingMinutes) {
        generateDetailedMeetingMinutes(newSpreadsheet, analysisResult, governanceResult);
        workflowLog.push(createLogEntry('SUCCESS', '議事録案作成完了', ''));
      }
    }
    
    // ステップ8: 統合通知送信（HTML形式メール＆Slack）
    const notificationResult = sendEnhancedNotifications(workflowResult);
    workflowResult.notifications = notificationResult;
    workflowLog.push(createLogEntry('SUCCESS', '通知送信完了', `メール: ${notificationResult.email}, Slack: ${notificationResult.slack}`));
    
    // ステップ9: 包括的ログ記録
    recordComprehensiveWorkflowLogs(workflowLog, workflowResult);
    
    workflowResult.success = true;
    
  } catch (error) {
    workflowResult.errors.push(error.toString());
    workflowLog.push(createLogEntry('ERROR', 'エラー発生', error.toString()));
    recordComprehensiveWorkflowLogs(workflowLog, workflowResult);
    sendErrorNotification(error, workflowResult);
  }
  
  const executionTime = (new Date() - startTime) / 1000;
  console.log(`ワークフロー完了: ${executionTime}秒`);
  
  if (!autoMode) {
    const ui = SpreadsheetApp.getUi();
    ui.alert('完了', `ワークフロー処理が完了しました\n実行時間: ${executionTime}秒\n${workflowResult.spreadsheetUrl ? 'スプレッドシート: ' + workflowResult.spreadsheetUrl : ''}`, ui.ButtonSet.OK);
  }
  
  return workflowResult;
}

// ========= アプリ統合で業務フロー生成＆通知（UI版） =========
function getMessagesAsAppWithWorkflow() {
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.prompt(
    '統合ワークフロー実行',
    'チャンネルID（例: C09BW2EEVAR）を入力してください:\n※空欄の場合は全参加チャンネルを処理',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() !== ui.Button.OK) return;
  
  const channelId = response.getResponseText().trim();
  
  if (channelId && !channelId.startsWith('C')) {
    ui.alert('エラー', 'チャンネルIDは「C」で始まる必要があります。', ui.ButtonSet.OK);
    return;
  }
  
  try {
    // 統合ワークフロー実行
    executeIntegratedContinuousWorkflow(channelId || null, false);
    
    // 既存の処理は統合ワークフローに移行済み
    
    // 6. 通知の準備と送信
    const notificationData = {
      channelName: channelName,
      messageCount: messages.length,
      taskCount: workflowData.tasks.length,
      flowSteps: workflowData.flowSteps.length,
      spreadsheetUrl: ss.getUrl(),
      sheets: {
        spec: specSheetName,
        flow: flowSheetName,
        task: taskSheetName
      }
    };
    
    // メール通知送信
    if (REPORT_EMAIL) {
      sendWorkflowNotificationEmail(notificationData);
    }
    
    // Slack通知送信
    const configSheet = ss.getSheetByName(SHEETS.CONFIG);
    if (configSheet) {
      const config = getConfigData(configSheet);
      if (config.notifySlackChannel) {
        sendWorkflowSlackNotification(config.notifySlackChannel, notificationData);
      }
    }
    
    // 7. 結果表示
    const resultMessage = `
業務フロー生成完了！

📊 分析結果:
- チャンネル: #${channelName}
- 分析メッセージ数: ${messages.length}件
- 抽出タスク数: ${workflowData.tasks.length}件
- フローステップ数: ${workflowData.flowSteps.length}件

📁 作成されたシート:
- 業務記述書: ${specSheetName}
- 業務フロー: ${flowSheetName}
- タスク管理: ${taskSheetName}

📧 通知:
- メール: ${REPORT_EMAIL ? '送信済み' : '未設定'}
- Slack: ${notificationData.slackNotified ? '送信済み' : '未設定'}

詳細はスプレッドシートをご確認ください。
    `;
    
    ui.alert('生成完了', resultMessage, ui.ButtonSet.OK);
    
  } catch (error) {
    console.error('エラー:', error);
    ui.alert('エラー', `処理中にエラーが発生しました:\n${error.toString()}`, ui.ButtonSet.OK);
  }
}

// タスク抽出と業務フロー生成
function extractTasksAndCreateWorkflow(messages) {
  const tasks = [];
  const flowSteps = [];
  const actors = new Set();
  
  // タスク関連キーワード
  const taskKeywords = {
    action: ['する', 'します', 'してください', 'お願い', '依頼', 'タスク', 'TODO', 'やること'],
    deadline: ['まで', '期限', '締切', 'いつまで', 'デッドライン'],
    responsible: ['担当', '責任者', 'オーナー', '@'],
    priority: ['至急', '緊急', '重要', '優先', 'ASAP'],
    process: ['手順', 'プロセス', 'フロー', '流れ', 'ステップ']
  };
  
  messages.forEach((msg, index) => {
    if (!msg.text) return;
    
    const text = msg.text;
    const msgDate = new Date(parseFloat(msg.ts) * 1000);
    
    // タスク抽出
    if (taskKeywords.action.some(kw => text.includes(kw))) {
      const task = {
        id: `TASK-${index + 1}`,
        description: text.substring(0, 200),
        createdAt: msgDate,
        user: msg.user || 'unknown',
        priority: taskKeywords.priority.some(kw => text.includes(kw)) ? '高' : '中',
        status: '未着手'
      };
      
      // 期限の抽出
      if (taskKeywords.deadline.some(kw => text.includes(kw))) {
        task.deadline = extractDeadline(text);
      }
      
      // 担当者の抽出
      if (text.includes('@')) {
        const mentions = text.match(/@[\w\-]+/g);
        if (mentions) {
          task.assignee = mentions[0].replace('@', '');
          actors.add(task.assignee);
        }
      }
      
      tasks.push(task);
    }
    
    // フローステップの抽出
    if (taskKeywords.process.some(kw => text.includes(kw))) {
      flowSteps.push({
        stepNo: flowSteps.length + 1,
        description: text.substring(0, 150),
        type: determineStepType(text),
        actor: msg.user || 'unknown'
      });
      actors.add(msg.user || 'unknown');
    }
  });
  
  // フローステップが少ない場合、タスクから生成
  if (flowSteps.length < 3 && tasks.length > 0) {
    tasks.forEach((task, index) => {
      flowSteps.push({
        stepNo: index + 1,
        description: task.description,
        type: '処理',
        actor: task.assignee || task.user
      });
    });
  }
  
  return {
    tasks: tasks,
    flowSteps: flowSteps,
    actors: Array.from(actors),
    summary: `${tasks.length}個のタスクと${flowSteps.length}個のプロセスステップを抽出`
  };
}

// 期限抽出
function extractDeadline(text) {
  const patterns = [
    /(\d{1,2}月\d{1,2}日)/,
    /(\d{4}年\d{1,2}月\d{1,2}日)/,
    /(今週|来週|今月|来月)末?/,
    /(\d+)日まで/
  ];
  
  for (const pattern of patterns) {
    const match = text.match(pattern);
    if (match) {
      return match[1];
    }
  }
  return '未定';
}

// ステップタイプ判定
function determineStepType(text) {
  if (text.includes('判断') || text.includes('確認') || text.includes('レビュー')) {
    return '判断';
  } else if (text.includes('承認') || text.includes('決裁')) {
    return '承認';
  } else if (text.includes('通知') || text.includes('連絡') || text.includes('共有')) {
    return '連絡';
  }
  return '処理';
}

// 業務記述書生成
function generateBusinessSpecification(workflowData, channelName) {
  return {
    title: `業務記述書 - ${channelName}`,
    purpose: `${channelName}チャンネルで議論された業務プロセスの文書化`,
    scope: 'Slackメッセージから抽出された業務タスクとフロー',
    overview: workflowData.summary,
    actors: workflowData.actors,
    tasks: workflowData.tasks,
    flowSteps: workflowData.flowSteps,
    createdDate: new Date()
  };
}

// 業務記述書シート作成
function createBusinessSpecSheet(sheet, spec) {
  let row = 1;
  
  // タイトル
  sheet.getRange(row, 1).setValue(spec.title);
  sheet.getRange(row, 1).setFontSize(16).setFontWeight('bold');
  row += 2;
  
  // 基本情報
  const basicInfo = [
    ['項目', '内容'],
    ['目的', spec.purpose],
    ['範囲', spec.scope],
    ['概要', spec.overview],
    ['作成日', spec.createdDate.toLocaleDateString('ja-JP')]
  ];
  
  sheet.getRange(row, 1, basicInfo.length, 2).setValues(basicInfo);
  sheet.getRange(row, 1, 1, 2).setFontWeight('bold').setBackground('#e3f2fd');
  row += basicInfo.length + 2;
  
  // 関係者
  sheet.getRange(row, 1).setValue('関係者');
  sheet.getRange(row, 1).setFontWeight('bold').setBackground('#f5f5f5');
  row++;
  
  spec.actors.forEach(actor => {
    sheet.getRange(row, 1).setValue(`• ${actor}`);
    row++;
  });
  row += 2;
  
  // タスク一覧
  sheet.getRange(row, 1).setValue('タスク一覧');
  sheet.getRange(row, 1).setFontWeight('bold').setBackground('#f5f5f5');
  row++;
  
  const taskHeaders = ['ID', 'タスク内容', '担当者', '優先度', '期限', 'ステータス'];
  sheet.getRange(row, 1, 1, taskHeaders.length).setValues([taskHeaders]);
  sheet.getRange(row, 1, 1, taskHeaders.length).setFontWeight('bold');
  row++;
  
  spec.tasks.forEach(task => {
    const taskRow = [
      task.id,
      task.description,
      task.assignee || '未割当',
      task.priority,
      task.deadline || '未定',
      task.status
    ];
    sheet.getRange(row, 1, 1, taskRow.length).setValues([taskRow]);
    row++;
  });
  
  sheet.autoResizeColumns(1, 6);
}

// 業務フローシート作成
function createWorkflowSheet(sheet, workflowData) {
  let row = 1;
  
  sheet.getRange(row, 1).setValue('業務フロー図');
  sheet.getRange(row, 1).setFontSize(16).setFontWeight('bold');
  row += 2;
  
  const headers = ['ステップNo', '作業内容', 'タイプ', '担当者', '備考'];
  sheet.getRange(row, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(row, 1, 1, headers.length).setFontWeight('bold').setBackground('#e3f2fd');
  row++;
  
  workflowData.flowSteps.forEach(step => {
    const stepRow = [
      step.stepNo,
      step.description,
      step.type,
      step.actor,
      ''
    ];
    sheet.getRange(row, 1, 1, stepRow.length).setValues([stepRow]);
    
    // タイプによる色分け
    if (step.type === '判断') {
      sheet.getRange(row, 3).setBackground('#fff3e0');
    } else if (step.type === '承認') {
      sheet.getRange(row, 3).setBackground('#e8f5e9');
    } else if (step.type === '連絡') {
      sheet.getRange(row, 3).setBackground('#e3f2fd');
    }
    
    row++;
  });
  
  // フローチャート風の視覚化
  row += 2;
  sheet.getRange(row, 1).setValue('フローチャート');
  sheet.getRange(row, 1).setFontWeight('bold');
  row++;
  
  workflowData.flowSteps.forEach((step, index) => {
    const symbol = step.type === '判断' ? '◆' : '□';
    const arrow = index < workflowData.flowSteps.length - 1 ? '↓' : '';
    
    sheet.getRange(row, 2).setValue(`${symbol} ${step.description}`);
    row++;
    
    if (arrow) {
      sheet.getRange(row, 2).setValue(arrow);
      row++;
    }
  });
  
  sheet.autoResizeColumns(1, 5);
}

// タスク管理シート作成
function createTaskManagementSheet(sheet, tasks) {
  let row = 1;
  
  sheet.getRange(row, 1).setValue('タスク管理表');
  sheet.getRange(row, 1).setFontSize(16).setFontWeight('bold');
  row += 2;
  
  const headers = ['ID', 'タスク', '担当者', '優先度', '期限', 'ステータス', '作成日', '更新日', '進捗', 'メモ'];
  sheet.getRange(row, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(row, 1, 1, headers.length).setFontWeight('bold').setBackground('#e3f2fd');
  row++;
  
  tasks.forEach(task => {
    const taskRow = [
      task.id,
      task.description,
      task.assignee || '',
      task.priority,
      task.deadline || '',
      task.status,
      task.createdAt,
      new Date(),
      '0%',
      ''
    ];
    sheet.getRange(row, 1, 1, taskRow.length).setValues([taskRow]);
    
    // 優先度による色分け
    if (task.priority === '高') {
      sheet.getRange(row, 4).setBackground('#ffebee');
    } else if (task.priority === '中') {
      sheet.getRange(row, 4).setBackground('#fff3e0');
    }
    
    // ステータスのドロップダウン設定
    const statusRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['未着手', '進行中', '完了', '保留', 'キャンセル'], true)
      .build();
    sheet.getRange(row, 6).setDataValidation(statusRule);
    
    // 進捗のドロップダウン設定
    const progressRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['0%', '25%', '50%', '75%', '100%'], true)
      .build();
    sheet.getRange(row, 9).setDataValidation(progressRule);
    
    row++;
  });
  
  sheet.autoResizeColumns(1, 10);
}

// ワークフロー通知メール送信
function sendWorkflowNotificationEmail(data) {
  const subject = `[業務フロー生成] ${data.channelName} - タスク${data.taskCount}件`;
  
  const htmlBody = `
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <style>
    body { font-family: 'Helvetica Neue', Arial, sans-serif; line-height: 1.6; color: #333; }
    .container { max-width: 600px; margin: 0 auto; padding: 20px; }
    .header { background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; padding: 30px; border-radius: 10px 10px 0 0; }
    .content { background: white; padding: 30px; border: 1px solid #e0e0e0; border-radius: 0 0 10px 10px; }
    .stats { display: flex; justify-content: space-around; margin: 20px 0; }
    .stat-box { text-align: center; padding: 15px; background: #f8f9fa; border-radius: 8px; }
    .stat-number { font-size: 24px; font-weight: bold; color: #667eea; }
    .button { display: inline-block; padding: 12px 24px; background: #667eea; color: white; text-decoration: none; border-radius: 5px; margin: 10px 5px; }
    .sheets-list { background: #f8f9fa; padding: 15px; border-radius: 8px; margin: 20px 0; }
  </style>
</head>
<body>
  <div class="container">
    <div class="header">
      <h1>📊 業務フロー生成完了</h1>
      <p>チャンネル: #${data.channelName}</p>
    </div>
    <div class="content">
      <div class="stats">
        <div class="stat-box">
          <div class="stat-number">${data.messageCount}</div>
          <div>メッセージ分析</div>
        </div>
        <div class="stat-box">
          <div class="stat-number">${data.taskCount}</div>
          <div>タスク抽出</div>
        </div>
        <div class="stat-box">
          <div class="stat-number">${data.flowSteps}</div>
          <div>フローステップ</div>
        </div>
      </div>
      
      <div class="sheets-list">
        <h3>📁 作成されたドキュメント</h3>
        <ul>
          <li>業務記述書: ${data.sheets.spec}</li>
          <li>業務フロー図: ${data.sheets.flow}</li>
          <li>タスク管理表: ${data.sheets.task}</li>
        </ul>
      </div>
      
      <div style="text-align: center; margin-top: 30px;">
        <a href="${data.spreadsheetUrl}" class="button">スプレッドシートを開く</a>
      </div>
      
      <hr style="margin: 30px 0; border: none; border-top: 1px solid #e0e0e0;">
      
      <p style="color: #666; font-size: 12px;">
        このメールは自動生成されました。<br>
        生成日時: ${new Date().toLocaleString('ja-JP')}
      </p>
    </div>
  </div>
</body>
</html>
  `;
  
  const plainBody = `
業務フロー生成完了

チャンネル: #${data.channelName}

【分析結果】
- メッセージ数: ${data.messageCount}件
- 抽出タスク: ${data.taskCount}件
- フローステップ: ${data.flowSteps}件

【作成ドキュメント】
- 業務記述書: ${data.sheets.spec}
- 業務フロー図: ${data.sheets.flow}
- タスク管理表: ${data.sheets.task}

スプレッドシート: ${data.spreadsheetUrl}

生成日時: ${new Date().toLocaleString('ja-JP')}
  `;
  
  try {
    MailApp.sendEmail({
      to: REPORT_EMAIL,
      subject: subject,
      body: plainBody,
      htmlBody: htmlBody
    });
    console.log('ワークフロー通知メール送信完了');
  } catch (error) {
    console.error('メール送信エラー:', error);
  }
}

// ワークフローSlack通知送信
function sendWorkflowSlackNotification(channelId, data) {
  const blocks = [
    {
      type: 'header',
      text: {
        type: 'plain_text',
        text: '📊 業務フロー生成完了'
      }
    },
    {
      type: 'section',
      fields: [
        {
          type: 'mrkdwn',
          text: `*チャンネル:*\n#${data.channelName}`
        },
        {
          type: 'mrkdwn',
          text: `*メッセージ数:*\n${data.messageCount}件`
        },
        {
          type: 'mrkdwn',
          text: `*タスク数:*\n${data.taskCount}件`
        },
        {
          type: 'mrkdwn',
          text: `*フローステップ:*\n${data.flowSteps}件`
        }
      ]
    },
    {
      type: 'section',
      text: {
        type: 'mrkdwn',
        text: `📁 *作成されたドキュメント:*\n• 業務記述書\n• 業務フロー図\n• タスク管理表`
      }
    },
    {
      type: 'actions',
      elements: [
        {
          type: 'button',
          text: {
            type: 'plain_text',
            text: 'スプレッドシートを開く'
          },
          url: data.spreadsheetUrl,
          style: 'primary'
        }
      ]
    }
  ];
  
  try {
    slackAPI('chat.postMessage', {
      channel: channelId,
      text: `業務フロー生成完了 - ${data.channelName}`,
      blocks: blocks
    });
    
    data.slackNotified = true;
    console.log('Slack通知送信完了');
  } catch (error) {
    console.error('Slack通知エラー:', error);
    data.slackNotified = false;
  }
}

// ========= アプリ統合でガバナンス分析付きメッセージ取得 =========
function getMessagesAsAppWithGovernance() {
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.prompt(
    'ガバナンス分析付きメッセージ取得',
    'チャンネルID（例: C09BW2EEVAR）を入力してください:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() !== ui.Button.OK) return;
  
  const channelId = response.getResponseText().trim();
  
  if (!channelId || !channelId.startsWith('C')) {
    ui.alert('エラー', 'チャンネルIDは「C」で始まる必要があります。', ui.ButtonSet.OK);
    return;
  }
  
  try {
    console.log(`ガバナンス分析モードでチャンネル ${channelId} にアクセス...`);
    
    // 1. チャンネル参加試行
    try {
      const joinResult = slackAPI('conversations.join', { channel: channelId });
      console.log('チャンネル参加成功:', joinResult.channel?.name);
    } catch (joinError) {
      console.log('チャンネル参加スキップ（プライベートチャンネルの可能性）');
    }
    
    // 2. メッセージ取得
    const history = slackAPI('conversations.history', {
      channel: channelId,
      limit: 200  // ガバナンス分析のため多めに取得
    });
    
    const messages = history.messages || [];
    console.log(`メッセージ取得成功: ${messages.length}件`);
    
    if (messages.length === 0) {
      ui.alert('情報', 'メッセージが見つかりませんでした。', ui.ButtonSet.OK);
      return;
    }
    
    // チャンネル名取得
    let channelName = channelId;
    try {
      const listResult = slackAPI('conversations.list', { limit: 1000 });
      const channel = listResult.channels?.find(ch => ch.id === channelId);
      if (channel) channelName = channel.name;
    } catch (error) {
      console.log('チャンネル名取得失敗');
    }
    
    ui.alert('処理開始', `${messages.length}件のメッセージを取得しました。\nガバナンス分析を開始します...`, ui.ButtonSet.OK);
    
    // 3. ガバナンス観点での分析
    const governanceResults = analyzeMessagesForGovernance(messages);
    
    // 4. スプレッドシート作成
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const timestamp = new Date().getTime();
    
    // 生データシート
    const rawSheetName = `Raw_${channelName}_${timestamp}`;
    const rawSheet = ss.insertSheet(rawSheetName);
    saveRawMessages(rawSheet, messages);
    
    // ガバナンス分析シート
    const govSheetName = `Gov_${channelName}_${timestamp}`;
    const govSheet = ss.insertSheet(govSheetName);
    createGovernanceAnalysisSheet(govSheet, governanceResults, channelName);
    
    // 5. 結果表示
    const resultMessage = `
ガバナンス分析完了！

📊 分析結果:
- チャンネル: #${channelName}
- 分析メッセージ数: ${messages.length}件

🏛️ ガバナンス観点:
- 承認・決裁関連: ${governanceResults.approvalItems.length}件
- 開示要件該当: ${governanceResults.disclosureItems.length}件
- リスク要因: ${governanceResults.riskItems.length}件
- コンプライアンス要注意: ${governanceResults.complianceItems.length}件

📁 作成されたシート:
- 生データ: ${rawSheetName}
- ガバナンス分析: ${govSheetName}

詳細はスプレッドシートをご確認ください。
    `;
    
    ui.alert('分析完了', resultMessage, ui.ButtonSet.OK);
    
  } catch (error) {
    console.error('エラー:', error);
    ui.alert('エラー', `処理中にエラーが発生しました:\n${error.toString()}`, ui.ButtonSet.OK);
  }
}

// ガバナンス観点でメッセージを分析
function analyzeMessagesForGovernance(messages) {
  const results = {
    approvalItems: [],      // 承認・決裁関連
    disclosureItems: [],    // 開示要件
    riskItems: [],          // リスク要因
    complianceItems: [],    // コンプライアンス
    decisions: [],          // 重要な決定事項
    advisorNeeds: []        // 専門家相談が必要な事項
  };
  
  // ガバナンス関連キーワード
  const governanceKeywords = {
    approval: ['承認', '決裁', '決議', '稟議', '許可', '認可', 'approve', 'approval'],
    disclosure: ['開示', '公表', '発表', 'リリース', 'IR', '適時開示', '決算', '業績'],
    risk: ['リスク', '問題', '課題', '懸念', '遅延', '損失', '事故', '違反'],
    compliance: ['コンプライアンス', '法令', '規制', '監査', '内部統制', '違反', '不正'],
    decision: ['決定', '決議', '方針', '戦略', '計画', '予算', '投資'],
    advisor: ['弁護士', '会計士', '税理士', '社労士', '専門家', '相談', 'コンサル']
  };
  
  messages.forEach(msg => {
    if (!msg.text) return;
    
    const text = msg.text.toLowerCase();
    const msgDate = new Date(parseFloat(msg.ts) * 1000);
    const messageInfo = {
      text: msg.text,
      date: msgDate.toLocaleString('ja-JP'),
      user: msg.user || 'unknown',
      ts: msg.ts
    };
    
    // 承認・決裁関連
    if (governanceKeywords.approval.some(kw => text.includes(kw))) {
      results.approvalItems.push({
        ...messageInfo,
        type: '承認・決裁',
        importance: determineImportance(msg.text),
        requiredAction: '承認フローの確認と記録'
      });
    }
    
    // 開示要件チェック
    if (governanceKeywords.disclosure.some(kw => text.includes(kw))) {
      const disclosureCheck = checkDisclosureRequirement(msg.text);
      if (disclosureCheck.requiresDisclosure) {
        results.disclosureItems.push({
          ...messageInfo,
          type: '開示要件',
          disclosureType: disclosureCheck.disclosureType,
          timeline: disclosureCheck.timeline,
          regulations: disclosureCheck.regulations
        });
      }
    }
    
    // リスク要因
    if (governanceKeywords.risk.some(kw => text.includes(kw))) {
      results.riskItems.push({
        ...messageInfo,
        type: 'リスク要因',
        riskLevel: assessRiskLevel(msg.text),
        mitigation: suggestMitigation(msg.text)
      });
    }
    
    // コンプライアンス
    if (governanceKeywords.compliance.some(kw => text.includes(kw))) {
      results.complianceItems.push({
        ...messageInfo,
        type: 'コンプライアンス',
        category: identifyComplianceCategory(msg.text),
        action: '法務部門への確認推奨'
      });
    }
    
    // 重要な決定事項
    if (governanceKeywords.decision.some(kw => text.includes(kw))) {
      results.decisions.push({
        ...messageInfo,
        type: '決定事項',
        impact: assessDecisionImpact(msg.text)
      });
    }
    
    // 専門家相談の必要性
    if (governanceKeywords.advisor.some(kw => text.includes(kw))) {
      results.advisorNeeds.push({
        ...messageInfo,
        type: '専門家相談',
        advisorType: identifyAdvisorType(msg.text),
        urgency: assessUrgency(msg.text)
      });
    }
  });
  
  return results;
}

// 重要度判定
function determineImportance(text) {
  if (text.includes('取締役') || text.includes('億円') || text.includes('重要')) {
    return '高';
  } else if (text.includes('部長') || text.includes('百万円')) {
    return '中';
  }
  return '低';
}

// リスクレベル評価
function assessRiskLevel(text) {
  if (text.includes('重大') || text.includes('深刻') || text.includes('違反')) {
    return '高';
  } else if (text.includes('懸念') || text.includes('課題')) {
    return '中';
  }
  return '低';
}

// リスク軽減策提案
function suggestMitigation(text) {
  if (text.includes('違反')) {
    return '即座に法務部門と相談し、是正措置を実施';
  } else if (text.includes('遅延')) {
    return 'スケジュール見直しとリソース追加検討';
  }
  return '継続的なモニタリングと早期対応';
}

// コンプライアンスカテゴリ特定
function identifyComplianceCategory(text) {
  if (text.includes('個人情報') || text.includes('プライバシー')) {
    return '個人情報保護';
  } else if (text.includes('インサイダー') || text.includes('内部情報')) {
    return 'インサイダー取引規制';
  } else if (text.includes('下請法') || text.includes('独占禁止')) {
    return '競争法';
  }
  return '一般コンプライアンス';
}

// 決定事項の影響度評価
function assessDecisionImpact(text) {
  if (text.includes('戦略') || text.includes('方針') || text.includes('億')) {
    return '全社レベル';
  } else if (text.includes('部門') || text.includes('プロジェクト')) {
    return '部門レベル';
  }
  return '個別案件';
}

// 必要な専門家タイプ特定
function identifyAdvisorType(text) {
  if (text.includes('弁護士') || text.includes('法的')) {
    return '弁護士';
  } else if (text.includes('会計') || text.includes('税')) {
    return '会計士・税理士';
  } else if (text.includes('労務') || text.includes('雇用')) {
    return '社労士';
  }
  return '専門コンサルタント';
}

// 緊急度評価
function assessUrgency(text) {
  if (text.includes('至急') || text.includes('緊急') || text.includes('即')) {
    return '緊急';
  } else if (text.includes('早急') || text.includes('速やか')) {
    return '高';
  }
  return '通常';
}

// ガバナンス分析シート作成
function createGovernanceAnalysisSheet(sheet, results, channelName) {
  let row = 1;
  
  // タイトル
  sheet.getRange(row, 1).setValue(`ガバナンス・コンプライアンス分析レポート - #${channelName}`);
  sheet.getRange(row, 1).setFontSize(14).setFontWeight('bold');
  row += 2;
  
  // サマリー
  sheet.getRange(row, 1).setValue('📊 分析サマリー');
  sheet.getRange(row, 1).setFontWeight('bold').setBackground('#e3f2fd');
  row++;
  
  const summary = [
    ['項目', '件数', '重要度'],
    ['承認・決裁事項', results.approvalItems.length, results.approvalItems.filter(i => i.importance === '高').length + '件が高重要度'],
    ['開示要件', results.disclosureItems.length, results.disclosureItems.length > 0 ? '要確認' : '-'],
    ['リスク要因', results.riskItems.length, results.riskItems.filter(i => i.riskLevel === '高').length + '件が高リスク'],
    ['コンプライアンス', results.complianceItems.length, results.complianceItems.length > 0 ? '要確認' : '-'],
    ['重要決定事項', results.decisions.length, results.decisions.filter(i => i.impact === '全社レベル').length + '件が全社影響'],
    ['専門家相談', results.advisorNeeds.length, results.advisorNeeds.filter(i => i.urgency === '緊急').length + '件が緊急']
  ];
  
  sheet.getRange(row, 1, summary.length, 3).setValues(summary);
  sheet.getRange(row, 1, 1, 3).setFontWeight('bold');
  row += summary.length + 2;
  
  // 各セクションの詳細
  const sections = [
    { title: '🔐 承認・決裁事項', data: results.approvalItems, headers: ['日時', 'メッセージ', '重要度', '必要アクション'] },
    { title: '📢 開示要件', data: results.disclosureItems, headers: ['日時', 'メッセージ', '開示種別', 'タイムライン'] },
    { title: '⚠️ リスク要因', data: results.riskItems, headers: ['日時', 'メッセージ', 'リスクレベル', '軽減策'] },
    { title: '⚖️ コンプライアンス', data: results.complianceItems, headers: ['日時', 'メッセージ', 'カテゴリ', 'アクション'] },
    { title: '✅ 重要決定事項', data: results.decisions, headers: ['日時', 'メッセージ', '影響度'] },
    { title: '👥 専門家相談', data: results.advisorNeeds, headers: ['日時', 'メッセージ', '専門家タイプ', '緊急度'] }
  ];
  
  sections.forEach(section => {
    if (section.data.length > 0) {
      // セクションタイトル
      sheet.getRange(row, 1).setValue(section.title);
      sheet.getRange(row, 1).setFontWeight('bold').setBackground('#f5f5f5');
      row++;
      
      // ヘッダー
      sheet.getRange(row, 1, 1, section.headers.length).setValues([section.headers]);
      sheet.getRange(row, 1, 1, section.headers.length).setFontWeight('bold');
      row++;
      
      // データ
      section.data.forEach(item => {
        const rowData = [];
        if (section.title.includes('承認')) {
          rowData.push(item.date, item.text.substring(0, 100), item.importance, item.requiredAction);
        } else if (section.title.includes('開示')) {
          rowData.push(item.date, item.text.substring(0, 100), 
            Array.isArray(item.disclosureType) ? item.disclosureType.join(', ') : item.disclosureType,
            item.timeline);
        } else if (section.title.includes('リスク')) {
          rowData.push(item.date, item.text.substring(0, 100), item.riskLevel, item.mitigation);
        } else if (section.title.includes('コンプライアンス')) {
          rowData.push(item.date, item.text.substring(0, 100), item.category, item.action);
        } else if (section.title.includes('決定')) {
          rowData.push(item.date, item.text.substring(0, 100), item.impact);
        } else if (section.title.includes('専門家')) {
          rowData.push(item.date, item.text.substring(0, 100), item.advisorType, item.urgency);
        }
        
        sheet.getRange(row, 1, 1, rowData.length).setValues([rowData]);
        
        // 高重要度/高リスク/緊急のハイライト
        if ((item.importance === '高') || (item.riskLevel === '高') || (item.urgency === '緊急')) {
          sheet.getRange(row, 1, 1, rowData.length).setBackground('#ffebee');
        }
        
        row++;
      });
      
      row += 2; // セクション間のスペース
    }
  });
  
  // 推奨アクション
  sheet.getRange(row, 1).setValue('💡 推奨アクション');
  sheet.getRange(row, 1).setFontWeight('bold').setBackground('#fff3e0');
  row++;
  
  const recommendations = [];
  if (results.disclosureItems.length > 0) {
    recommendations.push('• IR部門と法務部門に開示要件を確認してください');
  }
  if (results.riskItems.filter(i => i.riskLevel === '高').length > 0) {
    recommendations.push('• 高リスク事項について経営層への報告を検討してください');
  }
  if (results.complianceItems.length > 0) {
    recommendations.push('• コンプライアンス部門による詳細レビューを実施してください');
  }
  if (results.advisorNeeds.filter(i => i.urgency === '緊急').length > 0) {
    recommendations.push('• 緊急の専門家相談事項について早急に対応してください');
  }
  
  if (recommendations.length === 0) {
    recommendations.push('• 現時点で緊急の対応事項はありません');
  }
  
  recommendations.forEach(rec => {
    sheet.getRange(row, 1).setValue(rec);
    row++;
  });
  
  // 列幅調整
  sheet.autoResizeColumns(1, 4);
}

// 生メッセージ保存
function saveRawMessages(sheet, messages) {
  const headers = ['タイムスタンプ', 'ユーザーID', 'メッセージ', '日時', 'スレッドTS', '返信数'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  
  const messageData = messages.map(msg => [
    msg.ts,
    msg.user || '',
    msg.text || '',
    new Date(parseFloat(msg.ts) * 1000),
    msg.thread_ts || '',
    msg.reply_count || 0
  ]);
  
  if (messageData.length > 0) {
    sheet.getRange(2, 1, messageData.length, headers.length).setValues(messageData);
  }
  
  sheet.autoResizeColumns(1, headers.length);
}

// ========= アプリ統合でメッセージ取得＆分析（通常版） =========
function getMessagesAsAppAndAnalyze() {
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.prompt(
    'アプリ統合メッセージ取得＆分析',
    'チャンネルID（例: C09BW2EEVAR）を入力してください:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() !== ui.Button.OK) return;
  
  const channelId = response.getResponseText().trim();
  
  if (!channelId || !channelId.startsWith('C')) {
    ui.alert('エラー', 'チャンネルIDは「C」で始まる必要があります。', ui.ButtonSet.OK);
    return;
  }
  
  try {
    console.log(`アプリとしてチャンネル ${channelId} にアクセス...`);
    
    // 1. まずconversations.joinを試みる（パブリックチャンネルの場合）
    try {
      console.log('チャンネルへの参加を試行...');
      const joinResult = slackAPI('conversations.join', {
        channel: channelId
      });
      console.log('チャンネル参加成功:', joinResult.channel?.name);
    } catch (joinError) {
      console.log('チャンネル参加失敗（プライベートチャンネルの可能性）:', joinError.toString());
    }
    
    // 2. メッセージ取得を試行
    let messages = [];
    let channelName = channelId;
    
    try {
      console.log('メッセージ取得を試行...');
      const history = slackAPI('conversations.history', {
        channel: channelId,
        limit: 100  // 分析のために多めに取得
      });
      
      messages = history.messages || [];
      console.log(`メッセージ取得成功: ${messages.length}件`);
      
    } catch (historyError) {
      console.error('メッセージ取得失敗:', historyError);
      ui.alert('エラー', `メッセージ取得エラー: ${historyError.toString()}`, ui.ButtonSet.OK);
      return;
    }
    
    if (messages.length === 0) {
      ui.alert('情報', 'メッセージが見つかりませんでした。', ui.ButtonSet.OK);
      return;
    }
    
    // チャンネル名を取得
    try {
      const listResult = slackAPI('conversations.list', {
        limit: 1000
      });
      const channel = listResult.channels?.find(ch => ch.id === channelId);
      if (channel) {
        channelName = channel.name;
      }
    } catch (error) {
      console.log('チャンネル名取得失敗');
    }
    
    ui.alert('取得成功', `${messages.length}件のメッセージを取得しました。\n分析を開始します...`, ui.ButtonSet.OK);
    
    // 3. スプレッドシートに生データを保存
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const timestamp = new Date().getTime();
    const rawSheetName = `Raw_${channelName}_${timestamp}`;
    const rawSheet = ss.insertSheet(rawSheetName);
    
    const headers = ['タイムスタンプ', 'ユーザーID', 'メッセージ', '日時', 'スレッドTS', '返信数'];
    rawSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    rawSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    
    const messageData = messages.map(msg => [
      msg.ts,
      msg.user || '',
      msg.text || '',
      new Date(parseFloat(msg.ts) * 1000),
      msg.thread_ts || '',
      msg.reply_count || 0
    ]);
    
    rawSheet.getRange(2, 1, messageData.length, headers.length).setValues(messageData);
    rawSheet.autoResizeColumns(1, headers.length);
    
    // 4. メッセージ分析
    console.log('メッセージ分析開始...');
    let agendaItems = [];
    let decisions = [];
    let actionItems = [];
    
    // メッセージをバッチで分析
    const batchSize = 10;
    for (let i = 0; i < messages.length; i += batchSize) {
      const batch = messages.slice(i, Math.min(i + batchSize, messages.length));
      const batchText = batch.map(msg => {
        const date = new Date(parseFloat(msg.ts) * 1000);
        return `[${date.toLocaleString('ja-JP')}] ${msg.text || ''}`;
      }).join('\n\n');
      
      try {
        const analysis = analyzeMessageBatch(batchText);
        
        if (analysis.agenda_items && analysis.agenda_items.length > 0) {
          agendaItems = agendaItems.concat(analysis.agenda_items);
        }
        if (analysis.decisions && analysis.decisions.length > 0) {
          decisions = decisions.concat(analysis.decisions);
        }
        if (analysis.next_steps && analysis.next_steps.length > 0) {
          actionItems = actionItems.concat(analysis.next_steps);
        }
      } catch (error) {
        console.error(`バッチ ${Math.floor(i/batchSize) + 1} の分析エラー:`, error);
      }
    }
    
    // 5. 分析結果を新しいシートに保存
    if (agendaItems.length > 0 || decisions.length > 0 || actionItems.length > 0) {
      const analysisSheetName = `Analysis_${channelName}_${timestamp}`;
      const analysisSheet = ss.insertSheet(analysisSheetName);
      
      // 議題セクション
      let row = 1;
      analysisSheet.getRange(row, 1).setValue('📋 議題・トピック');
      analysisSheet.getRange(row, 1).setFontWeight('bold').setBackground('#f0f0f0');
      row++;
      
      const agendaHeaders = ['カテゴリ', '重要度', '議題', '概要', '関係者', 'アクション'];
      analysisSheet.getRange(row, 1, 1, agendaHeaders.length).setValues([agendaHeaders]);
      analysisSheet.getRange(row, 1, 1, agendaHeaders.length).setFontWeight('bold');
      row++;
      
      if (agendaItems.length > 0) {
        const agendaData = agendaItems.map(item => [
          item.category || '未分類',
          item.priority || '中',
          item.title || '',
          item.summary || '',
          item.people ? item.people.join(', ') : '',
          item.action_items ? item.action_items.join(', ') : ''
        ]);
        analysisSheet.getRange(row, 1, agendaData.length, agendaHeaders.length).setValues(agendaData);
        row += agendaData.length;
      }
      
      // 決定事項セクション
      row += 2;
      analysisSheet.getRange(row, 1).setValue('✅ 決定事項');
      analysisSheet.getRange(row, 1).setFontWeight('bold').setBackground('#e8f5e9');
      row++;
      
      if (decisions.length > 0) {
        decisions.forEach((decision, index) => {
          analysisSheet.getRange(row, 1).setValue(`${index + 1}. ${decision}`);
          row++;
        });
      }
      
      // アクションアイテムセクション
      row += 2;
      analysisSheet.getRange(row, 1).setValue('🎯 アクションアイテム');
      analysisSheet.getRange(row, 1).setFontWeight('bold').setBackground('#fff3e0');
      row++;
      
      if (actionItems.length > 0) {
        actionItems.forEach((action, index) => {
          analysisSheet.getRange(row, 1).setValue(`${index + 1}. ${action}`);
          row++;
        });
      }
      
      analysisSheet.autoResizeColumns(1, 6);
    }
    
    // 6. 結果表示
    const resultMessage = `
分析完了！

📊 分析結果:
- チャンネル: #${channelName}
- 分析メッセージ数: ${messages.length}件
- 抽出された議題: ${agendaItems.length}件
- 決定事項: ${decisions.length}件
- アクションアイテム: ${actionItems.length}件

📁 作成されたシート:
- 生データ: ${rawSheetName}
- 分析結果: Analysis_${channelName}_${timestamp}

詳細はスプレッドシートをご確認ください。
    `;
    
    ui.alert('分析完了', resultMessage, ui.ButtonSet.OK);
    
  } catch (error) {
    console.error('エラー:', error);
    ui.alert('エラー', `処理中にエラーが発生しました:\n${error.toString()}`, ui.ButtonSet.OK);
  }
}

// ========= アプリ統合用のメッセージ取得（分析なし） =========
function getMessagesAsApp() {
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.prompt(
    'アプリ統合メッセージ取得',
    'チャンネルID（例: C09BW2EEVAR）を入力してください:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() !== ui.Button.OK) return;
  
  const channelId = response.getResponseText().trim();
  
  if (!channelId || !channelId.startsWith('C')) {
    ui.alert('エラー', 'チャンネルIDは「C」で始まる必要があります。', ui.ButtonSet.OK);
    return;
  }
  
  try {
    console.log(`アプリとしてチャンネル ${channelId} にアクセス...`);
    
    // 1. まずconversations.joinを試みる（パブリックチャンネルの場合）
    try {
      console.log('チャンネルへの参加を試行...');
      const joinResult = slackAPI('conversations.join', {
        channel: channelId
      });
      console.log('チャンネル参加成功:', joinResult.channel?.name);
    } catch (joinError) {
      console.log('チャンネル参加失敗（プライベートチャンネルの可能性）:', joinError.toString());
    }
    
    // 2. メッセージ取得を試行
    let messages = [];
    let channelName = channelId;
    
    try {
      console.log('メッセージ取得を試行...');
      const history = slackAPI('conversations.history', {
        channel: channelId,
        limit: 50
      });
      
      messages = history.messages || [];
      console.log(`メッセージ取得成功: ${messages.length}件`);
      
    } catch (historyError) {
      console.error('メッセージ取得失敗:', historyError);
      
      const errorStr = historyError.toString();
      
      if (errorStr.includes('not_in_channel') || errorStr.includes('invalid_arguments')) {
        const helpMessage = `
Botがチャンネルにアクセスできません。

【パブリックチャンネルの場合】
このツールの「Botをチャンネルに追加」機能を使用してください。

【プライベートチャンネルの場合】
Slackで以下の手順を実行してください：

方法1: コマンドを使用
1. チャンネルで /invite @[bot-name] を実行

方法2: チャンネル設定から追加
1. チャンネル名をクリック
2. 「インテグレーション」タブを選択
3. 「アプリを追加」をクリック
4. お使いのBotアプリを選択

【確認事項】
- Bot Token Scopesに必要な権限があるか
  • channels:join (パブリックチャンネル参加用)
  • channels:history
  • channels:read
  • groups:history (プライベートチャンネル用)
  • groups:read (プライベートチャンネル用)
        `;
        ui.alert('アクセスエラー', helpMessage, ui.ButtonSet.OK);
        return;
      } else {
        ui.alert('エラー', `メッセージ取得エラー: ${errorStr}`, ui.ButtonSet.OK);
        return;
      }
    }
    
    if (messages.length === 0) {
      ui.alert('情報', 'メッセージが見つかりませんでした。', ui.ButtonSet.OK);
      return;
    }
    
    // チャンネル名を取得
    try {
      const listResult = slackAPI('conversations.list', {
        limit: 1000
      });
      const channel = listResult.channels?.find(ch => ch.id === channelId);
      if (channel) {
        channelName = channel.name;
      }
    } catch (error) {
      console.log('チャンネル名取得失敗');
    }
    
    // スプレッドシートに保存
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const timestamp = new Date().getTime();
    const sheetName = `App_${channelName}_${timestamp}`;
    const sheet = ss.insertSheet(sheetName);
    
    const headers = ['タイムスタンプ', 'ユーザーID', 'メッセージ', '日時', 'スレッドTS', '返信数'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    
    const messageData = messages.map(msg => [
      msg.ts,
      msg.user || '',
      msg.text || '',
      new Date(parseFloat(msg.ts) * 1000),
      msg.thread_ts || '',
      msg.reply_count || 0
    ]);
    
    sheet.getRange(2, 1, messageData.length, headers.length).setValues(messageData);
    sheet.autoResizeColumns(1, headers.length);
    
    ui.alert(
      '取得完了',
      `${messages.length}件のメッセージを取得しました。\nシート「${sheetName}」に保存されました。`,
      ui.ButtonSet.OK
    );
    
  } catch (error) {
    console.error('エラー:', error);
    ui.alert('エラー', `処理中にエラーが発生しました:\n${error.toString()}`, ui.ButtonSet.OK);
  }
}

// ========= 詳細なBot診断 =========
function detailedBotDiagnostics() {
  const ui = SpreadsheetApp.getUi();
  console.log('=== 詳細Bot診断開始 ===');
  
  let diagnosticResults = [];
  
  // 1. Token形式チェック
  diagnosticResults.push('【1. Token形式チェック】');
  if (!SLACK_BOT_TOKEN) {
    diagnosticResults.push('❌ TOKENが設定されていません');
    ui.alert('エラー', 'SLACK_BOT_TOKENが設定されていません', ui.ButtonSet.OK);
    return;
  }
  
  const tokenPrefix = SLACK_BOT_TOKEN.substring(0, 5);
  diagnosticResults.push(`Token形式: ${tokenPrefix}...`);
  
  if (tokenPrefix === 'xoxb-') {
    diagnosticResults.push('✅ Bot User OAuth Token（推奨）');
  } else if (tokenPrefix === 'xoxp-') {
    diagnosticResults.push('⚠️ User OAuth Token（非推奨）');
    diagnosticResults.push('   Bot User OAuth Tokenの使用を推奨します');
  } else {
    diagnosticResults.push('❌ 不明なToken形式');
  }
  
  // 2. auth.test API呼び出し
  diagnosticResults.push('\n【2. 認証テスト】');
  let authInfo;
  try {
    authInfo = slackAPI('auth.test', {});
    diagnosticResults.push('✅ 認証成功');
    diagnosticResults.push(`   User ID: ${authInfo.user_id}`);
    diagnosticResults.push(`   Team: ${authInfo.team}`);
    diagnosticResults.push(`   Bot ID: ${authInfo.bot_id || 'なし'}`);
    diagnosticResults.push(`   Is Bot: ${authInfo.is_bot}`);
  } catch (error) {
    diagnosticResults.push(`❌ 認証失敗: ${error.toString()}`);
    ui.alert('診断結果', diagnosticResults.join('\n'), ui.ButtonSet.OK);
    return;
  }
  
  // 3. 利用可能なメソッドのテスト
  diagnosticResults.push('\n【3. API メソッドテスト】');
  
  // conversations.list テスト
  try {
    const listResult = slackAPI('conversations.list', {
      limit: 1,
      types: 'public_channel'
    });
    diagnosticResults.push('✅ conversations.list: 成功');
    
    // Botが参加しているチャンネル数をカウント
    const fullList = slackAPI('conversations.list', {
      limit: 100,
      types: 'public_channel,private_channel'
    });
    const joinedChannels = (fullList.channels || []).filter(ch => ch.is_member);
    diagnosticResults.push(`   Botが参加中: ${joinedChannels.length}チャンネル`);
    
  } catch (error) {
    diagnosticResults.push(`❌ conversations.list: ${error.toString()}`);
  }
  
  // 4. チャンネル別アクセステスト
  diagnosticResults.push('\n【4. チャンネル別アクセステスト】');
  
  const testChannelId = ui.prompt(
    'チャンネルテスト',
    'テストするチャンネルID (例: C09BW2EEVAR) を入力:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (testChannelId.getSelectedButton() === ui.Button.OK) {
    const channelId = testChannelId.getResponseText();
    
    // 方法1: conversations.info (新しいスコープが必要な場合がある)
    diagnosticResults.push(`\nチャンネル ${channelId} のテスト:`);
    
    try {
      const infoResult = slackAPI('conversations.info', {
        channel: channelId,
        include_locale: false,
        include_num_members: false
      });
      diagnosticResults.push('✅ conversations.info: 成功');
      diagnosticResults.push(`   チャンネル名: ${infoResult.channel?.name}`);
      diagnosticResults.push(`   is_member: ${infoResult.channel?.is_member}`);
    } catch (error) {
      diagnosticResults.push(`❌ conversations.info: ${error.toString()}`);
      
      // 代替方法: conversations.listから探す
      try {
        diagnosticResults.push('\n代替方法: conversations.listから検索...');
        const listResult = slackAPI('conversations.list', {
          limit: 1000,
          types: 'public_channel,private_channel'
        });
        
        const targetChannel = listResult.channels?.find(ch => ch.id === channelId);
        if (targetChannel) {
          diagnosticResults.push(`✅ チャンネル発見: ${targetChannel.name}`);
          diagnosticResults.push(`   is_member: ${targetChannel.is_member}`);
          diagnosticResults.push(`   is_private: ${targetChannel.is_private}`);
        } else {
          diagnosticResults.push('❌ チャンネルが見つかりません');
        }
      } catch (listError) {
        diagnosticResults.push(`❌ リスト検索も失敗: ${listError.toString()}`);
      }
    }
    
    // 方法2: conversations.history (直接メッセージ取得)
    try {
      const historyResult = slackAPI('conversations.history', {
        channel: channelId,
        limit: 1
      });
      diagnosticResults.push('✅ conversations.history: 成功');
      diagnosticResults.push(`   メッセージ取得可能`);
    } catch (error) {
      diagnosticResults.push(`❌ conversations.history: ${error.toString()}`);
    }
  }
  
  // 5. 必要なスコープの確認
  diagnosticResults.push('\n【5. 推奨スコープ】');
  diagnosticResults.push('Bot Token Scopesに以下が必要:');
  diagnosticResults.push('   • channels:history');
  diagnosticResults.push('   • channels:read');
  diagnosticResults.push('   • groups:history (プライベートチャンネル用)');
  diagnosticResults.push('   • groups:read (プライベートチャンネル用)');
  diagnosticResults.push('   • users:read (ユーザー情報取得用)');
  
  // 結果表示
  const result = diagnosticResults.join('\n');
  console.log(result);
  ui.alert('詳細Bot診断結果', result, ui.ButtonSet.OK);
}

// ========= Botが参加しているチャンネルから選択して取得＆分析 =========
function fetchAndAnalyzeFromJoinedChannels() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    // 1. Bot情報取得
    console.log('Bot情報を取得中...');
    const authTest = slackAPI('auth.test', {});
    const botInfo = {
      user_id: authTest.user_id,
      team: authTest.team
    };
    
    // 2. Botが参加しているチャンネル一覧を取得（conversations.listを使用）
    console.log('チャンネル一覧を取得中...');
    const response = slackAPI('conversations.list', {
      types: 'public_channel,private_channel',
      exclude_archived: true,
      limit: 1000  // 制限を増やす
    });
    
    const channels = response.channels || [];
    
    // Botがメンバーのチャンネルのみフィルタ
    const joinedChannels = channels.filter(ch => ch.is_member);
    
    if (joinedChannels.length === 0) {
      ui.alert('情報', 'Botが参加しているチャンネルがありません。\nSlackでBotをチャンネルに招待してください。', ui.ButtonSet.OK);
      return;
    }
    
    // 3. チャンネル選択ダイアログ
    const channelList = joinedChannels.map((ch, index) => 
      `${index + 1}. #${ch.name} (${ch.num_members}人) ${ch.is_private ? '🔒' : ''}`
    ).join('\n');
    
    const selectionPrompt = ui.prompt(
      'チャンネル選択',
      `Botが参加しているチャンネル一覧:\n\n${channelList}\n\n番号を入力してください (1-${joinedChannels.length}):`,
      ui.ButtonSet.OK_CANCEL
    );
    
    if (selectionPrompt.getSelectedButton() !== ui.Button.OK) return;
    
    const selection = parseInt(selectionPrompt.getResponseText());
    if (isNaN(selection) || selection < 1 || selection > joinedChannels.length) {
      ui.alert('エラー', '無効な番号です。', ui.ButtonSet.OK);
      return;
    }
    
    const selectedChannel = joinedChannels[selection - 1];
    const channelId = selectedChannel.id;
    const channelName = selectedChannel.name;
    
    // 4. メッセージ取得
    ui.alert('処理開始', `チャンネル #${channelName} からメッセージを取得中...`, ui.ButtonSet.OK);
    
    const history = slackAPI('conversations.history', {
      channel: channelId,
      limit: 50
    });
    
    const messages = history.messages || [];
    console.log(`取得したメッセージ数: ${messages.length}`);
    
    if (messages.length === 0) {
      ui.alert('情報', 'メッセージが見つかりませんでした。', ui.ButtonSet.OK);
      return;
    }
    
    // 5. スプレッドシートに保存
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const timestamp = new Date().getTime();
    const rawSheetName = `Raw_${channelName}_${timestamp}`;
    const rawSheet = ss.insertSheet(rawSheetName);
    
    // ヘッダー設定
    const headers = ['タイムスタンプ', 'ユーザーID', 'メッセージ', '日時', 'スレッドTS', '返信数'];
    rawSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    rawSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    
    // メッセージデータを整形
    const messageData = messages.map(msg => [
      msg.ts,
      msg.user || '',
      msg.text || '',
      new Date(parseFloat(msg.ts) * 1000),
      msg.thread_ts || '',
      msg.reply_count || 0
    ]);
    
    rawSheet.getRange(2, 1, messageData.length, headers.length).setValues(messageData);
    
    // 6. メッセージ分析
    ui.alert('分析開始', 'AIでメッセージを分析中...', ui.ButtonSet.OK);
    
    const analysisResults = [];
    let agendaItems = [];
    
    // メッセージをバッチで分析
    const batchSize = 10;
    for (let i = 0; i < messages.length; i += batchSize) {
      const batch = messages.slice(i, Math.min(i + batchSize, messages.length));
      const batchText = batch.map(msg => `[${new Date(parseFloat(msg.ts) * 1000).toLocaleString('ja-JP')}] ${msg.text}`).join('\n\n');
      
      try {
        const analysis = analyzeMessageBatch(batchText);
        analysisResults.push(analysis);
        
        // 議題を抽出
        if (analysis.agenda_items && analysis.agenda_items.length > 0) {
          agendaItems = agendaItems.concat(analysis.agenda_items);
        }
      } catch (error) {
        console.error(`バッチ ${i/batchSize + 1} の分析エラー:`, error);
      }
    }
    
    // 7. 分析結果を新しいシートに保存
    const analysisSheetName = `Analysis_${channelName}_${timestamp}`;
    const analysisSheet = ss.insertSheet(analysisSheetName);
    
    const analysisHeaders = ['カテゴリ', '重要度', '議題', '概要', '関係者', 'アクションアイテム', '期限'];
    analysisSheet.getRange(1, 1, 1, analysisHeaders.length).setValues([analysisHeaders]);
    analysisSheet.getRange(1, 1, 1, analysisHeaders.length).setFontWeight('bold');
    
    if (agendaItems.length > 0) {
      const agendaData = agendaItems.map(item => [
        item.category || '未分類',
        item.priority || '中',
        item.title || '',
        item.summary || '',
        item.people ? item.people.join(', ') : '',
        item.action_items ? item.action_items.join(', ') : '',
        item.deadline || ''
      ]);
      
      analysisSheet.getRange(2, 1, agendaData.length, analysisHeaders.length).setValues(agendaData);
    }
    
    // 8. 結果表示
    const resultMessage = `
処理完了！

📊 取得結果:
- チャンネル: #${channelName}
- メッセージ数: ${messages.length}件
- 抽出された議題: ${agendaItems.length}件

📁 作成されたシート:
- 生データ: ${rawSheetName}
- 分析結果: ${analysisSheetName}

議題が${agendaItems.length}件抽出されました。
詳細は「${analysisSheetName}」シートをご確認ください。
    `;
    
    ui.alert('完了', resultMessage, ui.ButtonSet.OK);
    
  } catch (error) {
    console.error('エラー詳細:', error);
    
    let errorMessage = error.toString();
    
    if (error.toString().includes('invalid_auth')) {
      errorMessage = 'Bot Tokenが無効です。スクリプトプロパティを確認してください。';
    } else if (error.toString().includes('missing_scope')) {
      errorMessage = '必要な権限が不足しています。Bot Token Scopesを確認してください。\n必要なスコープ: channels:history, channels:read';
    }
    
    ui.alert('エラー', `処理中にエラーが発生しました:\n${errorMessage}`, ui.ButtonSet.OK);
  }
}


// ========= チャンネルIDを取得するヘルパー関数 =========
function getChannelIdByName() {
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.prompt(
    'チャンネル名からID検索',
    'チャンネル名を入力してください（#は不要）:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() !== ui.Button.OK) return;
  
  const channelName = response.getResponseText().toLowerCase().replace('#', '');
  
  try {
    const result = slackAPI('conversations.list', {
      types: 'public_channel,private_channel',
      limit: 1000
    });
    
    const channels = result.channels || [];
    const matchedChannels = channels.filter(ch => 
      ch.name.toLowerCase().includes(channelName)
    );
    
    if (matchedChannels.length === 0) {
      ui.alert('結果', `「${channelName}」に一致するチャンネルが見つかりませんでした。`, ui.ButtonSet.OK);
      return;
    }
    
    const resultText = matchedChannels.map(ch => 
      `#${ch.name}\nID: ${ch.id}\nメンバー: ${ch.num_members}人\nBotメンバー: ${ch.is_member ? '✅' : '❌'}\n`
    ).join('\n---\n');
    
    ui.alert('検索結果', resultText, ui.ButtonSet.OK);
    
  } catch (error) {
    ui.alert('エラー', `チャンネル検索中にエラーが発生しました:\n${error.toString()}`, ui.ButtonSet.OK);
  }
}

/**
 * Botの権限を確認
 */
function checkBotPermissions() {
  try {
    const authTest = slackAPI('auth.test', {});
    return {
      name: authTest.user || 'your-bot',
      userId: authTest.user_id,
      teamId: authTest.team_id,
      botId: authTest.bot_id || null
    };
  } catch (error) {
    console.error('Bot権限確認エラー:', error.toString());
    return {
      name: 'your-bot',
      userId: null,
      teamId: null,
      botId: null
    };
  }
}

/**
 * 利用可能なチャンネルリストを取得
 */
function listAvailableChannels() {
  const ui = SpreadsheetApp.getUi();
  console.log('=== チャンネルリスト取得開始 ===');
  
  let allChannels = [];
  let message = '=== 利用可能なチャンネル一覧 ===\n\n';
  
  try {
    // パブリックチャンネルを取得
    console.log('パブリックチャンネルを取得中...');
    const publicChannels = slackAPI('conversations.list', {
      types: 'public_channel',
      exclude_archived: true,
      limit: 1000
    });
    
    if (publicChannels.channels) {
      message += '【パブリックチャンネル】\n';
      publicChannels.channels.forEach(channel => {
        const memberStatus = channel.is_member ? '✅ メンバー' : '❌ 非メンバー';
        message += `• #${channel.name} (${channel.id}) ${memberStatus}\n`;
        allChannels.push({
          id: channel.id,
          name: channel.name,
          type: 'public',
          isMember: channel.is_member
        });
      });
      message += '\n';
    }
    
    // プライベートチャンネル（Botがメンバーのもののみ）を取得
    console.log('プライベートチャンネルを取得中...');
    try {
      const privateChannels = slackAPI('conversations.list', {
        types: 'private_channel',
        exclude_archived: true,
        limit: 1000
      });
      
      if (privateChannels.channels && privateChannels.channels.length > 0) {
        message += '【プライベートチャンネル（Botがメンバー）】\n';
        privateChannels.channels.forEach(channel => {
          message += `• 🔒#${channel.name} (${channel.id}) ✅ メンバー\n`;
          allChannels.push({
            id: channel.id,
            name: channel.name,
            type: 'private',
            isMember: true
          });
        });
        message += '\n';
      }
    } catch (privateError) {
      console.log('プライベートチャンネル取得エラー（権限不足の可能性）:', privateError.toString());
      message += '【プライベートチャンネル】\n';
      message += '※ Botがメンバーのプライベートチャンネルのみ表示されます\n\n';
    }
    
    // チャンネルIDの取得方法を追加
    message += '=== チャンネルIDの確認方法 ===\n';
    message += '1. Slackでチャンネルを右クリック\n';
    message += '2. 「リンクをコピー」を選択\n';
    message += '3. URLの最後の部分がチャンネルID\n';
    message += '   例: https://xxx.slack.com/archives/C09BW2EEVAR\n';
    message += '   → チャンネルID: C09BW2EEVAR\n\n';
    
    message += '=== 重要な注意事項 ===\n';
    message += '• プライベートチャンネルはBotを招待しないと表示されません\n';
    message += '• チャンネルにBotを招待: /invite @slack_governance\n';
    message += '• 招待後、このリストを再度実行してください\n\n';
    
    message += `合計: ${allChannels.length}チャンネル\n`;
    message += `Botがメンバー: ${allChannels.filter(c => c.isMember).length}チャンネル\n`;
    
  } catch (error) {
    console.error('チャンネルリスト取得エラー:', error.toString());
    message += '\n❌ エラー: ' + error.toString() + '\n';
    message += '\nBot Tokenが正しく設定されているか確認してください。\n';
  }
  
  ui.alert('チャンネルリスト', message, ui.ButtonSet.OK);
  
  return allChannels;
}

/**
 * チャンネルIDを検証
 */
function validateChannelId(channelId) {
  // SlackのチャンネルIDの形式をチェック
  // C で始まり、その後に8-11文字の英数字が続く
  const channelIdPattern = /^C[A-Z0-9]{8,11}$/;
  
  if (!channelIdPattern.test(channelId)) {
    console.error(`無効なチャンネルID形式: ${channelId}`);
    console.log('正しいチャンネルIDは「C」で始まり、その後に8-11文字の英数字が続きます');
    console.log('例: C09BW2EEVAR, C024BE7LR');
    return false;
  }
  
  return true;
}

/**
 * 高速版メイン処理（推奨）
 * 最適化された同期処理を使用
 */
function fastMainProcess() {
  console.log('=== 高速メイン処理開始 ===');
  console.log('実行時刻:', new Date().toLocaleString('ja-JP'));
  const startTime = new Date();
  
  try {
    // 1. 高速同期を実行
    console.log('\n1. 高速Slack同期中...');
    const syncResult = fastSync();
    
    if (!syncResult.success) {
      throw new Error('同期に失敗しました: ' + syncResult.error);
    }
    
    console.log(`同期完了: ${syncResult.messageCount}件のメッセージ (${syncResult.duration}秒)`);
    
    // 2. AI分析を実行（メッセージがある場合のみ）
    if (syncResult.messageCount > 0) {
      console.log('\n2. AI分析実行中...');
      runAIAnalysis();
      
      // 3. 重要な議題を検出して通知
      console.log('\n3. 重要議題の検出と通知...');
      detectAndNotifyImportantTopics();
    } else {
      console.log('新規メッセージなし。AI分析をスキップ');
    }
    
    const endTime = new Date();
    const duration = (endTime - startTime) / 1000;
    console.log(`\n=== 高速メイン処理完了 (処理時間: ${duration}秒) ===`);
    
    return {
      success: true,
      duration: duration,
      messageCount: syncResult.messageCount
    };
    
  } catch (error) {
    console.error('高速メイン処理エラー:', error.toString());
    logError('高速メイン処理', error.toString());
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * 自動実行用メイン関数（トリガー設定用）
 * Slackメッセージ取得 → AI分析 → 通知送信
 */
function mainAutoProcess() {
  console.log('=== 自動処理開始 ===');
  console.log('実行時刻:', new Date().toLocaleString('ja-JP'));
  const startTime = new Date();
  
  try {
    // 1. Slackから最新メッセージを同期
    console.log('\n1. Slackメッセージ同期中...');
    syncSlackMessages();
    
    // 2. AI分析を実行
    console.log('\n2. AI分析実行中...');
    runAIAnalysis();
    
    // 3. 重要な議題を検出して通知
    console.log('\n3. 重要議題の検出と通知...');
    detectAndNotifyImportantTopics();
    
    const endTime = new Date();
    const duration = (endTime - startTime) / 1000;
    console.log(`\n=== 自動処理完了 (処理時間: ${duration}秒) ===`);
    logInfo('自動処理完了', `全処理が正常に完了しました (${duration}秒)`);
    
  } catch (error) {
    console.error('自動処理エラー:', error.toString());
    logError('自動処理', error.toString());
    
    // エラー通知を管理者に送信
    sendErrorNotification(error);
  }
}

/**
 * 高速同期処理（パフォーマンス最適化版）
 * 最小限のAPI呼び出しとバッチ処理で高速化
 */
function fastSync() {
  console.log('=== 高速同期開始 ===');
  const startTime = new Date();
  
  try {
    // スプレッドシートを取得
    let ss;
    if (SPREADSHEET_ID) {
      ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    } else {
      // SPREADSHEET_IDが未設定の場合は現在のスプレッドシートを使用
      ss = SpreadsheetApp.getActiveSpreadsheet();
      if (!ss) {
        throw new Error('スプレッドシートが見つかりません。SPREADSHEET_IDを設定してください。');
      }
    }
    let syncSheet = ss.getSheetByName(SHEETS.SYNC_STATE);  // 正しいシート名
    let slackLogSheet = ss.getSheetByName(SHEETS.SLACK_LOG);
    
    if (!syncSheet) {
      console.log('SyncStateシートが見つかりません。作成します...');
      // SyncStateシートを自動作成（既存のシートがあれば再利用）
      try {
        syncSheet = ss.insertSheet(SHEETS.SYNC_STATE);
        syncSheet.getRange(1, 1, 1, 3).setValues([['チャンネルID', '最終同期タイムスタンプ', '最終同期日時']]);
        syncSheet.getRange(1, 1, 1, 3).setBackground('#4285F4').setFontColor('#FFFFFF').setFontWeight('bold');
        
        // デフォルトチャンネルを追加
        const defaultChannels = ['C09BW2EEVAR', 'C0854FC7S0H'];
        const channelData = defaultChannels.map(id => [id, '0', '']);
        if (channelData.length > 0) {
          syncSheet.getRange(2, 1, channelData.length, 3).setValues(channelData);
        }
      } catch (e) {
        console.log('SyncStateシート作成エラー:', e.toString());
        // 既存のシートを探す
        syncSheet = ss.getSheetByName(SHEETS.SYNC_STATE);
      }
      
      if (!syncSheet) {
        throw new Error('SyncStateシートの作成に失敗しました');
      }
    }
    
    if (!slackLogSheet) {
      console.log('slack_logシートが見つかりません。作成します...');
      slackLogSheet = ss.insertSheet(SHEETS.SLACK_LOG);
      slackLogSheet.getRange(1, 1, 1, 9).setValues([
        ['channel_id', 'channel_name', 'ts', 'thread_ts', 'user_name', 'message', 'date', 'reactions', 'files']
      ]);
    }
    
    // 全チャンネル情報を一度に取得してキャッシュ
    console.log('チャンネル情報を一括取得中...');
    const channelsResponse = slackAPI('conversations.list', {
      types: 'public_channel,private_channel',
      limit: 1000
    });
    
    const channelMap = {};
    if (channelsResponse.channels) {
      channelsResponse.channels.forEach(ch => {
        channelMap[ch.id] = ch.name;
      });
    }
    
    // 同期対象チャンネルを取得
    const syncData = syncSheet.getDataRange().getValues();
    const allMessageBatch = [];
    let totalMessages = 0;
    
    // ヘッダー行をスキップ
    for (let i = 1; i < syncData.length && i <= 3; i++) {  // 最大3チャンネルに制限
      const channelId = syncData[i][0];
      if (!channelId) continue;
      
      const channelName = channelMap[channelId] || channelId;
      const lastTs = syncData[i][1] || '0';
      
      console.log(`チャンネル ${channelName} を処理中...`);
      
      try {
        // メッセージを取得（最大30件）
        const response = slackAPI('conversations.history', {
          channel: channelId,
          oldest: String(lastTs),
          limit: 30,
          inclusive: false
        });
        
        const messages = response.messages || [];
        if (messages.length === 0) {
          console.log(`${channelName}: 新規メッセージなし`);
          continue;
        }
        
        console.log(`${channelName}: ${messages.length}件のメッセージ`);
        
        // メッセージデータを準備（最小限の処理）
        messages.forEach(msg => {
          const messageDate = new Date(Number(msg.ts.split('.')[0]) * 1000);
          allMessageBatch.push([
            channelId,
            channelName,
            msg.ts,
            msg.thread_ts || '',
            msg.user || 'unknown',
            msg.text || '',
            messageDate,
            '',  // reactions（スキップ）
            ''   // files（スキップ）
          ]);
        });
        
        totalMessages += messages.length;
        
        // 最新タイムスタンプを更新
        if (messages.length > 0) {
          syncSheet.getRange(i + 1, 2).setValue(messages[0].ts);
          syncSheet.getRange(i + 1, 3).setValue(new Date());
        }
        
      } catch (error) {
        console.error(`${channelName} エラー: ${error.toString()}`);
      }
    }
    
    // 全メッセージを一括で保存
    if (allMessageBatch.length > 0) {
      console.log(`${allMessageBatch.length}件のメッセージを一括保存中...`);
      const lastRow = slackLogSheet.getLastRow();
      slackLogSheet.getRange(lastRow + 1, 1, allMessageBatch.length, 9).setValues(allMessageBatch);
    }
    
    const duration = (Date.now() - startTime) / 1000;
    console.log(`=== 高速同期完了: ${duration}秒, ${totalMessages}件 ===`);
    
    return {
      success: true,
      messageCount: totalMessages,
      duration: duration
    };
    
  } catch (error) {
    console.error('高速同期エラー:', error.toString());
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * シンプルな同期処理（エラーを最小化）
 * スレッド返信の取得をスキップし、基本的なメッセージのみを同期
 */
function simplifiedSync() {
  console.log('=== シンプル同期開始 ===');
  const startTime = new Date();
  
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const syncSheet = ss.getSheetByName(SHEETS.SYNC_STATE);  // 正しいシート名
    const messagesSheet = ss.getSheetByName(SHEETS.MESSAGES);
    const slackLogSheet = ss.getSheetByName(SHEETS.SLACK_LOG);
    
    const data = syncSheet.getDataRange().getValues();
    let totalMessageCount = 0;
    
    // ヘッダー行をスキップ
    for (let i = 1; i < data.length; i++) {
      const channelId = data[i][0];
      if (!channelId) continue;
      
      try {
        console.log(`\nチャンネル同期: ${channelId}`);
        
        // 最終同期タイムスタンプを取得
        const lastTs = data[i][1] || '0';
        
        // メッセージ履歴を取得（最大50件に制限）
        const response = slackAPI('conversations.history', {
          channel: channelId,
          oldest: lastTs,
          limit: 50,
          inclusive: false
        });
        
        const messages = response.messages || [];
        if (messages.length === 0) {
          console.log('新規メッセージなし');
          continue;
        }
        
        console.log(`${messages.length}件の新規メッセージ`);
        
        // チャンネル名を取得（キャッシュから）
        const channelInfo = getChannelInfo(channelId);
        const channelName = channelInfo?.name || channelId;
        
        // バッチデータ準備
        const messageBatch = [];
        const slackLogBatch = [];
        
        messages.forEach(message => {
          // シンプルなメッセージデータのみ処理
          const enrichedMessage = {
            ...message,
            user_name: message.user || 'unknown',
            reactions: message.reactions ? 
              message.reactions.map(r => `${r.name}:${r.count}`).join(', ') : '',
            files: message.files ? 
              message.files.map(f => f.name || f.title).join(', ') : ''
          };
          
          messageBatch.push(prepareMessageRow(channelId, enrichedMessage));
          slackLogBatch.push(prepareSlackLogRow(channelId, channelName, enrichedMessage));
        });
        
        // バッチ保存
        if (messageBatch.length > 0) {
          saveMessagesBatch(messagesSheet, messageBatch);
          saveSlackLogBatch(slackLogSheet, slackLogBatch);
          totalMessageCount += messageBatch.length;
        }
        
        // 最終同期時刻を更新（最新メッセージのタイムスタンプ）
        if (messages.length > 0) {
          const latestTs = messages[0].ts;  // Slack APIは新しい順で返す
          syncSheet.getRange(i + 1, 2).setValue(latestTs);
          syncSheet.getRange(i + 1, 3).setValue(new Date());
        }
        
      } catch (error) {
        console.error(`チャンネル ${channelId} エラー: ${error.toString()}`);
      }
      
      // Rate limit対策
      Utilities.sleep(500);
    }
    
    const endTime = new Date();
    const duration = (endTime - startTime) / 1000;
    console.log(`\n=== シンプル同期完了 ===`);
    console.log(`処理時間: ${duration}秒`);
    console.log(`同期メッセージ数: ${totalMessageCount}件`);
    
    return {
      success: true,
      messageCount: totalMessageCount,
      duration: duration
    };
    
  } catch (error) {
    console.error('同期エラー:', error.toString());
    return {
      success: false,
      error: error.toString()
    };
  }
}

// ========= テスト用関数 =========
/**
 * 限定的な同期テスト（最初のチャンネルの最新10件のみ）
 * パフォーマンス問題のデバッグ用
 */
function testLimitedSync() {
  console.log('=== 限定同期テスト開始 ===');
  const startTime = new Date();
  
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const syncSheet = ss.getSheetByName(SHEETS.SYNC_STATE);  // 正しいシート名
    const messagesSheet = ss.getSheetByName(SHEETS.MESSAGES);
    const slackLogSheet = ss.getSheetByName(SHEETS.SLACK_LOG);
    
    // 最初のチャンネルのみテスト
    const lastRow = syncSheet.getLastRow();
    if (lastRow < 2) {
      console.log('同期するチャンネルがありません');
      return;
    }
    
    const channelId = syncSheet.getRange(2, 1).getValue();
    if (!channelId) {
      console.log('チャンネルIDが取得できません');
      return;
    }
    
    console.log(`テストチャンネル: ${channelId}`);
    
    // チャンネル情報を取得
    const channelInfo = getChannelInfo(channelId);
    const channelName = channelInfo?.name || channelId;
    console.log(`チャンネル名: ${channelName}`);
    
    // 最新10件のメッセージを取得
    console.log('メッセージ取得中...');
    const response = slackAPI('conversations.history', {
      channel: channelId,
      limit: 10
    });
    
    const messages = response.messages || [];
    console.log(`取得メッセージ数: ${messages.length}`);
    
    // バッチ処理用データ準備
    const messageBatch = [];
    const slackLogBatch = [];
    
    messages.forEach((message, index) => {
      console.log(`処理中 ${index + 1}/${messages.length}: ts=${message.ts.substring(0, 10)}`);
      
      // ユーザー情報は簡略化
      const userInfo = { name: message.user || 'unknown' };
      
      // リアクション情報を整形
      const reactions = message.reactions ? 
        message.reactions.map(r => `${r.name}:${r.count}`).join(', ') : '';
      
      // ファイル情報を整形
      const files = message.files ? 
        message.files.map(f => f.name || f.title).join(', ') : '';
      
      const enrichedMessage = {
        ...message,
        user_name: userInfo.name,
        reactions: reactions,
        files: files
      };
      
      messageBatch.push(prepareMessageRow(channelId, enrichedMessage));
      slackLogBatch.push(prepareSlackLogRow(channelId, channelName, enrichedMessage));
    });
    
    // バッチ保存
    console.log('バッチ保存開始...');
    saveMessagesBatch(messagesSheet, messageBatch);
    saveSlackLogBatch(slackLogSheet, slackLogBatch);
    
    const endTime = new Date();
    const duration = (endTime - startTime) / 1000;
    console.log(`=== テスト完了 (処理時間: ${duration}秒) ===`);
    
    return {
      success: true,
      channelId: channelId,
      channelName: channelName,
      messageCount: messages.length,
      duration: duration
    };
    
  } catch (error) {
    console.error('テストエラー:', error.toString());
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * 重要な議題を検出して通知
 */
function detectAndNotifyImportantTopics() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const messagesSheet = ss.getSheetByName(SHEETS.MESSAGES);
  const configSheet = ss.getSheetByName(SHEETS.CONFIG);
  
  if (!messagesSheet) {
    console.log('Messagesシートが見つかりません');
    return;
  }
  
  const config = getConfigData(configSheet);
  const data = messagesSheet.getDataRange().getValues();
  
  // 重要度の高いメッセージを抽出（最新24時間以内）
  const now = new Date();
  const oneDayAgo = new Date(now.getTime() - 24 * 60 * 60 * 1000);
  const importantMessages = [];
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const createdAt = row[14]; // created_at列
    const matchFlag = row[8]; // match_flag列
    const humanJudgement = row[9]; // human_judgement列
    const classificationJson = row[7]; // classification_json列
    
    // 最近のメッセージで、重要フラグがあるものを抽出
    if (createdAt && new Date(createdAt) > oneDayAgo) {
      if (matchFlag === '高' || humanJudgement === '必要') {
        let classification = [];
        try {
          classification = JSON.parse(classificationJson || '[]');
        } catch (e) {
          // JSONパースエラーは無視
        }
        
        // スコアが0.7以上のカテゴリがある場合
        const highScoreCategory = classification.find(c => c.score >= 0.7);
        if (highScoreCategory) {
          importantMessages.push({
            id: row[0],
            channelId: row[1],
            text: row[4],
            summary: row[6],
            category: highScoreCategory.category,
            score: highScoreCategory.score,
            permalink: row[10],
            createdAt: createdAt
          });
        }
      }
    }
  }
  
  // 重要なメッセージがある場合は通知
  if (importantMessages.length > 0) {
    console.log(`重要な議題を${importantMessages.length}件検出`);
    
    // Slack通知
    if (config.notifySlackChannel) {
      sendSlackNotification(config.notifySlackChannel, importantMessages);
    }
    
    // メール通知
    if (config.notifyEmails && config.notifyEmails.length > 0) {
      sendEmailNotification(config.notifyEmails, importantMessages);
    }
  } else {
    console.log('重要な議題は検出されませんでした');
  }
}

/**
 * エラー通知を管理者に送信
 */
function sendErrorNotification(error) {
  const adminEmail = PropertiesService.getScriptProperties().getProperty('ADMIN_EMAIL');
  if (!adminEmail) {
    console.log('管理者メールアドレスが設定されていません');
    return;
  }
  
  const subject = '[エラー] Slack議題生成システム - 自動処理エラー';
  const body = `
自動処理でエラーが発生しました。

発生時刻: ${new Date().toLocaleString('ja-JP')}
エラー内容:
${error.toString()}

スタックトレース:
${error.stack || 'なし'}

システムログを確認してください。
  `;
  
  try {
    MailApp.sendEmail(adminEmail, subject, body);
    console.log('エラー通知を管理者に送信しました');
  } catch (mailError) {
    console.error('エラー通知の送信に失敗:', mailError.toString());
  }
}

/**
 * トリガーの設定（手動実行用）
 */
function setupTriggers() {
  const ui = SpreadsheetApp.getUi();
  
  // 既存のトリガーを削除
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'mainAutoProcess') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  // 新しいトリガーを設定
  try {
    // 毎日朝9時に実行
    ScriptApp.newTrigger('mainAutoProcess')
      .timeBased()
      .everyDays(1)
      .atHour(9)
      .create();
    
    // 6時間ごとに実行（オプション）
    // ScriptApp.newTrigger('mainAutoProcess')
    //   .timeBased()
    //   .everyHours(6)
    //   .create();
    
    ui.alert('トリガー設定完了', '毎日朝9時に自動実行されるようになりました。', ui.ButtonSet.OK);
    console.log('トリガーを設定しました');
  } catch (error) {
    ui.alert('トリガー設定エラー', error.toString(), ui.ButtonSet.OK);
    console.error('トリガー設定エラー:', error.toString());
  }
}

/**
 * OpenAI APIテスト
 */
function testOpenAI() {
  const testMessage = "これはテストメッセージです。今日の会議で新製品の開発について話し合いました。";
  
  console.log('=== OpenAI API テスト開始 ===');
  console.log('APIキー存在確認:', OPENAI_API_KEY ? '設定済み' : '未設定');
  console.log('APIキー長さ:', OPENAI_API_KEY ? OPENAI_API_KEY.length : 0);
  console.log('APIキー先頭:', OPENAI_API_KEY ? OPENAI_API_KEY.substring(0, 10) + '...' : 'なし');
  
  // 直接APIを呼び出してレスポンスを確認
  try {
    console.log('\n1. シンプルなAPI呼び出しテスト');
    const simpleResponse = callOpenAIForAgenda([
      { role: 'user', content: 'Say "Hello World" in JSON format with a field called "message"' }
    ], 'gpt-5');
    
    console.log('シンプルレスポンス:', simpleResponse);
    console.log('レスポンスタイプ:', typeof simpleResponse);
    
    console.log('\n2. JSON形式指定でのAPI呼び出しテスト');
    const jsonResponse = callOpenAIForAgenda([
      { role: 'system', content: 'You must respond in valid JSON format only.' },
      { role: 'user', content: 'Create a JSON object with fields: status (set to "ok") and timestamp (current time)' }
    ], 'gpt-5', { type: 'json_object' });
    
    console.log('JSONレスポンス:', jsonResponse);
    console.log('JSONレスポンスタイプ:', typeof jsonResponse);
    
    console.log('\n3. 要約機能テスト');
    const summary = summarizeMessage(testMessage);
    console.log('要約結果:', JSON.stringify(summary, null, 2));
    
    SpreadsheetApp.getUi().alert('テスト完了', 'OpenAI APIのテストが完了しました。\nログ（Ctrl+Enter）を確認してください。', SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (error) {
    console.error('テストエラー:', error.toString());
    console.error('エラースタック:', error.stack);
    SpreadsheetApp.getUi().alert('テスト失敗', `エラー: ${error.toString()}\n\nログを確認してください。`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * 同期状態をリセット（テスト用）
 */
function resetSyncState(channelId = 'C09BW2EEVAR') {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const syncSheet = ss.getSheetByName(SHEETS.SYNC_STATE);
  
  if (!syncSheet) {
    SpreadsheetApp.getUi().alert('SyncStateシートが見つかりません');
    return;
  }
  
  const data = syncSheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === channelId) {
      // タイムスタンプを0にリセット
      syncSheet.getRange(i + 1, 2).setValue('0');
      syncSheet.getRange(i + 1, 3).setValue('リセット済み');
      syncSheet.getRange(i + 1, 4).setValue(0);
      syncSheet.getRange(i + 1, 5).setValue('reset');
      
      SpreadsheetApp.getUi().alert('同期状態をリセットしました', `チャンネル ${channelId} の同期状態をリセットしました。`, SpreadsheetApp.getUi().ButtonSet.OK);
      console.log(`チャンネル ${channelId} の同期状態をリセット`);
      return;
    }
  }
  
  SpreadsheetApp.getUi().alert('チャンネルが見つかりません', `チャンネル ${channelId} の同期状態が見つかりません。`, SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * メッセージ取得のデバッグ
 */
function debugMessageFetch(channelId = 'C09BW2EEVAR') {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const syncSheet = ss.getSheetByName(SHEETS.SYNC_STATE);
  
  console.log(`=== メッセージ取得デバッグ: ${channelId} ===`);
  
  let debugInfo = {
    channelId: channelId,
    tests: []
  };
  
  // 1. 最終同期時刻を確認
  const lastSync = getLastSyncTime(syncSheet, channelId);
  console.log(`最終同期タイムスタンプ: ${lastSync}`);
  debugInfo.lastSync = lastSync;
  
  // 2. 最新メッセージを取得（制限なし）
  try {
    console.log('\n=== 最新メッセージ取得（制限なし） ===');
    const allMessages = slackAPI('conversations.history', {
      channel: channelId,
      limit: 10
    });
    
    debugInfo.tests.push({
      test: '最新メッセージ（制限なし）',
      success: true,
      messageCount: allMessages.messages?.length || 0,
      messages: allMessages.messages?.map(m => ({
        ts: m.ts,
        text: m.text?.substring(0, 50),
        user: m.user,
        date: new Date(parseFloat(m.ts) * 1000).toLocaleString('ja-JP')
      }))
    });
    
    console.log(`取得数: ${allMessages.messages?.length || 0}`);
    if (allMessages.messages?.length > 0) {
      console.log(`最新: ${allMessages.messages[0].ts} - ${new Date(parseFloat(allMessages.messages[0].ts) * 1000).toLocaleString('ja-JP')}`);
      console.log(`最古: ${allMessages.messages[allMessages.messages.length - 1].ts}`);
    }
  } catch (error) {
    debugInfo.tests.push({
      test: '最新メッセージ（制限なし）',
      success: false,
      error: error.toString()
    });
  }
  
  // 3. 最終同期以降のメッセージを取得
  if (lastSync && lastSync !== '0') {
    try {
      console.log(`\n=== 最終同期以降のメッセージ取得 (oldest=${lastSync}) ===`);
      const newMessages = slackAPI('conversations.history', {
        channel: channelId,
        oldest: lastSync,
        limit: 100,
        inclusive: false
      });
      
      debugInfo.tests.push({
        test: '最終同期以降（inclusive=false）',
        success: true,
        messageCount: newMessages.messages?.length || 0,
        oldest: lastSync,
        messages: newMessages.messages?.map(m => ({
          ts: m.ts,
          text: m.text?.substring(0, 50),
          date: new Date(parseFloat(m.ts) * 1000).toLocaleString('ja-JP')
        }))
      });
      
      console.log(`取得数: ${newMessages.messages?.length || 0}`);
    } catch (error) {
      debugInfo.tests.push({
        test: '最終同期以降（inclusive=false）',
        success: false,
        error: error.toString()
      });
    }
    
    // inclusive=trueでも試す
    try {
      console.log(`\n=== inclusive=trueで再試行 ===`);
      const newMessagesInclusive = slackAPI('conversations.history', {
        channel: channelId,
        oldest: lastSync,
        limit: 100,
        inclusive: true
      });
      
      debugInfo.tests.push({
        test: '最終同期以降（inclusive=true）',
        success: true,
        messageCount: newMessagesInclusive.messages?.length || 0,
        messages: newMessagesInclusive.messages?.slice(0, 3).map(m => ({
          ts: m.ts,
          text: m.text?.substring(0, 50),
          date: new Date(parseFloat(m.ts) * 1000).toLocaleString('ja-JP')
        }))
      });
      
      console.log(`取得数: ${newMessagesInclusive.messages?.length || 0}`);
    } catch (error) {
      debugInfo.tests.push({
        test: '最終同期以降（inclusive=true）',
        success: false,
        error: error.toString()
      });
    }
  }
  
  // 4. SyncStateシートの状態を確認
  if (syncSheet) {
    const syncData = syncSheet.getDataRange().getValues();
    debugInfo.syncState = [];
    for (let i = 1; i < syncData.length; i++) {
      if (syncData[i][0] === channelId) {
        debugInfo.syncState.push({
          channelId: syncData[i][0],
          lastSyncTs: syncData[i][1],
          lastSyncDatetime: syncData[i][2],
          messageCount: syncData[i][3],
          status: syncData[i][4]
        });
      }
    }
  }
  
  // 結果表示
  let message = `=== メッセージ取得デバッグ結果 ===\n\n`;
  message += `チャンネル: ${channelId}\n`;
  message += `最終同期: ${lastSync || 'なし'}\n\n`;
  
  debugInfo.tests.forEach(test => {
    message += `【${test.test}】\n`;
    if (test.success) {
      message += `✅ 成功\n`;
      message += `メッセージ数: ${test.messageCount}\n`;
      if (test.messages && test.messages.length > 0) {
        message += `最新3件:\n`;
        test.messages.slice(0, 3).forEach(m => {
          message += `  - ${m.date}: ${m.text}\n`;
        });
      }
    } else {
      message += `❌ 失敗: ${test.error}\n`;
    }
    message += '\n';
  });
  
  if (debugInfo.syncState && debugInfo.syncState.length > 0) {
    message += `【同期状態】\n`;
    debugInfo.syncState.forEach(s => {
      message += `最終同期: ${s.lastSyncTs}\n`;
      message += `日時: ${s.lastSyncDatetime}\n`;
      message += `件数: ${s.messageCount}\n`;
    });
  }
  
  ui.alert('メッセージ取得デバッグ', message, ui.ButtonSet.OK);
  return debugInfo;
}

/**
 * 特定チャンネルのデバッグ診断
 */
function debugChannelAccess(channelId = 'C09BW2EEVAR') {
  const ui = SpreadsheetApp.getUi();
  console.log(`=== チャンネル ${channelId} の詳細デバッグ開始 ===`);
  
  let debugInfo = {
    channelId: channelId,
    tests: [],
    rawResponses: {}
  };
  
  // Bot情報を取得
  const botInfo = checkBotPermissions();
  debugInfo.botInfo = botInfo;
  
  // 1. conversations.listでチャンネルを検索
  console.log('\n1. conversations.listでチャンネルを検索...');
  try {
    const listResponse = slackAPI('conversations.list', {
      types: 'public_channel,private_channel',
      limit: 1000
    });
    
    const foundChannels = listResponse.channels?.filter(ch => ch.id === channelId) || [];
    debugInfo.tests.push({
      test: 'conversations.list',
      success: true,
      found: foundChannels.length,
      channels: foundChannels.map(ch => ({
        id: ch.id,
        name: ch.name,
        is_member: ch.is_member,
        is_private: ch.is_private,
        is_archived: ch.is_archived
      }))
    });
    
    debugInfo.rawResponses.list = foundChannels;
    console.log(`✅ 見つかったチャンネル数: ${foundChannels.length}`);
    foundChannels.forEach(ch => {
      console.log(`  - ${ch.name}: is_member=${ch.is_member}, is_private=${ch.is_private}`);
    });
  } catch (error) {
    debugInfo.tests.push({
      test: 'conversations.list',
      success: false,
      error: error.toString()
    });
    console.error(`❌ conversations.list失敗: ${error.toString()}`);
  }
  
  // 2. conversations.infoで直接取得
  console.log('\n2. conversations.infoで直接取得...');
  try {
    const infoResponse = slackAPI('conversations.info', {
      channel: channelId,
      include_num_members: true
    });
    
    debugInfo.tests.push({
      test: 'conversations.info',
      success: true,
      channel: {
        id: infoResponse.channel.id,
        name: infoResponse.channel.name,
        is_member: infoResponse.channel.is_member,
        is_private: infoResponse.channel.is_private,
        is_archived: infoResponse.channel.is_archived,
        num_members: infoResponse.channel.num_members
      }
    });
    
    debugInfo.rawResponses.info = infoResponse.channel;
    console.log(`✅ チャンネル情報取得成功`);
    console.log(`  - name: ${infoResponse.channel.name}`);
    console.log(`  - is_member: ${infoResponse.channel.is_member}`);
    console.log(`  - is_private: ${infoResponse.channel.is_private}`);
    console.log(`  - num_members: ${infoResponse.channel.num_members}`);
  } catch (error) {
    debugInfo.tests.push({
      test: 'conversations.info',
      success: false,
      error: error.toString()
    });
    console.error(`❌ conversations.info失敗: ${error.toString()}`);
  }
  
  // 3. conversations.membersでメンバーリストを取得
  console.log('\n3. conversations.membersでメンバーリスト取得...');
  try {
    const membersResponse = slackAPI('conversations.members', {
      channel: channelId,
      limit: 100
    });
    
    const botInList = membersResponse.members?.includes(botInfo.userId) || 
                      membersResponse.members?.includes(botInfo.botId);
    
    debugInfo.tests.push({
      test: 'conversations.members',
      success: true,
      totalMembers: membersResponse.members?.length || 0,
      botInMembersList: botInList,
      botUserId: botInfo.userId,
      botId: botInfo.botId,
      first5Members: membersResponse.members?.slice(0, 5)
    });
    
    console.log(`✅ メンバーリスト取得成功`);
    console.log(`  - 総メンバー数: ${membersResponse.members?.length || 0}`);
    console.log(`  - Bot (${botInfo.userId}) はリストに含まれる: ${botInList}`);
  } catch (error) {
    debugInfo.tests.push({
      test: 'conversations.members',
      success: false,
      error: error.toString()
    });
    console.error(`❌ conversations.members失敗: ${error.toString()}`);
  }
  
  // 4. conversations.historyでメッセージ履歴を取得
  console.log('\n4. conversations.historyでメッセージ履歴取得...');
  try {
    const historyResponse = slackAPI('conversations.history', {
      channel: channelId,
      limit: 1
    });
    
    debugInfo.tests.push({
      test: 'conversations.history',
      success: true,
      hasMessages: historyResponse.messages?.length > 0,
      messageCount: historyResponse.messages?.length || 0
    });
    
    console.log(`✅ メッセージ履歴取得成功`);
    console.log(`  - メッセージ数: ${historyResponse.messages?.length || 0}`);
  } catch (error) {
    debugInfo.tests.push({
      test: 'conversations.history',
      success: false,
      error: error.toString()
    });
    console.error(`❌ conversations.history失敗: ${error.toString()}`);
  }
  
  // 5. 結果の分析と表示
  let message = `=== チャンネル ${channelId} デバッグ結果 ===\n\n`;
  message += `Bot情報:\n`;
  message += `  名前: ${botInfo.name}\n`;
  message += `  User ID: ${botInfo.userId}\n`;
  message += `  Bot ID: ${botInfo.botId}\n\n`;
  
  message += `テスト結果:\n`;
  debugInfo.tests.forEach(test => {
    message += `\n【${test.test}】\n`;
    if (test.success) {
      message += `  ✅ 成功\n`;
      Object.entries(test).forEach(([key, value]) => {
        if (key !== 'test' && key !== 'success') {
          message += `  ${key}: ${JSON.stringify(value)}\n`;
        }
      });
    } else {
      message += `  ❌ 失敗: ${test.error}\n`;
    }
  });
  
  // 不整合の検出
  message += `\n=== 診断結果 ===\n`;
  
  const listTest = debugInfo.tests.find(t => t.test === 'conversations.list');
  const infoTest = debugInfo.tests.find(t => t.test === 'conversations.info');
  const membersTest = debugInfo.tests.find(t => t.test === 'conversations.members');
  
  if (listTest?.success && infoTest?.success) {
    const listMember = listTest.channels[0]?.is_member;
    const infoMember = infoTest.channel?.is_member;
    
    if (listMember !== infoMember) {
      message += `⚠️ 不整合検出: conversations.listでは is_member=${listMember}、conversations.infoでは is_member=${infoMember}\n`;
    }
    
    if (membersTest?.success) {
      if (infoMember && !membersTest.botInMembersList) {
        message += `⚠️ 不整合検出: is_member=trueだが、メンバーリストにBotが含まれていません\n`;
      } else if (!infoMember && membersTest.botInMembersList) {
        message += `⚠️ 不整合検出: is_member=falseだが、メンバーリストにBotが含まれています\n`;
      }
    }
  }
  
  // 推奨アクション
  message += `\n=== 推奨アクション ===\n`;
  if (infoTest?.success && !infoTest.channel?.is_member) {
    message += `1. Slackで /invite @${botInfo.name} を実行\n`;
    message += `2. 数秒待ってから再度このテストを実行\n`;
  } else if (infoTest?.error?.includes('invalid_arguments')) {
    message += `1. チャンネルがプライベートチャンネルの可能性があります\n`;
    message += `2. Slackで /invite @${botInfo.name} を実行\n`;
    message += `3. Botを一度削除して再度招待することも試してください\n`;
  }
  
  ui.alert('チャンネルデバッグ結果', message, ui.ButtonSet.OK);
  
  // コンソールにRAWレスポンスを出力
  console.log('\n=== RAWレスポンス ===');
  console.log('conversations.list response:', JSON.stringify(debugInfo.rawResponses.list, null, 2));
  console.log('conversations.info response:', JSON.stringify(debugInfo.rawResponses.info, null, 2));
  
  return debugInfo;
}

/**
 * Bot権限の詳細診断
 */
function diagnoseBotPermissions() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const configSheet = ss.getSheetByName(SHEETS.CONFIG);
  const config = getConfigData(configSheet);
  
  console.log('=== Bot権限詳細診断開始 ===');
  
  let diagnosticInfo = {
    token: {},
    permissions: {},
    channels: {},
    errors: []
  };
  
  // 1. トークン情報の確認
  try {
    const authTest = slackAPI('auth.test', {});
    diagnosticInfo.token = {
      valid: true,
      type: authTest.bot_id ? 'Bot User Token' : 'User Token',
      user: authTest.user,
      userId: authTest.user_id,
      botId: authTest.bot_id || 'N/A',
      team: authTest.team,
      teamId: authTest.team_id,
      url: authTest.url
    };
    console.log('✅ トークン認証成功');
    console.log(`トークンタイプ: ${diagnosticInfo.token.type}`);
    console.log(`Bot ID: ${diagnosticInfo.token.botId}`);
  } catch (error) {
    diagnosticInfo.token = {
      valid: false,
      error: error.toString()
    };
    diagnosticInfo.errors.push('トークン認証失敗: ' + error.toString());
  }
  
  // 2. 必要なスコープの確認（Bot Tokenの場合）
  const requiredScopes = [
    'channels:read',
    'channels:history', 
    'groups:read',
    'groups:history',
    'chat:write',
    'users:read'
  ];
  
  diagnosticInfo.permissions.requiredScopes = requiredScopes;
  diagnosticInfo.permissions.note = 'スコープはSlack Appの設定ページで確認・追加できます';
  
  // 3. 各チャンネルの詳細診断
  if (config.targetChannels && config.targetChannels.length > 0) {
    diagnosticInfo.channels.tested = [];
    
    config.targetChannels.forEach(channelId => {
      let channelDiag = {
        id: channelId,
        name: 'Unknown',
        tests: {}
      };
      
      // conversations.infoテスト
      try {
        const info = slackAPI('conversations.info', {
          channel: channelId,
          include_num_members: true
        });
        
        channelDiag.name = info.channel.name;
        channelDiag.tests.info = '✅ 取得成功';
        channelDiag.isPrivate = info.channel.is_private;
        channelDiag.isMember = info.channel.is_member;
        channelDiag.isArchived = info.channel.is_archived;
        channelDiag.numMembers = info.channel.num_members || 'N/A';
        
        // メンバー一覧の取得を試みる
        if (channelDiag.isMember) {
          try {
            const members = slackAPI('conversations.members', {
              channel: channelId,
              limit: 100
            });
            channelDiag.tests.members = `✅ メンバー取得成功 (${members.members.length}人)`;
            
            // Bot自身がメンバーリストに含まれているか確認
            const botInMembers = members.members.includes(diagnosticInfo.token.userId) || 
                                members.members.includes(diagnosticInfo.token.botId);
            channelDiag.botInMembersList = botInMembers;
            
            if (!botInMembers && channelDiag.isMember) {
              channelDiag.warning = '⚠️ is_memberはtrueですが、メンバーリストにBotが見つかりません';
            }
          } catch (memberError) {
            channelDiag.tests.members = `❌ メンバー取得失敗: ${memberError.toString()}`;
          }
        }
        
        // メッセージ履歴テスト
        if (channelDiag.isMember) {
          try {
            const history = slackAPI('conversations.history', {
              channel: channelId,
              limit: 1
            });
            channelDiag.tests.history = `✅ 履歴取得成功`;
            channelDiag.hasMessages = history.messages && history.messages.length > 0;
          } catch (historyError) {
            channelDiag.tests.history = `❌ 履歴取得失敗: ${historyError.toString()}`;
          }
        } else {
          channelDiag.tests.history = 'スキップ (メンバーではない)';
        }
        
      } catch (error) {
        channelDiag.tests.info = `❌ 情報取得失敗: ${error.toString()}`;
        
        // エラーの種類を判別
        if (error.toString().includes('channel_not_found')) {
          channelDiag.diagnosis = 'チャンネルが見つかりません';
        } else if (error.toString().includes('invalid_arguments')) {
          channelDiag.diagnosis = 'プライベートチャンネルでBotがメンバーではない';
        } else {
          channelDiag.diagnosis = 'アクセスエラー';
        }
      }
      
      diagnosticInfo.channels.tested.push(channelDiag);
    });
  }
  
  // 4. 診断結果の表示
  let message = '=== Bot権限詳細診断結果 ===\n\n';
  
  // トークン情報
  message += '【トークン情報】\n';
  if (diagnosticInfo.token.valid) {
    message += `✅ 認証成功\n`;
    message += `タイプ: ${diagnosticInfo.token.type}\n`;
    message += `Bot名: ${diagnosticInfo.token.user}\n`;
    message += `Bot ID: ${diagnosticInfo.token.botId}\n`;
    message += `Team: ${diagnosticInfo.token.team}\n\n`;
    
    if (diagnosticInfo.token.type === 'User Token') {
      message += '⚠️ 注意: User Tokenではなく、Bot User Tokenの使用を推奨します\n\n';
    }
  } else {
    message += `❌ 認証失敗: ${diagnosticInfo.token.error}\n\n`;
  }
  
  // 必要なスコープ
  message += '【必要なOAuth スコープ】\n';
  diagnosticInfo.permissions.requiredScopes.forEach(scope => {
    message += `• ${scope}\n`;
  });
  message += `\n${diagnosticInfo.permissions.note}\n\n`;
  
  // チャンネル診断結果
  if (diagnosticInfo.channels.tested) {
    message += '【チャンネル診断結果】\n';
    diagnosticInfo.channels.tested.forEach(ch => {
      message += `\n● ${ch.name} (${ch.id})\n`;
      message += `  タイプ: ${ch.isPrivate ? 'プライベート' : 'パブリック'}\n`;
      message += `  メンバー: ${ch.isMember ? '✅' : '❌'} (API応答)\n`;
      
      if (ch.botInMembersList !== undefined) {
        message += `  メンバーリスト確認: ${ch.botInMembersList ? '✅' : '❌'}\n`;
      }
      
      if (ch.warning) {
        message += `  ${ch.warning}\n`;
      }
      
      if (ch.isArchived) {
        message += `  ⚠️ アーカイブ済み\n`;
      }
      
      if (ch.numMembers) {
        message += `  メンバー数: ${ch.numMembers}\n`;
      }
      
      Object.entries(ch.tests).forEach(([test, result]) => {
        message += `  ${test}: ${result}\n`;
      });
      
      if (ch.diagnosis) {
        message += `  診断: ${ch.diagnosis}\n`;
      }
    });
  }
  
  // 推奨事項
  message += '\n【推奨事項】\n';
  
  const notMemberChannels = diagnosticInfo.channels.tested?.filter(ch => !ch.isMember);
  if (notMemberChannels && notMemberChannels.length > 0) {
    message += `\n${notMemberChannels.length}個のチャンネルへの招待が必要です:\n`;
    notMemberChannels.forEach(ch => {
      message += `1. Slackで #${ch.name} チャンネルを開く\n`;
      message += `2. /invite @${diagnosticInfo.token.user} を実行\n\n`;
    });
  }
  
  if (diagnosticInfo.errors.length > 0) {
    message += '\n【エラー】\n';
    diagnosticInfo.errors.forEach(err => {
      message += `• ${err}\n`;
    });
  }
  
  ui.alert('Bot権限詳細診断', message, ui.ButtonSet.OK);
  
  return diagnosticInfo;
}

/**
 * 通知テスト
 */
function testNotification() {
  const testMatches = [{
    id: 'test:123',
    summary: {
      summary: 'これはテスト通知です'
    },
    topCategory: '取締役会決議事項',
    topScore: 0.85,
    permalink: 'https://slack.com/test'
  }];
  
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const configSheet = ss.getSheetByName(SHEETS.CONFIG);
  const config = getConfigData(configSheet);
  
  if (config.notifySlackChannel) {
    sendSlackNotification(config.notifySlackChannel, testMatches);
    SpreadsheetApp.getUi().alert('Slack通知テストを送信しました');
  } else if (config.notifyEmails && config.notifyEmails.length > 0) {
    sendEmailNotification(config.notifyEmails, testMatches);
    SpreadsheetApp.getUi().alert('メール通知テストを送信しました');
  } else {
    SpreadsheetApp.getUi().alert('通知先が設定されていません。\nConfigシートで通知先を設定してください。');
  }
}

// ========= 複数議案の一括議事録作成 =========

/**
 * 複数議案の一括議事録作成
 */
function generateBatchMinutes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const ui = SpreadsheetApp.getUi();
  
  if (sheet.getName() !== 'Messages') {
    ui.alert('エラー', 'Messagesシートで実行してください', ui.ButtonSet.OK);
    return;
  }
  
  // 選択範囲を取得
  const selection = sheet.getActiveRange();
  if (!selection) {
    ui.alert('エラー', '議事録を作成したい行を選択してください', ui.ButtonSet.OK);
    return;
  }
  
  // 選択された行のデータを取得
  const rows = [];
  for (let i = selection.getRow(); i <= selection.getLastRow(); i++) {
    const row = sheet.getRange(i, 1, 1, sheet.getLastColumn()).getValues()[0];
    // human_judgementが「必要」の行のみ処理
    if (row[9] === '必要') {
      rows.push(row);
    }
  }
  
  if (rows.length === 0) {
    ui.alert('エラー', '人間判定が「必要」の行が選択されていません', ui.ButtonSet.OK);
    return;
  }
  
  // カテゴリごとにグループ化
  const groupedByCategory = {};
  rows.forEach(row => {
    const classification = JSON.parse(row[7] || '[]');
    if (classification.length > 0) {
      const topCategory = classification.reduce((prev, current) => 
        (prev.score > current.score) ? prev : current
      );
      
      if (!groupedByCategory[topCategory.category]) {
        groupedByCategory[topCategory.category] = [];
      }
      groupedByCategory[topCategory.category].push({
        id: row[0],
        summary: JSON.parse(row[6] || '{}'),
        classification: classification,
        text: row[4],
        permalink: row[10]
      });
    }
  });
  
  // 各カテゴリごとに統合議事録を作成
  const createdDocs = [];
  Object.entries(groupedByCategory).forEach(([category, items]) => {
    try {
      const doc = createBatchMinutesDocument(category, items);
      createdDocs.push({
        category: category,
        docUrl: doc.getUrl(),
        itemCount: items.length
      });
      
      // Draftsシートに記録
      const draftsSheet = ss.getSheetByName('Drafts');
      if (draftsSheet) {
        items.forEach(item => {
          saveDraftRecord(draftsSheet, item.id, category, doc.getUrl());
        });
      }
    } catch (error) {
      logError(`Batch minutes for ${category}`, error.toString());
    }
  });
  
  // 結果を表示
  if (createdDocs.length > 0) {
    let resultMessage = '以下の議事録を作成しました：\n\n';
    createdDocs.forEach(doc => {
      resultMessage += `• ${doc.category}（${doc.itemCount}件）\n`;
    });
    ui.alert('完了', resultMessage, ui.ButtonSet.OK);
  } else {
    ui.alert('エラー', '議事録の作成に失敗しました', ui.ButtonSet.OK);
  }
}

/**
 * カテゴリごとの統合議事録ドキュメントを作成
 */
function createBatchMinutesDocument(category, items) {
  const docName = `${category}_統合議事録_${Utilities.formatDate(new Date(), 'JST', 'yyyyMMdd_HHmmss')}`;
  const doc = DocumentApp.create(docName);
  const body = doc.getBody();
  
  // タイトル
  const title = body.appendParagraph(`${category} - 統合議事録`);
  title.setHeading(DocumentApp.ParagraphHeading.HEADING1);
  title.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  
  body.appendParagraph('');
  
  // 基本情報
  body.appendParagraph('基本情報').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph(`作成日時: ${Utilities.formatDate(new Date(), 'JST', 'yyyy年MM月dd日 HH:mm')}`);
  body.appendParagraph(`議案数: ${items.length}件`);
  body.appendParagraph('');
  
  // 議案サマリー
  body.appendParagraph('議案一覧').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  items.forEach((item, index) => {
    body.appendListItem(`${item.summary.summary || '要約なし'}`);
  });
  body.appendParagraph('');
  
  // 各議案の詳細
  body.appendParagraph('議案詳細').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  
  items.forEach((item, index) => {
    // 議案タイトル
    const agendaTitle = body.appendParagraph(`第${index + 1}号議案: ${item.summary.summary || '議案'}`);
    agendaTitle.setHeading(DocumentApp.ParagraphHeading.HEADING3);
    
    // 関係者
    if (item.summary.people && item.summary.people.length > 0) {
      body.appendParagraph(`関係者: ${item.summary.people.join(', ')}`);
    }
    
    // 決定事項
    if (item.summary.decisions && item.summary.decisions.length > 0) {
      body.appendParagraph('【決定事項】').setBold(true);
      item.summary.decisions.forEach(decision => {
        body.appendListItem(decision);
      });
    }
    
    // アクションアイテム
    if (item.summary.action_items && item.summary.action_items.length > 0) {
      body.appendParagraph('【アクションアイテム】').setBold(true);
      item.summary.action_items.forEach(action => {
        body.appendListItem(`${action.task} (担当: ${action.owner}, 期限: ${action.due})`);
      });
    }
    
    // 元のメッセージへのリンク
    if (item.permalink) {
      body.appendParagraph(`Slack: ${item.permalink}`);
    }
    
    body.appendParagraph('');
  });
  
  // 次回への申し送り事項
  body.appendParagraph('次回への申し送り事項').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph('（ここに申し送り事項を記載）');
  
  doc.saveAndClose();
  return doc;
}

// ========= Google Document テンプレート管理機能 =========

/**
 * テンプレート管理の初期設定
 * Templatesシートにサンプルテンプレート情報を追加
 */
function initializeTemplates() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Templates');
  
  if (!sheet) {
    SpreadsheetApp.getUi().alert('Templatesシートが見つかりません');
    return;
  }
  
  // サンプルテンプレート情報
  const templates = [
    ['取締役会議事録', '取締役会決議事項', '', '作成してください', new Date()],
    ['監査等委員会議事録', '監査等委員会決議事項', '', '作成してください', new Date()],
    ['株主総会議事録', '株主総会決議事項', '', '作成してください', new Date()],
    ['臨時取締役会議事録', '取締役会決議事項', '', '作成してください', new Date()],
    ['プロジェクト会議議事録', 'プロジェクト進捗', '', '作成してください', new Date()]
  ];
  
  // 既存データがない場合のみ追加
  if (sheet.getLastRow() <= 1) {
    sheet.getRange(2, 1, templates.length, 5).setValues(templates);
  }
  
  SpreadsheetApp.getUi().alert(
    'テンプレート情報を初期化しました。\n' +
    '各テンプレートのGoogle Doc IDを設定してください。'
  );
}

/**
 * テンプレートドキュメントを新規作成
 */
function createTemplateDocuments() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'テンプレート作成',
    'サンプルのテンプレートドキュメントを作成しますか？\n' +
    '既存のテンプレートがある場合は、そのDocument IDを直接Templatesシートに入力してください。',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) {
    return;
  }
  
  const folder = createTemplateFolder();
  const templates = createSampleTemplates(folder);
  updateTemplateSheet(templates);
  
  ui.alert(
    '✅ テンプレート作成完了',
    `${templates.length}個のテンプレートを作成しました。\n` +
    `フォルダ: ${folder.getUrl()}\n\n` +
    'Templatesシートで確認してください。'
  );
}

/**
 * テンプレート用フォルダを作成
 */
function createTemplateFolder() {
  const folderName = '議事録テンプレート_' + Utilities.formatDate(new Date(), 'JST', 'yyyyMMdd');
  
  // 既存のフォルダを確認
  const folders = DriveApp.getFoldersByName(folderName);
  if (folders.hasNext()) {
    return folders.next();
  }
  
  // 新規作成
  return DriveApp.createFolder(folderName);
}

/**
 * サンプルテンプレートを作成
 */
function createSampleTemplates(folder) {
  const templates = [];
  
  // 1. 取締役会議事録テンプレート
  const boardMeetingDoc = createBoardMeetingTemplate(folder);
  templates.push({
    name: '取締役会議事録',
    category: '取締役会決議事項',
    docId: boardMeetingDoc.getId(),
    url: boardMeetingDoc.getUrl()
  });
  
  // 2. 監査等委員会議事録テンプレート
  const auditCommitteeDoc = createAuditCommitteeTemplate(folder);
  templates.push({
    name: '監査等委員会議事録',
    category: '監査等委員会決議事項',
    docId: auditCommitteeDoc.getId(),
    url: auditCommitteeDoc.getUrl()
  });
  
  // 3. 株主総会議事録テンプレート
  const shareholderMeetingDoc = createShareholderMeetingTemplate(folder);
  templates.push({
    name: '株主総会議事録',
    category: '株主総会決議事項',
    docId: shareholderMeetingDoc.getId(),
    url: shareholderMeetingDoc.getUrl()
  });
  
  // 4. 臨時取締役会議事録テンプレート
  const extraordinaryBoardDoc = createExtraordinaryBoardTemplate(folder);
  templates.push({
    name: '臨時取締役会議事録',
    category: '取締役会決議事項',
    docId: extraordinaryBoardDoc.getId(),
    url: extraordinaryBoardDoc.getUrl()
  });
  
  // 5. プロジェクト会議議事録テンプレート
  const projectMeetingDoc = createProjectMeetingTemplate(folder);
  templates.push({
    name: 'プロジェクト会議議事録',
    category: 'プロジェクト進捗',
    docId: projectMeetingDoc.getId(),
    url: projectMeetingDoc.getUrl()
  });
  
  return templates;
}

/**
 * 取締役会議事録テンプレート作成
 */
function createBoardMeetingTemplate(folder) {
  const doc = DocumentApp.create('【テンプレート】取締役会議事録');
  const body = doc.getBody();
  
  // ヘッダー
  const header = body.appendParagraph('取締役会議事録');
  header.setHeading(DocumentApp.ParagraphHeading.HEADING1);
  header.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  
  body.appendParagraph('');
  
  // 基本情報セクション
  body.appendParagraph('1. 開催日時').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph('{{MEETING_DATE}} {{MEETING_TIME}}');
  body.appendParagraph('');
  
  body.appendParagraph('2. 開催場所').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph('{{MEETING_LOCATION}}');
  body.appendParagraph('');
  
  body.appendParagraph('3. 出席者').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph('取締役：');
  body.appendListItem('{{DIRECTOR_1}}');
  body.appendListItem('{{DIRECTOR_2}}');
  body.appendListItem('{{DIRECTOR_3}}');
  body.appendParagraph('監査等委員：');
  body.appendListItem('{{AUDITOR_1}}');
  body.appendListItem('{{AUDITOR_2}}');
  body.appendParagraph('');
  
  body.appendParagraph('4. 議長').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph('{{CHAIRPERSON}}');
  body.appendParagraph('');
  
  body.appendParagraph('5. 定足数の確認').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph('取締役総数{{TOTAL_DIRECTORS}}名中{{PRESENT_DIRECTORS}}名出席により、定款第{{ARTICLE_NUMBER}}条の定足数を満たすことを確認した。');
  body.appendParagraph('');
  
  body.appendParagraph('6. 議案').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  
  // 議案テンプレート
  body.appendParagraph('第1号議案：{{AGENDA_1_TITLE}}').setHeading(DocumentApp.ParagraphHeading.HEADING3);
  body.appendParagraph('【議案内容】');
  body.appendParagraph('{{AGENDA_1_CONTENT}}');
  body.appendParagraph('【審議経過】');
  body.appendParagraph('{{AGENDA_1_DISCUSSION}}');
  body.appendParagraph('【決議結果】');
  body.appendParagraph('{{AGENDA_1_RESOLUTION}}');
  body.appendParagraph('');
  
  body.appendParagraph('第2号議案：{{AGENDA_2_TITLE}}').setHeading(DocumentApp.ParagraphHeading.HEADING3);
  body.appendParagraph('【議案内容】');
  body.appendParagraph('{{AGENDA_2_CONTENT}}');
  body.appendParagraph('【審議経過】');
  body.appendParagraph('{{AGENDA_2_DISCUSSION}}');
  body.appendParagraph('【決議結果】');
  body.appendParagraph('{{AGENDA_2_RESOLUTION}}');
  body.appendParagraph('');
  
  body.appendParagraph('7. 報告事項').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph('{{REPORT_ITEMS}}');
  body.appendParagraph('');
  
  body.appendParagraph('8. 次回開催予定').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph('{{NEXT_MEETING}}');
  body.appendParagraph('');
  
  body.appendParagraph('9. 閉会').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph('以上をもって本日の議事を終了し、議長は{{CLOSING_TIME}}に閉会を宣言した。');
  body.appendParagraph('');
  
  // 署名欄
  body.appendParagraph('上記の決議を明確にするため、この議事録を作成し、出席取締役全員が記名押印する。');
  body.appendParagraph('');
  body.appendParagraph('{{MEETING_DATE}}');
  body.appendParagraph('');
  body.appendParagraph('{{COMPANY_NAME}}');
  body.appendParagraph('');
  body.appendParagraph('議長　　　{{CHAIRPERSON_NAME}}　　　印');
  body.appendParagraph('');
  body.appendParagraph('取締役　　{{DIRECTOR_NAME_1}}　　　印');
  body.appendParagraph('');
  body.appendParagraph('取締役　　{{DIRECTOR_NAME_2}}　　　印');
  
  // ドキュメントを保存してフォルダに移動
  doc.saveAndClose();
  const file = DriveApp.getFileById(doc.getId());
  folder.addFile(file);
  DriveApp.getRootFolder().removeFile(file);
  
  return doc;
}

/**
 * 監査等委員会議事録テンプレート作成
 */
function createAuditCommitteeTemplate(folder) {
  const doc = DocumentApp.create('【テンプレート】監査等委員会議事録');
  const body = doc.getBody();
  
  // ヘッダー
  const header = body.appendParagraph('監査等委員会議事録');
  header.setHeading(DocumentApp.ParagraphHeading.HEADING1);
  header.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  
  body.appendParagraph('');
  
  body.appendParagraph('1. 開催日時').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph('{{MEETING_DATE}} {{MEETING_TIME}}');
  body.appendParagraph('');
  
  body.appendParagraph('2. 開催場所').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph('{{MEETING_LOCATION}}');
  body.appendParagraph('');
  
  body.appendParagraph('3. 出席者').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph('監査等委員である取締役：');
  body.appendListItem('{{AUDIT_COMMITTEE_MEMBER_1}}（委員長）');
  body.appendListItem('{{AUDIT_COMMITTEE_MEMBER_2}}');
  body.appendListItem('{{AUDIT_COMMITTEE_MEMBER_3}}');
  body.appendParagraph('');
  
  body.appendParagraph('4. 議題').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  
  body.appendParagraph('(1) 監査計画について').setHeading(DocumentApp.ParagraphHeading.HEADING3);
  body.appendParagraph('{{AUDIT_PLAN_CONTENT}}');
  body.appendParagraph('');
  
  body.appendParagraph('(2) 内部統制システムの評価').setHeading(DocumentApp.ParagraphHeading.HEADING3);
  body.appendParagraph('{{INTERNAL_CONTROL_EVALUATION}}');
  body.appendParagraph('');
  
  body.appendParagraph('(3) 会計監査人との連携').setHeading(DocumentApp.ParagraphHeading.HEADING3);
  body.appendParagraph('{{AUDITOR_COORDINATION}}');
  body.appendParagraph('');
  
  body.appendParagraph('5. 審議事項').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph('{{DELIBERATION_ITEMS}}');
  body.appendParagraph('');
  
  body.appendParagraph('6. 決議事項').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph('{{RESOLUTION_ITEMS}}');
  body.appendParagraph('');
  
  body.appendParagraph('7. 次回開催予定').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph('{{NEXT_MEETING}}');
  body.appendParagraph('');
  
  // 署名欄
  body.appendParagraph('上記の通り監査等委員会を開催し、審議の結果を記録するため、この議事録を作成し、出席監査等委員全員が記名押印する。');
  body.appendParagraph('');
  body.appendParagraph('{{MEETING_DATE}}');
  body.appendParagraph('');
  body.appendParagraph('{{COMPANY_NAME}}');
  body.appendParagraph('');
  body.appendParagraph('監査等委員長　　{{COMMITTEE_CHAIR_NAME}}　　　印');
  body.appendParagraph('');
  body.appendParagraph('監査等委員　　　{{COMMITTEE_MEMBER_NAME_1}}　　　印');
  body.appendParagraph('');
  body.appendParagraph('監査等委員　　　{{COMMITTEE_MEMBER_NAME_2}}　　　印');
  
  doc.saveAndClose();
  const file = DriveApp.getFileById(doc.getId());
  folder.addFile(file);
  DriveApp.getRootFolder().removeFile(file);
  
  return doc;
}

/**
 * 株主総会議事録テンプレート作成
 */
function createShareholderMeetingTemplate(folder) {
  const doc = DocumentApp.create('【テンプレート】株主総会議事録');
  const body = doc.getBody();
  
  // ヘッダー
  const header = body.appendParagraph('第{{MEETING_NUMBER}}期 定時株主総会議事録');
  header.setHeading(DocumentApp.ParagraphHeading.HEADING1);
  header.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  
  body.appendParagraph('');
  
  body.appendParagraph('1. 開催日時').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph('{{MEETING_DATE}} {{MEETING_TIME}}');
  body.appendParagraph('');
  
  body.appendParagraph('2. 開催場所').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph('{{MEETING_LOCATION}}');
  body.appendParagraph('');
  
  body.appendParagraph('3. 出席株主').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph('株主総数：{{TOTAL_SHAREHOLDERS}}名');
  body.appendParagraph('発行済株式総数：{{TOTAL_SHARES}}株');
  body.appendParagraph('議決権を有する株式数：{{VOTING_SHARES}}株');
  body.appendParagraph('出席株主数：{{PRESENT_SHAREHOLDERS}}名（委任状による出席{{PROXY_SHAREHOLDERS}}名を含む）');
  body.appendParagraph('出席株主の議決権数：{{PRESENT_VOTING_RIGHTS}}個');
  body.appendParagraph('');
  
  body.appendParagraph('4. 議長').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph('代表取締役社長 {{PRESIDENT_NAME}}');
  body.appendParagraph('');
  
  body.appendParagraph('5. 議事の経過および結果').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  
  body.appendParagraph('第1号議案：{{PROPOSAL_1_TITLE}}').setHeading(DocumentApp.ParagraphHeading.HEADING3);
  body.appendParagraph('議長より、{{PROPOSAL_1_EXPLANATION}}');
  body.appendParagraph('審議の後、議長がその賛否を諮ったところ、{{PROPOSAL_1_RESULT}}');
  body.appendParagraph('');
  
  body.appendParagraph('第2号議案：{{PROPOSAL_2_TITLE}}').setHeading(DocumentApp.ParagraphHeading.HEADING3);
  body.appendParagraph('議長より、{{PROPOSAL_2_EXPLANATION}}');
  body.appendParagraph('審議の後、議長がその賛否を諮ったところ、{{PROPOSAL_2_RESULT}}');
  body.appendParagraph('');
  
  body.appendParagraph('第3号議案：{{PROPOSAL_3_TITLE}}').setHeading(DocumentApp.ParagraphHeading.HEADING3);
  body.appendParagraph('議長より、{{PROPOSAL_3_EXPLANATION}}');
  body.appendParagraph('審議の後、議長がその賛否を諮ったところ、{{PROPOSAL_3_RESULT}}');
  body.appendParagraph('');
  
  body.appendParagraph('6. 報告事項').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph('{{REPORT_ITEMS}}');
  body.appendParagraph('');
  
  body.appendParagraph('7. 閉会').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph('以上をもって本総会の議事を終了したので、議長は{{CLOSING_TIME}}閉会を宣言した。');
  body.appendParagraph('');
  
  // 署名欄
  body.appendParagraph('上記決議を明確にするため、議長および出席取締役が記名押印する。');
  body.appendParagraph('');
  body.appendParagraph('{{MEETING_DATE}}');
  body.appendParagraph('');
  body.appendParagraph('{{COMPANY_NAME}}');
  body.appendParagraph('');
  body.appendParagraph('議長　代表取締役社長　　{{PRESIDENT_NAME}}　　　印');
  body.appendParagraph('');
  body.appendParagraph('取締役　　{{DIRECTOR_NAME_1}}　　　印');
  body.appendParagraph('');
  body.appendParagraph('取締役　　{{DIRECTOR_NAME_2}}　　　印');
  
  doc.saveAndClose();
  const file = DriveApp.getFileById(doc.getId());
  folder.addFile(file);
  DriveApp.getRootFolder().removeFile(file);
  
  return doc;
}

/**
 * 臨時取締役会議事録テンプレート作成
 */
function createExtraordinaryBoardTemplate(folder) {
  const doc = DocumentApp.create('【テンプレート】臨時取締役会議事録');
  const body = doc.getBody();
  
  const header = body.appendParagraph('臨時取締役会議事録');
  header.setHeading(DocumentApp.ParagraphHeading.HEADING1);
  header.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  
  body.appendParagraph('');
  
  body.appendParagraph('1. 招集通知').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph('{{NOTICE_DATE}}付で臨時取締役会の招集通知を発し、全取締役の同意により開催した。');
  body.appendParagraph('');
  
  body.appendParagraph('2. 開催日時').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph('{{MEETING_DATE}} {{MEETING_TIME}}');
  body.appendParagraph('');
  
  body.appendParagraph('3. 開催方法').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph('{{MEETING_METHOD}}');
  body.appendParagraph('');
  
  body.appendParagraph('4. 出席者').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph('{{ATTENDEES}}');
  body.appendParagraph('');
  
  body.appendParagraph('5. 議案').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph('【緊急議案】{{URGENT_AGENDA_TITLE}}');
  body.appendParagraph('');
  body.appendParagraph('【背景】');
  body.appendParagraph('{{BACKGROUND}}');
  body.appendParagraph('');
  body.appendParagraph('【提案内容】');
  body.appendParagraph('{{PROPOSAL_CONTENT}}');
  body.appendParagraph('');
  body.appendParagraph('【審議】');
  body.appendParagraph('{{DELIBERATION}}');
  body.appendParagraph('');
  body.appendParagraph('【決議】');
  body.appendParagraph('{{RESOLUTION}}');
  body.appendParagraph('');
  
  body.appendParagraph('6. 今後の対応').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph('{{NEXT_STEPS}}');
  
  doc.saveAndClose();
  const file = DriveApp.getFileById(doc.getId());
  folder.addFile(file);
  DriveApp.getRootFolder().removeFile(file);
  
  return doc;
}

/**
 * プロジェクト会議議事録テンプレート作成
 */
function createProjectMeetingTemplate(folder) {
  const doc = DocumentApp.create('【テンプレート】プロジェクト会議議事録');
  const body = doc.getBody();
  
  const header = body.appendParagraph('{{PROJECT_NAME}} 会議議事録');
  header.setHeading(DocumentApp.ParagraphHeading.HEADING1);
  header.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  
  body.appendParagraph('');
  
  body.appendParagraph('基本情報').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph('日時：{{MEETING_DATE}} {{MEETING_TIME}}');
  body.appendParagraph('場所：{{MEETING_LOCATION}}');
  body.appendParagraph('参加者：{{PARTICIPANTS}}');
  body.appendParagraph('');
  
  body.appendParagraph('1. 進捗報告').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph('{{PROGRESS_REPORT}}');
  body.appendParagraph('');
  
  body.appendParagraph('2. 課題・リスク').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph('{{ISSUES_AND_RISKS}}');
  body.appendParagraph('');
  
  body.appendParagraph('3. 決定事項').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph('{{DECISIONS}}');
  body.appendParagraph('');
  
  body.appendParagraph('4. アクションアイテム').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph('{{ACTION_ITEMS}}');
  body.appendParagraph('');
  
  body.appendParagraph('5. 次回予定').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph('{{NEXT_MEETING}}');
  
  doc.saveAndClose();
  const file = DriveApp.getFileById(doc.getId());
  folder.addFile(file);
  DriveApp.getRootFolder().removeFile(file);
  
  return doc;
}

/**
 * Templatesシートを更新
 */
function updateTemplateSheet(templates) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Templates');
  
  if (!sheet) {
    return;
  }
  
  templates.forEach(template => {
    // 既存のテンプレートを検索
    const data = sheet.getDataRange().getValues();
    let rowIndex = -1;
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === template.name && data[i][1] === template.category) {
        rowIndex = i + 1;
        break;
      }
    }
    
    if (rowIndex > 0) {
      // 既存のレコードを更新
      sheet.getRange(rowIndex, 3).setValue(template.docId);
      sheet.getRange(rowIndex, 5).setValue(new Date());
    } else {
      // 新規レコードを追加
      const lastRow = sheet.getLastRow();
      sheet.getRange(lastRow + 1, 1, 1, 5).setValues([[
        template.name,
        template.category,
        template.docId,
        getPlaceholdersFromTemplate(template.docId),
        new Date()
      ]]);
    }
  });
}

/**
 * テンプレートからプレースホルダーを抽出
 */
function getPlaceholdersFromTemplate(docId) {
  try {
    const doc = DocumentApp.openById(docId);
    const text = doc.getBody().getText();
    
    // {{PLACEHOLDER}}形式のプレースホルダーを抽出
    const placeholders = text.match(/\{\{[A-Z_0-9]+\}\}/g) || [];
    const uniquePlaceholders = [...new Set(placeholders)];
    
    return uniquePlaceholders.join(', ');
  } catch (error) {
    console.error('テンプレート読み込みエラー:', error);
    return '';
  }
}

/**
 * テンプレートを使用して議事録を生成
 */
function generateMinutesFromTemplate(category, summaryData, messageData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const templatesSheet = ss.getSheetByName('Templates');
  
  if (!templatesSheet) {
    throw new Error('Templatesシートが見つかりません');
  }
  
  // カテゴリに対応するテンプレートを検索
  const templateData = templatesSheet.getDataRange().getValues();
  let templateDocId = null;
  
  for (let i = 1; i < templateData.length; i++) {
    if (templateData[i][1] === category) {
      templateDocId = templateData[i][2];
      break;
    }
  }
  
  if (!templateDocId) {
    console.log(`カテゴリ「${category}」のテンプレートが見つかりません。デフォルト生成を使用します。`);
    return null;
  }
  
  // テンプレートをコピーして新しいドキュメントを作成
  const templateDoc = DriveApp.getFileById(templateDocId);
  const newDocName = `${category}_議事録_${Utilities.formatDate(new Date(), 'JST', 'yyyyMMdd_HHmmss')}`;
  const newDoc = templateDoc.makeCopy(newDocName);
  
  // ドキュメントを開いてプレースホルダーを置換
  const doc = DocumentApp.openById(newDoc.getId());
  const body = doc.getBody();
  
  // プレースホルダーの置換
  replacePlaceholdersInDocument(body, summaryData, messageData);
  
  doc.saveAndClose();
  
  return {
    docId: newDoc.getId(),
    url: newDoc.getUrl(),
    name: newDocName
  };
}

/**
 * プレースホルダーを実際のデータで置換
 */
function replacePlaceholdersInDocument(body, summaryData, messageData) {
  const now = new Date();
  const companyName = getCompanyNameFromConfig();
  
  // 基本的なプレースホルダーの置換
  const basicReplacements = {
    '{{COMPANY_NAME}}': companyName,
    '{{MEETING_DATE}}': Utilities.formatDate(now, 'JST', 'yyyy年MM月dd日'),
    '{{MEETING_TIME}}': Utilities.formatDate(now, 'JST', 'HH時mm分'),
    '{{MEETING_LOCATION}}': 'オンライン会議（Slack）',
    '{{MEETING_METHOD}}': 'Web会議システム',
    '{{MEETING_NUMBER}}': new Date().getFullYear() - 2020  // 仮の期数
  };
  
  // AI分析から抽出した情報で置換
  const aiReplacements = extractAIReplacements(summaryData, messageData);
  
  // すべての置換を実行
  const allReplacements = { ...basicReplacements, ...aiReplacements };
  
  Object.entries(allReplacements).forEach(([placeholder, value]) => {
    body.replaceText(placeholder, value || '（未定）');
  });
  
  // 残ったプレースホルダーをデフォルト値で置換
  const remainingPlaceholders = body.getText().match(/\{\{[A-Z_0-9]+\}\}/g) || [];
  remainingPlaceholders.forEach(placeholder => {
    body.replaceText(placeholder, '（要確認）');
  });
}

/**
 * AI分析結果から置換データを抽出
 */
function extractAIReplacements(summaryData, messageData) {
  const replacements = {};
  
  if (summaryData) {
    // 参加者情報
    if (summaryData.people && summaryData.people.length > 0) {
      replacements['{{PARTICIPANTS}}'] = summaryData.people.join('、');
      replacements['{{CHAIRPERSON}}'] = summaryData.people[0];
      replacements['{{CHAIRPERSON_NAME}}'] = summaryData.people[0];
      replacements['{{PRESIDENT_NAME}}'] = summaryData.people[0];
      
      // 取締役として設定
      summaryData.people.forEach((person, index) => {
        replacements[`{{DIRECTOR_${index + 1}}}`] = person;
        replacements[`{{DIRECTOR_NAME_${index + 1}}}`] = person;
      });
    }
    
    // 決定事項
    if (summaryData.decisions && summaryData.decisions.length > 0) {
      replacements['{{DECISIONS}}'] = summaryData.decisions.map((d, i) => `${i + 1}. ${d}`).join('\n');
      replacements['{{RESOLUTION_ITEMS}}'] = summaryData.decisions.join('\n');
      
      // 各議案として設定
      summaryData.decisions.forEach((decision, index) => {
        replacements[`{{AGENDA_${index + 1}_TITLE}}`] = decision;
        replacements[`{{AGENDA_${index + 1}_CONTENT}}`] = decision;
        replacements[`{{AGENDA_${index + 1}_RESOLUTION}}`] = '全員一致で承認された。';
        replacements[`{{PROPOSAL_${index + 1}_TITLE}}`] = decision;
        replacements[`{{PROPOSAL_${index + 1}_RESULT}}`] = '賛成多数により可決された。';
      });
    }
    
    // アクションアイテム
    if (summaryData.action_items && summaryData.action_items.length > 0) {
      const actionText = summaryData.action_items.map(item => 
        `・${item.task}（担当：${item.owner}、期限：${item.due}）`
      ).join('\n');
      replacements['{{ACTION_ITEMS}}'] = actionText;
      replacements['{{NEXT_STEPS}}'] = actionText;
    }
    
    // 要約
    if (summaryData.summary) {
      replacements['{{BACKGROUND}}'] = summaryData.summary;
      replacements['{{PROPOSAL_CONTENT}}'] = summaryData.summary;
      replacements['{{PROGRESS_REPORT}}'] = summaryData.summary;
    }
  }
  
  // メッセージ内容から議論を抽出
  if (messageData) {
    const discussionText = extractDiscussion(messageData);
    replacements['{{DELIBERATION}}'] = discussionText;
    replacements['{{AGENDA_1_DISCUSSION}}'] = discussionText;
    replacements['{{PROPOSAL_1_EXPLANATION}}'] = discussionText;
  }
  
  // デフォルト値の設定
  replacements['{{TOTAL_DIRECTORS}}'] = '5';
  replacements['{{PRESENT_DIRECTORS}}'] = '5';
  replacements['{{ARTICLE_NUMBER}}'] = '23';
  replacements['{{CLOSING_TIME}}'] = Utilities.formatDate(new Date(), 'JST', 'HH時mm分');
  replacements['{{NEXT_MEETING}}'] = '次回は別途調整の上、開催する。';
  
  return replacements;
}

/**
 * メッセージから議論内容を抽出
 */
function extractDiscussion(messageData) {
  if (!messageData || !messageData.text) {
    return '（Slackでの議論内容を参照）';
  }
  
  // メッセージを整形
  const lines = messageData.text.split('\n').slice(0, 10); // 最初の10行
  return lines.join('\n') + '\n（以下、詳細はSlackログを参照）';
}

/**
 * 会社名を取得
 */
function getCompanyNameFromConfig() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName('Config');
  
  if (configSheet) {
    const data = configSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === 'company') {
        return data[i][1] || '株式会社〇〇';
      }
    }
  }
  
  return '株式会社〇〇';
}

/**
 * テンプレート一覧を取得（UI表示用）
 */
function getTemplateList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Templates');
  
  if (!sheet) {
    return [];
  }
  
  const data = sheet.getDataRange().getValues();
  const templates = [];
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][2]) {
      templates.push({
        name: data[i][0],
        category: data[i][1],
        docId: data[i][2],
        placeholders: data[i][3],
        lastUpdated: data[i][4]
      });
    }
  }
  
  return templates;
}

/**
 * テンプレートをプレビュー（読み取り専用で開く）
 */
function previewTemplate(docId) {
  try {
    const doc = DocumentApp.openById(docId);
    const url = doc.getUrl();
    
    const ui = SpreadsheetApp.getUi();
    ui.alert(
      'テンプレートプレビュー',
      `テンプレートを新しいタブで開きます。\n\n${url}\n\n※編集する場合は直接Google Documentで編集してください。`,
      ui.ButtonSet.OK
    );
    
    return url;
  } catch (error) {
    throw new Error('テンプレートを開けませんでした: ' + error.toString());
  }
}

/**
 * テンプレート一覧を表示
 */
function showTemplateList() {
  const templates = getTemplateList();
  
  if (templates.length === 0) {
    SpreadsheetApp.getUi().alert(
      'テンプレート一覧',
      'テンプレートが登録されていません。\n「テンプレート管理 > テンプレート作成」から作成してください。'
    );
    return;
  }
  
  let message = '📄 登録済みテンプレート一覧\n\n';
  
  templates.forEach((template, index) => {
    message += `${index + 1}. ${template.name}\n`;
    message += `   カテゴリ: ${template.category}\n`;
    message += `   更新日: ${Utilities.formatDate(new Date(template.lastUpdated), 'JST', 'yyyy/MM/dd')}\n`;
    message += `   プレースホルダー数: ${template.placeholders ? template.placeholders.split(',').length : 0}\n\n`;
  });
  
  message += '\n※テンプレートの編集は、Templatesシートから各ドキュメントを開いて行ってください。';
  
  SpreadsheetApp.getUi().alert('テンプレート一覧', message);
}

// ========= ユーティリティ関数群 =========

// ログエントリ作成
function createLogEntry(status, process, detail) {
  return {
    timestamp: new Date().toLocaleString('ja-JP'),
    status: status,
    process: process,
    detail: detail,
    error: status === 'ERROR' ? detail : null
  };
}

// ========= チャンネル一覧診断 =========
function diagnoseChannelAccess() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    console.log('チャンネルアクセス診断開始...');
    
    // 1. まず権限チェック
    console.log('権限チェック中...');
    const authTest = slackAPI('auth.test', {});
    console.log('Bot情報:', authTest);
    
    // 2. Botが参加しているチャンネルを取得
    const joinedChannels = getAllJoinedChannels();
    
    // 3. チャンネルを種類別に分類
    const publicChannels = joinedChannels.filter(ch => !ch.is_private);
    const privateChannels = joinedChannels.filter(ch => ch.is_private === true);
    const appAccessChannels = joinedChannels.filter(ch => ch.app_access);
    
    // 4. 診断結果を作成
    let result = `チャンネルアクセス診断結果：\n\n`;
    result += `🤖 Bot名: ${authTest.user || 'unknown'}\n`;
    result += `📍 ワークスペース: ${authTest.team || 'unknown'}\n`;
    result += `🔧 Bot ID: ${authTest.bot_id || 'なし（統合アプリケーション）'}\n\n`;
    
    result += `✅ アクセス可能なチャンネル総数: ${joinedChannels.length}\n`;
    if (appAccessChannels.length > 0) {
      result += `📱 アプリケーションアクセス: ${appAccessChannels.length}個\n`;
    }
    result += `\n`;
    
    result += `【パブリックチャンネル】${publicChannels.length}個\n`;
    publicChannels.slice(0, 10).forEach((ch, i) => {
      const accessType = ch.app_access ? ' (App統合)' : '';
      result += `  ${i + 1}. #${ch.name} (${ch.id})${accessType}\n`;
    });
    if (publicChannels.length > 10) {
      result += `  ... 他 ${publicChannels.length - 10} チャンネル\n`;
    }
    
    result += `\n【プライベートチャンネル】${privateChannels.length}個\n`;
    if (privateChannels.length === 0) {
      result += `  ⚠️ プライベートチャンネルが表示されない場合:\n`;
      result += `  1. Slack Appに groups:read と groups:history 権限を追加\n`;
      result += `  2. アプリを再インストール（Reinstall to Workspace）\n`;
      result += `  3. プライベートチャンネルでBotを招待 (/invite @bot-name)\n`;
      result += `  4. または、チャンネル設定の「Integrations」からアプリを追加\n`;
    } else {
      privateChannels.slice(0, 10).forEach((ch, i) => {
        const accessType = ch.app_access ? ' (App統合)' : '';
        result += `  ${i + 1}. 🔒${ch.name} (${ch.id})${accessType}\n`;
      });
      if (privateChannels.length > 10) {
        result += `  ... 他 ${privateChannels.length - 10} チャンネル\n`;
      }
    }
    
    result += `\n💡 ヒント:\n`;
    result += `- 対象チャンネルにBotが招待されていない場合は、/invite @bot-name で招待してください\n`;
    result += `- プライベートチャンネルも同様に招待が必要です\n`;
    result += `- groups:read 権限がないとプライベートチャンネルは表示されません\n`;
    result += `- 「Slack全チャンネル同期」を実行すると、上記すべてのチャンネルから情報を取得します\n\n`;
    
    // プライベートチャンネルが0の場合、追加の診断
    if (privateChannels.length === 0) {
      result += `\n🔍 プライベートチャンネル診断:\n`;
      
      // groups:read権限の確認
      try {
        const testPrivate = slackAPI('conversations.list', { types: 'private_channel', limit: 100 });
        if (testPrivate.ok) {
          result += `✅ groups:read権限はあります\n`;
          
          // すべてのプライベートチャンネルを確認
          const allPrivate = testPrivate.channels || [];
          if (allPrivate.length > 0) {
            result += `📊 ワークスペース内のプライベートチャンネル: ${allPrivate.length}個\n`;
            result += `❌ ただし、Botはどのプライベートチャンネルにも参加していません\n`;
            result += `💡 各プライベートチャンネルで /invite @${authTest.user} を実行してください\n`;
          } else {
            result += `📊 ワークスペースにプライベートチャンネルが存在しません\n`;
          }
        }
      } catch (e) {
        result += `❌ groups:read権限がありません: ${e.toString()}\n`;
      }
    }
    
    ui.alert('チャンネルアクセス診断', result, ui.ButtonSet.OK);
    
  } catch (error) {
    console.error('診断エラー:', error);
    ui.alert('エラー', `診断中にエラーが発生しました:\n${error.toString()}\n\n権限不足の可能性があります。`, ui.ButtonSet.OK);
  }
}

// ========= 権限診断 =========
function diagnoseBotScopes() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    console.log('権限診断開始...');
    
    // 1. 認証テスト
    const authTest = slackAPI('auth.test', {});
    
    // 2. 各APIをテストして権限を確認
    const scopeTests = {
      'channels:read': false,
      'channels:history': false,
      'groups:read': false,
      'groups:history': false,
      'users:read': false,
      'chat:write': false
    };
    
    // channels:read テスト
    try {
      const publicChannels = slackAPI('conversations.list', { types: 'public_channel', limit: 1 });
      if (publicChannels.ok) scopeTests['channels:read'] = true;
    } catch (e) {
      console.log('channels:read テスト失敗:', e.toString());
    }
    
    // groups:read テスト（プライベートチャンネル）
    try {
      const privateChannels = slackAPI('conversations.list', { types: 'private_channel', limit: 1 });
      if (privateChannels.ok) {
        scopeTests['groups:read'] = true;
        console.log(`プライベートチャンネル取得成功: ${privateChannels.channels?.length || 0}個`);
      }
    } catch (e) {
      console.log('groups:read テスト失敗:', e.toString());
    }
    
    // groups:history テスト（プライベートチャンネル履歴）
    try {
      // まずプライベートチャンネルを1つ取得
      const privateChannels = slackAPI('conversations.list', { types: 'private_channel', limit: 1 });
      if (privateChannels.ok && privateChannels.channels && privateChannels.channels.length > 0) {
        const privateChannel = privateChannels.channels.find(ch => ch.is_member);
        if (privateChannel) {
          const history = slackAPI('conversations.history', { 
            channel: privateChannel.id, 
            limit: 1 
          });
          if (history.ok) {
            scopeTests['groups:history'] = true;
            console.log(`プライベートチャンネル履歴取得成功: ${privateChannel.name}`);
          }
        }
      }
    } catch (e) {
      console.log('groups:history テスト失敗:', e.toString());
    }
    
    // channels:history テスト（パブリックチャンネル履歴）
    try {
      const publicChannels = slackAPI('conversations.list', { types: 'public_channel', limit: 1 });
      if (publicChannels.ok && publicChannels.channels && publicChannels.channels.length > 0) {
        const publicChannel = publicChannels.channels.find(ch => ch.is_member);
        if (publicChannel) {
          const history = slackAPI('conversations.history', { 
            channel: publicChannel.id, 
            limit: 1 
          });
          if (history.ok) {
            scopeTests['channels:history'] = true;
            console.log(`パブリックチャンネル履歴取得成功: ${publicChannel.name}`);
          }
        }
      }
    } catch (e) {
      console.log('channels:history テスト失敗:', e.toString());
    }
    
    // users:read テスト
    try {
      const users = slackAPI('users.list', { limit: 1 });
      if (users.ok) scopeTests['users:read'] = true;
    } catch (e) {
      console.log('users:read テスト失敗:', e.toString());
    }
    
    // 3. 診断結果を作成
    let result = `権限診断結果：\n\n`;
    result += `🤖 Bot名: ${authTest.user || 'unknown'}\n`;
    result += `📍 ワークスペース: ${authTest.team || 'unknown'}\n\n`;
    
    result += `【必須権限の状態】\n`;
    const requiredScopes = [
      { scope: 'channels:read', desc: 'パブリックチャンネル読み取り' },
      { scope: 'channels:history', desc: 'パブリックチャンネル履歴' },
      { scope: 'groups:read', desc: 'プライベートチャンネル読み取り' },
      { scope: 'groups:history', desc: 'プライベートチャンネル履歴' },
      { scope: 'users:read', desc: 'ユーザー情報読み取り' },
      { scope: 'chat:write', desc: 'メッセージ送信' }
    ];
    
    requiredScopes.forEach(({ scope, desc }) => {
      const status = scopeTests[scope] ? '✅' : '❌';
      result += `${status} ${scope} - ${desc}\n`;
    });
    
    // 不足している権限の警告
    const missingScopes = Object.entries(scopeTests)
      .filter(([scope, hasPermission]) => !hasPermission)
      .map(([scope]) => scope);
    
    if (missingScopes.length > 0) {
      result += `\n⚠️ 不足している権限:\n`;
      missingScopes.forEach(scope => {
        result += `  - ${scope}\n`;
      });
      result += `\n対処法:\n`;
      result += `1. https://api.slack.com/apps でアプリを開く\n`;
      result += `2. "OAuth & Permissions" で上記の権限を追加\n`;
      result += `3. "Reinstall to Workspace" をクリック\n`;
    } else {
      result += `\n✅ すべての必須権限が付与されています！`;
    }
    
    ui.alert('権限診断', result, ui.ButtonSet.OK);
    
  } catch (error) {
    console.error('権限診断エラー:', error);
    ui.alert('エラー', `権限診断中にエラーが発生しました:\n${error.toString()}`, ui.ButtonSet.OK);
  }
}

// ========= プライベートチャンネル詳細診断 =========
function diagnosePrivateChannels() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    console.log('チャンネルアクセス詳細診断開始...');
    
    let result = `チャンネルアクセス詳細診断結果：\n\n`;
    
    // Bot情報取得
    const authInfo = slackAPI('auth.test', {});
    result += `🤖 Bot名: @${authInfo.user || 'unknown'}\n`;
    result += `📍 ワークスペース: ${authInfo.team || 'unknown'}\n\n`;
    
    // デバッグ：API直接呼び出しで確認
    console.log('=== デバッグ: conversations.list直接呼び出し ===');
    const debugResponse = slackAPI('conversations.list', {
      types: 'private_channel',
      limit: 5
    });
    console.log('API Response:', JSON.stringify(debugResponse, null, 2));
    if (debugResponse.channels) {
      debugResponse.channels.forEach(ch => {
        console.log(`Channel: ${ch.name}, is_private: ${ch.is_private}, is_channel: ${ch.is_channel}, is_group: ${ch.is_group}`);
      });
    }
    
    // 1. 権限確認
    result += `【1. 権限チェック】\n`;
    
    // groups:read権限テスト
    try {
      const testList = slackAPI('conversations.list', { types: 'private_channel', limit: 1 });
      if (testList.ok) {
        result += `✅ groups:read権限: あり\n`;
      }
    } catch (e) {
      result += `❌ groups:read権限: なし（${e.toString()}）\n`;
    }
    
    // groups:history権限テスト（アクセス可能なプライベートチャンネルで）
    try {
      // まずBotがメンバーのプライベートチャンネルを探す
      const testList = slackAPI('conversations.list', { 
        types: 'private_channel', 
        limit: 100 
      });
      
      if (testList.ok && testList.channels) {
        let historyTestResult = false;
        
        // is_memberフラグが立っているチャンネルから試す
        const memberChannels = testList.channels.filter(ch => ch.is_member);
        
        for (const channel of memberChannels) {
          try {
            const history = slackAPI('conversations.history', { 
              channel: channel.id, 
              limit: 1 
            });
            if (history.ok) {
              historyTestResult = true;
              result += `✅ groups:history権限: あり（${channel.name}でテスト成功）\n`;
              break;
            }
          } catch (e) {
            // このチャンネルでは失敗、次を試す
          }
        }
        
        // is_memberでなくてもアクセスできるチャンネルを探す
        if (!historyTestResult) {
          for (const channel of testList.channels) {
            if (memberChannels.includes(channel)) continue;
            
            try {
              const history = slackAPI('conversations.history', { 
                channel: channel.id, 
                limit: 1 
              });
              if (history.ok) {
                historyTestResult = true;
                result += `✅ groups:history権限: あり（${channel.name}でApp統合アクセス確認）\n`;
                break;
              }
            } catch (e) {
              // このチャンネルでは失敗、次を試す
            }
          }
        }
        
        if (!historyTestResult) {
          result += `⚠️ groups:history権限: 権限はあるがアクセス可能なチャンネルがない\n`;
        }
      }
    } catch (e) {
      result += `⚠️ groups:history権限: 確認できず（${e.toString()}）\n`;
    }
    
    // 2. すべてのチャンネルを取得してから分類
    result += `\n【2. チャンネル一覧（タイプ別）】\n`;
    
    // すべてのチャンネルを取得（パブリック＋プライベート）
    const allChannels = [];
    let cursor = '';
    do {
      const params = {
        types: 'public_channel,private_channel',
        limit: 200,
        exclude_archived: true
      };
      if (cursor) params.cursor = cursor;
      
      const response = slackAPI('conversations.list', params);
      if (response.ok && response.channels) {
        allChannels.push(...response.channels);
        cursor = response.response_metadata?.next_cursor || '';
      } else {
        console.error('conversations.list エラー:', response.error);
        break;
      }
    } while (cursor);
    
    // チャンネルをタイプ別に分類
    const allPublicChannels = [];
    const allPrivateChannels = [];
    
    allChannels.forEach(channel => {
      // デバッグ情報出力
      console.log(`Channel: ${channel.name}, ID: ${channel.id}, is_private: ${channel.is_private}, is_channel: ${channel.is_channel}, is_group: ${channel.is_group}, is_member: ${channel.is_member}`);
      
      // Slack APIの新仕様: プライベートチャンネルもCで始まることがある
      // is_privateフラグが唯一の信頼できる判定基準
      const isPrivate = channel.is_private === true;
      
      // デバッグ: チャンネルタイプをログ
      if (isPrivate) {
        console.log(`🔒 プライベート: ${channel.name} (${channel.id})`);
      }
      
      // is_privateフラグのみで判定（新仕様対応）
      if (isPrivate) {
        allPrivateChannels.push(channel);
      } else {
        allPublicChannels.push(channel);
      }
    });
    
    result += `📢 パブリックチャンネル: ${allPublicChannels.length}個\n`;
    result += `🔒 プライベートチャンネル: ${allPrivateChannels.length}個\n`;
    result += `📦 合計: ${allChannels.length}個\n\n`;
    
    // デバッグ情報
    if (allPrivateChannels.length === 0 && allPublicChannels.length > 0) {
      result += `⚠️ 注意: プライベートチャンネルが検出されませんでした。\n`;
      result += `Botがプライベートチャンネルに招待されていない可能性があります。\n\n`;
    }
    
    // 3. 各チャンネルのアクセス状況を確認
    result += `【3. アクセス状況詳細】\n`;
    
    // デバッグ: チャンネルIDのプレフィックスを確認
    console.log('=== チャンネルIDプレフィックス分析 ===');
    const cPrefixChannels = allChannels.filter(ch => ch.id && ch.id.startsWith('C')).length;
    const gPrefixChannels = allChannels.filter(ch => ch.id && ch.id.startsWith('G')).length;
    const otherPrefixChannels = allChannels.filter(ch => ch.id && !ch.id.startsWith('C') && !ch.id.startsWith('G')).length;
    console.log(`Cで始まるチャンネル: ${cPrefixChannels}個`);
    console.log(`Gで始まるチャンネル: ${gPrefixChannels}個`);
    console.log(`その他: ${otherPrefixChannels}個`);
    
    // パブリックチャンネルのアクセス状況
    result += `\n📢 パブリックチャンネル（上位10個）:\n`;
    let publicAccessibleCount = 0;
    let publicMemberCount = 0;
    
    for (let i = 0; i < Math.min(allPublicChannels.length, 10); i++) {
      const channel = allPublicChannels[i];
      let status = '';
      
      if (channel.is_member) {
        status = '✅ メンバー';
        publicMemberCount++;
        publicAccessibleCount++;
      } else {
        status = '➕ 未参加（/invite で追加可能）';
      }
      
      result += `${i + 1}. #${channel.name} - ${status}\n`;
    }
    
    if (allPublicChannels.length > 10) {
      result += `... 他 ${allPublicChannels.length - 10} チャンネル\n`;
    }
    
    // プライベートチャンネルのアクセス状況
    result += `\n🔒 プライベートチャンネル（上位10個）:\n`;
    let privateAccessibleCount = 0;
    let privateMemberCount = 0;
    let privateAppAccessCount = 0;
    
    if (allPrivateChannels.length === 0) {
      result += `プライベートチャンネルが見つかりません。\n`;
      result += `※ Botがメンバーのプライベートチャンネルのみ表示されます\n`;
    } else {
      for (let i = 0; i < Math.min(allPrivateChannels.length, 10); i++) {
        const channel = allPrivateChannels[i];
        let status = '';
        
        // is_memberフラグ確認
        if (channel.is_member) {
          status = '✅ メンバー';
          privateMemberCount++;
          privateAccessibleCount++;
        } else {
          // アクセステスト
          try {
            const testHistory = slackAPI('conversations.history', {
              channel: channel.id,
              limit: 1
            });
            
            if (testHistory.ok) {
              status = '📱 App統合（アクセス可能）';
              privateAppAccessCount++;
              privateAccessibleCount++;
            } else {
              status = `❌ アクセス不可`;
            }
          } catch (e) {
            status = '❌ アクセス不可（招待が必要）';
          }
        }
        
        result += `${i + 1}. 🔒${channel.name} - ${status}\n`;
      }
      
      if (allPrivateChannels.length > 10) {
        result += `... 他 ${allPrivateChannels.length - 10} チャンネル\n`;
      }
    }
    
    // 4. サマリー
    result += `\n【4. サマリー】\n`;
    result += `\n📢 パブリックチャンネル:\n`;
    result += `  - 総数: ${allPublicChannels.length}個\n`;
    result += `  - メンバー: ${publicMemberCount}個\n`;
    result += `  - 未参加: ${allPublicChannels.length - publicMemberCount}個\n`;
    
    result += `\n🔒 プライベートチャンネル:\n`;
    result += `  - 総数: ${allPrivateChannels.length}個\n`;
    result += `  - メンバー: ${privateMemberCount}個\n`;
    result += `  - App統合: ${privateAppAccessCount}個\n`;
    result += `  - アクセス不可: ${allPrivateChannels.length - privateAccessibleCount}個\n`;
    
    // 5. 推奨事項
    result += `\n【5. 推奨アクション】\n`;
    
    // パブリックチャンネルの推奨事項
    if (publicMemberCount < allPublicChannels.length) {
      result += `\n📢 パブリックチャンネル:\n`;
      result += `${allPublicChannels.length - publicMemberCount}個のチャンネルに未参加です。\n`;
      result += `→ 必要に応じて「Botをチャンネルに追加」機能を使用してください\n`;
    }
    
    // プライベートチャンネルの推奨事項
    if (allPrivateChannels.length === 0) {
      result += `\n🔒 プライベートチャンネル:\n`;
      result += `プライベートチャンネルにBotを招待するには:\n`;
      result += `1. 対象のプライベートチャンネルで /invite @${authInfo.user || 'bot-name'} を実行\n`;
      result += `2. または、チャンネル設定 → Integrations → Add apps から追加\n`;
    } else if (privateAccessibleCount < allPrivateChannels.length) {
      result += `\n🔒 プライベートチャンネル:\n`;
      result += `${allPrivateChannels.length - privateAccessibleCount}個のチャンネルにアクセスできません。\n`;
      result += `\n各プライベートチャンネルで以下のコマンドを実行:\n`;
      result += `   /invite @${authInfo.user || 'Kushim Slack Governance'}\n`;
    } else if (allPrivateChannels.length > 0) {
      result += `\n✅ すべてのプライベートチャンネルにアクセス可能です！\n`;
    }
    
    result += `\n招待後は「🌐 Slack全チャンネル同期」でメッセージを取得してください。`;
    
    ui.alert('プライベートチャンネル詳細診断', result, ui.ButtonSet.OK);
    
  } catch (error) {
    console.error('プライベートチャンネル診断エラー:', error);
    ui.alert('エラー', `診断中にエラーが発生しました:\n${error.toString()}`, ui.ButtonSet.OK);
  }
}

// ========= プライベートチャンネル招待リスト生成 =========
function generateInviteList() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    console.log('プライベートチャンネル招待リスト生成開始...');
    
    // Bot情報取得
    const authInfo = slackAPI('auth.test', {});
    const botName = authInfo.user || 'Kushim Slack Governance';
    
    // すべてのチャンネルを取得してからプライベートを抽出
    const allChannels = [];
    let cursor = '';
    
    do {
      const params = {
        types: 'public_channel,private_channel',
        limit: 200,
        exclude_archived: true
      };
      if (cursor) params.cursor = cursor;
      
      const response = slackAPI('conversations.list', params);
      if (response.ok && response.channels) {
        allChannels.push(...response.channels);
        cursor = response.response_metadata?.next_cursor || '';
      } else {
        break;
      }
    } while (cursor);
    
    // プライベートチャンネルをフィルタリング（IDプレフィックスとis_privateフラグの両方を考慮）
    const allPrivateChannels = allChannels.filter(ch => {
      const isPrivateById = ch.id && ch.id.startsWith('G');
      const isPrivateByFlag = ch.is_private === true;
      return isPrivateById || isPrivateByFlag;
    });
    
    console.log(`全チャンネル数: ${allChannels.length}、プライベート: ${allPrivateChannels.length}`);
    
    // アクセスできないチャンネルを特定
    const inaccessibleChannels = [];
    
    for (const channel of allPrivateChannels) {
      let hasAccess = false;
      
      // is_memberチェック
      if (channel.is_member) {
        hasAccess = true;
      } else {
        // アクセステスト
        try {
          const testHistory = slackAPI('conversations.history', {
            channel: channel.id,
            limit: 1
          });
          if (testHistory.ok) {
            hasAccess = true;
          }
        } catch (e) {
          // アクセス不可
        }
      }
      
      if (!hasAccess) {
        inaccessibleChannels.push(channel);
      }
    }
    
    // 結果を表示
    if (inaccessibleChannels.length === 0) {
      ui.alert('完了', 'すべてのプライベートチャンネルにアクセス可能です！', ui.ButtonSet.OK);
      return;
    }
    
    // 招待コマンドリストを生成
    let result = `プライベートチャンネル招待リスト\n\n`;
    result += `以下の${inaccessibleChannels.length}個のチャンネルでBotの招待が必要です：\n\n`;
    
    result += `【招待コマンド一覧】\n`;
    result += `各チャンネルで以下のコマンドをコピー＆実行してください：\n\n`;
    
    inaccessibleChannels.forEach((channel, index) => {
      result += `${index + 1}. #${channel.name}\n`;
      result += `   /invite @${botName}\n\n`;
    });
    
    result += `\n【別の方法】\n`;
    result += `チャンネル設定（⚙️）→ Integrations → Add apps から\n`;
    result += `「${botName}」を検索して追加することもできます。\n\n`;
    
    result += `【注意事項】\n`;
    result += `- プライベートチャンネルのメンバーのみが招待を実行できます\n`;
    result += `- 招待後は「📊 チャンネルアクセス診断」で確認してください`;
    
    // スプレッドシートに書き出すオプション
    const response = ui.alert(
      'プライベートチャンネル招待リスト',
      `${inaccessibleChannels.length}個のチャンネルで招待が必要です。\n\nスプレッドシートに書き出しますか？`,
      ui.ButtonSet.YES_NO
    );
    
    if (response === ui.Button.YES) {
      const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
      const sheet = ss.insertSheet(`招待リスト_${new Date().toLocaleString('ja-JP')}`);
      
      // ヘッダー
      sheet.getRange(1, 1, 1, 3).setValues([['チャンネル名', 'チャンネルID', '招待コマンド']]);
      sheet.getRange(1, 1, 1, 3).setFontWeight('bold');
      
      // データ
      const data = inaccessibleChannels.map(ch => [
        ch.name,
        ch.id,
        `/invite @${botName}`
      ]);
      
      if (data.length > 0) {
        sheet.getRange(2, 1, data.length, 3).setValues(data);
      }
      
      // 列幅調整
      sheet.autoResizeColumns(1, 3);
      
      ui.alert('完了', `招待リストをスプレッドシートに書き出しました。\n\nシート名: ${sheet.getName()}`, ui.ButtonSet.OK);
    } else {
      ui.alert('招待リスト', result, ui.ButtonSet.OK);
    }
    
  } catch (error) {
    console.error('招待リスト生成エラー:', error);
    ui.alert('エラー', `招待リスト生成中にエラーが発生しました:\n${error.toString()}`, ui.ButtonSet.OK);
  }
}

// ========= ユーザー情報取得テスト =========
function testUserInfoFetch() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    console.log('ユーザー情報取得テスト開始...');
    
    // 1. users.listでユーザー一覧を取得
    const listResponse = slackAPI('users.list', { limit: 5 });
    if (!listResponse.ok) {
      ui.alert('エラー', `users.list APIでエラー: ${listResponse.error}`, ui.ButtonSet.OK);
      return;
    }
    
    const users = listResponse.members || [];
    console.log(`${users.length}人のユーザーを取得`);
    
    if (users.length === 0) {
      ui.alert('エラー', 'ユーザーが見つかりませんでした', ui.ButtonSet.OK);
      return;
    }
    
    // 2. 最初のユーザーで詳細情報を取得テスト
    const testUser = users[0];
    console.log(`テストユーザー: ${testUser.id}`);
    
    const infoResponse = slackAPI('users.info', { user: testUser.id });
    if (!infoResponse.ok) {
      ui.alert('エラー', `users.info APIでエラー: ${infoResponse.error}`, ui.ButtonSet.OK);
      return;
    }
    
    // 3. ユーザーID変換テスト
    const testText = `こんにちは <@${testUser.id}> さん`;
    const convertedText = convertSlackUserIdsToNames(testText);
    
    const result = `ユーザー情報取得テスト結果：
    
1. ユーザーID: ${testUser.id}
2. 表示名: ${testUser.name}
3. 実名: ${testUser.real_name || '未設定'}
4. メール: ${testUser.profile?.email || '取得不可'}

5. 変換テスト:
   変換前: ${testText}
   変換後: ${convertedText}

✅ users:read権限は正常に動作しています`;
    
    ui.alert('ユーザー情報取得テスト', result, ui.ButtonSet.OK);
    
  } catch (error) {
    ui.alert('エラー', `テスト中にエラーが発生しました:\n${error.toString()}`, ui.ButtonSet.OK);
  }
}

// コントロール番号生成
function generateControlNumber() {
  const date = new Date();
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  const random = Math.floor(Math.random() * 10000).toString().padStart(4, '0');
  return `CTRL-${year}${month}${day}-${random}`;
}

// 詳細シートフォーマット
function formatDetailedSheet(sheet) {
  if (!sheet) return;
  
  try {
    sheet.autoResizeColumns(1, sheet.getLastColumn());
    const headerRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());
    headerRange.setBackground('#4CAF50');
    headerRange.setFontColor('#FFFFFF');
    headerRange.setFontWeight('bold');
    sheet.setFrozenRows(1);
  } catch (e) {
    console.log('シートフォーマットエラー:', e);
  }
}

// ========= 詳細スプレッドシート作成（重要案件用） =========
function createDetailedWorkflowSpreadsheet(analysisResult, governanceResult, messages) {
  const spreadsheetName = `【重要】${analysisResult.categories?.[0] || '業務'}_${new Date().toISOString().split('T')[0]}`;
  const newSpreadsheet = SpreadsheetApp.create(spreadsheetName);
  
  // 全員編集可能に設定
  try {
    const file = DriveApp.getFileById(newSpreadsheet.getId());
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.EDIT);
  } catch (e) {
    console.log('共有設定エラー:', e);
  }
  
  // サマリーシート作成
  const summarySheet = newSpreadsheet.getActiveSheet();
  summarySheet.setName('サマリー');
  
  const summaryData = [
    ['項目', '内容'],
    ['作成日時', new Date().toLocaleString('ja-JP')],
    ['カテゴリ', analysisResult.categories?.join(', ') || ''],
    ['重要度', analysisResult.priority || 'MEDIUM'],
    ['リスクレベル', governanceResult.riskLevel || 'LOW'],
    ['承認要否', governanceResult.requiresApproval ? '要' : '不要'],
    ['承認レベル', governanceResult.approvalLevel || 'N/A'],
    ['開示要否', governanceResult.requiresDisclosure ? '要' : '不要'],
    ['開示種別', governanceResult.disclosureType || 'N/A'],
    ['コントロール番号', governanceResult.controlNumber || generateControlNumber()],
    ['概要', analysisResult.summary || '']
  ];
  
  summarySheet.getRange(1, 1, summaryData.length, 2).setValues(summaryData);
  formatDetailedSheet(summarySheet);
  
  // 議題・論点シート作成
  const topicsSheet = newSpreadsheet.insertSheet('議題・論点');
  const topicsData = [['No', '議題', '詳細', '優先度', '担当者', 'ステータス']];
  
  if (analysisResult.topics && Array.isArray(analysisResult.topics)) {
    analysisResult.topics.forEach((topic, index) => {
      topicsData.push([
        index + 1,
        typeof topic === 'object' ? (topic.title || '') : topic,
        typeof topic === 'object' ? (topic.description || '') : '',
        typeof topic === 'object' ? (topic.priority || index + 1) : index + 1,
        '',
        '未対応'
      ]);
    });
  }
  
  topicsSheet.getRange(1, 1, topicsData.length, 6).setValues(topicsData);
  formatDetailedSheet(topicsSheet);
  
  // アクションアイテムシート作成
  const actionSheet = newSpreadsheet.insertSheet('アクションアイテム');
  const actionData = [['No', 'タスク', '担当者', '期限', 'ステータス', '備考']];
  
  if (analysisResult.actionItems && Array.isArray(analysisResult.actionItems)) {
    analysisResult.actionItems.forEach((item, index) => {
      const itemObj = typeof item === 'object' ? item : { task: item };
      actionData.push([
        index + 1,
        itemObj.task || item,
        itemObj.owner || '',
        itemObj.deadline || '',
        '未着手',
        ''
      ]);
    });
  }
  
  actionSheet.getRange(1, 1, actionData.length, 6).setValues(actionData);
  formatDetailedSheet(actionSheet);
  
  return newSpreadsheet;
}

// ========= 包括的業務フロー文書生成 =========
function generateComprehensiveWorkflowDocuments(spreadsheet, analysisResult, governanceResult) {
  // 業務フローシート作成
  const workflowSheet = spreadsheet.insertSheet('業務フロー');
  
  const flowData = [
    ['業務フローチャート', '', '', '', '', '', '', ''],
    ['', '', '', '', '', '', '', ''],
    ['ステップ', 'プロセス', '担当者(R)', '承認者(A)', '相談先(C)', '情報共有(I)', '期限', 'ステータス']
  ];
  
  // 標準業務フロー生成
  const standardFlow = generateStandardBusinessFlow(analysisResult, governanceResult);
  standardFlow.forEach((step, index) => {
    flowData.push([
      index + 1,
      step.process,
      step.responsible,
      step.accountable,
      step.consulted,
      step.informed,
      step.deadline,
      '未着手'
    ]);
  });
  
  workflowSheet.getRange(1, 1, flowData.length, 8).setValues(flowData);
  formatDetailedSheet(workflowSheet);
  
  // 業務記述書シート作成
  const descriptionSheet = spreadsheet.insertSheet('業務記述書');
  generateEnhancedBusinessDescription(descriptionSheet, analysisResult, governanceResult);
  formatDetailedSheet(descriptionSheet);
}

// ========= 標準業務フロー生成 =========
function generateStandardBusinessFlow(analysisResult, governanceResult) {
  const workflow = [];
  
  // 基本フロー
  workflow.push({
    process: '課題・要件の確認',
    responsible: '担当者',
    accountable: '部門長',
    consulted: '関係部署',
    informed: 'チームメンバー',
    deadline: '1営業日'
  });
  
  workflow.push({
    process: '現状分析・課題整理',
    responsible: '担当者',
    accountable: '部門長',
    consulted: '分析チーム',
    informed: '関係者',
    deadline: '3営業日'
  });
  
  // 専門家相談が必要な場合
  if (governanceResult.requiresExpertConsultation) {
    workflow.push({
      process: '専門家への相談・助言取得',
      responsible: '法務・総務',
      accountable: '部門長',
      consulted: governanceResult.requiredExperts?.join('・') || '専門家',
      informed: '経営陣',
      deadline: '3営業日'
    });
  }
  
  // 承認が必要な場合
  if (governanceResult.requiresApproval) {
    workflow.push({
      process: '社内承認プロセス',
      responsible: '部門長',
      accountable: governanceResult.approvalLevel || '取締役',
      consulted: '法務・財務',
      informed: '監査役',
      deadline: '5営業日'
    });
  }
  
  // 開示が必要な場合
  if (governanceResult.requiresDisclosure) {
    workflow.push({
      process: '開示資料作成・確認',
      responsible: 'IR担当',
      accountable: 'CFO',
      consulted: '法務・会計士',
      informed: '取締役会',
      deadline: '適時'
    });
  }
  
  // 実行フェーズ
  workflow.push({
    process: '実行・実施',
    responsible: '担当者',
    accountable: '部門長',
    consulted: '関係部署',
    informed: '関係者全員',
    deadline: '計画通り'
  });
  
  workflow.push({
    process: 'モニタリング・報告',
    responsible: '担当者',
    accountable: '部門長',
    consulted: '品質管理',
    informed: '経営陣',
    deadline: '継続'
  });
  
  return workflow;
}

// ========= 強化版業務記述書生成 =========
function generateEnhancedBusinessDescription(sheet, analysisResult, governanceResult) {
  const description = [
    ['業務記述書'],
    [''],
    ['1. 目的'],
    [analysisResult.summary || '本業務の目的を記載'],
    [''],
    ['2. 適用範囲'],
    ['本記述書は以下の業務に適用される：'],
    ['・' + (analysisResult.categories?.join('\n・') || '適用範囲を記載')],
    [''],
    ['3. 責任と権限'],
    ['責任者：' + (governanceResult.approvalLevel || '部門長')],
    ['承認者：' + (governanceResult.approvalLevel || '取締役')],
    [''],
    ['4. 業務手順'],
    ['詳細は業務フローシート参照'],
    [''],
    ['5. リスクと統制'],
    ['リスクレベル：' + (governanceResult.riskLevel || 'MEDIUM')],
    ['統制番号：' + (governanceResult.controlNumber || generateControlNumber())],
    [''],
    ['6. 監査ポイント'],
    [(governanceResult.auditPoints?.join('\n') || '監査ポイントなし')],
    [''],
    ['7. 改訂履歴'],
    ['作成日：' + new Date().toLocaleDateString('ja-JP')],
    ['作成者：システム自動生成']
  ];
  
  description.forEach((row, index) => {
    sheet.getRange(index + 1, 1).setValue(row[0]);
    if (row[0].match(/^\d\./)) {
      sheet.getRange(index + 1, 1).setFontWeight('bold');
      sheet.getRange(index + 1, 1).setFontSize(12);
    }
  });
}

// ========= 詳細議事録案生成 =========
function generateDetailedMeetingMinutes(spreadsheet, analysisResult, governanceResult) {
  const minutesSheet = spreadsheet.insertSheet('議事録案');
  
  let meetingType = '経営会議';
  if (governanceResult.meetingType) {
    meetingType = governanceResult.meetingType;
  } else if (governanceResult.requiresApproval && analysisResult.priority === 'HIGH') {
    meetingType = '取締役会';
  }
  
  const minutesData = [
    [`${meetingType}議事録（案）`],
    [''],
    ['日時：' + new Date().toLocaleDateString('ja-JP')],
    ['場所：'],
    ['出席者：'],
    [''],
    ['【議題】']
  ];
  
  // 議題追加
  if (analysisResult.topics && Array.isArray(analysisResult.topics)) {
    analysisResult.topics.forEach((topic, index) => {
      const topicTitle = typeof topic === 'object' ? topic.title : topic;
      minutesData.push([`${index + 1}. ${topicTitle}`]);
    });
  }
  
  minutesData.push(
    [''],
    ['【審議内容】'],
    [''],
    ['【議案説明】'],
    [analysisResult.summary || ''],
    [''],
    ['【決議事項】'],
    ['以下の事項について、全会一致で承認可決された。']
  );
  
  // アクションアイテムを決議事項として追加
  if (analysisResult.actionItems && Array.isArray(analysisResult.actionItems)) {
    analysisResult.actionItems.forEach(item => {
      const itemText = typeof item === 'object' ? item.task : item;
      minutesData.push([`・${itemText}`]);
    });
  }
  
  minutesData.push(
    [''],
    ['以上'],
    [''],
    ['議事録作成者：'],
    ['確認者：']
  );
  
  minutesData.forEach((row, index) => {
    minutesSheet.getRange(index + 1, 1).setValue(row[0]);
    if (index === 0) {
      minutesSheet.getRange(1, 1).setFontSize(16);
      minutesSheet.getRange(1, 1).setFontWeight('bold');
      minutesSheet.getRange(1, 1).setHorizontalAlignment('center');
    }
    if (row[0].startsWith('【')) {
      minutesSheet.getRange(index + 1, 1).setFontWeight('bold');
    }
  });
  
  formatDetailedSheet(minutesSheet);
}

// ========= 強化版通知送信 =========
function sendEnhancedNotifications(workflowResult) {
  const result = {
    email: false,
    slack: false,
    errors: []
  };
  
  try {
    if (REPORT_EMAIL) {
      sendEnhancedHtmlEmail(workflowResult);
      result.email = true;
    }
  } catch (error) {
    result.errors.push('メール送信エラー: ' + error.toString());
  }
  
  try {
    if (SLACK_BOT_TOKEN) {
      sendEnhancedSlackNotification(workflowResult);
      result.slack = true;
    }
  } catch (error) {
    result.errors.push('Slack通知エラー: ' + error.toString());
  }
  
  return result;
}

// ========= 強化版HTMLメール送信 =========
function sendEnhancedHtmlEmail(workflowResult) {
  const subject = `【${workflowResult.analysis?.priority || 'INFO'}】Slackガバナンスレポート - ${new Date().toLocaleDateString('ja-JP')}`;
  
  // 優先度に応じた色設定（オーソドックスな色調）
  const priorityColors = {
    'HIGH': '#dc3545',    // 標準的な赤
    'MEDIUM': '#ffc107',  // 標準的な黄色
    'LOW': '#28a745',     // 標準的な緑
    'INFO': '#17a2b8'     // 標準的な青
  };
  const priorityColor = priorityColors[workflowResult.analysis?.priority] || '#667EEA';
  const spreadsheetUrl = workflowResult.spreadsheetUrl || `https://docs.google.com/spreadsheets/d/${SPREADSHEET_ID}/edit`;
  const effectiveRisk = workflowResult.governance?.riskLevel || 'LOW';
  
  const htmlBody = `
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <style>
    * { margin: 0; padding: 0; box-sizing: border-box; }
    body { 
      font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
      line-height: 1.6;
      color: #2C3E50;
      background: #F5F7FA;
    }
    .container { 
      max-width: 900px;
      margin: 0 auto;
      background: white;
      box-shadow: 0 0 30px rgba(0,0,0,0.1);
    }
    .header { 
      background: #2c3e50;
      color: white;
      padding: 40px 30px;
      position: relative;
      overflow: hidden;
    }
    .header::before {
      content: '';
      position: absolute;
      top: -50%;
      right: -10%;
      width: 300px;
      height: 300px;
      background: rgba(255,255,255,0.1);
      border-radius: 50%;
    }
    .header h1 { 
      font-size: 32px;
      margin-bottom: 10px;
      position: relative;
      z-index: 1;
    }
    .header p { 
      opacity: 0.95;
      font-size: 16px;
      position: relative;
      z-index: 1;
    }
    
    .metrics { 
      display: flex;
      justify-content: space-around;
      padding: 30px;
      background: #FAFBFC;
      border-bottom: 1px solid #E1E8ED;
    }
    .metric-card { 
      text-align: center;
      padding: 20px;
      background: white;
      border-radius: 12px;
      box-shadow: 0 4px 6px rgba(0,0,0,0.07);
      min-width: 150px;
      transition: transform 0.2s;
    }
    .metric-card:hover { transform: translateY(-3px); }
    .metric-card h3 { 
      font-size: 36px;
      color: ${priorityColor};
      margin-bottom: 8px;
      font-weight: bold;
    }
    .metric-card p { 
      color: #64748B;
      font-size: 14px;
      font-weight: 500;
    }
    
    .content { padding: 30px; }
    
    .section { 
      background: white;
      border: 1px solid #E1E8ED;
      border-radius: 12px;
      padding: 25px;
      margin-bottom: 25px;
      box-shadow: 0 2px 4px rgba(0,0,0,0.05);
    }
    .section h2 { 
      color: #1E293B;
      font-size: 22px;
      margin-bottom: 15px;
      padding-bottom: 10px;
      border-bottom: 2px solid ${priorityColor};
      display: inline-block;
    }
    .section h3 { 
      color: #334155;
      font-size: 18px;
      margin: 15px 0 10px;
    }
    .section p { 
      color: #475569;
      line-height: 1.8;
      margin-bottom: 10px;
    }
    
    .topic-list {
      list-style: none;
      padding: 0;
      margin: 15px 0;
    }
    .topic-item {
      background: #F8FAFC;
      border-left: 4px solid ${priorityColor};
      padding: 15px;
      margin: 10px 0;
      border-radius: 8px;
    }
    .topic-item strong {
      color: #1E293B;
      display: block;
      margin-bottom: 5px;
    }
    
    .action-items {
      background: #FEF3C7;
      border: 1px solid #FCD34D;
      border-radius: 8px;
      padding: 20px;
      margin: 20px 0;
    }
    .action-items h3 {
      color: #92400E;
      margin-bottom: 10px;
    }
    .action-item {
      background: white;
      padding: 12px;
      margin: 8px 0;
      border-radius: 6px;
      border-left: 3px solid #F59E0B;
    }
    
    .governance-alert {
      background: #FEE2E2;
      border: 1px solid #FCA5A5;
      border-radius: 8px;
      padding: 20px;
      margin: 20px 0;
    }
    .governance-alert h3 {
      color: #991B1B;
      margin-bottom: 10px;
    }
    
    .btn-container { 
      text-align: center;
      padding: 30px;
      background: #F8FAFC;
    }
    .btn { 
      display: inline-block;
      padding: 16px 40px;
      background: #2c3e50;
      color: white;
      text-decoration: none;
      border-radius: 4px;
      font-weight: 600;
      font-size: 16px;
      box-shadow: 0 2px 4px rgba(0,0,0,0.1);
      transition: all 0.3s;
    }
    .btn:hover { 
      transform: translateY(-2px);
      box-shadow: 0 12px 20px rgba(0,0,0,0.2);
    }
    
    .footer {
      background: #1E293B;
      color: #94A3B8;
      padding: 20px;
      text-align: center;
      font-size: 14px;
    }
    .footer a {
      color: #60A5FA;
      text-decoration: none;
    }
    
    .priority-badge {
      display: inline-block;
      padding: 6px 12px;
      border-radius: 20px;
      font-size: 12px;
      font-weight: bold;
      text-transform: uppercase;
      background: ${priorityColor};
      color: white;
      margin-left: 10px;
    }
  </style>
</head>
<body>
  <div class="container">
    <div class="header">
      <h1>📊 Slackガバナンスレポート
        <span class="priority-badge">${workflowResult.analysis?.priority || 'INFO'}</span>
      </h1>
      <p>生成日時: ${new Date().toLocaleString('ja-JP')}</p>
      <p style="margin-top:6px;">📑 スプレッドシート: <a href="${spreadsheetUrl}">${spreadsheetUrl}</a></p>
    </div>
    
    <div class="metrics">
      <div class="metric-card">
        <h3>${workflowResult.messages?.length || 0}</h3>
        <p>分析メッセージ数</p>
      </div>
      <div class="metric-card">
        <h3>${workflowResult.analysis?.topics?.length || 0}</h3>
        <p>抽出議題数</p>
      </div>
      <div class="metric-card">
        <h3>${workflowResult.analysis?.actionItems?.length || 0}</h3>
        <p>アクション項目</p>
      </div>
      <div class="metric-card">
        <h3>${effectiveRisk}</h3>
        <p>リスクレベル</p>
      </div>
    </div>
    
    <div class="content">
      <div class="section">
        <h2>📝 分析サマリー</h2>
        <p>${(workflowResult.analysis?.summary || 'メッセージの分析結果').replace(/\n/g, '<br>')}</p>
      </div>
      
      ${workflowResult.analysis?.topics && workflowResult.analysis.topics.length > 0 ? (() => {
        // 議題を分類
        const classified = classifyTopics(workflowResult.analysis, workflowResult.governance);
        
        let topicsHtml = '';
        
        // 重要議題がある場合
        if (classified.importantAgendas && classified.importantAgendas.length > 0) {
          topicsHtml += `
          <div class="section">
            <h2>💡 重要議題（決議事項・開示事項等）</h2>
            <ul class="topic-list">
              ${classified.importantAgendas.map((topic, i) => `
                <li class="topic-item" style="border-left-color: #dc3545;">
                  <strong>${i + 1}. ${typeof topic === 'object' ? (topic.title || topic.topic || topic) : topic}</strong>
                  ${topic.category ? `<span style="color: #dc3545; font-size: 12px;"> [ ${topic.category} ]</span>` : ''}
                  ${typeof topic === 'object' && topic.description ? `<p>${topic.description}</p>` : ''}
                </li>
              `).join('')}
            </ul>
          </div>`;
        }
        
        // 一般トピックがある場合
        if (classified.generalTopics && classified.generalTopics.length > 0) {
          topicsHtml += `
          <div class="section">
            <h2>📄 その他のトピック</h2>
            <ul class="topic-list">
              ${classified.generalTopics.slice(0, 5).map((topic, i) => `
                <li class="topic-item">
                  <strong>${i + 1}. ${typeof topic === 'object' ? (topic.title || topic.topic || topic) : topic}</strong>
                  ${typeof topic === 'object' && topic.description ? `<p>${topic.description}</p>` : ''}
                </li>
              `).join('')}
            </ul>
          </div>`;
        }
        
        return topicsHtml;
      })() : ''}
      
      ${workflowResult.analysis?.actionItems && workflowResult.analysis.actionItems.length > 0 ? `
      <div class="action-items">
        <h3>⚡ アクションアイテム</h3>
        ${workflowResult.analysis.actionItems.slice(0, 5).map(item => `
          <div class="action-item">
            ${typeof item === 'object' ? (item.task || item.action || item) : item}
            ${typeof item === 'object' && item.owner ? ` (担当: ${item.owner})` : ''}
            ${typeof item === 'object' && item.deadline ? ` - 期限: ${item.deadline}` : ''}
          </div>
        `).join('')}
      </div>
      ` : ''}
      
      ${workflowResult.governance?.requiresAction || workflowResult.governance?.riskLevel === 'HIGH' ? `
      <div class="governance-alert">
        <h3>⚠️ ガバナンス・コンプライアンス要対応事項</h3>
        <p>リスクレベル: <strong>${workflowResult.governance.riskLevel}</strong></p>
        ${workflowResult.governance.requiresApproval ? '<p>✓ 承認が必要です</p>' : ''}
        ${workflowResult.governance.requiresDisclosure ? '<p>✓ 開示が必要です</p>' : ''}
        ${workflowResult.governance.requiresExpertConsultation ? '<p>✓ 専門家への相談が必要です</p>' : ''}
        ${workflowResult.governance.controlNumber ? `<p>管理番号: ${workflowResult.governance.controlNumber}</p>` : ''}
      </div>
      ` : ''}
    </div>
    
    <div class="btn-container">
      <a href="${spreadsheetUrl}" class="btn">📑 詳細スプレッドシートを開く</a>
    </div>
    
    <div class="footer">
      <p>このメールは自動生成されました | Slackガバナンスシステム v2.0</p>
      <p style="margin-top: 10px;">© 2024 Automated Workflow System</p>
    </div>
  </div>
</body>
</html>`;
  
  // プレーンテキスト版も作成
  const plainBody = `
Slackガバナンスレポート [${workflowResult.analysis?.priority || 'INFO'}]
${'='.repeat(50)}

生成日時: ${new Date().toLocaleString('ja-JP')}
スプレッドシート: ${spreadsheetUrl}

【分析結果】
- メッセージ数: ${workflowResult.messages?.length || 0}件
- 議題数: ${workflowResult.analysis?.topics?.length || 0}件
- アクション項目: ${workflowResult.analysis?.actionItems?.length || 0}件
- リスクレベル: ${effectiveRisk}

【サマリー】
${workflowResult.analysis?.summary || 'メッセージの分析結果'}

詳細はこちら: ${spreadsheetUrl}
`;
  
  MailApp.sendEmail({
    to: REPORT_EMAIL,
    subject: subject,
    body: plainBody,
    htmlBody: htmlBody
  });
}

// ========= 強化版Slack通知 =========
function sendEnhancedSlackNotification(workflowResult) {
  // 優先度に応じた絵文字
  const priorityEmoji = {
    'HIGH': '🔴',
    'MEDIUM': '🟡',
    'LOW': '🔵',
    'INFO': 'ℹ️'
  };
  const emoji = priorityEmoji[workflowResult.analysis?.priority] || 'ℹ️';
  
  const blocks = [
    {
      type: "header",
      text: {
        type: "plain_text",
        text: `${emoji} Slack分析完了通知 [${workflowResult.analysis?.priority || 'INFO'}]`,
        emoji: true
      }
    },
    {
      type: "section",
      fields: [
        {
          type: "mrkdwn",
          text: `*📈 分析メッセージ数:*\n${workflowResult.messages?.length || 0}件`
        },
        {
          type: "mrkdwn",
          text: `*🎯 抽出議題数:*\n${workflowResult.analysis?.topics?.length || 0}件`
        },
        {
          type: "mrkdwn",
          text: `*⚡ アクション項目:*\n${workflowResult.analysis?.actionItems?.length || 0}件`
        },
        {
          type: "mrkdwn",
          text: `*⚠️ リスクレベル:*\n${workflowResult.governance?.riskLevel || 'N/A'}`
        }
      ]
    },
    {
      type: "divider"
    },
    {
      type: "section",
      text: {
        type: "mrkdwn",
        text: `*📝 分析サマリー:*\n${workflowResult.analysis?.summary || '分析結果サマリー'}`
      }
    }
  ];
  
  // 主要議題を追加
  if (workflowResult.analysis?.topics && workflowResult.analysis.topics.length > 0) {
    const topicsList = workflowResult.analysis.topics.slice(0, 3).map((topic, i) => {
      const topicText = typeof topic === 'object' ? (topic.title || topic.topic || topic) : topic;
      return `${i + 1}. ${topicText}`;
    }).join('\n');
    
    blocks.push({
      type: "section",
      text: {
        type: "mrkdwn",
        text: `*🎯 主要議題:*\n${topicsList}`
      }
    });
  }
  
  // アクションアイテムを追加
  if (workflowResult.analysis?.actionItems && workflowResult.analysis.actionItems.length > 0) {
    const actionsList = workflowResult.analysis.actionItems.slice(0, 3).map((item, i) => {
      const itemText = typeof item === 'object' ? (item.task || item.action || item) : item;
      return `• ${itemText}`;
    }).join('\n');
    
    blocks.push({
      type: "section",
      text: {
        type: "mrkdwn",
        text: `*⚡ 緊急アクション:*\n${actionsList}`
      }
    });
  }
  
  // ガバナンス警告
  if (workflowResult.governance?.requiresAction || workflowResult.governance?.riskLevel === 'HIGH') {
    blocks.push({
      type: "section",
      text: {
        type: "mrkdwn",
        text: `*⚠️ ガバナンス要対応:*\n${[
          workflowResult.governance.requiresApproval ? '• 承認が必要' : '',
          workflowResult.governance.requiresDisclosure ? '• 開示が必要' : '',
          workflowResult.governance.requiresExpertConsultation ? '• 専門家相談が必要' : ''
        ].filter(x => x).join('\n') || 'なし'}`
      }
    });
  }
  
  // スプレッドシートリンク
  if (workflowResult.spreadsheetUrl) {
    blocks.push({
      type: "divider"
    });
    blocks.push({
      type: "section",
      text: {
        type: "mrkdwn",
        text: `📋 *詳細資料:* <${workflowResult.spreadsheetUrl}|スプレッドシートを開く>`
      },
      accessory: {
        type: "button",
        text: {
          type: "plain_text",
          text: "詳細を見る",
          emoji: true
        },
        url: workflowResult.spreadsheetUrl,
        action_id: "view_spreadsheet"
      }
    });
  }
  
  // フッター
  blocks.push({
    type: "context",
    elements: [
      {
        type: "mrkdwn",
        text: `生成日時: ${new Date().toLocaleString('ja-JP')} | 管理番号: ${workflowResult.governance?.controlNumber || 'N/A'}`
      }
    ]
  });
  
  // 通知先チャンネルを設定から取得
  let notifyChannel = '#general';
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const configSheet = ss.getSheetByName(SHEETS.CONFIG);
    if (configSheet) {
      const configData = configSheet.getDataRange().getValues();
      for (let i = 1; i < configData.length; i++) {
        if (configData[i][0] === 'notifySlackChannel' && configData[i][1]) {
          notifyChannel = configData[i][1];
          break;
        }
      }
    }
  } catch (e) {
    console.log('設定チャンネル取得エラー:', e);
  }
  
  try {
    const response = slackAPI('chat.postMessage', {
      channel: notifyChannel,
      blocks: blocks,
      text: `Slack分析完了 [${workflowResult.analysis?.priority || 'INFO'}]`
    });
    console.log(`Slack通知送信成功: ${notifyChannel}`);
    return response;
  } catch (e) {
    console.log('Slack通知エラー:', e);
    throw e;
  }
}

// ========= 包括的ワークフローログ記録 =========
function recordComprehensiveWorkflowLogs(logs, result) {
  try {
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    let logSheet = spreadsheet.getSheetByName('Logs');
    
    if (!logSheet) {
      logSheet = spreadsheet.insertSheet('Logs');
      logSheet.getRange(1, 1, 1, 6).setValues([[
        '実行日時', '処理', 'ステータス', '詳細', 'エラー', '実行時間'
      ]]);
      formatDetailedSheet(logSheet);
    }
    
    const logData = logs.map(log => [
      log.timestamp || new Date().toLocaleString('ja-JP'),
      log.process || '',
      log.status || '',
      log.detail || '',
      log.error || '',
      log.executionTime || ''
    ]);
    
    const lastRow = logSheet.getLastRow();
    logSheet.getRange(lastRow + 1, 1, logData.length, 6).setValues(logData);
  } catch (e) {
    console.log('ログ記録エラー:', e);
  }
}

// ========= メッセージをシートに保存 =========
function saveMessagesToSheets(channelId, messages) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    
    // 通知チャンネルは除外
    const configSheet = ss.getSheetByName(SHEETS.CONFIG);
    if (configSheet) {
      const config = getConfigData(configSheet);
      if (config.notifySlackChannel && config.notifySlackChannel === channelId) {
        console.log(`通知チャンネル ${channelId} はログ記録から除外されます`);
        return;
      }
    }
    
    // ユーザー情報を事前に一括読み込み（パフォーマンス向上）
    if (Date.now() > userCacheExpiry) {
      loadAllUsers();
    }
    
    // SyncStateシートを取得・作成
    const syncSheet = getOrCreateSheet(ss, SHEETS.SYNC_STATE, [
      'channel_id', 'last_sync_ts', 'last_sync_datetime', 'message_count', 'status'
    ]);
    
    // Messagesシートを取得・作成
    const messagesSheet = getOrCreateSheet(ss, SHEETS.MESSAGES, [
      'id', 'channel_id', 'message_ts', 'thread_ts', 'text_raw', 'user_name',
      'summary_json', 'classification_json', 'match_flag', 'human_judgement',
      'permalink', 'checklist_proposed', 'agenda_selected', 'draft_doc_url',
      'timestamp', 'reactions', 'files'
    ]);
    
    // slack_logシートを取得・作成
    const slackLogSheet = getOrCreateSheet(ss, SHEETS.SLACK_LOG, [
      'channel_id', 'channel_name', 'ts', 'thread_ts', 
      'user_name', 'message', 'date', 'reactions', 'files'
    ]);
    
    // チャンネル名を取得（エラーの場合はチャンネルIDを使用）
    let channelName = channelId;
    try {
      // チャンネル情報取得を試行（プライベートチャンネルやBotが参加していないチャンネルでは失敗する可能性あり）
      const channelInfo = slackAPI('conversations.info', { channel: channelId });
      if (channelInfo && channelInfo.channel) {
        channelName = channelInfo.channel.name;
      }
    } catch (e) {
      // エラーは警告レベルで記録（処理は継続）
      console.warn(`チャンネル名取得スキップ: ${channelId} (${e.toString()})`);
      // チャンネルリストから名前を探す
      try {
        const channels = slackAPI('conversations.list', { types: 'public_channel,private_channel' });
        const found = channels.channels?.find(ch => ch.id === channelId);
        if (found) {
          channelName = found.name;
        }
      } catch (listError) {
        // リスト取得も失敗した場合はチャンネルIDをそのまま使用
      }
    }
    
    // 既存のメッセージIDを取得して重複チェック
    const existingIds = new Set();
    const lastRow = messagesSheet.getLastRow();
    if (lastRow > 1) {
      const existingData = messagesSheet.getRange(2, 1, lastRow - 1, 1).getValues();
      existingData.forEach(row => {
        if (row[0]) existingIds.add(row[0].toString());
      });
    }
    
    // バッチ用データ準備（重複を除外）
    const messageBatch = [];
    const slackLogBatch = [];
    let duplicateCount = 0;
    
    messages.forEach(message => {
      const messageId = `${channelId}_${message.ts}`;
      if (existingIds.has(messageId)) {
        duplicateCount++;
        console.log(`重複メッセージをスキップ: ${messageId}`);
      } else {
        // メッセージデータを準備
        messageBatch.push(prepareMessageRow(channelId, message));
        slackLogBatch.push(prepareSlackLogRow(channelId, channelName, message));
      }
    });
    
    // バッチ保存
    if (messageBatch.length > 0) {
      console.log(`${channelId}: ${messageBatch.length}件の新規メッセージを保存中... (${duplicateCount}件の重複をスキップ)`);
      saveMessagesBatch(messagesSheet, messageBatch);
      saveSlackLogBatch(slackLogSheet, slackLogBatch);
    } else {
      console.log(`${channelId}: 新規メッセージなし (${duplicateCount}件の重複メッセージをスキップ)`);
    }
    
    // 最終同期時刻を更新
    if (messages.length > 0) {
      const latestTs = messages[0].ts; // Slack APIは新しい順で返す
      updateLastSyncTime(syncSheet, channelId, latestTs);
      console.log(`${channelId}: 最終同期時刻を更新: ${latestTs}`);
    }
    
    console.log(`✅ ${channelId}: メッセージ保存完了`);
    
  } catch (error) {
    console.error(`メッセージ保存エラー (${channelId}):`, error);
  }
}

// ========= チャンネルメッセージ取得（リトライ付き） =========
function fetchChannelMessagesWithRetry(channelId, maxRetries = 3) {
  let retries = 0;
  let lastError;
  
  // 最終同期時刻を取得
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const syncSheet = getOrCreateSheet(ss, SHEETS.SYNC_STATE, [
    'channel_id', 'last_sync_ts', 'last_sync_datetime', 'message_count', 'status'
  ]);
  const lastSyncTs = getLastSyncTime(syncSheet, channelId);
  
  console.log(`チャンネル ${channelId} の最終同期時刻: ${lastSyncTs}`);
  
  // チャンネル情報を事前に確認
  const channelInfo = getChannelInfo(channelId);
  const isPrivate = channelInfo?.is_private || false;
  const isMember = channelInfo?.is_member || false;
  
  if (channelInfo) {
    console.log(`チャンネル情報: ${channelInfo.name} - プライベート: ${isPrivate}, メンバー: ${isMember}`);
    
    if (!isMember) {
      console.warn(`警告: Bot はチャンネル ${channelId} のメンバーではありません`);
      // プライベートチャンネルで Bot がメンバーでない場合はアクセス不可
      if (isPrivate) {
        console.error(`プライベートチャンネル ${channelId} へのアクセス権限がありません。Slackで /invite @bot_name を実行してください`);
        return [];
      }
    }
  }
  
  while (retries < maxRetries) {
    try {
      // 最終同期時刻より新しいメッセージのみを取得
      const params = {
        channel: channelId,
        limit: MAX_MESSAGES_PER_CHANNEL || 100
      };
      
      // 最終同期時刻がある場合は、それより新しいメッセージのみ取得
      if (lastSyncTs && lastSyncTs !== '0') {
        params.oldest = lastSyncTs;
        params.inclusive = false; // 最終同期時刻のメッセージ自体は含めない
        console.log(`${channelId}: タイムスタンプ ${lastSyncTs} より新しいメッセージのみ取得`);
      }
      
      const response = slackAPI('conversations.history', params);
      const messages = response.messages || [];
      
      console.log(`${channelId}: ${messages.length}件の新規メッセージを取得`);
      
      // メッセージをシートに保存
      if (messages.length > 0) {
        saveMessagesToSheets(channelId, messages);
      } else {
        console.log(`${channelId}: 新規メッセージなし`);
        // 新規メッセージがなくても最終同期時刻を現在時刻に更新
        const nowTs = Math.floor(Date.now() / 1000).toString();
        updateLastSyncTime(syncSheet, channelId, nowTs);
      }
      
      return messages;
      
    } catch (error) {
      lastError = error;
      const errorStr = error.toString();
      
      // エラーの種類によって処理を分岐
      if (errorStr.includes('channel_not_found')) {
        console.error(`チャンネル ${channelId} が見つかりません`);
        return [];
      }
      
      if (errorStr.includes('not_in_channel') || errorStr.includes('cant_read_channel')) {
        console.error(`チャンネル ${channelId} へのアクセス権限がありません。Bot をチャンネルに招待してください。`);
        return [];
      }
      
      if (errorStr.includes('invalid_auth')) {
        console.error('認証エラー: Slack Bot Tokenを確認してください');
        throw error;
      }
      
      if (errorStr.includes('missing_scope')) {
        const requiredScope = isPrivate ? 'groups:history' : 'channels:history';
        console.error(`権限不足: ${requiredScope} スコープが必要です。Slack Appの設定で追加してください`);
        return [];
      }
      
      retries++;
      if (retries < maxRetries) {
        console.log(`リトライ ${retries}/${maxRetries}...`);
        Utilities.sleep(1000 * retries);
      }
    }
  }
  
  console.error(`チャンネル${channelId}のメッセージ取得に失敗: ${lastError}`);
  return []; // エラー時は空配列を返す
}

// ========= 全参加チャンネルのメッセージ取得 =========
function fetchAllJoinedChannelsMessages() {
  const channels = getJoinedChannels();
  let allMessages = [];
  
  channels.forEach(channel => {
    try {
      const messages = fetchChannelMessagesWithRetry(channel.id);
      allMessages = allMessages.concat(messages.map(msg => ({
        ...msg,
        channel: channel.name,
        channelId: channel.id
      })));
    } catch (error) {
      console.error(`チャンネル ${channel.name} のエラー:`, error);
    }
  });
  
  return allMessages;
}

// ========= Bot参加済みチャンネル取得（改善版） =========
function getJoinedChannels() {
  const allChannels = [];
  
  // パブリックチャンネルを取得
  try {
    const publicResponse = slackAPI('conversations.list', {
      types: 'public_channel',
      exclude_archived: true,
      limit: 1000
    });
    
    const publicChannels = publicResponse.channels?.filter(channel => channel.is_member) || [];
    allChannels.push(...publicChannels);
    console.log(`取得したパブリックチャンネル数: ${publicChannels.length}`);
  } catch (e) {
    console.error('パブリックチャンネル取得エラー:', e);
  }
  
  // プライベートチャンネルを取得（groups:read スコープが必要）
  try {
    const privateResponse = slackAPI('conversations.list', {
      types: 'private_channel',
      exclude_archived: true,
      limit: 1000
    });
    
    // プライベートチャンネルは Bot がメンバーのもののみ取得可能
    const privateChannels = privateResponse.channels || [];
    allChannels.push(...privateChannels);
    console.log(`取得したプライベートチャンネル数: ${privateChannels.length}`);
  } catch (e) {
    // groups:read スコープがない場合はエラーになる
    if (e.toString().includes('missing_scope')) {
      console.warn('プライベートチャンネルへのアクセスには groups:read スコープが必要です');
    } else {
      console.error('プライベートチャンネル取得エラー:', e);
    }
  }
  
  console.log(`合計取得チャンネル数: ${allChannels.length}`);
  return allChannels;
}

// ========= チャンネルアクセス診断（改善版） =========
function diagnosePrivateChannelAccess(channelId) {
  const diagnostics = {
    channelId: channelId,
    hasAccess: false,
    channelFound: false,
    isPrivate: null,
    isMember: null,
    missingScopes: [],
    recommendations: [],
    errors: [],
    apiResponses: {}
  };
  
  console.log('\n=== チャンネルアクセス診断開始 ===');
  console.log(`対象チャンネル: ${channelId}`);
  
  // 重要: Botは自分がメンバーのチャンネルしか見られない
  console.log('\n※ 重要: Botはメンバーになっているチャンネルのみアクセス可能です');
  
  // Step 1: conversations.listでBotが参加しているチャンネルを確認
  console.log('\nStep 1: Botが参加しているチャンネルを検索...');
  
  // まずすべてのチャンネルを一度に取得
  try {
    const allChannelsResponse = slackAPI('conversations.list', {
      types: 'public_channel,private_channel',
      limit: 1000,
      exclude_archived: true
    });
    
    diagnostics.apiResponses.conversations_list = {
      ok: allChannelsResponse.ok,
      channel_count: allChannelsResponse.channels?.length || 0
    };
    
    if (allChannelsResponse.channels) {
      const channel = allChannelsResponse.channels.find(ch => ch.id === channelId);
      
      if (channel) {
        diagnostics.channelFound = true;
        diagnostics.isPrivate = channel.is_private || false;
        diagnostics.isMember = channel.is_member !== false; // Botはリストに表示されるチャンネルのメンバー
        diagnostics.channelName = channel.name;
        
        console.log(`✅ チャンネル発見: #${channel.name}`);
        console.log(`  - タイプ: ${diagnostics.isPrivate ? 'プライベート' : 'パブリック'}`);
        console.log(`  - Botメンバー: ${diagnostics.isMember ? 'はい' : 'いいえ'}`);
        console.log(`  - is_memberフラグ: ${channel.is_member}`);
      } else {
        console.log('❌ Botが参加しているチャンネルに見つかりません');
        diagnostics.recommendations.push('Botがチャンネルに参加していない可能性があります');
      }
    }
  } catch (listError) {
    diagnostics.errors.push(`conversations.list エラー: ${listError.toString()}`);
    console.error(`conversations.list エラー: ${listError}`);
    
    // スコープ不足の確認
    if (listError.toString().includes('missing_scope')) {
      const errorStr = listError.toString();
      if (errorStr.includes('channels:read')) {
        diagnostics.missingScopes.push('channels:read');
      }
      if (errorStr.includes('groups:read')) {
        diagnostics.missingScopes.push('groups:read');
      }
    }
  }
  
  // Step 2: conversations.infoで詳細情報を取得（オプション）
  console.log('\nStep 2: conversations.infoで詳細情報を取得...');
  try {
    const infoResponse = slackAPI('conversations.info', {
      channel: channelId,
      include_locale: false,
      include_num_members: true
    });
    
    diagnostics.apiResponses.conversations_info = {
      ok: infoResponse.ok,
      error: infoResponse.error
    };
    
    if (infoResponse.ok && infoResponse.channel) {
      const infoChannel = infoResponse.channel;
      
      // conversations.listで見つからなかった場合のみ更新
      if (!diagnostics.channelFound) {
        diagnostics.channelFound = true;
        diagnostics.isPrivate = infoChannel.is_private || false;
        diagnostics.isMember = infoChannel.is_member || false;
        diagnostics.channelName = infoChannel.name;
      }
      
      console.log(`conversations.info 結果:`);
      console.log(`  - チャンネル名: ${infoChannel.name}`);
      console.log(`  - is_member: ${infoChannel.is_member}`);
      console.log(`  - is_private: ${infoChannel.is_private}`);
      console.log(`  - num_members: ${infoChannel.num_members || 'N/A'}`);
    }
  } catch (infoError) {
    const errorStr = infoError.toString();
    diagnostics.apiResponses.conversations_info = {
      ok: false,
      error: errorStr
    };
    
    if (errorStr.includes('channel_not_found')) {
      diagnostics.errors.push('チャンネルが存在しないか、Botがアクセスできません');
      if (!diagnostics.channelFound) {
        diagnostics.recommendations.push('Botをチャンネルに招待する必要があります');
      }
    } else if (errorStr.includes('missing_scope')) {
      diagnostics.missingScopes.push('channels:read');
    } else {
      diagnostics.errors.push(`conversations.info エラー: ${errorStr}`);
    }
  }
  
  // Step 3: メッセージ取得テスト
  console.log('\nStep 3: conversations.historyでメッセージ取得テスト...');
  try {
    const historyResponse = slackAPI('conversations.history', {
      channel: channelId,
      limit: 1
    });
    
    diagnostics.apiResponses.conversations_history = {
      ok: historyResponse.ok,
      has_messages: (historyResponse.messages?.length || 0) > 0
    };
    
    if (historyResponse.ok) {
      diagnostics.hasAccess = true;
      console.log(`✅ メッセージ取得成功 (${historyResponse.messages?.length || 0}件)`);
    }
  } catch (historyError) {
    const errorStr = historyError.toString();
    diagnostics.apiResponses.conversations_history = {
      ok: false,
      error: errorStr
    };
    
    if (errorStr.includes('not_in_channel')) {
      diagnostics.errors.push('Botがチャンネルメンバーではありません');
      diagnostics.recommendations.push(`Slackで /invite @bot_name を実行してBotを招待してください`);
    } else if (errorStr.includes('channel_not_found')) {
      diagnostics.errors.push('チャンネルが存在しないか、Botがアクセスできません');
    } else if (errorStr.includes('missing_scope')) {
      const requiredScope = diagnostics.isPrivate ? 'groups:history' : 'channels:history';
      diagnostics.missingScopes.push(requiredScope);
      diagnostics.recommendations.push(`Slack Appの設定で ${requiredScope} スコープを追加してください`);
    } else if (errorStr.includes('invalid_auth')) {
      diagnostics.errors.push('認証エラー: Bot Tokenが無効です');
    } else {
      diagnostics.errors.push(`メッセージ取得エラー: ${errorStr}`);
    }
  }
  
  // 診断結果のサマリー
  console.log('\n=== 診断結果 ===');
  console.log(`チャンネルID: ${channelId}`);
  console.log(`Botが見つけたチャンネル: ${diagnostics.channelFound ? '✅ はい' : '❌ いいえ'}`);
  
  if (diagnostics.channelFound) {
    console.log(`チャンネル名: #${diagnostics.channelName}`);
    console.log(`チャンネルタイプ: ${diagnostics.isPrivate ? '🔒 プライベート' : '🌐 パブリック'}`);
    console.log(`Botメンバーシップ: ${diagnostics.isMember ? '✅ メンバー' : '❌ 非メンバー'}`);
    console.log(`メッセージアクセス: ${diagnostics.hasAccess ? '✅ 可能' : '❌ 不可'}`);
  } else {
    console.log('\n⚠️ Botがこのチャンネルを見つけられませんでした');
    console.log('可能性:');
    console.log('1. チャンネルIDが間違っている');
    console.log('2. プライベートチャンネルでBotがメンバーでない');
    console.log('3. チャンネルがアーカイブされている');
  }
  
  if (diagnostics.missingScopes.length > 0) {
    console.log(`\n不足しているスコープ:`);
    diagnostics.missingScopes.forEach(scope => console.log(`  - ${scope}`));
  }
  
  if (diagnostics.recommendations.length > 0) {
    console.log(`\n💡 推奨アクション:`);
    diagnostics.recommendations.forEach((rec, i) => console.log(`  ${i + 1}. ${rec}`));
  }
  
  // 重要な注意事項を追加
  if (!diagnostics.channelFound && diagnostics.isPrivate !== false) {
    diagnostics.recommendations.push('\n重要: プライベートチャンネルはBotがメンバーの場合のみアクセス可能です');
    diagnostics.recommendations.push('Slackでチャンネルに移動し、/invite @your_bot_name を実行してください');
  }
  
  if (diagnostics.errors.length > 0) {
    console.log(`\nエラー詳細:`);
    diagnostics.errors.forEach(err => console.log(`  - ${err}`));
  }
  
  return diagnostics;
}

// ========= OpenAI API呼び出し =========
function callOpenAIAPI(prompt, model = 'gpt-5') {
  
  // 通常のモデル（gpt-5）はResponses APIを使用
  const url = 'https://api.openai.com/v1/chat/completions';
  
  // プロンプト調整
  let systemMessage = 'あなたは業務分析の専門家です。必ず純粋なJSON形式のみで回答してください。マークダウンのコードブロック（```json）は使用しないでください。';
  let userMessage = prompt;
  
  let payload;
  
  if (false) {
    // (廃止) o3系特別処理
    userMessage = `${systemMessage}\n\n${prompt}\n\n重要: 回答は必ず有効なJSON形式のみで、余計な説明やマークダウンは一切含めないでください。`;
    payload = {
    model: model,
    messages: [
        { role: 'user', content: userMessage }
      ],
      max_completion_tokens: 2000
    };
  } else {
    payload = {
      model: model,
      messages: [
        { role: 'system', content: systemMessage },
        { role: 'user', content: userMessage }
    ],
    temperature: 0.3,
    max_tokens: 2000,
      response_format: { type: 'json_object' }
  };
  }
  
  const options = {
    method: 'post',
    headers: {
      'Authorization': `Bearer ${OPENAI_API_KEY}`,
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  
  try {
    const response = UrlFetchApp.fetch(url, options);
    const responseText = response.getContentText();
    
    // レスポンスが空の場合のチェック
    if (!responseText || responseText.trim() === '') {
      console.error('OpenAI APIから空のレスポンスが返されました');
      throw new Error('Empty response from OpenAI API');
    }
    
    let data;
    try {
      data = JSON.parse(responseText);
    } catch (parseError) {
      console.error('OpenAI APIレスポンスのJSONパースエラー:', parseError);
      console.error('レスポンステキスト:', responseText.substring(0, 500));
      throw new Error(`JSON parse error: ${parseError.message}`);
    }
    
    if (data.error) {
      // エラーメッセージにモデル情報を含める
      throw new Error(`OpenAI API Error (Model: ${model}): ${data.error.message}`);
    }
    
    const content = data.choices[0].message.content;
    
    console.log(`OpenAI API content type: ${typeof content}`);
    console.log(`OpenAI API content length: ${content ? content.length : 0}`);
    
    // contentが空でないことを確認
    if (!content || content.trim() === '') {
      console.error('OpenAI APIから空のコンテンツが返されました');
      throw new Error('Empty content from OpenAI API');
    }
    
    // o3モデルの場合、JSONとして返されることを期待しているが、
    // response_formatを使用していないため、手動で検証が必要
    if (false) {
      console.log('o3モデルの出力 (first 500 chars):', content.substring(0, 500));
      try {
        // コンテンツがJSONとして有効かチェック
        JSON.parse(content);
        console.log('o3モデルの出力: 有効なJSON形式です');
      } catch (jsonError) {
        console.warn('o3モデルの出力がJSON形式ではありません。');
        console.warn('JSONパースエラー:', jsonError.message);
        
        // JSON形式でない場合でも、とりあえずコンテンツを返す
        // 呼び出し側でパースエラーをハンドリングする
      }
    }
    
    return content;
  } catch (e) {
    console.error(`OpenAI APIエラー (使用モデル: ${model}):`, e);
    throw e;
  }
}

// ========= OpenAI Responses エンドポイント（現在は使用しない） =========
function callOpenAIResponsesEndpoint(prompt, model = 'gpt-5') {
  // Responsesエンドポイントは現在使用しないため、通常のAPIにフォールバック
  console.log('通常のOpenAI APIを使用します');
  return callOpenAIAPI(prompt, 'gpt-5');
}

// ========= 分析レスポンス用JSON Schema =========
function buildAnalysisResponseSchema() {
  return {
    type: 'object',
    properties: {
      categories: {
        type: 'array',
        items: { type: 'string' }
      },
      topics: {
        type: 'array',
        items: {
          type: 'object',
          properties: {
            title: { type: 'string' },
            description: { type: 'string' },
            priority: { type: 'number' }
          },
          required: ['title', 'description', 'priority']
        }
      },
      priority: {
        type: 'string',
        enum: ['HIGH', 'MEDIUM', 'LOW', 'INFO']
      },
      priorityReason: { type: 'string' },
      actionItems: {
        type: 'array',
        items: {
          type: 'object',
          properties: {
            task: { type: 'string' },
            owner: { type: 'string' },
            deadline: { type: 'string' }
          },
          required: ['task']
        }
      },
      stakeholders: {
        type: 'array',
        items: { type: 'string' }
      },
      urgency: {
        type: 'string',
        enum: ['critical', 'high', 'normal', 'low']
      },
      deadline: { type: 'string' },
      decisions: {
        type: 'array',
        items: { type: 'string' }
      },
      risks: {
        type: 'array',
        items: {
          type: 'object',
          properties: {
            risk: { type: 'string' },
            impact: { type: 'string' },
            mitigation: { type: 'string' }
          },
          required: ['risk']
        }
      },
      resources: {
        type: 'object',
        properties: {
          human: {
            type: 'array',
            items: { type: 'string' }
          },
          financial: { type: 'string' },
          time: { type: 'string' }
        }
      },
      kpis: {
        type: 'array',
        items: { type: 'string' }
      },
      summary: { type: 'string' }
    },
    required: ['categories', 'topics', 'priority', 'summary']
  };
}

// ========= エラー通知 =========
function sendErrorNotification(error, workflowResult) {
  if (REPORT_EMAIL) {
    try {
      MailApp.sendEmail({
        to: REPORT_EMAIL,
        subject: '【エラー】Slack統合システム処理エラー',
        body: `エラーが発生しました。\n\nエラー内容: ${error.toString()}\n\n処理状況:\n- メッセージ取得: ${workflowResult.messages?.length || 0}件\n- 分析完了: ${workflowResult.analysis ? 'Yes' : 'No'}\n- ガバナンスチェック: ${workflowResult.governance ? 'Yes' : 'No'}`
      });
    } catch (e) {
      console.log('エラー通知送信失敗:', e);
    }
  }
}

// ========= Botがアクセス可能なチャンネル一覧 =========
function listBotAccessibleChannels() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    console.log('Botがアクセス可能なチャンネルを取得中...');
    
    // Botが参加しているすべてのチャンネルを取得
    const response = slackAPI('conversations.list', {
      types: 'public_channel,private_channel',
      exclude_archived: true,
      limit: 1000
    });
    
    if (!response.ok) {
      ui.alert('エラー', `APIエラー: ${response.error}`, ui.ButtonSet.OK);
      return;
    }
    
    const channels = response.channels || [];
    
    // チャンネルをタイプ別に分類
    const publicChannels = channels.filter(ch => !ch.is_private && ch.is_member !== false);
    const privateChannels = channels.filter(ch => ch.is_private);
    
    let message = '🤖 Botがアクセス可能なチャンネル\n\n';
    
    message += `🌐 パブリックチャンネル (${publicChannels.length}件):\n`;
    if (publicChannels.length > 0) {
      publicChannels.forEach(ch => {
        message += `  #${ch.name} (${ch.id})\n`;
      });
    } else {
      message += '  なし\n';
    }
    
    message += `\n🔒 プライベートチャンネル (${privateChannels.length}件):\n`;
    if (privateChannels.length > 0) {
      privateChannels.forEach(ch => {
        message += `  #${ch.name} (${ch.id})\n`;
      });
    } else {
      message += '  なし\n';
      message += '\n※ プライベートチャンネルにアクセスするには：\n';
      message += '1. groups:read スコープが必要\n';
      message += '2. Botをチャンネルに招待する必要があります\n';
    }
    
    message += `\n合計: ${channels.length}チャンネル`;
    
    // 重要な注意事項
    message += '\n\n📌 重要な注意事項:\n';
    message += '• Botはメンバーとして招待されたチャンネルのみ表示されます\n';
    message += '• プライベートチャンネルは明示的な招待が必要です\n';
    message += '• /invite @bot_name でBotを招待できます';
    
    ui.alert('チャンネル一覧', message, ui.ButtonSet.OK);
    
  } catch (error) {
    console.error('チャンネル一覧取得エラー:', error);
    ui.alert('エラー', `チャンネル一覧の取得に失敗しました：\n${error.toString()}`, ui.ButtonSet.OK);
  }
}


// ========= ユーザー情報更新（メニュー用） =========
function refreshUserInfo() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    ui.alert('更新中', 'ユーザー情報を更新中...', ui.ButtonSet.OK);
    
    // ユーザー情報を再読み込み
    loadAllUsers();
    
    const userCount = Object.keys(userInfoCache).length;
    ui.alert('完了', `${userCount}人のユーザー情報を更新しました`, ui.ButtonSet.OK);
  } catch (error) {
    ui.alert('エラー', `ユーザー情報の更新に失敗しました: ${error.toString()}`, ui.ButtonSet.OK);
  }
}

/*
================================================================================
                                    終了
================================================================================
*/// ========= プライベートチャンネル完全デバッグ診断 =========
// この関数はプライベートチャンネルアクセス問題を根本的に診断します

function debugPrivateChannelsComplete() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    console.log('===== プライベートチャンネル完全診断開始 =====');
    
    let report = [];
    report.push('プライベートチャンネル完全診断レポート');
    report.push('=' .repeat(50));
    report.push('');
    
    // 1. Bot認証情報の確認
    report.push('【1. Bot認証情報】');
    const authInfo = slackAPI('auth.test', {});
    report.push(`Bot名: @${authInfo.user || 'unknown'}`);
    report.push(`Bot ID: ${authInfo.user_id || 'unknown'}`);
    report.push(`Team: ${authInfo.team || 'unknown'}`);
    report.push(`Token Type: ${SLACK_BOT_TOKEN.startsWith('xoxb-') ? 'Bot Token ✅' : 'User Token ⚠️'}`);
    report.push('');
    
    // 2. 必要なスコープの確認
    report.push('【2. 必要なスコープの確認】');
    report.push('プライベートチャンネルに必要なスコープ:');
    report.push('- groups:read（プライベートチャンネル一覧取得）');
    report.push('- groups:history（プライベートチャンネル履歴取得）');
    report.push('');
    
    // 3. conversations.listでプライベートチャンネルのみを取得
    report.push('【3. プライベートチャンネル取得テスト】');
    
    // 3-1. 重要: Slack APIの問題を回避するため、全チャンネルを取得してからフィルタリング
    console.log('全チャンネルを取得してフィルタリング...');
    
    // まず全チャンネルを取得（ページネーション対応）
    const allChannels = [];
    let cursor = '';
    
    do {
      const params = {
        limit: 200,  // typesを指定しない、または個別に取得
        exclude_archived: true
      };
      if (cursor) params.cursor = cursor;
      
      const response = slackAPI('conversations.list', params);
      if (response.ok && response.channels) {
        allChannels.push(...response.channels);
        cursor = response.response_metadata?.next_cursor || '';
      } else {
        break;
      }
    } while (cursor);
    
    const allChannelsResp = {
      ok: true,
      channels: allChannels
    };
    
    // プライベートチャンネルを正しくフィルタリング
    let privateChannels = [];
    if (allChannelsResp.ok && allChannelsResp.channels) {
      privateChannels = allChannelsResp.channels.filter(ch => {
        // プライベートチャンネルの正しい判定（is_privateフラグのみで判定）
        return ch.is_private === true;
      });
    }
    
    // レガシー互換性のため、privateResponseとして扱う
    const privateResponse = {
      ok: allChannelsResp.ok,
      channels: privateChannels,
      error: allChannelsResp.error
    };
    
    if (!privateResponse.ok) {
      report.push(`❌ エラー: ${privateResponse.error}`);
      if (privateResponse.error === 'missing_scope') {
        report.push('→ groups:read スコープが不足しています');
      }
    } else {
      const privateChannels = privateResponse.channels || [];
      report.push(`プライベートチャンネル数: ${privateChannels.length}個`);
      
      // 詳細表示
      if (privateChannels.length > 0) {
        report.push('');
        report.push('検出されたプライベートチャンネル:');
        privateChannels.forEach((ch, i) => {
          report.push(`${i + 1}. #${ch.name} (${ch.id})`);
          report.push(`   - is_member: ${ch.is_member ? '✅' : '❌'}`);
        });
      }
      
      if (privateChannels.length === 0) {
        report.push('⚠️ プライベートチャンネルが0個です');
        report.push('考えられる原因:');
        report.push('1. Botがどのプライベートチャンネルにも招待されていない');
        report.push('2. ワークスペースにプライベートチャンネルが存在しない');
      } else {
        report.push('');
        report.push('取得したプライベートチャンネル:');
        privateChannels.slice(0, 5).forEach((ch, i) => {
          report.push(`${i + 1}. #${ch.name} (${ch.id})`);
          report.push(`   - is_private: ${ch.is_private}`);
          report.push(`   - is_member: ${ch.is_member}`);
          report.push(`   - is_channel: ${ch.is_channel}`);
          report.push(`   - is_group: ${ch.is_group}`);
        });
        if (privateChannels.length > 5) {
          report.push(`... 他 ${privateChannels.length - 5} チャンネル`);
        }
      }
    }
    
    report.push('');
    
    // 4. パブリックチャンネルとプライベートチャンネルを両方取得して比較
    report.push('【4. 全チャンネル取得テスト（パブリック＋プライベート）】');
    
    const allResponse = slackAPI('conversations.list', {
      types: 'public_channel,private_channel',
      limit: 1000,
      exclude_archived: true
    });
    
    if (allResponse.ok) {
      const allChannels = allResponse.channels || [];
      
      // チャンネルIDのプレフィックスで分類
      const cChannels = allChannels.filter(ch => ch.id && ch.id.startsWith('C'));
      const gChannels = allChannels.filter(ch => ch.id && ch.id.startsWith('G'));
      const otherChannels = allChannels.filter(ch => ch.id && !ch.id.startsWith('C') && !ch.id.startsWith('G'));
      
      // is_privateフラグで分類
      const privateByFlag = allChannels.filter(ch => ch.is_private === true);
      const publicByFlag = allChannels.filter(ch => ch.is_private === false || ch.is_private === undefined);
      
      report.push(`全チャンネル数: ${allChannels.length}個`);
      report.push('');
      report.push('IDプレフィックスによる分類:');
      report.push(`- Cで始まる（通常パブリック）: ${cChannels.length}個`);
      report.push(`- Gで始まる（通常プライベート）: ${gChannels.length}個`);
      report.push(`- その他: ${otherChannels.length}個`);
      report.push('');
      report.push('is_privateフラグによる分類:');
      report.push(`- is_private=true: ${privateByFlag.length}個`);
      report.push(`- is_private=false/undefined: ${publicByFlag.length}個`);
      
      // 不一致の検出
      report.push('');
      report.push('【ID と is_private フラグの不一致チェック】');
      const mismatches = [];
      
      allChannels.forEach(ch => {
        const expectedPrivate = ch.id && ch.id.startsWith('G');
        const actualPrivate = ch.is_private === true;
        
        if (expectedPrivate !== actualPrivate) {
          mismatches.push({
            name: ch.name,
            id: ch.id,
            expectedPrivate: expectedPrivate,
            actualPrivate: actualPrivate
          });
        }
      });
      
      if (mismatches.length > 0) {
        report.push(`⚠️ ${mismatches.length}個のチャンネルで不一致を検出:`);
        mismatches.slice(0, 5).forEach(m => {
          report.push(`- #${m.name} (${m.id}): ID判定=${m.expectedPrivate}, フラグ=${m.actualPrivate}`);
        });
      } else {
        report.push('✅ すべてのチャンネルでIDとフラグが一致');
      }
    }
    
    report.push('');
    
    // 5. 特定のプライベートチャンネルへのアクセステスト
    report.push('【5. プライベートチャンネルアクセステスト】');
    
    if (privateResponse.ok && privateResponse.channels && privateResponse.channels.length > 0) {
      const testChannel = privateResponse.channels[0];
      report.push(`テスト対象: #${testChannel.name} (${testChannel.id})`);
      
      // conversations.history でアクセステスト
      try {
        const historyResponse = slackAPI('conversations.history', {
          channel: testChannel.id,
          limit: 1
        });
        
        if (historyResponse.ok) {
          report.push('✅ メッセージ履歴にアクセス可能');
        } else {
          report.push(`❌ アクセス不可: ${historyResponse.error}`);
          if (historyResponse.error === 'not_in_channel') {
            report.push('→ Botがチャンネルメンバーではありません');
          }
        }
      } catch (e) {
        report.push(`❌ エラー: ${e.toString()}`);
      }
    } else {
      report.push('テスト対象のプライベートチャンネルがありません');
    }
    
    report.push('');
    report.push('【6. 推奨アクション】');
    
    // プライベートチャンネルが0の場合の対処法
    if (!privateResponse.channels || privateResponse.channels.length === 0) {
      report.push('プライベートチャンネルにアクセスするには:');
      report.push('');
      report.push('1. Slack App の設定を確認:');
      report.push('   - https://api.slack.com/apps でアプリを選択');
      report.push('   - OAuth & Permissions → Scopes で以下を確認:');
      report.push('     ✓ groups:read');
      report.push('     ✓ groups:history');
      report.push('');
      report.push('2. アプリを再インストール:');
      report.push('   - スコープ追加後、"Reinstall to Workspace" をクリック');
      report.push('');
      report.push('3. プライベートチャンネルにBotを招待:');
      report.push('   - 各プライベートチャンネルで: /invite @' + (authInfo.user || 'bot-name'));
      report.push('   - または: チャンネル設定 → Integrations → Add apps');
      report.push('');
      report.push('4. Bot Token の確認:');
      report.push('   - xoxb- で始まるBot Tokenを使用しているか確認');
      report.push('   - User Token (xoxp-) では制限がある場合があります');
    }
    
    // 結果を表示
    const resultText = report.join('\n');
    console.log(resultText);
    
    // UIに表示（長すぎる場合は最初の部分のみ）
    const displayText = resultText.length > 3000 ? 
      resultText.substring(0, 2900) + '\n\n... (詳細はログを確認してください)' :
      resultText;
    
    ui.alert('診断結果', displayText, ui.ButtonSet.OK);
    
  } catch (error) {
    console.error('診断中にエラー:', error);
    ui.alert('エラー', `診断中にエラーが発生しました:\n${error.toString()}`, ui.ButtonSet.OK);
  }
}

// ========= プライベートチャンネルアクセステスト =========
/**
 * 検出されたプライベートチャンネルへのアクセステスト
 */
function testPrivateChannelAccess() {
  console.log('===== プライベートチャンネルアクセステスト =====');
  
  // 動的にプライベートチャンネルを検出
  const allChannels = [];
  let cursor = '';
  
  do {
    const params = {
      types: 'public_channel,private_channel',
      limit: 200,
      exclude_archived: true
    };
    if (cursor) params.cursor = cursor;
    
    const response = slackAPI('conversations.list', params);
    if (response.ok && response.channels) {
      allChannels.push(...response.channels);
      cursor = response.response_metadata?.next_cursor || '';
    } else {
      break;
    }
  } while (cursor);
  
  // プライベートチャンネルのみフィルタ
  const privateChannels = allChannels.filter(ch => ch.is_private === true);
  
  if (privateChannels.length === 0) {
    console.log('プライベートチャンネルが見つかりません');
    showAlertSafely('情報', 'プライベートチャンネルが見つかりません');
    return [];
  }
  
  const results = [];
  
  privateChannels.forEach(channel => {
    console.log(`\nテスト中: #${channel.name} (${channel.id})`);
    
    // 1. チャンネル情報取得
    try {
      const info = slackAPI('conversations.info', { channel: channel.id });
      if (info.ok && info.channel) {
        console.log(`✅ チャンネル情報取得成功`);
        console.log(`  - is_private: ${info.channel.is_private}`);
        console.log(`  - is_member: ${info.channel.is_member}`);
        console.log(`  - num_members: ${info.channel.num_members || 'N/A'}`);
        
        results.push({
          channel: channel.name,
          id: channel.id,
          infoAccess: true,
          isMember: info.channel.is_member,
          isPrivate: info.channel.is_private
        });
      }
    } catch (e) {
      console.log(`❌ チャンネル情報取得失敗: ${e.toString()}`);
      results.push({
        channel: channel.name,
        id: channel.id,
        infoAccess: false,
        error: e.toString()
      });
    }
    
    // 2. メッセージ履歴取得テスト
    try {
      const history = slackAPI('conversations.history', {
        channel: channel.id,
        limit: 1
      });
      
      if (history.ok) {
        console.log(`✅ メッセージ履歴にアクセス可能`);
        console.log(`  - メッセージ数: ${history.messages ? history.messages.length : 0}`);
        
        results[results.length - 1].historyAccess = true;
        results[results.length - 1].messageCount = history.messages ? history.messages.length : 0;
      }
    } catch (e) {
      console.log(`❌ メッセージ履歴アクセス失敗: ${e.toString()}`);
      
      if (e.toString().includes('not_in_channel')) {
        console.log('→ Botをチャンネルに招待してください:');
        console.log(`   /invite @kushim_slack_governan`);
      }
      
      results[results.length - 1].historyAccess = false;
      results[results.length - 1].historyError = e.toString();
    }
  });
  
  // 結果サマリー
  console.log('\n===== テスト結果サマリー =====');
  results.forEach(r => {
    console.log(`\n#${r.channel} (${r.id}):`);
    console.log(`  - チャンネル情報: ${r.infoAccess ? '✅' : '❌'}`);
    console.log(`  - プライベート: ${r.isPrivate ? '✅' : '❌'}`);
    console.log(`  - Botメンバー: ${r.isMember ? '✅' : '❌'}`);
    console.log(`  - メッセージ履歴: ${r.historyAccess ? '✅' : '❌'}`);
    
    if (!r.isMember && r.infoAccess) {
      console.log(`  📌 アクション: /invite @kushim_slack_governan を実行`);
    }
  });
  
  // UI表示（可能な場合）
  let message = 'プライベートチャンネルアクセステスト結果:\n\n';
  
  results.forEach(r => {
    message += `【${r.channel}】\n`;
    message += `・チャンネル情報: ${r.infoAccess ? '取得可能' : '取得不可'}\n`;
    message += `・Botメンバー: ${r.isMember ? 'はい' : 'いいえ'}\n`;
    message += `・メッセージ履歴: ${r.historyAccess ? 'アクセス可能' : 'アクセス不可'}\n`;
    
    if (!r.isMember && r.infoAccess) {
      message += `→ アクション: /invite @kushim_slack_governan\n`;
    }
      message += '\n';
  });
  
  showAlertSafely('テスト結果', message);
  
  return results;
}

/**
 * プライベートチャンネルのメッセージを取得して保存
 */
function syncPrivateChannels() {
  console.log('===== プライベートチャンネル同期開始 =====');
  
  // 動的にプライベートチャンネルを検出
  const allChannels = [];
  let cursor = '';
  
  do {
    const params = {
      types: 'public_channel,private_channel',
      limit: 200,
      exclude_archived: true
    };
    if (cursor) params.cursor = cursor;
    
    const response = slackAPI('conversations.list', params);
    if (response.ok && response.channels) {
      allChannels.push(...response.channels);
      cursor = response.response_metadata?.next_cursor || '';
    } else {
      break;
    }
  } while (cursor);
  
  // プライベートチャンネルかつBotがメンバーのチャンネルのみフィルタ
  const privateChannels = allChannels.filter(ch => 
    ch.is_private === true && ch.is_member === true
  );
  
  if (privateChannels.length === 0) {
    console.log('アクセス可能なプライベートチャンネルがありません');
    showAlertSafely('情報', 'アクセス可能なプライベートチャンネルがありません。\nBotを招待してください: /invite @kushim_slack_governan');
    return { success: 0, total: 0, messages: 0 };
  }
  
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let messagesSheet = ss.getSheetByName('Messages');
  if (!messagesSheet) {
    console.log('Messagesシートを作成中...');
    messagesSheet = createMessagesSheet(ss);
  }
  
  let totalMessages = 0;
  let successChannels = 0;
  
  privateChannels.forEach(channel => {
    console.log(`\n処理中: #${channel.name} (${channel.id})`);
    
    try {
      // メッセージ履歴取得
      const history = slackAPI('conversations.history', {
        channel: channel.id,
        limit: 100
      });
      
      if (history.ok && history.messages) {
        const messages = history.messages;
        console.log(`  - ${messages.length}件のメッセージを取得`);
        
        // メッセージを保存
        const rows = [];
        messages.forEach(msg => {
          const messageId = `${channel.id}_${msg.ts}`;
          const permalink = `https://slack.com/archives/${channel.id}/p${msg.ts.replace('.', '')}`;
          const messageDate = new Date(Number(msg.ts.split('.')[0]) * 1000);
          
          rows.push([
            messageId,           // id
            channel.id,          // channel_id
            msg.ts,              // message_ts
            msg.thread_ts || '', // thread_ts
            msg.text || '',      // text_raw
            msg.user || '',      // user_name
            '',                  // summary_json
            '',                  // classification_json
            '',                  // match_flag
            '',                  // human_judgement
            permalink,           // permalink
            '',                  // checklist_proposed
            '',                  // agenda_selected
            '',                  // draft_doc_url
            messageDate.toISOString(), // timestamp
            '',                  // reactions
            ''                   // files
          ]);
        });
        
        if (rows.length > 0) {
          const lastRow = messagesSheet.getLastRow();
          messagesSheet.getRange(lastRow + 1, 1, rows.length, rows[0].length)
            .setValues(rows);
          console.log(`  - ${rows.length}件を保存`);
        }
        
        totalMessages += messages.length;
        successChannels++;
        
      } else {
        console.log(`  - メッセージ取得失敗: ${history.error}`);
      }
      
    } catch (e) {
      console.log(`  - エラー: ${e.toString()}`);
      
      if (e.toString().includes('not_in_channel')) {
        console.log('  → Botを招待してください: /invite @kushim_slack_governan');
      }
    }
  });
  
  console.log('\n===== 同期完了 =====');
  console.log(`成功: ${successChannels}/${privateChannels.length}チャンネル`);
  console.log(`取得: ${totalMessages}メッセージ`);
  
  const summary = `プライベートチャンネル同期完了\n\n` +
    `処理: ${successChannels}/${privateChannels.length}チャンネル\n` +
    `取得: ${totalMessages}メッセージ`;
  
  showAlertSafely('同期結果', summary);
  
  return {
    success: successChannels,
    total: privateChannels.length,
    messages: totalMessages
  };
}

// メニューに追加するための関数
function addDebugMenuItems() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🔧 デバッグ')
    .addItem('プライベートチャンネル完全診断', 'debugPrivateChannelsComplete')
    .addToUi();
}

// ========= channel_not_foundエラー診断 =========
/**
 * channel_not_foundエラーの診断
 * チャンネルID C08UASCBHRB のアクセス問題を調査
 */
function diagnoseChannelNotFoundError() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    console.log('===== channel_not_found エラー診断開始 =====');
    
    let report = [];
    report.push('Channel Not Found エラー診断レポート');
    report.push('=' .repeat(50));
    report.push('');
    
    // 問題のチャンネルID
    const problematicChannelId = 'C08UASCBHRB';
    report.push('【問題のチャンネル】');
    report.push(`チャンネルID: ${problematicChannelId}`);
    report.push('');
    
    // 1. Bot認証情報の確認
    report.push('【1. Bot認証情報】');
    const authInfo = slackAPI('auth.test', {});
    report.push(`Bot名: @${authInfo.user || 'unknown'}`);
    report.push(`Bot ID: ${authInfo.user_id || 'unknown'}`);
    report.push(`Team: ${authInfo.team || 'unknown'}`);
    report.push('');
    
    // 2. チャンネル情報の取得試行
    report.push(`【2. チャンネル ${problematicChannelId} の情報取得試行】`);
    
    try {
      const channelInfo = slackAPI('conversations.info', {
        channel: problematicChannelId
      });
      
      if (channelInfo.ok) {
        report.push('✅ チャンネル情報取得成功！');
        report.push(`チャンネル名: #${channelInfo.channel.name}`);
        report.push(`プライベート: ${channelInfo.channel.is_private ? 'はい' : 'いいえ'}`);
        report.push(`アーカイブ済み: ${channelInfo.channel.is_archived ? 'はい' : 'いいえ'}`);
        report.push(`Botメンバー: ${channelInfo.channel.is_member ? '✅' : '❌'}`);
        
        if (!channelInfo.channel.is_member) {
          report.push('');
          report.push('⚠️ Botがメンバーではありません！');
          report.push('');
          report.push('【解決方法】');
          report.push('Slackでこのチャンネルに移動して、以下のコマンドを実行:');
          report.push(`/invite @${authInfo.user}`);
        }
      }
  } catch (error) {
      const errorStr = error.toString();
      report.push(`❌ チャンネル情報取得失敗: ${errorStr}`);
      
      if (errorStr.includes('channel_not_found')) {
        report.push('');
        report.push('【エラー分析】');
        report.push('チャンネルが見つかりません。考えられる原因:');
        report.push('1. チャンネルが削除された');
        report.push('2. チャンネルIDが間違っている');
        report.push('3. 別のワークスペースのチャンネルID');
        report.push('4. プライベートチャンネルでBotが一度も招待されていない');
      }
    }
    
    // 3. 全チャンネルリストでの存在確認
    report.push('');
    report.push('【3. 全チャンネルリストでの検索】');
    
    const allChannels = [];
    let cursor = '';
    
    do {
      const params = {
        types: 'public_channel,private_channel',
        limit: 1000,
        exclude_archived: false  // アーカイブ済みも含める
      };
      if (cursor) params.cursor = cursor;
      
      const response = slackAPI('conversations.list', params);
      if (response.ok && response.channels) {
        allChannels.push(...response.channels);
        cursor = response.response_metadata?.next_cursor || '';
      } else {
        break;
      }
    } while (cursor);
    
    const foundChannel = allChannels.find(ch => ch.id === problematicChannelId);
    
    if (foundChannel) {
      report.push(`✅ チャンネルがリストに存在します`);
      report.push(`名前: #${foundChannel.name}`);
      report.push(`プライベート: ${foundChannel.is_private ? 'はい' : 'いいえ'}`);
      report.push(`アーカイブ済み: ${foundChannel.is_archived ? 'はい' : 'いいえ'}`);
      report.push(`Botメンバー: ${foundChannel.is_member ? '✅' : '❌'}`);
      
      if (foundChannel.is_archived) {
        report.push('');
        report.push('⚠️ チャンネルはアーカイブされています');
        report.push('アーカイブされたチャンネルにはメッセージを送信できません');
      }
    } else {
      report.push('❌ チャンネルがリストに存在しません');
      report.push('');
      report.push('【可能性のある原因】');
      report.push('• チャンネルが削除された');
      report.push('• 異なるワークスペースのチャンネルID');
      report.push('• チャンネルIDが誤って記録された');
    }
    
    // 4. 通知設定の確認
    report.push('');
    report.push('【4. 通知設定の確認】');
    
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const configSheet = ss.getSheetByName(SHEETS.CONFIG);
    
    if (configSheet) {
      const configData = configSheet.getDataRange().getValues();
      const notifyChannelRow = configData.find(row => row[0] === 'notifySlackChannel');
      
      if (notifyChannelRow && notifyChannelRow[1]) {
        report.push(`現在の通知先チャンネル: ${notifyChannelRow[1]}`);
        
        if (notifyChannelRow[1] === problematicChannelId) {
          report.push('⚠️ このチャンネルが通知先に設定されています');
          report.push('');
          report.push('【推奨アクション】');
          report.push('1. 別の有効なチャンネルIDに変更する');
          report.push('2. または、Botをこのチャンネルに招待する');
        }
      } else {
        report.push('通知先チャンネルが設定されていません');
      }
    }
    
    // 5. 推奨される解決策
    report.push('');
    report.push('【推奨される解決策】');
    report.push('');
    
    if (foundChannel && !foundChannel.is_member) {
      report.push('1. Botをチャンネルに招待:');
      report.push(`   Slackで #${foundChannel.name} チャンネルに移動`);
      report.push(`   /invite @${authInfo.user} を実行`);
    } else if (foundChannel && foundChannel.is_archived) {
      report.push('1. チャンネルのアーカイブを解除するか、');
      report.push('2. 別のアクティブなチャンネルを通知先に設定');
    } else if (!foundChannel) {
      report.push('1. 通知先チャンネルIDを確認して修正');
      report.push('2. 正しいチャンネルIDを「⚙️ 設定」→「🎯 通知設定」で設定');
    }
    
    // 結果を表示
    const resultText = report.join('\n');
    console.log(resultText);
    
    // UIに表示
    ui.alert('診断結果', resultText, ui.ButtonSet.OK);
    
  } catch (error) {
    console.error('診断中にエラー:', error);
    ui.alert('エラー', `診断中にエラーが発生しました:\n${error.toString()}`, ui.ButtonSet.OK);
  }
}
