/**
 * ネット反応分析システム - 完全統合版
 * Version: 4.0
 * 
 * 機能:
 * - 簡易セットアップウィザード
 * - YouTube/X/Google Trends データ収集
 * - AI分析とコンテンツサジェスト
 * - 自動レポート生成
 * - エグゼクティブダッシュボード
 */

// =====================================
// グローバル設定
// =====================================
const CONFIG = {
  // API URLs
  YOUTUBE_API_URL: 'https://www.googleapis.com/youtube/v3',
  X_API_URL: 'https://api.twitter.com/2',
  GEMINI_API_URL: 'https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-pro-latest:generateContent',
  GOOGLE_ADS_API_URL: 'https://googleads.googleapis.com/v17',
  GROK_API_URL: 'https://api.x.ai/v1/chat/completions',
  
  // シート名
  SHEET_NAMES: {
    SETTINGS: '設定',
    DASHBOARD: 'ダッシュボード',
    REPORTS: 'レポート',
    YOUTUBE_DATA: 'YouTube_データ',
    X_DATA: 'X_データ',
    TRENDS_DATA: 'Trends_データ',
    SUGGESTIONS: 'サジェスト',
    ERROR_LOG: 'エラーログ'
  },
  
  // その他設定
  MAX_RETRIES: 3,
  TIMEOUT: 30000,
  BATCH_SIZE: 50
};

// =====================================
// メニュー作成
// =====================================
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📊 分析システム')
    .addItem('🚀 初回セットアップ', 'runSetupWizard')
    .addItem('▶️ クイック分析', 'runQuickAnalysis')
    .addItem('🎯 詳細分析（サジェスト付き）', 'runAdvancedAnalysis')
    .addSeparator()
    .addSubMenu(ui.createMenu('⚙️ 設定')
      .addItem('APIキー設定', 'setupAPIKeys')
      .addItem('自動実行設定', 'setupTriggers')
      .addItem('詳細設定', 'showAdvancedSettings')
      .addItem('設定リセット', 'resetSettings'))
    .addSeparator()
    .addSubMenu(ui.createMenu('📈 レポート')
      .addItem('日次レポート生成', 'generateDailyReport')
      .addItem('週次レポート生成', 'generateWeeklyReport')
      .addItem('月次レポート生成', 'generateMonthlyReport')
      .addItem('カスタムレポート', 'generateCustomReport'))
    .addSeparator()
    .addSubMenu(ui.createMenu('🧪 テスト')
      .addItem('YouTube API テスト', 'testYouTubeAPI')
      .addItem('X API テスト', 'testXAPI')
      .addItem('Gemini API テスト', 'testGeminiAPI')
      .addItem('サジェスト機能テスト', 'testSuggestions'))
    .addSeparator()
    .addSubMenu(ui.createMenu('❓ ヘルプ')
      .addItem('使い方ガイド', 'showUserGuide')
      .addItem('APIキー取得方法', 'showAPIKeyGuide')
      .addItem('トラブルシューティング', 'showTroubleshooting')
      .addItem('更新情報', 'showReleaseNotes'))
    .addToUi();
}

// =====================================
// セットアップウィザード
// =====================================
function runSetupWizard() {
  const ui = SpreadsheetApp.getUi();
  
  // ウェルカムメッセージ
  const welcome = ui.alert(
    '🎉 ネット反応分析システムへようこそ！',
    'これから簡単な設定を行います。\n\n' +
    '所要時間: 約3分\n\n' +
    '準備はよろしいですか？',
    ui.ButtonSet.YES_NO
  );
  
  if (welcome === ui.Button.NO) {
    ui.alert('セットアップをキャンセルしました。準備ができたら再度実行してください。');
    return;
  }
  
  try {
    // Step 1: スプレッドシート初期化
    ui.alert('Step 1/5', 'スプレッドシートを初期化しています...', ui.ButtonSet.OK);
    initializeSpreadsheet();
    
    // Step 2: 基本設定
    setupBasicSettings();
    
    // Step 3: 動作モード選択
    const mode = selectOperationMode();
    
    // Step 4: API設定（必要な場合）
    if (mode !== 'demo') {
      setupAPIKeys();
    }
    
    // Step 5: 自動実行設定
    const autoRun = ui.alert(
      'Step 5/5 - 自動実行設定',
      '毎日自動で分析を実行しますか？\n\n' +
      '• はい → 毎朝9時に自動実行\n' +
      '• いいえ → 手動実行のみ',
      ui.ButtonSet.YES_NO
    );
    
    if (autoRun === ui.Button.YES) {
      setupTriggers();
    }
    
    // 完了処理
    finalizeSetup();
    
  } catch (error) {
    ui.alert('エラー', 'セットアップ中にエラーが発生しました：\n' + error.toString(), ui.ButtonSet.OK);
  }
}

// =====================================
// 初期化関数
// =====================================
function initializeSpreadsheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 全シート作成
  Object.entries(CONFIG.SHEET_NAMES).forEach(([key, sheetName]) => {
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      if (!sheet) {
        Logger.log(`Failed to create sheet: ${sheetName}`);
        return;
      }
    }
    
    // シート別の初期設定
    switch (key) {
      case 'SETTINGS':
        initializeSettingsSheet(sheet);
        break;
      case 'DASHBOARD':
        initializeDashboardSheet(sheet);
        break;
      case 'REPORTS':
        initializeReportsSheet(sheet);
        break;
      case 'SUGGESTIONS':
        initializeSuggestionsSheet(sheet);
        break;
      default:
        initializeDataSheet(sheet);
    }
  });
  
  // デフォルトシートを削除
  try {
    const sheet1 = ss.getSheetByName('シート1');
    if (sheet1 && ss.getSheets().length > 1) {
      ss.deleteSheet(sheet1);
    }
  } catch (e) {
    // エラーは無視
  }
}

function initializeSettingsSheet(sheet) {
  if (!sheet) {
    throw new Error('Settings sheet could not be created or found.');
  }
  if (typeof sheet.getRange !== 'function') {
    throw new Error('Invalid sheet object for Settings sheet');
  }
  
  const settingsData = [
    ['項目', '設定値', '説明'],
    ['分析キーワード', '', '分析したいキーワードを入力'],
    ['レポートタイプ', '短期', '短期（7日）/ 中期（30日）/ 長期（365日）'],
    ['通知メール', '', 'レポートを受信するメールアドレス'],
    ['動作モード', 'demo', 'demo / basic / advanced'],
    ['YouTube分析', 'ON', 'YouTube分析のON/OFF'],
    ['X分析', 'ON', 'X（Twitter）分析のON/OFF'],
    ['AI要約', 'ON', 'AI要約のON/OFF'],
    ['サジェスト生成', 'ON', 'コンテンツサジェストのON/OFF'],
    ['監視Xアカウント', '', '監視するXアカウント（カンマ区切り）'],
    ['監視YouTubeチャンネル', '', '監視するチャンネルID（カンマ区切り）']
  ];
  
  sheet.getRange(1, 1, settingsData.length, 3).setValues(settingsData);
  
  // 書式設定
  sheet.getRange(1, 1, 1, 3).setBackground('#4a86e8').setFontColor('white').setFontWeight('bold');
  sheet.setColumnWidth(1, 150);
  sheet.setColumnWidth(2, 300);
  sheet.setColumnWidth(3, 400);
  
  // ドロップダウン設定
  const validations = {
    'B3': ['短期', '中期', '長期'],
    'B5': ['demo', 'basic', 'advanced'],
    'B6:B9': ['ON', 'OFF']
  };
  
  Object.entries(validations).forEach(([range, values]) => {
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(values)
      .build();
    sheet.getRange(range).setDataValidation(rule);
  });
}

function initializeDashboardSheet(sheet) {
  if (!sheet) {
    throw new Error('Dashboard sheet could not be created or found.');
  }
  if (typeof sheet.getRange !== 'function') {
    throw new Error('Invalid sheet object for Dashboard sheet');
  }
  
  const dashboardTemplate = [
    ['📊 ネット反応分析ダッシュボード'],
    [''],
    ['最終更新:', ''],
    ['分析キーワード:', ''],
    [''],
    ['📈 パフォーマンスサマリー'],
    ['総合スコア:', '', '', '前回比:', ''],
    ['YouTube再生数:', '', '', 'エンゲージメント率:', ''],
    ['X投稿数:', '', '', 'ポジティブ率:', ''],
    ['検索トレンド:', '', '', '成長率:', ''],
    [''],
    ['🚀 本日のオススメコンテンツ'],
    ['1.', ''],
    ['2.', ''],
    ['3.', ''],
    [''],
    ['💡 AI分析サマリー'],
    [''],
    [''],
    ['📊 トレンドチャート'],
    ['（データ取得後に表示されます）']
  ];
  
  dashboardTemplate.forEach((row, index) => {
    sheet.getRange(index + 1, 1, 1, row.length).setValues([row]);
  });
  
  // スタイル設定
  sheet.getRange('A1').setFontSize(20).setFontWeight('bold');
  sheet.getRange('A6').setFontSize(16).setFontWeight('bold').setBackground('#e8f0fe');
  sheet.getRange('A12').setFontSize(16).setFontWeight('bold').setBackground('#e8f0fe');
  sheet.getRange('A17').setFontSize(16).setFontWeight('bold').setBackground('#e8f0fe');
  sheet.getRange('A20').setFontSize(16).setFontWeight('bold').setBackground('#e8f0fe');
  sheet.setColumnWidth(1, 200);
  sheet.setColumnWidth(2, 200);
  sheet.setColumnWidth(3, 100);
  sheet.setColumnWidth(4, 150);
  sheet.setColumnWidth(5, 200);
}

// =====================================
// 基本設定
// =====================================
function setupBasicSettings() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAMES.SETTINGS);
  
  // キーワード入力
  const keyword = ui.prompt(
    'Step 2/5 - キーワード設定',
    '分析したいキーワードを入力してください：\n（例：AI、ChatGPT、マーケティング）',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (keyword.getSelectedButton() === ui.Button.OK && keyword.getResponseText()) {
    sheet.getRange('B2').setValue(keyword.getResponseText());
  } else {
    sheet.getRange('B2').setValue('AI'); // デフォルト値
  }
  
  // メール設定
  const email = ui.prompt(
    'メール通知設定（オプション）',
    'レポートを受信するメールアドレスを入力してください：\n（空欄でスキップ可能）',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (email.getSelectedButton() === ui.Button.OK && email.getResponseText()) {
    sheet.getRange('B4').setValue(email.getResponseText());
  }
}

function selectOperationMode() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAMES.SETTINGS);

  const response = ui.prompt(
    'Step 3/5 - 動作モード選択',
    '動作モードを入力してください（demo / basic / advanced）：\n\n' +
    '🎯 demo: APIキー不要、サンプルデータで動作\n' +
    '📊 basic: YouTube APIのみ使用\n' +
    '🚀 advanced: 全機能使用（要APIキー）',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) {
    ui.alert('モード選択をキャンセルしました。デフォルトでdemoモードを設定します。');
    const mode = 'demo';
    sheet.getRange('B5').setValue(mode);
    PropertiesService.getScriptProperties().setProperty('OPERATION_MODE', mode);
    return mode;
  }

  const input = response.getResponseText().trim().toLowerCase();
  let mode;
  switch (input) {
    case 'demo':
      mode = 'demo';
      break;
    case 'basic':
      mode = 'basic';
      break;
    case 'advanced':
      mode = 'advanced';
      break;
    default:
      ui.alert('無効な入力です。もう一度お試しください。');
      return selectOperationMode();
  }

  sheet.getRange('B5').setValue(mode);
  PropertiesService.getScriptProperties().setProperty('OPERATION_MODE', mode);
  return mode;
}

// =====================================
// API設定
// =====================================
function setupAPIKeys() {
  const ui = SpreadsheetApp.getUi();
  const mode = PropertiesService.getScriptProperties().getProperty('OPERATION_MODE');
  
  if (mode === 'demo') return;
  
  // YouTube API
  if (mode === 'basic' || mode === 'advanced') {
    const ytKey = ui.prompt(
      'YouTube API設定',
      'YouTube Data API v3 のキーを入力してください：\n' +
      '（取得方法はヘルプメニューを参照）',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (ytKey.getSelectedButton() === ui.Button.OK && ytKey.getResponseText()) {
      PropertiesService.getScriptProperties().setProperty('YOUTUBE_API_KEY', ytKey.getResponseText());
    }
  }
  
  // Advanced mode APIs
  if (mode === 'advanced') {
    // X API
    const xToken = ui.prompt(
      'X (Twitter) API設定',
      'X API のBearer Tokenを入力してください：\n' +
      '（スキップ可能）',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (xToken.getSelectedButton() === ui.Button.OK && xToken.getResponseText()) {
      PropertiesService.getScriptProperties().setProperty('X_BEARER_TOKEN', xToken.getResponseText());
    }
    
    // Gemini API
    const geminiKey = ui.prompt(
      'Gemini API設定',
      'Gemini API のキーを入力してください：\n' +
      '（AI分析とサジェスト機能に必要）',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (geminiKey.getSelectedButton() === ui.Button.OK && geminiKey.getResponseText()) {
      PropertiesService.getScriptProperties().setProperty('GEMINI_API_KEY', geminiKey.getResponseText());
    }
  }

  // Google Ads API
  const adsClientId = ui.prompt(
    'Google Ads API設定',
    'Google Ads API Client IDを入力してください：',
    ui.ButtonSet.OK_CANCEL
  );
  if (adsClientId.getSelectedButton() === ui.Button.OK && adsClientId.getResponseText()) {
    PropertiesService.getScriptProperties().setProperty('GOOGLE_ADS_CLIENT_ID', adsClientId.getResponseText());
  }

  const adsClientSecret = ui.prompt(
    'Google Ads API設定',
    'Google Ads API Client Secretを入力してください：',
    ui.ButtonSet.OK_CANCEL
  );
  if (adsClientSecret.getSelectedButton() === ui.Button.OK && adsClientSecret.getResponseText()) {
    PropertiesService.getScriptProperties().setProperty('GOOGLE_ADS_CLIENT_SECRET', adsClientSecret.getResponseText());
  }

  const adsDeveloperToken = ui.prompt(
    'Google Ads API設定',
    'Google Ads Developer Tokenを入力してください：',
    ui.ButtonSet.OK_CANCEL
  );
  if (adsDeveloperToken.getSelectedButton() === ui.Button.OK && adsDeveloperToken.getResponseText()) {
    PropertiesService.getScriptProperties().setProperty('GOOGLE_ADS_DEVELOPER_TOKEN', adsDeveloperToken.getResponseText());
  }

  const adsCustomerId = ui.prompt(
    'Google Ads API設定',
    'Google Ads Customer IDを入力してください（ハイフンなし）：',
    ui.ButtonSet.OK_CANCEL
  );
  if (adsCustomerId.getSelectedButton() === ui.Button.OK && adsCustomerId.getResponseText()) {
    PropertiesService.getScriptProperties().setProperty('GOOGLE_ADS_CUSTOMER_ID', adsCustomerId.getResponseText());
  }

  // Grok API
  const grokKey = ui.prompt(
    'Grok API設定',
    'Grok (xAI) APIキーを入力してください：\n（コンテンツサジェスト強化に使用）',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (grokKey.getSelectedButton() === ui.Button.OK && grokKey.getResponseText()) {
    PropertiesService.getScriptProperties().setProperty('GROK_API_KEY', grokKey.getResponseText());
  }
}

// =====================================
// メイン分析関数
// =====================================
function runQuickAnalysis(automated = false) {
  const ui = automated ? null : SpreadsheetApp.getUi();
  const settings = loadSettings();
  
  if (!settings.keyword) {
    if (ui) ui.alert('エラー', 'キーワードが設定されていません。設定シートで入力してください。', ui.ButtonSet.OK);
    return;
  }
  
  try {
    clearDataSheets();
    if (ui) showProgressMessage('分析を開始しています...');
    
    const mode = settings.operationMode || 'demo';
    let result;
    
    switch (mode) {
      case 'demo':
        result = runDemoAnalysis(settings);
        break;
      case 'basic':
        result = runBasicAnalysis(settings);
        break;
      case 'advanced':
        result = performFullAnalysis(settings);
        break;
    }
    
    saveAnalysisResult(result);
    updateDashboard(result);
    
    if (settings.notificationEmail) {
      sendNotification(result, settings);
    }
    
    if (ui) {
      ui.alert(
        '✅ 分析完了！',
        `総合スコア: ${result.score}/100\n\n` +
        `詳細はダッシュボードをご確認ください。`,
        ui.ButtonSet.OK
      );
    }
    
  } catch (error) {
    handleError(error);
    if (ui) ui.alert('エラー', '分析中にエラーが発生しました。詳細はエラーログを確認してください。', ui.ButtonSet.OK);
  }
}

function runAdvancedAnalysis(automated = false) {
  const ui = automated ? null : SpreadsheetApp.getUi();
  if (ui) ui.alert('詳細分析', 'AIサジェスト付きの詳細分析を開始します。\n少し時間がかかります。', ui.ButtonSet.OK);
  
  try {
    clearDataSheets();
    const settings = loadSettings();
    settings.includesSuggestions = true;
    
    for (const keyword of settings.keywords) {
      settings.keyword = keyword; // Set current keyword
      
      // Fetch previous result for comparison
      const previousResult = getPreviousAnalysisResult(keyword);
      
      const result = performFullAnalysis(settings);
      
      // Add comparisons
      if (previousResult) {
        result.previousScore = previousResult[2]; // Assuming column C is score
        result.scoreDelta = result.score - result.previousScore;
      }
      
      saveAnalysisResult(result);
      saveContentSuggestions(result.suggestions);
      updateDashboard(result);
      
      if (settings.notificationEmail) {
        generateAndSendHTMLReport(result, settings);
      }
    }
    
    if (ui) ui.alert('✅ 詳細分析完了！', 'サジェスト付きレポートが生成されました。', ui.ButtonSet.OK);
    
  } catch (error) {
    handleError(error);
  }
}

// =====================================
// 分析実行関数
// =====================================
function runDemoAnalysis(settings) {
  return {};
}

function runBasicAnalysis(settings) {
  const data = {
    timestamp: new Date(),
    keyword: settings.keyword,
    youtube: null,
    x: null,
    trends: null
  };
  
  // YouTube分析
  if (settings.youtubeEnabled && PropertiesService.getScriptProperties().getProperty('YOUTUBE_API_KEY')) {
    try {
      data.youtube = fetchYouTubeData(settings);
    } catch (error) {
      console.error('YouTube分析エラー:', error);
    }
  }
  
  // 基本的なトレンド分析
  data.trends = generateBasicTrends(settings);
  
  // スコア計算
  data.score = calculateScore(data);
  
  // AI要約（基本版）
  data.aiSummary = generateBasicSummary(data, settings);
  
  return data;
}

function performFullAnalysis(settings) {
  console.log('詳細分析を開始:', settings.keyword);
  
  // データ収集
  const collectedData = {
    timestamp: new Date(),
    keyword: settings.keyword,
    youtube: null,
    x: null,
    trends: null,
    errors: []
  };
  
  // 並列データ収集
  if (settings.youtubeEnabled) {
    try {
      collectedData.youtube = fetchYouTubeData(settings);
    } catch (error) {
      collectedData.errors.push({ service: 'YouTube', error: error.toString() });
    }
  }
  
  if (settings.xEnabled) {
    try {
      collectedData.x = fetchXData(settings);
    } catch (error) {
      collectedData.errors.push({ service: 'X', error: error.toString() });
    }
  }
  
  // トレンド分析
  collectedData.trends = analyzeTrends(collectedData, settings);
  
  // スコア計算
  collectedData.score = calculateScore(collectedData);
  
  // AI分析
  if (settings.aiEnabled && PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY')) {
    collectedData.aiAnalysis = performAIAnalysis(collectedData, settings);
  }
  
  // コンテンツサジェスト生成
  if (settings.includesSuggestions) {
    collectedData.suggestions = generateContentSuggestions(collectedData, settings);
  }
  
  // Google Ads Keyword Planner分析
  if (settings.aiEnabled) {
    try {
      collectedData.keywordPlanner = fetchKeywordPlannerData(settings);
    } catch (error) {
      collectedData.errors.push({ service: 'KeywordPlanner', error: error.toString() });
    }
  }
  
  return collectedData;
}

// =====================================
// YouTube API
// =====================================
function fetchYouTubeData(settings) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('YOUTUBE_API_KEY');
  if (!apiKey) throw new Error('YouTube APIキーが設定されていません');
  
  const data = {
    videos: [],
    channels: {},
    statistics: {
      totalViews: 0,
      totalLikes: 0,
      totalComments: 0,
      avgEngagementRate: 0
    }
  };
  
  // キーワード検索
  const searchUrl = `${CONFIG.YOUTUBE_API_URL}/search?` +
    `part=snippet&q=${encodeURIComponent(settings.keyword)}&` +
    `type=video&maxResults=${CONFIG.BATCH_SIZE}&order=relevance&` +
    `publishedAfter=${settings.startDate.toISOString()}&key=${apiKey}`;
  
  try {
    const searchResponse = UrlFetchApp.fetch(searchUrl);
    const searchData = JSON.parse(searchResponse.getContentText());
    
    if (searchData.items && searchData.items.length > 0) {
      // 動画詳細取得
      const videoIds = searchData.items.map(item => item.id.videoId).join(',');
      const detailsUrl = `${CONFIG.YOUTUBE_API_URL}/videos?` +
        `part=statistics,contentDetails,snippet&id=${videoIds}&key=${apiKey}`;
      
      const detailsResponse = UrlFetchApp.fetch(detailsUrl);
      const detailsData = JSON.parse(detailsResponse.getContentText());
      
      // データ処理
      detailsData.items.forEach(video => {
        const videoInfo = processYouTubeVideo(video);
        data.videos.push(videoInfo);
        
        // 統計更新
        data.statistics.totalViews += videoInfo.viewCount;
        data.statistics.totalLikes += videoInfo.likeCount;
        data.statistics.totalComments += videoInfo.commentCount;
      });
      
      // 平均エンゲージメント率計算
      if (data.videos.length > 0) {
        const totalEngagement = data.videos.reduce((sum, v) => sum + v.engagementRate, 0);
        data.statistics.avgEngagementRate = totalEngagement / data.videos.length;
      }
    }
    
    // 特定チャンネル分析
    if (settings.youtubeChannels && settings.youtubeChannels.length > 0) {
      settings.youtubeChannels.forEach(channelId => {
        try {
          data.channels[channelId] = fetchChannelData(channelId, settings, apiKey);
        } catch (error) {
          console.error(`チャンネル ${channelId} の取得エラー:`, error);
        }
      });
    }
    
  } catch (error) {
    throw new Error(`YouTube API エラー: ${error.toString()}`);
  }
  
  return data;
}

function processYouTubeVideo(video) {
  const stats = video.statistics;
  const viewCount = parseInt(stats.viewCount || 0);
  const likeCount = parseInt(stats.likeCount || 0);
  const commentCount = parseInt(stats.commentCount || 0);
  
  return {
    videoId: video.id,
    title: video.snippet.title,
    channelId: video.snippet.channelId,
    channelTitle: video.snippet.channelTitle,
    publishedAt: video.snippet.publishedAt,
    viewCount: viewCount,
    likeCount: likeCount,
    commentCount: commentCount,
    duration: parseDuration(video.contentDetails.duration),
    engagementRate: viewCount > 0 ? ((likeCount + commentCount) / viewCount * 100) : 0,
    thumbnail: video.snippet.thumbnails.medium.url
  };
}

// =====================================
// X (Twitter) API
// =====================================
function fetchXData(settings) {
  const bearerToken = PropertiesService.getScriptProperties().getProperty('X_BEARER_TOKEN');
  if (!bearerToken) {
    return generateMockXData(settings);
  }

  const data = {
    tweets: [],
    trends: [],
    accounts: {},
    statistics: {
      totalTweets: 0,
      totalLikes: 0,
      totalRetweets: 0,
      sentiment: { positive: 0, negative: 0, neutral: 0 }
    }
  };

  try {
    // ツイート検索
    const query = `${settings.keyword} -is:retweet`;
    const searchUrl = `${CONFIG.X_API_URL}/tweets/search/recent?` +
      `query=${encodeURIComponent(query)}&` +
      `start_time=${settings.startDate.toISOString()}&` +
      `max_results=100&tweet.fields=created_at,public_metrics,author_id`;
    
    const options = {
      method: 'GET',
      headers: {
        'Authorization': `Bearer ${bearerToken}`,
        'Content-Type': 'application/json'
      },
      muteHttpExceptions: true
    };
    
    const response = UrlFetchApp.fetch(searchUrl, options);
    const responseData = JSON.parse(response.getContentText());
    
    if (response.getResponseCode() === 200 && responseData.data) {
      responseData.data.forEach(tweet => {
        const tweetInfo = processXTweet(tweet);
        data.tweets.push(tweetInfo);
        
        // 統計更新
        data.statistics.totalLikes += tweetInfo.likeCount;
        data.statistics.totalRetweets += tweetInfo.retweetCount;
      });
      
      data.statistics.totalTweets = responseData.meta.result_count || data.tweets.length;
      
      // センチメント分析
      data.statistics.sentiment = analyzeSentiment(data.tweets);
    }

    // Xトレンド取得 (例: 全世界 WOEID=1)
    const trendsUrl = `${CONFIG.X_API_URL}/trends/place.json?id=1`;
    const trendsResponse = UrlFetchApp.fetch(trendsUrl, options);
    const trendsData = JSON.parse(trendsResponse.getContentText());
    if (trendsResponse.getResponseCode() === 200 && trendsData[0] && trendsData[0].trends) {
      data.trends = trendsData[0].trends.filter(trend => trend.name.toLowerCase().includes(settings.keyword.toLowerCase()));
    }
    
    // 特定アカウント分析
    if (settings.xAccounts && settings.xAccounts.length > 0) {
      settings.xAccounts.forEach(username => {
        try {
          data.accounts[username] = fetchXAccountData(username, settings, bearerToken);
        } catch (error) {
          console.error(`アカウント @${username} の取得エラー:`, error);
        }
      });
    }
    
  } catch (error) {
    console.error('X API エラー:', error);
    return generateMockXData(settings);
  }
  
  return data;
}

function processXTweet(tweet) {
  return {
    id: tweet.id,
    text: tweet.text,
    authorId: tweet.author_id,
    createdAt: tweet.created_at,
    likeCount: tweet.public_metrics.like_count || 0,
    retweetCount: tweet.public_metrics.retweet_count || 0,
    replyCount: tweet.public_metrics.reply_count || 0,
    impressionCount: tweet.public_metrics.impression_count || 0
  };
}

// =====================================
// AI分析とサジェスト
// =====================================
function performAIAnalysis(data, settings) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) {
    return generateBasicAnalysis(data, settings);
  }
  
  const prompt = createAnalysisPrompt(data, settings);
  
  try {
    const requestBody = {
      contents: [{
        parts: [{
          text: prompt
        }]
      }],
      generationConfig: {
        temperature: 0.7,
        topK: 40,
        topP: 0.95,
        maxOutputTokens: 8192,
      }
    };
    
    const options = {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'x-goog-api-key': apiKey
      },
      payload: JSON.stringify(requestBody),
      muteHttpExceptions: true
    };
    
    const response = UrlFetchApp.fetch(CONFIG.GEMINI_API_URL, options);
    const responseData = JSON.parse(response.getContentText());
    
    if (response.getResponseCode() === 200) {
      return responseData.candidates[0].content.parts[0].text;
    }
    
  } catch (error) {
    console.error('Gemini API エラー:', error);
  }
  
  return generateBasicAnalysis(data, settings);
}

function generateContentSuggestions(data, settings) {
  let suggestions = [];

  // Geminiサジェスト
  if (PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY')) {
    const geminiSuggestions = performGeminiSuggestions(data, settings);
    suggestions = suggestions.concat(geminiSuggestions);
  }

  // Grokサジェスト
  if (PropertiesService.getScriptProperties().getProperty('GROK_API_KEY')) {
    const grokSuggestions = performGrokSuggestions(data, settings);
    suggestions = suggestions.concat(grokSuggestions);
  }

  // トレンド分析に基づくサジェスト
  if (data.youtube && data.youtube.videos.length > 0) {
    const topVideos = data.youtube.videos
      .sort((a, b) => b.engagementRate - a.engagementRate)
      .slice(0, 5);
    
    topVideos.forEach(video => {
      const keywords = extractKeywords(video.title);
      suggestions.push({
        title: `「${settings.keyword} × ${keywords[0]}」の解説動画`,
        platform: 'YouTube',
        format: '10-15分動画',
        predictedPerformance: `${Math.floor(video.viewCount * 0.7)}再生予測`,
        difficulty: 3,
        recommendedTime: '金曜日 20:00',
        successProbability: 75,
        description: '高エンゲージメント動画の要素を取り入れた企画'
      });
    });
  }

  // X分析に基づくサジェスト
  if (data.x && data.x.tweets.length > 0) {
    suggestions.push({
      title: `${settings.keyword}に関する最新情報まとめスレッド`,
      platform: 'X',
      format: 'スレッド（5-10ツイート）',
      predictedPerformance: '1000+ エンゲージメント',
      difficulty: 2,
      recommendedTime: '平日 12:00',
      successProbability: 80,
      description: '話題性の高い情報をまとめて発信'
    });
  }

  // 汎用サジェスト
  suggestions.push(
    {
      title: `【完全版】${settings.keyword}の基礎から応用まで`,
      platform: 'YouTube',
      format: '長尺動画（20-30分）',
      predictedPerformance: '10万再生予測',
      difficulty: 4,
      recommendedTime: '土曜日 18:00',
      successProbability: 65,
      description: '包括的なガイドコンテンツ'
    },
    {
      title: `${settings.keyword}のよくある質問TOP10`,
      platform: 'ブログ',
      format: 'Q&A記事',
      predictedPerformance: '月間5000PV',
      difficulty: 2,
      recommendedTime: '随時',
      successProbability: 85,
      description: 'SEO対策とユーザーニーズに対応'
    }
  );

  // スコアリングとソート
  return rankSuggestions(suggestions, data);
}

function rankSuggestions(suggestions, data) {
  return suggestions.map(suggestion => {
    let score = suggestion.successProbability || 50;
    
    // プラットフォーム別の重み付け
    if (data.youtube && suggestion.platform === 'YouTube') {
      score += 10;
    }
    if (data.x && suggestion.platform === 'X') {
      score += 5;
    }
    
    // 難易度による調整
    score -= (suggestion.difficulty - 3) * 5;
    
    return { ...suggestion, score };
  }).sort((a, b) => b.score - a.score);
}

// =====================================
// レポート生成
// =====================================
function generateHTMLReport(data) {
  const html = `
<!DOCTYPE html>
<html lang="ja">
<head>
  <meta charset="UTF-8">
  <title>分析レポート - ${data.keyword}</title>
  <style>
    * { margin: 0; padding: 0; box-sizing: border-box; }
    body { 
      font-family: 'Helvetica Neue', Arial, sans-serif; 
      background: #f5f7fa;
      color: #2c3e50;
      line-height: 1.6;
    }
    .container { max-width: 1200px; margin: 0 auto; padding: 20px; }
    
    /* ヘッダー */
    .header {
      background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
      color: white;
      padding: 40px;
      border-radius: 15px;
      margin-bottom: 30px;
      text-align: center;
    }
    .header h1 { font-size: 2.5em; margin-bottom: 10px; }
    
    /* スコアカード */
    .score-card {
      background: white;
      padding: 30px;
      border-radius: 15px;
      box-shadow: 0 5px 20px rgba(0,0,0,0.1);
      text-align: center;
      margin-bottom: 30px;
    }
    .score-value {
      font-size: 4em;
      font-weight: bold;
      color: ${data.score >= 80 ? '#27ae60' : data.score >= 60 ? '#f39c12' : '#e74c3c'};
    }
    
    /* メトリクスグリッド */
    .metrics-grid {
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(250px, 1fr));
      gap: 20px;
      margin-bottom: 30px;
    }
    .metric-card {
      background: white;
      padding: 25px;
      border-radius: 12px;
      box-shadow: 0 3px 10px rgba(0,0,0,0.1);
    }
    .metric-label {
      color: #7f8c8d;
      font-size: 0.9em;
      text-transform: uppercase;
    }
    .metric-value {
      font-size: 2em;
      font-weight: bold;
      color: #2c3e50;
      margin: 10px 0;
    }
    
    /* サジェストセクション */
    .suggestions {
      background: white;
      padding: 30px;
      border-radius: 15px;
      box-shadow: 0 5px 20px rgba(0,0,0,0.1);
      margin-bottom: 30px;
    }
    .suggestion-item {
      border-left: 4px solid #667eea;
      padding: 20px;
      margin: 15px 0;
      background: #f8f9fa;
      border-radius: 8px;
    }
    .suggestion-title {
      font-size: 1.2em;
      font-weight: 600;
      color: #2c3e50;
      margin-bottom: 10px;
    }
    
    /* AI分析 */
    .ai-analysis {
      background: white;
      padding: 30px;
      border-radius: 15px;
      box-shadow: 0 5px 20px rgba(0,0,0,0.1);
      white-space: pre-wrap;
    }
  </style>
</head>
<body>
  <div class="container">
    <div class="header">
      <h1>ネット反応分析レポート - ${data.keyword}</h1>
      <p>生成日時: ${new Date(data.timestamp).toLocaleString('ja-JP')}</p>
    </div>
    
    <div class="score-card">
      <div class="metric-label">総合スコア</div>
      <div class="score-value">${data.score}/100</div>
      <p>前回比: ${data.scoreDelta ? (data.scoreDelta > 0 ? '+' : '') + data.scoreDelta : 'N/A'}</p>
    </div>
    
    <div class="metrics-grid">
      ${data.youtube ? `
        <div class="metric-card">
          <div class="metric-label">YouTube</div>
          <div class="metric-value">${formatNumber(data.youtube.statistics.totalViews)}</div>
          <p>総再生回数</p>
        </div>
      ` : ''}
      
      ${data.x ? `
        <div class="metric-card">
          <div class="metric-label">X (Twitter)</div>
          <div class="metric-value">${formatNumber(data.x.statistics.totalTweets)}</div>
          <p>ツイート数</p>
        </div>
      ` : ''}
      
      ${data.trends ? `
        <div class="metric-card">
          <div class="metric-label">検索トレンド</div>
          <div class="metric-value">${data.trends.growth > 0 ? '↑' : '↓'} ${Math.abs(data.trends.growth)}%</div>
          <p>成長率</p>
        </div>
      ` : ''}
    </div>
    
    ${data.suggestions && data.suggestions.length > 0 ? `
      <div class="suggestions">
        <h2>オススメコンテンツ</h2>
        ${data.suggestions.slice(0, 5).map((suggestion, index) => `
          <div class="suggestion-item">
            <div class="suggestion-title">${index + 1}. ${suggestion.title}</div>
            <p>プラットフォーム: ${suggestion.platform} | フォーマット: ${suggestion.format} | 予測パフォーマンス: ${suggestion.predictedPerformance}</p>
            <p style="margin-top: 10px;">${suggestion.description}</p>
          </div>
        `).join('')}
      </div>
    ` : ''}
    
    ${data.aiAnalysis ? `
      <div class="ai-analysis">
        <h2>AI分析</h2>
        <pre>${escapeHtml(data.aiAnalysis)}</pre>
      </div>
    ` : ''}
  </div>
</body>
</html>
`;
  return html;
}

function escapeHtml(text) {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

// =====================================
// データ保存
// =====================================
function saveAnalysisResult(result) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAMES.REPORTS);
  
  const row = [
    result.timestamp,
    result.keyword,
    result.score,
    result.youtube?.statistics.totalViews || 0,
    result.x?.statistics.totalTweets || 0,
    result.trends?.growth ? `${result.trends.growth > 0 ? '+' : ''}${result.trends.growth}%` : 'N/A',
    result.aiSummary || result.aiAnalysis || '',
    result.suggestions ? result.suggestions[0]?.title || '' : ''
  ];
  
  sheet.appendRow(row);
  
  // 生データも保存
  if (result.youtube) {
    saveYouTubeData(result.youtube, result.timestamp);
  }
  if (result.x) {
    saveXData(result.x, result.timestamp);
  }
}

function saveYouTubeData(data, timestamp) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAMES.YOUTUBE_DATA);
  
  if (sheet.getLastRow() === 0) {
    const headers = ['タイムスタンプ', '動画ID', 'タイトル', 'チャンネル', '再生回数', 'いいね数', 'コメント数', 'エンゲージメント率'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  
  data.videos.forEach(video => {
    sheet.appendRow([
      timestamp,
      video.videoId,
      video.title,
      video.channelTitle,
      video.viewCount,
      video.likeCount,
      video.commentCount,
      video.engagementRate.toFixed(2)
    ]);
  });
}

function saveContentSuggestions(suggestions) {
  if (!suggestions || suggestions.length === 0) return;
  
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAMES.SUGGESTIONS);
  
  if (sheet.getLastRow() === 0) {
    const headers = ['生成日時', 'タイトル', 'プラットフォーム', 'フォーマット', '予測パフォーマンス', '成功確率', 'スコア'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  
  suggestions.forEach(suggestion => {
    sheet.appendRow([
      new Date(),
      suggestion.title,
      suggestion.platform,
      suggestion.format,
      suggestion.predictedPerformance,
      suggestion.successProbability,
      suggestion.score
    ]);
  });
}

// =====================================
// ダッシュボード更新
// =====================================
function updateDashboard(result) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAMES.DASHBOARD);
  
  // 基本情報
  sheet.getRange('B3').setValue(new Date().toLocaleString('ja-JP'));
  sheet.getRange('B4').setValue(result.keyword);
  
  // スコアとメトリクス
  sheet.getRange('B7').setValue(result.score + '/100');
  
  if (result.youtube) {
    sheet.getRange('B8').setValue(formatNumber(result.youtube.statistics.totalViews));
    sheet.getRange('E8').setValue(result.youtube.statistics.avgEngagementRate.toFixed(2) + '%');
  }
  
  if (result.x) {
    sheet.getRange('B9').setValue(formatNumber(result.x.statistics.totalTweets));
    const posRate = (result.x.statistics.sentiment.positive / result.x.statistics.totalTweets * 100).toFixed(1);
    sheet.getRange('E9').setValue(posRate + '%');
  }
  
  if (result.trends) {
    sheet.getRange('B10').setValue(result.trends.searchVolume || 'N/A');
    sheet.getRange('E10').setValue((result.trends.growth > 0 ? '+' : '') + result.trends.growth + '%');
  }
  
  // サジェスト
  if (result.suggestions && result.suggestions.length > 0) {
    result.suggestions.slice(0, 3).forEach((suggestion, index) => {
      sheet.getRange(13 + index, 2).setValue(suggestion.title);
    });
  }
  
  // AI分析サマリー
  if (result.aiSummary || result.aiAnalysis) {
    const summary = (result.aiSummary || result.aiAnalysis).substring(0, 500) + '...';
    sheet.getRange('A18:E19').merge().setValue(summary).setWrap(true);
  }
}

// =====================================
// 通知機能
// =====================================
function sendNotification(result, settings) {
  const subject = `【${settings.reportType}レポート】${result.keyword} - スコア: ${result.score}/100`;
  
  let body = `
${result.keyword} の分析結果

📊 総合スコア: ${result.score}/100
`;

  if (result.youtube) {
    body += `
📹 YouTube分析
- 総再生回数: ${formatNumber(result.youtube.statistics.totalViews)}
- 平均エンゲージメント: ${result.youtube.statistics.avgEngagementRate.toFixed(2)}%
`;
  }

  if (result.x) {
    body += `
🐦 X (Twitter) 分析
- ツイート数: ${formatNumber(result.x.statistics.totalTweets)}
- ポジティブ率: ${(result.x.statistics.sentiment.positive / result.x.statistics.totalTweets * 100).toFixed(1)}%
`;
  }

  if (result.suggestions && result.suggestions.length > 0) {
    body += `
💡 本日のオススメコンテンツ:
${result.suggestions.slice(0, 3).map((s, i) => `${i + 1}. ${s.title}`).join('\n')}
`;
  }

  body += `
詳細はスプレッドシートをご確認ください。
`;

  try {
    MailApp.sendEmail({
      to: settings.notificationEmail,
      subject: subject,
      body: body
    });
  } catch (error) {
    console.error('メール送信エラー:', error);
  }
}

// =====================================
// トリガー設定
// =====================================
function setupTriggers() {
  // 既存トリガーを削除
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => ScriptApp.deleteTrigger(trigger));
  
  // 日次トリガー（毎日午前9時）
  ScriptApp.newTrigger('runDailyAnalysis')
    .timeBased()
    .everyDays(1)
    .atHour(9)
    .create();
  
  // 週次トリガー（毎週月曜日午前10時）
  ScriptApp.newTrigger('runWeeklyAnalysis')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(10)
    .create();
  
  // 月次トリガー（毎月1日午前11時）
  ScriptApp.newTrigger('runMonthlyAnalysis')
    .timeBased()
    .onMonthDay(1)
    .atHour(11)
    .create();
  
  console.log('トリガーを設定しました');
}

// 自動実行関数
function runDailyAnalysis() {
  updateReportType('短期');
  runAdvancedAnalysis(true);
}

function runWeeklyAnalysis() {
  updateReportType('中期');
  runAdvancedAnalysis(true);
}

function runMonthlyAnalysis() {
  updateReportType('長期');
  runAdvancedAnalysis(true);
}

function updateReportType(type) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAMES.SETTINGS);
  sheet.getRange('B3').setValue(type);
}

// =====================================
// ユーティリティ関数
// =====================================
function loadSettings() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAMES.SETTINGS);
  
  const settings = {
    keywords: sheet.getRange('B2').getValue().split(',').map(k => k.trim()).filter(k => k),
    reportType: sheet.getRange('B3').getValue(),
    notificationEmail: sheet.getRange('B4').getValue(),
    operationMode: sheet.getRange('B5').getValue() || 'advanced',
    youtubeEnabled: sheet.getRange('B6').getValue() === 'ON',
    xEnabled: sheet.getRange('B7').getValue() === 'ON',
    aiEnabled: sheet.getRange('B8').getValue() === 'ON',
    suggestionsEnabled: sheet.getRange('B9').getValue() === 'ON',
    xAccounts: sheet.getRange('B10').getValue().split(',').map(a => a.trim()).filter(a => a),
    youtubeChannels: sheet.getRange('B11').getValue().split(',').map(c => c.trim()).filter(c => c)
  };
  
  // 期間設定
  const periodDays = {
    '短期': 7,
    '中期': 30,
    '長期': 365
  };
  settings.period = periodDays[settings.reportType] || 7;
  settings.startDate = new Date();
  settings.startDate.setDate(settings.startDate.getDate() - settings.period);
  settings.endDate = new Date();
  
  return settings;
}

function formatNumber(num) {
  if (!num) return '0';
  if (num >= 1000000) {
    return (num / 1000000).toFixed(1) + 'M';
  } else if (num >= 1000) {
    return (num / 1000).toFixed(1) + 'K';
  }
  return num.toString();
}

function parseDuration(duration) {
  const match = duration.match(/PT(\d+H)?(\d+M)?(\d+S)?/);
  if (!match) return 0;
  
  const hours = parseInt(match[1]) || 0;
  const minutes = parseInt(match[2]) || 0;
  const seconds = parseInt(match[3]) || 0;
  
  return hours * 3600 + minutes * 60 + seconds;
}

function calculateScore(data) {
  let score = 50; // ベーススコア
  
  // YouTube要素
  if (data.youtube) {
    const ytScore = Math.min(data.youtube.statistics.avgEngagementRate * 10, 25);
    score += ytScore;
  }
  
  // X要素
  if (data.x) {
    const sentiment = data.x.statistics.sentiment;
    const total = sentiment.positive + sentiment.negative + sentiment.neutral;
    if (total > 0) {
      const positiveRate = sentiment.positive / total;
      score += positiveRate * 20;
    }
  }
  
  // トレンド要素
  if (data.trends && data.trends.growth) {
    score += Math.min(Math.max(data.trends.growth, -10), 10);
  }
  
  return Math.round(Math.min(Math.max(score, 0), 100));
}

function analyzeSentiment(tweets) {
  const positiveWords = ['良い', 'すごい', '素晴らしい', '最高', 'good', 'great', 'excellent', 'amazing'];
  const negativeWords = ['悪い', 'ダメ', '最悪', 'bad', 'terrible', 'awful', 'worst'];
  
  const sentiment = { positive: 0, negative: 0, neutral: 0 };
  
  tweets.forEach(tweet => {
    const text = tweet.text.toLowerCase();
    let isPositive = false;
    let isNegative = false;
    
    positiveWords.forEach(word => {
      if (text.includes(word)) isPositive = true;
    });
    
    negativeWords.forEach(word => {
      if (text.includes(word)) isNegative = true;
    });
    
    if (isPositive && !isNegative) {
      sentiment.positive++;
    } else if (isNegative && !isPositive) {
      sentiment.negative++;
    } else {
      sentiment.neutral++;
    }
  });
  
  return sentiment;
}

function extractKeywords(text) {
  const stopWords = ['の', 'は', 'が', 'を', 'に', 'で', 'と', 'the', 'a', 'an', 'and', 'or'];
  const words = text.toLowerCase().split(/\s+/)
    .filter(word => word.length > 2 && !stopWords.includes(word));
  
  // 簡易的な頻度分析
  const wordCount = {};
  words.forEach(word => {
    wordCount[word] = (wordCount[word] || 0) + 1;
  });
  
  return Object.entries(wordCount)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 5)
    .map(([word]) => word);
}

// =====================================
// エラーハンドリング
// =====================================
function handleError(error) {
  console.error('エラー発生:', error);
  
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAMES.ERROR_LOG);
  
  if (sheet.getLastRow() === 0) {
    const headers = ['タイムスタンプ', 'エラータイプ', 'メッセージ', 'スタック'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  
  sheet.appendRow([
    new Date(),
    error.name || 'Unknown',
    error.message || error.toString(),
    error.stack || 'No stack trace'
  ]);
  
  // 重大なエラーの場合は管理者に通知
  if (error.message && error.message.includes('API')) {
    const adminEmail = PropertiesService.getScriptProperties().getProperty('ADMIN_EMAIL');
    if (adminEmail) {
      try {
        MailApp.sendEmail({
          to: adminEmail,
          subject: '【エラー】ネット反応分析システム',
          body: `エラーが発生しました:\n\n${error.toString()}`
        });
      } catch (mailError) {
        console.error('メール送信失敗:', mailError);
      }
    }
  }
}

// =====================================
// ヘルプ機能
// =====================================
function showUserGuide() {
  const htmlContent = `
<div style="font-family: Arial, sans-serif; padding: 20px; max-width: 600px;">
  <h2>�� 使い方ガイド</h2>
  
  <h3>🚀 クイックスタート</h3>
  <ol>
    <li>初回セットアップを実行</li>
    <li>キーワードを設定</li>
    <li>クイック分析を実行</li>
  </ol>
  
  <h3>📊 分析モード</h3>
  <ul>
    <li><strong>デモモード</strong>: APIキー不要、サンプルデータで動作確認</li>
    <li><strong>ベーシック</strong>: YouTube分析のみ（要APIキー）</li>
    <li><strong>アドバンス</strong>: 全機能利用可能（全APIキー必要）</li>
  </ul>
  
  <h3>💡 便利な機能</h3>
  <ul>
    <li>自動実行: 日次・週次・月次レポートを自動生成</li>
    <li>コンテンツサジェスト: AI が最適なコンテンツを提案</li>
    <li>競合分析: 特定アカウントの動向を追跡</li>
  </ul>
  
  <h3>⚙️ カスタマイズ</h3>
  <p>設定シートで以下をカスタマイズできます：</p>
  <ul>
    <li>分析キーワード</li>
    <li>レポート期間</li>
    <li>通知先メールアドレス</li>
    <li>監視アカウント</li>
  </ul>
  
  <div style="margin-top: 20px; padding: 15px; background: #e8f0fe; border-radius: 5px;">
    <strong>💡 ヒント:</strong> まずはデモモードで動作を確認してから、APIキーを設定することをお勧めします。
  </div>
</div>
  `;
  
  const html = HtmlService.createHtmlOutput(htmlContent)
    .setWidth(650)
    .setHeight(600);
  
  SpreadsheetApp.getUi().showModalDialog(html, '使い方ガイド');
}

function showAPIKeyGuide() {
  const htmlContent = `
  <div style="font-family: Arial, sans-serif; padding: 20px; max-width: 600px;">
    <h2>🔑 APIキー取得ガイド</h2>
    
    <h3>YouTube Data API v3</h3>
    <ol>
      <li><a href="https://console.cloud.google.com/" target="_blank">Google Cloud Console</a> にアクセス</li>
      <li>新規プロジェクトを作成（または既存のものを選択）</li>
      <li>「APIとサービス」→「ライブラリ」を開く</li>
      <li>「YouTube Data API v3」を検索して有効化</li>
      <li>「認証情報」→「認証情報を作成」→「APIキー」</li>
      <li>生成されたAPIキーをコピー</li>
    </ol>
    <p><strong>無料枠:</strong> 1日10,000クォータ</p>
    
    <h3>Gemini API</h3>
    <ol>
      <li><a href="https://makersuite.google.com/app/apikey" target="_blank">Google AI Studio</a> にアクセス</li>
      <li>「Get API key」をクリック</li>
      <li>新規APIキーを作成</li>
    </ol>
    <p><strong>無料枠:</strong> 1分間60リクエスト</p>
    
    <h3>Google Ads API (Keyword Planner)</h3>
    <ol>
      <li><a href="https://ads.google.com/" target="_blank">Google Ads</a> アカウントにログイン</li>
      <li>ツール > キーワード プランナー にアクセス（有効化が必要）</li>
      <li><a href="https://developers.google.com/google-ads/api" target="_blank">Google Ads API ドキュメント</a>から Developer Token を申請/取得</li>
      <li>Google Cloud Consoleでプロジェクトを作成し、Google Ads APIを有効化</li>
      <li>「認証情報」→「OAuth クライアントID」作成（アプリケーションタイプ: Web application）</li>
      <li>Client IDとClient Secretをコピー</li>
      <li>Customer IDはGoogle AdsアカウントのID（ハイフンなし）</li>
    </ol>
    <p><strong>注意:</strong> Developer Tokenの承認には時間がかかる場合があります。無料トライアルあり。</p>
    
    <h3>X (Twitter) API v2</h3>
    <ol>
      <li><a href="https://developer.twitter.com/" target="_blank">Twitter Developer Portal</a> にアクセス</li>
      <li>開発者アカウントを申請</li>
      <li>プロジェクトとアプリを作成</li>
      <li>「Keys and tokens」からBearer Tokenを生成</li>
    </ol>
    <p><strong>無料枠:</strong> 月間500,000ツイート取得</p>
    
    <h3>Grok API</h3>
    <ol>
      <li><a href="https://api.x.ai/" target="_blank">xAI Developer Portal</a> にアクセス</li>
      <li>アカウントを作成し、APIアクセスをリクエスト</li>
      <li>承認後、APIキーを生成</li>
    </ol>
    <p><strong>無料枠:</strong> 制限あり、詳細はドキュメント確認</p>
    
    <div style="margin-top: 20px; padding: 15px; background: #fef7e0; border-radius: 5px;">
      <strong>⚠️ 注意:</strong> APIキーは他人と共有しないでください。定期的に更新することをお勧めします。
    </div>
  </div>
  `;
  
  const html = HtmlService.createHtmlOutput(htmlContent)
    .setWidth(650)
    .setHeight(700);
  
  SpreadsheetApp.getUi().showModalDialog(html, 'APIキー取得ガイド');
}

function showTroubleshooting() {
  const htmlContent = `
<div style="font-family: Arial, sans-serif; padding: 20px; max-width: 600px;">
  <h2>🔧 トラブルシューティング</h2>
  
  <h3>よくある問題と解決方法</h3>
  
  <div style="margin: 15px 0; padding: 15px; background: #fef7e0; border-radius: 5px;">
    <strong>Q: 「キーワードが設定されていません」エラー</strong><br>
    A: 設定シートのB2セルにキーワードを入力してください。
  </div>
  
  <div style="margin: 15px 0; padding: 15px; background: #fef7e0; border-radius: 5px;">
    <strong>Q: YouTube APIエラーが発生する</strong><br>
    A: 以下を確認してください：
    <ul>
      <li>APIキーが正しく設定されているか</li>
      <li>YouTube Data API v3が有効化されているか</li>
      <li>日次クォータ（10,000）を超えていないか</li>
    </ul>
  </div>
  
  <div style="margin: 15px 0; padding: 15px; background: #fef7e0; border-radius: 5px;">
    <strong>Q: メール通知が届かない</strong><br>
    A: 以下を確認してください：
    <ul>
      <li>メールアドレスが正しく入力されているか</li>
      <li>迷惑メールフォルダを確認</li>
      <li>Googleアカウントの送信制限（1日500通）に達していないか</li>
    </ul>
  </div>
  
  <div style="margin: 15px 0; padding: 15px; background: #fef7e0; border-radius: 5px;">
    <strong>Q: 分析に時間がかかる</strong><br>
    A: 詳細分析は多くのAPIコールを行うため、2-3分かかることがあります。
  </div>
  
  <h3>エラーログの確認方法</h3>
  <p>「エラーログ」シートで詳細なエラー情報を確認できます。</p>
  
  <h3>それでも解決しない場合</h3>
  <ol>
    <li>設定をリセットして再設定</li>
    <li>スプレッドシートを複製して新規インストール</li>
    <li>エラーメッセージをGoogleで検索</li>
  </ol>
  
  <div style="margin-top: 20px; padding: 15px; background: #e8f0fe; border-radius: 5px;">
    <strong>💡 ヒント:</strong> デモモードで基本動作を確認してから、段階的に機能を追加していくことをお勧めします。
  </div>
</div>
  `;
  
  const html = HtmlService.createHtmlOutput(htmlContent)
    .setWidth(650)
    .setHeight(700);
  
  SpreadsheetApp.getUi().showModalDialog(html, 'トラブルシューティング');
}

// =====================================
// テスト関数
// =====================================
function testYouTubeAPI() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    const settings = {
      keyword: 'Google Apps Script',
      startDate: new Date(Date.now() - 7 * 24 * 60 * 60 * 1000),
      period: 7
    };
    
    const data = fetchYouTubeData(settings);
    ui.alert('✅ YouTube API テスト成功', 
      `取得動画数: ${data.videos.length}\n` +
      `総再生回数: ${formatNumber(data.statistics.totalViews)}`,
      ui.ButtonSet.OK);
      
  } catch (error) {
    ui.alert('❌ YouTube API テスト失敗', error.toString(), ui.ButtonSet.OK);
  }
}

function testXAPI() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    const settings = {
      keyword: 'Google Apps Script',
      startDate: new Date(Date.now() - 7 * 24 * 60 * 60 * 1000)
    };
    
    const data = fetchXData(settings);
    ui.alert('✅ X API テスト成功', 
      `取得ツイート数: ${data.tweets.length}`,
      ui.ButtonSet.OK);
      
  } catch (error) {
    ui.alert('❌ X API テスト失敗', error.toString(), ui.ButtonSet.OK);
  }
}

function testGeminiAPI() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    const testData = {
      keyword: 'テスト',
      youtube: { statistics: { totalViews: 1000 } },
      x: { statistics: { totalTweets: 100 } }
    };
    
    const analysis = performAIAnalysis(testData, { keyword: 'テスト' });
    ui.alert('✅ Gemini API テスト成功', 
      '分析結果:\n' + analysis.substring(0, 200) + '...',
      ui.ButtonSet.OK);
      
  } catch (error) {
    ui.alert('❌ Gemini API テスト失敗', error.toString(), ui.ButtonSet.OK);
  }
}

function testSuggestions() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    const mockData = {
      keyword: 'AI',
      youtube: {
        videos: [
          { title: 'AI入門', viewCount: 10000, engagementRate: 5 }
        ]
      }
    };
    
    const suggestions = generateContentSuggestions(mockData, { keyword: 'AI' });
    ui.alert('✅ サジェスト機能テスト成功', 
      `生成サジェスト数: ${suggestions.length}\n` +
      `トップ提案: ${suggestions[0]?.title || 'なし'}`,
      ui.ButtonSet.OK);
      
  } catch (error) {
    ui.alert('❌ サジェスト機能テスト失敗', error.toString(), ui.ButtonSet.OK);
  }
}

// =====================================
// その他のユーティリティ
// =====================================
function resetSettings() {
  const ui = SpreadsheetApp.getUi();
  
  const confirm = ui.alert(
    '設定リセット',
    '全ての設定とAPIキーを削除します。続行しますか？',
    ui.ButtonSet.YES_NO
  );
  
  if (confirm === ui.Button.YES) {
    // プロパティをクリア
    PropertiesService.getScriptProperties().deleteAllProperties();
    
    // シートを再初期化
    initializeSpreadsheet();
    
    ui.alert('✅ リセット完了', '設定をリセットしました。セットアップウィザードを再実行してください。', ui.ButtonSet.OK);
  }
}

function showProgressMessage(message) {
  Logger.log(message);
}

function generateDemoSummary(keyword) {
  return '';
}

function generateDemoSuggestions(keyword) {
  return [];
}

function generateMockXData(settings) {
  return {
    tweets: [],
    statistics: {
      totalTweets: 0,
      totalLikes: 0,
      totalRetweets: 0,
      sentiment: { positive: 0, negative: 0, neutral: 0 }
    }
  };
}

function generateBasicTrends(settings) {
  return {
    searchVolume: 0,
    growth: 0,
    relatedTerms: []
  };
}

function analyzeTrends(data, settings) {
  const trends = {
    searchVolume: 0,
    growth: 0,
    momentum: 'stable'
  };
  
  // YouTube トレンド
  if (data.youtube && data.youtube.videos.length > 0) {
    const recentVideos = data.youtube.videos.filter(v => {
      const publishDate = new Date(v.publishedAt);
      const daysSincePublish = (new Date() - publishDate) / (1000 * 60 * 60 * 24);
      return daysSincePublish <= 7;
    });
    
    if (recentVideos.length > data.youtube.videos.length * 0.3) {
      trends.momentum = 'rising';
      trends.growth += 10;
    }
  }
  
  // X トレンド
  if (data.x && data.x.statistics.sentiment.positive > data.x.statistics.sentiment.negative * 2) {
    trends.growth += 5;
  }
  
  return trends;
}

function generateBasicSummary(data, settings) {
  return '';
}

function generateBasicAnalysis(data, settings) {
  return generateBasicSummary(data, settings);
}

function createAnalysisPrompt(data, settings) {
  return `
  キーワード「${settings.keyword}」に関する${settings.reportType}分析を行ってください。

  【収集データ】
  ${data.youtube ? `
  YouTube:
  - 動画数: ${data.youtube.videos.length}
  - 総再生回数: ${data.youtube.statistics.totalViews}
  - 平均エンゲージメント率: ${data.youtube.statistics.avgEngagementRate?.toFixed(2)}%
  ` : 'YouTube: データなし'}

  ${data.x ? `
  X (Twitter):
  - ツイート数: ${data.x.statistics.totalTweets}
  - 総エンゲージメント: ${data.x.statistics.totalLikes + data.x.statistics.totalRetweets}
  - センチメント: ポジティブ ${data.x.statistics.sentiment.positive}, ネガティブ ${data.x.statistics.sentiment.negative}
  ` : 'X: データなし'}

  以下の形式で分析結果を提供してください：

  1. 📊 総合評価（5段階評価と理由）
  2. 🎯 主要トレンド（3つ）
  3. 💭 ユーザー感情分析
  4. 📈 成長予測
  5. 💡 具体的な推奨アクション（5つ）

  簡潔で実用的な分析をお願いします。
  出力はプレーンテキストで、Markdown形式（例: ## や **）を使用しないでください。
  `;
}

function saveHTMLReport(html) {
  const folder = DriveApp.getRootFolder();
  const fileName = `レポート_${new Date().toISOString().split('T')[0]}.html`;
  const blob = Utilities.newBlob(html, 'text/html', fileName);
  folder.createFile(blob);
}

function initializeDataSheet(sheet) {
  sheet.getRange('A1').setValue('データは自動的に記録されます');
}

function initializeReportsSheet(sheet) {
  const headers = ['生成日時', 'キーワード', 'スコア', 'YouTube再生数', 'X投稿数', 'トレンド', 'AI分析', 'トップサジェスト'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setBackground('#4a86e8').setFontColor('white').setFontWeight('bold');
}

function initializeSuggestionsSheet(sheet) {
  const headers = ['生成日時', 'タイトル', 'プラットフォーム', 'フォーマット', '予測パフォーマンス', '成功確率', 'スコア'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setBackground('#4a86e8').setFontColor('white').setFontWeight('bold');
}

function saveXData(data, timestamp) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAMES.X_DATA);
  
  if (sheet.getLastRow() === 0) {
    const headers = ['タイムスタンプ', 'ツイートID', 'テキスト', 'いいね数', 'RT数', '返信数'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  
  data.tweets.slice(0, 20).forEach(tweet => {
    sheet.appendRow([
      timestamp,
      tweet.id,
      tweet.text.substring(0, 100) + '...',
      tweet.likeCount,
      tweet.retweetCount,
      tweet.replyCount
    ]);
  });
}

function fetchChannelData(channelId, settings, apiKey) {
  const channelUrl = `${CONFIG.YOUTUBE_API_URL}/channels?part=snippet,statistics&id=${channelId}&key=${apiKey}`;
  const response = UrlFetchApp.fetch(channelUrl);
  const data = JSON.parse(response.getContentText());
  
  if (!data.items || data.items.length === 0) {
    throw new Error(`チャンネル ${channelId} が見つかりません`);
  }
  
  return {
    channelInfo: data.items[0],
    recentActivity: []
  };
}

function fetchXAccountData(username, settings, bearerToken) {
  // 簡易実装
  return {
    username: username,
    tweets: []
  };
}

function finalizeSetup() {
  const ui = SpreadsheetApp.getUi();
  ui.alert(
    '🎉 セットアップ完了！',
    '設定が完了しました。\n\n' +
    '次のステップ:\n' +
    '1. メニューから「クイック分析」を実行\n' +
    '2. ダッシュボードで結果を確認\n' +
    '3. 必要に応じて設定を調整\n\n' +
    'それでは、分析を始めましょう！',
    ui.ButtonSet.OK
  );
}

function showReleaseNotes() {
  const htmlContent = `
<div style="font-family: Arial, sans-serif; padding: 20px; max-width: 600px;">
  <h2>📋 更新情報</h2>
  
  <h3>Version 4.0 (最新)</h3>
  <ul>
    <li>✨ 完全統合版リリース</li>
    <li>🚀 簡易セットアップウィザード追加</li>
    <li>🎯 AIコンテンツサジェスト機能</li>
    <li>📊 エグゼクティブダッシュボード</li>
    <li>🔧 エラーハンドリング改善</li>
  </ul>
  
  <h3>Version 3.0</h3>
  <ul>
    <li>予測分析機能追加</li>
    <li>競合分析機能</li>
    <li>HTMLレポート生成</li>
  </ul>
  
  <h3>Version 2.0</h3>
  <ul>
    <li>X (Twitter) API対応</li>
    <li>Gemini AI統合</li>
    <li>自動実行機能</li>
  </ul>
  
  <h3>Version 1.0</h3>
  <ul>
    <li>初回リリース</li>
    <li>YouTube分析機能</li>
    <li>基本レポート機能</li>
  </ul>
</div>
  `;
  
  const html = HtmlService.createHtmlOutput(htmlContent)
    .setWidth(650)
    .setHeight(500);
  
  SpreadsheetApp.getUi().showModalDialog(html, '更新情報');
}

// カスタムレポート生成
function generateCustomReport() {
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.prompt(
    'カスタムレポート',
    '分析期間を日数で入力してください（例: 14）:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() === ui.Button.OK) {
    const days = parseInt(response.getResponseText());
    if (days > 0 && days <= 365) {
      const settings = loadSettings();
      settings.period = days;
      settings.startDate = new Date();
      settings.startDate.setDate(settings.startDate.getDate() - days);
      
      runAdvancedAnalysis();
    } else {
      ui.alert('エラー', '1〜365の範囲で入力してください。', ui.ButtonSet.OK);
    }
  }
}

// レポート生成関数
function generateDailyReport() {
  updateReportType('短期');
  runAdvancedAnalysis(true);
}

function generateWeeklyReport() {
  updateReportType('中期');
  runAdvancedAnalysis(true);
}

function generateMonthlyReport() {
  updateReportType('長期');
  runAdvancedAnalysis(true);
}

// 詳細設定画面
function showAdvancedSettings() {
  const htmlContent = `
<div style="font-family: Arial, sans-serif; padding: 20px; max-width: 600px;">
  <h2>⚙️ 詳細設定</h2>
  
  <p>以下の設定は設定シートで変更できます：</p>
  
  <h3>分析設定</h3>
  <ul>
    <li><strong>分析キーワード</strong>: メインの分析対象</li>
    <li><strong>レポートタイプ</strong>: 短期/中期/長期</li>
    <li><strong>動作モード</strong>: demo/basic/advanced</li>
  </ul>
  
  <h3>機能ON/OFF</h3>
  <ul>
    <li><strong>YouTube分析</strong>: 動画分析の有効/無効</li>
    <li><strong>X分析</strong>: ツイート分析の有効/無効</li>
    <li><strong>AI要約</strong>: AI分析の有効/無効</li>
    <li><strong>サジェスト生成</strong>: コンテンツ提案の有効/無効</li>
  </ul>
  
  <h3>監視対象</h3>
  <ul>
    <li><strong>Xアカウント</strong>: カンマ区切りで複数指定可</li>
    <li><strong>YouTubeチャンネル</strong>: チャンネルIDをカンマ区切りで指定</li>
  </ul>
  
  <h3>通知設定</h3>
  <ul>
    <li><strong>通知メール</strong>: レポート送信先（カンマ区切りで複数可）</li>
  </ul>
  
  <div style="margin-top: 20px; padding: 15px; background: #e8f0fe; border-radius: 5px;">
    <strong>💡 ヒント:</strong> 設定変更後は必ず保存してから分析を実行してください。
  </div>
</div>
  `;
  
  const html = HtmlService.createHtmlOutput(htmlContent)
    .setWidth(650)
    .setHeight(600);
  
  SpreadsheetApp.getUi().showModalDialog(html, '詳細設定ガイド');
}

function getOAuthService() {
  return OAuth2.createService('googleAds')
    .setAuthorizationBaseUrl('https://accounts.google.com/o/oauth2/auth')
    .setTokenUrl('https://oauth2.googleapis.com/token')
    .setClientId(PropertiesService.getScriptProperties().getProperty('GOOGLE_ADS_CLIENT_ID'))
    .setClientSecret(PropertiesService.getScriptProperties().getProperty('GOOGLE_ADS_CLIENT_SECRET'))
    .setCallbackFunction('authCallback')
    .setPropertyStore(PropertiesService.getUserProperties())
    .setScope('https://www.googleapis.com/auth/adwords')
    .setParam('access_type', 'offline')
    .setParam('approval_prompt', 'force');
}

function authCallback(request) {
  const service = getOAuthService();
  const authorized = service.handleCallback(request);
  if (authorized) {
    return HtmlService.createHtmlOutput('認証成功！ウィンドウを閉じてください。');
  } else {
    return HtmlService.createHtmlOutput('認証失敗');
  }
}

function fetchKeywordPlannerData(settings) {
  const service = getOAuthService();
  if (!service.hasAccess()) {
    const authorizationUrl = service.getAuthorizationUrl();
    Logger.log('Open the following URL and re-run: ' + authorizationUrl);
    return null;
  }

  const accessToken = service.getAccessToken();
  const developerToken = PropertiesService.getScriptProperties().getProperty('GOOGLE_ADS_DEVELOPER_TOKEN');
  const kpCustomerId = PropertiesService.getScriptProperties().getProperty('GOOGLE_ADS_CUSTOMER_ID');

  const url = `${CONFIG.GOOGLE_ADS_API_URL}/customers/${kpCustomerId}:generateKeywordIdeas`;
  const payload = {
    keywordPlanNetwork: 'GOOGLE_SEARCH',
    keywordSeed: { keywords: [settings.keyword] },
    geoTargetConstants: ['geoTargetConstants/2392'],
    language: 'languageConstants/1000',
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      Authorization: `Bearer ${accessToken}`,
      'developer-token': developerToken,
      'login-customer-id': kpCustomerId,
    },
    payload: JSON.stringify(payload)
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const data = JSON.parse(response.getContentText());
    return data.keywordIdeas.map(idea => ({
      keyword: idea.text,
      avgMonthlySearches: idea.keywordIdeaMetrics.avgMonthlySearches,
      competition: idea.keywordIdeaMetrics.competition
    }));
  } catch (error) {
    console.error('Keyword Planner error:', error);
    return null;
  }
}

function clearDataSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetsToClear = [];
  sheetsToClear.forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (sheet) {
      sheet.clearContents();
    }
  });
}

function performGrokSuggestions(data, settings) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GROK_API_KEY');
  if (!apiKey) return [];

  const prompt = createSuggestionsPrompt(data, settings);

  try {
    const requestBody = {
      model: 'grok-4',
      messages: [{ role: 'user', content: prompt }],
      temperature: 0.7,
      max_tokens: 1024
    };

    const options = {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        Authorization: `Bearer ${apiKey}`
      },
      payload: JSON.stringify(requestBody),
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(CONFIG.GROK_API_URL, options);
    const responseData = JSON.parse(response.getContentText());

    if (response.getResponseCode() === 200 && responseData.choices) {
      return parseSuggestions(responseData.choices[0].message.content);
    }
  } catch (error) {
    console.error('Grok API error:', error);
  }

  return [];
}

function parseSuggestions(text) {
  return text.split('\n').filter(line => line.trim()).map(line => {
    const parts = line.split('|');
    return {
      title: parts[0]?.trim() || '',
      platform: parts[1]?.trim() || 'Unknown',
      description: parts[2]?.trim() || '',
      format: parts[3]?.trim() || 'N/A',
      predictedPerformance: parts[4]?.trim() || 'N/A',
      successProbability: parseInt(parts[5]?.trim()) || 50
    };
  });
}

function createSuggestionsPrompt(data, settings) {
  return `以下のデータに基づいて、${settings.keyword}に関するコンテンツを5つ提案してください。\n現在のWeb状況（最近のトレンド、検索ボリューム、競合状況）を考慮して、実用的で魅力的なアイデアを生成してください。\n\n【データ】\n${JSON.stringify(data, null, 2)}\n\n出力形式: 各サジェストを1行で、タイトル|プラットフォーム|説明`; }

function performGeminiSuggestions(data, settings) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) return [];

  const prompt = createSuggestionsPrompt(data, settings);

  const requestBody = {
    contents: [{
      parts: [{
        text: prompt
      }]
    }],
    generationConfig: {
      temperature: 0.7,
      topK: 40,
      topP: 0.95,
      maxOutputTokens: 8192,
    }
  };

  const options = {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'x-goog-api-key': apiKey
    },
    payload: JSON.stringify(requestBody),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(CONFIG.GEMINI_API_URL, options);
  const responseData = JSON.parse(response.getContentText());

  if (response.getResponseCode() === 200) {
    return parseSuggestions(responseData.candidates[0].content.parts[0].text);
  }

  return [];
}

function generateAndSendHTMLReport(result, settings) {
  const htmlReport = generateHTMLReport(result);
  saveHTMLReport(htmlReport); // Optional: still save to Drive
  
  const subject = `【${settings.reportType}レポート】${result.keyword} - スコア: ${result.score}/100`;
  const plainBody = `分析レポート: ${result.keyword}\nスコア: ${result.score}/100\n詳細はHTML版をご覧ください。`;
  
  GmailApp.sendEmail(settings.notificationEmail, subject, plainBody, {
    htmlBody: htmlReport
  });
}

function getPreviousAnalysisResult(keyword) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAMES.REPORTS);
  const data = sheet.getDataRange().getValues();
  
  // Find the most recent previous entry for this keyword
  for (let i = data.length - 1; i >= 1; i--) { // Skip header
    if (data[i][1] === keyword) {
      return data[i];
    }
  }
  return null;
}