// ===== Google Apps Script - ライセンス管理システム =====
// 
// このシステムは、複数のプラットフォームのライセンスを統合管理し、
// 定期的な棚卸しと差分分析を行います。
// 
// 必要な権限:
// - Google Admin SDK API
// - Google Drive API
// - Gmail API
// - 外部APIアクセス（UrlFetchApp）

// ===== Google Apps Script - ライセンス管理システム（修正版） =====

// ===== 設定値 =====
const CONFIG = {
  // スプレッドシートの設定
  SHEET_NAMES: {
    DASHBOARD: 'ダッシュボード',
    LICENSES: 'ライセンス一覧',
    USERS: 'ユーザー管理',
    DIFF_ANALYSIS: '差分分析',
    HISTORY: '変更履歴',
    SETTINGS: '設定',
    INVENTORY_HOLD: '棚卸し保留'
  },
  
  // API設定
  GOOGLE_WORKSPACE: {
    PRODUCT_ID: 'Google-Apps',
    SKU_IDS: {
      'Business Standard': '1010020020',
      'Business Plus': '1010020021',
      'Enterprise Standard': '1010020028',
      'Enterprise Plus': '1010020027'
    },
    // 月額料金（円）
    PRICING: {
      'Business Starter': 680,
      'Business Standard': 1360,
      'Business Plus': 2040,
      'Enterprise Standard': 2700,
      'Enterprise Plus': 4100,
      'Enterprise Essentials': 1080
    }
  },
  
  // Zoom料金設定（月額・円）
  ZOOM: {
    PRICING: {
      'Basic': 0,
      'Pro': 2000,
      'Business': 2700,
      'Enterprise': 3200,
      'Enterprise Plus': 3600
    }
  },
  
  // Notion料金設定（月額・円）
  NOTION: {
    PRICING: {
      'Free': 0,
      'Plus': 1500,              // 年額プラン $10/月
      'Plus_Monthly': 1800,      // 月額プラン $12/月
      'Business': 2250,          // 年額プラン $15/月（推定）
      'Business_Monthly': 2700,  // 月額プラン $18/月（推定）
      'Enterprise': 3750,        // お問い合わせベース $25/月（推定）
      'Enterprise_Monthly': 3750 // Enterpriseは通常年契約のみ
    }
  },
  
  // プラットフォーム一覧
  PLATFORMS: [
    'Google Workspace',
    'トラスト・ログイン',
    'ActiveGate',
    'Microsoft 365',
    'Canva',
    'ESET',
    'Slack',
    'FortiClient',
    'ファイルサーバー',
    'v-Client',
    'Zoom',
    'Notion',
    'QUDEN',
    'ChatGPT',
    'サクラサーバー'
  ],
  
  // アラート設定
  UNUSED_LICENSE_ALERT_DAYS: 30
};

// ===== 権限の初期化（初回実行時に使用） =====
function initializePermissions() {
  try {
    // Gmail権限を要求
    GmailApp.getInboxThreads(0, 1);
    
    // スプレッドシート権限を要求
    SpreadsheetApp.getActiveSpreadsheet();
    
    // Drive権限を要求（必要に応じて）
    DriveApp.getRootFolder();
    
    SpreadsheetApp.getUi().alert(
      '権限設定完了',
      '必要な権限が正常に設定されました。\n' +
      'メイン処理を実行できます。',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (error) {
    SpreadsheetApp.getUi().alert(
      'エラー',
      '権限の設定中にエラーが発生しました。\n' +
      error.toString(),
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

// ===== メイン処理（修正版） =====
function main() {
  try {
    // スプレッドシートオブジェクトを正しく取得
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    
    if (!spreadsheet) {
      throw new Error('スプレッドシートが見つかりません');
    }
    
    console.log('処理開始: ' + spreadsheet.getName());
    
    // シートの初期化
    initializeSheets(spreadsheet);
    
    // 各プラットフォームからライセンス情報を取得
    const allLicenses = fetchAllLicenses();
    
    // データをスプレッドシートに書き込み
    updateLicenseSheet(spreadsheet, allLicenses);
    
    // ユーザー管理シートを更新
    updateUserManagementSheet(spreadsheet, allLicenses);
    
    // ダッシュボードを更新
    updateDashboard(spreadsheet, allLicenses);
    
    // 差分分析を実行
    console.log('差分分析を実行します。spreadsheet:', spreadsheet ? 'あり' : 'なし', ', licenses数:', allLicenses.length);
    const analysisResults = performDifferenceAnalysis(spreadsheet, allLicenses);
    
    // Zoom利用分析を実行（エラーが発生してもスクリプトは継続）
    try {
      console.log('Zoom利用分析を実行します');
      setupZoomAnalytics();
      fetchZoomMeetings();
      
      // Zoom利用状況下位レポートを生成
      createZoomLowUsageReport();
    } catch (zoomError) {
      console.error('Zoom分析エラー:', zoomError);
      // エラーがあっても処理を継続
    }
    
    // Notion利用分析を実行（エラーが発生してもスクリプトは継続）
    try {
      console.log('Notion利用分析を実行します');
      setupNotionAnalytics();
      
      // データベースを自動検出
      const databases = discoverNotionDatabases();
      
      if (databases.length > 0) {
        // データベース設定をシートに保存
        const configSheet = spreadsheet.getSheetByName(NOTION_ANALYSIS_SHEET_NAMES.DATABASE_CONFIG);
        const pageListSheet = spreadsheet.getSheetByName(NOTION_ANALYSIS_SHEET_NAMES.PAGE_LIST);
        const userActivitySheet = spreadsheet.getSheetByName(NOTION_ANALYSIS_SHEET_NAMES.USER_ACTIVITY);
        
        if (configSheet) {
          // 既存データをクリア
          const lastRow = configSheet.getLastRow();
          if (lastRow > 1) {
            configSheet.getRange(2, 1, lastRow - 1, 6).clearContent();
          }
          
          // データベース情報を書き込み
          const configData = databases.map(db => [
            db.id,
            db.name,
            '有効',
            '',
            db.created || '',
            db.url || ''
          ]);
          
          if (configData.length > 0) {
            configSheet.getRange(2, 1, configData.length, 6).setValues(configData);
          }
        }
        
        // ページリストをクリア
        if (pageListSheet) {
          const pageLastRow = pageListSheet.getLastRow();
          if (pageLastRow > 1) {
            pageListSheet.getRange(2, 1, pageLastRow - 1, 5).clearContent();
          }
        }
        
        // ユーザーアクティビティシートをクリア
        if (userActivitySheet) {
          const activityLastRow = userActivitySheet.getLastRow();
          if (activityLastRow > 1) {
            userActivitySheet.getRange(2, 1, activityLastRow - 1, 6).clearContent();
          }
        }
        
        // 全ページを取得してユーザー情報を収集
        const allPageData = [];
        const allActivityData = [];
        
        for (const db of databases) {
          try {
            const pages = fetchNotionDatabasePages(db.id);
            console.log(`データベース ${db.name}: ${pages.length}ページを取得`);
            
            for (const page of pages) {
              const pageId = page.id;
              const title = extractNotionPageTitle(page);
              const lastEditedTime = page.last_edited_time || '';
              
              // 編集者情報を取得
              let editorName = '';
              
              // 最終編集者を取得
              if (page.last_edited_by && page.last_edited_by.id) {
                try {
                  const userInfo = fetchNotionUser(page.last_edited_by.id);
                  if (userInfo) {
                    editorName = userInfo.name || userInfo.type || 'Unknown User';
                  }
                } catch (e) {
                  // エラーログは出力しない（404エラーは正常な動作）
                }
              }
              
              // 編集者が取得できない場合は作成者を試す
              if (!editorName && page.created_by && page.created_by.id) {
                try {
                  const userInfo = fetchNotionUser(page.created_by.id);
                  if (userInfo) {
                    editorName = userInfo.name || userInfo.type || 'Unknown User';
                  }
                } catch (e) {
                  // エラーログは出力しない（404エラーは正常な動作）
                }
              }
              
              if (!editorName) {
                editorName = 'Unknown User';
              }
              
              allPageData.push([db.name, pageId, title, lastEditedTime, editorName]);
              
              // ユーザーアクティビティデータを準備
              if (editorName !== 'Unknown User' && lastEditedTime) {
                allActivityData.push([
                  new Date(lastEditedTime), // 実際の編集日時を使用
                  db.name,
                  editorName,
                  title,
                  'Page Edit',
                  pageId
                ]);
              }
              
              // Notion_Logシートにも記録
              const logSheet = spreadsheet.getSheetByName(NOTION_ANALYSIS_SHEET_NAMES.NOTION_LOG);
              if (logSheet) {
                const logData = [
                  new Date(), // 記録日時
                  db.name, // データベース名
                  pageId, // Page ID
                  title, // Page Title
                  lastEditedTime ? new Date(lastEditedTime) : '', // 最終編集日時
                  editorName, // 最終編集者
                  'New/Update' // 変更検出
                ];
                logSheet.appendRow(logData);
              }
            }
          } catch (error) {
            console.error(`データベース ${db.name} でエラー:`, error);
          }
        }
        
        // ページデータを書き込み
        if (pageListSheet && allPageData.length > 0) {
          pageListSheet.getRange(2, 1, allPageData.length, 5).setValues(allPageData);
        }
        
        // アクティビティデータを書き込み
        if (userActivitySheet && allActivityData.length > 0) {
          userActivitySheet.getRange(2, 1, allActivityData.length, 6).setValues(allActivityData);
        }
        
        // ユーザー分析を更新
        updateNotionUserAnalytics();
        
        // 利用状況下位レポートを生成
        createNotionLowUsageReport();
      }
    } catch (notionError) {
      console.error('Notion分析エラー:', notionError);
      // エラーがあっても処理を継続
    }
    
    // メールレポートを送信（エラーが発生してもスクリプトは継続）
    try {
      sendEmailReport(allLicenses, analysisResults);
      console.log('レポートメールを送信しました');
    } catch (emailError) {
      console.error('メールレポート送信エラー:', emailError);
      // エラーログのみ記録し、ポップアップは表示しない
    }
    
    // 履歴を記録
    recordHistory(spreadsheet, '手動更新実行');
    
    console.log('ライセンス情報の更新が完了しました');
    
  } catch (error) {
    console.error('エラーが発生しました:', error);
    SpreadsheetApp.getUi().alert('エラー', 'エラーが発生しました: ' + error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

// ===== シート初期化（修正版） =====
function initializeSheets(spreadsheet = SpreadsheetApp.getActiveSpreadsheet()) {
  if (!spreadsheet) {
    throw new Error('スプレッドシートオブジェクトが無効です');
  }
  
  Object.values(CONFIG.SHEET_NAMES).forEach(sheetName => {
    try {
      let sheet = spreadsheet.getSheetByName(sheetName);
      if (!sheet) {
        console.log('シートを作成: ' + sheetName);
        sheet = spreadsheet.insertSheet(sheetName);
        setupSheetHeaders(sheet, sheetName);
      }
    } catch (error) {
      console.error('シート作成エラー: ' + sheetName, error);
    }
  });
}

// ===== シートヘッダー設定 =====
function setupSheetHeaders(sheet, sheetName) {
  const headers = {
    [CONFIG.SHEET_NAMES.LICENSES]: [
      'プラットフォーム', 'ユーザーID', 'メールアドレス', '氏名', 
      'ライセンスタイプ', 'ステータス', '割当日', '最終ログイン', 
      '利用状況', 'コスト（月額）', '備考', '管理メモ', '棚卸し有効期限', '最終更新日時'
    ],
    [CONFIG.SHEET_NAMES.USERS]: [
      'メールアドレス', '氏名', '部署', '役職', 
      'Google Workspace', 'トラスト・ログイン', 'ActiveGate', 
      'Microsoft 365', 'Canva', 'ESET', 'Slack', 
      'FortiClient', 'ファイルサーバー', 'v-Client', 
      'Zoom', 'Notion', 'QUDEN', 'ChatGPT', 
      'サクラサーバー', '合計コスト（月額）', '最終更新日時'
    ],
    [CONFIG.SHEET_NAMES.DIFF_ANALYSIS]: [
      'カテゴリ', 'メールアドレス', '氏名', 'Google Workspace', 
      '利用中プラットフォーム数', '不足プラットフォーム数', '不足プラットフォーム',
      '推奨アクション', '備考'
    ],
    [CONFIG.SHEET_NAMES.DASHBOARD]: [
      '指標', '値', '前回比', '備考'
    ],
    [CONFIG.SHEET_NAMES.HISTORY]: [
      '日時', '操作', '対象', '変更内容', '実行者'
    ],
    [CONFIG.SHEET_NAMES.SETTINGS]: [
      '設定項目', '値', '説明'
    ]
  };
  
  if (headers[sheetName]) {
    const headerRow = headers[sheetName];
    sheet.getRange(1, 1, 1, headerRow.length).setValues([headerRow]);
    sheet.getRange(1, 1, 1, headerRow.length)
      .setBackground('#4a86e8')
      .setFontColor('#ffffff')
      .setFontWeight('bold');
  }
}

// ===== 全ライセンス取得（シンプル版） =====
function fetchAllLicenses() {
  const allLicenses = [];
  
  // テスト用のダミーデータ
  allLicenses.push({
    platform: 'Google Workspace',
    userId: 'test001',
    email: 'test@example.com',
    name: 'テストユーザー',
    licenseType: 'Business Standard',
    status: 'Active',
    assignedDate: new Date(),
    lastLogin: new Date(),
    usage: '本日',
    cost: 1360,
    lastUpdated: new Date()
  });
  
  // 実際の実装時は各プラットフォームのAPI呼び出しを行う
  // allLicenses.push(...fetchGoogleWorkspaceLicenses());
  // allLicenses.push(...fetchTrustLoginLicenses());
  // など
  
  return allLicenses;
}

// ===== ライセンスシート更新 =====
function updateLicenseSheet(spreadsheet, licenses) {
  // 引数チェック
  if (!spreadsheet) {
    console.error('ライセンスシート更新エラー: スプレッドシートオブジェクトが渡されていません');
    return;
  }
  
  const sheet = spreadsheet.getSheetByName(CONFIG.SHEET_NAMES.LICENSES);
  if (!sheet) {
    console.error('ライセンス一覧シートが見つかりません');
    return;
  }
  
  // 既存データをクリア（ヘッダー以外）
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
  }
  
  // データを配列に変換
  if (licenses.length > 0) {
    const data = licenses.map(license => [
      license.platform,
      license.userId,
      license.email,
      license.name,
      license.licenseType,
      license.status,
      license.assignedDate,
      license.lastLogin,
      license.usage,
      license.cost,
      '',  // 備考
      '',  // 管理メモ
      '',  // 棚卸し有効期限
      license.lastUpdated
    ]);
    
    // データを書き込み
    sheet.getRange(2, 1, data.length, data[0].length).setValues(data);
  }
}

// ===== ダッシュボード更新 =====
function updateDashboard(spreadsheet, licenses) {
  // 引数チェック
  if (!spreadsheet) {
    console.error('ダッシュボード更新エラー: スプレッドシートオブジェクトが渡されていません');
    return;
  }
  
  const sheet = spreadsheet.getSheetByName(CONFIG.SHEET_NAMES.DASHBOARD);
  if (!sheet) {
    console.error('ダッシュボードシートが見つかりません');
    return;
  }
  
  // 統計情報を計算
  const stats = {
    totalLicenses: licenses.length,
    activeLicenses: licenses.filter(l => l.status === 'Active').length,
    totalCost: licenses.reduce((sum, l) => sum + (l.cost || 0), 0),
    platformCounts: {}
  };
  
  // プラットフォーム別のユーザー数を集計
  CONFIG.PLATFORMS.forEach(platform => {
    stats.platformCounts[platform] = licenses.filter(l => l.platform === platform).length;
  });
  
  // ダッシュボードデータ
  const dashboardData = [
    ['総ライセンス数', stats.totalLicenses, '', ''],
    ['アクティブライセンス数', stats.activeLicenses, '', ''],
    ['月額合計コスト', stats.totalCost, '', '円'],
    ['', '', '', ''], // 空行
    ['システム別ユーザー数', '', '', '']
  ];
  
  // プラットフォーム別のユーザー数を追加
  Object.entries(stats.platformCounts).forEach(([platform, count]) => {
    dashboardData.push([platform, count, '', 'ユーザー']);
  });
  
  // 既存データをクリア（ヘッダー以外）
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
  }
  
  // データを書き込み
  sheet.getRange(2, 1, dashboardData.length, 4).setValues(dashboardData);
  
  // フォーマット設定
  // システム別ユーザー数のヘッダーを強調
  const headerRow = dashboardData.findIndex(row => row[0].includes('システム別'));
  if (headerRow >= 0) {
    sheet.getRange(headerRow + 2, 1, 1, 4)
      .setBackground('#e0e0e0')
      .setFontWeight('bold');
  }
}

// ===== 履歴記録 =====
function recordHistory(spreadsheet, action, target = '', details = '') {
  const sheet = spreadsheet.getSheetByName(CONFIG.SHEET_NAMES.HISTORY);
  if (!sheet) {
    console.error('履歴シートが見つかりません');
    return;
  }
  
  const row = [
    new Date(),
    action,
    target,
    details,
    Session.getActiveUser().getEmail()
  ];
  
  sheet.appendRow(row);
}

// ===== カスタムメニュー =====
function onOpen() {
  try {
    const ui = SpreadsheetApp.getUi();
    ui.createMenu('ライセンス管理')
      .addItem('更新', 'main')
      .addItem('更新（棚卸し機能付き）', 'mainWithInventory')
      .addItem('手動更新', 'main')
      .addItem('初期設定', 'runInitialSetup')
      .addItem('動作テスト', 'testFunction')
      .addItem('設定', 'showSettings')
      .addSeparator()
      .addSubMenu(ui.createMenu('棚卸し管理')
        .addItem('棚卸し期間のハイライト更新', 'applyInventoryHighlight')
        .addItem('棚卸しステータス確認', 'checkInventoryStatus')
        .addSeparator()
        .addItem('保留シート初期化', 'resetInventoryHoldSheet')
        .addItem('保留シート手動更新', 'manualUpdateInventoryHoldSheet')
        .addItem('保留状況確認', 'checkInventoryHoldStatus')
        .addSeparator()
        .addItem('自動転記を開始（5分ごと）', 'setupInventoryHoldTrigger')
        .addItem('自動転記を停止', 'removeInventoryHoldTrigger'))
      .addSeparator()
      .addSubMenu(ui.createMenu('利用分析')
        .addItem('Notion利用分析', 'runNotionAnalysis')
        .addItem('Zoom利用分析', 'runZoomAnalysis')
        .addItem('全プラットフォーム分析', 'runAllAnalytics')
        .addSeparator()
        .addItem('Notion分析診断', 'diagnoseNotionAnalysis')
        .addItem('Zoom分析診断', 'diagnoseZoomAnalysis')
        .addItem('全分析診断', 'diagnoseAllAnalytics'))
      .addSeparator()
      .addItem('権限の初期化', 'initializePermissions')
      .addItem('Zoomテスト実行', 'testZoomAPI')
      .addItem('GWSユーザー手動インポート', 'importGoogleWorkspaceUsersFromCSV')
      .addToUi();
  } catch (error) {
    console.error('メニュー作成エラー:', error);
  }
}

// ===== 初期設定 =====
function runInitialSetup() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    
    if (!spreadsheet) {
      throw new Error('スプレッドシートが見つかりません');
    }
    
    // シートを初期化
    initializeSheets(spreadsheet);
    
    // 設定を保存
    const scriptProperties = PropertiesService.getScriptProperties();
    scriptProperties.setProperty('GOOGLE_DOMAIN', 'your-domain.com');
    scriptProperties.setProperty('NOTIFICATION_EMAIL', Session.getActiveUser().getEmail());
    
    SpreadsheetApp.getUi().alert(
      '初期設定完了',
      'シートの作成が完了しました。\n\n次に各プラットフォームのAPI設定を行ってください。',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
  } catch (error) {
    SpreadsheetApp.getUi().alert('エラー', '初期設定中にエラーが発生しました: ' + error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

// ===== テスト関数 =====
function testFunction() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    
    if (!spreadsheet) {
      throw new Error('スプレッドシートが見つかりません');
    }
    
    // シート存在確認
    const sheetNames = spreadsheet.getSheets().map(sheet => sheet.getName());
    
    let message = '動作テスト結果:\n\n';
    message += '✓ スプレッドシート名: ' + spreadsheet.getName() + '\n';
    message += '✓ 存在するシート: ' + sheetNames.join(', ') + '\n';
    message += '✓ ユーザー: ' + Session.getActiveUser().getEmail() + '\n';
    
    SpreadsheetApp.getUi().alert('テスト成功', message, SpreadsheetApp.getUi().ButtonSet.OK);
    
  } catch (error) {
    SpreadsheetApp.getUi().alert('テストエラー', error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

// ===== ユーティリティ関数 =====
function getSettingValue(key) {
  const scriptProperties = PropertiesService.getScriptProperties();
  return scriptProperties.getProperty(key);
}

function getApiToken(service) {
  const scriptProperties = PropertiesService.getScriptProperties();
  return scriptProperties.getProperty(`${service}_API_TOKEN`);
}

function getLicenseCost(platform, licenseType) {
  const costs = {
    'Google Workspace': CONFIG.GOOGLE_WORKSPACE.PRICING,
    'Zoom': CONFIG.ZOOM.PRICING,
    'Notion': CONFIG.NOTION.PRICING,
    // 他のプラットフォームの料金を追加
  };
  
  return costs[platform] ? costs[platform][licenseType] || 0 : 0;
}

function getSkuMapping() {
  return {
    'f245ecc8-75af-4f8e-b61f-27d8114de5f3': 'Business Basic',
    'cbdc14ab-d96c-4c30-b9f4-6ada7cdc1de4': 'Business Standard',
    '3b555118-da6a-4418-894f-7df1e2096870': 'Business Premium',
    'o365_business': 'Business',
    // 追加のSKUをここに
  };
}


// ===== 全ライセンス取得 =====
function fetchAllLicenses() {
  const allLicenses = [];
  
  // 各プラットフォームからライセンスを取得
  allLicenses.push(...fetchGoogleWorkspaceLicenses());     // GWS
  allLicenses.push(...fetchTrustLoginLicenses());         // トラスト・ログイン
  allLicenses.push(...fetchActiveGateLicenses());         // ActiveGate
  allLicenses.push(...fetchMicrosoft365Licenses());       // M365
  allLicenses.push(...fetchCanvaLicenses());              // Canva
  allLicenses.push(...fetchESETLicenses());               // ESET
  allLicenses.push(...fetchSlackLicenses());              // Slack
  allLicenses.push(...fetchFortiClientLicenses());        // FortiClient
  allLicenses.push(...fetchFileServerLicenses());         // ファイルサーバー
  allLicenses.push(...fetchVClientLicenses());            // v-Client
  allLicenses.push(...fetchZoomLicenses());               // Zoom
  allLicenses.push(...fetchNotionLicenses());             // Notion
  allLicenses.push(...fetchQUDENLicenses());              // QUDEN
  allLicenses.push(...fetchChatGPTLicenses());            // ChatGPT4
  allLicenses.push(...fetchSakuraServerLicenses());       // サクラサーバー
  
  // 最終更新日時を追加
  allLicenses.forEach(license => {
    license.lastUpdated = new Date();
  });
  
  return allLicenses;
}

// ===== Google Workspace ライセンス取得 =====
function fetchGoogleWorkspaceLicenses() {
  const licenses = [];
  
  try {
    // Admin SDK API の存在確認
    if (typeof AdminDirectory === 'undefined') {
      console.error('Admin SDK API が有効化されていません。');
      console.error('Apps Script エディタの左メニュー「サービス」から Admin SDK API を追加してください。');
      return licenses;
    }
    
    const domain = getSettingValue('GOOGLE_DOMAIN');
    if (!domain || domain === 'domain.com') {
      console.log('Google Workspace ドメインが設定されていません');
      return licenses;
    }
    
    // Admin Directory API を使用してユーザー一覧を取得
    try {
      // domainではなくcustomerパラメータを使用
      const response = AdminDirectory.Users.list({
        customer: 'my_customer',  // 現在のドメインのユーザーを取得
        maxResults: 500,
        orderBy: 'email',
        projection: 'FULL'  // 全ての情報を取得
      });
      
      if (response.users && response.users.length > 0) {
        response.users.forEach(user => {
          // Google Workspaceプランを設定から取得（デフォルトはBusiness Standard）
          const gwsPlan = getSettingValue('GOOGLE_WORKSPACE_PLAN') || 'Business Standard';
          const cost = CONFIG.GOOGLE_WORKSPACE.PRICING[gwsPlan] || 1360;
          
          // ライセンスタイプの表示（AdminかUserかを表示）
          let licenseType = gwsPlan;
          if (user.isAdmin) {
            licenseType = `${gwsPlan} (Admin)`;
          } else {
            licenseType = `${gwsPlan} (User)`;
          }
          
          // 退職者（suspended）のチェック - 退職者も料金がかかる
          const status = user.suspended ? 'Suspended' : 'Active';
          const name = user.suspended ? `${user.name.fullName} (退職)` : user.name.fullName;
          
          licenses.push({
            platform: 'Google Workspace',
            userId: user.id,
            email: user.primaryEmail,
            name: name,
            licenseType: licenseType,
            status: status,
            assignedDate: user.creationTime ? new Date(user.creationTime) : new Date(),
            lastLogin: user.lastLoginTime ? new Date(user.lastLoginTime) : null,
            usage: calculateUsage(user.lastLoginTime),
            cost: cost,
            suspended: user.suspended || false,
            orgUnitPath: user.orgUnitPath || '/',
            lastUpdated: new Date()
          });
        });
      }
      
      console.log(`Google Workspace: ${licenses.length}件のユーザーを取得`);
      
    } catch (apiError) {
      console.error('Admin Directory API エラー:', apiError);
      
      // エラーの詳細を確認
      if (apiError.details) {
        console.error('エラーコード:', apiError.details.code);
        console.error('エラーメッセージ:', apiError.details.message);
      }
      
      if (apiError.message && apiError.message.includes('Not Authorized')) {
        console.error('========================================');
        console.error('権限エラー: 以下を確認してください');
        console.error('1. 実行ユーザーがGoogle Workspace管理者であること');
        console.error('2. Admin SDK APIが有効化されていること');
        console.error('3. 正しいOAuthスコープが設定されていること');
        console.error('');
        console.error('代替案:');
        console.error('- メニューから「GWSユーザー手動インポート」を使用');
        console.error('========================================');
        
        // ダミーデータを返して処理を継続
        return getGoogleWorkspaceDummyData();
      } else if (apiError.message && apiError.message.includes('Invalid Input')) {
        console.error('========================================');
        console.error('入力エラー: APIパラメータを確認してください');
        console.error('========================================');
        return getGoogleWorkspaceDummyData();
      }
      
      // その他のエラー
      return getGoogleWorkspaceDummyData();
    }
    
  } catch (error) {
    console.error('Google Workspaceライセンス取得エラー:', error);
  }
  
  return licenses;
}









// ===== Google Workspace ダミーデータ =====
function getGoogleWorkspaceDummyData() {
  console.log('Google Workspace: 権限エラーのため、スキップします');
  console.log('管理者権限を取得後、再度実行してください');
  
  // 空の配列を返して処理を継続
  return [];
}

// ===== トラスト・ログイン ライセンス取得 =====
function fetchTrustLoginLicenses() {
  const licenses = [];
  
  try {
    const apiToken = getApiToken('TRUSTLOGIN');
    const url = 'https://api.trustlogin.com/v1/licenses';
    
    if (!apiToken) {
      console.log('トラスト・ログイン API設定が不足しています');
      return licenses;
    }
    
    const response = UrlFetchApp.fetch(url, {
      method: 'GET',
      headers: {
        'Authorization': `Bearer ${apiToken}`,
        'Content-Type': 'application/json'
      },
      muteHttpExceptions: true
    });
    
    if (response.getResponseCode() === 200) {
      const data = JSON.parse(response.getContentText());
      
      if (data.licenses) {
        data.licenses.forEach(license => {
          licenses.push({
            platform: 'トラスト・ログイン',
            userId: license.user_id,
            email: license.email,
            name: license.name || '',
            licenseType: license.plan,
            status: license.status === 'assigned' ? 'Active' : 'Inactive',
            assignedDate: new Date(license.assigned_at),
            lastLogin: license.last_login ? new Date(license.last_login) : null,
            usage: calculateUsage(license.last_login),
            cost: getLicenseCost('トラスト・ログイン', license.plan)
          });
        });
      }
    }
  } catch (error) {
    console.error('トラスト・ログイン ライセンス取得エラー:', error);
  }
  
  return licenses;
}

// ===== ActiveGate ライセンス取得 =====
function fetchActiveGateLicenses() {
  const licenses = [];
  
  try {
    const apiKey = getApiToken('ACTIVEGATE');
    const domain = getSettingValue('ACTIVEGATE_DOMAIN');
    const url = `https://${domain}.activegate.io/api/v1/licenses`;
    
    if (!apiKey || !domain) {
      console.log('ActiveGate API設定が不足しています');
      return licenses;
    }
    
    const response = UrlFetchApp.fetch(url, {
      method: 'GET',
      headers: {
        'X-API-Key': apiKey,
        'Content-Type': 'application/json'
      },
      muteHttpExceptions: true
    });
    
    if (response.getResponseCode() === 200) {
      const data = JSON.parse(response.getContentText());
      
      if (data.licenses) {
        data.licenses.forEach(license => {
          licenses.push({
            platform: 'ActiveGate',
            userId: license.userId,
            email: license.userId,
            name: license.userName || '',
            licenseType: license.licenseType,
            status: license.assigned ? 'Active' : 'Inactive',
            assignedDate: new Date(license.assignedDate),
            lastLogin: license.lastAccess ? new Date(license.lastAccess) : null,
            usage: calculateUsage(license.lastAccess),
            cost: getLicenseCost('ActiveGate', license.licenseType)
          });
        });
      }
    }
  } catch (error) {
    console.error('ActiveGate ライセンス取得エラー:', error);
  }
  
  return licenses;
}

// ===== Microsoft 365 ライセンス取得 =====
function fetchMicrosoft365Licenses() {
  const licenses = [];
  
  try {
    const accessToken = getMicrosoftAccessToken();
    
    if (!accessToken) {
      console.log('Microsoft 365 アクセストークンの取得に失敗しました');
      return licenses;
    }
    
    const url = 'https://graph.microsoft.com/v1.0/users?$select=displayName,mail,userPrincipalName,accountEnabled,assignedLicenses';
    
    const response = UrlFetchApp.fetch(url, {
      method: 'GET',
      headers: {
        'Authorization': `Bearer ${accessToken}`,
        'Content-Type': 'application/json'
      },
      muteHttpExceptions: true
    });
    
    if (response.getResponseCode() === 200) {
      const data = JSON.parse(response.getContentText());
      
      if (data.value) {
        data.value.forEach(user => {
          const licenseType = user.assignedLicenses.length > 0 ? getSkuMapping()[user.assignedLicenses[0].skuId] || 'Unknown' : 'No License';
          
          licenses.push({
            platform: 'Microsoft 365',
            userId: user.id,
            email: user.userPrincipalName || user.mail,
            name: user.displayName,
            licenseType: licenseType,
            status: user.accountEnabled ? 'Active' : 'Disabled',
            assignedDate: new Date(),
            lastLogin: null,
            usage: '不明',
            cost: getLicenseCost('Microsoft 365', licenseType),
            lastUpdated: new Date()
          });
        });
      }
    } else {
      console.error('Microsoft 365 API エラー:', response.getContentText());
    }
    
  } catch (error) {
    console.error('Microsoft 365 ライセンス取得エラー:', error);
  }
  
  return licenses;
}

// ===== Canva ライセンス取得 =====
function fetchCanvaLicenses() {
  const licenses = [];
  
  try {
    const apiToken = getApiToken('CANVA');
    const teamId = getSettingValue('CANVA_TEAM_ID');
    
    if (!apiToken || !teamId) {
      console.log('Canva API設定が不足しています');
      return licenses;
    }
    
    const url = `https://api.canva.com/rest/v1/teams/${teamId}/members`;
    
    const response = UrlFetchApp.fetch(url, {
      method: 'GET',
      headers: {
        'Authorization': `Bearer ${apiToken}`,
        'Content-Type': 'application/json'
      },
      muteHttpExceptions: true
    });
    
    if (response.getResponseCode() === 200) {
      const data = JSON.parse(response.getContentText());
      
      if (data.members) {
        data.members.forEach(member => {
          licenses.push({
            platform: 'Canva',
            userId: member.user_id,
            email: member.email,
            name: member.display_name || member.email,
            licenseType: member.role || 'Member',
            status: member.status === 'active' ? 'Active' : 'Inactive',
            assignedDate: new Date(member.joined_at),
            lastLogin: member.last_active_at ? new Date(member.last_active_at) : null,
            usage: calculateUsage(member.last_active_at),
            cost: getLicenseCost('Canva', member.role)
          });
        });
      }
    }
  } catch (error) {
    console.error('Canva ライセンス取得エラー:', error);
  }
  
  return licenses;
}

// ===== ESET ライセンス取得 =====
function fetchESETLicenses() {
  const licenses = [];
  
  try {
    const apiKey = getApiToken('ESET');
    const companyId = getSettingValue('ESET_COMPANY_ID');
    
    if (!apiKey || !companyId) {
      console.log('ESET API設定が不足しています');
      return licenses;
    }
    
    const url = `https://era.eset.com/api/v1/companies/${companyId}/licenses`;
    
    const response = UrlFetchApp.fetch(url, {
      method: 'GET',
      headers: {
        'Authorization': `Bearer ${apiKey}`,
        'Accept': 'application/json'
      },
      muteHttpExceptions: true
    });
    
    if (response.getResponseCode() === 200) {
      const data = JSON.parse(response.getContentText());
      
      if (data.licenses) {
        data.licenses.forEach(license => {
          if (license.assigned_devices) {
            license.assigned_devices.forEach(device => {
              licenses.push({
                platform: 'ESET',
                userId: device.device_id,
                email: device.user_email || `device_${device.device_name}`,
                name: device.user_name || device.device_name,
                licenseType: license.product_name,
                status: license.status === 'active' ? 'Active' : 'Inactive',
                assignedDate: new Date(device.assigned_date),
                lastLogin: device.last_seen ? new Date(device.last_seen) : null,
                usage: calculateUsage(device.last_seen),
                cost: getLicenseCost('ESET', license.product_name)
              });
            });
          }
        });
      }
    }
  } catch (error) {
    console.error('ESET ライセンス取得エラー:', error);
  }
  
  return licenses;
}

// ===== Slack ライセンス取得 =====
function fetchSlackLicenses() {
  const licenses = [];
  
  try {
    const apiToken = getApiToken('SLACK');
    
    if (!apiToken) {
      console.log('Slack API設定が不足しています');
      return licenses;
    }
    
    // 1. まずユーザー一覧を取得
    const usersUrl = 'https://slack.com/api/users.list';
    
    const usersResponse = UrlFetchApp.fetch(usersUrl, {
      method: 'GET',
      headers: {
        'Authorization': `Bearer ${apiToken}`,
        'Content-Type': 'application/json'
      },
      muteHttpExceptions: true
    });
    
    if (usersResponse.getResponseCode() !== 200) {
      console.error('Slack users.list APIエラー:', usersResponse.getContentText());
      return licenses;
    }
    
    const usersData = JSON.parse(usersResponse.getContentText());
    
    if (!usersData.ok || !usersData.members) {
      console.error('Slack users.list レスポンスエラー:', usersData.error);
      return licenses;
    }
    
    // 2. 課金情報を取得（利用可能な場合）
    let billableInfo = {};
    try {
      const billableUrl = 'https://slack.com/api/team.billableInfo';
      const billableResponse = UrlFetchApp.fetch(billableUrl, {
        method: 'GET',
        headers: {
          'Authorization': `Bearer ${apiToken}`,
          'Content-Type': 'application/json'
        },
        muteHttpExceptions: true
      });
      
      if (billableResponse.getResponseCode() === 200) {
        const billableData = JSON.parse(billableResponse.getContentText());
        if (billableData.ok && billableData.billable_info) {
          billableInfo = billableData.billable_info;
        }
      }
    } catch (billableError) {
      console.log('Slack 課金情報の取得に失敗しました（管理者権限が必要）:', billableError);
    }
    
    // 3. ユーザー情報を処理
    usersData.members.forEach(member => {
      // ボット、アプリユーザー、Slackbotを除外
      if (member.is_bot || member.is_app_user || member.id === 'USLACKBOT') {
        return;
      }
      
      // 削除済み（無効化）ユーザーを除外
      if (member.deleted) {
        return;
      }
      
      // 課金情報がある場合、課金対象でないユーザーを除外
      if (billableInfo[member.id] && !billableInfo[member.id].billing_active) {
        return;
      }
      
      // ゲストユーザーでかつ、制限されているまたは超制限されている場合は除外（通常は課金対象外）
      if (member.is_restricted || member.is_ultra_restricted) {
        // マルチチャネルゲストまたはシングルチャネルゲストの場合
        // 課金ポリシーによってはカウントされない可能性があるため、設定で制御できるようにする
        const includeGuests = getSettingValue('SLACK_INCLUDE_GUESTS') !== 'false';
        if (!includeGuests) {
          return;
        }
      }
      
      const licenseType = member.is_admin ? 'Admin' : 
                         member.is_owner ? 'Owner' : 
                         member.is_restricted ? 'Multi-Channel Guest' : 
                         member.is_ultra_restricted ? 'Single-Channel Guest' : 'Member';
      
      licenses.push({
        platform: 'Slack',
        userId: member.id,
        email: member.profile.email || `${member.name}@slack`,
        name: member.real_name || member.name,
        licenseType: licenseType,
        status: 'Active',
        assignedDate: new Date(member.updated * 1000),
        lastLogin: null,
        usage: 'アクティブ',
        cost: getLicenseCost('Slack', licenseType)
      });
    });
    
    console.log(`Slack: ${licenses.length}名のアクティブユーザーを取得しました`);
    
  } catch (error) {
    console.error('Slack ライセンス取得エラー:', error);
  }
  
  return licenses;
}

// ===== その他のプラットフォーム用のスタブ関数 =====
// 実装が必要な場合は、上記の関数を参考に実装してください

function fetchFortiClientLicenses() {
  return [];
}

function fetchFileServerLicenses() {
  return [];
}

function fetchVClientLicenses() {
  return [];
}

// ===== Zoom ライセンス取得（実装版） =====
function fetchZoomLicenses() {
  const licenses = [];
  
  try {
    // Zoom OAuth トークンを取得
    const accessToken = getZoomAccessToken();
    
    if (!accessToken) {
      console.log('Zoom アクセストークンの取得に失敗しました');
      return licenses;
    }
    
    // ユーザー一覧を取得
    const url = 'https://api.zoom.us/v2/users?page_size=300&status=active';
    
    const response = UrlFetchApp.fetch(url, {
      method: 'GET',
      headers: {
        'Authorization': `Bearer ${accessToken}`,
        'Content-Type': 'application/json'
      },
      muteHttpExceptions: true
    });
    
    console.log('Zoom API レスポンスコード:', response.getResponseCode());
    
    if (response.getResponseCode() === 200) {
      const data = JSON.parse(response.getContentText());
      
      if (data.users) {
        console.log(`Zoom: ${data.users.length}件のユーザーを取得`);
        
        data.users.forEach(user => {
          // ライセンスタイプのマッピング
          // type: 1=Basic(無料), 2=Pro(有償), 3=Business/Enterprise(有償)
          let licenseType = 'Basic';
          
          // Zoomのライセンスタイプを詳細に判定
          if (user.type === 1) {
            licenseType = 'Basic';
          } else if (user.type === 2) {
            licenseType = 'Pro';
          } else if (user.type === 3) {
            // Business/Enterpriseの判定
            if (user.plan_type && user.plan_type.includes('enterprise')) {
              licenseType = 'Enterprise';
            } else {
              licenseType = 'Business';
            }
          }
          
          // ライセンスの詳細情報を含める（有償/無償の判別のため）
          const licenseDetail = user.type === 1 ? 'Basic (無料)' : `${licenseType} (有償)`;
          
          // コストの取得
          const cost = CONFIG.ZOOM.PRICING[licenseType] || 0;
          
          licenses.push({
            platform: 'Zoom',
            userId: user.id,
            email: user.email,
            name: `${user.first_name} ${user.last_name}`.trim() || user.display_name || user.email,
            licenseType: licenseDetail,  // 有償/無償の表示を含む
            status: user.status === 'active' ? 'Active' : 'Inactive',
            assignedDate: user.created_at ? new Date(user.created_at) : new Date(),
            lastLogin: user.last_login_time ? new Date(user.last_login_time) : null,
            usage: calculateUsage(user.last_login_time),
            cost: cost,
            lastUpdated: new Date()
          });
        });
        
        // ページネーション対応（必要に応じて）
        if (data.page_count > 1) {
          console.log(`Zoom: 複数ページあり（全${data.page_count}ページ）`);
          // 追加のページ取得処理をここに実装
        }
      }
    } else {
      console.error('Zoom API エラー:', response.getContentText());
    }
  } catch (error) {
    console.error('Zoom ライセンス取得エラー:', error);
  }
  
  return licenses;
}

// ===== Zoom アクセストークン取得 =====
function getZoomAccessToken() {
  try {
    const accountId = getSettingValue('ZOOM_ACCOUNT_ID');
    const clientId = getSettingValue('ZOOM_CLIENT_ID');
    const clientSecret = getSettingValue('ZOOM_CLIENT_SECRET');
    
    if (!accountId || !clientId || !clientSecret) {
      console.log('Zoom API認証情報が設定されていません');
      return null;
    }
    
    // Server-to-Server OAuth用のトークン取得
    const url = 'https://zoom.us/oauth/token';
    
    // Basic認証用のクレデンシャルをBase64エンコード
    const credentials = Utilities.base64Encode(`${clientId}:${clientSecret}`);
    
    const response = UrlFetchApp.fetch(url, {
      method: 'POST',
      headers: {
        'Authorization': `Basic ${credentials}`,
        'Content-Type': 'application/x-www-form-urlencoded'
      },
      payload: {
        'grant_type': 'account_credentials',
        'account_id': accountId
      },
      muteHttpExceptions: true
    });
    
    if (response.getResponseCode() === 200) {
      const data = JSON.parse(response.getContentText());
      return data.access_token;
    } else {
      console.error('Zoom トークン取得エラー:', response.getContentText());
      return null;
    }
    
  } catch (error) {
    console.error('Zoom アクセストークン取得エラー:', error);
    return null;
  }
}

// ===== Zoom関連の定数 =====
const ZOOM_ANALYSIS_SHEET_NAMES = {
  MEETING_LIST: 'Zoom_MeetingList',
  ZOOM_LOG: 'Zoom_Log',
  USER_ANALYTICS: 'Zoom_UserAnalytics',
  USER_ACTIVITY: 'Zoom_UserActivity',
  USAGE_STATS: 'Zoom_UsageStats',
  LOW_USAGE_REPORT: 'Zoom_利用状況下位'
};

// ===== Zoom 利用分析機能 =====

/**
 * Zoom分析機能の初期セットアップ
 */
function setupZoomAnalytics() {
  // ミーティングリストシート
  const meetingListSheet = getOrCreateSheet(ZOOM_ANALYSIS_SHEET_NAMES.MEETING_LIST);
  if (meetingListSheet) {
    meetingListSheet.getRange(1, 1, 1, 8).setValues([
      ['Meeting ID', 'Topic', 'Start Time', 'Duration', 'Host', 'Participants', 'Status', 'Join URL']
    ]);
    meetingListSheet.getRange(1, 1, 1, 8).setFontWeight('bold');
  }
  
  // ログシート
  const logSheet = getOrCreateSheet(ZOOM_ANALYSIS_SHEET_NAMES.ZOOM_LOG);
  if (logSheet) {
    logSheet.getRange(1, 1, 1, 7).setValues([
      ['記録日時', 'Activity Type', 'Meeting ID', 'Topic', 'User', 'Action', 'Details']
    ]);
    logSheet.getRange(1, 1, 1, 7).setFontWeight('bold');
  }
  
  // ユーザー分析シート
  const userAnalyticsSheet = getOrCreateSheet(ZOOM_ANALYSIS_SHEET_NAMES.USER_ANALYTICS);
  if (userAnalyticsSheet) {
    userAnalyticsSheet.getRange(1, 1, 1, 10).setValues([
      ['User Name', 'Email', 'Total Meetings', 'Total Duration (min)', 'Avg Duration (min)', 
       'Last Activity', 'This Week', 'This Month', 'Host Count', 'Participant Count']
    ]);
    userAnalyticsSheet.getRange(1, 1, 1, 10).setFontWeight('bold');
  }
  
  // ユーザー活動履歴シート
  const userActivitySheet = getOrCreateSheet(ZOOM_ANALYSIS_SHEET_NAMES.USER_ACTIVITY);
  if (userActivitySheet) {
    userActivitySheet.getRange(1, 1, 1, 8).setValues([
      ['記録日時', 'User Name', 'Email', 'Meeting Topic', 'Role', 'Join Time', 'Leave Time', 'Duration (min)']
    ]);
    userActivitySheet.getRange(1, 1, 1, 8).setFontWeight('bold');
  }
  
  // 使用統計シート
  const usageStatsSheet = getOrCreateSheet(ZOOM_ANALYSIS_SHEET_NAMES.USAGE_STATS);
  if (usageStatsSheet) {
    usageStatsSheet.getRange(1, 1, 1, 6).setValues([
      ['Date', 'Total Meetings', 'Total Participants', 'Total Duration (min)', 'Avg Meeting Duration', 'Peak Hour']
    ]);
    usageStatsSheet.getRange(1, 1, 1, 6).setFontWeight('bold');
  }
  
  // 利用状況下位レポートシート
  const lowUsageSheet = getOrCreateSheet(ZOOM_ANALYSIS_SHEET_NAMES.LOW_USAGE_REPORT);
  if (lowUsageSheet) {
    // ヘッダー設定は createZoomLowUsageReport 関数で行うため、ここでは初期化のみ
    console.log('Zoom利用状況下位レポートシートを初期化しました');
  }
  
  console.log('Zoom分析機能のセットアップが完了しました');
}

/**
 * Zoomミーティング履歴を取得
 */
function fetchZoomMeetings() {
  try {
    const accessToken = getZoomAccessToken();
    if (!accessToken) {
      console.log('Zoom アクセストークンの取得に失敗しました');
      return;
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const meetingListSheet = ss.getSheetByName(ZOOM_ANALYSIS_SHEET_NAMES.MEETING_LIST);
    
    // アクティブユーザー一覧を取得
    const users = fetchZoomActiveUsers(accessToken);
    console.log(`アクティブユーザー数: ${users.length}`);
    
    if (users.length === 0) {
      console.log('アクティブユーザーが見つかりませんでした');
      return;
    }
    
    // 各ユーザーのミーティングを取得
    const allMeetings = [];
    
    for (const user of users) {
      try {
        console.log(`ユーザー ${user.email} のミーティングを取得中...`);
        
        // スケジュール済みミーティング
        const scheduledMeetings = fetchUserMeetingsZoom(accessToken, user.id, 'scheduled');
        
        // 過去のミーティング
        const pastMeetings = fetchUserMeetingsZoom(accessToken, user.id, 'previous_meetings');
        
        // ミーティング情報にホスト情報を追加
        [...scheduledMeetings, ...pastMeetings].forEach(meeting => {
          meeting.host_email = user.email;
          meeting.host_name = `${user.first_name} ${user.last_name}`.trim();
          allMeetings.push(meeting);
        });
        
        // レート制限対策
        Utilities.sleep(100);
        
      } catch (error) {
        console.warn(`ユーザー ${user.email} でエラー:`, error);
      }
    }
    
    console.log(`総取得ミーティング数: ${allMeetings.length}`);
    
    // 過去7日間のミーティングをフィルタリング
    const weekAgo = new Date();
    weekAgo.setDate(weekAgo.getDate() - 7);
    
    const recentMeetings = allMeetings.filter(meeting => {
      if (!meeting.start_time && !meeting.created_at) return false;
      const meetingDate = new Date(meeting.start_time || meeting.created_at);
      return meetingDate >= weekAgo;
    });
    
    console.log(`過去7日間のミーティング: ${recentMeetings.length} 件`);
    
    // データをシートに書き込み
    writeMeetingsToSheet(meetingListSheet, recentMeetings.length > 0 ? recentMeetings : allMeetings);
    
    // ユーザー分析を更新
    updateZoomUserAnalytics();
    
    console.log(`ミーティング取得完了: 表示ミーティング数: ${recentMeetings.length > 0 ? recentMeetings.length : allMeetings.length}, 確認したユーザー数: ${users.length}, 期間: ${recentMeetings.length > 0 ? '過去7日間' : '全期間'}`);
    
  } catch (error) {
    console.error('ミーティング取得エラー:', error);
  }
}

/**
 * Zoomアクティブユーザー一覧を取得
 */
function fetchZoomActiveUsers(accessToken) {
  try {
    const url = 'https://api.zoom.us/v2/users?status=active&page_size=30';
    
    const options = {
      method: 'GET',
      headers: {
        'Authorization': `Bearer ${accessToken}`,
        'Content-Type': 'application/json'
      },
      muteHttpExceptions: true
    };
    
    const response = UrlFetchApp.fetch(url, options);
    const code = response.getResponseCode();
    
    if (code === 200) {
      const data = JSON.parse(response.getContentText());
      return data.users || [];
    } else {
      console.warn(`ユーザー取得失敗 (${code}): ${response.getContentText()}`);
      return [];
    }
    
  } catch (error) {
    console.error('ユーザー取得エラー:', error);
    return [];
  }
}

/**
 * 特定ユーザーのミーティングを取得
 */
function fetchUserMeetingsZoom(accessToken, userId, type = 'scheduled') {
  try {
    const url = `https://api.zoom.us/v2/users/${userId}/meetings?type=${type}&page_size=50`;
    
    const options = {
      method: 'GET',
      headers: {
        'Authorization': `Bearer ${accessToken}`,
        'Content-Type': 'application/json'
      },
      muteHttpExceptions: true
    };
    
    const response = UrlFetchApp.fetch(url, options);
    const code = response.getResponseCode();
    
    if (code === 200) {
      const data = JSON.parse(response.getContentText());
      return data.meetings || [];
    } else {
      console.warn(`ユーザー ${userId} のミーティング取得失敗 (${code})`);
      return [];
    }
    
  } catch (error) {
    console.error(`ユーザー ${userId} のミーティング取得エラー:`, error);
    return [];
  }
}

/**
 * ミーティングデータをシートに書き込み
 */
function writeMeetingsToSheet(sheet, meetings) {
  // 既存のデータをクリア（ヘッダー以外）
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, 8).clearContent();
  }
  
  if (meetings.length === 0) return;
  
  const meetingData = meetings.map(meeting => [
    meeting.id || meeting.uuid || 'Unknown',
    meeting.topic || 'No Topic',
    meeting.start_time || meeting.created_at || '',
    meeting.duration || meeting.planned_duration || 0,
    meeting.host_email || meeting.host_name || '',
    meeting.participants_count || 0,
    meeting.status || 'Scheduled',
    meeting.join_url || ''
  ]);
  
  sheet.getRange(2, 1, meetingData.length, 8).setValues(meetingData);
  sheet.autoResizeColumns(1, 8);
}

/**
 * Zoomユーザー別活動分析を更新
 */
function updateZoomUserAnalytics() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const meetingListSheet = ss.getSheetByName(ZOOM_ANALYSIS_SHEET_NAMES.MEETING_LIST);
  const userAnalyticsSheet = ss.getSheetByName(ZOOM_ANALYSIS_SHEET_NAMES.USER_ANALYTICS);
  
  const meetingLastRow = meetingListSheet.getLastRow();
  if (meetingLastRow <= 1) return;
  
  const meetingData = meetingListSheet.getRange(2, 1, meetingLastRow - 1, 8).getValues();
  
  // ホスト別統計を集計
  const hostStats = {};
  
  meetingData.forEach(([meetingId, topic, startTime, duration, host, participants, status, joinUrl]) => {
    if (!host || host === '') return;
    
    if (!hostStats[host]) {
      hostStats[host] = {
        name: host,
        email: host,
        totalMeetings: 0,
        totalDuration: 0,
        lastActivity: null,
        weekMeetings: 0,
        monthMeetings: 0,
        hostCount: 0,
        participantCount: 0
      };
    }
    
    hostStats[host].totalMeetings++;
    hostStats[host].totalDuration += duration || 0;
    hostStats[host].hostCount++;
    
    // 最終活動日時を更新
    if (startTime) {
      const activityDate = new Date(startTime);
      if (!hostStats[host].lastActivity || activityDate > hostStats[host].lastActivity) {
        hostStats[host].lastActivity = activityDate;
      }
      
      // 期間別カウント
      const now = new Date();
      const weekAgo = new Date(now.getTime() - 7 * 24 * 60 * 60 * 1000);
      const monthAgo = new Date(now.getTime() - 30 * 24 * 60 * 60 * 1000);
      
      if (activityDate >= weekAgo) {
        hostStats[host].weekMeetings++;
      }
      if (activityDate >= monthAgo) {
        hostStats[host].monthMeetings++;
      }
    }
  });
  
  // 分析結果をシートに書き込み
  const analyticsData = [];
  
  Object.values(hostStats).forEach(user => {
    const avgDuration = user.totalMeetings > 0 ? Math.round(user.totalDuration / user.totalMeetings) : 0;
    const lastActivityStr = user.lastActivity ? user.lastActivity.toLocaleString('ja-JP') : '';
    
    analyticsData.push([
      user.name,
      user.email,
      user.totalMeetings,
      user.totalDuration,
      avgDuration,
      lastActivityStr,
      user.weekMeetings,
      user.monthMeetings,
      user.hostCount,
      user.participantCount
    ]);
  });
  
  // 総参加時間でソート（降順）
  analyticsData.sort((a, b) => b[3] - a[3]);
  
  // 既存データをクリア
  const analyticsLastRow = userAnalyticsSheet.getLastRow();
  if (analyticsLastRow > 1) {
    userAnalyticsSheet.getRange(2, 1, analyticsLastRow - 1, 10).clearContent();
  }
  
  // 新しいデータを書き込み
  if (analyticsData.length > 0) {
    userAnalyticsSheet.getRange(2, 1, analyticsData.length, 10).setValues(analyticsData);
  }
  
  console.log(`Zoomユーザー分析を更新しました: ${analyticsData.length} ユーザー`);
}

// ===== Microsoft 365 アクセストークン取得 =====
function getMicrosoftAccessToken() {
  try {
    const tenantId = getSettingValue('M365_TENANT_ID');
    const clientId = getSettingValue('M365_CLIENT_ID');
    const clientSecret = getSettingValue('M365_CLIENT_SECRET');
    
    if (!tenantId || !clientId || !clientSecret) {
      console.log('Microsoft 365 API認証情報が設定されていません');
      return null;
    }
    
    const url = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;
    
    const response = UrlFetchApp.fetch(url, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded'
      },
      payload: {
        'client_id': clientId,
        'client_secret': clientSecret,
        'scope': 'https://graph.microsoft.com/.default',
        'grant_type': 'client_credentials'
      },
      muteHttpExceptions: true
    });
    
    if (response.getResponseCode() === 200) {
      const data = JSON.parse(response.getContentText());
      return data.access_token;
    } else {
      console.error('Microsoft 365 トークン取得エラー:', response.getContentText());
      return null;
    }
    
  } catch (error) {
    console.error('Microsoft 365 アクセストークン取得エラー:', error);
    return null;
  }
}

// ===== Zoom API テスト関数 =====
function testZoomAPI() {
  try {
    const licenses = fetchZoomLicenses();
    
    if (licenses.length > 0) {
      SpreadsheetApp.getUi().alert(
        'Zoom APIテスト成功',
        `${licenses.length}件のZoomライセンスを取得しました。\n\n` +
        `サンプル:\n` +
        `- メール: ${licenses[0].email}\n` +
        `- 名前: ${licenses[0].name}\n` +
        `- ライセンス: ${licenses[0].licenseType}\n` +
        `- コスト: ${licenses[0].cost}円/月`,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    } else {
      SpreadsheetApp.getUi().alert(
        'Zoom APIテスト',
        'Zoomライセンスが見つかりません。\n\n' +
        '設定を確認してください:\n' +
        '- Account ID\n' +
        '- Client ID\n' +
        '- Client Secret',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
  } catch (error) {
    SpreadsheetApp.getUi().alert(
      'Zoom APIエラー',
      'エラーが発生しました:\n' + error.toString(),
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

// ===== Notion関連の定数 =====
const NOTION_VERSION = '2022-06-28';
const NOTION_ANALYSIS_SHEET_NAMES = {
  DATABASE_CONFIG: 'Notion_DBConfig',
  PAGE_LIST: 'Notion_PageList',
  NOTION_LOG: 'Notion_Log',
  USER_ANALYTICS: 'Notion_UserAnalytics',
  USER_ACTIVITY: 'Notion_UserActivity',
  LOW_USAGE_REPORT: 'Notion_利用状況下位'
};

/**
 * シートを取得または作成するヘルパー関数
 */
function getOrCreateSheet(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    try {
      sheet = ss.insertSheet(sheetName);
    } catch (e) {
      console.error('シート作成エラー: ' + sheetName, e);
      // シート名が既に存在する場合もあるので再度取得を試みる
      sheet = ss.getSheetByName(sheetName);
    }
  }
  
  return sheet;
}

// ===== Notion ライセンス取得 =====
function fetchNotionLicenses() {
  const licenses = [];
  
  try {
    const notionToken = getSettingValue('NOTION_TOKEN');
    const notionPlan = getSettingValue('NOTION_PLAN') || 'Plus'; // デフォルトはPlusプラン
    
    if (!notionToken) {
      console.log('Notion APIトークンが設定されていません');
      return licenses;
    }
    
    // Notion APIでワークスペースのユーザー一覧を取得
    const url = 'https://api.notion.com/v1/users';
    const headers = {
      'Authorization': 'Bearer ' + notionToken,
      'Notion-Version': NOTION_VERSION
    };
    
    const options = {
      'method': 'get',
      'headers': headers,
      'muteHttpExceptions': true
    };
    
    const response = UrlFetchApp.fetch(url, options);
    
    if (response.getResponseCode() === 200) {
      const data = JSON.parse(response.getContentText());
      
      if (data.results) {
        data.results.forEach(function(user) {
          // ボットユーザーは除外
          if (user.type === 'bot') return;
          
          // プラン名を見やすく変換
          const displayPlanName = notionPlan
            .replace('_Monthly', ' (月契約)')
            .replace('Plus', 'Plus')
            .replace('Business', 'Business')
            .replace('Enterprise', 'Enterprise');
          
          const licenseInfo = {
            platform: 'Notion',
            userId: user.id,
            email: user.person && user.person.email ? user.person.email : '',
            name: user.name || '',
            licenseType: displayPlanName, // 表示用のプラン名
            status: user.object === 'user' ? 'Active' : 'Inactive',
            assignedDate: new Date(), // Notionは割当日の情報を提供しない
            lastLogin: null, // Notionは最終ログイン情報を提供しない
            usage: '利用中', // 詳細な利用状況はNotion APIでは取得不可
            cost: getLicenseCost('Notion', notionPlan)
          };
          licenses.push(licenseInfo);
        });
      }
    } else {
      console.error('Notion API エラー:', response.getContentText());
    }
  } catch (error) {
    console.error('Notion ライセンス取得エラー:', error);
  }
  
  return licenses;
}

// ===== Notion 利用分析機能 =====

/**
 * Notion分析機能の初期セットアップ
 */
function setupNotionAnalytics() {
  // データベース設定シート
  const configSheet = getOrCreateSheet(NOTION_ANALYSIS_SHEET_NAMES.DATABASE_CONFIG);
  if (configSheet) {
    configSheet.getRange(1, 1, 1, 6).setValues([
      ['データベースID', 'データベース名', '有効/無効', '最終取得日時', '作成日時', 'URL']
    ]);
    configSheet.getRange(1, 1, 1, 6).setFontWeight('bold');
  }
  
  // ページリストシート
  const pageListSheet = getOrCreateSheet(NOTION_ANALYSIS_SHEET_NAMES.PAGE_LIST);
  if (pageListSheet) {
    pageListSheet.getRange(1, 1, 1, 5).setValues([
      ['データベース名', 'Page ID', 'Page Title', '最終編集日時', '最終編集者']
    ]);
    pageListSheet.getRange(1, 1, 1, 5).setFontWeight('bold');
  }
  
  // ログシート
  const logSheet = getOrCreateSheet(NOTION_ANALYSIS_SHEET_NAMES.NOTION_LOG);
  if (logSheet) {
    logSheet.getRange(1, 1, 1, 7).setValues([
      ['記録日時', 'データベース名', 'Page ID', 'Page Title', '最終編集日時', '最終編集者', '変更検出']
    ]);
    logSheet.getRange(1, 1, 1, 7).setFontWeight('bold');
  }
  
  // ユーザー分析シート
  const userAnalyticsSheet = getOrCreateSheet(NOTION_ANALYSIS_SHEET_NAMES.USER_ANALYTICS);
  if (userAnalyticsSheet) {
    userAnalyticsSheet.getRange(1, 1, 1, 9).setValues([
      ['ユーザー名', '編集ページ数', '編集DB数', '最終活動日', '今日の編集数', '今週の編集数', '今月の編集数', '総編集回数', '平均編集間隔(日)']
    ]);
    userAnalyticsSheet.getRange(1, 1, 1, 9).setFontWeight('bold');
  }
  
  // ユーザー活動履歴シート
  const userActivitySheet = getOrCreateSheet(NOTION_ANALYSIS_SHEET_NAMES.USER_ACTIVITY);
  if (userActivitySheet) {
    userActivitySheet.getRange(1, 1, 1, 6).setValues([
      ['記録日時', 'データベース名', 'ユーザー名', 'Page Title', '活動種別', 'Page ID']
    ]);
    userActivitySheet.getRange(1, 1, 1, 6).setFontWeight('bold');
  }
  
  // 利用状況下位レポートシート
  const lowUsageSheet = getOrCreateSheet(NOTION_ANALYSIS_SHEET_NAMES.LOW_USAGE_REPORT);
  if (lowUsageSheet) {
    // ヘッダー設定は createNotionLowUsageReport 関数で行うため、ここでは初期化のみ
    console.log('Notion利用状況下位レポートシートを初期化しました');
  }
  
  console.log('Notion分析機能のセットアップが完了しました');
}

/**
 * Notion API を使用してデータベースを自動検出
 */
function discoverNotionDatabases() {
  try {
    const notionToken = getSettingValue('NOTION_TOKEN');
    if (!notionToken) {
      console.error('NotionのAPIトークンが設定されていません。設定画面から設定してください。');
      return [];
    }
    
    const url = 'https://api.notion.com/v1/search';
    const payload = {
      filter: {
        value: 'database',
        property: 'object'
      },
      page_size: 100
    };
    
    const options = {
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${notionToken}`,
        'Notion-Version': NOTION_VERSION,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };
    
    let allDatabases = [];
    let hasMore = true;
    let startCursor = null;
    
    while (hasMore) {
      if (startCursor) {
        payload.start_cursor = startCursor;
      }
      
      options.payload = JSON.stringify(payload);
      const response = UrlFetchApp.fetch(url, options);
      const code = response.getResponseCode();
      
      if (code !== 200) {
        throw new Error(`データベース検索エラー (Status: ${code}): ${response.getContentText()}`);
      }
      
      const data = JSON.parse(response.getContentText());
      allDatabases.push(...data.results);
      
      hasMore = data.has_more;
      startCursor = data.next_cursor;
    }
    
    // データベース情報を整理
    const databaseConfigs = allDatabases.map(db => {
      let dbName = 'Untitled Database';
      
      if (db.title && db.title.length > 0) {
        dbName = db.title.map(t => t.plain_text).join('');
      }
      
      return {
        id: db.id,
        name: dbName,
        enabled: true,
        url: db.url || '',
        created: db.created_time || '',
        lastEdited: db.last_edited_time || ''
      };
    });
    
    return databaseConfigs;
    
  } catch (error) {
    console.error('データベース検索エラー:', error);
    throw error;
  }
}

/**
 * Notionデータベースから全ページを取得
 */
function fetchNotionDatabasePages(databaseId) {
  const notionToken = getSettingValue('NOTION_TOKEN');
  const allPages = [];
  let hasMore = true;
  let nextCursor = null;
  
  while (hasMore) {
    const url = `https://api.notion.com/v1/databases/${databaseId}/query`;
    const payload = {
      page_size: 100
    };
    
    if (nextCursor) {
      payload.start_cursor = nextCursor;
    }
    
    const options = {
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${notionToken}`,
        'Notion-Version': NOTION_VERSION,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };
    
    const response = UrlFetchApp.fetch(url, options);
    const code = response.getResponseCode();
    
    if (code !== 200) {
      throw new Error(`データベース取得エラー (Status: ${code}): ${response.getContentText()}`);
    }
    
    const data = JSON.parse(response.getContentText());
    allPages.push(...data.results);
    
    hasMore = data.has_more;
    nextCursor = data.next_cursor;
  }
  
  return allPages;
}

/**
 * Notionページからタイトルを抽出
 */
function extractNotionPageTitle(page) {
  if (page.properties) {
    for (const [key, value] of Object.entries(page.properties)) {
      if (value.type === 'title' && value.title && value.title.length > 0) {
        return value.title.map(t => t.plain_text).join('');
      }
    }
  }
  return 'No Title';
}

/**
 * Notionユーザー情報を取得
 */
function fetchNotionUser(userId) {
  const notionToken = getSettingValue('NOTION_TOKEN');
  const url = `https://api.notion.com/v1/users/${userId}`;
  const options = {
    method: 'get',
    headers: {
      'Authorization': `Bearer ${notionToken}`,
      'Notion-Version': NOTION_VERSION,
      'Content-Type': 'application/json'
    },
    muteHttpExceptions: true
  };
  const response = UrlFetchApp.fetch(url, options);
  const code = response.getResponseCode();
  if (code === 404) {
    // ユーザーが見つからない場合は null を返す（削除されたユーザーの可能性）
    return null;
  }
  if (code !== 200) {
    throw new Error(`Notion ユーザー取得エラー (Status: ${code})`);
  }
  return JSON.parse(response.getContentText());
}

/**
 * Notionユーザー別活動分析を更新
 */
function updateNotionUserAnalytics() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const pageListSheet = ss.getSheetByName(NOTION_ANALYSIS_SHEET_NAMES.PAGE_LIST);
  const userActivitySheet = ss.getSheetByName(NOTION_ANALYSIS_SHEET_NAMES.USER_ACTIVITY);
  const userAnalyticsSheet = ss.getSheetByName(NOTION_ANALYSIS_SHEET_NAMES.USER_ANALYTICS);
  const logSheet = ss.getSheetByName(NOTION_ANALYSIS_SHEET_NAMES.NOTION_LOG);
  
  // 現在の日時情報を取得
  const now = new Date();
  const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  const weekAgo = new Date(today.getTime() - 7 * 24 * 60 * 60 * 1000);
  const monthAgo = new Date(today.getTime() - 30 * 24 * 60 * 60 * 1000);
  
  const userStats = {};
  
  // ページリストからユーザー別の最新編集情報を取得
  const pageLastRow = pageListSheet.getLastRow();
  const pageData = pageLastRow > 1 ? pageListSheet.getRange(2, 1, pageLastRow - 1, 5).getValues() : [];
  
  pageData.forEach(([databaseName, pageId, title, lastEditedTime, editor]) => {
    if (!editor || editor === 'Unknown User') return;
    
    if (!userStats[editor]) {
      userStats[editor] = {
        name: editor,
        editedPages: new Set(),
        databases: new Set(),
        lastActivity: null,
        todayEdits: 0,
        weekEdits: 0,
        monthEdits: 0,
        totalEdits: 0,
        uniqueActivities: new Set()
      };
    }
    
    userStats[editor].editedPages.add(pageId);
    userStats[editor].databases.add(databaseName);
    
    if (lastEditedTime) {
      const editDate = new Date(lastEditedTime);
      if (!userStats[editor].lastActivity || editDate > userStats[editor].lastActivity) {
        userStats[editor].lastActivity = editDate;
      }
    }
  });
  
  // 活動履歴から期間別の編集回数を集計
  const activityLastRow = userActivitySheet.getLastRow();
  if (activityLastRow > 1) {
    const activityData = userActivitySheet.getRange(2, 1, activityLastRow - 1, 6).getValues();
    
    activityData.forEach(([recordTime, databaseName, userName, pageTitle, activityType, pageId]) => {
      if (!userName || userName === 'Unknown User') return;
      
      if (!userStats[userName]) {
        userStats[userName] = {
          name: userName,
          editedPages: new Set(),
          databases: new Set(),
          lastActivity: null,
          todayEdits: 0,
          weekEdits: 0,
          monthEdits: 0,
          totalEdits: 0,
          uniqueActivities: new Set()
        };
      }
      
      // recordTimeは最終編集日時として記録されている
      const activityDate = new Date(recordTime);
      
      // ページIDを編集ページとして記録
      if (pageId) {
        userStats[userName].editedPages.add(pageId);
      }
      
      // データベースを記録
      if (databaseName) {
        userStats[userName].databases.add(databaseName);
      }
      
      // 編集カウントを増やす
      userStats[userName].totalEdits++;
      
      // 期間別の編集数をカウント
      if (activityDate >= today) {
        userStats[userName].todayEdits++;
      }
      if (activityDate >= weekAgo) {
        userStats[userName].weekEdits++;
      }
      if (activityDate >= monthAgo) {
        userStats[userName].monthEdits++;
      }
      
      // 最終活動日を更新
      if (!userStats[userName].lastActivity || activityDate > userStats[userName].lastActivity) {
        userStats[userName].lastActivity = activityDate;
      }
    });
  }
  
  // Notion_Logシートからも編集履歴を取得
  if (logSheet) {
    const logLastRow = logSheet.getLastRow();
    if (logLastRow > 1) {
      const logData = logSheet.getRange(2, 1, logLastRow - 1, 7).getValues();
      
      logData.forEach(([recordTime, databaseName, pageId, pageTitle, lastEditedTime, editorName, changeType]) => {
        if (!editorName || editorName === 'Unknown User') return;
        
        if (!userStats[editorName]) {
          userStats[editorName] = {
            name: editorName,
            editedPages: new Set(),
            databases: new Set(),
            lastActivity: null,
            todayEdits: 0,
            weekEdits: 0,
            monthEdits: 0,
            totalEdits: 0,
            uniqueActivities: new Set()
          };
        }
        
        // 最終編集日時を使用（存在する場合）
        const editDate = lastEditedTime ? new Date(lastEditedTime) : new Date(recordTime);
        
        // ページとデータベースを記録
        if (pageId) {
          userStats[editorName].editedPages.add(pageId);
        }
        if (databaseName) {
          userStats[editorName].databases.add(databaseName);
        }
        
        // 重複を避けるためのキー
        const activityKey = `${editorName}_${pageId}_${editDate.getTime()}`;
        
        if (!userStats[editorName].uniqueActivities.has(activityKey)) {
          userStats[editorName].uniqueActivities.add(activityKey);
          userStats[editorName].totalEdits++;
          
          // 期間別カウント
          if (editDate >= today) {
            userStats[editorName].todayEdits++;
          }
          if (editDate >= weekAgo) {
            userStats[editorName].weekEdits++;
          }
          if (editDate >= monthAgo) {
            userStats[editorName].monthEdits++;
          }
          
          // 最終活動日を更新
          if (!userStats[editorName].lastActivity || editDate > userStats[editorName].lastActivity) {
            userStats[editorName].lastActivity = editDate;
          }
        }
      });
    }
  }
  
  // 分析結果をシートに書き込み
  const analyticsData = [];
  
  Object.values(userStats).forEach(user => {
    const editedPagesCount = user.editedPages.size;
    const databasesCount = user.databases.size;
    const lastActivityStr = user.lastActivity ? user.lastActivity.toLocaleString('ja-JP') : '';
    
    let avgEditInterval = 0;
    if (user.totalEdits > 1) {
      const sortedActivities = Array.from(user.uniqueActivities)
        .map(key => new Date(parseInt(key.split('_')[2])))
        .sort((a, b) => a - b);
      
      if (sortedActivities.length > 1) {
        const firstActivity = sortedActivities[0];
        const lastActivity = sortedActivities[sortedActivities.length - 1];
        const totalDays = (lastActivity - firstActivity) / (1000 * 60 * 60 * 24);
        avgEditInterval = Math.round((totalDays / (sortedActivities.length - 1)) * 10) / 10;
      }
    }
    
    analyticsData.push([
      user.name,
      editedPagesCount,
      databasesCount,
      lastActivityStr,
      user.todayEdits,
      user.weekEdits,
      user.monthEdits,
      user.totalEdits,
      avgEditInterval
    ]);
  });
  
  // 総編集回数でソート（降順）
  analyticsData.sort((a, b) => b[7] - a[7]);
  
  // 既存データをクリア
  const analyticsLastRow = userAnalyticsSheet.getLastRow();
  if (analyticsLastRow > 1) {
    userAnalyticsSheet.getRange(2, 1, analyticsLastRow - 1, 9).clearContent();
  }
  
  // 新しいデータを書き込み
  if (analyticsData.length > 0) {
    userAnalyticsSheet.getRange(2, 1, analyticsData.length, 9).setValues(analyticsData);
  }
  
  console.log(`Notionユーザー分析を更新しました: ${analyticsData.length} ユーザー`);
}

function fetchQUDENLicenses() {
  return [];
}

function fetchChatGPTLicenses() {
  return [];
}

function fetchSakuraServerLicenses() {
  return [];
}

// ===== ヘルパー関数 =====
function calculateUsage(lastLoginTime) {
  if (!lastLoginTime) return '未使用';
  
  const now = new Date();
  const lastLogin = new Date(lastLoginTime);
  const daysSinceLogin = Math.floor((now - lastLogin) / (1000 * 60 * 60 * 24));
  
  if (daysSinceLogin === 0) return '本日';
  if (daysSinceLogin === 1) return '昨日';
  if (daysSinceLogin <= 7) return `${daysSinceLogin}日前`;
  if (daysSinceLogin <= 30) return `${Math.floor(daysSinceLogin / 7)}週間前`;
  if (daysSinceLogin <= 365) return `${Math.floor(daysSinceLogin / 30)}ヶ月前`;
  
  return '1年以上前';
}

// ===== ユーザー管理シートを更新 =====
function updateUserManagementSheet(spreadsheet, licenses) {
  // 引数チェック
  if (!spreadsheet) {
    console.error('ユーザー管理シート更新エラー: スプレッドシートオブジェクトが渡されていません');
    return;
  }
  
  const sheet = spreadsheet.getSheetByName(CONFIG.SHEET_NAMES.USERS);
  if (!sheet) {
    console.error('ユーザー管理シートが見つかりません');
    return;
  }
  
  // ユーザーごとにライセンスを集計
  const userMap = new Map();
  
  licenses.forEach(license => {
    const email = license.email;
    if (!userMap.has(email)) {
      userMap.set(email, {
        email: email,
        name: license.name,
        department: license.department || '',
        role: license.role || '',
        platforms: new Set(),
        costs: {},
        totalCost: 0
      });
    }
    
    const user = userMap.get(email);
    user.platforms.add(license.platform);
    
    // プラットフォームごとのコストを記録
    if (!user.costs[license.platform]) {
      user.costs[license.platform] = 0;
    }
    user.costs[license.platform] += (license.cost || 0);
    user.totalCost += (license.cost || 0);
  });
  
  // データを配列に変換
  const userData = [];
  userMap.forEach(user => {
    const row = [
      user.email,
      user.name,
      user.department,
      user.role
    ];
    
    // 各プラットフォームの利用状況
    CONFIG.PLATFORMS.forEach(platform => {
      if (user.platforms.has(platform)) {
        row.push('◯');
      } else {
        row.push('');
      }
    });
    
    // 合計コストと最終更新日時
    row.push(user.totalCost);
    row.push(new Date());
    
    userData.push(row);
  });
  
  // 既存データをクリア（ヘッダー以外）
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
  }
  
  // データを書き込み
  if (userData.length > 0) {
    sheet.getRange(2, 1, userData.length, userData[0].length).setValues(userData);
    
    // 条件付き書式を設定（利用しているシステムをハイライト）
    const platformStartCol = 5; // プラットフォーム列の開始位置
    const platformEndCol = platformStartCol + CONFIG.PLATFORMS.length - 1;
    
    const range = sheet.getRange(2, platformStartCol, userData.length, CONFIG.PLATFORMS.length);
    range.setHorizontalAlignment('center');
    
    // ◯がある箇所に背景色を設定
    const rule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('◯')
      .setBackground('#c8e6c9')
      .setRanges([range])
      .build();
    
    const rules = sheet.getConditionalFormatRules();
    rules.push(rule);
    sheet.setConditionalFormatRules(rules);
  }
  
  console.log(`ユーザー管理シート更新: ${userData.length}名のユーザーを登録`);
}

// ===== 差分分析 =====
function performDifferenceAnalysis(spreadsheet, licenses) {
  // 引数チェック
  if (!spreadsheet) {
    console.error('差分分析エラー: スプレッドシートオブジェクトが渡されていません');
    return;
  }
  
  if (!licenses || !Array.isArray(licenses)) {
    console.error('差分分析エラー: ライセンスデータが正しく渡されていません');
    return;
  }
  
  const sheet = spreadsheet.getSheetByName(CONFIG.SHEET_NAMES.DIFF_ANALYSIS);
  if (!sheet) {
    console.error('差分分析シートが見つかりません');
    return;
  }
  
  // Google Workspaceユーザーを基準として取得
  const googleUsers = new Map();
  const platformUserMap = {};
  
  // 全プラットフォームの初期化
  CONFIG.PLATFORMS.forEach(platform => {
    platformUserMap[platform] = new Set();
  });
  
  // ライセンス情報を整理
  licenses.forEach(license => {
    const email = license.email.toLowerCase();
    
    if (license.platform === 'Google Workspace' && (license.status === 'Active' || license.status === 'アクティブ') && !license.suspended) {
      googleUsers.set(email, {
        email: license.email,
        name: license.name,
        googleLicense: license.licenseType
      });
    }
    
    if (license.status === 'Active' || license.status === 'アクティブ') {
      platformUserMap[license.platform].add(email);
    }
  });
  
  // 分析結果を格納
  const analysisResults = [];
  
  // 1. 全ユーザーのメールアドレスを収集
  const allEmails = new Set();
  licenses.forEach(license => {
          if (license.email && (license.status === 'Active' || license.status === 'アクティブ')) {
        allEmails.add(license.email.toLowerCase());
      }
  });
  
  // 2. Googleアカウントがないが、他のプラットフォームのアカウントがあるユーザーを検出
  allEmails.forEach(email => {
    // Googleアカウントがあるかチェック
    if (!googleUsers.has(email)) {
      // 他のプラットフォームでのライセンス状況を確認
      const userPlatforms = [];
      let userName = '';
      
      CONFIG.PLATFORMS.forEach(platform => {
        if (platform !== 'Google Workspace' && platformUserMap[platform].has(email)) {
          userPlatforms.push(platform);
          // ユーザー名を取得（最初に見つかったもの）
          if (!userName) {
            const userLicense = licenses.find(l => 
              l.email.toLowerCase() === email && 
              l.platform === platform
            );
            if (userLicense) {
              userName = userLicense.name || '';
            }
          }
        }
      });
      
      // 他のプラットフォームでライセンスを持っている場合のみ追加
      if (userPlatforms.length > 0) {
        const recommendation = 'アカウントの棚卸し';
        
        analysisResults.push([
          'Googleアカウントなし',
          email,
          userName,
          'なし',
          userPlatforms.length,
          1, // Googleアカウントが不足
          'Google Workspace',
          recommendation,
          userPlatforms.join(', ')
        ]);
      }
    }
  });
  
  // 分析結果をソート（Googleアカウントなしを優先）
  analysisResults.sort((a, b) => {
    if (a[0] === 'Googleアカウントなし' && b[0] !== 'Googleアカウントなし') return -1;
    if (a[0] !== 'Googleアカウントなし' && b[0] === 'Googleアカウントなし') return 1;
    return 0;
  });
  
  // 既存データをクリア（ヘッダー以外）
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
  }
  
  // 分析結果を書き込み
  if (analysisResults.length > 0) {
    sheet.getRange(2, 1, analysisResults.length, analysisResults[0].length).setValues(analysisResults);
  }
  
  // コンソールログに結果を出力
  console.log(`差分分析完了: Googleアカウントなしで他プラットフォーム利用中のユーザー ${
    analysisResults.filter(r => r[0] === 'Googleアカウントなし').length
  }名を検出`);
  
  // 分析結果を返す
  return analysisResults;
}

// ===== メールレポート作成・送信 =====
function sendEmailReport(licenses, analysisResults) {
  try {
    // レポート送信先のメールアドレスを取得
    const recipientEmail = getSettingValue('REPORT_EMAIL') || Session.getActiveUser().getEmail();
    
    // 非アクティブアカウントの分析（棚卸し期間を考慮）
    const accountAnalysis = analyzeInactiveAccountsWithInventory(licenses);
    const inactiveAccounts = accountAnalysis.inactive;
    
    // Notion利用分析データを取得
    const notionAnalytics = getNotionAnalyticsSummary();
    
    // Zoom利用分析データを取得
    const zoomAnalytics = getZoomAnalyticsSummary();
    
    // 低利用ユーザー分析データを取得
    const zoomLowUsage = analyzeZoomLowUsage();
    const notionLowUsage = analyzeNotionLowUsage();
    
    // HTMLメールの作成
    const htmlBody = createHtmlEmailReport(licenses, analysisResults, inactiveAccounts, notionAnalytics, zoomAnalytics, zoomLowUsage, notionLowUsage);
    
    // メール送信
    const subject = `ライセンス管理レポート - ${Utilities.formatDate(new Date(), 'JST', 'yyyy年MM月dd日')}`;
    
    GmailApp.sendEmail(
      recipientEmail,
      subject,
      '', // プレーンテキスト版（空文字）
      {
        htmlBody: htmlBody,
        name: 'ライセンス管理システム'
      }
    );
    
    console.log('レポートメールを送信しました: ' + recipientEmail);
    
  } catch (error) {
    console.error('メール送信エラー:', error);
  }
}

// ===== Notion利用分析サマリー取得 =====
function getNotionAnalyticsSummary() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // ユーザー分析シートからデータ取得
    let userAnalyticsSheet;
    try {
      userAnalyticsSheet = ss.getSheetByName(NOTION_ANALYSIS_SHEET_NAMES.USER_ANALYTICS);
    } catch (e) {
      return null;
    }
    
    const lastRow = userAnalyticsSheet.getLastRow();
    if (lastRow <= 1) return null;
    
    const userData = userAnalyticsSheet.getRange(2, 1, lastRow - 1, 9).getValues();
    
    // アクティブユーザー（今週編集があったユーザー）
    const activeUsers = userData.filter(row => row[5] > 0);
    const totalEdits = userData.reduce((sum, row) => sum + row[7], 0);
    const weeklyEdits = userData.reduce((sum, row) => sum + row[5], 0);
    
    // データベース設定シートから情報取得
    let dbConfigSheet;
    try {
      dbConfigSheet = ss.getSheetByName(NOTION_ANALYSIS_SHEET_NAMES.DATABASE_CONFIG);
      const dbLastRow = dbConfigSheet.getLastRow();
      const dbCount = dbLastRow > 1 ? dbConfigSheet.getRange(2, 3, dbLastRow - 1, 1).getValues().filter(row => row[0] === '有効').length : 0;
      
      return {
        totalUsers: userData.length,
        activeUsers: activeUsers.length,
        totalEdits: totalEdits,
        weeklyEdits: weeklyEdits,
        monitoredDatabases: dbCount,
        topEditors: userData.slice(0, 5).map(row => ({
          name: row[0],
          totalEdits: row[7],
          weeklyEdits: row[5]
        }))
      };
    } catch (e) {
      return {
        totalUsers: userData.length,
        activeUsers: activeUsers.length,
        totalEdits: totalEdits,
        weeklyEdits: weeklyEdits,
        monitoredDatabases: 0,
        topEditors: userData.slice(0, 5).map(row => ({
          name: row[0],
          totalEdits: row[7],
          weeklyEdits: row[5]
        }))
      };
    }
    
  } catch (error) {
    console.error('Notion分析サマリー取得エラー:', error);
    return null;
  }
}

// ===== Zoom利用分析サマリー取得 =====
function getZoomAnalyticsSummary() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // ユーザー分析シートからデータ取得
    let userAnalyticsSheet;
    try {
      userAnalyticsSheet = ss.getSheetByName(ZOOM_ANALYSIS_SHEET_NAMES.USER_ANALYTICS);
    } catch (e) {
      return null;
    }
    
    const lastRow = userAnalyticsSheet.getLastRow();
    if (lastRow <= 1) return null;
    
    const userData = userAnalyticsSheet.getRange(2, 1, lastRow - 1, 10).getValues();
    
    // ミーティングリストからデータ取得
    let meetingListSheet;
    try {
      meetingListSheet = ss.getSheetByName(ZOOM_ANALYSIS_SHEET_NAMES.MEETING_LIST);
      const meetingLastRow = meetingListSheet.getLastRow();
      const meetingCount = meetingLastRow > 1 ? meetingLastRow - 1 : 0;
      
      // 統計計算
      const totalMeetings = userData.reduce((sum, row) => sum + row[2], 0);
      const totalDuration = userData.reduce((sum, row) => sum + row[3], 0);
      const weeklyMeetings = userData.reduce((sum, row) => sum + row[6], 0);
      
      return {
        totalUsers: userData.length,
        totalMeetings: totalMeetings,
        totalDuration: totalDuration,
        weeklyMeetings: weeklyMeetings,
        avgMeetingDuration: totalMeetings > 0 ? Math.round(totalDuration / totalMeetings) : 0,
        topHosts: userData.slice(0, 5).map(row => ({
          name: row[0],
          email: row[1],
          totalMeetings: row[2],
          totalDuration: row[3]
        }))
      };
      
    } catch (e) {
      return null;
    }
    
  } catch (error) {
    console.error('Zoom分析サマリー取得エラー:', error);
    return null;
  }
}

// ===== 非アクティブアカウントの分析 =====
function analyzeInactiveAccounts(licenses) {
  const inactiveThresholdDays = 30;
  const currentDate = new Date();
  const thresholdDate = new Date(currentDate.getTime() - (inactiveThresholdDays * 24 * 60 * 60 * 1000));
  
  const inactiveByPlatform = {};
  
  // プラットフォームごとに初期化
  CONFIG.PLATFORMS.forEach(platform => {
    inactiveByPlatform[platform] = [];
  });
  
  // ライセンスごとに最終ログインをチェック
  licenses.forEach(license => {
    if (license.status === 'Active' && license.lastLogin) {
      const lastLoginDate = new Date(license.lastLogin);
      if (lastLoginDate < thresholdDate) {
        const daysSinceLogin = Math.floor((currentDate - lastLoginDate) / (1000 * 60 * 60 * 24));
        inactiveByPlatform[license.platform].push({
          email: license.email,
          name: license.name,
          lastLogin: license.lastLogin,
          daysSinceLogin: daysSinceLogin,
          cost: license.cost || 0
        });
      }
    }
  });
  
  return inactiveByPlatform;
}

// ===== Zoom利用度合い分析 =====
function analyzeZoomLowUsage() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const userAnalyticsSheet = ss.getSheetByName(ZOOM_ANALYSIS_SHEET_NAMES.USER_ANALYTICS);
    
    if (!userAnalyticsSheet) {
      return null;
    }
    
    const lastRow = userAnalyticsSheet.getLastRow();
    if (lastRow <= 1) {
      return null;
    }
    
    const userData = userAnalyticsSheet.getRange(2, 1, lastRow - 1, 10).getValues();
    const lowUsageUsers = [];
    
    // 利用度合いの低いユーザーの条件：
    // - 週のミーティング数が1以下
    // - または総ミーティング時間が30分以下
    userData.forEach(row => {
      const [name, email, totalMeetings, totalDuration, avgDuration, todayMeetings, weekMeetings, monthMeetings, lastMeeting, joinUrl] = row;
      
      if (weekMeetings <= 1 || totalDuration <= 30) {
        lowUsageUsers.push({
          name: name || 'Unknown',
          email: email || '',
          weekMeetings: weekMeetings || 0,
          totalDuration: totalDuration || 0,
          totalMeetings: totalMeetings || 0,
          avgDuration: avgDuration || 0,
          lastMeeting: lastMeeting || 'なし'
        });
      }
    });
    
    // 週のミーティング数でソート（昇順）
    lowUsageUsers.sort((a, b) => a.weekMeetings - b.weekMeetings);
    
    return lowUsageUsers;
  } catch (error) {
    console.error('Zoom低利用ユーザー分析エラー:', error);
    return null;
  }
}

// ===== Zoom利用状況下位レポートシート作成 =====
function createZoomLowUsageReport() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // 低利用ユーザーデータを取得
    const lowUsageUsers = analyzeZoomLowUsage();
    
    if (!lowUsageUsers || lowUsageUsers.length === 0) {
      console.log('Zoom低利用ユーザーが見つかりませんでした');
      return;
    }
    
    // レポートシートを取得または作成
    const reportSheet = getOrCreateSheet(ZOOM_ANALYSIS_SHEET_NAMES.LOW_USAGE_REPORT);
    
    // シートをクリア
    reportSheet.clear();
    
    // ヘッダー設定
    const headers = [
      ['Zoom利用状況下位レポート', '', '', '', '', '', ''],
      ['更新日時: ' + Utilities.formatDate(new Date(), 'JST', 'yyyy/MM/dd HH:mm:ss'), '', '', '', '', '', ''],
      ['', '', '', '', '', '', ''],
      ['条件: 週のミーティング数が1以下、または総利用時間が30分以下のユーザー', '', '', '', '', '', ''],
      ['', '', '', '', '', '', ''],
      ['ユーザー名', 'メールアドレス', '週のミーティング数', '総ミーティング数', '総利用時間（分）', '平均会議時間（分）', '最終ミーティング']
    ];
    
    // ヘッダーを設定
    reportSheet.getRange(1, 1, headers.length, 7).setValues(headers);
    
    // タイトル行のフォーマット
    reportSheet.getRange(1, 1, 1, 7).merge();
    reportSheet.getRange(1, 1).setFontSize(16).setFontWeight('bold');
    reportSheet.getRange(2, 1, 1, 7).merge();
    reportSheet.getRange(4, 1, 1, 7).merge();
    
    // ヘッダー行のフォーマット
    const headerRow = reportSheet.getRange(6, 1, 1, 7);
    headerRow.setBackground('#ea4335');
    headerRow.setFontColor('#ffffff');
    headerRow.setFontWeight('bold');
    
    // データを準備
    const data = lowUsageUsers.map(user => [
      user.name,
      user.email,
      user.weekMeetings,
      user.totalMeetings,
      user.totalDuration,
      user.avgDuration,
      user.lastMeeting
    ]);
    
    // データを書き込み
    if (data.length > 0) {
      const dataRange = reportSheet.getRange(7, 1, data.length, 7);
      dataRange.setValues(data);
      
      // 週のミーティング数が0のユーザーを赤色で強調
      for (let i = 0; i < data.length; i++) {
        if (data[i][2] === 0) { // 週のミーティング数が0
          reportSheet.getRange(7 + i, 1, 1, 7).setBackground('#ffcccc');
        } else if (data[i][4] <= 30) { // 総利用時間が30分以下
          reportSheet.getRange(7 + i, 1, 1, 7).setBackground('#fff3cd');
        }
      }
    }
    
    // 統計情報を追加
    const statsStartRow = 7 + data.length + 2;
    const stats = [
      ['統計情報', ''],
      ['低利用ユーザー数', lowUsageUsers.length + '名'],
      ['週のミーティング数が0のユーザー', lowUsageUsers.filter(u => u.weekMeetings === 0).length + '名'],
      ['週のミーティング数が1以下のユーザー', lowUsageUsers.filter(u => u.weekMeetings <= 1).length + '名'],
      ['総利用時間が30分以下のユーザー', lowUsageUsers.filter(u => u.totalDuration <= 30).length + '名']
    ];
    
    reportSheet.getRange(statsStartRow, 1, stats.length, 2).setValues(stats);
    reportSheet.getRange(statsStartRow, 1).setFontWeight('bold').setFontSize(14);
    
    // 列幅を調整
    reportSheet.setColumnWidth(1, 200); // ユーザー名
    reportSheet.setColumnWidth(2, 250); // メールアドレス
    reportSheet.setColumnWidth(3, 150); // 週のミーティング数
    reportSheet.setColumnWidth(4, 150); // 総ミーティング数
    reportSheet.setColumnWidth(5, 150); // 総利用時間
    reportSheet.setColumnWidth(6, 150); // 平均会議時間
    reportSheet.setColumnWidth(7, 150); // 最終ミーティング
    
    console.log('Zoom利用状況下位レポートを作成しました');
    
  } catch (error) {
    console.error('Zoom利用状況下位レポート作成エラー:', error);
  }
}

// ===== Notion利用度合い分析 =====
function analyzeNotionLowUsage() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const userAnalyticsSheet = ss.getSheetByName(NOTION_ANALYSIS_SHEET_NAMES.USER_ANALYTICS);
    
    if (!userAnalyticsSheet) {
      return null;
    }
    
    const lastRow = userAnalyticsSheet.getLastRow();
    if (lastRow <= 1) {
      return null;
    }
    
    const userData = userAnalyticsSheet.getRange(2, 1, lastRow - 1, 9).getValues();
    const lowUsageUsers = [];
    
    // 利用度合いの低いユーザーの条件：
    // - 週の編集数が0
    // - または月の編集数が3以下
    userData.forEach(row => {
      const [name, editedPages, databases, lastActivity, todayEdits, weekEdits, monthEdits, totalEdits, avgInterval] = row;
      
      if (weekEdits === 0 || monthEdits <= 3) {
        lowUsageUsers.push({
          name: name || 'Unknown',
          weekEdits: weekEdits || 0,
          monthEdits: monthEdits || 0,
          totalEdits: totalEdits || 0,
          editedPages: editedPages || 0,
          lastActivity: lastActivity || 'なし',
          avgInterval: avgInterval || 0
        });
      }
    });
    
    // 週の編集数でソート（昇順）
    lowUsageUsers.sort((a, b) => a.weekEdits - b.weekEdits);
    
    return lowUsageUsers;
  } catch (error) {
    console.error('Notion低利用ユーザー分析エラー:', error);
    return null;
  }
}

// ===== Notion利用状況下位レポートシート作成 =====
function createNotionLowUsageReport() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // 低利用ユーザーデータを取得
    const lowUsageUsers = analyzeNotionLowUsage();
    
    if (!lowUsageUsers || lowUsageUsers.length === 0) {
      console.log('Notion低利用ユーザーが見つかりませんでした');
      return;
    }
    
    // レポートシートを取得または作成
    const reportSheet = getOrCreateSheet(NOTION_ANALYSIS_SHEET_NAMES.LOW_USAGE_REPORT);
    
    // シートをクリア
    reportSheet.clear();
    
    // ヘッダー設定
    const headers = [
      ['Notion利用状況下位レポート', '', '', '', '', '', ''],
      ['更新日時: ' + Utilities.formatDate(new Date(), 'JST', 'yyyy/MM/dd HH:mm:ss'), '', '', '', '', '', ''],
      ['', '', '', '', '', '', ''],
      ['条件: 週の編集数が0、または月の編集数が3以下のユーザー', '', '', '', '', '', ''],
      ['', '', '', '', '', '', ''],
      ['ユーザー名', '週の編集数', '月の編集数', '総編集数', '編集ページ数', '最終活動日時', '平均編集間隔']
    ];
    
    // ヘッダーを設定
    reportSheet.getRange(1, 1, headers.length, 7).setValues(headers);
    
    // タイトル行のフォーマット
    reportSheet.getRange(1, 1, 1, 7).merge();
    reportSheet.getRange(1, 1).setFontSize(16).setFontWeight('bold');
    reportSheet.getRange(2, 1, 1, 7).merge();
    reportSheet.getRange(4, 1, 1, 7).merge();
    
    // ヘッダー行のフォーマット
    const headerRow = reportSheet.getRange(6, 1, 1, 7);
    headerRow.setBackground('#4285f4');
    headerRow.setFontColor('#ffffff');
    headerRow.setFontWeight('bold');
    
    // データを準備
    const data = lowUsageUsers.map(user => [
      user.name,
      user.weekEdits,
      user.monthEdits,
      user.totalEdits,
      user.editedPages,
      user.lastActivity,
      user.avgInterval ? user.avgInterval + '日' : '-'
    ]);
    
    // データを書き込み
    if (data.length > 0) {
      const dataRange = reportSheet.getRange(7, 1, data.length, 7);
      dataRange.setValues(data);
      
      // 週の編集数が0のユーザーを赤色で強調
      for (let i = 0; i < data.length; i++) {
        if (data[i][1] === 0) { // 週の編集数が0
          reportSheet.getRange(7 + i, 1, 1, 7).setBackground('#ffcccc');
        } else if (data[i][2] <= 3) { // 月の編集数が3以下
          reportSheet.getRange(7 + i, 1, 1, 7).setBackground('#fff3cd');
        }
      }
    }
    
    // 統計情報を追加
    const statsStartRow = 7 + data.length + 2;
    const stats = [
      ['統計情報', ''],
      ['低利用ユーザー数', lowUsageUsers.length + '名'],
      ['週の編集数が0のユーザー', lowUsageUsers.filter(u => u.weekEdits === 0).length + '名'],
      ['月の編集数が3以下のユーザー', lowUsageUsers.filter(u => u.monthEdits <= 3).length + '名']
    ];
    
    reportSheet.getRange(statsStartRow, 1, stats.length, 2).setValues(stats);
    reportSheet.getRange(statsStartRow, 1).setFontWeight('bold').setFontSize(14);
    
    // 列幅を調整
    reportSheet.setColumnWidth(1, 200); // ユーザー名
    reportSheet.setColumnWidth(2, 100); // 週の編集数
    reportSheet.setColumnWidth(3, 100); // 月の編集数
    reportSheet.setColumnWidth(4, 100); // 総編集数
    reportSheet.setColumnWidth(5, 120); // 編集ページ数
    reportSheet.setColumnWidth(6, 150); // 最終活動日時
    reportSheet.setColumnWidth(7, 120); // 平均編集間隔
    
    console.log('Notion利用状況下位レポートを作成しました');
    
  } catch (error) {
    console.error('Notion利用状況下位レポート作成エラー:', error);
  }
}

// ===== HTMLメールレポートの作成 =====
function createHtmlEmailReport(licenses, analysisResults, inactiveAccounts, notionAnalytics, zoomAnalytics, zoomLowUsage, notionLowUsage) {
  let html = `
    <!DOCTYPE html>
    <html>
    <head>
      <style>
        body {
          font-family: Arial, sans-serif;
          line-height: 1.6;
          color: #333;
          max-width: 1200px;
          margin: 0 auto;
          padding: 20px;
        }
        h1 {
          color: #2c3e50;
          border-bottom: 3px solid #3498db;
          padding-bottom: 10px;
        }
        h2 {
          color: #34495e;
          margin-top: 30px;
          border-bottom: 2px solid #95a5a6;
          padding-bottom: 5px;
        }
        h3 {
          color: #7f8c8d;
          margin-top: 20px;
        }
        table {
          border-collapse: collapse;
          width: 100%;
          margin: 15px 0;
          box-shadow: 0 2px 5px rgba(0,0,0,0.1);
        }
        th {
          background-color: #3498db;
          color: white;
          padding: 12px;
          text-align: left;
          font-weight: bold;
        }
        td {
          padding: 10px;
          border-bottom: 1px solid #ecf0f1;
        }
        tr:nth-child(even) {
          background-color: #f8f9fa;
        }
        tr:hover {
          background-color: #e8f4fd;
        }
        .summary-box {
          background-color: #ecf0f1;
          border-left: 5px solid #3498db;
          padding: 15px;
          margin: 20px 0;
          border-radius: 5px;
        }
        .warning {
          color: #e74c3c;
          font-weight: bold;
        }
        .success {
          color: #27ae60;
          font-weight: bold;
        }
        .metric {
          display: inline-block;
          margin: 10px 20px 10px 0;
        }
        .metric-value {
          font-size: 24px;
          font-weight: bold;
          color: #2c3e50;
        }
        .metric-label {
          font-size: 14px;
          color: #7f8c8d;
        }
        .platform-section {
          margin: 20px 0;
          padding: 15px;
          background-color: #f8f9fa;
          border-radius: 8px;
        }
        .no-data {
          color: #95a5a6;
          font-style: italic;
        }
      </style>
    </head>
    <body>
      <h1>ライセンス管理レポート</h1>
      <p>レポート生成日: ${Utilities.formatDate(new Date(), 'JST', 'yyyy年MM月dd日 HH:mm')}</p>
      <p><a href="https://docs.google.com/spreadsheets/d/11OuDsWRUfWP_Bighhhjk6VocnMcn9YAmvgGifdjGEys/edit?pli=1&gid=1290738652#gid=1290738652" 
         style="color: #3498db; text-decoration: none; font-weight: bold;">
         &gt;&gt; スプレッドシートで詳細を確認する
      </a></p>
      
      <div class="summary-box">
        <h2>サマリー</h2>
        <div class="metric">
          <div class="metric-value">${licenses.length}</div>
          <div class="metric-label">総ライセンス数</div>
        </div>
        <div class="metric">
          <div class="metric-value">${licenses.filter(l => l.status === 'Active').length}</div>
          <div class="metric-label">アクティブライセンス</div>
        </div>
        <div class="metric">
          <div class="metric-value">¥${licenses.reduce((sum, l) => sum + (l.cost || 0), 0).toLocaleString()}</div>
          <div class="metric-label">月額合計コスト</div>
        </div>
      </div>
  `;
  
  // 非アクティブアカウント
  html += `<h2>30日以上ログインのないアカウント</h2>`;
  
  let hasInactiveAccounts = false;
  CONFIG.PLATFORMS.forEach(platform => {
    if (inactiveAccounts[platform] && inactiveAccounts[platform].length > 0) {
      hasInactiveAccounts = true;
      const totalInactiveCost = inactiveAccounts[platform].reduce((sum, account) => sum + account.cost, 0);
      
      html += `
        <div class="platform-section">
          <h3>${platform}</h3>
          <p>非アクティブアカウント数: <span class="warning">${inactiveAccounts[platform].length}</span> | 
             月額コスト: <span class="warning">¥${totalInactiveCost.toLocaleString()}</span></p>
          <table>
            <thead>
              <tr>
                <th>メールアドレス</th>
                <th>氏名</th>
                <th>最終ログイン</th>
                <th>経過日数</th>
                <th>月額コスト</th>
              </tr>
            </thead>
            <tbody>
      `;
      
      inactiveAccounts[platform].forEach(account => {
        html += `
          <tr>
            <td>${account.email}</td>
            <td>${account.name}</td>
            <td>${Utilities.formatDate(new Date(account.lastLogin), 'JST', 'yyyy/MM/dd')}</td>
            <td class="warning">${account.daysSinceLogin}日前</td>
            <td>¥${account.cost.toLocaleString()}</td>
          </tr>
        `;
      });
      
      html += `
            </tbody>
          </table>
        </div>
      `;
    }
  });
  
  if (!hasInactiveAccounts) {
    html += `<p class="no-data">非アクティブアカウントはありません</p>`;
  }
  
  // Notion利用分析
  if (notionAnalytics) {
    html += `
      <h2>Notion利用分析</h2>
      <div class="platform-section">
        <div class="metric">
          <div class="metric-value">${notionAnalytics.totalUsers}</div>
          <div class="metric-label">総ユーザー数</div>
        </div>
        <div class="metric">
          <div class="metric-value">${notionAnalytics.activeUsers}</div>
          <div class="metric-label">今週アクティブ</div>
        </div>
        <div class="metric">
          <div class="metric-value">${notionAnalytics.weeklyEdits.toLocaleString()}</div>
          <div class="metric-label">今週の編集数</div>
        </div>
        <div class="metric">
          <div class="metric-value">${notionAnalytics.monitoredDatabases}</div>
          <div class="metric-label">監視中DB数</div>
        </div>
      </div>
      
      <h3>最もアクティブなユーザー（TOP5）</h3>
      <table>
        <thead>
          <tr>
            <th>ユーザー名</th>
            <th>総編集回数</th>
            <th>今週の編集数</th>
          </tr>
        </thead>
        <tbody>
    `;
    
    notionAnalytics.topEditors.forEach(editor => {
      html += `
        <tr>
          <td>${editor.name}</td>
          <td>${editor.totalEdits.toLocaleString()}</td>
          <td>${editor.weeklyEdits.toLocaleString()}</td>
        </tr>
      `;
    });
    
    html += `
        </tbody>
      </table>
    `;
  }
  
  // Notion低利用ユーザー分析
  if (notionLowUsage && notionLowUsage.length > 0) {
    html += `
      <h3>Notion利用度が低いユーザー</h3>
      <p>週の編集数が0、または月の編集数が3以下のユーザー:</p>
      <table>
        <thead>
          <tr>
            <th>名前</th>
            <th>週の編集数</th>
            <th>月の編集数</th>
            <th>総編集数</th>
            <th>編集ページ数</th>
            <th>最終活動</th>
          </tr>
        </thead>
        <tbody>
    `;
    
    notionLowUsage.forEach(user => {
      html += `
        <tr>
          <td>${user.name}</td>
          <td class="warning">${user.weekEdits}</td>
          <td>${user.monthEdits}</td>
          <td>${user.totalEdits}</td>
          <td>${user.editedPages}</td>
          <td>${user.lastActivity}</td>
        </tr>
      `;
    });
    
    html += `
        </tbody>
      </table>
    `;
  }
  
  // Zoom利用分析
  if (zoomAnalytics) {
    html += `
      <h2>Zoom利用分析</h2>
      <div class="platform-section">
        <div class="metric">
          <div class="metric-value">${zoomAnalytics.totalUsers}</div>
          <div class="metric-label">総ユーザー数</div>
        </div>
        <div class="metric">
          <div class="metric-value">${zoomAnalytics.weeklyMeetings.toLocaleString()}</div>
          <div class="metric-label">今週のミーティング</div>
        </div>
        <div class="metric">
          <div class="metric-value">${zoomAnalytics.totalDuration.toLocaleString()}</div>
          <div class="metric-label">総利用時間（分）</div>
        </div>
        <div class="metric">
          <div class="metric-value">${zoomAnalytics.avgMeetingDuration}</div>
          <div class="metric-label">平均会議時間（分）</div>
        </div>
      </div>
      
      <h3>最も活発なホスト（TOP5）</h3>
      <table>
        <thead>
          <tr>
            <th>ホスト名</th>
            <th>メールアドレス</th>
            <th>総ミーティング数</th>
            <th>総利用時間（分）</th>
          </tr>
        </thead>
        <tbody>
    `;
    
    zoomAnalytics.topHosts.forEach(host => {
      html += `
        <tr>
          <td>${host.name}</td>
          <td>${host.email}</td>
          <td>${host.totalMeetings.toLocaleString()}</td>
          <td>${host.totalDuration.toLocaleString()}</td>
        </tr>
      `;
    });
    
    html += `
        </tbody>
      </table>
    `;
  }
  
  // Zoom低利用ユーザー分析
  if (zoomLowUsage && zoomLowUsage.length > 0) {
    html += `
      <h3>Zoom利用度が低いユーザー</h3>
      <p>週のミーティング数が1以下、または総利用時間が30分以下のユーザー:</p>
      <table>
        <thead>
          <tr>
            <th>名前</th>
            <th>メールアドレス</th>
            <th>週のミーティング数</th>
            <th>総ミーティング数</th>
            <th>総利用時間（分）</th>
            <th>最終ミーティング</th>
          </tr>
        </thead>
        <tbody>
    `;
    
    zoomLowUsage.forEach(user => {
      html += `
        <tr>
          <td>${user.name}</td>
          <td>${user.email}</td>
          <td class="warning">${user.weekMeetings}</td>
          <td>${user.totalMeetings}</td>
          <td>${user.totalDuration}</td>
          <td>${user.lastMeeting}</td>
        </tr>
      `;
    });
    
    html += `
        </tbody>
      </table>
    `;
  }
  
  // 差分分析結果（最後に配置）
  if (analysisResults && analysisResults.length > 0) {
    const googleAccountlessUsers = analysisResults.filter(r => r[0] === 'Googleアカウントなし');
    if (googleAccountlessUsers.length > 0) {
      html += `
        <h2>差分分析結果</h2>
        <p>Google Workspaceアカウントがないが、他のプラットフォームを利用しているユーザー:</p>
        <table>
          <thead>
            <tr>
              <th>メールアドレス</th>
              <th>氏名</th>
              <th>利用中のプラットフォーム</th>
              <th>推奨アクション</th>
            </tr>
          </thead>
          <tbody>
      `;
      
      googleAccountlessUsers.forEach(result => {
        html += `
          <tr>
            <td>${result[1]}</td>
            <td>${result[2]}</td>
            <td>${result[8]}</td>
            <td class="warning">${result[7]}</td>
          </tr>
        `;
      });
      
      html += `
          </tbody>
        </table>
      `;
    }
  }
  

  
  html += `
    </body>
    </html>
  `;
  
  return html;
}

// ===== Notion/Zoom分析実行関数 =====

/**
 * Notion利用分析を実行（メニューから実行）
 */
function runNotionAnalysis() {
  try {
    console.log('Notion利用分析を開始します');
    
    // Notion分析機能のセットアップ（既にセットアップ済みの場合はスキップ）
    setupNotionAnalytics();
    
    // データベースを自動検出
    const databases = discoverNotionDatabases();
    
    if (databases.length === 0) {
      console.error('Notionデータベースが見つかりませんでした。NotionのAPIトークンが正しく設定されているか確認してください。');
      return;
    }
    
    // データベース設定をシートに保存
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const configSheet = ss.getSheetByName(NOTION_ANALYSIS_SHEET_NAMES.DATABASE_CONFIG);
    
    if (!configSheet) {
      console.error('エラー: データベース設定シートが作成できませんでした。');
      return;
    }
    
    // 既存データをクリア
    const lastRow = configSheet.getLastRow();
    if (lastRow > 1) {
      configSheet.getRange(2, 1, lastRow - 1, 6).clearContent();
    }
    
    // データベース情報を書き込み
    const configData = databases.map(db => [
      db.id,
      db.name,
      '有効',
      '',
      db.created || '',
      db.url || ''
    ]);
    
    if (configData.length > 0) {
      configSheet.getRange(2, 1, configData.length, 6).setValues(configData);
    }
    
    // 全ページを取得
    const pageListSheet = ss.getSheetByName(NOTION_ANALYSIS_SHEET_NAMES.PAGE_LIST);
    
    if (!pageListSheet) {
      console.error('エラー: ページリストシートが作成できませんでした。');
      return;
    }
    
    const allData = [];
    
    for (const db of databases) {
      if (db.enabled) {
        try {
          const pages = fetchNotionDatabasePages(db.id);
          
          for (const page of pages) {
            const pageId = page.id;
            const title = extractNotionPageTitle(page);
            const lastEditedTime = page.last_edited_time || '';
            const createdTime = page.created_time || '';
            
            // 編集者情報を取得（last_edited_byとcreated_byの両方を確認）
            let editorName = '';
            let editorEmail = '';
            
            // 最終編集者を優先的に取得
            if (page.last_edited_by && page.last_edited_by.id) {
              try {
                const userInfo = fetchNotionUser(page.last_edited_by.id);
                editorName = userInfo.name || userInfo.type || 'Unknown User';
                editorEmail = userInfo.person && userInfo.person.email || '';
              } catch (e) {
                editorName = 'Unknown User';
                console.log('最終編集者の取得に失敗:', e);
              }
            }
            
            // 最終編集者が取得できない場合は作成者を取得
            if (!editorName || editorName === 'Unknown User') {
              if (page.created_by && page.created_by.id) {
                try {
                  const userInfo = fetchNotionUser(page.created_by.id);
                  editorName = userInfo.name || userInfo.type || 'Unknown User';
                  editorEmail = userInfo.person && userInfo.person.email || '';
                } catch (e) {
                  editorName = 'Unknown User';
                  console.log('作成者の取得に失敗:', e);
                }
              }
            }
            
            // ユーザーアクティビティを記録
            if (editorName && editorName !== 'Unknown User') {
              const activitySheet = ss.getSheetByName(NOTION_ANALYSIS_SHEET_NAMES.USER_ACTIVITY);
              if (activitySheet) {
                const activityData = [
                  new Date(), // 記録時刻
                  db.name,    // データベース名
                  editorName, // ユーザー名
                  title,      // ページタイトル
                  'Page Edit', // アクティビティタイプ
                  pageId      // ページID
                ];
                activitySheet.appendRow(activityData);
              }
            }
            
            allData.push([db.name, pageId, title, lastEditedTime, editorName]);
          }
        } catch (error) {
          console.error(`データベース ${db.name} でエラー:`, error);
        }
      }
    }
    
    // データをシートに書き込み
    if (allData.length > 0) {
      const pageLastRow = pageListSheet.getLastRow();
      if (pageLastRow > 1) {
        pageListSheet.getRange(2, 1, pageLastRow - 1, 5).clearContent();
      }
      pageListSheet.getRange(2, 1, allData.length, 5).setValues(allData);
    }
    
    // ユーザー分析を更新
    updateNotionUserAnalytics();
    
    console.log(`Notion利用分析が完了しました。検出されたデータベース: ${databases.length}、取得されたページ: ${allData.length}`);
    
  } catch (error) {
    console.error('Notion分析エラー:', error);

  }
}

/**
 * Zoom利用分析を実行（メニューから実行）
 */
function runZoomAnalysis() {
  try {
    console.log('Zoom利用分析を開始します');
    
    // Zoom分析機能のセットアップ（既にセットアップ済みの場合はスキップ）
    setupZoomAnalytics();
    
    // ミーティング履歴を取得
    fetchZoomMeetings();
    
    console.log('Zoom利用分析が完了しました');
    
  } catch (error) {
    console.error('Zoom分析エラー:', error);
  }
}

/**
 * 全プラットフォーム分析を実行（メニューから実行）
 */
function runAllAnalytics() {
  try {
    console.log('全プラットフォームの利用分析を開始します');
    
    // Notion分析
    console.log('Notion分析を開始...');
    runNotionAnalysis();
    
    // Zoom分析
    console.log('Zoom分析を開始...');
    runZoomAnalysis();
    
    console.log('全プラットフォームの利用分析が完了しました');
    
  } catch (error) {
    console.error('全プラットフォーム分析エラー:', error);
  }
}

// ===== デバッグ・診断機能 =====

/**
 * Notion分析機能の診断
 */
function diagnoseNotionAnalysis() {
  const ui = SpreadsheetApp.getUi();
  let report = '=== Notion分析機能診断 ===\n\n';
  
  // 1. APIトークンの確認
  const notionToken = getSettingValue('NOTION_TOKEN');
  if (!notionToken) {
    report += '❌ NotionのAPIトークンが設定されていません\n';
    report += '→ 設定画面からNotionトークンを設定してください\n\n';
  } else {
    report += '✅ NotionのAPIトークンが設定されています\n';
    report += 'トークン: ' + notionToken.substring(0, 10) + '...' + notionToken.substring(notionToken.length - 5) + '\n\n';
  }
  
  // 2. API接続テスト
  if (notionToken) {
    try {
      const url = 'https://api.notion.com/v1/users/me';
      const options = {
        method: 'GET',
        headers: {
          'Authorization': 'Bearer ' + notionToken,
          'Notion-Version': NOTION_VERSION,
          'Content-Type': 'application/json'
        },
        muteHttpExceptions: true
      };
      
      const response = UrlFetchApp.fetch(url, options);
      const code = response.getResponseCode();
      
      if (code === 200) {
        const data = JSON.parse(response.getContentText());
        report += '✅ Notion API接続成功\n';
        report += 'ボット名: ' + (data.bot.name || 'Unknown') + '\n';
        report += 'ワークスペース: ' + (data.bot.workspace_name || 'Unknown') + '\n\n';
      } else {
        report += '❌ Notion API接続失敗\n';
        report += 'ステータスコード: ' + code + '\n';
        report += 'エラー: ' + response.getContentText() + '\n\n';
      }
    } catch (error) {
      report += '❌ Notion API接続エラー\n';
      report += 'エラー: ' + error.message + '\n\n';
    }
  }
  
  // 3. シートの存在確認
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const requiredSheets = Object.values(NOTION_ANALYSIS_SHEET_NAMES);
  
  report += '=== シート状態 ===\n';
  requiredSheets.forEach(sheetName => {
    try {
      const sheet = ss.getSheetByName(sheetName);
      if (sheet) {
        const lastRow = sheet.getLastRow();
        report += '✅ ' + sheetName + ': ' + lastRow + ' 行\n';
      } else {
        report += '❌ ' + sheetName + ': 未作成\n';
      }
    } catch (e) {
      report += '❌ ' + sheetName + ': エラー\n';
    }
  });
  
  ui.alert('Notion分析機能診断結果', report, ui.ButtonSet.OK);
}

/**
 * Zoom分析機能の診断
 */
function diagnoseZoomAnalysis() {
  const ui = SpreadsheetApp.getUi();
  let report = '=== Zoom分析機能診断 ===\n\n';
  
  // 1. 認証情報の確認
  const accountId = getSettingValue('ZOOM_ACCOUNT_ID');
  const clientId = getSettingValue('ZOOM_CLIENT_ID');
  const clientSecret = getSettingValue('ZOOM_CLIENT_SECRET');
  
  if (!accountId || !clientId || !clientSecret) {
    report += '❌ Zoom認証情報が不完全です\n';
    if (!accountId) report += '→ ZOOM_ACCOUNT_IDが未設定\n';
    if (!clientId) report += '→ ZOOM_CLIENT_IDが未設定\n';
    if (!clientSecret) report += '→ ZOOM_CLIENT_SECRETが未設定\n';
    report += '\n設定画面から認証情報を設定してください\n\n';
  } else {
    report += '✅ Zoom認証情報が設定されています\n';
    report += 'アカウントID: ' + accountId.substring(0, 5) + '...' + accountId.substring(accountId.length - 3) + '\n';
    report += 'クライアントID: ' + clientId.substring(0, 5) + '...' + clientId.substring(clientId.length - 3) + '\n\n';
  }
  
  // 2. アクセストークン取得テスト
  if (accountId && clientId && clientSecret) {
    try {
      const accessToken = getZoomAccessToken();
      if (accessToken) {
        report += '✅ アクセストークン取得成功\n';
        report += 'トークン: ' + accessToken.substring(0, 10) + '...' + accessToken.substring(accessToken.length - 5) + '\n\n';
        
        // 3. API接続テスト
        try {
          const url = 'https://api.zoom.us/v2/users/me';
          const options = {
            method: 'GET',
            headers: {
              'Authorization': 'Bearer ' + accessToken,
              'Content-Type': 'application/json'
            },
            muteHttpExceptions: true
          };
          
          const response = UrlFetchApp.fetch(url, options);
          const code = response.getResponseCode();
          
          if (code === 200) {
            const data = JSON.parse(response.getContentText());
            report += '✅ Zoom API接続成功\n';
            report += 'ユーザー: ' + (data.email || 'Unknown') + '\n';
            report += 'アカウントID: ' + (data.account_id || 'Unknown') + '\n\n';
          } else {
            report += '❌ Zoom API接続失敗\n';
            report += 'ステータスコード: ' + code + '\n';
            report += 'エラー: ' + response.getContentText() + '\n\n';
          }
        } catch (error) {
          report += '❌ Zoom API接続エラー\n';
          report += 'エラー: ' + error.message + '\n\n';
        }
      } else {
        report += '❌ アクセストークン取得失敗\n\n';
      }
    } catch (error) {
      report += '❌ アクセストークン取得エラー\n';
      report += 'エラー: ' + error.message + '\n\n';
    }
  }
  
  // 4. シートの存在確認
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const requiredSheets = Object.values(ZOOM_ANALYSIS_SHEET_NAMES);
  
  report += '=== シート状態 ===\n';
  requiredSheets.forEach(sheetName => {
    try {
      const sheet = ss.getSheetByName(sheetName);
      if (sheet) {
        const lastRow = sheet.getLastRow();
        report += '✅ ' + sheetName + ': ' + lastRow + ' 行\n';
      } else {
        report += '❌ ' + sheetName + ': 未作成\n';
      }
    } catch (e) {
      report += '❌ ' + sheetName + ': エラー\n';
    }
  });
  
  ui.alert('Zoom分析機能診断結果', report, ui.ButtonSet.OK);
}

/**
 * 全分析機能の診断
 */
function diagnoseAllAnalytics() {
  diagnoseNotionAnalysis();
  diagnoseZoomAnalysis();
}

// ===== 設定画面表示 =====
function showSettings() {
  try {
    const html = HtmlService.createHtmlOutputFromFile('settings')
      .setWidth(500)
      .setHeight(700);
    SpreadsheetApp.getUi().showModalDialog(html, '設定');
  } catch (error) {
    console.error('設定画面表示エラー:', error);
    SpreadsheetApp.getUi().alert('エラー', '設定画面の表示に失敗しました: ' + error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

// ===== 基本設定を保存 =====
function saveBasicSettings(settings) {
  const scriptProperties = PropertiesService.getScriptProperties();
  
  Object.entries(settings).forEach(([key, value]) => {
    if (value && value.toString().trim() !== '') {
      scriptProperties.setProperty(key, value.toString());
    }
  });
  
  return true;
}

// ===== 全設定を保存（新しい設定画面用） =====
function saveAllSettings(settings) {
  try {
    const scriptProperties = PropertiesService.getScriptProperties();
    
    // デバッグ用ログ
    console.log('保存する設定:', settings);
    
    let savedCount = 0;
    Object.entries(settings).forEach(([key, value]) => {
      if (value && value.toString().trim() !== '') {
        scriptProperties.setProperty(key, value.toString());
        savedCount++;
        console.log(`保存: ${key} = ${value}`);
      }
    });
    
    console.log(`${savedCount}個の設定を保存しました`);
    
    // 保存確認
    const testValue = scriptProperties.getProperty('GOOGLE_DOMAIN');
    console.log('保存確認 - GOOGLE_DOMAIN:', testValue);
    
    return true;
  } catch (error) {
    console.error('設定保存エラー:', error);
    throw error;
  }
}

// ===== 基本設定を取得 =====
function getBasicSettings() {
  const scriptProperties = PropertiesService.getScriptProperties();
  return {
    GOOGLE_DOMAIN: scriptProperties.getProperty('GOOGLE_DOMAIN') || '',
    GOOGLE_WORKSPACE_PLAN: scriptProperties.getProperty('GOOGLE_WORKSPACE_PLAN') || 'Business Standard',
    NOTIFICATION_EMAIL: scriptProperties.getProperty('NOTIFICATION_EMAIL') || '',
    TRUSTLOGIN_API_TOKEN: scriptProperties.getProperty('TRUSTLOGIN_API_TOKEN') || '',
    ACTIVEGATE_DOMAIN: scriptProperties.getProperty('ACTIVEGATE_DOMAIN') || '',
    ACTIVEGATE_API_KEY: scriptProperties.getProperty('ACTIVEGATE_API_KEY') || ''
  };
}

// ===== 全設定を取得（新しい設定画面用） =====
function getAllSettings() {
  try {
    const scriptProperties = PropertiesService.getScriptProperties();
    console.log('設定を読み込み中...');
    
    // すべてのプロパティを取得してログ出力
    const allProps = scriptProperties.getProperties();
    console.log('保存されているすべての設定:', allProps);
    
    return {
      // 基本設定
      GOOGLE_DOMAIN: scriptProperties.getProperty('GOOGLE_DOMAIN') || '',
      GOOGLE_WORKSPACE_PLAN: scriptProperties.getProperty('GOOGLE_WORKSPACE_PLAN') || 'Business Standard',
      NOTIFICATION_EMAIL: scriptProperties.getProperty('NOTIFICATION_EMAIL') || '',
      
      // API設定
      TRUSTLOGIN_API_TOKEN: scriptProperties.getProperty('TRUSTLOGIN_API_TOKEN') || '',
      ACTIVEGATE_DOMAIN: scriptProperties.getProperty('ACTIVEGATE_DOMAIN') || '',
      ACTIVEGATE_API_KEY: scriptProperties.getProperty('ACTIVEGATE_API_KEY') || '',
      
      // Microsoft 365
      M365_TENANT_ID: scriptProperties.getProperty('M365_TENANT_ID') || '',
      M365_CLIENT_ID: scriptProperties.getProperty('M365_CLIENT_ID') || '',
      M365_CLIENT_SECRET: scriptProperties.getProperty('M365_CLIENT_SECRET') || '',
      
      // Zoom
      ZOOM_ACCOUNT_ID: scriptProperties.getProperty('ZOOM_ACCOUNT_ID') || '',
      ZOOM_CLIENT_ID: scriptProperties.getProperty('ZOOM_CLIENT_ID') || '',
      ZOOM_CLIENT_SECRET: scriptProperties.getProperty('ZOOM_CLIENT_SECRET') || '',
      
      // その他のプラットフォーム
      CANVA_API_TOKEN: scriptProperties.getProperty('CANVA_API_TOKEN') || '',
      ESET_API_TOKEN: scriptProperties.getProperty('ESET_API_TOKEN') || '',
      SLACK_API_TOKEN: scriptProperties.getProperty('SLACK_API_TOKEN') || '',
      
      // Notion
      NOTION_TOKEN: scriptProperties.getProperty('NOTION_TOKEN') || '',
      NOTION_PLAN: scriptProperties.getProperty('NOTION_PLAN') || 'Plus',
      
      // レポート設定
      REPORT_EMAIL: scriptProperties.getProperty('REPORT_EMAIL') || ''
    };
  } catch (error) {
    console.error('設定読み込みエラー:', error);
    // エラー時はデフォルト値を返す
    return {
      GOOGLE_DOMAIN: '',
      GOOGLE_WORKSPACE_PLAN: 'Business Standard',
      NOTIFICATION_EMAIL: '',
      TRUSTLOGIN_API_TOKEN: '',
      ACTIVEGATE_DOMAIN: '',
      ACTIVEGATE_API_KEY: '',
      M365_TENANT_ID: '',
      M365_CLIENT_ID: '',
      M365_CLIENT_SECRET: '',
      ZOOM_ACCOUNT_ID: '',
      ZOOM_CLIENT_ID: '',
      ZOOM_CLIENT_SECRET: '',
      CANVA_API_TOKEN: '',
      ESET_API_TOKEN: '',
      SLACK_API_TOKEN: '',
      // Notion defaults
      NOTION_TOKEN: '',
      NOTION_PLAN: 'Plus',
      // レポート設定
      REPORT_EMAIL: ''
    };
  }
}

// ===== 棚卸し管理機能 =====

/**
 * 棚卸し期間中かどうかをチェック
 * @param {Date} inventoryExpireDate - 棚卸し有効期限
 * @return {boolean} 棚卸し期間中の場合true
 */
function isInInventoryPeriod(inventoryExpireDate) {
  if (!inventoryExpireDate || inventoryExpireDate === '') return false;
  
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  
  const expireDate = new Date(inventoryExpireDate);
  expireDate.setHours(23, 59, 59, 999);
  
  return today <= expireDate;
}

/**
 * 棚卸しステータスを考慮した非アクティブアカウント分析
 * @param {Array} licenses - ライセンスデータ
 * @return {Object} 棚卸し期間を考慮した非アクティブアカウント情報
 */
function analyzeInactiveAccountsWithInventory(licenses) {
  const currentDate = new Date();
  const thresholdDate = new Date();
  thresholdDate.setDate(currentDate.getDate() - CONFIG.UNUSED_LICENSE_ALERT_DAYS);
  
  const inactiveByPlatform = {};
  const inventoryAccounts = {};  // 棚卸し中のアカウント
  
  // プラットフォームごとの初期化
  CONFIG.PLATFORMS.forEach(platform => {
    inactiveByPlatform[platform] = [];
    inventoryAccounts[platform] = [];
  });
  
  // スプレッドシートから棚卸し情報を取得
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const licenseSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.LICENSES);
  const holdSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.INVENTORY_HOLD);
  
  if (licenseSheet) {
    const dataRange = licenseSheet.getDataRange();
    const data = dataRange.getValues();
    const headers = data[0];
    
    // 棚卸し有効期限列のインデックスを取得
    const inventoryExpireIndex = headers.indexOf('棚卸し有効期限');
    const emailIndex = headers.indexOf('メールアドレス');
    const platformIndex = headers.indexOf('プラットフォーム');
    
    if (inventoryExpireIndex !== -1) {
      // 各行をチェック
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const inventoryExpireDate = row[inventoryExpireIndex];
        const email = row[emailIndex];
        const platform = row[platformIndex];
        
        // 棚卸し期間中のライセンスを別途記録
        if (isInInventoryPeriod(inventoryExpireDate)) {
          const license = licenses.find(l => 
            l.email === email && l.platform === platform
          );
          
          if (license && license.status === 'Active' && license.lastLogin) {
            const lastLoginDate = new Date(license.lastLogin);
            if (lastLoginDate < thresholdDate) {
              inventoryAccounts[platform].push({
                email: license.email,
                name: license.name,
                lastLogin: license.lastLogin,
                daysSinceLogin: Math.floor((currentDate - lastLoginDate) / (1000 * 60 * 60 * 24)),
                cost: license.cost || 0,
                isInventory: true,
                inventoryExpireDate: inventoryExpireDate
              });
            }
          }
        }
      }
    }
  }
  
  // 棚卸し保留シートのデータも確認
  if (holdSheet && holdSheet.getLastRow() > 1) {
    const holdData = holdSheet.getRange(2, 1, holdSheet.getLastRow() - 1, holdSheet.getLastColumn()).getValues();
    
    holdData.forEach(row => {
      const platform = row[0];
      const email = row[2];
      const inventoryExpireDate = row[12];
      
      if (isInInventoryPeriod(inventoryExpireDate)) {
        // 保留シートにあるライセンスも棚卸し中として記録
        const license = licenses.find(l => 
          l.email === email && l.platform === platform
        );
        
        if (license && license.status === 'Active' && license.lastLogin) {
          const lastLoginDate = new Date(license.lastLogin);
          if (lastLoginDate < thresholdDate) {
            if (!inventoryAccounts[platform].some(acc => acc.email === email)) {
              inventoryAccounts[platform].push({
                email: license.email,
                name: license.name,
                lastLogin: license.lastLogin,
                daysSinceLogin: Math.floor((currentDate - lastLoginDate) / (1000 * 60 * 60 * 24)),
                cost: license.cost || 0,
                isInventory: true,
                inventoryExpireDate: inventoryExpireDate,
                inHoldSheet: true
              });
            }
          }
        }
      }
    });
  }
  
  // 通常の非アクティブアカウント分析（棚卸し中を除外）
  licenses.forEach(license => {
    if (license.status === 'Active' && license.lastLogin) {
      const lastLoginDate = new Date(license.lastLogin);
      if (lastLoginDate < thresholdDate) {
        // 棚卸し中でないかチェック
        const isInventory = inventoryAccounts[license.platform].some(
          acc => acc.email === license.email
        );
        
        if (!isInventory) {
          const daysSinceLogin = Math.floor((currentDate - lastLoginDate) / (1000 * 60 * 60 * 24));
          inactiveByPlatform[license.platform].push({
            email: license.email,
            name: license.name,
            lastLogin: license.lastLogin,
            daysSinceLogin: daysSinceLogin,
            cost: license.cost || 0,
            isInventory: false
          });
        }
      }
    }
  });
  
  return {
    inactive: inactiveByPlatform,
    inventory: inventoryAccounts
  };
}

/**
 * ライセンス一覧シートに条件付き書式を適用（棚卸し期間のハイライト）
 */
function applyInventoryHighlight() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.LICENSES);
  
  if (!sheet) {
    console.error('ライセンス一覧シートが見つかりません');
    return;
  }
  
  const dataRange = sheet.getDataRange();
  const data = dataRange.getValues();
  const headers = data[0];
  
  // 棚卸し有効期限列のインデックスを取得
  const inventoryExpireIndex = headers.indexOf('棚卸し有効期限');
  
  if (inventoryExpireIndex === -1) {
    console.error('棚卸し有効期限列が見つかりません');
    return;
  }
  
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  
  // データ行をループして色を設定
  for (let i = 1; i < data.length; i++) {
    const inventoryExpireDate = data[i][inventoryExpireIndex];
    
    if (inventoryExpireDate && inventoryExpireDate !== '') {
      const expireDate = new Date(inventoryExpireDate);
      expireDate.setHours(23, 59, 59, 999);
      
      // 棚卸し期間中の場合、行全体を薄い黄色でハイライト
      if (today <= expireDate) {
        sheet.getRange(i + 1, 1, 1, headers.length).setBackground('#fff2cc');
      } else {
        // 期限切れの場合は通常の背景色に戻す
        sheet.getRange(i + 1, 1, 1, headers.length).setBackground(null);
      }
    } else {
      // 棚卸し期限が設定されていない場合は通常の背景色
      sheet.getRange(i + 1, 1, 1, headers.length).setBackground(null);
    }
  }
  
  console.log('棚卸し期間のハイライトを適用しました');
}

/**
 * メイン処理の拡張版（棚卸し機能を含む）
 */
function mainWithInventory() {
  try {
    // 通常のメイン処理を実行
    main();
    
    // 棚卸し期間のハイライトを適用
    applyInventoryHighlight();
    
    console.log('棚卸し機能を含むライセンス管理処理が完了しました');
    
  } catch (error) {
    console.error('エラーが発生しました:', error);
    SpreadsheetApp.getUi().alert('エラー', 'エラーが発生しました: ' + error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * 棚卸し保留シートの初期化とヘッダー設定
 */
function initializeInventoryHoldSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.INVENTORY_HOLD);
  
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_NAMES.INVENTORY_HOLD);
  }
  
  // ヘッダー設定
  const headers = [
    'プラットフォーム', 'ユーザーID', 'メールアドレス', '氏名', 
    'ライセンスタイプ', 'ステータス', '割当日', '最終ログイン', 
    '利用状況', 'コスト（月額）', '備考', '管理メモ', 
    '棚卸し有効期限', '保留開始日時', '自動削除予定日時'
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setBackground('#f4a460')
    .setFontColor('#ffffff')
    .setFontWeight('bold');
  
  return sheet;
}

/**
 * ライセンス一覧から棚卸し保留シートへの自動転記
 */
function updateInventoryHoldSheet() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const licenseSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.LICENSES);
    const holdSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.INVENTORY_HOLD) || initializeInventoryHoldSheet();
    
    if (!licenseSheet) {
      console.error('ライセンス一覧シートが見つかりません');
      return;
    }
    
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    
    // ライセンス一覧のデータを取得
    const licenseData = licenseSheet.getDataRange().getValues();
    const licenseHeaders = licenseData[0];
    
    // 各列のインデックスを取得
    const inventoryExpireIndex = licenseHeaders.indexOf('棚卸し有効期限');
    if (inventoryExpireIndex === -1) {
      console.error('棚卸し有効期限列が見つかりません');
      return;
    }
    
    // 棚卸し中のライセンスを抽出
    const inventoryLicenses = [];
    for (let i = 1; i < licenseData.length; i++) {
      const row = licenseData[i];
      const inventoryExpireDate = row[inventoryExpireIndex];
      
      if (inventoryExpireDate && inventoryExpireDate !== '') {
        const expireDate = new Date(inventoryExpireDate);
        expireDate.setHours(23, 59, 59, 999);
        
        // 棚卸し期間中の場合
        if (today <= expireDate) {
          // 行データに追加情報を付加
          const holdRow = [...row];
          holdRow.push(new Date()); // 保留開始日時
          holdRow.push(expireDate); // 自動削除予定日時
          inventoryLicenses.push(holdRow);
        }
      }
    }
    
    // 既存の保留シートのデータを取得
    const existingData = holdSheet.getLastRow() > 1 
      ? holdSheet.getRange(2, 1, holdSheet.getLastRow() - 1, holdSheet.getLastColumn()).getValues()
      : [];
    
    // 期限切れのデータを除外
    const validExistingData = existingData.filter(row => {
      const expireDate = row[12]; // 棚卸し有効期限のインデックス
      if (expireDate && expireDate !== '') {
        const expire = new Date(expireDate);
        expire.setHours(23, 59, 59, 999);
        return today <= expire;
      }
      return false;
    });
    
    // 新規追加するライセンスを特定
    const newLicenses = inventoryLicenses.filter(newRow => {
      const email = newRow[2]; // メールアドレス
      const platform = newRow[0]; // プラットフォーム
      
      return !validExistingData.some(existingRow => 
        existingRow[2] === email && existingRow[0] === platform
      );
    });
    
    // ヘッダーが存在しない場合は設定
    const currentHeaders = holdSheet.getRange(1, 1, 1, 15).getValues()[0];
    if (!currentHeaders[0] || currentHeaders[0] === '') {
      const headers = [
        'プラットフォーム', 'ユーザーID', 'メールアドレス', '氏名', 
        'ライセンスタイプ', 'ステータス', '割当日', '最終ログイン', 
        '利用状況', 'コスト（月額）', '備考', '管理メモ', 
        '棚卸し有効期限', '保留開始日時', '自動削除予定日時'
      ];
      holdSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      holdSheet.getRange(1, 1, 1, headers.length)
        .setBackground('#f4a460')
        .setFontColor('#ffffff')
        .setFontWeight('bold');
    }
    
    // 保留シートをクリアして更新（ヘッダー行は除外）
    if (holdSheet.getLastRow() > 1) {
      holdSheet.getRange(2, 1, holdSheet.getLastRow() - 1, holdSheet.getLastColumn()).clearContent();
      holdSheet.getRange(2, 1, holdSheet.getLastRow() - 1, holdSheet.getLastColumn()).setBackground(null);
    }
    
    // データを書き込み
    if (inventoryLicenses.length > 0) {
      // 最大15列のデータを書き込み
      const writeData = inventoryLicenses.map(row => {
        // 配列の長さを15に調整
        while (row.length < 15) {
          row.push('');
        }
        return row.slice(0, 15);
      });
      
      holdSheet.getRange(2, 1, writeData.length, 15).setValues(writeData);
      
      // 棚卸し期間中の行を薄いオレンジ色でハイライト
      holdSheet.getRange(2, 1, writeData.length, 15).setBackground('#ffe4b5');
    }
    
    console.log(`棚卸し保留シートを更新: ${inventoryLicenses.length}件のライセンスを管理中`);
    
    // 変更履歴を記録
    if (newLicenses.length > 0) {
      recordHistory(ss, `棚卸し保留シートに${newLicenses.length}件追加`);
    }
    
    return {
      total: inventoryLicenses.length,
      new: newLicenses.length,
      removed: existingData.length - validExistingData.length
    };
    
  } catch (error) {
    console.error('棚卸し保留シート更新エラー:', error);
  }
}

/**
 * 棚卸しステータスの確認と報告
 */
function checkInventoryStatus() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.LICENSES);
  
  if (!sheet) {
    SpreadsheetApp.getUi().alert('エラー', 'ライセンス一覧シートが見つかりません', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  const dataRange = sheet.getDataRange();
  const data = dataRange.getValues();
  const headers = data[0];
  
  const inventoryExpireIndex = headers.indexOf('棚卸し有効期限');
  const emailIndex = headers.indexOf('メールアドレス');
  const platformIndex = headers.indexOf('プラットフォーム');
  const statusIndex = headers.indexOf('ステータス');
  
  if (inventoryExpireIndex === -1) {
    SpreadsheetApp.getUi().alert('エラー', '棚卸し有効期限列が見つかりません', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  
  let activeInventoryCount = 0;
  let expiredInventoryCount = 0;
  const activeInventoryList = [];
  const expiredInventoryList = [];
  
  for (let i = 1; i < data.length; i++) {
    const inventoryExpireDate = data[i][inventoryExpireIndex];
    
    if (inventoryExpireDate && inventoryExpireDate !== '') {
      const expireDate = new Date(inventoryExpireDate);
      const email = data[i][emailIndex];
      const platform = data[i][platformIndex];
      const status = data[i][statusIndex];
      
      if (today <= expireDate) {
        activeInventoryCount++;
        activeInventoryList.push({
          email: email,
          platform: platform,
          expireDate: Utilities.formatDate(expireDate, 'JST', 'yyyy/MM/dd'),
          status: status
        });
      } else {
        expiredInventoryCount++;
        expiredInventoryList.push({
          email: email,
          platform: platform,
          expireDate: Utilities.formatDate(expireDate, 'JST', 'yyyy/MM/dd'),
          status: status
        });
      }
    }
  }
  
  let message = '=== 棚卸しステータス ===\n\n';
  message += `現在の日付: ${Utilities.formatDate(today, 'JST', 'yyyy/MM/dd')}\n\n`;
  
  message += `【棚卸し中のライセンス】 ${activeInventoryCount}件\n`;
  if (activeInventoryCount > 0 && activeInventoryCount <= 10) {
    activeInventoryList.forEach(item => {
      message += `  • ${item.platform}: ${item.email} (期限: ${item.expireDate})\n`;
    });
  } else if (activeInventoryCount > 10) {
    message += `  （詳細はライセンス一覧シートをご確認ください）\n`;
  }
  
  message += `\n【期限切れの棚卸し】 ${expiredInventoryCount}件\n`;
  if (expiredInventoryCount > 0 && expiredInventoryCount <= 5) {
    expiredInventoryList.forEach(item => {
      message += `  • ${item.platform}: ${item.email} (期限: ${item.expireDate})\n`;
    });
  } else if (expiredInventoryCount > 5) {
    message += `  （詳細はライセンス一覧シートをご確認ください）\n`;
  }
  
  if (activeInventoryCount === 0 && expiredInventoryCount === 0) {
    message += '\n現在、棚卸し設定されているライセンスはありません。\n';
  }
  
  message += '\n※棚卸し期間中のライセンスは通知対象から除外されます。';
  
  SpreadsheetApp.getUi().alert('棚卸しステータス', message, SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * 自動転記トリガーの設定（5分ごと）
 */
function setupInventoryHoldTrigger() {
  try {
    // 既存のトリガーを削除
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(trigger => {
      if (trigger.getHandlerFunction() === 'updateInventoryHoldSheet') {
        ScriptApp.deleteTrigger(trigger);
      }
    });
    
    // 新しいトリガーを作成（5分ごと）
    ScriptApp.newTrigger('updateInventoryHoldSheet')
      .timeBased()
      .everyMinutes(5)
      .create();
    
    SpreadsheetApp.getUi().alert(
      '自動転記設定完了',
      '棚卸し保留シートへの自動転記が5分ごとに実行されるよう設定されました。',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
    console.log('棚卸し保留シート自動転記トリガーを設定しました');
    
  } catch (error) {
    console.error('トリガー設定エラー:', error);
    SpreadsheetApp.getUi().alert(
      'エラー',
      'トリガーの設定中にエラーが発生しました: ' + error.toString(),
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * 自動転記トリガーの削除
 */
function removeInventoryHoldTrigger() {
  try {
    const triggers = ScriptApp.getProjectTriggers();
    let removed = 0;
    
    triggers.forEach(trigger => {
      if (trigger.getHandlerFunction() === 'updateInventoryHoldSheet') {
        ScriptApp.deleteTrigger(trigger);
        removed++;
      }
    });
    
    if (removed > 0) {
      SpreadsheetApp.getUi().alert(
        '自動転記停止',
        '棚卸し保留シートへの自動転記を停止しました。',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      console.log(`${removed}個の棚卸し保留シート自動転記トリガーを削除しました`);
    } else {
      SpreadsheetApp.getUi().alert(
        '情報',
        '自動転記は現在設定されていません。',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
    
  } catch (error) {
    console.error('トリガー削除エラー:', error);
    SpreadsheetApp.getUi().alert(
      'エラー',
      'トリガーの削除中にエラーが発生しました: ' + error.toString(),
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * 棚卸し保留シートの手動更新（メニュー用）
 */
function manualUpdateInventoryHoldSheet() {
  try {
    const result = updateInventoryHoldSheet();
    
    if (result) {
      let message = `棚卸し保留シートを更新しました。\n\n`;
      message += `管理中のライセンス: ${result.total}件\n`;
      message += `新規追加: ${result.new}件\n`;
      message += `期限切れ削除: ${result.removed}件`;
      
      SpreadsheetApp.getUi().alert('更新完了', message, SpreadsheetApp.getUi().ButtonSet.OK);
    } else {
      SpreadsheetApp.getUi().alert('更新完了', '棚卸し保留シートを更新しました。', SpreadsheetApp.getUi().ButtonSet.OK);
    }
  } catch (error) {
    console.error('手動更新エラー:', error);
    SpreadsheetApp.getUi().alert('エラー', 'エラーが発生しました: ' + error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * 棚卸し保留シートの初期化（メニュー用）
 */
function resetInventoryHoldSheet() {
  try {
    const sheet = initializeInventoryHoldSheet();
    
    // 既存データをクリア（ヘッダー以外）
    if (sheet.getLastRow() > 1) {
      sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clearContent();
      sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).setBackground(null);
    }
    
    SpreadsheetApp.getUi().alert(
      '初期化完了',
      '棚卸し保留シートを初期化しました。\nヘッダーが正しく設定されています。',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
    console.log('棚卸し保留シートを初期化しました');
    
  } catch (error) {
    console.error('初期化エラー:', error);
    SpreadsheetApp.getUi().alert(
      'エラー',
      '初期化中にエラーが発生しました: ' + error.toString(),
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * 棚卸し保留状況の確認
 */
function checkInventoryHoldStatus() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const holdSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.INVENTORY_HOLD);
    
    if (!holdSheet) {
      SpreadsheetApp.getUi().alert('情報', '棚卸し保留シートが存在しません。', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    const dataRange = holdSheet.getDataRange();
    const data = dataRange.getValues();
    
    if (data.length <= 1) {
      SpreadsheetApp.getUi().alert('情報', '現在、棚卸し保留中のライセンスはありません。', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    // プラットフォーム別に集計
    const platformCount = {};
    for (let i = 1; i < data.length; i++) {
      const platform = data[i][0];
      platformCount[platform] = (platformCount[platform] || 0) + 1;
    }
    
    // トリガーの状態を確認
    const triggers = ScriptApp.getProjectTriggers();
    const autoUpdateEnabled = triggers.some(trigger => 
      trigger.getHandlerFunction() === 'updateInventoryHoldSheet'
    );
    
    let message = '=== 棚卸し保留状況 ===\n\n';
    message += `保留中のライセンス総数: ${data.length - 1}件\n\n`;
    
    message += '【プラットフォーム別】\n';
    Object.entries(platformCount).forEach(([platform, count]) => {
      message += `  • ${platform}: ${count}件\n`;
    });
    
    message += `\n自動転記: ${autoUpdateEnabled ? '有効（5分ごと）' : '無効'}\n`;
    message += '\n※棚卸し保留中のライセンスは通知対象から自動的に除外されます。\n';
    message += '※有効期限を過ぎたライセンスは自動的に保留シートから削除されます。';
    
    SpreadsheetApp.getUi().alert('棚卸し保留状況', message, SpreadsheetApp.getUi().ButtonSet.OK);
    
  } catch (error) {
    console.error('保留状況確認エラー:', error);
    SpreadsheetApp.getUi().alert('エラー', 'エラーが発生しました: ' + error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}