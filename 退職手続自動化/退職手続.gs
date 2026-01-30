/**
 * 退職者処理システム - SSO検出強化版（エラー修正済み）
 * Version: 2.0.1
 * 
 * 機能:
 * - センシティブメール削除
 * - メール転送設定（個人レベル）
 * - 管理コンソールルーティング設定案内
 * - メール委任設定
 * - 外部SSOサービス検出（強化版）
 * - GASプロジェクト一覧取得
 * - 設定管理機能
 */

// ===== グローバル設定キャッシュ =====
let CONFIG_CACHE = null;

// ===== 初期化とメニュー =====

/**
 * スプレッドシート起動時の初期設定
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  ui.createMenu('👤 退職者処理')
    .addItem('📋 初期設定シートを作成', 'createInitialSheets')
    .addItem('▶️ 退職処理を実行', 'main')
    .addSeparator()
    .addSubMenu(ui.createMenu('🔧 個別機能')
      .addItem('📧 センシティブメールの削除', 'runDeleteEmails')
      .addItem('📋 センシティブメール抽出（確認用）', 'runExtractSensitiveEmails')
      .addItem('✅ 選択したメールを削除', 'runDeleteSelectedEmails')
      .addItem('🔗 外部SSOサービス詳細取得', 'runDetailedSSOAnalysis')
      .addItem('📱 利用デバイス一覧取得', 'runListUserDevices')
      .addItem('📄 データ一覧取得（ドキュメント・スプレッドシート）', 'runListDataFiles')
      .addItem('📅 カレンダー一覧取得', 'runListUserCalendars'))
    .addSeparator()
    .addSubMenu(ui.createMenu('📬 メールルーティング')
      .addItem('🔄 管理コンソール設定案内', 'showAdminRoutingGuide')
      .addItem('➡️ 個人転送設定（ユーザーレベル）', 'runEmailRouting')
      .addItem('👥 メール委任設定', 'runEmailDelegation')
      .addItem('📊 ルーティング設定CSV出力', 'exportRoutingConfigCSV')
      .addItem('🔍 現在の転送設定を確認', 'runCheckForwarding')
      .addItem('⏸️ 転送を無効化', 'runDisableForwarding'))
    .addSeparator()
    .addSubMenu(ui.createMenu('📊 分析・レポート')
      .addItem('📈 退職者総合レポート作成', 'generateComprehensiveReport')
      .addItem('🔍 外部サービス利用状況分析', 'analyzeExternalServices')
      .addItem('📱 デバイス利用状況分析', 'analyzeDeviceUsage'))
    .addSeparator()
    .addItem('♻️ 削除メールの復元', 'showRestoreDialog')
    .addItem('🔍 システム診断', 'runSystemDiagnostics')
    .addSeparator()
    .addSubMenu(ui.createMenu('⚙️ 設定管理')
      .addItem('📊 設定画面を開く', 'showConfigurationUI')
      .addItem('💾 設定をエクスポート', 'exportConfiguration')
      .addItem('🔄 設定を初期化', 'confirmInitializeConfiguration'))
    .addSeparator()
    .addItem('❓ ヘルプ', 'showHelp')
    .addToUi();
}

/**
 * 初回セットアップ（手動実行）
 */
function initialSetup() {
  initializeConfiguration();
  createInitialSheets();
  SpreadsheetApp.getUi().alert('セットアップ完了', '初期設定が完了しました。', SpreadsheetApp.getUi().ButtonSet.OK);
}

// ===== 設定管理 =====

/**
 * 設定の初期化
 */
function initializeConfiguration() {
  const scriptProperties = PropertiesService.getScriptProperties();
  
  const defaultConfig = {
    "system": {
      "version": "2.0.1",
      "name": "退職者処理システム（SSO検出強化版）"
    },
    "defaults": {
      "forwardToEmail": "admin@example.com",
      "maxEmailsToProcess": 50,
      "sensitiveKeywords": [
        "健康診断",
        "給与明細",
        "人事評価",
        "査定",
        "賞与",
        "昇進",
        "退職金",
        "機密",
        "パスワード",
        "個人情報"
      ],
      "routingMethod": "user",
      "ssoLookbackDays": 365,
      "popularSSOServices": [
        "Canva", "Zapier", "Slack", "Zoom", "Notion",
        "Miro", "Figma", "Dropbox", "Asana", "Trello",
        "Monday", "Airtable", "Calendly", "DocuSign", "Typeform",
        "Mailchimp", "HubSpot", "Salesforce", "Jira", "GitHub",
        "Chatwork", "LINE WORKS", "Microsoft Teams", "Box",
        "Adobe Creative Cloud", "Office 365", "Wrike", "ClickUp",
        "Sansan", "freee", "SmartHR", "マネーフォワード", "kintone", "Backlog"
      ]
    },
    "sheets": {
      "config": "設定",
      "deletedEmails": "削除メール一覧",
      "saas": "外部サービス一覧",
      "devices": "利用デバイス一覧",
      "dataFiles": "データ一覧",
      "calendars": "カレンダー一覧",
      "processLog": "処理ログ",
      "restoreLog": "復元ログ",
      "forwardingLog": "転送設定ログ",
      "routingCSV": "ルーティング設定CSV",
      "comprehensiveReport": "退職者総合レポート"
    },
    "cells": {
      "userEmail": "B3",
      "forwardEmail": "B4",
      "keywordStartRow": 7,
      "keywordStartColumn": 2
    }
  };
  
  scriptProperties.setProperty('CONFIG', JSON.stringify(defaultConfig));
  CONFIG_CACHE = null;
}

/**
 * 設定を取得
 */
function getConfiguration() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const configString = scriptProperties.getProperty('CONFIG');
  
  if (!configString) {
    initializeConfiguration();
    return getConfiguration();
  }
  
  try {
    return JSON.parse(configString);
  } catch (error) {
    console.error("設定のパースエラー:", error);
    initializeConfiguration();
    return getConfiguration();
  }
}

/**
 * 設定を更新
 */
function updateConfiguration(newConfig) {
  const scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperty('CONFIG', JSON.stringify(newConfig));
  CONFIG_CACHE = null;
}

/**
 * 設定を取得（キャッシュ付き）
 */
function getConfig() {
  if (!CONFIG_CACHE) {
    try {
      CONFIG_CACHE = getConfiguration();
      
      // 必要なプロパティが存在するか確認
      if (!CONFIG_CACHE.defaults) {
        CONFIG_CACHE.defaults = {};
      }
      if (!CONFIG_CACHE.defaults.sensitiveKeywords) {
        CONFIG_CACHE.defaults.sensitiveKeywords = [];
      }
      if (!CONFIG_CACHE.defaults.popularSSOServices) {
        CONFIG_CACHE.defaults.popularSSOServices = [];
      }
      if (!CONFIG_CACHE.sheets) {
        CONFIG_CACHE.sheets = {};
      }
      if (!CONFIG_CACHE.cells) {
        CONFIG_CACHE.cells = {};
      }
      
    } catch (error) {
      console.error("設定取得エラー:", error);
      // デフォルト設定を返す
      initializeConfiguration();
      CONFIG_CACHE = getConfiguration();
    }
  }
  return CONFIG_CACHE;
}

/**
 * 設定管理UIを表示
 */
function showConfigurationUI() {
  const html = HtmlService.createHtmlOutputFromFile('config-editor')
    .setWidth(600)
    .setHeight(500);
  SpreadsheetApp.getUi().showModalDialog(html, '設定管理');
}

/**
 * 設定をエクスポート
 */
function exportConfiguration() {
  const config = getConfiguration();
  const blob = Utilities.newBlob(JSON.stringify(config, null, 2), 'application/json', 'retirement-system-config.json');
  const file = DriveApp.createFile(blob);
  
  SpreadsheetApp.getUi().alert(
    '設定のエクスポート完了',
    '設定ファイルをGoogle Driveに保存しました。\n\nファイル名: retirement-system-config.json\nURL: ' + file.getUrl(),
    SpreadsheetApp.getUi().ButtonSet.OK
  );
  
  return file.getUrl();
}

/**
 * 設定の初期化確認
 */
function confirmInitializeConfiguration() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    '設定の初期化',
    '設定を初期状態に戻しますか？\nこの操作は取り消せません。',
    ui.ButtonSet.YES_NO
  );
  
  if (response === ui.Button.YES) {
    initializeConfiguration();
    ui.alert('設定を初期化しました。');
  }
}

// ===== 初期設定 =====

/**
 * 初期設定シートの作成
 */
function createInitialSheets() {
  const config = getConfig();
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // 設定シートの作成または取得
  let configSheet = spreadsheet.getSheetByName(config.sheets.config);
  if (!configSheet) {
    configSheet = spreadsheet.insertSheet(config.sheets.config, 0);
  }
  
  // 設定シートの初期化
  configSheet.clear();
  
  // タイトル
  configSheet.getRange("A1").setValue(config.system.name).setFontSize(16).setFontWeight("bold");
  configSheet.getRange("A2").setValue("Version " + config.system.version).setFontSize(10).setFontColor("#666666");
  
  // 入力欄
  configSheet.getRange("A3").setValue("退職者メールアドレス:");
  configSheet.getRange(config.cells.userEmail).setValue("").setBackground("#FFFACD")
    .setBorder(true, true, true, true, true, true);
    
  configSheet.getRange("A4").setValue("転送先メールアドレス:");
  configSheet.getRange(config.cells.forwardEmail).setValue(config.defaults.forwardToEmail)
    .setBackground("#E6E6FA").setBorder(true, true, true, true, true, true);
  
  // 削除対象キーワード
  configSheet.getRange("A6").setValue("削除対象キーワード:").setFontWeight("bold");
  
  // sensitiveKeywordsが存在するかチェック
  const keywords = config.defaults.sensitiveKeywords || [];
  keywords.forEach((keyword, index) => {
    configSheet.getRange(config.cells.keywordStartRow + index, 1).setValue("キーワード" + (index + 1) + ":");
    configSheet.getRange(config.cells.keywordStartRow + index, config.cells.keywordStartColumn)
      .setValue(keyword).setBackground("#FFE4E1");
  });
  
  // メールルーティング設定方法
  const routingRow = config.cells.keywordStartRow + keywords.length + 2;
  configSheet.getRange(routingRow, 1).setValue("メールルーティング設定:").setFontWeight("bold");
  configSheet.getRange(routingRow + 1, 1).setValue("推奨方法:");
  configSheet.getRange(routingRow + 1, 2).setValue("管理コンソール（組織レベル）").setFontColor("#0000FF");
  configSheet.getRange(routingRow + 2, 1).setValue("代替方法:");
  configSheet.getRange(routingRow + 2, 2).setValue("個人転送設定（ユーザーレベル）");
  
  // APIサービス状態
  const apiRow = routingRow + 4;
  configSheet.getRange(apiRow, 1).setValue("APIサービス状態:").setFontWeight("bold");
  
  const gmailStatus = typeof Gmail !== 'undefined' ? "✓ 有効" : "✗ 無効（追加が必要）";
  configSheet.getRange(apiRow + 1, 1).setValue("Gmail API:");
  configSheet.getRange(apiRow + 1, 2).setValue(gmailStatus)
    .setFontColor(gmailStatus.includes("✓") ? "green" : "red");
  
  const adminDirectoryStatus = typeof AdminDirectory !== 'undefined' ? "✓ 有効" : "✗ 無効（追加が必要）";
  configSheet.getRange(apiRow + 2, 1).setValue("Admin Directory API:");
  configSheet.getRange(apiRow + 2, 2).setValue(adminDirectoryStatus)
    .setFontColor(adminDirectoryStatus.includes("✓") ? "green" : "red");
  
  const adminReportsStatus = typeof AdminReports !== 'undefined' ? "✓ 有効" : "✗ 無効（追加が必要）";
  configSheet.getRange(apiRow + 3, 1).setValue("Admin Reports API:");
  configSheet.getRange(apiRow + 3, 2).setValue(adminReportsStatus)
    .setFontColor(adminReportsStatus.includes("✓") ? "green" : "red");
  
  // 実行方法
  configSheet.getRange(apiRow + 5, 1).setValue("実行方法:").setFontWeight("bold");
  configSheet.getRange(apiRow + 6, 1).setValue("1. 上記のB3セルに退職者のメールアドレスを入力");
  configSheet.getRange(apiRow + 7, 1).setValue("2. メニューの「退職者処理」→「退職処理を実行」を選択");
  configSheet.getRange(apiRow + 8, 1).setValue("※ メールルーティングは「メールルーティング」メニューから設定");
  
  // 列幅の調整
  configSheet.setColumnWidth(1, 250);
  configSheet.setColumnWidth(2, 300);
  
  SpreadsheetApp.getUi().alert(
    "初期設定完了",
    "初期設定シートを作成しました。\n\n" +
    "1. B3セルに退職者のメールアドレスを入力\n" +
    "2. メニューから処理を実行してください\n\n" +
    "メールルーティングは管理コンソールでの設定を推奨します。",
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * 設定シートから値を読み取る
 */
function getConfigFromSheet() {
  const config = getConfig();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(config.sheets.config);
  
  if (!sheet) {
    throw new Error("設定シートが見つかりません。「退職者処理」→「初期設定シートを作成」を実行してください。");
  }
  
  const userEmail = sheet.getRange(config.cells.userEmail).getValue().trim();
  const forwardEmail = sheet.getRange(config.cells.forwardEmail).getValue().trim();
  
  return {
    userEmail: userEmail,
    forwardEmail: forwardEmail || config.defaults.forwardToEmail,
    keywords: config.defaults.sensitiveKeywords || []
  };
}

// ===== メイン処理 =====

/**
 * メイン処理（全処理を実行）
 */
function main() {
  try {
    const userConfig = getConfigFromSheet();
    const systemConfig = getConfig();
    
    if (!userConfig.userEmail) {
      SpreadsheetApp.getUi().alert("エラー", "退職者のメールアドレスを入力してください。", SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    // 処理開始の確認
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      "退職者処理の確認",
      "対象: " + userConfig.userEmail + "\n" +
      "転送先: " + userConfig.forwardEmail + "\n" +
      "削除キーワード: " + userConfig.keywords.join(", ") + "\n\n" +
      "実行内容：\n" +
      "1. センシティブなメールの削除（ゴミ箱へ移動）\n" +
      "2. 外部SSOサービスの詳細取得\n" +
      "3. 利用デバイスの一覧取得\n" +
      "4. データ一覧の取得\n" +
      "5. カレンダー一覧の取得\n\n" +
      "※ メールルーティングは別途設定が必要です\n\n" +
      "続行しますか？",
      ui.ButtonSet.YES_NO
    );
    
    if (response !== ui.Button.YES) {
      ui.alert("処理をキャンセルしました。");
      return;
    }
    
    // 処理実行ログの開始
    logProcessStart(userConfig.userEmail);
    
    let successCount = 0;
    let errorMessages = [];
    
    // 各処理を実行（エラーが発生しても続行）
    try {
      const deletedEmails = deleteSensitiveEmails(userConfig, systemConfig);
      successCount++;
    } catch (error) {
      errorMessages.push("メール削除: " + error.message);
    }
    
    try {
      listExternalSSOServices(userConfig.userEmail, systemConfig);
      successCount++;
    } catch (error) {
      errorMessages.push("外部サービス一覧取得: " + error.message);
    }
    
    try {
      listUserDevices(userConfig.userEmail, systemConfig);
      successCount++;
    } catch (error) {
      errorMessages.push("デバイス一覧取得: " + error.message);
    }
    
    try {
      listUserDataFiles(userConfig.userEmail, systemConfig);
      successCount++;
    } catch (error) {
      errorMessages.push("データファイル一覧: " + error.message);
    }
    
    // 処理完了の通知
    let resultMessage = "処理が完了しました。\n\n成功: " + successCount + "/4 項目\n";
    if (errorMessages.length > 0) {
      resultMessage += "\n以下の処理でエラーが発生しました:\n" + errorMessages.join("\n");
    }
    resultMessage += "\n\n詳細は各シートをご確認ください。";
    resultMessage += "\n\n【重要】メールルーティング設定は別途行ってください。";
    resultMessage += "\nメニュー「メールルーティング」→「管理コンソール設定案内」を参照";
    
    ui.alert("処理完了", resultMessage, ui.ButtonSet.OK);
    
    // 処理実行ログの終了
    logProcessEnd(userConfig.userEmail, errorMessages.length > 0 ? "一部エラー" : "成功");
    
  } catch (error) {
    SpreadsheetApp.getUi().alert("エラー", "処理中にエラーが発生しました: " + error.message, SpreadsheetApp.getUi().ButtonSet.OK);
    console.error(error);
    logProcessEnd(userConfig.userEmail || "不明", "エラー: " + error.message);
  }
}

// ===== 外部SSOサービス検出関数 =====

/**
 * 外部SSOサービスの詳細を取得してシートに記録
 */
function listExternalSSOServices(userEmail, systemConfig) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(systemConfig.sheets.saas) 
      || SpreadsheetApp.getActiveSpreadsheet().insertSheet(systemConfig.sheets.saas);
    sheet.clear();
    
    // ヘッダー設定
    const headers = [
      "サービス名",
      "サービスタイプ",
      "最終利用日時",
      "認可日時",
      "認可タイプ",
      "スコープ",
      "ステータス"
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length)
      .setBackground('#4285F4')
      .setFontColor('#FFFFFF')
      .setFontWeight('bold');
    
    // Admin SDKが利用可能かチェック
    if (typeof AdminReports === 'undefined') {
      sheet.appendRow(["Admin Reports APIが有効になっていません", "", "", "", "", "", ""]);
      sheet.appendRow(["Apps Scriptエディタで「サービス」→「＋」→「Admin SDK API」を追加してください", "", "", "", "", "", ""]);
      return;
    }
    
    // 外部サービス情報を取得
    const services = getDetailedExternalSSOServices(userEmail, systemConfig.defaults.ssoLookbackDays || 365);
    
    if (services.length === 0) {
      sheet.appendRow(["外部SSOサービスは見つかりませんでした", "", "", "", "", "", ""]);
    } else {
      // データ行を追加
      const dataRows = services.map(service => [
        service.name,
        service.type,
        service.lastUsed ? formatDate(service.lastUsed) : 'N/A',
        service.authorizedDate ? formatDate(service.authorizedDate) : 'N/A',
        service.grantType,
        service.scope || 'N/A',
        service.status
      ]);
      
      sheet.getRange(2, 1, dataRows.length, headers.length).setValues(dataRows);
    }
    
    // サマリー情報を追加
    const summaryRow = sheet.getLastRow() + 2;
    sheet.getRange(summaryRow, 1).setValue("【サマリー】").setFontWeight("bold");
    sheet.getRange(summaryRow + 1, 1).setValue("総サービス数:");
    sheet.getRange(summaryRow + 1, 2).setValue(services.length);
    
    // サービスタイプ別集計
    const typeCount = {};
    services.forEach(service => {
      typeCount[service.type] = (typeCount[service.type] || 0) + 1;
    });
    
    let typeRow = summaryRow + 3;
    sheet.getRange(typeRow, 1).setValue("【タイプ別集計】").setFontWeight("bold");
    Object.entries(typeCount).forEach(([type, count], index) => {
      sheet.getRange(typeRow + 1 + index, 1).setValue(type + ":");
      sheet.getRange(typeRow + 1 + index, 2).setValue(count);
    });
    
    // 列幅の自動調整
    sheet.autoResizeColumns(1, headers.length);
    
    // フィルターの設定
    const lastRow = services.length + 1;
    if (lastRow > 1) {
      sheet.getRange(1, 1, lastRow, headers.length).createFilter();
    }
    
  } catch (error) {
    console.error("外部SSOサービス一覧取得エラー:", error);
    throw error;
  }
}

/**
 * 詳細な外部SSOサービス情報を取得
 */
function getDetailedExternalSSOServices(userEmail, lookbackDays) {
  const services = [];
  const serviceMap = new Map();
  
  const endTime = new Date();
  const startTime = new Date(endTime.getTime() - lookbackDays * 24 * 60 * 60 * 1000);
  
  try {
    // Token API から外部サービスを取得
    const response = AdminReports.Activities.list('all', 'token', {
      userKey: userEmail,
      startTime: startTime.toISOString(),
      endTime: endTime.toISOString(),
      maxResults: 500
    });
    
    // responseとitemsの存在確認
    if (response && response.items && Array.isArray(response.items)) {
      response.items.forEach(item => {
        const timestamp = item.id?.time;
        
        // eventsの存在確認
        if (item.events && Array.isArray(item.events)) {
          item.events.forEach(event => {
            if (event.name === 'authorize' || event.name === 'grant') {
              let serviceName = '';
              let scope = '';
              let grantType = event.name;
              
              // parametersの存在確認
              if (event.parameters && Array.isArray(event.parameters)) {
                event.parameters.forEach(param => {
                  if (param.name === 'app_name' || param.name === 'client_name') {
                    serviceName = param.value || param.stringValue || 'Unknown';
                  } else if (param.name === 'scope') {
                    scope = param.value || param.stringValue || '';
                  }
                });
              }
              
              if (serviceName && !serviceName.includes('Google')) {
                const serviceKey = serviceName.toLowerCase();
                
                if (!serviceMap.has(serviceKey)) {
                  serviceMap.set(serviceKey, {
                    name: serviceName,
                    type: categorizeService(serviceName),
                    lastUsed: timestamp,
                    authorizedDate: timestamp,
                    grantType: grantType,
                    scope: scope,
                    status: 'Active'
                  });
                } else {
                  // 最新の利用日時を更新
                  const existing = serviceMap.get(serviceKey);
                  if (timestamp > existing.lastUsed) {
                    existing.lastUsed = timestamp;
                  }
                  if (scope && !existing.scope) {
                    existing.scope = scope;
                  }
                }
              }
            }
          });
        }
      });
    }
  } catch (error) {
    console.error("Token API エラー:", error);
  }
  
  // MapからArrayに変換してソート
  return Array.from(serviceMap.values()).sort((a, b) => {
    return new Date(b.lastUsed) - new Date(a.lastUsed);
  });
}

/**
 * サービスをカテゴリ分類
 */
function categorizeService(serviceName) {
  const categories = {
    'コミュニケーション': ['Slack', 'Zoom', 'Teams', 'Chatwork', 'LINE WORKS', 'Discord'],
    'プロジェクト管理': ['Asana', 'Trello', 'Monday', 'Jira', 'Backlog', 'Wrike', 'ClickUp'],
    'デザイン・クリエイティブ': ['Canva', 'Figma', 'Miro', 'Adobe', 'Sketch'],
    'ストレージ・ファイル共有': ['Dropbox', 'Box', 'OneDrive'],
    'ドキュメント・ノート': ['Notion', 'Evernote', 'OneNote', 'Confluence'],
    'マーケティング': ['HubSpot', 'Mailchimp', 'Marketo'],
    'CRM・営業': ['Salesforce', 'Pipedrive', 'Zoho'],
    '会計・経理': ['freee', 'マネーフォワード', 'QuickBooks'],
    'HR・人事': ['SmartHR', 'WorkDay', 'BambooHR'],
    '自動化・連携': ['Zapier', 'IFTTT', 'Make'],
    '開発': ['GitHub', 'GitLab', 'Bitbucket'],
    'その他業務': ['DocuSign', 'Calendly', 'Typeform', 'SurveyMonkey', 'Airtable', 'kintone', 'Sansan']
  };
  
  const lowerServiceName = serviceName.toLowerCase();
  
  for (const [category, services] of Object.entries(categories)) {
    if (services.some(service => lowerServiceName.includes(service.toLowerCase()))) {
      return category;
    }
  }
  
  return 'その他';
}

/**
 * 利用デバイス一覧を取得（改善版：Chrome OS、モバイル、PC含む）
 */
function listUserDevices(userEmail, systemConfig) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(systemConfig.sheets.devices) 
      || SpreadsheetApp.getActiveSpreadsheet().insertSheet(systemConfig.sheets.devices);
    sheet.clear();
    
    // ヘッダー設定
    const headers = [
      "デバイスタイプ",
      "モデル",
      "OS",
      "シリアル番号",
      "識別子",
      "最終アクティブ",
      "ステータス",
      "IPアドレス",
      "組織単位"
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length)
      .setBackground('#34A853')
      .setFontColor('#FFFFFF')
      .setFontWeight('bold');
    
    if (typeof AdminDirectory === 'undefined') {
      sheet.appendRow(["Admin Directory APIが有効になっていません", "", "", "", "", "", "", "", ""]);
      sheet.appendRow(["Apps Scriptエディタで「サービス」→「＋」→「Admin SDK API」を追加してください", "", "", "", "", "", "", "", ""]);
      return;
    }
    
    const allDevices = getAllDevicesIncludingPC();
    console.log(`取得したデバイス数: ${allDevices.length}`);
    
    // デバイスタイプ別の内訳を表示
    const deviceCounts = {};
    allDevices.forEach(device => {
      const type = device.deviceType || 'Unknown';
      deviceCounts[type] = (deviceCounts[type] || 0) + 1;
    });
    console.log('デバイスタイプ別内訳:', deviceCounts);
    
    // 各デバイスの情報を処理
    let rowData = [];
    let errors = [];
    
    allDevices.forEach((device, index) => {
      try {
        if (index % 10 === 0) {
          console.log(`処理中: ${index + 1}/${allDevices.length}`);
        }
        
        const deviceInfo = processDevice(device, userEmail);
        if (deviceInfo) {
          rowData.push(deviceInfo);
        }
        
        // バッチで書き込み（メモリ効率化）
        if (rowData.length >= 100) {
          writeDataToSheet(sheet, rowData);
          rowData = [];
        }
      } catch (error) {
        console.error(`デバイス処理エラー: ${error.message}`);
        errors.push([
          new Date(),
          device.deviceId || device.resourceId || 'N/A',
          device.email?.[0] || device.annotatedUser || device.userEmail || 'N/A',
          error.message
        ]);
      }
    });
    
    // 残りのデータを書き込み
    if (rowData.length > 0) {
      writeDataToSheet(sheet, rowData);
    }
    
    // 分析結果を自動表示
    console.log('\n');
    analyzeSerialNumbers();
    
  } catch (error) {
    console.error("デバイス一覧取得エラー:", error);
    throw error;
  }
}

/**
 * すべてのデバイス情報を取得（PC含む）
 */
function getAllDevicesIncludingPC() {
  const devices = [];
  let pageToken = null;
  
  // 1. Chrome OSデバイス
  try {
    console.log('Chrome OSデバイスを取得中...');
    pageToken = null;
    do {
      const response = AdminDirectory.Chromeosdevices.list('my_customer', {
        pageToken: pageToken,
        maxResults: 100,
        projection: 'FULL'
      });
      
      if (response.chromeosdevices) {
        response.chromeosdevices.forEach(device => {
          device.deviceType = 'Chrome OS';
          devices.push(device);
        });
      }
      
      pageToken = response.nextPageToken;
    } while (pageToken);
    console.log(`Chrome OSデバイス: ${devices.length}台`);
  } catch (error) {
    console.log('Chrome OSデバイスの取得をスキップ:', error.message);
  }
  
  // 2. モバイルデバイス
  const mobileStartCount = devices.length;
  try {
    console.log('モバイルデバイスを取得中...');
    pageToken = null;
    do {
      const response = AdminDirectory.Mobiledevices.list('my_customer', {
        pageToken: pageToken,
        maxResults: 100,
        projection: 'FULL'
      });
      
      if (response.mobiledevices) {
        response.mobiledevices.forEach(device => {
          device.deviceType = 'Mobile';
          devices.push(device);
        });
      }
      
      pageToken = response.nextPageToken;
    } while (pageToken);
    console.log(`モバイルデバイス: ${devices.length - mobileStartCount}台`);
  } catch (error) {
    console.error('モバイルデバイス取得エラー:', error);
  }
  
  // 3. Chromeブラウザ（PC）
  const browserStartCount = devices.length;
  try {
    console.log('Chromeブラウザ（PC）を取得中...');
    pageToken = null;
    do {
      const response = AdminDirectory.Chromeosdevices.Browsers.list('my_customer', {
        pageToken: pageToken,
        maxResults: 100,
        projection: 'FULL'
      });
      
      if (response.browsers) {
        response.browsers.forEach(browser => {
          browser.deviceType = 'Browser';
          browser.deviceId = browser.deviceId || browser.browserId;
          browser.annotatedUser = browser.annotatedUser || browser.lastPolicyFetchTime;
          devices.push(browser);
        });
      }
      
      pageToken = response.nextPageToken;
    } while (pageToken);
    console.log(`Chromeブラウザ: ${devices.length - browserStartCount}台`);
  } catch (error) {
    console.log('Chromeブラウザの取得をスキップ:', error.message);
  }
  
  // 4. エンドポイント（Mac/Windows/Linux）
  const endpointStartCount = devices.length;
  try {
    console.log('エンドポイントデバイスを確認中...');
    
    // エンドポイント検証デバイス
    try {
      const endpointVerificationDevices = getEndpointVerificationDevices();
      endpointVerificationDevices.forEach(device => {
        devices.push(device);
      });
    } catch (error) {
      console.log('エンドポイント検証をスキップ:', error.message);
    }
    
    // ログイン履歴からPCを検出（IPアドレス付き）
    if (typeof AdminReports !== 'undefined') {
      try {
        console.log('ログイン履歴からPC/Macを検出中...');
        const detectedPCs = detectPCsFromLoginHistory();
        detectedPCs.forEach(pc => {
          // 重複チェック（既存のエンドポイントデバイスと重複しない場合のみ追加）
          const isDuplicate = devices.some(d => 
            d.annotatedUser === pc.userEmail && 
            d.deviceType === 'Endpoint' &&
            d.pcType === pc.deviceType
          );
          
          if (!isDuplicate) {
            devices.push({
              deviceType: 'Endpoint',
              pcType: pc.deviceType,
              osInfo: pc.osInfo,
              deviceId: `Login-${pc.userEmail}-${pc.ipAddress}`,
              serialNumber: undefined,
              lastSync: pc.lastLogin,
              status: 'ACTIVE',
              annotatedUser: pc.userEmail,
              userEmail: pc.userEmail,
              ipAddress: pc.ipAddress,
              source: 'Login History',
              model: pc.deviceType
            });
          }
        });
      } catch (error) {
        console.log('ログイン履歴検出エラー:', error.message);
      }
    }
    
    console.log(`エンドポイント（Mac/Windows/Linux）: ${devices.length - endpointStartCount}台`);
  } catch (error) {
    console.log('エンドポイントの取得をスキップ:', error.message);
  }
  
  return devices;
}

/**
 * エンドポイント検証を使用したデバイス取得
 */
function getEndpointVerificationDevices() {
  const devices = [];
  
  try {
    const response = AdminDirectory.Devices.list({
      customer: 'my_customer',
      maxResults: 100,
      orderBy: 'lastSync',
      sortOrder: 'DESCENDING'
    });
    
    if (response.devices) {
      response.devices.forEach(device => {
        let deviceType = 'Unknown PC';
        let osInfo = device.osVersion || 'Unknown OS';
        
        if (device.os) {
          if (device.os.toLowerCase().includes('mac')) {
            deviceType = 'Mac';
            osInfo = `macOS ${device.osVersion || ''}`.trim();
          } else if (device.os.toLowerCase().includes('windows')) {
            deviceType = 'Windows PC';
            osInfo = `Windows ${device.osVersion || ''}`.trim();
          } else if (device.os.toLowerCase().includes('linux')) {
            deviceType = 'Linux PC';
            osInfo = `Linux ${device.osVersion || ''}`.trim();
          }
        }
        
        // IPアドレスの取得
        let ipAddress = 'N/A';
        if (device.lastKnownNetwork && device.lastKnownNetwork.length > 0) {
          ipAddress = device.lastKnownNetwork[0].ipAddress || 'N/A';
        } else if (device.networkInfo && device.networkInfo.ipAddress) {
          ipAddress = device.networkInfo.ipAddress;
        }
        
        devices.push({
          deviceType: 'Endpoint',
          pcType: deviceType,
          osInfo: osInfo,
          deviceId: device.deviceId,
          serialNumber: device.serialNumber,
          lastSync: device.lastSync,
          status: device.status,
          annotatedUser: device.annotatedUser,
          annotatedAssetId: device.annotatedAssetId,
          model: device.model,
          orgUnitPath: device.orgUnitPath,
          ipAddress: ipAddress,
          source: 'Endpoint Verification'
        });
      });
    }
  } catch (error) {
    // エンドポイント検証が利用できない場合は空配列を返す
  }
  
  return devices;
}

/**
 * デバイス情報を処理
 */
function processDevice(device, userEmail) {
  let email = '';
  let deviceData = [];
  
  // デバイスタイプ別に処理
  switch (device.deviceType) {
    case 'Chrome OS':
      email = device.annotatedUser || '';
      if (!email || email !== userEmail) return null;
      deviceData = processChromeOSDevice(device, email);
      break;
      
    case 'Mobile':
      email = device.email?.[0] || '';
      if (!email || email !== userEmail) return null;
      deviceData = processMobileDevice(device, email);
      break;
      
    case 'Browser':
      email = device.annotatedUser || '';
      if (!email || email !== userEmail) return null;
      deviceData = processBrowserDevice(device, email);
      break;
      
    case 'Endpoint':
      email = device.userEmail || device.annotatedUser || '';
      if (!email || email !== userEmail) return null;
      deviceData = processEndpointDevice(device, email);
      break;
      
    default:
      return null;
  }
  
  return deviceData;
}

/**
 * Chrome OSデバイスの処理
 */
function processChromeOSDevice(device, email) {
  // シリアル番号の取得
  let serialNumber = 'N/A';
  let identifier = '';
  
  if (device.serialNumber && device.serialNumber !== '') {
    serialNumber = device.serialNumber;
    identifier = `SN:${device.serialNumber}`;
  } else if (device.annotatedAssetId && device.annotatedAssetId !== '') {
    identifier = `Asset:${device.annotatedAssetId}`;
  } else if (device.deviceId && device.deviceId !== '') {
    identifier = `ID:${device.deviceId}`;
  } else if (device.macAddress && device.macAddress !== '') {
    identifier = `MAC:${device.macAddress}`;
  }
  
  return [
    'Chrome OS',
    device.model || 'Chrome OS Device',
    device.osVersion || 'Chrome OS',
    serialNumber,
    identifier,
    formatDate(device.lastSync),
    device.status || 'N/A',
    device.lastKnownNetwork?.[0]?.ipAddress || 'N/A',
    device.orgUnitPath || 'N/A'
  ];
}

/**
 * モバイルデバイスの処理
 */
function processMobileDevice(device, email) {
  // シリアル番号の詳細取得
  let serialNumber = 'N/A';
  let identifier = '';
  
  if (device.serialNumber && device.serialNumber !== '') {
    serialNumber = device.serialNumber;
    identifier = `SN:${device.serialNumber}`;
  }
  
  if (device.hardwareInfo && typeof device.hardwareInfo === 'object') {
    if (serialNumber === 'N/A' && device.hardwareInfo.serialNumber && device.hardwareInfo.serialNumber !== '') {
      serialNumber = device.hardwareInfo.serialNumber;
      identifier = `SN:${device.hardwareInfo.serialNumber}`;
    }
    
    if (serialNumber === 'N/A' && device.hardwareInfo.imei && device.hardwareInfo.imei !== '') {
      serialNumber = device.hardwareInfo.imei;
      identifier = `IMEI:${device.hardwareInfo.imei}`;
    }
    
    if (serialNumber === 'N/A' && device.hardwareInfo.meid && device.hardwareInfo.meid !== '') {
      serialNumber = device.hardwareInfo.meid;
      identifier = `MEID:${device.hardwareInfo.meid}`;
    }
    
    if (serialNumber === 'N/A' && device.hardwareInfo.esn && device.hardwareInfo.esn !== '') {
      serialNumber = device.hardwareInfo.esn;
      identifier = `ESN:${device.hardwareInfo.esn}`;
    }
  }
  
  if (serialNumber === 'N/A') {
    if (device.hardwareId && device.hardwareId !== '') {
      identifier = `HW:${device.hardwareId}`;
    } else if (device.deviceId && device.deviceId !== '') {
      identifier = `DEV:${device.deviceId}`;
    } else if (device.androidId && device.androidId !== '') {
      identifier = `ANDROID:${device.androidId}`;
    } else if (device.resourceId && device.resourceId !== '') {
      identifier = `RES:${device.resourceId}`;
    }
  }
  
  let modelName = device.model || 'Mobile Device';
  if (!device.model && device.hardwareInfo && device.hardwareInfo.model) {
    modelName = device.hardwareInfo.model;
  }
  
  let osInfo = device.os || 'Mobile OS';
  if (device.osVersion) {
    osInfo = `${osInfo} ${device.osVersion}`;
  }
  
  let deviceSubType = 'Mobile';
  if (device.type === 'ANDROID') {
    deviceSubType = 'Android';
  } else if (device.type === 'IOS' || device.type === 'IOS_SYNC') {
    deviceSubType = 'iOS';
  }
  
  return [
    deviceSubType,
    modelName,
    osInfo.trim(),
    serialNumber,
    identifier,
    formatDate(device.lastSync || device.firstSync),
    device.status || 'N/A',
    device.ipAddress || 'N/A',
    device.orgUnitPath || 'N/A'
  ];
}

/**
 * Chromeブラウザ（PC）の処理
 */
function processBrowserDevice(browser, email) {
  let serialNumber = 'N/A';
  let identifier = '';
  
  if (browser.machineId && browser.machineId !== '') {
    identifier = `Machine:${browser.machineId}`;
  } else if (browser.virtualDeviceId && browser.virtualDeviceId !== '') {
    identifier = `Virtual:${browser.virtualDeviceId}`;
  } else if (browser.deviceId && browser.deviceId !== '') {
    const shortId = browser.deviceId.length > 20 ? 
      browser.deviceId.substring(0, 20) + '...' : browser.deviceId;
    identifier = `ID:${shortId}`;
  }
  
  let osInfo = 'Unknown OS';
  let deviceModel = 'PC';
  
  if (browser.osVersion) {
    osInfo = browser.osVersion;
  } else if (browser.platformVersion) {
    osInfo = browser.platformVersion;
  }
  
  if (browser.osPlatform) {
    if (browser.osPlatform.toLowerCase().includes('mac')) {
      deviceModel = 'Mac';
      osInfo = osInfo.includes('mac') ? osInfo : `macOS ${osInfo}`;
    } else if (browser.osPlatform.toLowerCase().includes('win')) {
      deviceModel = 'Windows PC';
      osInfo = osInfo.includes('Windows') ? osInfo : `Windows ${osInfo}`;
    } else if (browser.osPlatform.toLowerCase().includes('linux')) {
      deviceModel = 'Linux PC';
      osInfo = osInfo.includes('Linux') ? osInfo : `Linux ${osInfo}`;
    }
  }
  
  let deviceId = browser.deviceId || browser.browserId || `Chrome-${email}`;
  if (deviceId.length > 50) {
    deviceId = deviceId.substring(0, 20) + '...';
  }
  
  return [
    `${deviceModel} (Browser)`,
    browser.browserVersion || 'Chrome Browser',
    osInfo,
    serialNumber,
    identifier,
    formatDate(browser.lastActivityTime || browser.lastPolicyFetchTime),
    browser.enrollmentState || 'ACTIVE',
    'N/A',
    browser.orgUnitPath || 'N/A'
  ];
}

/**
 * エンドポイント（Windows PC等）の処理
 */
function processEndpointDevice(endpoint, email) {
  let serialNumber = 'N/A';
  let identifier = '';
  
  if (endpoint.source === 'Endpoint Verification' && endpoint.serialNumber) {
    serialNumber = endpoint.serialNumber;
    identifier = `SN:${endpoint.serialNumber}`;
  } else if (endpoint.source === 'Login History') {
    identifier = `IP:${endpoint.ipAddress || 'Unknown'}`;
  }
  
  if (!identifier && endpoint.annotatedAssetId) {
    identifier = `Asset:${endpoint.annotatedAssetId}`;
  }
  
  let deviceModel = 'PC';
  let osInfo = endpoint.osInfo || 'Unknown OS';
  
  if (endpoint.pcType) {
    deviceModel = endpoint.pcType;
  } else if (endpoint.osInfo) {
    if (endpoint.osInfo.toLowerCase().includes('mac')) {
      deviceModel = 'Mac';
    } else if (endpoint.osInfo.toLowerCase().includes('windows')) {
      deviceModel = 'Windows PC';
    } else if (endpoint.osInfo.toLowerCase().includes('linux')) {
      deviceModel = 'Linux PC';
    }
  }
  
  let displayDeviceId = endpoint.deviceId || endpoint.ipAddress || 'N/A';
  
  if (displayDeviceId.includes('-') && endpoint.ipAddress) {
    displayDeviceId = endpoint.ipAddress;
  }
  
  return [
    `${deviceModel} (${endpoint.source || 'Endpoint'})`,
    deviceModel,
    osInfo,
    serialNumber,
    identifier,
    formatDate(endpoint.lastLogin || endpoint.lastSync),
    endpoint.status || 'ACTIVE',
    endpoint.ipAddress || 'N/A',
    endpoint.orgUnitPath || 'N/A'
  ];
}

/**
 * データをシートに書き込み
 */
function writeDataToSheet(sheet, data) {
  if (data.length === 0) return;
  
  const startRow = sheet.getLastRow() + 1;
  sheet.getRange(startRow, 1, data.length, data[0].length).setValues(data);
}

/**
 * シリアル番号の取得状況を分析
 */
function analyzeSerialNumbers() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(getConfig().sheets.devices);
  
  if (!sheet || sheet.getLastRow() <= 1) {
    console.log('データがありません。');
    return;
  }
  
  console.log('=== シリアル番号取得状況分析 ===\n');
  
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 9).getValues();
  
  const stats = {
    total: 0,
    hasSerialNumber: 0,
    byType: {},
    byIdentifierType: {}
  };
  
  data.forEach(row => {
    const deviceType = row[0];
    const serialNumber = row[3];
    const identifier = row[4];
    
    stats.total++;
    
    if (!stats.byType[deviceType]) {
      stats.byType[deviceType] = {
        total: 0,
        hasSerial: 0,
        identifiers: {}
      };
    }
    
    stats.byType[deviceType].total++;
    
    if (serialNumber && serialNumber !== 'N/A' && serialNumber !== '') {
      stats.hasSerialNumber++;
      stats.byType[deviceType].hasSerial++;
    }
    
    if (identifier && identifier !== '') {
      const identifierType = identifier.split(':')[0];
      stats.byIdentifierType[identifierType] = (stats.byIdentifierType[identifierType] || 0) + 1;
      
      if (!stats.byType[deviceType].identifiers[identifierType]) {
        stats.byType[deviceType].identifiers[identifierType] = 0;
      }
      stats.byType[deviceType].identifiers[identifierType]++;
    }
  });
  
  console.log(`総デバイス数: ${stats.total}`);
  console.log(`シリアル番号取得済み: ${stats.hasSerialNumber} (${Math.round(stats.hasSerialNumber/stats.total*100)}%)\n`);
  
  console.log('【デバイスタイプ別】');
  Object.entries(stats.byType).forEach(([type, typeStats]) => {
    console.log(`\n${type}: ${typeStats.total}台`);
    console.log(`  シリアル番号: ${typeStats.hasSerial}台 (${Math.round(typeStats.hasSerial/typeStats.total*100)}%)`);
    if (Object.keys(typeStats.identifiers).length > 0) {
      console.log('  識別子タイプ:');
      Object.entries(typeStats.identifiers).forEach(([idType, count]) => {
        console.log(`    - ${idType}: ${count}台`);
      });
    }
  });
}

/**
 * ユーザー情報を取得
 */
function getUserInfo(email) {
  let userInfo = { 
    name: { fullName: 'N/A' }, 
    primaryEmail: email,
    orgUnitPath: 'N/A'
  };
  
  try {
    userInfo = AdminDirectory.Users.get(email);
  } catch (error) {
    console.log(`ユーザー情報取得エラー (${email}): ${error.message}`);
  }
  
  return userInfo;
}

/**
 * ログイン履歴からPC/Macを検出（すべてのユーザー用）
 */
function detectPCsFromLoginHistory() {
  const pcMap = new Map();
  
  try {
    const endTime = new Date();
    const startTime = new Date(endTime.getTime() - 30 * 24 * 60 * 60 * 1000); // 過去30日
    
    // 組織全体のログイン履歴を取得
    const response = AdminReports.Activities.list('all', 'login', {
      startTime: startTime.toISOString(),
      endTime: endTime.toISOString(),
      maxResults: 1000,
      eventName: 'login_success'
    });
    
    if (response.items) {
      response.items.forEach(item => {
        const userEmail = item.actor?.email;
        if (!userEmail) return;
        
        let deviceInfo = {
          userEmail: userEmail,
          ipAddress: item.ipAddress || 'Unknown',
          lastLogin: item.id?.time,
          deviceType: 'Unknown PC',
          osInfo: 'Unknown OS'
        };
        
        // イベントパラメータからUser-Agentを取得してOSを推測
        if (item.events) {
          item.events.forEach(event => {
            if (event.parameters) {
              event.parameters.forEach(param => {
                if (param.name === 'user_agent' && param.value) {
                  const ua = param.value;
                  if (ua.includes('Macintosh') || ua.includes('Mac OS')) {
                    deviceInfo.deviceType = 'Mac';
                    deviceInfo.osInfo = 'macOS';
                  } else if (ua.includes('Windows')) {
                    deviceInfo.deviceType = 'Windows PC';
                    deviceInfo.osInfo = 'Windows';
                  } else if (ua.includes('Linux') && !ua.includes('Android')) {
                    deviceInfo.deviceType = 'Linux PC';
                    deviceInfo.osInfo = 'Linux';
                  } else if (ua.includes('CrOS')) {
                    deviceInfo.deviceType = 'Chrome OS';
                    deviceInfo.osInfo = 'Chrome OS';
                  }
                }
              });
            }
          });
        }
        
        // モバイルデバイスを除外してPCのみを記録
        if (!['Unknown PC'].includes(deviceInfo.deviceType) && 
            !deviceInfo.osInfo.includes('Android') && 
            !deviceInfo.osInfo.includes('iOS')) {
          const deviceKey = `${userEmail}-${deviceInfo.deviceType}-${deviceInfo.ipAddress}`;
          if (!pcMap.has(deviceKey) || 
              new Date(deviceInfo.lastLogin) > new Date(pcMap.get(deviceKey).lastLogin)) {
            pcMap.set(deviceKey, deviceInfo);
          }
        }
      });
    }
  } catch (error) {
    console.error('ログイン履歴取得エラー:', error);
  }
  
  return Array.from(pcMap.values());
}

/**
 * ログイン履歴からPC/Macを検出
 */
function detectPCsFromLoginHistory(userEmail) {
  const pcMap = new Map();
  
  try {
    const endTime = new Date();
    const startTime = new Date(endTime.getTime() - 30 * 24 * 60 * 60 * 1000); // 過去30日
    
    const response = AdminReports.Activities.list('all', 'login', {
      userKey: userEmail,
      startTime: startTime.toISOString(),
      endTime: endTime.toISOString(),
      maxResults: 100,
      eventName: 'login_success'
    });
    
    if (response.items) {
      response.items.forEach(item => {
        let deviceInfo = {
          userEmail: userEmail,
          ipAddress: item.ipAddress || 'Unknown',
          lastLogin: item.id?.time,
          deviceType: 'Unknown PC',
          osInfo: 'Unknown OS'
        };
        
        // イベントパラメータからUser-Agentを取得してOSを推測
        if (item.events) {
          item.events.forEach(event => {
            if (event.parameters) {
              event.parameters.forEach(param => {
                if (param.name === 'user_agent' && param.value) {
                  const ua = param.value;
                  if (ua.includes('Macintosh') || ua.includes('Mac OS')) {
                    deviceInfo.deviceType = 'Mac';
                    deviceInfo.osInfo = 'macOS';
                  } else if (ua.includes('Windows')) {
                    deviceInfo.deviceType = 'Windows PC';
                    deviceInfo.osInfo = 'Windows';
                  } else if (ua.includes('Linux') && !ua.includes('Android')) {
                    deviceInfo.deviceType = 'Linux PC';
                    deviceInfo.osInfo = 'Linux';
                  } else if (ua.includes('CrOS')) {
                    deviceInfo.deviceType = 'Chrome OS';
                    deviceInfo.osInfo = 'Chrome OS';
                  }
                }
              });
            }
          });
        }
        
        // モバイルデバイスを除外してPCのみを記録
        if (!['Unknown PC'].includes(deviceInfo.deviceType) && 
            !deviceInfo.osInfo.includes('Android') && 
            !deviceInfo.osInfo.includes('iOS')) {
          const deviceKey = `${deviceInfo.deviceType}-${deviceInfo.ipAddress}`;
          if (!pcMap.has(deviceKey) || 
              new Date(deviceInfo.lastLogin) > new Date(pcMap.get(deviceKey).lastLogin)) {
            pcMap.set(deviceKey, deviceInfo);
          }
        }
      });
    }
  } catch (error) {
    console.error('ログイン履歴取得エラー:', error);
  }
  
  return Array.from(pcMap.values());
}

/**
 * ユーザーが所有するカレンダーの一覧を取得
 */
function listUserCalendars(userEmail, systemConfig) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(systemConfig.sheets.calendars) 
      || SpreadsheetApp.getActiveSpreadsheet().insertSheet(systemConfig.sheets.calendars);
    sheet.clear();
    
    // カレンダー一覧のヘッダーを設定
    const headers = ["カレンダー名", "カレンダーID", "タイプ", "アクセス権限", "説明", "タイムゾーン", "共有ユーザー数"];
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length)
      .setBackground('#0F9D58')
      .setFontColor('#FFFFFF')
      .setFontWeight('bold');
    
    const allCalendars = [];
    const organizerEvents = [];
    
    // Calendar APIが利用可能かチェック
    if (typeof Calendar === 'undefined') {
      sheet.appendRow(["Calendar APIが有効になっていません", "", "", "", "", "", ""]);
      sheet.appendRow(["Apps Scriptエディタで「サービス」→「＋」→「Google Calendar API」を追加してください", "", "", "", "", "", ""]);
      
      // 代替方法：CalendarAppを使用（制限あり）
      sheet.appendRow(["", "", "", "", "", "", ""]);
      sheet.appendRow(["【CalendarAppを使用した基本情報】", "", "", "", "", "", ""]);
      
      try {
        // プライマリカレンダー
        const primaryCalendar = CalendarApp.getDefaultCalendar();
        allCalendars.push({
          name: primaryCalendar.getName(),
          id: primaryCalendar.getId(),
          type: "プライマリ",
          accessRole: "owner",
          description: "メインカレンダー",
          timeZone: primaryCalendar.getTimeZone(),
          sharedUserCount: "不明"
        });
        
        // 他のカレンダー（所有・購読）
        const calendars = CalendarApp.getAllCalendars();
        calendars.forEach(calendar => {
          if (calendar.getId() !== primaryCalendar.getId()) {
            allCalendars.push({
              name: calendar.getName(),
              id: calendar.getId(),
              type: "その他",
              accessRole: calendar.isMyPrimaryCalendar() ? "owner" : "reader",
              description: calendar.getDescription() || "説明なし",
              timeZone: calendar.getTimeZone(),
              sharedUserCount: "不明"
            });
          }
        });
      } catch (error) {
        console.error("CalendarApp使用エラー:", error);
      }
    } else {
      // Calendar APIを使用して詳細情報を取得
      try {
        console.log("Calendar APIを使用してカレンダー情報を取得中...");
        
        // カレンダーリストを取得
        const calendarList = Calendar.CalendarList.list({
          showDeleted: false,
          showHidden: false
        });
        
        if (calendarList.items) {
          calendarList.items.forEach(calendar => {
            // オーナーまたは編集権限があるカレンダーのみ
            if (calendar.accessRole === 'owner' || calendar.accessRole === 'writer') {
              // ACL（アクセス制御リスト）を取得して共有ユーザー数をカウント
              let sharedUserCount = 0;
              try {
                const acl = Calendar.Acl.list(calendar.id);
                if (acl.items) {
                  sharedUserCount = acl.items.filter(rule => 
                    rule.scope.type === 'user' && 
                    rule.scope.value !== userEmail
                  ).length;
                }
              } catch (aclError) {
                console.log("ACL取得エラー:", aclError.message);
              }
              
              allCalendars.push({
                name: calendar.summary,
                id: calendar.id,
                type: calendar.primary ? "プライマリ" : "セカンダリ",
                accessRole: calendar.accessRole,
                description: calendar.description || "説明なし",
                timeZone: calendar.timeZone,
                sharedUserCount: sharedUserCount
              });
            }
          });
        }
      } catch (error) {
        console.error("Calendar API使用エラー:", error);
        sheet.appendRow(["Calendar APIエラー: " + error.message, "", "", "", "", "", ""]);
      }
    }
    
    // 結果をシートに記録
    if (allCalendars.length > 0) {
      allCalendars.forEach(calendar => {
        sheet.appendRow([
          calendar.name,
          calendar.id,
          calendar.type,
          calendar.accessRole,
          calendar.description,
          calendar.timeZone,
          calendar.sharedUserCount
        ]);
      });
      
      // サマリー情報を追加
      const summaryRow = sheet.getLastRow() + 2;
      sheet.getRange(summaryRow, 1).setValue("【サマリー】").setFontWeight("bold");
      sheet.getRange(summaryRow + 1, 1).setValue("総カレンダー数:");
      sheet.getRange(summaryRow + 1, 2).setValue(allCalendars.length);
      
      // オーナーカレンダー数
      const ownerCount = allCalendars.filter(c => c.accessRole === 'owner').length;
      sheet.getRange(summaryRow + 2, 1).setValue("オーナーのカレンダー数:");
      sheet.getRange(summaryRow + 2, 2).setValue(ownerCount);
      
      // 共有されているカレンダー数
      const sharedCount = allCalendars.filter(c => c.sharedUserCount > 0).length;
      sheet.getRange(summaryRow + 3, 1).setValue("他ユーザーと共有中:");
      sheet.getRange(summaryRow + 3, 2).setValue(sharedCount);
      
      // 注意事項
      const noteRow = summaryRow + 5;
      sheet.getRange(noteRow, 1).setValue("【注意事項】").setFontWeight("bold").setFontColor("#FF0000");
      sheet.getRange(noteRow + 1, 1).setValue("※ オーナーのカレンダーは退職処理時に適切に移管または削除が必要です");
      sheet.getRange(noteRow + 2, 1).setValue("※ 共有カレンダーは他のユーザーへの影響を確認してください");
      
    } else {
      sheet.appendRow(["カレンダーは見つかりませんでした", "", "", "", "", "", ""]);
    }
    
    // 列幅の自動調整
    sheet.autoResizeColumns(1, headers.length);
    
    // イベント（予定）の検索
    sheet.appendRow(["", "", "", "", "", "", ""]);
    sheet.appendRow(["", "", "", "", "", "", ""]);
    
    // イベント一覧のヘッダー
    const eventHeaderRow = sheet.getLastRow() + 1;
    sheet.getRange(eventHeaderRow, 1).setValue("【主催している予定一覧】").setFontWeight("bold").setFontSize(12);
    
    const eventHeaders = ["予定タイトル", "開始日時", "終了日時", "場所", "参加者数", "繰り返し", "カレンダー名"];
    sheet.appendRow(eventHeaders);
    sheet.getRange(sheet.getLastRow(), 1, 1, eventHeaders.length)
      .setBackground('#4285F4')
      .setFontColor('#FFFFFF')
      .setFontWeight('bold');
    
    // 主催しているイベントを検索
    if (typeof Calendar !== 'undefined') {
      try {
        console.log("主催している予定を検索中...");
        
        // 今日から1年後までの期間で検索
        const now = new Date();
        const oneYearLater = new Date(now.getTime() + 365 * 24 * 60 * 60 * 1000);
        
        // すべてのカレンダーから主催イベントを取得
        allCalendars.forEach(calendar => {
          try {
            const events = Calendar.Events.list(calendar.id, {
              timeMin: now.toISOString(),
              timeMax: oneYearLater.toISOString(),
              singleEvents: true,
              orderBy: 'startTime',
              maxResults: 250,
              q: userEmail // 検索クエリでユーザーのメールを含むイベントを絞り込み
            });
            
            if (events.items) {
              events.items.forEach(event => {
                // 主催者かどうかをチェック
                if (event.organizer && event.organizer.email === userEmail) {
                  let attendeeCount = 0;
                  if (event.attendees) {
                    attendeeCount = event.attendees.filter(attendee => 
                      !attendee.resource && attendee.email !== userEmail
                    ).length;
                  }
                  
                  organizerEvents.push({
                    title: event.summary || "（タイトルなし）",
                    start: event.start.dateTime || event.start.date,
                    end: event.end.dateTime || event.end.date,
                    location: event.location || "なし",
                    attendeeCount: attendeeCount,
                    recurring: event.recurringEventId ? "繰り返しあり" : "なし",
                    calendarName: calendar.name,
                    calendarId: calendar.id,
                    eventId: event.id
                  });
                }
              });
            }
          } catch (eventError) {
            console.log(`カレンダー ${calendar.name} のイベント取得エラー:`, eventError.message);
          }
        });
        
        // イベントを日付順にソート
        organizerEvents.sort((a, b) => new Date(a.start) - new Date(b.start));
        
        // イベントをシートに記録
        if (organizerEvents.length > 0) {
          organizerEvents.forEach(event => {
            sheet.appendRow([
              event.title,
              formatDateTime(event.start),
              formatDateTime(event.end),
              event.location,
              event.attendeeCount,
              event.recurring,
              event.calendarName
            ]);
          });
          
          // イベントのサマリー
          const eventSummaryRow = sheet.getLastRow() + 2;
          sheet.getRange(eventSummaryRow, 1).setValue("【イベントサマリー】").setFontWeight("bold");
          sheet.getRange(eventSummaryRow + 1, 1).setValue("主催している予定数:");
          sheet.getRange(eventSummaryRow + 1, 2).setValue(organizerEvents.length);
          
          // 参加者が多いイベント
          const largeEvents = organizerEvents.filter(e => e.attendeeCount >= 5);
          sheet.getRange(eventSummaryRow + 2, 1).setValue("参加者5名以上の予定:");
          sheet.getRange(eventSummaryRow + 2, 2).setValue(largeEvents.length);
          
          // 繰り返しイベント
          const recurringEvents = organizerEvents.filter(e => e.recurring === "繰り返しあり");
          sheet.getRange(eventSummaryRow + 3, 1).setValue("繰り返し予定:");
          sheet.getRange(eventSummaryRow + 3, 2).setValue(recurringEvents.length);
          
          // イベントに関する注意事項
          const eventNoteRow = eventSummaryRow + 5;
          sheet.getRange(eventNoteRow, 1).setValue("【重要】").setFontWeight("bold").setFontColor("#FF0000");
          sheet.getRange(eventNoteRow + 1, 1).setValue("※ 主催している予定は適切に処理する必要があります:");
          sheet.getRange(eventNoteRow + 2, 1).setValue("  - 他の主催者への変更");
          sheet.getRange(eventNoteRow + 3, 1).setValue("  - 予定のキャンセル");
          sheet.getRange(eventNoteRow + 4, 1).setValue("  - 参加者への通知");
        } else {
          sheet.appendRow(["主催している予定はありません", "", "", "", "", "", ""]);
        }
        
      } catch (error) {
        console.error("イベント検索エラー:", error);
        sheet.appendRow(["イベント検索エラー: " + error.message, "", "", "", "", "", ""]);
      }
    } else {
      // CalendarAppを使用した基本的なイベント検索
      try {
        console.log("CalendarAppで主催イベントを検索中...");
        
        const now = new Date();
        const oneMonthLater = new Date(now.getTime() + 30 * 24 * 60 * 60 * 1000);
        
        allCalendars.forEach(calendarInfo => {
          try {
            const calendar = CalendarApp.getCalendarById(calendarInfo.id);
            if (calendar) {
              const events = calendar.getEvents(now, oneMonthLater);
              
              events.forEach(event => {
                // 作成者をチェック（CalendarAppでは正確な主催者情報が取得しにくい）
                const creators = event.getCreators();
                if (creators.includes(userEmail)) {
                  const guests = event.getGuestList();
                  const guestCount = guests.filter(guest => 
                    guest.getEmail() !== userEmail
                  ).length;
                  
                  organizerEvents.push({
                    title: event.getTitle() || "（タイトルなし）",
                    start: event.getStartTime(),
                    end: event.getEndTime(),
                    location: event.getLocation() || "なし",
                    attendeeCount: guestCount,
                    recurring: event.isRecurringEvent() ? "繰り返しあり" : "なし",
                    calendarName: calendarInfo.name
                  });
                }
              });
            }
          } catch (calError) {
            console.log(`カレンダー ${calendarInfo.name} のアクセスエラー:`, calError.message);
          }
        });
        
        // CalendarAppで取得したイベントを記録
        if (organizerEvents.length > 0) {
          organizerEvents.forEach(event => {
            sheet.appendRow([
              event.title,
              formatDateTime(event.start),
              formatDateTime(event.end),
              event.location,
              event.attendeeCount,
              event.recurring,
              event.calendarName
            ]);
          });
          
          sheet.appendRow(["", "", "", "", "", "", ""]);
          sheet.appendRow(["※ CalendarApp使用のため、今後30日間の予定のみ表示", "", "", "", "", "", ""]);
        } else {
          sheet.appendRow(["主催している予定はありません（今後30日間）", "", "", "", "", "", ""]);
        }
        
      } catch (error) {
        console.error("CalendarAppエラー:", error);
      }
    }
    
    console.log("カレンダー検索完了: " + allCalendars.length + "件、主催イベント: " + organizerEvents.length + "件");
    
  } catch (error) {
    console.error("カレンダー一覧取得エラー:", error);
    throw new Error("カレンダー一覧の取得でエラーが発生しました: " + error.message);
  }
}

/**
 * 日時をフォーマット
 */
function formatDateTime(dateValue) {
  if (!dateValue) return "N/A";
  
  try {
    const date = new Date(dateValue);
    if (isNaN(date.getTime())) return dateValue;
    
    const year = date.getFullYear();
    const month = (date.getMonth() + 1).toString().padStart(2, '0');
    const day = date.getDate().toString().padStart(2, '0');
    const hours = date.getHours().toString().padStart(2, '0');
    const minutes = date.getMinutes().toString().padStart(2, '0');
    
    // 時刻が00:00の場合は終日イベント
    if (hours === "00" && minutes === "00") {
      return `${year}/${month}/${day}`;
    }
    
    return `${year}/${month}/${day} ${hours}:${minutes}`;
  } catch (error) {
    return dateValue;
  }
}

/**
 * 詳細なSSO分析を実行
 */
function runDetailedSSOAnalysis() {
  try {
    const userConfig = getConfigFromSheet();
    const systemConfig = getConfig();
    
    if (!userConfig.userEmail) {
      SpreadsheetApp.getUi().alert("エラー", "退職者のメールアドレスを入力してください。", SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    listExternalSSOServices(userConfig.userEmail, systemConfig);
    
    // 分析結果をコンソールに出力
    const services = getDetailedExternalSSOServices(userConfig.userEmail, systemConfig.defaults.ssoLookbackDays || 365);
    
    console.log("=== 外部SSOサービス分析結果 ===");
    console.log("対象ユーザー:", userConfig.userEmail);
    console.log("総サービス数:", services.length);
    console.log("\n【主要サービスの利用状況】");
    
    const popularServices = systemConfig.defaults.popularSSOServices || [];
    popularServices.forEach(popularService => {
      const found = services.find(s => 
        s.name.toLowerCase().includes(popularService.toLowerCase())
      );
      if (found) {
        console.log("✓", popularService, "- 最終利用:", formatDate(found.lastUsed));
      }
    });
    
    SpreadsheetApp.getUi().alert("完了", "外部SSOサービスの詳細分析が完了しました。", SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (error) {
    SpreadsheetApp.getUi().alert("エラー", error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * 利用デバイス一覧のみ取得
 */
function runListUserDevices() {
  try {
    const userConfig = getConfigFromSheet();
    const systemConfig = getConfig();
    
    if (!userConfig.userEmail) {
      SpreadsheetApp.getUi().alert("エラー", "退職者のメールアドレスを入力してください。", SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    listUserDevices(userConfig.userEmail, systemConfig);
    SpreadsheetApp.getUi().alert("完了", "利用デバイス一覧を取得しました。", SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (error) {
    SpreadsheetApp.getUi().alert("エラー", error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * 退職者の総合レポートを生成
 */
function generateComprehensiveReport() {
  try {
    const userConfig = getConfigFromSheet();
    const systemConfig = getConfig();
    
    if (!userConfig.userEmail) {
      SpreadsheetApp.getUi().alert("エラー", "退職者のメールアドレスを入力してください。", SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(systemConfig.sheets.comprehensiveReport) 
      || SpreadsheetApp.getActiveSpreadsheet().insertSheet(systemConfig.sheets.comprehensiveReport);
    sheet.clear();
    
    // レポートタイトル
    sheet.getRange("A1").setValue("退職者総合レポート").setFontSize(20).setFontWeight("bold");
    sheet.getRange("A2").setValue("生成日時: " + formatDateTime(new Date())).setFontSize(10).setFontColor("#666666");
    sheet.getRange("A3").setValue("対象者: " + userConfig.userEmail).setFontSize(14).setFontWeight("bold");
    
    // ユーザー基本情報の取得
    let userInfo = {};
    try {
      userInfo = getUserInfo(userConfig.userEmail);
      sheet.getRange("A4").setValue("氏名: " + userInfo.name).setFontSize(12);
      sheet.getRange("A5").setValue("部署: " + userInfo.department).setFontSize(12);
    } catch (e) {
      console.log("ユーザー情報取得エラー:", e);
    }
    
    let row = 7;
    
    // サマリー情報
    sheet.getRange(row, 1).setValue("【エグゼクティブサマリー】").setFontSize(16).setFontWeight("bold").setBackground("#4285F4").setFontColor("#FFFFFF");
    row += 2;
    
    const services = getDetailedExternalSSOServices(userConfig.userEmail, systemConfig.defaults.ssoLookbackDays || 365);
    const deviceSummary = getDeviceSummary(userConfig.userEmail);
    const calendarSummary = getCalendarSummary(userConfig.userEmail);
    const emailSummary = getEmailSummary(userConfig.userEmail);
    const fileSummary = getFileSummary(userConfig.userEmail);
    
    // サマリーテーブル
    const summaryData = [
      ["項目", "件数", "重要度"],
      ["外部連携サービス", services.length + "個", services.length > 10 ? "高" : "中"],
      ["利用デバイス", deviceSummary.totalCount + "台", deviceSummary.totalCount > 3 ? "高" : "低"],
      ["所有カレンダー", calendarSummary.ownedCount + "個", calendarSummary.ownedCount > 2 ? "中" : "低"],
      ["主催予定", calendarSummary.organizerEventCount + "件", calendarSummary.organizerEventCount > 10 ? "高" : "中"],
      ["削除対象メール", emailSummary.deletedCount + "件", emailSummary.deletedCount > 0 ? "要確認" : "-"],
      ["所有ファイル", fileSummary.totalFiles + "個", fileSummary.totalFiles > 100 ? "高" : "中"]
    ];
    
    summaryData.forEach((rowData, index) => {
      rowData.forEach((cellData, colIndex) => {
        const cell = sheet.getRange(row + index, colIndex + 1);
        cell.setValue(cellData);
        if (index === 0) {
          cell.setFontWeight("bold").setBackground("#E8E8E8");
        }
        if (colIndex === 2 && cellData === "高") {
          cell.setFontColor("#D50000");
        }
      });
    });
    
    row += summaryData.length + 2;
    
    // 外部サービス詳細情報
    sheet.getRange(row, 1).setValue("【外部連携サービス詳細】").setFontSize(16).setFontWeight("bold").setBackground("#EA4335").setFontColor("#FFFFFF");
    row += 2;
    
    // カテゴリ別集計
    const serviceCategories = {};
    services.forEach(service => {
      const category = service.type || "その他";
      if (!serviceCategories[category]) {
        serviceCategories[category] = [];
      }
      serviceCategories[category].push(service);
    });
    
    // カテゴリ別表示
    Object.keys(serviceCategories).sort().forEach(category => {
      sheet.getRange(row, 1).setValue("■ " + category + " (" + serviceCategories[category].length + "個)").setFontWeight("bold").setFontSize(12);
      row++;
      
      serviceCategories[category].slice(0, 10).forEach(service => {
        sheet.getRange(row, 2).setValue("• " + service.name);
        sheet.getRange(row, 3).setValue("最終利用: " + formatDate(service.lastAccess));
        sheet.getRange(row, 4).setValue("利用回数: " + service.eventCount);
        row++;
      });
      
      if (serviceCategories[category].length > 10) {
        sheet.getRange(row, 2).setValue("... 他 " + (serviceCategories[category].length - 10) + " サービス").setFontStyle("italic");
        row++;
      }
      row++;
    });
    
    // 重要な外部サービス（頻繁に利用）
    sheet.getRange(row, 1).setValue("【特に注意が必要な外部サービス（利用頻度高）】").setFontWeight("bold").setBackground("#FCE8B2");
    row += 2;
    
    const frequentServices = services
      .filter(s => s.eventCount > 10)
      .sort((a, b) => b.eventCount - a.eventCount)
      .slice(0, 10);
    
    if (frequentServices.length > 0) {
      frequentServices.forEach(service => {
        sheet.getRange(row, 1).setValue("⚠️ " + service.name);
        sheet.getRange(row, 2).setValue("利用回数: " + service.eventCount + "回");
        sheet.getRange(row, 3).setValue("タイプ: " + service.type);
        row++;
      });
    } else {
      sheet.getRange(row, 1).setValue("頻繁に利用されているサービスはありません");
      row++;
    }
    
    row += 2;
    
    // デバイス情報
    sheet.getRange(row, 1).setValue("【デバイス利用状況】").setFontSize(16).setFontWeight("bold").setBackground("#34A853").setFontColor("#FFFFFF");
    row += 2;
    
    const deviceTypes = {
      "Chrome OS": deviceSummary.chromeOS || 0,
      "モバイル": deviceSummary.mobile || 0,
      "Chromeブラウザ": deviceSummary.browser || 0,
      "PC/Mac": deviceSummary.pc || 0
    };
    
    Object.entries(deviceTypes).forEach(([type, count]) => {
      if (count > 0) {
        sheet.getRange(row, 1).setValue(type + ":");
        sheet.getRange(row, 2).setValue(count + "台");
        row++;
      }
    });
    
    if (deviceSummary.activeDevices && deviceSummary.activeDevices.length > 0) {
      row++;
      sheet.getRange(row, 1).setValue("最近利用されたデバイス:").setFontWeight("bold");
      row++;
      deviceSummary.activeDevices.slice(0, 5).forEach(device => {
        sheet.getRange(row, 2).setValue("• " + device.model + " - " + formatDate(device.lastSync));
        row++;
      });
    }
    
    row += 2;
    
    // カレンダー・予定情報
    sheet.getRange(row, 1).setValue("【カレンダー・予定情報】").setFontSize(16).setFontWeight("bold").setBackground("#FBBC04").setFontColor("#000000");
    row += 2;
    
    sheet.getRange(row, 1).setValue("所有カレンダー数:");
    sheet.getRange(row, 2).setValue(calendarSummary.ownedCount);
    row++;
    sheet.getRange(row, 1).setValue("主催している予定数:");
    sheet.getRange(row, 2).setValue(calendarSummary.organizerEventCount);
    row++;
    
    if (calendarSummary.importantEvents.length > 0) {
      row++;
      sheet.getRange(row, 1).setValue("【重要】参加者5名以上の予定:").setFontWeight("bold").setFontColor("#D50000");
      row++;
      calendarSummary.importantEvents.forEach(event => {
        sheet.getRange(row, 2).setValue("• " + event);
        row++;
      });
    }
    
    row += 2;
    
    // ファイル・データ情報
    sheet.getRange(row, 1).setValue("【ファイル・データ所有状況】").setFontSize(16).setFontWeight("bold").setBackground("#4285F4").setFontColor("#FFFFFF");
    row += 2;
    
    if (fileSummary.byType) {
      Object.entries(fileSummary.byType).forEach(([type, count]) => {
        if (count > 0) {
          sheet.getRange(row, 1).setValue(type + ":");
          sheet.getRange(row, 2).setValue(count + "個");
          row++;
        }
      });
    }
    
    row += 2;
    
    // メール処理状況
    sheet.getRange(row, 1).setValue("【メール処理状況】").setFontSize(16).setFontWeight("bold").setBackground("#EA4335").setFontColor("#FFFFFF");
    row += 2;
    
    sheet.getRange(row, 1).setValue("削除対象メール:");
    sheet.getRange(row, 2).setValue(emailSummary.deletedCount + "件");
    row++;
    
    if (emailSummary.keywords && emailSummary.keywords.length > 0) {
      sheet.getRange(row, 1).setValue("削除キーワード:");
      sheet.getRange(row, 2).setValue(emailSummary.keywords.join(", "));
      row++;
    }
    
    row += 2;
    
    // 推奨アクション
    sheet.getRange(row, 1).setValue("【推奨アクション】").setFontSize(16).setFontWeight("bold").setBackground("#FCE8B2");
    row += 2;
    
    // 優先度別アクション
    const highPriorityActions = [];
    const mediumPriorityActions = [];
    const lowPriorityActions = [];
    
    // 外部サービス関連
    if (services.length > 10) {
      highPriorityActions.push("【緊急】" + services.length + "個の外部サービスのアカウント無効化が必要");
    } else if (services.length > 0) {
      mediumPriorityActions.push(services.length + "個の外部サービスのアカウント確認と無効化");
    }
    
    // 特に重要な外部サービス
    if (frequentServices.length > 0) {
      highPriorityActions.push("【重要】頻繁利用サービス(" + frequentServices.slice(0, 3).map(s => s.name).join(", ") + ")の早急な対応");
    }
    
    // 自動化・連携関連の検出
    const automationServices = services.filter(s => 
      s.name.toLowerCase().includes('zapier') ||
      s.name.toLowerCase().includes('ifttt') ||
      s.name.toLowerCase().includes('make') ||
      s.name.toLowerCase().includes('integromat') ||
      s.name.toLowerCase().includes('power automate') ||
      s.name.toLowerCase().includes('workflow') ||
      s.type === "自動化" ||
      s.type === "連携ツール"
    );
    
    if (automationServices.length > 0) {
      highPriorityActions.push("【重要】自動化ツール(" + automationServices.slice(0, 3).map(s => s.name).join(", ") + ")の設定確認と移管");
    }
    
    // GASプロジェクト関連
    highPriorityActions.push("【重要】Google Apps Scriptプロジェクトの確認と移管");
    mediumPriorityActions.push("GASのトリガー設定の確認と必要に応じた無効化");
    mediumPriorityActions.push("GASプロジェクトの所有権移管または共同編集者の追加");
    
    // API連携関連
    const apiServices = services.filter(s => 
      s.name.toLowerCase().includes('api') ||
      s.type === "開発ツール" ||
      s.type === "API管理"
    );
    
    if (apiServices.length > 0) {
      highPriorityActions.push("【API連携】" + apiServices.length + "個のAPI連携サービスの認証情報確認");
      mediumPriorityActions.push("APIキーやトークンの再発行と更新");
    }
    
    // カレンダー関連
    if (calendarSummary.organizerEventCount > 10) {
      highPriorityActions.push("【緊急】" + calendarSummary.organizerEventCount + "件の主催予定の処理");
    } else if (calendarSummary.organizerEventCount > 0) {
      mediumPriorityActions.push(calendarSummary.organizerEventCount + "件の主催予定の引き継ぎ");
    }
    
    // デバイス関連
    if (deviceSummary.totalCount > 3) {
      mediumPriorityActions.push(deviceSummary.totalCount + "台のデバイスのリモートワイプ検討");
    }
    
    // Webhook・自動化の確認
    mediumPriorityActions.push("Webhook URLの確認と必要に応じた更新");
    mediumPriorityActions.push("自動化ワークフローの実行者変更");
    
    // 標準アクション
    mediumPriorityActions.push("Google Workspaceアカウントの無効化設定");
    mediumPriorityActions.push("メールルーティングの設定（管理コンソール）");
    mediumPriorityActions.push("OAuth認証を使用しているサービスの再認証");
    lowPriorityActions.push("ファイル・データの所有権移管");
    lowPriorityActions.push("共有ドライブへのファイル移動");
    lowPriorityActions.push("スプレッドシートの自動更新スクリプトの確認");
    
    // アクション表示
    sheet.getRange(row, 1).setValue("🔴 優先度：高").setFontWeight("bold").setFontColor("#D50000");
    row++;
    highPriorityActions.forEach((action, index) => {
      sheet.getRange(row, 1).setValue((index + 1) + ". " + action);
      row++;
    });
    
    row++;
    sheet.getRange(row, 1).setValue("🟡 優先度：中").setFontWeight("bold").setFontColor("#F57C00");
    row++;
    mediumPriorityActions.forEach((action, index) => {
      sheet.getRange(row, 1).setValue((index + 1) + ". " + action);
      row++;
    });
    
    row++;
    sheet.getRange(row, 1).setValue("🟢 優先度：低").setFontWeight("bold").setFontColor("#0F9D58");
    row++;
    lowPriorityActions.forEach((action, index) => {
      sheet.getRange(row, 1).setValue((index + 1) + ". " + action);
      row++;
    });
    
    row += 2;
    
    // 詳細な外部サービスリスト（付録）
    sheet.getRange(row, 1).setValue("【付録：外部サービス完全リスト】").setFontSize(14).setFontWeight("bold").setBackground("#E8E8E8");
    row += 2;
    
    if (services.length > 10) {
      // ヘッダー
      sheet.getRange(row, 1).setValue("サービス名").setFontWeight("bold");
      sheet.getRange(row, 2).setValue("カテゴリ").setFontWeight("bold");
      sheet.getRange(row, 3).setValue("最終利用").setFontWeight("bold");
      sheet.getRange(row, 4).setValue("利用回数").setFontWeight("bold");
      row++;
      
      services.forEach(service => {
        sheet.getRange(row, 1).setValue(service.name);
        sheet.getRange(row, 2).setValue(service.type);
        sheet.getRange(row, 3).setValue(formatDate(service.lastAccess));
        sheet.getRange(row, 4).setValue(service.eventCount);
        row++;
      });
    }
    
    // 列幅調整
    sheet.setColumnWidth(1, 350);
    sheet.setColumnWidth(2, 250);
    sheet.setColumnWidth(3, 150);
    sheet.setColumnWidth(4, 100);
    
    // 条件付き書式設定
    const lastRow = sheet.getLastRow();
    const range = sheet.getRange(1, 1, lastRow, 4);
    
    SpreadsheetApp.getUi().alert("完了", "退職者総合レポートを生成しました。\n\n外部連携サービス: " + services.length + "個を検出しました。", SpreadsheetApp.getUi().ButtonSet.OK);
    
  } catch (error) {
    SpreadsheetApp.getUi().alert("エラー", error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * デバイスのサマリー情報を取得
 */
function getDeviceSummary(userEmail) {
  const summary = {
    totalCount: 0,
    chromeOS: 0,
    mobile: 0,
    browser: 0,
    pc: 0,
    activeDevices: []
  };
  
  try {
    const devices = getAllDevicesIncludingPC();
    summary.totalCount = devices.length;
    
    devices.forEach(device => {
      switch (device.deviceType) {
        case 'chromeOS':
          summary.chromeOS++;
          break;
        case 'mobile':
          summary.mobile++;
          break;
        case 'browser':
          summary.browser++;
          break;
        case 'pc':
        case 'endpoint':
          summary.pc++;
          break;
      }
      
      // 最近同期されたデバイス
      if (device.lastSync && new Date(device.lastSync) > new Date(Date.now() - 30 * 24 * 60 * 60 * 1000)) {
        summary.activeDevices.push({
          model: device.model || device.deviceModel || device.deviceId,
          lastSync: device.lastSync
        });
      }
    });
    
    // 最新のものから5つまで
    summary.activeDevices.sort((a, b) => new Date(b.lastSync) - new Date(a.lastSync));
    summary.activeDevices = summary.activeDevices.slice(0, 5);
    
  } catch (error) {
    console.error("デバイスサマリー取得エラー:", error);
  }
  
  return summary;
}

/**
 * メールのサマリー情報を取得
 */
function getEmailSummary(userEmail) {
  const summary = {
    deletedCount: 0,
    keywords: []
  };
  
  try {
    const configSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("設定");
    if (configSheet) {
      const keywords = configSheet.getRange("C6").getValue();
      if (keywords) {
        summary.keywords = keywords.split(",").map(k => k.trim());
      }
      
      // 削除メール一覧シートから件数を取得
      const deletedSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("削除メール一覧");
      if (deletedSheet && deletedSheet.getLastRow() > 1) {
        summary.deletedCount = deletedSheet.getLastRow() - 1; // ヘッダーを除く
      }
    }
  } catch (error) {
    console.error("メールサマリー取得エラー:", error);
  }
  
  return summary;
}

/**
 * ファイルのサマリー情報を取得
 */
function getFileSummary(userEmail) {
  const summary = {
    totalFiles: 0,
    byType: {}
  };
  
  try {
    // データ一覧シートから情報を取得
    const dataSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("データ一覧");
    if (dataSheet && dataSheet.getLastRow() > 1) {
      const dataRange = dataSheet.getRange(2, 2, dataSheet.getLastRow() - 1, 1).getValues();
      
      dataRange.forEach(row => {
        const fileType = row[0];
        if (fileType) {
          summary.totalFiles++;
          summary.byType[fileType] = (summary.byType[fileType] || 0) + 1;
        }
      });
    }
  } catch (error) {
    console.error("ファイルサマリー取得エラー:", error);
  }
  
  return summary;
}

/**
 * カレンダーのサマリー情報を取得
 */
function getCalendarSummary(userEmail) {
  const summary = {
    ownedCount: 0,
    organizerEventCount: 0,
    importantEvents: []
  };
  
  try {
    // CalendarAppを使用して基本情報を取得
    const calendars = CalendarApp.getAllCalendars();
    calendars.forEach(calendar => {
      try {
        // プライマリカレンダーかどうかチェック
        if (calendar.isMyPrimaryCalendar() || calendar.getId().includes(userEmail)) {
          summary.ownedCount++;
        }
      } catch (e) {
        // アクセス権限エラーは無視
      }
    });
    
    // 今後30日間の主催イベントを確認
    const now = new Date();
    const oneMonthLater = new Date(now.getTime() + 30 * 24 * 60 * 60 * 1000);
    
    const primaryCalendar = CalendarApp.getDefaultCalendar();
    const events = primaryCalendar.getEvents(now, oneMonthLater);
    
    events.forEach(event => {
      const creators = event.getCreators();
      if (creators.includes(userEmail)) {
        summary.organizerEventCount++;
        
        const guests = event.getGuestList();
        const guestCount = guests.filter(guest => guest.getEmail() !== userEmail).length;
        
        if (guestCount >= 5) {
          summary.importantEvents.push(
            formatDate(event.getStartTime()) + " " + event.getTitle() + " (参加者: " + guestCount + "名)"
          );
        }
      }
    });
    
  } catch (error) {
    console.error("カレンダーサマリー取得エラー:", error);
  }
  
  return summary;
}

// ===== メールルーティング関連 =====

/**
 * 管理コンソールでのルーティング設定案内
 */
function showAdminRoutingGuide() {
  const config = getConfigFromSheet();
  const ui = SpreadsheetApp.getUi();
  
  const guideMessage = "【Google Workspace管理コンソールでのメールルーティング設定】\n\n" +
    "退職者のメールを組織レベルで転送するには、\n" +
    "管理コンソールでの設定が必要です。\n\n" +
    "◆ 設定情報\n" +
    "退職者: " + (config.userEmail || '（未設定）') + "\n" +
    "転送先: " + config.forwardEmail + "\n\n" +
    "◆ 設定手順\n" +
    "1. admin.google.com にアクセス\n" +
    "2. アプリ → Google Workspace → Gmail → ルーティング\n" +
    "3. 「設定」または「別のルートを追加」をクリック\n" +
    "4. 以下を設定：\n" +
    "   \n" +
    "   【1. 影響を受けるメール】\n" +
    "   - エンベロープ受信者: 一致タイプ「単一の受信者」\n" +
    "   - アドレス: " + (config.userEmail || '退職者のメールアドレス') + "\n" +
    "   \n" +
    "   【2. 上記のタイプのメッセージに対する処理】\n" +
    "   - ☑ エンベロープ受信者を変更\n" +
    "   - 受信者のアドレスを次に変更: " + config.forwardEmail + "\n" +
    "   - ☑ メッセージも配信する（元のメールも保持）\n" +
    "   \n" +
    "   【3. オプション】\n" +
    "   - ☑ スパムとして認識されたメッセージも影響を受ける\n\n" +
    "5. 「設定を追加」をクリック\n\n" +
    "◆ この設定のメリット\n" +
    "✓ 組織レベルでの確実な転送\n" +
    "✓ ユーザーが設定を変更不可\n" +
    "✓ 管理コンソールで一元管理\n" +
    "✓ 監査ログで追跡可能\n\n" +
    "◆ 代替案\n" +
    "個人レベルの転送設定を使用する場合は、\n" +
    "メニュー「メールルーティング」→「個人転送設定」を選択";
  
  ui.alert("管理コンソール設定ガイド", guideMessage, ui.ButtonSet.OK);
  
  // CSV出力の提案
  const csvResponse = ui.alert(
    "CSV出力",
    "この設定情報をCSVファイルとして出力しますか？\n" +
    "管理コンソールでの一括設定に使用できます。",
    ui.ButtonSet.YES_NO
  );
  
  if (csvResponse === ui.Button.YES) {
    exportRoutingConfigCSV();
  }
}

/**
 * 個人レベルの転送設定（ユーザーレベル）
 */
function runEmailRouting() {
  try {
    const userConfig = getConfigFromSheet();
    if (!userConfig.userEmail) {
      SpreadsheetApp.getUi().alert("エラー", "退職者のメールアドレスを入力してください。", SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    setupUserLevelForwarding(userConfig);
  } catch (error) {
    SpreadsheetApp.getUi().alert("エラー", error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * ユーザーレベルの転送設定
 */
function setupUserLevelForwarding(userConfig) {
  try {
    if (typeof Gmail === 'undefined') {
      throw new Error("Gmail APIが有効になっていません。Apps Scriptエディタで「サービス」→「＋」→「Gmail API」を追加してください。");
    }
    
    const ui = SpreadsheetApp.getUi();
    
    // 個人レベル転送の説明
    const response = ui.alert(
      "個人レベルの転送設定",
      "【注意】これは個人のGmail設定での転送です。\n" +
      "管理コンソールでの設定を推奨します。\n\n" +
      "対象: " + userConfig.userEmail + "\n" +
      "転送先: " + userConfig.forwardEmail + "\n\n" +
      "この設定により:\n" +
      "• 今後のメールが自動転送されます\n" +
      "• 元のメールも受信トレイに残ります\n" +
      "• ユーザーが設定を変更可能です\n\n" +
      "続行しますか？",
      ui.ButtonSet.YES_NO
    );
    
    if (response !== ui.Button.YES) {
      return;
    }
    
    // 転送先アドレスを追加
    Gmail.Users.Settings.ForwardingAddresses.create(
      { forwardingEmail: userConfig.forwardEmail }, 
      userConfig.userEmail
    );
    
    // 自動転送を有効化
    Gmail.Users.Settings.updateAutoForwarding({
      enabled: true,
      emailAddress: userConfig.forwardEmail,
      disposition: 'leaveInInbox'
    }, userConfig.userEmail);
    
    // ログに記録
    logForwardingSetup(userConfig, "個人レベル転送");
    
    ui.alert(
      "設定完了",
      "個人レベルの転送設定が完了しました。\n\n" +
      "より確実な転送のため、管理コンソールでの\n" +
      "ルーティング設定も検討してください。",
      ui.ButtonSet.OK
    );
    
  } catch (error) {
    console.error("転送設定エラー:", error);
    throw new Error("転送設定でエラーが発生しました: " + error.message);
  }
}

/**
 * メール委任設定
 */
function runEmailDelegation() {
  try {
    const userConfig = getConfigFromSheet();
    if (!userConfig.userEmail) {
      SpreadsheetApp.getUi().alert("エラー", "退職者のメールアドレスを入力してください。", SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      "メール委任設定",
      userConfig.forwardEmail + "に\n" +
      userConfig.userEmail + "のメールボックスへの\n" +
      "アクセス権限を付与しますか？\n\n" +
      "これにより:\n" +
      "• 過去のメールも含めて全てアクセス可能\n" +
      "• 委任先から送信も可能\n" +
      "• 監査ログで追跡可能",
      ui.ButtonSet.YES_NO
    );
    
    if (response !== ui.Button.YES) {
      return;
    }
    
    if (typeof Gmail === 'undefined') {
      throw new Error("Gmail APIが有効になっていません");
    }
    
    // メールの委任を追加
    Gmail.Users.Settings.Delegates.create({
      delegateEmail: userConfig.forwardEmail,
      verificationStatus: 'accepted'
    }, userConfig.userEmail);
    
    // ログに記録
    logForwardingSetup(userConfig, "メール委任");
    
    ui.alert("完了", "メール委任設定が完了しました。", ui.ButtonSet.OK);
    
  } catch (error) {
    SpreadsheetApp.getUi().alert("エラー", error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * ルーティング設定用CSVエクスポート
 */
function exportRoutingConfigCSV() {
  const config = getConfig();
  const userConfig = getConfigFromSheet();
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  let csvSheet = spreadsheet.getSheetByName(config.sheets.routingCSV);
  if (!csvSheet) {
    csvSheet = spreadsheet.insertSheet(config.sheets.routingCSV);
  }
  
  csvSheet.clear();
  
  // ヘッダー
  csvSheet.appendRow([
    "退職者メール",
    "転送先",
    "設定タイプ",
    "アクション",
    "オプション",
    "処理日"
  ]);
  
  // データ行
  csvSheet.appendRow([
    userConfig.userEmail,
    userConfig.forwardEmail,
    "エンベロープ受信者",
    "アドレスを変更",
    "メッセージも配信",
    new Date().toLocaleDateString()
  ]);
  
  // 管理コンソール用の詳細設定
  csvSheet.appendRow([]);
  csvSheet.appendRow(["管理コンソール設定用詳細情報"]);
  csvSheet.appendRow(["項目", "設定値"]);
  csvSheet.appendRow(["影響を受けるメール - タイプ", "エンベロープ受信者"]);
  csvSheet.appendRow(["影響を受けるメール - 一致タイプ", "単一の受信者"]);
  csvSheet.appendRow(["影響を受けるメール - アドレス", userConfig.userEmail]);
  csvSheet.appendRow(["アクション - エンベロープ受信者を変更", "有効"]);
  csvSheet.appendRow(["アクション - 変更先アドレス", userConfig.forwardEmail]);
  csvSheet.appendRow(["アクション - メッセージも配信", "有効"]);
  csvSheet.appendRow(["オプション - スパムも含む", "有効"]);
  
  // 使用方法
  csvSheet.appendRow([]);
  csvSheet.appendRow(["使用方法："]);
  csvSheet.appendRow(["1. このデータを参照して管理コンソールで設定"]);
  csvSheet.appendRow(["2. 複数の退職者がいる場合は行を追加"]);
  csvSheet.appendRow(["3. Google Workspace Admin APIでの一括処理も可能"]);
  
  // 列幅調整
  csvSheet.autoResizeColumns(1, 6);
  
  SpreadsheetApp.getUi().alert(
    "CSV出力完了",
    "「" + config.sheets.routingCSV + "」シートを確認してください。\n\n" +
    "このデータを管理コンソールでの設定に使用できます。",
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * 現在の転送設定を確認
 */
function runCheckForwarding() {
  try {
    const userConfig = getConfigFromSheet();
    if (!userConfig.userEmail) {
      SpreadsheetApp.getUi().alert("エラー", "退職者のメールアドレスを入力してください。", SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    if (typeof Gmail === 'undefined') {
      throw new Error("Gmail APIが有効になっていません");
    }
    
    // 現在の転送設定を取得
    const forwardingSettings = Gmail.Users.Settings.getAutoForwarding(userConfig.userEmail);
    const forwardingAddresses = Gmail.Users.Settings.ForwardingAddresses.list(userConfig.userEmail);
    
    let message = "【" + userConfig.userEmail + "の転送設定】\n\n";
    
    // 個人レベルの転送設定
    message += "◆ 個人レベルの転送設定\n";
    if (forwardingSettings && forwardingSettings.enabled) {
      message += "状態: 有効\n";
      message += "転送先: " + forwardingSettings.emailAddress + "\n";
      message += "処理: " + (forwardingSettings.disposition === 'leaveInInbox' ? 'コピーを転送' : '転送後に削除') + "\n";
    } else {
      message += "状態: 無効\n";
    }
    
    // 登録済み転送先
    if (forwardingAddresses && 
        forwardingAddresses.forwardingAddresses && 
        Array.isArray(forwardingAddresses.forwardingAddresses) && 
        forwardingAddresses.forwardingAddresses.length > 0) {
      message += "\n◆ 登録済み転送先:\n";
      forwardingAddresses.forwardingAddresses.forEach(addr => {
        message += "• " + addr.forwardingEmail + " (" + addr.verificationStatus + ")\n";
      });
    }
    
    message += "\n◆ 管理コンソールの設定\n";
    message += "管理コンソールでのルーティング設定は\n";
    message += "admin.google.com で確認してください。";
    
    SpreadsheetApp.getUi().alert("転送設定の確認", message, SpreadsheetApp.getUi().ButtonSet.OK);
    
  } catch (error) {
    SpreadsheetApp.getUi().alert("エラー", error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * 転送を無効化
 */
function runDisableForwarding() {
  try {
    const userConfig = getConfigFromSheet();
    if (!userConfig.userEmail) {
      SpreadsheetApp.getUi().alert("エラー", "退職者のメールアドレスを入力してください。", SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    if (typeof Gmail === 'undefined') {
      throw new Error("Gmail APIが有効になっていません");
    }
    
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      "転送設定の無効化",
      userConfig.userEmail + "の個人レベル転送を無効化しますか？\n\n" +
      "※ 管理コンソールの設定は影響を受けません",
      ui.ButtonSet.YES_NO
    );
    
    if (response !== ui.Button.YES) {
      return;
    }
    
    // 自動転送を無効化
    Gmail.Users.Settings.updateAutoForwarding({
      enabled: false
    }, userConfig.userEmail);
    
    ui.alert("完了", "個人レベルの転送を無効化しました。", ui.ButtonSet.OK);
    
  } catch (error) {
    SpreadsheetApp.getUi().alert("エラー", error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * 転送設定のログ記録
 */
function logForwardingSetup(userConfig, type) {
  const config = getConfig();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(config.sheets.forwardingLog) 
    || SpreadsheetApp.getActiveSpreadsheet().insertSheet(config.sheets.forwardingLog);
  
  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      "設定日時",
      "対象アドレス",
      "転送先/委任先",
      "設定タイプ",
      "設定者"
    ]);
  }
  
  sheet.appendRow([
    new Date(),
          userConfig.userEmail,
      userConfig.forwardEmail,
      type,
      getActiveUserEmail()
  ]);
  
  sheet.autoResizeColumns(1, 5);
}

// ===== 個別機能 =====

/**
 * メール削除のみ実行
 */
function runDeleteEmails() {
  try {
    const userConfig = getConfigFromSheet();
    const systemConfig = getConfig();
    
    if (!userConfig.userEmail) {
      SpreadsheetApp.getUi().alert("エラー", "退職者のメールアドレスを入力してください。", SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    const deleted = deleteSensitiveEmails(userConfig, systemConfig);
    SpreadsheetApp.getUi().alert("完了", deleted + "件のメールを削除しました。", SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (error) {
    SpreadsheetApp.getUi().alert("エラー", error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * センシティブメールの抽出のみ実行（削除前の確認用）
 */
function runExtractSensitiveEmails() {
  try {
    const userConfig = getConfigFromSheet();
    const systemConfig = getConfig();
    
    if (!userConfig.userEmail) {
      SpreadsheetApp.getUi().alert("エラー", "退職者のメールアドレスを入力してください。", SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    const sheet = createSensitiveEmailsSheet(userConfig, systemConfig);
    if (sheet) {
      const emailCount = sheet.getLastRow() - 1;
      SpreadsheetApp.getUi().alert(
        "抽出完了", 
        emailCount + "件のセンシティブメールを抽出しました。\n\n" +
        "「センシティブメール一覧」シートで内容を確認し、\n" +
        "削除したいメールにチェックを入れてから\n" +
        "「選択したメールを削除」を実行してください。", 
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
  } catch (error) {
    SpreadsheetApp.getUi().alert("エラー", error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * 選択したメールのみ削除
 */
function runDeleteSelectedEmails() {
  try {
    const deleted = deleteSelectedEmails();
    if (deleted > 0) {
      SpreadsheetApp.getUi().alert(
        "削除完了", 
        deleted + "件のメールをゴミ箱に移動しました。\n\n" +
        "削除履歴は「削除メール一覧」シートで確認できます。", 
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
  } catch (error) {
    SpreadsheetApp.getUi().alert("エラー", error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * データファイル一覧取得のみ実行（ドキュメント・スプレッドシート）
 */
function runListDataFiles() {
  try {
    const userConfig = getConfigFromSheet();
    const systemConfig = getConfig();
    
    if (!userConfig.userEmail) {
      SpreadsheetApp.getUi().alert("エラー", "退職者のメールアドレスを入力してください。", SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    listUserDataFiles(userConfig.userEmail, systemConfig);
    SpreadsheetApp.getUi().alert("完了", "データ一覧を取得しました。", SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (error) {
    SpreadsheetApp.getUi().alert("エラー", error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * カレンダー一覧取得のみ実行
 */
function runListUserCalendars() {
  try {
    const userConfig = getConfigFromSheet();
    const systemConfig = getConfig();
    
    if (!userConfig.userEmail) {
      SpreadsheetApp.getUi().alert("エラー", "退職者のメールアドレスを入力してください。", SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    listUserCalendars(userConfig.userEmail, systemConfig);
    SpreadsheetApp.getUi().alert("完了", "カレンダー一覧を取得しました。", SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (error) {
    SpreadsheetApp.getUi().alert("エラー", error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

// ===== コア機能 =====

/**
 * センシティブメールを検索してスプレッドシートに抽出
 */
function extractSensitiveEmails(userConfig, systemConfig) {
  const keywords = userConfig.keywords || [];
  if (keywords.length === 0) {
    throw new Error("削除対象のキーワードが設定されていません。");
  }
  
  const query = keywords.map(keyword => '(subject:"' + keyword + '" OR body:"' + keyword + '")').join(" OR ");
  const threads = GmailApp.search(query, 0, systemConfig.defaults.maxEmailsToProcess || 100);
  
  if (threads.length === 0) {
    return [];
  }
  
  const emailList = [];
  threads.forEach(thread => {
    const messages = thread.getMessages();
    messages.forEach(message => {
      // キーワードにマッチする理由を特定
      const subject = message.getSubject();
      const body = message.getPlainBody();
      const matchedKeywords = keywords.filter(keyword => 
        subject.includes(keyword) || body.includes(keyword)
      );
      
      emailList.push({
        threadId: thread.getId(),
        messageId: message.getId(),
        subject: subject,
        from: message.getFrom(),
        to: message.getTo(),
        date: message.getDate(),
        matchedKeywords: matchedKeywords.join(", "),
        snippet: message.getPlainBody().substring(0, 200).replace(/\n/g, " "),
        hasAttachments: message.getAttachments().length > 0
      });
    });
  });
  
  return emailList;
}

/**
 * センシティブメールの一覧をスプレッドシートに展開
 */
function createSensitiveEmailsSheet(userConfig, systemConfig) {
  const ui = SpreadsheetApp.getUi();
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // センシティブメール一覧シートを作成・更新
  let sheet = spreadsheet.getSheetByName("センシティブメール一覧");
  if (sheet) {
    sheet.clear();
  } else {
    sheet = spreadsheet.insertSheet("センシティブメール一覧");
  }
  
  // メールを抽出
  const emails = extractSensitiveEmails(userConfig, systemConfig);
  
  if (emails.length === 0) {
    ui.alert("センシティブなメールは見つかりませんでした。");
    return null;
  }
  
  // ヘッダーを設定
  const headers = [
    "削除対象",
    "件名",
    "送信者",
    "宛先",
    "受信日時",
    "マッチキーワード",
    "本文プレビュー",
    "添付ファイル",
    "スレッドID",
    "メッセージID"
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setBackground('#4285F4')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold');
  
  // データを設定
  const dataRows = emails.map(email => [
    true, // デフォルトで全て削除対象
    email.subject,
    email.from,
    email.to,
    email.date,
    email.matchedKeywords,
    email.snippet,
    email.hasAttachments ? "あり" : "なし",
    email.threadId,
    email.messageId
  ]);
  
  if (dataRows.length > 0) {
    sheet.getRange(2, 1, dataRows.length, headers.length).setValues(dataRows);
    
    // チェックボックスを設定
    sheet.getRange(2, 1, dataRows.length, 1).insertCheckboxes();
    
    // 列幅を調整
    sheet.setColumnWidth(1, 80);  // 削除対象
    sheet.setColumnWidth(2, 300); // 件名
    sheet.setColumnWidth(3, 200); // 送信者
    sheet.setColumnWidth(4, 200); // 宛先
    sheet.setColumnWidth(5, 150); // 受信日時
    sheet.setColumnWidth(6, 150); // マッチキーワード
    sheet.setColumnWidth(7, 400); // 本文プレビュー
    sheet.setColumnWidth(8, 100); // 添付ファイル
    sheet.setColumnWidth(9, 150); // スレッドID
    sheet.setColumnWidth(10, 150); // メッセージID
    
    // 条件付き書式を設定（削除対象がチェックされた行を薄い赤背景に）
    const range = sheet.getRange(2, 1, dataRows.length, headers.length);
    const rule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$A2=TRUE')
      .setBackground('#FFE4E1')
      .setRanges([range])
      .build();
    sheet.setConditionalFormatRules([rule]);
  }
  
  return sheet;
}

/**
 * スプレッドシートの選択に基づいてメールを削除
 */
function deleteSelectedEmails() {
  const ui = SpreadsheetApp.getUi();
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName("センシティブメール一覧");
  
  if (!sheet) {
    ui.alert("エラー", "センシティブメール一覧シートが見つかりません。\n先にメール抽出を実行してください。", ui.ButtonSet.OK);
    return 0;
  }
  
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    ui.alert("削除対象のメールがありません。");
    return 0;
  }
  
  // 削除対象のメールを収集
  const dataRange = sheet.getRange(2, 1, lastRow - 1, 10);
  const data = dataRange.getValues();
  const emailsToDelete = [];
  
  data.forEach((row, index) => {
    if (row[0] === true) { // 削除対象にチェックがある場合
      emailsToDelete.push({
        subject: row[1],
        from: row[2],
        date: row[4],
        threadId: row[8],
        messageId: row[9]
      });
    }
  });
  
  if (emailsToDelete.length === 0) {
    ui.alert("削除対象のメールが選択されていません。");
    return 0;
  }
  
  // 最終確認
  const confirmResponse = ui.alert(
    "メール削除の最終確認",
    emailsToDelete.length + "件のメールを削除します。\n\n" +
    "この操作はゴミ箱に移動します（30日後に完全削除）。\n" +
    "続行しますか？",
    ui.ButtonSet.YES_NO
  );
  
  if (confirmResponse !== ui.Button.YES) {
    return 0;
  }
  
  // 削除実行とログ記録
  const systemConfig = getConfig();
  const logSheet = spreadsheet.getSheetByName(systemConfig.sheets.deletedEmails) 
    || spreadsheet.insertSheet(systemConfig.sheets.deletedEmails);
  
  // ログシートのヘッダー設定
  if (logSheet.getLastRow() === 0) {
    logSheet.appendRow(["削除日時", "件名", "送信者", "受信日時", "スレッドID", "メッセージID", "復元状態"]);
    logSheet.getRange(1, 1, 1, 7)
      .setBackground('#FF6B6B')
      .setFontColor('#FFFFFF')
      .setFontWeight('bold');
  }
  
  const deletionTime = new Date();
  const processedThreads = new Set();
  let deletedCount = 0;
  
  emailsToDelete.forEach(email => {
    try {
      // スレッドIDで削除（同じスレッドは1回だけ処理）
      if (!processedThreads.has(email.threadId)) {
        const thread = GmailApp.getThreadById(email.threadId);
        if (thread) {
          thread.moveToTrash();
          processedThreads.add(email.threadId);
        }
      }
      
      // ログに記録
      logSheet.appendRow([
        deletionTime,
        email.subject,
        email.from,
        email.date,
        email.threadId,
        email.messageId,
        "未復元"
      ]);
      deletedCount++;
    } catch (error) {
      console.error("メール削除エラー:", error, email);
    }
  });
  
  // 列幅の自動調整
  logSheet.autoResizeColumns(1, 7);
  
  return deletedCount;
}

/**
 * センシティブなメールを削除（ゴミ箱へ移動）- 旧バージョン互換性のため保持
 */
function deleteSensitiveEmails(userConfig, systemConfig) {
  const ui = SpreadsheetApp.getUi();
  
  // 削除対象メールの検索
  const keywords = userConfig.keywords || [];
  if (keywords.length === 0) {
    ui.alert("削除対象のキーワードが設定されていません。");
    return 0;
  }
  
  // センシティブメール一覧シートを作成
  const sheet = createSensitiveEmailsSheet(userConfig, systemConfig);
  if (!sheet) {
    return 0;
  }
  
  // 削除前の確認
  const lastRow = sheet.getLastRow();
  const emailCount = lastRow - 1;
  
  const confirmResponse = ui.alert(
    "メール削除の確認",
    emailCount + "件のセンシティブなメールが見つかりました。\n" +
    "キーワード: " + keywords.join(", ") + "\n\n" +
    "スプレッドシートで詳細を確認し、削除対象を選択できます。\n\n" +
    "全て削除しますか？（「いいえ」を選択すると個別選択画面になります）",
    ui.ButtonSet.YES_NO_CANCEL
  );
  
  if (confirmResponse === ui.Button.CANCEL) {
    return 0;
  } else if (confirmResponse === ui.Button.NO) {
    // 個別選択を促す
    ui.alert(
      "個別選択モード",
      "「センシティブメール一覧」シートで削除したいメールにチェックを入れてから、\n" +
      "メニューの「退職者処理」→「個別機能」→「選択したメールを削除」を実行してください。",
      ui.ButtonSet.OK
    );
    return 0;
  }
  
  // 全て削除を選択した場合
  const deletedCount = deleteSelectedEmails();
  return deletedCount;
}

/**
 * ユーザーが所有するデータファイルの一覧を取得（ドキュメント・スプレッドシート）
 */
function listUserDataFiles(userEmail, systemConfig) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(systemConfig.sheets.dataFiles) 
      || SpreadsheetApp.getActiveSpreadsheet().insertSheet(systemConfig.sheets.dataFiles);
    sheet.clear();
    
    // ヘッダーを設定
    const headers = ["ファイル名", "タイプ", "最終更新日", "作成日", "URL", "説明", "親フォルダ"];
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length)
      .setBackground('#4285F4')
      .setFontColor('#FFFFFF')
      .setFontWeight('bold');
    
    const allFiles = [];
    const processedIds = new Set();
    
    // 1. スプレッドシートを検索
    try {
      console.log("スプレッドシートを検索中...");
      const spreadsheetQuery = 'mimeType="application/vnd.google-apps.spreadsheet" and "' + userEmail + '" in owners';
      const spreadsheets = DriveApp.searchFiles(spreadsheetQuery);
      
      while (spreadsheets.hasNext()) {
        const spreadsheet = spreadsheets.next();
        const fileId = spreadsheet.getId();
        if (!processedIds.has(fileId)) {
          processedIds.add(fileId);
          allFiles.push({
            name: spreadsheet.getName(),
            type: "スプレッドシート",
            lastUpdated: spreadsheet.getLastUpdated(),
            created: spreadsheet.getDateCreated(),
            url: spreadsheet.getUrl(),
            description: spreadsheet.getDescription() || "説明なし",
            parentFolder: getParentFolderName(spreadsheet)
          });
        }
      }
    } catch (error) {
      console.error("スプレッドシート検索エラー:", error);
    }
    
    // 2. Googleドキュメントを検索
    try {
      console.log("ドキュメントを検索中...");
      const docsQuery = 'mimeType="application/vnd.google-apps.document" and "' + userEmail + '" in owners';
      const docs = DriveApp.searchFiles(docsQuery);
      
      while (docs.hasNext()) {
        const doc = docs.next();
        const fileId = doc.getId();
        if (!processedIds.has(fileId)) {
          processedIds.add(fileId);
          allFiles.push({
            name: doc.getName(),
            type: "ドキュメント",
            lastUpdated: doc.getLastUpdated(),
            created: doc.getDateCreated(),
            url: doc.getUrl(),
            description: doc.getDescription() || "説明なし",
            parentFolder: getParentFolderName(doc)
          });
        }
      }
    } catch (error) {
      console.error("ドキュメント検索エラー:", error);
    }
    
    // 3. Googleフォームを検索
    try {
      console.log("フォームを検索中...");
      const formsQuery = 'mimeType="application/vnd.google-apps.form" and "' + userEmail + '" in owners';
      const forms = DriveApp.searchFiles(formsQuery);
      
      while (forms.hasNext()) {
        const form = forms.next();
        const fileId = form.getId();
        if (!processedIds.has(fileId)) {
          processedIds.add(fileId);
          allFiles.push({
            name: form.getName(),
            type: "フォーム",
            lastUpdated: form.getLastUpdated(),
            created: form.getDateCreated(),
            url: form.getUrl(),
            description: form.getDescription() || "説明なし",
            parentFolder: getParentFolderName(form)
          });
        }
      }
    } catch (error) {
      console.error("フォーム検索エラー:", error);
    }
    
    // 4. Googleスライドを検索
    try {
      console.log("スライドを検索中...");
      const slidesQuery = 'mimeType="application/vnd.google-apps.presentation" and "' + userEmail + '" in owners';
      const slides = DriveApp.searchFiles(slidesQuery);
      
      while (slides.hasNext()) {
        const slide = slides.next();
        const fileId = slide.getId();
        if (!processedIds.has(fileId)) {
          processedIds.add(fileId);
          allFiles.push({
            name: slide.getName(),
            type: "スライド",
            lastUpdated: slide.getLastUpdated(),
            created: slide.getDateCreated(),
            url: slide.getUrl(),
            description: slide.getDescription() || "説明なし",
            parentFolder: getParentFolderName(slide)
          });
        }
      }
    } catch (error) {
      console.error("スライド検索エラー:", error);
    }
    
    // 5. Google図形描画を検索
    try {
      console.log("図形描画を検索中...");
      const drawingQuery = 'mimeType="application/vnd.google-apps.drawing" and "' + userEmail + '" in owners';
      const drawings = DriveApp.searchFiles(drawingQuery);
      
      while (drawings.hasNext()) {
        const drawing = drawings.next();
        const fileId = drawing.getId();
        if (!processedIds.has(fileId)) {
          processedIds.add(fileId);
          allFiles.push({
            name: drawing.getName(),
            type: "図形描画",
            lastUpdated: drawing.getLastUpdated(),
            created: drawing.getDateCreated(),
            url: drawing.getUrl(),
            description: drawing.getDescription() || "説明なし",
            parentFolder: getParentFolderName(drawing)
          });
        }
      }
    } catch (error) {
      console.error("図形描画検索エラー:", error);
    }
    
    // 6. Google Sitesを検索
    try {
      console.log("Sitesを検索中...");
      const sitesQuery = 'mimeType="application/vnd.google-apps.site" and "' + userEmail + '" in owners';
      const sites = DriveApp.searchFiles(sitesQuery);
      
      while (sites.hasNext()) {
        const site = sites.next();
        const fileId = site.getId();
        if (!processedIds.has(fileId)) {
          processedIds.add(fileId);
          allFiles.push({
            name: site.getName(),
            type: "サイト",
            lastUpdated: site.getLastUpdated(),
            created: site.getDateCreated(),
            url: site.getUrl(),
            description: site.getDescription() || "説明なし",
            parentFolder: getParentFolderName(site)
          });
        }
      }
    } catch (error) {
      console.error("Sites検索エラー:", error);
    }
    
    // 7. Google Jamboardを検索
    try {
      console.log("Jamboardを検索中...");
      const jamboardQuery = 'mimeType="application/vnd.google-apps.jam" and "' + userEmail + '" in owners';
      const jamboards = DriveApp.searchFiles(jamboardQuery);
      
      while (jamboards.hasNext()) {
        const jamboard = jamboards.next();
        const fileId = jamboard.getId();
        if (!processedIds.has(fileId)) {
          processedIds.add(fileId);
          allFiles.push({
            name: jamboard.getName(),
            type: "Jamboard",
            lastUpdated: jamboard.getLastUpdated(),
            created: jamboard.getDateCreated(),
            url: jamboard.getUrl(),
            description: jamboard.getDescription() || "説明なし",
            parentFolder: getParentFolderName(jamboard)
          });
        }
      }
    } catch (error) {
      console.error("Jamboard検索エラー:", error);
    }
    
    // 8. Google My Mapsを検索
    try {
      console.log("My Mapsを検索中...");
      const mapsQuery = 'mimeType="application/vnd.google-apps.map" and "' + userEmail + '" in owners';
      const maps = DriveApp.searchFiles(mapsQuery);
      
      while (maps.hasNext()) {
        const map = maps.next();
        const fileId = map.getId();
        if (!processedIds.has(fileId)) {
          processedIds.add(fileId);
          allFiles.push({
            name: map.getName(),
            type: "マイマップ",
            lastUpdated: map.getLastUpdated(),
            created: map.getDateCreated(),
            url: map.getUrl(),
            description: map.getDescription() || "説明なし",
            parentFolder: getParentFolderName(map)
          });
        }
      }
    } catch (error) {
      console.error("My Maps検索エラー:", error);
    }
    
    // 9. Google Colaboratoryを検索
    try {
      console.log("Colaboratoryを検索中...");
      const colabQuery = 'mimeType="application/vnd.google.colaboratory" and "' + userEmail + '" in owners';
      const colabs = DriveApp.searchFiles(colabQuery);
      
      while (colabs.hasNext()) {
        const colab = colabs.next();
        const fileId = colab.getId();
        if (!processedIds.has(fileId)) {
          processedIds.add(fileId);
          allFiles.push({
            name: colab.getName(),
            type: "Colab",
            lastUpdated: colab.getLastUpdated(),
            created: colab.getDateCreated(),
            url: colab.getUrl(),
            description: colab.getDescription() || "説明なし",
            parentFolder: getParentFolderName(colab)
          });
        }
      }
    } catch (error) {
      console.error("Colaboratory検索エラー:", error);
    }
    
    // 10. その他のファイル（PDF、画像、動画など）を検索
    try {
      console.log("その他のファイルを検索中...");
      // PDFファイル
      const pdfQuery = 'mimeType="application/pdf" and "' + userEmail + '" in owners';
      const pdfs = DriveApp.searchFiles(pdfQuery);
      
      while (pdfs.hasNext()) {
        const pdf = pdfs.next();
        const fileId = pdf.getId();
        if (!processedIds.has(fileId)) {
          processedIds.add(fileId);
          allFiles.push({
            name: pdf.getName(),
            type: "PDF",
            lastUpdated: pdf.getLastUpdated(),
            created: pdf.getDateCreated(),
            url: pdf.getUrl(),
            description: pdf.getDescription() || "説明なし",
            parentFolder: getParentFolderName(pdf)
          });
        }
      }
      
      // 画像ファイル（主要な形式のみ）
      const imageTypes = [
        { mime: 'image/jpeg', type: 'JPEG画像' },
        { mime: 'image/png', type: 'PNG画像' },
        { mime: 'image/gif', type: 'GIF画像' }
      ];
      
      imageTypes.forEach(imageType => {
        const imageQuery = 'mimeType="' + imageType.mime + '" and "' + userEmail + '" in owners';
        const images = DriveApp.searchFiles(imageQuery);
        
        while (images.hasNext()) {
          const image = images.next();
          const fileId = image.getId();
          if (!processedIds.has(fileId)) {
            processedIds.add(fileId);
            allFiles.push({
              name: image.getName(),
              type: imageType.type,
              lastUpdated: image.getLastUpdated(),
              created: image.getDateCreated(),
              url: image.getUrl(),
              description: image.getDescription() || "説明なし",
              parentFolder: getParentFolderName(image)
            });
          }
        }
      });
      
      // 動画ファイル（主要な形式のみ）
      const videoTypes = [
        { mime: 'video/mp4', type: 'MP4動画' },
        { mime: 'video/quicktime', type: 'MOV動画' }
      ];
      
      videoTypes.forEach(videoType => {
        const videoQuery = 'mimeType="' + videoType.mime + '" and "' + userEmail + '" in owners';
        const videos = DriveApp.searchFiles(videoQuery);
        
        while (videos.hasNext()) {
          const video = videos.next();
          const fileId = video.getId();
          if (!processedIds.has(fileId)) {
            processedIds.add(fileId);
            allFiles.push({
              name: video.getName(),
              type: videoType.type,
              lastUpdated: video.getLastUpdated(),
              created: video.getDateCreated(),
              url: video.getUrl(),
              description: video.getDescription() || "説明なし",
              parentFolder: getParentFolderName(video)
            });
          }
        }
      });
      
    } catch (error) {
      console.error("その他のファイル検索エラー:", error);
    }
    
    // 結果をシートに記録
    if (allFiles.length > 0) {
      // 最終更新日でソート（新しい順）
      allFiles.sort((a, b) => b.lastUpdated - a.lastUpdated);
      
      allFiles.forEach(file => {
        sheet.appendRow([
          file.name,
          file.type,
          formatDate(file.lastUpdated),
          formatDate(file.created),
          file.url,
          file.description,
          file.parentFolder
        ]);
      });
      
      // サマリー情報を追加
      const summaryRow = sheet.getLastRow() + 2;
      sheet.getRange(summaryRow, 1).setValue("【サマリー】").setFontWeight("bold");
      sheet.getRange(summaryRow + 1, 1).setValue("総ファイル数:");
      sheet.getRange(summaryRow + 1, 2).setValue(allFiles.length);
      
      // タイプ別集計
      const typeCount = {};
      allFiles.forEach(file => {
        typeCount[file.type] = (typeCount[file.type] || 0) + 1;
      });
      
      let typeRow = summaryRow + 3;
      sheet.getRange(typeRow, 1).setValue("【タイプ別集計】").setFontWeight("bold");
      Object.entries(typeCount).forEach(([type, count], index) => {
        sheet.getRange(typeRow + 1 + index, 1).setValue(type + ":");
        sheet.getRange(typeRow + 1 + index, 2).setValue(count);
      });
      
    } else {
      sheet.appendRow(["データファイルは見つかりませんでした", "", "", "", "", "", ""]);
    }
    
    // 列幅の自動調整
    sheet.autoResizeColumns(1, headers.length);
    
    console.log("データファイル検索完了: " + allFiles.length + "件");
    
  } catch (error) {
    console.error("データファイル一覧取得エラー:", error);
    throw new Error("データファイル一覧の取得でエラーが発生しました: " + error.message);
  }
}

/**
 * 親フォルダ名を取得
 */
function getParentFolderName(file) {
  try {
    const parents = file.getParents();
    if (parents.hasNext()) {
      return parents.next().getName();
    }
    return "マイドライブ";
  } catch (error) {
    return "不明";
  }
}

// ===== 復元機能 =====

/**
 * 削除メール復元ダイアログを表示
 */
function showRestoreDialog() {
  const html = HtmlService.createHtmlOutputFromFile('restore-dialog')
    .setWidth(400)
    .setHeight(350);
  SpreadsheetApp.getUi().showModalDialog(html, 'メールの復元');
}

/**
 * HTMLダイアログから呼び出される復元関数
 */
function restoreEmailFromDialog(threadId) {
  try {
    const thread = GmailApp.getThreadById(threadId);
    if (!thread) {
      throw new Error('指定されたスレッドIDが見つかりません。');
    }
    
    if (!thread.isInTrash()) {
      return 'このメールは既に復元されています。';
    }
    
    thread.moveToInbox();
    
    // 削除メール一覧シートの復元状態を更新
    const config = getConfig();
    const deleteSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(config.sheets.deletedEmails);
    if (deleteSheet) {
      const data = deleteSheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (data[i][4] === threadId) {
          deleteSheet.getRange(i + 1, 6).setValue("復元済");
        }
      }
    }
    
    // 復元ログを記録
    const logSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(config.sheets.restoreLog) 
      || SpreadsheetApp.getActiveSpreadsheet().insertSheet(config.sheets.restoreLog);
    
    if (logSheet.getLastRow() === 0) {
      logSheet.appendRow(["復元日時", "スレッドID", "件名", "復元者"]);
    }
    
    const messages = thread.getMessages();
    if (messages.length > 0) {
      logSheet.appendRow([
        new Date(),
        threadId,
        messages[0].getSubject(),
        getActiveUserEmail()
      ]);
    }
    
    return 'メールを復元しました。受信トレイをご確認ください。';
  } catch (error) {
    console.error('復元エラー:', error);
    throw new Error(error.message);
  }
}

// ===== ユーティリティ関数 =====

/**
 * システム診断
 */
function runSystemDiagnostics() {
  const ui = SpreadsheetApp.getUi();
  const config = getConfig();
  let diagnosticsResult = "【システム診断結果】\n\n";
  diagnosticsResult += config.system.name + " v" + config.system.version + "\n\n";
  
  // APIサービスのチェック
  diagnosticsResult += "■ APIサービス状態\n";
  
  const apiChecks = [
    { name: "Gmail API", check: typeof Gmail !== 'undefined' },
    { name: "Admin Directory API", check: typeof AdminDirectory !== 'undefined' },
    { name: "Admin Reports API", check: typeof AdminReports !== 'undefined' }
  ];
  
  apiChecks.forEach(api => {
    diagnosticsResult += api.check ? "✓ " + api.name + ": 有効\n" : "✗ " + api.name + ": 無効（要追加）\n";
  });
  
  diagnosticsResult += "\n■ 権限チェック\n";
  
  // 権限チェック
  const permissionChecks = [
    {
      name: "メール読み取り権限",
      test: () => { GmailApp.getInboxThreads(0, 1); return true; }
    },
    {
      name: "ドライブアクセス権限",
      test: () => { DriveApp.getRootFolder(); return true; }
    },
    {
      name: "スプレッドシート権限",
      test: () => { SpreadsheetApp.getActiveSpreadsheet(); return true; }
    }
  ];
  
  permissionChecks.forEach(perm => {
    try {
      perm.test();
      diagnosticsResult += "✓ " + perm.name + ": OK\n";
    } catch (e) {
      diagnosticsResult += "✗ " + perm.name + ": NG\n";
    }
  });
  
  // 設定の確認
  diagnosticsResult += "\n■ 現在の設定\n";
  try {
    const userConfig = getConfigFromSheet();
    diagnosticsResult += "- 退職者メール: " + (userConfig.userEmail || '未設定') + "\n";
    diagnosticsResult += "- 転送先: " + userConfig.forwardEmail + "\n";
    diagnosticsResult += "- キーワード数: " + userConfig.keywords.length + "個\n";
  } catch (e) {
    diagnosticsResult += "- 設定シートが見つかりません\n";
  }
  
  // 推奨事項
  diagnosticsResult += "\n■ 推奨事項\n";
  if (!apiChecks.every(api => api.check)) {
    diagnosticsResult += "• Apps Scriptエディタで必要なサービスを追加してください\n";
  }
  diagnosticsResult += "• メールルーティングは管理コンソールでの設定を推奨\n";
  diagnosticsResult += "• 定期的に設定をエクスポートしてバックアップを取ることを推奨\n";
  
  ui.alert("システム診断", diagnosticsResult, ui.ButtonSet.OK);
}

/**
 * ヘルプの表示
 */
function showHelp() {
  const config = getConfig();
  const helpText = "【" + config.system.name + " v" + config.system.version + "】\n\n" +
    "■ 使い方\n\n" +
    "1. 初期設定\n" +
    "   メニュー「退職者処理」→「初期設定シートを作成」\n\n" +
    "2. 退職者情報の入力\n" +
    "   設定シートのB3セルに退職者のメールアドレスを入力\n\n" +
    "3. 処理の実行\n" +
    "   メニュー「退職者処理」→「退職処理を実行」\n\n" +
    "■ 新機能（v2.0.0）\n\n" +
    "◆ 外部SSOサービス検出\n" +
    "  - 退職者が利用していた外部サービスを詳細に検出\n" +
    "  - サービスタイプ別の分類\n" +
    "  - 最終利用日時の記録\n\n" +
    "◆ 総合レポート生成\n" +
    "  - 退職者の全体的な利用状況をレポート化\n" +
    "  - 推奨アクションの提示\n\n" +
    "■ メールルーティング設定\n\n" +
    "◆ 推奨: 管理コンソール（組織レベル）\n" +
    "  メニュー「メールルーティング」→「管理コンソール設定案内」\n" +
    "  - より確実な転送\n" +
    "  - ユーザーが変更不可\n" +
    "  - 管理者による一元管理\n\n" +
    "◆ 代替: 個人転送設定（ユーザーレベル）\n" +
    "  メニュー「メールルーティング」→「個人転送設定」\n" +
    "  - 簡単に設定可能\n" +
    "  - ユーザーが変更可能\n\n" +
    "◆ その他のオプション\n" +
    "  - メール委任設定\n" +
    "  - CSVエクスポート（一括処理用）\n\n" +
    "■ 設定管理\n\n" +
    "- キーワードや転送先の変更\n" +
    "  メニュー「設定管理」→「設定画面を開く」\n\n" +
    "- 設定のバックアップ\n" +
    "  メニュー「設定管理」→「設定をエクスポート」\n\n" +
    "■ メールの復元\n\n" +
    "誤って削除したメールは「削除メールの復元」から復元できます。\n" +
    "削除メール一覧シートでスレッドIDを確認してください。\n\n" +
    "■ 必要な権限\n\n" +
    "- Gmail API（メール操作）\n" +
    "- Admin Directory API（デバイス情報）\n" +
    "- Admin Reports API（SaaS連携情報）※管理者権限が必要\n" +
    "- Drive API（GASプロジェクト検索）\n\n" +
    "■ トラブルシューティング\n\n" +
    "APIエラーが発生する場合：\n" +
    "1. Apps Scriptエディタを開く\n" +
    "2. サービス → ＋ をクリック\n" +
    "3. 必要なAPIを追加\n\n" +
    "■ サポート\n\n" +
    "システム診断機能で状態を確認できます。\n" +
    "メニュー「退職者処理」→「システム診断」";
  
  SpreadsheetApp.getUi().alert("ヘルプ", helpText, SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * 処理開始ログ
 */
function logProcessStart(userEmail) {
  const config = getConfig();
  const logSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(config.sheets.processLog) 
    || SpreadsheetApp.getActiveSpreadsheet().insertSheet(config.sheets.processLog);
  
  if (logSheet.getLastRow() === 0) {
    logSheet.appendRow(["開始日時", "対象者", "実行者", "終了日時", "ステータス", "備考"]);
  }
  
  const rowIndex = logSheet.getLastRow() + 1;
  logSheet.getRange(rowIndex, 1).setValue(new Date());
  logSheet.getRange(rowIndex, 2).setValue(userEmail);
  logSheet.getRange(rowIndex, 3).setValue(getActiveUserEmail());
  
  // 処理中の行番号を保存
  PropertiesService.getScriptProperties().setProperty('currentProcessRow', rowIndex.toString());
}

/**
 * 処理終了ログ
 */
function logProcessEnd(userEmail, status) {
  const config = getConfig();
  const logSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(config.sheets.processLog);
  if (!logSheet) return;
  
  const rowIndex = parseInt(PropertiesService.getScriptProperties().getProperty('currentProcessRow') || '0');
  if (rowIndex > 0) {
    logSheet.getRange(rowIndex, 4).setValue(new Date());
    logSheet.getRange(rowIndex, 5).setValue(status);
    logSheet.getRange(rowIndex, 6).setValue("メールルーティングは別途設定が必要");
  }
  
  // プロパティをクリア
  PropertiesService.getScriptProperties().deleteProperty('currentProcessRow');
}

/**
 * 日付をフォーマット
 */
function formatDate(dateValue) {
  if (!dateValue) return 'N/A';
  
  try {
    const date = new Date(dateValue);
    if (isNaN(date.getTime())) return 'N/A';
    
    return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm');
  } catch (error) {
    return 'N/A';
  }
}

/**
 * アクティブユーザーのメールアドレスを取得（エラーハンドリング付き）
 */
function getActiveUserEmail() {
  try {
    return Session.getActiveUser().getEmail();
  } catch (error) {
    console.log('ユーザー情報取得エラー: userinfo.email スコープが必要です');
    return '不明';
  }
}

/**
 * 外部サービス利用状況分析
 */
function analyzeExternalServices() {
  try {
    const userConfig = getConfigFromSheet();
    if (!userConfig.userEmail) {
      SpreadsheetApp.getUi().alert("エラー", "退職者のメールアドレスを入力してください。", SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    const systemConfig = getConfig();
    const services = getDetailedExternalSSOServices(userConfig.userEmail, systemConfig.defaults.ssoLookbackDays || 365);
    
    let analysisResult = "=== 外部サービス利用状況分析 ===\n\n";
    analysisResult += "対象ユーザー: " + userConfig.userEmail + "\n";
    analysisResult += "分析期間: 過去" + (systemConfig.defaults.ssoLookbackDays || 365) + "日間\n\n";
    
    analysisResult += "■ サマリー\n";
    analysisResult += "総サービス数: " + services.length + "\n\n";
    
    // タイプ別集計
    const typeCount = {};
    services.forEach(service => {
      typeCount[service.type] = (typeCount[service.type] || 0) + 1;
    });
    
    analysisResult += "■ カテゴリ別利用状況\n";
    Object.entries(typeCount)
      .sort((a, b) => b[1] - a[1])
      .forEach(([type, count]) => {
        analysisResult += type + ": " + count + "サービス\n";
      });
    
    analysisResult += "\n■ 主要サービスの利用状況\n";
    const popularServices = systemConfig.defaults.popularSSOServices || [];
    popularServices.forEach(popularService => {
      const found = services.find(s => 
        s.name.toLowerCase().includes(popularService.toLowerCase())
      );
      if (found) {
        analysisResult += "✓ " + popularService + " - 最終利用: " + formatDate(found.lastUsed) + "\n";
      } else {
        analysisResult += "✗ " + popularService + " - 利用なし\n";
      }
    });
    
    analysisResult += "\n■ 最近利用したサービス（上位10件）\n";
    services.slice(0, 10).forEach((service, index) => {
      analysisResult += (index + 1) + ". " + service.name + " (" + service.type + ") - " + formatDate(service.lastUsed) + "\n";
    });
    
    SpreadsheetApp.getUi().alert("外部サービス利用状況分析", analysisResult, SpreadsheetApp.getUi().ButtonSet.OK);
    
  } catch (error) {
    SpreadsheetApp.getUi().alert("エラー", error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * デバイス利用状況分析
 */
function analyzeDeviceUsage() {
  const ui = SpreadsheetApp.getUi();
  ui.alert("デバイス利用状況分析", "この機能は開発中です。\n\n利用デバイス一覧は「個別機能」→「利用デバイス一覧取得」から確認できます。", ui.ButtonSet.OK);
}

// ===== デバッグ・テスト用関数 =====

/**
 * 設定をリセット（デバッグ用）
 */
function resetConfiguration() {
  PropertiesService.getScriptProperties().deleteAllProperties();
  CONFIG_CACHE = null;
  initializeConfiguration();
  console.log("設定をリセットしました");
}

/**
 * 設定のテスト（デバッグ用）
 */
function testConfiguration() {
  try {
    const config = getConfig();
    console.log("設定:", JSON.stringify(config, null, 2));
    console.log("キーワード数:", (config.defaults.sensitiveKeywords || []).length);
    console.log("人気サービス数:", (config.defaults.popularSSOServices || []).length);
  } catch (error) {
    console.error("エラー:", error);
  }
}

/**
 * デバイス情報のデバッグ（詳細版）
 */
function debugDeviceInfo() {
  try {
    const userConfig = getConfigFromSheet();
    if (!userConfig.userEmail) {
      console.log("エラー: 退職者のメールアドレスが設定されていません");
      return;
    }
    
    const userEmail = userConfig.userEmail;
    console.log("=== デバイス情報デバッグ ===");
    console.log("対象ユーザー:", userEmail);
    console.log("実行時刻:", new Date().toISOString());
    
    // API利用可能性チェック
    console.log("\n【APIチェック】");
    console.log("AdminDirectory:", typeof AdminDirectory !== 'undefined' ? "✓ 利用可能" : "✗ 利用不可");
    console.log("AdminReports:", typeof AdminReports !== 'undefined' ? "✓ 利用可能" : "✗ 利用不可");
    
    if (typeof AdminDirectory === 'undefined') {
      console.log("\n⚠️ Admin Directory APIが追加されていません。");
      console.log("Apps Scriptエディタで「サービス」→「＋」→「Admin SDK API」を追加してください。");
      return;
    }
    
    // 1. ユーザー情報の確認
    console.log("\n【ユーザー情報確認】");
    try {
      const user = AdminDirectory.Users.get(userEmail);
      console.log("ユーザー名:", user.name.fullName);
      console.log("プライマリメール:", user.primaryEmail);
      console.log("組織単位:", user.orgUnitPath);
      console.log("ステータス:", user.suspended ? "停止中" : "アクティブ");
    } catch (e) {
      console.log("ユーザー情報取得エラー:", e.message);
    }
    
    // 2. Chrome OSデバイス（全体）
    console.log("\n【Chrome OSデバイス（組織全体）】");
    try {
      const allChromeDevices = AdminDirectory.Chromeosdevices.list('my_customer', {
        maxResults: 5,
        projection: 'FULL'
      });
      
      if (allChromeDevices && allChromeDevices.chromeosdevices) {
        console.log("組織内の総Chrome OSデバイス数:", allChromeDevices.chromeosdevices.length);
        
        // ユーザーに関連するデバイスを探す
        let userDeviceCount = 0;
        allChromeDevices.chromeosdevices.forEach((device, index) => {
          console.log(`\nデバイス ${index + 1}:`);
          console.log("- annotatedUser:", device.annotatedUser || "未設定");
          console.log("- モデル:", device.model || "不明");
          console.log("- シリアル番号:", device.serialNumber || "N/A");
          
          if (device.annotatedUser === userEmail) {
            userDeviceCount++;
            console.log("→ このデバイスは対象ユーザーのものです！");
          }
          
          if (device.recentUsers) {
            console.log("- 最近のユーザー:");
            device.recentUsers.forEach(user => {
              console.log("  - " + user.email + " (タイプ: " + user.type + ")");
              if (user.email === userEmail) {
                userDeviceCount++;
                console.log("  → 対象ユーザーが使用しています！");
              }
            });
          }
        });
        
        console.log("\n対象ユーザーのChrome OSデバイス数:", userDeviceCount);
      } else {
        console.log("Chrome OSデバイスが見つかりません");
      }
    } catch (e) {
      console.log("Chrome OSデバイス取得エラー:", e.message);
    }
    
    // 3. モバイルデバイス（全体）
    console.log("\n【モバイルデバイス（組織全体）】");
    try {
      const allMobileDevices = AdminDirectory.Mobiledevices.list('my_customer', {
        maxResults: 5,
        projection: 'FULL'
      });
      
      if (allMobileDevices && allMobileDevices.mobiledevices) {
        console.log("組織内の総モバイルデバイス数:", allMobileDevices.mobiledevices.length);
        
        let userMobileCount = 0;
        allMobileDevices.mobiledevices.forEach((device, index) => {
          console.log(`\nモバイルデバイス ${index + 1}:`);
          console.log("- email:", device.email ? device.email.join(", ") : "未設定");
          console.log("- モデル:", device.model || "不明");
          console.log("- OS:", device.os || "不明");
          
          if (device.email && device.email.includes(userEmail)) {
            userMobileCount++;
            console.log("→ このデバイスは対象ユーザーのものです！");
          }
        });
        
        console.log("\n対象ユーザーのモバイルデバイス数:", userMobileCount);
      } else {
        console.log("モバイルデバイスが見つかりません");
      }
    } catch (e) {
      console.log("モバイルデバイス取得エラー:", e.message);
    }
    
    // 4. 異なるクエリ方法をテスト
    console.log("\n【クエリテスト】");
    
    // テスト1: query パラメータを使わない
    console.log("\nテスト1: 全デバイスを取得してフィルタリング");
    try {
      const devices = AdminDirectory.Mobiledevices.list('my_customer', {
        maxResults: 100
      });
      
      if (devices && devices.mobiledevices) {
        const userDevices = devices.mobiledevices.filter(d => 
          d.email && d.email.includes(userEmail)
        );
        console.log("フィルタリング結果:", userDevices.length + "台");
      }
    } catch (e) {
      console.log("エラー:", e.message);
    }
    
    // テスト2: 異なるクエリ形式
    console.log("\nテスト2: 異なるクエリ形式");
    const queryFormats = [
      `email:${userEmail}`,
      `email=${userEmail}`,
      `user:${userEmail}`,
      userEmail
    ];
    
    queryFormats.forEach(query => {
      try {
        console.log(`- クエリ "${query}" をテスト中...`);
        const result = AdminDirectory.Mobiledevices.list('my_customer', {
          query: query,
          maxResults: 10
        });
        
        if (result && result.mobiledevices) {
          console.log(`  結果: ${result.mobiledevices.length}台`);
        } else {
          console.log("  結果: 0台");
        }
      } catch (e) {
        console.log(`  エラー: ${e.message}`);
      }
    });
    
    // 5. エンドポイントデバイス
    console.log("\n【エンドポイントデバイス】");
    try {
      const endpoints = AdminDirectory.Devices.list({
        customer: 'my_customer',
        maxResults: 10
      });
      
      if (endpoints && endpoints.devices) {
        console.log("エンドポイントデバイス数:", endpoints.devices.length);
        endpoints.devices.forEach((device, index) => {
          if (index < 3) { // 最初の3台のみ表示
            console.log(`\nデバイス ${index + 1}:`);
            console.log("- annotatedUser:", device.annotatedUser || "未設定");
            console.log("- OS:", device.os || "不明");
            console.log("- モデル:", device.model || "不明");
          }
        });
      } else {
        console.log("エンドポイントデバイスが見つかりません");
      }
    } catch (e) {
      console.log("エンドポイントデバイス取得エラー:", e.message);
      console.log("（エンドポイント検証が有効になっていない可能性があります）");
    }
    
    // 6. 権限の確認
    console.log("\n【権限確認】");
    try {
      const me = AdminDirectory.Users.get('me');
      console.log("実行ユーザー:", me.primaryEmail);
      console.log("管理者権限:", me.isAdmin ? "あり" : "なし");
    } catch (e) {
      console.log("権限確認エラー:", e.message);
    }
    
  } catch (error) {
    console.error("デバッグ実行エラー:", error);
  }
}

/**
 * 特定ユーザーのデバイスを直接検索
 */
function searchUserDevicesDirectly() {
  const userConfig = getConfigFromSheet();
  if (!userConfig.userEmail) {
    console.log("エラー: 退職者のメールアドレスが設定されていません");
    return;
  }
  
  const userEmail = userConfig.userEmail;
  console.log("=== 直接検索 ===");
  console.log("対象:", userEmail);
  
  // モバイルデバイスの全取得とフィルタリング
  console.log("\n【モバイルデバイス検索】");
  try {
    let pageToken = null;
    let totalDevices = 0;
    let userDevices = [];
    
    do {
      const response = AdminDirectory.Mobiledevices.list('my_customer', {
        pageToken: pageToken,
        maxResults: 100,
        projection: 'FULL'
      });
      
      if (response.mobiledevices) {
        totalDevices += response.mobiledevices.length;
        
        response.mobiledevices.forEach(device => {
          // emailフィールドの確認
          if (device.email) {
            if (Array.isArray(device.email)) {
              if (device.email.includes(userEmail)) {
                userDevices.push(device);
              }
            } else if (device.email === userEmail) {
              userDevices.push(device);
            }
          }
          
          // ownerフィールドの確認
          if (device.owner && device.owner.includes(userEmail)) {
            if (!userDevices.includes(device)) {
              userDevices.push(device);
            }
          }
        });
      }
      
      pageToken = response.nextPageToken;
    } while (pageToken);
    
    console.log("総デバイス数:", totalDevices);
    console.log("対象ユーザーのデバイス数:", userDevices.length);
    
    if (userDevices.length > 0) {
      console.log("\n見つかったデバイス:");
      userDevices.forEach((device, index) => {
        console.log(`\nデバイス ${index + 1}:`);
        console.log("- タイプ:", device.type || "不明");
        console.log("- モデル:", device.model || "不明");
        console.log("- email:", device.email);
        console.log("- owner:", device.owner);
        console.log("- status:", device.status);
      });
    }
    
  } catch (e) {
    console.log("エラー:", e.message);
  }
  
  // Chrome OSデバイスの全取得とフィルタリング
  console.log("\n【Chrome OSデバイス検索】");
  try {
    let pageToken = null;
    let totalDevices = 0;
    let userDevices = [];
    
    do {
      const response = AdminDirectory.Chromeosdevices.list('my_customer', {
        pageToken: pageToken,
        maxResults: 100,
        projection: 'FULL'
      });
      
      if (response.chromeosdevices) {
        totalDevices += response.chromeosdevices.length;
        
        response.chromeosdevices.forEach(device => {
          // annotatedUserフィールドの確認
          if (device.annotatedUser === userEmail) {
            userDevices.push(device);
          }
          
          // recentUsersフィールドの確認
          if (device.recentUsers) {
            const hasUser = device.recentUsers.some(user => user.email === userEmail);
            if (hasUser && !userDevices.includes(device)) {
              userDevices.push(device);
            }
          }
        });
      }
      
      pageToken = response.nextPageToken;
    } while (pageToken);
    
    console.log("総Chrome OSデバイス数:", totalDevices);
    console.log("対象ユーザーのデバイス数:", userDevices.length);
    
    if (userDevices.length > 0) {
      console.log("\n見つかったデバイス:");
      userDevices.forEach((device, index) => {
        console.log(`\nデバイス ${index + 1}:`);
        console.log("- モデル:", device.model || "不明");
        console.log("- annotatedUser:", device.annotatedUser);
        console.log("- シリアル番号:", device.serialNumber || "N/A");
        console.log("- 最終同期:", device.lastSync);
      });
    }
    
  } catch (e) {
    console.log("エラー:", e.message);
  }
}

/**
 * Admin SDKの設定確認
 */
function checkAdminSDKSetup() {
  console.log("=== Admin SDK設定確認 ===");
  
  // 1. サービスの確認
  console.log("\n【追加されているサービス】");
  console.log("AdminDirectory:", typeof AdminDirectory !== 'undefined' ? "✓" : "✗");
  console.log("AdminReports:", typeof AdminReports !== 'undefined' ? "✓" : "✗");
  console.log("Gmail:", typeof Gmail !== 'undefined' ? "✓" : "✗");
  
  // 2. スコープの確認
  console.log("\n【必要なスコープ】");
  console.log("以下のスコープが必要です:");
  console.log("- https://www.googleapis.com/auth/admin.directory.device.chromeos.readonly");
  console.log("- https://www.googleapis.com/auth/admin.directory.device.mobile.readonly");
  console.log("- https://www.googleapis.com/auth/admin.directory.user.readonly");
  
  // 3. 実行ユーザーの確認
  console.log("\n【実行ユーザー】");
  try {
    console.log("実行者:", Session.getActiveUser().getEmail());
  } catch (e) {
    console.log("実行者: （権限不足により取得できません）");
    console.log("※ userinfo.email スコープが必要です");
  }
  console.log("※ 管理者権限が必要です");
  
  // 4. 組織の設定確認
  if (typeof AdminDirectory !== 'undefined') {
    try {
      console.log("\n【組織情報】");
      const customer = AdminDirectory.Customers.get('my_customer');
      console.log("組織ID:", customer.id);
      console.log("ドメイン:", customer.customerDomain);
    } catch (e) {
      console.log("組織情報取得エラー:", e.message);
    }
  }
}

/**
 * 必要な権限（スコープ）の設定ガイド
 */
function showPermissionSetupGuide() {
  const ui = SpreadsheetApp.getUi();
  
  const guide = `【デバイス情報取得の権限設定】

現在、デバイス情報を取得するための権限が不足しています。
以下の手順で権限を追加してください：

■ 方法1: マニフェストファイルで設定（推奨）
1. Apps Scriptエディタで「プロジェクト設定」（歯車アイコン）をクリック
2. 「「appsscript.json」マニフェスト ファイルをエディタで表示する」にチェック
3. エディタに戻り、「appsscript.json」ファイルを開く
4. 以下のコードをコピーして貼り付け：

{
  "timeZone": "Asia/Tokyo",
  "dependencies": {
    "enabledAdvancedServices": [
      {
        "userSymbol": "AdminDirectory",
        "version": "directory_v1",
        "serviceId": "admin"
      },
      {
        "userSymbol": "AdminReports",
        "version": "reports_v1",
        "serviceId": "admin"
      },
      {
        "userSymbol": "Gmail",
        "version": "v1",
        "serviceId": "gmail"
      }
    ]
  },
  "exceptionLogging": "STACKDRIVER",
  "runtimeVersion": "V8",
  "oauthScopes": [
    "https://www.googleapis.com/auth/spreadsheets.currentonly",
    "https://www.googleapis.com/auth/gmail.modify",
    "https://www.googleapis.com/auth/gmail.settings.basic",
    "https://www.googleapis.com/auth/gmail.settings.sharing",
    "https://www.googleapis.com/auth/admin.directory.user.readonly",
    "https://www.googleapis.com/auth/admin.directory.device.chromeos.readonly",
    "https://www.googleapis.com/auth/admin.directory.device.mobile.readonly",
    "https://www.googleapis.com/auth/admin.reports.audit.readonly",
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/userinfo.email"
  ]
}

5. ファイルを保存（Ctrl+S または Cmd+S）
6. スクリプトを実行し、新しい認証を行う

■ 方法2: 手動でサービスを追加
1. Apps Scriptエディタで「サービス」（+アイコン）をクリック
2. 「Admin SDK API」を探して選択
3. バージョンは最新を選択
4. 識別子は「AdminDirectory」のまま
5. 「追加」をクリック

■ 重要な注意事項
- Google Workspace管理者権限が必要です
- 初回実行時に新しい認証画面が表示されます
- すべての権限を承認してください

権限追加後、もう一度デバイス取得を実行してください。`;
  
  ui.alert('権限設定ガイド', guide, ui.ButtonSet.OK);
  
  // コンソールにも出力
  console.log(guide);
  
  // appsscript.jsonの内容をコンソールに出力
  console.log('\n【appsscript.jsonの完全な内容】');
  console.log(JSON.stringify({
    "timeZone": "Asia/Tokyo",
    "dependencies": {
      "enabledAdvancedServices": [
        {
          "userSymbol": "AdminDirectory",
          "version": "directory_v1",
          "serviceId": "admin"
        },
        {
          "userSymbol": "AdminReports",
          "version": "reports_v1",
          "serviceId": "admin"
        },
        {
          "userSymbol": "Gmail",
          "version": "v1",
          "serviceId": "gmail"
        }
      ]
    },
    "exceptionLogging": "STACKDRIVER",
    "runtimeVersion": "V8",
    "oauthScopes": [
      "https://www.googleapis.com/auth/spreadsheets.currentonly",
      "https://www.googleapis.com/auth/gmail.modify",
      "https://www.googleapis.com/auth/gmail.settings.basic",
      "https://www.googleapis.com/auth/gmail.settings.sharing",
      "https://www.googleapis.com/auth/admin.directory.user.readonly",
      "https://www.googleapis.com/auth/admin.directory.device.chromeos.readonly",
      "https://www.googleapis.com/auth/admin.directory.device.mobile.readonly",
      "https://www.googleapis.com/auth/admin.reports.audit.readonly",
      "https://www.googleapis.com/auth/drive",
      "https://www.googleapis.com/auth/userinfo.email"
    ]
  }, null, 2));
}

/**
 * 現在の権限状態を確認
 */
function checkCurrentPermissions() {
  console.log("=== 現在の権限状態 ===");
  
  // 現在のスコープを取得する方法はないため、APIを実際に呼び出してテスト
  const tests = [
    {
      name: "スプレッドシートアクセス",
      test: () => { SpreadsheetApp.getActiveSpreadsheet(); return true; }
    },
    {
      name: "Gmail読み取り",
      test: () => { GmailApp.getInboxThreads(0, 1); return true; }
    },
    {
      name: "ドライブアクセス",
      test: () => { DriveApp.getRootFolder(); return true; }
    },
    {
      name: "ユーザー情報読み取り",
      test: () => { 
        if (typeof AdminDirectory === 'undefined') return false;
        AdminDirectory.Users.get('me'); 
        return true; 
      }
    },
    {
      name: "モバイルデバイス読み取り",
      test: () => { 
        if (typeof AdminDirectory === 'undefined') return false;
        AdminDirectory.Mobiledevices.list('my_customer', {maxResults: 1}); 
        return true; 
      }
    },
    {
      name: "Chrome OSデバイス読み取り",
      test: () => { 
        if (typeof AdminDirectory === 'undefined') return false;
        AdminDirectory.Chromeosdevices.list('my_customer', {maxResults: 1}); 
        return true; 
      }
    },
    {
      name: "レポートAPI",
      test: () => { 
        if (typeof AdminReports === 'undefined') return false;
        const endTime = new Date();
        const startTime = new Date(endTime.getTime() - 24 * 60 * 60 * 1000);
        AdminReports.Activities.list('all', 'login', {
          startTime: startTime.toISOString(),
          endTime: endTime.toISOString(),
          maxResults: 1
        }); 
        return true; 
      }
    },
    {
      name: "Gmail設定変更",
      test: () => { 
        if (typeof Gmail === 'undefined') return false;
        Gmail.Users.Settings.getAutoForwarding('me'); 
        return true; 
      }
    }
  ];
  
  tests.forEach(test => {
    try {
      const result = test.test();
      console.log(`${result ? '✓' : '✗'} ${test.name}`);
    } catch (e) {
      console.log(`✗ ${test.name}: ${e.message}`);
    }
  });
  
  console.log("\n必要な権限が不足している場合は showPermissionSetupGuide() を実行してください");
}