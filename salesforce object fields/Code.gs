// Salesforce REST API Explorer for Google Sheets
// Salesforceの接続設定を管理

const CONFIG = {
  CLIENT_ID: PropertiesService.getScriptProperties().getProperty('SF_CLIENT_ID'),
  CLIENT_SECRET: PropertiesService.getScriptProperties().getProperty('SF_CLIENT_SECRET'),
  USERNAME: PropertiesService.getScriptProperties().getProperty('SF_USERNAME'),
  PASSWORD: PropertiesService.getScriptProperties().getProperty('SF_PASSWORD'),
  SECURITY_TOKEN: PropertiesService.getScriptProperties().getProperty('SF_SECURITY_TOKEN'),
  LOGIN_URL: PropertiesService.getScriptProperties().getProperty('SF_LOGIN_URL') || 'https://login.salesforce.com'
};

// メニュー作成
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Salesforce Explorer')
    .addItem('設定を初期化', 'showConfigDialog')
    .addItem('オブジェクト一覧を取得', 'fetchObjectList')
    .addItem('選択オブジェクトのフィールドを取得', 'fetchFieldsForSelectedObject')
    .addItem('リレーション分析を実行', 'analyzeRelationships')
    .addItem('データ構造サマリーを作成', 'createDataStructureSummary')
    .addSeparator()
    .addItem('全データを自動取得', 'runFullAnalysis')
    .addItem('自動実行をスケジュール設定', 'setupScheduledExecution')
    .addSeparator()
    .addItem('全データをクリア', 'clearAllData')
    .addToUi();
}

// 設定ダイアログを表示
function showConfigDialog() {
  const html = HtmlService.createHtmlOutputFromFile('ConfigDialog')
    .setWidth(500)
    .setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(html, 'Salesforce接続設定');
}

// 設定を保存
function saveConfig(config) {
  const scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperty('SF_CLIENT_ID', config.clientId);
  scriptProperties.setProperty('SF_CLIENT_SECRET', config.clientSecret);
  scriptProperties.setProperty('SF_USERNAME', config.username);
  scriptProperties.setProperty('SF_PASSWORD', config.password);
  scriptProperties.setProperty('SF_SECURITY_TOKEN', config.securityToken);
  scriptProperties.setProperty('SF_LOGIN_URL', config.loginUrl || 'https://login.salesforce.com');

  return '設定が保存されました';
}

// ログをシートに記録
function logToSheet(message, type = 'INFO') {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ExecutionLog') ||
                SpreadsheetApp.getActiveSpreadsheet().insertSheet('ExecutionLog');

  // 初回の場合はヘッダーを作成
  if (sheet.getLastRow() === 0) {
    const headers = ['タイムスタンプ', 'タイプ', 'メッセージ'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setBackground('#4285F4').setFontColor('#FFFFFF').setFontWeight('bold');
  }

  // ログを追加
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  sheet.appendRow([timestamp, type, message]);

  // 古いログを削除（1000行を超えたら古いものから削除）
  if (sheet.getLastRow() > 1000) {
    sheet.deleteRows(2, 100);
  }
}

// 全データを自動取得（メイン実行関数）
function runFullAnalysis() {
  const startTime = new Date();
  logToSheet('=== 自動実行開始 ===', 'START');

  const errors = [];
  let objectCount = 0;
  let fieldCount = 0;
  let relationCount = 0;
  let structureCount = 0;

  // 1. オブジェクト一覧を取得
  try {
    logToSheet('オブジェクト一覧取得を開始', 'INFO');
    objectCount = fetchObjectListSilent();
    logToSheet(`${objectCount}個のオブジェクトを取得しました`, 'SUCCESS');
  } catch (e) {
    const errorMsg = `オブジェクト一覧取得エラー: ${e.toString()}`;
    logToSheet(errorMsg, 'ERROR');
    errors.push(errorMsg);
  }

  // 2. 主要オブジェクトのフィールドを取得
  try {
    logToSheet('主要オブジェクトのフィールド取得を開始', 'INFO');
    const mainObjects = ['Account', 'Contact', 'Opportunity', 'Lead', 'Case', 'User'];
    let fieldErrors = 0;

    mainObjects.forEach(objName => {
      try {
        const count = fetchFieldsForObjectSilent(objName);
        fieldCount += count;
        logToSheet(`${objName}: ${count}個のフィールドを取得`, 'SUCCESS');
      } catch (e) {
        const errorMsg = `${objName}のフィールド取得エラー: ${e.toString()}`;
        logToSheet(errorMsg, 'ERROR');
        fieldErrors++;
      }
    });

    if (fieldErrors > 0) {
      errors.push(`${fieldErrors}個のオブジェクトでフィールド取得失敗`);
    }
  } catch (e) {
    const errorMsg = `フィールド取得処理全体エラー: ${e.toString()}`;
    logToSheet(errorMsg, 'ERROR');
    errors.push(errorMsg);
  }

  // 3. リレーション分析
  try {
    logToSheet('リレーション分析を開始', 'INFO');
    relationCount = analyzeRelationshipsSilent();
    logToSheet(`${relationCount}個のリレーションを分析しました`, 'SUCCESS');
  } catch (e) {
    const errorMsg = `リレーション分析エラー: ${e.toString()}`;
    logToSheet(errorMsg, 'ERROR');
    errors.push(errorMsg);
  }

  // 4. データ構造サマリー作成
  try {
    logToSheet('データ構造サマリー作成を開始', 'INFO');
    structureCount = createDataStructureSummarySilent();
    logToSheet(`${structureCount}個のオブジェクト構造を分析しました`, 'SUCCESS');
  } catch (e) {
    const errorMsg = `データ構造分析エラー: ${e.toString()}`;
    logToSheet(errorMsg, 'ERROR');
    errors.push(errorMsg);
  }

  // 実行時間を記防
  const endTime = new Date();
  const duration = Math.round((endTime - startTime) / 1000);

  // 結果サマリー
  if (errors.length > 0) {
    logToSheet(`=== 自動実行完了（一部エラーあり） ===`, 'WARNING');
    logToSheet(`エラー数: ${errors.length}`, 'WARNING');
    errors.forEach(err => logToSheet(`  - ${err}`, 'WARNING'));
  } else {
    logToSheet(`=== 自動実行完了 (実行時間: ${duration}秒) ===`, 'COMPLETE');
  }

  // サマリー情報を更新
  try {
    updateSummarySheet(objectCount, fieldCount, relationCount, structureCount, duration, errors);
  } catch (e) {
    logToSheet(`サマリーシート更新エラー: ${e.toString()}`, 'ERROR');
  }

  // 部分的に成功した場合もtrueを返す
  return objectCount > 0 || fieldCount > 0 || relationCount > 0 || structureCount > 0;
}

// サマリーシートを更新
function updateSummarySheet(objectCount, fieldCount, relationCount, structureCount, duration, errors = []) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Summary') ||
                SpreadsheetApp.getActiveSpreadsheet().insertSheet('Summary', 0);

  sheet.clear();

  // タイトル
  sheet.getRange(1, 1).setValue('Salesforce データ分析サマリー');
  sheet.getRange(1, 1).setFontSize(16).setFontWeight('bold');

  // 最終更新日時
  const lastUpdate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  sheet.getRange(2, 1).setValue('最終更新: ' + lastUpdate);

  // ステータス
  const statusText = errors.length > 0 ?
    `ステータス: 部分的成功（エラー${errors.length}件）` :
    'ステータス: 完全成功';
  sheet.getRange(3, 1).setValue(statusText);
  if (errors.length > 0) {
    sheet.getRange(3, 1).setFontColor('#FF6600');
  } else {
    sheet.getRange(3, 1).setFontColor('#008000');
  }

  // 統計情報
  const stats = [
    ['項目', '件数'],
    ['取得オブジェクト数', objectCount || 0],
    ['取得フィールド数', fieldCount || 0],
    ['分析リレーション数', relationCount || 0],
    ['構造分析オブジェクト数', structureCount || 0],
    ['実行時間（秒）', duration || 0],
    ['エラー数', errors.length]
  ];

  sheet.getRange(5, 1, stats.length, 2).setValues(stats);
  sheet.getRange(5, 1, 1, 2).setBackground('#4285F4').setFontColor('#FFFFFF').setFontWeight('bold');

  // エラー詳細を表示
  if (errors.length > 0) {
    const errorRow = 5 + stats.length + 2;
    sheet.getRange(errorRow, 1).setValue('エラー詳細:');
    sheet.getRange(errorRow, 1).setFontWeight('bold');

    errors.forEach((error, index) => {
      sheet.getRange(errorRow + 1 + index, 1).setValue(`${index + 1}. ${error}`);
      sheet.getRange(errorRow + 1 + index, 1).setFontColor('#CC0000');
    });
  }

  // 列幅調整
  sheet.autoResizeColumns(1, 2);
}

// スケジュール実行の設定
function setupScheduledExecution() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    'スケジュール設定',
    '実行間隔を選択してください:\n' +
    '1: 毎日\n' +
    '2: 毎週\n' +
    '3: 毎月\n' +
    '0: スケジュールを削除',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() === ui.Button.OK) {
    // 既存のトリガーを削除
    ScriptApp.getProjectTriggers().forEach(trigger => {
      if (trigger.getHandlerFunction() === 'runFullAnalysis') {
        ScriptApp.deleteTrigger(trigger);
      }
    });

    const choice = response.getResponseText();

    switch(choice) {
      case '1':
        ScriptApp.newTrigger('runFullAnalysis')
          .timeBased()
          .everyDays(1)
          .atHour(2) // 午前2時に実行
          .create();
        ui.alert('毎日午前2時に自動実行するよう設定しました');
        break;
      case '2':
        ScriptApp.newTrigger('runFullAnalysis')
          .timeBased()
          .everyWeeks(1)
          .onWeekDay(ScriptApp.WeekDay.MONDAY)
          .atHour(2)
          .create();
        ui.alert('毎週月曜日午前2時に自動実行するよう設定しました');
        break;
      case '3':
        ScriptApp.newTrigger('runFullAnalysis')
          .timeBased()
          .onMonthDay(1)
          .atHour(2)
          .create();
        ui.alert('毎月1日午前2時に自動実行するよう設定しました');
        break;
      case '0':
        ui.alert('スケジュールを削除しました');
        break;
      default:
        ui.alert('無効な選択です');
    }
  }
}

// Salesforceへの認証
function authenticateToSalesforce() {
  const url = `${CONFIG.LOGIN_URL}/services/oauth2/token`;

  // セキュリティトークンが空の場合はパスワードのみ使用（信頼済みIPの場合）
  const password = CONFIG.SECURITY_TOKEN ?
    CONFIG.PASSWORD + CONFIG.SECURITY_TOKEN :
    CONFIG.PASSWORD;

  const payload = {
    'grant_type': 'password',
    'client_id': CONFIG.CLIENT_ID,
    'client_secret': CONFIG.CLIENT_SECRET,
    'username': CONFIG.USERNAME,
    'password': password
  };

  const options = {
    'method': 'post',
    'payload': payload,
    'muteHttpExceptions': true
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const result = JSON.parse(response.getContentText());

    if (result.access_token) {
      return {
        accessToken: result.access_token,
        instanceUrl: result.instance_url
      };
    } else {
      throw new Error('認証に失敗しました: ' + response.getContentText());
    }
  } catch (e) {
    throw new Error('認証エラー: ' + e.toString());
  }
}

// オブジェクト一覧を取得（UI版）
function fetchObjectList() {
  const result = fetchObjectListSilent();
  SpreadsheetApp.getUi().alert(`${result}個のオブジェクトを取得しました`);
}

// オブジェクト一覧を取得（サイレント版）
function fetchObjectListSilent() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Objects') ||
                SpreadsheetApp.getActiveSpreadsheet().insertSheet('Objects');

  sheet.clear();

  try {
    const auth = authenticateToSalesforce();
    const url = `${auth.instanceUrl}/services/data/v59.0/sobjects`;

    const options = {
      'method': 'get',
      'headers': {
        'Authorization': 'Bearer ' + auth.accessToken
      },
      'muteHttpExceptions': true
    };

    const response = UrlFetchApp.fetch(url, options);
    const result = JSON.parse(response.getContentText());

    // ヘッダー行を作成
    const headers = ['オブジェクト名', 'ラベル', 'API名', 'カスタム', '作成可能', '更新可能', '削除可能', '検索可能'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setBackground('#4285F4').setFontColor('#FFFFFF').setFontWeight('bold');

    // データ行を作成
    const data = result.sobjects.map(obj => [
      obj.name,
      obj.label,
      obj.name,
      obj.custom ? 'はい' : 'いいえ',
      obj.createable ? 'はい' : 'いいえ',
      obj.updateable ? 'はい' : 'いいえ',
      obj.deletable ? 'はい' : 'いいえ',
      obj.searchable ? 'はい' : 'いいえ'
    ]);

    if (data.length > 0) {
      sheet.getRange(2, 1, data.length, headers.length).setValues(data);
    }

    // 列幅を自動調整
    sheet.autoResizeColumns(1, headers.length);

    return data.length;

  } catch (e) {
    logToSheet(`オブジェクト一覧取得エラー: ${e.toString()}`, 'ERROR');
    return 0; // エラー時は0を返す
  }
}

// 選択されたオブジェクトのフィールド情報を取得
function fetchFieldsForSelectedObject() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('オブジェクト名を入力', 'フィールド情報を取得したいオブジェクトのAPI名を入力してください（例: Account, Contact, CustomObject__c）', ui.ButtonSet.OK_CANCEL);

  if (response.getSelectedButton() == ui.Button.OK) {
    const objectName = response.getResponseText();
    fetchFieldsForObject(objectName);
  }
}

// 特定オブジェクトのフィールド情報を取得（UI版）
function fetchFieldsForObject(objectName) {
  const result = fetchFieldsForObjectSilent(objectName);
  SpreadsheetApp.getUi().alert(`${objectName}の${result}個のフィールドを取得しました`);
}

// 特定オブジェクトのフィールド情報を取得（サイレント版）
function fetchFieldsForObjectSilent(objectName) {
  const sheetName = `Fields_${objectName}`;
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName) ||
                SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);

  sheet.clear();

  try {
    const auth = authenticateToSalesforce();
    const url = `${auth.instanceUrl}/services/data/v59.0/sobjects/${objectName}/describe`;

    const options = {
      'method': 'get',
      'headers': {
        'Authorization': 'Bearer ' + auth.accessToken
      },
      'muteHttpExceptions': true
    };

    const response = UrlFetchApp.fetch(url, options);
    const result = JSON.parse(response.getContentText());

    // ヘッダー行を作成
    const headers = ['フィールド名', 'ラベル', 'API名', 'データ型', '長さ', '必須', 'ユニーク', 'カスタム', '作成可能', '更新可能', 'ナイラブル', '参照先'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setBackground('#4285F4').setFontColor('#FFFFFF').setFontWeight('bold');

    // データ行を作成
    const data = result.fields.map(field => [
      field.name,
      field.label,
      field.name,
      field.type,
      field.length || '',
      field.nillable === false ? 'はい' : 'いいえ',
      field.unique ? 'はい' : 'いいえ',
      field.custom ? 'はい' : 'いいえ',
      field.createable ? 'はい' : 'いいえ',
      field.updateable ? 'はい' : 'いいえ',
      field.nillable ? 'はい' : 'いいえ',
      field.referenceTo && field.referenceTo.length > 0 ? field.referenceTo.join(', ') : ''
    ]);

    if (data.length > 0) {
      sheet.getRange(2, 1, data.length, headers.length).setValues(data);
    }

    // 列幅を自動調整
    sheet.autoResizeColumns(1, headers.length);

    return data.length;

  } catch (e) {
    logToSheet(`${objectName}のフィールド取得エラー: ${e.toString()}`, 'ERROR');
    return 0; // エラー時は0を返す
  }
}

// 全データをクリア
function clearAllData() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert('確認', 'すべてのシートのデータをクリアしますか？', ui.ButtonSet.YES_NO);

  if (response == ui.Button.YES) {
    const sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
    sheets.forEach(sheet => {
      if (sheet.getName() !== 'シート1') {
        SpreadsheetApp.getActiveSpreadsheet().deleteSheet(sheet);
      }
    });
    ui.alert('すべてのデータがクリアされました');
  }
}

// 現在の設定を取得
function getCurrentConfig() {
  return {
    clientId: PropertiesService.getScriptProperties().getProperty('SF_CLIENT_ID') || '',
    clientSecret: PropertiesService.getScriptProperties().getProperty('SF_CLIENT_SECRET') || '',
    username: PropertiesService.getScriptProperties().getProperty('SF_USERNAME') || '',
    password: PropertiesService.getScriptProperties().getProperty('SF_PASSWORD') || '',
    securityToken: PropertiesService.getScriptProperties().getProperty('SF_SECURITY_TOKEN') || '',
    loginUrl: PropertiesService.getScriptProperties().getProperty('SF_LOGIN_URL') || 'https://login.salesforce.com'
  };
}

// リレーション分析を実行（UI版）
function analyzeRelationships() {
  const result = analyzeRelationshipsSilent();
  SpreadsheetApp.getUi().alert(`${result}個のリレーションを分析しました`);
}

// リレーション分析を実行（サイレント版）
function analyzeRelationshipsSilent() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Relationships') ||
                SpreadsheetApp.getActiveSpreadsheet().insertSheet('Relationships');

  sheet.clear();

  try {
    const auth = authenticateToSalesforce();
    const url = `${auth.instanceUrl}/services/data/v59.0/sobjects`;

    const options = {
      'method': 'get',
      'headers': {
        'Authorization': 'Bearer ' + auth.accessToken
      },
      'muteHttpExceptions': true
    };

    const response = UrlFetchApp.fetch(url, options);
    const result = JSON.parse(response.getContentText());

    // ヘッダー行を作成
    const headers = ['親オブジェクト', '子オブジェクト', 'リレーション名', 'フィールド名', 'リレーション型', '必須', 'カスケード削除', '親から子への参照名'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setBackground('#4285F4').setFontColor('#FFFFFF').setFontWeight('bold');

    const relationships = [];

    // 主要なオブジェクトのリレーション情報を取得
    const mainObjects = result.sobjects.filter(obj =>
      obj.queryable && !obj.name.endsWith('History') && !obj.name.endsWith('Feed')
    ).slice(0, 50); // API制限を考慮して上限を設定

    mainObjects.forEach(obj => {
      try {
        const describeUrl = `${auth.instanceUrl}/services/data/v59.0/sobjects/${obj.name}/describe`;
        const describeResponse = UrlFetchApp.fetch(describeUrl, {
          'method': 'get',
          'headers': {
            'Authorization': 'Bearer ' + auth.accessToken
          },
          'muteHttpExceptions': true
        });

        const describeResult = JSON.parse(describeResponse.getContentText());

        // リレーションフィールドを抽出
        describeResult.fields.forEach(field => {
          if (field.type === 'reference' && field.referenceTo && field.referenceTo.length > 0) {
            field.referenceTo.forEach(refObject => {
              relationships.push([
                refObject,
                obj.name,
                field.relationshipName || '',
                field.name,
                'Lookup',
                !field.nillable ? 'はい' : 'いいえ',
                field.cascadeDelete ? 'はい' : 'いいえ',
                field.relationshipName || field.name
              ]);
            });
          } else if (field.type === 'masterDetail') {
            field.referenceTo.forEach(refObject => {
              relationships.push([
                refObject,
                obj.name,
                field.relationshipName || '',
                field.name,
                'Master-Detail',
                'はい',
                'はい',
                field.relationshipName || field.name
              ]);
            });
          }
        });

        // 子リレーション情報も追加
        if (describeResult.childRelationships && describeResult.childRelationships.length > 0) {
          describeResult.childRelationships.forEach(childRel => {
            if (childRel.childSObject && childRel.field) {
              relationships.push([
                obj.name,
                childRel.childSObject,
                childRel.relationshipName || '',
                childRel.field,
                childRel.cascadeDelete ? 'Master-Detail' : 'Lookup',
                '',
                childRel.cascadeDelete ? 'はい' : 'いいえ',
                childRel.relationshipName || ''
              ]);
            }
          });
        }

      } catch (e) {
        // 個別のオブジェクトでエラーが発生した場合は続行
        console.log('Error fetching relationships for ' + obj.name + ': ' + e.toString());
      }
    });

    // 重複を削除
    const uniqueRelationships = [];
    const seen = new Set();

    relationships.forEach(rel => {
      const key = rel.join('|');
      if (!seen.has(key)) {
        seen.add(key);
        uniqueRelationships.push(rel);
      }
    });

    if (uniqueRelationships.length > 0) {
      sheet.getRange(2, 1, uniqueRelationships.length, headers.length).setValues(uniqueRelationships);
    }

    sheet.autoResizeColumns(1, headers.length);

    return uniqueRelationships.length;

  } catch (e) {
    logToSheet(`リレーション分析エラー: ${e.toString()}`, 'ERROR');
    return 0; // エラー時は0を返す
  }
}

// データ構造サマリーを作成（UI版）
function createDataStructureSummary() {
  const result = createDataStructureSummarySilent();
  SpreadsheetApp.getUi().alert(`${result}個のオブジェクトの構造を分析しました`);
}

// データ構造サマリーを作成（サイレント版）
function createDataStructureSummarySilent() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DataStructure') ||
                SpreadsheetApp.getActiveSpreadsheet().insertSheet('DataStructure');

  sheet.clear();

  try {
    const auth = authenticateToSalesforce();
    const url = `${auth.instanceUrl}/services/data/v59.0/sobjects`;

    const options = {
      'method': 'get',
      'headers': {
        'Authorization': 'Bearer ' + auth.accessToken
      },
      'muteHttpExceptions': true
    };

    const response = UrlFetchApp.fetch(url, options);
    const result = JSON.parse(response.getContentText());

    // ヘッダー行を作成
    const headers = [
      'オブジェクト名',
      'ラベル',
      '階層レベル',
      '親オブジェクト',
      '子オブジェクト数',
      'フィールド数',
      'リレーション数',
      'レコードタイプ対応',
      'トリガー対応',
      'カスタム/標準'
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setBackground('#4285F4').setFontColor('#FFFFFF').setFontWeight('bold');

    const dataStructure = [];
    const objectHierarchy = {};

    // 分析対象のオブジェクトをフィルタ
    const targetObjects = result.sobjects.filter(obj =>
      obj.queryable && !obj.name.endsWith('History') &&
      !obj.name.endsWith('Feed') && !obj.name.endsWith('Share')
    ).slice(0, 30); // API制限を考慮

    targetObjects.forEach(obj => {
      try {
        const describeUrl = `${auth.instanceUrl}/services/data/v59.0/sobjects/${obj.name}/describe`;
        const describeResponse = UrlFetchApp.fetch(describeUrl, {
          'method': 'get',
          'headers': {
            'Authorization': 'Bearer ' + auth.accessToken
          },
          'muteHttpExceptions': true
        });

        const describeResult = JSON.parse(describeResponse.getContentText());

        // 親オブジェクトを特定
        const parentObjects = [];
        const childCount = describeResult.childRelationships ? describeResult.childRelationships.length : 0;
        let relationCount = 0;

        describeResult.fields.forEach(field => {
          if ((field.type === 'reference' || field.type === 'masterDetail') &&
              field.referenceTo && field.referenceTo.length > 0) {
            relationCount++;
            // Master-Detailの場合は親とみなす
            if (field.type === 'masterDetail' || !field.nillable) {
              parentObjects.push(...field.referenceTo);
            }
          }
        });

        // 階層レベルを判定
        let hierarchyLevel = '独立';
        if (parentObjects.length > 0) {
          hierarchyLevel = '子';
        } else if (childCount > 5) {
          hierarchyLevel = '親';
        } else if (childCount > 0) {
          hierarchyLevel = '親/子';
        }

        dataStructure.push([
          obj.name,
          obj.label,
          hierarchyLevel,
          parentObjects.join(', ') || '-',
          childCount,
          describeResult.fields.length,
          relationCount,
          describeResult.recordTypeInfos && describeResult.recordTypeInfos.length > 1 ? 'はい' : 'いいえ',
          obj.triggerable ? 'はい' : 'いいえ',
          obj.custom ? 'カスタム' : '標準'
        ]);

      } catch (e) {
        console.log('Error analyzing ' + obj.name + ': ' + e.toString());
      }
    });

    // 階層レベルでソート
    dataStructure.sort((a, b) => {
      const levelOrder = {'親': 1, '親/子': 2, '独立': 3, '子': 4};
      return (levelOrder[a[2]] || 5) - (levelOrder[b[2]] || 5);
    });

    if (dataStructure.length > 0) {
      sheet.getRange(2, 1, dataStructure.length, headers.length).setValues(dataStructure);
    }

    // 条件付き書式を適用
    const hierarchyColumn = sheet.getRange(2, 3, dataStructure.length, 1);
    const rules = [];

    // 親オブジェクトを緑色
    const parentRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('親')
      .setBackground('#b7e1cd')
      .setRanges([hierarchyColumn])
      .build();
    rules.push(parentRule);

    // 子オブジェクトを黄色
    const childRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('子')
      .setBackground('#fce5cd')
      .setRanges([hierarchyColumn])
      .build();
    rules.push(childRule);

    sheet.setConditionalFormatRules(rules);
    sheet.autoResizeColumns(1, headers.length);

    // ER図用のMermaid記法を生成
    createERDiagram(targetObjects, auth);

    return dataStructure.length;

  } catch (e) {
    logToSheet(`データ構造分析エラー: ${e.toString()}`, 'ERROR');
    return 0; // エラー時は0を返す
  }
}

// ER図用のMermaid記法を生成
function createERDiagram(objects, auth) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ER_Diagram') ||
                SpreadsheetApp.getActiveSpreadsheet().insertSheet('ER_Diagram');

  sheet.clear();

  const headers = ['ER図 (Mermaid記法)'];
  sheet.getRange(1, 1, 1, 1).setValues([headers]);
  sheet.getRange(1, 1, 1, 1).setBackground('#4285F4').setFontColor('#FFFFFF').setFontWeight('bold');

  let mermaidCode = 'erDiagram\n';
  const processedRelations = new Set();

  // 主要なオブジェクトのリレーションを取得
  const mainObjects = objects.slice(0, 15); // 図が複雑になりすぎないよう制限

  mainObjects.forEach(obj => {
    try {
      const describeUrl = `${auth.instanceUrl}/services/data/v59.0/sobjects/${obj.name}/describe`;
      const describeResponse = UrlFetchApp.fetch(describeUrl, {
        'method': 'get',
        'headers': {
          'Authorization': 'Bearer ' + auth.accessToken
        },
        'muteHttpExceptions': true
      });

      const describeResult = JSON.parse(describeResponse.getContentText());

      // エンティティを追加
      mermaidCode += `    ${obj.name} {\n`;

      // 主要フィールドを追加（最大10個）
      const keyFields = describeResult.fields
        .filter(f => f.nameField || f.unique || f.externalId ||
                     f.type === 'id' || !f.nillable)
        .slice(0, 10);

      keyFields.forEach(field => {
        const fieldType = field.type.replace(/[^a-zA-Z]/g, '');
        const required = !field.nillable ? ' PK' : '';
        mermaidCode += `        ${fieldType} ${field.name}${required}\n`;
      });

      mermaidCode += `    }\n`;

      // リレーションを追加
      describeResult.fields.forEach(field => {
        if (field.type === 'reference' && field.referenceTo && field.referenceTo.length > 0) {
          field.referenceTo.forEach(refObject => {
            const relationKey = `${obj.name}-${refObject}`;
            if (!processedRelations.has(relationKey) && mainObjects.some(o => o.name === refObject)) {
              processedRelations.add(relationKey);
              const cardinality = field.nillable ? '}o--o{' : '}|--||';
              mermaidCode += `    ${obj.name} ${cardinality} ${refObject} : "${field.relationshipName || field.name}"\n`;
            }
          });
        }
      });

    } catch (e) {
      console.log('Error processing ER for ' + obj.name + ': ' + e.toString());
    }
  });

  // Mermaid記法を出力
  const rows = mermaidCode.split('\n').map(line => [line]);
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, 1).setValues(rows);
  }

  // 使用方法の説明を追加
  sheet.getRange(rows.length + 3, 1).setValue('使用方法:');
  sheet.getRange(rows.length + 4, 1).setValue('1. 上記のコードをコピー');
  sheet.getRange(rows.length + 5, 1).setValue('2. https://mermaid.live/ にアクセス');
  sheet.getRange(rows.length + 6, 1).setValue('3. コードを貼り付けてER図を表示');

  sheet.autoResizeColumn(1);
}