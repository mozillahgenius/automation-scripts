/**
 * 反社リスト変換処理システム - メインコード
 * Version: 3.1.0
 *
 * 機能:
 * - 初期セットアップ自動化
 * - 動的ファイル選択
 * - OpenAI API連携
 * - 350件自動分割
 * - CSV出力
 */

// ========================================
// グローバル設定
// ========================================

const OPENAI_API_URL = 'https://api.openai.com/v1/chat/completions';

const CONFIG = {
  batchSize: 350,
  currentSourceFile: null
};

// ========================================
// 初期セットアップ機能
// ========================================

/**
 * 初期セットアップ - 新規スプレッドシート作成
 */
function initialSetup() {
  const ui = SpreadsheetApp.getUi();

  const response = ui.alert(
    '初期セットアップ',
    '反社リスト変換システムの初期セットアップを開始します。\n\n以下の処理を実行します：\n1. 処理用スプレッドシートの作成\n2. 必要なシートの準備\n3. スクリプトプロパティの初期化\n\n続行しますか？',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    return;
  }

  try {
    // 新しいスプレッドシートを作成
    const newSpreadsheet = SpreadsheetApp.create('反社リスト処理_' + Utilities.formatDate(new Date(), 'JST', 'yyyyMMdd'));
    const spreadsheetId = newSpreadsheet.getId();

    // デフォルトシートを設定
    const defaultSheet = newSpreadsheet.getSheets()[0];
    defaultSheet.setName('Dashboard');

    // Dashboardシートの設定
    setupDashboardSheet(defaultSheet);

    // スクリプトをコピー
    const htmlMessage = `
セットアップが完了しました！

新しいスプレッドシートが作成されました：
${newSpreadsheet.getUrl()}

次の手順：
1. 上記URLを開く
2. 拡張機能 → Apps Script を選択
3. このプロジェクトのすべてのコード（.gsファイル）をコピー
4. 保存して実行

※このメッセージをメモしてください。
    `;

    ui.alert('セットアップ完了', htmlMessage, ui.ButtonSet.OK);

    // URLをクリップボードにコピー（可能な場合）
    return newSpreadsheet.getUrl();

  } catch (error) {
    ui.alert('エラー', 'セットアップ中にエラーが発生しました: ' + error.toString(), ui.ButtonSet.OK);
  }
}

/**
 * 現在のスプレッドシートをセットアップ
 */
function setupCurrentSpreadsheet() {
  const ui = SpreadsheetApp.getUi();

  const response = ui.alert(
    'スプレッドシートのセットアップ',
    '現在のスプレッドシートを反社リスト処理用にセットアップします。\n\n続行しますか？',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    return;
  }

  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

    // Dashboardシートを作成または更新
    let dashboardSheet = spreadsheet.getSheetByName('Dashboard');
    if (!dashboardSheet) {
      dashboardSheet = spreadsheet.insertSheet('Dashboard', 0);
    }
    setupDashboardSheet(dashboardSheet);

    // 設定シートを作成
    let configSheet = spreadsheet.getSheetByName('Config');
    if (!configSheet) {
      configSheet = spreadsheet.insertSheet('Config');
    }
    setupConfigSheet(configSheet);

    // 処理履歴シートを作成
    let historySheet = spreadsheet.getSheetByName('History');
    if (!historySheet) {
      historySheet = spreadsheet.insertSheet('History');
    }
    setupHistorySheet(historySheet);

    ui.alert('セットアップ完了', 'スプレッドシートのセットアップが完了しました。', ui.ButtonSet.OK);

  } catch (error) {
    ui.alert('エラー', 'セットアップ中にエラーが発生しました: ' + error.toString(), ui.ButtonSet.OK);
  }
}

/**
 * Dashboardシートをセットアップ
 */
function setupDashboardSheet(sheet) {
  sheet.clear();

  // ヘッダー設定（すべて4列に統一）
  const headers = [
    ['反社リスト変換システム Dashboard', '', '', ''],
    ['', '', '', ''],
    ['処理状況', '', '', ''],
    ['項目', '状態', '詳細', '最終更新'],
    ['現在の元データ', '未選択', '', ''],
    ['Step 1: データ読込', '未実行', '', ''],
    ['Step 2: AI予測', '未実行', '', ''],
    ['Step 3: データ分割', '未実行', '', ''],
    ['Step 4: CSV出力', '未実行', '', ''],
    ['', '', '', ''],
    ['処理統計', '', '', ''],
    ['総処理件数', '0', '', ''],
    ['AI予測件数', '0', '', ''],
    ['作成バッチ数', '0', '', ''],
    ['出力CSVファイル数', '0', '', ''],
    ['', '', '', ''],
    ['システム情報', '', '', ''],
    ['OpenAI APIキー', '未設定', '', ''],
    ['バッチサイズ', '350件', '', ''],
    ['最終処理日時', '', '', '']
  ];

  const range = sheet.getRange(1, 1, headers.length, 4);
  range.setValues(headers);

  // スタイリング
  sheet.getRange(1, 1, 1, 4).merge()
    .setFontSize(16)
    .setFontWeight('bold')
    .setBackground('#1a73e8')
    .setFontColor('#ffffff')
    .setHorizontalAlignment('center');

  sheet.getRange(3, 1).setFontSize(12).setFontWeight('bold');
  sheet.getRange(4, 1, 1, 4).setBackground('#f3f3f3').setFontWeight('bold');

  sheet.getRange(11, 1).setFontSize(12).setFontWeight('bold');
  sheet.getRange(17, 1).setFontSize(12).setFontWeight('bold');

  // 列幅調整
  sheet.setColumnWidth(1, 200);
  sheet.setColumnWidth(2, 150);
  sheet.setColumnWidth(3, 300);
  sheet.setColumnWidth(4, 150);

  // 条件付き書式（ステータスカラー）
  const statusRange = sheet.getRange(5, 2, 5, 1);
  const rules = [];

  // 完了 = 緑
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('完了')
    .setBackground('#b7e1cd')
    .setRanges([statusRange])
    .build());

  // 処理中 = 黄
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('処理中')
    .setBackground('#fce5cd')
    .setRanges([statusRange])
    .build());

  // エラー = 赤
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('エラー')
    .setBackground('#f4c7c3')
    .setRanges([statusRange])
    .build());

  sheet.setConditionalFormatRules(rules);
}

/**
 * 設定シートをセットアップ
 */
function setupConfigSheet(sheet) {
  sheet.clear();

  const config = [
    ['設定項目', '値', '説明'],
    ['OpenAI APIキー', '', 'OpenAI APIのシークレットキー'],
    ['バッチサイズ', '350', '1バッチあたりの件数'],
    ['API モデル', 'gpt-4o-mini', '使用するAIモデル'],
    ['API タイムアウト', '30000', 'API呼び出しのタイムアウト（ミリ秒）'],
    ['CSV文字コード', 'UTF-8 BOM', 'CSV出力時の文字コード'],
    ['自動バックアップ', 'TRUE', 'Master_Dataの自動バックアップ'],
    ['デバッグモード', 'FALSE', '詳細ログの出力']
  ];

  const range = sheet.getRange(1, 1, config.length, 3);
  range.setValues(config);

  // スタイリング
  sheet.getRange(1, 1, 1, 3).setBackground('#f3f3f3').setFontWeight('bold');

  // 列幅調整
  sheet.setColumnWidth(1, 200);
  sheet.setColumnWidth(2, 300);
  sheet.setColumnWidth(3, 400);

  // データ検証（TRUE/FALSE）
  const booleanRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['TRUE', 'FALSE'])
    .build();
  sheet.getRange(7, 2).setDataValidation(booleanRule);
  sheet.getRange(8, 2).setDataValidation(booleanRule);
}

/**
 * 処理履歴シートをセットアップ
 */
function setupHistorySheet(sheet) {
  sheet.clear();

  const headers = [
    ['処理日時', '処理種別', 'ソースファイル', '処理件数', 'AI予測件数', 'バッチ数', 'ステータス', '実行者', '備考']
  ];

  sheet.getRange(1, 1, 1, headers[0].length).setValues(headers);

  // スタイリング
  sheet.getRange(1, 1, 1, headers[0].length)
    .setBackground('#f3f3f3')
    .setFontWeight('bold');

  // 列幅調整
  sheet.setColumnWidth(1, 150); // 処理日時
  sheet.setColumnWidth(2, 120); // 処理種別
  sheet.setColumnWidth(3, 250); // ソースファイル
  sheet.setColumnWidth(4, 100); // 処理件数
  sheet.setColumnWidth(5, 120); // AI予測件数
  sheet.setColumnWidth(6, 100); // バッチ数
  sheet.setColumnWidth(7, 100); // ステータス
  sheet.setColumnWidth(8, 150); // 実行者
  sheet.setColumnWidth(9, 300); // 備考

  // フィルター設定
  sheet.getRange(1, 1, 1, headers[0].length).createFilter();
}

/**
 * Dashboardを更新
 */
function updateDashboard(step, status, details = '') {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    let dashboard = spreadsheet.getSheetByName('Dashboard');

    if (!dashboard) {
      return;
    }

    const now = Utilities.formatDate(new Date(), 'JST', 'yyyy/MM/dd HH:mm:ss');

    // ステップに応じて行を特定
    const stepRows = {
      'source': 5,
      'step1': 6,
      'step2': 7,
      'step3': 8,
      'step4': 9
    };

    const row = stepRows[step];
    if (row) {
      dashboard.getRange(row, 2).setValue(status);
      dashboard.getRange(row, 3).setValue(details);
      dashboard.getRange(row, 4).setValue(now);
    }

    // 最終処理日時を更新
    dashboard.getRange(20, 2).setValue(now);

  } catch (error) {
    console.error('Dashboard更新エラー:', error);
  }
}

/**
 * 処理履歴を追加
 */
function addHistory(processType, sourceFile, processedCount, aiCount, batchCount, status, notes = '') {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    let historySheet = spreadsheet.getSheetByName('History');

    if (!historySheet) {
      return;
    }

    const user = Session.getActiveUser().getEmail();
    const now = Utilities.formatDate(new Date(), 'JST', 'yyyy/MM/dd HH:mm:ss');

    const newRow = [
      now,
      processType,
      sourceFile || '',
      processedCount || 0,
      aiCount || 0,
      batchCount || 0,
      status,
      user,
      notes
    ];

    historySheet.appendRow(newRow);

  } catch (error) {
    console.error('履歴追加エラー:', error);
  }
}

// ========================================
// ファイル選択機能
// ========================================

/**
 * 元データファイルを選択
 */
function selectSourceFile() {
  const ui = SpreadsheetApp.getUi();

  const html = HtmlService.createHtmlOutputFromFile('file-picker')
    .setWidth(700)
    .setHeight(500);

  ui.showModalDialog(html, '元データファイルを選択');
}

/**
 * URLから元データファイルを設定
 */
function setSourceFileByUrl() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    '元データファイルの指定',
    'Google スプレッドシートのURLまたはIDを入力してください:',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() === ui.Button.OK) {
    const input = response.getResponseText();

    try {
      let fileId;

      // URLからIDを抽出
      if (input.includes('/spreadsheets/d/')) {
        const idMatch = input.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
        if (idMatch) {
          fileId = idMatch[1];
        }
      } else {
        // 入力が直接IDの場合
        fileId = input;
      }

      const file = DriveApp.getFileById(fileId);

      // ファイル情報を保存
      const fileInfo = {
        id: fileId,
        name: file.getName(),
        url: `https://docs.google.com/spreadsheets/d/${fileId}/edit`
      };

      saveSelectedFile(fileInfo);
      updateDashboard('source', '選択済', file.getName());

      ui.alert('成功', `元データファイルを設定しました:\n${file.getName()}`, ui.ButtonSet.OK);

    } catch (error) {
      ui.alert('エラー', `ファイルの設定に失敗しました:\n${error.toString()}`, ui.ButtonSet.OK);
    }
  }
}

/**
 * 最近使用したファイルから選択
 */
function selectFromRecentFiles() {
  const ui = SpreadsheetApp.getUi();
  const recentFiles = getRecentSourceFiles();

  if (recentFiles.length === 0) {
    ui.alert('情報', '最近使用したファイルはありません。', ui.ButtonSet.OK);
    return;
  }

  const html = HtmlService.createHtmlOutputFromFile('recent-files')
    .setWidth(600)
    .setHeight(400);

  ui.showModalDialog(html, '最近使用したファイル');
}

/**
 * Google Picker APIを使用してファイルを選択（Webアプリ用）
 */
function getPickerOAuthToken() {
  return ScriptApp.getOAuthToken();
}

/**
 * 選択されたファイル情報を保存
 */
function saveSelectedFile(fileInfo) {
  const props = PropertiesService.getUserProperties();

  // 現在のファイルとして保存
  props.setProperty('CURRENT_SOURCE_FILE', JSON.stringify(fileInfo));

  // 最近使用したファイルリストに追加
  let recentFiles = JSON.parse(props.getProperty('RECENT_SOURCE_FILES') || '[]');

  // 既存のエントリを削除（重複防止）
  recentFiles = recentFiles.filter(f => f.id !== fileInfo.id);

  // 先頭に追加
  recentFiles.unshift(fileInfo);

  // 最大10件まで保持
  recentFiles = recentFiles.slice(0, 10);

  props.setProperty('RECENT_SOURCE_FILES', JSON.stringify(recentFiles));

  return fileInfo;
}

/**
 * 現在選択中のファイル情報を取得
 */
function getCurrentSourceFile() {
  const props = PropertiesService.getUserProperties();
  const fileStr = props.getProperty('CURRENT_SOURCE_FILE');

  if (!fileStr) {
    return null;
  }

  return JSON.parse(fileStr);
}

/**
 * 最近使用したファイルリストを取得
 */
function getRecentSourceFiles() {
  const props = PropertiesService.getUserProperties();
  const filesStr = props.getProperty('RECENT_SOURCE_FILES');

  if (!filesStr) {
    return [];
  }

  return JSON.parse(filesStr);
}

// ========================================
// メイン処理機能（Step 1-4）
// ========================================

/**
 * Step 1: データ読み込みと基本変換
 */
function step1_loadAndTransform() {
  try {
    const ui = SpreadsheetApp.getUi();
    updateDashboard('step1', '処理中', 'データ読み込み開始');

    // 現在選択中のファイルを確認
    const sourceFile = getCurrentSourceFile();

    if (!sourceFile) {
      updateDashboard('step1', 'エラー', '元データファイル未選択');
      const response = ui.alert(
        '確認',
        '元データファイルが選択されていません。\n選択画面を開きますか？',
        ui.ButtonSet.YES_NO
      );

      if (response === ui.Button.YES) {
        selectSourceFile();
      }
      return;
    }

    // ファイル情報を表示して確認
    const confirmResponse = ui.alert(
      '確認',
      `以下のファイルからデータを読み込みます:\n\n${sourceFile.name}\n\n続行しますか？`,
      ui.ButtonSet.YES_NO
    );

    if (confirmResponse !== ui.Button.YES) {
      updateDashboard('step1', '未実行', 'キャンセル');
      return;
    }

    // 既存の変換データがあるか確認
    const targetSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    let masterSheet = targetSpreadsheet.getSheetByName('Master_Data');

    if (masterSheet) {
      const response = ui.alert(
        '確認',
        '既存のデータが存在します。上書きしますか？',
        ui.ButtonSet.YES_NO
      );

      if (response !== ui.Button.YES) {
        updateDashboard('step1', '未実行', 'キャンセル');
        ui.alert('処理をキャンセルしました。');
        return;
      }
    }

    // ソースファイルを開く
    const sourceSpreadsheet = SpreadsheetApp.openById(sourceFile.id);

    // シート名を検索
    const sheetNames = ['基本データ', 'Sheet1', 'データ', 'Data'];
    let sourceSheet = null;

    for (const sheetName of sheetNames) {
      try {
        sourceSheet = sourceSpreadsheet.getSheetByName(sheetName);
        if (sourceSheet) break;
      } catch (e) {
        // continue
      }
    }

    if (!sourceSheet) {
      sourceSheet = sourceSpreadsheet.getSheets()[0];
    }

    const sourceData = sourceSheet.getDataRange().getValues();
    const processedData = [];
    const today = Utilities.formatDate(new Date(), 'JST', 'yyyy-MM-dd');

    // ヘッダー行を設定
    const targetHeaders = [
      '등록일자',
      '고객구분',
      '한글명',
      '영문명',
      '성별',
      '생년월일(설립일)',
      '국적',
      '사용여부',
      '출처',
      '거주지',
      '비고',
      'フリガナ元データ',
      'ソースファイル'
    ];
    processedData.push(targetHeaders);

    // カラムインデックスを自動検出
    const headers = sourceData[0];
    const columnMapping = detectColumns(headers);

    // データを処理
    let processedCount = 0;
    for (let i = 1; i < sourceData.length; i++) {
      const row = sourceData[i];

      // 空行はスキップ
      if (!row[columnMapping.name]) continue;

      const processedRow = processDataRow(row, columnMapping, today, sourceFile.name);
      processedData.push(processedRow);
      processedCount++;
    }

    // Master_Dataシートに保存
    if (!masterSheet) {
      masterSheet = targetSpreadsheet.insertSheet('Master_Data');
    } else {
      masterSheet.clear();
    }

    masterSheet.getRange(1, 1, processedData.length, processedData[0].length).setValues(processedData);

    // 顧客区分列を文字列形式に設定
    const customerTypeRange = masterSheet.getRange(2, 2, processedData.length - 1, 1);
    customerTypeRange.setNumberFormat('@');

    // Dashboard更新
    updateDashboard('step1', '完了', `${processedCount}件処理`);

    // 統計情報更新
    const dashboard = targetSpreadsheet.getSheetByName('Dashboard');
    if (dashboard) {
      dashboard.getRange(12, 2).setValue(processedCount);
    }

    // 履歴追加
    addHistory('データ読込', sourceFile.name, processedCount, 0, 0, '成功');

    // フリガナがない件数を確認
    let missingKanaCount = 0;
    for (let i = 1; i < processedData.length; i++) {
      if (!processedData[i][11] && processedData[i][2]) {
        missingKanaCount++;
      }
    }

    if (missingKanaCount > 0) {
      ui.alert('Step 1 完了', `${processedCount}件のデータを処理しました。\n\n${missingKanaCount}件のデータにフリガナがありません。\nStep 2でAI予測を実行してください。`, ui.ButtonSet.OK);
    } else {
      ui.alert('Step 1 完了', `${processedCount}件のデータを処理しました。`, ui.ButtonSet.OK);
    }

  } catch (error) {
    updateDashboard('step1', 'エラー', error.toString());
    console.error('Step 1 エラー:', error);
    SpreadsheetApp.getUi().alert('エラー', 'データの読み込み中にエラーが発生しました: ' + error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * データ行を処理
 */
function processDataRow(row, columnMapping, today, sourceFileName) {
  const processedRow = [];

  // 1. 登録日付
  processedRow.push(today);

  // 2. 顧客区分
  const orgName = columnMapping.org !== -1 ? row[columnMapping.org] || '' : '';
  const isOrganization = orgName.includes('組') || orgName.includes('会') || orgName.includes('団');
  processedRow.push(isOrganization ? '02' : '01');

  // 3. 日本語名
  processedRow.push(row[columnMapping.name] || '');

  // 4. 英文名
  const kanaName = columnMapping.kana !== -1 ? row[columnMapping.kana] || '' : '';
  processedRow.push(kanaName ? convertKanaToRomaji(kanaName) : '');

  // 5. 性別
  const gender = columnMapping.gender !== -1 ? row[columnMapping.gender] || '' : '';
  let genderCode = '';
  if (gender === '男' || gender === 'M' || gender === '1') genderCode = '1';
  else if (gender === '女' || gender === 'F' || gender === '2') genderCode = '2';
  processedRow.push(genderCode);

  // 6. 生年月日
  const age = columnMapping.age !== -1 ? row[columnMapping.age] : null;
  let birthYear = '';
  if (age && !isNaN(age)) {
    birthYear = String(2025 - parseInt(age));
  }
  processedRow.push(birthYear);

  // 7. 国籍
  processedRow.push('JP');

  // 8. 使用有無
  processedRow.push('Y');

  // 9. 出典
  processedRow.push('暴力団追放運動推進都民センター\n폭력단 추방운동추진 도민센터\nAnti-Organized Crime Campaign Center of Tokyo');

  // 10. 居住地
  processedRow.push(columnMapping.address !== -1 ? row[columnMapping.address] || '' : '');

  // 11. 備考
  const remarks = [];
  if (columnMapping.alias !== -1 && row[columnMapping.alias]) {
    remarks.push('異名: ' + row[columnMapping.alias]);
  }
  if (orgName) {
    remarks.push('組織: ' + orgName);
  }
  if (columnMapping.content !== -1 && row[columnMapping.content]) {
    remarks.push(row[columnMapping.content]);
  }
  processedRow.push(remarks.join(' / '));

  // 12. フリガナ元データ
  processedRow.push(kanaName);

  // 13. ソースファイル
  processedRow.push(sourceFileName);

  return processedRow;
}

/**
 * カラムマッピングを自動検出
 */
function detectColumns(headers) {
  const mapping = {
    name: -1,
    kana: -1,
    gender: -1,
    age: -1,
    org: -1,
    address: -1,
    alias: -1,
    content: -1
  };

  for (let i = 0; i < headers.length; i++) {
    const header = String(headers[i] || '').toLowerCase();

    if (header.includes('名前') || header.includes('氏名') || header === 'name') {
      mapping.name = i;
    } else if (header.includes('ふりがな') || header.includes('フリガナ') || header.includes('カナ')) {
      mapping.kana = i;
    } else if (header.includes('性別') || header === 'gender') {
      mapping.gender = i;
    } else if (header.includes('年齢') || header === 'age') {
      mapping.age = i;
    } else if (header.includes('組織') || header.includes('団体')) {
      mapping.org = i;
    } else if (header.includes('住所') || header.includes('居住')) {
      mapping.address = i;
    } else if (header.includes('異名') || header.includes('別名')) {
      mapping.alias = i;
    } else if (header.includes('内容') || header.includes('備考')) {
      mapping.content = i;
    }
  }

  return mapping;
}

// 続きは次のメッセージで...