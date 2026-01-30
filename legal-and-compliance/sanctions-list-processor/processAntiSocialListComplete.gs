/**
 * 反社リスト変換処理システム - 完全統合版
 * Version: 4.0.0
 *
 * このファイルには全ての機能が統合されています
 * - 初期セットアップ
 * - 動的ファイル選択
 * - OpenAI API連携
 * - 350件ごとの個別スプレッドシート作成
 * - Dashboard/Config/History管理
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
 * 初期セットアップ - 現在のスプレッドシートで実行
 */
function initialSetup() {
  const ui = SpreadsheetApp.getUi();

  const response = ui.alert(
    '初期セットアップ',
    '反社リスト変換システムの初期セットアップを開始します。\n\n現在のスプレッドシートに以下のシートを作成します：\n• Dashboard（処理状況管理）\n• Config（設定管理）\n• History（履歴管理）\n\n続行しますか？',
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
      // 最初のシートをDashboardとして使用（通常は「シート1」）
      const sheets = spreadsheet.getSheets();
      if (sheets.length > 0 && sheets[0].getName() === 'シート1') {
        dashboardSheet = sheets[0];
        dashboardSheet.setName('Dashboard');
      } else {
        dashboardSheet = spreadsheet.insertSheet('Dashboard', 0);
      }
    }
    setupDashboardSheet(dashboardSheet);

    // Config シートを作成
    let configSheet = spreadsheet.getSheetByName('Config');
    if (!configSheet) {
      configSheet = spreadsheet.insertSheet('Config');
    }
    setupConfigSheet(configSheet);

    // History シートを作成
    let historySheet = spreadsheet.getSheetByName('History');
    if (!historySheet) {
      historySheet = spreadsheet.insertSheet('History');
    }
    setupHistorySheet(historySheet);

    // HTMLファイル作成ガイドを表示（実際のファイル作成はApps Script上で手動で行う）
    showHTMLCreationGuide();

    // セットアップ完了メッセージ
    const htmlMessage = `
初期セットアップが完了しました！

必要なシートとHTMLファイルが作成されました：
• Dashboard - 処理状況の確認
• Config - システム設定
• History - 処理履歴
• HTMLファイル - ダイアログ用

次の手順：
1. メニューから各機能を実行
2. Step 1: ファイルを選択して読み込み
3. Step 2: AI予測を実行（必要な場合）
4. Step 3: 350件ごとにスプレッドシート作成
    `;

    ui.alert('セットアップ完了', htmlMessage, ui.ButtonSet.OK);

  } catch (error) {
    ui.alert('エラー', 'セットアップ中にエラーが発生しました: ' + error.toString(), ui.ButtonSet.OK);
  }
}

/**
 * 現在のスプレッドシートをセットアップ（互換性のため残す）
 */
function setupCurrentSpreadsheet() {
  // initialSetupと同じ処理を実行
  initialSetup();
}

/**
 * Dashboardシートをセットアップ
 */
function setupDashboardSheet(sheet) {
  sheet.clear();

  const headers = [
    ['反社リスト変換システム Dashboard', '', '', ''],
    ['', '', '', ''],
    ['処理状況', '', '', ''],
    ['項目', '状態', '詳細', '最終更新'],
    ['現在の元データ', '未選択', '', ''],
    ['Step 1: データ読込', '未実行', '', ''],
    ['Step 2: AI予測', '未実行', '', ''],
    ['Step 3: データ分割', '未実行', '', ''],
    ['', '', '', ''],
    ['処理統計', '', '', ''],
    ['総処理件数', '0', '', ''],
    ['AI予測件数', '0', '', ''],
    ['作成ファイル数', '0', '', ''],
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

  sheet.getRange(10, 1).setFontSize(12).setFontWeight('bold');
  sheet.getRange(15, 1).setFontSize(12).setFontWeight('bold');

  // 列幅調整
  sheet.setColumnWidth(1, 200);
  sheet.setColumnWidth(2, 150);
  sheet.setColumnWidth(3, 300);
  sheet.setColumnWidth(4, 150);

  // 条件付き書式（ステータスカラー）
  const statusRange = sheet.getRange(5, 2, 4, 1);
  const rules = [];

  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('完了')
    .setBackground('#b7e1cd')
    .setRanges([statusRange])
    .build());

  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('処理中')
    .setBackground('#fce5cd')
    .setRanges([statusRange])
    .build());

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
  sheet.getRange(6, 2).setDataValidation(booleanRule);
  sheet.getRange(7, 2).setDataValidation(booleanRule);
}

/**
 * HTMLファイルを作成
 */
function createHTMLFiles() {
  // file-picker.htmlの内容
  const filePickerHTML = HtmlService.createHtmlOutputFromFile('file-picker')
    .getContent();

  // folder-picker.htmlの内容
  const folderPickerHTML = HtmlService.createHtmlOutputFromFile('folder-picker')
    .getContent();

  // recent-files.htmlの内容
  const recentFilesHTML = HtmlService.createHtmlOutputFromFile('recent-files')
    .getContent();
}

/**
 * HTMLファイル作成ガイド
 */
function showHTMLCreationGuide() {
  const ui = SpreadsheetApp.getUi();

  const guide = `
HTMLファイル作成手順：

1. 拡張機能 → Apps Script を開く

2. 左側の「＋」ボタンから「HTML」を選択

3. 以下の3つのHTMLファイルを作成：
   • file-picker.html
   • folder-picker.html
   • recent-files.html

4. 各ファイルの内容：
   このプロジェクトフォルダ内の同名のHTMLファイルの
   内容をコピー＆ペーストしてください。

5. 保存（Ctrl+S または Cmd+S）

6. 完了後、メニューから各機能をお使いください。

注意事項：
• ファイル名は正確に入力してください
• .html拡張子は自動的に付きます
• 内容は提供されたものをそのまま使用してください
  `;

  ui.alert('HTMLファイル作成ガイド', guide, ui.ButtonSet.OK);
}

/**
 * 処理履歴シートをセットアップ
 */
function setupHistorySheet(sheet) {
  sheet.clear();

  const headers = [
    ['処理日時', '処理種別', 'ソースファイル', '処理件数', 'AI予測件数', 'ファイル数', 'ステータス', '実行者', '備考']
  ];

  sheet.getRange(1, 1, 1, headers[0].length).setValues(headers);

  // スタイリング
  sheet.getRange(1, 1, 1, headers[0].length)
    .setBackground('#f3f3f3')
    .setFontWeight('bold');

  // 列幅調整
  sheet.setColumnWidth(1, 150);
  sheet.setColumnWidth(2, 120);
  sheet.setColumnWidth(3, 250);
  sheet.setColumnWidth(4, 100);
  sheet.setColumnWidth(5, 120);
  sheet.setColumnWidth(6, 100);
  sheet.setColumnWidth(7, 100);
  sheet.setColumnWidth(8, 150);
  sheet.setColumnWidth(9, 300);

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

    const stepRows = {
      'source': 5,
      'step1': 6,
      'step2': 7,
      'step3': 8
    };

    const row = stepRows[step];
    if (row) {
      dashboard.getRange(row, 2).setValue(status);
      dashboard.getRange(row, 3).setValue(details);
      dashboard.getRange(row, 4).setValue(now);
    }

    dashboard.getRange(18, 2).setValue(now);

  } catch (error) {
    console.error('Dashboard更新エラー:', error);
  }
}

/**
 * 処理履歴を追加
 */
function addHistory(processType, sourceFile, processedCount, aiCount, fileCount, status, notes = '') {
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
      fileCount || 0,
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
 * 元データファイルを選択（ファイルピッカー）
 */
function selectSourceFile() {
  const ui = SpreadsheetApp.getUi();

  try {
    // ドライブのルートフォルダからファイル一覧を取得
    const files = DriveApp.getFilesByType(MimeType.GOOGLE_SHEETS);
    let fileList = [];
    let count = 0;

    while (files.hasNext() && count < 50) {
      const file = files.next();
      fileList.push({
        id: file.getId(),
        name: file.getName(),
        lastUpdated: file.getLastUpdated()
      });
      count++;
    }

    if (fileList.length === 0) {
      ui.alert('情報', 'Google スプレッドシートが見つかりませんでした。', ui.ButtonSet.OK);
      return;
    }

    // ファイル選択ダイアログのHTMLを生成
    let html = '<div style="padding: 10px;">';
    html += '<h3>元データファイルを選択してください:</h3>';
    html += '<select id="fileSelect" size="10" style="width: 100%; height: 300px;">';

    fileList.forEach(file => {
      html += `<option value="${file.id}">${file.name}</option>`;
    });

    html += '</select><br><br>';
    html += '<button onclick="selectFile()">選択</button>';
    html += '<button onclick="google.script.host.close()">キャンセル</button>';
    html += '<script>';
    html += 'function selectFile() {';
    html += '  var select = document.getElementById("fileSelect");';
    html += '  if (select.selectedIndex >= 0) {';
    html += '    var fileId = select.options[select.selectedIndex].value;';
    html += '    var fileName = select.options[select.selectedIndex].text;';
    html += '    google.script.run.withSuccessHandler(function() {';
    html += '      google.script.host.close();';
    html += '    }).saveSelectedFileFromPicker(fileId, fileName);';
    html += '  } else {';
    html += '    alert("ファイルを選択してください");';
    html += '  }';
    html += '}';
    html += '</script>';
    html += '</div>';

    const htmlOutput = HtmlService.createHtmlOutput(html)
      .setWidth(600)
      .setHeight(450);

    ui.showModalDialog(htmlOutput, '元データファイルを選択');

  } catch (error) {
    ui.alert('エラー', 'ファイル選択中にエラーが発生しました: ' + error.toString(), ui.ButtonSet.OK);
  }
}

/**
 * ファイルピッカーから選択されたファイルを保存
 */
function saveSelectedFileFromPicker(fileId, fileName) {
  const fileInfo = {
    id: fileId,
    name: fileName,
    url: `https://docs.google.com/spreadsheets/d/${fileId}/edit`
  };

  saveSelectedFile(fileInfo);
  updateDashboard('source', '選択済', fileName);

  SpreadsheetApp.getUi().alert('成功', `元データファイルを設定しました:\n${fileName}`, SpreadsheetApp.getUi().ButtonSet.OK);
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
        fileId = input;
      }

      const file = DriveApp.getFileById(fileId);

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

  // HTMLでリストを作成
  let html = '<div style="padding: 10px;">';
  html += '<h3>最近使用したファイル:</h3>';
  html += '<select id="fileSelect" size="10" style="width: 100%; height: 300px;">';

  recentFiles.forEach((file, index) => {
    html += `<option value="${index}">${file.name}</option>`;
  });

  html += '</select><br><br>';
  html += '<button onclick="selectRecentFile()">選択</button>';
  html += '<button onclick="google.script.host.close()">キャンセル</button>';
  html += '<script>';
  html += 'function selectRecentFile() {';
  html += '  var select = document.getElementById("fileSelect");';
  html += '  if (select.selectedIndex >= 0) {';
  html += '    var index = select.options[select.selectedIndex].value;';
  html += '    google.script.run.withSuccessHandler(function() {';
  html += '      google.script.host.close();';
  html += '    }).selectRecentFileByIndex(index);';
  html += '  }';
  html += '}';
  html += '</script>';
  html += '</div>';

  const htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(500)
    .setHeight(400);

  ui.showModalDialog(htmlOutput, '最近使用したファイル');
}

/**
 * 最近使用したファイルをインデックスで選択
 */
function selectRecentFileByIndex(index) {
  const recentFiles = getRecentSourceFiles();
  if (index >= 0 && index < recentFiles.length) {
    const file = recentFiles[index];
    saveSelectedFile(file);
    updateDashboard('source', '選択済', file.name);
    SpreadsheetApp.getUi().alert('成功', `元データファイルを設定しました:\n${file.name}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * 選択されたファイル情報を保存
 */
function saveSelectedFile(fileInfo) {
  const props = PropertiesService.getUserProperties();

  props.setProperty('CURRENT_SOURCE_FILE', JSON.stringify(fileInfo));

  let recentFiles = JSON.parse(props.getProperty('RECENT_SOURCE_FILES') || '[]');
  recentFiles = recentFiles.filter(f => f.id !== fileInfo.id);
  recentFiles.unshift(fileInfo);
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

/**
 * 現在の選択を表示
 */
function showCurrentSelection() {
  const ui = SpreadsheetApp.getUi();
  const currentFile = getCurrentSourceFile();

  if (currentFile) {
    ui.alert(
      '現在の選択',
      `ファイル名: ${currentFile.name}\n\nファイルID: ${currentFile.id}\n\nURL: ${currentFile.url || 'N/A'}`,
      ui.ButtonSet.OK
    );
  } else {
    ui.alert('情報', '元データファイルが選択されていません。', ui.ButtonSet.OK);
  }
}

// ========================================
// Step 1: データ読み込みと基本変換
// ========================================

/**
 * Step 1: データ読み込みと基本変換
 */
function step1_loadAndTransform() {
  try {
    const ui = SpreadsheetApp.getUi();
    updateDashboard('step1', '処理中', 'データ読み込み開始');

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

    const confirmResponse = ui.alert(
      '確認',
      `以下のファイルからデータを読み込みます:\n\n${sourceFile.name}\n\n続行しますか？`,
      ui.ButtonSet.YES_NO
    );

    if (confirmResponse !== ui.Button.YES) {
      updateDashboard('step1', '未実行', 'キャンセル');
      return;
    }

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
        return;
      }
    }

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

    const headers = sourceData[0];
    const columnMapping = detectColumns(headers);

    // データを処理
    let processedCount = 0;
    for (let i = 1; i < sourceData.length; i++) {
      const row = sourceData[i];

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

    const customerTypeRange = masterSheet.getRange(2, 2, processedData.length - 1, 1);
    customerTypeRange.setNumberFormat('@');

    updateDashboard('step1', '完了', `${processedCount}件処理`);

    const dashboard = targetSpreadsheet.getSheetByName('Dashboard');
    if (dashboard) {
      dashboard.getRange(11, 2).setValue(processedCount);
    }

    addHistory('データ読込', sourceFile.name, processedCount, 0, 0, '成功');

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

  processedRow.push(today);

  // 顧客分類は全て"01"（テキスト形式）
  processedRow.push('"01"');

  processedRow.push(row[columnMapping.name] || '');

  const kanaName = columnMapping.kana !== -1 ? row[columnMapping.kana] || '' : '';
  processedRow.push(kanaName ? convertKanaToRomaji(kanaName) : '');

  const gender = columnMapping.gender !== -1 ? row[columnMapping.gender] || '' : '';
  let genderCode = '';
  if (gender === '男' || gender === 'M' || gender === '1') genderCode = '"01"';
  else if (gender === '女' || gender === 'F' || gender === '2') genderCode = '"02"';
  processedRow.push(genderCode);

  const age = columnMapping.age !== -1 ? row[columnMapping.age] : null;
  let birthYear = '';
  if (age && !isNaN(age)) {
    birthYear = String(2025 - parseInt(age));
  }
  processedRow.push(birthYear);

  processedRow.push('JP');

  processedRow.push('Y');

  processedRow.push('暴力団追放運動推進都民センター\n폭력단 추방운동추진 도민센터\nAnti-Organized Crime Campaign Center of Tokyo');

  // 住所欄は全てJPに統一
  processedRow.push('JP');

  const remarks = [];

  // 住所（住所詳細を備考に記載）
  const address = columnMapping.address !== -1 ? row[columnMapping.address] || '' : '';
  if (address) {
    remarks.push('住所: ' + address);
  }

  // 異名
  if (columnMapping.alias !== -1 && row[columnMapping.alias]) {
    remarks.push('異名: ' + row[columnMapping.alias]);
  }

  // 年齢（既に生年月日に変換済みだが、元の年齢も記録）
  if (age) {
    remarks.push('年齢: ' + age);
  }

  // 組織・団体
  const orgName = columnMapping.org !== -1 ? row[columnMapping.org] || '' : '';
  if (orgName) {
    remarks.push('組織・団体: ' + orgName);
  }

  // 内容
  if (columnMapping.content !== -1 && row[columnMapping.content]) {
    remarks.push('内容: ' + row[columnMapping.content]);
  }

  // 資料作成年月日
  if (columnMapping.createDate !== -1 && row[columnMapping.createDate]) {
    remarks.push('資料作成年月日: ' + row[columnMapping.createDate]);
  }

  // 県名
  if (columnMapping.prefecture !== -1 && row[columnMapping.prefecture]) {
    remarks.push('県名: ' + row[columnMapping.prefecture]);
  }

  // 更新年月日
  if (columnMapping.updateDate !== -1 && row[columnMapping.updateDate]) {
    remarks.push('更新年月日: ' + row[columnMapping.updateDate]);
  }

  // 新聞社名等
  if (columnMapping.newspaper !== -1 && row[columnMapping.newspaper]) {
    remarks.push('新聞社名等: ' + row[columnMapping.newspaper]);
  }

  // 刊行区分
  if (columnMapping.publicationType !== -1 && row[columnMapping.publicationType]) {
    remarks.push('刊行区分: ' + row[columnMapping.publicationType]);
  }

  // 区分
  if (columnMapping.category !== -1 && row[columnMapping.category]) {
    remarks.push('区分: ' + row[columnMapping.category]);
  }

  // 番号
  if (columnMapping.number !== -1 && row[columnMapping.number]) {
    remarks.push('番号: ' + row[columnMapping.number]);
  }

  processedRow.push(remarks.join(' / '));

  processedRow.push(kanaName);

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
    content: -1,
    createDate: -1,     // 資料作成年月日
    prefecture: -1,     // 県名
    updateDate: -1,     // 更新年月日
    newspaper: -1,      // 新聞社名等
    publicationType: -1, // 刊行区分
    category: -1,       // 区分
    number: -1          // 番号
  };

  for (let i = 0; i < headers.length; i++) {
    const header = String(headers[i] || '').toLowerCase();

    if (header.includes('名前') || header.includes('氏名') || header === 'name') {
      mapping.name = i;
    } else if (header.includes('ふりがな') || header.includes('フリガナ') || header.includes('カナ') || header.includes('ﾌﾘｶﾞﾅ')) {
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
    } else if (header.includes('内容') || (header.includes('備考') && !header.includes('年月日'))) {
      mapping.content = i;
    } else if (header.includes('資料作成年月日')) {
      mapping.createDate = i;
    } else if (header.includes('県名')) {
      mapping.prefecture = i;
    } else if (header.includes('更新年月日')) {
      mapping.updateDate = i;
    } else if (header.includes('新聞社名')) {
      mapping.newspaper = i;
    } else if (header.includes('刊行区分')) {
      mapping.publicationType = i;
    } else if (header === '区分' || header.includes('区分') && !header.includes('刊行')) {
      mapping.category = i;
    } else if (header === '番号' || header.includes('番号')) {
      mapping.number = i;
    }
  }

  return mapping;
}

// ========================================
// Step 2: AI予測処理
// ========================================

/**
 * Step 2: フリガナがないデータをAIで予測
 */
function step2_predictWithAI() {
  try {
    const ui = SpreadsheetApp.getUi();
    updateDashboard('step2', '処理中', 'AI予測開始');

    const targetSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const masterSheet = targetSpreadsheet.getSheetByName('Master_Data');

    if (!masterSheet) {
      updateDashboard('step2', 'エラー', 'Master_Data未検出');
      ui.alert('エラー', 'Master_Dataシートが見つかりません。Step 1を先に実行してください。', ui.ButtonSet.OK);
      return;
    }

    const data = masterSheet.getDataRange().getValues();
    const namesToProcess = [];
    const rowIndices = [];

    for (let i = 1; i < data.length; i++) {
      const kanjiName = data[i][2];
      const kanaData = data[i][11];

      if (kanjiName && !kanaData) {
        namesToProcess.push(kanjiName);
        rowIndices.push(i);
      }
    }

    if (namesToProcess.length === 0) {
      updateDashboard('step2', '完了', '予測不要');
      ui.alert('情報', '予測が必要なデータはありません。', ui.ButtonSet.OK);
      return;
    }

    const response = ui.alert(
      '確認',
      `${namesToProcess.length}件の名前でフリガナを予測します。\n推定コスト: $${(namesToProcess.length * 0.00005).toFixed(2)}\n続行しますか？`,
      ui.ButtonSet.YES_NO
    );

    if (response !== ui.Button.YES) {
      updateDashboard('step2', '未実行', 'キャンセル');
      return;
    }

    console.log(`${namesToProcess.length}件の名前でフリガナを予測開始`);
    const predictions = batchPredictKanaNames(namesToProcess);

    let successCount = 0;
    for (let j = 0; j < rowIndices.length; j++) {
      const rowIndex = rowIndices[j];
      const predictedKana = predictions[j];

      if (predictedKana) {
        const romaji = convertKanaToRomaji(predictedKana);
        masterSheet.getRange(rowIndex + 1, 4).setValue(romaji);
        masterSheet.getRange(rowIndex + 1, 12).setValue(predictedKana + ' (AI予測)');

        const currentRemarks = data[rowIndex][10];
        const updatedRemarks = currentRemarks ?
          currentRemarks + ' / ※フリガナはAI予測' :
          '※フリガナはAI予測';
        masterSheet.getRange(rowIndex + 1, 11).setValue(updatedRemarks);

        successCount++;
      }
    }

    updateDashboard('step2', '完了', `${successCount}件予測完了`);

    const dashboard = targetSpreadsheet.getSheetByName('Dashboard');
    if (dashboard) {
      dashboard.getRange(12, 2).setValue(successCount);
    }

    const sourceFile = getCurrentSourceFile();
    addHistory('AI予測', sourceFile ? sourceFile.name : '', 0, successCount, 0, '成功');

    ui.alert('Step 2 完了', `${successCount}件のフリガナ予測が完了しました。`, ui.ButtonSet.OK);

  } catch (error) {
    updateDashboard('step2', 'エラー', error.toString());
    console.error('Step 2 エラー:', error);
    SpreadsheetApp.getUi().alert('エラー', 'AI予測中にエラーが発生しました: ' + error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

// ========================================
// Step 3: データ分割処理（個別スプレッドシート作成）
// ========================================

/**
 * Step 3: データを350件ずつに分割して個別スプレッドシートとして保存
 */
function step3_splitToSpreadsheets() {
  try {
    const ui = SpreadsheetApp.getUi();
    updateDashboard('step3', '処理中', 'データ分割開始');

    const targetSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const masterSheet = targetSpreadsheet.getSheetByName('Master_Data');

    if (!masterSheet) {
      updateDashboard('step3', 'エラー', 'Master_Data未検出');
      ui.alert('エラー', 'Master_Dataシートが見つかりません。Step 1を先に実行してください。', ui.ButtonSet.OK);
      return;
    }

    const data = masterSheet.getDataRange().getValues();
    const headers = data[0].slice(0, 11);

    const currentFile = DriveApp.getFileById(targetSpreadsheet.getId());
    const parentFolder = currentFile.getParents().next();

    const timestamp = Utilities.formatDate(new Date(), 'JST', 'yyyyMMdd_HHmmss');
    const sourceFile = getCurrentSourceFile();
    const folderName = sourceFile ?
      `反社リスト_${sourceFile.name.replace(/\.[^/.]+$/, '')}_${timestamp}` :
      `反社リスト_分割データ_${timestamp}`;

    let outputFolder = null;
    const folders = parentFolder.getFoldersByName(folderName);
    if (folders.hasNext()) {
      const response = ui.alert(
        '確認',
        `同名のフォルダが存在します。上書きしますか？\nフォルダ名: ${folderName}`,
        ui.ButtonSet.YES_NO
      );

      if (response === ui.Button.YES) {
        outputFolder = folders.next();
        const existingFiles = outputFolder.getFiles();
        while (existingFiles.hasNext()) {
          existingFiles.next().setTrashed(true);
        }
      } else {
        updateDashboard('step3', '未実行', 'キャンセル');
        return;
      }
    } else {
      outputFolder = parentFolder.createFolder(folderName);
    }

    const configSheet = targetSpreadsheet.getSheetByName('Config');
    let batchSize = 350;
    if (configSheet) {
      const configData = configSheet.getDataRange().getValues();
      for (let i = 1; i < configData.length; i++) {
        if (configData[i][0] === 'バッチサイズ') {
          batchSize = parseInt(configData[i][1]) || 350;
          break;
        }
      }
    }

    let sheetIndex = 1;
    let totalProcessed = 0;
    const createdFiles = [];

    for (let i = 1; i < data.length; i += batchSize) {
      const batchData = [headers];
      const endIndex = Math.min(i + batchSize, data.length);

      for (let j = i; j < endIndex; j++) {
        const row = data[j].slice(0, 11);
        batchData.push(row);
      }

      const batchName = `Batch_${String(sheetIndex).padStart(3, '0')}_${batchData.length - 1}件`;
      const newSpreadsheet = SpreadsheetApp.create(batchName);
      const newSheet = newSpreadsheet.getActiveSheet();
      newSheet.setName('データ');

      newSheet.getRange(1, 1, batchData.length, batchData[0].length).setValues(batchData);

      // 顧客分類と性別をテキスト形式として設定
      if (batchData.length > 1) {
        // 顧客分類（B列）をテキスト形式に
        const customerTypeRange = newSheet.getRange(2, 2, batchData.length - 1, 1);
        customerTypeRange.setNumberFormat('@');

        // 性別（E列）もテキスト形式に
        const genderRange = newSheet.getRange(2, 5, batchData.length - 1, 1);
        genderRange.setNumberFormat('@');
      }

      const headerRange = newSheet.getRange(1, 1, 1, headers.length);
      headerRange.setBackground('#f3f3f3');
      headerRange.setFontWeight('bold');

      for (let col = 1; col <= headers.length; col++) {
        newSheet.autoResizeColumn(col);
      }

      const newFile = DriveApp.getFileById(newSpreadsheet.getId());
      newFile.moveTo(outputFolder);

      createdFiles.push({
        name: batchName,
        url: newSpreadsheet.getUrl(),
        count: batchData.length - 1
      });

      console.log(`${batchName}: ${batchData.length - 1}件のデータを出力`);
      totalProcessed += batchData.length - 1;
      sheetIndex++;
    }

    createSummarySheet(outputFolder, createdFiles, totalProcessed);

    updateDashboard('step3', '完了', `${sheetIndex - 1}ファイル作成`);

    const dashboard = targetSpreadsheet.getSheetByName('Dashboard');
    if (dashboard) {
      dashboard.getRange(13, 2).setValue(sheetIndex - 1);
    }

    addHistory('データ分割', sourceFile ? sourceFile.name : '', totalProcessed, 0, sheetIndex - 1, '成功', outputFolder.getName());

    const folderUrl = outputFolder.getUrl();
    ui.alert(
      'Step 3 完了',
      `データを${sheetIndex - 1}個のスプレッドシートに分割しました。\n\n` +
      `合計: ${totalProcessed}件\n` +
      `保存先フォルダ: ${folderName}\n\n` +
      `フォルダURL:\n${folderUrl}`,
      ui.ButtonSet.OK
    );

    return folderUrl;

  } catch (error) {
    updateDashboard('step3', 'エラー', error.toString());
    console.error('Step 3 エラー:', error);
    SpreadsheetApp.getUi().alert('エラー', 'データ分割中にエラーが発生しました: ' + error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * サマリーシートを作成
 */
function createSummarySheet(folder, createdFiles, totalCount) {
  try {
    const summarySpreadsheet = SpreadsheetApp.create('00_サマリー');
    const summarySheet = summarySpreadsheet.getActiveSheet();

    const headers = [
      ['反社リスト分割データ サマリー'],
      [''],
      ['作成日時', Utilities.formatDate(new Date(), 'JST', 'yyyy/MM/dd HH:mm:ss')],
      ['総件数', totalCount + '件'],
      ['ファイル数', createdFiles.length + '個'],
      [''],
      ['ファイル一覧'],
      ['No.', 'ファイル名', '件数', 'URL']
    ];

    let rowIndex = 1;
    headers.forEach(row => {
      if (row.length > 0) {
        summarySheet.getRange(rowIndex, 1, 1, row.length).setValues([row]);
      }
      rowIndex++;
    });

    createdFiles.forEach((file, index) => {
      const fileRow = [
        index + 1,
        file.name,
        file.count,
        file.url
      ];
      summarySheet.getRange(rowIndex, 1, 1, fileRow.length).setValues([fileRow]);
      rowIndex++;
    });

    summarySheet.getRange(1, 1, 1, 4).merge()
      .setFontSize(16)
      .setFontWeight('bold')
      .setBackground('#1a73e8')
      .setFontColor('#ffffff')
      .setHorizontalAlignment('center');

    summarySheet.getRange(7, 1, 1, 4).merge()
      .setFontSize(12)
      .setFontWeight('bold')
      .setBackground('#e3f2fd');

    summarySheet.getRange(8, 1, 1, 4)
      .setBackground('#f3f3f3')
      .setFontWeight('bold');

    summarySheet.setColumnWidth(1, 60);
    summarySheet.setColumnWidth(2, 250);
    summarySheet.setColumnWidth(3, 100);
    summarySheet.setColumnWidth(4, 400);

    const summaryFile = DriveApp.getFileById(summarySpreadsheet.getId());
    summaryFile.moveTo(folder);

    console.log('サマリーシートを作成しました');

  } catch (error) {
    console.error('サマリーシート作成エラー:', error);
  }
}

// ========================================
// 一括処理
// ========================================

/**
 * すべてのステップを連続実行
 */
function executeAllSteps() {
  const ui = SpreadsheetApp.getUi();

  const sourceFile = getCurrentSourceFile();
  if (!sourceFile) {
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

  const response = ui.alert(
    '確認',
    `以下のファイルを処理します:\n${sourceFile.name}\n\nすべてのステップを順番に実行します。\n続行しますか？`,
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    return;
  }

  try {
    step1_loadAndTransform();

    const targetSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const masterSheet = targetSpreadsheet.getSheetByName('Master_Data');

    if (masterSheet) {
      const data = masterSheet.getDataRange().getValues();

      let needsAI = false;
      for (let i = 1; i < data.length; i++) {
        if (data[i][2] && !data[i][11]) {
          needsAI = true;
          break;
        }
      }

      if (needsAI) {
        const aiResponse = ui.alert(
          '確認',
          'フリガナがないデータが見つかりました。AI予測を実行しますか？',
          ui.ButtonSet.YES_NO
        );

        if (aiResponse === ui.Button.YES) {
          step2_predictWithAI();
        }
      }

      step3_splitToSpreadsheets();
    }

    addHistory('一括処理', sourceFile.name, 0, 0, 0, '完了', '全ステップ実行');

    ui.alert('完了', 'すべての処理が完了しました。', ui.ButtonSet.OK);

  } catch (error) {
    console.error('一括処理エラー:', error);
    ui.alert('エラー', '処理中にエラーが発生しました: ' + error.toString(), ui.ButtonSet.OK);
  }
}

// ========================================
// ユーティリティ関数
// ========================================

/**
 * OpenAI APIを使用して漢字名からフリガナを予測（バッチ処理）
 */
function batchPredictKanaNames(kanjiNames) {
  const predictions = [];
  const batchSize = 10;
  const apiKey = getApiKey();

  if (!apiKey) {
    throw new Error('OpenAI APIキーが設定されていません。メニューから設定してください。');
  }

  for (let i = 0; i < kanjiNames.length; i += batchSize) {
    const batch = kanjiNames.slice(i, Math.min(i + batchSize, kanjiNames.length));

    try {
      const prompt = `以下の日本人の名前（漢字）について、それぞれのカタカナ読みを推測してください。
姓と名の間にはスペースを入れてください。
出力形式：漢字名|カタカナ読み

名前リスト：
${batch.join('\n')}`;

      const response = callOpenAI(prompt, apiKey);

      const lines = response.split('\n');
      for (const line of lines) {
        if (line.includes('|')) {
          const parts = line.split('|');
          if (parts.length >= 2) {
            const kana = parts[1].trim();
            predictions.push(kana);
          }
        }
      }

      while (predictions.length < Math.min(i + batchSize, kanjiNames.length)) {
        predictions.push('');
      }

    } catch (error) {
      console.error(`バッチ ${i / batchSize + 1} の処理中にエラー:`, error);
      for (let j = 0; j < batch.length; j++) {
        predictions.push('');
      }
    }

    if (i + batchSize < kanjiNames.length) {
      Utilities.sleep(1000);
    }
  }

  return predictions;
}

/**
 * OpenAI APIを呼び出す
 */
function callOpenAI(prompt, apiKey) {
  const configSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config');
  let model = 'gpt-4o-mini';
  let timeout = 30000;

  if (configSheet) {
    const configData = configSheet.getDataRange().getValues();
    for (let i = 1; i < configData.length; i++) {
      if (configData[i][0] === 'API モデル') {
        model = configData[i][1] || 'gpt-4o-mini';
      } else if (configData[i][0] === 'API タイムアウト') {
        timeout = parseInt(configData[i][1]) || 30000;
      }
    }
  }

  const payload = {
    model: model,
    messages: [
      {
        role: 'system',
        content: '日本人の名前の読み方を推測する専門家として回答してください。'
      },
      {
        role: 'user',
        content: prompt
      }
    ],
    temperature: 0.3,
    max_tokens: 500
  };

  const options = {
    method: 'post',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': `Bearer ${apiKey}`
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(OPENAI_API_URL, options);
    const json = JSON.parse(response.getContentText());

    if (json.error) {
      throw new Error(json.error.message);
    }

    return json.choices[0].message.content;
  } catch (error) {
    console.error('OpenAI API エラー:', error);
    return '';
  }
}

/**
 * カタカナをローマ字に変換（半角カタカナ対応）
 */
function convertKanaToRomaji(kana) {
  if (!kana) return '';

  kana = kana.replace(' (AI予測)', '');
  kana = kana.replace(/　/g, ' ');

  // 半角カタカナを全角カタカナに変換
  const hankakuToZenkaku = {
    'ｱ': 'ア', 'ｲ': 'イ', 'ｳ': 'ウ', 'ｴ': 'エ', 'ｵ': 'オ',
    'ｶ': 'カ', 'ｷ': 'キ', 'ｸ': 'ク', 'ｹ': 'ケ', 'ｺ': 'コ',
    'ｻ': 'サ', 'ｼ': 'シ', 'ｽ': 'ス', 'ｾ': 'セ', 'ｿ': 'ソ',
    'ﾀ': 'タ', 'ﾁ': 'チ', 'ﾂ': 'ツ', 'ﾃ': 'テ', 'ﾄ': 'ト',
    'ﾅ': 'ナ', 'ﾆ': 'ニ', 'ﾇ': 'ヌ', 'ﾈ': 'ネ', 'ﾉ': 'ノ',
    'ﾊ': 'ハ', 'ﾋ': 'ヒ', 'ﾌ': 'フ', 'ﾍ': 'ヘ', 'ﾎ': 'ホ',
    'ﾏ': 'マ', 'ﾐ': 'ミ', 'ﾑ': 'ム', 'ﾒ': 'メ', 'ﾓ': 'モ',
    'ﾔ': 'ヤ', 'ﾕ': 'ユ', 'ﾖ': 'ヨ',
    'ﾗ': 'ラ', 'ﾘ': 'リ', 'ﾙ': 'ル', 'ﾚ': 'レ', 'ﾛ': 'ロ',
    'ﾜ': 'ワ', 'ｦ': 'ヲ', 'ﾝ': 'ン',
    'ｧ': 'ァ', 'ｨ': 'ィ', 'ｩ': 'ゥ', 'ｪ': 'ェ', 'ｫ': 'ォ',
    'ｬ': 'ャ', 'ｭ': 'ュ', 'ｮ': 'ョ', 'ｯ': 'ッ',
    'ﾞ': '゛', 'ﾟ': '゜', 'ｰ': 'ー', '･': '・'
  };

  // 半角カタカナを全角カタカナに変換
  let zenkaku = kana;
  for (const [hankaku, zen] of Object.entries(hankakuToZenkaku)) {
    zenkaku = zenkaku.replace(new RegExp(hankaku, 'g'), zen);
  }

  // 濁音・半濁音の処理
  zenkaku = zenkaku.replace(/([カキクケコサシスセソタチツテトハヒフヘホ])゛/g, function(match, p1) {
    const dakuten = {
      'カ': 'ガ', 'キ': 'ギ', 'ク': 'グ', 'ケ': 'ゲ', 'コ': 'ゴ',
      'サ': 'ザ', 'シ': 'ジ', 'ス': 'ズ', 'セ': 'ゼ', 'ソ': 'ゾ',
      'タ': 'ダ', 'チ': 'ヂ', 'ツ': 'ヅ', 'テ': 'デ', 'ト': 'ド',
      'ハ': 'バ', 'ヒ': 'ビ', 'フ': 'ブ', 'ヘ': 'ベ', 'ホ': 'ボ'
    };
    return dakuten[p1] || match;
  });

  zenkaku = zenkaku.replace(/([ハヒフヘホ])゜/g, function(match, p1) {
    const handakuten = {
      'ハ': 'パ', 'ヒ': 'ピ', 'フ': 'プ', 'ヘ': 'ペ', 'ホ': 'ポ'
    };
    return handakuten[p1] || match;
  });

  // カタカナをひらがなに変換
  const hiragana = zenkaku.replace(/[\u30a1-\u30f6]/g, function(match) {
    const chr = match.charCodeAt(0) - 0x60;
    return String.fromCharCode(chr);
  });

  const conversionTable = {
    'あ': 'a', 'い': 'i', 'う': 'u', 'え': 'e', 'お': 'o',
    'か': 'ka', 'き': 'ki', 'く': 'ku', 'け': 'ke', 'こ': 'ko',
    'が': 'ga', 'ぎ': 'gi', 'ぐ': 'gu', 'げ': 'ge', 'ご': 'go',
    'さ': 'sa', 'し': 'shi', 'す': 'su', 'せ': 'se', 'そ': 'so',
    'ざ': 'za', 'じ': 'ji', 'ず': 'zu', 'ぜ': 'ze', 'ぞ': 'zo',
    'た': 'ta', 'ち': 'chi', 'つ': 'tsu', 'て': 'te', 'と': 'to',
    'だ': 'da', 'ぢ': 'ji', 'づ': 'zu', 'で': 'de', 'ど': 'do',
    'な': 'na', 'に': 'ni', 'ぬ': 'nu', 'ね': 'ne', 'の': 'no',
    'は': 'ha', 'ひ': 'hi', 'ふ': 'fu', 'へ': 'he', 'ほ': 'ho',
    'ば': 'ba', 'び': 'bi', 'ぶ': 'bu', 'べ': 'be', 'ぼ': 'bo',
    'ぱ': 'pa', 'ぴ': 'pi', 'ぷ': 'pu', 'ぺ': 'pe', 'ぽ': 'po',
    'ま': 'ma', 'み': 'mi', 'む': 'mu', 'め': 'me', 'も': 'mo',
    'や': 'ya', 'ゆ': 'yu', 'よ': 'yo',
    'ら': 'ra', 'り': 'ri', 'る': 'ru', 'れ': 're', 'ろ': 'ro',
    'わ': 'wa', 'ゐ': 'wi', 'ゑ': 'we', 'を': 'wo', 'ん': 'n',
    'ゃ': 'ya', 'ゅ': 'yu', 'ょ': 'yo',
    'ぁ': 'a', 'ぃ': 'i', 'ぅ': 'u', 'ぇ': 'e', 'ぉ': 'o',
    'っ': '', 'ー': ''
  };

  let romaji = '';
  let i = 0;

  while (i < hiragana.length) {
    if (i < hiragana.length - 1) {
      const char = hiragana[i];
      const nextChar = hiragana[i + 1];

      if (nextChar === 'ゃ' || nextChar === 'ゅ' || nextChar === 'ょ') {
        const baseRomaji = conversionTable[char] || char;
        const yoon = nextChar === 'ゃ' ? 'a' : nextChar === 'ゅ' ? 'u' : 'o';

        if (baseRomaji.length > 1) {
          romaji += baseRomaji.slice(0, -1) + 'y' + yoon;
        } else {
          romaji += baseRomaji + 'y' + yoon;
        }
        i += 2;
        continue;
      }

      if (char === 'っ' && nextChar) {
        const nextRomaji = conversionTable[nextChar] || nextChar;
        if (nextRomaji && nextRomaji[0]) {
          romaji += nextRomaji[0];
        }
        i++;
        continue;
      }
    }

    const char = hiragana[i];
    if (char === ' ') {
      romaji += ' ';
    } else {
      romaji += conversionTable[char] || char;
    }
    i++;
  }

  return romaji.toLowerCase();
}

// ========================================
// 設定管理
// ========================================

/**
 * APIキーを設定
 */
function setApiKey() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('OpenAI APIキー設定', 'OpenAI APIキーを入力してください:', ui.ButtonSet.OK_CANCEL);

  if (response.getSelectedButton() == ui.Button.OK) {
    const apiKey = response.getResponseText();
    PropertiesService.getScriptProperties().setProperty('OPENAI_API_KEY', apiKey);

    const configSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config');
    if (configSheet) {
      const configData = configSheet.getDataRange().getValues();
      for (let i = 1; i < configData.length; i++) {
        if (configData[i][0] === 'OpenAI APIキー') {
          configSheet.getRange(i + 1, 2).setValue('設定済み');
          break;
        }
      }
    }

    const dashboard = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Dashboard');
    if (dashboard) {
      dashboard.getRange(16, 2).setValue('設定済み');
    }

    ui.alert('APIキーを設定しました');
  }
}

/**
 * 保存されたAPIキーを取得
 */
function getApiKey() {
  const savedKey = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
  return savedKey;
}

// ========================================
// メニュー設定
// ========================================

/**
 * スプレッドシート開いた時にメニューを追加
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu('🔧 反社リスト処理')
    .addSubMenu(ui.createMenu('🏗️ セットアップ')
      .addItem('📊 現在のシートをセットアップ', 'initialSetup')
      .addItem('📝 HTMLファイル作成ガイド', 'showHTMLCreationGuide')
      .addSeparator()
      .addItem('🔑 OpenAI APIキー設定', 'setApiKey'))
    .addSeparator()
    .addSubMenu(ui.createMenu('📁 データ選択')
      .addItem('📂 ファイル一覧から選択', 'selectSourceFile')
      .addItem('🔗 URLで指定', 'setSourceFileByUrl')
      .addItem('⏰ 最近使用したファイル', 'selectFromRecentFiles')
      .addItem('📄 現在の選択を確認', 'showCurrentSelection'))
    .addSeparator()
    .addSubMenu(ui.createMenu('▶️ 個別実行')
      .addItem('Step 1: データ読み込み・基本変換', 'step1_loadAndTransform')
      .addItem('Step 2: AI予測（フリガナ）', 'step2_predictWithAI')
      .addItem('Step 3: 350件ごとに個別スプレッドシート作成', 'step3_splitToSpreadsheets'))
    .addSeparator()
    .addItem('🚀 すべて実行（一括処理）', 'executeAllSteps')
    .addSeparator()
    .addItem('📖 使い方', 'showInstructions')
    .addItem('ℹ️ バージョン情報', 'showVersion')
    .addToUi();
}

/**
 * 使い方を表示
 */
function showInstructions() {
  const instructions = `
【反社リスト変換システム 使い方】

◆ 初回セットアップ
1. 「セットアップ」→「現在のシートをセットアップ」
   - Dashboard, Config, Historyシートを自動作成
2. 「セットアップ」→「OpenAI APIキー設定」
   - OpenAI APIキーを入力

◆ データ選択（必須）
1. 「データ選択」メニューから選択方法を選ぶ
   - ファイル一覧: Driveから選択
   - URL指定: スプレッドシートURLを入力
   - 最近使用: 過去10件から選択

◆ 個別実行（ステップごと）
Step 1: データ読み込み・基本変換
   - 選択したファイルから自動でカラムを検出
   - Master_Dataシートに保存

Step 2: AI予測（必要な場合のみ）
   - フリガナがない名前をAIで予測
   - コスト見積もり表示あり

Step 3: 350件ごとに個別スプレッドシート作成
   - Master_Dataを350件ずつに分割
   - 個別のスプレッドシートファイルとして保存
   - 専用フォルダに整理して格納
   - サマリーシート自動作成

◆ 一括実行
「すべて実行」で全ステップを自動実行

◆ Dashboard
- 処理状況をリアルタイムで確認
- 統計情報の自動更新
- エラー状態の可視化

◆ History
- 全処理履歴を自動記録
- フィルター機能で検索可能

◆ Config
- システム設定のカスタマイズ
- バッチサイズ、APIモデル等を調整可能
  `;

  SpreadsheetApp.getUi().alert('使い方', instructions, SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * バージョン情報を表示
 */
function showVersion() {
  const info = `
反社リスト変換システム - 完全統合版
Version: 4.0.0

主な機能:
• 自動セットアップ機能
• Dashboard/Config/History管理
• 動的ファイル選択
• カラム自動検出
• OpenAI APIフリガナ予測
• 350件自動分割
• 個別スプレッドシート出力
• 専用フォルダ管理
• サマリーシート自動生成

更新履歴:
v4.0.0 - 完全統合版
v3.1.0 - 自動セットアップ追加
v3.0.0 - 動的ファイル選択対応
v2.0.0 - モジュール化
v1.0.0 - 初版リリース

開発: 2025年
  `;

  SpreadsheetApp.getUi().alert('バージョン情報', info, SpreadsheetApp.getUi().ButtonSet.OK);
}