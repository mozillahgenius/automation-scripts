/**
 * 反社リスト変換処理（動的ファイル選択版）
 * 各処理を個別に実行可能
 * OpenAI API対応
 * 毎回異なるデータファイルを選択可能
 */

// ========================================
// 設定項目
// ========================================

// OpenAI API設定（他のファイルと競合を避けるため、変数名を変更）
const DYNAMIC_OPENAI_API_KEY = 'your-openai-api-key-here'; // TODO: 実際のAPIキーに置き換え
const DYNAMIC_OPENAI_API_URL = 'https://api.openai.com/v1/chat/completions';

// デフォルト設定
const DYNAMIC_CONFIG = {
  batchSize: 350, // 分割サイズ
  currentSourceFile: null // 現在選択中のソースファイル
};

// ========================================
// ファイル選択機能
// ========================================

/**
 * 元データファイルを選択
 */
function selectSourceFile() {
  const ui = SpreadsheetApp.getUi();

  // HTMLダイアログを作成
  const html = HtmlService.createHtmlOutputFromFile('file-picker')
    .setWidth(600)
    .setHeight(400);

  ui.showModalDialog(html, '元データファイルを選択');
}

/**
 * URLから元データファイルを設定
 */
function setSourceFileByUrl() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    '元データファイルの指定',
    'Google スプレッドシートのURLを入力してください:',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() === ui.Button.OK) {
    const url = response.getResponseText();

    try {
      // URLからIDを抽出
      const idMatch = url.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
      if (!idMatch) {
        throw new Error('有効なスプレッドシートURLではありません');
      }

      const fileId = idMatch[1];
      const file = DriveApp.getFileById(fileId);

      // ファイル情報を保存
      saveSelectedFile({
        id: fileId,
        name: file.getName(),
        url: url
      });

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
  const html = HtmlService.createHtmlOutputFromFile('recent-files')
    .setWidth(500)
    .setHeight(400);

  ui.showModalDialog(html, '最近使用したファイル');
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
// Step 1: データ読み込みと基本変換
// ========================================

/**
 * Step 1: 選択されたデータを読み込んで基本的な変換を実行
 */
function step1_loadAndTransform() {
  try {
    const ui = SpreadsheetApp.getUi();

    // 現在選択中のファイルを確認
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

    // ファイル情報を表示して確認
    const confirmResponse = ui.alert(
      '確認',
      `以下のファイルからデータを読み込みます:\n\n${sourceFile.name}\n\n続行しますか？`,
      ui.ButtonSet.YES_NO
    );

    if (confirmResponse !== ui.Button.YES) {
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
        ui.alert('処理をキャンセルしました。');
        return;
      }
    }

    // ソースファイルを開く
    let sourceSpreadsheet;
    let sourceSheet;

    try {
      sourceSpreadsheet = SpreadsheetApp.openById(sourceFile.id);

      // シート名を検索（優先順位付き）
      const sheetNames = ['基本データ', 'Sheet1', 'データ', 'Data'];
      sourceSheet = null;

      for (const sheetName of sheetNames) {
        try {
          sourceSheet = sourceSpreadsheet.getSheetByName(sheetName);
          if (sourceSheet) break;
        } catch (e) {
          // シートが見つからない場合は次を試す
        }
      }

      // それでも見つからない場合は最初のシートを使用
      if (!sourceSheet) {
        sourceSheet = sourceSpreadsheet.getSheets()[0];
        ui.alert(
          '注意',
          `「基本データ」シートが見つからないため、「${sourceSheet.getName()}」シートを使用します。`,
          ui.ButtonSet.OK
        );
      }

    } catch (error) {
      ui.alert('エラー', `ファイルを開けませんでした:\n${error.toString()}`, ui.ButtonSet.OK);
      return;
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
      'フリガナ元データ', // 元のフリガナを保存
      'ソースファイル' // ソースファイル名を記録
    ];
    processedData.push(targetHeaders);

    // カラムインデックスを自動検出
    const headers = sourceData[0];
    const columnMapping = detectColumns(headers);

    if (!columnMapping.name) {
      ui.alert('エラー', '名前の列が見つかりません。ヘッダー行を確認してください。', ui.ButtonSet.OK);
      return;
    }

    // データを処理（2行目から開始）
    let processedCount = 0;
    for (let i = 1; i < sourceData.length; i++) {
      const row = sourceData[i];

      // 空行はスキップ
      if (!row[columnMapping.name]) continue;

      const processedRow = [];

      // 1. 登録日付
      processedRow.push(today);

      // 2. 顧客区分
      const orgName = columnMapping.org !== -1 ? row[columnMapping.org] || '' : '';
      const isOrganization = orgName.includes('組') || orgName.includes('会') || orgName.includes('団');
      processedRow.push(isOrganization ? '02' : '01');

      // 3. 日本語名
      processedRow.push(row[columnMapping.name] || '');

      // 4. 英文名（この段階では空欄またはフリガナをそのまま）
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

      // 異名
      if (columnMapping.alias !== -1 && row[columnMapping.alias]) {
        remarks.push('異名: ' + row[columnMapping.alias]);
      }

      // 年齢（既に生年月日に変換済みだが、元の年齢も記録）
      if (age) {
        remarks.push('年齢: ' + age);
      }

      // 組織・団体
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

      // 12. フリガナ元データ
      processedRow.push(kanaName);

      // 13. ソースファイル
      processedRow.push(sourceFile.name);

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

    ui.alert('Step 1 完了', `ファイル: ${sourceFile.name}\n${processedCount}件のデータを読み込み、基本変換を完了しました。`, ui.ButtonSet.OK);

    // フリガナがない件数を確認
    let missingKanaCount = 0;
    for (let i = 1; i < processedData.length; i++) {
      if (!processedData[i][11] && processedData[i][2]) { // フリガナがなく、名前がある場合
        missingKanaCount++;
      }
    }

    if (missingKanaCount > 0) {
      ui.alert('情報', `${missingKanaCount}件のデータにフリガナがありません。\nStep 2でAI予測を実行してください。`, ui.ButtonSet.OK);
    }

  } catch (error) {
    console.error('Step 1 エラー:', error);
    SpreadsheetApp.getUi().alert('エラー', 'データの読み込み中にエラーが発生しました: ' + error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
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

    // 名前
    if (header.includes('名前') || header.includes('氏名') || header === 'name') {
      mapping.name = i;
    }
    // フリガナ
    else if (header.includes('ふりがな') || header.includes('フリガナ') || header.includes('カナ') || header.includes('ﾌﾘｶﾞﾅ') || header.includes('kana')) {
      mapping.kana = i;
    }
    // 性別
    else if (header.includes('性別') || header === 'gender' || header === 'sex') {
      mapping.gender = i;
    }
    // 年齢
    else if (header.includes('年齢') || header === 'age') {
      mapping.age = i;
    }
    // 組織
    else if (header.includes('組織') || header.includes('団体') || header.includes('所属')) {
      mapping.org = i;
    }
    // 住所
    else if (header.includes('住所') || header.includes('居住') || header === 'address') {
      mapping.address = i;
    }
    // 異名
    else if (header.includes('異名') || header.includes('別名') || header === 'alias') {
      mapping.alias = i;
    }
    // 内容
    else if (header.includes('内容') || (header.includes('備考') && !header.includes('年月日')) || header.includes('詳細')) {
      mapping.content = i;
    }
    // 資料作成年月日
    else if (header.includes('資料作成年月日')) {
      mapping.createDate = i;
    }
    // 県名
    else if (header.includes('県名')) {
      mapping.prefecture = i;
    }
    // 更新年月日
    else if (header.includes('更新年月日')) {
      mapping.updateDate = i;
    }
    // 新聞社名
    else if (header.includes('新聞社名')) {
      mapping.newspaper = i;
    }
    // 刊行区分
    else if (header.includes('刊行区分')) {
      mapping.publicationType = i;
    }
    // 区分
    else if (header === '区分' || (header.includes('区分') && !header.includes('刊行'))) {
      mapping.category = i;
    }
    // 番号
    else if (header === '番号' || header.includes('番号')) {
      mapping.number = i;
    }
  }

  return mapping;
}

// ========================================
// Step 2-4: AI予測、分割、CSV出力
// （前のコードと同じ）
// ========================================

/**
 * Step 2: フリガナがないデータをAIで予測
 */
function step2_predictWithAI() {
  try {
    const ui = SpreadsheetApp.getUi();
    const targetSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const masterSheet = targetSpreadsheet.getSheetByName('Master_Data');

    if (!masterSheet) {
      ui.alert('エラー', 'Master_Dataシートが見つかりません。Step 1を先に実行してください。', ui.ButtonSet.OK);
      return;
    }

    const data = masterSheet.getDataRange().getValues();
    const namesToProcess = [];
    const rowIndices = [];

    // フリガナがない行を特定
    for (let i = 1; i < data.length; i++) {
      const kanjiName = data[i][2]; // 日本語名
      const kanaData = data[i][11]; // フリガナ元データ

      if (kanjiName && !kanaData) {
        namesToProcess.push(kanjiName);
        rowIndices.push(i);
      }
    }

    if (namesToProcess.length === 0) {
      ui.alert('情報', '予測が必要なデータはありません。', ui.ButtonSet.OK);
      return;
    }

    const response = ui.alert(
      '確認',
      `${namesToProcess.length}件の名前でフリガナを予測します。\n推定コスト: $${(namesToProcess.length * 0.00005).toFixed(2)}\n続行しますか？`,
      ui.ButtonSet.YES_NO
    );

    if (response !== ui.Button.YES) {
      ui.alert('処理をキャンセルしました。');
      return;
    }

    // AI予測を実行
    console.log(`${namesToProcess.length}件の名前でフリガナを予測開始`);
    const predictions = batchPredictKanaNames(namesToProcess);

    // 予測結果を反映
    for (let j = 0; j < rowIndices.length; j++) {
      const rowIndex = rowIndices[j];
      const predictedKana = predictions[j];

      if (predictedKana) {
        // 英文名を更新
        const romaji = convertKanaToRomaji(predictedKana);
        masterSheet.getRange(rowIndex + 1, 4).setValue(romaji);

        // フリガナ元データを更新
        masterSheet.getRange(rowIndex + 1, 12).setValue(predictedKana + ' (AI予測)');

        // 備考に追記
        const currentRemarks = data[rowIndex][10];
        const updatedRemarks = currentRemarks ?
          currentRemarks + ' / ※フリガナはAI予測' :
          '※フリガナはAI予測';
        masterSheet.getRange(rowIndex + 1, 11).setValue(updatedRemarks);
      }
    }

    ui.alert('Step 2 完了', `${namesToProcess.length}件のフリガナ予測が完了しました。`, ui.ButtonSet.OK);

  } catch (error) {
    console.error('Step 2 エラー:', error);
    SpreadsheetApp.getUi().alert('エラー', 'AI予測中にエラーが発生しました: ' + error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Step 3: データを350件ずつに分割して個別スプレッドシートとして保存
 */
function step3_splitToSpreadsheets() {
  try {
    const ui = SpreadsheetApp.getUi();
    const targetSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const masterSheet = targetSpreadsheet.getSheetByName('Master_Data');

    if (!masterSheet) {
      ui.alert('エラー', 'Master_Dataシートが見つかりません。Step 1を先に実行してください。', ui.ButtonSet.OK);
      return;
    }

    const data = masterSheet.getDataRange().getValues();
    const headers = data[0].slice(0, 11); // フリガナ元データとソースファイル列は除外

    // 現在のスプレッドシートが格納されているフォルダを取得
    const currentFile = DriveApp.getFileById(targetSpreadsheet.getId());
    const parentFolder = currentFile.getParents().next();

    // 出力用フォルダの作成または取得
    const timestamp = Utilities.formatDate(new Date(), 'JST', 'yyyyMMdd_HHmmss');
    const sourceFile = getCurrentSourceFile();
    const folderName = sourceFile ?
      `反社リスト_${sourceFile.name.replace(/\.[^/.]+$/, '')}_${timestamp}` :
      `反社リスト_分割データ_${timestamp}`;

    // 既存の同名フォルダをチェック
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
        // 既存のファイルを削除
        const existingFiles = outputFolder.getFiles();
        while (existingFiles.hasNext()) {
          existingFiles.next().setTrashed(true);
        }
      } else {
        ui.alert('処理をキャンセルしました。');
        return;
      }
    } else {
      outputFolder = parentFolder.createFolder(folderName);
    }

    // データを分割して個別スプレッドシートを作成
    const batchSize = DYNAMIC_CONFIG.batchSize;
    let sheetIndex = 1;
    let totalProcessed = 0;
    const createdFiles = [];

    for (let i = 1; i < data.length; i += batchSize) {
      const batchData = [headers];
      const endIndex = Math.min(i + batchSize, data.length);

      for (let j = i; j < endIndex; j++) {
        // フリガナ元データとソースファイル列を除外してコピー
        const row = data[j].slice(0, 11);
        batchData.push(row);
      }

      // 新しいスプレッドシートを作成
      const batchName = `Batch_${String(sheetIndex).padStart(3, '0')}_${batchData.length - 1}件`;
      const newSpreadsheet = SpreadsheetApp.create(batchName);
      const newSheet = newSpreadsheet.getActiveSheet();
      newSheet.setName('データ');

      // データを書き込み
      newSheet.getRange(1, 1, batchData.length, batchData[0].length).setValues(batchData);

      // 顧客区分列を文字列形式に設定
      if (batchData.length > 1) {
        const customerTypeRange = newSheet.getRange(2, 2, batchData.length - 1, 1);
        customerTypeRange.setNumberFormat('@');
      }

      // ヘッダー行のスタイリング
      const headerRange = newSheet.getRange(1, 1, 1, headers.length);
      headerRange.setBackground('#f3f3f3');
      headerRange.setFontWeight('bold');

      // 列幅の自動調整
      for (let col = 1; col <= headers.length; col++) {
        newSheet.autoResizeColumn(col);
      }

      // スプレッドシートを指定フォルダに移動
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

    // サマリーシートを作成
    createSummarySheet(outputFolder, createdFiles, totalProcessed);

    // 完了メッセージ
    const folderUrl = outputFolder.getUrl();
    ui.alert(
      'Step 3 完了',
      `データを${sheetIndex - 1}個のスプレッドシートに分割しました。\n\n` +
      `合計: ${totalProcessed}件\n` +
      `保存先フォルダ: ${folderName}\n\n` +
      `フォルダを開くには「OK」をクリック後、以下のURLにアクセスしてください:\n${folderUrl}`,
      ui.ButtonSet.OK
    );

    // フォルダURLをクリップボードにコピー（可能な場合）
    return folderUrl;

  } catch (error) {
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

    // ヘッダー
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

    // データを設定
    let rowIndex = 1;
    headers.forEach(row => {
      if (row.length > 0) {
        summarySheet.getRange(rowIndex, 1, 1, row.length).setValues([row]);
      }
      rowIndex++;
    });

    // ファイルリストを追加
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

    // スタイリング
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

    // 列幅調整
    summarySheet.setColumnWidth(1, 60);
    summarySheet.setColumnWidth(2, 250);
    summarySheet.setColumnWidth(3, 100);
    summarySheet.setColumnWidth(4, 400);

    // サマリーシートをフォルダに移動
    const summaryFile = DriveApp.getFileById(summarySpreadsheet.getId());
    summaryFile.moveTo(folder);

    console.log('サマリーシートを作成しました');

  } catch (error) {
    console.error('サマリーシート作成エラー:', error);
  }
}

// Step 4: CSV出力機能は削除されました
// データは Step 3 で個別スプレッドシートとして出力されます

// ========================================
// ユーティリティ関数（前のコードと同じ）
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

      // レスポンスを解析
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

      // レスポンスが不足している場合の処理
      while (predictions.length < Math.min(i + batchSize, kanjiNames.length)) {
        predictions.push('');
      }

    } catch (error) {
      console.error(`バッチ ${i / batchSize + 1} の処理中にエラー:`, error);
      for (let j = 0; j < batch.length; j++) {
        predictions.push('');
      }
    }

    // API制限を考慮して待機
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
  const payload = {
    model: 'gpt-4o-mini',
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
    const response = UrlFetchApp.fetch(DYNAMIC_OPENAI_API_URL, options);
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
 * カタカナをローマ字に変換
 */
function convertKanaToRomaji(kana) {
  if (!kana) return '';

  // (AI予測)の文字を削除
  kana = kana.replace(' (AI予測)', '');

  // 全角スペースを半角スペースに変換
  kana = kana.replace(/　/g, ' ');

  // カタカナをひらがなに変換
  const hiragana = kana.replace(/[\u30a1-\u30f6]/g, function(match) {
    const chr = match.charCodeAt(0) - 0x60;
    return String.fromCharCode(chr);
  });

  // ひらがな→ローマ字変換テーブル
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
    ui.alert('APIキーを設定しました');
  }
}

/**
 * 保存されたAPIキーを取得
 */
function getApiKey() {
  const savedKey = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
  return savedKey || DYNAMIC_OPENAI_API_KEY;
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
    .addSubMenu(ui.createMenu('📁 データ選択')
      .addItem('📂 ファイルピッカーで選択', 'selectSourceFile')
      .addItem('🔗 URLで指定', 'setSourceFileByUrl')
      .addItem('⏰ 最近使用したファイル', 'selectFromRecentFiles')
      .addItem('📄 現在の選択を確認', 'showCurrentSelection'))
    .addSeparator()
    .addSubMenu(ui.createMenu('⚙️ 初期設定')
      .addItem('🔑 OpenAI APIキー設定', 'setApiKey'))
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

/**
 * 一括実行
 */
function executeAllSteps() {
  const ui = SpreadsheetApp.getUi();

  // ファイルが選択されているか確認
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
    ui.alert('処理をキャンセルしました。');
    return;
  }

  try {
    // Step 1
    step1_loadAndTransform();

    // フリガナがないデータがあるか確認
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

      // Step 2 (必要な場合)
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

      // Step 3
      step3_splitToSpreadsheets();
    }

    ui.alert('完了', 'すべての処理が完了しました。', ui.ButtonSet.OK);

  } catch (error) {
    console.error('一括処理エラー:', error);
    ui.alert('エラー', '処理中にエラーが発生しました: ' + error.toString(), ui.ButtonSet.OK);
  }
}

/**
 * 使い方を表示
 */
function showInstructions() {
  const instructions = `
【反社リスト変換ツール 使い方】

◆ データ選択（必須）
1. 「データ選択」メニューから選択方法を選ぶ
   - ファイルピッカー: Driveから選択
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

◆ 特徴
- 毎回異なるファイルを選択可能
- カラム自動検出機能
- 最近使用したファイルの履歴管理
- ソースファイル名をCSVに含める
  `;

  SpreadsheetApp.getUi().alert('使い方', instructions, SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * バージョン情報を表示
 */
function showVersion() {
  const info = `
反社リスト変換ツール
Version: 3.0.0

主な機能:
• 動的ファイル選択
• カラム自動検出
• ファイル履歴管理
• OpenAI APIフリガナ予測
• 350件自動分割
• フォルダ選択CSV出力

更新履歴:
v3.0.0 - 動的ファイル選択対応
v2.0.0 - モジュール化
v1.0.0 - 初版リリース
  `;

  SpreadsheetApp.getUi().alert('バージョン情報', info, SpreadsheetApp.getUi().ButtonSet.OK);
}