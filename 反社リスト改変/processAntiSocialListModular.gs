/**
 * 反社リスト変換処理（モジュール版）
 * 各処理を個別に実行可能
 * OpenAI API対応
 */

// ========================================
// 設定項目
// ========================================

// OpenAI API設定
const OPENAI_API_KEY = 'your-openai-api-key-here'; // TODO: 実際のAPIキーに置き換え
const OPENAI_API_URL = 'https://api.openai.com/v1/chat/completions';

// スプレッドシートID設定
const CONFIG = {
  sourceSheetId: '元データのスプレッドシートID', // TODO: 実際のIDに置き換え
  targetSheetId: '出力先のスプレッドシートID', // TODO: 実際のIDに置き換え
  csvFolderId: 'Google DriveのフォルダID', // TODO: 実際のIDに置き換え
  batchSize: 350 // 分割サイズ
};

// ========================================
// Step 1: データ読み込みと基本変換
// ========================================

/**
 * Step 1: 元データを読み込んで基本的な変換を実行
 */
function step1_loadAndTransform() {
  try {
    const ui = SpreadsheetApp.getUi();

    // 既存の変換データがあるか確認
    const targetSpreadsheet = SpreadsheetApp.openById(CONFIG.targetSheetId);
    const masterSheet = targetSpreadsheet.getSheetByName('Master_Data');

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

    const sourceSheet = SpreadsheetApp.openById(CONFIG.sourceSheetId).getSheetByName('基本データ');
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
      'フリガナ元データ' // 追加：元のフリガナを保存
    ];
    processedData.push(targetHeaders);

    // データを処理（2行目から開始）
    let processedCount = 0;
    for (let i = 1; i < sourceData.length; i++) {
      const row = sourceData[i];

      // 空行はスキップ
      if (!row[2]) continue;

      const processedRow = [];

      // 1. 登録日付
      processedRow.push(today);

      // 2. 顧客区分
      const orgName = row[7] || '';
      const isOrganization = orgName.includes('組') || orgName.includes('会') || orgName.includes('団');
      processedRow.push(isOrganization ? '02' : '01');

      // 3. 日本語名
      processedRow.push(row[2] || '');

      // 4. 英文名（この段階では空欄またはフリガナをそのまま）
      const kanaName = row[4] || '';
      processedRow.push(kanaName ? convertKanaToRomaji(kanaName) : '');

      // 5. 性別
      const gender = row[6] || '';
      let genderCode = '';
      if (gender === '男') genderCode = '1';
      else if (gender === '女') genderCode = '2';
      processedRow.push(genderCode);

      // 6. 生年月日
      const age = row[5];
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
      processedRow.push(row[8] || '');

      // 11. 備考
      const remarks = [];
      if (row[3]) remarks.push('異名: ' + row[3]);
      if (row[7]) remarks.push('組織: ' + row[7]);
      if (row[9]) remarks.push(row[9]);
      processedRow.push(remarks.join(' / '));

      // 12. フリガナ元データ
      processedRow.push(kanaName);

      processedData.push(processedRow);
      processedCount++;
    }

    // Master_Dataシートに保存
    let masterSheet = targetSpreadsheet.getSheetByName('Master_Data');
    if (!masterSheet) {
      masterSheet = targetSpreadsheet.insertSheet('Master_Data');
    } else {
      masterSheet.clear();
    }

    masterSheet.getRange(1, 1, processedData.length, processedData[0].length).setValues(processedData);

    // 顧客区分列を文字列形式に設定
    const customerTypeRange = masterSheet.getRange(2, 2, processedData.length - 1, 1);
    customerTypeRange.setNumberFormat('@');

    ui.alert('Step 1 完了', `${processedCount}件のデータを読み込み、基本変換を完了しました。`, ui.ButtonSet.OK);

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

// ========================================
// Step 2: AI予測処理
// ========================================

/**
 * Step 2: フリガナがないデータをAIで予測
 */
function step2_predictWithAI() {
  try {
    const ui = SpreadsheetApp.getUi();
    const targetSpreadsheet = SpreadsheetApp.openById(CONFIG.targetSheetId);
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

    // プログレス表示
    const htmlOutput = HtmlService.createHtmlOutput('<p>AI予測処理中...</p>')
      .setWidth(300)
      .setHeight(100);
    ui.showModelessDialog(htmlOutput, '処理中');

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

// ========================================
// Step 3: データ分割処理
// ========================================

/**
 * Step 3: データを350件ずつに分割
 */
function step3_splitData() {
  try {
    const ui = SpreadsheetApp.getUi();
    const targetSpreadsheet = SpreadsheetApp.openById(CONFIG.targetSheetId);
    const masterSheet = targetSpreadsheet.getSheetByName('Master_Data');

    if (!masterSheet) {
      ui.alert('エラー', 'Master_Dataシートが見つかりません。Step 1を先に実行してください。', ui.ButtonSet.OK);
      return;
    }

    const data = masterSheet.getDataRange().getValues();
    const headers = data[0].slice(0, 11); // フリガナ元データ列は除外

    // 既存のBatchシートを確認
    const sheets = targetSpreadsheet.getSheets();
    const existingBatches = sheets.filter(sheet => sheet.getName().startsWith('Batch_'));

    if (existingBatches.length > 0) {
      const response = ui.alert(
        '確認',
        `${existingBatches.length}個の既存バッチシートが存在します。削除して新規作成しますか？`,
        ui.ButtonSet.YES_NO
      );

      if (response === ui.Button.YES) {
        existingBatches.forEach(sheet => targetSpreadsheet.deleteSheet(sheet));
      } else {
        ui.alert('処理をキャンセルしました。');
        return;
      }
    }

    // データを分割
    const batchSize = CONFIG.batchSize;
    let sheetIndex = 1;
    let totalProcessed = 0;

    for (let i = 1; i < data.length; i += batchSize) {
      const batchData = [headers];
      const endIndex = Math.min(i + batchSize, data.length);

      for (let j = i; j < endIndex; j++) {
        // フリガナ元データ列を除外してコピー
        const row = data[j].slice(0, 11);
        batchData.push(row);
      }

      // 新しいシートを作成
      const sheetName = `Batch_${sheetIndex}`;
      const newSheet = targetSpreadsheet.insertSheet(sheetName);

      // データを書き込み
      newSheet.getRange(1, 1, batchData.length, batchData[0].length).setValues(batchData);

      // 顧客区分列を文字列形式に設定
      const customerTypeRange = newSheet.getRange(2, 2, batchData.length - 1, 1);
      customerTypeRange.setNumberFormat('@');

      console.log(`${sheetName}: ${batchData.length - 1}件のデータを出力`);
      totalProcessed += batchData.length - 1;
      sheetIndex++;
    }

    ui.alert(
      'Step 3 完了',
      `データを${sheetIndex - 1}個のバッチに分割しました。\n合計: ${totalProcessed}件`,
      ui.ButtonSet.OK
    );

  } catch (error) {
    console.error('Step 3 エラー:', error);
    SpreadsheetApp.getUi().alert('エラー', 'データ分割中にエラーが発生しました: ' + error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

// ========================================
// Step 4: CSV出力処理
// ========================================

/**
 * Step 4: バッチシートをCSVファイルとして出力
 */
function step4_exportCSV() {
  try {
    const ui = SpreadsheetApp.getUi();
    const targetSpreadsheet = SpreadsheetApp.openById(CONFIG.targetSheetId);
    const sheets = targetSpreadsheet.getSheets();
    const batchSheets = sheets.filter(sheet => sheet.getName().startsWith('Batch_'));

    if (batchSheets.length === 0) {
      ui.alert('エラー', 'バッチシートが見つかりません。Step 3を先に実行してください。', ui.ButtonSet.OK);
      return;
    }

    const response = ui.alert(
      '確認',
      `${batchSheets.length}個のバッチシートをCSVファイルとして出力します。続行しますか？`,
      ui.ButtonSet.YES_NO
    );

    if (response !== ui.Button.YES) {
      ui.alert('処理をキャンセルしました。');
      return;
    }

    const folder = DriveApp.getFolderById(CONFIG.csvFolderId);
    let exportedCount = 0;

    batchSheets.forEach(sheet => {
      const data = sheet.getDataRange().getValues();

      // CSVコンテンツを作成
      let csvContent = '';
      data.forEach(row => {
        const csvRow = row.map(cell => {
          const cellStr = String(cell || '');
          if (cellStr.includes(',') || cellStr.includes('"') || cellStr.includes('\n')) {
            return '"' + cellStr.replace(/"/g, '""') + '"';
          }
          return cellStr;
        }).join(',');
        csvContent += csvRow + '\n';
      });

      // BOMを追加（Excel用）
      const bom = '\uFEFF';
      csvContent = bom + csvContent;

      // ファイル名に日時を追加
      const timestamp = Utilities.formatDate(new Date(), 'JST', 'yyyyMMdd_HHmmss');
      const fileName = `${sheet.getName()}_${timestamp}.csv`;

      // ファイルを作成
      const blob = Utilities.newBlob(csvContent, 'text/csv', fileName);
      folder.createFile(blob);

      console.log(`${fileName} を作成しました`);
      exportedCount++;
    });

    ui.alert(
      'Step 4 完了',
      `${exportedCount}個のCSVファイルを出力しました。\nフォルダ: ${folder.getName()}`,
      ui.ButtonSet.OK
    );

  } catch (error) {
    console.error('Step 4 エラー:', error);
    SpreadsheetApp.getUi().alert('エラー', 'CSV出力中にエラーが発生しました: ' + error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
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
  const response = ui.alert(
    '確認',
    'すべてのステップを順番に実行します。\n1. データ読み込み\n2. AI予測（必要な場合）\n3. データ分割\n4. CSV出力\n\n続行しますか？',
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
    const targetSpreadsheet = SpreadsheetApp.openById(CONFIG.targetSheetId);
    const masterSheet = targetSpreadsheet.getSheetByName('Master_Data');
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
    step3_splitData();

    // Step 4
    const csvResponse = ui.alert(
      '確認',
      'CSVファイルを出力しますか？',
      ui.ButtonSet.YES_NO
    );

    if (csvResponse === ui.Button.YES) {
      step4_exportCSV();
    }

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
  return savedKey || OPENAI_API_KEY;
}

/**
 * スプレッドシートIDを設定
 */
function configureSettings() {
  const ui = SpreadsheetApp.getUi();
  const html = HtmlService.createHtmlOutputFromFile('settings')
    .setWidth(500)
    .setHeight(400);
  ui.showModalDialog(html, '設定');
}

/**
 * 設定を保存
 */
function saveSettings(settings) {
  const props = PropertiesService.getScriptProperties();
  props.setProperty('SOURCE_SHEET_ID', settings.sourceSheetId);
  props.setProperty('TARGET_SHEET_ID', settings.targetSheetId);
  props.setProperty('CSV_FOLDER_ID', settings.csvFolderId);
  props.setProperty('BATCH_SIZE', settings.batchSize);

  // グローバル変数も更新
  CONFIG.sourceSheetId = settings.sourceSheetId;
  CONFIG.targetSheetId = settings.targetSheetId;
  CONFIG.csvFolderId = settings.csvFolderId;
  CONFIG.batchSize = parseInt(settings.batchSize);

  return '設定を保存しました';
}

/**
 * 設定を読み込み
 */
function loadSettings() {
  const props = PropertiesService.getScriptProperties();
  return {
    sourceSheetId: props.getProperty('SOURCE_SHEET_ID') || CONFIG.sourceSheetId,
    targetSheetId: props.getProperty('TARGET_SHEET_ID') || CONFIG.targetSheetId,
    csvFolderId: props.getProperty('CSV_FOLDER_ID') || CONFIG.csvFolderId,
    batchSize: props.getProperty('BATCH_SIZE') || CONFIG.batchSize
  };
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
    .addSubMenu(ui.createMenu('⚙️ 初期設定')
      .addItem('📋 スプレッドシートID設定', 'configureSettings')
      .addItem('🔑 OpenAI APIキー設定', 'setApiKey'))
    .addSeparator()
    .addSubMenu(ui.createMenu('▶️ 個別実行')
      .addItem('Step 1: データ読み込み・基本変換', 'step1_loadAndTransform')
      .addItem('Step 2: AI予測（フリガナ）', 'step2_predictWithAI')
      .addItem('Step 3: 350件ごとに分割', 'step3_splitData')
      .addItem('Step 4: CSV出力', 'step4_exportCSV'))
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
【反社リスト変換ツール 使い方】

◆ 初期設定（初回のみ）
1. 「初期設定」→「スプレッドシートID設定」
   - 元データ、出力先、CSVフォルダのIDを設定
2. 「初期設定」→「OpenAI APIキー設定」
   - OpenAI APIキーを入力

◆ 個別実行（ステップごと）
Step 1: データ読み込み・基本変換
   - 元データを読み込んで基本的な変換を実行
   - Master_Dataシートに保存

Step 2: AI予測（必要な場合のみ）
   - フリガナがない名前をAIで予測
   - コスト見積もり表示あり

Step 3: 350件ごとに分割
   - Master_DataをBatch_1, Batch_2...に分割
   - 既存バッチの上書き確認あり

Step 4: CSV出力
   - 各バッチをCSVファイルとして出力
   - タイムスタンプ付きファイル名

◆ 一括実行
「すべて実行」で全ステップを自動実行
   - AI予測の要否は自動判定
   - 各ステップで確認ダイアログ表示

◆ 注意事項
- 個人/法人区分は「01」「02」形式
- CSVはUTF-8 BOM付き（Excel対応）
- AI予測は10件ずつバッチ処理
  `;

  SpreadsheetApp.getUi().alert('使い方', instructions, SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * バージョン情報を表示
 */
function showVersion() {
  const info = `
反社リスト変換ツール
Version: 2.0.0

主な機能:
• ステップごとの個別実行
• OpenAI APIによるフリガナ予測
• 350件ごとの自動分割
• CSV一括出力

更新履歴:
v2.0.0 - モジュール化、個別実行対応
v1.0.0 - 初版リリース
  `;

  SpreadsheetApp.getUi().alert('バージョン情報', info, SpreadsheetApp.getUi().ButtonSet.OK);
}