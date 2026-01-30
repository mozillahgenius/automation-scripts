// Code.gsの続き

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
      ui.alert('処理をキャンセルしました。');
      return;
    }

    // AI予測を実行
    console.log(`${namesToProcess.length}件の名前でフリガナを予測開始`);
    const predictions = batchPredictKanaNames(namesToProcess);

    // 予測結果を反映
    let successCount = 0;
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

        successCount++;
      }
    }

    // Dashboard更新
    updateDashboard('step2', '完了', `${successCount}件予測完了`);

    // 統計情報更新
    const dashboard = targetSpreadsheet.getSheetByName('Dashboard');
    if (dashboard) {
      dashboard.getRange(13, 2).setValue(successCount);
    }

    // 履歴追加
    const sourceFile = getCurrentSourceFile();
    addHistory('AI予測', sourceFile ? sourceFile.name : '', 0, successCount, 0, '成功');

    ui.alert('Step 2 完了', `${successCount}件のフリガナ予測が完了しました。`, ui.ButtonSet.OK);

  } catch (error) {
    updateDashboard('step2', 'エラー', error.toString());
    console.error('Step 2 エラー:', error);
    SpreadsheetApp.getUi().alert('エラー', 'AI予測中にエラーが発生しました: ' + error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Step 3: データを350件ずつに分割
 */
function step3_splitData() {
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
    const headers = data[0].slice(0, 11); // フリガナ元データとソースファイル列は除外

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
        updateDashboard('step3', '未実行', 'キャンセル');
        ui.alert('処理をキャンセルしました。');
        return;
      }
    }

    // バッチサイズを取得
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

    // データを分割
    let sheetIndex = 1;
    let totalProcessed = 0;

    for (let i = 1; i < data.length; i += batchSize) {
      const batchData = [headers];
      const endIndex = Math.min(i + batchSize, data.length);

      for (let j = i; j < endIndex; j++) {
        // フリガナ元データとソースファイル列を除外してコピー
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

    // Dashboard更新
    updateDashboard('step3', '完了', `${sheetIndex - 1}バッチ作成`);

    // 統計情報更新
    const dashboard = targetSpreadsheet.getSheetByName('Dashboard');
    if (dashboard) {
      dashboard.getRange(14, 2).setValue(sheetIndex - 1);
    }

    // 履歴追加
    const sourceFile = getCurrentSourceFile();
    addHistory('データ分割', sourceFile ? sourceFile.name : '', totalProcessed, 0, sheetIndex - 1, '成功');

    ui.alert(
      'Step 3 完了',
      `データを${sheetIndex - 1}個のバッチに分割しました。\n合計: ${totalProcessed}件`,
      ui.ButtonSet.OK
    );

  } catch (error) {
    updateDashboard('step3', 'エラー', error.toString());
    console.error('Step 3 エラー:', error);
    SpreadsheetApp.getUi().alert('エラー', 'データ分割中にエラーが発生しました: ' + error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Step 4: バッチシートをCSVファイルとして出力
 */
function step4_exportCSV() {
  try {
    const ui = SpreadsheetApp.getUi();
    updateDashboard('step4', '処理中', 'CSV出力開始');

    const targetSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = targetSpreadsheet.getSheets();
    const batchSheets = sheets.filter(sheet => sheet.getName().startsWith('Batch_'));

    if (batchSheets.length === 0) {
      updateDashboard('step4', 'エラー', 'バッチシート未検出');
      ui.alert('エラー', 'バッチシートが見つかりません。Step 3を先に実行してください。', ui.ButtonSet.OK);
      return;
    }

    const response = ui.alert(
      'CSV出力先の選択',
      `${batchSheets.length}個のバッチシートをCSVファイルとして出力します。\n\n出力先を選択してください:\n1. OK: フォルダピッカーで選択\n2. キャンセル: 同じフォルダに保存`,
      ui.ButtonSet.OK_CANCEL
    );

    let folder;
    if (response === ui.Button.OK) {
      // フォルダピッカーを表示
      const html = HtmlService.createHtmlOutputFromFile('folder-picker')
        .setWidth(600)
        .setHeight(400);
      ui.showModalDialog(html, 'CSV出力先フォルダを選択');
      return; // フォルダ選択後、exportCSVToFolderが呼ばれる
    } else {
      // 同じフォルダに保存
      const fileId = targetSpreadsheet.getId();
      const file = DriveApp.getFileById(fileId);
      folder = file.getParents().next();
      exportCSVToFolder(folder.getId());
    }

  } catch (error) {
    updateDashboard('step4', 'エラー', error.toString());
    console.error('Step 4 エラー:', error);
    SpreadsheetApp.getUi().alert('エラー', 'CSV出力中にエラーが発生しました: ' + error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * 指定フォルダにCSVを出力
 */
function exportCSVToFolder(folderId) {
  try {
    const ui = SpreadsheetApp.getUi();
    const targetSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = targetSpreadsheet.getSheets();
    const batchSheets = sheets.filter(sheet => sheet.getName().startsWith('Batch_'));

    const folder = DriveApp.getFolderById(folderId);
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
      const sourceFile = getCurrentSourceFile();
      const sourcePrefix = sourceFile ? sourceFile.name.replace(/\.[^/.]+$/, '') + '_' : '';
      const fileName = `${sourcePrefix}${sheet.getName()}_${timestamp}.csv`;

      // ファイルを作成
      const blob = Utilities.newBlob(csvContent, 'text/csv', fileName);
      folder.createFile(blob);

      console.log(`${fileName} を作成しました`);
      exportedCount++;
    });

    // Dashboard更新
    updateDashboard('step4', '完了', `${exportedCount}ファイル出力`);

    // 統計情報更新
    const dashboard = targetSpreadsheet.getSheetByName('Dashboard');
    if (dashboard) {
      dashboard.getRange(15, 2).setValue(exportedCount);
    }

    // 履歴追加
    const sourceFile = getCurrentSourceFile();
    addHistory('CSV出力', sourceFile ? sourceFile.name : '', 0, 0, exportedCount, '成功', folder.getName());

    ui.alert(
      'Step 4 完了',
      `${exportedCount}個のCSVファイルを出力しました。\nフォルダ: ${folder.getName()}`,
      ui.ButtonSet.OK
    );

  } catch (error) {
    updateDashboard('step4', 'エラー', error.toString());
    console.error('CSV出力エラー:', error);
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
    }

    // 履歴追加
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
    muteHttpExceptions: true,
    timeout: timeout
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

    // Configシートも更新
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

    // Dashboard更新
    const dashboard = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Dashboard');
    if (dashboard) {
      dashboard.getRange(18, 2).setValue('設定済み');
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
// メニュー設定
// ========================================

/**
 * スプレッドシート開いた時にメニューを追加
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu('🔧 反社リスト処理')
    .addSubMenu(ui.createMenu('🏗️ セットアップ')
      .addItem('📊 現在のシートをセットアップ', 'setupCurrentSpreadsheet')
      .addItem('➕ 新規スプレッドシート作成', 'initialSetup')
      .addSeparator()
      .addItem('🔑 OpenAI APIキー設定', 'setApiKey'))
    .addSeparator()
    .addSubMenu(ui.createMenu('📁 データ選択')
      .addItem('📂 ファイルピッカーで選択', 'selectSourceFile')
      .addItem('🔗 URLで指定', 'setSourceFileByUrl')
      .addItem('⏰ 最近使用したファイル', 'selectFromRecentFiles')
      .addItem('📄 現在の選択を確認', 'showCurrentSelection'))
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
【反社リスト変換システム 使い方】

◆ 初回セットアップ
1. 「セットアップ」→「現在のシートをセットアップ」
   - Dashboard, Config, Historyシートを自動作成
2. 「セットアップ」→「OpenAI APIキー設定」
   - OpenAI APIキーを入力

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

Step 3: 350件ごとに分割
   - Master_DataをBatch_1, Batch_2...に分割
   - バッチサイズはConfigで変更可能

Step 4: CSV出力
   - 各バッチをCSVファイルとして出力
   - 出力先フォルダを選択可能

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
反社リスト変換システム
Version: 3.1.0

主な機能:
• 自動セットアップ機能
• Dashboard/Config/History管理
• 動的ファイル選択
• カラム自動検出
• OpenAI APIフリガナ予測
• 350件自動分割（可変）
• フォルダ選択CSV出力

更新履歴:
v3.1.0 - 自動セットアップ追加
v3.0.0 - 動的ファイル選択対応
v2.0.0 - モジュール化
v1.0.0 - 初版リリース

開発: 2025年
  `;

  SpreadsheetApp.getUi().alert('バージョン情報', info, SpreadsheetApp.getUi().ButtonSet.OK);
}