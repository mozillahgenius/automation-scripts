/**
 * 反社リスト変換処理（OpenAI API対応版）
 * 制裁リスト（原本）を指定フォーマットに変換する
 * 読み仮名がない場合はOpenAI APIを使用して予測
 */

// OpenAI API設定
const OPENAI_API_KEY = 'your-openai-api-key-here'; // TODO: 実際のAPIキーに置き換え
const OPENAI_API_URL = 'https://api.openai.com/v1/chat/completions';

function processAntiSocialList() {
  try {
    const sourceSheetId = '元データのスプレッドシートID'; // TODO: 実際のIDに置き換え
    const targetSheetId = '出力先のスプレッドシートID'; // TODO: 実際のIDに置き換え

    const sourceSheet = SpreadsheetApp.openById(sourceSheetId).getSheetByName('基本データ');
    const targetSpreadsheet = SpreadsheetApp.openById(targetSheetId);

    const sourceData = sourceSheet.getDataRange().getValues();
    const headers = sourceData[0];

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
      '비고'
    ];
    processedData.push(targetHeaders);

    // バッチ処理用の配列（OpenAI API呼び出し最適化）
    const namesToProcess = [];
    const rowIndices = [];

    // データを処理（2行目から開始）
    for (let i = 1; i < sourceData.length; i++) {
      const row = sourceData[i];

      // 空行はスキップ
      if (!row[2]) continue;

      const kanaName = row[4] || ''; // フリガナ
      const kanjiName = row[2] || ''; // 漢字名

      // フリガナがない場合、後でOpenAI APIで処理
      if (!kanaName && kanjiName) {
        namesToProcess.push(kanjiName);
        rowIndices.push(i);
      }
    }

    // OpenAI APIでフリガナを一括予測
    const predictedKanaNames = {};
    if (namesToProcess.length > 0) {
      console.log(`${namesToProcess.length}件の名前でフリガナを予測します...`);
      const predictions = batchPredictKanaNames(namesToProcess);

      for (let j = 0; j < namesToProcess.length; j++) {
        predictedKanaNames[rowIndices[j]] = predictions[j];
      }
    }

    // 再度データを処理して最終的な出力を作成
    for (let i = 1; i < sourceData.length; i++) {
      const row = sourceData[i];

      // 空行はスキップ
      if (!row[2]) continue;

      const processedRow = [];

      // 1. 登録日付（今日の日付）
      processedRow.push(today);

      // 2. 顧客区分（個人:01, 法人:02）
      const orgName = row[7] || '';
      const isOrganization = orgName.includes('組') || orgName.includes('会') || orgName.includes('団');
      processedRow.push(isOrganization ? '02' : '01');

      // 3. 日本語名（漢字名）
      const kanjiName = row[2] || '';
      processedRow.push(kanjiName);

      // 4. 英文名（ローマ字変換）
      let kanaName = row[4] || ''; // フリガナ

      // フリガナがない場合は予測値を使用
      if (!kanaName && predictedKanaNames[i]) {
        kanaName = predictedKanaNames[i];
        console.log(`予測フリガナ使用: ${kanjiName} → ${kanaName}`);
      }

      const romajiName = convertKanaToRomaji(kanaName);
      processedRow.push(romajiName);

      // 5. 性別（男:1, 女:2, 不明:空欄）
      const gender = row[6] || '';
      let genderCode = '';
      if (gender === '男') genderCode = '1';
      else if (gender === '女') genderCode = '2';
      processedRow.push(genderCode);

      // 6. 生年月日（年齢から逆算 - 概算）
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
      processedRow.push(row[8] || ''); // 住所

      // 11. 備考
      const remarks = [];
      if (row[3]) remarks.push('異名: ' + row[3]); // 異名
      if (row[7]) remarks.push('組織: ' + row[7]); // 組織・団体
      if (row[9]) remarks.push(row[9]); // 内容
      if (!row[4] && predictedKanaNames[i]) {
        remarks.push('※フリガナはAI予測');
      }
      processedRow.push(remarks.join(' / '));

      processedData.push(processedRow);
    }

    // データを350件ずつに分割して出力
    const batchSize = 350;
    let sheetIndex = 1;

    for (let i = 0; i < processedData.length - 1; i += batchSize) {
      const batchData = [targetHeaders].concat(
        processedData.slice(i + 1, Math.min(i + 1 + batchSize, processedData.length))
      );

      // 新しいシートを作成または既存のシートをクリア
      let targetSheet;
      const sheetName = `Batch_${sheetIndex}`;

      try {
        targetSheet = targetSpreadsheet.getSheetByName(sheetName);
        targetSheet.clear();
      } catch (e) {
        targetSheet = targetSpreadsheet.insertSheet(sheetName);
      }

      // データを書き込み
      if (batchData.length > 1) {
        targetSheet.getRange(1, 1, batchData.length, batchData[0].length).setValues(batchData);

        // 顧客区分列を文字列形式に設定
        const customerTypeRange = targetSheet.getRange(2, 2, batchData.length - 1, 1);
        customerTypeRange.setNumberFormat('@'); // テキスト形式

        console.log(`${sheetName}: ${batchData.length - 1}件のデータを出力しました`);
      }

      sheetIndex++;
    }

    console.log(`処理完了: 合計 ${processedData.length - 1}件のデータを処理しました`);

  } catch (error) {
    console.error('エラーが発生しました:', error);
    throw error;
  }
}

/**
 * OpenAI APIを使用して漢字名からフリガナを予測（バッチ処理）
 * @param {string[]} kanjiNames - 漢字名の配列
 * @return {string[]} 予測されたフリガナの配列
 */
function batchPredictKanaNames(kanjiNames) {
  const predictions = [];
  const batchSize = 10; // 一度に処理する名前の数

  for (let i = 0; i < kanjiNames.length; i += batchSize) {
    const batch = kanjiNames.slice(i, Math.min(i + batchSize, kanjiNames.length));

    try {
      const prompt = `以下の日本人の名前（漢字）について、それぞれのカタカナ読みを推測してください。
姓と名の間にはスペースを入れてください。
出力形式：漢字名|カタカナ読み

名前リスト：
${batch.join('\n')}`;

      const response = callOpenAI(prompt);

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
        predictions.push(''); // 空文字を追加
      }

    } catch (error) {
      console.error(`バッチ ${i / batchSize + 1} の処理中にエラー:`, error);
      // エラーの場合は空文字を追加
      for (let j = 0; j < batch.length; j++) {
        predictions.push('');
      }
    }

    // API制限を考慮して待機
    if (i + batchSize < kanjiNames.length) {
      Utilities.sleep(1000); // 1秒待機
    }
  }

  return predictions;
}

/**
 * OpenAI APIを呼び出す
 * @param {string} prompt - プロンプト
 * @return {string} レスポンステキスト
 */
function callOpenAI(prompt) {
  const payload = {
    model: 'gpt-4o-mini', // または gpt-3.5-turbo
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
      'Authorization': `Bearer ${OPENAI_API_KEY}`
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
 * @param {string} kana - カタカナ文字列
 * @return {string} ローマ字文字列
 */
function convertKanaToRomaji(kana) {
  if (!kana) return '';

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
    // 小文字の処理（拗音）
    if (i < hiragana.length - 1) {
      const char = hiragana[i];
      const nextChar = hiragana[i + 1];

      // 拗音の処理（きゃ、きゅ、きょ など）
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

      // 促音の処理（っ）
      if (char === 'っ' && nextChar) {
        const nextRomaji = conversionTable[nextChar] || nextChar;
        if (nextRomaji && nextRomaji[0]) {
          romaji += nextRomaji[0];
        }
        i++;
        continue;
      }
    }

    // 通常の変換
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

/**
 * CSVファイルとして出力
 */
function exportToCSV() {
  const targetSheetId = '出力先のスプレッドシートID'; // TODO: 実際のIDに置き換え
  const targetSpreadsheet = SpreadsheetApp.openById(targetSheetId);
  const sheets = targetSpreadsheet.getSheets();

  const folderId = 'Google DriveのフォルダID'; // TODO: 実際のIDに置き換え
  const folder = DriveApp.getFolderById(folderId);

  sheets.forEach(sheet => {
    if (sheet.getName().startsWith('Batch_')) {
      const data = sheet.getDataRange().getValues();

      // CSVコンテンツを作成
      let csvContent = '';
      data.forEach(row => {
        const csvRow = row.map(cell => {
          // セルの値を文字列に変換し、必要に応じてエスケープ
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

      // ファイルを作成
      const blob = Utilities.newBlob(csvContent, 'text/csv', `${sheet.getName()}.csv`);
      folder.createFile(blob);

      console.log(`${sheet.getName()}.csv を作成しました`);
    }
  });
}

/**
 * APIキーを設定（スクリプトプロパティ使用）
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
 * メニューバーに機能を追加
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('反社リスト処理')
    .addItem('データ変換処理を実行', 'processAntiSocialList')
    .addItem('CSVファイルを出力', 'exportToCSV')
    .addSeparator()
    .addItem('OpenAI APIキーを設定', 'setApiKey')
    .addItem('使い方', 'showInstructions')
    .addToUi();
}

/**
 * 使い方を表示
 */
function showInstructions() {
  const instructions = `
【反社リスト変換ツール 使い方（AI対応版）】

1. 初期設定
   ① OpenAI APIキーの設定
      - メニュー → 「OpenAI APIキーを設定」
      - APIキーを入力（https://platform.openai.com/api-keys で取得）

   ② スプレッドシートIDの設定
      - 元データのスプレッドシートID
      - 出力先のスプレッドシートID
      - CSV出力用のGoogle DriveフォルダID

2. 実行手順
   ① 「データ変換処理を実行」をクリック
      → フリガナがない名前は自動的にOpenAI APIで予測
      → 元データを指定フォーマットに変換し、350件ごとに分割

   ② 「CSVファイルを出力」をクリック
      → 各バッチをCSVファイルとしてDriveに保存

3. AI機能の特徴
   - フリガナが記載されていない漢字名を自動で読み仮名予測
   - 10件ずつバッチ処理で効率的に変換
   - 予測されたフリガナは備考欄に「※フリガナはAI予測」と記載

4. 注意事項
   - OpenAI APIの使用量に応じて課金される場合があります
   - API制限を考慮し、大量データの場合は処理時間がかかります
   - 個人/法人区分は「01」「02」のテキスト形式で出力
   - データは350件ごとに自動分割

5. 出力形式
   - Batch_1, Batch_2... という名前のシートに分割
   - CSVファイルはUTF-8 BOM付きで出力（Excel対応）
  `;

  SpreadsheetApp.getUi().alert('使い方', instructions, SpreadsheetApp.getUi().ButtonSet.OK);
}