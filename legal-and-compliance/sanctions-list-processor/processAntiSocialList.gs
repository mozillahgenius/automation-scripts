/**
 * 反社リスト変換処理
 * 制裁リスト（原本）を指定フォーマットに変換する
 */

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

    // データを処理（2行目から開始）
    for (let i = 1; i < sourceData.length; i++) {
      const row = sourceData[i];

      // 空行はスキップ
      if (!row[2]) continue; // 名前がない場合はスキップ

      const processedRow = [];

      // 1. 登録日付（今日の日付）
      processedRow.push(today);

      // 2. 顧客区分（個人:01, 法人:02）
      // 組織名に「組」「会」「団」が含まれる場合は法人と判定
      const orgName = row[7] || '';
      const isOrganization = orgName.includes('組') || orgName.includes('会') || orgName.includes('団');
      processedRow.push(isOrganization ? '02' : '01');

      // 3. 日本語名（漢字名）
      processedRow.push(row[2] || ''); // 名前

      // 4. 英文名（ローマ字変換）
      const kanaName = row[4] || ''; // フリガナ
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
 * メニューバーに機能を追加
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('反社リスト処理')
    .addItem('データ変換処理を実行', 'processAntiSocialList')
    .addItem('CSVファイルを出力', 'exportToCSV')
    .addSeparator()
    .addItem('使い方', 'showInstructions')
    .addToUi();
}

/**
 * 使い方を表示
 */
function showInstructions() {
  const instructions = `
【反社リスト変換ツール 使い方】

1. 準備
   - 元データのスプレッドシートIDを設定
   - 出力先のスプレッドシートIDを設定
   - CSV出力用のGoogle DriveフォルダIDを設定

2. 実行手順
   ① 「データ変換処理を実行」をクリック
      → 元データを指定フォーマットに変換し、350件ごとに分割

   ② 「CSVファイルを出力」をクリック
      → 各バッチをCSVファイルとしてDriveに保存

3. 注意事項
   - 個人/法人区分は「01」「02」のテキスト形式で出力
   - 英字名は自動的にローマ字変換
   - 生年月日は年齢から概算で算出
   - データは350件ごとに自動分割

4. 出力形式
   - Batch_1, Batch_2... という名前のシートに分割
   - CSVファイルはUTF-8 BOM付きで出力（Excel対応）
  `;

  SpreadsheetApp.getUi().alert('使い方', instructions, SpreadsheetApp.getUi().ButtonSet.OK);
}