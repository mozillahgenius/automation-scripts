/**
 * スプレッドシートのデータから Google Docs を自動生成するスクリプト
 *
 * 使用方法:
 * 1. このコードを Google スプレッドシートのスクリプトエディタに貼り付ける
 * 2. 以下の定数を実際の値に置き換える
 * 3. 実行前にスプレッドシートの構成を確認する
 */

// ========== 設定項目 ==========
// これらの値を実際のIDやURLに置き換えてください

// Google Docs テンプレートのID（URLから取得）
const TEMPLATE_DOC_ID = '1tdbloeGpx3O9zBRMxdFGPcNzJW-DD6HiQhoU9TIwZhM';

// 生成したドキュメントを保存するフォルダのID
const OUTPUT_FOLDER_ID = '1W_zUAeHpIK7mbCoIRqWkueTdqFSp1q6D';

// スプレッドシートの設定
const SHEET_NAME = '特許'; // 対象のシート名
const START_ROW = 2; // データの開始行（ヘッダーを除く）

// 処理範囲の設定（No列の値で指定）
const PROCESS_FROM = 1;  // 処理開始番号
const PROCESS_TO = 78;   // 処理終了番号（78番まで）

// 列の設定（A列=1, B列=2, ..., G列=7, H列=8）
const COL_NO = 1;                // A列: No
const COL_AMOUNT = 2;            // B列: 金額
const COL_SERVICE_NAME_ORIG = 7; // G列: 会社名・サービス名（元データ）
const COL_SERVICE_NAME_ADJ = 8;  // H列: 会社名・サービス名（調整後）

// ========== メイン処理 ==========

/**
 * メニューバーに実行ボタンを追加
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('ドキュメント生成')
    .addItem('テンプレートから生成', 'generateDocuments')
    .addItem('設定確認', 'checkSettings')
    .addToUi();
}

/**
 * 設定を確認する関数
 */
function checkSettings() {
  const messages = [];

  // テンプレートドキュメントの確認
  try {
    const templateDoc = DriveApp.getFileById(TEMPLATE_DOC_ID);
    messages.push(`✅ テンプレート: ${templateDoc.getName()}`);
  } catch (e) {
    messages.push('❌ テンプレートドキュメントが見つかりません。IDを確認してください。');
  }

  // 出力フォルダの確認
  try {
    const folder = DriveApp.getFolderById(OUTPUT_FOLDER_ID);
    messages.push(`✅ 出力フォルダ: ${folder.getName()}`);
  } catch (e) {
    messages.push('❌ 出力フォルダが見つかりません。IDを確認してください。');
  }

  // シートの確認
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (sheet) {
    messages.push(`✅ 対象シート: ${SHEET_NAME}`);
  } else {
    messages.push(`❌ シート「${SHEET_NAME}」が見つかりません。`);
  }

  SpreadsheetApp.getUi().alert('設定確認', messages.join('\\n'), SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * スプレッドシートからドキュメントを生成するメイン関数
 */
function generateDocuments() {
  try {
    // スプレッドシートとシートを取得
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName(SHEET_NAME);

    if (!sheet) {
      throw new Error(`シート「${SHEET_NAME}」が見つかりません`);
    }

    // データの最終行を取得
    const lastRow = sheet.getLastRow();
    if (lastRow < START_ROW) {
      SpreadsheetApp.getUi().alert('エラー', 'データが見つかりません', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }

    // テンプレートドキュメントを取得
    const templateFile = DriveApp.getFileById(TEMPLATE_DOC_ID);

    // 出力フォルダを取得
    const outputFolder = DriveApp.getFolderById(OUTPUT_FOLDER_ID);

    // 処理対象行を計算
    let processCount = 0;
    for (let row = START_ROW; row <= lastRow; row++) {
      const no = sheet.getRange(row, COL_NO).getValue();
      if (no >= PROCESS_FROM && no <= PROCESS_TO) {
        processCount++;
      }
    }

    // 処理の確認
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      '確認',
      `No ${PROCESS_FROM} から ${PROCESS_TO} までの ${processCount} 件のドキュメントを生成します。\\n続行しますか？`,
      ui.ButtonSet.YES_NO
    );

    if (response !== ui.Button.YES) {
      return;
    }

    let successCount = 0;
    let errorCount = 0;
    const errors = [];

    // 各行のデータを処理
    for (let row = START_ROW; row <= lastRow; row++) {
      try {
        // データを取得
        const no = sheet.getRange(row, COL_NO).getValue();

        // 処理範囲外の場合はスキップ
        if (no < PROCESS_FROM || no > PROCESS_TO) {
          continue;
        }

        const amount = sheet.getRange(row, COL_AMOUNT).getValue();

        // H列（調整後）のデータを優先的に使用、なければG列（元データ）を使用
        let serviceName = sheet.getRange(row, COL_SERVICE_NAME_ADJ).getValue();
        if (!serviceName || serviceName === '') {
          serviceName = sheet.getRange(row, COL_SERVICE_NAME_ORIG).getValue();
        }

        // 空行をスキップ
        if (!serviceName || serviceName === '') {
          continue;
        }

        // ファイル名を生成（例: "001_株式会社ABC"）
        const fileName = `${String(no).padStart(3, '0')}_${serviceName}`;

        // テンプレートをコピー
        const newDocFile = templateFile.makeCopy(fileName, outputFolder);

        // ドキュメントを開いて編集
        const newDoc = DocumentApp.openById(newDocFile.getId());
        const body = newDoc.getBody();

        // プレースホルダーを置換
        replaceText(body, '{{COMPANY_SERVICE}}', serviceName);
        replaceText(body, '{{Money_Amount}}', formatCurrency(amount));

        // ドキュメントを保存
        newDoc.saveAndClose();

        successCount++;

        // 進捗をログに記録
        console.log(`✅ 生成完了: ${fileName}`);

      } catch (e) {
        errorCount++;
        errors.push(`行${row}: ${e.message}`);
        console.error(`❌ エラー (行${row}): ${e.message}`);
      }
    }

    // 処理結果を表示
    let message = `処理が完了しました。\\n\\n`;
    message += `✅ 成功: ${successCount}件\\n`;
    if (errorCount > 0) {
      message += `❌ エラー: ${errorCount}件\\n\\n`;
      message += `エラー詳細:\\n${errors.slice(0, 5).join('\\n')}`;
      if (errors.length > 5) {
        message += `\\n... 他${errors.length - 5}件`;
      }
    }

    ui.alert('処理完了', message, ui.ButtonSet.OK);

  } catch (e) {
    SpreadsheetApp.getUi().alert('エラー', `処理中にエラーが発生しました:\\n${e.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
    console.error(e);
  }
}

/**
 * ドキュメント内のテキストを置換する関数
 * @param {Body} body - ドキュメントのボディ
 * @param {string} searchText - 検索するテキスト
 * @param {string} replaceText - 置換するテキスト
 */
function replaceText(body, searchText, replaceText) {
  body.replaceText(searchText, replaceText);
}

/**
 * 金額を通貨形式でフォーマットする関数
 * @param {number} amount - 金額
 * @return {string} フォーマット済みの金額文字列
 */
function formatCurrency(amount) {
  if (typeof amount !== 'number') {
    return String(amount);
  }
  return amount.toLocaleString('ja-JP');
}

/**
 * テスト用関数：1件だけ生成
 */
function testGenerateSingleDocument() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);

  if (!sheet) {
    SpreadsheetApp.getUi().alert('エラー', `シート「${SHEET_NAME}」が見つかりません`, SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  // 最初のデータ行を取得
  const row = START_ROW;
  const no = sheet.getRange(row, COL_NO).getValue();
  const amount = sheet.getRange(row, COL_AMOUNT).getValue();

  // H列（調整後）のデータを優先的に使用、なければG列（元データ）を使用
  let serviceName = sheet.getRange(row, COL_SERVICE_NAME_ADJ).getValue();
  if (!serviceName || serviceName === '') {
    serviceName = sheet.getRange(row, COL_SERVICE_NAME_ORIG).getValue();
  }

  console.log('テストデータ:');
  console.log(`No: ${no}`);
  console.log(`金額: ${amount}`);
  console.log(`会社・サービス名: ${serviceName}`);

  if (!serviceName) {
    SpreadsheetApp.getUi().alert('エラー', 'データが見つかりません', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  try {
    // テンプレートをコピー
    const templateFile = DriveApp.getFileById(TEMPLATE_DOC_ID);
    const outputFolder = DriveApp.getFolderById(OUTPUT_FOLDER_ID);
    const fileName = `TEST_${String(no).padStart(3, '0')}_${serviceName}`;

    const newDocFile = templateFile.makeCopy(fileName, outputFolder);
    const newDoc = DocumentApp.openById(newDocFile.getId());
    const body = newDoc.getBody();

    // プレースホルダーを置換
    replaceText(body, '{{COMPANY_SERVICE}}', serviceName);
    replaceText(body, '{{Money_Amount}}', formatCurrency(amount));

    newDoc.saveAndClose();

    SpreadsheetApp.getUi().alert('成功', `テストドキュメント「${fileName}」を生成しました`, SpreadsheetApp.getUi().ButtonSet.OK);

  } catch (e) {
    SpreadsheetApp.getUi().alert('エラー', `エラーが発生しました:\\n${e.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}