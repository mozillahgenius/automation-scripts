//=ARRAYFORMULA(if(R2:R="","",TEXT(R2:R, "yyyy-mm")))
// ▼▼▼【設定箇所】▼▼▼
// 1. コピー元のスプレッドシートのURLを貼り付けてください
const SOURCE_SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/SPREADSHEET_ID_PLACEHOLDER/edit?pli=1&gid=212937189#gid=212937189";

// 2. コピー元のシート名を指定してください
const SOURCE_SHEET_NAME = "ソアスク+台帳データ";

// 3. 貼り付け先のシート名を指定してください
const DESTINATION_SHEET_NAME = "実績DB";
// ▲▲▲【設定箇所】▲▲▲


/**
 * 別のスプレッドシートの指定シートからA列以降の全データをコピーし、
 * このスプレッドシートの指定シートのA列から値のみを貼り付けます。
 */
function copyAllColumns() {
  try {
    // コピー元のスプレッドシートとシートを名前で取得
    const sourceSpreadsheet = SpreadsheetApp.openByUrl(SOURCE_SPREADSHEET_URL);
    const sourceSheet = sourceSpreadsheet.getSheetByName(SOURCE_SHEET_NAME);

    if (!sourceSheet) {
      throw new Error("指定されたコピー元シートが見つかりませんでした: " + SOURCE_SHEET_NAME);
    }
    
    // コピー元の全データを取得（A列から全て）
    const sourceData = sourceSheet.getDataRange().getValues();

    // 貼り付け先のスプレッドシートとシートを取得
    const destinationSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    let destinationSheet = destinationSpreadsheet.getSheetByName(DESTINATION_SHEET_NAME);

    if (!destinationSheet) {
      destinationSheet = destinationSpreadsheet.insertSheet(DESTINATION_SHEET_NAME);
    }

    // 貼り付け先のシートの内容を一度すべてクリア
    destinationSheet.clearContents();

    // A列以降の全データをA列から貼り付け
    if (sourceData.length > 0 && sourceData[0].length > 0) {
      destinationSheet.getRange(1, 1, sourceData.length, sourceData[0].length).setValues(sourceData);
      
      // N列（売上金額）を日本円の通貨形式に設定（ヘッダー除外で2行目から）
      if (sourceData.length > 1) {
        destinationSheet
          .getRange(2, 14, sourceData.length - 1, 1) // N列
          .setNumberFormat('[$¥-ja-JP]#,##0;[Red]-[$¥-ja-JP]#,##0');
      }
    }
    
    console.log("A列以降の全データのコピーが正常に完了しました。");

  } catch (e) {
    console.error("エラーが発生しました: " + e.toString());
  }
}