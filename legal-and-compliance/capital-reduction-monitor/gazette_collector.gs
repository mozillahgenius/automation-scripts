/**
 * 官報 本紙PDF自動取得 Google Apps Script
 *
 * 使い方:
 * 1. Google Driveで新しいスプレッドシートを作成
 * 2. 拡張機能 > Apps Script を開く
 * 3. このコードを貼り付け
 * 4. FOLDER_ID を保存先フォルダのIDに変更
 * 5. 実行または定期実行のトリガーを設定
 */

// 設定
const CONFIG = {
  BASE_URL: 'https://www.kanpo.go.jp/',
  FOLDER_ID: 'DRIVE_FOLDER_ID_PLACEHOLDER', // Google DriveのフォルダIDを設定
  SHEET_NAME: '取得ログ' // ログ用シート名
};

/**
 * メイン関数 - 最新の本紙PDFを取得
 */
function fetchLatestKanpoHonshi() {
  const today = new Date();
  const dateStr = Utilities.formatDate(today, 'Asia/Tokyo', 'yyyyMMdd');

  try {
    // トップページをスクレイピングして最新の本紙情報を取得
    const honshiInfo = getLatestHonshiInfo();

    if (!honshiInfo) {
      Logger.log('本紙情報が取得できませんでした');
      return;
    }

    Logger.log('取得対象: ' + JSON.stringify(honshiInfo));

    // PDFをダウンロード
    const pdfUrl = honshiInfo.pdfUrl;
    const fileName = honshiInfo.fileName;

    // 重複チェック
    if (isAlreadyDownloaded(fileName)) {
      Logger.log('既にダウンロード済み: ' + fileName);
      return;
    }

    // PDFを取得してGoogle Driveに保存
    downloadAndSavePdf(pdfUrl, fileName, honshiInfo);

    Logger.log('ダウンロード完了: ' + fileName);

  } catch (error) {
    Logger.log('エラー発生: ' + error.message);
    throw error;
  }
}

/**
 * 官報トップページから最新の本紙情報を取得
 */
function getLatestHonshiInfo() {
  const response = UrlFetchApp.fetch(CONFIG.BASE_URL);
  const html = response.getContentText();

  // 本紙のPDFリンクを正規表現で抽出
  // パターン: ./YYYYMMDD/YYYYMMDDhXXXXX/YYYYMMDDhXXXXXfullXXXXXXXXf.html
  const honshiPattern = /\.\/(\d{8})\/(\d{8}h\d{5})\/(\d{8}h\d{5}full\d{8})f\.html/;
  const match = html.match(honshiPattern);

  if (!match) {
    // 代替パターンを試す
    return getHonshiInfoAlternative(html);
  }

  const dateStr = match[1];
  const folderName = match[2];
  const baseName = match[3];

  // 号数を抽出（例: h01627 → 1627）
  const issueMatch = folderName.match(/h(\d{5})/);
  const issueNumber = issueMatch ? parseInt(issueMatch[1], 10) : 0;

  return {
    date: dateStr,
    issueNumber: issueNumber,
    pdfUrl: CONFIG.BASE_URL + dateStr + '/' + folderName + '/pdf/' + baseName + '.pdf',
    fileName: '官報本紙_' + dateStr + '_第' + issueNumber + '号.pdf',
    htmlUrl: CONFIG.BASE_URL + match[0].substring(2)
  };
}

/**
 * 代替パターンで本紙情報を取得
 */
function getHonshiInfoAlternative(html) {
  // 「本紙」リンクの近くにあるfullリンクを探す
  const sections = html.split('本紙');

  if (sections.length < 2) return null;

  // 本紙セクション以降でfullリンクを探す
  const afterHonshi = sections[1].substring(0, 2000);
  const fullPattern = /href="\.\/(\d{8})\/(\d{8}h\d{5})\/(\d{8}h\d{5}full\d{8})f\.html"/;
  const match = afterHonshi.match(fullPattern);

  if (!match) return null;

  const dateStr = match[1];
  const folderName = match[2];
  const baseName = match[3];
  const issueMatch = folderName.match(/h(\d{5})/);
  const issueNumber = issueMatch ? parseInt(issueMatch[1], 10) : 0;

  return {
    date: dateStr,
    issueNumber: issueNumber,
    pdfUrl: CONFIG.BASE_URL + dateStr + '/' + folderName + '/pdf/' + baseName + '.pdf',
    fileName: '官報本紙_' + dateStr + '_第' + issueNumber + '号.pdf',
    htmlUrl: CONFIG.BASE_URL + dateStr + '/' + folderName + '/' + baseName + 'f.html'
  };
}

/**
 * PDFをダウンロードしてGoogle Driveに保存
 */
function downloadAndSavePdf(pdfUrl, fileName, info) {
  const response = UrlFetchApp.fetch(pdfUrl);
  const blob = response.getBlob().setName(fileName);

  const folder = DriveApp.getFolderById(CONFIG.FOLDER_ID);
  const file = folder.createFile(blob);

  // ログを記録
  logDownload(info, file.getUrl());

  return file;
}

/**
 * 重複ダウンロードチェック
 */
function isAlreadyDownloaded(fileName) {
  const folder = DriveApp.getFolderById(CONFIG.FOLDER_ID);
  const files = folder.getFilesByName(fileName);
  return files.hasNext();
}

/**
 * ダウンロードログを記録
 */
function logDownload(info, driveUrl) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_NAME);
    sheet.appendRow(['取得日時', '発行日', '号数', 'PDFファイル名', 'Drive URL', 'PDF元URL']);
  }

  sheet.appendRow([
    new Date(),
    info.date,
    info.issueNumber,
    info.fileName,
    driveUrl,
    info.pdfUrl
  ]);
}

/**
 * 指定日の本紙PDFを取得（過去分取得用）
 * @param {string} dateStr - YYYYMMDD形式の日付
 */
function fetchKanpoByDate(dateStr) {
  // トップページから対象日のリンクを探す
  const response = UrlFetchApp.fetch(CONFIG.BASE_URL);
  const html = response.getContentText();

  // 指定日のパターンを検索
  const pattern = new RegExp('\\./' + dateStr + '/' + dateStr + 'h(\\d{5})/' + dateStr + 'h\\d{5}full(\\d{8})f\\.html');
  const match = html.match(pattern);

  if (!match) {
    Logger.log(dateStr + 'の本紙が見つかりません');
    return null;
  }

  const issueNumber = parseInt(match[1], 10);
  const pageRange = match[2];
  const folderName = dateStr + 'h' + match[1];
  const baseName = dateStr + 'h' + match[1] + 'full' + pageRange;

  const info = {
    date: dateStr,
    issueNumber: issueNumber,
    pdfUrl: CONFIG.BASE_URL + dateStr + '/' + folderName + '/pdf/' + baseName + '.pdf',
    fileName: '官報本紙_' + dateStr + '_第' + issueNumber + '号.pdf'
  };

  if (isAlreadyDownloaded(info.fileName)) {
    Logger.log('既にダウンロード済み: ' + info.fileName);
    return null;
  }

  downloadAndSavePdf(info.pdfUrl, info.fileName, info);
  Logger.log('ダウンロード完了: ' + info.fileName);
  return info;
}

/**
 * 直近N日分の本紙を一括取得
 * @param {number} days - 取得する日数
 */
function fetchRecentKanpo(days) {
  days = days || 7;
  const response = UrlFetchApp.fetch(CONFIG.BASE_URL);
  const html = response.getContentText();

  // すべての本紙PDFリンクを抽出
  const pattern = /\.\/(\d{8})\/(\d{8}h\d{5})\/(\d{8}h\d{5}full\d{8})f\.html/g;
  const matches = [];
  let match;
  const seen = {};

  while ((match = pattern.exec(html)) !== null) {
    const dateStr = match[1];
    if (!seen[dateStr]) {
      seen[dateStr] = true;
      matches.push({
        date: dateStr,
        folder: match[2],
        base: match[3]
      });
    }
  }

  // 日付でソート（新しい順）
  matches.sort((a, b) => b.date.localeCompare(a.date));

  // 指定日数分を取得
  const toFetch = matches.slice(0, days);

  toFetch.forEach(item => {
    const issueMatch = item.folder.match(/h(\d{5})/);
    const issueNumber = issueMatch ? parseInt(issueMatch[1], 10) : 0;

    const info = {
      date: item.date,
      issueNumber: issueNumber,
      pdfUrl: CONFIG.BASE_URL + item.date + '/' + item.folder + '/pdf/' + item.base + '.pdf',
      fileName: '官報本紙_' + item.date + '_第' + issueNumber + '号.pdf'
    };

    if (!isAlreadyDownloaded(info.fileName)) {
      try {
        downloadAndSavePdf(info.pdfUrl, info.fileName, info);
        Logger.log('ダウンロード完了: ' + info.fileName);
        Utilities.sleep(1000); // サーバー負荷軽減
      } catch (e) {
        Logger.log('ダウンロード失敗: ' + info.fileName + ' - ' + e.message);
      }
    } else {
      Logger.log('スキップ（既存）: ' + info.fileName);
    }
  });
}

/**
 * 毎日定時実行用トリガーを設定
 */
function createDailyTrigger() {
  // 既存のトリガーを削除
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'fetchLatestKanpoHonshi') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // 毎日午前9時に実行するトリガーを作成
  ScriptApp.newTrigger('fetchLatestKanpoHonshi')
    .timeBased()
    .atHour(9)
    .everyDays(1)
    .inTimezone('Asia/Tokyo')
    .create();

  Logger.log('毎日午前9時に実行するトリガーを設定しました');
}
