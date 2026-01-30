/**
 * ============================================================
 * 共有ドライブ特定フォルダ内の日本語テキストPDFを日次増分で処理し、
 * キーワード（減資 / 資本金の額の減少 / 準備金の額の減少）に合致するPDFから
 * 会社名（法人格を含むもののみ）と住所（都道府県起点）を抽出して
 * スプレッドシートへPDF単位で追記する。
 *
 * - 処理済みはシートの file_id で管理（増分処理）
 * - 会社名と住所は1行1件で出力（行を分割）
 *
 * 重要：
 * - Advanced Google Services の「Drive API」を有効化してください（Drive.Files.copy / Drive.Files.get を使用）
 *   Apps Script エディタ → サービス → Drive API を追加
 * - 共有ドライブ対応：Drive API 呼び出しに supportsAllDrives: true を付与
 * - PDFがスキャン画像のみの場合は本方式ではテキスト抽出できません（OCRは未導入）
 * ============================================================
 */

/**
 * ========= 設定 =========
 * TODO: SPREADSHEET_ID と FOLDER_ID を必ず設定
 */
const CONFIG = {
  // 出力先スプレッドシートID
  SPREADSHEET_ID: "1f036RjFfzHyHQTLD1lM_73qKDQnZDXdeUXYC9RVce-w",

  // 出力先シート名（存在しなければ作成）
  SHEET_NAME: "減資抽出",

  // 共有ドライブ内の対象フォルダID
  FOLDER_ID: "1IVQTgWiMIftmc9ygO48WqWJsEGAoGEII",

  // キーワード（OR条件・単純一致）
  KEYWORDS: [
    "減資",
    "資本金の額の減少",
    "準備金の額の減少"
  ],

  // 会社名抽出：法人格を含むもののみ
  COMPANY_FORMS: ["株式会社", "合同会社", "有限会社", "合名会社", "合資会社"],

  // 住所抽出：都道府県起点
  PREFS: [
    "北海道","青森県","岩手県","宮城県","秋田県","山形県","福島県",
    "茨城県","栃木県","群馬県","埼玉県","千葉県","東京都","神奈川県",
    "新潟県","富山県","石川県","福井県","山梨県","長野県",
    "岐阜県","静岡県","愛知県","三重県",
    "滋賀県","京都府","大阪府","兵庫県","奈良県","和歌山県",
    "鳥取県","島根県","岡山県","広島県","山口県",
    "徳島県","香川県","愛媛県","高知県",
    "福岡県","佐賀県","長崎県","熊本県","大分県","宮崎県","鹿児島県",
    "沖縄県"
  ],

  // トリガ実行時刻（setupDailyTrigger() で使用）
  TRIGGER_HOUR: 2
};

/**
 * ========= スプレッドシート初期設定 =========
 * - シート作成/取得
 * - ヘッダ作成
 * - フィルタ/固定など最低限の見やすさを設定
 *
 * 使い方：
 *   1) CONFIG の SPREADSHEET_ID / SHEET_NAME を設定
 *   2) setupSpreadsheet() を1回手動実行
 */
function setupSpreadsheet() {
  assertConfig_();

  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = getOrCreateSheet_(ss, CONFIG.SHEET_NAME);

  ensureHeader_(sheet);
  sheet.setFrozenRows(1);

  // フィルタ（既にあれば無視）
  const lastCol = sheet.getLastColumn();
  const lastRow = sheet.getLastRow();
  if (lastRow >= 1) {
    const range = sheet.getRange(1, 1, Math.max(1, lastRow), lastCol);
    try {
      if (!sheet.getFilter()) range.createFilter();
    } catch (e) {
      // 既にある/作れない場合は無視
    }
  }

  // 列幅（任意）
  try {
    sheet.setColumnWidth(1, 170); // processed_at
    sheet.setColumnWidth(2, 300); // pdf_name
    sheet.setColumnWidth(3, 320); // pdf_url
    sheet.setColumnWidth(4, 260); // file_id
    sheet.setColumnWidth(5, 110); // has_match
    sheet.setColumnWidth(6, 520); // companies
    sheet.setColumnWidth(7, 520); // addresses
    sheet.setColumnWidth(8, 520); // note
  } catch (e) {}

  // 1行目見た目（任意）
  try {
    sheet.getRange(1, 1, 1, lastCol).setFontWeight("bold");
  } catch (e) {}
}

/**
 * ========= 日次トリガ設定（初回のみ手動実行） =========
 * - runDaily() を毎日1回実行するトリガを作成
 */
function setupDailyTrigger() {
  assertConfig_();

  ScriptApp.newTrigger("runDaily")
    .timeBased()
    .everyDays(1)
    .atHour(CONFIG.TRIGGER_HOUR)
    .create();
}

/**
 * 既存の runDaily トリガを消して作り直す
 */
function resetDailyTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  for (const t of triggers) {
    if (t.getHandlerFunction() === "runDaily") {
      ScriptApp.deleteTrigger(t);
    }
  }
  setupDailyTrigger();
}

/**
 * ========= メイン処理（トリガで毎日実行） =========
 * - 対象フォルダ内のPDFを列挙
 * - シートの file_id を参照して未処理だけ処理
 * - PDF -> Google Docs変換で本文抽出（Drive API）
 * - KEYWORDS のいずれかが含まれる場合のみ抽出して追記
 * - 会社名・住所は1行1件で出力（行を分割）
 */
function runDaily() {
  assertConfig_();

  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = getOrCreateSheet_(ss, CONFIG.SHEET_NAME);
  ensureHeader_(sheet);

  const processed = loadProcessedFileIds_(sheet);

  let folder;
  try {
    folder = DriveApp.getFolderById(CONFIG.FOLDER_ID);
  } catch (e) {
    throw new Error("FOLDER_ID が不正、またはアクセス権がありません: " + String(e));
  }

  const files = folder.getFilesByType(MimeType.PDF);

  const nowIso = new Date().toISOString();
  const rowsToAppend = [];

  while (files.hasNext()) {
    const file = files.next();
    const fileId = file.getId();

    // 増分のみ
    if (processed.has(fileId)) continue;

    const pdfName = file.getName();
    const pdfUrl = file.getUrl();

    // テキスト抽出
    let text = "";
    try {
      text = extractTextFromPdfViaDocs_(fileId);
    } catch (e) {
      rowsToAppend.push([
        nowIso,
        pdfName,
        pdfUrl,
        fileId,
        false,
        "",
        "",
        "TEXT_EXTRACT_ERROR: " + String(e)
      ]);
      continue;
    }

    // キーワード判定（OR条件・単純一致）
    const hitKeywords = CONFIG.KEYWORDS.filter(kw => text.includes(kw));
    const hasMatch = hitKeywords.length > 0;

    if (!hasMatch) {
      // 不一致でも処理済みにする（増分処理を安定させるため）
      rowsToAppend.push([
        nowIso,
        pdfName,
        pdfUrl,
        fileId,
        false,
        "",
        "",
        ""
      ]);
      continue;
    }

    const companies = extractCompanies_(text);
    const addresses = extractAddresses_(text);

    // 会社名と住所の多い方の件数に合わせて行を生成（1行1件）
    const maxLen = Math.max(companies.length, addresses.length, 1);

    for (let i = 0; i < maxLen; i++) {
      rowsToAppend.push([
        i === 0 ? nowIso : "",           // 最初の行のみ日時
        i === 0 ? pdfName : "",          // 最初の行のみPDF名
        i === 0 ? pdfUrl : "",           // 最初の行のみURL
        i === 0 ? fileId : "",           // 最初の行のみfile_id
        i === 0 ? true : "",             // 最初の行のみhas_match
        companies[i] || "",              // 会社名（1行1件）
        addresses[i] || "",              // 住所（1行1件）
        i === 0 ? "hit_keywords=" + hitKeywords.join(",") : ""
      ]);
    }
  }

  if (rowsToAppend.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rowsToAppend.length, rowsToAppend[0].length)
      .setValues(rowsToAppend);
  }
}

/**
 * ========= テキスト抽出 =========
 * Drive API（高度なGoogleサービス）を使い、PDFをGoogleドキュメントに変換して本文テキストを取得。
 * 取得後、変換ドキュメントは削除してDriveを汚さない。
 *
 * 共有ドライブ対応：
 * - Drive.Files.get / Drive.Files.copy の両方に supportsAllDrives: true を付ける
 */
function extractTextFromPdfViaDocs_(pdfFileId) {
  // 共有ドライブ対応：存在・権限確認（File not found の切り分け）
  Drive.Files.get(pdfFileId, { supportsAllDrives: true });

  const resource = {
    title: "tmp_pdf_extract_" + pdfFileId,
    mimeType: MimeType.GOOGLE_DOCS
  };

  // convert: true でGoogleドキュメント化（テキストPDF前提）
  const converted = Drive.Files.copy(resource, pdfFileId, {
    convert: true,
    supportsAllDrives: true
  });

  const docId = converted.id;

  try {
    const doc = DocumentApp.openById(docId);
    const text = doc.getBody().getText() || "";
    return normalizeText_(text);
  } finally {
    // 一時ドキュメント削除
    try { DriveApp.getFileById(docId).setTrashed(true); } catch (e) {}
  }
}

/**
 * ========= 抽出ロジック（会社名） =========
 * - 法人格を含む行を拾う（本文中前提）
 * - 会社名と住所は紐付けない
 */
function extractCompanies_(text) {
  const forms = CONFIG.COMPANY_FORMS;
  const lines = text.split("\n").map(l => l.trim()).filter(Boolean);

  const candidates = [];

  for (const line of lines) {
    if (!forms.some(f => line.includes(f))) continue;

    // 行が長すぎるとノイズなので制限
    const clipped = line.length > 160 ? line.slice(0, 160) : line;

    // 法人格単体などは除外
    if (forms.includes(clipped) || clipped === "株式会社" || clipped === "合同会社") continue;

    const cleaned = clipped
      .replace(/[（(].*?[）)]/g, "")     // 括弧書き除去
      .replace(/\s+/g, " ")
      .replace(/[,:;。．、]+$/g, "");   // 行末句読点

    if (cleaned.length < 4) continue;

    candidates.push(cleaned);
  }

  return uniq_(candidates).slice(0, 50);
}

/**
 * ========= 抽出ロジック（住所） =========
 * - 都道府県起点の行を拾う
 * - 会社と住所は紐付けない
 */
function extractAddresses_(text) {
  const prefs = CONFIG.PREFS;
  const lines = text.split("\n").map(l => l.trim()).filter(Boolean);

  const addresses = [];

  for (const line of lines) {
    const pref = prefs.find(p => line.startsWith(p));
    if (!pref) continue;

    if (line.length < 6) continue;

    const cleaned = line
      .replace(/\s+/g, " ")
      .replace(/[,:;。．、]+$/g, "");

    addresses.push(cleaned.length > 240 ? cleaned.slice(0, 240) : cleaned);
  }

  return uniq_(addresses).slice(0, 100);
}

/**
 * ========= 台帳管理（処理済みfileId） =========
 * - D列（file_id）をSet化
 */
function loadProcessedFileIds_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return new Set();

  const fileIdCol = 4; // D列
  const values = sheet.getRange(2, fileIdCol, lastRow - 1, 1).getValues();

  const s = new Set();
  for (const [v] of values) {
    if (v) s.add(String(v));
  }
  return s;
}

/**
 * ========= シート初期ヘッダ =========
 * - 既に1行目が存在する場合は触らない（上書きしない）
 * - note 列でエラー/ヒットキーワード等を保持
 */
function ensureHeader_(sheet) {
  if (sheet.getLastRow() >= 1) return;

  sheet.appendRow([
    "processed_at",  // A
    "pdf_name",      // B
    "pdf_url",       // C
    "file_id",       // D
    "has_match",     // E
    "companies",     // F
    "addresses",     // G
    "note"           // H
  ]);
}

/**
 * ========= Configチェック =========
 */
function assertConfig_() {
  if (!CONFIG.SPREADSHEET_ID || CONFIG.SPREADSHEET_ID.indexOf("PUT_YOUR") === 0) {
    throw new Error("CONFIG.SPREADSHEET_ID を設定してください");
  }
  if (!CONFIG.FOLDER_ID || CONFIG.FOLDER_ID.indexOf("PUT_YOUR") === 0) {
    throw new Error("CONFIG.FOLDER_ID を設定してください");
  }
}

/**
 * ========= 便利関数 =========
 */
function getOrCreateSheet_(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function normalizeText_(t) {
  return String(t || "")
    .replace(/\r\n/g, "\n")
    .replace(/\r/g, "\n");
}

function uniq_(arr) {
  const s = new Set();
  const out = [];
  for (const x of arr) {
    const k = String(x);
    if (s.has(k)) continue;
    s.add(k);
    out.push(k);
  }
  return out;
}

/**
 * ========= デバッグ（任意） =========
 * 共有ドライブの fileId に対して Drive API が見えているか確認する。
 */
function debugDriveApiAccess() {
  const fileId = "PUT_TEST_FILE_ID_HERE";
  const meta = Drive.Files.get(fileId, { supportsAllDrives: true });
  Logger.log(JSON.stringify({
    id: meta.id,
    title: meta.title,
    mimeType: meta.mimeType,
    driveId: meta.driveId,
    parents: meta.parents
  }, null, 2));
}

/**
 * ========= 初回セットアップを一括で実行（任意） =========
 * - スプレッドシート初期化
 * - トリガ作成（必要なら）
 *
 * 使い方：
 *   setupAll(true) でトリガ作成まで一括
 *   setupAll(false) でシート初期化だけ
 */
function setupAll(createTrigger) {
  setupSpreadsheet();
  if (createTrigger) {
    resetDailyTrigger();
  }
}
