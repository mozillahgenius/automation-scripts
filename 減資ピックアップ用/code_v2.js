/**
 * ============================================================
 * 官報 減資公告 抽出スクリプト v2
 *
 * 改善点:
 * - 会社名抽出: 「取締役」「代表取締役」の前から法人格を含む会社名を抽出
 * - 住所抽出: 都道府県から丁目・番・号までを一括取得
 * - 減資情報: 減少額・減少後資本金を抽出
 * - ホームページURL: Google Custom Search API で会社HPを検索
 *
 * 重要:
 * - Advanced Google Services の「Drive API」を有効化
 * - Google Custom Search API を使用する場合は API_KEY と SEARCH_ENGINE_ID を設定
 * ============================================================
 */

const CONFIG = {
  // 出力先スプレッドシートID
  SPREADSHEET_ID: "1f036RjFfzHyHQTLD1lM_73qKDQnZDXdeUXYC9RVce-w",

  // 出力先シート名
  SHEET_NAME: "減資抽出",

  // 共有ドライブ内の対象フォルダID
  FOLDER_ID: "1IVQTgWiMIftmc9ygO48WqWJsEGAoGEII",

  // キーワード（OR条件）
  KEYWORDS: [
    "減資",
    "資本金の額の減少",
    "準備金の額の減少"
  ],

  // 会社名抽出用の法人格
  COMPANY_FORMS: ["株式会社", "合同会社", "有限会社", "合名会社", "合資会社"],

  // 都道府県
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

  // トリガ実行時刻
  TRIGGER_HOUR: 2,

  // Google Custom Search API 設定（URL検索用）
  // ※ 使用する場合は設定が必要
  // Google Cloud Console で Custom Search API を有効化し、APIキーを取得
  // https://programmablesearchengine.google.com/ で検索エンジンを作成
  GOOGLE_API_KEY: "",  // 例: "AIzaSy..."
  SEARCH_ENGINE_ID: "", // 例: "012345678901234567890:abcdefghijk"

  // URL検索を有効にするか（APIキー未設定の場合はfalseに）
  ENABLE_URL_SEARCH: false
};

/**
 * ========= スプレッドシート初期設定 =========
 */
function setupSpreadsheet() {
  assertConfig_();

  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = getOrCreateSheet_(ss, CONFIG.SHEET_NAME);

  // 既存データがある場合はヘッダを作り直さない
  if (sheet.getLastRow() === 0) {
    ensureHeader_(sheet);
  }

  sheet.setFrozenRows(1);

  // フィルタ
  const lastCol = sheet.getLastColumn() || 12;
  const lastRow = sheet.getLastRow() || 1;
  const range = sheet.getRange(1, 1, Math.max(1, lastRow), lastCol);
  try {
    if (!sheet.getFilter()) range.createFilter();
  } catch (e) {}

  // 列幅
  try {
    sheet.setColumnWidth(1, 170);  // processed_at
    sheet.setColumnWidth(2, 250);  // pdf_name
    sheet.setColumnWidth(3, 280);  // pdf_url
    sheet.setColumnWidth(4, 220);  // file_id
    sheet.setColumnWidth(5, 80);   // has_match
    sheet.setColumnWidth(6, 280);  // company_name
    sheet.setColumnWidth(7, 350);  // address
    sheet.setColumnWidth(8, 120);  // capital_reduction
    sheet.setColumnWidth(9, 120);  // capital_after
    sheet.setColumnWidth(10, 350); // homepage_url
    sheet.setColumnWidth(11, 200); // note
  } catch (e) {}

  // ヘッダ装飾
  try {
    sheet.getRange(1, 1, 1, lastCol).setFontWeight("bold");
  } catch (e) {}
}

/**
 * ========= 日次トリガ設定 =========
 */
function setupDailyTrigger() {
  assertConfig_();
  ScriptApp.newTrigger("runDaily")
    .timeBased()
    .everyDays(1)
    .atHour(CONFIG.TRIGGER_HOUR)
    .create();
}

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
 * ========= メイン処理 =========
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

    if (processed.has(fileId)) continue;

    const pdfName = file.getName();
    const pdfUrl = file.getUrl();

    let text = "";
    try {
      text = extractTextFromPdfViaDocs_(fileId);
    } catch (e) {
      rowsToAppend.push([
        nowIso, pdfName, pdfUrl, fileId, false,
        "", "", "", "", "", "TEXT_EXTRACT_ERROR: " + String(e)
      ]);
      continue;
    }

    const hitKeywords = CONFIG.KEYWORDS.filter(kw => text.includes(kw));
    const hasMatch = hitKeywords.length > 0;

    if (!hasMatch) {
      rowsToAppend.push([
        nowIso, pdfName, pdfUrl, fileId, false,
        "", "", "", "", "", ""
      ]);
      continue;
    }

    // 公告単位で分割して処理
    const notices = splitIntoNotices_(text);

    if (notices.length === 0) {
      // 分割できない場合は従来方式
      const companies = extractCompaniesV2_(text);
      const addresses = extractAddressesV2_(text);
      const maxLen = Math.max(companies.length, addresses.length, 1);

      for (let i = 0; i < maxLen; i++) {
        const companyName = companies[i] || "";
        const homepageUrl = (CONFIG.ENABLE_URL_SEARCH && companyName)
          ? searchCompanyHomepage_(companyName)
          : "";

        rowsToAppend.push([
          i === 0 ? nowIso : "",
          i === 0 ? pdfName : "",
          i === 0 ? pdfUrl : "",
          i === 0 ? fileId : "",
          i === 0 ? true : "",
          companyName,
          addresses[i] || "",
          "",
          "",
          homepageUrl,
          i === 0 ? "hit_keywords=" + hitKeywords.join(",") : ""
        ]);
      }
    } else {
      // 公告単位で処理
      let isFirst = true;
      for (const notice of notices) {
        const extracted = extractFromNotice_(notice);

        const homepageUrl = (CONFIG.ENABLE_URL_SEARCH && extracted.companyName)
          ? searchCompanyHomepage_(extracted.companyName)
          : "";

        rowsToAppend.push([
          isFirst ? nowIso : "",
          isFirst ? pdfName : "",
          isFirst ? pdfUrl : "",
          isFirst ? fileId : "",
          isFirst ? true : "",
          extracted.companyName,
          extracted.address,
          extracted.capitalReduction,
          extracted.capitalAfter,
          homepageUrl,
          isFirst ? "hit_keywords=" + hitKeywords.join(",") : ""
        ]);
        isFirst = false;
      }
    }
  }

  if (rowsToAppend.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rowsToAppend.length, rowsToAppend[0].length)
      .setValues(rowsToAppend);
  }
}

/**
 * ========= 公告を分割 =========
 * 官報の減資公告は「優先出資」「資本金の額の減少公告」などの見出しで始まる
 */
function splitIntoNotices_(text) {
  // 公告の区切りパターン
  const patterns = [
    /優先出資[^\n]*公告/g,
    /資本金の額の減少[^\n]*公告/g,
    /準備金の額の減少[^\n]*公告/g,
    /減資[^\n]*公告/g
  ];

  const notices = [];
  const lines = text.split('\n');
  let currentNotice = [];
  let inNotice = false;

  for (const line of lines) {
    const trimmed = line.trim();

    // 新しい公告の開始を検出
    const isStart = patterns.some(p => {
      p.lastIndex = 0;
      return p.test(trimmed);
    });

    if (isStart) {
      if (currentNotice.length > 0) {
        notices.push(currentNotice.join('\n'));
      }
      currentNotice = [trimmed];
      inNotice = true;
    } else if (inNotice) {
      currentNotice.push(trimmed);
    }
  }

  if (currentNotice.length > 0) {
    notices.push(currentNotice.join('\n'));
  }

  return notices;
}

/**
 * ========= 1つの公告から情報を抽出 =========
 */
function extractFromNotice_(noticeText) {
  const result = {
    companyName: "",
    address: "",
    capitalReduction: "",
    capitalAfter: ""
  };

  // 会社名抽出（改善版）
  result.companyName = extractCompanyNameFromNotice_(noticeText);

  // 住所抽出（改善版）
  result.address = extractAddressFromNotice_(noticeText);

  // 減資額抽出
  const reductionMatch = noticeText.match(/(?:減少する(?:資本金の)?額|減少額)[^\d]*([0-9,，]+(?:万)?円)/);
  if (reductionMatch) {
    result.capitalReduction = normalizeAmount_(reductionMatch[1]);
  }

  // 減少後資本金額
  const afterMatch = noticeText.match(/(?:減少後の?(?:資本金の)?額|効力発生後[^\d]*)[^\d]*([0-9,，]+(?:万)?円)/);
  if (afterMatch) {
    result.capitalAfter = normalizeAmount_(afterMatch[1]);
  }

  return result;
}

/**
 * ========= 会社名抽出（改善版） =========
 * 「取締役」「代表取締役」の前にある法人格を含む会社名を抽出
 */
function extractCompanyNameFromNotice_(text) {
  const forms = CONFIG.COMPANY_FORMS;
  const formsPattern = forms.join('|');

  // パターン1: 「取締役」「代表取締役」の前にある会社名
  // 例: "株式会社ゴールドC　取締役　松澤和浩"
  const directorPatterns = [
    new RegExp(`((?:${formsPattern})[^\\s　取代]{1,30})\\s*(?:代表)?取締役`, 'g'),
    new RegExp(`((?:${formsPattern})[^\\n]{1,30})\\n\\s*(?:代表)?取締役`, 'g')
  ];

  for (const pattern of directorPatterns) {
    const match = pattern.exec(text);
    if (match) {
      return cleanCompanyName_(match[1]);
    }
  }

  // パターン2: 「特定目的会社」を含むパターン
  const tokuteiMatch = text.match(/((?:株式会社|合同会社)[^\s　]{1,20}特定目的会社)/);
  if (tokuteiMatch) {
    return cleanCompanyName_(tokuteiMatch[1]);
  }

  // パターン3: 1行目付近の法人格を含む文字列
  const lines = text.split('\n').slice(0, 10);
  for (const line of lines) {
    for (const form of forms) {
      if (line.includes(form)) {
        const companyMatch = line.match(new RegExp(`(${form}[^\\s　、。（(]{1,30})`));
        if (companyMatch) {
          return cleanCompanyName_(companyMatch[1]);
        }
      }
    }
  }

  return "";
}

/**
 * ========= 住所抽出（改善版） =========
 * 都道府県から丁目・番・号または建物名までを一括取得
 */
function extractAddressFromNotice_(text) {
  const prefs = CONFIG.PREFS;
  const prefsPattern = prefs.join('|');

  // 住所パターン: 都道府県から番地・建物名まで
  // 例: "東京都港区虎ノ門一丁目二番六号"
  const addressPattern = new RegExp(
    `((?:${prefsPattern})[^\\n]{5,80}?(?:` +
    `(?:\\d+|[一二三四五六七八九十]+)(?:丁目|番(?:地|町)?|号)|` +
    `(?:ビル|タワー|センター|マンション|アパート)[^\\n]{0,10}` +
    `))`,
    'g'
  );

  const matches = text.match(addressPattern);
  if (matches && matches.length > 0) {
    // 最も長いマッチを採用（より詳細な住所）
    let bestMatch = matches[0];
    for (const m of matches) {
      if (m.length > bestMatch.length) {
        bestMatch = m;
      }
    }
    return cleanAddress_(bestMatch);
  }

  // フォールバック: 従来方式
  for (const pref of prefs) {
    const idx = text.indexOf(pref);
    if (idx >= 0) {
      const substr = text.substring(idx, idx + 100);
      const endMatch = substr.match(/^(.+?(?:(?:\d+|[一二三四五六七八九十]+)(?:丁目|番|号)|[^\s]{2,10}(?:ビル|タワー)))/);
      if (endMatch) {
        return cleanAddress_(endMatch[1]);
      }
    }
  }

  return "";
}

/**
 * ========= 会社名クリーニング =========
 */
function cleanCompanyName_(name) {
  return name
    .replace(/[\s　]+/g, '')
    .replace(/[（(].*?[）)]/g, '')
    .replace(/[、。，．,:;]+$/g, '')
    .replace(/取締役.*$/, '')
    .replace(/代表.*$/, '')
    .trim();
}

/**
 * ========= 住所クリーニング =========
 */
function cleanAddress_(addr) {
  return addr
    .replace(/[\s　]+/g, '')
    .replace(/[、。，．]+$/g, '')
    .trim();
}

/**
 * ========= 金額正規化 =========
 */
function normalizeAmount_(amount) {
  return amount
    .replace(/[，,]/g, '')
    .replace(/\s/g, '');
}

/**
 * ========= 会社名抽出（フォールバック用・従来版改良） =========
 */
function extractCompaniesV2_(text) {
  const forms = CONFIG.COMPANY_FORMS;
  const formsPattern = forms.join('|');

  const pattern = new RegExp(
    `((?:${formsPattern})[^\\s　、。（(取代]{1,40})`,
    'g'
  );

  const matches = [];
  let m;
  while ((m = pattern.exec(text)) !== null) {
    const cleaned = cleanCompanyName_(m[1]);
    if (cleaned.length >= 4 && !forms.includes(cleaned)) {
      matches.push(cleaned);
    }
  }

  return uniq_(matches).slice(0, 50);
}

/**
 * ========= 住所抽出（フォールバック用・従来版改良） =========
 */
function extractAddressesV2_(text) {
  const prefs = CONFIG.PREFS;
  const prefsPattern = prefs.join('|');

  const pattern = new RegExp(
    `((?:${prefsPattern})[^\\n]{5,80}?(?:(?:\\d+|[一二三四五六七八九十]+)(?:丁目|番|号)|ビル|タワー))`,
    'g'
  );

  const matches = [];
  let m;
  while ((m = pattern.exec(text)) !== null) {
    const cleaned = cleanAddress_(m[1]);
    if (cleaned.length >= 8) {
      matches.push(cleaned);
    }
  }

  return uniq_(matches).slice(0, 100);
}

/**
 * ========= ホームページURL検索 =========
 * Google Custom Search API を使用
 */
function searchCompanyHomepage_(companyName) {
  if (!CONFIG.GOOGLE_API_KEY || !CONFIG.SEARCH_ENGINE_ID) {
    return "";
  }

  try {
    const query = encodeURIComponent(companyName + " 公式サイト");
    const url = `https://www.googleapis.com/customsearch/v1?key=${CONFIG.GOOGLE_API_KEY}&cx=${CONFIG.SEARCH_ENGINE_ID}&q=${query}&num=1`;

    const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const json = JSON.parse(response.getContentText());

    if (json.items && json.items.length > 0) {
      const link = json.items[0].link;
      // SNSや求人サイトは除外
      const excludePatterns = [
        /facebook\.com/i,
        /twitter\.com/i,
        /linkedin\.com/i,
        /instagram\.com/i,
        /youtube\.com/i,
        /indeed\.com/i,
        /recruit/i,
        /wikipedia/i
      ];

      if (!excludePatterns.some(p => p.test(link))) {
        return link;
      }

      // 2番目の結果を試す
      if (json.items.length > 1) {
        return json.items[1].link;
      }
    }
  } catch (e) {
    Logger.log("URL検索エラー: " + companyName + " - " + String(e));
  }

  return "";
}

/**
 * ========= 代替: UrlFetchApp で簡易検索（API不要版） =========
 * Google Custom Search API が使えない場合の代替
 * ※ 精度は劣る
 */
function searchCompanyHomepageSimple_(companyName) {
  try {
    // DuckDuckGo の Instant Answer API（制限あり）
    const query = encodeURIComponent(companyName);
    const url = `https://api.duckduckgo.com/?q=${query}&format=json&no_html=1`;

    const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const json = JSON.parse(response.getContentText());

    if (json.AbstractURL) {
      return json.AbstractURL;
    }

    if (json.Results && json.Results.length > 0) {
      return json.Results[0].FirstURL;
    }
  } catch (e) {
    Logger.log("簡易URL検索エラー: " + companyName + " - " + String(e));
  }

  return "";
}

/**
 * ========= テキスト抽出 =========
 */
function extractTextFromPdfViaDocs_(pdfFileId) {
  Drive.Files.get(pdfFileId, { supportsAllDrives: true });

  const resource = {
    title: "tmp_pdf_extract_" + pdfFileId,
    mimeType: MimeType.GOOGLE_DOCS
  };

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
    try { DriveApp.getFileById(docId).setTrashed(true); } catch (e) {}
  }
}

/**
 * ========= 台帳管理 =========
 */
function loadProcessedFileIds_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return new Set();

  const fileIdCol = 4;
  const values = sheet.getRange(2, fileIdCol, lastRow - 1, 1).getValues();

  const s = new Set();
  for (const [v] of values) {
    if (v) s.add(String(v));
  }
  return s;
}

/**
 * ========= ヘッダ =========
 */
function ensureHeader_(sheet) {
  if (sheet.getLastRow() >= 1) return;

  sheet.appendRow([
    "processed_at",      // A
    "pdf_name",          // B
    "pdf_url",           // C
    "file_id",           // D
    "has_match",         // E
    "company_name",      // F（会社名）
    "address",           // G（住所）
    "capital_reduction", // H（減少額）
    "capital_after",     // I（減少後資本金）
    "homepage_url",      // J（ホームページURL）
    "note"               // K
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
 * ========= デバッグ =========
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
 * ========= URL検索テスト =========
 */
function testUrlSearch() {
  const testCompanies = [
    "株式会社ゴールドC",
    "Luna3特定目的会社"
  ];

  for (const company of testCompanies) {
    const url = searchCompanyHomepage_(company);
    Logger.log(company + " => " + url);
  }
}

/**
 * ========= 抽出テスト（サンプルテキスト） =========
 */
function testExtraction() {
  const sampleText = `
優先出資証券提出公告
当社は、発行済優先出資証券につき当社の優先出資証券を消却する
東京都港区虎ノ門二丁目二番一号
取締役　松澤　和浩

資本金の額の減少公告
株式会社ゴールドC
東京都港区虎ノ門一丁目二番六号
減少する額　三億円
効力発生後の資本金の額　一億円
取締役　松澤　和浩
  `;

  const notices = splitIntoNotices_(sampleText);
  Logger.log("公告数: " + notices.length);

  for (let i = 0; i < notices.length; i++) {
    const result = extractFromNotice_(notices[i]);
    Logger.log("公告 " + (i + 1) + ":");
    Logger.log("  会社名: " + result.companyName);
    Logger.log("  住所: " + result.address);
    Logger.log("  減少額: " + result.capitalReduction);
    Logger.log("  減少後: " + result.capitalAfter);
  }
}

/**
 * ========= 初回セットアップ =========
 */
function setupAll(createTrigger) {
  setupSpreadsheet();
  if (createTrigger) {
    resetDailyTrigger();
  }
}
