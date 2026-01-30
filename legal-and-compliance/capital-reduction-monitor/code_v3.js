/**
 * ============================================================
 * 官報 減資公告 抽出スクリプト v3
 *
 * 機能:
 * - 会社名抽出: 「取締役」「代表取締役」の前から法人格を含む会社名を抽出
 * - 住所抽出: 都道府県から丁目・番・号までを一括取得
 * - 減資情報: 減少額・減少後資本金を抽出
 * - ホームページURL: Google Custom Search API で会社HPを検索
 * - 社内通知メール: 会社名・住所・URLが揃った案件を社内に通知
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
  GOOGLE_API_KEY: "YOUR_GOOGLE_API_KEY",  // 例: "AIzaSy..."
  SEARCH_ENGINE_ID: "2364cb900fd03458f", // 例: "012345678901234567890:abcdefghijk"

  // URL検索を有効にするか
  ENABLE_URL_SEARCH: true,

  // ========= 社内通知メール設定 =========
  // 通知メールの送信先（社内メンバー）
  NOTIFICATION_EMAIL: "your-email@example.com",

  // 通知メールを有効にするか
  ENABLE_NOTIFICATION: true,

  // 営業メールテンプレート参照用リンク
  SALES_TEMPLATE_DOC: "https://docs.google.com/document/d/13K2ME4xZgCb87t8oS3D_97bUw8XXx7h6HDU9zVooF7I/edit",

  // 打ち合わせ予約リンク
  CALENDAR_LINK: "https://calendar.app.google/JKPmN9kUPModECmy6",

  // リンク集
  PROFILE_LINK: "https://linktr.ee/hodakagoto_martin"
};

/**
 * ========= 営業メールテンプレート =========
 */
const SALES_EMAIL_TEMPLATE = `ご担当者様

はじめまして。合同会社IntelligentBeastの後藤と申します。

AIを活用した業務効率化・自動化・コストカットを中心に、
実務に踏み込む形で外部支援を行っています。

御社の事業内容・体制を拝見し、

・業務の属人化
・ツール乱立による非効率
・AI活用が業務に接続できていない点

といった課題に対し、短期間で効果が出る支援が可能と考え、
ご連絡しました。

【主な支援内容（例）】
・GAS／Zapier等による業務自動化、工数削減
・SaaS／データの整理・統合による判断コスト削減
・生成AI（GPT等）の業務組み込みと運用設計

いずれも「ツール導入」ではなく、
業務設計とコスト構造の改善まで含めて対応します。

経歴・実績の詳細や考え方については、以下のリンク集をご参照ください。
${CONFIG.PROFILE_LINK}

まずは30分ほど、現状を伺うお時間をいただければ、
御社に合う形で具体案をご提示します。

▼ 打ち合わせ予約
${CONFIG.CALENDAR_LINK}

何卒よろしくお願いいたします。

合同会社IntelligentBeast
代表　後藤 穂高
your-email@example.com
080-2066-1791`;

/**
 * ========= スプレッドシートメニュー作成 =========
 * スプレッドシートを開いたときに自動でメニューを追加
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('減資公告ツール')
    .addItem('今すぐ実行（新規PDF処理）', 'runDaily')
    .addItem('通知状況を更新（未通知を通知）', 'notifyExistingCandidates')
    .addSeparator()
    .addItem('初期セットアップ', 'setupSpreadsheet')
    .addItem('日次トリガーを設定', 'setupDailyTrigger')
    .addItem('日次トリガーをリセット', 'resetDailyTrigger')
    .addItem('トリガー状態を確認', 'checkTriggerStatus')
    .addSeparator()
    .addItem('テストメール送信', 'testSalesNotificationEmail')
    .addToUi();
}

/**
 * ========= インストール可能トリガー用 onOpen =========
 * 編集権限のないユーザーでもメニューを表示するための関数
 */
function onOpenInstallable(e) {
  onOpen();
}

/**
 * ========= スプレッドシート初期設定 =========
 */
function setupSpreadsheet() {
  assertConfig_();

  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = getOrCreateSheet_(ss, CONFIG.SHEET_NAME);

  if (sheet.getLastRow() === 0) {
    ensureHeader_(sheet);
  }

  sheet.setFrozenRows(1);

  const lastCol = sheet.getLastColumn() || 13;
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
    sheet.setColumnWidth(11, 100); // notified
    sheet.setColumnWidth(12, 200); // note
  } catch (e) {}

  try {
    sheet.getRange(1, 1, 1, lastCol).setFontWeight("bold");
  } catch (e) {}
}

/**
 * ========= 日次トリガ設定 =========
 */
function setupDailyTrigger() {
  assertConfig_();

  // 既存のrunDailyトリガーがあるかチェック
  const existingTriggers = ScriptApp.getProjectTriggers();
  for (const t of existingTriggers) {
    if (t.getHandlerFunction() === "runDaily") {
      SpreadsheetApp.getUi().alert(
        'トリガー設定',
        '既に日次トリガーが設定されています。\n' +
        '再設定する場合は「日次トリガーをリセット」を使用してください。',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }
  }

  ScriptApp.newTrigger("runDaily")
    .timeBased()
    .everyDays(1)
    .atHour(CONFIG.TRIGGER_HOUR)
    .create();

  // onOpenのインストール可能トリガーも設定
  let hasOnOpenTrigger = false;
  for (const t of ScriptApp.getProjectTriggers()) {
    if (t.getHandlerFunction() === "onOpenInstallable") {
      hasOnOpenTrigger = true;
      break;
    }
  }

  if (!hasOnOpenTrigger) {
    ScriptApp.newTrigger("onOpenInstallable")
      .forSpreadsheet(CONFIG.SPREADSHEET_ID)
      .onOpen()
      .create();
  }

  SpreadsheetApp.getUi().alert(
    'トリガー設定完了',
    '日次トリガーを設定しました。\n' +
    '毎日 ' + CONFIG.TRIGGER_HOUR + '時 に自動実行されます。',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function resetDailyTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  let deletedCount = 0;

  for (const t of triggers) {
    if (t.getHandlerFunction() === "runDaily") {
      ScriptApp.deleteTrigger(t);
      deletedCount++;
    }
  }

  ScriptApp.newTrigger("runDaily")
    .timeBased()
    .everyDays(1)
    .atHour(CONFIG.TRIGGER_HOUR)
    .create();

  SpreadsheetApp.getUi().alert(
    'トリガーリセット完了',
    (deletedCount > 0 ? deletedCount + '個の既存トリガーを削除し、' : '') +
    '新しい日次トリガーを設定しました。\n' +
    '毎日 ' + CONFIG.TRIGGER_HOUR + '時 に自動実行されます。',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * ========= トリガー状態確認 =========
 */
function checkTriggerStatus() {
  const triggers = ScriptApp.getProjectTriggers();
  let message = '現在のトリガー一覧:\n\n';

  if (triggers.length === 0) {
    message += 'トリガーが設定されていません。\n「日次トリガーを設定」を実行してください。';
  } else {
    for (const t of triggers) {
      message += '・' + t.getHandlerFunction() + ' (' + t.getEventType() + ')\n';
    }
  }

  SpreadsheetApp.getUi().alert('トリガー状態', message, SpreadsheetApp.getUi().ButtonSet.OK);
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

  // 営業対象候補を収集
  const salesCandidates = [];

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
        "", "", "", "", "", "", "TEXT_EXTRACT_ERROR: " + String(e)
      ]);
      continue;
    }

    const hitKeywords = CONFIG.KEYWORDS.filter(kw => text.includes(kw));
    const hasMatch = hitKeywords.length > 0;

    if (!hasMatch) {
      rowsToAppend.push([
        nowIso, pdfName, pdfUrl, fileId, false,
        "", "", "", "", "", "", ""
      ]);
      continue;
    }

    const notices = splitIntoNotices_(text);

    if (notices.length === 0) {
      const companies = extractCompaniesV2_(text);
      const addresses = extractAddressesV2_(text);
      const maxLen = Math.max(companies.length, addresses.length, 1);

      for (let i = 0; i < maxLen; i++) {
        const companyName = companies[i] || "";
        const address = addresses[i] || "";
        const homepageUrl = (CONFIG.ENABLE_URL_SEARCH && companyName)
          ? searchCompanyHomepage_(companyName)
          : "";

        // 会社名・住所・URLが揃っているか判定
        const isComplete = companyName && address && homepageUrl;

        rowsToAppend.push([
          i === 0 ? nowIso : "",
          i === 0 ? pdfName : "",
          i === 0 ? pdfUrl : "",
          i === 0 ? fileId : "",
          i === 0 ? true : "",
          companyName,
          address,
          "",
          "",
          homepageUrl,
          isComplete ? "未通知" : "",
          i === 0 ? "hit_keywords=" + hitKeywords.join(",") : ""
        ]);

        // 営業対象候補に追加
        if (isComplete) {
          salesCandidates.push({
            companyName: companyName,
            address: address,
            homepageUrl: homepageUrl,
            capitalReduction: "",
            capitalAfter: "",
            pdfName: pdfName,
            pdfUrl: pdfUrl
          });
        }
      }
    } else {
      let isFirst = true;
      for (const notice of notices) {
        const extracted = extractFromNotice_(notice);

        const homepageUrl = (CONFIG.ENABLE_URL_SEARCH && extracted.companyName)
          ? searchCompanyHomepage_(extracted.companyName)
          : "";

        // 会社名・住所・URLが揃っているか判定
        const isComplete = extracted.companyName && extracted.address && homepageUrl;

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
          isComplete ? "未通知" : "",
          isFirst ? "hit_keywords=" + hitKeywords.join(",") : ""
        ]);

        // 営業対象候補に追加
        if (isComplete) {
          salesCandidates.push({
            companyName: extracted.companyName,
            address: extracted.address,
            homepageUrl: homepageUrl,
            capitalReduction: extracted.capitalReduction,
            capitalAfter: extracted.capitalAfter,
            pdfName: pdfName,
            pdfUrl: pdfUrl
          });
        }

        isFirst = false;
      }
    }
  }

  // スプレッドシートに追記
  if (rowsToAppend.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rowsToAppend.length, rowsToAppend[0].length)
      .setValues(rowsToAppend);
  }

  // 社内通知メールを送信
  if (CONFIG.ENABLE_NOTIFICATION && salesCandidates.length > 0) {
    sendSalesNotificationEmail_(salesCandidates);
  }
}

/**
 * ========= 社内通知メール送信 =========
 * 会社名・住所・homepage_urlが揃っている案件を社内に通知
 */
function sendSalesNotificationEmail_(candidates) {
  if (!candidates || candidates.length === 0) return;

  const subject = `【営業リード】減資公告から${candidates.length}件の営業候補を検出しました`;

  let body = `お疲れ様です。

減資公告スクリプトが以下の営業候補を検出しました。
会社名・住所・ホームページURLが揃っている案件です。

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
■ 営業候補一覧（${candidates.length}件）
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

`;

  for (let i = 0; i < candidates.length; i++) {
    const c = candidates[i];
    body += `【${i + 1}】${c.companyName}
────────────────────────────────────
  住所: ${c.address}
  HP: ${c.homepageUrl}
`;
    if (c.capitalReduction) {
      body += `  減少額: ${c.capitalReduction}
`;
    }
    if (c.capitalAfter) {
      body += `  減少後資本金: ${c.capitalAfter}
`;
    }
    body += `  官報PDF: ${c.pdfUrl}

`;
  }

  body += `━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
■ 営業メールテンプレート
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

以下のテンプレートを参考に、各社のHPを確認した上で営業メールを送信してください。

--- テンプレート開始 ---

${SALES_EMAIL_TEMPLATE}

--- テンプレート終了 ---

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
■ 営業時のポイント
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

1. 先方のHPを確認し、事業内容を把握する
2. 減資の理由を推測する（事業縮小？組織再編？）
3. 減資に伴う課題（業務効率化、コスト削減ニーズ）を想定してアプローチ
4. お問い合わせフォームまたは代表メールに送信

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

▼ テンプレート元ドキュメント
${CONFIG.SALES_TEMPLATE_DOC}

▼ スプレッドシート（全データ）
https://docs.google.com/spreadsheets/d/${CONFIG.SPREADSHEET_ID}/edit

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

※ このメールは減資公告抽出スクリプトから自動送信されています。
`;

  try {
    MailApp.sendEmail({
      to: CONFIG.NOTIFICATION_EMAIL,
      subject: subject,
      body: body
    });
    Logger.log("社内通知メール送信完了: " + candidates.length + "件");
  } catch (e) {
    Logger.log("社内通知メール送信エラー: " + String(e));
  }
}

/**
 * ========= 手動で社内通知メールをテスト送信 =========
 */
function testSalesNotificationEmail() {
  const testCandidates = [
    {
      companyName: "株式会社テスト",
      address: "東京都港区虎ノ門一丁目二番三号",
      homepageUrl: "https://example.com",
      capitalReduction: "1億円",
      capitalAfter: "5000万円",
      pdfName: "test.pdf",
      pdfUrl: "https://example.com/test.pdf"
    },
    {
      companyName: "合同会社サンプル",
      address: "大阪府大阪市北区梅田一丁目一番一号",
      homepageUrl: "https://sample.co.jp",
      capitalReduction: "",
      capitalAfter: "",
      pdfName: "sample.pdf",
      pdfUrl: "https://example.com/sample.pdf"
    }
  ];

  sendSalesNotificationEmail_(testCandidates);
  Logger.log("テストメール送信完了");
}

/**
 * ========= 既存データから営業候補を抽出して通知 =========
 * スプレッドシートの既存データをスキャンし、
 * 会社名・住所・URLが揃っていてまだ通知していない行を通知
 */
function notifyExistingCandidates() {
  assertConfig_();

  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) {
    Logger.log("シートが見つかりません: " + CONFIG.SHEET_NAME);
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    Logger.log("データがありません");
    return;
  }

  // データ取得（A2:L最終行）
  const data = sheet.getRange(2, 1, lastRow - 1, 12).getValues();
  const candidates = [];
  const rowsToUpdate = [];

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const companyName = row[5];  // F列: company_name
    const address = row[6];      // G列: address
    const homepageUrl = row[9];  // J列: homepage_url
    const notified = row[10];    // K列: notified

    // 会社名・住所・URLが揃っていて、まだ通知していない
    if (companyName && address && homepageUrl && notified !== "通知済") {
      candidates.push({
        companyName: companyName,
        address: address,
        homepageUrl: homepageUrl,
        capitalReduction: row[7] || "",
        capitalAfter: row[8] || "",
        pdfName: row[1] || "",
        pdfUrl: row[2] || ""
      });

      // 通知済みに更新する行番号を記録
      rowsToUpdate.push(i + 2); // 1-indexed + ヘッダ行
    }
  }

  if (candidates.length === 0) {
    Logger.log("通知対象の候補がありません");
    return;
  }

  // 通知メール送信
  sendSalesNotificationEmail_(candidates);

  // 通知済みに更新
  for (const rowNum of rowsToUpdate) {
    sheet.getRange(rowNum, 11).setValue("通知済"); // K列
  }

  Logger.log(candidates.length + "件を通知し、ステータスを更新しました");
}

/**
 * ========= 公告を分割 =========
 */
function splitIntoNotices_(text) {
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

  result.companyName = extractCompanyNameFromNotice_(noticeText);
  result.address = extractAddressFromNotice_(noticeText);

  const reductionMatch = noticeText.match(/(?:減少する(?:資本金の)?額|減少額)[^\d]*([0-9,，]+(?:万)?円)/);
  if (reductionMatch) {
    result.capitalReduction = normalizeAmount_(reductionMatch[1]);
  }

  const afterMatch = noticeText.match(/(?:減少後の?(?:資本金の)?額|効力発生後[^\d]*)[^\d]*([0-9,，]+(?:万)?円)/);
  if (afterMatch) {
    result.capitalAfter = normalizeAmount_(afterMatch[1]);
  }

  return result;
}

/**
 * ========= 会社名抽出（改善版） =========
 */
function extractCompanyNameFromNotice_(text) {
  const forms = CONFIG.COMPANY_FORMS;
  const formsPattern = forms.join('|');

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

  const tokuteiMatch = text.match(/((?:株式会社|合同会社)[^\s　]{1,20}特定目的会社)/);
  if (tokuteiMatch) {
    return cleanCompanyName_(tokuteiMatch[1]);
  }

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
 */
function extractAddressFromNotice_(text) {
  const prefs = CONFIG.PREFS;
  const prefsPattern = prefs.join('|');

  const addressPattern = new RegExp(
    `((?:${prefsPattern})[^\\n]{5,80}?(?:` +
    `(?:\\d+|[一二三四五六七八九十]+)(?:丁目|番(?:地|町)?|号)|` +
    `(?:ビル|タワー|センター|マンション|アパート)[^\\n]{0,10}` +
    `))`,
    'g'
  );

  const matches = text.match(addressPattern);
  if (matches && matches.length > 0) {
    let bestMatch = matches[0];
    for (const m of matches) {
      if (m.length > bestMatch.length) {
        bestMatch = m;
      }
    }
    return cleanAddress_(bestMatch);
  }

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
 * ========= 会社名抽出（フォールバック用） =========
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
 * ========= 住所抽出（フォールバック用） =========
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
    "company_name",      // F
    "address",           // G
    "capital_reduction", // H
    "capital_after",     // I
    "homepage_url",      // J
    "notified",          // K（通知ステータス）
    "note"               // L
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
 * ========= 初回セットアップ =========
 */
function setupAll(createTrigger) {
  setupSpreadsheet();
  if (createTrigger) {
    resetDailyTrigger();
  }
}