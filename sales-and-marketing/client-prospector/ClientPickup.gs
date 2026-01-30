/**
 * 情シス・法務 業務委託クライアント自動ピックアップシステム
 *
 * 【GASデータソース】
 * - PR Times RSS: 調達・IPO・業務効率化ニュース
 * - Indeed RSS: 求人情報から企業を検知
 * - Google News API（カスタム検索）: 炎上・ガバナンス関連ニュース
 * - 手動インポート: CSVデータ
 *
 * 【Zapier連携データソース】
 * - Wantedly/Green: Web Parserでスクレイピング → Webhook送信
 * - X（Twitter）: 炎上・ネガティブ投稿監視 → Webhook送信
 * - Google Alerts: ニュース監視 → Webhook送信
 *
 * 【Zapier連携機能】
 * - 高スコア企業のSlack/メール通知
 * - HubSpot/Salesforce CRM連携
 */

// ========================================
// 設定
// ========================================
const CONFIG = {
  // スプレッドシートID（★要設定★）
  SPREADSHEET_ID: 'SPREADSHEET_ID_PLACEHOLDER',

  // シート名
  SHEETS: {
    DAILY_PICKUP: 'Daily_Pickup',
    COMPANIES: 'Companies_Master',
    SIGNALS: 'Signal_Log',
    SETTINGS: 'Settings'
  },

  // スコア閾値
  SCORE_THRESHOLD: 0.5,

  // 抽出件数
  DAILY_LIMIT: 20,

  // 対象従業員規模
  MIN_EMPLOYEES: 10,
  MAX_EMPLOYEES: 100,

  // 対象業界
  TARGET_INDUSTRIES: ['SaaS', 'IT', '金融', 'フィンテック', 'エンタメ', 'ゲーム', 'スタートアップ'],

  // 除外キーワード
  EXCLUDE_KEYWORDS: ['SIer', 'SES', '受託開発', '派遣', '人材紹介'],

  // Google Custom Search API（オプション）
  GOOGLE_API_KEY: '',
  GOOGLE_CSE_ID: '',

  // メール通知設定
  EMAIL: {
    // 通知先メールアドレス
    NOTIFICATION_TO: 'your-email@example.com',

    // 通知閾値（このスコア以上で通知）
    NOTIFICATION_SCORE_THRESHOLD: 0.7,

    // 送信者名
    SENDER_NAME: 'クライアントピックアップBot'
  },

  // Zapier連携設定（オプション）
  ZAPIER: {
    // Zapierからのデータ受信用シート
    INCOMING_SHEET: 'Zapier_Incoming',

    // CRM連携用Webhook URL（オプション）
    CRM_WEBHOOK: ''
  }
};

// ========================================
// スコアリング重み（仕様書準拠）
// ========================================
const WEIGHTS = {
  COST_CUT: 0.25,      // コストカットシグナル
  GROWTH: 0.20,        // 急成長シグナル
  GOVERNANCE: 0.20,    // ガバナンスシグナル
  OPS_DEBT: 0.15,      // 運用負債シグナル
  FIRE: 0.20,          // 炎上シグナル
  NOISE_PENALTY: 0.30  // ノイズペナルティ
};

// ========================================
// シグナルキーワード定義
// ========================================
const SIGNAL_KEYWORDS = {
  // A. コストカット（E4）
  COST_CUT: {
    strong: ['業務効率化', 'コスト削減', 'AI導入', '自動化推進', 'DX推進', 'バックオフィス改革'],
    medium: ['SaaS統廃合', '内製化', '外注切替', 'ツール見直し', '業務改善'],
    weak: ['効率化', '省力化', '生産性向上']
  },

  // B. 急成長（E1）
  GROWTH: {
    strong: ['シリーズA', 'シリーズB', 'シリーズC', '資金調達', '○億円調達'],
    medium: ['急成長', '事業拡大', '採用強化', 'エンジニア募集', '組織拡大'],
    weak: ['成長中', '拡大中', '積極採用']
  },

  // C. 監査・統制（E2）
  GOVERNANCE: {
    strong: ['IPO準備', 'ISMS取得', 'SOC2', 'Pマーク', '内部監査'],
    medium: ['コンプライアンス強化', '情報セキュリティ', '監査対応', '規程整備'],
    weak: ['セキュリティ', '内部統制', 'ガバナンス']
  },

  // D. 情シス・法務負債
  OPS_DEBT: {
    strong: ['情シス募集', '法務募集', '管理部門立ち上げ', '一人情シス', '専任不在'],
    medium: ['契約管理', 'ID管理', 'アカウント管理', '英文契約', '海外取引'],
    weak: ['バックオフィス', '管理部門', '総務']
  },

  // E. 炎上・レピュテーションリスク
  FIRE: {
    strong: ['情報漏えい', '個人情報流出', '不正アクセス', 'セキュリティインシデント', '炎上'],
    medium: ['障害発生', 'サービス停止', 'クレーム', '対応遅延', 'バグ'],
    weak: ['トラブル', '不具合', '問題発生']
  },

  // ノイズ（除外対象）
  NOISE: ['SIer', 'SES', '受託開発', '派遣', '人材紹介', 'コンサルティングファーム']
};

// ========================================
// メイン実行関数
// ========================================

/**
 * 日次実行：クライアントピックアップ
 */
function dailyPickup() {
  console.log('=== 日次ピックアップ開始 ===');

  try {
    // 1. データ収集
    const rawData = collectData();
    console.log(`収集データ数: ${rawData.length}`);

    // 2. 正規化・重複排除
    const companies = normalizeAndDedupe(rawData);
    console.log(`正規化後: ${companies.length}`);

    // 3. フィルタリング（規模・業界）
    const filtered = filterCompanies(companies);
    console.log(`フィルタ後: ${filtered.length}`);

    // 4. スコアリング
    const scored = scoreCompanies(filtered);
    console.log(`スコアリング完了`);

    // 5. 上位抽出
    const topCompanies = extractTopCompanies(scored, CONFIG.DAILY_LIMIT);
    console.log(`抽出数: ${topCompanies.length}`);

    // 6. シートに出力
    outputToSheet(topCompanies);

    // 7. 高スコア企業をZapier経由で通知
    notifyHighScoreCompanies(topCompanies);

    console.log('=== 日次ピックアップ完了 ===');
    return topCompanies;

  } catch (error) {
    console.error('エラー発生:', error);
    throw error;
  }
}

// ========================================
// データ収集
// ========================================

/**
 * 各ソースからデータを収集
 */
function collectData() {
  const allData = [];

  // PR Times RSS
  const prTimesData = fetchPRTimesRSS();
  allData.push(...prTimesData);

  // Indeed RSS（求人情報）
  const indeedData = fetchIndeedRSS();
  allData.push(...indeedData);

  // Zapierからの受信データ（Wantedly/Green/Twitter等）
  const zapierData = fetchZapierIncoming();
  allData.push(...zapierData);

  // 手動インポートシートからのデータ
  const manualData = fetchManualImport();
  allData.push(...manualData);

  // Google News検索（APIキー設定時のみ）
  if (CONFIG.GOOGLE_API_KEY && CONFIG.GOOGLE_CSE_ID) {
    const newsData = fetchGoogleNews();
    allData.push(...newsData);
  }

  return allData;
}

/**
 * PR Times RSSからデータ取得
 */
function fetchPRTimesRSS() {
  const results = [];
  const searchTerms = ['資金調達', 'IPO準備', '業務効率化', 'DX推進', 'セキュリティ'];

  searchTerms.forEach(term => {
    try {
      const url = `https://prtimes.jp/index.php?word=${encodeURIComponent(term)}&latest=1&aession=&articlecount=30&type=rss`;
      const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });

      if (response.getResponseCode() === 200) {
        const xml = XmlService.parse(response.getContentText());
        const items = xml.getRootElement().getChild('channel').getChildren('item');

        items.forEach(item => {
          const title = item.getChildText('title') || '';
          const description = item.getChildText('description') || '';
          const link = item.getChildText('link') || '';
          const pubDate = item.getChildText('pubDate') || '';

          // 企業名を抽出（タイトルから推測）
          const companyName = extractCompanyName(title);

          if (companyName) {
            results.push({
              source: 'PR_TIMES',
              companyName: companyName,
              title: title,
              description: description,
              url: link,
              date: pubDate,
              searchTerm: term
            });
          }
        });
      }
    } catch (e) {
      console.log(`PR Times取得エラー (${term}):`, e.message);
    }
  });

  return results;
}

/**
 * Indeed RSSから求人データ取得
 *
 * Indeed RSSフィード形式:
 * https://jp.indeed.com/rss?q=検索キーワード&l=勤務地
 *
 * 情シス・法務関連の求人を検索し、
 * 求人を出している企業をピックアップする
 */
function fetchIndeedRSS() {
  const results = [];

  // 検索クエリ（情シス・法務関連の求人キーワード）
  const searchQueries = [
    // 情シス系
    { q: '情シス', signal: 'OPS_DEBT' },
    { q: '社内SE', signal: 'OPS_DEBT' },
    { q: '一人情シス', signal: 'OPS_DEBT' },
    { q: 'コーポレートIT', signal: 'OPS_DEBT' },
    { q: 'IT管理者', signal: 'OPS_DEBT' },

    // 法務系
    { q: '法務', signal: 'GOVERNANCE' },
    { q: '契約法務', signal: 'GOVERNANCE' },
    { q: 'コンプライアンス', signal: 'GOVERNANCE' },

    // 成長シグナル系
    { q: 'IPO準備 管理部門', signal: 'GOVERNANCE' },
    { q: 'スタートアップ 情シス', signal: 'GROWTH' },
    { q: '急成長 バックオフィス', signal: 'GROWTH' },

    // セキュリティ系
    { q: 'ISMS', signal: 'GOVERNANCE' },
    { q: '情報セキュリティ 担当', signal: 'GOVERNANCE' }
  ];

  searchQueries.forEach(query => {
    try {
      // Indeed RSS URL（日本版）
      const url = `https://jp.indeed.com/rss?q=${encodeURIComponent(query.q)}&sort=date&limit=25`;

      const response = UrlFetchApp.fetch(url, {
        muteHttpExceptions: true,
        headers: {
          'User-Agent': 'Mozilla/5.0 (compatible; GoogleAppsScript)'
        }
      });

      if (response.getResponseCode() === 200) {
        const content = response.getContentText();

        // XMLパース
        try {
          const xml = XmlService.parse(content);
          const root = xml.getRootElement();
          const channel = root.getChild('channel');

          if (channel) {
            const items = channel.getChildren('item');

            items.forEach(item => {
              const title = item.getChildText('title') || '';
              const description = item.getChildText('description') || '';
              const link = item.getChildText('link') || '';
              const pubDate = item.getChildText('pubDate') || '';

              // 企業名を抽出
              const companyName = extractCompanyFromIndeed(title, description);

              // 除外チェック（派遣・SES等）
              if (companyName && !isExcludedCompany(title + ' ' + description)) {
                // 従業員数の推定（求人文から）
                const employeeEstimate = estimateEmployeeCount(description);

                results.push({
                  source: 'INDEED',
                  companyName: companyName,
                  title: title,
                  description: description,
                  url: link,
                  date: pubDate,
                  searchTerm: query.q,
                  signalType: query.signal,
                  employees: employeeEstimate
                });
              }
            });
          }
        } catch (parseError) {
          console.log(`Indeed XMLパースエラー (${query.q}):`, parseError.message);
        }
      }

      // レート制限対策（1秒待機）
      Utilities.sleep(1000);

    } catch (e) {
      console.log(`Indeed取得エラー (${query.q}):`, e.message);
    }
  });

  console.log(`Indeed RSS: ${results.length}件取得`);
  return results;
}

/**
 * Indeed求人から企業名を抽出
 */
function extractCompanyFromIndeed(title, description) {
  // Indeedの求人タイトル形式: "職種名 - 企業名 - 勤務地"
  const titleMatch = title.match(/\s-\s([^-]+)\s-\s/);
  if (titleMatch) {
    return titleMatch[1].trim();
  }

  // descriptionから企業名を抽出
  const companyPatterns = [
    /会社名[：:]\s*([^\n<]+)/,
    /企業名[：:]\s*([^\n<]+)/,
    /株式会社([^\s、。,\.<]+)/,
    /([^\s、。,\.<]+)株式会社/,
    /合同会社([^\s、。,\.<]+)/
  ];

  for (const pattern of companyPatterns) {
    const match = description.match(pattern);
    if (match) {
      return match[0].includes('株式会社') || match[0].includes('合同会社')
        ? match[0]
        : match[1].trim();
    }
  }

  return null;
}

/**
 * 除外対象企業かチェック
 */
function isExcludedCompany(text) {
  const excludePatterns = [
    /派遣/,
    /SES/,
    /受託開発/,
    /人材紹介/,
    /エージェント/,
    /求人サイト/,
    /転職支援/,
    /紹介予定派遣/
  ];

  return excludePatterns.some(pattern => pattern.test(text));
}

/**
 * 求人文から従業員数を推定
 */
function estimateEmployeeCount(description) {
  // 従業員数の記載パターン
  const patterns = [
    /従業員[：:数]?\s*約?(\d+)[名人]/,
    /社員数[：:]\s*約?(\d+)[名人]/,
    /(\d+)[名人]規模/,
    /(\d+)[-〜~](\d+)[名人]/
  ];

  for (const pattern of patterns) {
    const match = description.match(pattern);
    if (match) {
      // 範囲指定の場合は中間値
      if (match[2]) {
        return Math.round((parseInt(match[1]) + parseInt(match[2])) / 2);
      }
      return parseInt(match[1]);
    }
  }

  return 0; // 不明
}

// ========================================
// Zapier連携機能
// ========================================

/**
 * Zapierからのデータ受信用Webhook（GASをWebアプリとしてデプロイ）
 *
 * ZapierのWebhook送信先としてこのURLを設定:
 * https://script.google.com/macros/s/[DEPLOYMENT_ID]/exec
 *
 * 送信データ形式:
 * {
 *   "source": "WANTEDLY" | "GREEN" | "TWITTER" | "GOOGLE_ALERTS",
 *   "companyName": "株式会社〇〇",
 *   "title": "求人タイトル / ツイート内容",
 *   "description": "詳細",
 *   "url": "https://...",
 *   "industry": "SaaS",
 *   "employees": 50,
 *   "signalType": "GROWTH" | "FIRE" | "OPS_DEBT" | "GOVERNANCE"
 * }
 */
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    // スプレッドシートに保存
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    let sheet = ss.getSheetByName(CONFIG.ZAPIER.INCOMING_SHEET);

    if (!sheet) {
      sheet = ss.insertSheet(CONFIG.ZAPIER.INCOMING_SHEET);
      setupZapierIncomingSheet(sheet);
    }

    const row = [
      new Date().toISOString(),
      data.source || '',
      data.companyName || '',
      data.title || '',
      data.description || '',
      data.url || '',
      data.industry || '',
      data.employees || '',
      data.signalType || '',
      'PENDING'  // 処理ステータス
    ];

    sheet.appendRow(row);

    return ContentService.createTextOutput(JSON.stringify({
      status: 'success',
      message: 'Data received'
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    console.error('Webhook受信エラー:', error);
    return ContentService.createTextOutput(JSON.stringify({
      status: 'error',
      message: error.message
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Zapier受信シートの初期設定
 */
function setupZapierIncomingSheet(sheet) {
  const headers = [
    '受信日時', 'ソース', '企業名', 'タイトル', '説明',
    'URL', '業界', '従業員数', 'シグナルタイプ', '処理ステータス'
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setBackground('#f57c00')
    .setFontColor('white')
    .setFontWeight('bold');
}

/**
 * Zapier受信データを取得（未処理分）
 */
function fetchZapierIncoming() {
  const results = [];

  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.ZAPIER.INCOMING_SHEET);

    if (!sheet) return results;

    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const status = row[9];

      // 未処理データのみ取得
      if (status === 'PENDING') {
        results.push({
          source: row[1] || 'ZAPIER',
          companyName: row[2] || '',
          title: row[3] || '',
          description: row[4] || '',
          url: row[5] || '',
          industry: row[6] || '',
          employees: parseInt(row[7]) || 0,
          signalType: row[8] || '',
          date: row[0]
        });

        // ステータスを処理済みに更新
        sheet.getRange(i + 1, 10).setValue('PROCESSED');
      }
    }
  } catch (e) {
    console.log('Zapierデータ取得エラー:', e.message);
  }

  console.log(`Zapier受信データ: ${results.length}件`);
  return results;
}

/**
 * 高スコア企業をメールで通知
 */
function notifyHighScoreCompanies(companies) {
  if (!CONFIG.EMAIL.NOTIFICATION_TO) {
    console.log('通知メールアドレスが未設定のためスキップ');
    return;
  }

  const highScoreCompanies = companies.filter(
    c => c.totalScore >= CONFIG.EMAIL.NOTIFICATION_SCORE_THRESHOLD
  );

  if (highScoreCompanies.length === 0) {
    console.log('高スコア企業なし - メール送信スキップ');
    return;
  }

  try {
    const today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd');

    // メール件名
    const subject = `【リード検知】${today} - ${highScoreCompanies.length}社の高スコア企業を検出`;

    // メール本文（HTML形式）
    const htmlBody = generateEmailHtml(highScoreCompanies, today);

    // プレーンテキスト版
    const plainBody = generateEmailPlain(highScoreCompanies, today);

    // メール送信
    MailApp.sendEmail({
      to: CONFIG.EMAIL.NOTIFICATION_TO,
      subject: subject,
      body: plainBody,
      htmlBody: htmlBody,
      name: CONFIG.EMAIL.SENDER_NAME
    });

    console.log(`メール通知送信完了: ${highScoreCompanies.length}社`);

  } catch (e) {
    console.error('メール送信エラー:', e.message);
  }
}

/**
 * 通知メールのHTML本文を生成
 */
function generateEmailHtml(companies, date) {
  const rows = companies.map(company => {
    const urls = company.signals.slice(0, 3).map(s => s.url).filter(u => u);
    const urlLinks = urls.map(u => `<a href="${u}" style="color:#1a73e8;font-size:12px;">ソース</a>`).join(' | ');
    const summary = generateSummary(company);

    // 炎上度に応じた色
    const fireColor = company.fireLevel === 'High' ? '#dc3545' :
                      company.fireLevel === 'Mid' ? '#ffc107' : '#28a745';

    return `
      <tr style="border-bottom:1px solid #e0e0e0;">
        <td style="padding:12px 8px;font-weight:bold;">${company.companyName}</td>
        <td style="padding:12px 8px;text-align:center;">
          <span style="background-color:${company.totalScore >= 0.8 ? '#28a745' : '#ffc107'};color:white;padding:4px 8px;border-radius:4px;font-weight:bold;">
            ${(company.totalScore * 100).toFixed(0)}%
          </span>
        </td>
        <td style="padding:12px 8px;">${company.primarySignal}</td>
        <td style="padding:12px 8px;">${company.needsType}</td>
        <td style="padding:12px 8px;text-align:center;">
          <span style="color:${fireColor};font-weight:bold;">${company.fireLevel}</span>
        </td>
        <td style="padding:12px 8px;font-size:13px;color:#666;">${summary.replace(/\n/g, '<br>')}</td>
        <td style="padding:12px 8px;">${urlLinks}</td>
      </tr>
    `;
  }).join('');

  return `
    <!DOCTYPE html>
    <html>
    <head>
      <meta charset="utf-8">
    </head>
    <body style="font-family:'Helvetica Neue',Arial,sans-serif;margin:0;padding:20px;background-color:#f5f5f5;">
      <div style="max-width:900px;margin:0 auto;background-color:white;border-radius:8px;box-shadow:0 2px 4px rgba(0,0,0,0.1);">

        <!-- ヘッダー -->
        <div style="background-color:#4285f4;color:white;padding:20px;border-radius:8px 8px 0 0;">
          <h1 style="margin:0;font-size:24px;">🎯 クライアントピックアップ通知</h1>
          <p style="margin:8px 0 0;opacity:0.9;">${date} の高スコア企業レポート</p>
        </div>

        <!-- サマリー -->
        <div style="padding:20px;border-bottom:1px solid #e0e0e0;">
          <p style="margin:0;font-size:16px;">
            本日 <strong style="color:#4285f4;font-size:24px;">${companies.length}</strong> 社の高スコア企業を検出しました。
          </p>
          <p style="margin:8px 0 0;color:#666;font-size:14px;">
            スコア ${(CONFIG.EMAIL.NOTIFICATION_SCORE_THRESHOLD * 100).toFixed(0)}% 以上の企業のみ抽出
          </p>
        </div>

        <!-- テーブル -->
        <div style="padding:20px;overflow-x:auto;">
          <table style="width:100%;border-collapse:collapse;font-size:14px;">
            <thead>
              <tr style="background-color:#f8f9fa;border-bottom:2px solid #dee2e6;">
                <th style="padding:12px 8px;text-align:left;">企業名</th>
                <th style="padding:12px 8px;text-align:center;">スコア</th>
                <th style="padding:12px 8px;text-align:left;">主因</th>
                <th style="padding:12px 8px;text-align:left;">想定ニーズ</th>
                <th style="padding:12px 8px;text-align:center;">炎上度</th>
                <th style="padding:12px 8px;text-align:left;">課題要約</th>
                <th style="padding:12px 8px;text-align:left;">ソース</th>
              </tr>
            </thead>
            <tbody>
              ${rows}
            </tbody>
          </table>
        </div>

        <!-- フッター -->
        <div style="padding:20px;background-color:#f8f9fa;border-radius:0 0 8px 8px;text-align:center;">
          <p style="margin:0;color:#666;font-size:13px;">
            📊 詳細は <a href="https://docs.google.com/spreadsheets/d/${CONFIG.SPREADSHEET_ID}" style="color:#1a73e8;">スプレッドシート</a> をご確認ください
          </p>
          <p style="margin:8px 0 0;color:#999;font-size:12px;">
            このメールはクライアントピックアップシステムから自動送信されています
          </p>
        </div>

      </div>
    </body>
    </html>
  `;
}

/**
 * 通知メールのプレーンテキスト本文を生成
 */
function generateEmailPlain(companies, date) {
  const companyList = companies.map((company, i) => {
    const summary = generateSummary(company);
    const urls = company.signals.slice(0, 3).map(s => s.url).filter(u => u).join('\n    ');

    return `
${i + 1}. ${company.companyName}
   スコア: ${(company.totalScore * 100).toFixed(0)}%
   主因: ${company.primarySignal}
   想定ニーズ: ${company.needsType}
   炎上度: ${company.fireLevel}
   課題: ${summary}
   ソース:
    ${urls}
`;
  }).join('\n' + '─'.repeat(50) + '\n');

  return `
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
🎯 クライアントピックアップ通知
${date} の高スコア企業レポート
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

本日 ${companies.length} 社の高スコア企業を検出しました。
（スコア ${(CONFIG.EMAIL.NOTIFICATION_SCORE_THRESHOLD * 100).toFixed(0)}% 以上）

${'─'.repeat(50)}
${companyList}
${'─'.repeat(50)}

📊 詳細はスプレッドシートをご確認ください:
https://docs.google.com/spreadsheets/d/${CONFIG.SPREADSHEET_ID}

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
このメールはクライアントピックアップシステムから自動送信されています
`;
}

/**
 * レビューOK企業をCRMに送信（Zapier Webhook経由）
 */
function sendToCRM() {
  if (!CONFIG.ZAPIER.CRM_WEBHOOK) {
    console.log('CRM Webhookが未設定 - メール通知で代替します');
    sendReviewedCompaniesEmail();
    return;
  }

  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.SHEETS.DAILY_PICKUP);

  if (!sheet) return;

  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const reviewResult = row[11]; // レビュー結果列

    // OK かつ 未送信の企業を送信
    if (reviewResult === 'OK') {
      try {
        const payload = {
          company_name: row[1],
          industry: row[2],
          employees: row[3],
          score: row[4],
          primary_signal: row[5],
          needs_type: row[6],
          source_urls: row[7],
          summary: row[8],
          approach_text: row[9],
          fire_level: row[10],
          date: row[0]
        };

        const response = UrlFetchApp.fetch(CONFIG.ZAPIER.CRM_WEBHOOK, {
          method: 'POST',
          contentType: 'application/json',
          payload: JSON.stringify(payload),
          muteHttpExceptions: true
        });

        // 送信済みマークを付ける（レビュー結果を更新）
        sheet.getRange(i + 1, 12).setValue('OK→CRM送信済');
        console.log(`CRM送信: ${row[1]}`);

      } catch (e) {
        console.error(`CRM送信エラー (${row[1]}):`, e.message);
      }
    }
  }
}

/**
 * レビューOK企業をメールで送信（CRM Webhook未設定時の代替）
 */
function sendReviewedCompaniesEmail() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.SHEETS.DAILY_PICKUP);

  if (!sheet) return;

  const data = sheet.getDataRange().getValues();
  const okCompanies = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const reviewResult = row[11];

    if (reviewResult === 'OK') {
      okCompanies.push({
        date: row[0],
        companyName: row[1],
        industry: row[2],
        employees: row[3],
        score: row[4],
        primarySignal: row[5],
        needsType: row[6],
        urls: row[7],
        summary: row[8],
        approachText: row[9],
        fireLevel: row[10]
      });

      // 送信済みマーク
      sheet.getRange(i + 1, 12).setValue('OK→メール送信済');
    }
  }

  if (okCompanies.length === 0) {
    console.log('レビューOK企業なし');
    return;
  }

  // メール本文生成
  const today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd');
  const subject = `【営業対象確定】${okCompanies.length}社のレビューOK企業`;

  const body = okCompanies.map((c, i) => `
${i + 1}. ${c.companyName}
━━━━━━━━━━━━━━━━━━━━━
スコア: ${c.score}
想定ニーズ: ${c.needsType}
主因: ${c.primarySignal}
業界: ${c.industry}
従業員: ${c.employees}

【課題要約】
${c.summary}

【推奨アプローチ文】
${c.approachText}

【ソース】
${c.urls}

`).join('\n' + '═'.repeat(50) + '\n');

  MailApp.sendEmail({
    to: CONFIG.EMAIL.NOTIFICATION_TO,
    subject: subject,
    body: `レビューOK企業一覧（${today}）\n\n${body}`,
    name: CONFIG.EMAIL.SENDER_NAME
  });

  console.log(`レビューOK企業メール送信: ${okCompanies.length}社`);
}

/**
 * 手動インポートシートからデータ取得
 */
function fetchManualImport() {
  const results = [];

  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEETS.COMPANIES);

    if (!sheet) return results;

    const data = sheet.getDataRange().getValues();
    const headers = data[0];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[0]) continue; // 空行スキップ

      results.push({
        source: 'MANUAL',
        companyName: row[0] || '',
        industry: row[1] || '',
        employees: parseInt(row[2]) || 0,
        description: row[3] || '',
        url: row[4] || '',
        signals: row[5] || '',
        date: new Date().toISOString()
      });
    }
  } catch (e) {
    console.log('手動インポート取得エラー:', e.message);
  }

  return results;
}

/**
 * Google News検索（オプション）
 */
function fetchGoogleNews() {
  const results = [];
  const queries = [
    '情報漏えい スタートアップ',
    'IPO準備 SaaS',
    '資金調達 フィンテック',
    'セキュリティインシデント'
  ];

  queries.forEach(query => {
    try {
      const url = `https://www.googleapis.com/customsearch/v1?key=${CONFIG.GOOGLE_API_KEY}&cx=${CONFIG.GOOGLE_CSE_ID}&q=${encodeURIComponent(query)}&num=10&dateRestrict=d7`;
      const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });

      if (response.getResponseCode() === 200) {
        const json = JSON.parse(response.getContentText());

        if (json.items) {
          json.items.forEach(item => {
            const companyName = extractCompanyName(item.title);
            if (companyName) {
              results.push({
                source: 'GOOGLE_NEWS',
                companyName: companyName,
                title: item.title,
                description: item.snippet,
                url: item.link,
                date: new Date().toISOString(),
                searchTerm: query
              });
            }
          });
        }
      }
    } catch (e) {
      console.log(`Google News取得エラー (${query}):`, e.message);
    }
  });

  return results;
}

/**
 * タイトルから企業名を抽出
 */
function extractCompanyName(text) {
  if (!text) return null;

  // 「株式会社○○」「○○株式会社」パターン
  const patterns = [
    /株式会社([^\s、。,\.]+)/,
    /([^\s、。,\.]+)株式会社/,
    /合同会社([^\s、。,\.]+)/,
    /([^\s、。,\.]+)合同会社/,
    /（株）([^\s、。,\.]+)/,
    /([^\s、。,\.]+)（株）/
  ];

  for (const pattern of patterns) {
    const match = text.match(pattern);
    if (match) {
      return match[0].replace(/[（）\(\)]/g, '');
    }
  }

  return null;
}

// ========================================
// 正規化・フィルタリング
// ========================================

/**
 * データの正規化と重複排除
 */
function normalizeAndDedupe(data) {
  const companyMap = new Map();

  data.forEach(item => {
    const key = normalizeCompanyName(item.companyName);

    if (companyMap.has(key)) {
      // 既存データにシグナル情報を追加
      const existing = companyMap.get(key);
      existing.sources.push(item.source);
      existing.signals.push({
        source: item.source,
        title: item.title,
        description: item.description,
        url: item.url,
        searchTerm: item.searchTerm
      });
    } else {
      companyMap.set(key, {
        companyName: item.companyName,
        normalizedName: key,
        industry: item.industry || '',
        employees: item.employees || 0,
        sources: [item.source],
        signals: [{
          source: item.source,
          title: item.title,
          description: item.description,
          url: item.url,
          searchTerm: item.searchTerm
        }]
      });
    }
  });

  return Array.from(companyMap.values());
}

/**
 * 企業名の正規化
 */
function normalizeCompanyName(name) {
  if (!name) return '';
  return name
    .replace(/株式会社|合同会社|有限会社/g, '')
    .replace(/（株）|\(株\)|（有）|\(有\)/g, '')
    .replace(/\s+/g, '')
    .toLowerCase();
}

/**
 * 企業フィルタリング
 */
function filterCompanies(companies) {
  return companies.filter(company => {
    // 従業員規模チェック（データがある場合のみ）
    if (company.employees > 0) {
      if (company.employees < CONFIG.MIN_EMPLOYEES || company.employees > CONFIG.MAX_EMPLOYEES) {
        return false;
      }
    }

    // 除外キーワードチェック
    const allText = JSON.stringify(company).toLowerCase();
    for (const keyword of CONFIG.EXCLUDE_KEYWORDS) {
      if (allText.includes(keyword.toLowerCase())) {
        return false;
      }
    }

    return true;
  });
}

// ========================================
// スコアリング
// ========================================

/**
 * 企業群のスコアリング
 */
function scoreCompanies(companies) {
  return companies.map(company => {
    const scores = calculateSignalScores(company);

    // 総合スコア計算（仕様書の式）
    const totalScore =
      WEIGHTS.COST_CUT * scores.costCut +
      WEIGHTS.GROWTH * scores.growth +
      WEIGHTS.GOVERNANCE * scores.governance +
      WEIGHTS.OPS_DEBT * scores.opsDebt +
      WEIGHTS.FIRE * scores.fire -
      WEIGHTS.NOISE_PENALTY * scores.noise;

    return {
      ...company,
      scores: scores,
      totalScore: Math.max(0, Math.min(1, totalScore)),
      primarySignal: getPrimarySignal(scores),
      fireLevel: getFireLevel(scores.fire),
      needsType: determineNeedsType(scores)
    };
  });
}

/**
 * 各シグナルスコアの計算
 */
function calculateSignalScores(company) {
  const allText = company.signals.map(s =>
    `${s.title || ''} ${s.description || ''} ${s.searchTerm || ''}`
  ).join(' ').toLowerCase();

  return {
    costCut: calculateKeywordScore(allText, SIGNAL_KEYWORDS.COST_CUT),
    growth: calculateKeywordScore(allText, SIGNAL_KEYWORDS.GROWTH),
    governance: calculateKeywordScore(allText, SIGNAL_KEYWORDS.GOVERNANCE),
    opsDebt: calculateKeywordScore(allText, SIGNAL_KEYWORDS.OPS_DEBT),
    fire: calculateKeywordScore(allText, SIGNAL_KEYWORDS.FIRE),
    noise: calculateNoiseScore(allText)
  };
}

/**
 * キーワードベースのスコア計算
 */
function calculateKeywordScore(text, keywords) {
  let score = 0;

  // 強シグナル: 0.4
  keywords.strong.forEach(kw => {
    if (text.includes(kw.toLowerCase())) score += 0.4;
  });

  // 中シグナル: 0.2
  keywords.medium.forEach(kw => {
    if (text.includes(kw.toLowerCase())) score += 0.2;
  });

  // 弱シグナル: 0.1
  keywords.weak.forEach(kw => {
    if (text.includes(kw.toLowerCase())) score += 0.1;
  });

  return Math.min(1, score);
}

/**
 * ノイズスコア計算
 */
function calculateNoiseScore(text) {
  let score = 0;
  SIGNAL_KEYWORDS.NOISE.forEach(kw => {
    if (text.includes(kw.toLowerCase())) score += 0.3;
  });
  return Math.min(1, score);
}

/**
 * 主因シグナルの判定
 */
function getPrimarySignal(scores) {
  const signals = [
    { name: 'E4（コストカット）', score: scores.costCut },
    { name: 'E1（急成長）', score: scores.growth },
    { name: 'E2（ガバナンス）', score: scores.governance },
    { name: '運用負債', score: scores.opsDebt },
    { name: '炎上', score: scores.fire }
  ];

  signals.sort((a, b) => b.score - a.score);
  return signals[0].score > 0 ? signals[0].name : '不明';
}

/**
 * 炎上度判定
 */
function getFireLevel(fireScore) {
  if (fireScore >= 0.6) return 'High';
  if (fireScore >= 0.3) return 'Mid';
  return 'Low';
}

/**
 * 想定ニーズタイプ判定
 */
function determineNeedsType(scores) {
  const hasInfoSys = scores.opsDebt > 0.3 || scores.costCut > 0.3;
  const hasLegal = scores.governance > 0.3 || scores.fire > 0.3;

  if (hasInfoSys && hasLegal) return '両方';
  if (hasInfoSys) return '情シス';
  if (hasLegal) return '法務';
  return '要判断';
}

// ========================================
// 抽出・出力
// ========================================

/**
 * 上位企業の抽出
 */
function extractTopCompanies(scored, limit) {
  return scored
    .filter(c => c.totalScore >= CONFIG.SCORE_THRESHOLD)
    .sort((a, b) => b.totalScore - a.totalScore)
    .slice(0, limit);
}

/**
 * Google Sheetsへの出力
 */
function outputToSheet(companies) {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  let sheet = ss.getSheetByName(CONFIG.SHEETS.DAILY_PICKUP);

  // シートがなければ作成
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEETS.DAILY_PICKUP);
    setupDailyPickupSheet(sheet);
  }

  const today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');

  // 今日の日付のデータを追加
  companies.forEach(company => {
    const urls = company.signals.slice(0, 3).map(s => s.url).filter(u => u).join('\n');
    const summary = generateSummary(company);
    const approach = generateApproachText(company);

    const row = [
      today,                                    // 日付
      company.companyName,                      // A: 企業名
      company.industry || '要確認',            // B: 業界
      company.employees || '要確認',           // C: 従業員規模
      Math.round(company.totalScore * 100) / 100, // D: 総合Score
      company.primarySignal,                   // E: 主因
      company.needsType,                       // F: 想定ニーズ
      urls,                                     // G: 根拠URL
      summary,                                  // H: 想定課題要約
      approach,                                 // I: 推奨アプローチ文
      company.fireLevel,                       // J: 炎上度
      ''                                        // K: レビュー結果
    ];

    sheet.appendRow(row);
  });

  // フォーマット調整
  formatSheet(sheet);
}

/**
 * Daily_Pickupシートの初期設定
 */
function setupDailyPickupSheet(sheet) {
  const headers = [
    '日付',
    '企業名',
    '業界',
    '従業員規模',
    '総合Score',
    '主因',
    '想定ニーズ',
    '根拠URL',
    '想定課題要約',
    '推奨アプローチ文',
    '炎上度',
    'レビュー結果'
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setBackground('#4285f4')
    .setFontColor('white')
    .setFontWeight('bold');

  // 列幅調整
  sheet.setColumnWidth(1, 100);  // 日付
  sheet.setColumnWidth(2, 200);  // 企業名
  sheet.setColumnWidth(3, 100);  // 業界
  sheet.setColumnWidth(4, 100);  // 従業員
  sheet.setColumnWidth(5, 80);   // Score
  sheet.setColumnWidth(6, 120);  // 主因
  sheet.setColumnWidth(7, 100);  // 想定ニーズ
  sheet.setColumnWidth(8, 250);  // URL
  sheet.setColumnWidth(9, 300);  // 課題要約
  sheet.setColumnWidth(10, 400); // アプローチ文
  sheet.setColumnWidth(11, 80);  // 炎上度
  sheet.setColumnWidth(12, 100); // レビュー結果

  // フィルタ設定
  sheet.getRange(1, 1, 1, headers.length).createFilter();

  // レビュー結果のプルダウン
  const reviewRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['OK', 'NG', '保留'], true)
    .build();
  sheet.getRange('L2:L1000').setDataValidation(reviewRule);
}

/**
 * シートフォーマット調整
 */
function formatSheet(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  // スコアによる条件付き書式
  const scoreRange = sheet.getRange(2, 5, lastRow - 1, 1);
  const rules = sheet.getConditionalFormatRules();

  // 高スコア（緑）
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThanOrEqualTo(0.7)
    .setBackground('#b7e1cd')
    .setRanges([scoreRange])
    .build());

  // 中スコア（黄）
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberBetween(0.5, 0.69)
    .setBackground('#fce8b2')
    .setRanges([scoreRange])
    .build());

  sheet.setConditionalFormatRules(rules);

  // 炎上度による書式
  const fireRange = sheet.getRange(2, 11, lastRow - 1, 1);
  const fireRules = [
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('High')
      .setBackground('#f4c7c3')
      .setRanges([fireRange])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Mid')
      .setBackground('#fce8b2')
      .setRanges([fireRange])
      .build()
  ];

  sheet.setConditionalFormatRules([...rules, ...fireRules]);
}

/**
 * 課題要約の生成
 */
function generateSummary(company) {
  const parts = [];

  if (company.scores.growth > 0.3) {
    parts.push('急成長フェーズで管理体制が追いついていない可能性');
  }
  if (company.scores.governance > 0.3) {
    parts.push('監査・統制対応のニーズあり');
  }
  if (company.scores.costCut > 0.3) {
    parts.push('コスト効率化・業務自動化を推進中');
  }
  if (company.scores.fire > 0.3) {
    parts.push('インシデント対応・リスク管理の強化が必要');
  }
  if (company.scores.opsDebt > 0.3) {
    parts.push('情シス/法務体制の整備が課題');
  }

  return parts.slice(0, 2).join('\n') || '詳細確認が必要';
}

/**
 * アプローチ文の生成
 */
function generateApproachText(company) {
  const templates = {
    '情シス': `${company.companyName}様\n\n貴社の成長フェーズにおける情報システム基盤の整備について、月次顧問という形でご支援できればと考えております。\n\n・SaaS/ツールの選定・導入支援\n・セキュリティポリシー策定\n・ID管理・アクセス権限の整備\n\n等、専任を置くほどではないが対応が必要な領域を、外部パートナーとして柔軟にサポートいたします。`,

    '法務': `${company.companyName}様\n\n貴社の事業拡大に伴う法務体制の整備について、月次顧問という形でご支援できればと考えております。\n\n・契約書レビュー・作成\n・コンプライアンス体制構築\n・規程類の整備\n\n等、専任法務を置くフェーズではないが対応が必要な領域を、外部パートナーとして柔軟にサポートいたします。`,

    '両方': `${company.companyName}様\n\n貴社の成長フェーズにおける管理体制の整備について、情シス・法務の両面から月次顧問としてご支援できればと考えております。\n\n【情シス】SaaS導入、セキュリティ、ID管理\n【法務】契約レビュー、規程整備、コンプライアンス\n\n専任を置くほどではないが対応が必要な領域を、ワンストップで柔軟にサポートいたします。`
  };

  return templates[company.needsType] || templates['両方'];
}

// ========================================
// セットアップ・ユーティリティ
// ========================================

/**
 * 初期セットアップ
 */
function initialSetup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 設定シート
  let settingsSheet = ss.getSheetByName(CONFIG.SHEETS.SETTINGS);
  if (!settingsSheet) {
    settingsSheet = ss.insertSheet(CONFIG.SHEETS.SETTINGS);
    const settings = [
      ['設定項目', '値', '説明'],
      ['SCORE_THRESHOLD', '0.5', 'スコア閾値'],
      ['DAILY_LIMIT', '20', '日次抽出件数'],
      ['MIN_EMPLOYEES', '10', '最小従業員数'],
      ['MAX_EMPLOYEES', '100', '最大従業員数'],
      ['GOOGLE_API_KEY', '', 'Google Custom Search APIキー（オプション）'],
      ['GOOGLE_CSE_ID', '', 'Custom Search Engine ID（オプション）']
    ];
    settingsSheet.getRange(1, 1, settings.length, 3).setValues(settings);
  }

  // 企業マスターシート
  let companiesSheet = ss.getSheetByName(CONFIG.SHEETS.COMPANIES);
  if (!companiesSheet) {
    companiesSheet = ss.insertSheet(CONFIG.SHEETS.COMPANIES);
    const headers = [
      ['企業名', '業界', '従業員数', '説明/メモ', 'URL', 'シグナル情報']
    ];
    companiesSheet.getRange(1, 1, 1, 6).setValues(headers);
    companiesSheet.getRange(1, 1, 1, 6)
      .setBackground('#34a853')
      .setFontColor('white')
      .setFontWeight('bold');
  }

  // Daily_Pickupシート
  let pickupSheet = ss.getSheetByName(CONFIG.SHEETS.DAILY_PICKUP);
  if (!pickupSheet) {
    pickupSheet = ss.insertSheet(CONFIG.SHEETS.DAILY_PICKUP);
    setupDailyPickupSheet(pickupSheet);
  }

  console.log('セットアップ完了');
  SpreadsheetApp.getUi().alert('セットアップが完了しました！\n\n1. CONFIG内のSPREADSHEET_IDを設定してください\n2. Companies_Masterシートに企業データをインポートしてください\n3. dailyPickup()を実行またはトリガー設定してください');
}

/**
 * デイリートリガーの設定
 */
function setupDailyTrigger() {
  // 既存のトリガーを削除
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'dailyPickup') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // 新しいトリガーを設定（毎日午前9時）
  ScriptApp.newTrigger('dailyPickup')
    .timeBased()
    .atHour(9)
    .everyDays(1)
    .inTimezone('Asia/Tokyo')
    .create();

  console.log('デイリートリガーを設定しました（毎日9:00 JST）');
}

/**
 * テスト実行（少数データで動作確認）
 */
function testRun() {
  console.log('=== テスト実行開始 ===');

  // テストデータ
  const testCompanies = [
    {
      companyName: '株式会社テストSaaS',
      industry: 'SaaS',
      employees: 50,
      signals: [{
        source: 'TEST',
        title: 'シリーズA資金調達を実施、IPO準備も開始',
        description: '業務効率化ツールを提供する株式会社テストSaaSが5億円を調達。情シス体制の強化も課題。',
        url: 'https://example.com/test1'
      }]
    },
    {
      companyName: '株式会社フィンテック',
      industry: '金融',
      employees: 30,
      signals: [{
        source: 'TEST',
        title: 'ISMS取得に向けて体制整備',
        description: 'コンプライアンス強化のため法務・情シス機能の外部委託を検討中。',
        url: 'https://example.com/test2'
      }]
    }
  ];

  // スコアリング
  const scored = scoreCompanies(testCompanies);

  scored.forEach(c => {
    console.log(`\n企業: ${c.companyName}`);
    console.log(`総合スコア: ${c.totalScore}`);
    console.log(`主因: ${c.primarySignal}`);
    console.log(`想定ニーズ: ${c.needsType}`);
    console.log(`スコア詳細:`, c.scores);
  });

  console.log('\n=== テスト実行完了 ===');
  return scored;
}

/**
 * カスタムメニューの追加
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🎯 クライアントピックアップ')
    .addItem('📋 初期セットアップ', 'initialSetup')
    .addItem('▶️ 今すぐ実行', 'dailyPickup')
    .addItem('🧪 テスト実行', 'testRun')
    .addSeparator()
    .addItem('⏰ デイリートリガー設定', 'setupDailyTrigger')
    .addSeparator()
    .addSubMenu(ui.createMenu('📧 メール・通知')
      .addItem('📤 レビューOK企業をメール送信', 'sendToCRM')
      .addItem('🧪 テストメール送信', 'sendTestEmail'))
    .addSeparator()
    .addSubMenu(ui.createMenu('🔗 Zapier連携')
      .addItem('📊 Zapier受信データ確認', 'checkZapierIncoming'))
    .addToUi();
}

/**
 * テストメール送信
 */
function sendTestEmail() {
  const testCompanies = [
    {
      companyName: '株式会社テストSaaS',
      totalScore: 0.85,
      primarySignal: 'E1（急成長）',
      needsType: '両方',
      fireLevel: 'Low',
      scores: { growth: 0.6, governance: 0.4, costCut: 0.2, opsDebt: 0.5, fire: 0.1 },
      signals: [{ url: 'https://example.com/test1' }]
    },
    {
      companyName: '株式会社テストフィンテック',
      totalScore: 0.75,
      primarySignal: 'E2（ガバナンス）',
      needsType: '法務',
      fireLevel: 'Mid',
      scores: { growth: 0.3, governance: 0.7, costCut: 0.1, opsDebt: 0.2, fire: 0.3 },
      signals: [{ url: 'https://example.com/test2' }]
    }
  ];

  notifyHighScoreCompanies(testCompanies);
  SpreadsheetApp.getUi().alert('テストメールを送信しました。\n\n送信先: ' + CONFIG.EMAIL.NOTIFICATION_TO);
}

/**
 * Zapier受信データの確認
 */
function checkZapierIncoming() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  let sheet = ss.getSheetByName(CONFIG.ZAPIER.INCOMING_SHEET);

  if (!sheet) {
    SpreadsheetApp.getUi().alert('Zapier受信シートがありません。\nZapierからデータを送信するとシートが作成されます。');
    return;
  }

  const data = sheet.getDataRange().getValues();
  const pendingCount = data.filter((row, i) => i > 0 && row[9] === 'PENDING').length;
  const processedCount = data.filter((row, i) => i > 0 && row[9] === 'PROCESSED').length;

  SpreadsheetApp.getUi().alert(
    `Zapier受信データ状況\n\n` +
    `未処理: ${pendingCount}件\n` +
    `処理済: ${processedCount}件\n` +
    `合計: ${data.length - 1}件`
  );
}
