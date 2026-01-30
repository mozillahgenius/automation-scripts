/**
 * 日本進出初期外資系企業 検知システム
 *
 * 4レイヤー検知:
 * - Stage 0: 採用・Web変化（最速）
 * - Stage 1: VC/投資家ブログ
 * - Stage 2: 海外ニュース・PR
 * - Stage 3: 日本語メディア（回収網）
 *
 * 判定ロジック:
 * (A: 法人/法務 ∨ B: 採用/人) ∧ (C: Web/プロダクト ∨ D: 投資/VC) → 営業対象
 */

// ============================================
// 設定
// ============================================
const CONFIG = {
  // スプレッドシートID（出力先）
  SPREADSHEET_ID: 'SPREADSHEET_ID_PLACEHOLDER', // ← ここにスプレッドシートIDを設定

  // シート名
  SHEETS: {
    COMPANIES: '検知企業一覧',
    NEWS: 'ニュース検知',
    HIRING: '採用情報',
    WEB_CHANGES: 'Web変化',
    VC_BLOG: 'VC/投資家ブログ',
    SETTINGS: '設定',
    LOG: '実行ログ'
  },

  // API設定
  NEWS_API_KEY: 'YOUR_NEWS_API_KEY', // ← News APIのキーを設定

  // 検索キーワード
  SEARCH_KEYWORDS: {
    NEWS: [
      '"Japan expansion"',
      '"launch in Japan"',
      '"entering the Japanese market"',
      '"APAC expansion" AND "Japan"',
      '"open office in Tokyo"',
      '"Japanese market entry"'
    ],
    HIRING: [
      'Head of Japan',
      'Country Manager Japan',
      'Japan General Manager',
      'Japan Director',
      'VP Japan',
      'Japan Managing Director'
    ]
  },

  // 監視対象VCブログRSS
  VC_BLOG_RSS: [
    'https://a16z.com/feed/',
    'https://www.sequoiacap.com/feed/',
    // 他のVCブログRSSを追加
  ],

  // 監視対象企業リスト（Web変化検知用）
  TARGET_COMPANIES: [], // 設定シートから読み込み

  // 通知設定
  NOTIFICATION_EMAIL: 'your-email@example.com', // ← 通知先メールアドレス

  // 実行間隔
  TRIGGER_INTERVALS: {
    NEWS: 'daily',      // 日次
    HIRING: 'twice_weekly', // 週2回
    WEB: 'twice_weekly',    // 週2回
    VC_BLOG: 'daily'    // 日次
  }
};

// ============================================
// メイン実行関数
// ============================================

/**
 * 全検知処理を実行
 */
function runAllDetection() {
  const startTime = new Date();
  log('=== 検知処理開始 ===');

  try {
    // Stage 0: 採用・Web変化（最速）
    const hiringResults = detectHiringSignals();
    const webResults = detectWebChanges();

    // Stage 1: VC/投資家ブログ
    const vcResults = detectVCBlogMentions();

    // Stage 2-3: ニュース検知
    const newsResults = detectNewsSignals();

    // 複合判定
    const qualifiedCompanies = evaluateCompanies(hiringResults, webResults, vcResults, newsResults);

    // 結果出力
    outputResults(qualifiedCompanies);

    // 日次レポート送信（検知0件でも送信）
    sendDailyReport(qualifiedCompanies);

    const duration = (new Date() - startTime) / 1000;
    log(`=== 検知処理完了 (${duration}秒) ===`);
    log(`検知企業数: ${qualifiedCompanies.length}`);

  } catch (error) {
    log(`エラー: ${error.message}`, 'ERROR');
    throw error;
  }
}

/**
 * 日次実行（ニュース・VCブログ）
 */
function runDailyDetection() {
  log('=== 日次検知処理開始 ===');

  const newsResults = detectNewsSignals();
  const vcResults = detectVCBlogMentions();

  // 既存データと統合して判定（日次レポート送信あり）
  updateAndEvaluateWithReport(newsResults, vcResults, [], []);
}

/**
 * 週2回実行（採用・Web変化）
 */
function runTwiceWeeklyDetection() {
  log('=== 週2回検知処理開始 ===');

  const hiringResults = detectHiringSignals();
  const webResults = detectWebChanges();

  // 既存データと統合して判定
  updateAndEvaluate([], [], hiringResults, webResults);
}

// ============================================
// Stage 0: 採用シグナル検知
// ============================================

/**
 * 採用情報から日本進出シグナルを検知
 */
function detectHiringSignals() {
  log('採用シグナル検知開始...');
  const results = [];

  CONFIG.SEARCH_KEYWORDS.HIRING.forEach(keyword => {
    try {
      // LinkedIn Jobs API または他の求人APIを使用
      const jobs = searchJobPostings(keyword);

      jobs.forEach(job => {
        const signal = {
          type: 'HIRING',
          companyName: job.company,
          companyDomain: extractDomain(job.companyUrl),
          title: job.title,
          location: job.location,
          url: job.url,
          detectedAt: new Date(),
          confidence: calculateHiringConfidence(job),
          raw: job
        };

        results.push(signal);
        log(`採用検知: ${job.company} - ${job.title}`);
      });

    } catch (error) {
      log(`採用検索エラー (${keyword}): ${error.message}`, 'WARN');
    }

    // API制限対策
    Utilities.sleep(1000);
  });

  // 結果をシートに保存
  saveSignalsToSheet(results, CONFIG.SHEETS.HIRING);

  log(`採用シグナル検知完了: ${results.length}件`);
  return results;
}

/**
 * 求人検索（実装例：Google Custom Search API使用）
 */
function searchJobPostings(keyword) {
  const jobs = [];

  // LinkedIn求人ページをスクレイピング対象として検索
  const searchQuery = `site:linkedin.com/jobs ${keyword}`;

  try {
    // Google Custom Search APIを使用
    const apiKey = PropertiesService.getScriptProperties().getProperty('GOOGLE_API_KEY');
    const searchEngineId = PropertiesService.getScriptProperties().getProperty('SEARCH_ENGINE_ID');

    if (!apiKey || !searchEngineId) {
      log('Google API設定が不足しています', 'WARN');
      return jobs;
    }

    const url = `https://www.googleapis.com/customsearch/v1?key=${apiKey}&cx=${searchEngineId}&q=${encodeURIComponent(searchQuery)}&num=10`;

    const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const data = JSON.parse(response.getContentText());

    if (data.items) {
      data.items.forEach(item => {
        // タイトルと説明から企業名を抽出
        const companyMatch = item.title.match(/at\s+(.+?)(?:\s*\||$)/i);
        if (companyMatch) {
          jobs.push({
            company: companyMatch[1].trim(),
            title: item.title.split(' at ')[0] || item.title,
            url: item.link,
            companyUrl: '',
            location: 'Japan',
            snippet: item.snippet
          });
        }
      });
    }

  } catch (error) {
    log(`求人検索APIエラー: ${error.message}`, 'ERROR');
  }

  return jobs;
}

/**
 * 採用情報の確度スコア計算
 */
function calculateHiringConfidence(job) {
  let score = 50; // 基本スコア

  const title = job.title.toLowerCase();

  // 役職による加点
  if (title.includes('country manager') || title.includes('head of japan')) {
    score += 30;
  } else if (title.includes('general manager') || title.includes('managing director')) {
    score += 25;
  } else if (title.includes('director') || title.includes('vp')) {
    score += 20;
  }

  // 地域明示による加点
  if (title.includes('japan') || title.includes('tokyo')) {
    score += 10;
  }

  return Math.min(score, 100);
}

// ============================================
// Stage 0: Web変化検知
// ============================================

/**
 * 対象企業のWeb変化を検知
 */
function detectWebChanges() {
  log('Web変化検知開始...');
  const results = [];

  // 設定シートから監視対象企業を取得
  const targetCompanies = getTargetCompanies();

  targetCompanies.forEach(company => {
    try {
      const changes = checkWebsiteChanges(company);

      if (changes.hasJapaneseContent || changes.hasJapaneseLegalTerms) {
        const signal = {
          type: 'WEB_CHANGE',
          companyName: company.name,
          companyDomain: company.domain,
          changes: changes,
          detectedAt: new Date(),
          confidence: calculateWebConfidence(changes),
          raw: changes
        };

        results.push(signal);
        log(`Web変化検知: ${company.name} - ${JSON.stringify(changes.detected)}`);
      }

    } catch (error) {
      log(`Web検知エラー (${company.name}): ${error.message}`, 'WARN');
    }

    // レート制限対策
    Utilities.sleep(2000);
  });

  // 結果をシートに保存
  saveSignalsToSheet(results, CONFIG.SHEETS.WEB_CHANGES);

  log(`Web変化検知完了: ${results.length}件`);
  return results;
}

/**
 * Webサイトの変化をチェック
 */
function checkWebsiteChanges(company) {
  const result = {
    hasJapaneseContent: false,
    hasJapaneseLegalTerms: false,
    hasJapaneseLP: false,
    detected: []
  };

  try {
    // メインサイトをフェッチ
    const response = UrlFetchApp.fetch(`https://${company.domain}`, {
      muteHttpExceptions: true,
      followRedirects: true
    });

    const content = response.getContentText();

    // 日本語コンテンツの検出
    if (/[\u3040-\u309F\u30A0-\u30FF\u4E00-\u9FAF]/.test(content)) {
      result.hasJapaneseContent = true;
      result.detected.push('日本語コンテンツ');
    }

    // 日本語ページへのリンク検出
    if (/href=["'][^"']*\/ja\/|href=["'][^"']*\/jp\//i.test(content)) {
      result.hasJapaneseLP = true;
      result.detected.push('日本語LP');
    }

    // 日本法準拠の利用規約検出
    if (/日本法|Japanese law|Japan jurisdiction/i.test(content)) {
      result.hasJapaneseLegalTerms = true;
      result.detected.push('日本法準拠規約');
    }

    // 東京オフィスの記載検出
    if (/Tokyo office|東京オフィス|Japan office/i.test(content)) {
      result.detected.push('日本オフィス記載');
    }

  } catch (error) {
    log(`Webフェッチエラー (${company.domain}): ${error.message}`, 'WARN');
  }

  return result;
}

/**
 * Web変化の確度スコア計算
 */
function calculateWebConfidence(changes) {
  let score = 30; // 基本スコア

  if (changes.hasJapaneseLegalTerms) score += 30;
  if (changes.hasJapaneseLP) score += 25;
  if (changes.hasJapaneseContent) score += 15;

  return Math.min(score, 100);
}

// ============================================
// Stage 1: VC/投資家ブログ検知
// ============================================

/**
 * VCブログからJapan expansion言及を検知
 */
function detectVCBlogMentions() {
  log('VC/投資家ブログ検知開始...');
  const results = [];

  CONFIG.VC_BLOG_RSS.forEach(rssUrl => {
    try {
      const feed = fetchRSSFeed(rssUrl);

      feed.items.forEach(item => {
        // Japan関連キーワードでフィルタ
        if (containsJapanKeywords(item.title + ' ' + item.description)) {
          const signal = {
            type: 'VC_BLOG',
            source: feed.title,
            title: item.title,
            url: item.link,
            publishedAt: item.pubDate,
            detectedAt: new Date(),
            confidence: 70,
            companyName: extractCompanyFromText(item.title + ' ' + item.description),
            raw: item
          };

          results.push(signal);
          log(`VCブログ検知: ${feed.title} - ${item.title}`);
        }
      });

    } catch (error) {
      log(`RSSフェッチエラー (${rssUrl}): ${error.message}`, 'WARN');
    }

    Utilities.sleep(1000);
  });

  // 結果をシートに保存
  saveSignalsToSheet(results, CONFIG.SHEETS.VC_BLOG);

  log(`VCブログ検知完了: ${results.length}件`);
  return results;
}

/**
 * RSSフィードを取得
 */
function fetchRSSFeed(rssUrl) {
  const response = UrlFetchApp.fetch(rssUrl, { muteHttpExceptions: true });
  const xml = response.getContentText();
  const document = XmlService.parse(xml);
  const root = document.getRootElement();

  const feed = {
    title: '',
    items: []
  };

  // RSS 2.0形式の解析
  const channel = root.getChild('channel');
  if (channel) {
    feed.title = channel.getChildText('title') || '';

    const items = channel.getChildren('item');
    items.forEach(item => {
      feed.items.push({
        title: item.getChildText('title') || '',
        link: item.getChildText('link') || '',
        description: item.getChildText('description') || '',
        pubDate: item.getChildText('pubDate') || ''
      });
    });
  }

  return feed;
}

// ============================================
// Stage 2-3: ニュース検知
// ============================================

/**
 * ニュースAPIから日本進出シグナルを検知
 */
function detectNewsSignals() {
  log('ニュース検知開始...');
  const results = [];

  CONFIG.SEARCH_KEYWORDS.NEWS.forEach(keyword => {
    try {
      const articles = searchNews(keyword);

      articles.forEach(article => {
        const signal = {
          type: 'NEWS',
          source: article.source,
          title: article.title,
          url: article.url,
          publishedAt: article.publishedAt,
          detectedAt: new Date(),
          confidence: calculateNewsConfidence(article),
          companyName: extractCompanyFromText(article.title + ' ' + article.description),
          raw: article
        };

        results.push(signal);
        log(`ニュース検知: ${article.source} - ${article.title}`);
      });

    } catch (error) {
      log(`ニュース検索エラー (${keyword}): ${error.message}`, 'WARN');
    }

    Utilities.sleep(1000);
  });

  // 結果をシートに保存
  saveSignalsToSheet(results, CONFIG.SHEETS.NEWS);

  log(`ニュース検知完了: ${results.length}件`);
  return results;
}

/**
 * ニュース検索（News API使用）
 */
function searchNews(keyword) {
  const articles = [];

  const apiKey = CONFIG.NEWS_API_KEY || PropertiesService.getScriptProperties().getProperty('NEWS_API_KEY');

  if (!apiKey) {
    log('News API Keyが設定されていません', 'WARN');
    return articles;
  }

  try {
    const url = `https://newsapi.org/v2/everything?q=${encodeURIComponent(keyword)}&language=en&sortBy=publishedAt&pageSize=20&apiKey=${apiKey}`;

    const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const data = JSON.parse(response.getContentText());

    if (data.status === 'ok' && data.articles) {
      data.articles.forEach(article => {
        articles.push({
          source: article.source.name,
          title: article.title,
          description: article.description || '',
          url: article.url,
          publishedAt: article.publishedAt
        });
      });
    }

  } catch (error) {
    log(`News APIエラー: ${error.message}`, 'ERROR');
  }

  return articles;
}

/**
 * ニュースの確度スコア計算
 */
function calculateNewsConfidence(article) {
  let score = 60; // 基本スコア

  const text = (article.title + ' ' + article.description).toLowerCase();

  // 高確度キーワード
  if (text.includes('japan expansion') || text.includes('launch in japan')) {
    score += 25;
  }

  // 中確度キーワード
  if (text.includes('tokyo office') || text.includes('japanese market')) {
    score += 15;
  }

  // 信頼度の高いソース
  const trustedSources = ['techcrunch', 'reuters', 'bloomberg', 'nikkei'];
  if (trustedSources.some(s => article.source.toLowerCase().includes(s))) {
    score += 10;
  }

  return Math.min(score, 100);
}

// ============================================
// 複合判定ロジック
// ============================================

/**
 * 検知結果を統合して営業対象を判定
 * 判定ロジック: (A: 法人/法務 ∨ B: 採用/人) ∧ (C: Web/プロダクト ∨ D: 投資/VC) → 営業対象
 */
function evaluateCompanies(hiringResults, webResults, vcResults, newsResults) {
  log('複合判定開始...');

  // 企業ごとにシグナルを集約
  const companySignals = new Map();

  // 採用シグナル（B: 採用/人）
  hiringResults.forEach(signal => {
    const key = normalizeCompanyName(signal.companyName);
    if (!companySignals.has(key)) {
      companySignals.set(key, createCompanyRecord(signal.companyName));
    }
    const record = companySignals.get(key);
    record.signals.hiring.push(signal);
    record.hasHiring = true;
  });

  // Web変化シグナル（C: Web/プロダクト）
  webResults.forEach(signal => {
    const key = normalizeCompanyName(signal.companyName);
    if (!companySignals.has(key)) {
      companySignals.set(key, createCompanyRecord(signal.companyName));
    }
    const record = companySignals.get(key);
    record.signals.web.push(signal);
    record.hasWeb = true;
  });

  // VCブログシグナル（D: 投資/VC）
  vcResults.forEach(signal => {
    if (signal.companyName) {
      const key = normalizeCompanyName(signal.companyName);
      if (!companySignals.has(key)) {
        companySignals.set(key, createCompanyRecord(signal.companyName));
      }
      const record = companySignals.get(key);
      record.signals.vc.push(signal);
      record.hasVC = true;
    }
  });

  // ニュースシグナル（D: 投資/VC と同等扱い）
  newsResults.forEach(signal => {
    if (signal.companyName) {
      const key = normalizeCompanyName(signal.companyName);
      if (!companySignals.has(key)) {
        companySignals.set(key, createCompanyRecord(signal.companyName));
      }
      const record = companySignals.get(key);
      record.signals.news.push(signal);
      record.hasNews = true;
    }
  });

  // 判定: (B: 採用 ∨ C: Web) ∧ (D: ニュース or VC)
  const qualifiedCompanies = [];

  companySignals.forEach((record, key) => {
    const conditionBC = record.hasHiring || record.hasWeb;
    const conditionD = record.hasNews || record.hasVC;

    if (conditionBC && conditionD) {
      // 最優先営業対象
      record.priority = 'HIGH';
      record.reason = '複合シグナル検知';
      qualifiedCompanies.push(record);
      log(`営業対象判定: ${record.companyName} - HIGH (複合シグナル)`);
    } else if (conditionBC) {
      // 中優先（採用またはWeb変化のみ）
      record.priority = 'MEDIUM';
      record.reason = '採用/Web変化のみ';
      qualifiedCompanies.push(record);
      log(`営業対象判定: ${record.companyName} - MEDIUM (採用/Webのみ)`);
    } else if (conditionD) {
      // 低優先（ニュースまたはVCのみ）
      record.priority = 'LOW';
      record.reason = 'ニュース/VCのみ';
      qualifiedCompanies.push(record);
      log(`営業対象判定: ${record.companyName} - LOW (ニュース/VCのみ)`);
    }

    // 総合スコア計算
    record.totalScore = calculateTotalScore(record);
  });

  // スコア順にソート
  qualifiedCompanies.sort((a, b) => {
    const priorityOrder = { HIGH: 0, MEDIUM: 1, LOW: 2 };
    if (priorityOrder[a.priority] !== priorityOrder[b.priority]) {
      return priorityOrder[a.priority] - priorityOrder[b.priority];
    }
    return b.totalScore - a.totalScore;
  });

  log(`複合判定完了: ${qualifiedCompanies.length}社が営業対象`);
  return qualifiedCompanies;
}

/**
 * 企業レコードを作成
 */
function createCompanyRecord(companyName) {
  return {
    companyName: companyName,
    domain: '',
    signals: {
      hiring: [],
      web: [],
      vc: [],
      news: []
    },
    hasHiring: false,
    hasWeb: false,
    hasVC: false,
    hasNews: false,
    priority: '',
    reason: '',
    totalScore: 0,
    detectedAt: new Date()
  };
}

/**
 * 総合スコア計算
 */
function calculateTotalScore(record) {
  let score = 0;

  // 各シグナルの最高スコアを加算
  if (record.signals.hiring.length > 0) {
    score += Math.max(...record.signals.hiring.map(s => s.confidence));
  }
  if (record.signals.web.length > 0) {
    score += Math.max(...record.signals.web.map(s => s.confidence));
  }
  if (record.signals.vc.length > 0) {
    score += Math.max(...record.signals.vc.map(s => s.confidence));
  }
  if (record.signals.news.length > 0) {
    score += Math.max(...record.signals.news.map(s => s.confidence));
  }

  return score;
}

// ============================================
// 結果出力
// ============================================

/**
 * 検知結果をスプレッドシートに出力
 */
function outputResults(companies) {
  log('結果出力開始...');

  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.SHEETS.COMPANIES);

  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEETS.COMPANIES);
    // ヘッダー設定
    sheet.getRange(1, 1, 1, 12).setValues([[
      '検知日時',
      '企業名',
      'ドメイン',
      '優先度',
      '総合スコア',
      '採用シグナル',
      'Web変化',
      'VCブログ',
      'ニュース',
      '判定理由',
      'ステータス',
      '備考'
    ]]);
    sheet.getRange(1, 1, 1, 12).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }

  // データ追加
  const newRows = companies.map(company => [
    company.detectedAt,
    company.companyName,
    company.domain,
    company.priority,
    company.totalScore,
    company.signals.hiring.length > 0 ? '✓' : '',
    company.signals.web.length > 0 ? '✓' : '',
    company.signals.vc.length > 0 ? '✓' : '',
    company.signals.news.length > 0 ? '✓' : '',
    company.reason,
    '未対応',
    ''
  ]);

  if (newRows.length > 0) {
    const lastRow = sheet.getLastRow();
    sheet.getRange(lastRow + 1, 1, newRows.length, 12).setValues(newRows);

    // 優先度による条件付き書式
    applyConditionalFormatting(sheet);
  }

  log(`結果出力完了: ${newRows.length}件`);
}

/**
 * 条件付き書式を適用
 */
function applyConditionalFormatting(sheet) {
  const range = sheet.getDataRange();

  // 既存のルールをクリア
  sheet.clearConditionalFormatRules();

  const rules = [];

  // HIGH優先度：緑
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('HIGH')
    .setBackground('#c6efce')
    .setRanges([sheet.getRange('D:D')])
    .build());

  // MEDIUM優先度：黄
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('MEDIUM')
    .setBackground('#ffeb9c')
    .setRanges([sheet.getRange('D:D')])
    .build());

  // LOW優先度：赤
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('LOW')
    .setBackground('#ffc7ce')
    .setRanges([sheet.getRange('D:D')])
    .build());

  sheet.setConditionalFormatRules(rules);
}

/**
 * シグナルをシートに保存
 */
function saveSignalsToSheet(signals, sheetName) {
  if (signals.length === 0) return;

  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    // ヘッダー設定（シグナルタイプに応じて）
    const headers = getHeadersForSignalType(signals[0].type);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }

  // データ追加
  const rows = signals.map(signal => formatSignalRow(signal));

  if (rows.length > 0) {
    const lastRow = sheet.getLastRow();
    sheet.getRange(lastRow + 1, 1, rows.length, rows[0].length).setValues(rows);
  }
}

/**
 * シグナルタイプに応じたヘッダーを取得
 */
function getHeadersForSignalType(type) {
  switch (type) {
    case 'HIRING':
      return ['検知日時', '企業名', '役職', 'URL', '確度', '備考'];
    case 'WEB_CHANGE':
      return ['検知日時', '企業名', 'ドメイン', '検知内容', '確度', '備考'];
    case 'VC_BLOG':
      return ['検知日時', 'ソース', 'タイトル', 'URL', '企業名', '確度'];
    case 'NEWS':
      return ['検知日時', 'ソース', 'タイトル', 'URL', '企業名', '確度'];
    default:
      return ['検知日時', '種別', '内容', 'URL', '確度'];
  }
}

/**
 * シグナルを行データにフォーマット
 */
function formatSignalRow(signal) {
  switch (signal.type) {
    case 'HIRING':
      return [signal.detectedAt, signal.companyName, signal.title, signal.url, signal.confidence, ''];
    case 'WEB_CHANGE':
      return [signal.detectedAt, signal.companyName, signal.companyDomain, signal.changes.detected.join(', '), signal.confidence, ''];
    case 'VC_BLOG':
      return [signal.detectedAt, signal.source, signal.title, signal.url, signal.companyName || '', signal.confidence];
    case 'NEWS':
      return [signal.detectedAt, signal.source, signal.title, signal.url, signal.companyName || '', signal.confidence];
    default:
      return [signal.detectedAt, signal.type, JSON.stringify(signal), signal.url || '', signal.confidence];
  }
}

// ============================================
// 通知機能
// ============================================

/**
 * 検知結果をメール通知
 */
function sendNotification(companies) {
  const email = CONFIG.NOTIFICATION_EMAIL || PropertiesService.getScriptProperties().getProperty('NOTIFICATION_EMAIL');

  if (!email) {
    log('通知メールアドレスが設定されていません', 'WARN');
    return;
  }

  const highPriority = companies.filter(c => c.priority === 'HIGH');
  const mediumPriority = companies.filter(c => c.priority === 'MEDIUM');
  const lowPriority = companies.filter(c => c.priority === 'LOW');

  // スプレッドシートURL
  const spreadsheetUrl = `https://docs.google.com/spreadsheets/d/${CONFIG.SPREADSHEET_ID}/edit?gid=971308449#gid=971308449`;

  let body = `【日本進出外資系企業 検知レポート】\n\n`;
  body += `検知日時: ${new Date().toLocaleString('ja-JP')}\n`;
  body += `総検知数: ${companies.length}社\n\n`;

  // 検知概要サマリー
  body += `━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n`;
  body += `【検知概要】\n`;
  body += `  HIGH（複合シグナル）: ${highPriority.length}社\n`;
  body += `  MEDIUM（採用/Webのみ）: ${mediumPriority.length}社\n`;
  body += `  LOW（ニュース/VCのみ）: ${lowPriority.length}社\n`;
  body += `━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n`;

  if (highPriority.length > 0) {
    body += `■ 最優先営業対象 (${highPriority.length}社)\n`;
    highPriority.forEach(c => {
      body += `\n  【${c.companyName}】\n`;
      body += `    スコア: ${c.totalScore}\n`;
      body += `    シグナル: ${getSignalSummary(c)}\n`;
      body += `    ${getSignalDetails(c)}`;
    });
    body += '\n';
  }

  if (mediumPriority.length > 0) {
    body += `\n■ 中優先営業対象 (${mediumPriority.length}社)\n`;
    mediumPriority.forEach(c => {
      body += `\n  【${c.companyName}】\n`;
      body += `    スコア: ${c.totalScore}\n`;
      body += `    シグナル: ${getSignalSummary(c)}\n`;
      body += `    ${getSignalDetails(c)}`;
    });
    body += '\n';
  }

  if (lowPriority.length > 0) {
    body += `\n■ 低優先営業対象 (${lowPriority.length}社)\n`;
    lowPriority.forEach(c => {
      body += `  - ${c.companyName} (${getSignalSummary(c)})\n`;
    });
    body += '\n';
  }

  body += `\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n`;
  body += `詳細はスプレッドシートをご確認ください:\n`;
  body += `${spreadsheetUrl}\n`;
  body += `━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━`;

  GmailApp.sendEmail(email, `【検知】日本進出外資 ${highPriority.length}社（HIGH）/ ${mediumPriority.length}社（MEDIUM）`, body);
  log(`通知メール送信完了: ${email}`);
}

/**
 * 日次レポートをメール通知（検知0件でも送信）
 */
function sendDailyReport(companies) {
  const email = CONFIG.NOTIFICATION_EMAIL || PropertiesService.getScriptProperties().getProperty('NOTIFICATION_EMAIL');

  if (!email) {
    log('通知メールアドレスが設定されていません', 'WARN');
    return;
  }

  // スプレッドシートURL
  const spreadsheetUrl = `https://docs.google.com/spreadsheets/d/${CONFIG.SPREADSHEET_ID}/edit?gid=971308449#gid=971308449`;

  const today = new Date().toLocaleDateString('ja-JP', {
    year: 'numeric',
    month: 'long',
    day: 'numeric'
  });

  let subject;
  let body;

  if (companies.length === 0) {
    // 検知0件の場合
    subject = `【日次レポート】${today} - 新規検知なし`;

    body = `【日本進出外資系企業 日次レポート】\n\n`;
    body += `検知日時: ${new Date().toLocaleString('ja-JP')}\n\n`;
    body += `━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n`;
    body += `本日の新規検知はありませんでした。\n`;
    body += `━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n`;
    body += `引き続き以下のソースを監視中です:\n`;
    body += `  - ニュース（海外メディア）\n`;
    body += `  - 採用情報（LinkedIn等）\n`;
    body += `  - VCブログ（a16z, Sequoia等）\n`;
    body += `  - Web変化（日本語ページ追加等）\n\n`;
    body += `━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n`;
    body += `過去の検知結果はスプレッドシートをご確認ください:\n`;
    body += `${spreadsheetUrl}\n`;
    body += `━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━`;

  } else {
    // 検知ありの場合は通常の通知を使用
    sendNotification(companies);
    return;
  }

  GmailApp.sendEmail(email, subject, body);
  log(`日次レポートメール送信完了: ${email}`);
}

/**
 * シグナルサマリーを取得
 */
function getSignalSummary(company) {
  const signals = [];
  if (company.hasHiring) signals.push('採用');
  if (company.hasWeb) signals.push('Web変化');
  if (company.hasVC) signals.push('VCブログ');
  if (company.hasNews) signals.push('ニュース');
  return signals.join(', ');
}

/**
 * シグナル詳細を取得
 */
function getSignalDetails(company) {
  let details = '';

  // 採用シグナルの詳細
  if (company.signals.hiring && company.signals.hiring.length > 0) {
    details += `\n    [採用]\n`;
    company.signals.hiring.slice(0, 3).forEach(s => {
      details += `      - ${s.title || '役職不明'}\n`;
      if (s.url) details += `        ${s.url}\n`;
    });
  }

  // Web変化シグナルの詳細
  if (company.signals.web && company.signals.web.length > 0) {
    details += `\n    [Web変化]\n`;
    company.signals.web.slice(0, 3).forEach(s => {
      if (s.changes && s.changes.detected) {
        details += `      - 検知: ${s.changes.detected.join(', ')}\n`;
      }
      if (s.companyDomain) details += `        ドメイン: ${s.companyDomain}\n`;
    });
  }

  // VCブログシグナルの詳細
  if (company.signals.vc && company.signals.vc.length > 0) {
    details += `\n    [VCブログ]\n`;
    company.signals.vc.slice(0, 3).forEach(s => {
      details += `      - ${s.source || 'ソース不明'}: ${s.title || 'タイトル不明'}\n`;
      if (s.url) details += `        ${s.url}\n`;
    });
  }

  // ニュースシグナルの詳細
  if (company.signals.news && company.signals.news.length > 0) {
    details += `\n    [ニュース]\n`;
    company.signals.news.slice(0, 3).forEach(s => {
      details += `      - ${s.source || 'ソース不明'}: ${s.title || 'タイトル不明'}\n`;
      if (s.url) details += `        ${s.url}\n`;
    });
  }

  return details || '詳細情報なし';
}

// ============================================
// ユーティリティ関数
// ============================================

/**
 * スプレッドシートを取得
 */
function getSpreadsheet() {
  const id = CONFIG.SPREADSHEET_ID || PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');

  if (id) {
    return SpreadsheetApp.openById(id);
  }

  // IDが設定されていない場合は新規作成
  const ss = SpreadsheetApp.create('外資系企業検知システム');
  PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', ss.getId());
  log(`新規スプレッドシート作成: ${ss.getId()}`);
  return ss;
}

/**
 * 設定シートから監視対象企業を取得
 */
function getTargetCompanies() {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.SETTINGS);

  if (!sheet) {
    return [];
  }

  const data = sheet.getDataRange().getValues();
  const companies = [];

  // ヘッダー行をスキップ
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][1]) {
      companies.push({
        name: data[i][0],
        domain: data[i][1]
      });
    }
  }

  return companies;
}

/**
 * Japan関連キーワードを含むか判定
 */
function containsJapanKeywords(text) {
  const keywords = [
    'japan expansion',
    'launch in japan',
    'japanese market',
    'tokyo office',
    'apac expansion',
    'entering japan',
    'japan operations'
  ];

  const lowerText = text.toLowerCase();
  return keywords.some(kw => lowerText.includes(kw));
}

/**
 * テキストから企業名を抽出
 */
function extractCompanyFromText(text) {
  // 簡易的な企業名抽出（改善の余地あり）
  const patterns = [
    /([A-Z][a-zA-Z0-9]+(?:\s+[A-Z][a-zA-Z0-9]+)*)\s+(?:launches?|expands?|enters?|opens?)/i,
    /([A-Z][a-zA-Z0-9]+(?:\s+[A-Z][a-zA-Z0-9]+)*)\s+(?:to\s+)?(?:launch|expand|enter|open)/i
  ];

  for (const pattern of patterns) {
    const match = text.match(pattern);
    if (match) {
      return match[1].trim();
    }
  }

  return null;
}

/**
 * URLからドメインを抽出
 */
function extractDomain(url) {
  if (!url) return '';

  try {
    const match = url.match(/^(?:https?:\/\/)?(?:www\.)?([^\/]+)/i);
    return match ? match[1] : '';
  } catch {
    return '';
  }
}

/**
 * 企業名を正規化
 */
function normalizeCompanyName(name) {
  if (!name) return '';
  return name.toLowerCase()
    .replace(/,?\s*(inc\.?|corp\.?|ltd\.?|llc\.?|co\.?)$/i, '')
    .replace(/\s+/g, ' ')
    .trim();
}

/**
 * ログ出力
 */
function log(message, level = 'INFO') {
  const timestamp = new Date().toISOString();
  console.log(`[${timestamp}] [${level}] ${message}`);

  // ログシートにも記録
  try {
    const ss = getSpreadsheet();
    let sheet = ss.getSheetByName(CONFIG.SHEETS.LOG);

    if (!sheet) {
      sheet = ss.insertSheet(CONFIG.SHEETS.LOG);
      sheet.getRange(1, 1, 1, 3).setValues([['日時', 'レベル', 'メッセージ']]);
      sheet.getRange(1, 1, 1, 3).setFontWeight('bold');
    }

    sheet.appendRow([new Date(), level, message]);

    // ログが1000行を超えたら古いものを削除
    if (sheet.getLastRow() > 1000) {
      sheet.deleteRows(2, 100);
    }

  } catch (error) {
    // ログ記録エラーは無視
  }
}

/**
 * 既存データと統合して再判定
 */
function updateAndEvaluate(newsResults, vcResults, hiringResults, webResults) {
  // 既存のシグナルを読み込み
  const existingHiring = hiringResults.length > 0 ? hiringResults : loadExistingSignals(CONFIG.SHEETS.HIRING);
  const existingWeb = webResults.length > 0 ? webResults : loadExistingSignals(CONFIG.SHEETS.WEB_CHANGES);
  const existingVC = vcResults.length > 0 ? vcResults : loadExistingSignals(CONFIG.SHEETS.VC_BLOG);
  const existingNews = newsResults.length > 0 ? newsResults : loadExistingSignals(CONFIG.SHEETS.NEWS);

  const qualifiedCompanies = evaluateCompanies(existingHiring, existingWeb, existingVC, existingNews);

  outputResults(qualifiedCompanies);

  if (qualifiedCompanies.filter(c => c.priority === 'HIGH').length > 0) {
    sendNotification(qualifiedCompanies);
  }
}

/**
 * 既存データと統合して再判定（日次レポート送信あり）
 */
function updateAndEvaluateWithReport(newsResults, vcResults, hiringResults, webResults) {
  // 既存のシグナルを読み込み
  const existingHiring = hiringResults.length > 0 ? hiringResults : loadExistingSignals(CONFIG.SHEETS.HIRING);
  const existingWeb = webResults.length > 0 ? webResults : loadExistingSignals(CONFIG.SHEETS.WEB_CHANGES);
  const existingVC = vcResults.length > 0 ? vcResults : loadExistingSignals(CONFIG.SHEETS.VC_BLOG);
  const existingNews = newsResults.length > 0 ? newsResults : loadExistingSignals(CONFIG.SHEETS.NEWS);

  const qualifiedCompanies = evaluateCompanies(existingHiring, existingWeb, existingVC, existingNews);

  outputResults(qualifiedCompanies);

  // 日次レポート送信（検知0件でも送信）
  sendDailyReport(qualifiedCompanies);
}

/**
 * 既存シグナルを読み込み
 */
function loadExistingSignals(sheetName) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);

  if (!sheet || sheet.getLastRow() <= 1) {
    return [];
  }

  // 直近7日分のみ読み込み
  const data = sheet.getDataRange().getValues();
  const signals = [];
  const sevenDaysAgo = new Date();
  sevenDaysAgo.setDate(sevenDaysAgo.getDate() - 7);

  for (let i = 1; i < data.length; i++) {
    const detectedAt = new Date(data[i][0]);
    if (detectedAt >= sevenDaysAgo) {
      signals.push({
        type: getTypeFromSheetName(sheetName),
        companyName: data[i][1],
        detectedAt: detectedAt,
        confidence: data[i][4] || 50
      });
    }
  }

  return signals;
}

/**
 * シート名からシグナルタイプを取得
 */
function getTypeFromSheetName(sheetName) {
  switch (sheetName) {
    case CONFIG.SHEETS.HIRING: return 'HIRING';
    case CONFIG.SHEETS.WEB_CHANGES: return 'WEB_CHANGE';
    case CONFIG.SHEETS.VC_BLOG: return 'VC_BLOG';
    case CONFIG.SHEETS.NEWS: return 'NEWS';
    default: return 'UNKNOWN';
  }
}

// ============================================
// トリガー設定
// ============================================

/**
 * トリガーを設定
 */
function setupTriggers() {
  // 既存トリガーを削除
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => ScriptApp.deleteTrigger(trigger));

  // 日次トリガー（毎朝9時）
  ScriptApp.newTrigger('runDailyDetection')
    .timeBased()
    .atHour(9)
    .everyDays(1)
    .create();

  // 週2回トリガー（火曜・金曜の9時）
  ScriptApp.newTrigger('runTwiceWeeklyDetection')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.TUESDAY)
    .atHour(9)
    .create();

  ScriptApp.newTrigger('runTwiceWeeklyDetection')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.FRIDAY)
    .atHour(9)
    .create();

  log('トリガー設定完了');
}

/**
 * 初期設定シートを作成
 */
function createSettingsSheet() {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.SHEETS.SETTINGS);

  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEETS.SETTINGS);

    // ヘッダー
    sheet.getRange(1, 1, 1, 4).setValues([['企業名', 'ドメイン', '国', '備考']]);
    sheet.getRange(1, 1, 1, 4).setFontWeight('bold');

    // サンプルデータ
    sheet.getRange(2, 1, 3, 4).setValues([
      ['Stripe', 'stripe.com', 'US', '決済'],
      ['Notion', 'notion.so', 'US', 'ドキュメント'],
      ['Figma', 'figma.com', 'US', 'デザイン']
    ]);

    sheet.setFrozenRows(1);
  }

  log('設定シート作成完了');
}

// ============================================
// メニュー追加
// ============================================

/**
 * スプレッドシート開時にメニューを追加
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('外資検知システム')
    .addItem('全検知実行', 'runAllDetection')
    .addItem('日次検知実行', 'runDailyDetection')
    .addItem('週2回検知実行', 'runTwiceWeeklyDetection')
    .addSeparator()
    .addItem('トリガー設定', 'setupTriggers')
    .addItem('設定シート作成', 'createSettingsSheet')
    .addToUi();
}
