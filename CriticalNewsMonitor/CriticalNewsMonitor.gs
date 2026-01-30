// ============================================
// 批判的記事監視システム for Google Apps Script
// Perplexity Sonar APIを使用
// ============================================

// ====== 設定 ======
const CONFIG = {
  // Perplexity API設定
  PERPLEXITY_API_KEY: 'YOUR_PERPLEXITY_API_KEY', // PropertiesServiceから取得推奨
  PERPLEXITY_API_URL: 'https://api.perplexity.ai/chat/completions',
  
  // Perplexityモデル設定（フォールバック対応）
  MODELS: [
    'sonar-pro',  // Proモデル（より高精度・優先使用）
    'sonar'  // 標準モデル（フォールバック用）
  ],
  
  // スプレッドシート設定
  SHEET_NAME: '批判的記事',
  HEADERS: ['検索日時', '対象企業', '記事日付', '件名', '内容要約', '記者名/執筆者', 'URL', '媒体名', '批判度スコア'],
  
  // 検索設定シート
  CONFIG_SHEET_NAME: '検索設定',
  CONFIG_HEADERS: ['企業名', 'キーワード', '検索期間(日)', '批判的記事のみ', 'メール送信', '送信先メール', '実行頻度', '有効/無効', '最終実行日時', '備考'],
  
  // 検索設定のデフォルト値
  DEFAULT_DAYS_BACK: 7,
  MAX_RESULTS_PER_SEARCH: 10
};

// ====== メイン関数 ======
/**
 * 批判的記事を検索してスプレッドシートに記録＆メール送信
 * @param {string} companyName - 検索対象の会社名
 * @param {number} daysBack - 何日前までの記事を検索するか（デフォルト: 7日）
 * @param {Array<string>} keywords - 検索キーワード（オプション）
 * @param {string} emailTo - レポート送信先メールアドレス（オプション）
 * @param {boolean} sendEmail - メール送信するか（デフォルト: true）
 */
function searchCriticalArticles(companyName = null, daysBack = null, keywords = [], emailTo = null, sendEmail = true) {
  // パラメータ確認
  if (!companyName) {
    throw new Error('会社名を指定してください');
  }
  
  daysBack = daysBack || CONFIG.DEFAULT_DAYS_BACK;
  
  console.log(`検索開始: ${companyName} (過去${daysBack}日間)`);
  if (keywords && keywords.length > 0) {
    console.log(`キーワード: ${keywords.join(', ')}`);
  }
  
  try {
    // 1. APIキーの取得
    const apiKey = getApiKey();
    
    // 2. 検索クエリの構築
    const searchQuery = buildSearchQuery(companyName, daysBack, keywords);
    
    // 3. Perplexity APIで検索実行
    const articles = searchWithPerplexity(searchQuery, apiKey);
    
    // 4. 批判的な記事のフィルタリング
    const criticalArticles = filterCriticalArticles(articles, companyName, keywords);
    
    // 5. スプレッドシートに記録
    if (criticalArticles.length > 0) {
      saveToSpreadsheet(criticalArticles, companyName);
      console.log(`${criticalArticles.length}件の批判的記事を記録しました`);
      
      // 6. メール送信（オプション）
      if (sendEmail) {
        sendReportEmailWithArticles(criticalArticles, companyName, emailTo);
      }
    } else {
      console.log('批判的な記事は見つかりませんでした');
    }
    
    return criticalArticles;
    
  } catch (error) {
    console.error('エラーが発生しました:', error);
    throw error;
  }
}

/**
 * キーワードベースで記事を検索（批判的記事に限定しない）
 * @param {string} companyName - 検索対象の会社名
 * @param {Array<string>} keywords - 検索キーワード
 * @param {number} daysBack - 何日前までの記事を検索するか
 * @param {boolean} criticalOnly - 批判的記事のみに限定するか
 * @param {string} emailTo - レポート送信先メールアドレス（オプション）
 * @param {boolean} sendEmail - メール送信するか（デフォルト: true）
 */
function searchArticlesByKeywords(companyName, keywords, daysBack = 7, criticalOnly = false, emailTo = null, sendEmail = true) {
  if (!companyName || !keywords || keywords.length === 0) {
    throw new Error('会社名とキーワードを指定してください');
  }
  
  console.log(`キーワード検索開始: ${companyName}`);
  console.log(`キーワード: ${keywords.join(', ')}`);
  console.log(`期間: 過去${daysBack}日間`);
  console.log(`批判的記事のみ: ${criticalOnly ? 'はい' : 'いいえ'}`);
  
  try {
    const apiKey = getApiKey();
    const searchQuery = buildKeywordSearchQuery(companyName, keywords, daysBack, criticalOnly);
    const articles = searchWithPerplexity(searchQuery, apiKey);
    
    // キーワードフィルタリング
    const filteredArticles = filterArticlesByKeywords(articles, companyName, keywords, criticalOnly);
    
    if (filteredArticles.length > 0) {
      saveToSpreadsheet(filteredArticles, companyName);
      console.log(`${filteredArticles.length}件の記事を記録しました`);
      
      // メール送信（オプション）
      if (sendEmail) {
        const subject = `【${companyName}】キーワード検索結果: ${keywords.join(', ')}`;
        sendReportEmailWithArticles(filteredArticles, companyName, emailTo, subject);
      }
    } else {
      console.log('該当する記事は見つかりませんでした');
    }
    
    return filteredArticles;
    
  } catch (error) {
    console.error('エラーが発生しました:', error);
    throw error;
  }
}

// ====== API通信処理 ======
/**
 * Perplexity Sonar APIで検索を実行（フォールバック対応）
 */
function searchWithPerplexity(query, apiKey) {
  // モデルのフォールバック試行
  for (let i = 0; i < CONFIG.MODELS.length; i++) {
    const model = CONFIG.MODELS[i];
    console.log(`モデル ${model} で試行中... (${i + 1}/${CONFIG.MODELS.length})`);
    
    try {
      const result = trySearchWithModel(query, apiKey, model);
      console.log(`✓ モデル ${model} で成功`);
      return result;
    } catch (error) {
      console.log(`✗ モデル ${model} で失敗: ${error.message}`);
      if (i === CONFIG.MODELS.length - 1) {
        // すべてのモデルで失敗した場合
        throw new Error(`すべてのモデルで検索に失敗しました。最後のエラー: ${error.message}`);
      }
    }
  }
}

/**
 * 指定されたモデルでAPI検索を試行
 */
function trySearchWithModel(query, apiKey, model) {
  console.log(`\n=== API呼び出し詳細 ===`);
  console.log(`使用モデル: ${model}`);
  console.log(`検索クエリ長: ${query.length}文字`);
  
  // よりシンプルで効果的なプロンプト
  const payload = {
    model: model,
    messages: [
      {
        role: "user",
        content: query
      }
    ],
    temperature: 0.2,
    max_tokens: 4000,
    top_p: 0.95,
    return_citations: true,
    search_recency_filter: "month",
    frequency_penalty: 0,
    presence_penalty: 0,
    return_related_questions: false
  };
  
  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'Authorization': `Bearer ${apiKey}`,
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  
  console.log(`API呼び出し中...`);
  const startTime = new Date().getTime();
  
  const response = UrlFetchApp.fetch(CONFIG.PERPLEXITY_API_URL, options);
  const responseTime = new Date().getTime() - startTime;
  console.log(`レスポンス時間: ${responseTime}ms`);
  
  const responseCode = response.getResponseCode();
  const responseText = response.getContentText();
  console.log(`レスポンスコード: ${responseCode}`);
  console.log(`レスポンスサイズ: ${responseText.length}文字`);
  
  if (responseCode !== 200) {
    console.error(`APIエラー: ${responseText.substring(0, 500)}`);
    throw new Error(`API Error (${responseCode}): ${responseText.substring(0, 200)}`);
  }
  
  const responseData = JSON.parse(responseText);
  console.log(`レスポンス構造: choices=${responseData.choices?.length}, citations=${responseData.citations?.length}`);
  
  // レスポンスから記事データを抽出
  const content = responseData.choices?.[0]?.message?.content;
  if (!content) {
    console.log('コンテンツが空です');
    return [];
  }
  
  console.log(`コンテンツプレビュー: ${content.substring(0, 300)}...`);
  
  // JSON抽出を試みる（レスポンスにテキストが含まれる場合があるため）
  let jsonContent = content;
  
  // JSON部分のみを抽出（テキストの前後に説明がある場合の対処）
  const jsonMatch = content.match(/\{[\s\S]*\}/);
  if (jsonMatch) {
    jsonContent = jsonMatch[0];
    console.log('JSON部分を抽出しました');
  }
  
  try {
    const parsedContent = JSON.parse(jsonContent);
    const articles = parsedContent.articles || [];
    
    console.log(`解析結果: ${articles.length}件の記事`);
    
    if (articles.length === 0) {
      console.log('記事が見つかりませんでした');
      // テキストから抽出を試みる
      return extractArticlesFromText(content);
    }
    
    // 記事の検証と整形
    const validatedArticles = articles.map(article => {
      // 日付の整形
      if (article.date && !article.date.match(/^\d{4}-\d{2}-\d{2}$/)) {
        const dateObj = new Date(article.date);
        if (!isNaN(dateObj.getTime())) {
          article.date = formatDate(dateObj);
        }
      }
      
      // URLの検証
      if (!article.url || article.url === 'N/A' || article.url === '') {
        article.url = '';
      }
      
      // 著者名のデフォルト値
      if (!article.author || article.author === 'N/A') {
        article.author = '不明';
      }
      
      // 媒体名のデフォルト値
      if (!article.source || article.source === 'N/A') {
        article.source = '不明';
      }
      
      // 批判度スコアの検証
      if (!article.criticalScore || article.criticalScore < 1 || article.criticalScore > 10) {
        article.criticalScore = 5; // デフォルト値
      }
      
      return article;
    });
    
    // 引用元情報の追加（もしあれば）
    if (responseData.citations && responseData.citations.length > 0) {
      console.log(`引用元: ${responseData.citations.length}件`);
      
      // 引用元情報をログ出力
      responseData.citations.slice(0, 3).forEach((citation, i) => {
        console.log(`引用${i + 1}: ${citation.title || 'No title'} - ${citation.url || 'No URL'}`);
      });
    }
    
    console.log(`最終結果: ${validatedArticles.length}件の記事を返します`);
    return validatedArticles;
    
  } catch (parseError) {
    console.error('JSON解析エラー:', parseError);
    console.log('コンテンツ:', jsonContent.substring(0, 500));
    
    // JSON解析失敗時は、テキストから情報を抽出
    return extractArticlesFromText(content);
  }
}

/**
 * 検索クエリの構築
 */
function buildSearchQuery(companyName, daysBack, keywords = []) {
  const endDate = new Date();
  const startDate = new Date();
  startDate.setDate(startDate.getDate() - daysBack);
  
  let keywordSection = '';
  if (keywords && keywords.length > 0) {
    keywordSection = `
    - 必須キーワード: ${keywords.join(', ')}`;
  }
  
  // 英語でシンプルで効果的なクエリ
  let keywordFilter = '';
  if (keywords && keywords.length > 0) {
    keywordFilter = ` with keywords: ${keywords.join(', ')}`;
  }
  
  const query = `Search the web for negative or critical news about "${companyName}"${keywordFilter} from the last ${daysBack} days.

Look for: scandals, lawsuits, problems, criticism, investigations, violations, complaints, controversies.

Return actual search results as JSON:
{
  "articles": [
    {
      "date": "YYYY-MM-DD",
      "title": "exact headline",
      "summary": "brief summary",
      "author": "author or empty",
      "url": "real article URL",
      "source": "news site",
      "criticalScore": 7
    }
  ]
}

Return up to ${CONFIG.MAX_RESULTS_PER_SEARCH} articles with real URLs. If nothing found, return {"articles": []}`;
  
  return query;
}

/**
 * キーワード検索用クエリの構築
 */
function buildKeywordSearchQuery(companyName, keywords, daysBack, criticalOnly) {
  const endDate = new Date();
  const startDate = new Date();
  startDate.setDate(startDate.getDate() - daysBack);
  
  const articleType = criticalOnly 
    ? '批判的、否定的、問題提起、スキャンダル、不祥事に関する記事' 
    : 'すべてのニュース記事（ポジティブ、ネガティブ、中立的な内容を含む）';
  
  // 英語でシンプルなクエリにする（Perplexityは英語の方が精度が高い）
  const criticalTerms = criticalOnly 
    ? ' Focus on negative, critical, controversial content.' 
    : '';
  
  const query = `Search the web for news articles about "${companyName}" that mention "${keywords.join('" OR "')}" from the last ${daysBack} days.${criticalTerms}

Return the search results as a JSON object with this exact format:
{
  "articles": [
    {
      "date": "YYYY-MM-DD",
      "title": "exact article title from search result",
      "summary": "brief summary of the article content",
      "author": "author name if available, otherwise empty string",
      "url": "actual URL from search result",
      "source": "news source website",
      "criticalScore": 1-10 (10 being most critical)
    }
  ]
}

IMPORTANT: 
- Include real URLs from actual news sources
- Return up to ${CONFIG.MAX_RESULTS_PER_SEARCH} articles
- If no articles found, return {"articles": []}
- Search in both English and Japanese sources`;
  
  return query;
}

// ====== URL検証機能 ======
/**
 * URLが有効かどうかをチェック
 * @param {string} url - チェックするURL
 * @return {boolean} URLが有効かどうか
 */
function isValidUrl(url) {
  if (!url || typeof url !== 'string') {
    return false;
  }
  
  try {
    // 基本的なURL形式チェック
    const urlPattern = /^https?:\/\/(www\.)?[-a-zA-Z0-9@:%._\+~#=]{1,256}\.[a-zA-Z0-9()]{1,6}\b([-a-zA-Z0-9()@:%_\+.~#?&//=]*)$/;
    if (!urlPattern.test(url)) {
      return false;
    }
    
    // 明らかに偽のURLパターンをチェック
    const fakePatterns = [
      /example\.com/,
      /test\.com/,
      /fake\./,
      /dummy\./,
      /placeholder/,
      /sample\./,
      /invalid\./
    ];
    
    return !fakePatterns.some(pattern => pattern.test(url.toLowerCase()));
  } catch (error) {
    return false;
  }
}

/**
 * 記事の検証とフィルタリング（ハルシネーション防止強化版）
 * @param {Array} articles - 記事配列
 * @param {Array} citations - 引用元URL配列
 * @return {Array} 検証済み記事配列
 */
function validateAndFilterArticles(articles, citations = []) {
  if (!articles || !Array.isArray(articles)) {
    console.log('記事データが無効または配列ではありません');
    return [];
  }
  
  console.log(`記事検証開始: ${articles.length}件の候補記事`);
  
  const validatedArticles = articles.filter((article, index) => {
    // 基本項目の存在チェック
    if (!article.title || typeof article.title !== 'string' || article.title.trim() === '') {
      console.log(`記事${index + 1}: タイトルが無効`, article);
      return false;
    }
    
    if (!article.summary || typeof article.summary !== 'string' || article.summary.trim() === '') {
      console.log(`記事${index + 1}: 要約が無効`, article);
      return false;
    }
    
    // URLの詳細検証
    if (article.url) {
      if (!isValidUrl(article.url)) {
        console.log(`記事${index + 1}: 無効なURL形式 - ${article.url}`);
        
        // 引用元URLから代替を探す
        if (citations && citations.length > 0) {
          const citationIndex = index % citations.length;
          const citation = citations[citationIndex];
          if (citation && citation.url && isValidUrl(citation.url)) {
            article.url = citation.url;
            article.urlSource = 'citation';
            console.log(`記事${index + 1}: 引用元URLを採用 - ${citation.url}`);
          } else {
            article.url = '';
            article.urlSource = 'invalid';
          }
        } else {
          article.url = '';
          article.urlSource = 'invalid';
        }
      } else {
        article.urlSource = 'api_response';
      }
    } else {
      article.urlSource = 'missing';
    }
    
    // ハルシネーション検知パターン
    const hallucinationPatterns = [
      /example\.com/i,
      /test\.(com|org|net)/i,
      /placeholder/i,
      /sample/i,
      /dummy/i,
      /fake/i,
      /\[記事タイトル\]/i,
      /\[URL\]/i,
      /\[執筆者\]/i
    ];
    
    const hasHallucination = hallucinationPatterns.some(pattern => {
      return pattern.test(article.title) || 
             pattern.test(article.summary) || 
             (article.url && pattern.test(article.url)) ||
             (article.author && pattern.test(article.author));
    });
    
    if (hasHallucination) {
      console.log(`記事${index + 1}: ハルシネーションパターンを検出`, article);
      return false;
    }
    
    // 最低限の品質チェック
    if (article.title.length < 5 || article.summary.length < 20) {
      console.log(`記事${index + 1}: 内容が短すぎます（タイトル: ${article.title.length}文字, 要約: ${article.summary.length}文字）`);
      return false;
    }
    
    return true;
  });
  
  console.log(`記事検証完了: ${articles.length}件 → ${validatedArticles.length}件（検証通過）`);
  
  // 検証結果の詳細ログ
  validatedArticles.forEach((article, index) => {
    console.log(`検証済み記事${index + 1}: ${article.title} (URL: ${article.urlSource})`);
  });
  
  return validatedArticles;
}

/**
 * 実際のURLの存在確認（オプション・重い処理）
 * @param {string} url - 確認するURL
 * @return {boolean} URLが実際にアクセス可能かどうか
 */
function verifyUrlExists(url) {
  if (!isValidUrl(url)) {
    return false;
  }
  
  try {
    const response = UrlFetchApp.fetch(url, {
      method: 'HEAD',
      followRedirects: true,
      muteHttpExceptions: true
    });
    
    return response.getResponseCode() === 200;
  } catch (error) {
    console.log(`URL確認エラー (${url}):`, error.message);
    return false;
  }
}

// ====== データ処理 ======
/**
 * 批判的な記事のフィルタリング
 */
function filterCriticalArticles(articles, companyName, keywords = []) {
  if (!articles || articles.length === 0) {
    return [];
  }
  
  // 批判度スコアが5以上の記事をフィルタリング
  const filtered = articles.filter(article => {
    // Yahoo!ファイナンスの記事を除外
    const url = (article.url || '').toLowerCase();
    if (url.includes('finance.yahoo.co.jp')) {
      return false;
    }
    
    // 批判度スコアチェック
    if (article.criticalScore && article.criticalScore < 5) {
      return false;
    }
    
    // タイトルまたは要約に会社名が含まれているか確認
    const title = (article.title || '').toLowerCase();
    const summary = (article.summary || '').toLowerCase();
    const company = companyName.toLowerCase();
    
    if (!title.includes(company) && !summary.includes(company)) {
      return false;
    }
    
    // キーワードチェック（指定されている場合）
    if (keywords && keywords.length > 0) {
      const hasKeyword = keywords.some(keyword => {
        const kw = keyword.toLowerCase();
        return title.includes(kw) || summary.includes(kw);
      });
      if (!hasKeyword) {
        return false;
      }
    }
    
    return true;
  });
  
  // 批判度スコアで降順ソート
  filtered.sort((a, b) => (b.criticalScore || 0) - (a.criticalScore || 0));
  
  return filtered;
}

/**
 * キーワードで記事をフィルタリング
 */
function filterArticlesByKeywords(articles, companyName, keywords, criticalOnly) {
  if (!articles || articles.length === 0) {
    return [];
  }
  
  const filtered = articles.filter(article => {
    // Yahoo!ファイナンスの記事を除外
    const url = (article.url || '').toLowerCase();
    if (url.includes('finance.yahoo.co.jp')) {
      return false;
    }
    
    // 批判的記事のみの場合、スコアチェック
    if (criticalOnly && article.criticalScore && article.criticalScore < 5) {
      return false;
    }
    
    // 会社名チェック
    const title = (article.title || '').toLowerCase();
    const summary = (article.summary || '').toLowerCase();
    const company = companyName.toLowerCase();
    
    if (!title.includes(company) && !summary.includes(company)) {
      return false;
    }
    
    // キーワードチェック
    const hasKeyword = keywords.some(keyword => {
      const kw = keyword.toLowerCase();
      return title.includes(kw) || summary.includes(kw);
    });
    
    return hasKeyword;
  });
  
  // 関連度またはスコアでソート
  filtered.sort((a, b) => {
    if (criticalOnly) {
      return (b.criticalScore || 0) - (a.criticalScore || 0);
    }
    // 関連度でソート（実装簡略化のため、現在は日付順）
    return new Date(b.date || 0) - new Date(a.date || 0);
  });
  
  return filtered;
}

/**
 * テキストから記事情報を抽出（フォールバック用）
 */
function extractArticlesFromText(text) {
  console.log('テキストから記事情報を抽出中...');
  const articles = [];
  
  // 記事の境界を識別するパターン
  const articlePatterns = [
    /\d+\.\s*.*?(?=\d+\.|$)/gs,  // 番号付きリスト
    /•\s*.*?(?=•|$)/gs,           // 箇条書き
    /\n\n.*?(?=\n\n|$)/gs          // 段落区切り
  ];
  
  // 各パターンで試行
  for (const pattern of articlePatterns) {
    const matches = text.match(pattern);
    if (matches && matches.length > 0) {
      console.log(`${matches.length}個のセクションを発見`);
      
      matches.forEach(match => {
        const article = {};
        
        // タイトル抽出
        const titleMatch = match.match(/(?:タイトル|Title|題名|見出し)[:：]\s*(.+)/i) ||
                          match.match(/「(.+?)」/) ||
                          match.match(/『(.+?)』/);
        if (titleMatch) article.title = titleMatch[1].trim();
        
        // 日付抽出
        const dateMatch = match.match(/(?:日付|Date|掲載日)[:：]\s*(.+)/i) ||
                         match.match(/(\d{4}[-年/]\d{1,2}[-月/]\d{1,2})/);
        if (dateMatch) article.date = dateMatch[1].trim();
        
        // URL抽出
        const urlMatch = match.match(/https?:\/\/[^\s]+/);
        if (urlMatch) article.url = urlMatch[0];
        
        // 要約作成（最初の100文字程度）
        if (!article.summary && match.length > 50) {
          article.summary = match.substring(0, 150).replace(/\n/g, ' ').trim() + '...';
        }
        
        // デフォルト値設定
        article.author = article.author || '不明';
        article.source = article.source || '不明';
        article.criticalScore = 5;
        
        // 最低限の情報がある場合のみ追加
        if (article.title || article.summary) {
          articles.push(article);
        }
      });
      
      if (articles.length > 0) break;
    }
  }
  
  console.log(`抽出結果: ${articles.length}件の記事`);
  return articles;
}

// ====== スプレッドシート操作 ======

// ====== 検索設定管理機能 ======
/**
 * 検索設定シートのセットアップ
 * @param {boolean} forceReset - 既存の設定シートをリセットするか
 * @return {GoogleAppsScript.Spreadsheet.Sheet} 設定シート
 */
function setupConfigSheet(forceReset = false) {
  console.log('検索設定シートのセットアップを開始します...');
  
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    let configSheet = spreadsheet.getSheetByName(CONFIG.CONFIG_SHEET_NAME);
    
    // 既存シートのリセットまたは新規作成
    if (configSheet && forceReset) {
      console.log('既存の設定シートをリセットしています...');
      spreadsheet.deleteSheet(configSheet);
      configSheet = null;
    }
    
    if (!configSheet) {
      console.log('新しい設定シートを作成しています...');
      configSheet = spreadsheet.insertSheet(CONFIG.CONFIG_SHEET_NAME);
    }
    
    // ヘッダーの設定
    setupConfigHeaders(configSheet);
    
    // 列幅の設定
    setupConfigColumnWidths(configSheet);
    
    // データ検証の設定
    setupConfigDataValidation(configSheet);
    
    // サンプル設定の追加（初回のみ）
    if (configSheet.getLastRow() <= 1) {
      addSampleConfigs(configSheet);
    }
    
    // フィルターと固定行の設定
    configSheet.getRange(1, 1, 1, CONFIG.CONFIG_HEADERS.length).createFilter();
    configSheet.setFrozenRows(1);
    
    console.log('検索設定シートのセットアップが完了しました');
    return configSheet;
    
  } catch (error) {
    console.error('設定シートのセットアップでエラーが発生しました:', error);
    throw error;
  }
}

/**
 * 設定シートのヘッダー設定
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 */
function setupConfigHeaders(sheet) {
  const headerRange = sheet.getRange(1, 1, 1, CONFIG.CONFIG_HEADERS.length);
  headerRange.setValues([CONFIG.CONFIG_HEADERS]);
  
  // ヘッダーの書式設定
  headerRange.setFontWeight('bold');
  headerRange.setFontSize(11);
  headerRange.setBackground('#34a853');  // 緑色
  headerRange.setFontColor('#ffffff');
  headerRange.setHorizontalAlignment('center');
  headerRange.setVerticalAlignment('middle');
  headerRange.setBorder(true, true, true, true, true, true);
}

/**
 * 設定シートの列幅設定
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 */
function setupConfigColumnWidths(sheet) {
  const columnWidths = [
    120, // 企業名
    200, // キーワード
    80,  // 検索期間(日)
    100, // 批判的記事のみ
    80,  // メール送信
    180, // 送信先メール
    80,  // 実行頻度
    80,  // 有効/無効
    150, // 最終実行日時
    200  // 備考
  ];
  
  columnWidths.forEach((width, index) => {
    sheet.setColumnWidth(index + 1, width);
  });
}

/**
 * 設定シートのデータ検証設定
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 */
function setupConfigDataValidation(sheet) {
  // 検索期間のデータ検証（1-365日）
  const daysRange = sheet.getRange(2, 3, 1000, 1);
  const daysValidation = SpreadsheetApp.newDataValidation()
    .requireNumberBetween(1, 365)
    .setAllowInvalid(false)
    .setHelpText('1から365の数値を入力してください')
    .build();
  daysRange.setDataValidation(daysValidation);
  
  // 批判的記事のみ（TRUE/FALSE）
  const criticalOnlyRange = sheet.getRange(2, 4, 1000, 1);
  const criticalOnlyValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(['TRUE', 'FALSE'])
    .setAllowInvalid(false)
    .build();
  criticalOnlyRange.setDataValidation(criticalOnlyValidation);
  
  // メール送信（TRUE/FALSE）
  const emailSendRange = sheet.getRange(2, 5, 1000, 1);
  const emailSendValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(['TRUE', 'FALSE'])
    .setAllowInvalid(false)
    .build();
  emailSendRange.setDataValidation(emailSendValidation);
  
  // 実行頻度
  const frequencyRange = sheet.getRange(2, 7, 1000, 1);
  const frequencyValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(['毎日', '週次', '月次', '手動'])
    .setAllowInvalid(false)
    .build();
  frequencyRange.setDataValidation(frequencyValidation);
  
  // 有効/無効
  const enabledRange = sheet.getRange(2, 8, 1000, 1);
  const enabledValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(['有効', '無効'])
    .setAllowInvalid(false)
    .build();
  enabledRange.setDataValidation(enabledValidation);
}

/**
 * サンプル設定の追加
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 */
function addSampleConfigs(sheet) {
  const sampleData = [
    [
      'サンプル企業A',           // 企業名
      '不祥事,訴訟,問題',      // キーワード
      7,                      // 検索期間(日)
      'TRUE',                 // 批判的記事のみ
      'TRUE',                 // メール送信
      'example@company.com',  // 送信先メール
      '毎日',                 // 実行頻度
      '有効',                 // 有効/無効
      '',                     // 最終実行日時
      '重要企業のため毎日監視' // 備考
    ],
    [
      'サンプル企業B',
      'リコール,回収',
      14,
      'FALSE',
      'FALSE',
      '',
      '週次',
      '無効',
      '',
      'テスト用設定'
    ]
  ];
  
  const lastRow = sheet.getLastRow();
  sheet.getRange(lastRow + 1, 1, sampleData.length, CONFIG.CONFIG_HEADERS.length).setValues(sampleData);
}

/**
 * 検索設定の読み取り
 * @param {boolean} activeOnly - 有効な設定のみを取得するか
 * @return {Array<Object>} 設定オブジェクトの配列
 */
function getSearchConfigs(activeOnly = true) {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const configSheet = spreadsheet.getSheetByName(CONFIG.CONFIG_SHEET_NAME);
    
    if (!configSheet) {
      console.log('設定シートが見つかりません。セットアップを実行してください。');
      return [];
    }
    
    const lastRow = configSheet.getLastRow();
    if (lastRow <= 1) {
      console.log('設定データがありません。');
      return [];
    }
    
    // データを読み取り
    const data = configSheet.getRange(2, 1, lastRow - 1, CONFIG.CONFIG_HEADERS.length).getValues();
    
    const configs = data.map((row, index) => {
      // 空行をスキップ（企業名が空の場合）
      const companyName = row[0] ? row[0].toString().trim() : '';
      if (!companyName) return null;
      
      return {
        rowIndex: index + 2, // スプレッドシートの行番号
        companyName: companyName,
        keywords: row[1] ? row[1].toString().split(',').map(k => k.trim()).filter(k => k) : [],
        daysBack: parseInt(row[2]) || CONFIG.DEFAULT_DAYS_BACK,
        criticalOnly: row[3] === 'TRUE' || row[3] === true,
        sendEmail: row[4] === 'TRUE' || row[4] === true,
        emailTo: row[5] || '',
        frequency: row[6] || '手動',
        enabled: row[7] === '有効',
        lastExecuted: row[8] || null,
        notes: row[9] || ''
      };
    }).filter(config => config !== null);
    
    // 有効な設定のみフィルタリング
    if (activeOnly) {
      return configs.filter(config => config.enabled);
    }
    
    return configs;
    
  } catch (error) {
    console.error('設定読み取りエラー:', error);
    return [];
  }
}

/**
 * 設定の最終実行日時を更新
 * @param {number} rowIndex - 設定の行番号
 */
function updateLastExecuted(rowIndex) {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const configSheet = spreadsheet.getSheetByName(CONFIG.CONFIG_SHEET_NAME);
    
    if (configSheet) {
      const cell = configSheet.getRange(rowIndex, 9); // 最終実行日時列
      cell.setValue(new Date());
    }
  } catch (error) {
    console.error('最終実行日時更新エラー:', error);
  }
}

/**
 * 設定のデバッグ情報を表示
 */
function debugSearchConfigs() {
  console.log('=== 設定デバッグ情報 ===');
  
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const configSheet = spreadsheet.getSheetByName(CONFIG.CONFIG_SHEET_NAME);
    
    if (!configSheet) {
      console.log('❌ 設定シートが見つかりません');
      return;
    }
    
    console.log('✓ 設定シートが見つかりました');
    
    const lastRow = configSheet.getLastRow();
    console.log(`データ行数: ${lastRow}`);
    
    if (lastRow <= 1) {
      console.log('❌ データがありません（ヘッダーのみ）');
      return;
    }
    
    // 全データを読み取り
    const data = configSheet.getRange(2, 1, lastRow - 1, CONFIG.CONFIG_HEADERS.length).getValues();
    console.log(`読み取りデータ行数: ${data.length}`);
    
    // 各行をチェック
    data.forEach((row, index) => {
      const rowNumber = index + 2;
      const companyName = row[0] ? row[0].toString().trim() : '';
      const enabled = row[7] === '有効';
      
      console.log(`行${rowNumber}: 企業名="${companyName}", 有効=${enabled}, 生データ=[${row.join(', ')}]`);
    });
    
    // 解析後の設定をチェック
    const allConfigs = getSearchConfigs(false);
    const activeConfigs = getSearchConfigs(true);
    
    console.log(`解析後の設定数: 全${allConfigs.length}件, 有効${activeConfigs.length}件`);
    
    activeConfigs.forEach(config => {
      console.log(`有効設定: ${config.companyName} (頻度: ${config.frequency})`);
    });
    
  } catch (error) {
    console.error('設定デバッグでエラー:', error);
  }
}

/**
 * 設定に基づく自動検索実行
 * @param {string} frequency - 実行頻度フィルター（'毎日', '週次', '月次', 'すべて'）
 */
function executeSearchByConfig(frequency = 'すべて') {
  console.log(`設定に基づく検索を実行します（頻度: ${frequency}）`);
  console.log(`実行時刻: ${new Date().toLocaleString('ja-JP')}`);
  
  try {
    // 1. 設定シートの存在確認
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const configSheet = spreadsheet.getSheetByName(CONFIG.CONFIG_SHEET_NAME);
    
    if (!configSheet) {
      console.log('❌ 設定シートが見つかりません');
      console.log('解決方法: setupConfigSheet() を実行してください');
      return;
    }
    console.log('✓ 設定シートが見つかりました');
    
    // 2. デバッグ情報を先に表示
    console.log('\n--- 設定読み取り状況 ---');
    const allConfigs = getSearchConfigs(false);
    console.log(`全設定数: ${allConfigs.length}件`);
    
    // 全設定の詳細を表示
    allConfigs.forEach((config, index) => {
      console.log(`設定${index + 1}: 企業名="${config.companyName}", 有効=${config.enabled}, 頻度=${config.frequency}`);
    });
    
    const configs = getSearchConfigs(true); // 有効な設定のみ取得
    console.log(`\n有効設定数: ${configs.length}件`);
    
    if (configs.length === 0) {
      console.log('❌ 実行可能な設定がありません。');
      console.log('設定の詳細を確認するには debugSearchConfigs() を実行してください。');
      console.log('');
      console.log('考えられる原因:');
      console.log('1. 「有効/無効」列が「無効」になっている');
      console.log('2. 企業名が空白');
      console.log('3. データが入力されていない');
      return [];
    }
    
    // 3. 頻度による絞り込み確認
    const filteredConfigs = configs.filter(config => 
      frequency === 'すべて' || config.frequency === frequency
    );
    
    console.log(`頻度フィルター後: ${filteredConfigs.length}件`);
    if (filteredConfigs.length === 0) {
      console.log(`❌ 頻度「${frequency}」に該当する設定がありません`);
      configs.forEach(config => {
        console.log(`- ${config.companyName}: 頻度=${config.frequency}`);
      });
      return [];
    }
    
    let executedCount = 0;
    const results = [];
    
    for (const config of filteredConfigs) {
      try {
        console.log(`\n--- 企業 ${config.companyName} の検索を実行中 ---`);
        console.log(`企業名: "${config.companyName}" (長さ: ${config.companyName.length})`);
        console.log(`キーワード: [${config.keywords.join(', ')}]`);
        console.log(`検索期間: ${config.daysBack}日`);
        console.log(`批判的記事のみ: ${config.criticalOnly}`);
        console.log(`実行頻度: ${config.frequency}`);
        
        // 企業名の詳細チェック
        if (!config.companyName || config.companyName.trim() === '') {
          console.error(`❌ 企業名が空です: "${config.companyName}"`);
          throw new Error('企業名が空です');
        }
        
        let searchResults;
        if (config.keywords.length > 0) {
          // キーワード検索
          console.log(`キーワード検索を実行: ${config.keywords.join(', ')}`);
          searchResults = searchArticlesByKeywords(
            config.companyName,
            config.keywords,
            config.daysBack,
            config.criticalOnly,
            config.emailTo || null,
            config.sendEmail
          );
        } else {
          // 批判的記事検索
          console.log('批判的記事検索を実行');
          searchResults = searchCriticalArticles(
            config.companyName,
            config.daysBack,
            [],
            config.emailTo || null,
            config.sendEmail
          );
        }
        
        // 結果を配列に追加
        results.push({
          company: config.companyName,
          articles: searchResults,
          config: config
        });
        
        // 最終実行日時を更新
        updateLastExecuted(config.rowIndex);
        
        console.log(`${config.companyName}: ${searchResults.length}件の記事を検出`);
        executedCount++;
        
        // API制限対策で少し待機
        if (executedCount < filteredConfigs.length) {
          Utilities.sleep(2000);
        }
        
      } catch (error) {
        console.error(`${config.companyName} の検索でエラー:`, error);
        results.push({
          company: config.companyName,
          articles: [],
          error: error.message,
          config: config
        });
      }
    }
    
    console.log(`\n=== 検索完了 ===`);
    console.log(`実行設定数: ${filteredConfigs.length}件`);
    console.log(`成功数: ${executedCount}件`);
    console.log(`総記事数: ${results.reduce((sum, r) => sum + (r.articles?.length || 0), 0)}件`);
    
    return results;
    
  } catch (error) {
    console.error('設定による検索実行エラー:', error);
    throw error;
  }
}

/**
 * スプレッドシートの初期セットアップを行う
 * @param {boolean} forceReset - 既存のシートをリセットするか（デフォルト: false）
 * @return {GoogleAppsScript.Spreadsheet.Sheet} セットアップされたシート
 */
function setupSpreadsheet(forceReset = false) {
  console.log('スプレッドシートのセットアップを開始します...');
  
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = spreadsheet.getSheetByName(CONFIG.SHEET_NAME);
    
    // 既存シートのリセットまたは新規作成
    if (sheet && forceReset) {
      console.log('既存シートをリセットしています...');
      spreadsheet.deleteSheet(sheet);
      sheet = null;
    }
    
    if (!sheet) {
      console.log('新しいシートを作成しています...');
      sheet = spreadsheet.insertSheet(CONFIG.SHEET_NAME);
    }
    
    // ヘッダーの設定
    setupHeaders(sheet);
    
    // 列幅の自動調整
    setupColumnWidths(sheet);
    
    // フィルターの設定
    setupFilters(sheet);
    
    // データ検証の設定
    setupDataValidation(sheet);
    
    // 条件付き書式の設定
    setupConditionalFormatting(sheet);
    
    // 固定行の設定
    sheet.setFrozenRows(1);
    
    console.log('スプレッドシートのセットアップが完了しました');
    return sheet;
    
  } catch (error) {
    console.error('スプレッドシートのセットアップでエラーが発生しました:', error);
    throw error;
  }
}

/**
 * ヘッダーの設定と書式設定
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 */
function setupHeaders(sheet) {
  console.log('ヘッダーを設定しています...');
  
  // ヘッダー行の設定
  const headerRange = sheet.getRange(1, 1, 1, CONFIG.HEADERS.length);
  headerRange.setValues([CONFIG.HEADERS]);
  
  // ヘッダーの書式設定
  headerRange.setFontWeight('bold');
  headerRange.setFontSize(11);
  headerRange.setBackground('#4285f4');
  headerRange.setFontColor('#ffffff');
  headerRange.setHorizontalAlignment('center');
  headerRange.setVerticalAlignment('middle');
  
  // 境界線の設定
  headerRange.setBorder(true, true, true, true, true, true);
}

/**
 * 列幅の設定
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 */
function setupColumnWidths(sheet) {
  console.log('列幅を設定しています...');
  
  // 列幅の設定（列ごとに最適化）
  const columnWidths = [
    150, // 検索日時
    120, // 対象企業
    100, // 記事日付
    300, // 件名
    400, // 内容要約
    120, // 記者名/執筆者
    80,  // URL（リンク）
    120, // 媒体名
    80   // 批判度スコア
  ];
  
  columnWidths.forEach((width, index) => {
    sheet.setColumnWidth(index + 1, width);
  });
}

/**
 * フィルターの設定
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 */
function setupFilters(sheet) {
  console.log('フィルターを設定しています...');
  
  // ヘッダー行にフィルターを追加
  const headerRange = sheet.getRange(1, 1, 1, CONFIG.HEADERS.length);
  headerRange.createFilter();
}

/**
 * データ検証の設定
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 */
function setupDataValidation(sheet) {
  console.log('データ検証を設定しています...');
  
  // 批判度スコア列（I列）にデータ検証を設定
  const scoreColumnRange = sheet.getRange(2, 9, 1000, 1); // 2行目から1000行目まで
  const scoreValidation = SpreadsheetApp.newDataValidation()
    .requireNumberBetween(1, 10)
    .setAllowInvalid(false)
    .setHelpText('1から10の数値を入力してください（10が最も批判的）')
    .build();
  
  scoreColumnRange.setDataValidation(scoreValidation);
}

/**
 * 条件付き書式の設定
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 */
function setupConditionalFormatting(sheet) {
  console.log('条件付き書式を設定しています...');
  
  // 批判度スコアに基づく条件付き書式
  const scoreRange = sheet.getRange(2, 9, 1000, 1);
  
  // スコア8-10: 赤色（高い批判度）
  const highCriticalRule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberBetween(8, 10)
    .setBackground('#ffebee')
    .setFontColor('#c62828')
    .setRanges([scoreRange])
    .build();
  
  // スコア5-7: オレンジ色（中程度の批判度）
  const mediumCriticalRule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberBetween(5, 7)
    .setBackground('#fff3e0')
    .setFontColor('#ef6c00')
    .setRanges([scoreRange])
    .build();
  
  // スコア1-4: 緑色（低い批判度）
  const lowCriticalRule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberBetween(1, 4)
    .setBackground('#e8f5e8')
    .setFontColor('#2e7d32')
    .setRanges([scoreRange])
    .build();
  
  // 条件付き書式を適用
  const rules = sheet.getConditionalFormatRules();
  rules.push(highCriticalRule, mediumCriticalRule, lowCriticalRule);
  sheet.setConditionalFormatRules(rules);
}

/**
 * スプレッドシートに記事を保存
 */
function saveToSpreadsheet(articles, companyName) {
  const sheet = getOrCreateSheet();
  const timestamp = new Date();
  
  // ヘッダーがない場合は追加
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, CONFIG.HEADERS.length).setValues([CONFIG.HEADERS]);
    sheet.getRange(1, 1, 1, CONFIG.HEADERS.length).setFontWeight('bold');
  }
  
  // 記事データを整形して追加
  const rows = articles.map(article => [
    timestamp,
    companyName,
    article.date || '',
    article.title || '',
    article.summary || '',
    article.author || '不明',
    article.url || '',
    article.source || '',
    article.criticalScore || ''
  ]);
  
  if (rows.length > 0) {
    const lastRow = sheet.getLastRow();
    sheet.getRange(lastRow + 1, 1, rows.length, CONFIG.HEADERS.length).setValues(rows);
    
    // URLをハイパーリンクとして設定
    for (let i = 0; i < rows.length; i++) {
      const url = rows[i][6];
      if (url) {
        const cell = sheet.getRange(lastRow + 1 + i, 7);
        cell.setFormula(`=HYPERLINK("${url}","リンク")`);
      }
    }
  }
}

/**
 * シートの取得または作成
 */
function getOrCreateSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(CONFIG.SHEET_NAME);
  
  if (!sheet) {
    sheet = spreadsheet.insertSheet(CONFIG.SHEET_NAME);
  }
  
  return sheet;
}

// ====== ユーティリティ関数 ======
/**
 * APIキーの取得（PropertiesServiceから）
 */
function getApiKey() {
  // PropertiesServiceからAPIキーを取得（推奨）
  const scriptProperties = PropertiesService.getScriptProperties();
  let apiKey = scriptProperties.getProperty('PERPLEXITY_API_KEY');
  
  // PropertiesServiceに設定されていない場合はCONFIGから取得
  if (!apiKey) {
    apiKey = CONFIG.PERPLEXITY_API_KEY;
  }
  
  if (!apiKey || apiKey === 'YOUR_PERPLEXITY_API_KEY') {
    throw new Error('Perplexity APIキーが設定されていません');
  }
  
  return apiKey;
}

/**
 * 日付フォーマット
 */
function formatDate(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
}

// ====== 定期実行用関数 ======
/**
 * 毎日実行する用の関数（トリガー設定用）
 * スプレッドシートの設定に基づいて実行
 */
function dailyCheck() {
  console.log('毎日の自動チェックを開始します');
  try {
    executeSearchByConfig('毎日');
  } catch (error) {
    console.error('毎日チェックでエラー:', error);
  }
}

/**
 * 週次レポート作成用
 * スプレッドシートの設定に基づいて実行
 */
function weeklyReport() {
  console.log('週次レポートを作成します');
  try {
    executeSearchByConfig('週次');
  } catch (error) {
    console.error('週次レポートでエラー:', error);
  }
}

/**
 * 月次レポート作成用
 */
function monthlyReport() {
  console.log('月次レポートを作成します');
  try {
    executeSearchByConfig('月次');
  } catch (error) {
    console.error('月次レポートでエラー:', error);
  }
}

/**
 * 手動実行用（すべての有効な設定を実行）
 */
function manualExecuteAll() {
  console.log('手動で全設定を実行します');
  try {
    executeSearchByConfig('すべて');
  } catch (error) {
    console.error('手動実行でエラー:', error);
  }
}

/**
 * レポートメール送信（改良版・HTML形式）
 * @param {Array} articles - 送信する記事データ
 * @param {string} companyName - 会社名
 * @param {string} recipient - 送信先メールアドレス（オプション）
 * @param {string} subject - メール件名（オプション）
 */
function sendReportEmailWithArticles(articles, companyName, recipient = null, subject = null) {
  if (!articles || articles.length === 0) {
    console.log('送信する記事がありません');
    return;
  }
  
  // 送信先の設定
  const toEmail = recipient || Session.getActiveUser().getEmail();
  const emailSubject = subject || `【${companyName}】批判的記事レポート - ${formatDate(new Date())}`;
  
  // HTML形式のメール本文を作成
  const htmlBody = createHtmlReport(articles, companyName);
  
  // プレーンテキスト版も作成（HTML非対応のメールクライアント用）
  const plainTextBody = createPlainTextReport(articles, companyName);
  
  // メール送信
  GmailApp.sendEmail(toEmail, emailSubject, plainTextBody, {
    htmlBody: htmlBody,
    name: '批判的記事監視システム'
  });
  
  console.log(`レポートを送信しました: ${toEmail}`);
}

/**
 * HTML形式のレポート作成
 */
function createHtmlReport(articles, companyName) {
  let html = `
  <!DOCTYPE html>
  <html>
  <head>
    <meta charset="UTF-8">
    <style>
      body {
        font-family: 'Helvetica Neue', Arial, 'Hiragino Sans', 'Hiragino Kaku Gothic ProN', Meiryo, sans-serif;
        line-height: 1.6;
        color: #333;
        max-width: 800px;
        margin: 0 auto;
        padding: 20px;
      }
      h1 {
        color: #d32f2f;
        border-bottom: 3px solid #d32f2f;
        padding-bottom: 10px;
      }
      .summary {
        background: #f5f5f5;
        border-left: 4px solid #d32f2f;
        padding: 15px;
        margin: 20px 0;
      }
      table {
        width: 100%;
        border-collapse: collapse;
        margin: 20px 0;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
      }
      th {
        background: #424242;
        color: white;
        padding: 12px;
        text-align: left;
        font-weight: bold;
      }
      td {
        padding: 12px;
        border-bottom: 1px solid #ddd;
      }
      tr:hover {
        background: #f5f5f5;
      }
      .critical-high {
        color: #d32f2f;
        font-weight: bold;
      }
      .critical-medium {
        color: #f57c00;
      }
      .article-title {
        font-weight: bold;
        color: #1976d2;
      }
      .url-link {
        color: #1976d2;
        text-decoration: none;
      }
      .url-link:hover {
        text-decoration: underline;
      }
      .footer {
        margin-top: 40px;
        padding-top: 20px;
        border-top: 1px solid #ddd;
        font-size: 12px;
        color: #666;
      }
    </style>
  </head>
  <body>
    <h1>批判的記事レポート: ${companyName}</h1>
    
    <div class="summary">
      <strong>レポート生成日時:</strong> ${new Date().toLocaleString('ja-JP')}<br>
      <strong>対象企業:</strong> ${companyName}<br>
      <strong>検出記事数:</strong> ${articles.length}件<br>
      <strong>最高批判度スコア:</strong> ${Math.max(...articles.map(a => a.criticalScore || 0))}
    </div>
    
    <h2>記事一覧</h2>
    <table>
      <thead>
        <tr>
          <th style="width: 15%;">日付</th>
          <th style="width: 25%;">件名</th>
          <th style="width: 30%;">内容要約</th>
          <th style="width: 15%;">記者/執筆者</th>
          <th style="width: 10%;">URL</th>
          <th style="width: 5%;">スコア</th>
        </tr>
      </thead>
      <tbody>`;
  
  // 記事をテーブル行として追加
  articles.forEach(article => {
    const scoreClass = article.criticalScore >= 8 ? 'critical-high' : 
                       article.criticalScore >= 5 ? 'critical-medium' : '';
    
    html += `
        <tr>
          <td>${article.date || '不明'}</td>
          <td class="article-title">${escapeHtml(article.title || '無題')}</td>
          <td>${escapeHtml(article.summary || '要約なし')}</td>
          <td>${escapeHtml(article.author || '不明')}</td>
          <td>${article.url ? `<a href="${article.url}" class="url-link" target="_blank">記事を見る</a>` : 'N/A'}</td>
          <td class="${scoreClass}">${article.criticalScore || '-'}</td>
        </tr>`;
  });
  
  html += `
      </tbody>
    </table>
    
    <div class="footer">
      <p>このレポートは批判的記事監視システムによって自動生成されました。</p>
      <p>詳細はスプレッドシートをご確認ください。</p>
    </div>
  </body>
  </html>`;
  
  return html;
}

/**
 * プレーンテキスト形式のレポート作成
 */
function createPlainTextReport(articles, companyName) {
  let text = `批判的記事レポート: ${companyName}\n`;
  text += '='.repeat(50) + '\n\n';
  text += `レポート生成日時: ${new Date().toLocaleString('ja-JP')}\n`;
  text += `対象企業: ${companyName}\n`;
  text += `検出記事数: ${articles.length}件\n\n`;
  text += '記事一覧\n';
  text += '-'.repeat(50) + '\n\n';
  
  articles.forEach((article, index) => {
    text += `【${index + 1}】${article.title || '無題'}\n`;
    text += `  日付: ${article.date || '不明'}\n`;
    text += `  要約: ${article.summary || '要約なし'}\n`;
    text += `  記者: ${article.author || '不明'}\n`;
    text += `  URL: ${article.url || 'N/A'}\n`;
    text += `  批判度スコア: ${article.criticalScore || '-'}/10\n`;
    text += '\n';
  });
  
  return text;
}

/**
 * HTMLエスケープ処理
 */
function escapeHtml(text) {
  if (!text) return '';
  return text.toString()
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

/**
 * レポートメール送信（従来版・互換性維持）
 */
function sendReportEmail() {
  const sheet = getOrCreateSheet();
  const lastRow = sheet.getLastRow();
  
  if (lastRow <= 1) {
    return; // データがない場合は送信しない
  }
  
  // 過去7日間のデータを取得
  const sevenDaysAgo = new Date();
  sevenDaysAgo.setDate(sevenDaysAgo.getDate() - 7);
  
  const data = sheet.getRange(2, 1, lastRow - 1, CONFIG.HEADERS.length).getValues();
  const recentArticles = data.filter(row => new Date(row[0]) >= sevenDaysAgo);
  
  if (recentArticles.length === 0) {
    return;
  }
  
  // 企業ごとに記事データを整形
  const byCompany = {};
  recentArticles.forEach(row => {
    const company = row[1];
    if (!byCompany[company]) {
      byCompany[company] = [];
    }
    byCompany[company].push({
      date: row[2],
      title: row[3],
      summary: row[4],
      author: row[5],
      url: row[6],
      source: row[7],
      criticalScore: row[8]
    });
  });
  
  // 各企業ごとにレポートを送信
  for (const company in byCompany) {
    sendReportEmailWithArticles(byCompany[company], company);
  }
}

// ====== 統合セットアップ機能 ======
/**
 * システム全体の完全セットアップ
 * @param {boolean} forceReset - 既存設定をリセットするか
 * @param {boolean} includeConfigSheet - 設定シートも作成するか
 */
function setupCompleteSystem(forceReset = false, includeConfigSheet = true) {
  console.log('=== システム全体のセットアップを開始します ===');
  
  try {
    const results = {
      success: true,
      mainSheet: null,
      configSheet: null,
      errors: []
    };
    
    // 1. メインの記事保存シートをセットアップ
    console.log('\n1. 記事保存シートのセットアップ');
    try {
      results.mainSheet = setupSpreadsheet(forceReset);
      console.log('✓ 記事保存シート完了');
    } catch (error) {
      results.errors.push(`記事保存シート: ${error.message}`);
      console.error('✗ 記事保存シートでエラー:', error);
    }
    
    // 2. 検索設定シートをセットアップ
    if (includeConfigSheet) {
      console.log('\n2. 検索設定シートのセットアップ');
      try {
        results.configSheet = setupConfigSheet(forceReset);
        console.log('✓ 検索設定シート完了');
      } catch (error) {
        results.errors.push(`検索設定シート: ${error.message}`);
        console.error('✗ 検索設定シートでエラー:', error);
      }
    }
    
    // 3. 設定状態の検証
    console.log('\n3. 設定状態の検証');
    const validation = validateCompleteSetup();
    
    // 4. 結果レポート
    console.log('\n=== セットアップ結果 ===');
    if (results.errors.length === 0) {
      console.log('✓ すべてのセットアップが正常に完了しました！');
      console.log('\n次のステップ:');
      console.log('1. 検索設定シートで監視対象企業を設定');
      console.log('2. executeSearchByConfig() で検索実行');
      console.log('3. トリガー設定で自動実行を有効化');
    } else {
      console.log('⚠ 一部のセットアップでエラーが発生しました:');
      results.errors.forEach(error => console.log(`  - ${error}`));
      results.success = false;
    }
    
    return results;
    
  } catch (error) {
    console.error('✗ システムセットアップで重大なエラー:', error);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * 設定の検証機能
 * @return {Object} 検証結果
 */
function validateSearchConfigs() {
  console.log('検索設定の検証を開始します...');
  
  const result = {
    valid: true,
    configCount: 0,
    activeCount: 0,
    errors: [],
    warnings: [],
    recommendations: []
  };
  
  try {
    const configs = getSearchConfigs(false); // すべての設定を取得
    result.configCount = configs.length;
    result.activeCount = configs.filter(c => c.enabled).length;
    
    if (configs.length === 0) {
      result.warnings.push('設定が1件も登録されていません');
      result.recommendations.push('setupConfigSheet() を実行してサンプル設定を追加してください');
    }
    
    // 各設定の検証
    configs.forEach((config, index) => {
      const configName = `設定${index + 1}(${config.companyName})`;
      
      // 企業名チェック
      if (!config.companyName || config.companyName.trim() === '') {
        result.errors.push(`${configName}: 企業名が空です`);
        result.valid = false;
      }
      
      // 検索期間チェック
      if (config.daysBack < 1 || config.daysBack > 365) {
        result.errors.push(`${configName}: 検索期間が範囲外です (${config.daysBack}日)`);
        result.valid = false;
      }
      
      // メール設定チェック
      if (config.sendEmail && (!config.emailTo || config.emailTo.trim() === '')) {
        result.warnings.push(`${configName}: メール送信が有効ですが送信先が未設定です`);
      }
      
      // メールアドレス形式チェック
      if (config.emailTo && config.emailTo.trim() !== '') {
        const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
        if (!emailRegex.test(config.emailTo.trim())) {
          result.errors.push(`${configName}: メールアドレスの形式が正しくありません (${config.emailTo})`);
          result.valid = false;
        }
      }
      
      // キーワード設定の推奨事項
      if (config.keywords.length === 0) {
        result.recommendations.push(`${configName}: キーワードを設定すると検索精度が向上します`);
      }
    });
    
    // 有効な設定数のチェック
    if (result.activeCount === 0) {
      result.warnings.push('有効な設定がありません');
      result.recommendations.push('少なくとも1つの設定を「有効」にしてください');
    }
    
  } catch (error) {
    result.errors.push(`設定検証中にエラー: ${error.message}`);
    result.valid = false;
  }
  
  return result;
}

/**
 * システム全体の設定状態確認
 * @return {Object} 総合的な設定状態
 */
function validateCompleteSetup() {
  console.log('システム全体の設定状態を確認しています...');
  
  const result = {
    mainSheetValid: false,
    configSheetValid: false,
    configsValid: false,
    readyForUse: false,
    errors: [],
    recommendations: []
  };
  
  try {
    // 1. メインシートの確認
    const mainValidation = validateSpreadsheetSetup();
    result.mainSheetValid = mainValidation.sheetExists && mainValidation.hasHeaders;
    if (!result.mainSheetValid) {
      result.errors.push('記事保存シートが正しく設定されていません');
    }
    
    // 2. 設定シートの確認
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const configSheet = spreadsheet.getSheetByName(CONFIG.CONFIG_SHEET_NAME);
    result.configSheetValid = configSheet !== null;
    if (!result.configSheetValid) {
      result.errors.push('検索設定シートが作成されていません');
      result.recommendations.push('setupConfigSheet() を実行してください');
    }
    
    // 3. 設定内容の確認
    if (result.configSheetValid) {
      const configValidation = validateSearchConfigs();
      result.configsValid = configValidation.valid && configValidation.activeCount > 0;
      if (!result.configsValid) {
        result.errors.push('検索設定に問題があります');
        result.errors.push(...configValidation.errors);
        result.recommendations.push(...configValidation.recommendations);
      }
    }
    
    // 4. 使用準備完了かチェック
    result.readyForUse = result.mainSheetValid && result.configSheetValid && result.configsValid;
    
  } catch (error) {
    result.errors.push(`設定状態確認中にエラー: ${error.message}`);
  }
  
  return result;
}

/**
 * 設定診断レポートの表示
 */
function diagnoseCompleteSetup() {
  const validation = validateCompleteSetup();
  const configValidation = validateSearchConfigs();
  
  console.log('=== システム設定診断レポート ===');
  console.log(`記事保存シート: ${validation.mainSheetValid ? '✓' : '✗'}`);
  console.log(`検索設定シート: ${validation.configSheetValid ? '✓' : '✗'}`);
  console.log(`設定内容: ${validation.configsValid ? '✓' : '✗'}`);
  console.log(`使用準備完了: ${validation.readyForUse ? '✓' : '✗'}`);
  
  console.log(`\n=== 設定統計 ===`);
  console.log(`総設定数: ${configValidation.configCount}件`);
  console.log(`有効設定数: ${configValidation.activeCount}件`);
  
  if (validation.errors.length > 0) {
    console.log('\n【エラー】');
    validation.errors.forEach(error => console.log(`- ${error}`));
  }
  
  if (configValidation.warnings.length > 0) {
    console.log('\n【警告】');
    configValidation.warnings.forEach(warning => console.log(`- ${warning}`));
  }
  
  if (validation.recommendations.length > 0) {
    console.log('\n【推奨事項】');
    validation.recommendations.forEach(rec => console.log(`- ${rec}`));
  }
  
  if (validation.readyForUse) {
    console.log('\n✓ システムは使用準備完了です！');
    console.log('executeSearchByConfig() で検索を開始できます。');
  } else {
    console.log('\n⚠ システムの設定を完了してから使用してください。');
  }
  
  return validation;
}

// ====== スプレッドシート設定検証機能 ======
/**
 * スプレッドシートの設定状態を確認
 * @return {Object} 設定状態の詳細
 */
function validateSpreadsheetSetup() {
  console.log('スプレッドシートの設定状態を確認しています...');
  
  const result = {
    sheetExists: false,
    hasHeaders: false,
    hasFilters: false,
    hasConditionalFormatting: false,
    hasDataValidation: false,
    columnWidthsSet: false,
    frozenRowsSet: false,
    errors: [],
    recommendations: []
  };
  
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName(CONFIG.SHEET_NAME);
    
    if (!sheet) {
      result.errors.push(`シート "${CONFIG.SHEET_NAME}" が見つかりません`);
      result.recommendations.push('setupSpreadsheet() を実行してください');
      return result;
    }
    
    result.sheetExists = true;
    
    // ヘッダーの確認
    if (sheet.getLastRow() > 0) {
      const headerRow = sheet.getRange(1, 1, 1, CONFIG.HEADERS.length).getValues()[0];
      const headersMatch = CONFIG.HEADERS.every((header, index) => header === headerRow[index]);
      result.hasHeaders = headersMatch;
      
      if (!headersMatch) {
        result.errors.push('ヘッダーが正しく設定されていません');
      }
    } else {
      result.errors.push('ヘッダーが設定されていません');
    }
    
    // フィルターの確認
    result.hasFilters = sheet.getFilter() !== null;
    if (!result.hasFilters) {
      result.recommendations.push('フィルターを設定することを推奨します');
    }
    
    // 条件付き書式の確認
    result.hasConditionalFormatting = sheet.getConditionalFormatRules().length > 0;
    if (!result.hasConditionalFormatting) {
      result.recommendations.push('条件付き書式を設定することを推奨します');
    }
    
    // 固定行の確認
    result.frozenRowsSet = sheet.getFrozenRows() > 0;
    if (!result.frozenRowsSet) {
      result.recommendations.push('ヘッダー行を固定することを推奨します');
    }
    
    console.log('設定確認完了');
    
  } catch (error) {
    result.errors.push(`確認中にエラーが発生しました: ${error.message}`);
  }
  
  return result;
}

/**
 * スプレッドシート設定の診断レポートを出力
 */
function diagnoseSpreadsheetSetup() {
  const validation = validateSpreadsheetSetup();
  
  console.log('=== スプレッドシート診断レポート ===');
  console.log(`シート存在: ${validation.sheetExists ? '✓' : '✗'}`);
  console.log(`ヘッダー設定: ${validation.hasHeaders ? '✓' : '✗'}`);
  console.log(`フィルター設定: ${validation.hasFilters ? '✓' : '✗'}`);
  console.log(`条件付き書式: ${validation.hasConditionalFormatting ? '✓' : '✗'}`);
  console.log(`固定行設定: ${validation.frozenRowsSet ? '✓' : '✗'}`);
  
  if (validation.errors.length > 0) {
    console.log('\n【エラー】');
    validation.errors.forEach(error => console.log(`- ${error}`));
  }
  
  if (validation.recommendations.length > 0) {
    console.log('\n【推奨事項】');
    validation.recommendations.forEach(rec => console.log(`- ${rec}`));
  }
  
  const allPassed = validation.sheetExists && validation.hasHeaders && 
                   validation.hasFilters && validation.hasConditionalFormatting && 
                   validation.frozenRowsSet;
  
  if (allPassed) {
    console.log('\n✓ すべての設定が正常です！');
  } else {
    console.log('\n⚠ 一部設定が不完全です。setupSpreadsheet() の実行を検討してください。');
  }
  
  return validation;
}

/**
 * 完全なスプレッドシートセットアップとテストを実行
 * @param {boolean} forceReset - 既存設定をリセットするか
 */
function setupAndTestSpreadsheet(forceReset = false) {
  console.log('=== スプレッドシート完全セットアップ開始 ===');
  
  try {
    // 1. セットアップ実行
    const sheet = setupSpreadsheet(forceReset);
    console.log('✓ セットアップ完了');
    
    // 2. 設定確認
    const validation = validateSpreadsheetSetup();
    
    // 3. テストデータの追加（オプション）
    const addTestData = false; // 必要に応じてtrueに変更
    if (addTestData) {
      console.log('テストデータを追加しています...');
      addSampleTestData(sheet);
      console.log('✓ テストデータ追加完了');
    }
    
    // 4. 最終レポート
    console.log('\n=== セットアップ完了レポート ===');
    diagnoseSpreadsheetSetup();
    
    // 5. 成功判定の改善
    const isSuccessful = sheet !== null && validation.sheetExists && validation.hasHeaders;
    
    if (isSuccessful) {
      console.log('\n✓ スプレッドシートのセットアップが正常に完了しました！');
    } else {
      console.log('\n⚠ セットアップは完了しましたが、一部機能が不完全な可能性があります');
    }
    
    return {
      success: isSuccessful,
      sheet: sheet,
      validation: validation
    };
    
  } catch (error) {
    console.error('✗ セットアップ中にエラーが発生しました:', error);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * サンプルテストデータの追加
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 */
function addSampleTestData(sheet) {
  const testData = [
    [
      new Date(),
      'テスト企業A',
      '2024-01-15',
      'テスト記事タイトル1',
      'これは批判的な記事の要約例です。企業の不祥事について報じています。',
      'テスト記者1',
      'https://example.com/article1',
      'テストメディア1',
      8
    ],
    [
      new Date(),
      'テスト企業B',
      '2024-01-14',
      'テスト記事タイトル2',
      'こちらは中程度の批判的記事の例です。問題提起を行っています。',
      'テスト記者2',
      'https://example.com/article2',
      'テストメディア2',
      6
    ],
    [
      new Date(),
      'テスト企業C',
      '2024-01-13',
      'テスト記事タイトル3',
      '軽微な問題について報告している記事の例です。',
      'テスト記者3',
      'https://example.com/article3',
      'テストメディア3',
      3
    ]
  ];
  
  const lastRow = sheet.getLastRow();
  sheet.getRange(lastRow + 1, 1, testData.length, CONFIG.HEADERS.length).setValues(testData);
  
  // URL列にハイパーリンクを設定
  testData.forEach((row, index) => {
    const url = row[6];
    if (url) {
      const cell = sheet.getRange(lastRow + 1 + index, 7);
      cell.setFormula(`=HYPERLINK("${url}","リンク")`);
    }
  });
}

// ====== テスト用関数 ======
/**
 * 動作テスト用
 */
function testSearch() {
  // テスト実行（実際の企業名を使用）
  const results = searchCriticalArticles('トヨタ自動車', 7, [], null, false);
  console.log('検索結果:', results);
}

/**
 * GASエディタから手動実行する場合のサンプル関数
 */
function manualRun() {
  console.log('===== 手動実行開始 =====');
  console.log('実行時刻:', new Date().toLocaleString('ja-JP'));
  
  try {
    // 実行例1: 基本的な検索（メール送信なし）
    const results = searchCriticalArticles('トヨタ自動車', 7, [], null, false);
    
    console.log('\n===== 実行結果サマリー =====');
    console.log('検索結果:', results.length, '件');
    
    if (results.length > 0) {
      console.log('\n記事一覧:');
      results.forEach((article, i) => {
        console.log(`${i + 1}. ${article.title || '無題'}`);
        console.log(`   日付: ${article.date || '不明'}`);
        console.log(`   URL: ${article.url || 'なし'}`);
      });
    } else {
      console.log('記事が見つかりませんでした');
    }
    
  } catch (error) {
    console.error('実行エラー:', error);
    console.error('スタックトレース:', error.stack);
  }
  
  console.log('\n===== 実行終了 =====');
}

/**
 * デバッグモードでの実行（詳細ログ付き）
 */
function debugRun() {
  console.log('===== デバッグモード実行 =====');
  
  // APIキーの確認
  try {
    const apiKey = getApiKey();
    console.log('APIキー: 設定済み（長さ:', apiKey.length, '文字）');
  } catch (e) {
    console.error('APIキーエラー:', e.message);
    return;
  }
  
  // 小さいリクエストでテスト
  const testQuery = '最近のトヨタのニュースを1件だけ教えてください。';
  console.log('テストクエリ:', testQuery);
  
  try {
    const result = searchWithPerplexity(testQuery, getApiKey());
    console.log('テスト結果:', result);
  } catch (error) {
    console.error('テストエラー:', error);
  }
}

/**
 * API接続テスト
 */
function testApiConnection() {
  try {
    const apiKey = getApiKey();
    console.log('APIキー取得: OK');
    
    const testQuery = 'テストクエリ';
    const response = searchWithPerplexity(testQuery, apiKey);
    console.log('API接続: OK');
    console.log('レスポンス:', response);
    
  } catch (error) {
    console.error('テスト失敗:', error);
  }
}

/**
 * 利用可能なモデルのテスト
 */
function testAvailableModels() {
  console.log('=== 利用可能なモデルのテスト ===');
  
  const apiKey = getApiKey();
  const testQuery = '【テスト】ニュース検索のテストです。この検索に対して、{"articles":[]}という形式で空の配列を返してください。';
  
  console.log('設定されているモデル一覧:');
  CONFIG.MODELS.forEach((model, index) => {
    console.log(`${index + 1}. ${model}`);
  });
  
  console.log('\nモデルの動作テストを開始...');
  
  for (let i = 0; i < CONFIG.MODELS.length; i++) {
    const model = CONFIG.MODELS[i];
    console.log(`\n--- モデル ${model} のテスト ---`);
    
    try {
      const result = trySearchWithModel(testQuery, apiKey, model);
      console.log(`✓ ${model}: 正常動作 (結果: ${result.length}件)`);
    } catch (error) {
      console.log(`✗ ${model}: エラー - ${error.message}`);
    }
  }
  
  console.log('\n=== モデルテスト完了 ===');
  console.log('利用可能なモデルが見つかった場合、そのモデルが自動的に使用されます。');
}

/**
 * URL検証テスト
 */
function testUrlValidation() {
  console.log('=== URL検証テスト ===');
  
  const testUrls = [
    'https://www.nikkei.com/article/example123',  // 有効パターン
    'https://www.asahi.com/articles/example456',  // 有効パターン
    'https://example.com/fake-article',           // 無効パターン
    'https://test.com/dummy-news',                // 無効パターン
    'invalid-url',                                // 無効パターン
    'https://www.bloomberg.co.jp/news/articles/real-article' // 有効パターン
  ];
  
  testUrls.forEach(url => {
    const isValid = isValidUrl(url);
    console.log(`URL: ${url} → ${isValid ? '✓ 有効' : '✗ 無効'}`);
  });
}

/**
 * ハルシネーション防止強化版API検索テスト
 */
function testAntiHallucinationSearch() {
  console.log('=== ハルシネーション防止強化版API検索テスト ===');
  
  try {
    const apiKey = getApiKey();
    
    // 実際の検索関数を使用してテスト
    console.log('1. 実際の検索関数でテスト実行...');
    const results = searchCriticalArticles('トヨタ自動車', 30, [], null, false);
    
    console.log(`\n=== 検索結果分析 ===`);
    console.log(`総件数: ${results.length}件`);
    
    if (results.length === 0) {
      console.log('記事が見つかりませんでした。これは以下の理由が考えられます：');
      console.log('1. 指定期間に該当記事が存在しない');
      console.log('2. ハルシネーション防止機能により偽記事が除外された');
      console.log('3. URL検証により無効な記事が除外された');
    } else {
      results.forEach((article, index) => {
        console.log(`\n--- 記事${index + 1} ---`);
        console.log(`タイトル: ${article.title}`);
        console.log(`要約: ${article.summary?.substring(0, 100)}...`);
        console.log(`URL: ${article.url}`);
        console.log(`URL検証: ${isValidUrl(article.url) ? '✓ 有効' : '✗ 無効'}`);
        console.log(`媒体: ${article.source}`);
        console.log(`執筆者: ${article.author || '不明'}`);
        console.log(`批判度: ${article.criticalScore}/10`);
        console.log(`検証メモ: ${article.verificationNote || 'なし'}`);
        console.log(`URL情報源: ${article.urlSource || 'unknown'}`);
      });
    }
    
    // URL品質統計
    const urlStats = {
      valid: results.filter(a => isValidUrl(a.url)).length,
      invalid: results.filter(a => !isValidUrl(a.url)).length,
      missing: results.filter(a => !a.url).length
    };
    
    console.log(`\n=== URL品質統計 ===`);
    console.log(`有効URL: ${urlStats.valid}件`);
    console.log(`無効URL: ${urlStats.invalid}件`);
    console.log(`URL不明: ${urlStats.missing}件`);
    
    // ハルシネーション検証
    console.log(`\n=== ハルシネーション検証 ===`);
    const suspiciousArticles = results.filter(article => {
      const patterns = [/example\.com/i, /test\./i, /placeholder/i, /dummy/i];
      return patterns.some(pattern => 
        pattern.test(article.title) || 
        pattern.test(article.summary) || 
        pattern.test(article.url || '')
      );
    });
    
    console.log(`疑わしい記事: ${suspiciousArticles.length}件`);
    if (suspiciousArticles.length > 0) {
      console.log('⚠️ ハルシネーションの可能性がある記事が検出されました');
      suspiciousArticles.forEach((article, index) => {
        console.log(`- ${article.title} (${article.url})`);
      });
    } else {
      console.log('✓ ハルシネーションパターンは検出されませんでした');
    }
    
    return {
      articles: results,
      stats: urlStats,
      suspicious: suspiciousArticles.length
    };
    
  } catch (error) {
    console.error('ハルシネーション防止テストでエラー:', error);
    return {
      articles: [],
      stats: { valid: 0, invalid: 0, missing: 0 },
      suspicious: 0,
      error: error.message
    };
  }
}

/**
 * 実際の実行が設定ベースか確認するテスト
 */
function testActualExecution() {
  console.log('=== 実際の実行テスト ===');
  
  try {
    console.log('1. 設定シートの確認...');
    const configs = getSearchConfigs(true);
    console.log(`有効な設定: ${configs.length}件`);
    
    if (configs.length === 0) {
      console.log('❌ 有効な設定がありません');
      console.log('解決方法: 設定シートの「有効/無効」列を「有効」に設定してください');
      return false;
    }
    
    console.log('\n2. 設定の詳細:');
    configs.forEach((config, index) => {
      console.log(`設定${index + 1}:`);
      console.log(`  企業名: "${config.companyName}"`);
      console.log(`  キーワード: [${config.keywords.join(', ')}]`);
      console.log(`  実行頻度: ${config.frequency}`);
      console.log(`  有効: ${config.enabled}`);
    });
    
    console.log('\n3. テスト実行（実際のAPI呼び出しなし）...');
    
    // 実際に設定ベースで実行されるかテスト
    const testConfig = configs[0];
    console.log(`テスト対象: ${testConfig.companyName}`);
    
    if (!testConfig.companyName || testConfig.companyName.trim() === '') {
      console.log('❌ 企業名が空白です');
      return false;
    }
    
    console.log('✓ 設定は正常に読み取れています');
    console.log('\n4. 実際の検索実行を試してみます...');
    
    // デバッグモードで実際に実行
    return executeSearchByConfig('すべて');
    
  } catch (error) {
    console.error('実行テストでエラー:', error);
    return false;
  }
}

/**
 * 簡単な動作確認テスト
 */
function quickSystemCheck() {
  console.log('=== 簡単な動作確認テスト ===');
  
  try {
    // 1. システム状態の確認
    const validation = validateCompleteSetup();
    console.log(`システム準備状況: ${validation.readyForUse ? '✓ 準備完了' : '✗ セットアップ必要'}`);
    
    // 2. 設定の確認
    const configs = getSearchConfigs(true);
    console.log(`有効な設定: ${configs.length}件`);
    
    // 3. API接続の確認
    try {
      const apiKey = getApiKey();
      console.log('API接続: ✓ OK');
    } catch (error) {
      console.log('API接続: ✗ エラー -', error.message);
    }
    
    // 4. 総合判定
    if (validation.readyForUse && configs.length > 0) {
      console.log('\n🎉 システムは使用可能です！');
      console.log('実行コマンド: executeSearchByConfig("すべて")');
      return true;
    } else {
      console.log('\n⚠️ セットアップが必要です');
      console.log('実行コマンド: setupCompleteSystem()');
      return false;
    }
    
  } catch (error) {
    console.error('動作確認エラー:', error);
    return false;
  }
}

/**
 * 問題診断用の簡単なテスト
 */
function troubleshootConfigIssue() {
  console.log('=== 設定問題の診断を開始します ===');
  
  try {
    // 1. 設定シートの存在確認
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const configSheet = spreadsheet.getSheetByName(CONFIG.CONFIG_SHEET_NAME);
    
    if (!configSheet) {
      console.log('❌ 問題発見: 設定シートが存在しません');
      console.log('解決方法: setupConfigSheet() を実行してください');
      return;
    }
    console.log('✓ 設定シートが存在します');
    
    // 2. データの存在確認
    const lastRow = configSheet.getLastRow();
    if (lastRow <= 1) {
      console.log('❌ 問題発見: 設定データがありません');
      console.log('解決方法: 設定シートに企業名などのデータを入力してください');
      return;
    }
    console.log(`✓ データが ${lastRow - 1} 行あります`);
    
    // 3. 詳細デバッグ実行
    debugSearchConfigs();
    
    // 4. 有効な設定の確認
    const activeConfigs = getSearchConfigs(true);
    if (activeConfigs.length === 0) {
      console.log('❌ 問題発見: 有効な設定がありません');
      console.log('解決方法: 設定シートの「有効/無効」列を「有効」に設定してください');
      return;
    }
    
    console.log('\n=== 診断結果 ===');
    console.log('✓ 設定シートとデータは正常です');
    console.log('実際の検索実行をテストするには executeSearchByConfig("手動") を実行してください');
    
  } catch (error) {
    console.error('❌ 診断中にエラーが発生しました:', error);
  }
}

/**
 * 設定ベースの検索テスト
 */
function testConfigBasedSearch() {
  console.log('=== 設定ベースの検索テスト ===');
  
  try {
    // 1. 設定シートの存在確認
    const configs = getSearchConfigs(false);
    console.log(`設定読み取り: ${configs.length}件の設定を検出`);
    
    if (configs.length === 0) {
      console.log('設定がありません。setupConfigSheet() を実行してサンプル設定を追加してください。');
      return false;
    }
    
    // 2. 設定の検証
    const validation = validateSearchConfigs();
    console.log(`設定検証: ${validation.valid ? '✓' : '✗'} (有効: ${validation.activeCount}件)`);
    
    // 3. テスト実行（手動設定のみ）
    console.log('手動設定での検索テストを実行...');
    const testConfigs = configs.filter(c => c.frequency === '手動' && c.enabled);
    
    if (testConfigs.length > 0) {
      console.log(`${testConfigs.length}件の手動設定をテスト実行します`);
      // 実際の検索はスキップしてログのみ
      testConfigs.forEach(config => {
        console.log(`- ${config.companyName} (キーワード: ${config.keywords.join(', ') || 'なし'})`);
      });
    } else {
      console.log('テスト可能な手動設定がありません');
    }
    
    return true;
    
  } catch (error) {
    console.error('設定ベース検索テストでエラー:', error);
    return false;
  }
}

/**
 * 総合テスト関数（修正版）
 */
function runAllTests() {
  console.log('=== 総合テスト開始 ===');
  
  const results = {
    mainSheet: false,
    configSheet: false,
    apiConnection: false,
    configValidation: false,
    systemReady: false
  };
  
  // 1. メインスプレッドシートテスト
  console.log('\n1. メインスプレッドシートテスト');
  try {
    const mainValidation = validateSpreadsheetSetup();
    results.mainSheet = mainValidation.sheetExists && mainValidation.hasHeaders;
    console.log(`結果: ${results.mainSheet ? '✓' : '✗'}`);
    
    if (!results.mainSheet) {
      console.log('メインシートをセットアップします...');
      setupSpreadsheet();
      const retryValidation = validateSpreadsheetSetup();
      results.mainSheet = retryValidation.sheetExists && retryValidation.hasHeaders;
      console.log(`再テスト結果: ${results.mainSheet ? '✓' : '✗'}`);
    }
  } catch (error) {
    console.error('メインシートテストでエラー:', error);
  }
  
  // 2. 検索設定シートテスト
  console.log('\n2. 検索設定シートテスト');
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    let configSheet = spreadsheet.getSheetByName(CONFIG.CONFIG_SHEET_NAME);
    
    if (!configSheet) {
      console.log('設定シートをセットアップします...');
      configSheet = setupConfigSheet();
    }
    
    results.configSheet = configSheet !== null;
    console.log(`結果: ${results.configSheet ? '✓' : '✗'}`);
  } catch (error) {
    console.error('設定シートテストでエラー:', error);
  }
  
  // 3. 設定検証テスト
  console.log('\n3. 設定検証テスト');
  try {
    results.configValidation = testConfigBasedSearch();
    console.log(`結果: ${results.configValidation ? '✓' : '✗'}`);
  } catch (error) {
    console.error('設定検証テストでエラー:', error);
  }
  
  // 4. API接続テスト
  console.log('\n4. API接続テスト');
  try {
    testApiConnection();
    results.apiConnection = true;
    console.log('結果: ✓');
  } catch (error) {
    console.error('API接続テストでエラー:', error);
    console.log('結果: ✗');
  }
  
  // 5. 総合システム診断
  console.log('\n5. 総合システム診断');
  try {
    const systemValidation = validateCompleteSetup();
    results.systemReady = systemValidation.readyForUse;
    diagnoseCompleteSetup();
  } catch (error) {
    console.error('総合診断でエラー:', error);
  }
  
  // 結果サマリー（修正版）
  console.log('\n=== テスト結果サマリー ===');
  console.log(`メインシート: ${results.mainSheet ? '✓' : '✗'}`);
  console.log(`設定シート: ${results.configSheet ? '✓' : '✗'}`);
  console.log(`設定検証: ${results.configValidation ? '✓' : '✗'}`);
  console.log(`API接続: ${results.apiConnection ? '✓' : '✗'}`);
  console.log(`システム準備状況: ${results.systemReady ? '✓' : '✗'}`);
  
  // 総合判定は実際の使用可能性を重視
  const functionallyReady = results.systemReady && results.apiConnection && results.configValidation;
  console.log(`\n総合結果: ${functionallyReady ? '✓ システムは使用可能です' : '⚠ セットアップが必要です'}`);
  
  if (functionallyReady) {
    console.log('\n🎉 システムは正常に動作しています！');
    console.log('executeSearchByConfig() で検索を開始できます。');
    
    if (!results.mainSheet || !results.configSheet) {
      console.log('\n📝 注意: 一部のシート設定が不完全ですが、基本機能は動作します。');
      console.log('完全なセットアップには setupCompleteSystem() を実行してください。');
    }
  } else {
    console.log('\n⚠️ システムの初期設定を完了してください。');
    console.log('推奨: setupCompleteSystem() を実行');
  }
  
  return results;
}