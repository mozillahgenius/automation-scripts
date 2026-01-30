/**
 * Perplexity API連携 情報収集・分析システム
 * 
 * 使用方法:
 * 1. スプレッドシートに「シート1_検索設定」「シート2_規約チェック」を作成
 * 2. スクリプトプロパティにPERPLEXITY_API_KEYを設定
 * 3. トリガーで定期実行を設定
 */

// ==================== 設定 ====================
const CONFIG = {
  PERPLEXITY_API_KEY: PropertiesService.getScriptProperties().getProperty('PERPLEXITY_API_KEY'),
  PERPLEXITY_API_URL: 'https://api.perplexity.ai/chat/completions',
  
  // Grok-4 AI API の設定
  GROK_API_KEY: PropertiesService.getScriptProperties().getProperty('GROK_API_KEY'),
  GROK_API_URL: 'https://api.x.ai/v1/chat/completions',
  
  EMAIL_RECIPIENTS: 'your-email@example.com,channel@workspace.slack.com', // レポート送信先メールアドレス
  SHEET_NAMES: {
    SEARCH: 'シート1_検索設定',
    VIOLATION_CHECK: 'シート2_規約チェック'
  }
};

// ==================== メイン処理 ====================

/**
 * 書き込み検索のみを実行してメール送信（個別実行用）
 */
function executeKeywordSearchOnly() {
  try {
    Logger.log('=== 書き込み検索のみを実行開始 ===');
    
    // 初期チェック
    if (!CONFIG.PERPLEXITY_API_KEY) {
      throw new Error('PERPLEXITY_API_KEYが設定されていません。スクリプトプロパティを確認してください。');
    }
    
    // キーワード検索レポート実行
    const searchReport = executeKeywordSearch();
    
    if (searchReport) {
      sendSearchReport(searchReport);
      Logger.log('書き込み検索レポートの送信が完了しました');
    } else {
      Logger.log('検索レポートが生成されませんでした');
    }
    
  } catch (error) {
    Logger.log('書き込み検索エラー: ' + error.toString());
    sendErrorNotification(error, '書き込み検索');
  }
}

/**
 * 規約違反チェックのみを実行してメール送信（個別実行用）
 */
function executeViolationCheckOnly() {
  try {
    Logger.log('=== 規約違反チェックのみを実行開始 ===');
    
    // 初期チェック
    if (!CONFIG.PERPLEXITY_API_KEY) {
      throw new Error('PERPLEXITY_API_KEYが設定されていません。スクリプトプロパティを確認してください。');
    }
    
    // 規約違反チェックレポート実行
    const violationReport = executeViolationCheck();
    
    if (violationReport) {
      sendViolationReport(violationReport);
      Logger.log('規約違反チェックレポートの送信が完了しました');
    } else {
      Logger.log('規約違反レポートが生成されませんでした');
    }
    
  } catch (error) {
    Logger.log('規約違反チェックエラー: ' + error.toString());
    sendErrorNotification(error, '規約違反チェック');
  }
}

/**
 * メイン実行関数（両方同時実行・トリガーで実行）
 */
function executeMainProcess() {
  try {
    // 初期チェック
    if (!CONFIG.PERPLEXITY_API_KEY) {
      throw new Error('PERPLEXITY_API_KEYが設定されていません。スクリプトプロパティを確認してください。');
    }
    
    // シート1: キーワード検索レポート
    let searchReport = '';
    try {
      searchReport = executeKeywordSearch();
    } catch (error) {
      Logger.log('検索レポートエラー: ' + error.toString());
      searchReport = '<h2>🔍 キーワード検索レポート</h2>\n<p style="color: red;">エラー: ' + error.toString() + '</p>\n';
    }
    
    // シート2: 規約違反チェックレポート
    let violationReport = '';
    try {
      violationReport = executeViolationCheck();
    } catch (error) {
      Logger.log('規約チェックエラー: ' + error.toString());
      violationReport = '<h2>⚠️ 利用規約違反チェックレポート</h2>\n<p style="color: red;">エラー: ' + error.toString() + '</p>\n';
    }
    
    // レポートを個別にメール送信
    if (searchReport) {
      sendSearchReport(searchReport);
    }
    
    if (violationReport) {
      sendViolationReport(violationReport);
    }
    
    Logger.log('処理が正常に完了しました');
  } catch (error) {
    Logger.log('エラーが発生しました: ' + error.toString());
    sendErrorNotification(error);
  }
}

// ==================== シート1: キーワード検索 ====================

/**
 * キーワード検索を実行
 */
function executeKeywordSearch() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAMES.SEARCH);
  
  if (!sheet) {
    createSearchSheet();
    Logger.log('シート1_検索設定を作成しました。次回実行時から動作します。');
    return '<h2>🔍 キーワード検索レポート</h2>\n<p>シート1_検索設定を作成しました。設定を入力してから再度実行してください。</p>\n';
  }
  
  const data = sheet.getDataRange().getValues();
  
  // データが存在しない、またはヘッダーのみの場合
  if (!data || data.length <= 1) {
    Logger.log('検索設定が見つかりません');
    return '<h2>🔍 キーワード検索レポート</h2>\n<p>検索設定が入力されていません。シート1_検索設定に設定を入力してください。</p>\n';
  }
  
  const headers = data[0];
  const searchResults = [];
  
  // ヘッダー行をスキップして処理
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row || !row[0] || row[0] === '') continue; // 空行スキップ
    
    const searchConfig = {
      keyword: row[0] || '',
      instructions: row[1] || '',
      mediaName: row[2] || '',
      mediaURL: row[3] || '',
      daysBack: row[4] || 7, // デフォルトは7日前まで
      isActive: row[5] === true || row[5] === 'TRUE' || row[5] === 'true'
    };
    
    if (!searchConfig.isActive) {
      Logger.log(`スキップ: ${searchConfig.keyword} (無効化されています)`);
      continue;
    }
    
    if (!searchConfig.keyword) {
      Logger.log('キーワードが空のためスキップ');
      continue;
    }
    
    try {
      const result = performSearch(searchConfig);
      if (result) {
        searchResults.push(result);
      }
    } catch (error) {
      Logger.log(`検索エラー (${searchConfig.keyword}): ${error.toString()}`);
      searchResults.push({
        keyword: searchConfig.keyword,
        media: searchConfig.mediaName,
        content: 'エラー: 検索に失敗しました',
        error: error.toString(),
        timestamp: new Date()
      });
    }
    
    // API制限対策として待機
    Utilities.sleep(2000);
  }
  
  if (searchResults.length === 0) {
    return '<h2>🔍 キーワード検索レポート</h2>\n<p>有効な検索設定が見つかりませんでした。</p>\n';
  }
  
  return formatSearchReport(searchResults);
}

/**
 * Perplexity APIで検索を実行
 */
function performSearch(config) {
  if (!CONFIG.PERPLEXITY_API_KEY) {
    throw new Error('APIキーが設定されていません');
  }
  
  const prompt = buildSearchPrompt(config);
  
  const payload = {
    model: 'sonar-pro',
    messages: [
      {
        role: 'system',
        content: 'あなたは指定された媒体とURLを中心に情報を収集し、詳細なレポートを作成する専門家です。'
      },
      {
        role: 'user',
        content: prompt
      }
    ],
    temperature: 0.2,
    max_tokens: 2000
  };
  
  // search_domain_filterは、有効なURLがある場合のみ追加
  if (config.mediaURL && config.mediaURL.trim() !== '') {
    const domain = extractDomain(config.mediaURL);
    if (domain) {
      payload.search_domain_filter = [domain];
    }
  }
  
  const options = {
    method: 'post',
    headers: {
      'Authorization': `Bearer ${CONFIG.PERPLEXITY_API_KEY}`,
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  
  try {
    const response = UrlFetchApp.fetch(CONFIG.PERPLEXITY_API_URL, options);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();
    
    if (responseCode !== 200) {
      Logger.log(`API Error: Status ${responseCode}, Response: ${responseText}`);
      throw new Error(`API returned status ${responseCode}`);
    }
    
    const result = JSON.parse(responseText);
    
    if (result.choices && result.choices[0] && result.choices[0].message) {
      const daysBack = parseInt(config.daysBack) || 7;
      const targetDate = new Date();
      targetDate.setDate(targetDate.getDate() - daysBack);
      
      return {
        keyword: config.keyword,
        media: config.mediaName,
        content: result.choices[0].message.content || '結果が空でした',
        citations: result.citations || [],
        searchPeriod: `過去${daysBack}日間（${targetDate.toLocaleDateString('ja-JP')}以降）`,
        timestamp: new Date()
      };
    } else {
      throw new Error('APIレスポンスの形式が不正です');
    }
  } catch (error) {
    Logger.log('検索エラー: ' + error.toString());
    return {
      keyword: config.keyword,
      media: config.mediaName,
      content: 'エラー: 検索に失敗しました',
      error: error.toString(),
      timestamp: new Date()
    };
  }
}

/**
 * 検索用プロンプトを構築（個別投稿の詳細取得用）
 */
function buildSearchPrompt(config) {
  // 期間の計算
  const daysBack = parseInt(config.daysBack) || 7;
  const targetDate = new Date();
  targetDate.setDate(targetDate.getDate() - daysBack);
  const dateString = targetDate.toISOString().split('T')[0]; // YYYY-MM-DD形式
  
  let prompt = `以下の条件で個別の投稿を具体的に収集してください：\n\n`;
  prompt += `キーワード: ${config.keyword}\n`;
  prompt += `検索期間: ${dateString}以降（過去${daysBack}日間）\n`;
  
  if (config.instructions) {
    prompt += `追加条件: ${config.instructions}\n`;
  }
  
  if (config.mediaName && config.mediaURL) {
    prompt += `\n重要: 以下の媒体を中心に検索してください：\n`;
    prompt += `媒体名: ${config.mediaName}\n`;
    prompt += `URL: ${config.mediaURL}\n`;
  }
  
  prompt += `\n以下の形式で、個別の投稿をリスト化してください：\n\n`;
  prompt += `## 検索結果（個別投稿）\n\n`;
  prompt += `### 投稿 1\n`;
  prompt += `- **投稿日時**: YYYY-MM-DD HH:MM\n`;
  prompt += `- **投稿内容**: （具体的な投稿テキスト全文）\n`;
  prompt += `- **投稿 URL**: （直接リンク）\n`;
  prompt += `- **関連性**: なぜこの投稿がキーワードに関連しているのか簡潔に説明\n\n`;
  prompt += `### 投稿 2\n`;
  prompt += `（同様の形式で続く）\n\n`;
  prompt += `注意事項：\n`;
  prompt += `- キーワードに関連する可能性がある投稿であれば、断定的でなくても含めてください\n`;
  prompt += `- 投稿内容は全文を正確に転記し、省略しないでください\n`;
  prompt += `- URLは個別投稿への直接リンクを提供してください\n`;
  
  return prompt;
}

// ==================== シート2: 規約違反チェック ====================

/**
 * 違反チェックを実行
 */
function executeViolationCheck() {
  try {
    Logger.log('規約違反チェックを開始します');
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.VIOLATION_CHECK);
  
  if (!sheet) {
      throw new Error(`${CONFIG.SHEET_NAMES.VIOLATION_CHECK}シートが見つかりません。シートを作成してください。`);
  }
  
  const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    // 必須列のインデックスを取得（簡素化版）
    const accountCol = headers.indexOf('アカウント名');
    const keywordsCol = headers.indexOf('チェックキーワード（カンマ区切り）');
    const tosUrlCol = headers.indexOf('利用規約URL');
    const activeCol = headers.indexOf('有効/無効');
    
    if (accountCol === -1 || keywordsCol === -1 || tosUrlCol === -1 || activeCol === -1) {
      throw new Error('シートのヘッダーが不正です。シートを再作成してください。');
    }
    
    const violationReport = [];
    
    // 各行を処理
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
      if (!row[accountCol] || !row[activeCol]) continue;
    
    const checkConfig = {
        account: row[accountCol].toString().trim(),
        keywords: row[keywordsCol] ? row[keywordsCol].toString().split(',').map(k => k.trim()) : [],
        tosURL: row[tosUrlCol] ? row[tosUrlCol].toString().trim() : '',
        platform: 'Yahoo!ファイナンス掲示板', // デフォルト
        platformURL: `https://finance.yahoo.co.jp/cm/personal/history/comment?user=${row[accountCol]}&sort=2`, // 自動生成
        companyName: '(株)ネクスグループ' // デフォルト企業名
      };
      
      Logger.log(`違反チェック処理中: ${checkConfig.account}`);
      
      const result = checkViolations(checkConfig);
      if (result) {
        violationReport.push(result);
      }
    }
    
    if (violationReport.length === 0) {
      Logger.log('有効な違反チェック設定がありませんでした');
      return null;
    }
    
    const formattedReport = formatViolationReport(violationReport);
    Logger.log('規約違反レポートを生成しました');
    
    return formattedReport;
    
    } catch (error) {
    Logger.log(`規約違反チェックエラー: ${error.toString()}`);
    sendErrorNotification(error, '規約違反チェック');
    return null;
  }
}

/**
 * 違反チェックのメイン処理
 */
function checkViolations(config) {
  try {
    let termsContent = '';
    let userPosts = '';
    let analysisResult = '';
    
    // 利用規約を取得
    try {
      termsContent = fetchTermsOfService(config.tosURL);
    } catch (tosError) {
      Logger.log(`規約取得エラー: ${tosError.toString()}`);
      termsContent = getGenericTermsTemplate() + `\n\n※ 注意: 元のURL(${config.tosURL})からの取得に失敗したため、一般的な規約を使用しています。`;
    }
    
    // ユーザー投稿を収集（単一URLから全ての投稿を収集）
    try {
      userPosts = collectUserPosts(config);
    } catch (postsError) {
      Logger.log(`投稿収集エラー: ${postsError.toString()}`);
      userPosts = '投稿の収集に失敗しました。エラー: ' + postsError.toString();
    }
    
    // 違反分析を実行
    if (termsContent && userPosts && userPosts !== 'スクレイピングで投稿を取得できませんでした。') {
      try {
        analysisResult = analyzeViolationsBySection(termsContent, userPosts, config);
      } catch (analysisError) {
        Logger.log(`分析エラー: ${analysisError.toString()}`);
        analysisResult = '違反分析に失敗しました。エラー: ' + analysisError.toString();
      }
    } else {
      analysisResult = '利用規約または投稿の取得に失敗したため、分析をスキップしました。';
    }
    
    // 通報テンプレートを生成
    let reportTemplates = [];
    if (userPosts && userPosts.includes('### 投稿')) {
      // 各投稿からテンプレートを生成
      const postMatches = userPosts.match(/### 投稿 \d+[\s\S]*?(?=### 投稿 \d+|## |$)/g);
      if (postMatches) {
        postMatches.slice(0, 5).forEach(postText => {
          // 簡易的なパース
          const post = {
            datetime: (postText.match(/投稿日時: ([^\n]+)/) || ['', '不明'])[1],
            userName: config.account,
            content: (postText.match(/投稿内容: ([^\n]+)/) || ['', ''])[1],
            matchedKeyword: (postText.match(/マッチしたキーワード: ([^\n]+)/) || ['', ''])[1],
            postUrl: (postText.match(/投稿URL: ([^\n]+)/) || ['', ''])[1]
          };
          
          if (post.matchedKeyword && post.content) {
            reportTemplates.push(generateReportTemplate(post));
          }
        });
      }
    }
    
    return {
      account: config.account,
      platform: config.platform,
      termsContent: termsContent,
      userPosts: userPosts,
      analysis: analysisResult,
      reportTemplates: reportTemplates
    };
    
  } catch (error) {
    Logger.log(`違反チェックエラー (${config.account}): ${error.toString()}`);
    return {
      account: config.account,
      platform: config.platform,
      error: `処理中にエラーが発生しました: ${error.toString()}`
    };
  }
}

/**
 * 規約URLから規約内容を取得し、項目別に整理（改善版）
 */
function fetchTermsOfService(tosURL) {
  if (!tosURL || tosURL.trim() === '') {
    Logger.log('規約URLが空のため、一般的な規約で分析します');
    return getGenericTermsTemplate();
  }
  
  try {
    Logger.log(`規約取得を開始: ${tosURL}`);
    
    const prompt = `以下のURLから利用規約の内容を取得し、項目別に整理してください。
URLにアクセスできない場合は、そのプラットフォームの一般的な利用規約を基に分析してください。

URL: ${tosURL}

以下の形式で出力してください：

## 禁止事項
[項目1: 名誉毀損・中傷]
- 他のユーザーや第三者を中傷、誹謗中傷する内容
- 虚偽の情報で他者の評判を傷つける行為
- 判定基準: 特定個人への攻撃的発言、虚偽の事実の流布

[項目2: ハラスメント・いじめ]
- 特定の個人やグループへの繰り返しの嫌がらせ行為
- 脅迫、脅迫的なメッセージの送信
- 判定基準: 継続的な攻撃、脅迫的言動

[項目3: 著作権侵害]
- 許可なく他者のコンテンツを複製・配布
- コピーライト保護された素材の無断使用
- 判定基準: 引用範囲を超えた複製、出典明記なし

[項目4: スパム・迷惑行為]
- 同じ内容の大量投稿、無関係な広告投稿
- 詐欺的な情報や偽情報の拡散
- 判定基準: 繰り返し投稿、誤解を招く内容

[項目5: 差別・ヘイトスピーチ]
- 人種、性別、宗教、国籍等に基づく差別的発言
- 特定グループへの憎悪や偏見を助長する内容
- 判定基準: 属性に基づく一般化、攻撃的表現

## ペナルティ
- 警告: 初回違反時の注意喚起
- 一時停止: 1日〜30日のアカウント利用停止
- アカウント削除: 重大な違反や繰り返し違反時
- 投稿削除: 違反コンテンツの即座削除
- 機能制限: コメント、メッセージ等の一部機能制限`;
    
    const payload = {
      model: 'sonar-pro',
    messages: [
      {
        role: 'system',
          content: 'あなたは利用規約を分析し、項目別に整理する法務専門家です。URLにアクセスできない場合は、そのプラットフォームの一般的な規約を基に分析します。'
      },
      {
        role: 'user',
          content: prompt
      }
    ],
      temperature: 0.1,
      max_tokens: 3000
  };
  
  const options = {
    method: 'post',
    headers: {
      'Authorization': `Bearer ${CONFIG.PERPLEXITY_API_KEY}`,
      'Content-Type': 'application/json'
    },
      payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  
    const response = UrlFetchApp.fetch(CONFIG.PERPLEXITY_API_URL, options);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();
    
    Logger.log(`APIレスポンスコード: ${responseCode}`);
    
    if (responseCode !== 200) {
      Logger.log(`APIエラー: ${responseText}`);
      throw new Error(`APIエラー: ステータスコード ${responseCode}`);
    }
    
    const result = JSON.parse(responseText);
    
    if (result.choices && result.choices[0] && result.choices[0].message) {
      Logger.log('規約内容の取得に成功しました');
      return result.choices[0].message.content;
    } else {
      Logger.log('APIレスポンスの形式が不正です');
      throw new Error('APIレスポンスの形式が不正です');
    }
    
  } catch (error) {
    Logger.log(`規約取得エラー: ${error.toString()}`);
    Logger.log('フォールバック: 一般的な規約テンプレートを使用します');
    
    // フォールバック: 一般的な規約テンプレートを使用
    return getGenericTermsTemplate() + `

※ 注意: 元のURL(${tosURL})からの規約取得に失敗したため、一般的な規約で分析しています。エラー: ${error.toString()}`;
  }
}

/**
 * 一般的な規約テンプレートを取得
 */
function getGenericTermsTemplate() {
  return `## 一般的なソーシャルメディア利用規約

## 禁止事項
[項目1: 名誉毀損・中傷]
- 他のユーザーや第三者を中傷、誹謗中傷する内容の投稿
- 虚偽の情報で他者の評判を傷つける行為
- 判定基準: 特定個人への攻撃的発言、虚偽の事実の流布

[項目2: ハラスメント・いじめ]
- 特定の個人やグループへの繰り返しの嫌がらせ行為
- 脅迫、脅迫的なメッセージの送信
- 判定基準: 継続的な攻撃、脅迫的言動

[項目3: 著作権侵害]
- 許可なく他者のコンテンツを複製・配布
- コピーライト保護された素材の無断使用
- 判定基準: 引用範囲を超えた複製、出典明記なし

[項目4: スパム・迷惑行為]
- 同じ内容の大量投稿、無関係な広告投稿
- 詐欺的な情報や偽情報の拡散
- 判定基準: 繰り返し投稿、誤解を招く内容

[項目5: 差別・ヘイトスピーチ]
- 人種、性別、宗教、国籍等に基づく差別的発言
- 特定グループへの憎悪や偏見を助長する内容
- 判定基準: 属性に基づく一般化、攻撃的表現

[項目6: 暴力的コンテンツ]
- 物理的暴力を美化、助長する内容
- 自傷や自殺を助長する内容
- 判定基準: 暴力的表現、危険行為の推奨

[項目7: 成人コンテンツ]
- 未成年者に不適切な性的コンテンツ
- ヌードや部分的ヌードの投稿
- 判定基準: 性的な内容、露出的な表現

## ペナルティ
- 警告: 初回違反時の注意喚起
- 投稿削除: 違反コンテンツの即座削除
- 一時停止: 1日〜30日のアカウント利用停止
- 機能制限: コメント、メッセージ等の一部機能制限
- アカウント削除: 重大な違反や繰り返し違反時の永久停止`;
}

/**
 * Yahoo!ファイナンス掲示板からユーザー投稿を収集（キーワードベース抽出）
 */
function scrapeUserPostsFromYahoo(config) {
  try {
    Logger.log(`ウェブスクレイピング開始: ${config.platformURL}`);
    
    let fullContent = '';
    let currentUrl = config.platformURL;
    let pageCount = 0;
    const maxPages = 10; // 最大ページ数を制限
    
    while (currentUrl && pageCount < maxPages) {
      pageCount++;
      Logger.log(`ページ${pageCount}を取得中: ${currentUrl}`);
      
      const pageContent = fetchPageContent(currentUrl);
      
      if (!pageContent) {
        throw new Error(`ページ${pageCount}のコンテンツ取得に失敗しました`);
      }
      
      fullContent += pageContent;
      
      // 次のページURLを抽出（複数パターンで試行）
      let nextUrl = null;
      
      // パターン1: data-cl-paramsを持つリンク
      const nextUrlMatch1 = pageContent.match(/<a\s+href=\"([^\"]+)\"\s+data-cl-params=\"_cl_link:ne[^\"]*\">次のページ<\/a>/);
      if (nextUrlMatch1) {
        nextUrl = nextUrlMatch1[1];
      }
      
      // パターン2: classを持つリンク
      if (!nextUrl) {
        const nextUrlMatch2 = pageContent.match(/<a\s+[^>]*href=\"([^\"]+)\"[^>]*>次のページ<\/a>/);
        if (nextUrlMatch2) {
          nextUrl = nextUrlMatch2[1];
        }
      }
      
      // パターン3: next_post_dateパラメータを含むURL
      if (!nextUrl) {
        const nextUrlMatch3 = pageContent.match(/href=\"([^\"]*next_post_date=\d+[^\"]*)\"/);
        if (nextUrlMatch3) {
          nextUrl = nextUrlMatch3[1];
        }
      }
      
      if (nextUrl) {
        // URLのデコードと正規化
        nextUrl = nextUrl.replace(/&amp;/g, '&');
        if (!nextUrl.startsWith('http')) {
          currentUrl = 'https://finance.yahoo.co.jp' + nextUrl;
        } else {
          currentUrl = nextUrl;
        }
      } else {
        currentUrl = null;
      }
      
      if (currentUrl) {
        Logger.log(`次のページ検出: ${currentUrl}`);
        Utilities.sleep(1000); // レート制限回避のための待機
      }
    }
    
    Logger.log(`複数ページ取得完了: ${pageCount}ページ, 総文字数: ${fullContent.length}`);
    
    // commentBoxの数を確認
    const commentBoxMatches = fullContent.match(/<li\s+class=\"commentBox\"[^>]*>/gi);
    const commentBoxCount = commentBoxMatches ? commentBoxMatches.length : 0;
    Logger.log(`検出されたcommentBox数: ${commentBoxCount}`);
    
    // キーワードベースで投稿を抽出
    const posts = extractPostsByKeywords(fullContent, config);
    
    Logger.log(`スクレイピング完了: ${posts.length}件の投稿を取得`);
    
    return formatScrapedPosts(posts);
    
  } catch (error) {
    Logger.log(`ウェブスクレイピングエラー: ${error.toString()}`);
    
    // フォールバック: Perplexity APIを使用
    Logger.log('フォールバック: Perplexity APIで投稿収集を試みます');
    return collectUserPosts(config);
  }
}

/**
 * キーワードベースで投稿を抽出（前後300文字）
 */
function extractPostsByKeywords(htmlContent, config) {
  const posts = [];
  
  try {
    // HTMLからテキストを抽出（タグは保持して後で解析）
    const textWithoutScript = htmlContent.replace(/<script[\s\S]*?<\/script>/gi, '');
    const textWithoutStyle = textWithoutScript.replace(/<style[\s\S]*?<\/style>/gi, '');
    
    // テキストのみのバージョンも作成（検索用）
    const plainText = textWithoutStyle.replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ');
    
    Logger.log(`キーワードベース抽出開始: テキストサイズ ${plainText.length}文字`);
    
    // キーワードを準備（デフォルトキーワード含む）
    const keywords = config.keywords || ['KEYWORD_1', 'KEYWORD_2', 'KEYWORD_3', '中傷', '誹謗', '風説の流布'];
    Logger.log(`検索キーワード: ${keywords.join(', ')}`);
    
    let postIndex = 0;
    const contextLength = 300; // 前後の文字数
    const processedPositions = new Set(); // 重複を避けるため
    
    // 各キーワードで検索
    keywords.forEach(keyword => {
      const regex = new RegExp(keyword, 'gi');
      let match;
      
      while ((match = regex.exec(plainText)) !== null) {
        const position = match.index;
        
        // 既に処理済みの位置の近くなら、スキップ（重複回避）
        let skip = false;
        for (const processed of processedPositions) {
          if (Math.abs(position - processed) < contextLength) {
            skip = true;
            break;
          }
        }
        
        if (skip) continue;
        
        processedPositions.add(position);
        
        // 前後300文字を抽出
        const start = Math.max(0, position - contextLength);
        const end = Math.min(plainText.length, position + keyword.length + contextLength);
        const context = plainText.substring(start, end);
        
        // 元のHTMLから対応する部分を探して、投稿情報を抽出
        const htmlContext = extractHtmlContext(textWithoutStyle, start, end, position, keyword);
        
        postIndex++;
        const post = {
          postNumber: `keyword_${postIndex}`,
          title: `キーワード「${keyword}」を含む投稿`,
          content: context.trim(),
          datetime: htmlContext.datetime || '日付不明',
          company: htmlContext.company || config.companyName || '不明',
          userName: config.account || '不明',
          source: 'keyword_extraction',
          matchedKeyword: keyword,
          htmlContent: htmlContext.html,
          postUrl: htmlContext.postUrl || '',
          messageId: htmlContext.messageId || ''
        };
        
        posts.push(post);
        
        Logger.log(`キーワード「${keyword}」マッチ位置: ${position}`);
        Logger.log(`抽出コンテキスト: ${context.substring(0, 50)}...`);
      }
    });
    
    Logger.log(`キーワードベース抽出完了: ${posts.length}件の投稿を検出`);
    
    // 投稿番号と日付で並べ替え（新しい順）
    posts.sort((a, b) => {
      if (a.datetime !== '日付不明' && b.datetime !== '日付不明') {
        return b.datetime.localeCompare(a.datetime);
      }
      return 0;
    });
    
    // 最新5件のみを返す
    const latestPosts = posts.slice(0, 5);
    Logger.log(`最新5件の投稿に限定しました`);
    
    return latestPosts;
    
  } catch (error) {
    Logger.log(`キーワードベース抽出エラー: ${error.toString()}`);
    return [];
  }
}

/**
 * HTMLコンテキストから投稿情報を抽出
 */
function extractHtmlContext(html, start, end, keywordPosition, keyword) {
  try {
    // テキスト位置に対応するHTML部分を探す
    const plainHtml = html.replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ');
    
    // キーワード周辺のHTMLを探す（簡易的な方法）
    const keywordRegex = new RegExp(`([^>]{0,200}${keyword}[^<]{0,200})`, 'i');
    const contextMatch = html.match(keywordRegex);
    
    let result = {
      datetime: '日付不明',
      company: '不明',
      html: '',
      postUrl: '',
      messageId: ''
    };
    
    if (contextMatch) {
      // キーワードの前後1000文字のHTMLを取得
      const htmlStart = Math.max(0, html.indexOf(contextMatch[0]) - 500);
      const htmlEnd = Math.min(html.length, html.indexOf(contextMatch[0]) + contextMatch[0].length + 500);
      const htmlFragment = html.substring(htmlStart, htmlEnd);
      
      result.html = htmlFragment;
      
      // 日付を探す
      const dateMatch = htmlFragment.match(/(\d{4}\/\d{2}\/\d{2}\s+\d{2}:\d{2})/);
      if (dateMatch) {
        result.datetime = dateMatch[1];
      }
      
      // 企業名を探す
      const companyMatch = htmlFragment.match(/[\(（]([^）\)]+)[\)）]/);
      if (companyMatch) {
        result.company = companyMatch[1];
      }
      
      // メッセージIDを探す（投稿のURL生成用）
      const messageIdMatch = htmlFragment.match(/href="\/cm\/message\/(\d+\/\d+\/\d+\/\d+)"/);
      if (messageIdMatch) {
        result.messageId = messageIdMatch[1];
        result.postUrl = `https://finance.yahoo.co.jp/cm/message/${messageIdMatch[1]}`;
      } else {
        // 別のパターンも試す
        const altMessageMatch = htmlFragment.match(/href="\/cm\/message\/([^"]+)"/);
        if (altMessageMatch) {
          result.messageId = altMessageMatch[1];
          result.postUrl = `https://finance.yahoo.co.jp/cm/message/${altMessageMatch[1]}`;
        }
      }
    }
    
    return result;
    
  } catch (error) {
    Logger.log(`HTMLコンテキスト抽出エラー: ${error.toString()}`);
    return {
      datetime: '日付不明',
      company: '不明',
      html: ''
    };
  }
}

/**
 * ウェブページのコンテンツを取得（複数ページ対応強化）
 */
function fetchPageContent(url, retryCount = 0) {
  const maxRetries = 3;
  const delay = 2000; // 2秒
  
  try {
    const options = {
      method: 'GET',
      headers: {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36',
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
        'Accept-Language': 'ja,en-US;q=0.7,en;q=0.3',
        'Accept-Encoding': 'gzip, deflate',
        'DNT': '1',
        'Connection': 'keep-alive',
        'Upgrade-Insecure-Requests': '1'
      },
      followRedirects: true,
      muteHttpExceptions: true
    };
    
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    
    Logger.log(`HTTPレスポンスコード: ${responseCode} for ${url}`);
    
    if (responseCode === 200) {
      const content = response.getContentText('UTF-8');
      Logger.log(`ページ取得成功: ${content.length}文字 from ${url}`);
      return content;
    } else if (responseCode === 429 && retryCount < maxRetries) {
      Logger.log(`レート制限検出、${delay}ms待機後にリトライ (${retryCount + 1}/${maxRetries})`);
      Utilities.sleep(delay);
      return fetchPageContent(url, retryCount + 1);
    } else {
      throw new Error(`HTTPエラー: ${responseCode} for ${url}`);
    }
    
  } catch (error) {
    if (retryCount < maxRetries) {
      Logger.log(`取得エラー、リトライ中 (${retryCount + 1}/${maxRetries}): ${error.toString()}`);
      Utilities.sleep(delay);
      return fetchPageContent(url, retryCount + 1);
    } else {
      throw new Error(`ページ取得に失敗しました: ${error.toString()} for ${url}`);
    }
  }
}

/**
 * Yahoo!ファイナンスのHTMLから投稿データを解析（詳細セクション特化）
 */
function parseYahooFinancePosts(htmlContent, config) {
  const posts = [];
  
  try {
    // 投稿リストセクションを抽出（commentBoxを含む部分を広く取得）
    let targetContent = htmlContent;
    
    // まず、commentBoxが含まれる部分を探す
    const commentBoxMatch = htmlContent.match(/<ul\s+class=\"commentList\"[\s\S]*?<\/ul>/i);
    
    if (commentBoxMatch) {
      targetContent = commentBoxMatch[0];
      Logger.log(`コメントリストセクションを抽出: ${targetContent.length}文字`);
    } else {
      // フォールバック: cCommentListセクション全体を取得
      const commentListMatch = htmlContent.match(/<div\s+class=\"cCommentList\s+cf\">[\s\S]*?(?=<\/div>\s*<\/div>|$)/i);
      
      if (commentListMatch) {
        targetContent = commentListMatch[0];
        Logger.log(`cCommentListセクションを抽出: ${targetContent.length}文字`);
      } else {
        // 最後のフォールバック: myContentsセクション全体
        const myContentsMatch = htmlContent.match(/<div\s+class=\"myContents\s+cf\">[\s\S]*?(?=<\/div>\s*(?:<\/div>|<\/body>|<\/html>|$))/i);
        
        if (myContentsMatch) {
          targetContent = myContentsMatch[0];
          Logger.log(`myContentsセクションを抽出: ${targetContent.length}文字`);
        } else {
          Logger.log('特定セクションが見つからない、全HTMLを使用');
        }
      }
    }
    
    // デバッグ情報を出力
    Logger.log(`解析対象コンテンツサイズ: ${targetContent.length}文字`);
    debugHTMLStructure(targetContent);
    
    // ユーザー名を抽出（より幅広いパターン）
    const userNamePatterns = [
      /<h1[^>]*>([^<]+)</i,
      /<title[^>]*>([^<]*さん[^<]*)</i,
      /ユーザー名[\s\S]*?([\wぁ-ゖァ-ヶェ-ー一-龯]+)/i
    ];
    
    let userName = '不明';
    for (const pattern of userNamePatterns) {
      const match = targetContent.match(pattern);
      if (match && match[1]) {
        userName = match[1].trim();
        break;
      }
    }
    
    Logger.log(`ユーザー名を検出: ${userName}`);
    
    // Yahoo!ファイナンスの実際のHTML構造に基づいた改善されたパターン
    const postPatterns = [
      // パターン1: commentBox全体をキャプチャ（最も包括的）
      /<li\s+class=\"commentBox\"[^>]*>([\s\S]*?)<\/li>/gis,
      
      // パターン2: 詳細な構造パターン（企業名 + No. + タイトル + 日付 + detail内容）
      /<a\s+href=\"[^\"]*\/message\/[^\"]*\"[^>]*>([^<]+)<\/a>[\s\S]*?<span\s+class=\"commentNumber\"[^>]*>No\.(\d+)[\s\S]*?<h2[^>]*commentTitleArea[^>]*>[\s\S]*?<a[^>]*>([^<]+)<\/a>[\s\S]*?(\d{4}\/\d{2}\/\d{2}\s+\d{2}:\d{2})[\s\S]*?<div[^>]*class=\"detail\"[^>]*>([\s\S]*?)<\/div>/gis,
      
      // パターン3: 簡略化パターン（No. + 日付を基準）
      /<span[^>]*commentNumber[^>]*>No\.(\d+)[\s\S]*?(\d{4}\/\d{2}\/\d{2}\s+\d{2}:\d{2})[\s\S]*?<h2[^>]*>[\s\S]*?<a[^>]*>([^<]+)<\/a>[\s\S]*?<p[^>]*>([\s\S]*?)<\/p>/gis,
      
      // パターン4: より広範囲なテキスト抽出
      /No\.(\d+)[\s\S]{1,500}?(\d{4}\/\d{2}\/\d{2}\s+\d{2}:\d{2})[\s\S]{1,1000}?([\u3041-ゖァ-ヶェ-ー一-龯\w\s。、！？]{20,500})/gis,
      
      // パターン5: 投稿内容重視のパターン
      /<div[^>]*detail[^>]*>[\s\S]*?<p[^>]*>([\s\S]*?)<\/p>[\s\S]*?No\.(\d+)[\s\S]*?(\d{4}\/\d{2}\/\d{2}\s+\d{2}:\d{2})/gis
    ];
    
    // 各パターンで解析を試行
    for (let i = 0; i < postPatterns.length; i++) {
      const pattern = postPatterns[i];
      let match;
      let patternPosts = [];
      
      Logger.log(`パターン${i + 1}で解析中...`);
      
      while ((match = pattern.exec(targetContent)) !== null) {
        const post = extractPostDataImproved(match, i + 1, userName, config);
        if (post) {
          patternPosts.push(post);
        }
        
        // 無限ループ防止
        if (patternPosts.length > 100) {
          Logger.log('パターンマッチ数が上限に達しました');
          break;
        }
      }
      
      if (patternPosts.length > 0) {
        Logger.log(`パターン${i + 1}で${patternPosts.length}件の投稿を検出`);
        posts.push(...patternPosts);
        break; // 最初に成功したパターンを使用
      }
    }
    
    // フォールバック: 改善されたシンプル解析
    if (posts.length === 0) {
      Logger.log('正規表現での解析に失敗、改善されたシンプル解析を実行');
      return parseWithImprovedTextAnalysis(targetContent, userName, config);
    }
    
    // 重複を除去してソート
    const uniquePosts = removeDuplicatePosts(posts);
    Logger.log(`重複除去後: ${uniquePosts.length}件の投稿`);
    
    return uniquePosts;
    
  } catch (error) {
    Logger.log(`HTML解析エラー: ${error.toString()}`);
    return [];
  }
}

/**
 * HTML構造のデバッグ情報を出力
 */
function debugHTMLStructure(htmlContent) {
  try {
    // 主要なHTMLタグの数をカウント
    const tagCounts = {
      'article': (htmlContent.match(/<article/gi) || []).length,
      'li': (htmlContent.match(/<li/gi) || []).length,
      'div': (htmlContent.match(/<div/gi) || []).length,
      'a': (htmlContent.match(/<a/gi) || []).length,
      'time': (htmlContent.match(/<time/gi) || []).length,
      'h1': (htmlContent.match(/<h1/gi) || []).length,
      'h2': (htmlContent.match(/<h2/gi) || []).length,
      'h3': (htmlContent.match(/<h3/gi) || []).length
    };
    
    Logger.log('HTMLタグ統計: ' + JSON.stringify(tagCounts));
    
    // 日付パターンを検索
    const dateMatches = htmlContent.match(/\d{4}\/\d{2}\/\d{2}[\s\S]*?\d{2}:\d{2}/g);
    Logger.log(`日付パターン検出: ${dateMatches ? dateMatches.length : 0}件`);
    if (dateMatches && dateMatches.length > 0) {
      Logger.log('最初の日付例: ' + dateMatches[0]);
    }
    
    // リンクパターンを検索
    const linkMatches = htmlContent.match(/href="[^"]*\/message\/[^"]*"/g);
    Logger.log(`メッセージリンク検出: ${linkMatches ? linkMatches.length : 0}件`);
    if (linkMatches && linkMatches.length > 0) {
      Logger.log('最初のリンク例: ' + linkMatches[0]);
    }
    
    // サンプルHTMLを出力（最初の1000文字）
    const sample = htmlContent.substring(0, 1000).replace(/\s+/g, ' ');
    Logger.log('HTMLサンプル: ' + sample);
    
  } catch (error) {
    Logger.log('デバッグ情報出力エラー: ' + error.toString());
  }
}

/**
 * Yahoo!ファイナンスに特化した投稿データ抽出（改善版）
 */
function extractPostDataImproved(match, patternType, userName, config) {
  try {
    let postNumber = '', title = '', content = '', datetime = '', company = config.companyName || '不明';
    
    // マッチ結果のサイズをログ出力
    Logger.log(`パターン${patternType}マッチサイズ: ${match.length}`);
    
    switch (patternType) {
      case 1: // commentBox全体をキャプチャ（最優先）
        const commentBoxContent = match[1];
        return parseCommentBox(commentBoxContent, userName, config);
        
      case 2: // 詳細構造パターン（企業名 + No. + タイトル + 日付 + detail内容）
        [, company, postNumber, title, datetime, content] = match;
        // detail divから実際のテキストを抽出
        const detailPMatch = content.match(/<p[^>]*>([\s\S]*?)<\/p>/);
        if (detailPMatch) {
          content = detailPMatch[1];
        }
        break;
        
      case 3: // 簡略化パターン（No. + 日付 + タイトル + 内容）
        [, postNumber, datetime, title, content] = match;
        break;
        
      case 4: // 広範囲テキスト抽出
        [, postNumber, datetime, content] = match;
        title = content.substring(0, 50) + '...';
        break;
        
      case 5: // 投稿内容重視
        [, content, postNumber, datetime] = match;
        title = content.substring(0, 50) + '...';
        break;
        
      default:
        return null;
    }
    
    // データの清理と正規化
    postNumber = cleanText(postNumber);
    title = cleanText(title);
    content = cleanText(content);
    datetime = cleanText(datetime);
    company = cleanText(company) || config.companyName || '不明';
    
    // 空のデータをチェック
    if (!title && !content) {
      Logger.log(`パターン${patternType}: タイトルとコンテンツが空`);
      return null;
    }
    
    const result = {
      postNumber: postNumber || 'N/A',
      title: title || content.substring(0, 30) + '...',
      content: content || title,
      datetime: datetime || '日付不明',
      company: company,
      userName: userName,
      source: `scraping_pattern_${patternType}`
    };
    
    Logger.log(`投稿データ抽出成功: No.${result.postNumber}, ${result.title.substring(0, 20)}...`);
    return result;
    
  } catch (error) {
    Logger.log(`投稿データ抽出エラー (パターン${patternType}): ${error.toString()}`);
    return null;
  }
}

/**
 * commentBox全体を解析して投稿データを抽出
 */
function parseCommentBox(commentBoxContent, userName, config) {
  try {
    // 企業名を抽出
    const companyMatch = commentBoxContent.match(/<a\s+href="\/cm\/message\/[^"]*">\s*([^<]+)\s*<\/a>/);
    const company = companyMatch ? cleanText(companyMatch[1]) : (config.companyName || '不明');
    
    // 投稿番号を抽出
    const numberMatch = commentBoxContent.match(/<span\s+class="commentNumber">\s*No\.(\d+)/);
    const postNumber = numberMatch ? numberMatch[1] : 'N/A';
    
    // タイトルを抽出
    const titleMatch = commentBoxContent.match(/<h2\s+class="commentTitleArea">[\s\S]*?<a[^>]*>([^<]+)<\/a>/);
    const title = titleMatch ? cleanText(titleMatch[1]) : 'タイトル不明';
    
    // 日付を抽出
    const dateMatch = commentBoxContent.match(/<div\s+class="ttlInfoDateNum">[\s\S]*?<p>\s*(\d{4}\/\d{2}\/\d{2}\s+\d{2}:\d{2})/);
    const datetime = dateMatch ? dateMatch[1] : '日付不明';
    
    // コンテンツを抽出（改善版）
    // まず detail クラスの div を探す
    const detailMatch = commentBoxContent.match(/<div\s+class="detail">([\s\S]*?)<\/div>/);
    let content = '';
    
    if (detailMatch) {
      // detail内のpタグから内容を抽出
      const pMatch = detailMatch[1].match(/<p[^>]*>([\s\S]*?)<\/p>/);
      if (pMatch) {
        content = cleanText(pMatch[1]);
      } else {
        // pタグがない場合は、detail全体のテキストを使用
        content = cleanText(detailMatch[1]);
      }
    }
    
    // コンテンツが空の場合、別のパターンを試す
    if (!content) {
      // style属性を持つpタグを探す
      const styleMatch = commentBoxContent.match(/<p\s+style="[^"]*"[^>]*>([\s\S]*?)<\/p>/);
      if (styleMatch) {
        content = cleanText(styleMatch[1]);
      }
    }
    
    // それでも空の場合は、タイトルを使用
    if (!content) {
      content = title;
    }
    
    Logger.log(`commentBox解析結果: No.${postNumber}, ${company}, ${title.substring(0, 20)}...`);
    Logger.log(`抽出されたコンテンツ: ${content.substring(0, 100)}...`);
    
    return {
      postNumber: postNumber,
      title: title,
      content: content,
      datetime: datetime,
      company: company,
      userName: userName,
      source: 'scraping_commentbox'
    };
    
  } catch (error) {
    Logger.log(`commentBox解析エラー: ${error.toString()}`);
    return null;
  }
}

/**
 * Yahoo!ファイナンス用テキスト清理関数
 */
function cleanText(text) {
  if (!text) return '';
  
  return text
    .replace(/<br\s*\/?>/gi, '\n') // <br>を改行に変換
    .replace(/<[^>]+>/g, '') // その他のHTMLタグを除去
    .replace(/&hellip;/g, '...') // HTMLエンティティを変換
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/&[a-zA-Z0-9#]+;/g, ' ') // その他のHTMLエンティティ
    .replace(/\s+/g, ' ') // 連続する空白を一つに
    .trim(); // 前後の空白を除去
}

/**
 * 重複投稿を除去
 */
function removeDuplicatePosts(posts) {
  const seen = new Set();
  return posts.filter(post => {
    const key = `${post.datetime}-${post.title.substring(0, 20)}`;
    if (seen.has(key)) {
      return false;
    }
    seen.add(key);
    return true;
  });
}

/**
 * 改善されたシンプルテキスト解析（フォールバック）
 */
function parseWithImprovedTextAnalysis(htmlContent, userName, config) {
  const posts = [];
  
  try {
    // HTMLタグを除去してテキストのみを抽出
    const textContent = htmlContent.replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ');
    
    Logger.log('シンプル解析開始、テキストサイズ: ' + textContent.length);
    
    // Yahoo!ファイナンスの実際のテキスト構造に基づいたパターン（改善版）
    const simplePatterns = [
      // パターン1: 投稿コメント一覧の標準パターン（より広範囲）
      /No\.(\d+)[\s\S]{1,100}?(\d{4}\/\d{2}\/\d{2}\s+\d{2}:\d{2})[\s\S]{1,200}?([\u3041-ゖァ-ヶェ-ー一-龯\w\s。、！？\-\(\)（）]{10,1000}?)(?=No\.|\d{4}\/\d{2}\/\d{2}|$)/g,
      
      // パターン2: 企業名付きパターン（拡張版）
      /([\(（][^）\)]+[\)）][\s\S]{1,100}?)No\.(\d+)[\s\S]{1,100}?(\d{4}\/\d{2}\/\d{2}\s+\d{2}:\d{2})[\s\S]{1,200}?([\u3041-ゖァ-ヶェ-ー一-龯\w\s。、！？\-\(\)（）]{10,500})/g,
      
      // パターン3: キーワード重視パターン（設定されたキーワードを含むテキスト）
      config.keywords && config.keywords.length > 0 
        ? new RegExp(`([\\u3041-ゖァ-ヶェ-ー一-龯\\w\\s。、！？\\-\\(\\)（）]{0,200}(?:${config.keywords.join('|')})[\\u3041-ゖァ-ヶェ-ー一-龯\\w\\s。、！？\\-\\(\\)（）]{0,300})`, 'gi')
        : null,
      
      // パターン4: 日付ベースの抽出（日付の前後のテキストを抽出）
      /(\d{4}\/\d{2}\/\d{2}\s+\d{2}:\d{2})[\s\S]{1,50}?([\u3041-ゖァ-ヶェ-ー一-龯\w\s。、！？\-\(\)（）]{20,800}?)(?=\d{4}\/\d{2}\/\d{2}|No\.|$)/g,
      
      // パターン5: 改行ベースの区切り
      /([^\n]{20,1000})/g
    ].filter(Boolean); // nullの要素を除去
    
    // 各パターンで解析（改善版）
    simplePatterns.forEach((pattern, patternIndex) => {
      let match;
      let patternMatches = 0;
      
      Logger.log(`改善シンプルパターン${patternIndex + 1}で解析中...`);
      
      while ((match = pattern.exec(textContent)) !== null && patternMatches < 50) {
        let postNumber, title, content, datetime, company = config.companyName || '不明';
        
        switch (patternIndex) {
          case 0: // 標準パターン: No. + 日付 + 内容
            [, postNumber, datetime, content] = match;
            title = content.substring(0, 50) + '...';
            break;
            
          case 1: // 企業名付き: 企業名 + No. + 日付 + 内容
            [, company, postNumber, datetime, content] = match;
            title = content.substring(0, 50) + '...';
            break;
            
          case 2: // キーワード重視パターン
            [, content] = match;
            title = content.substring(0, 50) + '...';
            postNumber = `keyword_${patternMatches + 1}`;
            datetime = '日付抽出失敗';
            break;
            
          case 3: // 日付ベース: 日付 + 内容
            [, datetime, content] = match;
            title = content.substring(0, 50) + '...';
            postNumber = `date_${patternMatches + 1}`;
            break;
            
          case 4: // 改行ベース
            [, content] = match;
            title = content.substring(0, 50) + '...';
            postNumber = `line_${patternMatches + 1}`;
            datetime = '日付抽出失敗';
            break;
            
          default:
            continue;
        }
        
        // データの清理と検証
        content = cleanText(content);
        title = cleanText(title);
        datetime = cleanText(datetime);
        company = cleanText(company) || config.companyName || '不明';
        
        // キーワードフィルタリング（設定されている場合）
        if (config.keywords && config.keywords.length > 0) {
          const hasKeyword = config.keywords.some(keyword => 
            content.includes(keyword) || title.includes(keyword)
          );
          if (!hasKeyword && patternIndex !== 3) {
            continue; // キーワードが含まれない場合はスキップ
          }
        }
        
        if (content && content.length > 10) {
          const post = {
            postNumber: postNumber || `simple_${patternIndex + 1}_${patternMatches + 1}`,
            title: title || content.substring(0, 30) + '...',
            content: content,
            datetime: datetime || '日付不明',
            company: company,
            userName: userName,
            source: `simple_pattern_${patternIndex + 1}`
          };
          
          posts.push(post);
          patternMatches++;
          
          Logger.log(`パターン${patternIndex + 1}マッチ: No.${post.postNumber} "${post.title.substring(0, 30)}..."`);
          Logger.log(`コンテンツプレビュー: ${post.content.substring(0, 80)}...`);
        }
      }
      
      Logger.log(`改善シンプルパターン${patternIndex + 1}: ${patternMatches}件マッチ`);
      
      // 最初に成功したパターンで十分な結果がある場合は終了
      if (patternMatches > 5) {
        Logger.log(`パターン${patternIndex + 1}で十分な結果が得られたため、他パターンの処理を終了`);
        return;
      }
    });
    
    // デバッグ情報とフォールバックデータ
    if (posts.length === 0) {
      Logger.log('シンプル解析で投稿が検出されなかったため、デバッグ情報を作成');
      
      // テキストのサンプルを出力
      const textSample = textContent.substring(0, 1000);
      Logger.log('テキストサンプル (1000文字): ' + textSample);
      
      // キーワードの出現回数をチェック（詳細版）
      const keywordsToCheck = config.keywords || ['KEYWORD_1', 'KEYWORD_2', 'KEYWORD_3', '中傷', '誹謗', '風説の流布'];
      keywordsToCheck.forEach(keyword => {
        const regex = new RegExp(keyword, 'gi');
        const matches = textContent.match(regex);
        const count = matches ? matches.length : 0;
        Logger.log(`キーワード "${keyword}" の出現回数: ${count}`);
        
        // 出現位置も記録
        if (count > 0 && matches) {
          const positions = [];
          let match;
          const posRegex = new RegExp(keyword, 'gi');
          while ((match = posRegex.exec(textContent)) !== null) {
            positions.push(match.index);
            if (positions.length >= 3) break; // 最初の3つの位置のみ
          }
          Logger.log(`キーワード "${keyword}" の出現位置: ${positions.join(', ')}`);
        }
      });
      
      posts.push({
        postNumber: 'debug_001',
        title: `デバッグ: ページ取得成功、投稿解析失敗`,
        content: `HTML: ${htmlContent.length}文字, テキスト: ${textContent.length}文字. キーワード: [${config.keywords ? config.keywords.join(', ') : 'なし'}]. サンプル: ${textSample.substring(0, 100)}...`,
        datetime: new Date().toLocaleString('ja-JP'),
        company: config.companyName || '不明',
        userName: userName,
        source: 'debug_fallback'
      });
    }
    
    Logger.log(`改善されたシンプル解析で${posts.length}件の投稿を検出`);
    return posts;
    
  } catch (error) {
    Logger.log(`改善されたシンプル解析エラー: ${error.toString()}`);
    return [];
  }
}

/**
 * スクレイピングした投稿データをフォーマット
 */
function formatScrapedPosts(posts) {
  if (!posts || posts.length === 0) {
    return 'スクレイピングで投稿を取得できませんでした。';
  }
  
  let formattedResult = `スクレイピング結果 (${posts.length}件の投稿)\n\n`;
  
  posts.forEach((post, index) => {
    formattedResult += `投稿 ${index + 1}\n`;
    formattedResult += `投稿日時: ${post.datetime}\n`;
    formattedResult += `企業/銘柄: ${post.company}\n`;
    formattedResult += `投稿番号: ${post.postNumber}\n`;
    formattedResult += `投稿タイトル: ${post.title}\n`;
    formattedResult += `投稿内容: ${post.content}\n`;
    formattedResult += `投稿者: ${post.userName}\n`;
    formattedResult += `データソース: ${post.source}\n`;
    if (post.matchedKeyword) {
      formattedResult += `マッチしたキーワード: ${post.matchedKeyword}\n`;
    }
    if (post.postUrl) {
      formattedResult += `投稿URL: ${post.postUrl}\n`;
    }
    formattedResult += `\n`;
  });
  
  formattedResult += `\nスクレイピング結果の概要\n`;
  formattedResult += `総投稿数: ${posts.length}件\n`;
  formattedResult += `ユーザー名: ${posts[0]?.userName || '不明'}\n`;
  formattedResult += `主な企業: ${[...new Set(posts.map(p => p.company))].join(', ')}\n`;
  formattedResult += `取得方法: 直接スクレイピング\n`;
  
      return formattedResult;
}

/**
 * Yahoo!ファイナンスへの通報文章テンプレートを生成
 */
function generateReportTemplate(post, violationType = '名誉毀損・誹謗中傷') {
  const template = `
【通報内容】
違反種別: ${violationType}

【対象投稿】
投稿日時: ${post.datetime}
投稿者: ${post.userName}
投稿URL: ${post.postUrl || '不明'}

【投稿内容（抜粋）】
${post.content}

【違反理由】
上記の投稿は、特定の個人（「${post.matchedKeyword}」さん）に対する根拠のない誹謗中傷・名誉毀損にあたると考えられます。
Yahoo!ファイナンス掲示板の利用規約に違反しており、適切な対応をお願いいたします。

【該当する利用規約違反項目】
- 他人の名誉、信用、プライバシー等を侵害する行為
- 虚偽の情報を投稿する行為
- 第三者に不快感を与える行為

以上、ご確認・ご対応のほどよろしくお願いいたします。
`;
  
  return template;
}

/**
 * 規約違反レポートをフォーマット
 */
function formatViolationReport(violationResults) {
  if (!violationResults || violationResults.length === 0) {
    return '<h2 style="color: #e74c3c;">⚠️ 利用規約違反チェックレポート</h2>\n<p>チェック対象が見つかりませんでした。</p>';
  }
  
  let report = '<h2 style="color: #e74c3c;">⚠️ 利用規約違反チェックレポート</h2>\n';
  report += `<p>チェック実行日時: ${new Date().toLocaleString('ja-JP')}</p>\n`;
  report += `<p>チェック対象: ${violationResults.length}件のアカウント</p>\n\n`;
  
  violationResults.forEach((result, index) => {
    report += `<h3 style="color: #c0392b; margin-top: 20px;">${index + 1}. ${result.account}</h3>\n`;
    report += `<p>プラットフォーム: ${result.platform}</p>\n`;
    
    // ユーザーページへのリンクを追加
    const userPageUrl = `https://finance.yahoo.co.jp/cm/personal/history/comment?user=${result.account}`;
    report += `<div style="background-color: #e3f2fd; padding: 10px; margin: 10px 0; border-left: 4px solid #2196f3;">\n`;
    report += `<p style="font-weight: bold; color: #1976d2;">📌 一次対応のお願い</p>\n`;
    report += `<p>違反報告をするために、以下のユーザーページにアクセスしてください：</p>\n`;
    report += `<p><a href="${userPageUrl}" style="color: #1976d2; text-decoration: underline;">${userPageUrl}</a></p>\n`;
    report += `</div>\n`;
    
    if (result.error) {
      report += `<p style="color: red;">エラー: ${result.error}</p>\n`;
    } else {
      if (result.analysis) {
        report += `<div style="background-color: #f5f5f5; padding: 15px; margin: 10px 0; border-radius: 5px;">\n`;
        report += `<h4 style="color: #2c3e50; margin-top: 0;">🔍 違反分析結果</h4>\n`;
        // マークダウンを削除してHTMLに変換
        const cleanAnalysis = result.analysis
          .replace(/\*\*([^*]+)\*\*/g, '<strong>$1</strong>')
          .replace(/###\s*(.+)/g, '<h5 style="color: #34495e; margin: 10px 0;">$1</h5>')
          .replace(/##\s*(.+)/g, '<h4 style="color: #2c3e50; margin: 15px 0;">$1</h4>')
          .replace(/\n/g, '<br/>');
        report += `${cleanAnalysis}\n`;
        report += `</div>\n`;
      }
      
      if (result.userPosts) {
        report += `<div style="background-color: #e8f4f8; padding: 15px; margin: 10px 0; border-radius: 5px;">\n`;
        report += `<h4 style="color: #2c3e50; margin-top: 0;">📝 収集された投稿（最新5件）</h4>\n`;
        // マークダウンを削除
        const cleanPosts = result.userPosts
          .replace(/###\s*(.+)/g, '<h5 style="color: #34495e; margin: 10px 0;">$1</h5>')
          .replace(/##\s*(.+)/g, '<h4 style="color: #2c3e50; margin: 15px 0;">$1</h4>')
          .replace(/\*\*([^*]+)\*\*/g, '<strong>$1</strong>')
          .replace(/^-\s+(.+)$/gm, '<li>$1</li>')
          .replace(/(<li>.*<\/li>\n?)+/g, '<ul style="margin: 5px 0;">$&</ul>')
          .replace(/\n/g, '<br/>');
        const postsPreview = cleanPosts.substring(0, 3000);
        report += `${postsPreview}${cleanPosts.length > 3000 ? '...' : ''}\n`;
        report += `</div>\n`;
      }
      
      // 通報テンプレートを追加
      if (result.reportTemplates && result.reportTemplates.length > 0) {
        report += `<div style="background-color: #fff3cd; padding: 15px; margin: 10px 0; border: 1px solid #ffeaa7; border-radius: 5px;">\n`;
        report += `<h4 style="color: #856404; margin-top: 0;">📢 通報文章テンプレート</h4>\n`;
        report += `<p style="color: #856404;">以下の文章をコピーしてYahoo!ファイナンスの違反報告フォームでご利用ください：</p>\n`;
        result.reportTemplates.forEach((template, i) => {
          report += `<h5 style="color: #856404; margin: 15px 0 10px 0;">投稿${i + 1}の通報テンプレート</h5>\n`;
          report += `<pre style="background: white; padding: 15px; border: 1px solid #ddd; white-space: pre-wrap; font-family: monospace; font-size: 0.9em; border-radius: 3px;">${template}</pre>\n`;
        });
        report += `</div>\n`;
      }
    }
    
    report += '\n<hr/>\n';
  });
  
  return report;
}

/**
 * エラー通知を送信
 */
function sendErrorNotification(error, processType = '不明な処理') {
  try {
    const subject = `[エラー通知] ${processType}でエラーが発生しました`;
    const body = `
処理種別: ${processType}
発生日時: ${new Date().toLocaleString('ja-JP')}
エラー内容: ${error.toString()}

スタックトレース:
${error.stack || '不明'}

このエラーを確認し、必要に応じて設定を見直してください。
`;
    
    MailApp.sendEmail({
      to: CONFIG.EMAIL_RECIPIENTS,
      subject: subject,
      body: body
    });
    
    Logger.log('エラー通知メールを送信しました');
  } catch (emailError) {
    Logger.log('エラー通知メールの送信に失敗しました: ' + emailError.toString());
  }
}

/**
 * 検索レポートをメール送信
 */
function sendSearchReport(report) {
  try {
    const subject = '[情報収集レポート] キーワード検索結果';
    const htmlBody = `
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <style>
        body { 
            font-family: 'Helvetica Neue', Arial, sans-serif; 
            line-height: 1.6; 
            margin: 0;
            padding: 20px;
            background-color: #f5f5f5;
        }
        .container {
            max-width: 800px;
            margin: 0 auto;
            background-color: white;
            padding: 30px;
            border-radius: 10px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
        }
        h2 { 
            color: #2c3e50; 
            border-bottom: 3px solid #3498db; 
            padding-bottom: 10px;
            margin-bottom: 20px;
        }
        h3 { 
            color: #34495e; 
            margin-top: 30px;
            margin-bottom: 15px;
        }
        .timestamp { 
            color: #7f8c8d; 
            font-size: 0.9em;
            margin-bottom: 20px;
        }
        .content { 
            background-color: #f8f9fa; 
            padding: 20px; 
            border-left: 4px solid #3498db; 
            margin: 15px 0;
            border-radius: 0 5px 5px 0;
        }
        .citations { 
            background-color: #e8f5e9; 
            padding: 15px; 
            margin: 15px 0; 
            border-radius: 5px;
        }
        .error {
            color: #e74c3c;
            background-color: #fadbd8;
            padding: 10px;
            border-radius: 5px;
            margin: 10px 0;
        }
        hr {
            border: none;
            border-top: 1px solid #e0e0e0;
            margin: 30px 0;
        }
    </style>
</head>
<body>
    <div class="container">
        ${report}
    </div>
</body>
</html>
`;
    
    MailApp.sendEmail({
      to: CONFIG.EMAIL_RECIPIENTS,
      subject: subject,
      htmlBody: htmlBody
    });
    
    Logger.log('検索レポートメールを送信しました');
  } catch (error) {
    Logger.log('検索レポートメール送信エラー: ' + error.toString());
  }
}

/**
 * 規約違反レポートをメール送信
 */
function sendViolationReport(report) {
  try {
    const subject = '[規約違反チェック] 分析結果レポート';
    const htmlBody = `
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <style>
        body { 
            font-family: 'Helvetica Neue', Arial, sans-serif; 
            line-height: 1.6; 
            margin: 0;
            padding: 20px;
            background-color: #f5f5f5;
        }
        .container {
            max-width: 800px;
            margin: 0 auto;
            background-color: white;
            padding: 30px;
            border-radius: 10px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
        }
        h2 { 
            color: #e74c3c; 
            border-bottom: 3px solid #e74c3c; 
            padding-bottom: 10px;
            margin-bottom: 20px;
        }
        h3 { 
            color: #c0392b; 
            margin-top: 30px;
            margin-bottom: 15px;
        }
        h4 {
            color: #2c3e50;
            margin-top: 20px;
            margin-bottom: 10px;
        }
        h5 {
            color: #34495e;
            margin: 10px 0;
        }
        p {
            margin: 10px 0;
        }
        a {
            color: #3498db;
            text-decoration: none;
        }
        a:hover {
            text-decoration: underline;
        }
        ul {
            margin: 10px 0;
            padding-left: 20px;
        }
        li {
            margin: 5px 0;
        }
        pre {
            background-color: #f8f9fa;
            padding: 15px;
            border: 1px solid #dee2e6;
            border-radius: 5px;
            white-space: pre-wrap;
            font-family: 'Courier New', monospace;
            font-size: 0.9em;
            overflow-x: auto;
        }
        .info-box {
            background-color: #e3f2fd;
            border-left: 4px solid #2196f3;
            padding: 15px;
            margin: 15px 0;
            border-radius: 0 5px 5px 0;
        }
        .warning-box {
            background-color: #fff3cd;
            border: 1px solid #ffeaa7;
            padding: 15px;
            margin: 15px 0;
            border-radius: 5px;
        }
        .error {
            color: #e74c3c;
            background-color: #fadbd8;
            padding: 10px;
            border-radius: 5px;
            margin: 10px 0;
        }
        hr {
            border: none;
            border-top: 1px solid #e0e0e0;
            margin: 30px 0;
        }
    </style>
</head>
<body>
    <div class="container">
        ${report}
    </div>
</body>
</html>
`;
    
    MailApp.sendEmail({
      to: CONFIG.EMAIL_RECIPIENTS,
      subject: subject,
      htmlBody: htmlBody
    });
    
    Logger.log('規約違反レポートメールを送信しました');
  } catch (error) {
    Logger.log('規約違反レポートメール送信エラー: ' + error.toString());
  }
}

/**
 * 検索レポートをフォーマット
 */
function formatSearchReport(results) {
  if (!results || results.length === 0) {
    return '<h2 style="color: #2c3e50;">🔍 キーワード検索レポート</h2>\n<p>検索結果が見つかりませんでした。</p>';
  }
  
  let report = '<h2 style="color: #2c3e50;">🔍 キーワード検索レポート</h2>\n';
  report += `<p style="color: #7f8c8d; font-size: 0.9em;">生成日時: ${new Date().toLocaleString('ja-JP')}</p>\n`;
  report += `<p>検索件数: ${results.length}件</p>\n\n`;
  
  results.forEach((result, index) => {
    report += `<h3 style="color: #34495e; margin-top: 20px;">${index + 1}. ${result.keyword}</h3>\n`;
    report += `<p>検索媒体: ${result.media}</p>\n`;
    report += `<p>検索期間: ${result.searchPeriod}</p>\n`;
    
    if (result.error) {
      report += `<div class="error">エラー: ${result.error}</div>\n`;
    } else {
      report += `<div class="content">${result.content.replace(/\n/g, '<br/>')}</div>\n`;
      
      if (result.citations && result.citations.length > 0) {
        report += '<div class="citations"><h4>参考リンク:</h4><ul>\n';
        result.citations.forEach(citation => {
          report += `<li><a href="${citation}">${citation}</a></li>\n`;
        });
        report += '</ul></div>\n';
      }
    }
    
    report += '\n';
  });
  
  return report;
}

// ==================== シート作成機能 ====================

/**
 * 違反チェック設定シートを作成
 */
function createViolationCheckSheet() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.VIOLATION_CHECK);
  
  if (sheet) {
    const confirm = ui.alert('確認', `${CONFIG.SHEET_NAMES.VIOLATION_CHECK}シートが既に存在します。\n内容をクリアして再作成しますか？`, ui.ButtonSet.YES_NO);
    if (confirm !== ui.Button.YES) {
      ui.alert('情報', 'シートの作成をキャンセルしました。', ui.ButtonSet.OK);
      return;
    }
    sheet.clear();
  } else {
    sheet = ss.insertSheet(CONFIG.SHEET_NAMES.VIOLATION_CHECK);
  }
  
  // ヘッダー設定（簡素化版）
  const headers = ['アカウント名', 'チェックキーワード（カンマ区切り）', '利用規約URL', '有効/無効'];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#d9ead3')
    .setHorizontalAlignment('center');
  
  // 列幅調整
  sheet.setColumnWidths(1, headers.length, 200);
  
  // サンプルデータ（簡素化）
  const sampleData = [
    ['ACCOUNT_HASH_PLACEHOLDER', 'KEYWORD_1,KEYWORD_2,KEYWORD_3,中傷,誹謗,風説の流布', 'https://support.yahoo-net.jp/PccFinance/s/article/H000011273', true],
    ['another_user_id', '不適切な発言,ハラスメント,偽情報', 'https://example.com/terms', false]
  ];
  sheet.getRange(2, 1, sampleData.length, headers.length).setValues(sampleData);
  
  // データ検証（有効/無効）
  const range = sheet.getRange(2, headers.length, sheet.getMaxRows() - 1, 1);
  const rule = SpreadsheetApp.newDataValidation()
    .requireCheckbox()
    .setAllowInvalid(false)
    .build();
  range.setDataValidation(rule);
  
  ui.alert('成功', `${CONFIG.SHEET_NAMES.VIOLATION_CHECK}シートを作成しました。\nアカウント名、キーワード、利用規約URLを入力して使用してください。`, ui.ButtonSet.OK);
  
  Logger.log(`${CONFIG.SHEET_NAMES.VIOLATION_CHECK}シートを作成しました`);
}

/**
 * 検索設定シートを作成
 */
function createSearchSheet() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.SEARCH);
  
  if (sheet) {
    const confirm = ui.alert('確認', `${CONFIG.SHEET_NAMES.SEARCH}シートが既に存在します。\n内容をクリアして再作成しますか？`, ui.ButtonSet.YES_NO);
    if (confirm !== ui.Button.YES) {
      ui.alert('情報', 'シートの作成をキャンセルしました。', ui.ButtonSet.OK);
      return;
    }
    sheet.clear();
  } else {
    sheet = ss.insertSheet(CONFIG.SHEET_NAMES.SEARCH);
  }
  
  // ヘッダー設定
  const headers = ['キーワード', '追加指示', '媒体名', '媒体URL', '検索期間（日前）', '有効/無効'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#4285f4')
    .setFontColor('#ffffff')
    .setHorizontalAlignment('center');
  
  // サンプルデータ
  const sampleData = [
    ['AI技術の最新動向', '2024年以降の情報を中心に', 'TechCrunch', 'https://techcrunch.com', 7, true],
    ['Web3.0', '日本市場の動向を含める', 'CoinDesk', 'https://www.coindesk.com', 14, false]
  ];
  sheet.getRange(2, 1, sampleData.length, headers.length).setValues(sampleData);
  
  // 書式設定
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, headers.length);
  
  // 期間列にコメントを追加
  sheet.getRange(1, 5).setNote('本日から何日前まで遡って検索するかを数値で指定してください。例：7（一週間前まで）、30（一ヶ月前まで）');
  
  ui.alert('成功', `${CONFIG.SHEET_NAMES.SEARCH}シートを作成しました。\nキーワード、媒体、検索期間を設定して使用してください。`, ui.ButtonSet.OK);
  
  Logger.log(`${CONFIG.SHEET_NAMES.SEARCH}シートを作成しました`);
}

// ==================== メイン関数（Perplexity APIでの投稿収集＆違反分析） ====================

/**
 * Yahoo!ファイナンス掲示板からユーザー投稿を収集（スクレイピング優先）
 */
function collectUserPosts(config) {
  try {
    // まずウェブスクレイピングを試行
    Logger.log('ウェブスクレイピングで投稿収集を試みます');
    const scrapedResult = scrapeUserPostsFromYahoo(config);
    
    if (scrapedResult && !scrapedResult.includes('スクレイピングで投稿を取得できませんでした')) {
      Logger.log('ウェブスクレイピングに成功しました');
      return scrapedResult;
    }
    
  } catch (scrapingError) {
    Logger.log(`ウェブスクレイピング失敗: ${scrapingError.toString()}`);
  }
  
  // フォールバック: Perplexity APIを使用
  Logger.log('フォールバック: Perplexity APIで投稿収集を実行');
  return collectUserPostsWithAPI(config);
}

/**
 * Perplexity APIでユーザー投稿を収集（フォールバック）
 */
function collectUserPostsWithAPI(config) {
  const collectPrompt = buildCollectionPrompt(config);
  
  // domain filterをYahoo!ファイナンスに限定
  const payload = {
    model: 'sonar-pro',
    messages: [
      {
        role: 'system',
        content: 'あなたはYahoo!ファイナンス掲示板の投稿を精密に収集・分析する専門家です。投稿の内容、投稿者、日時を正確に記録します。'
      },
      {
        role: 'user',
        content: collectPrompt
      }
    ],
    temperature: 0.1,
    max_tokens: 3000,
    search_domain_filter: ['finance.yahoo.co.jp']
  };
  
  const options = {
    method: 'post',
    headers: {
      'Authorization': `Bearer ${CONFIG.PERPLEXITY_API_KEY}`,
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  
  try {
    const response = UrlFetchApp.fetch(CONFIG.PERPLEXITY_API_URL, options);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();
    
    Logger.log(`投稿収集APIレスポンスコード: ${responseCode}`);
    
    if (responseCode !== 200) {
      Logger.log(`投稿収集APIエラー: ${responseText}`);
      throw new Error(`APIエラー: ステータスコード ${responseCode}`);
    }
    
    const result = JSON.parse(responseText);
    
    if (result.choices && result.choices[0] && result.choices[0].message) {
      Logger.log('APIでのユーザー投稿収集に成功しました');
      return result.choices[0].message.content;
    } else {
      Logger.log('投稿収集APIレスポンスの形式が不正です');
      throw new Error('APIレスポンスの形式が不正です');
    }
    
  } catch (error) {
    Logger.log(`APIでのユーザー投稿収集エラー: ${error.toString()}`);
    throw new Error(`ユーザー投稿の収集に失敗しました: ${error.toString()}`);
  }
}

/**
 * 規約項目別に違反を分析（Grok-4優先、Perplexityフォールバック）
 */
function analyzeViolationsBySection(termsContent, userPosts, config) {
  try {
    Logger.log('規約項目別違反分析を開始します');
    
    // まずGrok-4で分析を試みる
    if (CONFIG.GROK_API_KEY) {
      try {
        Logger.log('Grok-4で違反分析を実行中...');
        return analyzeViolationsWithGrok(termsContent, userPosts, config);
      } catch (grokError) {
        Logger.log(`Grok-4分析失敗: ${grokError.toString()}`);
        Logger.log('フォールバックでPerplexityを使用します');
      }
    } else {
      Logger.log('Grok-4 APIキーが設定されていないため、Perplexityを使用します');
    }
    
    // フォールバック: Perplexity APIを使用
    return analyzeViolationsWithPerplexity(termsContent, userPosts, config);
    
  } catch (error) {
    Logger.log(`規約項目別違反分析エラー: ${error.toString()}`);
    throw new Error(`規約違反分析に失敗しました: ${error.toString()}`);
  }
}

/**
 * Grok-4での違反分析
 */
function analyzeViolationsWithGrok(termsContent, userPosts, config) {
  const prompt = buildGrokViolationAnalysisPrompt(termsContent, userPosts, config);
  
  const payload = {
    model: 'grok-2-1212',
    messages: [
      {
        role: 'system',
        content: 'あなたは日本の法務・コンプライアンスの世界最高水準の専門家です。利用規約違反を精密で客観的に分析し、実用的で具体的なアドバイスを提供してください。'
      },
      {
        role: 'user',
        content: prompt
      }
    ],
    temperature: 0.1,
    max_tokens: 4000
  };
  
  const options = {
    method: 'post',
    headers: {
      'Authorization': `Bearer ${CONFIG.GROK_API_KEY}`,
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  
  const response = UrlFetchApp.fetch(CONFIG.GROK_API_URL, options);
  const responseCode = response.getResponseCode();
  const responseText = response.getContentText();
  
  Logger.log(`Grok-4違反分析APIレスポンスコード: ${responseCode}`);
  
  if (responseCode !== 200) {
    Logger.log(`Grok-4違反分析APIエラー: ${responseText}`);
    throw new Error(`Grok-4 APIエラー: ステータスコード ${responseCode}`);
  }
  
  const result = JSON.parse(responseText);
  
  if (result.choices && result.choices[0] && result.choices[0].message) {
    Logger.log('Grok-4での違反分析が完了しました');
    return result.choices[0].message.content;
  } else {
    Logger.log('Grok-4違反分析APIレスポンスの形式が不正です');
    throw new Error('Grok-4 APIレスポンスの形式が不正です');
  }
}

/**
 * Grok-4用違反分析プロンプトを構築
 */
function buildGrokViolationAnalysisPrompt(termsContent, userPosts, config) {
  const prompt = `以下の利用規約とユーザー投稿を照合し、規約項目別に違反の有無を詳細に分析してください。

## 利用規約内容
${termsContent}

## ユーザー投稿内容
${userPosts}

## 分析結果の出力形式
以下の形式で、規約の各項目ごとに分析してください：

### [項目1: 項目名]
**違反判定**: 【違反あり/違反の可能性あり/違反なし】
**リスクレベル**: 【高/中/低】
**該当投稿**: （具体的な投稿内容を引用）
**理由**: （なぜ違反と判定したかの詳細な理由）
**推奨アクション**: （対応方法の推奨）

### [項目2: 項目名]
(同様の形式で続く)

### 総合評価
**全体的なリスクレベル**: 【高/中/低】
**総合的な対応推奨**: （優先順位と具体的なアクションプラン）

特に以下の点に注意して分析してください：
- 名誉毀損や中傷にあたる内容
- ハラスメントやいじめにあたる行為
- 著作権侵害の可能性
- スパムや迷惑行為
- 偏見や差別的な発言
- 偽情報や誤情報の拡散
- 風説の流布や市場操作`;

  return prompt;
}

/**
 * Perplexity APIでの違反分析（フォールバック）
 */
function analyzeViolationsWithPerplexity(termsContent, userPosts, config) {
  const prompt = `以下の利用規約とユーザー投稿を照合し、規約項目別に違反の有無を詳細に分析してください。

## 利用規約内容
${termsContent}

## ユーザー投稿内容
${userPosts}

## 分析結果の出力形式
以下の形式で、規約の各項目ごとに分析してください：

### [項目1: 項目名]
**違反判定**: 【違反あり/違反の可能性あり/違反なし】
**リスクレベル**: 【高/中/低】
**該当投稿**: （具体的な投稿内容を引用）
**理由**: （なぜ違反と判定したかの詳細な理由）
**推奨アクション**: （対応方法の推奨）

### [項目2: 項目名]
(同様の形式で続く)

### 総合評価
**全体的なリスクレベル**: 【高/中/低】
**編合的な対応推奨**: （優先順位と具体的なアクションプラン）

特に以下の点に注意して分析してください：
- 名誉毀損や中傷にあたる内容
- ハラスメントやいじめにあたる行為
- 著作権侵害の可能性
- スパムや迷惑行為
- 偏見や差別的な発言
- 偽情報や誤情報の拡散`;
  
  const payload = {
    model: 'sonar-pro',
    messages: [
      {
        role: 'system',
        content: 'あなたは利用規約違反を専門的に分析する法務・コンプライアンスの専門家です。精密で客観的な分析を行い、具体的で実用的なアドバイスを提供します。'
      },
      {
        role: 'user',
        content: prompt
      }
    ],
    temperature: 0.1,
    max_tokens: 4000
  };
  
  const options = {
    method: 'post',
    headers: {
      'Authorization': `Bearer ${CONFIG.PERPLEXITY_API_KEY}`,
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  
  const response = UrlFetchApp.fetch(CONFIG.PERPLEXITY_API_URL, options);
  const responseCode = response.getResponseCode();
  const responseText = response.getContentText();
  
  Logger.log(`Perplexity違反分析APIレスポンスコード: ${responseCode}`);
  
  if (responseCode !== 200) {
    Logger.log(`Perplexity違反分析APIエラー: ${responseText}`);
    throw new Error(`Perplexity APIエラー: ステータスコード ${responseCode}`);
  }
  
  const result = JSON.parse(responseText);
  
  if (result.choices && result.choices[0] && result.choices[0].message) {
    Logger.log('Perplexityでの違反分析が完了しました');
    return result.choices[0].message.content;
  } else {
    Logger.log('Perplexity違反分析APIレスポンスの形式が不正です');
    throw new Error('Perplexity APIレスポンスの形式が不正です');
  }
}

/**
 * Yahoo!ファイナンス掲示板用の投稿収集プロンプトを構築
 */
function buildCollectionPrompt(config) {
  let prompt = `Yahoo!ファイナンス掲示板から特定ユーザーの投稿を収集してください：\n\n`;
  
  // ユーザー情報
  prompt += `ユーザーID: ${config.account}\n`;
  prompt += `ユーザーページ: ${config.platformURL}\n`;
  
  if (config.keywords && config.keywords.length > 0) {
    prompt += `\n特に以下のキーワードを含む投稿に注目してください：\n`;
    config.keywords.forEach(keyword => {
      prompt += `- ${keyword}\n`;
    });
  }
  
  prompt += `\n以下の形式で投稿を収集してください：\n\n`;
  prompt += `## 投稿一覧\n\n`;
  prompt += `### 投稿 1\n`;
  prompt += `- **投稿日時**: 2025年8月14日 19:13\n`;
  prompt += `- **投稿番号**: 597\n`;
  prompt += `- **投稿内容**: （完全な投稿テキスト）\n`;
  prompt += `- **投稿者**: ユーザー名\n`;
  prompt += `- **投稿URL**: 直接リンク（可能であれば）\n\n`;
  prompt += `### 投稿 2\n`;
  prompt += `（同様の形式で続く）\n\n`;
  
  prompt += `注意事項：\n`;
  prompt += `- 最新の投稿から順に10件程度収集してください\n`;
  prompt += `- 投稿内容は省略せず全文を記載してください\n`;
  prompt += `- Yahoo!ファイナンスの掲示板システムに特化した検索を実行してください\n`;
  prompt += `- 投稿者の名前やユーザーIDも可能な限り記載してください\n`;
  
  return prompt;
}

// ==================== ユーティリティ ====================

/**
 * URLからドメインを抽出
 */
function extractDomain(url) {
  try {
    const match = url.match(/^https?:\/\/([^\/]+)/);
    return match ? match[1] : url;
  } catch (e) {
    return url;
  }
}

// ==================== 初期設定 ====================

/**
 * 初期設定を実行（最初に一度だけ実行）
 */
function initialSetup() {
  const ui = SpreadsheetApp.getUi();
  
  // スプレッドシートの確認
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let setupMessages = [];
  
  // APIキーの確認
  if (!CONFIG.PERPLEXITY_API_KEY) {
    setupMessages.push('❌ PERPLEXITY_API_KEYが設定されていません');
    setupMessages.push('設定方法: スクリプトエディタ → プロジェクト設定 → スクリプトプロパティ');
  } else {
    setupMessages.push('✅ APIキーが設定されています');
  }
  
  // メールアドレスの確認
  if (CONFIG.EMAIL_RECIPIENTS === 'your-email@example.com') {
    setupMessages.push('⚠️ メール送信先を変更してください (現在: your-email@example.com)');
  } else {
    setupMessages.push('✅ メール送信先: ' + CONFIG.EMAIL_RECIPIENTS);
  }
  
  // シート1の作成
  if (!spreadsheet.getSheetByName(CONFIG.SHEET_NAMES.SEARCH)) {
    createSearchSheet();
    setupMessages.push('✅ シート1_検索設定を作成しました');
  } else {
    setupMessages.push('✅ シート1_検索設定は既に存在します');
  }
  
  // シート2の作成
  if (!spreadsheet.getSheetByName(CONFIG.SHEET_NAMES.VIOLATION_CHECK)) {
    createViolationCheckSheet();
    setupMessages.push('✅ シート2_規約チェックを作成しました');
  } else {
    setupMessages.push('✅ シート2_規約チェックは既に存在します');
  }
  
  // セットアップ結果を表示
  const message = setupMessages.join('\n');
  ui.alert('初期設定状態', message, ui.ButtonSet.OK);
  
  // APIキーが設定されていない場合は終了
  if (!CONFIG.PERPLEXITY_API_KEY) {
    ui.alert('エラー', 'APIキーを設定してから再度実行してください', ui.ButtonSet.OK);
    return;
  }
  
  // テスト実行の確認
  const response = ui.alert(
    'テスト実行', 
    'テスト実行を行いますか？\n（APIを使用しない動作確認のみ）', 
    ui.ButtonSet.YES_NO
  );
  
  if (response === ui.Button.YES) {
    testRun();
  }
}

/**
 * テスト実行（API呼び出しなし）
 */
function testRun() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    Logger.log('=== テスト実行開始 ===');
    
    // シート1のテスト
    const sheet1 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAMES.SEARCH);
    if (sheet1) {
      const data1 = sheet1.getDataRange().getValues();
      Logger.log(`シート1: ${data1.length - 1}件の設定を確認`);
      
      let activeCount1 = 0;
      for (let i = 1; i < data1.length; i++) {
        if (data1[i][4] === true || data1[i][4] === 'TRUE') {
          activeCount1++;
        }
      }
      Logger.log(`シート1: ${activeCount1}件が有効`);
    }
    
    // シート2のテスト
    const sheet2 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAMES.VIOLATION_CHECK);
    if (sheet2) {
      const data2 = sheet2.getDataRange().getValues();
      Logger.log(`シート2: ${data2.length - 1}件の設定を確認`);
      
      let activeCount2 = 0;
      for (let i = 1; i < data2.length; i++) {
        if (data2[i][5] === true || data2[i][5] === 'TRUE') {
          activeCount2++;
        }
      }
      Logger.log(`シート2: ${activeCount2}件が有効`);
    }
    
    Logger.log('=== テスト実行完了 ===');
    ui.alert('テスト完了', 'シートの設定を確認しました。\n詳細はログを確認してください。', ui.ButtonSet.OK);
    
  } catch (error) {
    Logger.log('テストエラー: ' + error.toString());
    ui.alert('エラー', 'テスト中にエラーが発生しました:\n' + error.toString(), ui.ButtonSet.OK);
  }
}

/**
 * APIキーのテスト
 */
function testAPIKey() {
  const ui = SpreadsheetApp.getUi();
  
  if (!CONFIG.PERPLEXITY_API_KEY) {
    ui.alert('エラー', 'APIキーが設定されていません', ui.ButtonSet.OK);
    return;
  }
  
  const payload = {
    model: 'sonar-pro',
    messages: [
      {
        role: 'user',
        content: 'Hello, this is a test. Please respond with "API test successful".'
      }
    ],
    temperature: 0.1,
    max_tokens: 50
  };
  
  const options = {
    method: 'post',
    headers: {
      'Authorization': `Bearer ${CONFIG.PERPLEXITY_API_KEY}`,
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  
  try {
    const response = UrlFetchApp.fetch(CONFIG.PERPLEXITY_API_URL, options);
    const responseCode = response.getResponseCode();
    
    if (responseCode === 200) {
      const result = JSON.parse(response.getContentText());
      if (result.choices && result.choices[0]) {
        ui.alert('成功', 'APIキーのテストに成功しました！\n\nレスポンス:\n' + result.choices[0].message.content, ui.ButtonSet.OK);
      }
    } else if (responseCode === 401) {
      ui.alert('エラー', 'APIキーが無効です。正しいキーを設定してください。', ui.ButtonSet.OK);
    } else {
      ui.alert('エラー', `APIエラー: ステータスコード ${responseCode}`, ui.ButtonSet.OK);
    }
  } catch (error) {
    ui.alert('エラー', 'API接続エラー:\n' + error.toString(), ui.ButtonSet.OK);
  }
}

/**
 * ウェブスクレイピングのテスト実行
 */
function testWebScraping() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    Logger.log('=== ウェブスクレイピングテスト開始 ===');
    
    // テスト用の設定（キーワード更新）
    const testConfig = {
      accountName: 'ACCOUNT_HASH_PLACEHOLDER',
      companyName: 'COMPANY_NAME_PLACEHOLDER',
      platform: 'Yahoo!ファイナンス掲示板',
      platformURL: 'https://finance.yahoo.co.jp/cm/personal/history/comment?user=ACCOUNT_HASH_PLACEHOLDER&sort=2',
      keywords: ['KEYWORD_1', 'KEYWORD_2', 'KEYWORD_3'],
      tosURL: 'https://support.yahoo-net.jp/PccFinance/s/article/H000011273',
      isActive: true
    };
    
    // ウェブスクレイピングテスト
    const result = scrapeUserPostsFromYahoo(testConfig);
    
    if (result) {
      Logger.log('ウェブスクレイピングテスト成功');
      
      // 結果をダイアログで表示
      const summary = result.substring(0, 500) + (result.length > 500 ? '...' : '');
      ui.alert(
        'ウェブスクレイピングテスト結果',
        `テストに成功しました！\n\n${summary}\n\n詳細はログを確認してください。`,
        ui.ButtonSet.OK
      );
    } else {
      Logger.log('ウェブスクレイピングテスト失敗');
      ui.alert('エラー', 'ウェブスクレイピングテストに失敗しました。\nログを確認してください。', ui.ButtonSet.OK);
    }
    
  } catch (error) {
    Logger.log('ウェブスクレイピングテストエラー: ' + error.toString());
    ui.alert('エラー', 'ウェブスクレイピングテスト中にエラーが発生しました：\n' + error.toString(), ui.ButtonSet.OK);
  }
}

/**
 * ページ取得のシンプルテスト
 */
function testPageFetch() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    const testUrl = 'https://finance.yahoo.co.jp/cm/personal/history/comment?user=ACCOUNT_HASH_PLACEHOLDER&sort=2';
    
    Logger.log('ページ取得テスト開始: ' + testUrl);
    
    const content = fetchPageContent(testUrl);
    
    if (content) {
      Logger.log('ページ取得成功: ' + content.length + '文字');
      ui.alert('成功', `ページ取得に成功しました！\n\nコンテンツサイズ: ${content.length}文字`, ui.ButtonSet.OK);
    } else {
      ui.alert('失敗', 'ページの取得に失敗しました。', ui.ButtonSet.OK);
    }
    
  } catch (error) {
    Logger.log('ページ取得テストエラー: ' + error.toString());
    ui.alert('エラー', 'ページ取得テスト中にエラーが発生しました：\n' + error.toString(), ui.ButtonSet.OK);
  }
}

/**
 * Grok-4 APIテスト
 */
function testGrokAPI() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    Logger.log('=== Grok-4 APIテスト開始 ===');
    
    if (!CONFIG.GROK_API_KEY) {
      ui.alert('エラー', 'Grok-4 APIキーが設定されていません。\n「⚙️ Grok APIキー設定」で設定してください。', ui.ButtonSet.OK);
      return;
    }
    
    // テスト用のサンプル投稿
    const testPosts = `### テスト投稿 1
- **投稿番号**: 123
- **企業/銘柄**: テスト企業
- **内容**: KEYWORD_1さんのポストは事実ではないです。`;
    
    const testToS = `## 禁止事項
1. 他人を中傷、名誉毀損する行為
2. 虚偽情報の投稿`;
    
    const testConfig = {
      companyName: 'テスト企業',
      keywords: ['KEYWORD_1', 'KEYWORD_2']
    };
    
    // Grok-4でテスト分析を実行
    const result = analyzeViolationsWithGrok(testToS, testPosts, testConfig);
    
    if (result) {
      Logger.log('Grok-4 APIテスト成功');
      const summary = result.substring(0, 300) + (result.length > 300 ? '...' : '');
      ui.alert(
        'Grok-4 APIテスト結果',
        `テストに成功しました！\n\n結果サンプル:\n${summary}\n\n詳細はログを確認してください。`,
        ui.ButtonSet.OK
      );
    } else {
      ui.alert('エラー', 'Grok-4 APIテストに失敗しました。\nログを確認してください。', ui.ButtonSet.OK);
    }
    
  } catch (error) {
    Logger.log('Grok-4 APIテストエラー: ' + error.toString());
    ui.alert('エラー', 'Grok-4 APIテスト中にエラーが発生しました：\n' + error.toString(), ui.ButtonSet.OK);
  }
}

/**
 * Grok APIキーを自動設定（指定されたキーで）
 */
function autoSetGrokAPIKey() {
  try {
    const apiKey = 'YOUR_GROK_API_KEY';
    PropertiesService.getScriptProperties().setProperty('GROK_API_KEY', apiKey);
    Logger.log('Grok-4 APIキーが自動設定されました');
    
    const ui = SpreadsheetApp.getUi();
    ui.alert('成功', 'Grok-4 APIキーが設定されました！\n違反分析でGrok-4が優先的に使用されます。', ui.ButtonSet.OK);
    
    return true;
  } catch (error) {
    Logger.log('Grok APIキー自動設定エラー: ' + error.toString());
    return false;
  }
}

/**
 * Grok APIキーを手動設定
 */
function setGrokAPIKey() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    // 現在の設定を確認
    const currentKey = CONFIG.GROK_API_KEY;
    const currentStatus = currentKey ? '設定済み' : '未設定';
    
    const response = ui.prompt(
      'Grok-4 APIキー設定',
      `現在の状態: ${currentStatus}\n\nGrok-4 APIキーを入力してください：\n(空にすると削除されます)`,
      ui.ButtonSet.OK_CANCEL
    );
    
    if (response.getSelectedButton() === ui.Button.OK) {
      const newKey = response.getResponseText().trim();
      
      if (newKey === '') {
        // キーを削除
        PropertiesService.getScriptProperties().deleteProperty('GROK_API_KEY');
        Logger.log('Grok-4 APIキーが削除されました');
        ui.alert('成功', 'Grok-4 APIキーが削除されました。\n違反分析はPerplexity APIで実行されます。', ui.ButtonSet.OK);
      } else {
        // 新しいキーを設定
        PropertiesService.getScriptProperties().setProperty('GROK_API_KEY', newKey);
        Logger.log('Grok-4 APIキーが設定されました');
        ui.alert('成功', 'Grok-4 APIキーが設定されました。\n違反分析で優先的に使用されます。\n\nテストは「🤖 Grok-4 APIテスト」で実行できます。', ui.ButtonSet.OK);
      }
    } else {
      ui.alert('情報', 'Grok-4 APIキーの設定がキャンセルされました。', ui.ButtonSet.OK);
    }
    
  } catch (error) {
    Logger.log('Grok APIキー設定エラー: ' + error.toString());
    ui.alert('エラー', 'Grok APIキー設定中にエラーが発生しました：\n' + error.toString(), ui.ButtonSet.OK);
  }
}

/**
 * HTML構造デバッグテスト
 */
function debugHTMLStructureTest() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    Logger.log('=== HTML構造デバッグテスト開始 ===');
    
    const testUrl = 'https://finance.yahoo.co.jp/cm/personal/history/comment?user=ACCOUNT_HASH_PLACEHOLDER&sort=2';
    const content = fetchPageContent(testUrl);
    
    if (!content) {
      ui.alert('エラー', 'ページの取得に失敗しました。', ui.ButtonSet.OK);
      return;
    }
    
    // HTML構造を詳細にデバッグ
    debugHTMLStructure(content);
    
    // commentBoxの数をカウント
    const commentBoxes = content.match(/<li\s+class="commentBox">/gi);
    const commentBoxCount = commentBoxes ? commentBoxes.length : 0;
    
    // 日付パターンをチェック
    const datePatterns = content.match(/\d{4}\/\d{2}\/\d{2}\s+\d{2}:\d{2}/g);
    const dateCount = datePatterns ? datePatterns.length : 0;
    
    // 投稿番号パターンをチェック
    const numberPatterns = content.match(/No\.(\d+)/g);
    const numberCount = numberPatterns ? numberPatterns.length : 0;
    
    // 企業名パターンをチェック
    const companyPatterns = content.match(/[\(（][^\)）]*[\)）]/g);
    const companyCount = companyPatterns ? companyPatterns.length : 0;
    
    // 結果をログとダイアログで表示
    const summary = `HTML構造解析結果:
コンテンツサイズ: ${content.length}文字
commentBox: ${commentBoxCount}個
日付パターン: ${dateCount}個
投稿番号: ${numberCount}個
企業名パターン: ${companyCount}個`;
    
    Logger.log(summary);
    
    // サンプルデータを抜粋
    if (datePatterns && datePatterns.length > 0) {
      Logger.log('日付サンプル: ' + datePatterns.slice(0, 3).join(', '));
    }
    if (numberPatterns && numberPatterns.length > 0) {
      Logger.log('投稿番号サンプル: ' + numberPatterns.slice(0, 3).join(', '));
    }
    if (companyPatterns && companyPatterns.length > 0) {
      Logger.log('企業名サンプル: ' + companyPatterns.slice(0, 3).join(', '));
    }
    
    ui.alert(
      'HTML構造デバッグ結果',
      summary + '\n\n詳細はログを確認してください。',
      ui.ButtonSet.OK
    );
    
  } catch (error) {
    Logger.log('HTML構造デバッグテストエラー: ' + error.toString());
    ui.alert('エラー', 'HTML構造デバッグテスト中にエラーが発生しました：\n' + error.toString(), ui.ButtonSet.OK);
  }
}

/**
 * メニューを追加（スプレッドシート開いた時に実行）
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🤖 Perplexity連携')
    .addItem('📋 初期設定', 'initialSetup')
    .addSeparator()
    .addSubMenu(ui.createMenu('▶️ 実行メニュー')
      .addItem('🔄 両方同時実行', 'executeMainProcess')
    .addSeparator()
      .addItem('🔍 書き込み検索のみ', 'executeKeywordSearchOnly')
      .addItem('⚠️ 規約違反チェックのみ', 'executeViolationCheckOnly')
    )
    .addSeparator()
    .addSubMenu(ui.createMenu('🔧 シート管理')
      .addItem('検索設定シート作成', 'createSearchSheet')
      .addItem('規約チェックシート作成', 'createViolationCheckSheet')
    )
    .addSeparator()
    .addSubMenu(ui.createMenu('🧪 テストメニュー')
      .addItem('APIキーのテスト', 'testAPIKey')
      .addItem('テスト実行（APIなし）', 'testRun')
      .addSeparator()
      .addItem('🕷️ ウェブスクレイピングテスト', 'testWebScraping')
      .addItem('📝 ページ取得テスト', 'testPageFetch')
      .addItem('🔍 HTML構造デバッグ', 'debugHTMLStructureTest')
      .addSeparator()
      .addItem('🤖 Grok-4 APIテスト', 'testGrokAPI')
      .addItem('⚙️ Grok APIキー設定', 'setGrokAPIKey')
      .addItem('🚀 Grok APIキー自動設定', 'autoSetGrokAPIKey')
    )
    .addToUi();
}