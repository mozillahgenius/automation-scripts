/**
 * Perplexity API検証用テストコード
 * このファイルをGASに追加して実行し、API接続と応答を確認
 */

/**
 * シンプルなAPIテスト
 */
function testPerplexitySimple() {
  console.log('=== Perplexity APIシンプルテスト ===');
  
  const apiKey = 'YOUR_PERPLEXITY_API_KEY';
  const apiUrl = 'https://api.perplexity.ai/chat/completions';
  
  // 最もシンプルなリクエスト（動作確認済みのsonarモデルを使用）
  const payload = {
    model: 'sonar',
    messages: [
      {
        role: 'user',
        content: 'What are the latest news about Toyota in the last week? Return as JSON with articles array.'
      }
    ]
  };
  
  const options = {
    method: 'post',
    headers: {
      'Authorization': `Bearer ${apiKey}`,
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  
  try {
    console.log('リクエスト送信中...');
    const response = UrlFetchApp.fetch(apiUrl, options);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();
    
    console.log(`レスポンスコード: ${responseCode}`);
    
    if (responseCode === 200) {
      const data = JSON.parse(responseText);
      console.log('成功！レスポンス構造:');
      console.log('- choices数:', data.choices?.length);
      console.log('- citations数:', data.citations?.length);
      console.log('- usage:', data.usage);
      
      if (data.choices && data.choices[0]) {
        const content = data.choices[0].message.content;
        console.log('\nコンテンツ（最初の500文字）:');
        console.log(content.substring(0, 500));
      }
      
      if (data.citations && data.citations.length > 0) {
        console.log('\n引用元（最初の3件）:');
        data.citations.slice(0, 3).forEach(c => {
          console.log(`- ${c.title || 'No title'}: ${c.url || 'No URL'}`);
        });
      }
      
      return data;
    } else {
      console.error('エラーレスポンス:', responseText);
      return null;
    }
    
  } catch (error) {
    console.error('実行エラー:', error);
    return null;
  }
}

/**
 * 日本語での検索テスト
 */
function testPerplexityJapanese() {
  console.log('=== 日本語検索テスト ===');
  
  const apiKey = 'YOUR_PERPLEXITY_API_KEY';
  const apiUrl = 'https://api.perplexity.ai/chat/completions';
  
  const payload = {
    model: 'sonar-pro',  // 動作確認済みのsonar-proモデルを使用
    messages: [
      {
        role: 'system',
        content: 'You are a Japanese news researcher. Always search for Japanese news sources.'
      },
      {
        role: 'user',
        content: 'トヨタ自動車に関する最近の批判的なニュースを検索してください。JSON形式で返してください。'
      }
    ],
    temperature: 0.1,
    return_citations: true
  };
  
  const options = {
    method: 'post',
    headers: {
      'Authorization': `Bearer ${apiKey}`,
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  
  try {
    const response = UrlFetchApp.fetch(apiUrl, options);
    const data = JSON.parse(response.getContentText());
    
    if (response.getResponseCode() === 200) {
      console.log('成功！');
      const content = data.choices[0].message.content;
      
      // JSONを抽出
      const jsonMatch = content.match(/\{[\s\S]*\}/);
      if (jsonMatch) {
        try {
          const articles = JSON.parse(jsonMatch[0]);
          console.log('記事数:', articles.articles?.length || 0);
          if (articles.articles && articles.articles[0]) {
            console.log('最初の記事:', articles.articles[0]);
          }
        } catch (e) {
          console.log('JSON解析エラー:', e);
        }
      }
      
      return data;
    } else {
      console.error('エラー:', data);
      return null;
    }
    
  } catch (error) {
    console.error('実行エラー:', error);
    return null;
  }
}

/**
 * 各モデルの動作確認
 */
function testAllModels() {
  console.log('=== 全モデルテスト ===');
  
  // 動作確認済みのモデルのみをテスト
  const models = [
    'sonar-pro',  // 動作確認済み
    'sonar'       // 動作確認済み
  ];
  
  const apiKey = 'YOUR_PERPLEXITY_API_KEY';
  const apiUrl = 'https://api.perplexity.ai/chat/completions';
  
  const results = {};
  
  models.forEach(model => {
    console.log(`\nテスト中: ${model}`);
    
    const payload = {
      model: model,
      messages: [
        {
          role: 'user',
          content: 'Search for recent Toyota news. Return one article in JSON format.'
        }
      ],
      max_tokens: 500
    };
    
    const options = {
      method: 'post',
      headers: {
        'Authorization': `Bearer ${apiKey}`,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };
    
    try {
      const response = UrlFetchApp.fetch(apiUrl, options);
      const responseCode = response.getResponseCode();
      
      if (responseCode === 200) {
        results[model] = '✓ 成功';
        console.log(`✓ ${model}: 成功`);
      } else {
        const errorText = response.getContentText();
        results[model] = `✗ エラー ${responseCode}`;
        console.log(`✗ ${model}: エラー ${responseCode}`);
        
        // エラー詳細を解析
        try {
          const errorData = JSON.parse(errorText);
          if (errorData.error?.message) {
            console.log(`  詳細: ${errorData.error.message}`);
          }
        } catch (e) {
          // JSON解析エラーは無視
        }
      }
    } catch (error) {
      results[model] = `✗ 例外: ${error.message}`;
      console.log(`✗ ${model}: 例外 ${error.message}`);
    }
    
    // レート制限対策
    Utilities.sleep(1000);
  });
  
  console.log('\n=== 結果サマリー ===');
  Object.entries(results).forEach(([model, result]) => {
    console.log(`${model}: ${result}`);
  });
  
  return results;
}

/**
 * 実際の記事検索をシミュレート
 */
function testActualSearch() {
  console.log('=== 実際の検索シミュレーション ===');
  
  const apiKey = 'YOUR_PERPLEXITY_API_KEY';
  const companyName = 'トヨタ自動車';
  const daysBack = 7;
  
  const endDate = new Date();
  const startDate = new Date();
  startDate.setDate(startDate.getDate() - daysBack);
  
  const query = `
Search for critical or negative news articles about ${companyName} from the last ${daysBack} days.
Focus on: scandals, lawsuits, regulatory issues, product recalls, criticism.
Return the results in JSON format with this structure:
{
  "articles": [
    {
      "date": "YYYY-MM-DD",
      "title": "exact article title",
      "summary": "article summary",
      "author": "author name or empty",
      "url": "article URL",
      "source": "news source",
      "criticalScore": 7
    }
  ]
}
Include up to 5 articles, sorted by relevance.`;

  const payload = {
    model: 'sonar-pro',  // 動作確認済みのsonar-proモデルを使用
    messages: [
      {
        role: 'system',
        content: 'You are a news research assistant. Search for real news articles and return actual results with valid URLs.'
      },
      {
        role: 'user',
        content: query
      }
    ],
    temperature: 0.1,
    max_tokens: 2000,
    return_citations: true,
    search_recency_filter: 'week'
  };
  
  const options = {
    method: 'post',
    headers: {
      'Authorization': `Bearer ${apiKey}`,
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  
  try {
    console.log('検索実行中...');
    const startTime = new Date().getTime();
    
    const response = UrlFetchApp.fetch('https://api.perplexity.ai/chat/completions', options);
    
    const responseTime = new Date().getTime() - startTime;
    console.log(`レスポンス時間: ${responseTime}ms`);
    
    if (response.getResponseCode() === 200) {
      const data = JSON.parse(response.getContentText());
      const content = data.choices[0].message.content;
      
      console.log('レスポンス受信成功');
      console.log('引用元数:', data.citations?.length || 0);
      
      // JSON抽出
      const jsonMatch = content.match(/\{[\s\S]*\}/);
      if (jsonMatch) {
        try {
          const result = JSON.parse(jsonMatch[0]);
          console.log('\n検索結果:');
          console.log('記事数:', result.articles?.length || 0);
          
          if (result.articles && result.articles.length > 0) {
            result.articles.forEach((article, i) => {
              console.log(`\n記事${i + 1}:`);
              console.log('  タイトル:', article.title);
              console.log('  日付:', article.date);
              console.log('  URL:', article.url);
              console.log('  スコア:', article.criticalScore);
            });
          } else {
            console.log('記事が見つかりませんでした');
          }
          
          return result;
        } catch (e) {
          console.error('JSON解析エラー:', e);
          console.log('生のコンテンツ:', content.substring(0, 500));
        }
      } else {
        console.log('JSON形式のデータが見つかりません');
        console.log('生のコンテンツ:', content.substring(0, 500));
      }
    } else {
      console.error('APIエラー:', response.getResponseCode());
    }
    
  } catch (error) {
    console.error('実行エラー:', error);
  }
  
  return null;
}

/**
 * メインテスト実行
 */
function runAllPerplexityTests() {
  console.log('===== Perplexity API完全テスト開始 =====\n');
  
  // 1. シンプルテスト
  console.log('1. シンプルテスト');
  testPerplexitySimple();
  
  Utilities.sleep(2000);
  
  // 2. 日本語テスト
  console.log('\n2. 日本語テスト');
  testPerplexityJapanese();
  
  Utilities.sleep(2000);
  
  // 3. 実際の検索
  console.log('\n3. 実際の検索');
  testActualSearch();
  
  console.log('\n===== テスト完了 =====');
}