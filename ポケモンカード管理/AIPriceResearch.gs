// ==============================
// AI価格調査システム
// ==============================

/**
 * AIを使って最新価格を調査
 */
function getCardPriceByAI(cardData) {
  console.log(`AI価格調査開始: ${cardData.name} (${cardData.number})`);

  const config = getConfig();

  // OpenAI APIキーを確認
  if (!config.OPENAI_API_KEY) {
    console.error('OpenAI APIキーが設定されていません');
    return estimatePriceByRarity(cardData.rarity || 'R');
  }

  try {
    // 価格調査用のプロンプトを作成
    const prompt = createPriceResearchPrompt(cardData);

    // OpenAI APIで価格を調査
    const priceInfo = callOpenAIForPrice(config.OPENAI_API_KEY, prompt, cardData);

    if (priceInfo) {
      // 価格情報を解析して設定
      applyAIPriceInfo(cardData, priceInfo);
      console.log(`AI価格調査成功: ¥${cardData.price}`);
    } else {
      // AI調査失敗時はレアリティから推定
      cardData.price = estimatePriceByRarity(cardData.rarity || 'R');
      cardData.priceEstimated = true;
      console.log(`価格推定: ¥${cardData.price}`);
    }

  } catch (error) {
    console.error('AI価格調査エラー:', error);
    cardData.price = estimatePriceByRarity(cardData.rarity || 'R');
    cardData.priceError = error.toString();
  }

  // 価格履歴と予測を生成
  if (cardData.price) {
    cardData.priceHistory = generateAIPriceHistory(cardData);
    cardData.pricePrediction = generateAIPricePrediction(cardData);
  }

  return cardData.price;
}

/**
 * 価格調査用プロンプトを作成
 */
function createPriceResearchPrompt(cardData) {
  const today = new Date().toLocaleDateString('ja-JP');

  let prompt = `
今日は${today}です。以下のトレーディングカードの現在の市場価格を調査してください。

【カード情報】
- カード名: ${cardData.name || '不明'}
- ゲーム: ${cardData.game || '不明'}
- セット/シリーズ: ${cardData.set || '不明'}
- カード番号: ${cardData.number || '不明'}
- レアリティ: ${cardData.rarity || '不明'}
- 言語: ${cardData.language || '日本語'}
- 状態: ${cardData.condition || '美品'}

【調査内容】
1. 現在の日本市場での販売価格（円）
2. 最近の取引相場
3. 価格トレンド（上昇/下降/安定）
4. 3ヶ月前、6ヶ月前、12ヶ月前の推定価格
5. 6ヶ月後、12ヶ月後の価格予測

【重要な判断基準】
- メルカリ、ヤフオク、カードショップの相場
- 同じカード番号の正確な価格
- プロモカードの場合は配布時期も考慮
- 人気キャラクター（ピカチュウ、リザードン等）は高値傾向

必ず以下のJSON形式で回答してください：
{
  "currentPrice": 現在価格（数値、円）,
  "marketPrice": 市場平均価格（数値、円）,
  "trend": "上昇" | "下降" | "安定" | "変動",
  "confidence": "高" | "中" | "低",
  "priceHistory": {
    "12monthsAgo": 12ヶ月前価格,
    "6monthsAgo": 6ヶ月前価格,
    "3monthsAgo": 3ヶ月前価格
  },
  "pricePrediction": {
    "6monthsLater": 6ヶ月後予測,
    "12monthsLater": 12ヶ月後予測
  },
  "notes": "価格判定の根拠や特記事項"
}
`;

  // 特定のカードに関する追加情報
  if (cardData.game === 'ポケモン' || cardData.game === 'Pokemon') {
    prompt += '\n\n【ポケモンカード特有の考慮事項】\n';
    prompt += '- SAR、HR、SR、CSRは特に高値\n';
    prompt += '- 女性トレーナーカードは高値傾向\n';
    prompt += '- 最新弾は発売直後高く、徐々に下落\n';
    prompt += '- 絶版セットは価格上昇傾向\n';
  }

  return prompt;
}

/**
 * OpenAI APIで価格を調査
 */
function callOpenAIForPrice(apiKey, prompt, cardData) {
  const url = 'https://api.openai.com/v1/chat/completions';

  try {
    const payload = {
      model: 'gpt-4o-mini',
      messages: [
        {
          role: 'system',
          content: 'あなたはトレーディングカード市場の専門家です。最新の市場動向と価格情報に詳しく、正確な価格査定ができます。Web検索結果や最新の取引データを基に回答してください。'
        },
        {
          role: 'user',
          content: prompt
        }
      ],
      temperature: 0.3,  // より確実な回答を得るため低めに設定
      max_tokens: 1000,
      response_format: { type: "json_object" }  // JSON形式を強制
    };

    const response = UrlFetchApp.fetch(url, {
      method: 'POST',
      headers: {
        'Authorization': 'Bearer ' + apiKey,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    if (response.getResponseCode() === 200) {
      const result = JSON.parse(response.getContentText());
      const content = result.choices[0].message.content;

      console.log('AI回答:', content);

      // JSON形式の回答を解析
      try {
        const priceInfo = JSON.parse(content);
        return priceInfo;
      } catch (e) {
        console.error('JSON解析エラー:', e);
        // JSON解析失敗時は文字列から価格を抽出
        return extractPriceFromText(content);
      }
    } else {
      console.error('OpenAI APIエラー:', response.getContentText());
      return null;
    }

  } catch (error) {
    console.error('価格調査APIエラー:', error);
    return null;
  }
}

/**
 * AI価格情報をカードデータに適用
 */
function applyAIPriceInfo(cardData, priceInfo) {
  // 現在価格を設定
  if (priceInfo.currentPrice) {
    cardData.price = Math.max(priceInfo.currentPrice, 50);  // 最低50円
    cardData.priceSource = 'AI調査';
  } else if (priceInfo.marketPrice) {
    cardData.price = Math.max(priceInfo.marketPrice, 50);
    cardData.priceSource = 'AI市場価格';
  } else {
    cardData.price = estimatePriceByRarity(cardData.rarity || 'R');
    cardData.priceSource = '推定';
  }

  // その他の価格情報を設定
  cardData.marketPrice = priceInfo.marketPrice || cardData.price;
  cardData.priceTrend = priceInfo.trend || '不明';
  cardData.priceConfidence = priceInfo.confidence || '低';

  // 価格履歴を設定
  if (priceInfo.priceHistory) {
    cardData.priceHistory = {
      '12ヶ月前': priceInfo.priceHistory['12monthsAgo'] || 0,
      '9ヶ月前': priceInfo.priceHistory['9monthsAgo'] ||
                  Math.round((priceInfo.priceHistory['12monthsAgo'] + priceInfo.priceHistory['6monthsAgo']) / 2) || 0,
      '6ヶ月前': priceInfo.priceHistory['6monthsAgo'] || 0,
      '3ヶ月前': priceInfo.priceHistory['3monthsAgo'] || 0,
      '現在': cardData.price,
      'trend': cardData.priceTrend
    };
  }

  // 価格予測を設定
  if (priceInfo.pricePrediction) {
    cardData.pricePrediction = {
      '6ヶ月後': priceInfo.pricePrediction['6monthsLater'] || cardData.price,
      '12ヶ月後': priceInfo.pricePrediction['12monthsLater'] || cardData.price
    };
  }

  // AIの判定根拠を記録
  if (priceInfo.notes) {
    cardData.priceNotes = priceInfo.notes;
  }

  console.log(`AI価格設定: ¥${cardData.price} (${cardData.priceConfidence}信頼度)`);
}

/**
 * テキストから価格を抽出（フォールバック）
 */
function extractPriceFromText(text) {
  const priceInfo = {
    currentPrice: 0,
    marketPrice: 0,
    trend: '不明'
  };

  // 価格パターンを検索
  const pricePatterns = [
    /(\d{1,6})[,，]?(\d{3})?円/g,
    /￥(\d{1,6})[,，]?(\d{3})?/g,
    /¥(\d{1,6})[,，]?(\d{3})?/g
  ];

  let prices = [];
  pricePatterns.forEach(pattern => {
    const matches = text.matchAll(pattern);
    for (const match of matches) {
      let price = match[1];
      if (match[2]) {
        price += match[2];
      }
      prices.push(parseInt(price));
    }
  });

  if (prices.length > 0) {
    // 中央値を現在価格とする
    prices.sort((a, b) => a - b);
    priceInfo.currentPrice = prices[Math.floor(prices.length / 2)];
    priceInfo.marketPrice = Math.round(prices.reduce((a, b) => a + b, 0) / prices.length);
  }

  // トレンドを検出
  if (text.includes('上昇') || text.includes('高騰')) {
    priceInfo.trend = '上昇';
  } else if (text.includes('下降') || text.includes('下落')) {
    priceInfo.trend = '下降';
  } else if (text.includes('安定')) {
    priceInfo.trend = '安定';
  }

  return priceInfo;
}

/**
 * AI基準の価格履歴生成
 */
function generateAIPriceHistory(cardData) {
  const currentPrice = cardData.price || 100;
  const trend = cardData.priceTrend || '安定';

  // すでにAIが履歴を提供している場合はそれを使用
  if (cardData.priceHistory) {
    return cardData.priceHistory;
  }

  // トレンドに基づいて履歴を生成
  let history = {};

  switch (trend) {
    case '上昇':
      history = {
        '12ヶ月前': Math.round(currentPrice * 0.4),
        '9ヶ月前': Math.round(currentPrice * 0.55),
        '6ヶ月前': Math.round(currentPrice * 0.7),
        '3ヶ月前': Math.round(currentPrice * 0.85),
        '現在': currentPrice,
        'trend': '上昇'
      };
      break;

    case '下降':
      history = {
        '12ヶ月前': Math.round(currentPrice * 2.0),
        '9ヶ月前': Math.round(currentPrice * 1.7),
        '6ヶ月前': Math.round(currentPrice * 1.4),
        '3ヶ月前': Math.round(currentPrice * 1.15),
        '現在': currentPrice,
        'trend': '下降'
      };
      break;

    default:
      history = {
        '12ヶ月前': Math.round(currentPrice * 0.9),
        '9ヶ月前': Math.round(currentPrice * 0.92),
        '6ヶ月前': Math.round(currentPrice * 0.95),
        '3ヶ月前': Math.round(currentPrice * 0.98),
        '現在': currentPrice,
        'trend': '安定'
      };
  }

  return history;
}

/**
 * AI基準の価格予測生成
 */
function generateAIPricePrediction(cardData) {
  const currentPrice = cardData.price || 100;
  const trend = cardData.priceTrend || '安定';

  // すでにAIが予測を提供している場合はそれを使用
  if (cardData.pricePrediction) {
    return cardData.pricePrediction;
  }

  let prediction = {};

  switch (trend) {
    case '上昇':
      prediction = {
        '6ヶ月後': Math.round(currentPrice * 1.2),
        '12ヶ月後': Math.round(currentPrice * 1.5)
      };
      break;

    case '下降':
      prediction = {
        '6ヶ月後': Math.round(currentPrice * 0.85),
        '12ヶ月後': Math.round(currentPrice * 0.7)
      };
      break;

    default:
      prediction = {
        '6ヶ月後': Math.round(currentPrice * 1.02),
        '12ヶ月後': Math.round(currentPrice * 1.05)
      };
  }

  return prediction;
}

/**
 * 改良版：enrichCardData（AI価格調査を使用）
 */
function enrichCardDataWithAI(cardData) {
  console.log(`カードデータ補完開始（AI版）: ${cardData.name}`);

  try {
    // AI価格調査を実行
    getCardPriceByAI(cardData);

    // 追加情報の補完（必要に応じて）
    if (!cardData.set && cardData.number) {
      // カード番号からセット情報を推測
      inferSetFromNumber(cardData);
    }

    console.log(`補完完了: ${cardData.name} - ¥${cardData.price}`);

  } catch (error) {
    console.error('AI補完エラー:', error);
    cardData.price = estimatePriceByRarity(cardData.rarity || 'R');
    cardData.priceError = error.toString();
  }

  return cardData;
}

/**
 * カード番号からセット情報を推測
 */
function inferSetFromNumber(cardData) {
  if (!cardData.number) return;

  // ポケモンカードのセット番号パターン
  const patterns = {
    'S': '剣・盾シリーズ',
    'SV': 'スカーレット&バイオレット',
    'PROMO': 'プロモカード',
    'SM': 'サン&ムーン',
    'XY': 'XY'
  };

  for (const [prefix, setName] of Object.entries(patterns)) {
    if (cardData.number.toUpperCase().includes(prefix)) {
      cardData.set = cardData.set || setName;
      break;
    }
  }
}

/**
 * AI価格調査のテスト
 */
function testAIPriceResearch() {
  console.log('=== AI価格調査テスト ===\n');

  const testCards = [
    {
      name: 'ピカチュウex',
      game: 'ポケモン',
      number: 'SV-P 001',
      set: 'スカーレット&バイオレット',
      rarity: 'SR'
    },
    {
      name: 'リザードン',
      game: 'ポケモン',
      number: '006/150',
      set: 'ポケモンカード151',
      rarity: 'R'
    }
  ];

  testCards.forEach((card, index) => {
    console.log(`\nテスト${index + 1}: ${card.name}`);

    // AI価格調査
    getCardPriceByAI(card);

    console.log('結果:');
    console.log(`  現在価格: ¥${card.price}`);
    console.log(`  価格トレンド: ${card.priceTrend}`);
    console.log(`  信頼度: ${card.priceConfidence}`);
    console.log(`  情報源: ${card.priceSource}`);

    if (card.priceHistory) {
      console.log('  価格履歴:');
      Object.entries(card.priceHistory).forEach(([period, price]) => {
        if (period !== 'trend') {
          console.log(`    ${period}: ¥${price}`);
        }
      });
    }

    if (card.pricePrediction) {
      console.log('  価格予測:');
      Object.entries(card.pricePrediction).forEach(([period, price]) => {
        console.log(`    ${period}: ¥${price}`);
      });
    }
  });

  return testCards;
}