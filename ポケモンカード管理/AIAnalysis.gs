// ==============================
// AI画像解析（Perplexity API）
// ==============================

async function analyzeCardWithAI(driveFile, config) {
  const apiKey = config.PERPLEXITY_API_KEY;
  const model = config.AI_MODEL || 'llama-3.1-sonar-large-128k-online';

  // 画像をBase64エンコード
  const base64Image = Utilities.base64Encode(driveFile.blob.getBytes());
  const imageDataUrl = `data:${driveFile.blob.getContentType()};base64,${base64Image}`;

  // AI判定プロンプト
  const prompt = getCardAnalysisPrompt();

  try {
    // リトライロジック（最大3回）
    let retryCount = 0;
    const maxRetries = 3;

    while (retryCount < maxRetries) {
      try {
        const result = await callPerplexityVision(apiKey, model, imageDataUrl, prompt);
        const cardData = parseAIResponse(result);

        // Drive URLを追加
        cardData.driveUrl = driveFile.viewUrl;
        cardData.driveFileId = driveFile.id;
        cardData.originalFileName = driveFile.name;

        return cardData;

      } catch (error) {
        retryCount++;
        if (retryCount >= maxRetries) {
          throw error;
        }

        // エクスポネンシャルバックオフ
        const waitTime = Math.pow(2, retryCount) * 1000; // 2秒, 4秒, 8秒
        console.log(`AI判定リトライ ${retryCount}/${maxRetries}, 待機時間: ${waitTime}ms`);
        Utilities.sleep(waitTime);
      }
    }

  } catch (error) {
    console.error('AI判定エラー:', error);

    // フォールバック: 基本情報のみ返す
    return {
      name: driveFile.name.replace(/\.[^/.]+$/, ''), // 拡張子を除去
      game: 'Unknown',
      set: null,
      number: null,
      rarity: null,
      language: null,
      condition: null,
      price: null,
      notes: `AI判定失敗: ${error.toString()}`,
      status: '要確認',
      driveUrl: driveFile.viewUrl,
      driveFileId: driveFile.id,
      originalFileName: driveFile.name
    };
  }
}

function callPerplexityVision(apiKey, model, imageDataUrl, prompt) {
  const url = 'https://api.perplexity.ai/chat/completions';

  // 画像の説明を含むプロンプトを作成
  const imagePrompt = `画像を分析してください。これはトレーディングカードの画像です。\n\n${prompt}\n\n注意: 画像の内容から判断できる情報のみを抽出してください。`;

  const payload = {
    model: model,
    messages: [
      {
        role: 'system',
        content: 'あなたはトレーディングカードの専門家です。カード情報を正確に抽出してJSON形式で返答してください。'
      },
      {
        role: 'user',
        content: imagePrompt
      }
    ],
    max_tokens: 1000,
    temperature: 0.1, // 決定的な出力のため低めに設定
    top_p: 0.1,
    // 画像URLを検索クエリとして使用（Perplexityの画像検索機能を活用）
    search_domain_filter: [],
    return_images: false,
    return_related_questions: false
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

  const responseCode = response.getResponseCode();
  if (responseCode !== 200) {
    const errorBody = response.getContentText();
    throw new Error(`Perplexity API Error (${responseCode}): ${errorBody}`);
  }

  const result = JSON.parse(response.getContentText());

  // Perplexity APIのレスポンス構造を確認
  if (result.choices && result.choices.length > 0) {
    return result.choices[0].message.content;
  } else {
    throw new Error('Unexpected Perplexity API response structure');
  }
}

function getCardAnalysisPrompt() {
  return `
この画像のトレーディングカードを分析して、以下の情報をJSON形式で抽出してください。
不明な項目はnullとしてください。複数の可能性がある場合はnotesに記載してください。

必須フィールド:
- game: カードゲーム名（"Pokemon", "Yu-Gi-Oh!", "MTG", "Other"のいずれか）
- name: カード名（日本語または英語）
- set: セット名またはエキスパンション名
- number: カード番号（コレクター番号）
- rarity: レアリティ（C/UC/R/RR/SR/UR/HR/SAR等）
- language: 言語（"JP", "EN", "CN", "KR"等）
- condition: カードの状態推定（"新品", "美品", "良好", "やや傷", "傷あり", "ジャンク"のいずれか）
- notes: その他の情報、特記事項、不確実な情報

出力例:
{
  "game": "Pokemon",
  "name": "ピカチュウ",
  "set": "ポケモンカード151",
  "number": "025/165",
  "rarity": "R",
  "language": "JP",
  "condition": "美品",
  "notes": "ホロカード、中央に小さな白かけあり"
}

必ずJSON形式のみで回答してください。説明文は不要です。
`;
}

function parseAIResponse(aiResponse) {
  try {
    // JSONを抽出（マークダウンコードブロックを考慮）
    let jsonStr = aiResponse;

    // コードブロックで囲まれている場合の処理
    if (aiResponse.includes('```json')) {
      const match = aiResponse.match(/```json\n?([\s\S]*?)\n?```/);
      if (match) jsonStr = match[1];
    } else if (aiResponse.includes('```')) {
      const match = aiResponse.match(/```\n?([\s\S]*?)\n?```/);
      if (match) jsonStr = match[1];
    }

    // JSON以外のテキストが前後にある場合の処理
    const jsonMatch = jsonStr.match(/\{[\s\S]*\}/);
    if (jsonMatch) {
      jsonStr = jsonMatch[0];
    }

    const data = JSON.parse(jsonStr);

    // デフォルト値の設定
    return {
      name: data.name || 'Unknown Card',
      game: data.game || 'Unknown',
      set: data.set || null,
      number: data.number || null,
      rarity: data.rarity || null,
      language: data.language || null,
      condition: data.condition || null,
      price: data.price || null,
      notes: data.notes || null,
      status: '要確認' // デフォルトステータス
    };

  } catch (error) {
    console.error('AI応答のパースエラー:', error, 'Response:', aiResponse);
    throw new Error('AI応答の解析に失敗しました');
  }
}

// ==============================
// 外部API補完（オプション）
// ==============================

function enrichCardData(cardData) {
  try {
    switch (cardData.game) {
      case 'Pokemon':
        enrichPokemonCard(cardData);
        break;
      case 'Yu-Gi-Oh!':
        enrichYugiohCard(cardData);
        break;
      case 'MTG':
        enrichMTGCard(cardData);
        break;
      default:
        console.log('未対応のカードゲーム:', cardData.game);
    }
  } catch (error) {
    console.error('カードデータ補完エラー:', error);
    cardData.notes = (cardData.notes || '') + `\n補完エラー: ${error.toString()}`;
  }
}

function enrichPokemonCard(cardData) {
  // Pokemon TCG API を使用した補完
  if (!cardData.number || !cardData.set) {
    return;
  }

  const apiUrl = `https://api.pokemontcg.io/v2/cards?q=number:${cardData.number} set.name:"${cardData.set}"`;

  const response = UrlFetchApp.fetch(apiUrl, {
    muteHttpExceptions: true
  });

  if (response.getResponseCode() === 200) {
    const result = JSON.parse(response.getContentText());
    if (result.data && result.data.length > 0) {
      const apiCard = result.data[0];

      // APIデータで補完
      cardData.name = cardData.name || apiCard.name;
      cardData.set = apiCard.set.name;
      cardData.number = apiCard.number;
      cardData.rarity = cardData.rarity || apiCard.rarity;

      // 価格情報（TCGPlayer）
      if (apiCard.tcgplayer && apiCard.tcgplayer.prices) {
        const prices = apiCard.tcgplayer.prices;
        const marketPrice = prices.holofoil?.market || prices.normal?.market;
        if (marketPrice) {
          cardData.price = `$${marketPrice}`;
        }
      }
    }
  }
}

function enrichYugiohCard(cardData) {
  // YGOPRODeck API を使用した補完
  if (!cardData.name) {
    return;
  }

  const apiUrl = `https://db.ygoprodeck.com/api/v7/cardinfo.php?name=${encodeURIComponent(cardData.name)}`;

  const response = UrlFetchApp.fetch(apiUrl, {
    muteHttpExceptions: true
  });

  if (response.getResponseCode() === 200) {
    const result = JSON.parse(response.getContentText());
    if (result.data && result.data.length > 0) {
      const apiCard = result.data[0];

      // APIデータで補完
      cardData.name = apiCard.name;
      cardData.set = cardData.set || apiCard.card_sets?.[0]?.set_name;
      cardData.number = cardData.number || apiCard.card_sets?.[0]?.set_code;
      cardData.rarity = cardData.rarity || apiCard.card_sets?.[0]?.set_rarity;

      // 価格情報
      if (apiCard.card_prices && apiCard.card_prices.length > 0) {
        const price = apiCard.card_prices[0];
        cardData.price = `$${price.tcgplayer_price}`;
      }
    }
  }
}

function enrichMTGCard(cardData) {
  // Scryfall API を使用した補完
  if (!cardData.name) {
    return;
  }

  const apiUrl = `https://api.scryfall.com/cards/search?q=name:"${encodeURIComponent(cardData.name)}"`;

  const response = UrlFetchApp.fetch(apiUrl, {
    muteHttpExceptions: true
  });

  if (response.getResponseCode() === 200) {
    const result = JSON.parse(response.getContentText());
    if (result.data && result.data.length > 0) {
      const apiCard = result.data[0];

      // APIデータで補完
      cardData.name = apiCard.name;
      cardData.set = cardData.set || apiCard.set_name;
      cardData.number = cardData.number || apiCard.collector_number;
      cardData.rarity = cardData.rarity || apiCard.rarity;
      cardData.language = cardData.language || apiCard.lang.toUpperCase();

      // 価格情報
      if (apiCard.prices) {
        const price = apiCard.prices.usd || apiCard.prices.usd_foil;
        if (price) {
          cardData.price = `$${price}`;
        }
      }
    }
  }
}

// ==============================
// テスト関数
// ==============================

function testPerplexityConnection(apiKey) {
  const url = 'https://api.perplexity.ai/chat/completions';

  try {
    // シンプルなテストリクエスト
    const payload = {
      model: 'llama-3.1-sonar-small-128k-online',
      messages: [
        {
          role: 'user',
          content: 'Hello'
        }
      ],
      max_tokens: 10
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

    return response.getResponseCode() === 200;
  } catch (error) {
    return false;
  }
}

function testAIAnalysis() {
  const config = getConfig();

  // テスト用の画像URLを使用
  const testImageUrl = 'https://via.placeholder.com/400x600.png?text=Test+Card';

  console.log('AI分析テスト開始');

  try {
    const response = UrlFetchApp.fetch(testImageUrl);
    const blob = response.getBlob();

    const testFile = {
      id: 'test_id',
      name: 'test_card.png',
      viewUrl: 'https://example.com/test',
      blob: blob
    };

    const result = analyzeCardWithAI(testFile, config);
    console.log('AI分析結果:', result);

    return result;

  } catch (error) {
    console.error('AI分析テストエラー:', error);
    return {
      success: false,
      error: error.toString()
    };
  }
}

// ==============================
// Perplexityモデル一覧
// ==============================

function getAvailablePerplexityModels() {
  // Perplexityで利用可能なモデル
  return [
    'llama-3.1-sonar-large-128k-online', // sonar-pro (推奨)
    'llama-3.1-sonar-small-128k-online', // sonar
    'llama-3.1-sonar-huge-128k-online' // sonar-huge
  ];
}

function listPerplexityModels() {
  const models = getAvailablePerplexityModels();
  console.log('利用可能なPerplexityモデル:');
  models.forEach(model => {
    console.log(`- ${model}`);
  });
  return models;
}