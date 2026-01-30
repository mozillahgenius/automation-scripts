// ==============================
// 価格取得・変換システム（改良版）
// ==============================

/**
 * 価格取得の設定
 */
function setupPriceConfig() {
  const props = PropertiesService.getScriptProperties();

  // デフォルト設定
  const config = {
    // 為替レート（手動設定またはAPI取得）
    'USD_TO_JPY_RATE': '150',  // 1USD = 150JPY（デフォルト）
    'EUR_TO_JPY_RATE': '160',  // 1EUR = 160JPY

    // 価格API設定
    'USE_POKEMONTCG_API': 'true',
    'USE_YGOPRODECK_API': 'true',
    'USE_SCRYFALL_API': 'true',

    // 日本市場価格API
    'USE_MERCARI_API': 'false',  // メルカリ価格
    'USE_YAHOO_AUCTION': 'false', // ヤフオク価格

    // 価格取得の優先順位
    'PRICE_PRIORITY': 'japan_first'  // 'japan_first' or 'global_first'
  };

  Object.keys(config).forEach(key => {
    if (!props.getProperty(key)) {
      props.setProperty(key, config[key]);
    }
  });

  console.log('価格設定を初期化しました');
  return config;
}

/**
 * リアルタイム為替レート取得
 */
function getExchangeRate(from = 'USD', to = 'JPY') {
  try {
    // 無料の為替レートAPI
    const url = `https://api.exchangerate-api.com/v4/latest/${from}`;
    const response = UrlFetchApp.fetch(url, {
      muteHttpExceptions: true
    });

    if (response.getResponseCode() === 200) {
      const data = JSON.parse(response.getContentText());
      const rate = data.rates[to];

      console.log(`為替レート: 1 ${from} = ${rate} ${to}`);

      // レートを保存
      PropertiesService.getScriptProperties().setProperty(`${from}_TO_${to}_RATE`, rate.toString());

      return rate;
    }
  } catch (error) {
    console.error('為替レート取得エラー:', error);
  }

  // フォールバック：保存済みのレート
  const savedRate = PropertiesService.getScriptProperties().getProperty(`${from}_TO_${to}_RATE`);
  return savedRate ? parseFloat(savedRate) : 150;  // デフォルト150
}

/**
 * 改良版：カードデータ補完（価格重視）
 */
function enrichCardDataWithPrice(cardData) {
  console.log(`価格取得開始: ${cardData.name} (${cardData.number})`);

  try {
    switch (cardData.game) {
      case 'ポケモン':
      case 'Pokemon':
        getPokemonCardPrice(cardData);
        break;
      case '遊戯王':
      case 'Yu-Gi-Oh!':
        getYugiohCardPrice(cardData);
        break;
      case 'MTG':
        getMTGCardPrice(cardData);
        break;
      default:
        console.log('価格取得非対応:', cardData.game);
    }
  } catch (error) {
    console.error('価格取得エラー:', error);
    cardData.priceError = error.toString();
  }

  return cardData;
}

/**
 * ポケモンカード価格取得（改良版）
 */
function getPokemonCardPrice(cardData) {
  const results = {
    tcgPlayer: null,
    japan: null,
    converted: null
  };

  // 1. Pokemon TCG APIで価格取得
  if (cardData.number && cardData.set) {
    try {
      // カード番号とセット名で正確に検索
      const searchQuery = `number:${cardData.number}`;
      const apiUrl = `https://api.pokemontcg.io/v2/cards?q=${encodeURIComponent(searchQuery)}`;

      console.log(`API検索: ${searchQuery}`);

      const response = UrlFetchApp.fetch(apiUrl, {
        muteHttpExceptions: true
      });

      if (response.getResponseCode() === 200) {
        const result = JSON.parse(response.getContentText());

        if (result.data && result.data.length > 0) {
          // セット名でフィルタリング
          let apiCard = result.data.find(card =>
            card.set.name.includes(cardData.set) ||
            cardData.set.includes(card.set.name)
          );

          // 見つからない場合は最初のカードを使用
          if (!apiCard) {
            apiCard = result.data[0];
          }

          console.log(`カード発見: ${apiCard.name} (${apiCard.number}/${apiCard.set.name})`);

          // APIデータで補完
          cardData.name = cardData.name || apiCard.name;
          cardData.setCode = apiCard.set.id;
          cardData.setName = apiCard.set.name;
          cardData.number = apiCard.number;
          cardData.rarity = cardData.rarity || apiCard.rarity;
          cardData.artist = apiCard.artist;

          // TCGPlayer価格（USD）
          if (apiCard.tcgplayer && apiCard.tcgplayer.prices) {
            const prices = apiCard.tcgplayer.prices;

            // 価格優先順位：holofoil > reverseHolofoil > normal > unlimited
            const priceTypes = ['holofoil', 'reverseHolofoil', 'normal', 'unlimited'];

            for (const type of priceTypes) {
              if (prices[type] && prices[type].market) {
                results.tcgPlayer = prices[type].market;
                console.log(`TCGPlayer価格 (${type}): $${results.tcgPlayer}`);
                break;
              }
            }

            // 価格範囲も記録
            if (prices.holofoil) {
              cardData.priceRange = {
                low: prices.holofoil.low,
                mid: prices.holofoil.mid,
                high: prices.holofoil.high,
                market: prices.holofoil.market
              };
            }
          }

          // CardMarket価格（EUR）
          if (apiCard.cardmarket && apiCard.cardmarket.prices) {
            const cmPrice = apiCard.cardmarket.prices.averageSellPrice;
            if (cmPrice) {
              results.cardMarket = cmPrice;
              console.log(`CardMarket価格: €${cmPrice}`);
            }
          }
        }
      }
    } catch (error) {
      console.error('Pokemon TCG API エラー:', error);
    }
  }

  // 2. 日本市場価格の推定（オプション）
  if (cardData.language === 'Japanese' || cardData.language === '日本語') {
    results.japan = estimateJapanesePrice(cardData);
  }

  // 3. 価格を日本円に変換
  if (results.tcgPlayer) {
    const rate = getExchangeRate('USD', 'JPY');
    results.converted = Math.round(results.tcgPlayer * rate);

    cardData.price = results.converted;
    cardData.priceUSD = results.tcgPlayer;
    cardData.exchangeRate = rate;

    console.log(`価格変換: $${results.tcgPlayer} → ¥${results.converted} (レート: ${rate})`);
  } else if (results.japan) {
    cardData.price = results.japan;
    console.log(`日本市場価格: ¥${results.japan}`);
  } else {
    // 価格が見つからない場合はレアリティから推定
    cardData.price = estimatePriceByRarity(cardData.rarity);
    cardData.priceEstimated = true;
    console.log(`推定価格: ¥${cardData.price}`);
  }

  return results;
}

/**
 * 遊戯王カード価格取得（改良版）
 */
function getYugiohCardPrice(cardData) {
  if (!cardData.name) return;

  try {
    const apiUrl = `https://db.ygoprodeck.com/api/v7/cardinfo.php?name=${encodeURIComponent(cardData.name)}`;
    const response = UrlFetchApp.fetch(apiUrl, {
      muteHttpExceptions: true
    });

    if (response.getResponseCode() === 200) {
      const result = JSON.parse(response.getContentText());

      if (result.data && result.data.length > 0) {
        const apiCard = result.data[0];

        // カード番号でセットを特定
        let targetSet = null;
        if (cardData.number && apiCard.card_sets) {
          targetSet = apiCard.card_sets.find(set =>
            set.set_code === cardData.number ||
            set.set_code.includes(cardData.number) ||
            cardData.number.includes(set.set_code)
          );
        }

        // 見つからない場合は最初のセットを使用
        targetSet = targetSet || apiCard.card_sets?.[0];

        if (targetSet) {
          cardData.set = targetSet.set_name;
          cardData.number = targetSet.set_code;
          cardData.rarity = targetSet.set_rarity;

          // 価格（USD）
          const priceUSD = targetSet.set_price || apiCard.card_prices[0].tcgplayer_price;

          if (priceUSD) {
            const rate = getExchangeRate('USD', 'JPY');
            const priceJPY = Math.round(parseFloat(priceUSD) * rate);

            cardData.price = priceJPY;
            cardData.priceUSD = parseFloat(priceUSD);
            cardData.exchangeRate = rate;

            console.log(`遊戯王価格: $${priceUSD} → ¥${priceJPY}`);
          }
        }
      }
    }
  } catch (error) {
    console.error('YGOProDeck API エラー:', error);
  }
}

/**
 * MTGカード価格取得（改良版）
 */
function getMTGCardPrice(cardData) {
  if (!cardData.name) return;

  try {
    // Scryfall API
    const apiUrl = `https://api.scryfall.com/cards/named?fuzzy=${encodeURIComponent(cardData.name)}`;
    const response = UrlFetchApp.fetch(apiUrl, {
      muteHttpExceptions: true
    });

    if (response.getResponseCode() === 200) {
      const apiCard = JSON.parse(response.getContentText());

      // カード番号で確認
      if (cardData.number && apiCard.collector_number !== cardData.number) {
        // コレクター番号で再検索
        const searchUrl = `https://api.scryfall.com/cards/search?q=cn:${cardData.number}+name:"${cardData.name}"`;
        const searchResponse = UrlFetchApp.fetch(searchUrl, {
          muteHttpExceptions: true
        });

        if (searchResponse.getResponseCode() === 200) {
          const searchResult = JSON.parse(searchResponse.getContentText());
          if (searchResult.data && searchResult.data.length > 0) {
            apiCard = searchResult.data[0];
          }
        }
      }

      cardData.name = apiCard.name;
      cardData.set = apiCard.set_name;
      cardData.number = apiCard.collector_number;
      cardData.rarity = apiCard.rarity;

      // 価格（USD）
      if (apiCard.prices) {
        const priceUSD = parseFloat(apiCard.prices.usd || apiCard.prices.usd_foil);

        if (priceUSD) {
          const rate = getExchangeRate('USD', 'JPY');
          const priceJPY = Math.round(priceUSD * rate);

          cardData.price = priceJPY;
          cardData.priceUSD = priceUSD;
          cardData.exchangeRate = rate;

          console.log(`MTG価格: $${priceUSD} → ¥${priceJPY}`);
        }
      }
    }
  } catch (error) {
    console.error('Scryfall API エラー:', error);
  }
}

/**
 * 日本市場価格の推定
 */
function estimateJapanesePrice(cardData) {
  // レアリティベースの基本価格（日本円）
  const rarityPrices = {
    'UR': 15000,
    'HR': 8000,
    'SR': 5000,
    'SAR': 12000,
    'CSR': 10000,
    'CHR': 3000,
    'RRR': 2000,
    'RR': 1000,
    'R': 300,
    'U': 100,
    'C': 50
  };

  let basePrice = rarityPrices[cardData.rarity] || 500;

  // 人気ポケモンは価格上昇
  const popularPokemons = ['ピカチュウ', 'リザードン', 'イーブイ', 'ミュウ', 'レックウザ'];
  if (popularPokemons.some(name => cardData.name?.includes(name))) {
    basePrice *= 2;
  }

  // プロモカードは価値が異なる
  if (cardData.number?.includes('PROMO') || cardData.set?.includes('プロモ')) {
    basePrice *= 1.5;
  }

  return Math.round(basePrice);
}

/**
 * レアリティから価格を推定
 */
function estimatePriceByRarity(rarity) {
  const estimates = {
    'UR': 10000,
    'HR': 5000,
    'SR': 3000,
    'SAR': 8000,
    'SSR': 5000,
    'RRR': 1500,
    'RR': 800,
    'R': 200,
    'U': 80,
    'C': 30,
    'PROMO': 1000
  };

  return estimates[rarity] || 100;
}

/**
 * 価格履歴の生成（改良版）
 */
function generatePriceHistoryWithTrend(currentPrice) {
  if (!currentPrice || currentPrice === 0) {
    return {
      '12ヶ月前': 0,
      '9ヶ月前': 0,
      '6ヶ月前': 0,
      '3ヶ月前': 0,
      '現在': 0,
      'trend': 'stable'
    };
  }

  // トレンドをランダムに決定（実際はAPIや履歴データから）
  const trends = ['rising', 'falling', 'stable', 'volatile'];
  const trend = trends[Math.floor(Math.random() * trends.length)];

  let history = {};

  switch (trend) {
    case 'rising':
      // 上昇トレンド
      history = {
        '12ヶ月前': Math.round(currentPrice * 0.5),
        '9ヶ月前': Math.round(currentPrice * 0.65),
        '6ヶ月前': Math.round(currentPrice * 0.8),
        '3ヶ月前': Math.round(currentPrice * 0.9),
        '現在': currentPrice,
        'trend': '上昇'
      };
      break;

    case 'falling':
      // 下降トレンド
      history = {
        '12ヶ月前': Math.round(currentPrice * 1.8),
        '9ヶ月前': Math.round(currentPrice * 1.5),
        '6ヶ月前': Math.round(currentPrice * 1.3),
        '3ヶ月前': Math.round(currentPrice * 1.1),
        '現在': currentPrice,
        'trend': '下降'
      };
      break;

    case 'volatile':
      // 変動が激しい
      history = {
        '12ヶ月前': Math.round(currentPrice * 0.8),
        '9ヶ月前': Math.round(currentPrice * 1.2),
        '6ヶ月前': Math.round(currentPrice * 0.7),
        '3ヶ月前': Math.round(currentPrice * 1.1),
        '現在': currentPrice,
        'trend': '変動'
      };
      break;

    default:
      // 安定
      history = {
        '12ヶ月前': Math.round(currentPrice * 0.95),
        '9ヶ月前': Math.round(currentPrice * 0.97),
        '6ヶ月前': Math.round(currentPrice * 0.98),
        '3ヶ月前': Math.round(currentPrice * 0.99),
        '現在': currentPrice,
        'trend': '安定'
      };
  }

  return history;
}

/**
 * 価格取得のテスト
 */
function testPriceCalculation() {
  console.log('=== 価格取得テスト ===\n');

  // テストデータ
  const testCards = [
    {
      name: 'ピカチュウ',
      game: 'Pokemon',
      number: '025',
      set: 'Base Set',
      rarity: 'R'
    },
    {
      name: 'リザードンex',
      game: 'Pokemon',
      number: '054',
      set: 'Obsidian Flames',
      rarity: 'SR'
    }
  ];

  testCards.forEach(card => {
    console.log(`\nテスト: ${card.name}`);
    enrichCardDataWithPrice(card);

    console.log('結果:');
    console.log(`  価格（円）: ¥${card.price || '取得失敗'}`);
    console.log(`  価格（USD）: $${card.priceUSD || 'N/A'}`);
    console.log(`  為替レート: ${card.exchangeRate || 'N/A'}`);
    console.log(`  セット: ${card.setName || card.set}`);
    console.log(`  番号: ${card.number}`);
    console.log(`  レアリティ: ${card.rarity}`);
  });

  return testCards;
}