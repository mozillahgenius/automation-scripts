/**
 * 暗号資産期待値デイリーレポートシステム
 * Perplexity + Grok-4 統合版
 * CoinGecko価格取得、ステーキング対応、ブレークアウトポイント検出機能付き
 */

// =====================================
// 定数定義
// =====================================
const SHEETS = {
  ASSETS: 'Assets',
  HORIZON_CONFIG: 'HorizonConfig',
  PRICES: 'Prices',
  RESULTS: 'Results',
  BREAKOUTS: 'Breakouts',
  CONFIG: 'Config',
  LOG: 'Log'
};

const LLM_MODELS = {
  PERPLEXITY: 'perplexity',
  GROK: 'grok-4'
};

/**
 * Perplexityモデル名の取得
 * スクリプトプロパティ `PERPLEXITY_MODEL` があればそれを使用、なければ 'sonar-pro' を使用
 * 例: 'sonar-pro', 'sonar-medium', 'sonar-small', 'sonar-pro-reasoning'
 */
function getPerplexityModel() {
  // sonar-proのみを使用（固定）
  return 'sonar-pro';
}

const DEFAULT_HORIZONS = [14, 30, 90, 180, 365, 1095, 1825];

const PRICE_SOURCES = {
  // Perplexityのみ使用（実際には使用されないが互換性のため残す）
  'bitcoin': ['PERPLEXITY'],
  'ethereum': ['PERPLEXITY'],
  'ripple': ['PERPLEXITY'],
  'near': ['PERPLEXITY'],
  'cardano': ['PERPLEXITY'],
  'aave': ['PERPLEXITY'],
  'hedera-hashgraph': ['PERPLEXITY'],
  'the-graph': ['PERPLEXITY'],
  'algorand': ['PERPLEXITY'],
  'maker': ['PERPLEXITY'],
  'curve-dao-token': ['PERPLEXITY'],
  'cosmos': ['PERPLEXITY'],
  'polkadot': ['PERPLEXITY'],
  'polygon': ['PERPLEXITY'],
  'avalanche-2': ['PERPLEXITY'],
  'chainlink': ['PERPLEXITY'],
  'uniswap': ['PERPLEXITY'],
  'litecoin': ['PERPLEXITY'],
  'bitcoin-cash': ['PERPLEXITY'],
  'stellar': ['PERPLEXITY'],
  'synthetix-network-token': ['PERPLEXITY'],
  'solana': ['PERPLEXITY']
};

// =====================================
// メイン処理
// =====================================

/**
 * メイン実行関数
 */
function main() {
  try {
    console.log('暗号資産期待値レポート生成開始');
    
    // 初期化
    initializeSheets();
    
    // 設定読み込み
    const config = loadConfig();
    const assets = loadAssets();
    const horizons = loadHorizons();
    
    console.log(`読み込み完了 - 資産数: ${assets.length}, 評価期間: ${horizons.length}`);
    
    // 価格データ取得
    console.log('価格取得に渡すassets:', assets);
    if (!assets || assets.length === 0) {
      console.error('価格取得失敗: assetsが空または無効');
      return;
    }
    
    const prices = fetchPrices(assets);
    console.log('取得した価格データ:', prices);
    
    // ブレークアウトポイント検出
    const breakouts = detectBreakoutPoints(assets, prices);
    saveBreakouts(breakouts);
    
    // 各資産に対してマルチLLM分析
    const results = [];
    for (const asset of assets) {
      console.log(`分析中: ${asset.symbol}`);
      
      const basePrice = prices[asset.id];
      if (!basePrice || basePrice <= 0) {
        console.error(`価格データ取得失敗: ${asset.symbol} (${asset.id}) - 価格: ${basePrice}`);
        continue;
      }
      
      const assetResult = analyzeAssetWithMultiLLM(
        asset, 
        basePrice, 
        horizons,
        breakouts[asset.id] || []
      );
      
      if (assetResult) {
        results.push(assetResult);
      }
    }
    
    // 結果保存
    saveResults(results);
    
    // レポート生成・送信
    const report = generateHTMLReport(results, breakouts, config);
    sendEmailReport(report, config);
    
    // Notion更新
    try {
      updateNotionPage(results, breakouts, config);
      console.log('Notion更新完了');
    } catch (notionError) {
      console.error('Notion更新エラー:', notionError);
      // Notionエラーでもメイン処理は続行
    }
    
    // ログ記録
    logExecution('SUCCESS', `分析完了: ${results.length}資産`);
    
    console.log('暗号資産期待値レポート生成完了');
    
  } catch (error) {
    console.error('エラー:', error);
    logExecution('ERROR', error.toString());
    throw error;
  }
}

// =====================================
// 初期化処理
// =====================================

/**
 * スプレッドシート初期化
 */
function initializeSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Assets シート
  if (!ss.getSheetByName(SHEETS.ASSETS)) {
    const sheet = ss.insertSheet(SHEETS.ASSETS);
    sheet.getRange(1, 1, 1, 11).setValues([[
      'id', 'symbol', 'name', 'weight', 'staking_enabled', 
      'apy', 'compounding', 'fee_rate', 'haircut_risk', 'is_lst', 'notes'
    ]]);
    // 全銘柄の初期データ（デフォルトは無効化）
    const allAssets = [
      ['bitcoin', 'BTC', 'Bitcoin', 0, false, 0, 'annual', 0, 0, false, ''],
      ['ethereum', 'ETH', 'Ethereum', 0, false, 0, 'annual', 0, 0, false, ''],
      ['ripple', 'XRP', 'XRP', 0, false, 0, 'annual', 0, 0, false, ''],
      ['near', 'NEAR', 'NEAR Protocol', 0, false, 0, 'annual', 0, 0, false, ''],
      ['cardano', 'ADA', 'Cardano', 0, false, 0, 'annual', 0, 0, false, ''],
      ['aave', 'AAVE', 'Aave', 0, false, 0, 'annual', 0, 0, false, ''],
      ['hedera-hashgraph', 'HBAR', 'Hedera', 0, false, 0, 'annual', 0, 0, false, ''],
      ['the-graph', 'GRT', 'The Graph', 0, false, 0, 'annual', 0, 0, false, ''],
      ['algorand', 'ALGO', 'Algorand', 0, false, 0, 'annual', 0, 0, false, ''],
      ['maker', 'MKR', 'Maker', 0, false, 0, 'annual', 0, 0, false, ''],
      ['curve-dao-token', 'CRV', 'Curve', 0, false, 0, 'annual', 0, 0, false, ''],
      ['cosmos', 'ATOM', 'Cosmos', 0, false, 0, 'annual', 0, 0, false, ''],
      ['polkadot', 'DOT', 'Polkadot', 0, false, 0, 'annual', 0, 0, false, ''],
      ['polygon', 'POL', 'Polygon', 0, false, 0, 'annual', 0, 0, false, ''],
      ['avalanche-2', 'AVAX', 'Avalanche', 0, false, 0, 'annual', 0, 0, false, ''],
      ['chainlink', 'LINK', 'Chainlink', 0, false, 0, 'annual', 0, 0, false, ''],
      ['uniswap', 'UNI', 'Uniswap', 0, false, 0, 'annual', 0, 0, false, ''],
      ['litecoin', 'LTC', 'Litecoin', 0, false, 0, 'annual', 0, 0, false, ''],
      ['bitcoin-cash', 'BCH', 'Bitcoin Cash', 0, false, 0, 'annual', 0, 0, false, ''],
      ['stellar', 'XLM', 'Stellar', 0, false, 0, 'annual', 0, 0, false, ''],
      ['synthetix-network-token', 'SNX', 'Synthetix', 0, false, 0, 'annual', 0, 0, false, ''],
      ['solana', 'SOL', 'Solana', 0, false, 0, 'annual', 0, 0, false, '']
    ];
    
    sheet.getRange(2, 1, allAssets.length, 11).setValues(allAssets);
    
    // ヘルプコメントを追加
    sheet.getRange(1, 12).setValue('ヘルプ');
    sheet.getRange(2, 12).setValue('weight: 0=無効, 1=100%, 0.5=50%');
    sheet.getRange(3, 12).setValue('staking_enabled: true=有効, false=無効');
    sheet.getRange(4, 12).setValue('apy: 年率（例: 0.04=4%）');
    sheet.getRange(5, 12).setValue('compounding: annual/daily');
    sheet.getRange(6, 12).setValue('fee_rate: 手数料率（例: 0.1=10%）');
    sheet.getRange(7, 12).setValue('haircut_risk: リスク調整（例: 0.15=15%）');
    sheet.getRange(8, 12).setValue('is_lst: true=Liquid Staking Token');
  }
  
  // HorizonConfig シート
  if (!ss.getSheetByName(SHEETS.HORIZON_CONFIG)) {
    const sheet = ss.insertSheet(SHEETS.HORIZON_CONFIG);
    sheet.getRange(1, 1).setValue('horizon_days');
    const horizonData = DEFAULT_HORIZONS.map(h => [h]);
    sheet.getRange(2, 1, horizonData.length, 1).setValues(horizonData);
  }
  
  // Prices シート
  if (!ss.getSheetByName(SHEETS.PRICES)) {
    const sheet = ss.insertSheet(SHEETS.PRICES);
    sheet.getRange(1, 1, 1, 3).setValues([['id', 'asof_date', 'base_price_usd']]);
  }
  
  // Results シート
  if (!ss.getSheetByName(SHEETS.RESULTS)) {
    const sheet = ss.insertSheet(SHEETS.RESULTS);
    sheet.getRange(1, 1, 1, 10).setValues([[
      'timestamp', 'asset_id', 'horizon', 'price_ev_pct', 'staking_pct', 
      'total_ev_pct', 'confidence', 'models_used', 'scenarios', 'citations'
    ]]);
  }
  
  // Breakouts シート
  if (!ss.getSheetByName(SHEETS.BREAKOUTS)) {
    const sheet = ss.insertSheet(SHEETS.BREAKOUTS);
    sheet.getRange(1, 1, 1, 6).setValues([[
      'id', 'date', 'type', 'level', 'direction', 'rationale'
    ]]);
  }
  
  // Config シート
  if (!ss.getSheetByName(SHEETS.CONFIG)) {
    const sheet = ss.insertSheet(SHEETS.CONFIG);
    sheet.getRange(1, 1, 1, 2).setValues([['key', 'value']]);
    sheet.getRange(2, 1, 6, 2).setValues([
      ['report_hour', '9'],
      ['recipients', ''],
      ['subject_prefix', 'Crypto EV'],
      ['timezone', 'Asia/Kuala_Lumpur'],
      ['notion_database_id', ''],
      ['notion_page_id', '']
    ]);
  }
  
  // Log シート
  if (!ss.getSheetByName(SHEETS.LOG)) {
    const sheet = ss.insertSheet(SHEETS.LOG);
    sheet.getRange(1, 1, 1, 3).setValues([['timestamp', 'status', 'message']]);
  }
}

// =====================================
// データ読み込み
// =====================================

/**
 * 設定読み込み
 */
function loadConfig() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.CONFIG);
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
  
  const config = {};
  data.forEach(row => {
    if (row[0]) config[row[0]] = row[1];
  });
  
  return config;
}

/**
 * 資産情報読み込み
 */
function loadAssets() {
  try {
    console.log('loadAssets開始');
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.ASSETS);
    if (!sheet) {
      console.error('Assetsシートが見つかりません');
      // デフォルトの資産リストを返す
      return getDefaultAssets();
    }
    
    const lastRow = sheet.getLastRow();
    console.log(`Assetsシート最終行: ${lastRow}`);
    
    if (lastRow <= 1) {
      console.warn('Assetsシートにデータがありません。デフォルト資産を使用します。');
      return getDefaultAssets();
    }
    
    const data = sheet.getRange(2, 1, lastRow - 1, 11).getValues();
    console.log(`取得したデータ行数: ${data.length}`);
    
    const assets = data.filter(row => {
      console.log('資産行データ:', row);
      // weightが未設定の場合もデフォルト値1.0で処理
      return row[0]; // IDがあればOK
    }).map(row => ({
      id: row[0],
      symbol: row[1] || row[0].toUpperCase(),
      name: row[2] || row[0],
      weight: parseFloat(row[3]) || 1.0,
      staking_enabled: row[4] === true || row[4] === 'TRUE',
      apy: parseFloat(row[5]) || 0,
      compounding: row[6] || 'annual',
      fee_rate: parseFloat(row[7]) || 0,
      haircut_risk: parseFloat(row[8]) || 0,
      is_lst: row[9] === true || row[9] === 'TRUE',
      notes: row[10] || ''
    }));
    
    console.log(`読み込んだ有効資産数: ${assets.length}`);
    assets.forEach(asset => {
      console.log(`有効資産: ${asset.id} (${asset.symbol}) - weight: ${asset.weight}`);
    });
    
    // 資産が空の場合、デフォルト資産を使用
    if (assets.length === 0) {
      console.warn('有効な資産データがありません。デフォルト資産を適用します。');
      return getDefaultAssets();
    }
    
    return assets;
    
  } catch (error) {
    console.error('loadAssetsでエラー発生:', error);
    console.error('エラースタック:', error.stack);
    // エラー時はデフォルト資産を返す
    return getDefaultAssets();
  }
}

/**
 * デフォルトの資産リストを返す
 */
function getDefaultAssets() {
  console.log('デフォルト資産リストを生成');
  return [
    { id: 'bitcoin', symbol: 'BTC', name: 'Bitcoin', weight: 1.0, staking_enabled: false, apy: 0, compounding: 'annual', fee_rate: 0, haircut_risk: 0, is_lst: false, notes: '' },
    { id: 'ethereum', symbol: 'ETH', name: 'Ethereum', weight: 1.0, staking_enabled: true, apy: 3.5, compounding: 'annual', fee_rate: 0, haircut_risk: 0, is_lst: false, notes: '' },
    { id: 'solana', symbol: 'SOL', name: 'Solana', weight: 1.0, staking_enabled: true, apy: 5.0, compounding: 'annual', fee_rate: 0, haircut_risk: 0, is_lst: false, notes: '' }
  ];
}

/**
 * 評価期間読み込み
 */
function loadHorizons() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.HORIZON_CONFIG);
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
  return data.filter(row => row[0]).map(row => parseInt(row[0]));
}

// =====================================
// 価格データ取得
// =====================================

/**
 * 価格データ取得
 */
function fetchPrices(assets) {
  const prices = {};
  const today = new Date().toISOString().split('T')[0];
  
  // assetsパラメータの検証
  if (!assets) {
    console.warn('fetchPrices: assetsパラメータが未定義です');
    assets = [];
  }
  
  console.log(`価格取得開始 - 日付: ${today}, 対象資産数: ${assets.length}`);
  
  // Pricesシートから読み込み（今日のデータのみ）
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.PRICES);
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    const priceData = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
    console.log(`Pricesシートからのデータ行数: ${priceData.length}`);
    
    priceData.forEach(row => {
      if (row[0] && row[2] && row[1]) {
        const assetId = row[0];
        const dateStr = row[1] instanceof Date ? row[1].toISOString().split('T')[0] : row[1];
        const price = parseFloat(row[2]);
        
        // 今日のデータのみ使用
        if (dateStr === today) {
          prices[assetId] = price;
          console.log(`Pricesシートから取得: ${assetId} = $${price} (${dateStr})`);
        } else {
          console.log(`古いデータをスキップ: ${assetId} = $${price} (${dateStr})`);
        }
      }
    });
  }
  
  // 資産リストが空でないことを確認
  if (!assets || !Array.isArray(assets) || assets.length === 0) {
    console.warn('資産リストが空です。');
    return prices;
  }
  
  // 不足分の資産IDをリスト化
  console.log('assetsの内容:', assets);
  console.log('現在のpricesの内容:', prices);
  
  const missingAssetIds = assets
    .filter(asset => {
      console.log(`資産チェック: ${asset.id} - 価格あり: ${!!prices[asset.id]}`);
      return !prices[asset.id];
    })
    .map(asset => asset.id);
  
  console.log('missingAssetIds:', missingAssetIds);
  
  if (missingAssetIds.length > 0) {
    console.log(`価格取得が必要な資産: ${missingAssetIds.join(', ')}`);
    
    // Perplexityで一括価格取得
    try {
      console.log('Perplexity一括価格取得を開始...');
      
      // 再度引数の検証
      if (!missingAssetIds || !Array.isArray(missingAssetIds) || missingAssetIds.length === 0) {
        console.warn('missingAssetIdsが無効なため一括取得をスキップ');
        return prices;
      }
      
      const perplexityPrices = fetchAllPricesWithPerplexity(missingAssetIds);
      console.log('Perplexity一括取得結果:', perplexityPrices);
      
      // 取得できた価格のみ追加
      Object.keys(perplexityPrices).forEach(assetId => {
        const price = perplexityPrices[assetId];
        if (price && price > 0) {
          prices[assetId] = price;
          console.log(`Perplexity取得成功: ${assetId} = $${price}`);
        } else {
          console.warn(`無効な価格データ: ${assetId} = ${price}`);
        }
      });
      
      // 取得できなかった資産を個別に再試行
      const stillMissingIds = missingAssetIds.filter(id => !prices[id]);
      if (stillMissingIds.length > 0) {
        console.log(`取得できなかった資産を個別に再試行: ${stillMissingIds.join(', ')}`);
        
        stillMissingIds.forEach(assetId => {
          try {
            Utilities.sleep(1000); // レート制限対策
            console.log(`個別取得試行: ${assetId}`);
            const price = fetchPriceWithPerplexity(assetId);
            if (price && price > 0) {
              prices[assetId] = price;
              console.log(`個別取得成功: ${assetId} = $${price}`);
            } else {
              console.warn(`個別取得失敗: ${assetId} = ${price}`);
            }
          } catch (e) {
            console.error(`価格取得失敗 (${assetId}):`, e);
          }
        });
      }
      
    } catch (e) {
      console.error('Perplexity価格一括取得エラー:', e);
      
      // フォールバック: 個別取得
      console.log('個別価格取得にフォールバック...');
      missingAssetIds.forEach(assetId => {
        if (!prices[assetId]) {
          try {
            Utilities.sleep(1000); // レート制限対策
            console.log(`個別取得試行: ${assetId}`);
            const price = fetchPriceWithPerplexity(assetId);
            if (price && price > 0) {
              prices[assetId] = price;
              console.log(`個別取得成功: ${assetId} = $${price}`);
            } else {
              console.warn(`個別取得失敗: ${assetId} = ${price}`);
            }
          } catch (e) {
            console.error(`価格取得失敗 (${assetId}):`, e);
          }
        }
      });
    }
  }
  
  // 最終価格数をログ出力
  const finalPriceCount = Object.keys(prices).length;
  console.log(`最終価格データ数: ${finalPriceCount}`);
  Object.keys(prices).forEach(id => {
    console.log(`最終価格: ${id} = $${prices[id]}`);
  });
  
  // 取得した価格をPricesシートに保存（同日重複は回避）
  try {
    savePricesForDate(prices, today);
  } catch (e) {
    console.error('Pricesシート保存エラー:', e);
  }
  
  return prices;
}

/*
 * デフォルト価格取得機能を無効化
 * デフォルト価格は使用しない方針のためコメントアウト
 */
/*
function getDefaultPrice(assetId) {
  const defaultPrices = {
    'bitcoin': 100000,
    'ethereum': 3500,
    'cardano': 1.0,
    'solana': 200,
    'polygon': 1.0,
    'chainlink': 25,
    'polkadot': 8,
    'avalanche-2': 40,
    'cosmos': 12,
    'near': 6
  };
  
  return defaultPrices[assetId] || 1000; // 不明な資産には$1000をデフォルト設定
}
*/

/**
 * Pricesシートへ当日価格を保存（同日・同資産の重複はスキップ）
 * pricesByIdが空の場合は、現在の価格を取得して保存
 */
function savePricesForDate(pricesById, asOfDate) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.PRICES);
  if (!sheet) return;
  
  // pricesByIdが空または未定義の場合、現在の価格を取得
  if (!pricesById || Object.keys(pricesById).length === 0) {
    console.log('価格データが空のため、現在の価格を取得します');
    
    // 有効な資産リストを取得
    const assets = loadAssets();
    if (!assets || assets.length === 0) {
      console.log('有効な資産がないため、価格取得をスキップします');
      return;
    }
    
    // 不足分の資産IDをリスト化
    const assetIds = assets.map(asset => asset.id);
    console.log('現在価格を取得する資産:', assetIds);
    
    try {
      // Perplexityで一括価格取得
      const currentPrices = fetchAllPricesWithPerplexity(assetIds);
      console.log('取得した現在価格:', currentPrices);
      
      // 取得できなかった資産を個別に再試行
      const stillMissingIds = assetIds.filter(id => !currentPrices[id]);
      if (stillMissingIds.length > 0) {
        console.log(`取得できなかった資産を個別に再試行: ${stillMissingIds.join(', ')}`);
        
        stillMissingIds.forEach(assetId => {
          try {
            Utilities.sleep(1000); // レート制限対策
            console.log(`個別取得試行: ${assetId}`);
            const price = fetchPriceWithPerplexity(assetId);
            if (price && price > 0) {
              currentPrices[assetId] = price;
              console.log(`個別取得成功: ${assetId} = $${price}`);
            } else {
              console.warn(`個別取得失敗: ${assetId} = ${price}`);
            }
          } catch (e) {
            console.error(`価格取得失敗 (${assetId}):`, e);
          }
        });
      }
      
      pricesById = currentPrices;
    } catch (e) {
      console.error('現在価格取得エラー:', e);
      return;
    }
  }
  
  // 有効な価格データのみフィルタリング
  const validPrices = {};
  Object.keys(pricesById).forEach(id => {
    const price = pricesById[id];
    if (price && price > 0) {
      validPrices[id] = price;
    }
  });
  
  if (Object.keys(validPrices).length === 0) {
    console.log('保存する有効な価格データがありません');
    return;
  }
  
  const lastRow = sheet.getLastRow();
  const existing = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, 3).getValues() : [];
  const existingSet = new Set();
  existing.forEach(row => {
    const id = row[0];
    const date = row[1];
    if (id && date) existingSet.add(`${id}__${date}`);
  });
  
  const rowsToAppend = [];
  Object.keys(validPrices).forEach(id => {
    const key = `${id}__${asOfDate}`;
    if (!existingSet.has(key)) {
      rowsToAppend.push([id, asOfDate, validPrices[id]]);
    }
  });
  
  if (rowsToAppend.length > 0) {
    sheet.getRange(lastRow + 1, 1, rowsToAppend.length, 3).setValues(rowsToAppend);
    console.log(`価格データ保存完了: ${rowsToAppend.length}件`);
  }
}

// =====================================
// マルチLLM分析
// =====================================

/**
 * マルチLLMによる資産分析（簡素化されたワークフロー）
 * 1. Perplexity: 最新情報検索・事実調査
 * 2. Grok-4: Perplexityの結果を元に分析
 */
function analyzeAssetWithMultiLLM(asset, basePrice, horizons, recentBreakouts) {
  const modelResults = [];
  let perplexityResult = null;
  let grokResult = null;
  
  // ステップ1: Perplexityで最新情報検索・事実調査
  try {
    console.log(`  - perplexityで最新情報調査中...`);
    perplexityResult = analyzeWithPerplexity(asset, basePrice, horizons, recentBreakouts);
    if (perplexityResult && validateLLMResponse(perplexityResult)) {
      modelResults.push({
        name: LLM_MODELS.PERPLEXITY,
        weight: 1.0,
        result: perplexityResult
      });
    }
  } catch (e) {
    console.error(`  - perplexityエラー:`, e);
  }
  
  // ステップ2: Grok-4で分析（設定で有効な場合のみ）
  // 現在Grok-4に問題があるため、デフォルトで無効化
  const useGrok = PropertiesService.getScriptProperties().getProperty('USE_GROK') === 'true';
  
  if (useGrok) {
    try {
      console.log(`  - grok-4で分析中...`);
      grokResult = analyzeWithGrok(asset, basePrice, horizons, recentBreakouts, perplexityResult);
      if (grokResult && validateLLMResponse(grokResult)) {
        modelResults.push({
          name: LLM_MODELS.GROK,
          weight: 1.2, // 分析の重要性を考慮して重み増加
          result: grokResult
        });
      } else if (grokResult === null) {
        console.log(`  - grok-4が利用不可のためスキップ`);
      }
    } catch (e) {
      console.error(`  - grok-4エラー:`, e);
      console.log(`  - grok-4の結果を使用せずに続行`);
    }
  } else {
    console.log(`  - grok-4は無効化されています`);
  }
  
  // 最低1モデル必要（簡素化のため条件を緩和）
  if (modelResults.length < 1) {
    console.log(`  分析失敗: 有効モデル数不足 (${modelResults.length})`);
    return null;
  }
  
  // パネル平均で集計
  const aggregated = aggregateModelResults(modelResults, asset, basePrice, horizons);
  
  // ステーキング計算
  if (asset.staking_enabled && !asset.is_lst) {
    addStakingReturns(aggregated, asset, horizons);
  }
  
  return aggregated;
}

/**
 * Perplexityによる分析（事実調査重視 + 価格取得も可能）
 */
function analyzeWithPerplexity(asset, basePrice, horizons, breakouts, includePriceCheck = false) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('PERPLEXITY_API_KEY');
  if (!apiKey) throw new Error('Perplexity APIキー未設定');
  
  let prompt = buildAnalysisPrompt(asset, basePrice, horizons, breakouts, 'perplexity');
  
  // 価格確認を含める場合
  if (includePriceCheck && (!basePrice || basePrice <= 0)) {
    const currentTime = new Date().toISOString();
    const currentTimeJST = new Date().toLocaleString('ja-JP', {timeZone: 'Asia/Tokyo'});
    prompt = `🚨 URGENT: まず${asset.name} (${asset.symbol})の現在価格（${currentTime} / JST: ${currentTimeJST}時点）をリアルタイムで調査し、その最新価格を基に分析してください。古いデータではなく、現在の最新価格を使用してください。\n` + prompt;
  }
  
  const url = 'https://api.perplexity.ai/chat/completions';
  const payload = {
    model: getPerplexityModel(),
    messages: [
      {
        role: 'system',
        content: 'あなたは暗号資産の調査専門家です。最新の事実情報を徹底的に調査し、客観的なデータに基づいて期待値を分析してください。'
      },
      {
        role: 'user',
        content: prompt
      }
    ],
    temperature: 0.2,
    max_tokens: 2000
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
  
  const response = UrlFetchApp.fetch(url, options);
  const data = JSON.parse(response.getContentText());
  
  if (data.choices && data.choices[0]) {
    const content = data.choices[0].message.content;
    const result = parseJSONFromLLM(content);
    
    // 価格情報が含まれている場合は抽出
    if (includePriceCheck && result.current_price) {
      console.log(`Perplexity分析から価格取得: ${asset.symbol} = $${result.current_price}`);
    }
    
    return result;
  }
  
  throw new Error('Perplexity応答エラー');
}



/**
 * Grok-4による分析（Perplexityの事実調査を元に総合分析）
 */
function analyzeWithGrok(asset, basePrice, horizons, breakouts, perplexityResult = null) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GROK_API_KEY');
  if (!apiKey) throw new Error('Grok APIキー未設定');
  
  const prompt = buildAnalysisPrompt(asset, basePrice, horizons, breakouts, 'grok', perplexityResult);
  
  const url = 'https://api.x.ai/v1/chat/completions';
  const payload = {
    model: 'grok-4',
    messages: [
      {
        role: 'system',
        content: 'あなたは暗号資産の総合分析専門家です。出力は有効なJSONのみで返してください。'
      },
      {
        role: 'user',
        content: prompt
      }
    ],
    temperature: 0.2,
    max_tokens: 2000
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
  
  const response = UrlFetchApp.fetch(url, options);
  const responseCode = response.getResponseCode();
  const responseText = response.getContentText();
  
  if (responseCode === 502 || responseCode === 503 || responseCode === 504) {
    console.error(`Grok-4 サーバーエラー (${responseCode}): ${responseText.substring(0, 200)}`);
    console.log('Grok-4が利用できないため、スキップします');
    return null; // nullを返してスキップ
  }
  
  if (responseCode !== 200) {
    console.error(`Grok-4 HTTP ${responseCode}: ${responseText.substring(0, 200)}`);
    console.log('Grok-4エラーのため、スキップします');
    return null; // エラー時もnullを返す
  }
  
  const data = JSON.parse(responseText);
  
  if (data.error) {
    console.error('Grok-4 APIエラー:', data.error);
    throw new Error(`Grok-4 API: ${data.error.message || data.error.type || 'Unknown error'}`);
  }
  
  if (data.choices && data.choices[0]) {
    const content = data.choices[0].message.content || '';
    console.log('Grok-4応答長さ:', content.length);
    
    // 空の応答をチェック
    if (!content || content.trim().length === 0) {
      console.error('Grok-4: 空の応答を受信しました');
      return null;
    }
    
    try {
      return parseJSONFromLLM(content);
    } catch (e) {
      console.error('Grok-4 JSONパースエラー:', e.message);
      console.log('応答内容（先頭500文字）:', content.substring(0, 500));
      
      // より簡潔なプロンプトで再試行
      const simplePrompt = `
${asset.symbol}の価格期待値分析（現在価格$${basePrice}）
以下のJSONのみを返してください：
{"asset":"${asset.symbol}","horizons":[${horizons.join(',')}],"scenarios":[{"name":"bear","prob":0.3,"return_pct_by_horizon":{${horizons.map(h => `"${h}":-10`).join(',')}},"rationale":"弱気"},{"name":"base","prob":0.5,"return_pct_by_horizon":{${horizons.map(h => `"${h}":5`).join(',')}},"rationale":"中立"},{"name":"bull","prob":0.2,"return_pct_by_horizon":{${horizons.map(h => `"${h}":20`).join(',')}},"rationale":"強気"}],"ev_return_pct_by_horizon":{${horizons.map(h => `"${h}":2`).join(',')}},"confidence":0.7,"citations":[],"breakout_points":[]}`;
      
      const simplePayload = {
        model: 'grok-4',
        messages: [
          { role: 'user', content: simplePrompt }
        ],
        temperature: 0.0,
        max_tokens: 1000
      };
      
      try {
        const simpleResp = UrlFetchApp.fetch(url, {
          method: 'post',
          headers: {
            'Authorization': `Bearer ${apiKey}`,
            'Content-Type': 'application/json'
          },
          payload: JSON.stringify(simplePayload),
          muteHttpExceptions: true
        });
        
        const respCode = simpleResp.getResponseCode();
        if (respCode === 502 || respCode === 503 || respCode === 504) {
          console.log('Grok-4 再試行もサーバーエラー');
          return null;
        }
        
        if (respCode === 200) {
          const simpleData = JSON.parse(simpleResp.getContentText());
          const simpleContent = simpleData.choices && simpleData.choices[0] && simpleData.choices[0].message 
            ? simpleData.choices[0].message.content : '';
          if (simpleContent && simpleContent.trim().length > 0) {
            return parseJSONFromLLM(simpleContent);
          }
        }
      } catch (retryError) {
        console.error('Grok-4 再試行エラー:', retryError.message);
        return null;
      }
    }
  }
  
  console.log('Grok-4: 有効な応答が得られませんでした');
  return null;
}



/**
 * 分析プロンプト構築（簡素化版）
 */
function buildAnalysisPrompt(asset, basePrice, horizons, breakouts, modelType, ...previousResults) {
  const breakoutInfo = breakouts.length > 0 
    ? `\n最近のブレークアウト:\n${breakouts.slice(0, 3).map(b => 
        `- ${b.date}: ${b.type} ${b.direction} at $${b.level} (${b.rationale})`
      ).join('\n')}`
    : '';
  
  // 前のステップの結果をコンテキストとして追加
  let contextInfo = '';
  const [perplexityResult] = previousResults;
  
  if (modelType === 'grok' && perplexityResult) {
    contextInfo = `\n\n【Perplexityによる事実調査結果】\n${JSON.stringify(perplexityResult, null, 2)}`;
    contextInfo += `\n\n上記の事実情報を基に、総合的な分析と精密な期待値計算を行ってください。`;
  }
  
  const roleMap = {
    perplexity: '最新情報の徹底調査と事実収集',
    grok: 'Perplexityの事実情報を基にした総合分析と期待値計算'
  };
  
  return `
暗号資産 ${asset.symbol} (${asset.name}) の期待値分析を行ってください。
役割: ${roleMap[modelType]}

現在価格: $${basePrice}
評価期間（日数）: ${horizons.join(', ')}
${breakoutInfo}${contextInfo}

重要: 各期間で異なる価格変動率を設定してください。長期ほど変動幅が大きくなるはずです。
以下のJSONフォーマットで、実際の市場分析に基づいた予測を返してください：
{
  "asset": "${asset.symbol}",
  "horizons": [${horizons.join(',')}],
  "scenarios": [
    {
      "name": "bear",
      "prob": 0.3,
      "return_pct_by_horizon": {${horizons.map((h, i) => `"${h}":${-5 * (i + 1)}`).join(',')}},
      "rationale": "bearish scenario"
    },
    {
      "name": "base",
      "prob": 0.5,
      "return_pct_by_horizon": {${horizons.map((h, i) => `"${h}":${3 * (i + 1)}`).join(',')}},
      "rationale": "base scenario"
    },
    {
      "name": "bull",
      "prob": 0.2,
      "return_pct_by_horizon": {${horizons.map((h, i) => `"${h}":${10 * (i + 1)}`).join(',')}},
      "rationale": "bullish scenario"
    }
  ],
  "ev_return_pct_by_horizon": {${horizons.map((h, i) => `"${h}":${2 * (i + 1)}`).join(',')}},
  "confidence": 0.7,
  "citations": [],
  "breakout_points": []
}

制約：
- シナリオ確率の合計は1.0
- return_pct_by_horizonは全期間分必須
- citationsは最大3件
`;
}

/**
 * LLM応答からJSON解析
 */
function parseJSONFromLLM(content) {
  // 空のコンテンツチェック
  if (!content || content.trim().length === 0) {
    throw new Error('空の応答');
  }
  
  // JSONブロックを抽出
  const jsonMatch = content.match(/\{[\s\S]*\}/);
  if (!jsonMatch) {
    throw new Error('JSON形式が見つかりません');
  }
  
  try {
    return JSON.parse(sanitizeJson(jsonMatch[0]));
  } catch (e) {
    // より詳細なエラー情報を記録
    console.error('JSON解析エラー詳細:', e.message);
    console.error('解析対象（先頭200文字）:', jsonMatch[0].substring(0, 200));
    
    // リトライ: コードブロック除去
    const cleaned = content
      .replace(/```json/g, '')
      .replace(/```/g, '')
      .trim();
    
    const retryMatch = cleaned.match(/\{[\s\S]*\}/);
    if (retryMatch) {
      try {
        return JSON.parse(sanitizeJson(retryMatch[0]));
      } catch (retryError) {
        console.error('再試行も失敗:', retryError.message);
        throw new Error(`JSON解析失敗: ${e.message}`);
      }
    }
    
    throw new Error(`JSON解析失敗: ${e.message}`);
  }
}

// LLM出力JSONのサニタイズ（末尾カンマ・コメント・不正な改行の除去）
function sanitizeJson(text) {
  let t = text;
  
  // コメント除去（JSONコメントはサポートされていないため）
  t = t.replace(/\/\*[\s\S]*?\*\//g, '').replace(/\/\/.*$/gm, '');
  
  // 末尾カンマの除去 {"a":1,} → {"a":1}
  t = t.replace(/,\s*([}\]])/g, '$1');
  
  // 非表示の制御文字除去（改行とタブは保持）
  t = t.replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F\u007F]/g, '');
  
  // 文字列内でない改行を除去
  let inString = false;
  let escaped = false;
  let result = '';
  
  for (let i = 0; i < t.length; i++) {
    const char = t[i];
    const prevChar = i > 0 ? t[i-1] : '';
    
    if (char === '"' && !escaped) {
      inString = !inString;
    }
    
    escaped = (char === '\\' && !escaped);
    
    if ((char === '\n' || char === '\r') && !inString) {
      result += ' ';
    } else {
      result += char;
    }
  }
  
  // 連続スペースを単一スペースに
  result = result.replace(/\s+/g, ' ');
  
  return result.trim();
}

/**
 * LLM応答の検証
 */
function validateLLMResponse(response) {
  // 必須フィールドチェック
  if (!response.asset || !response.horizons || !response.scenarios) {
    return false;
  }
  
  // シナリオ確率の合計チェック
  const totalProb = response.scenarios.reduce((sum, s) => sum + (s.prob || 0), 0);
  if (Math.abs(totalProb - 1.0) > 0.01) {
    return false;
  }
  
  // 各シナリオの期間チェック
  for (const scenario of response.scenarios) {
    if (!scenario.return_pct_by_horizon) {
      return false;
    }
  }
  
  return true;
}

/**
 * モデル結果の集計
 */
function aggregateModelResults(modelResults, asset, basePrice, horizons) {
  const aggregated = {
    asset_id: asset.id,
    asset_symbol: asset.symbol,
    base_price: basePrice,
    horizons: horizons,
    models_used: modelResults.map(m => m.name),
    scenarios: {},
    ev_by_horizon: {},
    confidence: 0,
    citations: []
  };
  
  // 総重み
  const totalWeight = modelResults.reduce((sum, m) => sum + m.weight, 0);
  
  // 各期間のEV集計
  horizons.forEach(horizon => {
    let weightedSum = 0;
    let weightCount = 0;
    
    modelResults.forEach(model => {
      const ev = model.result.ev_return_pct_by_horizon[horizon];
      if (ev !== undefined) {
        weightedSum += ev * model.weight;
        weightCount += model.weight;
      }
    });
    
    aggregated.ev_by_horizon[horizon] = weightCount > 0 ? weightedSum / weightCount : 0;
  });
  
  // シナリオ集計
  const scenarioTypes = ['bear', 'base', 'bull'];
  scenarioTypes.forEach(scenarioName => {
    const scenario = {
      name: scenarioName,
      prob: 0,
      return_by_horizon: {},
      rationales: []
    };
    
    let probSum = 0;
    let probWeight = 0;
    
    modelResults.forEach(model => {
      const modelScenario = model.result.scenarios.find(s => s.name === scenarioName);
      if (modelScenario) {
        probSum += modelScenario.prob * model.weight;
        probWeight += model.weight;
        
        if (modelScenario.rationale) {
          scenario.rationales.push(modelScenario.rationale);
        }
        
        // リターン集計
        horizons.forEach(horizon => {
          if (!scenario.return_by_horizon[horizon]) {
            scenario.return_by_horizon[horizon] = 0;
          }
          const ret = modelScenario.return_pct_by_horizon[horizon];
          if (ret !== undefined) {
            scenario.return_by_horizon[horizon] += ret * model.weight / totalWeight;
          }
        });
      }
    });
    
    scenario.prob = probWeight > 0 ? probSum / probWeight : 0;
    aggregated.scenarios[scenarioName] = scenario;
  });
  
  // 信頼度平均
  const confidenceSum = modelResults.reduce((sum, m) => 
    sum + (m.result.confidence || 0.5) * m.weight, 0
  );
  aggregated.confidence = confidenceSum / totalWeight;
  
  // 引用集計（重複除去）
  const citationSet = new Set();
  modelResults.forEach(model => {
    if (model.result.citations) {
      model.result.citations.forEach(c => citationSet.add(c));
    }
  });
  aggregated.citations = Array.from(citationSet).slice(0, 3);
  
  return aggregated;
}

/**
 * ステーキングリターン追加
 */
function addStakingReturns(result, asset, horizons) {
  const effectiveAPY = asset.apy * (1 - asset.fee_rate) * (1 - asset.haircut_risk);
  
  result.staking_by_horizon = {};
  result.total_ev_by_horizon = {};
  
  horizons.forEach(horizon => {
    const years = horizon / 365;
    let stakingReturn = 0;
    
    if (asset.compounding === 'daily') {
      stakingReturn = Math.pow(1 + effectiveAPY / 365, horizon) - 1;
    } else {
      // annual compounding
      stakingReturn = Math.pow(1 + effectiveAPY, years) - 1;
    }
    
    result.staking_by_horizon[horizon] = stakingReturn * 100;
    
    // 総合EV = (1 + 価格EV) × (1 + ステーキング) - 1
    const priceEV = result.ev_by_horizon[horizon] / 100;
    const totalEV = (1 + priceEV) * (1 + stakingReturn) - 1;
    result.total_ev_by_horizon[horizon] = totalEV * 100;
  });
}

// =====================================
// ブレークアウトポイント検出
// =====================================

/**
 * ブレークアウトポイント検出
 */
function detectBreakoutPoints(assets, prices) {
  const breakouts = {};
  const today = new Date().toISOString().split('T')[0];
  
  // 現在の価格に基づいた動的なブレークアウトポイント検出
  assets.forEach(asset => {
    breakouts[asset.id] = [];
    
    const price = prices[asset.id];
    if (!price) return;
    
    // 価格に基づいた重要なレベルを計算
    const roundedPrice = Math.round(price / 1000) * 1000; // 1000単位で丸める
    const resistanceLevel = roundedPrice + (roundedPrice * 0.1); // 10%上
    const supportLevel = roundedPrice - (roundedPrice * 0.1); // 10%下
    
    // 主要な暗号通貨のブレークアウトポイント
    if (asset.id === 'bitcoin') {
      // BTCの上昇ブレークアウト
      const btcUpBreakout1 = Math.round(price * 1.05); // 5%上
      const btcUpBreakout2 = Math.round(price * 1.15); // 15%上
      const btcUpBreakout3 = Math.round(price * 1.30); // 30%上
      
      breakouts[asset.id].push({
        date: today,
        type: 'technical',
        level: btcUpBreakout1,
        direction: 'up',
        rationale: `短期レジスタンス$${btcUpBreakout1.toLocaleString()}`
      });
      
      breakouts[asset.id].push({
        date: today,
        type: 'technical',
        level: btcUpBreakout2,
        direction: 'up',
        rationale: `中期レジスタンス$${btcUpBreakout2.toLocaleString()}`
      });
      
      breakouts[asset.id].push({
        date: today,
        type: 'psychological',
        level: btcUpBreakout3,
        direction: 'up',
        rationale: `長期目標$${btcUpBreakout3.toLocaleString()}`
      });
      
      // BTCの下落ブレークアウト
      const btcDownBreakout1 = Math.round(price * 0.95); // 5%下
      const btcDownBreakout2 = Math.round(price * 0.85); // 15%下
      const btcDownBreakout3 = Math.round(price * 0.70); // 30%下
      
      breakouts[asset.id].push({
        date: today,
        type: 'technical',
        level: btcDownBreakout1,
        direction: 'down',
        rationale: `短期サポート$${btcDownBreakout1.toLocaleString()}`
      });
      
      breakouts[asset.id].push({
        date: today,
        type: 'technical',
        level: btcDownBreakout2,
        direction: 'down',
        rationale: `中期サポート$${btcDownBreakout2.toLocaleString()}`
      });
      
      breakouts[asset.id].push({
        date: today,
        type: 'psychological',
        level: btcDownBreakout3,
        direction: 'down',
        rationale: `危険水準$${btcDownBreakout3.toLocaleString()}`
      });
    }
    
    else if (asset.id === 'ethereum') {
      // ETHの上昇ブレークアウト
      const ethUpBreakout1 = Math.round(price * 1.08); // 8%上
      const ethUpBreakout2 = Math.round(price * 1.20); // 20%上
      const ethUpBreakout3 = Math.round(price * 1.40); // 40%上
      
      breakouts[asset.id].push({
        date: today,
        type: 'technical',
        level: ethUpBreakout1,
        direction: 'up',
        rationale: `短期レジスタンス$${ethUpBreakout1.toLocaleString()}`
      });
      
      breakouts[asset.id].push({
        date: today,
        type: 'technical',
        level: ethUpBreakout2,
        direction: 'up',
        rationale: `中期レジスタンス$${ethUpBreakout2.toLocaleString()}`
      });
      
      breakouts[asset.id].push({
        date: today,
        type: 'psychological',
        level: ethUpBreakout3,
        direction: 'up',
        rationale: `長期目標$${ethUpBreakout3.toLocaleString()}`
      });
      
      // ETHの下落ブレークアウト
      const ethDownBreakout1 = Math.round(price * 0.92); // 8%下
      const ethDownBreakout2 = Math.round(price * 0.80); // 20%下
      const ethDownBreakout3 = Math.round(price * 0.65); // 35%下
      
      breakouts[asset.id].push({
        date: today,
        type: 'technical',
        level: ethDownBreakout1,
        direction: 'down',
        rationale: `短期サポート$${ethDownBreakout1.toLocaleString()}`
      });
      
      breakouts[asset.id].push({
        date: today,
        type: 'technical',
        level: ethDownBreakout2,
        direction: 'down',
        rationale: `中期サポート$${ethDownBreakout2.toLocaleString()}`
      });
      
      breakouts[asset.id].push({
        date: today,
        type: 'psychological',
        level: ethDownBreakout3,
        direction: 'down',
        rationale: `危険水準$${ethDownBreakout3.toLocaleString()}`
      });
    }
    
    // その他の通貨には汎用的なブレークアウトポイントを設定
    else if (price > 0) {
      // 上昇ブレークアウト（3段階）
      const upBreakout1 = Math.round(price * 1.10 * 100) / 100; // 10%上
      const upBreakout2 = Math.round(price * 1.25 * 100) / 100; // 25%上
      const upBreakout3 = Math.round(price * 1.50 * 100) / 100; // 50%上
      
      breakouts[asset.id].push({
        date: today,
        type: 'technical',
        level: upBreakout1,
        direction: 'up',
        rationale: `短期レジスタンス$${upBreakout1.toLocaleString()}`
      });
      
      breakouts[asset.id].push({
        date: today,
        type: 'technical',
        level: upBreakout2,
        direction: 'up',
        rationale: `中期レジスタンス$${upBreakout2.toLocaleString()}`
      });
      
      breakouts[asset.id].push({
        date: today,
        type: 'psychological',
        level: upBreakout3,
        direction: 'up',
        rationale: `長期目標$${upBreakout3.toLocaleString()}`
      });
      
      // 下落ブレークアウト（3段階）
      const downBreakout1 = Math.round(price * 0.90 * 100) / 100; // 10%下
      const downBreakout2 = Math.round(price * 0.75 * 100) / 100; // 25%下
      const downBreakout3 = Math.round(price * 0.60 * 100) / 100; // 40%下
      
      breakouts[asset.id].push({
        date: today,
        type: 'technical',
        level: downBreakout1,
        direction: 'down',
        rationale: `短期サポート$${downBreakout1.toLocaleString()}`
      });
      
      breakouts[asset.id].push({
        date: today,
        type: 'technical',
        level: downBreakout2,
        direction: 'down',
        rationale: `中期サポート$${downBreakout2.toLocaleString()}`
      });
      
      breakouts[asset.id].push({
        date: today,
        type: 'psychological',
        level: downBreakout3,
        direction: 'down',
        rationale: `危険水準$${downBreakout3.toLocaleString()}`
      });
    }
  });
  
  return breakouts;
}

/**
 * ブレークアウト保存
 */
function saveBreakouts(breakouts) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.BREAKOUTS);
  
  const data = [];
  Object.keys(breakouts).forEach(assetId => {
    breakouts[assetId].forEach(breakout => {
      data.push([
        assetId,
        breakout.date,
        breakout.type,
        breakout.level,
        breakout.direction,
        breakout.rationale
      ]);
    });
  });
  
  if (data.length > 0) {
    const lastRow = sheet.getLastRow();
    sheet.getRange(lastRow + 1, 1, data.length, 6).setValues(data);
  }
}

// =====================================
// 結果保存
// =====================================

/**
 * 結果保存
 */
function saveResults(results) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.RESULTS);
  const timestamp = new Date();
  
  const data = [];
  results.forEach(result => {
    result.horizons.forEach(horizon => {
      data.push([
        timestamp,
        result.asset_id,
        horizon,
        result.ev_by_horizon[horizon] || 0,
        result.staking_by_horizon ? result.staking_by_horizon[horizon] || 0 : 0,
        result.total_ev_by_horizon ? result.total_ev_by_horizon[horizon] || result.ev_by_horizon[horizon] : result.ev_by_horizon[horizon],
        result.confidence,
        result.models_used.join(','),
        JSON.stringify(result.scenarios),
        result.citations.join(' | ')
      ]);
    });
  });
  
  if (data.length > 0) {
    const lastRow = sheet.getLastRow();
    sheet.getRange(lastRow + 1, 1, data.length, 10).setValues(data);
  }
}

// =====================================
// レポート生成・送信
// =====================================

/**
 * HTMLレポート生成
 */
function generateHTMLReport(results, allBreakouts, config) {
  const today = new Date().toISOString().split('T')[0];
  const horizonCount = results[0]?.horizons.length || 0;
  
  let html = `
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <style>
    body { font-family: Arial, sans-serif; line-height: 1.6; color: #333; }
    h1 { color: #2c3e50; border-bottom: 2px solid #3498db; padding-bottom: 10px; }
    h2 { color: #34495e; margin-top: 30px; }
    table { border-collapse: collapse; width: 100%; margin: 20px 0; }
    th, td { border: 1px solid #ddd; padding: 12px; text-align: left; }
    th { background-color: #3498db; color: white; }
    tr:nth-child(even) { background-color: #f2f2f2; }
    .positive { color: #27ae60; font-weight: bold; }
    .negative { color: #e74c3c; font-weight: bold; }
    .neutral { color: #95a5a6; }
    .scenario { background-color: #ecf0f1; padding: 10px; margin: 10px 0; border-radius: 5px; }
    .breakout { background-color: #fff3cd; padding: 10px; margin: 10px 0; border-radius: 5px; border-left: 4px solid #ffc107; }
    .confidence { display: inline-block; padding: 3px 8px; border-radius: 3px; }
    .high-confidence { background-color: #d4edda; color: #155724; }
    .medium-confidence { background-color: #fff3cd; color: #856404; }
    .low-confidence { background-color: #f8d7da; color: #721c24; }
    .footer { margin-top: 40px; padding-top: 20px; border-top: 1px solid #ddd; color: #7f8c8d; font-size: 0.9em; }
  </style>
</head>
<body>
  <h1>🔮 暗号資産期待値レポート - ${today}</h1>
  <p>評価期間: ${horizonCount}期間 | 最終更新: ${new Date().toLocaleTimeString('ja-JP')}</p>
`;
  
  // ポートフォリオサマリー
  html += '<h2>📊 ポートフォリオサマリー</h2>';
  html += '<table>';
  html += '<tr><th>資産</th><th>基準価格</th><th>30日EV%</th><th>90日EV%</th><th>365日EV%</th><th>信頼度</th></tr>';
  
  results.forEach(result => {
    // データ検証
    if (!result || !result.asset_symbol || result.base_price === undefined) {
      console.error('無効な結果データ:', result);
      return;
    }
    
    const ev30 = result.ev_by_horizon[30] || 0;
    const ev90 = result.ev_by_horizon[90] || 0;
    const ev365 = result.ev_by_horizon[365] || 0;
    const confidence = ((result.confidence || 0) * 100).toFixed(0);
    
    const confidenceClass = (result.confidence || 0) > 0.8 ? 'high-confidence' : 
                           (result.confidence || 0) > 0.6 ? 'medium-confidence' : 'low-confidence';
    
    html += `<tr>
      <td><strong>${result.asset_symbol}</strong></td>
      <td>$${(result.base_price || 0).toLocaleString()}</td>
      <td class="${ev30 >= 0 ? 'positive' : 'negative'}">${ev30.toFixed(1)}%</td>
      <td class="${ev90 >= 0 ? 'positive' : 'negative'}">${ev90.toFixed(1)}%</td>
      <td class="${ev365 >= 0 ? 'positive' : 'negative'}">${ev365.toFixed(1)}%</td>
      <td><span class="confidence ${confidenceClass}">${confidence}%</span></td>
    </tr>`;
  });
  
  html += '</table>';
  
  // 各資産の詳細
  results.forEach(result => {
    html += `<h2>💎 ${result.asset_symbol} 詳細分析</h2>`;
    
    // 期待値テーブル
    html += '<h3>期待値</h3>';
    html += '<table>';
    html += '<tr><th>期間（日）</th><th>価格EV%</th><th>ステーキング%</th><th>総合EV%</th><th>期待価格</th></tr>';
    
    result.horizons.forEach(horizon => {
      const priceEV = result.ev_by_horizon[horizon] || 0;
      const stakingReturn = result.staking_by_horizon ? result.staking_by_horizon[horizon] || 0 : 0;
      const totalEV = result.total_ev_by_horizon ? result.total_ev_by_horizon[horizon] || priceEV : priceEV;
      const expectedPrice = result.base_price * (1 + totalEV / 100);
      
      html += `<tr>
        <td>${horizon}</td>
        <td class="${priceEV >= 0 ? 'positive' : 'negative'}">${priceEV.toFixed(2)}%</td>
        <td class="${stakingReturn > 0 ? 'positive' : 'neutral'}">${stakingReturn.toFixed(2)}%</td>
        <td class="${totalEV >= 0 ? 'positive' : 'negative'}">${totalEV.toFixed(2)}%</td>
        <td>$${expectedPrice.toLocaleString()}</td>
      </tr>`;
    });
    
    html += '</table>';
    
    // シナリオ分析
    html += '<h3>シナリオ分析</h3>';
    ['bear', 'base', 'bull'].forEach(scenarioName => {
      const scenario = result.scenarios[scenarioName];
      if (!scenario) return;
      
      const emoji = scenarioName === 'bear' ? '🐻' : scenarioName === 'bull' ? '🐂' : '📊';
      const prob = (scenario.prob * 100).toFixed(0);
      
      html += `<div class="scenario">
        <strong>${emoji} ${scenarioName.toUpperCase()}シナリオ (確率: ${prob}%)</strong><br>
        30日: ${scenario.return_by_horizon[30]?.toFixed(1) || 'N/A'}% | 
        90日: ${scenario.return_by_horizon[90]?.toFixed(1) || 'N/A'}% | 
        365日: ${scenario.return_by_horizon[365]?.toFixed(1) || 'N/A'}%<br>
        <em>根拠: ${scenario.rationales[0] || '分析中'}</em>
      </div>`;
    });
    
    // ブレークアウトポイント
    const assetBreakouts = allBreakouts[result.asset_id] || [];
    if (assetBreakouts.length > 0) {
      html += '<h3>🚀 ブレークアウトポイント</h3>';
      assetBreakouts.slice(0, 3).forEach(breakout => {
        const arrow = breakout.direction === 'up' ? '↑' : '↓';
        html += `<div class="breakout">
          <strong>${breakout.date}</strong> ${arrow} 
          ${breakout.type} @ $${breakout.level.toLocaleString()}<br>
          ${breakout.rationale}
        </div>`;
      });
    }
    
    // 信頼度のみ表示（モデル名は非表示）
    html += `<p><strong>分析信頼度:</strong> ${(result.confidence * 100).toFixed(0)}%</p>`;
  });
  
  // 参考情報セクションを削除（ユーザーリクエストにより）
  // 参考文献のURLは表示しない
  
  // フッター
  html += `
  <div class="footer">
    <p>⚠️ 投資に関する判断は自己責任でお願いします。本レポートは情報提供のみを目的としています。</p>
    <p>生成日時: ${new Date().toLocaleString('ja-JP', {timeZone: config.timezone || 'Asia/Tokyo'})}</p>
  </div>
</body>
</html>`;
  
  return html;
}

/**
 * メールレポート送信
 */
function sendEmailReport(htmlContent, config) {
  const recipients = config.recipients;
  if (!recipients) {
    console.log('送信先未設定のためメール送信をスキップ');
    return;
  }
  
  const today = new Date().toISOString().split('T')[0];
  const subject = `${config.subject_prefix || 'Crypto EV'} ${today}`;
  
  try {
    MailApp.sendEmail({
      to: recipients,
      subject: subject,
      htmlBody: htmlContent
    });
    
    console.log(`レポート送信完了: ${recipients}`);
  } catch (error) {
    console.error('メール送信エラー:', error);
    throw error;
  }
}

/**
 * Notionページを更新
 */
function updateNotionPage(results, allBreakouts, config) {
  const notionToken = PropertiesService.getScriptProperties().getProperty('NOTION_TOKEN');
  const pageId = config.notion_page_id;
  
  if (!notionToken) {
    console.log('Notion APIトークンが未設定のため、Notion更新をスキップ');
    return;
  }
  
  if (!pageId) {
    console.log('NotionページIDが未設定のため、Notion更新をスキップ');
    return;
  }
  
  // ページIDの形式をチェック
  if (!validateNotionPageId(pageId)) {
    console.error('無効なNotionページID形式:', pageId);
    throw new Error(`Invalid Notion page ID format: ${pageId}. Please check the page ID in Config sheet.`);
  }
  
  console.log(`Notion更新開始: ${pageId}`);
  const today = new Date().toISOString().split('T')[0];
  
  // 新しいコンテンツを生成（先に生成しておく）
  console.log('新しいコンテンツを生成中...');
  const notionContent = generateNotionContent(results, allBreakouts, today);
  console.log(`生成されたブロック数: ${notionContent.length}`);
  
  let clearingSuccessful = false;
  
  try {
    // ページプロパティを更新（最終更新日時）
    console.log('ページプロパティを更新中...');
    updateNotionPageProperties(pageId, notionToken);
    console.log('ページプロパティ更新完了');
    
    // 既存のページ内容をクリア
    console.log('既存コンテンツをクリア中...');
    clearNotionPageContent(pageId, notionToken);
    console.log('既存コンテンツクリア完了');
    clearingSuccessful = true;
  } catch (clearError) {
    console.warn('既存コンテンツのクリアに失敗しました:', clearError.message);
    console.warn('コンテンツの追加は続行します（既存コンテンツの上に追加されます）');
    
    // クリアに失敗した場合の詳細ログ
    if (clearError.message.includes('404')) {
      console.warn('ページが見つからないかアクセス権限がない可能性があります');
      console.warn('それでも新しいコンテンツの追加を試行します');
    }
  }
  
  try {
    // 新しいコンテンツを追加（クリアが失敗しても実行）
    console.log('新しいコンテンツを追加中...');
    appendNotionContent(pageId, notionToken, notionContent);
    console.log('新しいコンテンツ追加完了');
    
    if (clearingSuccessful) {
      console.log('Notion更新完了（クリア+追加）');
    } else {
      console.log('Notion更新完了（追加のみ）- 既存コンテンツの上に追加されました');
    }
  } catch (appendError) {
    console.error('新しいコンテンツの追加に失敗:', appendError);
    
    // より詳細なエラー情報を提供
    if (appendError.message.includes('404')) {
      console.error('解決方法:');
      console.error('1. Configシートのnotion_page_idが正しいかチェック');
      console.error('2. NotionページでIntegrationが招待されているかチェック');
      console.error('3. ページが削除されていないかチェック');
      console.error('4. Integrationに適切な権限（読み取り・書き込み）があるかチェック');
    }
    
    throw appendError;
  }
}

/**
 * NotionページIDの形式を検証
 */
function validateNotionPageId(pageId) {
  if (!pageId || typeof pageId !== 'string') {
    return false;
  }
  
  // ハイフンを除去
  const cleanId = pageId.replace(/-/g, '');
  
  // 32文字の16進数文字列かチェック
  return cleanId.length === 32 && /^[0-9a-f]{32}$/i.test(cleanId);
}

/**
 * Notionページのプロパティを更新（最終更新日時）
 */
function updateNotionPageProperties(pageId, notionToken) {
  const url = `https://api.notion.com/v1/pages/${pageId}`;
  
  const now = new Date();
  const jstTime = new Date(now.getTime() + (9 * 60 * 60 * 1000)); // JST時間に変換
  const isoString = jstTime.toISOString();
  
  const payload = {
    properties: {
      "最終更新日時": {
        date: {
          start: isoString
        }
      }
    }
  };
  
  const options = {
    method: 'PATCH',
    headers: {
      'Authorization': `Bearer ${notionToken}`,
      'Content-Type': 'application/json',
      'Notion-Version': '2022-06-28'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  
  try {
    console.log(`ページプロパティ更新: ${pageId}`);
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();
    
    if (responseCode === 200) {
      console.log('✓ ページプロパティ更新成功');
      console.log(`更新日時: ${jstTime.toLocaleString('ja-JP', {timeZone: 'Asia/Tokyo'})}`);
    } else {
      console.warn(`ページプロパティ更新失敗: HTTP ${responseCode}`);
      console.warn('応答:', responseText.substring(0, 300));
      
      // プロパティが存在しない場合のヒント
      if (responseCode === 400 && responseText.includes('property')) {
        console.warn('ヒント: Notionページに「最終更新日時」プロパティ（Date型）を追加してください');
      }
    }
  } catch (error) {
    console.warn('ページプロパティ更新エラー:', error.message);
    console.warn('プロパティ更新に失敗しましたが、処理を続行します');
  }
}

/**
 * Notionページの既存コンテンツをクリア（改善版：全ブロック確実削除）
 */
function clearNotionPageContent(pageId, notionToken) {
  console.log(`Notionページ内容クリア開始: ${pageId}`);
  
  let totalDeletedCount = 0;
  let maxRetries = 10; // 最大10回まで繰り返し
  
  for (let retry = 0; retry < maxRetries; retry++) {
    try {
      console.log(`削除サイクル ${retry + 1}/${maxRetries}`);
      
      // 既存の子ブロックを取得
      const url = `https://api.notion.com/v1/blocks/${pageId}/children`;
      const getOptions = {
        method: 'GET',
        headers: {
          'Authorization': `Bearer ${notionToken}`,
          'Notion-Version': '2022-06-28'
        },
        muteHttpExceptions: true
      };
      
      const response = UrlFetchApp.fetch(url, getOptions);
      const responseCode = response.getResponseCode();
      const responseText = response.getContentText();
      
      if (responseCode === 404) {
        console.error('Notion APIエラー: ページが見つかりません');
        console.error('ページID:', pageId);
        console.error('応答:', responseText.substring(0, 500));
        throw new Error(`Notion page not found: ${pageId}. Please check the page ID in Config sheet.`);
      }
      
      if (responseCode !== 200) {
        console.error(`Notion GET失敗: HTTP ${responseCode}`);
        console.error('応答:', responseText.substring(0, 500));
        throw new Error(`Notion API error: HTTP ${responseCode}`);
      }
      
      const data = JSON.parse(responseText);
      
      // ブロックがなくなったら完了
      if (!data.results || data.results.length === 0) {
        console.log(`✓ 全ブロック削除完了 (合計${totalDeletedCount}ブロック削除)`);
        return;
      }
      
      console.log(`削除対象ブロック数: ${data.results.length}`);
      
      // 各ブロックを削除
      let deletedCount = 0;
      let failedCount = 0;
      
      // ブロックを逆順で削除（ネストしたブロック対策）
      const blocksToDelete = [...data.results].reverse();
      
      for (const block of blocksToDelete) {
        try {
          const deleteUrl = `https://api.notion.com/v1/blocks/${block.id}`;
          const deleteOptions = {
            method: 'DELETE',
            headers: {
              'Authorization': `Bearer ${notionToken}`,
              'Notion-Version': '2022-06-28'
            },
            muteHttpExceptions: true
          };
          
          const deleteResponse = UrlFetchApp.fetch(deleteUrl, deleteOptions);
          const deleteCode = deleteResponse.getResponseCode();
          
          if (deleteCode === 200) {
            deletedCount++;
            totalDeletedCount++;
          } else if (deleteCode === 404) {
            // ブロックが既に削除済みの場合
            console.log(`ブロック既に削除済み: ${block.id}`);
            deletedCount++;
          } else {
            console.warn(`ブロック削除失敗: ${block.id}, HTTP ${deleteCode}`);
            const deleteResponseText = deleteResponse.getContentText();
            console.warn('削除エラー詳細:', deleteResponseText.substring(0, 200));
            failedCount++;
          }
          
          // レート制限対策
          Utilities.sleep(150);
          
        } catch (deleteError) {
          console.error(`ブロック削除エラー: ${block.id}`, deleteError);
          failedCount++;
        }
      }
      
      console.log(`削除サイクル ${retry + 1} 完了: 成功${deletedCount}件, 失敗${failedCount}件`);
      
      // 全て削除できた場合は次のサイクルへ
      if (failedCount === 0) {
        // 少し待機してから次のチェック
        Utilities.sleep(500);
        continue;
      } else {
        // 失敗があった場合は少し長めに待機
        console.warn(`${failedCount}件の削除に失敗しました。1秒待機してリトライします。`);
        Utilities.sleep(1000);
      }
      
    } catch (error) {
      console.error(`削除サイクル ${retry + 1} でエラー:`, error);
      if (retry === maxRetries - 1) {
        throw error; // 最後の試行でエラーの場合は上位に伝播
      }
      // 次のサイクルまで待機
      Utilities.sleep(1000);
    }
  }
  
  console.warn(`最大試行回数に達しました。合計${totalDeletedCount}ブロックを削除しましたが、まだ残っている可能性があります。`);
}

/**
 * Notionページに新しいコンテンツを追加
 */
function appendNotionContent(pageId, notionToken, blocks) {
  if (!blocks || blocks.length === 0) {
    console.log('追加するブロックがありません');
    return;
  }
  
  const url = `https://api.notion.com/v1/blocks/${pageId}/children`;
  console.log(`Notion コンテンツ追加開始: ${blocks.length}ブロック`);
  
  // ブロックを100個ずつに分割（Notion APIの制限）
  const chunkSize = 100;
  const totalChunks = Math.ceil(blocks.length / chunkSize);
  
  for (let i = 0; i < blocks.length; i += chunkSize) {
    const chunkIndex = Math.floor(i / chunkSize) + 1;
    const chunk = blocks.slice(i, i + chunkSize);
    
    console.log(`チャンク ${chunkIndex}/${totalChunks} を処理中 (${chunk.length}ブロック)`);
    
    const payload = {
      children: chunk
    };
    
    const options = {
      method: 'PATCH',
      headers: {
        'Authorization': `Bearer ${notionToken}`,
        'Content-Type': 'application/json',
        'Notion-Version': '2022-06-28'
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };
    
    try {
      const response = UrlFetchApp.fetch(url, options);
      const responseCode = response.getResponseCode();
      const responseText = response.getContentText();
      
      console.log(`チャンク ${chunkIndex} 応答: HTTP ${responseCode}`);
      
      if (responseCode === 404) {
        console.error(`Notion API 404エラー (チャンク ${chunkIndex}): ページが見つかりません`);
        console.error('ページID:', pageId);
        console.error('応答:', responseText.substring(0, 500));
        throw new Error(`Notion page not found: ${pageId}`);
      }
      
      if (responseCode === 401) {
        console.error(`Notion API 401エラー (チャンク ${chunkIndex}): 認証失敗`);
        console.error('Notion APIトークンが無効か、ページへのアクセス権限がありません');
        throw new Error('Notion API authentication failed. Check token and page permissions.');
      }
      
      if (responseCode === 403) {
        console.error(`Notion API 403エラー (チャンク ${chunkIndex}): アクセス拒否`);
        console.error('Integrationがページに招待されていない可能性があります');
        throw new Error('Notion API access denied. Integration not invited to page.');
      }
      
      if (responseCode !== 200) {
        console.error(`Notion更新失敗 (チャンク ${chunkIndex}): HTTP ${responseCode}`);
        console.error('応答:', responseText.substring(0, 500));
        
        // レスポンスからエラー詳細を抽出
        try {
          const errorData = JSON.parse(responseText);
          if (errorData.message) {
            console.error('エラーメッセージ:', errorData.message);
          }
          if (errorData.code) {
            console.error('エラーコード:', errorData.code);
          }
        } catch (parseError) {
          console.error('エラー応答の解析に失敗');
        }
        
        throw new Error(`Notion API Error: ${responseCode}`);
      }
      
      console.log(`チャンク ${chunkIndex} 追加成功`);
      
      // レート制限対策
      if (i + chunkSize < blocks.length) {
        console.log('レート制限対策で300ms待機中...');
        Utilities.sleep(300);
      }
    } catch (error) {
      console.error(`Notion API呼び出しエラー (チャンク ${chunkIndex}):`, error);
      console.error('エラー詳細:', error.message);
      
      // ペイロードサイズの情報も出力
      const payloadSize = JSON.stringify(payload).length;
      console.error(`ペイロードサイズ: ${payloadSize} bytes`);
      
      throw error;
    }
  }
  
  console.log(`Notion コンテンツ追加完了: ${totalChunks}チャンク処理済み`);
}

/**
 * Notion用のコンテンツを生成
 */
function generateNotionContent(results, allBreakouts, date) {
  const blocks = [];
  
  // ヘッダー
  blocks.push({
    object: 'block',
    type: 'heading_1',
    heading_1: {
      rich_text: [
        {
          type: 'text',
          text: {
            content: `🔮 暗号資産期待値レポート - ${date}`
          }
        }
      ]
    }
  });
  
  // 更新時刻
  const updateTime = new Date().toLocaleString('ja-JP', {timeZone: 'Asia/Tokyo'});
  blocks.push({
    object: 'block',
    type: 'paragraph',
    paragraph: {
      rich_text: [
        {
          type: 'text',
          text: {
            content: `最終更新: ${updateTime} (JST)`
          }
        }
      ]
    }
  });
  
  // サマリーセクション
  blocks.push({
    object: 'block',
    type: 'heading_2',
    heading_2: {
      rich_text: [
        {
          type: 'text',
          text: {
            content: '📊 ポートフォリオサマリー'
          }
        }
      ]
    }
  });
  
  // サマリーテーブルの代わりに各資産を個別に表示
  results.forEach(result => {
    const ev30 = result.ev_by_horizon[30] || 0;
    const ev90 = result.ev_by_horizon[90] || 0;
    const ev365 = result.ev_by_horizon[365] || 0;
    const confidence = ((result.confidence || 0) * 100).toFixed(0);
    
    // 色を期待値に基づいて決定
    const getColor = (ev) => {
      if (ev > 10) return 'green';
      if (ev > 0) return 'yellow';
      return 'red';
    };
    
    blocks.push({
      object: 'block',
      type: 'callout',
      callout: {
        icon: {
          type: 'emoji',
          emoji: '💎'
        },
        color: getColor(ev365),
        rich_text: [
          {
            type: 'text',
            text: {
              content: `${result.asset_symbol}: $${(result.base_price || 0).toLocaleString()} | 30日: ${ev30.toFixed(1)}% | 90日: ${ev90.toFixed(1)}% | 365日: ${ev365.toFixed(1)}% | 信頼度: ${confidence}%`
            }
          }
        ]
      }
    });
  });
  
  // 各資産の詳細
  results.forEach(result => {
    blocks.push({
      object: 'block',
      type: 'heading_2',
      heading_2: {
        rich_text: [
          {
            type: 'text',
            text: {
              content: `💎 ${result.asset_symbol} 詳細分析`
            }
          }
        ]
      }
    });
    
    // 基本情報
    blocks.push({
      object: 'block',
      type: 'paragraph',
      paragraph: {
        rich_text: [
          {
            type: 'text',
            text: {
              content: `基準価格: $${(result.base_price || 0).toLocaleString()}`
            },
            annotations: {
              bold: true
            }
          }
        ]
      }
    });
    
    // 期待値テーブル（テキスト形式）
    const evRows = result.horizons.map(horizon => {
      const priceEV = result.ev_by_horizon[horizon] || 0;
      const stakingReturn = result.staking_by_horizon ? result.staking_by_horizon[horizon] || 0 : 0;
      const totalEV = result.total_ev_by_horizon ? result.total_ev_by_horizon[horizon] || priceEV : priceEV;
      const expectedPrice = result.base_price * (1 + totalEV / 100);
      
      return `${horizon}日: 価格EV ${priceEV.toFixed(2)}% | ステーキング ${stakingReturn.toFixed(2)}% | 総合EV ${totalEV.toFixed(2)}% | 期待価格 $${expectedPrice.toLocaleString()}`;
    }).join('\n');
    
    blocks.push({
      object: 'block',
      type: 'code',
      code: {
        language: 'plain text',
        rich_text: [
          {
            type: 'text',
            text: {
              content: evRows
            }
          }
        ]
      }
    });
    
    // シナリオ分析
    ['bear', 'base', 'bull'].forEach(scenarioName => {
      const scenario = result.scenarios[scenarioName];
      if (!scenario) return;
      
      const emoji = scenarioName === 'bear' ? '🐻' : scenarioName === 'bull' ? '🐂' : '📊';
      const prob = (scenario.prob * 100).toFixed(0);
      const color = scenarioName === 'bear' ? 'red' : scenarioName === 'bull' ? 'green' : 'default';
      
      const returns = result.horizons.map(h => 
        `${h}日: ${scenario.return_by_horizon[h]?.toFixed(1) || 'N/A'}%`
      ).join(' | ');
      
      blocks.push({
        object: 'block',
        type: 'callout',
        callout: {
          icon: {
            type: 'emoji',
            emoji: emoji
          },
          color: color,
          rich_text: [
            {
              type: 'text',
              text: {
                content: `${scenarioName.toUpperCase()}シナリオ (確率: ${prob}%)\n${returns}\n根拠: ${scenario.rationales[0] || '分析中'}`
              }
            }
          ]
        }
      });
    });
    
    // ブレークアウトポイント
    const assetBreakouts = allBreakouts[result.asset_id] || [];
    if (assetBreakouts.length > 0) {
      blocks.push({
        object: 'block',
        type: 'heading_3',
        heading_3: {
          rich_text: [
            {
              type: 'text',
              text: {
                content: '🚀 ブレークアウトポイント'
              }
            }
          ]
        }
      });
      
      assetBreakouts.slice(0, 3).forEach(breakout => {
        const arrow = breakout.direction === 'up' ? '↑' : '↓';
        const color = breakout.direction === 'up' ? 'green' : 'red';
        
        blocks.push({
          object: 'block',
          type: 'callout',
          callout: {
            icon: {
              type: 'emoji',
              emoji: breakout.direction === 'up' ? '📈' : '📉'
            },
            color: color,
            rich_text: [
              {
                type: 'text',
                text: {
                  content: `${breakout.date} ${arrow} ${breakout.type} @ $${breakout.level.toLocaleString()}\n${breakout.rationale}`
                }
              }
            ]
          }
        });
      });
    }
    
    // 区切り線
    blocks.push({
      object: 'block',
      type: 'divider',
      divider: {}
    });
  });
  
  // 注意事項
  blocks.push({
    object: 'block',
    type: 'callout',
    callout: {
      icon: {
        type: 'emoji',
        emoji: '⚠️'
      },
      color: 'yellow',
      rich_text: [
        {
          type: 'text',
          text: {
            content: '投資に関する判断は自己責任でお願いします。本レポートは情報提供のみを目的としています。'
          }
        }
      ]
    }
  });
  
  // フッター（AI/API言及削除）
  blocks.push({
    object: 'block',
    type: 'paragraph',
    paragraph: {
      rich_text: [
        {
          type: 'text',
          text: {
            content: '暗号資産期待値レポートシステム'
          },
          annotations: {
            italic: true,
            color: 'gray'
          }
        }
      ]
    }
  });
  
  return blocks;
}

// =====================================
// ログ記録
// =====================================

/**
 * 実行ログ記録
 */
function logExecution(status, message) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.LOG);
  const timestamp = new Date();
  
  sheet.appendRow([timestamp, status, message]);
  
  // 古いログの削除（1000行を超えたら削除）
  const lastRow = sheet.getLastRow();
  if (lastRow > 1000) {
    sheet.deleteRows(2, lastRow - 1000);
  }
}

// =====================================
// トリガー管理
// =====================================

/**
 * 日次トリガー作成
 */
function createDailyTrigger() {
  const config = loadConfig();
  const hour = parseInt(config.report_hour) || 9;
  
  // 既存トリガー削除
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'main') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  // 新規トリガー作成
  ScriptApp.newTrigger('main')
    .timeBased()
    .everyDays(1)
    .atHour(hour)
    .create();
  
  console.log(`日次トリガー設定完了: 毎日${hour}時実行`);
}

/**
 * トリガー削除
 */
function removeTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    ScriptApp.deleteTrigger(trigger);
  });
  
  console.log('全トリガー削除完了');
}

// =====================================
// メニュー
// =====================================

/**
 * スプレッドシート開いた時のメニュー作成
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🔮 暗号資産EV')
    .addItem('レポート生成', 'main')
    .addItem('シート初期化', 'initializeSheets')
    .addSeparator()
    .addItem('現在価格取得', 'fetchCurrentPrices')
    .addSeparator()
    .addItem('日次トリガー設定', 'createDailyTrigger')
    .addItem('トリガー削除', 'removeTriggers')
    .addSeparator()
    .addItem('古い価格データ削除', 'clearOldPrices')
    .addItem('全価格データクリア', 'clearAllPrices')
    .addSeparator()
    .addItem('APIキー設定', 'showAPIKeyDialog')
    .addItem('Notion設定', 'showNotionDialog')
    .addItem('Notion接続テスト', 'testNotionConnection')
    .addItem('Grok-4無効化', 'disableGrok')
    .addItem('Grok-4有効化', 'enableGrok')
    .addSeparator()
    .addItem('対応銘柄一覧', 'showSupportedAssets')
    .addItem('プリセット設定', 'showPresetSettings')
    .addItem('テスト実行', 'testRun')
    .addToUi();
}

/**
 * 現在価格を取得してPricesシートに保存
 */
function fetchCurrentPrices() {
  try {
    console.log('現在価格取得開始');
    SpreadsheetApp.getActiveSpreadsheet().toast('現在価格を取得中...', '価格取得', 3);
    
    const today = new Date().toISOString().split('T')[0];
    
    // 空の価格データで savePricesForDate を呼び出すと、自動的に現在価格を取得
    savePricesForDate({}, today);
    
    SpreadsheetApp.getActiveSpreadsheet().toast('現在価格の取得が完了しました', '完了', 3);
    console.log('現在価格取得完了');
    
  } catch (error) {
    console.error('現在価格取得エラー:', error);
    SpreadsheetApp.getUi().alert(`価格取得エラー: ${error.message}`);
  }
}

/**
 * Grok-4を無効化
 */
function disableGrok() {
  PropertiesService.getScriptProperties().setProperty('USE_GROK', 'false');
  SpreadsheetApp.getUi().alert('Grok-4を無効化しました。Perplexityのみを使用します。');
}

/**
 * Grok-4を有効化
 */
function enableGrok() {
  PropertiesService.getScriptProperties().setProperty('USE_GROK', 'true');
  SpreadsheetApp.getUi().alert('Grok-4を有効化しました。');
}

// =====================================
// データ管理・メンテナンス
// =====================================

/**
 * 古いPricesデータを削除
 */
function clearOldPrices(keepDays = 7) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.PRICES);
  if (!sheet) return;
  
  const cutoffDate = new Date();
  cutoffDate.setDate(cutoffDate.getDate() - keepDays);
  const cutoffDateStr = cutoffDate.toISOString().split('T')[0];
  
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return;
  
  const data = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
  const rowsToDelete = [];
  
  data.forEach((row, index) => {
    if (row[1]) {
      const dateStr = row[1] instanceof Date ? row[1].toISOString().split('T')[0] : row[1];
      if (dateStr < cutoffDateStr) {
        rowsToDelete.push(index + 2); // +2 because arrays are 0-indexed but sheets are 1-indexed, and we start from row 2
      }
    }
  });
  
  // 行を削除（後ろから削除して行番号のずれを防ぐ）
  rowsToDelete.reverse().forEach(rowNum => {
    sheet.deleteRow(rowNum);
  });
  
  console.log(`古いPricesデータを削除: ${rowsToDelete.length}行`);
}

/**
 * Pricesシートを完全にクリア
 */
function clearAllPrices() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.PRICES);
  if (!sheet) return;
  
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.deleteRows(2, lastRow - 1);
    console.log('全Pricesデータをクリアしました');
  }
}

// =====================================
// 設定・テスト用関数
// =====================================

/**
 * APIキー設定ダイアログ
 */
function showAPIKeyDialog() {
  const html = `
    <div style="padding: 20px;">
      <h3>APIキー設定</h3>
      <p>スクリプトプロパティに以下のキーを設定してください：</p>
      <ul>
        <li>PERPLEXITY_API_KEY （必須）</li>
        <li>GROK_API_KEY （必須）</li>
      </ul>
      <p>設定方法: プロジェクト設定 → スクリプトプロパティ</p>
    </div>
  `;
  
  const htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(400)
    .setHeight(300);
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'APIキー設定');
}

/**
 * Notion設定ダイアログ
 */
function showNotionDialog() {
  const html = `
    <div style="padding: 20px;">
      <h3>Notion統合設定</h3>
      <p>メール送信と同時にNotionページを自動更新するための設定です。</p>
      
      <h4>📋 必要な設定</h4>
      <ol>
        <li><strong>Notion APIトークン</strong><br>
        スクリプトプロパティに <code>NOTION_TOKEN</code> として設定</li>
        <li><strong>NotionページID</strong><br>
        Configシートの <code>notion_page_id</code> に設定</li>
      </ol>
      
      <h4>🔧 設定手順</h4>
      <ol>
        <li>Notion → Settings & members → Integrations</li>
        <li>「Develop or manage integrations」をクリック</li>
        <li>「New integration」を作成</li>
        <li>生成されたトークンをコピー</li>
        <li>Google Apps Script → プロジェクト設定 → スクリプトプロパティ</li>
        <li>プロパティ <code>NOTION_TOKEN</code> にトークンを設定</li>
        <li>更新したいNotionページでIntegrationを招待</li>
        <li>ページIDをConfigシートの <code>notion_page_id</code> に設定</li>
      </ol>
      
      <h4>📝 ページID取得方法</h4>
      <p>NotionページのURL例:<br>
      <code>https://notion.so/<span style="background: yellow;">ページID</span>?v=...</code></p>
      <p>ページIDは32文字のランダム文字列です（ハイフンは除く）</p>
      
      <h4>⚡ 機能</h4>
      <ul>
        <li>メール送信と同時にNotionページを自動更新</li>
        <li>期待値データを見やすいフォーマットで表示</li>
        <li>シナリオ分析とブレークアウトポイントを含む</li>
        <li>色分けされた視覚的なレポート</li>
      </ul>
      
      <div style="background-color: #f0f8ff; padding: 10px; border-radius: 5px; margin-top: 15px;">
        <strong>💡 ヒント:</strong> 設定完了後、「テスト実行」でNotion更新をテストできます。
      </div>
    </div>
  `;
  
  const htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(600)
    .setHeight(700);
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Notion統合設定');
}

/**
 * 対応銘柄一覧ダイアログ
 */
function showSupportedAssets() {
  const assets = [
    { id: 'bitcoin', symbol: 'BTC', name: 'Bitcoin' },
    { id: 'ethereum', symbol: 'ETH', name: 'Ethereum' },
    { id: 'ripple', symbol: 'XRP', name: 'XRP' },
    { id: 'near', symbol: 'NEAR', name: 'NEAR Protocol' },
    { id: 'cardano', symbol: 'ADA', name: 'Cardano' },
    { id: 'aave', symbol: 'AAVE', name: 'Aave' },
    { id: 'hedera-hashgraph', symbol: 'HBAR', name: 'Hedera' },
    { id: 'the-graph', symbol: 'GRT', name: 'The Graph' },
    { id: 'algorand', symbol: 'ALGO', name: 'Algorand' },
    { id: 'maker', symbol: 'MKR', name: 'Maker' },
    { id: 'curve-dao-token', symbol: 'CRV', name: 'Curve' },
    { id: 'cosmos', symbol: 'ATOM', name: 'Cosmos' },
    { id: 'polkadot', symbol: 'DOT', name: 'Polkadot' },
    { id: 'polygon', symbol: 'POL', name: 'Polygon' },
    { id: 'avalanche-2', symbol: 'AVAX', name: 'Avalanche' },
    { id: 'chainlink', symbol: 'LINK', name: 'Chainlink' },
    { id: 'uniswap', symbol: 'UNI', name: 'Uniswap' },
    { id: 'litecoin', symbol: 'LTC', name: 'Litecoin' },
    { id: 'bitcoin-cash', symbol: 'BCH', name: 'Bitcoin Cash' },
    { id: 'stellar', symbol: 'XLM', name: 'Stellar' },
    { id: 'synthetix-network-token', symbol: 'SNX', name: 'Synthetix' },
    { id: 'solana', symbol: 'SOL', name: 'Solana' }
  ];
  
  let html = `
    <div style="padding: 20px; max-height: 500px; overflow-y: auto;">
      <h3>対応銘柄一覧</h3>
      <p>Assetsシートに以下のIDを入力してください：</p>
      <table style="width: 100%; border-collapse: collapse;">
        <tr style="background-color: #f0f0f0;">
          <th style="border: 1px solid #ddd; padding: 8px;">ID (必須)</th>
          <th style="border: 1px solid #ddd; padding: 8px;">シンボル</th>
          <th style="border: 1px solid #ddd; padding: 8px;">名称</th>
        </tr>`;
  
  assets.forEach(asset => {
    html += `
        <tr>
          <td style="border: 1px solid #ddd; padding: 8px; font-family: monospace;">${asset.id}</td>
          <td style="border: 1px solid #ddd; padding: 8px;">${asset.symbol}</td>
          <td style="border: 1px solid #ddd; padding: 8px;">${asset.name}</td>
        </tr>`;
  });
  
  html += `
      </table>
      <p style="margin-top: 10px; font-size: 0.9em; color: #666;">
        ※ IDは正確に入力してください（CoinGecko APIのID形式）
      </p>
    </div>
  `;
  
  const htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(500)
    .setHeight(600);
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, '対応銘柄一覧');
}

/**
 * プリセット設定ダイアログ
 */
function showPresetSettings() {
  const html = `
    <div style="padding: 20px;">
      <h3>プリセット設定</h3>
      <p>よく使われる設定パターンを選択してください：</p>
      
      <div style="margin: 20px 0;">
        <h4>🔵 基本設定（BTC/ETH中心）</h4>
        <button onclick="google.script.run.applyPreset('basic')" style="padding: 10px 20px; background: #4285f4; color: white; border: none; border-radius: 5px; cursor: pointer;">
          適用
        </button>
        <p style="font-size: 0.9em; color: #666;">
          BTC(50%), ETH(30%), XRP(10%), ADA(10%) - ステーキング無効
        </p>
      </div>
      
      <div style="margin: 20px 0;">
        <h4>🟢 DeFi中心</h4>
        <button onclick="google.script.run.applyPreset('defi')" style="padding: 10px 20px; background: #34a853; color: white; border: none; border-radius: 5px; cursor: pointer;">
          適用
        </button>
        <p style="font-size: 0.9em; color: #666;">
          ETH(40%), AAVE(20%), UNI(15%), CRV(15%), MKR(10%) - ステーキング有効
        </p>
      </div>
      
      <div style="margin: 20px 0;">
        <h4>🟡 レイヤー1中心</h4>
        <button onclick="google.script.run.applyPreset('layer1')" style="padding: 10px 20px; background: #fbbc04; color: white; border: none; border-radius: 5px; cursor: pointer;">
          適用
        </button>
        <p style="font-size: 0.9em; color: #666;">
          BTC(30%), ETH(25%), SOL(15%), ADA(10%), DOT(10%), AVAX(10%) - ステーキング有効
        </p>
      </div>
      
      <div style="margin: 20px 0;">
        <h4>⚪ 全銘柄均等配分</h4>
        <button onclick="google.script.run.applyPreset('equal')" style="padding: 10px 20px; background: #9aa0a6; color: white; border: none; border-radius: 5px; cursor: pointer;">
          適用
        </button>
        <p style="font-size: 0.9em; color: #666;">
          全22銘柄を均等配分（各約4.5%） - ステーキング無効
        </p>
      </div>
      
      <div style="margin: 20px 0;">
        <h4>🔴 リセット</h4>
        <button onclick="google.script.run.applyPreset('reset')" style="padding: 10px 20px; background: #ea4335; color: white; border: none; border-radius: 5px; cursor: pointer;">
          適用
        </button>
        <p style="font-size: 0.9em; color: #666;">
          全銘柄を無効化（weight=0）
        </p>
      </div>
      
      <p style="margin-top: 20px; font-size: 0.9em; color: #666;">
        ※ 適用後はAssetsシートで個別調整が可能です
      </p>
    </div>
  `;
  
  const htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(500)
    .setHeight(700);
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'プリセット設定');
}

/**
 * プリセット設定を適用
 */
function applyPreset(presetType) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.ASSETS);
  if (!sheet) {
    SpreadsheetApp.getUi().alert('Assetsシートが見つかりません。先にシート初期化を行ってください。');
    return;
  }
  
  const presets = {
    basic: {
      'bitcoin': { weight: 0.5, staking: false },
      'ethereum': { weight: 0.3, staking: false },
      'ripple': { weight: 0.1, staking: false },
      'cardano': { weight: 0.1, staking: false }
    },
    defi: {
      'ethereum': { weight: 0.4, staking: true, apy: 0.04, fee_rate: 0.1, haircut_risk: 0.15 },
      'aave': { weight: 0.2, staking: false },
      'uniswap': { weight: 0.15, staking: false },
      'curve-dao-token': { weight: 0.15, staking: false },
      'maker': { weight: 0.1, staking: false }
    },
    layer1: {
      'bitcoin': { weight: 0.3, staking: false },
      'ethereum': { weight: 0.25, staking: true, apy: 0.04, fee_rate: 0.1, haircut_risk: 0.15 },
      'solana': { weight: 0.15, staking: true, apy: 0.07, fee_rate: 0.1, haircut_risk: 0.15 },
      'cardano': { weight: 0.1, staking: true, apy: 0.03, fee_rate: 0.05, haircut_risk: 0.1 },
      'polkadot': { weight: 0.1, staking: true, apy: 0.12, fee_rate: 0.1, haircut_risk: 0.15 },
      'avalanche-2': { weight: 0.1, staking: true, apy: 0.09, fee_rate: 0.1, haircut_risk: 0.15 }
    },
    equal: {
      'bitcoin': { weight: 0.045, staking: false },
      'ethereum': { weight: 0.045, staking: false },
      'ripple': { weight: 0.045, staking: false },
      'near': { weight: 0.045, staking: false },
      'cardano': { weight: 0.045, staking: false },
      'aave': { weight: 0.045, staking: false },
      'hedera-hashgraph': { weight: 0.045, staking: false },
      'the-graph': { weight: 0.045, staking: false },
      'algorand': { weight: 0.045, staking: false },
      'maker': { weight: 0.045, staking: false },
      'curve-dao-token': { weight: 0.045, staking: false },
      'cosmos': { weight: 0.045, staking: false },
      'polkadot': { weight: 0.045, staking: false },
      'polygon': { weight: 0.045, staking: false },
      'avalanche-2': { weight: 0.045, staking: false },
      'chainlink': { weight: 0.045, staking: false },
      'uniswap': { weight: 0.045, staking: false },
      'litecoin': { weight: 0.045, staking: false },
      'bitcoin-cash': { weight: 0.045, staking: false },
      'stellar': { weight: 0.045, staking: false },
      'synthetix-network-token': { weight: 0.045, staking: false },
      'solana': { weight: 0.045, staking: false }
    },
    reset: {
      'bitcoin': { weight: 0, staking: false },
      'ethereum': { weight: 0, staking: false },
      'ripple': { weight: 0, staking: false },
      'near': { weight: 0, staking: false },
      'cardano': { weight: 0, staking: false },
      'aave': { weight: 0, staking: false },
      'hedera-hashgraph': { weight: 0, staking: false },
      'the-graph': { weight: 0, staking: false },
      'algorand': { weight: 0, staking: false },
      'maker': { weight: 0, staking: false },
      'curve-dao-token': { weight: 0, staking: false },
      'cosmos': { weight: 0, staking: false },
      'polkadot': { weight: 0, staking: false },
      'polygon': { weight: 0, staking: false },
      'avalanche-2': { weight: 0, staking: false },
      'chainlink': { weight: 0, staking: false },
      'uniswap': { weight: 0, staking: false },
      'litecoin': { weight: 0, staking: false },
      'bitcoin-cash': { weight: 0, staking: false },
      'stellar': { weight: 0, staking: false },
      'synthetix-network-token': { weight: 0, staking: false },
      'solana': { weight: 0, staking: false }
    }
  };
  
  const preset = presets[presetType];
  if (!preset) {
    SpreadsheetApp.getUi().alert('無効なプリセットタイプです。');
    return;
  }
  
  // データを更新
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 11).getValues();
  const updatedData = data.map(row => {
    const assetId = row[0];
    if (preset[assetId]) {
      const config = preset[assetId];
      row[3] = config.weight; // weight
      row[4] = config.staking; // staking_enabled
      if (config.apy !== undefined) row[5] = config.apy; // apy
      if (config.fee_rate !== undefined) row[7] = config.fee_rate; // fee_rate
      if (config.haircut_risk !== undefined) row[8] = config.haircut_risk; // haircut_risk
    }
    return row;
  });
  
  sheet.getRange(2, 1, updatedData.length, 11).setValues(updatedData);
  
  SpreadsheetApp.getUi().alert(`${presetType}プリセットを適用しました！`);
}

/**
 * Notion設定をテスト
 */
function testNotionConnection() {
  try {
    console.log('Notion接続テスト開始');
    
    const config = loadConfig();
    const notionToken = PropertiesService.getScriptProperties().getProperty('NOTION_TOKEN');
    const pageId = config.notion_page_id;
    
    console.log('設定チェック:');
    console.log(`- Notion APIトークン: ${notionToken ? '設定済み' : '未設定'}`);
    console.log(`- NotionページID: ${pageId || '未設定'}`);
    
    if (!notionToken) {
      throw new Error('Notion APIトークンが未設定です。スクリプトプロパティに NOTION_TOKEN を設定してください。');
    }
    
    if (!pageId) {
      throw new Error('NotionページIDが未設定です。ConfigシートのNotion_page_idを設定してください。');
    }
    
    // ページID形式チェック
    if (!validateNotionPageId(pageId)) {
      throw new Error(`無効なNotionページID形式: ${pageId}`);
    }
    
    console.log('✓ 基本設定OK');
    
    // ページアクセステスト
    const url = `https://api.notion.com/v1/blocks/${pageId}/children`;
    const options = {
      method: 'GET',
      headers: {
        'Authorization': `Bearer ${notionToken}`,
        'Notion-Version': '2022-06-28'
      },
      muteHttpExceptions: true
    };
    
    console.log('ページアクセステスト中...');
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();
    
    console.log(`応答: HTTP ${responseCode}`);
    
    if (responseCode === 200) {
      console.log('✓ ページアクセス成功');
      const data = JSON.parse(responseText);
      console.log(`現在のブロック数: ${data.results ? data.results.length : 0}`);
      
      SpreadsheetApp.getUi().alert('✅ Notion接続テスト成功！\n\n設定は正しく動作します。');
      return true;
    } else if (responseCode === 404) {
      console.error('✗ ページが見つかりません');
      SpreadsheetApp.getUi().alert('❌ Notion接続エラー\n\nページが見つかりません。\n\n確認事項:\n1. ページIDが正しいか\n2. ページが削除されていないか\n3. IntegrationがページにInviteされているか');
      return false;
    } else if (responseCode === 401) {
      console.error('✗ 認証失敗');
      SpreadsheetApp.getUi().alert('❌ Notion接続エラー\n\n認証に失敗しました。\n\n確認事項:\n1. APIトークンが正しいか\n2. Integrationが有効か');
      return false;
    } else if (responseCode === 403) {
      console.error('✗ アクセス拒否');
      SpreadsheetApp.getUi().alert('❌ Notion接続エラー\n\nアクセスが拒否されました。\n\n確認事項:\n1. IntegrationがページにInviteされているか\n2. Integrationに読み取り権限があるか');
      return false;
    } else {
      console.error(`✗ 予期しないエラー: HTTP ${responseCode}`);
      console.error('応答:', responseText.substring(0, 300));
      SpreadsheetApp.getUi().alert(`❌ Notion接続エラー\n\nHTTP ${responseCode}\n\n詳細はログを確認してください。`);
      return false;
    }
    
  } catch (error) {
    console.error('Notion接続テストエラー:', error);
    SpreadsheetApp.getUi().alert(`❌ Notion接続テストエラー\n\n${error.message}`);
    return false;
  }
}

/**
 * テスト実行
 */
function testRun() {
  console.log('テスト実行開始');
  
  // 初期化
  initializeSheets();
  
  // 簡易テストデータで実行
  const testAssets = [{
    id: 'bitcoin',
    symbol: 'BTC',
    name: 'Bitcoin',
    weight: 1.0,
    staking_enabled: false,
    apy: 0,
    compounding: 'annual',
    fee_rate: 0,
    haircut_risk: 0,
    is_lst: false
  }];
  
  // 現在のBitcoin価格を取得
  let currentBtcPrice = 61000; // デフォルト値
  try {
    console.log('テスト用Bitcoin現在価格を取得中...');
    const fetchedPrice = fetchPriceWithPerplexity('bitcoin');
    if (fetchedPrice && fetchedPrice > 0) {
      currentBtcPrice = fetchedPrice;
      console.log(`現在のBitcoin価格を取得: $${currentBtcPrice}`);
    } else {
      console.warn('Bitcoin価格取得失敗、デフォルト値を使用');
    }
  } catch (e) {
    console.error('Bitcoin価格取得エラー:', e);
    console.log('デフォルト値を使用');
  }
  
  const testPrices = { bitcoin: currentBtcPrice };
  const testHorizons = [30, 90, 365];
  
  // モックレスポンス生成（現在価格を使用）
  const mockResult = {
    asset_id: 'bitcoin',
    asset_symbol: 'BTC',
    base_price: currentBtcPrice,
    horizons: testHorizons,
    models_used: ['test'],
    scenarios: {
      bear: { name: 'bear', prob: 0.3, return_by_horizon: {30: -10, 90: -15, 365: -20}, rationales: ['テスト'] },
      base: { name: 'base', prob: 0.5, return_by_horizon: {30: 5, 90: 10, 365: 20}, rationales: ['テスト'] },
      bull: { name: 'bull', prob: 0.2, return_by_horizon: {30: 20, 90: 40, 365: 100}, rationales: ['テスト'] }
    },
    ev_by_horizon: {30: 2.5, 90: 6.5, 365: 18},
    confidence: 0.75,
    citations: ['https://example.com']
  };
  
  // レポート生成
  const config = { recipients: '', subject_prefix: 'TEST', timezone: 'Asia/Tokyo' };
  const html = generateHTMLReport([mockResult], {}, config);
  
  // HTMLファイルとして保存
  const blob = Utilities.newBlob(html, 'text/html', 'test_report.html');
  DriveApp.createFile(blob);
  
  console.log(`テスト実行完了: 現在価格$${currentBtcPrice}でレポートをGoogleドライブに保存しました`);
  
  SpreadsheetApp.getActiveSpreadsheet().toast(`現在価格$${currentBtcPrice.toLocaleString()}でテストレポートを生成しました`, 'テスト完了', 5);
}

// getCryptoPriceは削除（Perplexityのみ使用）

/**
 * CoinMarketCap価格取得（無料版）
 */
function getCoinMarketCapPrice(assetId) {
  // CoinMarketCapの無料APIは制限があるため、Perplexityで代替
  return null;
}

/**
 * CryptoCompare価格取得（無料版）
 */
function getCryptoComparePrice(assetId) {
  const symbolMap = {
    'bitcoin': 'BTC',
    'ethereum': 'ETH',
    'solana': 'SOL',
    'cardano': 'ADA',
    'polygon': 'MATIC',
    'chainlink': 'LINK',
    'polkadot': 'DOT',
    'avalanche-2': 'AVAX',
    'cosmos': 'ATOM',
    'near': 'NEAR'
  };
  
  const symbol = symbolMap[assetId] || assetId.toUpperCase();
  const url = `https://api.cryptocompare.com/data/price?fsym=${symbol}&tsyms=USD`;
  
  try {
    const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const data = JSON.parse(response.getContentText());
    
    if (data.USD) {
      return data.USD;
    }
  } catch (e) {
    console.error(`CryptoCompare error for ${assetId}:`, e);
  }
  
  return null;
}

/**
 * CoinCap価格取得（無料版）
 */
function getCoinCapPrice(assetId) {
  const url = `https://api.coincap.io/v2/assets/${assetId}`;
  
  try {
    const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const data = JSON.parse(response.getContentText());
    
    if (data.data && data.data.priceUsd) {
      return parseFloat(data.data.priceUsd);
    }
  } catch (e) {
    console.error(`CoinCap error for ${assetId}:`, e);
  }
  
  return null;
}

/**
 * Coinbase 現物スポット価格取得（APIキー不要）
 * 例: https://api.coinbase.com/v2/prices/BTC-USD/spot
 */
function getCoinbaseSpotPrice(assetId) {
  const symbolMap = {
    'bitcoin': 'BTC',
    'ethereum': 'ETH',
    'solana': 'SOL',
    'cardano': 'ADA',
    'polygon': 'MATIC',
    'chainlink': 'LINK',
    'polkadot': 'DOT',
    'avalanche-2': 'AVAX',
    'cosmos': 'ATOM',
    'near': 'NEAR'
  };
  const symbol = symbolMap[assetId];
  if (!symbol) return null;
  const url = `https://api.coinbase.com/v2/prices/${symbol}-USD/spot`;
  try {
    const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (response.getResponseCode() !== 200) return null;
    const data = JSON.parse(response.getContentText());
    const amount = data && data.data && data.data.amount ? parseFloat(data.data.amount) : null;
    return amount && amount > 0 ? amount : null;
  } catch (e) {
    console.error(`Coinbase error for ${assetId}:`, e);
    return null;
  }
}

/**
 * Binance 価格取得（USDT建て → USD換算として使用）
 * 例: https://api.binance.com/api/v3/ticker/price?symbol=BTCUSDT
 */
function getBinancePrice(assetId) {
  const symbolMap = {
    'bitcoin': 'BTC',
    'ethereum': 'ETH',
    'solana': 'SOL',
    'cardano': 'ADA',
    'polygon': 'MATIC',
    'chainlink': 'LINK',
    'polkadot': 'DOT',
    'avalanche-2': 'AVAX',
    'cosmos': 'ATOM',
    'near': 'NEAR'
  };
  const symbol = symbolMap[assetId];
  if (!symbol) return null;
  const url = `https://api.binance.com/api/v3/ticker/price?symbol=${symbol}USDT`;
  try {
    const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (response.getResponseCode() !== 200) return null;
    const data = JSON.parse(response.getContentText());
    const price = data && data.price ? parseFloat(data.price) : null;
    return price && price > 0 ? price : null; // USDT≒USDとして採用
  } catch (e) {
    console.error(`Binance error for ${assetId}:`, e);
    return null;
  }
}

/**
 * CoinGecko価格取得（無料版・改良版）
 */
function getCoinGeckoPrice(assetId) {
  // バッチ取得に変更したため、個別取得は使用しない
  console.warn('getCoinGeckoPrice: バッチ取得を使用してください');
  return null;
}

/**
 * CoinGecko価格一括取得（レート制限対策・改善版）
 */
function getCoinGeckoPricesBatch(assetIds) {
  if (!assetIds || assetIds.length === 0) return {};
  
  console.log(`CoinGecko一括取得開始 - 対象: ${assetIds.length}資産`);
  
  // APIキーがある場合は使用（ProまたはFreeプラン）
  const apiKey = PropertiesService.getScriptProperties().getProperty('COINGECKO_API_KEY');
  
  const maxRetries = 3;
  const idsParam = assetIds.join(',');
  
  // 現在の市場データと24時間変動も取得
  const baseUrl = apiKey 
    ? 'https://pro-api.coingecko.com/api/v3/simple/price'
    : 'https://api.coingecko.com/api/v3/simple/price';
  
  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    try {
      // 24時間変動と最終更新時刻は省略してシンプルにする（レート制限回避）
      const url = `${baseUrl}?ids=${idsParam}&vs_currencies=usd`;
      
      const headers = {
        'Accept': 'application/json',
        'User-Agent': 'CryptoEVReport/1.0'
      };
      
      // APIキーがある場合はヘッダーに追加
      if (apiKey) {
        headers['x-cg-pro-api-key'] = apiKey;
      }
      
      const options = {
        method: 'GET',
        headers: headers,
        muteHttpExceptions: true
      };
      
      console.log(`CoinGecko API呼び出し (試行${attempt}回目)...`);
      const response = UrlFetchApp.fetch(url, options);
      const responseCode = response.getResponseCode();
      
      if (responseCode === 200) {
        const data = JSON.parse(response.getContentText());
        const prices = {};
        let successCount = 0;
        
        assetIds.forEach(id => {
          if (data[id] && data[id].usd) {
            prices[id] = data[id].usd;
            console.log(`✓ ${id}: $${data[id].usd.toLocaleString()}`);
            successCount++;
          } else {
            console.warn(`✗ ${id}: データなし`);
          }
        });
        
        console.log(`CoinGecko取得完了: ${successCount}/${assetIds.length}資産`);
        return prices;
        
      } else if (responseCode === 429) {
        console.warn(`CoinGecko レート制限 (${attempt}回目)`);
        if (attempt < maxRetries) {
          const waitTime = apiKey ? 5000 : 10000; // APIキーありなら5秒、なしなら10秒
          console.log(`${waitTime/1000}秒待機してリトライ...`);
          Utilities.sleep(waitTime);
        }
      } else {
        console.warn(`CoinGecko HTTP ${responseCode} (${attempt}回目)`);
        const responseText = response.getContentText();
        console.error('エラー詳細:', responseText.substring(0, 200));
        
        if (attempt < maxRetries) {
          Utilities.sleep(5000); // 5秒待機
        }
      }
    } catch (e) {
      console.error(`CoinGecko batch error (${attempt}回目):`, e.message);
      if (attempt < maxRetries) {
        Utilities.sleep(5000); // 5秒待機してリトライ
      }
    }
  }
  
  console.error('CoinGecko価格取得失敗（全試行終了）');
  return {};
}

/**
 * Perplexityで複数資産の価格を一括取得
 */
function fetchAllPricesWithPerplexity(assetIds) {
  console.log('fetchAllPricesWithPerplexity呼び出し - 引数:', assetIds);
  
  // デフォルトの資産リスト（引数が無い場合）
  if (!assetIds) {
    console.log('引数なしで呼び出されたため、デフォルトの資産リストを使用');
    assetIds = [
      'bitcoin', 'ethereum', 'ripple', 'near', 'cardano',
      'aave', 'hedera-hashgraph', 'the-graph', 'algorand',
      'maker', 'curve-dao-token', 'cosmos',
      'polkadot', 'polygon', 'avalanche-2', 'chainlink',
      'uniswap', 'litecoin', 'bitcoin-cash', 'stellar',
      'synthetix-network-token', 'solana'
    ];
  }
  
  // 引数のバリデーション
  if (!Array.isArray(assetIds) || assetIds.length === 0) {
    console.error('無効なassetIds:', assetIds);
    throw new Error('assetIdsが無効です（配列でない、または空）');
  }
  
  const apiKey = PropertiesService.getScriptProperties().getProperty('PERPLEXITY_API_KEY');
  if (!apiKey) throw new Error('Perplexity APIキー未設定');
  
  const symbolMap = {
    'bitcoin': 'Bitcoin BTC',
    'ethereum': 'Ethereum ETH', 
    'ripple': 'XRP',
    'near': 'NEAR Protocol NEAR',
    'cardano': 'Cardano ADA',
    'aave': 'Aave AAVE',
    'hedera-hashgraph': 'Hedera HBAR',
    'the-graph': 'The Graph GRT',
    'algorand': 'Algorand ALGO',

    'maker': 'Maker MKR',
    'curve-dao-token': 'Curve DAO Token CRV',
    'cosmos': 'Cosmos ATOM',
    'polkadot': 'Polkadot DOT',
    'polygon': 'Polygon POL',
    'avalanche-2': 'Avalanche AVAX',
    'chainlink': 'Chainlink LINK',
    'uniswap': 'Uniswap UNI',
    'litecoin': 'Litecoin LTC',
    'bitcoin-cash': 'Bitcoin Cash BCH',
    'stellar': 'Stellar XLM',
    'synthetix-network-token': 'Synthetix SNX',
    'solana': 'Solana SOL'
  };
  
  console.log('価格取得対象資産IDs:', assetIds);
  const assetNames = assetIds.map(id => {
    const name = symbolMap[id] || id;
    console.log(`アセットマッピング: ${id} -> ${name}`);
    return name;
  }).join(', ');
  
  const currentTime = new Date().toISOString();
  const currentTimeJST = new Date().toLocaleString('ja-JP', {timeZone: 'Asia/Tokyo'});
  
  // シンボルマッピングを使用して通貨名とシンボルを明確にする
  const currencyList = assetIds.map(id => {
    const name = symbolMap[id] || id;
    return `${id}: ${name}`;
  }).join('\n');
  
  const prompt = `🚨 URGENT: Get the LIVE CURRENT cryptocurrency prices in US DOLLARS (USD) RIGHT NOW as of ${currentTime} (JST: ${currentTimeJST}).

⚠️ CRITICAL: I need the MOST RECENT REAL-TIME prices, NOT historical or cached data!

IMPORTANT: All prices MUST be in USD (United States Dollar), NOT JPY, EUR, or any other currency.

I need the LIVE CURRENT USD prices for these cryptocurrencies:
${currencyList}

🔥 CRITICAL REQUIREMENTS:
1. ALL prices MUST be in USD (US Dollars) - NO OTHER CURRENCY
2. Use LIVE CURRENT market prices from major exchanges (Binance, Coinbase, Kraken) RIGHT NOW
3. Return ONLY numeric values in USD
4. Do NOT include currency symbols ($, ¥, €)
5. Do NOT convert from other currencies - use direct USD prices
6. Get the MOST RECENT prices available - NOT old/cached data

Return ONLY a JSON object with cryptocurrency IDs and their CURRENT USD prices.
For reference, typical CURRENT USD prices are:
- Bitcoin: around $100,000-120,000 USD
- Ethereum: around $3,500-4,500 USD
- Solana: around $180-250 USD

Example format (use ACTUAL CURRENT LIVE USD PRICES):
{
  "bitcoin": 115000.00,
  "ethereum": 4200.00,
  "ripple": 3.12,
  "near": 8.45,
  "cardano": 1.02,
  "aave": 315.00,
  "hedera-hashgraph": 0.34,
  "the-graph": 0.43,
  "algorand": 0.41,
  "maker": 3150.00,
  "curve-dao-token": 1.18,
  "cosmos": 13.20,
  "polkadot": 10.50,
  "polygon": 2.25,
  "avalanche-2": 47.00,
  "chainlink": 21.50,
  "uniswap": 15.20,
  "litecoin": 103.00,
  "bitcoin-cash": 515.00,
  "stellar": 0.44,
  "synthetix-network-token": 4.15,
  "solana": 195.00
}

🚨 Remember: ALL prices in USD only! Get LIVE CURRENT prices RIGHT NOW!`;
  
  const url = 'https://api.perplexity.ai/chat/completions';
  const payload = {
    model: getPerplexityModel(),
    messages: [
      {
        role: 'system',
        content: 'You are a cryptocurrency price API. CRITICAL: Return ONLY USD (United States Dollar) prices. Do NOT return JPY, EUR, or any other currency. All prices must be in USD. Get current prices from major USD exchanges like Coinbase, Binance US, Kraken. Return complete JSON with numeric USD values only.'
      },
      {
        role: 'user',
        content: prompt
      }
    ],
    temperature: 0.0,
    max_tokens: 1000,
    stream: false,
    return_citations: false,
    search_recency_filter: 'hour'
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
    const response = UrlFetchApp.fetch(url, options);
    if (response.getResponseCode() !== 200) {
      console.error(`Perplexity HTTP エラー: ${response.getResponseCode()}`);
      console.error('応答内容:', response.getContentText().substring(0, 300));
      throw new Error(`HTTP ${response.getResponseCode()}`);
    }
    
    const data = JSON.parse(response.getContentText());
    const content = data.choices && data.choices[0] && data.choices[0].message
      ? data.choices[0].message.content : '';
    
    if (!content) throw new Error('応答が空です');
    
    console.log('Perplexity原文応答（先頭500文字）:', content.substring(0, 500));
    
    // JSON抽出の改善
    const cleaned = content.replace(/```json|```/g, '').trim();
    console.log('クリーン後応答（先頭300文字）:', cleaned.substring(0, 300));
    
    // より柔軟なJSON抽出
    let jsonMatch = cleaned.match(/\{[\s\S]*\}/);
    
    if (jsonMatch) {
      try {
        const sanitized = sanitizeJson(jsonMatch[0]);
        console.log('サニタイズ後JSON:', sanitized);
        const prices = JSON.parse(sanitized);
        
        // 価格データの検証
        const validPrices = {};
        Object.keys(prices).forEach(key => {
          const value = prices[key];
          if (typeof value === 'number' && value > 0 && value < 10000000) {
            validPrices[key] = value;
          } else {
            console.warn(`無効な価格データを除外: ${key} = ${value}`);
          }
        });
        
        console.log('Perplexity一括価格取得成功:', validPrices);
        return validPrices;
      } catch (parseError) {
        console.error('JSON パースエラー:', parseError);
        console.error('パース対象文字列:', jsonMatch[0].substring(0, 200));
      }
    } else {
      console.warn('JSON形式が見つかりません。応答内容:', cleaned);
    }
  } catch (e) {
    console.error('Perplexity一括価格取得エラー:', e);
    console.error('エラー詳細:', e.stack);
  }
  
  throw new Error('価格一括取得失敗');
}

/**
 * Perplexityで価格取得（フォールバック）
 */
function fetchPriceWithPerplexity(assetId) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('PERPLEXITY_API_KEY');
  if (!apiKey) throw new Error('Perplexity APIキー未設定');
  
  const symbolMap = {
    'bitcoin': 'Bitcoin BTC',
    'ethereum': 'Ethereum ETH', 
    'ripple': 'XRP',
    'near': 'NEAR Protocol NEAR',
    'cardano': 'Cardano ADA',
    'aave': 'Aave AAVE',
    'hedera-hashgraph': 'Hedera HBAR',
    'the-graph': 'The Graph GRT',
    'algorand': 'Algorand ALGO',

    'maker': 'Maker MKR',
    'curve-dao-token': 'Curve DAO Token CRV',
    'cosmos': 'Cosmos ATOM',
    'polkadot': 'Polkadot DOT',
    'polygon': 'Polygon POL',
    'avalanche-2': 'Avalanche AVAX',
    'chainlink': 'Chainlink LINK',
    'uniswap': 'Uniswap UNI',
    'litecoin': 'Litecoin LTC',
    'bitcoin-cash': 'Bitcoin Cash BCH',
    'stellar': 'Stellar XLM',
    'synthetix-network-token': 'Synthetix SNX',
    'solana': 'Solana SOL'
  };
  
  const assetName = symbolMap[assetId] || assetId;
  const currentTime = new Date().toISOString();
  const currentTimeJST = new Date().toLocaleString('ja-JP', {timeZone: 'Asia/Tokyo'});
  
  console.log(`個別価格取得開始: ${assetId} -> ${assetName}`);
  const prompts = [
    // 1回目: 数字のみの応答を強制
    {
      system: 'You are a LIVE price checker. Get CURRENT REAL-TIME prices. Respond with only a plain numeric USD price, no symbols, no commas, no text.',
      user: `🚨 URGENT: RIGHT NOW at ${currentTime} (JST: ${currentTimeJST}), what is the LIVE CURRENT USD price of ${assetName}? Get the MOST RECENT price from exchanges. Respond with only the number like 61234.56`
    },
    // 2回目: JSONでの応答を強制
    {
      system: 'You are a LIVE price API. Get CURRENT REAL-TIME prices. Respond only with valid JSON: {"price": <number>} with CURRENT USD price. No extra text.',
      user: `🚨 URGENT: Return JSON only for the LIVE CURRENT USD price of ${assetName} RIGHT NOW at ${currentTime}: {"price": <number>}`
    }
  ];
  
  const url = 'https://api.perplexity.ai/chat/completions';
  
  for (let i = 0; i < prompts.length; i++) {
    const payload = {
      model: getPerplexityModel(),
      messages: [
        { role: 'system', content: prompts[i].system },
        { role: 'user', content: prompts[i].user }
      ],
      temperature: 0.0,
      max_tokens: 100
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
      const response = UrlFetchApp.fetch(url, options);
      if (response.getResponseCode() !== 200) continue;
      const data = JSON.parse(response.getContentText());
      const content = (data.choices && data.choices[0] && data.choices[0].message && data.choices[0].message.content)
        ? data.choices[0].message.content.trim()
        : '';
      if (!content) continue;
      
      // コードブロック除去
      const cleaned = content.replace(/```json|```/g, '').trim();
      
      // まずJSONを試す
      const jsonMatch = cleaned.match(/\{[\s\S]*\}/);
      if (jsonMatch) {
        try {
          const obj = JSON.parse(jsonMatch[0]);
          const priceFromJson = obj.price || obj.usd || obj.value;
          if (priceFromJson && Number(priceFromJson) > 0) {
            return Number(priceFromJson);
          }
        } catch (_) { /* JSONでなければ無視 */ }
      }
      
      // 数値抽出（汎用・高互換）
      const numberMatches = cleaned.match(/[0-9]+(?:,[0-9]{3})*(?:\.[0-9]+)?/g);
      if (numberMatches && numberMatches.length > 0) {
        // 最初に見つかった妥当な値を採用
        for (let m of numberMatches) {
          const numeric = parseFloat(m.replace(/,/g, ''));
          if (!isNaN(numeric) && numeric > 0.01 && numeric < 10000000) {
            return numeric;
          }
        }
      }
    } catch (e) {
      // 次のプロンプトへフォールバック
      console.error(`Perplexity価格取得試行${i + 1}回目でエラー:`, e);
    }
  }
  
  // 価格抽出に失敗した場合は null を返す（上位で他ソースにフォールバック）
  return null;
}