// ===== 設定定数 =====
const CONFIG = {
  // APIキー
  // 重要：本番環境では、APIキーをPropertiesServiceに保存することを推奨
  // PropertiesService.getScriptProperties().setProperty('COINGECKO_API_KEY', 'your-key');
  COINGECKO_API_KEY: 'CG-yMMAj3DyqfeuGeR8ZWa4FSA1', // 月間10,000コール制限
  
  // スプレッドシートの設定
  SHEETS: {
    PORTFOLIO: 'ポートフォリオ',
    ALERTS: 'アラート設定',
    HISTORY: '取引履歴',
    ACCUMULATION: '積立管理',
    REBALANCE: 'リバランス提案',
    PRICE_HISTORY: '価格履歴',
    BREAKOUT: 'ブレイクアウト',
    PSYCHOLOGY: '心理管理',
    OCO_ORDERS: 'OCO注文',
    EDGE_ANALYSIS: 'エッジ分析',
    SETTINGS: '設定' // 新規追加
  },
  
  // 通知設定（動的に読み込むため、関数で取得）
  get EMAIL_RECIPIENT() {
    return getEmailRecipient();
  },
  
  // 取引所設定
  EXCHANGE_FEE: 0.00075, // 0.075%
  
  // ポートフォリオ設定（ドキュメント反映版）
  TARGET_ALLOCATION: {
    'BTC': 0.10,
    'ETH': 0.25,
    'SOL': 0.20,
    'NEAR': 0.10,
    'ADA': 0.08,
    'AAVE': 0.07,
    'HBAR': 0.05,
    'GRT': 0.05,
    'ALGO': 0.04,
    'CASH': 0.06  // 現金待機（ドキュメントでは8%推奨も）
  },
  
  // トレンドフォロー設定（タートル流）
  BREAKOUT_PERIODS: {
    BASE: 20,    // 基盤層：20日ブレイクアウト
    GROWTH: 60   // 成長層：60日ブレイクアウト
  },
  
  // 階段指値設定（ドキュメント反映）
  LIMIT_ORDER_LEVELS: {
    BASE: [-0.02],                    // 基盤層：-2%
    GROWTH: [-0.05, -0.10, -0.15],   // 成長層：-5%, -10%, -15%
    SATELLITE: [-0.05]                // 衛星層：-5%
  },
  
  // リスク管理設定
  VOLATILITY_THRESHOLD: 0.25,  // 25%超でリバランス
  MAX_DRAWDOWN: -0.56,         // 最大許容ドローダウン -56%
  EDGE_RATIO_MIN: 1.0,         // 最小エッジ比率（MFE/MAE）
  
  // OCO設定（Take Profit / Stop Loss）
  OCO_SETTINGS: {
    'BTC': { tp: 0.30, sl: -0.20 },
    'ETH': { tp: 0.36, sl: -0.25 },
    'SOL': { tp: 0.43, sl: -0.29 },
    'NEAR': { tp: 0.50, sl: -0.35 },
    'ADA': { tp: 0.40, sl: -0.30 },
    'AAVE': { tp: 0.40, sl: -0.30 },
    'HBAR': { tp: 0.45, sl: -0.32 },
    'GRT': { tp: 0.45, sl: -0.32 },
    'ALGO': { tp: 0.40, sl: -0.30 }
  },
  
  // 心理管理設定（ZONE）
  PSYCHOLOGY_BELIEFS: [
    '何事も起こり得る（不確定性）',
    '利益を出すのに次に何が起こるか知る必要はない',
    '優位性のある変数の分布はランダム',
    'エッジは高い確率を示すに過ぎない',
    'あらゆる瞬間は独自なもの'
  ],
  
  // 心理バイアスチェック項目
  BIAS_CHECKLIST: [
    '損失回避',
    '直近偏向',
    'バンドワゴン効果',
    'アンカリング',
    '結果偏向',
    '処理効果'
  ],
  
  // リバランス閾値
  REBALANCE_THRESHOLD: 0.02, // 2%以上の乖離でリバランス提案
  
  // 積立設定（動的に読み込むため、関数で取得）
  get MONTHLY_INVESTMENT() {
    return getMonthlyInvestment();
  },
  ACCUMULATION_MONTHS: 3, // 3ヶ月ごとに投資
  
  // キャッシュ設定
  USE_CACHE: true,
  CACHE_DURATION: 5 // 5分
};

// ===== 設定管理関数 =====

/**
 * 設定シートから毎月の積立金額を取得
 */
function getMonthlyInvestment() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settingsSheet = ss.getSheetByName(CONFIG.SHEETS.SETTINGS);
  
  if (!settingsSheet) {
    console.warn('設定シートが見つかりません。デフォルト値を使用します。');
    return 1500000; // デフォルト値
  }
  
  // B2セルから積立金額を取得
  const amount = settingsSheet.getRange('B2').getValue();
  
  // 数値チェック
  if (!amount || isNaN(amount) || amount <= 0) {
    console.warn('無効な積立金額です。デフォルト値を使用します。');
    return 1500000; // デフォルト値
  }
  
  return amount;
}

/**
 * 設定シートから通知先メールアドレスを取得
 */
function getEmailRecipient() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settingsSheet = ss.getSheetByName(CONFIG.SHEETS.SETTINGS);
  
  if (!settingsSheet) {
    console.warn('設定シートが見つかりません。現在のユーザーのメールアドレスを使用します。');
    return Session.getActiveUser().getEmail();
  }
  
  // B3セルからメールアドレスを取得
  const email = settingsSheet.getRange('B3').getValue();
  
  // メールアドレスの簡易チェック
  if (!email || !email.toString().includes('@')) {
    console.warn('無効なメールアドレスです。現在のユーザーのメールアドレスを使用します。');
    return Session.getActiveUser().getEmail();
  }
  
  return email.toString();
}

/**
 * 設定ダイアログを表示
 */
function showSettingsDialog() {
  const html = `
    <div style="padding: 20px;">
      <h3>システム設定</h3>
      <form onsubmit="handleSubmit(event)">
        <div style="margin-bottom: 20px;">
          <label>毎月の積立金額（円）:</label><br>
          <input type="number" id="monthlyInvestment" min="0" step="10000" required 
                 style="width: 100%; padding: 5px; margin-top: 5px;">
          <small style="color: #666;">例: 1500000（150万円）</small>
        </div>
        <div style="margin-bottom: 20px;">
          <label>通知先メールアドレス:</label><br>
          <input type="email" id="emailRecipient" required 
                 style="width: 100%; padding: 5px; margin-top: 5px;">
          <small style="color: #666;">アラートや通知が送信されます</small>
        </div>
        <button type="submit" style="width: 100%; padding: 10px; background: #4285f4; color: white; border: none; cursor: pointer;">
          設定を保存
        </button>
      </form>
    </div>
    <script>
      // 現在の設定を読み込む
      google.script.run
        .withSuccessHandler(function(settings) {
          document.getElementById('monthlyInvestment').value = settings.monthlyInvestment;
          document.getElementById('emailRecipient').value = settings.emailRecipient;
        })
        .getCurrentSettings();
      
      function handleSubmit(event) {
        event.preventDefault();
        const data = {
          monthlyInvestment: parseFloat(document.getElementById('monthlyInvestment').value),
          emailRecipient: document.getElementById('emailRecipient').value
        };
        google.script.run
          .withSuccessHandler(() => {
            google.script.host.close();
          })
          .updateSettings(data);
      }
    </script>
  `;
  
  const ui = SpreadsheetApp.getUi();
  ui.showModalDialog(HtmlService.createHtmlOutput(html).setWidth(400).setHeight(350), 'システム設定');
}

/**
 * 現在の設定を取得（ダイアログ用）
 */
function getCurrentSettings() {
  return {
    monthlyInvestment: getMonthlyInvestment(),
    emailRecipient: getEmailRecipient()
  };
}

/**
 * 設定を更新
 */
function updateSettings(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let settingsSheet = ss.getSheetByName(CONFIG.SHEETS.SETTINGS);
  
  // 設定シートがない場合は作成
  if (!settingsSheet) {
    createSettingsSheet(ss);
    settingsSheet = ss.getSheetByName(CONFIG.SHEETS.SETTINGS);
  }
  
  // 設定を更新
  settingsSheet.getRange('B2').setValue(data.monthlyInvestment);
  settingsSheet.getRange('B3').setValue(data.emailRecipient);
  
  SpreadsheetApp.getActiveSpreadsheet().toast('設定を更新しました', '完了', 3);
}

// ===== メイン関数 =====

/**
 * API接続テスト（デバッグ用）
 */
function testAPIs() {
  const testSymbol = 'BTC';
  const results = [];
  
  // 1. Google Finance
  try {
    const price = getGoogleFinancePrice(testSymbol);
    results.push(`Google Finance: ${price ? price : 'Failed'}`);
  } catch (e) {
    results.push(`Google Finance: Error - ${e.message}`);
  }
  
  // 2. Binance
  try {
    const price = getBinancePrice(testSymbol);
    results.push(`Binance: ${price ? price : 'Failed'}`);
  } catch (e) {
    results.push(`Binance: Error - ${e.message}`);
  }
  
  // 3. CryptoCompare
  try {
    const price = getCryptoComparePrice(testSymbol);
    results.push(`CryptoCompare: ${price ? price : 'Failed'}`);
  } catch (e) {
    results.push(`CryptoCompare: Error - ${e.message}`);
  }
  
  // 4. CoinCap
  try {
    const price = getCoinCapPrice(testSymbol);
    results.push(`CoinCap: ${price ? price : 'Failed'}`);
  } catch (e) {
    results.push(`CoinCap: Error - ${e.message}`);
  }
  
  // 5. CoinGecko
  try {
    const price = getCoinGeckoPrice(testSymbol);
    results.push(`CoinGecko: ${price ? price : 'Failed'}`);
  } catch (e) {
    results.push(`CoinGecko: Error - ${e.message}`);
  }
  
  // 6. Kraken (新規追加)
  try {
    const price = getKrakenPrice(testSymbol);
    results.push(`Kraken: ${price ? price : 'Failed'}`);
  } catch (e) {
    results.push(`Kraken: Error - ${e.message}`);
  }
  
  // 結果を表示
  const ui = SpreadsheetApp.getUi();
  ui.alert('API接続テスト結果', results.join('\n'), ui.ButtonSet.OK);
}

/**
 * 心理チェックを実行してダイアログ表示
 */
function executePsychologyCheck() {
  const result = performPsychologyCheck();
  const ui = SpreadsheetApp.getUi();
  
  if (result.passed) {
    ui.alert('心理チェック完了', 
             '✅ すべての項目がクリアされました。\n取引を実行できます。', 
             ui.ButtonSet.OK);
  } else {
    ui.alert('心理チェック未完了', 
             '❌ ' + result.message, 
             ui.ButtonSet.OK);
  }
}

/**
 * API使用状況をダイアログで表示
 */
function showAPIUsageDialog() {
  const report = getAPIUsageReport();
  const html = `
    <div style="padding: 20px;">
      <h3>API使用状況</h3>
      <p><strong>CoinGecko API</strong></p>
      <p>使用量: ${report.used} / ${report.limit}</p>
      <p>使用率: ${report.percentage}%</p>
      <p>残り: ${report.remaining}コール</p>
      <hr>
      <p style="font-size: 12px; color: #666;">
        ※月間10,000コールの制限があります<br>
        ※毎月1日にリセットされます<br>
        ※80%を超えると警告メールが送信されます
      </p>
    </div>
  `;
  
  const ui = SpreadsheetApp.getUi();
  ui.showModalDialog(HtmlService.createHtmlOutput(html).setWidth(300).setHeight(250), 'API使用状況');
}

/**
 * キャッシュをクリア
 */
function clearPriceCache() {
  const cache = CacheService.getScriptCache();
  const symbols = Object.keys(CONFIG.TARGET_ALLOCATION).filter(s => s !== 'CASH');
  
  symbols.forEach(symbol => {
    cache.remove(`price_${symbol}`);
  });
  
  SpreadsheetApp.getActiveSpreadsheet().toast('価格キャッシュをクリアしました', '完了', 3);
}

/**
 * 心理チェックシートを開く
 */
function openPsychologySheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.PSYCHOLOGY);
  
  if (sheet) {
    ss.setActiveSheet(sheet);
  } else {
    SpreadsheetApp.getActiveSpreadsheet().toast('心理管理シートが見つかりません', 'エラー', 5);
  }
}

/**
 * 初期設定：スプレッドシートの構造を作成（拡張版）
 */
function initializeSpreadsheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 設定シートを最初に作成
  createSettingsSheet(ss);
  
  // 基本シートの作成
  createPortfolioSheet(ss);
  createAlertSheet(ss);
  createHistorySheet(ss);
  createAccumulationSheet(ss);
  createRebalanceSheet(ss);
  createPriceHistorySheet(ss);
  
  // タートル流・ZONE対応シートの作成
  createBreakoutSheet(ss);
  createPsychologySheet(ss);
  createOCOSheet(ss);
  createEdgeAnalysisSheet(ss);
  
  // 一時シートを作成（非表示）
  let tempSheet = ss.getSheetByName('_TEMP');
  if (!tempSheet) {
    tempSheet = ss.insertSheet('_TEMP');
    tempSheet.hideSheet();
  }
  
  // トリガーの設定
  setupTriggersWithBreakout();
  
  // 初期心理チェックリストの作成
  initializePsychologyChecklist(ss);
  
  // API使用量の初期化
  const props = PropertiesService.getScriptProperties();
  const monthKey = new Date().toISOString().slice(0, 7);
  if (!props.getProperty(`coingecko_usage_${monthKey}`)) {
    props.setProperty(`coingecko_usage_${monthKey}`, '0');
  }
  
  SpreadsheetApp.getActiveSpreadsheet().toast('初期設定が完了しました（設定可能版）', '成功', 5);
}

/**
 * 定期実行：価格更新とアラートチェック（ブレイクアウト通知強化版）
 */
function scheduledUpdateWithBreakout() {
  try {
    console.log('定期更新を開始します...');
    
    // API使用量をチェック
    const apiUsage = getAPIUsageReport();
    if (apiUsage.percentage >= 90) {
      console.log('API制限に達しているため、無料APIのみ使用');
      manualPriceUpdate(); // 無料APIのみで更新
    } else {
      // 通常の価格更新
      updateAllPrices();
    }
    
    // ブレイクアウトをチェック（通知を確実に送信）
    console.log('ブレイクアウトシグナルをチェックしています...');
    checkBreakoutSignals();
    
    // アラートをチェック
    checkPriceAlerts();
    
    // OCO注文をチェック
    checkOCOOrders();
    
    // 積立状況をチェック（3ヶ月ごとなので頻度は問題なし）
    checkAccumulationStatus();
    
    // ボラティリティをチェック（計算のみでAPI不要）
    checkVolatilityTarget();
    
    // リバランスの必要性をチェック
    checkRebalanceNeeded();
    
    // エッジ分析を更新（内部データのみ）
    updateEdgeAnalysis();
    
    // ログ記録
    logUpdate(`定期更新完了（API使用量: ${apiUsage.percentage}%）`);
    
  } catch (error) {
    console.error('定期更新エラー:', error);
    sendErrorNotification(error);
  }
}

// ===== 価格取得関数（API制限対応版） =====

/**
 * 価格取得の優先順位システム（修正版）
 */
const PRICE_SOURCES = {
  'BTC': ['GOOGLEFINANCE', 'BINANCE', 'CRYPTOCOMPARE', 'COINCAP', 'COINGECKO'],
  'ETH': ['GOOGLEFINANCE', 'BINANCE', 'CRYPTOCOMPARE', 'COINCAP', 'COINGECKO'],
  'SOL': ['BINANCE', 'CRYPTOCOMPARE', 'COINCAP', 'COINGECKO'],
  'NEAR': ['BINANCE', 'KRAKEN', 'CRYPTOCOMPARE', 'COINCAP', 'COINGECKO'],
  'ADA': ['BINANCE', 'CRYPTOCOMPARE', 'COINCAP', 'COINGECKO'],
  'AAVE': ['BINANCE', 'CRYPTOCOMPARE', 'COINCAP', 'COINGECKO'],
  'HBAR': ['BINANCE', 'KRAKEN', 'CRYPTOCOMPARE', 'COINCAP', 'COINGECKO'],
  'GRT': ['BINANCE', 'CRYPTOCOMPARE', 'COINCAP', 'COINGECKO'],
  'ALGO': ['BINANCE', 'CRYPTOCOMPARE', 'COINCAP', 'COINGECKO']
};

/**
 * 暗号資産の価格を取得（キャッシュ対応・修正版）
 */
function getCryptoPrice(symbol) {
  // キャッシュをチェック
  const cachedPrice = getCachedPrice(symbol);
  if (cachedPrice) return cachedPrice;
  
  // 優先順位に従って価格を取得
  const sources = PRICE_SOURCES[symbol] || ['BINANCE', 'CRYPTOCOMPARE', 'COINCAP', 'COINGECKO'];
  let price = null;
  let successSource = null;
  
  for (const source of sources) {
    try {
      switch(source) {
        case 'GOOGLEFINANCE':
          price = getGoogleFinancePrice(symbol);
          if (price) successSource = 'GoogleFinance';
          break;
        case 'BINANCE':
          price = getBinancePrice(symbol);
          if (price) successSource = 'Binance';
          break;
        case 'CRYPTOCOMPARE':
          price = getCryptoComparePrice(symbol);
          if (price) successSource = 'CryptoCompare';
          break;
        case 'COINCAP':
          price = getCoinCapPrice(symbol);
          if (price) successSource = 'CoinCap';
          break;
        case 'COINGECKO':
          price = getCoinGeckoPrice(symbol);
          if (price) successSource = 'CoinGecko';
          break;
        case 'KRAKEN':
          price = getKrakenPrice(symbol);
          if (price) successSource = 'Kraken';
          break;
      }
      
      if (price && price > 0) {
        console.log(`${symbol}: ${price} from ${successSource}`);
        setCachedPrice(symbol, price);
        return price;
      }
    } catch (e) {
      console.error(`Error getting ${symbol} price from ${source}:`, e);
    }
  }
  
  // すべて失敗した場合はバックアップ価格
  console.warn(`All APIs failed for ${symbol}, using backup price`);
  return getBackupPrice(symbol);
}

/**
 * Google Finance APIから価格取得（BTC, ETHのみ）
 */
function getGoogleFinancePrice(symbol) {
  try {
    // Google FinanceはBTCとETHのみサポート
    if (!['BTC', 'ETH'].includes(symbol)) {
      return null;
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let tempSheet = ss.getSheetByName('_TEMP');
    
    if (!tempSheet) {
      tempSheet = ss.insertSheet('_TEMP');
      tempSheet.hideSheet();
    }
    
    // 既存の値をクリア
    tempSheet.getRange('A1:B1').clearContent();
    
    // Google Financeの通貨ペア形式
    const currencyPair = symbol === 'BTC' ? 'BTCUSD' : 'ETHUSD';
    const formula = `=GOOGLEFINANCE("${currencyPair}")`;
    
    tempSheet.getRange('A1').setFormula(formula);
    SpreadsheetApp.flush(); // 強制的に再計算
    Utilities.sleep(2000); // 値が更新されるまで待機
    
    const price = tempSheet.getRange('A1').getValue();
    
    // クリーンアップ
    tempSheet.getRange('A1').clearContent();
    
    if (price && !isNaN(price) && price > 0) {
      console.log(`Google Finance価格取得成功: ${symbol} = ${price}`);
      return price;
    }
  } catch (e) {
    console.error(`Google Finance価格取得エラー (${symbol}):`, e);
  }
  return null;
}

/**
 * Binance APIから価格取得（修正版）
 */
function getBinancePrice(symbol) {
  try {
    // USDTペアとBUSDペアの両方を試す
    const pairs = ['USDT', 'BUSD'];
    
    for (const pair of pairs) {
      const binanceSymbol = symbol + pair;
      const url = `https://api.binance.com/api/v3/ticker/price?symbol=${binanceSymbol}`;
      const response = UrlFetchApp.fetch(url, {muteHttpExceptions: true});
      
      if (response.getResponseCode() === 200) {
        const data = JSON.parse(response.getContentText());
        const price = parseFloat(data.price);
        console.log(`Binance価格取得成功 (${binanceSymbol}): ${price}`);
        return price;
      }
    }
    
    console.log(`Binance: ${symbol}の価格が見つかりません`);
  } catch (e) {
    console.error(`Binance価格取得エラー (${symbol}):`, e);
  }
  return null;
}

/**
 * CoinGecko APIから価格取得（月間制限対応）
 */
function getCoinGeckoPrice(symbol) {
  // API使用回数をチェック
  const usage = getCoinGeckoUsage();
  if (usage >= 9000) { // 月間10,000の90%で制限
    console.log('CoinGecko API制限に近づいています');
    return null;
  }
  
  const coinIds = {
    'BTC': 'bitcoin',
    'ETH': 'ethereum',
    'SOL': 'solana',
    'NEAR': 'near',
    'ADA': 'cardano',
    'AAVE': 'aave',
    'HBAR': 'hedera',
    'GRT': 'the-graph',
    'ALGO': 'algorand'
  };
  
  const coinId = coinIds[symbol];
  if (!coinId) return null;
  
  try {
    const url = `https://api.coingecko.com/api/v3/simple/price?ids=${coinId}&vs_currencies=usd`;
    const options = {
      'headers': {
        'x-cg-demo-api-key': CONFIG.COINGECKO_API_KEY
      },
      'muteHttpExceptions': true
    };
    
    const response = UrlFetchApp.fetch(url, options);
    
    if (response.getResponseCode() === 200) {
      incrementCoinGeckoUsage();
      const data = JSON.parse(response.getContentText());
      return data[coinId].usd;
    }
  } catch (e) {
    console.error(`CoinGecko価格取得エラー (${symbol}):`, e);
  }
  return null;
}

// 新規追加: Kraken APIから価格取得
function getKrakenPrice(symbol) {
  const krakenPairs = {
    'BTC': 'XBTUSD',
    'ETH': 'ETHUSD',
    'SOL': 'SOLUSD',
    'NEAR': 'NEARUSD',
    'HBAR': 'HBARUSD'
  };
  
  const pair = krakenPairs[symbol];
  if (!pair) return null;
  
  try {
    const url = `https://api.kraken.com/0/public/Ticker?pair=${pair}`;
    const response = UrlFetchApp.fetch(url, {muteHttpExceptions: true});
    
    if (response.getResponseCode() === 200) {
      const data = JSON.parse(response.getContentText());
      if (data.result) {
        const resultKey = Object.keys(data.result)[0];
        if (resultKey && data.result[resultKey]) {
          const price = parseFloat(data.result[resultKey].c[0]);
          if (price > 0) {
            console.log(`Kraken価格取得成功 (${symbol}): ${price}`);
            return price;
          }
        }
      }
    }
  } catch (e) {
    console.error(`Kraken価格取得エラー (${symbol}): ${e.message}`);
  }
  return null;
}

/**
 * CryptoCompare APIから価格取得（専用関数）
 */
function getCryptoComparePrice(symbol) {
  try {
    const url = `https://min-api.cryptocompare.com/data/price?fsym=${symbol}&tsyms=USD`;
    const response = UrlFetchApp.fetch(url, {muteHttpExceptions: true});
    
    if (response.getResponseCode() === 200) {
      const data = JSON.parse(response.getContentText());
      if (data.USD && data.USD > 0) {
        return data.USD;
      }
    }
  } catch (e) {
    console.error(`CryptoCompare価格取得エラー (${symbol}):`, e);
  }
  return null;
}

/**
 * CoinCap APIから価格取得（専用関数）
 */
function getCoinCapPrice(symbol) {
  try {
    const coinCapIds = {
      'BTC': 'bitcoin',
      'ETH': 'ethereum',
      'SOL': 'solana',
      'NEAR': 'near-protocol',
      'ADA': 'cardano',
      'AAVE': 'aave',
      'HBAR': 'hedera-hashgraph',
      'GRT': 'the-graph',
      'ALGO': 'algorand'
    };
    
    const assetId = coinCapIds[symbol];
    if (assetId) {
      const url = `https://api.coincap.io/v2/assets/${assetId}`;
      const response = UrlFetchApp.fetch(url, {muteHttpExceptions: true});
      
      if (response.getResponseCode() === 200) {
        const data = JSON.parse(response.getContentText());
        if (data.data && data.data.priceUsd) {
          return parseFloat(data.data.priceUsd);
        }
      }
    }
  } catch (e) {
    console.error(`CoinCap価格取得エラー (${symbol}):`, e);
  }
  return null;
}

// ===== キャッシュ機能 =====

/**
 * キャッシュから価格を取得
 */
function getCachedPrice(symbol) {
  if (!CONFIG.USE_CACHE) return null;
  
  const cache = CacheService.getScriptCache();
  const cachedData = cache.get(`price_${symbol}`);
  
  if (cachedData) {
    const data = JSON.parse(cachedData);
    const now = new Date().getTime();
    const cacheAge = (now - data.timestamp) / 1000 / 60; // 分
    
    if (cacheAge < CONFIG.CACHE_DURATION) {
      console.log(`キャッシュから価格取得: ${symbol} = ${data.price}`);
      return data.price;
    }
  }
  
  return null;
}

/**
 * 価格をキャッシュに保存
 */
function setCachedPrice(symbol, price) {
  if (!CONFIG.USE_CACHE) return;
  
  const cache = CacheService.getScriptCache();
  const data = {
    price: price,
    timestamp: new Date().getTime()
  };
  
  cache.put(`price_${symbol}`, JSON.stringify(data), CONFIG.CACHE_DURATION * 60);
}

// ===== API使用量管理 =====

/**
 * CoinGecko API使用回数を取得
 */
function getCoinGeckoUsage() {
  const props = PropertiesService.getScriptProperties();
  const monthKey = new Date().toISOString().slice(0, 7); // YYYY-MM
  const usage = props.getProperty(`coingecko_usage_${monthKey}`);
  return usage ? parseInt(usage) : 0;
}

/**
 * CoinGecko API使用回数をインクリメント
 */
function incrementCoinGeckoUsage() {
  const props = PropertiesService.getScriptProperties();
  const monthKey = new Date().toISOString().slice(0, 7);
  const currentUsage = getCoinGeckoUsage();
  props.setProperty(`coingecko_usage_${monthKey}`, String(currentUsage + 1));
}

/**
 * API使用状況レポート
 */
function getAPIUsageReport() {
  const usage = getCoinGeckoUsage();
  const percentage = (usage / 10000 * 100).toFixed(2);
  
  return {
    used: usage,
    limit: 10000,
    percentage: percentage,
    remaining: 10000 - usage
  };
}

/**
 * バックアップ価格取得（簡易版）
 */
function getBackupPrice(symbol) {
  // 2024年の概算価格（実際の使用時は更新してください）
  const fallbackPrices = {
    'BTC': 95000,
    'ETH': 3500,
    'SOL': 200,
    'NEAR': 7,
    'ADA': 1.0,
    'AAVE': 350,
    'HBAR': 0.3,
    'GRT': 0.4,
    'ALGO': 0.35
  };
  console.warn(`Using backup price for ${symbol}: ${fallbackPrices[symbol]}`);
  return fallbackPrices[symbol] || 0;
}

// ===== ポートフォリオ管理関数 =====

/**
 * 全銘柄の価格を更新
 */
function updateAllPrices() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.PORTFOLIO);
  
  if (!sheet) {
    console.error('ポートフォリオシートが見つかりません');
    return;
  }
  
  const symbols = Object.keys(CONFIG.TARGET_ALLOCATION).filter(s => s !== 'CASH');
  const dataRange = sheet.getRange(2, 1, symbols.length, 11);
  const data = dataRange.getValues();
  
  // 価格履歴用のデータ
  const priceHistoryData = [[new Date()]];
  
  symbols.forEach((symbol, index) => {
    const price = getCryptoPrice(symbol);
    if (price && price > 0) {
      data[index][4] = price; // E列: 現在価格
      data[index][5] = data[index][2] * price; // F列: 評価額
      data[index][6] = (price - data[index][3]) * data[index][2]; // G列: 損益
      data[index][7] = data[index][3] > 0 ? (price - data[index][3]) / data[index][3] : 0; // H列: 損益率
      
      priceHistoryData[0].push(price);
    }
  });
  
  // ポートフォリオ比率の計算
  const totalValue = data.reduce((sum, row) => sum + (row[5] || 0), 0);
  data.forEach((row, index) => {
    data[index][8] = totalValue > 0 ? row[5] / totalValue : 0; // I列: 現在比率
    data[index][10] = data[index][8] - data[index][9]; // K列: 乖離率
  });
  
  // データを更新
  dataRange.setValues(data);
  
  // 更新時刻を記録
  sheet.getRange(1, 12).setValue('最終更新: ' + new Date().toLocaleString('ja-JP'));
  
  // 合計値を更新
  updateSummary(sheet, totalValue);
  
  // 価格履歴を記録
  recordPriceHistory(priceHistoryData);
}

/**
 * サマリー情報を更新
 */
function updateSummary(sheet, totalValue) {
  const summaryRow = 15; // サマリー行の位置
  
  sheet.getRange(summaryRow, 1).setValue('合計');
  sheet.getRange(summaryRow, 6).setValue(totalValue);
  sheet.getRange(summaryRow, 6).setNumberFormat('¥#,##0');
  
  // 総投資額を計算
  const dataRange = sheet.getRange(2, 3, 10, 2); // C列とD列
  const data = dataRange.getValues();
  const totalInvestment = data.reduce((sum, row) => sum + (row[0] * row[1]), 0);
  
  sheet.getRange(summaryRow + 1, 1).setValue('総投資額');
  sheet.getRange(summaryRow + 1, 6).setValue(totalInvestment);
  sheet.getRange(summaryRow + 1, 6).setNumberFormat('¥#,##0');
  
  sheet.getRange(summaryRow + 2, 1).setValue('総損益');
  sheet.getRange(summaryRow + 2, 6).setValue(totalValue - totalInvestment);
  sheet.getRange(summaryRow + 2, 6).setNumberFormat('¥#,##0');
  
  sheet.getRange(summaryRow + 3, 1).setValue('損益率');
  sheet.getRange(summaryRow + 3, 6).setValue(totalInvestment > 0 ? (totalValue - totalInvestment) / totalInvestment : 0);
  sheet.getRange(summaryRow + 3, 6).setNumberFormat('0.00%');
}

// ===== アラート機能 =====

/**
 * 価格アラートをチェック
 */
function checkPriceAlerts() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.ALERTS);
  
  if (!sheet || sheet.getLastRow() < 2) return;
  
  const dataRange = sheet.getRange(2, 1, sheet.getLastRow() - 1, 7);
  const data = dataRange.getValues();
  const alerts = [];
  
  data.forEach((row, index) => {
    if (row[5] === 'アクティブ') { // F列: ステータス
      const symbol = row[0];
      const condition = row[1];
      const targetPrice = row[2];
      const currentPrice = getCryptoPrice(symbol);
      
      if (currentPrice && shouldTriggerAlert(condition, currentPrice, targetPrice, row[3])) {
        alerts.push({
          symbol: symbol,
          condition: condition,
          targetPrice: targetPrice,
          currentPrice: currentPrice,
          previousPrice: row[3]
        });
        
        // アラートを発動済みに更新
        data[index][5] = '発動済み';
        data[index][6] = new Date();
      }
      
      // 現在価格を更新
      data[index][3] = currentPrice;
    }
  });
  
  // データを更新
  dataRange.setValues(data);
  
  // アラートメールを送信
  if (alerts.length > 0) {
    sendAlertEmail(alerts);
  }
}

/**
 * アラート条件の判定
 */
function shouldTriggerAlert(condition, currentPrice, targetPrice, previousPrice) {
  switch (condition) {
    case '以上':
      return currentPrice >= targetPrice;
    case '以下':
      return currentPrice <= targetPrice;
    case '上昇率':
      return previousPrice > 0 && (currentPrice - previousPrice) / previousPrice >= targetPrice / 100;
    case '下落率':
      return previousPrice > 0 && (previousPrice - currentPrice) / previousPrice >= targetPrice / 100;
    default:
      return false;
  }
}

/**
 * アラートメールを送信
 */
function sendAlertEmail(alerts) {
  const subject = '【暗号資産】価格アラート通知';
  let body = '以下の価格アラートが発動しました：\n\n';
  
  alerts.forEach(alert => {
    body += `━━━━━━━━━━━━━━━━━━━━\n`;
    body += `銘柄: ${alert.symbol}\n`;
    body += `条件: ${alert.condition}\n`;
    body += `目標: $${alert.targetPrice.toLocaleString()}\n`;
    body += `現在: $${alert.currentPrice.toLocaleString()}\n`;
    if (alert.previousPrice > 0) {
      const change = ((alert.currentPrice - alert.previousPrice) / alert.previousPrice * 100).toFixed(2);
      body += `変動: ${change > 0 ? '+' : ''}${change}%\n`;
    }
  });
  
  body += `\n━━━━━━━━━━━━━━━━━━━━\n`;
  body += `送信時刻: ${new Date().toLocaleString('ja-JP')}\n`;
  body += `\nスプレッドシート: ${SpreadsheetApp.getActiveSpreadsheet().getUrl()}`;
  
  MailApp.sendEmail({
    to: CONFIG.EMAIL_RECIPIENT,
    subject: subject,
    body: body
  });
}

// ===== ブレイクアウト戦略関数（タートル流） =====

/**
 * ブレイクアウトシグナルをチェック（通知強化版）
 */
function checkBreakoutSignals() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const priceHistorySheet = ss.getSheetByName(CONFIG.SHEETS.PRICE_HISTORY);
  const breakoutSheet = ss.getSheetByName(CONFIG.SHEETS.BREAKOUT);
  const signals = [];
  
  if (!priceHistorySheet || !breakoutSheet) {
    console.log('必要なシートが見つかりません');
    return;
  }
  
  const symbols = Object.keys(CONFIG.TARGET_ALLOCATION).filter(s => s !== 'CASH');
  
  symbols.forEach((symbol, index) => {
    const layer = getSymbolLayer(symbol);
    const period = layer === 'BASE' ? CONFIG.BREAKOUT_PERIODS.BASE : CONFIG.BREAKOUT_PERIODS.GROWTH;
    
    // 価格履歴を取得
    const priceData = getHistoricalPrices(priceHistorySheet, symbol, period);
    
    if (priceData.length >= period) {
      const currentPrice = getCryptoPrice(symbol);
      const highestHigh = Math.max(...priceData);
      const lowestLow = Math.min(...priceData);
      
      // ブレイクアウト判定
      if (currentPrice > highestHigh) {
        console.log(`${symbol}: BUYシグナル検出 - 現在価格 ${currentPrice} > ${period}日高値 ${highestHigh}`);
        signals.push({
          symbol: symbol,
          type: 'BUY',
          signal: `${period}日高値ブレイクアウト`,
          price: currentPrice,
          previousHigh: highestHigh,
          timestamp: new Date()
        });
      } else if (currentPrice < lowestLow) {
        console.log(`${symbol}: SELLシグナル検出 - 現在価格 ${currentPrice} < ${period}日安値 ${lowestLow}`);
        signals.push({
          symbol: symbol,
          type: 'SELL',
          signal: `${period}日安値ブレイクダウン`,
          price: currentPrice,
          previousLow: lowestLow,
          timestamp: new Date()
        });
      }
    }
  });
  
  // シグナルを記録して通知
  if (signals.length > 0) {
    console.log(`${signals.length}個のブレイクアウトシグナルを検出しました`);
    recordBreakoutSignals(breakoutSheet, signals);
    
    // 即座に通知を送信
    try {
      sendBreakoutNotification(signals);
      console.log('ブレイクアウト通知を送信しました');
    } catch (e) {
      console.error('ブレイクアウト通知の送信に失敗しました:', e);
    }
  } else {
    console.log('ブレイクアウトシグナルは検出されませんでした');
  }
}

/**
 * 銘柄のレイヤーを判定
 */
function getSymbolLayer(symbol) {
  if (['BTC', 'ETH', 'SOL'].includes(symbol)) return 'BASE';
  if (['ALGO'].includes(symbol)) return 'SATELLITE';
  return 'GROWTH';
}

/**
 * 過去の価格データを取得
 */
function getHistoricalPrices(sheet, symbol, period) {
  const lastRow = sheet.getLastRow();
  if (lastRow < period + 1) return [];
  
  const columnIndex = getSymbolColumnIndex(sheet, symbol);
  if (columnIndex === -1) return [];
  
  const range = sheet.getRange(lastRow - period + 1, columnIndex, period);
  return range.getValues().flat().filter(v => v && v > 0);
}

/**
 * ブレイクアウトシグナルを記録
 */
function recordBreakoutSignals(sheet, signals) {
  const lastRow = sheet.getLastRow();
  const data = signals.map(signal => [
    signal.timestamp,
    signal.symbol,
    signal.type,
    signal.signal,
    signal.price,
    signal.previousHigh || signal.previousLow,
    '', // エッジ比率（後で計算）
    'アクティブ'
  ]);
  
  sheet.getRange(lastRow + 1, 1, data.length, 8).setValues(data);
}

/**
 * ブレイクアウト通知（改良版）
 */
function sendBreakoutNotification(signals) {
  const subject = '【暗号資産】🚨 ブレイクアウトシグナル検出';
  let body = '重要：以下のブレイクアウトシグナルが検出されました！\n\n';
  
  // BUYシグナルとSELLシグナルを分けて表示
  const buySignals = signals.filter(s => s.type === 'BUY');
  const sellSignals = signals.filter(s => s.type === 'SELL');
  
  if (buySignals.length > 0) {
    body += '【📈 BUYシグナル】\n';
    buySignals.forEach(signal => {
      body += `━━━━━━━━━━━━━━━━━━━━\n`;
      body += `銘柄: ${signal.symbol}\n`;
      body += `シグナル: ${signal.signal}\n`;
      body += `現在価格: $${signal.price.toLocaleString()}\n`;
      body += `前回高値: $${signal.previousHigh.toLocaleString()}\n`;
      body += `上昇率: +${((signal.price - signal.previousHigh) / signal.previousHigh * 100).toFixed(2)}%\n`;
      
      // 推奨アクション
      const layer = getSymbolLayer(signal.symbol);
      const limits = CONFIG.LIMIT_ORDER_LEVELS[layer];
      body += `\n💡 推奨階段指値: ${limits.map(l => `${(l * 100).toFixed(0)}%`).join(', ')}\n`;
    });
    body += '\n';
  }
  
  if (sellSignals.length > 0) {
    body += '【📉 SELLシグナル】\n';
    sellSignals.forEach(signal => {
      body += `━━━━━━━━━━━━━━━━━━━━\n`;
      body += `銘柄: ${signal.symbol}\n`;
      body += `シグナル: ${signal.signal}\n`;
      body += `現在価格: $${signal.price.toLocaleString()}\n`;
      body += `前回安値: $${signal.previousLow.toLocaleString()}\n`;
      body += `下落率: ${((signal.price - signal.previousLow) / signal.previousLow * 100).toFixed(2)}%\n`;
    });
    body += '\n';
  }
  
  body += `━━━━━━━━━━━━━━━━━━━━\n`;
  body += `検出時刻: ${new Date().toLocaleString('ja-JP')}\n\n`;
  body += `⚠️【重要】取引前に必ず心理チェックリストを確認してください。\n`;
  body += `詳細分析: ${SpreadsheetApp.getActiveSpreadsheet().getUrl()}`;
  
  // メール送信
  MailApp.sendEmail({
    to: CONFIG.EMAIL_RECIPIENT,
    subject: subject,
    body: body
  });
  
  // スプレッドシートにも通知
  SpreadsheetApp.getActiveSpreadsheet().toast(
    `${signals.length}個のブレイクアウトシグナルを検出しました。メールで通知を送信しました。`,
    'ブレイクアウト通知',
    10
  );
}

/**
 * 手動でブレイクアウトをチェック
 */
function manualBreakoutCheck() {
  console.log('手動ブレイクアウトチェックを開始します...');
  checkBreakoutSignals();
  SpreadsheetApp.getActiveSpreadsheet().toast('ブレイクアウトチェックが完了しました', '完了', 3);
}

// ===== 心理管理関数（ZONE） - 修正版 =====

/**
 * 心理チェックリストを初期化（修正版）
 */
function initializePsychologyChecklist(ss) {
  const sheet = ss.getSheetByName(CONFIG.SHEETS.PSYCHOLOGY);
  if (!sheet) return;
  
  // 信念リストを作成（行3から開始、行2はヘッダー）
  const beliefsData = CONFIG.PSYCHOLOGY_BELIEFS.map((belief, index) => [
    `信念${index + 1}`,
    belief,
    false, // チェック状態
    new Date()
  ]);
  
  // ★修正点：行3から開始（元は行2から）
  sheet.getRange(3, 1, beliefsData.length, 4).setValues(beliefsData);
  
  // バイアスチェックリストを作成
  const biasStartRow = 3 + beliefsData.length + 2; // ★修正点：3 + length + 2
  sheet.getRange(biasStartRow, 1).setValue('心理バイアスチェック');
  sheet.getRange(biasStartRow, 1).setFontWeight('bold');
  
  // バイアスヘッダー
  const biasHeaders = [['バイアスの種類', '最近の事例', '発生回数', '最終発生日']];
  sheet.getRange(biasStartRow + 1, 1, 1, 4).setValues(biasHeaders);
  sheet.getRange(biasStartRow + 1, 1, 1, 4).setFontWeight('bold');
  sheet.getRange(biasStartRow + 1, 1, 1, 4).setBackground('#f3f3f3');
  
  const biasData = CONFIG.BIAS_CHECKLIST.map(bias => [
    bias,
    '', // 最近の事例
    0,  // 発生回数
    new Date()
  ]);
  
  // ★修正点：バイアスデータも正しい行から開始
  sheet.getRange(biasStartRow + 2, 1, biasData.length, 4).setValues(biasData);
}

/**
 * 取引前の心理チェック（修正版）
 */
function performPsychologyCheck() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.PSYCHOLOGY);
  
  if (!sheet) return { passed: true };
  
  // 信念の確認状態をチェック（行3から開始）
  const beliefsRange = sheet.getRange(3, 3, CONFIG.PSYCHOLOGY_BELIEFS.length, 1);
  const beliefsChecked = beliefsRange.getValues().flat();
  const allBeliefsChecked = beliefsChecked.every(checked => checked === true);
  
  if (!allBeliefsChecked) {
    return {
      passed: false,
      message: '取引前に全ての信念を確認してください'
    };
  }
  
  // 最近のバイアス発生をチェック
  const biasStartRow = 3 + CONFIG.PSYCHOLOGY_BELIEFS.length + 4; // ★修正点
  const biasRange = sheet.getRange(biasStartRow, 3, CONFIG.BIAS_CHECKLIST.length, 1);
  const biasCounts = biasRange.getValues().flat();
  const highBiasCount = biasCounts.some(count => count > 3);
  
  if (highBiasCount) {
    return {
      passed: false,
      message: '心理バイアスの発生が多いため、冷静になってから取引してください'
    };
  }
  
  return { passed: true };
}

/**
 * 心理チェックリストをリセット
 */
function resetPsychologyChecklist() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.PSYCHOLOGY);
  
  if (!sheet) return;
  
  // 信念チェックボックスをリセット（行3から）
  const checkboxRange = sheet.getRange(3, 3, CONFIG.PSYCHOLOGY_BELIEFS.length, 1);
  checkboxRange.setValue(false);
  
  console.log('心理チェックリストをリセットしました');
}

// ===== OCO注文管理（拡張版） =====

/**
 * OCO注文を自動設定
 */
function setOCOOrder(symbol, entryPrice, quantity) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.OCO_ORDERS);
  
  if (!sheet) return;
  
  const ocoSettings = CONFIG.OCO_SETTINGS[symbol];
  if (!ocoSettings) return;
  
  const takeProfitPrice = entryPrice * (1 + ocoSettings.tp);
  const stopLossPrice = entryPrice * (1 + ocoSettings.sl);
  
  const orderData = [[
    new Date(),
    symbol,
    quantity,
    entryPrice,
    takeProfitPrice,
    stopLossPrice,
    'アクティブ',
    '', // 約定日時
    '', // 約定価格
    ''  // 損益
  ]];
  
  const lastRow = sheet.getLastRow();
  sheet.getRange(lastRow + 1, 1, 1, 10).setValues(orderData);
}

/**
 * OCO注文をチェック
 */
function checkOCOOrders() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.OCO_ORDERS);
  
  if (!sheet || sheet.getLastRow() < 2) return;
  
  const dataRange = sheet.getRange(2, 1, sheet.getLastRow() - 1, 10);
  const data = dataRange.getValues();
  const executedOrders = [];
  
  data.forEach((row, index) => {
    if (row[6] === 'アクティブ') {
      const symbol = row[1];
      const currentPrice = getCryptoPrice(symbol);
      const entryPrice = row[3];
      const tpPrice = row[4];
      const slPrice = row[5];
      
      if (currentPrice >= tpPrice) {
        // Take Profit実行
        data[index][6] = 'TP約定';
        data[index][7] = new Date();
        data[index][8] = currentPrice;
        data[index][9] = (currentPrice - entryPrice) * row[2];
        
        executedOrders.push({
          type: 'Take Profit',
          symbol: symbol,
          price: currentPrice,
          profit: data[index][9]
        });
        
      } else if (currentPrice <= slPrice) {
        // Stop Loss実行
        data[index][6] = 'SL約定';
        data[index][7] = new Date();
        data[index][8] = currentPrice;
        data[index][9] = (currentPrice - entryPrice) * row[2];
        
        executedOrders.push({
          type: 'Stop Loss',
          symbol: symbol,
          price: currentPrice,
          loss: data[index][9]
        });
      }
    }
  });
  
  // データを更新
  dataRange.setValues(data);
  
  // 約定通知
  if (executedOrders.length > 0) {
    sendOCOExecutionNotification(executedOrders);
    updatePsychologyAfterTrade(executedOrders);
  }
}

// ===== 積立管理機能 =====

/**
 * 積立状況をチェック
 */
function checkAccumulationStatus() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.ACCUMULATION);
  
  if (!sheet) return;
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  
  // 最新の積立状況を取得
  const currentAmount = sheet.getRange(lastRow, 3).getValue(); // C列: 累積金額
  const startDate = sheet.getRange(lastRow, 1).getValue(); // A列: 開始日
  
  // 3ヶ月経過したかチェック
  const monthsPassed = getMonthsDifference(startDate, new Date());
  
  if (monthsPassed >= CONFIG.ACCUMULATION_MONTHS) {
    // ブレイクアウトシグナルを確認
    const breakoutSheet = ss.getSheetByName(CONFIG.SHEETS.BREAKOUT);
    const activeSignals = getActiveBreakoutSignals(breakoutSheet);
    
    // 投資実行を提案
    sendInvestmentNotificationWithSignals(currentAmount, activeSignals);
    
    // 新しい積立期間を開始
    startNewAccumulationPeriod(sheet);
  }
}

/**
 * 投資実行の通知
 */
function sendInvestmentNotification(amount) {
  const subject = '【暗号資産】投資実行タイミングのお知らせ';
  let body = `3ヶ月の積立期間が完了しました。\n\n`;
  body += `積立金額: ¥${amount.toLocaleString()}\n\n`;
  body += `投資配分案:\n`;
  
  // 配分計算
  Object.entries(CONFIG.TARGET_ALLOCATION).forEach(([symbol, ratio]) => {
    if (symbol !== 'CASH') {
      const investAmount = amount * ratio;
      body += `${symbol}: ¥${Math.floor(investAmount).toLocaleString()} (${(ratio * 100).toFixed(0)}%)\n`;
    }
  });
  
  body += `\n現金保有: ¥${Math.floor(amount * CONFIG.TARGET_ALLOCATION.CASH).toLocaleString()} (6%)\n`;
  body += `\n階段指値の設定をお忘れなく！\n`;
  body += `推奨指値レベル: -4%, -8%, -12%\n`;
  
  MailApp.sendEmail({
    to: CONFIG.EMAIL_RECIPIENT,
    subject: subject,
    body: body
  });
}

// ===== リバランス機能 =====

/**
 * リバランスの必要性をチェック
 */
function checkRebalanceNeeded() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const portfolioSheet = ss.getSheetByName(CONFIG.SHEETS.PORTFOLIO);
  const rebalanceSheet = ss.getSheetByName(CONFIG.SHEETS.REBALANCE);
  
  if (!portfolioSheet || !rebalanceSheet) return;
  
  // リバランスが必要な銘柄を特定
  const rebalanceItems = [];
  const dataRange = portfolioSheet.getRange(2, 1, 10, 11);
  const data = dataRange.getValues();
  
  data.forEach(row => {
    const symbol = row[0];
    const deviation = Math.abs(row[10]); // K列: 乖離率
    
    if (deviation > CONFIG.REBALANCE_THRESHOLD && symbol) {
      rebalanceItems.push({
        symbol: symbol,
        currentRatio: row[8],
        targetRatio: row[9],
        deviation: row[10],
        currentValue: row[5]
      });
    }
  });
  
  if (rebalanceItems.length > 0) {
    updateRebalanceSheet(rebalanceSheet, rebalanceItems);
    
    // 週次でのみ通知（月曜日）
    if (new Date().getDay() === 1) {
      sendRebalanceNotification(rebalanceItems);
    }
  }
}

/**
 * リバランスシートを更新
 */
function updateRebalanceSheet(sheet, items) {
  // シートをクリア
  sheet.clear();
  
  // ヘッダーを設定
  const headers = [['銘柄', '現在比率', '目標比率', '乖離', '推奨アクション', '金額']];
  sheet.getRange(1, 1, 1, 6).setValues(headers);
  sheet.getRange(1, 1, 1, 6).setFontWeight('bold');
  
  // データを設定
  const data = items.map(item => {
    const action = item.deviation > 0 ? '売却' : '購入';
    const totalValue = item.currentValue / item.currentRatio;
    const amount = Math.abs(item.deviation * totalValue);
    
    return [
      item.symbol,
      (item.currentRatio * 100).toFixed(2) + '%',
      (item.targetRatio * 100).toFixed(2) + '%',
      (item.deviation * 100).toFixed(2) + '%',
      action,
      Math.floor(amount)
    ];
  });
  
  if (data.length > 0) {
    sheet.getRange(2, 1, data.length, 6).setValues(data);
    sheet.getRange(2, 6, data.length, 1).setNumberFormat('¥#,##0');
  }
  
  // 更新時刻
  sheet.getRange(data.length + 3, 1).setValue('最終更新: ' + new Date().toLocaleString('ja-JP'));
}

// ===== エッジ分析（タートル流） =====

/**
 * エッジ分析を更新
 */
function updateEdgeAnalysis() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const historySheet = ss.getSheetByName(CONFIG.SHEETS.HISTORY);
  const edgeSheet = ss.getSheetByName(CONFIG.SHEETS.EDGE_ANALYSIS);
  
  if (!historySheet || !edgeSheet) return;
  
  // 取引履歴からエッジを計算
  const symbols = Object.keys(CONFIG.TARGET_ALLOCATION).filter(s => s !== 'CASH');
  const edgeData = [];
  
  symbols.forEach(symbol => {
    const trades = getTradeHistory(historySheet, symbol);
    if (trades.length > 0) {
      const edge = calculateEdgeRatio(trades);
      edgeData.push([
        symbol,
        trades.length,
        edge.winRate,
        edge.avgWin,
        edge.avgLoss,
        edge.edgeRatio,
        edge.expectancy,
        new Date()
      ]);
    }
  });
  
  // エッジ分析シートを更新
  if (edgeData.length > 0) {
    edgeSheet.clear();
    const headers = [['銘柄', '取引数', '勝率', '平均利益', '平均損失', 'エッジ比率', '期待値', '更新日時']];
    edgeSheet.getRange(1, 1, 1, 8).setValues(headers);
    edgeSheet.getRange(2, 1, edgeData.length, 8).setValues(edgeData);
  }
}

/**
 * エッジ比率を計算（MFE/MAE）
 */
function calculateEdgeRatio(trades) {
  let wins = 0;
  let totalWin = 0;
  let totalLoss = 0;
  
  trades.forEach(trade => {
    const profit = trade.profit;
    if (profit > 0) {
      wins++;
      totalWin += profit;
    } else {
      totalLoss += Math.abs(profit);
    }
  });
  
  const winRate = trades.length > 0 ? wins / trades.length : 0;
  const avgWin = wins > 0 ? totalWin / wins : 0;
  const avgLoss = (trades.length - wins) > 0 ? totalLoss / (trades.length - wins) : 0;
  const edgeRatio = avgLoss > 0 ? avgWin / avgLoss : 0;
  const expectancy = (winRate * avgWin) - ((1 - winRate) * avgLoss);
  
  return {
    winRate: winRate,
    avgWin: avgWin,
    avgLoss: avgLoss,
    edgeRatio: edgeRatio,
    expectancy: expectancy
  };
}

// ===== ボラティリティターゲティング =====

/**
 * ボラティリティをチェックして調整
 */
function checkVolatilityTarget() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const priceHistorySheet = ss.getSheetByName(CONFIG.SHEETS.PRICE_HISTORY);
  const portfolioSheet = ss.getSheetByName(CONFIG.SHEETS.PORTFOLIO);
  
  if (!priceHistorySheet || !portfolioSheet) return;
  
  const volatilities = {};
  const symbols = Object.keys(CONFIG.TARGET_ALLOCATION).filter(s => s !== 'CASH');
  
  // 各銘柄の30日ボラティリティを計算
  symbols.forEach(symbol => {
    const vol = calculate30DayVolatility(priceHistorySheet, symbol);
    volatilities[symbol] = vol;
  });
  
  // ポートフォリオ全体のボラティリティを計算
  const portfolioVol = calculatePortfolioVolatility(portfolioSheet, volatilities);
  
  // 閾値を超えた場合の調整
  if (portfolioVol > CONFIG.VOLATILITY_THRESHOLD) {
    const adjustment = proposeVolatilityAdjustment(portfolioSheet, volatilities);
    sendVolatilityAlertNotification(portfolioVol, adjustment);
  }
}

/**
 * 30日ボラティリティを計算
 */
function calculate30DayVolatility(sheet, symbol) {
  const prices = getHistoricalPrices(sheet, symbol, 30);
  if (prices.length < 2) return 0;
  
  // 日次リターンを計算
  const returns = [];
  for (let i = 1; i < prices.length; i++) {
    returns.push((prices[i] - prices[i-1]) / prices[i-1]);
  }
  
  // 標準偏差を計算
  const mean = returns.reduce((a, b) => a + b, 0) / returns.length;
  const variance = returns.reduce((a, b) => a + Math.pow(b - mean, 2), 0) / returns.length;
  const dailyVol = Math.sqrt(variance);
  
  // 年率換算
  return dailyVol * Math.sqrt(365);
}

// ===== シート作成関数 =====

/**
 * 設定シートを作成
 */
function createSettingsSheet(ss) {
  let sheet = ss.getSheetByName(CONFIG.SHEETS.SETTINGS);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEETS.SETTINGS);
  } else {
    sheet.clear();
  }
  
  // タイトル
  sheet.getRange(1, 1).setValue('システム設定');
  sheet.getRange(1, 1).setFontWeight('bold');
  sheet.getRange(1, 1).setFontSize(14);
  
  // 設定項目
  const settings = [
    ['毎月の積立金額（円）', 1500000],
    ['通知先メールアドレス', Session.getActiveUser().getEmail()]
  ];
  
  sheet.getRange(2, 1, settings.length, 2).setValues(settings);
  
  // フォーマット設定
  sheet.getRange(2, 1, settings.length, 1).setFontWeight('bold');
  sheet.getRange(2, 1, settings.length, 1).setBackground('#f3f3f3');
  sheet.getRange(2, 2).setNumberFormat('¥#,##0');
  
  // 説明文を追加
  sheet.getRange(4, 1).setValue('説明');
  sheet.getRange(4, 1).setFontWeight('bold');
  sheet.getRange(5, 1).setValue('毎月の積立金額：');
  sheet.getRange(5, 2).setValue('毎月積み立てる金額を設定します');
  sheet.getRange(6, 1).setValue('通知先メールアドレス：');
  sheet.getRange(6, 2).setValue('アラートや通知を受け取るメールアドレスを設定します');
  
  // 列幅調整
  sheet.setColumnWidth(1, 200);
  sheet.setColumnWidth(2, 300);
  
  // 保護設定（設定値以外は編集不可）
  const protection = sheet.protect().setDescription('設定シート保護');
  protection.setUnprotectedRanges([sheet.getRange('B2:B3')]);
}

/**
 * ポートフォリオシートを作成
 */
function createPortfolioSheet(ss) {
  let sheet = ss.getSheetByName(CONFIG.SHEETS.PORTFOLIO);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEETS.PORTFOLIO);
  } else {
    sheet.clear();
  }
  
  // ヘッダー設定
  const headers = [
    ['銘柄', 'ティッカー', '保有数量', '平均取得単価', '現在価格', '評価額', 
     '損益', '損益率', '現在比率', '目標比率', '乖離率']
  ];
  sheet.getRange(1, 1, 1, 11).setValues(headers);
  sheet.getRange(1, 1, 1, 11).setFontWeight('bold');
  sheet.setFrozenRows(1);
  
  // 初期データ設定
  const symbols = Object.keys(CONFIG.TARGET_ALLOCATION).filter(s => s !== 'CASH');
  const initialData = symbols.map(symbol => [
    symbol, 
    symbol, 
    0, // 保有数量
    0, // 平均取得単価
    0, // 現在価格
    0, // 評価額
    0, // 損益
    0, // 損益率
    0, // 現在比率
    CONFIG.TARGET_ALLOCATION[symbol], // 目標比率
    0  // 乖離率
  ]);
  
  sheet.getRange(2, 1, initialData.length, 11).setValues(initialData);
  
  // フォーマット設定
  sheet.getRange(2, 3, symbols.length, 1).setNumberFormat('#,##0.0000'); // 数量
  sheet.getRange(2, 4, symbols.length, 2).setNumberFormat('$#,##0.00'); // 価格
  sheet.getRange(2, 6, symbols.length, 1).setNumberFormat('¥#,##0'); // 評価額
  sheet.getRange(2, 7, symbols.length, 1).setNumberFormat('¥#,##0'); // 損益
  sheet.getRange(2, 8, symbols.length, 1).setNumberFormat('0.00%'); // 損益率
  sheet.getRange(2, 9, symbols.length, 3).setNumberFormat('0.00%'); // 比率
  
  // 列幅調整
  sheet.autoResizeColumns(1, 11);
}

/**
 * アラート設定シートを作成
 */
function createAlertSheet(ss) {
  let sheet = ss.getSheetByName(CONFIG.SHEETS.ALERTS);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEETS.ALERTS);
  } else {
    sheet.clear();
  }
  
  const headers = [
    ['銘柄', '条件', '目標値', '現在価格', '作成日', 'ステータス', '発動日時']
  ];
  sheet.getRange(1, 1, 1, 7).setValues(headers);
  sheet.getRange(1, 1, 1, 7).setFontWeight('bold');
  sheet.setFrozenRows(1);
  
  // 条件のドロップダウン設定
  const conditionRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['以上', '以下', '上昇率', '下落率'], true)
    .build();
  sheet.getRange(2, 2, 100, 1).setDataValidation(conditionRule);
  
  // ステータスのドロップダウン設定
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['アクティブ', '発動済み', '無効'], true)
    .build();
  sheet.getRange(2, 6, 100, 1).setDataValidation(statusRule);
  
  sheet.autoResizeColumns(1, 7);
}

/**
 * 取引履歴シートを作成
 */
function createHistorySheet(ss) {
  let sheet = ss.getSheetByName(CONFIG.SHEETS.HISTORY);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEETS.HISTORY);
  } else {
    sheet.clear();
  }
  
  const headers = [
    ['日時', '銘柄', '売買', '数量', '価格', '手数料', '合計金額', 'メモ']
  ];
  sheet.getRange(1, 1, 1, 8).setValues(headers);
  sheet.getRange(1, 1, 1, 8).setFontWeight('bold');
  sheet.setFrozenRows(1);
  
  // 売買のドロップダウン設定
  const actionRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['買い', '売り'], true)
    .build();
  sheet.getRange(2, 3, 1000, 1).setDataValidation(actionRule);
  
  sheet.autoResizeColumns(1, 8);
}

/**
 * 積立管理シートを作成
 */
function createAccumulationSheet(ss) {
  let sheet = ss.getSheetByName(CONFIG.SHEETS.ACCUMULATION);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEETS.ACCUMULATION);
  } else {
    sheet.clear();
  }
  
  const headers = [
    ['開始日', '月', '累積金額', '投資実行', 'ステータス']
  ];
  sheet.getRange(1, 1, 1, 5).setValues(headers);
  sheet.getRange(1, 1, 1, 5).setFontWeight('bold');
  sheet.setFrozenRows(1);
  
  // 初回データ（設定から金額を取得）
  const today = new Date();
  const monthlyInvestment = getMonthlyInvestment();
  const initialData = [[today, 1, monthlyInvestment, '', '積立中']];
  sheet.getRange(2, 1, 1, 5).setValues(initialData);
  
  sheet.getRange(2, 3, 100, 1).setNumberFormat('¥#,##0');
  sheet.autoResizeColumns(1, 5);
}

/**
 * リバランス提案シートを作成
 */
function createRebalanceSheet(ss) {
  let sheet = ss.getSheetByName(CONFIG.SHEETS.REBALANCE);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEETS.REBALANCE);
  } else {
    sheet.clear();
  }
  
  const headers = [
    ['銘柄', '現在比率', '目標比率', '乖離', '推奨アクション', '金額']
  ];
  sheet.getRange(1, 1, 1, 6).setValues(headers);
  sheet.getRange(1, 1, 1, 6).setFontWeight('bold');
  sheet.setFrozenRows(1);
  
  sheet.autoResizeColumns(1, 6);
}

/**
 * 価格履歴シートを作成（拡張版）
 */
function createPriceHistorySheet(ss) {
  let sheet = ss.getSheetByName(CONFIG.SHEETS.PRICE_HISTORY);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEETS.PRICE_HISTORY);
  } else {
    sheet.clear();
  }
  
  const symbols = Object.keys(CONFIG.TARGET_ALLOCATION).filter(s => s !== 'CASH');
  const headers = [['日時', ...symbols, 'データソース', 'API使用量']];
  sheet.getRange(1, 1, 1, symbols.length + 3).setValues(headers);
  sheet.getRange(1, 1, 1, symbols.length + 3).setFontWeight('bold');
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(1);
  
  sheet.autoResizeColumns(1, symbols.length + 3);
}

/**
 * ブレイクアウトシートを作成
 */
function createBreakoutSheet(ss) {
  let sheet = ss.getSheetByName(CONFIG.SHEETS.BREAKOUT);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEETS.BREAKOUT);
  } else {
    sheet.clear();
  }
  
  const headers = [
    ['日時', '銘柄', 'タイプ', 'シグナル', '現在価格', '前回高値/安値', 'エッジ比率', 'ステータス']
  ];
  sheet.getRange(1, 1, 1, 8).setValues(headers);
  sheet.getRange(1, 1, 1, 8).setFontWeight('bold');
  sheet.setFrozenRows(1);
  
  sheet.autoResizeColumns(1, 8);
}

/**
 * 心理管理シートを作成（修正版）
 */
function createPsychologySheet(ss) {
  let sheet = ss.getSheetByName(CONFIG.SHEETS.PSYCHOLOGY);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEETS.PSYCHOLOGY);
  } else {
    sheet.clear();
  }
  
  // タイトル
  sheet.getRange(1, 1).setValue('取引前信念チェックリスト');
  sheet.getRange(1, 1).setFontWeight('bold');
  sheet.getRange(1, 1).setFontSize(12);
  
  // ヘッダー（行2）
  const beliefHeaders = [['ID', '信念', 'チェック', '最終確認']];
  sheet.getRange(2, 1, 1, 4).setValues(beliefHeaders);
  sheet.getRange(2, 1, 1, 4).setFontWeight('bold');
  sheet.getRange(2, 1, 1, 4).setBackground('#f3f3f3');
  
  // チェックボックスを設定（行3から）
  const checkboxRange = sheet.getRange(3, 3, CONFIG.PSYCHOLOGY_BELIEFS.length, 1);
  checkboxRange.insertCheckboxes();
  
  // 列幅調整
  sheet.setColumnWidth(1, 100); // ID
  sheet.setColumnWidth(2, 400); // 信念
  sheet.setColumnWidth(3, 80);  // チェック
  sheet.setColumnWidth(4, 120); // 最終確認
  
  sheet.autoResizeRows(1, 20);
}

/**
 * OCO注文シートを作成
 */
function createOCOSheet(ss) {
  let sheet = ss.getSheetByName(CONFIG.SHEETS.OCO_ORDERS);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEETS.OCO_ORDERS);
  } else {
    sheet.clear();
  }
  
  const headers = [
    ['設定日時', '銘柄', '数量', 'エントリー価格', 'TP価格', 'SL価格', 
     'ステータス', '約定日時', '約定価格', '損益']
  ];
  sheet.getRange(1, 1, 1, 10).setValues(headers);
  sheet.getRange(1, 1, 1, 10).setFontWeight('bold');
  sheet.setFrozenRows(1);
  
  // フォーマット設定
  sheet.getRange(2, 4, 1000, 3).setNumberFormat('$#,##0.00'); // 価格
  sheet.getRange(2, 10, 1000, 1).setNumberFormat('¥#,##0'); // 損益
  
  sheet.autoResizeColumns(1, 10);
}

/**
 * エッジ分析シートを作成
 */
function createEdgeAnalysisSheet(ss) {
  let sheet = ss.getSheetByName(CONFIG.SHEETS.EDGE_ANALYSIS);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEETS.EDGE_ANALYSIS);
  } else {
    sheet.clear();
  }
  
  const headers = [
    ['銘柄', '取引数', '勝率', '平均利益', '平均損失', 'エッジ比率', '期待値', '更新日時']
  ];
  sheet.getRange(1, 1, 1, 8).setValues(headers);
  sheet.getRange(1, 1, 1, 8).setFontWeight('bold');
  sheet.setFrozenRows(1);
  
  // フォーマット設定
  sheet.getRange(2, 3, 100, 1).setNumberFormat('0.00%'); // 勝率
  sheet.getRange(2, 4, 100, 3).setNumberFormat('¥#,##0'); // 金額
  sheet.getRange(2, 6, 100, 1).setNumberFormat('0.00'); // エッジ比率
  
  sheet.autoResizeColumns(1, 8);
}

// ===== ヘルパー関数 =====

/**
 * 月数の差を計算
 */
function getMonthsDifference(date1, date2) {
  const months = (date2.getFullYear() - date1.getFullYear()) * 12;
  return months + date2.getMonth() - date1.getMonth();
}

/**
 * 価格履歴を記録
 */
function recordPriceHistory(priceHistoryData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.PRICE_HISTORY);
  
  if (!sheet) return;
  
  const lastRow = sheet.getLastRow();
  sheet.getRange(lastRow + 1, 1, 1, priceHistoryData[0].length).setValues(priceHistoryData);
}

/**
 * 新しい積立期間を開始
 */
function startNewAccumulationPeriod(sheet) {
  const lastRow = sheet.getLastRow();
  const monthlyInvestment = getMonthlyInvestment();
  const newData = [[
    new Date(),
    1,
    monthlyInvestment,
    '',
    '積立中'
  ]];
  sheet.getRange(lastRow + 1, 1, 1, 5).setValues(newData);
}

/**
 * エラー通知を送信
 */
function sendErrorNotification(error) {
  const subject = '【暗号資産管理】エラー通知';
  const body = `エラーが発生しました:\n\n${error.toString()}\n\nスタックトレース:\n${error.stack}`;
  
  MailApp.sendEmail({
    to: CONFIG.EMAIL_RECIPIENT,
    subject: subject,
    body: body
  });
}

/**
 * 更新ログを記録
 */
function logUpdate(message) {
  console.log(`[${new Date().toISOString()} ] ${message}`);
}

/**
 * 取引履歴を取得
 */
function getTradeHistory(sheet, symbol) {
  if (!sheet || sheet.getLastRow() < 2) return [];
  
  const dataRange = sheet.getRange(2, 1, sheet.getLastRow() - 1, 8);
  const data = dataRange.getValues();
  
  return data.filter(row => row[1] === symbol && row[2] === '売り').map(row => ({
    date: row[0],
    quantity: row[3],
    price: row[4],
    profit: row[6] - (row[3] * row[4])
  }));
}

/**
 * リバランス通知を送信
 */
function sendRebalanceNotification(items) {
  const subject = '【暗号資産】リバランス推奨のお知らせ';
  let body = 'ポートフォリオのリバランスが推奨されます。\n\n';
  
  items.forEach(item => {
    body += `${item.symbol}: ${(item.deviation * 100).toFixed(2)}%の乖離\n`;
  });
  
  body += `\n詳細はスプレッドシートをご確認ください。\n`;
  body += SpreadsheetApp.getActiveSpreadsheet().getUrl();
  
  MailApp.sendEmail({
    to: CONFIG.EMAIL_RECIPIENT,
    subject: subject,
    body: body
  });
}

/**
 * 銘柄の列インデックスを取得
 */
function getSymbolColumnIndex(sheet, symbol) {
  const headerRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());
  const headers = headerRange.getValues()[0];
  return headers.indexOf(symbol) + 1;
}

/**
 * ポートフォリオボラティリティを計算
 */
function calculatePortfolioVolatility(sheet, volatilities) {
  const dataRange = sheet.getRange(2, 1, 10, 11);
  const data = dataRange.getValues();
  
  let portfolioVol = 0;
  let totalWeight = 0;
  
  data.forEach(row => {
    const symbol = row[0];
    const weight = row[8]; // 現在比率
    
    if (volatilities[symbol] && weight > 0) {
      portfolioVol += Math.pow(weight * volatilities[symbol], 2);
      totalWeight += weight;
    }
  });
  
  return totalWeight > 0 ? Math.sqrt(portfolioVol) : 0;
}

/**
 * ボラティリティ調整を提案
 */
function proposeVolatilityAdjustment(sheet, volatilities) {
  let adjustment = '';
  
  // 高ボラティリティ銘柄を特定
  const highVolSymbols = Object.entries(volatilities)
    .filter(([_, vol]) => vol > CONFIG.VOLATILITY_THRESHOLD)
    .sort((a, b) => b[1] - a[1]);
  
  if (highVolSymbols.length > 0) {
    adjustment += '以下の高ボラティリティ銘柄の比率を削減することを推奨:\n';
    highVolSymbols.forEach(([symbol, vol]) => {
      adjustment += `- ${symbol}: ${(vol * 100).toFixed(1)}%\n`;
    });
    adjustment += '\n成長層の10%を基盤層（BTC/ETH）へシフトすることを検討してください。';
  }
  
  return adjustment;
}

/**
 * アクティブなブレイクアウトシグナルを取得
 */
function getActiveBreakoutSignals(sheet) {
  if (!sheet || sheet.getLastRow() < 2) return [];
  
  const dataRange = sheet.getRange(2, 1, sheet.getLastRow() - 1, 8);
  const data = dataRange.getValues();
  
  return data.filter(row => row[7] === 'アクティブ' && row[2] === 'BUY').map(row => ({
    symbol: row[1],
    signal: row[3],
    price: row[4]
  }));
}

/**
 * 取引後の心理状態を更新（修正版）
 */
function updatePsychologyAfterTrade(orders) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.PSYCHOLOGY);
  
  if (!sheet) return;
  
  // 損失があった場合のバイアスチェック
  const hasLoss = orders.some(order => order.loss);
  if (hasLoss) {
    const biasStartRow = 3 + CONFIG.PSYCHOLOGY_BELIEFS.length + 4; // ★修正点
    const lossAversionRow = biasStartRow; // 損失回避バイアスの行（最初の行）
    
    // 発生回数をインクリメント
    const currentCount = sheet.getRange(lossAversionRow, 3).getValue() || 0;
    sheet.getRange(lossAversionRow, 3).setValue(currentCount + 1);
    
    // 最近の事例を更新
    sheet.getRange(lossAversionRow, 2).setValue(`${new Date().toLocaleDateString('ja-JP')} - OCO損切り実行`);
    
    // 最終発生日を更新
    sheet.getRange(lossAversionRow, 4).setValue(new Date().toLocaleString('ja-JP'));
  }
}

// ===== 通知関数（拡張版） =====

/**
 * OCO約定通知
 */
function sendOCOExecutionNotification(orders) {
  const subject = '【暗号資産】OCO注文約定通知';
  let body = 'OCO注文が約定しました：\n\n';
  
  let totalProfit = 0;
  
  orders.forEach(order => {
    body += `━━━━━━━━━━━━━━━━━━━━\n`;
    body += `銘柄: ${order.symbol}\n`;
    body += `タイプ: ${order.type}\n`;
    body += `約定価格: $${order.price.toLocaleString()}\n`;
    
    if (order.profit) {
      body += `利益: ¥${order.profit.toLocaleString()}\n`;
      totalProfit += order.profit;
    } else if (order.loss) {
      body += `損失: ¥${order.loss.toLocaleString()}\n`;
      totalProfit += order.loss;
    }
  });
  
  body += `\n━━━━━━━━━━━━━━━━━━━━\n`;
  body += `合計損益: ¥${totalProfit.toLocaleString()}\n`;
  
  // 心理的アドバイス
  if (totalProfit < 0) {
    body += `\n【心理管理】\n`;
    body += `損失は避けられません。これも確率の一部です。\n`;
    body += `「不確定性」の信念を思い出し、次のエッジに集中しましょう。\n`;
  }
  
  MailApp.sendEmail({
    to: CONFIG.EMAIL_RECIPIENT,
    subject: subject,
    body: body
  });
}

/**
 * ボラティリティアラート通知
 */
function sendVolatilityAlertNotification(currentVol, adjustment) {
  const subject = '【暗号資産】ボラティリティ警告';
  let body = `ポートフォリオのボラティリティが閾値を超えました。\n\n`;
  body += `現在のボラティリティ: ${(currentVol * 100).toFixed(1)}%\n`;
  body += `設定閾値: ${(CONFIG.VOLATILITY_THRESHOLD * 100).toFixed(1)}%\n\n`;
  
  body += `【推奨調整】\n`;
  body += adjustment;
  
  body += `\n\n詳細分析: ${SpreadsheetApp.getActiveSpreadsheet().getUrl()}`;
  
  MailApp.sendEmail({
    to: CONFIG.EMAIL_RECIPIENT,
    subject: subject,
    body: body
  });
}

/**
 * 心理状態アラート
 */
function sendPsychologyAlertNotification(message) {
  const subject = '【暗号資産】心理チェック未完了';
  const body = `投資実行前に心理状態の確認が必要です。\n\n${message}\n\n` +
               `心理管理シートで確認を完了してください。\n` +
               `${SpreadsheetApp.getActiveSpreadsheet().getUrl()}`;
  
  MailApp.sendEmail({
    to: CONFIG.EMAIL_RECIPIENT,
    subject: subject,
    body: body
  });
}

/**
 * ブレイクアウトシグナル付き投資通知
 */
function sendInvestmentNotificationWithSignals(amount, signals) {
  const subject = '【暗号資産】投資実行タイミング（ブレイクアウト分析付き）';
  let body = `3ヶ月の積立期間が完了しました。\n\n`;
  body += `積立金額: ¥${amount.toLocaleString()}\n\n`;
  
  // ブレイクアウトシグナル
  if (signals.length > 0) {
    body += `【アクティブなブレイクアウトシグナル】\n`;
    signals.forEach(signal => {
      body += `${signal.symbol}: ${signal.signal} @ $${signal.price.toLocaleString()}\n`;
    });
    body += `\n`;
  } else {
    body += `【注意】現在アクティブなブレイクアウトシグナルはありません。\n\n`;
  }
  
  // 投資配分案（階段指値付き）
  body += `【推奨投資配分と階段指値】\n`;
  Object.entries(CONFIG.TARGET_ALLOCATION).forEach(([symbol, ratio]) => {
    if (symbol !== 'CASH') {
      const layer = getSymbolLayer(symbol);
      const limits = CONFIG.LIMIT_ORDER_LEVELS[layer];
      const investAmount = amount * ratio;
      
      body += `\n${symbol}: ¥${Math.floor(investAmount).toLocaleString()} (${(ratio * 100).toFixed(0)}%)\n`;
      body += `  階段指値: ${limits.map(l => `${(l * 100).toFixed(0)}%`).join(', ')}\n`;
    }
  });
  
  body += `\n現金保有: ¥${Math.floor(amount * CONFIG.TARGET_ALLOCATION.CASH).toLocaleString()} (6%)\n`;
  
  // 心理チェックリマインダー
  body += `\n【実行前チェックリスト】\n`;
  body += `□ 全ての信念を確認済み\n`;
  body += `□ 心理バイアスチェック完了\n`;
  body += `□ エッジ比率確認（MFE/MAE > 1）\n`;
  body += `□ OCO注文の準備完了\n`;
  
  body += `\n詳細分析: ${SpreadsheetApp.getActiveSpreadsheet().getUrl()}`;
  
  MailApp.sendEmail({
    to: CONFIG.EMAIL_RECIPIENT,
    subject: subject,
    body: body
  });
}

// ===== トリガー設定 =====

/**
 * トリガーを設定（ブレイクアウト通知対応版）
 */
function setupTriggersWithBreakout() {
  // 既存のトリガーを削除
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => ScriptApp.deleteTrigger(trigger));
  
  // 30分ごとの価格更新とブレイクアウトチェック
  ScriptApp.newTrigger('scheduledUpdateWithBreakout')
    .timeBased()
    .everyMinutes(30)
    .create();
  
  // 毎月1日の積立更新とAPI使用量リセット
  ScriptApp.newTrigger('monthlyAccumulationUpdate')
    .timeBased()
    .onMonthDay(1)
    .atHour(9)
    .create();
  
  // 毎日のサマリーレポート（心理チェック含む）
  ScriptApp.newTrigger('dailySummaryReportWithPsychology')
    .timeBased()
    .everyDays(1)
    .atHour(20)
    .create();
  
  // 週次エッジ分析（月曜日）
  ScriptApp.newTrigger('weeklyEdgeAnalysisReport')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(9)
    .create();
  
  // API使用量チェック（毎日午前9時）
  ScriptApp.newTrigger('checkAPIUsage')
    .timeBased()
    .everyDays(1)
    .atHour(9)
    .create();
  
  // スプレッドシートを開いたときのトリガー
  ScriptApp.newTrigger('onOpen')
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onOpen()
    .create();
  
  SpreadsheetApp.getActiveSpreadsheet().toast('トリガーを設定しました（ブレイクアウト通知対応）', '完了', 5);
}

// ===== 追加レポート機能 =====

/**
 * 日次サマリーレポート（心理状態含む）
 */
function dailySummaryReportWithPsychology() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const portfolioSheet = ss.getSheetByName(CONFIG.SHEETS.PORTFOLIO);
  const psychSheet = ss.getSheetByName(CONFIG.SHEETS.PSYCHOLOGY);
  const edgeSheet = ss.getSheetByName(CONFIG.SHEETS.EDGE_ANALYSIS);
  
  if (!portfolioSheet) return;
  
  // ポートフォリオデータを取得
  const totalValue = portfolioSheet.getRange(15, 6).getValue();
  const totalInvestment = portfolioSheet.getRange(16, 6).getValue();
  const totalProfit = portfolioSheet.getRange(17, 6).getValue();
  const profitRate = portfolioSheet.getRange(18, 6).getValue();
  
  const subject = '【暗号資産】日次レポート（タートル流・ZONE）';
  let body = `本日のポートフォリオサマリー\n`;
  body += `━━━━━━━━━━━━━━━━━━━━\n\n`;
  body += `評価額: ¥${totalValue.toLocaleString()}\n`;
  body += `投資額: ¥${totalInvestment.toLocaleString()}\n`;
  body += `損益: ¥${totalProfit.toLocaleString()}\n`;
  body += `損益率: ${(profitRate * 100).toFixed(2)}%\n\n`;
  
  // エッジ分析サマリー
  if (edgeSheet && edgeSheet.getLastRow() > 1) {
    body += `【エッジ分析】\n`;
    const topEdges = edgeSheet.getRange(2, 1, Math.min(3, edgeSheet.getLastRow() - 1), 6).getValues();
    topEdges.forEach(row => {
      if (row[0]) {
        body += `${row[0]}: エッジ比率 ${row[5].toFixed(2)}, 勝率 ${(row[2] * 100).toFixed(1)}%\n`;
      }
    });
    body += `\n`;
  }
  
  // API使用状況
  const apiUsage = getAPIUsageReport();
  body += `【API使用状況】\n`;
  body += `CoinGecko: ${apiUsage.used} / ${apiUsage.limit} (${apiUsage.percentage}%)\n`;
  if (apiUsage.percentage > 70) {
    body += `⚠️ API使用量が${apiUsage.percentage}%に達しています\n`;
  }
  body += `\n`;
  
  // 心理チェックリマインダー
  body += `【明日の取引前チェック】\n`;
  CONFIG.PSYCHOLOGY_BELIEFS.slice(0, 3).forEach(belief => {
    body += `□ ${belief}\n`;
  });
  
  body += `\n詳細: ${ss.getUrl()}`;
  
  MailApp.sendEmail({
    to: CONFIG.EMAIL_RECIPIENT,
    subject: subject,
    body: body
  });
}

/**
 * 週次エッジ分析レポート
 */
function weeklyEdgeAnalysisReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const edgeSheet = ss.getSheetByName(CONFIG.SHEETS.EDGE_ANALYSIS);
  
  if (!edgeSheet || edgeSheet.getLastRow() < 2) return;
  
  const subject = '【暗号資産】週次エッジ分析レポート';
  let body = `今週のエッジ分析結果\n`;
  body += `━━━━━━━━━━━━━━━━━━━━\n\n`;
  
  const dataRange = edgeSheet.getRange(2, 1, edgeSheet.getLastRow() - 1, 7);
  const data = dataRange.getValues();
  
  // エッジ比率でソート
  data.sort((a, b) => b[5] - a[5]);
  
  body += `【高エッジ銘柄 TOP3】\n`;
  data.slice(0, 3).forEach((row, index) => {
    body += `${index + 1}. ${row[0]}\n`;
    body += `   エッジ比率: ${row[5].toFixed(2)}\n`;
    body += `   期待値: ¥${row[6].toLocaleString()}\n`;
    body += `   取引数: ${row[1]}\n\n`;
  });
  
  // 警告
  const lowEdges = data.filter(row => row[5] < CONFIG.EDGE_RATIO_MIN);
  if (lowEdges.length > 0) {
    body += `【警告】エッジ比率が1.0未満の銘柄:\n`;
    lowEdges.forEach(row => {
      body += `- ${row[0]}: ${row[5].toFixed(2)}\n`;
    });
    body += `\nこれらの銘柄は取引を控えることを推奨します。\n`;
  }
  
  body += `\n詳細分析: ${ss.getUrl()}`;
  
  MailApp.sendEmail({
    to: CONFIG.EMAIL_RECIPIENT,
    subject: subject,
    body: body
  });
}

/**
 * 毎月の積立更新（心理リセット付き）
 */
function monthlyAccumulationUpdate() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const accSheet = ss.getSheetByName(CONFIG.SHEETS.ACCUMULATION);
  
  if (!accSheet) return;
  
  const lastRow = accSheet.getLastRow();
  const currentMonth = accSheet.getRange(lastRow, 2).getValue();
  const currentAmount = accSheet.getRange(lastRow, 3).getValue();
  const monthlyInvestment = getMonthlyInvestment();
  
  // 積立金額を更新
  accSheet.getRange(lastRow, 2).setValue(currentMonth + 1);
  accSheet.getRange(lastRow, 3).setValue(currentAmount + monthlyInvestment);
  
  // 月初の心理チェックリセット
  resetPsychologyChecklist();
  
  // API使用量カウンターをリセット（新月の場合）
  const currentMonthKey = new Date().toISOString().slice(0, 7);
  const props = PropertiesService.getScriptProperties();
  const allKeys = props.getKeys();
  
  // 古い月のカウンターを削除
  allKeys.forEach(key => {
    if (key.startsWith('coingecko_usage_') && !key.includes(currentMonthKey)) {
      props.deleteProperty(key);
    }
  });
}

/**
 * API使用量をチェックして警告
 */
function checkAPIUsage() {
  const report = getAPIUsageReport();
  
  // 80%以上使用で警告
  if (report.percentage >= 80) {
    const subject = '【暗号資産】API使用量警告';
    const body = `CoinGecko API使用量が制限に近づいています。\n\n` +
                 `使用量: ${report.used} / ${report.limit} (${report.percentage}%)\n` +
                 `残り: ${report.remaining}コール\n\n` +
                 `価格更新頻度の調整を検討してください。`;
    
    MailApp.sendEmail({
      to: CONFIG.EMAIL_RECIPIENT,
      subject: subject,
      body: body
    });
  }
}

/**
 * 手動価格更新（すべての無料APIを使用）
 */
function manualPriceUpdate() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.PORTFOLIO);
  
  if (!sheet) {
    SpreadsheetApp.getActiveSpreadsheet().toast('ポートフォリオシートが見つかりません', 'エラー', 5);
    return;
  }
  
  // 無料APIのみで価格更新
  const symbols = Object.keys(CONFIG.TARGET_ALLOCATION).filter(s => s !== 'CASH');
  let successCount = 0;
  let failedSymbols = [];
  
  symbols.forEach((symbol, index) => {
    // 複数の無料APIを順番に試す
    let price = null;
    let source = '';
    
    // 1. Binance
    price = getBinancePrice(symbol);
    if (price) {
      source = 'Binance';
    } else {
      // 2. CryptoCompare
      price = getCryptoComparePrice(symbol);
      if (price) source = 'CryptoCompare';
    }
    
    if (!price) {
      // 3. CoinCap
      price = getCoinCapPrice(symbol);
      if (price) source = 'CoinCap';
    }
    
    if (price && price > 0) {
      sheet.getRange(index + 2, 5).setValue(price); // E列に価格を設定
      console.log(`${symbol}: ${price} (${source})`);
      successCount++;
    } else {
      failedSymbols.push(symbol);
    }
  });
  
  // ポートフォリオの再計算
  updatePortfolioCalculations(sheet);
  
  // 更新時刻を記録
  sheet.getRange(1, 12).setValue('最終更新: ' + new Date().toLocaleString('ja-JP') + ' (無料API)');
  
  // 結果を通知
  let message = `${successCount}/${symbols.length} 銘柄の価格を更新しました（無料API使用）`;
  if (failedSymbols.length > 0) {
    message += `\n失敗: ${failedSymbols.join(', ')}`;
  }
  
  SpreadsheetApp.getActiveSpreadsheet().toast(message, '価格更新完了', 5);
}

/**
 * ポートフォリオの計算を更新（価格更新後）
 */
function updatePortfolioCalculations(sheet) {
  const symbols = Object.keys(CONFIG.TARGET_ALLOCATION).filter(s => s !== 'CASH');
  const dataRange = sheet.getRange(2, 1, symbols.length, 11);
  const data = dataRange.getValues();
  
  // 各行の計算を更新
  data.forEach((row, index) => {
    const quantity = row[2]; // C列: 保有数量
    const avgPrice = row[3]; // D列: 平均取得単価
    const currentPrice = row[4]; // E列: 現在価格
    
    if (currentPrice > 0) {
      data[index][5] = quantity * currentPrice; // F列: 評価額
      data[index][6] = (currentPrice - avgPrice) * quantity; // G列: 損益
      data[index][7] = avgPrice > 0 ? (currentPrice - avgPrice) / avgPrice : 0; // H列: 損益率
    }
  });
  
  // ポートフォリオ比率の計算
  const totalValue = data.reduce((sum, row) => sum + (row[5] || 0), 0);
  data.forEach((row, index) => {
    data[index][8] = totalValue > 0 ? row[5] / totalValue : 0; // I列: 現在比率
    data[index][10] = data[index][8] - data[index][9]; // K列: 乖離率
  });
  
  // データを更新
  dataRange.setValues(data);
  
  // サマリー更新
  updateSummary(sheet, totalValue);
}

// ===== 追加関数（デバッグ・修復用） =====

/**
 * 心理管理シートを修復（デバッグ用）
 */
function repairPsychologySheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 既存のシートを削除して再作成
  const oldSheet = ss.getSheetByName(CONFIG.SHEETS.PSYCHOLOGY);
  if (oldSheet) {
    ss.deleteSheet(oldSheet);
  }
  
  // 新しくシートを作成
  createPsychologySheet(ss);
  initializePsychologyChecklist(ss);
  
  SpreadsheetApp.getActiveSpreadsheet().toast('心理管理シートを修復しました', '完了', 5);
}

/**
 * 心理バイアスを手動で記録
 */
function recordPsychologyBias(biasType, description) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.PSYCHOLOGY);
  
  if (!sheet) return;
  
  const biasStartRow = 3 + CONFIG.PSYCHOLOGY_BELIEFS.length + 4; // ★修正点
  
  // バイアスタイプの行を探す
  const biasIndex = CONFIG.BIAS_CHECKLIST.indexOf(biasType);
  if (biasIndex === -1) return;
  
  const targetRow = biasStartRow + biasIndex;
  
  // 発生回数をインクリメント
  const currentCount = sheet.getRange(targetRow, 3).getValue() || 0;
  sheet.getRange(targetRow, 3).setValue(currentCount + 1);
  
  // 最近の事例を更新
  const example = description || `${new Date().toLocaleDateString('ja-JP')} - 手動記録`;
  sheet.getRange(targetRow, 2).setValue(example);
  
  // 最終発生日を更新
  sheet.getRange(targetRow, 4).setValue(new Date().toLocaleString('ja-JP'));
  
  // 警告チェック
  if (currentCount + 1 > 3) {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      `⚠️ ${biasType}の発生回数が${currentCount + 1}回になりました。取引を控えることを推奨します。`,
      '心理バイアス警告',
      10
    );
  }
}

/**
 * バイアス記録ダイアログを表示
 */
function showBiasRecordDialog() {
  const html = `
    <div style="padding: 20px;">
      <h3>心理バイアスを記録</h3>
      <form onsubmit="handleSubmit(event)">
        <div style="margin-bottom: 15px;">
          <label>バイアスの種類:</label><br>
          <select id="biasType" style="width: 100%; padding: 5px;">
            ${CONFIG.BIAS_CHECKLIST.map(bias => 
              `<option value="${bias}">${bias}</option>`
            ).join('')}
          </select>
        </div>
        <div style="margin-bottom: 15px;">
          <label>詳細（任意）:</label><br>
          <textarea id="description" style="width: 100%; height: 60px; padding: 5px;" 
                    placeholder="どのような状況で発生したか記録"></textarea>
        </div>
        <button type="submit" style="width: 100%; padding: 10px; background: #ff4444; color: white; border: none; cursor: pointer;">
          バイアスを記録
        </button>
      </form>
    </div>
    <script>
      function handleSubmit(event) {
        event.preventDefault();
        const data = {
          biasType: document.getElementById('biasType').value,
          description: document.getElementById('description').value
        };
        google.script.run
          .withSuccessHandler(() => {
            google.script.host.close();
          })
          .recordPsychologyBias(data.biasType, data.description);
      }
    </script>
  `;
  
  const ui = SpreadsheetApp.getUi();
  ui.showModalDialog(HtmlService.createHtmlOutput(html).setWidth(400).setHeight(300), 'バイアス記録');
}

/**
 * ブレイクアウトに基づく取引実行
 */
function executeBreakoutTrade(symbol, quantity, limitPrices) {
  // 心理チェック
  const psychCheck = performPsychologyCheck();
  if (!psychCheck.passed) {
    SpreadsheetApp.getActiveSpreadsheet().toast(psychCheck.message, '心理チェック未完了', 10);
    return;
  }
  
  const price = getCryptoPrice(symbol);
  const layer = getSymbolLayer(symbol);
  const limits = limitPrices || CONFIG.LIMIT_ORDER_LEVELS[layer];
  
  // 階段指値を記録
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const alertSheet = ss.getSheetByName(CONFIG.SHEETS.ALERTS);
  
  limits.forEach((limit, index) => {
    const targetPrice = price * (1 + limit);
    const alertData = [[
      symbol,
      '以下',
      targetPrice,
      price,
      new Date(),
      'アクティブ',
      ''
    ]];
    
    const lastRow = alertSheet.getLastRow();
    alertSheet.getRange(lastRow + 1, 1, 1, 7).setValues(alertData);
  });
  
  SpreadsheetApp.getActiveSpreadsheet().toast(
    `${symbol}の階段指値を設定しました: ${limits.map(l => `${(l * 100).toFixed(0)}%`).join(', ')}`,
    '指値設定完了',
    5
  );
}

/**
 * 取引記録ダイアログを表示
 */
function showTransactionDialog() {
  const html = `
    <div style="padding: 20px;">
      <h3>取引を記録</h3>
      <form onsubmit="handleSubmit(event)">
        <div style="margin-bottom: 10px;">
          <label>銘柄:</label><br>
          <select id="symbol" style="width: 100%;">
            <option value="BTC">BTC</option>
            <option value="ETH">ETH</option>
            <option value="SOL">SOL</option>
            <option value="NEAR">NEAR</option>
            <option value="ADA">ADA</option>
            <option value="AAVE">AAVE</option>
            <option value="HBAR">HBAR</option>
            <option value="GRT">GRT</option>
            <option value="ALGO">ALGO</option>
          </select>
        </div>
        <div style="margin-bottom: 10px;">
          <label>売買:</label><br>
          <select id="action" style="width: 100%;">
            <option value="買い">買い</option>
            <option value="売り">売り</option>
          </select>
        </div>
        <div style="margin-bottom: 10px;">
          <label>数量:</label><br>
          <input type="number" id="quantity" step="0.0001" required style="width: 100%;">
        </div>
        <div style="margin-bottom: 10px;">
          <label>価格 (USD):</label><br>
          <input type="number" id="price" step="0.01" required style="width: 100%;">
        </div>
        <div style="margin-bottom: 10px;">
          <label>メモ:</label><br>
          <input type="text" id="memo" style="width: 100%;">
        </div>
        <button type="submit" style="width: 100%; padding: 10px;">記録する</button>
      </form>
    </div>
    <script>
      function handleSubmit(event) {
        event.preventDefault();
        const data = {
          symbol: document.getElementById('symbol').value,
          action: document.getElementById('action').value,
          quantity: parseFloat(document.getElementById('quantity').value),
          price: parseFloat(document.getElementById('price').value),
          memo: document.getElementById('memo').value
        };
        google.script.run
          .withSuccessHandler(() => google.script.host.close())
          .recordTransactionFromDialog(data);
      }
    </script>
  `;
  
  const ui = SpreadsheetApp.getUi();
  ui.showModalDialog(HtmlService.createHtmlOutput(html).setWidth(350).setHeight(450), '取引記録');
}

/**
 * ダイアログから取引を記録
 */
function recordTransactionFromDialog(data) {
  recordTransactionWithOCO(data.symbol, data.action, data.quantity, data.price, data.memo);
}

/**
 * 手動で取引を記録（OCO付き）
 */
function recordTransactionWithOCO(symbol, action, quantity, price, memo = '') {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const historySheet = ss.getSheetByName(CONFIG.SHEETS.HISTORY);
  const portfolioSheet = ss.getSheetByName(CONFIG.SHEETS.PORTFOLIO);
  
  if (!historySheet || !portfolioSheet) return;
  
  // 手数料計算
  const fee = quantity * price * CONFIG.EXCHANGE_FEE;
  const total = action === '買い' ? (quantity * price + fee) : (quantity * price - fee);
  
  // 履歴に記録
  const newRow = [
    new Date(),
    symbol,
    action,
    quantity,
    price,
    fee,
    total,
    memo
  ];
  
  const lastRow = historySheet.getLastRow();
  historySheet.getRange(lastRow + 1, 1, 1, 8).setValues([newRow]);
  
  // ポートフォリオを更新
  updatePortfolioAfterTransaction(portfolioSheet, symbol, action, quantity, price);
  
  // 買いの場合はOCO注文を自動設定
  if (action === '買い') {
    setOCOOrder(symbol, price, quantity);
  }
  
  // エッジ分析を更新
  updateEdgeAnalysis();
  
  SpreadsheetApp.getActiveSpreadsheet().toast(
    `${symbol}の取引を記録しました。OCO注文も設定済みです。`,
    '取引記録完了',
    5
  );
}

/**
 * 取引後のポートフォリオ更新
 */
function updatePortfolioAfterTransaction(sheet, symbol, action, quantity, price) {
  const dataRange = sheet.getRange(2, 1, 10, 4);
  const data = dataRange.getValues();
  
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === symbol) {
      const currentQuantity = data[i][2];
      const currentAvgPrice = data[i][3];
      
      if (action === '買い') {
        // 平均取得単価を更新
        const newQuantity = currentQuantity + quantity;
        const newAvgPrice = ((currentQuantity * currentAvgPrice) + (quantity * price)) / newQuantity;
        data[i][2] = newQuantity;
        data[i][3] = newAvgPrice;
      } else {
        // 売却の場合
        data[i][2] = Math.max(0, currentQuantity - quantity);
      }
      
      break;
    }
  }
  
  dataRange.setValues(data);
  
  // 価格を更新して再計算
  updateAllPrices();
}

/**
 * ブレイクアウト履歴を表示
 */
function showBreakoutHistory() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.BREAKOUT);
  
  if (!sheet || sheet.getLastRow() < 2) {
    SpreadsheetApp.getActiveSpreadsheet().toast('ブレイクアウト履歴がありません', '情報', 3);
    return;
  }
  
  // ブレイクアウトシートをアクティブにする
  ss.setActiveSheet(sheet);
  
  // 最新のシグナルを取得
  const lastRow = sheet.getLastRow();
  const recentSignals = sheet.getRange(Math.max(2, lastRow - 4), 1, Math.min(5, lastRow - 1), 8).getValues();
  
  let summary = '【最近のブレイクアウトシグナル】\n';
  recentSignals.forEach(signal => {
    if (signal[0]) { // 日時が存在する場合
      summary += `${signal[1]} - ${signal[2]} - ${signal[3]}\n`;
    }
  });
  
  const ui = SpreadsheetApp.getUi();
  ui.alert('ブレイクアウト履歴', summary, ui.ButtonSet.OK);
}

/**
 * ブレイクアウト通知のテスト（デバッグ用）
 */
function testBreakoutNotification() {
  // テスト用のダミーシグナルを作成
  const testSignals = [
    {
      symbol: 'BTC',
      type: 'BUY',
      signal: '20日高値ブレイクアウト',
      price: 98000,
      previousHigh: 95000,
      timestamp: new Date()
    },
    {
      symbol: 'ETH',
      type: 'SELL',
      signal: '20日安値ブレイクダウン',
      price: 3200,
      previousLow: 3300,
      timestamp: new Date()
    }
  ];
  
  try {
    sendBreakoutNotification(testSignals);
    SpreadsheetApp.getActiveSpreadsheet().toast('テスト通知を送信しました', '成功', 5);
  } catch (e) {
    SpreadsheetApp.getActiveSpreadsheet().toast('通知送信に失敗しました: ' + e.toString(), 'エラー', 10);
  }
}

// ===== カスタムメニュー修正版 =====

/**
 * スプレッドシートを開いた時にカスタムメニューを作成
 * 注意：この関数が動作しない場合は、setupCustomMenu()を手動で実行してください
 */
function onOpen() {
  try {
    setupCustomMenu();
  } catch (e) {
    console.error('onOpen エラー:', e);
  }
}

/**
 * カスタムメニューを設定（手動実行可能）
 */
function setupCustomMenu() {
  const ui = SpreadsheetApp.getUi();
  
  ui.createMenu('暗号資産管理')
    .addItem('初期設定', 'initializeSpreadsheet')
    .addSeparator()
    .addSubMenu(ui.createMenu('価格更新')
      .addItem('手動で価格を更新', 'manualPriceUpdate')
      .addItem('ブレイクアウトをチェック', 'manualBreakoutCheck')
      .addItem('キャッシュをクリア', 'clearPriceCache'))
    .addSubMenu(ui.createMenu('設定')
      .addItem('システム設定', 'showSettingsDialog')
      .addItem('API使用状況', 'showAPIUsageDialog'))
    .addSubMenu(ui.createMenu('取引')
      .addItem('取引を記録', 'showTransactionDialog')
      .addItem('心理チェック', 'executePsychologyCheck')
      .addItem('バイアスを記録', 'showBiasRecordDialog'))
    .addSubMenu(ui.createMenu('レポート')
      .addItem('日次レポートを送信', 'dailySummaryReportWithPsychology')
      .addItem('週次エッジ分析を送信', 'weeklyEdgeAnalysisReport')
      .addItem('ブレイクアウト履歴', 'showBreakoutHistory'))
    .addSubMenu(ui.createMenu('デバッグ')
      .addItem('API接続テスト', 'testAPIs')
      .addItem('心理管理シートを修復', 'repairPsychologySheet')
      .addItem('メニューを再設定', 'setupCustomMenu'))
    .addToUi();
    
  SpreadsheetApp.getActiveSpreadsheet().toast('カスタムメニューを設定しました', '完了', 3);
}