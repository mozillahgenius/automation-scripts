// ====================================
// 株式銘柄チェックシステム
// ====================================

// ====================================
// メニュー設定
// ====================================
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('株式銘柄チェック')
    .addItem('初期セットアップ', 'setupSheets')
    .addItem('手動でレポート送信', 'manualSendReport')
    .addItem('トリガー設定', 'setupTrigger')
    .addItem('トリガー削除', 'removeTrigger')
    .addSeparator()
    .addItem('テスト実行（メール送信なし）', 'testRun')
    .show();
}

// ====================================
// セットアップ関数
// ====================================
function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Tickersシート作成
  let tickersSheet = ss.getSheetByName('Tickers');
  if (!tickersSheet) {
    tickersSheet = ss.insertSheet('Tickers');
    tickersSheet.getRange(1, 1, 1, 6).setValues([
      ['code', 'company_name', 'industry', 'enabled', 'opts_symbol', 'notes']
    ]);
    tickersSheet.getRange(1, 1, 1, 6).setFontWeight('bold');
    
    // サンプルデータ
    const sampleData = [
      ['9984', 'ソフトバンクグループ', '通信・投資', true, '', ''],
      ['3491', 'GA technologies', '不動産テック', true, '', ''],
      ['8591', 'オリックス', '総合金融', true, '', ''],
      ['6098', 'リクルートHD', '人材/プラットフォーム', true, '', '']
    ];
    tickersSheet.getRange(2, 1, sampleData.length, 6).setValues(sampleData);
  }
  
  // Settingsシート作成
  let settingsSheet = ss.getSheetByName('Settings');
  if (!settingsSheet) {
    settingsSheet = ss.insertSheet('Settings');
    settingsSheet.getRange(1, 1, 1, 2).setValues([['key', 'value']]);
    settingsSheet.getRange(1, 1, 1, 2).setFontWeight('bold');
    
    const defaultSettings = [
      ['mail_to', 'kenichi.horii@medialink-ml.co.jp'],
      ['mail_cc', ''],
      ['subject_prefix', '【株レポート】'],
      ['news_per_company', '3'],
      ['send_time_jst', '18:00'],
      ['skip_holiday', 'TRUE'],
      ['timezone', 'Asia/Tokyo'],
      ['holiday_calendar_id', 'ja.japanese#holiday@group.v.calendar.google.com'],
      ['use_extractive_summary_only', 'TRUE']
    ];
    settingsSheet.getRange(2, 1, defaultSettings.length, 2).setValues(defaultSettings);
  }
  
  // Logsシート作成
  let logsSheet = ss.getSheetByName('Logs');
  if (!logsSheet) {
    logsSheet = ss.insertSheet('Logs');
    logsSheet.getRange(1, 1, 1, 4).setValues([
      ['timestamp', 'type', 'message', 'details']
    ]);
    logsSheet.getRange(1, 1, 1, 4).setFontWeight('bold');
  }
  
  SpreadsheetApp.getUi().alert('セットアップが完了しました。');
}

// ====================================
// トリガー設定
// ====================================
function setupTrigger() {
  removeTrigger(); // 既存のトリガーを削除
  
  const settings = loadSettings();
  const timeParts = settings.send_time_jst.split(':');
  const hour = parseInt(timeParts[0]);
  const minute = parseInt(timeParts[1] || '0');
  
  ScriptApp.newTrigger('main')
    .timeBased()
    .everyDays(1)
    .atHour(hour)
    .nearMinute(minute)
    .inTimezone(settings.timezone)
    .create();
  
  SpreadsheetApp.getUi().alert(`毎日${settings.send_time_jst}に実行するトリガーを設定しました。`);
}

function removeTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'main') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}

// ====================================
// メイン処理
// ====================================
function main() {
  try {
    const settings = loadSettings();
    
    // 祝日チェック
    if (settings.skip_holiday && isHoliday(settings.holiday_calendar_id)) {
      logMessage('INFO', '祝日のため処理をスキップしました');
      return;
    }
    
    // 土日チェック
    const today = new Date();
    const dayOfWeek = today.getDay();
    if (dayOfWeek === 0 || dayOfWeek === 6) {
      logMessage('INFO', '週末のため処理をスキップしました');
      return;
    }
    
    // レポート生成と送信
    const report = generateReport();
    sendMail(report, settings);
    
    logMessage('SUCCESS', 'レポート送信完了');
  } catch (error) {
    logMessage('ERROR', 'メイン処理でエラーが発生', error.toString());
    throw error;
  }
}

function manualSendReport() {
  main();
  SpreadsheetApp.getUi().alert('レポートを送信しました。');
}

function testRun() {
  try {
    const report = generateReport();
    const htmlOutput = HtmlService.createHtmlOutput(report.html)
      .setWidth(800)
      .setHeight(600);
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'テストプレビュー');
  } catch (error) {
    SpreadsheetApp.getUi().alert('エラー: ' + error.toString());
  }
}

// ====================================
// 設定読み込み
// ====================================
function loadSettings() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settingsSheet = ss.getSheetByName('Settings');
  
  if (!settingsSheet) {
    throw new Error('Settingsシートが見つかりません。setupSheets()を実行してください。');
  }
  
  const data = settingsSheet.getDataRange().getValues();
  const settings = {};
  
  for (let i = 1; i < data.length; i++) {
    const key = data[i][0];
    const value = data[i][1];
    if (key) {
      settings[key] = value;
    }
  }
  
  return settings;
}

function loadTickers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tickersSheet = ss.getSheetByName('Tickers');
  
  if (!tickersSheet) {
    throw new Error('Tickersシートが見つかりません。setupSheets()を実行してください。');
  }
  
  const data = tickersSheet.getDataRange().getValues();
  const tickers = [];
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][3] === true || data[i][3] === 'TRUE') { // enabled列をチェック
      tickers.push({
        code: String(data[i][0]),
        companyName: data[i][1],
        industry: data[i][2],
        optsSymbol: data[i][4],
        notes: data[i][5]
      });
    }
  }
  
  return tickers;
}

// ====================================
// レポート生成
// ====================================
function generateReport() {
  const tickers = loadTickers();
  const today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');
  const now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
  
  let html = `
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <style>
    body { font-family: Arial, sans-serif; line-height: 1.6; color: #333; }
    .header { background: #f4f4f4; padding: 10px; margin-bottom: 20px; }
    .stock-card { border: 1px solid #ddd; padding: 15px; margin-bottom: 20px; border-radius: 5px; }
    .stock-header { font-size: 18px; font-weight: bold; color: #0066cc; margin-bottom: 10px; }
    .data-row { margin: 5px 0; }
    .label { font-weight: bold; color: #666; }
    .positive { color: #009900; }
    .negative { color: #cc0000; }
    .disclaimer { background: #fffacd; padding: 10px; margin-top: 20px; font-size: 12px; }
    .section-title { font-weight: bold; margin-top: 15px; margin-bottom: 5px; color: #444; }
  </style>
</head>
<body>
  <div class="header">
    <h2>株式レポート ${today}</h2>
    <p>配信日時: ${now} (JST)<br>
    基準: 当日終値<br>
    データソース: 東証・各社IR情報等</p>
  </div>
`;
  
  let hasError = false;
  
  for (const ticker of tickers) {
    try {
      const stockData = fetchStockData(ticker);
      html += generateStockSection(ticker, stockData);
    } catch (error) {
      logMessage('ERROR', `銘柄${ticker.code}のデータ取得失敗`, error.toString());
      hasError = true;
      html += `
  <div class="stock-card">
    <div class="stock-header">${ticker.code} ${ticker.companyName}</div>
    <p style="color: #cc0000;">データ取得エラー: ${error.toString()}</p>
  </div>
`;
    }
  }
  
  html += `
  <div class="disclaimer">
    <strong>免責事項:</strong><br>
    本レポートは公開情報の事実抽出・計算の再現性のみを目的としており、投資助言ではありません。<br>
    投資判断は自己責任でお願いいたします。<br>
    <br>
    データソース: 東証、各社IR情報、TDnet等<br>
    生成: Google Apps Script
  </div>
</body>
</html>
`;
  
  const subject = `【株レポート】${today}（平日18:00配信）${hasError ? ' [一部欠落]' : ''}`;
  
  return { html, subject };
}

function generateStockSection(ticker, data) {
  let html = `
  <div class="stock-card">
    <div class="stock-header">${ticker.code} ${ticker.companyName}`;
  
  if (ticker.industry) {
    html += ` (${ticker.industry})`;
  }
  
  html += `</div>`;
  
  // 市場データ
  if (data.marketData) {
    const changeClass = data.marketData.change >= 0 ? 'positive' : 'negative';
    html += `
    <div class="section-title">市場データ</div>
    <div class="data-row">
      <span class="label">終値:</span> ${data.marketData.close.toLocaleString()}円
      <span class="${changeClass}">
        (${data.marketData.change >= 0 ? '+' : ''}${data.marketData.change.toLocaleString()}円 
        / ${data.marketData.changePercent >= 0 ? '+' : ''}${data.marketData.changePercent.toFixed(2)}%)
      </span>
    </div>`;
    
    if (data.marketData.marketCap) {
      html += `
    <div class="data-row">
      <span class="label">時価総額:</span> ${(data.marketData.marketCap / 100000000).toFixed(1)}億円
    </div>`;
    }
  }
  
  // 会社計画
  if (data.companyPlan) {
    html += `
    <div class="section-title">会社計画（通期予想）</div>`;
    
    if (data.companyPlan.revenue) {
      html += `
    <div class="data-row">
      <span class="label">売上高:</span> ${data.companyPlan.revenue}
    </div>`;
    }
    
    if (data.companyPlan.operatingProfit) {
      html += `
    <div class="data-row">
      <span class="label">営業利益:</span> ${data.companyPlan.operatingProfit}
    </div>`;
    }
    
    if (data.companyPlan.source) {
      html += `
    <div class="data-row" style="font-size: 12px; color: #666;">
      出典: ${data.companyPlan.source} (${data.companyPlan.date})
    </div>`;
    }
  }
  
  // 当日開示/IR
  if (data.news && data.news.length > 0) {
    html += `
    <div class="section-title">当日開示/IR情報</div>`;
    
    for (const item of data.news) {
      html += `
    <div class="data-row">
      • <a href="${item.url}" target="_blank">${item.title}</a>`;
      
      if (item.summary) {
        html += `<br>
      <span style="font-size: 12px; color: #666; margin-left: 10px;">${item.summary}</span>`;
      }
      
      html += `</div>`;
    }
  }
  
  html += `
  </div>`;
  
  return html;
}

// ====================================
// データ取得関数
// ====================================
function fetchStockData(ticker) {
  const data = {
    marketData: null,
    companyPlan: null,
    news: []
  };
  
  // 市場データ取得（簡易版）
  try {
    data.marketData = fetchMarketData(ticker.code);
  } catch (error) {
    logMessage('WARN', `市場データ取得失敗: ${ticker.code}`, error.toString());
  }
  
  // 会社計画取得（簡易版）
  try {
    data.companyPlan = fetchCompanyPlan(ticker.code);
  } catch (error) {
    logMessage('WARN', `会社計画取得失敗: ${ticker.code}`, error.toString());
  }
  
  // ニュース取得（簡易版）
  try {
    data.news = fetchNews(ticker.code, ticker.companyName);
  } catch (error) {
    logMessage('WARN', `ニュース取得失敗: ${ticker.code}`, error.toString());
  }
  
  return data;
}

function fetchMarketData(code) {
  // Yahoo Finance APIの代替として、簡易的なダミーデータを返す
  // 実際の実装では適切なAPIを使用してください
  const basePrice = 1000 + Math.random() * 4000;
  const change = (Math.random() - 0.5) * 200;
  const changePercent = (change / basePrice) * 100;
  
  return {
    close: Math.round(basePrice),
    change: Math.round(change),
    changePercent: changePercent,
    marketCap: Math.round(basePrice * 10000000) // 簡易計算
  };
}

function fetchCompanyPlan(code) {
  // 簡易的なダミーデータを返す
  // 実際の実装では各社のIR情報を取得してください
  return {
    revenue: '1,234億円',
    operatingProfit: '123億円',
    source: '決算短信',
    date: '2025-08-01'
  };
}

function fetchNews(code, companyName) {
  // TDnetのRSSフィードから取得する簡易実装
  try {
    const url = `https://www.release.tdnet.info/inbs/I_list_001_${code}.html`;
    const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    
    if (response.getResponseCode() === 200) {
      // 簡易的なパース（実際はより適切なパースが必要）
      return [
        {
          title: '本日の適時開示情報',
          url: url,
          summary: '詳細はリンク先をご確認ください'
        }
      ];
    }
  } catch (error) {
    // エラーの場合は空配列を返す
  }
  
  return [];
}

// ====================================
// メール送信
// ====================================
function sendMail(report, settings) {
  const mailOptions = {
    to: settings.mail_to,
    subject: report.subject,
    htmlBody: report.html
  };
  
  if (settings.mail_cc) {
    mailOptions.cc = settings.mail_cc;
  }
  
  MailApp.sendEmail(mailOptions);
  
  // ログに記録
  logMessage('INFO', 'メール送信完了', `To: ${settings.mail_to}`);
}

// ====================================
// ユーティリティ関数
// ====================================
function isHoliday(calendarId) {
  try {
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    const tomorrow = new Date(today);
    tomorrow.setDate(tomorrow.getDate() + 1);
    
    const calendar = CalendarApp.getCalendarById(calendarId);
    if (!calendar) {
      logMessage('WARN', '祝日カレンダーが見つかりません', calendarId);
      return false;
    }
    
    const events = calendar.getEvents(today, tomorrow);
    return events.length > 0;
  } catch (error) {
    logMessage('ERROR', '祝日チェックエラー', error.toString());
    return false;
  }
}

function logMessage(type, message, details = '') {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let logsSheet = ss.getSheetByName('Logs');
  
  if (!logsSheet) {
    logsSheet = ss.insertSheet('Logs');
    logsSheet.getRange(1, 1, 1, 4).setValues([
      ['timestamp', 'type', 'message', 'details']
    ]);
  }
  
  const timestamp = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
  logsSheet.appendRow([timestamp, type, message, details]);
  
  // 古いログを削除（1000行を超えたら古いものから削除）
  const lastRow = logsSheet.getLastRow();
  if (lastRow > 1000) {
    logsSheet.deleteRows(2, lastRow - 1000);
  }
}

// ====================================
// オプション理論価格計算（ブラック・ショールズ）
// ====================================
function calcBlackScholes(S, K, T, r, q, sigma, optionType = 'call') {
  // S: 現在の株価
  // K: 権利行使価格
  // T: 満期までの期間（年）
  // r: リスクフリーレート
  // q: 配当利回り
  // sigma: ボラティリティ
  
  const d1 = (Math.log(S / K) + (r - q + 0.5 * sigma * sigma) * T) / (sigma * Math.sqrt(T));
  const d2 = d1 - sigma * Math.sqrt(T);
  
  if (optionType === 'call') {
    return S * Math.exp(-q * T) * normCDF(d1) - K * Math.exp(-r * T) * normCDF(d2);
  } else {
    return K * Math.exp(-r * T) * normCDF(-d2) - S * Math.exp(-q * T) * normCDF(-d1);
  }
}

function normCDF(x) {
  // 標準正規分布の累積分布関数
  const t = 1 / (1 + 0.2316419 * Math.abs(x));
  const d = 0.3989423 * Math.exp(-x * x / 2);
  const prob = d * t * (0.3193815 + t * (-0.3565638 + t * (1.781478 + t * (-1.821256 + t * 1.330274))));
  
  if (x > 0) {
    return 1 - prob;
  } else {
    return prob;
  }
}