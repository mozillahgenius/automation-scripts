// ====================================
// 株式銘柄チェックシステム
// ====================================
// 
// 機能概要:
// - 平日18:00に指定銘柄の株価情報・会社計画・開示情報をメール配信
// - Perplexity APIを使用して最新の公開情報を自動取得
// - 事実ベースの情報のみを抽出（予測・推測は除外）
//
// 使い方:
// 1. メニュー「株式銘柄チェック」→「初期セットアップ」を実行
// 2. メニュー「株式銘柄チェック」→「Perplexity API Key設定」でAPIキーを設定
// 3. Tickersシートに監視したい銘柄を追加
// 4. Settingsシートでメール送信先等を設定
// 5. メニュー「株式銘柄チェック」→「トリガー設定」で自動実行を設定
//
// 必要なAPI:
// - Perplexity API (https://www.perplexity.ai/settings/api)
// ====================================

// ====================================
// メニュー設定
// ====================================
function onOpen(e) {
  try {
    // スプレッドシートのUIを取得
    const ui = SpreadsheetApp.getUi();
    
    // メニューを作成
    const menu = ui.createMenu('株式銘柄チェック');
    menu.addItem('初期セットアップ', 'setupSheets');
    menu.addItem('手動でレポート送信', 'manualSendReport');
    menu.addItem('トリガー設定', 'setupTrigger');
    menu.addItem('トリガー削除', 'removeTrigger');
    menu.addSeparator();
    menu.addItem('Perplexity API Key設定', 'setPerplexityApiKey');
    menu.addItem('テスト実行（メール送信なし）', 'testRun');
    menu.addItem('価格取得テスト', 'testPriceRetrieval');
    menu.addItem('業界情報テスト', 'testIndustryInfo');
    menu.addToUi();
  } catch (error) {
    console.error('メニュー作成エラー:', error);
    // エラーが発生した場合、シンプルなメニューを試す
    SpreadsheetApp.getActiveSpreadsheet().addMenu('株式銘柄チェック', [
      {name: '初期セットアップ', functionName: 'setupSheets'},
      {name: '手動でレポート送信', functionName: 'manualSendReport'},
      {name: 'メニュー再読み込み', functionName: 'reloadMenu'}
    ]);
  }
}

// メニューを手動で再読み込みする関数
function reloadMenu() {
  onOpen();
  SpreadsheetApp.getUi().alert('メニューを再読み込みしました。');
}

// インストール可能なトリガーとしてonOpenを設定する関数
function installOnOpenTrigger() {
  const ss = SpreadsheetApp.getActive();
  ScriptApp.newTrigger('onOpen')
    .forSpreadsheet(ss)
    .onOpen()
    .create();
  SpreadsheetApp.getUi().alert('onOpenトリガーを設定しました。スプレッドシートを再度開いてください。');
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
      ['news_per_company', '5'],
      ['send_time_jst', '18:00'],
      ['skip_holiday', 'TRUE'],
      ['timezone', 'Asia/Tokyo'],
      ['holiday_calendar_id', 'ja.japanese#holiday@group.v.calendar.google.com'],
      ['use_extractive_summary_only', 'TRUE'],
      ['perplexity_api_key', ''],
      ['perplexity_model', 'sonar-pro'],
      ['perplexity_web_citations', 'TRUE']
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
// API Key設定
// ====================================
function setPerplexityApiKey() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    'Perplexity API Key設定',
    'Perplexity API Keyを入力してください:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() === ui.Button.OK) {
    const apiKey = response.getResponseText().trim();
    if (apiKey) {
      // スクリプトプロパティに安全に保存
      PropertiesService.getScriptProperties().setProperty('PERPLEXITY_API_KEY', apiKey);
      ui.alert('API Keyを設定しました。');
    } else {
      ui.alert('API Keyが入力されていません。');
    }
  }
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

// 価格取得テスト関数
function testPriceRetrieval() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('価格取得テスト', '証券コードを入力してください（例: 9984）:', ui.ButtonSet.OK_CANCEL);
  
  if (response.getSelectedButton() === ui.Button.OK) {
    const code = response.getResponseText().trim();
    if (code) {
      ui.alert('価格取得テスト開始', `証券コード ${code} の価格を取得中...`, ui.ButtonSet.OK);
      
      try {
        const marketData = fetchMarketData(code);
        
        let resultText = `証券コード: ${code}\n\n`;
        resultText += `データソース: ${marketData.dataSource}\n`;
        resultText += `終値: ${marketData.close}円\n`;
        resultText += `前日比: ${marketData.change}円 (${marketData.changePercent.toFixed(2)}%)\n`;
        resultText += `高値: ${marketData.dayHigh}円\n`;
        resultText += `安値: ${marketData.dayLow}円\n`;
        
        // 発行済み株式総数と時価総額の表示
        if (marketData.sharesOutstanding > 0) {
          resultText += `発行済株式総数: ${marketData.sharesOutstanding.toLocaleString()}株\n`;
        }
        
        if (marketData.marketCap > 0) {
          const marketCapBillion = marketData.marketCap / 100000000;
          let marketCapDisplay = '';
          if (marketCapBillion >= 10000) {
            marketCapDisplay = `${(marketCapBillion / 10000).toFixed(2)}兆円`;
          } else {
            marketCapDisplay = `${marketCapBillion.toFixed(0).toLocaleString()}億円`;
          }
          resultText += `時価総額: ${marketCapDisplay}`;
          if (marketData.sharesOutstanding > 0) {
            resultText += ` (株価 × 発行済株式総数)`;
          }
          resultText += `\n`;
        } else {
          resultText += `時価総額: 計算できませんでした\n`;
        }
        resultText += `取得時刻: ${marketData.fetchTime || '不明'}\n\n`;
        
        // ログも確認
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const logsSheet = ss.getSheetByName('Logs');
        if (logsSheet) {
          resultText += '※詳細はLogsシートをご確認ください。';
        }
        
        ui.alert('価格取得成功', resultText, ui.ButtonSet.OK);
      } catch (error) {
        ui.alert('価格取得失敗', `エラー: ${error.toString()}\n\n詳細はLogsシートをご確認ください。`, ui.ButtonSet.OK);
      }
    }
  }
}

// 業界情報取得テスト関数
function testIndustryInfo() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('業界情報テスト', '証券コードを入力してください（例: 9984）:', ui.ButtonSet.OK_CANCEL);
  
  if (response.getSelectedButton() === ui.Button.OK) {
    const code = response.getResponseText().trim();
    if (code) {
      const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Tickers');
      
      if (!sheet) {
        ui.alert('エラー', 'Tickersシートが見つかりません', ui.ButtonSet.OK);
        return;
      }
      
      // 証券コードから企業名を取得
      const data = sheet.getDataRange().getValues();
      let companyName = '';
      for (let i = 1; i < data.length; i++) {
        if (data[i][0].toString() === code) {
          companyName = data[i][1];
          break;
        }
      }
      
      if (!companyName) {
        ui.alert('エラー', `証券コード ${code} が見つかりません`, ui.ButtonSet.OK);
        return;
      }
      
      ui.alert('テスト開始', `${code} ${companyName} の業界情報を取得しています...`, ui.ButtonSet.OK);
      
      try {
        const insights = fetchCompanyInsights(code, companyName);
        
        let resultText = `【${code} ${companyName}】\n\n`;
        resultText += `■業界: ${insights.industry.sector}\n\n`;
        resultText += `■主な収益源:\n${insights.industry.revenueSource}\n\n`;
        resultText += `■将来への投資分野:\n${insights.industry.futureInvestment}\n\n`;
        resultText += `■投資テーマ: ${insights.industry.themes}\n`;
        
        // 長いテキストの場合は、最初の1000文字だけ表示
        if (resultText.length > 1000) {
          resultText = resultText.substring(0, 1000) + '\n...(以下省略)';
        }
        
        ui.alert('業界情報取得結果', resultText, ui.ButtonSet.OK);
        
      } catch (error) {
        ui.alert('エラー', `業界情報の取得に失敗しました: ${error.toString()}`, ui.ButtonSet.OK);
      }
    }
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
// ユーティリティ関数
// ====================================
function cleanMarkdown(text) {
  if (!text) return text;
  
  // マークダウン記法を削除
  return text
    .replace(/\*\*\*([^*]+)\*\*\*/g, '$1')  // ***bold italic*** を削除
    .replace(/\*\*([^*]+)\*\*/g, '$1')      // **bold** を削除
    .replace(/\*([^*]+)\*/g, '$1')          // *italic* を削除
    .replace(/__([^_]+)__/g, '$1')          // __underline__ を削除
    .replace(/_([^_]+)_/g, '$1')            // _italic_ を削除
    .replace(/~~([^~]+)~~/g, '$1')          // ~~strikethrough~~ を削除
    .replace(/#{1,6}\s+/g, '')              // # heading を削除
    .replace(/`([^`]+)`/g, '$1')            // `code` を削除
    .replace(/\[([^\]]+)\]\([^)]+\)/g, '$1') // [link](url) をテキストのみに
    .replace(/^[*+-]\s+/gm, '')             // リストマーカーを削除
    .replace(/^\d+\.\s+/gm, '');            // 番号付きリストを削除
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
    .stock-header { font-size: 20px; font-weight: bold; color: #0066cc; margin-bottom: 15px; border-bottom: 2px solid #0066cc; padding-bottom: 5px; }
    .data-row { margin: 8px 0; }
    .label { font-weight: bold; color: #666; }
    .positive { color: #009900; }
    .negative { color: #cc0000; }
    .disclaimer { background: #fffacd; padding: 10px; margin-top: 20px; font-size: 12px; }
    .section-title { font-weight: bold; margin-top: 20px; margin-bottom: 10px; color: #333; font-size: 16px; border-bottom: 1px solid #ddd; padding-bottom: 3px; }
    strong { font-weight: bold; }
  </style>
</head>
<body>
  <div class="header">
    <h2>株式レポート ${today}</h2>
    <p>配信日時: ${now} (JST)<br>
    基準: 当日終値<br>
    データソース: Yahoo Finance・東証・各社IR情報等</p>
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
    データソース: Yahoo Finance、東証、各社IR情報、Perplexity AI等<br>
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
    <div class="stock-header">${ticker.code} ${ticker.companyName}</div>`;
  
  // 基本情報（時価総額を最初に表示）
  if (data.marketData && data.marketData.close > 0) {
    html += `
    <div class="section-title">基本情報</div>
    <div class="data-row" style="font-size: 16px; line-height: 1.8;">`;
    
    // 時価総額を目立つように表示
    const marketCapBillion = data.marketData.marketCap / 100000000;
    let marketCapDisplay = '';
    if (data.marketData.marketCap > 0) {
      if (marketCapBillion >= 10000) {
        marketCapDisplay = `${(marketCapBillion / 10000).toFixed(2)}兆円`;
      } else {
        marketCapDisplay = `${marketCapBillion.toFixed(0).toLocaleString()}億円`;
      }
    } else {
      marketCapDisplay = '計算中';
    }
    
    // 時価総額の計算式を含めて表示
    html += `<div style="background-color: #f0f8ff; padding: 10px; border-radius: 5px; margin-bottom: 10px;">
      <strong style="color: #0066cc; font-size: 18px;">時価総額: </strong>
      <span style="font-size: 20px; font-weight: bold; color: #333;">${marketCapDisplay}</span>`;
    
    // 発行済み株式総数を表示
    if (data.marketData.sharesOutstanding > 0) {
      html += `<div style="font-size: 14px; color: #666; margin-top: 5px;">
        株価 ${data.marketData.close.toLocaleString()}円 × 発行済株式総数 ${(data.marketData.sharesOutstanding / 1000000).toFixed(0).toLocaleString()}百万株
      </div>`;
    }
    
    html += `</div>`;
    
    // 現在値
    const changeClass = data.marketData.change >= 0 ? 'positive' : 'negative';
    html += `<strong style="color: #0066cc;">現在値:</strong> ${data.marketData.close.toLocaleString()}円 
      <span class="${changeClass}">(${data.marketData.change >= 0 ? '+' : ''}${data.marketData.change.toLocaleString()}円 / ${data.marketData.changePercent >= 0 ? '+' : ''}${data.marketData.changePercent.toFixed(2)}%)</span>
    </div>`;
  }
  
  // 業界・テーマ情報
  if (data.industryInfo) {
    html += `
    <div class="section-title">業界・投資テーマ</div>
    <div class="data-row" style="font-size: 14px; line-height: 1.6;">
      <strong style="color: #0066cc;">業界:</strong> ${cleanMarkdown(data.industryInfo.sector)}<br>`;
    
    // 主な収益源
    if (data.industryInfo.revenueSource && 
        data.industryInfo.revenueSource !== '情報なし' && 
        data.industryInfo.revenueSource !== '情報を取得中' &&
        data.industryInfo.revenueSource !== '情報を取得できませんでした') {
      const cleanRevenue = cleanMarkdown(data.industryInfo.revenueSource);
      html += `<strong style="color: #0066cc;">主な収益源:</strong> ${cleanRevenue.replace(/\n/g, '<br>')}<br>`;
    }
    
    // 将来への投資分野
    if (data.industryInfo.futureInvestment && 
        data.industryInfo.futureInvestment !== '情報なし' &&
        data.industryInfo.futureInvestment !== '情報を取得中' &&
        data.industryInfo.futureInvestment !== '情報を取得できませんでした') {
      const cleanInvestment = cleanMarkdown(data.industryInfo.futureInvestment);
      html += `<strong style="color: #0066cc;">将来への投資分野:</strong> ${cleanInvestment.replace(/\n/g, '<br>')}<br>`;
    }
    
    html += `<strong style="color: #0066cc;">投資テーマ:</strong> ${cleanMarkdown(data.industryInfo.themes)}
    </div>`;
  }
  
  // 日本・世界経済における位置付け
  if (data.economicPosition && data.economicPosition.japanRole !== '取得失敗') {
    html += `
    <div class="section-title">経済的位置付け</div>
    <div class="data-row" style="font-size: 14px; line-height: 1.6;">
      ${cleanMarkdown(data.economicPosition.japanRole).replace(/\n/g, '<br>')}
    </div>`;
  }
  
  // マクロ経済の影響
  if (data.macroImpact && data.macroImpact.analysis !== '取得失敗') {
    html += `
    <div class="section-title">マクロ経済環境の影響</div>
    <div class="data-row" style="font-size: 14px; line-height: 1.6;">
      ${cleanMarkdown(data.macroImpact.analysis).replace(/\n/g, '<br>')}
    </div>`;
    
    // 専門家の見解
    if (data.macroImpact.expertViews && data.macroImpact.expertViews.length > 0) {
      html += `
    <div style="margin-top: 10px;">
      <div style="font-weight: bold; color: #666; margin-bottom: 5px;">専門家・経営陣の見解：</div>`;
      
      for (const view of data.macroImpact.expertViews) {
        html += `
      <div style="margin: 8px 0; padding-left: 15px; border-left: 3px solid #009900;">
        <div style="font-style: italic; color: #333;">「${cleanMarkdown(view.quote)}」</div>
        <div style="font-size: 12px; color: #666; margin-top: 3px;">― ${cleanMarkdown(view.source)}</div>
      </div>`;
      }
      
      html += `</div>`;
    }
  }
  
  // 詳細市場データ
  if (data.marketData && data.marketData.close > 0 && data.marketData.dayHigh > 0) {
    html += `
    <div class="section-title">詳細市場データ</div>`;
    
    if (data.marketData.dayHigh > 0 && data.marketData.dayLow > 0) {
      html += `
    <div class="data-row">
      <span class="label">高値/安値:</span> ${data.marketData.dayHigh.toLocaleString()}円 / ${data.marketData.dayLow.toLocaleString()}円
      <span style="color: #666;"> (レンジ: ${((data.marketData.dayHigh / data.marketData.dayLow - 1) * 100).toFixed(1)}%)</span>
    </div>`;
    }
  }
  
  // 財務指標（投資判断に重要）
  if (data.fundamentals) {
    html += `
    <div class="section-title">財務指標・バリュエーション</div>`;
    
    if (data.fundamentals.per !== null) {
      html += `
    <div class="data-row">
      <span class="label">PER:</span> ${data.fundamentals.per}倍
      ${data.fundamentals.per < 15 ? '<span style="color: green;">（割安水準）</span>' : 
        data.fundamentals.per > 30 ? '<span style="color: orange;">（割高水準）</span>' : ''}
    </div>`;
    }
    
    if (data.fundamentals.pbr !== null) {
      html += `
    <div class="data-row">
      <span class="label">PBR:</span> ${data.fundamentals.pbr}倍
      ${data.fundamentals.pbr < 1 ? '<span style="color: green;">（解散価値割れ）</span>' : ''}
    </div>`;
    }
    
    if (data.fundamentals.dividendYield !== null) {
      html += `
    <div class="data-row">
      <span class="label">配当利回り:</span> ${data.fundamentals.dividendYield}%
      ${data.fundamentals.dividendYield > 3 ? '<span style="color: green;">（高配当）</span>' : ''}
    </div>`;
    }
    
    if (data.fundamentals.roe !== null) {
      html += `
    <div class="data-row">
      <span class="label">ROE:</span> ${data.fundamentals.roe}%
      ${data.fundamentals.roe > 10 ? '<span style="color: green;">（高収益性）</span>' : 
        data.fundamentals.roe < 5 ? '<span style="color: orange;">（改善余地あり）</span>' : ''}
    </div>`;
    }
    
    if (data.fundamentals.revenue || data.fundamentals.operatingProfit) {
      html += `
    <div class="data-row">
      ${data.fundamentals.revenue ? `<span class="label">売上高:</span> ${data.fundamentals.revenue}　` : ''}
      ${data.fundamentals.operatingProfit ? `<span class="label">営業利益:</span> ${data.fundamentals.operatingProfit}` : ''}
    </div>`;
    }
    
    html += `
    <div class="data-row" style="font-size: 12px; color: #666;">
      出典: <a href="${data.fundamentals.url}" target="_blank">${data.fundamentals.source}</a>
    </div>`;
  }
  
  // 経営方針と経営者の発言
  if (data.managementInsights && data.managementInsights.visionAndStrategy !== '取得失敗') {
    html += `
    <div class="section-title">経営方針・経営者メッセージ</div>`;
    
    // 経営ビジョン・戦略
    if (data.managementInsights.visionAndStrategy !== '情報なし') {
      html += `
    <div class="data-row" style="font-size: 14px; line-height: 1.6; margin-bottom: 10px;">
      ${cleanMarkdown(data.managementInsights.visionAndStrategy).replace(/\n/g, '<br>')}
    </div>`;
    }
    
    // 経営者の発言引用
    if (data.managementInsights.executiveQuotes && data.managementInsights.executiveQuotes.length > 0) {
      html += `
    <div style="margin-top: 10px;">
      <div style="font-weight: bold; color: #666; margin-bottom: 5px;">経営陣の重要発言：</div>`;
      
      for (const quote of data.managementInsights.executiveQuotes) {
        html += `
      <div style="margin: 8px 0; padding-left: 15px; border-left: 3px solid #0066cc;">
        <div style="font-style: italic; color: #333;">「${cleanMarkdown(quote.quote)}」</div>
        <div style="font-size: 12px; color: #666; margin-top: 3px;">― ${cleanMarkdown(quote.source)}</div>
      </div>`;
      }
      
      html += `</div>`;
    }
    
    // 参考URL
    if (data.managementInsights.referenceUrls && data.managementInsights.referenceUrls.length > 0) {
      html += `
    <div style="margin-top: 10px; font-size: 12px;">
      <span style="color: #666;">参考記事：</span>`;
      
      for (const url of data.managementInsights.referenceUrls.slice(0, 3)) {
        html += ` <a href="${url}" target="_blank" style="color: #0066cc;">[記事]</a>`;
      }
      
      html += `</div>`;
    }
  }
  
  // 組織課題・取り組み
  if (data.organizationalChallenges && data.organizationalChallenges.summary !== '取得失敗' && data.organizationalChallenges.summary !== '情報なし') {
    html += `
    <div class="section-title">直近の組織課題・取り組み</div>
    <div class="data-row" style="font-size: 14px; line-height: 1.6;">
      ${cleanMarkdown(data.organizationalChallenges.summary).replace(/\n/g, '<br>')}
    </div>`;
  }
  
  // ニュース（必須表示）
  html += `
    <div class="section-title">最新ニュース（過去3日間）</div>`;
  
  if (data.news && data.news.length > 0) {
    for (const item of data.news) {
      const importanceColor = item.importance === 'high' ? '#d00' : '#666';
      html += `
    <div class="data-row" style="margin-bottom: 10px;">
      • `;
      
      if (item.publishTime) {
        html += `<span style="color: ${importanceColor}; font-size: 12px;">[${item.publishTime}]</span> `;
      }
      
      html += `<a href="${item.url}" target="_blank" style="color: #0066cc; font-weight: bold;">${cleanMarkdown(item.title)}</a>`;
      
      if (item.summary) {
        html += `<br>
      <span style="font-size: 13px; color: #333; margin-left: 10px; display: block; margin-top: 3px;">
        ${cleanMarkdown(item.summary).replace(/\n/g, '<br>')}
      </span>`;
      }
      
      html += `</div>`;
    }
  } else {
    html += `
    <div class="data-row" style="color: #666;">
      ニュースが取得できませんでした。<a href="https://finance.yahoo.co.jp/quote/${ticker.code}.T" target="_blank">Yahoo Finance</a>で最新情報をご確認ください。
    </div>`;
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
    fundamentals: null,
    industryInfo: null,
    economicPosition: null,
    macroImpact: null,
    organizationalChallenges: null,
    managementInsights: null,
    news: []
  };
  
  // 市場データ取得（Yahoo Finance）
  try {
    data.marketData = fetchMarketData(ticker.code);
  } catch (error) {
    logMessage('WARN', `市場データ取得失敗: ${ticker.code}`, error.toString());
  }
  
  // ファンダメンタルズ（財務情報）取得
  try {
    data.fundamentals = fetchYahooFinanceComprehensive(ticker.code);
  } catch (error) {
    logMessage('WARN', `ファンダメンタルズ取得失敗: ${ticker.code}`, error.toString());
  }
  
  // 業界情報、経済的位置付け、組織課題をPerplexityで取得
  try {
    const companyInsights = fetchCompanyInsights(ticker.code, ticker.companyName);
    data.industryInfo = companyInsights.industry;
    data.economicPosition = companyInsights.economicPosition;
    data.macroImpact = companyInsights.macroImpact;
    data.organizationalChallenges = companyInsights.challenges;
    data.managementInsights = companyInsights.management;
  } catch (error) {
    logMessage('WARN', `企業インサイト取得失敗: ${ticker.code}`, error.toString());
  }
  
  // ニュース取得（必須）
  try {
    data.news = fetchComprehensiveNews(ticker.code, ticker.companyName);
    // ニュースが取得できない場合は最低限のニュースを生成
    if (!data.news || data.news.length === 0) {
      data.news = [{
        title: '最新ニュースの取得に失敗しました',
        url: `https://finance.yahoo.co.jp/quote/${ticker.code}.T`,
        summary: 'Yahoo Financeで最新情報をご確認ください',
        publishTime: Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm')
      }];
    }
  } catch (error) {
    logMessage('ERROR', `ニュース取得失敗: ${ticker.code}`, error.toString());
    // フォールバック
    data.news = [{
      title: 'ニュース取得エラー',
      url: `https://finance.yahoo.co.jp/quote/${ticker.code}.T`,
      summary: error.toString(),
      publishTime: Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm')
    }];
  }
  
  return data;
}

function fetchMarketData(code) {
  try {
    // 価格情報を取得
    const marketData = fetchYahooFinanceData(code);
    if (marketData && marketData.close > 0) {
      // 発行済み株式総数を取得
      const sharesOutstanding = fetchSharesOutstanding(code);
      
      // 時価総額を計算（発行済み株式総数 × 株価）
      if (sharesOutstanding > 0) {
        marketData.sharesOutstanding = sharesOutstanding;
        marketData.marketCap = marketData.close * sharesOutstanding;
        logMessage('SUCCESS', `時価総額計算: ${code}`, `株価: ${marketData.close}円 × 発行済株式数: ${sharesOutstanding.toLocaleString()}株 = ${(marketData.marketCap / 100000000).toFixed(0)}億円`);
      }
      
      logMessage('SUCCESS', `価格取得成功: ${code}`, `終値: ${marketData.close}円, ソース: ${marketData.dataSource}`);
      return marketData;
    }
  } catch (error) {
    logMessage('ERROR', `価格取得エラー: ${code}`, error.toString());
  }
  
  // フォールバック：ダミーデータ（テスト用）
  logMessage('WARN', `フォールバックデータ使用: ${code}`, 'リアルデータの取得に失敗したため、テストデータを使用');
  const testShares = 1000000000; // 10億株
  const testPrice = 3000 + Math.floor(Math.random() * 2000);
  return {
    close: testPrice,
    change: Math.floor((Math.random() - 0.5) * 200),
    changePercent: (Math.random() - 0.5) * 5,
    sharesOutstanding: testShares,
    marketCap: testPrice * testShares,
    previousClose: 3000,
    dayHigh: 3100,
    dayLow: 2900,
    dataSource: 'テストデータ（実データ取得失敗）',
    fetchTime: Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss')
  };
}

// 発行済み株式総数を取得
function fetchSharesOutstanding(code) {
  const settings = loadSettings();
  
  // まずPerplexityで取得を試みる
  try {
    const prompt = `
証券コード${code}の発行済株式総数を教えてください。

以下の形式で、数値のみを返してください：
発行済株式総数: [数値]株

最新の有価証券報告書や決算短信の情報を使用してください。
自己株式を除いた実質的な流通株式数があれば、それも併記してください。
`;
    
    const perplexityResult = callPerplexityAPI(prompt, settings);
    const content = perplexityResult.choices[0].message.content;
    
    // 発行済株式総数を抽出（複数のパターンに対応）
    const sharesPatterns = [
      /発行済株式総数[：:]*\s*([\d,\.]+)\s*(千株|百万株|億株|株)/,
      /発行済株式数[：:]*\s*([\d,\.]+)\s*(千株|百万株|億株|株)/,
      /流通株式数[：:]*\s*([\d,\.]+)\s*(千株|百万株|億株|株)/,
      /自己株式控除後[：:]*\s*([\d,\.]+)\s*(千株|百万株|億株|株)/
    ];
    
    for (const pattern of sharesPatterns) {
      const match = content.match(pattern);
      if (match && match[1]) {
        let shares = parseFloat(match[1].replace(/,/g, ''));
        // 単位に応じて株数に変換
        if (match[2]) {
          switch(match[2]) {
            case '億株': shares *= 100000000; break;
            case '百万株': shares *= 1000000; break;
            case '千株': shares *= 1000; break;
            default: break;
          }
        }
        if (shares > 0) {
          logMessage('SUCCESS', `発行済株式総数取得（Perplexity）: ${code}`, `${shares.toLocaleString()}株`);
          return shares;
        }
      }
    }
  } catch (error) {
    logMessage('WARN', `Perplexityで発行済株式総数取得失敗: ${code}`, error.toString());
  }
  
  // Yahoo FinanceのHTMLから取得を試みる
  try {
    const url = `https://finance.yahoo.co.jp/quote/${code}.T`;
    const response = UrlFetchApp.fetch(url, {
      muteHttpExceptions: true,
      headers: {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
      }
    });
    
    if (response.getResponseCode() === 200) {
      const html = response.getContentText();
      
      // 発行済株式数のパターン
      const sharesPatterns = [
        /発行済株式数[^>]*>([\d,\.]+)\s*(千株|百万株|億株|株)/,
        /発行済株式総数[^>]*>([\d,\.]+)\s*(千株|百万株|億株|株)/,
        /shares\s*outstanding[^>]*>([\d,\.]+)\s*(千株|百万株|億株|株)/i
      ];
      
      for (const pattern of sharesPatterns) {
        const match = html.match(pattern);
        if (match && match[1]) {
          let shares = parseFloat(match[1].replace(/,/g, ''));
          // 単位に応じて株数に変換
          switch(match[2]) {
            case '億株': shares *= 100000000; break;
            case '百万株': shares *= 1000000; break;
            case '千株': shares *= 1000; break;
            default: break;
          }
          if (shares > 0) {
            logMessage('SUCCESS', `発行済株式総数取得（Yahoo Finance）: ${code}`, `${shares.toLocaleString()}株`);
            return shares;
          }
        }
      }
    }
  } catch (error) {
    logMessage('WARN', `Yahoo Financeで発行済株式総数取得失敗: ${code}`, error.toString());
  }
  
  // 取得できなかった場合
  logMessage('WARN', `発行済株式総数を取得できませんでした: ${code}`);
  return 0;
}

function fetchYahooFinanceData(code) {
  // 複数のデータソースを試す
  const dataSources = [
    { name: 'Yahoo Finance API', func: () => fetchYahooFinanceAPI(code) },
    { name: 'Perplexity API', func: () => fetchStockPriceWithPerplexity(code) },
    { name: 'Yahoo Finance HTML', func: () => fetchYahooFinanceHTML(code) }
  ];
  
  for (const source of dataSources) {
    try {
      logMessage('INFO', `${source.name}から価格取得試行: ${code}`);
      const data = source.func();
      if (data && data.close > 0) {
        data.dataSource = source.name;
        logMessage('SUCCESS', `${source.name}で価格取得成功: ${code} - ${data.close}円`);
        return data;
      }
    } catch (error) {
      logMessage('WARN', `${source.name}取得失敗: ${code}`, error.toString());
    }
  }
  
  throw new Error('すべてのデータソースから価格情報を取得できませんでした');
}

// Yahoo Finance APIから取得
function fetchYahooFinanceAPI(code) {
  try {
    // Yahoo Finance APIエンドポイント（日本株）
    const symbol = `${code}.T`;
    const url = `https://query2.finance.yahoo.com/v8/finance/chart/${symbol}?interval=1d&range=1d`;
    
    const response = UrlFetchApp.fetch(url, {
      muteHttpExceptions: true,
      headers: {
        'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36',
        'Accept': 'application/json'
      }
    });
    
    if (response.getResponseCode() === 200) {
      const data = JSON.parse(response.getContentText());
      
      if (data.chart && data.chart.result && data.chart.result.length > 0) {
        const result = data.chart.result[0];
        const quote = result.indicators.quote[0];
        const meta = result.meta;
        
        // 最新の価格データ
        const timestamps = result.timestamp;
        const lastIndex = timestamps.length - 1;
        const currentPrice = quote.close[lastIndex] || meta.regularMarketPrice || 0;
        const previousClose = meta.previousClose || meta.chartPreviousClose || 0;
        
        if (currentPrice > 0) {
          return {
            close: currentPrice,
            change: currentPrice - previousClose,
            changePercent: previousClose > 0 ? ((currentPrice - previousClose) / previousClose) * 100 : 0,
            marketCap: meta.marketCap || 0,
            previousClose: previousClose,
            dayHigh: Math.max(...quote.high.filter(h => h !== null && h !== undefined)),
            dayLow: Math.min(...quote.low.filter(l => l !== null && l !== undefined)),
            fetchTime: Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss')
          };
        }
      }
    }
  } catch (error) {
    throw error;
  }
  
  return null;
}

// Perplexity APIを使用して価格取得
function fetchStockPriceWithPerplexity(code) {
  const settings = loadSettings();
  const prompt = `
証券コード${code}の本日の株価情報を教えてください。

以下の形式で、数値のみを返してください：
終値: [数値]
前日比: [数値]
前日比率: [数値]%
時価総額: [数値]億円（または兆円）
高値: [数値]
安値: [数値]

最新の市場データのみを提供してください。
`;
  
  try {
    const perplexityResult = callPerplexityAPI(prompt, settings);
    const content = perplexityResult.choices[0].message.content;
    
    // 価格情報を抽出
    const priceMatch = content.match(/終値[：:]*\s*([\d,\.]+)/);
    const changeMatch = content.match(/前日比[：:]*\s*([+-]?[\d,\.]+)/);
    const changePercentMatch = content.match(/前日比率[：:]*\s*([+-]?[\d\.]+)/);
    const highMatch = content.match(/高値[：:]*\s*([\d,\.]+)/);
    const lowMatch = content.match(/安値[：:]*\s*([\d,\.]+)/);
    
    // 時価総額を抽出
    let marketCap = 0;
    const marketCapMatch = content.match(/時価総額[：:]*\s*([\d,\.]+)\s*(億円|兆円)/);
    if (marketCapMatch) {
      let value = parseFloat(marketCapMatch[1].replace(/,/g, ''));
      if (marketCapMatch[2] === '兆円') {
        marketCap = value * 1000000000000;
      } else if (marketCapMatch[2] === '億円') {
        marketCap = value * 100000000;
      }
    }
    
    const currentPrice = priceMatch ? parseFloat(priceMatch[1].replace(/,/g, '')) : 0;
    const change = changeMatch ? parseFloat(changeMatch[1].replace(/,/g, '')) : 0;
    const changePercent = changePercentMatch ? parseFloat(changePercentMatch[1]) : 0;
    const dayHigh = highMatch ? parseFloat(highMatch[1].replace(/,/g, '')) : currentPrice;
    const dayLow = lowMatch ? parseFloat(lowMatch[1].replace(/,/g, '')) : currentPrice;
    
    if (currentPrice > 0) {
      return {
        close: currentPrice,
        change: change,
        changePercent: changePercent,
        marketCap: marketCap,
        previousClose: currentPrice - change,
        dayHigh: dayHigh,
        dayLow: dayLow,
        fetchTime: Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss')
      };
    }
  } catch (error) {
    throw error;
  }
  
  return null;
}

// Yahoo Finance HTMLから取得（フォールバック）
function fetchYahooFinanceHTML(code) {
  const url = `https://finance.yahoo.co.jp/quote/${code}.T`;
  
  try {
    const response = UrlFetchApp.fetch(url, {
      muteHttpExceptions: true,
      headers: {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
      }
    });
    
    if (response.getResponseCode() !== 200) {
      throw new Error('Yahoo Financeからのデータ取得に失敗');
    }
    
    const html = response.getContentText();
    
    // より確実な正規表現パターン
    // 現在値（複数のパターンを試す）
    let currentPrice = 0;
    const pricePatterns = [
      /<span class="stoksPrice">([\d,\.]+)<\/span>/,
      /<td class="stoksPrice">([\d,\.]+)<\/td>/,
      /"price"[^>]*>([\d,\.]+)</,
      /現在値[^>]*>([\d,\.]+)</
    ];
    
    for (const pattern of pricePatterns) {
      const match = html.match(pattern);
      if (match && match[1]) {
        currentPrice = parseFloat(match[1].replace(/,/g, ''));
        if (currentPrice > 0) break;
      }
    }
    
    // 時価総額を取得
    let marketCap = 0;
    const marketCapPatterns = [
      /時価総額[^>]*>([\d,\.]+)\s*(億円|兆円|万円|円)/,
      /時価総額.*?>([\d,\.]+)\s*(億円|兆円|万円|円)/,
      /market\s*cap[^>]*>([\d,\.]+)\s*(億円|兆円|万円|円)/i
    ];
    
    for (const pattern of marketCapPatterns) {
      const match = html.match(pattern);
      if (match && match[1] && match[2]) {
        let value = parseFloat(match[1].replace(/,/g, ''));
        // 単位に応じて円に変換
        switch(match[2]) {
          case '兆円': value *= 1000000000000; break;
          case '億円': value *= 100000000; break;
          case '万円': value *= 10000; break;
          default: break;
        }
        marketCap = value;
        if (marketCap > 0) break;
      }
    }
    
    // 発行済株式数から時価総額を計算（バックアップ）
    if (marketCap === 0 && currentPrice > 0) {
      const sharesPattern = /発行済株式数[^>]*>([\d,\.]+)\s*(千株|百万株|株)/;
      const sharesMatch = html.match(sharesPattern);
      if (sharesMatch && sharesMatch[1]) {
        let shares = parseFloat(sharesMatch[1].replace(/,/g, ''));
        if (sharesMatch[2] === '千株') shares *= 1000;
        else if (sharesMatch[2] === '百万株') shares *= 1000000;
        marketCap = currentPrice * shares;
      }
    }
    
    if (currentPrice > 0) {
      return {
        close: currentPrice,
        change: 0,
        changePercent: 0,
        marketCap: marketCap,
        previousClose: currentPrice,
        dayHigh: currentPrice,
        dayLow: currentPrice,
        fetchTime: Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss')
      };
    }
  } catch (error) {
    throw error;
  }
  
  return null;
}

// ====================================
// Perplexity API関連関数
// ====================================
function callPerplexityAPI(prompt, settings) {
  // スクリプトプロパティから取得を優先
  let apiKey = PropertiesService.getScriptProperties().getProperty('PERPLEXITY_API_KEY');
  
  // なければ設定シートから取得（後方互換性のため）
  if (!apiKey) {
    apiKey = settings.perplexity_api_key;
  }
  
  if (!apiKey) {
    throw new Error('Perplexity API Keyが設定されていません。メニューから「Perplexity API Key設定」を実行してください。');
  }
  
  const url = 'https://api.perplexity.ai/chat/completions';
  const payload = {
    model: settings.perplexity_model || 'sonar-small-online',
    messages: [
      {
        role: 'system',
        content: `あなたは日本の株式市場と企業情報の専門的な調査アシスタントです。
以下のルールを厳守してください：
1. 公式発表されたファクトとURLのみを提供する
2. 推測や予測は一切含めない
3. 信頼できる情報源（会社IR、取引所、主要メディア）のみを参照する
4. URLは必ず完全な形式で提供する
5. 情報が見つからない場合は「該当情報なし」と明確に回答する`
      },
      {
        role: 'user',
        content: prompt
      }
    ],
    max_tokens: 1000,
    temperature: 0.2,
    return_citations: settings.perplexity_web_citations === 'TRUE'
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
  const result = JSON.parse(response.getContentText());
  
  if (response.getResponseCode() !== 200) {
    throw new Error(`Perplexity API Error: ${result.error?.message || 'Unknown error'}`);
  }
  
  return result;
}

// Yahoo Financeから包括的な財務データを取得
function fetchYahooFinanceComprehensive(code) {
  try {
    const url = `https://finance.yahoo.co.jp/quote/${code}.T`;
    const response = UrlFetchApp.fetch(url, {
      muteHttpExceptions: true,
      headers: {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
      }
    });
    
    if (response.getResponseCode() === 200) {
      const html = response.getContentText();
      
      // 財務指標を抽出
      const perMatch = html.match(/PER[^>]*>([\\d,\\.]+)/);
      const pbrMatch = html.match(/PBR[^>]*>([\\d,\\.]+)/);
      const divYieldMatch = html.match(/配当利回り[^>]*>([\\d,\\.]+)%/);
      const revenueMatch = html.match(/売上高[^>]*>[^>]*>([\\d,\\.]+)(兆円|億円|百万円)/);
      const profitMatch = html.match(/営業利益[^>]*>[^>]*>([\\d,\\.]+)(兆円|億円|百万円)/);
      const epsMatch = html.match(/EPS[^>]*>([\\d,\\.]+)/);
      const roeMatch = html.match(/ROE[^>]*>([\\d,\\.]+)%/);
      
      const fundamentals = {
        per: perMatch ? parseFloat(perMatch[1]) : null,
        pbr: pbrMatch ? parseFloat(pbrMatch[1]) : null,
        dividendYield: divYieldMatch ? parseFloat(divYieldMatch[1]) : null,
        revenue: revenueMatch ? `${revenueMatch[1]}${revenueMatch[2]}` : null,
        operatingProfit: profitMatch ? `${profitMatch[1]}${profitMatch[2]}` : null,
        eps: epsMatch ? parseFloat(epsMatch[1]) : null,
        roe: roeMatch ? parseFloat(roeMatch[1]) : null,
        source: 'Yahoo Finance',
        url: url,
        fetchTime: Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss')
      };
      
      return fundamentals;
    }
  } catch (error) {
    logMessage('ERROR', `Yahoo Finance包括データ取得失敗: ${code}`, error.toString());
  }
  
  return null;
}

// 企業の業界情報と組織課題を取得
function fetchCompanyInsights(code, companyName) {
  const settings = loadSettings();
  
  const prompt = `
証券コード${code} ${companyName}について、公開情報（有価証券報告書、決算説明資料、会社ホームページ等）に基づいて以下の情報を教えてください：

【業界情報】
業界セクター: ${companyName}が属する主要な業界分類を記載（例：情報通信業、小売業、製造業等）

【主な収益源】
主な収益源: 現在の売上高の内訳と各事業の構成比を具体的に記載してください。必ず以下の形式で：
- 事業名1: 売上高○○億円（構成比○○％）
- 事業名2: 売上高○○億円（構成比○○％）
※最新の決算資料やセグメント情報から正確な数値を引用してください

【将来への投資分野】
将来への投資分野: 以下の観点から具体的に記載：
- 研究開発投資：重点分野と投資額
- 新規事業：具体的な事業内容と投資規模
- 設備投資：投資対象と金額
- M&A・提携：方向性と想定規模

【投資テーマ】
投資テーマ: 関連する市場テーマを列挙（AI、脱炭素、DX、メタバース、半導体、EV、バイオ、宇宙、量子等）

【その他の項目】
2. 日本経済・世界経済における位置付け
3. マクロ経済の影響分析（為替、金利、インフレ、地政学的リスク）
4. 経営方針と経営者の発言（必ず発言者の役職名と発言日、出典URLを併記）
5. 直近の組織課題と取り組み

重要事項：
- 必ず公開されている具体的な数値やファクトを使用してください
- 推測や一般論ではなく、その企業固有の情報を記載してください
- 回答にマークダウン記法（**、*、#など）は使用しないでください
- 各項目は必ず上記のラベル（【業界情報】等）を付けて回答してください
`;
  
  try {
    const perplexityResult = callPerplexityAPI(prompt, settings);
    const content = perplexityResult.choices[0].message.content;
    
    // デバッグ用ログ
    logMessage('DEBUG', `Perplexity応答（${companyName}）`, content.substring(0, 500) + '...');
    
    // 経営者発言の抽出
    const executiveQuotes = [];
    const quoteRegex = /「([^」]+)」.*?（([^）]+)）/g;
    let match;
    while ((match = quoteRegex.exec(content)) !== null) {
      executiveQuotes.push({
        quote: match[1],
        source: match[2]
      });
    }
    
    // URL抽出
    const urlRegex = /https?:\/\/[^\s<>"{}|\\^`\[\]]+/g;
    const urls = content.match(urlRegex) || [];
    
    // 情報を構造化 - 【】で囲まれたセクションを確実に抽出
    const extractSection = (sectionName) => {
      const regex = new RegExp(`【${sectionName}】[\\s\\S]*?(?=【|$)`, 'g');
      const match = content.match(regex);
      if (match && match[0]) {
        // セクション名を除いた内容を返す
        return match[0].replace(new RegExp(`【${sectionName}】\\s*`), '').trim();
      }
      return null;
    };
    
    // 各セクションを抽出
    const industrySection = extractSection('業界情報');
    const revenueSection = extractSection('主な収益源');
    const investmentSection = extractSection('将来への投資分野');
    const themeSection = extractSection('投資テーマ');
    
    // 業界セクターを抽出（「業界セクター:」の後の内容）
    let sectorName = '情報なし';
    if (industrySection) {
      const sectorMatch = industrySection.match(/業界セクター[：:]\s*([^\n]+)/);
      sectorName = sectorMatch ? sectorMatch[1].trim() : industrySection.split('\n')[0].trim();
    }
    
    // 主な収益源を抽出（「主な収益源:」の後の内容）
    let revenueSource = '情報なし';
    if (revenueSection) {
      const revenueMatch = revenueSection.match(/主な収益源[：:]\s*([\s\S]+)/);
      if (revenueMatch) {
        revenueSource = revenueMatch[1].trim();
      } else {
        // 「主な収益源:」がない場合は全体を使用
        revenueSource = revenueSection;
      }
    }
    
    // 将来への投資分野を抽出
    let futureInvestment = '情報なし';
    if (investmentSection) {
      const investMatch = investmentSection.match(/将来への投資分野[：:]\s*([\s\S]+)/);
      if (investMatch) {
        futureInvestment = investMatch[1].trim();
      } else {
        futureInvestment = investmentSection;
      }
    }
    
    // 投資テーマを抽出
    let themes = '情報なし';
    if (themeSection) {
      const themeMatch = themeSection.match(/投資テーマ[：:]\s*([^\n]+)/);
      themes = themeMatch ? themeMatch[1].trim() : themeSection.split('\n')[0].trim();
    }
    
    // その他のセクションも同様に抽出
    const challengesSection = content.match(/組織課題[\s\S]*?(?=【|\d\.|$)/);
    const managementSection = content.match(/経営方針[\s\S]*?(?=【|\d\.|$)/);
    const economicPositionSection = content.match(/日本経済・世界経済[\s\S]*?(?=【|\d\.|$)/);
    const macroImpactSection = content.match(/マクロ経済[\s\S]*?(?=【|\d\.|$)/);
    
    return {
      industry: {
        sector: sectorName,
        revenueSource: revenueSource,
        futureInvestment: futureInvestment,
        themes: themes,
        fullDescription: content
      },
      economicPosition: {
        japanRole: economicPositionSection ? economicPositionSection[0].replace(/^[^\n]*\n/, '').trim() : '情報なし',
        globalPosition: '',
        fullDescription: content
      },
      macroImpact: {
        analysis: macroImpactSection ? macroImpactSection[0].replace(/^[^\n]*\n/, '').trim() : '情報なし',
        expertViews: executiveQuotes.filter(q => q.quote.includes('為替') || q.quote.includes('金利') || q.quote.includes('経済')),
        fullDescription: content
      },
      challenges: {
        summary: challengesSection ? challengesSection[0].replace(/^[^\n]*\n/, '').trim() : '情報なし',
        fullDescription: content
      },
      management: {
        executiveQuotes: executiveQuotes,
        visionAndStrategy: managementSection ? managementSection[0].replace(/^[^\n]*\n/, '').trim() : '情報なし',
        referenceUrls: urls,
        fullDescription: content
      }
    };
  } catch (error) {
    logMessage('WARN', `企業インサイト取得失敗: ${code}`, error.toString());
    return {
      industry: { 
        sector: '情報を取得できませんでした', 
        revenueSource: '情報を取得できませんでした',
        futureInvestment: '情報を取得できませんでした',
        themes: '情報を取得できませんでした', 
        fullDescription: '' 
      },
      economicPosition: { japanRole: '取得失敗', globalPosition: '取得失敗', fullDescription: '' },
      macroImpact: { analysis: '取得失敗', expertViews: [], fullDescription: '' },
      challenges: { summary: '取得失敗', fullDescription: '' },
      management: { executiveQuotes: [], visionAndStrategy: '取得失敗', referenceUrls: [], fullDescription: '' }
    };
  }
}

// 包括的なニュース取得（Yahoo Finance + Perplexity）
function fetchComprehensiveNews(code, companyName) {
  const settings = loadSettings();
  const newsLimit = parseInt(settings.news_per_company) || 5;
  
  const now = new Date();
  const threeDaysAgo = new Date(now);
  threeDaysAgo.setDate(threeDaysAgo.getDate() - 3);
  
  const prompt = `
証券コード${code} ${companyName}の過去3日間の重要ニュースを探しています。

以下の観点で最重要ニュースを${newsLimit}件選んでください：
1. 株価に大きな影響を与える可能性があるもの
2. 業績・財務に関連するもの
3. 経営戦略・M&A・提携に関するもの
4. 新製品・新事業に関するもの
5. 規制・法的問題に関するもの

各ニュースについて以下の形式で提供してください：
[日時] タイトル
URL: [ニュースのURL]
要約: [株価への影響を含む2-3文の要約]

情報源：日経新聞、ロイター、ブルームバーグ、Yahoo Finance、会社IR、東洋経済、ダイヤモンド等の信頼できるメディア

注意：回答にマークダウン記法（**、*、#など）は一切使用しないでください。
`;
  
  try {
    const perplexityResult = callPerplexityAPI(prompt, settings);
    const content = perplexityResult.choices[0].message.content;
    
    const newsItems = [];
    const blocks = content.split(/\n\n+/);
    
    for (const block of blocks) {
      const lines = block.trim().split('\n');
      if (lines.length >= 2) {
        const titleLine = lines[0];
        const urlLine = lines.find(l => l.includes('URL:')) || '';
        const summaryLine = lines.find(l => l.includes('要約:')) || '';
        
        const timeMatch = titleLine.match(/\[([^\]]+)\]/);
        const titleMatch = titleLine.replace(/\[[^\]]+\]/, '').trim();
        const urlMatch = urlLine.match(/https?:\/\/[^\s]+/);
        const summaryMatch = summaryLine.replace(/要約[:：]\s*/, '');
        
        if (titleMatch) {
          newsItems.push({
            title: titleMatch,
            url: urlMatch ? urlMatch[0] : `https://finance.yahoo.co.jp/quote/${code}.T`,
            summary: summaryMatch || '',
            publishTime: timeMatch ? timeMatch[1] : '',
            importance: 'high'
          });
        }
        
        if (newsItems.length >= newsLimit) break;
      }
    }
    
    // Yahoo Financeからも補完
    if (newsItems.length < newsLimit) {
      try {
        const yahooNews = fetchYahooFinanceNews(code);
        for (const news of yahooNews) {
          if (newsItems.length >= newsLimit) break;
          newsItems.push(news);
        }
      } catch (error) {
        logMessage('WARN', 'Yahoo Financeニュース取得失敗', error.toString());
      }
    }
    
    return newsItems;
  } catch (error) {
    logMessage('ERROR', `包括的ニュース取得失敗: ${code}`, error.toString());
    // フォールバック：Yahoo Financeのみ
    return fetchYahooFinanceNews(code);
  }
}

// Yahoo Financeからニュース取得
function fetchYahooFinanceNews(code) {
  const url = `https://finance.yahoo.co.jp/quote/${code}.T/news`;
  const newsItems = [];
  
  try {
    const response = UrlFetchApp.fetch(url, {
      muteHttpExceptions: true,
      headers: {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
      }
    });
    
    if (response.getResponseCode() === 200) {
      // 簡易的な実装（実際にはより詳細なパースが必要）
      newsItems.push({
        title: `${code}の最新ニュース`,
        url: url,
        summary: 'Yahoo Financeで詳細をご確認ください',
        publishTime: Utilities.formatDate(new Date(), 'Asia/Tokyo', 'MM/dd HH:mm'),
        importance: 'medium'
      });
    }
  } catch (error) {
    logMessage('WARN', `Yahoo Financeニュース取得エラー: ${code}`, error.toString());
  }
  
  return newsItems;
}

// Yahoo Financeから財務データを取得
function fetchYahooFinancialData(code) {
  try {
    const url = `https://finance.yahoo.co.jp/quote/${code}.T`;
    const response = UrlFetchApp.fetch(url, {
      muteHttpExceptions: true,
      headers: {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
      }
    });
    
    if (response.getResponseCode() === 200) {
      const html = response.getContentText();
      
      // 財務情報を抽出（簡易実装）
      const revenueMatch = html.match(/売上高[^>]*>[^>]*>([\\d,\\.]+)(\u5104\u5186|\u5146\u5186)/);
      const profitMatch = html.match(/営業利益[^>]*>[^>]*>([\\d,\\.]+)(\u5104\u5186|\u5146\u5186)/);
      
      if (revenueMatch || profitMatch) {
        return {
          revenue: revenueMatch ? `${revenueMatch[1]}${revenueMatch[2]}` : '情報なし',
          operatingProfit: profitMatch ? `${profitMatch[1]}${profitMatch[2]}` : '情報なし',
          source: 'Yahoo Finance',
          url: url,
          date: Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd')
        };
      }
    }
  } catch (error) {
    logMessage('WARN', `Yahoo Finance財務データ取得失敗: ${code}`, error.toString());
  }
  return null;
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

