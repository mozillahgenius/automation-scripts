/**
 * HubSpot Deal Status Reminder - 改良版
 * 成約前の案件で3日間更新されていないものを通知するGoogle Apps Script
 * レポートスプレッドシート作成機能付き
 */

// ==================== 設定値 ====================
const CONFIG = {
  // 通知先メールアドレス（必須）
  EMAIL_RECIPIENTS: 'user1@example.com,user2@example.com,admin@example.com',
  
  // スプレッドシートのシート名（正確に入力）
  SHEETS: {
    DEALS: 'HS/deals/09Sep',        // 案件データのシート名
    STATUS_MASTER: 'StatusMaster'   // ステータスマスターのシート名
  },
  
  // 列のインデックス（0から開始）
  COLUMNS: {
    DEAL_NAME: 0,        // A列: Deal Name
    AMOUNT: 1,           // B列: Amount
    DEAL_STAGE: 2,       // C列: Deal Stage
    CREATE_DATE: 3,      // D列: Create Date
    CLOSE_DATE: 4,       // E列: Close Date
    DEAL_OWNER: 5,       // F列: Deal Owner
    LAST_MODIFIED: 6     // G列: Last Modified Date
  },
  
  // 更新日数の閾値
  DAYS_THRESHOLD: 3
};

// ==================== メイン関数 ====================

/**
 * メニュー作成（自動実行）
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('HubSpot通知')
    .addItem('今すぐチェック', 'checkAndNotifyStaleDealsManual')
    .addItem('データを診断', 'debugDataCheck')
    .addItem('成約前ステータスを確認', 'checkPreContractStatuses')
    .addItem('テストメール送信', 'testEmailNotification')
    .addItem('実行ログを確認', 'showExecutionLog')
    .addToUi();
}

/**
 * メイン実行関数（自動実行対応）
 */
function checkAndNotifyStaleDeals() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    
    // 成約前ステータスと失注ステータスの取得
    const { preContractStatuses, excludedStatuses } = getPreContractStatuses(spreadsheet);
    console.log('成約前ステータス数:', preContractStatuses.length);
    console.log('失注ステータス数:', excludedStatuses.length);
    
    // 3日間更新されていない成約前案件の取得（失注は除外）
    const staleDeals = getStaleDeals(spreadsheet, preContractStatuses, excludedStatuses);
    console.log('未更新案件数:', staleDeals.length);
    
    // 案件が存在する場合、レポートスプレッドシート作成とメール通知を送信
    if (staleDeals.length > 0) {
      const reportUrl = createStaleDealsReportSpreadsheet(staleDeals);
      sendEmailNotification(staleDeals, reportUrl);
      console.log(`通知完了: ${staleDeals.length}件の未更新案件を通知しました`);
    } else {
      console.log('確認完了: 未更新案件はありません');
    }
    
    // 実行ログを記録
    logExecution(staleDeals.length);
    
  } catch (error) {
    console.error('エラーが発生しました:', error);
    sendErrorNotification(error);
  }
}

/**
 * 手動実行用関数（UIアラート付き）
 */
function checkAndNotifyStaleDealsManual() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const ui = SpreadsheetApp.getUi();
    
    // 成約前ステータスと失注ステータスの取得
    const { preContractStatuses, excludedStatuses } = getPreContractStatuses(spreadsheet);
    console.log('成約前ステータス数:', preContractStatuses.length);
    console.log('失注ステータス数:', excludedStatuses.length);
    
    // 3日間更新されていない成約前案件の取得（失注は除外）
    const staleDeals = getStaleDeals(spreadsheet, preContractStatuses, excludedStatuses);
    console.log('未更新案件数:', staleDeals.length);
    
    // 案件が存在する場合、レポートスプレッドシート作成とメール通知を送信
    if (staleDeals.length > 0) {
      const reportUrl = createStaleDealsReportSpreadsheet(staleDeals);
      sendEmailNotification(staleDeals, reportUrl);
      ui.alert('通知完了', `${staleDeals.length}件の未更新案件を通知しました\n\nレポート: ${reportUrl}`, ui.ButtonSet.OK);
    } else {
      ui.alert('確認完了', '未更新案件はありません', ui.ButtonSet.OK);
    }
    
  } catch (error) {
    console.error('エラーが発生しました:', error);
    SpreadsheetApp.getUi().alert('エラー', `エラーが発生しました: ${error}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * 成約前のステータス一覧を取得（失注は除外）
 */
function getPreContractStatuses(spreadsheet) {
  const statusMasterSheet = spreadsheet.getSheetByName(CONFIG.SHEETS.STATUS_MASTER);
  
  if (!statusMasterSheet) {
    throw new Error(`StatusMasterシート「${CONFIG.SHEETS.STATUS_MASTER}」が見つかりません`);
  }
  
  const statusData = statusMasterSheet.getDataRange().getValues();
  const preContractStatuses = [];
  const excludedStatuses = [];  // 失注ステータス用
  
  // ヘッダー行をスキップして処理
  for (let i = 1; i < statusData.length; i++) {
    const category = statusData[i][0]; // A列: 成約前/成約後/失注
    const status = statusData[i][1];   // B列: ステータス名
    
    if (category === '成約前' && status) {
      preContractStatuses.push(status);
    } else if (category === '失注' && status) {
      excludedStatuses.push(status);
      console.log('失注ステータスを除外:', status);
    }
  }
  
  return { preContractStatuses, excludedStatuses };
}

/**
 * 3日間更新されていない成約前案件を取得（失注は除外）
 */
function getStaleDeals(spreadsheet, preContractStatuses, excludedStatuses = []) {
  const dealsSheet = spreadsheet.getSheetByName(CONFIG.SHEETS.DEALS);
  
  if (!dealsSheet) {
    const sheets = spreadsheet.getSheets();
    const sheetNames = sheets.map(s => s.getName()).join(', ');
    throw new Error(`シート「${CONFIG.SHEETS.DEALS}」が見つかりません。存在するシート: ${sheetNames}`);
  }
  
  const dealsData = dealsSheet.getDataRange().getValues();
  const staleDeals = [];
  const excludedDeals = []; // 失注案件を別途カウント
  const now = new Date();
  const thresholdDate = new Date(now.getTime() - (CONFIG.DAYS_THRESHOLD * 24 * 60 * 60 * 1000));
  
  // ヘッダー行をスキップして処理
  for (let i = 1; i < dealsData.length; i++) {
    const dealStage = dealsData[i][CONFIG.COLUMNS.DEAL_STAGE];
    const lastModified = dealsData[i][CONFIG.COLUMNS.LAST_MODIFIED];
    
    // 失注ステータスの場合は別途カウントしてスキップ
    if (excludedStatuses.includes(dealStage)) {
      if (lastModified) {
        const lastModifiedDate = new Date(lastModified);
        if (lastModifiedDate < thresholdDate) {
          excludedDeals.push(dealsData[i][CONFIG.COLUMNS.DEAL_NAME]);
        }
      }
      console.log(`失注案件をスキップ: ${dealsData[i][CONFIG.COLUMNS.DEAL_NAME]} (${dealStage})`);
      continue;
    }
    
    // 成約前ステータスかつ最終更新日が閾値を超えている場合
    if (preContractStatuses.includes(dealStage) && lastModified) {
      const lastModifiedDate = new Date(lastModified);
      
      if (lastModifiedDate < thresholdDate) {
        const daysSinceUpdate = Math.floor((now - lastModifiedDate) / (24 * 60 * 60 * 1000));
        
        staleDeals.push({
          dealName: dealsData[i][CONFIG.COLUMNS.DEAL_NAME],
          amount: dealsData[i][CONFIG.COLUMNS.AMOUNT],
          dealStage: dealStage,
          createDate: formatDate(dealsData[i][CONFIG.COLUMNS.CREATE_DATE]),
          closeDate: formatDate(dealsData[i][CONFIG.COLUMNS.CLOSE_DATE]),
          dealOwner: dealsData[i][CONFIG.COLUMNS.DEAL_OWNER] || '未割当',
          lastModified: formatDate(lastModifiedDate),
          daysSinceUpdate: daysSinceUpdate,
          rowNumber: i + 1  // スプレッドシート上の行番号
        });
      }
    }
  }
  
  // 更新日数の多い順にソート
  staleDeals.sort((a, b) => b.daysSinceUpdate - a.daysSinceUpdate);
  
  console.log(`失注案件（${CONFIG.DAYS_THRESHOLD}日以上未更新）: ${excludedDeals.length}件（通知対象外）`);
  
  return staleDeals;
}

// ==================== レポート作成関数 ====================

/**
 * 未更新案件のレポート用スプレッドシートを新規作成
 */
function createStaleDealsReportSpreadsheet(staleDeals) {
  // 現在の日時を取得
  const now = new Date();
  const dateStr = Utilities.formatDate(now, 'JST', 'yyyy-MM-dd HH:mm');
  const fileName = `未更新案件レポート_${Utilities.formatDate(now, 'JST', 'yyyyMMdd_HHmmss')}`;
  
  // 新規スプレッドシートを作成
  const newSpreadsheet = SpreadsheetApp.create(fileName);
  
  // スプレッドシートのアクセス権限を設定（URLを知っている全員が編集可能）
  const file = DriveApp.getFileById(newSpreadsheet.getId());
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.EDIT);
  
  const sheet = newSpreadsheet.getActiveSheet();
  sheet.setName('未更新案件一覧');
  
  // ヘッダー行を設定
  const headers = [
    '優先度',
    '案件名',
    'ステータス',
    '担当者',
    '金額',
    '未更新日数',
    '最終更新日',
    '作成日',
    '成約予定日'
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // ヘッダー行のスタイル設定
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#667eea');
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');
  headerRange.setHorizontalAlignment('center');
  
  // データを追加
  const dataRows = staleDeals.map(deal => {
    let priority = '通常';
    if (deal.daysSinceUpdate >= 7) {
      priority = '🔴 緊急';
    } else if (deal.daysSinceUpdate >= 5) {
      priority = '🟡 要注意';
    } else {
      priority = '🟢 通常';
    }
    
    return [
      priority,
      deal.dealName,
      deal.dealStage,
      deal.dealOwner,
      deal.amount || 0,
      deal.daysSinceUpdate,
      deal.lastModified,
      deal.createDate,
      deal.closeDate
    ];
  });
  
  if (dataRows.length > 0) {
    sheet.getRange(2, 1, dataRows.length, headers.length).setValues(dataRows);
    
    // 優先度によって行の背景色を設定
    for (let i = 0; i < dataRows.length; i++) {
      const row = i + 2;
      const daysSince = staleDeals[i].daysSinceUpdate;
      let bgColor = '#ffffff';
      
      if (daysSince >= 7) {
        bgColor = '#ffebee'; // 赤系（緊急）
      } else if (daysSince >= 5) {
        bgColor = '#fff3cd'; // 黄系（要注意）
      } else {
        bgColor = '#f1f8e9'; // 緑系（通常）
      }
      
      sheet.getRange(row, 1, 1, headers.length).setBackground(bgColor);
    }
    
    // 金額列に通貨形式を適用
    sheet.getRange(2, 5, dataRows.length, 1).setNumberFormat('¥#,##0');
  }
  
  // 列幅を自動調整
  sheet.autoResizeColumns(1, headers.length);
  
  // サマリーシートを追加
  const summarySheet = newSpreadsheet.insertSheet('サマリー');
  
  // ステータス別集計を作成
  const statusSummary = createStatusSummary(staleDeals);
  
  // サマリー情報を作成
  const summaryData = [
    ['レポート作成日時', dateStr],
    ['', ''],
    ['【概要】', ''],
    ['未更新案件総数', staleDeals.length],
    ['対象金額合計', staleDeals.reduce((sum, d) => sum + (d.amount || 0), 0)],
    ['', ''],
    ['【優先度別内訳】', ''],
    ['緊急（7日以上）', staleDeals.filter(d => d.daysSinceUpdate >= 7).length],
    ['要注意（5-6日）', staleDeals.filter(d => d.daysSinceUpdate >= 5 && d.daysSinceUpdate < 7).length],
    ['通常（3-4日）', staleDeals.filter(d => d.daysSinceUpdate < 5).length],
    ['', ''],
    ['【ステータス別集計】', '']
  ];
  
  // ステータス別の詳細を追加
  Object.keys(statusSummary).sort((a, b) => statusSummary[b].count - statusSummary[a].count).forEach(status => {
    const stat = statusSummary[status];
    summaryData.push([status, `${stat.count}件 (平均${stat.avgDaysStale}日停滞)`]);
  });
  
  summarySheet.getRange(1, 1, summaryData.length, 2).setValues(summaryData);
  
  // サマリーシートのスタイル設定
  summarySheet.getRange('A3').setFontWeight('bold');
  summarySheet.getRange('A7').setFontWeight('bold');
  summarySheet.getRange('A12').setFontWeight('bold');
  summarySheet.getRange('A1:B1').setBackground('#e8f0fe');
  summarySheet.autoResizeColumns(1, 2);
  
  // 金額のフォーマット
  summarySheet.getRange('B5').setNumberFormat('¥#,##0');
  
  // スプレッドシートのURLを返す
  return newSpreadsheet.getUrl();
}

// ==================== メール送信関数 ====================

/**
 * HTMLメール通知を送信
 */
function sendEmailNotification(staleDeals, reportUrl) {
  const subject = `[HubSpot] ${CONFIG.DAYS_THRESHOLD}日以上更新されていない成約前案件: ${staleDeals.length}件`;
  
  // ステータス別集計を作成
  const statusSummary = createStatusSummary(staleDeals);
  
  const htmlBody = createSimplifiedHtmlEmailBody(staleDeals, statusSummary, reportUrl);
  const plainTextBody = createSimplifiedPlainTextEmailBody(staleDeals, statusSummary, reportUrl);
  
  MailApp.sendEmail({
    to: CONFIG.EMAIL_RECIPIENTS,
    subject: subject,
    body: plainTextBody,
    htmlBody: htmlBody,
    noReply: true
  });
}

/**
 * ステータス別の集計を作成
 */
function createStatusSummary(staleDeals) {
  const summary = {};
  
  staleDeals.forEach(deal => {
    if (!summary[deal.dealStage]) {
      summary[deal.dealStage] = {
        count: 0,
        totalAmount: 0,
        avgDaysStale: 0,
        deals: []
      };
    }
    
    summary[deal.dealStage].count++;
    summary[deal.dealStage].totalAmount += (deal.amount || 0);
    summary[deal.dealStage].avgDaysStale += deal.daysSinceUpdate;
    summary[deal.dealStage].deals.push(deal);
  });
  
  // 平均日数を計算
  Object.keys(summary).forEach(status => {
    summary[status].avgDaysStale = Math.round(summary[status].avgDaysStale / summary[status].count);
  });
  
  return summary;
}

/**
 * 簡略化されたHTML形式のメール本文を作成
 */
function createSimplifiedHtmlEmailBody(staleDeals, statusSummary, reportUrl) {
  const totalDeals = staleDeals.length;
  const urgentDeals = staleDeals.filter(d => d.daysSinceUpdate >= 7).length;
  const warningDeals = staleDeals.filter(d => d.daysSinceUpdate >= 5 && d.daysSinceUpdate < 7).length;
  const totalAmount = staleDeals.reduce((sum, d) => sum + (d.amount || 0), 0);
  
  let html = `
<!DOCTYPE html>
<html>
<head>
  <style>
    body { font-family: Arial, sans-serif; line-height: 1.6; color: #333; }
    .container { max-width: 600px; margin: 0 auto; background: #fff; padding: 20px; }
    .header { background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; padding: 25px; text-align: center; border-radius: 5px; margin-bottom: 20px; }
    .header h1 { margin: 0; font-size: 24px; }
    .header p { margin: 10px 0 0; opacity: 0.9; font-size: 14px; }
    .summary { display: flex; justify-content: space-around; margin: 20px 0; padding: 20px; background: #f8f9fa; border-radius: 5px; }
    .summary-item { text-align: center; }
    .summary-label { font-size: 12px; color: #6c757d; margin-top: 5px; }
    .summary-number { font-size: 28px; font-weight: bold; }
    .urgent { color: #dc3545; }
    .warning { color: #ffc107; }
    .normal { color: #28a745; }
    .report-button { display: block; background: #007bff; color: white; text-decoration: none; padding: 15px 30px; border-radius: 5px; text-align: center; font-weight: bold; margin: 25px 0; }
    .report-button:hover { background: #0056b3; }
    .status-summary { margin: 20px 0; padding: 15px; background: #f8f9fa; border-left: 4px solid #6c757d; }
    .status-summary h4 { margin: 0 0 10px 0; color: #495057; font-size: 16px; }
    .status-row { display: flex; justify-content: space-between; padding: 5px 0; border-bottom: 1px solid #e9ecef; font-size: 14px; }
    .status-name { color: #495057; }
    .status-count { color: #6c757d; font-weight: bold; }
    .action-box { margin: 20px 0; padding: 15px; background: #e3f2fd; border-left: 4px solid #2196f3; }
    .action-box h4 { margin: 0 0 10px 0; color: #1976d2; font-size: 16px; }
    .action-box ol { margin: 10px 0; padding-left: 20px; font-size: 14px; }
    .footer { margin-top: 30px; padding-top: 20px; border-top: 1px solid #dee2e6; text-align: center; color: #6c757d; font-size: 12px; }
  </style>
</head>
<body>
  <div class="container">
    <div class="header">
      <h1>📊 HubSpot 案件更新アラート</h1>
      <p>成約前案件の更新状況チェック</p>
    </div>
    
    <div class="summary">
      <div class="summary-item">
        <div class="summary-number urgent">${totalDeals}</div>
        <div class="summary-label">未更新案件</div>
      </div>
      <div class="summary-item">
        <div class="summary-number urgent">${urgentDeals}</div>
        <div class="summary-label">緊急（7日以上）</div>
      </div>
      <div class="summary-item">
        <div class="summary-number warning">${warningDeals}</div>
        <div class="summary-label">要注意（5-6日）</div>
      </div>
    </div>
    
    <div class="status-summary">
      <h4>ステータス別集計（上位5件）</h4>`;
  
  // ステータス別集計（上位5件のみ）
  Object.keys(statusSummary)
    .sort((a, b) => statusSummary[b].count - statusSummary[a].count)
    .slice(0, 5)
    .forEach(status => {
      const stat = statusSummary[status];
      html += `
      <div class="status-row">
        <span class="status-name">${status}</span>
        <span class="status-count">${stat.count}件（平均${stat.avgDaysStale}日）</span>
      </div>`;
    });
  
  if (Object.keys(statusSummary).length > 5) {
    html += `
      <div class="status-row" style="font-style: italic; color: #6c757d;">
        <span>他${Object.keys(statusSummary).length - 5}ステータス...</span>
      </div>`;
  }
  
  html += `
    </div>
    
    <a href="${reportUrl}" class="report-button">
      📈 詳細レポートを開く
    </a>
    
    <div class="action-box">
      <h4>対応方法</h4>
      <ol>
        <li>上記ボタンから詳細レポートを開く</li>
        <li>優先度の高い案件から確認</li>
        <li>担当者へ更新を依頼</li>
        <li>HubSpotで案件情報を更新</li>
      </ol>
    </div>
    
    <div class="footer">
      <p>対象金額合計: ¥${totalAmount.toLocaleString()}</p>
      <p>対象シート: ${CONFIG.SHEETS.DEALS}</p>
      <p>このメールは自動送信されています</p>
    </div>
  </div>
</body>
</html>`;
  
  return html;
}

/**
 * 簡略化されたプレーンテキスト形式のメール本文を作成
 */
function createSimplifiedPlainTextEmailBody(staleDeals, statusSummary, reportUrl) {
  const totalAmount = staleDeals.reduce((sum, d) => sum + (d.amount || 0), 0);
  
  let text = `HubSpot 案件更新アラート
━━━━━━━━━━━━━━━━━━

【サマリー】
未更新案件数: ${staleDeals.length}件
緊急対応(7日以上): ${staleDeals.filter(d => d.daysSinceUpdate >= 7).length}件
要注意(5-6日): ${staleDeals.filter(d => d.daysSinceUpdate >= 5 && d.daysSinceUpdate < 7).length}件
通常(3-4日): ${staleDeals.filter(d => d.daysSinceUpdate < 5).length}件

【ステータス別集計（上位5件）】`;

  // ステータス別集計を追加
  Object.keys(statusSummary)
    .sort((a, b) => statusSummary[b].count - statusSummary[a].count)
    .slice(0, 5)
    .forEach(status => {
      const stat = statusSummary[status];
      text += `\n・${status}: ${stat.count}件 (平均${stat.avgDaysStale}日停滞)`;
    });
  
  text += `

対象金額合計: ¥${totalAmount.toLocaleString()}

【詳細レポート】
${reportUrl}

【対応方法】
1. 上記URLから詳細レポートを開く
2. 優先度の高い案件から確認
3. 担当者へ更新を依頼
4. HubSpotで案件情報を更新

対象シート: ${CONFIG.SHEETS.DEALS}
━━━━━━━━━━━━━━━━━━
このメールは自動送信されています`;
  
  return text;
}

// ==================== デバッグ関数 ====================

/**
 * データ診断
 */
function debugDataCheck() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const dealsSheet = spreadsheet.getSheetByName(CONFIG.SHEETS.DEALS);
    const ui = SpreadsheetApp.getUi();
    
    if (!dealsSheet) {
      ui.alert('エラー', `シート「${CONFIG.SHEETS.DEALS}」が見つかりません`, ui.ButtonSet.OK);
      return;
    }
    
    const data = dealsSheet.getDataRange().getValues();
    let message = `【データ診断結果】\n\n`;
    
    message += `データ行数: ${data.length}行\n`;
    message += `データ列数: ${data[0] ? data[0].length : 0}列\n\n`;
    
    if (data.length > 0) {
      message += `【ヘッダー行】\n`;
      message += `C列: "${data[0][2]}"\n`;
      message += `G列: "${data[0][6]}"\n\n`;
    }
    
    const now = new Date();
    const thresholdDate = new Date(now.getTime() - (CONFIG.DAYS_THRESHOLD * 24 * 60 * 60 * 1000));
    
    message += `【データサンプル（2-4行目）】\n`;
    for (let i = 1; i < Math.min(4, data.length); i++) {
      message += `\n${i + 1}行目:\n`;
      message += `  ステータス: ${data[i][2]}\n`;
      message += `  最終更新: ${data[i][6]}\n`;
      
      if (data[i][6]) {
        const lastModified = new Date(data[i][6]);
        const daysSince = Math.floor((now - lastModified) / (24 * 60 * 60 * 1000));
        message += `  → ${daysSince}日前に更新\n`;
        message += `  → 3日以上前?: ${lastModified < thresholdDate ? 'はい' : 'いいえ'}\n`;
      }
    }
    
    ui.alert('データ診断', message, ui.ButtonSet.OK);
    
  } catch (error) {
    SpreadsheetApp.getUi().alert('エラー', `診断中にエラー: ${error}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * 成約前ステータスの確認
 */
function checkPreContractStatuses() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const statusMasterSheet = spreadsheet.getSheetByName(CONFIG.SHEETS.STATUS_MASTER);
    const ui = SpreadsheetApp.getUi();
    
    if (!statusMasterSheet) {
      ui.alert('エラー', `シート「${CONFIG.SHEETS.STATUS_MASTER}」が見つかりません`, ui.ButtonSet.OK);
      return;
    }
    
    const statusData = statusMasterSheet.getDataRange().getValues();
    let preContractList = [];
    let postContractList = [];
    let lostList = [];  // 失注リスト
    
    let message = `【ステータスマスター診断】\n\n`;
    message += `総行数: ${statusData.length}行\n\n`;
    
    // データ確認
    for (let i = 1; i < statusData.length; i++) {
      const category = statusData[i][0];
      const status = statusData[i][1];
      
      if (category === '成約前' && status) {
        preContractList.push(status);
      } else if (category === '成約後' && status) {
        postContractList.push(status);
      } else if (category === '失注' && status) {
        lostList.push(status);
      }
    }
    
    message += `成約前ステータス: ${preContractList.length}個\n`;
    message += `成約後ステータス: ${postContractList.length}個\n`;
    message += `失注ステータス: ${lostList.length}個\n\n`;
    
    if (preContractList.length > 0) {
      message += `【成約前ステータス一覧】\n`;
      preContractList.forEach((status, idx) => {
        if (idx < 5) {  // 最初の5個のみ表示
          message += `${idx + 1}. ${status}\n`;
        }
      });
      if (preContractList.length > 5) {
        message += `... 他${preContractList.length - 5}個\n`;
      }
    }
    
    if (lostList.length > 0) {
      message += `\n【失注ステータス一覧（通知対象外）】\n`;
      lostList.forEach((status, idx) => {
        if (idx < 5) {  // 最初の5個のみ表示
          message += `${idx + 1}. ${status}\n`;
        }
      });
      if (lostList.length > 5) {
        message += `... 他${lostList.length - 5}個\n`;
      }
    }
    
    if (preContractList.length === 0) {
      message += `⚠️ 成約前ステータスが見つかりません！\n`;
      message += `StatusMasterシートのA列に「成約前」と記載してください。`;
    }
    
    ui.alert('成約前ステータス確認', message, ui.ButtonSet.OK);
    
  } catch (error) {
    SpreadsheetApp.getUi().alert('エラー', `確認中にエラー: ${error}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * テストメール送信
 */
function testEmailNotification() {
  // テストデータを作成
  const testData = [
    {
      dealName: 'テスト案件1 - 緊急',
      amount: 1000000,
      dealStage: 'リード (souco-tenant)',
      createDate: '2024/01/01',
      closeDate: '2024/03/31',
      dealOwner: 'テスト太郎',
      lastModified: '2024/02/01',
      daysSinceUpdate: 10,
      rowNumber: 10
    },
    {
      dealName: 'テスト案件2 - 要注意',
      amount: 500000,
      dealStage: '商談中',
      createDate: '2024/01/15',
      closeDate: '2024/03/15',
      dealOwner: 'テスト花子',
      lastModified: '2024/02/05',
      daysSinceUpdate: 6,
      rowNumber: 15
    },
    {
      dealName: 'テスト案件3 - 通常',
      amount: 300000,
      dealStage: '提案済み',
      createDate: '2024/02/01',
      closeDate: '2024/04/01',
      dealOwner: 'テスト次郎',
      lastModified: '2024/02/07',
      daysSinceUpdate: 4,
      rowNumber: 20
    }
  ];
  
  // テスト用レポートスプレッドシートを作成
  const reportUrl = createStaleDealsReportSpreadsheet(testData);
  
  // テストメールを送信
  sendEmailNotification(testData, reportUrl);
  SpreadsheetApp.getUi().alert('テスト送信', `テストメールを送信しました\n\nレポート: ${reportUrl}`, SpreadsheetApp.getUi().ButtonSet.OK);
}

// ==================== ユーティリティ関数 ====================

/**
 * 日付をフォーマット
 */
function formatDate(date) {
  if (!date) return '-';
  
  const d = new Date(date);
  if (isNaN(d.getTime())) return '-';
  
  const year = d.getFullYear();
  const month = String(d.getMonth() + 1).padStart(2, '0');
  const day = String(d.getDate()).padStart(2, '0');
  
  return `${year}/${month}/${day}`;
}

/**
 * 実行ログを記録
 */
function logExecution(dealCount) {
  try {
    const props = PropertiesService.getScriptProperties();
    const logs = JSON.parse(props.getProperty('executionLogs') || '[]');
    
    // 新しいログエントリを追加
    logs.push({
      timestamp: new Date().toISOString(),
      dealCount: dealCount,
      status: dealCount > 0 ? 'notified' : 'no_deals'
    });
    
    // 最新30件のみ保持
    if (logs.length > 30) {
      logs.splice(0, logs.length - 30);
    }
    
    props.setProperty('executionLogs', JSON.stringify(logs));
  } catch (error) {
    console.error('ログ記録エラー:', error);
  }
}

/**
 * 実行ログを表示
 */
function showExecutionLog() {
  try {
    const ui = SpreadsheetApp.getUi();
    const props = PropertiesService.getScriptProperties();
    const logs = JSON.parse(props.getProperty('executionLogs') || '[]');
    
    if (logs.length === 0) {
      ui.alert('実行ログ', 'まだ実行ログがありません', ui.ButtonSet.OK);
      return;
    }
    
    let message = '【最近の実行ログ（最新10件）】\n\n';
    const recentLogs = logs.slice(-10).reverse();
    
    recentLogs.forEach((log, index) => {
      const date = new Date(log.timestamp);
      const formattedDate = `${date.getFullYear()}/${String(date.getMonth() + 1).padStart(2, '0')}/${String(date.getDate()).padStart(2, '0')} ${String(date.getHours()).padStart(2, '0')}:${String(date.getMinutes()).padStart(2, '0')}`;
      const status = log.status === 'notified' ? `${log.dealCount}件通知` : '対象なし';
      message += `${formattedDate} - ${status}\n`;
    });
    
    ui.alert('実行ログ', message, ui.ButtonSet.OK);
  } catch (error) {
    SpreadsheetApp.getUi().alert('エラー', `ログ表示エラー: ${error}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * エラー通知を送信
 */
function sendErrorNotification(error) {
  const subject = '[エラー] HubSpot案件通知システム';
  const body = `
案件通知システムでエラーが発生しました。

エラー内容:
${error.toString()}

スタックトレース:
${error.stack || 'なし'}

発生日時: ${formatDate(new Date())} ${new Date().toLocaleTimeString('ja-JP')}
`;
  
  try {
    MailApp.sendEmail({
      to: CONFIG.EMAIL_RECIPIENTS,
      subject: subject,
      body: body,
      noReply: true
    });
  } catch (mailError) {
    console.error('エラー通知の送信に失敗:', mailError);
  }
}