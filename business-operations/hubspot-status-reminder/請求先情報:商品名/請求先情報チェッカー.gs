/**
 * HubSpot Deal 請求先情報チェッカー
 * 成約後の案件で請求先企業名、商品名、souco案件ID、soucoIDが抜けているものを通知
 */

// ==================== 設定値 ====================
const CONFIG = {
  // 通知先メールアドレス（必須）
  EMAIL_RECIPIENTS: 'user1@example.com,user2@example.com,admin@example.com',
  
  // スプレッドシートのシート名
  SHEETS: {
    DEALS: 'HS/deals/09Sep',        // 案件データのシート名
    STATUS_MASTER: 'StatusMaster'   // ステータスマスターのシート名
  },
  
  // 列のインデックス（0から開始）
  COLUMNS: {
    DEAL_NAME: 0,           // A列: [Deals] Deal Name
    AMOUNT: 1,              // B列: [Deals] Amount
    DEAL_STAGE: 2,          // C列: [Deals] Deal Stage
    CREATE_DATE: 3,         // D列: [Deals] Create Date
    CLOSE_DATE: 4,          // E列: [Deals] Close Date
    DEAL_OWNER: 5,          // F列: [Deals] Deal owner
    BILLING_COMPANY: 6,     // G列: [Deals] 利用者側_請求先_企業名
    PRODUCT_NAME: 7,        // H列: [Deals] 契約書に記載する商品名
    SOUCO_CASE_ID: 8,       // I列: [Deals] souco案件ID
    WAREHOUSE_ID: 9         // J列: [Deals] 契約した倉庫ID
  }
};

// ==================== メイン関数 ====================

/**
 * メニュー作成（自動実行）
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('請求先情報チェック')
    .addItem('今すぐチェック', 'checkMissingBillingInfoManual')
    .addItem('成約後ステータスを確認', 'checkPostContractStatuses')
    .addItem('テストメール送信', 'testBillingInfoEmail')
    .addItem('データを診断', 'debugBillingDataCheck')
    .addToUi();
}

/**
 * メイン実行関数（自動実行対応）
 */
function checkMissingBillingInfo() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    
    // 成約後ステータスの取得
    const postContractStatuses = getPostContractStatuses(spreadsheet);
    console.log('成約後ステータス数:', postContractStatuses.length);
    
    // 請求先情報が不足している案件の取得
    const incompleteDeals = getIncompleteBillingDeals(spreadsheet, postContractStatuses);
    console.log('情報不足案件数:', incompleteDeals.length);
    
    // 案件が存在する場合、レポートスプレッドシート作成とメール通知を送信
    if (incompleteDeals.length > 0) {
      const reportUrl = createReportSpreadsheet(incompleteDeals);
      sendBillingInfoEmail(incompleteDeals, reportUrl);
      console.log(`通知完了: ${incompleteDeals.length}件の情報不足案件を通知しました`);
    } else {
      console.log('確認完了: 情報不足の案件はありません');
    }
    
  } catch (error) {
    console.error('エラーが発生しました:', error);
    sendErrorNotification(error);
  }
}

/**
 * 手動実行用関数（UIアラート付き）
 */
function checkMissingBillingInfoManual() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const ui = SpreadsheetApp.getUi();
    
    // 成約後ステータスの取得
    const postContractStatuses = getPostContractStatuses(spreadsheet);
    console.log('成約後ステータス数:', postContractStatuses.length);
    
    // 請求先情報が不足している案件の取得
    const incompleteDeals = getIncompleteBillingDeals(spreadsheet, postContractStatuses);
    console.log('情報不足案件数:', incompleteDeals.length);
    
    // 案件が存在する場合、レポートスプレッドシート作成とメール通知を送信
    if (incompleteDeals.length > 0) {
      const reportUrl = createReportSpreadsheet(incompleteDeals);
      sendBillingInfoEmail(incompleteDeals, reportUrl);
      ui.alert('通知完了', `${incompleteDeals.length}件の情報不足案件を通知しました\n\nレポート: ${reportUrl}`, ui.ButtonSet.OK);
    } else {
      ui.alert('確認完了', '情報不足の案件はありません', ui.ButtonSet.OK);
    }
    
  } catch (error) {
    console.error('エラーが発生しました:', error);
    SpreadsheetApp.getUi().alert('エラー', `エラーが発生しました: ${error}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * 成約後のステータス一覧を取得
 */
function getPostContractStatuses(spreadsheet) {
  const statusMasterSheet = spreadsheet.getSheetByName(CONFIG.SHEETS.STATUS_MASTER);
  
  if (!statusMasterSheet) {
    throw new Error(`StatusMasterシート「${CONFIG.SHEETS.STATUS_MASTER}」が見つかりません`);
  }
  
  const statusData = statusMasterSheet.getDataRange().getValues();
  const postContractStatuses = [];
  
  // ヘッダー行をスキップして処理
  for (let i = 1; i < statusData.length; i++) {
    const category = statusData[i][0]; // A列: 成約前/成約後/失注
    const status = statusData[i][1];   // B列: ステータス名
    
    if (category === '成約後' && status) {
      postContractStatuses.push(status);
      console.log('成約後ステータス:', status);
    }
  }
  
  return postContractStatuses;
}

/**
 * 請求先情報が不足している成約後案件を取得
 */
function getIncompleteBillingDeals(spreadsheet, postContractStatuses) {
  const dealsSheet = spreadsheet.getSheetByName(CONFIG.SHEETS.DEALS);
  
  if (!dealsSheet) {
    const sheets = spreadsheet.getSheets();
    const sheetNames = sheets.map(s => s.getName()).join(', ');
    throw new Error(`シート「${CONFIG.SHEETS.DEALS}」が見つかりません。存在するシート: ${sheetNames}`);
  }
  
  const dealsData = dealsSheet.getDataRange().getValues();
  const incompleteDeals = [];
  
  // ヘッダー行をスキップして処理
  for (let i = 1; i < dealsData.length; i++) {
    const dealStage = dealsData[i][CONFIG.COLUMNS.DEAL_STAGE];
    
    // 成約後ステータスの案件のみチェック
    if (postContractStatuses.includes(dealStage)) {
      const billingCompany = dealsData[i][CONFIG.COLUMNS.BILLING_COMPANY];
      const productName = dealsData[i][CONFIG.COLUMNS.PRODUCT_NAME];
      const soucoCaseId = dealsData[i][CONFIG.COLUMNS.SOUCO_CASE_ID];
      const warehouseId = dealsData[i][CONFIG.COLUMNS.WAREHOUSE_ID];
      
      // いずれかの情報が欠けている場合
      const missingFields = [];
      if (!billingCompany || billingCompany === '') missingFields.push('利用者側_請求先_企業名');
      if (!productName || productName === '') missingFields.push('契約書に記載する商品名');
      if (!soucoCaseId || soucoCaseId === '') missingFields.push('souco案件ID');
      if (!warehouseId || warehouseId === '') missingFields.push('契約した倉庫ID');
      
      if (missingFields.length > 0) {
        incompleteDeals.push({
          dealName: dealsData[i][CONFIG.COLUMNS.DEAL_NAME] || '(名称未設定)',
          amount: dealsData[i][CONFIG.COLUMNS.AMOUNT],
          dealStage: dealStage,
          dealOwner: dealsData[i][CONFIG.COLUMNS.DEAL_OWNER] || '未割当',
          closeDate: formatDate(dealsData[i][CONFIG.COLUMNS.CLOSE_DATE]),
          billingCompany: billingCompany || '-',
          productName: productName || '-',
          soucoCaseId: soucoCaseId || '-',
          warehouseId: warehouseId || '-',
          missingFields: missingFields,
          rowNumber: i + 1  // スプレッドシート上の行番号
        });
      }
    }
  }
  
  // 不足フィールド数でソート（多い順）
  incompleteDeals.sort((a, b) => b.missingFields.length - a.missingFields.length);
  
  return incompleteDeals;
}

// ==================== レポート作成関数 ====================

/**
 * レポート用スプレッドシートを新規作成
 */
function createReportSpreadsheet(incompleteDeals) {
  // 現在の日時を取得
  const now = new Date();
  const dateStr = Utilities.formatDate(now, 'JST', 'yyyy-MM-dd HH:mm');
  const fileName = `請求先情報不足レポート_${Utilities.formatDate(now, 'JST', 'yyyyMMdd_HHmmss')}`;
  
  // 新規スプレッドシートを作成
  const newSpreadsheet = SpreadsheetApp.create(fileName);
  
  // スプレッドシートのアクセス権限を設定（URLを知っている全員が編集可能）
  const file = DriveApp.getFileById(newSpreadsheet.getId());
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.EDIT);
  
  const sheet = newSpreadsheet.getActiveSheet();
  sheet.setName('情報不足案件一覧');
  
  // ヘッダー行を設定
  const headers = [
    '行番号',
    '案件名',
    'ステータス',
    '担当者',
    '成約日',
    '金額',
    '利用者側_請求先_企業名',
    '契約書に記載する商品名',
    'souco案件ID',
    '契約した倉庫ID',
    '不足項目',
    '不足項目数'
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // ヘッダー行のスタイル設定
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#4a86e8');
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');
  headerRange.setHorizontalAlignment('center');
  
  // データを追加
  const dataRows = incompleteDeals.map(deal => [
    deal.rowNumber,
    deal.dealName,
    deal.dealStage,
    deal.dealOwner,
    deal.closeDate,
    deal.amount || 0,
    deal.billingCompany === '-' ? '' : deal.billingCompany,
    deal.productName === '-' ? '' : deal.productName,
    deal.soucoCaseId === '-' ? '' : deal.soucoCaseId,
    deal.warehouseId === '-' ? '' : deal.warehouseId,
    deal.missingFields.join(', '),
    deal.missingFields.length
  ]);
  
  if (dataRows.length > 0) {
    sheet.getRange(2, 1, dataRows.length, headers.length).setValues(dataRows);
    
    // 不足項目数によって行の背景色を設定
    for (let i = 0; i < dataRows.length; i++) {
      const row = i + 2;
      const missingCount = incompleteDeals[i].missingFields.length;
      let bgColor = '#ffffff';
      
      if (missingCount >= 3) {
        bgColor = '#ffebee'; // 赤系（重要）
      } else if (missingCount === 2) {
        bgColor = '#fff3cd'; // 黄系（要対応）
      } else if (missingCount === 1) {
        bgColor = '#f1f8e9'; // 緑系（軽微）
      }
      
      sheet.getRange(row, 1, 1, headers.length).setBackground(bgColor);
      
      // 不足している項目のセルを赤く塗る
      if (incompleteDeals[i].billingCompany === '-') {
        sheet.getRange(row, 7).setBackground('#ffcdd2');
      }
      if (incompleteDeals[i].productName === '-') {
        sheet.getRange(row, 8).setBackground('#ffcdd2');
      }
      if (incompleteDeals[i].soucoCaseId === '-') {
        sheet.getRange(row, 9).setBackground('#ffcdd2');
      }
      if (incompleteDeals[i].warehouseId === '-') {
        sheet.getRange(row, 10).setBackground('#ffcdd2');
      }
    }
  }
  
  // 列幅を自動調整
  sheet.autoResizeColumns(1, headers.length);
  
  // サマリーシートを追加
  const summarySheet = newSpreadsheet.insertSheet('サマリー');
  
  // サマリー情報を作成
  const missingSummary = createMissingSummary(incompleteDeals);
  const summaryData = [
    ['レポート作成日時', dateStr],
    ['', ''],  // 空行
    ['【概要】', ''],
    ['情報不足案件総数', incompleteDeals.length],
    ['対象金額合計', missingSummary.totalAmount],
    ['', ''],  // 空行
    ['【重要度別内訳】', ''],
    ['重要（3項目以上不足）', incompleteDeals.filter(d => d.missingFields.length >= 3).length],
    ['要対応（2項目不足）', incompleteDeals.filter(d => d.missingFields.length === 2).length],
    ['軽微（1項目不足）', incompleteDeals.filter(d => d.missingFields.length === 1).length],
    ['', ''],  // 空行
    ['【不足項目別集計】', ''],
    ['利用者側_請求先_企業名', missingSummary['利用者側_請求先_企業名']],
    ['契約書に記載する商品名', missingSummary['契約書に記載する商品名']],
    ['souco案件ID', missingSummary['souco案件ID']],
    ['契約した倉庫ID', missingSummary['契約した倉庫ID']]
  ];
  
  summarySheet.getRange(1, 1, summaryData.length, 2).setValues(summaryData);
  
  // サマリーシートのスタイル設定
  summarySheet.getRange('A3').setFontWeight('bold');
  summarySheet.getRange('A7').setFontWeight('bold');
  summarySheet.getRange('A12').setFontWeight('bold');
  summarySheet.getRange('A1:B1').setBackground('#e8f0fe');
  summarySheet.autoResizeColumns(1, 2);
  
  // スプレッドシートのURLを返す
  return newSpreadsheet.getUrl();
}

// ==================== メール送信関数 ====================

/**
 * HTMLメール通知を送信
 */
function sendBillingInfoEmail(incompleteDeals, reportUrl) {
  const subject = `[HubSpot] 請求先情報が不足している成約後案件: ${incompleteDeals.length}件`;
  
  // 不足項目別の集計を作成
  const missingSummary = createMissingSummary(incompleteDeals);
  
  const htmlBody = createBillingHtmlEmailBody(incompleteDeals, missingSummary, reportUrl);
  const plainTextBody = createBillingPlainTextEmailBody(incompleteDeals, missingSummary, reportUrl);
  
  MailApp.sendEmail({
    to: CONFIG.EMAIL_RECIPIENTS,
    subject: subject,
    body: plainTextBody,
    htmlBody: htmlBody,
    noReply: true
  });
}

/**
 * 不足項目別の集計を作成
 */
function createMissingSummary(incompleteDeals) {
  const summary = {
    '利用者側_請求先_企業名': 0,
    '契約書に記載する商品名': 0,
    'souco案件ID': 0,
    '契約した倉庫ID': 0,
    totalAmount: 0
  };
  
  incompleteDeals.forEach(deal => {
    deal.missingFields.forEach(field => {
      summary[field]++;
    });
    summary.totalAmount += (deal.amount || 0);
  });
  
  return summary;
}

/**
 * HTML形式のメール本文を作成
 */
function createBillingHtmlEmailBody(incompleteDeals, missingSummary, reportUrl) {
  const criticalDeals = incompleteDeals.filter(d => d.missingFields.length >= 3);
  const warningDeals = incompleteDeals.filter(d => d.missingFields.length === 2);
  const minorDeals = incompleteDeals.filter(d => d.missingFields.length === 1);
  
  let html = `
<!DOCTYPE html>
<html>
<head>
  <style>
    body { font-family: Arial, sans-serif; line-height: 1.6; color: #333; }
    .container { max-width: 600px; margin: 0 auto; background: #fff; padding: 20px; }
    .header { background: linear-gradient(135deg, #ff6b6b 0%, #ff8e53 100%); color: white; padding: 25px; text-align: center; border-radius: 5px; margin-bottom: 20px; }
    .header h1 { margin: 0; font-size: 24px; }
    .header p { margin: 10px 0 0; opacity: 0.9; font-size: 14px; }
    .summary { display: flex; justify-content: space-around; margin: 20px 0; padding: 20px; background: #f8f9fa; border-radius: 5px; }
    .summary-item { text-align: center; }
    .summary-label { font-size: 12px; color: #6c757d; margin-top: 5px; }
    .summary-number { font-size: 28px; font-weight: bold; }
    .critical { color: #dc3545; }
    .warning { color: #ffc107; }
    .minor { color: #28a745; }
    .report-button { display: block; background: #007bff; color: white; text-decoration: none; padding: 15px 30px; border-radius: 5px; text-align: center; font-weight: bold; margin: 25px 0; }
    .report-button:hover { background: #0056b3; }
    .field-summary { margin: 20px 0; padding: 15px; background: #fff5f5; border-left: 4px solid #dc3545; }
    .field-summary h4 { margin: 0 0 10px 0; color: #dc3545; font-size: 16px; }
    .field-row { display: flex; justify-content: space-between; padding: 5px 0; border-bottom: 1px solid #ffe0e0; }
    .field-name { color: #495057; font-size: 14px; }
    .field-count { color: #dc3545; font-weight: bold; font-size: 14px; }
    .action-box { margin: 20px 0; padding: 15px; background: #e3f2fd; border-left: 4px solid #2196f3; }
    .action-box h4 { margin: 0 0 10px 0; color: #1976d2; font-size: 16px; }
    .action-box ol { margin: 10px 0; padding-left: 20px; font-size: 14px; }
    .footer { margin-top: 30px; padding-top: 20px; border-top: 1px solid #dee2e6; text-align: center; color: #6c757d; font-size: 12px; }
  </style>
</head>
<body>
  <div class="container">
    <div class="header">
      <h1>📋 請求先情報不足アラート</h1>
      <p>成約後案件の必須情報チェック</p>
    </div>
    
    <div class="summary">
      <div class="summary-item">
        <div class="summary-number critical">${incompleteDeals.length}</div>
        <div class="summary-label">情報不足案件</div>
      </div>
      <div class="summary-item">
        <div class="summary-number critical">${criticalDeals.length}</div>
        <div class="summary-label">重要（3項目以上）</div>
      </div>
      <div class="summary-item">
        <div class="summary-number warning">${warningDeals.length}</div>
        <div class="summary-label">要対応（2項目）</div>
      </div>
    </div>
    
    <div class="field-summary">
      <h4>不足項目別集計</h4>
      <div class="field-row">
        <span class="field-name">利用者側_請求先_企業名</span>
        <span class="field-count">${missingSummary['利用者側_請求先_企業名']}件</span>
      </div>
      <div class="field-row">
        <span class="field-name">契約書に記載する商品名</span>
        <span class="field-count">${missingSummary['契約書に記載する商品名']}件</span>
      </div>
      <div class="field-row">
        <span class="field-name">souco案件ID</span>
        <span class="field-count">${missingSummary['souco案件ID']}件</span>
      </div>
      <div class="field-row">
        <span class="field-name">契約した倉庫ID</span>
        <span class="field-count">${missingSummary['契約した倉庫ID']}件</span>
      </div>
    </div>`;
  
  
  html += `
    <a href="${reportUrl}" class="report-button">
      📊 詳細レポートを開く
    </a>
    
    <div class="action-box">
      <h4>対応方法</h4>
      <ol>
        <li>上記ボタンから詳細レポートを開く</li>
        <li>不足情報を確認して入力する</li>
        <li>全項目の入力完了を確認</li>
      </ol>
    </div>
    
    <div class="footer">
      <p>このメールは自動送信されています</p>
      <p>対象シート: ${CONFIG.SHEETS.DEALS}</p>
    </div>
  </div>
</body>
</html>`;
  
  return html;
}

/**
 * プレーンテキスト形式のメール本文を作成
 */
function createBillingPlainTextEmailBody(incompleteDeals, missingSummary, reportUrl) {
  let text = `請求先情報不足アラート
━━━━━━━━━━━━━━━━━━

【サマリー】
情報不足案件数: ${incompleteDeals.length}件
重要(3項目以上): ${incompleteDeals.filter(d => d.missingFields.length >= 3).length}件
要対応(2項目): ${incompleteDeals.filter(d => d.missingFields.length === 2).length}件
軽微(1項目): ${incompleteDeals.filter(d => d.missingFields.length === 1).length}件

【不足項目別集計】
・利用者側_請求先_企業名: ${missingSummary['利用者側_請求先_企業名']}件
・契約書に記載する商品名: ${missingSummary['契約書に記載する商品名']}件
・souco案件ID: ${missingSummary['souco案件ID']}件
・契約した倉庫ID: ${missingSummary['契約した倉庫ID']}件

対象金額合計: ¥${missingSummary.totalAmount.toLocaleString()}

【詳細レポート】
${reportUrl}

【対応方法】
1. 上記URLから詳細レポートを開く
2. 不足情報を確認して入力する
3. 全項目の入力完了を確認

対象シート: ${CONFIG.SHEETS.DEALS}
━━━━━━━━━━━━━━━━━━
このメールは自動送信されています`;
  
  return text;
}

// ==================== デバッグ関数 ====================

/**
 * データ診断
 */
function debugBillingDataCheck() {
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
      message += `G列: "${data[0][6]}"\n`;
      message += `H列: "${data[0][7]}"\n`;
      message += `I列: "${data[0][8]}"\n`;
      message += `J列: "${data[0][9]}"\n\n`;
    }
    
    // 成約後ステータスを取得
    const postContractStatuses = getPostContractStatuses(spreadsheet);
    let postContractCount = 0;
    let incompleteCount = 0;
    
    message += `【データサンプル（最初の成約後案件3件）】\n`;
    let sampleCount = 0;
    for (let i = 1; i < data.length && sampleCount < 3; i++) {
      if (postContractStatuses.includes(data[i][2])) {
        postContractCount++;
        const missingFields = [];
        if (!data[i][6]) missingFields.push('利用者側_請求先_企業名');
        if (!data[i][7]) missingFields.push('契約書に記載する商品名');
        if (!data[i][8]) missingFields.push('souco案件ID');
        if (!data[i][9]) missingFields.push('契約した倉庫ID');
        
        if (missingFields.length > 0) {
          incompleteCount++;
          message += `\n行${i + 1}: ${data[i][0]}\n`;
          message += `  ステータス: ${data[i][2]}\n`;
          message += `  不足項目: ${missingFields.join(', ')}\n`;
          sampleCount++;
        }
      }
    }
    
    message += `\n【集計】\n`;
    message += `成約後案件総数: ${postContractCount}件\n`;
    message += `情報不足案件数: ${incompleteCount}件\n`;
    
    ui.alert('データ診断', message, ui.ButtonSet.OK);
    
  } catch (error) {
    SpreadsheetApp.getUi().alert('エラー', `診断中にエラー: ${error}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * 成約後ステータスの確認
 */
function checkPostContractStatuses() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const statusMasterSheet = spreadsheet.getSheetByName(CONFIG.SHEETS.STATUS_MASTER);
    const ui = SpreadsheetApp.getUi();
    
    if (!statusMasterSheet) {
      ui.alert('エラー', `シート「${CONFIG.SHEETS.STATUS_MASTER}」が見つかりません`, ui.ButtonSet.OK);
      return;
    }
    
    const statusData = statusMasterSheet.getDataRange().getValues();
    let postContractList = [];
    
    let message = `【ステータスマスター診断】\n\n`;
    message += `総行数: ${statusData.length}行\n\n`;
    
    // データ確認
    for (let i = 1; i < statusData.length; i++) {
      const category = statusData[i][0];
      const status = statusData[i][1];
      
      if (category === '成約後' && status) {
        postContractList.push(status);
      }
    }
    
    message += `成約後ステータス: ${postContractList.length}個\n\n`;
    
    if (postContractList.length > 0) {
      message += `【成約後ステータス一覧】\n`;
      postContractList.forEach((status, idx) => {
        message += `${idx + 1}. ${status}\n`;
      });
    } else {
      message += `⚠️ 成約後ステータスが見つかりません！\n`;
      message += `StatusMasterシートのA列に「成約後」と記載してください。`;
    }
    
    ui.alert('成約後ステータス確認', message, ui.ButtonSet.OK);
    
  } catch (error) {
    SpreadsheetApp.getUi().alert('エラー', `確認中にエラー: ${error}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * テストメール送信
 */
function testBillingInfoEmail() {
  // テストデータを作成
  const testData = [
    {
      dealName: 'テスト案件1',
      amount: 1000000,
      dealStage: '成約',
      dealOwner: 'テスト太郎',
      closeDate: '2024/03/31',
      billingCompany: 'テスト株式会社',
      productName: '-',
      soucoCaseId: '-',
      warehouseId: 'WH12345',
      missingFields: ['契約書に記載する商品名', 'souco案件ID'],
      rowNumber: 10
    },
    {
      dealName: 'テスト案件2',
      amount: 500000,
      dealStage: '契約締結済み',
      dealOwner: 'テスト花子',
      closeDate: '2024/03/15',
      billingCompany: '-',
      productName: '-',
      soucoCaseId: '-',
      warehouseId: '-',
      missingFields: ['利用者側_請求先_企業名', '契約書に記載する商品名', 'souco案件ID', '契約した倉庫ID'],
      rowNumber: 15
    }
  ];
  
  // テスト用レポートスプレッドシートを作成
  const reportUrl = createReportSpreadsheet(testData);
  
  // テストメールを送信
  sendBillingInfoEmail(testData, reportUrl);
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
 * エラー通知を送信
 */
function sendErrorNotification(error) {
  const subject = '[エラー] HubSpot請求先情報チェックシステム';
  const body = `
請求先情報チェックシステムでエラーが発生しました。

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