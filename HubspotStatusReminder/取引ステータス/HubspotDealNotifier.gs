/**
 * HubSpot Deal Status Reminder
 * 成約前の案件で3日間更新されていないものを通知するGoogle Apps Script
 */

// 設定値（必要に応じて変更してください）
const CONFIG = {
  // 通知先メールアドレス（カンマ区切りで複数指定可能）
  EMAIL_RECIPIENTS: 'your-email@example.com',
  
  // スプレッドシートのURL（使用する場合はここに記入）
  SPREADSHEET_URL: '',  // 例: 'https://docs.google.com/spreadsheets/d/xxxxx/edit'
  
  // スプレッドシートのシート名
  SHEETS: {
    DEALS: 'HS/deals/05Sep',  // 実際のシート名
    STATUS_MASTER: 'StatusMaster'
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

/**
 * メイン実行関数
 * トリガーから実行される
 */
function checkAndNotifyStaleDeals() {
  try {
    // スプレッドシートの取得（URLが設定されている場合はそれを使用）
    const spreadsheet = CONFIG.SPREADSHEET_URL 
      ? SpreadsheetApp.openByUrl(CONFIG.SPREADSHEET_URL)
      : SpreadsheetApp.getActiveSpreadsheet();
    
    // 成約前ステータスの取得
    const preContractStatuses = getPreContractStatuses(spreadsheet);
    
    // 3日間更新されていない成約前案件の取得
    const staleDeals = getStaleDeals(spreadsheet, preContractStatuses);
    
    // 案件が存在する場合、メール通知を送信
    if (staleDeals.length > 0) {
      sendEmailNotification(staleDeals);
      Logger.log(`${staleDeals.length}件の未更新案件を通知しました`);
    } else {
      Logger.log('未更新案件はありません');
    }
    
  } catch (error) {
    console.error('エラーが発生しました:', error);
    sendErrorNotification(error);
  }
}

/**
 * 成約前のステータス一覧を取得
 */
function getPreContractStatuses(spreadsheet) {
  const statusMasterSheet = spreadsheet.getSheetByName(CONFIG.SHEETS.STATUS_MASTER);
  
  if (!statusMasterSheet) {
    throw new Error('StatusMasterシートが見つかりません');
  }
  
  const statusData = statusMasterSheet.getDataRange().getValues();
  const preContractStatuses = [];
  
  // ヘッダー行をスキップして処理
  for (let i = 1; i < statusData.length; i++) {
    const category = statusData[i][0]; // A列: 成約前/成約後
    const status = statusData[i][1];   // B列: ステータス名
    
    if (category === '成約前' && status) {
      preContractStatuses.push(status);
    }
  }
  
  return preContractStatuses;
}

/**
 * 3日間更新されていない成約前案件を取得
 */
function getStaleDeals(spreadsheet, preContractStatuses) {
  const dealsSheet = spreadsheet.getSheetByName(CONFIG.SHEETS.DEALS);
  
  if (!dealsSheet) {
    // より詳細なエラーメッセージ
    const sheets = spreadsheet.getSheets();
    const sheetNames = sheets.map(s => s.getName()).join(', ');
    throw new Error(`シート「${CONFIG.SHEETS.DEALS}」が見つかりません。存在するシート: ${sheetNames}`);
  }
  
  const dealsData = dealsSheet.getDataRange().getValues();
  const staleDeals = [];
  const now = new Date();
  const thresholdDate = new Date(now.getTime() - (CONFIG.DAYS_THRESHOLD * 24 * 60 * 60 * 1000));
  
  // ヘッダー行をスキップして処理
  for (let i = 1; i < dealsData.length; i++) {
    const dealStage = dealsData[i][CONFIG.COLUMNS.DEAL_STAGE];
    const lastModified = dealsData[i][CONFIG.COLUMNS.LAST_MODIFIED];
    
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
          daysSinceUpdate: daysSinceUpdate
        });
      }
    }
  }
  
  // 更新日数の多い順にソート
  staleDeals.sort((a, b) => b.daysSinceUpdate - a.daysSinceUpdate);
  
  return staleDeals;
}

/**
 * HTMLメール通知を送信
 */
function sendEmailNotification(staleDeals) {
  const subject = `[HubSpot] ${CONFIG.DAYS_THRESHOLD}日以上更新されていない成約前案件: ${staleDeals.length}件`;
  
  // HTML形式のメール本文を作成
  const htmlBody = createHtmlEmailBody(staleDeals);
  
  // プレーンテキスト版も作成（HTMLが表示できない環境用）
  const plainTextBody = createPlainTextEmailBody(staleDeals);
  
  // メール送信
  MailApp.sendEmail({
    to: CONFIG.EMAIL_RECIPIENTS,
    subject: subject,
    body: plainTextBody,
    htmlBody: htmlBody,
    noReply: true
  });
}

/**
 * HTML形式のメール本文を作成
 */
function createHtmlEmailBody(staleDeals) {
  const totalDeals = staleDeals.length;
  const urgentDeals = staleDeals.filter(d => d.daysSinceUpdate >= 7).length;
  const warningDeals = staleDeals.filter(d => d.daysSinceUpdate >= 5 && d.daysSinceUpdate < 7).length;
  
  let html = `
<!DOCTYPE html>
<html>
<head>
  <style>
    body {
      font-family: 'Helvetica Neue', Arial, sans-serif;
      line-height: 1.6;
      color: #333;
      background-color: #f5f5f5;
    }
    .container {
      max-width: 900px;
      margin: 0 auto;
      background-color: #ffffff;
      padding: 0;
      box-shadow: 0 2px 5px rgba(0,0,0,0.1);
    }
    .header {
      background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
      color: white;
      padding: 30px;
      text-align: center;
      border-radius: 5px 5px 0 0;
    }
    .header h1 {
      margin: 0;
      font-size: 24px;
      font-weight: 600;
    }
    .summary {
      padding: 25px 30px;
      background-color: #f8f9fa;
      border-bottom: 1px solid #e9ecef;
    }
    .summary-grid {
      display: flex;
      justify-content: space-around;
      margin-top: 15px;
    }
    .summary-item {
      text-align: center;
      padding: 10px;
    }
    .summary-number {
      font-size: 32px;
      font-weight: bold;
      margin-bottom: 5px;
    }
    .summary-label {
      font-size: 14px;
      color: #6c757d;
      text-transform: uppercase;
    }
    .urgent { color: #dc3545; }
    .warning { color: #ffc107; }
    .normal { color: #28a745; }
    .content {
      padding: 30px;
    }
    table {
      width: 100%;
      border-collapse: collapse;
      margin-top: 20px;
    }
    th {
      background-color: #6c757d;
      color: white;
      padding: 12px;
      text-align: left;
      font-weight: 600;
      font-size: 14px;
      border: 1px solid #6c757d;
    }
    td {
      padding: 10px 12px;
      border: 1px solid #dee2e6;
      font-size: 14px;
    }
    tr:nth-child(even) {
      background-color: #f8f9fa;
    }
    tr:hover {
      background-color: #e9ecef;
    }
    .priority-urgent {
      background-color: #ffebee !important;
    }
    .priority-warning {
      background-color: #fff3cd !important;
    }
    .days-badge {
      display: inline-block;
      padding: 3px 8px;
      border-radius: 12px;
      font-size: 12px;
      font-weight: bold;
      text-align: center;
    }
    .days-urgent {
      background-color: #dc3545;
      color: white;
    }
    .days-warning {
      background-color: #ffc107;
      color: #333;
    }
    .days-normal {
      background-color: #28a745;
      color: white;
    }
    .footer {
      padding: 20px 30px;
      background-color: #f8f9fa;
      border-top: 1px solid #dee2e6;
      text-align: center;
      color: #6c757d;
      font-size: 12px;
    }
    .action-required {
      background-color: #fff3cd;
      border-left: 4px solid #ffc107;
      padding: 15px;
      margin: 20px 0;
    }
    .action-required h3 {
      margin-top: 0;
      color: #856404;
    }
    @media only screen and (max-width: 600px) {
      .container {
        width: 100%;
      }
      .summary-grid {
        flex-direction: column;
      }
      table {
        font-size: 12px;
      }
      th, td {
        padding: 8px;
      }
    }
  </style>
</head>
<body>
  <div class="container">
    <div class="header">
      <h1>🔔 HubSpot 案件更新アラート</h1>
      <p style="margin: 10px 0 0; opacity: 0.9;">成約前案件の更新状況レポート</p>
    </div>
    
    <div class="summary">
      <h2 style="margin: 0 0 10px; color: #495057;">サマリー</h2>
      <div class="summary-grid">
        <div class="summary-item">
          <div class="summary-number urgent">${totalDeals}</div>
          <div class="summary-label">未更新案件</div>
        </div>
        <div class="summary-item">
          <div class="summary-number urgent">${urgentDeals}</div>
          <div class="summary-label">緊急対応<br>(7日以上)</div>
        </div>
        <div class="summary-item">
          <div class="summary-number warning">${warningDeals}</div>
          <div class="summary-label">要注意<br>(5-6日)</div>
        </div>
      </div>
    </div>
    
    <div class="content">
      <div class="action-required">
        <h3>⚠️ 対応が必要です</h3>
        <p style="margin: 0;">以下の成約前案件が${CONFIG.DAYS_THRESHOLD}日以上更新されていません。<br>
        早急に状況確認と更新をお願いします。</p>
      </div>
      
      <h3 style="color: #495057; margin-top: 30px;">📋 未更新案件リスト</h3>
      
      <table>
        <thead>
          <tr>
            <th style="width: 25%;">案件名</th>
            <th style="width: 8%;">未更新</th>
            <th style="width: 15%;">ステータス</th>
            <th style="width: 12%;">担当者</th>
            <th style="width: 12%;">金額</th>
            <th style="width: 10%;">作成日</th>
            <th style="width: 10%;">最終更新</th>
            <th style="width: 10%;">クローズ予定</th>
          </tr>
        </thead>
        <tbody>`;
  
  // 各案件の情報を追加
  staleDeals.forEach((deal, index) => {
    let rowClass = '';
    let daysClass = 'days-normal';
    let daysBadgeClass = 'days-badge days-normal';
    
    if (deal.daysSinceUpdate >= 7) {
      rowClass = 'priority-urgent';
      daysBadgeClass = 'days-badge days-urgent';
    } else if (deal.daysSinceUpdate >= 5) {
      rowClass = 'priority-warning';
      daysBadgeClass = 'days-badge days-warning';
    }
    
    const amount = deal.amount ? `¥${Number(deal.amount).toLocaleString()}` : '-';
    
    html += `
          <tr class="${rowClass}">
            <td style="font-weight: 600;">${deal.dealName || '(名称未設定)'}</td>
            <td style="text-align: center;">
              <span class="${daysBadgeClass}">${deal.daysSinceUpdate}日</span>
            </td>
            <td>${deal.dealStage}</td>
            <td>${deal.dealOwner}</td>
            <td style="text-align: right;">${amount}</td>
            <td>${deal.createDate}</td>
            <td>${deal.lastModified}</td>
            <td>${deal.closeDate}</td>
          </tr>`;
  });
  
  html += `
        </tbody>
      </table>
      
      <div style="margin-top: 30px; padding: 15px; background-color: #e7f3ff; border-radius: 5px;">
        <h4 style="margin: 0 0 10px; color: #004085;">💡 推奨アクション</h4>
        <ul style="margin: 5px 0; padding-left: 20px;">
          <li>各案件の現在の状況を確認し、HubSpotで最新情報に更新</li>
          <li>停滞している案件は、次のアクションを決定して実行</li>
          <li>不要な案件はステータスを「失注」または「削除」に変更</li>
        </ul>
      </div>
    </div>
    
    <div class="footer">
      <p>このメールは自動送信されています。<br>
      設定変更が必要な場合は、Google Apps Scriptの設定をご確認ください。</p>
      <p style="margin-top: 10px;">
        生成日時: ${formatDate(new Date())} | 
        閾値設定: ${CONFIG.DAYS_THRESHOLD}日
      </p>
    </div>
  </div>
</body>
</html>`;
  
  return html;
}

/**
 * プレーンテキスト形式のメール本文を作成
 */
function createPlainTextEmailBody(staleDeals) {
  let text = `HubSpot 案件更新アラート
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

【サマリー】
未更新案件数: ${staleDeals.length}件
緊急対応必要(7日以上): ${staleDeals.filter(d => d.daysSinceUpdate >= 7).length}件
要注意(5-6日): ${staleDeals.filter(d => d.daysSinceUpdate >= 5 && d.daysSinceUpdate < 7).length}件

【未更新案件リスト】
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n`;
  
  staleDeals.forEach((deal, index) => {
    const priority = deal.daysSinceUpdate >= 7 ? '[緊急]' : 
                    deal.daysSinceUpdate >= 5 ? '[要注意]' : '';
    
    text += `${index + 1}. ${priority} ${deal.dealName}
   未更新日数: ${deal.daysSinceUpdate}日
   ステータス: ${deal.dealStage}
   担当者: ${deal.dealOwner}
   金額: ${deal.amount ? `¥${Number(deal.amount).toLocaleString()}` : '-'}
   最終更新: ${deal.lastModified}
   クローズ予定: ${deal.closeDate}
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n`;
  });
  
  text += `
【推奨アクション】
1. 各案件の現在の状況を確認し、HubSpotで最新情報に更新
2. 停滞している案件は、次のアクションを決定して実行
3. 不要な案件はステータスを「失注」または「削除」に変更

このメールは自動送信されています。
生成日時: ${formatDate(new Date())}`;
  
  return text;
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

発生日時: ${formatDate(new Date())}
`;
  
  MailApp.sendEmail({
    to: CONFIG.EMAIL_RECIPIENTS,
    subject: subject,
    body: body,
    noReply: true
  });
}

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
 * 手動テスト用関数
 */
function testEmailNotification() {
  // テスト用のダミーデータ
  const testData = [
    {
      dealName: 'テスト案件1',
      amount: 1000000,
      dealStage: 'リード (souco-tenant)',
      createDate: '2024/01/01',
      closeDate: '2024/03/31',
      dealOwner: 'テスト太郎',
      lastModified: '2024/02/01',
      daysSinceUpdate: 7
    },
    {
      dealName: 'テスト案件2',
      amount: 500000,
      dealStage: '提案中 (souco-reborn)',
      createDate: '2024/01/15',
      closeDate: '2024/04/30',
      dealOwner: 'テスト花子',
      lastModified: '2024/02/05',
      daysSinceUpdate: 5
    }
  ];
  
  sendEmailNotification(testData);
  Logger.log('テストメールを送信しました');
}

/**
 * 初期設定関数
 * スプレッドシートのメニューに機能を追加
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('HubSpot通知')
    .addItem('今すぐチェック', 'checkAndNotifyStaleDeals')
    .addItem('テストメール送信', 'testEmailNotification')
    .addSeparator()
    .addItem('設定を表示', 'showConfig')
    .addItem('シート名を確認', 'checkSheetNames')
    .addItem('データを診断', 'debugDataCheck')
    .addItem('成約前ステータスを確認', 'checkPreContractStatuses')
    .addToUi();
}

/**
 * 設定を表示
 */
function showConfig() {
  const ui = SpreadsheetApp.getUi();
  const message = `
現在の設定:
━━━━━━━━━━━━━━━━━━
通知先: ${CONFIG.EMAIL_RECIPIENTS}
閾値: ${CONFIG.DAYS_THRESHOLD}日
━━━━━━━━━━━━━━━━━━

設定を変更する場合は、スクリプトエディタで
CONFIG変数を編集してください。
`;
  
  ui.alert('設定情報', message, ui.ButtonSet.OK);
}

/**
 * シート名を確認（デバッグ用）
 */
function checkSheetNames() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = spreadsheet.getSheets();
  const ui = SpreadsheetApp.getUi();
  
  let message = '現在のシート名一覧:\n━━━━━━━━━━━━━━━━━━\n';
  
  sheets.forEach((sheet, index) => {
    const sheetName = sheet.getName();
    message += `${index + 1}. "${sheetName}"\n`;
  });
  
  message += `━━━━━━━━━━━━━━━━━━\n\n`;
  message += `設定されているシート名:\n`;
  message += `案件シート: "${CONFIG.SHEETS.DEALS}"\n`;
  message += `ステータスマスター: "${CONFIG.SHEETS.STATUS_MASTER}"\n\n`;
  message += `※シート名が異なる場合は、CONFIG設定を更新してください。`;
  
  ui.alert('シート名確認', message, ui.ButtonSet.OK);
  
  // ログにも出力
  Logger.log('シート名一覧:');
  sheets.forEach((sheet, index) => {
    Logger.log(`${index + 1}. "${sheet.getName()}"`);
  });
}

/**
 * データ診断（デバッグ用）
 */
function debugDataCheck() {
  try {
    const spreadsheet = CONFIG.SPREADSHEET_URL 
      ? SpreadsheetApp.openByUrl(CONFIG.SPREADSHEET_URL)
      : SpreadsheetApp.getActiveSpreadsheet();
    
    const dealsSheet = spreadsheet.getSheetByName(CONFIG.SHEETS.DEALS);
    const ui = SpreadsheetApp.getUi();
    
    if (!dealsSheet) {
      ui.alert('エラー', `シート「${CONFIG.SHEETS.DEALS}」が見つかりません`, ui.ButtonSet.OK);
      return;
    }
    
    const data = dealsSheet.getDataRange().getValues();
    let message = `【${CONFIG.SHEETS.DEALS}シートのデータ診断】\n\n`;
    
    message += `データ行数: ${data.length}行\n`;
    message += `データ列数: ${data[0] ? data[0].length : 0}列\n\n`;
    
    if (data.length > 0) {
      message += `【ヘッダー行（1行目）】\n`;
      data[0].forEach((col, idx) => {
        if (col) {
          message += `${String.fromCharCode(65 + idx)}列: "${col}"\n`;
        }
      });
    }
    
    message += `\n【データサンプル（2-4行目）】\n`;
    const now = new Date();
    const thresholdDate = new Date(now.getTime() - (CONFIG.DAYS_THRESHOLD * 24 * 60 * 60 * 1000));
    
    for (let i = 1; i < Math.min(4, data.length); i++) {
      message += `\n--- ${i + 1}行目 ---\n`;
      message += `案件名(A列): ${data[i][0]}\n`;
      message += `ステータス(C列): ${data[i][2]}\n`;
      message += `最終更新(G列): ${data[i][6]}\n`;
      
      if (data[i][6]) {
        const lastModified = new Date(data[i][6]);
        const daysSince = Math.floor((now - lastModified) / (24 * 60 * 60 * 1000));
        message += `→ ${daysSince}日前に更新\n`;
        message += `→ 3日以上前?: ${lastModified < thresholdDate ? 'はい' : 'いいえ'}\n`;
      }
    }
    
    ui.alert('データ診断結果', message, ui.ButtonSet.OK);
    
  } catch (error) {
    SpreadsheetApp.getUi().alert('エラー', `診断中にエラーが発生: ${error}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * 成約前ステータスの確認
 */
function checkPreContractStatuses() {
  try {
    const spreadsheet = CONFIG.SPREADSHEET_URL 
      ? SpreadsheetApp.openByUrl(CONFIG.SPREADSHEET_URL)
      : SpreadsheetApp.getActiveSpreadsheet();
    
    const statusMasterSheet = spreadsheet.getSheetByName(CONFIG.SHEETS.STATUS_MASTER);
    const ui = SpreadsheetApp.getUi();
    
    if (!statusMasterSheet) {
      ui.alert('エラー', `シート「${CONFIG.SHEETS.STATUS_MASTER}」が見つかりません`, ui.ButtonSet.OK);
      return;
    }
    
    const statusData = statusMasterSheet.getDataRange().getValues();
    let preContractCount = 0;
    let postContractCount = 0;
    let preContractList = [];
    
    let message = `【ステータスマスター診断】\n\n`;
    message += `総行数: ${statusData.length}行\n\n`;
    
    // ヘッダー確認
    if (statusData.length > 0) {
      message += `【ヘッダー（1行目）】\n`;
      message += `A列: "${statusData[0][0]}"\n`;
      message += `B列: "${statusData[0][1]}"\n\n`;
    }
    
    // データ確認
    for (let i = 1; i < statusData.length; i++) {
      const category = statusData[i][0];
      const status = statusData[i][1];
      
      if (category === '成約前' && status) {
        preContractCount++;
        preContractList.push(status);
      } else if (category === '成約後' && status) {
        postContractCount++;
      }
    }
    
    message += `【集計結果】\n`;
    message += `成約前ステータス: ${preContractCount}個\n`;
    message += `成約後ステータス: ${postContractCount}個\n\n`;
    
    if (preContractList.length > 0) {
      message += `【成約前ステータス一覧】\n`;
      preContractList.forEach((status, idx) => {
        message += `${idx + 1}. ${status}\n`;
      });
    } else {
      message += `⚠️ 成約前ステータスが見つかりません！\n`;
      message += `A列に「成約前」と記載されているか確認してください。`;
    }
    
    ui.alert('成約前ステータス確認', message, ui.ButtonSet.OK);
    
  } catch (error) {
    SpreadsheetApp.getUi().alert('エラー', `確認中にエラーが発生: ${error}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}