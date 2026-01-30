/**
 * レポート生成とメール送信機能
 * 違反情報のレポート作成とメール通知
 */

// =================================================================
// レポート生成メイン関数
// =================================================================
function generateAndSendReport() {
  console.log('レポート生成を開始...');
  
  try {
    // 違反情報の集計
    const summary = getViolationsSummary();
    
    // 詳細データの取得
    const violations = getDetailedViolations();
    
    // HTMLレポートの作成
    const htmlReport = createHtmlReport(summary, violations);
    
    // メール送信
    sendReportEmail(htmlReport, summary);
    
    // 最終実行情報を更新
    updateLastRunInfo(summary);
    
    console.log('レポート生成完了');
    
    return true;
    
  } catch (error) {
    console.error('レポート生成エラー:', error);
    throw error;
  }
}

// =================================================================
// 詳細違反情報取得
// =================================================================
function getDetailedViolations() {
  const spreadsheet = SpreadsheetApp.openById(getSpreadsheetId());
  const violationsSheet = spreadsheet.getSheetByName('Violations_Log');
  
  if (!violationsSheet || violationsSheet.getLastRow() <= 1) {
    return [];
  }
  
  const data = violationsSheet.getRange(2, 1, violationsSheet.getLastRow() - 1, 14).getValues();
  const violations = [];
  
  for (const row of data) {
    // アクティブな違反のみ（NEWまたはONGOING）
    if (row[11] === 'NEW' || row[11] === 'ONGOING') {
      violations.push({
        violationId: row[0],
        type: row[1],
        resourceId: row[2],
        resourceName: row[3],
        fullPath: row[4],
        violationType: row[5],
        violationMessage: row[6],
        matchedRule: row[7],
        severity: row[8],
        firstDetected: row[9],
        lastConfirmed: row[10],
        status: row[11],
        daysSinceDetection: getDaysSince(row[9])
      });
    }
  }
  
  // ステータスと重要度でソート
  violations.sort((a, b) => {
    const statusOrder = { 'NEW': 0, 'ONGOING': 1 };
    const severityOrder = { 'ERROR': 0, 'WARN': 1 };
    
    if (a.status !== b.status) {
      return statusOrder[a.status] - statusOrder[b.status];
    }
    if (a.severity !== b.severity) {
      return severityOrder[a.severity] - severityOrder[b.severity];
    }
    return b.daysSinceDetection - a.daysSinceDetection;
  });
  
  return violations;
}

// =================================================================
// HTMLレポート作成
// =================================================================
function createHtmlReport(summary, violations) {
  const today = new Date().toLocaleDateString('ja-JP');
  const spreadsheetUrl = SpreadsheetApp.openById(getSpreadsheetId()).getUrl();
  
  let html = `
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <style>
    body { 
      font-family: Arial, sans-serif; 
      line-height: 1.6;
      color: #333;
      max-width: 900px;
      margin: 0 auto;
      padding: 20px;
    }
    h1 { 
      color: #2c3e50;
      border-bottom: 3px solid #3498db;
      padding-bottom: 10px;
    }
    h2 { 
      color: #34495e;
      margin-top: 30px;
      border-left: 4px solid #3498db;
      padding-left: 10px;
    }
    .summary-grid {
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
      gap: 15px;
      margin: 20px 0;
    }
    .summary-card {
      background: #f8f9fa;
      border-radius: 8px;
      padding: 15px;
      border-left: 4px solid #3498db;
    }
    .summary-card.error {
      border-left-color: #e74c3c;
      background: #fff5f5;
    }
    .summary-card.warning {
      border-left-color: #f39c12;
      background: #fffaf0;
    }
    .summary-card.new {
      border-left-color: #e74c3c;
      background: #fff5f5;
    }
    .metric-value {
      font-size: 24px;
      font-weight: bold;
      color: #2c3e50;
    }
    .metric-label {
      font-size: 12px;
      color: #7f8c8d;
      text-transform: uppercase;
    }
    table { 
      border-collapse: collapse; 
      width: 100%;
      margin-top: 20px;
      box-shadow: 0 2px 4px rgba(0,0,0,0.1);
    }
    th { 
      background-color: #34495e;
      color: white;
      padding: 12px;
      text-align: left;
      font-size: 14px;
    }
    td { 
      padding: 10px;
      border-bottom: 1px solid #ecf0f1;
      font-size: 13px;
    }
    tr:hover {
      background-color: #f5f6fa;
    }
    .status-new { 
      background-color: #ff6b6b;
      color: white;
      padding: 3px 8px;
      border-radius: 12px;
      font-size: 11px;
      font-weight: bold;
    }
    .status-ongoing { 
      background-color: #feca57;
      color: #333;
      padding: 3px 8px;
      border-radius: 12px;
      font-size: 11px;
      font-weight: bold;
    }
    .severity-error { 
      color: #e74c3c;
      font-weight: bold;
    }
    .severity-warn { 
      color: #f39c12;
      font-weight: bold;
    }
    .footer {
      margin-top: 40px;
      padding-top: 20px;
      border-top: 1px solid #ecf0f1;
      font-size: 12px;
      color: #7f8c8d;
    }
    .cta-button {
      display: inline-block;
      background: #3498db;
      color: white;
      padding: 10px 20px;
      text-decoration: none;
      border-radius: 5px;
      margin-top: 20px;
    }
    .cta-button:hover {
      background: #2980b9;
    }
  </style>
</head>
<body>
  <h1>📋 命名ルール監査レポート</h1>
  <p>実行日: ${today}</p>
  
  <h2>📊 サマリー</h2>
  <div class="summary-grid">
    <div class="summary-card">
      <div class="metric-label">総違反数</div>
      <div class="metric-value">${summary.total}</div>
    </div>
    <div class="summary-card new">
      <div class="metric-label">新規違反</div>
      <div class="metric-value">${summary.new}</div>
    </div>
    <div class="summary-card warning">
      <div class="metric-label">継続違反</div>
      <div class="metric-value">${summary.ongoing}</div>
    </div>
    <div class="summary-card">
      <div class="metric-label">解決済み</div>
      <div class="metric-value">${summary.resolved}</div>
    </div>
  </div>
  
  <h3>📊 タイプ別違反</h3>
  <div class="summary-grid">
    <div class="summary-card">
      <div class="metric-label">Slack</div>
      <div class="metric-value">${summary.byType.Slack || 0}</div>
    </div>
    <div class="summary-card">
      <div class="metric-label">Drive</div>
      <div class="metric-value">${summary.byType.Drive || 0}</div>
    </div>
    <div class="summary-card">
      <div class="metric-label">Folder</div>
      <div class="metric-value">${summary.byType.DriveFolder || 0}</div>
    </div>
  </div>
  
  <h3>⚠️ 重要度別</h3>
  <div class="summary-grid">
    <div class="summary-card error">
      <div class="metric-label">ERROR</div>
      <div class="metric-value">${summary.bySeverity.ERROR || 0}</div>
    </div>
    <div class="summary-card warning">
      <div class="metric-label">WARN</div>
      <div class="metric-value">${summary.bySeverity.WARN || 0}</div>
    </div>
  </div>
`;
  
  // 違反詳細テーブル
  if (violations.length > 0) {
    html += `
  <h2>📝 違反詳細</h2>
  <table>
    <thead>
      <tr>
        <th>ステータス</th>
        <th>タイプ</th>
        <th>名称</th>
        <th>フルパス</th>
        <th>重要度</th>
        <th>違反内容</th>
        <th>初回検出</th>
        <th>経過日数</th>
      </tr>
    </thead>
    <tbody>
`;
    
    // 最大100件まで表示
    const displayViolations = violations.slice(0, 100);
    
    for (const violation of displayViolations) {
      const statusClass = violation.status.toLowerCase();
      const severityClass = violation.severity.toLowerCase();
      const firstDetectedDate = new Date(violation.firstDetected).toLocaleDateString('ja-JP');
      
      html += `
      <tr>
        <td><span class="status-${statusClass}">${violation.status}</span></td>
        <td>${violation.type}</td>
        <td><strong>${escapeHtml(violation.resourceName)}</strong></td>
        <td>${escapeHtml(violation.fullPath)}</td>
        <td><span class="severity-${severityClass}">${violation.severity}</span></td>
        <td>${escapeHtml(violation.violationMessage)}</td>
        <td>${firstDetectedDate}</td>
        <td>${violation.daysSinceDetection}日</td>
      </tr>
`;
    }
    
    html += `
    </tbody>
  </table>
`;
    
    if (violations.length > 100) {
      html += `
  <p><em>※ 表示されているのは先頭100件です。全件を確認するにはスプレッドシートをご覧ください。</em></p>
`;
    }
  } else {
    html += `
  <h2>✅ 違反なし</h2>
  <p>現在、アクティブな命名ルール違反はありません。</p>
`;
  }
  
  // フッター
  html += `
  <div class="footer">
    <a href="${spreadsheetUrl}" class="cta-button">📋 詳細をスプレッドシートで確認</a>
    <p>
      このレポートは自動生成されました。<br>
      問い合わせ: IT管理部門<br>
      次回実行: 翌営業日 8:30
    </p>
  </div>
</body>
</html>
`;
  
  return html;
}

// =================================================================
// メール送信
// =================================================================
function sendReportEmail(htmlReport, summary) {
  const config = getConfig();
  const recipients = config.notificationEmail;
  
  if (!recipients || recipients.trim() === '') {
    console.log('通知先メールが設定されていないため、メール送信をスキップ');
    return;
  }
  
  const today = new Date().toLocaleDateString('ja-JP');
  const subject = `[命名監査] ${today} Slack/Drive/Folder 違反 ${summary.total}件`;
  
  // テキスト版も作成
  const textBody = createTextReport(summary);
  
  // メール送信
  MailApp.sendEmail({
    to: recipients,
    subject: subject,
    body: textBody,
    htmlBody: htmlReport,
    name: 'Naming Audit System'
  });
  
  console.log(`レポートを送信しました: ${recipients}`);
}

// =================================================================
// テキストレポート作成
// =================================================================
function createTextReport(summary) {
  const today = new Date().toLocaleDateString('ja-JP');
  const spreadsheetUrl = SpreadsheetApp.openById(getSpreadsheetId()).getUrl();
  
  let text = `
命名ルール監査レポート
========================
実行日: ${today}

[サマリー]
- 総違反数: ${summary.total}件
- 新規違反: ${summary.new}件
- 継続違反: ${summary.ongoing}件
- 解決済み: ${summary.resolved}件

[タイプ別]
- Slack: ${summary.byType.Slack || 0}件
- Drive: ${summary.byType.Drive || 0}件
- Folder: ${summary.byType.DriveFolder || 0}件

[重要度別]
- ERROR: ${summary.bySeverity.ERROR || 0}件
- WARN: ${summary.bySeverity.WARN || 0}件

詳細は以下のスプレッドシートをご確認ください:
${spreadsheetUrl}

---
このレポートは自動生成されました。
`;
  
  return text;
}

// =================================================================
// 最終実行情報更新
// =================================================================
function updateLastRunInfo(summary) {
  const spreadsheet = SpreadsheetApp.openById(getSpreadsheetId());
  const sheet = spreadsheet.getSheetByName('Report_LastRun');
  
  if (!sheet) {
    console.error('Report_LastRun シートが見つかりません');
    return;
  }
  
  const now = new Date();
  const data = [
    ['最終実行日時', now],
    ['実行ステータス', 'SUCCESS'],
    ['Slackチャンネル取得数', getSheetRowCount('Slack_Channels')],
    ['Drive取得数', getSheetRowCount('Drive_SharedDrives')],
    ['フォルダ取得数', getSheetRowCount('Drive_Folders')],
    ['新規違反検出数', summary.new],
    ['継続違反数', summary.ongoing],
    ['解決済み違反数', summary.resolved],
    ['エラー内容', ''],
    ['実行時間(秒)', ''] // 実行時間はメイン関数で計測
  ];
  
  sheet.getRange(2, 1, data.length, 2).setValues(data);
}

// =================================================================
// ユーティリティ関数
// =================================================================
function getDaysSince(date) {
  if (!date) return 0;
  const now = new Date();
  const then = new Date(date);
  const diffTime = Math.abs(now - then);
  const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
  return diffDays;
}

function escapeHtml(text) {
  const map = {
    '&': '&amp;',
    '<': '&lt;',
    '>': '&gt;',
    '"': '&quot;',
    "'": '&#039;'
  };
  return String(text).replace(/[&<>"']/g, m => map[m]);
}

function getSheetRowCount(sheetName) {
  try {
    const spreadsheet = SpreadsheetApp.openById(getSpreadsheetId());
    const sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) return 0;
    const lastRow = sheet.getLastRow();
    return lastRow > 1 ? lastRow - 1 : 0; // ヘッダーを除く
  } catch (error) {
    return 0;
  }
}