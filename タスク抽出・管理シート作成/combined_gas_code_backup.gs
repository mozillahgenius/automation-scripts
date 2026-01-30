/*
================================================================================
                    タスク抽出・管理シート作成 - 統合ファイル
================================================================================

【目次 - Table of Contents】
1. config.gs - 基本設定とユーティリティ機能
2. flow_visualizer.gs - フロービジュアライザー機能
3. gmail_inbound.gs - Gmail受信処理機能
4. gmail_outbound.gs - Gmail送信処理機能
5. menu.gs - カスタムメニューとセットアップ機能
6. openai_client.gs - OpenAI API連携機能
7. parser_and_writer.gs - データ解析・書き込み処理機能

作成日: 2025-08-16
説明: Google Apps Scriptによるタスク管理システムの全ソースコード

================================================================================
*/

// ================================================================================
// 1. config.gs - 基本設定とユーティリティ機能
// ================================================================================

// 基本定数
const CONFIG_SHEET = 'Config';
const INBOX_SHEET = 'Inbox';
const SPEC_SHEET = '業務記述書';
const FLOW_SHEET = 'フロー';
const VISUAL_SHEET = 'ビジュアルフロー';

// CSV行を解析する関数
function parseCSVLine(line) {
  const result = [];
  let current = '';
  let inQuotes = false;
  
  for (let i = 0; i < line.length; i++) {
    const char = line[i];
    const nextChar = line[i + 1];
    
    if (char === '"') {
      if (inQuotes && nextChar === '"') {
        // エスケープされた引用符
        current += '"';
        i++; // 次の引用符をスキップ
      } else {
        // 引用符の開始/終了
        inQuotes = !inQuotes;
      }
    } else if (char === ',' && !inQuotes) {
      // フィールドの区切り
      result.push(current.trim());
      current = '';
    } else {
      current += char;
    }
  }
  
  // 最後のフィールドを追加
  result.push(current.trim());
  
  return result;
}
const ACTIVITY_LOG_SHEET = 'ActivityLog';

// スプレッドシート関連ユーティリティ
function ss() {
  return SpreadsheetApp.getActive();
}

function file() {
  return DriveApp.getFileById(ss().getId());
}

// Config管理
function getConfig(key) {
  const sh = ss().getSheetByName(CONFIG_SHEET);
  if (!sh) return null;
  
  const values = sh.getRange(1, 1, sh.getLastRow(), 2).getValues();
  const configMap = new Map(values.map(r => [String(r[0]).trim(), String(r[1]).trim()]));
  return configMap.get(key);
}

function setConfig(key, value) {
  const sh = ss().getSheetByName(CONFIG_SHEET);
  if (!sh) return;
  
  const values = sh.getRange(1, 1, sh.getLastRow(), 2).getValues();
  let found = false;
  
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]).trim() === key) {
      sh.getRange(i + 1, 2).setValue(value);
      found = true;
      break;
    }
  }
  
  if (!found) {
    sh.appendRow([key, value]);
  }
}

// 初期Config設定
function initializeConfig() {
  const sh = ss().getSheetByName(CONFIG_SHEET) || ss().insertSheet(CONFIG_SHEET);
  
  const defaultConfigs = [
    ['PROCESSING_QUERY', 'label:inbox is:unread subject:"[WORK-REQ]"'],
    ['DEFAULT_TO_EMAIL', ''],
    ['OPENAI_MODEL', 'gpt-4o-mini'],
    ['ORG_PROFILE_JSON', '{"listing":"上場区分","industry":"業種","jurisdictions":["JP"],"policies":["内部統制準拠"]}'],
    ['SHARE_ANYONE_WITH_LINK', 'TRUE'],
    ['FLOW_SHEET_NAME', 'フロー'],
    ['VISUAL_SHEET_NAME', 'ビジュアルフロー'],
    ['LEGAL_JURISDICTIONS', 'JP, global'],
    ['MAX_RETRY_COUNT', '3'],
    ['RETRY_DELAY_MS', '2000']
  ];
  
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, defaultConfigs.length, 2).setValues(defaultConfigs);
  }
}

// 共有設定
function shareSheetAnyWithLink() {
  try {
    file().setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.EDIT);
    logActivity('SHARE', 'Set ANYONE_WITH_LINK permissions');
    return true;
  } catch (e) {
    logActivity('SHARE_ERROR', `Failed to set ANYONE_WITH_LINK: ${e.toString()}`);
    return false;
  }
}

function addEditor(email) {
  try {
    file().addEditor(email);
    logActivity('SHARE', `Added editor: ${email}`);
    return true;
  } catch (e) {
    logActivity('SHARE_ERROR', `Failed to add editor ${email}: ${e.toString()}`);
    return false;
  }
}

// メールアドレス抽出
function extractEmail(fromHeader) {
  const match = fromHeader.match(/<([^>]+)>/);
  return match ? match[1] : fromHeader.replace(/.*\s/, '').trim();
}

// HTMLをテキストに変換
function htmlToText(html) {
  return html
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<\/p>/gi, '\n\n')
    .replace(/<[^>]+>/g, ' ')
    .replace(/&nbsp;/g, ' ')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&amp;/g, '&')
    .replace(/\s+/g, ' ')
    .trim();
}

// メール署名を除去
function removeEmailSignature(text) {
  if (!text) return text;
  
  // 一般的な署名の開始パターン
  const signaturePatterns = [
    /--\s*\n/,  // -- で始まる署名
    /—\s*\n/,   // — で始まる署名
    /＿+\s*\n/, // アンダースコアの連続
    /━+\s*\n/,  // 罫線
    /※この.*$/s, // ※この...で始まる免責事項
    /\n\n(敬具|よろしくお願い|Best regards|Regards|Sincerely|Thanks)[\s\S]*$/i,
    /\n\n-{3,}[\s\S]*$/, // 3つ以上のハイフン
    /\n\n_{3,}[\s\S]*$/, // 3つ以上のアンダースコア
    /\n\n={3,}[\s\S]*$/, // 3つ以上のイコール
    /\n\n\*{3,}[\s\S]*$/  // 3つ以上のアスタリスク
  ];
  
  let cleanedText = text;
  
  // 各パターンでマッチする最初の位置を探す
  let earliestIndex = text.length;
  for (const pattern of signaturePatterns) {
    const match = text.match(pattern);
    if (match && match.index < earliestIndex) {
      earliestIndex = match.index;
    }
  }
  
  // 署名部分を除去
  if (earliestIndex < text.length) {
    cleanedText = text.substring(0, earliestIndex).trim();
  }
  
  // 連絡先情報のパターンも除去（メールアドレス、電話番号が連続する部分）
  const contactPattern = /(\n.*[@].*\n.*[0-9-()]+.*\n)/g;
  cleanedText = cleanedText.replace(contactPattern, '\n');
  
  return cleanedText.trim();
}

// 処理済みチェック
function isProcessed(messageId) {
  const sh = ss().getSheetByName(INBOX_SHEET);
  if (!sh || sh.getLastRow() <= 1) return false;
  
  const lastRow = sh.getLastRow();
  const dataRows = Math.max(1, lastRow - 1);
  const values = sh.getRange(2, 3, dataRows, 1).getValues();
  return values.some(row => row[0] === messageId);
}

function markProcessed(messageId) {
  const sh = ss().getSheetByName(INBOX_SHEET);
  if (!sh || sh.getLastRow() <= 1) return;
  
  const lastRow = sh.getLastRow();
  const dataRows = Math.max(1, lastRow - 1);
  const values = sh.getRange(2, 3, dataRows, 1).getValues();
  for (let i = 0; i < values.length; i++) {
    if (values[i][0] === messageId) {
      sh.getRange(i + 2, 7).setValue('PROCESSED');
      sh.getRange(i + 2, 8).setValue(new Date());
      break;
    }
  }
}

// Inboxログ記録
function logInbox(messageId, threadId, from, subject, summary, status) {
  const sh = ss().getSheetByName(INBOX_SHEET) || createInboxSheet();
  sh.appendRow([
    new Date(),
    threadId,
    messageId,
    from,
    subject,
    summary,
    status,
    status === 'PROCESSED' ? new Date() : '',
    ''
  ]);
}

function createInboxSheet() {
  const sh = ss().insertSheet(INBOX_SHEET);
  sh.getRange(1, 1, 1, 9).setValues([[
    '受信日時', 'ThreadId', 'MessageId', 'From', 'Subject', 
    '要約', 'ステータス', '処理日時', 'エラー'
  ]]);
  sh.getRange(1, 1, 1, 9).setFontWeight('bold');
  return sh;
}

// エラーログ記録
function logError(messageId, error) {
  const sh = ss().getSheetByName(INBOX_SHEET);
  if (!sh || sh.getLastRow() <= 1) return;
  
  const lastRow = sh.getLastRow();
  const dataRows = Math.max(1, lastRow - 1);
  const values = sh.getRange(2, 3, dataRows, 1).getValues();
  for (let i = 0; i < values.length; i++) {
    if (values[i][0] === messageId) {
      sh.getRange(i + 2, 7).setValue('ERROR');
      sh.getRange(i + 2, 9).setValue(error.toString());
      break;
    }
  }
  
  logActivity('ERROR', `MessageId: ${messageId}, Error: ${error.toString()}`);
}

// アクティビティログ
function logActivity(type, details) {
  const sh = ss().getSheetByName(ACTIVITY_LOG_SHEET) || createActivityLogSheet();
  let userEmail;
  try {
    userEmail = Session.getActiveUser().getEmail();
  } catch (e) {
    userEmail = 'Unknown User';
  }
  sh.appendRow([
    new Date(),
    type,
    details,
    userEmail
  ]);
}

function createActivityLogSheet() {
  const sh = ss().insertSheet(ACTIVITY_LOG_SHEET);
  sh.getRange(1, 1, 1, 4).setValues([[
    'タイムスタンプ', 'タイプ', '詳細', 'ユーザー'
  ]]);
  sh.getRange(1, 1, 1, 4).setFontWeight('bold');
  sh.hideSheet();
  return sh;
}

// リトライ処理
function retryWithBackoff(func, maxRetries = 3) {
  const configRetries = parseInt(getConfig('MAX_RETRY_COUNT') || '3');
  const retryDelay = parseInt(getConfig('RETRY_DELAY_MS') || '2000');
  maxRetries = configRetries;
  
  for (let i = 0; i < maxRetries; i++) {
    try {
      return func();
    } catch (e) {
      if (i === maxRetries - 1) throw e;
      Utilities.sleep(retryDelay * Math.pow(2, i));
    }
  }
}

// ================================================================================
// 2. flow_visualizer.gs - フロービジュアライザー機能
// ================================================================================

// 高度なカラーパレット定義
const ADVANCED_COLORS = {
  // ヘッダー系
  MAIN_HEADER: '#2C3E50',
  SUB_HEADER: '#34495E',
  SECTION_HEADER: '#5D6D7E',
  
  // プロセス系
  START_END: '#27AE60',
  PROCESS: '#3498DB',
  DECISION: '#F39C12',
  SUBPROCESS: '#9B59B6',
  
  // 背景系
  TIMELINE_BG: '#ECF0F1',
  DEPT_BG: '#E8F5E9',
  EMPTY_BG: '#FAFAFA',
  HIGHLIGHT_BG: '#FFF9C4',
  
  // ツール系
  TOOL_BG: '#E3F2FD',
  DATASOURCE_BG: '#FFF3E0',
  
  // ステータス系
  SUCCESS: '#4CAF50',
  WARNING: '#FF9800',
  ERROR: '#F44336',
  INFO: '#2196F3'
};

// 部署別カラーパレット
const DEPT_COLOR_PALETTE = [
  '#E3F2FD', '#FCE4EC', '#F3E5F5', '#EDE7F6', '#E8EAF6',
  '#E1F5FE', '#E0F2F1', '#E8F5E9', '#F9FBE7', '#FFF8E1',
  '#FFF3E0', '#FBE9E7', '#EFEBE9', '#FAFAFA', '#ECEFF1'
];

// ツール別アイコンとカラー
const TOOL_ICONS = {
  'Word': { icon: '📝', color: '#2B579A' },
  'Excel': { icon: '📊', color: '#217346' },
  'PowerPoint': { icon: '📰', color: '#D24726' },
  'PPT': { icon: '📰', color: '#D24726' },
  'Teams': { icon: '👥', color: '#5B5FC7' },
  'Outlook': { icon: '📧', color: '#0078D4' },
  'Gmail': { icon: '📨', color: '#EA4335' },
  'Slack': { icon: '💬', color: '#4A154B' },
  'GitHub': { icon: '🐙', color: '#24292E' },
  'Jira': { icon: '📋', color: '#0052CC' },
  'Notion': { icon: '📓', color: '#000000' },
  'Google Drive': { icon: '☁️', color: '#4285F4' },
  'Zoom': { icon: '📹', color: '#2D8CFF' },
  'メール': { icon: '✉️', color: '#EA4335' },
  'ブラウザ': { icon: '🌐', color: '#4CAF50' },
  'システム': { icon: '⚙️', color: '#607D8B' },
  'データベース': { icon: '🗄️', color: '#FF6F00' }
};

// フロービジュアライザー - ビジュアルフロー生成機能

function generateVisualFlow() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const flowSheet = ss.getSheetByName(getConfig('FLOW_SHEET_NAME') || 'フロー');
  const visualSheet = ss.getSheetByName(getConfig('VISUAL_SHEET_NAME') || 'ビジュアルフロー') || 
                      ss.insertSheet(getConfig('VISUAL_SHEET_NAME') || 'ビジュアルフロー');
  
  if (!flowSheet) {
    SpreadsheetApp.getUi().alert('フローシートが見つかりません。');
    return;
  }
  
  // ビジュアルフローシートをクリア
  visualSheet.clear();
  visualSheet.clearFormats();
  
  // フローデータを取得
  const flowData = flowSheet.getDataRange().getValues();
  if (flowData.length <= 1) {
    SpreadsheetApp.getUi().alert('フローデータがありません。');
    return;
  }
  
  const headers = flowData[0];
  const rows = flowData.slice(1).filter(row => row[0]); // 工程が入力されている行のみ
  
  if (rows.length === 0) {
    SpreadsheetApp.getUi().alert('有効なフローデータがありません。');
    return;
  }
  
  // 部署リストを作成
  const departments = [...new Set(rows.map(row => row[2]).filter(d => d))];
  
  // ビジュアルフローのレイアウト設定
  const startRow = 3;
  const startCol = 2;
  const boxWidth = 3;
  const boxHeight = 3;
  const horizontalGap = 1;
  const verticalGap = 1;
  
  // タイトル設定
  visualSheet.getRange(1, 1).setValue('業務フロー図');
  visualSheet.getRange(1, 1).setFontSize(16).setFontWeight('bold');
  
  // 部署別のレーン作成
  let currentCol = startCol;
  const deptColumns = {};
  
  departments.forEach((dept, index) => {
    const deptCol = currentCol + index * (boxWidth + horizontalGap);
    deptColumns[dept] = deptCol;
    
    // 部署名を表示
    visualSheet.getRange(startRow - 1, deptCol, 1, boxWidth).merge();
    visualSheet.getRange(startRow - 1, deptCol).setValue(dept);
    visualSheet.getRange(startRow - 1, deptCol).setBackground('#e8eaf6')
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setBorder(true, true, true, true, true, false);
  });
  
  // フローボックスの描画
  let currentRow = startRow + 1;
  const processedSteps = [];
  
  rows.forEach((row, rowIndex) => {
    const step = row[0]; // 工程
    const timing = row[1]; // 実施タイミング
    const dept = row[2]; // 部署
    const role = row[3]; // 担当役割
    const task = row[4]; // 作業内容
    const condition = row[5]; // 条件分岐
    const tool = row[6]; // 利用ツール
    const url = row[7]; // URLリンク
    const note = row[8]; // 備考
    
    const col = deptColumns[dept] || startCol;
    const currentRowPos = currentRow;
    
    // ボックスのセル範囲を取得
    const boxRange = visualSheet.getRange(currentRowPos, col, boxHeight, boxWidth);
    boxRange.merge();
    
    // ボックスの内容設定
    let boxContent = `【${step}】\n${task}`;
    if (role) boxContent += `\n(${role})`;
    if (tool) boxContent += `\n[${tool}]`;
    
    boxRange.setValue(boxContent);
    
    // ボックスのスタイル設定
    if (condition) {
      // 条件分岐は菱形風に黄色背景
      boxRange.setBackground('#fff9c4')
        .setBorder(true, true, true, true, false, false, '#ff9800', SpreadsheetApp.BorderStyle.SOLID_THICK);
    } else if (rowIndex === 0) {
      // 開始は緑背景
      boxRange.setBackground('#c8e6c9')
        .setBorder(true, true, true, true, false, false, '#4caf50', SpreadsheetApp.BorderStyle.SOLID_THICK);
    } else if (rowIndex === rows.length - 1) {
      // 終了は赤背景
      boxRange.setBackground('#ffcdd2')
        .setBorder(true, true, true, true, false, false, '#f44336', SpreadsheetApp.BorderStyle.SOLID_THICK);
    } else {
      // 通常処理は青背景
      boxRange.setBackground('#e3f2fd')
        .setBorder(true, true, true, true, false, false, '#2196f3', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    }
    
    boxRange.setWrap(true)
      .setVerticalAlignment('middle')
      .setHorizontalAlignment('center')
      .setFontSize(10);
    
    // URLリンクがある場合
    if (url) {
      const linkRange = visualSheet.getRange(row + boxHeight, col, 1, boxWidth);
      linkRange.merge();
      linkRange.setValue('📎 リンク');
      linkRange.setFormula(`=HYPERLINK("${url}", "📎 リンク")`);
      linkRange.setFontSize(9).setFontColor('#1a73e8');
    }
    
    // 備考がある場合
    if (note) {
      const noteRange = visualSheet.getRange(row, col + boxWidth + 1);
      noteRange.setValue(`💡 ${note}`);
      noteRange.setFontSize(9).setFontColor('#666').setWrap(true);
    }
    
    // 矢印の描画（次のステップがある場合）
    if (rowIndex < rows.length - 1) {
      const nextDept = rows[rowIndex + 1][2];
      const nextCol = deptColumns[nextDept] || startCol;
      
      if (col === nextCol) {
        // 同じ部署内での移動（下向き矢印）
        const arrowRange = visualSheet.getRange(row + boxHeight, col + Math.floor(boxWidth / 2));
        arrowRange.setValue('↓');
        arrowRange.setFontSize(16).setHorizontalAlignment('center');
      } else {
        // 異なる部署への移動（横向き矢印）
        const direction = nextCol > col ? '→' : '←';
        const arrowCol = col < nextCol ? col + boxWidth : col - 1;
        const arrowRange = visualSheet.getRange(row + Math.floor(boxHeight / 2), arrowCol);
        arrowRange.setValue(direction);
        arrowRange.setFontSize(16).setHorizontalAlignment('center');
      }
    }
    
    processedSteps.push({
      step: step,
      row: row,
      col: col,
      dept: dept,
      condition: condition
    });
    
    currentRow += boxHeight + verticalGap + 1;
  });
  
  // 凡例の追加
  const legendRow = currentRow + 3;
  visualSheet.getRange(legendRow, startCol).setValue('【凡例】');
  visualSheet.getRange(legendRow, startCol).setFontWeight('bold');
  
  const legends = [
    { color: '#c8e6c9', text: '開始', border: '#4caf50' },
    { color: '#e3f2fd', text: '通常処理', border: '#2196f3' },
    { color: '#fff9c4', text: '条件分岐', border: '#ff9800' },
    { color: '#ffcdd2', text: '終了', border: '#f44336' }
  ];
  
  legends.forEach((legend, index) => {
    const legendCol = startCol + index * 3;
    const legendRange = visualSheet.getRange(legendRow + 1, legendCol, 1, 2);
    legendRange.merge();
    legendRange.setValue(legend.text);
    legendRange.setBackground(legend.color)
      .setBorder(true, true, true, true, false, false, legend.border, SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
      .setHorizontalAlignment('center');
  });
  
  // 列幅と行高の調整
  for (let i = 1; i <= visualSheet.getMaxColumns(); i++) {
    visualSheet.setColumnWidth(i, 120);
  }
  
  for (let i = startRow; i <= currentRow; i++) {
    visualSheet.setRowHeight(i, 60);
  }
  
  // シート全体の書式設定
  visualSheet.getRange(1, 1, visualSheet.getMaxRows(), visualSheet.getMaxColumns())
    .setFontFamily('Noto Sans JP');
  
  logActivity('VISUAL_FLOW', 'Visual flow generated successfully');
  
  SpreadsheetApp.getUi().alert('ビジュアルフローを生成しました。');
}

// サンプルデータ作成（開発・テスト用）
function createSampleFlowData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const flowSheet = ss.getSheetByName('フロー') || ss.insertSheet('フロー');
  
  // シートをクリア
  flowSheet.clear();
  
  // ヘッダー設定
  const headers = [
    '工程', '実施タイミング', '部署', '担当役割', '作業内容', 
    '条件分岐', '利用ツール', 'URLリンク', '備考'
  ];
  
  flowSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  flowSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#e8f5e9');
  
  // サンプルデータ
  const sampleData = [
    ['要件定義', 'Day 1-5', '企画部', 'プロジェクトマネージャー', '業務要件のヒアリングと整理', '', 'Teams, Miro', 'https://example.com/requirements', '関係者全員参加必須'],
    ['承認判断', 'Day 6', '経営企画部', '部長', '要件の承認可否を判断', '承認/差戻し', '', '', '予算上限確認'],
    ['基本設計', 'Day 7-15', 'IT部', 'システムアーキテクト', 'システム構成の設計', '', 'draw.io, Confluence', 'https://example.com/design', ''],
    ['詳細設計', 'Day 16-25', 'IT部', '開発リード', '機能仕様の詳細化', '', 'GitHub, Figma', '', 'UI/UXチームと連携'],
    ['開発', 'Day 26-50', '開発部', '開発チーム', 'コーディングと単体テスト', '', 'VS Code, Git', 'https://github.com/example', 'アジャイル開発'],
    ['品質チェック', 'Day 51-55', '品質管理部', 'QAエンジニア', 'テスト実施と不具合修正', '合格/再テスト', 'Selenium, JIRA', '', ''],
    ['リリース準備', 'Day 56-58', 'IT部', 'インフラチーム', '本番環境へのデプロイ準備', '', 'Jenkins, Docker', '', ''],
    ['本番リリース', 'Day 59', 'IT部', 'リリースマネージャー', '本番環境への展開', '', 'Kubernetes', '', '夜間作業'],
    ['運用引継ぎ', 'Day 60', '運用部', '運用チーム', '運用手順書の確認と引継ぎ', '', 'ServiceNow', 'https://example.com/operations', '24時間体制確立']
  ];
  
  flowSheet.getRange(2, 1, sampleData.length, headers.length).setValues(sampleData);
  
  // 列幅調整
  flowSheet.setColumnWidth(1, 100); // 工程
  flowSheet.setColumnWidth(2, 120); // 実施タイミング
  flowSheet.setColumnWidth(3, 100); // 部署
  flowSheet.setColumnWidth(4, 150); // 担当役割
  flowSheet.setColumnWidth(5, 250); // 作業内容
  flowSheet.setColumnWidth(6, 150); // 条件分岐
  flowSheet.setColumnWidth(7, 120); // 利用ツール
  flowSheet.setColumnWidth(8, 200); // URLリンク
  flowSheet.setColumnWidth(9, 200); // 備考
  
  SpreadsheetApp.getUi().alert('サンプルフローデータを作成しました。');
}

// 高度なビジュアルフロー生成（業務フローチャート図作成.js参考版）
function generateAdvancedVisualFlow() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(FLOW_SHEET);
    if (!sheet) {
      SpreadsheetApp.getUi().alert('エラー', 'フローシートが見つかりません。', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) {
      SpreadsheetApp.getUi().alert('エラー', 'データがありません。', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    // ビジュアルシートの準備
    const visualSheet = getOrCreateSheet(VISUAL_SHEET);
    visualSheet.clear();
    visualSheet.clearFormats();
    
    // データの解析
    const flowData = parseAdvancedFlowData(data);
    
    // フローチャートの描画
    drawAdvancedFlowChart(visualSheet, flowData);
    
    // 業務サマリーシート作成
    createBusinessSummarySheet(flowData);
    
    SpreadsheetApp.getUi().alert('成功', '高度なビジュアルフローが生成されました。', SpreadsheetApp.getUi().ButtonSet.OK);
    
  } catch (error) {
    console.error('エラー:', error);
    SpreadsheetApp.getUi().alert('エラー', 'フロー生成中にエラーが発生しました: ' + error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

// 高度なフローデータ解析
function parseAdvancedFlowData(data) {
  const headers = data[0];
  const columnIndex = {};
  headers.forEach((header, index) => {
    columnIndex[header] = index;
  });
  
  const flowData = {
    departments: {},
    departmentList: [],
    timings: [],
    tools: new Set(),
    datasources: {},
    processName: "",
    statistics: {
      totalTasks: 0,
      totalDepartments: 0,
      totalTools: 0,
      decisionPoints: 0
    }
  };
  
  // プロセス名の取得
  if (data.length > 1 && data[1][columnIndex["工程"]]) {
    flowData.processName = data[1][columnIndex["工程"]];
  }
  
  // データの整理
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[columnIndex["工程"]] || row[columnIndex["工程"]] === "") continue;
    
    const dept = row[columnIndex["部署"]] || "その他";
    const timing = row[columnIndex["実施タイミング"]] || "";
    const tool = row[columnIndex["利用ツール"]] || "";
    const url = row[columnIndex["URLリンク"]] || "";
    const condition = row[columnIndex["条件分岐"]] || "";
    
    // 部署の初期化
    if (!flowData.departments[dept]) {
      flowData.departments[dept] = {};
      if (!flowData.departmentList.includes(dept)) {
        flowData.departmentList.push(dept);
      }
    }
    
    // タイミングの追加
    if (timing && !flowData.timings.includes(timing)) {
      flowData.timings.push(timing);
    }
    
    // タスクの追加
    if (!flowData.departments[dept][timing]) {
      flowData.departments[dept][timing] = [];
    }
    
    flowData.departments[dept][timing].push({
      task: row[columnIndex["作業内容"]] || "",
      role: row[columnIndex["担当役割"]] || "",
      condition: condition,
      tool: tool,
      url: url,
      note: row[columnIndex["備考"]] || ""
    });
    
    // 統計更新
    flowData.statistics.totalTasks++;
    if (condition && condition !== "-") {
      flowData.statistics.decisionPoints++;
    }
    
    // ツールの収集
    if (tool && tool !== "-") {
      const tools = tool.split(/[／、,]/);
      tools.forEach(t => {
        const trimmedTool = t.trim();
        if (trimmedTool) {
          flowData.tools.add(trimmedTool);
        }
      });
    }
    
    // データソース管理
    if (url && url !== "-") {
      if (!flowData.datasources[timing]) {
        flowData.datasources[timing] = [];
      }
      if (!flowData.datasources[timing].includes(url)) {
        flowData.datasources[timing].push(url);
      }
    }
  }
  
  flowData.statistics.totalDepartments = flowData.departmentList.length;
  flowData.statistics.totalTools = flowData.tools.size;
  
  return flowData;
}

// 高度なフローチャート描画
function drawAdvancedFlowChart(sheet, flowData) {
  let currentRow = 1;
  const maxCols = Math.max(flowData.departmentList.length + 2, 10);
  
  // タイトル行
  const flowTitle = flowData.processName || "業務フロー";
  sheet.getRange(currentRow, 1, 1, maxCols).merge();
  const titleCell = sheet.getRange(currentRow, 1);
  titleCell.setValue(flowTitle + "（" + new Date().getFullYear() + "年" + (new Date().getMonth() + 1) + "月版）");
  titleCell.setBackground(ADVANCED_COLORS.MAIN_HEADER);
  titleCell.setFontColor("#FFFFFF");
  titleCell.setFontSize(18);
  titleCell.setFontWeight("bold");
  titleCell.setHorizontalAlignment("center");
  titleCell.setVerticalAlignment("middle");
  sheet.setRowHeight(currentRow, 50);
  currentRow++;
  
  // 統計情報行
  sheet.getRange(currentRow, 1, 1, maxCols).merge();
  const statsCell = sheet.getRange(currentRow, 1);
  statsCell.setValue(`📊 総タスク: ${flowData.statistics.totalTasks} | 👥 部署: ${flowData.statistics.totalDepartments} | 🔧 ツール: ${flowData.statistics.totalTools} | ⚡ 判断: ${flowData.statistics.decisionPoints}`);
  statsCell.setBackground(ADVANCED_COLORS.SUB_HEADER);
  statsCell.setFontColor("#FFFFFF");
  statsCell.setHorizontalAlignment("center");
  sheet.setRowHeight(currentRow, 35);
  currentRow++;
  
  // ツール行
  if (flowData.tools.size > 0) {
    drawAdvancedToolsRow(sheet, currentRow, flowData.tools, maxCols);
    currentRow++;
  }
  
  // ヘッダー行
  drawAdvancedHeaderRow(sheet, currentRow, flowData.departmentList, Object.keys(flowData.datasources).length > 0);
  currentRow++;
  
  // 開始行
  drawAdvancedStartRow(sheet, currentRow, maxCols);
  currentRow++;
  
  // 各タイミングの行
  flowData.timings.forEach((timing, index) => {
    drawAdvancedTimingRow(sheet, currentRow, timing, flowData, index);
    currentRow++;
  });
  
  // 終了行
  drawAdvancedEndRow(sheet, currentRow, flowData.departmentList.length, Object.keys(flowData.datasources).length > 0);
  currentRow++;
  
  // 凡例行
  drawAdvancedLegendRow(sheet, currentRow, maxCols);
  
  // 列幅の調整
  adjustAdvancedColumnWidths(sheet, flowData.departmentList.length, Object.keys(flowData.datasources).length > 0);
  
  // 罫線の設定
  applyAdvancedBorders(sheet, currentRow, maxCols);
}

// ツール行の描画（高度版）
function drawAdvancedToolsRow(sheet, row, tools, maxCols) {
  sheet.getRange(row, 1).setValue("使用ツール");
  sheet.getRange(row, 1).setBackground(ADVANCED_COLORS.SECTION_HEADER);
  sheet.getRange(row, 1).setFontColor("#FFFFFF");
  sheet.getRange(row, 1).setFontWeight("bold");
  
  sheet.getRange(row, 2, 1, maxCols - 1).merge();
  const toolCell = sheet.getRange(row, 2);
  toolCell.setBackground(ADVANCED_COLORS.TOOL_BG);
  
  let toolText = "";
  tools.forEach(tool => {
    const toolInfo = TOOL_ICONS[tool] || { icon: '🔧', color: '#666666' };
    toolText += ` ${toolInfo.icon} ${tool} `;
  });
  
  toolCell.setValue(toolText);
  toolCell.setHorizontalAlignment("left");
  sheet.setRowHeight(row, 35);
}

// ヘッダー行の描画（高度版）
function drawAdvancedHeaderRow(sheet, row, departments, hasDataSource) {
  // 日程列
  sheet.getRange(row, 1).setValue("タイミング");
  sheet.getRange(row, 1).setBackground(ADVANCED_COLORS.SECTION_HEADER);
  sheet.getRange(row, 1).setFontColor("#FFFFFF");
  sheet.getRange(row, 1).setFontWeight("bold");
  
  // 部署列
  departments.forEach((dept, index) => {
    const col = index + 2;
    const cell = sheet.getRange(row, col);
    cell.setValue(dept);
    cell.setBackground(DEPT_COLOR_PALETTE[index % DEPT_COLOR_PALETTE.length]);
    cell.setFontWeight("bold");
    cell.setWrap(true);
    cell.setHorizontalAlignment("center");
    cell.setVerticalAlignment("middle");
  });
  
  // データソース列
  if (hasDataSource) {
    const dataCol = departments.length + 2;
    sheet.getRange(row, dataCol).setValue("📚 関連資料");
    sheet.getRange(row, dataCol).setBackground(ADVANCED_COLORS.DATASOURCE_BG);
    sheet.getRange(row, dataCol).setFontWeight("bold");
  }
  
  sheet.setRowHeight(row, 50);
}

// 開始行の描画（高度版）
function drawAdvancedStartRow(sheet, row, maxCols) {
  sheet.getRange(row, 1, 1, maxCols).merge();
  const cell = sheet.getRange(row, 1);
  cell.setValue("🚀 【プロセス開始】");
  cell.setBackground(ADVANCED_COLORS.START_END);
  cell.setFontColor("#FFFFFF");
  cell.setFontWeight("bold");
  cell.setFontSize(14);
  cell.setHorizontalAlignment("center");
  cell.setBorder(true, true, true, true, false, false, "#228B22", SpreadsheetApp.BorderStyle.SOLID_THICK);
  sheet.setRowHeight(row, 40);
}

// タイミング行の描画（高度版）
function drawAdvancedTimingRow(sheet, row, timing, flowData, timingIndex) {
  // タイミング列
  const timingCell = sheet.getRange(row, 1);
  timingCell.setValue(timing);
  timingCell.setBackground(ADVANCED_COLORS.TIMELINE_BG);
  timingCell.setFontWeight("bold");
  timingCell.setWrap(true);
  
  // 各部署のタスク
  flowData.departmentList.forEach((dept, deptIndex) => {
    const col = deptIndex + 2;
    const tasks = flowData.departments[dept][timing];
    
    if (tasks && tasks.length > 0) {
      const task = tasks[0];
      const cell = sheet.getRange(row, col);
      
      // タスク内容の設定
      let content = task.task;
      if (task.role && task.role !== "-") {
        content = "【" + task.role + "】\n" + content;
      }
      if (task.tool && task.tool !== "-") {
        const toolInfo = TOOL_ICONS[task.tool.split(/[／、,]/)[0].trim()];
        if (toolInfo) {
          content += "\n" + toolInfo.icon + " " + task.tool;
        } else {
          content += "\n🔧 " + task.tool;
        }
      }
      
      cell.setValue(content);
      cell.setWrap(true);
      cell.setHorizontalAlignment("center");
      cell.setVerticalAlignment("middle");
      
      // スタイルの設定
      if (task.condition && task.condition !== "-") {
        // 判断ボックス
        cell.setBackground(ADVANCED_COLORS.DECISION);
        cell.setBorder(true, true, true, true, false, false, "#FF8C00", SpreadsheetApp.BorderStyle.SOLID_THICK);
        cell.setFontWeight("bold");
        
        const noteContent = "⚡ 条件分岐: " + task.condition + 
                          (task.note ? "\n📝 備考: " + task.note : "");
        cell.setNote(noteContent);
      } else {
        // プロセスボックス
        cell.setBackground(ADVANCED_COLORS.PROCESS);
        cell.setBorder(true, true, true, true, false, false, "#4682B4", SpreadsheetApp.BorderStyle.SOLID_THICK);
        
        if (task.note) {
          cell.setNote("📝 備考: " + task.note);
        }
      }
      
      // 矢印の追加
      if (timingIndex < flowData.timings.length - 1) {
        const nextTiming = flowData.timings[timingIndex + 1];
        if (flowData.departments[dept][nextTiming]) {
          addAdvancedArrowToCell(cell, "↓");
        }
      }
    } else {
      // 空のセル
      const cell = sheet.getRange(row, col);
      cell.setBackground(ADVANCED_COLORS.EMPTY_BG);
    }
  });
  
  // データソース
  const hasDataSource = Object.keys(flowData.datasources).length > 0;
  if (hasDataSource) {
    const dataCol = flowData.departmentList.length + 2;
    const dataCell = sheet.getRange(row, dataCol);
    
    if (flowData.datasources[timing] && flowData.datasources[timing].length > 0) {
      const urls = flowData.datasources[timing].join("\n");
      dataCell.setValue("📎 " + urls);
      dataCell.setBackground(ADVANCED_COLORS.DATASOURCE_BG);
      dataCell.setBorder(true, true, true, true, false, false, "#2196F3", SpreadsheetApp.BorderStyle.DASHED);
      dataCell.setWrap(true);
      dataCell.setHorizontalAlignment("center");
      dataCell.setVerticalAlignment("middle");
    } else {
      dataCell.setValue("");
      dataCell.setBackground(ADVANCED_COLORS.EMPTY_BG);
    }
  }
  
  sheet.setRowHeight(row, 90);
}

// 終了行の描画（高度版）
function drawAdvancedEndRow(sheet, row, deptCount, hasDataSource) {
  const mergeCols = hasDataSource ? deptCount + 2 : deptCount + 1;
  sheet.getRange(row, 1, 1, mergeCols).merge();
  const cell = sheet.getRange(row, 1);
  cell.setValue("✅ 【プロセス完了】");
  cell.setBackground(ADVANCED_COLORS.START_END);
  cell.setFontColor("#FFFFFF");
  cell.setFontSize(14);
  cell.setFontWeight("bold");
  cell.setHorizontalAlignment("center");
  cell.setBorder(true, true, true, true, false, false, "#228B22", SpreadsheetApp.BorderStyle.SOLID_THICK);
  sheet.setRowHeight(row, 50);
}

// 凡例行の描画（高度版）
function drawAdvancedLegendRow(sheet, row, maxCols) {
  sheet.getRange(row, 1, 1, maxCols).merge();
  const legendCell = sheet.getRange(row, 1);
  legendCell.setValue("【凡例】 📦 処理・作業　⚡ 判断・分岐　→ 処理の流れ　📎 関連資料　📝 備考（セルの注記に詳細）");
  legendCell.setBackground(ADVANCED_COLORS.TIMELINE_BG);
  legendCell.setFontWeight("bold");
  legendCell.setHorizontalAlignment("left");
  legendCell.setVerticalAlignment("middle");
  sheet.setRowHeight(row, 40);
}

// セルに矢印を追加（高度版）
function addAdvancedArrowToCell(cell, arrow) {
  const currentValue = cell.getValue();
  const richText = SpreadsheetApp.newRichTextValue()
    .setText(currentValue + "\n" + arrow)
    .setTextStyle(currentValue.length + 1, currentValue.length + arrow.length + 1, 
      SpreadsheetApp.newTextStyle()
        .setForegroundColor(ADVANCED_COLORS.INFO)
        .setFontSize(16)
        .setBold(true)
        .build())
    .build();
  cell.setRichTextValue(richText);
}

// 列幅の調整（高度版）
function adjustAdvancedColumnWidths(sheet, deptCount, hasDataSource) {
  sheet.setColumnWidth(1, 150); // タイムライン列
  for (let i = 2; i <= deptCount + 1; i++) {
    sheet.setColumnWidth(i, 200); // 部署列
  }
  if (hasDataSource) {
    sheet.setColumnWidth(deptCount + 2, 150); // データソース列
  }
}

// 罫線の設定（高度版）
function applyAdvancedBorders(sheet, lastRow, lastCol) {
  const range = sheet.getRange(1, 1, lastRow, lastCol);
  range.setBorder(true, true, true, true, true, true, "#d0d0d0", SpreadsheetApp.BorderStyle.SOLID);
}

// 業務サマリーシート作成
function createBusinessSummarySheet(flowData) {
  const summarySheet = getOrCreateSheet('業務サマリー');
  summarySheet.clear();
  
  let row = 1;
  
  // タイトル
  summarySheet.getRange(row, 1, 1, 6).merge();
  const titleCell = summarySheet.getRange(row, 1);
  titleCell.setValue('業務プロセスサマリー');
  titleCell.setBackground(ADVANCED_COLORS.MAIN_HEADER);
  titleCell.setFontColor('#FFFFFF');
  titleCell.setFontSize(18);
  titleCell.setFontWeight('bold');
  titleCell.setHorizontalAlignment('center');
  summarySheet.setRowHeight(row, 50);
  row += 2;
  
  // 基本情報
  summarySheet.getRange(row, 1).setValue('プロセス名');
  summarySheet.getRange(row, 2, 1, 5).merge();
  summarySheet.getRange(row, 2).setValue(flowData.processName || '未設定');
  summarySheet.getRange(row, 1).setBackground(ADVANCED_COLORS.TIMELINE_BG);
  summarySheet.getRange(row, 1).setFontWeight('bold');
  row++;
  
  summarySheet.getRange(row, 1).setValue('総タスク数');
  summarySheet.getRange(row, 2, 1, 5).merge();
  summarySheet.getRange(row, 2).setValue(flowData.statistics.totalTasks);
  summarySheet.getRange(row, 1).setBackground(ADVANCED_COLORS.TIMELINE_BG);
  summarySheet.getRange(row, 1).setFontWeight('bold');
  row++;
  
  summarySheet.getRange(row, 1).setValue('関連部署数');
  summarySheet.getRange(row, 2, 1, 5).merge();
  summarySheet.getRange(row, 2).setValue(flowData.statistics.totalDepartments);
  summarySheet.getRange(row, 1).setBackground(ADVANCED_COLORS.TIMELINE_BG);
  summarySheet.getRange(row, 1).setFontWeight('bold');
  row++;
  
  summarySheet.getRange(row, 1).setValue('使用ツール数');
  summarySheet.getRange(row, 2, 1, 5).merge();
  summarySheet.getRange(row, 2).setValue(flowData.statistics.totalTools);
  summarySheet.getRange(row, 1).setBackground(ADVANCED_COLORS.TIMELINE_BG);
  summarySheet.getRange(row, 1).setFontWeight('bold');
  row++;
  
  summarySheet.getRange(row, 1).setValue('判断ポイント数');
  summarySheet.getRange(row, 2, 1, 5).merge();
  summarySheet.getRange(row, 2).setValue(flowData.statistics.decisionPoints);
  summarySheet.getRange(row, 1).setBackground(ADVANCED_COLORS.TIMELINE_BG);
  summarySheet.getRange(row, 1).setFontWeight('bold');
  row += 2;
  
  // 部署別タスク
  summarySheet.getRange(row, 1, 1, 6).merge();
  summarySheet.getRange(row, 1).setValue('👥 部署別タスク分布');
  summarySheet.getRange(row, 1).setBackground(ADVANCED_COLORS.SUB_HEADER);
  summarySheet.getRange(row, 1).setFontColor('#FFFFFF');
  summarySheet.getRange(row, 1).setFontWeight('bold');
  row++;
  
  flowData.departmentList.forEach(dept => {
    let taskCount = 0;
    Object.values(flowData.departments[dept]).forEach(timingTasks => {
      taskCount += timingTasks.length;
    });
    
    summarySheet.getRange(row, 1).setValue(dept);
    summarySheet.getRange(row, 2, 1, 5).merge();
    summarySheet.getRange(row, 2).setValue(`${taskCount} タスク`);
    summarySheet.getRange(row, 1).setBackground(DEPT_COLOR_PALETTE[flowData.departmentList.indexOf(dept) % DEPT_COLOR_PALETTE.length]);
    summarySheet.getRange(row, 1).setFontWeight('bold');
    row++;
  });
  
  // 書式調整
  summarySheet.autoResizeColumns(1, 6);
}

// シート取得または作成
function getOrCreateSheet(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }
  
  return sheet;
}

// 業務サマリーのみ作成
function createBusinessSummaryOnly() {
  try {
    const flowSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(FLOW_SHEET);
    if (!flowSheet) {
      SpreadsheetApp.getUi().alert('フローシートが見つかりません');
      return;
    }
    
    const data = flowSheet.getDataRange().getValues();
    if (data.length < 2) {
      SpreadsheetApp.getUi().alert('データがありません');
      return;
    }
    
    const flowData = parseAdvancedFlowData(data);
    createBusinessSummarySheet(flowData);
    
    SpreadsheetApp.getUi().alert('業務サマリーを作成しました');
  } catch (error) {
    SpreadsheetApp.getUi().alert('エラー: ' + error.toString());
  }
}

// ================================================================================
// 3. gmail_inbound.gs - Gmail受信処理機能
// ================================================================================

// Gmail受信処理

// 新着メール処理（メイン関数）
function processNewEmails() {
  const query = getConfig('PROCESSING_QUERY') || 'label:inbox is:unread';
  logActivity('PROCESS_START', `Processing emails with query: ${query}`);
  
  try {
    const threads = GmailApp.search(query);
    
    if (threads.length === 0) {
      logActivity('PROCESS_INFO', 'No new emails found');
      return;
    }
    
    threads.forEach(thread => {
      processThread(thread);
    });
    
    logActivity('PROCESS_END', `Processed ${threads.length} threads`);
  } catch (e) {
    logActivity('PROCESS_ERROR', e.toString());
    throw e;
  }
}

// スレッド処理
function processThread(thread) {
  const messages = thread.getMessages();
  
  messages.forEach(msg => {
    try {
      processMessage(msg, thread);
    } catch (e) {
      logActivity('MESSAGE_ERROR', `Failed to process message: ${e.toString()}`);
    }
  });
}

// メッセージ処理
function processMessage(msg, thread) {
  const messageId = msg.getId();
  
  // 処理済みチェック
  if (isProcessed(messageId)) {
    logActivity('SKIP', `Message ${messageId} already processed`);
    return;
  }
  
  // メール情報抽出
  const from = extractEmail(msg.getFrom());
  const subject = msg.getSubject();
  const htmlBody = msg.getBody();
  let plainBody = msg.getPlainBody() || htmlToText(htmlBody);
  const receivedDate = msg.getDate();
  
  // 署名部分を除去
  plainBody = removeEmailSignature(plainBody);
  
  // 件名から[WORK-REQ]を除去して、本文と結合
  const cleanSubject = subject.replace(/\[WORK-REQ\]/gi, '').trim();
  const combinedContent = `【件名】${cleanSubject}\n\n【本文】\n${plainBody}`;
  
  // Inboxにログ記録
  logInbox(messageId, thread.getId(), from, subject, plainBody.substring(0, 200), 'NEW');
  
  try {
    // OpenAI呼び出し（件名と本文を結合したものを送信）
    const orgProfile = getConfig('ORG_PROFILE_JSON') || '{}';
    const result = callOpenAI(combinedContent, orgProfile);
    
    // 検証
    validateOpenAIResponse(result);
    
    // データ書き込み（改善されたエンジンを使用）
    writeWorkSpec(result.work_spec);
    
    // 新しいデータ処理エンジンを強制使用
    writeFlowRowsImproved(result.flow_rows);
    
    // ビジュアルフロー生成
    if (typeof generateVisualFlow === 'function') {
      generateVisualFlow();
    }
    
    // 共有設定
    const shareSuccess = handleSharing(from);
    
    // 返信メール送信
    sendNotificationEmail(from, result.work_spec, ss().getUrl());
    
    // 処理済みマーク
    markProcessed(messageId);
    labelThreadProcessed(thread);
    
    logActivity('PROCESS_SUCCESS', `Successfully processed message ${messageId}`);
  } catch (e) {
    logError(messageId, e);
    
    // エラー通知メール送信
    sendErrorNotificationEmail(from, subject, e.toString());
    
    throw e;
  }
}

// 共有設定処理
function handleSharing(senderEmail) {
  let shareSuccess = false;
  
  // ANYONE_WITH_LINKの設定を試行
  if (String(getConfig('SHARE_ANYONE_WITH_LINK')).toUpperCase() === 'TRUE') {
    shareSuccess = shareSheetAnyWithLink();
  }
  
  // 送信者を編集者として追加
  const editorSuccess = addEditor(senderEmail);
  
  return shareSuccess && editorSuccess;
}

// 成功通知メール送信
function sendNotificationEmail(to, workSpec, sheetUrl) {
  const subject = `[WORK-SPEC READY] ${workSpec.title}`;
  const plainBody = buildPlainTextNotification(workSpec, sheetUrl);
  const htmlBody = buildHtmlNotification(workSpec, sheetUrl);
  
  GmailApp.sendEmail(to, subject, plainBody, {
    htmlBody: htmlBody,
    name: 'タスク管理システム'
  });
  
  logActivity('EMAIL_SENT', `Notification sent to ${to}`);
}

// プレーンテキスト通知作成
function buildPlainTextNotification(workSpec, sheetUrl) {
  return `業務記述書の作成が完了しました。

タイトル: ${workSpec.title}
概要: ${workSpec.summary}

スプレッドシートURL: ${sheetUrl}

このスプレッドシートでは以下の内容を確認・編集できます：
- 業務記述書（詳細仕様）
- タスクフロー表
- ビジュアルフロー図

【重要な注意事項】
- 本書面は自動生成されたものです。最終的な判断は専門家にご確認ください。
- 法令・規制に関する記載は参考情報であり、法的助言ではありません。
- スプレッドシートは編集可能です。必要に応じて内容を更新してください。

---
タスク管理システム by Google Apps Script`;
}

// HTML通知作成
function buildHtmlNotification(workSpec, sheetUrl) {
  return `
    <div style="font-family: 'Noto Sans JP', Arial, sans-serif; max-width: 600px; margin: 0 auto;">
      <div style="background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; padding: 20px; text-align: center; border-radius: 10px 10px 0 0;">
        <h1 style="margin: 0; font-size: 24px;">📋 業務記述書が完成しました</h1>
      </div>
      
      <div style="padding: 20px; background-color: #f8f9fa; border: 1px solid #e9ecef; border-top: none;">
        <h2 style="color: #495057; margin-top: 0;">${workSpec.title}</h2>
        <p style="font-size: 16px; line-height: 1.5; color: #6c757d;">${workSpec.summary}</p>
        
        <div style="background-color: white; padding: 15px; border-radius: 8px; margin: 20px 0; border-left: 4px solid #28a745;">
          <h3 style="margin-top: 0; color: #28a745;">📊 スプレッドシート（編集可能）</h3>
          <p style="margin-bottom: 10px;">以下のリンクから業務記述書とタスク管理シートにアクセス・編集できます：</p>
          <a href="${sheetUrl}" style="display: inline-block; background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; padding: 12px 24px; text-decoration: none; border-radius: 4px; font-weight: bold;">📝 スプレッドシートを開く</a>
        </div>
        
        ${workSpec.timeline && workSpec.timeline.length > 0 ? `
        <div style="background-color: white; padding: 15px; border-radius: 8px; margin: 20px 0; border-left: 4px solid #ffc107;">
          <h3 style="margin-top: 0; color: #ff9800;">⏰ 主要マイルストーン</h3>
          <ul style="margin: 0; padding-left: 20px;">
            ${workSpec.timeline.map(phase => `
              <li style="margin-bottom: 8px;">
                <strong>${phase.phase}</strong> (${phase.duration_hint})
                ${phase.milestones && phase.milestones.length > 0 ? 
                  `<ul style="margin-top: 5px;">${phase.milestones.map(milestone => 
                    `<li style="color: #6c757d;">${milestone}</li>`
                  ).join('')}</ul>` 
                  : ''}
              </li>
            `).join('')}
          </ul>
        </div>
        ` : ''}
        
        <div style="background-color: white; padding: 15px; border-radius: 8px; margin: 20px 0; border-left: 4px solid #dc3545;">
          <h3 style="margin-top: 0; color: #dc3545;">⚠️ 重要な注意事項</h3>
          <ul style="margin: 0; padding-left: 20px; color: #6c757d;">
            <li>本書面は自動生成されたものです。最終的な判断は専門家にご確認ください。</li>
            <li>法令・規制に関する記載は参考情報であり、法的助言ではありません。</li>
            <li>スプレッドシートは編集可能です。必要に応じて内容を更新してください。</li>
          </ul>
        </div>
        
        <div style="text-align: center; margin-top: 30px; padding-top: 20px; border-top: 1px solid #dee2e6;">
          <p style="color: #6c757d; font-size: 14px; margin: 0;">
            このメールは自動送信されています。<br>
            タスク管理システム by Google Apps Script
          </p>
        </div>
      </div>
    </div>
  `;
}

// エラー通知メール送信
function sendErrorNotificationEmail(to, originalSubject, errorMessage) {
  const subject = `[WORK-SPEC ERROR] 処理エラー: ${originalSubject}`;
  const body = `業務記述書の作成中にエラーが発生しました。

元の件名: ${originalSubject}
エラー内容: ${errorMessage}

お手数ですが、システム管理者にお問い合わせください。

---
タスク管理システム by Google Apps Script`;
  
  try {
    GmailApp.sendEmail(to, subject, body);
    logActivity('ERROR_EMAIL_SENT', `Error notification sent to ${to}`);
  } catch (e) {
    logActivity('ERROR_EMAIL_FAILED', `Failed to send error notification: ${e.toString()}`);
  }
}

// スレッドに処理済みラベルを付与
function labelThreadProcessed(thread) {
  try {
    // 既存のラベルを取得または作成
    let label = GmailApp.getUserLabelByName('PROCESSED');
    if (!label) {
      label = GmailApp.createLabel('PROCESSED');
    }
    
    thread.addLabel(label);
    thread.markRead();
    
    logActivity('LABEL', `Added PROCESSED label to thread ${thread.getId()}`);
  } catch (e) {
    logActivity('LABEL_ERROR', `Failed to label thread: ${e.toString()}`);
  }
}

// ================================================================================
// 4. gmail_outbound.gs - Gmail送信処理機能
// ================================================================================

// Gmail送信処理（任意の業務メール送信機能）

// サイドバーUI表示
function showEmailComposer() {
  const html = HtmlService.createHtmlOutput(getEmailComposerHtml())
    .setTitle('業務メール作成')
    .setWidth(400);
  SpreadsheetApp.getUi().showSidebar(html);
}

// メール送信（サイドバーから呼び出し）
function sendBusinessEmail(to, subject, body) {
  try {
    // 入力検証
    if (!to || !subject || !body) {
      throw new Error('宛先、件名、本文はすべて必須です。');
    }
    
    // メールアドレス検証
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    if (!emailRegex.test(to)) {
      throw new Error('有効なメールアドレスを入力してください。');
    }
    
    // デフォルトの件名プレフィックスを追加
    const prefixedSubject = subject.startsWith('[WORK-REQ]') ? subject : `[WORK-REQ] ${subject}`;
    
    // HTML形式のメール本文作成
    const htmlBody = createBusinessEmailHtml(body);
    
    // メール送信
    GmailApp.sendEmail(to, prefixedSubject, body, {
      htmlBody: htmlBody,
      name: 'タスク管理システム'
    });
    
    // ログ記録
    logActivity('OUTBOUND_EMAIL', `Sent to: ${to}, Subject: ${prefixedSubject}`);
    
    return {
      success: true,
      message: 'メールを送信しました。'
    };
    
  } catch (e) {
    logActivity('OUTBOUND_ERROR', e.toString());
    return {
      success: false,
      message: `エラー: ${e.toString()}`
    };
  }
}

// ビジネスメールHTML作成
function createBusinessEmailHtml(body) {
  // 改行をHTMLのbrタグに変換
  const htmlBody = body.replace(/\n/g, '<br>');
  
  return `
    <div style="font-family: 'Noto Sans JP', Arial, sans-serif; max-width: 600px; margin: 0 auto;">
      <div style="background-color: #f8f9fa; padding: 20px; border: 1px solid #dee2e6; border-radius: 8px;">
        <div style="background-color: white; padding: 20px; border-radius: 4px;">
          <p style="font-size: 16px; line-height: 1.6; color: #333; margin: 0;">
            ${htmlBody}
          </p>
        </div>
        
        <div style="margin-top: 20px; padding-top: 20px; border-top: 1px solid #dee2e6;">
          <p style="color: #6c757d; font-size: 14px; margin: 0;">
            このメールはタスク管理システムから送信されています。<br>
            業務内容に基づいて自動的に業務記述書とタスクフローが生成されます。
          </p>
        </div>
      </div>
    </div>
  `;
}

// サイドバーHTML取得
function getEmailComposerHtml() {
  const defaultTo = getConfig('DEFAULT_TO_EMAIL') || '';
  
  return `
    <!DOCTYPE html>
    <html>
      <head>
        <base target="_top">
        <style>
          body {
            font-family: 'Noto Sans JP', Arial, sans-serif;
            padding: 15px;
            margin: 0;
          }
          .form-group {
            margin-bottom: 15px;
          }
          label {
            display: block;
            margin-bottom: 5px;
            font-weight: bold;
            color: #333;
          }
          input, textarea {
            width: 100%;
            padding: 8px;
            border: 1px solid #ddd;
            border-radius: 4px;
            font-size: 14px;
            box-sizing: border-box;
          }
          textarea {
            resize: vertical;
            min-height: 150px;
          }
          button {
            width: 100%;
            padding: 10px;
            border: none;
            border-radius: 4px;
            font-size: 16px;
            font-weight: bold;
            cursor: pointer;
            transition: background-color 0.3s;
          }
          .btn-primary {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            margin-bottom: 10px;
          }
          .btn-primary:hover {
            opacity: 0.9;
          }
          .btn-secondary {
            background-color: #6c757d;
            color: white;
          }
          .btn-secondary:hover {
            background-color: #5a6268;
          }
          .loading {
            display: none;
            text-align: center;
            padding: 20px;
          }
          .spinner {
            border: 3px solid #f3f3f3;
            border-top: 3px solid #667eea;
            border-radius: 50%;
            width: 40px;
            height: 40px;
            animation: spin 1s linear infinite;
            margin: 0 auto;
          }
          @keyframes spin {
            0% { transform: rotate(0deg); }
            100% { transform: rotate(360deg); }
          }
          .message {
            padding: 10px;
            border-radius: 4px;
            margin-bottom: 15px;
            display: none;
          }
          .message.success {
            background-color: #d4edda;
            color: #155724;
            border: 1px solid #c3e6cb;
          }
          .message.error {
            background-color: #f8d7da;
            color: #721c24;
            border: 1px solid #f5c6cb;
          }
          .info {
            background-color: #e3f2fd;
            border: 1px solid #90caf9;
            border-radius: 4px;
            padding: 10px;
            margin-bottom: 15px;
            color: #1565c0;
            font-size: 13px;
          }
        </style>
      </head>
      <body>
        <h2 style="color: #333; margin-top: 0;">業務メール作成</h2>
        
        <div class="info">
          ℹ️ このフォームから送信されたメールは自動的に処理され、業務記述書とタスクフローが生成されます。
        </div>
        
        <div id="message" class="message"></div>
        
        <form id="emailForm">
          <div class="form-group">
            <label for="to">宛先メールアドレス *</label>
            <input type="email" id="to" name="to" value="${defaultTo}" required placeholder="example@example.com">
          </div>
          
          <div class="form-group">
            <label for="subject">件名 *</label>
            <input type="text" id="subject" name="subject" required placeholder="業務依頼のタイトル">
            <small style="color: #666; font-size: 12px;">※ [WORK-REQ] プレフィックスは自動付与されます</small>
          </div>
          
          <div class="form-group">
            <label for="body">業務内容 *</label>
            <textarea id="body" name="body" required placeholder="実施したい業務の詳細を記入してください。&#10;&#10;例：&#10;- 業務の目的&#10;- 必要な成果物&#10;- 期限&#10;- 関係者&#10;- その他要件"></textarea>
          </div>
          
          <button type="submit" class="btn-primary">送信</button>
          <button type="button" class="btn-secondary" onclick="clearForm()">クリア</button>
        </form>
        
        <div id="loading" class="loading">
          <div class="spinner"></div>
          <p>送信中...</p>
        </div>
        
        <script>
          document.getElementById('emailForm').addEventListener('submit', function(e) {
            e.preventDefault();
            
            const to = document.getElementById('to').value;
            const subject = document.getElementById('subject').value;
            const body = document.getElementById('body').value;
            
            // フォームを非表示、ローディング表示
            document.getElementById('emailForm').style.display = 'none';
            document.getElementById('loading').style.display = 'block';
            document.getElementById('message').style.display = 'none';
            
            // GASの関数を呼び出し
            google.script.run
              .withSuccessHandler(function(result) {
                document.getElementById('loading').style.display = 'none';
                document.getElementById('emailForm').style.display = 'block';
                
                const messageDiv = document.getElementById('message');
                messageDiv.className = result.success ? 'message success' : 'message error';
                messageDiv.textContent = result.message;
                messageDiv.style.display = 'block';
                
                if (result.success) {
                  // 成功時はフォームをクリア
                  clearForm();
                  
                  // 3秒後にメッセージを非表示
                  setTimeout(function() {
                    messageDiv.style.display = 'none';
                  }, 3000);
                }
              })
              .withFailureHandler(function(error) {
                document.getElementById('loading').style.display = 'none';
                document.getElementById('emailForm').style.display = 'block';
                
                const messageDiv = document.getElementById('message');
                messageDiv.className = 'message error';
                messageDiv.textContent = 'エラーが発生しました: ' + error.toString();
                messageDiv.style.display = 'block';
              })
              .sendBusinessEmail(to, subject, body);
          });
          
          function clearForm() {
            document.getElementById('subject').value = '';
            document.getElementById('body').value = '';
            // 宛先はデフォルト値があれば保持
          }
        </script>
      </body>
    </html>
  `;
}

// ================================================================================
// 5. menu.gs - カスタムメニューとセットアップ機能
// ================================================================================

// カスタムメニューとセットアップ機能

// スプレッドシート開いた時の処理
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  ui.createMenu('📋 タスク管理システム')
    .addSubMenu(ui.createMenu('⚙️ システム')
      .addItem('🚀 初回セットアップ', 'setupSystem')
      .addItem('🔧 設定を開く', 'openConfigSheet')
      .addSeparator()
      .addItem('🔑 APIキーを設定', 'setApiKey')
      .addItem('⏰ トリガーを設定', 'setupTriggers')
      .addItem('🗑️ トリガーを削除', 'deleteTriggers'))
    .addSubMenu(ui.createMenu('📧 メール')
      .addItem('✉️ 業務メール作成', 'showEmailComposer')
      .addItem('📥 新着メール処理を今すぐ実行', 'processNewEmailsManually')
      .addSeparator()
      .addItem('🏷️ 処理済みラベルを作成', 'createProcessedLabel'))
    .addSubMenu(ui.createMenu('📊 フロー')
      .addItem('🎨 ビジュアルフロー生成', 'generateVisualFlow')
      .addItem('✨ 高度なビジュアルフロー生成', 'generateAdvancedVisualFlow')
      .addItem('📋 業務サマリー作成', 'createBusinessSummaryOnly')
      .addItem('📝 サンプルデータ作成', 'createSampleFlowData')
      .addSeparator()
      .addItem('🔄 フローシートをリセット', 'resetFlowSheet'))
    .addSubMenu(ui.createMenu('📈 レポート')
      .addItem('📊 処理統計を表示', 'showProcessingStats')
      .addItem('📋 アクティビティログを表示', 'showActivityLog'))
    .addSeparator()
    .addItem('❓ ヘルプ', 'showHelp')
    .addItem('ℹ️ バージョン情報', 'showAbout')
    .addToUi();
    
  // 初回起動チェック
  checkFirstRun();
}

// 初回セットアップ
function setupSystem() {
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.alert(
    '初回セットアップ',
    '以下の処理を実行します：\n\n' +
    '1. 必要なシートの作成\n' +
    '2. 初期設定の配置\n' +
    '3. APIキーの設定確認\n' +
    '4. タイマートリガーの設定\n\n' +
    '続行しますか？',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) {
    return;
  }
  
  try {
    // プログレス表示
    const progressHtml = HtmlService.createHtmlOutput(getProgressHtml())
      .setWidth(400)
      .setHeight(200);
    ui.showModalDialog(progressHtml, 'セットアップ中...');
    
    // 1. シート作成
    createRequiredSheets();
    
    // 2. 初期設定
    initializeConfig();
    
    // 3. APIキー確認
    const apiKey = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
    if (!apiKey) {
      ui.alert(
        '⚠️ APIキー未設定',
        'OpenAI APIキーが設定されていません。\n' +
        'メニューから「システム > APIキーを設定」を選択して設定してください。',
        ui.ButtonSet.OK
      );
    }
    
    // 4. トリガー設定
    setupTriggers();
    
    // 完了メッセージ
    ui.alert(
      '✅ セットアップ完了',
      'システムのセットアップが完了しました。\n\n' +
      '次のステップ：\n' +
      '1. Config シートで設定を確認\n' +
      '2. APIキーを設定（未設定の場合）\n' +
      '3. テストメールを送信して動作確認',
      ui.ButtonSet.OK
    );
    
  } catch (e) {
    ui.alert(
      '❌ セットアップエラー',
      'セットアップ中にエラーが発生しました：\n' + e.toString(),
      ui.ButtonSet.OK
    );
    logActivity('SETUP_ERROR', e.toString());
  }
}

// 必要なシートの作成
function createRequiredSheets() {
  const requiredSheets = [
    CONFIG_SHEET,
    INBOX_SHEET,
    SPEC_SHEET,
    FLOW_SHEET,
    VISUAL_SHEET,
    ACTIVITY_LOG_SHEET
  ];
  
  requiredSheets.forEach(sheetName => {
    if (!ss().getSheetByName(sheetName)) {
      if (sheetName === CONFIG_SHEET) {
        initializeConfig();
      } else if (sheetName === INBOX_SHEET) {
        createInboxSheet();
      } else if (sheetName === SPEC_SHEET) {
        createWorkSpecSheet();
      } else if (sheetName === FLOW_SHEET) {
        createFlowSheet(sheetName);
      } else if (sheetName === ACTIVITY_LOG_SHEET) {
        createActivityLogSheet();
      } else {
        ss().insertSheet(sheetName);
      }
    }
  });
  
  logActivity('SETUP', 'Required sheets created');
}

// APIキー設定
function setApiKey() {
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.prompt(
    'OpenAI APIキー設定',
    'OpenAI APIキーを入力してください：\n' +
    '（キーは安全に保存されます）',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() === ui.Button.OK) {
    const apiKey = response.getResponseText().trim();
    
    if (apiKey) {
      PropertiesService.getScriptProperties().setProperty('OPENAI_API_KEY', apiKey);
      ui.alert('✅ APIキーを設定しました。');
      logActivity('API_KEY', 'API key configured');
    } else {
      ui.alert('⚠️ APIキーが入力されていません。');
    }
  }
}

// トリガー設定
function setupTriggers() {
  // 既存のトリガーを削除
  deleteTriggers();
  
  // 時間ベーストリガーを作成（5分ごと）
  ScriptApp.newTrigger('processNewEmails')
    .timeBased()
    .everyMinutes(5)
    .create();
    
  logActivity('TRIGGER', 'Time-based trigger created (every 5 minutes)');
}

// トリガー削除
function deleteTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'processNewEmails') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  logActivity('TRIGGER', 'Existing triggers deleted');
}

// 手動でメール処理実行
function processNewEmailsManually() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    ui.alert('📥 処理中...', 'メールを処理しています。しばらくお待ちください。', ui.ButtonSet.OK);
    processNewEmails();
    ui.alert('✅ 完了', 'メール処理が完了しました。', ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('❌ エラー', 'メール処理中にエラーが発生しました：\n' + e.toString(), ui.ButtonSet.OK);
  }
}

// Config シートを開く
function openConfigSheet() {
  const sheet = ss().getSheetByName(CONFIG_SHEET);
  if (sheet) {
    ss().setActiveSheet(sheet);
  } else {
    SpreadsheetApp.getUi().alert('Config シートが見つかりません。');
  }
}

// 処理済みラベル作成
function createProcessedLabel() {
  try {
    let label = GmailApp.getUserLabelByName('PROCESSED');
    if (!label) {
      label = GmailApp.createLabel('PROCESSED');
      SpreadsheetApp.getUi().alert('✅ PROCESSEDラベルを作成しました。');
    } else {
      SpreadsheetApp.getUi().alert('ℹ️ PROCESSEDラベルは既に存在します。');
    }
  } catch (e) {
    SpreadsheetApp.getUi().alert('❌ エラー：' + e.toString());
  }
}

// フローシートリセット
function resetFlowSheet() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    '確認',
    'フローシートのデータをすべて削除しますか？',
    ui.ButtonSet.YES_NO
  );
  
  if (response === ui.Button.YES) {
    const sheet = ss().getSheetByName(FLOW_SHEET);
    if (sheet) {
      sheet.clear();
      const headers = FLOW_HEADERS;
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#e8f5e9');
      ui.alert('✅ フローシートをリセットしました。');
    }
  }
}

// 処理統計表示
function showProcessingStats() {
  const inboxSheet = ss().getSheetByName(INBOX_SHEET);
  if (!inboxSheet || inboxSheet.getLastRow() <= 1) {
    SpreadsheetApp.getUi().alert('処理データがありません。');
    return;
  }
  
  const lastRow = inboxSheet.getLastRow();
  const dataRows = Math.max(1, lastRow - 1);
  const data = inboxSheet.getRange(2, 7, dataRows, 1).getValues();
  const stats = {
    total: data.length,
    processed: data.filter(row => row[0] === 'PROCESSED').length,
    error: data.filter(row => row[0] === 'ERROR').length,
    new: data.filter(row => row[0] === 'NEW').length
  };
  
  const message = `📊 処理統計\n\n` +
    `合計: ${stats.total} 件\n` +
    `処理済み: ${stats.processed} 件\n` +
    `エラー: ${stats.error} 件\n` +
    `未処理: ${stats.new} 件`;
    
  SpreadsheetApp.getUi().alert(message);
}

// アクティビティログ表示
function showActivityLog() {
  const sheet = ss().getSheetByName(ACTIVITY_LOG_SHEET);
  if (sheet) {
    sheet.showSheet();
    ss().setActiveSheet(sheet);
  } else {
    SpreadsheetApp.getUi().alert('アクティビティログがありません。');
  }
}

// ヘルプ表示
function showHelp() {
  const helpText = `📋 タスク管理システム - ヘルプ\n\n` +
    `【基本的な使い方】\n` +
    `1. 初回セットアップを実行\n` +
    `2. OpenAI APIキーを設定\n` +
    `3. Config シートで設定を調整\n` +
    `4. メール受信または送信で業務記述書を自動生成\n\n` +
    `【メール処理】\n` +
    `- 件名に [WORK-REQ] を含むメールを自動処理\n` +
    `- 5分ごとに自動チェック（変更可能）\n` +
    `- 処理結果は送信者にメール通知\n\n` +
    `【トラブルシューティング】\n` +
    `- エラーが発生した場合はInboxシートを確認\n` +
    `- APIキーが正しく設定されているか確認\n` +
    `- アクティビティログで詳細を確認`;
    
  SpreadsheetApp.getUi().alert(helpText);
}

// バージョン情報表示
function showAbout() {
  const about = `📋 タスク管理システム\n\n` +
    `バージョン: 1.0.0\n` +
    `作成日: 2024\n` +
    `説明: メールから業務記述書とタスクフローを自動生成\n\n` +
    `機能:\n` +
    `- OpenAI GPTによる業務記述書生成\n` +
    `- ビジュアルフロー自動描画\n` +
    `- Gmail連携による自動処理\n` +
    `- 上場企業レベルの品質管理`;
    
  SpreadsheetApp.getUi().alert(about);
}

// 初回起動チェック
function checkFirstRun() {
  const isFirstRun = PropertiesService.getDocumentProperties().getProperty('FIRST_RUN_COMPLETE');
  
  if (!isFirstRun) {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      '👋 ようこそ！',
      'タスク管理システムへようこそ！\n\n' +
      '初回セットアップを実行しますか？',
      ui.ButtonSet.YES_NO
    );
    
    if (response === ui.Button.YES) {
      setupSystem();
    }
    
    PropertiesService.getDocumentProperties().setProperty('FIRST_RUN_COMPLETE', 'true');
  }
}

// プログレス表示用HTML
function getProgressHtml() {
  return `
    <!DOCTYPE html>
    <html>
      <head>
        <style>
          body {
            font-family: Arial, sans-serif;
            text-align: center;
            padding: 20px;
          }
          .spinner {
            border: 4px solid #f3f3f3;
            border-top: 4px solid #3498db;
            border-radius: 50%;
            width: 50px;
            height: 50px;
            animation: spin 1s linear infinite;
            margin: 20px auto;
          }
          @keyframes spin {
            0% { transform: rotate(0deg); }
            100% { transform: rotate(360deg); }
          }
        </style>
      </head>
      <body>
        <h3>セットアップ中...</h3>
        <div class="spinner"></div>
        <p>しばらくお待ちください</p>
      </body>
    </html>
  `;
}

// ================================================================================
// 6. openai_client.gs - OpenAI API連携機能
// ================================================================================

// OpenAI API設定
const OPENAI_URL_CHAT = 'https://api.openai.com/v1/chat/completions';

// JSON Schema定義
function buildWorkSpecSchema() {
  return {
    type: 'object',
    additionalProperties: false,
    properties: {
      work_spec: {
        type: 'object',
        additionalProperties: false,
        properties: {
          title: { type: 'string' },
          summary: { type: 'string' },
          scope: { type: 'string' },
          deliverables: { type: 'array', items: { type: 'string' } },
          org_structure: { type: 'array', items: { type: 'string' } },
          raci: {
            type: 'array',
            items: {
              type: 'object',
              additionalProperties: false,
              properties: {
                role: { type: 'string' },
                dept: { type: 'string' },
                R: { type: 'boolean' },
                A: { type: 'boolean' },
                C: { type: 'boolean' },
                I: { type: 'boolean' }
              },
              required: ['role', 'dept', 'R', 'A', 'C', 'I']
            }
          },
          timeline: {
            type: 'array',
            items: {
              type: 'object',
              additionalProperties: false,
              properties: {
                phase: { type: 'string' },
                duration_hint: { type: 'string' },
                milestones: { type: 'array', items: { type: 'string' } },
                dependencies: { type: 'array', items: { type: 'string' } }
              },
              required: ['phase', 'duration_hint', 'milestones', 'dependencies']
            }
          },
          requirements_constraints: { type: 'array', items: { type: 'string' } },
          risks_mitigations: { type: 'array', items: { type: 'string' } },
          pro_considerations: { type: 'array', items: { type: 'string' } },
          kpi_sla: { type: 'array', items: { type: 'string' } },
          approvals: { type: 'array', items: { type: 'string' } },
          security_privacy_controls: { type: 'array', items: { type: 'string' } },
          legal_regulations: {
            type: 'array',
            items: {
              type: 'object',
              additionalProperties: false,
              properties: {
                name: { type: 'string' },
                scope: { type: 'string' },
                note: { type: 'string' }
              },
              required: ['name', 'scope', 'note']
            }
          },
          references: { type: 'array', items: { type: 'string' } },
          assumptions: { type: 'array', items: { type: 'string' } }
        },
        required: [
          'title', 'summary', 'scope', 'deliverables', 'org_structure',
          'raci', 'timeline', 'requirements_constraints', 'risks_mitigations',
          'pro_considerations', 'kpi_sla', 'approvals', 'security_privacy_controls',
          'legal_regulations', 'references', 'assumptions'
        ]
      },
      flow_rows: {
        type: 'array',
        items: {
          type: 'object',
          additionalProperties: false,
          properties: {
            '工程': { type: 'string' },
            '実施タイミング': { type: 'string' },
            '部署': { type: 'string' },
            '担当役割': { type: 'string' },
            '作業内容': { type: 'string' },
            '条件分岐': { type: 'string' },
            '利用ツール': { type: 'string' },
            'URLリンク': { type: 'string' },
            '備考': { type: 'string' }
          },
          required: ['工程', '実施タイミング', '部署', '担当役割', '作業内容', '条件分岐', '利用ツール', 'URLリンク', '備考']
        }
      }
    },
    required: ['work_spec', 'flow_rows']
  };
}

// システムプロンプト構築
function buildSystemPrompt() {
  return `あなたは上場企業レベルのプロジェクトマネジメント、法務、内部統制の実務に精通したアシスタントです。

重要: 与えられた情報が少ない場合でも、あなたの専門知識を活用して、実務的で具体的な提案を積極的に行ってください。

業務内容の記述から、以下の観点で業務記述書とフロー表を作成してください：

1. 上場企業のプロ水準の品質
   - 内部統制（J-SOX）への配慮
   - 説明責任と透明性の確保
   - リスク管理と対策の明確化

2. 法令・規制への対応
   - 該当する法令・ガイドラインを名称レベルで列挙
   - 業界特有の規制やルールも考慮
   - 必ず「最終的には法務・専門家の確認が必要」と明記

3. 詳細な体制とスケジュール
   - RACI（推奨）による役割分担の明確化
   - フェーズ分割と各フェーズの所要期間目安
   - マイルストーンと依存関係の明示

4. セキュリティと個人情報保護
   - データ分類とアクセス権限
   - 個人情報保護法への準拠
   - セキュリティ管理策の具体化

5. 成果物と受入基準
   - 各成果物の明確な定義
   - 完了基準（DoD）の設定
   - KPI/SLAの設定

すべての出力は日本語で作成し、指定されたJSON Schemaに完全準拠してください。
法的助言の代替ではないことを必ず明記してください。`;
}

// ユーザープロンプト構築
function buildUserPrompt(mailBody, orgProfileJson) {
  const orgProfile = orgProfileJson ? JSON.parse(orgProfileJson) : {};
  
  return `以下の業務内容から、業務記述書とフロー表を作成してください。

【業務内容】
${mailBody}

【組織プロフィール】
- 上場区分: ${orgProfile.listing || '未設定'}
- 業種: ${orgProfile.industry || '未設定'}
- 対象地域: ${(orgProfile.jurisdictions || ['JP']).join(', ')}
- 社内基準: ${(orgProfile.policies || []).join(', ')}

【要求事項】
1. 上場企業として必要な観点をすべて網羅
2. 法令・規制は具体的な名称を記載（最終確認は専門家が必要な旨も明記）
3. フロー表は実行可能な詳細レベルで作成
4. リスクと対策を具体的に記載
5. セキュリティ・個人情報保護・内部統制の観点を含める

【重要な指示】
- 情報が不足している場合は、業界標準やベストプラクティスに基づいて、具体的な提案を行ってください
- 「上場準備」のような抽象的な依頼でも、IPO準備の標準的なプロセスを想定して詳細な計画を作成してください
- すべてのフィールドに必ず値を設定してください（空配列の場合も最低1つの要素を含める）
- scope: 業務の範囲を明確に記載
- deliverables: 成果物リストを必ず含める（最低1つ）
- org_structure: 組織体制を必ず含める（最低1つ）

【flow_rowsの必須項目】
- 工程: 各フェーズ名を必ず記載（例：「事前準備」「内部統制構築」「監査法人選定」など）
- 実施タイミング: 時期を必ず記載（例：「N-3期」「第1四半期」「1月～3月」など）
- 部署: 担当部署を必ず記載（例：「経営企画部」「経理部」「法務部」など）
- 担当役割: 役割を必ず記載（例：「CFO」「プロジェクトマネージャー」「担当者」など）
- 作業内容: 具体的なタスクを必ず記載
- 条件分岐: 分岐がない場合は「なし」と記載
- 利用ツール: ツールが不要な場合は「手動作業」と記載
- URLリンク: 参考URLがない場合は「なし」と記載
- 備考: 特記事項がない場合は「特になし」と記載

重要：空文字列やnullを返さず、必ず意味のある値を設定してください。`;
}

// OpenAI API呼び出し
function callOpenAI(mailBody, orgProfileJson) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
  if (!apiKey) {
    throw new Error('OpenAI APIキーが設定されていません。スクリプトプロパティに OPENAI_API_KEY を設定してください。');
  }
  
  const schema = buildWorkSpecSchema();
  const messages = [
    { role: 'system', content: buildSystemPrompt() },
    { role: 'user', content: buildUserPrompt(mailBody, orgProfileJson) }
  ];
  
  const payload = {
    model: getConfig('OPENAI_MODEL') || 'gpt-4o',
    messages: messages,
    response_format: {
      type: 'json_schema',
      json_schema: {
        name: 'WorkSpecSchema',
        schema: schema,
        strict: true  // strictモードを有効化してデータ品質を向上
      }
    },
    temperature: 0.1,  // より一貫した出力のために温度を下げる
    max_tokens: 6000,  // より詳細な出力のためにトークン数を増加
    seed: 42  // 再現性のためのシード値
  };
  
  logActivity('OPENAI_CALL', `Calling OpenAI with model: ${payload.model}`);
  
  const response = retryWithBackoff(() => {
    const res = UrlFetchApp.fetch(OPENAI_URL_CHAT, {
      method: 'post',
      headers: {
        'Authorization': `Bearer ${apiKey}`,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    
    const status = res.getResponseCode();
    if (status >= 300) {
      const errorBody = res.getContentText();
      logActivity('OPENAI_ERROR', `Status: ${status}, Body: ${errorBody}`);
      throw new Error(`OpenAI API error ${status}: ${errorBody}`);
    }
    
    return res;
  });
  
  const responseData = JSON.parse(response.getContentText());
  const content = responseData.choices[0].message.content;
  
  logActivity('OPENAI_SUCCESS', 'Successfully received response from OpenAI');
  
  return JSON.parse(content);
}

// JSON検証
function validateOpenAIResponse(data) {
  if (!data || typeof data !== 'object') {
    throw new Error('Invalid response format');
  }
  
  if (!data.work_spec || typeof data.work_spec !== 'object') {
    throw new Error('Missing or invalid work_spec');
  }
  
  if (!data.flow_rows || !Array.isArray(data.flow_rows)) {
    throw new Error('Missing or invalid flow_rows');
  }
  
  const ws = data.work_spec;
  const required = ['title', 'summary'];
  for (const field of required) {
    if (!ws[field]) {
      throw new Error(`Missing required field in work_spec: ${field}`);
    }
  }
  
  // flow_rowsの検証をより柔軟に
  for (let i = 0; i < data.flow_rows.length; i++) {
    const row = data.flow_rows[i];
    const requiredFlow = ['工程', '実施タイミング', '部署', '担当役割', '作業内容'];
    
    // 空値の自動補完
    if (!row['工程'] || row['工程'].trim() === '') {
      row['工程'] = `フェーズ${i + 1}`;
    }
    if (!row['実施タイミング'] || row['実施タイミング'].trim() === '') {
      row['実施タイミング'] = `第${i + 1}期`;
    }
    if (!row['部署'] || row['部署'].trim() === '') {
      row['部署'] = '経営企画部';
    }
    if (!row['担当役割'] || row['担当役割'].trim() === '') {
      row['担当役割'] = '担当者';
    }
    if (!row['作業内容'] || row['作業内容'].trim() === '') {
      row['作業内容'] = 'タスク実施';
    }
    
    // オプションフィールドのデフォルト値
    if (!row['条件分岐']) row['条件分岐'] = 'なし';
    if (!row['利用ツール']) row['利用ツール'] = '手動作業';
    if (!row['URLリンク']) row['URLリンク'] = 'なし';
    if (!row['備考']) row['備考'] = '特になし';
  }
  
  return true;
}

// ================================================================================
// 7. parser_and_writer.gs - データ解析・書き込み処理機能
// ================================================================================

// データ解析・書き込み処理

// フローシートのヘッダー定義
const FLOW_HEADERS = [
  '工程', 
  '実施タイミング', 
  '部署', 
  '担当役割', 
  '作業内容', 
  '条件分岐', 
  '利用ツール', 
  'URLリンク', 
  '備考'
];

// 業務記述書の書き込み
function writeWorkSpec(workSpec) {
  const sh = ss().getSheetByName(SPEC_SHEET) || createWorkSpecSheet();
  
  // IDを生成
  const id = Utilities.getUuid();
  const timestamp = new Date();
  
  // データ整形
  const rowData = [
    id,
    timestamp,
    workSpec.title || '',
    workSpec.summary || '',
    workSpec.scope || '',
    formatArray(workSpec.deliverables),
    formatArray(workSpec.org_structure),
    formatRaci(workSpec.raci),
    formatTimeline(workSpec.timeline),
    formatArray(workSpec.requirements_constraints),
    formatArray(workSpec.risks_mitigations),
    formatArray(workSpec.pro_considerations),
    formatArray(workSpec.kpi_sla),
    formatArray(workSpec.approvals),
    formatArray(workSpec.security_privacy_controls),
    formatLegalRegulations(workSpec.legal_regulations),
    formatArray(workSpec.references),
    formatArray(workSpec.assumptions)
  ];
  
  // データ書き込み
  sh.appendRow(rowData);
  
  // 書式設定
  const lastRow = sh.getLastRow();
  sh.getRange(lastRow, 1, 1, rowData.length).setWrap(true);
  sh.getRange(lastRow, 3).setFontWeight('bold'); // タイトルを太字
  
  logActivity('WRITE_SPEC', `Written work spec: ${workSpec.title}`);
}

// 業務記述書シート作成
function createWorkSpecSheet() {
  const sh = ss().insertSheet(SPEC_SHEET);
  
  const headers = [
    'ID',
    '作成日時',
    'タイトル',
    '概要',
    'スコープ',
    '成果物',
    '体制',
    'RACI',
    'スケジュール',
    '要件・制約',
    'リスク・対策',
    'プロ水準留意事項',
    'KPI/SLA',
    '承認フロー',
    'セキュリティ/個情保/内部統制',
    '法令・規制',
    '参考URL',
    '仮定条件'
  ];
  
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  sh.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  sh.getRange(1, 1, 1, headers.length).setBackground('#f0f0f0');
  sh.setFrozenRows(1);
  
  // 列幅調整
  sh.setColumnWidth(1, 150); // ID
  sh.setColumnWidth(2, 120); // 作成日時
  sh.setColumnWidth(3, 200); // タイトル
  sh.setColumnWidth(4, 300); // 概要
  
  return sh;
}

// フロー行の書き込み（レガシー関数 - 改善版にリダイレクト）
function writeFlowRows(flowRows) {
  // 新しい安全な実装を使用
  return writeFlowRowsSafe(flowRows);
}

// フローシート作成
function createFlowSheet(sheetName) {
  const sh = ss().insertSheet(sheetName);
  
  // ヘッダー行を設定
  sh.getRange(1, 1, 1, FLOW_HEADERS.length).setValues([FLOW_HEADERS]);
  sh.getRange(1, 1, 1, FLOW_HEADERS.length).setFontWeight('bold');
  sh.getRange(1, 1, 1, FLOW_HEADERS.length).setBackground('#f0f0f0');
  sh.setFrozenRows(1);
  
  // 列幅を調整
  sh.setColumnWidth(1, 120); // 工程
  sh.setColumnWidth(2, 150); // 実施タイミング
  sh.setColumnWidth(3, 120); // 部署
  sh.setColumnWidth(4, 150); // 担当役割
  sh.setColumnWidth(5, 300); // 作業内容
  sh.setColumnWidth(6, 150); // 条件分岐
  sh.setColumnWidth(7, 150); // 利用ツール
  sh.setColumnWidth(8, 200); // URLリンク
  sh.setColumnWidth(9, 200); // 備考
  
  return sh;
}

// 2番目のcreateFlowSheet関数に移動
    }
    
    if (values.length === 0) {
      console.log('データが空のため書き込みをスキップします');
      return;
    }
    
    // 各行の構造をチェック
    for (let i = 0; i < values.length; i++) {
      if (!Array.isArray(values[i])) {
        console.error(`Row ${i + 1} が配列ではありません:`, typeof values[i], values[i]);
        throw new Error(`Row ${i + 1} が配列ではありません: ${typeof values[i]}`);
      }
      
      if (values[i].length !== FLOW_HEADERS.length) {
        console.warn(`Row ${i + 1} の列数が不正: 期待値=${FLOW_HEADERS.length}, 実際=${values[i].length}`);
        // 列数を調整
        while (values[i].length < FLOW_HEADERS.length) {
          values[i].push('');
        }
        values[i] = values[i].slice(0, FLOW_HEADERS.length);
      }
      
      // 各セルの型をチェック
      for (let j = 0; j < values[i].length; j++) {
        if (typeof values[i][j] !== 'string') {
          console.warn(`非文字列データ検出: Row ${i + 1}, Col ${j + 1}, Type: ${typeof values[i][j]}, Value: ${values[i][j]}`);
          values[i][j] = String(values[i][j]);
        }
      }
    }
    
    console.log('最終データサンプル:', values.slice(0, 2));
    
    // setValues の引数を強制的に2次元配列として再構築
    const safeValues = values.map(row => {
      if (Array.isArray(row)) {
        return row.slice(); // 配列のコピーを作成
      } else {
        // 文字列の場合はカンマで分割して配列に変換
        console.warn('文字列データを配列に変換:', row);
        const stringRow = String(row);
        // CSVの適切な解析（引用符内のカンマを考慮）
        const splitRow = parseCSVLine(stringRow);
        // 列数を調整
        while (splitRow.length < FLOW_HEADERS.length) {
          splitRow.push('');
        }
        return splitRow.slice(0, FLOW_HEADERS.length);
      }
    });
    
    console.log('安全な形式に変換後:', safeValues.slice(0, 2));
    
    // 配列の長さを安全に取得
    const rowCount = Array.isArray(safeValues) ? safeValues.length : 0;
    const colCount = FLOW_HEADERS.length;
    
    if (rowCount === 0) {
      console.log('書き込むデータがありません');
      return;
    }
    
    // 数値として明示的に扱う
    const startRow = 2;
    const startCol = 1;
    
    console.log(`書き込み範囲: 行${startRow}から${rowCount}行、列${startCol}から${colCount}列`);
    
    sh.getRange(startRow, startCol, rowCount, colCount).setValues(safeValues);
    console.log('データ書き込み成功');
  } catch (error) {
    console.error('データ書き込みエラー:', error.message);
    console.error('エラー詳細:', error.stack);
    console.error('問題のあるデータ構造:', {
      type: typeof values,
      isArray: Array.isArray(values),
      length: values ? values.length : 'N/A',
      sample: values ? values.slice(0, 2) : 'N/A'
    });
    
    // 代替処理：1行ずつ書き込みを試行
    console.log('代替処理を開始：1行ずつ書き込み');
    try {
      for (let i = 0; i < values.length; i++) {
        try {
          // 行データを安全な配列形式に変換
          let rowData = values[i];
          if (!Array.isArray(rowData)) {
            console.warn(`Row ${i + 1} を配列に変換:`, rowData);
            rowData = parseCSVLine(String(rowData));
            while (rowData.length < FLOW_HEADERS.length) {
              rowData.push('');
            }
            rowData = rowData.slice(0, FLOW_HEADERS.length);
          }
          
          sh.getRange(i + 2, 1, 1, FLOW_HEADERS.length).setValues([rowData]);
          console.log(`Row ${i + 1} 書き込み成功`);
        } catch (rowError) {
          console.error(`Row ${i + 1} 書き込みエラー:`, rowError.message);
          console.error(`問題のある行データ:`, values[i]);
          
          // さらなる代替処理：セルごとに書き込み
          let rowData = values[i];
          if (!Array.isArray(rowData)) {
            rowData = parseCSVLine(String(rowData));
          }
          
          for (let j = 0; j < FLOW_HEADERS.length; j++) {
            try {
              const cellValue = rowData[j] || '';
              sh.getRange(i + 2, j + 1).setValue(String(cellValue).trim());
            } catch (cellError) {
              console.error(`Row ${i + 1}, Col ${j + 1} セル書き込みエラー:`, cellError.message);
              sh.getRange(i + 2, j + 1).setValue('エラー');
            }
          }
        }
      }
      console.log('代替処理完了');
    } catch (fallbackError) {
      console.error('代替処理も失敗:', fallbackError.message);
      throw error; // 元のエラーを再スロー
    }
  }
  
  // 書式設定（安全な変数を使用）
  const finalRowCount = Math.max(1, flowRows.length);
  try {
    sh.getRange(2, 1, finalRowCount, FLOW_HEADERS.length).setWrap(true);
    sh.getRange(2, 1, finalRowCount, 1).setFontWeight('bold'); // 工程列を太字
    
    // 条件分岐がある行に背景色設定
    for (let i = 0; i < finalRowCount; i++) {
      try {
        const rowData = sh.getRange(i + 2, 6).getValue(); // 条件分岐列の値を取得
        if (rowData && rowData !== 'なし' && rowData !== '') {
          sh.getRange(i + 2, 1, 1, FLOW_HEADERS.length).setBackground('#fff3cd');
        }
      } catch (formatError) {
        console.warn(`Row ${i + 2} の書式設定エラー:`, formatError.message);
      }
    }
  } catch (formatError) {
    console.warn('書式設定エラー:', formatError.message);
  }
  
  logActivity('WRITE_FLOW', `Written ${flowRows.length} flow rows`);
}

// フローシート作成
function createFlowSheet(sheetName) {
  const sh = ss().insertSheet(sheetName);
  
  sh.getRange(1, 1, 1, FLOW_HEADERS.length).setValues([FLOW_HEADERS]);
  sh.getRange(1, 1, 1, FLOW_HEADERS.length).setFontWeight('bold');
  sh.getRange(1, 1, 1, FLOW_HEADERS.length).setBackground('#e8f5e9');
  sh.setFrozenRows(1);
  
  // 列幅調整
  sh.setColumnWidth(1, 100); // 工程
  sh.setColumnWidth(2, 120); // 実施タイミング
  sh.setColumnWidth(3, 100); // 部署
  sh.setColumnWidth(4, 100); // 担当役割
  sh.setColumnWidth(5, 250); // 作業内容
  sh.setColumnWidth(6, 150); // 条件分岐
  sh.setColumnWidth(7, 120); // 利用ツール
  sh.setColumnWidth(8, 150); // URLリンク
  sh.setColumnWidth(9, 200); // 備考
  
  return sh;
}

// 配列データのフォーマット
function formatArray(arr) {
  if (!arr || !Array.isArray(arr)) return '';
  return arr.filter(item => item).join('\n');
}

// RACIマトリクスのフォーマット
function formatRaci(raciArray) {
  if (!raciArray || !Array.isArray(raciArray)) return '';
  
  return raciArray.map(item => {
    const roles = [];
    if (item.R) roles.push('R');
    if (item.A) roles.push('A');
    if (item.C) roles.push('C');
    if (item.I) roles.push('I');
    
    return `${item.dept || ''} - ${item.role || ''}: ${roles.join('')}`;
  }).join('\n');
}

// タイムラインのフォーマット
function formatTimeline(timeline) {
  if (!timeline || !Array.isArray(timeline)) return '';
  
  return timeline.map(phase => {
    let result = `【${phase.phase}】 ${phase.duration_hint || ''}`;
    
    if (phase.milestones && phase.milestones.length > 0) {
      result += '\nマイルストーン:\n' + phase.milestones.map(m => `  ・${m}`).join('\n');
    }
    
    if (phase.dependencies && phase.dependencies.length > 0) {
      result += '\n依存関係:\n' + phase.dependencies.map(d => `  ・${d}`).join('\n');
    }
    
    return result;
  }).join('\n\n');
}

// 法令・規制のフォーマット
function formatLegalRegulations(regulations) {
  if (!regulations || !Array.isArray(regulations)) return '';
  
  const formatted = regulations.map(reg => {
    let result = reg.name || '';
    if (reg.scope) result += `（${reg.scope}）`;
    if (reg.note) result += `: ${reg.note}`;
    return result;
  }).join('\n');
  
  // 法的助言の免責事項を追加
  return formatted + '\n\n※ 上記は参考情報です。最終的な判断は法務・専門家にご確認ください。法的助言ではありません。';
}

// 生データ保存（エラー時のフォールバック）
function saveRawData(data, error) {
  const sheetName = '業務記述書（Raw）';
  let sh = ss().getSheetByName(sheetName);
  
  if (!sh) {
    sh = ss().insertSheet(sheetName);
    sh.getRange(1, 1, 1, 4).setValues([['タイムスタンプ', 'エラー', 'データタイプ', '生データ']]);
    sh.getRange(1, 1, 1, 4).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  
  sh.appendRow([
    new Date(),
    error.toString(),
    typeof data,
    JSON.stringify(data, null, 2)
  ]);
  
  logActivity('SAVE_RAW', 'Saved raw data due to error');
}

// データ検証とクリーニング
function sanitizeData(data) {
  if (!data || typeof data !== 'object') return data;
  
  // 再帰的にオブジェクトをクリーニング
  const cleaned = {};
  for (const key in data) {
    if (data.hasOwnProperty(key)) {
      const value = data[key];
      
      if (value === null || value === undefined) {
        cleaned[key] = '';
      } else if (Array.isArray(value)) {
        cleaned[key] = value.map(item => 
          typeof item === 'object' ? sanitizeData(item) : item
        );
      } else if (typeof value === 'object') {
        cleaned[key] = sanitizeData(value);
      } else {
        cleaned[key] = value;
      }
    }
  }
  
  return cleaned;
}

// 個人情報マスキング（オプション）
function maskSensitiveInfo(text) {
  if (!text || typeof text !== 'string') return text;
  
  // メールアドレスのマスキング
  text = text.replace(/([a-zA-Z0-9._%+-]+)@([a-zA-Z0-9.-]+\.[a-zA-Z]{2,})/g, 
    (match, p1, p2) => p1.substring(0, 2) + '***@' + p2);
  
  // 電話番号のマスキング（日本の形式）
  text = text.replace(/(\d{2,4})-(\d{2,4})-(\d{4})/g, '$1-****-****');
  text = text.replace(/0\d{1,4}-\d{1,4}-\d{4}/g, '0**-****-****');
  
  // 郵便番号のマスキング
  text = text.replace(/〒?\d{3}-\d{4}/g, '〒***-****');
  
  return text;
}

// ================================================================================
// 9. data_processor.gs - 根本的に改善されたデータ処理エンジン
// ================================================================================

// データ処理の根本的改善
// 新しいデータ処理エンジン

// データ正規化クラス
class DataNormalizer {
  constructor() {
    this.cleaningPatterns = [
      // 末尾数字パターン
      { pattern: /\d+$/, replacement: '', condition: (str) => str.length > 1 },
      // 特定文字列後の数字
      { pattern: /特になし\d+$/, replacement: '特になし', condition: () => true },
      { pattern: /なし\d+$/, replacement: 'なし', condition: () => true },
      // 一般的な末尾数字（文字の後に数字）
      { pattern: /([^\d])\d+$/, replacement: '$1', condition: (str) => str.length > 2 }
    ];
  }

  // 文字列のクリーニング
  cleanString(value) {
    if (!value || typeof value !== 'string') {
      return String(value || '').trim();
    }

    let cleaned = value.trim();
    const original = cleaned;

    for (const rule of this.cleaningPatterns) {
      if (rule.condition(cleaned) && rule.pattern.test(cleaned)) {
        const newValue = cleaned.replace(rule.pattern, rule.replacement).trim();
        if (newValue.length > 0 && newValue !== cleaned) {
          console.log(`データクリーニング: "${cleaned}" -> "${newValue}"`);
          cleaned = newValue;
          break; // 最初にマッチしたルールのみ適用
        }
      }
    }

    return cleaned;
  }

  // フロー行データの正規化
  normalizeFlowRow(row, index) {
    const normalizedRow = {};
    const requiredFields = ['工程', '実施タイミング', '部署', '担当役割', '作業内容'];
    const optionalFields = ['条件分岐', '利用ツール', 'URLリンク', '備考'];

    // 必須フィールドの処理
    for (const field of requiredFields) {
      let value = row[field];
      
      if (!value || String(value).trim() === '') {
        // デフォルト値を設定
        switch (field) {
          case '工程':
            value = `フェーズ${index + 1}`;
            break;
          case '実施タイミング':
            value = `第${index + 1}期`;
            break;
          case '部署':
            value = '経営企画部';
            break;
          case '担当役割':
            value = '担当者';
            break;
          case '作業内容':
            value = 'タスク実施';
            break;
        }
        console.log(`デフォルト値設定: ${field} = "${value}"`);
      }

      normalizedRow[field] = this.cleanString(value);
    }

    // オプションフィールドの処理
    for (const field of optionalFields) {
      let value = row[field] || '';
      if (field === '条件分岐' && !value) {
        value = 'なし';
      }
      normalizedRow[field] = this.cleanString(value);
    }

    return normalizedRow;
  }

  // 2次元配列への変換（厳密な型保証）
  convertToSpreadsheetArray(flowRows, headers) {
    if (!Array.isArray(flowRows) || !Array.isArray(headers)) {
      throw new Error('Invalid input: flowRows and headers must be arrays');
    }

    const result = [];
    
    for (let i = 0; i < flowRows.length; i++) {
      const row = flowRows[i];
      const arrayRow = [];

      for (const header of headers) {
        const value = row[header] || '';
        arrayRow.push(String(value)); // 明示的に文字列変換
      }

      // 配列の長さを確認
      if (arrayRow.length !== headers.length) {
        throw new Error(`Row ${i + 1}: Expected ${headers.length} columns, got ${arrayRow.length}`);
      }

      result.push(arrayRow);
    }

    console.log(`変換完了: ${result.length}行 x ${headers.length}列の2次元配列`);
    return result;
  }
}

// 安全なスプレッドシート書き込みクラス
class SafeSpreadsheetWriter {
  constructor(sheet, headers) {
    this.sheet = sheet;
    this.headers = headers;
    this.normalizer = new DataNormalizer();
  }

  // データの書き込み（複数の安全策を実装）
  writeData(flowRows) {
    if (!flowRows || flowRows.length === 0) {
      console.log('書き込むデータがありません');
      return;
    }

    try {
      // Step 1: データの正規化
      const normalizedRows = flowRows.map((row, index) => 
        this.normalizer.normalizeFlowRow(row, index)
      );

      // Step 2: 2次元配列への変換
      const spreadsheetArray = this.normalizer.convertToSpreadsheetArray(normalizedRows, this.headers);

      // Step 3: 既存データのクリア
      this.clearExistingData();

      // Step 4: 安全な書き込み
      this.performSafeWrite(spreadsheetArray);

      // Step 5: 書式設定
      this.applyFormatting(spreadsheetArray.length);

      console.log(`データ書き込み完了: ${spreadsheetArray.length}行`);

    } catch (error) {
      console.error('データ書き込みエラー:', error);
      throw new Error(`データ書き込み失敗: ${error.message}`);
    }
  }

  // 既存データのクリア
  clearExistingData() {
    const lastRow = this.sheet.getLastRow();
    if (lastRow > 1) {
      const dataRows = Math.max(1, lastRow - 1);
      this.sheet.getRange(2, 1, dataRows, this.headers.length).clearContent();
    }
  }

  // 安全な書き込み実行
  performSafeWrite(data) {
    try {
      // データの検証
      if (!Array.isArray(data) || data.length === 0) {
        console.log('書き込むデータが無効または空です');
        return;
      }
      
      // データ構造のログ出力
      console.log(`書き込みデータ: ${data.length}行 x ${this.headers.length}列`);
      console.log('最初の行サンプル:', data[0]);
      
      // 数値パラメータの検証
      const startRow = 2;
      const startCol = 1;
      const numRows = Number(data.length);
      const numCols = Number(this.headers.length);
      
      if (isNaN(numRows) || isNaN(numCols)) {
        throw new Error(`無効な行数または列数: rows=${numRows}, cols=${numCols}`);
      }
      
      // 一括書き込みを試行
      this.sheet.getRange(startRow, startCol, numRows, numCols).setValues(data);
      console.log('一括書き込み成功');
    } catch (error) {
      console.warn('一括書き込み失敗、行ごと書き込みに切り替え:', error.message);
      console.error('エラー詳細:', error);
      this.writeRowByRow(data);
    }
  }

  // 行ごとの書き込み
  writeRowByRow(data) {
    for (let i = 0; i < data.length; i++) {
      try {
        const rowNum = Number(i + 2);
        const colStart = 1;
        const numRows = 1;
        const numCols = Number(this.headers.length);
        
        if (!Array.isArray(data[i])) {
          console.error(`Row ${i + 1} が配列ではありません:`, data[i]);
          continue;
        }
        
        this.sheet.getRange(rowNum, colStart, numRows, numCols).setValues([data[i]]);
        console.log(`Row ${i + 1} 書き込み成功`);
      } catch (error) {
        console.error(`Row ${i + 1} 書き込みエラー:`, error.message);
        console.error(`問題のデータ:`, data[i]);
        // セルごとの書き込みにフォールバック
        this.writeCellByCell(i + 2, data[i]);
      }
    }
  }

  // セルごとの書き込み
  writeCellByCell(rowIndex, rowData) {
    for (let j = 0; j < rowData.length; j++) {
      try {
        this.sheet.getRange(rowIndex, j + 1).setValue(rowData[j]);
      } catch (error) {
        console.error(`Cell (${rowIndex}, ${j + 1}) 書き込みエラー:`, error.message);
        this.sheet.getRange(rowIndex, j + 1).setValue('エラー');
      }
    }
  }

  // 書式設定
  applyFormatting(rowCount) {
    try {
      // テキストの折り返し
      this.sheet.getRange(2, 1, rowCount, this.headers.length).setWrap(true);
      
      // 工程列を太字
      this.sheet.getRange(2, 1, rowCount, 1).setFontWeight('bold');
      
      // 条件分岐がある行の背景色設定
      for (let i = 0; i < rowCount; i++) {
        const conditionValue = this.sheet.getRange(i + 2, 6).getValue(); // 条件分岐列
        if (conditionValue && conditionValue !== 'なし' && conditionValue !== '') {
          this.sheet.getRange(i + 2, 1, 1, this.headers.length).setBackground('#fff3cd');
        }
      }
    } catch (error) {
      console.warn('書式設定エラー:', error.message);
    }
  }
}

// 改善されたフロー行書き込み関数（新しい安全な実装にリダイレクト）
function writeFlowRowsImproved(flowRows) {
  return writeFlowRowsSafe(flowRows);
}

// デバッグ用関数：データ型と内容を詳細に出力
function debugDataStructure(data, label = 'データ') {
  console.log('\n========== デバッグ情報開始 ==========');
  console.log(`【${label}】`);
  console.log('データ型:', typeof data);
  console.log('null/undefined?:', data === null || data === undefined);
  console.log('配列?:', Array.isArray(data));
  
  if (data === null || data === undefined) {
    console.log('データは null または undefined です');
    console.log('========== デバッグ情報終了 ==========\n');
    return;
  }
  
  if (Array.isArray(data)) {
    console.log('配列の長さ:', data.length);
    console.log('最初の3要素の詳細:');
    for (let i = 0; i < Math.min(3, data.length); i++) {
      console.log(`  [${i}] 型: ${typeof data[i]}`);
      if (typeof data[i] === 'string') {
        console.log(`      値: "${data[i].substring(0, 100)}${data[i].length > 100 ? '...' : ''}"`);
        console.log(`      長さ: ${data[i].length}文字`);
        console.log(`      カンマの数: ${(data[i].match(/,/g) || []).length}`);
      } else if (Array.isArray(data[i])) {
        console.log(`      配列長: ${data[i].length}`);
        console.log(`      内容: [${data[i].slice(0, 3).map(v => typeof v).join(', ')}...]`);
      } else if (typeof data[i] === 'object' && data[i] !== null) {
        console.log(`      キー: ${Object.keys(data[i]).slice(0, 5).join(', ')}`);
      } else {
        console.log(`      値: ${data[i]}`);
      }
    }
  } else if (typeof data === 'string') {
    console.log('文字列の長さ:', data.length);
    console.log('最初の200文字:', data.substring(0, 200) + (data.length > 200 ? '...' : ''));
    console.log('改行の数:', (data.match(/\n/g) || []).length);
    console.log('カンマの数:', (data.match(/,/g) || []).length);
    console.log('最初の行:', data.split('\n')[0]);
  } else if (typeof data === 'object') {
    const keys = Object.keys(data);
    console.log('オブジェクトのキー数:', keys.length);
    console.log('最初の10個のキー:', keys.slice(0, 10).join(', '));
    console.log('最初の3つのキーと値:');
    for (let i = 0; i < Math.min(3, keys.length); i++) {
      const key = keys[i];
      const value = data[key];
      console.log(`  ${key}: (${typeof value}) ${String(value).substring(0, 50)}${String(value).length > 50 ? '...' : ''}`);
    }
  } else {
    console.log('その他の型のデータ:', data);
  }
  
  console.log('========== デバッグ情報終了 ==========\n');
}

// 完全に安全な新しいフロー行書き込み関数
function writeFlowRowsSafe(flowRows) {
  const sheetName = getConfig('FLOW_SHEET_NAME') || FLOW_SHEET;
  const sheet = ss().getSheetByName(sheetName) || createFlowSheet(sheetName);
  const headers = [
    '工程', '実施タイミング', '部署', '担当役割', '作業内容', 
    '条件分岐', '利用ツール', 'URLリンク', '備考'
  ];

  console.log('=== 安全なフロー行書き込み開始 ===');
  
  // 詳細なデバッグ情報を出力
  debugDataStructure(flowRows, '入力データ (flowRows)');
  
  // データを安全に処理
  let processedData = [];
  
  try {
    // flowRowsが配列かどうかチェック
    if (!flowRows) {
      console.log('データがnullまたはundefined');
      return;
    }
    
    if (Array.isArray(flowRows)) {
      console.log(`配列として受信: ${flowRows.length}個の要素`);
      
      // 各要素を安全に処理
      for (let i = 0; i < flowRows.length; i++) {
        const row = flowRows[i];
        console.log(`\n--- 行${i + 1}の処理開始 ---`);
        debugDataStructure(row, `行${i + 1}`);
        
        if (typeof row === 'object' && row !== null) {
          // オブジェクトの場合、ヘッダーに基づいて配列を作成
          const rowArray = [];
          for (const header of headers) {
            const value = row[header] || '';
            // 末尾の不要な数字を削除
            const cleanValue = String(value).replace(/特になし\d+$/, '特になし').replace(/なし\d+$/, 'なし');
            rowArray.push(cleanValue);
          }
          processedData.push(rowArray);
        } else if (typeof row === 'string') {
          // 文字列の場合、カンマで分割して配列に変換
          console.log(`行${i + 1}は文字列です。解析を試みます`);
          console.log('文字列の内容（最初の100文字）:', row.substring(0, 100));
          const parts = row.split(',').map(part => part.trim());
          console.log('分割後の要素数:', parts.length);
          console.log('分割結果:', parts);
          const rowArray = [];
          for (let j = 0; j < headers.length; j++) {
            const value = parts[j] || '';
            const cleanValue = String(value).replace(/特になし\d+$/, '特になし').replace(/なし\d+$/, 'なし');
            rowArray.push(cleanValue);
          }
          processedData.push(rowArray);
        } else if (Array.isArray(row)) {
          // 既に配列の場合
          const rowArray = [];
          for (let j = 0; j < headers.length; j++) {
            const value = row[j] || '';
            const cleanValue = String(value).replace(/特になし\d+$/, '特になし').replace(/なし\d+$/, 'なし');
            rowArray.push(cleanValue);
          }
          processedData.push(rowArray);
        }
      }
    } else if (typeof flowRows === 'string') {
      // 全体が文字列の場合、行ごとに分割してから処理
      console.log('全体が文字列として受信');
      console.log('文字列の長さ:', flowRows.length);
      console.log('最初の200文字:', flowRows.substring(0, 200));
      const lines = flowRows.split('\n').filter(line => line.trim());
      console.log('行数:', lines.length);
      if (lines.length > 0) {
        console.log('最初の行:', lines[0]);
      }
      for (let i = 0; i < lines.length; i++) {
        const parts = lines[i].split(',').map(part => part.trim());
        const rowArray = [];
        for (let j = 0; j < headers.length; j++) {
          const value = parts[j] || '';
          const cleanValue = String(value).replace(/特になし\d+$/, '特になし').replace(/なし\d+$/, 'なし');
          rowArray.push(cleanValue);
        }
        processedData.push(rowArray);
      }
    } else if (typeof flowRows === 'object') {
      // 単一のオブジェクトの場合
      console.log('単一のオブジェクトとして受信');
      const rowArray = [];
      for (const header of headers) {
        const value = flowRows[header] || '';
        const cleanValue = String(value).replace(/特になし\d+$/, '特になし').replace(/なし\d+$/, 'なし');
        rowArray.push(cleanValue);
      }
      processedData.push(rowArray);
    } else {
      console.error('サポートされていないデータ型:', typeof flowRows);
      return;
    }
    
    // 処理済みデータのデバッグ情報
    console.log('\n=== 処理済みデータの確認 ===');
    console.log('処理済み行数:', processedData.length);
    if (processedData.length > 0) {
      console.log('最初の行のデータ:', processedData[0]);
    }
    
    // データが存在しない場合は終了
    if (processedData.length === 0) {
      console.log('処理可能なデータがありません');
      return;
    }
    
    // 既存データをクリア
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      const clearRows = lastRow - 1;
      const clearCols = headers.length;
      console.log(`既存データをクリア: ${clearRows}行 x ${clearCols}列`);
      sheet.getRange(2, 1, clearRows, clearCols).clearContent();
    }
    
    // セルごとに安全に書き込み（エラーを完全に回避）
    console.log(`書き込み開始: ${processedData.length}行`);
    for (let i = 0; i < processedData.length; i++) {
      for (let j = 0; j < headers.length; j++) {
        try {
          const cellValue = processedData[i][j] || '';
          // セルごとに個別に書き込み（最も安全）
          sheet.getRange(i + 2, j + 1).setValue(cellValue);
        } catch (cellError) {
          console.error(`セル(${i + 2}, ${j + 1})書き込みエラー:`, cellError.message);
          console.error('エラーの詳細:', cellError);
          console.error('書き込もうとした値:', processedData[i][j]);
          console.error('値の型:', typeof processedData[i][j]);
          // エラーが発生してもデフォルト値を設定
          try {
            sheet.getRange(i + 2, j + 1).setValue('');
          } catch (e) {
            // それでも失敗したら無視
          }
        }
      }
      console.log(`行${i + 1}書き込み完了`);
    }
    
    // 書式設定（エラーが発生しても続行）
    try {
      sheet.getRange(2, 1, processedData.length, headers.length).setWrap(true);
      sheet.getRange(2, 1, processedData.length, 1).setFontWeight('bold');
    } catch (e) {
      console.warn('書式設定エラー:', e.message);
    }
    
    console.log('=== フロー行書き込み完了 ===');
    logActivity('WRITE_FLOW_SAFE', `Successfully written ${processedData.length} flow rows`);
    
  } catch (error) {
    console.error('致命的なエラー:', error.message);
    console.error('スタックトレース:', error.stack);
    // エラーが発生してもアプリケーションは続行
  }
}

/*
================================================================================
                                    終了
================================================================================
*/