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
    ['PROCESSING_QUERY', 'subject:[task] is:unread'],
    ['DEFAULT_TO_EMAIL', ''],
    ['OPENAI_MODEL', 'o3-deep-research'],  // 最新のgpt-5を使用
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
  // 既存のシートがないか再確認
  let sh = ss().getSheetByName(ACTIVITY_LOG_SHEET);
  if (sh) {
    console.log('ActivityLogシートは既に存在します');
    return sh;
  }
  
  try {
    sh = ss().insertSheet(ACTIVITY_LOG_SHEET);
    sh.getRange(1, 1, 1, 4).setValues([[
      'タイムスタンプ', 'タイプ', '詳細', 'ユーザー'
    ]]);
    sh.getRange(1, 1, 1, 4).setFontWeight('bold');
    sh.hideSheet();
    console.log('ActivityLogシートを作成しました');
  } catch (e) {
    console.error('ActivityLogシート作成エラー:', e.toString());
    // エラーが発生した場合は、既存のシートを探す
    const sheets = ss().getSheets();
    for (const sheet of sheets) {
      if (sheet.getName().toLowerCase() === 'activitylog') {
        console.log('ActivityLogシートが別の形式で存在しています');
        return sheet;
      }
    }
  }
  
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
    console.error('フローシートが見つかりません。');
    return;
  }
  
  // ビジュアルフローシートをクリア
  visualSheet.clear();
  visualSheet.clearFormats();
  
  // フローデータを取得
  const flowData = flowSheet.getDataRange().getValues();
  if (flowData.length <= 1) {
    console.error('フローデータがありません。');
    return;
  }
  
  const headers = flowData[0];
  const rows = flowData.slice(1).filter(row => row[0]); // 工程が入力されている行のみ
  
  if (rows.length === 0) {
    console.error('有効なフローデータがありません。');
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
  
  console.log('ビジュアルフローを生成しました。');
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
  
  console.log('サンプルフローデータを作成しました。');
}

// 高度なビジュアルフロー生成（業務フローチャート図作成.js参考版）
function generateAdvancedVisualFlow() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(FLOW_SHEET);
    if (!sheet) {
      console.error('フローシートが見つかりません。');
      return;
    }
    
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) {
      console.error('データがありません。');
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
    
    console.log('高度なビジュアルフローが生成されました。');
    
  } catch (error) {
    console.error('エラー:', error);
    console.error('フロー生成中にエラーが発生しました:', error);
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
      console.error('フローシートが見つかりません');
      return;
    }
    
    const data = flowSheet.getDataRange().getValues();
    if (data.length < 2) {
      console.error('データがありません');
      return;
    }
    
    const flowData = parseAdvancedFlowData(data);
    createBusinessSummarySheet(flowData);
    
    console.log('業務サマリーを作成しました');
  } catch (error) {
    console.error('エラー:', error);
  }
}

// ================================================================================
// 3. gmail_inbound.gs - Gmail受信処理機能
// ================================================================================

// Gmail受信処理

// 新着メール処理（メイン関数）
function processNewEmails() {
  const query = getConfig('PROCESSING_QUERY') || 'subject:[task] is:unread';
  logActivity('PROCESS_START', `Processing emails with query: ${query}`);
  console.log('検索クエリ:', query);
  
  try {
    const threads = GmailApp.search(query);
    console.log('検索結果:', threads.length, '件のスレッドが見つかりました');
    
    if (threads.length === 0) {
      logActivity('PROCESS_INFO', 'No new emails found');
      
      // デバッグ用: 全ての未読メールを確認
      const allUnread = GmailApp.search('is:unread', 0, 5);
      console.log('未読メール総数:', allUnread.length);
      if (allUnread.length > 0) {
        console.log('未読メールの件名リスト:');
        allUnread.forEach(thread => {
          const firstMessage = thread.getMessages()[0];
          const subject = firstMessage.getSubject();
          console.log('  - "' + subject + '"');
          if (subject.toLowerCase().includes('task')) {
            console.log('    → このメールは"task"を含んでいます');
          }
        });
      }
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
  
  // 件名から[task]を除去して、本文と結合
  const cleanSubject = subject.replace(/\[task\]/gi, '').trim();
  const combinedContent = `【件名】${cleanSubject}\n\n【本文】\n${plainBody}`;
  
  // Inboxにログ記録
  logInbox(messageId, thread.getId(), from, subject, plainBody.substring(0, 200), 'NEW');
  
  try {
    // OpenAI呼び出し（件名と本文を結合したものを送信）
    const orgProfile = getConfig('ORG_PROFILE_JSON') || '{}';
    const result = callOpenAI(combinedContent, orgProfile);
    
    // 検証
    validateOpenAIResponse(result);
    
    // ガバナンスチェックを自動実行
    console.log('=== ガバナンスチェック開始 ===');
    const governanceCheck = performComprehensiveGovernanceCheck(result.work_spec, result.flow_rows);
    console.log('ガバナンススコア:', governanceCheck.overallScore);
    console.log('開示要件:', governanceCheck.disclosureRequirements.length, '件');
    console.log('要専門家相談:', governanceCheck.advisorConsultations ? governanceCheck.advisorConsultations.length : 0, '件');
    
    // 新規スプレッドシートを作成（メールごとに独立）
    const newSpreadsheet = createIndependentSpreadsheetWithGovernance(cleanSubject, result, governanceCheck);
    
    // 共有設定（URLを知っている人は誰でも編集可能）
    setPublicEditAccess(newSpreadsheet);
    
    // 返信メール送信（新規スプレッドシートのURLを送信）
    sendNotificationEmail(from, result.work_spec, newSpreadsheet.getUrl());
    
    // 処理済みマーク
    markProcessed(messageId);
    labelThreadProcessed(thread);
    
    logActivity('PROCESS_SUCCESS', `Successfully processed message ${messageId}`);
  } catch (e) {
    logError(messageId, e);
    
    // エラーの詳細情報を取得
    let errorDetails = '';
    if (e.message) {
      errorDetails = e.message;
    } else if (e.toString) {
      errorDetails = e.toString();
    } else {
      errorDetails = String(e);
    }
    
    // スタックトレースがある場合は追加
    if (e.stack) {
      console.error('エラースタックトレース:', e.stack);
    }
    
    // エラー通知メール送信（詳細情報付き）
    sendErrorNotificationEmail(from, subject, errorDetails);
    
    // 元のエラーを再スロー（ただしint変換エラーは特別処理）
    if (errorDetails.includes('Cannot convert') && errorDetails.includes('to int')) {
      console.error('int変換エラーを検出。データ形式の問題の可能性があります。');
      // int変換エラーの場合は処理を続行しない
      return;
    }
    
    throw e;
  }
}

// メールごとに独立したスプレッドシートを作成
function createIndependentSpreadsheet(subject, result) {
  // 後方互換性のため、ガバナンスチェックなしで作成
  return createIndependentSpreadsheetWithGovernance(subject, result, null);
}

// ガバナンスチェック結果を含むスプレッドシートを作成
function createIndependentSpreadsheetWithGovernance(subject, result, governanceCheck) {
  console.log('=== 新規スプレッドシート作成開始（ガバナンス機能付き） ===');
  
  // スプレッドシート名を生成（日時を含む）
  const now = new Date();
  const dateStr = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyyMMdd_HHmmss');
  const spreadsheetName = `【業務記述書】${subject}_${dateStr}`;
  
  // 新規スプレッドシートを作成
  const newSpreadsheet = SpreadsheetApp.create(spreadsheetName);
  console.log('新規スプレッドシート作成:', newSpreadsheet.getUrl());
  
  // 1. デフォルトのシートを取得して業務サマリシートを作成
  const defaultSheet = newSpreadsheet.getSheets()[0];
  defaultSheet.setName('業務サマリ');
  console.log('業務サマリシート作成開始');
  createSummarySheetWithGovernance(defaultSheet, result.work_spec, governanceCheck);
  console.log('業務サマリシート作成完了');
  
  // スプレッドシートを明示的に保存
  SpreadsheetApp.flush();
  
  // 2. 業務記述書シートを作成
  console.log('業務記述書シート作成開始');
  const specSheet = newSpreadsheet.insertSheet('業務記述書');
  writeWorkSpecToSheet(specSheet, result.work_spec);
  console.log('業務記述書シート作成完了');
  
  // スプレッドシートを明示的に保存
  SpreadsheetApp.flush();
  
  // 3. フローシートを作成（ガバナンス情報付き）
  console.log('フローシート作成開始');
  const flowSheet = newSpreadsheet.insertSheet('フロー');
  const cleanedFlowRows = cleanFlowRowsData(result.flow_rows);
  writeFlowToSheetWithGovernance(flowSheet, cleanedFlowRows, governanceCheck);
  console.log('フローシート作成完了');
  
  // スプレッドシートを明示的に保存
  SpreadsheetApp.flush();
  
  // 4. ガバナンスレポートシートを作成
  if (governanceCheck) {
    console.log('ガバナンスレポート作成開始');
    try {
      const govSheet = newSpreadsheet.insertSheet('ガバナンス・コンプライアンス');
      createGovernanceReportSheet(govSheet, governanceCheck);
      console.log('ガバナンスレポート作成完了');
    } catch (error) {
      console.error('ガバナンスレポート作成エラー:', error);
    }
    SpreadsheetApp.flush();
  }
  
  // 5. 外部専門家相談チェックリストシートを作成
  if (governanceCheck && governanceCheck.advisorConsultations && governanceCheck.advisorConsultations.length > 0) {
    console.log('専門家相談チェックリスト作成開始');
    try {
      const consultSheet = newSpreadsheet.insertSheet('専門家相談チェックリスト');
      createConsultationChecklistSheet(consultSheet, governanceCheck.advisorConsultations);
      console.log('専門家相談チェックリスト作成完了');
    } catch (error) {
      console.error('専門家相談チェックリスト作成エラー:', error);
    }
    SpreadsheetApp.flush();
  }
  
  // 6. ビジュアルフローシートを作成（最後に作成して確実に実行）
  console.log('ビジュアルフローシート作成開始');
  try {
    const visualSheet = newSpreadsheet.insertSheet('ビジュアルフロー');
    createVisualFlowInSheet(visualSheet, flowSheet);
    console.log('ビジュアルフローシート作成完了');
  } catch (error) {
    console.error('ビジュアルフロー作成エラー:', error);
    // エラーが発生しても処理を継続
  }
  
  // 最終保存
  SpreadsheetApp.flush();
  
  console.log('=== 新規スプレッドシート作成完了 ===');
  return newSpreadsheet;
}

// 業務サマリシートを作成
function createSummarySheet(sheet, workSpec) {
  // 後方互換性のため、ガバナンスチェックなしで作成
  createSummarySheetWithGovernance(sheet, workSpec, null);
}

// ガバナンス情報を含む業務サマリシートを作成
function createSummarySheetWithGovernance(sheet, workSpec, governanceCheck) {
  // タイトル
  sheet.getRange('A1').setValue('業務サマリ');
  sheet.getRange('A1').setFontSize(18).setFontWeight('bold').setBackground('#4285F4').setFontColor('#FFFFFF');
  sheet.getRange('A1:D1').merge();
  
  // 基本情報
  const summaryData = [
    ['項目', '内容'],
    ['タイトル', workSpec.title || ''],
    ['概要', workSpec.summary || ''],
    ['目的', workSpec.purpose || ''],
    ['対象範囲', workSpec.scope || ''],
    ['前提条件', formatArray(workSpec.prerequisites) || ''],
    ['成果物', formatArray(workSpec.deliverables) || ''],
    ['関係者', formatArray(workSpec.stakeholders) || ''],
    ['作成日時', new Date()],
    ['最終更新', new Date()]
  ];
  
  sheet.getRange(3, 1, summaryData.length, 2).setValues(summaryData);
  sheet.getRange(3, 1, summaryData.length, 1).setFontWeight('bold').setBackground('#F0F0F0');
  sheet.getRange(3, 2, summaryData.length, 1).setWrap(true);
  
  // 列幅調整
  sheet.setColumnWidth(1, 150);
  sheet.setColumnWidth(2, 500);
  
  // 罫線
  sheet.getRange(3, 1, summaryData.length, 2).setBorder(true, true, true, true, true, true);
}

// 業務記述書をシートに書き込み
function writeWorkSpecToSheet(sheet, workSpec) {
  // ヘッダー
  const headers = ['項目', '内容', '詳細', '備考'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#E8F5E9');
  
  // データ
  const specData = [
    ['タイトル', workSpec.title || '', '', ''],
    ['概要', workSpec.summary || '', '', ''],
    ['目的', workSpec.purpose || '', '', ''],
    ['対象範囲', workSpec.scope || '', '', ''],
    ['前提条件', formatArray(workSpec.prerequisites) || '', '', ''],
    ['必要なリソース', formatArray(workSpec.resources) || '', '', ''],
    ['成果物', formatArray(workSpec.deliverables) || '', '', ''],
    ['関係者', formatArray(workSpec.stakeholders) || '', '', ''],
    ['承認プロセス', workSpec.approval_process || '', '', ''],
    ['リスクと対策', formatRisks(workSpec.risks) || '', '', ''],
    ['期限・頻度', formatTimeline(workSpec.timeline) || '', '', ''],
    ['KPI/成功基準', formatArray(workSpec.kpis) || '', '', '']
  ];
  
  sheet.getRange(2, 1, specData.length, headers.length).setValues(specData);
  sheet.getRange(2, 1, specData.length, 4).setWrap(true);
  
  // 列幅調整
  sheet.setColumnWidth(1, 150);
  sheet.setColumnWidth(2, 400);
  sheet.setColumnWidth(3, 300);
  sheet.setColumnWidth(4, 200);
  
  // 罫線
  sheet.getRange(1, 1, specData.length + 1, headers.length).setBorder(true, true, true, true, true, true);
}

// フローをシートに書き込み
function writeFlowToSheet(sheet, flowRows) {
  // 後方互換性のため、ガバナンスチェックなしで書き込み
  writeFlowToSheetWithGovernance(sheet, flowRows, null);
}

// ガバナンス情報を含むフローをシートに書き込み
function writeFlowToSheetWithGovernance(sheet, flowRows, governanceCheck) {
  const headers = FLOW_HEADERS; // 定数を使用して一貫性を保つ
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#E8F5E9');
  sheet.setFrozenRows(1);
  
  // フローデータを処理して書き込み
  const processedData = [];
  
  if (Array.isArray(flowRows)) {
    for (let i = 0; i < flowRows.length; i++) {
      const row = flowRows[i];
      
      if (typeof row === 'object' && row !== null) {
        const workContent = row['作業内容'] || '';
        const actions = splitIntoActions(workContent);
        const processName = row['工程'] || '';
        const timing = row['実施タイミング'] || '';
        const dept = row['部署'] || '';
        const condition = row['条件分岐'] || '';
        
        for (let j = 0; j < actions.length; j++) {
          const rowArray = [];
          
          for (const header of headers) {
            let value = '';
            
            if (header === '作業内容') {
              value = actions[j];
            } else if (header === '法令・規制') {
              value = checkLegalRegulations(processName, actions[j], timing, dept);
            } else if (header === '内部統制') {
              value = checkInternalControl(processName, actions[j], condition, dept);
            } else if (header === 'コンプライアンス留意点') {
              value = j === 0 ? generateComplianceNotes(processName, actions[j], timing, dept, condition) : '';
            } else if (j === 0) {
              value = row[header] || '';
            } else {
              if (header === '工程' || header === '実施タイミング' || header === '部署' || header === '担当役割') {
                value = row[header] || '';
              } else {
                value = '';
              }
            }
            
            rowArray.push(value);
          }
          processedData.push(rowArray);
        }
      }
    }
  }
  
  if (processedData.length > 0) {
    sheet.getRange(2, 1, processedData.length, headers.length).setValues(processedData);
    sheet.getRange(2, 1, processedData.length, headers.length).setWrap(true);
  }
  
  // 列幅調整
  sheet.setColumnWidth(1, 120); // 工程
  sheet.setColumnWidth(2, 150); // 実施タイミング
  sheet.setColumnWidth(3, 120); // 部署
  sheet.setColumnWidth(4, 150); // 担当役割
  sheet.setColumnWidth(5, 300); // 作業内容
  sheet.setColumnWidth(6, 150); // 条件分岐
  sheet.setColumnWidth(7, 150); // 利用ツール
  sheet.setColumnWidth(8, 200); // URLリンク
  sheet.setColumnWidth(9, 200); // 備考
  sheet.setColumnWidth(10, 250); // 法令・規制
  sheet.setColumnWidth(11, 250); // 内部統制
  sheet.setColumnWidth(12, 300); // コンプライアンス留意点
  
  // 罫線
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(1, 1, lastRow, headers.length).setBorder(true, true, true, true, true, true);
  }
}

// ビジュアルフローをシートに作成
function createVisualFlowInSheet(visualSheet, flowSheet) {
  const data = flowSheet.getDataRange().getValues();
  if (data.length < 2) {
    console.log('フローデータがないためビジュアルフロー作成をスキップ');
    return;
  }
  
  const flowData = parseFlowDataForVisual(data);
  drawVisualFlowChart(visualSheet, flowData);
}

// URLを知っている人は誰でも編集可能な共有設定
function setPublicEditAccess(spreadsheet) {
  try {
    const file = DriveApp.getFileById(spreadsheet.getId());
    
    // リンクを知っている全員が編集可能に設定
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.EDIT);
    
    console.log('共有設定完了: URLを知っている人は誰でも編集可能');
    return true;
  } catch (error) {
    console.error('共有設定エラー:', error);
    return false;
  }
}

// 共有設定処理（旧関数、互換性のため残す）
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
  const subject = `[業務記述書完成] ${workSpec.title}`;
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

// HTML通知作成（UTF-8エンコーディング対応）
function buildHtmlNotification(workSpec, sheetUrl) {
  return `
    <!DOCTYPE html>
    <html>
    <head>
      <meta charset="UTF-8">
      <meta name="viewport" content="width=device-width, initial-scale=1.0">
    </head>
    <body>
    <div style="font-family: 'Noto Sans JP', Arial, sans-serif; max-width: 600px; margin: 0 auto;">
      <div style="background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; padding: 20px; text-align: center; border-radius: 10px 10px 0 0;">
        <h1 style="margin: 0; font-size: 24px;">業務記述書が完成しました</h1>
      </div>
      
      <div style="padding: 20px; background-color: #f8f9fa; border: 1px solid #e9ecef; border-top: none;">
        <h2 style="color: #495057; margin-top: 0;">${workSpec.title}</h2>
        <p style="font-size: 16px; line-height: 1.5; color: #6c757d;">${workSpec.summary}</p>
        
        <div style="background-color: white; padding: 15px; border-radius: 8px; margin: 20px 0; border-left: 4px solid #28a745;">
          <h3 style="margin-top: 0; color: #28a745;">スプレッドシート（編集可能）</h3>
          <p style="margin-bottom: 10px;">以下のリンクから業務記述書とタスク管理シートにアクセス・編集できます：</p>
          <a href="${sheetUrl}" style="display: inline-block; background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; padding: 12px 24px; text-decoration: none; border-radius: 4px; font-weight: bold;">スプレッドシートを開く</a>
        </div>
        
        ${workSpec.timeline && workSpec.timeline.length > 0 ? `
        <div style="background-color: white; padding: 15px; border-radius: 8px; margin: 20px 0; border-left: 4px solid #ffc107;">
          <h3 style="margin-top: 0; color: #ff9800;">主要マイルストーン</h3>
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
          <h3 style="margin-top: 0; color: #dc3545;">重要な注意事項</h3>
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
    </body>
    </html>
  `;
}

// エラー通知メール送信
function sendErrorNotificationEmail(to, originalSubject, errorMessage) {
  const subject = `[処理エラー] ${originalSubject}`;
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
    const prefixedSubject = subject.startsWith('[task]') ? subject : `[task] ${subject}`;
    
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
            <small style="color: #666; font-size: 12px;">※ [task] プレフィックスは自動付与されます</small>
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
  
  // GPT-5へのアップグレード提案（初回のみ）
  const hasShownGPT5 = PropertiesService.getDocumentProperties().getProperty('GPT5_UPGRADE_SHOWN');
  const currentModel = getConfig('OPENAI_MODEL');
  if (!hasShownGPT5 && currentModel !== 'gpt-5') {
    PropertiesService.getDocumentProperties().setProperty('GPT5_UPGRADE_SHOWN', 'true');
    // 少し遅延させて実行
    Utilities.sleep(1000);
    upgradeToGPT5();
  }
  
  ui.createMenu('📋 タスク管理システム')
    .addSubMenu(ui.createMenu('⚙️ システム')
      .addItem('🚀 初回セットアップ', 'setupSystem')
      .addItem('🔧 設定を開く', 'openConfigSheet')
      .addSeparator()
      .addItem('🔑 APIキーを設定', 'setApiKey')
      .addItem('🤖 AIモデルを選択', 'selectOpenAIModel')
      .addItem('⏰ トリガーを設定', 'setupTriggers')
      .addItem('🗑️ トリガーを削除', 'deleteTriggers'))
    .addSubMenu(ui.createMenu('📧 メール')
      .addItem('✉️ 業務メール作成', 'showEmailComposer')
      .addItem('📥 新着メール処理を今すぐ実行', 'processNewEmailsManually')
      .addSeparator()
      .addItem('🔧 検索クエリを修正', 'fixProcessingQuery')
      .addItem('🔍 メール検索テスト', 'testEmailSearch')
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
      console.log('✅ APIキーを設定しました。');
      logActivity('API_KEY', 'API key configured');
    } else {
      console.warn('⚠️ APIキーが入力されていません。');
    }
  }
}

// モデル選択ダイアログ
function selectOpenAIModel() {
  const ui = SpreadsheetApp.getUi();
  const currentModel = getConfig('OPENAI_MODEL') || 'gpt-5';
  
  const htmlContent = `
    <div style="font-family: Arial, sans-serif; padding: 20px;">
      <h3>OpenAIモデル選択</h3>
      <p>現在のモデル: <strong>${currentModel}</strong></p>
      <br>
      <label style="background-color: #fff3e0; padding: 5px; border-radius: 5px;">
        <input type="radio" name="model" value="o3-deep-research" ${currentModel === 'o3-deep-research' ? 'checked' : ''}>
        <strong>o3-deep-research</strong> 🔬 (最高精度) - 深層分析、複雑な推論
      </label><br><br>
      <label style="background-color: #e8f5e9; padding: 5px; border-radius: 5px;">
        <input type="radio" name="model" value="gpt-5" ${currentModel === 'gpt-5' ? 'checked' : ''}>
        <strong>gpt-5</strong> 🆕 (推奨) - 最新技術、高速・高性能
      </label><br><br>
      <label>
        <input type="radio" name="model" value="gpt-4o" ${currentModel === 'gpt-4o' ? 'checked' : ''}>
        <strong>gpt-4o</strong> - バランス型
      </label><br><br>
      <label>
        <input type="radio" name="model" value="gpt-4-turbo" ${currentModel === 'gpt-4-turbo' ? 'checked' : ''}>
        <strong>gpt-4-turbo</strong> - 詳細な分析
      </label><br><br>
      <label>
        <input type="radio" name="model" value="gpt-3.5-turbo" ${currentModel === 'gpt-3.5-turbo' ? 'checked' : ''}>
        <strong>gpt-3.5-turbo</strong> - 高速・低コスト
      </label><br><br>
      <hr>
      <p style="color: #666; font-size: 12px;">
        ※ o3-deep-researchは最も詳細な分析が可能ですが、処理に時間がかかる場合があります
      </p>
      <br>
      <button onclick="google.script.run.updateOpenAIModel(document.querySelector('input[name=model]:checked').value); google.script.host.close();">
        保存
      </button>
      <button onclick="google.script.host.close();">キャンセル</button>
    </div>
  `;
  
  const html = HtmlService.createHtmlOutput(htmlContent)
    .setWidth(450)
    .setHeight(420);
  
  ui.showModalDialog(html, 'AIモデル選択');
}

// モデル更新
function updateOpenAIModel(model) {
  setConfig('OPENAI_MODEL', model);
  const ui = SpreadsheetApp.getUi();
  
  // モデル別のメッセージ
  let message = `AIモデルを ${model} に変更しました。`;
  if (model === 'gpt-5') {
    message += '\n\n🎉 最新のgpt-5モデルを使用します。より高精度な業務分析と法令遵守チェックが可能になります。';
  }
  
  ui.alert('設定完了', message, ui.ButtonSet.OK);
  logActivity('MODEL_CHANGED', `OpenAI model changed to: ${model}`);
}

// gpt-5への自動アップグレード（既存ユーザー向け）
function upgradeToGPT5() {
  const currentModel = getConfig('OPENAI_MODEL');
  const ui = SpreadsheetApp.getUi();
  
  if (currentModel !== 'gpt-5') {
    const response = ui.alert(
      '🎉 GPT-5が利用可能です！',
      '最新のGPT-5モデルが利用可能になりました。\n\n' +
      '【GPT-5の特徴】\n' +
      '• より深い業務理解と分析力\n' +
      '• 法令・規制の最新情報に対応\n' +
      '• MECEな構造でのタスク分解精度向上\n' +
      '• 内部統制とリスク管理の詳細化\n\n' +
      'GPT-5にアップグレードしますか？',
      ui.ButtonSet.YES_NO
    );
    
    if (response === ui.Button.YES) {
      setConfig('OPENAI_MODEL', 'gpt-5');
      ui.alert('アップグレード完了', 'GPT-5モデルに切り替えました。より高精度な分析が可能になります。', ui.ButtonSet.OK);
      logActivity('MODEL_UPGRADE', 'Upgraded to GPT-5');
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
  try {
    console.log('📥 手動メール処理開始');
    processNewEmails();
    console.log('✅ メール処理完了');
  } catch (e) {
    console.error('❌ メール処理エラー:', e.toString());
    throw e; // エラーをコンソールに表示
  }
}

// Config シートを開く
function openConfigSheet() {
  const sheet = ss().getSheetByName(CONFIG_SHEET);
  if (sheet) {
    ss().setActiveSheet(sheet);
  } else {
    console.error('Config シートが見つかりません。');
  }
}

// 処理済みラベル作成
function createProcessedLabel() {
  try {
    let label = GmailApp.getUserLabelByName('PROCESSED');
    if (!label) {
      label = GmailApp.createLabel('PROCESSED');
      console.log('✅ PROCESSEDラベルを作成しました。');
    } else {
      console.log('ℹ️ PROCESSEDラベルは既に存在します。');
    }
  } catch (e) {
    console.error('❌ エラー：', e);
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
      console.log('✅ フローシートをリセットしました。');
    }
  }
}

// 処理統計表示
function showProcessingStats() {
  const inboxSheet = ss().getSheetByName(INBOX_SHEET);
  if (!inboxSheet || inboxSheet.getLastRow() <= 1) {
    console.warn('処理データがありません。');
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
  let sheet = ss().getSheetByName(ACTIVITY_LOG_SHEET);
  
  // シートが存在しない場合は作成
  if (!sheet) {
    console.log('ActivityLogシートが存在しないため作成します');
    sheet = createActivityLogSheet();
  }
  
  if (sheet) {
    // 隠されている場合は表示
    if (sheet.isSheetHidden()) {
      sheet.showSheet();
    }
    ss().setActiveSheet(sheet);
    const ui = SpreadsheetApp.getUi();
    ui.alert('アクティビティログ', 'アクティビティログを表示しました。', ui.ButtonSet.OK);
  } else {
    console.error('アクティビティログシートの作成に失敗しました');
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
    `- 件名に [task] を含むメールを自動処理\n` +
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
const OPENAI_URL_RESPONSES = 'https://api.openai.com/v1/responses'; // o3-deep-research用（将来的に対応）

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
  return `あなたは日本の上場企業（東証プライム市場）において、プロジェクトマネジメント、法務、内部統制、リスク管理の実務経験を20年以上持つ専門家です。

【あなたの専門性】
- 金融商品取引法、会社法、J-SOX法に精通
- 内部統制報告制度の構築・運用経験豊富
- ISO9001/27001、プライバシーマーク認証取得支援経験
- 監査法人対応、コーポレートガバナンス・コード対応の実績多数
- PMBOK、COBIT、COSOフレームワークの実装経験

【作成方針】
MECE（Mutually Exclusive, Collectively Exhaustive）の原則に基づき、以下の観点で業務を詳細に分解してください：

1. 法令・規制の具体的対応
   - 金融商品取引法（開示規制、内部統制報告制度）
   - 会社法（取締役会規程、監査役監査基準）
   - 個人情報保護法（プライバシーポリシー、同意取得）
   - 下請法、独占禁止法、労働基準法など関連法令
   - 業界特有の規制（金融業：銀行法、製造業：PL法など）
   - 各法令の具体的な条文番号まで特定
   - 違反時の罰則規定と影響範囲を明記

2. 内部統制・リスク管理の詳細設計
   - 3点セット（業務記述書、フローチャート、RCM）の作成
   - キーコントロールの特定と評価手続き
   - IT全般統制（ITGC）とIT業務処理統制（ITAC）
   - 職務分離（SoD）マトリックスの設計
   - 不正のトライアングル理論に基づく予防的統制
   - リスクアセスメント（発生可能性×影響度）
   - BCP/DR計画との連携

3. 具体的なアクション分解（MECE原則）
   - WBSレベル3以上の詳細度で作業を分解
   - 各タスクを15分～2時間単位の作業に細分化
   - 前工程・後工程の依存関係を明確化
   - クリティカルパスの特定
   - バッファ時間の設定（リスク対応）
   - チェックポイント、承認ポイントの明示
   - エスカレーションパスの定義

4. 実務的な具体例の提示
   - 使用する具体的な文書テンプレート名
   - 参照すべき社内規程・マニュアル名
   - 利用システム・ツールの具体名（SAP、Salesforce等）
   - 承認フロー（稟議システムの承認ルート）
   - 監査証跡の取得方法
   - KPIの計算式と測定頻度

5. ステークホルダー管理
   - RACIマトリックス（Responsible/Accountable/Consulted/Informed）
   - コミュニケーション計画（頻度、手段、参加者）
   - 報告書フォーマット（取締役会、監査役会向け）
   - 外部機関対応（監査法人、規制当局、証券取引所）

6. 上場企業特有の考慮事項（上場企業の場合に適用）
   - 東京証券取引所との関係：適時開示、コーポレートガバナンス報告書の提出、株式事務など
   - 金融庁・財務局との関係：有価証券報告書、内部統制報告書の提出、検査対応
   - 開示に関する観点：適時開示規則、インサイダー取引防止、IR活動の実施
   - これらの観点を業務フローに組み込み、必要なタスクとして明示
   - 入力に「株主総会」や「取締役会」などのキーワードが含まれる場合、開示義務（適時開示、法定開示）に関連するかを評価し、関連する場合、開示手続き、内部統制、監査対応などのタスクを追加

【重要な指示】
- 抽象的な表現を避け、実行可能な具体的アクションを記載
- 「検討する」→「〇〇の基準に基づき△△を評価し、□□を決定する」
- 「確認する」→「〇〇チェックリストの全項目が基準値を満たすことを確認する」
- 「管理する」→「〇〇管理台帳に記録し、週次で△△指標をモニタリングする」
- すべての法的記載には「※最終的には顧問弁護士・専門家による確認が必要」を付記

出力は日本語で、JSON Schema準拠。法的助言の代替ではないことを明記。`;
}

// ユーザープロンプト構築
function buildUserPrompt(mailBody, orgProfileJson) {
  const orgProfile = orgProfileJson ? JSON.parse(orgProfileJson) : {};
  
  return `以下の業務内容をMECE原則に基づき、実行可能な詳細タスクに分解して業務記述書とフロー表を作成してください。

【業務内容】
${mailBody}

【組織プロフィール】
- 上場区分: ${orgProfile.listing || '東証プライム'}
- 業種: ${orgProfile.industry || '製造業'}
- 対象地域: ${(orgProfile.jurisdictions || ['JP']).join(', ')}
- 社内基準: ${(orgProfile.policies || ['J-SOX対応', 'ISO27001認証']).join(', ')}
- 従業員数: ${orgProfile.employees || '1000名以上'}
- 売上規模: ${orgProfile.revenue || '100億円以上'}

【詳細化の要求水準】

1. タスクの粒度
   - 各タスクは最大2時間で完了可能な単位に分解
   - 具体的な成果物・アウトプットを明記
   - 判断基準・チェックポイントを数値化
   例：「確認する」→「〇〇チェックリスト25項目中23項目以上が基準値80%を超えることを確認」

2. 法令・規制の具体的記載
   - 法令名と該当条文を明記（例：金融商品取引法第24条）
   - 違反時のペナルティを記載（例：5年以下の懲役又は500万円以下の罰金）
   - 監督官庁への届出期限（例：変更後2週間以内に関東財務局へ届出）
   - 業界ガイドライン（例：日本証券業協会自主規制規則第〇条）

3. 内部統制の具体的設計
   - 予防的統制：承認権限規程（例：100万円以上は部長決裁）
   - 発見的統制：月次照合作業（例：売掛金残高と補助簿の照合）
   - IT統制：アクセスログの定期レビュー（例：特権IDの使用記録を週次確認）
   - 証跡保存：7年間の文書保存（電子帳簿保存法準拠）

4. リスク対応の詳細
   - リスクシナリオ：具体的な事象と発生確率（H/M/L）
   - 影響額：定量評価（売上の〇%相当）
   - 対応策：予防策、発生時対応、復旧計画
   - モニタリング指標：KRI（Key Risk Indicator）の設定

5. 実務ツール・システム
   - ERP：SAP S/4HANA（モジュール名まで特定）
   - ワークフロー：ServiceNow（申請フォームID）
   - 文書管理：SharePoint（フォルダ構成）
   - コミュニケーション：Teams（チャネル名）

【flow_rows作成の詳細要求】
各行は以下の粒度で記載：

- 工程：WBSレベル2（例：「1.2 要件定義フェーズ」）
- 実施タイミング：具体的な日付・期間（例：「2024年4月1日～4月15日（10営業日）」）
- 部署：正式部署名と人数（例：「経理部決算チーム（5名）」）
- 担当役割：RACI形式（例：「R:主任、A:課長、C:部長、I:監査役」）
- 作業内容：5W1H形式の具体的記述
  例：「売掛金年齢表を作成し、90日超の債権リスト（想定20件）を抽出。
       各債権について営業担当者へヒアリング（1件30分）を実施し、
       回収可能性を5段階評価。評価結果を貸倒引当金算定表に反映。」
- 条件分岐：判断基準を数値化（例：「売上高1000万円以上の場合は役員承認ルートへ」）
- 利用ツール：バージョンまで特定（例：「Excel 2021 売掛金管理テンプレートv3.2」）
- URLリンク：具体的な参照先（例：「社内ポータル/規程集/与信管理規程」）
- 備考：注意事項、過去の失敗事例、改善提案など
- 法令・規制：該当する具体的な法令・条文・ガイドライン
- 内部統制：統制活動の種類とコントロール番号（例：「予防的統制 CC-AR-001」）
- コンプライアンス留意点：過去の違反事例、監査指摘事項、業界のベストプラクティス

【最低限含めるべきタスク数】
- メインプロセス：最低15タスク
- 各タスクのサブタスク：3～5個
- チェックポイント：各フェーズに2箇所以上
- 承認ポイント：重要な意思決定箇所すべて

重要：抽象的な表現は使用禁止。すべて測定可能・実行可能な具体的記述にすること。`;
}

// OpenAI API呼び出し
function callOpenAI(mailBody, orgProfileJson) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
  if (!apiKey) {
    throw new Error('OpenAI APIキーが設定されていません。スクリプトプロパティに OPENAI_API_KEY を設定してください。');
  }
  
  const modelName = getConfig('OPENAI_MODEL') || 'gpt-4o';
  const schema = buildWorkSpecSchema();
  
  // gpt-5やo3-deep-researchの場合はv1/responsesエンドポイントを使用
  const useResponsesEndpoint = ['gpt-5', 'o3-deep-research'].includes(modelName);
  
  if (useResponsesEndpoint) {
    return callOpenAIResponses(mailBody, orgProfileJson, apiKey, modelName, schema);
  }
  
  // 通常のチャットモデル（gpt-4o, gpt-4-turbo, gpt-3.5-turbo）
  const messages = [
    { role: 'system', content: buildSystemPrompt() },
    { role: 'user', content: buildUserPrompt(mailBody, orgProfileJson) }
  ];
  
  const payload = {
    model: modelName,
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
  
  logActivity('OPENAI_CALL', `Calling OpenAI Chat API with model: ${modelName}`);
  
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
    if (status === 404) {
      const errorBody = res.getContentText();
      if (errorBody.includes('v1/responses') || errorBody.includes('This model is only supported')) {
        // このモデルはv1/responsesエンドポイントを使用する必要がある
        logActivity('ENDPOINT_SWITCH', `Model ${modelName} requires v1/responses endpoint, switching...`);
        return callOpenAIResponses(mailBody, orgProfileJson, apiKey, modelName, schema);
      }
    }
    
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

// v1/responsesエンドポイント用のAPI呼び出し（GPT-5, o3-deep-research用）
function callOpenAIResponses(mailBody, orgProfileJson, apiKey, modelName, schema) {
  const requestId = `req_${new Date().getTime()}_${Math.random().toString(36).substr(2, 9)}`;
  logActivity('OPENAI_CALL', `Calling OpenAI Responses API with model: ${modelName}, Request ID: ${requestId}`);

  // システムプロンプトとユーザープロンプトを結合
  const systemPrompt = buildSystemPrompt();
  const userPrompt = buildUserPrompt(mailBody, orgProfileJson);

  // JSON出力を明示的に指示
  const enhancedPrompt = `${systemPrompt}\n\n${userPrompt}\n\n重要: 必ず有効なJSONフォーマットで出力してください。出力は以下のJSONスキーマに厳密に従ってください。追加のテキストや説明は一切含めないでください。\nJSON Schema: ${JSON.stringify(schema, null, 2)}`;

  let payload;
  let url = OPENAI_URL_CHAT; // Use chat completions endpoint
  let effectiveModel = modelName;

  if (modelName === 'o3-deep-research') {
    effectiveModel = 'o1-preview'; // Map to actual model
    payload = {
      model: effectiveModel,
      messages: [
        { role: 'user', content: enhancedPrompt }
      ],
      max_tokens: 8000 // Supported parameter
      // Omit temperature, response_format as they are not supported
    };
  } else if (modelName === 'gpt-5') {
    // Existing logic for gpt-5, but since it's not real, fallback or handle similarly
    effectiveModel = 'gpt-4o'; // Fallback
    payload = {
      model: effectiveModel,
      messages: [
        { role: 'user', content: enhancedPrompt }
      ],
      max_tokens: 8000
    };
  } else {
    // Fallback for other models
    payload = {
      model: modelName,
      messages: [
        { role: 'user', content: enhancedPrompt }
      ],
      max_tokens: 8000
    };
  }

  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'Authorization': `Bearer ${apiKey}`,
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const response = retryWithBackoff(() => {
    const res = UrlFetchApp.fetch(url, options);
    const status = res.getResponseCode();
    
    if (status === 200) {
      return res;
    }
    
    if (status === 429) {
      const retryAfter = res.getHeaders()['Retry-After'] || '60';
      logActivity('OPENAI_RATE_LIMIT', `Rate limited. Retry after ${retryAfter} seconds`);
      throw new Error(`Rate limit exceeded. Retry after ${retryAfter} seconds`);
    }
    
    if (status >= 500) {
      logActivity('OPENAI_SERVER_ERROR', `Server error: ${status}`);
      throw new Error(`OpenAI server error: ${status}`);
    }
    
    if (status === 404) {
      // モデルが利用できない場合はgpt-4oにフォールバック
      const errorBody = res.getContentText();
      logActivity('OPENAI_MODEL_ERROR', `Model ${effectiveModel} not available: ${errorBody}`);
      
      // gpt-4oにフォールバック
      setConfig('OPENAI_MODEL', 'gpt-4o');
      logActivity('MODEL_FALLBACK', 'Falling back to gpt-4o');
      return callOpenAI(mailBody, orgProfileJson);  // 再帰的に呼び出し
    }
    
    const errorBody = res.getContentText();
    const errorDetails = JSON.stringify(payload, null, 2);
    logActivity('OPENAI_ERROR', `Status: ${status}, Body: ${errorBody}`);
    logActivity('OPENAI_PAYLOAD', `Failed payload: ${errorDetails}`);
    
    // より詳細なエラーメッセージ
    console.error('OpenAI API Error:', {
      status: status,
      error: errorBody,
      model: effectiveModel,
      endpoint: url
    });
    
    throw new Error(`OpenAI API error ${status}: ${errorBody}`);
  });
  
  // フォールバックの場合の型チェック
  if (response && typeof response.getContentText !== 'function') {
    // 既に解析済みのJSONの場合
    logActivity('FALLBACK_RESPONSE', 'Using fallback response');
    return response;
  }
  
  const responseData = JSON.parse(response.getContentText());
  
  logActivity('OPENAI_SUCCESS', 'Successfully received response from OpenAI Responses API');
  
  // JSONとしてパースを試みる
  try {
    return JSON.parse(responseData.choices[0].message.content);
  } catch (e) {
    // 既にオブジェクトの場合はそのまま返す
    if (typeof responseData.choices[0].message.content === 'object') {
      return responseData.choices[0].message.content;
    }
    throw new Error('Failed to parse OpenAI response: ' + e.toString());
  }
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

// Config修正関数
function fixProcessingQuery() {
  const ui = SpreadsheetApp.getUi();
  const sh = ss().getSheetByName(CONFIG_SHEET);
  
  if (!sh) {
    ui.alert('エラー', 'Configシートが見つかりません。初回セットアップを実行してください。', ui.ButtonSet.OK);
    return;
  }
  
  // 現在の値を確認
  const currentQuery = getConfig('PROCESSING_QUERY');
  console.log('現在の検索クエリ:', currentQuery);
  
  // 古い値が含まれていたら修正
  if (currentQuery && (currentQuery.includes('WORK-REQ') || currentQuery.includes('label:inbox'))) {
    setConfig('PROCESSING_QUERY', 'subject:[task] is:unread');
    console.log('検索クエリを修正しました: subject:[task] is:unread');
    ui.alert('修正完了', '検索クエリを [task] 用に修正しました。\n新しいクエリ: subject:[task] is:unread', ui.ButtonSet.OK);
  } else if (!currentQuery) {
    setConfig('PROCESSING_QUERY', 'subject:[task] is:unread');
    console.log('検索クエリを設定しました: subject:[task] is:unread');
    ui.alert('設定完了', '検索クエリを設定しました。\nクエリ: subject:[task] is:unread', ui.ButtonSet.OK);
  } else {
    console.log('検索クエリは既に正しく設定されています');
    ui.alert('確認', `現在の検索クエリ:\n${currentQuery}\n\n既に正しく設定されています。`, ui.ButtonSet.OK);
  }
}

// メール検索テスト関数
function testEmailSearch() {
  console.log('===== メール検索テスト開始 =====');
  
  // 先に設定を修正
  fixProcessingQuery();
  
  // 1. 現在の設定を確認
  const currentQuery = getConfig('PROCESSING_QUERY');
  console.log('修正後の検索クエリ:', currentQuery);
  
  // 2. 様々なパターンで検索をテスト
  const testQueries = [
    'subject:[task]',
    'subject:"[task]"',
    'subject:task',
    '[task]',
    'is:unread subject:[task]',
    'is:unread subject:task',
    'is:unread',
    'in:anywhere [task]',
    'in:anywhere subject:task'
  ];
  
  console.log('\n各検索パターンの結果:');
  testQueries.forEach(query => {
    try {
      const threads = GmailApp.search(query, 0, 5);
      console.log(`  "${query}": ${threads.length}件`);
      
      if (threads.length > 0 && query.includes('task')) {
        // 最初のメールの件名を表示
        const firstMessage = threads[0].getMessages()[0];
        console.log(`    例: "${firstMessage.getSubject()}"`);
      }
    } catch (e) {
      console.log(`  "${query}": エラー - ${e.toString()}`);
    }
  });
  
  // 3. 未読メール全体から[task]を含むものを探す
  console.log('\n未読メールから[task]を含むものを検索:');
  const allUnread = GmailApp.search('is:unread', 0, 20);
  console.log(`未読メール総数: ${allUnread.length}件`);
  
  let taskCount = 0;
  allUnread.forEach(thread => {
    const firstMessage = thread.getMessages()[0];
    const subject = firstMessage.getSubject();
    
    // 様々なパターンでマッチング
    if (subject.includes('[task]') || 
        subject.toLowerCase().includes('[task]') ||
        subject.includes('task') ||
        subject.includes('【task】')) {
      taskCount++;
      console.log(`  ✓ "${subject}"`);
    }
  });
  
  console.log(`\n[task]関連メール: ${taskCount}件`);
  
  // 4. 推奨クエリの提案
  console.log('\n推奨される検索クエリ:');
  if (taskCount > 0) {
    console.log('  ・ "subject:task is:unread" - より広範囲にマッチ');
    console.log('  ・ "is:unread [task]" - 件名と本文から検索');
  }
  
  console.log('\n===== テスト完了 =====');
  
  // 結果をUIに表示
  const ui = SpreadsheetApp.getUi();
  ui.alert('検索テスト結果', 
    `現在のクエリ: ${currentQuery}\n` +
    `未読メール: ${allUnread.length}件\n` +
    `[task]関連: ${taskCount}件\n\n` +
    '詳細はログを確認してください', 
    ui.ButtonSet.OK);
}

// データ解析・書き込み処理

// フローシートのヘッダー定義（法令・内部統制の観点を追加）
const FLOW_HEADERS = [
  '工程', 
  '実施タイミング', 
  '部署', 
  '担当役割', 
  '作業内容', 
  '条件分岐', 
  '利用ツール', 
  'URLリンク', 
  '備考',
  '法令・規制',
  '内部統制',
  'コンプライアンス留意点'
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
  
  // 列幅を調整（法令・内部統制の列を追加）
  sh.setColumnWidth(1, 120); // 工程
  sh.setColumnWidth(2, 150); // 実施タイミング
  sh.setColumnWidth(3, 120); // 部署
  sh.setColumnWidth(4, 150); // 担当役割
  sh.setColumnWidth(5, 300); // 作業内容
  sh.setColumnWidth(6, 150); // 条件分岐
  sh.setColumnWidth(7, 150); // 利用ツール
  sh.setColumnWidth(8, 200); // URLリンク
  sh.setColumnWidth(9, 200); // 備考
  sh.setColumnWidth(10, 250); // 法令・規制
  sh.setColumnWidth(11, 250); // 内部統制
  sh.setColumnWidth(12, 300); // コンプライアンス留意点
  
  return sh;
}

// フローシート作成
function createFlowSheet(sheetName) {
  const sh = ss().insertSheet(sheetName);
  
  sh.getRange(1, 1, 1, FLOW_HEADERS.length).setValues([FLOW_HEADERS]);
  sh.getRange(1, 1, 1, FLOW_HEADERS.length).setFontWeight('bold');
  sh.getRange(1, 1, 1, FLOW_HEADERS.length).setBackground('#e8f5e9');
  sh.setFrozenRows(1);
  
  // 列幅調整（法令・内部統制の列を追加）
  sh.setColumnWidth(1, 100); // 工程
  sh.setColumnWidth(2, 120); // 実施タイミング
  sh.setColumnWidth(3, 100); // 部署
  sh.setColumnWidth(4, 100); // 担当役割
  sh.setColumnWidth(5, 250); // 作業内容
  sh.setColumnWidth(6, 150); // 条件分岐
  sh.setColumnWidth(7, 120); // 利用ツール
  sh.setColumnWidth(8, 150); // URLリンク
  sh.setColumnWidth(9, 200); // 備考
  sh.setColumnWidth(10, 200); // 法令・規制
  sh.setColumnWidth(11, 200); // 内部統制
  sh.setColumnWidth(12, 250); // コンプライアンス留意点
  
  return sh;
}

// ビジュアルフロー生成関数
function generateVisualFlow() {
  try {
    console.log('=== ビジュアルフロー生成開始 ===');
    
    const flowSheetName = getConfig('FLOW_SHEET_NAME') || FLOW_SHEET;
    const sheet = ss().getSheetByName(flowSheetName);
    
    if (!sheet) {
      console.error('フローシートが見つかりません');
      return;
    }
    
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) {
      console.error('データがありません');
      return;
    }
    
    // ビジュアルシートの準備
    const visualSheetName = 'ビジュアルフロー';
    let visualSheet = ss().getSheetByName(visualSheetName);
    if (!visualSheet) {
      visualSheet = ss().insertSheet(visualSheetName);
    }
    visualSheet.clear();
    
    // データの解析
    const flowData = parseFlowDataForVisual(data);
    
    // フローチャートの描画
    drawVisualFlowChart(visualSheet, flowData);
    
    console.log('=== ビジュアルフロー生成完了 ===');
    
  } catch (error) {
    console.error('ビジュアルフロー生成エラー:', error);
  }
}

// ビジュアル用フローデータの解析
function parseFlowDataForVisual(data) {
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
    processName: ''
  };
  
  // プロセス名の取得
  if (data.length > 1 && data[1][columnIndex['工程']]) {
    flowData.processName = data[1][columnIndex['工程']];
  }
  
  // データの整理
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[columnIndex['工程']] || row[columnIndex['工程']] === '') continue;
    
    const dept = row[columnIndex['部署']] || 'その他';
    const timing = row[columnIndex['実施タイミング']] || '';
    const tool = row[columnIndex['利用ツール']] || '';
    const url = row[columnIndex['URLリンク']] || '';
    
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
      task: row[columnIndex['作業内容']] || '',
      role: row[columnIndex['担当役割']] || '',
      condition: row[columnIndex['条件分岐']] || '',
      tool: tool,
      url: url,
      note: row[columnIndex['備考']] || '',
      legal: row[columnIndex['法令・規制']] || '',
      control: row[columnIndex['内部統制']] || '',
      compliance: row[columnIndex['コンプライアンス留意点']] || ''
    });
    
    // ツールの収集
    if (tool && tool !== '-' && tool !== 'なし') {
      const tools = tool.split(/[／、,]/);
      tools.forEach(t => {
        const trimmedTool = t.trim();
        if (trimmedTool) {
          flowData.tools.add(trimmedTool);
        }
      });
    }
    
    // URLをタイミングごとに管理
    if (url && url !== '-' && url !== 'なし') {
      if (!flowData.datasources[timing]) {
        flowData.datasources[timing] = [];
      }
      if (!flowData.datasources[timing].includes(url)) {
        flowData.datasources[timing].push(url);
      }
    }
  }
  
  return flowData;
}

// 配列データのフォーマット
function formatArray(arr) {
  if (!arr || !Array.isArray(arr)) return '';
  return arr.filter(item => item).join('\n');
}

// リスクデータのフォーマット
function formatRisks(risks) {
  if (!risks || !Array.isArray(risks)) return '';
  
  return risks.map(risk => {
    if (typeof risk === 'string') {
      return risk;
    } else if (typeof risk === 'object' && risk !== null) {
      const parts = [];
      if (risk.risk) parts.push(`リスク: ${risk.risk}`);
      if (risk.mitigation) parts.push(`対策: ${risk.mitigation}`);
      if (risk.probability) parts.push(`確率: ${risk.probability}`);
      if (risk.impact) parts.push(`影響: ${risk.impact}`);
      return parts.join(' / ');
    }
    return '';
  }).filter(item => item).join('\n');
}

// ビジュアルフローチャートの描画
function drawVisualFlowChart(sheet, flowData) {
  // カラーパレット
  const COLORS = {
    HEADER: '#4A5568',
    TIMELINE: '#F7FAFC',
    PROCESS: '#87CEEB',
    DECISION: '#FFD700',
    START_END: '#90EE90',
    EMPTY: '#FAFAFA',
    DATASOURCE: '#E3F2FD'
  };
  
  let currentRow = 1;
  
  // タイトル行
  const flowTitle = flowData.processName || '業務フロー';
  const maxCols = Math.max(flowData.departmentList.length + 2, 10);
  sheet.getRange(currentRow, 1, 1, maxCols).merge();
  const titleCell = sheet.getRange(currentRow, 1);
  titleCell.setValue(flowTitle);
  titleCell.setBackground(COLORS.HEADER);
  titleCell.setFontColor('#FFFFFF');
  titleCell.setFontSize(16);
  titleCell.setFontWeight('bold');
  titleCell.setHorizontalAlignment('center');
  titleCell.setVerticalAlignment('middle');
  sheet.setRowHeight(currentRow, 50);
  currentRow++;
  
  // ヘッダー行（タイムライン + 部署）
  sheet.getRange(currentRow, 1).setValue('タイミング');
  sheet.getRange(currentRow, 1).setBackground(COLORS.TIMELINE);
  sheet.getRange(currentRow, 1).setFontWeight('bold');
  
  flowData.departmentList.forEach((dept, index) => {
    sheet.getRange(currentRow, index + 2).setValue(dept);
    sheet.getRange(currentRow, index + 2).setBackground('#E8EAF6');
    sheet.getRange(currentRow, index + 2).setFontWeight('bold');
    sheet.getRange(currentRow, index + 2).setHorizontalAlignment('center');
  });
  
  if (Object.keys(flowData.datasources).length > 0) {
    const dataCol = flowData.departmentList.length + 2;
    sheet.getRange(currentRow, dataCol).setValue('関連資料');
    sheet.getRange(currentRow, dataCol).setBackground('#E3F2FD');
    sheet.getRange(currentRow, dataCol).setFontWeight('bold');
  }
  
  sheet.setRowHeight(currentRow, 40);
  currentRow++;
  
  // 開始行
  sheet.getRange(currentRow, 1).setValue('【開始】');
  sheet.getRange(currentRow, 1).setBackground(COLORS.START_END);
  sheet.getRange(currentRow, 1).setFontWeight('bold');
  currentRow++;
  
  // 各タイミングの行
  flowData.timings.forEach((timing) => {
    sheet.getRange(currentRow, 1).setValue(timing);
    sheet.getRange(currentRow, 1).setBackground(COLORS.TIMELINE);
    sheet.getRange(currentRow, 1).setFontWeight('bold');
    sheet.getRange(currentRow, 1).setVerticalAlignment('middle');
    
    // 各部署のタスク
    flowData.departmentList.forEach((dept, deptIndex) => {
      const col = deptIndex + 2;
      if (flowData.departments[dept] && flowData.departments[dept][timing]) {
        const tasks = flowData.departments[dept][timing];
        const taskTexts = tasks.map(t => {
          let text = t.task;
          if (t.role) text = `[${t.role}] ${text}`;
          if (t.condition && t.condition !== 'なし') text = `◆ ${text}`;
          return text;
        });
        
        const cell = sheet.getRange(currentRow, col);
        cell.setValue(taskTexts.join('\n'));
        
        // 条件分岐があるかチェック
        const hasCondition = tasks.some(t => t.condition && t.condition !== 'なし');
        cell.setBackground(hasCondition ? COLORS.DECISION : COLORS.PROCESS);
        
        cell.setWrap(true);
        cell.setVerticalAlignment('top');
        cell.setBorder(true, true, true, true, false, false);
        
        // ツール情報と法令・内部統制情報をノートに追加
        const noteItems = [];
        const tools = tasks.map(t => t.tool).filter(t => t && t !== 'なし').join(', ');
        if (tools) {
          noteItems.push(`使用ツール: ${tools}`);
        }
        
        const legals = tasks.map(t => t.legal).filter(t => t && t !== 'なし');
        if (legals.length > 0) {
          noteItems.push(`法令・規制: ${[...new Set(legals)].join(', ')}`);
        }
        
        const controls = tasks.map(t => t.control).filter(t => t && t !== 'なし');
        if (controls.length > 0) {
          noteItems.push(`内部統制: ${[...new Set(controls)].join(', ')}`);
        }
        
        const compliances = tasks.map(t => t.compliance).filter(t => t && t !== '特になし');
        if (compliances.length > 0) {
          noteItems.push(`留意点: ${compliances[0]}`); // 最初の留意点のみ表示
        }
        
        if (noteItems.length > 0) {
          cell.setNote(noteItems.join('\n\n'));
        }
      } else {
        sheet.getRange(currentRow, col).setBackground(COLORS.EMPTY);
      }
    });
    
    // データソース列
    if (Object.keys(flowData.datasources).length > 0) {
      const dataCol = flowData.departmentList.length + 2;
      if (flowData.datasources[timing] && flowData.datasources[timing].length > 0) {
        const urls = flowData.datasources[timing].join('\n');
        const dataCell = sheet.getRange(currentRow, dataCol);
        dataCell.setValue('📄 ' + urls);
        dataCell.setBackground(COLORS.DATASOURCE);
        dataCell.setWrap(true);
      }
    }
    
    sheet.setRowHeight(currentRow, 90);
    currentRow++;
  });
  
  // 終了行
  sheet.getRange(currentRow, 1).setValue('【完了】');
  sheet.getRange(currentRow, 1).setBackground(COLORS.START_END);
  sheet.getRange(currentRow, 1).setFontWeight('bold');
  
  const mergeCols = Object.keys(flowData.datasources).length > 0 ? 
                     flowData.departmentList.length + 1 : 
                     flowData.departmentList.length;
  sheet.getRange(currentRow, 2, 1, mergeCols).merge();
  const msgCell = sheet.getRange(currentRow, 2);
  msgCell.setValue('✅ プロセス完了');
  msgCell.setBackground('#E8F5E9');
  msgCell.setFontSize(14);
  msgCell.setFontWeight('bold');
  msgCell.setHorizontalAlignment('center');
  sheet.setRowHeight(currentRow, 50);
  currentRow++;
  
  // 凡例行
  sheet.getRange(currentRow, 1, 1, maxCols).merge();
  const legendCell = sheet.getRange(currentRow, 1);
  legendCell.setValue('凡例： □ 処理・作業　◆ 判断・分岐　📄 関連資料　※セルのノートに法令・内部統制・コンプライアンス情報があります');
  legendCell.setBackground(COLORS.TIMELINE);
  legendCell.setFontWeight('bold');
  sheet.setRowHeight(currentRow, 40);
  
  // 列幅の調整
  sheet.setColumnWidth(1, 150); // タイムライン列
  for (let i = 2; i <= flowData.departmentList.length + 1; i++) {
    sheet.setColumnWidth(i, 200); // 部署列
  }
  if (Object.keys(flowData.datasources).length > 0) {
    sheet.setColumnWidth(flowData.departmentList.length + 2, 150); // データソース列
  }
  
  // 全体に罫線を設定
  const range = sheet.getRange(1, 1, currentRow, maxCols);
  range.setBorder(true, true, true, true, true, true, '#d0d0d0', SpreadsheetApp.BorderStyle.SOLID);
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

// 法令・規制をチェックする関数（ガバナンス強化版）
function checkLegalRegulations(processName, workContent, timing, dept) {
  const regulations = [];
  
  // 外部専門家相談の必要性を判定
  const advisorsNeeded = determineRequiredAdvisors(processName + ' ' + workContent);
  if (advisorsNeeded.length > 0) {
    regulations.push('【要専門家相談】' + advisorsNeeded.map(a => a.type).join('、'));
  }
  
  // 開示要件チェック
  const disclosureCheck = checkDisclosureRequirement(processName + ' ' + workContent);
  if (disclosureCheck.requiresDisclosure) {
    regulations.push('【要開示】' + disclosureCheck.disclosureType.join('、'));
  }
  
  // 株主総会関連
  if (processName.includes('株主総会') || workContent.includes('株主総会')) {
    regulations.push('会社法（第295条〜第325条）');
    if (timing.includes('6月')) {
      regulations.push('定時株主総会（会社法第296条）');
    }
    if (workContent.includes('招集通知')) {
      regulations.push('招集通知期限（会社法第299条：2週間前）');
    }
    if (workContent.includes('議決権')) {
      regulations.push('議決権行使（会社法第308条〜第313条）');
    }
  }
  
  // 決算・開示関連
  if (processName.includes('決算') || workContent.includes('決算') || workContent.includes('開示')) {
    regulations.push('金融商品取引法');
    if (workContent.includes('四半期')) {
      regulations.push('四半期報告書（金商法第24条の4の7）45日以内');
    }
    if (workContent.includes('有価証券報告書')) {
      regulations.push('有価証券報告書（金商法第24条）3ヶ月以内');
    }
    if (workContent.includes('内部統制報告書')) {
      regulations.push('内部統制報告書（金商法第24条の4の4）');
    }
  }
  
  // 取締役会関連（ガバナンス強化）
  if (workContent.includes('取締役会') || dept.includes('取締役')) {
    regulations.push('会社法第362条（取締役会の権限）');
    if (workContent.includes('議事録')) {
      regulations.push('会社法第369条（取締役会議事録）');
    }
    // 重要事項の判定
    if (workContent.includes('重要') || workContent.includes('決議')) {
      regulations.push('【重要決議】東証への適時開示検討');
      regulations.push('【専門家相談】法律事務所への事前確認推奨');
    }
  }
  
  // 監査関連
  if (workContent.includes('監査') || dept.includes('監査')) {
    if (workContent.includes('会計監査')) {
      regulations.push('会社法第436条（計算書類の監査）');
      regulations.push('金商法第193条の2（監査証明）');
    }
    if (workContent.includes('内部監査')) {
      regulations.push('J-SOX（金商法第24条の4の4）');
    }
  }
  
  // 個人情報保護
  if (workContent.includes('個人情報') || workContent.includes('顧客情報')) {
    regulations.push('個人情報保護法');
    regulations.push('GDPR（EU居住者データを扱う場合）');
  }
  
  // インサイダー取引規制
  if (workContent.includes('重要事実') || workContent.includes('適時開示')) {
    regulations.push('金商法第166条（インサイダー取引規制）');
    regulations.push('東証適時開示規則');
  }
  
  // 労働関連
  if (dept.includes('人事') || workContent.includes('労働') || workContent.includes('雇用')) {
    regulations.push('労働基準法');
    if (workContent.includes('36協定')) {
      regulations.push('労基法第36条（時間外労働）');
    }
  }
  
  return regulations.length > 0 ? regulations.join('、') : 'なし';
}

// 内部統制の観点をチェックする関数
function checkInternalControl(processName, workContent, condition, dept) {
  const controls = [];
  
  // 職務分離
  if (condition && condition !== 'なし') {
    controls.push('職務分離の原則');
  }
  
  // 承認権限
  if (workContent.includes('承認') || workContent.includes('決裁')) {
    controls.push('承認権限規程の遵守');
    if (workContent.includes('金額')) {
      controls.push('金額基準による承認権限の設定');
    }
  }
  
  // 文書化
  if (workContent.includes('記録') || workContent.includes('議事録') || workContent.includes('文書')) {
    controls.push('文書化（Documentation）');
    controls.push('監査証跡の保持');
  }
  
  // IT統制
  if (workContent.includes('システム') || workContent.includes('データ')) {
    controls.push('IT全般統制（ITGC）');
    if (workContent.includes('アクセス')) {
      controls.push('アクセス権限管理');
    }
    if (workContent.includes('バックアップ')) {
      controls.push('データバックアップ体制');
    }
  }
  
  // 財務報告
  if (processName.includes('決算') || workContent.includes('財務') || workContent.includes('会計')) {
    controls.push('財務報告に係る内部統制（J-SOX）');
    if (workContent.includes('仕訳')) {
      controls.push('仕訳承認プロセス');
    }
  }
  
  // リスク評価
  if (workContent.includes('リスク') || workContent.includes('評価')) {
    controls.push('リスク評価と対応');
    controls.push('COSOフレームワーク準拠');
  }
  
  // モニタリング
  if (workContent.includes('確認') || workContent.includes('検証') || workContent.includes('レビュー')) {
    controls.push('独立的モニタリング');
    controls.push('予防的統制');
  }
  
  // 相互牽制
  if (dept.includes('経理') || dept.includes('財務')) {
    controls.push('相互牽制体制');
    if (workContent.includes('出納') || workContent.includes('支払')) {
      controls.push('出納業務の分離');
    }
  }
  
  return controls.length > 0 ? controls.join('、') : 'なし';
}

// コンプライアンス留意点を生成する関数（ガバナンス強化版）
function generateComplianceNotes(processName, workContent, timing, dept, condition) {
  const notes = [];
  
  // 外部専門家相談チェックリスト生成
  const taskDescription = `${processName} - ${workContent}`;
  const requiredAdvisors = determineRequiredAdvisors(taskDescription);
  if (requiredAdvisors.length > 0) {
    const checklist = generateConsultationChecklist(taskDescription, requiredAdvisors);
    notes.push('【最優先】外部専門家への事前相談実施');
    checklist.consultationSteps.forEach(step => {
      if (step.phase.includes('専門家')) {
        notes.push(`- ${step.phase}: ${step.timeline}`);
      }
    });
  }
  
  // 株主総会特有の留意点
  if (processName.includes('株主総会')) {
    notes.push('【重要】招集通知は法定期限（2週間前）を厳守');
    if (workContent.includes('議決権')) {
      notes.push('議決権行使書の管理を徹底（改ざん防止）');
    }
    if (workContent.includes('質問')) {
      notes.push('想定問答集の事前準備と法務確認');
    }
  }
  
  // 開示関連
  if (workContent.includes('開示') || workContent.includes('IR')) {
    notes.push('【開示】東証への事前相談を検討');
    notes.push('公平開示の原則を遵守（フェア・ディスクロージャー）');
    if (workContent.includes('業績')) {
      notes.push('業績予想の修正は速やかに開示（軽微基準の確認）');
    }
  }
  
  // 決算関連
  if (processName.includes('決算') || workContent.includes('決算')) {
    notes.push('【決算】会計監査人との事前協議を実施');
    notes.push('重要な会計上の見積りは文書化');
    if (timing.includes('四半期')) {
      notes.push('四半期レビュー対応（監査より簡易だが重要）');
    }
  }
  
  // インサイダー情報管理
  if (workContent.includes('重要事実') || workContent.includes('未公表')) {
    notes.push('【インサイダー】情報管理を徹底（need to know原則）');
    notes.push('役職員の自社株売買は事前申請制');
  }
  
  // データ保護
  if (workContent.includes('個人情報') || workContent.includes('データ')) {
    notes.push('【個人情報】取得時に利用目的を明示');
    notes.push('第三者提供には本人同意が必要');
    if (workContent.includes('削除') || workContent.includes('廃棄')) {
      notes.push('データ削除は復元不可能な方法で実施');
    }
  }
  
  // 契約関連
  if (workContent.includes('契約') || workContent.includes('締結')) {
    notes.push('【契約】法務部門の事前レビュー必須');
    notes.push('利益相反取引は取締役会承認が必要');
  }
  
  // 監査対応
  if (workContent.includes('監査')) {
    notes.push('【監査】監査調書は7年間保存');
    notes.push('監査人の独立性を阻害する行為は禁止');
  }
  
  // リスク管理全般
  if (condition && condition !== 'なし') {
    notes.push('【統制】判断基準を明文化し、恣意性を排除');
    notes.push('例外処理は必ず上位者の承認を取得');
  }
  
  // タイミングに関する留意点
  if (timing.includes('期限') || timing.includes('以内')) {
    notes.push('【期限】法定期限がある場合は余裕を持ったスケジュール設定');
  }
  
  return notes.length > 0 ? notes.join('\n') : '特になし';
}

// 作業内容を個別のアクションに分割する関数
function splitIntoActions(workContent) {
  if (!workContent || typeof workContent !== 'string') {
    return [''];
  }
  
  // 複数の区切り文字で分割（句読点、改行、「・」など）
  const separators = [
    '。',      // 句点
    '\n',      // 改行
    '・',      // 中黒
    '、その後', // 順序を示す表現
    '、次に',   // 順序を示す表現
    '→',       // 矢印
    '①', '②', '③', '④', '⑤', // 番号付きリスト
    '1.', '2.', '3.', '4.', '5.', // 番号付きリスト
    '；'        // セミコロン
  ];
  
  let actions = [workContent];
  
  // 各区切り文字で分割を試みる
  for (const separator of separators) {
    let tempActions = [];
    for (const action of actions) {
      if (action.includes(separator)) {
        const parts = action.split(separator);
        tempActions.push(...parts);
      } else {
        tempActions.push(action);
      }
    }
    actions = tempActions;
  }
  
  // 「および」「また」「さらに」で始まる部分も分割
  let finalActions = [];
  for (const action of actions) {
    if (action.match(/^(および|また|さらに|そして)/)) {
      // 接続詞で始まる場合は独立したアクションとして扱う
      finalActions.push(action);
    } else if (action.includes('および') || action.includes('また')) {
      // 文中に接続詞がある場合も分割を検討
      const subParts = action.split(/(?=および|また)/);
      finalActions.push(...subParts);
    } else {
      finalActions.push(action);
    }
  }
  
  // 空白のみのアクションを除去し、トリミング
  finalActions = finalActions
    .map(action => action.trim())
    .filter(action => action.length > 0);
  
  // アクションが空の場合は元のテキストを返す
  if (finalActions.length === 0) {
    return [workContent];
  }
  
  // 各アクションに連番を付ける（オプション）
  const numbered = finalActions.map((action, index) => {
    // すでに番号が付いている場合はそのまま
    if (action.match(/^[①-⑩\d+\.]/)) {
      return action;
    }
    // 短いアクション（10文字以下）の場合は番号を付けない
    if (action.length <= 10) {
      return action;
    }
    // それ以外は番号を付ける
    return `${index + 1}. ${action}`;
  });
  
  console.log(`アクション分割結果: ${numbered.length}個`);
  numbered.forEach((action, i) => {
    console.log(`  アクション${i + 1}: ${action.substring(0, 50)}${action.length > 50 ? '...' : ''}`);
  });
  
  return numbered;
}

// flow_rowsデータのクリーニング関数
function cleanFlowRowsData(flowRows) {
  console.log('flow_rowsクリーニング開始');
  
  if (!flowRows) {
    console.log('flow_rowsがnullまたはundefined');
    return [];
  }
  
  // 配列でない場合は配列に変換
  if (!Array.isArray(flowRows)) {
    console.log('flow_rowsが配列ではないため変換');
    flowRows = [flowRows];
  }
  
  const cleaned = flowRows.map((row, index) => {
    console.log(`行${index + 1}のクリーニング開始`);
    
    // オブジェクトの場合はそのまま処理
    if (typeof row === 'object' && row !== null && !Array.isArray(row)) {
      const cleanedRow = {};
      for (const key in row) {
        let value = row[key];
        
        // 値をクリーニング
        if (typeof value === 'string') {
          // 末尾の数字を削除（「特になし3」→「特になし」）
          value = value.replace(/(\D+)\d+$/, '$1').trim();
          // 「なし」の末尾の数字も削除
          value = value.replace(/^(なし)\d+$/, '$1');
          value = value.replace(/^(特になし)\d+$/, '$1');
        }
        
        cleanedRow[key] = value || '';
      }
      console.log(`行${index + 1}クリーニング完了（オブジェクト）`);
      return cleanedRow;
    }
    
    // 文字列の場合
    if (typeof row === 'string') {
      console.log(`行${index + 1}は文字列: ${row.substring(0, 100)}`);
      // カンマ区切りで分割
      const parts = row.split(',').map(part => {
        let cleaned = part.trim();
        // 末尾の数字を削除
        cleaned = cleaned.replace(/(\D+)\d+$/, '$1').trim();
        cleaned = cleaned.replace(/^(なし)\d+$/, '$1');
        cleaned = cleaned.replace(/^(特になし)\d+$/, '$1');
        return cleaned;
      });
      
      // ヘッダーに基づいてオブジェクトを作成
      const headers = ['工程', '実施タイミング', '部署', '担当役割', '作業内容', '条件分岐', '利用ツール', 'URLリンク', '備考'];
      const cleanedRow = {};
      headers.forEach((header, i) => {
        cleanedRow[header] = parts[i] || '';
      });
      console.log(`行${index + 1}クリーニング完了（文字列→オブジェクト）`);
      return cleanedRow;
    }
    
    // 配列の場合
    if (Array.isArray(row)) {
      console.log(`行${index + 1}は配列`);
      const headers = ['工程', '実施タイミング', '部署', '担当役割', '作業内容', '条件分岐', '利用ツール', 'URLリンク', '備考'];
      const cleanedRow = {};
      headers.forEach((header, i) => {
        let value = row[i] || '';
        if (typeof value === 'string') {
          // 末尾の数字を削除
          value = value.replace(/(\D+)\d+$/, '$1').trim();
          value = value.replace(/^(なし)\d+$/, '$1');
          value = value.replace(/^(特になし)\d+$/, '$1');
        }
        cleanedRow[header] = value;
      });
      console.log(`行${index + 1}クリーニング完了（配列→オブジェクト）`);
      return cleanedRow;
    }
    
    console.warn(`行${index + 1}は未対応の型: ${typeof row}`);
    return null;
  }).filter(row => row !== null);
  
  console.log(`クリーニング完了: ${cleaned.length}行`);
  return cleaned;
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

// 完全に安全な新しいフロー行書き込み関数（1アクション1セル形式）
function writeFlowRowsSafe(flowRows) {
  const sheetName = getConfig('FLOW_SHEET_NAME') || FLOW_SHEET;
  const sheet = ss().getSheetByName(sheetName) || createFlowSheet(sheetName);
  const headers = FLOW_HEADERS; // 定数を使用（法令・規制等を含む）

  console.log('=== 安全なフロー行書き込み開始（1アクション1セル形式） ===');
  
  // 詳細なデバッグ情報を出力
  debugDataStructure(flowRows, '入力データ (flowRows)');
  
  // データを安全に処理（1アクション1セル形式）
  let processedData = [];
  
  try {
    // flowRowsが配列かどうかチェック
    if (!flowRows) {
      console.log('データがnullまたはundefined');
      return;
    }
    
    if (Array.isArray(flowRows)) {
      console.log(`配列として受信: ${flowRows.length}個の要素`);
      
      // 各要素を安全に処理（作業内容を分割）
      for (let i = 0; i < flowRows.length; i++) {
        const row = flowRows[i];
        console.log(`\n--- 行${i + 1}の処理開始 ---`);
        debugDataStructure(row, `行${i + 1}`);
        
        if (typeof row === 'object' && row !== null) {
          // 作業内容を分割して複数行に展開（1アクション1セル）
          const workContent = row['作業内容'] || '';
          const actions = splitIntoActions(workContent);
          
          console.log(`作業内容を${actions.length}個のアクションに分割`);
          
          // 各アクションごとに行を作成（法令・内部統制の観点を追加）
          for (let j = 0; j < actions.length; j++) {
            const rowArray = [];
            const processName = row['工程'] || '';
            const timing = row['実施タイミング'] || '';
            const dept = row['部署'] || '';
            const condition = row['条件分岐'] || '';
            
            for (const header of headers) {
              let value = '';
              
              if (header === '作業内容') {
                // 作業内容は分割されたアクション
                value = actions[j];
              } else if (header === '法令・規制') {
                // 法令・規制を自動判定
                value = checkLegalRegulations(processName, actions[j], timing, dept);
              } else if (header === '内部統制') {
                // 内部統制の観点を自動判定
                value = checkInternalControl(processName, actions[j], condition, dept);
              } else if (header === 'コンプライアンス留意点') {
                // コンプライアンス留意点を自動生成
                value = j === 0 ? generateComplianceNotes(processName, actions[j], timing, dept, condition) : '';
              } else if (j === 0) {
                // 最初のアクションの場合は全ての情報を含める
                value = row[header] || '';
              } else {
                // 2番目以降のアクションは作業内容以外を空にするか、継続する情報のみ
                if (header === '工程' || header === '実施タイミング' || header === '部署' || header === '担当役割') {
                  value = row[header] || '';
                } else {
                  value = '';
                }
              }
              
              // 末尾の不要な数字を削除
              const cleanValue = String(value).replace(/特になし\d+$/, '特になし').replace(/なし\d+$/, 'なし');
              rowArray.push(cleanValue);
            }
            processedData.push(rowArray);
          }
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

// ================================================================================
// 8. governance_functions.gs - ガバナンス・コンプライアンス機能
// ================================================================================

/**
 * ガバナンス機能設定
 */
const GOVERNANCE_CONFIG = {
  enableDisclosureCheck: true,
  enableAdvisorConsultation: true,
  enableMECEClassification: true,
  autoGenerateTimeline: true,
  strictComplianceMode: true
};

// 開示判定マスターデータ
const DISCLOSURE_TRIGGERS = {
  TIMELY_DISCLOSURE_DECISIONS: {
    '株式発行': {
      criteria: ['新株発行', '増資', '第三者割当', '公募', '株主割当'],
      threshold: '発行済株式総数の10%以上',
      timeline: '決議後直ちに',
      authority: '取締役会決議',
      documents: ['有価証券届出書', '適時開示資料'],
      regulations: ['金商法第4条', '東証適時開示規則第2条']
    },
    '資本政策': {
      criteria: ['自己株式取得', '資本金減少', '株式分割', '株式併合'],
      threshold: '資本金の10%以上',
      timeline: '決議後直ちに',
      authority: '取締役会決議（一部株主総会）',
      documents: ['適時開示資料', '臨時報告書'],
      regulations: ['会社法第156条', '東証適時開示規則']
    },
    'M&A': {
      criteria: ['合併', '会社分割', '株式交換', '株式移転', '事業譲渡'],
      threshold: '純資産の30%以上',
      timeline: '基本合意時及び決議後',
      authority: '取締役会決議及び株主総会特別決議',
      documents: ['適時開示資料', '臨時報告書', '公開買付届出書'],
      regulations: ['会社法第783条', '金商法第27条の3']
    },
    '業務提携': {
      criteria: ['資本提携', '業務提携', '技術提携'],
      threshold: '売上高の10%以上の影響',
      timeline: '契約締結後直ちに',
      authority: '取締役会決議',
      documents: ['適時開示資料'],
      regulations: ['東証適時開示規則第2条']
    }
  },
  TIMELY_DISCLOSURE_EVENTS: {
    '災害・事故': {
      criteria: ['火災', '爆発', '自然災害', '事故'],
      threshold: '純資産の3%以上の損害',
      timeline: '発生後直ちに',
      authority: '代表取締役',
      documents: ['適時開示資料'],
      regulations: ['東証適時開示規則第2条']
    },
    '訴訟': {
      criteria: ['訴訟提起', '仲裁申立', '調停申立'],
      threshold: '純資産の15%以上の請求',
      timeline: '提起後直ちに',
      authority: '法務部門',
      documents: ['適時開示資料'],
      regulations: ['東証適時開示規則']
    }
  },
  FINANCIAL_DISCLOSURE: {
    '決算短信': {
      criteria: ['四半期決算', '通期決算'],
      timeline: '決算後45日以内（推奨30日）',
      authority: '取締役会承認',
      documents: ['決算短信', '四半期報告書'],
      regulations: ['東証決算短信作成要領']
    },
    '業績予想修正': {
      criteria: ['売上高', '営業利益', '経常利益', '純利益', '配当'],
      threshold: '10%以上の乖離',
      timeline: '判明後直ちに',
      authority: '取締役会決議',
      documents: ['適時開示資料'],
      regulations: ['東証適時開示規則']
    }
  },
  STATUTORY_DISCLOSURE: {
    '有価証券報告書': {
      timeline: '事業年度終了後3か月以内',
      authority: '代表取締役',
      edinet: true,
      regulations: ['金商法第24条']
    },
    '四半期報告書': {
      timeline: '四半期終了後45日以内',
      authority: '代表取締役',
      edinet: true,
      regulations: ['金商法第24条の4の7']
    },
    '臨時報告書': {
      triggers: ['主要株主異動', '代表取締役異動', '監査人異動'],
      timeline: '発生後遅滞なく',
      authority: '代表取締役',
      edinet: true,
      regulations: ['金商法第24条の5']
    },
    '内部統制報告書': {
      timeline: '有価証券報告書と同時',
      authority: '代表取締役',
      edinet: true,
      regulations: ['金商法第24条の4の4']
    }
  }
};

// 外部専門家マスターデータ
const EXTERNAL_ADVISORS = {
  '法律事務所': {
    specialties: ['M&A・企業再編', 'コーポレートガバナンス', '株主総会対応', '取締役会運営', '契約法務', 'コンプライアンス'],
    consultationTiming: '重要な法的判断が必要な段階の初期',
    deliverables: ['リーガルオピニオン', '契約書レビュー', 'デューデリジェンス報告書'],
    urgencyLevels: { 'CRITICAL': '即日対応', 'HIGH': '2-3営業日', 'MEDIUM': '1週間', 'LOW': '2週間' }
  },
  '監査法人': {
    specialties: ['会計監査', '内部統制監査（J-SOX）', '四半期レビュー', 'M&Aデューデリジェンス'],
    consultationTiming: '決算前・重要な会計処理の変更前',
    deliverables: ['監査報告書', '内部統制監査報告書', 'マネジメントレター'],
    urgencyLevels: { 'CRITICAL': '即日対応', 'HIGH': '3-5営業日', 'MEDIUM': '2週間', 'LOW': '1か月' }
  },
  '税理士事務所': {
    specialties: ['税務申告', '税務調査対応', '移転価格', '国際税務', '組織再編税制'],
    consultationTiming: '税務判断が必要な取引の実行前',
    deliverables: ['税務意見書', '税務デューデリジェンス報告書', 'タックスストラクチャリング提案'],
    urgencyLevels: { 'CRITICAL': '即日対応', 'HIGH': '2-3営業日', 'MEDIUM': '1週間', 'LOW': '2週間' }
  },
  '司法書士事務所': {
    specialties: ['商業登記', '不動産登記', '役員変更登記', '定款変更', '増資・減資登記'],
    consultationTiming: '登記が必要な決議の前',
    deliverables: ['登記申請書', '定款', '議事録作成支援'],
    urgencyLevels: { 'CRITICAL': '当日対応', 'HIGH': '1-2営業日', 'MEDIUM': '3-5営業日', 'LOW': '1週間' }
  },
  '社会保険労務士事務所': {
    specialties: ['就業規則作成・変更', '労働契約', '労使協定', '社会保険手続き'],
    consultationTiming: '人事労務施策の実施前',
    deliverables: ['就業規則', '労使協定書', '労務監査報告書'],
    urgencyLevels: { 'CRITICAL': '即日対応', 'HIGH': '2-3営業日', 'MEDIUM': '1週間', 'LOW': '2週間' }
  }
};

/**
 * タスクが開示対象かを判定
 */
function checkDisclosureRequirement(task) {
  const result = {
    requiresDisclosure: false,
    disclosureType: [],
    timeline: [],
    authorities: [],
    documents: [],
    regulations: [],
    notes: []
  };

  const taskLower = task.toLowerCase();
  
  // 株主総会関連
  if (taskLower.includes('株主総会')) {
    result.requiresDisclosure = true;
    result.disclosureType.push('株主総会関連開示');
    
    if (taskLower.includes('定時')) {
      result.timeline.push('招集通知: 総会2週間前');
      result.timeline.push('招集通知Web開示: 発送前');
      result.documents.push('招集通知', '事業報告', '計算書類');
      result.regulations.push('会社法第299条');
    }
    
    if (taskLower.includes('臨時')) {
      result.timeline.push('適時開示: 招集決定後直ちに');
      result.documents.push('適時開示資料', '招集通知');
      result.regulations.push('東証適時開示規則');
    }
    
    result.authorities.push('取締役会', '代表取締役');
    result.notes.push('TDnetでの開示と自社Webサイトでの公表を並行実施');
  }

  // 取締役会関連
  if (taskLower.includes('取締役会')) {
    const boardItems = ['決算', '配当', '自己株式', '役員', '組織再編', '業務提携'];
    for (const item of boardItems) {
      if (taskLower.includes(item)) {
        result.requiresDisclosure = true;
        result.disclosureType.push('取締役会決議事項');
        result.timeline.push('決議後直ちに');
        result.authorities.push('取締役会');
        result.documents.push('適時開示資料');
        result.regulations.push('東証適時開示規則第2条');
        break;
      }
    }
  }

  // M&A・組織再編
  if (taskLower.includes('合併') || taskLower.includes('買収') || 
      taskLower.includes('m&a') || taskLower.includes('事業譲渡')) {
    result.requiresDisclosure = true;
    result.disclosureType.push('組織再編・M&A');
    result.timeline.push('基本合意時', '最終契約時', '効力発生時');
    result.authorities.push('取締役会', '株主総会（特別決議）');
    result.documents.push('適時開示資料', '臨時報告書', '公開買付届出書');
    result.regulations.push('金商法第27条の3', '会社法第783条');
    result.notes.push('財務アドバイザー・法務アドバイザーとの連携必須');
  }

  // 決算・業績関連
  if (taskLower.includes('決算') || taskLower.includes('業績')) {
    result.requiresDisclosure = true;
    result.disclosureType.push('決算開示');
    
    if (taskLower.includes('四半期')) {
      result.timeline.push('四半期終了後45日以内');
      result.documents.push('四半期決算短信', '四半期報告書');
      result.regulations.push('金商法第24条の4の7');
    } else if (taskLower.includes('通期') || taskLower.includes('年度')) {
      result.timeline.push('期末後45日以内（決算短信）', '期末後3か月以内（有価証券報告書）');
      result.documents.push('決算短信', '有価証券報告書');
      result.regulations.push('金商法第24条');
    }
    
    if (taskLower.includes('修正') || taskLower.includes('予想')) {
      result.timeline.push('判明後直ちに（業績予想修正）');
      result.notes.push('売上高・利益が10%以上乖離する場合は開示必須');
    }
    
    result.authorities.push('取締役会', '監査役会', '会計監査人');
  }

  return result;
}

/**
 * タスクに対して必要な外部専門家を判定
 */
function determineRequiredAdvisors(task, context = {}) {
  const requiredAdvisors = [];
  const taskLower = task.toLowerCase();
  
  // 必須相談パターンの定義
  const mandatoryPatterns = [
    {
      keywords: ['株主総会', '株主', '総会'],
      advisors: ['法律事務所', '司法書士事務所'],
      reason: '株主総会の適法な運営と手続きの確認',
      checkpoints: ['招集手続きの適法性確認', '議案の適法性確認', '決議要件の確認', '議事録作成要領の確認']
    },
    {
      keywords: ['取締役会', '取締役', '役員', '執行役'],
      advisors: ['法律事務所', '司法書士事務所'],
      reason: '取締役会運営の適法性と役員変更登記',
      checkpoints: ['決議事項の適法性確認', '利益相反取引の確認', '特別利害関係の確認', '登記手続きの確認']
    },
    {
      keywords: ['M&A', '買収', '合併', '事業譲渡', '会社分割'],
      advisors: ['法律事務所', '監査法人', '税理士事務所'],
      reason: 'M&A取引の法務・財務・税務面での総合的検証',
      checkpoints: ['ストラクチャーの検討', 'デューデリジェンスの実施', '価格の妥当性検証', '契約条件の交渉']
    },
    {
      keywords: ['決算', '財務諸表', '有価証券報告書', '四半期報告書'],
      advisors: ['監査法人', '税理士事務所'],
      reason: '適正な財務報告と税務申告',
      checkpoints: ['会計処理の妥当性確認', '開示内容の適切性確認', '内部統制の有効性評価', '税務リスクの確認']
    },
    {
      keywords: ['増資', '減資', '自己株式', '新株', '社債'],
      advisors: ['法律事務所', '司法書士事務所'],
      reason: '資本政策の適法性と実行可能性の確認',
      checkpoints: ['発行条件の妥当性', '既存株主への影響分析', '開示書類の作成', '登記手続きの準備']
    },
    {
      keywords: ['労働', '雇用', '解雇', '就業規則', 'ハラスメント'],
      advisors: ['社会保険労務士事務所', '法律事務所'],
      reason: '労働法令遵守と労使紛争の予防',
      checkpoints: ['労働法令の遵守確認', '就業規則の整備', '労使協定の締結', '紛争リスクの評価']
    },
    {
      keywords: ['契約', '締結', '変更', '解除'],
      advisors: ['法律事務所'],
      reason: '契約リスクの評価と条件交渉',
      checkpoints: ['契約条件の妥当性確認', 'リスク条項の確認', '責任範囲の明確化', '紛争解決条項の確認']
    },
    {
      keywords: ['コンプライアンス', '違反', '不正', '内部統制'],
      advisors: ['法律事務所', '監査法人'],
      reason: 'コンプライアンス体制の強化と違反防止',
      checkpoints: ['現状のリスク評価', '改善策の立案', 'モニタリング体制の構築', '教育研修の実施']
    },
    {
      keywords: ['訴訟', '紛争', '係争', '調停', '仲裁'],
      advisors: ['法律事務所'],
      reason: '法的紛争の適切な解決',
      checkpoints: ['勝訴可能性の評価', '和解条件の検討', '証拠の収集・保全', '訴訟戦略の立案']
    },
    {
      keywords: ['個人情報', 'プライバシー', 'GDPR', '情報漏洩'],
      advisors: ['法律事務所'],
      reason: '個人情報保護法令の遵守',
      checkpoints: ['現行体制の評価', '規程・手順の整備', 'セキュリティ対策の確認', 'インシデント対応体制の構築']
    }
  ];

  // パターンマッチング
  mandatoryPatterns.forEach(pattern => {
    const hasKeyword = pattern.keywords.some(keyword => taskLower.includes(keyword));
    if (hasKeyword) {
      pattern.advisors.forEach(advisor => {
        requiredAdvisors.push({
          type: advisor,
          reason: pattern.reason,
          priority: 'MANDATORY',
          checkpoints: pattern.checkpoints,
          timing: EXTERNAL_ADVISORS[advisor].consultationTiming
        });
      });
    }
  });

  // 金額基準での判定
  if (context.amount) {
    const amount = parseInt(context.amount);
    if (amount > 100000000) { // 1億円以上
      requiredAdvisors.push({
        type: '法律事務所',
        reason: '高額取引のため法的リスク評価が必要',
        priority: 'HIGH',
        checkpoints: ['契約条件の精査', 'リスク分析', '交渉戦略']
      });
    }
  }

  return requiredAdvisors;
}

/**
 * 外部専門家相談チェックリスト生成
 */
function generateConsultationChecklist(task, advisors) {
  const checklist = {
    task: task,
    consultationSteps: [],
    documentationRequired: [],
    timeline: [],
    budgetConsiderations: []
  };

  // ステップ1: 事前準備
  checklist.consultationSteps.push({
    step: 1,
    phase: '事前準備',
    actions: [
      '相談事項の明確化と論点整理',
      '関連資料の収集と整理',
      '社内での事前検討と方針案の作成',
      '予算の確保と決裁取得'
    ],
    timeline: 'T-14日',
    responsible: '担当部門'
  });

  // ステップ2: 専門家選定
  checklist.consultationSteps.push({
    step: 2,
    phase: '専門家選定',
    actions: [
      '複数の専門家候補のリストアップ',
      '見積もり取得と比較検討',
      '利益相反チェック',
      '秘密保持契約（NDA）の締結'
    ],
    timeline: 'T-10日',
    responsible: '法務部・総務部'
  });

  // ステップ3: 各専門家への相談実施
  advisors.forEach((advisor, index) => {
    checklist.consultationSteps.push({
      step: 3 + index,
      phase: `${advisor.type}への相談`,
      actions: [
        '初回ミーティングの実施',
        '詳細情報の提供と質疑応答',
        '中間報告の受領とフィードバック',
        '最終意見書・報告書の受領'
      ],
      timeline: `T-${7 - index}日`,
      responsible: `担当部門・${advisor.type}`,
      deliverables: EXTERNAL_ADVISORS[advisor.type].deliverables,
      checkpoints: advisor.checkpoints || []
    });
  });

  // ステップ4: 社内検討
  checklist.consultationSteps.push({
    step: 3 + advisors.length,
    phase: '社内検討・意思決定',
    actions: [
      '専門家意見の社内共有と検討',
      'リスク評価と対応策の決定',
      '実行計画の策定',
      '必要な社内承認の取得'
    ],
    timeline: 'T-2日',
    responsible: '経営陣・関連部門'
  });

  // 必要書類リスト
  checklist.documentationRequired = [
    '相談依頼書',
    '背景説明資料',
    '関連契約書・規程類',
    '財務データ（必要に応じて）',
    '過去の類似案件資料',
    '社内検討資料',
    '取締役会・経営会議資料'
  ];

  return checklist;
}

/**
 * MECEなタスク分類体系
 */
const TASK_CLASSIFICATION_MECE = {
  'ガバナンス・コンプライアンス': {
    '株主総会運営': {
      tasks: ['定時株主総会の準備・開催', '臨時株主総会の準備・開催', '株主総会招集通知の作成・送付'],
      disclosure: true,
      priority: 'HIGH'
    },
    '取締役会運営': {
      tasks: ['取締役会の開催・運営', '取締役会議事録の作成', '取締役会規程の管理'],
      disclosure: true,
      priority: 'HIGH'
    },
    '監査対応': {
      tasks: ['監査役監査への対応', '内部監査への対応', '会計監査人監査への対応'],
      disclosure: false,
      priority: 'HIGH'
    },
    'コンプライアンス管理': {
      tasks: ['コンプライアンス違反の防止・発見', '内部通報制度の運営', 'コンプライアンス研修の実施'],
      disclosure: false,
      priority: 'MEDIUM'
    }
  },
  '情報開示・IR': {
    '適時開示': {
      tasks: ['決定事実の開示', '発生事実の開示', '決算情報の開示', '業績予想修正の開示'],
      disclosure: true,
      priority: 'CRITICAL'
    },
    '法定開示': {
      tasks: ['有価証券報告書の作成・提出', '四半期報告書の作成・提出', '臨時報告書の作成・提出'],
      disclosure: true,
      priority: 'CRITICAL'
    },
    'IR活動': {
      tasks: ['決算説明会の開催', 'アナリスト・機関投資家対応', '個人投資家向け説明会'],
      disclosure: false,
      priority: 'HIGH'
    }
  },
  '内部統制・リスク管理': {
    '内部統制システム': {
      tasks: ['J-SOX対応', '内部統制の整備・運用', '内部統制の評価', '内部統制報告書の作成'],
      disclosure: true,
      priority: 'HIGH'
    },
    'リスク管理': {
      tasks: ['リスクアセスメント', 'リスク対応策の策定', 'BCP（事業継続計画）の策定・更新'],
      disclosure: false,
      priority: 'HIGH'
    }
  },
  '経営管理': {
    '経営企画': {
      tasks: ['中期経営計画の策定', '年度事業計画の策定', '予算策定・管理', 'KPI管理'],
      disclosure: false,
      priority: 'HIGH'
    },
    '組織管理': {
      tasks: ['組織変更・改編', '規程・規則の制定・改廃', '権限委譲・決裁権限の管理'],
      disclosure: false,
      priority: 'MEDIUM'
    }
  }
};

/**
 * タスクをMECE分類に振り分け
 */
function classifyTaskMECE(task) {
  const classification = {
    level1: null,
    level2: null,
    level3: null,
    requiresDisclosure: false,
    priority: 'LOW',
    relatedTasks: []
  };

  const taskLower = task.toLowerCase();

  // 各分類を検査
  for (const [l1Key, l1Value] of Object.entries(TASK_CLASSIFICATION_MECE)) {
    for (const [l2Key, l2Value] of Object.entries(l1Value)) {
      for (const l3Task of l2Value.tasks) {
        if (taskLower.includes(l3Task.toLowerCase()) || 
            l3Task.toLowerCase().includes(taskLower)) {
          classification.level1 = l1Key;
          classification.level2 = l2Key;
          classification.level3 = l3Task;
          classification.requiresDisclosure = l2Value.disclosure;
          classification.priority = l2Value.priority;
          classification.relatedTasks = l2Value.tasks.filter(t => t !== l3Task);
          return classification;
        }
      }
    }
  }

  // マッチしない場合はキーワードベースで推定
  if (taskLower.includes('開示') || taskLower.includes('報告書')) {
    classification.level1 = '情報開示・IR';
    classification.requiresDisclosure = true;
    classification.priority = 'HIGH';
  } else if (taskLower.includes('監査') || taskLower.includes('統制')) {
    classification.level1 = '内部統制・リスク管理';
    classification.priority = 'HIGH';
  } else if (taskLower.includes('取締役') || taskLower.includes('株主')) {
    classification.level1 = 'ガバナンス・コンプライアンス';
    classification.requiresDisclosure = true;
    classification.priority = 'HIGH';
  }

  return classification;
}

/**
 * 統合的なガバナンスチェック
 */
function performComprehensiveGovernanceCheck(workSpec, flowData) {
  const governanceReport = {
    overallScore: 0,
    disclosureRequirements: [],
    advisorConsultations: [],
    complianceGaps: [],
    recommendations: [],
    timeline: [],
    riskAssessment: []
  };

  // 1. 業務仕様書からガバナンス要素を抽出
  if (workSpec) {
    const specText = JSON.stringify(workSpec).toLowerCase();
    const disclosureCheck = checkDisclosureRequirement(specText);
    if (disclosureCheck.requiresDisclosure) {
      governanceReport.disclosureRequirements.push({
        type: disclosureCheck.disclosureType.join(', '),
        timeline: disclosureCheck.timeline,
        documents: disclosureCheck.documents,
        regulations: disclosureCheck.regulations
      });
    }
  }

  // 2. フローデータから承認プロセスを分析
  if (flowData && Array.isArray(flowData)) {
    const approvalSteps = flowData.filter(row => 
      row['作業内容'] && (
        row['作業内容'].includes('承認') ||
        row['作業内容'].includes('決裁') ||
        row['作業内容'].includes('決議')
      )
    );

    // 承認階層の適切性を評価
    const requiredApprovers = new Set();
    approvalSteps.forEach(step => {
      if (step['担当役割']) {
        requiredApprovers.add(step['担当役割']);
      }
    });

    // 必要な承認者が不足していないかチェック
    const essentialApprovers = ['取締役会', '代表取締役', '監査役'];
    essentialApprovers.forEach(approver => {
      if (!Array.from(requiredApprovers).some(r => r.includes(approver))) {
        if (governanceReport.disclosureRequirements.length > 0) {
          governanceReport.complianceGaps.push(
            `重要な承認者「${approver}」が承認フローに含まれていません`
          );
        }
      }
    });
  }

  // 3. リスク評価
  const risks = [
    {
      category: '開示遅延リスク',
      probability: governanceReport.disclosureRequirements.length > 2 ? 'HIGH' : 'MEDIUM',
      impact: 'HIGH',
      mitigation: 'IR部門との事前調整、開示チェックリストの活用'
    },
    {
      category: 'コンプライアンス違反リスク',
      probability: governanceReport.complianceGaps.length > 0 ? 'HIGH' : 'LOW',
      impact: 'CRITICAL',
      mitigation: '法務部門による事前レビュー、コンプライアンスチェックの実施'
    }
  ];
  governanceReport.riskAssessment = risks;

  // 4. 推奨事項の生成
  if (governanceReport.disclosureRequirements.length > 0) {
    governanceReport.recommendations.push(
      '東証への事前相談を検討してください（複雑な開示案件の場合）'
    );
    governanceReport.recommendations.push(
      'IR部門と法務部門の早期巻き込みを推奨します'
    );
  }

  if (governanceReport.complianceGaps.length > 0) {
    governanceReport.recommendations.push(
      '承認フローの見直しと必要な承認者の追加を検討してください'
    );
  }

  // 5. フローデータから専門家相談要件を抽出
  if (flowData && Array.isArray(flowData)) {
    flowData.forEach((row, index) => {
      const taskDescription = `${row['工程'] || ''} ${row['作業内容'] || ''}`;
      const advisors = determineRequiredAdvisors(taskDescription);
      
      if (advisors.length > 0) {
        const checklist = generateConsultationChecklist(taskDescription, advisors);
        governanceReport.advisorConsultations.push({
          taskId: index + 1,
          task: taskDescription,
          advisors: advisors,
          checklist: checklist
        });
      }
    });
  }

  // 6. スコアリング（100点満点）
  let score = 100;
  score -= governanceReport.complianceGaps.length * 10;
  score -= governanceReport.riskAssessment.filter(r => r.probability === 'HIGH').length * 5;
  score = Math.max(0, score);
  governanceReport.overallScore = score;

  return governanceReport;
}

// ガバナンスレポートシートを作成
function createGovernanceReportSheet(sheet, governanceCheck) {
  let row = 1;

  // タイトル
  sheet.getRange(row, 1, 1, 8).merge();
  sheet.getRange(row, 1).setValue('ガバナンス・コンプライアンスチェックレポート');
  sheet.getRange(row, 1).setFontSize(16).setFontWeight('bold');
  sheet.getRange(row, 1).setBackground('#1a73e8').setFontColor('#ffffff');
  row += 2;

  // サマリー
  sheet.getRange(row, 1).setValue('【サマリー】');
  sheet.getRange(row, 1).setFontWeight('bold').setBackground('#e8f0fe');
  row++;
  
  const summaryData = [
    ['ガバナンススコア', governanceCheck.overallScore + '/100点'],
    ['開示要件', governanceCheck.disclosureRequirements.length + '件'],
    ['外部専門家相談', (governanceCheck.advisorConsultations ? governanceCheck.advisorConsultations.length : 0) + '件'],
    ['コンプライアンスギャップ', governanceCheck.complianceGaps.length + '件']
  ];
  
  sheet.getRange(row, 1, summaryData.length, 2).setValues(summaryData);
  sheet.getRange(row, 1, summaryData.length, 1).setFontWeight('bold');
  row += summaryData.length + 2;

  // 開示要件
  if (governanceCheck.disclosureRequirements.length > 0) {
    sheet.getRange(row, 1).setValue('【開示要件】');
    sheet.getRange(row, 1).setFontWeight('bold').setBackground('#fce8b2');
    row++;
    
    const disclosureHeaders = ['No.', '開示種別', '期限', '必要書類', '関連法規'];
    sheet.getRange(row, 1, 1, disclosureHeaders.length).setValues([disclosureHeaders]);
    sheet.getRange(row, 1, 1, disclosureHeaders.length).setFontWeight('bold');
    row++;
    
    governanceCheck.disclosureRequirements.forEach((req, index) => {
      const rowData = [
        index + 1,
        req.type || '',
        Array.isArray(req.timeline) ? req.timeline.join(', ') : '',
        Array.isArray(req.documents) ? req.documents.join(', ') : '',
        Array.isArray(req.regulations) ? req.regulations.join(', ') : ''
      ];
      sheet.getRange(row, 1, 1, rowData.length).setValues([rowData]);
      row++;
    });
    row++;
  }

  // 外部専門家相談
  if (governanceCheck.advisorConsultations && governanceCheck.advisorConsultations.length > 0) {
    sheet.getRange(row, 1).setValue('【外部専門家相談が必要なタスク】');
    sheet.getRange(row, 1).setFontWeight('bold').setBackground('#d9ead3');
    row++;
    
    governanceCheck.advisorConsultations.forEach((consultation, index) => {
      sheet.getRange(row, 1).setValue(`${index + 1}. ${consultation.task}`);
      sheet.getRange(row, 1).setFontWeight('bold');
      row++;
      
      consultation.advisors.forEach(advisor => {
        sheet.getRange(row, 2).setValue(`・${advisor.type}: ${advisor.reason}`);
        row++;
      });
      row++;
    });
  }

  // 推奨事項
  if (governanceCheck.recommendations.length > 0) {
    sheet.getRange(row, 1).setValue('【推奨事項】');
    sheet.getRange(row, 1).setFontWeight('bold').setBackground('#f4cccc');
    row++;
    
    governanceCheck.recommendations.forEach(rec => {
      sheet.getRange(row, 1).setValue(`・${rec}`);
      row++;
    });
    row++;
  }

  // リスク評価
  if (governanceCheck.riskAssessment && governanceCheck.riskAssessment.length > 0) {
    sheet.getRange(row, 1).setValue('【リスク評価】');
    sheet.getRange(row, 1).setFontWeight('bold').setBackground('#ffe599');
    row++;
    
    const riskHeaders = ['リスクカテゴリ', '発生確率', '影響度', '対策'];
    sheet.getRange(row, 1, 1, riskHeaders.length).setValues([riskHeaders]);
    sheet.getRange(row, 1, 1, riskHeaders.length).setFontWeight('bold');
    row++;
    
    governanceCheck.riskAssessment.forEach(risk => {
      const rowData = [
        risk.category,
        risk.probability,
        risk.impact,
        risk.mitigation
      ];
      sheet.getRange(row, 1, 1, rowData.length).setValues([rowData]);
      row++;
    });
  }

  // 書式調整
  sheet.autoResizeColumns(1, 8);
}

// 専門家相談チェックリストシートを作成
function createConsultationChecklistSheet(sheet, consultations) {
  let row = 1;

  // タイトル
  sheet.getRange(row, 1, 1, 6).merge();
  sheet.getRange(row, 1).setValue('外部専門家相談チェックリスト');
  sheet.getRange(row, 1).setFontSize(16).setFontWeight('bold');
  sheet.getRange(row, 1).setBackground('#34a853').setFontColor('#ffffff');
  row += 2;

  consultations.forEach((consultation, consultIndex) => {
    // タスクタイトル
    sheet.getRange(row, 1, 1, 6).merge();
    sheet.getRange(row, 1).setValue(`【タスク${consultIndex + 1}】 ${consultation.task}`);
    sheet.getRange(row, 1).setFontWeight('bold').setBackground('#e8f5e9');
    row++;

    // 必要な専門家
    sheet.getRange(row, 1).setValue('必要な専門家:');
    sheet.getRange(row, 1).setFontWeight('bold');
    row++;
    
    consultation.advisors.forEach(advisor => {
      sheet.getRange(row, 2).setValue(`${advisor.type}`);
      sheet.getRange(row, 3).setValue(`理由: ${advisor.reason}`);
      sheet.getRange(row, 4).setValue(`優先度: ${advisor.priority}`);
      row++;
    });
    row++;

    // 相談ステップ
    if (consultation.checklist && consultation.checklist.consultationSteps) {
      sheet.getRange(row, 1).setValue('相談ステップ:');
      sheet.getRange(row, 1).setFontWeight('bold');
      row++;
      
      const stepHeaders = ['ステップ', 'フェーズ', 'アクション', 'タイミング', '担当'];
      sheet.getRange(row, 1, 1, stepHeaders.length).setValues([stepHeaders]);
      sheet.getRange(row, 1, 1, stepHeaders.length).setBackground('#f0f0f0').setFontWeight('bold');
      row++;
      
      consultation.checklist.consultationSteps.forEach(step => {
        const stepData = [
          step.step,
          step.phase,
          step.actions.join('\n'),
          step.timeline,
          step.responsible
        ];
        sheet.getRange(row, 1, 1, stepData.length).setValues([stepData]);
        sheet.getRange(row, 3).setWrap(true);
        row++;
      });
      row += 2;
    }
  });

  // 必要書類リスト
  sheet.getRange(row, 1).setValue('【準備が必要な書類】');
  sheet.getRange(row, 1).setFontWeight('bold').setBackground('#fce8b2');
  row++;
  
  const documents = [
    '相談依頼書',
    '背景説明資料',
    '関連契約書・規程類',
    '財務データ（必要に応じて）',
    '過去の類似案件資料',
    '社内検討資料',
    '取締役会・経営会議資料'
  ];
  
  documents.forEach(doc => {
    sheet.getRange(row, 1).setValue(`□ ${doc}`);
    row++;
  });

  // 書式調整
  sheet.autoResizeColumns(1, 6);
}

// ガバナンス情報を追加したフローシートを作成
function writeFlowToSheetWithGovernance(sheet, flowRows, governanceCheck) {
  const headers = FLOW_HEADERS; // 定数を使用して一貫性を保つ
  
  // ガバナンス情報を追加したヘッダー
  const enhancedHeaders = [...headers, '開示要件', '要専門家相談', '優先度'];
  
  sheet.getRange(1, 1, 1, enhancedHeaders.length).setValues([enhancedHeaders]);
  sheet.getRange(1, 1, 1, enhancedHeaders.length).setFontWeight('bold').setBackground('#E8F5E9');
  sheet.setFrozenRows(1);
  
  // フローデータを処理して書き込み
  const processedData = [];
  
  if (Array.isArray(flowRows)) {
    for (let i = 0; i < flowRows.length; i++) {
      const row = flowRows[i];
      
      if (typeof row === 'object' && row !== null) {
        const workContent = row['作業内容'] || '';
        const actions = splitIntoActions(workContent);
        const processName = row['工程'] || '';
        const timing = row['実施タイミング'] || '';
        const dept = row['部署'] || '';
        const condition = row['条件分岐'] || '';
        
        // ガバナンス情報を取得
        const taskDescription = `${processName} ${workContent}`;
        const disclosureCheck = checkDisclosureRequirement(taskDescription);
        const advisors = determineRequiredAdvisors(taskDescription);
        const classification = classifyTaskMECE(taskDescription);
        
        for (let j = 0; j < actions.length; j++) {
          const rowArray = [];
          
          for (const header of headers) {
            let value = '';
            
            if (header === '作業内容') {
              value = actions[j];
            } else if (header === '法令・規制') {
              value = checkLegalRegulations(processName, actions[j], timing, dept);
            } else if (header === '内部統制の観点') {
              value = checkInternalControl(processName, actions[j], condition, dept);
            } else if (header === 'コンプライアンス留意点') {
              value = generateComplianceNotes(processName, actions[j], timing, dept, condition);
            } else if (row.hasOwnProperty(header)) {
              value = j === 0 ? row[header] : '';
            } else {
              value = '';
            }
            
            rowArray.push(value || '');
          }
          
          // ガバナンス情報を追加
          rowArray.push(disclosureCheck.requiresDisclosure ? '要開示' : '');
          rowArray.push(advisors.length > 0 ? advisors.map(a => a.type).join(', ') : '');
          rowArray.push(classification.priority || '');
          
          processedData.push(rowArray);
        }
      }
    }
    
    // データを書き込み
    if (processedData.length > 0) {
      sheet.getRange(2, 1, processedData.length, enhancedHeaders.length).setValues(processedData);
      sheet.getRange(2, 1, processedData.length, enhancedHeaders.length).setWrap(true);
      
      // 優先度による色分け
      for (let i = 0; i < processedData.length; i++) {
        const priority = processedData[i][enhancedHeaders.length - 1];
        let bgColor = '#ffffff';
        
        switch(priority) {
          case 'CRITICAL': bgColor = '#f4cccc'; break;
          case 'HIGH': bgColor = '#fce5cd'; break;
          case 'MEDIUM': bgColor = '#fff2cc'; break;
          case 'LOW': bgColor = '#d9ead3'; break;
        }
        
        if (priority) {
          sheet.getRange(i + 2, enhancedHeaders.length).setBackground(bgColor);
        }
      }
    }
  }
  
  console.log('ガバナンス情報付きフロー書き込み完了');
}

/*
================================================================================
                                    終了
================================================================================
*/