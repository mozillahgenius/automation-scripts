/**
 * スプレッドシート業務フロー自動生成スクリプト（汎用版）
 * セルの背景色、罫線、結合を使用して視覚的なフローチャートを作成
 * 
 * 修正履歴:
 * - 関連資料（URL）を正しいタイミングと関連付けて表示するよう修正
 * - 参考資料が業務フローの上部に配置される問題を解決
 */

// 定数定義
const SHEET_NAMES = {
  FLOW: "フロー",
  VISUAL: "ビジュアルフロー"
};

const REQUIRED_HEADERS = ["工程", "実施タイミング", "部署", "担当役割", "作業内容", "条件分岐", "利用ツール"];

// 基本カラーパレット
const COLORS = {
  HEADER: "#4A5568",
  SUBHEADER: "#718096",
  TIMELINE: "#F7FAFC",
  PROCESS: "#87CEEB",
  DECISION: "#FFD700",
  START_END: "#90EE90",
  EMPTY: "#FAFAFA",
  DATASOURCE: "#E3F2FD",
  TOOL_BG: "#EDF2F7"
};

// 部署カラーパレット（自動割り当て用）
const DEPT_COLORS = [
  "#E3F2FD", // 薄い青
  "#FFF3E0", // 薄いオレンジ
  "#FCE4EC", // 薄いピンク
  "#F3E5F5", // 薄い紫
  "#E8F5E9", // 薄い緑
  "#FFF8E1", // 薄い黄
  "#E0F2F1", // 薄いシアン
  "#FAFAFA", // 薄いグレー
  "#FBE9E7", // 薄い赤
  "#E8EAF6"  // 薄い藍
];

// ツールカラー（汎用的なツール名に対応）
const TOOL_COLORS = {
  "Word": "#2B579A",
  "Excel": "#217346",
  "PPT": "#D24726",
  "PowerPoint": "#D24726",
  "Teams": "#5B5FC7",
  "Outlook": "#0078D4",
  "Photoshop": "#31A8FF",
  "CMS": "#FF6B6B",
  "ブラウザ": "#4CAF50",
  "Browser": "#4CAF50",
  "メール": "#EA4335",
  "Email": "#EA4335",
  "Slack": "#4A154B",
  "GitHub": "#24292E",
  "Jira": "#0052CC",
  "Notion": "#000000",
  "Google Drive": "#4285F4",
  "Zoom": "#2D8CFF"
};

// カスタムメニューの追加
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('フローチャート生成')
    .addItem('初回セットアップ', 'setupSystem')
    .addSeparator()
    .addItem('ビジュアルフロー生成', 'generateVisualFlow')
    .addItem('フロー設定', 'showFlowSettings')
    .addSeparator()
    .addItem('サンプルデータ作成', 'createSampleData')
    .addToUi();
}

// フロー設定ダイアログ
function showFlowSettings() {
  const html = HtmlService.createHtmlOutput(`
    <h3>フロー設定</h3>
    <p>フローのタイトルを入力してください：</p>
    <input type="text" id="flowTitle" value="業務フロー" style="width: 300px;">
    <br><br>
    <button onclick="google.script.run.setFlowTitle(document.getElementById('flowTitle').value); google.script.host.close();">
      設定して閉じる
    </button>
  `)
  .setWidth(400)
  .setHeight(200);
  
  SpreadsheetApp.getUi().showModalDialog(html, 'フロー設定');
}

// フロータイトルの設定
function setFlowTitle(title) {
  PropertiesService.getDocumentProperties().setProperty('flowTitle', title);
  SpreadsheetApp.getUi().alert('設定完了', 'フロータイトルを「' + title + '」に設定しました。', SpreadsheetApp.getUi().ButtonSet.OK);
}

// 初回セットアップ
function setupSystem() {
  try {
    createSampleData();
    getOrCreateSheet(SHEET_NAMES.VISUAL);
    SpreadsheetApp.getUi().alert('成功', 'セットアップが完了しました。', SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (error) {
    SpreadsheetApp.getUi().alert('エラー', 'セットアップ中にエラーが発生しました: ' + error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

// ビジュアルフロー生成
function generateVisualFlow() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.FLOW);
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
    const visualSheet = getOrCreateSheet(SHEET_NAMES.VISUAL);
    visualSheet.clear();
    
    // データの解析
    const flowData = parseFlowData(data);
    
    // フローチャートの描画
    drawFlowChart(visualSheet, flowData);
    
    SpreadsheetApp.getUi().alert('成功', 'ビジュアルフローが生成されました。', SpreadsheetApp.getUi().ButtonSet.OK);
    
  } catch (error) {
    console.error('エラー:', error);
    SpreadsheetApp.getUi().alert('エラー', 'フロー生成中にエラーが発生しました: ' + error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

// フローデータの解析
function parseFlowData(data) {
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
    datasources: {},  // タイミングごとの関連資料を管理
    processName: ""
  };
  
  // プロセス名の取得（最初のデータ行の工程名）
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
    
    // 部署の初期化とリスト作成
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
      condition: row[columnIndex["条件分岐"]] || "",
      tool: tool,
      url: url,
      note: row[columnIndex["備考"]] || ""
    });
    
    // ツールの収集
    if (tool && tool !== "-") {
      // 複数ツールの場合は分割
      const tools = tool.split(/[／、,]/);
      tools.forEach(t => {
        const trimmedTool = t.trim();
        if (trimmedTool) {
          flowData.tools.add(trimmedTool);
        }
      });
    }
    
    // データソース（URL）をタイミングごとに管理
    if (url && url !== "-") {
      if (!flowData.datasources[timing]) {
        flowData.datasources[timing] = [];
      }
      // 同じURLが既に登録されていなければ追加
      if (!flowData.datasources[timing].includes(url)) {
        flowData.datasources[timing].push(url);
      }
    }
  }
  
  return flowData;
}

// フローチャートの描画
function drawFlowChart(sheet, flowData) {
  let currentRow = 1;
  
  // タイトル行
  const flowTitle = PropertiesService.getDocumentProperties().getProperty('flowTitle') || 
                    flowData.processName || 
                    "業務フロー";
  
  const maxCols = Math.max(flowData.departmentList.length + 2, 10);
  sheet.getRange(currentRow, 1, 1, maxCols).merge();
  const titleCell = sheet.getRange(currentRow, 1);
  titleCell.setValue(flowTitle + "（" + new Date().getFullYear() + "年" + (new Date().getMonth() + 1) + "月〜）");
  titleCell.setBackground(COLORS.HEADER);
  titleCell.setFontColor("#FFFFFF");
  titleCell.setFontSize(16);
  titleCell.setFontWeight("bold");
  titleCell.setHorizontalAlignment("center");
  titleCell.setVerticalAlignment("middle");
  sheet.setRowHeight(currentRow, 50);
  currentRow++;
  
  // ツール行
  if (flowData.tools.size > 0) {
    drawToolsRow(sheet, currentRow, flowData.tools, maxCols);
    currentRow++;
  }
  
  // ヘッダー行
  drawHeaderRow(sheet, currentRow, flowData.departmentList, Object.keys(flowData.datasources).length > 0);
  currentRow++;
  
  // 開始行
  drawStartRow(sheet, currentRow, maxCols);
  currentRow++;
  
  // 各タイミングの行
  flowData.timings.forEach((timing, index) => {
    drawTimingRow(sheet, currentRow, timing, flowData, index);
    currentRow++;
  });
  
  // 終了行
  drawEndRow(sheet, currentRow, flowData.departmentList.length, Object.keys(flowData.datasources).length > 0);
  currentRow++;
  
  // 凡例行
  drawLegendRow(sheet, currentRow, maxCols);
  
  // 列幅の調整
  adjustColumnWidths(sheet, flowData.departmentList.length, Object.keys(flowData.datasources).length > 0);
  
  // 罫線の設定
  applyBorders(sheet, currentRow, maxCols);
}

// ツール行の描画
function drawToolsRow(sheet, row, tools, maxCols) {
  sheet.getRange(row, 1).setValue("使用ツール");
  sheet.getRange(row, 1).setBackground("#2D3748");
  sheet.getRange(row, 1).setFontColor("#FFFFFF");
  sheet.getRange(row, 1).setFontWeight("bold");
  
  sheet.getRange(row, 2, 1, maxCols - 1).merge();
  const toolCell = sheet.getRange(row, 2);
  toolCell.setBackground(COLORS.TOOL_BG);
  
  let toolText = "";
  tools.forEach(tool => {
    const color = TOOL_COLORS[tool] || "#666666";
    toolText += ` [${tool}] `;
  });
  
  toolCell.setValue(toolText);
  toolCell.setHorizontalAlignment("left");
  sheet.setRowHeight(row, 35);
}

// ヘッダー行の描画
function drawHeaderRow(sheet, row, departments, hasDataSource) {
  // 日程列
  sheet.getRange(row, 1).setValue("日程");
  sheet.getRange(row, 1).setBackground("#2D3748");
  sheet.getRange(row, 1).setFontColor("#FFFFFF");
  sheet.getRange(row, 1).setFontWeight("bold");
  
  // 部署列（動的に色を割り当て）
  departments.forEach((dept, index) => {
    const col = index + 2;
    const cell = sheet.getRange(row, col);
    
    // 部署名と役割を分離して表示
    const displayName = dept;
    
    cell.setValue(displayName);
    cell.setBackground(DEPT_COLORS[index % DEPT_COLORS.length]);
    cell.setFontWeight("bold");
    cell.setWrap(true);
    cell.setHorizontalAlignment("center");
    cell.setVerticalAlignment("middle");
  });
  
  // データソース列
  if (hasDataSource) {
    const dataCol = departments.length + 2;
    sheet.getRange(row, dataCol).setValue("関連資料");
    sheet.getRange(row, dataCol).setBackground("#CBD5E0");
    sheet.getRange(row, dataCol).setFontWeight("bold");
  }
  
  sheet.setRowHeight(row, 50);
}

// 開始行の描画
function drawStartRow(sheet, row, maxCols) {
  const cell = sheet.getRange(row, 1);
  cell.setValue("【開始】");
  cell.setBackground(COLORS.START_END);
  cell.setFontWeight("bold");
  cell.setBorder(true, true, true, true, false, false, "#228B22", SpreadsheetApp.BorderStyle.SOLID_THICK);
  
  sheet.setRowHeight(row, 40);
}

// タイミング行の描画
function drawTimingRow(sheet, row, timing, flowData, timingIndex) {
  // タイミング列
  const timingCell = sheet.getRange(row, 1);
  timingCell.setValue(timing);
  timingCell.setBackground(COLORS.TIMELINE);
  timingCell.setFontWeight("bold");
  timingCell.setWrap(true);
  
  // 各部署のタスク
  flowData.departmentList.forEach((dept, deptIndex) => {
    const col = deptIndex + 2;
    const tasks = flowData.departments[dept][timing];
    
    if (tasks && tasks.length > 0) {
      const task = tasks[0]; // 最初のタスクを表示
      const cell = sheet.getRange(row, col);
      
      // タスク内容の設定
      let content = task.task;
      if (task.role && task.role !== "-") {
        content = "【" + task.role + "】\n" + content;
      }
      if (task.tool && task.tool !== "-") {
        content += "\n(" + task.tool + ")";
      }
      
      cell.setValue(content);
      cell.setWrap(true);
      cell.setHorizontalAlignment("center");
      cell.setVerticalAlignment("middle");
      
      // スタイルの設定
      if (task.condition && task.condition !== "-") {
        // 判断ボックス
        cell.setBackground(COLORS.DECISION);
        cell.setBorder(true, true, true, true, false, false, "#FF8C00", SpreadsheetApp.BorderStyle.SOLID_THICK);
        cell.setFontWeight("bold");
        
        // 条件を注記として追加
        const noteContent = "条件分岐: " + task.condition + 
                          (task.note ? "\n備考: " + task.note : "");
        cell.setNote(noteContent);
      } else {
        // プロセスボックス
        cell.setBackground(COLORS.PROCESS);
        cell.setBorder(true, true, true, true, false, false, "#4682B4", SpreadsheetApp.BorderStyle.SOLID_THICK);
        
        // 備考があれば注記として追加
        if (task.note) {
          cell.setNote("備考: " + task.note);
        }
      }
      
      // 矢印の追加（次のタイミングがある場合）
      if (timingIndex < flowData.timings.length - 1) {
        // 同じ部署に次のタスクがあるか確認
        const nextTiming = flowData.timings[timingIndex + 1];
        if (flowData.departments[dept][nextTiming]) {
          addArrowToCell(cell, "↓");
        }
      }
      
    } else {
      // 空のセル
      const cell = sheet.getRange(row, col);
      cell.setBackground(COLORS.EMPTY);
    }
  });
  
  // データソース（タイミングに対応した関連資料を表示）
  const hasDataSource = Object.keys(flowData.datasources).length > 0;
  if (hasDataSource) {
    const dataCol = flowData.departmentList.length + 2;
    const dataCell = sheet.getRange(row, dataCol);
    
    // 現在のタイミングに対応する関連資料を取得
    if (flowData.datasources[timing] && flowData.datasources[timing].length > 0) {
      // 複数のURLがある場合は改行で区切って表示
      const urls = flowData.datasources[timing].join("\n");
      dataCell.setValue("📄 " + urls);
      dataCell.setBackground(COLORS.DATASOURCE);
      dataCell.setBorder(true, true, true, true, false, false, "#2196F3", SpreadsheetApp.BorderStyle.DASHED);
      dataCell.setWrap(true);
      dataCell.setHorizontalAlignment("center");
      dataCell.setVerticalAlignment("middle");
    } else {
      // このタイミングに関連資料がない場合は空欄
      dataCell.setValue("");
      dataCell.setBackground(COLORS.EMPTY);
    }
  }
  
  sheet.setRowHeight(row, 90);
}

// 終了行の描画
function drawEndRow(sheet, row, deptCount, hasDataSource) {
  const cell = sheet.getRange(row, 1);
  cell.setValue("【完了】");
  cell.setBackground(COLORS.START_END);
  cell.setFontWeight("bold");
  cell.setBorder(true, true, true, true, false, false, "#228B22", SpreadsheetApp.BorderStyle.SOLID_THICK);
  
  // 完了メッセージ
  const mergeCols = hasDataSource ? deptCount + 1 : deptCount;
  sheet.getRange(row, 2, 1, mergeCols).merge();
  const msgCell = sheet.getRange(row, 2);
  msgCell.setValue("✅ プロセス完了");
  msgCell.setBackground("#E8F5E9");
  msgCell.setFontSize(16);
  msgCell.setFontWeight("bold");
  msgCell.setHorizontalAlignment("center");
  
  sheet.setRowHeight(row, 50);
}

// 凡例行の描画
function drawLegendRow(sheet, row, maxCols) {
  sheet.getRange(row, 1, 1, maxCols).merge();
  const legendCell = sheet.getRange(row, 1);
  legendCell.setValue("凡例： □ 処理・作業　◆ 判断・分岐　→ 処理の流れ　📄 関連資料　※セルの注記（メモ）に詳細情報があります");
  legendCell.setBackground(COLORS.TIMELINE);
  legendCell.setFontWeight("bold");
  legendCell.setHorizontalAlignment("left");
  legendCell.setVerticalAlignment("middle");
  sheet.setRowHeight(row, 40);
}

// セルに矢印を追加
function addArrowToCell(cell, arrow) {
  const currentValue = cell.getValue();
  const richText = SpreadsheetApp.newRichTextValue()
    .setText(currentValue + "\n" + arrow)
    .setTextStyle(currentValue.length + 1, currentValue.length + arrow.length + 1, 
      SpreadsheetApp.newTextStyle()
        .setForegroundColor("#4682B4")
        .setFontSize(16)
        .setBold(true)
        .build())
    .build();
  cell.setRichTextValue(richText);
}

// 列幅の調整
function adjustColumnWidths(sheet, deptCount, hasDataSource) {
  sheet.setColumnWidth(1, 150); // タイムライン列
  for (let i = 2; i <= deptCount + 1; i++) {
    sheet.setColumnWidth(i, 200); // 部署列
  }
  if (hasDataSource) {
    sheet.setColumnWidth(deptCount + 2, 150); // データソース列
  }
}

// 罫線の設定
function applyBorders(sheet, lastRow, lastCol) {
  const range = sheet.getRange(1, 1, lastRow, lastCol);
  range.setBorder(true, true, true, true, true, true, "#d0d0d0", SpreadsheetApp.BorderStyle.SOLID);
}

// シートの取得または作成
function getOrCreateSheet(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }
  
  return sheet;
}

// サンプルデータの作成（汎用的な内容に変更）
function createSampleData() {
  const sheet = getOrCreateSheet(SHEET_NAMES.FLOW);
  sheet.clear();
  
  const headers = ["工程", "実施タイミング", "部署", "担当役割", "作業内容", "条件分岐", "利用ツール", "URLリンク", "備考"];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  const sampleData = [
    ["承認プロセス", "申請時", "営業部", "申請者", "申請書作成・提出", "-", "Word／メール", "申請書テンプレート", "-"],
    ["承認プロセス", "1営業日以内", "営業部", "部門長", "申請内容確認", "承認可否判定", "Teams／メール", "-", "金額により承認権限が異なる"],
    ["承認プロセス", "2営業日以内", "経理部", "経理担当", "予算確認", "予算内/予算超過", "Excel", "予算管理表", "-"],
    ["承認プロセス", "3営業日以内", "総務部", "総務担当", "最終承認処理", "-", "承認システム", "-", "承認後の場合"],
    ["承認プロセス", "承認後即日", "営業部", "申請者", "承認通知受領・実行", "-", "メール", "-", "-"]
  ];
  
  sheet.getRange(2, 1, sampleData.length, headers.length).setValues(sampleData);
  
  // 書式設定
  sheet.autoResizeColumns(1, headers.length);
  sheet.getRange(1, 1, 1, headers.length).setBackground("#4285F4").setFontColor("#FFFFFF").setFontWeight("bold");
  
  SpreadsheetApp.getUi().alert("サンプルデータが作成されました。");
}