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
    const row = currentRow;
    
    // ボックスのセル範囲を取得
    const boxRange = visualSheet.getRange(row, col, boxHeight, boxWidth);
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