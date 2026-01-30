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
  // 新しいエンジンが利用可能な場合はそちらを使用
  if (typeof writeFlowRowsImproved === 'function') {
    console.log('レガシー関数から改善版にリダイレクト');
    return writeFlowRowsImproved(flowRows);
  }
  
  // 以下はフォールバック処理
  const sheetName = getConfig('FLOW_SHEET_NAME') || FLOW_SHEET;
  const sh = ss().getSheetByName(sheetName) || createFlowSheet(sheetName);
  
  // 既存データをクリア（ヘッダーは残す）
  if (sh.getLastRow() > 1) {
    sh.getRange(2, 1, sh.getLastRow() - 1, FLOW_HEADERS.length).clearContent();
  }
  
  // データがない場合は終了
  if (!flowRows || flowRows.length === 0) {
    logActivity('WRITE_FLOW', 'No flow rows to write');
    return;
  }
  
  // データマッピング（すべての値を文字列として扱う）
  console.log('flowRows の型と構造:', typeof flowRows, Array.isArray(flowRows), flowRows.length);
  console.log('flowRows サンプル:', flowRows.slice(0, 2));
  
  const values = flowRows.map((row, rowIndex) => {
    console.log(`Row ${rowIndex + 1} の型:`, typeof row, Array.isArray(row));
    console.log(`Row ${rowIndex + 1} の内容:`, row);
    
    const mappedRow = FLOW_HEADERS.map((header, colIndex) => {
      const value = row[header] || '';
      const stringValue = String(value).trim();
      
      // 末尾の数字を除去（OpenAIレスポンスの問題対応）
      let cleanValue = stringValue;
      if (/\d+$/.test(stringValue) && stringValue.length > 1) {
        const withoutNumber = stringValue.replace(/\d+$/, '').trim();
        if (withoutNumber.length > 0) {
          console.log(`末尾の数字を除去: "${stringValue}" -> "${withoutNumber}"`);
          cleanValue = withoutNumber;
        }
      }
      
      // デバッグ用：問題のあるデータをログ出力
      if (cleanValue.includes('3') && colIndex === 8) { // 備考列
        console.log(`Row ${rowIndex + 1}, Column ${header}: "${cleanValue}"`);
      }
      
      return cleanValue;
    });
    
    console.log(`Row ${rowIndex + 1} マッピング結果:`, mappedRow);
    return mappedRow;
  });
  
  // データ書き込み
  try {
    console.log(`書き込み予定データ: ${values.length}行 x ${FLOW_HEADERS.length}列`);
    console.log('データ構造確認:', typeof values, Array.isArray(values));
    
    // データ構造の詳細チェック
    if (!Array.isArray(values)) {
      throw new Error(`valuesが配列ではありません: ${typeof values}`);
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
        const splitRow = stringRow.split(',');
        // 列数を調整
        while (splitRow.length < FLOW_HEADERS.length) {
          splitRow.push('');
        }
        return splitRow.slice(0, FLOW_HEADERS.length);
      }
    });
    
    console.log('安全な形式に変換後:', safeValues.slice(0, 2));
    sh.getRange(2, 1, safeValues.length, FLOW_HEADERS.length).setValues(safeValues);
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
            rowData = String(rowData).split(',');
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
            rowData = String(rowData).split(',');
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