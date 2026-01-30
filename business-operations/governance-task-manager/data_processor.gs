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
      this.sheet.getRange(2, 1, lastRow - 1, this.headers.length).clearContent();
    }
  }

  // 安全な書き込み実行
  performSafeWrite(data) {
    try {
      // 一括書き込みを試行
      this.sheet.getRange(2, 1, data.length, this.headers.length).setValues(data);
      console.log('一括書き込み成功');
    } catch (error) {
      console.warn('一括書き込み失敗、行ごと書き込みに切り替え:', error.message);
      this.writeRowByRow(data);
    }
  }

  // 行ごとの書き込み
  writeRowByRow(data) {
    for (let i = 0; i < data.length; i++) {
      try {
        this.sheet.getRange(i + 2, 1, 1, this.headers.length).setValues([data[i]]);
      } catch (error) {
        console.error(`Row ${i + 1} 書き込みエラー:`, error.message);
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

// 改善されたフロー行書き込み関数
function writeFlowRowsImproved(flowRows) {
  const sheetName = getConfig('FLOW_SHEET_NAME') || FLOW_SHEET;
  const sheet = ss().getSheetByName(sheetName) || createFlowSheet(sheetName);
  const headers = [
    '工程', '実施タイミング', '部署', '担当役割', '作業内容', 
    '条件分岐', '利用ツール', 'URLリンク', '備考'
  ];

  console.log('改善されたフロー行書き込み開始');
  console.log('入力データ:', typeof flowRows, Array.isArray(flowRows), flowRows?.length);

  const writer = new SafeSpreadsheetWriter(sheet, headers);
  writer.writeData(flowRows);

  logActivity('WRITE_FLOW_IMPROVED', `Successfully written ${flowRows.length} flow rows`);
}
