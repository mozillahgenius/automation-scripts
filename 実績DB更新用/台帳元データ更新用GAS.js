function processAll() {
    importExcelSheetToTarget(0, '台帳自動連携_1', 7, 8, 7);
    importExcelSheetToTarget(1, '台帳自動連携_2', 10, 11, 10);
    
    // まず転記シート（全データ）を作成
    mergeSheetsTo転記シート();
    
    // 転記シートから期間フィルタリングして「ソアスク+台帳データ」に出力
    filterFromTransferSheet();
  }
  
  // 処理結果を確認する関数
  function verifyProcessingResults() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('転記シート（全データ）');
    
    if (!sheet) {
      logUsage('エラー: 転記シート（全データ）が見つかりません');
      return;
    }
    
    logUsage(`=== 処理結果の検証 ===`);
    
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    
    if (lastRow >= 2 && lastCol >= 18) {
      // M列の名前確認
      const mColumnData = sheet.getRange(2, 13, Math.min(5, lastRow - 1), 1).getValues();
      logUsage(`\nM列（名前）のサンプル（空白削除後）:`);
      mColumnData.forEach((row, index) => {
        if (row[0]) {
          logUsage(`  行${index + 2}: "${row[0]}"`);
        }
      });
      
      // N列の計上日確認
      const nColumnData = sheet.getRange(2, 14, Math.min(5, lastRow - 1), 1).getValues();
      logUsage(`\nN列（計上日）のサンプル（形式変換後）:`);
      nColumnData.forEach((row, index) => {
        if (row[0]) {
          logUsage(`  行${index + 2}: "${row[0]}"`);
        }
      });
      
      // S列の売上計上日確認
      const sColumnData = sheet.getRange(2, 19, Math.min(5, lastRow - 1), 1).getValues();
      logUsage(`\nS列（売上計上日）のサンプル（形式変換後）:`);
      sColumnData.forEach((row, index) => {
        if (row[0]) {
          logUsage(`  行${index + 2}: "${row[0]}"`);
        }
      });
    }
  }
  
  // 新しい関数：転記シートから期間フィルタリングして「ソアスク+台帳データ」へ
  function filterFromTransferSheet() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sourceSheet = ss.getSheetByName('転記シート（全データ）');
    
    if (!sourceSheet) {
      logUsage('エラー: 転記シート（全データ）が見つかりません');
      return;
    }
    
    // フィルタリング後のシート（ソアスク+台帳データ）を取得または作成
    let targetSheet = ss.getSheetByName('ソアスク+台帳データ');
    if (!targetSheet) {
      targetSheet = ss.insertSheet('ソアスク+台帳データ');
    } else {
      // 1行目（ヘッダー）を保持して、2行目以降をクリア
      const lastRow = targetSheet.getLastRow();
      const lastCol = targetSheet.getLastColumn();
      
      // ヘッダーを保存
      let headerRow = null;
      if (lastRow >= 1 && lastCol >= 1) {
        headerRow = targetSheet.getRange(1, 1, 1, lastCol).getValues()[0];
      }
      
      // 2行目以降にデータがある場合はクリア
      if (lastRow >= 2) {
        targetSheet.getRange(2, 1, lastRow - 1, lastCol).clearContent();
        targetSheet.getRange(2, 1, lastRow - 1, lastCol).clearFormat();
      }
      
      logUsage(`既存のヘッダー行を保持`);
    }
    
    // 期間の設定
    const startDate = new Date('2024/12/01');
    const endDate = new Date('2025/11/30');
    endDate.setHours(23, 59, 59, 999); // 終了日の最後の時刻まで含める
    
    logUsage(`=== 期間フィルタリング開始 ===`);
    logUsage(`対象期間: ${startDate.toLocaleDateString()} 〜 ${endDate.toLocaleDateString()}`);
    
    // データを効率的に取得
    const lastRow = sourceSheet.getLastRow();
    const lastCol = sourceSheet.getLastColumn();
    
    logUsage(`転記シートのサイズ: ${lastRow}行 x ${lastCol}列`);
    
    if (lastRow <= 1) {
      logUsage('データがありません（ヘッダーのみ）');
      return;
    }
    
    // データを一括で取得（1行目はヘッダーなので、2行目から取得）
    const data = sourceSheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
    
    // M列（インデックス12）に計上日があることを確認
    const dateColumnIndex = 12; // M列は13列目（0ベースで12）
    
    logUsage(`計上日列: M列（インデックス: ${dateColumnIndex}）`);
    
    // 背景色も取得（2行目から）
    const backgrounds = sourceSheet.getRange(2, 1, lastRow - 1, lastCol).getBackgrounds();
    
    // フィルタリングされたデータを格納
    const filteredData = [];
    const filteredBackgrounds = [];
    let filteredCount = 0;
    let skippedCount = 0;
    let emptyDateCount = 0;
    let parseFailCount = 0;
    let outOfRangeCount = 0;
    let debugCount = 0;
    
    // データ行をフィルタリング
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      
      // 空行はスキップ
      if (!row || row.every(cell => cell === '' || cell === null)) {
        continue;
      }
      
      // デバッグ：最初の10件の日付情報を詳細にログ出力
      if (debugCount < 10 && row[dateColumnIndex] !== '') {
        const rawDate = row[dateColumnIndex];
        const dateType = typeof rawDate;
        
        logUsage(`デバッグ[${debugCount}] 行${i+2}:`);  // 実際の行番号は+2（ヘッダー分）
        logUsage(`  - 元の値: "${rawDate}" (型: ${dateType})`);
        
        if (typeof rawDate === 'string') {
          const match = rawDate.match(/(\d{4})\/(\d{2})\/(\d{2})/);
          if (match) {
            const date = new Date(parseInt(match[1]), parseInt(match[2]) - 1, parseInt(match[3]));
            const inRange = date >= startDate && date <= endDate;
            logUsage(`  - 日付: ${date.toLocaleDateString()}`);
            logUsage(`  - 期間判定: ${inRange ? '期間内' : '期間外'}`);
          } else {
            logUsage(`  - 日付形式が不正`);
          }
        } else if (rawDate instanceof Date) {
          const inRange = rawDate >= startDate && rawDate <= endDate;
          logUsage(`  - 日付: ${rawDate.toLocaleDateString()}`);
          logUsage(`  - 期間判定: ${inRange ? '期間内' : '期間外'}`);
        }
        debugCount++;
      }
      
      // 計上日をチェック（N列）
      if (row.length <= dateColumnIndex || row[dateColumnIndex] === '' || row[dateColumnIndex] === null) {
        // 計上日が空の場合
        emptyDateCount++;
        skippedCount++;
      } else {
        const dateValue = row[dateColumnIndex];
        let date = null;
        
        // 文字列形式の日付（yyyy/MM/dd）をパース
        if (typeof dateValue === 'string') {
          const match = dateValue.match(/(\d{4})\/(\d{2})\/(\d{2})/);
          if (match) {
            date = new Date(parseInt(match[1]), parseInt(match[2]) - 1, parseInt(match[3]));
          } else {
            // その他の形式の場合はparseDate関数を使用
            date = parseDate(dateValue);
          }
        } else if (dateValue instanceof Date) {
          date = dateValue;
        } else {
          date = parseDate(dateValue);
        }
        
        if (!date || isNaN(date.getTime())) {
          // パース失敗
          parseFailCount++;
          skippedCount++;
        } else if (date >= startDate && date <= endDate) {
          // 期間内
          filteredData.push(row);
          filteredBackgrounds.push(backgrounds[i]);
          filteredCount++;
        } else {
          // 期間外
          outOfRangeCount++;
          skippedCount++;
        }
      }
      
      // 定期的にメモリ解放
      if (i % 1000 === 0) {
        SpreadsheetApp.flush();
      }
    }
    
    logUsage(`=== フィルタリング結果 ===`);
    logUsage(`処理した総行数: ${data.length}`);
    logUsage(`対象期間内: ${filteredCount}件`);
    logUsage(`期間外: ${outOfRangeCount}件`);
    logUsage(`計上日が空: ${emptyDateCount}件`);
    logUsage(`日付パース失敗: ${parseFailCount}件`);
    logUsage(`合計スキップ: ${skippedCount}件`);
    
    // フィルタリングしたデータを「ソアスク+台帳データ」シートに出力
    if (filteredData.length > 0) {
      logUsage(`${filteredData.length}行を出力開始（2行目から）`);
      
      try {
        // データを2行目から書き込み（1行目はヘッダーを保持）
        targetSheet.getRange(2, 1, filteredData.length, filteredData[0].length).setValues(filteredData);
        logUsage(`データの書き込み完了`);
        
        // 背景色も効率的に適用（2行目から）
        if (filteredBackgrounds.length > 0) {
          targetSheet.getRange(2, 1, filteredBackgrounds.length, filteredBackgrounds[0].length).setBackgrounds(filteredBackgrounds);
          logUsage(`背景色の適用完了`);
        }
        
        // N列（売上金額）を日本円の通貨形式に設定
        if (filteredData.length > 0) {
          targetSheet
            .getRange(2, 14, filteredData.length, 1) // N列（2行目から）
            .setNumberFormat('[$¥-ja-JP]#,##0;[Red]-[$¥-ja-JP]#,##0');
        }
      } catch (e) {
        logUsage(`エラー: データ書き込み中にエラーが発生しました: ${e.message}`);
      }
    } else {
      logUsage(`警告: フィルタリング結果が0件です`);
    }
    
    logUsage(`フィルタリング完了: ${filteredCount}件のデータを「ソアスク+台帳データ」に出力`);
    logUsage(`=== 期間フィルタリング終了 ===`);
  }
  
  // 日付をパースする補助関数（改良版）
  function parseDate(dateValue) {
    if (!dateValue) return null;
    
    // すでにDateオブジェクトの場合
    if (dateValue instanceof Date) {
      return dateValue;
    }
    
    // 数値の場合（Excelのシリアル値の可能性）
    if (typeof dateValue === 'number') {
      // Excelのシリアル値を日付に変換（1900年1月1日を基準）
      // JavaScriptは1970年1月1日を基準とするため、調整が必要
      const excelEpoch = new Date(1899, 11, 30); // 1899年12月30日
      const msPerDay = 24 * 60 * 60 * 1000;
      const date = new Date(excelEpoch.getTime() + dateValue * msPerDay);
      if (!isNaN(date.getTime())) {
        return date;
      }
    }
    
    // 文字列の場合
    const dateStr = dateValue.toString().trim();
    if (!dateStr) return null;
    
    // 一般的な日付フォーマットを試す
    const patterns = [
      /(\d{4})[\/\-年](\d{1,2})[\/\-月](\d{1,2})日?/,  // 2024/12/01, 2024-12-01, 2024年12月01日
      /(\d{1,2})[\/\-月](\d{1,2})日?[\/\-,\s]+(\d{4})年?/,  // 12/01/2024, 12月01日, 2024年
      /令和(\d+)年(\d{1,2})月(\d{1,2})日?/  // 令和6年12月1日
    ];
    
    for (const pattern of patterns) {
      const match = dateStr.match(pattern);
      if (match) {
        let year, month, day;
        
        if (pattern.source.includes('令和')) {
          // 令和の場合
          year = 2018 + parseInt(match[1]);
          month = parseInt(match[2]);
          day = parseInt(match[3]);
        } else if (match[1].length === 4) {
          // 年が最初
          year = parseInt(match[1]);
          month = parseInt(match[2]);
          day = parseInt(match[3]);
        } else {
          // 年が最後
          year = parseInt(match[3]);
          month = parseInt(match[1]);
          day = parseInt(match[2]);
        }
        
        const date = new Date(year, month - 1, day);
        if (!isNaN(date.getTime())) {
          return date;
        }
      }
    }
    
    // 最後の手段として、Dateコンストラクタに任せる
    const date = new Date(dateStr);
    if (!isNaN(date.getTime())) {
      return date;
    }
    
    return null;
  }
  
  // 列インデックスを列文字に変換
  function columnIndexToLetter(index) {
    let letter = '';
    while (index >= 0) {
      letter = String.fromCharCode(65 + (index % 26)) + letter;
      index = Math.floor(index / 26) - 1;
    }
    return letter;
  }
  
  // 既存の関数はそのまま維持
  function importExcelSheetToTarget(sheetIndex, targetSheetName, mappingRowHeader, mappingRowMap, startRowInExcel) {
    const excelFileId = '1wyr5nJJc-4HFcXPpf3T5PL-fbQehSlVO';
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const outputSheet = ss.getSheetByName(targetSheetName);
    const mappingSheet = ss.getSheetByName('マッピング定義');
  
    try {
      logUsage(`=== ${targetSheetName}：処理開始 ===`);
  
      const blob = DriveApp.getFileById(excelFileId).getBlob();
      const tempFile = Drive.Files.insert({
        title: 'TempExcelConverted',
        mimeType: MimeType.GOOGLE_SHEETS
      }, blob);
      const tempSS = SpreadsheetApp.openById(tempFile.id);
      const sourceSheet = tempSS.getSheets()[sheetIndex];
  
      const totalRows = sourceSheet.getLastRow() - startRowInExcel + 1;
      const totalCols = sourceSheet.getLastColumn();
      const raw = sourceSheet.getRange(startRowInExcel, 1, totalRows, totalCols).getValues();
      const headers = raw[0];
      const originalData = raw.slice(1);
  
      // マッピング定義のB列以降を読み取る（A列を無視）
      const headerMapRow = mappingSheet.getRange(mappingRowHeader, 2, 1, mappingSheet.getLastColumn() - 1).getValues()[0];
      const targetMapRow = mappingSheet.getRange(mappingRowMap, 2, 1, mappingSheet.getLastColumn() - 1).getValues()[0];
  
      const columnMap = [];
      for (let i = 0; i < targetMapRow.length; i++) {
        const mapping = targetMapRow[i];
        if (mapping && typeof mapping === 'string') {
          const targets = mapping.split(',').map(c => c.trim()).filter(c => /^[A-Z]+$/.test(c));
          if (targets.length > 0) {
            const toIndexes = targets.map(columnLetterToIndex);
            columnMap.push({ fromIndex: i + 1, toIndexes, label: headerMapRow[i] });  // B列から読んでいるのでi+1のまま
          }
        }
      }
  
      const maxCol = Math.max(...columnMap.flatMap(m => m.toIndexes)) + 1;
      const transformedData = [];
      for (let i = 0; i < originalData.length; i++) {
        const newRow = Array(maxCol).fill('');
        columnMap.forEach(({ fromIndex, toIndexes }) => {
          toIndexes.forEach(toCol => {
            let val = originalData[i][fromIndex];
            newRow[toCol] = val;
          });
        });
  
        const isEmpty = newRow.every(cell => cell === '' || cell === null);
        const isZeroOnly = newRow.every(cell => cell === 0 || cell === '0' || cell === '' || cell === null);
        if (!isEmpty && !isZeroOnly) {
          transformedData.push(newRow);
        }
      }
  
      const headerRowOut = Array(maxCol).fill('');
      columnMap.forEach(({ label, toIndexes }) => {
        toIndexes.forEach(toCol => {
          headerRowOut[toCol] = label;
        });
      });
  
      outputSheet.clearContents();
      outputSheet.getRange(1, 1, 1, maxCol).setValues([headerRowOut]);
      if (transformedData.length > 0) {
        outputSheet.getRange(2, 1, transformedData.length, maxCol).setValues(transformedData);
      }
  
      logUsage(`${targetSheetName} 出力完了：${transformedData.length}行`);
      DriveApp.getFileById(tempFile.id).setTrashed(true);
    } catch (e) {
      logUsage(`エラー（${targetSheetName}）：${e.message}`);
    }
  }
  
  function mergeSheetsTo転記シート() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sourceSheetNames = ['ソアスク自動連携', '台帳自動連携_1', '台帳自動連携_2'];
    const colors = ['#fce4d6', '#d9e1f2', '#e2f0d9']; // オレンジ, 青, 緑
  
    const 転記シート = ss.getSheetByName('転記シート（全データ）') || ss.insertSheet('転記シート（全データ）');
  
    // 1行目（ヘッダー）を保持して、2行目以降をクリア
    const lastRow = 転記シート.getLastRow();
    const lastCol = 転記シート.getLastColumn();
    
    // ヘッダーを保存
    let headerRow = null;
    let headerBackgrounds = null;
    if (lastRow >= 1 && lastCol >= 1) {
      headerRow = 転記シート.getRange(1, 1, 1, lastCol).getValues()[0];
      headerBackgrounds = 転記シート.getRange(1, 1, 1, lastCol).getBackgrounds()[0];
      logUsage(`既存のヘッダー行を保持`);
    }
    
    // 2行目以降にデータがある場合はクリア
    if (lastRow >= 2) {
      転記シート.getRange(2, 1, lastRow - 1, lastCol).clearContent();
      転記シート.getRange(2, 1, lastRow - 1, lastCol).clearFormat();
    }
    
    // 全データを一時的に格納する配列
    let allData = [];
    let allBackgrounds = [];
    let maxCols = 0;
    
    const 台帳1データMap = new Map();
  
    // 列の説明（参考）- 転記シートではA列から開始
    // A列: エンドユーザ計上部門（インデックス0）
    // B列: 案件番号（インデックス1）
    // C列: 案件担当（インデックス2）
    // D列: 品目（インデックス3）
    // E列: 売上分類（インデックス4）※台帳自動連携で上書き
    // F列: 取引区分（インデックス5）※台帳自動連携で上書き
    // G列: 商流（インデックス6）※台帳自動連携で上書き
    // H列: 販売区分（インデックス7）※台帳自動連携で「通常」を設定
    // ...
    // L列: 担当者名（インデックス11）※空白を削除
    // M列: 計上日（インデックス12）※yyyy/MM/dd形式に変換
    // ...
    // R列: 売上計上日（インデックス17）※yyyy/MM/dd形式に変換
  
    for (let s = 0; s < sourceSheetNames.length; s++) {
      const sheetName = sourceSheetNames[s];
      const sheet = ss.getSheetByName(sheetName);
      if (!sheet) {
        logUsage(`⚠ シートが見つかりません: ${sheetName}`);
        continue;
      }
  
      // データを一括で取得
      const lastRow = sheet.getLastRow();
      const lastCol = sheet.getLastColumn();
      
      if (lastRow <= 1 || lastCol === 0) continue;
      
      const data = sheet.getRange(1, 1, lastRow, lastCol).getValues();
      if (data.length <= 1) continue;
  
      let body = data.slice(1); // ヘッダー除去
  
      const rowCount = body.length;
      const colCount = body[0] ? body[0].length : 0;
      
      if (rowCount > 0 && colCount > 0) {
        // 台帳自動連携の場合、A列を削除して左にシフト
        const is台帳 = sheetName === '台帳自動連携_1' || sheetName === '台帳自動連携_2';
        
        if (is台帳) {
          // A列を削除（インデックス0を除外）
          body = body.map(row => row.slice(1));
          logUsage(`${sheetName}: A列を削除して左にシフト`);
        }
        
        // 調整後の列数を取得
        const adjustedColCount = body[0] ? body[0].length : 0;
        maxCols = Math.max(maxCols, adjustedColCount);
        
        // ✅ 台帳自動連携_1の場合、案件番号とE,F,G列の値を保存（A列削除後のインデックス）
        if (sheetName === '台帳自動連携_1') {
          logUsage(`台帳自動連携_1 データ収集開始`);
          for (let i = 0; i < body.length; i++) {
            // A列削除後、案件番号は元のB列→新しいA列（インデックス0）
            const 案件番号 = String(body[i][0] || '').trim();
            if (案件番号) {
              const データ = {
                売上分類: body[i][4] || '', // 元のF列→新しいE列（インデックス4）
                取引区分: body[i][5] || '', // 元のG列→新しいF列（インデックス5）  
                商流: body[i][6] || ''      // 元のH列→新しいG列（インデックス6）
              };
              台帳1データMap.set(案件番号, データ);
            }
          }
          logUsage(`台帳自動連携_1 から ${台帳1データMap.size} 件のデータを収集`);
        }
  
        // ✅ 台帳自動連携_2の場合、案件番号が一致したらE,F,G列の値を上書き（A列削除後のインデックス）
        if (sheetName === '台帳自動連携_2') {
          logUsage(`台帳自動連携_2 データ照合開始`);
          let matchCount = 0;
          for (let i = 0; i < body.length; i++) {
            // A列削除後、案件番号は元のB列→新しいA列（インデックス0）
            const 案件番号 = String(body[i][0] || '').trim();
            if (案件番号 && 台帳1データMap.has(案件番号)) {
              const matchedData = 台帳1データMap.get(案件番号);
              body[i][4] = matchedData.売上分類;  // 新しいE列
              body[i][5] = matchedData.取引区分;  // 新しいF列
              body[i][6] = matchedData.商流;      // 新しいG列
              matchCount++;
            }
          }
          logUsage(`台帳自動連携_2 で ${matchCount} 件のデータを引き継ぎ`);
        }
  
        // ✅ 台帳自動連携_1と台帳自動連携_2の場合、H列に「通常」を設定（A列削除後のインデックス）
        if (sheetName === '台帳自動連携_1' || sheetName === '台帳自動連携_2') {
          for (let i = 0; i < body.length; i++) {
            // 新しいH列（インデックス7）に「通常」を設定
            if (body[i].length < 8) {
              // 配列を拡張
              while (body[i].length < 8) {
                body[i].push('');
              }
            }
            body[i][7] = '通常';
          }
        }
  
        // データと背景色を配列に追加
        for (let i = 0; i < body.length; i++) {
          const rowData = body[i].slice(); // 配列をコピー
          
          // L列（インデックス11）の名前から空白を削除
          if (rowData.length > 11 && rowData[11]) {
            rowData[11] = removeSpacesFromName(rowData[11]);
          }
          
          // M列（インデックス12）の計上日を yyyy/MM/dd 形式に変換
          if (rowData.length > 12 && rowData[12]) {
            rowData[12] = formatDateToString(rowData[12]);
          }
          
          // R列（インデックス17）の売上計上日も yyyy/MM/dd 形式に変換
          if (rowData.length > 17 && rowData[17]) {
            rowData[17] = formatDateToString(rowData[17]);
          }
          
          allData.push(rowData);
          
          // 背景色の配列を作成
          const rowBackground = new Array(rowData.length).fill(colors[s % colors.length]);
          allBackgrounds.push(rowBackground);
        }
      }
      
      // メモリ解放のため定期的にflush
      if (s > 0 && s % 2 === 0) {
        SpreadsheetApp.flush();
      }
    }
  
    // ヘッダーが存在していれば復元、なければデフォルトヘッダーを設定
    if (headerRow && headerRow.length > 0) {
      // 既存のヘッダーを復元
      転記シート.getRange(1, 1, 1, headerRow.length).setValues([headerRow]);
      if (headerBackgrounds) {
        転記シート.getRange(1, 1, 1, headerBackgrounds.length).setBackgrounds([headerBackgrounds]);
      }
    } else if (allData.length > 0) {
      // ヘッダーがない場合は、ソアスク自動連携のヘッダーをコピー
      const sourceSheet = ss.getSheetByName('ソアスク自動連携');
      if (sourceSheet && sourceSheet.getLastRow() >= 1) {
        const headers = sourceSheet.getRange(1, 1, 1, sourceSheet.getLastColumn()).getValues();
        if (headers && headers[0]) {
          転記シート.getRange(1, 1, 1, headers[0].length).setValues(headers);
          logUsage('ソアスク自動連携からヘッダーを設定');
        }
      }
    }
  
    // 全データを一括で書き込み（2行目から）
    if (allData.length > 0) {
      // 最大列数に合わせて配列を調整
      maxCols = Math.max(...allData.map(row => row.length));
      
      for (let i = 0; i < allData.length; i++) {
        while (allData[i].length < maxCols) {
          allData[i].push('');
          allBackgrounds[i].push(allBackgrounds[i][allBackgrounds[i].length - 1]);
        }
      }
      
      // データを一括で設定（2行目から開始）
      転記シート.getRange(2, 1, allData.length, maxCols).setValues(allData);
      
      // 背景色を一括で設定
      転記シート.getRange(2, 1, allData.length, maxCols).setBackgrounds(allBackgrounds);
      
      // N列（売上金額）を日本円の通貨形式に設定
      if (allData.length > 0) {
        転記シート
          .getRange(2, 14, allData.length, 1) // N列（2行目から）
          .setNumberFormat('[$¥-ja-JP]#,##0;[Red]-[$¥-ja-JP]#,##0');
      }
    }
  
    logUsage(`転記シート（全データ） に ${allData.length} 行を統合出力（2行目から）`);
    logUsage(`台帳自動連携シートはA列を削除して左にシフト済み`);
  }
  
  function columnLetterToIndex(letter) {
    let column = 0;
    for (let i = 0; i < letter.length; i++) {
      column *= 26;
      column += letter.charCodeAt(i) - 64;
    }
    return column - 1;
  }
  
  function logUsage(message) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const logSheet = ss.getSheetByName('台帳連携ログ');
    logSheet.appendRow([new Date(), message]);
  }
  
  // 処理を分割実行する関数（タイムアウト対策）
  function processStep1() {
    importExcelSheetToTarget(0, '台帳自動連携_1', 7, 8, 7);
    logUsage('ステップ1完了：台帳自動連携_1の処理');
  }
  
  function processStep2() {
    importExcelSheetToTarget(1, '台帳自動連携_2', 10, 11, 10);
    logUsage('ステップ2完了：台帳自動連携_2の処理');
  }
  
  function processStep3() {
    mergeSheetsTo転記シート();
    logUsage('ステップ3完了：転記シートの作成');
  }
  
  function processStep4() {
    filterFromTransferSheet();
    logUsage('ステップ4完了：期間フィルタリング');
  }
  
  // タイムアウト対策版の実行関数
  function processAllWithTimeout() {
    try {
      processStep1();
      Utilities.sleep(1000); // 1秒待機
      
      processStep2();
      Utilities.sleep(1000); // 1秒待機
      
      processStep3();
      Utilities.sleep(1000); // 1秒待機
      
      processStep4();
      
      logUsage('=== 全処理完了 ===');
    } catch (e) {
      logUsage(`エラー発生: ${e.message}`);
      throw e;
    }
  }
  
  // デバッグ用：日付をテストする関数
  function testDateParsing() {
    const testDates = [
      '2024/12/01',
      '2025/02/01',
      '2024年12月1日',
      '2025-03-15',
      '2024/11/30',  // 期間外
      '2025/12/01',  // 期間外
      45623,  // Excelシリアル値の例
      null,
      ''
    ];
    
    const startDate = new Date('2024/12/01');
    const endDate = new Date('2025/11/30');
    
    testDates.forEach((dateValue, index) => {
      const parsed = parseDate(dateValue);
      const inRange = parsed && parsed >= startDate && parsed <= endDate;
      logUsage(`テスト[${index}] 値: "${dateValue}" → パース: ${parsed ? parsed.toLocaleDateString() : 'null'} → 期間内: ${inRange}`);
    });
  }
  
  // ヘッダーを設定する関数（必要に応じて使用）
  function setupHeaders() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // ソアスク+台帳データのヘッダー設定
    let targetSheet = ss.getSheetByName('ソアスク+台帳データ');
    if (!targetSheet) {
      targetSheet = ss.insertSheet('ソアスク+台帳データ');
    }
    
    // 転記シート（全データ）のヘッダー設定
    let transferSheet = ss.getSheetByName('転記シート（全データ）');
    if (!transferSheet) {
      transferSheet = ss.insertSheet('転記シート（全データ）');
    }
    
    // ソアスク自動連携シートからヘッダーをコピー
    const sourceSheet = ss.getSheetByName('ソアスク自動連携');
    if (sourceSheet && sourceSheet.getLastRow() >= 1) {
      const headers = sourceSheet.getRange(1, 1, 1, sourceSheet.getLastColumn()).getValues();
      if (headers && headers[0]) {
        // ソアスク+台帳データにヘッダーを設定
        targetSheet.getRange(1, 1, 1, headers[0].length).setValues(headers);
        
        // 転記シート（全データ）にもA列から始まるようにヘッダーを設定
        transferSheet.getRange(1, 1, 1, headers[0].length).setValues(headers);
        
        logUsage('両シートにヘッダーを設定しました');
      }
    } else {
      // デフォルトのヘッダーを設定
      const defaultHeaders = [
        'エンドユーザ計上部門','案件担当','売上計上日','売上（税抜）','案件名','品目',
        '売上品目分類','契約開始日','契約終了日','計上日','直販','商談（税抜）',
        '原価（税抜）','品目','発注先','仕入勘定科目','納品日'
      ];
      
      // ソアスク+台帳データにヘッダーを設定
      targetSheet.getRange(1, 1, 1, defaultHeaders.length).setValues([defaultHeaders]);
      
      // 転記シートにもA列から始まるようにヘッダーを設定
      transferSheet.getRange(1, 1, 1, defaultHeaders.length).setValues([defaultHeaders]);
      
      logUsage('デフォルトヘッダーを設定しました');
    }
  }
  
  // 名前から空白を削除する補助関数
  function removeSpacesFromName(name) {
    if (!name) return '';
    // 半角スペースと全角スペースの両方を削除
    return String(name).replace(/[\s　]+/g, '');
  }
  
  // 日付をyyyy/MM/dd形式に変換する補助関数
  function formatDateToString(dateValue) {
    if (!dateValue) return '';
    
    let date = null;
    
    if (dateValue instanceof Date) {
      date = dateValue;
    } else if (typeof dateValue === 'string' && dateValue.includes('GMT')) {
      date = new Date(dateValue);
    } else {
      date = parseDate(dateValue);
    }
    
    if (date && !isNaN(date.getTime())) {
      const year = date.getFullYear();
      const month = String(date.getMonth() + 1).padStart(2, '0');
      const day = String(date.getDate()).padStart(2, '0');
      return `${year}/${month}/${day}`;
    }
    
    return dateValue; // 変換できない場合は元の値を返す
  }
  
  // デバッグ用：転記シートのN列データを確認
  function checkTransferSheetDates() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sourceSheet = ss.getSheetByName('転記シート（全データ）');
    
    if (!sourceSheet) {
      logUsage('エラー: 転記シート（全データ）が見つかりません');
      return;
    }
    
    const lastRow = sourceSheet.getLastRow();
    const lastCol = sourceSheet.getLastColumn();
    
    logUsage(`=== 転記シートのN列データ確認 ===`);
    logUsage(`シートサイズ: ${lastRow}行 x ${lastCol}列`);
    
    // ヘッダー行を確認
    if (lastRow >= 1 && lastCol >= 14) {
      const headers = sourceSheet.getRange(1, 1, 1, Math.min(lastCol, 20)).getValues()[0];
      logUsage(`ヘッダー（最初の20列）:`);
      headers.forEach((header, index) => {
        if (header) {
          logUsage(`  ${columnIndexToLetter(index)}列: "${header}"`);
        }
      });
    }
    
    // M列（13列目）のデータをサンプル取得 - A列から始まるので13列目
    if (lastRow >= 2 && lastCol >= 13) {
      const mColumnData = sourceSheet.getRange(2, 13, Math.min(lastRow - 1, 20), 1).getValues();
      logUsage(`\nM列のサンプルデータ（最初の20行）:`);
      
      const startDate = new Date('2024/12/01');
      const endDate = new Date('2025/11/30');
      
      mColumnData.forEach((row, index) => {
        const value = row[0];
        const parsed = parseDate(value);
        const inRange = parsed && parsed >= startDate && parsed <= endDate;
        
        logUsage(`  行${index + 2}: 値="${value}" → ${parsed ? parsed.toLocaleDateString() : 'パース失敗'} → ${inRange ? '期間内' : '期間外'}`);
      });
    }
  }
  
  // L列の名前データを確認する関数
  function checkNameData() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sourceSheet = ss.getSheetByName('転記シート（全データ）');
    
    if (!sourceSheet) {
      logUsage('エラー: 転記シート（全データ）が見つかりません');
      return;
    }
    
    const lastRow = sourceSheet.getLastRow();
    const lastCol = sourceSheet.getLastColumn();
    
    logUsage(`=== L列の名前データ確認 ===`);
    
    if (lastRow >= 2 && lastCol >= 12) {
      const lColumnData = sourceSheet.getRange(2, 12, Math.min(lastRow - 1, 20), 1).getValues();
      logUsage(`L列のサンプルデータ（最初の20行）:`);
      
      let spacesFoundCount = 0;
      lColumnData.forEach((row, index) => {
        const value = row[0];
        if (value) {
          const hasSpaces = String(value).includes(' ') || String(value).includes('　');
          if (hasSpaces) {
            spacesFoundCount++;
            logUsage(`  行${index + 2}: "${value}" → 空白あり`);
          }
        }
      });
      
      if (spacesFoundCount === 0) {
        logUsage(`空白を含む名前は見つかりませんでした`);
      } else {
        logUsage(`${spacesFoundCount}件の名前に空白が含まれています`);
      }
    }
  }
  
  // N列が正しいか確認する関数
  function verifyDateColumn() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sourceSheet = ss.getSheetByName('転記シート（全データ）');
    
    if (!sourceSheet) {
      logUsage('エラー: 転記シート（全データ）が見つかりません');
      return;
    }
    
    logUsage(`=== 計上日列の検証 ===`);
    
    // ヘッダー行全体を取得
    const headers = sourceSheet.getRange(1, 1, 1, sourceSheet.getLastColumn()).getValues()[0];
    
    // 「計上日」を含む列をすべて探す
    const dateColumns = [];
    headers.forEach((header, index) => {
      if (header && header.toString().includes('計上日')) {
        dateColumns.push({
          index: index,
          column: columnIndexToLetter(index),
          header: header
        });
      }
    });
    
    if (dateColumns.length === 0) {
      logUsage('警告: 「計上日」を含む列が見つかりません');
    } else {
      logUsage(`「計上日」を含む列が${dateColumns.length}個見つかりました:`);
      dateColumns.forEach(col => {
        logUsage(`  ${col.column}列（インデックス${col.index}）: "${col.header}"`);
        
        // 各列のサンプルデータを取得
        if (sourceSheet.getLastRow() >= 2) {
          const sampleData = sourceSheet.getRange(2, col.index + 1, Math.min(5, sourceSheet.getLastRow() - 1), 1).getValues();
          logUsage(`    サンプルデータ:`);
          sampleData.forEach((row, i) => {
            logUsage(`      行${i + 2}: "${row[0]}"`);
          });
        }
      });
    }
  }