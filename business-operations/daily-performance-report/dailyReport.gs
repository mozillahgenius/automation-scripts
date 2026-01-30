/**
 * 前日比レポート自動送信システム (GAS版)
 * 仕様書: 業績Dailyレポート
 * 
 * 主要機能:
 * - 複数列対応の前日比較
 * - HTML形式メール送信
 * - CSV添付機能
 * - 差分0時の送信抑止
 * - 自動スケジューリング
 * - エラーハンドリング・通知
 */

// =============================================================================
// メイン実行関数
// =============================================================================

/**
 * 日次レポート実行（定時実行される関数）
 */
function runDailyReport() {
  const startTime = Date.now();
  let logMessage = '';
  let diffCount = 0;
  let sentCount = 0;
  
  try {
    // 1. 設定読み込み
    const settings = getSettings();
    
    // 2. データ取得
    const todayData = getSheetMap(settings['シート_今日'] || '今日');
    const prevData = getSheetMap(settings['シート_前日'] || '前日');
    
    // 3. 差分計算
    const diffResult = calcDiffs(todayData, prevData, settings);
    diffCount = diffResult.items.length;
    
    // 4. 差分0時の送信抑止チェック
    const sendOnNoDiff = settings['差分0でも送信'] !== 'false';
    if (diffCount === 0 && !sendOnNoDiff) {
      logRun('OK(NO_DIFF)', 0, 0, Date.now() - startTime, '差分0件のため送信なし');
      return;
    }
    
    // 5. HTMLメール生成
    const htmlContent = buildHtml(diffResult, settings);
    
    // 6. CSV添付ファイル生成（オプション）
    let attachment = null;
    if (settings['CSV添付'] === 'true') {
      attachment = createCsvAttachment(diffResult, settings);
    }
    
    // 7. メール送信
    const subject = `${settings['件名プレフィックス'] || '[日次レポート]'} ${Utilities.formatDate(new Date(), settings['タイムゾーン'] || 'Asia/Tokyo', 'yyyy-MM-dd')} 前日比レポート（差分${diffCount}件）`;
    
    const to = settings['送信先'];
    if (!to) {
      throw new Error('送信先が設定されていません');
    }
    
    const mailOptions = {
      to: to,
      subject: subject,
      htmlBody: htmlContent
    };
    
    if (settings['Cc']) mailOptions.cc = settings['Cc'];
    if (settings['Bcc']) mailOptions.bcc = settings['Bcc'];
    if (attachment) mailOptions.attachments = [attachment];
    
    MailApp.sendEmail(mailOptions);
    sentCount = to.split(',').length;
    
    // 8. スナップショット更新（成功時のみ）
    updateSnapshot(settings);
    
    // 9. ログ記録
    logMessage = diffCount > 0 ? `差分${diffCount}件を${sentCount}件の宛先に送信` : '差分0件で送信実行';
    logRun('OK', sentCount, diffCount, Date.now() - startTime, logMessage);
    
  } catch (error) {
    // エラーハンドリング
    const errorMsg = `エラー: ${error.message}`;
    logRun('ERROR', 0, 0, Date.now() - startTime, errorMsg);
    
    // エラー通知メール
    sendErrorNotification(error, settings || {});
    
    throw error; // 再throw for trigger error handling
  }
}

// =============================================================================
// 設定管理
// =============================================================================

/**
 * 設定シートから設定値を読み込み
 * @return {Object} 設定オブジェクト
 */
function getSettings() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('設定');
  if (!sheet) {
    throw new Error('設定シートが見つかりません');
  }
  
  const data = sheet.getDataRange().getValues();
  const settings = {};
  
  // デフォルト値
  const defaults = {
    '送信時刻(24h)': '09:10',
    'タイムゾーン': 'Asia/Tokyo',
    'シート_今日': '今日',
    'シート_前日': '前日',
    '増減率表示小数': '1',
    '金額表示桁区切り': 'true',
    '上位件数': '',
    '閾値_最小変化量': '0',
    '件名プレフィックス': '[日次レポート]',
    'メール署名HTML': '',
    'HTML_軽量表示': 'false',
    '差分判定対象列': '値',
    '差分0でも送信': 'true',
    'エラー通知先': '',
    'メール出力列': '',
    'CSV添付': 'false'
  };
  
  // 設定値読み込み
  for (let i = 0; i < data.length; i++) {
    const key = data[i][0];
    const value = data[i][1];
    if (key && value !== undefined && value !== '') {
      settings[key] = value;
    }
  }
  
  // デフォルト値適用
  for (const [key, defaultValue] of Object.entries(defaults)) {
    if (!(key in settings)) {
      settings[key] = defaultValue;
    }
  }
  
  // 必須項目チェック
  if (!settings['送信先']) {
    throw new Error('送信先が設定されていません');
  }
  
  return settings;
}

// =============================================================================
// データ取得・処理
// =============================================================================

/**
 * シートデータをMapオブジェクトに変換
 * @param {string} sheetName シート名
 * @return {Object} {headers: [...], dataMap: Map<ID, rowData>}
 */
function getSheetMap(sheetName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) {
    throw new Error(`シート「${sheetName}」が見つかりません`);
  }
  
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) {
    return { headers: [], dataMap: new Map() };
  }
  
  const headers = data[0];
  const dataMap = new Map();
  let skipCount = 0;
  
  // IDカラムのインデックスを特定
  const idIndex = headers.findIndex(h => h === 'ID');
  if (idIndex === -1) {
    throw new Error(`シート「${sheetName}」にIDカラムが見つかりません`);
  }
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const id = String(row[idIndex] || '');
    
    if (!id) {
      skipCount++;
      continue;
    }
    
    if (dataMap.has(id)) {
      skipCount++;
      continue; // ID重複スキップ
    }
    
    const rowData = {};
    headers.forEach((header, index) => {
      rowData[header] = row[index];
    });
    
    dataMap.set(id, rowData);
  }
  
  if (skipCount > 0) {
    console.log(`${sheetName}: ID欠落・重複により${skipCount}件をスキップ`);
  }
  
  return { headers, dataMap };
}

/**
 * 数値変換（非数値の場合はログ出力）
 * @param {*} value 値
 * @return {number|null} 数値またはnull
 */
function toNum(value) {
  if (value === null || value === undefined || value === '') {
    return null;
  }
  
  const num = Number(value);
  return isNaN(num) ? null : num;
}

// =============================================================================
// 差分計算
// =============================================================================

/**
 * 前日比差分計算（複数列対応）
 * @param {Object} todayData 今日のデータ
 * @param {Object} prevData 前日のデータ
 * @param {Object} settings 設定
 * @return {Object} 差分結果
 */
function calcDiffs(todayData, prevData, settings) {
  const targetColumns = settings['差分判定対象列'].split(',').map(col => col.trim());
  const threshold = Number(settings['閾値_最小変化量'] || 0);
  const topCount = settings['上位件数'] ? Number(settings['上位件数']) : null;
  
  const items = [];
  const allIds = new Set([...todayData.dataMap.keys(), ...prevData.dataMap.keys()]);
  
  for (const id of allIds) {
    const todayRow = todayData.dataMap.get(id);
    const prevRow = prevData.dataMap.get(id);
    
    let status;
    if (!prevRow) {
      status = '新規';
    } else if (!todayRow) {
      status = '削除';
    } else {
      status = '更新';
    }
    
    const details = {};
    let hasSignificantChange = false;
    let firstColumnAbsDiff = 0; // ソート用
    
    for (let i = 0; i < targetColumns.length; i++) {
      const column = targetColumns[i];
      const todayValue = todayRow ? toNum(todayRow[column]) : null;
      const prevValue = prevRow ? toNum(prevRow[column]) : null;
      
      let diff = null;
      let rate = null;
      
      if (status === '新規') {
        diff = todayValue;
        rate = 'N/A';
      } else if (status === '削除') {
        diff = prevValue ? -prevValue : null;
        rate = 'N/A';
      } else {
        // 更新
        if (todayValue !== null && prevValue !== null) {
          diff = todayValue - prevValue;
          rate = prevValue !== 0 ? ((todayValue - prevValue) / prevValue) * 100 : 'N/A';
        } else if (todayValue !== null) {
          diff = todayValue;
          rate = 'N/A';
        } else if (prevValue !== null) {
          diff = -prevValue;
          rate = 'N/A';
        }
      }
      
      details[column] = {
        前日: prevValue,
        今日: todayValue,
        差: diff,
        率: rate
      };
      
      // 閾値チェック（第一列基準）
      if (i === 0) {
        firstColumnAbsDiff = diff !== null ? Math.abs(diff) : 0;
        if (firstColumnAbsDiff >= threshold) {
          hasSignificantChange = true;
        }
      }
    }
    
    // 閾値フィルタ
    if (!hasSignificantChange && status === '更新') {
      continue;
    }
    
    const item = {
      ID: id,
      名称: todayRow ? todayRow['名称'] : (prevRow ? prevRow['名称'] : ''),
      状態: status,
      明細: details,
      _sortKey: firstColumnAbsDiff // ソート用内部フィールド
    };
    
    items.push(item);
  }
  
  // ソート（第一列の絶対差降順）
  items.sort((a, b) => b._sortKey - a._sortKey);
  
  // 上位N件抽出
  if (topCount && items.length > topCount) {
    items.splice(topCount);
  }
  
  return { items };
}

// =============================================================================
// HTMLメール生成
// =============================================================================

/**
 * HTML形式のメール本文生成
 * @param {Object} diffResult 差分結果
 * @param {Object} settings 設定
 * @return {string} HTML文字列
 */
function buildHtml(diffResult, settings) {
  const isLightMode = settings['HTML_軽量表示'] === 'true';
  const decimalPlaces = Number(settings['増減率表示小数'] || 1);
  const useThousandsSeparator = settings['金額表示桁区切り'] !== 'false';
  const targetColumns = settings['差分判定対象列'].split(',').map(col => col.trim());
  const timezone = settings['タイムゾーン'] || 'Asia/Tokyo';
  
  const now = new Date();
  const dateStr = Utilities.formatDate(now, timezone, 'yyyy-MM-dd');
  const timeStr = Utilities.formatDate(now, timezone, 'HH:mm:ss');
  
  let html = `
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <title>前日比レポート</title>
  <style>
    body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; margin: 20px; background-color: #f5f5f5; }
    .container { max-width: 1200px; margin: 0 auto; background-color: white; padding: 20px; border-radius: 8px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }
    .header { background-color: #2c5aa0; color: white; padding: 15px; border-radius: 5px; margin-bottom: 20px; }
    .summary { background-color: #e8f4fd; padding: 15px; border-radius: 5px; margin-bottom: 20px; border-left: 4px solid #2c5aa0; }
    .table-container { margin-bottom: 30px; }
    .column-title { font-size: 18px; font-weight: bold; color: #2c5aa0; margin-bottom: 10px; padding-bottom: 5px; border-bottom: 2px solid #2c5aa0; }
    table { width: 100%; border-collapse: collapse; font-size: 14px; }
    th { background-color: #f8f9fa; padding: 12px 8px; text-align: left; border: 1px solid #dee2e6; font-weight: 600; }
    td { padding: 10px 8px; border: 1px solid #dee2e6; }
    .increase { background-color: #d4edda; }
    .decrease { background-color: #f8d7da; }
    .new { background-color: #cce5ff; }
    .delete { background-color: #e2e3e5; }
    .number { text-align: right; font-family: 'Courier New', monospace; }
    .signature { margin-top: 20px; padding-top: 20px; border-top: 1px solid #ddd; color: #666; }
  </style>
</head>
<body>
  <div class="container">
    <div class="header">
      <h1>📊 前日比レポート</h1>
      <p>生成日時: ${dateStr} ${timeStr} (${timezone})<br>
         対象日: ${dateStr}<br>
         差分件数: ${diffResult.items.length}件</p>
    </div>
`;

  // 概要カード（増加TOP3・減少TOP3）
  if (diffResult.items.length > 0 && targetColumns.length > 0) {
    const firstColumn = targetColumns[0];
    const increases = diffResult.items
      .filter(item => item.明細[firstColumn] && item.明細[firstColumn].差 > 0)
      .slice(0, 3);
    const decreases = diffResult.items
      .filter(item => item.明細[firstColumn] && item.明細[firstColumn].差 < 0)
      .slice(0, 3);
    
    html += `
    <div class="summary">
      <h3>📈 概要サマリー (${firstColumn}基準)</h3>
      <div style="display: flex; gap: 20px;">
        <div style="flex: 1;">
          <h4 style="color: #28a745;">🔼 増加 TOP3</h4>
          <ul>
`;
    increases.forEach(item => {
      const diff = formatNumber(item.明細[firstColumn].差, useThousandsSeparator);
      html += `<li>${escapeHtml(item.名称 || item.ID)}: +${diff}</li>`;
    });
    
    html += `
          </ul>
        </div>
        <div style="flex: 1;">
          <h4 style="color: #dc3545;">🔽 減少 TOP3</h4>
          <ul>
`;
    decreases.forEach(item => {
      const diff = formatNumber(Math.abs(item.明細[firstColumn].差), useThousandsSeparator);
      html += `<li>${escapeHtml(item.名称 || item.ID)}: -${diff}</li>`;
    });
    
    html += `
          </ul>
        </div>
      </div>
    </div>
`;
  }

  // 列ごとのテーブル生成
  for (const column of targetColumns) {
    html += `
    <div class="table-container">
      <div class="column-title">[${escapeHtml(column)}] 前日比</div>
      <table>
        <thead>
          <tr>
            <th>ID</th>
            <th>名称</th>
            <th>今日</th>
            <th>前日</th>
            <th>差分</th>
            <th>前日比(%)</th>
            <th>ステータス</th>
          </tr>
        </thead>
        <tbody>
`;
    
    diffResult.items.forEach(item => {
      const detail = item.明細[column];
      if (!detail) return;
      
      const rowClass = getRowClass(item.状態, detail.差);
      const todayValue = detail.今日 !== null ? formatNumber(detail.今日, useThousandsSeparator) : '-';
      const prevValue = detail.前日 !== null ? formatNumber(detail.前日, useThousandsSeparator) : '-';
      const diffValue = detail.差 !== null ? formatNumber(detail.差, useThousandsSeparator, true) : '-';
      const rateValue = detail.率 !== 'N/A' && detail.率 !== null ? 
        `${detail.率.toFixed(decimalPlaces)}%` : 'N/A';
      
      html += `
        <tr class="${rowClass}">
          <td>${escapeHtml(item.ID)}</td>
          <td>${escapeHtml(item.名称 || '')}</td>
          <td class="number">${todayValue}</td>
          <td class="number">${prevValue}</td>
          <td class="number">${diffValue}</td>
          <td class="number">${rateValue}</td>
          <td>${item.状態}</td>
        </tr>
`;
    });
    
    html += `
        </tbody>
      </table>
    </div>
`;
  }
  
  // 署名
  const signature = settings['メール署名HTML'];
  if (signature) {
    html += `<div class="signature">${signature}</div>`;
  }
  
  html += `
  </div>
</body>
</html>
`;
  
  return html;
}

/**
 * 行のCSSクラス決定
 * @param {string} status ステータス
 * @param {number} diff 差分値
 * @return {string} CSSクラス名
 */
function getRowClass(status, diff) {
  if (status === '新規') return 'new';
  if (status === '削除') return 'delete';
  if (diff > 0) return 'increase';
  if (diff < 0) return 'decrease';
  return '';
}

/**
 * 数値フォーマット
 * @param {number} value 数値
 * @param {boolean} useThousandsSeparator 桁区切り使用
 * @param {boolean} showSign 符号表示
 * @return {string} フォーマット済み文字列
 */
function formatNumber(value, useThousandsSeparator = true, showSign = false) {
  if (value === null || value === undefined) return '-';
  
  let formatted = Math.round(value).toString();
  
  if (useThousandsSeparator) {
    formatted = Number(value).toLocaleString();
  }
  
  if (showSign && value > 0) {
    formatted = '+' + formatted;
  }
  
  return formatted;
}

/**
 * HTMLエスケープ
 * @param {string} text テキスト
 * @return {string} エスケープ済みテキスト
 */
function escapeHtml(text) {
  if (!text) return '';
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

// =============================================================================
// CSV添付生成
// =============================================================================

/**
 * CSV添付ファイル生成
 * @param {Object} diffResult 差分結果
 * @param {Object} settings 設定
 * @return {Object} 添付ファイルオブジェクト
 */
function createCsvAttachment(diffResult, settings) {
  const targetColumns = settings['差分判定対象列'].split(',').map(col => col.trim());
  
  // CSVヘッダー
  let csvContent = 'ID,名称,ステータス';
  for (const column of targetColumns) {
    csvContent += `,${column}_今日,${column}_前日,${column}_差分,${column}_前日比%`;
  }
  csvContent += '\n';
  
  // データ行
  diffResult.items.forEach(item => {
    let row = `"${item.ID}","${item.名称 || ''}","${item.状態}"`;
    
    for (const column of targetColumns) {
      const detail = item.明細[column];
      if (detail) {
        const todayVal = detail.今日 !== null ? detail.今日 : '';
        const prevVal = detail.前日 !== null ? detail.前日 : '';
        const diffVal = detail.差 !== null ? detail.差 : '';
        const rateVal = detail.率 !== 'N/A' && detail.率 !== null ? detail.率.toFixed(1) : '';
        
        row += `,"${todayVal}","${prevVal}","${diffVal}","${rateVal}"`;
      } else {
        row += ',"","","",""';
      }
    }
    
    csvContent += row + '\n';
  });
  
  // UTF-8 BOM付きで作成
  const blob = Utilities.newBlob('\ufeff' + csvContent, 'text/csv', '前日比レポート.csv');
  
  return blob;
}

// =============================================================================
// スナップショット更新
// =============================================================================

/**
 * 前日シートを今日シートで更新
 * @param {Object} settings 設定
 */
function updateSnapshot(settings) {
  const todaySheetName = settings['シート_今日'] || '今日';
  const prevSheetName = settings['シート_前日'] || '前日';
  
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const todaySheet = spreadsheet.getSheetByName(todaySheetName);
  const prevSheet = spreadsheet.getSheetByName(prevSheetName);
  
  if (!todaySheet) {
    throw new Error(`今日シート「${todaySheetName}」が見つかりません`);
  }
  
  if (!prevSheet) {
    throw new Error(`前日シート「${prevSheetName}」が見つかりません`);
  }
  
  // 前日シートをクリア
  prevSheet.clear();
  
  // 今日シートの全データをコピー
  const todayData = todaySheet.getDataRange().getValues();
  if (todayData.length > 0) {
    prevSheet.getRange(1, 1, todayData.length, todayData[0].length).setValues(todayData);
  }
}

// =============================================================================
// ログ記録
// =============================================================================

/**
 * 実行ログを記録
 * @param {string} result 結果 (OK|OK(NO_DIFF)|ERROR)
 * @param {number} sentCount 送信件数
 * @param {number} diffCount 差分件数
 * @param {number} duration 所要時間(ms)
 * @param {string} message メッセージ
 */
function logRun(result, sentCount, diffCount, duration, message) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ログ');
    if (!sheet) {
      console.log('ログシートが見つかりません');
      return;
    }
    
    const now = new Date();
    const row = [now, result, sentCount, diffCount, duration, message];
    sheet.appendRow(row);
  } catch (error) {
    console.log('ログ記録エラー:', error.message);
  }
}

// =============================================================================
// エラーハンドリング
// =============================================================================

/**
 * エラー通知メール送信
 * @param {Error} error エラーオブジェクト
 * @param {Object} settings 設定
 */
function sendErrorNotification(error, settings) {
  try {
    const errorTo = settings['エラー通知先'] || settings['送信先']?.split(',')[0];
    if (!errorTo) {
      console.log('エラー通知先が設定されていません');
      return;
    }
    
    const subject = '[日次レポート] ERROR';
    const message = error.message || 'Unknown error';
    const stack = error.stack ? error.stack.substring(0, 1000) : 'No stack trace';
    
    const body = `
日次レポートの実行中にエラーが発生しました。

エラーメッセージ:
${message}

スタックトレース:
${stack}

発生時刻: ${new Date().toLocaleString('ja-JP', {timeZone: settings['タイムゾーン'] || 'Asia/Tokyo'})}

設定を確認し、必要に応じて runDailyReport() を手動実行してください。
`;
    
    MailApp.sendEmail({
      to: errorTo,
      subject: subject,
      body: body
    });
    
  } catch (notificationError) {
    console.log('エラー通知の送信に失敗:', notificationError.message);
  }
}

// =============================================================================
// トリガー管理
// =============================================================================

/**
 * 定時実行トリガーをインストール
 */
function installTrigger() {
  try {
    // 既存トリガーを削除
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(trigger => {
      if (trigger.getHandlerFunction() === 'runDailyReport') {
        ScriptApp.deleteTrigger(trigger);
      }
    });
    
    // 設定読み込み
    const settings = getSettings();
    const timeStr = settings['送信時刻(24h)'] || '09:10';
    const timezone = settings['タイムゾーン'] || 'Asia/Tokyo';
    
    // 時刻解析
    const [hour, minute] = timeStr.split(':').map(Number);
    if (isNaN(hour) || isNaN(minute) || hour < 0 || hour > 23 || minute < 0 || minute > 59) {
      throw new Error('送信時刻の形式が正しくありません (HH:MM)');
    }
    
    // 新しいトリガーを作成
    ScriptApp.newTrigger('runDailyReport')
      .timeBased()
      .everyDays(1)
      .atHour(hour)
      .nearMinute(minute)
      .inTimezone(timezone)
      .create();
    
    console.log(`トリガーを設定しました: 毎日 ${timeStr} (${timezone})`);
    return `トリガーを設定しました: 毎日 ${timeStr} (${timezone})`;
    
  } catch (error) {
    console.log('トリガー設定エラー:', error.message);
    throw error;
  }
}

/**
 * トリガーを削除
 */
function uninstallTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  let deletedCount = 0;
  
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'runDailyReport') {
      ScriptApp.deleteTrigger(trigger);
      deletedCount++;
    }
  });
  
  console.log(`${deletedCount}個のトリガーを削除しました`);
  return `${deletedCount}個のトリガーを削除しました`;
}

// =============================================================================
// ユーティリティ・テスト関数
// =============================================================================

/**
 * 設定シートのテンプレートを作成
 */
function createSettingsTemplate() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName('設定');
  
  if (!sheet) {
    sheet = spreadsheet.insertSheet('設定');
  }
  
  const template = [
    ['キー', '値'],
    ['送信先', 'your-email@example.com'],
    ['Cc', ''],
    ['Bcc', ''],
    ['送信時刻(24h)', '09:10'],
    ['タイムゾーン', 'Asia/Tokyo'],
    ['シート_今日', '今日'],
    ['シート_前日', '前日'],
    ['増減率表示小数', '1'],
    ['金額表示桁区切り', 'true'],
    ['上位件数', '20'],
    ['閾値_最小変化量', '0'],
    ['件名プレフィックス', '[日次レポート]'],
    ['メール署名HTML', '<p>自動送信メールです。返信しないでください。</p>'],
    ['HTML_軽量表示', 'false'],
    ['差分判定対象列', '値'],
    ['差分0でも送信', 'true'],
    ['エラー通知先', 'ops@example.com'],
    ['メール出力列', 'ID,名称,値'],
    ['CSV添付', 'false']
  ];
  
  sheet.clear();
  sheet.getRange(1, 1, template.length, 2).setValues(template);
  
  // ヘッダー行をフォーマット
  sheet.getRange(1, 1, 1, 2).setFontWeight('bold').setBackground('#e8f4fd');
  
  return '設定シートのテンプレートを作成しました';
}

/**
 * ログシートを作成
 */
function createLogSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName('ログ');
  
  if (!sheet) {
    sheet = spreadsheet.insertSheet('ログ');
  }
  
  const headers = ['実行日時', '結果', '送信件数', '差分件数', '所要ms', 'メッセージ'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#e8f4fd');
  
  return 'ログシートを作成しました';
}

/**
 * テスト実行（トリガーを使わずに手動実行）
 */
function testRun() {
  console.log('=== テスト実行開始 ===');
  try {
    runDailyReport();
    console.log('=== テスト実行完了 ===');
    return 'テスト実行が完了しました。ログシートとメールを確認してください。';
  } catch (error) {
    console.log('=== テスト実行エラー ===');
    console.log(error.message);
    throw error;
  }
}

/**
 * 現在の設定を表示
 */
function showCurrentSettings() {
  try {
    const settings = getSettings();
    console.log('現在の設定:');
    for (const [key, value] of Object.entries(settings)) {
      console.log(`${key}: ${value}`);
    }
    return settings;
  } catch (error) {
    console.log('設定読み込みエラー:', error.message);
    throw error;
  }
}
