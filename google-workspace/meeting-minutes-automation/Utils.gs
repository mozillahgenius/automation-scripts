/**
 * ユーティリティ関数集
 * 運用・保守・デバッグ用の補助機能
 */

/**
 * システムの動作テスト用関数
 */
function testSystem() {
  console.log('=== システムテスト開始 ===');
  
  try {
    // 1. 設定の読み込みテスト
    console.log('1. 設定読み込みテスト');
    const config = getConfig();
    console.log('設定読み込み成功:', config);
    
    // 2. Calendar API テスト
    console.log('2. Calendar API テスト');
    const events = getRecentEvents(config.calendarId, 24); // 24時間以内
    console.log(`取得イベント数: ${events.length}`);
    
    // 3. Vertex AI テスト（サンプルテキスト）
    console.log('3. Vertex AI テスト');
    const sampleTranscript = "参加者A: おはようございます。今日は定例会議ですね。参加者B: はい、今週の進捗を確認しましょう。";
    const testMinutes = generateMinutes(config.minutesPrompt, sampleTranscript);
    console.log('議事録生成成功:', testMinutes.substring(0, 100) + '...');
    
    console.log('=== システムテスト完了（正常） ===');
    
  } catch (error) {
    console.error('=== システムテストエラー ===', error);
    throw error;
  }
}

/**
 * 特定のイベントIDを手動処理する関数
 */
function processSpecificEvent(eventId) {
  try {
    console.log(`イベント ${eventId} を手動処理開始`);
    
    const config = getConfig();
    // カレンダーIDが設定されていない場合はデフォルトカレンダーを使用
    let calendar;
    if (!config.calendarId || config.calendarId === 'your-calendar-id@group.calendar.google.com') {
      calendar = CalendarApp.getDefaultCalendar();
    } else {
      calendar = CalendarApp.getCalendarById(config.calendarId);
    }
    const event = calendar.getEventById(eventId);
    
    if (!event) {
      throw new Error(`イベント ${eventId} が見つかりません`);
    }
    
    // 既に処理済みかチェック（強制実行の場合はスキップ）
    if (isProcessed(eventId)) {
      console.log('既に処理済みですが、強制実行します');
    }
    
    // トランスクリプトを取得
    const transcript = getTranscriptByEvent(event);
    if (!transcript) {
      throw new Error('トランスクリプトが見つかりません');
    }
    
    // 議事録を生成
    const minutes = generateMinutes(config.minutesPrompt, transcript);
    
    // メール宛先を収集
    const recipients = collectRecipients(event, config.additionalRecipients);
    
    // メールを送信
    const subject = `【議事録】${event.getTitle()}`;
    sendMinutesEmail(recipients, subject, minutes);
    
    // ログに記録
    logExecution(eventId, 'SUCCESS', '手動実行', event.getTitle());
    
    console.log(`イベント ${eventId} の手動処理が完了しました`);
    
  } catch (error) {
    console.error(`イベント ${eventId} の手動処理中にエラー:`, error);
    logExecution(eventId, 'ERROR', `手動実行エラー: ${error.toString()}`, '');
    throw error;
  }
}

/**
 * 実行ログのクリーンアップ（古いログを削除）
 */
function cleanupLogs(daysToKeep = 30) {
  try {
    const sheet = SpreadsheetApp.openById(CONFIG_SPREADSHEET_ID).getSheetByName(LOG_SHEET_NAME);
    const values = sheet.getDataRange().getValues();
    
    if (values.length <= 1) {
      console.log('削除対象のログがありません');
      return;
    }
    
    const cutoffDate = new Date(Date.now() - (daysToKeep * 24 * 60 * 60 * 1000));
    const rowsToDelete = [];
    
    // ヘッダー行をスキップして、古いログを特定
    for (let i = 1; i < values.length; i++) {
      const processedAt = values[i][1]; // ProcessedAt列
      if (processedAt && new Date(processedAt) < cutoffDate) {
        rowsToDelete.push(i + 1); // 1-based row number
      }
    }
    
    // 後ろの行から削除（行番号がずれないように）
    rowsToDelete.reverse();
    for (const rowNumber of rowsToDelete) {
      sheet.deleteRow(rowNumber);
    }
    
    console.log(`${rowsToDelete.length}件の古いログを削除しました（${daysToKeep}日より古い）`);
    
  } catch (error) {
    console.error('ログクリーンアップエラー:', error);
    throw error;
  }
}

/**
 * システム状態の健全性チェック
 */
function healthCheck() {
  const results = {
    timestamp: new Date(),
    checks: []
  };
  
  try {
    // 1. 設定ファイルの確認
    try {
      const config = getConfig();
      results.checks.push({
        name: '設定ファイル',
        status: 'OK',
        details: `カレンダーID: ${config.calendarId.substring(0, 10)}...`
      });
    } catch (error) {
      results.checks.push({
        name: '設定ファイル',
        status: 'ERROR',
        details: error.toString()
      });
    }
    
    // 2. スプレッドシートアクセス確認
    try {
      const sheet = getConfigSpreadsheet();
      results.checks.push({
        name: 'スプレッドシートアクセス',
        status: 'OK',
        details: `シート名: ${sheet.getName()}`
      });
    } catch (error) {
      results.checks.push({
        name: 'スプレッドシートアクセス',
        status: 'ERROR',
        details: error.toString()
      });
    }
    
    // 3. OAuth権限確認
    try {
      const token = ScriptApp.getOAuthToken();
      results.checks.push({
        name: 'OAuth権限',
        status: 'OK',
        details: 'トークン取得成功'
      });
    } catch (error) {
      results.checks.push({
        name: 'OAuth権限',
        status: 'ERROR',
        details: error.toString()
      });
    }
    
    // 4. 最近のログ確認
    try {
      const sheet = SpreadsheetApp.openById(CONFIG_SPREADSHEET_ID).getSheetByName(LOG_SHEET_NAME);
      const lastRow = sheet.getLastRow();
      const recentLogs = lastRow > 1 ? sheet.getRange(Math.max(2, lastRow - 4), 1, Math.min(5, lastRow - 1), 5).getValues() : [];
      
      results.checks.push({
        name: '最近の実行ログ',
        status: 'OK',
        details: `最近の実行: ${recentLogs.length}件`
      });
      
      // エラーログの確認
      const errorCount = recentLogs.filter(row => row[2] === 'ERROR').length;
      if (errorCount > 0) {
        results.checks.push({
          name: '最近のエラー',
          status: 'WARNING',
          details: `直近5件中${errorCount}件のエラーを検出`
        });
      }
      
    } catch (error) {
      results.checks.push({
        name: '実行ログ確認',
        status: 'ERROR',
        details: error.toString()
      });
    }
    
    console.log('=== ヘルスチェック結果 ===');
    console.log(`実行時刻: ${results.timestamp}`);
    
    for (const check of results.checks) {
      console.log(`${check.name}: ${check.status} - ${check.details}`);
    }
    
    const hasErrors = results.checks.some(check => check.status === 'ERROR');
    const hasWarnings = results.checks.some(check => check.status === 'WARNING');
    
    if (hasErrors) {
      console.log('🔴 システムにエラーが検出されました');
    } else if (hasWarnings) {
      console.log('🟡 システムに警告が検出されました');
    } else {
      console.log('🟢 システムは正常に動作しています');
    }
    
    return results;
    
  } catch (error) {
    console.error('ヘルスチェック実行エラー:', error);
    throw error;
  }
}

/**
 * 統計情報の取得
 */
function getStatistics(days = 7) {
  try {
    const sheet = SpreadsheetApp.openById(CONFIG_SPREADSHEET_ID).getSheetByName(LOG_SHEET_NAME);
    const values = sheet.getDataRange().getValues();
    
    if (values.length <= 1) {
      console.log('統計データがありません');
      return null;
    }
    
    const cutoffDate = new Date(Date.now() - (days * 24 * 60 * 60 * 1000));
    const recentLogs = values.slice(1).filter(row => {
      const processedAt = row[1];
      return processedAt && new Date(processedAt) >= cutoffDate;
    });
    
    const stats = {
      period: `過去${days}日間`,
      totalExecutions: recentLogs.length,
      successCount: recentLogs.filter(row => row[2] === 'SUCCESS').length,
      errorCount: recentLogs.filter(row => row[2] === 'ERROR').length,
      uniqueMeetings: new Set(recentLogs.map(row => row[4])).size,
      successRate: 0
    };
    
    if (stats.totalExecutions > 0) {
      stats.successRate = Math.round((stats.successCount / stats.totalExecutions) * 100);
    }
    
    console.log('=== 統計情報 ===');
    console.log(`期間: ${stats.period}`);
    console.log(`総実行回数: ${stats.totalExecutions}`);
    console.log(`成功: ${stats.successCount}`);
    console.log(`エラー: ${stats.errorCount}`);
    console.log(`ユニーク会議数: ${stats.uniqueMeetings}`);
    console.log(`成功率: ${stats.successRate}%`);
    
    return stats;
    
  } catch (error) {
    console.error('統計情報取得エラー:', error);
    throw error;
  }
}

/**
 * 設定値の検証
 */
function validateConfiguration() {
  try {
    console.log('=== 設定検証開始 ===');
    
    const config = getConfig();
    const errors = [];
    const warnings = [];
    
    // スプレッドシートIDの検証
    try {
      const spreadsheet = getConfigSpreadsheet();
      console.log(`設定スプレッドシートにアクセス成功: ${spreadsheet.getName()}`);
    } catch (error) {
      errors.push(`設定スプレッドシートアクセスエラー: ${error.toString()}`);
    }

    // カレンダーIDの検証
    if (!config.calendarId || !config.calendarId.includes('@')) {
      errors.push('CalendarId が正しい形式ではありません');
    } else {
      try {
        const calendar = CalendarApp.getCalendarById(config.calendarId);
        console.log(`カレンダー "${calendar.getName()}" にアクセス成功`);
      } catch (error) {
        errors.push(`カレンダーアクセスエラー: ${error.toString()}`);
      }
    }
    
    // プロンプトの検証
    if (!config.minutesPrompt || config.minutesPrompt.length < 100) {
      warnings.push('MinutesPrompt が短すぎる可能性があります');
    }
    if (!config.minutesPrompt.includes('{transcript_text}')) {
      errors.push('MinutesPrompt に {transcript_text} プレースホルダーがありません');
    }
    
    // CheckHours の検証
    if (isNaN(config.checkHours) || config.checkHours < 1 || config.checkHours > 24) {
      warnings.push('CheckHours は 1〜24 の範囲で設定することを推奨します');
    }
    
    // メールアドレスの検証
    if (config.additionalRecipients) {
      const emails = config.additionalRecipients.split(',').map(email => email.trim());
      for (const email of emails) {
        if (email && !email.includes('@')) {
          warnings.push(`無効なメールアドレス形式: ${email}`);
        }
      }
    }
    
    // GCP設定の検証
    if (!GCP_PROJECT_ID || GCP_PROJECT_ID === 'your-gcp-project-id') {
      errors.push('GCP_PROJECT_ID が設定されていません');
    }
    
    // 結果の表示
    if (errors.length > 0) {
      console.log('🔴 設定エラー:');
      errors.forEach(error => console.log(`  - ${error}`));
    }
    
    if (warnings.length > 0) {
      console.log('🟡 設定警告:');
      warnings.forEach(warning => console.log(`  - ${warning}`));
    }
    
    if (errors.length === 0 && warnings.length === 0) {
      console.log('🟢 設定は正常です');
    }
    
    console.log('=== 設定検証完了 ===');
    
    return {
      valid: errors.length === 0,
      errors,
      warnings
    };
    
  } catch (error) {
    console.error('設定検証エラー:', error);
    throw error;
  }
}

/**
 * デバッグ用：Meet URLから会議コードを抽出
 */
function extractMeetingCodeFromUrl(meetUrl) {
  const match = meetUrl.match(/https:\/\/meet\.google\.com\/([a-z-]+)/);
  return match ? match[1] : null;
}

/**
 * デバッグ用：Gemini APIのレスポンス詳細表示
 */
function debugGeminiResponse(promptTemplate, transcript) {
  try {
    const prompt = promptTemplate.replace('{transcript_text}', transcript);
    console.log('=== Gemini API デバッグ ===');
    console.log('プロンプト長:', prompt.length);
    console.log('プロンプト（最初の200文字）:', prompt.substring(0, 200) + '...');
    
    const response = generateMinutes(promptTemplate, transcript);
    console.log('レスポンス長:', response.length);
    console.log('レスポンス:', response);
    
    return response;
  } catch (error) {
    console.error('Gemini API デバッグエラー:', error);
    throw error;
  }
}
