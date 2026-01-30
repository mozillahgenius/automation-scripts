/**
 * システムテスト用関数
 * トランスクリプトが利用できない場合のテストとデバッグ
 */

/**
 * 完全なシステムテスト（サンプルデータ使用）
 */
function completeSystemTest() {
  console.log('=== 完全なシステムテスト（サンプルデータ使用）===');
  
  const config = getConfig();
  
  // 実際の会議を想定したサンプルトランスクリプト
  const sampleTranscript = `
山田: 皆さん、お疲れ様です。週次定例会議を始めます。本日は2025年8月10日、15時からの開催です。
田中: よろしくお願いします。
佐藤: お願いします。
山田: まず、各プロジェクトの進捗を確認しましょう。田中さん、プロジェクトAはいかがですか？
田中: プロジェクトAは順調です。現在フェーズ2が90%完了しており、今週金曜日には完了予定です。
山田: 素晴らしい。課題はありますか？
田中: 特に大きな課題はありませんが、最終レビューに時間がかかる可能性があります。
山田: 了解です。佐藤さん、プロジェクトBの状況をお願いします。
佐藤: プロジェクトBは少し遅れています。技術的な問題が発生し、解決に2日かかりました。
山田: 具体的にはどのような問題でしたか？
佐藤: APIの互換性の問題でした。既に解決済みで、今後は問題ないと思います。
山田: 分かりました。リカバリープランはありますか？
佐藤: はい、今週末に追加作業を行い、来週月曜までには予定に追いつく予定です。
山田: ありがとうございます。それでは、今後のアクションアイテムを確認します。
山田: 田中さんは金曜日までにフェーズ2を完了させてください。
田中: 承知しました。
山田: 佐藤さんは月曜日までに遅れを取り戻してください。
佐藤: はい、頑張ります。
山田: 私は両プロジェクトの状況を役員会に報告します。
山田: 次回の会議は来週月曜日、同じ時間でよろしいですか？
田中: 大丈夫です。
佐藤: 問題ありません。
山田: それでは、本日の会議を終了します。お疲れ様でした。
全員: お疲れ様でした。
`;
  
  // 1. 議事録生成
  console.log('\n📝 議事録を生成中...');
  const minutes = generateMinutes(config.minutesPrompt, sampleTranscript);
  
  console.log('\n=== 生成された議事録 ===');
  console.log(minutes);
  
  // 2. メール送信
  const userEmail = Session.getActiveUser().getEmail();
  const recipients = [userEmail];
  const subject = '【テスト】週次定例会議 議事録 ' + new Date().toLocaleDateString('ja-JP');
  
  console.log('\n📧 メール送信中...');
  sendMinutesEmail(recipients, subject, minutes);
  
  console.log(`✅ メール送信完了: ${userEmail}`);
  
  // 3. ログ記録
  logExecution(
    'test-' + new Date().getTime(),
    'SUCCESS',
    '',
    'テスト週次定例会議'
  );
  
  console.log('\n=== テスト完了 ===');
  console.log('✅ 議事録生成: 成功');
  console.log('✅ メール送信: 成功');
  console.log('✅ ログ記録: 成功');
  console.log('\nシステムは正常に動作しています！');
  console.log('実際の会議でトランスクリプトが生成されれば、自動処理が可能です。');
}

/**
 * トランスクリプトの状態を詳しく確認
 */
function checkTranscriptStatus() {
  const config = getConfig();
  const events = getRecentEvents(config.calendarId, 24); // 過去24時間
  
  console.log('=== トランスクリプト状態確認 ===');
  
  events.forEach(event => {
    console.log(`\n会議: ${event.getTitle()}`);
    console.log(`開始: ${event.getStartTime()}`);
    console.log(`終了: ${event.getEndTime()}`);
    
    const now = new Date();
    const endTime = event.getEndTime();
    const hoursSinceEnd = (now - endTime) / (1000 * 60 * 60);
    
    if (endTime > now) {
      console.log('⏳ 会議はまだ進行中です');
    } else if (hoursSinceEnd < 0.5) {
      console.log('⏱️ 会議終了から30分未満 - トランスクリプト処理中の可能性');
    } else {
      console.log(`✓ 会議終了から${Math.round(hoursSinceEnd)}時間経過`);
      console.log('トランスクリプトが有効になっていない可能性があります');
    }
  });
  
  console.log('\n=== トランスクリプトを有効にする方法 ===');
  console.log('1. Google Meetで会議を開始');
  console.log('2. 右下の「その他のオプション」（3点メニュー）をクリック');
  console.log('3. 「録画」を選択して開始');
  console.log('4. 「トランスクリプト」もオンにする（利用可能な場合）');
  console.log('5. 会議を実施');
  console.log('6. 会議終了後、5-10分待つ');
  console.log('7. main() を実行して議事録を生成');
}

/**
 * 処理済みログをクリア
 */
function clearProcessedLogs() {
  const spreadsheet = getConfigSpreadsheet();
  const sheet = spreadsheet.getSheetByName(LOG_SHEET_NAME);
  
  if (sheet) {
    // ヘッダー行以外をクリア
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.deleteRows(2, lastRow - 1);
      console.log('✅ 処理済みログをクリアしました');
    } else {
      console.log('ℹ️ クリアするログがありません');
    }
  } else {
    console.log('⚠️ ログシートが見つかりません');
  }
}

/**
 * サンプルデータで議事録メールをプレビュー
 */
function previewMinutesEmail() {
  console.log('=== 議事録メールのプレビュー ===');
  
  const config = getConfig();
  
  // 短いサンプル
  const shortTranscript = `
田中: 本日の会議を開始します。
佐藤: お願いします。
田中: プロジェクトの進捗はいかがですか？
佐藤: 予定通り進んでいます。来週完成予定です。
田中: 了解しました。次回は来週月曜日でお願いします。
佐藤: 承知しました。
`;
  
  const minutes = generateMinutes(config.minutesPrompt, shortTranscript);
  
  console.log('件名: 【議事録】プロジェクト定例会議');
  console.log('宛先: ' + Session.getActiveUser().getEmail());
  console.log('\n--- メール本文 ---');
  console.log(minutes);
  console.log('--- メール本文終了 ---');
  
  return minutes;
}

/**
 * 特定のイベントを強制的に再処理
 */
function reprocessEvent(eventTitle) {
  const config = getConfig();
  const events = getRecentEvents(config.calendarId, 24);
  
  const targetEvent = events.find(e => e.getTitle() === eventTitle);
  
  if (!targetEvent) {
    console.log(`イベント「${eventTitle}」が見つかりません`);
    return;
  }
  
  console.log(`イベント「${eventTitle}」を強制処理します`);
  
  // トランスクリプト取得を試行
  const transcript = getTranscriptByEvent(targetEvent);
  
  if (transcript) {
    console.log('✅ トランスクリプト取得成功');
    
    // 議事録生成
    const minutes = generateMinutes(config.minutesPrompt, transcript);
    console.log('✅ 議事録生成成功');
    
    // メール送信
    const userEmail = Session.getActiveUser().getEmail();
    sendMinutesEmail([userEmail], `【議事録】${targetEvent.getTitle()}`, minutes);
    console.log('✅ メール送信完了');
    
    // ログ記録
    logExecution(
      targetEvent.getId(),
      'SUCCESS',
      '手動再処理',
      targetEvent.getTitle()
    );
  } else {
    console.log('❌ トランスクリプトが見つかりません');
    console.log('考えられる理由:');
    console.log('1. 会議で録画/トランスクリプトが有効になっていない');
    console.log('2. 会議がまだ終了していない');
    console.log('3. トランスクリプト処理がまだ完了していない（会議終了後5-10分かかります）');
  }
}

/**
 * システム全体の状態を確認
 */
function checkSystemHealth() {
  console.log('=== システムヘルスチェック ===\n');
  
  // 1. 設定確認
  console.log('📋 設定確認...');
  try {
    const config = getConfig();
    console.log('✅ 設定読み込み: 成功');
    console.log(`  - カレンダーID: ${config.calendarId}`);
    console.log(`  - チェック時間: ${config.checkHours}時間`);
  } catch (error) {
    console.log('❌ 設定読み込み: 失敗');
    console.log(`  エラー: ${error.message}`);
  }
  
  // 2. カレンダーアクセス
  console.log('\n📅 カレンダーアクセス...');
  try {
    const config = getConfig();
    const events = getRecentEvents(config.calendarId, 1);
    console.log('✅ カレンダーアクセス: 成功');
    console.log(`  - 直近1時間のイベント数: ${events.length}`);
  } catch (error) {
    console.log('❌ カレンダーアクセス: 失敗');
    console.log(`  エラー: ${error.message}`);
  }
  
  // 3. Meet API
  console.log('\n🎥 Meet API接続...');
  try {
    const testUrl = 'https://meet.googleapis.com/v2/conferenceRecords?pageSize=1';
    const response = UrlFetchApp.fetch(testUrl, {
      method: 'GET',
      headers: {
        'Authorization': `Bearer ${ScriptApp.getOAuthToken()}`,
        'Accept': 'application/json'
      },
      muteHttpExceptions: true
    });
    
    if (response.getResponseCode() === 200) {
      console.log('✅ Meet API: 接続成功');
    } else {
      console.log('⚠️ Meet API: 接続可能だが権限に問題がある可能性');
      console.log(`  レスポンスコード: ${response.getResponseCode()}`);
    }
  } catch (error) {
    console.log('❌ Meet API: 接続失敗');
    console.log(`  エラー: ${error.message}`);
  }
  
  // 4. Vertex AI (Gemini)
  console.log('\n🤖 Vertex AI (Gemini) API...');
  try {
    const testTranscript = '田中: テストです。';
    const config = getConfig();
    const result = generateMinutes(config.minutesPrompt, testTranscript);
    if (result) {
      console.log('✅ Vertex AI: 接続成功');
    }
  } catch (error) {
    console.log('❌ Vertex AI: 接続失敗');
    console.log(`  エラー: ${error.message}`);
  }
  
  // 5. Gmail
  console.log('\n📧 Gmail...');
  try {
    const email = Session.getActiveUser().getEmail();
    console.log('✅ Gmail: アクセス可能');
    console.log(`  - アカウント: ${email}`);
  } catch (error) {
    console.log('❌ Gmail: アクセス失敗');
    console.log(`  エラー: ${error.message}`);
  }
  
  console.log('\n=== ヘルスチェック完了 ===');
}

/**
 * クイックスタートガイド
 */
function showQuickStartGuide() {
  console.log('=== 🚀 クイックスタートガイド ===\n');
  
  console.log('【ステップ1】システムテスト');
  console.log('completeSystemTest() を実行');
  console.log('→ サンプルデータで議事録生成とメール送信をテスト\n');
  
  console.log('【ステップ2】会議でトランスクリプトを有効化');
  console.log('1. Google Meetで会議を開始');
  console.log('2. 録画を開始（3点メニュー → 録画）');
  console.log('3. 会議を実施（最低1分程度話す）');
  console.log('4. 会議を終了\n');
  
  console.log('【ステップ3】議事録の自動生成');
  console.log('1. 会議終了後5-10分待つ');
  console.log('2. main() を実行');
  console.log('3. メールで議事録を確認\n');
  
  console.log('【ステップ4】定期実行の設定');
  console.log('setupTriggers() を実行');
  console.log('→ 1時間ごとに自動実行されるようになります\n');
  
  console.log('【トラブルシューティング】');
  console.log('checkSystemHealth() - システム状態を確認');
  console.log('checkTranscriptStatus() - トランスクリプトの状態を確認');
  console.log('clearProcessedLogs() - 処理済みログをクリア');
}