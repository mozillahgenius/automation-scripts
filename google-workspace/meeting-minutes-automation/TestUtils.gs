/**
 * テスト用ユーティリティ関数
 * システムのテストと検証用の関数群
 */

/**
 * 議事録システムの基本機能テスト（イベントオブジェクトなし）
 */
function testBasicFunctions() {
  console.log('=== 基本機能テスト開始 ===');
  
  try {
    // 1. 設定の読み込みテスト
    console.log('1. 設定読み込みテスト');
    const config = getConfig();
    console.log('✓ 設定読み込み成功');
    console.log('  カレンダーID:', config.calendarId || 'デフォルト');
    console.log('  チェック時間:', config.checkHours + '時間');
    
    // 2. カレンダーアクセステスト
    console.log('\n2. カレンダーアクセステスト');
    try {
      const calendar = config.calendarId && config.calendarId !== 'your-calendar-id@group.calendar.google.com' 
        ? CalendarApp.getCalendarById(config.calendarId)
        : CalendarApp.getDefaultCalendar();
      
      if (calendar) {
        console.log('✓ カレンダーアクセス成功');
        console.log('  カレンダー名:', calendar.getName());
      } else {
        console.log('✗ カレンダーが見つかりません');
      }
    } catch (error) {
      console.log('✗ カレンダーアクセスエラー:', error.message);
    }
    
    // 3. Vertex AI テスト（サンプルテキスト）
    console.log('\n3. Vertex AI テスト');
    const sampleTranscript = `
参加者A: おはようございます。本日の定例会議を開始します。
参加者B: おはようございます。今週の進捗を報告させていただきます。
参加者A: はい、お願いします。
参加者B: プロジェクトAは予定通り進行中で、来週には完了予定です。
参加者A: 了解しました。プロジェクトBはどうですか？
参加者B: プロジェクトBは少し遅れていますが、今週中にキャッチアップする予定です。
参加者A: わかりました。サポートが必要な場合は連絡してください。
参加者B: ありがとうございます。
参加者A: それでは、次回は来週の同じ時間にお願いします。
参加者B: 承知しました。ありがとうございました。
    `;
    
    try {
      const testMinutes = generateMinutes(config.minutesPrompt, sampleTranscript);
      console.log('✓ 議事録生成成功');
      console.log('  生成文字数:', testMinutes.length);
      console.log('  先頭100文字:', testMinutes.substring(0, 100) + '...');
    } catch (error) {
      console.log('✗ 議事録生成エラー:', error.message);
    }
    
    // 4. メール送信機能テスト（ドライラン）
    console.log('\n4. メール送信機能テスト');
    try {
      const userEmail = Session.getActiveUser().getEmail();
      console.log('✓ メール送信準備OK');
      console.log('  送信元アドレス:', userEmail);
    } catch (error) {
      console.log('✗ メールアクセスエラー:', error.message);
    }
    
    console.log('\n=== 基本機能テスト完了 ===');
    
  } catch (error) {
    console.error('基本機能テストエラー:', error);
    throw error;
  }
}

/**
 * Meet URL抽出機能のテスト
 */
function testMeetUrlExtraction() {
  console.log('=== Meet URL抽出テスト ===');
  
  // テストケース
  const testCases = [
    {
      description: 'Google Meet URL (xxx-xxxx-xxx形式)',
      input: '会議リンク: https://meet.google.com/abc-defg-hij',
      expected: 'abc-defg-hij'
    },
    {
      description: 'Google Meet URL (短縮形式)',
      input: 'Join: meet.google.com/xyz-1234-abc',
      expected: 'xyz-1234-abc'
    },
    {
      description: 'テキスト内のMeet URL',
      input: '本日の会議はこちら: https://meet.google.com/foo-bar-baz 開始時間は10時です',
      expected: 'foo-bar-baz'
    },
    {
      description: 'URLなし',
      input: '会議の詳細は別途連絡します',
      expected: null
    }
  ];
  
  // 各テストケースを実行
  for (const testCase of testCases) {
    console.log(`\nテスト: ${testCase.description}`);
    console.log(`入力: ${testCase.input}`);
    
    // URL抽出パターン（Code.gsから同じロジック）
    const meetPatterns = [
      /https:\/\/meet\.google\.com\/([a-z]{3}-[a-z]{4}-[a-z]{3})/i,
      /https:\/\/meet\.google\.com\/([a-z0-9\-]+)/i,
      /meet\.google\.com\/([a-z]{3}-[a-z]{4}-[a-z]{3})/i,
      /meet\.google\.com\/([a-z0-9\-]+)/i
    ];
    
    let meetingCode = null;
    for (const pattern of meetPatterns) {
      const match = testCase.input.match(pattern);
      if (match && match[1]) {
        meetingCode = match[1];
        break;
      }
    }
    
    console.log(`抽出結果: ${meetingCode}`);
    console.log(`期待値: ${testCase.expected}`);
    console.log(`判定: ${meetingCode === testCase.expected ? '✓ 成功' : '✗ 失敗'}`);
  }
  
  console.log('\n=== Meet URL抽出テスト完了 ===');
}

/**
 * サンプルイベントでのトランスクリプト取得テスト
 */
function testWithSampleEvent() {
  console.log('=== サンプルイベントテスト開始 ===');
  
  try {
    // サンプルイベントオブジェクトを作成（モック）
    const mockEvent = {
      getId: function() { return 'test-event-123'; },
      getTitle: function() { return 'テスト会議'; },
      getDescription: function() { 
        return '会議リンク: https://meet.google.com/abc-defg-hij\n参加者: test@example.com';
      },
      getStartTime: function() { return new Date(); },
      getEndTime: function() { return new Date(Date.now() + 3600000); },
      getGuestList: function() { 
        return [
          { getEmail: function() { return 'test1@example.com'; } },
          { getEmail: function() { return 'test2@example.com'; } }
        ];
      }
    };
    
    console.log('モックイベント作成完了');
    console.log('  ID:', mockEvent.getId());
    console.log('  タイトル:', mockEvent.getTitle());
    console.log('  説明:', mockEvent.getDescription());
    
    // トランスクリプト取得をテスト
    console.log('\nトランスクリプト取得テスト...');
    const transcript = getTranscriptByEvent(mockEvent);
    
    if (transcript) {
      console.log('✓ トランスクリプト取得成功');
      console.log('  文字数:', transcript.length);
    } else {
      console.log('ℹ️ トランスクリプトが見つかりません（Meet APIの応答待ちの可能性）');
    }
    
  } catch (error) {
    console.error('サンプルイベントテストエラー:', error);
  }
  
  console.log('\n=== サンプルイベントテスト完了 ===');
}

/**
 * 実際のカレンダーイベントでテスト
 */
function testWithRealEvents() {
  console.log('=== 実カレンダーイベントテスト ===');
  
  try {
    const config = getConfig();
    
    // 過去24時間のイベントを取得
    console.log('過去24時間のイベントを検索中...');
    const events = getRecentEvents(config.calendarId, 24);
    
    if (events.length === 0) {
      console.log('Meet リンクを含むイベントが見つかりません');
      console.log('ヒント: カレンダーに Google Meet リンクを含むイベントを作成してください');
      return;
    }
    
    console.log(`${events.length}件のMeetイベントが見つかりました`);
    
    // 最初のイベントでテスト
    const event = events[0];
    console.log('\n最初のイベントでテスト:');
    console.log('  タイトル:', event.getTitle());
    console.log('  開始時刻:', event.getStartTime());
    
    // Meet URLの抽出テスト
    const description = event.getDescription() || '';
    console.log('  説明（最初の100文字）:', description.substring(0, 100));
    
    // トランスクリプト取得をテスト
    console.log('\nトランスクリプト取得を試行中...');
    const transcript = getTranscriptByEvent(event);
    
    if (transcript) {
      console.log('✓ トランスクリプト取得成功');
      console.log('  文字数:', transcript.length);
      console.log('  先頭100文字:', transcript.substring(0, 100));
    } else {
      console.log('ℹ️ トランスクリプトが見つかりません');
      console.log('  考えられる理由:');
      console.log('  - 会議がまだ終了していない');
      console.log('  - 録画/トランスクリプトが有効になっていない');
      console.log('  - Meet APIの処理待ち（会議終了後数分かかる場合があります）');
    }
    
  } catch (error) {
    console.error('実イベントテストエラー:', error);
  }
  
  console.log('\n=== 実カレンダーイベントテスト完了 ===');
}

/**
 * 完全なエンドツーエンドテスト
 */
function runFullTest() {
  console.log('=== 完全テスト実行 ===\n');
  
  // 1. 基本機能テスト
  testBasicFunctions();
  
  // 2. URL抽出テスト
  console.log('\n');
  testMeetUrlExtraction();
  
  // 3. サンプルイベントテスト
  console.log('\n');
  testWithSampleEvent();
  
  // 4. 実イベントテスト
  console.log('\n');
  testWithRealEvents();
  
  // 5. システムヘルスチェック
  console.log('\n');
  healthCheck();
  
  console.log('\n=== 全テスト完了 ===');
}