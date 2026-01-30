/**
 * Google Meet 検出用のユーティリティ関数
 * カレンダーイベントからMeet情報を検出する各種方法
 */

/**
 * イベントからGoogle Meet情報を取得（複数の方法を試行）
 */
function detectMeetInfo(event) {
  const meetInfo = {
    hasMeet: false,
    meetUrl: null,
    meetCode: null,
    detectionMethod: null
  };
  
  const eventTitle = event.getTitle() || '';
  const eventId = event.getId();
  
  console.log(`Meet情報検出: ${eventTitle}`);
  
  // 方法1: イベントの説明から検出
  try {
    const description = event.getDescription() || '';
    if (description.includes('meet.google.com')) {
      const meetPatterns = [
        /https:\/\/meet\.google\.com\/([a-z]{3}-[a-z]{4}-[a-z]{3})/i,
        /https:\/\/meet\.google\.com\/([a-z0-9\-]+)/i,
        /meet\.google\.com\/([a-z]{3}-[a-z]{4}-[a-z]{3})/i,
        /meet\.google\.com\/([a-z0-9\-]+)/i
      ];
      
      for (const pattern of meetPatterns) {
        const match = description.match(pattern);
        if (match && match[1]) {
          meetInfo.hasMeet = true;
          meetInfo.meetCode = match[1];
          meetInfo.meetUrl = `https://meet.google.com/${match[1]}`;
          meetInfo.detectionMethod = 'description';
          console.log(`  ✓ 説明欄からMeet URL検出: ${meetInfo.meetUrl}`);
          return meetInfo;
        }
      }
    }
  } catch (error) {
    console.log(`  説明取得エラー: ${error.message}`);
  }
  
  // 方法2: Calendar APIのConference Dataを使用
  try {
    const calendarId = event.getOriginalCalendarId ? event.getOriginalCalendarId() : 'primary';
    const eventDetails = Calendar.Events.get(calendarId, eventId.replace(/@.*/g, ''));
    
    if (eventDetails.conferenceData) {
      const conferenceData = eventDetails.conferenceData;
      console.log(`  Conference Data found:`, JSON.stringify(conferenceData).substring(0, 200));
      
      // Meet URLを探す
      if (conferenceData.entryPoints) {
        for (const entryPoint of conferenceData.entryPoints) {
          if (entryPoint.entryPointType === 'video' && entryPoint.uri) {
            if (entryPoint.uri.includes('meet.google.com')) {
              const match = entryPoint.uri.match(/meet\.google\.com\/([a-z0-9\-]+)/i);
              if (match && match[1]) {
                meetInfo.hasMeet = true;
                meetInfo.meetCode = match[1];
                meetInfo.meetUrl = entryPoint.uri;
                meetInfo.detectionMethod = 'conferenceData';
                console.log(`  ✓ Conference DataからMeet URL検出: ${meetInfo.meetUrl}`);
                return meetInfo;
              }
            }
          }
        }
      }
      
      // Conference IDから直接Meet URLを生成
      if (conferenceData.conferenceSolution && 
          conferenceData.conferenceSolution.name === 'Google Meet' && 
          conferenceData.conferenceId) {
        meetInfo.hasMeet = true;
        meetInfo.meetCode = conferenceData.conferenceId;
        meetInfo.meetUrl = `https://meet.google.com/${conferenceData.conferenceId}`;
        meetInfo.detectionMethod = 'conferenceId';
        console.log(`  ✓ Conference IDからMeet URL生成: ${meetInfo.meetUrl}`);
        return meetInfo;
      }
    }
  } catch (error) {
    console.log(`  Calendar API アクセスエラー: ${error.message}`);
  }
  
  // 方法3: HTMLリンクから検出
  try {
    const description = event.getDescription() || '';
    // HTMLタグ内のリンクも検出
    const htmlLinkPattern = /href=["']([^"']*meet\.google\.com[^"']*)/i;
    const htmlMatch = description.match(htmlLinkPattern);
    if (htmlMatch && htmlMatch[1]) {
      const url = htmlMatch[1];
      const codeMatch = url.match(/meet\.google\.com\/([a-z0-9\-]+)/i);
      if (codeMatch && codeMatch[1]) {
        meetInfo.hasMeet = true;
        meetInfo.meetCode = codeMatch[1];
        meetInfo.meetUrl = url;
        meetInfo.detectionMethod = 'htmlLink';
        console.log(`  ✓ HTMLリンクからMeet URL検出: ${meetInfo.meetUrl}`);
        return meetInfo;
      }
    }
  } catch (error) {
    console.log(`  HTMLリンク検出エラー: ${error.message}`);
  }
  
  // 方法4: Location（場所）フィールドから検出
  try {
    const location = event.getLocation() || '';
    if (location.includes('meet.google.com')) {
      const match = location.match(/meet\.google\.com\/([a-z0-9\-]+)/i);
      if (match && match[1]) {
        meetInfo.hasMeet = true;
        meetInfo.meetCode = match[1];
        meetInfo.meetUrl = location.includes('http') ? location : `https://meet.google.com/${match[1]}`;
        meetInfo.detectionMethod = 'location';
        console.log(`  ✓ 場所フィールドからMeet URL検出: ${meetInfo.meetUrl}`);
        return meetInfo;
      }
    }
  } catch (error) {
    console.log(`  場所フィールド取得エラー: ${error.message}`);
  }
  
  if (!meetInfo.hasMeet) {
    console.log(`  ✗ Meet情報が見つかりません`);
  }
  
  return meetInfo;
}

/**
 * デバッグ用: すべてのイベントのMeet情報を確認
 */
function debugAllEvents() {
  console.log('=== 全イベントのMeet情報デバッグ ===');
  
  const config = getConfig();
  const now = new Date();
  const startTime = new Date(now.getTime() - (24 * 60 * 60 * 1000)); // 24時間前
  
  let calendar;
  if (!config.calendarId || config.calendarId === 'your-calendar-id@group.calendar.google.com') {
    calendar = CalendarApp.getDefaultCalendar();
  } else {
    calendar = CalendarApp.getCalendarById(config.calendarId);
  }
  
  const events = calendar.getEvents(startTime, now);
  console.log(`検査するイベント数: ${events.length}`);
  
  let meetEventsCount = 0;
  
  events.forEach((event, index) => {
    console.log(`\n--- イベント ${index + 1} ---`);
    console.log(`タイトル: ${event.getTitle()}`);
    console.log(`開始: ${event.getStartTime()}`);
    
    const meetInfo = detectMeetInfo(event);
    
    if (meetInfo.hasMeet) {
      meetEventsCount++;
      console.log(`🟢 Meet情報あり`);
      console.log(`  URL: ${meetInfo.meetUrl}`);
      console.log(`  Code: ${meetInfo.meetCode}`);
      console.log(`  検出方法: ${meetInfo.detectionMethod}`);
    } else {
      console.log(`🔴 Meet情報なし`);
      
      // デバッグ情報を表示
      const description = event.getDescription() || '';
      const location = event.getLocation() || '';
      
      if (description || location) {
        console.log(`  デバッグ情報:`);
        if (description) {
          console.log(`    説明: ${description.substring(0, 100)}${description.length > 100 ? '...' : ''}`);
        }
        if (location) {
          console.log(`    場所: ${location}`);
        }
      }
    }
  });
  
  console.log(`\n=== サマリー ===`);
  console.log(`総イベント数: ${events.length}`);
  console.log(`Meetイベント数: ${meetEventsCount}`);
  console.log(`Meet検出率: ${events.length > 0 ? Math.round(meetEventsCount / events.length * 100) : 0}%`);
}

/**
 * テスト用: サンプルイベントでMeet情報を追加
 */
function createTestEventWithMeet() {
  const calendar = CalendarApp.getDefaultCalendar();
  const now = new Date();
  const startTime = new Date(now.getTime() + (60 * 60 * 1000)); // 1時間後
  const endTime = new Date(startTime.getTime() + (60 * 60 * 1000)); // 2時間後
  
  const event = calendar.createEvent(
    'テスト会議（自動生成）',
    startTime,
    endTime,
    {
      description: 'このイベントはテスト用です。\nMeet URL: https://meet.google.com/abc-defg-hij',
      location: 'https://meet.google.com/abc-defg-hij'
    }
  );
  
  console.log('テストイベントを作成しました:');
  console.log('  タイトル:', event.getTitle());
  console.log('  開始時刻:', event.getStartTime());
  console.log('  Meet URL:', 'https://meet.google.com/abc-defg-hij');
  
  return event.getId();
}