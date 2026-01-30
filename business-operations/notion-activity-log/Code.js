/**
 * Notion API を使用してページを自動取得し、定期的にログを記録するシステム
 * + ユーザー別活動分析機能
 */

// Notion設定
const NOTION_TOKEN = PropertiesService.getScriptProperties().getProperty('NOTION_API_TOKEN');
const NOTION_VERSION = '2022-06-28';

// ログ設定
const LOG_SHEET_NAME = 'NotionLog'; // ログ用シート名
const PAGE_LIST_SHEET_NAME = 'PageList'; // ページリスト用シート名
const USER_ANALYTICS_SHEET_NAME = 'UserAnalytics'; // ユーザー分析用シート名
const USER_ACTIVITY_SHEET_NAME = 'UserActivity'; // ユーザー活動履歴用シート名

/**
 * 初期セットアップ - 全シートを作成し、ヘッダーを設定
 */
function setupNotionMonitoring() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // ページリストシートの作成/設定
  let pageListSheet;
  try {
    pageListSheet = ss.getSheetByName(PAGE_LIST_SHEET_NAME);
  } catch (e) {
    pageListSheet = ss.insertSheet(PAGE_LIST_SHEET_NAME);
  }
  
  // ヘッダー設定（ページリスト）
  pageListSheet.getRange(1, 1, 1, 4).setValues([
    ['Page ID', 'Page Title', '最終編集日時', '最終編集者']
  ]);
  pageListSheet.getRange(1, 1, 1, 4).setFontWeight('bold');
  
  // ログシートの作成/設定
  let logSheet;
  try {
    logSheet = ss.getSheetByName(LOG_SHEET_NAME);
  } catch (e) {
    logSheet = ss.insertSheet(LOG_SHEET_NAME);
  }
  
  // ヘッダー設定（ログ）
  logSheet.getRange(1, 1, 1, 6).setValues([
    ['記録日時', 'Page ID', 'Page Title', '最終編集日時', '最終編集者', '変更検出']
  ]);
  logSheet.getRange(1, 1, 1, 6).setFontWeight('bold');
  
  // ユーザー分析シートの作成/設定
  let userAnalyticsSheet;
  try {
    userAnalyticsSheet = ss.getSheetByName(USER_ANALYTICS_SHEET_NAME);
  } catch (e) {
    userAnalyticsSheet = ss.insertSheet(USER_ANALYTICS_SHEET_NAME);
  }
  
  // ヘッダー設定（ユーザー分析）
  userAnalyticsSheet.getRange(1, 1, 1, 8).setValues([
    ['ユーザー名', '編集ページ数', '最終活動日', '今日の編集数', '今週の編集数', '今月の編集数', '総編集回数', '平均編集間隔(日)']
  ]);
  userAnalyticsSheet.getRange(1, 1, 1, 8).setFontWeight('bold');
  
  // ユーザー活動履歴シートの作成/設定
  let userActivitySheet;
  try {
    userActivitySheet = ss.getSheetByName(USER_ACTIVITY_SHEET_NAME);
  } catch (e) {
    userActivitySheet = ss.insertSheet(USER_ACTIVITY_SHEET_NAME);
  }
  
  // ヘッダー設定（ユーザー活動履歴）
  userActivitySheet.getRange(1, 1, 1, 5).setValues([
    ['記録日時', 'ユーザー名', 'Page Title', '活動種別', 'Page ID']
  ]);
  userActivitySheet.getRange(1, 1, 1, 5).setFontWeight('bold');
  
  SpreadsheetApp.getUi().alert('セットアップ完了！次にdatabaseIdを設定して getAllPagesFromDatabase() を実行してください。');
}

/**
 * Notionデータベースから全ページを取得してシートに記録
 * 注意: DATABASE_ID を実際のデータベースIDに変更してください
 */
function getAllPagesFromDatabase() {
  // ここに監視したいNotionデータベースのIDを設定
  const DATABASE_ID = '1b70009db84b8075a22ae6ee74aebff3'; // 実際のデータベースIDに変更してください
  
  if (DATABASE_ID === 'YOUR_DATABASE_ID_HERE') {
    SpreadsheetApp.getUi().alert('DATABASE_ID を実際の値に変更してください');
    return;
  }
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(PAGE_LIST_SHEET_NAME);
  
  try {
    // 既存のデータをクリア（ヘッダー以外）
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.getRange(2, 1, lastRow - 1, 4).clearContent();
    }
    
    // データベースからページを取得
    const pages = fetchAllPagesFromDatabase(DATABASE_ID);
    
    if (pages.length === 0) {
      SpreadsheetApp.getUi().alert('ページが見つかりませんでした');
      return;
    }
    
    // シートにデータを書き込み
    const data = [];
    for (const page of pages) {
      const pageId = page.id;
      const title = extractPageTitle(page);
      const lastEditedTime = page.last_edited_time || '';
      
      // ユーザー名を取得
      let lastEditorName = '';
      const lastEditorUserId = (page.last_edited_by && page.last_edited_by.id) || '';
      if (lastEditorUserId) {
        try {
          const userInfo = fetchNotionUser(lastEditorUserId);
          lastEditorName = userInfo.name || '';
        } catch (e) {
          lastEditorName = 'Unknown User';
        }
      }
      
      data.push([pageId, title, lastEditedTime, lastEditorName]);
    }
    
    if (data.length > 0) {
      sheet.getRange(2, 1, data.length, 4).setValues(data);
    }
    
    // 初回取得時にユーザー分析を実行
    updateUserAnalytics();
    
    SpreadsheetApp.getUi().alert(`${pages.length} ページを取得しました`);
    
  } catch (error) {
    SpreadsheetApp.getUi().alert('エラー: ' + error.message);
    console.error(error);
  }
}

/**
 * ページリストの情報を更新し、変更があればログに記録
 */
function updateAndLogNotionPages() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const pageListSheet = ss.getSheetByName(PAGE_LIST_SHEET_NAME);
  const logSheet = ss.getSheetByName(LOG_SHEET_NAME);
  const userActivitySheet = ss.getSheetByName(USER_ACTIVITY_SHEET_NAME);
  
  const lastRow = pageListSheet.getLastRow();
  if (lastRow < 2) {
    console.log('監視対象のページがありません');
    return;
  }
  
  // 現在の記録日時
  const now = new Date();
  
  // 既存データを取得
  const existingData = pageListSheet.getRange(2, 1, lastRow - 1, 4).getValues();
  
  for (let i = 0; i < existingData.length; i++) {
    const [pageId, oldTitle, oldLastEditedTime, oldLastEditor] = existingData[i];
    
    if (!pageId) continue;
    
    try {
      // 最新のページ情報を取得
      const pageInfo = fetchNotionPage(pageId.toString().trim());
      const newTitle = extractPageTitle(pageInfo);
      const newLastEditedTime = pageInfo.last_edited_time || '';
      
      // ユーザー名を取得
      let newLastEditor = '';
      const lastEditorUserId = (pageInfo.last_edited_by && pageInfo.last_edited_by.id) || '';
      if (lastEditorUserId) {
        try {
          const userInfo = fetchNotionUser(lastEditorUserId);
          newLastEditor = userInfo.name || '';
        } catch (e) {
          newLastEditor = 'Unknown User';
        }
      }
      
      // 変更検出
      let hasChanged = false;
      let changeNote = '';
      let activityType = '';
      
      if (oldLastEditedTime !== newLastEditedTime) {
        hasChanged = true;
        changeNote = '編集日時変更';
        activityType = 'ページ編集';
      }
      
      if (oldLastEditor !== newLastEditor) {
        if (hasChanged) {
          changeNote += ' + 編集者変更';
        } else {
          hasChanged = true;
          changeNote = '編集者変更';
        }
        activityType = 'ページ編集';
      }
      
      if (oldTitle !== newTitle) {
        if (hasChanged) {
          changeNote += ' + タイトル変更';
        } else {
          hasChanged = true;
          changeNote = 'タイトル変更';
        }
        activityType = 'タイトル変更';
      }
      
      // ページリストを更新
      pageListSheet.getRange(i + 2, 2, 1, 3).setValues([[newTitle, newLastEditedTime, newLastEditor]]);
      
      // 変更があった場合、またはログが必要な場合はログに記録
      const shouldLog = hasChanged || isPeriodicLog();
      
      if (shouldLog) {
        const logData = [
          now,
          pageId,
          newTitle,
          newLastEditedTime,
          newLastEditor,
          hasChanged ? changeNote : '定期記録'
        ];
        
        // ログシートに追加
        const logLastRow = logSheet.getLastRow();
        logSheet.getRange(logLastRow + 1, 1, 1, 6).setValues([logData]);
        
        // ユーザー活動履歴に追加（変更があった場合のみ）
        if (hasChanged && newLastEditor && newLastEditor !== 'Unknown User') {
          const activityData = [
            now,
            newLastEditor,
            newTitle,
            activityType,
            pageId
          ];
          
          const activityLastRow = userActivitySheet.getLastRow();
          userActivitySheet.getRange(activityLastRow + 1, 1, 1, 5).setValues([activityData]);
        }
      }
      
    } catch (error) {
      console.error(`ページ ${pageId} の処理でエラー:`, error);
      
      // エラーログを記録
      const logData = [
        now,
        pageId,
        oldTitle,
        '',
        '',
        'エラー: ' + error.message
      ];
      
      const logLastRow = logSheet.getLastRow();
      logSheet.getRange(logLastRow + 1, 1, 1, 6).setValues([logData]);
    }
  }
  
  // ページ更新後にユーザー分析を更新
  updateUserAnalytics();
}

/**
 * ユーザー別活動分析を更新
 */
function updateUserAnalytics() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const pageListSheet = ss.getSheetByName(PAGE_LIST_SHEET_NAME);
  const userActivitySheet = ss.getSheetByName(USER_ACTIVITY_SHEET_NAME);
  const userAnalyticsSheet = ss.getSheetByName(USER_ANALYTICS_SHEET_NAME);
  
  // 現在の日時情報を取得
  const now = new Date();
  const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  const weekAgo = new Date(today.getTime() - 7 * 24 * 60 * 60 * 1000);
  const monthAgo = new Date(today.getTime() - 30 * 24 * 60 * 60 * 1000);
  
  // 現在アクティブなページからユーザー情報を収集
  const pageData = pageListSheet.getRange(2, 1, Math.max(pageListSheet.getLastRow() - 1, 0), 4).getValues();
  const userStats = {};
  
  // ページリストからユーザー別統計を作成
  pageData.forEach(([pageId, title, lastEditedTime, editor]) => {
    if (!editor || editor === 'Unknown User') return;
    
    if (!userStats[editor]) {
      userStats[editor] = {
        name: editor,
        editedPages: new Set(),
        lastActivity: null,
        todayEdits: 0,
        weekEdits: 0,
        monthEdits: 0,
        totalEdits: 0,
        editDates: []
      };
    }
    
    userStats[editor].editedPages.add(pageId);
    
    if (lastEditedTime) {
      const editDate = new Date(lastEditedTime);
      userStats[editor].editDates.push(editDate);
      
      if (!userStats[editor].lastActivity || editDate > userStats[editor].lastActivity) {
        userStats[editor].lastActivity = editDate;
      }
    }
  });
  
  // 活動履歴からより詳細な統計を取得
  const activityLastRow = userActivitySheet.getLastRow();
  if (activityLastRow > 1) {
    const activityData = userActivitySheet.getRange(2, 1, activityLastRow - 1, 5).getValues();
    
    activityData.forEach(([recordTime, userName, pageTitle, activityType, pageId]) => {
      if (!userName || userName === 'Unknown User') return;
      
      if (!userStats[userName]) {
        userStats[userName] = {
          name: userName,
          editedPages: new Set(),
          lastActivity: null,
          todayEdits: 0,
          weekEdits: 0,
          monthEdits: 0,
          totalEdits: 0,
          editDates: []
        };
      }
      
      const activityDate = new Date(recordTime);
      userStats[userName].totalEdits++;
      userStats[userName].editDates.push(activityDate);
      
      if (activityDate >= today) {
        userStats[userName].todayEdits++;
      }
      if (activityDate >= weekAgo) {
        userStats[userName].weekEdits++;
      }
      if (activityDate >= monthAgo) {
        userStats[userName].monthEdits++;
      }
      
      if (!userStats[userName].lastActivity || activityDate > userStats[userName].lastActivity) {
        userStats[userName].lastActivity = activityDate;
      }
    });
  }
  
  // 分析結果をシートに書き込み
  const analyticsData = [];
  
  Object.values(userStats).forEach(user => {
    const editedPagesCount = user.editedPages.size;
    const lastActivityStr = user.lastActivity ? user.lastActivity.toLocaleString('ja-JP') : '';
    
    // 平均編集間隔を計算
    let avgEditInterval = 0;
    if (user.editDates.length > 1) {
      const sortedDates = user.editDates.sort((a, b) => a - b);
      let totalInterval = 0;
      for (let i = 1; i < sortedDates.length; i++) {
        totalInterval += (sortedDates[i] - sortedDates[i-1]) / (1000 * 60 * 60 * 24); // 日単位
      }
      avgEditInterval = Math.round(totalInterval / (sortedDates.length - 1) * 10) / 10;
    }
    
    analyticsData.push([
      user.name,
      editedPagesCount,
      lastActivityStr,
      user.todayEdits,
      user.weekEdits,
      user.monthEdits,
      user.totalEdits,
      avgEditInterval
    ]);
  });
  
  // 総編集回数でソート（降順）
  analyticsData.sort((a, b) => b[6] - a[6]);
  
  // 既存データをクリア
  const analyticsLastRow = userAnalyticsSheet.getLastRow();
  if (analyticsLastRow > 1) {
    userAnalyticsSheet.getRange(2, 1, analyticsLastRow - 1, 8).clearContent();
  }
  
  // 新しいデータを書き込み
  if (analyticsData.length > 0) {
    userAnalyticsSheet.getRange(2, 1, analyticsData.length, 8).setValues(analyticsData);
    
    // 条件付き書式を適用（活動レベルに応じて色分け）
    const dataRange = userAnalyticsSheet.getRange(2, 4, analyticsData.length, 4); // 今日〜総編集回数の列
    
    // 今日の編集数の色分け
    const todayRange = userAnalyticsSheet.getRange(2, 4, analyticsData.length, 1);
    todayRange.setBackgroundColors(analyticsData.map(row => {
      const todayEdits = row[3];
      if (todayEdits >= 3) return ['#d4edda']; // 緑
      if (todayEdits >= 1) return ['#fff3cd']; // 黄
      return ['#ffffff']; // 白
    }));
    
    // 総編集回数の色分け
    const totalRange = userAnalyticsSheet.getRange(2, 7, analyticsData.length, 1);
    const maxTotal = Math.max(...analyticsData.map(row => row[6]));
    totalRange.setBackgroundColors(analyticsData.map(row => {
      const total = row[6];
      const ratio = total / maxTotal;
      if (ratio >= 0.8) return ['#d4edda']; // 緑
      if (ratio >= 0.5) return ['#fff3cd']; // 黄
      if (ratio >= 0.2) return ['#f8d7da']; // 薄赤
      return ['#ffffff']; // 白
    }));
  }
  
  console.log(`ユーザー分析を更新しました: ${analyticsData.length} ユーザー`);
}

/**
 * ユーザー活動レポートを生成（手動実行用）
 */
function generateUserActivityReport() {
  updateUserAnalytics();
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const userAnalyticsSheet = ss.getSheetByName(USER_ANALYTICS_SHEET_NAME);
  
  // レポート用の新しいシートを作成
  const reportSheetName = `活動レポート_${new Date().toISOString().split('T')[0]}`;
  let reportSheet;
  
  try {
    reportSheet = ss.insertSheet(reportSheetName);
  } catch (e) {
    reportSheet = ss.getSheetByName(reportSheetName);
    reportSheet.clear();
  }
  
  // レポートヘッダー
  const reportDate = new Date().toLocaleString('ja-JP');
  reportSheet.getRange(1, 1).setValue(`Notion ユーザー活動レポート - ${reportDate}`);
  reportSheet.getRange(1, 1).setFontSize(14).setFontWeight('bold');
  
  // 分析データをコピー
  const analyticsData = userAnalyticsSheet.getDataRange().getValues();
  reportSheet.getRange(3, 1, analyticsData.length, analyticsData[0].length).setValues(analyticsData);
  
  // 書式設定
  reportSheet.getRange(3, 1, 1, analyticsData[0].length).setFontWeight('bold');
  reportSheet.autoResizeColumns(1, analyticsData[0].length);
  
  // サマリー統計を追加
  const summaryStartRow = 3 + analyticsData.length + 2;
  reportSheet.getRange(summaryStartRow, 1).setValue('=== サマリー統計 ===');
  reportSheet.getRange(summaryStartRow, 1).setFontWeight('bold');
  
  const totalUsers = analyticsData.length - 1; // ヘッダーを除く
  const activeToday = analyticsData.slice(1).filter(row => row[3] > 0).length;
  const activeWeek = analyticsData.slice(1).filter(row => row[4] > 0).length;
  
  reportSheet.getRange(summaryStartRow + 1, 1, 3, 2).setValues([
    ['総ユーザー数:', totalUsers],
    ['今日アクティブ:', activeToday],
    ['今週アクティブ:', activeWeek]
  ]);
  
  SpreadsheetApp.getUi().alert(`ユーザー活動レポートを生成しました: ${reportSheetName}`);
}

/**
 * 定期実行用の関数（変更検出 + 定期ログ + ユーザー分析）
 */
function scheduledUpdate() {
  updateAndLogNotionPages();
}

/**
 * 定期的なトリガー設定（4時間ごと実行）
 */
function setupPeriodicTrigger() {
  // 既存のトリガーを削除
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'scheduledUpdate') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  // 新しいトリガーを設定（4時間ごと実行）
  ScriptApp.newTrigger('scheduledUpdate')
    .timeBased()
    .everyHours(4)
    .create();
    
  SpreadsheetApp.getUi().alert('定期実行トリガーを設定しました（4時間ごと実行）');
}

/**
 * 定期的なトリガー設定（毎日実行）
 */
function setupDailyTrigger() {
  // 既存のトリガーを削除
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'scheduledUpdate') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  // 新しいトリガーを設定（毎日午前9時実行）
  ScriptApp.newTrigger('scheduledUpdate')
    .timeBased()
    .everyDays(1)
    .atHour(9)
    .create();
    
  SpreadsheetApp.getUi().alert('定期実行トリガーを設定しました（毎日午前9時実行）');
}

// ===== ユーティリティ関数 =====

/**
 * Notionデータベースから全ページを取得
 */
function fetchAllPagesFromDatabase(databaseId) {
  const allPages = [];
  let hasMore = true;
  let nextCursor = null;
  
  while (hasMore) {
    const url = `https://api.notion.com/v1/databases/${databaseId}/query`;
    const payload = {
      page_size: 100
    };
    
    if (nextCursor) {
      payload.start_cursor = nextCursor;
    }
    
    const options = {
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${NOTION_TOKEN}`,
        'Notion-Version': NOTION_VERSION,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };
    
    const response = UrlFetchApp.fetch(url, options);
    const code = response.getResponseCode();
    
    if (code !== 200) {
      throw new Error(`データベース取得エラー (Status: ${code}): ${response.getContentText()}`);
    }
    
    const data = JSON.parse(response.getContentText());
    allPages.push(...data.results);
    
    hasMore = data.has_more;
    nextCursor = data.next_cursor;
  }
  
  return allPages;
}

/**
 * ページからタイトルを抽出
 */
function extractPageTitle(page) {
  if (page.properties) {
    // データベースページの場合、titleプロパティを探す
    for (const [key, value] of Object.entries(page.properties)) {
      if (value.type === 'title' && value.title && value.title.length > 0) {
        return value.title.map(t => t.plain_text).join('');
      }
    }
  }
  
  // 通常のページの場合
  if (page.object === 'page' && page.parent && page.parent.type === 'workspace') {
    // ワークスペース直下のページ
    return 'Untitled';
  }
  
  return 'No Title';
}

/**
 * 定期ログが必要かどうかの判定（4時間ごとに記録）
 */
function isPeriodicLog() {
  // 現在時刻
  const now = new Date();
  const hour = now.getHours();
  
  // 4時間ごと（0時、4時、8時、12時、16時、20時）にログを記録
  return hour % 4 === 0;
}

/**
 * Notion API からページ情報を取得する（元のコードから）
 */
function fetchNotionPage(pageId) {
  const url = `https://api.notion.com/v1/pages/${pageId}`;
  const options = {
    method: 'get',
    headers: {
      'Authorization': `Bearer ${NOTION_TOKEN}`,
      'Notion-Version': NOTION_VERSION,
      'Content-Type': 'application/json'
    },
    muteHttpExceptions: true
  };
  const response = UrlFetchApp.fetch(url, options);
  const code = response.getResponseCode();
  if (code !== 200) {
    throw new Error(`Notion ページ取得エラー (Status: ${code})`);
  }
  return JSON.parse(response.getContentText());
}

/**
 * Notion API からユーザー情報を取得する（元のコードから）
 */
function fetchNotionUser(userId) {
  const url = `https://api.notion.com/v1/users/${userId}`;
  const options = {
    method: 'get',
    headers: {
      'Authorization': `Bearer ${NOTION_TOKEN}`,
      'Notion-Version': NOTION_VERSION,
      'Content-Type': 'application/json'
    },
    muteHttpExceptions: true
  };
  const response = UrlFetchApp.fetch(url, options);
  const code = response.getResponseCode();
  if (code !== 200) {
    throw new Error(`Notion ユーザー取得エラー (Status: ${code})`);
  }
  return JSON.parse(response.getContentText());
}