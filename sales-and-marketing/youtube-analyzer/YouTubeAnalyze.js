/**
 * YouTube動画管理スクリプト - 完全版（Analytics対応・修正版v3）
 * 設定シート対応、チャンネル取得、タグ管理、サムネイル取得、定期実行対応
 * 動画URL取得機能付き + YouTube Analytics API対応
 * ブランドアカウント対応版
 * 
 * 【事前準備】
 * 1. Google Apps Scriptエディタで「サービス」→「YouTube Data API v3」を追加
 * 2. Google Apps Scriptエディタで「サービス」→「YouTube Analytics API」を追加
 * 3. Google Cloud ConsoleでYouTube Analytics APIを有効化
 */

// ==================== 設定 ====================
var CONFIG = {
  SHEET_NAME: 'YouTube管理',
  CONFIG_SHEET_NAME: '設定',
  THUMBNAIL_SIZE: 'medium',
  AUTO_FETCH_INTERVAL_HOURS: 6,
  ANALYTICS_DAYS: 28
};

// カラム定義（1始まり）- YouTube管理シート
var COLUMNS = {
  VIDEO_ID: 1,
  VIDEO_URL: 2,
  CURRENT_TITLE: 3,
  CURRENT_DESCRIPTION: 4,
  CURRENT_TAGS: 5,
  NEW_TITLE: 6,
  NEW_DESCRIPTION: 7,
  NEW_TAGS: 8,
  VIEW_COUNT: 9,
  LIKE_COUNT: 10,
  COMMENT_COUNT: 11,
  PUBLISHED_AT: 12,
  AVG_VIEW_DURATION: 13,
  AVG_VIEW_PERCENTAGE: 14,
  SUBSCRIBERS_GAINED: 15,
  SUBSCRIBERS_LOST: 16,
  EST_MINUTES_WATCHED: 17,
  THUMBNAIL_URL: 18,
  THUMBNAIL_IMAGE: 19,
  CATEGORY_ID: 20,
  UPDATE_FLAG: 21,
  LAST_FETCHED: 22,
  IMPROVEMENT_COMMENT: 23
};

// 設定シートのカラム定義
var CONFIG_COLUMNS = {
  CHANNEL_ID: 1,
  CHANNEL_NAME: 2,
  PLAYLIST_ID: 3,
  ENABLED: 4,
  LAST_FETCHED: 5,
  NOTE: 6
};

// 設定シートの一般設定（右側）の定義
var GENERAL_SETTINGS_COL = 8;  // H列から開始

// ==================== メニュー ====================
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.Menu('YouTube管理')
    .addItem('動画情報を取得', 'fetchVideoInfo')
    .addItem('Analytics情報を取得', 'fetchAnalyticsInfo')
    .addItem('全情報を取得（動画+Analytics）', 'fetchAllInfo')
    .addItem('YouTubeに反映（更新）', 'updateVideoDetails')
    .addSeparator()
    .addSubMenu(ui.createMenu('チャンネル')
      .addItem('自分のチャンネルから動画を取得', 'fetchMyChannelVideos')
      .addItem('チャンネルIDから動画を取得', 'fetchChannelVideosById')
      .addItem('チャンネル情報を表示', 'showChannelInfo'))
    .addSubMenu(ui.createMenu('編集サポート')
      .addItem('現在値を編集列にコピー', 'copyCurrentToNew')
      .addItem('差分をハイライト', 'highlightDifferences')
      .addItem('ハイライトをクリア', 'clearHighlights'))
    .addSeparator()
    .addSubMenu(ui.createMenu('サムネイル')
      .addItem('サムネイルURLを取得', 'fetchThumbnails')
      .addItem('サムネイル画像を表示', 'showThumbnailImages'))
    .addSeparator()
    .addSubMenu(ui.createMenu('定期実行設定')
      .addItem('設定シートにチャンネルを追加', 'addChannelToConfig')
      .addItem('設定シートを開く', 'openConfigSheet')
      .addSeparator()
      .addItem('定期実行を開始', 'startAutoFetch')
      .addItem('定期実行を停止', 'stopAutoFetch')
      .addItem('トリガー状態を確認', 'checkTriggerStatus')
      .addSeparator()
      .addItem('今すぐ手動実行', 'runAutoFetchManually'))
    .addSubMenu(ui.createMenu('初期設定')
      .addItem('初期セットアップ（全シート作成）', 'initialSetup')
      .addItem('設定シートのみ作成', 'createConfigSheet'))
    .addSeparator()
    .addItem('ヘルプ', 'showHelp')
    .addToUi();
}

// ==================== 初期セットアップ ====================
function initialSetup() {
  createMainSheet();
  createConfigSheet();
  showAlert('初期セットアップが完了しました！\n\n' +
    '【作成されたシート】\n' +
    '・YouTube管理: 動画一覧と編集\n' +
    '・設定: チャンネル情報と各種設定\n\n' +
    '【次のステップ】\n' +
    '1. 設定シートにチャンネルを追加\n' +
    '2. チャンネルから動画を取得\n' +
    '3. 動画情報を取得\n\n' +
    '【重要】YouTube Analytics APIを使用するには、\n' +
    'Google Apps Scriptエディタで「サービス」→\n' +
    '「YouTube Analytics API」を追加してください。\n\n' +
    '【ブランドアカウントの場合】\n' +
    '設定シートの「AnalyticsチャンネルID」に\n' +
    'ブランドアカウントのチャンネルIDを設定してください。');
}

function createMainSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = CONFIG.SHEET_NAME;

  Logger.log('createMainSheet: シート名 = ' + sheetName);

  var sheet = ss.getSheetByName(sheetName);

  if (sheet) {
    var ui = SpreadsheetApp.getUi();
    var response = ui.alert(
      '確認',
      sheetName + 'シートは既に存在します。再設定しますか？',
      ui.ButtonSet.YES_NO
    );
    if (response !== ui.Button.YES) return;
    sheet.clear();
  } else {
    // 明示的に名前を指定してシートを作成
    try {
      sheet = ss.insertSheet(sheetName);
      Logger.log('シート作成成功: ' + sheet.getName());
    } catch (e) {
      Logger.log('insertSheet エラー: ' + e.message);
      sheet = ss.insertSheet();
      sheet.setName(sheetName);
      Logger.log('setName後: ' + sheet.getName());
    }
  }
  
  var headerRow1 = [
    '', '', '【現在の値】', '', '', '【修正後】', '', '',
    '【基本統計】', '', '', '',
    '【Analytics統計】', '', '', '', '',
    '【サムネイル】', '', '【管理】', '', '', ''
  ];
  
  var headerRow2 = [
    '動画ID', '動画URL',
    '現在のタイトル', '現在の説明', '現在のタグ',
    '新タイトル', '新説明', '新タグ',
    '再生回数', '高評価数', 'コメント数', '公開日',
    '平均視聴時間', '視聴維持率(%)', '登録者増加', '登録者減少', '総視聴時間(分)',
    'サムネイルURL', 'サムネイル',
    'カテゴリID', '更新フラグ', '最終取得日時', '改善コメント'
  ];
  
  sheet.getRange(1, 1, 1, headerRow1.length).setValues([headerRow1]);
  sheet.getRange(2, 1, 1, headerRow2.length).setValues([headerRow2]);
  
  sheet.getRange(1, 1, 1, headerRow1.length)
    .setFontWeight('bold')
    .setFontSize(10)
    .setHorizontalAlignment('center');
  
  // 色分け
  sheet.getRange(1, 3, 1, 3).setBackground('#e8f0fe');
  sheet.getRange(1, 6, 1, 3).setBackground('#fce8e6');
  sheet.getRange(1, 9, 1, 4).setBackground('#e6f4ea');
  sheet.getRange(1, 13, 1, 5).setBackground('#fff3e0');
  sheet.getRange(1, 18, 1, 2).setBackground('#fef7e0');
  sheet.getRange(1, 20, 1, 4).setBackground('#f3e8fd');
  
  sheet.getRange(2, 1, 1, headerRow2.length)
    .setFontWeight('bold')
    .setBackground('#4285f4')
    .setFontColor('white')
    .setHorizontalAlignment('center')
    .setBorder(true, true, true, true, true, true);
  
  // 列幅設定
  sheet.setColumnWidth(COLUMNS.VIDEO_ID, 130);
  sheet.setColumnWidth(COLUMNS.VIDEO_URL, 280);
  sheet.setColumnWidth(COLUMNS.CURRENT_TITLE, 280);
  sheet.setColumnWidth(COLUMNS.CURRENT_DESCRIPTION, 300);
  sheet.setColumnWidth(COLUMNS.CURRENT_TAGS, 200);
  sheet.setColumnWidth(COLUMNS.NEW_TITLE, 280);
  sheet.setColumnWidth(COLUMNS.NEW_DESCRIPTION, 300);
  sheet.setColumnWidth(COLUMNS.NEW_TAGS, 200);
  sheet.setColumnWidth(COLUMNS.VIEW_COUNT, 90);
  sheet.setColumnWidth(COLUMNS.LIKE_COUNT, 80);
  sheet.setColumnWidth(COLUMNS.COMMENT_COUNT, 90);
  sheet.setColumnWidth(COLUMNS.PUBLISHED_AT, 140);
  sheet.setColumnWidth(COLUMNS.AVG_VIEW_DURATION, 110);
  sheet.setColumnWidth(COLUMNS.AVG_VIEW_PERCENTAGE, 100);
  sheet.setColumnWidth(COLUMNS.SUBSCRIBERS_GAINED, 100);
  sheet.setColumnWidth(COLUMNS.SUBSCRIBERS_LOST, 100);
  sheet.setColumnWidth(COLUMNS.EST_MINUTES_WATCHED, 110);
  sheet.setColumnWidth(COLUMNS.THUMBNAIL_URL, 180);
  sheet.setColumnWidth(COLUMNS.THUMBNAIL_IMAGE, 160);
  sheet.setColumnWidth(COLUMNS.CATEGORY_ID, 90);
  sheet.setColumnWidth(COLUMNS.UPDATE_FLAG, 90);
  sheet.setColumnWidth(COLUMNS.LAST_FETCHED, 150);
  sheet.setColumnWidth(COLUMNS.IMPROVEMENT_COMMENT, 300);
  
  // 背景色（データ行）
  sheet.getRange(3, COLUMNS.CURRENT_TITLE, 997, 3).setBackground('#f8f9fa');
  sheet.getRange(3, COLUMNS.NEW_TITLE, 997, 3).setBackground('#ffffff');
  sheet.setRowHeightsForced(3, 1, 90);
  
  // 条件付き書式
  var rules = [];
  var flagRange = sheet.getRange('U3:U1000');
  
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('○')
    .setBackground('#fff2cc')
    .setRanges([flagRange])
    .build());
  
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('完了')
    .setBackground('#d9ead3')
    .setRanges([flagRange])
    .build());
  
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('エラー')
    .setBackground('#f4cccc')
    .setRanges([flagRange])
    .build());
  
  sheet.setConditionalFormatRules(rules);
  
  // フィルター
  if (sheet.getFilter()) {
    sheet.getFilter().remove();
  }
  sheet.getRange(2, 1, sheet.getMaxRows() - 1, headerRow2.length).createFilter();
  sheet.setFrozenRows(2);
  
  // 更新フラグのデータ検証
  var flagValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(['', '○', '完了', 'エラー'], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange('U3:U1000').setDataValidation(flagValidation);
  
  // 保護
  var protection = sheet.getRange(3, COLUMNS.CURRENT_TITLE, 997, 3).protect();
  protection.setDescription('現在の値（自動取得）');
  protection.setWarningOnly(true);
  
  // セル結合（グループヘッダー）
  sheet.getRange(1, 3, 1, 3).merge();
  sheet.getRange(1, 6, 1, 3).merge();
  sheet.getRange(1, 9, 1, 4).merge();
  sheet.getRange(1, 13, 1, 5).merge();
  sheet.getRange(1, 18, 1, 2).merge();
  sheet.getRange(1, 20, 1, 4).merge();
  
  Logger.log('YouTube管理シート作成完了');
}

function createConfigSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = CONFIG.CONFIG_SHEET_NAME;

  Logger.log('createConfigSheet: シート名 = ' + sheetName);

  var sheet = ss.getSheetByName(sheetName);

  if (sheet) {
    sheet.clear();
    Logger.log('既存の設定シートをクリアしました');
  } else {
    // 明示的に名前を指定してシートを作成
    try {
      sheet = ss.insertSheet(sheetName);
      Logger.log('シート作成成功: ' + sheet.getName());
    } catch (e) {
      Logger.log('insertSheet エラー: ' + e.message);
      sheet = ss.insertSheet();
      sheet.setName(sheetName);
      Logger.log('setName後: ' + sheet.getName());
    }
  }
  
  // ========== 左側：チャンネル設定 ==========
  // 行1: タイトル
  sheet.getRange(1, 1).setValue('【監視チャンネル設定】');
  sheet.getRange(1, 1, 1, 6).merge()
    .setBackground('#4285f4')
    .setFontColor('white')
    .setFontWeight('bold')
    .setFontSize(11);
  
  // 行2: ヘッダー
  var channelHeaders = ['チャンネルID', 'チャンネル名', 'プレイリストID', '有効', '最終取得日時', 'メモ'];
  sheet.getRange(2, 1, 1, channelHeaders.length).setValues([channelHeaders]);
  sheet.getRange(2, 1, 1, channelHeaders.length)
    .setFontWeight('bold')
    .setBackground('#e8f0fe')
    .setBorder(true, true, true, true, true, true);
  
  // チェックボックス
  var checkboxValidation = SpreadsheetApp.newDataValidation().requireCheckbox().build();
  sheet.getRange('D3:D100').setDataValidation(checkboxValidation);
  
  // 列幅（左側）
  sheet.setColumnWidth(1, 220);
  sheet.setColumnWidth(2, 200);
  sheet.setColumnWidth(3, 220);
  sheet.setColumnWidth(4, 60);
  sheet.setColumnWidth(5, 150);
  sheet.setColumnWidth(6, 200);
  
  // ========== 右側：一般設定 ==========
  var col = GENERAL_SETTINGS_COL;  // H列 = 8
  
  // 行1: タイトル
  sheet.getRange(1, col).setValue('【一般設定】');
  sheet.getRange(1, col, 1, 3).merge()
    .setBackground('#34a853')
    .setFontColor('white')
    .setFontWeight('bold')
    .setFontSize(11);
  
  // 行2: ヘッダー
  sheet.getRange(2, col, 1, 3).setValues([['設定項目', '値', '説明']]);
  sheet.getRange(2, col, 1, 3)
    .setFontWeight('bold')
    .setBackground('#e6f4ea')
    .setBorder(true, true, true, true, true, true);
  
  // 行3以降: 設定データ（設定項目, 値, 説明の順）
  var settingsData = [
    ['定期実行間隔（時間）', CONFIG.AUTO_FETCH_INTERVAL_HOURS, '新規動画チェックの間隔'],
    ['サムネイルサイズ', CONFIG.THUMBNAIL_SIZE, 'default/medium/high/maxres'],
    ['新規動画取得数', 100, 'チャンネルから取得する最新動画数'],
    ['Analytics取得期間（日）', CONFIG.ANALYTICS_DAYS, '過去何日分のAnalyticsを取得'],
    ['AnalyticsチャンネルID', '', 'ブランドアカウントの場合は入力（UCxxxx形式）'],
    ['最終実行日時', '', '自動更新']
  ];
  sheet.getRange(3, col, settingsData.length, 3).setValues(settingsData);
  
  // 列幅（右側）
  sheet.setColumnWidth(col, 200);
  sheet.setColumnWidth(col + 1, 220);
  sheet.setColumnWidth(col + 2, 300);
  
  // 行設定
  sheet.setRowHeight(1, 30);
  sheet.setFrozenRows(2);
  
  // ========== 使い方セクション ==========
  var instructionRow = 12;
  sheet.getRange(instructionRow, 1).setValue('【使い方】');
  sheet.getRange(instructionRow, 1, 1, 6).merge()
    .setFontWeight('bold')
    .setBackground('#f0f0f0');
  
  var instructions = [
    ['1. 「設定シートにチャンネルを追加」からチャンネルを登録'],
    ['2. 「有効」列にチェックを入れたチャンネルが定期実行の対象'],
    ['3. チャンネルIDは手入力も可能（UCxxxx形式）'],
    ['4. 複数チャンネルを登録して一括管理可能'],
    [''],
    ['【Analytics APIについて】'],
    ['・YouTube Analytics APIは自分のチャンネルの動画のみ取得可能'],
    ['・他人のチャンネルの動画はAnalytics情報を取得できません'],
    ['・サービスの追加が必要（メニュー→サービス→YouTube Analytics API）'],
    [''],
    ['【ブランドアカウントの場合】'],
    ['・右側の「AnalyticsチャンネルID」にブランドアカウントのチャンネルIDを入力'],
    ['・チャンネルIDはYouTube Studioまたはチャンネルページから確認可能']
  ];
  sheet.getRange(instructionRow + 1, 1, instructions.length, 1).setValues(instructions);
  sheet.getRange(instructionRow + 1, 1, instructions.length, 1).setFontColor('#666666');
  
  Logger.log('設定シート作成完了');
}

function openConfigSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.CONFIG_SHEET_NAME);
  
  if (!sheet) {
    createConfigSheet();
    sheet = ss.getSheetByName(CONFIG.CONFIG_SHEET_NAME);
  }
  
  ss.setActiveSheet(sheet);
}

// ==================== 設定シートの読み書き ====================

function getEnabledChannels() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.CONFIG_SHEET_NAME);
  
  if (!sheet) return [];
  
  var data = sheet.getDataRange().getValues();
  var channels = [];
  
  for (var i = 2; i < data.length; i++) {
    var channelId = data[i][CONFIG_COLUMNS.CHANNEL_ID - 1];
    var enabled = data[i][CONFIG_COLUMNS.ENABLED - 1];
    
    if (!channelId || String(channelId).indexOf('【') === 0) continue;
    
    if (enabled === true) {
      channels.push({
        row: i + 1,
        channelId: channelId,
        channelName: data[i][CONFIG_COLUMNS.CHANNEL_NAME - 1],
        playlistId: data[i][CONFIG_COLUMNS.PLAYLIST_ID - 1],
        note: data[i][CONFIG_COLUMNS.NOTE - 1]
      });
    }
  }
  
  return channels;
}

function getGeneralSettings() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.CONFIG_SHEET_NAME);
  
  // デフォルト値
  var settings = {
    interval: CONFIG.AUTO_FETCH_INTERVAL_HOURS || 6,
    thumbnailSize: CONFIG.THUMBNAIL_SIZE || 'medium',
    maxVideos: 100,
    analyticsDays: CONFIG.ANALYTICS_DAYS || 28,
    analyticsChannelId: ''
  };
  
  if (!sheet) {
    Logger.log('設定シートがないため、デフォルト値を使用');
    return settings;
  }
  
  var data = sheet.getDataRange().getValues();
  var settingsColIndex = GENERAL_SETTINGS_COL - 1;  // 0始まりインデックス（H列=7）
  
  // 行3から読み取り開始（インデックス2）
  for (var i = 2; i < data.length; i++) {
    var key = String(data[i][settingsColIndex] || '').trim();
    var value = data[i][settingsColIndex + 1];
    
    if (!key) continue;
    
    if (key === '定期実行間隔（時間）' && value !== '' && value !== null) {
      settings.interval = parseInt(value) || settings.interval;
    } else if (key === 'サムネイルサイズ' && value) {
      settings.thumbnailSize = String(value);
    } else if (key === '新規動画取得数' && value !== '' && value !== null) {
      settings.maxVideos = parseInt(value) || settings.maxVideos;
    } else if (key === 'Analytics取得期間（日）' && value !== '' && value !== null) {
      settings.analyticsDays = parseInt(value) || settings.analyticsDays;
    } else if (key === 'AnalyticsチャンネルID' && value) {
      settings.analyticsChannelId = String(value).trim();
    }
  }
  
  Logger.log('設定読み込み: ' + JSON.stringify(settings));
  
  return settings;
}

function updateLastRunTime() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.CONFIG_SHEET_NAME);
  if (!sheet) return;
  
  var data = sheet.getDataRange().getValues();
  var settingsColIndex = GENERAL_SETTINGS_COL - 1;
  
  for (var i = 0; i < data.length; i++) {
    if (data[i][settingsColIndex] === '最終実行日時') {
      sheet.getRange(i + 1, GENERAL_SETTINGS_COL + 1).setValue(new Date());
      break;
    }
  }
}

function updateChannelLastFetched(row) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.CONFIG_SHEET_NAME);
  if (!sheet) return;
  sheet.getRange(row, CONFIG_COLUMNS.LAST_FETCHED).setValue(new Date());
}

function addChannelToConfig() {
  var ui = SpreadsheetApp.getUi();
  
  var response = ui.prompt(
    'チャンネルを追加',
    '追加するチャンネルを入力してください。\n\n' +
    '・空欄: 自分のチャンネル\n' +
    '・チャンネルID: UCxxxx 形式\n' +
    '・ハンドル: @username 形式',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() !== ui.Button.OK) return;
  
  var input = response.getResponseText().trim();
  
  try {
    var channelResponse;
    
    if (!input) {
      channelResponse = YouTube.Channels.list('snippet,contentDetails', { mine: true });
      if (!channelResponse.items || channelResponse.items.length === 0) {
        showAlert('自分のチャンネルが見つかりません');
        return;
      }
    } else if (input.indexOf('@') === 0) {
      channelResponse = YouTube.Channels.list('snippet,contentDetails', { forHandle: input.substring(1) });
    } else {
      channelResponse = YouTube.Channels.list('snippet,contentDetails', { id: input });
    }
    
    if (!channelResponse.items || channelResponse.items.length === 0) {
      showAlert('チャンネルが見つかりません: ' + input);
      return;
    }
    
    var channel = channelResponse.items[0];
    var channelId = channel.id;
    var channelName = channel.snippet.title;
    var playlistId = channel.contentDetails.relatedPlaylists.uploads;
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(CONFIG.CONFIG_SHEET_NAME);
    
    if (!sheet) {
      createConfigSheet();
      sheet = ss.getSheetByName(CONFIG.CONFIG_SHEET_NAME);
    }
    
    var data = sheet.getDataRange().getValues();
    for (var i = 2; i < data.length; i++) {
      if (data[i][0] === channelId) {
        showAlert('このチャンネルは既に登録されています。\n\nチャンネル名: ' + channelName);
        return;
      }
    }
    
    // 空の行を探す（行3から）
    var insertRow = 3;
    for (var i = 2; i < Math.min(data.length, 100); i++) {
      if (!data[i][0] || String(data[i][0]).indexOf('【') === 0) {
        insertRow = i + 1;
        break;
      }
      insertRow = i + 2;
    }
    
    var noteResponse = ui.prompt(
      'メモ（任意）',
      'このチャンネルのメモを入力してください（空欄可）:',
      ui.ButtonSet.OK_CANCEL
    );
    
    var note = '';
    if (noteResponse.getSelectedButton() === ui.Button.OK) {
      note = noteResponse.getResponseText();
    }
    
    var rowData = [channelId, channelName, playlistId, true, '', note];
    sheet.getRange(insertRow, 1, 1, rowData.length).setValues([rowData]);
    
    var checkboxValidation = SpreadsheetApp.newDataValidation().requireCheckbox().build();
    sheet.getRange(insertRow, CONFIG_COLUMNS.ENABLED).setDataValidation(checkboxValidation);
    sheet.getRange(insertRow, CONFIG_COLUMNS.ENABLED).setValue(true);
    
    showAlert('チャンネルを追加しました！\n\n' +
      'チャンネル名: ' + channelName + '\n' +
      'チャンネルID: ' + channelId + '\n\n' +
      '設定シートで「有効」列を確認し、定期実行を開始してください。\n\n' +
      '【ブランドアカウントの場合】\n' +
      '「AnalyticsチャンネルID」にも同じIDを設定してください。');
    
    ss.setActiveSheet(sheet);
    
  } catch (e) {
    showAlert('エラー: ' + e.message);
  }
}

// ==================== YouTube Analytics API ====================

function getMyChannelId() {
  try {
    var response = YouTube.Channels.list('id', { mine: true });
    if (response.items && response.items.length > 0) {
      return response.items[0].id;
    }
  } catch (e) {
    Logger.log('チャンネルID取得エラー: ' + e.message);
  }
  return null;
}

function getAnalyticsChannelId() {
  var settings = getGeneralSettings();
  
  if (settings.analyticsChannelId) {
    Logger.log('設定シートからAnalyticsチャンネルID取得: ' + settings.analyticsChannelId);
    return settings.analyticsChannelId;
  }
  
  var myChannelId = getMyChannelId();
  Logger.log('デフォルトチャンネルID使用: ' + myChannelId);
  return myChannelId;
}

function fetchAnalyticsInfo() {
  var sheet = getSheet();
  var data = sheet.getDataRange().getValues();
  var startRow = 3;
  
  if (data.length < startRow) {
    showAlert('動画IDを3行目以降のA列に入力してください');
    return;
  }
  
  var analyticsChannelId = getAnalyticsChannelId();
  if (!analyticsChannelId) {
    showAlert('チャンネルIDを取得できません。\n\n' +
      '【対処法】\n' +
      '・YouTubeチャンネルを持つアカウントで認証してください\n' +
      '・ブランドアカウントの場合は、設定シートの\n' +
      '  「AnalyticsチャンネルID」に入力してください');
    return;
  }
  
  var videoIds = [];
  for (var i = startRow - 1; i < data.length; i++) {
    var cellValue = data[i][0];
    if (cellValue) {
      var extractedId = extractVideoId(String(cellValue));
      if (extractedId) {
        videoIds.push({
          row: i + 1,
          id: extractedId
        });
      }
    }
  }
  
  if (videoIds.length === 0) {
    showAlert('動画IDが見つかりません');
    return;
  }
  
  var settings = getGeneralSettings();
  var successCount = 0;
  var errorCount = 0;
  var notOwnedCount = 0;
  
  var endDate = new Date();
  var startDate = new Date();
  startDate.setDate(startDate.getDate() - settings.analyticsDays);
  
  var startDateStr = Utilities.formatDate(startDate, 'Asia/Tokyo', 'yyyy-MM-dd');
  var endDateStr = Utilities.formatDate(endDate, 'Asia/Tokyo', 'yyyy-MM-dd');
  
  Logger.log('Analytics取得開始');
  Logger.log('チャンネルID: ' + analyticsChannelId);
  Logger.log('期間: ' + startDateStr + ' ～ ' + endDateStr);
  Logger.log('対象動画数: ' + videoIds.length);
  
  var batchSize = 50;
  
  for (var i = 0; i < videoIds.length; i += batchSize) {
    var batch = videoIds.slice(i, i + batchSize);
    var ids = [];
    for (var j = 0; j < batch.length; j++) {
      if (batch[j].id) {
        ids.push(batch[j].id);
      }
    }
    
    if (ids.length === 0) continue;
    
    try {
      var analyticsData = fetchAnalyticsForVideos(analyticsChannelId, ids, startDateStr, endDateStr);
      
      for (var k = 0; k < batch.length; k++) {
        var videoId = batch[k].id;
        var row = batch[k].row;
        
        if (analyticsData[videoId]) {
          var stats = analyticsData[videoId];
          
          sheet.getRange(row, COLUMNS.AVG_VIEW_DURATION).setValue(formatDuration(stats.avgViewDuration || 0));
          sheet.getRange(row, COLUMNS.AVG_VIEW_PERCENTAGE).setValue(stats.avgViewPercentage ? stats.avgViewPercentage.toFixed(1) : 0);
          sheet.getRange(row, COLUMNS.SUBSCRIBERS_GAINED).setValue(stats.subscribersGained || 0);
          sheet.getRange(row, COLUMNS.SUBSCRIBERS_LOST).setValue(stats.subscribersLost || 0);
          sheet.getRange(row, COLUMNS.EST_MINUTES_WATCHED).setValue(Math.round(stats.estimatedMinutesWatched || 0));
          
          successCount++;
        } else {
          notOwnedCount++;
        }
      }
      
      SpreadsheetApp.flush();
      
    } catch (e) {
      Logger.log('Analytics取得エラー: ' + e.message);
      errorCount++;
    }
    
    Utilities.sleep(200);
  }
  
  var lastRow = sheet.getLastRow();
  if (lastRow >= startRow) {
    sheet.getRange(startRow, COLUMNS.SUBSCRIBERS_GAINED, lastRow - startRow + 1, 1).setNumberFormat('#,##0');
    sheet.getRange(startRow, COLUMNS.SUBSCRIBERS_LOST, lastRow - startRow + 1, 1).setNumberFormat('#,##0');
    sheet.getRange(startRow, COLUMNS.EST_MINUTES_WATCHED, lastRow - startRow + 1, 1).setNumberFormat('#,##0');
  }
  
  var message = 'Analytics情報取得完了\n\n' +
    '取得成功: ' + successCount + '件\n' +
    '取得期間: 過去' + settings.analyticsDays + '日間\n' +
    '使用チャンネルID: ' + analyticsChannelId;
  
  if (notOwnedCount > 0) {
    message += '\n\n※ ' + notOwnedCount + '件は指定チャンネルの動画ではないため\nAnalytics情報を取得できませんでした。';
  }
  
  if (errorCount > 0) {
    message += '\nエラー: ' + errorCount + '件';
  }
  
  showAlert(message);
}

function fetchAnalyticsForVideos(channelId, videoIds, startDate, endDate) {
  var result = {};
  
  if (!videoIds || !Array.isArray(videoIds) || videoIds.length === 0) {
    Logger.log('fetchAnalyticsForVideos: videoIdsが空または無効です');
    return result;
  }
  
  if (!channelId) {
    Logger.log('fetchAnalyticsForVideos: channelIdが空です');
    return result;
  }
  
  Logger.log('fetchAnalyticsForVideos: チャンネルID=' + channelId + ', 動画数=' + videoIds.length);
  
  try {
    var response = YouTubeAnalytics.Reports.query({
      ids: 'channel==' + channelId,
      startDate: startDate,
      endDate: endDate,
      metrics: 'views,estimatedMinutesWatched,averageViewDuration,averageViewPercentage,subscribersGained,subscribersLost,likes,shares',
      dimensions: 'video',
      filters: 'video==' + videoIds.join(','),
      maxResults: videoIds.length
    });
    
    Logger.log('Analytics API応答: ' + (response.rows ? response.rows.length : 0) + '件');
    
    if (response.rows) {
      for (var i = 0; i < response.rows.length; i++) {
        var row = response.rows[i];
        var videoId = row[0];
        
        result[videoId] = {
          views: row[1] || 0,
          estimatedMinutesWatched: row[2] || 0,
          avgViewDuration: row[3] || 0,
          avgViewPercentage: row[4] || 0,
          subscribersGained: row[5] || 0,
          subscribersLost: row[6] || 0,
          likes: row[7] || 0,
          shares: row[8] || 0
        };
      }
    }
    
  } catch (e) {
    Logger.log('Analytics API エラー: ' + e.message);
    
    if (e.message.indexOf('YouTubeAnalytics') !== -1 || e.message.indexOf('not defined') !== -1) {
      throw new Error('YouTube Analytics APIサービスが追加されていません。\n' +
        'スクリプトエディタで「サービス」→「YouTube Analytics API」を追加してください。');
    }
    throw e;
  }
  
  return result;
}

function formatDuration(seconds) {
  if (!seconds || seconds <= 0) return '0:00';
  
  var mins = Math.floor(seconds / 60);
  var secs = Math.floor(seconds % 60);
  
  return mins + ':' + (secs < 10 ? '0' : '') + secs;
}

function fetchAllInfo() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert(
    '全情報取得',
    '動画情報とAnalytics情報を両方取得します。\n\n' +
    '※ Analytics情報は自分のチャンネルの動画のみ取得可能です。\n' +
    '※ ブランドアカウントの場合は設定シートにチャンネルIDを設定してください。\n\n' +
    '続行しますか？',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) return;
  
  fetchVideoInfoSilent();
  
  try {
    fetchAnalyticsInfoSilent();
    showAlert('全情報の取得が完了しました！');
  } catch (e) {
    showAlert('動画情報は取得しました。\n\nAnalytics情報の取得中にエラー:\n' + e.message);
  }
}

function fetchVideoInfoSilent() {
  var sheet = getSheet();
  var data = sheet.getDataRange().getValues();
  var startRow = 3;
  
  if (data.length < startRow) return 0;
  
  var videoIds = [];
  for (var i = startRow - 1; i < data.length; i++) {
    var cellValue = data[i][0];
    if (cellValue) {
      var extractedId = extractVideoId(String(cellValue));
      if (extractedId) {
        videoIds.push({
          row: i + 1,
          id: extractedId
        });
      }
    }
  }
  
  if (videoIds.length === 0) return 0;
  
  var batchSize = 50;
  var successCount = 0;
  
  for (var i = 0; i < videoIds.length; i += batchSize) {
    var batch = videoIds.slice(i, i + batchSize);
    var ids = [];
    for (var j = 0; j < batch.length; j++) {
      ids.push(batch[j].id);
    }
    var idsStr = ids.join(',');
    
    try {
      var response = YouTube.Videos.list('snippet,statistics', { id: idsStr });
      
      if (response.items && response.items.length > 0) {
        for (var k = 0; k < response.items.length; k++) {
          var video = response.items[k];
          var target = null;
          
          for (var m = 0; m < batch.length; m++) {
            if (batch[m].id === video.id) {
              target = batch[m];
              break;
            }
          }
          
          if (target) {
            var thumbnailUrl = getThumbnailUrl(video.snippet.thumbnails);
            var tags = video.snippet.tags ? video.snippet.tags.join(', ') : '';
            var videoUrl = getVideoUrl(video.id);
            var categoryId = video.snippet.categoryId || '';
            
            sheet.getRange(target.row, COLUMNS.VIDEO_ID).setValue(video.id);
            sheet.getRange(target.row, COLUMNS.VIDEO_URL).setValue(videoUrl);
            sheet.getRange(target.row, COLUMNS.CURRENT_TITLE).setValue(video.snippet.title);
            sheet.getRange(target.row, COLUMNS.CURRENT_DESCRIPTION).setValue(video.snippet.description);
            sheet.getRange(target.row, COLUMNS.CURRENT_TAGS).setValue(tags);
            sheet.getRange(target.row, COLUMNS.VIEW_COUNT).setValue(parseInt(video.statistics.viewCount) || 0);
            sheet.getRange(target.row, COLUMNS.LIKE_COUNT).setValue(parseInt(video.statistics.likeCount) || 0);
            sheet.getRange(target.row, COLUMNS.COMMENT_COUNT).setValue(parseInt(video.statistics.commentCount) || 0);
            sheet.getRange(target.row, COLUMNS.PUBLISHED_AT).setValue(formatDate(video.snippet.publishedAt));
            sheet.getRange(target.row, COLUMNS.THUMBNAIL_URL).setValue(thumbnailUrl);
            sheet.getRange(target.row, COLUMNS.THUMBNAIL_IMAGE).setFormula('=IMAGE("' + thumbnailUrl + '")');
            sheet.getRange(target.row, COLUMNS.CATEGORY_ID).setValue(categoryId);
            sheet.getRange(target.row, COLUMNS.LAST_FETCHED).setValue(Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss'));
            
            successCount++;
          }
        }
        
        SpreadsheetApp.flush();
      }
    } catch (e) {
      Logger.log('バッチ処理エラー: ' + e.message);
    }
    
    Utilities.sleep(200);
  }
  
  var lastRow = sheet.getLastRow();
  if (lastRow >= startRow) {
    sheet.getRange(startRow, COLUMNS.VIEW_COUNT, lastRow - startRow + 1, 3).setNumberFormat('#,##0');
    sheet.setRowHeightsForced(startRow, lastRow - startRow + 1, 90);
  }
  
  return successCount;
}

function fetchAnalyticsInfoSilent() {
  var sheet = getSheet();
  var data = sheet.getDataRange().getValues();
  var startRow = 3;
  
  if (data.length < startRow) return 0;
  
  var analyticsChannelId = getAnalyticsChannelId();
  if (!analyticsChannelId) {
    Logger.log('Analytics用チャンネルIDが取得できません');
    return 0;
  }
  
  var videoIds = [];
  for (var i = startRow - 1; i < data.length; i++) {
    var cellValue = data[i][0];
    if (cellValue) {
      var extractedId = extractVideoId(String(cellValue));
      if (extractedId) {
        videoIds.push({
          row: i + 1,
          id: extractedId
        });
      }
    }
  }
  
  if (videoIds.length === 0) return 0;
  
  var settings = getGeneralSettings();
  var successCount = 0;
  
  var endDate = new Date();
  var startDate = new Date();
  startDate.setDate(startDate.getDate() - settings.analyticsDays);
  
  var startDateStr = Utilities.formatDate(startDate, 'Asia/Tokyo', 'yyyy-MM-dd');
  var endDateStr = Utilities.formatDate(endDate, 'Asia/Tokyo', 'yyyy-MM-dd');
  
  var batchSize = 50;
  
  for (var i = 0; i < videoIds.length; i += batchSize) {
    var batch = videoIds.slice(i, i + batchSize);
    var ids = [];
    for (var j = 0; j < batch.length; j++) {
      if (batch[j].id) {
        ids.push(batch[j].id);
      }
    }
    
    if (ids.length === 0) continue;
    
    try {
      var analyticsData = fetchAnalyticsForVideos(analyticsChannelId, ids, startDateStr, endDateStr);
      
      for (var k = 0; k < batch.length; k++) {
        var videoId = batch[k].id;
        var row = batch[k].row;
        
        if (analyticsData[videoId]) {
          var stats = analyticsData[videoId];
          
          sheet.getRange(row, COLUMNS.AVG_VIEW_DURATION).setValue(formatDuration(stats.avgViewDuration || 0));
          sheet.getRange(row, COLUMNS.AVG_VIEW_PERCENTAGE).setValue(stats.avgViewPercentage ? stats.avgViewPercentage.toFixed(1) : 0);
          sheet.getRange(row, COLUMNS.SUBSCRIBERS_GAINED).setValue(stats.subscribersGained || 0);
          sheet.getRange(row, COLUMNS.SUBSCRIBERS_LOST).setValue(stats.subscribersLost || 0);
          sheet.getRange(row, COLUMNS.EST_MINUTES_WATCHED).setValue(Math.round(stats.estimatedMinutesWatched || 0));
          
          successCount++;
        }
      }
      
      SpreadsheetApp.flush();
      
    } catch (e) {
      Logger.log('Analytics取得エラー: ' + e.message);
    }
    
    Utilities.sleep(200);
  }
  
  var lastRow = sheet.getLastRow();
  if (lastRow >= startRow) {
    sheet.getRange(startRow, COLUMNS.SUBSCRIBERS_GAINED, lastRow - startRow + 1, 1).setNumberFormat('#,##0');
    sheet.getRange(startRow, COLUMNS.EST_MINUTES_WATCHED, lastRow - startRow + 1, 1).setNumberFormat('#,##0');
  }
  
  return successCount;
}

// ==================== 定期実行（トリガー管理） ====================

function startAutoFetch() {
  var channels = getEnabledChannels();
  
  if (channels.length === 0) {
    showAlert('有効なチャンネルが設定されていません。\n\n' +
      '1. 「設定シートにチャンネルを追加」でチャンネルを登録\n' +
      '2. 設定シートの「有効」列にチェックを入れてください');
    return;
  }
  
  stopAutoFetchSilent();
  
  var settings = getGeneralSettings();
  
  ScriptApp.newTrigger('autoFetchAll')
    .timeBased()
    .everyHours(settings.interval)
    .create();
  
  var channelNames = [];
  for (var i = 0; i < channels.length; i++) {
    channelNames.push('・' + channels[i].channelName);
  }
  
  showAlert('定期実行を開始しました\n\n' +
    '【監視チャンネル】\n' + channelNames.join('\n') + '\n\n' +
    '【実行間隔】\n' + settings.interval + '時間ごと\n\n' +
    '【実行内容】\n' +
    '1. 各チャンネルから新規動画IDを取得\n' +
    '2. 全動画の詳細情報・統計を更新\n' +
    '3. Analytics情報を更新');
}

function stopAutoFetch() {
  var count = stopAutoFetchSilent();
  
  if (count > 0) {
    showAlert('定期実行を停止しました（' + count + '件のトリガーを削除）');
  } else {
    showAlert('実行中の定期実行はありません');
  }
}

function stopAutoFetchSilent() {
  var triggers = ScriptApp.getProjectTriggers();
  var deletedCount = 0;
  
  for (var i = 0; i < triggers.length; i++) {
    var funcName = triggers[i].getHandlerFunction();
    if (funcName === 'autoFetchAll' || funcName === 'autoFetchVideoInfo') {
      ScriptApp.deleteTrigger(triggers[i]);
      deletedCount++;
    }
  }
  
  return deletedCount;
}

function checkTriggerStatus() {
  var triggers = ScriptApp.getProjectTriggers();
  var autoTriggers = [];
  
  for (var i = 0; i < triggers.length; i++) {
    var funcName = triggers[i].getHandlerFunction();
    if (funcName === 'autoFetchAll' || funcName === 'autoFetchVideoInfo') {
      autoTriggers.push(triggers[i]);
    }
  }
  
  var channels = getEnabledChannels();
  var settings = getGeneralSettings();
  
  var channelInfo = '監視チャンネル: ';
  if (channels.length === 0) {
    channelInfo += 'なし';
  } else {
    var names = [];
    for (var i = 0; i < channels.length; i++) {
      names.push(channels[i].channelName);
    }
    channelInfo += names.join(', ');
  }
  
  var analyticsChannelInfo = 'AnalyticsチャンネルID: ';
  if (settings.analyticsChannelId) {
    analyticsChannelInfo += settings.analyticsChannelId;
  } else {
    analyticsChannelInfo += '（未設定・デフォルト使用）';
  }
  
  if (autoTriggers.length === 0) {
    showAlert('定期実行の状態\n\n' +
      'ステータス: 停止中\n' +
      channelInfo + '\n' +
      analyticsChannelInfo + '\n' +
      '設定間隔: ' + settings.interval + '時間\n' +
      'Analytics取得期間: ' + settings.analyticsDays + '日');
    return;
  }
  
  showAlert('定期実行の状態\n\n' +
    'ステータス: 実行中\n' +
    channelInfo + '\n' +
    analyticsChannelInfo + '\n' +
    '実行間隔: ' + settings.interval + '時間\n' +
    'Analytics取得期間: ' + settings.analyticsDays + '日');
}

function runAutoFetchManually() {
  var channels = getEnabledChannels();
  
  if (channels.length === 0) {
    showAlert('有効なチャンネルが設定されていません。\n\n' +
      '設定シートでチャンネルを追加し、「有効」にチェックを入れてください。');
    return;
  }
  
  var ui = SpreadsheetApp.getUi();
  
  var channelNames = [];
  for (var i = 0; i < channels.length; i++) {
    channelNames.push('・' + channels[i].channelName);
  }
  
  var response = ui.alert(
    '手動実行',
    '以下のチャンネルに対して処理を実行しますか？\n\n' +
    channelNames.join('\n') + '\n\n' +
    '1. 新規動画を取得\n' +
    '2. 全動画の情報を更新\n' +
    '3. Analytics情報を更新',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) return;
  
  autoFetchAll();
  
  showAlert('手動実行が完了しました');
}

function autoFetchAll() {
  try {
    Logger.log('=== 定期実行開始: ' + new Date() + ' ===');
    
    var channels = getEnabledChannels();
    var settings = getGeneralSettings();
    
    if (channels.length === 0) {
      Logger.log('有効なチャンネルがありません');
      return;
    }
    
    var totalNewCount = 0;
    
    for (var i = 0; i < channels.length; i++) {
      var channel = channels[i];
      Logger.log('チャンネル処理: ' + channel.channelName);
      
      var newCount = autoFetchNewVideosFromChannel(channel.playlistId, settings.maxVideos);
      totalNewCount += newCount;
      
      updateChannelLastFetched(channel.row);
      Utilities.sleep(500);
    }
    
    Logger.log('新規動画追加合計: ' + totalNewCount + '件');
    
    var updateCount = autoUpdateAllVideoInfo();
    Logger.log('情報更新: ' + updateCount + '件');
    
    try {
      var analyticsCount = fetchAnalyticsInfoSilent();
      Logger.log('Analytics更新: ' + analyticsCount + '件');
    } catch (e) {
      Logger.log('Analytics更新スキップ: ' + e.message);
    }
    
    updateLastRunTime();
    
    Logger.log('=== 定期実行完了 ===');
    
  } catch (e) {
    Logger.log('定期実行エラー: ' + e.message);
  }
}

function autoFetchNewVideosFromChannel(playlistId, maxResults) {
  if (!playlistId) {
    Logger.log('プレイリストIDが指定されていません');
    return 0;
  }
  
  var sheet = getSheet();
  var startRow = 3;
  
  var existingIds = {};
  var data = sheet.getDataRange().getValues();
  for (var i = startRow - 1; i < data.length; i++) {
    if (data[i][0]) {
      var id = extractVideoId(String(data[i][0]));
      existingIds[id] = true;
    }
  }
  
  var videoIds = [];
  var pageToken = '';
  
  while (videoIds.length < maxResults) {
    var params = {
      playlistId: playlistId,
      maxResults: Math.min(50, maxResults - videoIds.length)
    };
    
    if (pageToken) {
      params.pageToken = pageToken;
    }
    
    try {
      var response = YouTube.PlaylistItems.list('contentDetails', params);
      
      if (response.items) {
        for (var j = 0; j < response.items.length; j++) {
          videoIds.push(response.items[j].contentDetails.videoId);
        }
      }
      
      pageToken = response.nextPageToken;
      if (!pageToken) break;
    } catch (e) {
      Logger.log('プレイリスト取得エラー: ' + e.message);
      break;
    }
    
    Utilities.sleep(100);
  }
  
  var newVideoIds = [];
  for (var k = 0; k < videoIds.length; k++) {
    if (!existingIds[videoIds[k]]) {
      newVideoIds.push(videoIds[k]);
    }
  }
  
  if (newVideoIds.length === 0) {
    return 0;
  }
  
  var lastRow = sheet.getLastRow();
  var insertRow = Math.max(lastRow + 1, startRow);
  
  var videoIdData = [];
  for (var m = 0; m < newVideoIds.length; m++) {
    videoIdData.push([newVideoIds[m]]);
  }
  sheet.getRange(insertRow, 1, newVideoIds.length, 1).setValues(videoIdData);
  
  return newVideoIds.length;
}

function autoUpdateAllVideoInfo() {
  var sheet = getSheet();
  var data = sheet.getDataRange().getValues();
  var startRow = 3;
  
  if (data.length < startRow) return 0;
  
  var videoIds = [];
  for (var i = startRow - 1; i < data.length; i++) {
    var cellValue = data[i][0];
    if (cellValue) {
      var extractedId = extractVideoId(String(cellValue));
      if (extractedId) {
        videoIds.push({
          row: i + 1,
          id: extractedId
        });
      }
    }
  }
  
  if (videoIds.length === 0) return 0;
  
  var updateCount = 0;
  var batchSize = 50;
  
  for (var i = 0; i < videoIds.length; i += batchSize) {
    var batch = videoIds.slice(i, i + batchSize);
    var ids = [];
    for (var j = 0; j < batch.length; j++) {
      ids.push(batch[j].id);
    }
    var idsStr = ids.join(',');
    
    try {
      var response = YouTube.Videos.list('snippet,statistics', { id: idsStr });
      
      if (response.items && response.items.length > 0) {
        for (var k = 0; k < response.items.length; k++) {
          var video = response.items[k];
          var target = null;
          
          for (var m = 0; m < batch.length; m++) {
            if (batch[m].id === video.id) {
              target = batch[m];
              break;
            }
          }
          
          if (target) {
            var tags = video.snippet.tags ? video.snippet.tags.join(', ') : '';
            var thumbnailUrl = getThumbnailUrl(video.snippet.thumbnails);
            var videoUrl = getVideoUrl(video.id);
            var categoryId = video.snippet.categoryId || '';
            
            sheet.getRange(target.row, COLUMNS.VIDEO_ID).setValue(video.id);
            sheet.getRange(target.row, COLUMNS.VIDEO_URL).setValue(videoUrl);
            sheet.getRange(target.row, COLUMNS.CURRENT_TITLE).setValue(video.snippet.title);
            sheet.getRange(target.row, COLUMNS.CURRENT_DESCRIPTION).setValue(video.snippet.description);
            sheet.getRange(target.row, COLUMNS.CURRENT_TAGS).setValue(tags);
            sheet.getRange(target.row, COLUMNS.VIEW_COUNT).setValue(parseInt(video.statistics.viewCount) || 0);
            sheet.getRange(target.row, COLUMNS.LIKE_COUNT).setValue(parseInt(video.statistics.likeCount) || 0);
            sheet.getRange(target.row, COLUMNS.COMMENT_COUNT).setValue(parseInt(video.statistics.commentCount) || 0);
            sheet.getRange(target.row, COLUMNS.PUBLISHED_AT).setValue(formatDate(video.snippet.publishedAt));
            sheet.getRange(target.row, COLUMNS.THUMBNAIL_URL).setValue(thumbnailUrl);
            sheet.getRange(target.row, COLUMNS.THUMBNAIL_IMAGE).setFormula('=IMAGE("' + thumbnailUrl + '")');
            sheet.getRange(target.row, COLUMNS.CATEGORY_ID).setValue(categoryId);
            sheet.getRange(target.row, COLUMNS.LAST_FETCHED).setValue(Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss'));
            
            updateCount++;
          }
        }
        
        SpreadsheetApp.flush();
      }
    } catch (e) {
      Logger.log('バッチ処理エラー: ' + e.message);
    }
    
    Utilities.sleep(200);
  }
  
  var lastRow = sheet.getLastRow();
  if (lastRow >= startRow) {
    sheet.getRange(startRow, COLUMNS.VIEW_COUNT, lastRow - startRow + 1, 3).setNumberFormat('#,##0');
    sheet.setRowHeightsForced(startRow, lastRow - startRow + 1, 90);
  }
  
  return updateCount;
}

// ==================== チャンネル関連機能（手動） ====================

function fetchMyChannelVideos() {
  var ui = SpreadsheetApp.getUi();
  
  var response = ui.prompt(
    '自分のチャンネルから動画を取得',
    '取得する動画の最大件数を入力してください（最大500件）:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() !== ui.Button.OK) return;
  
  var maxResults = Math.min(parseInt(response.getResponseText()) || 50, 500);
  
  try {
    var channelResponse = YouTube.Channels.list('contentDetails,snippet', { mine: true });
    
    if (!channelResponse.items || channelResponse.items.length === 0) {
      showAlert('チャンネルが見つかりません。\nYouTubeチャンネルを持つアカウントで認証してください。');
      return;
    }
    
    var channel = channelResponse.items[0];
    var uploadsPlaylistId = channel.contentDetails.relatedPlaylists.uploads;
    var channelTitle = channel.snippet.title;
    
    ui.alert('チャンネル: ' + channelTitle + '\n\n動画を取得中...');
    
    fetchVideosFromPlaylist(uploadsPlaylistId, maxResults);
    
  } catch (e) {
    showAlert('エラー: ' + e.message);
  }
}

function fetchChannelVideosById() {
  var ui = SpreadsheetApp.getUi();
  
  var idResponse = ui.prompt(
    'チャンネルから動画を取得',
    'チャンネルID（UCxxxx）または ハンドル（@username）を入力:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (idResponse.getSelectedButton() !== ui.Button.OK) return;
  
  var channelInput = idResponse.getResponseText().trim();
  
  if (!channelInput) {
    showAlert('チャンネルIDまたはハンドルを入力してください');
    return;
  }
  
  var countResponse = ui.prompt(
    '取得件数',
    '取得する動画の最大件数を入力してください（最大500件）:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (countResponse.getSelectedButton() !== ui.Button.OK) return;
  
  var maxResults = Math.min(parseInt(countResponse.getResponseText()) || 50, 500);
  
  try {
    var channelId = channelInput;
    
    if (channelInput.indexOf('@') === 0) {
      var handleResponse = YouTube.Channels.list('id', { forHandle: channelInput.substring(1) });
      
      if (!handleResponse.items || handleResponse.items.length === 0) {
        showAlert('ハンドル「' + channelInput + '」が見つかりません');
        return;
      }
      
      channelId = handleResponse.items[0].id;
    }
    
    var channelResponse = YouTube.Channels.list('contentDetails,snippet', { id: channelId });
    
    if (!channelResponse.items || channelResponse.items.length === 0) {
      showAlert('チャンネルID「' + channelId + '」が見つかりません');
      return;
    }
    
    var channel = channelResponse.items[0];
    var uploadsPlaylistId = channel.contentDetails.relatedPlaylists.uploads;
    var channelTitle = channel.snippet.title;
    
    SpreadsheetApp.getUi().alert('チャンネル: ' + channelTitle + '\n\n動画を取得中...');
    
    fetchVideosFromPlaylist(uploadsPlaylistId, maxResults);
    
  } catch (e) {
    showAlert('エラー: ' + e.message);
  }
}

function fetchVideosFromPlaylist(playlistId, maxResults) {
  var sheet = getSheet();
  var startRow = 3;
  var ui = SpreadsheetApp.getUi();
  var lastRow = sheet.getLastRow();
  var insertRow = startRow;
  
  if (lastRow >= startRow) {
    var response = ui.alert(
      '既存データがあります',
      '既存データの下に追加しますか？\n\n' +
      '「はい」→ 下に追加\n' +
      '「いいえ」→ 既存データをクリアして追加',
      ui.ButtonSet.YES_NO_CANCEL
    );
    
    if (response === ui.Button.CANCEL) return;
    
    if (response === ui.Button.YES) {
      insertRow = lastRow + 1;
    } else {
      sheet.getRange(startRow, 1, lastRow - startRow + 1, sheet.getLastColumn()).clearContent();
    }
  }
  
  var videoIds = [];
  var pageToken = '';
  
  while (videoIds.length < maxResults) {
    var params = {
      playlistId: playlistId,
      maxResults: Math.min(50, maxResults - videoIds.length)
    };
    
    if (pageToken) {
      params.pageToken = pageToken;
    }
    
    var response = YouTube.PlaylistItems.list('contentDetails', params);
    
    if (response.items) {
      for (var i = 0; i < response.items.length; i++) {
        videoIds.push(response.items[i].contentDetails.videoId);
      }
    }
    
    pageToken = response.nextPageToken;
    if (!pageToken) break;
    
    Utilities.sleep(100);
  }
  
  if (videoIds.length === 0) {
    showAlert('動画が見つかりませんでした');
    return;
  }
  
  var videoIdData = [];
  for (var j = 0; j < videoIds.length; j++) {
    videoIdData.push([videoIds[j]]);
  }
  sheet.getRange(insertRow, 1, videoIds.length, 1).setValues(videoIdData);
  
  showAlert(videoIds.length + '件の動画IDを取得しました\n\n' +
    '「動画情報を取得」または「全情報を取得」を実行して詳細情報を取得してください。');
}

function showChannelInfo() {
  var ui = SpreadsheetApp.getUi();
  
  var response = ui.prompt(
    'チャンネル情報を表示',
    'チャンネルID（UCxxxx）、ハンドル（@username）、\nまたは空欄（自分のチャンネル）:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() !== ui.Button.OK) return;
  
  var input = response.getResponseText().trim();
  
  try {
    var channelResponse;
    
    if (!input) {
      channelResponse = YouTube.Channels.list('snippet,statistics,contentDetails', { mine: true });
    } else if (input.indexOf('@') === 0) {
      channelResponse = YouTube.Channels.list('snippet,statistics,contentDetails', { forHandle: input.substring(1) });
    } else {
      channelResponse = YouTube.Channels.list('snippet,statistics,contentDetails', { id: input });
    }
    
    if (!channelResponse.items || channelResponse.items.length === 0) {
      showAlert('チャンネルが見つかりません');
      return;
    }
    
    var channel = channelResponse.items[0];
    var stats = channel.statistics;
    var desc = channel.snippet.description || '';
    if (desc.length > 100) {
      desc = desc.substring(0, 100) + '...';
    }
    
    var info = 'チャンネル情報\n\n' +
      '【基本情報】\n' +
      'チャンネル名: ' + channel.snippet.title + '\n' +
      'チャンネルID: ' + channel.id + '\n' +
      '説明: ' + desc + '\n\n' +
      '【統計】\n' +
      '登録者数: ' + parseInt(stats.subscriberCount).toLocaleString() + '人\n' +
      '総再生回数: ' + parseInt(stats.viewCount).toLocaleString() + '回\n' +
      '動画数: ' + parseInt(stats.videoCount).toLocaleString() + '本\n\n' +
      '【アップロードプレイリストID】\n' +
      channel.contentDetails.relatedPlaylists.uploads + '\n\n' +
      '【ブランドアカウントでAnalyticsを取得する場合】\n' +
      '設定シートの「AnalyticsチャンネルID」に\n' +
      channel.id + ' を入力してください。';
    
    showAlert(info);
    
  } catch (e) {
    showAlert('エラー: ' + e.message);
  }
}

// ==================== 動画情報取得（手動） ====================

function fetchVideoInfo() {
  var sheet = getSheet();
  var data = sheet.getDataRange().getValues();
  var startRow = 3;
  
  if (data.length < startRow) {
    showAlert('動画IDを3行目以降のA列に入力してください');
    return;
  }
  
  var videoIds = [];
  for (var i = startRow - 1; i < data.length; i++) {
    var cellValue = data[i][0];
    if (cellValue) {
      var extractedId = extractVideoId(String(cellValue));
      if (extractedId) {
        videoIds.push({
          row: i + 1,
          id: extractedId
        });
      }
    }
  }
  
  if (videoIds.length === 0) {
    showAlert('動画IDが見つかりません');
    return;
  }
  
  var batchSize = 50;
  var successCount = 0;
  var errorCount = 0;
  
  for (var i = 0; i < videoIds.length; i += batchSize) {
    var batch = videoIds.slice(i, i + batchSize);
    var ids = [];
    for (var j = 0; j < batch.length; j++) {
      ids.push(batch[j].id);
    }
    var idsStr = ids.join(',');
    
    try {
      var response = YouTube.Videos.list('snippet,statistics', { id: idsStr });
      
      if (response.items && response.items.length > 0) {
        for (var k = 0; k < response.items.length; k++) {
          var video = response.items[k];
          var target = null;
          
          for (var m = 0; m < batch.length; m++) {
            if (batch[m].id === video.id) {
              target = batch[m];
              break;
            }
          }
          
          if (target) {
            var thumbnailUrl = getThumbnailUrl(video.snippet.thumbnails);
            var tags = video.snippet.tags ? video.snippet.tags.join(', ') : '';
            var videoUrl = getVideoUrl(video.id);
            var categoryId = video.snippet.categoryId || '';
            
            sheet.getRange(target.row, COLUMNS.VIDEO_ID).setValue(video.id);
            sheet.getRange(target.row, COLUMNS.VIDEO_URL).setValue(videoUrl);
            sheet.getRange(target.row, COLUMNS.CURRENT_TITLE).setValue(video.snippet.title);
            sheet.getRange(target.row, COLUMNS.CURRENT_DESCRIPTION).setValue(video.snippet.description);
            sheet.getRange(target.row, COLUMNS.CURRENT_TAGS).setValue(tags);
            sheet.getRange(target.row, COLUMNS.VIEW_COUNT).setValue(parseInt(video.statistics.viewCount) || 0);
            sheet.getRange(target.row, COLUMNS.LIKE_COUNT).setValue(parseInt(video.statistics.likeCount) || 0);
            sheet.getRange(target.row, COLUMNS.COMMENT_COUNT).setValue(parseInt(video.statistics.commentCount) || 0);
            sheet.getRange(target.row, COLUMNS.PUBLISHED_AT).setValue(formatDate(video.snippet.publishedAt));
            sheet.getRange(target.row, COLUMNS.THUMBNAIL_URL).setValue(thumbnailUrl);
            sheet.getRange(target.row, COLUMNS.THUMBNAIL_IMAGE).setFormula('=IMAGE("' + thumbnailUrl + '")');
            sheet.getRange(target.row, COLUMNS.CATEGORY_ID).setValue(categoryId);
            sheet.getRange(target.row, COLUMNS.LAST_FETCHED).setValue(Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss'));
            
            successCount++;
          }
        }
        
        SpreadsheetApp.flush();
      }
    } catch (e) {
      Logger.log('バッチ処理エラー: ' + e.message);
      errorCount++;
    }
    
    Utilities.sleep(200);
  }
  
  var lastRow = sheet.getLastRow();
  if (lastRow >= startRow) {
    sheet.getRange(startRow, COLUMNS.VIEW_COUNT, lastRow - startRow + 1, 3).setNumberFormat('#,##0');
    sheet.setRowHeightsForced(startRow, lastRow - startRow + 1, 90);
  }
  
  var message = successCount + '件の動画情報を取得しました';
  if (errorCount > 0) {
    message += '\n（エラー: ' + errorCount + '件）';
  }
  message += '\n\n【次のステップ】\n' +
    '・Analytics情報も取得する場合は「Analytics情報を取得」を実行\n' +
    '・修正する場合は、F-H列を編集し、U列に「○」を入力';
  
  showAlert(message);
}

// ==================== 動画情報更新 ====================

function updateVideoDetails() {
  var sheet = getSheet();
  var data = sheet.getDataRange().getValues();
  var startRow = 3;
  
  var updatedCount = 0;
  var errors = [];
  var processedRows = [];
  
  for (var i = startRow - 1; i < data.length; i++) {
    var row = i + 1;
    var videoIdRaw = data[i][COLUMNS.VIDEO_ID - 1];
    var videoId = extractVideoId(videoIdRaw ? String(videoIdRaw) : '');
    var updateFlag = data[i][COLUMNS.UPDATE_FLAG - 1];
    
    if (!updateFlag || (updateFlag !== '○' && updateFlag !== true && updateFlag !== 'TRUE')) {
      continue;
    }
    
    var currentTitle = data[i][COLUMNS.CURRENT_TITLE - 1];
    var currentDescription = data[i][COLUMNS.CURRENT_DESCRIPTION - 1] || '';
    var currentTags = data[i][COLUMNS.CURRENT_TAGS - 1] || '';
    
    var newTitle = data[i][COLUMNS.NEW_TITLE - 1] || currentTitle;
    var newDescription = data[i][COLUMNS.NEW_DESCRIPTION - 1] !== '' ? data[i][COLUMNS.NEW_DESCRIPTION - 1] : currentDescription;
    var newTags = data[i][COLUMNS.NEW_TAGS - 1] !== '' ? data[i][COLUMNS.NEW_TAGS - 1] : currentTags;
    var categoryId = data[i][COLUMNS.CATEGORY_ID - 1];
    
    if (!videoId) {
      errors.push('行' + row + ': 動画IDが空です');
      sheet.getRange(row, COLUMNS.UPDATE_FLAG).setValue('エラー');
      continue;
    }
    
    if (!newTitle) {
      errors.push('行' + row + ': タイトルが空です');
      sheet.getRange(row, COLUMNS.UPDATE_FLAG).setValue('エラー');
      continue;
    }
    
    try {
      var response = YouTube.Videos.list('snippet', { id: videoId });
      
      if (!response.items || response.items.length === 0) {
        errors.push('行' + row + ': 動画が見つかりません (' + videoId + ')');
        sheet.getRange(row, COLUMNS.UPDATE_FLAG).setValue('エラー');
        continue;
      }
      
      var video = response.items[0];
      
      video.snippet.title = newTitle;
      video.snippet.description = newDescription;
      
      if (newTags) {
        var tagsArray = String(newTags).split(/[,、]/);
        var cleanTags = [];
        for (var t = 0; t < tagsArray.length; t++) {
          var tag = tagsArray[t].trim();
          if (tag.length > 0) {
            cleanTags.push(tag);
          }
        }
        video.snippet.tags = cleanTags;
      } else {
        video.snippet.tags = [];
      }
      
      if (!video.snippet.categoryId) {
        video.snippet.categoryId = categoryId || '22';
      }
      
      YouTube.Videos.update(video, 'snippet');
      
      sheet.getRange(row, COLUMNS.CURRENT_TITLE).setValue(newTitle);
      sheet.getRange(row, COLUMNS.CURRENT_DESCRIPTION).setValue(newDescription);
      sheet.getRange(row, COLUMNS.CURRENT_TAGS).setValue(newTags);
      
      sheet.getRange(row, COLUMNS.NEW_TITLE, 1, 3).clearContent();
      
      processedRows.push(row);
      updatedCount++;
      
    } catch (e) {
      errors.push('行' + row + ': ' + e.message);
      sheet.getRange(row, COLUMNS.UPDATE_FLAG).setValue('エラー');
    }
    
    Utilities.sleep(200);
  }
  
  for (var p = 0; p < processedRows.length; p++) {
    sheet.getRange(processedRows[p], COLUMNS.UPDATE_FLAG).setValue('完了');
  }
  
  var message = updatedCount + '件の動画を更新しました';
  if (errors.length > 0) {
    var errorList = errors.slice(0, 10);
    message += '\n\nエラー:\n' + errorList.join('\n');
    if (errors.length > 10) {
      message += '\n... 他' + (errors.length - 10) + '件';
    }
  }
  showAlert(message);
}

// ==================== 編集サポート機能 ====================

function copyCurrentToNew() {
  var sheet = getSheet();
  var ui = SpreadsheetApp.getUi();
  
  var response = ui.alert(
    'コピー方法を選択',
    '選択中の行のみコピーしますか？\n\n' +
    '「はい」→ 選択行のみ\n' +
    '「いいえ」→ 全データ行',
    ui.ButtonSet.YES_NO_CANCEL
  );
  
  if (response === ui.Button.CANCEL) return;
  
  var startRow = 3;
  var targetRows = [];
  
  if (response === ui.Button.YES) {
    var selection = sheet.getActiveRange();
    var selStartRow = selection.getRow();
    var selNumRows = selection.getNumRows();
    
    for (var i = 0; i < selNumRows; i++) {
      var row = selStartRow + i;
      if (row >= startRow) {
        targetRows.push(row);
      }
    }
  } else {
    var lastRow = sheet.getLastRow();
    for (var row = startRow; row <= lastRow; row++) {
      targetRows.push(row);
    }
  }
  
  if (targetRows.length === 0) {
    showAlert('対象の行がありません');
    return;
  }
  
  var copyCount = 0;
  for (var j = 0; j < targetRows.length; j++) {
    var r = targetRows[j];
    var currentTitle = sheet.getRange(r, COLUMNS.CURRENT_TITLE).getValue();
    var currentDesc = sheet.getRange(r, COLUMNS.CURRENT_DESCRIPTION).getValue();
    var currentTags = sheet.getRange(r, COLUMNS.CURRENT_TAGS).getValue();
    
    if (currentTitle) {
      sheet.getRange(r, COLUMNS.NEW_TITLE).setValue(currentTitle);
      sheet.getRange(r, COLUMNS.NEW_DESCRIPTION).setValue(currentDesc);
      sheet.getRange(r, COLUMNS.NEW_TAGS).setValue(currentTags);
      copyCount++;
    }
  }
  
  showAlert(copyCount + '行の現在値を修正後列にコピーしました\n\n' +
    'F-H列を編集して、U列に「○」を入力してください。');
}

function highlightDifferences() {
  var sheet = getSheet();
  var startRow = 3;
  var lastRow = sheet.getLastRow();
  
  if (lastRow < startRow) {
    showAlert('データがありません');
    return;
  }
  
  var data = sheet.getRange(startRow, 1, lastRow - startRow + 1, COLUMNS.NEW_TAGS).getValues();
  var diffCount = 0;
  
  sheet.getRange(startRow, COLUMNS.NEW_TITLE, lastRow - startRow + 1, 3).setBackground('#ffffff');
  
  for (var i = 0; i < data.length; i++) {
    var row = startRow + i;
    var currentTitle = data[i][COLUMNS.CURRENT_TITLE - 1] || '';
    var currentDesc = data[i][COLUMNS.CURRENT_DESCRIPTION - 1] || '';
    var currentTags = data[i][COLUMNS.CURRENT_TAGS - 1] || '';
    var newTitle = data[i][COLUMNS.NEW_TITLE - 1] || '';
    var newDesc = data[i][COLUMNS.NEW_DESCRIPTION - 1] || '';
    var newTags = data[i][COLUMNS.NEW_TAGS - 1] || '';
    
    if (newTitle && newTitle !== currentTitle) {
      sheet.getRange(row, COLUMNS.NEW_TITLE).setBackground('#fff2cc');
      diffCount++;
    }
    if (newDesc && newDesc !== currentDesc) {
      sheet.getRange(row, COLUMNS.NEW_DESCRIPTION).setBackground('#fff2cc');
      diffCount++;
    }
    if (newTags && newTags !== currentTags) {
      sheet.getRange(row, COLUMNS.NEW_TAGS).setBackground('#fff2cc');
      diffCount++;
    }
  }
  
  showAlert(diffCount + '箇所の差分をハイライトしました');
}

function clearHighlights() {
  var sheet = getSheet();
  var startRow = 3;
  var lastRow = sheet.getLastRow();
  
  if (lastRow >= startRow) {
    sheet.getRange(startRow, COLUMNS.NEW_TITLE, lastRow - startRow + 1, 3).setBackground('#ffffff');
  }
  
  showAlert('ハイライトをクリアしました');
}

// ==================== サムネイル関連 ====================

function fetchThumbnails() {
  var sheet = getSheet();
  var data = sheet.getDataRange().getValues();
  var startRow = 3;
  
  var count = 0;
  for (var i = startRow - 1; i < data.length; i++) {
    var videoIdRaw = data[i][0];
    var videoId = extractVideoId(videoIdRaw ? String(videoIdRaw) : '');
    if (!videoId) continue;
    
    try {
      var response = YouTube.Videos.list('snippet', { id: videoId });
      if (response.items && response.items.length > 0) {
        var thumbnailUrl = getThumbnailUrl(response.items[0].snippet.thumbnails);
        sheet.getRange(i + 1, COLUMNS.THUMBNAIL_URL).setValue(thumbnailUrl);
        count++;
      }
    } catch (e) {
      Logger.log('サムネイル取得エラー (行' + (i + 1) + '): ' + e.message);
    }
  }
  
  showAlert(count + '件のサムネイルURLを取得しました');
}

function showThumbnailImages() {
  var sheet = getSheet();
  var data = sheet.getDataRange().getValues();
  var startRow = 3;
  
  var count = 0;
  for (var i = startRow - 1; i < data.length; i++) {
    var thumbnailUrl = data[i][COLUMNS.THUMBNAIL_URL - 1];
    if (thumbnailUrl) {
      sheet.getRange(i + 1, COLUMNS.THUMBNAIL_IMAGE).setFormula('=IMAGE("' + thumbnailUrl + '")');
      count++;
    }
  }
  
  if (count > 0) {
    sheet.setRowHeightsForced(startRow, sheet.getLastRow() - startRow + 1, 90);
  }
  
  showAlert(count + '件のサムネイル画像を表示しました');
}

// ==================== ヘルプ ====================

function showHelp() {
  var helpText = 'YouTube管理ツール ヘルプ（v3）\n\n' +
    '【初期設定の流れ】\n' +
    '1. 「初期セットアップ」を実行\n' +
    '2. 「設定シートにチャンネルを追加」でチャンネル登録\n' +
    '3. 設定シートで「有効」にチェック\n' +
    '4. ブランドアカウントの場合は「AnalyticsチャンネルID」を設定\n' +
    '5. 「定期実行を開始」で自動取得開始\n\n' +
    '【列の構成】\n' +
    '・A列: 動画ID\n' +
    '・B列: 動画URL\n' +
    '・C-E列: 現在のYouTube値\n' +
    '・F-H列: 修正後の値\n' +
    '・I-L列: 基本統計\n' +
    '・M-Q列: Analytics統計\n\n' +
    '【動画情報の更新】\n' +
    '1. F-H列を編集\n' +
    '2. U列に「○」を入力\n' +
    '3. 「YouTubeに反映」を実行\n\n' +
    '【ブランドアカウントの場合】\n' +
    '設定シートの「AnalyticsチャンネルID」に\n' +
    'チャンネルID（UCxxxx形式）を入力してください。';
  
  showAlert(helpText);
}

// ==================== ユーティリティ関数 ====================

function getSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) {
    sheet = ss.getActiveSheet();
  }
  return sheet;
}

function extractVideoId(input) {
  if (!input) return '';
  var str = String(input).trim();
  
  var patterns = [
    /[?&]v=([a-zA-Z0-9_-]{11})/,
    /youtu\.be\/([a-zA-Z0-9_-]{11})/,
    /youtube\.com\/embed\/([a-zA-Z0-9_-]{11})/,
    /youtube\.com\/v\/([a-zA-Z0-9_-]{11})/,
    /^([a-zA-Z0-9_-]{11})$/
  ];
  
  for (var i = 0; i < patterns.length; i++) {
    var match = str.match(patterns[i]);
    if (match) return match[1];
  }
  
  return str;
}

function getVideoUrl(videoId) {
  return 'https://www.youtube.com/watch?v=' + videoId;
}

function getThumbnailUrl(thumbnails) {
  var settings = getGeneralSettings();
  var size = settings.thumbnailSize || 'medium';
  var priorities = ['medium', 'high', 'standard', 'maxres', 'default'];
  
  if (thumbnails[size]) {
    return thumbnails[size].url;
  }
  
  for (var i = 0; i < priorities.length; i++) {
    if (thumbnails[priorities[i]]) {
      return thumbnails[priorities[i]].url;
    }
  }
  
  return '';
}

function formatDate(isoString) {
  if (!isoString) return '';
  var date = new Date(isoString);
  return Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');
}

function showAlert(message) {
  try {
    SpreadsheetApp.getUi().alert(message);
  } catch (e) {
    Logger.log(message);
  }
}

// ==================== デバッグ ====================

function debugSettings() {
  var settings = getGeneralSettings();
  Logger.log('=== 設定内容 ===');
  Logger.log(JSON.stringify(settings, null, 2));
}

function debugAnalyticsChannelId() {
  Logger.log('=== チャンネルID確認 ===');
  Logger.log('個人: ' + getMyChannelId());
  Logger.log('Analytics用: ' + getAnalyticsChannelId());
}