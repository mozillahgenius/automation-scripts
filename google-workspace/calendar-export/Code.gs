function onOpen() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  // 初回セットアップチェック
  const configSheet = spreadsheet.getSheetByName('Config');
  const groupsSheet = spreadsheet.getSheetByName('Groups');
  const eventsSheet = spreadsheet.getSheetByName('Events');

  if (!configSheet || !groupsSheet || !eventsSheet) {
    // シートが存在しない場合は自動セットアップを実行
    const response = ui.alert(
      '初期セットアップ',
      'カレンダー抽出ツールの初期セットアップを行います。\n必要なシートを自動作成しますか？',
      ui.ButtonSet.YES_NO
    );

    if (response === ui.Button.YES) {
      setupSheets();
      showInitialGuide();
    }
  }

  ui.createMenu('Calendar Export')
    .addItem('Setup Sheets', 'setupSheets')
    .addItem('Run Export', 'runExport')
    .addItem('Clear Events Sheet', 'clearEventsSheet')
    .addSeparator()
    .addItem('Quick Setup Guide', 'showQuickGuide')
    .addItem('Test Connection', 'testConnection')
    .addToUi();
}

function onInstall(e) {
  onOpen(e);
}

function setupSheets() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  let configSheet = spreadsheet.getSheetByName('Config');
  if (!configSheet) {
    configSheet = spreadsheet.insertSheet('Config');
    configSheet.getRange('A1:B1').setValues([['KEY', 'VALUE']]);

    // 日付を文字列として設定（書式なしテキストとして扱う）
    const configValues = [
      ['START_DATE', '2025-09-01'],
      ['END_DATE', '2025-09-30'],
      ['CALENDAR_IDS', 'primary'],
      ['GROUP_MODE', 'DIRECT_GROUP'],
      ['GROUP_MATCH', 'ANY'],
      ['TIMEZONE', 'Asia/Tokyo'],
      ['MAX_RESULTS', '500']
    ];

    configSheet.getRange('A2:B8').setValues(configValues);

    // 日付セルを文字列形式に設定
    configSheet.getRange('B2:B3').setNumberFormat('@');  // @は文字列形式

    configSheet.getRange('A1:B1').setFontWeight('bold');
    configSheet.autoResizeColumns(1, 2);

    // 注意事項を追加
    configSheet.getRange('D2').setValue('※日付はYYYY-MM-DD形式で入力');
    configSheet.getRange('D2').setFontColor('#666666');
    configSheet.getRange('D2').setFontSize(10);
  }

  let groupsSheet = spreadsheet.getSheetByName('Groups');
  if (!groupsSheet) {
    groupsSheet = spreadsheet.insertSheet('Groups');
    groupsSheet.getRange('A1').setValue('Group Email');
    groupsSheet.getRange('A2').setValue('team@example.com');
    groupsSheet.getRange('A3').setValue('all@example.com');
    groupsSheet.getRange('A1').setFontWeight('bold');
    groupsSheet.autoResizeColumn(1);
  }

  let eventsSheet = spreadsheet.getSheetByName('Events');
  if (!eventsSheet) {
    eventsSheet = spreadsheet.insertSheet('Events');
    const headers = [
      'CalendarId', 'EventId', 'Summary', 'Start', 'End',
      'Organizer', 'Attendees', 'MatchMode', 'MatchedGroup',
      'MatchDetail', 'Location', 'HangoutLink', 'Created',
      'Updated', 'Visibility', 'Status', 'Description'
    ];
    eventsSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    eventsSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    eventsSheet.setFrozenRows(1);
    eventsSheet.autoResizeColumns(1, headers.length);
  }

  SpreadsheetApp.getActiveSpreadsheet().toast('シートのセットアップが完了しました', 'セットアップ完了', 3);
}

function getConfig() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = spreadsheet.getSheetByName('Config');

  if (!configSheet) {
    throw new Error('Configシートが見つかりません。先にSetup Sheetsを実行してください。');
  }

  const lastRow = configSheet.getLastRow();
  const configData = lastRow >= 2 ? configSheet.getRange(2, 1, lastRow - 1, 2).getValues() : [];
  const config = {};

  configData.forEach(row => {
    const key = row[0];
    let value = row[1];

    if (key) {
      // 日付オブジェクトの場合は、YYYY-MM-DD形式に変換
      if (value instanceof Date) {
        const year = value.getFullYear();
        const month = String(value.getMonth() + 1).padStart(2, '0');
        const day = String(value.getDate()).padStart(2, '0');
        value = `${year}-${month}-${day}`;
        console.log(`日付を変換: ${key} = ${value}`);
      } else {
        value = String(value).trim();
      }

      config[key] = value;
    }
  });

  // 日付フォーマットの検証と修正
  if (!config.START_DATE || !config.END_DATE) {
    throw new Error('START_DATEとEND_DATEを設定してください');
  }

  // 日付形式の検証（YYYY-MM-DD）
  const datePattern = /^\d{4}-\d{2}-\d{2}$/;

  // もし日付形式が正しくない場合、自動修正を試みる
  if (!datePattern.test(config.START_DATE)) {
    // "2025/09/01" や "2025.09.01" などの形式を修正
    config.START_DATE = config.START_DATE.replace(/[\/\.]/g, '-');

    // まだ正しくない場合はエラー
    if (!datePattern.test(config.START_DATE)) {
      throw new Error(`START_DATE の形式が正しくありません: ${config.START_DATE}（YYYY-MM-DD形式で入力してください）`);
    }
  }

  if (!datePattern.test(config.END_DATE)) {
    config.END_DATE = config.END_DATE.replace(/[\/\.]/g, '-');

    if (!datePattern.test(config.END_DATE)) {
      throw new Error(`END_DATE の形式が正しくありません: ${config.END_DATE}（YYYY-MM-DD形式で入力してください）`);
    }
  }

  config.calendarIds = config.CALENDAR_IDS ?
    config.CALENDAR_IDS.split(',').map(id => id.trim()).filter(id => id) : ['primary'];

  // デフォルト値の設定
  config.GROUP_MODE = config.GROUP_MODE || 'DIRECT_GROUP';
  config.GROUP_MATCH = config.GROUP_MATCH || 'ANY';
  config.TIMEZONE = config.TIMEZONE || 'Asia/Tokyo';
  config.MAX_RESULTS = config.MAX_RESULTS || '500';

  console.log('設定を読み込みました:', {
    START_DATE: config.START_DATE,
    END_DATE: config.END_DATE,
    CALENDAR_IDS: config.calendarIds.join(', '),
    GROUP_MODE: config.GROUP_MODE
  });

  return config;
}

function getGroups() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const groupsSheet = spreadsheet.getSheetByName('Groups');

  if (!groupsSheet) {
    throw new Error('Groupsシートが見つかりません。先にSetup Sheetsを実行してください。');
  }

  const lastRow = groupsSheet.getLastRow();
  if (lastRow < 2) return [];

  const groupsData = groupsSheet.getRange(2, 1, lastRow - 1, 1).getValues();
  return groupsData.flat().filter(email => email && email.trim());
}

function clearEventsSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const eventsSheet = spreadsheet.getSheetByName('Events');

  if (!eventsSheet) {
    SpreadsheetApp.getActiveSpreadsheet().toast('Eventsシートが見つかりません', 'エラー', 3);
    return;
  }

  const lastRow = eventsSheet.getLastRow();
  if (lastRow > 1) {
    eventsSheet.getRange(2, 1, lastRow - 1, eventsSheet.getLastColumn()).clear();
  }

  SpreadsheetApp.getActiveSpreadsheet().toast('Eventsシートをクリアしました', '完了', 3);
}

function runExport() {
  try {
    const config = getConfig();
    const groups = getGroups();

    if (groups.length === 0) {
      SpreadsheetApp.getActiveSpreadsheet().toast('Groupsシートにグループメールアドレスを入力してください', 'エラー', 5);
      return;
    }

    SpreadsheetApp.getActiveSpreadsheet().toast('エクスポートを開始します...', '処理中', 3);

    let groupMembers = {};
    if (config.GROUP_MODE === 'MEMBER_OF_GROUP') {
      groupMembers = fetchAllGroupMembers(groups);
    }

    // 日付のバリデーションと修正
    let startDate, endDate;
    try {
      // タイムゾーンを考慮した日付生成
      const timezone = config.TIMEZONE || 'Asia/Tokyo';

      // 日付文字列をチェック
      if (!config.START_DATE || !config.END_DATE) {
        throw new Error('開始日と終了日を設定してください');
      }

      // ISO形式で日付を作成（タイムゾーンを明示）
      startDate = new Date(config.START_DATE + 'T00:00:00');
      endDate = new Date(config.END_DATE + 'T23:59:59');

      // 日付が有効かチェック
      if (isNaN(startDate.getTime()) || isNaN(endDate.getTime())) {
        throw new Error('日付形式が正しくありません。YYYY-MM-DD形式で入力してください');
      }

      // 開始日が終了日より後でないかチェック
      if (startDate > endDate) {
        throw new Error('開始日は終了日より前の日付を設定してください');
      }

    } catch (dateError) {
      SpreadsheetApp.getActiveSpreadsheet().toast(
        `日付エラー: ${dateError.message}`,
        'エラー',
        10
      );
      return;
    }

    let allEvents = [];
    let errorCalendars = [];

    for (const calendarId of config.calendarIds) {
      try {
        const events = fetchCalendarEvents(calendarId, startDate, endDate, config);

        for (const event of events) {
          const matchResult = checkGroupMatch(event, groups, groupMembers, config);

          if (matchResult.isMatch) {
            allEvents.push({
              calendarId: calendarId,
              event: event,
              matchResult: matchResult
            });
          }
        }
      } catch (error) {
        console.error(`カレンダー ${calendarId} の処理中にエラー:`, error);
        errorCalendars.push(calendarId);
      }
    }

    if (errorCalendars.length > 0) {
      SpreadsheetApp.getActiveSpreadsheet().toast(
        `一部のカレンダーでエラーが発生しました: ${errorCalendars.join(', ')}`,
        '警告',
        5
      );
    }

    writeEventsToSheet(allEvents, config);

    SpreadsheetApp.getActiveSpreadsheet().toast(
      `エクスポート完了: ${allEvents.length}件のイベントを抽出しました`,
      '完了',
      5
    );

  } catch (error) {
    console.error('エクスポート中にエラーが発生しました:', error);
    SpreadsheetApp.getActiveSpreadsheet().toast(
      `エラー: ${error.message}`,
      'エラー',
      10
    );
  }
}

function fetchCalendarEvents(calendarId, startDate, endDate, config) {
  const events = [];
  let pageToken = null;
  const maxResults = parseInt(config.MAX_RESULTS) || 500;

  try {
    // カレンダーIDの存在確認
    try {
      const calendarInfo = Calendar.Calendars.get(calendarId);
      console.log(`カレンダー確認: ${calendarInfo.summary || calendarId}`);
    } catch (calError) {
      console.error(`カレンダー ${calendarId} が見つからない、またはアクセス権限がありません`);
      throw new Error(`カレンダー ${calendarId} へのアクセスエラー`);
    }

    // 日付の再確認
    if (!startDate || !endDate || isNaN(startDate.getTime()) || isNaN(endDate.getTime())) {
      throw new Error('日付が無効です');
    }

    do {
      const options = {
        timeMin: startDate.toISOString(),
        timeMax: endDate.toISOString(),
        singleEvents: true,
        orderBy: 'startTime',
        maxResults: Math.min(maxResults - events.length, 250),
        showDeleted: false
      };

      if (pageToken) {
        options.pageToken = pageToken;
      }

      console.log(`イベント取得中: ${calendarId}, 期間: ${options.timeMin} ～ ${options.timeMax}`);

      const response = Calendar.Events.list(calendarId, options);

      if (response.items && response.items.length > 0) {
        console.log(`${response.items.length}件のイベントを取得`);
        events.push(...response.items);
      }

      pageToken = response.nextPageToken;

      if (events.length >= maxResults) {
        console.log(`最大取得件数（${maxResults}件）に達しました`);
        break;
      }

    } while (pageToken);

    console.log(`カレンダー ${calendarId} から合計 ${events.length}件のイベントを取得しました`);

  } catch (error) {
    console.error(`Calendar API エラー (${calendarId}):`, error);
    console.error('エラーの詳細:', {
      message: error.message,
      stack: error.stack,
      toString: error.toString()
    });
    throw new Error(`カレンダー ${calendarId} の処理エラー: ${error.message}`);
  }

  return events;
}

function fetchAllGroupMembers(groups) {
  const membersMap = {};

  for (const groupEmail of groups) {
    try {
      const members = fetchGroupMembers(groupEmail);
      membersMap[groupEmail] = members;
    } catch (error) {
      console.log(`グループ ${groupEmail} のメンバー取得をスキップ (エラー: ${error.message})`);
      membersMap[groupEmail] = [];
    }
  }

  return membersMap;
}

function fetchGroupMembers(groupEmail) {
  const members = [];
  let pageToken = null;

  try {
    if (typeof AdminDirectory === 'undefined') {
      console.log('Admin SDK Directory APIが有効になっていません。DIRECT_GROUPモードで動作します。');
      return [];
    }

    do {
      const options = { maxResults: 200 };
      if (pageToken) {
        options.pageToken = pageToken;
      }

      const response = AdminDirectory.Members.list(groupEmail, options);

      if (response.members) {
        response.members.forEach(member => {
          members.push(member.email.toLowerCase());
        });
      }

      pageToken = response.nextPageToken;
    } while (pageToken);

  } catch (error) {
    console.error(`グループメンバー取得エラー (${groupEmail}):`, error);
    return [];
  }

  return members;
}

function checkGroupMatch(event, groups, groupMembers, config) {
  const result = {
    isMatch: false,
    matchMode: config.GROUP_MODE,
    matchedGroups: [],
    matchDetails: []
  };

  const attendees = [];
  if (event.organizer && event.organizer.email) {
    attendees.push(event.organizer.email.toLowerCase());
  }
  if (event.attendees) {
    event.attendees.forEach(attendee => {
      if (attendee.email) {
        attendees.push(attendee.email.toLowerCase());
      }
    });
  }

  const uniqueAttendees = [...new Set(attendees)];

  for (const group of groups) {
    const groupLower = group.toLowerCase();
    let groupMatched = false;

    if (config.GROUP_MODE === 'DIRECT_GROUP') {
      if (uniqueAttendees.includes(groupLower)) {
        groupMatched = true;
        result.matchedGroups.push(group);
        result.matchDetails.push('GROUP_INVITED');
      }
    } else if (config.GROUP_MODE === 'MEMBER_OF_GROUP') {
      const members = groupMembers[group] || [];
      const matchedMembers = uniqueAttendees.filter(attendee => members.includes(attendee));

      if (matchedMembers.length > 0) {
        groupMatched = true;
        result.matchedGroups.push(group);
        result.matchDetails.push(`MEMBER_INVITED:${matchedMembers.join(',')}`);
      }
    }

    if (groupMatched && config.GROUP_MATCH === 'ANY') {
      result.isMatch = true;
      break;
    }
  }

  if (config.GROUP_MATCH === 'ALL' && result.matchedGroups.length === groups.length) {
    result.isMatch = true;
  } else if (config.GROUP_MATCH === 'ANY' && result.matchedGroups.length > 0) {
    result.isMatch = true;
  }

  return result;
}

function writeEventsToSheet(eventsData, config) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const eventsSheet = spreadsheet.getSheetByName('Events');

  if (!eventsSheet) {
    throw new Error('Eventsシートが見つかりません');
  }

  clearEventsSheet();

  if (eventsData.length === 0) {
    return;
  }

  const timezone = config.TIMEZONE || 'Asia/Tokyo';
  const rows = [];

  for (const data of eventsData) {
    const event = data.event;
    const matchResult = data.matchResult;

    let startTime = '';
    let endTime = '';

    if (event.start) {
      if (event.start.dateTime) {
        startTime = Utilities.formatDate(new Date(event.start.dateTime), timezone, 'yyyy-MM-dd HH:mm:ss');
      } else if (event.start.date) {
        startTime = event.start.date;
      }
    }

    if (event.end) {
      if (event.end.dateTime) {
        endTime = Utilities.formatDate(new Date(event.end.dateTime), timezone, 'yyyy-MM-dd HH:mm:ss');
      } else if (event.end.date) {
        endTime = event.end.date;
      }
    }

    const attendeesList = event.attendees ?
      event.attendees.map(a => a.email).filter(e => e).join(', ') : '';

    const hangoutLink = event.hangoutLink ||
      (event.conferenceData && event.conferenceData.entryPoints ?
        event.conferenceData.entryPoints[0].uri : '');

    const createdDate = event.created ?
      Utilities.formatDate(new Date(event.created), timezone, 'yyyy-MM-dd HH:mm:ss') : '';
    const updatedDate = event.updated ?
      Utilities.formatDate(new Date(event.updated), timezone, 'yyyy-MM-dd HH:mm:ss') : '';

    rows.push([
      data.calendarId,
      event.id || '',
      event.summary || '(タイトルなし)',
      startTime,
      endTime,
      event.organizer ? event.organizer.email : '',
      attendeesList,
      matchResult.matchMode,
      matchResult.matchedGroups.join(', '),
      matchResult.matchDetails.join(' | '),
      event.location || '',
      hangoutLink,
      createdDate,
      updatedDate,
      event.visibility || 'default',
      event.status || '',
      event.description ? event.description.substring(0, 500) : ''
    ]);
  }

  if (rows.length > 0) {
    eventsSheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
    eventsSheet.autoResizeColumns(1, 17);
  }
}

function testConnection() {
  try {
    const calendars = Calendar.CalendarList.list();
    console.log('利用可能なカレンダー:');
    calendars.items.forEach(cal => {
      console.log(`- ${cal.summary} (${cal.id})`);
    });

    SpreadsheetApp.getActiveSpreadsheet().toast('カレンダーAPIへの接続に成功しました', '成功', 3);
  } catch (error) {
    console.error('接続テストエラー:', error);
    SpreadsheetApp.getActiveSpreadsheet().toast(`エラー: ${error.message}`, 'エラー', 5);
  }
}

function showInitialGuide() {
  const ui = SpreadsheetApp.getUi();
  const htmlContent = HtmlService.createHtmlOutput(`
    <div style="font-family: Arial, sans-serif; padding: 20px; max-width: 500px;">
      <h2>🎉 セットアップ完了！</h2>
      <p>カレンダー抽出ツールのセットアップが完了しました。</p>

      <h3>📋 次のステップ：</h3>
      <ol style="line-height: 1.8;">
        <li><b>Configシート</b>で抽出期間とカレンダーIDを設定</li>
        <li><b>Groupsシート</b>に対象グループのメールアドレスを入力</li>
        <li>メニューの<b>「Calendar Export」→「Run Export」</b>を実行</li>
        <li><b>Eventsシート</b>に結果が出力されます</li>
      </ol>

      <h3>⚙️ 設定のポイント：</h3>
      <ul style="line-height: 1.6;">
        <li><b>CALENDAR_IDS</b>: 「primary」で自分のカレンダー</li>
        <li><b>GROUP_MODE</b>:
          <ul>
            <li>DIRECT_GROUP: グループ直接招待のみ</li>
            <li>MEMBER_OF_GROUP: メンバー招待も含む（要Admin権限）</li>
          </ul>
        </li>
        <li><b>GROUP_MATCH</b>: ANY（いずれか） or ALL（すべて）</li>
      </ul>

      <h3>💡 ヒント：</h3>
      <p style="background: #f0f0f0; padding: 10px; border-left: 3px solid #4285f4;">
        初回実行時には権限の承認が必要です。<br>
        「承認が必要です」と表示されたら指示に従ってください。
      </p>

      <div style="margin-top: 20px; text-align: center;">
        <button onclick="google.script.host.close()"
                style="background: #4285f4; color: white; padding: 10px 20px;
                       border: none; border-radius: 4px; cursor: pointer;">
          閉じる
        </button>
      </div>
    </div>
  `)
  .setWidth(550)
  .setHeight(600);

  ui.showModalDialog(htmlContent, 'カレンダー抽出ツール - セットアップガイド');
}

function showQuickGuide() {
  const ui = SpreadsheetApp.getUi();
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  // 現在の設定を取得
  let currentConfig = '未設定';
  let currentGroups = '未設定';

  try {
    const configSheet = spreadsheet.getSheetByName('Config');
    if (configSheet) {
      const config = getConfig();
      currentConfig = `
        期間: ${config.START_DATE} ～ ${config.END_DATE}<br>
        カレンダー: ${config.CALENDAR_IDS}<br>
        モード: ${config.GROUP_MODE}
      `;
    }

    const groups = getGroups();
    if (groups.length > 0) {
      currentGroups = groups.slice(0, 3).join('<br>');
      if (groups.length > 3) {
        currentGroups += `<br>他 ${groups.length - 3} 件`;
      }
    }
  } catch (e) {
    // エラーは無視
  }

  const htmlContent = HtmlService.createHtmlOutput(`
    <div style="font-family: Arial, sans-serif; padding: 20px;">
      <h2>📊 カレンダー抽出ツール</h2>

      <h3>現在の設定：</h3>
      <div style="background: #f8f8f8; padding: 10px; border-radius: 4px;">
        <b>Config設定:</b><br>
        ${currentConfig}
        <br><br>
        <b>対象グループ:</b><br>
        ${currentGroups}
      </div>

      <h3>使い方：</h3>
      <table style="width: 100%; border-collapse: collapse;">
        <tr>
          <td style="padding: 8px; border-bottom: 1px solid #ddd;">
            <b>1. Setup Sheets</b>
          </td>
          <td style="padding: 8px; border-bottom: 1px solid #ddd;">
            必要なシートを作成/リセット
          </td>
        </tr>
        <tr>
          <td style="padding: 8px; border-bottom: 1px solid #ddd;">
            <b>2. Run Export</b>
          </td>
          <td style="padding: 8px; border-bottom: 1px solid #ddd;">
            カレンダーデータを抽出
          </td>
        </tr>
        <tr>
          <td style="padding: 8px; border-bottom: 1px solid #ddd;">
            <b>3. Clear Events</b>
          </td>
          <td style="padding: 8px; border-bottom: 1px solid #ddd;">
            Eventsシートをクリア
          </td>
        </tr>
        <tr>
          <td style="padding: 8px; border-bottom: 1px solid #ddd;">
            <b>4. Test Connection</b>
          </td>
          <td style="padding: 8px;">
            API接続テスト
          </td>
        </tr>
      </table>

      <h3>トラブルシューティング：</h3>
      <ul style="line-height: 1.6;">
        <li>「Calendar is not defined」エラー<br>
          → Calendar APIを有効化してください</li>
        <li>イベントが抽出されない<br>
          → 日付形式（YYYY-MM-DD）を確認</li>
        <li>グループメンバーが取得できない<br>
          → Admin権限が必要です</li>
      </ul>

      <div style="margin-top: 20px; text-align: center;">
        <button onclick="google.script.host.close()"
                style="background: #4285f4; color: white; padding: 10px 20px;
                       border: none; border-radius: 4px; cursor: pointer;">
          閉じる
        </button>
      </div>
    </div>
  `)
  .setWidth(500)
  .setHeight(600);

  ui.showModalDialog(htmlContent, 'クイックガイド');
}