// ==========================================
// スプレッドシート初期化スクリプト
// ==========================================

function initializeSpreadsheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 既存のシートをクリア（最初のシートは残す）
  const sheets = ss.getSheets();
  for (let i = sheets.length - 1; i > 0; i--) {
    ss.deleteSheet(sheets[i]);
  }
  
  // 各シートを作成
  createConfigSheet(ss);
  createSyncStateSheet(ss);
  createMessagesSheet(ss);
  createCategoriesSheet(ss);
  createChecklistsSheet(ss);
  createTemplatesSheet(ss);
  createDraftsSheet(ss);
  createLogsSheet(ss);
  
  // 最初のデフォルトシートを削除
  try {
    ss.deleteSheet(sheets[0]);
  } catch (e) {
    // 削除できない場合は無視
  }
  
  SpreadsheetApp.getUi().alert('スプレッドシートの初期化が完了しました');
}

// ========= Config シート =========
function createConfigSheet(ss) {
  const sheet = ss.insertSheet('Config');
  
  const headers = ['設定項目', '値', '説明'];
  const data = [
    ['company', '', '会社名'],
    ['targetChannels', '', 'Slack監視対象チャンネルID（カンマ区切り）'],
    ['notifySlackChannel', '', '通知先SlackチャンネルID'],
    ['notifyEmails', '', '通知先メールアドレス（カンマ区切り）'],
    ['openaiModel', 'gpt-4o-mini', 'OpenAIモデル名（要約・分類用）'],
    ['openaiModelDraft', 'gpt-4o', 'OpenAIモデル名（ドラフト生成用）'],
    ['OPENAI_MODEL', 'gpt-4o', 'メイン処理用OpenAIモデル名'],
    ['classificationThreshold', '0.6', '該当判定しきい値（0-1）'],
    ['syncIntervalMinutes', '5', 'Slack同期間隔（分）'],
    ['analysisIntervalHours', '1', 'AI分析実行間隔（時間）'],
    ['notificationHours', '9,15', '通知時刻（カンマ区切り）']
  ];
  
  sheet.getRange(1, 1, 1, 3).setValues([headers]);
  sheet.getRange(2, 1, data.length, 3).setValues(data);
  
  // ヘッダーのフォーマット
  sheet.getRange(1, 1, 1, 3)
    .setBackground('#4285F4')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold');
  
  // 列幅の調整
  sheet.setColumnWidth(1, 200);
  sheet.setColumnWidth(2, 300);
  sheet.setColumnWidth(3, 400);
  
  // フリーズ
  sheet.setFrozenRows(1);
}

// ========= SyncState シート =========
function createSyncStateSheet(ss) {
  const sheet = ss.insertSheet('SyncState');
  
  const headers = ['channel_id', 'last_sync_ts', 'last_sync_datetime', 'message_count', 'status'];
  
  sheet.getRange(1, 1, 1, 5).setValues([headers]);
  
  // ヘッダーのフォーマット
  sheet.getRange(1, 1, 1, 5)
    .setBackground('#34A853')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold');
  
  // 列幅の調整
  sheet.setColumnWidth(1, 150);
  sheet.setColumnWidth(2, 150);
  sheet.setColumnWidth(3, 180);
  sheet.setColumnWidth(4, 120);
  sheet.setColumnWidth(5, 100);
  
  sheet.setFrozenRows(1);
}

// ========= Messages シート =========
function createMessagesSheet(ss) {
  const sheet = ss.insertSheet('Messages');
  
  const headers = [
    'id',
    'channel_id',
    'message_ts',
    'thread_ts',
    'text_raw',
    'user_name',
    'summary_json',
    'classification_json',
    'match_flag',
    'human_judgement',
    'permalink',
    'checklist_proposed',
    'agenda_selected',
    'draft_doc_url',
    'created_at',
    'updated_at'
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // ヘッダーのフォーマット
  sheet.getRange(1, 1, 1, headers.length)
    .setBackground('#EA4335')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold');
  
  // 列幅の調整
  sheet.setColumnWidth(1, 200); // id
  sheet.setColumnWidth(2, 120); // channel_id
  sheet.setColumnWidth(3, 150); // message_ts
  sheet.setColumnWidth(4, 150); // thread_ts
  sheet.setColumnWidth(5, 400); // text_raw
  sheet.setColumnWidth(6, 150); // user_name
  sheet.setColumnWidth(7, 300); // summary_json
  sheet.setColumnWidth(8, 300); // classification_json
  sheet.setColumnWidth(9, 100); // match_flag
  sheet.setColumnWidth(10, 120); // human_judgement
  sheet.setColumnWidth(11, 250); // permalink
  
  // データ検証の設定（human_judgement列）
  const validationRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['必要', '不要', '保留'], true)
    .setAllowInvalid(false)
    .build();
  
  sheet.getRange(2, 10, 1000, 1).setDataValidation(validationRule);
  
  // チェックボックスの設定（match_flag列）
  sheet.getRange(2, 9, 1000, 1).insertCheckboxes();
  
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(1);
}

// ========= Categories シート =========
function createCategoriesSheet(ss) {
  const sheet = ss.insertSheet('Categories');
  
  const headers = ['カテゴリ名', '説明', '判定基準', 'キーワード（カンマ区切り）', '重要度'];
  
  const sampleData = [
    ['開示事項', '適時開示が必要な事項', '金融商品取引法、取引所規則に基づく開示が必要な事項', '決算,業績予想,配当,買収,合併,提携,新株発行', '高'],
    ['取締役会決議事項', '取締役会での決議が必要な事項', '会社法および定款で定められた取締役会決議事項', '重要な財産,借入,投資,組織変更,人事,規程改定', '高'],
    ['監査等委員会決議事項', '監査等委員会での決議が必要な事項', '監査等委員会の職務に関する事項', '監査計画,監査報告,会計監査人,内部統制', '中'],
    ['株主総会決議事項', '株主総会での決議が必要な事項', '会社法および定款で定められた株主総会決議事項', '定款変更,取締役選任,剰余金配当,資本政策', '高']
  ];
  
  sheet.getRange(1, 1, 1, 5).setValues([headers]);
  sheet.getRange(2, 1, sampleData.length, 5).setValues(sampleData);
  
  // ヘッダーのフォーマット
  sheet.getRange(1, 1, 1, 5)
    .setBackground('#FBBC04')
    .setFontColor('#000000')
    .setFontWeight('bold');
  
  // 列幅の調整
  sheet.setColumnWidth(1, 180);
  sheet.setColumnWidth(2, 300);
  sheet.setColumnWidth(3, 400);
  sheet.setColumnWidth(4, 350);
  sheet.setColumnWidth(5, 100);
  
  // データ検証の設定（重要度列）
  const validationRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['高', '中', '低'], true)
    .setAllowInvalid(false)
    .build();
  
  sheet.getRange(2, 5, 100, 1).setDataValidation(validationRule);
  
  sheet.setFrozenRows(1);
}

// ========= Checklists シート =========
function createChecklistsSheet(ss) {
  const sheet = ss.insertSheet('Checklists');
  
  const headers = ['カテゴリ', 'チェック項目', '説明', '必須/任意', 'テンプレート文'];
  
  const sampleData = [
    ['取締役会決議事項', '決議事項の明確化', '議案の内容を具体的に記載', '必須', '第○号議案：○○について'],
    ['取締役会決議事項', '出席者の確認', '定足数の充足を確認', '必須', '出席取締役○名（定足数○名）'],
    ['取締役会決議事項', '決議結果', '賛成・反対・棄権の記録', '必須', '全員一致で承認/賛成○名、反対○名'],
    ['開示事項', '開示時期', '適時開示のタイミング', '必須', '決議後速やかに開示'],
    ['開示事項', '開示内容', '開示する情報の範囲', '必須', '決議内容、理由、今後の見通し'],
    ['監査等委員会決議事項', '監査計画', '年度監査計画の策定', '必須', '○○年度監査計画について'],
    ['株主総会決議事項', '招集通知', '株主総会招集通知の発送', '必須', '総会の○週間前までに発送']
  ];
  
  sheet.getRange(1, 1, 1, 5).setValues([headers]);
  sheet.getRange(2, 1, sampleData.length, 5).setValues(sampleData);
  
  // ヘッダーのフォーマット
  sheet.getRange(1, 1, 1, 5)
    .setBackground('#9333EA')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold');
  
  // 列幅の調整
  sheet.setColumnWidth(1, 180);
  sheet.setColumnWidth(2, 250);
  sheet.setColumnWidth(3, 350);
  sheet.setColumnWidth(4, 100);
  sheet.setColumnWidth(5, 300);
  
  // データ検証の設定（必須/任意列）
  const validationRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['必須', '任意'], true)
    .setAllowInvalid(false)
    .build();
  
  sheet.getRange(2, 4, 100, 1).setDataValidation(validationRule);
  
  sheet.setFrozenRows(1);
}

// ========= Templates シート =========
function createTemplatesSheet(ss) {
  const sheet = ss.insertSheet('Templates');
  
  const headers = ['テンプレート名', 'カテゴリ', 'Google Doc ID', 'プレースホルダー', '最終更新日'];
  
  const sampleData = [
    ['取締役会議事録', '取締役会決議事項', '', '{{company_name}}, {{meeting_date}}, {{agenda_sections}}', ''],
    ['監査等委員会議事録', '監査等委員会決議事項', '', '{{company_name}}, {{meeting_date}}, {{audit_items}}', ''],
    ['適時開示文書', '開示事項', '', '{{company_name}}, {{disclosure_date}}, {{disclosure_content}}', ''],
    ['株主総会議事録', '株主総会決議事項', '', '{{company_name}}, {{meeting_date}}, {{resolutions}}', '']
  ];
  
  sheet.getRange(1, 1, 1, 5).setValues([headers]);
  sheet.getRange(2, 1, sampleData.length, 5).setValues(sampleData);
  
  // ヘッダーのフォーマット
  sheet.getRange(1, 1, 1, 5)
    .setBackground('#00ACC1')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold');
  
  // 列幅の調整
  sheet.setColumnWidth(1, 200);
  sheet.setColumnWidth(2, 200);
  sheet.setColumnWidth(3, 300);
  sheet.setColumnWidth(4, 400);
  sheet.setColumnWidth(5, 150);
  
  sheet.setFrozenRows(1);
}

// ========= Drafts シート =========
function createDraftsSheet(ss) {
  const sheet = ss.insertSheet('Drafts');
  
  const headers = [
    'message_id',
    'category',
    'doc_url',
    'created_at',
    'created_by',
    'status',
    'reviewed_by',
    'reviewed_at',
    'notes'
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // ヘッダーのフォーマット
  sheet.getRange(1, 1, 1, headers.length)
    .setBackground('#FF6D00')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold');
  
  // 列幅の調整
  sheet.setColumnWidth(1, 200); // message_id
  sheet.setColumnWidth(2, 150); // category
  sheet.setColumnWidth(3, 350); // doc_url
  sheet.setColumnWidth(4, 150); // created_at
  sheet.setColumnWidth(5, 150); // created_by
  sheet.setColumnWidth(6, 100); // status
  sheet.setColumnWidth(7, 150); // reviewed_by
  sheet.setColumnWidth(8, 150); // reviewed_at
  sheet.setColumnWidth(9, 300); // notes
  
  // データ検証の設定（status列）
  const validationRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['下書き', 'レビュー中', '承認済み', '却下'], true)
    .setAllowInvalid(false)
    .build();
  
  sheet.getRange(2, 6, 100, 1).setDataValidation(validationRule);
  
  sheet.setFrozenRows(1);
}

// ========= Logs シート =========
function createLogsSheet(ss) {
  const sheet = ss.insertSheet('Logs');
  
  const headers = ['timestamp', 'level', 'message', 'details'];
  
  sheet.getRange(1, 1, 1, 4).setValues([headers]);
  
  // ヘッダーのフォーマット
  sheet.getRange(1, 1, 1, 4)
    .setBackground('#616161')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold');
  
  // 列幅の調整
  sheet.setColumnWidth(1, 180);
  sheet.setColumnWidth(2, 80);
  sheet.setColumnWidth(3, 500);
  sheet.setColumnWidth(4, 300);
  
  // 条件付き書式（エラーレベルを赤色に）
  const range = sheet.getRange(2, 2, 1000, 1);
  const rule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('ERROR')
    .setBackground('#FFCDD2')
    .setFontColor('#B71C1C')
    .setRanges([range])
    .build();
  
  const rules = sheet.getConditionalFormatRules();
  rules.push(rule);
  sheet.setConditionalFormatRules(rules);
  
  sheet.setFrozenRows(1);
}

// ========= サンプルデータ挿入 =========
function insertSampleData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Messagesシートにサンプルデータを挿入
  const messagesSheet = ss.getSheetByName('Messages');
  if (messagesSheet) {
    const sampleMessages = [
      [
        'C12345678:1234567890.123456',
        'C12345678',
        '1234567890.123456',
        '',
        '来週の取締役会で新規事業の投資案件について議論したいと思います。投資額は約5億円を予定しています。',
        '山田太郎',
        '{"summary":"新規事業投資案件の取締役会審議","decisions":[],"action_items":[{"owner":"山田","task":"投資案件資料準備","due":"来週"}]}',
        '[{"category":"取締役会決議事項","score":0.85,"rationale":"5億円の投資案件は重要な財産の処分に該当"}]',
        true,
        '',
        'https://example.slack.com/archives/C12345678/p1234567890123456',
        '',
        '',
        '',
        new Date(),
        new Date()
      ],
      [
        'C12345678:1234567890.234567',
        'C12345678',
        '1234567890.234567',
        '1234567890.123456',
        '投資案件の詳細資料を共有します。ROIは3年で150%を見込んでいます。',
        '鈴木花子',
        '{"summary":"投資案件詳細とROI見込み","decisions":[],"action_items":[]}',
        '[{"category":"取締役会決議事項","score":0.75,"rationale":"投資案件の詳細情報"}]',
        true,
        '',
        'https://example.slack.com/archives/C12345678/p1234567890234567',
        '',
        '',
        '',
        new Date(),
        new Date()
      ]
    ];
    
    const lastRow = messagesSheet.getLastRow();
    messagesSheet.getRange(lastRow + 1, 1, sampleMessages.length, sampleMessages[0].length)
      .setValues(sampleMessages);
  }
  
  SpreadsheetApp.getUi().alert('サンプルデータを挿入しました');
}