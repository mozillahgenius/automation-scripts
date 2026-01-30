// ==========================================
// 議事録作成・承認ワークフロー システム
// 統合版 Google Apps Script
// ==========================================

// === セットアップ機能 ===
function setup() {
  const ui = SpreadsheetApp.getUi();

  try {
    // 現在のスプレッドシートでセットアップ
    const response = ui.alert(
      'セットアップ',
      '議事録管理システムをセットアップします。\n現在のスプレッドシートに必要なシートを作成します。\n続行しますか？',
      ui.ButtonSet.YES_NO
    );

    if (response !== ui.Button.YES) {
      return;
    }

    // 現在のスプレッドシートを使用
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const ssId = spreadsheet.getId();
    const url = spreadsheet.getUrl();

    // スプレッドシート名を変更（必要に応じて）
    spreadsheet.rename('議事録管理システム');

    // プロパティストアに保存
    PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', ssId);

    // シート初期化
    initializeSheets(spreadsheet);

    // プロパティの初期設定
    setupProperties();

    // トリガー設定
    setupTriggers();

    ui.alert(
      'セットアップ完了',
      `議事録管理システムのセットアップが完了しました。\n\n現在のスプレッドシートに必要なシートが作成されました。\n\nメニューから各機能をご利用ください。`,
      ui.ButtonSet.OK
    );

  } catch (error) {
    ui.alert('エラー', 'セットアップ中にエラーが発生しました: ' + error.toString(), ui.ButtonSet.OK);
  }
}

// === プロパティ設定 ===
function setupProperties() {
  const properties = PropertiesService.getScriptProperties();

  // デフォルト値を設定（既存の値がある場合は上書きしない）
  const defaults = {
    'FROM_NAME': '議事録管理システム',
    'REPLY_TO': Session.getActiveUser().getEmail(),
    'DOMAIN_RESTRICTION': 'false',
    'AI_PROMPT': '議案を標準フォーマットに整形してください。',
    'AUDIT_LOG_RETENTION_DAYS': '365',
    'APPROVAL_CHECK_DAYS': '7', // 承認確認期間（会議日からの日数）
    'APPROVAL_CONTACT_EMAIL': Session.getActiveUser().getEmail(), // 承認確認通知先
    'OPENAI_API_KEY': '', // OpenAI APIキー（ユーザーが設定）
    'USE_OPENAI': 'false', // OpenAI API使用フラグ
    'OPENAI_MODEL': 'gpt-5', // 使用するOpenAIモデル（固定）
    'DOCS_FOLDER_ID': '' // ドキュメント保存先フォルダID（空の場合はルートフォルダ）
  };

  for (const [key, value] of Object.entries(defaults)) {
    if (!properties.getProperty(key)) {
      properties.setProperty(key, value);
    }
  }
}

// === トリガー設定 ===
function setupTriggers() {
  // 既存のトリガーを削除
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    const funcName = trigger.getHandlerFunction();
    if (funcName === 'processReminders' || funcName === 'checkApprovalStatus') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // 承認確認用トリガーを設定（毎日午前9時）
  ScriptApp.newTrigger('checkApprovalStatus')
    .timeBased()
    .everyDays(1)
    .atHour(9)
    .create();
}

// === 定数定義 ===
const CONFIG = {
  SHEET_NAMES: {
    DOCS: 'Docs',
    MOTIONS: 'Motions',
    TEMPLATES: 'Templates',
    CONFIG: 'Config',
    AUDIT_LOG: 'AuditLog',
    OFFICERS: 'Officers'  // 役員マスタ
  },
  STATUS: {
    DRAFT: '編集中',
    APPROVED: '承認済',
    FINALIZED: '最終版'
  }
};

// === スプレッドシート取得 ===
function getSpreadsheet() {
  const ssId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (!ssId) {
    // 現在のスプレッドシートを使用
    return SpreadsheetApp.getActiveSpreadsheet();
  }
  try {
    return SpreadsheetApp.openById(ssId);
  } catch (e) {
    // IDが無効な場合は現在のスプレッドシートを使用
    return SpreadsheetApp.getActiveSpreadsheet();
  }
}

// === 初期化処理 ===
function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu('議事録ワークフロー')
    .addItem('🚀 初期セットアップ', 'setup')
    .addSeparator()
    .addItem('📝 新規ドラフト作成', 'showCreateDraftDialog')
    .addItem('➕ 決議事項追加', 'showAddMotionDialog')
    .addItem('📢 報告事項追加', 'showAddReportDialog')
    .addItem('✍️ 会議終了後 議事録追記', 'showPostMeetingEditorDialog')
    .addSeparator()
    .addItem('📄 テンプレート管理', 'showTemplateDialog')
    .addItem('🧪 テンプレートテスト', 'showTemplateTestDialog')
    .addItem('📑 テンプレート適用', 'applyTemplate')
    .addSeparator()
    .addItem('👥 役員管理', 'showOfficersManagementDialog')
    .addSeparator()
    .addItem('✅ ステータス変更', 'showStatusChangeDialog')
    .addSeparator()
    .addItem('⚙️ システム設定', 'showConfigDialog')
    .addItem('📋 監査ログ閲覧', 'showAuditLogDialog')
    .addSeparator()
    .addItem('🔄 シート初期化', 'initializeSheetsMenu')
    .addToUi();
}

// === シート初期化（メニュー用） ===
function initializeSheetsMenu() {
  const ss = getSpreadsheet();
  initializeSheets(ss);
}

// === シート初期化 ===
function initializeSheets(spreadsheet) {
  const ss = spreadsheet || getSpreadsheet();

  try {
    // Docsシートの作成・初期化
    let sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.DOCS);
    if (!sheet) {
      sheet = ss.insertSheet(CONFIG.SHEET_NAMES.DOCS);
      sheet.getRange(1, 1, 1, 10).setValues([[
        'docId', '会議種別', '会議日', 'タイトル', '下書きURL',
        '申請者', '期限', 'ステータス', '最終更新', 'バージョン'
      ]]);
      sheet.getRange(1, 1, 1, 10).setFontWeight('bold');
      sheet.setFrozenRows(1);
    }

    // Motionsシートの作成・初期化
    sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.MOTIONS);
    if (!sheet) {
      sheet = ss.insertSheet(CONFIG.SHEET_NAMES.MOTIONS);
      sheet.getRange(1, 1, 1, 15).setValues([[
        'docId', '種別', '議案番号', '議案タイトル', '入力HTML', '生成本文',
        '添付資料有無', '添付資料メモ', '生成時刻', '最終編集者',
        '説明者', '決議結果', '賛否詳細', '特別利害関係人', '付帯条件'
      ]]);
      sheet.getRange(1, 1, 1, 15).setFontWeight('bold');
      sheet.setFrozenRows(1);
    }

    // Templatesシートの作成・初期化
    sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.TEMPLATES);
    if (!sheet) {
      sheet = ss.insertSheet(CONFIG.SHEET_NAMES.TEMPLATES);
      sheet.getRange(1, 1, 1, 5).setValues([[
        '会議種別', 'templateDocId', 'テンプレ名称', 'バージョン', '備考'
      ]]);
      sheet.getRange(1, 1, 1, 5).setFontWeight('bold');
      sheet.setFrozenRows(1);

      // サンプルテンプレート情報を追加
      sheet.getRange(2, 1, 3, 5).setValues([
        ['監査等委員会', '', '監査等委員会議事録テンプレート', '1.0', '標準フォーマット'],
        ['取締役会', '', '取締役会議事録テンプレート', '1.0', '標準フォーマット'],
        ['取締役会(書面)', '', '取締役会書面決議議事録テンプレート', '1.0', '書面決議用']
      ]);
    }

    // Configシートの作成・初期化
    sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.CONFIG);
    if (!sheet) {
      sheet = ss.insertSheet(CONFIG.SHEET_NAMES.CONFIG);
      sheet.getRange(1, 1, 2, 2).setValues([
        ['設定項目', '設定値'],
        ['システム名', '議事録管理システム']
      ]);
      sheet.getRange(3, 1, 6, 2).setValues([
        ['会社名', '株式会社○○'],
        ['管理者メール', Session.getActiveUser().getEmail()],
        ['ドメイン制限', 'false'],
        ['AI整形プロンプト', '議案を標準フォーマットに整形してください。'],
        ['承認確認期間(日)', '7'],
        ['承認確認通知先', Session.getActiveUser().getEmail()]
      ]);
      sheet.getRange(1, 1, 1, 2).setFontWeight('bold');
    }

    // AuditLogシートの作成・初期化
    sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.AUDIT_LOG);
    if (!sheet) {
      sheet = ss.insertSheet(CONFIG.SHEET_NAMES.AUDIT_LOG);
      sheet.getRange(1, 1, 1, 6).setValues([[
        '日時', '操作種別', '実行者', 'docId', '詳細', '成否'
      ]]);
      sheet.getRange(1, 1, 1, 6).setFontWeight('bold');
      sheet.setFrozenRows(1);
    }

    // Officersシートの作成・初期化
    sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.OFFICERS);
    if (!sheet) {
      sheet = ss.insertSheet(CONFIG.SHEET_NAMES.OFFICERS);
      sheet.getRange(1, 1, 1, 6).setValues([[
        '氏名', '役職', 'メールアドレス', '役員区分', '在任開始日', '備考'
      ]]);
      sheet.getRange(1, 1, 1, 6).setFontWeight('bold');
      sheet.setFrozenRows(1);

      // サンプル役員データ
      sheet.getRange(2, 1, 3, 6).setValues([
        ['山田太郎', '代表取締役社長', 'yamada@example.com', '取締役', '2020/04/01', ''],
        ['鈴木花子', '取締役CFO', 'suzuki@example.com', '取締役', '2021/04/01', ''],
        ['佐藤次郎', '監査役', 'sato@example.com', '監査役', '2019/04/01', '常勤']
      ]);
    }

    // デフォルトシートを削除
    const defaultSheet = ss.getSheetByName('Sheet1') || ss.getSheetByName('シート1');
    if (defaultSheet && ss.getSheets().length > 1) {
      ss.deleteSheet(defaultSheet);
    }

    // 初期化完了ログ
    addAuditLog('INITIALIZE', null, 'シート初期化完了', true);

    SpreadsheetApp.getUi().alert('初期化完了', 'すべてのシートが初期化されました。', SpreadsheetApp.getUi().ButtonSet.OK);

  } catch (error) {
    console.error('シート初期化エラー:', error);
    SpreadsheetApp.getUi().alert('エラー', 'シート初期化中にエラーが発生しました: ' + error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
    addAuditLog('INITIALIZE_ERROR', null, error.toString(), false);
  }
}

// ==========================================
// 設定管理機能
// ==========================================

// === Config値取得 ===
function getConfigValue(key) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.CONFIG);
    if (!sheet) return null;

    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === key) {
        return data[i][1];
      }
    }
    return null;
  } catch (error) {
    console.error('Config値取得エラー:', error);
    return null;
  }
}

// === Config値設定 ===
function setConfigValue(key, value) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.CONFIG);
    if (!sheet) return false;

    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === key) {
        sheet.getRange(i + 1, 2).setValue(value);
        return true;
      }
    }

    // 新規追加
    const lastRow = sheet.getLastRow();
    sheet.getRange(lastRow + 1, 1, 1, 2).setValues([[key, value]]);
    return true;
  } catch (error) {
    console.error('Config値設定エラー:', error);
    return false;
  }
}

// === 利用可能なプレースホルダー一覧取得 ===
function getAvailablePlaceholders() {
  return {
    basic: [
      { key: '{{COMPANY_NAME}}', description: '会社名' },
      { key: '{{MEETING_TYPE}}', description: '会議種別（取締役会、株主総会等）' },
      { key: '{{MEETING_TITLE}}', description: '会議タイトル（第○回○○会等）' },
      { key: '{{MEETING_DATE}}', description: '会議日（フォーマット済み）' },
      { key: '{{YEAR}}', description: '開催年' },
      { key: '{{MONTH}}', description: '開催月' },
      { key: '{{DAY}}', description: '開催日' },
      { key: '{{HOUR}}', description: '開始時間（時）' },
      { key: '{{MINUTE}}', description: '開始時間（分）' },
      { key: '{{LOCATION}}', description: '開催場所' },
      { key: '{{CHAIR}}', description: '議長名' }
    ],
    officers: [
      { key: '{{ATTENDEES}}', description: '出席者リスト' },
      { key: '{{ABSENTEES}}', description: '欠席者リスト' },
      { key: '{{ATTENDING_OFFICERS}}', description: '出席役員リスト' },
      { key: '{{ABSENT_OFFICERS}}', description: '欠席役員リスト' },
      { key: '{{SECRETARY}}', description: '議事録作成者' }
    ],
    content: [
      { key: '{{RESOLUTIONS_BLOCK}}', description: '決議事項ブロック（決議事項が挿入される場所）' },
      { key: '{{REPORTS_BLOCK}}', description: '報告事項ブロック（報告事項が挿入される場所）' },
      { key: '{{RESOLUTION_RESULT}}', description: '決議結果' },
      { key: '{{NEXT_MEETING}}', description: '次回会議予定' }
    ]
  };
}

// ==========================================
// ドキュメント管理機能
// ==========================================

// === ドラフト作成 ===
function createDraft(params) {
  try {
    const {
      meetingType,
      meetingDate,
      title,
      location,
      chair,
      attendees,
      absentees,
      approvers,
      deadline
    } = params;

    // テンプレート取得
    const templateId = getTemplateId(meetingType);
    if (!templateId) {
      throw new Error(`${meetingType}のテンプレートが設定されていません`);
    }

    // テンプレートからドキュメントをコピー
    const template = DriveApp.getFileById(templateId);
    const docTitle = `${meetingType}_${meetingDate}_${title}`;

    // 保存先フォルダを取得
    const folderId = PropertiesService.getScriptProperties().getProperty('DOCS_FOLDER_ID');
    let newDoc;
    if (folderId) {
      const folder = DriveApp.getFolderById(folderId);
      newDoc = template.makeCopy(docTitle, folder);
    } else {
      newDoc = template.makeCopy(docTitle);
    }
    const docId = newDoc.getId();

    // ドキュメントを開いてプレースホルダを置換
    const doc = DocumentApp.openById(docId);
    const body = doc.getBody();

    // 出席者・欠席者を氏名 役職形式に変換（半角スペース区切り）
    const formatAttendees = (list) => {
      if (!list || list.length === 0) return '';
      return list.map(person => {
        if (typeof person === 'string') {
          // 旧形式（文字列）の場合
          return person;
        } else if (person.name) {
          // 新形式（オブジェクト）の場合
          return person.position ? `${person.name} ${person.position}` : person.name;
        }
        return '';
      }).filter(p => p).join('、');
    };

    // 署名欄用に役員を改行とスペースで整形
    const formatSignatureOfficers = (list, chairName) => {
      if (!list || list.length === 0) return '';

      // 議長の情報を出席者リストから探す
      let signatureText = '';
      let chairPosition = '';
      let chairFullName = '';

      if (chairName) {
        for (const person of list) {
          let name = '';
          let position = '';

          if (typeof person === 'string') {
            const parts = person.trim().split(/\s+/);
            if (parts.length > 1) {
              name = parts[0];
              position = parts.slice(1).join(' ');
            } else {
              name = person;
            }
          } else if (person.name) {
            name = person.name;
            position = person.position || '';
          }

          // 議長名と一致する場合、その役職を保存
          if (name === chairName) {
            chairFullName = name;
            chairPosition = position || '代表取締役社長';
            break;
          }
        }

        // 議長を最初に配置（役職付き）
        if (chairPosition) {
          signatureText += `議長 ${chairPosition}    ${chairFullName}\n\n\n\n\n`;
        } else {
          signatureText += `議長 代表取締役    ${chairName}\n\n\n\n\n`;
        }
      }

      // その他の出席役員
      const officers = list.map(person => {
        let name = '';
        let position = '';

        if (typeof person === 'string') {
          // 旧形式（文字列）の場合、スペースで分割
          const parts = person.trim().split(/\s+/);
          if (parts.length > 1) {
            name = parts[0];
            position = parts.slice(1).join(' ');
          } else {
            name = person;
          }
        } else if (person.name) {
          // 新形式（オブジェクト）の場合
          name = person.name;
          position = person.position || '';
        }

        // 議長と同じ名前の場合はスキップ（重複を避ける）
        if (chairName && (name === chairName || name === chairFullName)) {
          return '';
        }

        // 役職がある場合は先に役職、なければ「取締役」と仮定
        const title = position || '取締役';

        // 各役員の署名欄（5行分の改行でスペースを確保）
        return `${title}    ${name}\n\n\n\n\n`;
      }).filter(p => p).join('');

      signatureText += officers;
      return signatureText;
    };


    // 日付情報を解析
    const meetingDateObj = new Date(meetingDate);
    const year = meetingDateObj.getFullYear();
    const month = meetingDateObj.getMonth() + 1;
    const day = meetingDateObj.getDate();
    const hour = meetingDateObj.getHours();
    const minute = meetingDateObj.getMinutes();

    // 会社情報を取得（Configシートから）
    const companyName = getConfigValue('会社名') || '株式会社○○';

    // プレースホルダ置換
    // 注意: {{RESOLUTIONS_BLOCK}}と{{REPORTS_BLOCK}}は議案追加時に使用するため、ここでは置換しない
    replacePlaceholders(body, {
      '{{COMPANY_NAME}}': companyName,
      '{{MEETING_TYPE}}': meetingType,
      '{{MEETING_TITLE}}': title || meetingType,  // タイトルまたは会議体名
      '{{MEETING_DATE}}': meetingDate,
      '{{YEAR}}': year.toString(),
      '{{MONTH}}': month.toString(),
      '{{DAY}}': day.toString(),
      '{{HOUR}}': hour.toString().padStart(2, '0'),
      '{{MINUTE}}': minute.toString().padStart(2, '0'),
      '{{LOCATION}}': location || '本社会議室',
      '{{CHAIR}}': chair || '',
      '{{ATTENDEES}}': formatAttendees(attendees),
      '{{ABSENTEES}}': formatAttendees(absentees),
      '{{ATTENDING_OFFICERS}}': formatAttendees(attendees),  // 出席役員
      '{{ABSENT_OFFICERS}}': formatAttendees(absentees),     // 欠席役員
      '{{SECRETARY}}': Session.getActiveUser().getEmail(),
      '{{CREATED_DATE}}': formatDate(new Date()),
      '{{TITLE}}': title
      // {{RESOLUTIONS_BLOCK}}, {{REPORTS_BLOCK}}, {{RESOLUTION_RESULT}}, {{NEXT_MEETING}} は残す
      // {{APPROVALS_TABLE}} は既にテンプレートから削除済み
    });

    doc.saveAndClose();

    // スプレッドシートに記録
    const ss = getSpreadsheet();
    const docsSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.DOCS);

    docsSheet.appendRow([
      docId,
      meetingType,
      meetingDate,
      title,
      doc.getUrl(),
      Session.getActiveUser().getEmail(),
      deadline,
      CONFIG.STATUS.DRAFT,
      new Date(),
      '1.0'
    ]);

    // 監査ログ
    addAuditLog('CREATE_DRAFT', docId, `ドラフト作成: ${docTitle}`, true);

    return {
      success: true,
      docId: docId,
      url: doc.getUrl(),
      message: 'ドラフトを作成しました'
    };

  } catch (error) {
    console.error('ドラフト作成エラー:', error);
    addAuditLog('CREATE_DRAFT_ERROR', null, error.toString(), false);
    return {
      success: false,
      error: error.toString()
    };
  }
}

// === テンプレートID取得 ===
function getTemplateId(meetingType) {
  const ss = getSpreadsheet();
  const templatesSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.TEMPLATES);

  if (!templatesSheet) {
    throw new Error('Templatesシートが見つかりません');
  }

  const data = templatesSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === meetingType) {
      return data[i][1];
    }
  }

  return null;
}

// === プレースホルダ置換 ===
function replacePlaceholders(body, replacements) {
  for (const [placeholder, value] of Object.entries(replacements)) {
    // 空のプレースホルダーはスキップ
    if (!placeholder || placeholder.trim() === '') {
      continue;
    }
    // 正規表現の特殊文字をエスケープ
    const escapedPlaceholder = placeholder.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    body.replaceText(escapedPlaceholder, value || '');
  }
}

// === プレースホルダー抽出 ===
function extractPlaceholders(text) {
  const regex = /\{\{([^}]+)\}\}/g;
  const matches = text.match(regex);
  if (matches) {
    const unique = [...new Set(matches)];
    return unique.join(', ');
  }
  return '';
}

// === ドキュメントから議長情報を取得 ===
function getChairFromDocument(docId) {
  try {
    const doc = DocumentApp.openById(docId);
    const text = doc.getBody().getText();

    // 議長パターンを検索
    const patterns = [
      /議長[:：]\s*(.+?)[\s\n]/,
      /議長\s+(.+?)[\s\n]/,
      /議長たる(.+?)[\s\n]/
    ];

    for (const pattern of patterns) {
      const match = text.match(pattern);
      if (match && match[1]) {
        return match[1].trim();
      }
    }

    // Docsシートから取得を試みる
    const ss = getSpreadsheet();
    const docsSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.DOCS);
    if (docsSheet) {
      const data = docsSheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] === docId && data[i][10]) { // 議長列があれば
          return data[i][10];
        }
      }
    }

    return null;
  } catch (error) {
    console.error('getChairFromDocument error:', error);
    return null;
  }
}

// === 議長と説明者が同じかチェック ===
function isChairAndPresenterSame(chair, presenter) {
  // 氏名部分だけを比較（役職を除く）
  const chairName = chair.split(/[\s　]/)[0];
  const presenterName = presenter.split(/[\s　]/)[0];

  return chairName && presenterName && chairName === presenterName;
}

// === 議案ブロック更新 ===

// === 議案セクション処理（マーカーベース） ===
function processMotionSection(body, paragraphs, motions, motionsContent, startMarker, endMarker, placeholder, typeName, docId) {
  let startIndex = -1;
  let endIndex = -1;
  let placeholderIndex = -1;

  // マーカーまたはプレースホルダーを探す（最新のparagraphs配列を取得し直す）
  const currentParagraphs = body.getParagraphs();
  for (let i = 0; i < currentParagraphs.length; i++) {
    const text = currentParagraphs[i].getText();

    if (text.includes(startMarker)) {
      startIndex = i;
    } else if (text.includes(endMarker)) {
      endIndex = i;
    } else if (text.includes(placeholder)) {
      placeholderIndex = i;
    }
  }

  // 既存のマーカーで囲まれたセクションがある場合
  if (startIndex >= 0 && endIndex >= 0 && startIndex < endIndex) {
    // 既存の議案番号を収集
    const existingNumbers = new Set();
    for (let i = startIndex + 1; i < endIndex; i++) {
      const text = currentParagraphs[i].getText();
      let match;
      if (typeName === '決議事項') {
        match = text.match(/第(\d+)号議案/);
      } else {
        match = text.match(/【報告事項(\d+)】/);
      }
      if (match) {
        existingNumbers.add(parseInt(match[1]));
      }
    }

    // 新しい議案だけをフィルタ
    const newMotions = motions.filter(m => {
      const num = parseInt(String(m.motionNumber).replace(/[^0-9]/g, ''));
      return !existingNumbers.has(num);
    });

    if (newMotions.length > 0) {
      // 新しい議案のテキストを生成
      let newMotionsContent = '';
      const hasExisting = existingNumbers.size > 0;

      newMotions.sort((a, b) => {
        const numA = parseInt(String(a.motionNumber).replace(/[^0-9]/g, ''));
        const numB = parseInt(String(b.motionNumber).replace(/[^0-9]/g, ''));
        return numA - numB;
      });

      newMotions.forEach((motion, index) => {
        // 既存議案がある場合は区切りを追加
        if (hasExisting || index > 0) {
          newMotionsContent += '\n\n';
        }

        // 議案内容を構築
        if (typeName === '決議事項') {
          newMotionsContent += `第${motion.motionNumber}号議案　${motion.title}\n\n`;

          if (motion.presenter) {
            const chairInfo = getChairFromDocument(docId);
            if (chairInfo && isChairAndPresenterSame(chairInfo, motion.presenter)) {
              newMotionsContent += `議長たる${motion.presenter}から次の議案について説明がなされた。\n\n`;
            } else {
              newMotionsContent += `議長の指名により、${motion.presenter}から次の議案について説明がなされた。\n\n`;
            }
          }

          newMotionsContent += `${motion.content}\n`;

          if (motion.resolutionResult) {
            newMotionsContent += '\n';
            if (motion.resolutionResult === '全員一致承認' || motion.resolutionResult === '全員一致で承認') {
              newMotionsContent += '議長が本議案を議場に諮ったところ、出席取締役全員異議なく、これを承認可決した。';
            } else if (motion.resolutionResult === '賛成多数承認' || motion.resolutionResult === '賛成多数で承認') {
              newMotionsContent += '議長が本議案を議場に諮ったところ、賛成多数をもってこれを承認可決した';
              if (motion.votingDetails) {
                newMotionsContent += `（${motion.votingDetails}）`;
              }
              newMotionsContent += '。';
            } else if (motion.resolutionResult === '継続審議') {
              newMotionsContent += '本議案については、さらに検討を要するため、継続審議とすることとした。';
            } else if (motion.resolutionResult === '否決') {
              newMotionsContent += '議長が本議案を議場に諮ったところ、賛成少数により否決された。';
            } else {
              newMotionsContent += `議長が本議案を議場に諮ったところ、${motion.resolutionResult}。`;
            }
          }

          if (motion.specialInterest && motion.specialInterest !== 'なし' && motion.specialInterest !== '') {
            newMotionsContent += `\nなお、${motion.specialInterest}は、本議案について特別の利害関係を有するため、その審議および決議に参加しなかった。`;
          }

          if (motion.conditions && motion.conditions !== 'なし' && motion.conditions !== '') {
            newMotionsContent += `\nまた、本議案の実施にあたっては、${motion.conditions}ことが付帯条件として決議された。`;
          }
        } else {
          // 報告事項
          newMotionsContent += `【報告事項${motion.motionNumber}】${motion.title}\n\n`;

          if (motion.presenter) {
            const chairInfo = getChairFromDocument(docId);
            if (chairInfo && isChairAndPresenterSame(chairInfo, motion.presenter)) {
              newMotionsContent += `議長たる${motion.presenter}から次の報告がなされた。\n\n`;
            } else {
              newMotionsContent += `議長の指名により、${motion.presenter}から次の報告がなされた。\n\n`;
            }
          }

          newMotionsContent += `${motion.content}\n`;

          if (motion.resolutionResult) {
            newMotionsContent += '\n';
            if (motion.resolutionResult === '報告事項') {
              newMotionsContent += '本件は報告事項として説明がなされ、出席取締役はこれを了承した。';
            } else {
              newMotionsContent += `${motion.resolutionResult}。`;
            }
          }
        }
      });

      // 終了マーカーの直前に新しい議案を挿入
      const insertIndex = body.getChildIndex(currentParagraphs[endIndex]);

      // 既存テキストのスタイルを取得
      let templateStyle = null;
      for (let i = startIndex + 1; i < endIndex; i++) {
        const para = currentParagraphs[i];
        if (para.getText().trim()) {
          const text = para.editAsText();
          templateStyle = {
            fontSize: text.getFontSize(0),
            fontFamily: text.getFontFamily(0),
            foregroundColor: text.getForegroundColor(0)
          };
          break;
        }
      }

      const lines = newMotionsContent.split('\n');
      lines.forEach((line, idx) => {
        const newPara = body.insertParagraph(insertIndex + idx, line);
        if (templateStyle && templateStyle.fontSize) {
          const text = newPara.editAsText();
          if (templateStyle.fontSize) text.setFontSize(templateStyle.fontSize);
          if (templateStyle.fontFamily) text.setFontFamily(templateStyle.fontFamily);
          if (templateStyle.foregroundColor) text.setForegroundColor(templateStyle.foregroundColor);
        }
      });
    }

  // プレースホルダーがある場合
  } else if (placeholderIndex >= 0) {
    // プレースホルダーをマーカーに置換
    currentParagraphs[placeholderIndex].setText(startMarker);

    if (motionsContent) {
      const insertIndex = body.getChildIndex(currentParagraphs[placeholderIndex]) + 1;
      const lines = motionsContent.split('\n');
      lines.forEach((line, idx) => {
        body.insertParagraph(insertIndex + idx, line);
      });
      body.insertParagraph(insertIndex + lines.length, endMarker);
    } else {
      body.insertParagraph(body.getChildIndex(currentParagraphs[placeholderIndex]) + 1, endMarker);
    }
  }
  // プレースホルダーもマーカーもない場合は何もしない（ドキュメントを汚さない）
}

// === ドキュメント最終化 ===
function finalizeDocument(docId) {
  try {
    const doc = DocumentApp.openById(docId);

    // PDFエクスポート
    const pdf = DriveApp.getFileById(docId).getAs('application/pdf');
    const pdfName = doc.getName() + '_最終版.pdf';
    const pdfFile = DriveApp.createFile(pdf);
    pdfFile.setName(pdfName);

    // ドキュメントと同じフォルダに保存
    const docFile = DriveApp.getFileById(docId);
    const parents = docFile.getParents();
    if (parents.hasNext()) {
      const folder = parents.next();
      folder.addFile(pdfFile);
      DriveApp.getRootFolder().removeFile(pdfFile);
    }

    // ステータス更新
    updateDocStatus(docId, CONFIG.STATUS.FINALIZED);

    // バージョン名設定
    const version = `最終版_${formatDate(new Date())}`;
    updateDocVersion(docId, version);

    addAuditLog('FINALIZE', docId, `文書最終化: ${pdfName}`, true);

    return {
      success: true,
      pdfUrl: pdfFile.getUrl(),
      message: '文書を最終化しました'
    };

  } catch (error) {
    console.error('文書最終化エラー:', error);
    addAuditLog('FINALIZE_ERROR', docId, error.toString(), false);
    return {
      success: false,
      error: error.toString()
    };
  }
}

// === ドキュメントステータス更新 ===
function updateDocStatus(docId, status) {
  const ss = getSpreadsheet();
  const docsSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.DOCS);

  const data = docsSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === docId) {
      docsSheet.getRange(i + 1, 8).setValue(status);
      docsSheet.getRange(i + 1, 9).setValue(new Date());
      break;
    }
  }
}

// === ドキュメントバージョン更新 ===
function updateDocVersion(docId, version) {
  const ss = getSpreadsheet();
  const docsSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.DOCS);

  const data = docsSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === docId) {
      if (!version) {
        // 自動バージョンアップ
        let currentVersion = data[i][9] || '1.0';
        // バージョンを文字列に変換
        if (typeof currentVersion !== 'string') {
          currentVersion = String(currentVersion || '1.0');
        }
        const versionParts = currentVersion.split('.');
        const minorVersion = parseInt(versionParts[1] || '0') + 1;
        version = `${versionParts[0]}.${minorVersion}`;
      }
      docsSheet.getRange(i + 1, 10).setValue(version);
      break;
    }
  }
}

// === ドキュメント最終更新日時のみ記録 ===
function updateDocumentLastModified(docId) {
  try {
    const ss = getSpreadsheet();
    const docsSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.DOCS);

    if (!docsSheet) {
      console.error('Docsシートが見つかりません');
      return;
    }

    const data = docsSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === docId) {
        // 最終更新日時を更新（9列目）
        docsSheet.getRange(i + 1, 9).setValue(new Date());
        console.log('ドキュメント最終更新日時を記録:', docId);
        break;
      }
    }
  } catch (error) {
    console.error('updateDocumentLastModified error:', error);
  }
}

// ==========================================
// 議案管理機能
// ==========================================

// === テスト用：ドキュメント一覧取得テスト ===
function testGetDocsList() {
  console.log('===== testGetDocsList開始 =====');

  try {
    // 直接getDocsListを呼び出してテスト
    const docs = getDocsList();

    console.log('取得結果:');
    console.log('- ドキュメント数:', docs.length);
    console.log('- データ型:', typeof docs);
    console.log('- 配列かどうか:', Array.isArray(docs));

    if (docs.length > 0) {
      console.log('最初のドキュメント:', JSON.stringify(docs[0]));
    }

    return {
      success: true,
      count: docs.length,
      docs: docs,
      message: `${docs.length}件のドキュメントを取得しました`
    };

  } catch (error) {
    console.error('testGetDocsList エラー:', error);
    return {
      success: false,
      error: error.toString()
    };
  }
}

// === テスト用：文字列パラメータテスト ===
function testSimpleString(str) {
  console.log('testSimpleString received:', str);
  console.log('Type:', typeof str);
  console.log('Value:', str);
  return {
    success: true,
    received: str,
    type: typeof str,
    length: str ? str.length : 0
  };
}

// === テスト用：複数パラメータテスト ===
function testMultipleParams(param1, param2, param3, param4) {
  console.log('testMultipleParams received:');
  console.log('- param1:', param1, 'type:', typeof param1);
  console.log('- param2:', param2, 'type:', typeof param2);
  console.log('- param3:', param3, 'type:', typeof param3);
  console.log('- param4:', param4, 'type:', typeof param4);

  return {
    success: true,
    params: {
      param1: param1,
      param2: param2,
      param3: param3,
      param4: param4
    }
  };
}

// === 議案追加（拡張版） ===
// === 議案追加（個別パラメータで受け取る） ===
function addMotionDirect(docId, motionNumber, title, inputHtml, hasAttachment, attachmentMemo, useAI) {
  console.log('addMotionDirect called');
  console.log('- docId:', docId);
  console.log('- motionNumber:', motionNumber);
  console.log('- title:', title);
  console.log('- inputHtml:', inputHtml);
  console.log('- hasAttachment:', hasAttachment);
  console.log('- attachmentMemo:', attachmentMemo);
  console.log('- useAI:', useAI);

  // オブジェクトに再構成
  const params = {
    docId: docId,
    motionNumber: motionNumber,
    title: title,
    inputHtml: inputHtml,
    hasAttachment: hasAttachment,
    attachmentMemo: attachmentMemo,
    useAI: useAI
  };

  // 既存の関数を呼び出す
  return addMotion(params);
}

// === 会議終了後議案追加 ===
function addPostMeetingMotion(data) {
  try {
    console.log('addPostMeetingMotion called with:', JSON.stringify(data));

    if (!data || !data.docId || !data.motionNumber || !data.title) {
      throw new Error('必須パラメータが不足しています');
    }

    const {
      docId,
      motionNumber,
      title,
      content,
      attendingOfficers,
      absentOfficers,
      presenter,
      resolutionResult,
      votingDetails,
      specialInterest,
      conditions,
      type
    } = data;

    // ドキュメントを取得
    const doc = DocumentApp.openById(docId);
    const body = doc.getBody();

    // 議案番号の重複チェック
    if (isMotionNumberDuplicate(docId, motionNumber, type || '決議事項')) {
      // 既存の議案を更新する
      updateExistingMotion(docId, motionNumber, data);
    } else {
      // 新しい議案として追加
      const motionText = formatPostMeetingMotion(data);

      // ドキュメントに議案を追加
      insertMotionToDocument(doc, motionText, motionNumber);

      // 出席者情報を更新
      updateAttendanceInfo(doc, attendingOfficers, absentOfficers);
    }

    // 決議一覧表を更新
    updateApprovalTable(doc);

    // データベースに記録
    const motionSheet = getSpreadsheet().getSheetByName(CONFIG.SHEET_NAMES.MOTIONS);
    motionSheet.appendRow([
      docId,
      motionNumber,
      title,
      content,
      presenter,
      resolutionResult,
      votingDetails,
      specialInterest,
      conditions,
      new Date(),
      '会議終了後追記'
    ]);

    // ドキュメントの更新日時を記録
    updateDocumentLastModified(docId);

    // 監査ログ記録
    addAuditLog('POST_MEETING_MOTION_ADD', docId, {
      motionNumber: motionNumber,
      title: title,
      resolutionResult: resolutionResult
    });

    return {
      success: true,
      message: `議案${motionNumber}「${title}」を追記しました`
    };

  } catch (error) {
    console.error('addPostMeetingMotion error:', error);
    return {
      success: false,
      error: error.toString()
    };
  }
}

// === 会議終了後議案のフォーマット ===
function formatPostMeetingMotion(data) {
  let text = `第${data.motionNumber}号議案　${data.title}\n\n`;

  if (data.presenter) {
    text += `【提案者】\n${data.presenter}\n\n`;
  }

  text += `【議案内容】\n${data.content}\n\n`;

  text += `【決議結果】\n議長が議場に諮ったところ、${data.resolutionResult}した。\n`;
  if (data.votingDetails) {
    text += `（${data.votingDetails}）\n`;
  }
  text += '\n';

  if (data.specialInterest) {
    text += `【特別利害関係人】\n${data.specialInterest}\n\n`;
  }

  if (data.conditions) {
    text += `【付帯条件・留意事項】\n${data.conditions}\n\n`;
  }

  return text;
}

// === 出席者情報の更新 ===
function updateAttendanceInfo(doc, attendingOfficers, absentOfficers) {
  const body = doc.getBody();
  const text = body.getText();

  // 出席者のプレースホルダーを置換
  const attendingPlaceholder = '{{ATTENDING_OFFICERS}}';
  const absentPlaceholder = '{{ABSENT_OFFICERS}}';

  if (text.includes(attendingPlaceholder)) {
    body.replaceText(attendingPlaceholder, attendingOfficers || '');
  }

  if (text.includes(absentPlaceholder)) {
    body.replaceText(absentPlaceholder, absentOfficers || '');
  }
}

// === 既存議案の更新 ===
function updateExistingMotion(docId, motionNumber, data) {
  const doc = DocumentApp.openById(docId);
  const body = doc.getBody();

  // 既存の議案を探して更新
  const searchPattern = `第${motionNumber}号議案`;
  const foundElement = body.findText(searchPattern);

  if (foundElement) {
    // 既存の議案を更新ロジック
    // （実装の詳細は要件に応じて調整）
    console.log('既存議案を更新:', motionNumber);
  }
}

// === 次の議案番号を取得 ===
function getNextMotionNumber(docId) {
  try {
    const motions = getMotionsList(docId);
    if (!motions || motions.length === 0) {
      return 1;
    }

    let maxNumber = 0;
    motions.forEach(motion => {
      const num = parseInt(motion.motionNumber);
      if (num > maxNumber) {
        maxNumber = num;
      }
    });

    return maxNumber + 1;
  } catch (error) {
    console.error('getNextMotionNumber error:', error);
    return 1;
  }
}

// === 議案追加（JSON文字列で受け取る） ===
function addMotionJSON(jsonString) {
  try {
    console.log('addMotionJSON called with:', jsonString);
    console.log('Type of jsonString:', typeof jsonString);

    // JSON文字列をパース
    const params = JSON.parse(jsonString);

    console.log('Parsed params:', params);

    return addMotion(params);
  } catch (error) {
    console.error('JSONパースエラー:', error);
    return {
      success: false,
      error: 'パラメータの解析に失敗しました: ' + error.toString()
    };
  }
}

// === 議案追加（内部処理） ===
function addMotion(params) {
  try {
    console.log('addMotion called');
    console.log('Type of params:', typeof params);
    console.log('params:', params);
    console.log('params stringified:', JSON.stringify(params));
    console.log('params keys:', params ? Object.keys(params) : 'null');

    if (!params) {
      throw new Error('パラメータが指定されていません');
    }

    // デストラクチャリングを避けて個別に取得
    const docId = params.docId;
    const motionNumber = params.motionNumber;
    const title = params.title;
    const inputHtml = params.inputHtml;
    const hasAttachment = params.hasAttachment;
    const attachmentMemo = params.attachmentMemo;
    const useAI = params.useAI;
    const type = params.type || '決議事項'; // デフォルトは決議事項
    const presenter = params.presenter || '';
    const resolutionResult = params.resolutionResult || (type === '報告事項' ? '報告事項' : '');
    const votingDetails = params.votingDetails || '';
    const specialInterest = params.specialInterest || '';
    const conditions = params.conditions || '';

    console.log('Extracted parameters:');
    console.log('- docId:', docId);
    console.log('- motionNumber:', motionNumber);
    console.log('- title:', title);
    console.log('- inputHtml length:', inputHtml ? inputHtml.length : 'null');
    console.log('- hasAttachment:', hasAttachment);
    console.log('- attachmentMemo:', attachmentMemo);
    console.log('- useAI:', useAI);
    console.log('- type:', type);
    console.log('- presenter:', presenter);
    console.log('- resolutionResult:', resolutionResult);

    // 必須パラメータのチェック
    if (!docId) {
      throw new Error('ドキュメントIDが指定されていません');
    }
    if (!motionNumber) {
      throw new Error('議案番号が指定されていません');
    }
    if (!title) {
      throw new Error('議案タイトルが指定されていません');
    }
    if (!inputHtml) {
      throw new Error('議案内容が指定されていません');
    }

    // 議案番号の重複チェック
    if (isMotionNumberDuplicate(docId, motionNumber, type)) {
      const typeName = type === '報告事項' ? '報告事項' : '議案';
      throw new Error(`${typeName}番号${motionNumber}は既に使用されています`);
    }

    // AI整形処理
    let generatedText = inputHtml;
    if (useAI) {
      generatedText = formatMotionWithAI(inputHtml, title, type);
    }

    // スプレッドシートに保存
    const ss = getSpreadsheet();
    const motionsSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.MOTIONS);

    if (!motionsSheet) {
      throw new Error('Motionsシートが見つかりません');
    }

    motionsSheet.appendRow([
      docId,
      type,
      motionNumber,
      title,
      inputHtml,
      generatedText,
      hasAttachment ? '有' : '無',
      attachmentMemo || '',
      new Date(),
      Session.getActiveUser().getEmail(),
      presenter,
      resolutionResult,
      votingDetails,
      specialInterest,
      conditions
    ]);

    // ドキュメントに反映
    updateMotionsBlockExtended(docId);

    // 監査ログ
    addAuditLog('ADD_MOTION', docId, `議案追加: ${motionNumber} - ${title}`, true);

    return {
      success: true,
      message: `議案第${motionNumber}号を追加しました`,
      generatedText: generatedText
    };

  } catch (error) {
    console.error('議案追加エラー:', error);
    const docIdForLog = params?.docId || null;
    addAuditLog('ADD_MOTION_ERROR', docIdForLog, error.toString(), false);
    return {
      success: false,
      error: error.toString()
    };
  }
}

// === 拡張版議案ブロック更新 ===
function updateMotionsBlockExtended(docId) {
  try {
    const doc = DocumentApp.openById(docId);
    const body = doc.getBody();

    // 議案データ取得
    const ss = getSpreadsheet();
    const motionsSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.MOTIONS);

    if (!motionsSheet || motionsSheet.getLastRow() <= 1) {
      return;
    }

    const data = motionsSheet.getDataRange().getValues();

    // 該当ドキュメントの議案を抽出
    const motions = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === docId) {
        motions.push({
          type: data[i][1] || '決議事項',
          motionNumber: data[i][2],
          title: data[i][3],
          content: data[i][5] || data[i][4], // AI整形後のテキスト、なければ元のテキスト
          presenter: data[i][10] || '',
          resolutionResult: data[i][11] || (data[i][1] === '報告事項' ? '報告事項' : '全員一致承認'),
          votingDetails: data[i][12] || '',
          specialInterest: data[i][13] || '',
          conditions: data[i][14] || ''
        });
      }
    }

    // 決議事項と報告事項を分離
    const resolutions = motions.filter(m => m.type === '決議事項');
    const reports = motions.filter(m => m.type === '報告事項');

    // 決議事項ブロックを構築
    let resolutionsContent = '';
    if (resolutions.length > 0) {
      resolutions.sort((a, b) => {
        const numA = parseInt(String(a.motionNumber).replace(/[^0-9]/g, ''));
        const numB = parseInt(String(b.motionNumber).replace(/[^0-9]/g, ''));
        return numA - numB;
      });

      resolutions.forEach((motion, index) => {
        if (index > 0) {
          resolutionsContent += '\n\n';
        }

        resolutionsContent += `第${motion.motionNumber}号議案　${motion.title}\n\n`;

        // 説明者の記載（議長と同じかチェック）
        if (motion.presenter) {
          const chairInfo = getChairFromDocument(docId);
          if (chairInfo && isChairAndPresenterSame(chairInfo, motion.presenter)) {
            resolutionsContent += `議長たる${motion.presenter}から次の議案について説明がなされた。\n\n`;
          } else {
            resolutionsContent += `議長の指名により、${motion.presenter}から次の議案について説明がなされた。\n\n`;
          }
        }

        // 議案内容（AI整形済みまたは元のテキスト）
        resolutionsContent += `${motion.content}\n`;

        // 決議結果を自然な文章で記載
        if (motion.resolutionResult) {
          resolutionsContent += '\n';

          if (motion.resolutionResult === '全員一致承認' || motion.resolutionResult === '全員一致で承認') {
            resolutionsContent += '議長が本議案を議場に諮ったところ、出席取締役全員異議なく、これを承認可決した。';
          } else if (motion.resolutionResult === '賛成多数承認' || motion.resolutionResult === '賛成多数で承認') {
            resolutionsContent += '議長が本議案を議場に諮ったところ、賛成多数をもってこれを承認可決した';
            if (motion.votingDetails) {
              resolutionsContent += `（${motion.votingDetails}）`;
            }
            resolutionsContent += '。';
          } else if (motion.resolutionResult === '継続審議') {
            resolutionsContent += '本議案については、さらに検討を要するため、継続審議とすることとした。';
          } else if (motion.resolutionResult === '否決') {
            resolutionsContent += '議長が本議案を議場に諮ったところ、賛成少数により否決された。';
          } else {
            resolutionsContent += `議長が本議案を議場に諮ったところ、${motion.resolutionResult}。`;
          }
        }

        // 特別利害関係人がいる場合
        if (motion.specialInterest && motion.specialInterest !== 'なし' && motion.specialInterest !== '') {
          resolutionsContent += `\nなお、${motion.specialInterest}は、本議案について特別の利害関係を有するため、その審議および決議に参加しなかった。`;
        }

        // 付帯条件がある場合
        if (motion.conditions && motion.conditions !== 'なし' && motion.conditions !== '') {
          resolutionsContent += `\nまた、本議案の実施にあたっては、${motion.conditions}ことが付帯条件として決議された。`;
        }
      });
    }

    // 報告事項ブロックを構築
    let reportsContent = '';
    if (reports.length > 0) {
      reports.sort((a, b) => {
        const numA = parseInt(String(a.motionNumber).replace(/[^0-9]/g, ''));
        const numB = parseInt(String(b.motionNumber).replace(/[^0-9]/g, ''));
        return numA - numB;
      });

      reports.forEach((motion, index) => {
        if (index > 0) {
          reportsContent += '\n\n';
        }

        reportsContent += `【報告事項${motion.motionNumber}】${motion.title}\n\n`;

        // 説明者の記載（議長と同じかチェック）
        if (motion.presenter) {
          const chairInfo = getChairFromDocument(docId);
          if (chairInfo && isChairAndPresenterSame(chairInfo, motion.presenter)) {
            reportsContent += `議長たる${motion.presenter}から次の報告がなされた。\n\n`;
          } else {
            reportsContent += `議長の指名により、${motion.presenter}から次の報告がなされた。\n\n`;
          }
        }

        // 報告内容（AI整形済みまたは元のテキスト）
        reportsContent += `${motion.content}\n`;

        // 報告事項の結果
        if (motion.resolutionResult) {
          reportsContent += '\n';
          if (motion.resolutionResult === '報告事項') {
            reportsContent += '本件は報告事項として説明がなされ、出席取締役はこれを了承した。';
          } else {
            reportsContent += `${motion.resolutionResult}。`;
          }
        }
      });
    }

    // マーカーとプレースホルダーを定義
    const RESOLUTIONS_START = '【決議事項開始】';
    const RESOLUTIONS_END = '【決議事項終了】';
    const REPORTS_START = '【報告事項開始】';
    const REPORTS_END = '【報告事項終了】';
    const RESOLUTIONS_PLACEHOLDER = '{{RESOLUTIONS_BLOCK}}';
    const REPORTS_PLACEHOLDER = '{{REPORTS_BLOCK}}';

    const paragraphs = body.getParagraphs();

    // 決議事項セクションを処理
    processMotionSection(body, paragraphs, resolutions, resolutionsContent,
                        RESOLUTIONS_START, RESOLUTIONS_END, RESOLUTIONS_PLACEHOLDER, '決議事項', docId);

    // 報告事項セクションを処理
    processMotionSection(body, paragraphs, reports, reportsContent,
                        REPORTS_START, REPORTS_END, REPORTS_PLACEHOLDER, '報告事項', docId);

    doc.saveAndClose();

    // ドキュメントのバージョン更新
    updateDocVersion(docId);

    addAuditLog('UPDATE_MOTIONS_EXT', docId, `議案ブロック更新（決議事項: ${resolutions.length}件、報告事項: ${reports.length}件）`, true);

  } catch (error) {
    console.error('議案ブロック更新エラー:', error);
    throw error;
  }
}

// === AI整形処理 ===
function formatMotionWithAI(inputHtml, title, type = '決議事項') {
  try {
    const properties = PropertiesService.getScriptProperties();
    const useOpenAI = properties.getProperty('USE_OPENAI') === 'true';
    const apiKey = properties.getProperty('OPENAI_API_KEY');

    // HTMLをテキストに変換
    const inputText = htmlToText(inputHtml);

    // OpenAI APIを使用する場合
    if (useOpenAI && apiKey) {
      return formatMotionWithOpenAI(inputText, title, apiKey, type);
    } else {
      // 簡易整形処理にフォールバック
      return simpleFormatMotion(inputText, title, type);
    }

  } catch (error) {
    console.error('AI整形エラー:', error);
    return htmlToText(inputHtml);
  }
}

// === OpenAI APIを使った議案整形 ===
function formatMotionWithOpenAI(inputText, title, apiKey, type = '決議事項') {
  try {
    const properties = PropertiesService.getScriptProperties();
    // モデルはGPT-5に固定
    const model = 'gpt-5';
    const customPrompt = properties.getProperty('AI_PROMPT') || '';

    let systemPrompt = '';

    if (type === '報告事項') {
      systemPrompt = `あなたは企業の法務部門で働く議事録作成のスペシャリストです。
入力された報告事項を、日本の会社法に準拠した自然な日本語の議事録として整形してください。

【重要】出力形式：
- 見出しや項目番号（【】や1.2.3.など）は使用せず、自然な文章で記述してください
- 「議長の指名により、〇〇から以下の報告がなされた」という書き出しで始めてください
- 報告内容は段落で区切りながら、流れるような文章で記述してください
- 最後は「本件は報告事項として説明がなされ、出席取締役はこれを了承した。」で締めくくってください`;
    } else {
      systemPrompt = `あなたは企業の法務部門で働く議事録作成のスペシャリストです。
入力された議案内容を、日本の会社法に準拠した自然な日本語の議事録として整形してください。

【重要】出力形式：
- 見出しや項目番号（【】や1.2.3.など）は使用せず、自然な文章で記述してください
- 「議長の指名により、〇〇から以下の説明がなされた」という書き出しで始めてください
- 議案内容は段落で区切りながら、流れるような文章で記述してください
- 最後は「議長が本議案について議場に諮ったところ、出席取締役全員異議なく、原案どおり承認可決された。」で締めくくってください

文体の注意：
- 議事録特有の丁寧な文体（である調）を使用
- 箇条書きは最小限に留め、文章で説明する
- 「〜について説明がなされた」「〜することとした」などの議事録らしい表現を使用
- 具体的な数値や事実は明確に記載し、曖昧な表現は避ける

【会社法・開示規制の観点からの確認】
議案内容に以下の要素が含まれる場合は、必ず法的留意事項を文末に追加してください：

1. 取締役会決議事項（会社法362条4項）に該当する場合
   - 重要な財産の処分・譲受け
   - 多額の借財
   - 重要な使用人の選任・解任
   - 重要な組織の設置・変更・廃止
   → 「なお、本件は会社法第362条第4項に定める重要な〇〇に該当するため、取締役会の決議事項である。」

2. 利益相反取引（会社法356条）に該当する場合
   - 取締役の自己取引
   - 取締役と会社の利益が相反する取引
   → 「本取引は会社法第356条に定める利益相反取引に該当するため、取締役会の承認を要する。」

3. 上場企業の適時開示事項に該当する可能性がある場合
   - 決算情報、業績予想の修正
   - 配当予想の修正
   - 新たな事業の開始、重要な事業からの撤退
   - 合併、会社分割、株式交換、事業譲渡
   - 第三者割当増資、新株予約権の発行
   → 「本件は東京証券取引所の適時開示規則に該当する可能性があるため、開示の要否について確認を行うこととする。」

4. インサイダー取引規制に関わる事項
   - 重要事実に該当する可能性がある情報
   → 「本件は金融商品取引法上の重要事実に該当する可能性があるため、情報管理に十分留意することとする。」

5. その他の法的検討事項
   - 定款変更を要する事項
   - 株主総会決議を要する事項
   - 官公庁への届出・許認可を要する事項
   → 該当する法令や手続きを明記

${customPrompt}`;
    }

    const userPrompt = `
議案タイトル: ${title}

入力内容:
${inputText}

上記の内容を標準フォーマットに整形してください。`;

    // OpenAI APIリクエスト
    const url = 'https://api.openai.com/v1/chat/completions';
    const payload = {
      model: model,
      messages: [
        {
          role: 'system',
          content: systemPrompt
        },
        {
          role: 'user',
          content: userPrompt
        }
      ],
      temperature: 0.3,
      max_tokens: 2000
    };

    const options = {
      method: 'post',
      headers: {
        'Authorization': `Bearer ${apiKey}`,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();

    if (responseCode === 200) {
      const jsonResponse = JSON.parse(response.getContentText());
      const formattedText = jsonResponse.choices[0].message.content;

      // タイトルを追加
      return `【${title}】\n\n${formattedText}`;
    } else {
      // APIエラーの場合は簡易整形にフォールバック
      console.error('OpenAI APIエラー:', response.getContentText());
      return simpleFormatMotion(inputText, title, type);
    }

  } catch (error) {
    console.error('OpenAI API呼び出しエラー:', error);
    // エラー時は簡易整形にフォールバック
    return simpleFormatMotion(inputText, title, type);
  }
}

// === 簡易整形処理（AIが使えない場合） ===
function simpleFormatMotion(inputText, title, type = '決議事項') {
  // 入力テキストをクリーンアップ
  const cleanedText = inputText.trim().replace(/\n{3,}/g, '\n\n');

  // 記号やラベルを削除
  let formattedText = cleanedText
    .replace(/【.*?】/g, '') // 【】ラベルを削除
    .replace(/[■□◆◇○●]/g, '') // 記号を削除
    .replace(/^\d+[\.\)．）]\s*/gm, '') // 番号付きリストを削除
    .replace(/^[・・･-]\s*/gm, '') // 箇条書き記号を削除
    .trim();

  // タイトルから議案の種類を推測
  const titleLower = title.toLowerCase();

  // 短い入力を詳細に膨らませる
  if (formattedText.length < 100) {
    // 短い文章の場合、自動的に詳細を補完
    formattedText = expandShortMotion(formattedText, title);
  } else {
    // 通常の整形処理
    // 本文を自然な段落に整形
    const paragraphs = formattedText.split(/\n\n+/);
    formattedText = '';
    paragraphs.forEach((para, index) => {
      if (para.trim()) {
        // 各段落を自然な文章に整形
        let cleanPara = para.trim();
        // 最初の段落は導入文として整形
        if (index === 0 && !cleanPara.match(/^当社|本件|今回/)) {
          cleanPara = `本件に関し、${cleanPara}`;
        }
        formattedText += cleanPara;
        if (index < paragraphs.length - 1) {
          formattedText += '\n\n';
        }
      }
    });
  }

  // 議案の内容から法的留意点を簡易的にチェック
  const legalNotes = checkLegalRequirements(formattedText, title);
  if (legalNotes.length > 0) {
    formattedText += '\n\n';
    legalNotes.forEach(note => {
      formattedText += note + '\n';
    });
  }

  // 文章の一貫性をチェック
  formattedText = ensureTextCoherence(formattedText, title);

  return formattedText;
}

// === 文章の一貫性を確保 ===
function ensureTextCoherence(text, title) {
  // 文章が断片化していないか確認
  let coherentText = text;

  // 文末の統一（である調に統一）
  coherentText = coherentText
    .replace(/です。/g, 'である。')
    .replace(/ます。/g, 'る。')
    .replace(/ました。/g, 'た。')
    .replace(/でした。/g, 'であった。')
    .replace(/ません。/g, 'ない。')
    .replace(/ください。/g, 'こととする。');

  // 文章の接続を確認
  const sentences = coherentText.split(/。/);
  let rebuiltText = '';

  sentences.forEach((sentence, index) => {
    let trimmedSentence = sentence.trim();
    if (!trimmedSentence) return;

    // 文頭が小文字や接続語なしで始まる場合の修正
    if (index > 0 && rebuiltText.length > 0) {
      // 前の文との関係を確認
      if (!trimmedSentence.match(/^(また|なお|ただし|さらに|加えて|一方|他方|これにより|その結果|したがって|よって|以上)/)) {
        // 接続が不自然な場合は適切な接続語を追加
        if (trimmedSentence.match(/^(当社|本件|今回|同)/)) {
          // 主語で始まる場合はそのまま
        } else if (index === sentences.length - 2) {
          // 最後から2番目の文の場合
          trimmedSentence = 'なお、' + trimmedSentence;
        }
      }
    }

    rebuiltText += trimmedSentence;
    if (index < sentences.length - 1 && trimmedSentence) {
      rebuiltText += '。';
    }
  });

  // 段落の整合性チェック
  const paragraphs = rebuiltText.split(/\n\n+/);
  let finalText = '';

  paragraphs.forEach((para, index) => {
    let trimmedPara = para.trim();
    if (!trimmedPara) return;

    // 段落の文字数が極端に短い場合は前後の段落と結合を検討
    if (trimmedPara.length < 30 && index > 0 && finalText) {
      // 前の段落に結合
      finalText += 'また、' + trimmedPara;
    } else {
      if (finalText) finalText += '\n\n';
      finalText += trimmedPara;
    }
  });

  // 最終的な文末処理
  if (!finalText.endsWith('。')) {
    finalText += '。';
  }

  return finalText;
}

// === 短い議案を詳細に展開 ===
function expandShortMotion(text, title) {
  let expanded = '';
  const textLower = text.toLowerCase();
  const titleLower = title.toLowerCase();

  // キーワードベースで内容を推測して展開
  if (titleLower.includes('ラブアン') || textLower.includes('ラブアン') ||
      titleLower.includes('国際') || textLower.includes('海外')) {
    // ラブアン法人関連の場合
    expanded = `当社は、グローバルな事業展開の一環として、マレーシア・ラブアン国際ビジネス金融センターにおける法人運営を検討してまいりました。\n\n`;
    expanded += `ラブアン法人は、国際金融センターとしての制度的優位性を有しており、以下のメリットが期待されます。\n`;
    expanded += `第一に、税制面での優遇措置により、効率的な資金管理が可能となります。`;
    expanded += `第二に、規制面での柔軟性により、国際取引の円滑化が図れます。`;
    expanded += `第三に、アジア太平洋地域へのアクセスポイントとして、戦略的な立地を活用できます。\n\n`;
    expanded += `${text}\n\n`;
    expanded += `本件の実施により、当社のグローバル競争力の強化と、新たな収益機会の創出が期待されます。`;
    expanded += `なお、実施にあたっては、現地法規制の遵守と適切なガバナンス体制の構築に万全を期すこととします。`;
  } else if (titleLower.includes('ポケモン') || textLower.includes('カード') ||
             titleLower.includes('事業') || textLower.includes('新規')) {
    // 新規事業関連の場合
    expanded = `当社は、事業ポートフォリオの多様化と収益基盤の強化を目的として、新規事業への参入を検討してまいりました。\n\n`;
    expanded += `${text}\n\n`;
    expanded += `本事業は、当社の既存事業とのシナジー効果が期待でき、以下の点で戦略的意義があります。\n`;
    expanded += `第一に、新たな顧客層へのアプローチが可能となり、収益源の多様化が図れます。`;
    expanded += `第二に、デジタルトランスフォーメーションの推進により、事業効率の向上が期待されます。`;
    expanded += `第三に、市場の成長性が高く、中長期的な収益拡大が見込まれます。\n\n`;
    expanded += `事業開始にあたっては、適切な人員配置と投資計画を策定し、リスク管理体制を整備した上で、段階的に展開することとします。`;
  } else if (titleLower.includes('資金') || textLower.includes('資金') ||
             titleLower.includes('借入') || textLower.includes('融資')) {
    // 資金調達関連の場合
    expanded = `当社の事業拡大および運転資金の確保を目的として、以下の資金調達を実施したく存じます。\n\n`;
    expanded += `${text}\n\n`;
    expanded += `本件資金調達の必要性および合理性については以下のとおりです。\n`;
    expanded += `第一に、事業拡大に伴う設備投資資金として活用し、生産能力の向上を図ります。`;
    expanded += `第二に、運転資金の充実により、財務基盤の安定化を実現します。`;
    expanded += `第三に、新規事業への投資資金として活用し、成長機会を確実に捕捉します。\n\n`;
    expanded += `調達条件については、複数の金融機関と協議の上、最も有利な条件で実施することとし、返済計画についても無理のない範囲で設定いたします。`;
  } else {
    // その他の一般的な議案
    expanded = `${text}\n\n`;
    expanded += `本件は、当社の経営戦略上重要な案件であり、以下の観点から実施が必要と判断されます。\n`;
    expanded += `第一に、事業の継続性と発展性の観点から、適時適切な対応が求められます。`;
    expanded += `第二に、競争優位性の確保と市場での地位向上に寄与することが期待されます。`;
    expanded += `第三に、ステークホルダーの期待に応え、企業価値の向上につながります。\n\n`;
    expanded += `実施にあたっては、関係部門との連携を密にし、適切な進捗管理を行うこととします。`;
  }

  return expanded;
}

// === 法的要件の簡易チェック ===
function checkLegalRequirements(text, title) {
  const notes = [];
  const lowerText = (text + title).toLowerCase();

  // 重要な財産の処分・譲受けのチェック
  if (lowerText.includes('譲渡') || lowerText.includes('売却') ||
      lowerText.includes('取得') || lowerText.includes('買収')) {
    notes.push('なお、本件は会社法第362条第4項に定める重要な財産の処分・譲受けに該当する可能性があるため、取締役会の決議事項である。');
  }

  // 借入のチェック
  if (lowerText.includes('借入') || lowerText.includes('融資') ||
      lowerText.includes('ローン')) {
    notes.push('本件は会社法第362条第4項に定める多額の借財に該当する可能性があるため、その金額規模に留意する必要がある。');
  }

  // 新事業のチェック
  if (lowerText.includes('新事業') || lowerText.includes('新規事業') ||
      lowerText.includes('事業開始') || lowerText.includes('参入')) {
    notes.push('本件は東京証券取引所の適時開示規則における新たな事業の開始に該当する可能性があるため、開示の要否について確認を行うこととする。');
  }

  // 組織変更のチェック
  if (lowerText.includes('組織') || lowerText.includes('部門') ||
      lowerText.includes('支店') || lowerText.includes('営業所')) {
    notes.push('本件は会社法第362条第4項に定める重要な組織の設置・変更に該当する可能性がある。');
  }

  // 人事のチェック
  if (lowerText.includes('選任') || lowerText.includes('解任') ||
      lowerText.includes('任命') || lowerText.includes('異動')) {
    if (lowerText.includes('執行役員') || lowerText.includes('部長') ||
        lowerText.includes('支店長')) {
      notes.push('本件は会社法第362条第4項に定める重要な使用人の選任・解任に該当する。');
    }
  }

  // 利益相反のチェック
  if (lowerText.includes('利益相反') || lowerText.includes('自己取引') ||
      lowerText.includes('競業')) {
    notes.push('本取引は会社法第356条に定める利益相反取引に該当するため、取締役会の承認を要する。');
  }

  // 配当のチェック
  if (lowerText.includes('配当') || lowerText.includes('剰余金')) {
    notes.push('本件は配当に関する事項であり、東京証券取引所の適時開示規則に該当する可能性があるため、適切な開示手続きを行う必要がある。');
  }

  // M&A関連のチェック
  if (lowerText.includes('合併') || lowerText.includes('買収') ||
      lowerText.includes('統合') || lowerText.includes('株式交換')) {
    notes.push('本件は東京証券取引所の適時開示規則および金融商品取引法上の重要事実に該当する可能性があるため、情報管理および開示手続きに十分留意することとする。');
  }

  return notes;
}

// === 議案をコンテキスト付きで整形 ===
function formatMotionWithContext(context) {
  try {
    const properties = PropertiesService.getScriptProperties();
    const useOpenAI = properties.getProperty('USE_OPENAI') === 'true';
    const apiKey = properties.getProperty('OPENAI_API_KEY');

    // HTMLをテキストに変換
    const contentText = htmlToText(context.content);

    // OpenAI APIを使用する場合
    if (useOpenAI && apiKey) {
      return formatMotionWithOpenAIContext(contentText, context, apiKey);
    } else {
      // 簡易整形処理にフォールバック
      return simpleFormatMotionWithContext(contentText, context);
    }

  } catch (error) {
    console.error('AI整形エラー:', error);
    return htmlToText(context.content);
  }
}

// === OpenAI APIを使った議案整形（コンテキスト版） ===
function formatMotionWithOpenAIContext(contentText, context, apiKey) {
  try {
    const systemPrompt = `あなたは企業の法務部門で働く議事録作成のスペシャリストです。
入力された議案内容を、日本の会社法に準拠した自然な日本語の議事録として整形してください。

【重要】出力形式：
- 見出しや項目番号（【】や1.2.3.など）は使用せず、自然な文章で記述してください
- 「議長の指名により、〇〇から以下の説明がなされた」という書き出しで始めてください
- 議案内容は段落で区切りながら、流れるような文章で記述してください
- 最後は「議長が本議案について議場に諮ったところ、${context.resolutionResult}した。」で締めくくってください
${context.votingDetails ? `- 賛否の詳細は自然な形で文章に組み込んでください` : ''}
${context.specialInterest ? `- 特別利害関係人については「なお、${context.specialInterest}は本議案について特別の利害関係を有するため、本議案の議決には参加しなかった。」という形で記載してください` : ''}
${context.conditions ? `- 付帯条件は「ただし、${context.conditions}を条件とする。」という形で追記してください` : ''}`;

    const userPrompt = `以下の議案を整形してください：

議案タイトル: ${context.title}
説明者: ${context.presenter}
内容: ${contentText}`;

    // OpenAI API呼び出し（実装済みの関数を使用）
    const url = 'https://api.openai.com/v1/chat/completions';
    const payload = {
      model: 'gpt-5',
      messages: [
        { role: 'system', content: systemPrompt },
        { role: 'user', content: userPrompt }
      ],
      temperature: 0.3,
      max_tokens: 2000
    };

    const options = {
      method: 'post',
      contentType: 'application/json',
      headers: {
        'Authorization': `Bearer ${apiKey}`
      },
      payload: JSON.stringify(payload)
    };

    const response = UrlFetchApp.fetch(url, options);
    const json = JSON.parse(response.getContentText());

    if (json.choices && json.choices[0] && json.choices[0].message) {
      return json.choices[0].message.content.trim();
    }

    // フォールバック
    return simpleFormatMotionWithContext(contentText, context);

  } catch (error) {
    console.error('OpenAI API エラー:', error);
    return simpleFormatMotionWithContext(contentText, context);
  }
}

// === 簡易整形処理（コンテキスト版） ===
function simpleFormatMotionWithContext(contentText, context) {
  let result = '';

  // 記号やラベルを削除
  let cleanText = contentText
    .replace(/【.*?】/g, '') // 【】ラベルを削除
    .replace(/[■□◆◇○●]/g, '') // 記号を削除
    .replace(/^\d+[\.\)．）]\s*/gm, '') // 番号付きリストを削除
    .replace(/^[・・･-]\s*/gm, '') // 箇条書き記号を削除
    .trim();

  // 説明者の導入（自然な文章として）
  if (context.presenter && context.presenter !== '議長') {
    result += `議長の指名により、${context.presenter}から次のとおり説明があった。\n\n`;
  } else {
    result += `議長から、${context.title}について次の説明があった。\n\n`;
  }

  // 短い文章の場合は拡張
  if (contentText.length < 100) {
    const expanded = expandShortMotion(contentText, context.title);
    result += expanded + '\n\n';
  } else {
    // 長い文章はそのまま使用
    result += contentText + '\n\n';
  }

  // 特別利害関係人
  if (context.specialInterest) {
    result += `なお、${context.specialInterest}は本議案について特別の利害関係を有するため、本議案の議決には参加しなかった。\n\n`;
  }

  // 決議結果
  result += `議長が本議案について議場に諮ったところ、${context.resolutionResult}した。`;

  // 賛否詳細
  if (context.votingDetails) {
    result += `（${context.votingDetails}）`;
  }

  // 付帯条件
  if (context.conditions) {
    result += `\nただし、${context.conditions}を条件とする。`;
  }

  // 法的要件チェック
  const legalNotes = checkLegalRequirements(contentText, context.title);
  if (legalNotes.length > 0) {
    result += '\n\n' + legalNotes.join('\n');
  }

  // 文章の一貫性を確保
  result = ensureTextCoherence(result, context.title);

  return result;
}

// === HTMLからテキスト変換 ===
function htmlToText(html) {
  let text = html.toString();

  text = text.replace(/<br\s*\/?>/gi, '\n');
  text = text.replace(/<\/p>/gi, '\n\n');
  text = text.replace(/<\/div>/gi, '\n');
  text = text.replace(/<\/li>/gi, '\n');
  text = text.replace(/<li>/gi, '・ ');
  text = text.replace(/<[^>]*>/g, '');
  text = text.replace(/&nbsp;/g, ' ');
  text = text.replace(/&lt;/g, '<');
  text = text.replace(/&gt;/g, '>');
  text = text.replace(/&amp;/g, '&');
  text = text.replace(/&quot;/g, '"');
  text = text.replace(/&#39;/g, "'");
  text = text.replace(/\n\n+/g, '\n\n');

  return text.trim();
}

// === 議案番号重複チェック ===
function isMotionNumberDuplicate(docId, motionNumber, type = '決議事項') {
  const ss = getSpreadsheet();
  const motionsSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.MOTIONS);

  if (!motionsSheet || motionsSheet.getLastRow() <= 1) {
    return false;
  }

  const data = motionsSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    // docId、種別、議案番号の3つが一致する場合のみ重複とみなす
    if (data[i][0] === docId &&
        (data[i][1] || '決議事項') === type &&
        data[i][2].toString() === motionNumber.toString()) {
      return true;
    }
  }

  return false;
}

// === 議案プレビュー生成 ===
function previewMotion(inputHtml, title, useAI, type = '決議事項') {
  try {
    let previewText = inputHtml;

    if (useAI) {
      previewText = formatMotionWithAI(inputHtml, title, type);
    } else {
      previewText = htmlToText(inputHtml);
    }

    return {
      success: true,
      preview: previewText
    };

  } catch (error) {
    console.error('議案プレビューエラー:', error);
    return {
      success: false,
      error: error.toString(),
      preview: htmlToText(inputHtml)
    };
  }
}

// ==========================================
// ユーティリティ関数
// ==========================================

// === 監査ログ記録 ===
function addAuditLog(operation, docId, details, success) {
  try {
    const ss = getSpreadsheet();
    const auditSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.AUDIT_LOG);

    if (!auditSheet) {
      console.error('AuditLogシートが見つかりません');
      return;
    }

    const timestamp = new Date();
    const user = Session.getActiveUser().getEmail();

    auditSheet.appendRow([
      timestamp,
      operation,
      user,
      docId || '',
      details || '',
      success ? '成功' : '失敗'
    ]);
  } catch (error) {
    console.error('監査ログ記録エラー:', error);
  }
}

// === ドキュメントID生成 ===
function generateDocId() {
  return Utilities.getUuid();
}

// === 日付フォーマット ===
function formatDate(date) {
  return Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
}

// === メールバリデーション ===
function validateEmail(email) {
  const regex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return regex.test(email);
}

// === ヘルパー関数 ===
function getMotionsByDocId(docId) {
  const ss = getSpreadsheet();
  const motionsSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.MOTIONS);

  if (!motionsSheet || motionsSheet.getLastRow() <= 1) {
    return [];
  }

  const data = motionsSheet.getDataRange().getValues();
  const motions = [];

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === docId) {
      motions.push({
        docId: data[i][0],
        type: data[i][1] || '決議事項',
        motionNumber: data[i][2],
        title: data[i][3],
        inputHtml: data[i][4],
        generatedText: data[i][5],
        hasAttachment: data[i][6],
        attachmentMemo: data[i][7],
        generatedTime: data[i][8],
        lastEditor: data[i][9],
        presenter: data[i][10],
        resolutionResult: data[i][11],
        votingDetails: data[i][12],
        specialInterest: data[i][13],
        conditions: data[i][14]
      });
    }
  }

  // 議案番号でソート
  motions.sort((a, b) => {
    return parseInt(a.motionNumber) - parseInt(b.motionNumber);
  });

  return motions;
}

function getDocumentInfo(docId) {
  const ss = getSpreadsheet();
  const docsSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.DOCS);
  const data = docsSheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === docId) {
      return {
        docId: data[i][0],
        meetingType: data[i][1],
        meetingDate: data[i][2],
        title: data[i][3],
        url: data[i][4],
        applicant: data[i][5],
        deadline: data[i][6],
        status: data[i][7]
      };
    }
  }

  return null;
}

function clearExistingApprovers(docId) {
  const ss = getSpreadsheet();
  const approversSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.APPROVERS);
  const data = approversSheet.getDataRange().getValues();

  // 後ろから削除（インデックスがずれないように）
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][0] === docId) {
      approversSheet.deleteRow(i + 1);
    }
  }
}

function updateDocDeadline(docId, deadline) {
  const ss = getSpreadsheet();
  const docsSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.DOCS);
  const data = docsSheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === docId) {
      docsSheet.getRange(i + 1, 7).setValue(deadline);
      break;
    }
  }
}

// === データ取得関数（UI用） ===

// 役員リスト取得
function getOfficersList() {
  try {
    console.log('getOfficersList開始');
    const ss = getSpreadsheet();
    const officersSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.OFFICERS);

    if (!officersSheet) {
      console.log('Officersシートが見つかりません - 初期化を試行');
      initializeSheets(ss);
      const retrySheet = ss.getSheetByName(CONFIG.SHEET_NAMES.OFFICERS);
      if (!retrySheet) {
        console.error('Officersシート作成失敗');
        return [];
      }
      return getOfficersList(); // 再帰呼び出し
    }

    const lastRow = officersSheet.getLastRow();
    console.log('Officers最終行:', lastRow);

    if (lastRow <= 1) {
      console.log('役員データがありません - サンプルデータを追加');
      // サンプルデータを追加
      officersSheet.getRange(2, 1, 3, 6).setValues([
        ['山田太郎', '代表取締役社長', 'yamada@example.com', '取締役', '2020/04/01', ''],
        ['鈴木花子', '取締役CFO', 'suzuki@example.com', '取締役', '2021/04/01', ''],
        ['佐藤次郎', '監査役', 'sato@example.com', '監査役', '2019/04/01', '常勤']
      ]);
    }

    const data = officersSheet.getRange(2, 1, officersSheet.getLastRow() - 1, 6).getValues();
    console.log('取得したデータ行数:', data.length);

    const officers = data.map(row => ({
      name: row[0],
      position: row[1],
      email: row[2],
      category: row[3],
      startDate: row[4],
      note: row[5]
    })).filter(officer => officer.name);

    console.log('フィルター後の役員数:', officers.length);
    return officers;

  } catch (error) {
    console.error('役員リスト取得エラー:', error);
    return [];
  }
}

// === 役員管理ダイアログ表示 ===
function showOfficersManagementDialog() {
  const html = HtmlService.createHtmlOutputFromFile('officers_management')
    .setWidth(800)
    .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, '役員管理');
}

// === 役員追加 ===
function addOfficer(officerData) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.OFFICERS);
    if (!sheet) throw new Error('Officersシートが見つかりません');

    const lastRow = sheet.getLastRow();
    sheet.getRange(lastRow + 1, 1, 1, 6).setValues([[
      officerData.name,
      officerData.position,
      officerData.email || '',
      officerData.category || '',
      officerData.startDate || '',
      officerData.note || ''
    ]]);

    addAuditLog('ADD_OFFICER', null, `役員追加: ${officerData.name}`, true);
    return { success: true, message: '役員を追加しました' };
  } catch (error) {
    console.error('役員追加エラー:', error);
    return { success: false, error: error.toString() };
  }
}

// === 役員削除 ===
function deleteOfficer(name) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.OFFICERS);
    if (!sheet) throw new Error('Officersシートが見つかりません');

    const data = sheet.getDataRange().getValues();
    for (let i = data.length - 1; i >= 1; i--) {
      if (data[i][0] === name) {
        sheet.deleteRow(i + 1);
        addAuditLog('DELETE_OFFICER', null, `役員削除: ${name}`, true);
        return { success: true, message: '役員を削除しました' };
      }
    }

    return { success: false, error: '該当する役員が見つかりません' };
  } catch (error) {
    console.error('役員削除エラー:', error);
    return { success: false, error: error.toString() };
  }
}

// === 会議後データ更新 ===
function updatePostMeetingData(params) {
  try {
    const {
      docId,
      attendingOfficers,
      absentOfficers,
      actualMeetingDate,
      actualMeetingTime,
      actualLocation,
      actualChair,
      motions,
      nextMeetingDate,
      notes
    } = params;

    if (!docId) {
      throw new Error('文書IDが指定されていません');
    }

    const doc = DocumentApp.openById(docId);
    const body = doc.getBody();

    // 出席者・欠席者を氏名 役職形式で整形
    const formatOfficers = (officers) => {
      if (!officers || officers.length === 0) return '';
      return officers.join('、');
    };

    // プレースホルダ置換
    const replacements = {
      '{{ATTENDING_OFFICERS}}': formatOfficers(attendingOfficers),
      '{{ABSENT_OFFICERS}}': formatOfficers(absentOfficers),
      '{{ATTENDEES}}': formatOfficers(attendingOfficers),
      '{{ABSENTEES}}': formatOfficers(absentOfficers)
    };

    // ドキュメント全体を検索して置換
    const text = body.getText();
    Object.keys(replacements).forEach(key => {
      body.replaceText(key, replacements[key]);
    });

    // 議案の決議結果を更新
    if (motions && motions.length > 0) {
      const ss = getSpreadsheet();
      const motionsSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.MOTIONS);

      motions.forEach(motion => {
        // スプレッドシートの該当議案を更新
        const data = motionsSheet.getDataRange().getValues();
        for (let i = 1; i < data.length; i++) {
          if (data[i][0] === docId && data[i][1].toString() === motion.number.toString()) {
            // 決議結果等を更新
            motionsSheet.getRange(i + 1, 11).setValue(motion.resolutionResult || '全員一致承認');
            motionsSheet.getRange(i + 1, 12).setValue(motion.votingDetails || '');
            motionsSheet.getRange(i + 1, 13).setValue(motion.specialInterest || '');
            motionsSheet.getRange(i + 1, 14).setValue(motion.conditions || '');
            break;
          }
        }
      });

      // ドキュメントの議案ブロックを更新
      updateMotionsBlockExtended(docId);
    }

    doc.saveAndClose();
    addAuditLog('UPDATE_POST_MEETING', docId, '会議後データ更新完了', true);

    return { success: true, message: '議事録を更新しました' };
  } catch (error) {
    console.error('会議後データ更新エラー:', error);
    return { success: false, error: error.toString() };
  }
}

function getDocsList() {
  try {
    const ss = getSpreadsheet();
    console.log('スプレッドシート取得:', ss ? ss.getName() : 'null');

    const docsSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.DOCS);
    console.log('Docsシート取得:', docsSheet ? 'OK' : 'NOT FOUND');

    if (!docsSheet) {
      console.log('Docsシートが見つかりません');
      return [];
    }

    const lastRow = docsSheet.getLastRow();
    console.log('Docsシートの行数:', lastRow);

    if (lastRow <= 1) {
      console.log('Docsシートにデータがありません（ヘッダーのみ）');
      return [];
    }

    const data = docsSheet.getRange(2, 1, lastRow - 1, 10).getValues();
    console.log('取得したデータ行数:', data.length);

    // 最初の数行をログ出力
    if (data.length > 0) {
      console.log('サンプルデータ（最初の行）:');
      console.log('docId:', data[0][0]);
      console.log('title:', data[0][3]);
      console.log('status:', data[0][7]);
    }

    const result = data.map(row => ({
      docId: row[0],        // A列
      meetingType: row[1],  // B列
      meetingDate: row[2],  // C列
      title: row[3],        // D列
      draftUrl: row[4],     // E列
      applicant: row[5],    // F列
      deadline: row[6],     // G列
      status: row[7],       // H列
      lastUpdate: row[8],   // I列
      version: row[9]       // J列
    })).filter(doc => doc.docId); // docIdが空の行を除外

    console.log('返すドキュメント数:', result.length);

    // HTML Serviceの制限を回避するため、JSON文字列として返す
    return JSON.stringify(result);

  } catch (error) {
    console.error('getDocsList エラー:', error);
    return JSON.stringify([]);  // 空配列をJSON文字列として返す
  }
}

// === テスト関数：Docsシートの内容を確認 ===
function testDocsSheet() {
  const ss = getSpreadsheet();
  const docsSheet = ss.getSheetByName('Docs');

  if (!docsSheet) {
    console.log('Docsシートが見つかりません');
    return;
  }

  const lastRow = docsSheet.getLastRow();
  const lastCol = docsSheet.getLastColumn();

  console.log('=== Docsシート診断 ===');
  console.log('行数:', lastRow);
  console.log('列数:', lastCol);

  if (lastRow > 1) {
    console.log('\n=== ヘッダー行 ===');
    const headers = docsSheet.getRange(1, 1, 1, lastCol).getValues()[0];
    headers.forEach((header, index) => {
      console.log(`${String.fromCharCode(65 + index)}列: ${header}`);
    });

    console.log('\n=== データ行（最初の3行） ===');
    const dataRows = Math.min(3, lastRow - 1);
    const data = docsSheet.getRange(2, 1, dataRows, lastCol).getValues();

    data.forEach((row, rowIndex) => {
      console.log(`\n--- 行${rowIndex + 2} ---`);
      console.log('A列(docId):', row[0]);
      console.log('B列(meetingType):', row[1]);
      console.log('C列(meetingDate):', row[2]);
      console.log('D列(title):', row[3]);
      console.log('H列(status):', row[7]);
    });
  } else {
    console.log('データ行がありません');
  }

  // getDocsList関数も実行してみる
  console.log('\n=== getDocsList関数の結果 ===');
  const docs = getDocsList();
  console.log('取得したドキュメント数:', docs.length);
  if (docs.length > 0) {
    console.log('最初のドキュメント:', docs[0]);
  }
}

function getMeetingTypes() {
  const ss = getSpreadsheet();
  const templatesSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.TEMPLATES);

  if (!templatesSheet || templatesSheet.getLastRow() <= 1) {
    return [];
  }

  const data = templatesSheet.getRange(2, 1, templatesSheet.getLastRow() - 1, 1).getValues();
  return data.map(row => row[0]).filter(type => type);
}

function getMotionsList(docId) {
  return getMotionsByDocId(docId);
}

function getApprovalStatus(docId) {
  const approvals = getApprovalsByDocId(docId);
  const docInfo = getDocumentInfo(docId);

  return {
    docInfo: docInfo,
    approvals: approvals,
    summary: {
      total: approvals.length,
      approved: approvals.filter(a => a.status === CONFIG.APPROVAL_STATUS.APPROVED).length,
      rejected: approvals.filter(a => a.status === CONFIG.APPROVAL_STATUS.REJECTED).length,
      pending: approvals.filter(a => a.status === CONFIG.APPROVAL_STATUS.PENDING).length
    }
  };
}

// === ダイアログ表示関数（未実装） ===
function showCreateDraftDialog() {
  const html = HtmlService.createHtmlOutputFromFile('create_draft')
    .setWidth(600)
    .setHeight(500);
  SpreadsheetApp.getUi().showModalDialog(html, '新規ドラフト作成');
}

function showAddMotionDialog() {
  const html = HtmlService.createHtmlOutputFromFile('add_motion')
    .setWidth(700)
    .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, '決議事項追加');
}

function showAddReportDialog() {
  const html = HtmlService.createHtmlOutputFromFile('add_report')
    .setWidth(700)
    .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, '報告事項追加');
}

function showPostMeetingEditorDialog() {
  const html = HtmlService.createHtmlOutputFromFile('post_meeting_editor')
    .setWidth(700)
    .setHeight(700);
  SpreadsheetApp.getUi().showModalDialog(html, '会議終了後 議事録追記');
}

function showConfigDialog() {
  const html = HtmlService.createHtmlOutputFromFile('config')
    .setWidth(600)
    .setHeight(500);
  SpreadsheetApp.getUi().showModalDialog(html, 'システム設定');
}

function showAuditLogDialog() {
  const html = HtmlService.createHtmlOutputFromFile('audit_log')
    .setWidth(900)
    .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, '監査ログ閲覧');
}

// === ステータス変更ダイアログ ===
function showStatusChangeDialog() {
  const html = HtmlService.createHtmlOutputFromFile('status_change')
    .setWidth(500)
    .setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(html, 'ステータス変更');
}

// === ドキュメント一覧取得（ステータス変更用） ===
function getDocumentsForStatusChange() {
  try {
    const ss = getSpreadsheet();
    const docsSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.DOCS);

    if (!docsSheet || docsSheet.getLastRow() <= 1) {
      return [];
    }

    const data = docsSheet.getDataRange().getValues();
    const docs = [];

    for (let i = 1; i < data.length; i++) {
      docs.push({
        docId: data[i][0],
        meetingType: data[i][1],
        meetingDate: data[i][2] ? formatDate(new Date(data[i][2])) : '',
        title: data[i][3],
        status: data[i][7] || CONFIG.STATUS.DRAFT
      });
    }

    return docs;
  } catch (error) {
    console.error('ドキュメント一覧取得エラー:', error);
    return [];
  }
}

// === ステータス変更 ===
function changeDocumentStatus(docId, newStatus) {
  try {
    const ss = getSpreadsheet();
    const docsSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.DOCS);
    const data = docsSheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === docId) {
        const oldStatus = data[i][7];
        docsSheet.getRange(i + 1, 8).setValue(newStatus);

        addAuditLog('CHANGE_STATUS', docId,
          `ステータス変更: ${oldStatus} → ${newStatus}`, true);

        return {
          success: true,
          message: `ステータスを「${newStatus}」に変更しました`
        };
      }
    }

    throw new Error('ドキュメントが見つかりません');

  } catch (error) {
    console.error('ステータス変更エラー:', error);
    addAuditLog('CHANGE_STATUS_ERROR', docId, error.toString(), false);
    return {
      success: false,
      error: error.toString()
    };
  }
}

// ==========================================
// 承認確認機能
// ==========================================

// === 承認確認メール送信 ===
function sendApprovalCheckEmail(docId, contactEmail) {
  try {
    const docInfo = getDocumentInfo(docId);
    if (!docInfo) {
      throw new Error('ドキュメント情報が見つかりません');
    }

    const properties = PropertiesService.getScriptProperties();
    const fromName = properties.getProperty('FROM_NAME') || '議事録管理システム';

    const subject = `【承認確認】議事録承認状況の確認: ${docInfo.title}`;

    const body = `
承認確認担当者様

以下の議事録について、承認期限が経過しましたので、承認状況をご確認ください。

【議事録情報】
会議種別: ${docInfo.meetingType}
会議日: ${formatDate(new Date(docInfo.meetingDate))}
件名: ${docInfo.title}
ドキュメントURL: ${docInfo.url}
現在のステータス: ${docInfo.status}

※承認が完了している場合は、メニューの「ステータス変更」から「承認済」に変更してください。

---
議事録管理システム
`;

    GmailApp.sendEmail(contactEmail, subject, body, {
      name: fromName
    });

    addAuditLog('SEND_APPROVAL_CHECK', docId, `承認確認メール送信: ${contactEmail}`, true);

    return {
      success: true,
      message: '承認確認メールを送信しました'
    };

  } catch (error) {
    console.error('承認確認メール送信エラー:', error);
    addAuditLog('SEND_APPROVAL_CHECK_ERROR', docId, error.toString(), false);
    return {
      success: false,
      error: error.toString()
    };
  }
}

// === 定期チェック（承認確認） ===
function checkApprovalStatus() {
  try {
    const properties = PropertiesService.getScriptProperties();
    const checkDays = parseInt(properties.getProperty('APPROVAL_CHECK_DAYS') || '7');
    const contactEmail = properties.getProperty('APPROVAL_CONTACT_EMAIL');

    if (!contactEmail) {
      console.log('承認確認通知先が設定されていません');
      return;
    }

    const ss = getSpreadsheet();
    const docsSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.DOCS);

    if (!docsSheet || docsSheet.getLastRow() <= 1) {
      return;
    }

    const data = docsSheet.getDataRange().getValues();
    const now = new Date();
    let sentCount = 0;

    for (let i = 1; i < data.length; i++) {
      const docId = data[i][0];
      const meetingDate = new Date(data[i][2]);
      const status = data[i][7];

      // ステータスが「編集中」で、会議日からcheckDays日経過している場合
      if (status === CONFIG.STATUS.DRAFT) {
        const daysSinceMeeting = (now - meetingDate) / (1000 * 60 * 60 * 24);

        if (daysSinceMeeting >= checkDays) {
          const result = sendApprovalCheckEmail(docId, contactEmail);
          if (result.success) {
            sentCount++;
          }
        }
      }
    }

    if (sentCount > 0) {
      addAuditLog('CHECK_APPROVAL_STATUS', null, `承認確認実施: ${sentCount}件送信`, true);
    }

  } catch (error) {
    console.error('承認確認チェックエラー:', error);
    addAuditLog('CHECK_APPROVAL_STATUS_ERROR', null, error.toString(), false);
  }
}

// ==========================================
// 設定管理機能（config.html用）
// ==========================================

// === すべての設定を取得 ===
function getAllSettings() {
  const properties = PropertiesService.getScriptProperties();
  return properties.getProperties();
}

// === 設定を保存 ===
function saveSettings(newSettings, templates) {
  try {
    const properties = PropertiesService.getScriptProperties();

    // プロパティを更新
    for (const [key, value] of Object.entries(newSettings)) {
      properties.setProperty(key, value);
    }

    // テンプレート設定を更新
    if (templates && templates.length > 0) {
      updateTemplateSettings(templates);
    }

    addAuditLog('UPDATE_SETTINGS', null, '設定更新', true);

    return {
      success: true,
      message: '設定を保存しました'
    };

  } catch (error) {
    console.error('設定保存エラー:', error);
    addAuditLog('UPDATE_SETTINGS_ERROR', null, error.toString(), false);
    return {
      success: false,
      error: error.toString()
    };
  }
}

// === テンプレート設定を取得 ===
function getTemplateSettings() {
  const ss = getSpreadsheet();
  const templatesSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.TEMPLATES);

  if (!templatesSheet || templatesSheet.getLastRow() <= 1) {
    return [];
  }

  const data = templatesSheet.getRange(2, 1, templatesSheet.getLastRow() - 1, 2).getValues();
  return data.map(row => ({
    type: row[0],
    docId: row[1]
  })).filter(t => t.type);
}

// === テンプレート設定を更新 ===
function updateTemplateSettings(templates) {
  const ss = getSpreadsheet();
  const templatesSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.TEMPLATES);

  if (!templatesSheet) {
    throw new Error('Templatesシートが見つかりません');
  }

  // 既存データを取得
  const existingData = templatesSheet.getDataRange().getValues();

  // ヘッダー以外をクリア
  if (existingData.length > 1) {
    templatesSheet.getRange(2, 1, existingData.length - 1, 5).clearContent();
  }

  // 新規データを設定
  templates.forEach((template, index) => {
    templatesSheet.getRange(index + 2, 1).setValue(template.type);
    templatesSheet.getRange(index + 2, 2).setValue(template.docId);
    templatesSheet.getRange(index + 2, 3).setValue(`${template.type}議事録テンプレート`);
    templatesSheet.getRange(index + 2, 4).setValue('1.0');
    templatesSheet.getRange(index + 2, 5).setValue('ユーザー設定');
  });
}

// === テンプレート管理ダイアログ表示 ===
function showTemplateDialog() {
  try {
    // シンプル版を使用
    const html = HtmlService.createHtmlOutputFromFile('simple_template_manager')
      .setWidth(800)
      .setHeight(700);
    SpreadsheetApp.getUi().showModalDialog(html, 'テンプレート管理');
  } catch (error) {
    SpreadsheetApp.getUi().alert('エラー: ' + error.toString());
  }
}

// === テンプレートテストダイアログ表示 ===
function showTemplateTestDialog() {
  const html = HtmlService.createHtmlOutputFromFile('test_template')
    .setWidth(600)
    .setHeight(700);
  SpreadsheetApp.getUi().showModalDialog(html, 'テンプレート機能テスト');
}

// === 旧テンプレート管理ダイアログ表示 ===
function showOldTemplateDialog() {
  const html = HtmlService.createHtmlOutputFromFile('manage_templates')
    .setWidth(900)
    .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, 'テンプレート管理（旧版）');
}

// === テンプレート一覧取得 ===
function getTemplateList() {
  try {
    const ss = getSpreadsheet();
    const templatesSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.TEMPLATES);

    console.log('Templates sheet exists:', !!templatesSheet);

    if (!templatesSheet) {
      console.log('Templates sheet not found');
      return [];
    }

    const lastRow = templatesSheet.getLastRow();
    console.log('Last row:', lastRow);

    if (lastRow <= 1) {
      console.log('No data in Templates sheet');
      return [];
    }

    // データ範囲を安全に取得
    const numRows = lastRow - 1;
    const numCols = Math.min(templatesSheet.getLastColumn(), 5);

    if (numRows <= 0 || numCols <= 0) {
      console.log('Invalid data range');
      return [];
    }

    const data = templatesSheet.getRange(2, 1, numRows, numCols).getValues();
    console.log('Data rows:', data.length);

    const templates = [];

    data.forEach((row, index) => {
      try {
        if (row[0]) { // 会議種別が存在する場合
          const template = {
            type: String(row[0]),
            docId: row[1] ? String(row[1]) : '',
            description: row[2] ? String(row[2]) : '',
            version: row[3] ? String(row[3]) : '1.0',
            source: row[4] ? String(row[4]) : ''
          };

          // ドキュメントURLを追加（エラーが発生してもスキップ）
          if (template.docId) {
            try {
              const doc = DocumentApp.openById(template.docId);
              template.docUrl = doc.getUrl();
            } catch (e) {
              console.log('Cannot access document:', template.docId);
              template.docUrl = null;
            }
          }

          templates.push(template);
        }
      } catch (rowError) {
        console.error('Error processing row ' + index + ':', rowError);
      }
    });

    console.log('Templates found:', templates.length);
    return templates;

  } catch (error) {
    console.error('テンプレート一覧取得エラー:', error);
    console.error('Error details:', error.toString());
    // エラーが発生しても空配列を返す
    return [];
  }
}

// === シンプルなテンプレート一覧取得（デバッグ用） ===
function getSimpleTemplateList() {
  try {
    const ss = getSpreadsheet();
    let templatesSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.TEMPLATES);

    // シートが存在しない場合は作成
    if (!templatesSheet) {
      templatesSheet = ss.insertSheet(CONFIG.SHEET_NAMES.TEMPLATES);
      templatesSheet.getRange(1, 1, 1, 7).setValues([[
        '会議種別', 'ドキュメントID', 'URL', 'プレースホルダー', '説明', 'バージョン', '作成日'
      ]]);
      templatesSheet.getRange(1, 1, 1, 7).setFontWeight('bold');
      templatesSheet.setFrozenRows(1);

      return {
        success: true,
        message: 'Templatesシートを作成しました',
        templates: []
      };
    }

    const lastRow = templatesSheet.getLastRow();
    if (lastRow <= 1) {
      return { error: 'No data', templates: [] };
    }

    const data = templatesSheet.getRange(2, 1, lastRow - 1, 2).getValues();
    const templates = [];

    data.forEach(row => {
      if (row[0]) {
        templates.push({
          type: row[0],
          docId: row[1] || 'N/A'
        });
      }
    });

    return {
      success: true,
      count: templates.length,
      templates: templates
    };

  } catch (error) {
    return {
      error: error.toString(),
      templates: []
    };
  }
}

// === 特定のテンプレート取得 ===
function getTemplateByType(type) {
  try {
    const templates = getTemplateList();
    return templates.find(t => t.type === type) || null;
  } catch (error) {
    console.error('テンプレート取得エラー:', error);
    return null;
  }
}

// === テンプレートドキュメント作成 ===
function createTemplateDocument(params) {
  try {
    const { type, content } = params;

    // 新しいドキュメントを作成
    const doc = DocumentApp.create(`${type}議事録テンプレート`);
    const docId = doc.getId();
    const body = doc.getBody();

    // テンプレート内容を設定
    body.setText(content);

    // スタイル設定
    const title = body.getParagraphs()[0];
    if (title) {
      title.setHeading(DocumentApp.ParagraphHeading.TITLE);
    }

    doc.saveAndClose();

    // ドキュメントのURLを取得
    const docUrl = doc.getUrl();

    // プレースホルダーを抽出
    const placeholders = extractPlaceholders(content);

    // スプレッドシートに記録
    const ss = getSpreadsheet();
    const templatesSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.TEMPLATES);

    // 既存のテンプレートをチェック
    const existingData = templatesSheet.getDataRange().getValues();
    let rowToUpdate = -1;

    for (let i = 1; i < existingData.length; i++) {
      if (existingData[i][0] === type) {
        rowToUpdate = i + 1;
        break;
      }
    }

    if (rowToUpdate > 0) {
      // 既存のテンプレートを更新
      templatesSheet.getRange(rowToUpdate, 2).setValue(docId);
      templatesSheet.getRange(rowToUpdate, 3).setValue(docUrl);
      templatesSheet.getRange(rowToUpdate, 4).setValue(placeholders);
      templatesSheet.getRange(rowToUpdate, 5).setValue(`${type}議事録テンプレート`);
      templatesSheet.getRange(rowToUpdate, 6).setValue('1.0');
      templatesSheet.getRange(rowToUpdate, 7).setValue(new Date());
    } else {
      // 新規追加
      templatesSheet.appendRow([
        type,
        docId,
        docUrl,
        placeholders,
        `${type}議事録テンプレート`,
        '1.0',
        new Date()
      ]);
    }

    addAuditLog('CREATE_TEMPLATE', docId, `テンプレート作成: ${type}`, true);

    return {
      success: true,
      type: type,
      docId: docId,
      url: doc.getUrl()
    };

  } catch (error) {
    console.error('テンプレート作成エラー:', error);
    addAuditLog('CREATE_TEMPLATE_ERROR', null, error.toString(), false);
    return {
      success: false,
      error: error.toString()
    };
  }
}

// === 既存ドキュメントからテンプレートインポート ===
function importTemplateFromDoc(params) {
  try {
    const { docId, type } = params;

    // ドキュメントを開く
    const sourceDoc = DocumentApp.openById(docId);

    // コピーを作成
    const driveFile = DriveApp.getFileById(docId);
    const newFile = driveFile.makeCopy(`${type}議事録テンプレート`);
    const newDocId = newFile.getId();

    // スプレッドシートに記録
    const ss = getSpreadsheet();
    const templatesSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.TEMPLATES);

    // 既存のテンプレートをチェック
    const existingData = templatesSheet.getDataRange().getValues();
    let rowToUpdate = -1;

    for (let i = 1; i < existingData.length; i++) {
      if (existingData[i][0] === type) {
        rowToUpdate = i + 1;
        break;
      }
    }

    if (rowToUpdate > 0) {
      // 既存のテンプレートを更新
      templatesSheet.getRange(rowToUpdate, 2).setValue(newDocId);
      templatesSheet.getRange(rowToUpdate, 3).setValue(`${type}議事録テンプレート`);
      templatesSheet.getRange(rowToUpdate, 4).setValue('1.0');
      templatesSheet.getRange(rowToUpdate, 5).setValue(new Date());
    } else {
      // 新規追加
      templatesSheet.appendRow([
        type,
        newDocId,
        `${type}議事録テンプレート`,
        '1.0',
        new Date()
      ]);
    }

    addAuditLog('IMPORT_TEMPLATE', newDocId, `テンプレートインポート: ${type}`, true);

    return {
      success: true,
      type: type,
      docId: newDocId,
      url: DocumentApp.openById(newDocId).getUrl()
    };

  } catch (error) {
    console.error('テンプレートインポートエラー:', error);
    addAuditLog('IMPORT_TEMPLATE_ERROR', null, error.toString(), false);
    return {
      success: false,
      error: error.toString()
    };
  }
}

// === テンプレート削除 ===
function deleteTemplate(type) {
  try {
    const ss = getSpreadsheet();
    const templatesSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.TEMPLATES);

    if (!templatesSheet) {
      throw new Error('Templatesシートが見つかりません');
    }

    const data = templatesSheet.getDataRange().getValues();
    let rowToDelete = -1;
    let docIdToDelete = null;

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === type) {
        rowToDelete = i + 1;
        docIdToDelete = data[i][1];
        break;
      }
    }

    if (rowToDelete < 0) {
      throw new Error(`${type}のテンプレートが見つかりません`);
    }

    // スプレッドシートから削除
    templatesSheet.deleteRow(rowToDelete);

    // ドキュメントも削除（オプション）
    if (docIdToDelete) {
      try {
        DriveApp.getFileById(docIdToDelete).setTrashed(true);
      } catch (e) {
        console.log('ドキュメントの削除に失敗:', e);
      }
    }

    addAuditLog('DELETE_TEMPLATE', docIdToDelete, `テンプレート削除: ${type}`, true);

    return {
      success: true,
      message: `${type}のテンプレートを削除しました`
    };

  } catch (error) {
    console.error('テンプレート削除エラー:', error);
    addAuditLog('DELETE_TEMPLATE_ERROR', null, error.toString(), false);
    return {
      success: false,
      error: error.toString()
    };
  }
}

// === サンプルテンプレート作成 ===
function createSampleTemplates() {
  try {
    const samples = [
      {
        type: '取締役会',
        content: `{{COMPANY_NAME}}
{{MEETING_TITLE}}議事録

日時: {{YEAR}}年{{MONTH}}月{{DAY}}日 {{HOUR}}時{{MINUTE}}分
場所: {{LOCATION}}

議長: {{CHAIR}}

1. 出席役員
{{ATTENDING_OFFICERS}}

2. 欠席役員
{{ABSENT_OFFICERS}}

3. 開会
定刻、{{CHAIR}}が議長となり、本取締役会が適法に成立した旨を宣言し、開会した。

4. 決議事項
{{RESOLUTIONS_BLOCK}}

5. 報告事項
{{REPORTS_BLOCK}}

6. 決議結果
{{RESOLUTION_RESULT}}

7. 次回予定
{{NEXT_MEETING}}

8. 閉会
以上をもって本日の議事をすべて終了したので、議長は閉会を宣言した。

上記の決議を明確にするため、この議事録を作成し、出席取締役及び出席監査役が次に記名押印する。

{{YEAR}}年{{MONTH}}月{{DAY}}日

{{COMPANY_NAME}}

議長　{{CHAIR}}

`
      },
      {
        type: '株主総会',
        content: `{{COMPANY_NAME}}
{{MEETING_TITLE}}議事録

日時: {{YEAR}}年{{MONTH}}月{{DAY}}日 {{HOUR}}時{{MINUTE}}分
場所: {{LOCATION}}

議長: {{CHAIR}}

1. 出席状況
総株主の議決権数: ○○個
出席株主数: ○○名
出席株主の議決権数: ○○個

2. 開会
{{CHAIR}}は、定刻、議長席につき、本総会が適法に成立したことを宣言し、開会した。

3. 決議事項
{{RESOLUTIONS_BLOCK}}

4. 報告事項
{{REPORTS_BLOCK}}

5. 決議結果
{{RESOLUTION_RESULT}}

6. 閉会
以上をもって本日の議事をすべて終了したので、議長は閉会を宣言した。

上記の決議を明確にするため、この議事録を作成し、議長及び出席取締役が次に記名押印する。

{{YEAR}}年{{MONTH}}月{{DAY}}日

{{COMPANY_NAME}}

議長　{{CHAIR}}

`
      },
      {
        type: '監査役会',
        content: `{{COMPANY_NAME}}
{{MEETING_TITLE}}議事録

日時: {{YEAR}}年{{MONTH}}月{{DAY}}日 {{HOUR}}時{{MINUTE}}分
場所: {{LOCATION}}

議長: {{CHAIR}}

1. 出席監査役
{{ATTENDING_OFFICERS}}

2. 欠席監査役
{{ABSENT_OFFICERS}}

3. 開会
{{CHAIR}}が議長となり、本監査役会が適法に成立した旨を宣言し、開会した。

4. 決議事項
{{RESOLUTIONS_BLOCK}}

5. 報告事項
{{REPORTS_BLOCK}}

6. 決議結果
{{RESOLUTION_RESULT}}

7. 次回予定
{{NEXT_MEETING}}

8. 閉会
以上をもって本日の議事をすべて終了したので、議長は閉会を宣言した。

上記の決議を明確にするため、この議事録を作成し、出席監査役が次に記名押印する。

{{YEAR}}年{{MONTH}}月{{DAY}}日

{{COMPANY_NAME}}

議長　{{CHAIR}}

`
      }
    ];

    const results = [];
    samples.forEach(sample => {
      const result = createTemplateDocument(sample);
      results.push({
        type: sample.type,
        success: result.success,
        docId: result.docId
      });
    });

    addAuditLog('CREATE_SAMPLE_TEMPLATES', null,
      `サンプルテンプレート作成: ${results.length}件`, true);

    return {
      success: true,
      message: `${results.length}件のサンプルテンプレートを作成しました`,
      results: results
    };

  } catch (error) {
    console.error('サンプルテンプレート作成エラー:', error);
    addAuditLog('CREATE_SAMPLE_TEMPLATES_ERROR', null, error.toString(), false);
    return {
      success: false,
      error: error.toString()
    };
  }
}

// === OpenAI接続テスト ===
function testOpenAIConnection(apiKey, _model) {
  try {
    const url = 'https://api.openai.com/v1/chat/completions';
    // モデルはGPT-5に固定
    const payload = {
      model: 'gpt-5',
      messages: [
        {
          role: 'system',
          content: 'You are a helpful assistant.'
        },
        {
          role: 'user',
          content: 'テスト接続です。「接続成功」と返答してください。'
        }
      ],
      temperature: 0.3,
      max_tokens: 100
    };

    const options = {
      method: 'post',
      headers: {
        'Authorization': `Bearer ${apiKey}`,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();

    if (responseCode === 200) {
      const jsonResponse = JSON.parse(response.getContentText());
      return {
        success: true,
        message: 'OpenAI API接続成功',
        response: jsonResponse.choices[0].message.content
      };
    } else {
      const errorResponse = JSON.parse(response.getContentText());
      return {
        success: false,
        error: errorResponse.error?.message || 'API接続エラー'
      };
    }

  } catch (error) {
    console.error('OpenAI接続テストエラー:', error);
    return {
      success: false,
      error: error.toString()
    };
  }
}