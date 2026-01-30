/**
 * 会議背景作成GAS — スライド自動生成 + 共有招待
 *
 * 前提:
 * - コンテナバインド（対象スプレッドシートに紐付け）での利用を想定
 * - Advanced Google Services の Drive API を有効化（サービス名: Drive, v3）
 * - GCP 側の Drive API も有効化
 * - シート構成:
 *   - Config:   A列=Key, B列=Value（OUTPUT_FOLDER_ID/TEMPLATE_FILE_ID）
 *   - TemplateMapping: A列=部署名, B列=テンプレートページ番号（1始まり）
 *   - Employees: （ヘッダ行あり）社員ID/氏名/氏名英語/部署/役職/メール/ステータス/生成スライドURL
 */

// ====== シート名・キー定義 ======
const SHEET = {
  CONFIG: 'Config',
  TEMPLATE_MAPPING: 'TemplateMapping',
  EMPLOYEES: 'Employees',
};

const CONFIG_KEY = {
  OUTPUT_FOLDER_ID: 'OUTPUT_FOLDER_ID',
  TEMPLATE_FILE_ID: 'TEMPLATE_FILE_ID',
};

// Employees シートの推奨ヘッダ名（存在チェック＆参照に使用）
const EMP_HEADERS = {
  ID: '社員ID',
  NAME: '氏名',
  NAME_EN: '氏名英語',
  DEPT: '部署',
  TITLE: '役職',
  EMAIL: 'メール',
  STATUS: 'ステータス',
  URL: '生成スライドURL',
};

// ステータス値
const STATUS = {
  PENDING: '未処理',
  DONE: '完了',
  ERROR_PREFIX: 'エラー: ',
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('会議背景作成')
    .addItem('初期セットアップ（デモ作成）', 'setupSpreadsheetAndDemoAssets')
    .addItem('未処理を実行', 'runForPendingEmployees')
    .addToUi();
}

// メインエントリ: 未処理行のみ実行
function runForPendingEmployees() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = loadConfig_(ss);
  const mapping = loadTemplateMapping_(ss);

  const sheet = ss.getSheetByName(SHEET.EMPLOYEES);
  if (!sheet) throw new Error(`シートが見つかりません: ${SHEET.EMPLOYEES}`);

  // 表示値を用いて取得（表示書式を保持するため）
  const data = getSheetObjectsDisplay_(sheet); // {rows: [...], headers: [...], colIndex: {}} 形式
  const idx = data.colIndex; // 見出し→列番号

  const updates = []; // {row, url, status}
  const folderCache = new Map(); // dept -> {id, url}

  for (const row of data.rows) {
    const rowNumber = row.__rowNumber; // 実データ行番号

    // ステータスが未処理以外はスキップ（空欄も未処理として扱う）
    const status = String(row[EMP_HEADERS.STATUS] || '').trim();
    if (status && status !== STATUS.PENDING) continue;

    try {
      const email = String(row[EMP_HEADERS.EMAIL] || '').trim();
      if (!isValidEmail_(email)) {
        throw new Error('メールアドレスが不正/空です');
      }

      const dept = String(row[EMP_HEADERS.DEPT] || '').trim();
      const pageNumber = mapping.get(dept); // 1-based index
      if (!pageNumber) {
        throw new Error('mapping not found（部署に紐づくテンプレートページ番号がありません）');
      }

      // テンプレート（単一ファイル）をコピー
      const outputFolderId = getConfigValue_(config, CONFIG_KEY.OUTPUT_FOLDER_ID);
      const templateFileId = getConfigValue_(config, CONFIG_KEY.TEMPLATE_FILE_ID);
      if (!outputFolderId) throw new Error('Config: OUTPUT_FOLDER_ID が未設定です');
      if (!templateFileId) throw new Error('Config: TEMPLATE_FILE_ID が未設定です');

      const fileName = buildFileName_(row);
      const deptName = dept || '未分類';
      let deptFolder = folderCache.get(deptName);
      if (!deptFolder) {
        deptFolder = getOrCreateSubFolder_(outputFolderId, deptName);
        folderCache.set(deptName, deptFolder);
      }
      const newFile = copyFileToFolder_(templateFileId, deptFolder.id, fileName);

      // 対象ページ以外を削除
      keepOnlyPage_(newFile.id, pageNumber);

      // プレースホルダ置換
      replacePlaceholders_(newFile.id, row);

      // 共有招待（Drive Advanced Service）
      shareWithNotification_(newFile.id, [email]);

      updates.push({
        row: rowNumber,
        url: (newFile.url || (newFile.getUrl && newFile.getUrl()) || ''),
        status: STATUS.DONE,
      });
    } catch (err) {
      const message = err && err.message ? err.message : String(err);
      updates.push({
        row: rowNumber,
        url: row[EMP_HEADERS.URL] || '',
        status: STATUS.ERROR_PREFIX + message,
      });
      Logger.log(`Row ${rowNumber} error: ${message}`);
    }
  }

  // バッチ書き戻し
  applyUpdates_(sheet, updates, data.colIndex);
}

// ========= ユーティリティ =========

// Config シートの Key-Value を読み込み
function loadConfig_(ss) {
  const sheet = ss.getSheetByName(SHEET.CONFIG);
  if (!sheet) throw new Error(`シートが見つかりません: ${SHEET.CONFIG}`);
  const values = sheet.getDataRange().getDisplayValues();
  const config = {};
  for (let i = 0; i < values.length; i++) {
    const rawKey = String(values[i][0] || '').trim();
    const val = String(values[i][1] || '').trim();
    if (!rawKey) continue;
    const normKey = normalizeKey_(rawKey);
    // オリジナルキーと正規化キーの両方で参照可能にする
    config[rawKey] = val;
    config[normKey] = val;
  }
  return config;
}

function normalizeKey_(s) {
  // 大文字化し、英数字以外を除去（例: "OUTPUT_FOLDER_ID" → "OUTPUTFOLDERID"、"出力先フォルダ ID" → "ID"部以外除外）
  return String(s).toUpperCase().replace(/[^A-Z0-9]/g, '');
}

function getConfigValue_(config, keyLiteral) {
  return config[keyLiteral] || config[normalizeKey_((keyLiteral))] || '';
}

// TemplateMapping を Map<部署名, ページ番号(1始まり)> として読み込み
function loadTemplateMapping_(ss) {
  const sheet = ss.getSheetByName(SHEET.TEMPLATE_MAPPING);
  if (!sheet) throw new Error(`シートが見つかりません: ${SHEET.TEMPLATE_MAPPING}`);
  const values = sheet.getDataRange().getValues();
  const map = new Map();
  for (let i = 1; i < values.length; i++) {
    const dept = String(values[i][0] || '').trim();
    const pageNumRaw = values[i][1];
    const pageNum = Number(pageNumRaw);
    if (dept && pageNum && isFinite(pageNum)) map.set(dept, Math.floor(pageNum));
  }
  return map;
}

// シートからオブジェクト配列を構築（ヘッダ行ベース）
function getSheetObjects_(sheet) {
  const values = sheet.getDataRange().getValues();
  if (values.length === 0) return { rows: [], headers: [], colIndex: {} };
  const headers = values[0].map(h => String(h || '').trim());
  const colIndex = headers.reduce((acc, h, i) => (acc[h] = i, acc), {});
  const rows = [];
  for (let r = 1; r < values.length; r++) {
    const obj = {};
    for (let c = 0; c < headers.length; c++) {
      obj[headers[c]] = values[r][c];
    }
    obj.__rowNumber = r + 1; // シート行番号（1始まり）
    rows.push(obj);
  }
  return { rows, headers, colIndex };
}

// ファイル名生成: 例) 会議背景_氏名_部署
function buildFileName_(row) {
  const name = String(row[EMP_HEADERS.NAME] || '').trim();
  const dept = String(row[EMP_HEADERS.DEPT] || '').trim();
  return `会議背景_${name || 'NO_NAME'}_${dept || 'NO_DEPT'}`;
}

// プレースホルダ置換
// スライド側の想定プレースホルダ: {{NAME}}, {{DEPARTMENT}}, {{TITLE}}, {{EMAIL}}
function replacePlaceholders_(presentationId, row) {
  const pres = SlidesApp.openById(presentationId);
  const replacements = {
    '{{NAME}}': String(row[EMP_HEADERS.NAME] || ''),
    '{{NAME_EN}}': String(row[EMP_HEADERS.NAME_EN] || ''),
    '{{DEPARTMENT}}': String(row[EMP_HEADERS.DEPT] || ''),
    '{{TITLE}}': String(row[EMP_HEADERS.TITLE] || ''),
    '{{EMAIL}}': String(row[EMP_HEADERS.EMAIL] || ''),
  };
  Object.keys(replacements).forEach(k => {
    pres.replaceAllText(k, replacements[k]);
  });
  pres.saveAndClose();
}

// 指定ページ（1-based）以外のスライドを削除
function keepOnlyPage_(presentationId, pageNumber) {
  const pres = SlidesApp.openById(presentationId);
  const slides = pres.getSlides();
  const targetIndex = Math.max(1, Math.floor(pageNumber)) - 1; // 0-based
  if (targetIndex < 0 || targetIndex >= slides.length) {
    throw new Error(`テンプレートページ番号が不正です: ${pageNumber}`);
  }
  for (let i = slides.length - 1; i >= 0; i--) {
    if (i !== targetIndex) slides[i].remove();
  }
  pres.saveAndClose();
}

// Advanced Drive: 共有ドライブ対応のファイルコピー（親フォルダ直下に作成）
function copyFileToFolder_(srcFileId, destFolderId, newName) {
  const resource = {
    name: newName,
    parents: [destFolderId],
  };
  const file = Drive.Files.copy(resource, srcFileId, {
    supportsAllDrives: true,
    fields: 'id,webViewLink',
  });
  return { id: file.id, url: file.webViewLink || '' };
}

// 共有ドライブへの移動は Drive API v2 の update/addParents/removeParents を使うが
// Apps Script では mediaData 引数との混同でエラーになりやすいため、
// 本スクリプトでは「copy → 元をゴミ箱へ」で代替しています。

// （不要になったため未使用）

// 共有招待（Drive Advanced Service 必須）
// emails: string[] 複数宛先にも対応
function shareWithNotification_(fileId, emails) {
  // 共有メールに同梱する依頼メッセージ
  const message = [
    '以下の手順でご対応をお願いします。',
    '',
    '1. スライドにアクセスし、ご自身の「所属」「肩書き」「氏名」等が正しいか確認してください。',
    '2. 問題なければ「ファイル > ダウンロード > PNG 画像（現在のスライド）」からPNGをダウンロードしてください。',
    '3. ダウンロードした画像をWeb会議の背景（Zoom / Google Meet）に設定してください。',
  ].join('\n');

  emails.forEach(email => {
    const resource = {
      role: 'writer', // 編集権限に設定（閲覧のみは 'reader'）
      type: 'user',
      emailAddress: email,
    };
    // Advanced Drive Service: Drive.Permissions.create
    Drive.Permissions.create(resource, fileId, {
      sendNotificationEmail: true,
      emailMessage: message,
      supportsAllDrives: true,
    });
  });
}

// バッチ書き戻し
function applyUpdates_(sheet, updates, colIndex) {
  if (!updates.length) return;
  const rows = updates.map(u => u.row);
  const minRow = Math.min.apply(null, rows);
  const maxRow = Math.max.apply(null, rows);

  // 対象範囲を読み出して上書き（URL, ステータスのみ）
  const headerRow = 1;
  const startRow = minRow;
  const numRows = maxRow - minRow + 1;
  const numCols = sheet.getLastColumn();
  const range = sheet.getRange(startRow, 1, numRows, numCols);
  const values = range.getValues();

  const urlCol = (colIndex[EMP_HEADERS.URL] || 0) + 1; // 1-based
  const statusCol = (colIndex[EMP_HEADERS.STATUS] || 0) + 1;

  const byRow = new Map(updates.map(u => [u.row, u]));
  for (let r = 0; r < numRows; r++) {
    const realRow = startRow + r;
    const upd = byRow.get(realRow);
    if (!upd) continue;
    if (urlCol > 0) values[r][urlCol - 1] = upd.url || '';
    if (statusCol > 0) values[r][statusCol - 1] = upd.status || '';
  }

  range.setValues(values);
}

function isValidEmail_(s) {
  if (!s) return false;
  // シンプル判定（組織のドメイン要件などがあれば適宜強化）
  const re = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return re.test(String(s));
}

// Employees は表示値（書式適用済の文字列）を利用して読み込み
function getSheetObjectsDisplay_(sheet) {
  const values = sheet.getDataRange().getDisplayValues();
  if (values.length === 0) return { rows: [], headers: [], colIndex: {} };
  const headers = values[0].map(h => String(h || '').trim());
  const colIndex = headers.reduce((acc, h, i) => (acc[h] = i, acc), {});
  const rows = [];
  for (let r = 1; r < values.length; r++) {
    const obj = {};
    for (let c = 0; c < headers.length; c++) {
      obj[headers[c]] = values[r][c];
    }
    obj.__rowNumber = r + 1;
    rows.push(obj);
  }
  return { rows, headers, colIndex };
}

// ========== セットアップ（シート作成 + デモ用テンプレ/データ作成） ==========

function setupSpreadsheetAndDemoAssets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // シート準備（存在しなければ作成、あればクリア）
  const configSheet = ensureSheet_(ss, SHEET.CONFIG);
  const mappingSheet = ensureSheet_(ss, SHEET.TEMPLATE_MAPPING);
  const employeesSheet = ensureSheet_(ss, SHEET.EMPLOYEES);

  clearAndSetHeader_(configSheet, ['Key', 'Value']);
  clearAndSetHeader_(mappingSheet, ['部署名', 'テンプレートページ番号']);
  clearAndSetHeader_(employeesSheet, [
    EMP_HEADERS.ID,
    EMP_HEADERS.NAME,
    EMP_HEADERS.NAME_EN,
    EMP_HEADERS.DEPT,
    EMP_HEADERS.TITLE,
    EMP_HEADERS.EMAIL,
    EMP_HEADERS.STATUS,
    EMP_HEADERS.URL,
  ]);

  // 共有ドライブのルートIDは不要。既存の出力フォルダIDを直接設定して利用します。

  // 出力フォルダIDを取得（未設定ならプロンプト）。既存の共有ドライブ内フォルダを指定してください。
  let outputFolderId = getConfigValue_(loadConfig_(ss), CONFIG_KEY.OUTPUT_FOLDER_ID);
  if (!outputFolderId) {
    const ui = SpreadsheetApp.getUi();
    const res = ui.prompt('出力先フォルダIDを入力', '共有ドライブ内の既存フォルダIDを入力してください（OUTPUT_FOLDER_ID に保存します）。', ui.ButtonSet.OK_CANCEL);
    if (res.getSelectedButton() !== ui.Button.OK) {
      ui.alert('キャンセルされました。先に Config に OUTPUT_FOLDER_ID を設定してください。');
      return;
    }
    outputFolderId = res.getResponseText().trim();
    if (!outputFolderId) {
      ui.alert('フォルダIDが空です。処理を中断します。');
      return;
    }
    // 存在チェック
    try {
      Drive.Files.get(outputFolderId, {
        supportsAllDrives: true,
      });
    } catch (e) {
      ui.alert('指定のフォルダIDが無効またはアクセス権がありません。');
      return;
    }
    configSheet.getRange(2, 1, 1, 2).setValues([[CONFIG_KEY.OUTPUT_FOLDER_ID, outputFolderId]]);
  }

  // 既存テンプレートIDの取得（未設定ならプロンプト）
  let templateFileId = getConfigValue_(loadConfig_(ss), CONFIG_KEY.TEMPLATE_FILE_ID);
  if (!templateFileId) {
    const ui = SpreadsheetApp.getUi();
    const res = ui.prompt('テンプレートのファイルIDを入力', '単一ファイルに部署ごとのページを持つGoogleスライドのID（TEMPLATE_FILE_ID）を入力してください。', ui.ButtonSet.OK_CANCEL);
    if (res.getSelectedButton() !== ui.Button.OK) {
      ui.alert('キャンセルされました。先に Config に TEMPLATE_FILE_ID を設定してください。');
      return;
    }
    templateFileId = res.getResponseText().trim();
    if (!templateFileId) {
      ui.alert('テンプレートIDが空です。処理を中断します。');
      return;
    }
    try {
      const f = Drive.Files.get(templateFileId, {
        supportsAllDrives: true,
      });
      // 可能なら形式チェック（Slides）
      if (f.mimeType && f.mimeType !== 'application/vnd.google-apps.presentation') {
        ui.alert('指定のファイルはGoogleスライドではない可能性があります。続行しますがご確認ください。');
      }
    } catch (e) {
      ui.alert('指定のテンプレートIDが無効またはアクセス権がありません。');
      return;
    }
    configSheet.getRange(3, 1, 1, 2).setValues([[CONFIG_KEY.TEMPLATE_FILE_ID, templateFileId]]);
  }
  // マッピングはページ番号
  mappingSheet.getRange(2, 1, 2, 2).setValues([
    ['営業部', 1],
    ['開発部', 2],
  ]);

  // ダミー社員データ
  const demoRows = [
    ['E001', '山田 太郎', 'Taro Yamada', '営業部', 'マネージャー', 'taro.yamada@example.com', STATUS.PENDING, ''],
    ['E002', '佐藤 花子', 'Hanako Sato', '開発部', 'エンジニア', 'hanako.sato@example.com', STATUS.PENDING, ''],
    ['E003', '鈴木 次郎', 'Jiro Suzuki', '開発部', 'リード', 'jiro.suzuki@example.com', STATUS.PENDING, ''],
  ];
  employeesSheet.getRange(2, 1, demoRows.length, demoRows[0].length).setValues(demoRows);

  SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert('初期セットアップが完了しました。Config/TemplateMapping/Employees を確認してください。\n' +
    'OUTPUT_FOLDER_ID: ' + outputFolderId + '\n' +
    'TEMPLATE_FILE_ID: ' + (templateFileId || getConfigValue_(loadConfig_(ss), CONFIG_KEY.TEMPLATE_FILE_ID)) + '\n' +
    '注意: TemplateMapping のページ番号はテンプレ内のスライド順（1始まり）に合わせてください。');
}

// シート存在保証
function ensureSheet_(ss, name) {
  const s = ss.getSheetByName(name);
  return s || ss.insertSheet(name);
}

// ヘッダ設定（既存データはクリア）
function clearAndSetHeader_(sheet, headers) {
  sheet.clear();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.setFrozenRows(1);
}

// デモ用テンプレート作成（単一ファイルに複数ページ）
function createDemoTemplateMulti_(deptList) {
  const title = `会議背景テンプレート_共通_${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss')}`;
  const pres = SlidesApp.create(title);

  const addDeptSlide = (slide, deptName) => {
    slide.getPageElements().forEach(e => e.remove());
    slide.insertTextBox('Welcome to the Meeting', 50, 60, 640, 60)
      .getText().getTextStyle().setFontSize(28);

    const placeholders = [
      '氏名: {{NAME}}',
      '氏名(英字): {{NAME_EN}}',
      '部署: {{DEPARTMENT}}',
      '役職: {{TITLE}}',
      'メール: {{EMAIL}}',
    ];
    let top = 140;
    placeholders.forEach(t => {
      slide.insertTextBox(t, 60, top, 620, 36).getText().getTextStyle().setFontSize(18);
      top += 48;
    });

    slide.insertTextBox(`Template: ${deptName}`, 60, 320, 620, 30)
      .getText().getTextStyle().setFontSize(12).setForegroundColor('#777777');
  };

  const first = pres.getSlides()[0];
  addDeptSlide(first, deptList[0] || '部署A');
  for (let i = 1; i < deptList.length; i++) {
    const s = pres.appendSlide(SlidesApp.PredefinedLayout.BLANK);
    addDeptSlide(s, deptList[i]);
  }

  pres.saveAndClose();
  const presId = pres.getId();
  // URLは移動後に変わらないが、ここではIDのみ返す
  return { id: presId, url: SlidesApp.openById(presId).getUrl(), name: title };
}
// Advanced Drive: 親フォルダ配下に特定名称のサブフォルダを取得 or 作成
function getOrCreateSubFolder_(parentFolderId, name) {
  const q = "mimeType = 'application/vnd.google-apps.folder' and trashed = false and name = '" + name.replace(/'/g, "\\'") + "' and '" + parentFolderId + "' in parents";
  const res = Drive.Files.list({
    q: q,
    pageSize: 1,
    includeItemsFromAllDrives: true,
    supportsAllDrives: true,
  });
  if (res && res.files && res.files.length > 0) {
    const f = res.files[0];
    return { id: f.id, url: (f.webViewLink || f.alternateLink || '') };
  }
  // 作成
  const resource = {
    name: name,
    mimeType: 'application/vnd.google-apps.folder',
    parents: [parentFolderId],
  };
  const folder = Drive.Files.create(resource, null, {
    supportsAllDrives: true,
  });
  return { id: folder.id, url: folder.webViewLink || folder.alternateLink || '' };
}
