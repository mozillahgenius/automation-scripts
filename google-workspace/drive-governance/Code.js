/**
 * Drive 総合管理システム（完全版）
 * =============================================================================
 * PDFガイドに基づく包括的なGoogle Drive管理システム
 * 
 * 【主要機能】
 * 1) ファイル・フォルダの命名規則チェック
 * 2) ファイル更新リマインド機能
 * 3) 重複ファイル検知
 * 4) 3つのドライブ（Daily/External-Share/Archive）の管理
 * 5) 権限の自動失効機能
 * 6) DriveMembers/MemberDashboardによる権限の可視化
 * 7) zzzz_ファイルの自動アーカイブ移動
 * 8) 外部共有の安全管理機能
 * 9) アクセス権限の定期監査レポート
 * 
 * 【導入手順】
 * 1. Apps Script エディタで「サービス」→「Drive API」を追加
 * 2. このスクリプトに紐づくGCPプロジェクトで Drive API を有効化
 * 3. Admin SDK Directory API も有効化（メンバー管理用）
 * 4. スクリプト実行者が、対象の共有ドライブ全ての「管理者」であることを確認
 * 5. 定数設定後、一度 setupExtended() を実行しシートを初期化
 * 6. mainCheckExtended() をトリガー登録（毎日実行推奨）
 */

// ===== 設定定数 =====
const CONFIG_SHEET_ID   = 'SPREADSHEET_ID_PLACEHOLDER_1';  // 管理用スプレッドシートID
const CONFIG_SHEET_NAME = 'FolderConfig';
const DUP_SHEET_NAME    = 'DuplicateScan';
const VIOL_SHEET_NAME   = 'RuleViolations';
const REMINDER_SHEET_NAME = 'FileReminders';

// 新規追加シート
const DRIVE_CONFIG_SHEET_NAME = 'DriveConfig';
const DRIVE_MEMBERS_SHEET_NAME = 'DriveMembers';
const MEMBER_DASHBOARD_SHEET_NAME = 'MemberDashboard';
const AUDIT_LOG_SHEET_NAME = 'AuditLog';
const EXTERNAL_SHARE_SHEET_NAME = 'ExternalShareLog';

const ADMIN_EMAIL       = 'automation@example-company.test';

// ドライブの種類
const DRIVE_TYPES = {
  DAILY: 'Daily',
  EXTERNAL: 'External-Share',
  ARCHIVE: 'Archive'
};

// 権限の自動失効期間（日数）
const PERMISSION_EXPIRY_DAYS = {
  CONTRIBUTOR: 90,  // 編集者は90日
  VIEWER: 30        // 閲覧者は30日
};

/**
 * 拡張版メイン関数：すべての管理機能を実行
 */
function mainCheckExtended() {
  try {
    Logger.log('=== Drive総合管理チェック開始 ===');
    
    // 既存の機能
    scanForDuplicates();
    scanNamingViolations();
    sendFileReviewReminders();
    
    // 新規機能
    moveZzzzFilesToArchive();
    revokeExpiredPermissions();
    updateDriveMembersList();
    updateMemberDashboard();
    
    Logger.log('=== Drive総合管理チェック完了 ===');
  } catch (error) {
    Logger.log('致命的なエラー発生: ' + error.stack);
    try {
      MailApp.sendEmail({
        to: ADMIN_EMAIL,
        subject: '【Drive管理】エラー発生',
        body: 'エラーが発生しました:\n' + error.toString() + '\n' + error.stack
      });
    } catch (e) {
      Logger.log('エラー通知メールの送信に失敗しました: ' + e);
    }
  }
}

/**
 * 旧バージョンとの互換性のためのメイン関数
 */
function mainCheck() {
  mainCheckExtended();
}

/**
 * 拡張版セットアップ関数：すべての管理シートを作成
 */
function setupExtended() {
  const ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
  
  // すべてのシート設定
  const allSheets = [
    { 
      name: CONFIG_SHEET_NAME, 
      headers: ['監視対象フォルダID','通知先メール','リマインド閾値（日）','リマインド頻度（日）','最終リマインド送信日時','最終チェック日時','備考'] 
    },
    { 
      name: DUP_SHEET_NAME,    
      headers: ['スキャン日時','ファイルID','ファイル名','フォルダパス','サイズ(byte)','MD5ハッシュ','最終更新日'] 
    },
    { 
      name: VIOL_SHEET_NAME,   
      headers: ['検知日時','パス','種別','名前','ルール種別','詳細','URL'] 
    },
    { 
      name: REMINDER_SHEET_NAME, 
      headers: ['送信日時','対象フォルダ','ファイル名','最終更新日','経過日数','ファイルパス','ファイルURL','送信先メール'] 
    },
    {
      name: DRIVE_CONFIG_SHEET_NAME,
      headers: ['ドライブ種別','ドライブID','ドライブ名','Gemini検索','備考']
    },
    {
      name: DRIVE_MEMBERS_SHEET_NAME,
      headers: ['更新日時','ドライブID','ドライブ名','メールアドレス','権限タイプ','最終アクセス','有効期限','外部ドメイン']
    },
    {
      name: MEMBER_DASHBOARD_SHEET_NAME,
      headers: ['更新日時','サマリー'] // ピボットテーブル用
    },
    {
      name: AUDIT_LOG_SHEET_NAME,
      headers: ['監査日時','ドライブ名','対象者','権限','アクション','理由','実行者']
    },
    {
      name: EXTERNAL_SHARE_SHEET_NAME,
      headers: ['共有日時','ファイルID','ファイル名','共有先メール','権限','有効期限','承認者','備考']
    }
  ];

  allSheets.forEach(function(c) {
    let sh = ss.getSheetByName(c.name);
    if (!sh) {
      sh = ss.insertSheet(c.name);
      Logger.log('シート作成: ' + c.name);
    }
    
    if (sh.getLastRow() === 0) {
      sh.clear();
      sh.appendRow(c.headers).setFrozenRows(1);
      sh.getRange("A1:" + String.fromCharCode(64 + c.headers.length) + "1").setFontWeight("bold");
      Logger.log('ヘッダー設定: ' + c.name);
    }
  });

  // サンプルデータの追加
  const cfgSheet = ss.getSheetByName(CONFIG_SHEET_NAME);
  if (cfgSheet.getLastRow() === 1) {
    cfgSheet.appendRow([
      '監視対象フォルダIDを入力', 
      'manager@example.com',
      60,
      7,
      '',
      '',
      'サンプル行（削除してください）'
    ]);
  }
  
  // ドライブ設定のサンプル行を追加
  const driveConfigSheet = ss.getSheetByName(DRIVE_CONFIG_SHEET_NAME);
  if (driveConfigSheet.getLastRow() === 1) {
    driveConfigSheet.appendRow([DRIVE_TYPES.DAILY, 'ドライブIDを入力', 'Daily業務用', 'はい', '日常業務用']);
    driveConfigSheet.appendRow([DRIVE_TYPES.EXTERNAL, 'ドライブIDを入力', 'External共有用', 'はい', '外部共有専用']);
    driveConfigSheet.appendRow([DRIVE_TYPES.ARCHIVE, 'ドライブIDを入力', 'Archiveアーカイブ', 'いいえ', '長期保管用']);
  }

  Logger.log('拡張セットアップが完了しました。');
  Logger.log('DriveConfigシートで3つのドライブIDを設定してください。');
  Logger.log('FolderConfigシートで監視対象フォルダを設定してください。');
}

/**
 * 旧バージョンとの互換性のためのセットアップ関数
 */
function setup() {
  setupExtended();
}

/**
 * フォルダ設定を取得します。
 */
function _getFolderConfigs(configSheet) {
  try {
    const data = configSheet.getDataRange().getValues();
    const configs = [];
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const folderId = row[0];
      const notifyEmail = row[1];
      const reminderThreshold = row[2];
      const reminderFrequency = row[3];
      const lastReminderDate = row[4];
      const lastCheckDate = row[5];
      const remarks = row[6];
      
      if (folderId && 
          folderId !== '監視対象フォルダIDを入力' && 
          reminderThreshold && 
          !isNaN(reminderThreshold)) {
        configs.push({
          folderId: folderId,
          notifyEmail: notifyEmail || ADMIN_EMAIL,
          reminderThreshold: parseInt(reminderThreshold) || 60,
          reminderFrequency: parseInt(reminderFrequency) || 7,
          lastReminderDate: lastReminderDate,
          lastCheckDate: lastCheckDate,
          remarks: remarks,
          rowIndex: i + 1
        });
      }
    }
    
    Logger.log('監視対象フォルダ数: ' + configs.length + '個');
    return configs;
  } catch (e) {
    Logger.log('フォルダ設定の取得でエラー: ' + e);
    return [];
  }
}

/**
 * 監視対象フォルダのIDリストを取得します（チェックが必要なフォルダのみ）。
 */
function _getTargetFolderIds(configSheet) {
  try {
    const folderConfigs = _getFolderConfigs(configSheet);
    const targetFolderIds = [];
    const now = new Date();
    
    folderConfigs.forEach(function(config) {
      // リマインド頻度に基づいてチェックが必要かどうかを判定
      const shouldCheck = _shouldCheckFolder(config.lastCheckDate, config.reminderFrequency, now);
      
      if (shouldCheck) {
        // 同じフォルダIDが既に追加されていない場合のみ追加
        if (targetFolderIds.indexOf(config.folderId) === -1) {
          targetFolderIds.push(config.folderId);
        }
      } else {
        Logger.log('チェックスキップ: ' + config.folderId + ' (次回チェックまであと' + 
                   _getDaysUntilNextCheck(config.lastCheckDate, config.reminderFrequency, now) + '日)');
      }
    });
    
    Logger.log('チェック対象フォルダ数: ' + targetFolderIds.length + '個（全' + folderConfigs.length + '個中）');
    return targetFolderIds;
  } catch (e) {
    Logger.log('監視対象フォルダIDの取得でエラー: ' + e);
    return [];
  }
}

/**
 * フォルダのチェックが必要かどうかを判定します。
 */
function _shouldCheckFolder(lastCheckDate, reminderFrequency, now) {
  if (!lastCheckDate) {
    return true; // 一度もチェックしていない場合は必ずチェック
  }
  
  const lastDate = lastCheckDate instanceof Date ? lastCheckDate : new Date(lastCheckDate);
  const daysSinceLastCheck = (now.getTime() - lastDate.getTime()) / (24 * 60 * 60 * 1000);
  
  return daysSinceLastCheck >= reminderFrequency;
}

/**
 * 次回チェックまでの日数を計算します。
 */
function _getDaysUntilNextCheck(lastCheckDate, reminderFrequency, now) {
  if (!lastCheckDate) return 0;
  
  const lastDate = lastCheckDate instanceof Date ? lastCheckDate : new Date(lastCheckDate);
  const daysSinceLastCheck = (now.getTime() - lastDate.getTime()) / (24 * 60 * 60 * 1000);
  const daysUntilNext = reminderFrequency - daysSinceLastCheck;
  
  return Math.max(0, Math.ceil(daysUntilNext));
}

/**
 * 重複ファイルをスキャンして記録します。
 */
function scanForDuplicates() {
  const ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
  const configSheet = ss.getSheetByName(CONFIG_SHEET_NAME);
  const sht = ss.getSheetByName(DUP_SHEET_NAME);
  
  if (sht.getLastRow() > 1) {
    sht.getRange(2, 1, sht.getMaxRows() - 1, 7).clearContent();
  }
  
  // チェックが必要なフォルダのみを取得
  const targetFolderIds = _getTargetFolderIds(configSheet);
  if (targetFolderIds.length === 0) {
    Logger.log('チェックが必要なフォルダがありません。重複スキャンをスキップします。');
    return;
  }
  
  const rows = [];
  const now = new Date();
  const configs = _getFolderConfigs(configSheet);
  
  targetFolderIds.forEach(function(folderId, index) {
    Logger.log('重複スキャン (' + (index + 1) + '/' + targetFolderIds.length + '): ' + folderId);
    _collectFiles(folderId, '', now, rows);
    
    // このフォルダIDに関連するすべての設定の最終チェック日時を更新
    configs.forEach(function(config) {
      if (config.folderId === folderId) {
        configSheet.getRange(config.rowIndex, 6).setValue(now);
      }
    });
  });
  
  if (rows.length > 0) {
    sht.getRange(2, 1, rows.length, 7).setValues(rows);
    _highlightDuplicates(sht);
    Logger.log('重複スキャン完了: ' + rows.length + '件のファイルを記録');
  }
}

/**
 * 再帰的にファイルを収集し、情報を配列に追加します。
 */
function _collectFiles(folderId, path, ts, rows) {
  let folder;
  try {
     folder = DriveApp.getFolderById(folderId);
  } catch(e) {
    Logger.log('フォルダが見つかりません: ' + folderId);
    return;
  }

  const currentPath = path + '/' + folder.getName();
  const files = folder.getFiles();
  while (files.hasNext()) {
    const f = files.next();
    try {
      const fileDetails = Drive.Files.get(f.getId(), {
        fields: 'md5Checksum,size,modifiedTime',
        supportsAllDrives: true 
      });
      rows.push([
        ts, 
        f.getId(), 
        f.getName(), 
        currentPath, 
        fileDetails.size || 0, 
        fileDetails.md5Checksum || 'N/A', 
        new Date(fileDetails.modifiedTime)
      ]);
    } catch (e) { 
      rows.push([
        ts, 
        f.getId(), 
        f.getName(), 
        currentPath, 
        f.getSize(), 
        '取得エラー', 
        f.getLastUpdated()
      ]); 
    }
  }
  
  const subFolders = folder.getFolders(); 
  while (subFolders.hasNext()) {
    _collectFiles(subFolders.next().getId(), currentPath, ts, rows);
  }
}

/**
 * 重複ファイルの行をハイライトします。
 */
function _highlightDuplicates(sht) {
  const data = sht.getDataRange().getValues();
  const fileData = data.slice(1);
  const md5Map = {};
  
  fileData.forEach(function(row, index) { 
    const md5 = row[5];
    if (md5 && md5 !== 'N/A' && md5 !== '取得エラー') {
      if (!md5Map[md5]) {
        md5Map[md5] = [];
      }
      md5Map[md5].push(index + 2);
    }
  });
  
  if (sht.getLastRow() > 1) {
    sht.getRange(2, 1, sht.getLastRow() - 1, sht.getLastColumn()).setBackground(null);
  }
  
  for (const md5 in md5Map) {
    const rowNumbers = md5Map[md5];
    if (rowNumbers.length > 1) {
      rowNumbers.forEach(function(rowNum) {
        sht.getRange(rowNum, 1, 1, sht.getLastColumn()).setBackground('#FFF2CC');
      });
    }
  }
}

/**
 * 命名規則が有効かチェックします。
 * XX_（数字2桁）、YYYYMMDD_（日付8桁）、zzzz_（アーカイブ用）、keep_（永久保存）、X.（アルファベット1文字+ピリオド）の形式を許可
 */
function _isValidNamingConvention(name) {
  // パターン1: 数字2桁_
  const twoDigitPattern = /^\d{2}_.+/;
  
  // パターン2: YYYYMMDD_（日付8桁）
  const datePattern = /^\d{8}_.+/;
  
  // パターン3: zzzz_（アーカイブ用）
  const archivePattern = /^zzzz_.+/i;
  
  // パターン4: keep_（永久保存用）
  const keepPattern = /^keep_.+/i;
  
  // パターン5: アルファベット1文字+ピリオド
  const alphabetPattern = /^[a-zA-Z]\..+/;
  
  if (twoDigitPattern.test(name)) {
    return true;
  }
  
  if (archivePattern.test(name)) {
    return true;
  }
  
  if (keepPattern.test(name)) {
    return true;
  }
  
  if (alphabetPattern.test(name)) {
    return true;
  }
  
  if (datePattern.test(name)) {
    const dateStr = name.substring(0, 8);
    const year = parseInt(dateStr.substring(0, 4));
    const month = parseInt(dateStr.substring(4, 6));
    const day = parseInt(dateStr.substring(6, 8));
    
    if (year >= 1900 && year <= 2100 && 
        month >= 1 && month <= 12 && 
        day >= 1 && day <= 31) {
      return true;
    }
  }
  
  return false;
}

/**
 * ルール違反をスキャンして通知します。
 */
function scanNamingViolations() {
  const ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
  const configSheet = ss.getSheetByName(CONFIG_SHEET_NAME);
  const sht = ss.getSheetByName(VIOL_SHEET_NAME);
  
  if (sht.getLastRow() > 1) {
    sht.getRange(2, 1, sht.getMaxRows() - 1, 7).clearContent();
  }

  // チェックが必要なフォルダのみを取得
  const targetFolderIds = _getTargetFolderIds(configSheet);
  if (targetFolderIds.length === 0) {
    Logger.log('チェックが必要なフォルダがありません。ルール違反スキャンをスキップします。');
    return;
  }

  const rows = [];
  const now = new Date();
  const configs = _getFolderConfigs(configSheet);
  
  targetFolderIds.forEach(function(folderId, index) {
    Logger.log('ルール違反スキャン (' + (index + 1) + '/' + targetFolderIds.length + '): ' + folderId);
    _checkFolderRules(folderId, '', rows, now);
    
    // このフォルダIDに関連するすべての設定の最終チェック日時を更新
    configs.forEach(function(config) {
      if (config.folderId === folderId) {
        configSheet.getRange(config.rowIndex, 6).setValue(now);
      }
    });
  });
  
  if (rows.length > 0) {
    sht.getRange(2, 1, rows.length, 7).setValues(rows);
    try { 
      const emailBody = _createViolationEmailBody(rows);
      MailApp.sendEmail({
        to: ADMIN_EMAIL,
        subject: '【Drive通知】ルール違反が検知されました',
        htmlBody: emailBody
      });
    } catch (e) { 
      Logger.log('ルール違反通知メールの送信に失敗しました: ' + e); 
    }
    Logger.log('ルール違反検知: ' + rows.length + '件');
  } else {
    Logger.log('ルール違反なし');
  }
}

/**
 * フォルダ内の命名規則などをチェックします。
 */
function _checkFolderRules(folderId, path, rows, ts) {
  let folder;
  try { 
    folder = DriveApp.getFolderById(folderId); 
  } catch (error) { 
    Logger.log('フォルダ取得失敗: ' + folderId + ' - ' + error.toString());
    return; 
  }
  
  const fullPath = path + '/' + folder.getName();
  
  // ルートフォルダ以外の場合、命名規則をチェック
  if (path) { 
    if (!_isValidNamingConvention(folder.getName())) {
      rows.push([
        ts, 
        fullPath, 
        'フォルダ', 
        folder.getName(), 
        '命名規則違反', 
        'フォルダ名の先頭が許可された形式（数字2桁_ / YYYYMMDD_ / zzzz_ / keep_ / アルファベット.）ではありません。',
        folder.getUrl()
      ]); 
    }
  }

  // ファイルの命名規則チェック
  const files = folder.getFiles();
  while (files.hasNext()) {
    const file = files.next(); 
    const fileName = file.getName();
    
    // 命名規則チェック
    if (!_isValidNamingConvention(fileName)) {
      rows.push([
        ts, 
        fullPath, 
        'ファイル', 
        fileName, 
        '命名規則違反', 
        'ファイル名の先頭が許可された形式（数字2桁_ / YYYYMMDD_ / zzzz_ / keep_ / アルファベット.）ではありません。',
        file.getUrl()
      ]);
    }
    
    // 禁止語チェック
    if (/関連|その他/i.test(fileName)) {
      rows.push([
        ts, 
        fullPath, 
        'ファイル', 
        fileName, 
        '命名規則違反', 
        '名前に禁止語（関連、その他）が含まれています。',
        file.getUrl()
      ]); 
    }
    
    // Microsoftファイル形式チェック
    const microsoftFileCheck = _checkMicrosoftFileFormat(fileName);
    if (microsoftFileCheck.isMicrosoftFile) {
      rows.push([
        ts, 
        fullPath, 
        'ファイル', 
        fileName, 
        'Microsoftファイル形式', 
        microsoftFileCheck.fileType + 'ファイルです。' + microsoftFileCheck.recommendation + 'に変換することを推奨します。',
        file.getUrl()
      ]);
    }
  }

  // サブフォルダの再帰チェック
  const subs = folder.getFolders();
  while (subs.hasNext()) {
    _checkFolderRules(subs.next().getId(), fullPath, rows, ts);
  }
}

/**
 * Microsoftファイル形式をチェックします。
 */
function _checkMicrosoftFileFormat(fileName) {
  const microsoftFormats = {
    '.doc': { type: 'Word文書', google: 'Googleドキュメント' },
    '.docx': { type: 'Word文書', google: 'Googleドキュメント' },
    '.docm': { type: 'Word文書（マクロ有効）', google: 'Googleドキュメント' },
    '.xls': { type: 'Excel表計算', google: 'Googleスプレッドシート' },
    '.xlsx': { type: 'Excel表計算', google: 'Googleスプレッドシート' },
    '.xlsm': { type: 'Excel表計算（マクロ有効）', google: 'Googleスプレッドシート' },
    '.xlsb': { type: 'Excel表計算（バイナリ）', google: 'Googleスプレッドシート' },
    '.ppt': { type: 'PowerPointプレゼンテーション', google: 'Googleスライド' },
    '.pptx': { type: 'PowerPointプレゼンテーション', google: 'Googleスライド' },
    '.pptm': { type: 'PowerPointプレゼンテーション（マクロ有効）', google: 'Googleスライド' },
    '.vsd': { type: 'Visio図面', google: 'Google図形描画' },
    '.vsdx': { type: 'Visio図面', google: 'Google図形描画' },
    '.pub': { type: 'Publisher文書', google: 'Googleドキュメント' },
    '.one': { type: 'OneNote', google: 'Google Keep' },
    '.mpp': { type: 'Microsoft Project', google: 'Googleスプレッドシート' },
    '.mdb': { type: 'Access データベース', google: 'Googleスプレッドシート' },
    '.accdb': { type: 'Access データベース', google: 'Googleスプレッドシート' }
  };
  
  const lowerFileName = fileName.toLowerCase();
  
  for (const extension in microsoftFormats) {
    if (lowerFileName.endsWith(extension)) {
      const info = microsoftFormats[extension];
      return {
        isMicrosoftFile: true,
        fileType: info.type,
        recommendation: info.google,
        extension: extension
      };
    }
  }
  
  return {
    isMicrosoftFile: false
  };
}

/**
 * ルール違反の詳細なメール本文を作成します。
 */
function _createViolationEmailBody(violationRows) {
  const spreadsheetUrl = 'https://docs.google.com/spreadsheets/d/' + CONFIG_SHEET_ID + '/edit';
  
  let html = '<h2>🚨 Google Drive ルール違反レポート</h2>';
  html += '<p><strong>検知日時:</strong> ' + new Date().toLocaleString('ja-JP') + '</p>';
  html += '<p><strong>検知件数:</strong> ' + violationRows.length + '件</p>';
  
  html += '<div style="margin: 20px 0;">';
  html += '<a href="' + spreadsheetUrl + '" target="_blank" style="background-color: #1a73e8; color: white; padding: 12px 24px; text-decoration: none; border-radius: 4px; display: inline-block; font-weight: bold;">📊 スプレッドシートで詳細を確認</a>';
  html += '</div>';
  html += '<p style="color: #666; font-size: 12px;">※スプレッドシートを開いた後、「' + VIOL_SHEET_NAME + '」シートをご確認ください。</p>';
  
  html += '<h3>📋 違反詳細一覧</h3>';
  html += '<table border="1" cellpadding="8" cellspacing="0" style="border-collapse: collapse; width: 100%; font-family: Arial, sans-serif;">';
  html += '<thead style="background-color: #f2f2f2;">';
  html += '<tr><th>種別</th><th>ファイル/フォルダ名</th><th>ルール違反内容</th><th>リンク</th></tr>';
  html += '</thead><tbody>';
  
  violationRows.forEach(function(row) {
    const timestamp = row[0];
    const path = row[1];
    const type = row[2];
    const name = row[3];
    const ruleType = row[4];
    const details = row[5];
    const url = row[6];
    const rowColor = _getViolationRowColor(ruleType);
    
    html += '<tr style="background-color: ' + rowColor + ';">';
    html += '<td>' + type + '</td>';
    html += '<td style="font-weight: bold;">' + name + '</td>';
    html += '<td>' + ruleType + '<br><small style="color: #666;">' + details + '</small></td>';
    html += '<td><a href="' + url + '" target="_blank" style="color: #1a73e8; text-decoration: none;">📂 開く</a></td>';
    html += '</tr>';
  });
  
  html += '</tbody></table>';
  html += '<h3>📊 違反種別サマリー</h3>';
  html += _createViolationSummary(violationRows);
  html += '<hr>';
  html += '<p style="color: #666; font-size: 12px;">';
  html += '詳細は<a href="' + spreadsheetUrl + '" target="_blank">スプレッドシート</a>でもご確認いただけます。<br>';
  html += 'このメールは自動送信されています。';
  html += '</p>';
  
  return html;
}

/**
 * 違反種別に応じた行の背景色を返します。
 */
function _getViolationRowColor(ruleType) {
  switch (ruleType) {
    case '命名規則違反': return '#ffebee';
    case 'Microsoftファイル形式': return '#e8f5e8';
    default: return '#f5f5f5';
  }
}

/**
 * 違反種別のサマリーを作成します。
 */
function _createViolationSummary(violationRows) {
  const summary = {};
  violationRows.forEach(function(row) {
    const ruleType = row[4];
    summary[ruleType] = (summary[ruleType] || 0) + 1;
  });
  
  let html = '<ul>';
  for (const rule in summary) {
    const count = summary[rule];
    html += '<li><strong>' + rule + ':</strong> ' + count + '件</li>';
  }
  html += '</ul>';
  
  return html;
}

/**
 * ファイル更新確認リマインドメールを送信します。
 */
function sendFileReviewReminders() {
  const ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
  const configSheet = ss.getSheetByName(CONFIG_SHEET_NAME);
  const reminderSheet = ss.getSheetByName(REMINDER_SHEET_NAME);
  
  const folderConfigs = _getFolderConfigs(configSheet);
  if (folderConfigs.length === 0) {
    Logger.log('フォルダ設定が見つかりません。リマインド確認をスキップします。');
    return;
  }
  
  const now = new Date();
  
  folderConfigs.forEach(function(config, index) {
    Logger.log('リマインド確認処理 (' + (index + 1) + '/' + folderConfigs.length + '): ' + config.folderId);
    
    const folderId = config.folderId;
    const reminderThreshold = config.reminderThreshold;
    const reminderFrequency = config.reminderFrequency;
    const lastReminderDate = config.lastReminderDate;
    const notifyEmail = config.notifyEmail;
    const rowIndex = config.rowIndex;
    
    const shouldSendReminder = _shouldSendReminderEmail(lastReminderDate, reminderFrequency, now);
    
    if (!shouldSendReminder) {
      Logger.log('リマインド頻度未到達: ' + folderId);
      configSheet.getRange(rowIndex, 6).setValue(now);
      return;
    }
    
    try {
      const folder = DriveApp.getFolderById(folderId);
      Logger.log('フォルダ取得成功: ' + folder.getName());
      
      const reminderFiles = [];
      const reminderCutoff = new Date(now.getTime() - reminderThreshold * 24 * 60 * 60 * 1000);
      _collectReminderFiles(folder, '', reminderCutoff, reminderFiles);
      
      if (reminderFiles.length > 0) {
        const emailSent = _sendReminderEmail(folder, reminderFiles, notifyEmail, reminderThreshold);
        
        if (emailSent) {
          configSheet.getRange(rowIndex, 5).setValue(now);
          configSheet.getRange(rowIndex, 6).setValue(now);
          
          reminderFiles.forEach(function(file) {
            const daysSinceUpdate = Math.floor((now.getTime() - file.lastUpdated.getTime()) / (24 * 60 * 60 * 1000));
            reminderSheet.appendRow([
              now,
              folder.getName(),
              file.name,
              file.lastUpdated,
              daysSinceUpdate,
              file.path,
              file.url,
              notifyEmail
            ]);
          });
          
          Logger.log('リマインドメール送信成功: ' + folder.getName() + ' (' + reminderFiles.length + '件) → ' + notifyEmail);
        } else {
          Logger.log('リマインドメール送信失敗: ' + folder.getName() + ' → ' + notifyEmail);
        }
      } else {
        Logger.log('リマインド対象ファイルなし: ' + folder.getName());
        configSheet.getRange(rowIndex, 6).setValue(now);
      }
      
    } catch (error) {
      Logger.log('リマインド処理エラー: ' + folderId + ' - ' + error.toString());
    }
  });
}

/**
 * リマインドメールを送信する必要があるかチェックします。
 */
function _shouldSendReminderEmail(lastReminderDate, reminderFrequency, now) {
  if (!lastReminderDate) {
    return true;
  }
  
  const lastDate = lastReminderDate instanceof Date ? lastReminderDate : new Date(lastReminderDate);
  const daysSinceLastReminder = (now.getTime() - lastDate.getTime()) / (24 * 60 * 60 * 1000);
  
  return daysSinceLastReminder >= reminderFrequency;
}

/**
 * リマインド対象ファイルを再帰的に収集します。
 */
function _collectReminderFiles(folder, path, cutoff, reminderFiles) {
  try {
    const fullPath = path + '/' + folder.getName();
    
    const files = folder.getFiles();
    while (files.hasNext()) {
      const file = files.next();
      const fileName = file.getName();
      if (!fileName.startsWith('keep_') && !fileName.toLowerCase().startsWith('zzzz_') && file.getLastUpdated() < cutoff) {
        reminderFiles.push({
          name: fileName,
          lastUpdated: file.getLastUpdated(),
          url: file.getUrl(),
          path: fullPath,
          size: file.getSize()
        });
      }
    }
    
    const subFolders = folder.getFolders();
    while (subFolders.hasNext()) {
      _collectReminderFiles(subFolders.next(), fullPath, cutoff, reminderFiles);
    }
  } catch (error) {
    Logger.log('リマインドファイル収集エラー: ' + folder.getName() + ' - ' + error.toString());
  }
}

/**
 * リマインド確認メールを送信します。
 */
function _sendReminderEmail(folder, reminderFiles, notifyEmail, reminderThreshold) {
  try {
    const spreadsheetUrl = 'https://docs.google.com/spreadsheets/d/' + CONFIG_SHEET_ID + '/edit';
    const subject = '【Drive確認】' + folder.getName() + 'フォルダのファイル更新確認（' + reminderThreshold + '日以上未更新）';
    
    let html = '<h2>📋 Google Drive ファイル更新確認のお願い</h2>';
    html += '<p><strong>対象フォルダ:</strong> ' + folder.getName() + '</p>';
    html += '<p><strong>確認日時:</strong> ' + new Date().toLocaleString('ja-JP') + '</p>';
    html += '<p><strong>対象ファイル数:</strong> ' + reminderFiles.length + '件</p>';
    
    html += '<div style="margin: 20px 0;">';
    html += '<a href="' + spreadsheetUrl + '" target="_blank" style="background-color: #1a73e8; color: white; padding: 12px 24px; text-decoration: none; border-radius: 4px; display: inline-block; font-weight: bold;">📊 スプレッドシートで履歴を確認</a>';
    html += '</div>';
    html += '<p style="color: #666; font-size: 12px;">※スプレッドシートを開いた後、「' + REMINDER_SHEET_NAME + '」シートをご確認ください。</p>';
    
    html += '<h3>📝 確認が必要なファイル</h3>';
    html += '<p>以下のファイルは<strong>' + reminderThreshold + '日以上</strong>更新されていません。<br>';
    html += '各ファイルについて、以下のいずれかの対応をお願いします：</p>';
    
    html += '<ul>';
    html += '<li>✅ <strong>更新が必要</strong> → ファイルを更新してください</li>';
    html += '<li>📌 <strong>永久保存版</strong> → ファイル名の先頭に "keep_" を追加してください</li>';
    html += '<li>📦 <strong>アーカイブ</strong> → ファイル名の先頭に "zzzz_" を追加してください</li>';
    html += '<li>🗑️ <strong>不要</strong> → ファイルを削除してください</li>';
    html += '</ul>';
    
    html += '<h3>📂 対象ファイル一覧</h3>';
    html += '<table border="1" cellpadding="8" cellspacing="0" style="border-collapse: collapse; width: 100%; font-family: Arial, sans-serif;">';
    html += '<thead style="background-color: #f2f2f2;">';
    html += '<tr><th>ファイル名</th><th>最終更新日</th><th>経過日数</th><th>場所</th><th>サイズ</th><th>アクション</th></tr>';
    html += '</thead><tbody>';
    
    reminderFiles.sort(function(a, b) { return a.lastUpdated - b.lastUpdated; });
    
    reminderFiles.forEach(function(file, index) {
      const rowColor = index % 2 === 0 ? '#ffffff' : '#f9f9f9';
      const daysSinceUpdate = Math.floor((new Date() - file.lastUpdated) / (24 * 60 * 60 * 1000));
      const sizeInMB = (file.size / (1024 * 1024)).toFixed(2);
      
      let dayStyle = '';
      if (daysSinceUpdate > 180) dayStyle = 'color: #d32f2f; font-weight: bold;';
      else if (daysSinceUpdate > 90) dayStyle = 'color: #f57c00; font-weight: bold;';
      else dayStyle = 'color: #388e3c;';
      
      html += '<tr style="background-color: ' + rowColor + ';">';
      html += '<td style="font-weight: bold;">' + file.name + '</td>';
      html += '<td>' + file.lastUpdated.toLocaleDateString('ja-JP') + '</td>';
      html += '<td style="' + dayStyle + '">' + daysSinceUpdate + '日前</td>';
      html += '<td><small>' + file.path + '</small></td>';
      html += '<td>' + sizeInMB + ' MB</td>';
      html += '<td><a href="' + file.url + '" target="_blank" style="color: #1a73e8; text-decoration: none;">📂 開く</a></td>';
      html += '</tr>';
    });
    
    html += '</tbody></table>';
    
    html += '<h3>⚠️ ご注意</h3>';
    html += '<ul>';
    html += '<li>永久保存版として残したいファイルは、ファイル名の先頭に "keep_" を追加してください</li>';
    html += '<li>アーカイブ（古いが一応残す）ファイルは、ファイル名の先頭に "zzzz_" を追加してください</li>';
    html += '<li>このリマインドは' + reminderThreshold + '日以上未更新のファイルに対して送信されています</li>';
    html += '<li>ご質問がございましたら、システム管理者までお問い合わせください</li>';
    html += '</ul>';
    
    html += '<hr>';
    html += '<p style="color: #666; font-size: 12px;">';
    html += 'このメールは自動送信されています。<br>';
    html += '履歴は<a href="' + spreadsheetUrl + '" target="_blank">スプレッドシート</a>でご確認いただけます。<br>';
    html += 'システム管理者: ' + ADMIN_EMAIL;
    html += '</p>';
    
    MailApp.sendEmail({
      to: notifyEmail,
      subject: subject,
      htmlBody: html
    });
    
    Logger.log('リマインドメール送信成功: ' + notifyEmail);
    return true;
    
  } catch (error) {
    Logger.log('リマインドメール送信失敗: ' + notifyEmail + ' - ' + error.toString());
    return false;
  }
}

// ===== 新規追加機能 =====

/**
 * zzzz_プレフィックスのファイルをArchiveドライブへ自動移動
 */
function moveZzzzFilesToArchive() {
  const ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
  const driveConfigSheet = ss.getSheetByName(DRIVE_CONFIG_SHEET_NAME);
  const auditSheet = ss.getSheetByName(AUDIT_LOG_SHEET_NAME);
  
  if (!driveConfigSheet || driveConfigSheet.getLastRow() <= 1) {
    Logger.log('ドライブ設定が見つかりません');
    return;
  }
  
  const driveConfigs = _getDriveConfigs(driveConfigSheet);
  const dailyDrive = driveConfigs.find(d => d.type === DRIVE_TYPES.DAILY);
  const archiveDrive = driveConfigs.find(d => d.type === DRIVE_TYPES.ARCHIVE);
  
  if (!dailyDrive || !archiveDrive) {
    Logger.log('Daily/Archiveドライブが設定されていません');
    return;
  }
  
  let movedCount = 0;
  const now = new Date();
  
  try {
    const query = `'${dailyDrive.id}' in parents and name contains 'zzzz_' and trashed = false`;
    const files = DriveApp.searchFiles(query);
    
    while (files.hasNext()) {
      const file = files.next();
      const fileName = file.getName();
      
      if (fileName.toLowerCase().startsWith('zzzz_')) {
        try {
          const archiveFolder = DriveApp.getFolderById(archiveDrive.id);
          file.moveTo(archiveFolder);
          
          auditSheet.appendRow([
            now,
            DRIVE_TYPES.DAILY + ' → ' + DRIVE_TYPES.ARCHIVE,
            file.getOwner().getEmail(),
            'ファイル移動',
            'zzzz_自動アーカイブ',
            fileName,
            'システム自動'
          ]);
          
          movedCount++;
          Logger.log('アーカイブ移動: ' + fileName);
          
        } catch (e) {
          Logger.log('ファイル移動エラー: ' + fileName + ' - ' + e.toString());
        }
      }
    }
    
    if (movedCount > 0) {
      Logger.log('アーカイブ移動完了: ' + movedCount + '件');
      
      MailApp.sendEmail({
        to: ADMIN_EMAIL,
        subject: '【Drive管理】ファイルアーカイブ完了',
        body: 'zzzz_プレフィックスのファイル ' + movedCount + '件をArchiveドライブへ移動しました。\n\n' +
              '詳細はAuditLogシートをご確認ください。'
      });
    }
    
  } catch (error) {
    Logger.log('アーカイブ移動エラー: ' + error.toString());
  }
}

/**
 * 期限切れの外部共有権限を自動失効
 */
function revokeExpiredPermissions() {
  const ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
  const externalShareSheet = ss.getSheetByName(EXTERNAL_SHARE_SHEET_NAME);
  const auditSheet = ss.getSheetByName(AUDIT_LOG_SHEET_NAME);
  
  if (!externalShareSheet || externalShareSheet.getLastRow() <= 1) {
    Logger.log('外部共有ログが見つかりません');
    return;
  }
  
  const now = new Date();
  const data = externalShareSheet.getDataRange().getValues();
  let revokedCount = 0;
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const fileId = row[1];
    const email = row[3];
    const permission = row[4];
    const expiryDate = row[5];
    
    if (expiryDate && expiryDate instanceof Date && expiryDate < now) {
      try {
        const file = DriveApp.getFileById(fileId);
        file.removeEditor(email);
        file.removeViewer(email);
        
        externalShareSheet.getRange(i + 1, 8).setValue('失効済み');
        
        auditSheet.appendRow([
          now,
          'External-Share',
          email,
          permission,
          '権限失効',
          '有効期限切れ',
          'システム自動'
        ]);
        
        revokedCount++;
        Logger.log('権限失効: ' + email + ' - ' + file.getName());
        
      } catch (e) {
        Logger.log('権限失効エラー: ' + fileId + ' - ' + e.toString());
      }
    }
  }
  
  if (revokedCount > 0) {
    Logger.log('権限失効完了: ' + revokedCount + '件');
  }
}

/**
 * ドライブメンバーリストを更新
 */
function updateDriveMembersList() {
  const ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
  const driveConfigSheet = ss.getSheetByName(DRIVE_CONFIG_SHEET_NAME);
  const membersSheet = ss.getSheetByName(DRIVE_MEMBERS_SHEET_NAME);
  
  if (!driveConfigSheet || !membersSheet) {
    Logger.log('必要なシートが見つかりません');
    return;
  }
  
  if (membersSheet.getLastRow() > 1) {
    membersSheet.getRange(2, 1, membersSheet.getMaxRows() - 1, 8).clearContent();
  }
  
  const driveConfigs = _getDriveConfigs(driveConfigSheet);
  const now = new Date();
  const rows = [];
  
  driveConfigs.forEach(function(driveConfig) {
    try {
      const driveId = driveConfig.id;
      const driveName = driveConfig.name;
      
      const permissions = Drive.Permissions.list(driveId, {
        supportsAllDrives: true,
        fields: 'items(emailAddress,role,type,expirationTime)',
        maxResults: 100
      });
      
      permissions.items.forEach(function(permission) {
        if (permission.emailAddress) {
          const isExternal = !permission.emailAddress.endsWith('@' + Session.getActiveUser().getEmail().split('@')[1]);
          
          rows.push([
            now,
            driveId,
            driveName,
            permission.emailAddress,
            permission.role,
            '',
            permission.expirationTime || '',
            isExternal ? '外部' : '内部'
          ]);
        }
      });
      
    } catch (e) {
      Logger.log('ドライブメンバー取得エラー: ' + driveConfig.name + ' - ' + e.toString());
    }
  });
  
  if (rows.length > 0) {
    membersSheet.getRange(2, 1, rows.length, 8).setValues(rows);
    Logger.log('ドライブメンバーリスト更新: ' + rows.length + '件');
    _highlightProblematicPermissions(membersSheet);
  }
}

/**
 * メンバーダッシュボードを更新（権限の可視化）
 */
function updateMemberDashboard() {
  const ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
  const membersSheet = ss.getSheetByName(DRIVE_MEMBERS_SHEET_NAME);
  const dashboardSheet = ss.getSheetByName(MEMBER_DASHBOARD_SHEET_NAME);
  
  if (!membersSheet || !dashboardSheet || membersSheet.getLastRow() <= 1) {
    Logger.log('データが不足しています');
    return;
  }
  
  dashboardSheet.clear();
  dashboardSheet.appendRow(['メンバーダッシュボード', new Date()]);
  dashboardSheet.appendRow(['']);
  
  const driveConfigs = _getDriveConfigs(ss.getSheetByName(DRIVE_CONFIG_SHEET_NAME));
  const memberData = membersSheet.getDataRange().getValues();
  
  driveConfigs.forEach(function(driveConfig, index) {
    const driveName = driveConfig.name;
    const driveType = driveConfig.type;
    
    const driveMembers = memberData.filter(row => row[2] === driveName);
    
    const roleCounts = {};
    let externalCount = 0;
    
    driveMembers.forEach(function(member) {
      const role = member[4];
      const isExternal = member[7] === '外部';
      
      roleCounts[role] = (roleCounts[role] || 0) + 1;
      if (isExternal) externalCount++;
    });
    
    const startRow = dashboardSheet.getLastRow() + 2;
    dashboardSheet.appendRow([driveType + ' (' + driveName + ')']);
    
    let bgColor = '#ffffff';
    if (driveType === DRIVE_TYPES.DAILY) bgColor = '#e3f2fd';
    else if (driveType === DRIVE_TYPES.EXTERNAL) bgColor = '#ffebee';
    else if (driveType === DRIVE_TYPES.ARCHIVE) bgColor = '#fff3e0';
    
    dashboardSheet.getRange(startRow, 1, 1, 2).setBackground(bgColor).setFontWeight('bold');
    
    for (const role in roleCounts) {
      dashboardSheet.appendRow(['  ' + role, roleCounts[role] + '名']);
    }
    
    if (externalCount > 0) {
      dashboardSheet.appendRow(['  外部ユーザー', externalCount + '名']).getRange(dashboardSheet.getLastRow(), 1, 1, 2).setFontColor('#d32f2f');
    }
    
    if (driveType === DRIVE_TYPES.ARCHIVE && roleCounts['writer']) {
      dashboardSheet.appendRow(['  ⚠️ 警告: Archiveに編集権限あり', roleCounts['writer'] + '名'])
        .getRange(dashboardSheet.getLastRow(), 1, 1, 2).setBackground('#ffcdd2');
    }
  });
  
  Logger.log('メンバーダッシュボード更新完了');
}

/**
 * 権限の定期監査レポートを作成
 */
function auditAccounts() {
  const ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
  const membersSheet = ss.getSheetByName(DRIVE_MEMBERS_SHEET_NAME);
  const auditSheet = ss.getSheetByName(AUDIT_LOG_SHEET_NAME);
  
  if (!membersSheet || membersSheet.getLastRow() <= 1) {
    Logger.log('メンバーデータがありません');
    return;
  }
  
  const now = new Date();
  const memberData = membersSheet.getDataRange().getValues();
  const issues = [];
  
  for (let i = 1; i < memberData.length; i++) {
    const row = memberData[i];
    const driveName = row[2];
    const email = row[3];
    const role = row[4];
    const isExternal = row[7] === '外部';
    
    let issue = null;
    
    if (driveName.includes('Archive') && (role === 'writer' || role === 'owner')) {
      issue = 'Archiveドライブに編集権限';
    }
    
    if (isExternal && role === 'writer') {
      issue = '外部ユーザーに編集権限';
    }
    
    if (issue) {
      issues.push({
        driveName: driveName,
        email: email,
        role: role,
        issue: issue
      });
      
      auditSheet.appendRow([
        now,
        driveName,
        email,
        role,
        '監査指摘',
        issue,
        'システム監査'
      ]);
    }
  }
  
  if (issues.length > 0) {
    const emailBody = _createAuditReportEmail(issues);
    MailApp.sendEmail({
      to: ADMIN_EMAIL,
      subject: '【Drive管理】権限監査レポート - 要対応',
      htmlBody: emailBody
    });
    Logger.log('監査レポート送信: ' + issues.length + '件の問題を検出');
  } else {
    Logger.log('監査完了: 問題なし');
  }
}

/**
 * 外部共有を一時的に全て無効化（緊急時用）
 */
function emergencyDisableExternalSharing() {
  const ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
  const driveConfigSheet = ss.getSheetByName(DRIVE_CONFIG_SHEET_NAME);
  const auditSheet = ss.getSheetByName(AUDIT_LOG_SHEET_NAME);
  
  const driveConfigs = _getDriveConfigs(driveConfigSheet);
  const externalDrive = driveConfigs.find(d => d.type === DRIVE_TYPES.EXTERNAL);
  
  if (!externalDrive) {
    Logger.log('External-Shareドライブが設定されていません');
    return;
  }
  
  const now = new Date();
  let revokedCount = 0;
  
  try {
    const permissions = Drive.Permissions.list(externalDrive.id, {
      supportsAllDrives: true,
      fields: 'items(id,emailAddress,role,type)',
      maxResults: 100
    });
    
    permissions.items.forEach(function(permission) {
      if (permission.emailAddress && !permission.emailAddress.endsWith('@' + Session.getActiveUser().getEmail().split('@')[1])) {
        try {
          Drive.Permissions.remove(externalDrive.id, permission.id, {
            supportsAllDrives: true
          });
          
          auditSheet.appendRow([
            now,
            externalDrive.name,
            permission.emailAddress,
            permission.role,
            '緊急権限削除',
            '外部共有の一時停止',
            Session.getActiveUser().getEmail()
          ]);
          
          revokedCount++;
        } catch (e) {
          Logger.log('権限削除エラー: ' + permission.emailAddress + ' - ' + e.toString());
        }
      }
    });
    
    MailApp.sendEmail({
      to: ADMIN_EMAIL,
      subject: '【緊急】外部共有を一時停止しました',
      body: '外部共有の緊急停止を実行しました。\n\n' +
            '削除した権限数: ' + revokedCount + '件\n' +
            '実行者: ' + Session.getActiveUser().getEmail() + '\n' +
            '実行時刻: ' + now.toLocaleString('ja-JP')
    });
    
    Logger.log('緊急停止完了: ' + revokedCount + '件の外部共有を削除');
    
  } catch (error) {
    Logger.log('緊急停止エラー: ' + error.toString());
  }
}

// ===== ヘルパー関数 =====

/**
 * ドライブ設定を取得
 */
function _getDriveConfigs(driveConfigSheet) {
  if (!driveConfigSheet || driveConfigSheet.getLastRow() <= 1) return [];
  
  const data = driveConfigSheet.getDataRange().getValues();
  const configs = [];
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[1] && row[1] !== 'ドライブIDを入力') {
      configs.push({
        type: row[0],
        id: row[1],
        name: row[2],
        geminiSearch: row[3] === 'はい',
        remarks: row[4]
      });
    }
  }
  
  return configs;
}

/**
 * 問題のある権限をハイライト
 */
function _highlightProblematicPermissions(membersSheet) {
  const data = membersSheet.getDataRange().getValues();
  
  membersSheet.getRange(2, 1, membersSheet.getLastRow() - 1, 8).setBackground(null);
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const driveName = row[2];
    const role = row[4];
    const isExternal = row[7] === '外部';
    
    let shouldHighlight = false;
    
    if (driveName.includes('Archive') && (role === 'writer' || role === 'owner')) {
      shouldHighlight = true;
    }
    
    if (isExternal && role === 'writer') {
      shouldHighlight = true;
    }
    
    if (shouldHighlight) {
      membersSheet.getRange(i + 1, 1, 1, 8).setBackground('#ffcdd2');
    }
  }
}

/**
 * 監査レポートメールの作成
 */
function _createAuditReportEmail(issues) {
  let html = '<h2>🔍 Google Drive 権限監査レポート</h2>';
  html += '<p><strong>監査日時:</strong> ' + new Date().toLocaleString('ja-JP') + '</p>';
  html += '<p><strong>検出された問題:</strong> ' + issues.length + '件</p>';
  
  html += '<h3>⚠️ 要対応項目</h3>';
  html += '<table border="1" cellpadding="8" cellspacing="0" style="border-collapse: collapse;">';
  html += '<thead style="background-color: #f2f2f2;">';
  html += '<tr><th>ドライブ</th><th>メールアドレス</th><th>権限</th><th>問題</th></tr>';
  html += '</thead><tbody>';
  
  issues.forEach(function(issue) {
    html += '<tr>';
    html += '<td>' + issue.driveName + '</td>';
    html += '<td>' + issue.email + '</td>';
    html += '<td>' + issue.role + '</td>';
    html += '<td style="color: #d32f2f; font-weight: bold;">' + issue.issue + '</td>';
    html += '</tr>';
  });
  
  html += '</tbody></table>';
  
  html += '<h3>📋 推奨アクション</h3>';
  html += '<ul>';
  html += '<li>Archiveドライブの編集権限は削除し、閲覧権限のみに変更してください</li>';
  html += '<li>外部ユーザーへの編集権限は、必要性を確認の上、期限を設定してください</li>';
  html += '<li>不要な権限は速やかに削除してください</li>';
  html += '</ul>';
  
  html += '<p style="margin-top: 20px;">';
  html += '<a href="https://docs.google.com/spreadsheets/d/' + CONFIG_SHEET_ID + '/edit#gid=' + _getSheetGid(AUDIT_LOG_SHEET_NAME) + '" ';
  html += 'style="background-color: #1a73e8; color: white; padding: 10px 20px; text-decoration: none; border-radius: 4px;">監査ログを確認</a>';
  html += '</p>';
  
  return html;
}

/**
 * 外部共有を記録する関数（手動実行用）
 */
function recordExternalShare(fileId, email, permission, expiryDays, approver) {
  const ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
  const externalShareSheet = ss.getSheetByName(EXTERNAL_SHARE_SHEET_NAME);
  
  const now = new Date();
  const expiryDate = new Date(now.getTime() + expiryDays * 24 * 60 * 60 * 1000);
  
  try {
    const file = DriveApp.getFileById(fileId);
    
    externalShareSheet.appendRow([
      now,
      fileId,
      file.getName(),
      email,
      permission,
      expiryDate,
      approver || Session.getActiveUser().getEmail(),
      ''
    ]);
    
    Logger.log('外部共有を記録: ' + file.getName() + ' → ' + email);
    
  } catch (e) {
    Logger.log('外部共有記録エラー: ' + e.toString());
  }
}

/**
 * ドライブ容量レポートを生成（四半期用）
 */
function generateStorageReport() {
  const ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
  const driveConfigSheet = ss.getSheetByName(DRIVE_CONFIG_SHEET_NAME);
  const driveConfigs = _getDriveConfigs(driveConfigSheet);
  
  let reportHtml = '<h2>📊 Google Drive 容量レポート</h2>';
  reportHtml += '<p><strong>レポート日時:</strong> ' + new Date().toLocaleString('ja-JP') + '</p>';
  
  reportHtml += '<table border="1" cellpadding="8" cellspacing="0" style="border-collapse: collapse;">';
  reportHtml += '<thead style="background-color: #f2f2f2;">';
  reportHtml += '<tr><th>ドライブ種別</th><th>ドライブ名</th><th>ファイル数</th><th>推定容量</th></tr>';
  reportHtml += '</thead><tbody>';
  
  driveConfigs.forEach(function(driveConfig) {
    try {
      const folder = DriveApp.getFolderById(driveConfig.id);
      const fileCount = _countFilesInFolder(folder);
      
      reportHtml += '<tr>';
      reportHtml += '<td>' + driveConfig.type + '</td>';
      reportHtml += '<td>' + driveConfig.name + '</td>';
      reportHtml += '<td>' + fileCount + '</td>';
      reportHtml += '<td>計測中...</td>';
      reportHtml += '</tr>';
      
    } catch (e) {
      Logger.log('容量レポートエラー: ' + driveConfig.name + ' - ' + e.toString());
    }
  });
  
  reportHtml += '</tbody></table>';
  
  reportHtml += '<h3>💡 推奨アクション</h3>';
  reportHtml += '<ul>';
  reportHtml += '<li>Archiveドライブの容量が増大している場合は、古いファイルの削除を検討してください</li>';
  reportHtml += '<li>重複ファイルの削除により、容量を節約できます</li>';
  reportHtml += '<li>大容量ファイルは圧縮またはクラウドストレージへの移行を検討してください</li>';
  reportHtml += '</ul>';
  
  MailApp.sendEmail({
    to: ADMIN_EMAIL,
    subject: '【Drive管理】四半期容量レポート',
    htmlBody: reportHtml
  });
  
  Logger.log('容量レポート送信完了');
}

/**
 * フォルダ内のファイル数をカウント（再帰的）
 */
function _countFilesInFolder(folder) {
  let count = 0;
  
  const files = folder.getFiles();
  while (files.hasNext()) {
    files.next();
    count++;
  }
  
  const subFolders = folder.getFolders();
  while (subFolders.hasNext()) {
    count += _countFilesInFolder(subFolders.next());
  }
  
  return count;
}

/**
 * シートのGIDを取得する関数
 */
function _getSheetGid(sheetName) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
    const sheet = ss.getSheetByName(sheetName);
    if (sheet) {
      return sheet.getSheetId();
    }
  } catch (e) {
    Logger.log('シートGID取得エラー: ' + sheetName + ' - ' + e);
  }
  return 0;
}

/**
 * 設定の妥当性をチェックする関数
 */
function validateConfiguration() {
  const ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
  const configSheet = ss.getSheetByName(CONFIG_SHEET_NAME);
  const driveConfigSheet = ss.getSheetByName(DRIVE_CONFIG_SHEET_NAME);
  const issues = [];
  
  // フォルダ設定の検証
  const folderConfigs = _getFolderConfigs(configSheet);
  
  if (folderConfigs.length === 0) {
    issues.push('❌ 監視対象フォルダが設定されていません');
  } else {
    issues.push('✅ 監視対象フォルダ: ' + folderConfigs.length + '個設定済み');
    
    folderConfigs.forEach(function(config, index) {
      let folderStatus = '❌';
      let emailStatus = '❌';
      
      try {
        const folder = DriveApp.getFolderById(config.folderId);
        folderStatus = '✅ ' + folder.getName();
      } catch (e) {
        folderStatus = '❌ アクセス不可';
      }
      
      if (config.notifyEmail && config.notifyEmail.includes('@')) {
        emailStatus = '✅';
      } else {
        emailStatus = '❌ 無効';
      }
      
      issues.push('  ' + (index + 1) + '. フォルダ' + folderStatus + ' | メール' + emailStatus + ' (' + config.notifyEmail + ')');
      issues.push('     リマインド: ' + config.reminderThreshold + '日/' + config.reminderFrequency + '日間隔');
    });
  }
  
  // ドライブ設定の検証
  const driveConfigs = _getDriveConfigs(driveConfigSheet);
  
  if (driveConfigs.length === 0) {
    issues.push('❌ ドライブが設定されていません');
  } else {
    const types = [DRIVE_TYPES.DAILY, DRIVE_TYPES.EXTERNAL, DRIVE_TYPES.ARCHIVE];
    types.forEach(function(type) {
      const drive = driveConfigs.find(d => d.type === type);
      if (!drive) {
        issues.push('❌ ' + type + 'ドライブが設定されていません');
      } else {
        try {
          DriveApp.getFolderById(drive.id);
          issues.push('✅ ' + type + 'ドライブ: ' + drive.name);
        } catch (e) {
          issues.push('❌ ' + type + 'ドライブ: アクセス不可');
        }
      }
    });
  }
  
  if (ADMIN_EMAIL && ADMIN_EMAIL.includes('@')) {
    issues.push('✅ システム管理者メール: ' + ADMIN_EMAIL);
  } else {
    issues.push('❌ システム管理者メールアドレスが正しく設定されていません');
  }
  
  Logger.log('=== 設定チェック結果 ===');
  issues.forEach(function(issue) {
    Logger.log(issue);
  });
  
  return issues;
}

/**
 * スプレッドシートの権限を確認・付与する関数
 */
function checkAndGrantSpreadsheetAccess() {
  try {
    const ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
    const file = DriveApp.getFileById(CONFIG_SHEET_ID);
    
    const access = file.getSharingAccess();
    const permission = file.getSharingPermission();
    
    Logger.log('=== スプレッドシート権限情報 ===');
    Logger.log('ファイル名: ' + ss.getName());
    Logger.log('アクセスレベル: ' + access);
    Logger.log('権限レベル: ' + permission);
    Logger.log('オーナー: ' + file.getOwner().getEmail());
    
    const editors = file.getEditors();
    if (editors.length > 0) {
      Logger.log('編集者:');
      editors.forEach(function(editor) {
        Logger.log('  - ' + editor.getEmail());
      });
    }
    
    const viewers = file.getViewers();
    if (viewers.length > 0) {
      Logger.log('閲覧者:');
      viewers.forEach(function(viewer) {
        Logger.log('  - ' + viewer.getEmail());
      });
    }
    
    Logger.log('スプレッドシートURL: https://docs.google.com/spreadsheets/d/' + CONFIG_SHEET_ID);
    
    return {
      name: ss.getName(),
      url: 'https://docs.google.com/spreadsheets/d/' + CONFIG_SHEET_ID,
      access: access,
      permission: permission,
      owner: file.getOwner().getEmail()
    };
    
  } catch (error) {
    Logger.log('スプレッドシート権限確認エラー: ' + error.toString());
    return null;
  }
}

/**
 * 特定のメールアドレスにスプレッドシートの編集権限を付与する関数
 */
function grantEditAccessToEmail(email) {
  try {
    const file = DriveApp.getFileById(CONFIG_SHEET_ID);
    file.addEditor(email);
    Logger.log('編集権限を付与しました: ' + email);
    
    MailApp.sendEmail({
      to: email,
      subject: '【Drive管理】スプレッドシートへのアクセス権限付与',
      body: 'Drive管理スプレッドシートへの編集権限が付与されました。\n\n' +
            'スプレッドシートURL: https://docs.google.com/spreadsheets/d/' + CONFIG_SHEET_ID + '\n\n' +
            'このスプレッドシートで以下の情報を確認できます：\n' +
            '- ルール違反一覧\n' +
            '- ファイル更新リマインド履歴\n' +
            '- 重複ファイル一覧\n' +
            '- フォルダ設定\n' +
            '- ドライブメンバー管理\n' +
            '- 権限監査ログ'
    });
    
    return true;
  } catch (error) {
    Logger.log('権限付与エラー: ' + error.toString());
    return false;
  }
}

/**
 * レポートのサマリーを生成する関数
 */
function generateReportSummary() {
  const ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
  const now = new Date();
  
  const violSheet = ss.getSheetByName(VIOL_SHEET_NAME);
  const reminderSheet = ss.getSheetByName(REMINDER_SHEET_NAME);
  const dupSheet = ss.getSheetByName(DUP_SHEET_NAME);
  const membersSheet = ss.getSheetByName(DRIVE_MEMBERS_SHEET_NAME);
  const auditSheet = ss.getSheetByName(AUDIT_LOG_SHEET_NAME);
  
  const summary = {
    date: now.toLocaleString('ja-JP'),
    violations: violSheet.getLastRow() - 1,
    reminders: reminderSheet.getLastRow() - 1,
    duplicates: _countDuplicateFiles(dupSheet),
    totalMembers: membersSheet.getLastRow() - 1,
    auditIssues: _countRecentAuditIssues(auditSheet),
    spreadsheetUrl: 'https://docs.google.com/spreadsheets/d/' + CONFIG_SHEET_ID
  };
  
  Logger.log('=== レポートサマリー ===');
  Logger.log('作成日時: ' + summary.date);
  Logger.log('ルール違反: ' + summary.violations + '件');
  Logger.log('リマインド: ' + summary.reminders + '件');
  Logger.log('重複ファイル: ' + summary.duplicates + '件');
  Logger.log('総メンバー数: ' + summary.totalMembers + '名');
  Logger.log('監査指摘事項: ' + summary.auditIssues + '件');
  Logger.log('詳細URL: ' + summary.spreadsheetUrl);
  
  return summary;
}

/**
 * 重複ファイル数をカウントする補助関数
 */
function _countDuplicateFiles(dupSheet) {
  if (dupSheet.getLastRow() <= 1) return 0;
  
  const data = dupSheet.getDataRange().getValues();
  const md5Map = {};
  let duplicateCount = 0;
  
  for (let i = 1; i < data.length; i++) {
    const md5 = data[i][5];
    if (md5 && md5 !== 'N/A' && md5 !== '取得エラー') {
      if (md5Map[md5]) {
        duplicateCount++;
      } else {
        md5Map[md5] = true;
      }
    }
  }
  
  return duplicateCount;
}

/**
 * 最近の監査指摘事項をカウント
 */
function _countRecentAuditIssues(auditSheet) {
  if (auditSheet.getLastRow() <= 1) return 0;
  
  const data = auditSheet.getDataRange().getValues();
  const thirtyDaysAgo = new Date();
  thirtyDaysAgo.setDate(thirtyDaysAgo.getDate() - 30);
  
  let count = 0;
  for (let i = 1; i < data.length; i++) {
    const auditDate = data[i][0];
    const action = data[i][4];
    if (auditDate instanceof Date && auditDate > thirtyDaysAgo && action === '監査指摘') {
      count++;
    }
  }
  
  return count;
}

// ===== テスト用関数 =====

/**
 * リマインド確認のみを実行する関数（テスト用）
 */
function testFileReminders() {
  Logger.log('=== ファイルリマインドテスト開始 ===');
  sendFileReviewReminders();
  Logger.log('=== ファイルリマインドテスト完了 ===');
}

/**
 * ルール違反チェックのみを実行する関数（テスト用）
 */
function testRuleCheck() {
  Logger.log('=== ルール違反チェックテスト開始 ===');
  scanNamingViolations();
  Logger.log('=== ルール違反チェックテスト完了 ===');
}

/**
 * 重複ファイルスキャンのみを実行する関数（テスト用）
 */
function testDuplicateCheck() {
  Logger.log('=== 重複ファイルスキャンテスト開始 ===');
  scanForDuplicates();
  Logger.log('=== 重複ファイルスキャンテスト完了 ===');
}

/**
 * アーカイブ移動のみを実行する関数（テスト用）
 */
function testArchiveMove() {
  Logger.log('=== アーカイブ移動テスト開始 ===');
  moveZzzzFilesToArchive();
  Logger.log('=== アーカイブ移動テスト完了 ===');
}

/**
 * メンバー管理機能のみを実行する関数（テスト用）
 */
function testMemberManagement() {
  Logger.log('=== メンバー管理テスト開始 ===');
  updateDriveMembersList();
  updateMemberDashboard();
  Logger.log('=== メンバー管理テスト完了 ===');
}

/**
 * チェックステータスのサマリーを表示する関数
 */
function showCheckStatus() {
  const ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
  const configSheet = ss.getSheetByName(CONFIG_SHEET_NAME);
  const configs = _getFolderConfigs(configSheet);
  const now = new Date();
  
  Logger.log('=== フォルダチェックステータス ===');
  Logger.log('現在日時: ' + now.toLocaleString('ja-JP'));
  Logger.log('');
  
  let needCheckCount = 0;
  let skipCount = 0;
  
  configs.forEach(function(config, index) {
    const shouldCheck = _shouldCheckFolder(config.lastCheckDate, config.reminderFrequency, now);
    const lastCheck = config.lastCheckDate ? new Date(config.lastCheckDate) : null;
    const daysUntilNext = _getDaysUntilNextCheck(config.lastCheckDate, config.reminderFrequency, now);
    
    let status = '';
    if (!lastCheck) {
      status = '⚠️ 未チェック';
      needCheckCount++;
    } else if (shouldCheck) {
      const daysSinceLastCheck = Math.floor((now.getTime() - lastCheck.getTime()) / (24 * 60 * 60 * 1000));
      status = '✅ チェック必要（' + daysSinceLastCheck + '日経過）';
      needCheckCount++;
    } else {
      status = '⏭️ スキップ（次回まで' + daysUntilNext + '日）';
      skipCount++;
    }
    
    Logger.log((index + 1) + '. フォルダID: ' + config.folderId);
    Logger.log('   リマインド頻度: ' + config.reminderFrequency + '日ごと');
    Logger.log('   最終チェック: ' + (lastCheck ? lastCheck.toLocaleString('ja-JP') : 'なし'));
    Logger.log('   ステータス: ' + status);
    Logger.log('');
  });
  
  Logger.log('=== サマリー ===');
  Logger.log('チェック必要: ' + needCheckCount + '個');
  Logger.log('スキップ可能: ' + skipCount + '個');
  Logger.log('合計: ' + configs.length + '個');
}

/**
 * 強制的にすべてのフォルダをチェックする関数
 */
function forceCheckAllFolders() {
  const ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);
  const configSheet = ss.getSheetByName(CONFIG_SHEET_NAME);
  const configs = _getFolderConfigs(configSheet);
  
  // 一時的にすべての最終チェック日時をクリア
  configs.forEach(function(config) {
    configSheet.getRange(config.rowIndex, 6).setValue('');
  });
  
  Logger.log('すべてのフォルダを強制チェック対象にしました。');
  
  // メインチェックを実行
  mainCheckExtended();
}