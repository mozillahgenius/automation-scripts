/**
 * Google Drive 共有ドライブ階層構造可視化ツール
 *
 * 【初回セットアップ】
 * 1. このコードをApps Scriptに貼り付けて保存
 * 2. 関数「setup」を実行（Drive APIの有効化と権限承認を自動で行います）
 * 3. スプレッドシートに戻り、メニュー「ドライブ構造」から実行
 */

// ===========================================
// 初期セットアップ
// ===========================================

/**
 * 初期セットアップを実行
 * この関数を最初に実行してください
 */
function setup() {
  // Drive APIの有効化確認
  try {
    Drive.Drives.list({ pageSize: 1 });
  } catch (e) {
    console.log('セットアップエラー: Drive APIが有効化されていません。');
    console.log('以下の手順で有効化してください:');
    console.log('1. Apps Scriptエディタの左側メニューで「サービス」の「+」をクリック');
    console.log('2. 「Drive API」を選択して「追加」をクリック');
    console.log('3. 再度この「setup」関数を実行してください');
    throw new Error('Drive APIが有効化されていません。左側メニューの「サービス」から「Drive API」を追加してください。');
  }

  // 権限確認のためスプレッドシートにアクセス
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    console.log('セットアップ完了（権限承認済み）');
    console.log('');
    console.log('次のステップ:');
    console.log('1. スプレッドシートを開く（または再読み込み）');
    console.log('2. メニューに「ドライブ構造」が追加されています');
    console.log('3. 「ドライブ構造」→「共有ドライブを選択」または「マイドライブを処理」を実行');
    return;
  }

  // メニューを追加
  onOpen();

  // 完了メッセージ
  const ui = SpreadsheetApp.getUi();
  ui.alert(
    'セットアップ完了',
    'セットアップが完了しました。\n\n' +
    'スプレッドシートに「ドライブ構造」メニューが追加されました。\n' +
    'メニューから「共有ドライブを選択」または「マイドライブを処理」を選んで実行してください。',
    ui.ButtonSet.OK
  );
}

// ===========================================
// メニュー・UI関連
// ===========================================

/**
 * スプレッドシートを開いた時にカスタムメニューを追加
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('ドライブ構造')
    .addItem('共有ドライブを選択', 'showDriveSelector')
    .addItem('マイドライブを処理', 'processMyDrive')
    .addItem('フォルダIDを指定', 'showFolderIdInput')
    .addSeparator()
    .addSubMenu(ui.createMenu('定期更新')
      .addItem('設定シートを作成/開く', 'openConfigSheet')
      .addItem('定期更新を有効化', 'enableScheduledUpdate')
      .addItem('定期更新を無効化', 'disableScheduledUpdate')
      .addItem('今すぐ定期更新を実行', 'runScheduledUpdate'))
    .addSeparator()
    .addItem('使い方', 'showHelp')
    .addToUi();
}

/**
 * 共有ドライブ選択ダイアログを表示
 */
function showDriveSelector() {
  const drives = getSharedDrives();

  if (drives.length === 0) {
    SpreadsheetApp.getUi().alert(
      'アクセス可能な共有ドライブが見つかりませんでした。\n\n' +
      '・共有ドライブへのアクセス権限を確認してください\n' +
      '・Google Workspaceアカウントが必要です\n\n' +
      '「マイドライブを処理」を使用することもできます。'
    );
    return;
  }

  const html = createDriveSelectorHtml(drives);
  const htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(400)
    .setHeight(300);

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, '共有ドライブを選択');
}

/**
 * ドライブ選択用HTMLを生成
 */
function createDriveSelectorHtml(drives) {
  let options = '';
  drives.forEach(function(drive) {
    options += '<option value="' + drive.id + '">' + escapeHtml(drive.name) + '</option>';
  });

  return `
    <!DOCTYPE html>
    <html>
    <head>
      <base target="_top">
      <style>
        body { font-family: Arial, sans-serif; padding: 20px; }
        select { width: 100%; padding: 10px; font-size: 14px; margin: 10px 0; }
        button {
          width: 100%; padding: 12px; font-size: 14px;
          background-color: #4285f4; color: white;
          border: none; border-radius: 4px; cursor: pointer;
          margin-top: 10px;
        }
        button:hover { background-color: #3367d6; }
        button:disabled { background-color: #ccc; cursor: not-allowed; }
        .info { color: #666; font-size: 12px; margin-top: 15px; }
        .loading { display: none; text-align: center; margin-top: 10px; }
      </style>
    </head>
    <body>
      <p>処理する共有ドライブを選択してください:</p>
      <select id="driveSelect">${options}</select>
      <button id="submitBtn" onclick="processSelected()">構造を取得</button>
      <div id="loading" class="loading">処理中...</div>
      <p class="info">※ファイル数が多い場合、処理に時間がかかることがあります</p>
      <script>
        function processSelected() {
          const btn = document.getElementById('submitBtn');
          const loading = document.getElementById('loading');
          btn.disabled = true;
          btn.textContent = '処理中...';
          loading.style.display = 'block';

          const driveId = document.getElementById('driveSelect').value;
          const driveName = document.getElementById('driveSelect').selectedOptions[0].text;
          google.script.run
            .withSuccessHandler(function() { google.script.host.close(); })
            .withFailureHandler(function(e) {
              alert('エラー: ' + e.message);
              btn.disabled = false;
              btn.textContent = '構造を取得';
              loading.style.display = 'none';
            })
            .processSharedDrive(driveId, driveName);
        }
      </script>
    </body>
    </html>
  `;
}

/**
 * フォルダID入力ダイアログを表示
 */
function showFolderIdInput() {
  const html = createFolderIdInputHtml();
  const htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(500)
    .setHeight(320);

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'フォルダIDを指定して取得');
}

/**
 * フォルダID入力用HTMLを生成
 */
function createFolderIdInputHtml() {
  return `
    <!DOCTYPE html>
    <html>
    <head>
      <base target="_top">
      <style>
        body { font-family: Arial, sans-serif; padding: 20px; }
        input[type="text"] {
          width: 100%; padding: 10px; font-size: 14px;
          margin: 5px 0 10px 0; box-sizing: border-box;
          border: 1px solid #ccc; border-radius: 4px;
        }
        label { font-weight: bold; display: block; margin-top: 10px; }
        button {
          width: 100%; padding: 12px; font-size: 14px;
          background-color: #4285f4; color: white;
          border: none; border-radius: 4px; cursor: pointer;
          margin-top: 15px;
        }
        button:hover { background-color: #3367d6; }
        button:disabled { background-color: #ccc; cursor: not-allowed; }
        .info { color: #666; font-size: 12px; margin-top: 10px; line-height: 1.5; }
        .loading { display: none; text-align: center; margin-top: 10px; }
      </style>
    </head>
    <body>
      <label>フォルダID:</label>
      <input type="text" id="folderId" placeholder="例: 1ABC2DEF3GHI4JKL5MNO6PQR">

      <label>シート名（任意）:</label>
      <input type="text" id="sheetName" placeholder="例: プロジェクトA_構造">

      <button id="submitBtn" onclick="processFolder()">構造を取得</button>
      <div id="loading" class="loading">処理中...</div>

      <p class="info">
        【フォルダIDの取得方法】<br>
        Google DriveでフォルダをブラウザIDを開き、URLの末尾部分がIDです:<br>
        https://drive.google.com/drive/folders/<strong>ここがフォルダID</strong>
      </p>

      <script>
        function processFolder() {
          const folderId = document.getElementById('folderId').value.trim();
          let sheetName = document.getElementById('sheetName').value.trim();

          if (!folderId) {
            alert('フォルダIDを入力してください');
            return;
          }

          const btn = document.getElementById('submitBtn');
          const loading = document.getElementById('loading');
          btn.disabled = true;
          btn.textContent = '処理中...';
          loading.style.display = 'block';

          google.script.run
            .withSuccessHandler(function() { google.script.host.close(); })
            .withFailureHandler(function(e) {
              alert('エラー: ' + e.message);
              btn.disabled = false;
              btn.textContent = '構造を取得';
              loading.style.display = 'none';
            })
            .processFolderById(folderId, sheetName);
        }
      </script>
    </body>
    </html>
  `;
}

/**
 * 使い方を表示
 */
function showHelp() {
  const message =
    '【Google Drive 共有ドライブ階層構造可視化ツール】\n\n' +
    '■ 初期セットアップ\n' +
    '  関数「setup」を実行してください\n\n' +
    '■ 共有ドライブを選択\n' +
    '  アクセス可能な共有ドライブを選択して構造を取得します\n\n' +
    '■ マイドライブを処理\n' +
    '  マイドライブ内のファイル構造を取得します\n\n' +
    '■ フォルダIDを指定\n' +
    '  任意のフォルダIDを指定して、そのフォルダ以下の構造を取得します\n' +
    '  共有ドライブ内の第2層以下のフォルダも指定可能です\n\n' +
    '【出力内容】\n' +
    '・階層（深さ）\n' +
    '・名前（インデント付き）\n' +
    '・種別（フォルダ/ファイル）\n' +
    '・カテゴリ（Google Docs, PDF, 画像など）\n' +
    '・拡張子\n' +
    '・サイズ\n' +
    '・作成日・更新日\n' +
    '・フルパス';

  SpreadsheetApp.getUi().alert(message);
}

// ===========================================
// メイン処理
// ===========================================

/**
 * 共有ドライブを処理
 */
function processSharedDrive(driveId, driveName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = driveName + '_構造';

  // 既存シートがあれば削除して新規作成
  let sheet = ss.getSheetByName(sheetName);
  if (sheet) {
    ss.deleteSheet(sheet);
  }
  sheet = ss.insertSheet(sheetName);

  // ステータス表示
  SpreadsheetApp.getActiveSpreadsheet().toast('ファイル構造を取得中...', 'Processing', -1);

  // ファイル取得
  const files = getFilesFromSharedDrive(driveId, driveName);

  // シートに書き込み
  writeToSheet(sheet, files);

  // 完了通知
  SpreadsheetApp.getActiveSpreadsheet().toast(
    files.length + '件のファイル/フォルダを出力しました',
    '完了',
    5
  );
}

/**
 * マイドライブを処理
 */
function processMyDrive() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = 'マイドライブ_構造';

  // 既存シートがあれば削除して新規作成
  let sheet = ss.getSheetByName(sheetName);
  if (sheet) {
    ss.deleteSheet(sheet);
  }
  sheet = ss.insertSheet(sheetName);

  // ステータス表示
  SpreadsheetApp.getActiveSpreadsheet().toast('ファイル構造を取得中...', 'Processing', -1);

  // ファイル取得
  const files = getFilesFromMyDrive();

  // シートに書き込み
  writeToSheet(sheet, files);

  // 完了通知
  SpreadsheetApp.getActiveSpreadsheet().toast(
    files.length + '件のファイル/フォルダを出力しました',
    '完了',
    5
  );
}

/**
 * 任意のフォルダIDを指定して処理
 */
function processFolderById(folderId, customSheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // フォルダ情報を取得
  let folderInfo;
  try {
    folderInfo = Drive.Files.get(folderId, {
      supportsAllDrives: true,
      fields: 'id, name, mimeType, driveId'
    });
  } catch (e) {
    throw new Error('フォルダが見つかりません。フォルダIDを確認してください。\n' + e.message);
  }

  // フォルダかどうか確認
  if (folderInfo.mimeType !== 'application/vnd.google-apps.folder') {
    throw new Error('指定されたIDはフォルダではありません。');
  }

  const folderName = folderInfo.name;
  const sheetName = customSheetName || (folderName + '_構造');

  // 既存シートがあれば削除して新規作成
  let sheet = ss.getSheetByName(sheetName);
  if (sheet) {
    ss.deleteSheet(sheet);
  }
  sheet = ss.insertSheet(sheetName);

  // ステータス表示
  SpreadsheetApp.getActiveSpreadsheet().toast('ファイル構造を取得中...', 'Processing', -1);

  // ファイル取得
  const files = getFilesFromFolder(folderId, folderName, folderInfo.driveId);

  // シートに書き込み
  writeToSheet(sheet, files);

  // 完了通知
  SpreadsheetApp.getActiveSpreadsheet().toast(
    files.length + '件のファイル/フォルダを出力しました',
    '完了',
    5
  );
}

// ===========================================
// Drive API関連
// ===========================================

/**
 * 共有ドライブ一覧を取得
 */
function getSharedDrives() {
  const drives = [];
  let pageToken = null;

  do {
    const response = Drive.Drives.list({
      pageSize: 100,
      pageToken: pageToken,
      fields: 'nextPageToken, drives(id, name)'
    });

    if (response.drives) {
      response.drives.forEach(function(drive) {
        drives.push({
          id: drive.id,
          name: drive.name
        });
      });
    }

    pageToken = response.nextPageToken;
  } while (pageToken);

  return drives;
}

/**
 * 共有ドライブからファイル取得
 */
function getFilesFromSharedDrive(driveId, driveName) {
  const allFiles = [];
  const pathMap = {};

  pathMap[driveId] = '';

  // ルートエントリ
  allFiles.push({
    id: driveId,
    name: driveName,
    mimeType: 'application/vnd.google-apps.folder',
    extension: '-',
    path: '/',
    depth: 0,
    parentId: null,
    size: '-',
    createdTime: '-',
    modifiedTime: '-',
    isFolder: true,
    url: 'https://drive.google.com/drive/folders/' + driveId
  });

  // 再帰的に取得
  fetchFilesRecursively(driveId, driveId, 1, allFiles, pathMap, true);

  return allFiles;
}

/**
 * マイドライブからファイル取得
 */
function getFilesFromMyDrive() {
  const allFiles = [];
  const pathMap = {};
  const rootId = 'root';

  pathMap[rootId] = '';

  // ルートエントリ
  allFiles.push({
    id: rootId,
    name: 'マイドライブ',
    mimeType: 'application/vnd.google-apps.folder',
    extension: '-',
    path: '/',
    depth: 0,
    parentId: null,
    size: '-',
    createdTime: '-',
    modifiedTime: '-',
    isFolder: true,
    url: 'https://drive.google.com/drive/my-drive'
  });

  // 再帰的に取得
  fetchFilesRecursively(null, rootId, 1, allFiles, pathMap, false);

  return allFiles;
}

/**
 * 任意のフォルダからファイル取得
 */
function getFilesFromFolder(folderId, folderName, driveId) {
  const allFiles = [];
  const pathMap = {};

  pathMap[folderId] = '';

  // ルートエントリ
  allFiles.push({
    id: folderId,
    name: folderName,
    mimeType: 'application/vnd.google-apps.folder',
    extension: '-',
    path: '/',
    depth: 0,
    parentId: null,
    size: '-',
    createdTime: '-',
    modifiedTime: '-',
    isFolder: true,
    url: 'https://drive.google.com/drive/folders/' + folderId
  });

  // 共有ドライブ内かどうかで処理を分岐
  const isSharedDrive = !!driveId;
  fetchFilesRecursively(driveId, folderId, 1, allFiles, pathMap, isSharedDrive);

  return allFiles;
}

/**
 * ファイルを再帰的に取得
 */
function fetchFilesRecursively(driveId, parentId, depth, allFiles, pathMap, isSharedDrive) {
  let pageToken = null;

  do {
    const query = "'" + parentId + "' in parents and trashed = false";

    const params = {
      q: query,
      pageSize: 1000,
      pageToken: pageToken,
      fields: 'nextPageToken, files(id, name, mimeType, size, createdTime, modifiedTime, parents)',
      orderBy: 'folder, name'
    };

    if (isSharedDrive) {
      params.driveId = driveId;
      params.corpora = 'drive';
      params.includeItemsFromAllDrives = true;
      params.supportsAllDrives = true;
    }

    const response = Drive.Files.list(params);

    if (response.files) {
      // フォルダを先に処理
      const folders = [];
      const files = [];

      response.files.forEach(function(file) {
        if (file.mimeType === 'application/vnd.google-apps.folder') {
          folders.push(file);
        } else {
          files.push(file);
        }
      });

      const sortedFiles = folders.concat(files);

      sortedFiles.forEach(function(file) {
        const parentPath = pathMap[parentId] || '';
        const currentPath = parentPath + '/' + file.name;
        pathMap[file.id] = currentPath;

        const isFolder = file.mimeType === 'application/vnd.google-apps.folder';

        const fileItem = {
          id: file.id,
          name: file.name,
          mimeType: file.mimeType || 'unknown',
          extension: isFolder ? '-' : getExtensionFromName(file.name),
          path: currentPath,
          depth: depth,
          parentId: parentId,
          size: isFolder ? '-' : formatFileSize(file.size),
          createdTime: formatDate(file.createdTime),
          modifiedTime: formatDate(file.modifiedTime),
          isFolder: isFolder,
          url: getFileUrl(file.id, file.mimeType)
        };

        allFiles.push(fileItem);

        if (isFolder) {
          fetchFilesRecursively(driveId, file.id, depth + 1, allFiles, pathMap, isSharedDrive);
        }
      });
    }

    pageToken = response.nextPageToken;
  } while (pageToken);
}

// ===========================================
// シート書き込み関連
// ===========================================

/**
 * シートにデータを書き込み
 */
function writeToSheet(sheet, files) {
  const headers = ['階層', '名前', '種別', 'カテゴリ', '拡張子', 'サイズ', '作成日', '更新日', 'パス', 'URL'];

  // データ準備
  const data = [headers];

  files.forEach(function(file) {
    const indent = file.depth > 0 ? '　'.repeat(file.depth - 1) + '└ ' : '';
    const category = getCategoryFromMimeType(file.mimeType);

    data.push([
      file.depth,
      indent + file.name,
      file.isFolder ? 'フォルダ' : 'ファイル',
      category,
      file.extension,
      file.size,
      file.createdTime,
      file.modifiedTime,
      file.path,
      file.url || '-'
    ]);
  });

  // 一括書き込み
  sheet.getRange(1, 1, data.length, headers.length).setValues(data);

  // 書式設定
  applyFormatting(sheet, files, headers.length);
}

/**
 * 書式を適用
 */
function applyFormatting(sheet, files, columnCount) {
  const lastRow = files.length + 1;

  // ヘッダー行の書式
  const headerRange = sheet.getRange(1, 1, 1, columnCount);
  headerRange
    .setBackground('#336699')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  // 列幅設定
  const columnWidths = [50, 400, 70, 120, 70, 80, 90, 90, 400, 350];
  columnWidths.forEach(function(width, index) {
    sheet.setColumnWidth(index + 1, width);
  });

  // 行の色分け（カテゴリ別）- バッチ処理で高速化
  if (files.length > 0) {
    const backgrounds = [];
    for (let i = 0; i < files.length; i++) {
      const category = getCategoryFromMimeType(files[i].mimeType);
      const color = getCategoryColor(category);
      backgrounds.push(Array(columnCount).fill(color));
    }
    sheet.getRange(2, 1, files.length, columnCount).setBackgrounds(backgrounds);
  }

  // ヘッダー行を固定
  sheet.setFrozenRows(1);

  // フィルターを設定
  if (lastRow > 1) {
    const dataRange = sheet.getRange(1, 1, lastRow, columnCount);
    dataRange.createFilter();
  }
}

// ===========================================
// ユーティリティ関数
// ===========================================

/**
 * MIMEタイプからカテゴリを取得
 */
function getCategoryFromMimeType(mimeType) {
  const categories = {
    'application/vnd.google-apps.folder': 'フォルダ',
    'application/vnd.google-apps.document': 'Google Docs',
    'application/vnd.google-apps.spreadsheet': 'Google Sheets',
    'application/vnd.google-apps.presentation': 'Google Slides',
    'application/vnd.google-apps.form': 'Google Forms',
    'application/vnd.google-apps.drawing': 'Google Drawing',
    'application/vnd.google-apps.site': 'Google Sites',
    'application/vnd.google-apps.map': 'Google Maps',
    'application/vnd.google-apps.shortcut': 'ショートカット',
    'application/pdf': 'PDF',
    'application/msword': 'Word',
    'application/vnd.openxmlformats-officedocument.wordprocessingml.document': 'Word',
    'application/vnd.oasis.opendocument.text': 'Word',
    'application/vnd.ms-excel': 'Excel',
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': 'Excel',
    'application/vnd.oasis.opendocument.spreadsheet': 'Excel',
    'text/csv': 'Excel',
    'application/vnd.ms-powerpoint': 'PowerPoint',
    'application/vnd.openxmlformats-officedocument.presentationml.presentation': 'PowerPoint',
    'application/vnd.oasis.opendocument.presentation': 'PowerPoint',
    'image/jpeg': '画像',
    'image/png': '画像',
    'image/gif': '画像',
    'image/svg+xml': '画像',
    'image/webp': '画像',
    'image/bmp': '画像',
    'image/tiff': '画像',
    'image/heic': '画像',
    'image/heif': '画像',
    'video/mp4': '動画',
    'video/quicktime': '動画',
    'video/x-msvideo': '動画',
    'video/x-matroska': '動画',
    'video/webm': '動画',
    'video/mpeg': '動画',
    'video/3gpp': '動画',
    'audio/mpeg': '音声',
    'audio/wav': '音声',
    'audio/aac': '音声',
    'audio/flac': '音声',
    'audio/ogg': '音声',
    'audio/mp4': '音声',
    'audio/x-m4a': '音声',
    'application/zip': 'アーカイブ',
    'application/x-tar': 'アーカイブ',
    'application/gzip': 'アーカイブ',
    'application/x-rar-compressed': 'アーカイブ',
    'application/x-7z-compressed': 'アーカイブ',
    'text/plain': 'テキスト',
    'text/rtf': 'テキスト',
    'text/html': 'コード',
    'text/css': 'コード',
    'text/xml': 'コード',
    'application/javascript': 'コード',
    'application/typescript': 'コード',
    'application/json': 'コード',
    'application/x-python': 'コード',
    'text/x-python': 'コード',
    'text/x-java-source': 'コード'
  };

  return categories[mimeType] || 'その他';
}

/**
 * カテゴリから背景色を取得
 */
function getCategoryColor(category) {
  const colors = {
    'フォルダ': '#fff3cd',
    'Google Docs': '#cce5ff',
    'Google Sheets': '#d4edda',
    'Google Slides': '#ffe5cc',
    'Google Forms': '#e2d5f1',
    'Google Drawing': '#fce4ec',
    'Google Sites': '#e0f7fa',
    'Google Maps': '#c8e6c9',
    'ショートカット': '#f5f5f5',
    'PDF': '#ffcdd2',
    'Word': '#bbdefb',
    'Excel': '#c8e6c9',
    'PowerPoint': '#ffe0b2',
    '画像': '#b2ebf2',
    '動画': '#f8bbd9',
    '音声': '#fff9c4',
    'アーカイブ': '#cfd8dc',
    'テキスト': '#fff8e1',
    'コード': '#dcedc8',
    'その他': '#f5f5f5'
  };

  return colors[category] || '#ffffff';
}

/**
 * ファイル名から拡張子を取得
 */
function getExtensionFromName(name) {
  if (!name) return '-';
  const parts = name.split('.');
  if (parts.length > 1) {
    return parts[parts.length - 1].toLowerCase();
  }
  return '-';
}

/**
 * ファイルサイズをフォーマット
 */
function formatFileSize(bytes) {
  if (!bytes) return '-';
  const size = parseInt(bytes, 10);
  if (isNaN(size)) return '-';

  const units = ['B', 'KB', 'MB', 'GB', 'TB'];
  let unitIndex = 0;
  let fileSize = size;

  while (fileSize >= 1024 && unitIndex < units.length - 1) {
    fileSize /= 1024;
    unitIndex++;
  }

  return fileSize.toFixed(1) + ' ' + units[unitIndex];
}

/**
 * 日付をフォーマット
 */
function formatDate(dateString) {
  if (!dateString) return '-';
  const date = new Date(dateString);
  const year = date.getFullYear();
  const month = ('0' + (date.getMonth() + 1)).slice(-2);
  const day = ('0' + date.getDate()).slice(-2);
  return year + '/' + month + '/' + day;
}

/**
 * HTMLエスケープ
 */
function escapeHtml(text) {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

/**
 * ファイルURLを生成
 */
function getFileUrl(fileId, mimeType) {
  if (!fileId) return '-';

  // Google形式のファイルは専用URLを使用
  const googleAppsUrls = {
    'application/vnd.google-apps.document': 'https://docs.google.com/document/d/' + fileId,
    'application/vnd.google-apps.spreadsheet': 'https://docs.google.com/spreadsheets/d/' + fileId,
    'application/vnd.google-apps.presentation': 'https://docs.google.com/presentation/d/' + fileId,
    'application/vnd.google-apps.form': 'https://docs.google.com/forms/d/' + fileId,
    'application/vnd.google-apps.drawing': 'https://docs.google.com/drawings/d/' + fileId,
    'application/vnd.google-apps.folder': 'https://drive.google.com/drive/folders/' + fileId
  };

  if (googleAppsUrls[mimeType]) {
    return googleAppsUrls[mimeType];
  }

  // その他のファイルは汎用URL
  return 'https://drive.google.com/file/d/' + fileId + '/view';
}

// ===========================================
// 定期更新設定関連
// ===========================================

const CONFIG_SHEET_NAME = '_config';

/**
 * 設定シートを作成または開く
 */
function openConfigSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let configSheet = ss.getSheetByName(CONFIG_SHEET_NAME);

  if (!configSheet) {
    configSheet = createConfigSheet(ss);
    SpreadsheetApp.getUi().alert(
      '設定シート作成完了',
      '設定シート「' + CONFIG_SHEET_NAME + '」を作成しました。\n\n' +
      '定期更新したいフォルダIDとシート名を入力してください。\n' +
      '有効列をTRUEにすると、その行が定期更新の対象になります。',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }

  // 設定シートをアクティブにする
  ss.setActiveSheet(configSheet);
}

/**
 * 設定シートを作成
 */
function createConfigSheet(ss) {
  const configSheet = ss.insertSheet(CONFIG_SHEET_NAME);

  // ヘッダー
  const headers = ['有効', 'フォルダID', 'シート名', '種別', '最終更新'];
  configSheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // ヘッダー書式
  configSheet.getRange(1, 1, 1, headers.length)
    .setBackground('#336699')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  // サンプル行
  const sampleData = [
    [true, '', 'マイドライブ_構造', 'マイドライブ', ''],
    [false, '（ここにフォルダIDを入力）', 'サンプル_構造', 'フォルダID', '']
  ];
  configSheet.getRange(2, 1, sampleData.length, headers.length).setValues(sampleData);

  // 列幅設定
  configSheet.setColumnWidth(1, 50);   // 有効
  configSheet.setColumnWidth(2, 350);  // フォルダID
  configSheet.setColumnWidth(3, 200);  // シート名
  configSheet.setColumnWidth(4, 100);  // 種別
  configSheet.setColumnWidth(5, 150);  // 最終更新

  // チェックボックス設定
  configSheet.getRange(2, 1, 100, 1).insertCheckboxes();

  // 説明行を追加
  configSheet.getRange(5, 1, 1, 5).setValues([[
    '【説明】', '有効=TRUE の行が定期更新対象', '種別: マイドライブ/共有ドライブID/フォルダID', '', ''
  ]]);
  configSheet.getRange(5, 1, 1, 5).setFontColor('#666666').setFontStyle('italic');

  // ヘッダー固定
  configSheet.setFrozenRows(1);

  return configSheet;
}

/**
 * 定期更新を有効化
 */
function enableScheduledUpdate() {
  const ui = SpreadsheetApp.getUi();

  // 既存のトリガーを削除
  deleteScheduleTriggers();

  // 頻度選択ダイアログ
  const result = ui.alert(
    '定期更新の頻度を選択',
    '毎日更新を設定しますか？\n\n' +
    '「はい」→ 毎日午前3時に更新\n' +
    '「いいえ」→ 毎週月曜午前3時に更新',
    ui.ButtonSet.YES_NO_CANCEL
  );

  if (result === ui.Button.CANCEL) {
    return;
  }

  // トリガーを作成
  if (result === ui.Button.YES) {
    ScriptApp.newTrigger('runScheduledUpdate')
      .timeBased()
      .everyDays(1)
      .atHour(3)
      .create();
    ui.alert('定期更新を有効化しました', '毎日午前3時に自動更新されます。', ui.ButtonSet.OK);
  } else {
    ScriptApp.newTrigger('runScheduledUpdate')
      .timeBased()
      .onWeekDay(ScriptApp.WeekDay.MONDAY)
      .atHour(3)
      .create();
    ui.alert('定期更新を有効化しました', '毎週月曜日の午前3時に自動更新されます。', ui.ButtonSet.OK);
  }
}

/**
 * 定期更新を無効化
 */
function disableScheduledUpdate() {
  deleteScheduleTriggers();
  SpreadsheetApp.getUi().alert(
    '定期更新を無効化しました',
    'すべての定期更新トリガーを削除しました。',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * スケジュールトリガーを削除
 */
function deleteScheduleTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'runScheduledUpdate') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}

/**
 * 定期更新を実行
 */
function runScheduledUpdate() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName(CONFIG_SHEET_NAME);

  if (!configSheet) {
    console.log('設定シートが見つかりません。');
    return;
  }

  const data = configSheet.getDataRange().getValues();
  let updatedCount = 0;

  // ヘッダー行をスキップして処理
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const enabled = row[0];
    const folderId = row[1];
    const sheetName = row[2];
    const type = row[3];

    // 空行や説明行をスキップ
    if (!sheetName || sheetName.startsWith('【')) {
      continue;
    }

    // 有効でない行はスキップ
    if (!enabled) {
      continue;
    }

    try {
      if (type === 'マイドライブ') {
        // マイドライブを処理
        processMyDriveWithSheetName(sheetName);
      } else if (folderId && folderId.length > 10 && !folderId.startsWith('（')) {
        // フォルダIDまたは共有ドライブIDを処理
        processFolderById(folderId, sheetName);
      } else {
        continue;
      }

      // 最終更新日時を記録
      const now = new Date();
      configSheet.getRange(i + 1, 5).setValue(Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm'));
      updatedCount++;

    } catch (e) {
      console.log('エラー（行' + (i + 1) + '）: ' + e.message);
      configSheet.getRange(i + 1, 5).setValue('エラー: ' + e.message);
    }
  }

  console.log('定期更新完了: ' + updatedCount + '件を更新しました。');
}

/**
 * マイドライブを指定シート名で処理
 */
function processMyDriveWithSheetName(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 既存シートがあれば削除して新規作成
  let sheet = ss.getSheetByName(sheetName);
  if (sheet) {
    ss.deleteSheet(sheet);
  }
  sheet = ss.insertSheet(sheetName);

  // ファイル取得
  const files = getFilesFromMyDrive();

  // シートに書き込み
  writeToSheet(sheet, files);
}
