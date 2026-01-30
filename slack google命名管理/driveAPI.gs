/**
 * Google Drive API 連携機能
 * 共有ドライブと全階層フォルダの取得、スプレッドシートへの保存
 */

// =================================================================
// Drive API メイン関数
// =================================================================
function fetchGoogleDriveData() {
  const startTime = new Date();
  console.log('Google Driveデータ取得を開始...');
  
  try {
    // 設定読み込み
    const config = getConfig();
    
    // 共有ドライブ一覧を取得
    const sharedDrives = fetchSharedDrives(config);
    
    // 各ドライブのフォルダを取得
    const allFolders = fetchAllFolders(sharedDrives, config);
    
    const executionTime = (new Date() - startTime) / 1000;
    console.log(`Driveデータ取得完了: ドライブ${sharedDrives.length}件, フォルダ${allFolders.length}件 (${executionTime}秒)`);
    
    return {
      sharedDrives: sharedDrives,
      folders: allFolders
    };
    
  } catch (error) {
    console.error('Driveデータ取得エラー:', error);
    throw error;
  }
}

// =================================================================
// 共有ドライブ一覧取得
// =================================================================
function fetchSharedDrives(config) {
  console.log('共有ドライブ一覧を取得中...');
  
  if (config.driveTarget === 'MY_DRIVE') {
    // マイドライブのみの場合
    return [];
  }
  
  const sharedDrives = [];
  let pageToken = null;
  
  do {
    try {
      const response = Drive.Drives.list({
        pageSize: 100,
        pageToken: pageToken,
        fields: 'nextPageToken, drives(id, name, createdTime)'
      });
      
      if (response.drives && response.drives.length > 0) {
        sharedDrives.push(...response.drives);
      }
      
      pageToken = response.nextPageToken;
      
    } catch (error) {
      console.error('共有ドライブ一覧取得エラー:', error);
      throw error;
    }
  } while (pageToken);
  
  // スプレッドシートに保存
  saveSharedDrivesToSheet(sharedDrives);
  
  return sharedDrives;
}

// =================================================================
// 全フォルダ取得（全階層）
// =================================================================
function fetchAllFolders(sharedDrives, config) {
  console.log('全フォルダを取得中...');
  
  const allFolders = [];
  const maxDepth = parseInt(config.maxFolderDepth) || 10;
  const excludeRegex = config.excludePathRegex ? new RegExp(config.excludePathRegex) : null;
  
  // 共有ドライブのフォルダを取得
  if (config.driveTarget === 'SHARED_DRIVES' || config.driveTarget === 'BOTH') {
    for (const drive of sharedDrives) {
      const driveFolders = fetchFoldersFromDrive(drive.id, drive.name, maxDepth, excludeRegex);
      allFolders.push(...driveFolders);
      
      // 実行時間制限チェック
      if (isExecutionTimeLimitNear()) {
        console.log('実行時間制限が近づいたため、処理を中断');
        saveCheckpoint({ lastProcessedDrive: drive.id });
        break;
      }
    }
  }
  
  // マイドライブのフォルダを取得
  if (config.driveTarget === 'MY_DRIVE' || config.driveTarget === 'BOTH') {
    const myDriveFolders = fetchFoldersFromDrive('root', 'My Drive', maxDepth, excludeRegex);
    allFolders.push(...myDriveFolders);
  }
  
  // スプレッドシートに保存
  saveFoldersToSheet(allFolders);
  
  return allFolders;
}

// =================================================================
// 特定ドライブのフォルダ取得
// =================================================================
function fetchFoldersFromDrive(driveId, driveName, maxDepth, excludeRegex) {
  console.log(`ドライブ "${driveName}" のフォルダを取得中...`);
  
  const folders = [];
  const folderMap = {}; // ID -> フォルダ情報のマッピング
  let pageToken = null;
  
  // クエリ構築
  const query = "mimeType='application/vnd.google-apps.folder' and trashed=false";
  
  do {
    try {
      const options = {
        q: query,
        pageSize: 100,
        pageToken: pageToken,
        fields: 'nextPageToken, files(id, name, parents, createdTime, modifiedTime)',
        supportsAllDrives: true,
        includeItemsFromAllDrives: true
      };
      
      // 共有ドライブの場合
      if (driveId !== 'root') {
        options.driveId = driveId;
        options.corpora = 'drive';
      } else {
        // マイドライブの場合
        options.corpora = 'user';
      }
      
      const response = Drive.Files.list(options);
      
      if (response.files && response.files.length > 0) {
        for (const file of response.files) {
          folderMap[file.id] = {
            id: file.id,
            name: file.name,
            parentId: file.parents ? file.parents[0] : null,
            driveId: driveId,
            driveName: driveName,
            createdTime: file.createdTime,
            modifiedTime: file.modifiedTime,
            fullPath: null, // 後で構築
            depth: null     // 後で計算
          };
        }
      }
      
      pageToken = response.nextPageToken;
      
    } catch (error) {
      console.error(`フォルダ取得エラー (${driveName}):`, error);
      throw error;
    }
  } while (pageToken);
  
  // フルパスと階層深度を構築
  buildFolderPaths(folderMap, driveId, driveName, maxDepth, excludeRegex);
  
  // マップから配列に変換
  for (const folderId in folderMap) {
    const folder = folderMap[folderId];
    if (folder.fullPath && folder.depth <= maxDepth) {
      // 除外パスチェック
      if (!excludeRegex || !excludeRegex.test(folder.fullPath)) {
        folders.push(folder);
      }
    }
  }
  
  return folders;
}

// =================================================================
// フォルダパス構築
// =================================================================
function buildFolderPaths(folderMap, rootId, rootName, maxDepth, excludeRegex) {
  // 各フォルダのフルパスを再帰的に構築
  for (const folderId in folderMap) {
    if (!folderMap[folderId].fullPath) {
      buildSingleFolderPath(folderMap[folderId], folderMap, rootId, rootName, 0, maxDepth);
    }
  }
}

// =================================================================
// 単一フォルダのパス構築（再帰）
// =================================================================
function buildSingleFolderPath(folder, folderMap, rootId, rootName, currentDepth, maxDepth) {
  if (currentDepth > maxDepth) {
    return null;
  }
  
  // 既にパスが構築されている場合
  if (folder.fullPath) {
    return folder.fullPath;
  }
  
  // ルートレベルのフォルダ
  if (!folder.parentId || folder.parentId === rootId) {
    folder.fullPath = `/${rootName}/${folder.name}`;
    folder.depth = 1;
    return folder.fullPath;
  }
  
  // 親フォルダのパスを取得
  const parentFolder = folderMap[folder.parentId];
  if (!parentFolder) {
    // 親フォルダが見つからない場合（アクセス権限など）
    folder.fullPath = `/${rootName}/[Unknown]/${folder.name}`;
    folder.depth = 2;
    return folder.fullPath;
  }
  
  // 親フォルダのパスを再帰的に構築
  const parentPath = buildSingleFolderPath(parentFolder, folderMap, rootId, rootName, currentDepth + 1, maxDepth);
  
  if (parentPath) {
    folder.fullPath = `${parentPath}/${folder.name}`;
    folder.depth = parentFolder.depth + 1;
    return folder.fullPath;
  }
  
  return null;
}

// =================================================================
// 共有ドライブ情報をシートに保存
// =================================================================
function saveSharedDrivesToSheet(sharedDrives) {
  const spreadsheet = SpreadsheetApp.openById(getSpreadsheetId());
  const sheet = spreadsheet.getSheetByName('Drive_SharedDrives');
  
  if (!sheet) {
    throw new Error('Drive_SharedDrives シートが見つかりません');
  }
  
  // 既存データをクリア
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clear();
  }
  
  // データ整形
  const now = new Date();
  const data = sharedDrives.map(drive => [
    drive.id,
    drive.name,
    drive.createdTime ? new Date(drive.createdTime) : '',
    now,
    '', // Violation
    '', // ViolationType
    '', // ViolationMessage
    ''  // MatchedRule
  ]);
  
  if (data.length > 0) {
    sheet.getRange(2, 1, data.length, data[0].length).setValues(data);
  }
  
  // 違反判定を実行
  validateSharedDrives();
}

// =================================================================
// フォルダ情報をシートに保存
// =================================================================
function saveFoldersToSheet(folders) {
  const spreadsheet = SpreadsheetApp.openById(getSpreadsheetId());
  const sheet = spreadsheet.getSheetByName('Drive_Folders');
  
  if (!sheet) {
    throw new Error('Drive_Folders シートが見つかりません');
  }
  
  // 既存データをクリア
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clear();
  }
  
  // データ整形
  const now = new Date();
  const data = folders.map(folder => [
    folder.id,
    folder.name,
    folder.parentId || '',
    folder.driveId,
    folder.fullPath,
    folder.depth,
    folder.createdTime ? new Date(folder.createdTime) : '',
    folder.modifiedTime ? new Date(folder.modifiedTime) : '',
    now,
    '', // Violation
    '', // ViolationType
    '', // ViolationMessage
    ''  // MatchedRule
  ]);
  
  if (data.length > 0) {
    sheet.getRange(2, 1, data.length, data[0].length).setValues(data);
  }
  
  // 違反判定を実行
  validateDriveFolders();
}

// =================================================================
// 共有ドライブ名違反判定
// =================================================================
function validateSharedDrives() {
  const spreadsheet = SpreadsheetApp.openById(getSpreadsheetId());
  const drivesSheet = spreadsheet.getSheetByName('Drive_SharedDrives');
  const rulesSheet = spreadsheet.getSheetByName('Rules');
  const whitelistSheet = spreadsheet.getSheetByName('Whitelist');
  
  if (!drivesSheet || !rulesSheet || !whitelistSheet) {
    throw new Error('必要なシートが見つかりません');
  }
  
  // ルールとホワイトリストを取得
  const rules = getDriveRules(rulesSheet);
  const whitelist = getDriveWhitelist(whitelistSheet);
  
  // ドライブデータを取得
  const lastRow = drivesSheet.getLastRow();
  if (lastRow <= 1) return;
  
  const drivesData = drivesSheet.getRange(2, 1, lastRow - 1, 8).getValues();
  const validationResults = [];
  
  // 各ドライブを判定
  for (let i = 0; i < drivesData.length; i++) {
    const driveName = drivesData[i][1]; // DriveName列
    const result = validateName(driveName, rules, whitelist, 'Drive');
    
    validationResults.push([
      result.violation ? 'TRUE' : 'FALSE',
      result.violationType || '',
      result.violationMessage || '',
      result.matchedRule || ''
    ]);
  }
  
  // 結果をシートに書き込み
  if (validationResults.length > 0) {
    drivesSheet.getRange(2, 5, validationResults.length, 4).setValues(validationResults);
  }
}

// =================================================================
// フォルダ名違反判定
// =================================================================
function validateDriveFolders() {
  const spreadsheet = SpreadsheetApp.openById(getSpreadsheetId());
  const foldersSheet = spreadsheet.getSheetByName('Drive_Folders');
  const rulesSheet = spreadsheet.getSheetByName('Rules');
  const whitelistSheet = spreadsheet.getSheetByName('Whitelist');
  
  if (!foldersSheet || !rulesSheet || !whitelistSheet) {
    throw new Error('必要なシートが見つかりません');
  }
  
  // ルールとホワイトリストを取得
  const rules = getDriveFolderRules(rulesSheet);
  const whitelist = getDriveFolderWhitelist(whitelistSheet);
  
  // フォルダデータを取得
  const lastRow = foldersSheet.getLastRow();
  if (lastRow <= 1) return;
  
  const foldersData = foldersSheet.getRange(2, 1, lastRow - 1, 13).getValues();
  const validationResults = [];
  
  // 各フォルダを判定
  for (let i = 0; i < foldersData.length; i++) {
    const folderName = foldersData[i][1]; // FolderName列
    const result = validateName(folderName, rules, whitelist, 'DriveFolder');
    
    validationResults.push([
      result.violation ? 'TRUE' : 'FALSE',
      result.violationType || '',
      result.violationMessage || '',
      result.matchedRule || ''
    ]);
  }
  
  // 結果をシートに書き込み
  if (validationResults.length > 0) {
    foldersSheet.getRange(2, 10, validationResults.length, 4).setValues(validationResults);
  }
}

// =================================================================
// Drive用ルール取得
// =================================================================
function getDriveRules(rulesSheet) {
  const lastRow = rulesSheet.getLastRow();
  if (lastRow <= 1) return [];
  
  const rulesData = rulesSheet.getRange(2, 1, lastRow - 1, 7).getValues();
  
  return rulesData
    .filter(row => row[1] === 'Drive' && row[6] === 'TRUE')
    .map(row => ({
      ruleName: row[0],
      regex: row[2],
      severity: row[3],
      priority: row[4],
      description: row[5]
    }))
    .sort((a, b) => a.priority - b.priority);
}

// =================================================================
// DriveFolder用ルール取得
// =================================================================
function getDriveFolderRules(rulesSheet) {
  const lastRow = rulesSheet.getLastRow();
  if (lastRow <= 1) return [];
  
  const rulesData = rulesSheet.getRange(2, 1, lastRow - 1, 7).getValues();
  
  return rulesData
    .filter(row => row[1] === 'DriveFolder' && row[6] === 'TRUE')
    .map(row => ({
      ruleName: row[0],
      regex: row[2],
      severity: row[3],
      priority: row[4],
      description: row[5]
    }))
    .sort((a, b) => a.priority - b.priority);
}

// =================================================================
// Drive用ホワイトリスト取得
// =================================================================
function getDriveWhitelist(whitelistSheet) {
  const lastRow = whitelistSheet.getLastRow();
  if (lastRow <= 1) return [];
  
  const whitelistData = whitelistSheet.getRange(2, 1, lastRow - 1, 5).getValues();
  const now = new Date();
  
  return whitelistData
    .filter(row => {
      if (row[0] !== 'Drive') return false;
      if (row[4] && new Date(row[4]) < now) return false;
      return true;
    })
    .map(row => ({
      pattern: row[1],
      isRegex: row[2] === 'TRUE',
      reason: row[3]
    }));
}

// =================================================================
// DriveFolder用ホワイトリスト取得
// =================================================================
function getDriveFolderWhitelist(whitelistSheet) {
  const lastRow = whitelistSheet.getLastRow();
  if (lastRow <= 1) return [];
  
  const whitelistData = whitelistSheet.getRange(2, 1, lastRow - 1, 5).getValues();
  const now = new Date();
  
  return whitelistData
    .filter(row => {
      if (row[0] !== 'DriveFolder') return false;
      if (row[4] && new Date(row[4]) < now) return false;
      return true;
    })
    .map(row => ({
      pattern: row[1],
      isRegex: row[2] === 'TRUE',
      reason: row[3]
    }));
}