// ==============================
// Google Photos → Drive 同期機能
// ==============================

/**
 * Google Photos Library APIを使用してアルバムから画像を取得し、
 * Google Driveにコピーして処理する
 */

// ==============================
// 設定と初期化
// ==============================

/**
 * Google Photos同期の初期設定
 */
function setupPhotosSyncConfig() {
  const props = PropertiesService.getScriptProperties();

  // デフォルト設定
  const defaultConfig = {
    'PHOTOS_ALBUM_NAME': 'ポケモンカード',  // 同期するアルバム名
    'PHOTOS_SYNC_ENABLED': 'true',           // 同期を有効化
    'PHOTOS_LAST_SYNC': '',                  // 最後の同期日時
    'PHOTOS_SYNC_INTERVAL': '1',             // 同期間隔（時間）
    'PHOTOS_MAX_ITEMS': '50'                 // 一度に処理する最大枚数
  };

  // 既存の設定を保持しながらデフォルト値を設定
  Object.keys(defaultConfig).forEach(key => {
    if (!props.getProperty(key)) {
      props.setProperty(key, defaultConfig[key]);
    }
  });

  console.log('Google Photos同期設定を初期化しました');

  return {
    albumName: props.getProperty('PHOTOS_ALBUM_NAME'),
    enabled: props.getProperty('PHOTOS_SYNC_ENABLED') === 'true',
    lastSync: props.getProperty('PHOTOS_LAST_SYNC'),
    syncInterval: parseInt(props.getProperty('PHOTOS_SYNC_INTERVAL')),
    maxItems: parseInt(props.getProperty('PHOTOS_MAX_ITEMS'))
  };
}

// ==============================
// Google Photos API関連
// ==============================

/**
 * Google Photos APIのアクセストークンを取得
 * @return {string} アクセストークン
 */
function getPhotosAccessToken() {
  try {
    const service = ScriptApp.getOAuthToken();
    return service;
  } catch (error) {
    console.error('アクセストークン取得エラー:', error);
    throw new Error('Google Photos APIへのアクセス権限が必要です');
  }
}

/**
 * 指定した名前のアルバムIDを取得
 * @param {string} albumName - アルバム名
 * @return {string|null} アルバムID
 */
function getAlbumIdByName(albumName) {
  const token = getPhotosAccessToken();
  const url = 'https://photoslibrary.googleapis.com/v1/albums';

  try {
    const response = UrlFetchApp.fetch(url, {
      headers: {
        'Authorization': 'Bearer ' + token,
        'Content-Type': 'application/json'
      },
      muteHttpExceptions: true
    });

    if (response.getResponseCode() !== 200) {
      console.error('アルバム一覧取得エラー:', response.getContentText());
      return null;
    }

    const data = JSON.parse(response.getContentText());
    const albums = data.albums || [];

    // 指定名のアルバムを探す
    const targetAlbum = albums.find(album => album.title === albumName);

    if (targetAlbum) {
      console.log(`アルバム "${albumName}" を見つけました: ${targetAlbum.id}`);
      return targetAlbum.id;
    } else {
      console.log(`アルバム "${albumName}" が見つかりません`);

      // アルバムを作成
      return createPhotosAlbum(albumName);
    }

  } catch (error) {
    console.error('アルバムID取得エラー:', error);
    return null;
  }
}

/**
 * Google Photosにアルバムを作成
 * @param {string} albumName - アルバム名
 * @return {string|null} 作成されたアルバムのID
 */
function createPhotosAlbum(albumName) {
  const token = getPhotosAccessToken();
  const url = 'https://photoslibrary.googleapis.com/v1/albums';

  try {
    const response = UrlFetchApp.fetch(url, {
      method: 'POST',
      headers: {
        'Authorization': 'Bearer ' + token,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify({
        album: {
          title: albumName
        }
      }),
      muteHttpExceptions: true
    });

    if (response.getResponseCode() === 200) {
      const album = JSON.parse(response.getContentText());
      console.log(`アルバム "${albumName}" を作成しました: ${album.id}`);
      return album.id;
    } else {
      console.error('アルバム作成エラー:', response.getContentText());
      return null;
    }

  } catch (error) {
    console.error('アルバム作成エラー:', error);
    return null;
  }
}

/**
 * アルバムから画像一覧を取得
 * @param {string} albumId - アルバムID
 * @param {number} maxResults - 取得する最大件数
 * @return {Array} 画像アイテムの配列
 */
function getPhotosFromAlbum(albumId, maxResults = 50) {
  const token = getPhotosAccessToken();
  const url = 'https://photoslibrary.googleapis.com/v1/mediaItems:search';

  const items = [];
  let pageToken = null;

  try {
    do {
      const payload = {
        albumId: albumId,
        pageSize: Math.min(maxResults - items.length, 100)
      };

      if (pageToken) {
        payload.pageToken = pageToken;
      }

      const response = UrlFetchApp.fetch(url, {
        method: 'POST',
        headers: {
          'Authorization': 'Bearer ' + token,
          'Content-Type': 'application/json'
        },
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      });

      if (response.getResponseCode() !== 200) {
        console.error('画像取得エラー:', response.getContentText());
        break;
      }

      const data = JSON.parse(response.getContentText());
      const mediaItems = data.mediaItems || [];

      items.push(...mediaItems);
      pageToken = data.nextPageToken;

    } while (pageToken && items.length < maxResults);

    console.log(`${items.length}枚の画像を取得しました`);
    return items;

  } catch (error) {
    console.error('画像一覧取得エラー:', error);
    return [];
  }
}

/**
 * Google Photos画像をGoogle Driveにダウンロード
 * @param {Object} mediaItem - Google PhotosのメディアアイテムAPI:
 * @param {string} targetFolderId - 保存先のDriveフォルダID
 * @return {Object|null} Driveファイル情報
 */
function downloadPhotoToDrive(mediaItem, targetFolderId) {
  try {
    // ダウンロードURL生成（最高品質）
    const downloadUrl = `${mediaItem.baseUrl}=d`;

    // 画像データを取得
    const response = UrlFetchApp.fetch(downloadUrl, {
      muteHttpExceptions: true
    });

    if (response.getResponseCode() !== 200) {
      console.error('画像ダウンロードエラー:', response.getResponseCode());
      return null;
    }

    const blob = response.getBlob();

    // ファイル名を生成（タイムスタンプ付き）
    const originalName = mediaItem.filename || 'untitled';
    const timestamp = new Date(mediaItem.mediaMetadata.creationTime).getTime();
    const fileName = `photos_${timestamp}_${originalName}`;

    blob.setName(fileName);

    // HEIC形式の場合はJPEGに変換
    if (fileName.toLowerCase().match(/\.(heic|heif)$/)) {
      console.log('HEIC形式を検出、JPEGに変換中...');
      const jpegBlob = blob.getAs('image/jpeg');
      jpegBlob.setName(fileName.replace(/\.(heic|heif)$/i, '.jpg'));
      blob = jpegBlob;
    }

    // Driveフォルダに保存
    const folder = DriveApp.getFolderById(targetFolderId);
    const file = folder.createFile(blob);

    // メタデータを設定
    file.setDescription(JSON.stringify({
      photosId: mediaItem.id,
      photosUrl: mediaItem.productUrl,
      creationTime: mediaItem.mediaMetadata.creationTime,
      width: mediaItem.mediaMetadata.width,
      height: mediaItem.mediaMetadata.height,
      syncedAt: new Date().toISOString()
    }));

    console.log(`画像をDriveに保存: ${file.getName()}`);

    return {
      id: file.getId(),
      name: file.getName(),
      url: file.getUrl(),
      photosId: mediaItem.id,
      mimeType: file.getMimeType(),
      size: file.getSize()
    };

  } catch (error) {
    console.error('画像ダウンロードエラー:', error);
    return null;
  }
}

// ==============================
// 同期処理
// ==============================

/**
 * Google PhotosアルバムをDriveと同期
 * @return {Object} 同期結果
 */
function syncPhotosAlbumToDrive() {
  const startTime = Date.now();
  console.log('=== Google Photos同期開始 ===');

  try {
    // 設定を取得
    const syncConfig = setupPhotosSyncConfig();

    if (!syncConfig.enabled) {
      console.log('Photos同期が無効になっています');
      return { success: false, message: '同期が無効' };
    }

    // アルバムIDを取得
    const albumId = getAlbumIdByName(syncConfig.albumName);
    if (!albumId) {
      throw new Error(`アルバム "${syncConfig.albumName}" が見つかりません`);
    }

    // Driveフォルダを準備
    const driveFolderId = getOrCreateSyncFolder();

    // 処理済みIDを取得
    const processedIds = getProcessedPhotosIds();

    // アルバムから画像を取得
    const photos = getPhotosFromAlbum(albumId, syncConfig.maxItems);

    // 新着画像をフィルタリング
    const newPhotos = photos.filter(photo => !processedIds.includes(photo.id));

    if (newPhotos.length === 0) {
      console.log('新着画像なし');
      updateLastSyncTime();
      return { success: true, message: '新着画像なし', count: 0 };
    }

    console.log(`新着画像: ${newPhotos.length}枚`);

    // 各画像をDriveにダウンロード
    const results = [];
    for (const photo of newPhotos) {
      try {
        console.log(`処理中: ${photo.filename}`);

        // Driveにダウンロード
        const driveFile = downloadPhotoToDrive(photo, driveFolderId);

        if (driveFile) {
          // 処理済みとしてマーク
          markPhotosAsProcessed(photo.id);

          results.push({
            success: true,
            photosId: photo.id,
            driveId: driveFile.id,
            fileName: driveFile.name
          });
        } else {
          results.push({
            success: false,
            photosId: photo.id,
            error: 'ダウンロード失敗'
          });
        }

      } catch (error) {
        console.error(`画像処理エラー: ${photo.filename}`, error);
        results.push({
          success: false,
          photosId: photo.id,
          error: error.toString()
        });
      }
    }

    // 同期時刻を更新
    updateLastSyncTime();

    // 結果サマリー
    const successCount = results.filter(r => r.success).length;
    const failureCount = results.filter(r => !r.success).length;
    const processingTime = (Date.now() - startTime) / 1000;

    console.log(`=== 同期完了 ===`);
    console.log(`成功: ${successCount}枚, 失敗: ${failureCount}枚`);
    console.log(`処理時間: ${processingTime}秒`);

    // 同期後にDrive画像を処理
    if (successCount > 0) {
      console.log('Drive画像処理を開始します...');
      processImagesFromDrive();
    }

    return {
      success: true,
      message: `${successCount}枚を同期`,
      successCount: successCount,
      failureCount: failureCount,
      processingTime: processingTime,
      results: results
    };

  } catch (error) {
    console.error('Photos同期エラー:', error);
    sendNotification('Google Photos同期エラー: ' + error.toString());
    return {
      success: false,
      message: error.toString()
    };
  }
}

/**
 * 同期用Driveフォルダを取得または作成
 * @return {string} フォルダID
 */
function getOrCreateSyncFolder() {
  let folderId = PropertiesService.getScriptProperties().getProperty('DRIVE_FOLDER_ID');

  if (!folderId) {
    const folder = DriveApp.createFolder('ポケモンカード_Photos同期');
    folderId = folder.getId();
    PropertiesService.getScriptProperties().setProperty('DRIVE_FOLDER_ID', folderId);
    console.log(`同期フォルダ作成: ${folder.getUrl()}`);
  }

  // Photos同期用のサブフォルダを作成
  const mainFolder = DriveApp.getFolderById(folderId);
  let syncFolder;

  const syncFolders = mainFolder.getFoldersByName('Photos同期');
  if (syncFolders.hasNext()) {
    syncFolder = syncFolders.next();
  } else {
    syncFolder = mainFolder.createFolder('Photos同期');
    console.log('Photos同期フォルダを作成しました');
  }

  return syncFolder.getId();
}

/**
 * 処理済みPhotos IDを取得
 * @return {Array<string>} 処理済みID配列
 */
function getProcessedPhotosIds() {
  const processedIds = PropertiesService.getScriptProperties().getProperty('PROCESSED_PHOTOS_IDS');
  return processedIds ? JSON.parse(processedIds) : [];
}

/**
 * Photos画像を処理済みとしてマーク
 * @param {string} photosId - Google Photos ID
 */
function markPhotosAsProcessed(photosId) {
  const processedIds = getProcessedPhotosIds();

  if (!processedIds.includes(photosId)) {
    processedIds.push(photosId);

    // 最大1000件まで保持（古いものから削除）
    if (processedIds.length > 1000) {
      processedIds.splice(0, processedIds.length - 1000);
    }

    PropertiesService.getScriptProperties().setProperty(
      'PROCESSED_PHOTOS_IDS',
      JSON.stringify(processedIds)
    );
  }
}

/**
 * 最終同期時刻を更新
 */
function updateLastSyncTime() {
  PropertiesService.getScriptProperties().setProperty(
    'PHOTOS_LAST_SYNC',
    new Date().toISOString()
  );
}

// ==============================
// トリガー設定
// ==============================

/**
 * Photos同期トリガーを設定
 */
function setupPhotosSyncTriggers() {
  // 既存のPhotos同期トリガーを削除
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'syncPhotosAlbumToDrive') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // 新しいトリガーを設定（1時間ごと）
  ScriptApp.newTrigger('syncPhotosAlbumToDrive')
    .timeBased()
    .everyHours(1)
    .create();

  console.log('Photos同期トリガー設定完了: 1時間ごと');
}

// ==============================
// 管理関数
// ==============================

/**
 * Photos同期状態を確認
 */
function checkPhotosSyncStatus() {
  const config = setupPhotosSyncConfig();

  console.log('=== Google Photos同期状態 ===');
  console.log(`同期有効: ${config.enabled}`);
  console.log(`アルバム名: ${config.albumName}`);
  console.log(`最終同期: ${config.lastSync || '未実行'}`);
  console.log(`同期間隔: ${config.syncInterval}時間`);
  console.log(`最大処理数: ${config.maxItems}枚/回`);

  const processedIds = getProcessedPhotosIds();
  console.log(`処理済み画像数: ${processedIds.length}枚`);

  // アルバムの存在確認
  const albumId = getAlbumIdByName(config.albumName);
  if (albumId) {
    console.log(`アルバムID: ${albumId}`);

    // アルバム内の画像数を確認
    const photos = getPhotosFromAlbum(albumId, 1);
    if (photos.length > 0) {
      console.log('アルバムへのアクセス: 成功');
    }
  } else {
    console.log('アルバム: 見つかりません');
  }

  return config;
}

/**
 * Photos同期設定を変更
 * @param {string} albumName - 同期するアルバム名
 * @param {boolean} enabled - 同期を有効化するか
 */
function configurePhotosSync(albumName, enabled = true) {
  const props = PropertiesService.getScriptProperties();

  if (albumName) {
    props.setProperty('PHOTOS_ALBUM_NAME', albumName);
  }

  props.setProperty('PHOTOS_SYNC_ENABLED', enabled.toString());

  console.log('Photos同期設定を更新しました');
  console.log(`アルバム名: ${albumName || '変更なし'}`);
  console.log(`同期有効: ${enabled}`);

  if (enabled) {
    // トリガーを設定
    setupPhotosSyncTriggers();
  }
}

/**
 * 手動で同期を実行
 */
function manualPhotosSync() {
  console.log('手動同期を開始します...');
  const result = syncPhotosAlbumToDrive();

  if (result.success) {
    console.log(`同期成功: ${result.message}`);
  } else {
    console.error(`同期失敗: ${result.message}`);
  }

  return result;
}

// ==============================
// 初期セットアップ
// ==============================

/**
 * Google Photos同期の初期セットアップ
 */
function setupPhotosSync() {
  console.log('=== Google Photos同期セットアップ ===');

  try {
    // 1. 設定を初期化
    const config = setupPhotosSyncConfig();
    console.log('1. 設定を初期化しました');

    // 2. アルバムを確認/作成
    const albumId = getAlbumIdByName(config.albumName);
    if (albumId) {
      console.log(`2. アルバム "${config.albumName}" を準備しました`);
    } else {
      console.log(`2. アルバム "${config.albumName}" の作成に失敗しました`);
      throw new Error('アルバムの作成に失敗しました');
    }

    // 3. Driveフォルダを準備
    const folderId = getOrCreateSyncFolder();
    const folder = DriveApp.getFolderById(folderId);
    console.log(`3. 同期フォルダを準備しました: ${folder.getUrl()}`);

    // 4. トリガーを設定
    setupPhotosSyncTriggers();
    console.log('4. 自動同期トリガーを設定しました');

    console.log('\n=== セットアップ完了 ===');
    console.log(`Google Photosアルバム "${config.albumName}" とDriveの同期が設定されました`);
    console.log('使い方:');
    console.log(`1. Google Photosで "${config.albumName}" アルバムに画像を追加`);
    console.log('2. 1時間ごとに自動同期されます');
    console.log('3. 手動同期: manualPhotosSync()を実行');

    return {
      success: true,
      albumName: config.albumName,
      folderId: folderId,
      folderUrl: folder.getUrl()
    };

  } catch (error) {
    console.error('セットアップエラー:', error);
    return {
      success: false,
      error: error.toString()
    };
  }
}