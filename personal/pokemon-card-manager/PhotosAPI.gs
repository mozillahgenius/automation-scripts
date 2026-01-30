// ==============================
// Google Photos API関連
// ==============================

function getPhotosAlbumInfo(albumId) {
  const token = getPhotosAccessToken();
  const url = `https://photoslibrary.googleapis.com/v1/albums/${albumId}`;

  const response = UrlFetchApp.fetch(url, {
    headers: {
      'Authorization': 'Bearer ' + token,
      'Content-Type': 'application/json'
    },
    muteHttpExceptions: true
  });

  if (response.getResponseCode() !== 200) {
    throw new Error(`アルバム情報取得失敗: ${response.getContentText()}`);
  }

  return JSON.parse(response.getContentText());
}

function getNewPhotosFromAlbum(albumId, processedIds) {
  const token = getPhotosAccessToken();
  const config = getConfig();
  const maxPhotos = config.MAX_PHOTOS_PER_RUN;

  const url = 'https://photoslibrary.googleapis.com/v1/mediaItems:search';

  const payload = {
    albumId: albumId,
    pageSize: Math.min(maxPhotos, 100) // APIの最大値は100
  };

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
    throw new Error(`写真取得失敗: ${response.getContentText()}`);
  }

  const data = JSON.parse(response.getContentText());
  const mediaItems = data.mediaItems || [];

  // 未処理の画像のみフィルタリング
  const newPhotos = mediaItems
    .filter(item => {
      // 画像のみ（動画は除外）
      if (!item.mimeType || !item.mimeType.startsWith('image/')) {
        return false;
      }
      // 未処理のもののみ
      return !processedIds.includes(item.id);
    })
    .slice(0, maxPhotos)
    .map(item => ({
      id: item.id,
      filename: item.filename,
      mimeType: item.mimeType,
      creationTime: item.mediaMetadata.creationTime,
      width: item.mediaMetadata.width,
      height: item.mediaMetadata.height,
      baseUrl: item.baseUrl,
      productUrl: item.productUrl
    }));

  return newPhotos;
}

function downloadPhotoData(photo) {
  // 最大解像度でダウンロード（w=幅, h=高さ, d=ダウンロード）
  const downloadUrl = `${photo.baseUrl}=w${photo.width}-h${photo.height}-d`;

  const response = UrlFetchApp.fetch(downloadUrl, {
    muteHttpExceptions: true
  });

  if (response.getResponseCode() !== 200) {
    throw new Error(`画像ダウンロード失敗: ${response.getResponseCode()}`);
  }

  return response.getBlob();
}

function savePhotoToDrive(photo, folderId) {
  // 画像データをダウンロード
  const blob = downloadPhotoData(photo);

  // ファイル名を設定
  const fileName = photo.filename || `card_${new Date().getTime()}.jpg`;
  blob.setName(fileName);

  // Driveフォルダに保存
  const folder = DriveApp.getFolderById(folderId);
  const file = folder.createFile(blob);

  // 説明を追加
  file.setDescription(`Google Photos ID: ${photo.id}\n作成日時: ${photo.creationTime}`);

  // 共有設定（閲覧リンクを取得可能に）
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  return {
    id: file.getId(),
    name: file.getName(),
    url: file.getUrl(),
    downloadUrl: file.getDownloadUrl(),
    viewUrl: `https://drive.google.com/file/d/${file.getId()}/view`,
    blob: file.getBlob()
  };
}

// ==============================
// OAuth2認証
// ==============================

function getPhotosAccessToken() {
  // Google Photos APIにアクセスするためのトークンを取得
  // GASの高度なGoogle サービスを有効にする必要があります

  // 方法1: OAuth2ライブラリを使用する場合
  const service = getPhotosOAuth2Service();
  if (service.hasAccess()) {
    return service.getAccessToken();
  } else {
    throw new Error('Google Photos APIへのアクセス権限がありません');
  }
}

function getPhotosOAuth2Service() {
  // OAuth2ライブラリを使用してサービスを作成
  // 事前にライブラリを追加する必要があります:
  // ライブラリID: 1B7FSrk5Zi6L1rSxxTDgDEUsPzlukDsi4KGuTMorsTQHhGBzBkMun4iDF

  const clientId = PropertiesService.getScriptProperties().getProperty('PHOTOS_CLIENT_ID');
  const clientSecret = PropertiesService.getScriptProperties().getProperty('PHOTOS_CLIENT_SECRET');

  return OAuth2.createService('GooglePhotos')
    .setAuthorizationBaseUrl('https://accounts.google.com/o/oauth2/v2/auth')
    .setTokenUrl('https://oauth2.googleapis.com/token')
    .setClientId(clientId)
    .setClientSecret(clientSecret)
    .setCallbackFunction('authCallback')
    .setPropertyStore(PropertiesService.getUserProperties())
    .setScope('https://www.googleapis.com/auth/photoslibrary.readonly')
    .setParam('access_type', 'offline')
    .setParam('prompt', 'consent');
}

function authCallback(request) {
  const service = getPhotosOAuth2Service();
  const isAuthorized = service.handleCallback(request);

  if (isAuthorized) {
    return HtmlService.createHtmlOutput('認証成功！このタブを閉じてください。');
  } else {
    return HtmlService.createHtmlOutput('認証失敗。もう一度お試しください。');
  }
}

function getAuthorizationUrl() {
  const service = getPhotosOAuth2Service();
  return service.getAuthorizationUrl();
}

// ==============================
// 簡易版（ScriptApp.getOAuthToken()を使用）
// ==============================

function getPhotosAccessTokenSimple() {
  // 簡易版: GASのデフォルト認証を使用
  // この方法を使う場合は、マニフェストファイル（appsscript.json）に
  // 以下のスコープを追加する必要があります：
  // "https://www.googleapis.com/auth/photoslibrary.readonly"

  try {
    return ScriptApp.getOAuthToken();
  } catch (error) {
    console.error('トークン取得エラー:', error);
    throw new Error('Google Photos APIアクセストークンの取得に失敗しました');
  }
}

// ==============================
// バッチ処理用
// ==============================

function batchDownloadPhotos(photos, folderId) {
  const results = [];

  for (const photo of photos) {
    try {
      const driveFile = savePhotoToDrive(photo, folderId);
      results.push({
        success: true,
        photo: photo,
        driveFile: driveFile
      });
    } catch (error) {
      console.error(`写真保存エラー: ${photo.filename}`, error);
      results.push({
        success: false,
        photo: photo,
        error: error.toString()
      });
    }

    // レート制限対策（1秒待機）
    Utilities.sleep(1000);
  }

  return results;
}

// ==============================
// デバッグ用
// ==============================

function testPhotosAPI() {
  const config = getConfig();
  const albumId = config.PHOTOS_ALBUM_ID;

  console.log('Google Photos APIテスト開始');

  try {
    // アルバム情報取得
    const albumInfo = getPhotosAlbumInfo(albumId);
    console.log('アルバム情報:', albumInfo);

    // 写真取得（最新5枚）
    const photos = getNewPhotosFromAlbum(albumId, []);
    console.log(`取得した写真: ${photos.length}枚`);

    if (photos.length > 0) {
      console.log('最初の写真:', photos[0]);
    }

    return {
      success: true,
      albumTitle: albumInfo.title,
      photoCount: photos.length
    };

  } catch (error) {
    console.error('Photos APIテストエラー:', error);
    return {
      success: false,
      error: error.toString()
    };
  }
}