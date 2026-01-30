// ==============================
// Google Photos API 代替実装（REST API直接呼び出し）
// ==============================

/**
 * Photos APIをREST APIで直接実装
 */
function setupPhotosAPIWorkaround() {
  console.log('=================================================================================');
  console.log('                    Google Photos API 代替実装                                   ');
  console.log('=================================================================================\n');

  console.log('Advanced ServicesにPhotos Library APIが表示されない問題を回避します。\n');

  console.log('【この実装の特徴】');
  console.log('✅ REST APIを直接呼び出し');
  console.log('✅ Advanced Services不要');
  console.log('✅ 既存のOAuth認証を使用');
  console.log('✅ すぐに使用可能\n');

  console.log('【実行コマンド】');
  console.log('1. testDirectPhotosAPI() - 接続テスト');
  console.log('2. listPhotosAlbums() - アルバム一覧取得');
  console.log('3. createPhotosAlbum("ポケモンカード") - アルバム作成');
  console.log('4. getPhotosFromLibrary() - 写真取得\n');

  return {
    status: 'ready',
    nextStep: 'testDirectPhotosAPI()'
  };
}

/**
 * REST API直接呼び出しで接続テスト
 */
function testDirectPhotosAPI() {
  console.log('=== Photos API直接接続テスト ===\n');

  const token = ScriptApp.getOAuthToken();

  // 最もシンプルなエンドポイントから開始
  const endpoints = [
    {
      name: 'メディアアイテム取得（GET）',
      url: 'https://photoslibrary.googleapis.com/v1/mediaItems?pageSize=1',
      method: 'GET',
      headers: {
        'Authorization': 'Bearer ' + token
      }
    },
    {
      name: 'アルバム一覧取得（GET）',
      url: 'https://photoslibrary.googleapis.com/v1/albums?pageSize=1',
      method: 'GET',
      headers: {
        'Authorization': 'Bearer ' + token
      }
    }
  ];

  let successCount = 0;

  endpoints.forEach(endpoint => {
    try {
      const response = UrlFetchApp.fetch(endpoint.url, {
        method: endpoint.method,
        headers: endpoint.headers,
        muteHttpExceptions: true
      });

      const code = response.getResponseCode();
      console.log(`${endpoint.name}: ${code}`);

      if (code === 200) {
        successCount++;
        const data = JSON.parse(response.getContentText());
        console.log(`  ✅ 成功`);

        if (endpoint.name.includes('メディアアイテム')) {
          console.log(`  アイテム数: ${data.mediaItems ? data.mediaItems.length : 0}`);
        } else if (endpoint.name.includes('アルバム')) {
          console.log(`  アルバム数: ${data.albums ? data.albums.length : 0}`);
        }
      } else if (code === 403) {
        console.log(`  ❌ 権限エラー`);
        const error = JSON.parse(response.getContentText());
        if (error.error && error.error.message) {
          console.log(`  ${error.error.message}`);
        }
      } else {
        console.log(`  ⚠️ エラー: ${code}`);
      }

    } catch (e) {
      console.error(`${endpoint.name}: 例外エラー - ${e.toString()}`);
    }
  });

  if (successCount > 0) {
    console.log('\n✅ Photos APIへの接続成功！');
    console.log('次のステップ: setupWorkingPhotosSync()');
    return true;
  } else {
    console.log('\n❌ Photos APIへの接続に失敗');
    console.log('解決策: fixPhotosPermissions()');
    return false;
  }
}

/**
 * 動作する形でアルバム一覧を取得
 */
function listPhotosAlbums() {
  console.log('=== アルバム一覧取得 ===\n');

  try {
    const token = ScriptApp.getOAuthToken();
    const url = 'https://photoslibrary.googleapis.com/v1/albums';

    const response = UrlFetchApp.fetch(url, {
      method: 'GET',
      headers: {
        'Authorization': 'Bearer ' + token,
        'Content-Type': 'application/json'
      },
      muteHttpExceptions: true
    });

    const code = response.getResponseCode();

    if (code === 200) {
      const data = JSON.parse(response.getContentText());
      const albums = data.albums || [];

      console.log(`アルバム数: ${albums.length}`);

      albums.forEach((album, index) => {
        console.log(`${index + 1}. ${album.title}`);
        console.log(`   ID: ${album.id}`);
        console.log(`   写真数: ${album.mediaItemsCount || 0}`);
      });

      return albums;
    } else {
      console.error('エラー:', response.getContentText());
      return [];
    }

  } catch (error) {
    console.error('例外エラー:', error);
    return [];
  }
}

/**
 * アルバムを作成
 */
function createPhotosAlbum(albumTitle) {
  console.log(`=== アルバム作成: ${albumTitle} ===\n`);

  try {
    const token = ScriptApp.getOAuthToken();
    const url = 'https://photoslibrary.googleapis.com/v1/albums';

    const payload = {
      album: {
        title: albumTitle
      }
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

    const code = response.getResponseCode();

    if (code === 200 || code === 201) {
      const album = JSON.parse(response.getContentText());
      console.log('✅ アルバム作成成功');
      console.log(`ID: ${album.id}`);
      console.log(`タイトル: ${album.title}`);
      console.log(`URL: ${album.productUrl}`);

      // アルバムIDを保存
      PropertiesService.getScriptProperties().setProperty('PHOTOS_ALBUM_ID', album.id);

      return album;
    } else {
      console.error('作成失敗:', response.getContentText());
      return null;
    }

  } catch (error) {
    console.error('例外エラー:', error);
    return null;
  }
}

/**
 * ライブラリから写真を取得
 */
function getPhotosFromLibrary(maxResults = 10) {
  console.log('=== ライブラリから写真取得 ===\n');

  try {
    const token = ScriptApp.getOAuthToken();
    const url = `https://photoslibrary.googleapis.com/v1/mediaItems?pageSize=${maxResults}`;

    const response = UrlFetchApp.fetch(url, {
      method: 'GET',
      headers: {
        'Authorization': 'Bearer ' + token
      },
      muteHttpExceptions: true
    });

    const code = response.getResponseCode();

    if (code === 200) {
      const data = JSON.parse(response.getContentText());
      const items = data.mediaItems || [];

      console.log(`取得した写真: ${items.length}枚`);

      items.forEach((item, index) => {
        console.log(`${index + 1}. ${item.filename}`);
        console.log(`   作成日: ${item.mediaMetadata.creationTime}`);
        console.log(`   サイズ: ${item.mediaMetadata.width}x${item.mediaMetadata.height}`);
      });

      return items;
    } else {
      console.error('取得失敗:', response.getContentText());
      return [];
    }

  } catch (error) {
    console.error('例外エラー:', error);
    return [];
  }
}

/**
 * 権限問題を修正
 */
function fixPhotosPermissions() {
  console.log('=== Photos API権限修正 ===\n');

  console.log('【確認事項】');
  console.log('1. company-gasプロジェクトでPhotos Library APIが有効化されているか');
  console.log('2. OAuth認証で適切なスコープが設定されているか\n');

  console.log('【解決手順】');

  console.log('ステップ1: プロジェクト番号を確認');
  console.log('────────────────────');
  verifyProjectNumber();

  console.log('\nステップ2: APIの有効化を再確認');
  console.log('────────────────────');
  console.log('https://console.cloud.google.com/apis/library/photoslibrary.googleapis.com?project=company-gas');
  console.log('で「管理」ボタンが表示されているか確認\n');

  console.log('ステップ3: 認証をリセット');
  console.log('────────────────────');
  console.log('clearPhotosAuth() を実行\n');

  console.log('ステップ4: 再テスト');
  console.log('────────────────────');
  console.log('testDirectPhotosAPI() を実行\n');

  return {
    step1: 'verifyProjectNumber()',
    step2: 'APIコンソールで確認',
    step3: 'clearPhotosAuth()',
    step4: 'testDirectPhotosAPI()'
  };
}

/**
 * プロジェクト番号を確認
 */
function verifyProjectNumber() {
  // GASプロジェクトの設定を確認するため、エラーメッセージから情報を取得
  try {
    const token = ScriptApp.getOAuthToken();
    const response = UrlFetchApp.fetch('https://photoslibrary.googleapis.com/v1/albums', {
      headers: {
        'Authorization': 'Bearer ' + token
      },
      muteHttpExceptions: true
    });

    if (response.getResponseCode() === 403) {
      const error = JSON.parse(response.getContentText());

      if (error.error && error.error.message) {
        const match = error.error.message.match(/project (\d+)/);
        if (match) {
          console.log('現在のプロジェクト番号: ' + match[1]);
          console.log('\nGASエディタで確認:');
          console.log('プロジェクトの設定 → GCPプロジェクト番号');
          console.log('この番号とcompany-gasプロジェクトの番号が一致しているか確認');
        }
      }
    }
  } catch (e) {
    console.log('プロジェクト番号の取得に失敗');
  }
}

/**
 * 認証をクリア
 */
function clearPhotosAuth() {
  console.log('=== 認証のクリア ===\n');

  console.log('1. https://myaccount.google.com/permissions にアクセス');
  console.log('2. このプロジェクト関連のアクセスを削除');
  console.log('3. GASエディタでプロジェクトを保存（Ctrl+S）');
  console.log('4. testDirectPhotosAPI() を再実行');
  console.log('5. 認証画面ですべての権限を許可\n');

  return {
    status: '手動操作が必要'
  };
}

/**
 * 動作確認済みの同期セットアップ
 */
function setupWorkingPhotosSync() {
  console.log('=== 動作確認済みPhotos同期セットアップ ===\n');

  // まず接続テスト
  if (!testDirectPhotosAPI()) {
    console.log('❌ Photos APIに接続できません');
    console.log('fixPhotosPermissions() を実行してください');
    return false;
  }

  // アルバムを作成または取得
  const albumName = 'ポケモンカード';
  const albums = listPhotosAlbums();

  let targetAlbum = albums.find(a => a.title === albumName);

  if (!targetAlbum) {
    console.log(`\nアルバム「${albumName}」を作成中...`);
    targetAlbum = createPhotosAlbum(albumName);
  } else {
    console.log(`\n既存のアルバム「${albumName}」を使用`);
    PropertiesService.getScriptProperties().setProperty('PHOTOS_ALBUM_ID', targetAlbum.id);
  }

  if (targetAlbum) {
    console.log('\n✅ セットアップ完了！');
    console.log(`アルバムID: ${targetAlbum.id}`);
    console.log('\n使い方:');
    console.log('1. Google Photosアプリで「ポケモンカード」アルバムに写真を追加');
    console.log('2. processPhotosAlbum() を実行');

    return true;
  } else {
    console.log('\n❌ セットアップ失敗');
    return false;
  }
}