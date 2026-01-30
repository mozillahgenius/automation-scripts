// ==============================
// Google Photos API デバッグ
// ==============================

/**
 * 詳細なデバッグ情報を取得
 */
function debugPhotosAPI() {
  console.log('=== Google Photos API デバッグ ===\n');

  // 1. プロジェクト情報
  console.log('【プロジェクト情報】');
  console.log('スクリプトID: ' + ScriptApp.getScriptId());

  // 2. 有効なスコープを確認
  console.log('\n【OAuth スコープ確認】');
  try {
    const token = ScriptApp.getOAuthToken();
    console.log('トークン取得: 成功');
    console.log('トークン長: ' + token.length);
  } catch (e) {
    console.error('トークン取得エラー:', e);
  }

  // 3. 異なるエンドポイントをテスト
  console.log('\n【エンドポイントテスト】');

  const endpoints = [
    {
      name: 'Albums (GET)',
      url: 'https://photoslibrary.googleapis.com/v1/albums',
      method: 'GET'
    },
    {
      name: 'MediaItems (GET)',
      url: 'https://photoslibrary.googleapis.com/v1/mediaItems',
      method: 'GET'
    },
    {
      name: 'SharedAlbums (GET)',
      url: 'https://photoslibrary.googleapis.com/v1/sharedAlbums',
      method: 'GET'
    }
  ];

  endpoints.forEach(endpoint => {
    testEndpoint(endpoint);
  });

  // 4. 完全なエラー情報を取得
  console.log('\n【詳細エラー情報】');
  getDetailedError();

  return {
    status: 'デバッグ完了',
    nextStep: 'fixPhotosAPIAuth()'
  };
}

/**
 * 個別のエンドポイントをテスト
 */
function testEndpoint(endpoint) {
  try {
    const token = ScriptApp.getOAuthToken();
    const options = {
      method: endpoint.method,
      headers: {
        'Authorization': 'Bearer ' + token,
        'Content-Type': 'application/json'
      },
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(endpoint.url, options);
    const code = response.getResponseCode();

    console.log(`${endpoint.name}: ${code}`);

    if (code !== 200) {
      const error = JSON.parse(response.getContentText());
      if (error.error && error.error.message) {
        console.log(`  エラー: ${error.error.message}`);
      }
    }
  } catch (e) {
    console.log(`${endpoint.name}: エラー - ${e.toString()}`);
  }
}

/**
 * 詳細なエラー情報を取得
 */
function getDetailedError() {
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

      console.log('エラー詳細:');
      console.log(JSON.stringify(error, null, 2));

      // エラーの種類を判定
      if (error.error) {
        const msg = error.error.message;

        if (msg.includes('has not been used in project')) {
          console.log('\n診断: APIが有効化されていません');
          console.log('解決策: enablePhotosAPIInCompanyGas() を実行');
        } else if (msg.includes('insufficient authentication scopes')) {
          console.log('\n診断: スコープ不足');
          console.log('解決策: fixPhotosAPIAuth() を実行');
        } else if (msg.includes('Request had insufficient authentication scopes')) {
          console.log('\n診断: OAuth認証のスコープが不足');
          console.log('解決策: resetAndReauthorize() を実行');
        } else {
          console.log('\n診断: 不明なエラー');
        }
      }
    }
  } catch (e) {
    console.error('詳細エラー取得失敗:', e);
  }
}

/**
 * Photos API認証を修正
 */
function fixPhotosAPIAuth() {
  console.log('=== Photos API認証の修正 ===\n');

  console.log('【解決方法1: スコープをリセット】');
  console.log('resetAndReauthorize() を実行\n');

  console.log('【解決方法2: 別の認証方法を使用】');
  console.log('useServiceAccount() を実行\n');

  console.log('【解決方法3: Advanced Google Servicesを使用】');
  console.log('enablePhotosAdvancedService() を実行\n');

  return {
    option1: 'resetAndReauthorize()',
    option2: 'useServiceAccount()',
    option3: 'enablePhotosAdvancedService()'
  };
}

/**
 * 認証をリセットして再認証
 */
function resetAndReauthorize() {
  console.log('=== 認証のリセットと再認証 ===\n');

  console.log('【手順】\n');

  console.log('1. 既存の認証を削除');
  console.log('────────────────────');
  console.log('https://myaccount.google.com/permissions');
  console.log('でこのプロジェクトのアクセスを削除\n');

  console.log('2. GASエディタでプロジェクトを保存');
  console.log('────────────────────');
  console.log('Ctrl+S / Cmd+S\n');

  console.log('3. テスト関数を実行');
  console.log('────────────────────');
  console.log('testPhotosWithNewAuth() を実行\n');

  console.log('4. 認証画面で以下を確認');
  console.log('────────────────────');
  console.log('✓ Google Photos関連の権限がすべて表示される');
  console.log('✓ すべてにチェックを入れて許可\n');

  return {
    nextStep: 'testPhotosWithNewAuth()'
  };
}

/**
 * 新しい認証でテスト
 */
function testPhotosWithNewAuth() {
  // まずDriveで認証を強制
  DriveApp.getRootFolder();

  // 次にPhotos APIをテスト
  try {
    const token = ScriptApp.getOAuthToken();

    // メディアアイテムを検索（アルバム不要）
    const url = 'https://photoslibrary.googleapis.com/v1/mediaItems';

    const response = UrlFetchApp.fetch(url, {
      headers: {
        'Authorization': 'Bearer ' + token,
        'Content-Type': 'application/json'
      },
      muteHttpExceptions: true
    });

    const code = response.getResponseCode();
    console.log('Photos API レスポンス: ' + code);

    if (code === 200) {
      console.log('✅ 成功！Photos APIに接続できました');
      console.log('\n次: setupPhotosSync() を実行');
      return true;
    } else {
      const error = JSON.parse(response.getContentText());
      console.log('エラー: ' + JSON.stringify(error, null, 2));

      if (code === 403) {
        console.log('\n追加の対処法:');
        console.log('tryAlternativePhotosAccess() を実行');
      }
      return false;
    }
  } catch (e) {
    console.error('エラー:', e);
    return false;
  }
}

/**
 * Advanced Google Servicesを有効化
 */
function enablePhotosAdvancedService() {
  console.log('=== Advanced Google Services経由でPhotos APIを使用 ===\n');

  console.log('【手順】\n');

  console.log('1. GASエディタで「サービス」を開く');
  console.log('────────────────────');
  console.log('左メニューの「サービス」（＋マーク）をクリック\n');

  console.log('2. Google Photos Library APIを追加');
  console.log('────────────────────');
  console.log('一覧から「Google Photos Library API」を探す');
  console.log('※見つからない場合は、一覧の最下部を確認\n');

  console.log('3. 追加ボタンをクリック');
  console.log('────────────────────');
  console.log('サービスID: PhotosLibrary');
  console.log('バージョン: v1\n');

  console.log('4. 追加後、テスト実行');
  console.log('────────────────────');
  console.log('testPhotosAdvancedService() を実行\n');

  return {
    status: '手動設定が必要',
    nextStep: 'GASエディタで設定後、testPhotosAdvancedService()を実行'
  };
}

/**
 * 代替アクセス方法
 */
function tryAlternativePhotosAccess() {
  console.log('=== 代替Photos APIアクセス方法 ===\n');

  try {
    // OAuth 2.0を直接使用
    const token = ScriptApp.getOAuthToken();

    // シンプルなリクエストから開始
    const url = 'https://photoslibrary.googleapis.com/v1/mediaItems?pageSize=1';

    const response = UrlFetchApp.fetch(url, {
      method: 'GET',
      headers: {
        'Authorization': 'Bearer ' + token
      },
      muteHttpExceptions: true
    });

    const code = response.getResponseCode();
    console.log('レスポンスコード: ' + code);

    if (code === 200) {
      console.log('✅ 代替方法で成功！');
      const data = JSON.parse(response.getContentText());
      console.log('メディアアイテム数: ' + (data.mediaItems ? data.mediaItems.length : 0));

      console.log('\n次のステップ:');
      console.log('useAlternativePhotosSync() を実行');
      return true;
    } else {
      console.log('代替方法も失敗: ' + response.getContentText());

      console.log('\n最終手段:');
      console.log('usePhotosAPIWorkaround() を実行');
      return false;
    }
  } catch (e) {
    console.error('エラー:', e);
    return false;
  }
}