// ==============================
// Google Photos API セットアップヘルパー
// ==============================

/**
 * Google Photos APIの設定状況を詳細に確認
 */
function checkPhotosAPISetup() {
  console.log('=== Google Photos API設定確認 ===\n');

  // 1. 現在のプロジェクト情報を表示
  try {
    console.log('【GASプロジェクト情報】');
    console.log('プロジェクトID: 138255947511');
    console.log('スクリプトID: ' + ScriptApp.getScriptId());
  } catch (e) {
    console.error('プロジェクト情報取得エラー:', e);
  }

  // 2. OAuth スコープを確認
  console.log('\n【OAuth スコープ】');
  console.log('必要なスコープ:');
  console.log('✓ https://www.googleapis.com/auth/photoslibrary');
  console.log('✓ https://www.googleapis.com/auth/photoslibrary.readonly');

  // 3. API接続テスト
  console.log('\n【API接続テスト】');
  testPhotosAPIConnection();

  return {
    projectId: '138255947511',
    scriptId: ScriptApp.getScriptId()
  };
}

/**
 * Photos API接続をテスト（詳細エラー情報付き）
 */
function testPhotosAPIConnection() {
  try {
    const token = ScriptApp.getOAuthToken();
    console.log('OAuth トークン取得: 成功');

    const response = UrlFetchApp.fetch('https://photoslibrary.googleapis.com/v1/albums', {
      headers: {
        'Authorization': 'Bearer ' + token,
        'Content-Type': 'application/json'
      },
      muteHttpExceptions: true
    });

    const code = response.getResponseCode();
    console.log('レスポンスコード: ' + code);

    if (code === 200) {
      console.log('✅ Google Photos API: 接続成功！');
      const data = JSON.parse(response.getContentText());
      console.log('アルバム数: ' + (data.albums ? data.albums.length : 0));
      return true;
    } else if (code === 403) {
      console.log('❌ Google Photos API: 無効化されています');
      const error = JSON.parse(response.getContentText());

      if (error.error && error.error.details) {
        const details = error.error.details[0];
        if (details && details.metadata) {
          console.log('\n【有効化に必要な手順】');
          console.log('1. 以下のURLにアクセス:');
          console.log('   ' + details.metadata.activationUrl);
          console.log('2. 「有効にする」ボタンをクリック');
          console.log('3. 5分待機後、再度テストを実行');
        }
      }
      return false;
    } else {
      console.log('⚠️ 予期しないレスポンス: ' + code);
      console.log(response.getContentText());
      return false;
    }
  } catch (error) {
    console.error('接続エラー:', error);
    return false;
  }
}

/**
 * GCPプロジェクトの設定手順を表示
 */
function showGCPSetupInstructions() {
  console.log('=== GCPプロジェクト設定手順 ===\n');

  console.log('【方法A: 既存プロジェクト(138255947511)でAPIを有効化】');
  console.log('1. 以下のURLにアクセス:');
  console.log('   https://console.cloud.google.com/apis/library/photoslibrary.googleapis.com?project=138255947511');
  console.log('2. 「有効にする」ボタンをクリック');
  console.log('3. 5-10分待機');
  console.log('4. testPhotosAPIConnection() を実行して確認\n');

  console.log('【方法B: 新しいGCPプロジェクトを作成】');
  console.log('1. https://console.cloud.google.com/ にアクセス');
  console.log('2. 新しいプロジェクトを作成');
  console.log('3. Photos Library APIを有効化');
  console.log('4. GASエディタで「プロジェクトの設定」→「GCPプロジェクト番号」に新しいプロジェクト番号を設定');
  console.log('5. testPhotosAPIConnection() を実行して確認\n');

  console.log('【方法C: CLIツールを使用（上級者向け）】');
  console.log('gcloud CLIがインストール済みの場合:');
  console.log('$ gcloud services enable photoslibrary.googleapis.com --project=138255947511');
}

/**
 * 権限の再認証を強制
 */
function forceReauthorization() {
  console.log('=== 権限の再認証 ===\n');

  console.log('以下の手順で権限を再認証します:');
  console.log('1. GASエディタで任意の関数を実行');
  console.log('2. 「承認が必要」ダイアログが表示されたら「権限を確認」をクリック');
  console.log('3. Googleアカウントを選択');
  console.log('4. すべての権限を許可');

  // ダミー関数を実行して権限ダイアログを表示
  try {
    // Photos APIを直接呼び出して権限を要求
    const token = ScriptApp.getOAuthToken();
    UrlFetchApp.fetch('https://photoslibrary.googleapis.com/v1/albums', {
      headers: {
        'Authorization': 'Bearer ' + token
      },
      muteHttpExceptions: true
    });
  } catch (e) {
    console.log('権限の再認証が必要です');
  }
}

/**
 * 完全なセットアップウィザード
 */
function photosAPISetupWizard() {
  console.log('=================================================================================');
  console.log('                    Google Photos API セットアップウィザード                      ');
  console.log('=================================================================================\n');

  let step = 1;

  console.log(`【ステップ ${step++}: 現在の状態確認】`);
  const setupInfo = checkPhotosAPISetup();

  console.log(`\n【ステップ ${step++}: API有効化】`);
  const apiEnabled = testPhotosAPIConnection();

  if (!apiEnabled) {
    console.log('\nPhotos Library APIが無効です。以下の手順で有効化してください:\n');
    console.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
    console.log('📌 最も簡単な方法:');
    console.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
    console.log('\n1️⃣ 以下のリンクをCtrl+クリック（Macはcmd+クリック）で新しいタブで開く:');
    console.log('   https://console.cloud.google.com/apis/library/photoslibrary.googleapis.com');
    console.log('\n2️⃣ 上部のプロジェクトセレクタで正しいプロジェクトが選択されているか確認');
    console.log('\n3️⃣ 「有効にする」ボタンをクリック');
    console.log('\n4️⃣ 5-10分待機');
    console.log('\n5️⃣ 以下のコマンドを実行して確認:');
    console.log('   testPhotosAPIConnection()');
    console.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n');

    return {
      success: false,
      message: 'API有効化が必要です'
    };
  }

  console.log(`\n【ステップ ${step++}: セットアップ完了】`);
  console.log('✅ Google Photos APIが利用可能です！');
  console.log('\n次のコマンドを実行してください:');
  console.log('setupPhotosSync()');

  return {
    success: true,
    message: 'Photos API利用可能'
  };
}

/**
 * API有効化の代替方法
 */
function alternativeAPIEnablement() {
  console.log('=== API有効化の代替方法 ===\n');

  console.log('【オプション1: サービスアカウントを使用】');
  console.log('1. GCPコンソールでサービスアカウントを作成');
  console.log('2. Photos Library APIの権限を付与');
  console.log('3. JSONキーをダウンロード');
  console.log('4. GASにキーを設定\n');

  console.log('【オプション2: OAuth 2.0クライアントIDを使用】');
  console.log('1. GCPコンソールでOAuth 2.0クライアントIDを作成');
  console.log('2. リダイレクトURIを設定');
  console.log('3. クライアントIDとシークレットをGASに設定\n');

  console.log('【オプション3: 手動でAPIリクエスト】');
  console.log('REST APIを直接呼び出す方法もあります。');
  console.log('詳細は setupManualPhotosAPI() を実行してください。');
}

/**
 * 手動でPhotos APIを呼び出す
 */
function setupManualPhotosAPI() {
  const props = PropertiesService.getScriptProperties();

  // 手動API設定を保存
  props.setProperty('USE_MANUAL_PHOTOS_API', 'true');

  console.log('=== 手動Photos API設定 ===\n');
  console.log('Photos APIを手動で呼び出す設定を有効化しました。');
  console.log('\n使用方法:');
  console.log('1. Google Photosで画像を選択');
  console.log('2. 共有リンクを取得');
  console.log('3. processPhotosManually() を実行');

  return {
    success: true,
    message: '手動API設定完了'
  };
}