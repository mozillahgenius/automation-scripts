// ==============================
// API有効化状態の詳細確認
// ==============================

/**
 * Photos APIの有効化状態を詳細に確認
 */
function checkPhotosAPIStatus() {
  console.log('=== Photos API状態の詳細確認 ===\n');

  try {
    const token = ScriptApp.getOAuthToken();
    const response = UrlFetchApp.fetch('https://photoslibrary.googleapis.com/v1/albums', {
      headers: {
        'Authorization': 'Bearer ' + token,
        'Content-Type': 'application/json'
      },
      muteHttpExceptions: true
    });

    const code = response.getResponseCode();
    const content = response.getContentText();

    console.log('レスポンスコード: ' + code);

    if (code === 403) {
      const error = JSON.parse(content);

      if (error.error && error.error.message) {
        console.log('エラーメッセージ: ' + error.error.message);

        // SERVICE_DISABLEDエラーの場合
        if (error.error.message.includes('has not been used in project')) {
          console.log('\n❌ Photos Library APIが無効です\n');

          // プロジェクトIDを抽出
          const match = error.error.message.match(/project (\d+)/);
          if (match) {
            const projectId = match[1];
            console.log('現在のプロジェクトID: ' + projectId);
            console.log('\n【解決方法】');
            console.log('以下のURLでAPIを有効化してください:');
            console.log(`https://console.developers.google.com/apis/api/photoslibrary.googleapis.com/overview?project=${projectId}`);
            console.log('\nまたは:');
            console.log('https://console.cloud.google.com/apis/library/photoslibrary.googleapis.com?project=company-gas');
          }

          // 有効化URLを抽出
          if (error.error.details) {
            error.error.details.forEach(detail => {
              if (detail.metadata && detail.metadata.activationUrl) {
                console.log('\n直接有効化URL:');
                console.log(detail.metadata.activationUrl);
              }
            });
          }

        } else if (error.error.message.includes('insufficient authentication scopes')) {
          console.log('\n❌ 認証スコープが不足しています\n');
          console.log('【解決方法】');
          console.log('1. GASエディタで任意の関数を手動実行');
          console.log('2. 認証ダイアログですべての権限を許可');
          console.log('3. manualAuthSteps() を参照');

        } else {
          console.log('\n❌ その他のエラー');
        }
      }

    } else if (code === 200) {
      console.log('\n✅ Photos API: 有効化されています！');
      console.log('認証も正常です。');
      console.log('\n次のステップ:');
      console.log('setupPhotosSync() を実行');
      return true;

    } else {
      console.log('\n⚠️ 予期しないレスポンス');
      console.log(content);
    }

  } catch (error) {
    console.error('エラー:', error);
  }

  return false;
}

/**
 * company-gasプロジェクトでAPIを有効化する手順
 */
function enablePhotosAPIInCompanyGas() {
  console.log('=================================================================================');
  console.log('                 company-gasプロジェクトでPhotos APIを有効化                      ');
  console.log('=================================================================================\n');

  console.log('【確認事項】\n');

  console.log('1. プロジェクトの確認');
  console.log('────────────────────');
  console.log('現在使用中のプロジェクト: company-gas');
  console.log('確認URL: https://console.cloud.google.com/home/dashboard?project=company-gas\n');

  console.log('2. APIライブラリにアクセス');
  console.log('────────────────────');
  console.log('以下のURLを新しいタブで開く:');
  console.log('https://console.cloud.google.com/apis/library?project=company-gas\n');

  console.log('3. Photos Library APIを検索');
  console.log('────────────────────');
  console.log('検索ボックスに「Photos Library API」と入力\n');

  console.log('4. APIを有効化');
  console.log('────────────────────');
  console.log('「Photos Library API」をクリック');
  console.log('「有効にする」ボタンをクリック\n');

  console.log('5. 有効化の確認');
  console.log('────────────────────');
  console.log('「管理」ボタンが表示されれば有効化完了\n');

  console.log('6. 5-10分待機');
  console.log('────────────────────');
  console.log('APIの有効化が反映されるまで時間がかかります\n');

  console.log('7. 確認テスト');
  console.log('────────────────────');
  console.log('checkPhotosAPIStatus() を実行\n');

  console.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  console.log('直接リンク（Ctrl/Cmd + クリックで開く）:');
  console.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  console.log('\n📌 Photos Library API有効化ページ:');
  console.log('https://console.cloud.google.com/apis/library/photoslibrary.googleapis.com?project=company-gas\n');

  return {
    project: 'company-gas',
    action: 'Photos Library APIを有効化',
    url: 'https://console.cloud.google.com/apis/library/photoslibrary.googleapis.com?project=company-gas'
  };
}

/**
 * APIが有効化されるまで待機してテスト
 */
function waitAndTestPhotosAPI() {
  console.log('=== API有効化待機中 ===\n');

  console.log('Photos Library APIを有効化した場合、反映まで5-10分かかります。\n');

  console.log('【チェックリスト】');
  console.log('□ company-gasプロジェクトを選択した');
  console.log('□ Photos Library APIを検索した');
  console.log('□ 「有効にする」ボタンをクリックした');
  console.log('□ 「管理」ボタンが表示されている');
  console.log('□ 5分以上待った\n');

  console.log('すべて完了したら、以下を実行:');
  console.log('checkPhotosAPIStatus()\n');

  console.log('まだ403エラーが出る場合:');
  console.log('1. さらに5分待つ');
  console.log('2. ブラウザをリフレッシュ');
  console.log('3. GASエディタを再読み込み');
  console.log('4. 再度 checkPhotosAPIStatus() を実行');

  return {
    status: 'waiting',
    nextCheck: 'checkPhotosAPIStatus()'
  };
}