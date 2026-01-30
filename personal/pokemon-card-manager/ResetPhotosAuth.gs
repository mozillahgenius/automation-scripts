// ==============================
// Google Photos API 認証リセット
// ==============================

/**
 * 認証をリセットして再認証を促す
 */
function resetPhotosAuthorization() {
  console.log('=== Google Photos認証リセット ===\n');

  console.log('【手順1: 現在の認証をクリア】');
  console.log('1. https://myaccount.google.com/permissions にアクセス');
  console.log('2. 「ポケモンカード管理」または関連するGASアプリを探す');
  console.log('3. 「アクセスを削除」をクリック\n');

  console.log('【手順2: 新しい認証を実行】');
  console.log('以下のコマンドを実行してください:');
  console.log('requestPhotosPermission()\n');

  return {
    nextStep: 'requestPhotosPermission()を実行'
  };
}

/**
 * Photos APIの権限を明示的に要求
 */
function requestPhotosPermission() {
  console.log('=== Photos API権限のリクエスト ===\n');

  try {
    // Google Photos APIを明示的に呼び出して権限を要求
    const token = ScriptApp.getOAuthToken();

    // まずシンプルなAPIコールでテスト
    const testUrl = 'https://photoslibrary.googleapis.com/v1/mediaItems';

    const response = UrlFetchApp.fetch(testUrl, {
      method: 'GET',
      headers: {
        'Authorization': 'Bearer ' + token,
        'Content-Type': 'application/json'
      },
      muteHttpExceptions: true
    });

    const code = response.getResponseCode();
    console.log('レスポンスコード: ' + code);

    if (code === 403) {
      console.log('\n⚠️ 権限が不足しています。');
      console.log('\n以下の手順を実行してください:');
      console.log('1. このスクリプトを一度保存（Ctrl+S / Cmd+S）');
      console.log('2. forceNewAuthorization() を実行');
      console.log('3. 表示される認証画面ですべての権限を許可');
      console.log('4. 再度 testPhotosAPIConnection() を実行');

      return false;
    } else if (code === 200) {
      console.log('✅ Photos API権限: 正常に設定されています！');
      console.log('\n次のコマンドを実行:');
      console.log('setupPhotosSync()');
      return true;
    } else {
      console.log('予期しないレスポンス: ' + code);
      console.log(response.getContentText());
      return false;
    }

  } catch (error) {
    console.error('エラー:', error);
    console.log('\n認証が必要です。forceNewAuthorization() を実行してください。');
    return false;
  }
}

/**
 * 強制的に新しい認証を要求
 */
function forceNewAuthorization() {
  console.log('=== 新規認証の強制実行 ===\n');

  // ダミー関数を作成して実行権限を要求
  const testFunction = function() {
    // Photos Library APIへのアクセスを試みる
    try {
      const token = ScriptApp.getOAuthToken();

      // アルバム一覧を取得（これにより権限ダイアログが表示される）
      UrlFetchApp.fetch('https://photoslibrary.googleapis.com/v1/albums', {
        headers: {
          'Authorization': 'Bearer ' + token
        },
        muteHttpExceptions: false  // エラーを表示させる
      });

    } catch (e) {
      // エラーは予期されている（権限がない場合）
      console.log('認証ダイアログが表示されます...');
    }
  };

  // 実行
  testFunction();

  console.log('\n認証ダイアログが表示された場合:');
  console.log('1. 「権限を確認」をクリック');
  console.log('2. Googleアカウントを選択');
  console.log('3. 「詳細」をクリック（警告が出た場合）');
  console.log('4. 「安全でないページに移動」をクリック');
  console.log('5. すべての権限にチェックを入れて「許可」');
  console.log('\n認証完了後、testPhotosAPIConnection() を実行してください');

  return {
    status: '認証プロセス開始',
    nextStep: 'testPhotosAPIConnection()'
  };
}

/**
 * スコープを確認して修正方法を提示
 */
function checkAndFixScopes() {
  console.log('=== スコープの確認と修正 ===\n');

  console.log('【現在の設定】');
  console.log('appsscript.jsonに以下のスコープが必要です:\n');
  console.log('"oauthScopes": [');
  console.log('  "https://www.googleapis.com/auth/photoslibrary",');
  console.log('  "https://www.googleapis.com/auth/photoslibrary.readonly",');
  console.log('  "https://www.googleapis.com/auth/photoslibrary.sharing",');
  console.log('  // ... 他のスコープ');
  console.log(']\n');

  console.log('【修正手順】');
  console.log('1. GASエディタで appsscript.json を開く');
  console.log('2. oauthScopes セクションを確認');
  console.log('3. 上記のPhotos関連スコープが含まれているか確認');
  console.log('4. 含まれていない場合は追加');
  console.log('5. ファイルを保存（Ctrl+S / Cmd+S）');
  console.log('6. forceNewAuthorization() を実行\n');

  console.log('【重要】');
  console.log('スコープを変更した後は、必ず再認証が必要です。');
  console.log('forceNewAuthorization() で再認証してください。');

  return {
    requiredScopes: [
      'https://www.googleapis.com/auth/photoslibrary',
      'https://www.googleapis.com/auth/photoslibrary.readonly',
      'https://www.googleapis.com/auth/photoslibrary.sharing'
    ]
  };
}

/**
 * 完全リセットと再設定
 */
function completePhotosReset() {
  console.log('=================================================================================');
  console.log('                    Google Photos API 完全リセット                               ');
  console.log('=================================================================================\n');

  console.log('以下の手順を順番に実行してください:\n');

  console.log('📝 ステップ1: 既存の認証を削除');
  console.log('────────────────────────────────');
  console.log('1. https://myaccount.google.com/permissions にアクセス');
  console.log('2. このGASプロジェクトのアクセスを削除\n');

  console.log('📝 ステップ2: appsscript.jsonを確認');
  console.log('────────────────────────────────');
  console.log('GASエディタで appsscript.json を開き、以下が含まれているか確認:');
  console.log('"https://www.googleapis.com/auth/photoslibrary"');
  console.log('"https://www.googleapis.com/auth/photoslibrary.readonly"\n');

  console.log('📝 ステップ3: プロジェクトを保存');
  console.log('────────────────────────────────');
  console.log('Ctrl+S（Windows）または Cmd+S（Mac）で保存\n');

  console.log('📝 ステップ4: 新規認証を実行');
  console.log('────────────────────────────────');
  console.log('forceNewAuthorization() を実行\n');

  console.log('📝 ステップ5: 接続テスト');
  console.log('────────────────────────────────');
  console.log('testPhotosAPIConnection() を実行\n');

  console.log('📝 ステップ6: セットアップ');
  console.log('────────────────────────────────');
  console.log('setupPhotosSync() を実行\n');

  return {
    step1: 'https://myaccount.google.com/permissions',
    step2: 'appsscript.json確認',
    step3: 'プロジェクト保存',
    step4: 'forceNewAuthorization()',
    step5: 'testPhotosAPIConnection()',
    step6: 'setupPhotosSync()'
  };
}