// ==============================
// OAuth スコープ修正
// ==============================

/**
 * スコープ不足を解決する完全ガイド
 */
function fixInsufficientScopes() {
  console.log('=================================================================================');
  console.log('              「insufficient authentication scopes」エラーの解決                  ');
  console.log('=================================================================================\n');

  console.log('このエラーは、Google Photos APIへのアクセス権限が付与されていないことを意味します。\n');

  console.log('【解決手順】\n');

  console.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  console.log('ステップ1: 既存の認証を完全に削除');
  console.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  console.log('1. 以下のURLにアクセス:');
  console.log('   https://myaccount.google.com/permissions\n');
  console.log('2. 検索ボックスに「ポケモン」または「card」と入力');
  console.log('3. このプロジェクトを見つけて「アクセスを削除」をクリック\n');

  console.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  console.log('ステップ2: GASエディタで再認証を強制');
  console.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  console.log('1. GASエディタを開く');
  console.log('2. 上部の関数選択から「forcePhotosReauth」を選択');
  console.log('3. 「実行」ボタン（▶️）をクリック\n');

  console.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  console.log('ステップ3: 認証画面での重要な操作');
  console.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  console.log('認証画面が表示されたら:');
  console.log('1. Googleアカウントを選択');
  console.log('2. 「詳細」をクリック（重要！）');
  console.log('3. 「ポケモンカード管理（安全ではないページ）に移動」をクリック');
  console.log('4. 以下の権限がすべて表示されることを確認:');
  console.log('   ✓ Google Photos のライブラリを表示');
  console.log('   ✓ Google Drive のファイルの表示、編集、作成、削除');
  console.log('   ✓ Google スプレッドシートのすべてのスプレッドシートの参照、編集、作成、削除');
  console.log('5. すべての権限にチェックを入れて「許可」\n');

  console.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  console.log('ステップ4: 再テスト');
  console.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  console.log('testDirectPhotosAPI() を実行\n');

  return {
    step1: 'https://myaccount.google.com/permissions でアクセス削除',
    step2: 'forcePhotosReauth() をGASエディタから実行',
    step3: '認証画面ですべての権限を許可',
    step4: 'testDirectPhotosAPI() でテスト'
  };
}

/**
 * Photos API用の再認証を強制
 */
function forcePhotosReauth() {
  console.log('=== Photos API再認証 ===\n');

  try {
    // まずDriveで基本認証
    const files = DriveApp.getFiles();
    console.log('Drive認証: OK');

    // 次にPhotos APIを呼び出して追加スコープを要求
    const token = ScriptApp.getOAuthToken();

    // このリクエストで認証エラーを発生させる
    const response = UrlFetchApp.fetch('https://photoslibrary.googleapis.com/v1/albums', {
      headers: {
        'Authorization': 'Bearer ' + token
      },
      muteHttpExceptions: false  // エラーを表示
    });

    console.log('Photos API接続: 成功！');
    console.log('testDirectPhotosAPI() を実行してください');

  } catch (error) {
    // エラーは予期されている
    if (error.toString().includes('insufficient')) {
      console.log('認証が必要です。');
      console.log('\n⚠️ 認証画面が表示されない場合:');
      console.log('1. GASエディタの「実行」ボタンから実行してください');
      console.log('2. コンソールからの実行では認証画面が出ません');
    } else {
      console.log('エラー: ' + error.toString());
    }
  }

  return {
    status: '認証プロセス開始',
    nextStep: 'testDirectPhotosAPI()'
  };
}

/**
 * 手動での完全リセット手順
 */
function manualScopeReset() {
  console.log('=== 手動でのスコープリセット ===\n');

  console.log('【方法A: プロジェクトトリガーを使用】');
  console.log('────────────────────────────────');
  console.log('1. GASエディタの左メニューから「トリガー」をクリック');
  console.log('2. 「トリガーを追加」をクリック');
  console.log('3. 実行する関数: testPhotosAuth');
  console.log('4. イベントのソース: 時間主導型');
  console.log('5. 「保存」をクリック');
  console.log('6. 認証画面が表示される\n');

  console.log('【方法B: ウェブアプリとしてデプロイ】');
  console.log('────────────────────────────────');
  console.log('1. GASエディタで「デプロイ」→「新しいデプロイ」');
  console.log('2. 種類: ウェブアプリ');
  console.log('3. 「デプロイ」をクリック');
  console.log('4. 「アクセスを承認」で認証\n');

  console.log('【方法C: 別のGASプロジェクトを作成】');
  console.log('────────────────────────────────');
  console.log('1. https://script.google.com で新規プロジェクト');
  console.log('2. appsscript.jsonでスコープを設定');
  console.log('3. コードをコピー');
  console.log('4. 実行して認証\n');

  return {
    methodA: 'トリガー経由',
    methodB: 'デプロイ経由',
    methodC: '新規プロジェクト'
  };
}

/**
 * トリガー用テスト関数
 */
function testPhotosAuth() {
  const token = ScriptApp.getOAuthToken();

  const response = UrlFetchApp.fetch('https://photoslibrary.googleapis.com/v1/albums', {
    headers: {
      'Authorization': 'Bearer ' + token
    },
    muteHttpExceptions: true
  });

  const code = response.getResponseCode();

  if (code === 200) {
    console.log('✅ Photos API認証成功！');

    // トリガーを削除
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(trigger => {
      if (trigger.getHandlerFunction() === 'testPhotosAuth') {
        ScriptApp.deleteTrigger(trigger);
      }
    });

    return true;
  } else if (code === 403) {
    console.log('❌ まだスコープが不足しています');
    return false;
  }
}

/**
 * 最終手段: スコープを明示的に要求
 */
function requestPhotosScope() {
  console.log('=== スコープを明示的に要求 ===\n');

  // GASエディタでの手動実行を促す
  console.log('【重要】この関数は必ずGASエディタから実行してください\n');

  console.log('手順:');
  console.log('1. GASエディタの上部で「requestPhotosScope」を選択');
  console.log('2. 「実行」ボタン（▶️）をクリック');
  console.log('3. 「承認が必要です」が表示されたら「権限を確認」');
  console.log('4. すべての権限を許可\n');

  // Photos APIアクセスを試みる
  try {
    // DriveとPhotosの両方のスコープを要求
    DriveApp.getRootFolder();

    const token = ScriptApp.getOAuthToken();
    UrlFetchApp.fetch('https://photoslibrary.googleapis.com/v1/mediaItems?pageSize=1', {
      headers: {
        'Authorization': 'Bearer ' + token
      },
      muteHttpExceptions: false
    });

    console.log('✅ 認証成功！');
    return true;

  } catch (e) {
    if (e.toString().includes('承認が必要')) {
      console.log('認証ダイアログが表示されるはずです');
    } else if (e.toString().includes('insufficient')) {
      console.log('スコープ不足: 認証画面ですべての権限を許可してください');
    } else {
      console.log('エラー: ' + e.toString());
    }
    return false;
  }
}

/**
 * スコープ確認
 */
function checkCurrentScopes() {
  console.log('=== 現在のスコープ確認 ===\n');

  console.log('【必要なスコープ】');
  console.log('appsscript.jsonに以下が含まれている必要があります:\n');
  console.log('"oauthScopes": [');
  console.log('  "https://www.googleapis.com/auth/photoslibrary",');
  console.log('  "https://www.googleapis.com/auth/photoslibrary.readonly",');
  console.log('  "https://www.googleapis.com/auth/drive",');
  console.log('  "https://www.googleapis.com/auth/spreadsheets"');
  console.log(']\n');

  console.log('【確認方法】');
  console.log('1. GASエディタで appsscript.json を開く');
  console.log('2. oauthScopes セクションを確認');
  console.log('3. 不足しているスコープを追加');
  console.log('4. ファイルを保存（Ctrl+S）');
  console.log('5. forcePhotosReauth() を実行\n');

  return {
    requiredScope: 'https://www.googleapis.com/auth/photoslibrary',
    action: 'appsscript.jsonを確認して追加'
  };
}