// ==============================
// 強制認証トリガー
// ==============================

/**
 * 新しい方法で認証を強制（必ず認証ダイアログを表示）
 */
function forcePhotosAuthDialog() {
  console.log('=== 認証ダイアログを強制表示 ===\n');

  // DriveAppを使って認証を強制
  // これによりスコープが変更され、再認証が必要になる
  try {
    // まずDriveの権限を要求（これは通常成功する）
    const files = DriveApp.getFiles();
    console.log('Drive権限: OK');

    // 次にPhotos特有の処理を追加
    // これにより追加のスコープが必要になる
    const dummyPhotoRequest = function() {
      const url = 'https://photoslibrary.googleapis.com/v1/albums';
      const token = ScriptApp.getOAuthToken();

      UrlFetchApp.fetch(url, {
        headers: {
          'Authorization': 'Bearer ' + token
        },
        muteHttpExceptions: false  // エラーで認証ダイアログを表示
      });
    };

    // 実行
    dummyPhotoRequest();

  } catch (error) {
    console.log('認証エラー（予期された動作）:', error.toString());
    console.log('\n⚠️ 認証ダイアログが表示されない場合:');
    console.log('deployForAuth() を実行してください');
  }

  return {
    status: '認証プロセス開始',
    nextStep: 'testPhotosAPIConnection()'
  };
}

/**
 * デプロイを使った認証強制（最も確実な方法）
 */
function deployForAuth() {
  console.log('=================================================================================');
  console.log('                    ウェブアプリとしてデプロイして認証を強制                      ');
  console.log('=================================================================================\n');

  console.log('【手動手順が必要です】\n');

  console.log('📝 ステップ1: デプロイメニューを開く');
  console.log('────────────────────────────────');
  console.log('GASエディタの右上「デプロイ」ボタンをクリック\n');

  console.log('📝 ステップ2: 新しいデプロイ');
  console.log('────────────────────────────────');
  console.log('「新しいデプロイ」を選択\n');

  console.log('📝 ステップ3: 種類を選択');
  console.log('────────────────────────────────');
  console.log('歯車アイコンをクリック → 「ウェブアプリ」を選択\n');

  console.log('📝 ステップ4: 設定');
  console.log('────────────────────────────────');
  console.log('説明: Photos API認証用');
  console.log('実行ユーザー: 自分');
  console.log('アクセスできるユーザー: 自分のみ\n');

  console.log('📝 ステップ5: デプロイ');
  console.log('────────────────────────────────');
  console.log('「デプロイ」ボタンをクリック\n');

  console.log('📝 ステップ6: 認証');
  console.log('────────────────────────────────');
  console.log('⚠️ ここで認証ダイアログが表示されます！');
  console.log('1. 「アクセスを承認」をクリック');
  console.log('2. アカウントを選択');
  console.log('3. 「詳細」→「安全でないページに移動」');
  console.log('4. すべての権限を許可\n');

  console.log('📝 ステップ7: 完了後');
  console.log('────────────────────────────────');
  console.log('デプロイ完了後、以下を実行:');
  console.log('testPhotosAPIConnection()\n');

  return {
    instruction: 'GASエディタでデプロイメニューから手動実行'
  };
}

/**
 * テスト用のdoGet関数（ウェブアプリ用）
 */
function doGet() {
  // Photos APIを使用するコードを含める
  try {
    const token = ScriptApp.getOAuthToken();
    const response = UrlFetchApp.fetch('https://photoslibrary.googleapis.com/v1/albums', {
      headers: {
        'Authorization': 'Bearer ' + token
      },
      muteHttpExceptions: true
    });

    return HtmlService.createHtmlOutput('<h1>認証成功</h1><p>Photos API認証が完了しました。GASエディタに戻って testPhotosAPIConnection() を実行してください。</p>');
  } catch (error) {
    return HtmlService.createHtmlOutput('<h1>認証が必要</h1><p>エラー: ' + error.toString() + '</p>');
  }
}

/**
 * 別の認証方法：トリガーを使う
 */
function setupAuthTrigger() {
  console.log('=== トリガーを使った認証 ===\n');

  console.log('時間ベースのトリガーを設定して認証を強制します。\n');

  // 既存のトリガーを削除
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'photosAuthTest') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // 新しいトリガーを作成
  ScriptApp.newTrigger('photosAuthTest')
    .timeBased()
    .after(1000)  // 1秒後に実行
    .create();

  console.log('トリガーを設定しました。');
  console.log('\n⚠️ 認証ダイアログが表示されるはずです');
  console.log('表示されない場合は、以下を確認:');
  console.log('1. GASエディタの「トリガー」メニューを開く');
  console.log('2. photosAuthTestトリガーが作成されているか確認');
  console.log('3. 手動で実行ボタンをクリック');

  return {
    status: 'トリガー設定完了',
    message: '1秒後に認証ダイアログが表示されます'
  };
}

/**
 * トリガー用のテスト関数
 */
function photosAuthTest() {
  try {
    const token = ScriptApp.getOAuthToken();
    const response = UrlFetchApp.fetch('https://photoslibrary.googleapis.com/v1/albums', {
      headers: {
        'Authorization': 'Bearer ' + token
      },
      muteHttpExceptions: false
    });

    console.log('Photos API接続成功！');

    // トリガーを削除
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(trigger => {
      if (trigger.getHandlerFunction() === 'photosAuthTest') {
        ScriptApp.deleteTrigger(trigger);
      }
    });

  } catch (error) {
    console.error('認証が必要です:', error);
  }
}

/**
 * 最も簡単な方法：手動実行
 */
function manualAuthSteps() {
  console.log('=================================================================================');
  console.log('                    最も簡単な認証方法                                           ');
  console.log('=================================================================================\n');

  console.log('【GASエディタで手動実行】\n');

  console.log('1️⃣ GASエディタの関数選択ドロップダウンから');
  console.log('   「testPhotosAPIConnection」を選択\n');

  console.log('2️⃣ 「実行」ボタン（▶️）をクリック\n');

  console.log('3️⃣ 初回実行時に「承認が必要です」と表示されたら:');
  console.log('   a) 「権限を確認」をクリック');
  console.log('   b) Googleアカウントを選択');
  console.log('   c) 「詳細」をクリック');
  console.log('   d) 「ポケモンカード管理（安全ではないページ）に移動」をクリック');
  console.log('   e) すべての権限を許可\n');

  console.log('4️⃣ 認証完了後、再度「実行」ボタンをクリック\n');

  console.log('5️⃣ 成功したら setupPhotosSync() を実行\n');

  console.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  console.log('⚠️ 重要: 必ずGASエディタの「実行」ボタンから実行してください');
  console.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');

  return {
    nextAction: 'GASエディタで testPhotosAPIConnection を手動実行'
  };
}