// ==============================
// GASプロジェクトとGCPプロジェクトの関連付け修正
// ==============================

/**
 * 現在のプロジェクト設定を確認して修正方法を提示
 */
function checkAndFixProjectSettings() {
  console.log('=================================================================================');
  console.log('                    プロジェクト設定の確認と修正                                  ');
  console.log('=================================================================================\n');

  // 現在のGASプロジェクト情報
  console.log('【現在のGASプロジェクト情報】');
  console.log('スクリプトID: ' + ScriptApp.getScriptId());
  console.log('エラーメッセージのプロジェクトID: 138255947511');
  console.log('※このプロジェクトではPhotos APIが有効化できていません\n');

  console.log('【解決方法を選択してください】\n');

  console.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  console.log('方法1: 既にAPIを有効化したプロジェクトをGASに関連付ける（推奨）');
  console.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  console.log('\n1. GASエディタで以下の手順を実行:');
  console.log('   a) 左メニューの「プロジェクトの設定」をクリック');
  console.log('   b) 「Google Cloud Platform（GCP）プロジェクト」セクションを探す');
  console.log('   c) 「プロジェクトを変更」をクリック');
  console.log('   d) APIを有効化したプロジェクトのプロジェクト番号を入力');
  console.log('      ※プロジェクト番号は以下で確認できます:');
  console.log('      https://console.cloud.google.com/home/dashboard');
  console.log('   e) 「プロジェクトを設定」をクリック\n');

  console.log('2. プロジェクト番号の確認方法:');
  console.log('   a) https://console.cloud.google.com/ にアクセス');
  console.log('   b) 上部のプロジェクトセレクタで使用中のプロジェクトを選択');
  console.log('   c) ダッシュボードに表示される「プロジェクト番号」をコピー\n');

  console.log('3. 設定後、以下を実行:');
  console.log('   testPhotosAPIConnection()\n');

  console.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  console.log('方法2: 新しいGCPプロジェクトを作成して関連付ける');
  console.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  console.log('\n1. 新しいプロジェクトを作成:');
  console.log('   setupNewGCPProject() を実行\n');

  console.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  console.log('方法3: スタンドアロンスクリプトとして再作成');
  console.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  console.log('\n1. 新しいスタンドアロンGASプロジェクトを作成:');
  console.log('   https://script.google.com/ で新規作成');
  console.log('2. コードをコピー');
  console.log('3. GCPプロジェクトを正しく設定');
  console.log('4. APIを有効化\n');

  return {
    currentProjectIssue: 'プロジェクトID不一致',
    solution: '上記の方法1を推奨'
  };
}

/**
 * 新しいGCPプロジェクトの作成手順
 */
function setupNewGCPProject() {
  console.log('=== 新しいGCPプロジェクトの作成手順 ===\n');

  console.log('【ステップ1: 新しいプロジェクトを作成】');
  console.log('1. 以下のURLにアクセス:');
  console.log('   https://console.cloud.google.com/projectcreate\n');

  console.log('2. プロジェクト情報を入力:');
  console.log('   プロジェクト名: pokemon-card-manager');
  console.log('   プロジェクトID: 自動生成されたものを使用');
  console.log('   場所: 組織なし（個人アカウントの場合）\n');

  console.log('3. 「作成」をクリック\n');

  console.log('【ステップ2: Photos Library APIを有効化】');
  console.log('1. プロジェクト作成後、以下にアクセス:');
  console.log('   https://console.cloud.google.com/apis/library/photoslibrary.googleapis.com\n');

  console.log('2. 「有効にする」をクリック\n');

  console.log('【ステップ3: プロジェクト番号を取得】');
  console.log('1. プロジェクトダッシュボードで「プロジェクト番号」を確認');
  console.log('   https://console.cloud.google.com/home/dashboard\n');

  console.log('2. プロジェクト番号をコピー（例: 123456789012）\n');

  console.log('【ステップ4: GASにプロジェクトを関連付け】');
  console.log('1. GASエディタに戻る');
  console.log('2. 「プロジェクトの設定」→「GCPプロジェクト」');
  console.log('3. プロジェクト番号を入力して「プロジェクトを設定」\n');

  console.log('【ステップ5: 確認】');
  console.log('testPhotosAPIConnection() を実行\n');

  return {
    nextStep: 'GCPコンソールで新しいプロジェクトを作成'
  };
}

/**
 * 既存のプロジェクト番号を確認する方法
 */
function findExistingProjectNumber() {
  console.log('=== 既存プロジェクトの番号を確認 ===\n');

  console.log('Photos APIを有効化したプロジェクトの番号を確認します:\n');

  console.log('1. Google Cloud Consoleにアクセス');
  console.log('   https://console.cloud.google.com/\n');

  console.log('2. 上部のプロジェクトセレクタをクリック\n');

  console.log('3. 「すべて」タブを選択\n');

  console.log('4. Photos APIを有効化したプロジェクトを見つける');
  console.log('   ※「Untitled project」や最近作成したプロジェクト\n');

  console.log('5. プロジェクトをクリックして選択\n');

  console.log('6. ホーム/ダッシュボードで以下を確認:');
  console.log('   - プロジェクト名');
  console.log('   - プロジェクトID');
  console.log('   - プロジェクト番号 ← これをコピー\n');

  console.log('7. GASエディタで:');
  console.log('   a) プロジェクトの設定を開く');
  console.log('   b) GCPプロジェクト番号に貼り付け');
  console.log('   c) 「プロジェクトを設定」をクリック\n');

  console.log('コピーしたプロジェクト番号をメモしておいてください。');

  return {
    instruction: 'GCPコンソールでプロジェクト番号を確認'
  };
}

/**
 * プロジェクト設定後の確認
 */
function verifyProjectSetup() {
  console.log('=== プロジェクト設定の確認 ===\n');

  // API接続テスト
  console.log('【Photos API接続テスト】');
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

    if (code === 200) {
      console.log('✅ 成功！Photos APIに接続できました！');
      console.log('\n次のステップ:');
      console.log('setupPhotosSync() を実行してください');
      return true;
    } else if (code === 403) {
      console.log('❌ まだAPIが有効化されていません');
      const error = JSON.parse(response.getContentText());

      if (error.error && error.error.details && error.error.details[0]) {
        const projectId = error.error.details[0].metadata.consumer;
        console.log('\n現在のプロジェクト: ' + projectId);
        console.log('このプロジェクトでPhotos APIを有効化する必要があります');
      }

      console.log('\n解決方法:');
      console.log('1. checkAndFixProjectSettings() を実行');
      console.log('2. 表示される手順に従ってプロジェクトを設定');
      return false;
    } else {
      console.log('⚠️ 予期しないエラー: ' + code);
      console.log(response.getContentText());
      return false;
    }
  } catch (error) {
    console.error('エラー:', error);
    console.log('\n権限の再認証が必要かもしれません');
    console.log('forceReauthorization() を実行してください');
    return false;
  }
}

/**
 * クイックセットアップガイド
 */
function quickProjectSetup() {
  console.log('=================================================================================');
  console.log('                    クイックセットアップガイド                                    ');
  console.log('=================================================================================\n');

  console.log('📌 最も簡単な解決方法:\n');

  console.log('【オプションA: 既存のAPIが有効なプロジェクトを使用】');
  console.log('────────────────────────────────────────');
  console.log('1. findExistingProjectNumber() を実行');
  console.log('2. 表示される手順でプロジェクト番号を取得');
  console.log('3. GASエディタでプロジェクト番号を設定');
  console.log('4. verifyProjectSetup() で確認\n');

  console.log('【オプションB: 新規プロジェクトを作成】');
  console.log('────────────────────────────────────────');
  console.log('1. setupNewGCPProject() を実行');
  console.log('2. 表示される手順で新しいプロジェクトを作成');
  console.log('3. Photos APIを有効化');
  console.log('4. GASエディタでプロジェクト番号を設定');
  console.log('5. verifyProjectSetup() で確認\n');

  console.log('どちらかの方法を選んで実行してください。');

  return {
    optionA: 'findExistingProjectNumber()',
    optionB: 'setupNewGCPProject()'
  };
}