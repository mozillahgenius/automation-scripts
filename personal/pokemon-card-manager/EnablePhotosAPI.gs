// ==============================
// Google Photos API 有効化ヘルパー
// ==============================

/**
 * Google Photos APIを手動で有効化するための手順を表示
 */
function showPhotosAPIEnableInstructions() {
  console.log('=== Google Photos API 有効化手順 ===\n');

  console.log('プロジェクト権限の問題により、自動有効化ができません。');
  console.log('以下の手順で手動で有効化してください：\n');

  console.log('【方法1: 新しいGCPプロジェクトを作成】');
  console.log('1. https://console.cloud.google.com/ にアクセス');
  console.log('2. 新しいプロジェクトを作成');
  console.log('3. Photos Library APIを有効化');
  console.log('4. GASプロジェクトをこの新しいGCPプロジェクトに関連付け\n');

  console.log('【方法2: Drive版を使用（推奨）】');
  console.log('Google Photos APIを使わず、Drive版を使用します：');
  console.log('1. initialDriveSetup() を実行');
  console.log('2. 作成されたフォルダに画像をアップロード');
  console.log('3. processImagesFromDrive() を実行\n');

  console.log('【方法3: 手動でGoogle Photosから画像を取得】');
  console.log('以下の関数を実行してください：');
  console.log('setupManualPhotosSync()\n');

  return {
    success: false,
    message: 'Google Photos APIの有効化が必要です',
    driveAlternative: true
  };
}

/**
 * 手動でGoogle Photosの画像URLを入力して処理
 */
function setupManualPhotosSync() {
  console.log('=== 手動Photos同期セットアップ ===\n');

  console.log('Google Photosから画像を手動でダウンロードして処理します。');
  console.log('\n使い方：');
  console.log('1. Google Photosでアルバムを作成');
  console.log('2. アルバムの共有リンクを取得');
  console.log('3. processPhotosFromShareLink("共有リンク") を実行');

  // Drive版のセットアップも同時に行う
  const driveSetup = initialDriveSetup();

  console.log('\n代替案として、Drive版も設定しました：');
  console.log(`フォルダURL: ${driveSetup.folderUrl}`);

  return {
    success: true,
    driveFolder: driveSetup.folderUrl,
    message: 'Drive版のセットアップが完了しました'
  };
}

/**
 * Google Photos共有リンクから画像を処理（手動）
 * @param {string} shareLink - Google Photosの共有リンク
 */
function processPhotosFromShareLink(shareLink) {
  console.log('この機能は手動でのダウンロードが必要です：');
  console.log('1. 共有リンクから画像をダウンロード');
  console.log('2. Driveフォルダにアップロード');
  console.log('3. processImagesFromDrive() を実行');

  return {
    success: false,
    message: '手動でのダウンロードとアップロードが必要です'
  };
}

/**
 * Drive版への移行を推奨
 */
function recommendDriveVersion() {
  console.log('=== Drive版への移行を推奨 ===\n');

  console.log('Google Photos APIの権限問題を回避するため、');
  console.log('Drive版の使用を強く推奨します。\n');

  console.log('【Drive版の利点】');
  console.log('✅ 権限設定不要');
  console.log('✅ すぐに使用可能');
  console.log('✅ 安定動作');
  console.log('✅ エラーが少ない\n');

  console.log('【セットアップコマンド】');
  console.log('initialDriveSetup()\n');

  console.log('【使用方法】');
  console.log('1. 上記コマンドで表示されるフォルダに画像をアップロード');
  console.log('2. processImagesFromDrive() を実行\n');

  // 実際にDrive版をセットアップ
  const result = initialDriveSetup();

  console.log('✅ Drive版のセットアップが完了しました！');
  console.log(`📁 フォルダURL: ${result.folderUrl}`);
  console.log(`📊 スプレッドシート: ${result.spreadsheetUrl}`);

  return result;
}

/**
 * 現在の最適な処理方法を提案
 */
function getBestProcessingMethod() {
  console.log('=== 現在の最適な処理方法 ===\n');

  // Photos APIの状態を確認
  let photosAPIAvailable = false;
  try {
    const token = ScriptApp.getOAuthToken();
    const response = UrlFetchApp.fetch('https://photoslibrary.googleapis.com/v1/albums', {
      headers: {
        'Authorization': 'Bearer ' + token
      },
      muteHttpExceptions: true
    });

    if (response.getResponseCode() === 200) {
      photosAPIAvailable = true;
      console.log('✅ Google Photos API: 利用可能');
    } else {
      console.log('❌ Google Photos API: 利用不可');
    }
  } catch (error) {
    console.log('❌ Google Photos API: エラー');
  }

  // Drive設定を確認
  const driveFolderId = PropertiesService.getScriptProperties().getProperty('DRIVE_FOLDER_ID');
  const driveAvailable = driveFolderId ? true : false;

  if (driveAvailable) {
    console.log('✅ Drive版: 設定済み');
  } else {
    console.log('⚠️ Drive版: 未設定');
  }

  console.log('\n【推奨される処理方法】');

  if (photosAPIAvailable) {
    console.log('1. Google Photos同期を使用');
    console.log('   実行: manualPhotosSync()');
  } else if (driveAvailable) {
    console.log('1. Drive版を使用（推奨）');
    console.log('   実行: processImagesFromDrive()');
  } else {
    console.log('1. Drive版をセットアップ');
    console.log('   実行: initialDriveSetup()');
    console.log('2. フォルダに画像をアップロード');
    console.log('3. processImagesFromDrive() を実行');
  }

  return {
    photosAPI: photosAPIAvailable,
    driveSetup: driveAvailable,
    recommendation: driveAvailable ? 'Drive版を使用' : 'Drive版をセットアップ'
  };
}