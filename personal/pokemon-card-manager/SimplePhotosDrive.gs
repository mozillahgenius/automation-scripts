// ==============================
// シンプルなGoogle Photos → Drive バックアップ
// ==============================

/**
 * Google PhotosからDriveへの簡単なバックアップ方法
 */
function simplePhotosDriveSetup() {
  console.log('=================================================================================');
  console.log('           シンプルなGoogle Photos → Drive バックアップ設定                      ');
  console.log('=================================================================================\n');

  console.log('複雑なAPIを使わず、簡単な方法でカード画像を管理します。\n');

  console.log('【方法1: Google Takeoutを使用（最も簡単）】');
  console.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  console.log('1. https://takeout.google.com にアクセス');
  console.log('2. 「Google Photos」のみを選択');
  console.log('3. 特定のアルバム「ポケモンカード」を選択');
  console.log('4. エクスポート先を「Driveに追加」に設定');
  console.log('5. エクスポート実行（定期的に自動実行も可能）\n');

  console.log('【方法2: Google Photos → Drive 手動コピー】');
  console.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  console.log('1. Google Photos（photos.google.com）を開く');
  console.log('2. 「ポケモンカード」アルバムの写真を選択');
  console.log('3. 右上のメニューから「ダウンロード」');
  console.log('4. ダウンロードした写真をDriveの指定フォルダにアップロード\n');

  console.log('【方法3: Google Photos共有リンク → Drive保存】');
  console.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  console.log('1. Google Photosでアルバムを共有リンク化');
  console.log('2. processSharedPhotosLink() を使用して処理\n');

  console.log('【推奨: Drive版を直接使用】');
  console.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  console.log('Photos APIの複雑な認証を避けて、Drive版を使用することを推奨します。\n');

  console.log('setupSimpleDriveWorkflow() を実行\n');

  return {
    recommended: 'setupSimpleDriveWorkflow()',
    alternative1: 'Google Takeout',
    alternative2: '手動ダウンロード＆アップロード',
    alternative3: '共有リンク経由'
  };
}

/**
 * シンプルなDriveワークフローのセットアップ
 */
function setupSimpleDriveWorkflow() {
  console.log('=== シンプルなDriveワークフロー ===\n');

  console.log('【セットアップ内容】');
  console.log('1. Driveに画像アップロード用フォルダを作成');
  console.log('2. 自動処理スクリプトを設定');
  console.log('3. 処理結果の管理\n');

  // フォルダを作成
  let folderId = PropertiesService.getScriptProperties().getProperty('DRIVE_FOLDER_ID');

  if (!folderId) {
    const rootFolder = DriveApp.createFolder('ポケモンカード管理');
    const uploadFolder = rootFolder.createFolder('新規アップロード');
    const processedFolder = rootFolder.createFolder('処理済み');
    const errorFolder = rootFolder.createFolder('エラー');

    folderId = uploadFolder.getId();
    PropertiesService.getScriptProperties().setProperty('DRIVE_FOLDER_ID', uploadFolder.getId());
    PropertiesService.getScriptProperties().setProperty('PROCESSED_FOLDER_ID', processedFolder.getId());
    PropertiesService.getScriptProperties().setProperty('ERROR_FOLDER_ID', errorFolder.getId());

    console.log('✅ フォルダ構成を作成しました');
    console.log(`📁 メインフォルダ: ${rootFolder.getUrl()}`);
    console.log(`📁 アップロード用: ${uploadFolder.getUrl()}`);
    console.log(`📁 処理済み: ${processedFolder.getUrl()}`);
    console.log(`📁 エラー: ${errorFolder.getUrl()}`);
  } else {
    const folder = DriveApp.getFolderById(folderId);
    console.log(`既存のフォルダを使用: ${folder.getUrl()}`);
  }

  console.log('\n【使い方】');
  console.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  console.log('方法A: PCから直接アップロード');
  console.log('────────────────────────────');
  console.log('1. 上記の「アップロード用」フォルダを開く');
  console.log('2. カード画像をドラッグ&ドロップ');
  console.log('3. processDriveImages() を実行\n');

  console.log('方法B: スマホから（Googleドライブアプリ）');
  console.log('────────────────────────────');
  console.log('1. Googleドライブアプリを開く');
  console.log('2. 「ポケモンカード管理/新規アップロード」フォルダに移動');
  console.log('3. ＋ボタンから「アップロード」→「写真や動画」');
  console.log('4. カード画像を選択してアップロード');
  console.log('5. processDriveImages() を実行\n');

  console.log('方法C: iPhoneから（ショートカット経由）');
  console.log('────────────────────────────');
  console.log('setupiPhoneShortcut() を実行して手順を確認\n');

  // トリガーを設定
  setupSimpleTriggers();

  return {
    status: 'セットアップ完了',
    nextStep: 'processDriveImages()',
    uploadFolder: DriveApp.getFolderById(folderId).getUrl()
  };
}

/**
 * シンプルなDrive画像処理
 */
function processDriveImages() {
  console.log('=== Drive画像処理開始 ===\n');

  const folderId = PropertiesService.getScriptProperties().getProperty('DRIVE_FOLDER_ID');
  const processedFolderId = PropertiesService.getScriptProperties().getProperty('PROCESSED_FOLDER_ID');
  const errorFolderId = PropertiesService.getScriptProperties().getProperty('ERROR_FOLDER_ID');

  if (!folderId) {
    console.log('❌ フォルダが設定されていません');
    console.log('setupSimpleDriveWorkflow() を実行してください');
    return;
  }

  const uploadFolder = DriveApp.getFolderById(folderId);
  const processedFolder = DriveApp.getFolderById(processedFolderId);
  const errorFolder = DriveApp.getFolderById(errorFolderId);

  // 画像ファイルを取得
  const imageTypes = [MimeType.JPEG, MimeType.PNG];
  let processedCount = 0;
  let errorCount = 0;

  imageTypes.forEach(mimeType => {
    const files = uploadFolder.getFilesByType(mimeType);

    while (files.hasNext()) {
      const file = files.next();

      try {
        console.log(`処理中: ${file.getName()}`);

        // 画像をAI分析（既存の関数を使用）
        const driveFile = {
          id: file.getId(),
          name: file.getName(),
          url: file.getUrl(),
          blob: file.getBlob(),
          driveFile: file
        };

        const config = getConfig();
        const cardData = analyzeCardWithAI(driveFile, config);

        // カード情報を基にファイル名を更新
        const newFileName = generateCardFileName(cardData);
        file.setName(newFileName);

        // Notionに登録
        const notionPageId = createNotionRecord(cardData, driveFile, config);

        // スプレッドシートに記録
        logCardToSpreadsheet(cardData, notionPageId);

        // 処理済みフォルダに移動
        file.moveTo(processedFolder);

        console.log(`✅ 処理完了: ${newFileName}`);
        processedCount++;

      } catch (error) {
        console.error(`❌ エラー: ${file.getName()}`, error);

        // エラーフォルダに移動
        file.moveTo(errorFolder);
        errorCount++;
      }
    }
  });

  console.log(`\n処理完了: 成功=${processedCount}, エラー=${errorCount}`);

  if (processedCount > 0) {
    console.log(`処理済みフォルダ: ${processedFolder.getUrl()}`);
  }

  if (errorCount > 0) {
    console.log(`エラーフォルダ: ${errorFolder.getUrl()}`);
  }

  return {
    success: processedCount,
    error: errorCount
  };
}

/**
 * カード情報からファイル名を生成
 */
function generateCardFileName(cardData) {
  const date = new Date().toISOString().split('T')[0];
  const game = cardData.game || 'Unknown';
  const name = cardData.name || 'Unknown';
  const number = cardData.number || 'XXX';

  // ファイル名をクリーンアップ
  const cleanName = name.replace(/[\/\\:*?"<>|]/g, '_').substring(0, 30);

  return `${game}_${cleanName}_${number}_${date}.jpg`;
}

/**
 * シンプルなトリガー設定
 */
function setupSimpleTriggers() {
  // 既存のトリガーを削除
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'processDriveImages') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // 1時間ごとに実行
  ScriptApp.newTrigger('processDriveImages')
    .timeBased()
    .everyHours(1)
    .create();

  console.log('✅ 自動処理トリガーを設定しました（1時間ごと）');
}

/**
 * iPhone用ショートカット設定
 */
function setupiPhoneShortcut() {
  console.log('=== iPhone用ショートカット設定 ===\n');

  console.log('iPhoneの「ショートカット」アプリで以下を設定:\n');

  console.log('【ショートカットの作成手順】');
  console.log('1. ショートカットアプリを開く');
  console.log('2. 「＋」で新規ショートカット作成');
  console.log('3. アクションを追加:\n');

  console.log('   アクション1: 写真を選択');
  console.log('   ├─ 複数を選択: オン\n');

  console.log('   アクション2: 写真をJPEGに変換');
  console.log('   ├─ 品質: 高\n');

  console.log('   アクション3: Googleドライブに保存');
  console.log('   ├─ 保存先: ポケモンカード管理/新規アップロード');
  console.log('   └─ 既存を置き換え: オフ\n');

  console.log('4. ショートカット名: 「カード登録」');
  console.log('5. ホーム画面に追加\n');

  console.log('【使い方】');
  console.log('1. カードの写真を撮影');
  console.log('2. ホーム画面の「カード登録」をタップ');
  console.log('3. 写真を選択');
  console.log('4. 自動的にDriveにアップロード');
  console.log('5. 1時間後に自動処理（または手動で processDriveImages() を実行）\n');

  return {
    status: '手動設定が必要',
    app: 'iOSショートカット'
  };
}

/**
 * 共有Photosリンクから画像を処理
 */
function processSharedPhotosLink(sharedLink) {
  console.log('=== 共有リンクから画像処理 ===\n');

  console.log('⚠️ 注意: Google Photos共有リンクから直接画像を取得することはできません。\n');

  console.log('【代替手順】');
  console.log('1. 共有リンクを開く');
  console.log('2. 画像を選択してダウンロード');
  console.log('3. Driveの指定フォルダにアップロード');
  console.log('4. processDriveImages() を実行\n');

  console.log('または、以下の方法を使用:');
  console.log('setupSimpleDriveWorkflow() でDrive版を使用');

  return {
    status: '手動操作が必要',
    alternative: 'setupSimpleDriveWorkflow()'
  };
}