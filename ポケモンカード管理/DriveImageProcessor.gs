// ==============================
// Google Drive から直接画像を処理する代替版
// ==============================

/**
 * Google Driveの指定フォルダから新着画像を処理する
 * （Google Photos APIを使わない簡易版）
 */
function processImagesFromDrive() {
  const startTime = Date.now();

  try {
    console.log('ドライブ画像処理開始');

    const config = getConfig();
    const processedIds = getProcessedIds();

    // Driveフォルダから画像を取得
    const folder = DriveApp.getFolderById(config.DRIVE_FOLDER_ID);
    const files = folder.getFilesByType(MimeType.JPEG);

    const newImages = [];

    // 新着画像を収集（最大処理数まで）
    while (files.hasNext() && newImages.length < config.MAX_PHOTOS_PER_RUN) {
      const file = files.next();
      const fileId = file.getId();

      // 処理済みチェック
      if (!processedIds.includes(fileId)) {
        newImages.push({
          id: fileId,
          file: file,
          name: file.getName(),
          createdDate: file.getDateCreated(),
          blob: file.getBlob()
        });
      }
    }

    // PNG画像も処理
    const pngFiles = folder.getFilesByType(MimeType.PNG);
    while (pngFiles.hasNext() && newImages.length < config.MAX_PHOTOS_PER_RUN) {
      const file = pngFiles.next();
      const fileId = file.getId();

      if (!processedIds.includes(fileId)) {
        newImages.push({
          id: fileId,
          file: file,
          name: file.getName(),
          createdDate: file.getDateCreated(),
          blob: file.getBlob()
        });
      }
    }

    if (newImages.length === 0) {
      console.log('新着画像なし');
      return;
    }

    console.log(`新着画像: ${newImages.length}枚`);

    // 各画像を処理
    const results = [];
    for (const image of newImages) {
      try {
        console.log(`処理中: ${image.name}`);

        // Driveファイルオブジェクトをそのまま使用
        const driveFile = {
          id: image.file.getId(),
          name: image.file.getName(),
          url: image.file.getUrl(),
          downloadUrl: image.file.getDownloadUrl(),
          viewUrl: `https://drive.google.com/file/d/${image.file.getId()}/view`,
          blob: image.blob,
          driveFile: image.file
        };

        // AI判定
        const cardData = analyzeCardWithAI(driveFile, config);

        // ユニークIDを生成
        cardData.uniqueId = generateUniqueCardId(cardData, image);
        cardData.driveFileId = image.id;

        // 外部API補完（オプション）
        if (config.USE_EXTERNAL_API) {
          enrichCardData(cardData);
        }

        // 重複チェック
        const duplicateCount = countDuplicateCards(cardData, config);
        cardData.duplicateNumber = duplicateCount + 1;

        // AI判定結果を基にファイル名を更新
        const newFileName = renameDriveFile(driveFile, cardData);
        cardData.driveFileName = newFileName;

        // Notionへ登録
        const notionPageId = createNotionRecord(cardData, driveFile, config);

        // スプレッドシートに記録
        logCardToSpreadsheet(cardData, notionPageId);

        // 処理済みとしてマーク
        markAsProcessed(image.id);

        results.push({
          success: true,
          fileId: image.id,
          notionPageId: notionPageId
        });

      } catch (error) {
        console.error(`画像処理エラー: ${image.name}`, error);
        results.push({
          success: false,
          fileId: image.id,
          error: error.toString()
        });

        // エラーを記録
        logError(image, error);
      }
    }

    // 処理結果サマリー
    const successCount = results.filter(r => r.success).length;
    const failureCount = results.filter(r => !r.success).length;

    console.log(`処理完了: 成功=${successCount}, 失敗=${failureCount}`);

    // 処理履歴を記録
    logProcessingHistory(results, startTime);

    // 通知（失敗がある場合）
    if (failureCount > 0) {
      sendNotification(`カード処理: ${successCount}枚成功, ${failureCount}枚失敗`);
    }

  } catch (error) {
    console.error('Drive画像処理エラー:', error);
    sendNotification('Drive画像処理で重大なエラーが発生しました: ' + error.toString());
  }
}

/**
 * 手動アップロード用フォルダを作成
 */
function createUploadFolder() {
  const parentFolderId = PropertiesService.getScriptProperties().getProperty('DRIVE_FOLDER_ID');

  if (!parentFolderId) {
    // ルートフォルダに作成
    const folder = DriveApp.createFolder('ポケモンカード_アップロード');
    PropertiesService.getScriptProperties().setProperty('DRIVE_FOLDER_ID', folder.getId());
    console.log(`アップロードフォルダ作成: ${folder.getUrl()}`);
    return folder;
  }

  try {
    const folder = DriveApp.getFolderById(parentFolderId);
    console.log(`既存フォルダ: ${folder.getUrl()}`);
    return folder;
  } catch (e) {
    // フォルダが見つからない場合は新規作成
    const folder = DriveApp.createFolder('ポケモンカード_アップロード');
    PropertiesService.getScriptProperties().setProperty('DRIVE_FOLDER_ID', folder.getId());
    console.log(`アップロードフォルダ作成: ${folder.getUrl()}`);
    return folder;
  }
}

/**
 * 処理済みフォルダへ移動
 */
function moveToProcessedFolder(file) {
  try {
    const parentFolderId = PropertiesService.getScriptProperties().getProperty('DRIVE_FOLDER_ID');
    const parentFolder = DriveApp.getFolderById(parentFolderId);

    // 処理済みフォルダを取得または作成
    let processedFolder;
    const folders = parentFolder.getFoldersByName('処理済み');

    if (folders.hasNext()) {
      processedFolder = folders.next();
    } else {
      processedFolder = parentFolder.createFolder('処理済み');
    }

    // ファイルを移動
    file.moveTo(processedFolder);
    console.log(`ファイル移動: ${file.getName()} → 処理済みフォルダ`);

  } catch (error) {
    console.error('ファイル移動エラー:', error);
  }
}

/**
 * トリガー設定（Drive版）
 */
function setupDriveTriggers() {
  // 既存のトリガーをすべて削除
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    ScriptApp.deleteTrigger(trigger);
  });

  // Drive画像処理用トリガー（1時間ごと）
  ScriptApp.newTrigger('processImagesFromDrive')
    .timeBased()
    .everyHours(1)
    .create();

  // 価格更新用トリガー（週1回、月曜日の午前9時）
  ScriptApp.newTrigger('updateCardPrices')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(9)
    .create();

  console.log('Driveトリガー設定完了:');
  console.log('- 画像処理: 1時間ごと');
  console.log('- 価格更新: 週1回（月曜9時）');
}

/**
 * 初期セットアップ（Drive版）
 */
function initialDriveSetup() {
  console.log('=== Drive版セットアップ開始 ===');

  // 1. スクリプトプロパティの初期化
  console.log('1. スクリプトプロパティを設定中...');
  setupScriptProperties();

  // 2. アップロードフォルダの作成
  console.log('2. アップロードフォルダを作成中...');
  const folder = createUploadFolder();
  console.log(`   フォルダURL: ${folder.getUrl()}`);

  // 3. スプレッドシートの作成
  console.log('3. 管理用スプレッドシートを作成中...');
  const spreadsheetResult = setupCardManagementSpreadsheet();
  console.log(`   スプレッドシートURL: ${spreadsheetResult.url}`);

  // 4. トリガーの設定
  console.log('4. 自動実行トリガーを設定中...');
  setupDriveTriggers();

  console.log('\n=== セットアップ完了 ===');
  console.log('使い方:');
  console.log(`1. ${folder.getUrl()} に画像をアップロード`);
  console.log('2. processImagesFromDrive()を実行（または1時間待機）');
  console.log('3. スプレッドシートで結果を確認');

  return {
    folderUrl: folder.getUrl(),
    spreadsheetUrl: spreadsheetResult.url,
    message: 'Drive版セットアップが完了しました'
  };
}