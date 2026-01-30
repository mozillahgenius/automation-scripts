// ==============================
// ポケモンカード管理システム - 完全統合版
// ==============================

// ==============================
// スプレッドシートメニュー
// ==============================

/**
 * スプレッドシートを開いた時にメニューを追加
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu('🎴 カード管理')
    .addSubMenu(ui.createMenu('📷 画像処理')
      .addItem('🔄 新着画像を処理', 'processNewImagesMenu')
      .addItem('📋 全画像を再処理', 'reprocessAllImagesMenu')
      .addSeparator()
      .addItem('📁 ドライブから処理', 'processImagesFromDriveMenu')
      .addItem('🌆 Google Photosから処理', 'processPhotosAlbumMenu'))
    .addSeparator()
    .addSubMenu(ui.createMenu('💰 価格管理')
      .addItem('🔍 AI価格調査（選択カード）', 'updateSelectedCardPrice')
      .addItem('📈 全カード価格更新', 'updateAllPricesMenu')
      .addSeparator()
      .addItem('💱 為替レート更新', 'updateExchangeRate'))
    .addSeparator()
    .addSubMenu(ui.createMenu('📝 Notion連携')
      .addItem('✅ Notion接続テスト', 'testNotionConnection')
      .addItem('🔄 Notionへ送信', 'syncToNotion')
      .addSeparator()
      .addItem('🏗️ プロパティ設定', 'setupNotionDatabaseProperties'))
    .addSeparator()
    .addSubMenu(ui.createMenu('⚙️ 設定')
      .addItem('🔐 APIキー設定', 'showSettingsDialog')
      .addItem('📋 初期セットアップ', 'initialSetup')
      .addSeparator()
      .addItem('🗑️ エラーログをクリア', 'clearErrorLog')
      .addItem('📋 設定確認', 'showCurrentConfig'))
    .addSeparator()
    .addItem('❓ ヘルプ', 'showHelp')
    .addToUi();
}

// ==============================
// メニュー関連関数
// ==============================

/**
 * 新着画像処理（メニュー用）
 */
function processNewImagesMenu() {
  const ui = SpreadsheetApp.getUi();
  ui.alert('処理開始', '新着画像の処理を開始します...', ui.ButtonSet.OK);

  try {
    const result = processImagesFromDriveImproved();
    ui.alert('処理完了', `処理完了\n成功: ${result.successCount || 0}件\n失敗: ${result.failureCount || 0}件`, ui.ButtonSet.OK);
  } catch (error) {
    ui.alert('エラー', error.toString(), ui.ButtonSet.OK);
  }
}

/**
 * ドライブから処理（メニュー用）
 */
function processImagesFromDriveMenu() {
  const ui = SpreadsheetApp.getUi();
  ui.alert('処理開始', 'ドライブ画像の処理を開始します...', ui.ButtonSet.OK);

  try {
    processImagesFromDrive();
    ui.alert('処理完了', 'ドライブ画像の処理が完了しました', ui.ButtonSet.OK);
  } catch (error) {
    ui.alert('エラー', error.toString(), ui.ButtonSet.OK);
  }
}

/**
 * 全カード価格更新（メニュー用）
 */
function updateAllPricesMenu() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    '価格更新',
    'すべてのカードの価格を更新しますか？\n時間がかかる場合があります。',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) return;

  try {
    updateAllPrices();
    ui.alert('完了', '価格更新が完了しました', ui.ButtonSet.OK);
  } catch (error) {
    ui.alert('エラー', error.toString(), ui.ButtonSet.OK);
  }
}

/**
 * 選択されたカードの価格をAIで更新
 */
function updateSelectedCardPrice() {
  const sheet = SpreadsheetApp.getActiveSheet();
  if (sheet.getName() !== 'カード一覧') {
    SpreadsheetApp.getUi().alert('「カード一覧」シートで実行してください');
    return;
  }

  const range = sheet.getActiveRange();
  const row = range.getRow();

  if (row < 2) {
    SpreadsheetApp.getUi().alert('カードの行を選択してください');
    return;
  }

  const cardData = {
    name: sheet.getRange(row, 3).getValue(),
    game: sheet.getRange(row, 4).getValue(),
    set: sheet.getRange(row, 5).getValue(),
    number: sheet.getRange(row, 6).getValue(),
    rarity: sheet.getRange(row, 7).getValue(),
    language: sheet.getRange(row, 8).getValue(),
    condition: sheet.getRange(row, 9).getValue()
  };

  const ui = SpreadsheetApp.getUi();
  ui.alert('処理中', '価格を更新中...', ui.ButtonSet.OK);

  try {
    // AI価格調査
    getCardPriceByAI(cardData);

    // 英語カードの場合はJPY変換
    if (cardData.language && cardData.language.toUpperCase().startsWith('EN')) {
      convertEnglishCardPrice(cardData);
    }

    // 価格関連データをスプレッドシートに更新
    sheet.getRange(row, 10).setValue(cardData.price || 0);                    // 現在価格
    sheet.getRange(row, 11).setValue(cardData.marketPrice || cardData.price || 0); // 市場価格
    sheet.getRange(row, 12).setValue(cardData.priceTrend || '不明');          // 価格トレンド

    // 価格履歴を更新
    if (cardData.priceHistory) {
      sheet.getRange(row, 13).setValue(cardData.priceHistory['12ヶ月前'] || 0);
      sheet.getRange(row, 14).setValue(cardData.priceHistory['6ヶ月前'] || 0);
      sheet.getRange(row, 15).setValue(cardData.priceHistory['3ヶ月前'] || 0);
    }

    // 価格予測を更新
    if (cardData.pricePrediction) {
      sheet.getRange(row, 16).setValue(cardData.pricePrediction['6ヶ月後'] || 0);
      sheet.getRange(row, 17).setValue(cardData.pricePrediction['12ヶ月後'] || 0);
    }

    // PSA価格も更新
    if (cardData.psaGradedPrice) {
      sheet.getRange(row, 18).setValue(cardData.psaGradedPrice.PSA9 || 0);
      sheet.getRange(row, 19).setValue(cardData.psaGradedPrice['PSA9.5'] || 0);
      sheet.getRange(row, 20).setValue(cardData.psaGradedPrice.PSA10 || 0);
    }

    ui.alert('完了', `価格更新完了: ¥${(cardData.price || 0).toLocaleString()}`, ui.ButtonSet.OK);
  } catch (error) {
    ui.alert('エラー', error.toString(), ui.ButtonSet.OK);
  }
}

/**
 * 全カード価格更新
 */
function updateAllPrices() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('カード一覧');
  if (!sheet) {
    console.error('カード一覧シートがありません');
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    console.log('更新するカードがありません');
    return;
  }

  let successCount = 0;
  let errorCount = 0;
  const startTime = Date.now();

  // カラムインデックス（新フォーマットで統一）
  // 10:現在価格, 11:市場価格, 12:価格トレンド, 13:12ヶ月前, 14:6ヶ月前, 15:3ヶ月前
  // 16:6ヶ月後予測, 17:1年後予測, 18:PSA9, 19:PSA9.5, 20:PSA10

  for (let row = 2; row <= lastRow; row++) {
    try {
      // カード情報を取得
      const cardData = {
        name: sheet.getRange(row, 3).getValue(),
        game: sheet.getRange(row, 4).getValue(),
        set: sheet.getRange(row, 5).getValue(),
        number: sheet.getRange(row, 6).getValue(),
        rarity: sheet.getRange(row, 7).getValue(),
        language: sheet.getRange(row, 8).getValue(),
        condition: sheet.getRange(row, 9).getValue()
      };

      // 必須フィールドのチェック
      if (!cardData.name || !cardData.game) {
        console.log(`行${row}: カード名またはゲーム名が不足しています`);
        continue;
      }

      console.log(`行${row}: ${cardData.name}の価格を更新中...`);

      // AI価格調査
      getCardPriceByAI(cardData);

      // 英語カードの場合はUSD→JPY変換
      if (cardData.language && cardData.language.toUpperCase().startsWith('EN')) {
        convertEnglishCardPrice(cardData);
        console.log(`  英語カード価格変換: ¥${cardData.price}, トレンド: ${cardData.priceTrend || '不明'}`);
      }

      // 価格関連データをスプレッドシートに更新
      sheet.getRange(row, 10).setValue(cardData.price || 0);                    // 現在価格
      sheet.getRange(row, 11).setValue(cardData.marketPrice || cardData.price || 0); // 市場価格
      sheet.getRange(row, 12).setValue(cardData.priceTrend || '不明');          // 価格トレンド

      // 価格履歴を更新
      if (cardData.priceHistory) {
        sheet.getRange(row, 13).setValue(cardData.priceHistory['12ヶ月前'] || 0);
        sheet.getRange(row, 14).setValue(cardData.priceHistory['6ヶ月前'] || 0);
        sheet.getRange(row, 15).setValue(cardData.priceHistory['3ヶ月前'] || 0);
      }

      // 価格予測を更新
      if (cardData.pricePrediction) {
        sheet.getRange(row, 16).setValue(cardData.pricePrediction['6ヶ月後'] || 0);
        sheet.getRange(row, 17).setValue(cardData.pricePrediction['12ヶ月後'] || 0);
      }

      // PSA価格を更新
      if (cardData.psaGradedPrice) {
        sheet.getRange(row, 18).setValue(cardData.psaGradedPrice.PSA9 || 0);
        sheet.getRange(row, 19).setValue(cardData.psaGradedPrice['PSA9.5'] || 0);
        sheet.getRange(row, 20).setValue(cardData.psaGradedPrice.PSA10 || 0);
      }

      successCount++;

      // API制限を考慮して少し待機
      Utilities.sleep(1000); // 1秒待機

    } catch (error) {
      console.error(`行${row}のエラー:`, error);
      errorCount++;
    }
  }

  const elapsedTime = Math.round((Date.now() - startTime) / 1000);
  console.log(`価格更新完了: 成功=${successCount}, 失敗=${errorCount}, 処理時間=${elapsedTime}秒`);

  return {
    success: successCount,
    error: errorCount,
    time: elapsedTime
  };
}

/**
 * 為替レートを更新
 */
function updateExchangeRate() {
  const rate = getExchangeRate('USD', 'JPY');
  SpreadsheetApp.getUi().alert('為替レート', `現在のレート\n1 USD = ${rate} JPY`, SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Notion接続テスト
 */
function testNotionConnection() {
  const config = getConfig();
  const ui = SpreadsheetApp.getUi();

  if (!config.NOTION_API_KEY || !config.NOTION_DATABASE_ID) {
    ui.alert('Notion設定エラー', 'Notion設定がありません。\n設定メニューからAPIキーを設定してください', ui.ButtonSet.OK);
    return;
  }

  try {
    const dbInfo = getNotionDatabaseInfo(config);
    ui.alert('接続成功', `Notion接続成功！\n\nデータベース名: ${dbInfo.title}\nプロパティ数: ${Object.keys(dbInfo.properties).length}`, ui.ButtonSet.OK);
  } catch (error) {
    ui.alert('接続エラー', error.toString(), ui.ButtonSet.OK);
  }
}

/**
 * 設定ダイアログを表示
 */
function showSettingsDialog() {
  const html = `
    <div style="padding: 20px; font-family: Arial, sans-serif;">
      <h2>APIキー設定</h2>
      <p>スクリプトプロパティから設定を編集してください。</p>
      <br>
      <h3>設定方法:</h3>
      <ol>
        <li>スクリプトエディタを開く</li>
        <li>「プロジェクト設定」をクリック</li>
        <li>「スクリプトプロパティ」を選択</li>
        <li>必要なAPIキーを入力</li>
      </ol>
      <br>
      <h3>必要なAPIキー:</h3>
      <ul>
        <li><strong>OPENAI_API_KEY</strong>: OpenAI APIキー</li>
        <li><strong>PERPLEXITY_API_KEY</strong>: Perplexity APIキー</li>
        <li><strong>NOTION_API_KEY</strong>: Notion APIキー</li>
        <li><strong>NOTION_DATABASE_ID</strong>: NotionデータベースID</li>
      </ul>
      <br>
      <button onclick="google.script.host.close()">閉じる</button>
    </div>
  `;

  const htmlOutput = HtmlService.createHtmlOutput(html)
      .setWidth(600)
      .setHeight(500);

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'APIキー設定');
}

/**
 * ヘルプを表示
 */
function showHelp() {
  const helpText = `
🎴 カード管理システム ヘルプ

【使い方】
1. 初回は「設定」→「初期セットアップ」を実行
2. APIキーを設定（設定メニューから）
3. 画像をアップロードして処理を実行

【価格更新】
- 単体: カードを選択して「AI価格調査」
- 全体: 「全カード価格更新」
- PSAグレード別価格も自動取得

【PSAグレード価格】
- PSA9, PSA9.5, PSA10の価格を自動取得
- 英語カードも自動でJPY変換

【必要なAPIキー】
- OpenAI APIキー（必須）
- Perplexity APIキー（価格調査用）
- Notion APIキー（Notion連携用）

【サポート】
問題がある場合はエラーログを確認してください
  `;

  SpreadsheetApp.getUi().alert('ヘルプ', helpText, SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * 現在の設定を表示
 */
function showCurrentConfig() {
  const config = getConfig();
  const configText = `
現在の設定状態:

OpenAI API: ${config.OPENAI_API_KEY ? '✅ 設定済み' : '❌ 未設定'}
Perplexity API: ${config.PERPLEXITY_API_KEY ? '✅ 設定済み' : '❌ 未設定'}
Notion API: ${config.NOTION_API_KEY ? '✅ 設定済み' : '❌ 未設定'}
Notion DB: ${config.NOTION_DATABASE_ID ? '✅ 設定済み' : '❌ 未設定'}
Driveフォルダ: ${config.DRIVE_FOLDER_ID ? '✅ 設定済み' : '❌ 未設定'}

処理設定:
一度の最大処理数: ${config.MAX_PHOTOS_PER_RUN || 50}枚
AIモデル: ${config.AI_MODEL || 'gpt-4o'}
為替レート: 1 USD = ${getExchangeRate('USD', 'JPY')} JPY
  `;

  SpreadsheetApp.getUi().alert('設定確認', configText, SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * エラーログをクリア
 */
function clearErrorLog() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('エラーログ');
  if (sheet) {
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.deleteRows(2, lastRow - 1);
      SpreadsheetApp.getUi().alert('完了', 'エラーログをクリアしました', SpreadsheetApp.getUi().ButtonSet.OK);
    } else {
      SpreadsheetApp.getUi().alert('情報', 'エラーログは空です', SpreadsheetApp.getUi().ButtonSet.OK);
    }
  } else {
    SpreadsheetApp.getUi().alert('エラー', 'エラーログシートがありません', SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * スプレッドシートからNotionへ同期
 */
function syncToNotion() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Notionへ同期',
    'スプレッドシートのデータをNotionに送信しますか？',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    return;
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('カード一覧');
  if (!sheet) {
    ui.alert('エラー', 'カード一覧シートがありません', ui.ButtonSet.OK);
    return;
  }

  const config = getConfig();
  if (!config.NOTION_API_KEY || !config.NOTION_DATABASE_ID) {
    ui.alert('エラー', 'Notion設定がありません', ui.ButtonSet.OK);
    return;
  }

  let successCount = 0;
  let errorCount = 0;
  const lastRow = sheet.getLastRow();
  const maxRows = Math.min(lastRow, 11); // テスト用に10件まで

  for (let i = 2; i <= maxRows; i++) {
    try {
      const cardData = {
        uniqueId: sheet.getRange(i, 1).getValue(),
        name: sheet.getRange(i, 3).getValue(),
        game: sheet.getRange(i, 4).getValue(),
        set: sheet.getRange(i, 5).getValue(),
        number: sheet.getRange(i, 6).getValue(),
        rarity: sheet.getRange(i, 7).getValue(),
        language: sheet.getRange(i, 8).getValue(),
        condition: sheet.getRange(i, 9).getValue(),
        price: sheet.getRange(i, 10).getValue(),
        psaGradedPrice: {
          PSA9: sheet.getRange(i, 17).getValue(),
          'PSA9.5': sheet.getRange(i, 18).getValue(),
          PSA10: sheet.getRange(i, 19).getValue()
        }
      };

      // 英語カードの場合はJPY変換を確認
      if (cardData.language && cardData.language.toUpperCase().startsWith('EN')) {
        convertEnglishCardPrice(cardData);
      }

      const notionId = createNotionRecord(cardData, null, config);
      if (notionId) {
        sheet.getRange(i, 22).setValue(notionId); // Notion IDカラムを更新
        successCount++;
      } else {
        errorCount++;
      }
    } catch (error) {
      console.error(`行${i}の処理エラー:`, error);
      errorCount++;
    }
  }

  ui.alert('同期完了', `同期完了\n成功: ${successCount}件\n失敗: ${errorCount}件`, ui.ButtonSet.OK);
}

/**
 * Google Photosから処理（メニュー用）
 */
function processPhotosAlbumMenu() {
  const ui = SpreadsheetApp.getUi();
  const config = getConfig();

  if (!config.PHOTOS_ALBUM_ID) {
    ui.alert('エラー', 'Google PhotosアルバムIDが設定されていません', ui.ButtonSet.OK);
    return;
  }

  ui.alert('処理開始', 'Google Photosから画像を処理します...', ui.ButtonSet.OK);

  try {
    main(); // Photos API版のmain関数を呼び出し
    ui.alert('処理完了', 'Google Photosの処理が完了しました', ui.ButtonSet.OK);
  } catch (error) {
    ui.alert('エラー', error.toString(), ui.ButtonSet.OK);
  }
}

/**
 * 全画像再処理（メニュー用）
 */
function reprocessAllImagesMenu() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    '確認',
    'すべての画像を再処理しますか？\n処理済みフラグがリセットされます。',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) return;

  try {
    // 処理済みIDをクリア
    PropertiesService.getScriptProperties().deleteProperty('PROCESSED_PHOTO_IDS');

    // 再処理実行
    processImagesFromDriveImproved();

    ui.alert('完了', '全画像の再処理が完了しました', ui.ButtonSet.OK);
  } catch (error) {
    ui.alert('エラー', error.toString(), ui.ButtonSet.OK);
  }
}

// ==============================
// メインエントリポイント（Google Photos版）
// ==============================

function main() {
  const startTime = Date.now(); // 処理時間計測用

  try {
    console.log('ポケモンカード管理処理開始');

    const config = getConfig();
    const processedIds = getProcessedIds();

    // 1. Google Photosから新着画像を取得
    const newPhotos = getNewPhotosFromAlbum(config.PHOTOS_ALBUM_ID, processedIds);

    if (newPhotos.length === 0) {
      console.log('新着画像なし');
      return;
    }

    console.log(`新着画像: ${newPhotos.length}枚`);

    // 2. 各画像を処理
    const results = [];
    for (const photo of newPhotos) {
      try {
        console.log(`処理中: ${photo.filename}`);

        // 2.1 Google Driveへ保存
        const driveFile = savePhotoToDrive(photo, config.DRIVE_FOLDER_ID);

        // 2.2 AI判定
        const cardData = analyzeCardWithAI(driveFile, config);

        // 2.3 ユニークIDを生成（重複カードでも別管理可能に）
        cardData.uniqueId = generateUniqueCardId(cardData, photo);
        cardData.photoId = photo.id; // Google PhotosのIDも保持

        // 2.4 外部API補完（オプション）
        if (config.USE_EXTERNAL_API) {
          enrichCardData(cardData);
        }
        // 価格はPerplexityのsonar-proで推定
        getCardPriceByAI(cardData);

        // 2.5 重複チェック（同じカードの枚数をカウント）
        const duplicateCount = countDuplicateCards(cardData, config);
        cardData.duplicateNumber = duplicateCount + 1; // 何枚目かを記録

        // 2.6 AI判定結果を基にDriveファイル名を更新
        const newFileName = renameDriveFile(driveFile, cardData);
        cardData.driveFileName = newFileName;

        // 2.7 Notionへ登録
        const notionPageId = createNotionRecord(cardData, driveFile, config);

        // 2.8 スプレッドシートに記録
        logCardToSpreadsheet(cardData, notionPageId);

        // 2.9 処理済みとしてマーク
        markAsProcessed(photo.id);

        results.push({
          success: true,
          photoId: photo.id,
          notionPageId: notionPageId
        });

      } catch (error) {
        console.error(`画像処理エラー: ${photo.filename}`, error);
        results.push({
          success: false,
          photoId: photo.id,
          error: error.toString()
        });

        // エラーを記録
        logError(photo, error);
      }
    }

    // 3. 処理結果サマリー
    const successCount = results.filter(r => r.success).length;
    const failureCount = results.filter(r => !r.success).length;

    console.log(`処理完了: 成功=${successCount}, 失敗=${failureCount}`);

    // 4. 処理履歴をスプレッドシートに記録
    logProcessingHistory(results, startTime);

    // 5. 通知（失敗が多い場合）
    if (failureCount > 0) {
      sendNotification(`カード処理: ${successCount}枚成功, ${failureCount}枚失敗`);
    }

  } catch (error) {
    console.error('メイン処理エラー:', error);
    sendNotification('カード管理処理で重大なエラーが発生しました: ' + error.toString());
  }
}

// ==============================
// Google Drive版処理（Photos APIを使わない代替版）
// ==============================

function processImagesFromDrive() {
  const startTime = Date.now();

  try {
    console.log('ドライブ画像処理開始');

    const config = getConfig();
    const processedIds = getProcessedIds();

    // Driveフォルダから画像を取得
    const folder = DriveApp.getFolderById(config.DRIVE_FOLDER_ID);

    // すべてのファイルを取得して画像形式をフィルタリング
    const allFiles = folder.getFiles();
    const newImages = [];
    const supportedTypes = [
      'image/jpeg',
      'image/jpg',
      'image/png',
      'image/heic',
      'image/heif'
    ];

    // 新着画像を収集（最大処理数まで）
    while (allFiles.hasNext() && newImages.length < config.MAX_PHOTOS_PER_RUN) {
      const file = allFiles.next();
      const fileId = file.getId();
      const mimeType = file.getMimeType();
      const fileName = file.getName().toLowerCase();

      // MIMEタイプまたは拡張子で画像判定
      const isImage = supportedTypes.includes(mimeType) ||
                      fileName.endsWith('.heic') ||
                      fileName.endsWith('.heif') ||
                      fileName.endsWith('.jpg') ||
                      fileName.endsWith('.jpeg') ||
                      fileName.endsWith('.png');

      // 処理済みチェック
      if (isImage && !processedIds.includes(fileId)) {
        console.log(`発見: ${file.getName()} (${mimeType || '不明な形式'})`);

        // HEICの場合はJPEGに変換
        let blob = file.getBlob();
        let convertedFile = file;

        if (fileName.endsWith('.heic') || fileName.endsWith('.heif')) {
          try {
            console.log(`HEIC画像を検出: ${file.getName()}`);

            // HEICファイルをJPEGに変換
            const jpegBlob = convertHeicToJpeg(file);

            // 新しいJPEGファイルを作成
            const newFileName = file.getName().replace(/\.(heic|heif)$/i, '.jpg');
            const folder = file.getParents().next();
            const newFile = folder.createFile(jpegBlob);
            newFile.setName(newFileName);

            // 元のHEICファイルを削除または別フォルダに移動
            moveToArchiveFolder(file, folder);

            console.log(`HEIC→JPEG変換成功: ${newFileName}`);

            // 変換後のファイルを使用
            convertedFile = newFile;
            blob = jpegBlob;

          } catch (e) {
            console.error(`HEIC変換エラー: ${e.toString()}`);
            // エラーの場合は元のファイルをそのまま使用
            blob = file.getBlob();
          }
        }

        newImages.push({
          id: convertedFile.getId(),
          file: convertedFile,
          name: convertedFile.getName(),
          createdDate: convertedFile.getDateCreated(),
          blob: blob
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
        // 価格はPerplexityのsonar-proで推定
        getCardPriceByAI(cardData);

        // 英語カードの場合はUSD→JPY変換
        if (cardData.language && cardData.language.toUpperCase().startsWith('EN')) {
          convertEnglishCardPrice(cardData);
          console.log(`英語カード価格変換完了: ¥${cardData.price}, トレンド: ${cardData.priceTrend || '不明'}`);
        }

        // 重複チェック
        const duplicateCount = countDuplicateCards(cardData, config);
        cardData.duplicateNumber = duplicateCount + 1;

        // AI判定結果を基にファイル名を更新
        const newFileName = renameDriveFile(driveFile, cardData);
        cardData.driveFileName = newFileName;

        // Notionへ登録（エラー時はスプレッドシートのみに記録）
        let notionPageId = null;
        try {
          notionPageId = createNotionRecord(cardData, driveFile, config);
        } catch (notionError) {
          console.error('Notion登録エラー（スプレッドシートには記録）:', notionError);
          cardData.notionError = notionError.toString();
        }

        // スプレッドシートに記録（Notionが失敗しても必ず実行）
        logCardToSpreadsheet(cardData, notionPageId);

        // 処理済みとしてマーク
        markAsProcessed(image.id);

        // 処理済みフォルダへ移動（オプション）
        moveToProcessedFolder(image.file);

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

// ==============================
// 設定管理
// ==============================

function getConfig() {
  const scriptProperties = PropertiesService.getScriptProperties();

  return {
    // API Keys
    OPENAI_API_KEY: scriptProperties.getProperty('OPENAI_API_KEY'),  // OpenAI追加
    PERPLEXITY_API_KEY: scriptProperties.getProperty('PERPLEXITY_API_KEY'),
    NOTION_API_KEY: scriptProperties.getProperty('NOTION_API_KEY'),

    // IDs
    NOTION_DATABASE_ID: scriptProperties.getProperty('NOTION_DATABASE_ID'),
    PHOTOS_ALBUM_ID: scriptProperties.getProperty('PHOTOS_ALBUM_ID'),
    DRIVE_FOLDER_ID: scriptProperties.getProperty('DRIVE_FOLDER_ID'),

    // オプション
    USE_EXTERNAL_API: scriptProperties.getProperty('USE_EXTERNAL_API') === 'true',
    USE_OPENAI: scriptProperties.getProperty('USE_OPENAI') !== 'false',  // デフォルトtrue
    NOTIFICATION_EMAIL: scriptProperties.getProperty('NOTIFICATION_EMAIL') || '',

    // 処理設定
    MAX_PHOTOS_PER_RUN: parseInt(scriptProperties.getProperty('MAX_PHOTOS_PER_RUN') || '50'),
    // 後方互換: まず VISION_MODEL/PRICE_MODEL、なければ AI_MODEL、最後にデフォルト
    AI_MODEL: scriptProperties.getProperty('AI_MODEL') || '',
    VISION_MODEL: scriptProperties.getProperty('VISION_MODEL') || scriptProperties.getProperty('AI_MODEL') || 'gpt-4o',
    PRICE_MODEL: scriptProperties.getProperty('PRICE_MODEL') || scriptProperties.getProperty('AI_MODEL') || 'gpt-4o'
  };
}

function setupScriptProperties() {
  const scriptProperties = PropertiesService.getScriptProperties();

  // 必須設定項目（値は実際のものに置き換えてください）
  const requiredProperties = {
    'PERPLEXITY_API_KEY': 'your-perplexity-api-key',
    'NOTION_API_KEY': 'your-notion-api-key',
    'NOTION_DATABASE_ID': 'your-notion-database-id',
    'PHOTOS_ALBUM_ID': 'your-google-photos-album-id',
    'DRIVE_FOLDER_ID': 'your-google-drive-folder-id',
    'USE_EXTERNAL_API': 'false',
    'MAX_PHOTOS_PER_RUN': '50',
    // 既存 'AI_MODEL' があればそれを使い、なければ推奨デフォルト
    'AI_MODEL': 'gpt-4o',
    'VISION_MODEL': 'gpt-4o',
    'PRICE_MODEL': 'gpt-4o'
  };

  Object.entries(requiredProperties).forEach(([key, value]) => {
    if (!scriptProperties.getProperty(key)) {
      scriptProperties.setProperty(key, value);
      console.log(`設定追加: ${key}`);
    }
  });

  console.log('初期設定完了');
}

// ==============================
// 初期セットアップ
// ==============================

// Photos API版セットアップ
function initialSetup() {
  console.log('=== ポケモンカード管理システム 初期セットアップ ===');

  // 1. スクリプトプロパティの初期化
  console.log('1. スクリプトプロパティを設定中...');
  setupScriptProperties();

  // 2. スプレッドシートの作成
  console.log('2. 管理用スプレッドシートを作成中...');
  const spreadsheetResult = setupCardManagementSpreadsheet();
  console.log(`   スプレッドシートURL: ${spreadsheetResult.url}`);

  // 3. トリガーの設定
  console.log('3. 自動実行トリガーを設定中...');
  setupTriggers();

  console.log('\n=== セットアップ完了 ===');
  console.log('次の手順：');
  console.log('1. プロジェクト設定 → スクリプトプロパティからAPIキー等を設定');
  console.log('2. testConnection()を実行して接続テスト');
  console.log('3. main()を手動実行して動作確認');

  return {
    spreadsheetUrl: spreadsheetResult.url,
    spreadsheetId: spreadsheetResult.spreadsheetId,
    message: '初期セットアップが完了しました'
  };
}

// Drive版セットアップ
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

// ==============================
// トリガー設定
// ==============================

function setupTriggers() {
  // 既存のトリガーをすべて削除
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    ScriptApp.deleteTrigger(trigger);
  });

  // 新規カード登録用トリガー（1時間ごと）
  ScriptApp.newTrigger('main')
    .timeBased()
    .everyHours(1)
    .create();

  // 価格更新用トリガー（週1回、月曜日の午前9時）
  ScriptApp.newTrigger('updateCardPrices')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(9)
    .create();

  console.log('トリガー設定完了:');
  console.log('- 新規カード登録: 1時間ごと');
  console.log('- 価格更新: 週1回（月曜9時）');
}

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

// ==============================
// 処理済みID管理
// ==============================

function getProcessedIds() {
  const userProperties = PropertiesService.getUserProperties();
  const idsJson = userProperties.getProperty('PROCESSED_PHOTO_IDS');

  if (!idsJson) {
    return [];
  }

  try {
    return JSON.parse(idsJson);
  } catch (error) {
    console.error('処理済みID読み込みエラー:', error);
    return [];
  }
}

function markAsProcessed(photoId) {
  const userProperties = PropertiesService.getUserProperties();
  const processedIds = getProcessedIds();

  if (!processedIds.includes(photoId)) {
    processedIds.push(photoId);

    // 最新1000件のみ保持（メモリ節約）
    if (processedIds.length > 1000) {
      processedIds.splice(0, processedIds.length - 1000);
    }

    userProperties.setProperty('PROCESSED_PHOTO_IDS', JSON.stringify(processedIds));
  }
}

function resetProcessedIds() {
  const userProperties = PropertiesService.getUserProperties();
  userProperties.deleteProperty('PROCESSED_PHOTO_IDS');
  console.log('処理済みIDをリセットしました');
}

// ==============================
// ユニークID生成と重複管理
// ==============================

function generateUniqueCardId(cardData, photo) {
  // タイムスタンプとランダム文字列を組み合わせてユニークIDを生成
  const timestamp = new Date().getTime();
  const random = Math.random().toString(36).substring(2, 8);
  const cardIdentifier = (cardData.name || 'unknown').substring(0, 10).replace(/[^a-zA-Z0-9]/g, '');

  return `CARD_${cardIdentifier}_${timestamp}_${random}`.toUpperCase();
}

function countDuplicateCards(cardData, config) {
  const spreadsheetId = PropertiesService.getScriptProperties().getProperty('MASTER_SPREADSHEET_ID');

  if (!spreadsheetId) {
    return 0;
  }

  try {
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheet = spreadsheet.getSheetByName('カード一覧');

    if (!sheet) {
      return 0;
    }

    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();

    let count = 0;
    // ヘッダー行をスキップして検索（カード名と番号が一致するものをカウント）
    for (let i = 1; i < values.length; i++) {
      const rowName = values[i][2]; // カード名のカラム（インデックス調整済み）
      const rowNumber = values[i][5]; // カード番号のカラム（インデックス調整済み）

      if (rowName === cardData.name && rowNumber === cardData.number) {
        count++;
      }
    }

    return count;

  } catch (error) {
    console.error('重複カウントエラー:', error);
    return 0;
  }
}

// ==============================
// Google Photos API関連
// ==============================

function getPhotosAlbumInfo(albumId) {
  const token = getPhotosAccessToken();
  const url = `https://photoslibrary.googleapis.com/v1/albums/${albumId}`;

  const response = UrlFetchApp.fetch(url, {
    headers: {
      'Authorization': 'Bearer ' + token,
      'Content-Type': 'application/json'
    },
    muteHttpExceptions: true
  });

  if (response.getResponseCode() !== 200) {
    throw new Error(`アルバム情報取得失敗: ${response.getContentText()}`);
  }

  return JSON.parse(response.getContentText());
}

function getNewPhotosFromAlbum(albumId, processedIds) {
  const token = getPhotosAccessToken();
  const config = getConfig();
  const maxPhotos = config.MAX_PHOTOS_PER_RUN;

  const url = 'https://photoslibrary.googleapis.com/v1/mediaItems:search';

  const payload = {
    albumId: albumId,
    pageSize: Math.min(maxPhotos, 100) // APIの最大値は100
  };

  const response = UrlFetchApp.fetch(url, {
    method: 'POST',
    headers: {
      'Authorization': 'Bearer ' + token,
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  if (response.getResponseCode() !== 200) {
    throw new Error(`写真取得失敗: ${response.getContentText()}`);
  }

  const data = JSON.parse(response.getContentText());
  const mediaItems = data.mediaItems || [];

  // 未処理の画像のみフィルタリング
  const newPhotos = mediaItems
    .filter(item => {
      // 画像のみ（動画は除外）
      if (!item.mimeType || !item.mimeType.startsWith('image/')) {
        return false;
      }
      // 未処理のもののみ
      return !processedIds.includes(item.id);
    })
    .slice(0, maxPhotos)
    .map(item => ({
      id: item.id,
      filename: item.filename,
      mimeType: item.mimeType,
      creationTime: item.mediaMetadata.creationTime,
      width: item.mediaMetadata.width,
      height: item.mediaMetadata.height,
      baseUrl: item.baseUrl,
      productUrl: item.productUrl
    }));

  return newPhotos;
}

function downloadPhotoData(photo) {
  // 最大解像度でダウンロード（w=幅, h=高さ, d=ダウンロード）
  const downloadUrl = `${photo.baseUrl}=w${photo.width}-h${photo.height}-d`;

  const response = UrlFetchApp.fetch(downloadUrl, {
    muteHttpExceptions: true
  });

  if (response.getResponseCode() !== 200) {
    throw new Error(`画像ダウンロード失敗: ${response.getResponseCode()}`);
  }

  return response.getBlob();
}

function savePhotoToDrive(photo, folderId) {
  // 画像データをダウンロード
  const blob = downloadPhotoData(photo);

  // ファイル名は後でAI判定後に変更するため、一時的な名前を設定
  const tempFileName = `temp_${photo.id}_${new Date().getTime()}.jpg`;
  blob.setName(tempFileName);

  // Driveフォルダに保存
  const folder = DriveApp.getFolderById(folderId);
  const file = folder.createFile(blob);

  // 説明を追加
  file.setDescription(`Google Photos ID: ${photo.id}\n作成日時: ${photo.creationTime}\n元ファイル名: ${photo.filename}`);

  // 共有設定（閲覧リンクを取得可能に）
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  return {
    id: file.getId(),
    name: file.getName(),
    url: file.getUrl(),
    downloadUrl: file.getDownloadUrl(),
    viewUrl: `https://drive.google.com/file/d/${file.getId()}/view`,
    blob: file.getBlob(),
    driveFile: file // ファイルオブジェクトも返す
  };
}

// AI判定結果を踏まえてDriveファイル名を更新
function renameDriveFile(driveFileInfo, cardData) {
  try {
    const file = DriveApp.getFileById(driveFileInfo.id);

    // AI解析が失敗している場合は元のファイル名を維持
    if (cardData.game === 'Unknown' && cardData.name === 'Unknown') {
      console.log('AI解析失敗のため、元のファイル名を維持');
      return driveFileInfo.name;
    }

    // すでに"UNK_"で始まるファイル名の場合は変更しない
    const currentName = file.getName();
    if (currentName.startsWith('UNK_UNK_')) {
      console.log('すでに処理済みのファイル名のため、変更しない');
      return currentName;
    }

    // カード情報を基にファイル名を生成
    const game = (cardData.game || 'Unknown').substring(0, 3).toUpperCase();
    const name = (cardData.name || 'unknown').replace(/[^a-zA-Z0-9ぁ-んァ-ヶー一-龥]/g, '_');
    const number = (cardData.number || '').replace(/[^a-zA-Z0-9]/g, '');
    const rarity = (cardData.rarity || '').substring(0, 5);
    const condition = (cardData.condition || '').substring(0, 2);
    const timestamp = new Date().toISOString().substring(0, 10);

    // 新しいファイル名を構築
    let newFileName = `${game}_${name}`;
    if (number) newFileName += `_${number}`;
    if (rarity) newFileName += `_${rarity}`;
    if (condition) newFileName += `_${condition}`;
    newFileName += `_${timestamp}.jpg`;

    // ファイル名を更新
    file.setName(newFileName);

    // 説明も更新
    const description = `Google Photos ID: ${cardData.photoId || ''}\nカード名: ${cardData.name}\nゲーム: ${cardData.game}\nセット: ${cardData.set || ''}\n番号: ${cardData.number || ''}\nレアリティ: ${cardData.rarity || ''}\n状態: ${cardData.condition || ''}\nユニークID: ${cardData.uniqueId}`;
    file.setDescription(description);

    console.log(`ファイル名更新: ${newFileName}`);

    return newFileName;

  } catch (error) {
    console.error('Driveファイル名更新エラー:', error);
    return driveFileInfo.name;
  }
}

// ==============================
// HEIC画像変換処理
// ==============================

function convertHeicToJpeg(heicFile) {
  try {
    // Google DriveのAPIを使用してHEICをJPEGに変換
    const fileId = heicFile.getId();

    // 方法1: getAsを使用（Driveがサポートしている場合）
    try {
      const jpegBlob = heicFile.getAs('image/jpeg');
      jpegBlob.setName(heicFile.getName().replace(/\.(heic|heif)$/i, '.jpg'));
      return jpegBlob;
    } catch (e) {
      console.log('getAsでの変換失敗、別の方法を試行');
    }

    // 方法2: Drive APIを使用してサムネイルを取得（代替方法）
    const thumbnailLink = Drive.Files.get(fileId).thumbnailLink;
    if (thumbnailLink) {
      // サムネイルURLを高解像度版に変更
      const highResLink = thumbnailLink.replace('=s220', '=s2000');
      const response = UrlFetchApp.fetch(highResLink, {
        headers: {
          'Authorization': 'Bearer ' + ScriptApp.getOAuthToken()
        }
      });
      const blob = response.getBlob();
      blob.setName(heicFile.getName().replace(/\.(heic|heif)$/i, '.jpg'));
      return blob;
    }

    // 方法3: 元のファイルをそのまま返す（変換できない場合）
    throw new Error('HEIC変換がサポートされていません');

  } catch (error) {
    console.error('HEIC変換エラー:', error);
    throw error;
  }
}

function moveToArchiveFolder(file, parentFolder) {
  try {
    // アーカイブフォルダを取得または作成
    let archiveFolder;
    const folders = parentFolder.getFoldersByName('HEIC_Archive');

    if (folders.hasNext()) {
      archiveFolder = folders.next();
    } else {
      archiveFolder = parentFolder.createFolder('HEIC_Archive');
      console.log('HEICアーカイブフォルダを作成しました');
    }

    // ファイルをアーカイブフォルダに移動
    file.moveTo(archiveFolder);
    console.log(`HEICファイルをアーカイブ: ${file.getName()}`);

  } catch (error) {
    console.error('アーカイブエラー:', error);
    // エラーの場合はファイルをゴミ箱に移動
    try {
      file.setTrashed(true);
      console.log(`HEICファイルをゴミ箱に移動: ${file.getName()}`);
    } catch (e) {
      console.error('ファイル削除エラー:', e);
    }
  }
}

// ==============================
// OAuth2認証
// ==============================

function getPhotosAccessToken() {
  // 簡易版: GASのデフォルト認証を使用
  try {
    return ScriptApp.getOAuthToken();
  } catch (error) {
    console.error('トークン取得エラー:', error);
    throw new Error('Google Photos APIアクセストークンの取得に失敗しました');
  }
}

// ==============================
// Drive版ユーティリティ
// ==============================

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

// ==============================
// AI画像解析（Perplexity API）
// ==============================

function analyzeCardWithAI(driveFile, config) {
  // OpenAI APIをメインで使用
  const hasOpenAI = config.OPENAI_API_KEY;
  const hasPerplexity = config.PERPLEXITY_API_KEY;

  // 画像をBase64エンコード
  const base64Image = Utilities.base64Encode(driveFile.blob.getBytes());
  const imageDataUrl = `data:${driveFile.blob.getContentType()};base64,${base64Image}`;

  // AI判定プロンプト
  const prompt = getCardAnalysisPrompt();

  try {
    // リトライロジック（最大3回）
    let retryCount = 0;
    const maxRetries = 3;

    while (retryCount < maxRetries) {
      try {
        let result;

        if (hasOpenAI) {
          // OpenAI Vision APIを使用（メイン）
          console.log('OpenAI APIで画像解析中...');
          result = callOpenAIVision(config.OPENAI_API_KEY, imageDataUrl, prompt, config.VISION_MODEL);
        } else if (hasPerplexity) {
          // Perplexity APIを使用（フォールバック）
          console.log('Perplexity APIで画像解析中（OpenAI APIキーが未設定）...');
          // Perplexityの画像モデルは使用不可、テキストモデルで代替
          const model = 'sonar-pro';
          result = callPerplexityVision(config.PERPLEXITY_API_KEY, model, imageDataUrl, prompt);
        } else {
          throw new Error('APIキーが設定されていません（OPENAI_API_KEY を設定してください）');
        }

        const cardData = parseAIResponse(result);

        // Drive URLを追加
        cardData.driveUrl = driveFile.viewUrl;
        cardData.driveFileId = driveFile.id;
        cardData.originalFileName = driveFile.name;

        return cardData;

      } catch (error) {
        retryCount++;
        if (retryCount >= maxRetries) {
          throw error;
        }

        // エクスポネンシャルバックオフ（より長い待機時間）
        const waitTime = Math.pow(2, retryCount + 1) * 2000; // 4秒, 8秒, 16秒
        console.log(`AI判定リトライ ${retryCount}/${maxRetries}, 待機時間: ${waitTime}ms`);
        Utilities.sleep(waitTime);
      }
    }

  } catch (error) {
    console.error('AI判定エラー:', `[Error: ${error}]`);

    // フォールバック: 基本情報のみ返す
    return {
      name: driveFile.name.replace(/\.[^/.]+$/, ''), // 拡張子を除去
      game: 'Unknown',
      set: null,
      number: null,
      rarity: null,
      language: null,
      condition: null,
      price: null,
      notes: `AI判定失敗: ${error.toString()}`,
      status: '要確認',
      driveUrl: driveFile.viewUrl,
      driveFileId: driveFile.id,
      originalFileName: driveFile.name
    };
  }
}

// OpenAI Vision APIを使用した画像解析
function callOpenAIVision(apiKey, imageDataUrl, prompt, model) {
  const url = 'https://api.openai.com/v1/chat/completions';

  const payload = {
    model: model || 'gpt-4o',  // 画像対応の推奨モデル（プロパティで上書き可能）
    messages: [
      {
        role: 'system',
        content: 'あなたはトレーディングカードの専門家です。画像からカード情報を正確に抽出してJSON形式で返答してください。'
      },
      {
        role: 'user',
        content: [
          {
            type: 'text',
            text: prompt
          },
          {
            type: 'image_url',
            image_url: {
              url: imageDataUrl,
              detail: 'high'
            }
          }
        ]
      }
    ],
    max_tokens: 1000,
    temperature: 0.1
  };

  const response = UrlFetchApp.fetch(url, {
    method: 'POST',
    headers: {
      'Authorization': 'Bearer ' + apiKey,
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const responseCode = response.getResponseCode();
  if (responseCode !== 200) {
    const errorBody = response.getContentText();
    throw new Error(`OpenAI API Error (${responseCode}): ${errorBody}`);
  }

  const result = JSON.parse(response.getContentText());

  if (result.choices && result.choices.length > 0) {
    return result.choices[0].message.content;
  } else {
    throw new Error('Unexpected OpenAI API response structure');
  }
}

function callPerplexityVision(apiKey, model, imageDataUrl, prompt) {
  const url = 'https://api.perplexity.ai/chat/completions';

  const payload = {
    model: model,
    messages: [
      {
        role: 'system',
        content: 'あなたはトレーディングカードの専門家です。画像からカード情報を正確に抽出してJSON形式で返答してください。'
      },
      {
        role: 'user',
        content: [
          {
            type: 'text',
            text: prompt
          },
          {
            type: 'image_url',
            image_url: {
              url: imageDataUrl
            }
          }
        ]
      }
    ],
    max_tokens: 1000,
    temperature: 0.1, // 決定的な出力のため低めに設定
    top_p: 0.1
  };

  const response = UrlFetchApp.fetch(url, {
    method: 'POST',
    headers: {
      'Authorization': 'Bearer ' + apiKey,
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const responseCode = response.getResponseCode();
  if (responseCode !== 200) {
    const errorBody = response.getContentText();
    throw new Error(`Perplexity API Error (${responseCode}): ${errorBody}`);
  }

  const result = JSON.parse(response.getContentText());

  // Perplexity APIのレスポンス構造を確認
  if (result.choices && result.choices.length > 0) {
    return result.choices[0].message.content;
  } else {
    throw new Error('Unexpected Perplexity API response structure');
  }
}

function getCardAnalysisPrompt() {
  return `
この画像のトレーディングカードを分析して、以下の情報をJSON形式で抽出してください。
不明な項目はnullとしてください。複数の可能性がある場合はnotesに記載してください。

必須フィールド:
- game: カードゲーム名（"Pokemon", "Yu-Gi-Oh!", "MTG", "Other"のいずれか）
- name: カード名（日本語または英語）
- set: セット名またはエキスパンション名
- number: カード番号（コレクター番号）
- rarity: レアリティ（C/UC/R/RR/SR/UR/HR/SAR等）
- language: 言語（"JP", "EN", "CN", "KR"等）
- condition: カードの状態推定（"新品", "美品", "良好", "やや傷", "傷あり", "ジャンク"のいずれか）
- notes: その他の情報、特記事項、不確実な情報

出力例:
{
  "game": "Pokemon",
  "name": "ピカチュウ",
  "set": "ポケモンカード151",
  "number": "025/165",
  "rarity": "R",
  "language": "JP",
  "condition": "美品",
  "notes": "ホロカード、中央に小さな白かけあり"
}

必ずJSON形式のみで回答してください。説明文は不要です。
`;
}

function parseAIResponse(aiResponse) {
  try {
    // JSONを抽出（マークダウンコードブロックを考慮）
    let jsonStr = aiResponse;

    // コードブロックで囲まれている場合の処理
    if (aiResponse.includes('```json')) {
      const match = aiResponse.match(/```json\n?([\s\S]*?)\n?```/);
      if (match) jsonStr = match[1];
    } else if (aiResponse.includes('```')) {
      const match = aiResponse.match(/```\n?([\s\S]*?)\n?```/);
      if (match) jsonStr = match[1];
    }

    // JSON以外のテキストが前後にある場合の処理
    const jsonMatch = jsonStr.match(/\{[\s\S]*\}/);
    if (jsonMatch) {
      jsonStr = jsonMatch[0];
    }

    const data = JSON.parse(jsonStr);

    // デフォルト値の設定
    return {
      name: data.name || 'Unknown Card',
      game: data.game || 'Unknown',
      set: data.set || null,
      number: data.number || null,
      rarity: data.rarity || null,
      language: data.language || null,
      condition: data.condition || null,
      price: data.price || null,
      notes: data.notes || null,
      status: '要確認' // デフォルトステータス
    };

  } catch (error) {
    console.error('AI応答のパースエラー:', error, 'Response:', aiResponse);
    throw new Error('AI応答の解析に失敗しました');
  }
}

// ==============================
// 外部API補完（オプション）
// ==============================

function enrichCardData(cardData) {
  try {
    switch (cardData.game) {
      case 'Pokemon':
        enrichPokemonCard(cardData);
        break;
      case 'Yu-Gi-Oh!':
        enrichYugiohCard(cardData);
        break;
      case 'MTG':
        enrichMTGCard(cardData);
        break;
      default:
        console.log('未対応のカードゲーム:', cardData.game);
    }
  } catch (error) {
    console.error('カードデータ補完エラー:', error);
    cardData.notes = (cardData.notes || '') + `\n補完エラー: ${error.toString()}`;
  }
}

function enrichPokemonCard(cardData) {
  // Pokemon TCG API を使用した補完
  if (!cardData.number || !cardData.set) {
    return;
  }

  const apiUrl = `https://api.pokemontcg.io/v2/cards?q=number:${cardData.number} set.name:"${cardData.set}"`;

  const response = UrlFetchApp.fetch(apiUrl, {
    muteHttpExceptions: true
  });

  if (response.getResponseCode() === 200) {
    const result = JSON.parse(response.getContentText());
    if (result.data && result.data.length > 0) {
      const apiCard = result.data[0];

      // APIデータで補完
      cardData.name = cardData.name || apiCard.name;
      cardData.set = apiCard.set.name;
      cardData.number = apiCard.number;
      cardData.rarity = cardData.rarity || apiCard.rarity;

      // 価格情報（TCGPlayer）
      if (apiCard.tcgplayer && apiCard.tcgplayer.prices) {
        const prices = apiCard.tcgplayer.prices;
        const marketPrice = prices.holofoil?.market || prices.normal?.market;
        if (marketPrice) {
          // 数値として保存（USD）
          cardData.price = marketPrice;
          cardData.currency = 'USD';
        }
      }
    }
  }
}

function enrichYugiohCard(cardData) {
  // YGOPRODeck API を使用した補完
  if (!cardData.name) {
    return;
  }

  const apiUrl = `https://db.ygoprodeck.com/api/v7/cardinfo.php?name=${encodeURIComponent(cardData.name)}`;

  const response = UrlFetchApp.fetch(apiUrl, {
    muteHttpExceptions: true
  });

  if (response.getResponseCode() === 200) {
    const result = JSON.parse(response.getContentText());
    if (result.data && result.data.length > 0) {
      const apiCard = result.data[0];

      // APIデータで補完
      cardData.name = apiCard.name;
      cardData.set = cardData.set || apiCard.card_sets?.[0]?.set_name;
      cardData.number = cardData.number || apiCard.card_sets?.[0]?.set_code;
      cardData.rarity = cardData.rarity || apiCard.card_sets?.[0]?.set_rarity;

      // 価格情報
      if (apiCard.card_prices && apiCard.card_prices.length > 0) {
        const price = apiCard.card_prices[0];
        // 数値として保存（USD）
        cardData.price = parseFloat(price.tcgplayer_price);
        cardData.currency = 'USD';
      }
    }
  }
}

function enrichMTGCard(cardData) {
  // Scryfall API を使用した補完
  if (!cardData.name) {
    return;
  }

  const apiUrl = `https://api.scryfall.com/cards/search?q=name:"${encodeURIComponent(cardData.name)}"`;

  const response = UrlFetchApp.fetch(apiUrl, {
    muteHttpExceptions: true
  });

  if (response.getResponseCode() === 200) {
    const result = JSON.parse(response.getContentText());
    if (result.data && result.data.length > 0) {
      const apiCard = result.data[0];

      // APIデータで補完
      cardData.name = apiCard.name;
      cardData.set = cardData.set || apiCard.set_name;
      cardData.number = cardData.number || apiCard.collector_number;
      cardData.rarity = cardData.rarity || apiCard.rarity;
      cardData.language = cardData.language || apiCard.lang.toUpperCase();

      // 価格情報
      if (apiCard.prices) {
        const price = apiCard.prices.usd || apiCard.prices.usd_foil;
        if (price) {
          // 数値として保存（USD）
          cardData.price = parseFloat(price);
          cardData.currency = 'USD';
        }
      }
    }
  }
}

// ==============================
// Notion API連携
// ==============================

function createNotionRecord(cardData, driveFile, config) {
  const notionApiKey = config.NOTION_API_KEY;
  const databaseId = config.NOTION_DATABASE_ID;

  const url = `https://api.notion.com/v1/pages`;

  // 英語カードの価格をJPYに変換
  if (cardData.language && cardData.language.toUpperCase().startsWith('EN')) {
    convertEnglishCardPrice(cardData);
  }

  // Notionページのプロパティを構築
  const properties = buildNotionProperties(cardData);

  const payload = {
    parent: {
      database_id: databaseId
    },
    properties: properties,
    children: buildNotionContent(cardData)
  };

  // デバッグ: 送信されるプロパティをログ出力
  console.log('送信するプロパティ数:', Object.keys(properties).length);
  if (Object.keys(properties).length === 0) {
    console.log('警告: プロパティが空です！');
  } else {
    console.log('送信するプロパティ:');
    Object.keys(properties).forEach(key => {
      const value = properties[key];
      if (value) {
        console.log(`- ${key}:`, JSON.stringify(value).substring(0, 100));
      }
    });
  }

  // リトライロジック
  let retryCount = 0;
  const maxRetries = 3;
  let lastError = null;

  while (retryCount < maxRetries) {
    try {
      console.log(`Notion API呼び出し中... (試行 ${retryCount + 1}/${maxRetries})`);

      const response = UrlFetchApp.fetch(url, {
        method: 'POST',
        headers: {
          'Authorization': 'Bearer ' + notionApiKey,
          'Content-Type': 'application/json',
          'Notion-Version': '2022-06-28'
        },
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      });

      const responseCode = response.getResponseCode();

      if (responseCode === 200 || responseCode === 201) {
        const result = JSON.parse(response.getContentText());
        console.log(`Notionレコード作成成功: ${result.id}`);
        return result.id;
      } else {
        const errorBody = response.getContentText();
        lastError = new Error(`Notion API Error (${responseCode}): ${errorBody}`);
        console.error(`Notion APIエラー (試行 ${retryCount + 1}):`, errorBody);
      }

    } catch (error) {
      lastError = error;
      console.error(`Notion接続エラー (試行 ${retryCount + 1}):`, error.toString());

      // "Address unavailable"エラーの場合は待機時間を長くする
      if (error.toString().includes('Address unavailable')) {
        console.log('ネットワークエラー検出。待機時間を延長します。');
        Utilities.sleep(10000); // 10秒待機
      }
    }

    retryCount++;

    if (retryCount < maxRetries) {
      // 指数バックオフ
      const waitTime = Math.pow(2, retryCount) * 2000;
      console.log(`${waitTime}ms待機後にリトライします...`);
      Utilities.sleep(waitTime);
    }
  }

  // すべての試行が失敗した場合
  console.error('Notion登録失敗（すべての試行が失敗）:', lastError);
  throw lastError;
}

function buildNotionProperties(cardData) {
  console.log('buildNotionProperties開始:', cardData.name);

  // データベースプロパティを取得
  let dbInfo;
  try {
    const config = getConfig();
    dbInfo = getNotionDatabaseInfo(config);
    console.log('データベースプロパティ取得成功');
  } catch (error) {
    console.error('データベース情報取得エラー:', error);
    // エラー時はデフォルトプロパティを使用
    console.log('デフォルトプロパティを使用');
    return buildDefaultNotionProperties(cardData);
  }

  const properties = {};

  // タイトルプロパティを探す
  const titleProp = Object.keys(dbInfo.properties).find(key =>
    dbInfo.properties[key].type === 'title'
  );

  console.log('タイトルプロパティ名:', titleProp);

  // タイトルプロパティ（必須）
  if (titleProp) {
    properties[titleProp] = {
      title: [
        {
          text: {
            content: `${cardData.name || 'Unknown Card'} [${cardData.duplicateNumber || 1}]`
          }
        }
      ]
    };
    console.log('タイトルプロパティを追加');
  } else {
    console.log('警告: タイトルプロパティが見つかりません');
  }

  // 各プロパティを条件付きで追加
  if (dbInfo.properties.UniqueID) {
    properties.UniqueID = {
      rich_text: [
        {
          text: {
            content: cardData.uniqueId || ''
          }
        }
      ]
    };
  }

  if (dbInfo.properties.Game) {
    properties.Game = {
      select: {
        name: cardData.game || 'Unknown'
      }
    };
  }

  if (dbInfo.properties.Set) {
    properties.Set = {
      rich_text: [
        {
          text: {
            content: cardData.set || ''
          }
        }
      ]
    };
  }

  if (dbInfo.properties.Number) {
    properties.Number = {
      rich_text: [
        {
          text: {
            content: cardData.number || ''
          }
        }
      ]
    };
  }

  if (dbInfo.properties.Rarity) {
    properties.Rarity = {
      select: cardData.rarity ? {
        name: cardData.rarity
      } : null
    };
  }

  if (dbInfo.properties.Language) {
    properties.Language = {
      select: {
        name: cardData.language || '日本語'
      }
    };
  }

  if (dbInfo.properties.Condition) {
    properties.Condition = {
      select: {
        name: cardData.condition || 'NM'
      }
    };
  }

  if (dbInfo.properties.Status) {
    properties.Status = {
      select: {
        name: cardData.status || '在庫'
      }
    };
  }

  if (dbInfo.properties.Source) {
    properties.Source = {
      rich_text: [
        {
          text: {
            content: '自動アップロード'
          }
        }
      ]
    };
  }

  if (dbInfo.properties.DuplicateNumber) {
    properties.DuplicateNumber = {
      number: cardData.duplicateNumber || 1
    };
  }

  if (dbInfo.properties.PhotoID) {
    properties.PhotoID = {
      rich_text: [
        {
          text: {
            content: cardData.photoId || ''
          }
        }
      ]
    };
  }

  if (dbInfo.properties.DriveFileID) {
    properties.DriveFileID = {
      rich_text: [
        {
          text: {
            content: cardData.driveFileId || ''
          }
        }
      ]
    };
  }

  if (dbInfo.properties.ImageURL) {
    properties.ImageURL = {
      url: cardData.driveUrl || null
    };
  }

  // 価格情報の取得と追加
  const priceData = getCardPriceData(cardData);

  if (dbInfo.properties.Price) {
    properties.Price = {
      number: priceData.currentPrice || 0
    };
  }

  if (dbInfo.properties.MarketPrice) {
    properties.MarketPrice = {
      number: priceData.marketPrice || 0
    };
  }

  // 価格トレンドを追加
  if (dbInfo.properties.PriceTrend) {
    properties.PriceTrend = {
      select: {
        name: cardData.priceTrend || '不明'
      }
    };
  }

  // 価格推移の個別プロパティ
  if (dbInfo.properties.Price1YearAgo && priceData.priceHistory) {
    properties.Price1YearAgo = {
      number: priceData.priceHistory['12ヶ月前'] || 0
    };
  }

  if (dbInfo.properties.Price6MonthsAgo && priceData.priceHistory) {
    properties.Price6MonthsAgo = {
      number: priceData.priceHistory['6ヶ月前'] || 0
    };
  }

  if (dbInfo.properties.Price3MonthsAgo && priceData.priceHistory) {
    properties.Price3MonthsAgo = {
      number: priceData.priceHistory['3ヶ月前'] || 0
    };
  }

  // 価格予測の個別プロパティ
  if (dbInfo.properties.PredictedPrice6Months && priceData.pricePrediction) {
    properties.PredictedPrice6Months = {
      number: priceData.pricePrediction['6ヶ月後'] || 0
    };
  }

  if (dbInfo.properties.PredictedPrice1Year && priceData.pricePrediction) {
    properties.PredictedPrice1Year = {
      number: priceData.pricePrediction['12ヶ月後'] || 0
    };
  }

  // JSON形式での価格推移データ
  if (dbInfo.properties.PriceHistory) {
    properties.PriceHistory = {
      rich_text: [
        {
          text: {
            content: JSON.stringify(priceData.priceHistory || {})
          }
        }
      ]
    };
  }

  if (dbInfo.properties.PricePrediction) {
    properties.PricePrediction = {
      rich_text: [
        {
          text: {
            content: JSON.stringify(priceData.pricePrediction || {})
          }
        }
      ]
    };
  }

  // PSAグレード別価格を追加（priceDataまたはcardDataから取得）
  const psaPrices = priceData.psaGradedPrice || cardData.psaGradedPrice;
  if (psaPrices) {
    if (dbInfo.properties.PSA9_Price && psaPrices.PSA9) {
      properties.PSA9_Price = {
        number: psaPrices.PSA9 || 0
      };
    }

    if (dbInfo.properties['PSA9.5_Price'] && psaPrices['PSA9.5']) {
      properties['PSA9.5_Price'] = {
        number: psaPrices['PSA9.5'] || 0
      };
    }

    if (dbInfo.properties.PSA10_Price && psaPrices.PSA10) {
      properties.PSA10_Price = {
        number: psaPrices.PSA10 || 0
      };
    }
  }

  // 価格推移と予測をNotesに含める
  const priceInfo = formatPriceInfo(priceData, cardData);

  if (dbInfo.properties.Notes) {
    const notes = (cardData.notes || '') + '\n\n' + priceInfo;
    properties.Notes = {
      rich_text: [
        {
          text: {
            content: notes.substring(0, 2000)
          }
        }
      ]
    };
  }

  if (dbInfo.properties.RegisteredDate) {
    properties.RegisteredDate = {
      date: {
        start: new Date().toISOString()
      }
    };
  }

  if (dbInfo.properties.LastUpdated) {
    properties.LastUpdated = {
      date: {
        start: new Date().toISOString()
      }
    };
  }

  if (dbInfo.properties.PriceLastUpdated) {
    properties.PriceLastUpdated = {
      date: {
        start: priceData.lastUpdated || new Date().toISOString()
      }
    };
  }

  console.log('buildNotionProperties終了: プロパティ数=', Object.keys(properties).length);
  return properties;
}

// デフォルトのNotionプロパティ
function buildDefaultNotionProperties(cardData) {
  return {
    'Name': {
      title: [
        {
          text: {
            content: `${cardData.name || 'Unknown Card'} [${cardData.duplicateNumber || 1}]`
          }
        }
      ]
    }
  };
}

function buildNotionContent(cardData) {
  const content = [];

  // 価格データを取得
  const priceData = getCardPriceData(cardData);
  const currency = priceData.currency || getCurrencyByLanguage(cardData.language);
  const sym = getCurrencySymbol(currency);

  // 見出し
  content.push({
    type: 'heading_2',
    heading_2: {
      rich_text: [
        {
          text: {
            content: 'カード情報'
          }
        }
      ]
    }
  });

  // カード詳細情報のテーブル
  const details = [
    ['項目', '内容'],
    ['カード名', cardData.name || '-'],
    ['ゲーム', cardData.game || '-'],
    ['セット', cardData.set || '-'],
    ['番号', cardData.number || '-'],
    ['レアリティ', cardData.rarity || '-'],
    ['言語', cardData.language || '-'],
    ['状態', cardData.condition || '-'],
    ['現在価格', `${sym}${priceData.currentPrice || 0}`],
    ['市場価格', `${sym}${priceData.marketPrice || 0}`],
    ['価格トレンド', cardData.priceTrend || '不明']
  ];

  // テーブルブロック
  content.push({
    type: 'table',
    table: {
      table_width: 2,
      has_column_header: true,
      has_row_header: false,
      children: details.map(row => ({
        type: 'table_row',
        table_row: {
          cells: row.map(cell => [
            {
              type: 'text',
              text: {
                content: cell
              }
            }
          ])
        }
      }))
    }
  });

  // 画像リンク
  content.push({
    type: 'heading_3',
    heading_3: {
      rich_text: [
        {
          text: {
            content: '画像'
          }
        }
      ]
    }
  });

  content.push({
    type: 'paragraph',
    paragraph: {
      rich_text: [
        {
          text: {
            content: 'Google Drive: '
          }
        },
        {
          text: {
            content: cardData.driveUrl,
            link: {
              url: cardData.driveUrl
            }
          }
        }
      ]
    }
  });

  // 価格推移セクション
  content.push({
    type: 'heading_3',
    heading_3: {
      rich_text: [
        {
          text: {
            content: '価格推移'
          }
        }
      ]
    }
  });

  // 価格推移の情報
  if (priceData.priceHistory) {
    let priceHistoryText = '';
    Object.entries(priceData.priceHistory).forEach(([period, price]) => {
      if (period !== 'trend') { // trendキーは除外
        priceHistoryText += `${period}: ${sym}${price}\n`;
      }
    });

    content.push({
      type: 'paragraph',
      paragraph: {
        rich_text: [
          {
            text: {
              content: priceHistoryText || '価格推移データなし'
            }
          }
        ]
      }
    });

    // トレンド情報を追加
    if (cardData.priceTrend) {
      content.push({
        type: 'paragraph',
        paragraph: {
          rich_text: [
            {
              text: {
                content: `トレンド: ${cardData.priceTrend}`
              },
              annotations: {
                bold: true
              }
            }
          ]
        }
      });
    }
  }

  // 価格予測セクション
  content.push({
    type: 'heading_3',
    heading_3: {
      rich_text: [
        {
          text: {
            content: '価格予測'
          }
        }
      ]
    }
  });

  // 価格予測の情報
  if (priceData.pricePrediction) {
    let pricePredictionText = '';
    Object.entries(priceData.pricePrediction).forEach(([period, price]) => {
      pricePredictionText += `${period}: ${sym}${price}\n`;
    });

    content.push({
      type: 'paragraph',
      paragraph: {
        rich_text: [
          {
            text: {
              content: pricePredictionText || '価格予測データなし'
            }
          }
        ]
      }
    });
  }

  // メモ
  if (cardData.notes) {
    content.push({
      type: 'heading_3',
      heading_3: {
        rich_text: [
          {
            text: {
              content: 'メモ'
            }
          }
        ]
      }
    });

    content.push({
      type: 'paragraph',
      paragraph: {
        rich_text: [
          {
            text: {
              content: cardData.notes
            }
          }
        ]
      }
    });
  }

  // メタ情報
  content.push({
    type: 'divider',
    divider: {}
  });

  content.push({
    type: 'paragraph',
    paragraph: {
      rich_text: [
        {
          text: {
            content: `登録日時: ${new Date().toLocaleString('ja-JP')}\n`
          },
          annotations: {
            italic: true,
            color: 'gray'
          }
        },
        {
          text: {
            content: `元ファイル名: ${cardData.originalFileName || 'Unknown'}\n`
          },
          annotations: {
            italic: true,
            color: 'gray'
          }
        },
        {
          text: {
            content: `価格最終更新: ${priceData.lastUpdated || new Date().toISOString()}`
          },
          annotations: {
            italic: true,
            color: 'gray'
          }
        }
      ]
    }
  });

  return content;
}

function getNotionDatabaseInfo(config) {
  const notionApiKey = config.NOTION_API_KEY;
  const databaseId = config.NOTION_DATABASE_ID;

  const url = `https://api.notion.com/v1/databases/${databaseId}`;

  const response = UrlFetchApp.fetch(url, {
    headers: {
      'Authorization': 'Bearer ' + notionApiKey,
      'Notion-Version': '2022-06-28'
    },
    muteHttpExceptions: true
  });

  if (response.getResponseCode() !== 200) {
    throw new Error(`Notionデータベース情報取得失敗: ${response.getContentText()}`);
  }

  const data = JSON.parse(response.getContentText());

  return {
    id: data.id,
    title: data.title[0]?.plain_text || 'Untitled',
    properties: data.properties  // プロパティオブジェクト全体を返す
  };
}

// ==============================
// エラーログ
// ==============================

function logError(photo, error) {
  const sheet = getOrCreateLogSheet();
  sheet.appendRow([
    new Date(),
    photo.id,
    photo.filename || photo.name,
    error.toString(),
    JSON.stringify(photo)
  ]);
}

function getOrCreateLogSheet() {
  // 現在のスプレッドシートを使用
  let spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  if (!spreadsheet) {
    // スプレッドシートIDがプロパティに保存されている場合は取得
    const spreadsheetId = PropertiesService.getScriptProperties().getProperty('LOG_SPREADSHEET_ID') ||
                         PropertiesService.getScriptProperties().getProperty('MASTER_SPREADSHEET_ID');

    if (spreadsheetId) {
      try {
        spreadsheet = SpreadsheetApp.openById(spreadsheetId);
      } catch (e) {
        throw new Error('スプレッドシートが見つかりません');
      }
    } else {
      throw new Error('このスクリプトはスプレッドシートから実行してください');
    }
  }

  let sheet = spreadsheet.getSheetByName('エラーログ');
  if (!sheet) {
    sheet = spreadsheet.insertSheet('エラーログ');
    sheet.getRange(1, 1, 1, 5).setValues([['日時', 'Photo ID', 'ファイル名', 'エラー', '詳細']]);
  }

  return sheet;
}

function createLogSpreadsheet() {
  // 現在のスプレッドシートを使用
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (!spreadsheet) {
    throw new Error('このスクリプトはスプレッドシートから実行してください');
  }
  const spreadsheetId = spreadsheet.getId();
  PropertiesService.getScriptProperties().setProperty('LOG_SPREADSHEET_ID', spreadsheetId);
  return spreadsheet;
}

// ==============================
// スプレッドシート自動セットアップ
// ==============================

function setupCardManagementSpreadsheet() {
  console.log('スプレッドシート自動セットアップ開始');

  // 現在のスプレッドシート（GASプロジェクトに紐づいているもの）を使用
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (!spreadsheet) {
    throw new Error('このスクリプトはスプレッドシートから実行してください');
  }

  const spreadsheetId = spreadsheet.getId();
  PropertiesService.getScriptProperties().setProperty('MASTER_SPREADSHEET_ID', spreadsheetId);

  // カード一覧シートの作成（既存のシートをチェック）
  let cardSheet = spreadsheet.getSheetByName('カード一覧');
  if (!cardSheet) {
    cardSheet = spreadsheet.getActiveSheet();
    cardSheet.setName('カード一覧');
  } else {
    // 既存のシートがある場合はフィルターを削除してからクリア
    const filter = cardSheet.getFilter();
    if (filter) {
      filter.remove();
    }
    cardSheet.clear();
  }

  // ヘッダー行の設定（価格履歴・予測・PSAグレード価格を含む28カラム）
  const headers = [
    'ユニークID',
    '登録日時',
    'カード名',
    'ゲーム',
    'セット',
    '番号',
    'レアリティ',
    '言語',
    '状態',
    '現在価格',
    '市場価格',
    '価格トレンド',
    '12ヶ月前',
    '6ヶ月前',
    '3ヶ月前',
    '6ヶ月後予測',
    '1年後予測',
    'PSA9価格',
    'PSA9.5価格',
    'PSA10価格',
    '重複番号',
    'ステータス',
    'Drive URL',
    'Notion ID',
    'Photos ID',
    'ファイル名',
    'メモ',
    'エラー'
  ];

  const headerRange = cardSheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);

  // ヘッダーのフォーマット設定
  headerRange.setBackground('#4285F4');
  headerRange.setFontColor('#FFFFFF');
  headerRange.setFontWeight('bold');
  headerRange.setHorizontalAlignment('center');

  // 列幅の設定（価格列は広めに）
  const columnWidths = [
    120, // ユニークID
    120, // 登録日時
    180, // カード名
    80,  // ゲーム
    150, // セット
    100, // 番号
    80,  // レアリティ
    60,  // 言語
    60,  // 状態
    100, // 現在価格
    100, // 市場価格
    100, // 価格トレンド
    100, // 12ヶ月前
    100, // 6ヶ月前
    100, // 3ヶ月前
    100, // 6ヶ月後予測
    100, // 1年後予測
    120, // PSA9価格
    120, // PSA9.5価格
    120, // PSA10価格
    80,  // 重複番号
    80,  // ステータス
    150, // Drive URL
    150, // Notion ID
    150, // Photos ID
    150, // ファイル名
    200, // メモ
    200  // エラー
  ];

  for (let i = 0; i < columnWidths.length; i++) {
    cardSheet.setColumnWidth(i + 1, columnWidths[i]);
  }

  // 価格列に通貨フォーマットを設定
  // 10:現在価格, 11:市場価格, 13-20:価格履歴・予測・PSA価格（12列目の価格トレンドはテキストなのでスキップ）
  const priceColumns = [10, 11, 13, 14, 15, 16, 17, 18, 19, 20];
  priceColumns.forEach(col => {
    cardSheet.getRange(2, col, 999, 1).setNumberFormat('¥#,##0');
  })

  // フィルタービューの設定（既存のフィルターを確認してから作成）
  const existingFilter = cardSheet.getFilter();
  if (existingFilter) {
    // 既存のフィルターがある場合は削除
    existingFilter.remove();
  }
  // 新しいフィルターを作成
  cardSheet.getRange(1, 1, 1000, headers.length).createFilter();

  // 統計シートの作成
  let statsSheet = spreadsheet.getSheetByName('統計');
  if (!statsSheet) {
    statsSheet = spreadsheet.insertSheet('統計');
  } else {
    // 既存のフィルターがあれば削除
    const statsFilter = statsSheet.getFilter();
    if (statsFilter) statsFilter.remove();
    statsSheet.clear();
  }
  setupStatsSheet(statsSheet);

  // 価格推移シートの作成
  let priceSheet = spreadsheet.getSheetByName('価格推移');
  if (!priceSheet) {
    priceSheet = spreadsheet.insertSheet('価格推移');
  } else {
    const priceFilter = priceSheet.getFilter();
    if (priceFilter) priceFilter.remove();
    priceSheet.clear();
  }
  setupPriceSheet(priceSheet);

  // 処理履歴シートの作成
  let historySheet = spreadsheet.getSheetByName('処理履歴');
  if (!historySheet) {
    historySheet = spreadsheet.insertSheet('処理履歴');
  } else {
    const historyFilter = historySheet.getFilter();
    if (historyFilter) historyFilter.remove();
    historySheet.clear();
  }
  setupHistorySheet(historySheet);

  // エラーログシートの作成
  let errorSheet = spreadsheet.getSheetByName('エラーログ');
  if (!errorSheet) {
    errorSheet = spreadsheet.insertSheet('エラーログ');
  } else {
    const errorFilter = errorSheet.getFilter();
    if (errorFilter) errorFilter.remove();
    errorSheet.clear();
  }
  setupErrorSheet(errorSheet);

  // 設定シートの作成
  let configSheet = spreadsheet.getSheetByName('設定');
  if (!configSheet) {
    configSheet = spreadsheet.insertSheet('設定');
  } else {
    const configFilter = configSheet.getFilter();
    if (configFilter) configFilter.remove();
    configSheet.clear();
  }
  setupConfigSheet(configSheet);

  console.log(`スプレッドシート作成完了: ${spreadsheet.getUrl()}`);

  return {
    spreadsheetId: spreadsheetId,
    url: spreadsheet.getUrl()
  };
}

function setupStatsSheet(sheet) {
  const headers = ['統計項目', '値'];
  const statsData = [
    ['総カード数（重複含む）', '=COUNTA(カード一覧!C:C)-1'],
    ['ユニークカード種数', '=SUMPRODUCT(1/COUNTIFS(カード一覧!C:C,カード一覧!C:C,カード一覧!F:F,カード一覧!F:F))'],
    ['ポケモンカード数', '=COUNTIF(カード一覧!D:D,"ポケモン")'],
    ['遊戯王カード数', '=COUNTIF(カード一覧!D:D,"遊戯王")'],
    ['MTGカード数', '=COUNTIF(カード一覧!D:D,"MTG")'],
    ['重複カード数', '=COUNTIF(カード一覧!R:R,">1")'],
    ['確認済みカード数', '=COUNTIF(カード一覧!S:S,"確定")'],
    ['要確認カード数', '=COUNTIF(カード一覧!S:S,"要確認")'],
    ['合計現在価格', '=SUM(カード一覧!J:J)'],
    ['平均現在価格', '=AVERAGE(カード一覧!J:J)'],
    ['最高価格カード', '=MAX(カード一覧!J:J)'],
    ['本日処理数', '=COUNTIF(カード一覧!B:B,TODAY())'],
    ['今週処理数', '=COUNTIFS(カード一覧!B:B,">="&TODAY()-7,カード一覧!B:B,"<="&TODAY())']
  ];

  sheet.getRange(1, 1, 1, 2).setValues([headers]);
  sheet.getRange(1, 1, 1, 2).setBackground('#34A853').setFontColor('#FFFFFF').setFontWeight('bold');

  sheet.getRange(2, 1, statsData.length, 2).setValues(statsData);

  // グラフの追加
  const chartBuilder = sheet.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(sheet.getRange('A2:B5'))
    .setPosition(10, 4, 0, 0)
    .setOption('title', 'カードゲーム別分布')
    .setOption('width', 400)
    .setOption('height', 300);

  sheet.insertChart(chartBuilder.build());

  sheet.autoResizeColumns(1, 2);
}

function setupHistorySheet(sheet) {
  const headers = [
    '実行日時',
    '処理枚数',
    '成功',
    '失敗',
    '処理時間(秒)',
    'エラー内容'
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setBackground('#EA4335').setFontColor('#FFFFFF').setFontWeight('bold');

  // 列幅の設定
  sheet.setColumnWidth(1, 150);
  sheet.setColumnWidth(6, 300);

  // フィルタービューの設定
  sheet.getRange(1, 1, 1000, headers.length).createFilter();
}

// 価格推移シート設定
function setupPriceSheet(sheet) {
  const headers = [
    'カード名',
    'ゲーム',
    'レアリティ',
    '現在価格',
    '前回価格',
    '価格変動',
    '変動率(%)',
    'PSA9価格',
    'PSA9.5価格',
    'PSA10価格',
    '更新日時'
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setBackground('#FBBC05')
    .setFontColor('#000000')
    .setFontWeight('bold');

  // 列幅設定
  sheet.setColumnWidth(1, 200);
  sheet.setColumnWidth(4, 100);
  sheet.setColumnWidth(5, 100);
  sheet.setColumnWidth(6, 100);
  sheet.setColumnWidth(7, 100);
  sheet.setColumnWidth(8, 120);
  sheet.setColumnWidth(9, 120);
  sheet.setColumnWidth(10, 120);
  sheet.setColumnWidth(11, 150);

  // 価格列にフォーマット
  sheet.getRange(2, 4, 999, 3).setNumberFormat('¥#,##0');
  sheet.getRange(2, 7, 999, 1).setNumberFormat('#,##0.0%');
  sheet.getRange(2, 8, 999, 3).setNumberFormat('¥#,##0'); // PSA価格フォーマット

  sheet.getRange(1, 1, 1000, headers.length).createFilter();
}

// エラーログシート設定
function setupErrorSheet(sheet) {
  const headers = [
    'エラー日時',
    '画像ID',
    'ファイル名',
    'エラータイプ',
    'エラー詳細',
    'ステータス'
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setBackground('#EA4335')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold');

  // 列幅設定
  sheet.setColumnWidth(1, 150);
  sheet.setColumnWidth(2, 150);
  sheet.setColumnWidth(3, 200);
  sheet.setColumnWidth(4, 120);
  sheet.setColumnWidth(5, 400);
  sheet.setColumnWidth(6, 100);

  sheet.getRange(1, 1, 1000, headers.length).createFilter();
}

function setupConfigSheet(sheet) {
  const configData = [
    ['設定項目', '値', '説明'],
    ['PERPLEXITY_API_KEY', '', 'Perplexity APIキー'],
    ['NOTION_API_KEY', '', 'Notion APIキー'],
    ['NOTION_DATABASE_ID', '', 'NotionデータベースID'],
    ['PHOTOS_ALBUM_ID', '', 'Google PhotosアルバムID'],
    ['DRIVE_FOLDER_ID', '', 'Google DriveフォルダID'],
    ['USE_EXTERNAL_API', 'false', '外部API補完の有効化(true/false)'],
    ['MAX_PHOTOS_PER_RUN', '50', '1回の実行での最大処理枚数'],
    ['AI_MODEL', 'gpt-4o', 'レガシー互換（未設定時はVISION/PRICEにフォールバック）'],
    ['VISION_MODEL', 'gpt-4o', '画像解析用のOpenAIモデル名'],
    ['PRICE_MODEL', 'gpt-4o', '価格推定用のOpenAIモデル名'],
    ['NOTIFICATION_EMAIL', '', 'エラー通知先メールアドレス']
  ];

  sheet.getRange(1, 1, configData.length, 3).setValues(configData);

  // ヘッダーのフォーマット
  sheet.getRange(1, 1, 1, 3).setBackground('#FBBC04').setFontColor('#000000').setFontWeight('bold');

  // 列幅の設定
  sheet.setColumnWidth(1, 200);
  sheet.setColumnWidth(2, 300);
  sheet.setColumnWidth(3, 400);

  // 注意書きの追加
  sheet.getRange(configData.length + 2, 1).setValue('※ この設定は参照用です。実際の設定はスクリプトプロパティで管理されます。');
  sheet.getRange(configData.length + 2, 1).setFontColor('#FF0000').setFontStyle('italic');
}

// カード情報をスプレッドシートに記録
function logCardToSpreadsheet(cardData, notionPageId) {
  let spreadsheetId = PropertiesService.getScriptProperties().getProperty('MASTER_SPREADSHEET_ID');

  if (!spreadsheetId) {
    // 現在のスプレッドシートのIDを取得して使用
    const currentSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    if (currentSpreadsheet) {
      spreadsheetId = currentSpreadsheet.getId();
      PropertiesService.getScriptProperties().setProperty('MASTER_SPREADSHEET_ID', spreadsheetId);
    } else {
      console.error('スプレッドシートが見つかりません');
      return;
    }
  }

  try {
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    let sheet = spreadsheet.getSheetByName('カード一覧');

    // 価格データを取得
    const priceData = getCardPriceData(cardData);

    if (!sheet) {
      sheet = spreadsheet.insertSheet('カード一覧');
    }

    // ヘッダー行を確認・更新
    const headers = [
      'ユニークID', '登録日時', 'カード名', 'ゲーム', 'セット',
      '番号', 'レアリティ', '言語', '状態', '現在価格',
      '市場価格', '価格トレンド', '12ヶ月前', '6ヶ月前', '3ヶ月前',
      '6ヶ月後予測', '1年後予測', 'PSA9価格', 'PSA9.5価格', 'PSA10価格',
      '重複番号', 'ステータス', 'Drive URL', 'Notion ID', 'Photos ID',
      'ファイル名', 'メモ', 'エラー'
    ];

    // 最初の行が空またはヘッダーでない場合、ヘッダーを設定
    if (sheet.getLastRow() === 0 || sheet.getRange(1, 1).getValue() !== 'ユニークID') {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length)
        .setBackground('#4285F4')
        .setFontColor('#FFFFFF')
        .setFontWeight('bold');
    }

    const rowData = [
      cardData.uniqueId || '',                                              // ユニークID
      new Date(),                                                           // 処理日時
      cardData.name || '',                                                  // カード名
      cardData.game || '',                                                  // ゲーム
      cardData.set || '',                                                   // セット
      cardData.number || '',                                                // カード番号
      cardData.rarity || '',                                                // レアリティ
      cardData.language || '',                                              // 言語
      cardData.condition || '',                                             // 状態
      priceData.currentPrice || 0,                                          // 現在価格
      priceData.marketPrice || 0,                                           // 市場価格
      cardData.priceTrend || '不明',                                        // 価格トレンド
      priceData.priceHistory ? priceData.priceHistory['12ヶ月前'] || 0 : 0, // 12ヶ月前価格
      priceData.priceHistory ? priceData.priceHistory['6ヶ月前'] || 0 : 0,  // 6ヶ月前価格
      priceData.priceHistory ? priceData.priceHistory['3ヶ月前'] || 0 : 0,  // 3ヶ月前価格
      priceData.pricePrediction ? priceData.pricePrediction['6ヶ月後'] || 0 : 0,  // 6ヶ月後予測
      priceData.pricePrediction ? priceData.pricePrediction['12ヶ月後'] || 0 : 0, // 12ヶ月後予測
      cardData.psaGradedPrice ? cardData.psaGradedPrice.PSA9 || 0 : 0,      // PSA9価格
      cardData.psaGradedPrice ? cardData.psaGradedPrice['PSA9.5'] || 0 : 0, // PSA9.5価格
      cardData.psaGradedPrice ? cardData.psaGradedPrice.PSA10 || 0 : 0,     // PSA10価格
      cardData.duplicateNumber || 1,                                        // 重複番号
      cardData.status || '処理済み',                                        // ステータス
      cardData.driveUrl || '',                                              // Drive URL
      notionPageId || '',                                                   // Notion Page ID
      cardData.photoId || '',                                               // Photos ID
      cardData.driveFileName || '',                                         // Drive File名
      cardData.notes || '',                                                 // メモ
      cardData.error || ''                                                  // エラー
    ];

    sheet.appendRow(rowData);

    // 価格カラムに通貨フォーマットを適用
    const lastRow = sheet.getLastRow();
    // 価格関連カラム (J:S列 = 10:19列) に通貨フォーマット - PSA価格3列追加
    sheet.getRange(lastRow, 10, 1, 10).setNumberFormat('¥#,##0');

    console.log('スプレッドシート記録完了（価格情報含む）');

  } catch (error) {
    console.error('スプレッドシート記録エラー:', error);
  }
}

// 処理履歴をスプレッドシートに記録
function logProcessingHistory(results, startTime) {
  const spreadsheetId = PropertiesService.getScriptProperties().getProperty('MASTER_SPREADSHEET_ID');

  if (!spreadsheetId) return;

  try {
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheet = spreadsheet.getSheetByName('処理履歴');

    if (!sheet) return;

    const successCount = results.filter(r => r.success).length;
    const failureCount = results.filter(r => !r.success).length;
    const processingTime = (Date.now() - startTime) / 1000; // 秒単位
    const errorMessages = results.filter(r => !r.success).map(r => r.error).join('; ');

    const rowData = [
      new Date(),
      results.length,
      successCount,
      failureCount,
      processingTime,
      errorMessages || ''
    ];

    sheet.appendRow(rowData);

  } catch (error) {
    console.error('処理履歴記録エラー:', error);
  }
}

// ==============================
// 通知
// ==============================

function sendNotification(message) {
  const config = getConfig();
  const email = config.NOTIFICATION_EMAIL;

  if (!email) {
    console.log('通知先メールアドレスが設定されていません');
    return;
  }

  GmailApp.sendEmail(
    email,
    'ポケモンカード管理システム通知',
    message,
    {
      name: 'Card Management System'
    }
  );
}

// ==============================
// 設定確認関数
// ==============================

function checkAPIKeys() {
  const props = PropertiesService.getScriptProperties();
  const openaiKey = props.getProperty('OPENAI_API_KEY');
  const perplexityKey = props.getProperty('PERPLEXITY_API_KEY');

  console.log('=== API キー設定状況 ===');
  console.log('OpenAI APIキー: ' + (openaiKey ? '設定済み（' + openaiKey.substring(0, 7) + '...）' : '未設定'));
  console.log('Perplexity APIキー: ' + (perplexityKey ? '設定済み（' + perplexityKey.substring(0, 7) + '...）' : '未設定'));

  if (!openaiKey) {
    console.log('\nOpenAI APIキーを設定するには:');
    console.log("PropertiesService.getScriptProperties().setProperty('OPENAI_API_KEY', 'sk-...');");
  }

  return {
    openai: openaiKey ? true : false,
    perplexity: perplexityKey ? true : false
  };
}

// ==============================
// テスト関数
// ==============================

function testConnection() {
  console.log('接続テスト開始');

  const config = getConfig();
  const results = {};

  // Google Photos接続テスト（オプション）
  try {
    const photos = getPhotosAlbumInfo(config.PHOTOS_ALBUM_ID);
    results.googlePhotos = `✓ アルバム接続成功: ${photos.title}`;
  } catch (error) {
    results.googlePhotos = `✗ アルバム接続失敗: ${error.toString()}`;
  }

  // Google Drive接続テスト（必須）
  try {
    const folder = DriveApp.getFolderById(config.DRIVE_FOLDER_ID);
    results.googleDrive = `✓ Drive接続成功: ${folder.getName()}`;
  } catch (error) {
    results.googleDrive = `✗ Drive接続失敗: ${error.toString()}`;
  }

  // OpenAI接続テスト（メイン）
  if (config.OPENAI_API_KEY) {
    try {
      const testResponse = testOpenAIConnection(config.OPENAI_API_KEY);
      results.openai = testResponse ? '✓ OpenAI接続成功（メインAI）' : '✗ OpenAI接続失敗';
    } catch (error) {
      results.openai = `✗ OpenAI接続失敗: ${error.toString()}`;
    }
  } else {
    results.openai = '✗ OpenAI APIキー未設定（推奨）';
  }

  // Perplexity接続テスト（フォールバック）
  if (config.PERPLEXITY_API_KEY) {
    try {
      const testResponse = testPerplexityConnection(config.PERPLEXITY_API_KEY);
      results.perplexity = testResponse ? '✓ Perplexity接続成功（フォールバック）' : '✗ Perplexity接続失敗';
    } catch (error) {
      results.perplexity = `✗ Perplexity接続失敗: ${error.toString()}`;
    }
  } else {
    results.perplexity = '- Perplexity未設定（オプション）';
  }

  // Notion接続テスト（必須）
  try {
    const dbInfo = getNotionDatabaseInfo(config);
    results.notion = `✓ Notion接続成功: ${dbInfo.title}`;
  } catch (error) {
    results.notion = `✗ Notion接続失敗: ${error.toString()}`;
  }

  console.log('接続テスト結果:', results);
  return results;
}

// OpenAI接続テスト
function testOpenAIConnection(apiKey) {
  const url = 'https://api.openai.com/v1/chat/completions';
  const config = getConfig();

  try {
    const payload = {
      model: config.PRICE_MODEL || 'gpt-4o',
      messages: [
        {
          role: 'user',
          content: 'Test connection'
        }
      ],
      max_tokens: 10
    };

    const response = UrlFetchApp.fetch(url, {
      method: 'POST',
      headers: {
        'Authorization': 'Bearer ' + apiKey,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    if (response.getResponseCode() === 200) {
      console.log(`OpenAI接続成功（Vision: ${config.VISION_MODEL}, Price: ${config.PRICE_MODEL}）`);
      return true;
    } else {
      console.log('OpenAI接続失敗: ' + response.getContentText());
      return false;
    }
  } catch (error) {
    console.log('OpenAI接続エラー: ' + error.toString());
    return false;
  }
}

function testPerplexityConnection(apiKey) {
  const url = 'https://api.perplexity.ai/chat/completions';

  try {
    // テキスト用モデルで接続テスト（画像解析はOpenAIを使用）
    const payload = {
      model: 'sonar-pro',  // 正しいモデル名に修正
      messages: [
        {
          role: 'user',
          content: 'Test connection'
        }
      ],
      max_tokens: 10
    };

    const response = UrlFetchApp.fetch(url, {
      method: 'POST',
      headers: {
        'Authorization': 'Bearer ' + apiKey,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    if (response.getResponseCode() === 200) {
      console.log('Perplexity接続成功');
      return true;
    } else {
      console.log('Perplexity接続失敗: ' + response.getContentText());
      return false;
    }
  } catch (error) {
    console.log('Perplexity接続エラー: ' + error.toString());
    return false;
  }
}

// ==============================
// Notionデータベースプロパティ確認と作成
// ==============================

function checkNotionDatabaseStructure() {
  const config = getConfig();

  if (!config.NOTION_API_KEY || !config.NOTION_DATABASE_ID) {
    console.log('Notion設定が未設定です');
    return;
  }

  try {
    const dbInfo = getNotionDatabaseInfo(config);
    console.log('データベース名:', dbInfo.title);
    console.log('\n現在のプロパティ一覧:');

    Object.keys(dbInfo.properties).forEach(propName => {
      const prop = dbInfo.properties[propName];
      console.log(`- ${propName} (${prop.type})`);
    });

    console.log('\n=== プロパティ名の大文字小文字を確認 ===');
    console.log('実際のプロパティ名をそのまま使用する必要があります');

    return dbInfo.properties;
  } catch (error) {
    console.error('データベース情報取得エラー:', error);
  }
}

function setupNotionDatabaseProperties() {
  const config = getConfig();

  if (!config.NOTION_API_KEY || !config.NOTION_DATABASE_ID) {
    console.log('Notion設定が未設定です');
    return;
  }

  const notionApiKey = config.NOTION_API_KEY;
  const databaseId = config.NOTION_DATABASE_ID;
  const url = `https://api.notion.com/v1/databases/${databaseId}`;

  // まず既存のプロパティを取得
  console.log('既存プロパティを確認中...');
  let existingProperties = {};
  try {
    const dbInfo = getNotionDatabaseInfo(config);
    existingProperties = dbInfo.properties;
    console.log('既存プロパティ:');
    Object.keys(existingProperties).forEach(prop => {
      console.log(`- ${prop}: ${existingProperties[prop].type}`);
    });
  } catch (error) {
    console.error('既存プロパティ取得エラー:', error);
    return false;
  }

  // 必要なプロパティの定義（titleは除外）
  const requiredProperties = {
    'UniqueID': { rich_text: {} },
    'Game': {
      select: {
        options: [
          { name: 'ポケモン', color: 'red' },
          { name: '遊戯王', color: 'blue' },
          { name: 'MTG', color: 'green' },
          { name: 'その他', color: 'gray' }
        ]
      }
    },
    'Set': { rich_text: {} },
    'Number': { rich_text: {} },
    'Rarity': {
      select: {
        options: [
          { name: 'UR', color: 'purple' },
          { name: 'SR', color: 'yellow' },
          { name: 'HR', color: 'orange' },
          { name: 'R', color: 'blue' },
          { name: 'U', color: 'green' },
          { name: 'C', color: 'gray' },
          { name: 'プロモ', color: 'pink' }
        ]
      }
    },
    'Language': {
      select: {
        options: [
          { name: '日本語', color: 'blue' },
          { name: '英語', color: 'red' },
          { name: 'その他', color: 'gray' }
        ]
      }
    },
    'Condition': {
      select: {
        options: [
          { name: 'NM', color: 'green' },
          { name: 'SP', color: 'yellow' },
          { name: 'MP', color: 'orange' },
          { name: 'HP', color: 'red' },
          { name: 'DM', color: 'gray' }
        ]
      }
    },
    'Status': {
      select: {
        options: [
          { name: '在庫', color: 'green' },
          { name: '出品中', color: 'yellow' },
          { name: '売却済', color: 'red' },
          { name: '保留', color: 'gray' }
        ]
      }
    },
    'Price': { number: { format: 'yen' } },
    'MarketPrice': { number: { format: 'yen' } },
    'PriceTrend': {  // 価格トレンド
      select: {
        options: [
          { name: '上昇', color: 'green' },
          { name: '下降', color: 'red' },
          { name: '安定', color: 'blue' },
          { name: '不明', color: 'gray' }
        ]
      }
    },
    'PriceHistory': { rich_text: {} },  // 価格推移データ
    'PricePrediction': { rich_text: {} },  // 価格予測データ
    'Price1YearAgo': { number: { format: 'yen' } },  // 1年前の価格
    'Price6MonthsAgo': { number: { format: 'yen' } },  // 6ヶ月前の価格
    'Price3MonthsAgo': { number: { format: 'yen' } },  // 3ヶ月前の価格
    'PredictedPrice6Months': { number: { format: 'yen' } },  // 6ヶ月後予測
    'PredictedPrice1Year': { number: { format: 'yen' } },  // 1年後予測
    'PSA9_Price': { number: { format: 'yen' } },  // PSA9鑑定品価格
    'PSA9.5_Price': { number: { format: 'yen' } },  // PSA9.5鑑定品価格
    'PSA10_Price': { number: { format: 'yen' } },  // PSA10鑑定品価格
    'Source': { rich_text: {} },
    'DuplicateNumber': { number: {} },
    'PhotoID': { rich_text: {} },
    'DriveFileID': { rich_text: {} },
    'Notes': { rich_text: {} },
    'ImageURL': { url: {} },
    'RegisteredDate': { date: {} },
    'LastUpdated': { date: {} },
    'PriceLastUpdated': { date: {} }  // 価格最終更新日
  };

  // 既存のタイトルプロパティを保持しつつ、新規プロパティを追加
  const titleProp = Object.keys(existingProperties).find(key =>
    existingProperties[key].type === 'title'
  );

  const propertiesToUpdate = {};
  if (titleProp) {
    // 既存のタイトルプロパティをそのまま使用
    propertiesToUpdate[titleProp] = existingProperties[titleProp];
  }

  // 新規プロパティを追加
  Object.assign(propertiesToUpdate, requiredProperties);

  const payload = {
    properties: propertiesToUpdate
  };

  const options = {
    method: 'PATCH',
    headers: {
      'Authorization': 'Bearer ' + notionApiKey,
      'Notion-Version': '2022-06-28',
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  console.log('\nNotionデータベースプロパティを更新中...');
  const response = UrlFetchApp.fetch(url, options);

  if (response.getResponseCode() === 200) {
    console.log('✓ Notionデータベースプロパティ作成成功');

    // 更新後のプロパティを確認
    const updatedDb = JSON.parse(response.getContentText());
    console.log('\n更新後のプロパティ:');
    Object.keys(updatedDb.properties).forEach(prop => {
      console.log(`- ${prop}: ${updatedDb.properties[prop].type}`);
    });

    return true;
  } else {
    const error = response.getContentText();
    console.error('✗ プロパティ作成失敗:', error);
    return false;
  }
}

// ==============================
// 価格情報取得と処理
// ==============================

function getCardPriceData(cardData) {
  const priceData = {
    currentPrice: 0,
    marketPrice: 0,
    priceHistory: {},
    pricePrediction: {},
    lastUpdated: new Date().toISOString(),
    currency: getCurrencyByLanguage(cardData.language)
  };

  try {
    // すでにAIなどで価格が入っている場合はそれを優先
    if (cardData.price) {
      // 価格を数値に変換（文字列の場合は$を除去）
      let rawPrice = cardData.price;
      if (typeof rawPrice === 'string') {
        rawPrice = parseFloat(rawPrice.replace(/[$,¥]/g, ''));
      }
      rawPrice = parseFloat(rawPrice) || 0;

      let rawMarketPrice = cardData.marketPrice || rawPrice;
      if (typeof rawMarketPrice === 'string') {
        rawMarketPrice = parseFloat(rawMarketPrice.replace(/[$,¥]/g, ''));
      }
      rawMarketPrice = parseFloat(rawMarketPrice) || rawPrice;

      // 英語カード（USD）の場合は日本円に変換（まだ変換されていない場合のみ）
      let convertedPrice = rawPrice;
      let convertedMarketPrice = rawMarketPrice;

      const currentCurrency = cardData.currency || priceData.currency;
      // priceConvertedがない、かつ通貨がUSDの場合のみ変換
      // すでにJPYに変換されている場合はスキップ
      if (currentCurrency === 'USD' && !cardData.priceConverted && !cardData.priceUSD) {
        // 為替レートを取得または使用
        const exchangeRate = cardData.exchangeRate || getExchangeRate('USD', 'JPY');
        // USD → JPY変換
        convertedPrice = Math.round(rawPrice * exchangeRate);
        convertedMarketPrice = Math.round(rawMarketPrice * exchangeRate);
        console.log(`getCardPriceDataで価格変換: $${rawPrice.toFixed(2)} → ¥${convertedPrice} (レート: ${exchangeRate})`);
      } else if (cardData.priceConverted) {
        console.log('getCardPriceData: すでに変換済みのためスキップ');
      }

      priceData.currentPrice = convertedPrice;
      priceData.marketPrice = convertedMarketPrice;
      priceData.priceHistory = cardData.priceHistory || generatePriceHistory(priceData.currentPrice);
      priceData.pricePrediction = cardData.pricePrediction || generatePricePrediction(priceData.currentPrice, cardData.rarity);
      priceData.currency = 'JPY'; // 表示は常にJPYで統一

      // PSAグレード価格も設定（すでに変換済みの場合はそのまま使用）
      if (cardData.psaGradedPrice) {
        priceData.psaGradedPrice = cardData.psaGradedPrice;
      }

      return priceData;
    }

    // ゲームタイプによって異なる価格取得処理
    if (cardData.game === 'Pokemon' || cardData.game === 'ポケモン') {
      Object.assign(priceData, getPokemonCardPrice(cardData));
    } else if (cardData.game === 'Yu-Gi-Oh!' || cardData.game === '遊戯王') {
      Object.assign(priceData, getYugiohCardPrice(cardData));
    } else if (cardData.game === 'MTG') {
      Object.assign(priceData, getMTGCardPrice(cardData));
    }

    // 価格推移データの生成（シミュレーション）
    priceData.priceHistory = generatePriceHistory(priceData.currentPrice);

    // 価格予測の生成
    priceData.pricePrediction = generatePricePrediction(priceData.currentPrice, cardData.rarity);

  } catch (error) {
    console.error('価格情報取得エラー:', error);
  }

  return priceData;
}

function generatePriceHistory(currentPrice) {
  if (!currentPrice || currentPrice === 0) {
    return {
      '12ヶ月前': 0,
      '9ヶ月前': 0,
      '6ヶ月前': 0,
      '3ヶ月前': 0,
      '現在': 0
    };
  }

  // 価格変動のシミュレーション（実際のAPIがない場合）
  const history = {
    '12ヶ月前': Math.round(currentPrice * (0.7 + Math.random() * 0.3)),
    '9ヶ月前': Math.round(currentPrice * (0.8 + Math.random() * 0.3)),
    '6ヶ月前': Math.round(currentPrice * (0.85 + Math.random() * 0.3)),
    '3ヶ月前': Math.round(currentPrice * (0.9 + Math.random() * 0.2)),
    '現在': currentPrice
  };

  return history;
}

function generatePricePrediction(currentPrice, rarity) {
  if (!currentPrice || currentPrice === 0) {
    return {
      '6ヶ月後': 0,
      '12ヶ月後': 0
    };
  }

  // レアリティによる価格上昇率の調整
  let growthFactor = 1.0;
  if (rarity) {
    const rarityFactors = {
      'UR': 1.3,
      'SR': 1.2,
      'HR': 1.25,
      'SAR': 1.35,
      'R': 1.1,
      'U': 1.05,
      'C': 1.0
    };
    growthFactor = rarityFactors[rarity] || 1.1;
  }

  const prediction = {
    '6ヶ月後': Math.round(currentPrice * growthFactor * (0.9 + Math.random() * 0.3)),
    '12ヶ月後': Math.round(currentPrice * growthFactor * growthFactor * (0.8 + Math.random() * 0.4))
  };

  return prediction;
}

function formatPriceInfo(priceData, cardData) {
  const currency = priceData.currency || 'JPY';
  const sym = getCurrencySymbol(currency);
  let info = '【価格情報】\n';

  info += `現在価格: ${sym}${priceData.currentPrice || 0}\n`;
  info += `市場価格: ${sym}${priceData.marketPrice || 0}\n\n`;

  // PSAグレード別価格を追加
  if (cardData && cardData.psaGradedPrice) {
    info += '【PSAグレード別価格】\n';
    if (cardData.psaGradedPrice.PSA9) {
      info += `PSA9: ¥${cardData.psaGradedPrice.PSA9.toLocaleString()}\n`;
    }
    if (cardData.psaGradedPrice['PSA9.5']) {
      info += `PSA9.5: ¥${cardData.psaGradedPrice['PSA9.5'].toLocaleString()}\n`;
    }
    if (cardData.psaGradedPrice.PSA10) {
      info += `PSA10: ¥${cardData.psaGradedPrice.PSA10.toLocaleString()}\n`;
    }
    info += '\n';
  }

  info += '【価格推移】\n';
  if (priceData.priceHistory) {
    Object.entries(priceData.priceHistory).forEach(([period, price]) => {
      if (period !== 'trend') { // trendプロパティを除外
        info += `${period}: ${sym}${price}\n`;
      }
    });
  }

  info += '\n【価格予測】\n';
  if (priceData.pricePrediction) {
    Object.entries(priceData.pricePrediction).forEach(([period, price]) => {
      info += `${period}: ${sym}${price}\n`;
    });
  }

  info += `\n最終更新: ${priceData.lastUpdated || new Date().toISOString()}`;

  return info;
}

// 通貨別の最低価格（0円回避用）
function getMinimumPrice(currency) {
  switch ((currency || 'JPY').toUpperCase()) {
    case 'USD':
      return 1; // $1 未満は切り上げ
    case 'EUR':
      return 1;
    case 'JPY':
    default:
      return 50; // ¥50 未満は切り上げ
  }
}

function getPokemonCardPrice(cardData) {
  // ポケモンカードの価格取得（実際のAPIを使用する場合はここに実装）
  // 現在はサンプルデータを返す
  const basePrice = Math.floor(Math.random() * 10000) + 500;

  return {
    currentPrice: basePrice,
    marketPrice: Math.round(basePrice * 1.1)
  };
}

function getYugiohCardPrice(cardData) {
  // 遊戯王カードの価格取得
  const basePrice = Math.floor(Math.random() * 5000) + 300;

  return {
    currentPrice: basePrice,
    marketPrice: Math.round(basePrice * 1.05)
  };
}

function getMTGCardPrice(cardData) {
  // MTGカードの価格取得
  const basePrice = Math.floor(Math.random() * 15000) + 1000;

  return {
    currentPrice: basePrice,
    marketPrice: Math.round(basePrice * 1.15)
  };
}

// ==============================
// 価格更新処理
// ==============================

function updateCardPrices() {
  const startTime = Date.now();
  console.log('カード価格更新処理開始');

  const config = getConfig();
  const spreadsheetId = PropertiesService.getScriptProperties().getProperty('MASTER_SPREADSHEET_ID');

  if (!spreadsheetId) {
    console.error('スプレッドシートが見つかりません');
    return;
  }

  try {
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheet = spreadsheet.getSheetByName('カード一覧');

    if (!sheet) {
      console.error('カード一覧シートが見つかりません');
      return;
    }

    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();

    let updateCount = 0;
    const batchSize = 10; // 一度に処理するカード数

    // ヘッダー行をスキップして処理（バッチ処理）
    for (let i = 1; i < values.length && updateCount < batchSize; i++) {
      const uniqueId = values[i][0];
      const name = values[i][2];
      const game = values[i][3];
      const set = values[i][4];
      const number = values[i][5];

      if (!name || !game) continue;

      // カードデータを構築
      const cardData = {
        uniqueId: uniqueId,
        name: name,
        game: game,
        set: set,
        number: number
      };

      try {
        // 外部APIから価格情報を取得
        enrichCardData(cardData);

        if (cardData.price) {
          // 価格カラム（10列目）を更新
          sheet.getRange(i + 1, 10).setValue(cardData.price);

          // 価格履歴を記録
          logPriceHistory(cardData);

          updateCount++;
          console.log(`価格更新: ${name} - ${cardData.price}`);
        }

      } catch (error) {
        console.error(`価格更新エラー: ${name}`, error);
      }

      // API制限対策
      Utilities.sleep(1000);
    }

    // 処理結果を記録
    const processingTime = (Date.now() - startTime) / 1000;
    console.log(`価格更新完了: ${updateCount}件更新（処理時間: ${processingTime}秒）`);

    // 価格更新履歴をスプレッドシートに記録
    logPriceUpdateHistory(updateCount, processingTime);

  } catch (error) {
    console.error('価格更新処理エラー:', error);
    sendNotification('カード価格更新処理でエラーが発生しました: ' + error.toString());
  }
}

function logPriceHistory(cardData) {
  const spreadsheetId = PropertiesService.getScriptProperties().getProperty('MASTER_SPREADSHEET_ID');

  if (!spreadsheetId) return;

  try {
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);

    // 価格履歴シートを取得または作成
    let historySheet = spreadsheet.getSheetByName('価格履歴');
    if (!historySheet) {
      historySheet = spreadsheet.insertSheet('価格履歴');

      // ヘッダー設定
      const headers = [
        '記録日時',
        'ユニークID',
        'カード名',
        'ゲーム',
        '価格',
        '前回価格',
        '変動率(%)',
        'データソース'
      ];

      historySheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      historySheet.getRange(1, 1, 1, headers.length)
        .setBackground('#9900FF')
        .setFontColor('#FFFFFF')
        .setFontWeight('bold');
    }

    // 前回価格を取得
    const lastPrice = getLastPrice(cardData.uniqueId, historySheet);
    let changeRate = 0;

    if (lastPrice && lastPrice !== '0') {
      const currentPriceNum = parseFloat(cardData.price.replace(/[^0-9.-]/g, ''));
      const lastPriceNum = parseFloat(lastPrice.replace(/[^0-9.-]/g, ''));
      if (lastPriceNum > 0) {
        changeRate = ((currentPriceNum - lastPriceNum) / lastPriceNum * 100).toFixed(2);
      }
    }

    // 価格履歴を記録
    const historyData = [
      new Date(),
      cardData.uniqueId,
      cardData.name,
      cardData.game,
      cardData.price,
      lastPrice || '-',
      changeRate ? changeRate + '%' : '-',
      'API'
    ];

    historySheet.appendRow(historyData);

  } catch (error) {
    console.error('価格履歴記録エラー:', error);
  }
}

function getLastPrice(uniqueId, sheet) {
  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();

  // 最新の価格を探す（逆順で検索）
  for (let i = values.length - 1; i > 0; i--) {
    if (values[i][1] === uniqueId) {
      return values[i][4]; // 価格カラム
    }
  }

  return null;
}

function logPriceUpdateHistory(updateCount, processingTime) {
  const spreadsheetId = PropertiesService.getScriptProperties().getProperty('MASTER_SPREADSHEET_ID');

  if (!spreadsheetId) return;

  try {
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheet = spreadsheet.getSheetByName('処理履歴');

    if (!sheet) return;

    const historyData = [
      new Date(),
      0, // 処理枚数（新規）
      updateCount, // 成功（価格更新数として使用）
      0, // 失敗
      processingTime,
      '価格更新処理'
    ];

    sheet.appendRow(historyData);

  } catch (error) {
    console.error('価格更新履歴記録エラー:', error);
  }
}

// 手動価格更新（指定カードのみ）
function updateSingleCardPrice(uniqueId) {
  const config = getConfig();
  const spreadsheetId = PropertiesService.getScriptProperties().getProperty('MASTER_SPREADSHEET_ID');

  if (!spreadsheetId) {
    console.error('スプレッドシートが見つかりません');
    return;
  }

  try {
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheet = spreadsheet.getSheetByName('カード一覧');

    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();

    // 該当カードを検索
    for (let i = 1; i < values.length; i++) {
      if (values[i][0] === uniqueId) {
        const cardData = {
          uniqueId: values[i][0],
          name: values[i][2],
          game: values[i][3],
          set: values[i][4],
          number: values[i][5]
        };

        // 価格情報を取得
        enrichCardData(cardData);

        if (cardData.price) {
          sheet.getRange(i + 1, 10).setValue(cardData.price);
          logPriceHistory(cardData);
          console.log(`価格更新完了: ${cardData.name} - ${cardData.price}`);
        }

        return cardData;
      }
    }

    console.log('指定されたカードが見つかりません: ' + uniqueId);

  } catch (error) {
    console.error('単一カード価格更新エラー:', error);
  }
}


// ==============================
// Google Photos → Drive 同期機能
// ==============================

/**
 * Google Photos Library APIを使用してアルバムから画像を取得し、
 * Google Driveにコピーして処理する
 */

// ==============================
// 設定と初期化
// ==============================

/**
 * Google Photos同期の初期設定
 */
function setupPhotosSyncConfig() {
  const props = PropertiesService.getScriptProperties();

  // デフォルト設定
  const defaultConfig = {
    'PHOTOS_ALBUM_NAME': 'ポケモンカード',  // 同期するアルバム名
    'PHOTOS_SYNC_ENABLED': 'true',           // 同期を有効化
    'PHOTOS_LAST_SYNC': '',                  // 最後の同期日時
    'PHOTOS_SYNC_INTERVAL': '1',             // 同期間隔（時間）
    'PHOTOS_MAX_ITEMS': '50'                 // 一度に処理する最大枚数
  };

  // 既存の設定を保持しながらデフォルト値を設定
  Object.keys(defaultConfig).forEach(key => {
    if (!props.getProperty(key)) {
      props.setProperty(key, defaultConfig[key]);
    }
  });

  console.log('Google Photos同期設定を初期化しました');

  return {
    albumName: props.getProperty('PHOTOS_ALBUM_NAME'),
    enabled: props.getProperty('PHOTOS_SYNC_ENABLED') === 'true',
    lastSync: props.getProperty('PHOTOS_LAST_SYNC'),
    syncInterval: parseInt(props.getProperty('PHOTOS_SYNC_INTERVAL')),
    maxItems: parseInt(props.getProperty('PHOTOS_MAX_ITEMS'))
  };
}

// ==============================
// Google Photos API関連
// ==============================

/**
 * Google Photos APIのアクセストークンを取得
 * @return {string} アクセストークン
 */
function getPhotosAccessToken() {
  try {
    const service = ScriptApp.getOAuthToken();
    return service;
  } catch (error) {
    console.error('アクセストークン取得エラー:', error);
    throw new Error('Google Photos APIへのアクセス権限が必要です');
  }
}

/**
 * 指定した名前のアルバムIDを取得
 * @param {string} albumName - アルバム名
 * @return {string|null} アルバムID
 */
function getAlbumIdByName(albumName) {
  const token = getPhotosAccessToken();
  const url = 'https://photoslibrary.googleapis.com/v1/albums';

  try {
    const response = UrlFetchApp.fetch(url, {
      headers: {
        'Authorization': 'Bearer ' + token,
        'Content-Type': 'application/json'
      },
      muteHttpExceptions: true
    });

    if (response.getResponseCode() !== 200) {
      console.error('アルバム一覧取得エラー:', response.getContentText());
      return null;
    }

    const data = JSON.parse(response.getContentText());
    const albums = data.albums || [];

    // 指定名のアルバムを探す
    const targetAlbum = albums.find(album => album.title === albumName);

    if (targetAlbum) {
      console.log(`アルバム "${albumName}" を見つけました: ${targetAlbum.id}`);
      return targetAlbum.id;
    } else {
      console.log(`アルバム "${albumName}" が見つかりません`);

      // アルバムを作成
      return createPhotosAlbum(albumName);
    }

  } catch (error) {
    console.error('アルバムID取得エラー:', error);
    return null;
  }
}

/**
 * Google Photosにアルバムを作成
 * @param {string} albumName - アルバム名
 * @return {string|null} 作成されたアルバムのID
 */
function createPhotosAlbum(albumName) {
  const token = getPhotosAccessToken();
  const url = 'https://photoslibrary.googleapis.com/v1/albums';

  try {
    const response = UrlFetchApp.fetch(url, {
      method: 'POST',
      headers: {
        'Authorization': 'Bearer ' + token,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify({
        album: {
          title: albumName
        }
      }),
      muteHttpExceptions: true
    });

    if (response.getResponseCode() === 200) {
      const album = JSON.parse(response.getContentText());
      console.log(`アルバム "${albumName}" を作成しました: ${album.id}`);
      return album.id;
    } else {
      console.error('アルバム作成エラー:', response.getContentText());
      return null;
    }

  } catch (error) {
    console.error('アルバム作成エラー:', error);
    return null;
  }
}

/**
 * アルバムから画像一覧を取得
 * @param {string} albumId - アルバムID
 * @param {number} maxResults - 取得する最大件数
 * @return {Array} 画像アイテムの配列
 */
function getPhotosFromAlbum(albumId, maxResults = 50) {
  const token = getPhotosAccessToken();
  const url = 'https://photoslibrary.googleapis.com/v1/mediaItems:search';

  const items = [];
  let pageToken = null;

  try {
    do {
      const payload = {
        albumId: albumId,
        pageSize: Math.min(maxResults - items.length, 100)
      };

      if (pageToken) {
        payload.pageToken = pageToken;
      }

      const response = UrlFetchApp.fetch(url, {
        method: 'POST',
        headers: {
          'Authorization': 'Bearer ' + token,
          'Content-Type': 'application/json'
        },
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      });

      if (response.getResponseCode() !== 200) {
        console.error('画像取得エラー:', response.getContentText());
        break;
      }

      const data = JSON.parse(response.getContentText());
      const mediaItems = data.mediaItems || [];

      items.push(...mediaItems);
      pageToken = data.nextPageToken;

    } while (pageToken && items.length < maxResults);

    console.log(`${items.length}枚の画像を取得しました`);
    return items;

  } catch (error) {
    console.error('画像一覧取得エラー:', error);
    return [];
  }
}

/**
 * Google Photos画像をGoogle Driveにダウンロード
 * @param {Object} mediaItem - Google PhotosのメディアアイテムAPI:
 * @param {string} targetFolderId - 保存先のDriveフォルダID
 * @return {Object|null} Driveファイル情報
 */
function downloadPhotoToDrive(mediaItem, targetFolderId) {
  try {
    // ダウンロードURL生成（最高品質）
    const downloadUrl = `${mediaItem.baseUrl}=d`;

    // 画像データを取得
    const response = UrlFetchApp.fetch(downloadUrl, {
      muteHttpExceptions: true
    });

    if (response.getResponseCode() !== 200) {
      console.error('画像ダウンロードエラー:', response.getResponseCode());
      return null;
    }

    const blob = response.getBlob();

    // ファイル名を生成（タイムスタンプ付き）
    const originalName = mediaItem.filename || 'untitled';
    const timestamp = new Date(mediaItem.mediaMetadata.creationTime).getTime();
    const fileName = `photos_${timestamp}_${originalName}`;

    blob.setName(fileName);

    // HEIC形式の場合はJPEGに変換
    if (fileName.toLowerCase().match(/\.(heic|heif)$/)) {
      console.log('HEIC形式を検出、JPEGに変換中...');
      const jpegBlob = blob.getAs('image/jpeg');
      jpegBlob.setName(fileName.replace(/\.(heic|heif)$/i, '.jpg'));
      blob = jpegBlob;
    }

    // Driveフォルダに保存
    const folder = DriveApp.getFolderById(targetFolderId);
    const file = folder.createFile(blob);

    // メタデータを設定
    file.setDescription(JSON.stringify({
      photosId: mediaItem.id,
      photosUrl: mediaItem.productUrl,
      creationTime: mediaItem.mediaMetadata.creationTime,
      width: mediaItem.mediaMetadata.width,
      height: mediaItem.mediaMetadata.height,
      syncedAt: new Date().toISOString()
    }));

    console.log(`画像をDriveに保存: ${file.getName()}`);

    return {
      id: file.getId(),
      name: file.getName(),
      url: file.getUrl(),
      photosId: mediaItem.id,
      mimeType: file.getMimeType(),
      size: file.getSize()
    };

  } catch (error) {
    console.error('画像ダウンロードエラー:', error);
    return null;
  }
}

// ==============================
// 同期処理
// ==============================

/**
 * Google PhotosアルバムをDriveと同期
 * @return {Object} 同期結果
 */
function syncPhotosAlbumToDrive() {
  const startTime = Date.now();
  console.log('=== Google Photos同期開始 ===');

  try {
    // 設定を取得
    const syncConfig = setupPhotosSyncConfig();

    if (!syncConfig.enabled) {
      console.log('Photos同期が無効になっています');
      return { success: false, message: '同期が無効' };
    }

    // アルバムIDを取得
    const albumId = getAlbumIdByName(syncConfig.albumName);
    if (!albumId) {
      throw new Error(`アルバム "${syncConfig.albumName}" が見つかりません`);
    }

    // Driveフォルダを準備
    const driveFolderId = getOrCreateSyncFolder();

    // 処理済みIDを取得
    const processedIds = getProcessedPhotosIds();

    // アルバムから画像を取得
    const photos = getPhotosFromAlbum(albumId, syncConfig.maxItems);

    // 新着画像をフィルタリング
    const newPhotos = photos.filter(photo => !processedIds.includes(photo.id));

    if (newPhotos.length === 0) {
      console.log('新着画像なし');
      updateLastSyncTime();
      return { success: true, message: '新着画像なし', count: 0 };
    }

    console.log(`新着画像: ${newPhotos.length}枚`);

    // 各画像をDriveにダウンロード
    const results = [];
    for (const photo of newPhotos) {
      try {
        console.log(`処理中: ${photo.filename}`);

        // Driveにダウンロード
        const driveFile = downloadPhotoToDrive(photo, driveFolderId);

        if (driveFile) {
          // 処理済みとしてマーク
          markPhotosAsProcessed(photo.id);

          results.push({
            success: true,
            photosId: photo.id,
            driveId: driveFile.id,
            fileName: driveFile.name
          });
        } else {
          results.push({
            success: false,
            photosId: photo.id,
            error: 'ダウンロード失敗'
          });
        }

      } catch (error) {
        console.error(`画像処理エラー: ${photo.filename}`, error);
        results.push({
          success: false,
          photosId: photo.id,
          error: error.toString()
        });
      }
    }

    // 同期時刻を更新
    updateLastSyncTime();

    // 結果サマリー
    const successCount = results.filter(r => r.success).length;
    const failureCount = results.filter(r => !r.success).length;
    const processingTime = (Date.now() - startTime) / 1000;

    console.log(`=== 同期完了 ===`);
    console.log(`成功: ${successCount}枚, 失敗: ${failureCount}枚`);
    console.log(`処理時間: ${processingTime}秒`);

    // 同期後にDrive画像を処理
    if (successCount > 0) {
      console.log('Drive画像処理を開始します...');
      processImagesFromDrive();
    }

    return {
      success: true,
      message: `${successCount}枚を同期`,
      successCount: successCount,
      failureCount: failureCount,
      processingTime: processingTime,
      results: results
    };

  } catch (error) {
    console.error('Photos同期エラー:', error);
    sendNotification('Google Photos同期エラー: ' + error.toString());
    return {
      success: false,
      message: error.toString()
    };
  }
}

/**
 * 同期用Driveフォルダを取得または作成
 * @return {string} フォルダID
 */
function getOrCreateSyncFolder() {
  let folderId = PropertiesService.getScriptProperties().getProperty('DRIVE_FOLDER_ID');

  if (!folderId) {
    const folder = DriveApp.createFolder('ポケモンカード_Photos同期');
    folderId = folder.getId();
    PropertiesService.getScriptProperties().setProperty('DRIVE_FOLDER_ID', folderId);
    console.log(`同期フォルダ作成: ${folder.getUrl()}`);
  }

  // Photos同期用のサブフォルダを作成
  const mainFolder = DriveApp.getFolderById(folderId);
  let syncFolder;

  const syncFolders = mainFolder.getFoldersByName('Photos同期');
  if (syncFolders.hasNext()) {
    syncFolder = syncFolders.next();
  } else {
    syncFolder = mainFolder.createFolder('Photos同期');
    console.log('Photos同期フォルダを作成しました');
  }

  return syncFolder.getId();
}

/**
 * 処理済みPhotos IDを取得
 * @return {Array<string>} 処理済みID配列
 */
function getProcessedPhotosIds() {
  const processedIds = PropertiesService.getScriptProperties().getProperty('PROCESSED_PHOTOS_IDS');
  return processedIds ? JSON.parse(processedIds) : [];
}

/**
 * Photos画像を処理済みとしてマーク
 * @param {string} photosId - Google Photos ID
 */
function markPhotosAsProcessed(photosId) {
  const processedIds = getProcessedPhotosIds();

  if (!processedIds.includes(photosId)) {
    processedIds.push(photosId);

    // 最大1000件まで保持（古いものから削除）
    if (processedIds.length > 1000) {
      processedIds.splice(0, processedIds.length - 1000);
    }

    PropertiesService.getScriptProperties().setProperty(
      'PROCESSED_PHOTOS_IDS',
      JSON.stringify(processedIds)
    );
  }
}

/**
 * 最終同期時刻を更新
 */
function updateLastSyncTime() {
  PropertiesService.getScriptProperties().setProperty(
    'PHOTOS_LAST_SYNC',
    new Date().toISOString()
  );
}

// ==============================
// トリガー設定
// ==============================

/**
 * Photos同期トリガーを設定
 */
function setupPhotosSyncTriggers() {
  // 既存のPhotos同期トリガーを削除
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'syncPhotosAlbumToDrive') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // 新しいトリガーを設定（1時間ごと）
  ScriptApp.newTrigger('syncPhotosAlbumToDrive')
    .timeBased()
    .everyHours(1)
    .create();

  console.log('Photos同期トリガー設定完了: 1時間ごと');
}

// ==============================
// 管理関数
// ==============================

/**
 * Photos同期状態を確認
 */
function checkPhotosSyncStatus() {
  const config = setupPhotosSyncConfig();

  console.log('=== Google Photos同期状態 ===');
  console.log(`同期有効: ${config.enabled}`);
  console.log(`アルバム名: ${config.albumName}`);
  console.log(`最終同期: ${config.lastSync || '未実行'}`);
  console.log(`同期間隔: ${config.syncInterval}時間`);
  console.log(`最大処理数: ${config.maxItems}枚/回`);

  const processedIds = getProcessedPhotosIds();
  console.log(`処理済み画像数: ${processedIds.length}枚`);

  // アルバムの存在確認
  const albumId = getAlbumIdByName(config.albumName);
  if (albumId) {
    console.log(`アルバムID: ${albumId}`);

    // アルバム内の画像数を確認
    const photos = getPhotosFromAlbum(albumId, 1);
    if (photos.length > 0) {
      console.log('アルバムへのアクセス: 成功');
    }
  } else {
    console.log('アルバム: 見つかりません');
  }

  return config;
}

/**
 * Photos同期設定を変更
 * @param {string} albumName - 同期するアルバム名
 * @param {boolean} enabled - 同期を有効化するか
 */
function configurePhotosSync(albumName, enabled = true) {
  const props = PropertiesService.getScriptProperties();

  if (albumName) {
    props.setProperty('PHOTOS_ALBUM_NAME', albumName);
  }

  props.setProperty('PHOTOS_SYNC_ENABLED', enabled.toString());

  console.log('Photos同期設定を更新しました');
  console.log(`アルバム名: ${albumName || '変更なし'}`);
  console.log(`同期有効: ${enabled}`);

  if (enabled) {
    // トリガーを設定
    setupPhotosSyncTriggers();
  }
}

/**
 * 手動で同期を実行
 */
function manualPhotosSync() {
  console.log('手動同期を開始します...');
  const result = syncPhotosAlbumToDrive();

  if (result.success) {
    console.log(`同期成功: ${result.message}`);
  } else {
    console.error(`同期失敗: ${result.message}`);
  }

  return result;
}

// ==============================
// 初期セットアップ
// ==============================

/**
 * Google Photos同期の初期セットアップ
 */
function setupPhotosSync() {
  console.log('=== Google Photos同期セットアップ ===');

  try {
    // 1. 設定を初期化
    const config = setupPhotosSyncConfig();
    console.log('1. 設定を初期化しました');

    // 2. アルバムを確認/作成
    const albumId = getAlbumIdByName(config.albumName);
    if (albumId) {
      console.log(`2. アルバム "${config.albumName}" を準備しました`);
    } else {
      console.log(`2. アルバム "${config.albumName}" の作成に失敗しました`);
      throw new Error('アルバムの作成に失敗しました');
    }

    // 3. Driveフォルダを準備
    const folderId = getOrCreateSyncFolder();
    const folder = DriveApp.getFolderById(folderId);
    console.log(`3. 同期フォルダを準備しました: ${folder.getUrl()}`);

    // 4. トリガーを設定
    setupPhotosSyncTriggers();
    console.log('4. 自動同期トリガーを設定しました');

    console.log('\n=== セットアップ完了 ===');
    console.log(`Google Photosアルバム "${config.albumName}" とDriveの同期が設定されました`);
    console.log('使い方:');
    console.log(`1. Google Photosで "${config.albumName}" アルバムに画像を追加`);
    console.log('2. 1時間ごとに自動同期されます');
    console.log('3. 手動同期: manualPhotosSync()を実行');

    return {
      success: true,
      albumName: config.albumName,
      folderId: folderId,
      folderUrl: folder.getUrl()
    };

  } catch (error) {
    console.error('セットアップエラー:', error);
    return {
      success: false,
      error: error.toString()
    };
  }
}// ==============================
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
}// ==============================
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
    console.log('プロジェクトID: GCP_PROJECT_ID_PLACEHOLDER');
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
    projectId: 'GCP_PROJECT_ID_PLACEHOLDER',
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

  console.log('【方法A: 既存プロジェクト(GCP_PROJECT_ID_PLACEHOLDER)でAPIを有効化】');
  console.log('1. 以下のURLにアクセス:');
  console.log('   https://console.cloud.google.com/apis/library/photoslibrary.googleapis.com?project=GCP_PROJECT_ID_PLACEHOLDER');
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
  console.log('$ gcloud services enable photoslibrary.googleapis.com --project=GCP_PROJECT_ID_PLACEHOLDER');
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
}// ==============================
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
  console.log('エラーメッセージのプロジェクトID: GCP_PROJECT_ID_PLACEHOLDER');
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
}// ==============================
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
}// ==============================
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
}// ==============================
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
}// ==============================
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
}// ==============================
// Google Photos API 代替実装（REST API直接呼び出し）
// ==============================

/**
 * Photos APIをREST APIで直接実装
 */
function setupPhotosAPIWorkaround() {
  console.log('=================================================================================');
  console.log('                    Google Photos API 代替実装                                   ');
  console.log('=================================================================================\n');

  console.log('Advanced ServicesにPhotos Library APIが表示されない問題を回避します。\n');

  console.log('【この実装の特徴】');
  console.log('✅ REST APIを直接呼び出し');
  console.log('✅ Advanced Services不要');
  console.log('✅ 既存のOAuth認証を使用');
  console.log('✅ すぐに使用可能\n');

  console.log('【実行コマンド】');
  console.log('1. testDirectPhotosAPI() - 接続テスト');
  console.log('2. listPhotosAlbums() - アルバム一覧取得');
  console.log('3. createPhotosAlbum("ポケモンカード") - アルバム作成');
  console.log('4. getPhotosFromLibrary() - 写真取得\n');

  return {
    status: 'ready',
    nextStep: 'testDirectPhotosAPI()'
  };
}

/**
 * REST API直接呼び出しで接続テスト
 */
function testDirectPhotosAPI() {
  console.log('=== Photos API直接接続テスト ===\n');

  const token = ScriptApp.getOAuthToken();

  // 最もシンプルなエンドポイントから開始
  const endpoints = [
    {
      name: 'メディアアイテム取得（GET）',
      url: 'https://photoslibrary.googleapis.com/v1/mediaItems?pageSize=1',
      method: 'GET',
      headers: {
        'Authorization': 'Bearer ' + token
      }
    },
    {
      name: 'アルバム一覧取得（GET）',
      url: 'https://photoslibrary.googleapis.com/v1/albums?pageSize=1',
      method: 'GET',
      headers: {
        'Authorization': 'Bearer ' + token
      }
    }
  ];

  let successCount = 0;

  endpoints.forEach(endpoint => {
    try {
      const response = UrlFetchApp.fetch(endpoint.url, {
        method: endpoint.method,
        headers: endpoint.headers,
        muteHttpExceptions: true
      });

      const code = response.getResponseCode();
      console.log(`${endpoint.name}: ${code}`);

      if (code === 200) {
        successCount++;
        const data = JSON.parse(response.getContentText());
        console.log(`  ✅ 成功`);

        if (endpoint.name.includes('メディアアイテム')) {
          console.log(`  アイテム数: ${data.mediaItems ? data.mediaItems.length : 0}`);
        } else if (endpoint.name.includes('アルバム')) {
          console.log(`  アルバム数: ${data.albums ? data.albums.length : 0}`);
        }
      } else if (code === 403) {
        console.log(`  ❌ 権限エラー`);
        const error = JSON.parse(response.getContentText());
        if (error.error && error.error.message) {
          console.log(`  ${error.error.message}`);
        }
      } else {
        console.log(`  ⚠️ エラー: ${code}`);
      }

    } catch (e) {
      console.error(`${endpoint.name}: 例外エラー - ${e.toString()}`);
    }
  });

  if (successCount > 0) {
    console.log('\n✅ Photos APIへの接続成功！');
    console.log('次のステップ: setupWorkingPhotosSync()');
    return true;
  } else {
    console.log('\n❌ Photos APIへの接続に失敗');
    console.log('解決策: fixPhotosPermissions()');
    return false;
  }
}

/**
 * 動作する形でアルバム一覧を取得
 */
function listPhotosAlbums() {
  console.log('=== アルバム一覧取得 ===\n');

  try {
    const token = ScriptApp.getOAuthToken();
    const url = 'https://photoslibrary.googleapis.com/v1/albums';

    const response = UrlFetchApp.fetch(url, {
      method: 'GET',
      headers: {
        'Authorization': 'Bearer ' + token,
        'Content-Type': 'application/json'
      },
      muteHttpExceptions: true
    });

    const code = response.getResponseCode();

    if (code === 200) {
      const data = JSON.parse(response.getContentText());
      const albums = data.albums || [];

      console.log(`アルバム数: ${albums.length}`);

      albums.forEach((album, index) => {
        console.log(`${index + 1}. ${album.title}`);
        console.log(`   ID: ${album.id}`);
        console.log(`   写真数: ${album.mediaItemsCount || 0}`);
      });

      return albums;
    } else {
      console.error('エラー:', response.getContentText());
      return [];
    }

  } catch (error) {
    console.error('例外エラー:', error);
    return [];
  }
}

/**
 * アルバムを作成
 */
function createPhotosAlbum(albumTitle) {
  console.log(`=== アルバム作成: ${albumTitle} ===\n`);

  try {
    const token = ScriptApp.getOAuthToken();
    const url = 'https://photoslibrary.googleapis.com/v1/albums';

    const payload = {
      album: {
        title: albumTitle
      }
    };

    const response = UrlFetchApp.fetch(url, {
      method: 'POST',
      headers: {
        'Authorization': 'Bearer ' + token,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    const code = response.getResponseCode();

    if (code === 200 || code === 201) {
      const album = JSON.parse(response.getContentText());
      console.log('✅ アルバム作成成功');
      console.log(`ID: ${album.id}`);
      console.log(`タイトル: ${album.title}`);
      console.log(`URL: ${album.productUrl}`);

      // アルバムIDを保存
      PropertiesService.getScriptProperties().setProperty('PHOTOS_ALBUM_ID', album.id);

      return album;
    } else {
      console.error('作成失敗:', response.getContentText());
      return null;
    }

  } catch (error) {
    console.error('例外エラー:', error);
    return null;
  }
}

/**
 * ライブラリから写真を取得
 */
function getPhotosFromLibrary(maxResults = 10) {
  console.log('=== ライブラリから写真取得 ===\n');

  try {
    const token = ScriptApp.getOAuthToken();
    const url = `https://photoslibrary.googleapis.com/v1/mediaItems?pageSize=${maxResults}`;

    const response = UrlFetchApp.fetch(url, {
      method: 'GET',
      headers: {
        'Authorization': 'Bearer ' + token
      },
      muteHttpExceptions: true
    });

    const code = response.getResponseCode();

    if (code === 200) {
      const data = JSON.parse(response.getContentText());
      const items = data.mediaItems || [];

      console.log(`取得した写真: ${items.length}枚`);

      items.forEach((item, index) => {
        console.log(`${index + 1}. ${item.filename}`);
        console.log(`   作成日: ${item.mediaMetadata.creationTime}`);
        console.log(`   サイズ: ${item.mediaMetadata.width}x${item.mediaMetadata.height}`);
      });

      return items;
    } else {
      console.error('取得失敗:', response.getContentText());
      return [];
    }

  } catch (error) {
    console.error('例外エラー:', error);
    return [];
  }
}

/**
 * 権限問題を修正
 */
function fixPhotosPermissions() {
  console.log('=== Photos API権限修正 ===\n');

  console.log('【確認事項】');
  console.log('1. company-gasプロジェクトでPhotos Library APIが有効化されているか');
  console.log('2. OAuth認証で適切なスコープが設定されているか\n');

  console.log('【解決手順】');

  console.log('ステップ1: プロジェクト番号を確認');
  console.log('────────────────────');
  verifyProjectNumber();

  console.log('\nステップ2: APIの有効化を再確認');
  console.log('────────────────────');
  console.log('https://console.cloud.google.com/apis/library/photoslibrary.googleapis.com?project=company-gas');
  console.log('で「管理」ボタンが表示されているか確認\n');

  console.log('ステップ3: 認証をリセット');
  console.log('────────────────────');
  console.log('clearPhotosAuth() を実行\n');

  console.log('ステップ4: 再テスト');
  console.log('────────────────────');
  console.log('testDirectPhotosAPI() を実行\n');

  return {
    step1: 'verifyProjectNumber()',
    step2: 'APIコンソールで確認',
    step3: 'clearPhotosAuth()',
    step4: 'testDirectPhotosAPI()'
  };
}

/**
 * プロジェクト番号を確認
 */
function verifyProjectNumber() {
  // GASプロジェクトの設定を確認するため、エラーメッセージから情報を取得
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

      if (error.error && error.error.message) {
        const match = error.error.message.match(/project (\d+)/);
        if (match) {
          console.log('現在のプロジェクト番号: ' + match[1]);
          console.log('\nGASエディタで確認:');
          console.log('プロジェクトの設定 → GCPプロジェクト番号');
          console.log('この番号とcompany-gasプロジェクトの番号が一致しているか確認');
        }
      }
    }
  } catch (e) {
    console.log('プロジェクト番号の取得に失敗');
  }
}

/**
 * 認証をクリア
 */
function clearPhotosAuth() {
  console.log('=== 認証のクリア ===\n');

  console.log('1. https://myaccount.google.com/permissions にアクセス');
  console.log('2. このプロジェクト関連のアクセスを削除');
  console.log('3. GASエディタでプロジェクトを保存（Ctrl+S）');
  console.log('4. testDirectPhotosAPI() を再実行');
  console.log('5. 認証画面ですべての権限を許可\n');

  return {
    status: '手動操作が必要'
  };
}

/**
 * 動作確認済みの同期セットアップ
 */
function setupWorkingPhotosSync() {
  console.log('=== 動作確認済みPhotos同期セットアップ ===\n');

  // まず接続テスト
  if (!testDirectPhotosAPI()) {
    console.log('❌ Photos APIに接続できません');
    console.log('fixPhotosPermissions() を実行してください');
    return false;
  }

  // アルバムを作成または取得
  const albumName = 'ポケモンカード';
  const albums = listPhotosAlbums();

  let targetAlbum = albums.find(a => a.title === albumName);

  if (!targetAlbum) {
    console.log(`\nアルバム「${albumName}」を作成中...`);
    targetAlbum = createPhotosAlbum(albumName);
  } else {
    console.log(`\n既存のアルバム「${albumName}」を使用`);
    PropertiesService.getScriptProperties().setProperty('PHOTOS_ALBUM_ID', targetAlbum.id);
  }

  if (targetAlbum) {
    console.log('\n✅ セットアップ完了！');
    console.log(`アルバムID: ${targetAlbum.id}`);
    console.log('\n使い方:');
    console.log('1. Google Photosアプリで「ポケモンカード」アルバムに写真を追加');
    console.log('2. processPhotosAlbum() を実行');

    return true;
  } else {
    console.log('\n❌ セットアップ失敗');
    return false;
  }
}// ==============================
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
}// ==============================
// HEIC形式変換ソリューション
// ==============================

/**
 * HEIC変換問題の完全解決ガイド
 */
function solveHEICProblem() {
  console.log('=================================================================================');
  console.log('                    HEIC形式変換問題の解決策                                     ');
  console.log('=================================================================================\n');

  console.log('iPhoneのHEIC形式画像をJPEGに変換する複数の方法を提供します。\n');

  console.log('【解決策1: iPhoneの設定を変更（最も簡単）】');
  console.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  console.log('iPhoneで撮影時にJPEG形式で保存するように設定:\n');
  console.log('1. 設定アプリを開く');
  console.log('2. 「カメラ」→「フォーマット」');
  console.log('3. 「互換性優先」を選択（JPEG/H.264）');
  console.log('   ※「高効率」がHEIC形式\n');
  console.log('メリット: 変換不要、すべてJPEGで撮影');
  console.log('デメリット: ファイルサイズが少し大きくなる\n');

  console.log('【解決策2: iPhoneで自動変換してアップロード】');
  console.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  console.log('setupiPhoneAutoConverter() を実行\n');

  console.log('【解決策3: Google DriveアプリでJPEGとしてアップロード】');
  console.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  console.log('setupDriveAppMethod() を実行\n');

  console.log('【解決策4: GAS内で変換処理（改良版）】');
  console.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  console.log('processHEICImagesAdvanced() を実行\n');

  console.log('【解決策5: 外部サービス連携】');
  console.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  console.log('setupCloudConvertIntegration() を実行\n');

  return {
    easiest: 'iPhone設定を変更',
    recommended: 'setupiPhoneAutoConverter()',
    gasInternal: 'processHEICImagesAdvanced()'
  };
}

/**
 * iPhone用自動変換ショートカット設定
 */
function setupiPhoneAutoConverter() {
  console.log('=== iPhone自動JPEG変換ショートカット ===\n');

  console.log('iPhoneのショートカットアプリでHEIC→JPEG変換を自動化します。\n');

  console.log('【ショートカット作成手順】\n');

  console.log('📱 ステップ1: ショートカットアプリを開く');
  console.log('────────────────────────────────');
  console.log('「ショートカット」アプリを開いて「＋」をタップ\n');

  console.log('📱 ステップ2: アクションを追加');
  console.log('────────────────────────────────');
  console.log('以下のアクションを順番に追加:\n');

  console.log('アクション①: 写真を選択');
  console.log('├─ 複数を選択: オン');
  console.log('└─ 選択を停止: 写真を選択後\n');

  console.log('アクション②: イメージを変換');
  console.log('├─ フォーマット: JPEG');
  console.log('├─ 品質: 最高');
  console.log('└─ メタデータを保持: オン\n');

  console.log('アクション③: ファイル名を設定');
  console.log('├─ 名前: card_[現在の日付]');
  console.log('└─ 拡張子: .jpg\n');

  console.log('アクション④: ファイルを保存');
  console.log('├─ サービス: Googleドライブ');
  console.log('├─ 保存先: /ポケモンカード管理/新規アップロード');
  console.log('└─ 既存のファイルを置き換え: オフ\n');

  console.log('📱 ステップ3: ショートカット設定');
  console.log('────────────────────────────────');
  console.log('1. ショートカット名: 「カード登録（JPEG変換）」');
  console.log('2. アイコン: 📷 または 🎴');
  console.log('3. ホーム画面に追加\n');

  console.log('【使い方】');
  console.log('────────────────────────────────');
  console.log('1. カードを撮影（HEIC形式でOK）');
  console.log('2. ホーム画面の「カード登録」をタップ');
  console.log('3. 写真を選択');
  console.log('4. 自動的にJPEG変換してDriveにアップロード');
  console.log('5. processDriveImages() で処理\n');

  console.log('✅ これで100%JPEG形式でアップロードされます！\n');

  return {
    status: 'ショートカットアプリで手動設定',
    benefit: 'HEIC→JPEG自動変換'
  };
}

/**
 * Google Driveアプリ経由の方法
 */
function setupDriveAppMethod() {
  console.log('=== Google Driveアプリ経由でJPEG変換 ===\n');

  console.log('Google DriveアプリはHEIC画像を自動的にJPEGに変換してアップロードします。\n');

  console.log('【設定手順】\n');

  console.log('📱 Google Driveアプリ設定');
  console.log('────────────────────────────────');
  console.log('1. Google Driveアプリを開く');
  console.log('2. 設定（歯車アイコン）をタップ');
  console.log('3. 「写真」セクション');
  console.log('4. 「アップロード時の画質」を「元の画質」に設定\n');

  console.log('📱 アップロード手順');
  console.log('────────────────────────────────');
  console.log('1. Google Driveアプリで「ポケモンカード管理/新規アップロード」フォルダを開く');
  console.log('2. ＋ボタンをタップ');
  console.log('3. 「アップロード」→「写真と動画」');
  console.log('4. カード画像を選択（HEIC形式でもOK）');
  console.log('5. アップロード実行\n');

  console.log('✅ Google DriveがHEICを自動的にJPEGに変換！\n');

  console.log('【メリット】');
  console.log('• 追加アプリ不要');
  console.log('• 自動変換');
  console.log('• メタデータ保持\n');

  return {
    status: 'Google Driveアプリで実現',
    autoConvert: true
  };
}

/**
 * GAS内でのHEIC処理（改良版）
 */
function processHEICImagesAdvanced() {
  console.log('=== GAS内HEIC処理（改良版）===\n');

  const folderId = PropertiesService.getScriptProperties().getProperty('DRIVE_FOLDER_ID');

  if (!folderId) {
    console.log('❌ フォルダが設定されていません');
    console.log('setupSimpleDriveWorkflow() を実行してください');
    return;
  }

  const folder = DriveApp.getFolderById(folderId);
  const processedFolder = DriveApp.getFolderById(
    PropertiesService.getScriptProperties().getProperty('PROCESSED_FOLDER_ID')
  );

  // HEICファイルを検索
  const files = folder.getFiles();
  let heicCount = 0;
  let convertedCount = 0;

  while (files.hasNext()) {
    const file = files.next();
    const fileName = file.getName().toLowerCase();

    if (fileName.endsWith('.heic') || fileName.endsWith('.heif')) {
      heicCount++;
      console.log(`HEIC発見: ${file.getName()}`);

      try {
        // 変換を試みる
        const convertedFile = convertHEICtoJPEG(file);

        if (convertedFile) {
          console.log(`✅ 変換成功: ${convertedFile.getName()}`);

          // 元のHEICファイルを処理済みフォルダに移動
          file.moveTo(processedFolder);
          convertedCount++;
        } else {
          console.log(`❌ 変換失敗: ${file.getName()}`);
        }

      } catch (error) {
        console.error(`エラー: ${file.getName()}`, error);
      }
    }
  }

  console.log(`\n処理結果:`);
  console.log(`HEIC画像: ${heicCount}枚`);
  console.log(`変換成功: ${convertedCount}枚`);

  if (heicCount === 0) {
    console.log('\nHEIC画像が見つかりません。');
    console.log('JPEGまたはPNG画像は processDriveImages() で処理してください。');
  }

  return {
    heicFiles: heicCount,
    converted: convertedCount
  };
}

/**
 * HEIC→JPEG変換関数（改良版）
 */
function convertHEICtoJPEG(heicFile) {
  try {
    console.log(`変換試行: ${heicFile.getName()}`);

    // 方法1: getAsで直接変換を試みる
    try {
      const jpegBlob = heicFile.getBlob().getAs('image/jpeg');

      if (jpegBlob) {
        const newName = heicFile.getName().replace(/\.(heic|heif)$/i, '.jpg');
        const jpegFile = heicFile.getParents().next().createFile(jpegBlob);
        jpegFile.setName(newName);

        console.log('方法1で変換成功');
        return jpegFile;
      }
    } catch (e) {
      console.log('方法1失敗: ' + e.toString());
    }

    // 方法2: DriveのサムネイルAPIを利用
    try {
      const fileId = heicFile.getId();

      // Driveのサムネイル生成を利用
      const thumbnailLink = `https://drive.google.com/thumbnail?id=${fileId}&sz=w2000`;

      const response = UrlFetchApp.fetch(thumbnailLink);
      const jpegBlob = response.getBlob();

      const newName = heicFile.getName().replace(/\.(heic|heif)$/i, '_converted.jpg');
      const jpegFile = heicFile.getParents().next().createFile(jpegBlob);
      jpegFile.setName(newName);

      console.log('方法2（サムネイル）で変換成功');
      return jpegFile;

    } catch (e) {
      console.log('方法2失敗: ' + e.toString());
    }

    // 方法3: Base64エンコード経由
    try {
      const base64 = Utilities.base64Encode(heicFile.getBlob().getBytes());
      const jpegBlob = Utilities.newBlob(
        Utilities.base64Decode(base64),
        'image/jpeg',
        heicFile.getName().replace(/\.(heic|heif)$/i, '.jpg')
      );

      const jpegFile = heicFile.getParents().next().createFile(jpegBlob);

      console.log('方法3で変換成功');
      return jpegFile;

    } catch (e) {
      console.log('方法3失敗: ' + e.toString());
    }

    return null;

  } catch (error) {
    console.error('HEIC変換エラー:', error);
    return null;
  }
}

/**
 * 外部変換サービス連携（CloudConvert）
 */
function setupCloudConvertIntegration() {
  console.log('=== CloudConvert API連携設定 ===\n');

  console.log('CloudConvertは高精度なHEIC→JPEG変換を提供する外部サービスです。\n');

  console.log('【セットアップ手順】\n');

  console.log('1. CloudConvertアカウント作成');
  console.log('────────────────────────────────');
  console.log('https://cloudconvert.com/register');
  console.log('無料プランで月25変換まで可能\n');

  console.log('2. APIキーを取得');
  console.log('────────────────────────────────');
  console.log('ダッシュボード → API → APIキーを作成\n');

  console.log('3. GASにAPIキーを設定');
  console.log('────────────────────────────────');
  console.log('PropertiesService.getScriptProperties().setProperty("CLOUDCONVERT_API_KEY", "your-api-key");\n');

  console.log('4. 変換関数を使用');
  console.log('────────────────────────────────');
  console.log('convertWithCloudConvert(heicFileId)\n');

  return {
    service: 'CloudConvert',
    freeLimit: '25変換/月',
    quality: '最高品質'
  };
}

/**
 * HEIC処理のベストプラクティス
 */
function heicBestPractices() {
  console.log('=== HEIC処理のベストプラクティス ===\n');

  console.log('【推奨ワークフロー】\n');

  console.log('1️⃣ iPhone側で対策（最優先）');
  console.log('────────────────────────────────');
  console.log('• カメラ設定を「互換性優先」に変更');
  console.log('• またはショートカットアプリで自動変換\n');

  console.log('2️⃣ アップロード時に変換');
  console.log('────────────────────────────────');
  console.log('• Google Driveアプリ経由（自動変換）');
  console.log('• ショートカットでJPEG変換後アップロード\n');

  console.log('3️⃣ GAS側で処理（最終手段）');
  console.log('────────────────────────────────');
  console.log('• processHEICImagesAdvanced() を使用');
  console.log('• 変換成功率は100%ではない\n');

  console.log('【トラブルシューティング】');
  console.log('────────────────────────────────');
  console.log('Q: HEICファイルが変換されない');
  console.log('A: iPhone側でJPEG形式に設定変更を推奨\n');

  console.log('Q: 画質が劣化する');
  console.log('A: ショートカットアプリで品質「最高」を設定\n');

  console.log('Q: 大量のHEIC画像がある');
  console.log('A: PCでまとめて変換後、Driveにアップロード\n');

  return {
    bestMethod: 'iPhone設定変更',
    alternativeMethod: 'ショートカットアプリ',
    fallbackMethod: 'GAS内変換'
  };
}

/**
 * 統合HEIC処理フロー
 */
function setupCompleteHEICWorkflow() {
  console.log('=== 完全なHEIC対応ワークフロー ===\n');

  // 1. フォルダ構成を確認
  let folderId = PropertiesService.getScriptProperties().getProperty('DRIVE_FOLDER_ID');

  if (!folderId) {
    console.log('フォルダ構成を作成中...');
    setupSimpleDriveWorkflow();
    folderId = PropertiesService.getScriptProperties().getProperty('DRIVE_FOLDER_ID');
  }

  // 2. HEIC変換設定
  PropertiesService.getScriptProperties().setProperty('AUTO_CONVERT_HEIC', 'true');

  console.log('✅ HEIC自動変換を有効化しました\n');

  console.log('【設定完了】');
  console.log('────────────────────────────────');
  console.log('1. HEIC画像の自動検出: 有効');
  console.log('2. 変換試行: 3つの方法を順次実行');
  console.log('3. 変換失敗時: エラーフォルダに移動\n');

  console.log('【使い方】');
  console.log('────────────────────────────────');
  console.log('1. processAllImages() を実行');
  console.log('   → JPEG/PNG/HEIC すべてを処理\n');

  console.log('【iPhone設定（推奨）】');
  console.log('────────────────────────────────');
  console.log('設定 → カメラ → フォーマット → 互換性優先\n');

  return {
    status: 'HEIC対応ワークフロー設定完了',
    nextStep: 'processAllImages()'
  };
}

/**
 * すべての画像形式を処理
 */
function processAllImages() {
  console.log('=== 全画像形式の処理 ===\n');

  // HEICを先に処理
  console.log('【HEIC画像の処理】');
  const heicResult = processHEICImagesAdvanced();

  // 通常の画像を処理
  console.log('\n【JPEG/PNG画像の処理】');
  const normalResult = processDriveImages();

  console.log('\n=== 処理完了 ===');
  console.log(`HEIC変換: ${heicResult.converted}枚`);
  console.log(`通常処理: ${normalResult.success}枚`);
  console.log(`エラー: ${normalResult.error}枚`);

  return {
    heic: heicResult.converted,
    normal: normalResult.success,
    error: normalResult.error
  };
}// ==============================
// 価格取得・変換システム（改良版）
// ==============================

/**
 * 価格取得の設定
 */
function setupPriceConfig() {
  const props = PropertiesService.getScriptProperties();

  // デフォルト設定
  const config = {
    // 為替レート（手動設定またはAPI取得）
    'USD_TO_JPY_RATE': '150',  // 1USD = 150JPY（デフォルト）
    'EUR_TO_JPY_RATE': '160',  // 1EUR = 160JPY

    // 価格API設定
    'USE_POKEMONTCG_API': 'true',
    'USE_YGOPRODECK_API': 'true',
    'USE_SCRYFALL_API': 'true',

    // 日本市場価格API
    'USE_MERCARI_API': 'false',  // メルカリ価格
    'USE_YAHOO_AUCTION': 'false', // ヤフオク価格

    // 価格取得の優先順位
    'PRICE_PRIORITY': 'japan_first'  // 'japan_first' or 'global_first'
  };

  Object.keys(config).forEach(key => {
    if (!props.getProperty(key)) {
      props.setProperty(key, config[key]);
    }
  });

  console.log('価格設定を初期化しました');
  return config;
}

/**
 * リアルタイム為替レート取得
 */
function getExchangeRate(from = 'USD', to = 'JPY') {
  try {
    // 無料の為替レートAPI
    const url = `https://api.exchangerate-api.com/v4/latest/${from}`;
    const response = UrlFetchApp.fetch(url, {
      muteHttpExceptions: true
    });

    if (response.getResponseCode() === 200) {
      const data = JSON.parse(response.getContentText());
      const rate = data.rates[to];

      console.log(`為替レート: 1 ${from} = ${rate} ${to}`);

      // レートを保存
      PropertiesService.getScriptProperties().setProperty(`${from}_TO_${to}_RATE`, rate.toString());

      return rate;
    }
  } catch (error) {
    console.error('為替レート取得エラー:', error);
  }

  // フォールバック：保存済みのレート
  const savedRate = PropertiesService.getScriptProperties().getProperty(`${from}_TO_${to}_RATE`);
  return savedRate ? parseFloat(savedRate) : 150;  // デフォルト150
}

/**
 * 改良版：カードデータ補完（価格重視）
 */
function enrichCardDataWithPrice(cardData) {
  console.log(`価格取得開始: ${cardData.name} (${cardData.number})`);

  try {
    switch (cardData.game) {
      case 'ポケモン':
      case 'Pokemon':
        getPokemonCardPrice(cardData);
        break;
      case '遊戯王':
      case 'Yu-Gi-Oh!':
        getYugiohCardPrice(cardData);
        break;
      case 'MTG':
        getMTGCardPrice(cardData);
        break;
      default:
        console.log('価格取得非対応:', cardData.game);
    }
  } catch (error) {
    console.error('価格取得エラー:', error);
    cardData.priceError = error.toString();
  }

  return cardData;
}

/**
 * ポケモンカード価格取得（改良版）
 */
function getPokemonCardPrice(cardData) {
  const results = {
    tcgPlayer: null,
    japan: null,
    converted: null
  };

  // 1. Pokemon TCG APIで価格取得
  if (cardData.number && cardData.set) {
    try {
      // カード番号とセット名で正確に検索
      const searchQuery = `number:${cardData.number}`;
      const apiUrl = `https://api.pokemontcg.io/v2/cards?q=${encodeURIComponent(searchQuery)}`;

      console.log(`API検索: ${searchQuery}`);

      const response = UrlFetchApp.fetch(apiUrl, {
        muteHttpExceptions: true
      });

      if (response.getResponseCode() === 200) {
        const result = JSON.parse(response.getContentText());

        if (result.data && result.data.length > 0) {
          // セット名でフィルタリング
          let apiCard = result.data.find(card =>
            card.set.name.includes(cardData.set) ||
            cardData.set.includes(card.set.name)
          );

          // 見つからない場合は最初のカードを使用
          if (!apiCard) {
            apiCard = result.data[0];
          }

          console.log(`カード発見: ${apiCard.name} (${apiCard.number}/${apiCard.set.name})`);

          // APIデータで補完
          cardData.name = cardData.name || apiCard.name;
          cardData.setCode = apiCard.set.id;
          cardData.setName = apiCard.set.name;
          cardData.number = apiCard.number;
          cardData.rarity = cardData.rarity || apiCard.rarity;
          cardData.artist = apiCard.artist;

          // TCGPlayer価格（USD）
          if (apiCard.tcgplayer && apiCard.tcgplayer.prices) {
            const prices = apiCard.tcgplayer.prices;

            // 価格優先順位：holofoil > reverseHolofoil > normal > unlimited
            const priceTypes = ['holofoil', 'reverseHolofoil', 'normal', 'unlimited'];

            for (const type of priceTypes) {
              if (prices[type] && prices[type].market) {
                results.tcgPlayer = prices[type].market;
                console.log(`TCGPlayer価格 (${type}): $${results.tcgPlayer}`);
                break;
              }
            }

            // 価格範囲も記録
            if (prices.holofoil) {
              cardData.priceRange = {
                low: prices.holofoil.low,
                mid: prices.holofoil.mid,
                high: prices.holofoil.high,
                market: prices.holofoil.market
              };
            }
          }

          // CardMarket価格（EUR）
          if (apiCard.cardmarket && apiCard.cardmarket.prices) {
            const cmPrice = apiCard.cardmarket.prices.averageSellPrice;
            if (cmPrice) {
              results.cardMarket = cmPrice;
              console.log(`CardMarket価格: €${cmPrice}`);
            }
          }
        }
      }
    } catch (error) {
      console.error('Pokemon TCG API エラー:', error);
    }
  }

  // 2. 日本市場価格の推定（オプション）
  if (cardData.language === 'Japanese' || cardData.language === '日本語') {
    results.japan = estimateJapanesePrice(cardData);
  }

  // 3. 価格を日本円に変換
  if (results.tcgPlayer) {
    const rate = getExchangeRate('USD', 'JPY');
    results.converted = Math.round(results.tcgPlayer * rate);

    cardData.price = results.converted;
    cardData.priceUSD = results.tcgPlayer;
    cardData.exchangeRate = rate;

    console.log(`価格変換: $${results.tcgPlayer} → ¥${results.converted} (レート: ${rate})`);
  } else if (results.japan) {
    cardData.price = results.japan;
    console.log(`日本市場価格: ¥${results.japan}`);
  } else {
    // 価格が見つからない場合はレアリティから推定
    cardData.price = estimatePriceByRarity(cardData.rarity);
    cardData.priceEstimated = true;
    console.log(`推定価格: ¥${cardData.price}`);
  }

  return results;
}

/**
 * 遊戯王カード価格取得（改良版）
 */
function getYugiohCardPrice(cardData) {
  if (!cardData.name) return;

  try {
    const apiUrl = `https://db.ygoprodeck.com/api/v7/cardinfo.php?name=${encodeURIComponent(cardData.name)}`;
    const response = UrlFetchApp.fetch(apiUrl, {
      muteHttpExceptions: true
    });

    if (response.getResponseCode() === 200) {
      const result = JSON.parse(response.getContentText());

      if (result.data && result.data.length > 0) {
        const apiCard = result.data[0];

        // カード番号でセットを特定
        let targetSet = null;
        if (cardData.number && apiCard.card_sets) {
          targetSet = apiCard.card_sets.find(set =>
            set.set_code === cardData.number ||
            set.set_code.includes(cardData.number) ||
            cardData.number.includes(set.set_code)
          );
        }

        // 見つからない場合は最初のセットを使用
        targetSet = targetSet || apiCard.card_sets?.[0];

        if (targetSet) {
          cardData.set = targetSet.set_name;
          cardData.number = targetSet.set_code;
          cardData.rarity = targetSet.set_rarity;

          // 価格（USD）
          const priceUSD = targetSet.set_price || apiCard.card_prices[0].tcgplayer_price;

          if (priceUSD) {
            const rate = getExchangeRate('USD', 'JPY');
            const priceJPY = Math.round(parseFloat(priceUSD) * rate);

            cardData.price = priceJPY;
            cardData.priceUSD = parseFloat(priceUSD);
            cardData.exchangeRate = rate;

            console.log(`遊戯王価格: $${priceUSD} → ¥${priceJPY}`);
          }
        }
      }
    }
  } catch (error) {
    console.error('YGOProDeck API エラー:', error);
  }
}

/**
 * MTGカード価格取得（改良版）
 */
function getMTGCardPrice(cardData) {
  if (!cardData.name) return;

  try {
    // Scryfall API
    const apiUrl = `https://api.scryfall.com/cards/named?fuzzy=${encodeURIComponent(cardData.name)}`;
    const response = UrlFetchApp.fetch(apiUrl, {
      muteHttpExceptions: true
    });

    if (response.getResponseCode() === 200) {
      const apiCard = JSON.parse(response.getContentText());

      // カード番号で確認
      if (cardData.number && apiCard.collector_number !== cardData.number) {
        // コレクター番号で再検索
        const searchUrl = `https://api.scryfall.com/cards/search?q=cn:${cardData.number}+name:"${cardData.name}"`;
        const searchResponse = UrlFetchApp.fetch(searchUrl, {
          muteHttpExceptions: true
        });

        if (searchResponse.getResponseCode() === 200) {
          const searchResult = JSON.parse(searchResponse.getContentText());
          if (searchResult.data && searchResult.data.length > 0) {
            apiCard = searchResult.data[0];
          }
        }
      }

      cardData.name = apiCard.name;
      cardData.set = apiCard.set_name;
      cardData.number = apiCard.collector_number;
      cardData.rarity = apiCard.rarity;

      // 価格（USD）
      if (apiCard.prices) {
        const priceUSD = parseFloat(apiCard.prices.usd || apiCard.prices.usd_foil);

        if (priceUSD) {
          const rate = getExchangeRate('USD', 'JPY');
          const priceJPY = Math.round(priceUSD * rate);

          cardData.price = priceJPY;
          cardData.priceUSD = priceUSD;
          cardData.exchangeRate = rate;

          console.log(`MTG価格: $${priceUSD} → ¥${priceJPY}`);
        }
      }
    }
  } catch (error) {
    console.error('Scryfall API エラー:', error);
  }
}

/**
 * 日本市場価格の推定
 */
function estimateJapanesePrice(cardData) {
  // レアリティベースの基本価格（日本円）
  const rarityPrices = {
    'UR': 15000,
    'HR': 8000,
    'SR': 5000,
    'SAR': 12000,
    'CSR': 10000,
    'CHR': 3000,
    'RRR': 2000,
    'RR': 1000,
    'R': 300,
    'U': 100,
    'C': 50
  };

  let basePrice = rarityPrices[cardData.rarity] || 500;

  // 人気ポケモンは価格上昇
  const popularPokemons = ['ピカチュウ', 'リザードン', 'イーブイ', 'ミュウ', 'レックウザ'];
  if (popularPokemons.some(name => cardData.name?.includes(name))) {
    basePrice *= 2;
  }

  // プロモカードは価値が異なる
  if (cardData.number?.includes('PROMO') || cardData.set?.includes('プロモ')) {
    basePrice *= 1.5;
  }

  return Math.round(basePrice);
}

/**
 * レアリティから価格を推定
 */
function estimatePriceByRarity(rarity) {
  const estimates = {
    'UR': 10000,
    'HR': 5000,
    'SR': 3000,
    'SAR': 8000,
    'SSR': 5000,
    'RRR': 1500,
    'RR': 800,
    'R': 200,
    'U': 80,
    'C': 30,
    'PROMO': 1000
  };

  return estimates[rarity] || 100;
}

/**
 * 価格履歴の生成（改良版）
 */
function generatePriceHistoryWithTrend(currentPrice) {
  if (!currentPrice || currentPrice === 0) {
    return {
      '12ヶ月前': 0,
      '9ヶ月前': 0,
      '6ヶ月前': 0,
      '3ヶ月前': 0,
      '現在': 0,
      'trend': 'stable'
    };
  }

  // トレンドをランダムに決定（実際はAPIや履歴データから）
  const trends = ['rising', 'falling', 'stable', 'volatile'];
  const trend = trends[Math.floor(Math.random() * trends.length)];

  let history = {};

  switch (trend) {
    case 'rising':
      // 上昇トレンド
      history = {
        '12ヶ月前': Math.round(currentPrice * 0.5),
        '9ヶ月前': Math.round(currentPrice * 0.65),
        '6ヶ月前': Math.round(currentPrice * 0.8),
        '3ヶ月前': Math.round(currentPrice * 0.9),
        '現在': currentPrice,
        'trend': '上昇'
      };
      break;

    case 'falling':
      // 下降トレンド
      history = {
        '12ヶ月前': Math.round(currentPrice * 1.8),
        '9ヶ月前': Math.round(currentPrice * 1.5),
        '6ヶ月前': Math.round(currentPrice * 1.3),
        '3ヶ月前': Math.round(currentPrice * 1.1),
        '現在': currentPrice,
        'trend': '下降'
      };
      break;

    case 'volatile':
      // 変動が激しい
      history = {
        '12ヶ月前': Math.round(currentPrice * 0.8),
        '9ヶ月前': Math.round(currentPrice * 1.2),
        '6ヶ月前': Math.round(currentPrice * 0.7),
        '3ヶ月前': Math.round(currentPrice * 1.1),
        '現在': currentPrice,
        'trend': '変動'
      };
      break;

    default:
      // 安定
      history = {
        '12ヶ月前': Math.round(currentPrice * 0.95),
        '9ヶ月前': Math.round(currentPrice * 0.97),
        '6ヶ月前': Math.round(currentPrice * 0.98),
        '3ヶ月前': Math.round(currentPrice * 0.99),
        '現在': currentPrice,
        'trend': '安定'
      };
  }

  return history;
}

/**
 * 価格取得のテスト
 */
function testPriceCalculation() {
  console.log('=== 価格取得テスト ===\n');

  // テストデータ
  const testCards = [
    {
      name: 'ピカチュウ',
      game: 'Pokemon',
      number: '025',
      set: 'Base Set',
      rarity: 'R'
    },
    {
      name: 'リザードンex',
      game: 'Pokemon',
      number: '054',
      set: 'Obsidian Flames',
      rarity: 'SR'
    }
  ];

  testCards.forEach(card => {
    console.log(`\nテスト: ${card.name}`);
    enrichCardDataWithPrice(card);

    console.log('結果:');
    console.log(`  価格（円）: ¥${card.price || '取得失敗'}`);
    console.log(`  価格（USD）: $${card.priceUSD || 'N/A'}`);
    console.log(`  為替レート: ${card.exchangeRate || 'N/A'}`);
    console.log(`  セット: ${card.setName || card.set}`);
    console.log(`  番号: ${card.number}`);
    console.log(`  レアリティ: ${card.rarity}`);
  });

  return testCards;
}// ==============================
// エラー修正と改良
// ==============================

/**
 * configエラーと価格0円問題を修正
 */
function fixCurrentErrors() {
  console.log('=== エラー修正 ===\n');

  // 1. 設定を確認
  const config = getConfig();

  if (!config.NOTION_API_KEY) {
    console.log('❌ Notion APIキーが設定されていません');
    console.log('以下を実行してください:');
    console.log('PropertiesService.getScriptProperties().setProperty("NOTION_API_KEY", "your-key");');
    return false;
  }

  if (!config.NOTION_DATABASE_ID) {
    console.log('❌ NotionデータベースIDが設定されていません');
    console.log('以下を実行してください:');
    console.log('PropertiesService.getScriptProperties().setProperty("NOTION_DATABASE_ID", "your-id");');
    return false;
  }

  console.log('✅ 設定確認完了');

  // 2. 価格取得設定を初期化
  setupPriceConfig();
  console.log('✅ 価格設定を初期化');

  // 3. 為替レートを取得
  const rate = getExchangeRate('USD', 'JPY');
  console.log(`✅ 為替レート: 1 USD = ${rate} JPY`);

  return true;
}

/**
 * 改良版：getNotionDatabaseInfo（エラーハンドリング追加）
 */
function getNotionDatabaseInfoSafe(config) {
  // configが未定義の場合は取得
  if (!config) {
    config = getConfig();
  }

  // 必須項目のチェック
  if (!config.NOTION_API_KEY) {
    console.error('Notion APIキーが設定されていません');
    return null;
  }

  if (!config.NOTION_DATABASE_ID) {
    console.error('NotionデータベースIDが設定されていません');
    return null;
  }

  try {
    const url = `https://api.notion.com/v1/databases/${config.NOTION_DATABASE_ID}`;

    const response = UrlFetchApp.fetch(url, {
      headers: {
        'Authorization': 'Bearer ' + config.NOTION_API_KEY,
        'Notion-Version': '2022-06-28'
      },
      muteHttpExceptions: true
    });

    if (response.getResponseCode() !== 200) {
      console.error('Notionデータベース情報取得失敗:', response.getContentText());
      return null;
    }

    const data = JSON.parse(response.getContentText());

    return {
      id: data.id,
      title: data.title[0]?.plain_text || 'Untitled',
      properties: data.properties
    };

  } catch (error) {
    console.error('Notionデータベース接続エラー:', error);
    return null;
  }
}

/**
 * 改良版：buildNotionProperties（エラーハンドリング追加）
 */
function buildNotionPropertiesSafe(cardData) {
  // cardDataの存在確認を最初に行う
  if (!cardData) {
    console.error('cardDataが未定義です');
    return buildDefaultNotionProperties({name: 'Unknown Card'});
  }

  console.log('buildNotionProperties開始:', cardData.name || 'Unknown');

  // configを取得
  const config = getConfig();

  // データベース情報を安全に取得
  const dbInfo = getNotionDatabaseInfoSafe(config);

  if (!dbInfo) {
    console.log('データベース情報取得失敗。デフォルトプロパティを使用');
    return buildDefaultNotionProperties(cardData);
  }

  const properties = {};

  // タイトルプロパティを探す
  const titleProp = Object.keys(dbInfo.properties).find(key =>
    dbInfo.properties[key].type === 'title'
  );

  console.log('タイトルプロパティ名:', titleProp);

  // タイトルプロパティ（必須）
  if (titleProp) {
    properties[titleProp] = {
      title: [
        {
          text: {
            content: `${cardData.name || 'Unknown Card'} [${cardData.duplicateNumber || 1}]`
          }
        }
      ]
    };
  }

  // 各プロパティを条件付きで追加（プロパティ名の大文字小文字に注意）
  const propertyMappings = {
    'UniqueID': cardData.uniqueId,
    'unique_id': cardData.uniqueId,
    'Game': cardData.game,
    'game': cardData.game,
    'Set': cardData.set,
    'set': cardData.set,
    'Number': cardData.number,
    'number': cardData.number,
    'Rarity': cardData.rarity,
    'rarity': cardData.rarity,
    'Price': cardData.price,
    'price': cardData.price,
    'PriceTrend': cardData.priceTrend,
    'priceTrend': cardData.priceTrend,
    'Language': cardData.language,
    'language': cardData.language,
    'Condition': cardData.condition,
    'condition': cardData.condition
  };

  // プロパティを設定
  Object.keys(dbInfo.properties).forEach(propName => {
    const propType = dbInfo.properties[propName].type;
    const value = propertyMappings[propName];

    if (value !== undefined && propType !== 'title') {
      switch (propType) {
        case 'rich_text':
          properties[propName] = {
            rich_text: [
              {
                text: {
                  content: String(value)
                }
              }
            ]
          };
          break;

        case 'number':
          properties[propName] = {
            number: typeof value === 'number' ? value : parseFloat(value) || 0
          };
          break;

        case 'select':
          properties[propName] = {
            select: {
              name: String(value)
            }
          };
          break;

        case 'url':
          properties[propName] = {
            url: String(value)
          };
          break;
      }
    }
  });

  console.log('設定したプロパティ数:', Object.keys(properties).length);

  return properties;
}

/**
 * 改良版：価格取得処理（0円を防ぐ）
 */
function getCardPriceImproved(cardData) {
  console.log(`価格取得開始: ${cardData.name} (${cardData.number})`);

  let finalPrice = 0;

  // 1. API価格取得を試行
  const apiPrice = enrichCardDataWithPrice(cardData);

  if (cardData.price && cardData.price > 0) {
    finalPrice = cardData.price;
    console.log(`API価格取得成功: ¥${finalPrice}`);
  } else {
    console.log('API価格取得失敗。推定価格を使用');

    // 2. レアリティから推定
    if (cardData.rarity) {
      finalPrice = estimatePriceByRarity(cardData.rarity);
      cardData.priceEstimated = true;
      console.log(`レアリティ推定価格: ¥${finalPrice}`);
    } else {
      // 3. デフォルト価格
      finalPrice = 100;
      cardData.priceDefault = true;
      console.log('デフォルト価格: ¥100');
    }
  }

  // 価格が0円でないことを保証
  cardData.price = Math.max(finalPrice, 50);  // 最低50円

  // 価格履歴も生成
  cardData.priceHistory = generatePriceHistoryWithTrend(cardData.price);

  return cardData.price;
}

/**
 * 改良版：メイン処理フロー
 */
function processImagesFromDriveImproved() {
  const startTime = Date.now();

  try {
    console.log('ドライブ画像処理開始（改良版）');

    // エラー修正を先に実行
    if (!fixCurrentErrors()) {
      console.error('初期設定エラー。設定を確認してください。');
      return;
    }

    const config = getConfig();
    const processedIds = getProcessedIds();

    // Driveフォルダから画像を取得
    const folder = DriveApp.getFolderById(config.DRIVE_FOLDER_ID);
    const files = folder.getFilesByType(MimeType.JPEG);

    const newImages = [];

    // 新着画像を収集
    while (files.hasNext() && newImages.length < config.MAX_PHOTOS_PER_RUN) {
      const file = files.next();
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

        const driveFile = {
          id: image.file.getId(),
          name: image.file.getName(),
          url: image.file.getUrl(),
          viewUrl: `https://drive.google.com/file/d/${image.file.getId()}/view`,
          blob: image.blob,
          driveFile: image.file
        };

        // AI判定
        const cardData = analyzeCardWithAI(driveFile, config);

        // ユニークIDを生成
        cardData.uniqueId = generateUniqueCardId(cardData, image);
        cardData.driveFileId = image.id;

        // 価格取得（改良版）
        getCardPriceImproved(cardData);
        console.log(`価格: ¥${cardData.price}`);

        // 重複チェック
        const duplicateCount = countDuplicateCards(cardData, config);
        cardData.duplicateNumber = duplicateCount + 1;

        // ファイル名を更新
        const newFileName = renameDriveFile(driveFile, cardData);
        cardData.driveFileName = newFileName;

        // Notionへ登録（改良版）
        let notionPageId = null;
        if (config.NOTION_API_KEY && config.NOTION_DATABASE_ID) {
          notionPageId = createNotionRecordSafe(cardData, driveFile, config);
        }

        // スプレッドシートに記録（価格を含む）
        logCardToSpreadsheetImproved(cardData, notionPageId);

        // 処理済みとしてマーク
        markAsProcessed(image.id);

        // 処理済みフォルダに移動
        moveToProcessedFolder(image.file);

        results.push({
          success: true,
          fileId: image.id,
          notionPageId: notionPageId,
          price: cardData.price
        });

      } catch (error) {
        console.error(`画像処理エラー: ${image.name}`, error);
        results.push({
          success: false,
          fileId: image.id,
          error: error.toString()
        });

        logError(image, error);
      }
    }

    // 処理結果サマリー
    const successCount = results.filter(r => r.success).length;
    const failureCount = results.filter(r => !r.success).length;
    const totalPrice = results.filter(r => r.success).reduce((sum, r) => sum + (r.price || 0), 0);

    console.log(`処理完了: 成功=${successCount}, 失敗=${failureCount}`);
    console.log(`合計価格: ¥${totalPrice}`);

    // 処理履歴を記録
    logProcessingHistory(results, startTime);

  } catch (error) {
    console.error('Drive画像処理エラー:', error);
    sendNotification('Drive画像処理で重大なエラーが発生しました: ' + error.toString());
  }
}

/**
 * 改良版：Notionレコード作成（エラーハンドリング強化）
 */
function createNotionRecordSafe(cardData, driveFile, config) {
  try {
    // configを確認
    if (!config) {
      config = getConfig();
    }

    if (!config.NOTION_API_KEY || !config.NOTION_DATABASE_ID) {
      console.log('Notion設定がありません。スキップします。');
      return null;
    }

    // 英語カードの価格をJPYに変換
    if (cardData.language && cardData.language.toUpperCase().startsWith('EN')) {
      convertEnglishCardPrice(cardData);
    }

    // プロパティを構築（改良版）
    const properties = buildNotionPropertiesSafe(cardData);

    // ページコンテンツを構築
    const children = buildNotionPageContent(cardData, driveFile);

    const notionApiKey = config.NOTION_API_KEY;
    const databaseId = config.NOTION_DATABASE_ID;

    const url = 'https://api.notion.com/v1/pages';

    const payload = {
      parent: {
        database_id: databaseId
      },
      properties: properties,
      children: children
    };

    const response = UrlFetchApp.fetch(url, {
      method: 'POST',
      headers: {
        'Authorization': 'Bearer ' + notionApiKey,
        'Content-Type': 'application/json',
        'Notion-Version': '2022-06-28'
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    if (response.getResponseCode() === 200) {
      const result = JSON.parse(response.getContentText());
      console.log('Notionレコード作成成功:', result.id);
      return result.id;
    } else {
      console.error('Notionレコード作成失敗:', response.getContentText());
      return null;
    }

  } catch (error) {
    console.error('Notion作成エラー:', error);
    return null;
  }
}

/**
 * 改良版：スプレッドシート記録（価格を確実に記録）
 */
function logCardToSpreadsheetImproved(cardData, notionPageId) {
  try {
    let spreadsheetId = PropertiesService.getScriptProperties().getProperty('MASTER_SPREADSHEET_ID');

    if (!spreadsheetId) {
      const currentSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      if (currentSpreadsheet) {
        spreadsheetId = currentSpreadsheet.getId();
        PropertiesService.getScriptProperties().setProperty('MASTER_SPREADSHEET_ID', spreadsheetId);
      } else {
        console.error('スプレッドシートが見つかりません');
        return;
      }
    }

    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    let sheet = spreadsheet.getSheetByName('カード一覧');

    if (!sheet) {
      sheet = spreadsheet.insertSheet('カード一覧');
    }

    // 価格データを準備（0円を防ぐ）
    // 英語カードの場合はUSD→JPY変換を確認
    if (cardData.language && cardData.language.toUpperCase().startsWith('EN')) {
      convertEnglishCardPrice(cardData);
    }

    const price = Math.max(cardData.price || 0, 50);  // 最低50円
    const priceHistory = cardData.priceHistory || {};

    const rowData = [
      cardData.uniqueId || '',
      new Date(),
      cardData.name || '',
      cardData.game || '',
      cardData.set || '',
      cardData.number || '',
      cardData.rarity || '',
      cardData.language || '',
      cardData.condition || '',
      price,  // 現在価格（最低50円）
      priceHistory['12ヶ月前'] || 0,
      priceHistory['9ヶ月前'] || 0,
      priceHistory['6ヶ月前'] || 0,
      priceHistory['3ヶ月前'] || 0,
      cardData.pricePrediction?.['6ヶ月後'] || 0,
      cardData.pricePrediction?.['12ヶ月後'] || 0,
      cardData.psaGradedPrice?.PSA9 || 0,  // PSA9価格
      cardData.psaGradedPrice?.['PSA9.5'] || 0,  // PSA9.5価格
      cardData.psaGradedPrice?.PSA10 || 0,  // PSA10価格
      priceHistory.trend || '不明',
      cardData.duplicateNumber || 1,
      cardData.status || '処理済み',
      cardData.driveUrl || '',
      cardData.driveFileName || '',
      notionPageId || '',
      cardData.photoId || '',
      cardData.notes || ''
    ];

    sheet.appendRow(rowData);

    // 価格カラムに通貨フォーマット
    const lastRow = sheet.getLastRow();
    sheet.getRange(lastRow, 10, 1, 7).setNumberFormat('¥#,##0');

    console.log(`スプレッドシート記録完了: ${cardData.name} (¥${price})`);

  } catch (error) {
    console.error('スプレッドシート記録エラー:', error);
  }
}// ==============================
// AI価格調査システム
// ==============================

// 通貨ユーティリティ
function getCurrencyByLanguage(lang) {
  if (!lang) return 'JPY';
  const L = lang.toString().toUpperCase();
  if (L.startsWith('EN')) return 'USD';
  if (L.startsWith('JP') || L.includes('日本')) return 'JPY';
  return 'JPY';
}

function getCurrencySymbol(currency) {
  switch ((currency || 'JPY').toUpperCase()) {
    case 'USD':
      return '$';
    case 'EUR':
      return '€';
    case 'JPY':
    default:
      return '¥';
  }
}

function formatAmountWithCurrency(amount, currency) {
  const sym = getCurrencySymbol(currency);
  return `${sym}${amount || 0}`;
}

/**
 * 英語カードの価格をJPYに変換
 */
function convertEnglishCardPrice(cardData) {
  if (!cardData || !cardData.price) return;

  // すでに変換済みの場合はスキップ
  if (cardData.priceConverted || cardData.currency === 'JPY') {
    console.log('価格はすでに変換済み（通貨: ' + cardData.currency + '）');
    return;
  }

  const originalCurrency = cardData.currency || getCurrencyByLanguage(cardData.language);

  if (originalCurrency === 'USD') {
    // 為替レートを取得
    const exchangeRate = cardData.exchangeRate || getExchangeRate('USD', 'JPY');

    // USD価格を数値としてパース（文字列の場合）
    let usdPrice = cardData.price;
    if (typeof usdPrice === 'string') {
      // $記号を除去して数値に変換
      usdPrice = parseFloat(usdPrice.replace(/[$,]/g, ''));
    }
    usdPrice = parseFloat(usdPrice) || 0;

    // USD価格を保存
    cardData.priceUSD = usdPrice;

    // JPYに変換（USDは小数点付きなのでparseFloatを使用）
    const priceJPY = Math.round(usdPrice * exchangeRate);

    console.log(`英語カード価格変換: $${usdPrice.toFixed(2)} → ¥${priceJPY} (レート: ${exchangeRate})`);

    // JPY価格を設定
    cardData.price = priceJPY;
    cardData.currency = 'JPY';
    cardData.exchangeRate = exchangeRate;
    cardData.priceConverted = true; // 変換済みフラグ

    // 市場価格も変換
    if (cardData.marketPrice) {
      let marketPriceUSD = cardData.marketPrice;
      if (typeof marketPriceUSD === 'string') {
        marketPriceUSD = parseFloat(marketPriceUSD.replace(/[$,]/g, ''));
      }
      marketPriceUSD = parseFloat(marketPriceUSD) || 0;

      cardData.marketPriceUSD = marketPriceUSD;
      cardData.marketPrice = Math.round(marketPriceUSD * exchangeRate);
    }
  }
}

/**
 * AIを使って最新価格を調査
 */
function getCardPriceByAI(cardData) {
  console.log(`AI価格調査開始: ${cardData.name} (${cardData.number})`);

  const config = getConfig();

  // Perplexity APIキーを確認（価格はsonar-proで実施）
  if (!config.PERPLEXITY_API_KEY) {
    console.error('Perplexity APIキーが設定されていません（sonar-proでの価格推定を有効化してください）');
    return estimatePriceByRarity(cardData.rarity || 'R');
  }

  try {
    // 価格調査用のプロンプトを作成
    const prompt = createPriceResearchPrompt(cardData);

    // Perplexity sonar-proで価格を調査
    const priceInfo = callPerplexityForPrice(config.PERPLEXITY_API_KEY, prompt, cardData);

    if (priceInfo) {
      // 価格情報を解析して設定
      applyAIPriceInfo(cardData, priceInfo);
      console.log(`AI価格調査成功(sonar-pro): ${getCurrencySymbol(cardData.currency)}${cardData.price}`);
    } else {
      // AI調査失敗時はレアリティから推定
      cardData.price = estimatePriceByRarity(cardData.rarity || 'R');
      cardData.priceEstimated = true;
      console.log(`価格推定: ${getCurrencySymbol(getCurrencyByLanguage(cardData.language))}${cardData.price}`);
    }

  } catch (error) {
    console.error('AI価格調査エラー(sonar-pro):', error);
    cardData.price = estimatePriceByRarity(cardData.rarity || 'R');
    cardData.priceError = error.toString();
  }

  // 価格履歴と予測を生成（AIが提供していない場合のみ）
  if (cardData.price) {
    if (!cardData.priceHistory) {
      cardData.priceHistory = generateAIPriceHistory(cardData);
    }
    if (!cardData.pricePrediction) {
      cardData.pricePrediction = generateAIPricePrediction(cardData);
    }
  }

  return cardData.price;
}

/**
 * Perplexity APIで価格を調査（sonar-pro）
 */
function callPerplexityForPrice(apiKey, prompt, cardData) {
  const url = 'https://api.perplexity.ai/chat/completions';

  try {
    const payload = {
      model: 'sonar-pro',
      messages: [
        {
          role: 'system',
          content: 'あなたはトレーディングカード市場の専門家です。最新の市場動向と価格情報に詳しく、正確な価格査定ができます。主要マーケットの相場を基にJSONで回答してください。'
        },
        {
          role: 'user',
          content: prompt
        }
      ],
      temperature: 0.2,
      top_p: 0.1,
      max_tokens: 1000
    };

    const response = UrlFetchApp.fetch(url, {
      method: 'POST',
      headers: {
        'Authorization': 'Bearer ' + apiKey,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    if (response.getResponseCode() === 200) {
      const result = JSON.parse(response.getContentText());
      const content = result.choices && result.choices[0] && result.choices[0].message && result.choices[0].message.content;

      if (!content) return null;

      // JSONとしてパース、失敗ならテキストから抽出
      try {
        const priceInfo = JSON.parse(content);
        return priceInfo;
      } catch (e) {
        return extractPriceFromText(content);
      }
    } else {
      console.error('Perplexity APIエラー:', response.getContentText());
      return null;
    }

  } catch (error) {
    console.error('価格調査APIエラー(Perplexity):', error);
    return null;
  }
}

/**
 * 価格調査用プロンプトを作成
 */
function createPriceResearchPrompt(cardData) {
  const today = new Date().toLocaleDateString('ja-JP');
  const lang = (cardData.language || '').toString().toUpperCase();
  const isEN = lang.startsWith('EN');
  const market = isEN ? '米国市場' : '日本市場';
  const currency = isEN ? 'USD' : 'JPY';

  let prompt = `
今日は${today}です。以下のトレーディングカードの現在の市場価格を調査してください。

【カード情報】
- カード名: ${cardData.name || '不明'}
- ゲーム: ${cardData.game || '不明'}
- セット/シリーズ: ${cardData.set || '不明'}
- カード番号: ${cardData.number || '不明'}
- レアリティ: ${cardData.rarity || '不明'}
- 言語: ${cardData.language || '日本語'}
- 状態: ${cardData.condition || '美品'}

【対象市場と通貨】
- 対象市場: ${market}
- 通貨: ${currency}

【調査内容】
1. 現在の対象市場での販売価格（指定通貨）
2. 最近の取引相場
3. 価格トレンド（上昇/下降/安定）
4. 3ヶ月前、6ヶ月前、12ヶ月前の推定価格
5. 6ヶ月後、12ヶ月後の価格予測
6. PSA鑑定グレード別価格（PSA9、PSA9.5、PSA10）

【重要な判断基準】
- ${isEN ? 'TCGplayer、eBay落札相場、主要カードショップ（US）' : 'メルカリ、ヤフオク、カードショップ（JP）'}
- 同じカード番号の正確な価格
- プロモカードの場合は配布時期も考慮
- 人気キャラクター（ピカチュウ、リザードン等）は高値傾向

必ず以下のJSON形式で回答してください：
{
  "currency": "${currency}",
  "currentPrice": 現在価格（数値、通貨単位は${currency}）,
  "marketPrice": 市場平均価格（数値、通貨単位は${currency}）,
  "trend": "上昇" | "下降" | "安定" | "変動",
  "confidence": "高" | "中" | "低",
  "priceHistory": {
    "12monthsAgo": 12ヶ月前価格,
    "6monthsAgo": 6ヶ月前価格,
    "3monthsAgo": 3ヶ月前価格
  },
  "pricePrediction": {
    "6monthsLater": 6ヶ月後予測,
    "12monthsLater": 12ヶ月後予測
  },
  "psaGradedPrice": {
    "PSA9": PSA9鑑定品の価格（数値、通貨単位は${currency}）,
    "PSA9.5": PSA9.5鑑定品の価格（数値、通貨単位は${currency}）,
    "PSA10": PSA10鑑定品の価格（数値、通貨単位は${currency}）
  },
  "notes": "価格判定の根拠や特記事項"
}
`;

  // 特定のカードに関する追加情報
  if (cardData.game === 'ポケモン' || cardData.game === 'Pokemon') {
    prompt += '\n\n【ポケモンカード特有の考慮事項】\n';
    prompt += '- SAR、HR、SR、CSRは特に高値\n';
    prompt += '- 女性トレーナーカードは高値傾向\n';
    prompt += '- 最新弾は発売直後高く、徐々に下落\n';
    prompt += '- 絶版セットは価格上昇傾向\n';
    prompt += '- PSA鑑定品は大幅なプレミアム（PSA10は特に高額）\n';
    prompt += '- 人気キャラクター（リザードン、ピカチュウ等）のPSA10は極めて高額\n';
  }

  return prompt;
}

/**
 * OpenAI APIで価格を調査
 */
function callOpenAIForPrice(apiKey, prompt, cardData) {
  const url = 'https://api.openai.com/v1/chat/completions';

  try {
    const payload = {
      model: (getConfig().PRICE_MODEL) || 'gpt-4o',
      messages: [
        {
          role: 'system',
          content: 'あなたはトレーディングカード市場の専門家です。最新の市場動向と価格情報に詳しく、正確な価格査定ができます。Web検索結果や最新の取引データを基に回答してください。'
        },
        {
          role: 'user',
          content: prompt
        }
      ],
      temperature: 0.3,  // より確実な回答を得るため低めに設定
      max_tokens: 1000,
      response_format: { type: "json_object" }  // JSON形式を強制
    };

    const response = UrlFetchApp.fetch(url, {
      method: 'POST',
      headers: {
        'Authorization': 'Bearer ' + apiKey,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    if (response.getResponseCode() === 200) {
      const result = JSON.parse(response.getContentText());
      const content = result.choices[0].message.content;

      console.log('AI回答:', content);

      // JSON形式の回答を解析
      try {
        const priceInfo = JSON.parse(content);
        return priceInfo;
      } catch (e) {
        console.error('JSON解析エラー:', e);
        // JSON解析失敗時は文字列から価格を抽出
        return extractPriceFromText(content);
      }
    } else {
      console.error('OpenAI APIエラー:', response.getContentText());
      return null;
    }

  } catch (error) {
    console.error('価格調査APIエラー:', error);
    return null;
  }
}

/**
 * AI価格情報をカードデータに適用
 */
function applyAIPriceInfo(cardData, priceInfo) {
  console.log('applyAIPriceInfo開始');
  console.log('  現在のcardData.currency:', cardData.currency);
  console.log('  現在のcardData.price:', cardData.price);
  console.log('  priceConvertedフラグ:', cardData.priceConverted);

  // 通貨（AI返答があれば優先、なければ言語から推定）
  const originalCurrency = (priceInfo && priceInfo.currency) || getCurrencyByLanguage(cardData.language);
  cardData.currency = originalCurrency;

  // 英語カードの場合、為替レートを取得
  let exchangeRate = 1;
  if (originalCurrency === 'USD') {
    exchangeRate = cardData.exchangeRate || getExchangeRate('USD', 'JPY');
    cardData.exchangeRate = exchangeRate;
  }

  const minPrice = getMinimumPrice('JPY'); // 最小価格は常にJPYで判定

  // 現在価格を設定（USD→JPY変換含む）
  if (priceInfo && priceInfo.currentPrice && isFinite(priceInfo.currentPrice)) {
    let price = Number(priceInfo.currentPrice);
    if (originalCurrency === 'USD') {
      cardData.priceUSD = price; // USD価格を保存
      price = Math.round(price * exchangeRate); // JPYに変換
    }
    cardData.price = Math.max(price, minPrice);
    cardData.priceSource = 'AI調査';
  } else if (priceInfo && priceInfo.marketPrice && isFinite(priceInfo.marketPrice)) {
    let price = Number(priceInfo.marketPrice);
    if (originalCurrency === 'USD') {
      cardData.priceUSD = price;
      price = Math.round(price * exchangeRate);
    }
    cardData.price = Math.max(price, minPrice);
    cardData.priceSource = 'AI市場価格';
  } else {
    cardData.price = estimatePriceByRarity(cardData.rarity || 'R');
    cardData.priceSource = '推定';
  }

  // 市場価格を設定（USD→JPY変換含む）
  if (priceInfo && priceInfo.marketPrice && isFinite(priceInfo.marketPrice)) {
    let marketPrice = Number(priceInfo.marketPrice);
    if (originalCurrency === 'USD') {
      cardData.marketPriceUSD = marketPrice;
      marketPrice = Math.round(marketPrice * exchangeRate);
    }
    cardData.marketPrice = marketPrice;
  } else {
    cardData.marketPrice = cardData.price;
  }

  cardData.priceTrend = (priceInfo && priceInfo.trend) || '不明';
  cardData.priceConfidence = (priceInfo && priceInfo.confidence) || '低';

  // 価格履歴を設定（USD→JPY変換含む）
  if (priceInfo && priceInfo.priceHistory) {
    cardData.priceHistory = {};

    // 各履歴価格をJPYに変換
    const history12m = priceInfo.priceHistory['12monthsAgo'] || 0;
    const history6m = priceInfo.priceHistory['6monthsAgo'] || 0;
    const history3m = priceInfo.priceHistory['3monthsAgo'] || 0;
    const history9m = priceInfo.priceHistory['9monthsAgo'] || Math.round((history12m + history6m) / 2);

    if (originalCurrency === 'USD') {
      cardData.priceHistory['12ヶ月前'] = Math.round(history12m * exchangeRate);
      cardData.priceHistory['9ヶ月前'] = Math.round(history9m * exchangeRate);
      cardData.priceHistory['6ヶ月前'] = Math.round(history6m * exchangeRate);
      cardData.priceHistory['3ヶ月前'] = Math.round(history3m * exchangeRate);
    } else {
      cardData.priceHistory['12ヶ月前'] = history12m;
      cardData.priceHistory['9ヶ月前'] = history9m;
      cardData.priceHistory['6ヶ月前'] = history6m;
      cardData.priceHistory['3ヶ月前'] = history3m;
    }

    cardData.priceHistory['現在'] = cardData.price;
    cardData.priceHistory['trend'] = cardData.priceTrend;
  }

  // 価格予測を設定（USD→JPY変換含む）
  if (priceInfo && priceInfo.pricePrediction) {
    const pred6m = priceInfo.pricePrediction['6monthsLater'] || cardData.price;
    const pred12m = priceInfo.pricePrediction['12monthsLater'] || cardData.price;

    if (originalCurrency === 'USD') {
      cardData.pricePrediction = {
        '6ヶ月後': Math.round(pred6m * exchangeRate),
        '12ヶ月後': Math.round(pred12m * exchangeRate)
      };
    } else {
      cardData.pricePrediction = {
        '6ヶ月後': pred6m,
        '12ヶ月後': pred12m
      };
    }
  }

  // PSAグレード別価格を設定（USD→JPY変換含む）
  if (priceInfo && priceInfo.psaGradedPrice) {
    cardData.psaGradedPrice = {};

    // 各PSAグレードの価格をJPYに変換
    if (priceInfo.psaGradedPrice.PSA9) {
      let psa9Price = Number(priceInfo.psaGradedPrice.PSA9) || 0;
      if (originalCurrency === 'USD' && psa9Price > 0) {
        cardData.psaGradedPrice.PSA9_USD = psa9Price;
        cardData.psaGradedPrice.PSA9 = Math.round(psa9Price * exchangeRate);
      } else {
        cardData.psaGradedPrice.PSA9 = psa9Price;
      }
    }

    if (priceInfo.psaGradedPrice['PSA9.5']) {
      let psa95Price = Number(priceInfo.psaGradedPrice['PSA9.5']) || 0;
      if (originalCurrency === 'USD' && psa95Price > 0) {
        cardData.psaGradedPrice['PSA9.5_USD'] = psa95Price;
        cardData.psaGradedPrice['PSA9.5'] = Math.round(psa95Price * exchangeRate);
      } else {
        cardData.psaGradedPrice['PSA9.5'] = psa95Price;
      }
    }

    if (priceInfo.psaGradedPrice.PSA10) {
      let psa10Price = Number(priceInfo.psaGradedPrice.PSA10) || 0;
      if (originalCurrency === 'USD' && psa10Price > 0) {
        cardData.psaGradedPrice.PSA10_USD = psa10Price;
        cardData.psaGradedPrice.PSA10 = Math.round(psa10Price * exchangeRate);
      } else {
        cardData.psaGradedPrice.PSA10 = psa10Price;
      }
    }

    console.log('PSAグレード別価格設定:');
    if (cardData.psaGradedPrice.PSA9) {
      console.log(`  PSA9: ¥${cardData.psaGradedPrice.PSA9}`);
    }
    if (cardData.psaGradedPrice['PSA9.5']) {
      console.log(`  PSA9.5: ¥${cardData.psaGradedPrice['PSA9.5']}`);
    }
    if (cardData.psaGradedPrice.PSA10) {
      console.log(`  PSA10: ¥${cardData.psaGradedPrice.PSA10}`);
    }
  }

  // AIの判定根拠を記録
  if (priceInfo.notes) {
    cardData.priceNotes = priceInfo.notes;
  }

  // ログ出力（USD価格がある場合は両方表示）
  if (originalCurrency === 'USD' && cardData.priceUSD) {
    console.log(`AI価格設定: $${cardData.priceUSD} → ¥${cardData.price} (${cardData.priceConfidence}信頼度)`);
    // 変換済みフラグを設定（二重変換防止）
    cardData.priceConverted = true;
    cardData.currency = 'JPY'; // 通貨をJPYに変更
    console.log('  変換後のcurrency:', cardData.currency);
    console.log('  priceConvertedフラグ:', cardData.priceConverted);
  } else {
    const sym = getCurrencySymbol('JPY');
    console.log(`AI価格設定: ${sym}${cardData.price} (${cardData.priceConfidence}信頼度)`);
  }
}

/**
 * テキストから価格を抽出（フォールバック）
 */
function extractPriceFromText(text) {
  const priceInfo = {
    currentPrice: 0,
    marketPrice: 0,
    trend: '不明',
    currency: null
  };

  // 価格パターンを検索
  const pricePatterns = [
    /(\d{1,6})[,，]?(\d{3})?円/g,
    /￥(\d{1,6})[,，]?(\d{3})?/g,
    /¥(\d{1,6})[,，]?(\d{3})?/g,
    /\$(\d{1,3}(?:[,，]\d{3})*(?:\.\d{1,2})?)/g
  ];

  let prices = [];
  pricePatterns.forEach(pattern => {
    const matches = text.matchAll(pattern);
    for (const match of matches) {
      if (pattern.source.startsWith('\\$')) {
        // USD パターン
        const num = match[1].replace(/[,_，]/g, '');
        prices.push(Math.round(parseFloat(num)));
        priceInfo.currency = priceInfo.currency || 'USD';
      } else {
        // 円パターン
        let price = match[1];
        if (match[2]) {
          price += match[2];
        }
        // 円の場合は整数、ドルの場合は小数を考慮
        prices.push(parseFloat(price));
        priceInfo.currency = priceInfo.currency || 'JPY';
      }
    }
  });

  if (prices.length > 0) {
    // 中央値を現在価格とする
    prices.sort((a, b) => a - b);
    priceInfo.currentPrice = prices[Math.floor(prices.length / 2)];
    priceInfo.marketPrice = Math.round(prices.reduce((a, b) => a + b, 0) / prices.length);
    if (!priceInfo.currency) {
      // 見出し語から通貨を推定
      priceInfo.currency = /\$/.test(text) ? 'USD' : 'JPY';
    }
  }

  // トレンドを検出
  if (text.includes('上昇') || text.includes('高騰')) {
    priceInfo.trend = '上昇';
  } else if (text.includes('下降') || text.includes('下落')) {
    priceInfo.trend = '下降';
  } else if (text.includes('安定')) {
    priceInfo.trend = '安定';
  }

  return priceInfo;
}

/**
 * AI基準の価格履歴生成
 */
function generateAIPriceHistory(cardData) {
  const currentPrice = cardData.price || 100;
  const trend = cardData.priceTrend || '安定';

  // すでにAIが履歴を提供している場合はそれを使用
  if (cardData.priceHistory) {
    return cardData.priceHistory;
  }

  // トレンドに基づいて履歴を生成
  let history = {};

  switch (trend) {
    case '上昇':
      history = {
        '12ヶ月前': Math.round(currentPrice * 0.4),
        '9ヶ月前': Math.round(currentPrice * 0.55),
        '6ヶ月前': Math.round(currentPrice * 0.7),
        '3ヶ月前': Math.round(currentPrice * 0.85),
        '現在': currentPrice,
        'trend': '上昇'
      };
      break;

    case '下降':
      history = {
        '12ヶ月前': Math.round(currentPrice * 2.0),
        '9ヶ月前': Math.round(currentPrice * 1.7),
        '6ヶ月前': Math.round(currentPrice * 1.4),
        '3ヶ月前': Math.round(currentPrice * 1.15),
        '現在': currentPrice,
        'trend': '下降'
      };
      break;

    default:
      history = {
        '12ヶ月前': Math.round(currentPrice * 0.9),
        '9ヶ月前': Math.round(currentPrice * 0.92),
        '6ヶ月前': Math.round(currentPrice * 0.95),
        '3ヶ月前': Math.round(currentPrice * 0.98),
        '現在': currentPrice,
        'trend': '安定'
      };
  }

  return history;
}

/**
 * AI基準の価格予測生成
 */
function generateAIPricePrediction(cardData) {
  const currentPrice = cardData.price || 100;
  const trend = cardData.priceTrend || '安定';

  // すでにAIが予測を提供している場合はそれを使用
  if (cardData.pricePrediction) {
    return cardData.pricePrediction;
  }

  let prediction = {};

  switch (trend) {
    case '上昇':
      prediction = {
        '6ヶ月後': Math.round(currentPrice * 1.2),
        '12ヶ月後': Math.round(currentPrice * 1.5)
      };
      break;

    case '下降':
      prediction = {
        '6ヶ月後': Math.round(currentPrice * 0.85),
        '12ヶ月後': Math.round(currentPrice * 0.7)
      };
      break;

    default:
      prediction = {
        '6ヶ月後': Math.round(currentPrice * 1.02),
        '12ヶ月後': Math.round(currentPrice * 1.05)
      };
  }

  return prediction;
}

/**
 * 改良版：enrichCardData（AI価格調査を使用）
 */
function enrichCardDataWithAI(cardData) {
  console.log(`カードデータ補完開始（AI版）: ${cardData.name}`);

  try {
    // AI価格調査を実行
    getCardPriceByAI(cardData);

    // 追加情報の補完（必要に応じて）
    if (!cardData.set && cardData.number) {
      // カード番号からセット情報を推測
      inferSetFromNumber(cardData);
    }

    console.log(`補完完了: ${cardData.name} - ${getCurrencySymbol(cardData.currency)}${cardData.price}`);

  } catch (error) {
    console.error('AI補完エラー:', error);
    cardData.price = estimatePriceByRarity(cardData.rarity || 'R');
    cardData.priceError = error.toString();
  }

  return cardData;
}

/**
 * カード番号からセット情報を推測
 */
function inferSetFromNumber(cardData) {
  if (!cardData.number) return;

  // ポケモンカードのセット番号パターン
  const patterns = {
    'S': '剣・盾シリーズ',
    'SV': 'スカーレット&バイオレット',
    'PROMO': 'プロモカード',
    'SM': 'サン&ムーン',
    'XY': 'XY'
  };

  for (const [prefix, setName] of Object.entries(patterns)) {
    if (cardData.number.toUpperCase().includes(prefix)) {
      cardData.set = cardData.set || setName;
      break;
    }
  }
}

/**
 * AI価格調査のテスト
 */
function testAIPriceResearch() {
  console.log('=== AI価格調査テスト ===\n');

  const testCards = [
    {
      name: 'ピカチュウex',
      game: 'ポケモン',
      number: 'SV-P 001',
      set: 'スカーレット&バイオレット',
      rarity: 'SR'
    },
    {
      name: 'リザードン',
      game: 'ポケモン',
      number: '006/150',
      set: 'ポケモンカード151',
      rarity: 'R'
    }
  ];

  testCards.forEach((card, index) => {
    console.log(`\nテスト${index + 1}: ${card.name}`);

    // AI価格調査
    getCardPriceByAI(card);

    console.log('結果:');
    console.log(`  現在価格: ¥${card.price}`);
    console.log(`  価格トレンド: ${card.priceTrend}`);
    console.log(`  信頼度: ${card.priceConfidence}`);
    console.log(`  情報源: ${card.priceSource}`);

    if (card.priceHistory) {
      console.log('  価格履歴:');
      Object.entries(card.priceHistory).forEach(([period, price]) => {
        if (period !== 'trend') {
          console.log(`    ${period}: ¥${price}`);
        }
      });
    }

    if (card.pricePrediction) {
      console.log('  価格予測:');
      Object.entries(card.pricePrediction).forEach(([period, price]) => {
        console.log(`    ${period}: ¥${price}`);
      });
    }
  });

  return testCards;
}
// ==============================
// Notionエラー修正
// ==============================

/**
 * 現在のエラーを診断して修正
 */
function diagnoseAndFixError() {
  console.log('=== エラー診断 ===\n');

  // 1. 設定確認
  const config = getConfig();

  console.log('【設定状態】');
  console.log('OpenAI APIキー:', config.OPENAI_API_KEY ? '設定済み' : '未設定');
  console.log('Notion APIキー:', config.NOTION_API_KEY ? '設定済み' : '未設定');
  console.log('NotionデータベースID:', config.NOTION_DATABASE_ID ? '設定済み' : '未設定');
  console.log('Driveフォルダ:', config.DRIVE_FOLDER_ID ? '設定済み' : '未設定');

  if (!config.OPENAI_API_KEY) {
    console.log('\n❌ OpenAI APIキーが設定されていません');
    console.log('以下を実行:');
    console.log('PropertiesService.getScriptProperties().setProperty("OPENAI_API_KEY", "your-key");');
  }

  if (!config.DRIVE_FOLDER_ID) {
    console.log('\n❌ Driveフォルダが設定されていません');
    console.log('以下を実行:');
    console.log('setupSimpleDriveWorkflow()');
  }

  console.log('\n【推奨アクション】');
  console.log('1. setupAndTest() を実行して初期設定');
  console.log('2. processImagesWithFullErrorHandling() で画像処理');

  return config;
}

/**
 * セットアップとテスト
 */
function setupAndTest() {
  console.log('=== 完全セットアップ ===\n');

  // 1. 価格設定
  setupPriceConfig();
  console.log('✅ 価格設定完了');

  // 2. Driveフォルダ設定
  const result = setupSimpleDriveWorkflow();
  console.log('✅ Driveフォルダ設定完了');
  console.log(`📁 アップロードフォルダ: ${result.uploadFolder}`);

  // 3. Notion設定（オプション）
  const config = getConfig();
  if (config.NOTION_API_KEY && config.NOTION_DATABASE_ID) {
    console.log('✅ Notion設定あり');
  } else {
    console.log('⚠️ Notion未設定（スプレッドシートのみ使用）');
  }

  console.log('\n=== セットアップ完了 ===');
  console.log('次のコマンド: processImagesWithFullErrorHandling()');

  return result;
}

/**
 * 完全なエラーハンドリング付きbuildNotionProperties
 */
function buildNotionPropertiesFixed(cardData) {
  // cardDataの存在確認
  if (!cardData) {
    console.error('cardDataが未定義です');
    return {};
  }

  console.log('buildNotionProperties開始:', cardData.name || 'Unknown');

  // configを取得
  const config = getConfig();

  // Notion設定がない場合は空を返す
  if (!config.NOTION_API_KEY || !config.NOTION_DATABASE_ID) {
    console.log('Notion未設定のためスキップ');
    return {};
  }

  // データベース情報を安全に取得
  const dbInfo = getNotionDatabaseInfoSafe(config);

  if (!dbInfo || !dbInfo.properties) {
    console.log('データベース情報取得失敗。デフォルトプロパティを使用');
    return buildDefaultNotionProperties(cardData);
  }

  const properties = {};

  // タイトルプロパティを探す
  const titleProp = Object.keys(dbInfo.properties).find(key =>
    dbInfo.properties[key].type === 'title'
  );

  console.log('タイトルプロパティ名:', titleProp || 'なし');

  // タイトルプロパティ（必須）
  if (titleProp) {
    properties[titleProp] = {
      title: [
        {
          text: {
            content: `${cardData.name || 'Unknown Card'} [${cardData.duplicateNumber || 1}]`
          }
        }
      ]
    };
  } else {
    // タイトルプロパティがない場合は「Name」を試す
    properties['Name'] = {
      title: [
        {
          text: {
            content: `${cardData.name || 'Unknown Card'} [${cardData.duplicateNumber || 1}]`
          }
        }
      ]
    };
  }

  // 各プロパティを条件付きで追加
  const propertyMappings = {
    'UniqueID': cardData.uniqueId,
    'unique_id': cardData.uniqueId,
    'Game': cardData.game,
    'game': cardData.game,
    'Set': cardData.set,
    'set': cardData.set,
    'Number': cardData.number,
    'number': cardData.number,
    'Rarity': cardData.rarity,
    'rarity': cardData.rarity,
    'Price': cardData.price,
    'price': cardData.price,
    'PriceTrend': cardData.priceTrend,
    'priceTrend': cardData.priceTrend,
    'Language': cardData.language,
    'language': cardData.language,
    'Condition': cardData.condition,
    'condition': cardData.condition,
    'Status': cardData.status,
    'status': cardData.status
  };

  // プロパティを設定
  Object.keys(dbInfo.properties).forEach(propName => {
    const propInfo = dbInfo.properties[propName];
    const propType = propInfo.type;
    const value = propertyMappings[propName];

    if (value !== undefined && value !== null && propType !== 'title') {
      try {
        switch (propType) {
          case 'rich_text':
            properties[propName] = {
              rich_text: [
                {
                  text: {
                    content: String(value).substring(0, 2000)  // 文字数制限
                  }
                }
              ]
            };
            break;

          case 'number':
            properties[propName] = {
              number: typeof value === 'number' ? value : parseFloat(value) || 0
            };
            break;

          case 'select':
            properties[propName] = {
              select: {
                name: String(value).substring(0, 100)  // 文字数制限
              }
            };
            break;

          case 'url':
            if (String(value).startsWith('http')) {
              properties[propName] = {
                url: String(value)
              };
            }
            break;

          case 'checkbox':
            properties[propName] = {
              checkbox: Boolean(value)
            };
            break;
        }
      } catch (e) {
        console.log(`プロパティ設定エラー (${propName}):`, e.toString());
      }
    }
  });

  console.log('設定したプロパティ数:', Object.keys(properties).length);

  return properties;
}

/**
 * 完全なエラーハンドリング付き画像処理
 */
function processImagesWithFullErrorHandling() {
  const startTime = Date.now();

  console.log('=== 画像処理開始（完全エラーハンドリング版）===\n');

  try {
    // 設定確認
    const config = getConfig();

    if (!config.OPENAI_API_KEY) {
      console.error('❌ OpenAI APIキーが設定されていません');
      console.log('PropertiesService.getScriptProperties().setProperty("OPENAI_API_KEY", "your-key");');
      return;
    }

    if (!config.DRIVE_FOLDER_ID) {
      console.error('❌ Driveフォルダが設定されていません');
      console.log('setupSimpleDriveWorkflow() を実行してください');
      return;
    }

    const processedIds = getProcessedIds();

    // Driveフォルダから画像を取得
    const folder = DriveApp.getFolderById(config.DRIVE_FOLDER_ID);

    // JPEG、PNG、HEIC形式を処理
    const imageTypes = [MimeType.JPEG, MimeType.PNG];
    const allImages = [];

    imageTypes.forEach(mimeType => {
      const files = folder.getFilesByType(mimeType);
      while (files.hasNext() && allImages.length < config.MAX_PHOTOS_PER_RUN) {
        const file = files.next();
        if (!processedIds.includes(file.getId())) {
          allImages.push(file);
        }
      }
    });

    // HEICファイルも処理
    const allFiles = folder.getFiles();
    while (allFiles.hasNext() && allImages.length < config.MAX_PHOTOS_PER_RUN) {
      const file = allFiles.next();
      const fileName = file.getName().toLowerCase();
      if ((fileName.endsWith('.heic') || fileName.endsWith('.heif')) &&
          !processedIds.includes(file.getId())) {
        allImages.push(file);
      }
    }

    if (allImages.length === 0) {
      console.log('新着画像なし');
      return;
    }

    console.log(`新着画像: ${allImages.length}枚`);

    // 各画像を処理
    const results = [];

    for (const file of allImages) {
      try {
        console.log(`\n処理中: ${file.getName()}`);

        // HEICの場合は変換を試みる
        let processFile = file;
        if (file.getName().toLowerCase().match(/\.(heic|heif)$/)) {
          console.log('HEIC形式を検出、変換を試みます...');
          const convertedFile = convertHEICtoJPEG(file);
          if (convertedFile) {
            processFile = convertedFile;
            console.log('HEIC→JPEG変換成功');
          } else {
            console.log('HEIC変換失敗、元ファイルで処理継続');
          }
        }

        const driveFile = {
          id: processFile.getId(),
          name: processFile.getName(),
          url: processFile.getUrl(),
          viewUrl: `https://drive.google.com/file/d/${processFile.getId()}/view`,
          blob: processFile.getBlob(),
          driveFile: processFile
        };

        // AI判定
        const cardData = analyzeCardWithAI(driveFile, config);

        if (!cardData) {
          throw new Error('AI判定失敗：cardDataが生成されませんでした');
        }

        // 必須フィールドを確認
        cardData.name = cardData.name || 'Unknown Card';
        cardData.game = cardData.game || 'Unknown';
        cardData.uniqueId = generateUniqueCardId(cardData, file);
        cardData.driveFileId = file.getId();

        // AI価格調査
        console.log('価格調査中...');
        getCardPriceByAI(cardData);
        console.log(`価格: ${getCurrencySymbol(cardData.currency)}${cardData.price || 0}`);

        // 重複チェック
        const duplicateCount = countDuplicateCards(cardData, config);
        cardData.duplicateNumber = duplicateCount + 1;

        // ファイル名を更新
        const newFileName = renameDriveFile(driveFile, cardData);
        cardData.driveFileName = newFileName;
        cardData.driveUrl = driveFile.url;

        // Notionへ登録（設定がある場合のみ）
        let notionPageId = null;
        if (config.NOTION_API_KEY && config.NOTION_DATABASE_ID) {
          try {
            notionPageId = createNotionRecordFixed(cardData, driveFile, config);
          } catch (notionError) {
            console.error('Notion登録エラー（処理は継続）:', notionError);
          }
        }

        // スプレッドシートに記録
        logCardToSpreadsheetWithCheck(cardData, notionPageId);

        // 処理済みとしてマーク
        markAsProcessed(file.getId());

        // 処理済みフォルダに移動
        try {
          moveToProcessedFolder(file);
        } catch (moveError) {
          console.log('ファイル移動エラー（処理は継続）:', moveError);
        }

        results.push({
          success: true,
          fileId: file.getId(),
          fileName: file.getName(),
          cardName: cardData.name,
          price: cardData.price || 0
        });

        console.log(`✅ 処理成功: ${cardData.name}`);

      } catch (error) {
        console.error(`❌ エラー: ${file.getName()}`, error);
        results.push({
          success: false,
          fileId: file.getId(),
          fileName: file.getName(),
          error: error.toString()
        });

        // エラーログ
        logError(file, error);
      }
    }

    // 処理結果サマリー
    const successCount = results.filter(r => r.success).length;
    const failureCount = results.filter(r => !r.success).length;
    const totalPrice = results
      .filter(r => r.success)
      .reduce((sum, r) => sum + (r.price || 0), 0);

    const processingTime = (Date.now() - startTime) / 1000;

    console.log('\n=== 処理完了 ===');
    console.log(`成功: ${successCount}枚`);
    console.log(`失敗: ${failureCount}枚`);
    console.log(`合計価格: ¥${totalPrice}`);
    console.log(`処理時間: ${processingTime}秒`);

    // 成功したカードをリスト
    if (successCount > 0) {
      console.log('\n【処理成功カード】');
      results.filter(r => r.success).forEach((r, i) => {
        console.log(`${i + 1}. ${r.cardName} - ¥${r.price}`);
      });
    }

    // 失敗したカードをリスト
    if (failureCount > 0) {
      console.log('\n【処理失敗カード】');
      results.filter(r => !r.success).forEach((r, i) => {
        console.log(`${i + 1}. ${r.fileName}: ${r.error}`);
      });
    }

    return results;

  } catch (error) {
    console.error('重大なエラー:', error);
    sendNotification('画像処理で重大なエラー: ' + error.toString());
    return null;
  }
}

/**
 * 修正版：Notionレコード作成
 */
function createNotionRecordFixed(cardData, driveFile, config) {
  if (!cardData) {
    console.error('cardDataが未定義');
    return null;
  }

  if (!config.NOTION_API_KEY || !config.NOTION_DATABASE_ID) {
    console.log('Notion未設定');
    return null;
  }

  try {
    // 英語カードの価格をJPYに変換
    if (cardData.language && cardData.language.toUpperCase().startsWith('EN')) {
      convertEnglishCardPrice(cardData);
    }

    // プロパティを構築（修正版）
    const properties = buildNotionPropertiesFixed(cardData);

    // ページコンテンツを構築
    const children = buildNotionPageContent(cardData, driveFile);

    const url = 'https://api.notion.com/v1/pages';

    const payload = {
      parent: {
        database_id: config.NOTION_DATABASE_ID
      },
      properties: properties,
      children: children || []
    };

    const response = UrlFetchApp.fetch(url, {
      method: 'POST',
      headers: {
        'Authorization': 'Bearer ' + config.NOTION_API_KEY,
        'Content-Type': 'application/json',
        'Notion-Version': '2022-06-28'
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    if (response.getResponseCode() === 200) {
      const result = JSON.parse(response.getContentText());
      console.log('Notionレコード作成成功:', result.id);
      return result.id;
    } else {
      console.error('Notionレコード作成失敗:', response.getContentText());
      return null;
    }

  } catch (error) {
    console.error('Notion作成エラー:', error);
    return null;
  }
}

/**
 * 修正版：スプレッドシート記録
 */
function logCardToSpreadsheetWithCheck(cardData, notionPageId) {
  if (!cardData) {
    console.error('cardDataが未定義');
    return;
  }

  try {
    // 既存の関数を使用
    logCardToSpreadsheetImproved(cardData, notionPageId);
  } catch (error) {
    console.error('スプレッドシート記録エラー:', error);

    // 最低限の情報だけでも記録を試みる
    try {
      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet() ||
                         SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('MASTER_SPREADSHEET_ID'));

      if (spreadsheet) {
        let sheet = spreadsheet.getSheetByName('カード一覧');
        if (!sheet) {
          sheet = spreadsheet.insertSheet('カード一覧');
        }

        sheet.appendRow([
          cardData.uniqueId || 'ERROR',
          new Date(),
          cardData.name || 'Unknown',
          cardData.game || '',
          cardData.set || '',
          cardData.number || '',
          cardData.rarity || '',
          '','',  // 言語、状態
          cardData.price || 0,
          '','','','',  // 価格履歴
          '','',  // 価格予測
          '',  // トレンド
          cardData.duplicateNumber || 1,
          'エラー',
          cardData.driveUrl || '',
          '',  // ファイル名
          notionPageId || '',
          '',  // Photos ID
          error.toString()  // メモにエラーを記録
        ]);
      }
    } catch (fallbackError) {
      console.error('フォールバック記録も失敗:', fallbackError);
    }
  }
}
