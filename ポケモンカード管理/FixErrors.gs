// ==============================
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
  console.log('buildNotionProperties開始:', cardData.name);

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
}