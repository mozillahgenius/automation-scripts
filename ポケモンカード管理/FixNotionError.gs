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
        console.log(`価格: ¥${cardData.price || 0}`);

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