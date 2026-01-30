// ==============================
// ポケモンカード管理システム - 統合版
// ==============================

// ==============================
// メインエントリポイント
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
// 設定管理
// ==============================

function getConfig() {
  const scriptProperties = PropertiesService.getScriptProperties();

  return {
    // API Keys
    PERPLEXITY_API_KEY: scriptProperties.getProperty('PERPLEXITY_API_KEY'),
    NOTION_API_KEY: scriptProperties.getProperty('NOTION_API_KEY'),

    // IDs
    NOTION_DATABASE_ID: scriptProperties.getProperty('NOTION_DATABASE_ID'),
    PHOTOS_ALBUM_ID: scriptProperties.getProperty('PHOTOS_ALBUM_ID'),
    DRIVE_FOLDER_ID: scriptProperties.getProperty('DRIVE_FOLDER_ID'),

    // オプション
    USE_EXTERNAL_API: scriptProperties.getProperty('USE_EXTERNAL_API') === 'true',
    NOTIFICATION_EMAIL: scriptProperties.getProperty('NOTIFICATION_EMAIL') || '',

    // 処理設定
    MAX_PHOTOS_PER_RUN: parseInt(scriptProperties.getProperty('MAX_PHOTOS_PER_RUN') || '50'),
    AI_MODEL: scriptProperties.getProperty('AI_MODEL') || 'llava-v1.6-34b'
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
    'AI_MODEL': 'llava-v1.6-34b'
  };

  Object.entries(requiredProperties).forEach(([key, value]) => {
    if (!scriptProperties.getProperty(key)) {
      scriptProperties.setProperty(key, value);
      console.log(`設定追加: ${key}`);
    }
  });

  console.log('初期設定完了');
}

// 初期セットアップ（すべての設定を一度に行う）
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
// AI画像解析（Perplexity API）
// ==============================

function analyzeCardWithAI(driveFile, config) {
  const apiKey = config.PERPLEXITY_API_KEY;
  const model = config.AI_MODEL || 'llava-v1.6-34b';

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
        const result = callPerplexityVision(apiKey, model, imageDataUrl, prompt);
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

        // エクスポネンシャルバックオフ
        const waitTime = Math.pow(2, retryCount) * 1000; // 2秒, 4秒, 8秒
        console.log(`AI判定リトライ ${retryCount}/${maxRetries}, 待機時間: ${waitTime}ms`);
        Utilities.sleep(waitTime);
      }
    }

  } catch (error) {
    console.error('AI判定エラー:', error);

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
          cardData.price = `$${marketPrice}`;
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
        cardData.price = `$${price.tcgplayer_price}`;
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
          cardData.price = `$${price}`;
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

  // Notionページのプロパティを構築
  const properties = buildNotionProperties(cardData);

  const payload = {
    parent: {
      database_id: databaseId
    },
    properties: properties,
    children: buildNotionContent(cardData)
  };

  try {
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
    if (responseCode !== 200) {
      const errorBody = response.getContentText();
      throw new Error(`Notion API Error (${responseCode}): ${errorBody}`);
    }

    const result = JSON.parse(response.getContentText());
    console.log(`Notionレコード作成成功: ${result.id}`);

    return result.id;

  } catch (error) {
    console.error('Notion登録エラー:', error);
    throw error;
  }
}

function buildNotionProperties(cardData) {
  const properties = {
    // タイトル（必須） - ユニークIDを含む
    'Name': {
      title: [
        {
          text: {
            content: `${cardData.name || 'Unknown Card'} [${cardData.duplicateNumber || 1}]`
          }
        }
      ]
    },

    // ユニークID
    'UniqueID': {
      rich_text: [
        {
          text: {
            content: cardData.uniqueId || ''
          }
        }
      ]
    },

    // ゲーム種別
    'Game': {
      select: {
        name: cardData.game || 'Unknown'
      }
    },

    // セット名
    'Set': {
      rich_text: [
        {
          text: {
            content: cardData.set || ''
          }
        }
      ]
    },

    // カード番号
    'Number': {
      rich_text: [
        {
          text: {
            content: cardData.number || ''
          }
        }
      ]
    },

    // レアリティ
    'Rarity': {
      rich_text: [
        {
          text: {
            content: cardData.rarity || ''
          }
        }
      ]
    },

    // 言語
    'Language': {
      select: {
        name: cardData.language || 'Unknown'
      }
    },

    // 状態
    'Condition': {
      rich_text: [
        {
          text: {
            content: cardData.condition || ''
          }
        }
      ]
    },

    // ステータス
    'Status': {
      select: {
        name: cardData.status || '要確認'
      }
    },

    // DriveのURL
    'Source': {
      url: cardData.driveUrl
    },

    // 重複番号
    'DuplicateNumber': {
      number: cardData.duplicateNumber || 1
    },

    // Google Photos ID
    'PhotoID': {
      rich_text: [
        {
          text: {
            content: cardData.photoId || ''
          }
        }
      ]
    }
  };

  // 価格（オプション）
  if (cardData.price) {
    properties['Price'] = {
      rich_text: [
        {
          text: {
            content: cardData.price.toString()
          }
        }
      ]
    };
  }

  // ノート（オプション）
  if (cardData.notes) {
    properties['Notes'] = {
      rich_text: [
        {
          text: {
            content: cardData.notes.substring(0, 2000) // Notionの文字数制限
          }
        }
      ]
    };
  }

  return properties;
}

function buildNotionContent(cardData) {
  const content = [];

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
    ['価格', cardData.price || '-']
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
            content: `登録日時: ${new Date().toLocaleString('ja-JP')}\n`,
            annotations: {
              italic: true,
              color: 'gray'
            }
          }
        },
        {
          text: {
            content: `元ファイル名: ${cardData.originalFileName || 'Unknown'}`,
            annotations: {
              italic: true,
              color: 'gray'
            }
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
    properties: Object.keys(data.properties)
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
    photo.filename,
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
    // 既存のシートがある場合はクリアしてから設定
    cardSheet.clear();
  }

  // ヘッダー行の設定
  const headers = [
    'ユニークID',
    '処理日時',
    'カード名',
    'ゲーム',
    'セット',
    'カード番号',
    'レアリティ',
    '言語',
    '状態',
    '価格',
    '重複番号',
    'ステータス',
    'Google Drive URL',
    'Notion Page ID',
    'Google Photos ID',
    'メモ'
  ];

  const headerRange = cardSheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);

  // ヘッダーのフォーマット設定
  headerRange.setBackground('#4285F4');
  headerRange.setFontColor('#FFFFFF');
  headerRange.setFontWeight('bold');
  headerRange.setHorizontalAlignment('center');

  // 列幅の自動調整
  for (let i = 1; i <= headers.length; i++) {
    cardSheet.autoResizeColumn(i);
  }

  // フィルタービューの設定
  cardSheet.getRange(1, 1, 1000, headers.length).createFilter();

  // 統計シートの作成
  let statsSheet = spreadsheet.getSheetByName('統計');
  if (!statsSheet) {
    statsSheet = spreadsheet.insertSheet('統計');
  } else {
    statsSheet.clear();
  }
  setupStatsSheet(statsSheet);

  // 処理履歴シートの作成
  let historySheet = spreadsheet.getSheetByName('処理履歴');
  if (!historySheet) {
    historySheet = spreadsheet.insertSheet('処理履歴');
  } else {
    historySheet.clear();
  }
  setupHistorySheet(historySheet);

  // 設定シートの作成
  let configSheet = spreadsheet.getSheetByName('設定');
  if (!configSheet) {
    configSheet = spreadsheet.insertSheet('設定');
  } else {
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
    ['ポケモンカード数', '=COUNTIF(カード一覧!D:D,"Pokemon")'],
    ['遊戯王カード数', '=COUNTIF(カード一覧!D:D,"Yu-Gi-Oh!")'],
    ['MTGカード数', '=COUNTIF(カード一覧!D:D,"MTG")'],
    ['重複カード数', '=COUNTIF(カード一覧!K:K,">1")'],
    ['確認済みカード数', '=COUNTIF(カード一覧!L:L,"確定")'],
    ['要確認カード数', '=COUNTIF(カード一覧!L:L,"要確認")'],
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
    ['AI_MODEL', 'llava-v1.6-34b', 'Perplexityモデル名（画像対応）'],
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
    const sheet = spreadsheet.getSheetByName('カード一覧') || spreadsheet.getSheets()[0];

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
      cardData.price || '',
      cardData.duplicateNumber || 1,
      cardData.status || '要確認',
      cardData.driveUrl || '',
      notionPageId || '',
      cardData.photoId || '',
      cardData.notes || ''
    ];

    sheet.appendRow(rowData);

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
// テスト関数
// ==============================

function testConnection() {
  console.log('接続テスト開始');

  const config = getConfig();
  const results = {};

  // Google Photos接続テスト
  try {
    const photos = getPhotosAlbumInfo(config.PHOTOS_ALBUM_ID);
    results.googlePhotos = `✓ アルバム接続成功: ${photos.title}`;
  } catch (error) {
    results.googlePhotos = `✗ アルバム接続失敗: ${error.toString()}`;
  }

  // Google Drive接続テスト
  try {
    const folder = DriveApp.getFolderById(config.DRIVE_FOLDER_ID);
    results.googleDrive = `✓ Drive接続成功: ${folder.getName()}`;
  } catch (error) {
    results.googleDrive = `✗ Drive接続失敗: ${error.toString()}`;
  }

  // Perplexity接続テスト
  try {
    const testResponse = testPerplexityConnection(config.PERPLEXITY_API_KEY);
    results.perplexity = testResponse ? '✓ Perplexity接続成功' : '✗ Perplexity接続失敗';
  } catch (error) {
    results.perplexity = `✗ Perplexity接続失敗: ${error.toString()}`;
  }

  // Notion接続テスト
  try {
    const dbInfo = getNotionDatabaseInfo(config);
    results.notion = `✓ Notion接続成功: ${dbInfo.title}`;
  } catch (error) {
    results.notion = `✗ Notion接続失敗: ${error.toString()}`;
  }

  console.log('接続テスト結果:', results);
  return results;
}

function testPerplexityConnection(apiKey) {
  const url = 'https://api.perplexity.ai/chat/completions';

  try {
    // シンプルなテストリクエスト
    const payload = {
      model: 'llama-3.1-sonar-small-128k-online',
      messages: [
        {
          role: 'user',
          content: 'Hello'
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

    return response.getResponseCode() === 200;
  } catch (error) {
    return false;
  }
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