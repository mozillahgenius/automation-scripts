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
    // タイトル（必須）
    'Name': {
      title: [
        {
          text: {
            content: cardData.name || 'Unknown Card'
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

// ==============================
// Notionデータベース情報取得
// ==============================

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
// Notion検索・重複チェック
// ==============================

function searchNotionRecords(config, query) {
  const notionApiKey = config.NOTION_API_KEY;
  const databaseId = config.NOTION_DATABASE_ID;

  const url = `https://api.notion.com/v1/databases/${databaseId}/query`;

  const payload = {
    filter: query,
    page_size: 100
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

  if (response.getResponseCode() !== 200) {
    throw new Error(`Notion検索失敗: ${response.getContentText()}`);
  }

  const data = JSON.parse(response.getContentText());
  return data.results;
}

function checkDuplicateCard(cardData, config) {
  // カード名と番号で重複チェック
  if (!cardData.name || !cardData.number) {
    return false;
  }

  const filter = {
    and: [
      {
        property: 'Name',
        title: {
          equals: cardData.name
        }
      },
      {
        property: 'Number',
        rich_text: {
          equals: cardData.number
        }
      }
    ]
  };

  try {
    const results = searchNotionRecords(config, filter);
    return results.length > 0;
  } catch (error) {
    console.error('重複チェックエラー:', error);
    return false;
  }
}

// ==============================
// バッチ処理
// ==============================

function batchCreateNotionRecords(cardsData, config) {
  const results = [];

  for (const cardData of cardsData) {
    try {
      // 重複チェック
      if (checkDuplicateCard(cardData, config)) {
        console.log(`重複スキップ: ${cardData.name}`);
        results.push({
          success: false,
          cardData: cardData,
          error: 'Duplicate card'
        });
        continue;
      }

      const pageId = createNotionRecord(cardData, cardData, config);
      results.push({
        success: true,
        cardData: cardData,
        pageId: pageId
      });

    } catch (error) {
      console.error(`Notion登録エラー: ${cardData.name}`, error);
      results.push({
        success: false,
        cardData: cardData,
        error: error.toString()
      });
    }

    // レート制限対策
    Utilities.sleep(500);
  }

  return results;
}

// ==============================
// テスト関数
// ==============================

function testNotionConnection() {
  const config = getConfig();

  console.log('Notion接続テスト開始');

  try {
    const dbInfo = getNotionDatabaseInfo(config);
    console.log('データベース情報:', dbInfo);

    // テストデータ
    const testCard = {
      name: 'テストカード_' + new Date().getTime(),
      game: 'Pokemon',
      set: 'テストセット',
      number: 'TEST-001',
      rarity: 'R',
      language: 'JP',
      condition: '美品',
      price: '¥1,000',
      notes: 'これはテストレコードです',
      status: '要確認',
      driveUrl: 'https://drive.google.com/test',
      driveFileId: 'test_id',
      originalFileName: 'test.jpg'
    };

    const pageId = createNotionRecord(testCard, testCard, config);
    console.log('テストレコード作成成功:', pageId);

    return {
      success: true,
      databaseTitle: dbInfo.title,
      pageId: pageId
    };

  } catch (error) {
    console.error('Notion接続テストエラー:', error);
    return {
      success: false,
      error: error.toString()
    };
  }
}

// ==============================
// Notionページ更新
// ==============================

function updateNotionRecord(pageId, updates, config) {
  const notionApiKey = config.NOTION_API_KEY;
  const url = `https://api.notion.com/v1/pages/${pageId}`;

  const payload = {
    properties: buildNotionProperties(updates)
  };

  const response = UrlFetchApp.fetch(url, {
    method: 'PATCH',
    headers: {
      'Authorization': 'Bearer ' + notionApiKey,
      'Content-Type': 'application/json',
      'Notion-Version': '2022-06-28'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  if (response.getResponseCode() !== 200) {
    throw new Error(`Notion更新失敗: ${response.getContentText()}`);
  }

  return JSON.parse(response.getContentText());
}