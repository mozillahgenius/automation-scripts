/**
 * Slackファイル処理モジュール
 * Word、PDF、Googleドキュメントの読み取り機能
 */

/**
 * Slackイベントからファイル情報を取得して処理
 */
function handleFileShared(event, channel) {
  const files = event.files || [];
  if (files.length === 0) return null;
  
  const results = [];
  
  files.forEach(file => {
    try {
      const fileContent = downloadAndProcessFile(file);
      if (fileContent) {
        results.push({
          name: file.name,
          type: file.mimetype,
          content: fileContent
        });
      }
    } catch (e) {
      Logger.log(`ファイル処理エラー: ${file.name} - ${e.toString()}`);
      results.push({
        name: file.name,
        type: file.mimetype,
        error: `ファイル処理中にエラーが発生しました: ${e.toString()}`
      });
    }
  });
  
  return results;
}

/**
 * Slackからファイルをダウンロードして処理
 */
function downloadAndProcessFile(file) {
  const config = Settings();
  if (!config?.SLACK_TOKEN) return null;
  
  // Slack APIを使用してファイルをダウンロード
  const downloadUrl = file.url_private_download || file.url_private;
  
  const options = {
    method: 'get',
    headers: {
      'Authorization': 'Bearer ' + config.SLACK_TOKEN
    },
    muteHttpExceptions: true
  };
  
  const response = UrlFetchApp.fetch(downloadUrl, options);
  const blob = response.getBlob();
  
  // ファイルタイプに応じて処理
  const mimeType = file.mimetype;
  
  if (mimeType === 'application/pdf') {
    return processPDF(blob);
  } else if (mimeType === 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' || 
             mimeType === 'application/msword') {
    return processWord(blob);
  } else if (mimeType === 'text/plain') {
    return blob.getDataAsString();
  } else if (file.name && file.name.includes('docs.google.com')) {
    // GoogleドキュメントのURLの場合
    return processGoogleDoc(file.url_private);
  } else {
    return `ファイルタイプ ${mimeType} はサポートされていません。`;
  }
}

/**
 * PDFファイルを処理
 */
function processPDF(blob) {
  try {
    // PDFをGoogle Docsに変換して読み取る
    const resource = {
      title: 'temp-pdf-' + Utilities.getUuid(),
      mimeType: MimeType.GOOGLE_DOCS
    };
    
    // Drive APIを使用してPDFをGoogle Docsに変換
    const file = Drive.Files.insert(resource, blob, {
      ocr: true,
      ocrLanguage: 'ja'
    });
    
    // 変換されたドキュメントからテキストを抽出
    const doc = DocumentApp.openById(file.id);
    const text = doc.getBody().getText();
    
    // 一時ファイルを削除
    DriveApp.getFileById(file.id).setTrashed(true);
    
    return text;
  } catch (e) {
    Logger.log('PDF処理エラー: ' + e.toString());
    throw new Error('PDFの読み取りに失敗しました: ' + e.toString());
  }
}

/**
 * Wordファイルを処理
 */
function processWord(blob) {
  try {
    // WordファイルをGoogle Docsに変換
    const resource = {
      title: 'temp-word-' + Utilities.getUuid(),
      mimeType: MimeType.GOOGLE_DOCS
    };
    
    // Drive APIを使用してWordをGoogle Docsに変換
    const file = Drive.Files.insert(resource, blob);
    
    // 変換されたドキュメントからテキストを抽出
    const doc = DocumentApp.openById(file.id);
    const text = doc.getBody().getText();
    
    // 一時ファイルを削除
    DriveApp.getFileById(file.id).setTrashed(true);
    
    return text;
  } catch (e) {
    Logger.log('Word処理エラー: ' + e.toString());
    throw new Error('Wordファイルの読み取りに失敗しました: ' + e.toString());
  }
}

/**
 * GoogleドキュメントのURLから内容を取得
 */
function processGoogleDoc(url) {
  try {
    // URLからドキュメントIDを抽出
    const docIdMatch = url.match(/\/d\/([a-zA-Z0-9-_]+)/);
    if (!docIdMatch) {
      throw new Error('Google ドキュメントのIDが見つかりません');
    }
    
    const docId = docIdMatch[1];
    
    // ドキュメントを開いてテキストを取得
    const doc = DocumentApp.openById(docId);
    const text = doc.getBody().getText();
    
    return text;
  } catch (e) {
    Logger.log('Google Docs処理エラー: ' + e.toString());
    throw new Error('Google ドキュメントの読み取りに失敗しました: ' + e.toString());
  }
}

/**
 * ファイル内容を要約（必要に応じて）
 */
function summarizeFileContent(content, maxLength = 2000) {
  if (content.length <= maxLength) {
    return content;
  }
  
  // 長すぎる場合は先頭と末尾を抽出
  const halfLength = Math.floor(maxLength / 2);
  return content.substring(0, halfLength) + 
         '\n\n[... 中略 ...]\n\n' + 
         content.substring(content.length - halfLength);
}

/**
 * ファイル処理結果をSlackメッセージ用にフォーマット
 */
function formatFileResults(results) {
  if (!results || results.length === 0) {
    return 'ファイルが見つかりませんでした。';
  }
  
  let message = '📎 *添付ファイルの内容:*\n\n';
  
  results.forEach(result => {
    message += `*ファイル名:* ${result.name}\n`;
    message += `*タイプ:* ${result.type}\n`;
    
    if (result.error) {
      message += `⚠️ *エラー:* ${result.error}\n`;
    } else if (result.content) {
      const summary = summarizeFileContent(result.content, 1000);
      message += '```\n' + summary + '\n```\n';
    }
    
    message += '\n---\n\n';
  });
  
  return message;
}

/**
 * ファイル付きメッセージのハンドラ
 */
function handleMessageWithFiles(event) {
  const { text, channel, thread_ts, ts, files } = event;
  
  if (!files || files.length === 0) {
    return null;
  }
  
  // ファイルを処理
  const fileResults = handleFileShared(event, channel);
  
  // ファイル内容を含めて応答を生成
  if (fileResults && fileResults.length > 0) {
    const fileContents = fileResults.map(r => r.content || r.error).join('\n\n');
    
    // ファイル内容を含めたコンテキストを作成
    const context = {
      userMessage: text || 'ファイルが添付されました',
      fileContents: fileContents,
      fileInfo: fileResults.map(r => ({
        name: r.name,
        type: r.type
      }))
    };
    
    return context;
  }
  
  return null;
}