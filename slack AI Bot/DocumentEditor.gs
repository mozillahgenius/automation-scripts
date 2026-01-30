/**
 * ドキュメント編集・修正案作成モジュール
 * Word/Googleドキュメントをコピーして修正案を記載し、共有URLを生成
 */

/**
 * ドキュメントを処理して修正案を作成
 */
function processDocumentForReview(file, userMessage, channel) {
  try {
    let docId = null;
    let originalFileName = '';
    
    // ファイルタイプに応じて処理
    if (file.mimetype === 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' || 
        file.mimetype === 'application/msword') {
      // Wordファイルの場合
      const result = convertWordToGoogleDoc(file);
      docId = result.docId;
      originalFileName = file.name;
    } else if (file.url_private && file.url_private.includes('docs.google.com')) {
      // GoogleドキュメントのURLの場合
      const originalDocId = extractDocIdFromUrl(file.url_private);
      docId = copyGoogleDoc(originalDocId);
      originalFileName = file.name || 'Google Document';
    } else {
      throw new Error('このファイルタイプはドキュメント編集に対応していません');
    }
    
    if (!docId) {
      throw new Error('ドキュメントの処理に失敗しました');
    }
    
    // ドキュメントの内容を取得
    const doc = DocumentApp.openById(docId);
    const body = doc.getBody();
    const originalText = body.getText();
    
    // ドキュメント名を設定
    const timestamp = Utilities.formatDate(new Date(), 'JST', 'yyyy-MM-dd_HH:mm');
    doc.setName(`【修正版】${originalFileName}_${timestamp}`);
    
    // AIに修正案を依頼
    const suggestions = getAISuggestions(originalText, userMessage);
    
    // 修正案をドキュメントに適用
    applyAISuggestionsToDocument(doc, suggestions);
    
    // ドキュメントに編集履歴のヘッダーを追加
    addReviewHeader(doc, originalFileName, userMessage);
    
    // 共有設定（元のドキュメントと同じ設定をコピー）
    const originalDocId = file.url_private && file.url_private.includes('docs.google.com') 
      ? extractDocIdFromUrl(file.url_private) 
      : null;
    setDocumentSharing(docId, originalDocId);
    
    // ドキュメントのURLを取得
    const docUrl = doc.getUrl();
    
    return {
      success: true,
      url: docUrl,
      docId: docId,
      fileName: doc.getName(),
      originalFileName: originalFileName,
      suggestionsCount: suggestions.length
    };
    
  } catch (error) {
    Logger.log('ドキュメント処理エラー: ' + error.toString());
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * WordファイルをGoogle Documentに変換
 */
function convertWordToGoogleDoc(file) {
  const config = Settings();
  if (!config?.SLACK_TOKEN) throw new Error('Slack Tokenが設定されていません');
  
  // Slackからファイルをダウンロード
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
  
  // Google Documentに変換
  const resource = {
    title: file.name.replace(/\.(docx?|doc)$/i, ''),
    mimeType: MimeType.GOOGLE_DOCS
  };
  
  const convertedFile = Drive.Files.insert(resource, blob);
  
  return {
    docId: convertedFile.id,
    fileName: convertedFile.title
  };
}

/**
 * Google DocumentのURLからドキュメントIDを抽出
 */
function extractDocIdFromUrl(url) {
  const patterns = [
    /\/document\/d\/([a-zA-Z0-9-_]+)/,
    /id=([a-zA-Z0-9-_]+)/,
    /\/([a-zA-Z0-9-_]+)\/edit/
  ];
  
  for (const pattern of patterns) {
    const match = url.match(pattern);
    if (match && match[1]) {
      return match[1];
    }
  }
  
  throw new Error('Google DocumentのIDを抽出できませんでした');
}

/**
 * Google Documentをコピー
 */
function copyGoogleDoc(originalDocId) {
  try {
    const originalDoc = DriveApp.getFileById(originalDocId);
    const copy = originalDoc.makeCopy();
    return copy.getId();
  } catch (error) {
    throw new Error('ドキュメントのコピーに失敗しました: ' + error.toString());
  }
}

/**
 * AIから修正案を取得
 */
function getAISuggestions(documentText, userRequest) {
  const config = Settings();
  
  // プロンプトを構築
  const systemPrompt = `あなたは優秀な文書校正者です。以下の文書を読んで、ユーザーのリクエストに基づいて修正案を提供してください。
修正案は以下のJSON形式で返してください：
[
  {
    "type": "correction" | "comment" | "addition",
    "originalText": "修正対象の元のテキスト（最大50文字）",
    "suggestion": "修正案または追加するテキスト",
    "comment": "修正理由やコメント",
    "position": "before" | "after" | "replace"
  }
]

重要な注意事項：
- originalTextは必ず元の文書に存在する一意のテキストを指定してください
- 同じテキストが複数ある場合は、より長い文脈を含めて一意になるようにしてください
- commentには修正の理由を簡潔に記載してください
- 修正案は具体的で実装可能なものにしてください`;

  const userPrompt = `以下の文書を確認して、「${userRequest || '全般的な改善提案をお願いします'}」という観点で修正案を提供してください。

文書内容：
${documentText}`;

  let suggestions = [];
  
  if (config?.OPEN_AI_TOKEN) {
    const response = chatGPTResponse(userPrompt, {
      optionRole: [{ role: 'system', content: systemPrompt }]
    });
    
    try {
      // レスポンスからJSON部分を抽出
      const jsonMatch = response.match(/\[[\s\S]*\]/);
      if (jsonMatch) {
        suggestions = JSON.parse(jsonMatch[0]);
      }
    } catch (e) {
      Logger.log('JSON解析エラー: ' + e.toString());
      // フォールバック: 基本的な修正案を作成
      suggestions = [{
        type: 'comment',
        originalText: documentText.substring(0, 50),
        suggestion: '',
        comment: 'AIによる詳細な分析が必要です: ' + response.substring(0, 200),
        position: 'after'
      }];
    }
  }
  
  return suggestions;
}

/**
 * AIの修正案をドキュメントに適用
 */
function applyAISuggestionsToDocument(doc, suggestions) {
  const body = doc.getBody();
  const paragraphs = body.getParagraphs();
  
  // 修正案を適用（逆順で処理して位置ずれを防ぐ）
  suggestions.reverse().forEach(suggestion => {
    try {
      const searchResult = body.findText(suggestion.originalText);
      
      if (searchResult) {
        const element = searchResult.getElement();
        const text = element.asText();
        const startOffset = searchResult.getStartOffset();
        const endOffset = searchResult.getEndOffsetInclusive();
        
        // 修正タイプに応じて処理
        switch (suggestion.type) {
          case 'correction':
            // 取り消し線を追加して修正案を併記
            text.setStrikethrough(startOffset, endOffset, true);
            text.setForegroundColor(startOffset, endOffset, '#FF0000');
            
            // 修正案を追加
            const correctionText = ` [修正案: ${suggestion.suggestion}]`;
            text.insertText(endOffset + 1, correctionText);
            text.setForegroundColor(endOffset + 1, endOffset + correctionText.length, '#0000FF');
            text.setBold(endOffset + 1, endOffset + correctionText.length, true);
            
            // コメントを追加
            if (suggestion.comment) {
              const commentText = ` (理由: ${suggestion.comment})`;
              text.insertText(endOffset + correctionText.length + 1, commentText);
              text.setForegroundColor(endOffset + correctionText.length + 1, 
                                     endOffset + correctionText.length + commentText.length, '#666666');
              text.setItalic(endOffset + correctionText.length + 1, 
                            endOffset + correctionText.length + commentText.length, true);
            }
            break;
            
          case 'comment':
            // コメントを黄色ハイライトで追加
            text.setBackgroundColor(startOffset, endOffset, '#FFFF00');
            const commentOnlyText = ` [コメント: ${suggestion.comment}]`;
            text.insertText(endOffset + 1, commentOnlyText);
            text.setForegroundColor(endOffset + 1, endOffset + commentOnlyText.length, '#008000');
            text.setItalic(endOffset + 1, endOffset + commentOnlyText.length, true);
            break;
            
          case 'addition':
            // 追加提案を緑色で表示
            const additionText = ` [追加提案: ${suggestion.suggestion}]`;
            text.insertText(endOffset + 1, additionText);
            text.setForegroundColor(endOffset + 1, endOffset + additionText.length, '#008000');
            text.setBold(endOffset + 1, endOffset + additionText.length, true);
            text.setBackgroundColor(endOffset + 1, endOffset + additionText.length, '#E8F5E9');
            break;
        }
      } else {
        Logger.log(`テキストが見つかりません: ${suggestion.originalText}`);
      }
    } catch (e) {
      Logger.log(`修正案の適用エラー: ${e.toString()}`);
    }
  });
}

/**
 * ドキュメントにレビューヘッダーを追加
 */
function addReviewHeader(doc, originalFileName, userRequest) {
  const body = doc.getBody();
  const timestamp = Utilities.formatDate(new Date(), 'JST', 'yyyy/MM/dd HH:mm:ss');
  
  // ヘッダーセクションを作成
  const headerParagraph = body.insertParagraph(0, '━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  headerParagraph.setHeading(DocumentApp.ParagraphHeading.NORMAL);
  
  const titleParagraph = body.insertParagraph(1, '📝 AI文書レビュー結果');
  titleParagraph.setHeading(DocumentApp.ParagraphHeading.HEADING1);
  titleParagraph.setBold(true);
  
  body.insertParagraph(2, `元ファイル: ${originalFileName}`);
  body.insertParagraph(3, `レビュー日時: ${timestamp}`);
  body.insertParagraph(4, `レビュー要求: ${userRequest || '全般的な改善'}`);
  
  const legendParagraph = body.insertParagraph(5, '\n凡例:');
  legendParagraph.setBold(true);
  
  body.insertParagraph(6, '• 赤字取り消し線: 修正が必要な箇所');
  body.insertParagraph(7, '• 青字: 修正案');
  body.insertParagraph(8, '• 黄色ハイライト: コメント箇所');
  body.insertParagraph(9, '• 緑字: 追加提案');
  
  body.insertParagraph(10, '━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n');
}

/**
 * ドキュメントの共有設定（元のドキュメントと同じ設定をコピー）
 */
function setDocumentSharing(docId, originalDocId) {
  try {
    const file = DriveApp.getFileById(docId);
    
    // 元のドキュメントが存在する場合は、その共有設定をコピー
    if (originalDocId) {
      try {
        const originalFile = DriveApp.getFileById(originalDocId);
        
        // 元のファイルの共有設定を取得
        const originalAccess = originalFile.getSharingAccess();
        const originalPermission = originalFile.getSharingPermission();
        
        // 同じ共有設定を適用
        file.setSharing(originalAccess, originalPermission);
        
        // 元のファイルのエディターをコピー
        const editors = originalFile.getEditors();
        editors.forEach(editor => {
          try {
            file.addEditor(editor.getEmail());
          } catch (e) {
            Logger.log(`エディター追加エラー: ${editor.getEmail()}`);
          }
        });
        
        // 元のファイルのビューアーをコピー
        const viewers = originalFile.getViewers();
        viewers.forEach(viewer => {
          try {
            file.addViewer(viewer.getEmail());
          } catch (e) {
            Logger.log(`ビューアー追加エラー: ${viewer.getEmail()}`);
          }
        });
        
        Logger.log(`元のドキュメントの共有設定をコピーしました: ${docId}`);
      } catch (e) {
        Logger.log('元のドキュメントの共有設定取得エラー: ' + e.toString());
        // フォールバック: デフォルトの共有設定を適用
        setDefaultSharing(file);
      }
    } else {
      // 元のドキュメントがない場合（Wordファイルの場合など）はデフォルト設定
      setDefaultSharing(file);
    }
    
  } catch (error) {
    Logger.log('共有設定エラー: ' + error.toString());
    throw new Error('ドキュメントの共有設定に失敗しました');
  }
}

/**
 * デフォルトの共有設定を適用
 */
function setDefaultSharing(file) {
  // デフォルト: 組織内のユーザーがリンクで編集可能
  try {
    // まず組織内での共有を試みる
    file.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.EDIT);
    Logger.log('組織内共有設定を適用しました');
  } catch (e) {
    // 組織設定が使えない場合は、リンクを知っている人が編集可能に
    try {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.EDIT);
      Logger.log('リンク共有設定を適用しました');
    } catch (e2) {
      Logger.log('共有設定の適用に失敗しました: ' + e2.toString());
    }
  }
}

/**
 * Slackメッセージ用のフォーマット
 */
function formatDocumentReviewResult(result) {
  if (!result.success) {
    return `❌ ドキュメントの処理中にエラーが発生しました:\n${result.error}`;
  }
  
  const message = `✅ **ドキュメントレビューが完了しました！**

📄 **元ファイル:** ${result.originalFileName}
📝 **修正版ファイル:** ${result.fileName}
💡 **修正提案数:** ${result.suggestionsCount}件

🔗 **編集可能なリンク:**
${result.url}

💬 このリンクから直接ドキュメントを編集できます。
   URLを知っている人は誰でも編集可能です。

📌 修正案の見方:
• 赤字取り消し線: 修正が必要な箇所
• 青字: 修正案
• 黄色ハイライト: コメント箇所
• 緑字: 追加提案`;
  
  return message;
}