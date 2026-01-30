/**
 * ========================================
 * マニュアル・FAQ生成改良版
 * 小さなタスク単位で独立した文書を生成
 * ========================================
 */

// 既存の設定値
const SLACK_BOT_TOKEN = 'YOUR_SLACK_BOT_TOKEN';
const GOOGLE_DOC_ID = '1dkxrY8mtC28bWyDtxm0NVDohlESzNwqHJqq4PQFimqY';
const LOG_SHEET_NAME = 'slack_log';
const MANUAL_SHEET_NAME = 'business_manual';
const FAQ_SHEET_NAME = 'faq_list';
const OPENAI_API_KEY = 'YOUR_OPENAI_API_KEY';

/**
 * メッセージをトピック別に分類する
 */
function classifyMessagesByTopic(messages) {
  const topics = [];
  let currentTopic = [];
  let lastThreadTs = null;
  
  for (const msg of messages) {
    const threadTs = msg.thread_ts || msg.ts;
    
    // 新しいスレッドまたは時間差が大きい場合は新トピック
    if (lastThreadTs !== threadTs || 
        (lastThreadTs && Math.abs(parseFloat(msg.ts) - parseFloat(lastThreadTs)) > 3600)) {
      
      if (currentTopic.length > 0) {
        topics.push([...currentTopic]);
        currentTopic = [];
      }
    }
    
    currentTopic.push(msg);
    lastThreadTs = threadTs;
  }
  
  // 最後のトピックを追加
  if (currentTopic.length > 0) {
    topics.push(currentTopic);
  }
  
  console.log(`${messages.length}件のメッセージを${topics.length}個のトピックに分類`);
  return topics;
}

/**
 * メッセージのコンテキストを分析
 */
function analyzeMessageContext(messages) {
  const keywords = new Set();
  const participants = new Set();
  let hasQuestion = false;
  let hasDecision = false;
  let hasInstruction = false;
  let hasTroubleshooting = false;
  
  for (const msg of messages) {
    // 参加者を記録
    if (msg.user_name) participants.add(msg.user_name);
    
    // メッセージの特徴を分析
    const text = msg.message.toLowerCase();
    
    // 質問パターン
    if (text.match(/[?？]|どう|なぜ|いつ|どこ|誰|何|how|what|when|where|why|who/)) {
      hasQuestion = true;
    }
    
    // 決定パターン
    if (text.match(/決定|決まり|確定|承認|approved|decided|confirmed/)) {
      hasDecision = true;
    }
    
    // 指示パターン
    if (text.match(/してください|お願い|やって|実行|please|execute|run/)) {
      hasInstruction = true;
    }
    
    // トラブルシューティングパターン
    if (text.match(/エラー|失敗|問題|修正|解決|error|failed|issue|fix|solve/)) {
      hasTroubleshooting = true;
    }
    
    // キーワード抽出（簡易版）
    const words = text.split(/[\s　,、。.!！?？]+/);
    words.forEach(word => {
      if (word.length > 3 && !word.match(/^(です|ます|した|して|これ|それ|あれ|this|that|have|will|would)$/)) {
        keywords.add(word);
      }
    });
  }
  
  return {
    messageCount: messages.length,
    participantCount: participants.size,
    participants: Array.from(participants),
    keywords: Array.from(keywords).slice(0, 10),
    hasQuestion,
    hasDecision,
    hasInstruction,
    hasTroubleshooting,
    estimatedType: determineDocumentType(hasQuestion, hasDecision, hasInstruction, hasTroubleshooting)
  };
}

/**
 * ドキュメントタイプを判定
 */
function determineDocumentType(hasQuestion, hasDecision, hasInstruction, hasTroubleshooting) {
  if (hasTroubleshooting) return 'TROUBLESHOOTING';
  if (hasQuestion && messages.length > 1) return 'FAQ';
  if (hasDecision) return 'DECISION';
  if (hasInstruction) return 'PROCEDURE';
  return 'INFORMATION';
}

/**
 * 改良版：小さな単位でマニュアルを生成
 */
function generateBusinessManualImproved(messages) {
  console.log(`=== 改良版マニュアル生成開始: ${messages.length}件のメッセージ ===`);
  
  if (!OPENAI_API_KEY || OPENAI_API_KEY === '***') {
    console.error('OpenAI APIキーが設定されていません');
    return null;
  }
  
  const sheet = getOrCreateManualSheet();
  const results = [];
  
  // メッセージをトピック別に分類
  const topics = classifyMessagesByTopic(messages);
  console.log(`${topics.length}個の独立したトピックを検出`);
  
  // 各トピックを個別に処理
  for (let i = 0; i < topics.length; i++) {
    const topicMessages = topics[i];
    const context = analyzeMessageContext(topicMessages);
    
    console.log(`トピック ${i + 1}/${topics.length}: ${context.messageCount}件のメッセージ, タイプ: ${context.estimatedType}`);
    
    // メッセージが少なすぎる場合はスキップ
    if (topicMessages.length < 1) continue;
    
    // トピックごとにマニュアルを生成
    const manualItem = generateSingleManualItem(topicMessages, context);
    if (manualItem) {
      results.push(manualItem);
      saveManualToSheet(sheet, manualItem, topicMessages);
    }
    
    // API制限対策のため少し待機
    Utilities.sleep(500);
  }
  
  console.log(`生成完了: ${results.length}件のマニュアル項目`);
  return results;
}

/**
 * 単一のマニュアル項目を生成
 */
function generateSingleManualItem(messages, context) {
  const conversationText = formatMessagesForAI(messages);
  
  // コンテキストに応じたプロンプトを生成
  const prompt = createContextAwarePrompt(conversationText, context);
  
  try {
    const url = 'https://api.openai.com/v1/responses';
    const response = UrlFetchApp.fetch(url, {
      method: 'post',
      headers: {
        'Authorization': 'Bearer ' + OPENAI_API_KEY,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify({
        model: 'gpt-5',
        input: `SYSTEM:\nあなたは業務文書作成の専門家です。与えられた会話から、独立した1つの明確なタスクや手順を抽出してください。\n複数の異なるタスクを無理に1つにまとめないでください。最も重要な1つのポイントに焦点を当ててください。\n\nUSER:\n${prompt}`,
        temperature: 0.3,
        max_output_tokens: 2000
      })
    });
    
    const data = JSON.parse(response.getContentText());
    const content = extractTextFromOpenAIResponse(data);
    
    return parseManualContent(content, context);
    
  } catch (error) {
    console.error('マニュアル生成エラー:', error);
    return null;
  }
}

/**
 * コンテキストに応じたプロンプトを生成
 */
function createContextAwarePrompt(conversationText, context) {
  let promptType = '';
  
  switch (context.estimatedType) {
    case 'TROUBLESHOOTING':
      promptType = `
この会話から、具体的な問題と解決方法を1つ抽出してください。
複数の問題がある場合は、最も重要な1つに絞ってください。

出力形式：
カテゴリ: トラブルシューティング
タイトル: [具体的な問題]
問題の症状: [具体的な症状]
原因: [判明した原因]
解決手順:
1. [手順1]
2. [手順2]
...
確認方法: [解決を確認する方法]
予防策: [再発防止策]`;
      break;
      
    case 'FAQ':
      promptType = `
この会話から、最も重要な質問と回答を1つ抽出してください。
複数の質問がある場合は、それぞれ独立して扱い、ここでは1つだけ出力してください。

出力形式：
カテゴリ: FAQ
質問: [明確な質問文]
回答: [簡潔で正確な回答]
補足情報: [必要に応じて]
関連事項: [関連する他の情報]`;
      break;
      
    case 'DECISION':
      promptType = `
この会話から、行われた意思決定を1つ抽出してください。
複数の決定がある場合は、最も重要な1つに絞ってください。

出力形式：
カテゴリ: 意思決定記録
タイトル: [決定事項]
背景: [決定に至った背景]
決定内容: [具体的な決定内容]
理由: [決定の根拠]
実行事項: [必要なアクション]
責任者: [担当者または部門]
期限: [実施期限]`;
      break;
      
    case 'PROCEDURE':
      promptType = `
この会話から、具体的な作業手順を1つ抽出してください。
複数の手順がある場合は、最も完結した1つのタスクに絞ってください。

出力形式：
カテゴリ: 作業手順
タイトル: [作業名]
目的: [この作業の目的]
前提条件: [必要な準備や条件]
手順:
1. [手順1]
2. [手順2]
...
確認事項: [完了確認の方法]
注意点: [気をつけるべきこと]`;
      break;
      
    default:
      promptType = `
この会話から、業務に有用な情報を1つ抽出してください。
複数のトピックがある場合は、最も重要な1つに絞ってください。

出力形式：
カテゴリ: [適切なカテゴリ]
タイトル: [内容を表す明確なタイトル]
内容: [詳細な説明]
ポイント: [重要な点]
関連情報: [参考になる情報]`;
  }
  
  return `${promptType}

会話内容：
${conversationText}

注意事項：
- 1つの独立したトピックとして完結させてください
- 無関係な複数のタスクを混ぜないでください
- 具体的で実用的な内容にしてください
- 推測や一般化は最小限にしてください`;
}

/**
 * マニュアルコンテンツをパース
 */
function parseManualContent(content, context) {
  const lines = content.split('\n');
  const manual = {
    category: '',
    title: '',
    content: '',
    keywords: context.keywords.join(', '),
    participants: context.participants.join(', '),
    messageCount: context.messageCount
  };
  
  let currentSection = '';
  
  for (const line of lines) {
    if (line.startsWith('カテゴリ:')) {
      manual.category = line.replace('カテゴリ:', '').trim();
    } else if (line.startsWith('タイトル:')) {
      manual.title = line.replace('タイトル:', '').trim();
    } else if (line.startsWith('質問:')) {
      manual.title = '【FAQ】' + line.replace('質問:', '').trim();
      currentSection = 'content';
    } else if (currentSection || (!manual.category && !manual.title)) {
      manual.content += line + '\n';
    }
  }
  
  // 内容が空でなければ返す
  if (manual.title && manual.content) {
    manual.content = manual.content.trim();
    return manual;
  }
  
  return null;
}

/**
 * マニュアルをシートに保存
 */
function saveManualToSheet(sheet, manual, originalMessages) {
  const timestamp = new Date();
  const channelName = originalMessages[0]?.channel_name || '';
  const messageIds = originalMessages.map(m => `${m.channel_id}_${m.ts}`).join(', ');
  
  sheet.appendRow([
    timestamp,
    manual.category || 'その他',
    manual.title,
    manual.content,
    channelName,
    messageIds,
    'アクティブ',
    manual.participants || '',
    manual.keywords || '',
    manual.messageCount || originalMessages.length
  ]);
  
  console.log(`保存: ${manual.title}`);
}

/**
 * 改良版FAQ生成
 */
function generateFAQImproved(messages) {
  console.log(`=== 改良版FAQ生成開始: ${messages.length}件のメッセージ ===`);
  
  if (!OPENAI_API_KEY || OPENAI_API_KEY === '***') {
    console.error('OpenAI APIキーが設定されていません');
    return null;
  }
  
  const sheet = getOrCreateFAQSheet();
  const results = [];
  
  // Q&Aパターンを検出
  const qaPairs = detectQAPairs(messages);
  console.log(`${qaPairs.length}個のQ&Aペアを検出`);
  
  // 各Q&Aペアを個別に処理
  for (const qaPair of qaPairs) {
    const faqItem = generateSingleFAQ(qaPair);
    if (faqItem) {
      results.push(faqItem);
      saveFAQToSheet(sheet, faqItem, qaPair.messages);
    }
    
    // API制限対策
    Utilities.sleep(500);
  }
  
  console.log(`生成完了: ${results.length}件のFAQ`);
  return results;
}

/**
 * Q&Aペアを検出
 */
function detectQAPairs(messages) {
  const pairs = [];
  
  for (let i = 0; i < messages.length; i++) {
    const msg = messages[i];
    const text = msg.message.toLowerCase();
    
    // 質問パターンを検出
    if (text.match(/[?？]|どう|なぜ|いつ|どこ|誰|何/)) {
      // 次の数メッセージを回答候補として収集
      const relatedMessages = [msg];
      const threadTs = msg.thread_ts || msg.ts;
      
      for (let j = i + 1; j < Math.min(i + 10, messages.length); j++) {
        const nextMsg = messages[j];
        
        // 同じスレッドまたは直後のメッセージ
        if (nextMsg.thread_ts === threadTs || 
            (Math.abs(parseFloat(nextMsg.ts) - parseFloat(msg.ts)) < 300)) {
          relatedMessages.push(nextMsg);
        } else {
          break;
        }
      }
      
      // 回答が含まれていそうな場合のみ追加
      if (relatedMessages.length > 1) {
        pairs.push({
          question: msg,
          messages: relatedMessages
        });
        
        // 処理済みメッセージをスキップ
        i += relatedMessages.length - 1;
      }
    }
  }
  
  return pairs;
}

/**
 * 単一のFAQを生成
 */
function generateSingleFAQ(qaPair) {
  const conversationText = formatMessagesForAI(qaPair.messages);
  
  const prompt = `以下の会話から、1つの明確な質問と回答を抽出してください。

出力形式：
質問: [ユーザーの質問を明確に]
回答: [簡潔で分かりやすい回答]
カテゴリ: [適切なカテゴリ]
タグ: [関連キーワード、カンマ区切り]
補足: [必要に応じて追加情報]

会話内容：
${conversationText}

注意：
- 質問と回答は1対1で明確にしてください
- 複数の質問を混ぜないでください
- 回答は実用的で具体的にしてください`;
  
  try {
    const url = 'https://api.openai.com/v1/responses';
    const response = UrlFetchApp.fetch(url, {
      method: 'post',
      headers: {
        'Authorization': 'Bearer ' + OPENAI_API_KEY,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify({
        model: 'gpt-5',
        input: `SYSTEM:\nFAQ作成の専門家として、明確で有用なQ&Aを作成してください。\n\nUSER:\n${prompt}`,
        temperature: 0.3,
        max_output_tokens: 1000
      })
    });
    
    const data = JSON.parse(response.getContentText());
    const content = extractTextFromOpenAIResponse(data);
    
    return parseFAQContent(content);
    
  } catch (error) {
    console.error('FAQ生成エラー:', error);
    return null;
  }
}

/**
 * FAQコンテンツをパース
 */
function parseFAQContent(content) {
  const lines = content.split('\n');
  const faq = {
    question: '',
    answer: '',
    category: '',
    tags: '',
    supplement: ''
  };
  
  for (const line of lines) {
    if (line.startsWith('質問:')) {
      faq.question = line.replace('質問:', '').trim();
    } else if (line.startsWith('回答:')) {
      faq.answer = line.replace('回答:', '').trim();
    } else if (line.startsWith('カテゴリ:')) {
      faq.category = line.replace('カテゴリ:', '').trim();
    } else if (line.startsWith('タグ:')) {
      faq.tags = line.replace('タグ:', '').trim();
    } else if (line.startsWith('補足:')) {
      faq.supplement = line.replace('補足:', '').trim();
    }
  }
  
  // 質問と回答があれば返す
  if (faq.question && faq.answer) {
    return faq;
  }
  
  return null;
}

/**
 * FAQをシートに保存
 */
function saveFAQToSheet(sheet, faq, originalMessages) {
  const timestamp = new Date();
  const channelName = originalMessages[0]?.channel_name || '';
  const messageIds = originalMessages.map(m => `${m.channel_id}_${m.ts}`).join(', ');
  
  const fullAnswer = faq.answer + (faq.supplement ? '\n\n補足: ' + faq.supplement : '');
  
  sheet.appendRow([
    timestamp,
    faq.question,
    fullAnswer,
    faq.category || 'その他',
    faq.tags || '',
    channelName,
    messageIds,
    'アクティブ'
  ]);
  
  console.log(`FAQ保存: ${faq.question.substring(0, 50)}...`);
}

/**
 * メッセージをAI用にフォーマット
 */
function formatMessagesForAI(messages) {
  return messages.map(msg => {
    const time = new Date(Number(msg.ts.split('.')[0]) * 1000);
    const timeStr = Utilities.formatDate(time, 'JST', 'HH:mm');
    return `[${timeStr}] ${msg.user_name || 'Unknown'}: ${msg.message}`;
  }).join('\n');
}

/**
 * OpenAI Responses APIのレスポンスからテキストを抽出（後方互換）
 */
function extractTextFromOpenAIResponse(data) {
  try {
    if (!data) return '';
    if (typeof data.output_text === 'string' && data.output_text.trim()) {
      return data.output_text;
    }
    if (Array.isArray(data.output)) {
      const parts = [];
      data.output.forEach(block => {
        if (block && Array.isArray(block.content)) {
          block.content.forEach(c => {
            if (c && typeof c.text === 'string') parts.push(c.text);
            if (c && c.type === 'text' && c.text && c.text.value) parts.push(c.text.value);
          });
        }
      });
      const joined = parts.join('\n').trim();
      if (joined) return joined;
    }
    if (data.choices && data.choices[0]) {
      const choice = data.choices[0];
      if (choice.message && typeof choice.message.content === 'string') {
        return choice.message.content;
      }
      if (typeof choice.text === 'string') return choice.text;
    }
  } catch (e) {
    console.error('レスポンス解析エラー:', e);
  }
  return '';
}

/**
 * 既存関数との互換性維持
 */
function getOrCreateManualSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(MANUAL_SHEET_NAME);
  
  if (!sheet) {
    sheet = ss.insertSheet(MANUAL_SHEET_NAME);
    const headers = [
      '作成日時', 'カテゴリ', 'タイトル', '内容', 
      '元のチャンネル', '関連メッセージ', 'ステータス',
      '参加者', 'キーワード', 'メッセージ数'
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  
  return sheet;
}

function getOrCreateFAQSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(FAQ_SHEET_NAME);
  
  if (!sheet) {
    sheet = ss.insertSheet(FAQ_SHEET_NAME);
    const headers = [
      '作成日時', '質問', '回答', 'カテゴリ', 'タグ', 
      '元のチャンネル', '関連メッセージ', 'ステータス'
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  
  return sheet;
}

/**
 * メインエントリーポイント：改良版で既存関数を置き換え
 */
function generateManualAndFAQImproved() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LOG_SHEET_NAME);
  if (!sheet) {
    SpreadsheetApp.getUi().alert('エラー', 'slack_logシートが見つかりません', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  // 過去24時間のメッセージを取得
  const now = new Date();
  const yesterday = new Date(now.getTime() - 24 * 60 * 60 * 1000);
  
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const messages = [];
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const date = row[6]; // date列
    
    if (date instanceof Date && date >= yesterday) {
      messages.push({
        channel_id: row[0],
        channel_name: row[1],
        ts: row[2],
        thread_ts: row[3],
        user_name: row[4],
        message: row[5],
        date: row[6]
      });
    }
  }
  
  console.log(`過去24時間のメッセージ: ${messages.length}件`);
  
  if (messages.length === 0) {
    SpreadsheetApp.getUi().alert('情報', '過去24時間にメッセージがありません', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  // チャンネルごとに処理
  const channelMap = {};
  messages.forEach(msg => {
    if (!channelMap[msg.channel_name]) {
      channelMap[msg.channel_name] = [];
    }
    channelMap[msg.channel_name].push(msg);
  });
  
  let totalManuals = 0;
  let totalFAQs = 0;
  
  for (const [channelName, channelMessages] of Object.entries(channelMap)) {
    console.log(`\nチャンネル: ${channelName} (${channelMessages.length}件)`);
    
    // マニュアル生成
    const manuals = generateBusinessManualImproved(channelMessages);
    if (manuals) totalManuals += manuals.length;
    
    // FAQ生成
    const faqs = generateFAQImproved(channelMessages);
    if (faqs) totalFAQs += faqs.length;
  }
  
  const ui = SpreadsheetApp.getUi();
  ui.alert(
    '生成完了',
    `マニュアル: ${totalManuals}件\nFAQ: ${totalFAQs}件\n\n各項目は独立したタスクとして生成されました。`,
    ui.ButtonSet.OK
  );
}
