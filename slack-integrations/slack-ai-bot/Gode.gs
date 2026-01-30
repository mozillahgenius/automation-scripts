/**
 * 環境変数取得
 */
function Settings() {
  const env = ScriptProperties.getProperties();
  return {
    ...env,
  };
}

/**
 * ログ出力
 */
function Log(title, text) {
  Logger.log(title, text);
}

/**
 * チャンネル情報取得
 */
function getChannelInfo(channelId) {
  const url = 'https://slack.com/api/conversations.info';
  const config = Settings();
  if (!config?.SLACK_TOKEN) return;
  const payload = {
    token: config.SLACK_TOKEN,
    channel: channelId,
  };
  const options = {
    method: 'post',
    payload,
  };
  const response = UrlFetchApp.fetch(url, options);
  const data = JSON.parse(response.getContentText());
  return data.channel;
}

/**
 * メインエントリポイント
 */
function doPost(e) {
  const params = JSON.parse(e.postData.contents);

  // 重複受信防止（3秒超過リトライ対策）
  const cache = CacheService.getScriptCache();
  if (
    params.hasOwnProperty('event') &&
    params.event.hasOwnProperty('client_msg_id')
  ) {
    if (cache.get(params.event.client_msg_id) === 'done') {
      return ContentService.createTextOutput();
    } else {
      cache.put(params.event.client_msg_id, 'done', 600);
    }
  }

  const Bot = new SlackBot(e);

  // メンションイベント用ハンドラ
  const run = ({ event, message, channel, threadTs }) => {
    // ファイルが添付されている場合の処理
    let fileContext = null;
    let documentReviewResult = null;
    
    // ドキュメントレビューのキーワードをチェック
    const isDocumentReview = message && (
      message.includes('修正') || 
      message.includes('レビュー') || 
      message.includes('校正') ||
      message.includes('チェック') ||
      message.includes('改善') ||
      message.includes('コメント')
    );
    
    if (event.files && event.files.length > 0) {
      const file = event.files[0]; // 最初のファイルを処理
      
      // Word/Googleドキュメントの場合でレビューが要求されている場合
      if (isDocumentReview && (
        file.mimetype === 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' ||
        file.mimetype === 'application/msword' ||
        (file.url_private && file.url_private.includes('docs.google.com'))
      )) {
        // ドキュメントレビューモードで処理
        documentReviewResult = processDocumentForReview(file, message, channel);
        
        if (documentReviewResult && documentReviewResult.success) {
          // レビュー結果をSlackに送信
          const reviewMessage = formatDocumentReviewResult(documentReviewResult);
          postMessage(reviewMessage, channel, threadTs || event.ts);
          return; // ここで処理終了
        }
      } else {
        // 通常のファイル読み取り処理
        fileContext = handleMessageWithFiles(event);
      }
    }
    
    // チャンネル説明を取得
    const channelInfo = getChannelInfo(channel);
    const channelDescription =
      channelInfo?.purpose?.value ||
      'チャンネル説明がありません。以下の標準プロンプトを使用します。';

    // ベースロール定義
    const baseRole = [
      { role: 'system', content: 'You are a helpful assistant.' },
      {
        role: 'system',
        content:
          "When asked about the latest, be sure to give the Perplexity search URL link of 'https://www.perplexity.ai/?q='",
      },
      {
        role: 'system',
        content: 'When creating graphs, diagrams, or charts, use Mermaid syntax wrapped in ```mermaid code blocks. Slack will automatically render these as visual diagrams. Example formats: flowchart, sequence diagram, gantt chart, pie chart, etc.'
      },
      {
        role: 'system',
        content:
          'If you have any questions, please ask clearly. Also, never answer by guessing.',
      },
      { role: 'system', content: 'Please answer as concisely as possible.' },
      {
        role: 'system',
        content:
          'If you have a reference source, please list the URL in list form at the end.',
      },
    ];

    // メンション部分を除去して本文を取得
    let text = message;
    const mentionStart = text.indexOf('<@');
    const mentionEnd = text.indexOf('>');
    if (mentionStart === 0 && mentionEnd !== -1) {
      text = text.substring(mentionEnd + 1).trim();
    }

    // 簡易コマンド対応
    if (text === 'Hello') {
      postMessage('Hi there!', channel, event.ts);
      return;
    } else if (text === 'help') {
      postMessage(
        'このボットはSlackでChatGPTに質問を投げられるボットです！',
        channel,
        event.ts
      );
      return;
    }

    // system プロンプトを追加
    const optionRole = [...baseRole];
    optionRole.push(
      {
        role: 'system',
        content: 'Please explain the response results in Japanese.',
      },
      {
        role: 'system',
        content:
          '今から説明するslackチャンネルとしてふさわしい回答を望む。付与する情報を前提として回答してください。FAQのスプレッドシートでの処理が可能な場合にはそのFAQの内容を踏まえて回答しつつ、スプレッドシートのリンク(https://docs.google.com/spreadsheets/d/1MKMjUp2F3r71-VCsT4wVfZo1G6IjEpxobQ7fJKtRPlA/edit?usp=sharing)を掲載して。もしもFAQのシートを参照する必要がなければ、FAQのシートには言及しないで。またFAQのスプレッドシートや参考情報のURLはなるべく自然な形で伝えるようにして。' +
          channelDescription,
      }
    );

    // FAQから追加ロールを取得
    const faq = getFaqRole(text);
    if (faq) optionRole.push(faq);

    // ファイルコンテキストがある場合は追加
    if (fileContext) {
      optionRole.push({
        role: 'system',
        content: `ユーザーが添付したファイルの内容:\n${fileContext.fileContents}\n\nこの内容を考慮して回答してください。`
      });
      // ユーザーメッセージにファイル情報を追加
      text = `${text || ''}\n\n[添付ファイル: ${fileContext.fileInfo.map(f => f.name).join(', ')}]`;
    }

    // スレッド履歴をマージ
    const threadMessages = event.thread_ts
      ? getThreadMessages(channel, event.thread_ts)
      : [];
    mergeRoleAndThread(optionRole, threadMessages);

    // AI レスポンス取得
    let responseText = '';
    if (Settings().GEMINI_TOKEN) {
      responseText = geminiResponse(text, optionRole);
    } else {
      responseText = chatGPTResponse(text, { optionRole });
    }

    // スレッドIDを決定（既存スレッド or 新規スレッドを起こす）
    const replyThread = threadTs || event.ts;

    // メッセージ送信（スレッド内）
    postMessage(responseText, channel, replyThread);
  };

  // メンションされたときのみ処理
  Bot.handleMentionEventBase(run);

  return Bot.response();
}

/**
 * FAQロールを取得
 */
function getFaqRole(question) {
  try {
    const morpths = filterGNL(gNL(question), ['NOUN', 'NUM', 'NUMBER']);
    let words = [];
    for (let i = 0; i < morpths.length; i++) {
      let d = katakanaToHiragana(
        toHalfWidth(morpths[i]).toLowerCase().replace(',', '')
      );
      if (d.indexOf('-')) {
        const arr = morpths[i].split('-');
        for (let n = 0; n < arr.length; n++) {
          words.push(
            katakanaToHiragana(
              toHalfWidth(arr[n]).toLowerCase().replace(',', '')
            )
          );
        }
        continue;
      }
      words.push(d);
    }

    const faqs = SpreadsheetApp.getActive()
      .getSheetByName('faq')
      .getRange('A:B')
      .getValues()
      .filter((row) => !row.every((cell) => cell.trim() === ''));
    let sfaqs = [], result = [];
    for (let i = 1; i < faqs.length; i++) {
      sfaqs[i] = faqs[i].map((cell) =>
        katakanaToHiragana(toHalfWidth(cell).toLowerCase().replace(',', ''))
      );
    }
    for (let i = 1; i < sfaqs.length; i++) {
      if (sfaqs[i].some((faq) => words.some((w) => faq.includes(w)))) {
        if (result.length === 0) result.push(faqs[0]);
        result.push(faqs[i]);
      }
    }
    if (!result.length) return null;
    return {
      role: 'system',
      content:
        '今から記載するJSON形式のFAQを踏まえて回答を望む(FAQの回答とは言わない)' +
        JSON.stringify(result),
    };
  } catch (e) {
    return null;
  }
}

/**
 * スレッド履歴とロールをマージ
 */
function mergeRoleAndThread(optionRole, threadMessages) {
  for (let i = 0; i < threadMessages.length; i++) {
    optionRole.push({
      role: threadMessages[i].hasOwnProperty('app_id') ? 'assistant' : 'user',
      content: threadMessages[i].text || '',
    });
  }
}

/**
 * スレッドメッセージ取得
 */
function getThreadMessages(channelId, threadTs) {
  const url = 'https://slack.com/api/conversations.replies';
  const config = Settings();
  if (!config?.SLACK_TOKEN) return [];
  const payload = {
    token: config.SLACK_TOKEN,
    channel: channelId,
    ts: threadTs,
  };
  const options = {
    method: 'get',
    payload,
  };
  const response = UrlFetchApp.fetch(url, options);
  const data = JSON.parse(response.getContentText());
  return data.messages || [];
}

/**
 * メッセージ送信
 */
function postMessage(message, channel, threadTs = null) {
  const url = 'https://slack.com/api/chat.postMessage';
  const config = Settings();
  if (!config?.SLACK_TOKEN) return;
  const payload = {
    token: config.SLACK_TOKEN,
    channel: channel,
    unfurl_links: true,
    text: message,
    ...(threadTs ? { thread_ts: threadTs } : {}),
  };
  UrlFetchApp.fetch(url, { method: 'post', payload });
}

/**
 * OpenAI Chat Completion 呼び出し
 */
function chatGPTResponse(message, { optionRole = [], temperature }) {
  const config = Settings();
  if (!config?.OPEN_AI_TOKEN) return '';
  const apiKey = config.OPEN_AI_TOKEN;
  const url = 'https://api.openai.com/v1/chat/completions';
  const payload = {
    model: 'o3',
    messages: [...optionRole, { role: 'user', content: message }],
    temperature: temperature ?? 1,
  };
  const options = {
    method: 'post',
    muteHttpExceptions: true,
    headers: {
      Authorization: 'Bearer ' + apiKey,
      'Content-Type': 'application/json',
    },
    payload: JSON.stringify(payload),
  };
  const res = UrlFetchApp.fetch(url, options);
  const result = JSON.parse(res.getContentText());
  return result?.choices?.[0]?.message?.content || '';
}
