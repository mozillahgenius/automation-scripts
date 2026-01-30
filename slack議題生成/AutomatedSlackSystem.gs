// ==========================================
// 自動実行版：Slack議題生成＆メッセージ分析システム
// ==========================================

// ========= 設定値 =========
const SPREADSHEET_ID = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID') || '';
const SLACK_BOT_TOKEN = PropertiesService.getScriptProperties().getProperty('SLACK_BOT_TOKEN') || '';
const OPENAI_API_KEY = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY') || '';
const REPORT_EMAIL = PropertiesService.getScriptProperties().getProperty('REPORT_EMAIL') || '';

// パフォーマンス設定
const FETCH_THREAD_REPLIES = true;  // スレッド返信の取得
const MAX_MESSAGES_PER_CHANNEL = 50;  // チャンネルごとの最大取得メッセージ数
const BATCH_SIZE = 100;  // スプレッドシートへの一括書き込みサイズ

// シート名定義
const SHEETS = {
  CONFIG: 'Config',
  SYNC_STATE: 'SyncState',
  MESSAGES: 'Messages',
  CATEGORIES: 'Categories',
  CHECKLISTS: 'Checklists',
  TEMPLATES: 'Templates',
  DRAFTS: 'Drafts',
  LOGS: 'Logs',
  SLACK_LOG: 'slack_log',
  BUSINESS_MANUAL: 'business_manual',
  FAQ_LIST: 'faq_list',
  DAILY_REPORT: 'daily_report'
};

// ========= メイン自動実行関数（時間トリガー用） =========
function mainAutoProcess() {
  console.log('=== 自動処理開始 ===');
  console.log('実行時刻:', new Date().toLocaleString('ja-JP'));
  const startTime = new Date();
  
  try {
    // 1. Slackから最新メッセージを同期
    console.log('\n1. Slackメッセージ同期中...');
    const syncResult = syncBotJoinedChannels();
    
    if (syncResult.messageCount === 0) {
      console.log('新規メッセージなし。処理を終了します。');
      return;
    }
    
    // 2. AI分析を実行
    console.log('\n2. AI分析実行中...');
    runAIAnalysis();
    
    // 3. 議題抽出と業務フロー生成
    console.log('\n3. 議題抽出＆業務フロー生成中...');
    analyzeSlackAndSendReport();
    
    // 4. ガバナンスチェックの実行
    console.log('\n4. ガバナンスチェック実行中...');
    performGovernanceAnalysis();
    
    // 5. 重要な議題を検出して通知
    console.log('\n5. 重要議題の検出と通知...');
    detectAndNotifyImportantTopics();
    
    const endTime = new Date();
    const duration = (endTime - startTime) / 1000;
    console.log(`\n=== 自動処理完了 (処理時間: ${duration}秒) ===`);
    logInfo('自動処理完了', `全処理が正常に完了しました (${duration}秒)`);
    
  } catch (error) {
    console.error('自動処理エラー:', error.toString());
    logError('自動処理', error.toString());
    
    // エラー通知を管理者に送信
    sendErrorNotification(error);
  }
}

// ========= Slack API基本関数 =========
function slackAPI(method, params = {}) {
  const url = `https://slack.com/api/${method}`;
  
  const options = {
    method: 'post',
    headers: {
      'Authorization': `Bearer ${SLACK_BOT_TOKEN}`,
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify(params),
    muteHttpExceptions: true
  };
  
  try {
    const response = UrlFetchApp.fetch(url, options);
    const data = JSON.parse(response.getContentText());
    
    if (!data.ok) {
      console.error(`Slack API Error [${method}]:`, data.error);
      throw new Error(`Slack API Error: ${data.error}`);
    }
    
    return data;
  } catch (error) {
    logError(`Slack API ${method}`, error.toString());
    throw error;
  }
}

// ========= Bot参加チャンネルの同期 =========
function syncBotJoinedChannels() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const syncSheet = getOrCreateSheet(ss, SHEETS.SYNC_STATE, ['チャンネルID', '最終同期タイムスタンプ', '最終同期日時']);
  const slackLogSheet = getOrCreateSheet(ss, SHEETS.SLACK_LOG, [
    'タイムスタンプ', 'チャンネルID', 'チャンネル名', 'ユーザーID',
    'メッセージ', '日時', 'スレッドTS', '返信数', 'permalink'
  ]);
  const messagesSheet = getOrCreateSheet(ss, SHEETS.MESSAGES, [
    'id', 'channel_id', 'user_id', 'timestamp', 'text', 'thread_ts',
    'summary_json', 'classification_json', 'match_flag', 'human_judgement',
    'permalink', 'draft_url', 'processed_at', 'error', 'created_at'
  ]);
  
  let totalMessageCount = 0;
  
  try {
    // 最初にBot情報を取得して権限を確認
    try {
      const authTest = slackAPI('auth.test');
      console.log(`Bot情報: user=${authTest.user}, user_id=${authTest.user_id}, team=${authTest.team}`);
    } catch (e) {
      console.warn('Bot権限情報の取得に失敗:', e);
    }
    
    // Botが参加しているチャンネルを取得
    const result = slackAPI('conversations.list', {
      types: 'public_channel,private_channel',
      exclude_archived: true,
      limit: 1000
    });
    
    const channels = result.channels || [];
    const joinedChannels = channels.filter(ch => ch.is_member);
    
    console.log(`Botが参加しているチャンネル数: ${joinedChannels.length}`);
    
    // チャンネル名をログに出力
    joinedChannels.forEach(ch => {
      console.log(`  - ${ch.name} (${ch.id})`);
    });
    
    // 各チャンネルからメッセージを取得
    for (const channel of joinedChannels) {
      const messageCount = syncChannelMessages(channel.id, channel.name, syncSheet, slackLogSheet, messagesSheet);
      totalMessageCount += messageCount;
    }
    
    return {
      success: true,
      messageCount: totalMessageCount,
      channelCount: joinedChannels.length
    };
    
  } catch (error) {
    console.error('チャンネル同期エラー:', error);
    return {
      success: false,
      error: error.toString(),
      messageCount: 0
    };
  }
}

// ========= 個別チャンネルのメッセージ同期 =========
function syncChannelMessages(channelId, channelName, syncSheet, slackLogSheet, messagesSheet) {
  const lastSyncTs = getLastSyncTime(syncSheet, channelId) || '0';
  const messages = fetchChannelHistory(channelId, lastSyncTs);
  
  if (messages.length === 0) {
    return 0;
  }
  
  console.log(`${channelName}: ${messages.length}件の新規メッセージ`);
  
  // バッチでメッセージを保存
  const messageBatch = [];
  const slackLogBatch = [];
  
  // まずメインメッセージを処理
  for (const message of messages) {
    // メッセージデータの準備
    const messageRow = prepareMessageRow(channelId, message);
    const slackLogRow = prepareSlackLogRow(channelId, channelName, message);
    
    // 有効なデータのみ追加
    if (messageRow) messageBatch.push(messageRow);
    if (slackLogRow) slackLogBatch.push(slackLogRow);
    
    // スレッド返信も取得
    if (FETCH_THREAD_REPLIES && message && message.thread_ts && message.thread_ts === message.ts && message.reply_count > 0) {
      try {
        const replies = fetchThreadReplies(channelId, message.thread_ts);
        // 返信メッセージも処理
        for (const reply of replies) {
          const replyMessageRow = prepareMessageRow(channelId, reply);
          const replySlackLogRow = prepareSlackLogRow(channelId, channelName, reply);
          
          // 有効なデータのみ追加
          if (replyMessageRow) messageBatch.push(replyMessageRow);
          if (replySlackLogRow) slackLogBatch.push(replySlackLogRow);
        }
      } catch (error) {
        console.error(`スレッド ${message.thread_ts} の返信取得エラー:`, error);
      }
    }
  }
  
  // バッチ保存
  if (messageBatch.length > 0) {
    saveMessagesBatch(messagesSheet, messageBatch);
    saveSlackLogBatch(slackLogSheet, slackLogBatch);
  }
  
  // 最終同期時刻を更新（最初のメッセージが最新）
  if (messages.length > 0) {
    const latestTs = messages[0].ts;
    updateLastSyncTime(syncSheet, channelId, latestTs);
  }
  
  return messages.length;
}

// ========= チャンネル履歴取得 =========
function fetchChannelHistory(channelId, oldest = '0') {
  const params = {
    channel: channelId,
    oldest: oldest,
    inclusive: false,
    limit: MAX_MESSAGES_PER_CHANNEL
  };
  
  try {
    const response = slackAPI('conversations.history', params);
    const messages = response.messages || [];
    
    // デバッグ: 取得したメッセージのタイムスタンプを確認
    if (messages.length > 0) {
      const firstMsg = messages[0];
      const now = Date.now() / 1000; // 現在のUNIXタイムスタンプ（秒）
      const msgTime = parseFloat(firstMsg.ts);
      
      // タイムスタンプが未来の場合は警告
      if (msgTime > now) {
        console.warn(`異常なタイムスタンプを検出: ${firstMsg.ts} (現在時刻より未来)`);
        console.warn(`メッセージ内容: ${firstMsg.text?.substring(0, 50)}`);
      }
      
      // タイムスタンプが異常に大きい場合（2030年以降）
      if (msgTime > 1893456000) { // 2030-01-01
        console.error(`異常に大きなタイムスタンプ: ${firstMsg.ts}`);
        return []; // 異常なデータは処理しない
      }
    }
    
    return messages;
  } catch (error) {
    console.error(`チャンネル ${channelId} の履歴取得エラー:`, error);
    return [];
  }
}

// ========= スレッド返信取得 =========
function fetchThreadReplies(channelId, threadTs) {
  // タイムスタンプの検証
  if (!threadTs || typeof threadTs !== 'string' || !threadTs.match(/^\d+\.\d+$/)) {
    console.warn(`無効なスレッドタイムスタンプ: ${threadTs}`);
    return [];
  }
  
  try {
    const response = slackAPI('conversations.replies', {
      channel: channelId,
      ts: threadTs,
      inclusive: false,
      limit: 100
    });
    
    return response.messages || [];
  } catch (error) {
    console.error(`スレッド ${threadTs} の返信取得エラー:`, error);
    return [];
  }
}

// ========= メッセージデータ準備 =========
function prepareMessageRow(channelId, message) {
  // メッセージのバリデーション
  if (!message || !message.ts) {
    console.warn('無効なメッセージオブジェクト:', message);
    return null;
  }
  
  // タイムスタンプの妥当性チェック
  const tsValue = parseFloat(message.ts);
  const now = Date.now() / 1000;
  
  if (tsValue > now + 86400) { // 1日以上未来
    console.warn(`未来のタイムスタンプをスキップ: ${message.ts}`);
    return null;
  }
  
  if (tsValue > 1893456000) { // 2030年以降
    console.warn(`異常なタイムスタンプをスキップ: ${message.ts}`);
    return null;
  }
  
  const messageId = `${channelId}_${message.ts}`;
  const timestamp = new Date(tsValue * 1000);
  
  // Permalinkの取得は一旦スキップ（エラーを避けるため）
  const permalink = ''; // getMessagePermalink(channelId, message.ts);
  
  return [
    messageId,
    channelId,
    message.user || 'bot',
    message.ts,
    message.text || '',
    message.thread_ts || '',
    '', // summary_json
    '', // classification_json
    '', // match_flag
    '', // human_judgement
    permalink,
    '', // draft_url
    '', // processed_at
    '', // error
    timestamp
  ];
}

function prepareSlackLogRow(channelId, channelName, message) {
  // メッセージのバリデーション
  if (!message || !message.ts) {
    console.warn('無効なメッセージオブジェクト:', message);
    return null;
  }
  
  // タイムスタンプの妥当性チェック
  const tsValue = parseFloat(message.ts);
  const now = Date.now() / 1000;
  
  if (tsValue > now + 86400 || tsValue > 1893456000) {
    return null; // 異常なタイムスタンプはスキップ
  }
  
  const timestamp = new Date(tsValue * 1000);
  
  // Permalinkの取得は一旦スキップ（エラーを避けるため）
  const permalink = ''; // getMessagePermalink(channelId, message.ts);
  
  return [
    message.ts,
    channelId,
    channelName,
    message.user || 'bot',
    message.text || '',
    timestamp,
    message.thread_ts || '',
    message.reply_count || 0,
    permalink
  ];
}

// ========= AI分析実行 =========
function runAIAnalysis() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const messagesSheet = ss.getSheetByName(SHEETS.MESSAGES);
  const categoriesSheet = getOrCreateSheet(ss, SHEETS.CATEGORIES, ['カテゴリ名', '説明', 'キーワード', '重要度']);
  
  if (!messagesSheet) {
    console.error('Messagesシートが見つかりません');
    return;
  }
  
  const categories = getCategoriesData(categoriesSheet);
  const unanalyzedMessages = getUnanalyzedMessages(messagesSheet);
  
  console.log(`未分析メッセージ: ${unanalyzedMessages.length}件`);
  
  let processedCount = 0;
  const maxExecutionTime = 4.5 * 60 * 1000; // 4.5分
  const startTime = new Date().getTime();
  
  for (const message of unanalyzedMessages) {
    // 実行時間チェック
    if (new Date().getTime() - startTime > maxExecutionTime) {
      console.log('実行時間制限に近づいたため、処理を中断します');
      break;
    }
    
    try {
      // メッセージを要約
      const summary = summarizeMessage(message.text);
      
      // カテゴリ分類
      const classification = classifyMessage(summary, categories);
      
      // 結果を更新
      updateAnalysisResult(messagesSheet, message.id, {
        summary_json: JSON.stringify(summary),
        classification_json: JSON.stringify(classification),
        match_flag: classification.length > 0 && classification[0].score > 0.7 ? '高' : '低',
        processed_at: new Date()
      });
      
      processedCount++;
      
    } catch (error) {
      console.error(`メッセージ ${message.id} の分析エラー:`, error);
      updateAnalysisResult(messagesSheet, message.id, {
        error: error.toString(),
        processed_at: new Date()
      });
    }
  }
  
  console.log(`AI分析完了: ${processedCount}件処理`);
  logInfo(`AI分析完了: ${processedCount}件のメッセージを分析`);
}

// ========= メッセージ要約（OpenAI） =========
function summarizeMessage(text) {
  if (!text || text.trim() === '') {
    return { summary: '', key_points: [] };
  }
  
  const messages = [
    {
      role: 'system',
      content: '以下のSlackメッセージを簡潔に要約し、重要なポイントを抽出してください。JSON形式で返してください。'
    },
    {
      role: 'user',
      content: `メッセージ: ${text}\n\n以下のJSON形式で返してください:\n{\n  "summary": "要約文",\n  "key_points": ["ポイント1", "ポイント2"],\n  "has_action_item": true/false,\n  "urgency": "high/medium/low"\n}`
    }
  ];
  
  try {
    const result = callOpenAI(messages, 'gpt-5', { type: 'json_object' });
    return result;
  } catch (error) {
    console.error('要約エラー:', error);
    return { summary: text.substring(0, 100), key_points: [], error: error.toString() };
  }
}

// ========= カテゴリ分類 =========
function classifyMessage(summary, categories) {
  const systemPrompt = `あなたは企業のメッセージ分類の専門家です。
以下のカテゴリのいずれかに分類し、関連度スコア（0-1）を付けてください。
カテゴリ: ${categories.map(c => c.name).join(', ')}`;
  
  const userPrompt = `要約: ${summary.summary || ''}\nキーポイント: ${(summary.key_points || []).join(', ')}`;
  
  const messages = [
    { role: 'system', content: systemPrompt },
    {
      role: 'user',
      content: `${userPrompt}\n\n以下のJSON形式で返してください:\n[{"category": "カテゴリ名", "score": 0.8, "reason": "分類理由"}]`
    }
  ];
  
  try {
    const result = callOpenAI(messages, 'gpt-5', { type: 'json_object' });
    // 配列として返されることを期待
    if (Array.isArray(result)) {
      return result;
    } else if (result.classifications) {
      return result.classifications;
    }
    return [];
  } catch (error) {
    console.error('分類エラー:', error);
    return [];
  }
}

// ========= 議題抽出＆業務フロー生成・メール送信 =========
function analyzeSlackAndSendReport() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const slackLogSheet = ss.getSheetByName(SHEETS.SLACK_LOG);
  
  if (!slackLogSheet) {
    console.log('slack_logシートが見つかりません');
    return;
  }
  
  // 過去24時間のメッセージを取得
  const now = new Date();
  const yesterday = new Date(now.getTime() - 24 * 60 * 60 * 1000);
  const recentMessages = getRecentSlackMessages(slackLogSheet, yesterday, now);
  
  if (recentMessages.length === 0) {
    console.log('分析対象のメッセージがありません');
    return;
  }
  
  // メッセージから議題を抽出
  const agendaItems = extractAgendasFromSlackMessages(recentMessages);
  
  if (agendaItems.length === 0) {
    console.log('議題が抽出されませんでした');
    return;
  }
  
  // 業務フローチャートを生成
  const flowchart = generateBusinessFlowchart(agendaItems);
  
  // レポート生成
  const report = {
    date: now,
    messageCount: recentMessages.length,
    agendaItems: agendaItems,
    flowchart: flowchart,
    summary: summarizeAgendaItems(agendaItems)
  };
  
  // メール送信
  if (REPORT_EMAIL) {
    sendAgendaReportWithFlowchart(report);
  }
  
  // スプレッドシートに記録
  recordAgendaAnalysis(ss, report);
  
  console.log(`議題分析完了: ${agendaItems.length}件の議題を抽出`);
}

// ========= 最近のSlackメッセージ取得 =========
function getRecentSlackMessages(slackLogSheet, startDate, endDate) {
  const data = slackLogSheet.getDataRange().getValues();
  const headers = data[0];
  const dateIndex = headers.indexOf('日時');
  
  if (dateIndex === -1) {
    console.error('日時列が見つかりません');
    return [];
  }
  
  const recentMessages = [];
  
  for (let i = 1; i < data.length; i++) {
    const messageDate = new Date(data[i][dateIndex]);
    
    if (messageDate >= startDate && messageDate <= endDate) {
      recentMessages.push({
        timestamp: data[i][0],
        channelId: data[i][1],
        channelName: data[i][2],
        userId: data[i][3],
        text: data[i][4],
        date: messageDate,
        threadTs: data[i][6],
        replyCount: data[i][7],
        permalink: data[i][8]
      });
    }
  }
  
  return recentMessages.sort((a, b) => a.date - b.date);
}

// ========= 議題抽出 =========
function extractAgendasFromSlackMessages(messages) {
  if (messages.length === 0) return [];
  
  // チャンネルごとにメッセージをグループ化
  const messagesByChannel = {};
  messages.forEach(msg => {
    if (!messagesByChannel[msg.channelName]) {
      messagesByChannel[msg.channelName] = [];
    }
    messagesByChannel[msg.channelName].push(msg);
  });
  
  const allAgendaItems = [];
  
  // 各チャンネルのメッセージを分析
  for (const [channelName, channelMessages] of Object.entries(messagesByChannel)) {
    const analysis = analyzeMessagesWithAI(channelName, channelMessages);
    if (analysis.agenda_items && analysis.agenda_items.length > 0) {
      allAgendaItems.push(...analysis.agenda_items);
    }
  }
  
  return allAgendaItems;
}

// ========= AIによるメッセージ分析 =========
function analyzeMessagesWithAI(channelName, messages) {
  const messagesText = messages.map(msg => 
    `[${msg.date.toLocaleString('ja-JP')}] ${msg.text}`
  ).join('\n\n');
  
  const prompt = `チャンネル「${channelName}」の以下のSlackメッセージから、重要な議題・決定事項・課題を抽出してください。

メッセージ:
${messagesText}

以下のJSON形式で返してください:
{
  "agenda_items": [
    {
      "title": "議題タイトル",
      "category": "予算/契約/人事/システム/その他",
      "priority": "高/中/低",
      "summary": "概要",
      "people": ["関係者名"],
      "action_items": ["アクション項目"],
      "deadline": "期限（あれば）",
      "decision_required": true/false
    }
  ]
}`;
  
  const aiMessages = [
    { role: 'system', content: 'あなたは議事録作成と議題抽出の専門家です。' },
    { role: 'user', content: prompt }
  ];
  
  try {
    const result = callOpenAI(aiMessages, 'gpt-5', { type: 'json_object' });
    return result;
  } catch (error) {
    console.error('議題抽出エラー:', error);
    return { agenda_items: [] };
  }
}

// ========= 業務フローチャート生成 =========
function generateBusinessFlowchart(agendaItems) {
  if (agendaItems.length === 0) return null;
  
  const flowPrompt = `以下の議題項目から業務フローチャートを生成してください:
${JSON.stringify(agendaItems, null, 2)}

以下のJSON形式で返してください:
{
  "title": "フローチャートタイトル",
  "steps": [
    {
      "id": "step1",
      "name": "ステップ名",
      "description": "説明",
      "responsible": "担当者/部署",
      "inputs": ["必要な入力"],
      "outputs": ["出力/成果物"],
      "next_steps": ["step2"],
      "decision_point": false
    }
  ],
  "start": "step1",
  "end": ["stepN"]
}`;
  
  const messages = [
    { role: 'system', content: '業務プロセス設計の専門家として、効率的な業務フローを設計してください。' },
    { role: 'user', content: flowPrompt }
  ];
  
  try {
    const result = callOpenAI(messages, 'gpt-5', { type: 'json_object' });
    return result;
  } catch (error) {
    console.error('フローチャート生成エラー:', error);
    return null;
  }
}

// ========= ガバナンス分析 =========
function performGovernanceAnalysis() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const messagesSheet = ss.getSheetByName(SHEETS.MESSAGES);
  
  if (!messagesSheet) {
    console.log('Messagesシートが見つかりません');
    return;
  }
  
  // 重要度の高いメッセージを取得
  const importantMessages = getImportantMessages(messagesSheet);
  
  if (importantMessages.length === 0) {
    console.log('ガバナンス分析対象のメッセージがありません');
    return;
  }
  
  // ガバナンスチェックを実行
  const governanceResults = [];
  
  for (const message of importantMessages) {
    const check = performMessageGovernanceCheck(message);
    if (check.requiresAction) {
      governanceResults.push(check);
    }
  }
  
  // 結果をレポート
  if (governanceResults.length > 0) {
    saveGovernanceReport(ss, governanceResults);
    
    // 重要な案件は通知
    const criticalItems = governanceResults.filter(r => r.severity === 'critical');
    if (criticalItems.length > 0) {
      sendGovernanceAlert(criticalItems);
    }
  }
  
  console.log(`ガバナンス分析完了: ${governanceResults.length}件の要対応事項`);
}

// ========= 重要メッセージの取得 =========
function getImportantMessages(messagesSheet) {
  const data = messagesSheet.getDataRange().getValues();
  const headers = data[0];
  const importantMessages = [];
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const matchFlag = row[headers.indexOf('match_flag')];
    const classificationJson = row[headers.indexOf('classification_json')];
    
    if (matchFlag === '高' && classificationJson) {
      try {
        const classification = JSON.parse(classificationJson);
        if (classification.length > 0 && classification[0].score > 0.7) {
          importantMessages.push({
            id: row[headers.indexOf('id')],
            text: row[headers.indexOf('text')],
            summary: JSON.parse(row[headers.indexOf('summary_json')] || '{}'),
            classification: classification,
            timestamp: row[headers.indexOf('timestamp')]
          });
        }
      } catch (e) {
        // JSONパースエラーは無視
      }
    }
  }
  
  return importantMessages;
}

// ========= 個別メッセージのガバナンスチェック =========
function performMessageGovernanceCheck(message) {
  const checkResult = {
    messageId: message.id,
    requiresAction: false,
    severity: 'low',
    issues: [],
    recommendations: []
  };
  
  const text = message.text.toLowerCase();
  const category = message.classification[0]?.category;
  
  // 開示要件チェック
  const disclosureKeywords = ['決算', '業績', '予想', '修正', '開示', '発表', 'IR'];
  if (disclosureKeywords.some(kw => text.includes(kw))) {
    checkResult.requiresAction = true;
    checkResult.severity = 'high';
    checkResult.issues.push('適時開示要件の可能性');
    checkResult.recommendations.push('IR部門への確認');
  }
  
  // 承認要件チェック
  if (category === '予算' || category === '契約') {
    const amountMatch = text.match(/(\d{1,3}(,\d{3})*|\d+)万円|(\d{1,3}(,\d{3})*|\d+)千円/);
    if (amountMatch) {
      checkResult.requiresAction = true;
      checkResult.severity = 'medium';
      checkResult.issues.push('金額承認が必要な可能性');
      checkResult.recommendations.push('承認権限規程の確認');
    }
  }
  
  // コンプライアンスチェック
  const complianceKeywords = ['違反', '不正', '問題', 'リスク', '監査', '指摘'];
  if (complianceKeywords.some(kw => text.includes(kw))) {
    checkResult.requiresAction = true;
    checkResult.severity = 'critical';
    checkResult.issues.push('コンプライアンス上の懸念');
    checkResult.recommendations.push('法務部門への相談');
    checkResult.recommendations.push('内部監査部門への報告');
  }
  
  return checkResult;
}

// ========= 重要議題の検出と通知 =========
function detectAndNotifyImportantTopics() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const messagesSheet = ss.getSheetByName(SHEETS.MESSAGES);
  const configSheet = ss.getSheetByName(SHEETS.CONFIG);
  
  if (!messagesSheet) {
    console.log('Messagesシートが見つかりません');
    return;
  }
  
  const config = getConfigData(configSheet);
  const data = messagesSheet.getDataRange().getValues();
  
  // 最新24時間以内の重要メッセージを抽出
  const now = new Date();
  const oneDayAgo = new Date(now.getTime() - 24 * 60 * 60 * 1000);
  const importantMessages = [];
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const createdAt = row[14]; // created_at列
    const matchFlag = row[8]; // match_flag列
    const classificationJson = row[7]; // classification_json列
    
    if (createdAt && new Date(createdAt) > oneDayAgo && matchFlag === '高') {
      try {
        const classification = JSON.parse(classificationJson || '[]');
        const highScoreCategory = classification.find(c => c.score >= 0.7);
        
        if (highScoreCategory) {
          importantMessages.push({
            id: row[0],
            channelId: row[1],
            text: row[4],
            summary: JSON.parse(row[6] || '{}'),
            category: highScoreCategory.category,
            score: highScoreCategory.score,
            permalink: row[10],
            createdAt: createdAt
          });
        }
      } catch (e) {
        // JSONパースエラーは無視
      }
    }
  }
  
  // 重要なメッセージがある場合は通知
  if (importantMessages.length > 0) {
    // Slack通知
    if (config.notifySlackChannel) {
      sendSlackNotification(config.notifySlackChannel, importantMessages);
    }
    
    // メール通知
    if (config.notifyEmails && config.notifyEmails.length > 0) {
      sendEmailNotification(config.notifyEmails, importantMessages);
    }
    
    console.log(`${importantMessages.length}件の重要議題を通知`);
  }
}

// ========= OpenAI API呼び出し =========
function callOpenAI(messages, model = 'gpt-5', responseFormat = null) {
  const url = 'https://api.openai.com/v1/responses';
  
  if (!OPENAI_API_KEY) {
    throw new Error('OpenAI APIキーが設定されていません');
  }
  
  // Responses APIではmessagesではなくinputを使用
  const combinedInput = messages.map(m => {
    const content = (typeof m.content === 'string') ? m.content : JSON.stringify(m.content);
    return `${String(m.role || 'user').toUpperCase()}: ${content}`;
  }).join("\n\n");

  const payload = {
    model: model,
    input: combinedInput,
    max_output_tokens: 2000
  };
  
  if (responseFormat && responseFormat.type === 'json_object') {
    payload.text = { format: { type: 'json_schema', name: 'response', schema: { type: 'object', properties: {}, additionalProperties: true }, strict: false } };
  }
  
  const options = {
    method: 'post',
    headers: {
      'Authorization': `Bearer ${OPENAI_API_KEY}`,
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  
  try {
    const response = UrlFetchApp.fetch(url, options);
    const data = JSON.parse(response.getContentText());
    
    if (data.error) {
      throw new Error(`OpenAI API Error: ${data.error.message}`);
    }
    const content = extractTextFromOpenAIResponse_A_(data);
    if (responseFormat && responseFormat.type === 'json_object') {
      try { return JSON.parse(content); } catch (_) { return { error: 'JSON parse error', raw: content || '' }; }
    }
    return content;
  } catch (error) {
    logError('OpenAI API', error.toString());
    throw error;
  }
}

// Responses API テキスト抽出（後方互換）
function extractTextFromOpenAIResponse_A_(data) {
  try {
    if (!data) return '';
    if (typeof data.output_text === 'string' && data.output_text.trim()) return data.output_text;
    if (Array.isArray(data.output)) {
      const parts = [];
      data.output.forEach(block => {
        if (block && Array.isArray(block.content)) {
          block.content.forEach(c => {
            if (typeof c === 'string') parts.push(c);
            else if (c && typeof c.text === 'string') parts.push(c.text);
            else if (c && c.type === 'text' && c.text && c.text.value) parts.push(c.text.value);
          });
        }
      });
      const joined = parts.join('\n').trim();
      if (joined) return joined;
    }
    if (data.choices && data.choices[0]) {
      const ch = data.choices[0];
      if (ch.message && typeof ch.message.content === 'string') return ch.message.content;
      if (typeof ch.text === 'string') return ch.text;
    }
  } catch (e) {
    console.error('extractTextFromOpenAIResponse_A_ error:', e.toString());
  }
  return '';
}

// ========= 通知関数 =========
function sendSlackNotification(channel, messages) {
  const blocks = [
    {
      type: 'header',
      text: {
        type: 'plain_text',
        text: '📋 重要な議題の通知'
      }
    },
    {
      type: 'section',
      text: {
        type: 'mrkdwn',
        text: `*${messages.length}件の重要な議題があります*`
      }
    }
  ];
  
  messages.forEach((msg, index) => {
    blocks.push({
      type: 'divider'
    });
    blocks.push({
      type: 'section',
      text: {
        type: 'mrkdwn',
        text: `*${index + 1}. ${msg.summary.summary || msg.text.substring(0, 50)}*\n` +
               `カテゴリ: ${msg.category}\n` +
               `重要度スコア: ${msg.score}\n` +
               `<${msg.permalink}|元のメッセージを見る>`
      }
    });
  });
  
  slackAPI('chat.postMessage', {
    channel: channel,
    blocks: blocks
  });
}

function sendEmailNotification(emails, messages) {
  const subject = `【重要】Slack議題通知 - ${messages.length}件の重要案件`;
  
  let htmlBody = `
    <h2>重要な議題の通知</h2>
    <p>${messages.length}件の重要な議題が検出されました。</p>
    <hr>
  `;
  
  let plainBody = `重要な議題の通知\n\n${messages.length}件の重要な議題が検出されました。\n\n`;
  
  messages.forEach((msg, index) => {
    htmlBody += `
      <h3>${index + 1}. ${msg.summary.summary || msg.text.substring(0, 50)}</h3>
      <ul>
        <li>カテゴリ: ${msg.category}</li>
        <li>重要度スコア: ${msg.score}</li>
        <li>作成日時: ${new Date(msg.createdAt).toLocaleString('ja-JP')}</li>
        <li><a href="${msg.permalink}">Slackで確認</a></li>
      </ul>
      <hr>
    `;
    
    plainBody += `${index + 1}. ${msg.summary.summary || msg.text.substring(0, 50)}\n`;
    plainBody += `   カテゴリ: ${msg.category}\n`;
    plainBody += `   重要度スコア: ${msg.score}\n`;
    plainBody += `   Slackリンク: ${msg.permalink}\n\n`;
  });
  
  MailApp.sendEmail({
    to: emails.join(','),
    subject: subject,
    body: plainBody,
    htmlBody: htmlBody
  });
}

// ========= レポート送信 =========
function sendAgendaReportWithFlowchart(report) {
  const subject = `Slack議題分析レポート - ${new Date(report.date).toLocaleDateString('ja-JP')}`;
  
  const htmlBody = createAgendaReportHtml(report);
  const plainBody = createAgendaReportPlainText(report);
  
  GmailApp.sendEmail(REPORT_EMAIL, subject, plainBody, {
    htmlBody: htmlBody,
    name: 'Slack議題分析システム'
  });
  
  console.log('議題分析レポートを送信しました');
}

function createAgendaReportHtml(report) {
  let html = `
    <h2>Slack議題分析レポート</h2>
    <p>日付: ${new Date(report.date).toLocaleString('ja-JP')}</p>
    <p>分析メッセージ数: ${report.messageCount}件</p>
    <p>抽出された議題数: ${report.agendaItems.length}件</p>
    
    <h3>議題サマリー</h3>
    <p>${report.summary}</p>
    
    <h3>議題詳細</h3>
  `;
  
  report.agendaItems.forEach((item, index) => {
    html += `
      <h4>${index + 1}. ${item.title}</h4>
      <ul>
        <li>カテゴリ: ${item.category}</li>
        <li>優先度: ${item.priority}</li>
        <li>概要: ${item.summary}</li>
        <li>関係者: ${item.people ? item.people.join(', ') : 'なし'}</li>
        <li>アクション: ${item.action_items ? item.action_items.join(', ') : 'なし'}</li>
        ${item.deadline ? `<li>期限: ${item.deadline}</li>` : ''}
      </ul>
    `;
  });
  
  if (report.flowchart) {
    html += `
      <h3>業務フローチャート</h3>
      <h4>${report.flowchart.title}</h4>
      <p>開始: ${report.flowchart.start} → 終了: ${report.flowchart.end.join(', ')}</p>
    `;
  }
  
  return html;
}

function createAgendaReportPlainText(report) {
  let text = `Slack議題分析レポート
日付: ${new Date(report.date).toLocaleString('ja-JP')}
分析メッセージ数: ${report.messageCount}件
抽出された議題数: ${report.agendaItems.length}件

【議題サマリー】
${report.summary}

【議題詳細】
`;
  
  report.agendaItems.forEach((item, index) => {
    text += `
${index + 1}. ${item.title}
   カテゴリ: ${item.category}
   優先度: ${item.priority}
   概要: ${item.summary}
   関係者: ${item.people ? item.people.join(', ') : 'なし'}
   アクション: ${item.action_items ? item.action_items.join(', ') : 'なし'}
   ${item.deadline ? `期限: ${item.deadline}` : ''}
`;
  });
  
  return text;
}

// ========= ユーティリティ関数 =========
function getOrCreateSheet(spreadsheet, sheetName, headers) {
  let sheet = spreadsheet.getSheetByName(sheetName);
  
  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
    if (headers && headers.length > 0) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    }
  }
  
  return sheet;
}

function getLastSyncTime(sheet, channelId) {
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === channelId) {
      return data[i][1];
    }
  }
  
  return null;
}

function updateLastSyncTime(sheet, channelId, timestamp) {
  const data = sheet.getDataRange().getValues();
  let updated = false;
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === channelId) {
      sheet.getRange(i + 1, 2).setValue(timestamp);
      sheet.getRange(i + 1, 3).setValue(new Date());
      updated = true;
      break;
    }
  }
  
  if (!updated) {
    const lastRow = sheet.getLastRow();
    sheet.getRange(lastRow + 1, 1, 1, 3).setValues([
      [channelId, timestamp, new Date()]
    ]);
  }
}

function saveMessagesBatch(sheet, rows) {
  if (rows.length === 0) return;
  
  const lastRow = sheet.getLastRow();
  sheet.getRange(lastRow + 1, 1, rows.length, rows[0].length).setValues(rows);
}

function saveSlackLogBatch(sheet, rows) {
  if (rows.length === 0) return;
  
  const lastRow = sheet.getLastRow();
  sheet.getRange(lastRow + 1, 1, rows.length, rows[0].length).setValues(rows);
}

function getUnanalyzedMessages(sheet) {
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const messages = [];
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const summaryJson = row[headers.indexOf('summary_json')];
    const text = row[headers.indexOf('text')];
    
    if (!summaryJson && text) {
      messages.push({
        id: row[headers.indexOf('id')],
        text: text,
        rowIndex: i + 1
      });
    }
  }
  
  return messages;
}

function updateAnalysisResult(sheet, messageId, result) {
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][headers.indexOf('id')] === messageId) {
      if (result.summary_json) {
        sheet.getRange(i + 1, headers.indexOf('summary_json') + 1).setValue(result.summary_json);
      }
      if (result.classification_json) {
        sheet.getRange(i + 1, headers.indexOf('classification_json') + 1).setValue(result.classification_json);
      }
      if (result.match_flag) {
        sheet.getRange(i + 1, headers.indexOf('match_flag') + 1).setValue(result.match_flag);
      }
      if (result.processed_at) {
        sheet.getRange(i + 1, headers.indexOf('processed_at') + 1).setValue(result.processed_at);
      }
      if (result.error) {
        sheet.getRange(i + 1, headers.indexOf('error') + 1).setValue(result.error);
      }
      break;
    }
  }
}

function getMessagePermalink(channelId, messageTs) {
  // パラメータの検証
  if (!channelId || !messageTs) {
    console.warn(`無効なパラメータ: channelId=${channelId}, messageTs=${messageTs}`);
    return '';
  }
  
  // タイムスタンプを文字列に変換
  const tsString = String(messageTs);
  
  // タイムスタンプの形式検証（例: 1234567890.123456）
  if (!tsString.match(/^\d+\.\d+$/)) {
    console.warn(`無効なタイムスタンプ形式: ${tsString}`);
    return '';
  }
  
  try {
    const response = slackAPI('chat.getPermalink', {
      channel: channelId,
      message_ts: tsString
    });
    return response.permalink || '';
  } catch (error) {
    // エラーログを削減するため、permalinkの取得失敗は警告レベルに
    console.warn(`Permalink取得失敗 (channel: ${channelId}, ts: ${tsString}): ${error}`);
    return '';
  }
}

function getCategoriesData(sheet) {
  const data = sheet.getDataRange().getValues();
  const categories = [];
  
  // デフォルトカテゴリ
  const defaultCategories = [
    { name: '予算', description: '予算関連の議論', keywords: '予算,予算案,費用,コスト', importance: '高' },
    { name: '契約', description: '契約関連の議論', keywords: '契約,契約書,合意,取引', importance: '高' },
    { name: '人事', description: '人事関連の議論', keywords: '採用,退職,人事,評価', importance: '中' },
    { name: 'システム', description: 'システム関連の議論', keywords: 'システム,開発,バグ,リリース', importance: '中' },
    { name: '営業', description: '営業関連の議論', keywords: '営業,売上,顧客,商談', importance: '高' },
    { name: 'その他', description: 'その他の議論', keywords: '', importance: '低' }
  ];
  
  if (data.length <= 1) {
    // ヘッダーのみの場合はデフォルトカテゴリを設定
    defaultCategories.forEach(cat => {
      const lastRow = sheet.getLastRow();
      sheet.getRange(lastRow + 1, 1, 1, 4).setValues([
        [cat.name, cat.description, cat.keywords, cat.importance]
      ]);
    });
    return defaultCategories;
  }
  
  for (let i = 1; i < data.length; i++) {
    categories.push({
      name: data[i][0],
      description: data[i][1],
      keywords: data[i][2],
      importance: data[i][3]
    });
  }
  
  return categories;
}

function getConfigData(sheet) {
  if (!sheet) {
    return {
      notifySlackChannel: '',
      notifyEmails: []
    };
  }
  
  const data = sheet.getDataRange().getValues();
  const config = {};
  
  for (let i = 1; i < data.length; i++) {
    const key = data[i][0];
    const value = data[i][1];
    
    if (key === 'notify_slack_channel') {
      config.notifySlackChannel = value;
    } else if (key === 'notify_emails') {
      config.notifyEmails = value ? value.split(',').map(e => e.trim()) : [];
    }
  }
  
  return config;
}

function summarizeAgendaItems(agendaItems) {
  if (agendaItems.length === 0) return '議題なし';
  
  const byCategory = {};
  agendaItems.forEach(item => {
    if (!byCategory[item.category]) {
      byCategory[item.category] = 0;
    }
    byCategory[item.category]++;
  });
  
  const highPriority = agendaItems.filter(item => item.priority === '高').length;
  const withDeadline = agendaItems.filter(item => item.deadline).length;
  
  let summary = `総議題数: ${agendaItems.length}件\n`;
  summary += `高優先度: ${highPriority}件\n`;
  summary += `期限付き: ${withDeadline}件\n\n`;
  summary += 'カテゴリ別:\n';
  
  Object.entries(byCategory).forEach(([cat, count]) => {
    summary += `- ${cat}: ${count}件\n`;
  });
  
  return summary;
}

function recordAgendaAnalysis(ss, report) {
  const reportSheet = getOrCreateSheet(ss, SHEETS.DAILY_REPORT, [
    '実行日時', 'メッセージ数', '議題数', 'サマリー', '詳細'
  ]);
  
  const lastRow = reportSheet.getLastRow();
  reportSheet.getRange(lastRow + 1, 1, 1, 5).setValues([[
    report.date,
    report.messageCount,
    report.agendaItems.length,
    report.summary,
    JSON.stringify(report.agendaItems)
  ]]);
}

function saveGovernanceReport(ss, results) {
  const sheet = getOrCreateSheet(ss, 'governance_report', [
    '日時', 'メッセージID', '重要度', '問題', '推奨事項'
  ]);
  
  const rows = results.map(r => [
    new Date(),
    r.messageId,
    r.severity,
    r.issues.join(', '),
    r.recommendations.join(', ')
  ]);
  
  if (rows.length > 0) {
    const lastRow = sheet.getLastRow();
    sheet.getRange(lastRow + 1, 1, rows.length, 5).setValues(rows);
  }
}

function sendGovernanceAlert(criticalItems) {
  if (!REPORT_EMAIL) return;
  
  const subject = '【緊急】ガバナンス上の重要事項の通知';
  let body = `${criticalItems.length}件の重要なガバナンス事項が検出されました。\n\n`;
  
  criticalItems.forEach((item, index) => {
    body += `${index + 1}. メッセージID: ${item.messageId}\n`;
    body += `   問題: ${item.issues.join(', ')}\n`;
    body += `   推奨対応: ${item.recommendations.join(', ')}\n\n`;
  });
  
  MailApp.sendEmail({
    to: REPORT_EMAIL,
    subject: subject,
    body: body
  });
}

function sendErrorNotification(error) {
  if (!REPORT_EMAIL) return;
  
  const subject = '【エラー】Slack議題分析システム エラー通知';
  const body = `自動処理中にエラーが発生しました。\n\n` +
               `エラー内容:\n${error.toString()}\n\n` +
               `発生時刻: ${new Date().toLocaleString('ja-JP')}`;
  
  MailApp.sendEmail({
    to: REPORT_EMAIL,
    subject: subject,
    body: body
  });
}

// ========= ログ関数 =========
function logInfo(message, details = '') {
  console.log(`[INFO] ${message}`, details);
  logToSheet('INFO', message, details);
}

function logError(context, error) {
  console.error(`[ERROR] ${context}:`, error);
  logToSheet('ERROR', context, error);
}

function logToSheet(level, message, details) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const logsSheet = getOrCreateSheet(ss, SHEETS.LOGS, ['タイムスタンプ', 'レベル', 'メッセージ', '詳細']);
    
    const lastRow = logsSheet.getLastRow();
    logsSheet.getRange(lastRow + 1, 1, 1, 4).setValues([[
      new Date(),
      level,
      message,
      details.toString()
    ]]);
  } catch (e) {
    // ログ記録自体のエラーは無視
    console.error('ログ記録エラー:', e);
  }
}

// ========= 時間トリガーの設定 =========
// この関数は手動で一度だけ実行してトリガーを設定
function setupTimeTrigger() {
  // 既存のトリガーを削除
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'mainAutoProcess') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  // 1時間ごとにメイン処理を実行
  ScriptApp.newTrigger('mainAutoProcess')
    .timeBased()
    .everyHours(1)
    .create();
  
  console.log('時間トリガーを設定しました（1時間ごと）');
}

// ========= 初期設定関数（手動実行用） =========
function initialSetup() {
  // スプレッドシートIDの設定
  const spreadsheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
  PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', spreadsheetId);
  
  // 必要なシートの作成
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  getOrCreateSheet(ss, SHEETS.CONFIG, ['設定項目', '値', '説明']);
  getOrCreateSheet(ss, SHEETS.CATEGORIES, ['カテゴリ名', '説明', 'キーワード', '重要度']);
  
  console.log('初期設定完了');
  console.log('次の手順:');
  console.log('1. スクリプトプロパティに以下を設定:');
  console.log('   - SLACK_BOT_TOKEN: Slack Botトークン');
  console.log('   - OPENAI_API_KEY: OpenAI APIキー');
  console.log('   - REPORT_EMAIL: レポート送信先メールアドレス');
  console.log('2. setupTimeTrigger()を実行して時間トリガーを設定');
}
