// ==========================================
// 整理版：Slack議題生成＆メッセージ分析システム
// ==========================================

// ========= 設定値 =========
const SPREADSHEET_ID = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID') || '';
const SLACK_BOT_TOKEN = PropertiesService.getScriptProperties().getProperty('SLACK_BOT_TOKEN') || '';
const OPENAI_API_KEY = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY') || '';
const REPORT_EMAIL = PropertiesService.getScriptProperties().getProperty('REPORT_EMAIL') || '';

// パフォーマンス設定
const FETCH_THREAD_REPLIES = true;  // スレッド返信の取得を有効化
const MAX_MESSAGES_PER_CHANNEL = 100;  // チャンネルごとの最大取得メッセージ数
const BATCH_SIZE = 100;  // スプレッドシートへの一括書き込みサイズ

// ========= Slack API 基本関数 =========
function slackAPI(method, params = {}) {
  if (!SLACK_BOT_TOKEN || SLACK_BOT_TOKEN === '') {
    const errorMsg = 'Slack Bot Tokenが設定されていません。';
    console.error(errorMsg);
    throw new Error(errorMsg);
  }

  const url = `https://slack.com/api/${method}`;
  const options = {
    method: 'post',
    contentType: 'application/x-www-form-urlencoded',
    headers: {
      'Authorization': 'Bearer ' + SLACK_BOT_TOKEN
    },
    payload: params,
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const responseText = response.getContentText();
    
    if (!responseText) {
      throw new Error('Slack APIから空のレスポンスが返されました');
    }
    
    const data = JSON.parse(responseText);
    
    if (!data.ok) {
      console.error(`Slack APIエラー: ${data.error}`);
      throw new Error(`Slack API Error: ${data.error}`);
    }
    
    return data;
  } catch (error) {
    console.error(`Slack API呼び出しエラー: ${error.toString()}`);
    throw error;
  }
}

// ========= メッセージ取得（アプリ統合） =========
function getMessagesAsApp() {
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.prompt(
    'アプリ統合でメッセージ取得',
    'チャンネルID（例: C09BW2EEVAR）を入力してください:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() !== ui.Button.OK) return;
  
  const channelId = response.getResponseText().trim();
  
  if (!channelId || !channelId.startsWith('C')) {
    ui.alert('エラー', 'チャンネルIDは「C」で始まる必要があります。', ui.ButtonSet.OK);
    return;
  }
  
  try {
    // チャンネル参加試行（オプション）
    try {
      slackAPI('conversations.join', { channel: channelId });
    } catch (joinError) {
      console.log('チャンネル参加スキップ（既に参加済みまたは権限なし）');
    }
    
    // メッセージ取得
    const history = slackAPI('conversations.history', {
      channel: channelId,
      limit: MAX_MESSAGES_PER_CHANNEL
    });
    
    const messages = history.messages || [];
    
    if (messages.length === 0) {
      ui.alert('情報', 'メッセージが見つかりませんでした。', ui.ButtonSet.OK);
      return;
    }
    
    // スプレッドシートに保存
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Messages') || ss.insertSheet('Messages');
    saveMessagesToSheet(sheet, channelId, messages);
    
    ui.alert('完了', `${messages.length}件のメッセージを取得しました。`, ui.ButtonSet.OK);
    
  } catch (error) {
    ui.alert('エラー', `処理中にエラーが発生しました:\n${error.toString()}`, ui.ButtonSet.OK);
  }
}

// ========= メッセージ取得＆AI分析 =========
function getMessagesAsAppAndAnalyze() {
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.prompt(
    'アプリ統合で取得＆分析',
    'チャンネルID（例: C09BW2EEVAR）を入力してください:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() !== ui.Button.OK) return;
  
  const channelId = response.getResponseText().trim();
  
  try {
    // メッセージ取得
    const history = slackAPI('conversations.history', {
      channel: channelId,
      limit: MAX_MESSAGES_PER_CHANNEL
    });
    
    const messages = history.messages || [];
    
    if (messages.length === 0) {
      ui.alert('情報', 'メッセージが見つかりませんでした。', ui.ButtonSet.OK);
      return;
    }
    
    // AI分析実行
    const analysisResults = analyzeMessagesWithAI(messages);
    
    // 結果を保存
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Analysis') || ss.insertSheet('Analysis');
    saveAnalysisResults(sheet, analysisResults);
    
    ui.alert('完了', `分析完了:\n- ${messages.length}件のメッセージ\n- ${analysisResults.topics.length}個の議題を抽出`, ui.ButtonSet.OK);
    
  } catch (error) {
    ui.alert('エラー', `処理中にエラーが発生しました:\n${error.toString()}`, ui.ButtonSet.OK);
  }
}

// ========= ガバナンス分析 =========
function getMessagesAsAppWithGovernance() {
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.prompt(
    'ガバナンス分析',
    'チャンネルID（例: C09BW2EEVAR）を入力してください:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() !== ui.Button.OK) return;
  
  const channelId = response.getResponseText().trim();
  
  try {
    // メッセージ取得
    const history = slackAPI('conversations.history', {
      channel: channelId,
      limit: MAX_MESSAGES_PER_CHANNEL
    });
    
    const messages = history.messages || [];
    
    // ガバナンス分析実行
    const governanceResults = analyzeMessagesForGovernance(messages);
    
    // 結果を保存
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const timestamp = new Date().getTime();
    const sheetName = `ガバナンス_${timestamp}`;
    const sheet = ss.insertSheet(sheetName);
    createGovernanceAnalysisSheet(sheet, governanceResults, channelId);
    
    // 通知送信
    if (REPORT_EMAIL) {
      sendGovernanceNotificationEmail(governanceResults, channelId);
    }
    
    ui.alert('完了', `ガバナンス分析完了:\n- リスク項目: ${governanceResults.risks.length}件\n- 承認要件: ${governanceResults.approvals.length}件`, ui.ButtonSet.OK);
    
  } catch (error) {
    ui.alert('エラー', `処理中にエラーが発生しました:\n${error.toString()}`, ui.ButtonSet.OK);
  }
}

// ========= 業務フロー生成＆通知 =========
function getMessagesAsAppWithWorkflow() {
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.prompt(
    '業務フロー生成＆通知',
    'チャンネルID（例: C09BW2EEVAR）を入力してください:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() !== ui.Button.OK) return;
  
  const channelId = response.getResponseText().trim();
  
  try {
    // メッセージ取得
    const history = slackAPI('conversations.history', {
      channel: channelId,
      limit: MAX_MESSAGES_PER_CHANNEL
    });
    
    const messages = history.messages || [];
    
    if (messages.length === 0) {
      ui.alert('情報', 'メッセージが見つかりませんでした。', ui.ButtonSet.OK);
      return;
    }
    
    // タスク抽出と業務フロー生成
    const workflowData = extractTasksAndCreateWorkflow(messages);
    
    // 業務記述書の生成
    const businessSpec = generateBusinessSpecification(workflowData, channelId);
    
    // スプレッドシート作成
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const timestamp = new Date().getTime();
    
    // 業務記述書シート
    const specSheetName = `業務記述書_${timestamp}`;
    const specSheet = ss.insertSheet(specSheetName);
    createBusinessSpecSheet(specSheet, businessSpec);
    
    // 業務フローシート
    const flowSheetName = `業務フロー_${timestamp}`;
    const flowSheet = ss.insertSheet(flowSheetName);
    createWorkflowSheet(flowSheet, workflowData);
    
    // タスク管理シート
    const taskSheetName = `タスク_${timestamp}`;
    const taskSheet = ss.insertSheet(taskSheetName);
    createTaskManagementSheet(taskSheet, workflowData.tasks);
    
    // 通知送信
    const notificationData = {
      channelName: channelId,
      messageCount: messages.length,
      taskCount: workflowData.tasks.length,
      flowSteps: workflowData.flowSteps.length,
      sheets: {
        spec: specSheetName,
        flow: flowSheetName,
        task: taskSheetName
      },
      spreadsheetUrl: ss.getUrl()
    };
    
    if (REPORT_EMAIL) {
      sendWorkflowNotificationEmail(notificationData);
    }
    
    sendWorkflowSlackNotification(notificationData, channelId);
    
    ui.alert('完了', `業務フロー生成完了！\n📊 タスク: ${workflowData.tasks.length}件\n📈 フローステップ: ${workflowData.flowSteps.length}件`, ui.ButtonSet.OK);
    
  } catch (error) {
    ui.alert('エラー', `処理中にエラーが発生しました:\n${error.toString()}`, ui.ButtonSet.OK);
  }
}

// ========= AI分析関数 =========
function analyzeMessagesWithAI(messages) {
  if (!OPENAI_API_KEY) {
    throw new Error('OpenAI APIキーが設定されていません');
  }
  
  const messageText = messages.map(m => m.text || '').join('\n');
  
  const prompt = `
以下のSlackメッセージから重要な議題、論点、カテゴリを抽出してください。

メッセージ:
${messageText}

以下の形式でJSON出力してください:
{
  "topics": [{"title": "議題タイトル", "summary": "要約", "category": "カテゴリ", "priority": "高/中/低"}],
  "summary": "全体要約",
  "categories": ["カテゴリ1", "カテゴリ2"]
}`;
  
  const response = callOpenAI(prompt);
  
  try {
    return JSON.parse(response);
  } catch (e) {
    return {
      topics: [{title: "分析結果", summary: response, category: "その他", priority: "中"}],
      summary: response,
      categories: ["その他"]
    };
  }
}

// ========= ガバナンス分析関数 =========
function analyzeMessagesForGovernance(messages) {
  const messageText = messages.map(m => m.text || '').join('\n');
  
  const prompt = `
以下のSlackメッセージから、ガバナンス・コンプライアンスの観点で重要な項目を抽出してください。

メッセージ:
${messageText}

以下の観点で分析してください:
1. 開示要件（適時開示・決算開示）
2. 承認フローの適切性
3. リスク項目
4. 必要な専門家（弁護士・会計士等）
5. 内部統制のポイント

JSON形式で出力してください:
{
  "disclosures": [{"type": "種別", "content": "内容", "urgency": "緊急度"}],
  "approvals": [{"item": "項目", "level": "承認レベル", "risk": "リスク"}],
  "risks": [{"category": "カテゴリ", "description": "説明", "impact": "影響度"}],
  "experts": [{"type": "専門家種別", "reason": "理由", "urgency": "緊急度"}],
  "controls": [{"point": "統制ポイント", "description": "説明", "importance": "重要度"}]
}`;
  
  const response = callOpenAI(prompt);
  
  try {
    return JSON.parse(response);
  } catch (e) {
    return {
      disclosures: [],
      approvals: [],
      risks: [{category: "分析エラー", description: response, impact: "低"}],
      experts: [],
      controls: []
    };
  }
}

// ========= タスク抽出と業務フロー生成 =========
function extractTasksAndCreateWorkflow(messages) {
  const tasks = [];
  const flowSteps = [];
  const actors = new Set();
  
  // タスク関連キーワード
  const taskKeywords = {
    action: ['する', 'します', 'してください', 'お願い', '依頼', 'タスク', 'TODO'],
    deadline: ['まで', '期限', '締切', 'いつまで'],
    responsible: ['担当', '責任者', '@'],
    priority: ['至急', '緊急', '重要', '優先', 'ASAP'],
    process: ['手順', 'プロセス', 'フロー', '流れ', 'ステップ']
  };
  
  messages.forEach((msg, index) => {
    if (!msg.text) return;
    
    const text = msg.text;
    const msgDate = new Date(parseFloat(msg.ts) * 1000);
    
    // タスク抽出
    if (taskKeywords.action.some(kw => text.includes(kw))) {
      const task = {
        id: `TASK-${index + 1}`,
        description: text.substring(0, 200),
        createdAt: msgDate,
        user: msg.user || 'unknown',
        priority: taskKeywords.priority.some(kw => text.includes(kw)) ? '高' : '中',
        status: '未着手'
      };
      
      // 期限の抽出
      if (taskKeywords.deadline.some(kw => text.includes(kw))) {
        task.deadline = extractDeadline(text);
      }
      
      // 担当者の抽出
      if (text.includes('@')) {
        const mentions = text.match(/@[\w\-]+/g);
        if (mentions) {
          task.assignee = mentions[0].replace('@', '');
          actors.add(task.assignee);
        }
      }
      
      tasks.push(task);
    }
    
    // フローステップの抽出
    if (taskKeywords.process.some(kw => text.includes(kw))) {
      flowSteps.push({
        stepNo: flowSteps.length + 1,
        description: text.substring(0, 150),
        type: determineStepType(text),
        actor: msg.user || 'unknown'
      });
      actors.add(msg.user || 'unknown');
    }
  });
  
  // フローステップが少ない場合、タスクから生成
  if (flowSteps.length < 3 && tasks.length > 0) {
    tasks.forEach((task, index) => {
      flowSteps.push({
        stepNo: index + 1,
        description: task.description,
        type: '処理',
        actor: task.assignee || task.user
      });
    });
  }
  
  return {
    tasks: tasks,
    flowSteps: flowSteps,
    actors: Array.from(actors),
    summary: `${tasks.length}個のタスクと${flowSteps.length}個のプロセスステップを抽出`
  };
}

// ========= 期限抽出 =========
function extractDeadline(text) {
  const patterns = [
    /(\d{1,2}月\d{1,2}日)/,
    /(\d{4}年\d{1,2}月\d{1,2}日)/,
    /(今週|来週|今月|来月)末?/,
    /(\d+)日まで/
  ];
  
  for (const pattern of patterns) {
    const match = text.match(pattern);
    if (match) {
      return match[1];
    }
  }
  return '未定';
}

// ========= ステップタイプ判定 =========
function determineStepType(text) {
  if (text.includes('判断') || text.includes('確認') || text.includes('レビュー')) {
    return '判断';
  } else if (text.includes('承認') || text.includes('決裁')) {
    return '承認';
  } else if (text.includes('連絡') || text.includes('報告') || text.includes('共有')) {
    return '連絡';
  } else {
    return '処理';
  }
}

// ========= 業務記述書生成 =========
function generateBusinessSpecification(workflowData, channelName) {
  return {
    title: `業務記述書 - ${channelName}`,
    purpose: `${channelName}チャンネルで議論された業務プロセスの文書化`,
    scope: 'Slackメッセージから抽出された業務タスクとフロー',
    overview: workflowData.summary,
    actors: workflowData.actors,
    tasks: workflowData.tasks,
    flowSteps: workflowData.flowSteps,
    createdDate: new Date()
  };
}

// ========= OpenAI API呼び出し =========
function callOpenAI(prompt, model = 'gpt-4o') {
  const url = 'https://api.openai.com/v1/chat/completions';
  
  const payload = {
    model: model,
    messages: [
      { role: 'system', content: 'あなたは優秀なビジネスアナリストです。' },
      { role: 'user', content: prompt }
    ],
    temperature: 0.7,
    max_tokens: 2000
  };
  
  const options = {
    method: 'post',
    headers: {
      'Authorization': 'Bearer ' + OPENAI_API_KEY,
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
    
    return data.choices[0].message.content;
  } catch (error) {
    console.error('OpenAI API呼び出しエラー:', error);
    throw error;
  }
}

// ========= シート作成関数 =========
function saveMessagesToSheet(sheet, channelId, messages) {
  const headers = ['Timestamp', 'Channel ID', 'User', 'Text', 'Thread TS'];
  
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  
  const rows = messages.map(msg => [
    new Date(parseFloat(msg.ts) * 1000),
    channelId,
    msg.user || '',
    msg.text || '',
    msg.thread_ts || ''
  ]);
  
  if (rows.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, headers.length).setValues(rows);
  }
}

function saveAnalysisResults(sheet, results) {
  sheet.clear();
  sheet.getRange(1, 1).setValue('AI分析結果');
  sheet.getRange(2, 1).setValue('全体要約');
  sheet.getRange(2, 2).setValue(results.summary);
  
  sheet.getRange(4, 1).setValue('議題一覧');
  const headers = ['タイトル', '要約', 'カテゴリ', '優先度'];
  sheet.getRange(5, 1, 1, headers.length).setValues([headers]);
  
  if (results.topics && results.topics.length > 0) {
    const topicRows = results.topics.map(topic => [
      topic.title,
      topic.summary,
      topic.category,
      topic.priority
    ]);
    sheet.getRange(6, 1, topicRows.length, headers.length).setValues(topicRows);
  }
}

function createGovernanceAnalysisSheet(sheet, results, channelName) {
  let row = 1;
  
  // タイトル
  sheet.getRange(row, 1).setValue(`ガバナンス分析 - ${channelName}`);
  sheet.getRange(row, 1).setFontSize(16).setFontWeight('bold');
  row += 2;
  
  // 開示要件
  if (results.disclosures && results.disclosures.length > 0) {
    sheet.getRange(row, 1).setValue('開示要件');
    sheet.getRange(row, 1).setFontWeight('bold');
    row++;
    
    const discHeaders = ['種別', '内容', '緊急度'];
    sheet.getRange(row, 1, 1, discHeaders.length).setValues([discHeaders]);
    row++;
    
    const discRows = results.disclosures.map(d => [d.type, d.content, d.urgency]);
    sheet.getRange(row, 1, discRows.length, discHeaders.length).setValues(discRows);
    row += discRows.length + 1;
  }
  
  // リスク項目
  if (results.risks && results.risks.length > 0) {
    sheet.getRange(row, 1).setValue('リスク項目');
    sheet.getRange(row, 1).setFontWeight('bold');
    row++;
    
    const riskHeaders = ['カテゴリ', '説明', '影響度'];
    sheet.getRange(row, 1, 1, riskHeaders.length).setValues([riskHeaders]);
    row++;
    
    const riskRows = results.risks.map(r => [r.category, r.description, r.impact]);
    sheet.getRange(row, 1, riskRows.length, riskHeaders.length).setValues(riskRows);
  }
}

function createBusinessSpecSheet(sheet, spec) {
  let row = 1;
  
  // タイトル
  sheet.getRange(row, 1).setValue(spec.title);
  sheet.getRange(row, 1).setFontSize(16).setFontWeight('bold');
  row += 2;
  
  // 基本情報
  const basicInfo = [
    ['目的', spec.purpose],
    ['スコープ', spec.scope],
    ['概要', spec.overview],
    ['作成日', spec.createdDate]
  ];
  
  sheet.getRange(row, 1, basicInfo.length, 2).setValues(basicInfo);
  row += basicInfo.length + 1;
  
  // 関係者
  if (spec.actors && spec.actors.length > 0) {
    sheet.getRange(row, 1).setValue('関係者');
    sheet.getRange(row, 1).setFontWeight('bold');
    row++;
    sheet.getRange(row, 1).setValue(spec.actors.join(', '));
    row += 2;
  }
}

function createWorkflowSheet(sheet, workflowData) {
  let row = 1;
  
  sheet.getRange(row, 1).setValue('業務フロー図');
  sheet.getRange(row, 1).setFontSize(16).setFontWeight('bold');
  row += 2;
  
  const headers = ['ステップ番号', '説明', 'タイプ', '担当者'];
  sheet.getRange(row, 1, 1, headers.length).setValues([headers]);
  row++;
  
  if (workflowData.flowSteps && workflowData.flowSteps.length > 0) {
    const flowRows = workflowData.flowSteps.map(step => [
      step.stepNo,
      step.description,
      step.type,
      step.actor
    ]);
    sheet.getRange(row, 1, flowRows.length, headers.length).setValues(flowRows);
  }
}

function createTaskManagementSheet(sheet, tasks) {
  let row = 1;
  
  sheet.getRange(row, 1).setValue('タスク管理表');
  sheet.getRange(row, 1).setFontSize(16).setFontWeight('bold');
  row += 2;
  
  const headers = ['タスクID', '説明', '優先度', '期限', '担当者', 'ステータス'];
  sheet.getRange(row, 1, 1, headers.length).setValues([headers]);
  row++;
  
  if (tasks && tasks.length > 0) {
    const taskRows = tasks.map(task => [
      task.id,
      task.description,
      task.priority,
      task.deadline || '未定',
      task.assignee || task.user,
      task.status
    ]);
    sheet.getRange(row, 1, taskRows.length, headers.length).setValues(taskRows);
  }
}

// ========= 通知関数 =========
function sendGovernanceNotificationEmail(results, channelName) {
  const subject = `[ガバナンス分析] ${channelName}`;
  
  const htmlBody = `
    <h2>ガバナンス分析結果</h2>
    <p>チャンネル: ${channelName}</p>
    <h3>リスク項目: ${results.risks.length}件</h3>
    <h3>承認要件: ${results.approvals.length}件</h3>
    <h3>開示要件: ${results.disclosures.length}件</h3>
  `;
  
  MailApp.sendEmail({
    to: REPORT_EMAIL,
    subject: subject,
    htmlBody: htmlBody
  });
}

function sendWorkflowNotificationEmail(data) {
  const subject = `[業務フロー生成] ${data.channelName}`;
  
  const htmlBody = `
    <h2>業務フロー生成完了</h2>
    <p>チャンネル: ${data.channelName}</p>
    <ul>
      <li>メッセージ数: ${data.messageCount}</li>
      <li>タスク数: ${data.taskCount}</li>
      <li>フローステップ: ${data.flowSteps}</li>
    </ul>
    <p><a href="${data.spreadsheetUrl}">スプレッドシートを開く</a></p>
  `;
  
  MailApp.sendEmail({
    to: REPORT_EMAIL,
    subject: subject,
    htmlBody: htmlBody
  });
}

function sendWorkflowSlackNotification(data, channelId) {
  try {
    const blocks = [
      {
        type: 'header',
        text: {
          type: 'plain_text',
          text: '📊 業務フロー生成完了'
        }
      },
      {
        type: 'section',
        fields: [
          {
            type: 'mrkdwn',
            text: `*メッセージ数:*\n${data.messageCount}`
          },
          {
            type: 'mrkdwn',
            text: `*タスク数:*\n${data.taskCount}`
          }
        ]
      },
      {
        type: 'actions',
        elements: [
          {
            type: 'button',
            text: {
              type: 'plain_text',
              text: 'スプレッドシートを開く'
            },
            url: data.spreadsheetUrl
          }
        ]
      }
    ];
    
    slackAPI('chat.postMessage', {
      channel: channelId,
      text: `業務フロー生成完了`,
      blocks: JSON.stringify(blocks)
    });
  } catch (error) {
    console.error('Slack通知エラー:', error);
  }
}

// ========= メール処理機能 =========
function processIncomingEmails() {
  const threads = GmailApp.search('is:unread subject:"[task]"', 0, 10);
  
  threads.forEach(thread => {
    const messages = thread.getMessages();
    messages.forEach(message => {
      if (message.isUnread()) {
        try {
          const body = message.getPlainBody();
          const subject = message.getSubject();
          const from = message.getFrom();
          
          // 業務仕様書生成
          const spec = generateBusinessSpecFromEmail(body);
          
          // 独立したスプレッドシート作成
          const newSS = SpreadsheetApp.create(`業務仕様書_${new Date().getTime()}`);
          const sheet = newSS.getActiveSheet();
          sheet.getRange(1, 1).setValue(spec);
          
          // 返信メール送信
          message.reply(`業務仕様書を作成しました。\n${newSS.getUrl()}`);
          
          // 既読にする
          message.markRead();
        } catch (error) {
          console.error('メール処理エラー:', error);
        }
      }
    });
  });
}

function generateBusinessSpecFromEmail(emailBody) {
  const prompt = `
以下のメール内容から業務仕様書を生成してください:

${emailBody}

業務仕様書として構造化して出力してください。
`;
  
  return callOpenAI(prompt);
}

// ========= 統計分析 =========
function analyzeMessageStatistics(messages) {
  const stats = {
    totalMessages: messages.length,
    uniqueUsers: new Set(messages.map(m => m.user)).size,
    messagesPerHour: {},
    topUsers: {},
    categories: {}
  };
  
  messages.forEach(msg => {
    const hour = new Date(parseFloat(msg.ts) * 1000).getHours();
    stats.messagesPerHour[hour] = (stats.messagesPerHour[hour] || 0) + 1;
    
    const user = msg.user || 'unknown';
    stats.topUsers[user] = (stats.topUsers[user] || 0) + 1;
  });
  
  // ピーク時間を特定
  stats.peakHour = Object.entries(stats.messagesPerHour)
    .sort((a, b) => b[1] - a[1])[0];
  
  // 最もアクティブなユーザー
  stats.mostActiveUser = Object.entries(stats.topUsers)
    .sort((a, b) => b[1] - a[1])[0];
  
  return stats;
}

// ========= エラーハンドリング =========
function retryWithBackoff(func, maxRetries = 3) {
  let lastError;
  
  for (let i = 0; i < maxRetries; i++) {
    try {
      return func();
    } catch (error) {
      lastError = error;
      console.error(`試行 ${i + 1} 失敗:`, error);
      
      if (i < maxRetries - 1) {
        const waitTime = Math.pow(2, i) * 1000; // 指数バックオフ
        Utilities.sleep(waitTime);
      }
    }
  }
  
  throw lastError;
}

// ========= 実行時間管理 =========
function checkExecutionTime(startTime, limitMinutes = 5) {
  const elapsed = (new Date().getTime() - startTime) / 1000 / 60;
  if (elapsed > limitMinutes) {
    console.log(`実行時間制限（${limitMinutes}分）に達しました`);
    return true;
  }
  return false;
}