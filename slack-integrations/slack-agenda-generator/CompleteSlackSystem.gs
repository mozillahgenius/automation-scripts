// ==========================================
// 完全統合版：Slack議題生成＆分析システム
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

// ========= メイン統合実行関数 =========
function executeCompleteAnalysis() {
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.prompt(
    '完全統合分析',
    'チャンネルID（例: C09BW2EEVAR）を入力してください:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() !== ui.Button.OK) return;
  
  const channelId = response.getResponseText().trim();
  
  if (!channelId || !channelId.startsWith('C')) {
    ui.alert('エラー', 'チャンネルIDは「C」で始まる必要があります。', ui.ButtonSet.OK);
    return;
  }
  
  const startTime = new Date().getTime();
  
  try {
    ui.alert('処理開始', '完全統合分析を開始します。処理には数分かかる場合があります。', ui.ButtonSet.OK);
    
    // 1. メッセージ取得
    console.log('ステップ1: メッセージ取得中...');
    const messages = fetchChannelMessages(channelId);
    
    if (messages.length === 0) {
      ui.alert('情報', 'メッセージが見つかりませんでした。', ui.ButtonSet.OK);
      return;
    }
    
    // 2. AI分析
    console.log('ステップ2: AI分析実行中...');
    const analysisResults = analyzeMessagesWithAI(messages);
    
    // 3. ガバナンス分析
    console.log('ステップ3: ガバナンス分析実行中...');
    const governanceResults = analyzeMessagesForGovernance(messages);
    
    // 4. 業務フロー生成
    console.log('ステップ4: 業務フロー生成中...');
    const workflowData = extractTasksAndCreateWorkflow(messages);
    
    // 5. 議事録生成
    console.log('ステップ5: 議事録生成中...');
    const minutes = generateMinutesFromMessages(messages, analysisResults);
    
    // 6. スプレッドシート作成
    console.log('ステップ6: スプレッドシート作成中...');
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const timestamp = new Date().getTime();
    
    // メッセージシート
    const msgSheet = ss.insertSheet(`Messages_${timestamp}`);
    saveMessagesToSheet(msgSheet, channelId, messages);
    
    // 分析結果シート
    const analysisSheet = ss.insertSheet(`Analysis_${timestamp}`);
    saveAnalysisResults(analysisSheet, analysisResults);
    
    // ガバナンスシート
    const govSheet = ss.insertSheet(`Governance_${timestamp}`);
    createGovernanceAnalysisSheet(govSheet, governanceResults, channelId);
    
    // 業務記述書シート
    const businessSpec = generateBusinessSpecification(workflowData, channelId);
    const specSheet = ss.insertSheet(`Spec_${timestamp}`);
    createBusinessSpecSheet(specSheet, businessSpec);
    
    // 業務フローシート
    const flowSheet = ss.insertSheet(`Flow_${timestamp}`);
    createWorkflowSheet(flowSheet, workflowData);
    
    // ビジュアルフローシート
    const visualSheet = ss.insertSheet(`VisualFlow_${timestamp}`);
    createVisualFlowChart(visualSheet, workflowData);
    
    // タスク管理シート
    const taskSheet = ss.insertSheet(`Tasks_${timestamp}`);
    createTaskManagementSheet(taskSheet, workflowData.tasks);
    
    // RACI責任分担表
    const raciSheet = ss.insertSheet(`RACI_${timestamp}`);
    createRACIMatrix(raciSheet, workflowData);
    
    // タイムラインシート
    const timelineSheet = ss.insertSheet(`Timeline_${timestamp}`);
    createProcessTimeline(timelineSheet, workflowData);
    
    // 議事録シート
    const minutesSheet = ss.insertSheet(`Minutes_${timestamp}`);
    saveMinutesToSheet(minutesSheet, minutes);
    
    // 7. 通知送信
    console.log('ステップ7: 通知送信中...');
    const notificationData = {
      channelId: channelId,
      messageCount: messages.length,
      analysisResults: analysisResults,
      governanceResults: governanceResults,
      workflowData: workflowData,
      spreadsheetUrl: ss.getUrl(),
      timestamp: timestamp
    };
    
    // メール通知
    if (REPORT_EMAIL) {
      sendComprehensiveNotificationEmail(notificationData);
    }
    
    // Slack通知
    sendComprehensiveSlackNotification(notificationData, channelId);
    
    // 実行時間チェック
    const executionTime = (new Date().getTime() - startTime) / 1000;
    
    ui.alert('完了', `
完全統合分析が完了しました！

📊 処理結果:
- メッセージ数: ${messages.length}件
- 抽出議題: ${analysisResults.topics.length}件
- リスク項目: ${governanceResults.risks.length}件
- タスク数: ${workflowData.tasks.length}件
- 実行時間: ${executionTime.toFixed(1)}秒

詳細はスプレッドシートをご確認ください。
${ss.getUrl()}
    `, ui.ButtonSet.OK);
    
  } catch (error) {
    console.error('統合分析エラー:', error);
    ui.alert('エラー', `処理中にエラーが発生しました:\n${error.toString()}`, ui.ButtonSet.OK);
    
    // エラー通知
    if (REPORT_EMAIL) {
      sendErrorNotificationEmail(REPORT_EMAIL, 'Complete Analysis Error', error.toString());
    }
  }
}

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

// ========= メッセージ取得関数 =========
function fetchChannelMessages(channelId) {
  try {
    // チャンネル参加試行
    try {
      slackAPI('conversations.join', { channel: channelId });
    } catch (joinError) {
      console.log('チャンネル参加スキップ');
    }
    
    // メッセージ取得
    const history = slackAPI('conversations.history', {
      channel: channelId,
      limit: MAX_MESSAGES_PER_CHANNEL
    });
    
    const messages = history.messages || [];
    
    // スレッド返信を取得
    if (FETCH_THREAD_REPLIES) {
      const messagesWithReplies = [];
      
      messages.forEach(msg => {
        messagesWithReplies.push(msg);
        
        if (msg.thread_ts && msg.reply_count > 0) {
          try {
            const replies = slackAPI('conversations.replies', {
              channel: channelId,
              ts: msg.thread_ts,
              limit: 100
            });
            
            if (replies.messages && replies.messages.length > 1) {
              // 最初のメッセージは親メッセージなのでスキップ
              replies.messages.slice(1).forEach(reply => {
                messagesWithReplies.push(reply);
              });
            }
          } catch (error) {
            console.error(`スレッド返信取得エラー: ${error}`);
          }
        }
      });
      
      return messagesWithReplies;
    }
    
    return messages;
  } catch (error) {
    console.error(`メッセージ取得エラー: ${error}`);
    throw error;
  }
}

// ========= AI分析関数 =========
function analyzeMessagesWithAI(messages) {
  if (!OPENAI_API_KEY) {
    throw new Error('OpenAI APIキーが設定されていません');
  }
  
  const messageText = messages.map(m => m.text || '').join('\n');
  
  const prompt = `
以下のSlackメッセージから重要な議題、論点、カテゴリを詳細に分析してください。

メッセージ:
${messageText.substring(0, 10000)}

以下の形式でJSON出力してください:
{
  "topics": [
    {
      "title": "議題タイトル",
      "summary": "詳細な要約",
      "category": "カテゴリ（予算・契約・人事・システム・営業・その他）",
      "priority": "高/中/低",
      "keyPoints": ["重要ポイント1", "重要ポイント2"],
      "actionItems": ["アクション1", "アクション2"],
      "relatedTopics": ["関連議題1", "関連議題2"]
    }
  ],
  "summary": "全体要約（200文字以上）",
  "categories": ["カテゴリ1", "カテゴリ2"],
  "recommendations": ["推奨事項1", "推奨事項2"],
  "risks": ["リスク1", "リスク2"],
  "nextSteps": ["次のステップ1", "次のステップ2"]
}`;
  
  const response = callOpenAI(prompt, 'gpt-4o');
  
  try {
    return JSON.parse(response);
  } catch (e) {
    console.error('JSON解析エラー:', e);
    return {
      topics: [{title: "分析結果", summary: response, category: "その他", priority: "中", keyPoints: [], actionItems: [], relatedTopics: []}],
      summary: response,
      categories: ["その他"],
      recommendations: [],
      risks: [],
      nextSteps: []
    };
  }
}

// ========= ガバナンス分析関数 =========
function analyzeMessagesForGovernance(messages) {
  const messageText = messages.map(m => m.text || '').join('\n');
  
  const prompt = `
以下のSlackメッセージから、ガバナンス・コンプライアンスの観点で重要な項目を詳細に分析してください。

メッセージ:
${messageText.substring(0, 10000)}

以下の観点で分析してください:
1. 開示要件（適時開示・決算開示・法定開示）
2. 承認フローの適切性（取締役会・経営会議・稟議）
3. リスクアセスメント（財務・法務・オペレーショナル・レピュテーション）
4. 必要な専門家（弁護士・会計士・税理士・社労士・その他）
5. 内部統制のポイント（J-SOX・監査・統制活動）
6. コンプライアンスギャップ

JSON形式で出力してください:
{
  "disclosures": [
    {
      "type": "適時開示/決算開示/法定開示",
      "content": "開示内容",
      "urgency": "即時/1日以内/1週間以内",
      "requirement": "要件詳細",
      "deadline": "期限"
    }
  ],
  "approvals": [
    {
      "item": "承認項目",
      "level": "取締役会/経営会議/部長決裁/課長決裁",
      "risk": "高/中/低",
      "currentStatus": "未承認/承認済み/承認中",
      "requiredDocuments": ["必要書類1", "必要書類2"]
    }
  ],
  "risks": [
    {
      "category": "財務/法務/オペレーショナル/レピュテーション",
      "description": "リスク詳細説明",
      "impact": "重大/高/中/低",
      "probability": "高/中/低",
      "mitigation": "軽減策",
      "owner": "責任者"
    }
  ],
  "experts": [
    {
      "type": "弁護士/会計士/税理士/社労士/その他",
      "reason": "相談理由",
      "urgency": "即時/1日以内/1週間以内",
      "scope": "相談範囲",
      "estimatedCost": "想定費用"
    }
  ],
  "controls": [
    {
      "point": "統制ポイント",
      "description": "詳細説明",
      "importance": "重要/高/中/低",
      "controlNumber": "CTRL-XXX",
      "testProcedure": "テスト手続き",
      "frequency": "日次/週次/月次/四半期/年次"
    }
  ],
  "complianceGaps": [
    {
      "area": "ギャップ領域",
      "description": "ギャップ詳細",
      "impact": "影響度",
      "remediation": "是正措置"
    }
  ]
}`;
  
  const response = callOpenAI(prompt, 'gpt-4o');
  
  try {
    return JSON.parse(response);
  } catch (e) {
    console.error('ガバナンス分析JSON解析エラー:', e);
    return {
      disclosures: [],
      approvals: [],
      risks: [],
      experts: [],
      controls: [],
      complianceGaps: []
    };
  }
}

// ========= タスク抽出と業務フロー生成 =========
function extractTasksAndCreateWorkflow(messages) {
  const tasks = [];
  const flowSteps = [];
  const actors = new Set();
  const timeline = [];
  
  // タスク関連キーワード
  const taskKeywords = {
    action: ['する', 'します', 'してください', 'お願い', '依頼', 'タスク', 'TODO', 'やること', '実施', '作成', '確認', '承認'],
    deadline: ['まで', '期限', '締切', 'いつまで', 'デッドライン', '納期'],
    responsible: ['担当', '責任者', 'オーナー', '@', '誰が'],
    priority: ['至急', '緊急', '重要', '優先', 'ASAP', '最優先', '早急'],
    process: ['手順', 'プロセス', 'フロー', '流れ', 'ステップ', '工程', '段階']
  };
  
  messages.forEach((msg, index) => {
    if (!msg.text) return;
    
    const text = msg.text;
    const msgDate = new Date(parseFloat(msg.ts) * 1000);
    
    // タスク抽出
    if (taskKeywords.action.some(kw => text.includes(kw))) {
      const task = {
        id: `TASK-${tasks.length + 1}`,
        description: text.substring(0, 200),
        createdAt: msgDate,
        user: msg.user || 'unknown',
        priority: taskKeywords.priority.some(kw => text.includes(kw)) ? '高' : '中',
        status: '未着手',
        estimatedHours: estimateTaskHours(text),
        dependencies: extractDependencies(text, tasks),
        deliverables: extractDeliverables(text)
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
      
      // タイムラインに追加
      timeline.push({
        date: msgDate,
        type: 'task',
        content: task.description,
        assignee: task.assignee || task.user
      });
    }
    
    // フローステップの抽出
    if (taskKeywords.process.some(kw => text.includes(kw))) {
      const step = {
        stepNo: flowSteps.length + 1,
        description: text.substring(0, 150),
        type: determineStepType(text),
        actor: msg.user || 'unknown',
        tools: extractTools(text),
        inputs: extractInputs(text),
        outputs: extractOutputs(text),
        conditions: extractConditions(text)
      };
      
      flowSteps.push(step);
      actors.add(step.actor);
      
      // タイムラインに追加
      timeline.push({
        date: msgDate,
        type: 'process',
        content: step.description,
        actor: step.actor
      });
    }
  });
  
  // フローステップが少ない場合、タスクから生成
  if (flowSteps.length < 3 && tasks.length > 0) {
    tasks.forEach((task, index) => {
      flowSteps.push({
        stepNo: flowSteps.length + 1,
        description: task.description,
        type: '処理',
        actor: task.assignee || task.user,
        tools: [],
        inputs: [],
        outputs: task.deliverables,
        conditions: []
      });
    });
  }
  
  // RACI情報の生成
  const raciMatrix = generateRACIMatrix(tasks, actors);
  
  return {
    tasks: tasks,
    flowSteps: flowSteps,
    actors: Array.from(actors),
    timeline: timeline.sort((a, b) => a.date - b.date),
    raciMatrix: raciMatrix,
    summary: `${tasks.length}個のタスクと${flowSteps.length}個のプロセスステップを抽出`,
    totalEstimatedHours: tasks.reduce((sum, task) => sum + (task.estimatedHours || 0), 0)
  };
}

// ========= 補助関数群 =========
function extractDeadline(text) {
  const patterns = [
    /(\d{1,2}月\d{1,2}日)/,
    /(\d{4}年\d{1,2}月\d{1,2}日)/,
    /(今週|来週|今月|来月)末?/,
    /(\d+)日まで/,
    /(月曜|火曜|水曜|木曜|金曜|土曜|日曜)まで/
  ];
  
  for (const pattern of patterns) {
    const match = text.match(pattern);
    if (match) {
      return match[1];
    }
  }
  return '未定';
}

function determineStepType(text) {
  if (text.includes('判断') || text.includes('確認') || text.includes('レビュー')) {
    return '判断';
  } else if (text.includes('承認') || text.includes('決裁')) {
    return '承認';
  } else if (text.includes('連絡') || text.includes('報告') || text.includes('共有')) {
    return '連絡';
  } else if (text.includes('作成') || text.includes('作業') || text.includes('実施')) {
    return '処理';
  } else {
    return '処理';
  }
}

function estimateTaskHours(text) {
  if (text.includes('簡単') || text.includes('すぐ')) return 1;
  if (text.includes('複雑') || text.includes('詳細')) return 8;
  if (text.includes('大規模') || text.includes('プロジェクト')) return 40;
  return 4; // デフォルト
}

function extractDependencies(text, existingTasks) {
  const deps = [];
  existingTasks.forEach(task => {
    if (text.includes(task.id) || text.includes('前の') || text.includes('完了後')) {
      deps.push(task.id);
    }
  });
  return deps;
}

function extractDeliverables(text) {
  const deliverables = [];
  const patterns = ['資料', 'レポート', 'ドキュメント', '報告書', '提案書', 'プレゼン'];
  
  patterns.forEach(pattern => {
    if (text.includes(pattern)) {
      deliverables.push(pattern);
    }
  });
  
  return deliverables;
}

function extractTools(text) {
  const tools = [];
  const toolPatterns = ['Excel', 'Word', 'PowerPoint', 'Slack', 'メール', 'システム', 'ツール'];
  
  toolPatterns.forEach(tool => {
    if (text.toLowerCase().includes(tool.toLowerCase())) {
      tools.push(tool);
    }
  });
  
  return tools;
}

function extractInputs(text) {
  const inputs = [];
  if (text.includes('データ')) inputs.push('データ');
  if (text.includes('情報')) inputs.push('情報');
  if (text.includes('要件')) inputs.push('要件');
  return inputs;
}

function extractOutputs(text) {
  const outputs = [];
  if (text.includes('結果')) outputs.push('結果');
  if (text.includes('成果物')) outputs.push('成果物');
  if (text.includes('レポート')) outputs.push('レポート');
  return outputs;
}

function extractConditions(text) {
  const conditions = [];
  if (text.includes('もし') || text.includes('場合')) {
    conditions.push('条件分岐あり');
  }
  return conditions;
}

function generateRACIMatrix(tasks, actors) {
  const matrix = {};
  
  actors.forEach(actor => {
    matrix[actor] = {};
    
    tasks.forEach(task => {
      if (task.assignee === actor) {
        matrix[actor][task.id] = 'R'; // Responsible
      } else if (task.user === actor) {
        matrix[actor][task.id] = 'A'; // Accountable
      } else if (task.description.includes(actor)) {
        matrix[actor][task.id] = 'C'; // Consulted
      } else {
        matrix[actor][task.id] = 'I'; // Informed
      }
    });
  });
  
  return matrix;
}

// ========= 議事録生成 =========
function generateMinutesFromMessages(messages, analysisResults) {
  const minutes = {
    date: new Date(),
    title: '議事録',
    attendees: [...new Set(messages.map(m => m.user).filter(u => u))],
    agenda: analysisResults.topics.map(t => t.title),
    discussions: [],
    decisions: [],
    actionItems: [],
    nextSteps: analysisResults.nextSteps || []
  };
  
  // 議論内容の抽出
  analysisResults.topics.forEach(topic => {
    minutes.discussions.push({
      topic: topic.title,
      summary: topic.summary,
      keyPoints: topic.keyPoints || []
    });
    
    // アクションアイテムの追加
    if (topic.actionItems && topic.actionItems.length > 0) {
      topic.actionItems.forEach(item => {
        minutes.actionItems.push({
          action: item,
          topic: topic.title,
          priority: topic.priority
        });
      });
    }
  });
  
  // 決定事項の抽出（キーワードベース）
  messages.forEach(msg => {
    if (msg.text && (msg.text.includes('決定') || msg.text.includes('承認') || msg.text.includes('合意'))) {
      minutes.decisions.push({
        decision: msg.text.substring(0, 200),
        timestamp: new Date(parseFloat(msg.ts) * 1000)
      });
    }
  });
  
  return minutes;
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
    estimatedTotalHours: workflowData.totalEstimatedHours,
    timeline: workflowData.timeline,
    raciMatrix: workflowData.raciMatrix,
    createdDate: new Date(),
    version: '1.0',
    approver: '',
    reviewer: ''
  };
}

// ========= OpenAI API呼び出し =========
function callOpenAI(prompt, model = 'gpt-4o') {
  const url = 'https://api.openai.com/v1/chat/completions';
  
  const payload = {
    model: model,
    messages: [
      { role: 'system', content: 'あなたは優秀なビジネスアナリストです。日本語で詳細かつ構造化された分析を提供してください。' },
      { role: 'user', content: prompt }
    ],
    temperature: 0.7,
    max_tokens: 4000
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
  
  return retryWithBackoff(() => {
    const response = UrlFetchApp.fetch(url, options);
    const data = JSON.parse(response.getContentText());
    
    if (data.error) {
      throw new Error(`OpenAI API Error: ${data.error.message}`);
    }
    
    return data.choices[0].message.content;
  });
}

// ========= シート作成関数群 =========
function saveMessagesToSheet(sheet, channelId, messages) {
  const headers = ['Timestamp', 'Channel ID', 'User', 'Text', 'Thread TS', 'Reply Count'];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  
  const rows = messages.map(msg => [
    new Date(parseFloat(msg.ts) * 1000),
    channelId,
    msg.user || '',
    msg.text || '',
    msg.thread_ts || '',
    msg.reply_count || 0
  ]);
  
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }
}

function saveAnalysisResults(sheet, results) {
  sheet.getRange(1, 1).setValue('AI分析結果');
  sheet.getRange(1, 1).setFontSize(16).setFontWeight('bold');
  
  sheet.getRange(3, 1).setValue('全体要約');
  sheet.getRange(3, 2).setValue(results.summary);
  sheet.getRange(3, 1).setFontWeight('bold');
  
  // 議題一覧
  sheet.getRange(5, 1).setValue('議題一覧');
  sheet.getRange(5, 1).setFontWeight('bold');
  
  const headers = ['タイトル', '要約', 'カテゴリ', '優先度'];
  sheet.getRange(6, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(6, 1, 1, headers.length).setFontWeight('bold');
  
  if (results.topics && results.topics.length > 0) {
    const topicRows = results.topics.map(topic => [
      topic.title,
      topic.summary,
      topic.category,
      topic.priority
    ]);
    sheet.getRange(7, 1, topicRows.length, headers.length).setValues(topicRows);
  }
  
  // 推奨事項
  let currentRow = 7 + (results.topics ? results.topics.length : 0) + 2;
  
  if (results.recommendations && results.recommendations.length > 0) {
    sheet.getRange(currentRow, 1).setValue('推奨事項');
    sheet.getRange(currentRow, 1).setFontWeight('bold');
    currentRow++;
    
    results.recommendations.forEach((rec, index) => {
      sheet.getRange(currentRow + index, 1).setValue(`${index + 1}. ${rec}`);
    });
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
    sheet.getRange(row, 1).setFontWeight('bold').setBackground('#FFF2CC');
    row++;
    
    const discHeaders = ['種別', '内容', '緊急度', '要件', '期限'];
    sheet.getRange(row, 1, 1, discHeaders.length).setValues([discHeaders]);
    sheet.getRange(row, 1, 1, discHeaders.length).setFontWeight('bold');
    row++;
    
    const discRows = results.disclosures.map(d => [
      d.type, d.content, d.urgency, d.requirement || '', d.deadline || ''
    ]);
    sheet.getRange(row, 1, discRows.length, discHeaders.length).setValues(discRows);
    row += discRows.length + 1;
  }
  
  // リスク項目
  if (results.risks && results.risks.length > 0) {
    sheet.getRange(row, 1).setValue('リスク項目');
    sheet.getRange(row, 1).setFontWeight('bold').setBackground('#FFE6E6');
    row++;
    
    const riskHeaders = ['カテゴリ', '説明', '影響度', '発生可能性', '軽減策', '責任者'];
    sheet.getRange(row, 1, 1, riskHeaders.length).setValues([riskHeaders]);
    sheet.getRange(row, 1, 1, riskHeaders.length).setFontWeight('bold');
    row++;
    
    const riskRows = results.risks.map(r => [
      r.category, r.description, r.impact, r.probability || '', r.mitigation || '', r.owner || ''
    ]);
    sheet.getRange(row, 1, riskRows.length, riskHeaders.length).setValues(riskRows);
    row += riskRows.length + 1;
  }
  
  // 専門家相談
  if (results.experts && results.experts.length > 0) {
    sheet.getRange(row, 1).setValue('必要な専門家');
    sheet.getRange(row, 1).setFontWeight('bold').setBackground('#E6F4FF');
    row++;
    
    const expertHeaders = ['専門家種別', '相談理由', '緊急度', '相談範囲', '想定費用'];
    sheet.getRange(row, 1, 1, expertHeaders.length).setValues([expertHeaders]);
    sheet.getRange(row, 1, 1, expertHeaders.length).setFontWeight('bold');
    row++;
    
    const expertRows = results.experts.map(e => [
      e.type, e.reason, e.urgency, e.scope || '', e.estimatedCost || ''
    ]);
    sheet.getRange(row, 1, expertRows.length, expertHeaders.length).setValues(expertRows);
    row += expertRows.length + 1;
  }
  
  // 内部統制
  if (results.controls && results.controls.length > 0) {
    sheet.getRange(row, 1).setValue('内部統制ポイント');
    sheet.getRange(row, 1).setFontWeight('bold').setBackground('#E6FFE6');
    row++;
    
    const controlHeaders = ['統制ポイント', '説明', '重要度', 'コントロール番号', 'テスト手続き', '頻度'];
    sheet.getRange(row, 1, 1, controlHeaders.length).setValues([controlHeaders]);
    sheet.getRange(row, 1, 1, controlHeaders.length).setFontWeight('bold');
    row++;
    
    const controlRows = results.controls.map(c => [
      c.point, c.description, c.importance, c.controlNumber || '', c.testProcedure || '', c.frequency || ''
    ]);
    sheet.getRange(row, 1, controlRows.length, controlHeaders.length).setValues(controlRows);
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
    ['文書バージョン', spec.version],
    ['作成日', spec.createdDate],
    ['目的', spec.purpose],
    ['スコープ', spec.scope],
    ['概要', spec.overview],
    ['推定総工数（時間）', spec.estimatedTotalHours]
  ];
  
  basicInfo.forEach(info => {
    sheet.getRange(row, 1).setValue(info[0]);
    sheet.getRange(row, 1).setFontWeight('bold');
    sheet.getRange(row, 2).setValue(info[1]);
    row++;
  });
  
  row++;
  
  // 関係者
  if (spec.actors && spec.actors.length > 0) {
    sheet.getRange(row, 1).setValue('関係者');
    sheet.getRange(row, 1).setFontWeight('bold');
    row++;
    sheet.getRange(row, 1).setValue(spec.actors.join(', '));
    row += 2;
  }
  
  // タスク一覧
  if (spec.tasks && spec.tasks.length > 0) {
    sheet.getRange(row, 1).setValue('タスク一覧');
    sheet.getRange(row, 1).setFontWeight('bold');
    row++;
    
    const taskHeaders = ['ID', '説明', '優先度', '期限', '担当者', '推定時間'];
    sheet.getRange(row, 1, 1, taskHeaders.length).setValues([taskHeaders]);
    sheet.getRange(row, 1, 1, taskHeaders.length).setFontWeight('bold');
    row++;
    
    const taskRows = spec.tasks.map(task => [
      task.id,
      task.description,
      task.priority,
      task.deadline || '未定',
      task.assignee || task.user,
      task.estimatedHours || ''
    ]);
    sheet.getRange(row, 1, taskRows.length, taskHeaders.length).setValues(taskRows);
  }
}

function createWorkflowSheet(sheet, workflowData) {
  let row = 1;
  
  sheet.getRange(row, 1).setValue('業務フロー図');
  sheet.getRange(row, 1).setFontSize(16).setFontWeight('bold');
  row += 2;
  
  const headers = ['ステップ番号', '説明', 'タイプ', '担当者', 'ツール', '入力', '出力', '条件'];
  sheet.getRange(row, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(row, 1, 1, headers.length).setFontWeight('bold');
  row++;
  
  if (workflowData.flowSteps && workflowData.flowSteps.length > 0) {
    const flowRows = workflowData.flowSteps.map(step => [
      step.stepNo,
      step.description,
      step.type,
      step.actor,
      step.tools ? step.tools.join(', ') : '',
      step.inputs ? step.inputs.join(', ') : '',
      step.outputs ? step.outputs.join(', ') : '',
      step.conditions ? step.conditions.join(', ') : ''
    ]);
    sheet.getRange(row, 1, flowRows.length, headers.length).setValues(flowRows);
  }
}

function createVisualFlowChart(sheet, workflowData) {
  sheet.getRange(1, 1).setValue('ビジュアルフローチャート');
  sheet.getRange(1, 1).setFontSize(16).setFontWeight('bold');
  
  // フローチャートの描画
  let row = 3;
  let col = 2;
  
  workflowData.flowSteps.forEach((step, index) => {
    // ステップボックスの作成
    const range = sheet.getRange(row, col, 3, 4);
    range.merge();
    range.setValue(`${step.stepNo}. ${step.description}`);
    range.setBorder(true, true, true, true, false, false);
    range.setHorizontalAlignment('center');
    range.setVerticalAlignment('middle');
    range.setWrap(true);
    
    // タイプによる色分け
    if (step.type === '判断') {
      range.setBackground('#FFE6B3'); // オレンジ
    } else if (step.type === '承認') {
      range.setBackground('#FFB3B3'); // 赤
    } else if (step.type === '連絡') {
      range.setBackground('#B3D9FF'); // 青
    } else {
      range.setBackground('#B3FFB3'); // 緑
    }
    
    // 担当者を下に表示
    sheet.getRange(row + 3, col, 1, 4).merge();
    sheet.getRange(row + 3, col).setValue(`担当: ${step.actor}`);
    sheet.getRange(row + 3, col).setFontSize(10);
    sheet.getRange(row + 3, col).setHorizontalAlignment('center');
    
    // 矢印を追加（最後のステップ以外）
    if (index < workflowData.flowSteps.length - 1) {
      sheet.getRange(row + 4, col + 1, 1, 2).merge();
      sheet.getRange(row + 4, col + 1).setValue('↓');
      sheet.getRange(row + 4, col + 1).setFontSize(20);
      sheet.getRange(row + 4, col + 1).setHorizontalAlignment('center');
    }
    
    row += 5;
  });
}

function createTaskManagementSheet(sheet, tasks) {
  let row = 1;
  
  sheet.getRange(row, 1).setValue('タスク管理表');
  sheet.getRange(row, 1).setFontSize(16).setFontWeight('bold');
  row += 2;
  
  const headers = ['タスクID', '説明', '優先度', '期限', '担当者', 'ステータス', '推定時間', '依存関係', '成果物'];
  sheet.getRange(row, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(row, 1, 1, headers.length).setFontWeight('bold');
  row++;
  
  if (tasks && tasks.length > 0) {
    const taskRows = tasks.map(task => [
      task.id,
      task.description,
      task.priority,
      task.deadline || '未定',
      task.assignee || task.user,
      task.status,
      task.estimatedHours || '',
      task.dependencies ? task.dependencies.join(', ') : '',
      task.deliverables ? task.deliverables.join(', ') : ''
    ]);
    sheet.getRange(row, 1, taskRows.length, headers.length).setValues(taskRows);
    
    // 優先度による色分け
    tasks.forEach((task, index) => {
      if (task.priority === '高') {
        sheet.getRange(row + index, 3).setBackground('#FFB3B3');
      } else if (task.priority === '中') {
        sheet.getRange(row + index, 3).setBackground('#FFFFB3');
      } else {
        sheet.getRange(row + index, 3).setBackground('#B3FFB3');
      }
    });
  }
}

function createRACIMatrix(sheet, workflowData) {
  sheet.getRange(1, 1).setValue('RACI責任分担表');
  sheet.getRange(1, 1).setFontSize(16).setFontWeight('bold');
  
  sheet.getRange(3, 1).setValue('R: Responsible (実行責任者)');
  sheet.getRange(4, 1).setValue('A: Accountable (説明責任者)');
  sheet.getRange(5, 1).setValue('C: Consulted (相談先)');
  sheet.getRange(6, 1).setValue('I: Informed (報告先)');
  
  const startRow = 8;
  
  if (workflowData.raciMatrix && workflowData.tasks.length > 0) {
    // ヘッダー行（タスクID）
    const taskIds = ['担当者 / タスク'].concat(workflowData.tasks.map(t => t.id));
    sheet.getRange(startRow, 1, 1, taskIds.length).setValues([taskIds]);
    sheet.getRange(startRow, 1, 1, taskIds.length).setFontWeight('bold');
    
    // 各アクターの行
    let row = startRow + 1;
    Object.keys(workflowData.raciMatrix).forEach(actor => {
      const rowData = [actor];
      workflowData.tasks.forEach(task => {
        rowData.push(workflowData.raciMatrix[actor][task.id] || '');
      });
      sheet.getRange(row, 1, 1, rowData.length).setValues([rowData]);
      
      // RACI値による色分け
      rowData.slice(1).forEach((value, col) => {
        const cell = sheet.getRange(row, col + 2);
        if (value === 'R') cell.setBackground('#FFB3B3');
        else if (value === 'A') cell.setBackground('#FFE6B3');
        else if (value === 'C') cell.setBackground('#FFFFB3');
        else if (value === 'I') cell.setBackground('#E6E6E6');
      });
      
      row++;
    });
    
    // 枠線を追加
    sheet.getRange(startRow, 1, row - startRow, taskIds.length).setBorder(true, true, true, true, true, true);
  }
}

function createProcessTimeline(sheet, workflowData) {
  sheet.getRange(1, 1).setValue('プロセスタイムライン');
  sheet.getRange(1, 1).setFontSize(16).setFontWeight('bold');
  
  if (workflowData.timeline && workflowData.timeline.length > 0) {
    const headers = ['日時', 'タイプ', '内容', '担当者'];
    sheet.getRange(3, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(3, 1, 1, headers.length).setFontWeight('bold');
    
    const timelineRows = workflowData.timeline.map(item => [
      item.date,
      item.type,
      item.content,
      item.assignee || item.actor || ''
    ]);
    
    sheet.getRange(4, 1, timelineRows.length, headers.length).setValues(timelineRows);
    
    // タイプによる色分け
    workflowData.timeline.forEach((item, index) => {
      const rowNum = 4 + index;
      if (item.type === 'task') {
        sheet.getRange(rowNum, 2).setBackground('#B3FFB3');
      } else if (item.type === 'process') {
        sheet.getRange(rowNum, 2).setBackground('#B3D9FF');
      }
    });
  }
}

function saveMinutesToSheet(sheet, minutes) {
  let row = 1;
  
  // タイトル
  sheet.getRange(row, 1).setValue(minutes.title);
  sheet.getRange(row, 1).setFontSize(16).setFontWeight('bold');
  row += 2;
  
  // 基本情報
  sheet.getRange(row, 1).setValue('日時');
  sheet.getRange(row, 2).setValue(minutes.date);
  row++;
  
  sheet.getRange(row, 1).setValue('出席者');
  sheet.getRange(row, 2).setValue(minutes.attendees.join(', '));
  row += 2;
  
  // 議題
  sheet.getRange(row, 1).setValue('議題');
  sheet.getRange(row, 1).setFontWeight('bold');
  row++;
  minutes.agenda.forEach((item, index) => {
    sheet.getRange(row + index, 1).setValue(`${index + 1}. ${item}`);
  });
  row += minutes.agenda.length + 1;
  
  // 議論内容
  sheet.getRange(row, 1).setValue('議論内容');
  sheet.getRange(row, 1).setFontWeight('bold');
  row++;
  
  minutes.discussions.forEach(discussion => {
    sheet.getRange(row, 1).setValue(`【${discussion.topic}】`);
    sheet.getRange(row, 1).setFontWeight('bold');
    row++;
    sheet.getRange(row, 1).setValue(discussion.summary);
    row++;
    
    if (discussion.keyPoints && discussion.keyPoints.length > 0) {
      sheet.getRange(row, 1).setValue('重要ポイント:');
      row++;
      discussion.keyPoints.forEach(point => {
        sheet.getRange(row, 1).setValue(`・${point}`);
        row++;
      });
    }
    row++;
  });
  
  // 決定事項
  if (minutes.decisions && minutes.decisions.length > 0) {
    sheet.getRange(row, 1).setValue('決定事項');
    sheet.getRange(row, 1).setFontWeight('bold');
    row++;
    
    minutes.decisions.forEach((decision, index) => {
      sheet.getRange(row + index, 1).setValue(`${index + 1}. ${decision.decision}`);
    });
    row += minutes.decisions.length + 1;
  }
  
  // アクションアイテム
  if (minutes.actionItems && minutes.actionItems.length > 0) {
    sheet.getRange(row, 1).setValue('アクションアイテム');
    sheet.getRange(row, 1).setFontWeight('bold');
    row++;
    
    const actionHeaders = ['アクション', '関連議題', '優先度'];
    sheet.getRange(row, 1, 1, actionHeaders.length).setValues([actionHeaders]);
    sheet.getRange(row, 1, 1, actionHeaders.length).setFontWeight('bold');
    row++;
    
    const actionRows = minutes.actionItems.map(item => [
      item.action,
      item.topic,
      item.priority
    ]);
    sheet.getRange(row, 1, actionRows.length, actionHeaders.length).setValues(actionRows);
  }
}

// ========= 通知関数群 =========
function sendComprehensiveNotificationEmail(data) {
  const subject = `[完全分析レポート] チャンネル ${data.channelId}`;
  
  const htmlBody = `
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <style>
    body { font-family: 'Helvetica Neue', Arial, sans-serif; line-height: 1.6; color: #333; background-color: #f5f5f5; }
    .container { max-width: 800px; margin: 0 auto; padding: 20px; background: white; border-radius: 10px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }
    .header { background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; padding: 30px; border-radius: 10px 10px 0 0; margin: -20px -20px 20px -20px; }
    h1 { margin: 0; font-size: 28px; }
    h2 { color: #667eea; border-bottom: 2px solid #667eea; padding-bottom: 10px; }
    h3 { color: #764ba2; }
    .stats { display: grid; grid-template-columns: repeat(auto-fit, minmax(150px, 1fr)); gap: 15px; margin: 20px 0; }
    .stat-box { text-align: center; padding: 20px; background: #f8f9fa; border-radius: 8px; border: 1px solid #e9ecef; }
    .stat-number { font-size: 32px; font-weight: bold; color: #667eea; }
    .stat-label { color: #6c757d; margin-top: 5px; }
    .section { margin: 30px 0; padding: 20px; background: #f8f9fa; border-radius: 8px; }
    .button { display: inline-block; padding: 12px 24px; background: #667eea; color: white; text-decoration: none; border-radius: 5px; margin: 10px 5px; }
    .button:hover { background: #5a67d8; }
    ul { padding-left: 20px; }
    .risk-high { color: #dc3545; font-weight: bold; }
    .risk-medium { color: #ffc107; font-weight: bold; }
    .risk-low { color: #28a745; }
    .footer { margin-top: 40px; padding-top: 20px; border-top: 1px solid #e9ecef; text-align: center; color: #6c757d; }
  </style>
</head>
<body>
  <div class="container">
    <div class="header">
      <h1>📊 完全統合分析レポート</h1>
      <p style="margin: 10px 0 0 0;">チャンネル: ${data.channelId}</p>
      <p style="margin: 5px 0 0 0;">生成日時: ${new Date().toLocaleString('ja-JP')}</p>
    </div>
    
    <div class="stats">
      <div class="stat-box">
        <div class="stat-number">${data.messageCount}</div>
        <div class="stat-label">メッセージ数</div>
      </div>
      <div class="stat-box">
        <div class="stat-number">${data.analysisResults.topics.length}</div>
        <div class="stat-label">抽出議題</div>
      </div>
      <div class="stat-box">
        <div class="stat-number">${data.governanceResults.risks.length}</div>
        <div class="stat-label">リスク項目</div>
      </div>
      <div class="stat-box">
        <div class="stat-number">${data.workflowData.tasks.length}</div>
        <div class="stat-label">タスク数</div>
      </div>
    </div>
    
    <div class="section">
      <h2>📋 分析サマリー</h2>
      <p>${data.analysisResults.summary}</p>
    </div>
    
    <div class="section">
      <h2>🎯 主要議題</h2>
      <ul>
        ${data.analysisResults.topics.slice(0, 5).map(topic => `
          <li>
            <strong>${topic.title}</strong> 
            <span class="${topic.priority === '高' ? 'risk-high' : topic.priority === '中' ? 'risk-medium' : 'risk-low'}">
              [${topic.priority}]
            </span>
            <br>${topic.summary}
          </li>
        `).join('')}
      </ul>
    </div>
    
    <div class="section">
      <h2>⚠️ ガバナンス・リスク</h2>
      ${data.governanceResults.risks.length > 0 ? `
        <ul>
          ${data.governanceResults.risks.slice(0, 3).map(risk => `
            <li>
              <strong>${risk.category}</strong>: ${risk.description}
              <span class="${risk.impact === '重大' || risk.impact === '高' ? 'risk-high' : risk.impact === '中' ? 'risk-medium' : 'risk-low'}">
                [影響度: ${risk.impact}]
              </span>
            </li>
          `).join('')}
        </ul>
      ` : '<p>重大なリスクは検出されませんでした。</p>'}
      
      ${data.governanceResults.experts && data.governanceResults.experts.length > 0 ? `
        <h3>必要な専門家相談</h3>
        <ul>
          ${data.governanceResults.experts.map(expert => `
            <li>${expert.type}: ${expert.reason}</li>
          `).join('')}
        </ul>
      ` : ''}
    </div>
    
    <div class="section">
      <h2>📊 業務フロー</h2>
      <p>タスク数: ${data.workflowData.tasks.length}</p>
      <p>プロセスステップ: ${data.workflowData.flowSteps.length}</p>
      <p>推定総工数: ${data.workflowData.totalEstimatedHours || 0}時間</p>
      
      <h3>主要タスク</h3>
      <ul>
        ${data.workflowData.tasks.slice(0, 5).map(task => `
          <li>${task.id}: ${task.description.substring(0, 100)}...</li>
        `).join('')}
      </ul>
    </div>
    
    <div style="text-align: center; margin: 40px 0;">
      <a href="${data.spreadsheetUrl}" class="button">📊 詳細スプレッドシートを開く</a>
    </div>
    
    <div class="footer">
      <p>このレポートは自動生成されました。</p>
      <p>© Slack統合分析システム</p>
    </div>
  </div>
</body>
</html>
  `;
  
  const plainBody = `
完全統合分析レポート
=====================

チャンネル: ${data.channelId}
生成日時: ${new Date().toLocaleString('ja-JP')}

【処理結果】
- メッセージ数: ${data.messageCount}件
- 抽出議題: ${data.analysisResults.topics.length}件
- リスク項目: ${data.governanceResults.risks.length}件
- タスク数: ${data.workflowData.tasks.length}件

【分析サマリー】
${data.analysisResults.summary}

【主要議題】
${data.analysisResults.topics.slice(0, 5).map((topic, i) => 
  `${i + 1}. ${topic.title} [${topic.priority}]\n   ${topic.summary}`
).join('\n')}

【ガバナンス・リスク】
${data.governanceResults.risks.slice(0, 3).map((risk, i) => 
  `${i + 1}. ${risk.category}: ${risk.description} [影響度: ${risk.impact}]`
).join('\n')}

【業務フロー】
- タスク数: ${data.workflowData.tasks.length}
- プロセスステップ: ${data.workflowData.flowSteps.length}
- 推定総工数: ${data.workflowData.totalEstimatedHours || 0}時間

詳細はスプレッドシートをご確認ください:
${data.spreadsheetUrl}
  `;
  
  MailApp.sendEmail({
    to: REPORT_EMAIL,
    subject: subject,
    body: plainBody,
    htmlBody: htmlBody
  });
}

function sendComprehensiveSlackNotification(data, channelId) {
  try {
    const blocks = [
      {
        type: 'header',
        text: {
          type: 'plain_text',
          text: '📊 完全統合分析完了'
        }
      },
      {
        type: 'section',
        text: {
          type: 'mrkdwn',
          text: `*分析結果サマリー*\n${data.analysisResults.summary.substring(0, 200)}...`
        }
      },
      {
        type: 'section',
        fields: [
          {
            type: 'mrkdwn',
            text: `*📝 メッセージ数:*\n${data.messageCount}`
          },
          {
            type: 'mrkdwn',
            text: `*🎯 議題数:*\n${data.analysisResults.topics.length}`
          },
          {
            type: 'mrkdwn',
            text: `*⚠️ リスク数:*\n${data.governanceResults.risks.length}`
          },
          {
            type: 'mrkdwn',
            text: `*✅ タスク数:*\n${data.workflowData.tasks.length}`
          }
        ]
      },
      {
        type: 'section',
        text: {
          type: 'mrkdwn',
          text: '*主要議題:*\n' + data.analysisResults.topics.slice(0, 3).map((topic, i) => 
            `${i + 1}. ${topic.title} [${topic.priority}]`
          ).join('\n')
        }
      },
      {
        type: 'actions',
        elements: [
          {
            type: 'button',
            text: {
              type: 'plain_text',
              text: '📊 スプレッドシートを開く'
            },
            url: data.spreadsheetUrl,
            style: 'primary'
          }
        ]
      }
    ];
    
    slackAPI('chat.postMessage', {
      channel: channelId,
      text: `完全統合分析が完了しました`,
      blocks: JSON.stringify(blocks)
    });
  } catch (error) {
    console.error('Slack通知エラー:', error);
  }
}

function sendErrorNotificationEmail(to, subject, errorMessage) {
  const htmlBody = `
    <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto;">
      <div style="background: #dc3545; color: white; padding: 20px; border-radius: 5px 5px 0 0;">
        <h2 style="margin: 0;">⚠️ エラー通知</h2>
      </div>
      <div style="padding: 20px; background: #f8f9fa; border: 1px solid #dee2e6;">
        <p><strong>件名:</strong> ${subject}</p>
        <p><strong>エラー内容:</strong></p>
        <pre style="background: white; padding: 15px; border-radius: 5px; overflow-x: auto;">${errorMessage}</pre>
        <p><strong>発生時刻:</strong> ${new Date().toLocaleString('ja-JP')}</p>
      </div>
    </div>
  `;
  
  MailApp.sendEmail({
    to: to,
    subject: `[エラー] ${subject}`,
    htmlBody: htmlBody
  });
}

// ========= メール処理機能 =========
function processIncomingEmails() {
  const threads = GmailApp.search('is:unread subject:"[task]"', 0, 10);
  
  threads.forEach(thread => {
    const messages = thread.getMessages();
    messages.forEach(message => {
      if (message.isUnread()) {
        try {
          const body = cleanEmailBody(message.getPlainBody());
          const subject = message.getSubject();
          const from = message.getFrom();
          
          // 業務仕様書生成
          const spec = generateBusinessSpecFromEmail(body);
          
          // 独立したスプレッドシート作成
          const newSS = SpreadsheetApp.create(`業務仕様書_${new Date().getTime()}`);
          const sheet = newSS.getActiveSheet();
          
          // 仕様書をシートに書き込み
          let row = 1;
          sheet.getRange(row, 1).setValue('業務仕様書');
          sheet.getRange(row, 1).setFontSize(16).setFontWeight('bold');
          row += 2;
          
          sheet.getRange(row, 1).setValue(spec);
          
          // 返信メール送信
          const replyBody = `
業務仕様書を作成しました。

スプレッドシートURL: ${newSS.getUrl()}

作成日時: ${new Date().toLocaleString('ja-JP')}
          `;
          
          message.reply(replyBody);
          
          // 既読にする
          message.markRead();
          
          console.log(`メール処理完了: ${subject}`);
        } catch (error) {
          console.error('メール処理エラー:', error);
        }
      }
    });
  });
}

function cleanEmailBody(body) {
  // メール署名の削除
  const signaturePatterns = [
    /--\s*\n[\s\S]*$/,  // -- で始まる署名
    /^-{3,}[\s\S]*$/m,  // --- で始まる署名
    /^_{3,}[\s\S]*$/m,  // ___ で始まる署名
    /Sent from .*/i,     // "Sent from" で始まる行
    /^.*について、.*より/m  // 日本語の一般的な署名パターン
  ];
  
  let cleanedBody = body;
  signaturePatterns.forEach(pattern => {
    cleanedBody = cleanedBody.replace(pattern, '');
  });
  
  return cleanedBody.trim();
}

function generateBusinessSpecFromEmail(emailBody) {
  const prompt = `
以下のメール内容から詳細な業務仕様書を生成してください:

${emailBody}

以下の構成で業務仕様書を作成してください:
1. 概要
2. 目的
3. スコープ
4. 要件定義
5. 機能仕様
6. 非機能要件
7. 制約事項
8. スケジュール
9. リスクと対策
10. 成果物

詳細かつ実用的な内容にしてください。
`;
  
  return callOpenAI(prompt, 'gpt-4o');
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
        console.log(`${waitTime}ms 待機中...`);
        Utilities.sleep(waitTime);
      }
    }
  }
  
  console.error('すべての再試行が失敗しました');
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

// ========= バッチ処理 =========
function processBatch(items, batchSize, processor) {
  const results = [];
  
  for (let i = 0; i < items.length; i += batchSize) {
    const batch = items.slice(i, i + batchSize);
    try {
      const batchResults = processor(batch);
      results.push(...batchResults);
    } catch (error) {
      console.error(`バッチ処理エラー (${i}-${i + batchSize}):`, error);
    }
  }
  
  return results;
}

// ========= デバッグ関数 =========
function debugLog(message, data = null) {
  const timestamp = new Date().toISOString();
  console.log(`[DEBUG ${timestamp}] ${message}`);
  if (data) {
    console.log(JSON.stringify(data, null, 2));
  }
}