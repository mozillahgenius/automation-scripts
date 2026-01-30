// ==========================================
// 改良版：重要議題に限定したSlack分析システム
// ==========================================

// ========= 重要な論点のみをフィルタリング =========
function filterCriticalTopics(topics, priority) {
  if (!topics || topics.length === 0) return [];

  // HIGHの場合は全て重要
  if (priority === 'HIGH') {
    return topics.slice(0, 3); // 最大3つまで
  }

  // MEDIUMの場合は上位2つ
  if (priority === 'MEDIUM') {
    return topics.slice(0, 2);
  }

  // LOWの場合は1つだけ、または空配列
  if (priority === 'LOW') {
    // 本当に重要そうなキーワードが含まれる場合のみ1つ返す
    const criticalKeywords = [
      '予算', '承認', '決裁', '契約', '法的', 'リスク', '損失',
      '取締役', '監査', '開示', 'M&A', '投資', '人事', '組織'
    ];

    const criticalTopic = topics.find(topic => {
      const text = (topic.title + ' ' + (topic.description || '')).toLowerCase();
      return criticalKeywords.some(keyword => text.includes(keyword));
    });

    return criticalTopic ? [criticalTopic] : [];
  }

  return topics;
}

// ========= 詳細業務フロー生成 =========
function generateDetailedBusinessFlow(analysisResult, governanceResult, messages) {
  const workflow = [];

  // メッセージから具体的なタスクを抽出
  const extractedTasks = extractTasksFromMessages(messages);

  // 1. 現状分析・課題整理
  workflow.push({
    process: '現状分析・課題整理',
    detail: `${analysisResult.summary || '議題内容の確認と現状把握'}`,
    responsible: '担当者',
    accountable: '部門長',
    consulted: '関係部署',
    informed: 'チームメンバー',
    deadline: '1営業日'
  });

  // 2. 論点ごとの具体的なアクション
  if (analysisResult.topics && analysisResult.topics.length > 0) {
    // フィルタリングされた重要論点のみ処理
    const filteredTopics = filterCriticalTopics(analysisResult.topics, analysisResult.priority);

    filteredTopics.forEach((topic, index) => {
      workflow.push({
        process: `論点${index + 1}: ${topic.title}`,
        detail: topic.description || '詳細検討',
        responsible: analysisResult.stakeholders?.[0] || '担当者',
        accountable: '部門長',
        consulted: '専門チーム',
        informed: '関係部署',
        deadline: priorityToDeadline(topic.priority || 3)
      });
    });
  }

  // 3. アクションアイテムの実行（重要なもののみ）
  if (analysisResult.actionItems && analysisResult.actionItems.length > 0) {
    // 重要度HIGHまたはMEDIUMの場合のみアクションアイテムを含める
    if (analysisResult.priority === 'HIGH' || analysisResult.priority === 'MEDIUM') {
      analysisResult.actionItems.slice(0, 3).forEach((item, index) => {
        workflow.push({
          process: `アクション${index + 1}`,
          detail: item.task,
          responsible: item.owner || '担当者',
          accountable: '部門長',
          consulted: '関連チーム',
          informed: '関係者',
          deadline: item.deadline || '5営業日'
        });
      });
    }
  }

  // 4. 専門家相談が必要な場合
  if (governanceResult.requiresExpertConsultation) {
    workflow.push({
      process: '専門家相談',
      detail: `${governanceResult.requiredExperts?.join('、') || '専門家'}への相談と助言取得`,
      responsible: '法務・総務',
      accountable: '部門長',
      consulted: governanceResult.requiredExperts?.join('・') || '専門家',
      informed: '経営陣',
      deadline: '3営業日'
    });
  }

  // 5. 承認が必要な場合
  if (governanceResult.requiresApproval) {
    const approvalSteps = generateApprovalSteps(governanceResult.approvalLevel, analysisResult);
    approvalSteps.forEach(step => workflow.push(step));
  }

  // 6. 開示が必要な場合
  if (governanceResult.requiresDisclosure) {
    workflow.push({
      process: '開示資料作成',
      detail: `${governanceResult.disclosureType || '適時開示'}資料の作成と確認`,
      responsible: 'IR担当',
      accountable: 'CFO',
      consulted: '法務・会計士',
      informed: '取締役会',
      deadline: '適時'
    });
  }

  // 7. 実行・モニタリング
  workflow.push({
    process: '実行・モニタリング',
    detail: '決定事項の実施と進捗管理',
    responsible: extractedTasks.primary_owner || '担当者',
    accountable: '部門長',
    consulted: '関係部署',
    informed: '関係者全員',
    deadline: '計画通り'
  });

  // 8. フォローアップ（重要案件のみ）
  if (analysisResult.priority === 'HIGH') {
    workflow.push({
      process: 'フォローアップ',
      detail: '実施結果の評価と次のアクション検討',
      responsible: '部門長',
      accountable: '経営陣',
      consulted: '評価チーム',
      informed: '取締役会',
      deadline: '実施後1週間'
    });
  }

  return workflow;
}

// メッセージからタスクを抽出
function extractTasksFromMessages(messages) {
  const tasks = {
    primary_owner: null,
    tasks_list: []
  };

  if (!messages) return tasks;

  messages.forEach(msg => {
    if (!msg.text) return;
    // @メンションがあれば担当者として抽出
    const mentions = msg.text.match(/@[\w\-]+/g);
    if (mentions && !tasks.primary_owner) {
      tasks.primary_owner = mentions[0].replace('@', '');
    }

    // タスクキーワードを検出
    const taskKeywords = ['する', 'します', 'してください', 'お願い', '依頼', 'TODO'];
    if (taskKeywords.some(kw => msg.text.includes(kw))) {
      tasks.tasks_list.push({
        text: msg.text.substring(0, 100),
        owner: mentions ? mentions[0].replace('@', '') : null
      });
    }
  });

  return tasks;
}

// 優先度から期限を算出
function priorityToDeadline(priority) {
  if (priority === 1) return '2営業日';
  if (priority === 2) return '3営業日';
  return '5営業日';
}

// 承認ステップを生成
function generateApprovalSteps(approvalLevel, analysisResult) {
  const steps = [];

  if (approvalLevel === '部長承認') {
    steps.push({
      process: '部長承認',
      detail: '部門内での意思決定と承認',
      responsible: '担当者',
      accountable: '部長',
      consulted: '課長',
      informed: 'チーム',
      deadline: '2営業日'
    });
  } else if (approvalLevel === '取締役会') {
    steps.push({
      process: '取締役会付議',
      detail: '取締役会資料作成と上程',
      responsible: '部門長',
      accountable: 'CEO',
      consulted: '法務・財務',
      informed: '監査役',
      deadline: '5営業日'
    });
  } else if (approvalLevel === '株主総会') {
    steps.push({
      process: '株主総会決議',
      detail: '株主総会招集と決議',
      responsible: '取締役会',
      accountable: '代表取締役',
      consulted: '法務・監査役',
      informed: '全株主',
      deadline: '法定期限'
    });
  }

  return steps;
}

// ========= スプレッドシートへの記載を最適化 =========
function writeFilteredTopicsToSheet(sheet, analysisResult, governanceResult) {
  const filteredTopics = filterCriticalTopics(analysisResult.topics, analysisResult.priority);

  if (filteredTopics.length === 0) {
    // 重要な論点がない場合は、サマリーのみ記載
    sheet.getRange(1, 1).setValue('重要な議題はありません');
    sheet.getRange(2, 1).setValue(`分析概要: ${analysisResult.summary}`);
    return;
  }

  // ヘッダー
  const headers = ['No', '論点', '詳細', '重要度', 'アクション', '期限', '担当者'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#4285f4').setFontColor('#ffffff');

  // 重要論点のみを記載
  const rows = [];
  filteredTopics.forEach((topic, index) => {
    // 対応するアクションアイテムを取得
    const relatedAction = analysisResult.actionItems?.find(item =>
      item.task?.includes(topic.title) || topic.description?.includes(item.task)
    );

    rows.push([
      index + 1,
      topic.title,
      topic.description || '',
      analysisResult.priority,
      relatedAction?.task || '要検討',
      relatedAction?.deadline || priorityToDeadline(topic.priority || 3),
      relatedAction?.owner || analysisResult.stakeholders?.[0] || '未定'
    ]);
  });

  sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);

  // ガバナンス要件を追記（重要案件のみ）
  if (analysisResult.priority === 'HIGH' || governanceResult.requiresAction) {
    const govRow = sheet.getLastRow() + 2;
    sheet.getRange(govRow, 1).setValue('【ガバナンス要件】');
    sheet.getRange(govRow, 1).setFontWeight('bold').setBackground('#ea4335');

    const govData = [];
    if (governanceResult.requiresApproval) {
      govData.push(['承認レベル', governanceResult.approvalLevel]);
    }
    if (governanceResult.requiresDisclosure) {
      govData.push(['開示種別', governanceResult.disclosureType]);
    }
    if (governanceResult.riskLevel === 'HIGH') {
      govData.push(['リスク評価', governanceResult.riskReasons?.join('、')]);
    }

    if (govData.length > 0) {
      sheet.getRange(govRow + 1, 1, govData.length, 2).setValues(govData);
    }
  }
}

// ========= 使用例 =========
function testEnhancedSystem() {
  // テストメッセージ
  const testMessages = [
    { text: '来期予算3000万円の承認をお願いします', user: 'user1', ts: Date.now() / 1000 },
    { text: '会議の日程調整です', user: 'user2', ts: Date.now() / 1000 },
    { text: 'システム障害が発生しました', user: 'user3', ts: Date.now() / 1000 }
  ];

  // AI分析結果（モック）
  const analysisResult = {
    categories: ['予算', 'システム'],
    topics: [
      { title: '予算承認', description: '3000万円の来期予算承認', priority: 1 },
      { title: '会議調整', description: '定例会議の日程調整', priority: 3 },
      { title: 'システム障害', description: 'システム障害対応', priority: 2 }
    ],
    priority: 'HIGH',
    priorityReason: '1000万円を超える予算承認が必要',
    actionItems: [
      { task: '予算申請書作成', owner: '財務部', deadline: '3営業日' }
    ],
    stakeholders: ['財務部', 'IT部'],
    summary: '来期予算承認とシステム障害対応'
  };

  // ガバナンスチェック結果（モック）
  const governanceResult = {
    requiresApproval: true,
    approvalLevel: '取締役会',
    requiresDisclosure: false,
    riskLevel: 'MEDIUM',
    requiresAction: true
  };

  // フィルタリング実行
  const filtered = filterCriticalTopics(analysisResult.topics, analysisResult.priority);
  console.log('フィルタリング後の論点:', filtered);

  // 詳細フロー生成
  const detailedFlow = generateDetailedBusinessFlow(analysisResult, governanceResult, testMessages);
  console.log('詳細業務フロー:', detailedFlow);

  // スプレッドシートに記載（テスト用）
  const ss = SpreadsheetApp.create('テスト_重要議題のみ');
  const sheet = ss.getActiveSheet();
  writeFilteredTopicsToSheet(sheet, analysisResult, governanceResult);

  console.log('テスト完了。スプレッドシートURL:', ss.getUrl());
}