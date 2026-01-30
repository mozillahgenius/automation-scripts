/**
 * 外部専門家相談管理システム
 * 重要事項の実行前に外部専門家への相談を必須化
 */

// ================================================================================
// 外部専門家マスターデータ
// ================================================================================

const EXTERNAL_ADVISORS = {
  '法律事務所': {
    specialties: [
      'M&A・企業再編',
      'コーポレートガバナンス',
      '株主総会対応',
      '取締役会運営',
      '契約法務',
      'コンプライアンス',
      '労働法',
      '知的財産権',
      '訴訟・紛争解決',
      '金融商品取引法',
      '独占禁止法',
      '個人情報保護法'
    ],
    consultationTiming: '重要な法的判断が必要な段階の初期',
    deliverables: [
      'リーガルオピニオン',
      '契約書レビュー',
      'デューデリジェンス報告書',
      '法的リスク分析書'
    ],
    urgencyLevels: {
      'CRITICAL': '即日対応',
      'HIGH': '2-3営業日',
      'MEDIUM': '1週間',
      'LOW': '2週間'
    }
  },
  
  '監査法人': {
    specialties: [
      '会計監査',
      '内部統制監査（J-SOX）',
      '四半期レビュー',
      'IPO支援',
      'IFRS導入支援',
      '不正調査',
      'M&Aデューデリジェンス（財務）',
      '企業価値評価'
    ],
    consultationTiming: '決算前・重要な会計処理の変更前',
    deliverables: [
      '監査報告書',
      '内部統制監査報告書',
      'マネジメントレター',
      '改善提案書'
    ],
    urgencyLevels: {
      'CRITICAL': '即日対応',
      'HIGH': '3-5営業日',
      'MEDIUM': '2週間',
      'LOW': '1か月'
    }
  },
  
  '会計事務所・税理士事務所': {
    specialties: [
      '税務申告',
      '税務調査対応',
      '移転価格',
      '国際税務',
      '組織再編税制',
      'M&A税務',
      '事業承継',
      'タックスプランニング',
      '消費税',
      '法人税'
    ],
    consultationTiming: '税務判断が必要な取引の実行前',
    deliverables: [
      '税務意見書',
      '税務デューデリジェンス報告書',
      'タックスストラクチャリング提案',
      '税務申告書'
    ],
    urgencyLevels: {
      'CRITICAL': '即日対応',
      'HIGH': '2-3営業日',
      'MEDIUM': '1週間',
      'LOW': '2週間'
    }
  },
  
  '司法書士事務所': {
    specialties: [
      '商業登記',
      '不動産登記',
      '役員変更登記',
      '定款変更',
      '増資・減資登記',
      '合併・分割登記',
      '本店移転',
      '支店設置',
      '株式関連手続き'
    ],
    consultationTiming: '登記が必要な決議の前',
    deliverables: [
      '登記申請書',
      '定款',
      '議事録作成支援',
      '登記完了証明書'
    ],
    urgencyLevels: {
      'CRITICAL': '当日対応',
      'HIGH': '1-2営業日',
      'MEDIUM': '3-5営業日',
      'LOW': '1週間'
    }
  },
  
  '社会保険労務士事務所': {
    specialties: [
      '就業規則作成・変更',
      '労働契約',
      '労使協定',
      '社会保険手続き',
      '労働基準監督署対応',
      '労災対応',
      '人事制度設計',
      '給与計算',
      'ハラスメント対策',
      '労使紛争解決'
    ],
    consultationTiming: '人事労務施策の実施前',
    deliverables: [
      '就業規則',
      '労使協定書',
      '労務監査報告書',
      '改善提案書'
    ],
    urgencyLevels: {
      'CRITICAL': '即日対応',
      'HIGH': '2-3営業日',
      'MEDIUM': '1週間',
      'LOW': '2週間'
    }
  },
  
  '特許事務所': {
    specialties: [
      '特許出願',
      '商標登録',
      '意匠登録',
      '知財戦略立案',
      '侵害調査',
      'ライセンス契約',
      '知財デューデリジェンス',
      '知財訴訟支援'
    ],
    consultationTiming: '新製品・サービス開発の初期段階',
    deliverables: [
      '特許明細書',
      '先行技術調査報告書',
      '侵害鑑定書',
      '知財戦略提案書'
    ],
    urgencyLevels: {
      'CRITICAL': '即日対応',
      'HIGH': '3-5営業日',
      'MEDIUM': '2週間',
      'LOW': '1か月'
    }
  },
  
  'コンサルティングファーム': {
    specialties: [
      '経営戦略',
      '事業戦略',
      'PMI（買収後統合）',
      'デジタルトランスフォーメーション',
      '業務改革',
      'リスクマネジメント',
      'サイバーセキュリティ',
      'ESG・サステナビリティ'
    ],
    consultationTiming: '戦略的意思決定の前',
    deliverables: [
      '戦略提案書',
      'フィージビリティスタディ',
      '実行計画書',
      'プロジェクト管理支援'
    ],
    urgencyLevels: {
      'CRITICAL': '1週間',
      'HIGH': '2週間',
      'MEDIUM': '1か月',
      'LOW': '2か月'
    }
  },
  
  '証券会社・投資銀行': {
    specialties: [
      'IPO（新規上場）',
      '資金調達',
      'M&Aアドバイザリー',
      '株式公開買付（TOB）',
      'IR支援',
      '株価算定',
      'フェアネス・オピニオン'
    ],
    consultationTiming: '資本政策・M&A検討の初期段階',
    deliverables: [
      'バリュエーション報告書',
      'フェアネス・オピニオン',
      'IM（インフォメーション・メモランダム）',
      'プロセスレター'
    ],
    urgencyLevels: {
      'CRITICAL': '3営業日',
      'HIGH': '1週間',
      'MEDIUM': '2週間',
      'LOW': '1か月'
    }
  }
};

// ================================================================================
// 相談要否判定ロジック
// ================================================================================

/**
 * タスクに対して必要な外部専門家を判定
 */
function determineRequiredAdvisors(task, context = {}) {
  const requiredAdvisors = [];
  const taskLower = task.toLowerCase();
  
  // 必須相談パターンの定義
  const mandatoryPatterns = [
    {
      keywords: ['株主総会', '株主', '総会'],
      advisors: ['法律事務所', '司法書士事務所'],
      reason: '株主総会の適法な運営と手続きの確認',
      checkpoints: [
        '招集手続きの適法性確認',
        '議案の適法性確認',
        '決議要件の確認',
        '議事録作成要領の確認'
      ]
    },
    {
      keywords: ['取締役会', '取締役', '役員', '執行役'],
      advisors: ['法律事務所', '司法書士事務所'],
      reason: '取締役会運営の適法性と役員変更登記',
      checkpoints: [
        '決議事項の適法性確認',
        '利益相反取引の確認',
        '特別利害関係の確認',
        '登記手続きの確認'
      ]
    },
    {
      keywords: ['M&A', '買収', '合併', '事業譲渡', '会社分割', '株式交換', '株式移転'],
      advisors: ['法律事務所', '監査法人', '税理士事務所', '証券会社・投資銀行'],
      reason: 'M&A取引の法務・財務・税務面での総合的検証',
      checkpoints: [
        'ストラクチャーの検討',
        'デューデリジェンスの実施',
        '価格の妥当性検証',
        '契約条件の交渉',
        'クロージング手続きの確認',
        'PMI計画の策定'
      ]
    },
    {
      keywords: ['決算', '財務諸表', '有価証券報告書', '四半期報告書'],
      advisors: ['監査法人', '税理士事務所'],
      reason: '適正な財務報告と税務申告',
      checkpoints: [
        '会計処理の妥当性確認',
        '開示内容の適切性確認',
        '内部統制の有効性評価',
        '税務リスクの確認'
      ]
    },
    {
      keywords: ['増資', '減資', '自己株式', '新株', '社債'],
      advisors: ['法律事務所', '証券会社・投資銀行', '司法書士事務所'],
      reason: '資本政策の適法性と実行可能性の確認',
      checkpoints: [
        '発行条件の妥当性',
        '既存株主への影響分析',
        '開示書類の作成',
        '登記手続きの準備'
      ]
    },
    {
      keywords: ['上場', 'IPO', '株式公開'],
      advisors: ['証券会社・投資銀行', '監査法人', '法律事務所'],
      reason: 'IPO準備の包括的支援',
      checkpoints: [
        '上場基準の充足確認',
        '内部管理体制の整備',
        '開示体制の構築',
        '審査対応準備'
      ]
    },
    {
      keywords: ['労働', '雇用', '解雇', '就業規則', 'ハラスメント'],
      advisors: ['社会保険労務士事務所', '法律事務所'],
      reason: '労働法令遵守と労使紛争の予防',
      checkpoints: [
        '労働法令の遵守確認',
        '就業規則の整備',
        '労使協定の締結',
        '紛争リスクの評価'
      ]
    },
    {
      keywords: ['特許', '商標', '知財', '著作権', 'ライセンス'],
      advisors: ['特許事務所', '法律事務所'],
      reason: '知的財産権の保護と活用',
      checkpoints: [
        '権利化戦略の検討',
        '侵害リスクの確認',
        'ライセンス条件の検討',
        '紛争対応策の準備'
      ]
    },
    {
      keywords: ['契約', '締結', '変更', '解除'],
      advisors: ['法律事務所'],
      reason: '契約リスクの評価と条件交渉',
      checkpoints: [
        '契約条件の妥当性確認',
        'リスク条項の確認',
        '責任範囲の明確化',
        '紛争解決条項の確認'
      ]
    },
    {
      keywords: ['コンプライアンス', '違反', '不正', '内部統制'],
      advisors: ['法律事務所', '監査法人', 'コンサルティングファーム'],
      reason: 'コンプライアンス体制の強化と違反防止',
      checkpoints: [
        '現状のリスク評価',
        '改善策の立案',
        'モニタリング体制の構築',
        '教育研修の実施'
      ]
    },
    {
      keywords: ['訴訟', '紛争', '係争', '調停', '仲裁'],
      advisors: ['法律事務所'],
      reason: '法的紛争の適切な解決',
      checkpoints: [
        '勝訴可能性の評価',
        '和解条件の検討',
        '証拠の収集・保全',
        '訴訟戦略の立案'
      ]
    },
    {
      keywords: ['個人情報', 'プライバシー', 'GDPR', '情報漏洩'],
      advisors: ['法律事務所', 'コンサルティングファーム'],
      reason: '個人情報保護法令の遵守',
      checkpoints: [
        '現行体制の評価',
        '規程・手順の整備',
        'セキュリティ対策の確認',
        'インシデント対応体制の構築'
      ]
    }
  ];

  // パターンマッチング
  mandatoryPatterns.forEach(pattern => {
    const hasKeyword = pattern.keywords.some(keyword => taskLower.includes(keyword));
    if (hasKeyword) {
      pattern.advisors.forEach(advisor => {
        requiredAdvisors.push({
          type: advisor,
          reason: pattern.reason,
          priority: 'MANDATORY',
          checkpoints: pattern.checkpoints,
          timing: EXTERNAL_ADVISORS[advisor].consultationTiming
        });
      });
    }
  });

  // 金額基準での判定
  if (context.amount) {
    const amount = parseInt(context.amount);
    if (amount > 100000000) { // 1億円以上
      requiredAdvisors.push({
        type: '法律事務所',
        reason: '高額取引のため法的リスク評価が必要',
        priority: 'HIGH',
        checkpoints: ['契約条件の精査', 'リスク分析', '交渉戦略']
      });
    }
    if (amount > 500000000) { // 5億円以上
      requiredAdvisors.push({
        type: '監査法人',
        reason: '重要な取引のため会計上の影響評価が必要',
        priority: 'HIGH',
        checkpoints: ['会計処理の検討', '開示への影響', '内部統制への影響']
      });
    }
  }

  // 緊急度による追加判定
  if (context.urgency === 'CRITICAL') {
    // 既存のアドバイザーの優先度を引き上げ
    requiredAdvisors.forEach(advisor => {
      advisor.priority = 'CRITICAL';
      advisor.responseTime = EXTERNAL_ADVISORS[advisor.type].urgencyLevels['CRITICAL'];
    });
  }

  return requiredAdvisors;
}

/**
 * 外部専門家相談チェックリスト生成
 */
function generateConsultationChecklist(task, advisors) {
  const checklist = {
    task: task,
    consultationSteps: [],
    documentationRequired: [],
    timeline: [],
    budgetConsiderations: []
  };

  // ステップ1: 事前準備
  checklist.consultationSteps.push({
    step: 1,
    phase: '事前準備',
    actions: [
      '相談事項の明確化と論点整理',
      '関連資料の収集と整理',
      '社内での事前検討と方針案の作成',
      '予算の確保と決裁取得'
    ],
    timeline: 'T-14日',
    responsible: '担当部門'
  });

  // ステップ2: 専門家選定
  checklist.consultationSteps.push({
    step: 2,
    phase: '専門家選定',
    actions: [
      '複数の専門家候補のリストアップ',
      '見積もり取得と比較検討',
      '利益相反チェック',
      '秘密保持契約（NDA）の締結'
    ],
    timeline: 'T-10日',
    responsible: '法務部・総務部'
  });

  // ステップ3: 各専門家への相談実施
  advisors.forEach((advisor, index) => {
    checklist.consultationSteps.push({
      step: 3 + index,
      phase: `${advisor.type}への相談`,
      actions: [
        '初回ミーティングの実施',
        '詳細情報の提供と質疑応答',
        '中間報告の受領とフィードバック',
        '最終意見書・報告書の受領'
      ],
      timeline: `T-${7 - index}日`,
      responsible: `担当部門・${advisor.type}`,
      deliverables: EXTERNAL_ADVISORS[advisor.type].deliverables,
      checkpoints: advisor.checkpoints || []
    });
  });

  // ステップ4: 社内検討
  checklist.consultationSteps.push({
    step: 3 + advisors.length,
    phase: '社内検討・意思決定',
    actions: [
      '専門家意見の社内共有と検討',
      'リスク評価と対応策の決定',
      '実行計画の策定',
      '必要な社内承認の取得'
    ],
    timeline: 'T-2日',
    responsible: '経営陣・関連部門'
  });

  // ステップ5: 実行とフォローアップ
  checklist.consultationSteps.push({
    step: 4 + advisors.length,
    phase: '実行・フォローアップ',
    actions: [
      '決定事項の実行',
      '専門家との継続的な連携',
      '実行状況のモニタリング',
      '追加相談の必要性評価'
    ],
    timeline: 'T-0日以降',
    responsible: '実行責任部門'
  });

  // 必要書類リスト
  checklist.documentationRequired = [
    '相談依頼書',
    '背景説明資料',
    '関連契約書・規程類',
    '財務データ（必要に応じて）',
    '過去の類似案件資料',
    '社内検討資料',
    '取締役会・経営会議資料'
  ];

  // 予算考慮事項
  checklist.budgetConsiderations = [
    {
      item: 'アドバイザリー費用',
      estimation: '案件規模の0.5-2%程度',
      factors: ['案件の複雑さ', '緊急度', '専門家の知名度']
    },
    {
      item: 'デューデリジェンス費用',
      estimation: '500万円〜5000万円',
      factors: ['調査範囲', '対象会社の規模', '調査期間']
    },
    {
      item: 'その他実費',
      estimation: '100万円〜500万円',
      factors: ['交通費', '翻訳費用', '印刷費用']
    }
  ];

  return checklist;
}

/**
 * 相談記録管理
 */
function createConsultationRecord(consultation) {
  return {
    id: Utilities.getUuid(),
    date: new Date(),
    task: consultation.task,
    advisorType: consultation.advisorType,
    advisorName: consultation.advisorName,
    contactPerson: consultation.contactPerson,
    consultationSummary: consultation.summary,
    adviceReceived: consultation.advice,
    documentsProvided: consultation.documents || [],
    documentsReceived: consultation.receivedDocs || [],
    actionItems: consultation.actionItems || [],
    followUpRequired: consultation.followUp || false,
    followUpDate: consultation.followUpDate || null,
    cost: consultation.cost || 0,
    status: consultation.status || 'IN_PROGRESS',
    internalAttendees: consultation.internalAttendees || [],
    nextSteps: consultation.nextSteps || [],
    risks: consultation.risks || [],
    decisions: consultation.decisions || []
  };
}

/**
 * 専門家意見の統合分析
 */
function analyzeExpertOpinions(opinions) {
  const analysis = {
    consensus: [],
    divergence: [],
    risks: [],
    recommendations: [],
    actionPlan: []
  };

  // 意見の集約
  const opinionMap = new Map();
  opinions.forEach(opinion => {
    const key = opinion.topic || 'general';
    if (!opinionMap.has(key)) {
      opinionMap.set(key, []);
    }
    opinionMap.get(key).push(opinion);
  });

  // コンセンサスと相違点の分析
  opinionMap.forEach((topicOpinions, topic) => {
    if (topicOpinions.length > 1) {
      // 共通見解を抽出
      const commonPoints = findCommonPoints(topicOpinions);
      if (commonPoints.length > 0) {
        analysis.consensus.push({
          topic: topic,
          points: commonPoints,
          advisors: topicOpinions.map(o => o.advisorType)
        });
      }

      // 相違点を抽出
      const differences = findDifferences(topicOpinions);
      if (differences.length > 0) {
        analysis.divergence.push({
          topic: topic,
          differences: differences,
          requiresFurtherDiscussion: true
        });
      }
    }
  });

  // リスク統合
  opinions.forEach(opinion => {
    if (opinion.risks && opinion.risks.length > 0) {
      opinion.risks.forEach(risk => {
        const existingRisk = analysis.risks.find(r => r.description === risk.description);
        if (existingRisk) {
          existingRisk.mentionedBy.push(opinion.advisorType);
          existingRisk.severity = Math.max(existingRisk.severity, risk.severity || 0);
        } else {
          analysis.risks.push({
            description: risk.description,
            severity: risk.severity || 'MEDIUM',
            mentionedBy: [opinion.advisorType],
            mitigation: risk.mitigation || ''
          });
        }
      });
    }
  });

  // 推奨事項の優先順位付け
  const allRecommendations = [];
  opinions.forEach(opinion => {
    if (opinion.recommendations) {
      opinion.recommendations.forEach(rec => {
        allRecommendations.push({
          recommendation: rec,
          source: opinion.advisorType,
          priority: opinion.priority || 'MEDIUM'
        });
      });
    }
  });

  // 優先度でソート
  allRecommendations.sort((a, b) => {
    const priorityOrder = { 'CRITICAL': 0, 'HIGH': 1, 'MEDIUM': 2, 'LOW': 3 };
    return priorityOrder[a.priority] - priorityOrder[b.priority];
  });

  analysis.recommendations = allRecommendations;

  // 統合アクションプラン
  analysis.actionPlan = generateIntegratedActionPlan(analysis);

  return analysis;
}

/**
 * 統合アクションプラン生成
 */
function generateIntegratedActionPlan(analysis) {
  const actionPlan = [];
  let stepNumber = 1;

  // リスク対応を最優先
  analysis.risks
    .filter(r => r.severity === 'CRITICAL' || r.severity === 'HIGH')
    .forEach(risk => {
      actionPlan.push({
        step: stepNumber++,
        action: `リスク対応: ${risk.description}`,
        priority: 'CRITICAL',
        responsible: '法務部・リスク管理部',
        timeline: '即時',
        details: risk.mitigation
      });
    });

  // コンセンサスに基づく行動
  analysis.consensus.forEach(consensus => {
    consensus.points.forEach(point => {
      actionPlan.push({
        step: stepNumber++,
        action: point,
        priority: 'HIGH',
        responsible: '担当部門',
        timeline: '計画通り実行',
        supportedBy: consensus.advisors
      });
    });
  });

  // 相違点の解決
  analysis.divergence.forEach(divergence => {
    actionPlan.push({
      step: stepNumber++,
      action: `追加検討: ${divergence.topic}`,
      priority: 'MEDIUM',
      responsible: '経営企画部',
      timeline: '1週間以内',
      details: '専門家間の意見相違について追加協議が必要'
    });
  });

  // 推奨事項の実行
  analysis.recommendations.slice(0, 10).forEach(rec => {
    actionPlan.push({
      step: stepNumber++,
      action: rec.recommendation,
      priority: rec.priority,
      responsible: '関連部門',
      timeline: '優先度に応じて',
      recommendedBy: rec.source
    });
  });

  return actionPlan;
}

/**
 * 共通見解の抽出（ヘルパー関数）
 */
function findCommonPoints(opinions) {
  // 実装は簡略化
  const commonPoints = [];
  const allPoints = opinions.flatMap(o => o.keyPoints || []);
  
  // 頻度分析
  const pointFrequency = {};
  allPoints.forEach(point => {
    pointFrequency[point] = (pointFrequency[point] || 0) + 1;
  });

  // 複数の専門家が言及した点を共通見解とする
  Object.entries(pointFrequency).forEach(([point, count]) => {
    if (count > 1) {
      commonPoints.push(point);
    }
  });

  return commonPoints;
}

/**
 * 相違点の抽出（ヘルパー関数）
 */
function findDifferences(opinions) {
  const differences = [];
  
  // 各意見の立場を比較
  opinions.forEach((opinion, index) => {
    opinions.slice(index + 1).forEach(otherOpinion => {
      if (opinion.stance && otherOpinion.stance && opinion.stance !== otherOpinion.stance) {
        differences.push({
          advisor1: opinion.advisorType,
          stance1: opinion.stance,
          advisor2: otherOpinion.advisorType,
          stance2: otherOpinion.stance,
          topic: opinion.topic
        });
      }
    });
  });

  return differences;
}

/**
 * 相談優先度スコアリング
 */
function calculateConsultationPriority(task, context = {}) {
  let score = 0;
  const factors = [];

  // 法的リスク
  if (task.includes('契約') || task.includes('法') || task.includes('規制')) {
    score += 30;
    factors.push('法的リスクあり');
  }

  // 財務インパクト
  if (context.financialImpact > 100000000) {
    score += 40;
    factors.push('重要な財務影響');
  }

  // 開示義務
  if (task.includes('開示') || task.includes('報告') || task.includes('株主')) {
    score += 35;
    factors.push('開示義務あり');
  }

  // 緊急度
  if (context.urgency === 'CRITICAL') {
    score += 25;
    factors.push('緊急対応必要');
  }

  // 前例なし
  if (context.noPrecedent) {
    score += 20;
    factors.push('前例なし');
  }

  // ステークホルダー影響
  if (context.stakeholderImpact === 'HIGH') {
    score += 15;
    factors.push('ステークホルダー影響大');
  }

  return {
    score: Math.min(score, 100),
    priority: score >= 70 ? 'CRITICAL' : score >= 50 ? 'HIGH' : score >= 30 ? 'MEDIUM' : 'LOW',
    factors: factors,
    recommendation: score >= 50 ? '外部専門家への相談を強く推奨' : '必要に応じて相談を検討'
  };
}

/**
 * 専門家相談フローのビジュアル化データ生成
 */
function generateConsultationFlowData(task, requiredAdvisors) {
  const flowData = {
    nodes: [],
    edges: [],
    timeline: []
  };

  // 開始ノード
  flowData.nodes.push({
    id: 'start',
    label: 'タスク開始',
    type: 'start',
    position: { x: 0, y: 0 }
  });

  // 社内検討ノード
  flowData.nodes.push({
    id: 'internal_review',
    label: '社内事前検討',
    type: 'process',
    position: { x: 0, y: 100 },
    duration: '2-3日'
  });

  flowData.edges.push({
    from: 'start',
    to: 'internal_review'
  });

  // 各専門家相談ノード
  let yPos = 200;
  requiredAdvisors.forEach((advisor, index) => {
    const nodeId = `advisor_${index}`;
    flowData.nodes.push({
      id: nodeId,
      label: advisor.type,
      type: 'consultation',
      position: { x: index * 150, y: yPos },
      priority: advisor.priority,
      duration: EXTERNAL_ADVISORS[advisor.type].urgencyLevels[advisor.priority] || '1週間'
    });

    flowData.edges.push({
      from: 'internal_review',
      to: nodeId,
      label: advisor.reason
    });
  });

  // 意見統合ノード
  flowData.nodes.push({
    id: 'integration',
    label: '専門家意見の統合',
    type: 'process',
    position: { x: 0, y: 300 },
    duration: '1-2日'
  });

  requiredAdvisors.forEach((advisor, index) => {
    flowData.edges.push({
      from: `advisor_${index}`,
      to: 'integration'
    });
  });

  // 意思決定ノード
  flowData.nodes.push({
    id: 'decision',
    label: '最終意思決定',
    type: 'decision',
    position: { x: 0, y: 400 }
  });

  flowData.edges.push({
    from: 'integration',
    to: 'decision'
  });

  // 実行ノード
  flowData.nodes.push({
    id: 'execution',
    label: 'タスク実行',
    type: 'end',
    position: { x: 0, y: 500 }
  });

  flowData.edges.push({
    from: 'decision',
    to: 'execution',
    label: '承認'
  });

  return flowData;
}

/**
 * 過去の相談履歴から学習
 */
function learnFromConsultationHistory(currentTask) {
  // 実際の実装では、過去の相談記録DBから類似案件を検索
  const suggestions = {
    similarCases: [],
    lessonsLearned: [],
    recommendedAdvisors: [],
    estimatedTimeline: null,
    estimatedCost: null
  };

  // ダミーデータ（実際はDBから取得）
  const historicalData = [
    {
      task: 'M&A実行',
      advisors: ['法律事務所', '監査法人', '税理士事務所'],
      duration: 45,
      cost: 50000000,
      outcome: 'SUCCESS',
      lessons: ['早期の専門家巻き込みが成功の鍵', 'DD期間は余裕を持って設定']
    }
  ];

  // 類似案件の抽出
  historicalData.forEach(history => {
    const similarity = calculateSimilarity(currentTask, history.task);
    if (similarity > 0.7) {
      suggestions.similarCases.push(history);
      suggestions.lessonsLearned.push(...history.lessons);
      suggestions.recommendedAdvisors.push(...history.advisors);
      suggestions.estimatedTimeline = history.duration;
      suggestions.estimatedCost = history.cost;
    }
  });

  return suggestions;
}

/**
 * タスク類似度計算（簡易版）
 */
function calculateSimilarity(task1, task2) {
  // 実際の実装では、より高度なテキスト類似度アルゴリズムを使用
  const words1 = task1.toLowerCase().split(/\s+/);
  const words2 = task2.toLowerCase().split(/\s+/);
  
  const commonWords = words1.filter(word => words2.includes(word));
  const similarity = commonWords.length / Math.max(words1.length, words2.length);
  
  return similarity;
}