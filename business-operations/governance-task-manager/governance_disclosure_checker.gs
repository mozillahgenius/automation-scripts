/**
 * ガバナンス・開示判定システム
 * 東証・財務局への開示要件を体系的にチェック
 */

// ================================================================================
// 開示判定マスターデータ
// ================================================================================

const DISCLOSURE_TRIGGERS = {
  // 適時開示（東証）が必要な決定事実
  TIMELY_DISCLOSURE_DECISIONS: {
    '株式発行': {
      criteria: ['新株発行', '増資', '第三者割当', '公募', '株主割当'],
      threshold: '発行済株式総数の10%以上',
      timeline: '決議後直ちに',
      authority: '取締役会決議',
      documents: ['有価証券届出書', '適時開示資料'],
      regulations: ['金商法第4条', '東証適時開示規則第2条']
    },
    '資本政策': {
      criteria: ['自己株式取得', '資本金減少', '株式分割', '株式併合'],
      threshold: '資本金の10%以上',
      timeline: '決議後直ちに',
      authority: '取締役会決議（一部株主総会）',
      documents: ['適時開示資料', '臨時報告書'],
      regulations: ['会社法第156条', '東証適時開示規則']
    },
    'M&A': {
      criteria: ['合併', '会社分割', '株式交換', '株式移転', '事業譲渡'],
      threshold: '純資産の30%以上',
      timeline: '基本合意時及び決議後',
      authority: '取締役会決議及び株主総会特別決議',
      documents: ['適時開示資料', '臨時報告書', '公開買付届出書'],
      regulations: ['会社法第783条', '金商法第27条の3']
    },
    '業務提携': {
      criteria: ['資本提携', '業務提携', '技術提携'],
      threshold: '売上高の10%以上の影響',
      timeline: '契約締結後直ちに',
      authority: '取締役会決議',
      documents: ['適時開示資料'],
      regulations: ['東証適時開示規則第2条']
    },
    '重要な契約': {
      criteria: ['ライセンス契約', '販売契約', '製造委託契約'],
      threshold: '売上高の10%以上',
      timeline: '契約締結後直ちに',
      authority: '取締役会決議',
      documents: ['適時開示資料'],
      regulations: ['東証適時開示規則']
    }
  },

  // 適時開示が必要な発生事実
  TIMELY_DISCLOSURE_EVENTS: {
    '災害・事故': {
      criteria: ['火災', '爆発', '自然災害', '事故'],
      threshold: '純資産の3%以上の損害',
      timeline: '発生後直ちに',
      authority: '代表取締役',
      documents: ['適時開示資料'],
      regulations: ['東証適時開示規則第2条']
    },
    '訴訟': {
      criteria: ['訴訟提起', '仲裁申立', '調停申立'],
      threshold: '純資産の15%以上の請求',
      timeline: '提起後直ちに',
      authority: '法務部門',
      documents: ['適時開示資料'],
      regulations: ['東証適時開示規則']
    },
    '行政処分': {
      criteria: ['業務停止', '業務改善命令', '課徴金'],
      threshold: '軽微基準なし',
      timeline: '処分後直ちに',
      authority: '代表取締役',
      documents: ['適時開示資料', '臨時報告書'],
      regulations: ['金商法第24条の5']
    }
  },

  // 決算関連開示
  FINANCIAL_DISCLOSURE: {
    '決算短信': {
      criteria: ['四半期決算', '通期決算'],
      timeline: '決算後45日以内（推奨30日）',
      authority: '取締役会承認',
      documents: ['決算短信', '四半期報告書'],
      regulations: ['東証決算短信作成要領']
    },
    '業績予想修正': {
      criteria: ['売上高', '営業利益', '経常利益', '純利益', '配当'],
      threshold: '10%以上の乖離',
      timeline: '判明後直ちに',
      authority: '取締役会決議',
      documents: ['適時開示資料'],
      regulations: ['東証適時開示規則']
    }
  },

  // 法定開示（財務局）
  STATUTORY_DISCLOSURE: {
    '有価証券報告書': {
      timeline: '事業年度終了後3か月以内',
      authority: '代表取締役',
      edinet: true,
      regulations: ['金商法第24条']
    },
    '四半期報告書': {
      timeline: '四半期終了後45日以内',
      authority: '代表取締役',
      edinet: true,
      regulations: ['金商法第24条の4の7']
    },
    '臨時報告書': {
      triggers: ['主要株主異動', '代表取締役異動', '監査人異動'],
      timeline: '発生後遅滞なく',
      authority: '代表取締役',
      edinet: true,
      regulations: ['金商法第24条の5']
    },
    '内部統制報告書': {
      timeline: '有価証券報告書と同時',
      authority: '代表取締役',
      edinet: true,
      regulations: ['金商法第24条の4の4']
    }
  }
};

// 株主総会関連の開示要件
const SHAREHOLDER_MEETING_DISCLOSURE = {
  '定時株主総会': {
    timeline: {
      '招集通知発送': '総会2週間前',
      '招集通知開示': '発送前にTDnet・自社Web',
      '決議通知': '総会後遅滞なく',
      '臨時報告書': '総会後遅滞なく'
    },
    documents: ['招集通知', '参考書類', '事業報告', '計算書類'],
    regulations: ['会社法第299条', '金商法第24条の5']
  },
  '臨時株主総会': {
    triggers: ['定款変更', '合併承認', '重要な事業譲渡'],
    timeline: {
      '招集通知': '総会2週間前',
      '適時開示': '招集決定後直ちに'
    },
    regulations: ['会社法第309条']
  }
};

// 取締役会決議事項の開示判定
const BOARD_RESOLUTION_DISCLOSURE = {
  '開示必須事項': [
    '決算承認',
    '配当決定',
    '自己株式取得',
    '重要な人事（代表取締役、役員）',
    '組織再編',
    '重要な業務提携'
  ],
  '開示検討事項': [
    '重要な設備投資（総資産の10%以上）',
    '重要な資産の取得・処分（総資産の10%以上）',
    '重要な訴訟の提起・和解',
    '内部統制システムの重要な変更'
  ]
};

// ================================================================================
// 開示判定関数
// ================================================================================

/**
 * タスクが開示対象かを判定
 */
function checkDisclosureRequirement(task) {
  const result = {
    requiresDisclosure: false,
    disclosureType: [],
    timeline: [],
    authorities: [],
    documents: [],
    regulations: [],
    notes: []
  };

  // キーワードベースで開示要件をチェック
  const taskLower = task.toLowerCase();
  
  // 株主総会関連
  if (taskLower.includes('株主総会')) {
    result.requiresDisclosure = true;
    result.disclosureType.push('株主総会関連開示');
    
    if (taskLower.includes('定時')) {
      result.timeline.push('招集通知: 総会2週間前');
      result.timeline.push('招集通知Web開示: 発送前');
      result.documents.push('招集通知', '事業報告', '計算書類');
      result.regulations.push('会社法第299条');
    }
    
    if (taskLower.includes('臨時')) {
      result.timeline.push('適時開示: 招集決定後直ちに');
      result.documents.push('適時開示資料', '招集通知');
      result.regulations.push('東証適時開示規則');
    }
    
    result.authorities.push('取締役会', '代表取締役');
    result.notes.push('TDnetでの開示と自社Webサイトでの公表を並行実施');
  }

  // 取締役会関連
  if (taskLower.includes('取締役会')) {
    // 決議事項から開示要否を判定
    for (const item of BOARD_RESOLUTION_DISCLOSURE['開示必須事項']) {
      if (taskLower.includes(item)) {
        result.requiresDisclosure = true;
        result.disclosureType.push('取締役会決議事項');
        result.timeline.push('決議後直ちに');
        result.authorities.push('取締役会');
        result.documents.push('適時開示資料');
        result.regulations.push('東証適時開示規則第2条');
        break;
      }
    }
  }

  // M&A・組織再編
  if (taskLower.includes('合併') || taskLower.includes('買収') || 
      taskLower.includes('m&a') || taskLower.includes('事業譲渡')) {
    result.requiresDisclosure = true;
    result.disclosureType.push('組織再編・M&A');
    result.timeline.push('基本合意時', '最終契約時', '効力発生時');
    result.authorities.push('取締役会', '株主総会（特別決議）');
    result.documents.push('適時開示資料', '臨時報告書', '公開買付届出書');
    result.regulations.push('金商法第27条の3', '会社法第783条');
    result.notes.push('財務アドバイザー・法務アドバイザーとの連携必須');
  }

  // 決算・業績関連
  if (taskLower.includes('決算') || taskLower.includes('業績')) {
    result.requiresDisclosure = true;
    result.disclosureType.push('決算開示');
    
    if (taskLower.includes('四半期')) {
      result.timeline.push('四半期終了後45日以内');
      result.documents.push('四半期決算短信', '四半期報告書');
      result.regulations.push('金商法第24条の4の7');
    } else if (taskLower.includes('通期') || taskLower.includes('年度')) {
      result.timeline.push('期末後45日以内（決算短信）', '期末後3か月以内（有価証券報告書）');
      result.documents.push('決算短信', '有価証券報告書');
      result.regulations.push('金商法第24条');
    }
    
    if (taskLower.includes('修正') || taskLower.includes('予想')) {
      result.timeline.push('判明後直ちに（業績予想修正）');
      result.notes.push('売上高・利益が10%以上乖離する場合は開示必須');
    }
    
    result.authorities.push('取締役会', '監査役会', '会計監査人');
  }

  // 資本政策
  if (taskLower.includes('増資') || taskLower.includes('自己株') || 
      taskLower.includes('配当') || taskLower.includes('資本')) {
    result.requiresDisclosure = true;
    result.disclosureType.push('資本政策');
    result.timeline.push('決議後直ちに');
    result.authorities.push('取締役会');
    result.documents.push('適時開示資料');
    
    if (taskLower.includes('増資')) {
      result.documents.push('有価証券届出書');
      result.regulations.push('金商法第4条');
      result.notes.push('財務局への届出が必要（1億円以上の場合）');
    }
  }

  // コンプライアンス・リスク事象
  if (taskLower.includes('不正') || taskLower.includes('違反') || 
      taskLower.includes('事故') || taskLower.includes('訴訟')) {
    result.requiresDisclosure = true;
    result.disclosureType.push('リスク事象');
    result.timeline.push('発生後直ちに');
    result.authorities.push('代表取締役', 'リスク管理委員会');
    result.documents.push('適時開示資料');
    result.regulations.push('東証適時開示規則（発生事実）');
    result.notes.push('投資判断に重要な影響を与える場合は軽微基準なし');
  }

  // 内部統制
  if (taskLower.includes('内部統制') || taskLower.includes('j-sox')) {
    result.requiresDisclosure = true;
    result.disclosureType.push('内部統制');
    result.timeline.push('有価証券報告書と同時提出');
    result.authorities.push('代表取締役', '監査役会');
    result.documents.push('内部統制報告書');
    result.regulations.push('金商法第24条の4の4');
    result.notes.push('重要な欠陥がある場合は適時開示も必要');
  }

  return result;
}

/**
 * ガバナンス体制に基づく承認フロー判定
 */
function determineApprovalFlow(task, disclosureResult) {
  const approvalFlow = {
    levels: [],
    timeline: [],
    checkpoints: []
  };

  // 開示が必要な場合の承認フロー
  if (disclosureResult.requiresDisclosure) {
    // レベル1: 実務担当
    approvalFlow.levels.push({
      level: 1,
      role: '実務担当者',
      action: 'ドラフト作成・事実確認',
      timeline: 'T-10営業日'
    });

    // レベル2: 部門責任者
    approvalFlow.levels.push({
      level: 2,
      role: '部門責任者',
      action: '内容精査・部門承認',
      timeline: 'T-7営業日'
    });

    // レベル3: 法務・コンプライアンス
    approvalFlow.levels.push({
      level: 3,
      role: '法務・コンプライアンス部',
      action: '法的確認・開示要否最終判定',
      timeline: 'T-5営業日'
    });

    // レベル4: 経営層
    if (disclosureResult.authorities.includes('取締役会')) {
      approvalFlow.levels.push({
        level: 4,
        role: '取締役会',
        action: '決議・承認',
        timeline: 'T-2営業日'
      });
    }

    if (disclosureResult.authorities.includes('株主総会')) {
      approvalFlow.levels.push({
        level: 5,
        role: '株主総会',
        action: '特別決議・普通決議',
        timeline: '総会当日'
      });
    }

    // チェックポイント
    approvalFlow.checkpoints.push('利益相反チェック');
    approvalFlow.checkpoints.push('インサイダー情報管理');
    approvalFlow.checkpoints.push('同時開示の原則確認');
    
    if (disclosureResult.documents.includes('有価証券届出書')) {
      approvalFlow.checkpoints.push('財務局事前相談');
    }
    
    if (disclosureResult.disclosureType.includes('M&A')) {
      approvalFlow.checkpoints.push('独立第三者委員会の意見取得');
      approvalFlow.checkpoints.push('フェアネス・オピニオン取得検討');
    }
  }

  return approvalFlow;
}

/**
 * 開示タイムラインの生成
 */
function generateDisclosureTimeline(task, targetDate) {
  const timeline = [];
  const disclosure = checkDisclosureRequirement(task);
  
  if (!disclosure.requiresDisclosure) {
    return timeline;
  }

  // 基準日からの逆算
  const baseDate = targetDate || new Date();
  
  // 各開示タイプに応じたタイムライン生成
  if (disclosure.disclosureType.includes('株主総会関連開示')) {
    timeline.push({
      date: 'T-60日',
      action: '株主総会日程の取締役会決議',
      responsible: '取締役会'
    });
    timeline.push({
      date: 'T-45日',
      action: '招集通知原案作成開始',
      responsible: '総務部・法務部'
    });
    timeline.push({
      date: 'T-30日',
      action: '監査役会・会計監査人との調整',
      responsible: '経理部・監査役会'
    });
    timeline.push({
      date: 'T-21日',
      action: '招集通知の取締役会承認',
      responsible: '取締役会'
    });
    timeline.push({
      date: 'T-16日',
      action: '招集通知Web開示（TDnet・自社サイト）',
      responsible: 'IR部'
    });
    timeline.push({
      date: 'T-14日',
      action: '招集通知発送',
      responsible: '総務部'
    });
    timeline.push({
      date: 'T-0日',
      action: '株主総会開催',
      responsible: '全役員'
    });
    timeline.push({
      date: 'T+1日',
      action: '臨時報告書提出（EDINET）',
      responsible: '法務部'
    });
  }

  if (disclosure.disclosureType.includes('決算開示')) {
    timeline.push({
      date: '決算日+5営業日',
      action: '決算数値確定',
      responsible: '経理部'
    });
    timeline.push({
      date: '決算日+20営業日',
      action: '監査法人レビュー完了',
      responsible: '監査法人・経理部'
    });
    timeline.push({
      date: '決算日+25営業日',
      action: '決算短信原案作成',
      responsible: 'IR部・経理部'
    });
    timeline.push({
      date: '決算日+28営業日',
      action: '取締役会承認',
      responsible: '取締役会'
    });
    timeline.push({
      date: '決算日+30営業日',
      action: '決算短信開示（15:00）',
      responsible: 'IR部'
    });
    timeline.push({
      date: '決算日+30営業日',
      action: '決算説明会開催',
      responsible: '経営陣・IR部'
    });
  }

  return timeline;
}

/**
 * MECEなタスク分類体系
 */
const TASK_CLASSIFICATION_MECE = {
  // レベル1: 大分類
  'ガバナンス・コンプライアンス': {
    // レベル2: 中分類
    '株主総会運営': {
      // レベル3: 小分類
      tasks: [
        '定時株主総会の準備・開催',
        '臨時株主総会の準備・開催',
        '株主総会招集通知の作成・送付',
        '議決権行使書の集計',
        '株主総会議事録の作成'
      ],
      disclosure: true,
      priority: 'HIGH'
    },
    '取締役会運営': {
      tasks: [
        '取締役会の開催・運営',
        '取締役会議事録の作成',
        '取締役会規程の管理',
        '決議事項の管理・フォロー'
      ],
      disclosure: true,
      priority: 'HIGH'
    },
    '監査対応': {
      tasks: [
        '監査役監査への対応',
        '内部監査への対応',
        '会計監査人監査への対応',
        '監査指摘事項の改善'
      ],
      disclosure: false,
      priority: 'HIGH'
    },
    'コンプライアンス管理': {
      tasks: [
        'コンプライアンス違反の防止・発見',
        '内部通報制度の運営',
        'コンプライアンス研修の実施',
        '規程・マニュアルの整備'
      ],
      disclosure: false,
      priority: 'MEDIUM'
    }
  },
  
  '情報開示・IR': {
    '適時開示': {
      tasks: [
        '決定事実の開示',
        '発生事実の開示',
        '決算情報の開示',
        '業績予想修正の開示'
      ],
      disclosure: true,
      priority: 'CRITICAL'
    },
    '法定開示': {
      tasks: [
        '有価証券報告書の作成・提出',
        '四半期報告書の作成・提出',
        '臨時報告書の作成・提出',
        '内部統制報告書の作成・提出'
      ],
      disclosure: true,
      priority: 'CRITICAL'
    },
    'IR活動': {
      tasks: [
        '決算説明会の開催',
        'アナリスト・機関投資家対応',
        '個人投資家向け説明会',
        'IR資料の作成・更新'
      ],
      disclosure: false,
      priority: 'HIGH'
    }
  },
  
  '内部統制・リスク管理': {
    '内部統制システム': {
      tasks: [
        'J-SOX対応',
        '内部統制の整備・運用',
        '内部統制の評価',
        '内部統制報告書の作成'
      ],
      disclosure: true,
      priority: 'HIGH'
    },
    'リスク管理': {
      tasks: [
        'リスクアセスメント',
        'リスク対応策の策定',
        'BCP（事業継続計画）の策定・更新',
        'インシデント対応'
      ],
      disclosure: false,
      priority: 'HIGH'
    }
  },
  
  '経営管理': {
    '経営企画': {
      tasks: [
        '中期経営計画の策定',
        '年度事業計画の策定',
        '予算策定・管理',
        'KPI管理'
      ],
      disclosure: false,
      priority: 'HIGH'
    },
    '組織管理': {
      tasks: [
        '組織変更・改編',
        '規程・規則の制定・改廃',
        '権限委譲・決裁権限の管理',
        '関係会社管理'
      ],
      disclosure: false,
      priority: 'MEDIUM'
    }
  }
};

/**
 * タスクをMECE分類に振り分け
 */
function classifyTaskMECE(task) {
  const classification = {
    level1: null,
    level2: null,
    level3: null,
    requiresDisclosure: false,
    priority: 'LOW',
    relatedTasks: []
  };

  const taskLower = task.toLowerCase();

  // 各分類を検査
  for (const [l1Key, l1Value] of Object.entries(TASK_CLASSIFICATION_MECE)) {
    for (const [l2Key, l2Value] of Object.entries(l1Value)) {
      for (const l3Task of l2Value.tasks) {
        if (taskLower.includes(l3Task.toLowerCase()) || 
            l3Task.toLowerCase().includes(taskLower)) {
          classification.level1 = l1Key;
          classification.level2 = l2Key;
          classification.level3 = l3Task;
          classification.requiresDisclosure = l2Value.disclosure;
          classification.priority = l2Value.priority;
          
          // 関連タスクを特定
          classification.relatedTasks = l2Value.tasks.filter(t => t !== l3Task);
          
          return classification;
        }
      }
    }
  }

  // マッチしない場合はキーワードベースで推定
  if (taskLower.includes('開示') || taskLower.includes('報告書')) {
    classification.level1 = '情報開示・IR';
    classification.requiresDisclosure = true;
    classification.priority = 'HIGH';
  } else if (taskLower.includes('監査') || taskLower.includes('統制')) {
    classification.level1 = '内部統制・リスク管理';
    classification.priority = 'HIGH';
  } else if (taskLower.includes('取締役') || taskLower.includes('株主')) {
    classification.level1 = 'ガバナンス・コンプライアンス';
    classification.requiresDisclosure = true;
    classification.priority = 'HIGH';
  }

  return classification;
}

/**
 * 統合的なガバナンスチェック
 */
function performComprehensiveGovernanceCheck(workSpec, flowData) {
  const governanceReport = {
    overallScore: 0,
    disclosureRequirements: [],
    complianceGaps: [],
    recommendations: [],
    timeline: [],
    riskAssessment: []
  };

  // 1. 業務仕様書からガバナンス要素を抽出
  if (workSpec) {
    const specText = JSON.stringify(workSpec).toLowerCase();
    
    // 開示要件チェック
    const disclosureCheck = checkDisclosureRequirement(specText);
    if (disclosureCheck.requiresDisclosure) {
      governanceReport.disclosureRequirements.push({
        type: disclosureCheck.disclosureType.join(', '),
        timeline: disclosureCheck.timeline,
        documents: disclosureCheck.documents,
        regulations: disclosureCheck.regulations
      });
    }
  }

  // 2. フローデータから承認プロセスを分析
  if (flowData && Array.isArray(flowData)) {
    const approvalSteps = flowData.filter(row => 
      row['作業内容'] && (
        row['作業内容'].includes('承認') ||
        row['作業内容'].includes('決裁') ||
        row['作業内容'].includes('決議')
      )
    );

    // 承認階層の適切性を評価
    const requiredApprovers = new Set();
    approvalSteps.forEach(step => {
      if (step['担当役割']) {
        requiredApprovers.add(step['担当役割']);
      }
    });

    // 必要な承認者が不足していないかチェック
    const essentialApprovers = ['取締役会', '代表取締役', '監査役'];
    essentialApprovers.forEach(approver => {
      if (!Array.from(requiredApprovers).some(r => r.includes(approver))) {
        if (governanceReport.disclosureRequirements.length > 0) {
          governanceReport.complianceGaps.push(
            `重要な承認者「${approver}」が承認フローに含まれていません`
          );
        }
      }
    });
  }

  // 3. リスク評価
  const risks = [
    {
      category: '開示遅延リスク',
      probability: governanceReport.disclosureRequirements.length > 2 ? 'HIGH' : 'MEDIUM',
      impact: 'HIGH',
      mitigation: 'IR部門との事前調整、開示チェックリストの活用'
    },
    {
      category: 'コンプライアンス違反リスク',
      probability: governanceReport.complianceGaps.length > 0 ? 'HIGH' : 'LOW',
      impact: 'CRITICAL',
      mitigation: '法務部門による事前レビュー、コンプライアンスチェックの実施'
    }
  ];
  governanceReport.riskAssessment = risks;

  // 4. 推奨事項の生成
  if (governanceReport.disclosureRequirements.length > 0) {
    governanceReport.recommendations.push(
      '東証への事前相談を検討してください（複雑な開示案件の場合）'
    );
    governanceReport.recommendations.push(
      'IR部門と法務部門の早期巻き込みを推奨します'
    );
  }

  if (governanceReport.complianceGaps.length > 0) {
    governanceReport.recommendations.push(
      '承認フローの見直しと必要な承認者の追加を検討してください'
    );
  }

  // 5. スコアリング（100点満点）
  let score = 100;
  score -= governanceReport.complianceGaps.length * 10;
  score -= governanceReport.riskAssessment.filter(r => r.probability === 'HIGH').length * 5;
  score = Math.max(0, score);
  governanceReport.overallScore = score;

  return governanceReport;
}

/**
 * エクスポート用：開示チェック結果をフォーマット
 */
function formatDisclosureCheckForSheet(task) {
  const result = checkDisclosureRequirement(task);
  
  if (!result.requiresDisclosure) {
    return '';
  }

  let formatted = '【要開示】\n';
  formatted += `種別: ${result.disclosureType.join(', ')}\n`;
  formatted += `期限: ${result.timeline.join(', ')}\n`;
  formatted += `規制: ${result.regulations.join(', ')}\n`;
  
  if (result.notes.length > 0) {
    formatted += `留意点: ${result.notes.join('; ')}`;
  }

  return formatted;
}

/**
 * バッチ処理：複数タスクの開示要件を一括チェック
 */
function batchCheckDisclosure(tasks) {
  const results = [];
  
  tasks.forEach((task, index) => {
    const checkResult = checkDisclosureRequirement(task);
    const classification = classifyTaskMECE(task);
    
    results.push({
      taskId: index + 1,
      task: task,
      requiresDisclosure: checkResult.requiresDisclosure,
      disclosureType: checkResult.disclosureType,
      priority: classification.priority,
      category: classification.level1,
      subcategory: classification.level2,
      timeline: checkResult.timeline,
      documents: checkResult.documents,
      regulations: checkResult.regulations
    });
  });

  // 優先度順にソート
  const priorityOrder = { 'CRITICAL': 0, 'HIGH': 1, 'MEDIUM': 2, 'LOW': 3 };
  results.sort((a, b) => priorityOrder[a.priority] - priorityOrder[b.priority]);

  return results;
}