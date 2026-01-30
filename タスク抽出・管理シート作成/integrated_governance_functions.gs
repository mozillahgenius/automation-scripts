/**
 * 統合ガバナンス機能
 * 既存コードに組み込むための関数群
 */

// ================================================================================
// グローバル設定：ガバナンスチェック機能を有効化
// ================================================================================

const GOVERNANCE_CONFIG = {
  enableDisclosureCheck: true,
  enableAdvisorConsultation: true,
  enableMECEClassification: true,
  autoGenerateTimeline: true,
  strictComplianceMode: true
};

// ================================================================================
// メイン統合関数：タスク処理時に自動実行
// ================================================================================

/**
 * タスク処理時の統合ガバナンスチェック
 * @param {Object} workSpec - 業務仕様書データ
 * @param {Array} flowRows - フローデータ
 * @return {Object} ガバナンスチェック結果
 */
function performIntegratedGovernanceCheck(workSpec, flowRows) {
  console.log('===== 統合ガバナンスチェック開始 =====');
  
  const governanceResults = {
    timestamp: new Date(),
    summary: {
      totalTasks: 0,
      disclosureRequired: 0,
      advisorConsultationRequired: 0,
      highPriorityItems: 0,
      complianceGaps: []
    },
    disclosureRequirements: [],
    advisorConsultations: [],
    timeline: [],
    recommendations: [],
    enrichedFlowRows: []
  };

  // 1. フローデータの各タスクをチェック
  if (flowRows && Array.isArray(flowRows)) {
    governanceResults.summary.totalTasks = flowRows.length;
    
    flowRows.forEach((row, index) => {
      const taskDescription = `${row['工程'] || ''} ${row['作業内容'] || ''}`;
      const enrichedRow = {...row};
      
      // 開示要件チェック
      if (GOVERNANCE_CONFIG.enableDisclosureCheck) {
        const disclosureCheck = checkDisclosureRequirement(taskDescription);
        if (disclosureCheck.requiresDisclosure) {
          governanceResults.summary.disclosureRequired++;
          governanceResults.disclosureRequirements.push({
            taskId: index + 1,
            task: taskDescription,
            ...disclosureCheck
          });
          
          // フローデータに開示情報を追加
          enrichedRow['開示要件'] = disclosureCheck.disclosureType.join(', ');
          enrichedRow['開示期限'] = disclosureCheck.timeline.join(', ');
        }
      }
      
      // 外部専門家相談チェック
      if (GOVERNANCE_CONFIG.enableAdvisorConsultation) {
        const advisorCheck = determineRequiredAdvisors(taskDescription);
        if (advisorCheck.length > 0) {
          governanceResults.summary.advisorConsultationRequired++;
          
          const consultation = {
            taskId: index + 1,
            task: taskDescription,
            advisors: advisorCheck,
            checklist: generateConsultationChecklist(taskDescription, advisorCheck)
          };
          governanceResults.advisorConsultations.push(consultation);
          
          // フローデータに専門家相談情報を追加
          enrichedRow['要専門家相談'] = advisorCheck.map(a => a.type).join(', ');
          enrichedRow['相談タイミング'] = advisorCheck[0].timing;
        }
      }
      
      // MECE分類
      if (GOVERNANCE_CONFIG.enableMECEClassification) {
        const classification = classifyTaskMECE(taskDescription);
        enrichedRow['MECE分類'] = classification.level1;
        enrichedRow['優先度'] = classification.priority;
        
        if (classification.priority === 'CRITICAL' || classification.priority === 'HIGH') {
          governanceResults.summary.highPriorityItems++;
        }
      }
      
      governanceResults.enrichedFlowRows.push(enrichedRow);
    });
  }

  // 2. タイムライン生成
  if (GOVERNANCE_CONFIG.autoGenerateTimeline) {
    governanceResults.timeline = generateMasterTimeline(
      governanceResults.disclosureRequirements,
      governanceResults.advisorConsultations
    );
  }

  // 3. 総合的な推奨事項生成
  governanceResults.recommendations = generateGovernanceRecommendations(governanceResults);

  // 4. コンプライアンスギャップ分析
  governanceResults.summary.complianceGaps = identifyComplianceGaps(workSpec, flowRows);

  console.log('===== 統合ガバナンスチェック完了 =====');
  console.log(`開示必要: ${governanceResults.summary.disclosureRequired}件`);
  console.log(`専門家相談必要: ${governanceResults.summary.advisorConsultationRequired}件`);
  console.log(`高優先度: ${governanceResults.summary.highPriorityItems}件`);

  return governanceResults;
}

// ================================================================================
// タイムライン生成
// ================================================================================

/**
 * マスタータイムライン生成
 */
function generateMasterTimeline(disclosureRequirements, advisorConsultations) {
  const timeline = [];
  const today = new Date();
  
  // 専門家相談フェーズ（最初に実施）
  advisorConsultations.forEach(consultation => {
    consultation.checklist.consultationSteps.forEach(step => {
      timeline.push({
        date: calculateDate(today, step.timeline),
        phase: '専門家相談',
        action: step.phase,
        responsible: step.responsible,
        priority: consultation.advisors[0].priority || 'HIGH'
      });
    });
  });

  // 開示準備フェーズ
  disclosureRequirements.forEach(disclosure => {
    disclosure.timeline.forEach(timing => {
      timeline.push({
        date: calculateDate(today, timing),
        phase: '開示対応',
        action: `${disclosure.disclosureType.join(', ')}の準備`,
        responsible: 'IR部・法務部',
        priority: 'CRITICAL',
        documents: disclosure.documents
      });
    });
  });

  // 日付順にソート
  timeline.sort((a, b) => a.date - b.date);

  return timeline;
}

/**
 * 日付計算ヘルパー
 */
function calculateDate(baseDate, timeExpression) {
  // T-X日形式のパース
  const match = timeExpression.match(/T([+-])(\d+)/);
  if (match) {
    const sign = match[1];
    const days = parseInt(match[2]);
    const date = new Date(baseDate);
    date.setDate(date.getDate() + (sign === '-' ? -days : days));
    return date;
  }
  return baseDate;
}

// ================================================================================
// コンプライアンスギャップ分析
// ================================================================================

/**
 * コンプライアンスギャップの特定
 */
function identifyComplianceGaps(workSpec, flowRows) {
  const gaps = [];
  
  // 必須プロセスの確認
  const requiredProcesses = [
    { name: '内部監査', regulation: 'J-SOX' },
    { name: 'リスク評価', regulation: 'COSO' },
    { name: '承認プロセス', regulation: '内部統制' }
  ];

  requiredProcesses.forEach(process => {
    const exists = flowRows && flowRows.some(row => 
      (row['作業内容'] || '').includes(process.name)
    );
    
    if (!exists) {
      gaps.push({
        type: 'MISSING_PROCESS',
        process: process.name,
        regulation: process.regulation,
        severity: 'HIGH',
        recommendation: `${process.name}を業務フローに追加してください`
      });
    }
  });

  // 外部専門家相談の抜け漏れチェック
  const criticalTasks = flowRows ? flowRows.filter(row => {
    const task = `${row['工程'] || ''} ${row['作業内容'] || ''}`;
    const advisors = determineRequiredAdvisors(task);
    return advisors.some(a => a.priority === 'MANDATORY') && !row['要専門家相談'];
  }) : [];

  criticalTasks.forEach(task => {
    gaps.push({
      type: 'MISSING_ADVISOR_CONSULTATION',
      task: `${task['工程']} - ${task['作業内容']}`,
      severity: 'CRITICAL',
      recommendation: '外部専門家への事前相談が必須です'
    });
  });

  return gaps;
}

// ================================================================================
// 推奨事項生成
// ================================================================================

/**
 * ガバナンス推奨事項の生成
 */
function generateGovernanceRecommendations(governanceResults) {
  const recommendations = [];
  
  // 開示関連
  if (governanceResults.summary.disclosureRequired > 0) {
    recommendations.push({
      category: '開示対応',
      priority: 'CRITICAL',
      items: [
        '東証への事前相談を実施してください',
        'IR部門と法務部門の早期巻き込みが必要です',
        '開示スケジュールの詳細化と関係者への共有を行ってください',
        'インサイダー情報管理体制の確認が必要です'
      ]
    });
  }

  // 専門家相談関連
  if (governanceResults.summary.advisorConsultationRequired > 0) {
    recommendations.push({
      category: '外部専門家活用',
      priority: 'HIGH',
      items: [
        '複数の専門家から見積もりを取得し、比較検討してください',
        'NDA（秘密保持契約）の締結を最初に行ってください',
        '相談事項を明確化し、必要資料を事前に準備してください',
        '専門家の意見を統合し、リスク評価を実施してください'
      ]
    });
  }

  // コンプライアンスギャップ
  if (governanceResults.summary.complianceGaps.length > 0) {
    const criticalGaps = governanceResults.summary.complianceGaps.filter(
      gap => gap.severity === 'CRITICAL' || gap.severity === 'HIGH'
    );
    
    if (criticalGaps.length > 0) {
      recommendations.push({
        category: 'コンプライアンス改善',
        priority: 'CRITICAL',
        items: criticalGaps.map(gap => gap.recommendation)
      });
    }
  }

  // 一般的な推奨事項
  recommendations.push({
    category: '継続的改善',
    priority: 'MEDIUM',
    items: [
      '定期的なガバナンス体制の見直しを実施してください',
      '内部監査による業務プロセスの検証を行ってください',
      '従業員向けコンプライアンス研修を定期実施してください',
      'リスク管理委員会での定期的なモニタリングを行ってください'
    ]
  });

  return recommendations;
}

// ================================================================================
// レポート生成
// ================================================================================

/**
 * ガバナンスチェックレポートをスプレッドシートに出力
 */
function outputGovernanceReport(spreadsheet, governanceResults) {
  // レポートシートを作成または取得
  let reportSheet = spreadsheet.getSheetByName('ガバナンスレポート');
  if (!reportSheet) {
    reportSheet = spreadsheet.insertSheet('ガバナンスレポート');
  } else {
    reportSheet.clear();
  }

  let row = 1;

  // タイトル
  reportSheet.getRange(row, 1, 1, 8).merge();
  reportSheet.getRange(row, 1).setValue('ガバナンス・コンプライアンスチェックレポート');
  reportSheet.getRange(row, 1).setFontSize(16).setFontWeight('bold');
  reportSheet.getRange(row, 1).setBackground('#1a73e8').setFontColor('#ffffff');
  row += 2;

  // サマリー
  reportSheet.getRange(row, 1).setValue('【サマリー】');
  reportSheet.getRange(row, 1).setFontWeight('bold').setBackground('#e8f0fe');
  row++;
  
  const summaryData = [
    ['チェック日時', governanceResults.timestamp],
    ['総タスク数', governanceResults.summary.totalTasks],
    ['開示必要件数', governanceResults.summary.disclosureRequired],
    ['専門家相談必要件数', governanceResults.summary.advisorConsultationRequired],
    ['高優先度項目数', governanceResults.summary.highPriorityItems],
    ['コンプライアンスギャップ', governanceResults.summary.complianceGaps.length]
  ];
  
  reportSheet.getRange(row, 1, summaryData.length, 2).setValues(summaryData);
  reportSheet.getRange(row, 1, summaryData.length, 1).setFontWeight('bold');
  row += summaryData.length + 2;

  // 開示要件
  if (governanceResults.disclosureRequirements.length > 0) {
    reportSheet.getRange(row, 1).setValue('【開示要件】');
    reportSheet.getRange(row, 1).setFontWeight('bold').setBackground('#fce8b2');
    row++;
    
    const disclosureHeaders = ['No.', 'タスク', '開示種別', '期限', '必要書類', '関連法規'];
    reportSheet.getRange(row, 1, 1, disclosureHeaders.length).setValues([disclosureHeaders]);
    reportSheet.getRange(row, 1, 1, disclosureHeaders.length).setFontWeight('bold');
    row++;
    
    governanceResults.disclosureRequirements.forEach((req, index) => {
      const rowData = [
        index + 1,
        req.task,
        req.disclosureType.join(', '),
        req.timeline.join(', '),
        req.documents.join(', '),
        req.regulations.join(', ')
      ];
      reportSheet.getRange(row, 1, 1, rowData.length).setValues([rowData]);
      row++;
    });
    row++;
  }

  // 外部専門家相談
  if (governanceResults.advisorConsultations.length > 0) {
    reportSheet.getRange(row, 1).setValue('【外部専門家相談】');
    reportSheet.getRange(row, 1).setFontWeight('bold').setBackground('#d9ead3');
    row++;
    
    governanceResults.advisorConsultations.forEach((consultation, index) => {
      reportSheet.getRange(row, 1).setValue(`${index + 1}. ${consultation.task}`);
      reportSheet.getRange(row, 1).setFontWeight('bold');
      row++;
      
      consultation.advisors.forEach(advisor => {
        reportSheet.getRange(row, 2).setValue(`・${advisor.type}: ${advisor.reason}`);
        row++;
      });
      row++;
    });
  }

  // 推奨事項
  reportSheet.getRange(row, 1).setValue('【推奨事項】');
  reportSheet.getRange(row, 1).setFontWeight('bold').setBackground('#f4cccc');
  row++;
  
  governanceResults.recommendations.forEach(rec => {
    reportSheet.getRange(row, 1).setValue(`[${rec.priority}] ${rec.category}`);
    reportSheet.getRange(row, 1).setFontWeight('bold');
    row++;
    
    rec.items.forEach(item => {
      reportSheet.getRange(row, 2).setValue(`・${item}`);
      row++;
    });
    row++;
  });

  // 書式調整
  reportSheet.autoResizeColumns(1, 8);
  
  console.log('ガバナンスレポート出力完了');
}

// ================================================================================
// エクスポート用関数（既存コードから呼び出し）
// ================================================================================

/**
 * フロー処理時の自動ガバナンスチェック
 */
function enhanceFlowWithGovernance(flowSheet, flowRows) {
  // ガバナンスチェック実行
  const governanceResults = performIntegratedGovernanceCheck(null, flowRows);
  
  // 強化されたフローデータをシートに反映
  if (governanceResults.enrichedFlowRows.length > 0) {
    // 新しい列を追加
    const additionalHeaders = ['開示要件', '開示期限', '要専門家相談', '相談タイミング', 'MECE分類', '優先度'];
    const lastCol = flowSheet.getLastColumn();
    
    // ヘッダー追加
    flowSheet.getRange(1, lastCol + 1, 1, additionalHeaders.length).setValues([additionalHeaders]);
    flowSheet.getRange(1, lastCol + 1, 1, additionalHeaders.length).setFontWeight('bold').setBackground('#ffe599');
    
    // データ追加
    const enrichedData = governanceResults.enrichedFlowRows.map(row => [
      row['開示要件'] || '',
      row['開示期限'] || '',
      row['要専門家相談'] || '',
      row['相談タイミング'] || '',
      row['MECE分類'] || '',
      row['優先度'] || ''
    ]);
    
    flowSheet.getRange(2, lastCol + 1, enrichedData.length, additionalHeaders.length).setValues(enrichedData);
    
    // 優先度による色分け
    for (let i = 0; i < enrichedData.length; i++) {
      const priority = enrichedData[i][5];
      let bgColor = '#ffffff';
      
      switch(priority) {
        case 'CRITICAL': bgColor = '#f4cccc'; break;
        case 'HIGH': bgColor = '#fce5cd'; break;
        case 'MEDIUM': bgColor = '#fff2cc'; break;
        case 'LOW': bgColor = '#d9ead3'; break;
      }
      
      if (priority) {
        flowSheet.getRange(i + 2, lastCol + 6).setBackground(bgColor);
      }
    }
  }
  
  return governanceResults;
}