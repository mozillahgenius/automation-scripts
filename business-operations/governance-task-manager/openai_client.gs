// OpenAI API設定
const OPENAI_URL_CHAT = 'https://api.openai.com/v1/chat/completions';

// JSON Schema定義
function buildWorkSpecSchema() {
  return {
    type: 'object',
    additionalProperties: false,
    properties: {
      work_spec: {
        type: 'object',
        additionalProperties: false,
        properties: {
          title: { type: 'string' },
          summary: { type: 'string' },
          scope: { type: 'string' },
          deliverables: { type: 'array', items: { type: 'string' } },
          org_structure: { type: 'array', items: { type: 'string' } },
          raci: {
            type: 'array',
            items: {
              type: 'object',
              additionalProperties: false,
              properties: {
                role: { type: 'string' },
                dept: { type: 'string' },
                R: { type: 'boolean' },
                A: { type: 'boolean' },
                C: { type: 'boolean' },
                I: { type: 'boolean' }
              },
              required: ['role', 'dept']
            }
          },
          timeline: {
            type: 'array',
            items: {
              type: 'object',
              additionalProperties: false,
              properties: {
                phase: { type: 'string' },
                duration_hint: { type: 'string' },
                milestones: { type: 'array', items: { type: 'string' } },
                dependencies: { type: 'array', items: { type: 'string' } }
              },
              required: ['phase', 'duration_hint']
            }
          },
          requirements_constraints: { type: 'array', items: { type: 'string' } },
          risks_mitigations: { type: 'array', items: { type: 'string' } },
          pro_considerations: { type: 'array', items: { type: 'string' } },
          kpi_sla: { type: 'array', items: { type: 'string' } },
          approvals: { type: 'array', items: { type: 'string' } },
          security_privacy_controls: { type: 'array', items: { type: 'string' } },
          legal_regulations: {
            type: 'array',
            items: {
              type: 'object',
              additionalProperties: false,
              properties: {
                name: { type: 'string' },
                scope: { type: 'string' },
                note: { type: 'string' }
              },
              required: ['name']
            }
          },
          references: { type: 'array', items: { type: 'string' } },
          assumptions: { type: 'array', items: { type: 'string' } }
        },
        required: [
          'title', 'summary', 'scope', 'deliverables', 'org_structure',
          'raci', 'timeline', 'requirements_constraints', 'risks_mitigations',
          'pro_considerations', 'kpi_sla', 'approvals', 'security_privacy_controls',
          'legal_regulations', 'references', 'assumptions'
        ]
      },
      flow_rows: {
        type: 'array',
        items: {
          type: 'object',
          additionalProperties: false,
          properties: {
            '工程': { type: 'string' },
            '実施タイミング': { type: 'string' },
            '部署': { type: 'string' },
            '担当役割': { type: 'string' },
            '作業内容': { type: 'string' },
            '条件分岐': { type: 'string' },
            '利用ツール': { type: 'string' },
            'URLリンク': { type: 'string' },
            '備考': { type: 'string' }
          },
          required: ['工程', '実施タイミング', '部署', '担当役割', '作業内容', '条件分岐', '利用ツール', 'URLリンク', '備考']
        }
      }
    },
    required: ['work_spec', 'flow_rows']
  };
}

// システムプロンプト構築
function buildSystemPrompt() {
  return `あなたは上場企業レベルのプロジェクトマネジメント、法務、内部統制の実務に精通したアシスタントです。

重要: 与えられた情報が少ない場合でも、あなたの専門知識を活用して、実務的で具体的な提案を積極的に行ってください。

業務内容の記述から、以下の観点で業務記述書とフロー表を作成してください：

1. 上場企業のプロ水準の品質
   - 内部統制（J-SOX）への配慮
   - 説明責任と透明性の確保
   - リスク管理と対策の明確化

2. 法令・規制への対応
   - 該当する法令・ガイドラインを名称レベルで列挙
   - 業界特有の規制やルールも考慮
   - 必ず「最終的には法務・専門家の確認が必要」と明記

3. 詳細な体制とスケジュール
   - RACI（推奨）による役割分担の明確化
   - フェーズ分割と各フェーズの所要期間目安
   - マイルストーンと依存関係の明示

4. セキュリティと個人情報保護
   - データ分類とアクセス権限
   - 個人情報保護法への準拠
   - セキュリティ管理策の具体化

5. 成果物と受入基準
   - 各成果物の明確な定義
   - 完了基準（DoD）の設定
   - KPI/SLAの設定

すべての出力は日本語で作成し、指定されたJSON Schemaに完全準拠してください。
法的助言の代替ではないことを必ず明記してください。`;
}

// ユーザープロンプト構築
function buildUserPrompt(mailBody, orgProfileJson) {
  const orgProfile = orgProfileJson ? JSON.parse(orgProfileJson) : {};
  
  return `以下の業務内容から、業務記述書とフロー表を作成してください。

【業務内容】
${mailBody}

【組織プロフィール】
- 上場区分: ${orgProfile.listing || '未設定'}
- 業種: ${orgProfile.industry || '未設定'}
- 対象地域: ${(orgProfile.jurisdictions || ['JP']).join(', ')}
- 社内基準: ${(orgProfile.policies || []).join(', ')}

【要求事項】
1. 上場企業として必要な観点をすべて網羅
2. 法令・規制は具体的な名称を記載（最終確認は専門家が必要な旨も明記）
3. フロー表は実行可能な詳細レベルで作成
4. リスクと対策を具体的に記載
5. セキュリティ・個人情報保護・内部統制の観点を含める

【重要な指示】
- 情報が不足している場合は、業界標準やベストプラクティスに基づいて、具体的な提案を行ってください
- 「上場準備」のような抽象的な依頼でも、IPO準備の標準的なプロセスを想定して詳細な計画を作成してください
- すべてのフィールドに必ず値を設定してください（空配列の場合も最低1つの要素を含める）
- scope: 業務の範囲を明確に記載
- deliverables: 成果物リストを必ず含める（最低1つ）
- org_structure: 組織体制を必ず含める（最低1つ）
- 条件分岐がない場合は「なし」、ツールが不要な場合は「手動作業」など、適切な値を設定してください`;
}

// OpenAI API呼び出し
function callOpenAI(mailBody, orgProfileJson) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
  if (!apiKey) {
    throw new Error('OpenAI APIキーが設定されていません。スクリプトプロパティに OPENAI_API_KEY を設定してください。');
  }
  
  const schema = buildWorkSpecSchema();
  const messages = [
    { role: 'system', content: buildSystemPrompt() },
    { role: 'user', content: buildUserPrompt(mailBody, orgProfileJson) }
  ];
  
  const payload = {
    model: getConfig('OPENAI_MODEL') || 'gpt-4o',
    messages: messages,
    response_format: {
      type: 'json_schema',
      json_schema: {
        name: 'WorkSpecSchema',
        schema: schema,
        strict: true  // strictモードを有効化してデータ品質を向上
      }
    },
    temperature: 0.1,  // より一貫した出力のために温度を下げる
    max_tokens: 6000,  // より詳細な出力のためにトークン数を増加
    seed: 42  // 再現性のためのシード値
  };
  
  logActivity('OPENAI_CALL', `Calling OpenAI with model: ${payload.model}`);
  
  const response = retryWithBackoff(() => {
    const res = UrlFetchApp.fetch(OPENAI_URL_CHAT, {
      method: 'post',
      headers: {
        'Authorization': `Bearer ${apiKey}`,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    
    const status = res.getResponseCode();
    if (status >= 300) {
      const errorBody = res.getContentText();
      logActivity('OPENAI_ERROR', `Status: ${status}, Body: ${errorBody}`);
      throw new Error(`OpenAI API error ${status}: ${errorBody}`);
    }
    
    return res;
  });
  
  const responseData = JSON.parse(response.getContentText());
  const content = responseData.choices[0].message.content;
  
  logActivity('OPENAI_SUCCESS', 'Successfully received response from OpenAI');
  
  return JSON.parse(content);
}

// JSON検証
function validateOpenAIResponse(data) {
  console.log('バリデーション開始:', typeof data, data ? Object.keys(data) : 'null');
  
  if (!data || typeof data !== 'object') {
    console.error('レスポンス形式エラー:', data);
    throw new Error('Invalid response format');
  }
  
  if (!data.work_spec || typeof data.work_spec !== 'object') {
    console.error('work_spec エラー:', data.work_spec);
    throw new Error('Missing or invalid work_spec');
  }
  
  if (!data.flow_rows || !Array.isArray(data.flow_rows)) {
    console.error('flow_rows エラー:', data.flow_rows);
    throw new Error(`Missing or invalid flow_rows. Expected array, got: ${typeof data.flow_rows}`);
  }
  
  console.log(`flow_rows 件数: ${data.flow_rows.length}`);
  
  const ws = data.work_spec;
  const required = ['title', 'summary', 'timeline', 'pro_considerations', 'legal_regulations'];
  console.log('work_spec フィールド確認:', Object.keys(ws));
  
  for (const field of required) {
    if (!ws[field] || (typeof ws[field] === 'string' && ws[field].trim() === '')) {
      console.warn(`work_spec の ${field} が空です。デフォルト値を設定します。`);
      // デフォルト値を設定
      switch (field) {
        case 'title':
          ws[field] = '業務記述書';
          break;
        case 'summary':
          ws[field] = '業務概要を記載してください';
          break;
        case 'timeline':
          ws[field] = 'スケジュールを調整中';
          break;
        case 'pro_considerations':
          ws[field] = '専門的な考慮事項を確認中';
          break;
        case 'legal_regulations':
          ws[field] = '法的規制を確認中';
          break;
        default:
          ws[field] = '要確認';
      }
    }
  }
  
  for (let i = 0; i < data.flow_rows.length; i++) {
    const row = data.flow_rows[i];
    const requiredFlow = ['工程', '実施タイミング', '部署', '担当役割', '作業内容'];
    
    for (const field of requiredFlow) {
      if (!row[field] || row[field].trim() === '') {
        // デフォルト値を設定（combined_gas_code.gsと統一）
        switch (field) {
          case '工程':
            row[field] = `フェーズ${i + 1}`;
            break;
          case '実施タイミング':
            row[field] = `第${i + 1}期`;
            break;
          case '部署':
            row[field] = '経営企画部';
            break;
          case '担当役割':
            row[field] = '担当者';
            break;
          case '作業内容':
            row[field] = 'タスク実施';
            break;
          default:
            row[field] = '未設定';
        }
        console.log(`デフォルト値を設定: row ${i + 1}, field: ${field}, value: ${row[field]}`);
      }
    }
    
    // オプションフィールドのデフォルト値（combined_gas_code.gsと統一）
    if (!row['条件分岐']) row['条件分岐'] = 'なし';
    if (!row['利用ツール']) row['利用ツール'] = '';
    if (!row['URLリンク']) row['URLリンク'] = '';
    if (!row['備考']) row['備考'] = '';
    
    // すべてのフィールドを文字列として正規化
    for (const field of ['工程', '実施タイミング', '部署', '担当役割', '作業内容', '条件分岐', '利用ツール', 'URLリンク', '備考']) {
      if (row[field] !== undefined && row[field] !== null) {
        row[field] = String(row[field]).trim();
        
        // より積極的な末尾数字の除去
        const originalValue = row[field];
        
        // パターン1: 末尾に単独の数字が付いている場合
        if (/\d+$/.test(originalValue) && originalValue.length > 1) {
          const withoutNumber = originalValue.replace(/\d+$/, '').trim();
          if (withoutNumber.length > 0) {
            console.log(`末尾の数字を除去 (パターン1): "${originalValue}" -> "${withoutNumber}"`);
            row[field] = withoutNumber;
          }
        }
        
        // パターン2: 特定の文字列パターンの後に数字が付いている場合
        if (/特になし\d+$/.test(originalValue)) {
          const cleaned = originalValue.replace(/\d+$/, '').trim();
          console.log(`特になし後の数字を除去: "${originalValue}" -> "${cleaned}"`);
          row[field] = cleaned;
        }
        
        // パターン3: その他の一般的な末尾数字パターン
        if (/[^\d]\d+$/.test(originalValue) && originalValue.length > 2) {
          const cleaned = originalValue.replace(/\d+$/, '').trim();
          if (cleaned.length > 0 && cleaned !== originalValue) {
            console.log(`一般的な末尾数字を除去: "${originalValue}" -> "${cleaned}"`);
            row[field] = cleaned;
          }
        }
      }
    }
  }
  
  console.log('バリデーション完了: すべてのフィールドが適切に設定されました');
  return true;
}