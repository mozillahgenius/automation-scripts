// ==========================================
// JSON Schema定義とプロンプトテンプレート
// ==========================================

// ========= JSON Schemas =========

// メッセージ要約用スキーマ
const SUMMARY_SCHEMA = {
  name: "message_summary",
  strict: true,
  schema: {
    type: "object",
    required: ["summary", "decisions", "action_items", "people", "dates"],
    additionalProperties: false,
    properties: {
      summary: {
        type: "string",
        description: "メッセージの要約（100文字以内）"
      },
      decisions: {
        type: "array",
        description: "決定事項のリスト",
        items: {
          type: "string"
        }
      },
      action_items: {
        type: "array",
        description: "アクションアイテムのリスト",
        items: {
          type: "object",
          required: ["owner", "task", "due"],
          additionalProperties: false,
          properties: {
            owner: {
              type: "string",
              description: "担当者名"
            },
            task: {
              type: "string",
              description: "タスク内容"
            },
            due: {
              type: "string",
              description: "期限"
            }
          }
        }
      },
      people: {
        type: "array",
        description: "関係者のリスト",
        items: {
          type: "string"
        }
      },
      dates: {
        type: "array",
        description: "言及された日付のリスト",
        items: {
          type: "string"
        }
      }
    }
  }
};

// カテゴリ分類用スキーマ
const CLASSIFICATION_SCHEMA = {
  name: "category_classification",
  strict: true,
  schema: {
    type: "object",
    required: ["classifications"],
    additionalProperties: false,
    properties: {
      classifications: {
        type: "array",
        description: "各カテゴリの分類結果",
        items: {
          type: "object",
          required: ["category", "score", "rationale", "key_quotes"],
          additionalProperties: false,
          properties: {
            category: {
              type: "string",
              description: "カテゴリ名"
            },
            score: {
              type: "number",
              description: "該当度スコア（0.0-1.0）",
              minimum: 0,
              maximum: 1
            },
            rationale: {
              type: "string",
              description: "判定理由"
            },
            key_quotes: {
              type: "array",
              description: "関連する引用",
              items: {
                type: "string"
              }
            }
          }
        }
      }
    }
  }
};

// チェックリスト提案用スキーマ
const CHECKLIST_SCHEMA = {
  name: "checklist_proposal",
  strict: true,
  schema: {
    type: "object",
    required: ["category", "checklist_items"],
    additionalProperties: false,
    properties: {
      category: {
        type: "string",
        description: "対象カテゴリ"
      },
      checklist_items: {
        type: "array",
        description: "チェックリスト項目",
        items: {
          type: "object",
          required: ["item", "why", "example", "priority"],
          additionalProperties: false,
          properties: {
            item: {
              type: "string",
              description: "チェック項目"
            },
            why: {
              type: "string",
              description: "必要な理由"
            },
            example: {
              type: "string",
              description: "記載例"
            },
            priority: {
              type: "string",
              enum: ["必須", "推奨", "任意"],
              description: "優先度"
            }
          }
        }
      }
    }
  }
};

// 議事録ドラフト用スキーマ
const MINUTES_DRAFT_SCHEMA = {
  name: "minutes_draft",
  strict: true,
  schema: {
    type: "object",
    required: ["meeting_type", "meeting_date", "meeting_title", "attendees", "agenda_items", "next_actions"],
    additionalProperties: false,
    properties: {
      meeting_type: {
        type: "string",
        description: "会議種別"
      },
      meeting_date: {
        type: "string",
        description: "開催日時"
      },
      meeting_title: {
        type: "string",
        description: "議題名"
      },
      attendees: {
        type: "array",
        description: "出席者リスト",
        items: {
          type: "object",
          required: ["name", "role"],
          additionalProperties: false,
          properties: {
            name: {
              type: "string"
            },
            role: {
              type: "string"
            }
          }
        }
      },
      agenda_items: {
        type: "array",
        description: "議案リスト",
        items: {
          type: "object",
          required: ["number", "title", "proposer", "summary", "discussion", "resolution", "vote_result"],
          additionalProperties: false,
          properties: {
            number: {
              type: "string",
              description: "議案番号"
            },
            title: {
              type: "string",
              description: "議案名"
            },
            proposer: {
              type: "string",
              description: "提案者"
            },
            summary: {
              type: "string",
              description: "議案概要"
            },
            discussion: {
              type: "string",
              description: "審議内容"
            },
            resolution: {
              type: "string",
              description: "決議内容"
            },
            vote_result: {
              type: "string",
              description: "採決結果"
            }
          }
        }
      },
      next_actions: {
        type: "array",
        description: "今後の対応事項",
        items: {
          type: "object",
          required: ["action", "responsible", "deadline"],
          additionalProperties: false,
          properties: {
            action: {
              type: "string"
            },
            responsible: {
              type: "string"
            },
            deadline: {
              type: "string"
            }
          }
        }
      }
    }
  }
};

// ========= プロンプトテンプレート =========

const PROMPTS = {
  // システムプロンプト
  SYSTEM: {
    LEGAL_ASSISTANT: "あなたは日本の会社法および金融商品取引法に精通した法務アシスタントです。正確で実務的な支援を提供してください。",
    GOVERNANCE_EXPERT: "あなたは企業ガバナンスの専門家です。取締役会、監査等委員会、株主総会の運営に関する実務知識を活用してください。",
    DISCLOSURE_SPECIALIST: "あなたは適時開示の専門家です。東京証券取引所の規則に基づいた開示判断を支援してください。"
  },
  
  // 要約プロンプト
  SUMMARIZE: `以下のSlackメッセージを分析し、ビジネス上の重要な情報を抽出してください。

メッセージ:
{{MESSAGE_TEXT}}

以下の観点で要約してください：
1. 議論の要点（何について話しているか）
2. 決定事項（合意や決定された内容）
3. アクションアイテム（誰が何をいつまでに行うか）
4. 関係者（議論に関わっている人物）
5. 重要な日付（締切、予定日など）

JSON形式で構造化して出力してください。`,

  // 分類プロンプト
  CLASSIFY: `以下の要約内容を、企業ガバナンスの観点から分類してください。

要約内容:
{{SUMMARY_JSON}}

カテゴリ定義:
{{CATEGORIES_DEFINITION}}

各カテゴリについて：
1. 該当度を0.0-1.0のスコアで評価
2. そのスコアをつけた理由を説明
3. 判断の根拠となる要約内容からの引用を提示

特に以下の観点を重視してください：
- 法的要件の有無
- 金額的重要性
- 組織への影響度
- 開示義務の可能性

JSON形式で出力してください。`,

  // チェックリスト生成プロンプト
  GENERATE_CHECKLIST: `以下の案件について、必要な確認事項のチェックリストを作成してください。

カテゴリ: {{CATEGORY}}
要約: {{SUMMARY}}

以下の観点でチェックリストを作成：
1. 法令上の要件
2. 社内規程上の要件
3. 実務上の留意点
4. 必要書類
5. 関係者への連絡事項

各項目について：
- なぜ必要か
- 具体的な記載例
- 優先度（必須/推奨/任意）

を含めてJSON形式で出力してください。`,

  // 議事録ドラフト生成プロンプト
  GENERATE_MINUTES: `以下の情報を基に、{{MEETING_TYPE}}の議事録案を作成してください。

要約情報:
{{SUMMARY}}

チェックリスト:
{{CHECKLIST}}

議事録に含める内容：
1. 基本情報（日時、場所、出席者）
2. 議案内容
3. 審議経過
4. 決議結果
5. 今後の対応事項

以下の形式要件を満たしてください：
- 法的要件を充足する記載
- 客観的で正確な表現
- 時系列に沿った記述
- 決議の成立要件の明記

JSON形式で構造化して出力してください。`,

  // 開示文書ドラフト生成プロンプト
  GENERATE_DISCLOSURE: `以下の情報を基に、適時開示文書の案を作成してください。

決議内容:
{{RESOLUTION}}

開示に含める内容：
1. 開示事項の概要
2. 決定/発生の経緯
3. 今後の見通し
4. 業績への影響

東証の開示様式に準拠し、以下に留意してください：
- 投資判断に必要な情報の網羅
- 誤解を招かない表現
- 数値の正確性
- 将来見通しに関する注意事項

適切な開示文案を作成してください。`
};

// ========= プロンプト生成関数 =========

function buildSummarizePrompt(messageText) {
  return PROMPTS.SUMMARIZE.replace('{{MESSAGE_TEXT}}', messageText);
}

function buildClassifyPrompt(summaryJson, categoriesDefinition) {
  return PROMPTS.CLASSIFY
    .replace('{{SUMMARY_JSON}}', JSON.stringify(summaryJson, null, 2))
    .replace('{{CATEGORIES_DEFINITION}}', categoriesDefinition);
}

function buildChecklistPrompt(category, summary) {
  return PROMPTS.GENERATE_CHECKLIST
    .replace('{{CATEGORY}}', category)
    .replace('{{SUMMARY}}', JSON.stringify(summary, null, 2));
}

function buildMinutesPrompt(meetingType, summary, checklist) {
  return PROMPTS.GENERATE_MINUTES
    .replace('{{MEETING_TYPE}}', meetingType)
    .replace('{{SUMMARY}}', JSON.stringify(summary, null, 2))
    .replace('{{CHECKLIST}}', JSON.stringify(checklist, null, 2));
}

function buildDisclosurePrompt(resolution) {
  return PROMPTS.GENERATE_DISCLOSURE
    .replace('{{RESOLUTION}}', JSON.stringify(resolution, null, 2));
}

// ========= バリデーション関数 =========

function validateSummary(data) {
  try {
    const parsed = typeof data === 'string' ? JSON.parse(data) : data;
    return parsed.summary && 
           Array.isArray(parsed.decisions) && 
           Array.isArray(parsed.action_items) &&
           Array.isArray(parsed.people) &&
           Array.isArray(parsed.dates);
  } catch (e) {
    return false;
  }
}

function validateClassification(data) {
  try {
    const parsed = typeof data === 'string' ? JSON.parse(data) : data;
    if (!parsed.classifications || !Array.isArray(parsed.classifications)) {
      return false;
    }
    return parsed.classifications.every(c => 
      c.category && 
      typeof c.score === 'number' && 
      c.score >= 0 && 
      c.score <= 1 &&
      c.rationale &&
      Array.isArray(c.key_quotes)
    );
  } catch (e) {
    return false;
  }
}

function validateChecklist(data) {
  try {
    const parsed = typeof data === 'string' ? JSON.parse(data) : data;
    return parsed.category && 
           Array.isArray(parsed.checklist_items) &&
           parsed.checklist_items.every(item => 
             item.item && 
             item.why && 
             item.example && 
             ['必須', '推奨', '任意'].includes(item.priority)
           );
  } catch (e) {
    return false;
  }
}

function validateMinutes(data) {
  try {
    const parsed = typeof data === 'string' ? JSON.parse(data) : data;
    return parsed.meeting_type &&
           parsed.meeting_date &&
           parsed.meeting_title &&
           Array.isArray(parsed.attendees) &&
           Array.isArray(parsed.agenda_items) &&
           Array.isArray(parsed.next_actions);
  } catch (e) {
    return false;
  }
}