/**
 * AI最新情報配信 & Notion更新システム
 * 朝夕のニュースと昼のアイデアを別々のNotionページに更新
 */

// ===== プロフィール情報 =====
const AUTHOR_PROFILE = {
  // 基本情報
  basic: {
    name: 'SENDER_NAME',
    nameReading: 'やまだ たろう',
    age: 36,
    birthDate: '19XX年X月X日',
    location: 'マレーシア',
    email: 'your-email@example.com',
    currentPosition: '企業A 代表取締役、企業B 代表'
  },
  
  // 学歴・資格
  education: {
    graduate: '修士XX大学大学院（2010年〜2013年）',
    undergraduate: '法学士 XX大学法学部（2007年〜2010年、成績優秀につき3年卒業）',
    highSchool: '私立XX高等学校（2004年〜2007年）',
    languages: ['英語：ビジネス上級・ネイティブレベル', '中国語：日常会話レベル'],
    awards: ['各社でMVP多数', 'XX大学学業奨励賞（学科首席）']
  },
  
  // 職歴・経験
  career: {
    current: [
      '企業A（2023年4月〜現在）設立、代表取締役（海外法人）',
      '企業B（2018年9月〜現在）設立、代表'
    ],
    past: [
      '企業D（2019年9月〜2020年5月）設立、代表取締役',
      '企業C 取締役、執行役員（2019年10月〜2020年5月）',
      '企業E 法務部員（2017年8月〜2018年8月）',
      '企業F 法務部員（2016年7月〜2017年7月）',
      '企業G 法務部員（2015年9月〜2016年6月）',
      '家業でコンサルティング業務（2013年〜2015年8月）'
    ]
  },
  
  // 専門分野・コア領域
  expertise: {
    management: {
      title: '管理部・上場準備のエキスパート',
      details: [
        '20社以上の上場準備経験、6社の上場実現',
        'バックオフィス全般の効率化・システム導入',
        'リスクマネジメント体制構築',
        '業務フロー整備・内部監査実施'
      ]
    },
    legal: {
      title: '法務部門のスペシャリスト',
      details: [
        '法務部立ち上げから運用まで一貫対応',
        '法的グレー領域のクリアランス・ロビイング経験',
        'YouTube関連法務、Web3領域の法的整理',
        'M&A・資金調達のDD対応',
        'AI活用による法務業務の自動化・効率化'
      ]
    },
    it: {
      title: '情報システム部門の構築・運用',
      details: [
        '情シス部門の立ち上げ・立て直し',
        'ISMS対応・セキュリティ施策運用',
        '各種システム導入・統廃合',
        '業務自動化（Zapier、Google Apps Script等）'
      ]
    },
    business: {
      title: '新規事業立ち上げ',
      details: [
        '法的スキーム検討・適法性確保',
        'マーケティング・資金調達',
        'システム整備・物流インフラ',
        '海外ビジネス（マレーシア）展開'
      ]
    },
    strategy: {
      title: '経営企画・組織運営',
      details: [
        '戦略立案・事業計画策定',
        '予実管理・コストカット実行',
        '人事制度・教育制度構築',
        '意思決定プロセス整備'
      ]
    }
  },
  
  // 特徴的なスキル・強み
  skills: {
    technical: [
      'SaaSツール専門知識（Freeeシリーズ、マネーフォワードシリーズ、ジョブカンシリーズ、Bakurakuシリーズ）',
      'CRM・営業支援（Salesforce・ソアスクのカスタマイズ、ワークフロー開発）',
      'プロジェクト管理（Asana、Backlog、Jooto、Trelloの導入・テンプレート化・運用）',
      'ナレッジ管理（Notion、Docbase、Google Site、Kibellaでのマニュアル整備・浸透）',
      '業務自動化（Zapier、IFTTT、Microsoft Power Automate、Google Apps Script）',
      'チャットツール（Slack、Teamsの導入・ルール整備・ワークフロー構築）',
      'AI活用（社内向けオリジナルAI開発・導入、契約書チェックAI、教育Bot開発）',
      '動画活用（YouTubeを社内マニュアル整備に活用した業務効率化）'
    ],
    legal: [
      'YouTube関連法務（複数企業でのYouTube関連法的整理・コンプライアンス対策）',
      'Web3・ブロックチェーン（賭博系、金融系、国際スキームの法的整理）',
      'フィンテック・新金融（自動車相乗り適法化、賭博系論点クリアランス）',
      'グレー領域事業化（法的論点整理・意見書取得・事業スキーム構築）',
      '規制当局対応（官庁向けロビイング、業界団体での規制制定議論）',
      '反社チェック自動化（Googleによる検索自動化、チェック体制構築）',
      '契約管理自動化（契約書保存・リスト作成・期限通知の完全自動化）'
    ],
    management: [
      '上場準備チーム（20社以上の上場準備チーム構築・運営、6社の上場実現）',
      '人事評価制度（MVV中心の評価制度整備・再構築、「メンテナンス」人材の適切評価）',
      '教育・研修制度（社内研修資料・動画作成、Bot活用による教育自動化）',
      '採用設計（人事評価制度から逆算した採用設計・運用）',
      '意思決定体制（会議体設計、職務決裁権限表整理、稟議制度整備）',
      '中間管理職育成（オペレーション落とし込みのための管理職・事務方育成）',
      '組織化推進（システム統合プロジェクト、ボトムアップ進行ルール化、ナレッジ整備）'
    ]
  },
  
  // 取引実績企業
  clients: ['企業A', '企業B', '企業C', '...（複数社）'],
  
  // 価値観・理念
  values: {
    principles: [
      '効率化と自動化による生産性向上',
      '法的リスクを適切に管理した事業運営',
      '国際的な視野での事業展開',
      '持続可能な組織体制の構築'
    ],
    approach: [
      '実務的で実行可能なソリューション提供',
      '複雑な問題のシンプル化・体系化',
      '継続的な改善と効率化の追求',
      'リスクとリターンのバランス重視'
    ]
  },
  
  // メルマガ提供価値
  newsletter: {
    content_areas: [
      {
        title: 'スタートアップ・上場準備の実務知識',
        details: [
          '上場準備の具体的なステップ（20社以上の経験から）',
          '管理部門構築のポイント（法務・情シス・経営企画の統合視点）',
          '効率的な業務フロー設計・内部監査実施',
          'リスクマネジメント体制構築・運営'
        ]
      },
      {
        title: '法務・コンプライアンスの実践的アドバイス',
        details: [
          'YouTube・Web3・フィンテック等の最新領域法務',
          'グレー領域の事業化手法（規制当局との折衝実例）',
          'AI活用による法務業務効率化（契約書チェック・反社チェック自動化）',
          'M&A・資金調達のDD実務'
        ]
      },
      {
        title: '業務効率化・DXの具体的手法',
        details: [
          '50以上のSaaSツール導入・統合実例',
          'Zapier・GAS等による業務自動化の実装',
          '社内向けAI開発・導入の実践法',
          'YouTube活用による社内教育・マニュアル整備'
        ]
      },
      {
        title: '海外展開・国際ビジネス',
        details: [
          'マレーシア・ラブアンでのタックスヘイブン活用実例',
          '海外法人設立・運営の実務（オペレーションまで）',
          '対日本向け・対海外向けビジネス立ち上げ経験',
          '国際的な法的スキーム構築'
        ]
      },
      {
        title: '経営企画・組織運営の実践',
        details: [
          '予実管理体制整備・コストカット実行',
          '人事制度・教育制度構築（MVV中心設計）',
          '意思決定プロセス整備・会議体設計',
          '新規事業立ち上げサポート（法的スキーム〜実務まで）'
        ]
      }
    ],
    target_readers: [
      'スタートアップ経営者・役員',
      '上場準備中の企業担当者',
      '管理部門・法務部門責任者',
      '新規事業開発担当者',
      '海外展開を検討する企業'
    ]
  },
  
  // 文体・トーン設定
  writing_style: {
    basic_policy: [
      '敬語：丁寧語ベース、親しみやすさも重視',
      '口調：実務的で具体的、経験に基づいた説得力',
      '一人称：「私」',
      '特徴：数字・具体例を多用、実践的なアドバイス重視'
    ],
    avoid: [
      '抽象的・理論的すぎる内容',
      '根拠のない楽観論',
      '法的アドバイスと誤解される表現',
      '機密情報・個人情報の開示'
    ],
    usage_notes: [
      '経験の具体性を重視：数字（20社以上、6社上場等）や具体的な企業名を積極的に活用',
      '実務的な視点：理論よりも実際の経験・実践に基づく内容を優先',
      '複合的な専門性：法務・情シス・経営企画の複合的な視点を活かす',
      '国際的な視野：海外展開・タックスヘイブン等の国際ビジネス経験を活用',
      '効率化・自動化：テクノロジーを活用した業務改善の視点を重視'
    ]
  }
};

// ==============================
// 設定・定数
// ==============================
const CONFIG = {
  // API Keys (PropertiesServiceに保存推奨)
  PERPLEXITY_API_KEY: 'YOUR_PERPLEXITY_API_KEY',
  GROK_API_KEY: 'YOUR_GROK_API_KEY', 
  NOTION_API_KEY: 'YOUR_NOTION_API_KEY',
  
  // Notion設定
  NOTION: {
    // ニュース用ページID（既存ページを更新）
    NEWS_PAGE_ID: '2590009db84b807cb01bddcc7e81372d',  // ハイフンなしの32文字形式
    // アイデア用ページID（既存ページを更新）
    IDEAS_PAGE_ID: '2590009db84b80afa5acee72dd29abf0'  // ハイフンなしの32文字形式
  },
  
  // メール配信設定
  EMAIL: {
    RECIPIENTS: [AUTHOR_PROFILE.basic.email], // 配信先メールアドレス
    SENDER_NAME: AUTHOR_PROFILE.basic.name + ' AI情報配信'
  },
  
  // 配信スケジュール
  SCHEDULE: {
    MORNING: { hour: 7, type: 'morning_news' },    // 朝: 昨日のニュースまとめ
    NOON: { hour: 12, type: 'business_ideas' },    // 昼: ビジネスアイデア
    EVENING: { hour: 18, type: 'evening_news' }    // 夕方: 本日のニュースまとめ
  }
};

// ==============================
// Perplexity API連携
// ==============================

/**
 * Perplexity APIで最新AI情報を取得
 */
function fetchAIInfoFromPerplexity(query, date) {
  try {
    const url = 'https://api.perplexity.ai/chat/completions';
    
    const payload = {
      model: 'sonar-pro',
      messages: [
        {
          role: 'system',
          content: 'You are an AI news analyst. Provide comprehensive, factual information about recent AI developments in Japanese.'
        },
        {
          role: 'user',
          content: `${date}の${query}について、最新の情報を日本語で詳しく教えてください。具体的な企業名、数値、発表内容を含めてください。`
        }
      ],
      temperature: 0.3,
      max_tokens: 2000,
      return_citations: true,
      search_recency_filter: "day"
    };

    const options = {
      method: 'post',
      contentType: 'application/json',
      headers: {
        'Authorization': `Bearer ${getApiKey('PERPLEXITY_API_KEY')}`,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(url, options);
    const responseData = JSON.parse(response.getContentText());
    
    if (responseData.choices && responseData.choices[0]) {
      return {
        content: responseData.choices[0].message.content,
        citations: responseData.citations || []
      };
    }
    
    return null;
  } catch (error) {
    console.error('Perplexity APIエラー:', error);
    return null;
  }
}

/**
 * 複数トピックの情報収集
 */
function gatherMultipleTopics(topics, date) {
  const results = {};
  
  topics.forEach(topic => {
    const data = fetchAIInfoFromPerplexity(topic, date);
    if (data) {
      results[topic] = data;
    }
    Utilities.sleep(1000); // レート制限対策
  });
  
  return results;
}

// ==============================
// Grok API連携
// ==============================

/**
 * Grok APIでコンテンツ生成
 */
function generateContentWithGrok(prompt, type) {
  try {
    const url = 'https://api.x.ai/v1/chat/completions';

    const payload = {
      model: 'grok-4',
      messages: [
        { 
          role: 'system', 
          content: 'You are a Japanese AI business analyst and newsletter writer. Create structured, actionable content based on the latest AI trends.' 
        },
        { 
          role: 'user', 
          content: prompt 
        }
      ],
      temperature: type === 'business_ideas' ? 0.8 : 0.6,
      max_tokens: 4096
    };

    const options = {
      method: 'post',
      contentType: 'application/json',
      headers: {
        'Authorization': `Bearer ${getApiKey('GROK_API_KEY')}`,
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(url, options);
    const data = JSON.parse(response.getContentText());

    if (data && data.choices && data.choices[0]) {
      return data.choices[0].message.content;
    }

    return null;
  } catch (error) {
    console.error('Grok APIエラー:', error);
    return null;
  }
}

// ==============================
// コンテンツ生成
// ==============================

/**
 * プロフィール情報のコンテキスト生成
 */
function generateProfileContext() {
  return `
## 執筆者情報
- 執筆者：${AUTHOR_PROFILE.basic.name}
- 現職：${AUTHOR_PROFILE.basic.currentPosition}
- 専門：上場準備・法務・情報システム・海外展開
- 実績：20社以上の上場準備経験、6社の上場実現

## 執筆方針
- 実務経験に基づく実践的な内容
- 丁寧語ベースで親しみやすく
- 一人称「私」で具体的な経験を交える
- 理論よりも実用性を重視
`;
}

/**
 * 朝のニュースコンテンツ生成
 */
function generateMorningNews() {
  const yesterday = new Date();
  yesterday.setDate(yesterday.getDate() - 1);
  
  // Perplexityで情報収集
  const topics = [
    'AI関連の最新ニュース',
    'OpenAI Anthropic Google Microsoft Meta などのAI関連企業の発表',
    'AI スタートアップ 資金調達',
    'AI 規制 法律 ガイドライン',
    '生成AI 新サービス リリース'
  ];
  
  const perplexityData = gatherMultipleTopics(topics, yesterday.toLocaleDateString('ja-JP'));
  const profileContext = generateProfileContext();
  
  // Grokでコンテンツ生成（AI Mail.jsのプロンプトを採用）
  const prompt = `
${profileContext}

あなたは${AUTHOR_PROFILE.basic.name}として、AI・テクノロジー専門のニュースアナリストの視点でメルマガを執筆します。
昨日（${yesterday.toLocaleDateString('ja-JP')}）のAI関連ニュースをまとめてください。

## Perplexityから収集した最新情報：
${JSON.stringify(perplexityData, null, 2)}

上記の最新情報を参考にしながら、以下の構成で執筆してください：

## 🌅 おはようございます！昨日のAIニュースまとめ

皆さん、おはようございます。${AUTHOR_PROFILE.basic.name}です。
上場準備や法務・情シス分野での実務経験を活かし、昨日のAI関連ニュースの中から、特にビジネス実務に直結する情報をピックアップしてお届けします。

## 1. 📈 トップニュース（最重要3件）

昨日最も注目された出来事を3つ選び、それぞれについて：
- ニュースの概要
- **実務への影響**（上場準備・管理部門の観点から）
- **私の実体験から**：類似の課題に直面した際の対応事例
- **法務・コンプライアンス視点**での注意点
- **業務効率化**への応用可能性（実際の導入事例含む）

## 2. 🏢 企業動向

### 大手テック企業のAI関連発表
具体的な企業名と発表内容を記載：
- [企業名]: [発表内容の詳細]
- **実務的な視点**：類似事例との比較と実装時の考慮点
- **私の経験談**：過去の企業でのAI導入時に直面した課題と解決策

### スタートアップの資金調達
具体的な企業名と調達額を記載：
- [企業名]: [調達額・投資家・用途]
- **上場準備の観点**：資金調達戦略と成長可能性の分析
- **実際の事例**：支援した企業での類似フェーズでの経験談

### 新サービス・製品のローンチ
- [企業名・サービス名]: [サービスの特徴・機能]
- **導入検討ポイント**：SaaSツール導入の知見から
- **実装時の注意点**：実際の導入現場で遭遇したトラブルと対策

## 3. ⚖️ 規制・政策動向

法務実務の経験を踏まえた分析：
- 各国のAI規制動向
- 倫理的な議論
- ガイドラインの発表
- **実務対応のポイント**
- **私の対応経験**：グレー領域の事業化で実際に行った規制当局との折衝事例

## 4. 🔧 業務効率化・自動化の視点

業務自動化の実装経験から見た活用可能性：
- 今回のニュースで紹介された技術の業務応用
- 導入時の注意点
- ROI算出のポイント
- **成功・失敗事例**：実際の導入プロジェクトで学んだ教訓

## 5. 🌏 海外展開への影響

海外ビジネスの経験から：
- 海外市場への影響
- 国際的な法的スキーム構築への示唆
- グローバル展開での考慮点
- **現地での実体験**：マレーシアでの事業運営で感じた課題と解決策

## 6. 明日への展望

昨日のニュースから予想される今後の動き

---

**重要な執筆指示:**
- 自己紹介は冒頭のみ簡潔に
- 実務経験は軽く触れる程度に留める
- 「私の経験では...」等の表現は必要最小限に
- 記事の主体はAI情報の分析と実用的アドバイス
- 丁寧語ベースで親しみやすい文体
- 読者に直接役立つ実践的な情報を重視
- 具体的な企業名、数値、日付は事実に基づいたもののみ使用
- 最新のAI情報を基に実用的な分析を提供
`;

  return generateContentWithGrok(prompt, 'morning_news');
}

/**
 * 昼のビジネスアイデア生成
 */
function generateBusinessIdeas() {
  const today = new Date();
  
  // 最新トレンド収集
  const trendTopics = [
    'AI 最新技術 トレンド ビジネス応用',
    '業務効率化 AI ツール 自動化',
    'AIエージェント RAG LLM 実装事例',
    '日本企業 AI導入 成功事例'
  ];
  
  const trendsData = gatherMultipleTopics(trendTopics, today.toLocaleDateString('ja-JP'));
  const profileContext = generateProfileContext();
  
  // 重複回避のため過去のテーマを取得
  const recentThemes = getRecentBusinessIdeaThemes();
  
  const prompt = `
${profileContext}

あなたは${AUTHOR_PROFILE.basic.name}として、実務経験豊富なビジネスコンサルタントの視点でメルマガを執筆します。
${today.toLocaleDateString('ja-JP')}時点の最新のAI技術トレンドを踏まえて、実践的なビジネスチャンス・アイデアと業務効率化ソリューションを提案してください。

## Perplexityから収集した最新AIトレンド：
${JSON.stringify(trendsData, null, 2)}

上記の最新トレンドを参考にしながら、以下の構成で執筆してください：

## 💡 こんにちは！AIビジネス&効率化アイデア

皆さん、こんにちは。${AUTHOR_PROFILE.basic.name}です。
上場準備・法務・情シス分野での実務経験を活かし、今日は最新のAI技術トレンドから、実際の現場で使える業務効率化アイデアとビジネスチャンスをお届けします。

【重要: 重複回避】
以下は過去2週間に配信済みのビジネスアイデアのテーマです。これらとタイトル・主題・切り口が重複しない新規テーマを必ず選定してください（同義語・言い換えも不可）。
${recentThemes.length ? recentThemes.map(t => `- ${t}`).join('\n') : '- （なし）'}

## パート1: 🚀 業務効率化・自動化アイデア

### 今週注目のAI技術を活用した効率化提案

**業務自動化の実装経験から、最も実用的なアイデアを1つ選んで詳細解説します。**

#### 🎯 [選定したアイデアのタイトル]

##### 1. なぜこのアイデアを選んだか
実務経験から、以下の理由で今最も注目すべきと判断：
- 関連する最新のAI技術やニュース
- 実際の現場で遭遇した類似課題
- なぜ今がベストタイミングなのか
- **私の失敗談**：過去に見送って後悔した類似案件の事例

##### 2. 現状の課題分析（実務経験から）
- **対象業務**: 管理部門・上場準備の観点から特定
- **現状の問題点**: 実際に目にした定量的データ（時間、コスト、エラー率等）
- **なぜ従来の方法では解決できないのか**（SaaSツール導入経験から）
- **私が直面した具体例**：某企業での同様課題とその時の対応

##### 3. 実装アプローチ（導入実績から）

**ステップ1: [初期段階の名称]**（期間：X週間）
- 具体的な作業内容（システム導入経験の活用）
- 達成すべき成果物
- **私の実体験**：某社で同様のステップ1実行時に発生した予期せぬ課題

**ステップ2-5: [各段階の詳細]**
（業務自動化の実装手順に基づく）
- **要注意ポイント**：過去のプロジェクトで実際に躓いた箇所

##### 4. 期待される効果（実績ベース）
**時間削減効果:**
過去の業務自動化導入実績から算出：
- 現状作業時間 → AI導入後の時間（X%削減）
- 月間削減時間：X時間（年間X万円の人件費削減）
- **実際の測定例**：A社での導入前後の具体的な数値変化

**ROI計算:**
予実管理の経験から：
- 投資回収期間：X ヶ月
- 3年間の累積価値：X万円
- **成功事例との比較**：類似規模の企業での実際のROI実績

##### 5. 実装時の注意点
法務・情シスの経験から：
- セキュリティ・コンプライアンス対応
- 既存システムとの統合リスク
- 組織変更管理のポイント

---

## パート2: 💼 AIビジネスアイデア

### 海外ビジネス経験から見る新機会

**海外展開の経験を活かした国際的視点のビジネス提案**

#### 💡 [選定したビジネスアイデアのタイトル]

##### 1. なぜ今このビジネスなのか
海外ビジネスの経験から：
- 海外市場での類似事例
- 日本市場での展開可能性
- マレーシアでの活動で得た市場インサイト
- **私の気づき**：現地でビジネスを実際に始めて分かった意外な障壁と機会

##### 2. 市場分析（実務経験ベース）
**市場規模:**
経営企画の経験から算出：
- TAM（全体市場）: X億円
- SOM（獲得可能市場）: X億円
- **実感ベースの検証**：実際の顧客ヒアリングで感じた市場の温度感

**競合分析:**
上場企業支援で経験した類似領域から：
- **競合との差別化で苦労した実例**：某社での競合対策の成功・失敗事例

##### 3. 法務・規制対応（専門領域）
法務実務の経験から：
- 必要な許認可・届出
- グレー領域の適法性確保
- 規制当局との折衝ポイント
- **実際の折衝体験**：某事業でのグレー領域クリアランス交渉の具体的な進め方

##### 4. 実装ロードマップ
新規事業立ち上げの実績に基づく具体的手順：

**第1四半期〜第4四半期の詳細プラン**
（上場準備支援の経験を踏まえた現実的スケジュール）
- **私の教訓**：スケジュール遅延の主な原因と対策（実体験から）

##### 5. 資金調達・投資計画
M&A・資金調達の経験から：
- 必要資金の算出根拠
- 投資家へのアプローチ方法
- バリュエーション算定のポイント
- **失敗から学んだ教訓**：某社の資金調達で犯した失敗とその後の軌道修正

## 最後に

今日ご紹介したアイデアは、上場準備・法務・情シス分野での実務経験と業務自動化の実装経験に基づく実践的な提案です。

ご質問がございましたら、${AUTHOR_PROFILE.basic.email}までお気軽にお問い合わせください。
次回の夕方のニュースレターでも、最新情報をお届けします。

${AUTHOR_PROFILE.basic.name}

---

**重要な執筆指示:**
- 自己紹介は冒頭のみ簡潔に
- 実務経験は軽く触れる程度に留める
- 「私の経験では...」等の表現は必要最小限に
- 記事の主体はAI情報の分析と実用的アドバイス
- 読者に即座に実行できる具体的なアクションプランを提供
- 丁寧語ベースで親しみやすい文体で執筆
- 最新のAI技術トレンドを基にした実践的な提案
`;

  return generateContentWithGrok(prompt, 'business_ideas');
}

/**
 * 過去のビジネスアイデアテーマを取得（重複回避用）
 */
function getRecentBusinessIdeaThemes() {
  // 実装は後で追加可能（スプレッドシートやPropertiesServiceから取得）
  return [];
}

/**
 * 夕方のニュースコンテンツ生成
 */
function generateEveningNews() {
  const today = new Date();
  
  // 本日の最新情報収集
  const todayTopics = [
    'AI 本日 今日 最新ニュース 速報',
    'OpenAI Anthropic Google 本日の発表',
    'AI企業 今日の動向 株価',
    '日本 AI 本日のニュース',
    'AIサービス 新機能 アップデート 今日'
  ];
  
  const todayData = gatherMultipleTopics(todayTopics, today.toLocaleDateString('ja-JP'));
  const profileContext = generateProfileContext();
  
  const prompt = `
${profileContext}

あなたは${AUTHOR_PROFILE.basic.name}として、AI・テクノロジー専門のニュースアナリストの視点でメルマガを執筆します。
本日（${today.toLocaleDateString('ja-JP')}）これまでに報じられたAI関連の最新ニュースをまとめてください。

## Perplexityから収集した本日の最新情報：
${JSON.stringify(todayData, null, 2)}

【重要な制約事項】
1. 上記のPerplexityデータに含まれている情報のみを使用すること
2. 時刻・企業名・数値は、データに実際に存在するもののみ記載
3. データにない情報は含めない（「詳細は未発表」等と明記）
4. 推測や予想は「〜の可能性がある」と明確に区別
5. 株価や数値データは、データソースに含まれるもののみ使用

以下の構成で執筆してください：

## 🌆 お疲れ様です！本日のAI最新ニュース

皆さん、お疲れ様です。${AUTHOR_PROFILE.basic.name}です。
本日も一日お疲れ様でした。上場準備・法務・情シス分野での実務経験を活かし、今日一日で発表された重要なAI関連ニュースを、ビジネス実務の観点から解説します。

## 1. 🚨 本日の速報・Breaking News

データに含まれる本日のニュースを時系列で整理：
（データにある時刻情報のみ記載）

## 2. 📊 本日のハイライト TOP5

データから最も重要な5つの出来事：

**1. [データに基づくタイトル]**
- 概要：（データの内容を要約）
- **実務的な見解**：ビジネスへの影響
- **実務への影響**：スタートアップ経営者への示唆

**2-5. [同様にデータに基づいて記載]**

## 3. 🏢 業界別の動き

### テック大手の動向
データに含まれる企業のみ：
- 企業名と発表内容（データに基づく）

### AI専門企業の動向
データに含まれる企業のみ：
- 企業名と動向（データに基づく）

### 国内企業の動向
データに含まれる企業のみ：
- 企業名と動き（データに基づく）

## 4. ⚖️ 法務・規制面での注目ポイント

データに含まれる規制関連情報のみ：
- 新たな規制動向（データに基づく）
- コンプライアンス対応のポイント

## 5. 🔧 業務効率化の新機会

本日発表された技術から読み取れる活用可能性：
- データに含まれる新技術の概要
- 一般的な導入検討ポイント

## 6. 📈 投資・資金調達動向

データに含まれる案件のみ：
- 資金調達案件（データに基づく具体的な金額）
- M&A情報（データに確認できるもののみ）

## 7. 明日への注目ポイント

データから確認できる予定のみ：
- 明日以降の予定（データに基づく）

---

**本日のまとめ**

収集したデータから見える本日の主要トレンドを簡潔に総括。

明日朝には昨日のニュースまとめをお届けします。
ご質問は${AUTHOR_PROFILE.basic.email}までお気軽にどうぞ。

${AUTHOR_PROFILE.basic.name}

---

**執筆上の注意:**
- Perplexityデータにない情報は一切創作しない
- 時刻や数値はデータで確認できるもののみ記載
- 株価情報はデータソースにあるもののみ使用
- 不確実な情報には必ず「報道によると」を付ける
- 夕方らしい労いの言葉を含めた親しみやすい文体
`;

  return generateContentWithGrok(prompt, 'evening_news');
}

// ==============================
// Notion API連携
// ==============================

/**
 * Notionページ更新（crypto notion.gsを参考に実装）
 */
function updateNotionPage(pageId, title, content, type) {
  try {
    const notionToken = getApiKey('NOTION_API_KEY');
    
    if (!notionToken) {
      console.error('Notion APIトークンが設定されていません');
      return null;
    }
    
    if (!pageId) {
      console.error('NotionページIDが設定されていません');
      return null;
    }
    
    console.log(`Notionページ更新開始: ${pageId}`);
    const today = new Date().toISOString().split('T')[0];
    
    // コンテンツをNotionブロック形式に変換
    const blocks = convertToNotionBlocks(content);
    console.log(`生成されたブロック数: ${blocks.length}`);
    
    let clearingSuccessful = false;
    
    try {
      // ページプロパティを更新（最終更新日時）
      console.log('ページプロパティを更新中...');
      updateNotionPageProperties(pageId, notionToken, title, type);
      console.log('ページプロパティ更新完了');
      
      // 既存のページ内容をクリア
      console.log('既存コンテンツをクリア中...');
      clearNotionPageContent(pageId, notionToken);
      console.log('既存コンテンツクリア完了');
      clearingSuccessful = true;
    } catch (clearError) {
      console.warn('既存コンテンツのクリアに失敗しました:', clearError.message);
      console.warn('コンテンツの追加は続行します（既存コンテンツの上に追加されます）');
    }
    
    try {
      // 新しいコンテンツを追加
      console.log('新しいコンテンツを追加中...');
      appendNotionContent(pageId, notionToken, blocks);
      console.log('新しいコンテンツ追加完了');
      
      if (clearingSuccessful) {
        console.log('Notion更新完了（クリア+追加）');
      } else {
        console.log('Notion更新完了（追加のみ）');
      }
      
      return { success: true, pageId: pageId };
    } catch (appendError) {
      console.error('新しいコンテンツの追加に失敗:', appendError);
      throw appendError;
    }
  } catch (error) {
    console.error('Notionページ更新エラー:', error);
    return null;
  }
}

/**
 * Notionページプロパティ更新
 */
function updateNotionPageProperties(pageId, notionToken, title, type) {
  const url = `https://api.notion.com/v1/pages/${pageId}`;
  
  const now = new Date();
  const jstTime = new Date(now.getTime() + (9 * 60 * 60 * 1000)); // JST時間に変換
  
  const payload = {
    properties: {
      "最終更新日時": {
        date: {
          start: jstTime.toISOString()
        }
      }
    }
  };
  
  const options = {
    method: 'PATCH',
    headers: {
      'Authorization': `Bearer ${notionToken}`,
      'Content-Type': 'application/json',
      'Notion-Version': '2022-06-28'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  
  try {
    console.log(`ページプロパティ更新: ${pageId}`);
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();
    
    if (responseCode === 200) {
      console.log('✓ ページプロパティ更新成功');
      console.log(`更新日時: ${jstTime.toLocaleString('ja-JP', {timeZone: 'Asia/Tokyo'})}`);
    } else {
      console.warn(`ページプロパティ更新失敗: HTTP ${responseCode}`);
      if (responseCode === 400 && responseText.includes('property')) {
        console.warn('ヒント: Notionページに「最終更新日時」プロパティ（Date型）を追加してください');
      }
    }
  } catch (error) {
    console.warn('ページプロパティ更新エラー:', error.message);
  }
}

/**
 * Notionページのコンテンツをクリア（crypto notion.gsより）
 */
function clearNotionPageContent(pageId, notionToken) {
  console.log(`Notionページ内容クリア開始: ${pageId}`);
  
  let totalDeletedCount = 0;
  let maxRetries = 10; // 最大10回まで繰り返し
  
  for (let retry = 0; retry < maxRetries; retry++) {
    try {
      console.log(`削除サイクル ${retry + 1}/${maxRetries}`);
      
      // 既存の子ブロックを取得
      const url = `https://api.notion.com/v1/blocks/${pageId}/children`;
      const getOptions = {
        method: 'GET',
        headers: {
          'Authorization': `Bearer ${notionToken}`,
          'Notion-Version': '2022-06-28'
        },
        muteHttpExceptions: true
      };
      
      const response = UrlFetchApp.fetch(url, getOptions);
      const responseCode = response.getResponseCode();
      const responseText = response.getContentText();
      
      if (responseCode === 404) {
        throw new Error(`Notionページが見つかりません: ${pageId}`);
      }
      
      if (responseCode !== 200) {
        throw new Error(`Notion APIエラー: HTTP ${responseCode}`);
      }
      
      const data = JSON.parse(responseText);
      
      // ブロックがなくなったら完了
      if (!data.results || data.results.length === 0) {
        console.log(`✓ 全ブロック削除完了 (合計${totalDeletedCount}ブロック削除)`);
        return;
      }
      
      console.log(`削除対象ブロック数: ${data.results.length}`);
      
      // 各ブロックを削除
      let deletedCount = 0;
      for (const block of data.results) {
        const deleteUrl = `https://api.notion.com/v1/blocks/${block.id}`;
        const deleteOptions = {
          method: 'DELETE',
          headers: {
            'Authorization': `Bearer ${notionToken}`,
            'Notion-Version': '2022-06-28'
          },
          muteHttpExceptions: true
        };
        
        const deleteResponse = UrlFetchApp.fetch(deleteUrl, deleteOptions);
        if (deleteResponse.getResponseCode() === 200) {
          deletedCount++;
          totalDeletedCount++;
        }
      }
      
      console.log(`サイクル${retry + 1}で${deletedCount}ブロック削除`);
      
      // 少し待機
      Utilities.sleep(500);
      
    } catch (error) {
      console.error(`クリアエラー (サイクル${retry + 1}):`, error.message);
      if (error.message.includes('見つかりません')) {
        throw error;
      }
    }
  }
  
  console.log(`最大リトライ回数到達 (合計${totalDeletedCount}ブロック削除)`);
}

/**
 * Notionページにコンテンツを追加（crypto notion.gsより）
 */
function appendNotionContent(pageId, notionToken, blocks) {
  if (!blocks || blocks.length === 0) {
    console.log('追加するブロックがありません');
    return;
  }
  
  const url = `https://api.notion.com/v1/blocks/${pageId}/children`;
  console.log(`Notion コンテンツ追加開始: ${blocks.length}ブロック`);
  
  // ブロックを100個ずつに分割（Notion APIの制限）
  const chunkSize = 100;
  const totalChunks = Math.ceil(blocks.length / chunkSize);
  
  for (let i = 0; i < blocks.length; i += chunkSize) {
    const chunkIndex = Math.floor(i / chunkSize) + 1;
    const chunk = blocks.slice(i, i + chunkSize);
    
    console.log(`チャンク ${chunkIndex}/${totalChunks} を処理中 (${chunk.length}ブロック)`);
    
    const payload = {
      children: chunk
    };
    
    const options = {
      method: 'PATCH',
      headers: {
        'Authorization': `Bearer ${notionToken}`,
        'Content-Type': 'application/json',
        'Notion-Version': '2022-06-28'
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };
    
    try {
      const response = UrlFetchApp.fetch(url, options);
      const responseCode = response.getResponseCode();
      
      console.log(`チャンク ${chunkIndex} 応答: HTTP ${responseCode}`);
      
      if (responseCode === 404) {
        throw new Error(`Notionページが見つかりません: ${pageId}`);
      }
      
      if (responseCode !== 200) {
        console.error(`チャンク ${chunkIndex} エラー: HTTP ${responseCode}`);
      } else {
        console.log(`✓ チャンク ${chunkIndex} 追加成功`);
      }
      
      // APIレート制限対策
      if (chunkIndex < totalChunks) {
        Utilities.sleep(300);
      }
      
    } catch (error) {
      console.error(`チャンク ${chunkIndex} エラー:`, error.message);
      throw error;
    }
  }
  
  console.log(`✓ 全${blocks.length}ブロックの追加完了`);
}

/**
 * マークダウンをNotionブロックに変換
 */
function convertToNotionBlocks(markdown) {
  const blocks = [];
  
  // コンテンツが空の場合のエラーハンドリング
  if (!markdown || markdown.trim() === '') {
    console.error('空のコンテンツ: Notionブロックを生成できません');
    return [{
      object: 'block',
      type: 'paragraph',
      paragraph: {
        rich_text: [{
          type: 'text',
          text: { content: 'コンテンツの生成に失敗しました。' }
        }]
      }
    }];
  }
  
  const lines = markdown.split('\n');
  
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    
    // 見出し処理（h6からh1の順で処理して正確にマッチ）
    if (line.startsWith('###### ')) {
      // h6 -> heading_3として扱う（Notionにheading_6はない）
      blocks.push({
        object: 'block',
        type: 'heading_3',
        heading_3: {
          rich_text: [{
            type: 'text',
            text: { content: line.substring(7) }
          }]
        }
      });
    } else if (line.startsWith('##### ')) {
      // h5 -> heading_3として扱う
      blocks.push({
        object: 'block',
        type: 'heading_3',
        heading_3: {
          rich_text: [{
            type: 'text',
            text: { content: line.substring(6) }
          }]
        }
      });
    } else if (line.startsWith('#### ')) {
      // h4 -> heading_3として扱う
      blocks.push({
        object: 'block',
        type: 'heading_3',
        heading_3: {
          rich_text: [{
            type: 'text',
            text: { content: line.substring(5) }
          }]
        }
      });
    } else if (line.startsWith('### ')) {
      // H3見出し
      blocks.push({
        object: 'block',
        type: 'heading_3',
        heading_3: {
          rich_text: [{
            type: 'text',
            text: { content: line.substring(4) }
          }]
        }
      });
    } else if (line.startsWith('## ')) {
      // H2見出し
      blocks.push({
        object: 'block',
        type: 'heading_2',
        heading_2: {
          rich_text: [{
            type: 'text',
            text: { content: line.substring(3) }
          }]
        }
      });
    } else if (line.startsWith('# ')) {
      // H1見出し
      blocks.push({
        object: 'block',
        type: 'heading_1',
        heading_1: {
          rich_text: [{
            type: 'text',
            text: { content: line.substring(2) }
          }]
        }
      });
    } else if (line.match(/^\d+\./)) {
      // 番号付きリスト
      blocks.push({
        object: 'block',
        type: 'numbered_list_item',
        numbered_list_item: {
          rich_text: [{
            type: 'text',
            text: { content: line.replace(/^\d+\.\s*/, '') }
          }]
        }
      });
    } else if (line.startsWith('- ') || line.startsWith('* ')) {
      // 箇条書き
      blocks.push({
        object: 'block',
        type: 'bulleted_list_item',
        bulleted_list_item: {
          rich_text: [{
            type: 'text',
            text: { content: line.replace(/^[-*]\s*/, '') }
          }]
        }
      });
    } else if (line.trim() !== '') {
      // 通常の段落
      blocks.push({
        object: 'block',
        type: 'paragraph',
        paragraph: {
          rich_text: [{
            type: 'text',
            text: { content: line }
          }]
        }
      });
    }
  }
  
  // Notionは最大100ブロックまで
  return blocks.slice(0, 100);
}

// ==============================
// メール送信
// ==============================

/**
 * HTMLメールテンプレート生成
 */
function createEmailHTML(content, type) {
  const getTheme = (type) => {
    switch(type) {
      case 'morning_news':
        return {
          emoji: '🌅',
          title: 'AIニュース朝刊',
          gradient: 'linear-gradient(135deg, #667eea 0%, #764ba2 100%)'
        };
      case 'business_ideas':
        return {
          emoji: '💡',
          title: 'AIビジネス&効率化アイデア',
          gradient: 'linear-gradient(135deg, #f093fb 0%, #f5576c 100%)'
        };
      case 'evening_news':
        return {
          emoji: '🌆',
          title: 'AI最新ニュース',
          gradient: 'linear-gradient(135deg, #4facfe 0%, #00f2fe 100%)'
        };
      default:
        return {
          emoji: '📰',
          title: 'AI情報',
          gradient: 'linear-gradient(135deg, #667eea 0%, #764ba2 100%)'
        };
    }
  };
  
  const theme = getTheme(type);
  
  // マークダウンをHTMLに変換（改良版）
  let formattedContent = content
    // 見出しの変換（h4, h5, h6も追加）
    .replace(/^######\s+(.+)$/gm, '<h6>$1</h6>')
    .replace(/^#####\s+(.+)$/gm, '<h5>$1</h5>')
    .replace(/^####\s+(.+)$/gm, '<h4>$1</h4>')
    .replace(/^###\s+(.+)$/gm, '<h3>$1</h3>')
    .replace(/^##\s+(.+)$/gm, '<h2>$1</h2>')
    .replace(/^#\s+(.+)$/gm, '<h1>$1</h1>')
    // 太字の変換
    .replace(/\*\*([^*]+)\*\*/g, '<strong>$1</strong>')
    // 箇条書きの変換
    .replace(/^[\s]*[-*]\s+(.+)$/gm, '<li>$1</li>')
    .replace(/^\d+\.\s+(.+)$/gm, '<li>$1</li>');
  
  // 連続するliタグをul/olで囲む
  formattedContent = formattedContent.replace(/(<li[^>]*>.*?<\/li>\s*)+/gs, function(match) {
    // 数字リストかどうかを判断
    if (content.includes('1.') || content.includes('2.')) {
      return '<ol>' + match + '</ol>';
    } else {
      return '<ul>' + match + '</ul>';
    }
  });
  
  // 段落の処理
  formattedContent = formattedContent
    .split('\n\n')
    .map(paragraph => {
      // 既にHTMLタグで囲まれている場合はそのまま
      if (paragraph.match(/^<[^>]+>/)) {
        return paragraph;
      }
      // 空行は無視
      if (paragraph.trim() === '') {
        return '';
      }
      // 通常のテキストはpタグで囲む
      return '<p>' + paragraph + '</p>';
    })
    .filter(p => p !== '')
    .join('\n');
  
  return `
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <style>
    body {
      font-family: 'Helvetica Neue', -apple-system, BlinkMacSystemFont, 'Segoe UI', Arial, sans-serif;
      line-height: 1.6;
      color: #333;
      max-width: 800px;
      margin: 0 auto;
      padding: 20px;
      background: ${theme.gradient};
    }
    .container {
      background: white;
      border-radius: 16px;
      padding: 40px;
      box-shadow: 0 20px 60px rgba(0,0,0,0.15);
    }
    .header {
      text-align: center;
      margin-bottom: 35px;
      padding-bottom: 25px;
      border-bottom: 3px solid #667eea;
    }
    .emoji {
      font-size: 48px;
      margin-bottom: 10px;
    }
    .title {
      color: #2c3e50;
      font-size: 32px;
      margin: 10px 0;
      font-weight: 700;
    }
    .date {
      color: #7f8c8d;
      font-size: 14px;
      margin-top: 5px;
    }
    h1 {
      color: #2c3e50;
      font-size: 28px;
      border-bottom: 3px solid #667eea;
      padding-bottom: 10px;
      margin-top: 30px;
    }
    h2 {
      color: #34495e;
      font-size: 24px;
      margin-top: 25px;
      margin-bottom: 15px;
      padding-bottom: 8px;
      border-bottom: 2px solid #ecf0f1;
    }
    h3 {
      color: #555;
      font-size: 20px;
      margin-top: 20px;
      margin-bottom: 10px;
    }
    h4 {
      color: #666;
      font-size: 18px;
      margin-top: 15px;
      margin-bottom: 10px;
      font-weight: 600;
    }
    h5 {
      color: #777;
      font-size: 16px;
      margin-top: 15px;
      margin-bottom: 8px;
      font-weight: 600;
    }
    h6 {
      color: #888;
      font-size: 14px;
      margin-top: 10px;
      margin-bottom: 8px;
      font-weight: 600;
    }
    p {
      margin: 12px 0;
      line-height: 1.7;
    }
    ul, ol {
      padding-left: 25px;
      margin: 15px 0;
    }
    li {
      margin: 8px 0;
      line-height: 1.6;
    }
    strong {
      color: #2c3e50;
      font-weight: 600;
    }
    .footer {
      text-align: center;
      margin-top: 40px;
      padding-top: 25px;
      border-top: 1px solid #ecf0f1;
      color: #7f8c8d;
      font-size: 13px;
    }
    .footer a {
      color: #667eea;
      text-decoration: none;
    }
    .footer a:hover {
      text-decoration: underline;
    }
    @media (max-width: 600px) {
      .container {
        padding: 25px;
      }
      .title {
        font-size: 26px;
      }
      h1 {
        font-size: 24px;
      }
      h2 {
        font-size: 20px;
      }
      h3 {
        font-size: 18px;
      }
    }
  </style>
</head>
<body>
  <div class="container">
    <div class="header">
      <div class="emoji">${theme.emoji}</div>
      <div class="title">${theme.title}</div>
      <div class="date">${new Date().toLocaleString('ja-JP', {
        year: 'numeric',
        month: 'long',
        day: 'numeric',
        weekday: 'long',
        hour: '2-digit',
        minute: '2-digit'
      })}</div>
    </div>
    
    <div class="content">
      ${formattedContent}
    </div>
    
    <div class="footer">
      <p><strong>配信スケジュール</strong></p>
      <p>🌅 朝7時: 昨日のニュースまとめ | 💡 昼12時: ビジネス&効率化アイデア | 🌆 夕方18時: 本日の最新ニュース</p>
      <p style="margin-top: 20px;">
        お問い合わせ: <a href="mailto:${AUTHOR_PROFILE.basic.email}">${AUTHOR_PROFILE.basic.email}</a>
      </p>
      <p style="margin-top: 10px; font-size: 11px; color: #95a5a6;">
        © 2024 ${AUTHOR_PROFILE.basic.name} - AI Newsletter System<br>
        Powered by Perplexity AI & Grok-4
      </p>
    </div>
  </div>
</body>
</html>`;
}

/**
 * メール送信
 */
function sendEmail(subject, htmlContent, recipients) {
  recipients.forEach(to => {
    try {
      MailApp.sendEmail({
        to: to,
        subject: subject,
        htmlBody: htmlContent,
        name: CONFIG.EMAIL.SENDER_NAME
      });
      console.log(`メール送信成功: ${to}`);
    } catch (error) {
      console.error(`メール送信失敗 (${to}):`, error);
    }
  });
}

// ==============================
// 配信実行関数
// ==============================

/**
 * 朝の配信
 */
function deliverMorningNews() {
  const startTime = Date.now();
  let notionSuccess = false;
  let emailSuccess = false;
  let recipients = [];
  
  try {
    console.log('朝のニュース配信開始...');
    
    // 設定とレシピエント取得
    const settings = getSettings();
    recipients = getRecipients();
    
    // 配信有効チェック
    if (settings['配信有効'] === 'FALSE') {
      console.log('配信が無効になっています。');
      return;
    }
    
    const content = generateMorningNews();
    if (!content) {
      throw new Error('朝のニュースコンテンツ生成失敗');
    }
    
    // コンテンツ重複チェック
    if (isContentDuplicate(content, 3)) {
      console.warn('重複コンテンツを検出しました。再生成します...');
      // 重複が検出された場合はスキップまたは再生成
      throw new Error('重複コンテンツ検出');
    }
    
    const today = new Date();
    const title = `AIニュース朝刊 - ${today.toLocaleDateString('ja-JP')}`;
    
    // Notion更新（既存ページを更新）
    const newsPageId = settings['Notion News Page ID'] || CONFIG.NOTION.NEWS_PAGE_ID;
    const notionResult = updateNotionPage(
      newsPageId,
      title,
      content,
      'news'
    );
    notionSuccess = !!notionResult;
    
    // メール送信
    const htmlContent = createEmailHTML(content, 'morning_news');
    sendEmail(
      `🌅 ${title}`,
      htmlContent,
      recipients
    );
    emailSuccess = true;
    
    // ログ記録
    const duration = Date.now() - startTime;
    logDelivery('morning_news', title, recipients.length, notionSuccess, emailSuccess, duration, null, 'SUCCESS');
    
    // コンテンツログ記録
    appendContentLog({
      type: 'morning_news',
      subject: `🌅 ${title}`,
      to: recipients,
      content: content
    });
    
    console.log('朝のニュース配信完了');
    
  } catch (error) {
    console.error('朝のニュース配信エラー:', error);
    
    // ログ記録（エラー）
    const duration = Date.now() - startTime;
    const title = `AIニュース朝刊 - ${new Date().toLocaleDateString('ja-JP')}`;
    logDelivery('morning_news', title, recipients.length, notionSuccess, emailSuccess, duration, error, 'ERROR');
    
    sendErrorNotification(error, 'morning');
  }
}

/**
 * 昼の配信
 */
function deliverBusinessIdeas() {
  const startTime = Date.now();
  let notionSuccess = false;
  let emailSuccess = false;
  let recipients = [];
  
  try {
    console.log('ビジネスアイデア配信開始...');
    
    // 設定とレシピエント取得
    const settings = getSettings();
    recipients = getRecipients();
    
    // 配信有効チェック
    if (settings['配信有効'] === 'FALSE') {
      console.log('配信が無効になっています。');
      return;
    }
    
    // 過去のビジネスアイデアテーマを取得（重複回避用）
    const recentThemes = getRecentBusinessIdeaThemes(7);
    console.log(`過去7日間のテーマ数: ${recentThemes.length}`);
    
    const content = generateBusinessIdeas();
    if (!content) {
      throw new Error('ビジネスアイデアコンテンツ生成失敗');
    }
    
    // コンテンツ重複チェック
    if (isContentDuplicate(content, 7)) {
      console.warn('重複コンテンツを検出しました。再生成します...');
      // 重複が検出された場合はスキップまたは再生成
      throw new Error('重複コンテンツ検出');
    }
    
    const today = new Date();
    const title = `AIビジネス&効率化アイデア - ${today.toLocaleDateString('ja-JP')}`;
    
    // Notion更新（アイデア専用ページを更新）
    const ideasPageId = settings['Notion Ideas Page ID'] || CONFIG.NOTION.IDEAS_PAGE_ID;
    const notionResult = updateNotionPage(
      ideasPageId,
      title,
      content,
      'business_ideas'
    );
    notionSuccess = !!notionResult;
    
    // メール送信
    const htmlContent = createEmailHTML(content, 'business_ideas');
    sendEmail(
      `💡 ${title}`,
      htmlContent,
      recipients
    );
    emailSuccess = true;
    
    // ログ記録
    const duration = Date.now() - startTime;
    logDelivery('business_ideas', title, recipients.length, notionSuccess, emailSuccess, duration, null, 'SUCCESS');
    
    // コンテンツログ記録
    appendContentLog({
      type: 'business_ideas',
      subject: `💡 ${title}`,
      to: recipients,
      content: content
    });
    
    console.log('ビジネスアイデア配信完了');
    
  } catch (error) {
    console.error('ビジネスアイデア配信エラー:', error);
    
    // ログ記録（エラー）
    const duration = Date.now() - startTime;
    const title = `AIビジネス&効率化アイデア - ${new Date().toLocaleDateString('ja-JP')}`;
    logDelivery('business_ideas', title, recipients.length, notionSuccess, emailSuccess, duration, error, 'ERROR');
    
    sendErrorNotification(error, 'noon');
  }
}

/**
 * 夕方の配信
 */
function deliverEveningNews() {
  const startTime = Date.now();
  let notionSuccess = false;
  let emailSuccess = false;
  let recipients = [];
  
  try {
    console.log('夕方のニュース配信開始...');
    
    // 設定とレシピエント取得
    const settings = getSettings();
    recipients = getRecipients();
    
    // 配信有効チェック
    if (settings['配信有効'] === 'FALSE') {
      console.log('配信が無効になっています。');
      return;
    }
    
    const content = generateEveningNews();
    if (!content) {
      throw new Error('夕方のニュースコンテンツ生成失敗');
    }
    
    // コンテンツ重複チェック
    if (isContentDuplicate(content, 3)) {
      console.warn('重複コンテンツを検出しました。再生成します...');
      // 重複が検出された場合はスキップまたは再生成
      throw new Error('重複コンテンツ検出');
    }
    
    const today = new Date();
    const title = `AI最新ニュース - ${today.toLocaleDateString('ja-JP')}`;
    
    // Notion更新（既存ページを更新）
    const newsPageId = settings['Notion News Page ID'] || CONFIG.NOTION.NEWS_PAGE_ID;
    const notionResult = updateNotionPage(
      newsPageId,
      title,
      content,
      'news'
    );
    notionSuccess = !!notionResult;
    
    // メール送信
    const htmlContent = createEmailHTML(content, 'evening_news');
    sendEmail(
      `🌆 ${title}`,
      htmlContent,
      recipients
    );
    emailSuccess = true;
    
    // ログ記録
    const duration = Date.now() - startTime;
    logDelivery('evening_news', title, recipients.length, notionSuccess, emailSuccess, duration, null, 'SUCCESS');
    
    // コンテンツログ記録
    appendContentLog({
      type: 'evening_news',
      subject: `🌆 ${title}`,
      to: recipients,
      content: content
    });
    
    console.log('夕方のニュース配信完了');
    
  } catch (error) {
    console.error('夕方のニュース配信エラー:', error);
    
    // ログ記録（エラー）
    const duration = Date.now() - startTime;
    const title = `AI最新ニュース - ${new Date().toLocaleDateString('ja-JP')}`;
    logDelivery('evening_news', title, recipients.length, notionSuccess, emailSuccess, duration, error, 'ERROR');
    
    sendErrorNotification(error, 'evening');
  }
}

// ==============================
// ユーティリティ関数
// ==============================

/**
 * APIキー取得（PropertiesService推奨）
 */
function getApiKey(keyName) {
  // PropertiesServiceから取得を試みる
  const scriptProperties = PropertiesService.getScriptProperties();
  const key = scriptProperties.getProperty(keyName);
  
  if (key) {
    console.log(`APIキー取得成功 (PropertiesService): ${keyName}`);
    return key;
  }
  
  // フォールバック: CONFIGから取得
  if (CONFIG[keyName]) {
    console.log(`APIキー取得成功 (CONFIG): ${keyName}`);
    return CONFIG[keyName];
  }
  
  console.error(`APIキーが見つかりません: ${keyName}`);
  return '';
}

/**
 * エラー通知送信
 */
function sendErrorNotification(error, timing) {
  const subject = `⚠️ AI情報配信エラー（${timing}）`;
  const body = `
配信エラーが発生しました。

タイミング: ${timing}
エラー内容: ${error.toString()}
スタックトレース: ${error.stack || 'なし'}
発生時刻: ${new Date().toLocaleString('ja-JP')}
`;
  
  CONFIG.EMAIL.RECIPIENTS.forEach(to => {
    try {
      MailApp.sendEmail({
        to: to,
        subject: subject,
        body: body
      });
    } catch (e) {
      console.error('エラー通知送信失敗:', e);
    }
  });
}

// ==============================
// トリガー設定
// ==============================

/**
 * 自動実行トリガーを設定
 */
function setupTriggers() {
  // 既存トリガーを削除
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    const handler = trigger.getHandlerFunction();
    if (handler === 'deliverMorningNews' || 
        handler === 'deliverBusinessIdeas' || 
        handler === 'deliverEveningNews') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  // 朝の配信（7時）
  ScriptApp.newTrigger('deliverMorningNews')
    .timeBased()
    .everyDays(1)
    .atHour(7)
    .create();
  
  // 昼の配信（12時）
  ScriptApp.newTrigger('deliverBusinessIdeas')
    .timeBased()
    .everyDays(1)
    .atHour(12)
    .create();
  
  // 夕方の配信（18時）
  ScriptApp.newTrigger('deliverEveningNews')
    .timeBased()
    .everyDays(1)
    .atHour(18)
    .create();
  
  console.log('トリガー設定完了');
}

// ==============================
// スプレッドシート初期化
// ==============================

/**
 * スプレッドシート定数
 */
const SHEETS = {
  SETTINGS: 'Settings',
  RECIPIENTS: 'Recipients', 
  LOGS: 'Logs',
  CONTENT_LOGS: 'ContentLogs' // コンテンツログ用シート追加
};

/**
 * スプレッドシート初期セットアップ
 */
function setupSpreadsheet() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 既存のシートをチェック
  const existingSheets = ss.getSheets().map(sheet => sheet.getName());
  const requiredSheets = [SHEETS.SETTINGS, SHEETS.RECIPIENTS, SHEETS.LOGS, SHEETS.CONTENT_LOGS];
  const existingRequired = requiredSheets.filter(name => existingSheets.includes(name));
  
  // 既存シートがある場合は確認
  if (existingRequired.length > 0) {
    const result = ui.alert(
      'スプレッドシート初期化',
      `既存のシート（${existingRequired.join(', ')}）が見つかりました。\n上書きしますか？`,
      ui.ButtonSet.YES_NO_CANCEL
    );
    
    if (result === ui.Button.NO) {
      ui.alert('キャンセル', '初期化をキャンセルしました。', ui.ButtonSet.OK);
      return;
    } else if (result === ui.Button.CANCEL) {
      return;
    }
    
    // 既存シートを削除
    existingRequired.forEach(sheetName => {
      const sheet = ss.getSheetByName(sheetName);
      if (sheet) {
        ss.deleteSheet(sheet);
      }
    });
  }
  
  try {
    // Settingsシートを作成
    createSettingsSheet(ss);
    
    // Recipientsシートを作成
    createRecipientsSheet(ss);
    
    // Logsシートを作成
    createLogsSheet(ss);
    
    // ContentLogsシートを作成
    createContentLogsSheet(ss);
    
    ui.alert(
      '初期化完了',
      'スプレッドシートの初期化が完了しました。\n\n作成されたシート:\n• Settings（基本設定）\n• Recipients（配信先）\n• Logs（配信ログ）\n• ContentLogs（コンテンツログ）',
      ui.ButtonSet.OK
    );
    
  } catch (error) {
    console.error('スプレッドシート初期化エラー:', error);
    ui.alert(
      'エラー',
      `初期化中にエラーが発生しました：\n${error.toString()}`,
      ui.ButtonSet.OK
    );
  }
}

/**
 * Settingsシート作成
 */
function createSettingsSheet(spreadsheet) {
  const sheet = spreadsheet.insertSheet(SHEETS.SETTINGS);
  
  // ヘッダー設定
  const headers = [
    ['項目', '値', '説明']
  ];
  
  // 初期データ
  const data = [
    ['配信有効', 'TRUE', 'メール配信を有効にするかどうか（TRUE/FALSE）'],
    ['配信先メールアドレス', AUTHOR_PROFILE.basic.email, 'デフォルトの配信先メールアドレス'],
    ['Perplexity API Key', '', 'Perplexity AIのAPIキー（PropertiesServiceでの管理推奨）'],
    ['Grok API Key', '', 'Grok AIのAPIキー（PropertiesServiceでの管理推奨）'],
    ['Notion API Key', '', 'Notion APIキー（PropertiesServiceでの管理推奨）'],
    ['Notion News Page ID', CONFIG.NOTION.NEWS_PAGE_ID, 'NotionニュースページID'],
    ['Notion Ideas Page ID', CONFIG.NOTION.IDEAS_PAGE_ID, 'NotionアイデアページID'],
    ['メール送信者名', CONFIG.EMAIL.SENDER_NAME, 'メール送信時の送信者名'],
    ['朝の配信時刻', '7', '朝のニュース配信時刻（24時間制）'],
    ['昼の配信時刻', '12', 'ビジネスアイデア配信時刻（24時間制）'],
    ['夕方の配信時刻', '18', '夕方のニュース配信時刻（24時間制）']
  ];
  
  // データを設定
  sheet.getRange(1, 1, headers.length, headers[0].length).setValues(headers);
  sheet.getRange(2, 1, data.length, data[0].length).setValues(data);
  
  // 書式設定
  sheet.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground('#4285F4').setFontColor('white');
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, 3);
  
  // 列幅調整
  sheet.setColumnWidth(1, 200);
  sheet.setColumnWidth(2, 300);
  sheet.setColumnWidth(3, 400);
  
  // データ検証（配信有効）
  const validationRange = sheet.getRange('B2');
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['TRUE', 'FALSE'])
    .setAllowInvalid(false)
    .build();
  validationRange.setDataValidation(rule);
}

/**
 * Recipientsシート作成
 */
function createRecipientsSheet(spreadsheet) {
  const sheet = spreadsheet.insertSheet(SHEETS.RECIPIENTS);
  
  // ヘッダー設定
  const headers = [
    ['メールアドレス', '氏名', '組織', '配信有効', '登録日', '備考']
  ];
  
  // デフォルトデータ
  const data = [
    [AUTHOR_PROFILE.basic.email, AUTHOR_PROFILE.basic.name, AUTHOR_PROFILE.basic.currentPosition, 'TRUE', new Date(), '管理者']
  ];
  
  // データを設定
  sheet.getRange(1, 1, headers.length, headers[0].length).setValues(headers);
  sheet.getRange(2, 1, data.length, data[0].length).setValues(data);
  
  // 書式設定
  sheet.getRange(1, 1, 1, 6).setFontWeight('bold').setBackground('#34A853').setFontColor('white');
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, 6);
  
  // データ検証（配信有効）
  const validationRange = sheet.getRange('D:D');
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['TRUE', 'FALSE'])
    .setAllowInvalid(false)
    .build();
  validationRange.setDataValidation(rule);
  
  // 日付書式設定
  sheet.getRange('E:E').setNumberFormat('yyyy/mm/dd');
}

/**
 * Logsシート作成
 */
function createLogsSheet(spreadsheet) {
  const sheet = spreadsheet.insertSheet(SHEETS.LOGS);
  
  // ヘッダー設定
  const headers = [
    ['実行日時', '配信タイプ', 'タイトル', '配信先数', 'Notion更新', 'メール送信', '所要時間(秒)', 'エラー詳細', 'ステータス', '備考']
  ];
  
  // データを設定
  sheet.getRange(1, 1, headers.length, headers[0].length).setValues(headers);
  
  // 書式設定
  sheet.getRange(1, 1, 1, 10).setFontWeight('bold').setBackground('#EA4335').setFontColor('white');
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, 10);
  
  // 列幅調整
  sheet.setColumnWidth(1, 140); // 実行日時
  sheet.setColumnWidth(2, 120); // 配信タイプ
  sheet.setColumnWidth(3, 200); // タイトル
  sheet.setColumnWidth(8, 250); // エラー詳細
  
  // 条件付き書式（ステータス）
  const statusRange = sheet.getRange('I:I');
  
  // 成功時（緑）
  const successRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('SUCCESS')
    .setBackground('#D9EAD3')
    .setFontColor('#137333')
    .setRanges([statusRange])
    .build();
  
  // エラー時（赤）
  const errorRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('ERROR')
    .setBackground('#F4CCCC')
    .setFontColor('#CC0000')
    .setRanges([statusRange])
    .build();
  
  sheet.setConditionalFormatRules([successRule, errorRule]);
}

/**
 * ContentLogsシート作成
 */
function createContentLogsSheet(spreadsheet) {
  const sheet = spreadsheet.insertSheet(SHEETS.CONTENT_LOGS);
  
  // ヘッダー設定（AI Mail.jsを参考に拡張）
  const headers = [
    ['送信日時', '配信種別', '件名', '配信先', 'メインテーマ', 'サブテーマ', 'キーワード', '詳細要約（300文字）', '主要トピック', 'コンテンツハッシュ']
  ];
  
  // データを設定
  sheet.getRange(1, 1, headers.length, headers[0].length).setValues(headers);
  
  // 書式設定
  sheet.getRange(1, 1, 1, 10).setFontWeight('bold').setBackground('#4285F4').setFontColor('white');
  sheet.setFrozenRows(1);
  
  // 列幅調整
  sheet.setColumnWidth(1, 140); // 送信日時
  sheet.setColumnWidth(2, 100); // 配信種別
  sheet.setColumnWidth(3, 250); // 件名
  sheet.setColumnWidth(4, 200); // 配信先
  sheet.setColumnWidth(5, 200); // メインテーマ
  sheet.setColumnWidth(6, 200); // サブテーマ
  sheet.setColumnWidth(7, 250); // キーワード
  sheet.setColumnWidth(8, 350); // 詳細要約
  sheet.setColumnWidth(9, 200); // 主要トピック
  sheet.setColumnWidth(10, 150); // コンテンツハッシュ
}

/**
 * 設定値取得
 */
function getSettings() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.SETTINGS);
  if (!sheet) {
    console.warn('Settingsシートが存在しません。デフォルト値を使用します。');
    return {};
  }
  
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
  const settings = {};
  
  data.forEach(row => {
    if (row[0] && row[1] !== '') {
      settings[row[0]] = row[1];
    }
  });
  
  return settings;
}

/**
 * 配信先リスト取得
 */
function getRecipients() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.RECIPIENTS);
  if (!sheet) {
    console.warn('Recipientsシートが存在しません。デフォルト配信先を使用します。');
    return [AUTHOR_PROFILE.basic.email];
  }
  
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues();
  return data
    .filter(row => row[0] && row[3] === 'TRUE') // メールアドレスがあり、配信有効
    .map(row => row[0]);
}

/**
 * 配信ログ記録
 */
function logDelivery(type, title, recipientCount, notionSuccess, emailSuccess, duration, error = null, status = 'SUCCESS') {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.LOGS);
  if (!sheet) {
    console.warn('Logsシートが存在しません。ログ記録をスキップします。');
    return;
  }
  
  const logData = [
    new Date(),
    type,
    title,
    recipientCount,
    notionSuccess ? 'OK' : 'NG',
    emailSuccess ? 'OK' : 'NG', 
    Math.round(duration / 1000),
    error ? error.toString().substring(0, 200) : '',
    status,
    ''
  ];
  
  sheet.appendRow(logData);
}

/**
 * コンテンツログ追記（AI Mail.jsを参考に実装）
 */
function appendContentLog(entry) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.CONTENT_LOGS);
    if (!sheet) {
      console.warn('ContentLogsシートが存在しません');
      return;
    }
    
    // コンテンツから詳細情報を抽出
    const contentAnalysis = analyzeNewsletterContent(entry.content || '', entry.type);
    
    // コンテンツのハッシュ値を生成（重複チェック用）
    const contentHash = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, entry.content || '')
      .map(byte => ('0' + (byte & 0xFF).toString(16)).slice(-2))
      .join('')
      .substring(0, 8); // 最初の8文字
    
    const row = [
      new Date(),
      entry.type,
      entry.subject,
      Array.isArray(entry.to) ? entry.to.join(', ') : entry.to,
      contentAnalysis.mainTheme,
      contentAnalysis.subThemes.join(', '),
      contentAnalysis.keywords.join(', '),
      contentAnalysis.detailedSummary,
      contentAnalysis.mainTopics.join(' | '),
      contentHash
    ];
    
    sheet.appendRow(row);
  } catch (e) {
    console.error('コンテンツログ書き込みエラー:', e);
  }
}

/**
 * ニュースレターコンテンツ分析（AI Mail.jsより移植）
 */
function analyzeNewsletterContent(content, type) {
  const analysis = {
    mainTheme: '',
    subThemes: [],
    keywords: [],
    detailedSummary: '',
    mainTopics: []
  };
  
  if (!content) return analysis;
  
  try {
    // メインテーマの抽出
    let mainTheme = '';
    
    // パターン1: #### 💡 [タイトル]
    let match = /####\s*💡\s*\[(.+?)\]/.exec(content);
    if (match) mainTheme = match[1].trim();
    
    // パターン2: #### 🎯 [タイトル]
    if (!mainTheme) {
      match = /####\s*🎯\s*\[(.+?)\]/.exec(content);
      if (match) mainTheme = match[1].trim();
    }
    
    // パターン3: 最初の主要な見出し
    if (!mainTheme) {
      match = /^##\s+(.+)$/m.exec(content);
      if (match) mainTheme = match[1].trim();
    }
    
    analysis.mainTheme = mainTheme || 'テーマ不明';
    
    // サブテーマの抽出
    const subThemeMatches = content.match(/^###\s+(.+)$/gm) || [];
    analysis.subThemes = subThemeMatches.map(h => h.replace(/^#+\s*/, '').trim()).slice(0, 3);
    
    // キーワードの抽出
    const keywords = new Set();
    
    // 太字テキストからキーワード抽出
    const boldMatches = content.match(/\*\*([^*]+)\*\*/g) || [];
    boldMatches.forEach(bold => {
      const keyword = bold.replace(/\*\*/g, '').trim();
      if (keyword.length > 1 && keyword.length < 20) {
        keywords.add(keyword);
      }
    });
    
    // 技術用語、ビジネス用語を抽出
    const techTerms = content.match(/AI|機械学習|深層学習|LLM|GPT|Claude|Gemini|RAG|自動化|効率化|DX|デジタル変革|SaaS|API|クラウド/g) || [];
    techTerms.forEach(term => keywords.add(term));
    
    analysis.keywords = Array.from(keywords).slice(0, 8);
    
    // 詳細要約の生成（300文字程度）
    let summary = content.replace(/^#+\s*[^\n]*\n/gm, '').trim().substring(0, 297);
    if (summary.length === 297) summary += '...';
    analysis.detailedSummary = summary;
    
    // 主要トピックの抽出
    const topics = [];
    if (type === 'business_ideas') {
      if (content.includes('業務効率化') || content.includes('自動化')) topics.push('業務効率化');
      if (content.includes('ビジネスアイデア') || content.includes('新規事業')) topics.push('ビジネス戦略');
      if (content.includes('AI導入') || content.includes('システム')) topics.push('AI導入');
      if (content.includes('ROI') || content.includes('投資回収')) topics.push('ROI分析');
      if (content.includes('海外') || content.includes('国際')) topics.push('海外展開');
    } else {
      if (content.includes('OpenAI') || content.includes('Anthropic') || content.includes('Google')) topics.push('AI企業動向');
      if (content.includes('資金調達') || content.includes('投資')) topics.push('資金調達');
      if (content.includes('規制') || content.includes('法律')) topics.push('規制・法務');
      if (content.includes('新サービス') || content.includes('リリース')) topics.push('新サービス');
    }
    
    analysis.mainTopics = topics.length > 0 ? topics : ['一般AI情報'];
    
  } catch (e) {
    console.error('コンテンツ分析エラー:', e);
    analysis.mainTheme = 'エラー';
  }
  
  return analysis;
}

/**
 * 過去N日分のビジネスアイデアテーマを取得（重複回避用）
 */
function getRecentBusinessIdeaThemes(days) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.CONTENT_LOGS);
    if (!sheet) return [];
    
    const range = sheet.getDataRange();
    const values = range.getValues();
    const now = new Date();
    const since = new Date(now.getTime() - days * 24 * 60 * 60 * 1000);
    
    const headers = values[0];
    const idxDate = headers.indexOf('送信日時');
    const idxType = headers.indexOf('配信種別');
    const idxMainTheme = headers.indexOf('メインテーマ');
    const idxSubThemes = headers.indexOf('サブテーマ');
    const idxKeywords = headers.indexOf('キーワード');
    
    const themes = [];
    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      const date = new Date(row[idxDate]);
      const type = row[idxType];
      
      if (date >= since && type === 'business_ideas') {
        themes.push({
          mainTheme: row[idxMainTheme],
          subThemes: row[idxSubThemes] ? row[idxSubThemes].split(', ') : [],
          keywords: row[idxKeywords] ? row[idxKeywords].split(', ') : [],
          date: date
        });
      }
    }
    
    return themes;
  } catch (e) {
    console.error('過去テーマ取得エラー:', e);
    return [];
  }
}

/**
 * コンテンツ重複チェック
 */
function isContentDuplicate(content, days = 7) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.CONTENT_LOGS);
    if (!sheet) return false;
    
    // コンテンツのハッシュ値を計算
    const contentHash = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, content)
      .map(byte => ('0' + (byte & 0xFF).toString(16)).slice(-2))
      .join('')
      .substring(0, 8);
    
    const range = sheet.getDataRange();
    const values = range.getValues();
    const headers = values[0];
    const idxHash = headers.indexOf('コンテンツハッシュ');
    const idxDate = headers.indexOf('送信日時');
    
    const now = new Date();
    const since = new Date(now.getTime() - days * 24 * 60 * 60 * 1000);
    
    // 過去のハッシュをチェック
    for (let i = 1; i < values.length; i++) {
      const date = new Date(values[i][idxDate]);
      if (date >= since && values[i][idxHash] === contentHash) {
        console.log(`重複コンテンツ検出: ${contentHash}`);
        return true;
      }
    }
    
    return false;
  } catch (e) {
    console.error('重複チェックエラー:', e);
    return false;
  }
}

// ==============================
// 設定管理UI
// ==============================

/**
 * スプレッドシートのカスタムメニュー
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📰 AI情報配信')
    .addItem('📋 スプレッドシート初期化', 'setupSpreadsheet')
    .addItem('⚙️ 初期設定', 'showSetupDialog')
    .addItem('🔄 トリガー設定', 'setupTriggers')
    .addSeparator()
    .addSubMenu(ui.createMenu('🧪 テスト配信')
      .addItem('🌅 朝のニュース', 'testMorningDelivery')
      .addItem('💡 ビジネスアイデア', 'testNoonDelivery')
      .addItem('🌆 夕方のニュース', 'testEveningDelivery'))
    .addSeparator()
    .addItem('📊 ステータス確認', 'checkSystemStatus')
    .addToUi();
}

/**
 * 初期設定ダイアログ
 */
function showSetupDialog() {
  const html = HtmlService.createHtmlOutputFromFile('setup')
    .setWidth(600)
    .setHeight(500);
  SpreadsheetApp.getUi()
    .showModalDialog(html, 'AI情報配信システム設定');
}

/**
 * 設定保存
 */
function saveSettings(settings) {
  const scriptProperties = PropertiesService.getScriptProperties();
  
  // APIキー保存
  if (settings.perplexityKey) {
    scriptProperties.setProperty('PERPLEXITY_API_KEY', settings.perplexityKey);
  }
  if (settings.grokKey) {
    scriptProperties.setProperty('GROK_API_KEY', settings.grokKey);
  }
  if (settings.notionKey) {
    scriptProperties.setProperty('NOTION_API_KEY', settings.notionKey);
  }
  
  // Notion Database ID保存
  if (settings.newsDbId) {
    scriptProperties.setProperty('NEWS_DATABASE_ID', settings.newsDbId);
  }
  if (settings.ideasDbId) {
    scriptProperties.setProperty('IDEAS_DATABASE_ID', settings.ideasDbId);
  }
  
  // メールアドレス保存
  if (settings.emails) {
    scriptProperties.setProperty('EMAIL_RECIPIENTS', settings.emails);
  }
  
  return '設定を保存しました';
}

/**
 * エラー通知送信
 */
function sendErrorNotification(error, timing) {
  const subject = `⚠️ AI情報配信エラー（${timing}）`;
  const body = `
配信エラーが発生しました。

タイミング: ${timing}
エラー内容: ${error.toString()}
スタックトレース: ${error.stack || 'なし'}
発生時刻: ${new Date().toLocaleDateString('ja-JP')}
`;
  
  // 設定から管理者メールを取得
  const settings = getSettings();
  const adminEmail = settings['配信先メールアドレス'] || AUTHOR_PROFILE.basic.email;
  
  try {
    MailApp.sendEmail({
      to: adminEmail,
      subject: subject,
      body: body
    });
  } catch (e) {
    console.error('エラー通知送信失敗:', e);
  }
}

/**
 * システムステータス確認
 */
function checkSystemStatus() {
  const ui = SpreadsheetApp.getUi();
  const settings = getSettings();
  const recipients = getRecipients();
  
  const triggers = ScriptApp.getProjectTriggers();
  const morningTrigger = triggers.find(t => t.getHandlerFunction() === 'deliverMorningNews');
  const noonTrigger = triggers.find(t => t.getHandlerFunction() === 'deliverBusinessIdeas');
  const eveningTrigger = triggers.find(t => t.getHandlerFunction() === 'deliverEveningNews');
  
  let status = '📊 システムステータス\n\n';
  
  status += '⚙️ 基本設定:\n';
  status += `  配信有効: ${settings['配信有効'] === 'TRUE' ? '✅' : '❌'}\n`;
  status += `  配信先数: ${recipients.length}件\n\n`;
  
  status += '🔑 API設定:\n';
  const scriptProperties = PropertiesService.getScriptProperties();
  status += `  Perplexity: ${scriptProperties.getProperty('PERPLEXITY_API_KEY') || CONFIG.PERPLEXITY_API_KEY ? '✅' : '❌'}\n`;
  status += `  Grok: ${scriptProperties.getProperty('GROK_API_KEY') || CONFIG.GROK_API_KEY ? '✅' : '❌'}\n`;
  status += `  Notion: ${scriptProperties.getProperty('NOTION_API_KEY') ? '✅' : '❌'}\n\n`;
  
  status += '📝 Notionページ:\n';
  status += `  ニュース用: ${settings['Notion News Page ID'] ? '✅' : '❌'}\n`;
  status += `  アイデア用: ${settings['Notion Ideas Page ID'] ? '✅' : '❌'}\n\n`;
  
  status += '⏰ 配信スケジュール:\n';
  status += `  朝7時: ${morningTrigger ? '✅' : '❌'} - 昨日のニュースまとめ\n`;
  status += `  昼12時: ${noonTrigger ? '✅' : '❌'} - ビジネスアイデア\n`;
  status += `  夕方18時: ${eveningTrigger ? '✅' : '❌'} - 本日の最新ニュース\n\n`;
  
  // 最近の配信状況
  const logsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.LOGS);
  if (logsSheet) {
    const lastRow = logsSheet.getLastRow();
    if (lastRow > 1) {
      const recentLogs = logsSheet.getRange(Math.max(2, lastRow - 4), 1, Math.min(5, lastRow - 1), 10).getValues();
      status += '📋 最近の配信:\n';
      recentLogs.reverse().forEach(log => {
        const date = new Date(log[0]);
        const type = log[1];
        const statusEmoji = log[8] === 'SUCCESS' ? '✅' : '❌';
        status += `  ${statusEmoji} ${date.toLocaleDateString('ja-JP')} ${date.toLocaleTimeString('ja-JP', {hour: '2-digit', minute: '2-digit'})} - ${type}\n`;
      });
    }
  }
  
  ui.alert('システムステータス', status, ui.ButtonSet.OK);
}

/**
 * 配信先追加ダイアログ
 */
function showAddRecipientDialog() {
  const ui = SpreadsheetApp.getUi();
  
  const result = ui.prompt(
    '配信先追加',
    'メールアドレスを入力してください:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (result.getSelectedButton() === ui.Button.OK) {
    const email = result.getResponseText().trim();
    
    if (email && email.includes('@')) {
      const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.RECIPIENTS);
      sheet.appendRow([email, '', '', 'TRUE', new Date(), '']);
      ui.alert('追加完了', `${email} を配信先に追加しました。`, ui.ButtonSet.OK);
    } else {
      ui.alert('エラー', '有効なメールアドレスを入力してください。', ui.ButtonSet.OK);
    }
  }
}

/**
 * 配信統計表示
 */
function showDeliveryStats() {
  const ui = SpreadsheetApp.getUi();
  const logsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.LOGS);
  
  if (!logsSheet || logsSheet.getLastRow() <= 1) {
    ui.alert('配信統計', '配信ログがありません。', ui.ButtonSet.OK);
    return;
  }
  
  const logs = logsSheet.getRange(2, 1, logsSheet.getLastRow() - 1, 10).getValues();
  
  let stats = {
    total: logs.length,
    morning: 0,
    noon: 0,
    evening: 0,
    success: 0,
    error: 0
  };
  
  logs.forEach(log => {
    const type = log[1];
    const status = log[8];
    
    if (type === 'morning_news') stats.morning++;
    if (type === 'business_ideas') stats.noon++;
    if (type === 'evening_news') stats.evening++;
    if (status === 'SUCCESS') stats.success++;
    if (status === 'ERROR') stats.error++;
  });
  
  const message = `
📊 配信統計

総配信数: ${stats.total}件

種別別:
  朝のニュース: ${stats.morning}件
  ビジネスアイデア: ${stats.noon}件
  夕方のニュース: ${stats.evening}件

ステータス:
  成功: ${stats.success}件 (${Math.round(stats.success / stats.total * 100)}%)
  エラー: ${stats.error}件 (${Math.round(stats.error / stats.total * 100)}%)
`;
  
  ui.alert('配信統計', message, ui.ButtonSet.OK);
}

/**
 * 古いログを削除
 */
function cleanOldLogs() {
  const ui = SpreadsheetApp.getUi();
  
  const result = ui.alert(
    '古いログの削除',
    '30日以上前のログを削除しますか？',
    ui.ButtonSet.YES_NO
  );
  
  if (result === ui.Button.YES) {
    const logsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.LOGS);
    if (!logsSheet || logsSheet.getLastRow() <= 1) {
      ui.alert('削除対象なし', 'ログが存在しません。', ui.ButtonSet.OK);
      return;
    }
    
    const cutoffDate = new Date();
    cutoffDate.setDate(cutoffDate.getDate() - 30);
    
    const values = logsSheet.getDataRange().getValues();
    const newValues = [values[0]]; // ヘッダー行を保持
    let deletedCount = 0;
    
    for (let i = 1; i < values.length; i++) {
      const logDate = new Date(values[i][0]);
      if (logDate >= cutoffDate) {
        newValues.push(values[i]);
      } else {
        deletedCount++;
      }
    }
    
    if (deletedCount > 0) {
      logsSheet.clear();
      logsSheet.getRange(1, 1, newValues.length, newValues[0].length).setValues(newValues);
      logsSheet.getRange(1, 1, 1, 10).setFontWeight('bold').setBackground('#EA4335').setFontColor('white');
      logsSheet.setFrozenRows(1);
      
      ui.alert('削除完了', `${deletedCount}件の古いログを削除しました。`, ui.ButtonSet.OK);
    } else {
      ui.alert('削除対象なし', '30日以上前のログはありません。', ui.ButtonSet.OK);
    }
  }
}

/**
 * 全配信テスト
 */
function testAllDeliveries() {
  const ui = SpreadsheetApp.getUi();
  
  const result = ui.alert(
    'テスト配信',
    '3つすべての配信をテストします。続行しますか？',
    ui.ButtonSet.YES_NO
  );
  
  if (result === ui.Button.YES) {
    testMorningDelivery();
    Utilities.sleep(2000);
    testNoonDelivery();
    Utilities.sleep(2000);
    testEveningDelivery();
    
    ui.alert('テスト完了', 'すべてのテスト配信が完了しました。', ui.ButtonSet.OK);
  }
}

// ==============================
// テスト関数
// ==============================

/**
 * 朝の配信テスト
 */
function testMorningDelivery() {
  deliverMorningNews();
  SpreadsheetApp.getActiveSpreadsheet().toast('朝のニューステスト配信完了', 'テスト', 3);
}

/**
 * 昼の配信テスト
 */
function testNoonDelivery() {
  deliverBusinessIdeas();
  SpreadsheetApp.getActiveSpreadsheet().toast('ビジネスアイデアテスト配信完了', 'テスト', 3);
}

/**
 * 夕方の配信テスト
 */
function testEveningDelivery() {
  deliverEveningNews();
  SpreadsheetApp.getActiveSpreadsheet().toast('夕方のニューステスト配信完了', 'テスト', 3);
}