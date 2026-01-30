/**
 * 配列または単一の宛先に対して個別送信
 */
function sendEmailIndividually(toField, subject, htmlBody) {
  try {
    const recipients = Array.isArray(toField) ? toField : [toField];
    recipients.forEach((to) => {
      if (!to) return;
      MailApp.sendEmail({
        to: to,
        subject: subject,
        htmlBody: htmlBody,
        attachments: []
      });
    });
  } catch (e) {
    console.error('個別送信エラー:', e);
  }
}
// ===== AI メルマガ配信システム =====
// 1日3回異なるコンテンツを配信するニュースレターシステム

// ===== プロフィール情報 =====
const AUTHOR_PROFILE = {
  // 基本情報
  basic: {
    name: '後藤穂高',
    nameReading: 'ごとう ほだか',
    age: 36,
    birthDate: '1988年6月11日',
    location: 'マレーシア',
    email: 'your-email@example.com',
    currentPosition: 'Good Light Inc. 代表取締役、合同会社Intelligent Beast 代表'
  },
  
  // 学歴・資格
  education: {
    graduate: '法務博士（修士）慶應義塾大学法科大学院（2010年〜2013年）',
    undergraduate: '法学士 上智大学法学部国際関係法学科（2007年〜2010年、成績優秀につき3年卒業）',
    highSchool: '私立國學院久我山高等学校（2004年〜2007年）',
    languages: ['英語：ビジネス上級・ネイティブレベル', '中国語：日常会話レベル'],
    awards: ['各社でMVP多数', '上智大学学業奨励賞（学科首席）']
  },
  
  // 職歴・経験
  career: {
    current: [
      'Good Light Inc.（2023年4月〜現在）設立、代表取締役（マレーシア・ラブアンのタックスヘイブン・カンパニー）',
      '合同会社Intelligent Beast（2018年9月〜現在）設立、代表'
    ],
    past: [
      '株式会社Intelligent Beast（2019年9月〜2020年5月）設立、代表取締役（株式会社EYS-STYLEの社内ベンチャー）',
      '株式会社EYS-STYLE（現2nd Community株式会社）取締役、執行役員（2019年10月〜2020年5月）',
      'UUUM株式会社 法務部員（2017年8月〜2018年8月）',
      '株式会社gloops 法務部員（2016年7月〜2017年7月）',
      '株式会社フリーセル（現・ブランディングテクノロジー株式会社）法務部員（2015年9月〜2016年6月）',
      '実家の家業でコンサルティング業務（2013年〜2015年8月）※母親がAOL Japan代表取締役等を歴任'
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
      'YouTube関連法務（UUUM、AnyColor他多数企業でのYouTube関連法的整理・コンプライアンス対策）',
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
  clients: [
    '株式会社バンク', 'SCデジタル株式会社（住友商事子会社）', 'AI CROSS株式会社', '株式会社ジーエルシー',
    'ユニファ株式会社', '株式会社BitStar', 'ススメル株式会社', 'プライズ株式会社',
    'ブランディングテクノロジー株式会社', '株式会社hayaoki', '株式会社ミクシィ', '株式会社カンム',
    '株式会社いちから（AnyColor）', 'ナイル株式会社', '株式会社EYS-STYLE', '株式会社エイスリー',
    '株式会社アール・アンド・エー・シー', '株式会社flaggs', '株式会社キッズライン',
    'DouYu Japan株式会社', '株式会社366', 'グリーンモンスター株式会社', '株式会社タナクロ',
    '株式会社STAGEON', 'グローバルパートナーズ株式会社（CHRO）', 'ディープレイ株式会社',
    'Digital Entertainment Asset Pte.Ltd.', 'oVice株式会社', 'ユビ電株式会社',
    'チューリンガム株式会社', 'NextMedica株式会社', 'SecureNavi株式会社', 'メディアリンク株式会社',
    '株式会社Iranoan', '株式会社SHIFT AI', '株式会社クシム', '株式会社クシムインサイト'
  ],
  
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

const CONFIG = {
  // APIキー
  // 重要：本番環境では、APIキーをPropertiesServiceに保存することを推奨
  PERPLEXITY_API_KEY: 'YOUR_PERPLEXITY_API_KEY',
  GROK_API_KEY: 'YOUR_GROK_API_KEY',
  
  // スプレッドシート設定（設定シートのみ使用）
  SHEETS: {
    SETTINGS: '設定',
    LOGS: '配信ログ',
    RECIPIENTS: '配信先'
  },
  
  // 通知設定（動的に読み込み）
  get EMAIL_RECIPIENT() {
    return getEmailRecipient();
  },
  
  // 配信時間設定
  DELIVERY_SCHEDULE: {
    MORNING: { hour: 7, type: 'yesterday_news' },    // 朝：昨日のニュース
    NOON: { hour: 12, type: 'business_ideas' },      // 昼：ビジネス&効率化アイデア
    EVENING: { hour: 18, type: 'today_news' }        // 夕方：本日のニュース
  }
};

// ===== プロフィール情報生成関数 =====

/**
 * プロフィール情報をプロンプトに組み込む形式で生成（簡潔版）
 */
function generateProfileContext() {
  return `
## 執筆者情報

### 基本プロフィール
- 執筆者：${AUTHOR_PROFILE.basic.name}
- 現職：${AUTHOR_PROFILE.basic.currentPosition}
- 専門：上場準備・法務・情報システム・海外展開

### 執筆方針
- 実務経験に基づく実践的な内容
- 丁寧語ベースで親しみやすく
- 一人称「私」で具体的な経験を交える
- 理論よりも実用性を重視

**プロフィールは自己紹介部分とコンテンツの導入で軽く触れる程度に留め、記事の主体は最新AI情報の分析と実用的アドバイスに集中してください。**
`;
}

// ===== Perplexity API連携関数 =====

/**
 * Perplexity APIを使用して最新のAI情報を収集
 */
function fetchAINewsFromPerplexity(query, date) {
  try {
    const url = 'https://api.perplexity.ai/chat/completions';
    
    const payload = {
      model: 'sonar-deep-research',
      messages: [
        {
          role: 'system',
          content: 'You are a helpful AI news researcher. Provide comprehensive, factual information about recent AI developments in Japanese.'
        },
        {
          role: 'user',
          content: `${date}の${query}について、最新の情報を日本語で詳しく教えてください。具体的な企業名、数値、発表内容を含めてください。`
        }
      ],
      temperature: 0.2,
      max_tokens: 2000,
      return_citations: true,
      search_domain_filter: ["perplexity.ai"],
      search_recency_filter: "day"
    };

    const options = {
      method: 'post',
      contentType: 'application/json',
      headers: {
        'Authorization': `Bearer ${CONFIG.PERPLEXITY_API_KEY}`,
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
 * 複数のトピックについてPerplexityから情報を収集
 */
function gatherMultipleTopicsFromPerplexity(topics, date) {
  const results = {};
  
  topics.forEach(topic => {
    const data = fetchAINewsFromPerplexity(topic, date);
    if (data) {
      results[topic] = data;
    }
    Utilities.sleep(1000); // レート制限対策
  });
  
  return results;
}

// ... existing code ...

/**
 * 昨日のAIニュースを取得（プロフィール情報を反映）
 */
function fetchYesterdayNews() {
  const yesterday = new Date();
  yesterday.setDate(yesterday.getDate() - 1);
  
  // Perplexityから最新情報を収集
  const topics = [
    'AI関連の最新ニュース',
    'OpenAI Anthropic Google Microsoft Meta などのAI関連企業の発表',
    'AI スタートアップ 資金調達',
    'AI 規制 法律 ガイドライン',
    '生成AI 新サービス リリース'
  ];
  
  const perplexityData = gatherMultipleTopicsFromPerplexity(
    topics, 
    yesterday.toLocaleDateString('ja-JP')
  );
  
  const profileContext = generateProfileContext();
  
  // Perplexityの情報を含めたプロンプト
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
  
  return callGrokAPI(prompt, 'yesterday_news');
}

/**
 * AIビジネスアイデアを生成（プロフィール情報を反映）
 */
function fetchBusinessIdeas() {
  const today = new Date();
  
  // Perplexityから最新のAI技術トレンドを収集
  const trendTopics = [
    'AI 最新技術 トレンド ビジネス応用',
    '業務効率化 AI ツール 自動化',
    'AIエージェント RAG LLM 実装事例',
    '日本企業 AI導入 成功事例'
  ];
  
  const perplexityTrends = gatherMultipleTopicsFromPerplexity(
    trendTopics,
    today.toLocaleDateString('ja-JP')
  );
  
  const profileContext = generateProfileContext();
  
  // 重複回避用: 過去2週間のビジネスアイデアテーマを取得
  const recentThemes = getRecentBusinessIdeaThemes(14);
  
  let prompt = `
${profileContext}

あなたは${AUTHOR_PROFILE.basic.name}として、実務経験豊富なビジネスコンサルタントの視点でメルマガを執筆します。
${today.toLocaleDateString('ja-JP')}時点の最新のAI技術トレンドを踏まえて、実践的なビジネスチャンス・アイデアと業務効率化ソリューションを提案してください。

## Perplexityから収集した最新AIトレンド：
${JSON.stringify(perplexityTrends, null, 2)}

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
  
  // 最大3回まで重複回避を試みる
  for (let attempt = 0; attempt < 3; attempt++) {
    const result = callGrokAPI(prompt, 'business_ideas');
    if (!result || !result.content) return result;
    const theme = parseBusinessIdeaTitleFromContent(result.content);
    if (theme && !isThemeDuplicate(theme, recentThemes)) {
      return result;
    }
    // 重複時はプロンプトに追加の指示を付けて再生成
    const avoidNote = `\n【注意】前案は「${theme || '（不明）'}」が過去配信テーマと重複しました。全く異なる新規テーマに変更し、タイトル語句も被らないように再提案してください。`;
    prompt += avoidNote;
  }
  
  // 最後に生成したものを返す（重複する可能性はあるが3回試行済み）
  return callGrokAPI(prompt, 'business_ideas');
}

/**
 * 本日の最新AIニュースを取得（プロフィール情報を反映）
 */
function fetchTodayNews() {
  const today = new Date();
  
  // Perplexityから本日の最新情報を収集
  const todayTopics = [
    'AI 本日 今日 最新ニュース 速報',
    'OpenAI Anthropic Google 本日の発表',
    'AI企業 今日の動向 株価',
    '日本 AI 本日のニュース',
    'AIサービス 新機能 アップデート 今日'
  ];
  
  const perplexityToday = gatherMultipleTopicsFromPerplexity(
    todayTopics,
    today.toLocaleDateString('ja-JP')
  );
  
  const profileContext = generateProfileContext();
  
  const prompt = `
${profileContext}

あなたは${AUTHOR_PROFILE.basic.name}として、AI・テクノロジー専門のニュースアナリストの視点でメルマガを執筆します。
本日（${today.toLocaleDateString('ja-JP')}）これまでに報じられたAI関連の最新ニュースをまとめてください。

## Perplexityから収集した本日の最新情報：
${JSON.stringify(perplexityToday, null, 2)}

上記の最新情報を参考にしながら、以下の構成で執筆してください：

## 🌆 お疲れ様です！本日のAI最新ニュース

皆さん、お疲れ様です。${AUTHOR_PROFILE.basic.name}です。
本日も一日お疲れ様でした。上場準備・法務・情シス分野での実務経験を活かし、今日一日で発表された重要なAI関連ニュースを、ビジネス実務の観点から解説します。

## 1. 🚨 本日の速報・Breaking News

本日発表された重要ニュース（法務実務の経験から見た重要度順）

## 2. 📊 本日のハイライト TOP5

実務経験から見て最も注目すべき出来事：

**1. [出来事1のタイトル]**
- 概要：
- **実務的な見解**：社内AI開発の経験から見たインパクト
- **私の体験談**：類似技術を導入した際の現場での反応と課題
- **実務への影響**：スタートアップ経営者への示唆

**2-5. [同様の形式で]**
各項目に実体験ベースの教訓や気づきを含める

## 3. 🏢 業界別の動き

### テック大手の動向
- **Google/Microsoft/Meta等の世界的なビッグテック**：
- **実務的な関連**：上場企業支援で経験した類似動向
- **私が感じた変化**：過去の案件と比較して感じる市場の変化

### AI専門企業の動向
- **OpenAI/Anthropic等**：
- **技術的分析**：社内AI開発の知見から
- **導入現場での実感**：実際にAI技術を業務に組み込んだ時の体験

### 国内企業の動向
- **日本企業の動き**：
- **上場準備への影響**：上場準備支援の経験から
- **現場で見た課題**：日本企業特有のAI導入の障壁と解決策の実例

## 4. ⚖️ 法務・規制面での注目ポイント

法務実務の経験から：
- 新たな規制動向
- コンプライアンス対応のポイント
- グレー領域の事業化への影響
- **実際の対応事例**：規制変更に直面した時の対応と学び

## 5. 🔧 業務効率化の新機会

業務自動化の実装経験から見た実用性：
- 本日発表された新技術の業務応用可能性
- 導入検討時のポイント
- SaaSツールとの統合可能性
- **現場での試行錯誤**：新技術導入時の予期せぬ問題と解決プロセス

## 6. 🌏 国際ビジネスへの影響

海外ビジネスの運営経験から：
- 海外展開への示唆
- タックスヘイブン活用との関連
- 国際的な法的スキーム構築への影響
- **リアルな海外体験**：マレーシアでの事業運営で学んだ文化的な違いと対応策

## 7. 📈 投資・資金調達動向

M&A・資金調達の経験から：
- 注目すべき資金調達案件
- 投資トレンドの分析
- 上場準備企業への示唆
- **投資家との対話で感じること**：最近の投資家の関心の変化と実感

## 8. 明日への注目ポイント

明日以降に予定されているイベントや発表（私が注目する理由含む）

---

**本日のまとめ**

今日も様々な動きがありましたが、上場準備・法務・情シスの経験から見て、特に注目すべきは...

明日朝には昨日のニュースまとめをお届けします。
ご質問は${AUTHOR_PROFILE.basic.email}までお気軽にどうぞ。

${AUTHOR_PROFILE.basic.name}

---

**重要な執筆指示:**
- 自己紹介は冒頭のみ簡潔に
- 実務経験は軽く触れる程度に留める
- 「私の経験では...」等の表現は必要最小限に
- 夕方らしい「お疲れ様」の労いの気持ちを込めた親しみやすい文体
- 実務に直結する具体的なアドバイスを重視
- 記事の主体はAI情報の分析と実用的な示唆
- 本日のAI動向を基にした総括的な視点
`;
  
  return callGrokAPI(prompt, 'today_news');
}

// ... existing code continues with the same structure ...

/**
 * Grok-4 APIを呼び出す共通関数（x.ai）
 */
function callGrokAPI(prompt, type) {
  try {
    const url = 'https://api.x.ai/v1/chat/completions';

    const payload = {
      model: 'grok-4',
      messages: [
        { role: 'system', content: 'You are a Japanese newsletter writer specializing in AI news and business productivity. Follow the provided structure exactly and keep tone polite and practical.' },
        { role: 'user', content: prompt }
      ],
      temperature: type === 'business_ideas' ? 0.9 : 0.7,
      max_tokens: 4096
    };

    const options = {
      method: 'post',
      contentType: 'application/json',
      headers: {
        Authorization: `Bearer ${CONFIG.GROK_API_KEY}`,
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(url, options);
    const status = response.getResponseCode();
    const text = response.getContentText();

    if (status < 200 || status >= 300) {
      console.error(`Grok API HTTP ${status}: ${text}`);
      return null;
    }

    const data = JSON.parse(text);

    // x.ai chat completions response shape
    // { id, object, created, model, choices: [{ message: { role, content }, ...}] }
    if (data && data.choices && data.choices[0] && data.choices[0].message && typeof data.choices[0].message.content === 'string') {
      const content = data.choices[0].message.content;
      return {
        content: content,
        type: type,
        generatedAt: new Date()
      };
    }

    console.error('Grok API: Unexpected response shape', text);
    return null;

  } catch (error) {
    console.error('Grok APIエラー:', error);
    return null;
  }
}


// ===== HTMLテンプレート生成 =====

/**
 * HTMLメールテンプレートを生成
 */
function createNewsletterHTML(newsData, timeSlot) {
  const getThemeColors = (slot) => {
    switch(slot) {
      case 'morning':
        return {
          gradient: 'linear-gradient(135deg, #667eea 0%, #764ba2 100%)',
          primary: '#667eea',
          secondary: '#764ba2',
          emoji: '🌅',
          title: 'AIニュース朝刊'
        };
              case 'noon':
        return {
          gradient: 'linear-gradient(135deg, #f093fb 0%, #f5576c 100%)',
          primary: '#f093fb',
          secondary: '#f5576c',
          emoji: '💡',
          title: 'AIビジネス&効率化アイデア'
        };
      case 'evening':
        return {
          gradient: 'linear-gradient(135deg, #4facfe 0%, #00f2fe 100%)',
          primary: '#4facfe',
          secondary: '#00f2fe',
          emoji: '🌆',
          title: 'AI最新ニュース'
        };
      default:
        return {
          gradient: 'linear-gradient(135deg, #667eea 0%, #764ba2 100%)',
          primary: '#667eea',
          secondary: '#764ba2',
          emoji: '📰',
          title: 'AIニュースレター'
        };
    }
  };
  
  const theme = getThemeColors(timeSlot);
  
  const formatContent = (content) => {
    // 改行を正規化
    let formatted = content
      .replace(/\r\n/g, '\n')
      .replace(/\r/g, '\n');
    
    // マークダウン形式をHTMLに変換
    formatted = formatted
      // 見出しの変換（h2, h3）
      .replace(/^##\s*(.+)$/gm, '<h2 style="color: #2c3e50; margin-top: 25px; margin-bottom: 15px; padding-bottom: 8px; border-bottom: 2px solid ' + theme.primary + ';">$1</h2>')
      .replace(/^###\s*(.+)$/gm, '<h3 style="color: #34495e; margin-top: 20px; margin-bottom: 10px; font-size: 18px;">$1</h3>')
      
      // 太字の変換
      .replace(/\*\*([^*]+)\*\*/g, '<strong style="color: #2c3e50;">$1</strong>')
      
      // リストアイテムの変換（先にリストアイテムを変換）
      .replace(/^[\s]*-\s+(.+)$/gm, '<li style="margin: 8px 0; line-height: 1.7;">$1</li>')
      .replace(/^[\s]*\d+\.\s+(.+)$/gm, '<li style="margin: 8px 0; line-height: 1.7;">$1</li>')
      
      // 連続するliタグをul/olで囲む
      .replace(/(<li[^>]*>.*?<\/li>\s*)+/g, function(match) {
        // 数字リストかどうかを判断
        if (match.includes('1.') || match.includes('2.') || match.includes('3.')) {
          return '<ol style="padding-left: 20px; margin: 10px 0;">' + match + '</ol>';
        } else {
          return '<ul style="padding-left: 20px; margin: 10px 0;">' + match + '</ul>';
        }
      })
      
      // パラグラフの処理（最後に実行）
      .replace(/\n\s*\n/g, '</p><p style="margin: 12px 0; line-height: 1.7;">')
      
      // 最初と最後にpタグを追加
      .replace(/^/, '<p style="margin: 12px 0; line-height: 1.7;">')
      .replace(/$/, '</p>')
      
      // 空のpタグを削除
      .replace(/<p[^>]*>\s*<\/p>/g, '')
      
      // h2, h3タグがpタグ内にある場合は修正
      .replace(/<p[^>]*>\s*(<h[23][^>]*>.*?<\/h[23]>)\s*<\/p>/g, '$1')
      
      // ul, olタグがpタグ内にある場合は修正
      .replace(/<p[^>]*>\s*(<[uo]l[^>]*>.*?<\/[uo]l>)\s*<\/p>/gs, '$1');
    
    return formatted;
  };
  
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
      border-bottom: 3px solid ${theme.primary};
    }
    h1 {
      color: #2c3e50;
      font-size: 32px;
      margin: 0;
      font-weight: 700;
    }
    .subtitle {
      color: #7f8c8d;
      font-size: 16px;
      margin-top: 10px;
    }
    .content-wrapper {
      margin: 30px 0;
    }
    .content-section {
      background: #f8f9fa;
      border-radius: 12px;
      padding: 25px;
      margin: 20px 0;
      border-left: 4px solid ${theme.primary};
    }
    .highlight-box {
      background: linear-gradient(135deg, ${theme.primary}15 0%, ${theme.secondary}15 100%);
      border-radius: 8px;
      padding: 15px;
      margin: 15px 0;
      border: 1px solid ${theme.primary}30;
    }
    .badge {
      display: inline-block;
      background: ${theme.gradient};
      color: white;
      padding: 4px 12px;
      border-radius: 20px;
      font-size: 12px;
      font-weight: 600;
      margin-right: 8px;
    }
    h2 {
      color: #2c3e50;
      margin-top: 25px;
      margin-bottom: 15px;
      padding-bottom: 8px;
      border-bottom: 2px solid ${theme.primary};
    }
    h3 {
      color: #34495e;
      margin-top: 20px;
      margin-bottom: 10px;
      font-size: 18px;
    }
    ul, ol {
      padding-left: 20px;
      margin: 10px 0;
    }
    li {
      margin: 8px 0;
      line-height: 1.7;
    }
    p {
      margin: 12px 0;
      line-height: 1.7;
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
      color: ${theme.primary};
      text-decoration: none;
    }
    .footer a:hover {
      text-decoration: underline;
    }
    .time-info {
      background: #ecf0f1;
      border-radius: 8px;
      padding: 10px 15px;
      margin: 20px 0;
      text-align: center;
      font-size: 14px;
      color: #555;
    }
    @media (max-width: 600px) {
      .container {
        padding: 25px;
      }
      h1 {
        font-size: 26px;
      }
    }
  </style>
</head>
<body>
  <div class="container">
    <div class="header">
      <h1>${theme.emoji} ${theme.title}</h1>
      <div class="subtitle">${new Date().toLocaleString('ja-JP', { 
        year: 'numeric', 
        month: 'long', 
        day: 'numeric', 
        weekday: 'long',
        hour: '2-digit',
        minute: '2-digit'
      })}</div>
    </div>
    
    <div class="time-info">
      配信時刻：${getDeliveryTimeDescription(timeSlot)}
    </div>
    
    <div class="content-wrapper">
      <div class="content-section">
        ${formatContent(newsData.content)}
      </div>
    </div>
    
    <div class="footer">
      <p><strong>配信スケジュール</strong></p>
      <p>🌅 朝7時: 昨日のニュースまとめ | ☀️ 昼12時: ビジネス&効率化アイデア | 🌆 夕方18時: 本日の最新ニュース</p>
      <p style="margin-top: 10px; font-size: 11px; color: #95a5a6;">
        © 2025 AI Newsletter System - Powered by Grok-4<br>
        このレポートは公開情報に基づいてAIによって生成されています<br>
        投資やビジネス判断の際は、複数の情報源での確認をお勧めします
      </p>
    </div>
  </div>
</body>
</html>`;
}

/**
 * 配信時刻の説明を取得
 */
function getDeliveryTimeDescription(timeSlot) {
  switch(timeSlot) {
    case 'morning':
      return '朝の配信（昨日のニュースまとめ）';
    case 'noon':
      return '昼の配信（AIビジネス&効率化アイデア）';
    case 'evening':
      return '夕方の配信（本日の最新ニュース）';
    default:
      return 'AIニュースレター';
  }
}

// ===== エラー処理 =====

/**
 * エラー通知を送信
 */
function sendErrorNotification(error, timeSlot) {
  const subject = `⚠️ ニュースレター配信エラー（${timeSlot}）`;
  const body = `
ニュースレター配信中にエラーが発生しました。

配信タイミング: ${timeSlot}
エラー内容:
${error.toString()}

スタックトレース:
${error.stack || 'なし'}

発生時刻: ${new Date().toLocaleString('ja-JP')}
`;
  
  // Use configured recipients (array or string) and fallback to author email
  let to = normalizeRecipients(CONFIG.EMAIL_RECIPIENT);
  if (!to) to = normalizeRecipients(AUTHOR_PROFILE.basic.email);
  try {
    MailApp.sendEmail({ to, subject, body });
  } catch (sendErr) {
    console.error('Failed to send error notification (sendEmail)', sendErr);
  }
}

// ===== トリガー管理 =====

/**
 * ニュース配信トリガーを設定（1日3回）
 */
function setupNewsDeliveryTriggers() {
  // 既存のトリガーを削除
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    const handler = trigger.getHandlerFunction();
    if (handler === 'sendMorningNewsletter' || 
        handler === 'sendNoonNewsletter' || 
        handler === 'sendEveningNewsletter') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  // 朝の配信（7時）
  ScriptApp.newTrigger('sendMorningNewsletter')
    .timeBased()
    .everyDays(1)
    .atHour(7)
    .create();
  
  // 昼の配信（12時）
  ScriptApp.newTrigger('sendNoonNewsletter')
    .timeBased()
    .everyDays(1)
    .atHour(12)
    .create();
  
  // 夕方の配信（18時）
  ScriptApp.newTrigger('sendEveningNewsletter')
    .timeBased()
    .everyDays(1)
    .atHour(18)
    .create();
  
  console.log('ニュース配信トリガーを設定しました（7:00, 12:00, 18:00）');
  SpreadsheetApp.getActiveSpreadsheet().toast(
    '配信スケジュールを設定しました\n朝7時: 昨日のまとめ\n昼12時: ビジネスアイデア\n夕方18時: 本日の最新',
    '設定完了',
    5
  );
}

// ===== 配信関数 =====

/**
 * 朝のニュースレターを配信
 */
function sendMorningNewsletter() {
  try {
    console.log('朝のニュースレター配信を開始...');
    const newsData = fetchYesterdayNews();
    
    if (!newsData || !newsData.content) {
      throw new Error('朝のニュースデータの取得に失敗しました');
    }
    
    const htmlContent = createNewsletterHTML(newsData, 'morning');
    const subject = `🌅 ${new Date().toLocaleDateString('ja-JP')} AIニュース朝刊 - 昨日のハイライト`;
    
    sendEmailIndividually(CONFIG.EMAIL_RECIPIENT, subject, htmlContent);
    // ログ記録
    appendDeliveryLog({
      type: 'morning',
      subject: subject,
      to: CONFIG.EMAIL_RECIPIENT,
      content: newsData.content // コンテンツ全体を渡して詳細分析
    });
    
    console.log('朝のニュースレター配信完了');
    
  } catch (error) {
    console.error('朝のニュースレター配信エラー:', error);
    try {
      sendErrorNotification(error, 'morning');
    } catch (notifyErr) {
      console.error('Notifier failed', notifyErr);
    }
  }
}

/**
 * 昼のニュースレターを配信（ビジネス&効率化アイデア）
 */
function sendNoonNewsletter() {
  try {
    console.log('昼のニュースレター配信を開始...');
    const newsData = fetchBusinessIdeas();
    
    if (!newsData || !newsData.content) {
      throw new Error('昼のニュースデータの取得に失敗しました');
    }
    
    const htmlContent = createNewsletterHTML(newsData, 'noon');
    const subject = `💡 ${new Date().toLocaleDateString('ja-JP')} AIビジネス&効率化アイデア`;
    
    sendEmailIndividually(CONFIG.EMAIL_RECIPIENT, subject, htmlContent);
    // ログ記録
    appendDeliveryLog({
      type: 'noon',
      subject: subject,
      to: CONFIG.EMAIL_RECIPIENT,
      content: newsData.content // コンテンツ全体を渡して詳細分析
    });
    
    console.log('昼のニュースレター配信完了');
    
  } catch (error) {
    console.error('昼のニュースレター配信エラー:', error);
    try {
      sendErrorNotification(error, 'noon');
    } catch (notifyErr) {
      console.error('Notifier failed', notifyErr);
    }
  }
}

/**
 * 夕方のニュースレターを配信
 */
function sendEveningNewsletter() {
  try {
    console.log('夕方のニュースレター配信を開始...');
    const newsData = fetchTodayNews();
    
    if (!newsData || !newsData.content) {
      throw new Error('夕方のニュースデータの取得に失敗しました');
    }
    
    const htmlContent = createNewsletterHTML(newsData, 'evening');
    const subject = `🌆 ${new Date().toLocaleDateString('ja-JP')} AI最新ニュース - 本日のまとめ`;
    
    sendEmailIndividually(CONFIG.EMAIL_RECIPIENT, subject, htmlContent);
    // ログ記録
    appendDeliveryLog({
      type: 'evening',
      subject: subject,
      to: CONFIG.EMAIL_RECIPIENT,
      content: newsData.content // コンテンツ全体を渡して詳細分析
    });
    
    console.log('夕方のニュースレター配信完了');
    
  } catch (error) {
    console.error('夕方のニュースレター配信エラー:', error);
    try {
      sendErrorNotification(error, 'evening');
    } catch (notifyErr) {
      console.error('Notifier failed', notifyErr);
    }
  }
}

// ===== 設定管理関数 =====

/**
 * 配信先メールアドレスを取得
 */
function getEmailRecipient() {
  try {
    // 配信先シートがあれば、購読中の全メールアドレス配列を返す
    const recipientsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.RECIPIENTS);
    if (recipientsSheet) {
      const values = recipientsSheet.getDataRange().getValues();
      const emails = [];
      for (let i = 1; i < values.length; i++) {
        const email = (values[i][0] || '').toString().trim();
        const subscribed = String(values[i][3]).toLowerCase() !== 'false';
        if (email && subscribed) emails.push(email);
      }
      if (emails.length > 0) return emails;
    }

    // フォールバック: 設定シートの単一メール
    const setSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.SETTINGS);
    if (!setSheet) {
      console.log('設定シートが見つかりません。デフォルト値を使用します。');
      return AUTHOR_PROFILE.basic.email;
    }
    const sValues = setSheet.getDataRange().getValues();
    for (let i = 1; i < sValues.length; i++) { // ヘッダーをスキップ
      if (sValues[i][0] === '配信先メールアドレス') {
        const email = sValues[i][1];
        return email || AUTHOR_PROFILE.basic.email;
      }
    }
    return AUTHOR_PROFILE.basic.email;
    
  } catch (error) {
    console.error('メールアドレス取得エラー:', error);
    return AUTHOR_PROFILE.basic.email;
  }
}

/**
 * ニュースレター設定シートを作成
 */
function createNewsletterSettingsSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(CONFIG.SHEETS.SETTINGS);
  let logSheet = spreadsheet.getSheetByName(CONFIG.SHEETS.LOGS);
  let recipientsSheet = spreadsheet.getSheetByName(CONFIG.SHEETS.RECIPIENTS);
  
  if (!sheet) {
    sheet = spreadsheet.insertSheet(CONFIG.SHEETS.SETTINGS);
  }
  if (!logSheet) {
    logSheet = spreadsheet.insertSheet(CONFIG.SHEETS.LOGS);
  }
  if (!recipientsSheet) {
    recipientsSheet = spreadsheet.insertSheet(CONFIG.SHEETS.RECIPIENTS);
  }
  
  // シートをクリア
  sheet.clear();
  
  // ヘッダーと設定項目を作成
  const settingsData = [
    ['設定項目', '値', '説明'],
    ['配信先メールアドレス', AUTHOR_PROFILE.basic.email, 'ニュースレターの配信先'],
    ['配信スケジュール', '7:00, 12:00, 18:00', '1日3回の配信時刻'],
    ['最終更新', new Date().toLocaleString('ja-JP'), 'システム最終更新日時']
  ];
  
  // データを設定
  sheet.getRange(1, 1, settingsData.length, settingsData[0].length).setValues(settingsData);
  
  // フォーマットを適用
  const headerRange = sheet.getRange(1, 1, 1, 3);
  headerRange.setFontWeight('bold')
    .setBackground('#4285F4')
    .setFontColor('white');
  
  // 列幅を調整
  sheet.setColumnWidth(1, 200);
  sheet.setColumnWidth(2, 300);
  sheet.setColumnWidth(3, 250);
  
  console.log('設定シートを作成しました');

  // 配信ログシートを初期化（ヘッダーが無ければ作成）
  const logHeader = [['送信日時', '配信種別', '件名', '配信先', 'メインテーマ', 'サブテーマ', 'キーワード', '詳細要約（300文字）', '主要トピック']];
  const logFirstRow = logSheet.getRange(1, 1, 1, logHeader[0].length).getValues()[0];
  const headerMissing = logFirstRow.some(v => v === '' || v == null) || logFirstRow.length < logHeader[0].length;
  if (headerMissing) {
    logSheet.clear();
    logSheet.getRange(1, 1, 1, logHeader[0].length).setValues(logHeader);
    logSheet.setFrozenRows(1);
    // 列幅を内容に応じて調整
    const columnWidths = [180, 100, 200, 150, 200, 180, 150, 400, 250];
    for (let c = 1; c <= logHeader[0].length; c++) {
      logSheet.setColumnWidth(c, columnWidths[c-1] || 200);
    }
  }

  // 配信先シートを初期化（ヘッダーが無ければ作成）
  const recipientsHeader = [['メールアドレス', '氏名（任意）', 'タグ（任意）', '購読中（TRUE/FALSE）']];
  const recFirst = recipientsSheet.getRange(1, 1, 1, recipientsHeader[0].length).getValues()[0];
  const recHeaderMissing = recFirst.some(v => v === '' || v == null);
  if (recHeaderMissing) {
    recipientsSheet.clear();
    recipientsSheet.getRange(1, 1, 1, recipientsHeader[0].length).setValues(recipientsHeader);
    recipientsSheet.setFrozenRows(1);
    recipientsSheet.setColumnWidth(1, 260);
    recipientsSheet.setColumnWidth(2, 160);
    recipientsSheet.setColumnWidth(3, 160);
    recipientsSheet.setColumnWidth(4, 140);
  }
}

// ===== ログ・重複回避ユーティリティ =====

/**
 * 配信ログを追記（改良版）
 */
function appendDeliveryLog(entry) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.LOGS);
    if (!sheet) return;
    
    // コンテンツから詳細情報を抽出
    const contentAnalysis = analyzeNewsletterContent(entry.content || '', entry.type);
    
    const row = [
      new Date(),
      entry.type,
      entry.subject,
      Array.isArray(entry.to) ? entry.to.join(', ') : entry.to,
      contentAnalysis.mainTheme,
      contentAnalysis.subThemes.join(', '),
      contentAnalysis.keywords.join(', '),
      contentAnalysis.detailedSummary,
      contentAnalysis.mainTopics.join(' | ')
    ];
    sheet.appendRow(row);
  } catch (e) {
    console.error('配信ログ書き込みエラー:', e);
  }
}

/**
 * 過去N日分のビジネスアイデアテーマを取得（改良版）
 */
function getRecentBusinessIdeaThemes(days) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.LOGS);
    if (!sheet) return [];
    const range = sheet.getDataRange();
    const values = range.getValues();
    const now = new Date();
    const since = new Date(now.getTime() - days * 24 * 60 * 60 * 1000);
    const headers = values[0];
    const idxDate = headers.indexOf('送信日時');
    const idxType = headers.indexOf('配信種別');
    
    // 新しいカラム構成に対応
    let idxMainTheme = headers.indexOf('メインテーマ');
    let idxSubThemes = headers.indexOf('サブテーマ');
    let idxKeywords = headers.indexOf('キーワード');
    
    // 後方互換性のため、古いカラム名もチェック
    if (idxMainTheme === -1) idxMainTheme = headers.indexOf('テーマ/タイトル');
    
    const themes = [];
    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      const sentAt = new Date(row[idxDate]);
      if (row[idxType] === 'noon' && sentAt >= since) {
        // メインテーマ、サブテーマ、キーワードを組み合わせて重複チェック
        let themeInfo = '';
        if (idxMainTheme >= 0 && row[idxMainTheme]) {
          themeInfo += String(row[idxMainTheme]).trim();
        }
        if (idxSubThemes >= 0 && row[idxSubThemes]) {
          themeInfo += ' | ' + String(row[idxSubThemes]).trim();
        }
        if (idxKeywords >= 0 && row[idxKeywords]) {
          themeInfo += ' | ' + String(row[idxKeywords]).trim();
        }
        if (themeInfo) themes.push(themeInfo);
      }
    }
    return themes.filter(Boolean);
  } catch (e) {
    console.error('過去テーマ取得エラー:', e);
    return [];
  }
}

/**
 * ニュースレターコンテンツを詳細分析
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
    // メインテーマの抽出（複数のパターンに対応）
    let mainTheme = '';
    
    // パターン1: #### 💡 [タイトル]
    let match = /####\s*💡\s*\[(.+?)\]/.exec(content);
    if (match) mainTheme = match[1].trim();
    
    // パターン2: #### 🎯 [タイトル]
    if (!mainTheme) {
      match = /####\s*🎯\s*\[(.+?)\]/.exec(content);
      if (match) mainTheme = match[1].trim();
    }
    
    // パターン3: ## パート1: 🚀 業務効率化・自動化アイデア の後の見出し
    if (!mainTheme) {
      match = /##\s*パート1.*?###\s*(.+)$/m.exec(content);
      if (match) mainTheme = match[1].replace(/^#+\s*/, '').trim();
    }
    
    // パターン4: 最初の主要な見出し
    if (!mainTheme) {
      match = /^##\s+(.+)$/m.exec(content);
      if (match) mainTheme = match[1].trim();
    }
    
    analysis.mainTheme = mainTheme || 'テーマ不明';
    
    // サブテーマの抽出（h3, h4見出し）
    const subThemeMatches = content.match(/^###\s+(.+)$/gm) || [];
    analysis.subThemes = subThemeMatches.map(h => h.replace(/^#+\s*/, '').trim()).slice(0, 3);
    
    // キーワードの抽出（太字、重要語句）
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
    const techTerms = content.match(/AI|機械学習|深層学習|LLM|GPT|RAG|自動化|効率化|DX|デジタル変革|SaaS|API|クラウド/g) || [];
    techTerms.forEach(term => keywords.add(term));
    
    analysis.keywords = Array.from(keywords).slice(0, 8);
    
    // 詳細要約の生成（300文字程度）
    let summary = '';
    
    // 冒頭部分を取得
    const intro = content.match(/^##\s+[^\n]*[\s\S]*?(?=^##|$)/m);
    if (intro) {
      summary = intro[0].replace(/^#+\s*[^\n]*\n/, '').trim();
    }
    
    // 長すぎる場合は調整
    if (summary.length > 300) {
      summary = summary.substring(0, 297) + '...';
    } else if (summary.length < 100) {
      // 短すぎる場合は補完
      const additionalContent = content.substring(0, 500).replace(/#+[^\n]*\n/g, '').trim();
      summary = additionalContent.substring(0, 297) + '...';
    }
    
    analysis.detailedSummary = summary;
    
    // 主要トピックの抽出
    const topics = [];
    if (type === 'noon') {
      if (content.includes('業務効率化') || content.includes('自動化')) topics.push('業務効率化');
      if (content.includes('ビジネスアイデア') || content.includes('新規事業')) topics.push('ビジネス戦略');
      if (content.includes('AI導入') || content.includes('システム')) topics.push('AI導入');
      if (content.includes('ROI') || content.includes('投資回収')) topics.push('ROI分析');
      if (content.includes('海外') || content.includes('国際')) topics.push('海外展開');
    } else {
      if (content.includes('OpenAI') || content.includes('Anthropic')) topics.push('AI企業動向');
      if (content.includes('資金調達') || content.includes('投資')) topics.push('資金調達');
      if (content.includes('規制') || content.includes('法律')) topics.push('規制・法務');
      if (content.includes('新サービス') || content.includes('リリース')) topics.push('新サービス');
    }
    
    analysis.mainTopics = topics.length > 0 ? topics : ['一般AI情報'];
    
  } catch (e) {
    console.error('コンテンツ分析エラー:', e);
    analysis.mainTheme = 'エラー';
    analysis.detailedSummary = content.substring(0, 300) || 'コンテンツ取得失敗';
  }
  
  return analysis;
}

/**
 * コンテンツからビジネスアイデアのタイトルを抽出（後方互換性のため保持）
 */
function extractBusinessIdeaTitle(content) {
  const analysis = analyzeNewsletterContent(content, 'noon');
  return analysis.mainTheme;
}

/**
 * コンテンツから主要タイトルを抽出
 */
function extractPrimaryTitle(content) {
  const h2 = /^##\s+(.+)$/m.exec(content || '');
  return h2 ? h2[1].trim() : '';
}

/**
 * テーマが直近テーマ配列に重複しているか
 */
function isThemeDuplicate(theme, recentThemes) {
  if (!theme) return false;
  const target = normalizeTheme(theme);
  return recentThemes.some(t => normalizeTheme(t) === target);
}

function normalizeTheme(s) {
  return String(s)
    .toLowerCase()
    .replace(/\s+/g, '')
    .replace(/[\u3000\s]/g, '') // 全角空白
    .replace(/[\uFF10-\uFF19]/g, d => String.fromCharCode(d.charCodeAt(0) - 0xFF10 + 48)); // 全角数字→半角
}

/**
 * ビジネスアイデア本文からタイトルを抽出（ログ用）
 */
function parseBusinessIdeaTitleFromContent(content) {
  return extractBusinessIdeaTitle(content);
}

/**
 * ニュースレター設定ダイアログを表示
 */
function showNewsletterSettingsDialog() {
  const ui = SpreadsheetApp.getUi();
  
  // 現在の設定を取得
  const currentEmail = getEmailRecipient();
  
  // メールアドレス設定ダイアログ
  const result = ui.prompt(
    '📧 配信設定',
    `現在の配信先: ${currentEmail}\n\n新しい配信先メールアドレスを入力してください:`,
    ui.ButtonSet.OK_CANCEL
  );
  
  if (result.getSelectedButton() === ui.Button.OK) {
    const newEmail = result.getResponseText().trim();
    
    if (newEmail && newEmail.includes('@')) {
      // 設定シートを更新
      const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.SETTINGS);
      if (sheet) {
        const values = sheet.getDataRange().getValues();
        for (let i = 1; i < values.length; i++) {
          if (values[i][0] === '配信先メールアドレス') {
            sheet.getRange(i + 1, 2).setValue(newEmail);
            sheet.getRange(i + 1, 4).setValue(new Date().toLocaleString('ja-JP')); // 更新日時も更新
            break;
          }
        }
      }
      
      ui.alert(
        '設定更新完了',
        `配信先を「${newEmail}」に更新しました。\n\n配信スケジュール:\n🌅 朝7時: 昨日のニュース\n💡 昼12時: ビジネス&効率化アイデア\n🌆 夕方18時: 本日の最新ニュース`,
        ui.ButtonSet.OK
      );
    } else {
      ui.alert('エラー', '有効なメールアドレスを入力してください。', ui.ButtonSet.OK);
    }
  }
}

// ===== テスト関数 =====

/**
 * 朝の配信をテスト
 */
function testMorningDelivery() {
  sendMorningNewsletter();
  SpreadsheetApp.getActiveSpreadsheet().toast(
    '朝のテストメール（昨日のニュース）を送信しました',
    'テスト完了',
    3
  );
}

/**
 * 昼の配信をテスト
 */
function testNoonDelivery() {
  sendNoonNewsletter();
  SpreadsheetApp.getActiveSpreadsheet().toast(
    '昼のテストメール（ビジネス&効率化アイデア）を送信しました',
    'テスト完了',
    3
  );
}

/**
 * 夕方の配信をテスト
 */
function testEveningDelivery() {
  sendEveningNewsletter();
  SpreadsheetApp.getActiveSpreadsheet().toast(
    '夕方のテストメール（本日の最新）を送信しました',
    'テスト完了',
    3
  );
}

/**
 * 全ての配信をテスト
 */
function testAllDeliveries() {
  sendMorningNewsletter();
  Utilities.sleep(2000);
  sendNoonNewsletter();
  Utilities.sleep(2000);
  sendEveningNewsletter();
  
  SpreadsheetApp.getActiveSpreadsheet().toast(
    '全てのテストメールを送信しました',
    'テスト完了',
    5
  );
}

// ===== 初期化・ユーティリティ =====

/**
 * システムを初期化
 */
function initializeNewsletterSystem() {
  // 設定シートを作成
  createNewsletterSettingsSheet();
  
  // トリガーを設定
  setupNewsDeliveryTriggers();
  
  SpreadsheetApp.getActiveSpreadsheet().toast(
    'ニュースレターシステムの初期化が完了しました\n1日3回の配信を設定しました',
    '初期化完了',
    5
  );
}

/**
 * システムステータスをチェック
 */
function checkSystemStatus() {
  const ui = SpreadsheetApp.getUi();
  
  // トリガーの状態を確認
  const triggers = ScriptApp.getProjectTriggers();
  const morningTrigger = triggers.find(t => t.getHandlerFunction() === 'sendMorningNewsletter');
  const noonTrigger = triggers.find(t => t.getHandlerFunction() === 'sendNoonNewsletter');
  const eveningTrigger = triggers.find(t => t.getHandlerFunction() === 'sendEveningNewsletter');
  
  let status = '📊 システムステータス\n\n';
  status += `📧 配信先: ${CONFIG.EMAIL_RECIPIENT}\n\n`;
  status += `  朝7時: ${morningTrigger ? '✅ 設定済み' : '❌ 未設定'} - 昨日のニュース\n`;
  status += '⏰ 配信スケジュール:\n';
  status += `  昼12時: ${noonTrigger ? '✅ 設定済み' : '❌ 未設定'} - ビジネスアイデア\n`;
  status += `  夕方18時: ${eveningTrigger ? '✅ 設定済み' : '❌ 未設定'} - 本日の最新\n\n`;
  status += `🔑 API Key: ${CONFIG.GROK_API_KEY ? '✅ 設定済み' : '❌ 未設定'}\n`;
  
  ui.alert('システムステータス', status, ui.ButtonSet.OK);
}

/**
 * カスタムメニューを作成
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📧 AIニュースレター')
    .addItem('⚙️ 配信設定', 'showNewsletterSettingsDialog')
    .addSeparator()
    .addItem('📊 システムステータス', 'checkSystemStatus')
    .addItem('🔄 システム初期化', 'initializeNewsletterSystem')
    .addSeparator()
    .addSubMenu(ui.createMenu('🧪 テスト配信')
      .addItem('🌅 朝の配信テスト', 'testMorningDelivery')
      .addItem('☀️ 昼の配信テスト（ビジネス&効率化）', 'testNoonDelivery')
      .addItem('🌆 夕方の配信テスト', 'testEveningDelivery')
      .addSeparator()
      .addItem('📬 全配信テスト', 'testAllDeliveries'))
    .addToUi();
}

// Put this once, e.g., near the top-level utilities
function normalizeRecipients(input) {
  try {
    if (input == null) return '';
    if (Array.isArray(input)) {
      return input.flat(Infinity)
        .map(v => (v == null ? '' : String(v).trim()))
        .filter(v => v && /.+@.+\..+/.test(v))
        .join(',');
    }
    const s = String(input).trim();
    if (!s) return '';
    // If already comma/semicolon separated, keep commas
    return s.replace(/;/g, ',');
  } catch (e) {
    console.error('normalizeRecipients failed', e);
    return '';
  }
}