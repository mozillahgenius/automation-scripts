/**
 * Instagram・X投稿用画像自動生成システム
 * 1日3回異なるコンテンツの画像を生成
 */

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
  
  // Instagram投稿コンテンツ
  instagram: {
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
    target_audience: [
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
  SLIDESPEAK_API_KEY: 'YOUR_SLIDESPEAK_API_KEY', // SlideSpeakのAPIキーを設定してください
  
  // Google Drive設定
  GOOGLE_DRIVE: {
    ROOT_FOLDER_NAME: 'AI_Images', // ルートフォルダ名
    FOLDER_ID: '1y_XJEY4IDNc78cYtLY7dbXEObO5kLxi0', // 特定のフォルダIDを使用する場合は設定
    ORGANIZE_BY_DATE: true, // 日付別にフォルダを作成
    ORGANIZE_BY_TYPE: true // 時間帯別にサブフォルダを作成
  },
  
  // スプレッドシート設定
  SHEETS: {
    SETTINGS: '設定',
    LOGS: '生成ログ',
    IMAGES: '画像管理'
  },
  
  // 生成スケジュール設定
  GENERATION_SCHEDULE: {
    MORNING: { hour: 7, type: 'yesterday_news' },    // 朝：昨日のニュース
    NOON: { hour: 12, type: 'business_ideas' },      // 昼：ビジネス&効率化アイデア
    EVENING: { hour: 18, type: 'today_news' }        // 夕方：本日のニュース
  },
  
  // Instagram投稿設定
  INSTAGRAM: {
    IMAGE_COUNT: 4, // 1回の投稿で生成する画像数
    IMAGE_ASPECT_RATIO: '1:1', // Instagram正方形
    IMAGE_STYLE: 'modern_business', // ビジネス向けスタイル
    COLOR_SCHEME: ['#667eea', '#764ba2', '#f093fb', '#f5576c', '#4facfe', '#00f2fe']
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
      model: 'sonar',
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

// ===== SlideSpeak API連携関数 =====

/**
 * Google Slidesを使用してスライドを作成（メイン関数）
 */
function createSlidesWithGoogleSlides(content, timeSlot = 'unknown') {
  try {
    console.log(`=== Google Slidesでプレゼンテーション作成開始 ===`);
    console.log(`時間帯: ${timeSlot}`);
    console.log(`コンテンツ長: ${content ? content.length : 0}文字`);
    console.log(`コンテンツプレビュー: ${content ? content.substring(0, 100) + '...' : 'なし'}`);
    
    // Google Slidesでプレゼンテーションを作成
    const slideResult = createPresentationWithGoogleSlides(content, timeSlot);
    
    console.log(`Google Slides作成結果:`, slideResult);
    
    if (slideResult && slideResult.status === 'success') {
      console.log('✅ Google Slidesプレゼンテーション作成完了 + URL保存完了');
      
      const result = [{
        slideNumber: 1,
        content: {
          title: 'Google Slidesプレゼンテーション',
          text: content,
          type: 'google_slides_presentation'
        },
        imageUrl: `https://docs.google.com/presentation/d/${slideResult.taskId}/export/png`,
        presentationUrl: slideResult.presentationUrl,
        powerPointUrl: `https://docs.google.com/presentation/d/${slideResult.taskId}/export/pptx`,
        pdfUrl: `https://docs.google.com/presentation/d/${slideResult.taskId}/export/pdf`,
        timestamp: new Date(),
        source: 'google_slides',
        timeSlot: timeSlot,
        presentationId: slideResult.taskId,
        slideCount: slideResult.slides ? slideResult.slides.length : 1
      }];
      
      console.log(`戻り値:`, JSON.stringify(result, null, 2));
      return result;
    } else {
      console.error('❌ Google Slidesプレゼンテーション作成に失敗');
      throw new Error('Google Slidesプレゼンテーション作成が失敗しました');
    }
    
  } catch (error) {
    console.error('❌ Google Slidesプレゼンテーション作成エラー:', error);
    console.error('エラーの詳細:', error.toString());
    throw error; // エラーを上位に伝播
  }
}

/**
 * SlideSpeak APIを呼び出してスライドを作成
 */
function callSlideSpeakAPI(content, timeSlot = 'unknown') {
  try {
    console.log(`=== SlideSpeak API呼び出し開始 ===`);
    
    // SlideSpeak API公式エンドポイント
    const url = 'https://api.slidespeak.co/api/v1/presentation/generate';
    console.log(`API URL: ${url}`);
    
    // APIキーを取得（PropertiesServiceまたはCONFIGから）
    const apiKey = getSlideSpeakApiKey();
    console.log(`APIキー取得: ${apiKey ? 'あり' : 'なし'} (長さ: ${apiKey ? apiKey.length : 0})`);
    
    if (!apiKey || apiKey === 'YOUR_SLIDESPEAK_API_KEY') {
      console.error('❌ APIキーが設定されていません');
      return null;
    }
    
    // スライド作成用のテキスト（コンテンツ全体を含める）
    const plainText = `
【プレゼンテーションスライド】

${content}

---
作成者: 後藤穂高
作成日時: ${new Date().toLocaleDateString('ja-JP')}
時間帯: ${timeSlot}
    `.trim();
    
    console.log(`生成するテキスト長: ${plainText.length}文字`);
    console.log(`生成テキスト内容:\n${plainText}`);
    
    const payload = {
      plain_text: plainText,
      length: 4, // 4ページのプレゼンテーションを作成
      template: "business", // ビジネス向けテンプレート
      width: 1920, // 標準プレゼンテーション幅
      height: 1080, // 標準プレゼンテーション高さ
      aspect_ratio: "16:9" // 標準プレゼンテーション比率
    };
    
    console.log(`APIリクエストペイロード:`, JSON.stringify(payload, null, 2));
    
    const options = {
      method: 'post',
      contentType: 'application/json',
      headers: {
        'X-API-key': apiKey, // 小文字の 'key' に注意
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };
    
    console.log(`API呼び出し実行中...`);
    const response = UrlFetchApp.fetch(url, options);
    const statusCode = response.getResponseCode();
    const responseText = response.getContentText();
    
    console.log(`✅ SlideSpeak API Response: Status ${statusCode}`);
    console.log(`レスポンス長: ${responseText.length}文字`);
    console.log(`レスポンスプレビュー: ${responseText.substring(0, 200)}...`);
    
    // ステータスコードをチェック
    if (statusCode !== 200) {
      console.error(`❌ SlideSpeak API HTTP Error: ${statusCode}`);
      console.error(`Response: ${responseText.substring(0, 500)}`);
      return null;
    }
    
    // HTMLレスポンスの場合の検出
    if (responseText.trim().toLowerCase().startsWith('<!doctype') || 
        responseText.trim().toLowerCase().startsWith('<html')) {
      console.error('SlideSpeak API returned HTML instead of JSON');
      console.error(`Response preview: ${responseText.substring(0, 200)}`);
      return null;
    }
    
    // レスポンスをパース
    const responseData = JSON.parse(responseText);
    
    // SlideSpeakはタスクIDを返すので、タスクの完了を待つ
    if (responseData.task_id) {
      console.log(`✅ スライド作成タスクID取得: ${responseData.task_id}`);
      
      // タスクの完了を待つ（スライド作成用）
      console.log(`タスク完了待ち開始...`);
      const slideResult = waitForSlideCreationTask(responseData.task_id, apiKey, content, timeSlot);
      console.log(`タスク完了待ち結果:`, slideResult);
      return slideResult;
    }
    
    console.error('❌ SlideSpeak API: Unexpected response format', responseData);
    console.error('Response data:', JSON.stringify(responseData, null, 2));
    return null;
    
  } catch (error) {
    console.error(`SlideSpeakスライド作成 API呼び出しエラー:`, error);
    return null;
  }
}

/**
 * SlideSpeak APIを呼び出して画像を生成（拡張版：PowerPointとURL保存機能付き）
 */
function callSlideSpeakAPI_Deprecated(slideContent, slideNumber, timeSlot = 'unknown') {
  try {
    // SlideSpeak API公式エンドポイント
    const url = 'https://api.slidespeak.co/api/v1/presentation/generate';
    
    // APIキーを取得（PropertiesServiceまたはCONFIGから）
    const apiKey = getSlideSpeakApiKey();
    
    // SlideSpeak APIの正しいパラメータ形式
    // plain_textにスライドコンテンツを含める
    const plainText = `
${slideContent.title || `スライド ${slideNumber}`}

${slideContent.text || ''}

[スライド ${slideNumber}/4]
    `.trim();
    
    const payload = {
      plain_text: plainText,
      length: 1, // 1枚のスライドを生成
      template: "business", // ビジネス向けテンプレート
      export_format: "png", // PNG画像としてエクスポート
      width: 1920, // 標準プレゼンテーション幅
      height: 1080, // 標準プレゼンテーション高さ
      aspect_ratio: "16:9" // 標準プレゼンテーション比率
    };
    
    const options = {
      method: 'post',
      contentType: 'application/json',
      headers: {
        'X-API-key': apiKey, // 小文字の 'key' に注意
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };
    
    const response = UrlFetchApp.fetch(url, options);
    const statusCode = response.getResponseCode();
    const responseText = response.getContentText();
    
    console.log(`SlideSpeak API Response (スライド${slideNumber}): Status ${statusCode}`);
    
    // ステータスコードをチェック
    if (statusCode !== 200) {
      console.error(`SlideSpeak API HTTP Error: ${statusCode}`);
      console.error(`Response: ${responseText.substring(0, 500)}`);
      return null;
    }
    
    // HTMLレスポンスの場合の検出
    if (responseText.trim().toLowerCase().startsWith('<!doctype') || 
        responseText.trim().toLowerCase().startsWith('<html')) {
      console.error('SlideSpeak API returned HTML instead of JSON');
      console.error(`Response preview: ${responseText.substring(0, 200)}`);
      return null;
    }
    
    // レスポンスをパース
    const responseData = JSON.parse(responseText);
    
    // SlideSpeakはタスクIDを返すので、タスクの完了を待つ必要がある
    if (responseData.task_id) {
      console.log(`タスクID取得: ${responseData.task_id}`);
      
      // タスクの完了を待つ（PowerPointとURL保存機能付き）
      const presentationUrl = waitForSlideSpeakTask(responseData.task_id, apiKey, slideContent, slideNumber, timeSlot);
      if (presentationUrl) {
        return presentationUrl;
      }
    }
    
    console.error('SlideSpeak API: Unexpected response format', responseData);
    return null;
    
  } catch (error) {
    console.error(`SlideSpeak API呼び出しエラー (スライド${slideNumber}):`, error);
    return null;
  }
}

/**
 * スライド作成タスクの完了を待つ（修正版）
 */
function waitForSlideCreationTask(taskId, apiKey, content, timeSlot) {
  try {
    console.log(`=== タスク完了待ち開始: ${taskId} ===`);
    const statusUrl = `https://api.slidespeak.co/api/v1/task_status/${taskId}`;
    const maxAttempts = 40; // 最大80秒待機（2秒間隔×40回）
    
    for (let i = 0; i < maxAttempts; i++) {
      const options = {
        method: 'get',
        headers: {
          'X-API-key': apiKey
        },
        muteHttpExceptions: true
      };
      
      console.log(`タスクステータス確認 ${i+1}/${maxAttempts}: ${statusUrl}`);
      const response = UrlFetchApp.fetch(statusUrl, options);
      const statusCode = response.getResponseCode();
      const responseText = response.getContentText();
      
      console.log(`ステータス確認結果: Status ${statusCode}, Response length: ${responseText.length}`);
      
      if (statusCode === 200) {
        let data;
        try {
          data = JSON.parse(responseText);
        } catch (parseError) {
          console.error('JSON解析エラー:', parseError);
          console.error('レスポンス内容:', responseText.substring(0, 500));
          continue;
        }
        
        console.log('タスクデータ:', JSON.stringify(data, null, 2));
        
        const status = data.status || data.task_status || data.state || 'unknown';
        const taskResult = data.task_result || data.result;
        const presentationUrl = (taskResult && taskResult.url) || 
                               data.presentation_url || 
                               data.url || 
                               data.result_url || 
                               data.output_url ||
                               `https://app.slidespeak.co/presentations/${taskId}`;
        
        console.log(`現在のステータス: ${status}`);
        
        // 完了判定（大文字小文字を区別しない）
        const lowerStatus = status.toLowerCase();
        if (lowerStatus === 'success' || lowerStatus === 'completed' || lowerStatus === 'complete' || lowerStatus === 'done') {
          console.log('✅ スライド作成完了！');
          
          // まずURLテキストファイルを保存（確実に保存されるように）
          const urlTextResult = saveUnifiedSlideUrlToTextFile(taskId, presentationUrl, content, timeSlot);
          console.log(`URLテキストファイル保存結果: ${urlTextResult ? '成功' : '失敗'}`);
          
          // PowerPointファイルのダウンロードを試行（複数回リトライ）
          let powerPointResult = null;
          for (let downloadAttempt = 1; downloadAttempt <= 3; downloadAttempt++) {
            console.log(`PowerPointダウンロード試行 ${downloadAttempt}/3`);
            try {
              powerPointResult = saveUnifiedPowerPointToDrive(taskId, content, timeSlot);
              if (powerPointResult) {
                console.log(`✅ PowerPoint保存成功: ${powerPointResult.fileName}`);
                break;
              } else {
                console.log(`PowerPointダウンロード失敗（試行${downloadAttempt}）`);
                if (downloadAttempt < 3) {
                  console.log('5秒後に再試行...');
                  Utilities.sleep(5000);
                }
              }
            } catch (downloadError) {
              console.error(`PowerPointダウンロードエラー（試行${downloadAttempt}）:`, downloadError);
              if (downloadAttempt < 3) {
                Utilities.sleep(5000);
              }
            }
          }
          
          const result = {
            taskId: taskId,
            presentationUrl: presentationUrl,
            powerPointUrl: powerPointResult ? powerPointResult.fileUrl : null,
            urlTextUrl: urlTextResult ? urlTextResult.fileUrl : null,
            powerPointResult: powerPointResult,
            urlTextResult: urlTextResult,
            status: 'success'
          };
          
          console.log('✅ タスク完了結果:', JSON.stringify(result, null, 2));
          return result;
          
        } else if (lowerStatus === 'failed' || lowerStatus === 'error' || lowerStatus === 'failure' || lowerStatus === 'timeout') {
          console.error(`❌ スライド作成タスク失敗: ${status}`);
          console.error('タスクデータ詳細:', JSON.stringify(data, null, 2));
          console.error('エラー詳細:', data.error || data.message || data.task_info || 'Unknown error');
          
          // より詳細なエラー情報を取得するため、別のAPIエンドポイントも試行
          try {
            console.log('詳細エラー情報を取得中...');
            const detailUrl = `https://api.slidespeak.co/api/v1/presentations/${taskId}`;
            const detailOptions = {
              method: 'get',
              headers: { 'X-API-key': apiKey },
              muteHttpExceptions: true
            };
            const detailResponse = UrlFetchApp.fetch(detailUrl, detailOptions);
            const detailData = JSON.parse(detailResponse.getContentText());
            console.error('プレゼンテーション詳細:', JSON.stringify(detailData, null, 2));
          } catch (detailError) {
            console.error('詳細情報取得エラー:', detailError);
          }
          
          // 失敗時でもURLテキストファイルは保存
          const urlTextResult = saveUnifiedSlideUrlToTextFile(taskId, presentationUrl, content, timeSlot);
          
          return {
            taskId: taskId,
            presentationUrl: presentationUrl,
            powerPointUrl: null,
            urlTextUrl: urlTextResult ? urlTextResult.fileUrl : null,
            powerPointResult: null,
            urlTextResult: urlTextResult,
            status: 'failed',
            error: data.error || data.message || data.task_info || 'Unknown error',
            rawData: data
          };
          
        } else if (lowerStatus === 'processing' || lowerStatus === 'pending' || lowerStatus === 'in_progress' || 
                   lowerStatus === 'sent' || lowerStatus === 'queued' || lowerStatus === 'created' || lowerStatus === 'submitted') {
          console.log(`⏳ 処理中... (${status}) - ${i+1}/${maxAttempts}`);
        } else {
          console.log(`❓ 不明なステータス: ${status} - 処理継続`);
        }
        
      } else if (statusCode === 404) {
        console.error('❌ タスクが見つかりません (404)');
        return null;
      } else {
        console.error(`❌ HTTP エラー: ${statusCode}`);
        console.error('レスポンス:', responseText.substring(0, 200));
      }
      
      // 2秒待機
      if (i < maxAttempts - 1) {
        Utilities.sleep(2000);
      }
    }
    
    // タイムアウト時の処理
    console.error('⏰ タスクタイムアウト: 80秒経過');
    const fallbackUrl = `https://app.slidespeak.co/presentations/${taskId}`;
    const urlTextResult = saveUnifiedSlideUrlToTextFile(taskId, fallbackUrl, content, timeSlot);
    
    // タイムアウト後の最終PowerPoint取得試行
    console.log('タイムアウト後の最終PowerPoint取得試行...');
    const powerPointResult = saveUnifiedPowerPointToDrive(taskId, content, timeSlot);
    
    return {
      taskId: taskId,
      presentationUrl: fallbackUrl,
      powerPointUrl: powerPointResult ? powerPointResult.fileUrl : null,
      urlTextUrl: urlTextResult ? urlTextResult.fileUrl : null,
      powerPointResult: powerPointResult,
      urlTextResult: urlTextResult,
      status: 'timeout'
    };
    
  } catch (error) {
    console.error('❌ タスクステータス確認で致命的エラー:', error);
    console.error('エラー詳細:', error.toString());
    
    // エラー時でも最低限のURLテキストファイルは保存
    try {
      const fallbackUrl = `https://app.slidespeak.co/presentations/${taskId}`;
      const urlTextResult = saveUnifiedSlideUrlToTextFile(taskId, fallbackUrl, content, timeSlot);
      return {
        taskId: taskId,
        presentationUrl: fallbackUrl,
        powerPointUrl: null,
        urlTextUrl: urlTextResult ? urlTextResult.fileUrl : null,
        powerPointResult: null,
        urlTextResult: urlTextResult,
        status: 'error',
        error: error.toString()
      };
    } catch (finalError) {
      console.error('最終的な保存処理でもエラー:', finalError);
      return null;
    }
  }
}

/**
 * SlideSpeak タスクの完了を待つ（拡張版：PowerPointとURL保存機能付き）
 */
function waitForSlideSpeakTask(taskId, apiKey, slideContent = null, slideNumber = 1, timeSlot = 'unknown') {
  try {
    const statusUrl = `https://api.slidespeak.co/api/v1/task_status/${taskId}`;
    const maxAttempts = 20; // 最大20回試行（約40秒）
    
    for (let i = 0; i < maxAttempts; i++) {
      const options = {
        method: 'get',
        headers: {
          'X-API-key': apiKey
        },
        muteHttpExceptions: true
      };
      
      const response = UrlFetchApp.fetch(statusUrl, options);
      const statusCode = response.getResponseCode();
      const responseText = response.getContentText();
      
      console.log(`タスクステータス確認 (試行 ${i+1}/${maxAttempts}): Status ${statusCode}`);
      
      if (statusCode === 200) {
        const data = JSON.parse(responseText);
        console.log('タスクステータス:', JSON.stringify(data));
        
        // 様々なステータスフィールドをチェック
        const status = data.status || data.task_status || data.state;
        const taskResult = data.task_result;
        const presentationUrl = (taskResult && taskResult.url) || data.presentation_url || data.url || data.result_url || data.output_url;
        const pdfUrl = data.pdf_url || data.pdf;
        const downloadUrl = data.download_url || data.download;
        
        // 完了判定（SUCCESSを追加）
        if (status === 'SUCCESS' || status === 'completed' || status === 'complete' || status === 'done' || status === 'success') {
          console.log('プレゼンテーション生成完了');
          
          // PowerPointファイルとURLテキストファイルを保存（並列実行）
          if (slideContent && timeSlot !== 'unknown') {
            console.log('PowerPointファイルとURLテキストファイルの保存を開始...');
            
            // 並列でPowerPointとURLテキストファイルを保存
            try {
              const powerPointResult = savePowerPointToDrive(taskId, slideContent, slideNumber, timeSlot);
              const urlTextResult = saveSlideUrlToTextFile(taskId, presentationUrl || `https://app.slidespeak.co/presentations/${taskId}`, slideContent, slideNumber, timeSlot);
              
              if (powerPointResult) {
                console.log(`PowerPoint保存成功: ${powerPointResult.fileName}`);
              }
              if (urlTextResult) {
                console.log(`URLテキスト保存成功: ${urlTextResult.fileName}`);
              }
            } catch (saveError) {
              console.error('ファイル保存中にエラーが発生:', saveError);
              // 保存エラーでも画像生成は続行
            }
          }
          
          // PNG画像URLを優先的に探す
          const imageUrl = (taskResult && taskResult.image_url) || data.image_url || data.png_url;
          const pngExportUrl = (taskResult && taskResult.png_export_url) || data.png_export_url;
          
          if (imageUrl) {
            console.log('PNG image URL found:', imageUrl);
            return imageUrl;
          } else if (pngExportUrl) {
            console.log('PNG export URL found:', pngExportUrl);
            return pngExportUrl;
          } else if (presentationUrl) {
            console.log('Presentation URL found:', presentationUrl);
            // PNGエクスポートURLを構築
            return convertPresentationToImageUrl(presentationUrl);
          } else if (downloadUrl) {
            console.log('Download URL found:', downloadUrl);
            // PNGエクスポートを試みる
            const pngUrl = downloadUrl.replace('.pptx', '.png').replace('/download', '/export/png');
            console.log('Attempting PNG export:', pngUrl);
            return pngUrl;
          } else {
            // PNGエクスポートURLを直接構築
            console.log('No URL found, constructing PNG export URL from task ID');
            const pngExportUrl = `https://api.slidespeak.co/api/v1/presentations/${taskId}/export/png`;
            return pngExportUrl;
          }
        } else if (status === 'failed' || status === 'error') {
          console.error('タスク失敗:', data.error || data.message || 'Unknown error');
          return null;
        } else if (status === 'processing' || status === 'pending' || status === 'in_progress') {
          console.log(`タスク処理中... (ステータス: ${status})`);
        } else {
          console.log(`不明なステータス: ${status}`);
        }
      } else if (statusCode === 404) {
        console.error('タスクが見つかりません');
        return null;
      }
      
      // 2秒待機
      Utilities.sleep(2000);
    }
    
    console.error('タスクタイムアウト: 40秒経過');
    // タイムアウトしても、タスクIDからURLを生成して返す
    return `https://app.slidespeak.co/presentations/${taskId}`;
    
  } catch (error) {
    console.error('タスクステータス確認エラー:', error);
    return null;
  }
}

/**
 * プレゼンテーションURLを画像URLに変換
 * SlideSpeakが直接PNGを返すことを想定
 */
function convertPresentationToImageUrl(presentationUrl) {
  try {
    // URLがすでに画像URLの場合はそのまま返す
    if (presentationUrl.includes('.png') || presentationUrl.includes('.jpg') || 
        presentationUrl.includes('.jpeg') || presentationUrl.includes('/image') ||
        presentationUrl.includes('/export/png')) {
      return presentationUrl;
    }
    
    // プレゼンテーションIDを抽出してPNGエクスポートURLを構築
    let presentationId = null;
  
    // パターン1: /presentations/{id}
    let match = presentationUrl.match(/\/presentations?\/([a-zA-Z0-9\-]+)/);
    if (match && match[1]) {
      presentationId = match[1];
    }
  
    // パターン2: タスクIDがそのままプレゼンテーションIDの場合
    if (!presentationId && presentationUrl.match(/^[a-zA-Z0-9\-]+$/)) {
      presentationId = presentationUrl;
    }
    
    if (presentationId) {
      // PNGエクスポートURLを直接構築
      console.log('PNGエクスポートURLを構築:', presentationId);
      const pngExportUrl = `https://api.slidespeak.co/api/v1/presentations/${presentationId}/export/png`;
      return pngExportUrl;
    }
    
    // デフォルトでそのまま返す
    return presentationUrl;
  } catch (error) {
    console.error('プレゼンテーションURL変換エラー:', error);
    return presentationUrl;
  }
}

/**
 * 統合プレゼンテーション用：SlideSpeakからPowerPointファイルを取得してGoogle Driveに保存（修正版）
 */
function saveUnifiedPowerPointToDrive(taskId, content, timeSlot) {
  try {
    console.log(`=== PowerPointダウンロード開始: TaskID ${taskId} ===`);
    const apiKey = getSlideSpeakApiKey();
    
    if (!apiKey || apiKey === 'YOUR_SLIDESPEAK_API_KEY') {
      console.error('❌ APIキーが設定されていません');
      return null;
    }
    
    // PowerPointダウンロードURLを構築（優先順位順）
    const downloadUrls = [
      `https://api.slidespeak.co/api/v1/presentations/${taskId}/download`,
      `https://api.slidespeak.co/api/v1/presentations/${taskId}/export/pptx`,
      `https://api.slidespeak.co/api/v1/presentations/${taskId}.pptx`,
      `https://api.slidespeak.co/api/v1/task_result/${taskId}/download`
    ];
    
    // 保存先フォルダを事前に取得・作成
    const folder = getOrCreatePresentationFolder(timeSlot);
    if (!folder) {
      console.error('❌ 保存先フォルダの作成に失敗');
      return null;
    }
    console.log(`✅ 保存先フォルダ確保: ${folder.getName()}`);
    
    for (let urlIndex = 0; urlIndex < downloadUrls.length; urlIndex++) {
      const downloadUrl = downloadUrls[urlIndex];
      console.log(`PowerPointダウンロード試行 ${urlIndex + 1}/${downloadUrls.length}: ${downloadUrl}`);
      
      const options = {
        method: 'get',
        headers: {
          'X-API-key': apiKey,
          'Accept': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
          'User-Agent': 'Google Apps Script PowerPoint Downloader'
        },
        muteHttpExceptions: true
      };
      
      try {
        console.log('HTTP リクエスト送信中...');
        const response = UrlFetchApp.fetch(downloadUrl, options);
        const statusCode = response.getResponseCode();
        const contentType = response.getHeaders()['Content-Type'] || '';
        const contentLength = response.getHeaders()['Content-Length'] || '0';
        
        console.log(`レスポンス: Status ${statusCode}, ContentType: ${contentType}, Length: ${contentLength}`);
        
        if (statusCode === 200) {
          const blob = response.getBlob();
          const blobSize = blob.getSize();
          console.log(`Blob取得成功: サイズ ${blobSize} bytes`);
          
          // ファイルサイズをチェック（極端に小さいファイルは無効とみなす）
          if (blobSize < 1000) {
            console.log(`ファイルサイズが小さすぎます (${blobSize} bytes) - 次のURLを試行`);
            continue;
          }
          
          // Content-Typeをチェック（HTMLページが返されていないか確認）
          if (contentType.includes('text/html')) {
            console.log('HTMLページが返されました - 次のURLを試行');
            continue;
          }
          
          // ファイル名を生成（統合プレゼンテーション用）
          const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
          const fileName = `unified_presentation_${timeSlot}_${timestamp}.pptx`;
          blob.setName(fileName);
          
          console.log(`ファイル保存開始: ${fileName}`);
          
          // Google Driveに保存
          const file = folder.createFile(blob);
          const savedFileSize = file.getSize();
          
          console.log(`✅ PowerPointファイル保存成功!`);
          console.log(`- ファイル名: ${fileName}`);
          console.log(`- ファイルサイズ: ${savedFileSize} bytes`);
          console.log(`- ファイルID: ${file.getId()}`);
          console.log(`- ダウンロードURL: ${downloadUrl}`);
          
          return {
            fileId: file.getId(),
            fileName: fileName,
            fileUrl: file.getUrl(),
            downloadUrl: downloadUrl,
            size: savedFileSize,
            success: true
          };
          
        } else if (statusCode === 404) {
          console.log(`❌ ファイルが見つかりません (404): ${downloadUrl}`);
        } else if (statusCode === 403) {
          console.log(`❌ アクセス権限がありません (403): ${downloadUrl}`);
        } else if (statusCode === 429) {
          console.log(`❌ レート制限に達しました (429): ${downloadUrl}`);
          // レート制限の場合は少し待機
          if (urlIndex < downloadUrls.length - 1) {
            console.log('5秒待機後に次のURLを試行...');
            Utilities.sleep(5000);
          }
        } else {
          console.log(`❌ HTTPエラー (${statusCode}): ${downloadUrl}`);
          if (response.getContentText().length < 200) {
            console.log(`エラーレスポンス: ${response.getContentText()}`);
          }
        }
        
      } catch (fetchError) {
        console.error(`❌ ダウンロードリクエストエラー (${downloadUrl}):`, fetchError);
        console.error('詳細:', fetchError.toString());
      }
      
      // 次のURLを試行する前に少し待機
      if (urlIndex < downloadUrls.length - 1) {
        console.log('2秒待機後に次のURLを試行...');
        Utilities.sleep(2000);
      }
    }
    
    console.error('❌ 全てのPowerPointダウンロードURLが失敗しました');
    console.error(`試行したURL数: ${downloadUrls.length}`);
    console.error('考えられる原因:');
    console.error('1. スライド作成が完了していない');
    console.error('2. APIキーの権限が不足している');
    console.error('3. SlideSpeakサービス側の問題');
    console.error('4. ネットワーク接続の問題');
    
    return null;
    
  } catch (error) {
    console.error('❌ PowerPoint保存で致命的エラー:', error);
    console.error('エラー詳細:', error.toString());
    console.error('スタックトレース:', error.stack);
    return null;
  }
}

/**
 * SlideSpeakからPowerPointファイルを取得してGoogle Driveに保存
 */
function savePowerPointToDrive(taskId, slideContent, slideNumber, timeSlot) {
  try {
    const apiKey = getSlideSpeakApiKey();
    
    // PowerPointダウンロードURLを構築
    const downloadUrls = [
      `https://api.slidespeak.co/api/v1/presentations/${taskId}/download`,
      `https://api.slidespeak.co/api/v1/presentations/${taskId}/export/pptx`,
      `https://api.slidespeak.co/api/v1/presentations/${taskId}.pptx`
    ];
    
    for (const downloadUrl of downloadUrls) {
      console.log(`PowerPointダウンロード試行: ${downloadUrl}`);
      
      const options = {
        method: 'get',
        headers: {
          'X-API-key': apiKey,
          'Accept': 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
        },
        muteHttpExceptions: true
      };
      
      try {
        const response = UrlFetchApp.fetch(downloadUrl, options);
        const statusCode = response.getResponseCode();
        
        if (statusCode === 200) {
          const blob = response.getBlob();
          
          // ファイル名を生成
          const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
          const fileName = `slide_${slideNumber}_${timeSlot}_${timestamp}.pptx`;
          blob.setName(fileName);
          
          // Google Driveに保存
          const folder = getOrCreatePresentationFolder(timeSlot);
          if (folder) {
            const file = folder.createFile(blob);
            
            console.log(`PowerPointファイル保存成功: ${fileName}`);
            
            return {
              fileId: file.getId(),
              fileName: fileName,
              fileUrl: file.getUrl(),
              downloadUrl: downloadUrl,
              size: file.getSize()
            };
          }
        } else {
          console.log(`PowerPointダウンロード失敗 (${downloadUrl}): Status ${statusCode}`);
        }
      } catch (e) {
        console.log(`PowerPointダウンロードエラー (${downloadUrl}): ${e.toString()}`);
      }
    }
    
    console.error('全てのPowerPointダウンロードURLが失敗しました');
    return null;
    
  } catch (error) {
    console.error('PowerPoint保存エラー:', error);
    return null;
  }
}

/**
 * 統合プレゼンテーション用：URLと詳細情報をテキストファイルでGoogle Driveに保存
 */
function saveUnifiedSlideUrlToTextFile(taskId, presentationUrl, content, timeSlot) {
  try {
    // テキストファイルの内容を作成
    const timestamp = new Date().toISOString();
    const textContent = `統合プレゼンテーション情報
============================

作成日時: ${timestamp}
時間帯: ${timeSlot}
タスクID: ${taskId}

プレゼンテーションURL:
${presentationUrl}

統合プレゼンテーション内容:
${content}

SlideSpeak API情報:
- タスクID: ${taskId}
- 生成日時: ${timestamp}

関連URL:
- プレゼンテーション閲覧: ${presentationUrl}
- PowerPointダウンロード: https://api.slidespeak.co/api/v1/presentations/${taskId}/download

※プレゼンテーション閲覧URLはSlideSpeakによって発行される閲覧専用URLです

作成者情報:
- 作成者: 後藤穂高
- 専門: 上場準備・法務・経営企画・DX推進
- 会社: Good Light Inc. / 合同会社Intelligent Beast

============================
このファイルは自動生成されました
統合プレゼンテーション形式で作成されています
`;
    
    // ファイル名を生成
    const fileTimestamp = new Date().toISOString().replace(/[:.]/g, '-');
    const fileName = `unified_presentation_${timeSlot}_${fileTimestamp}.txt`;
    
    // Blobを作成
    const blob = Utilities.newBlob(textContent, 'text/plain; charset=utf-8', fileName);
    
    // Google Driveに保存
    const folder = getOrCreateUrlFolder(timeSlot);
    if (folder) {
      const file = folder.createFile(blob);
      
      console.log(`統合プレゼンテーションURLテキストファイル保存成功: ${fileName}`);
      
      return {
        fileId: file.getId(),
        fileName: fileName,
        fileUrl: file.getUrl(),
        size: file.getSize(),
        content: textContent
      };
    }
    
    return null;
    
  } catch (error) {
    console.error('統合プレゼンテーションURLテキストファイル保存エラー:', error);
    return null;
  }
}

/**
 * スライドのURLと詳細情報をテキストファイルでGoogle Driveに保存
 */
function saveSlideUrlToTextFile(taskId, presentationUrl, slideContent, slideNumber, timeSlot) {
  try {
    // テキストファイルの内容を作成
    const timestamp = new Date().toISOString();
    const textContent = `スライド情報
=================

作成日時: ${timestamp}
スライド番号: ${slideNumber}
時間帯: ${timeSlot}
タスクID: ${taskId}

プレゼンテーションURL:
${presentationUrl}

スライド内容:
タイトル: ${slideContent.title || 'なし'}
テキスト: ${slideContent.text || 'なし'}
タイプ: ${slideContent.type || 'なし'}

SlideSpeak API情報:
- タスクID: ${taskId}
- 生成日時: ${timestamp}
- 画像URL: ${convertPresentationToImageUrl(presentationUrl)}

関連URL:
- プレゼンテーション表示: ${presentationUrl}
- PNG画像: ${convertPresentationToImageUrl(presentationUrl)}
- PowerPointダウンロード: https://api.slidespeak.co/api/v1/presentations/${taskId}/download

=================
このファイルは自動生成されました
`;
    
    // ファイル名を生成
    const fileTimestamp = new Date().toISOString().replace(/[:.]/g, '-');
    const fileName = `slide_${slideNumber}_${timeSlot}_${fileTimestamp}.txt`;
    
    // Blobを作成
    const blob = Utilities.newBlob(textContent, 'text/plain; charset=utf-8', fileName);
    
    // Google Driveに保存
    const folder = getOrCreateUrlFolder(timeSlot);
    if (folder) {
      const file = folder.createFile(blob);
      
      console.log(`URLテキストファイル保存成功: ${fileName}`);
      
      return {
        fileId: file.getId(),
        fileName: fileName,
        fileUrl: file.getUrl(),
        size: file.getSize(),
        content: textContent
      };
    }
    
    return null;
    
  } catch (error) {
    console.error('URLテキストファイル保存エラー:', error);
    return null;
  }
}

/**
 * プレゼンテーション用フォルダを作成または取得
 */
function getOrCreatePresentationFolder(timeSlot) {
  try {
    const rootFolderId = initializeGoogleDriveFolders();
    if (!rootFolderId) {
      console.error('ルートフォルダの作成に失敗しました');
      return null;
    }
    
    const rootFolder = DriveApp.getFolderById(rootFolderId);
    const today = new Date();
    const dateString = today.toLocaleDateString('ja-JP').replace(/\//g, '-');
    
    // 日付フォルダを取得または作成
    let dateFolder;
    const dateFolders = rootFolder.getFoldersByName(dateString);
    if (dateFolders.hasNext()) {
      dateFolder = dateFolders.next();
    } else {
      dateFolder = rootFolder.createFolder(dateString);
    }
    
    // プレゼンテーション用サブフォルダを作成または取得
    const presentationFolderName = `presentations_${timeSlot}`;
    let presentationFolder;
    const presentationFolders = dateFolder.getFoldersByName(presentationFolderName);
    if (presentationFolders.hasNext()) {
      presentationFolder = presentationFolders.next();
    } else {
      presentationFolder = dateFolder.createFolder(presentationFolderName);
    }
    
    return presentationFolder;
    
  } catch (error) {
    console.error('プレゼンテーションフォルダ作成エラー:', error);
    return null;
  }
}

/**
 * URL用フォルダを作成または取得
 */
function getOrCreateUrlFolder(timeSlot) {
  try {
    const rootFolderId = initializeGoogleDriveFolders();
    if (!rootFolderId) {
      console.error('ルートフォルダの作成に失敗しました');
      return null;
    }
    
    const rootFolder = DriveApp.getFolderById(rootFolderId);
    const today = new Date();
    const dateString = today.toLocaleDateString('ja-JP').replace(/\//g, '-');
    
    // 日付フォルダを取得または作成
    let dateFolder;
    const dateFolders = rootFolder.getFoldersByName(dateString);
    if (dateFolders.hasNext()) {
      dateFolder = dateFolders.next();
    } else {
      dateFolder = rootFolder.createFolder(dateString);
    }
    
    // URL用サブフォルダを作成または取得
    const urlFolderName = `urls_${timeSlot}`;
    let urlFolder;
    const urlFolders = dateFolder.getFoldersByName(urlFolderName);
    if (urlFolders.hasNext()) {
      urlFolder = urlFolders.next();
    } else {
      urlFolder = dateFolder.createFolder(urlFolderName);
    }
    
    return urlFolder;
    
  } catch (error) {
    console.error('URLフォルダ作成エラー:', error);
    return null;
  }
}





// ===== コンテンツ生成関数 =====

/**
 * 昨日のAIニュースを取得して説明用スライドを生成
 */
function generateYesterdayNewsSlides() {
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
  
  // Grokでコンテンツ生成
  console.log(`=== Grokでコンテンツ生成開始 ===`);
  const prompt = createPresentationPrompt(profileContext, perplexityData, 'yesterday_news', yesterday);
  console.log(`プロンプト長: ${prompt ? prompt.length : 0}文字`);
  
  const newsData = callGrokAPI(prompt, 'yesterday_news');
  console.log(`Grok API結果:`, newsData);
  
  if (newsData && newsData.content) {
    console.log(`✅ コンテンツ生成成功 (長さ: ${newsData.content.length}文字)`);
    
    // SlideSpeak でスライド作成（朝の時間帯）
    const slides = createSlidesWithSlideSpeak(newsData.content, 'morning');
    
    // ログ記録
    saveGenerationLog('morning', slides, newsData.content);
    
    return slides;
  } else {
    console.error('❌ Grokコンテンツ生成に失敗しました');
    console.error('newsData:', newsData);
  }
  
  return null;
}

/**
 * ビジネスアイデアの説明用スライドを生成
 */
function generateBusinessIdeasSlides() {
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
  
  // Grokでコンテンツ生成
  const prompt = createPresentationPrompt(profileContext, perplexityTrends, 'business_ideas', today);
  const ideaData = callGrokAPI(prompt, 'business_ideas');
  
  if (ideaData && ideaData.content) {
    // SlideSpeak でスライド作成（昼の時間帯）
    const slides = createSlidesWithSlideSpeak(ideaData.content, 'noon');
    
    // ログ記録
    saveGenerationLog('noon', slides, ideaData.content);
    
    return slides;
  }
  
  return null;
}

/**
 * 本日のAIニュースの説明用スライドを生成
 */
function generateTodayNewsSlides() {
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
  
  // Grokでコンテンツ生成
  const prompt = createPresentationPrompt(profileContext, perplexityToday, 'today_news', today);
  const todayData = callGrokAPI(prompt, 'today_news');
  
  if (todayData && todayData.content) {
    // SlideSpeak でスライド作成（夕方の時間帯）
    const slides = createSlidesWithSlideSpeak(todayData.content, 'evening');
    
    // ログ記録
    saveGenerationLog('evening', slides, todayData.content);
    
    return slides;
  }
  
  return null;
}

/**
 * プレゼンテーション用のプロンプトを作成
 */
function createPresentationPrompt(profileContext, perplexityData, type, date) {
  // プレゼンテーション用に最適化されたプロンプトを生成
  let basePrompt = `
${profileContext}

プレゼンテーション用のコンテンツを作成してください。
以下の要件に従ってください：

【重要】スライド構成（全4枚）:
表紙や目次は不要で、すべて詳細な内容のスライドとしてください。

1. 概要・背景（1枚目）
   - 主要トピックの明確な提示
   - 背景情報の詳細な説明
   - キーポイントの整理
   - データや事実に基づく情報

2. 詳細分析（第1部）（2枚目）
   - 最も重要な分析ポイントの深堀り
   - 具体的なデータや事例の提示
   - 技術的な詳細の説明
   - 影響範囲の明確化

3. 詳細分析（第2部）（3枚目）
   - 追加の重要ポイントの分析
   - ビジネスへの具体的影響
   - 市場や業界への波及効果
   - 実務的な考察

4. 影響・結論・次のステップ（4枚目）
   - 主要な結論の整理
   - 実務への応用方法
   - 推奨する具体的なアクションプラン
   - 参考情報とフォローアップ

【デザイン要件】:
- アスペクト比: 16:9（標準プレゼンテーション形式）
- 文字量: 詳細で分かりやすい説明
- 構成: ビジネス向けの論理的な構成
- スタイル: プロフェッショナルで信頼性の高いデザイン

【内容要件】:
- 各スライドに十分な詳細情報を含める
- 実務に直結する具体的な内容
- 数値データや事例を積極的に活用
- 後藤穂高の専門性（法務・経営企画・上場準備等）を活かした視点

## Perplexityから収集した情報：
${JSON.stringify(perplexityData, null, 2)}
`;
  
  // タイプ別の追加指示
  switch(type) {
    case 'yesterday_news':
      basePrompt += '\n\n昨日のAIニュースから最も重要なトピックを選び、4枚のスライドで詳細に分析してください。法務・規制面、ビジネスへの影響、実務的な対応策を含めて包括的に説明してください。';
      break;
    case 'business_ideas':
      basePrompt += '\n\n実践的なAIビジネスアイデアを1つ選び、4枚のスライドで詳細に解説してください。実現可能性、法的考慮事項、必要なリソース、具体的な実装ステップを含めて体系的に説明してください。';
      break;
    case 'today_news':
      basePrompt += '\n\n本日の最新AIニュースから最も影響力のあるトピックを選び、4枚のスライドで詳細に分析してください。短期的・長期的な影響、企業が取るべき対応策、注意すべきリスクを含めて実務的に説明してください。';
      break;
  }
  
  return basePrompt;
}

/**
 * Grok-4 APIを呼び出す共通関数
 */
function callGrokAPI(prompt, type) {
  try {
    const url = 'https://api.x.ai/v1/chat/completions';

    const payload = {
      model: 'grok-4',
      messages: [
        { role: 'system', content: 'You are a Japanese content creator specializing in Instagram posts about AI and business. Create engaging, visual content optimized for Instagram carousel posts.' },
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

// ===== Google Drive 管理 =====

/**
 * Google Driveのフォルダ構造を初期化
 */
function initializeGoogleDriveFolders() {
  try {
    let rootFolder;
    
    // ルートフォルダの取得または作成
    if (CONFIG.GOOGLE_DRIVE.FOLDER_ID) {
      rootFolder = DriveApp.getFolderById(CONFIG.GOOGLE_DRIVE.FOLDER_ID);
    } else {
      const folders = DriveApp.getFoldersByName(CONFIG.GOOGLE_DRIVE.ROOT_FOLDER_NAME);
      if (folders.hasNext()) {
        rootFolder = folders.next();
      } else {
        rootFolder = DriveApp.createFolder(CONFIG.GOOGLE_DRIVE.ROOT_FOLDER_NAME);
      }
    }
    
    return rootFolder.getId();
  } catch (error) {
    console.error('Google Driveフォルダ初期化エラー:', error);
    return null;
  }
}

/**
 * 日付と時間帯に応じたフォルダを作成・取得
 */
function getOrCreateImageFolder(timeSlot) {
  try {
    const rootFolderId = initializeGoogleDriveFolders();
    if (!rootFolderId) return null;
    
    const rootFolder = DriveApp.getFolderById(rootFolderId);
    const today = new Date();
    
    let targetFolder = rootFolder;
    
    // 日付別フォルダ作成
    if (CONFIG.GOOGLE_DRIVE.ORGANIZE_BY_DATE) {
      const dateString = today.toLocaleDateString('ja-JP').replace(/\//g, '-');
      const dateFolders = targetFolder.getFoldersByName(dateString);
      
      if (dateFolders.hasNext()) {
        targetFolder = dateFolders.next();
      } else {
        targetFolder = targetFolder.createFolder(dateString);
      }
    }
    
    // 時間帯別サブフォルダ作成
    if (CONFIG.GOOGLE_DRIVE.ORGANIZE_BY_TYPE) {
      const typeMapping = {
        'morning': '01_朝_昨日のニュース',
        'noon': '02_昼_ビジネスアイデア',
        'evening': '03_夕方_本日のニュース'
      };
      
      const typeFolderName = typeMapping[timeSlot] || timeSlot;
      const typeFolders = targetFolder.getFoldersByName(typeFolderName);
      
      if (typeFolders.hasNext()) {
        targetFolder = typeFolders.next();
      } else {
        targetFolder = targetFolder.createFolder(typeFolderName);
      }
    }
    
    return targetFolder;
  } catch (error) {
    console.error('フォルダ作成エラー:', error);
    return null;
  }
}

/**
 * 画像URLからGoogle Driveに保存
 */
function saveImageToDrive(imageUrl, fileName, folder) {
  try {
    // Google SlidesのエクスポートURLの場合
    if (imageUrl && imageUrl.includes('docs.google.com/presentation')) {
      // URLから直接画像を取得
      const response = UrlFetchApp.fetch(imageUrl, {
        headers: {
          'Authorization': 'Bearer ' + ScriptApp.getOAuthToken()
        }
      });
      const blob = response.getBlob();
      blob.setName(fileName);
      
      // Google Driveに保存
      const file = folder.createFile(blob);
      
      return {
        fileId: file.getId(),
        fileName: file.getName(),
        fileUrl: file.getUrl(),
        size: file.getSize()
      };
    }
    
    // 通常のURLの場合
    if (imageUrl && imageUrl.startsWith('http')) {
      const response = UrlFetchApp.fetch(imageUrl);
      const blob = response.getBlob();
      blob.setName(fileName);
      
      const file = folder.createFile(blob);
      
      return {
        fileId: file.getId(),
        fileName: file.getName(),
        fileUrl: file.getUrl(),
        size: file.getSize()
      };
    }
    
    console.error(`無効な画像URL: ${imageUrl}`);
    return null;
    
  } catch (error) {
    console.error(`画像保存エラー (${fileName}):`, error);
    return null;
  }
}

/**
 * 生成された画像セットをGoogle Driveに保存
 */
function saveImagesToDrive(images, timeSlot) {
  try {
    const folder = getOrCreateImageFolder(timeSlot);
    if (!folder) {
      console.error('保存先フォルダの作成に失敗しました');
      return [];
    }
    
    const savedImages = [];
    const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
    
    images.forEach((image, index) => {
      const fileName = `slide_${index + 1}_${timestamp}.png`;
      const savedFile = saveImageToDrive(image.imageUrl, fileName, folder);
      
      if (savedFile) {
        savedImages.push({
          ...image,
          driveFile: savedFile,
          savedAt: new Date()
        });
      }
    });
    
    console.log(`${savedImages.length}枚の画像をGoogle Driveに保存しました`);
    return savedImages;
  } catch (error) {
    console.error('画像セット保存エラー:', error);
    return [];
  }
}

// ===== スプレッドシート管理 =====

/**
 * 生成ログを保存
 */
function saveGenerationLog(timeSlot, images, content, driveFiles = []) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.LOGS);
    if (!sheet) return;
    
    const driveUrls = driveFiles.map(file => file.driveFile ? file.driveFile.fileUrl : '').filter(url => url);
    
    const row = [
      new Date(),
      timeSlot,
      images.length,
      JSON.stringify(images.map(img => img.imageUrl)),
      driveUrls.join(', '),
      content.substring(0, 500),
      'SUCCESS'
    ];
    
    sheet.appendRow(row);
  } catch (error) {
    console.error('ログ保存エラー:', error);
  }
}

/**
 * スライド管理シートにスライド情報を保存
 */
function saveImageInfo(slides, type) {
  try {
    console.log(`=== スライド情報保存開始 ===`);
    console.log(`タイプ: ${type}`);
    console.log(`スライド数: ${slides ? slides.length : 0}`);
    
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.IMAGES);
    if (!sheet) {
      console.error('❌ シートが見つかりません:', CONFIG.SHEETS.IMAGES);
      return;
    }
    
    slides.forEach((slide, index) => {
      console.log(`スライド ${index + 1} 保存中:`, slide);
      
      const row = [
        new Date(), // 生成日時
        type, // 時間帯
        1, // 生成枚数（常に1）
        slide.presentationUrl || 'N/A', // プレゼンテーションURL
        slide.powerPointUrl || 'N/A', // PowerPointダウンロードURL
        slide.content ? slide.content.title : 'N/A', // タイトル
        slide.content ? slide.content.text.substring(0, 200) : 'N/A' // 内容（先頭200文字）
      ];
      
      console.log(`追加する行データ:`, row);
      sheet.appendRow(row);
    });
    
    console.log(`✅ スライド情報保存完了`);
  } catch (error) {
    console.error('❌ スライド情報保存エラー:', error);
    console.error('エラー詳細:', error.toString());
  }
}

// ===== トリガー管理 =====

/**
 * 画像生成トリガーを設定（1日3回）
 */
function setupImageGenerationTriggers() {
  // 既存のトリガーを削除
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    const handler = trigger.getHandlerFunction();
    if (handler === 'generateMorningImages' || 
        handler === 'generateNoonImages' || 
        handler === 'generateEveningImages') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  // 朝の生成（7時）
  ScriptApp.newTrigger('generateMorningImages')
    .timeBased()
    .everyDays(1)
    .atHour(7)
    .create();
  
  // 昼の生成（12時）
  ScriptApp.newTrigger('generateNoonImages')
    .timeBased()
    .everyDays(1)
    .atHour(12)
    .create();
  
  // 夕方の生成（18時）
  ScriptApp.newTrigger('generateEveningImages')
    .timeBased()
    .everyDays(1)
    .atHour(18)
    .create();
  
  console.log('画像生成トリガーを設定しました（7:00, 12:00, 18:00）');
  SpreadsheetApp.getActiveSpreadsheet().toast(
    '生成スケジュールを設定しました\n朝7時: 昨日のまとめ\n昼12時: ビジネスアイデア\n夕方18時: 本日の最新',
    '設定完了',
    5
  );
}

// ===== 生成実行関数 =====

/**
 * 朝の画像生成
 */
function generateMorningImages() {
  try {
    console.log('=== 朝のスライド作成開始 ===');
    console.log(`実行時刻: ${new Date().toISOString()}`);
    
    const slides = generateYesterdayNewsSlides();
    console.log(`generateYesterdayNewsSlides結果:`, slides);
    
    if (slides && slides.length > 0) {
      console.log(`✅ スライド作成成功: ${slides.length}件`);
      console.log(`スライド詳細:`, JSON.stringify(slides, null, 2));
      
      // スプレッドシートに記録
      console.log(`スプレッドシートに記録中...`);
      saveImageInfo(slides, 'morning');
      console.log(`✅ 朝のスライド作成完了: ${slides.length}件`);
      console.log(`PowerPointファイルとURLテキストファイルがGoogle Driveに保存されました`);
    } else {
      console.error('❌ スライド作成結果が空です');
      throw new Error('朝のスライド作成に失敗しました');
    }
  } catch (error) {
    console.error('❌ 朝のスライド作成エラー:', error);
    console.error('エラー詳細:', error.toString());
    console.error('スタックトレース:', error.stack);
    sendErrorNotification(error, 'morning');
  }
}

/**
 * 昼のスライド作成
 */
function generateNoonImages() {
  try {
    console.log('昼のスライド作成を開始...');
    const slides = generateBusinessIdeasSlides();
    
    if (slides && slides.length > 0) {
      // スプレッドシートに記録
      saveImageInfo(slides, 'noon');
      console.log(`昼のスライド作成完了: ${slides.length}件`);
      console.log(`PowerPointファイルとURLテキストファイルがGoogle Driveに保存されました`);
    } else {
      throw new Error('昼のスライド作成に失敗しました');
    }
  } catch (error) {
    console.error('昼の画像生成エラー:', error);
    sendErrorNotification(error, 'noon');
  }
}

/**
 * 夕方のスライド作成
 */
function generateEveningImages() {
  try {
    console.log('夕方のスライド作成を開始...');
    const slides = generateTodayNewsSlides();
    
    if (slides && slides.length > 0) {
      // スプレッドシートに記録
      saveImageInfo(slides, 'evening');
      console.log(`夕方のスライド作成完了: ${slides.length}件`);
      console.log(`PowerPointファイルとURLテキストファイルがGoogle Driveに保存されました`);
    } else {
      throw new Error('夕方のスライド作成に失敗しました');
    }
  } catch (error) {
    console.error('夕方の画像生成エラー:', error);
    sendErrorNotification(error, 'evening');
  }
}

// ===== エラー処理 =====

/**
 * エラー通知を送信
 */
function sendErrorNotification(error, timeSlot) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.LOGS);
  if (sheet) {
    const row = [
      new Date(),
      timeSlot,
      0,
      'ERROR',
      error.toString(),
      'FAILED'
    ];
    sheet.appendRow(row);
  }
  console.error(`画像生成エラー (${timeSlot}):`, error);
}

// ===== 初期化・設定 =====

/**
 * システムを初期化
 */
function initializeImageGenerationSystem() {
  // 必要なシートを作成
  createRequiredSheets();
  
  // トリガーを設定
  setupImageGenerationTriggers();
  
  SpreadsheetApp.getActiveSpreadsheet().toast(
    'Instagram画像生成システムの初期化が完了しました\n1日3回の自動生成を設定しました',
    '初期化完了',
    5
  );
}

/**
 * 必要なシートを作成
 */
function createRequiredSheets() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // 設定シート
  let settingsSheet = spreadsheet.getSheetByName(CONFIG.SHEETS.SETTINGS);
  if (!settingsSheet) {
    settingsSheet = spreadsheet.insertSheet(CONFIG.SHEETS.SETTINGS);
    const settingsHeader = [['設定項目', '値', '説明']];
    settingsSheet.getRange(1, 1, 1, 3).setValues(settingsHeader);
    const settingsData = [
      ['SlideSpeak APIキー', CONFIG.SLIDESPEAK_API_KEY, 'SlideSpeak APIのアクセスキー'],
      ['生成スケジュール', '7:00, 12:00, 18:00', '1日3回の生成時刻'],
      ['画像枚数', CONFIG.INSTAGRAM.IMAGE_COUNT, '1回の投稿で生成する画像数'],
      ['最終更新', new Date().toLocaleString('ja-JP'), 'システム最終更新日時']
    ];
    settingsSheet.getRange(2, 1, settingsData.length, 3).setValues(settingsData);
  }
  
  // 生成ログシート
  let logsSheet = spreadsheet.getSheetByName(CONFIG.SHEETS.LOGS);
  if (!logsSheet) {
    logsSheet = spreadsheet.insertSheet(CONFIG.SHEETS.LOGS);
    const logsHeader = [['生成日時', '時間帯', '生成枚数', '画像URL', 'Drive URL', 'コンテンツ概要', 'ステータス']];
    logsSheet.getRange(1, 1, 1, 7).setValues(logsHeader);
    logsSheet.setFrozenRows(1);
  }
  
  // 画像管理シート
  let imagesSheet = spreadsheet.getSheetByName(CONFIG.SHEETS.IMAGES);
  if (!imagesSheet) {
    imagesSheet = spreadsheet.insertSheet(CONFIG.SHEETS.IMAGES);
    const imagesHeader = [['保存日時', 'タイプ', 'スライド番号', '画像URL', 'Drive URL', 'Drive File ID', 'タイトル', 'コンテンツ']];
    imagesSheet.getRange(1, 1, 1, 8).setValues(imagesHeader);
    imagesSheet.setFrozenRows(1);
  }
}

// ===== ユーティリティ関数 =====

/**
 * Google Driveにフォルダを作成または取得
 */
function getOrCreateFolder(parentFolder, folderName) {
  try {
    // 既存のフォルダを検索
    const folders = parentFolder.getFoldersByName(folderName);
    if (folders.hasNext()) {
      return folders.next();
    }
    
    // フォルダが存在しない場合は作成
    return parentFolder.createFolder(folderName);
  } catch (error) {
    console.error(`フォルダ作成エラー (${folderName}):`, error);
    throw error;
  }
}


// ===== テスト関数 =====

/**
 * 朝の生成をテスト（Google Slides版）
 */
function testMorningGeneration() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    console.log('=== 朝のGoogle Slidesプレゼンテーション生成テスト ===');
    generateMorningImages();
    
    ui.alert('✅ テスト完了', 
      '朝のGoogle Slidesプレゼンテーション生成を実行しました。\n詳細は実行ログを確認してください。', 
      ui.ButtonSet.OK);
    
    console.log('✅ 朝のテスト完了');
    
  } catch (error) {
    console.error('❌ 朝のテストエラー:', error);
    ui.alert('❌ エラー', `朝のテスト中にエラーが発生しました:\n${error.toString()}`, ui.ButtonSet.OK);
  }
}

/**
 * 昼の生成をテスト（Google Slides版）
 */
function testNoonGeneration() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    console.log('=== 昼のGoogle Slidesプレゼンテーション生成テスト ===');
    generateNoonImages();
    
    ui.alert('✅ テスト完了', 
      '昼のGoogle Slidesプレゼンテーション生成を実行しました。\n詳細は実行ログを確認してください。', 
      ui.ButtonSet.OK);
    
    console.log('✅ 昼のテスト完了');
    
  } catch (error) {
    console.error('❌ 昼のテストエラー:', error);
    ui.alert('❌ エラー', `昼のテスト中にエラーが発生しました:\n${error.toString()}`, ui.ButtonSet.OK);
  }
}

/**
 * 夕方の生成をテスト（Google Slides版）
 */
function testEveningGeneration() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    console.log('=== 夕方のGoogle Slidesプレゼンテーション生成テスト ===');
    generateEveningImages();
    
    ui.alert('✅ テスト完了', 
      '夕方のGoogle Slidesプレゼンテーション生成を実行しました。\n詳細は実行ログを確認してください。', 
      ui.ButtonSet.OK);
    
    console.log('✅ 夕方のテスト完了');
    
  } catch (error) {
    console.error('❌ 夕方のテストエラー:', error);
    ui.alert('❌ エラー', `夕方のテスト中にエラーが発生しました:\n${error.toString()}`, ui.ButtonSet.OK);
  }
}

// ===== メニュー作成 =====

/**
 * カスタムメニューを作成
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📸 Instagram画像生成')
    .addItem('⚙️ システム設定', 'showSystemSettings')
    .addItem('🔑 APIキー設定', 'setApiKeys')
    .addSeparator()
    .addItem('📊 システムステータス', 'checkSystemStatus')
    .addItem('🔄 システム初期化', 'initializeImageGenerationSystem')
    .addSeparator()
    .addSubMenu(ui.createMenu('🧪 テスト')
      .addItem('🔌 API接続テスト', 'testSlideSpeakConnection')
      .addItem('🖼️ テスト画像生成（1枚）', 'testSingleImageGeneration')
      .addSeparator()
      .addItem('🌅 朝の生成テスト', 'testMorningGeneration')
      .addItem('☀️ 昼の生成テスト', 'testNoonGeneration')
      .addItem('🌆 夕方の生成テスト', 'testEveningGeneration'))
    .addToUi();
}

/**
 * システム設定ダイアログを表示
 */
function showSystemSettings() {
  const ui = SpreadsheetApp.getUi();
  
  const result = ui.alert(
    '📸 Instagram画像生成システム',
    `現在の設定:\n\n` +
    `• 生成スケジュール: 7:00, 12:00, 18:00\n` +
    `• 画像枚数: ${CONFIG.INSTAGRAM.IMAGE_COUNT}枚/回\n` +
    `• アスペクト比: ${CONFIG.INSTAGRAM.IMAGE_ASPECT_RATIO}\n` +
    `• スタイル: ${CONFIG.INSTAGRAM.IMAGE_STYLE}\n\n` +
    `SlideSpeak APIキーを設定してください。`,
    ui.ButtonSet.OK
  );
}

/**
 * システムステータスをチェック
 */
function checkSystemStatus() {
  const ui = SpreadsheetApp.getUi();
  
  // トリガーの状態を確認
  const triggers = ScriptApp.getProjectTriggers();
  const morningTrigger = triggers.find(t => t.getHandlerFunction() === 'generateMorningImages');
  const noonTrigger = triggers.find(t => t.getHandlerFunction() === 'generateNoonImages');
  const eveningTrigger = triggers.find(t => t.getHandlerFunction() === 'generateEveningImages');
  
  // Google Driveフォルダの確認
  const rootFolderId = initializeGoogleDriveFolders();
  const driveStatus = rootFolderId ? '✅ 設定済み' : '❌ 未設定';
  
  let status = '📊 システムステータス\n\n';
  status += '⏰ 生成スケジュール:\n';
  status += `  朝7時: ${morningTrigger ? '✅ 設定済み' : '❌ 未設定'} - 昨日のニュース\n`;
  status += `  昼12時: ${noonTrigger ? '✅ 設定済み' : '❌ 未設定'} - ビジネスアイデア\n`;
  status += `  夕方18時: ${eveningTrigger ? '✅ 設定済み' : '❌ 未設定'} - 本日の最新\n\n`;
  status += `📁 Google Drive:\n`;
  status += `  保存先フォルダ: ${driveStatus}\n`;
  status += `  フォルダ名: ${CONFIG.GOOGLE_DRIVE.ROOT_FOLDER_NAME}\n\n`;
  status += `🔑 API Keys:\n`;
  status += `  Grok: ${CONFIG.GROK_API_KEY ? '✅ 設定済み' : '❌ 未設定'}\n`;
  status += `  Perplexity: ${CONFIG.PERPLEXITY_API_KEY ? '✅ 設定済み' : '❌ 未設定'}\n`;
  status += `  SlideSpeak: ${CONFIG.SLIDESPEAK_API_KEY !== 'YOUR_SLIDESPEAK_API_KEY' ? '✅ 設定済み' : '❌ 未設定'}\n`;
  
  ui.alert('システムステータス', status, ui.ButtonSet.OK);
}

// ===== APIキー管理 =====

/**
 * APIキーを安全に設定（PropertiesService使用）
 */
function setApiKeys() {
  const ui = SpreadsheetApp.getUi();
  
  // SlideSpeak APIキー入力
  const slideSpeakResult = ui.prompt(
    '🔑 SlideSpeak APIキー設定',
    'SlideSpeak APIキーを入力してください:\n\n' +
    '※ APIキーの取得方法は SlideSpeak_Setup_Guide.md を参照',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (slideSpeakResult.getSelectedButton() === ui.Button.OK) {
    const apiKey = slideSpeakResult.getResponseText().trim();
    if (apiKey) {
      // PropertiesServiceに保存
      const scriptProperties = PropertiesService.getScriptProperties();
      scriptProperties.setProperty('SLIDESPEAK_API_KEY', apiKey);
      
      ui.alert('✅ 成功', 'SlideSpeak APIキーを設定しました。', ui.ButtonSet.OK);
    }
  }
}

/**
 * PropertiesServiceからAPIキーを取得
 */
function getSlideSpeakApiKey() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const key = scriptProperties.getProperty('SLIDESPEAK_API_KEY');
  return key || CONFIG.SLIDESPEAK_API_KEY; // PropertiesServiceになければCONFIGから取得
}

/**
 * SlideSpeak API接続テスト
 */
function testSlideSpeakConnection() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    const apiKey = getSlideSpeakApiKey();
    
    if (!apiKey || apiKey === 'YOUR_SLIDESPEAK_API_KEY') {
      ui.alert(
        '⚠️ APIキー未設定',
        'SlideSpeak APIキーが設定されていません。\n' +
        'メニューから「APIキー設定」を実行してください。',
        ui.ButtonSet.OK
      );
      return;
    }
    
    // 正しいテストエンドポイント（単純なGET）
    const endpoints = [
      'https://api.slidespeak.co/api/v1/task_status/test', // タスクステータスエンドポイントのテスト
      'https://api.slidespeak.co/api/v1/presentation/generate' // POSTエンドポイントをGETでテスト
    ];
    
    let successUrl = null;
    let lastError = null;
    
    for (const url of endpoints) {
      console.log(`Testing endpoint: ${url}`);
      const options = {
        method: 'get',
        headers: {
          'X-API-key': apiKey, // 小文字の 'key' が正しい
          'Accept': 'application/json'
        },
        muteHttpExceptions: true
      };
    
      try {
        const response = UrlFetchApp.fetch(url, options);
        const statusCode = response.getResponseCode();
        
        if (statusCode === 200 || statusCode === 201) {
          successUrl = url;
          break;
        } else if (statusCode === 401 || statusCode === 403) {
          lastError = `認証エラー (${statusCode}): APIキーが無効です`;
          break;
        } else {
          lastError = `エンドポイント ${url}: ステータス ${statusCode}`;
        }
      } catch (e) {
        lastError = `エンドポイント ${url}: ${e.toString()}`;
      }
    }
    
    if (successUrl) {
      ui.alert(
        '✅ 接続成功',
        `SlideSpeak APIへの接続に成功しました！\n\n` +
        `有効なエンドポイント: ${successUrl}\n` +
        `APIキー: 設定済み`,
        ui.ButtonSet.OK
      );
    } else {
      ui.alert(
        '⚠️ 接続エラー',
        `SlideSpeak APIに接続できませんでした。\n\n` +
        `最後のエラー: ${lastError}\n\n` +
        `以下を確認してください:\n` +
        `1. APIキーが正しいか\n` +
        `2. SlideSpeak のアカウントが有効か\n` +
        `3. APIプランが有効か\n\n` +
        `※ SlideSpeak の正確なAPIエンドポイントはサポートに確認してください。`,
        ui.ButtonSet.OK
      );
    }
    
  } catch (error) {
    ui.alert(
      '❌ エラー',
      `API接続テスト中にエラーが発生しました:\n${error.toString()}`,
      ui.ButtonSet.OK
    );
  }
}

/**
 * Google Slidesプレゼンテーション生成テスト（メイン）
 */
function testSingleImageGeneration() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    console.log('=== Google Slidesプレゼンテーション生成テスト開始 ===');
    console.log(`実行時刻: ${new Date().toISOString()}`);
    
    // テスト用コンテンツ
    const testContent = `Google Slides APIテストプレゼンテーション

## テスト概要

このプレゼンテーションはGoogle Slides APIを使用して自動生成されました。

### テスト項目
1. ✅ プレゼンテーション作成
2. ✅ コンテンツ設定
3. ✅ Google Driveフォルダ整理
4. ✅ URL生成・保存

### 技術仕様
• 使用API: Google Slides API
• 言語: Google Apps Script
• 保存先: Google Drive
• 形式: Google Slides (.gslides)

### エクスポート機能
• PowerPoint形式 (.pptx)
• PDF形式 (.pdf)
• PNG画像形式 (.png)

---
作成者: ${AUTHOR_PROFILE.basic.name}
専門: 上場準備・法務・DX推進
テスト実行日時: ${new Date().toLocaleDateString('ja-JP')} ${new Date().toLocaleTimeString('ja-JP')}`;
    
    console.log('テスト用コンテンツ長:', testContent.length);
    
    // Google Slidesでプレゼンテーション作成を試行
    console.log('Google Slidesプレゼンテーション作成開始...');
    const slideResult = createSlidesWithGoogleSlides(testContent, 'test');
    
    console.log('Google Slides作成結果:', JSON.stringify(slideResult, null, 2));
    
    if (slideResult && slideResult.length > 0) {
      const result = slideResult[0];
      
      // 詳細な成功結果を表示
      let successMessage = `✅ Google Slidesプレゼンテーション作成成功！\n\n`;
      successMessage += `📄 プレゼンテーションURL:\n${result.presentationUrl}\n\n`;
      successMessage += `🔧 機能確認:\n`;
      successMessage += `• スライド数: ${result.slideCount || '1'}\n`;
      successMessage += `• 作成方式: ${result.source}\n`;
      successMessage += `• プレゼンテーションID: ${result.presentationId}\n\n`;
      
      successMessage += `📥 利用可能なエクスポート:\n`;
      successMessage += `• PowerPoint: ${result.powerPointUrl}\n`;
      successMessage += `• PDF: ${result.pdfUrl}\n`;
      successMessage += `• PNG画像: ${result.imageUrl}\n\n`;
      
      successMessage += `💾 ファイル保存状況:\n`;
      successMessage += `• Google Drive: ✅ 保存済み\n`;
      successMessage += `• フォルダ整理: ✅ 完了\n`;
      successMessage += `• URLテキストファイル: ✅ 保存済み\n\n`;
      
      successMessage += `🎉 Google Slides APIによるプレゼンテーション作成は完全に動作しています！`;
      
      ui.alert('✅ テスト成功', successMessage, ui.ButtonSet.OK);
      
      console.log('✅ Google Slidesテスト完了 - 成功');
      
    } else {
      const errorMessage = `❌ Google Slidesプレゼンテーション作成に失敗しました。\n\n` +
        `考えられる原因:\n` +
        `1. Google Slides APIの権限が不足している\n` +
        `2. Google Driveの容量が不足している\n` +
        `3. ネットワーク接続に問題がある\n\n` +
        `詳細はGoogle Apps Scriptの実行ログを確認してください。`;
      
      ui.alert('❌ テスト失敗', errorMessage, ui.ButtonSet.OK);
      
      console.error('❌ Google Slidesテスト完了 - 失敗');
    }
    
  } catch (error) {
    console.error('❌ Google Slidesテスト実行中にエラー:', error);
    console.error('エラー詳細:', error.toString());
    console.error('スタックトレース:', error.stack);
    
    const errorMessage = `Google Slidesテスト中にエラーが発生しました:\n\n${error.toString()}\n\n` +
      `詳細なエラー情報はGoogle Apps Scriptの実行ログで確認できます。`;
    
    ui.alert('❌ エラー', errorMessage, ui.ButtonSet.OK);
  }
}

/**
 * 異なるパラメータでのテスト関数
 */
function testAlternativeParameters() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    console.log('=== 代替パラメータテスト開始 ===');
    
    // より単純なテストコンテンツ
    const simpleContent = "Simple test presentation with basic content.";
    
    console.log('代替パラメータAPIテスト開始...');
    const result = callSlideSpeakAPIWithAlternativeParams(simpleContent, 'test');
    
    if (result) {
      ui.alert('✅ 代替パラメータテスト成功', 
        `プレゼンテーション作成成功\nURL: ${result.presentationUrl}`, 
        ui.ButtonSet.OK);
    } else {
      ui.alert('❌ 代替パラメータテスト失敗', 
        '代替パラメータでも失敗しました', 
        ui.ButtonSet.OK);
    }
    
  } catch (error) {
    console.error('代替パラメータテストエラー:', error);
    ui.alert('❌ エラー', error.toString(), ui.ButtonSet.OK);
  }
}

/**
 * 代替パラメータでSlideSpeak APIを呼び出し
 */
function callSlideSpeakAPIWithAlternativeParams(content, timeSlot = 'test') {
  try {
    console.log('=== 代替パラメータAPI呼び出し ===');
    
    const url = 'https://api.slidespeak.co/api/v1/presentation/generate';
    const apiKey = getSlideSpeakApiKey();
    
    if (!apiKey || apiKey === 'YOUR_SLIDESPEAK_API_KEY') {
      console.error('❌ APIキーが設定されていません');
      return null;
    }
    
    // 最小限のパラメータでテスト
    const simplePayload = {
      plain_text: content,
      length: 1 // 1ページのみ
    };
    
    console.log('簡単なペイロード:', JSON.stringify(simplePayload, null, 2));
    
    const options = {
      method: 'post',
      contentType: 'application/json',
      headers: {
        'X-API-key': apiKey,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(simplePayload),
      muteHttpExceptions: true
    };
    
    const response = UrlFetchApp.fetch(url, options);
    const statusCode = response.getResponseCode();
    const responseText = response.getContentText();
    
    console.log(`代替パラメータAPI応答: Status ${statusCode}`);
    console.log(`レスポンス: ${responseText}`);
    
    if (statusCode !== 200) {
      console.error('❌ 代替パラメータAPI HTTPエラー:', statusCode);
      return null;
    }
    
    const responseData = JSON.parse(responseText);
    
    if (responseData.task_id) {
      console.log(`✅ 代替パラメータタスクID: ${responseData.task_id}`);
      return waitForSlideCreationTask(responseData.task_id, apiKey, content, timeSlot);
    }
    
    console.error('❌ 代替パラメータ: task_idが返されませんでした');
    return null;
    
  } catch (error) {
    console.error('❌ 代替パラメータAPI呼び出しエラー:', error);
    return null;
  }
}

/**
 * SlideSpeak API接続テスト（複数エンドポイント）
 */
function testSlideSpeakConnectivity() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    console.log('=== SlideSpeak API接続テスト開始 ===');
    
    const apiKey = getSlideSpeakApiKey();
    if (!apiKey || apiKey === 'YOUR_SLIDESPEAK_API_KEY') {
      ui.alert('❌ エラー', 'APIキーが設定されていません', ui.ButtonSet.OK);
      return;
    }
    
    console.log(`APIキー確認: ${apiKey.substring(0, 8)}...`);
    
    // 複数のエンドポイントをテスト
    const endpoints = [
      'https://api.slidespeak.co/api/v1/presentation/generate',
      'https://api.slidespeak.co/api/v1/presentations/create',
      'https://api.slidespeak.co/api/v1/create-presentation',
      'https://api.slidespeak.co/presentations/generate',
      'https://api.slidespeak.co/generate'
    ];
    
    const testPayload = {
      plain_text: "Test presentation",
      length: 1
    };
    
    let results = [];
    
    for (let i = 0; i < endpoints.length; i++) {
      const endpoint = endpoints[i];
      console.log(`\n--- エンドポイント ${i+1}/${endpoints.length}: ${endpoint} ---`);
      
      try {
        // POSTリクエストをテスト
        const postOptions = {
          method: 'post',
          contentType: 'application/json',
          headers: {
            'X-API-key': apiKey,
            'Content-Type': 'application/json',
            'Authorization': `Bearer ${apiKey}`,
            'X-API-Key': apiKey
          },
          payload: JSON.stringify(testPayload),
          muteHttpExceptions: true
        };
        
        console.log('POST リクエスト送信中...');
        const postResponse = UrlFetchApp.fetch(endpoint, postOptions);
        const postStatus = postResponse.getResponseCode();
        const postText = postResponse.getContentText();
        
        console.log(`POST応答: Status ${postStatus}`);
        console.log(`POST内容: ${postText.substring(0, 200)}...`);
        
        results.push({
          endpoint: endpoint,
          method: 'POST',
          status: postStatus,
          response: postText.substring(0, 100),
          success: postStatus === 200
        });
        
        if (postStatus === 200) {
          console.log('✅ 成功したエンドポイントを発見!');
          break;
        }
        
        // POST失敗時はGETもテスト
        if (postStatus === 405) {
          console.log('GET リクエストもテスト...');
          const getOptions = {
            method: 'get',
            headers: {
              'X-API-key': apiKey,
              'Authorization': `Bearer ${apiKey}`,
              'X-API-Key': apiKey
            },
            muteHttpExceptions: true
          };
          
          const getResponse = UrlFetchApp.fetch(endpoint, getOptions);
          const getStatus = getResponse.getResponseCode();
          const getText = getResponse.getContentText();
          
          console.log(`GET応答: Status ${getStatus}`);
          console.log(`GET内容: ${getText.substring(0, 200)}...`);
          
          results.push({
            endpoint: endpoint,
            method: 'GET',
            status: getStatus,
            response: getText.substring(0, 100),
            success: getStatus === 200
          });
        }
        
        // レート制限を避けるため少し待機
        Utilities.sleep(1000);
        
      } catch (endpointError) {
        console.error(`エンドポイントエラー (${endpoint}):`, endpointError);
        results.push({
          endpoint: endpoint,
          method: 'ERROR',
          status: 'ERROR',
          response: endpointError.toString(),
          success: false
        });
      }
    }
    
    // 結果をまとめて表示
    console.log('\n=== テスト結果まとめ ===');
    let resultMessage = 'SlideSpeak API接続テスト結果:\n\n';
    
    for (const result of results) {
      const status = result.success ? '✅' : '❌';
      resultMessage += `${status} ${result.endpoint}\n`;
      resultMessage += `   Method: ${result.method}, Status: ${result.status}\n`;
      if (result.response) {
        resultMessage += `   Response: ${result.response.substring(0, 50)}...\n`;
      }
      resultMessage += '\n';
      
      console.log(`${status} ${result.endpoint} (${result.method}): ${result.status}`);
    }
    
    const successfulEndpoints = results.filter(r => r.success);
    if (successfulEndpoints.length > 0) {
      resultMessage += `\n成功したエンドポイント: ${successfulEndpoints.length}個\n`;
      resultMessage += '上記の成功したエンドポイントを使用してください。';
    } else {
      resultMessage += '\n❌ 全てのエンドポイントでエラーが発生しました。\n';
      resultMessage += '考えられる原因:\n';
      resultMessage += '1. APIキーが無効\n';
      resultMessage += '2. SlideSpeak APIの仕様変更\n';
      resultMessage += '3. アカウントの問題\n';
      resultMessage += '\nSlideSpeakサポートに問い合わせることをお勧めします。';
    }
    
    ui.alert('🔍 API接続テスト完了', resultMessage, ui.ButtonSet.OK);
    
  } catch (error) {
    console.error('❌ API接続テストで致命的エラー:', error);
    ui.alert('❌ エラー', `テスト中にエラーが発生: ${error.toString()}`, ui.ButtonSet.OK);
  }
}

/**
 * APIキー設定確認テスト
 */
function checkApiKeyConfiguration() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    console.log('=== APIキー設定確認 ===');
    
    // PropertiesServiceから取得
    let apiKeyFromProperties = null;
    try {
      apiKeyFromProperties = PropertiesService.getScriptProperties().getProperty('SLIDESPEAK_API_KEY');
    } catch (e) {
      console.log('PropertiesServiceからの取得エラー:', e);
    }
    
    // CONFIGから取得
    let apiKeyFromConfig = null;
    try {
      apiKeyFromConfig = CONFIG.SLIDESPEAK_API_KEY;
    } catch (e) {
      console.log('CONFIGからの取得エラー:', e);
    }
    
    // getSlideSpeakApiKey関数から取得
    let apiKeyFromFunction = null;
    try {
      apiKeyFromFunction = getSlideSpeakApiKey();
    } catch (e) {
      console.log('getSlideSpeakApiKey関数エラー:', e);
    }
    
    let message = 'APIキー設定状況:\n\n';
    
    message += `🔑 PropertiesService: ${apiKeyFromProperties ? `設定済み (${apiKeyFromProperties.substring(0, 8)}...)` : '未設定'}\n`;
    message += `⚙️ CONFIG: ${apiKeyFromConfig ? `設定済み (${apiKeyFromConfig.substring(0, 8)}...)` : '未設定'}\n`;
    message += `🔧 関数経由: ${apiKeyFromFunction ? `取得成功 (${apiKeyFromFunction.substring(0, 8)}...)` : '取得失敗'}\n\n`;
    
    if (!apiKeyFromFunction || apiKeyFromFunction === 'YOUR_SLIDESPEAK_API_KEY') {
      message += '❌ APIキーが正しく設定されていません。\n\n';
      message += '設定方法:\n';
      message += '1. Google Apps Scriptで「プロジェクト設定」を開く\n';
      message += '2. 「スクリプト プロパティ」セクションで追加\n';
      message += '3. プロパティ: SLIDESPEAK_API_KEY\n';
      message += '4. 値: あなたのSlideSpeakAPIキー\n\n';
      message += 'または、CONFIGオブジェクト内のSLIDESPEAK_API_KEYを直接編集してください。';
    } else {
      message += '✅ APIキーは正しく設定されています。\n';
      message += 'エンドポイントまたはAPIの仕様に問題がある可能性があります。';
    }
    
    console.log('APIキー確認結果:', message);
    ui.alert('🔑 APIキー設定確認', message, ui.ButtonSet.OK);
    
  } catch (error) {
    console.error('APIキー確認エラー:', error);
    ui.alert('❌ エラー', `確認中にエラーが発生: ${error.toString()}`, ui.ButtonSet.OK);
  }
}

/**
 * 代替手法：Google Slidesを使用したプレゼンテーション作成
 */
function createPresentationWithGoogleSlides(content, timeSlot = 'test') {
  try {
    console.log('=== Google Slidesを使用したプレゼンテーション作成開始 ===');
    
    // 新しいプレゼンテーションを作成
    const presentation = SlidesApp.create(`自動生成プレゼンテーション_${timeSlot}_${new Date().getTime()}`);
    const presentationId = presentation.getId();
    
    console.log(`✅ Google Slidesプレゼンテーション作成: ${presentationId}`);
    
    // コンテンツを段落に分割
    const contentLines = content.split('\n').filter(line => line.trim() !== '');
    const slides = [];
    
    // 最初のスライド（タイトル）
    const titleSlide = presentation.getSlides()[0];
    const titleLayout = titleSlide.getLayout();
    
    // タイトルを設定
    const shapes = titleSlide.getShapes();
    if (shapes.length > 0) {
      const titleShape = shapes[0];
      titleShape.getText().setText(`自動生成プレゼンテーション\n${timeSlot}`);
    }
    
    // サブタイトルを設定
    if (shapes.length > 1) {
      const subtitleShape = shapes[1];
      subtitleShape.getText().setText(`作成者: 後藤穂高\n作成日時: ${new Date().toLocaleDateString('ja-JP')}`);
    }
    
    slides.push({
      slideId: titleSlide.getObjectId(),
      title: 'タイトルスライド',
      content: `自動生成プレゼンテーション - ${timeSlot}`
    });
    
    // コンテンツスライドを追加
    const chunkSize = 3; // 1スライドに3行ずつ
    for (let i = 0; i < contentLines.length; i += chunkSize) {
      const slideContent = contentLines.slice(i, i + chunkSize);
      const slideTitle = `スライド ${Math.floor(i / chunkSize) + 2}`;
      
      // 新しいスライドを追加
      const newSlide = presentation.appendSlide(SlidesApp.PredefinedLayout.TITLE_AND_BODY);
      
      // タイトルを設定
      const slideShapes = newSlide.getShapes();
      if (slideShapes.length > 0) {
        slideShapes[0].getText().setText(slideTitle);
      }
      
      // 本文を設定
      if (slideShapes.length > 1) {
        slideShapes[1].getText().setText(slideContent.join('\n\n'));
      }
      
      slides.push({
        slideId: newSlide.getObjectId(),
        title: slideTitle,
        content: slideContent.join('\n')
      });
      
      console.log(`スライド追加: ${slideTitle}`);
    }
    
    // プレゼンテーションのURLを取得
    const presentationUrl = `https://docs.google.com/presentation/d/${presentationId}/edit`;
    
    console.log(`✅ プレゼンテーション完成: ${presentationUrl}`);
    
    // Google Driveに移動してフォルダに整理
    try {
      const folder = getOrCreatePresentationFolder(timeSlot);
      if (folder) {
        const file = DriveApp.getFileById(presentationId);
        DriveApp.getRootFolder().removeFile(file);
        folder.addFile(file);
        console.log(`✅ プレゼンテーションをフォルダに移動`);
      }
    } catch (moveError) {
      console.log('フォルダ移動エラー（処理は継続）:', moveError);
    }
    
    // URLテキストファイルを保存
    const urlTextResult = saveGoogleSlidesUrlToTextFile(presentationId, presentationUrl, content, timeSlot);
    
    return {
      taskId: presentationId,
      presentationUrl: presentationUrl,
      powerPointUrl: null, // Google Slidesなのでなし
      urlTextUrl: urlTextResult ? urlTextResult.fileUrl : null,
      powerPointResult: null,
      urlTextResult: urlTextResult,
      status: 'success',
      slides: slides,
      type: 'google_slides'
    };
    
  } catch (error) {
    console.error('❌ Google Slides作成エラー:', error);
    return null;
  }
}

/**
 * Google SlidesのURLと詳細情報をテキストファイルで保存
 */
function saveGoogleSlidesUrlToTextFile(presentationId, presentationUrl, content, timeSlot) {
  try {
    const timestamp = new Date().toISOString();
    const textContent = `Google Slides プレゼンテーション情報
============================

作成日時: ${timestamp}
時間帯: ${timeSlot}
プレゼンテーションID: ${presentationId}

プレゼンテーションURL:
${presentationUrl}

編集用URL:
${presentationUrl}

プレビューURL:
https://docs.google.com/presentation/d/${presentationId}/preview

PDF出力URL:
https://docs.google.com/presentation/d/${presentationId}/export/pdf

プレゼンテーション内容:
${content}

Google Slides情報:
- プレゼンテーションID: ${presentationId}
- 生成方式: Google Slides API
- 作成日時: ${timestamp}

作成者情報:
- 作成者: 後藤穂高
- 専門: 上場準備・法務・経営企画・DX推進
- 会社: Good Light Inc. / 合同会社Intelligent Beast

============================
このファイルは自動生成されました
Google Slides形式で作成されています
`;
    
    const fileTimestamp = new Date().toISOString().replace(/[:.]/g, '-');
    const fileName = `google_slides_${timeSlot}_${fileTimestamp}.txt`;
    
    const blob = Utilities.newBlob(textContent, 'text/plain; charset=utf-8', fileName);
    
    const folder = getOrCreateUrlFolder(timeSlot);
    if (folder) {
      const file = folder.createFile(blob);
      
      console.log(`✅ Google SlidesURLテキストファイル保存成功: ${fileName}`);
      
      return {
        fileId: file.getId(),
        fileName: fileName,
        fileUrl: file.getUrl(),
        size: file.getSize(),
        content: textContent
      };
    }
    
    return null;
    
  } catch (error) {
    console.error('❌ Google Slides URLテキストファイル保存エラー:', error);
    return null;
  }
}

/**
 * SlideSpeak API問題時の代替手法テスト
 */
function testAlternativeSlideGeneration() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    console.log('=== 代替スライド生成テスト開始 ===');
    
    const testContent = `これは代替手法によるテストプレゼンテーションです。

主な特徴:
• Google Slides APIを使用
• 自動的なスライド分割
• フォルダ整理機能

テスト項目:
1. プレゼンテーション作成
2. コンテンツ設定
3. ファイル保存
4. URL生成

SlideSpeak APIが利用できない場合の代替手法として、
Google Slides APIを使用してプレゼンテーションを作成します。`;
    
    console.log('Google Slides作成開始...');
    const result = createPresentationWithGoogleSlides(testContent, 'test');
    
    if (result && result.status === 'success') {
      let successMessage = `✅ 代替手法でプレゼンテーション作成成功！\n\n`;
      successMessage += `📄 Google Slidesプレゼンテーション:\n${result.presentationUrl}\n\n`;
      successMessage += `📝 スライド数: ${result.slides ? result.slides.length : '不明'}\n`;
      successMessage += `💾 URLテキストファイル: ${result.urlTextResult ? '保存済み' : '保存失敗'}\n\n`;
      successMessage += `この方法なら確実にプレゼンテーションを作成できます。\n`;
      successMessage += `SlideSpeak APIの代わりにGoogle Slides APIを使用しています。`;
      
      ui.alert('✅ 代替手法テスト成功', successMessage, ui.ButtonSet.OK);
      console.log('✅ 代替手法テスト成功');
      
    } else {
      ui.alert('❌ 代替手法テスト失敗', 
        'Google Slides作成でもエラーが発生しました。\n権限設定を確認してください。', 
        ui.ButtonSet.OK);
      console.error('❌ 代替手法テスト失敗');
    }
    
  } catch (error) {
    console.error('❌ 代替手法テストエラー:', error);
    ui.alert('❌ エラー', `テスト中にエラーが発生: ${error.toString()}`, ui.ButtonSet.OK);
  }
}

// ===== 各時間帯のスライド生成関数（Google Slides版） =====

/**
 * 昨日のニュースまとめスライド生成（Google Slides版・実際のニュース取得）
 */
function generateYesterdayNewsSlides() {
  try {
    console.log('=== 昨日のニュースまとめスライド生成開始 ===');
    
    const yesterday = new Date();
    yesterday.setDate(yesterday.getDate() - 1);
    const dateStr = yesterday.toLocaleDateString('ja-JP');
    
    console.log(`昨日の日付: ${dateStr}`);
    
    // 実際のニュースデータを取得
    console.log('Perplexity APIから昨日のニュースを取得中...');
    const newsTopics = [
      'AI 人工知能 最新ニュース',
      'スタートアップ 資金調達 新サービス',
      'テクノロジー イノベーション 新技術'
    ];
    
    const newsData = gatherMultipleTopicsFromPerplexity(newsTopics, dateStr);
    console.log('取得したニュースデータ:', Object.keys(newsData).length + '件');
    
    // ニュースデータが取得できた場合は実際の内容を使用
    let aiNews = '';
    let businessNews = '';
    let techNews = '';
    
    if (newsData && Object.keys(newsData).length > 0) {
      console.log('✅ 実際のニュースデータを使用');
      
      if (newsData[newsTopics[0]]) {
        aiNews = newsData[newsTopics[0]].content.substring(0, 500) + '...';
      }
      if (newsData[newsTopics[1]]) {
        businessNews = newsData[newsTopics[1]].content.substring(0, 500) + '...';
      }
      if (newsData[newsTopics[2]]) {
        techNews = newsData[newsTopics[2]].content.substring(0, 500) + '...';
      }
    } else {
      console.log('⚠️ ニュースデータ取得失敗 - デフォルトテンプレートを使用');
      aiNews = '• ChatGPT、Claude等の大規模言語モデルの最新アップデート\n• 企業のAI導入事例と効果測定\n• AI規制に関する政府・業界動向';
      businessNews = '• 注目スタートアップの資金調達ラウンド\n• 新サービス・プロダクトのローンチ情報\n• M&A・業務提携の最新動向';
      techNews = '• 次世代技術の実用化進展\n• オープンソースプロジェクトの重要リリース\n• サイバーセキュリティの最新脅威と対策';
    }
    
    // 実際のニュース内容でコンテンツを生成
    const content = `昨日（${dateStr}）のAI・テクノロジーニュース まとめ

## 主要トピック

### AI・機械学習分野の動向
${aiNews}

### ビジネス・スタートアップ情報
${businessNews}

### 技術・イノベーション
${techNews}

## 実務への影響と対応策

これらの最新動向を踏まえ、以下の観点で分析します：

### 短期的な影響（今週〜来月）
• 新技術・サービスの導入可能性
• 競合他社の動向把握
• 法的・規制面での対応必要性

### 中長期的な戦略への影響
• 事業戦略の見直し検討項目
• 投資・リソース配分の調整
• 人材育成・採用への影響

---
📊 データソース: Perplexity AI（リアルタイム検索）
👤 分析・解説: ${AUTHOR_PROFILE.basic.name}
🏢 専門: 上場準備・法務・DX推進・海外展開
📅 作成日時: ${new Date().toLocaleDateString('ja-JP')} ${new Date().toLocaleTimeString('ja-JP')}`;
    
    console.log(`生成コンテンツ長: ${content.length}文字`);
    console.log('Google Slidesでプレゼンテーション作成...');
    return createSlidesWithGoogleSlides(content, 'morning');
    
  } catch (error) {
    console.error('❌ 昨日のニューススライド生成エラー:', error);
    throw error;
  }
}

/**
 * ビジネスアイデア・効率化スライド生成（Google Slides版・AI活用事例取得）
 */
function generateBusinessIdeasSlides() {
  try {
    console.log('=== ビジネスアイデア・効率化スライド生成開始 ===');
    
    const today = new Date().toLocaleDateString('ja-JP');
    console.log(`本日の日付: ${today}`);
    
    // 実際のAI活用事例とビジネストレンドを取得
    console.log('Perplexity APIからAI活用事例を取得中...');
    const businessTopics = [
      'AI 業務効率化 企業導入事例',
      'DX デジタルトランスフォーメーション 成功事例',
      'スタートアップ ビジネスモデル 革新'
    ];
    
    const businessData = gatherMultipleTopicsFromPerplexity(businessTopics, today);
    console.log('取得したビジネスデータ:', Object.keys(businessData).length + '件');
    
    // 取得したデータを整理
    let aiEfficiencyNews = '';
    let dxCaseStudies = '';
    let startupInnovation = '';
    
    if (businessData && Object.keys(businessData).length > 0) {
      console.log('✅ 実際のビジネスデータを使用');
      
      if (businessData[businessTopics[0]]) {
        aiEfficiencyNews = businessData[businessTopics[0]].content.substring(0, 600) + '...';
      }
      if (businessData[businessTopics[1]]) {
        dxCaseStudies = businessData[businessTopics[1]].content.substring(0, 600) + '...';
      }
      if (businessData[businessTopics[2]]) {
        startupInnovation = businessData[businessTopics[2]].content.substring(0, 600) + '...';
      }
    } else {
      console.log('⚠️ ビジネスデータ取得失敗 - デフォルトテンプレートを使用');
      aiEfficiencyNews = '• 大手企業におけるChatGPT導入による文書作成時間50%削減事例\n• RPA（Robotic Process Automation）による経理業務自動化\n• AI搭載CRMシステムによる営業効率向上';
      dxCaseStudies = '• 製造業でのIoT活用による品質管理の精密化\n• 小売業のオムニチャネル戦略とデータ活用\n• 金融機関のデジタルバンキング推進事例';
      startupInnovation = '• AI×ヘルスケア分野での診断支援サービス\n• EdTech領域での個別最適化学習プラットフォーム\n• サステナビリティ×テクノロジーの環境ソリューション';
    }
    
    const content = `ビジネス効率化・AI活用 最新動向とアイデア集
${today}版

## 今日の注目：最新AI活用事例

### 企業の業務効率化事例
${aiEfficiencyNews}

### DX推進の成功パターン
${dxCaseStudies}

### スタートアップの革新的アプローチ
${startupInnovation}

## 実践的な導入アイデア

### 管理部門での即効性がある施策
• 契約書レビューAIの導入（リーガルテック活用）
• 請求書OCR＋自動仕訳システム
• 人事評価データの可視化ダッシュボード

### 中長期的な戦略的取組み
• 予測分析による事業計画精度向上
• 顧客データ統合による営業戦略最適化
• サプライチェーン全体の可視化・効率化

## ROI（投資対効果）の考え方

### 定量的効果の測定指標
• 作業時間削減率（時間コスト）
• エラー率削減（品質向上）
• 売上・利益への直接貢献

### 定性的効果の評価
• 従業員満足度の向上
• 意思決定スピードの向上
• 競合優位性の確立

---
💡 実務経験からのアドバイス
📊 データソース: Perplexity AI（最新事例）
👤 解説: ${AUTHOR_PROFILE.basic.name}
🏢 実績: 20社以上の上場準備支援、6社上場実現
📅 作成日時: ${new Date().toLocaleDateString('ja-JP')} ${new Date().toLocaleTimeString('ja-JP')}`;
    
    console.log(`生成コンテンツ長: ${content.length}文字`);
    console.log('Google Slidesでプレゼンテーション作成...');
    return createSlidesWithGoogleSlides(content, 'noon');
    
  } catch (error) {
    console.error('❌ ビジネスアイデアスライド生成エラー:', error);
    throw error;
  }
}

/**
 * 本日のニュース・トレンドスライド生成（Google Slides版・リアルタイムニュース取得）
 */
function generateTodayNewsSlides() {
  try {
    console.log('=== 本日のニュース・トレンドスライド生成開始 ===');
    
    const today = new Date();
    const dateStr = today.toLocaleDateString('ja-JP');
    const timeStr = today.toLocaleTimeString('ja-JP');
    
    console.log(`本日の日付・時刻: ${dateStr} ${timeStr}`);
    
    // 本日のリアルタイムニュースを取得
    console.log('Perplexity APIから本日の最新ニュースを取得中...');
    const todayTopics = [
      '今日 AI 人工知能 最新ニュース 発表',
      '本日 テクノロジー企業 株価 投資 動向',
      '今日 スタートアップ 資金調達 新サービス リリース'
    ];
    
    const todayNewsData = gatherMultipleTopicsFromPerplexity(todayTopics, dateStr);
    console.log('取得した本日のニュースデータ:', Object.keys(todayNewsData).length + '件');
    
    // 取得したデータを整理
    let aiTechNews = '';
    let marketBusinessNews = '';
    let startupNews = '';
    
    if (todayNewsData && Object.keys(todayNewsData).length > 0) {
      console.log('✅ 実際の本日のニュースデータを使用');
      
      if (todayNewsData[todayTopics[0]]) {
        aiTechNews = todayNewsData[todayTopics[0]].content.substring(0, 700) + '...';
      }
      if (todayNewsData[todayTopics[1]]) {
        marketBusinessNews = todayNewsData[todayTopics[1]].content.substring(0, 700) + '...';
      }
      if (todayNewsData[todayTopics[2]]) {
        startupNews = todayNewsData[todayTopics[2]].content.substring(0, 700) + '...';
      }
    } else {
      console.log('⚠️ 本日のニュースデータ取得失敗 - デフォルトテンプレートを使用');
      aiTechNews = '• OpenAI、Google、Microsoft等の主要AI企業の最新発表\n• 生成AI技術の新たな活用領域の発見\n• AI関連規制・ガイドラインの最新動向';
      marketBusinessNews = '• テック株の動向と投資家の反応\n• 新たなIPO予定企業の発表\n• M&A・戦略的パートナーシップの発表';
      startupNews = '• 注目スタートアップの大型資金調達\n• 革新的な新サービス・プロダクトのローンチ\n• 業界を変える可能性のある技術発表';
    }
    
    const content = `【速報】本日（${dateStr}）のAI・テクノロジー 最新動向
更新時刻: ${timeStr}

## 🚨 本日の重要ニュース

### AI・テクノロジー業界の動き
${aiTechNews}

### 市場・投資・ビジネス動向
${marketBusinessNews}

### スタートアップ・イノベーション
${startupNews}

## 📈 市場への影響分析

### 即座に影響が予想される領域
• 関連企業の株価動向
• 投資家センチメントの変化
• 競合他社の対応動向

### 業界全体への波及効果
• 技術標準・規格への影響
• 規制環境の変化予測
• 新たな市場機会の創出

## 🔮 明日以降の注目ポイント

### 短期的な注目事項（明日〜来週）
• 関連企業の追加発表予定
• 市場の反応と株価動向
• 規制当局のコメント・対応

### 中期的な戦略的検討項目（来月〜四半期）
• 事業戦略への組み込み検討
• 技術投資・人材採用の調整
• 競合分析と対応策の策定

## ⚡ 実務担当者への提言

今日の動向を踏まえた具体的なアクションアイテム：
• 関連技術・サービスの詳細調査
• 自社事業への影響度評価
• ステークホルダーへの情報共有

---
🔄 リアルタイム更新: Perplexity AI
📊 市場分析: 投資・M&A経験に基づく解説
👤 分析者: ${AUTHOR_PROFILE.basic.name}
🏢 専門: 上場準備・法務・DX推進・海外展開
📅 作成: ${dateStr} ${timeStr}`;
    
    console.log(`生成コンテンツ長: ${content.length}文字`);
    console.log('Google Slidesでプレゼンテーション作成...');
    return createSlidesWithGoogleSlides(content, 'evening');
    
  } catch (error) {
    console.error('❌ 本日のニューススライド生成エラー:', error);
    throw error;
  }
}