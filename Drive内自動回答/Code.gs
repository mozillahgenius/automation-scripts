/**
 * Drive参照RAGで質問票を自動下書き（GAS＋外部AI）
 *
 * 機能概要:
 * - inbox_questionnaires/ フォルダを監視
 * - 過去QAシートから類似QAを検索（Jaccard類似度）
 * - OpenAI APIで回答を整形
 * - drafts/ フォルダに下書きシートを出力
 */

// ===========================================
// 設定（Script Propertiesから取得）
// ===========================================
const CONFIG = {
  // フォルダID（初回実行時に自動取得またはScript Propertiesから）
  get INBOX_FOLDER_ID() {
    return PropertiesService.getScriptProperties().getProperty('INBOX_FOLDER_ID') || '';
  },
  get DRAFTS_FOLDER_ID() {
    return PropertiesService.getScriptProperties().getProperty('DRAFTS_FOLDER_ID') || '';
  },
  get LOGS_FOLDER_ID() {
    return PropertiesService.getScriptProperties().getProperty('LOGS_FOLDER_ID') || '';
  },
  get QA_SHEET_ID() {
    return PropertiesService.getScriptProperties().getProperty('QA_SHEET_ID') || '';
  },
  get OPENAI_API_KEY() {
    return PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY') || '';
  },

  // 検索パラメータ
  TOP_K: 3,
  THRESHOLD_AUTO: 0.50,    // 自動採用（0.50以上）※日本語Jaccardは低めに出るため緩和
  THRESHOLD_REVIEW: 0.25,  // 要確認（0.25以上）※これ未満は「該当なし」

  // OpenAI設定
  get OPENAI_MODEL() {
    // Script Properties の OPENAI_MODEL があれば優先。未設定時は最新系の軽量モデルを既定に。
    return PropertiesService.getScriptProperties().getProperty('OPENAI_MODEL') || 'gpt-4o-mini';
  },
  OPENAI_MAX_TOKENS: 2000
};

// ===========================================
// メイン処理
// ===========================================

/**
 * メイン実行関数（トリガーから呼び出し）
 */
function processInboxFiles() {
  const startTime = new Date();
  const logs = [];

  try {
    // 設定チェック
    validateConfig();

    // 過去QAを読み込み
    const qaIndex = loadQAIndex();
    logs.push(`QAインデックス読み込み完了: ${qaIndex.length}件`);

    // inboxフォルダのファイルを取得
    const inboxFolder = DriveApp.getFolderById(CONFIG.INBOX_FOLDER_ID);
    const files = inboxFolder.getFiles();

    let processedCount = 0;

    while (files.hasNext()) {
      const file = files.next();
      const fileName = file.getName();
      const mimeType = file.getMimeType();

      try {
        logs.push(`処理開始: ${fileName}`);

        // Sheets または Excel を処理
        let spreadsheet;
        if (mimeType === MimeType.GOOGLE_SHEETS) {
          spreadsheet = SpreadsheetApp.openById(file.getId());
        } else if (mimeType === MimeType.MICROSOFT_EXCEL || fileName.endsWith('.xlsx') || fileName.endsWith('.xls')) {
          // ExcelをSheetsに変換
          const convertedFile = convertExcelToSheets(file);
          spreadsheet = SpreadsheetApp.openById(convertedFile.getId());
          logs.push(`Excel変換完了: ${fileName}`);
        } else {
          logs.push(`スキップ（非対応形式）: ${fileName}`);
          continue;
        }

        // 質問を抽出
        const questions = extractQuestions(spreadsheet);
        logs.push(`質問抽出: ${questions.length}件`);

        if (questions.length === 0) {
          logs.push(`質問が見つかりません: ${fileName}`);
          continue;
        }

        // 各質問に対して検索・下書き生成
        const drafts = [];
        for (let i = 0; i < questions.length; i++) {
          const question = questions[i];

          // 類似QA検索
          const candidates = searchSimilarQA(question.text, qaIndex);

          // 下書き生成
          const draft = generateDraft(question, candidates);
          drafts.push(draft);
        }

        // 下書きシートを作成
        createDraftSheet(file.getName(), drafts);
        logs.push(`下書き作成完了: ${drafts.length}件`);

        // 処理済みファイルを移動（オプション：アーカイブフォルダへ）
        // moveToArchive(file);

        processedCount++;

      } catch (fileError) {
        logs.push(`エラー（${fileName}）: ${fileError.message}`);
      }
    }

    logs.push(`処理完了: ${processedCount}ファイル`);

  } catch (error) {
    logs.push(`致命的エラー: ${error.message}`);
  }

  // ログを保存
  saveLog(startTime, logs);
}

// ===========================================
// QAインデックス管理
// ===========================================

/**
 * 過去QAシートを読み込み、インデックス化
 * @returns {Array} QAオブジェクトの配列
 */
function loadQAIndex() {
  const sheet = SpreadsheetApp.openById(CONFIG.QA_SHEET_ID).getActiveSheet();
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  // 列インデックスを特定
  const colIndex = {
    question: headers.indexOf('Question'),
    answer: headers.indexOf('Answer'),
    sourceUrl: headers.indexOf('SourceURL'),
    updatedAt: headers.indexOf('UpdatedAt'),
    tags: headers.indexOf('Tags')
  };

  // 必須列チェック
  if (colIndex.question === -1 || colIndex.answer === -1 || colIndex.sourceUrl === -1) {
    throw new Error('QAシートに必須列（Question/Answer/SourceURL）がありません');
  }

  const qaIndex = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const question = String(row[colIndex.question] || '').trim();
    const answer = String(row[colIndex.answer] || '').trim();

    if (!question || !answer) continue;

    qaIndex.push({
      question: question,
      answer: answer,
      sourceUrl: String(row[colIndex.sourceUrl] || ''),
      updatedAt: colIndex.updatedAt !== -1 ? row[colIndex.updatedAt] : null,
      tags: colIndex.tags !== -1 ? String(row[colIndex.tags] || '') : '',
      // 正規化済みトークン（検索用）
      tokens: tokenize(question)
    });
  }

  return qaIndex;
}

// ===========================================
// 質問抽出
// ===========================================

/**
 * スプレッドシートから質問を抽出
 * @param {Spreadsheet} spreadsheet - 対象スプレッドシート
 * @returns {Array} 質問オブジェクトの配列
 */
function extractQuestions(spreadsheet) {
  const sheet = spreadsheet.getActiveSheet();
  const data = sheet.getDataRange().getValues();

  if (data.length < 2) return [];

  // 質問列を検出
  const questionCol = detectQuestionColumn(data);
  if (questionCol === -1) {
    // フォールバック：2列目（B列）を質問列と仮定
    Logger.log('質問列を自動検出できませんでした。B列を使用します。');
    return extractFromColumn(data, 1);
  }

  return extractFromColumn(data, questionCol);
}

/**
 * 質問列を自動検出
 * @param {Array} data - シートデータ
 * @returns {number} 列インデックス（-1: 検出失敗）
 */
function detectQuestionColumn(data) {
  const headers = data[0];

  // ヘッダーから質問列を推定
  const questionKeywords = ['質問', '問', 'question', 'q', '内容', '設問', '問い合わせ'];

  for (let col = 0; col < headers.length; col++) {
    const header = String(headers[col]).toLowerCase().trim();
    for (const keyword of questionKeywords) {
      if (header.includes(keyword)) {
        return col;
      }
    }
  }

  // ヘッダーで見つからない場合、最も長いテキストが多い列を質問列と推定
  const avgLengths = [];
  for (let col = 0; col < headers.length; col++) {
    let totalLength = 0;
    let count = 0;
    for (let row = 1; row < Math.min(data.length, 10); row++) {
      const cellValue = String(data[row][col] || '');
      if (cellValue.length > 10) {
        totalLength += cellValue.length;
        count++;
      }
    }
    avgLengths.push(count > 0 ? totalLength / count : 0);
  }

  const maxAvgLength = Math.max(...avgLengths);
  if (maxAvgLength > 20) {
    return avgLengths.indexOf(maxAvgLength);
  }

  return -1;
}

/**
 * 指定列から質問を抽出
 * @param {Array} data - シートデータ
 * @param {number} col - 列インデックス
 * @returns {Array} 質問オブジェクトの配列
 */
function extractFromColumn(data, col) {
  const questions = [];

  for (let row = 1; row < data.length; row++) {
    const cellValue = String(data[row][col] || '').trim();

    // 空セル、短すぎるテキストはスキップ
    if (cellValue.length < 5) continue;

    // 番号や記号のみの場合はスキップ
    if (/^[\d\.\-\(\)]+$/.test(cellValue)) continue;

    questions.push({
      rowNumber: row + 1,
      text: cellValue
    });
  }

  return questions;
}

// ===========================================
// 類似度検索（Jaccard）
// ===========================================

/**
 * テキストをトークン化（正規化）
 * @param {string} text - 入力テキスト
 * @returns {Set} トークンのセット
 */
function tokenize(text) {
  // 小文字化
  let normalized = text.toLowerCase();

  // 記号除去（日本語対応）
  normalized = normalized.replace(/[、。！？「」『』（）\[\]【】・…―]/g, ' ');
  normalized = normalized.replace(/[.,!?;:'"()\[\]{}<>]/g, ' ');

  // 空白正規化
  normalized = normalized.replace(/\s+/g, ' ').trim();

  // 単語分割（スペース区切り + 2文字以上のN-gram）
  const words = normalized.split(' ').filter(w => w.length > 0);
  const tokens = new Set(words);

  // 日本語用：2文字N-gramを追加
  const noSpaceText = normalized.replace(/\s/g, '');
  for (let i = 0; i < noSpaceText.length - 1; i++) {
    tokens.add(noSpaceText.substring(i, i + 2));
  }

  return tokens;
}

/**
 * Jaccard類似度を計算
 * @param {Set} tokens1 - トークンセット1
 * @param {Set} tokens2 - トークンセット2
 * @returns {number} 類似度（0〜1）
 */
function jaccardSimilarity(tokens1, tokens2) {
  if (tokens1.size === 0 || tokens2.size === 0) return 0;

  let intersection = 0;
  for (const token of tokens1) {
    if (tokens2.has(token)) intersection++;
  }

  const union = tokens1.size + tokens2.size - intersection;
  return union > 0 ? intersection / union : 0;
}

/**
 * 類似QAを検索（複合スコアリング）
 * @param {string} questionText - 検索クエリ
 * @param {Array} qaIndex - QAインデックス
 * @returns {Array} 候補QA（スコア順、上位K件）
 */
function searchSimilarQA(questionText, qaIndex) {
  const queryTokens = tokenize(questionText);
  const queryText = questionText.toLowerCase();

  const scored = qaIndex.map(qa => {
    // 1. Jaccard類似度（基本スコア）
    let jaccardScore = jaccardSimilarity(queryTokens, qa.tokens);

    // 2. キーワード包含率（クエリのトークンがQAにどれだけ含まれるか）
    let containmentScore = 0;
    let matchedTokens = 0;
    for (const token of queryTokens) {
      if (qa.tokens.has(token)) {
        matchedTokens++;
      }
    }
    if (queryTokens.size > 0) {
      containmentScore = matchedTokens / queryTokens.size;
    }

    // 3. 部分文字列マッチ（重要キーワードの直接マッチ）
    let substringBonus = 0;
    const qaQuestion = qa.question.toLowerCase();
    // 3文字以上の単語で部分一致チェック
    const queryWords = queryText.split(/[\s、。！？]+/).filter(w => w.length >= 3);
    for (const word of queryWords) {
      if (qaQuestion.includes(word)) {
        substringBonus += 0.15;
      }
    }
    substringBonus = Math.min(substringBonus, 0.3); // 上限30%

    // 4. タグマッチボーナス
    let tagBonus = 0;
    if (qa.tags) {
      const tags = qa.tags.toLowerCase().split(',').map(t => t.trim());
      for (const tag of tags) {
        if (tag && queryText.includes(tag)) {
          tagBonus += 0.1;
        }
      }
      tagBonus = Math.min(tagBonus, 0.2); // 上限20%
    }

    // 複合スコア計算（重み付け）
    let score = (jaccardScore * 0.4) + (containmentScore * 0.4) + substringBonus + tagBonus;

    // 5. UpdatedAtによるブースト（新しいほど有利）
    if (qa.updatedAt) {
      const ageInDays = (new Date() - new Date(qa.updatedAt)) / (1000 * 60 * 60 * 24);
      if (ageInDays < 365) {
        score *= 1 + (0.1 * (1 - ageInDays / 365));
      }
    }

    return {
      ...qa,
      score: Math.min(score, 1.0), // 1.0を上限に
      debugScores: { jaccardScore, containmentScore, substringBonus, tagBonus } // デバッグ用
    };
  });

  // スコア降順でソート、上位K件を返す
  scored.sort((a, b) => b.score - a.score);
  return scored.slice(0, CONFIG.TOP_K);
}

// ===========================================
// 下書き生成（OpenAI API）
// ===========================================

/**
 * 下書きを生成
 * @param {Object} question - 質問オブジェクト
 * @param {Array} candidates - 候補QA
 * @returns {Object} 下書きオブジェクト
 */
function generateDraft(question, candidates) {
  const topCandidate = candidates[0];
  const matchScore = topCandidate ? topCandidate.score : 0;

  let status, draftAnswer, sourceUrl;

  if (matchScore >= CONFIG.THRESHOLD_AUTO) {
    // 高一致：自動採用
    status = 'auto';
    const aiResult = callOpenAI(question.text, candidates);
    draftAnswer = aiResult.answer;
    sourceUrl = aiResult.sourceUrl || topCandidate.sourceUrl;
  } else if (matchScore >= CONFIG.THRESHOLD_REVIEW) {
    // 中一致：要確認
    status = 'needs_review';
    const aiResult = callOpenAI(question.text, candidates);
    draftAnswer = aiResult.answer;
    sourceUrl = aiResult.sourceUrl || topCandidate.sourceUrl;
  } else if (matchScore > 0 && topCandidate) {
    // 低一致だが候補あり：参考情報として提示
    status = 'needs_review';
    const aiResult = callOpenAI(question.text, candidates);
    draftAnswer = aiResult.answer;
    sourceUrl = aiResult.sourceUrl || topCandidate.sourceUrl;
  } else {
    // 該当なし
    status = 'needs_review';
    draftAnswer = '【該当するQAが見つかりませんでした。手動での回答が必要です。】';
    sourceUrl = '';
  }

  // 類似度のステータスラベルを生成
  let matchStatusLabel;
  if (matchScore >= CONFIG.THRESHOLD_AUTO) {
    matchStatusLabel = '高一致（自動採用）';
  } else if (matchScore >= CONFIG.THRESHOLD_REVIEW) {
    matchStatusLabel = '中一致（要確認）';
  } else if (matchScore > 0) {
    matchStatusLabel = '低一致（参考情報）';
  } else {
    matchStatusLabel = '該当なし';
  }

  return {
    rowNumber: question.rowNumber,
    originalQuestion: question.text,
    draftAnswer: draftAnswer,
    sourceUrl: sourceUrl,
    matchScore: matchScore,
    matchStatusLabel: matchStatusLabel,
    status: status,
    reviewerNote: ''
  };
}

/**
 * OpenAI APIを呼び出して回答を整形
 * @param {string} questionText - 質問テキスト
 * @param {Array} candidates - 候補QA
 * @returns {Object} {answer, sourceUrl}
 */
function callOpenAI(questionText, candidates) {
  if (!CONFIG.OPENAI_API_KEY) {
    // APIキーがない場合は候補の回答をそのまま返す
    return {
      answer: candidates[0]?.answer || '',
      sourceUrl: candidates[0]?.sourceUrl || ''
    };
  }

  // プロンプト構築
  const candidateText = candidates.map((c, i) =>
    `【候補${i + 1}】\n質問: ${c.question}\n回答: ${c.answer}\n出典: ${c.sourceUrl}`
  ).join('\n\n');

  const systemPrompt = `あなたは質問回答の整形アシスタントです。
以下のルールを厳守してください：
1. 提供された候補QAの情報のみを使用して回答を作成してください
2. 候補QAに含まれない新しい情報や事実を追加することは禁止です
3. 回答は日本語で、丁寧かつ簡潔に整形してください
4. 最も適切な候補を選び、その出典URLを必ず返してください
5. 出力はJSON形式で {"answer": "回答文", "sourceUrl": "URL"} としてください`;

  const userPrompt = `【質問】
${questionText}

【候補QA】
${candidateText}

上記の候補QAを参考に、質問への回答を整形してください。`;

  try {
    const response = UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions', {
      method: 'post',
      contentType: 'application/json',
      headers: {
        'Authorization': `Bearer ${CONFIG.OPENAI_API_KEY}`
      },
      payload: JSON.stringify({
        model: CONFIG.OPENAI_MODEL,
        messages: [
          { role: 'system', content: systemPrompt },
          { role: 'user', content: userPrompt }
        ],
        max_tokens: CONFIG.OPENAI_MAX_TOKENS,
        temperature: 0.3
      }),
      muteHttpExceptions: true
    });

    const result = JSON.parse(response.getContentText());

    if (result.error) {
      Logger.log('OpenAI API Error: ' + result.error.message);
      return {
        answer: candidates[0]?.answer || '',
        sourceUrl: candidates[0]?.sourceUrl || ''
      };
    }

    const content = result.choices[0].message.content;

    // JSONをパース
    try {
      const parsed = JSON.parse(content);
      return {
        answer: parsed.answer || '',
        sourceUrl: parsed.sourceUrl || candidates[0]?.sourceUrl || ''
      };
    } catch (parseError) {
      // JSONパース失敗時はテキストをそのまま使用
      return {
        answer: content,
        sourceUrl: candidates[0]?.sourceUrl || ''
      };
    }

  } catch (error) {
    Logger.log('OpenAI API呼び出しエラー: ' + error.message);
    return {
      answer: candidates[0]?.answer || '',
      sourceUrl: candidates[0]?.sourceUrl || ''
    };
  }
}

// ===========================================
// 下書きシート作成
// ===========================================

/**
 * 下書きシートを作成
 * @param {string} originalFileName - 元ファイル名
 * @param {Array} drafts - 下書き配列
 */
function createDraftSheet(originalFileName, drafts) {
  const draftsFolder = DriveApp.getFolderById(CONFIG.DRAFTS_FOLDER_ID);

  // ファイル名を生成
  const timestamp = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss');
  const draftFileName = `下書き_${originalFileName}_${timestamp}`;

  // 新規スプレッドシート作成
  const spreadsheet = SpreadsheetApp.create(draftFileName);
  const sheet = spreadsheet.getActiveSheet();
  sheet.setName('下書き');

  // ヘッダー設定
  const headers = ['No', 'OriginalQuestion', 'DraftAnswer', 'MatchScore(%)', 'MatchStatus', 'SourceURL', 'Status', 'ReviewerNote'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setBackground('#4285f4')
    .setFontColor('#ffffff')
    .setFontWeight('bold');

  // データ設定
  const dataRows = drafts.map((draft, index) => [
    index + 1,
    draft.originalQuestion,
    draft.draftAnswer,
    Math.round(draft.matchScore * 100),  // パーセント表示
    draft.matchStatusLabel,
    draft.sourceUrl,
    draft.status,
    draft.reviewerNote
  ]);

  if (dataRows.length > 0) {
    sheet.getRange(2, 1, dataRows.length, headers.length).setValues(dataRows);
  }

  // 列幅調整
  sheet.setColumnWidth(1, 50);   // No
  sheet.setColumnWidth(2, 400);  // OriginalQuestion
  sheet.setColumnWidth(3, 400);  // DraftAnswer
  sheet.setColumnWidth(4, 100);  // MatchScore(%)
  sheet.setColumnWidth(5, 150);  // MatchStatus
  sheet.setColumnWidth(6, 200);  // SourceURL
  sheet.setColumnWidth(7, 120);  // Status
  sheet.setColumnWidth(8, 200);  // ReviewerNote

  // 条件付き書式（Status列）
  const statusRange = sheet.getRange(2, 7, dataRows.length, 1);

  // auto: 緑
  const ruleAuto = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('auto')
    .setBackground('#d9ead3')
    .setRanges([statusRange])
    .build();

  // needs_review: 黄
  const ruleReview = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('needs_review')
    .setBackground('#fff2cc')
    .setRanges([statusRange])
    .build();

  // approved: 青
  const ruleApproved = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('approved')
    .setBackground('#cfe2f3')
    .setRanges([statusRange])
    .build();

  sheet.setConditionalFormatRules([ruleAuto, ruleReview, ruleApproved]);

  // ファイルをdraftsフォルダに移動
  const file = DriveApp.getFileById(spreadsheet.getId());
  file.moveTo(draftsFolder);

  Logger.log(`下書きシート作成: ${draftFileName}`);
}

// ===========================================
// Excel変換
// ===========================================

/**
 * ExcelファイルをGoogle Sheetsに変換
 * @param {File} file - Excelファイル
 * @returns {File} 変換後のSheetsファイル
 */
function convertExcelToSheets(file) {
  const blob = file.getBlob();
  const resource = {
    title: file.getName().replace(/\.xlsx?$/i, ''),
    mimeType: MimeType.GOOGLE_SHEETS
  };

  const convertedFile = Drive.Files.insert(resource, blob, {
    convert: true
  });

  return DriveApp.getFileById(convertedFile.id);
}

// ===========================================
// ログ管理
// ===========================================

/**
 * ログを保存
 * @param {Date} startTime - 開始時刻
 * @param {Array} logs - ログメッセージ配列
 */
function saveLog(startTime, logs) {
  const logsFolder = DriveApp.getFolderById(CONFIG.LOGS_FOLDER_ID);
  const timestamp = Utilities.formatDate(startTime, 'Asia/Tokyo', 'yyyyMMdd_HHmmss');
  const fileName = `log_${timestamp}.txt`;

  const endTime = new Date();
  const duration = (endTime - startTime) / 1000;

  const logContent = [
    `=== 実行ログ ===`,
    `開始: ${startTime.toLocaleString('ja-JP')}`,
    `終了: ${endTime.toLocaleString('ja-JP')}`,
    `所要時間: ${duration}秒`,
    ``,
    `--- 詳細 ---`,
    ...logs
  ].join('\n');

  logsFolder.createFile(fileName, logContent, MimeType.PLAIN_TEXT);
}

// ===========================================
// 設定・初期化
// ===========================================

/**
 * 設定を検証
 */
function validateConfig() {
  const required = ['INBOX_FOLDER_ID', 'DRAFTS_FOLDER_ID', 'LOGS_FOLDER_ID', 'QA_SHEET_ID'];
  const missing = required.filter(key => !CONFIG[key]);

  if (missing.length > 0) {
    throw new Error(`設定が不足しています: ${missing.join(', ')}\nsetup()を実行して設定してください。`);
  }
}

/**
 * ========================================
 * 初期セットアップ（完全自動・UIなし）
 * ========================================
 *
 * この関数を実行するだけで以下が自動作成されます：
 * - ルートフォルダ（QA自動下書きシステム）※スクリプトがバインドされたシートと同じ階層
 * - inbox_questionnaires/ フォルダ
 * - drafts/ フォルダ
 * - logs/ フォルダ
 * - 過去QAシート（ヘッダー設定済み）
 * - 時間トリガー（1時間ごと）
 *
 * @param {Object} options - オプション設定
 * @param {string} options.rootFolderName - ルートフォルダ名（デフォルト: 'QA自動下書きシステム'）
 * @param {string} options.openaiApiKey - OpenAI APIキー（オプション）
 * @param {string} options.parentFolderId - 親フォルダID（指定しない場合はスクリプトのシートと同じフォルダ）
 * @param {boolean} options.createTrigger - 時間トリガーを作成するか（デフォルト: true）
 * @param {number} options.triggerIntervalHours - トリガー間隔（時間、デフォルト: 1）
 * @returns {Object} 作成されたリソースのID情報
 */
function setup(options = {}) {
  // デフォルト設定
  const config = {
    rootFolderName: options.rootFolderName || 'QA自動下書きシステム',
    openaiApiKey: options.openaiApiKey || '',
    parentFolderId: options.parentFolderId || null,
    createTrigger: options.createTrigger !== false,
    triggerIntervalHours: options.triggerIntervalHours || 1
  };

  Logger.log('=== 初期セットアップ開始 ===');
  Logger.log(`ルートフォルダ名: ${config.rootFolderName}`);

  // 既存設定のチェック
  const existingConfig = checkExistingSetup();
  if (existingConfig.isComplete) {
    Logger.log('既存の設定が見つかりました。再セットアップする場合は resetSetup() を先に実行してください。');
    Logger.log('既存設定:');
    Logger.log(`  ROOT_FOLDER_ID: ${existingConfig.rootFolderId}`);
    Logger.log(`  QA_SHEET_ID: ${existingConfig.qaSheetId}`);
    return existingConfig;
  }

  // 親フォルダを決定（スクリプトがバインドされたシートの親フォルダを優先）
  let parentFolder = null;

  if (config.parentFolderId) {
    // オプションで明示的に指定された場合
    parentFolder = DriveApp.getFolderById(config.parentFolderId);
    Logger.log(`指定された親フォルダを使用: ${parentFolder.getName()}`);
  } else {
    // スクリプトがバインドされたスプレッドシートの親フォルダを取得
    parentFolder = getScriptParentFolder();
    if (parentFolder) {
      Logger.log(`スクリプトのシートと同じフォルダを使用: ${parentFolder.getName()}`);
    } else {
      Logger.log('親フォルダを特定できませんでした。マイドライブ直下に作成します。');
    }
  }

  // ルートフォルダ作成
  let rootFolder;
  if (parentFolder) {
    rootFolder = parentFolder.createFolder(config.rootFolderName);
    Logger.log(`親フォルダ「${parentFolder.getName()}」内にルートフォルダを作成`);
  } else {
    rootFolder = DriveApp.createFolder(config.rootFolderName);
    Logger.log('マイドライブ直下にルートフォルダを作成');
  }
  Logger.log(`ルートフォルダID: ${rootFolder.getId()}`);

  // サブフォルダ作成
  const inboxFolder = rootFolder.createFolder('inbox_questionnaires');
  Logger.log(`inbox_questionnaires フォルダ作成: ${inboxFolder.getId()}`);

  const draftsFolder = rootFolder.createFolder('drafts');
  Logger.log(`drafts フォルダ作成: ${draftsFolder.getId()}`);

  const logsFolder = rootFolder.createFolder('logs');
  Logger.log(`logs フォルダ作成: ${logsFolder.getId()}`);

  // 処理済みファイル用フォルダ（オプション）
  const processedFolder = rootFolder.createFolder('processed');
  Logger.log(`processed フォルダ作成: ${processedFolder.getId()}`);

  // 過去QAシート作成
  const qaSpreadsheet = SpreadsheetApp.create('過去QA');
  const qaSheet = qaSpreadsheet.getActiveSheet();
  qaSheet.setName('QA');

  // QAシートのヘッダー設定
  const qaHeaders = ['Q_ID', 'Question', 'Answer', 'SourceURL', 'UpdatedAt', 'Tags'];
  qaSheet.getRange(1, 1, 1, qaHeaders.length).setValues([qaHeaders]);
  qaSheet.getRange(1, 1, 1, qaHeaders.length)
    .setBackground('#4285f4')
    .setFontColor('#ffffff')
    .setFontWeight('bold');

  // 列幅設定
  qaSheet.setColumnWidth(1, 80);   // Q_ID
  qaSheet.setColumnWidth(2, 400);  // Question
  qaSheet.setColumnWidth(3, 400);  // Answer
  qaSheet.setColumnWidth(4, 250);  // SourceURL
  qaSheet.setColumnWidth(5, 120);  // UpdatedAt
  qaSheet.setColumnWidth(6, 150);  // Tags

  // サンプルデータを挿入（動作確認用）
  const sampleData = [
    ['QA001', '請求書の発行方法を教えてください', '請求書は管理画面の「請求書発行」メニューから発行できます。発行後はPDFでダウンロード可能です。', 'https://example.com/help/invoice', new Date(), '請求,発行'],
    ['QA002', 'パスワードをリセットするにはどうすればいいですか', 'ログイン画面の「パスワードを忘れた方」リンクをクリックし、登録メールアドレスを入力してください。リセット用のリンクが送信されます。', 'https://example.com/help/password', new Date(), 'パスワード,リセット,ログイン'],
    ['QA003', '契約を解約したい場合の手続きは', '解約は「設定」→「契約管理」→「解約手続き」から申請できます。解約は月末締めとなり、翌月から適用されます。', 'https://example.com/help/cancel', new Date(), '解約,契約']
  ];
  qaSheet.getRange(2, 1, sampleData.length, qaHeaders.length).setValues(sampleData);
  Logger.log('サンプルQAデータを挿入しました（3件）');

  // QAシートをルートフォルダに移動
  const qaFile = DriveApp.getFileById(qaSpreadsheet.getId());
  qaFile.moveTo(rootFolder);
  Logger.log(`過去QAシート作成: ${qaSpreadsheet.getId()}`);

  // Script Propertiesに保存
  const props = PropertiesService.getScriptProperties();
  const properties = {
    'ROOT_FOLDER_ID': rootFolder.getId(),
    'INBOX_FOLDER_ID': inboxFolder.getId(),
    'DRAFTS_FOLDER_ID': draftsFolder.getId(),
    'LOGS_FOLDER_ID': logsFolder.getId(),
    'PROCESSED_FOLDER_ID': processedFolder.getId(),
    'QA_SHEET_ID': qaSpreadsheet.getId()
  };

  // OpenAI APIキーが指定されていれば保存
  if (config.openaiApiKey) {
    properties['OPENAI_API_KEY'] = config.openaiApiKey;
    Logger.log('OpenAI APIキーを設定しました');
  }

  props.setProperties(properties);
  Logger.log('Script Propertiesに設定を保存しました');

  // 時間トリガー作成
  if (config.createTrigger) {
    setupTimeTrigger(config.triggerIntervalHours);
    Logger.log(`時間トリガーを作成しました（${config.triggerIntervalHours}時間ごと）`);
  }

  // 結果をまとめる
  const result = {
    rootFolderId: rootFolder.getId(),
    rootFolderUrl: rootFolder.getUrl(),
    inboxFolderId: inboxFolder.getId(),
    inboxFolderUrl: inboxFolder.getUrl(),
    draftsFolderId: draftsFolder.getId(),
    draftsFolderUrl: draftsFolder.getUrl(),
    logsFolderId: logsFolder.getId(),
    processedFolderId: processedFolder.getId(),
    qaSheetId: qaSpreadsheet.getId(),
    qaSheetUrl: qaSpreadsheet.getUrl()
  };

  Logger.log('');
  Logger.log('=== セットアップ完了 ===');
  Logger.log('');
  Logger.log('【作成されたリソース】');
  Logger.log(`ルートフォルダ: ${result.rootFolderUrl}`);
  Logger.log(`  inbox_questionnaires/: 質問票をここに投入`);
  Logger.log(`  drafts/: 下書きがここに出力`);
  Logger.log(`  logs/: 実行ログ`);
  Logger.log(`  processed/: 処理済みファイル`);
  Logger.log(`過去QAシート: ${result.qaSheetUrl}`);
  Logger.log('');
  Logger.log('【次のステップ】');
  Logger.log('1. 過去QAシートに実際のQAデータを入力');
  Logger.log('2. inbox_questionnaires に質問票ファイルを投入');
  Logger.log('3. processInboxFiles() を実行（または時間トリガーを待つ）');
  Logger.log('');
  Logger.log('※ OpenAI APIキーを後から設定する場合: setOpenAIKey("your-api-key")');

  return result;
}

/**
 * スクリプトがバインドされたスプレッドシートの親フォルダを取得
 * @returns {Folder|null} 親フォルダ（取得できない場合はnull）
 */
function getScriptParentFolder() {
  try {
    // コンテナバインドスクリプトの場合、SpreadsheetApp.getActiveSpreadsheet()で取得可能
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (ss) {
      const file = DriveApp.getFileById(ss.getId());
      const parents = file.getParents();
      if (parents.hasNext()) {
        return parents.next();
      }
    }
  } catch (e) {
    Logger.log('コンテナバインドスクリプトではありません: ' + e.message);
  }

  try {
    // スタンドアロンスクリプトの場合、スクリプトファイル自体の親フォルダを取得
    const scriptId = ScriptApp.getScriptId();
    const scriptFile = DriveApp.getFileById(scriptId);
    const parents = scriptFile.getParents();
    if (parents.hasNext()) {
      return parents.next();
    }
  } catch (e) {
    Logger.log('スクリプトファイルの親フォルダを取得できませんでした: ' + e.message);
  }

  return null;
}

/**
 * 既存セットアップの確認
 * @returns {Object} 既存設定の情報
 */
function checkExistingSetup() {
  const props = PropertiesService.getScriptProperties();
  const rootFolderId = props.getProperty('ROOT_FOLDER_ID');
  const qaSheetId = props.getProperty('QA_SHEET_ID');
  const inboxFolderId = props.getProperty('INBOX_FOLDER_ID');

  return {
    isComplete: !!(rootFolderId && qaSheetId && inboxFolderId),
    rootFolderId: rootFolderId,
    qaSheetId: qaSheetId,
    inboxFolderId: inboxFolderId
  };
}

/**
 * セットアップをリセット（Script Propertiesをクリア）
 * ※ 実際のフォルダ・ファイルは削除されません
 */
function resetSetup() {
  const props = PropertiesService.getScriptProperties();
  const keys = [
    'ROOT_FOLDER_ID',
    'INBOX_FOLDER_ID',
    'DRAFTS_FOLDER_ID',
    'LOGS_FOLDER_ID',
    'PROCESSED_FOLDER_ID',
    'QA_SHEET_ID'
  ];

  keys.forEach(key => props.deleteProperty(key));

  // トリガーも削除
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'processInboxFiles') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  Logger.log('セットアップをリセットしました。setup()で再セットアップできます。');
  Logger.log('※ 実際のDriveフォルダ・ファイルは削除されていません。必要に応じて手動で削除してください。');
}

/**
 * OpenAI APIキーを設定
 * @param {string} apiKey - OpenAI APIキー
 */
function setOpenAIKey(apiKey) {
  if (!apiKey || typeof apiKey !== 'string') {
    Logger.log('エラー: 有効なAPIキーを指定してください');
    return;
  }

  PropertiesService.getScriptProperties().setProperty('OPENAI_API_KEY', apiKey);
  Logger.log('OpenAI APIキーを設定しました');
}

/**
 * 時間トリガーを設定
 * @param {number} intervalHours - 実行間隔（時間）
 */
function setupTimeTrigger(intervalHours = 1) {
  // 既存のトリガーを削除
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'processInboxFiles') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // 新しいトリガーを作成
  ScriptApp.newTrigger('processInboxFiles')
    .timeBased()
    .everyHours(intervalHours)
    .create();

  Logger.log(`時間トリガーを設定しました: ${intervalHours}時間ごとに processInboxFiles() を実行`);
}

/**
 * トリガーを削除
 */
function removeTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  let removed = 0;
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'processInboxFiles') {
      ScriptApp.deleteTrigger(trigger);
      removed++;
    }
  });
  Logger.log(`${removed}件のトリガーを削除しました`);
}

/**
 * 手動設定（既存フォルダ/シートのIDを指定してセットアップ）
 * @param {Object} ids - 各リソースのID
 * @param {string} ids.inboxFolderId - inboxフォルダID（必須）
 * @param {string} ids.draftsFolderId - draftsフォルダID（必須）
 * @param {string} ids.logsFolderId - logsフォルダID（必須）
 * @param {string} ids.qaSheetId - 過去QAシートID（必須）
 * @param {string} ids.rootFolderId - ルートフォルダID（オプション）
 * @param {string} ids.processedFolderId - 処理済みフォルダID（オプション）
 * @param {string} ids.openaiApiKey - OpenAI APIキー（オプション）
 */
function manualSetup(ids) {
  if (!ids || !ids.inboxFolderId || !ids.draftsFolderId || !ids.logsFolderId || !ids.qaSheetId) {
    Logger.log('エラー: 必須パラメータが不足しています');
    Logger.log('必須: inboxFolderId, draftsFolderId, logsFolderId, qaSheetId');
    Logger.log('');
    Logger.log('使用例:');
    Logger.log('manualSetup({');
    Logger.log('  inboxFolderId: "xxx",');
    Logger.log('  draftsFolderId: "xxx",');
    Logger.log('  logsFolderId: "xxx",');
    Logger.log('  qaSheetId: "xxx"');
    Logger.log('});');
    return;
  }

  const props = PropertiesService.getScriptProperties();

  // 必須設定
  props.setProperty('INBOX_FOLDER_ID', ids.inboxFolderId);
  props.setProperty('DRAFTS_FOLDER_ID', ids.draftsFolderId);
  props.setProperty('LOGS_FOLDER_ID', ids.logsFolderId);
  props.setProperty('QA_SHEET_ID', ids.qaSheetId);

  // オプション設定
  if (ids.rootFolderId) {
    props.setProperty('ROOT_FOLDER_ID', ids.rootFolderId);
  }
  if (ids.processedFolderId) {
    props.setProperty('PROCESSED_FOLDER_ID', ids.processedFolderId);
  }
  if (ids.openaiApiKey) {
    props.setProperty('OPENAI_API_KEY', ids.openaiApiKey);
  }

  // 設定内容を検証
  try {
    DriveApp.getFolderById(ids.inboxFolderId);
    DriveApp.getFolderById(ids.draftsFolderId);
    DriveApp.getFolderById(ids.logsFolderId);
    SpreadsheetApp.openById(ids.qaSheetId);
    Logger.log('手動設定が完了しました。すべてのリソースにアクセス可能です。');
  } catch (e) {
    Logger.log('警告: 一部のリソースにアクセスできません: ' + e.message);
    Logger.log('IDが正しいか、アクセス権限があるか確認してください。');
  }
}

/**
 * メニューを追加（スプレッドシートから操作する場合用）
 */
function onOpen() {
  try {
    const ui = SpreadsheetApp.getUi();
    ui.createMenu('QA自動下書き')
      .addItem('初期セットアップ', 'setup')
      .addItem('設定リセット', 'resetSetup')
      .addSeparator()
      .addItem('今すぐ実行', 'processInboxFiles')
      .addItem('設定確認', 'checkConfig')
      .addSeparator()
      .addItem('トリガー設定（1時間）', 'setupTimeTrigger')
      .addItem('トリガー削除', 'removeTrigger')
      .addToUi();
  } catch (e) {
    // スプレッドシート以外から実行された場合は無視
  }
}

// ===========================================
// テスト用関数
// ===========================================

/**
 * Jaccard類似度のテスト
 */
function testJaccard() {
  const text1 = '請求書の発行方法を教えてください';
  const text2 = '請求書はどのように発行しますか';

  const tokens1 = tokenize(text1);
  const tokens2 = tokenize(text2);

  const similarity = jaccardSimilarity(tokens1, tokens2);
  Logger.log(`類似度: ${similarity}`);
  Logger.log(`トークン1: ${[...tokens1].join(', ')}`);
  Logger.log(`トークン2: ${[...tokens2].join(', ')}`);
}

/**
 * 設定確認（詳細情報付き）
 */
function checkConfig() {
  const props = PropertiesService.getScriptProperties().getProperties();

  Logger.log('=== 現在の設定 ===');
  Logger.log('');

  // 設定値の表示
  const displayKeys = [
    'ROOT_FOLDER_ID',
    'INBOX_FOLDER_ID',
    'DRAFTS_FOLDER_ID',
    'LOGS_FOLDER_ID',
    'PROCESSED_FOLDER_ID',
    'QA_SHEET_ID',
    'OPENAI_API_KEY'
  ];

  for (const key of displayKeys) {
    const value = props[key];
    if (key === 'OPENAI_API_KEY') {
      Logger.log(`  ${key}: ${value ? '(設定済み)' : '(未設定)'}`);
    } else {
      Logger.log(`  ${key}: ${value || '(未設定)'}`);
    }
  }

  // リソースのアクセス確認
  Logger.log('');
  Logger.log('=== リソースアクセス確認 ===');

  const checks = [
    { key: 'INBOX_FOLDER_ID', type: 'folder', name: 'inbox' },
    { key: 'DRAFTS_FOLDER_ID', type: 'folder', name: 'drafts' },
    { key: 'LOGS_FOLDER_ID', type: 'folder', name: 'logs' },
    { key: 'QA_SHEET_ID', type: 'sheet', name: '過去QA' }
  ];

  let allOk = true;
  for (const check of checks) {
    const id = props[check.key];
    if (!id) {
      Logger.log(`  ${check.name}: 未設定`);
      allOk = false;
      continue;
    }

    try {
      if (check.type === 'folder') {
        const folder = DriveApp.getFolderById(id);
        Logger.log(`  ${check.name}: OK (${folder.getName()})`);
      } else {
        const sheet = SpreadsheetApp.openById(id);
        Logger.log(`  ${check.name}: OK (${sheet.getName()})`);
      }
    } catch (e) {
      Logger.log(`  ${check.name}: エラー - ${e.message}`);
      allOk = false;
    }
  }

  // トリガー確認
  Logger.log('');
  Logger.log('=== トリガー状態 ===');
  const triggers = ScriptApp.getProjectTriggers();
  const relevantTriggers = triggers.filter(t => t.getHandlerFunction() === 'processInboxFiles');

  if (relevantTriggers.length === 0) {
    Logger.log('  時間トリガー: 未設定');
  } else {
    relevantTriggers.forEach(t => {
      Logger.log(`  時間トリガー: 設定済み`);
    });
  }

  Logger.log('');
  if (allOk) {
    Logger.log('セットアップ状態: 正常');
  } else {
    Logger.log('セットアップ状態: 要確認（setup() を実行してください）');
  }

  return allOk;
}

// ===========================================
// モデル設定ヘルパー
// ===========================================

/**
 * OpenAIのモデルを設定（Script Properties: OPENAI_MODEL）
 * 例: setOpenAIModel('gpt-5') または setOpenAIModel('gpt-4o')
 * @param {string} model - モデルID（例: 'gpt-5', 'gpt-4o', 'gpt-4o-mini'）
 */
function setOpenAIModel(model) {
  if (!model || typeof model !== 'string') {
    Logger.log('エラー: 有効なモデル名を指定してください');
    return;
  }
  PropertiesService.getScriptProperties().setProperty('OPENAI_MODEL', model);
  Logger.log(`OPENAI_MODEL を '${model}' に設定しました`);
}

/**
 * 現在のモデル設定をログに表示
 */
function showOpenAIModel() {
  const model = PropertiesService.getScriptProperties().getProperty('OPENAI_MODEL') || '(未設定: 既定=gpt-4o-mini)';
  Logger.log(`現在のOPENAI_MODEL: ${model}`);
}

/**
 * モデル動作の簡易テスト（APIキー必須）
 * 実行後、Apps Scriptの実行ログで結果を確認してください。
 */
function testOpenAIModel() {
  try {
    if (!CONFIG.OPENAI_API_KEY) {
      Logger.log('エラー: OPENAI_API_KEY が未設定です');
      return;
    }
    const systemPrompt = 'You are a helpful assistant.';
    const userPrompt = '返信はOKのみ。';
    const response = UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions', {
      method: 'post',
      contentType: 'application/json',
      headers: { 'Authorization': `Bearer ${CONFIG.OPENAI_API_KEY}` },
      payload: JSON.stringify({
        model: CONFIG.OPENAI_MODEL,
        messages: [ { role: 'system', content: systemPrompt }, { role: 'user', content: userPrompt } ],
        max_tokens: 3,
        temperature: 0
      }),
      muteHttpExceptions: true
    });

    const status = response.getResponseCode();
    const text = response.getContentText();
    Logger.log(`HTTP ${status}`);

    let body;
    try { body = JSON.parse(text); } catch (e) {}

    if (status >= 200 && status < 300 && body && body.choices && body.choices[0]?.message?.content) {
      Logger.log('呼び出し成功');
      Logger.log(`content: ${body.choices[0].message.content}`);
    } else {
      const msg = body?.error?.message || text;
      Logger.log('呼び出し失敗');
      Logger.log(`error: ${msg}`);
      Logger.log('注: モデル未開放、またはエンドポイント非対応の可能性があります');
    }
  } catch (e) {
    Logger.log('テスト中エラー: ' + e.message);
  }
}
