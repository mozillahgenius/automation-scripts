/**
 * 標的型メール訓練システム - 完全統合版
 * 全機能を含む完全なシステム実装
 */

// ==================== 初期設定・メニュー ====================

/**
 * スプレッドシート開いた時の初期化
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🎯 標的型メール訓練')
    .addItem('📋 初期設定', 'initializeSpreadsheet')
    .addSeparator()
    .addItem('📝 キャンペーン本文生成', 'generateCampaignMails')
    .addItem('✉️ テスト送信（自分宛）', 'sendTestMail')
    .addItem('🚀 本送信開始', 'startCampaignSending')
    .addSeparator()
    .addItem('📊 集計更新', 'updateResults')
    .addItem('📈 ダッシュボード表示', 'showDashboard')
    .addSeparator()
    .addItem('🔗 WebApp URL表示', 'showWebAppUrl')
    .addItem('⚙️ 設定確認', 'checkConfiguration')
    .addToUi();
}

/**
 * スプレッドシートを初期化
 */
function initializeSpreadsheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 既存のシートを削除（初回実行時）
  const sheets = ss.getSheets();
  if (sheets.length === 1 && sheets[0].getName() === 'シート1') {
    sheets[0].setName('Config');
  }

  // 各シートを作成
  createConfigSheet(ss);
  createTargetsSheet(ss);
  createCampaignsSheet(ss);
  createMailsSheet(ss);
  createClicksSheet(ss);
  createResultsSheet(ss);

  // スクリプトプロパティの初期設定
  initializeScriptProperties();

  SpreadsheetApp.getUi().alert('スプレッドシートの初期化が完了しました。');
}

// ==================== シート作成関数 ====================

/**
 * Configシート作成
 */
function createConfigSheet(ss) {
  let sheet = ss.getSheetByName('Config');
  if (!sheet) {
    sheet = ss.insertSheet('Config');
  }

  sheet.clear();

  const headers = ['Key', 'Value', '備考'];
  const data = [
    ['PPLX_API_KEY', '', 'Perplexity APIキー（Script Propertiesでも設定可）'],
    ['COMPANY_URLS', 'https://example.com,https://example.com/news', '自社/関連/採用ページURL（カンマ区切り）'],
    ['SENDER_ALIAS', 'alerts@training.example.com', '許諾済みエイリアス'],
    ['SENDER_NAME_WEAK_SUS', 'exarnple HR', 'わずかな不審性（rn/m混同など）'],
    ['REPLY_TO', 'no-reply@training.example.com', 'Reply-To設定'],
    ['LANDING_PAGE_URL', '', 'GAS WebApp URL（デプロイ後に設定）'],
    ['EXPLAIN_LP_URL', 'https://intra.example.com/security/phishing-lesson', '教育ページURL'],
    ['RATE_LIMIT_PER_MIN', '80', '1分あたりの送信上限'],
    ['PIXEL_TRACKING', 'FALSE', '開封率トラッキング（TRUE/FALSE）'],
    ['TOKEN_LENGTH', '24', 'URLトークン長'],
    ['MAIL_DIFFICULTY_LEVELS', 'Easy,Std,Hard', '難易度レベル'],
    ['AUTO_EXPLAIN_ENABLED', 'TRUE', '自動解説メール送信（TRUE/FALSE）']
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#f0f0f0');
  sheet.getRange(2, 1, data.length, 3).setValues(data);
  sheet.autoResizeColumns(1, 3);

  // 条件付き書式（重要項目をハイライト）
  const range = sheet.getRange('A2:A13');
  const rule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('PPLX_API_KEY')
    .setBackground('#fff3cd')
    .setRanges([range])
    .build();
  sheet.setConditionalFormatRules([rule]);
}

/**
 * Targetsシート作成
 */
function createTargetsSheet(ss) {
  let sheet = ss.getSheetByName('Targets');
  if (!sheet) {
    sheet = ss.insertSheet('Targets');
  }

  sheet.clear();

  const headers = ['enabled', 'email', 'name', 'dept', 'title', 'uid', 'manager_email'];
  const sampleData = [
    ['TRUE', 'tanaka@example.com', '田中太郎', '営業部', '課長', 'U001', 'suzuki@example.com'],
    ['TRUE', 'yamada@example.com', '山田花子', '人事部', '主任', 'U002', 'sato@example.com'],
    ['FALSE', 'test@example.com', 'テストユーザー', 'IT部', 'エンジニア', 'U999', '']
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#f0f0f0');
  sheet.getRange(2, 1, sampleData.length, headers.length).setValues(sampleData);
  sheet.autoResizeColumns(1, headers.length);

  // データ検証（enabled列）
  const enabledRange = sheet.getRange('A2:A1000');
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['TRUE', 'FALSE'], true)
    .setAllowInvalid(false)
    .build();
  enabledRange.setDataValidation(rule);
}

/**
 * Campaignsシート作成
 */
function createCampaignsSheet(ss) {
  let sheet = ss.getSheetByName('Campaigns');
  if (!sheet) {
    sheet = ss.insertSheet('Campaigns');
  }

  sheet.clear();

  const headers = ['campaign_id', 'name', 'scheduled_at', 'difficulty', 'persona_hint', 'prompt_template', 'suspicious_flags', 'status'];
  const sampleData = [
    [
      'C001',
      '2025Q1_標的型訓練',
      '2025/01/15 10:00',
      'Easy',
      '新入社員向け',
      '会社の最新ニュースを基に、部署{dept}の{title}向けにメールを作成。緊急性を演出。',
      'homoglyph_from,time_pressure',
      'draft'
    ]
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#f0f0f0');
  sheet.getRange(2, 1, sampleData.length, headers.length).setValues(sampleData);
  sheet.autoResizeColumns(1, headers.length);

  // データ検証
  const difficultyRange = sheet.getRange('D2:D1000');
  const difficultyRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Easy', 'Std', 'Hard'], true)
    .build();
  difficultyRange.setDataValidation(difficultyRule);

  const statusRange = sheet.getRange('H2:H1000');
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['draft', 'ready', 'sent', 'completed'], true)
    .build();
  statusRange.setDataValidation(statusRule);
}

/**
 * Mailsシート作成
 */
function createMailsSheet(ss) {
  let sheet = ss.getSheetByName('Mails');
  if (!sheet) {
    sheet = ss.insertSheet('Mails');
  }

  sheet.clear();

  const headers = ['campaign_id', 'uid', 'email', 'token', 'subject', 'body_html', 'body_text', 'send_status', 'sent_at', 'explained'];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#f0f0f0');
  sheet.autoResizeColumns(1, headers.length);
}

/**
 * Clicksシート作成
 */
function createClicksSheet(ss) {
  let sheet = ss.getSheetByName('Clicks');
  if (!sheet) {
    sheet = ss.insertSheet('Clicks');
  }

  sheet.clear();

  const headers = ['ts', 'campaign_id', 'uid', 'email', 'token', 'ua', 'ip_hash', 'referer'];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#f0f0f0');
  sheet.autoResizeColumns(1, headers.length);
}

/**
 * Resultsシート作成
 */
function createResultsSheet(ss) {
  let sheet = ss.getSheetByName('Results');
  if (!sheet) {
    sheet = ss.insertSheet('Results');
  }

  sheet.clear();

  const headers = ['campaign_id', 'campaign_name', '送信数', 'クリック数', 'クリック率', '部署別クリック率', '役職別クリック率', '更新日時'];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#f0f0f0');
  sheet.autoResizeColumns(1, headers.length);
}

// ==================== メール生成機能 ====================

/**
 * キャンペーンメール生成のメイン処理
 */
function generateCampaignMails() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // キャンペーン選択
  const campaignSheet = ss.getSheetByName('Campaigns');
  const campaigns = campaignSheet.getDataRange().getValues();

  if (campaigns.length <= 1) {
    ui.alert('エラー', 'Campaignsシートにキャンペーンが登録されていません。', ui.ButtonSet.OK);
    return;
  }

  // キャンペーン選択ダイアログ
  const campaignList = [];
  for (let i = 1; i < campaigns.length; i++) {
    campaignList.push(`${campaigns[i][0]}: ${campaigns[i][1]} (${campaigns[i][3]})`);
  }

  const response = ui.prompt(
    'キャンペーン選択',
    `生成するキャンペーンIDを入力してください：\n\n${campaignList.join('\n')}\n`,
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  const campaignId = response.getResponseText();

  // キャンペーン情報取得
  let campaign = null;
  for (let i = 1; i < campaigns.length; i++) {
    if (campaigns[i][0] === campaignId) {
      campaign = {
        id: campaigns[i][0],
        name: campaigns[i][1],
        scheduledAt: campaigns[i][2],
        difficulty: campaigns[i][3],
        personaHint: campaigns[i][4],
        promptTemplate: campaigns[i][5],
        suspiciousFlags: campaigns[i][6],
        status: campaigns[i][7]
      };
      break;
    }
  }

  if (!campaign) {
    ui.alert('エラー', `キャンペーンID ${campaignId} が見つかりません。`, ui.ButtonSet.OK);
    return;
  }

  // 対象者取得
  const targetsSheet = ss.getSheetByName('Targets');
  const targets = targetsSheet.getDataRange().getValues();
  const enabledTargets = [];

  for (let i = 1; i < targets.length; i++) {
    if (targets[i][0] === 'TRUE') {
      enabledTargets.push({
        email: targets[i][1],
        name: targets[i][2],
        dept: targets[i][3],
        title: targets[i][4],
        uid: targets[i][5],
        managerEmail: targets[i][6]
      });
    }
  }

  if (enabledTargets.length === 0) {
    ui.alert('エラー', '有効な対象者が登録されていません。', ui.ButtonSet.OK);
    return;
  }

  ui.alert('処理開始',
    `${enabledTargets.length}名分のメールを生成します。\n` +
    `キャンペーン: ${campaign.name}\n` +
    `難易度: ${campaign.difficulty}`,
    ui.ButtonSet.OK);

  // メール生成処理
  try {
    const generatedMails = generateMailsForCampaign(campaign, enabledTargets);
    saveMails(generatedMails);

    ui.alert('完了',
      `${generatedMails.length}件のメールを生成しました。\n` +
      'Mailsシートで確認してください。',
      ui.ButtonSet.OK);
  } catch (error) {
    ui.alert('エラー', `メール生成中にエラーが発生しました：\n${error.toString()}`, ui.ButtonSet.OK);
  }
}

/**
 * キャンペーン用メール生成
 */
function generateMailsForCampaign(campaign, targets) {
  const mails = [];
  const apiKey = getConfig('PPLX_API_KEY');
  const companyUrls = getConfig('COMPANY_URLS');
  const landingUrl = getConfig('LANDING_PAGE_URL');

  // 企業情報を取得（Perplexity APIを使用する場合）
  let companyContext = '';
  if (apiKey && companyUrls) {
    companyContext = fetchCompanyContext(apiKey, companyUrls);
  }

  for (const target of targets) {
    const token = generateToken();
    const trackingUrl = `${landingUrl}?c=${campaign.id}&t=${token}`;

    // メール本文を生成
    const mailContent = generatePhishingMailContent(campaign, target, companyContext, trackingUrl);

    mails.push({
      campaignId: campaign.id,
      uid: target.uid,
      email: target.email,
      token: token,
      subject: mailContent.subject,
      bodyHtml: mailContent.bodyHtml,
      bodyText: mailContent.bodyText,
      sendStatus: 'pending',
      sentAt: '',
      explained: 'FALSE'
    });
  }

  return mails;
}

/**
 * フィッシングメール本文生成
 */
function generatePhishingMailContent(campaign, target, companyContext, trackingUrl) {
  const difficulty = campaign.difficulty;
  const suspiciousFlags = campaign.suspiciousFlags ? campaign.suspiciousFlags.split(',') : [];

  // 難易度に応じた件名パターン
  const subjectPatterns = {
    'Easy': [
      '【緊急】システムメンテナンスのお知らせ',
      '【重要】セキュリティアップデートが必要です',
      '【至急確認】あなたのアカウントに異常なアクセスが検出されました'
    ],
    'Std': [
      'パスワードの有効期限が近づいています',
      `${target.name}様 - 福利厚生制度の変更について`,
      '社内システム移行に関するご案内'
    ],
    'Hard': [
      `Re: ${target.dept}の${target.title}様へのご連絡`,
      '先日の件について',
      `${target.name}様 - ご確認をお願いします`
    ]
  };

  const subjects = subjectPatterns[difficulty] || subjectPatterns['Easy'];
  const subject = subjects[Math.floor(Math.random() * subjects.length)];

  // 不審性フラグの適用
  let fromName = '社内システム管理';
  if (suspiciousFlags.includes('homoglyph_from')) {
    fromName = fromName.replace('内', '內'); // 似た文字に置換
  }

  // HTML本文生成
  const bodyHtml = createPhishingHtmlBody(target, subject, trackingUrl, difficulty, suspiciousFlags);

  // テキスト本文生成
  const bodyText = createPhishingTextBody(target, subject, trackingUrl, difficulty);

  return {
    subject: subject,
    bodyHtml: bodyHtml,
    bodyText: bodyText
  };
}

/**
 * フィッシングメールHTML本文作成
 */
function createPhishingHtmlBody(target, subject, trackingUrl, difficulty, suspiciousFlags) {
  const urgency = suspiciousFlags.includes('time_pressure');
  const deadline = urgency ? '本日17:00まで' : '今週中';

  let buttonColor = '#007bff';
  if (difficulty === 'Easy') {
    buttonColor = '#dc3545'; // 赤（緊急感を演出）
  }

  return `
<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8">
  <style>
    body { font-family: 'メイリオ', sans-serif; line-height: 1.6; }
    .container { max-width: 600px; margin: 0 auto; padding: 20px; }
    .header { background: #f0f0f0; padding: 15px; border-radius: 5px; }
    .content { margin: 20px 0; }
    .button {
      display: inline-block;
      padding: 12px 35px;
      background: ${buttonColor};
      color: white;
      text-decoration: none;
      border-radius: 5px;
      font-weight: bold;
    }
    .footer { margin-top: 30px; font-size: 0.9em; color: #666; }
    ${urgency ? '.urgent { color: red; font-weight: bold; }' : ''}
  </style>
</head>
<body>
  <div class="container">
    <div class="header">
      <h2>${subject}</h2>
    </div>

    <div class="content">
      <p>${target.name} 様</p>
      <p>${target.dept}の${target.title}として、重要な確認事項があります。</p>

      ${urgency ? '<p class="urgent">※緊急対応が必要です</p>' : ''}

      <p>詳細は以下のリンクからご確認ください。</p>
      <p><strong>期限：${deadline}</strong></p>

      <p style="text-align: center; margin: 30px 0;">
        <a href="${trackingUrl}" class="button">確認する</a>
      </p>

      <p>ご不明な点がございましたら、IT管理部までお問い合わせください。</p>
    </div>

    <div class="footer">
      <p>社内システム管理部</p>
      ${suspiciousFlags.includes('typo') ? '<p>※このメールは自動送信でs。</p>' : '<p>※このメールは自動送信です。</p>'}
    </div>
  </div>
</body>
</html>
`;
}

/**
 * フィッシングメールテキスト本文作成
 */
function createPhishingTextBody(target, subject, trackingUrl, difficulty) {
  return `${target.name} 様

${target.dept}の${target.title}として、重要な確認事項があります。

詳細は以下のリンクからご確認ください：
${trackingUrl}

ご不明な点がございましたら、IT管理部までお問い合わせください。

社内システム管理部
`;
}

/**
 * メール保存
 */
function saveMails(mails) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mailsSheet = ss.getSheetByName('Mails');

  const data = mails.map(mail => [
    mail.campaignId,
    mail.uid,
    mail.email,
    mail.token,
    mail.subject,
    mail.bodyHtml,
    mail.bodyText,
    mail.sendStatus,
    mail.sentAt,
    mail.explained
  ]);

  if (data.length > 0) {
    const lastRow = mailsSheet.getLastRow();
    mailsSheet.getRange(lastRow + 1, 1, data.length, 10).setValues(data);
  }
}

// ==================== メール送信機能 ====================

/**
 * テスト送信
 */
function sendTestMail() {
  const ui = SpreadsheetApp.getUi();
  const userEmail = Session.getActiveUser().getEmail();

  const response = ui.alert(
    'テスト送信確認',
    `${userEmail} 宛にテストメールを送信します。\nよろしいですか？`,
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    return;
  }

  try {
    // サンプルメール作成
    const testMail = {
      to: userEmail,
      subject: '[テスト] 重要：システムアップデートのお知らせ',
      body: createTestMailBody(),
      htmlBody: createTestMailHtmlBody()
    };

    // 送信
    sendPhishingMail(testMail);

    ui.alert('完了', 'テストメールを送信しました。', ui.ButtonSet.OK);
  } catch (error) {
    ui.alert('エラー', `送信エラー：${error.toString()}`, ui.ButtonSet.OK);
  }
}

/**
 * テストメール本文作成
 */
function createTestMailBody() {
  const token = generateToken();
  const landingUrl = getConfig('LANDING_PAGE_URL');
  const trackingUrl = `${landingUrl}?c=TEST&t=${token}`;

  return `社員各位

システムの重要なセキュリティアップデートを実施いたします。

本日17:00までに、以下のリンクから確認をお願いします：
${trackingUrl}

※このメールは標的型メール訓練のテストメールです。

IT管理部
`;
}

/**
 * テストメールHTML本文作成
 */
function createTestMailHtmlBody() {
  const token = generateToken();
  const landingUrl = getConfig('LANDING_PAGE_URL');
  const trackingUrl = `${landingUrl}?c=TEST&t=${token}`;

  return `
<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8">
  <style>
    body { font-family: 'メイリオ', sans-serif; line-height: 1.6; }
    .container { max-width: 600px; margin: 0 auto; padding: 20px; }
    .header { background: #f0f0f0; padding: 15px; border-radius: 5px; }
    .content { margin: 20px 0; }
    .button {
      display: inline-block;
      padding: 10px 30px;
      background: #007bff;
      color: white;
      text-decoration: none;
      border-radius: 5px;
    }
    .footer { margin-top: 30px; font-size: 0.9em; color: #666; }
    .warning {
      background: #fff3cd;
      border: 1px solid #ffc107;
      padding: 10px;
      margin-top: 20px;
      border-radius: 5px;
    }
  </style>
</head>
<body>
  <div class="container">
    <div class="header">
      <h2>重要：システムアップデートのお知らせ</h2>
    </div>

    <div class="content">
      <p>社員各位</p>

      <p>システムの重要なセキュリティアップデートを実施いたします。</p>

      <p><strong>本日17:00まで</strong>に、以下のリンクから確認をお願いします：</p>

      <p style="text-align: center;">
        <a href="${trackingUrl}" class="button">確認はこちら</a>
      </p>
    </div>

    <div class="warning">
      <strong>⚠️ 注意：</strong>このメールは標的型メール訓練のテストメールです。
    </div>

    <div class="footer">
      <p>IT管理部</p>
    </div>
  </div>
</body>
</html>
`;
}

/**
 * フィッシングメール送信
 */
function sendPhishingMail(mailData) {
  const senderAlias = getConfig('SENDER_ALIAS');
  const replyTo = getConfig('REPLY_TO');

  const options = {
    htmlBody: mailData.htmlBody,
    name: mailData.fromName || '社内システム管理'
  };

  if (senderAlias) {
    options.from = senderAlias;
  }

  if (replyTo) {
    options.replyTo = replyTo;
  }

  GmailApp.sendEmail(
    mailData.to,
    mailData.subject,
    mailData.body || mailData.bodyText,
    options
  );
}

/**
 * 本送信開始
 */
function startCampaignSending() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 未送信メール確認
  const mailsSheet = ss.getSheetByName('Mails');
  const mails = mailsSheet.getDataRange().getValues();

  const pendingMails = [];
  for (let i = 1; i < mails.length; i++) {
    if (mails[i][7] !== 'sent') {
      pendingMails.push(i);
    }
  }

  if (pendingMails.length === 0) {
    ui.alert('情報', '送信待ちのメールがありません。', ui.ButtonSet.OK);
    return;
  }

  const response = ui.alert(
    '送信確認',
    `${pendingMails.length}件のメールを送信します。\nよろしいですか？`,
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    return;
  }

  // バッチ送信実行
  const result = batchSendMails(pendingMails);

  ui.alert('送信完了',
    `送信完了: ${result.success}件\n` +
    `送信失敗: ${result.failed}件\n` +
    `詳細はMailsシートを確認してください。`,
    ui.ButtonSet.OK);
}

/**
 * バッチメール送信
 */
function batchSendMails(mailIndices) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mailsSheet = ss.getSheetByName('Mails');
  const mails = mailsSheet.getDataRange().getValues();

  const rateLimit = parseInt(getConfig('RATE_LIMIT_PER_MIN') || '80');
  const delayMs = Math.ceil(60000 / rateLimit);

  let successCount = 0;
  let failedCount = 0;

  for (const index of mailIndices) {
    const mail = mails[index];

    try {
      const mailData = {
        to: mail[2], // email
        subject: mail[4], // subject
        bodyText: mail[6], // body_text
        htmlBody: mail[5] // body_html
      };

      sendPhishingMail(mailData);

      // 送信ステータス更新
      mailsSheet.getRange(index + 1, 8).setValue('sent');
      mailsSheet.getRange(index + 1, 9).setValue(new Date());

      successCount++;

      // レート制限のための待機
      Utilities.sleep(delayMs);

    } catch (error) {
      console.error(`送信エラー (${mail[2]}):`, error);
      mailsSheet.getRange(index + 1, 8).setValue('failed');
      failedCount++;
    }
  }

  return {
    success: successCount,
    failed: failedCount
  };
}

// ==================== WebApp機能（クリック追跡） ====================

/**
 * WebApp GET処理
 */
function doGet(e) {
  const params = e.parameter;
  const token = params.t;
  const campaignId = params.c;

  if (!token || !campaignId) {
    return createEducationalPage('パラメータが不足しています');
  }

  // クリック記録
  recordClick(campaignId, token, e);

  // 教育ページを返す
  return createEducationalPage();
}

/**
 * クリック記録
 */
function recordClick(campaignId, token, request) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const clicksSheet = ss.getSheetByName('Clicks');
  const mailsSheet = ss.getSheetByName('Mails');

  // トークンからメール情報を取得
  const mails = mailsSheet.getDataRange().getValues();
  let targetMail = null;

  for (let i = 1; i < mails.length; i++) {
    if (mails[i][3] === token && mails[i][0] === campaignId) {
      targetMail = {
        uid: mails[i][1],
        email: mails[i][2],
        rowIndex: i + 1
      };
      break;
    }
  }

  if (!targetMail) {
    console.error('Invalid token or campaign ID');
    return;
  }

  // クリック情報を記録
  const clickData = [
    new Date(),
    campaignId,
    targetMail.uid,
    targetMail.email,
    token,
    request.parameter['user-agent'] || 'unknown',
    hashIP(request.parameter['x-forwarded-for'] || 'unknown'),
    request.parameter.referer || ''
  ];

  clicksSheet.appendRow(clickData);

  // 自動解説メール送信（有効な場合）
  if (getConfig('AUTO_EXPLAIN_ENABLED') === 'TRUE') {
    sendEducationalMail(targetMail.email, campaignId);
    mailsSheet.getRange(targetMail.rowIndex, 10).setValue('TRUE');
  }
}

/**
 * 教育ページ作成
 */
function createEducationalPage(errorMessage) {
  const explainUrl = getConfig('EXPLAIN_LP_URL');

  const html = `
<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8">
  <title>標的型メール訓練 - 教育ページ</title>
  <style>
    body {
      font-family: 'メイリオ', sans-serif;
      margin: 0;
      padding: 0;
      background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
      min-height: 100vh;
      display: flex;
      align-items: center;
      justify-content: center;
    }
    .container {
      background: white;
      border-radius: 10px;
      padding: 40px;
      max-width: 600px;
      box-shadow: 0 10px 40px rgba(0,0,0,0.2);
    }
    h1 {
      color: #e74c3c;
      text-align: center;
      margin-bottom: 30px;
    }
    .warning {
      background: #fff3cd;
      border: 2px solid #ffc107;
      border-radius: 5px;
      padding: 20px;
      margin: 20px 0;
    }
    .tips {
      background: #d4edda;
      border: 2px solid #28a745;
      border-radius: 5px;
      padding: 20px;
      margin: 20px 0;
    }
    .tips h3 {
      color: #155724;
      margin-top: 0;
    }
    ul {
      line-height: 1.8;
    }
    .button {
      display: inline-block;
      padding: 12px 30px;
      background: #007bff;
      color: white;
      text-decoration: none;
      border-radius: 5px;
      margin-top: 20px;
    }
    .error {
      color: #e74c3c;
      font-weight: bold;
    }
  </style>
</head>
<body>
  <div class="container">
    <h1>⚠️ 標的型メール訓練</h1>

    ${errorMessage ? `<p class="error">${errorMessage}</p>` : ''}

    <div class="warning">
      <h3>このメールは訓練用のフィッシングメールでした</h3>
      <p>あなたがクリックしたリンクは、標的型メール訓練の一環として送信されたものです。</p>
      <p>実際の攻撃だった場合、以下のリスクがありました：</p>
      <ul>
        <li>個人情報の漏洩</li>
        <li>マルウェア感染</li>
        <li>アカウントの乗っ取り</li>
        <li>金銭的被害</li>
      </ul>
    </div>

    <div class="tips">
      <h3>✅ フィッシングメールを見破るポイント</h3>
      <ul>
        <li><strong>送信元を確認：</strong>差出人のメールアドレスは正規のものか？</li>
        <li><strong>緊急性の演出：</strong>「至急」「今すぐ」などの文言に注意</li>
        <li><strong>リンク先の確認：</strong>リンクにカーソルを合わせてURLを確認</li>
        <li><strong>日本語の違和感：</strong>不自然な日本語や誤字脱字がないか</li>
        <li><strong>添付ファイル：</strong>予期しない添付ファイルは開かない</li>
      </ul>
    </div>

    <p>不審なメールを受信した場合は、IT部門に報告してください。</p>

    ${explainUrl ? `<a href="${explainUrl}" class="button">詳しい対策を学ぶ</a>` : ''}
  </div>
</body>
</html>
`;

  return HtmlService.createHtmlOutput(html);
}

// ==================== 教育メール送信機能 ====================

/**
 * 教育メール送信
 */
function sendEducationalMail(email, campaignId) {
  const subject = '【重要】標的型メール訓練の結果について';

  const body = `
先ほどクリックされたメールは、セキュリティ訓練の一環として送信された
標的型メール（フィッシングメール）でした。

実際の攻撃であった場合、以下のような被害が発生する可能性がありました：

1. 個人情報の漏洩
2. マルウェア感染による情報流出
3. アカウントの乗っ取り
4. 金銭的被害

今後、不審なメールを受信した際は、以下の点にご注意ください：

✓ 送信元メールアドレスの確認
✓ 緊急性を装う文言への警戒
✓ リンク先URLの事前確認
✓ 不自然な日本語の有無
✓ 予期しない添付ファイルを開かない

ご不明な点がございましたら、IT部門までお問い合わせください。

IT管理部
`;

  const htmlBody = createEducationalHtmlMail(campaignId);

  try {
    GmailApp.sendEmail(email, subject, body, {
      htmlBody: htmlBody,
      name: 'セキュリティ訓練事務局'
    });
  } catch (error) {
    console.error('教育メール送信エラー:', error);
  }
}

/**
 * 教育メールHTML作成
 */
function createEducationalHtmlMail(campaignId) {
  const explainUrl = getConfig('EXPLAIN_LP_URL');

  return `
<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8">
  <style>
    body { font-family: 'メイリオ', sans-serif; line-height: 1.6; color: #333; }
    .container { max-width: 600px; margin: 0 auto; padding: 20px; }
    .header {
      background: #dc3545;
      color: white;
      padding: 20px;
      border-radius: 5px 5px 0 0;
      text-align: center;
    }
    .content {
      background: #f8f9fa;
      padding: 30px;
      border: 1px solid #dee2e6;
      border-top: none;
    }
    .warning-box {
      background: #fff3cd;
      border-left: 4px solid #ffc107;
      padding: 15px;
      margin: 20px 0;
    }
    .tips-box {
      background: #d4edda;
      border-left: 4px solid #28a745;
      padding: 15px;
      margin: 20px 0;
    }
    .tips-box h3 { color: #155724; margin-top: 0; }
    ul { line-height: 2; }
    .button {
      display: inline-block;
      padding: 12px 30px;
      background: #007bff;
      color: white;
      text-decoration: none;
      border-radius: 5px;
      margin-top: 20px;
    }
    .footer {
      margin-top: 30px;
      padding-top: 20px;
      border-top: 1px solid #dee2e6;
      font-size: 0.9em;
      color: #6c757d;
    }
  </style>
</head>
<body>
  <div class="container">
    <div class="header">
      <h1>⚠️ 標的型メール訓練の結果</h1>
    </div>

    <div class="content">
      <div class="warning-box">
        <strong>重要：</strong>先ほどクリックされたメールは、セキュリティ訓練の一環として送信された標的型メール（フィッシングメール）でした。
      </div>

      <h2>もし実際の攻撃だったら...</h2>
      <p>以下のような被害が発生する可能性がありました：</p>
      <ul>
        <li>🔓 個人情報やパスワードの漏洩</li>
        <li>🦠 マルウェア感染による機密情報の流出</li>
        <li>👤 アカウントの乗っ取りと不正利用</li>
        <li>💰 金銭的被害や身代金要求</li>
        <li>🏢 会社全体への被害拡大</li>
      </ul>

      <div class="tips-box">
        <h3>✅ 今後の対策ポイント</h3>
        <ul>
          <li><strong>送信元の確認：</strong>メールアドレスのドメインが正規のものか確認</li>
          <li><strong>緊急性への警戒：</strong>「至急」「今すぐ」などの文言に注意</li>
          <li><strong>リンクの検証：</strong>クリック前にURLを確認（ホバーで表示）</li>
          <li><strong>文章の違和感：</strong>不自然な日本語や誤字脱字をチェック</li>
          <li><strong>添付ファイル：</strong>予期しないファイルは開かない</li>
          <li><strong>確認の徹底：</strong>不審な場合は送信元に別途確認</li>
        </ul>
      </div>

      <p><strong>次回は必ず見破れるように、この経験を活かしてください。</strong></p>

      ${explainUrl ? `
      <p style="text-align: center;">
        <a href="${explainUrl}" class="button">詳しい対策方法を学ぶ</a>
      </p>
      ` : ''}

      <div class="footer">
        <p>ご不明な点がございましたら、IT部門までお問い合わせください。</p>
        <p>キャンペーンID: ${campaignId}</p>
      </div>
    </div>
  </div>
</body>
</html>
`;
}

// ==================== 集計・レポート機能 ====================

/**
 * 結果集計更新
 */
function updateResults() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const campaignsSheet = ss.getSheetByName('Campaigns');
  const mailsSheet = ss.getSheetByName('Mails');
  const clicksSheet = ss.getSheetByName('Clicks');
  const resultsSheet = ss.getSheetByName('Results');
  const targetsSheet = ss.getSheetByName('Targets');

  // 既存の結果をクリア
  if (resultsSheet.getLastRow() > 1) {
    resultsSheet.getRange(2, 1, resultsSheet.getLastRow() - 1, 8).clear();
  }

  const campaigns = campaignsSheet.getDataRange().getValues();
  const mails = mailsSheet.getDataRange().getValues();
  const clicks = clicksSheet.getDataRange().getValues();
  const targets = targetsSheet.getDataRange().getValues();

  const results = [];

  // キャンペーンごとに集計
  for (let i = 1; i < campaigns.length; i++) {
    const campaignId = campaigns[i][0];
    const campaignName = campaigns[i][1];

    // 送信数カウント
    let sentCount = 0;
    const sentEmails = new Set();
    for (let j = 1; j < mails.length; j++) {
      if (mails[j][0] === campaignId && mails[j][7] === 'sent') {
        sentCount++;
        sentEmails.add(mails[j][2]); // email
      }
    }

    // クリック数カウント
    const clickedEmails = new Set();
    const deptClicks = {};
    const titleClicks = {};

    for (let j = 1; j < clicks.length; j++) {
      if (clicks[j][1] === campaignId) {
        clickedEmails.add(clicks[j][3]); // email

        // 対象者情報を取得
        for (let k = 1; k < targets.length; k++) {
          if (targets[k][1] === clicks[j][3]) { // email match
            const dept = targets[k][3];
            const title = targets[k][4];

            deptClicks[dept] = (deptClicks[dept] || 0) + 1;
            titleClicks[title] = (titleClicks[title] || 0) + 1;
            break;
          }
        }
      }
    }

    const clickCount = clickedEmails.size;
    const clickRate = sentCount > 0 ? (clickCount / sentCount * 100).toFixed(1) + '%' : '0%';

    // 部署別・役職別クリック率
    const deptStats = Object.entries(deptClicks)
      .map(([dept, count]) => `${dept}:${count}`)
      .join(', ');

    const titleStats = Object.entries(titleClicks)
      .map(([title, count]) => `${title}:${count}`)
      .join(', ');

    results.push([
      campaignId,
      campaignName,
      sentCount,
      clickCount,
      clickRate,
      deptStats || '-',
      titleStats || '-',
      new Date()
    ]);
  }

  // 結果を書き込み
  if (results.length > 0) {
    resultsSheet.getRange(2, 1, results.length, 8).setValues(results);
  }

  // 条件付き書式（クリック率による色分け）
  const clickRateRange = resultsSheet.getRange('E2:E' + (results.length + 1));

  // 高クリック率（30%以上）を赤
  const highRule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThanOrEqualTo(30)
    .setBackground('#ffcccc')
    .setRanges([clickRateRange])
    .build();

  // 中クリック率（10-30%）を黄色
  const midRule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberBetween(10, 30)
    .setBackground('#fff3cd')
    .setRanges([clickRateRange])
    .build();

  // 低クリック率（10%未満）を緑
  const lowRule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThan(10)
    .setBackground('#d4edda')
    .setRanges([clickRateRange])
    .build();

  resultsSheet.setConditionalFormatRules([highRule, midRule, lowRule]);

  SpreadsheetApp.getUi().alert('集計完了', '結果を更新しました。', SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * ダッシュボード表示
 */
function showDashboard() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Resultsシートに移動
  const resultsSheet = ss.getSheetByName('Results');
  if (resultsSheet) {
    ss.setActiveSheet(resultsSheet);
    ui.alert('ダッシュボード', 'Resultsシートに集計結果が表示されています。', ui.ButtonSet.OK);
  } else {
    ui.alert('エラー', 'Resultsシートが見つかりません。', ui.ButtonSet.OK);
  }
}

// ==================== ユーティリティ関数 ====================

/**
 * WebApp URLを表示
 */
function showWebAppUrl() {
  const url = ScriptApp.getService().getUrl();
  const ui = SpreadsheetApp.getUi();
  ui.alert('WebApp URL',
    `クリック計測用のWebApp URLは以下です：\n\n${url}\n\n` +
    'このURLをConfigシートのLANDING_PAGE_URLに設定してください。',
    ui.ButtonSet.OK);
}

/**
 * 設定確認
 */
function checkConfiguration() {
  const ui = SpreadsheetApp.getUi();
  const messages = [];

  // 必須設定の確認
  const requiredConfigs = [
    'PPLX_API_KEY',
    'COMPANY_URLS',
    'SENDER_ALIAS',
    'LANDING_PAGE_URL'
  ];

  for (const key of requiredConfigs) {
    const value = getConfig(key);
    if (!value || value === '') {
      messages.push(`❌ ${key} が設定されていません`);
    } else {
      messages.push(`✅ ${key}: 設定済み`);
    }
  }

  // WebAppのデプロイ状態確認
  try {
    const url = ScriptApp.getService().getUrl();
    if (url) {
      messages.push(`✅ WebApp: デプロイ済み`);
    } else {
      messages.push(`❌ WebApp: 未デプロイ`);
    }
  } catch (e) {
    messages.push(`❌ WebApp: 未デプロイ`);
  }

  ui.alert('設定確認', messages.join('\n'), ui.ButtonSet.OK);
}

/**
 * トークン生成
 */
function generateToken() {
  const tokenLength = parseInt(getConfig('TOKEN_LENGTH') || '24');
  const chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
  let token = '';

  for (let i = 0; i < tokenLength; i++) {
    token += chars.charAt(Math.floor(Math.random() * chars.length));
  }

  return token;
}

/**
 * IPアドレスのハッシュ化
 */
function hashIP(ip) {
  if (!ip) return 'unknown';

  // 簡易的なハッシュ化（実運用では適切なハッシュアルゴリズムを使用）
  const hash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, ip);
  const hashStr = hash.map(byte => ('0' + (byte & 0xFF).toString(16)).slice(-2)).join('');

  // 最初の8文字のみ返す（プライバシー配慮）
  return hashStr.substring(0, 8);
}

/**
 * スクリプトプロパティ初期化
 */
function initializeScriptProperties() {
  const properties = PropertiesService.getScriptProperties();

  // 既存のプロパティがない場合のみ設定
  if (!properties.getProperty('INITIALIZED')) {
    properties.setProperties({
      'INITIALIZED': 'true',
      'PPLX_API_KEY': '', // 実際のAPIキーは後で設定
      'VERSION': '1.0.0'
    });
  }
}

/**
 * 設定値を取得するヘルパー関数
 */
function getConfig(key) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName('Config');

  if (!configSheet) {
    throw new Error('Configシートが見つかりません');
  }

  const data = configSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === key) {
      return data[i][1];
    }
  }

  // スクリプトプロパティからも探す
  const scriptValue = PropertiesService.getScriptProperties().getProperty(key);
  if (scriptValue) {
    return scriptValue;
  }

  return null;
}

/**
 * 設定値を保存するヘルパー関数
 */
function setConfig(key, value) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName('Config');

  if (!configSheet) {
    throw new Error('Configシートが見つかりません');
  }

  const data = configSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === key) {
      configSheet.getRange(i + 1, 2).setValue(value);
      return;
    }
  }

  // 新規行として追加
  const lastRow = configSheet.getLastRow();
  configSheet.getRange(lastRow + 1, 1, 1, 3).setValues([[key, value, '']]);
}

/**
 * 企業情報取得（Perplexity API使用）
 * 注：実装にはPerplexity APIの詳細仕様が必要
 */
function fetchCompanyContext(apiKey, urls) {
  // Perplexity APIを使用して企業情報を取得
  // この実装は仕様に応じて調整が必要

  try {
    // APIエンドポイント（仮）
    const endpoint = 'https://api.perplexity.ai/v1/completions';

    const payload = {
      model: 'pplx-70b-online',
      messages: [
        {
          role: 'system',
          content: 'あなたは企業情報を収集するアシスタントです。'
        },
        {
          role: 'user',
          content: `以下のURLから企業の最新ニュースや情報を3つ抽出してください：${urls}`
        }
      ]
    };

    const options = {
      method: 'post',
      headers: {
        'Authorization': `Bearer ${apiKey}`,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(payload)
    };

    const response = UrlFetchApp.fetch(endpoint, options);
    const data = JSON.parse(response.getContentText());

    return data.choices[0].message.content;

  } catch (error) {
    console.error('Perplexity API エラー:', error);
    return '';
  }
}

/**
 * トリガー設定（自動送信用）
 */
function setupTriggers() {
  // 既存のトリガーを削除
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => ScriptApp.deleteTrigger(trigger));

  // 定期実行トリガーを設定（1時間ごと）
  ScriptApp.newTrigger('checkScheduledMails')
    .timeBased()
    .everyHours(1)
    .create();
}

/**
 * スケジュール送信チェック
 */
function checkScheduledMails() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const campaignsSheet = ss.getSheetByName('Campaigns');
  const campaigns = campaignsSheet.getDataRange().getValues();

  const now = new Date();

  for (let i = 1; i < campaigns.length; i++) {
    const scheduledAt = new Date(campaigns[i][2]);
    const status = campaigns[i][7];

    // スケジュール時刻を過ぎていて、ステータスがreadyの場合
    if (scheduledAt <= now && status === 'ready') {
      // 送信処理を実行
      sendCampaignMails(campaigns[i][0]);

      // ステータスを更新
      campaignsSheet.getRange(i + 1, 8).setValue('sent');
    }
  }
}

/**
 * キャンペーンメール送信
 */
function sendCampaignMails(campaignId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mailsSheet = ss.getSheetByName('Mails');
  const mails = mailsSheet.getDataRange().getValues();

  const pendingMails = [];
  for (let i = 1; i < mails.length; i++) {
    if (mails[i][0] === campaignId && mails[i][7] !== 'sent') {
      pendingMails.push(i);
    }
  }

  if (pendingMails.length > 0) {
    batchSendMails(pendingMails);
  }
}