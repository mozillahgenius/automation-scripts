/**
 * 標的型メール訓練システム - メインコード
 * カスタムメニューと主要機能のエントリーポイント
 */

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

  // 最後の8文字のみ返す（プライバシー配慮）
  return hashStr.substring(0, 8);
}