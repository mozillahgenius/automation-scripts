// =========================================
// Google フォーム 回答通知スクリプト
// 教育移住に関するお問い合わせフォーム用
// =========================================

// 通知先メールアドレス
const NOTIFICATION_EMAIL = "your-email@example.com";

// メール件名
const EMAIL_SUBJECT = "【新規問い合わせ】教育移住に関するお問い合わせがありました";

/**
 * フォーム送信時に自動実行される関数
 * @param {Object} e - フォーム送信イベントオブジェクト
 */
function onFormSubmit(e) {
  try {
    const responses = e.namedValues;
    const emailBody = createEmailBody(responses);
    const htmlBody = createHtmlEmailBody(responses);

    MailApp.sendEmail({
      to: NOTIFICATION_EMAIL,
      subject: EMAIL_SUBJECT,
      body: emailBody,
      htmlBody: htmlBody
    });

    // 自動返信メール（任意）
    sendAutoReply(responses);

  } catch (error) {
    console.error("Error: " + error.message);
    // エラー通知
    MailApp.sendEmail({
      to: NOTIFICATION_EMAIL,
      subject: "【エラー】フォーム通知でエラーが発生しました",
      body: "エラー内容: " + error.message
    });
  }
}

/**
 * プレーンテキスト形式のメール本文を作成
 * @param {Object} responses - フォーム回答データ
 * @returns {string} メール本文
 */
function createEmailBody(responses) {
  const getValue = (key) => {
    return responses[key] ? responses[key].join(", ") : "未回答";
  };

  // 移住希望時期をまとめる
  const migrationTiming = [];
  if (getValue("移住希望時期 [半年以内]") !== "未回答") migrationTiming.push("半年以内");
  if (getValue("移住希望時期 [1年以内]") !== "未回答") migrationTiming.push("1年以内");
  if (getValue("移住希望時期 [2年以内]") !== "未回答") migrationTiming.push("2年以内");
  if (getValue("移住希望時期 [未定]") !== "未回答") migrationTiming.push("未定");
  const timingText = migrationTiming.length > 0 ? migrationTiming.join(", ") : "未回答";

  let body = `
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
　新しいお問い合わせが届きました
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

【受付日時】
${getValue("タイムスタンプ")}

────────────────────────────────
■ お客様情報
────────────────────────────────
お名前: ${getValue("お名前 (フルネーム)")}
メールアドレス: ${getValue("メールアドレス(ご都合のよろしい連絡先)")}
電話番号: ${getValue("電話番号 (任意)")}
現在お住まいの国: ${getValue("現在お住まいの国")}

────────────────────────────────
■ 移住・教育について
────────────────────────────────
移住を検討されている国: ${getValue("移住を検討されている国")}
お子様の人数: ${getValue("教育移住の対象となるお子様の人数")}
お子様の学年/年齢: ${getValue("お子様の現在の学年または年齢")}
移住希望時期: ${timingText}

────────────────────────────────
■ お問い合わせ内容
────────────────────────────────
主な目的: ${getValue("お問い合わせの主な目的")}

【ご相談内容の詳細】
${getValue("ご相談内容の詳細をご記入ください。")}

【教育方針・ご希望】
${getValue("家庭内での教育方針・ご希望など(将来お子様にどんな人生にしてもらいたいかなど)")}

────────────────────────────────
■ 流入経路
────────────────────────────────
${getValue("当社のWebサイトをどこでお知りになりましたか？")}

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
※ このメールは自動送信です
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
`;

  return body;
}

/**
 * HTML形式のメール本文を作成
 * @param {Object} responses - フォーム回答データ
 * @returns {string} HTML形式のメール本文
 */
function createHtmlEmailBody(responses) {
  const getValue = (key) => {
    return responses[key] ? responses[key].join(", ") : "未回答";
  };

  // 移住希望時期をまとめる
  const migrationTiming = [];
  if (getValue("移住希望時期 [半年以内]") !== "未回答") migrationTiming.push("半年以内");
  if (getValue("移住希望時期 [1年以内]") !== "未回答") migrationTiming.push("1年以内");
  if (getValue("移住希望時期 [2年以内]") !== "未回答") migrationTiming.push("2年以内");
  if (getValue("移住希望時期 [未定]") !== "未回答") migrationTiming.push("未定");
  const timingText = migrationTiming.length > 0 ? migrationTiming.join(", ") : "未回答";

  const escapeHtml = (text) => {
    return text
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/\n/g, "<br>");
  };

  return `
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <style>
    body { font-family: 'Helvetica Neue', Arial, 'Hiragino Sans', sans-serif; line-height: 1.6; color: #333; }
    .container { max-width: 600px; margin: 0 auto; padding: 20px; }
    .header { background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; padding: 20px; border-radius: 8px 8px 0 0; text-align: center; }
    .header h1 { margin: 0; font-size: 20px; }
    .content { background: #f9f9f9; padding: 20px; border: 1px solid #e0e0e0; }
    .section { background: white; margin-bottom: 15px; padding: 15px; border-radius: 6px; box-shadow: 0 1px 3px rgba(0,0,0,0.1); }
    .section-title { color: #667eea; font-weight: bold; font-size: 14px; margin-bottom: 10px; border-bottom: 2px solid #667eea; padding-bottom: 5px; }
    .field { margin-bottom: 8px; }
    .field-label { font-weight: bold; color: #555; font-size: 12px; }
    .field-value { color: #333; margin-top: 2px; }
    .long-text { background: #f5f5f5; padding: 10px; border-radius: 4px; white-space: pre-wrap; }
    .footer { text-align: center; padding: 15px; color: #888; font-size: 11px; }
    .timestamp { background: #667eea; color: white; display: inline-block; padding: 5px 10px; border-radius: 4px; font-size: 12px; }
  </style>
</head>
<body>
  <div class="container">
    <div class="header">
      <h1>📩 新しいお問い合わせが届きました</h1>
    </div>
    <div class="content">
      <p style="text-align: center;">
        <span class="timestamp">受付日時: ${escapeHtml(getValue("タイムスタンプ"))}</span>
      </p>

      <div class="section">
        <div class="section-title">👤 お客様情報</div>
        <div class="field">
          <div class="field-label">お名前</div>
          <div class="field-value">${escapeHtml(getValue("お名前 (フルネーム)"))}</div>
        </div>
        <div class="field">
          <div class="field-label">メールアドレス</div>
          <div class="field-value"><a href="mailto:${escapeHtml(getValue("メールアドレス(ご都合のよろしい連絡先)"))}">${escapeHtml(getValue("メールアドレス(ご都合のよろしい連絡先)"))}</a></div>
        </div>
        <div class="field">
          <div class="field-label">電話番号</div>
          <div class="field-value">${escapeHtml(getValue("電話番号 (任意)"))}</div>
        </div>
        <div class="field">
          <div class="field-label">現在お住まいの国</div>
          <div class="field-value">${escapeHtml(getValue("現在お住まいの国"))}</div>
        </div>
      </div>

      <div class="section">
        <div class="section-title">🌏 移住・教育について</div>
        <div class="field">
          <div class="field-label">移住を検討されている国</div>
          <div class="field-value">${escapeHtml(getValue("移住を検討されている国"))}</div>
        </div>
        <div class="field">
          <div class="field-label">お子様の人数</div>
          <div class="field-value">${escapeHtml(getValue("教育移住の対象となるお子様の人数"))}</div>
        </div>
        <div class="field">
          <div class="field-label">お子様の学年/年齢</div>
          <div class="field-value">${escapeHtml(getValue("お子様の現在の学年または年齢"))}</div>
        </div>
        <div class="field">
          <div class="field-label">移住希望時期</div>
          <div class="field-value">${escapeHtml(timingText)}</div>
        </div>
      </div>

      <div class="section">
        <div class="section-title">💬 お問い合わせ内容</div>
        <div class="field">
          <div class="field-label">主な目的</div>
          <div class="field-value">${escapeHtml(getValue("お問い合わせの主な目的"))}</div>
        </div>
        <div class="field">
          <div class="field-label">ご相談内容の詳細</div>
          <div class="field-value long-text">${escapeHtml(getValue("ご相談内容の詳細をご記入ください。"))}</div>
        </div>
        <div class="field">
          <div class="field-label">教育方針・ご希望</div>
          <div class="field-value long-text">${escapeHtml(getValue("家庭内での教育方針・ご希望など(将来お子様にどんな人生にしてもらいたいかなど)"))}</div>
        </div>
      </div>

      <div class="section">
        <div class="section-title">📊 流入経路</div>
        <div class="field-value">${escapeHtml(getValue("当社のWebサイトをどこでお知りになりましたか？"))}</div>
      </div>
    </div>
    <div class="footer">
      このメールは自動送信されています
    </div>
  </div>
</body>
</html>
`;
}

/**
 * お客様への自動返信メール送信（任意機能）
 * @param {Object} responses - フォーム回答データ
 */
function sendAutoReply(responses) {
  const customerEmail = responses["メールアドレス(ご都合のよろしい連絡先)"];
  const customerName = responses["お名前 (フルネーム)"];

  if (!customerEmail || customerEmail[0] === "") return;

  const subject = "【自動返信】お問い合わせありがとうございます";
  const body = `
${customerName ? customerName[0] : "お客"} 様

この度は教育移住に関するお問い合わせをいただき、
誠にありがとうございます。

お問い合わせ内容を確認の上、担当者より
2営業日以内にご連絡させていただきます。

しばらくお待ちくださいますようお願い申し上げます。

━━━━━━━━━━━━━━━━━━━━━━━━
COMPANY_X
your-email@example.com
━━━━━━━━━━━━━━━━━━━━━━━━
`;

  MailApp.sendEmail({
    to: customerEmail[0],
    subject: subject,
    body: body
  });
}

/**
 * テスト用関数
 */
function testEmailNotification() {
  const testEvent = {
    namedValues: {
      "タイムスタンプ": ["2026/01/17 10:00:00"],
      "メールアドレス": ["form@example.com"],
      "お名前 (フルネーム)": ["テスト 太郎"],
      "メールアドレス(ご都合のよろしい連絡先)": ["test@example.com"],
      "電話番号 (任意)": ["090-1234-5678"],
      "現在お住まいの国": ["日本"],
      "移住を検討されている国": ["マレーシア、シンガポール"],
      "教育移住の対象となるお子様の人数": ["2人"],
      "お子様の現在の学年または年齢": ["小学3年生、小学1年生"],
      "お問い合わせの主な目的": ["学校の情報収集"],
      "移住希望時期 [半年以内]": [""],
      "移住希望時期 [1年以内]": ["希望"],
      "移住希望時期 [2年以内]": [""],
      "移住希望時期 [未定]": [""],
      "ご相談内容の詳細をご記入ください。": ["マレーシアのインターナショナルスクールについて詳しく知りたいです。\n特にクアラルンプール近郊の学校を検討しています。"],
      "家庭内での教育方針・ご希望など(将来お子様にどんな人生にしてもらいたいかなど)": ["グローバルに活躍できる人材になってほしいと考えています。\n英語力と多様性への理解を身につけてほしいです。"],
      "当社のWebサイトをどこでお知りになりましたか？": ["Google検索"]
    }
  };

  onFormSubmit(testEvent);
  console.log("テストメールを送信しました");
}
