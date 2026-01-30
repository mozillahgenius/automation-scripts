// =========================================
// Google フォーム 回答通知スクリプト
// 個人向け顧問サービス相談申込フォーム用
// =========================================

// 通知先メールアドレス
const NOTIFICATION_EMAIL = "your-email@example.com";

// メール件名
const EMAIL_SUBJECT = "【新規申込】個人向け顧問サービス相談申込がありました";

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

  let body = `
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
　個人向け顧問サービス相談申込がありました
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

【受付日時】
${getValue("タイムスタンプ")}

────────────────────────────────
■ お客様情報
────────────────────────────────
お名前: ${getValue("Q1. お名前（必須）")}
ご本人の立場: ${getValue("Q2. ご本人の立場（必須）")}
連絡用メールアドレス: ${getValue("Q8. 連絡用メールアドレス（必須）")}

────────────────────────────────
■ ご相談内容
────────────────────────────────
相談の分野: ${getValue("Q3. ご相談の大枠の分野（必須）")}

【現在の状況】
${getValue("Q4. 現在の状況について（必須）")}

判断が必要な時期: ${getValue("Q5. いつ頃までに判断が必要ですか（必須）")}

────────────────────────────────
■ サービス理解・ご希望
────────────────────────────────
サービス理解確認: ${getValue("Q6. 本サービスの理解確認（必須）")}
ご希望の顧問料レンジ: ${getValue("Q7. ご希望の顧問料レンジ（必須）")}

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

  const escapeHtml = (text) => {
    return text
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/\n/g, "<br>");
  };

  // 緊急度に応じた色を設定
  const urgency = getValue("Q5. いつ頃までに判断が必要ですか（必須）");
  let urgencyColor = "#667eea";
  let urgencyBadge = "";
  if (urgency.includes("緊急")) {
    urgencyColor = "#e53e3e";
    urgencyBadge = '<span style="background: #e53e3e; color: white; padding: 2px 8px; border-radius: 4px; font-size: 11px; margin-left: 8px;">緊急</span>';
  }

  return `
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <style>
    body { font-family: 'Helvetica Neue', Arial, 'Hiragino Sans', sans-serif; line-height: 1.6; color: #333; }
    .container { max-width: 600px; margin: 0 auto; padding: 20px; }
    .header { background: linear-gradient(135deg, #2d3748 0%, #4a5568 100%); color: white; padding: 20px; border-radius: 8px 8px 0 0; text-align: center; }
    .header h1 { margin: 0; font-size: 20px; }
    .content { background: #f9f9f9; padding: 20px; border: 1px solid #e0e0e0; }
    .section { background: white; margin-bottom: 15px; padding: 15px; border-radius: 6px; box-shadow: 0 1px 3px rgba(0,0,0,0.1); }
    .section-title { color: #2d3748; font-weight: bold; font-size: 14px; margin-bottom: 10px; border-bottom: 2px solid #4a5568; padding-bottom: 5px; }
    .field { margin-bottom: 8px; }
    .field-label { font-weight: bold; color: #555; font-size: 12px; }
    .field-value { color: #333; margin-top: 2px; }
    .long-text { background: #f5f5f5; padding: 10px; border-radius: 4px; white-space: pre-wrap; }
    .footer { text-align: center; padding: 15px; color: #888; font-size: 11px; }
    .timestamp { background: #4a5568; color: white; display: inline-block; padding: 5px 10px; border-radius: 4px; font-size: 12px; }
    .tag { display: inline-block; background: #e2e8f0; color: #4a5568; padding: 2px 8px; border-radius: 4px; font-size: 11px; margin: 2px; }
    .price-tag { background: #48bb78; color: white; padding: 4px 12px; border-radius: 4px; font-weight: bold; }
  </style>
</head>
<body>
  <div class="container">
    <div class="header">
      <h1>個人向け顧問サービス相談申込${urgencyBadge}</h1>
    </div>
    <div class="content">
      <p style="text-align: center;">
        <span class="timestamp">受付日時: ${escapeHtml(getValue("タイムスタンプ"))}</span>
      </p>

      <div class="section">
        <div class="section-title">お客様情報</div>
        <div class="field">
          <div class="field-label">お名前</div>
          <div class="field-value">${escapeHtml(getValue("Q1. お名前（必須）"))}</div>
        </div>
        <div class="field">
          <div class="field-label">ご本人の立場</div>
          <div class="field-value">${escapeHtml(getValue("Q2. ご本人の立場（必須）"))}</div>
        </div>
        <div class="field">
          <div class="field-label">連絡用メールアドレス</div>
          <div class="field-value"><a href="mailto:${escapeHtml(getValue("Q8. 連絡用メールアドレス（必須）"))}">${escapeHtml(getValue("Q8. 連絡用メールアドレス（必須）"))}</a></div>
        </div>
      </div>

      <div class="section">
        <div class="section-title">ご相談内容</div>
        <div class="field">
          <div class="field-label">相談の分野</div>
          <div class="field-value">${formatTags(getValue("Q3. ご相談の大枠の分野（必須）"))}</div>
        </div>
        <div class="field">
          <div class="field-label">現在の状況</div>
          <div class="field-value long-text">${escapeHtml(getValue("Q4. 現在の状況について（必須）"))}</div>
        </div>
        <div class="field">
          <div class="field-label">判断が必要な時期</div>
          <div class="field-value" style="color: ${urgencyColor}; font-weight: bold;">${escapeHtml(getValue("Q5. いつ頃までに判断が必要ですか（必須）"))}</div>
        </div>
      </div>

      <div class="section">
        <div class="section-title">サービス理解・ご希望</div>
        <div class="field">
          <div class="field-label">サービス理解確認</div>
          <div class="field-value">${formatCheckItems(getValue("Q6. 本サービスの理解確認（必須）"))}</div>
        </div>
        <div class="field">
          <div class="field-label">ご希望の顧問料レンジ</div>
          <div class="field-value"><span class="price-tag">${escapeHtml(getValue("Q7. ご希望の顧問料レンジ（必須）"))}</span></div>
        </div>
      </div>
    </div>
    <div class="footer">
      このメールは自動送信されています<br>
      フォーム設置ページ: <a href="https://well-being.life/personal-advisory-service">https://well-being.life/personal-advisory-service</a>
    </div>
  </div>
</body>
</html>
`;
}

/**
 * カンマ区切りの値をタグ形式にフォーマット
 * @param {string} value - カンマ区切りの値
 * @returns {string} HTML形式のタグ
 */
function formatTags(value) {
  if (value === "未回答") return "未回答";
  const items = value.split(", ");
  return items.map(item => `<span class="tag">${item}</span>`).join(" ");
}

/**
 * チェック項目をリスト形式にフォーマット
 * @param {string} value - カンマ区切りの値
 * @returns {string} HTML形式のリスト
 */
function formatCheckItems(value) {
  if (value === "未回答") return "未回答";
  const items = value.split(", ");
  return items.map(item => `<div style="margin: 4px 0;"><span style="color: #48bb78;">✓</span> ${item}</div>`).join("");
}

/**
 * お客様への自動返信メール送信（任意機能）
 * @param {Object} responses - フォーム回答データ
 */
function sendAutoReply(responses) {
  const customerEmail = responses["Q8. 連絡用メールアドレス（必須）"];
  const customerName = responses["Q1. お名前（必須）"];

  if (!customerEmail || customerEmail[0] === "") return;

  const subject = "【自動返信】顧問サービス相談申込を受け付けました";
  const body = `
${customerName ? customerName[0] : "お客"} 様

この度は個人向け顧問サービスへの相談申込をいただき、
誠にありがとうございます。

お申し込み内容を確認の上、担当者より
2営業日以内にご連絡させていただきます。

※ 本サービスは、判断の整理・意思決定支援を目的とした顧問サービスです。
※ 内容によっては、顧問契約をお受けできない場合がございます。

しばらくお待ちくださいますようお願い申し上げます。

━━━━━━━━━━━━━━━━━━━━━━━━
Intelligent Beast
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
      "タイムスタンプ": ["2026/01/17 22:00:00"],
      "メールアドレス": ["form@example.com"],
      "Q1. お名前（必須）": ["テスト 太郎"],
      "Q2. ご本人の立場（必須）": ["旦那様"],
      "Q3. ご相談の大枠の分野（必須）": ["移住・海外関連, 教育・進学・学校選択, ラブアン法人運営"],
      "Q4. 現在の状況について（必須）": ["マレーシアへの移住を検討しています。\n子供の教育とラブアン法人の設立について相談したいです。\n現在は日本在住で、半年以内に移住を考えています。"],
      "Q5. いつ頃までに判断が必要ですか（必須）": ["緊急（2週間以内）"],
      "Q6. 本サービスの理解確認（必須）": ["実務代行サービスではないことを理解しています, 感情的な相談・慰めを主目的としないことを理解しています, 判断・契約・金銭の最終責任は自分にあることを理解しています"],
      "Q7. ご希望の顧問料レンジ（必須）": ["月額5万円"],
      "Q8. 連絡用メールアドレス（必須）": ["test@example.com"]
    }
  };

  onFormSubmit(testEvent);
  console.log("テストメールを送信しました");
}
