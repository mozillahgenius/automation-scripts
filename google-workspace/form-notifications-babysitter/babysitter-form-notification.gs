// =========================================
// Google フォーム 回答通知スクリプト
// Home Services / Tutors 応募フォーム用
// =========================================

// 通知先メールアドレス
const NOTIFICATION_EMAIL = "your-email@example.com";

// メール件名
const EMAIL_SUBJECT = "【新規応募】Home Services / Tutors の応募がありました";

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

    // 応募者への自動返信メール
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
    return responses[key] ? responses[key].join(", ") : "Not answered";
  };

  // Availability をまとめる
  const availability = [];
  const morningAvail = getValue("Availability [Morning (9:00 AM - 12:00 PM)]");
  const afternoonAvail = getValue("Availability [Afternoon (1:00 PM - 5:00 PM)]");
  const eveningAvail = getValue("Availability [Evening (6:00 PM - 9:00 PM)]");

  if (morningAvail !== "Not answered" && morningAvail !== "") availability.push("Morning (9:00 AM - 12:00 PM): " + morningAvail);
  if (afternoonAvail !== "Not answered" && afternoonAvail !== "") availability.push("Afternoon (1:00 PM - 5:00 PM): " + afternoonAvail);
  if (eveningAvail !== "Not answered" && eveningAvail !== "") availability.push("Evening (6:00 PM - 9:00 PM): " + eveningAvail);
  const availabilityText = availability.length > 0 ? availability.join("\n") : "Not answered";

  // 希望職種
  const preferredRole = getValue("Preferred Role/Service Applying For");
  const otherRole = getValue("If 'Other' was selected above, please specify the role/service.");
  const roleText = otherRole !== "Not answered" && otherRole !== ""
    ? preferredRole + " (" + otherRole + ")"
    : preferredRole;

  let body = `
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
  New Application Received
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

【Received Date/Time】
${getValue("タイムスタンプ")}

────────────────────────────────
■ Applicant Information
────────────────────────────────
Full Name: ${getValue("Full Name")}
Email Address: ${getValue("Email Address")}
Phone Number: ${getValue("Phone Number")}

────────────────────────────────
■ Application Details
────────────────────────────────
Preferred Role/Service: ${roleText}
Years of Experience: ${getValue("Years of relevant experience in this role/service")}
Expected Hourly Rate (MYR): ${getValue("Expected Hourly Rate (in MYR)")}

────────────────────────────────
■ Availability
────────────────────────────────
${availabilityText}

────────────────────────────────
■ Qualifications & Experience
────────────────────────────────
${getValue("Please describe your qualifications, certifications, and experience in detail.")}

────────────────────────────────
■ Additional Information
────────────────────────────────
Own Transportation: ${getValue("Do you have your own transportation?")}
Communication Skills: ${getValue("How would you rate your communication skills?")}
Resume/CV: ${getValue("Please upload your resume/CV.")}

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
※ This is an automated email
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
    return responses[key] ? responses[key].join(", ") : "Not answered";
  };

  // Availability をまとめる
  const morningAvail = getValue("Availability [Morning (9:00 AM - 12:00 PM)]");
  const afternoonAvail = getValue("Availability [Afternoon (1:00 PM - 5:00 PM)]");
  const eveningAvail = getValue("Availability [Evening (6:00 PM - 9:00 PM)]");

  // 希望職種
  const preferredRole = getValue("Preferred Role/Service Applying For");
  const otherRole = getValue("If 'Other' was selected above, please specify the role/service.");
  const roleText = otherRole !== "Not answered" && otherRole !== ""
    ? preferredRole + " (" + otherRole + ")"
    : preferredRole;

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
    .header { background: linear-gradient(135deg, #FF6B6B 0%, #FFE66D 100%); color: white; padding: 20px; border-radius: 8px 8px 0 0; text-align: center; }
    .header h1 { margin: 0; font-size: 20px; text-shadow: 1px 1px 2px rgba(0,0,0,0.2); }
    .content { background: #f9f9f9; padding: 20px; border: 1px solid #e0e0e0; }
    .section { background: white; margin-bottom: 15px; padding: 15px; border-radius: 6px; box-shadow: 0 1px 3px rgba(0,0,0,0.1); }
    .section-title { color: #FF6B6B; font-weight: bold; font-size: 14px; margin-bottom: 10px; border-bottom: 2px solid #FF6B6B; padding-bottom: 5px; }
    .field { margin-bottom: 8px; }
    .field-label { font-weight: bold; color: #555; font-size: 12px; }
    .field-value { color: #333; margin-top: 2px; }
    .long-text { background: #f5f5f5; padding: 10px; border-radius: 4px; white-space: pre-wrap; }
    .footer { text-align: center; padding: 15px; color: #888; font-size: 11px; }
    .timestamp { background: #FF6B6B; color: white; display: inline-block; padding: 5px 10px; border-radius: 4px; font-size: 12px; }
    .highlight { background: #FFF3CD; padding: 3px 8px; border-radius: 3px; }
    .availability-table { width: 100%; border-collapse: collapse; margin-top: 5px; }
    .availability-table td { padding: 5px 10px; border-bottom: 1px solid #eee; }
    .availability-table td:first-child { font-weight: bold; color: #555; width: 50%; }
  </style>
</head>
<body>
  <div class="container">
    <div class="header">
      <h1>New Application Received</h1>
    </div>
    <div class="content">
      <p style="text-align: center;">
        <span class="timestamp">Received: ${escapeHtml(getValue("タイムスタンプ"))}</span>
      </p>

      <div class="section">
        <div class="section-title">Applicant Information</div>
        <div class="field">
          <div class="field-label">Full Name</div>
          <div class="field-value">${escapeHtml(getValue("Full Name"))}</div>
        </div>
        <div class="field">
          <div class="field-label">Email Address</div>
          <div class="field-value"><a href="mailto:${escapeHtml(getValue("Email Address"))}">${escapeHtml(getValue("Email Address"))}</a></div>
        </div>
        <div class="field">
          <div class="field-label">Phone Number</div>
          <div class="field-value">${escapeHtml(getValue("Phone Number"))}</div>
        </div>
      </div>

      <div class="section">
        <div class="section-title">Application Details</div>
        <div class="field">
          <div class="field-label">Preferred Role/Service</div>
          <div class="field-value"><span class="highlight">${escapeHtml(roleText)}</span></div>
        </div>
        <div class="field">
          <div class="field-label">Years of Experience</div>
          <div class="field-value">${escapeHtml(getValue("Years of relevant experience in this role/service"))}</div>
        </div>
        <div class="field">
          <div class="field-label">Expected Hourly Rate (MYR)</div>
          <div class="field-value">${escapeHtml(getValue("Expected Hourly Rate (in MYR)"))}</div>
        </div>
      </div>

      <div class="section">
        <div class="section-title">Availability</div>
        <table class="availability-table">
          <tr>
            <td>Morning (9:00 AM - 12:00 PM)</td>
            <td>${escapeHtml(morningAvail)}</td>
          </tr>
          <tr>
            <td>Afternoon (1:00 PM - 5:00 PM)</td>
            <td>${escapeHtml(afternoonAvail)}</td>
          </tr>
          <tr>
            <td>Evening (6:00 PM - 9:00 PM)</td>
            <td>${escapeHtml(eveningAvail)}</td>
          </tr>
        </table>
      </div>

      <div class="section">
        <div class="section-title">Qualifications & Experience</div>
        <div class="field-value long-text">${escapeHtml(getValue("Please describe your qualifications, certifications, and experience in detail."))}</div>
      </div>

      <div class="section">
        <div class="section-title">Additional Information</div>
        <div class="field">
          <div class="field-label">Own Transportation</div>
          <div class="field-value">${escapeHtml(getValue("Do you have your own transportation?"))}</div>
        </div>
        <div class="field">
          <div class="field-label">Communication Skills</div>
          <div class="field-value">${escapeHtml(getValue("How would you rate your communication skills?"))}</div>
        </div>
        <div class="field">
          <div class="field-label">Resume/CV</div>
          <div class="field-value">${escapeHtml(getValue("Please upload your resume/CV."))}</div>
        </div>
      </div>
    </div>
    <div class="footer">
      This is an automated email
    </div>
  </div>
</body>
</html>
`;
}

/**
 * 応募者への自動返信メール送信
 * @param {Object} responses - フォーム回答データ
 */
function sendAutoReply(responses) {
  const applicantEmail = responses["Email Address"];
  const applicantName = responses["Full Name"];

  if (!applicantEmail || applicantEmail[0] === "") return;

  const subject = "Thank you for your application - Home Services / Tutors";
  const body = `
Dear ${applicantName ? applicantName[0] : "Applicant"},

Thank you for submitting your application for our Home Services / Tutors position.

We have received your application and our team will review it carefully.
You can expect to hear from us within 3 business days.

Selection Process:
1. Application Review
2. Interview (Online or In-person)
3. Final Decision

If you have any questions, please feel free to contact us.

Best regards,

━━━━━━━━━━━━━━━━━━━━━━━━
Home Services / Tutors Recruitment Team
your-email@example.com
━━━━━━━━━━━━━━━━━━━━━━━━
`;

  MailApp.sendEmail({
    to: applicantEmail[0],
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
      "タイムスタンプ": ["2026/01/19 10:00:00"],
      "メールアドレス": ["form@example.com"],
      "Full Name": ["John Smith"],
      "Email Address": ["john.smith@example.com"],
      "Phone Number": ["+60 12-345 6789"],
      "Preferred Role/Service Applying For": ["Babysitter"],
      "If 'Other' was selected above, please specify the role/service.": [""],
      "Years of relevant experience in this role/service": ["3-5 years"],
      "Availability [Morning (9:00 AM - 12:00 PM)]": ["Available"],
      "Availability [Afternoon (1:00 PM - 5:00 PM)]": ["Available"],
      "Availability [Evening (6:00 PM - 9:00 PM)]": ["Not Available"],
      "Expected Hourly Rate (in MYR)": ["50-80"],
      "Please describe your qualifications, certifications, and experience in detail.": ["I have 5 years of experience working as a babysitter and tutor.\nI hold a diploma in Early Childhood Education.\nI am CPR certified and have experience with children aged 2-12."],
      "Do you have your own transportation?": ["Yes"],
      "How would you rate your communication skills?": ["Excellent"],
      "Please upload your resume/CV.": ["https://drive.google.com/file/d/xxxxx"]
    }
  };

  onFormSubmit(testEvent);
  console.log("Test email sent successfully");
}
