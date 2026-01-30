/**
 * Gmail連携機能
 * メールの送受信履歴を文書管理システムと連携
 */

/**
 * 指定したDocKeyに関連するメールを検索
 */
function searchRelatedEmails(docKey, title) {
  const queries = [];
  
  // DocKeyまたはタイトルで検索
  if (docKey) {
    queries.push(`"${docKey}"`);
  }
  if (title) {
    queries.push(`"${title}"`);
  }
  
  if (queries.length === 0) {
    return [];
  }
  
  const searchQuery = queries.join(' OR ');
  const threads = GmailApp.search(searchQuery, 0, 50);
  const emails = [];
  
  threads.forEach(thread => {
    const messages = thread.getMessages();
    messages.forEach(message => {
      emails.push({
        id: message.getId(),
        threadId: thread.getId(),
        subject: message.getSubject(),
        from: message.getFrom(),
        to: message.getTo(),
        date: message.getDate(),
        body: message.getPlainBody().substring(0, 200),
        hasAttachments: message.getAttachments().length > 0,
        messageUrl: `https://mail.google.com/mail/u/0/#inbox/${message.getId()}`
      });
    });
  });
  
  return emails;
}

/**
 * 最新のメール送受信情報を取得
 */
function getLatestEmailInfo(docKey, title) {
  const emails = searchRelatedEmails(docKey, title);
  
  if (emails.length === 0) {
    return null;
  }
  
  // 最新のメールを取得
  emails.sort((a, b) => b.date - a.date);
  const latestEmail = emails[0];
  
  // 送信者を判定
  const myEmail = Session.getActiveUser().getEmail();
  const lastSentBy = latestEmail.from.includes(myEmail) ? SENDER_TYPE.SELF : SENDER_TYPE.PARTNER;
  
  return {
    lastMailUrl: latestEmail.messageUrl,
    lastSentBy: lastSentBy,
    lastMailDate: latestEmail.date,
    subject: latestEmail.subject,
    from: latestEmail.from,
    to: latestEmail.to
  };
}

/**
 * メールツリーを構築
 */
function buildMailTree(docKey, title) {
  const emails = searchRelatedEmails(docKey, title);
  
  if (emails.length === 0) {
    return '';
  }
  
  // スレッドごとにグループ化
  const threads = {};
  emails.forEach(email => {
    if (!threads[email.threadId]) {
      threads[email.threadId] = [];
    }
    threads[email.threadId].push(email);
  });
  
  // ツリー形式で文字列化
  let tree = [];
  Object.keys(threads).forEach(threadId => {
    const threadEmails = threads[threadId];
    threadEmails.sort((a, b) => a.date - b.date);
    
    threadEmails.forEach((email, index) => {
      const indent = '  '.repeat(index);
      const dateStr = formatDateTime(email.date);
      tree.push(`${indent}[${dateStr}] ${email.from} → ${email.subject}`);
    });
  });
  
  return tree.join('\n');
}

/**
 * 文書に関連するメールを自動更新
 */
function updateEmailInfoForDocument(docKey) {
  const documents = searchDocuments({ docKey: docKey });
  
  if (documents.length === 0) {
    return { success: false, error: 'Document not found' };
  }
  
  const doc = documents[0];
  const emailInfo = getLatestEmailInfo(doc.docKey, doc.title);
  
  if (emailInfo) {
    const mailTree = buildMailTree(doc.docKey, doc.title);
    
    return updateDocument(docKey, {
      lastSentBy: emailInfo.lastSentBy,
      lastMailUrl: emailInfo.lastMailUrl,
      mailTree: mailTree
    });
  }
  
  return { success: false, error: 'No related emails found' };
}

/**
 * すべての文書のメール情報を更新
 */
function updateAllEmailInfo() {
  const sheet = getOrCreateSheet();
  const data = sheet.getDataRange().getValues();
  let updatedCount = 0;
  
  for (let i = 1; i < data.length; i++) {
    const docKey = data[i][COLUMNS.DOC_KEY];
    const title = data[i][COLUMNS.TITLE];
    
    if (docKey) {
      const emailInfo = getLatestEmailInfo(docKey, title);
      
      if (emailInfo) {
        const mailTree = buildMailTree(docKey, title);
        
        data[i][COLUMNS.LAST_SENT_BY] = emailInfo.lastSentBy;
        data[i][COLUMNS.LAST_MAIL_URL] = emailInfo.lastMailUrl;
        data[i][COLUMNS.MAIL_TREE] = mailTree;
        
        updatedCount++;
      }
    }
  }
  
  // データを書き戻し
  if (updatedCount > 0) {
    sheet.getDataRange().setValues(data);
  }
  
  return {
    success: true,
    updatedCount: updatedCount
  };
}

/**
 * メール送信時に自動でLastSentByを更新するトリガー
 */
function onEmailSent(e) {
  // GmailApp.sendEmailをラップして使用
  const originalSendEmail = GmailApp.sendEmail;
  
  GmailApp.sendEmail = function(recipient, subject, body, options) {
    // オリジナルのメール送信
    originalSendEmail.apply(this, arguments);
    
    // 文書管理システムの更新
    setTimeout(() => {
      updateEmailInfoBySubject(subject);
    }, 1000);
  };
}

/**
 * 件名から文書を特定して更新
 */
function updateEmailInfoBySubject(subject) {
  const sheet = getOrCreateSheet();
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    const docKey = data[i][COLUMNS.DOC_KEY];
    const title = data[i][COLUMNS.TITLE];
    
    if ((docKey && subject.includes(docKey)) || (title && subject.includes(title))) {
      updateDocument(docKey, {
        lastSentBy: SENDER_TYPE.SELF,
        lastMailUrl: getLatestEmailUrl(subject)
      });
      break;
    }
  }
}

/**
 * 最新のメールURLを取得
 */
function getLatestEmailUrl(subject) {
  const threads = GmailApp.search(`subject:"${subject}"`, 0, 1);
  
  if (threads.length > 0) {
    const messages = threads[0].getMessages();
    if (messages.length > 0) {
      const latestMessage = messages[messages.length - 1];
      return `https://mail.google.com/mail/u/0/#inbox/${latestMessage.getId()}`;
    }
  }
  
  return '';
}

/**
 * 日時をフォーマット
 */
function formatDateTime(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  const hours = String(date.getHours()).padStart(2, '0');
  const minutes = String(date.getMinutes()).padStart(2, '0');
  return `${year}-${month}-${day} ${hours}:${minutes}`;
}

/**
 * 定期的にメール情報を更新（トリガーで実行）
 */
function scheduledEmailUpdate() {
  const result = updateAllEmailInfo();
  
  if (result.success) {
    console.log(`Updated email info for ${result.updatedCount} documents`);
  }
}