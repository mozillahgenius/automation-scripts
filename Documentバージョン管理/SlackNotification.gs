/**
 * Slack通知機能
 * 文書管理システムからSlackへ通知を送信
 */

// Slack設定（スクリプトプロパティから取得）
function getSlackConfig() {
  const scriptProperties = PropertiesService.getScriptProperties();
  return {
    webhookUrl: scriptProperties.getProperty('SLACK_WEBHOOK_URL') || '',
    channel: scriptProperties.getProperty('SLACK_CHANNEL') || '#general',
    username: scriptProperties.getProperty('SLACK_USERNAME') || '文書管理システム',
    iconEmoji: scriptProperties.getProperty('SLACK_ICON_EMOJI') || ':page_facing_up:'
  };
}

/**
 * Slack設定を保存
 */
function setSlackConfig(webhookUrl, channel, username, iconEmoji) {
  const scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperties({
    'SLACK_WEBHOOK_URL': webhookUrl,
    'SLACK_CHANNEL': channel || '#general',
    'SLACK_USERNAME': username || '文書管理システム',
    'SLACK_ICON_EMOJI': iconEmoji || ':page_facing_up:'
  });
}

/**
 * Slackにメッセージを送信
 */
function sendToSlack(message, attachments = []) {
  const config = getSlackConfig();
  
  if (!config.webhookUrl) {
    console.error('Slack Webhook URLが設定されていません');
    return false;
  }
  
  const payload = {
    channel: config.channel,
    username: config.username,
    icon_emoji: config.iconEmoji,
    text: message,
    attachments: attachments
  };
  
  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  
  try {
    const response = UrlFetchApp.fetch(config.webhookUrl, options);
    return response.getResponseCode() === 200;
  } catch (e) {
    console.error('Slack送信エラー:', e);
    return false;
  }
}

/**
 * 新規文書追加を通知
 */
function notifyNewDocument(docData) {
  const message = '新しい文書が追加されました';
  
  const attachment = {
    fallback: message,
    color: 'good',
    title: '新規文書',
    fields: [
      {
        title: 'DocKey',
        value: docData.docKey,
        short: true
      },
      {
        title: 'タイトル',
        value: docData.title,
        short: true
      },
      {
        title: 'ステージ',
        value: docData.stage,
        short: true
      },
      {
        title: '期限',
        value: docData.dueDate || '未設定',
        short: true
      },
      {
        title: '作成者',
        value: Session.getActiveUser().getEmail(),
        short: true
      }
    ],
    footer: '文書管理システム',
    footer_icon: 'https://cdn-icons-png.flaticon.com/512/2991/2991110.png',
    ts: Math.floor(Date.now() / 1000)
  };
  
  return sendToSlack(message, [attachment]);
}

/**
 * 文書ステータス変更を通知
 */
function notifyStatusChange(docKey, oldStatus, newStatus, docTitle) {
  const message = `文書のステータスが変更されました`;
  
  const statusColors = {
    'DRAFT': '#808080',
    'FOR-REVIEW': '#FFA500',
    'APPROVED': '#008000',
    'ARCHIVED': '#4B0082'
  };
  
  const attachment = {
    fallback: message,
    color: statusColors[newStatus] || 'warning',
    title: 'ステータス変更',
    fields: [
      {
        title: 'DocKey',
        value: docKey,
        short: true
      },
      {
        title: 'タイトル',
        value: docTitle,
        short: false
      },
      {
        title: '変更前',
        value: oldStatus,
        short: true
      },
      {
        title: '変更後',
        value: newStatus,
        short: true
      },
      {
        title: '変更者',
        value: Session.getActiveUser().getEmail(),
        short: true
      }
    ],
    footer: '文書管理システム',
    footer_icon: 'https://cdn-icons-png.flaticon.com/512/2991/2991110.png',
    ts: Math.floor(Date.now() / 1000)
  };
  
  return sendToSlack(message, [attachment]);
}

/**
 * 期限切れ文書を通知
 */
function notifyOverdueDocuments() {
  const overdueDocuments = getOverdueDocuments();
  
  if (overdueDocuments.length === 0) {
    return;
  }
  
  const message = `⚠️ ${overdueDocuments.length}件の文書が期限切れです`;
  
  const attachments = overdueDocuments.map(doc => {
    return {
      fallback: `${doc.docKey}: ${doc.title} - ${doc.daysPastDue}日超過`,
      color: doc.daysPastDue > 7 ? 'danger' : 'warning',
      title: doc.title,
      fields: [
        {
          title: 'DocKey',
          value: doc.docKey,
          short: true
        },
        {
          title: '期限',
          value: doc.dueDate,
          short: true
        },
        {
          title: '超過日数',
          value: `${doc.daysPastDue}日`,
          short: true
        },
        {
          title: 'ステータス',
          value: doc.projectStatus,
          short: true
        }
      ]
    };
  });
  
  // 最大5件まで表示
  const limitedAttachments = attachments.slice(0, 5);
  if (attachments.length > 5) {
    limitedAttachments.push({
      fallback: `他${attachments.length - 5}件`,
      color: '#808080',
      text: `他${attachments.length - 5}件の期限切れ文書があります`
    });
  }
  
  return sendToSlack(message, limitedAttachments);
}

/**
 * メール送受信を通知
 */
function notifyEmailActivity(docKey, docTitle, emailInfo) {
  const senderType = emailInfo.lastSentBy === SENDER_TYPE.SELF ? '自社' : '相手先';
  const message = `📧 文書に関連するメールが${senderType}から送信されました`;
  
  const attachment = {
    fallback: message,
    color: emailInfo.lastSentBy === SENDER_TYPE.SELF ? '#0084FF' : '#00C853',
    title: 'メール送受信',
    fields: [
      {
        title: 'DocKey',
        value: docKey,
        short: true
      },
      {
        title: 'タイトル',
        value: docTitle,
        short: false
      },
      {
        title: '送信者',
        value: senderType,
        short: true
      },
      {
        title: '件名',
        value: emailInfo.subject,
        short: false
      },
      {
        title: 'From',
        value: emailInfo.from,
        short: true
      },
      {
        title: 'To',
        value: emailInfo.to,
        short: true
      }
    ],
    actions: [
      {
        type: 'button',
        text: 'メールを開く',
        url: emailInfo.lastMailUrl
      }
    ],
    footer: '文書管理システム',
    footer_icon: 'https://cdn-icons-png.flaticon.com/512/2991/2991110.png',
    ts: Math.floor(Date.now() / 1000)
  };
  
  return sendToSlack(message, [attachment]);
}

/**
 * プロジェクト完了を通知
 */
function notifyProjectCompletion(docKey, docTitle) {
  const message = '✅ プロジェクトが完了しました';
  
  const attachment = {
    fallback: message,
    color: 'good',
    title: 'プロジェクト完了',
    fields: [
      {
        title: 'DocKey',
        value: docKey,
        short: true
      },
      {
        title: 'タイトル',
        value: docTitle,
        short: false
      },
      {
        title: '完了日',
        value: formatDate(new Date()),
        short: true
      },
      {
        title: '完了者',
        value: Session.getActiveUser().getEmail(),
        short: true
      }
    ],
    footer: '文書管理システム',
    footer_icon: 'https://cdn-icons-png.flaticon.com/512/2991/2991110.png',
    ts: Math.floor(Date.now() / 1000)
  };
  
  return sendToSlack(message, [attachment]);
}

/**
 * 週次サマリーを送信
 */
function sendWeeklySummary() {
  const sheet = getOrCreateSheet();
  const data = sheet.getDataRange().getValues();
  
  // 統計情報を集計
  const stats = {
    total: data.length - 1,
    draft: 0,
    forReview: 0,
    approved: 0,
    archived: 0,
    open: 0,
    inProgress: 0,
    closed: 0,
    delayed: 0,
    overdue: 0
  };
  
  const today = new Date();
  
  for (let i = 1; i < data.length; i++) {
    // ステージ別
    switch (data[i][COLUMNS.STAGE]) {
      case STAGES.DRAFT: stats.draft++; break;
      case STAGES.FOR_REVIEW: stats.forReview++; break;
      case STAGES.APPROVED: stats.approved++; break;
      case STAGES.ARCHIVED: stats.archived++; break;
    }
    
    // プロジェクトステータス別
    switch (data[i][COLUMNS.PROJECT_STATUS]) {
      case PROJECT_STATUS.OPEN: stats.open++; break;
      case PROJECT_STATUS.IN_PROGRESS: stats.inProgress++; break;
      case PROJECT_STATUS.CLOSED: stats.closed++; break;
      case PROJECT_STATUS.DELAYED: stats.delayed++; break;
    }
    
    // 期限切れチェック
    const dueDate = data[i][COLUMNS.DUE_DATE];
    if (dueDate && new Date(dueDate) < today) {
      stats.overdue++;
    }
  }
  
  const message = '📊 週次文書管理レポート';
  
  const attachment = {
    fallback: message,
    color: '#36a64f',
    title: '週次サマリー',
    pretext: `${formatDate(today)} 時点の文書管理状況`,
    fields: [
      {
        title: '総文書数',
        value: stats.total.toString(),
        short: true
      },
      {
        title: '期限切れ',
        value: stats.overdue > 0 ? `⚠️ ${stats.overdue}件` : '0件',
        short: true
      },
      {
        title: '文書ステージ',
        value: `下書き: ${stats.draft}\nレビュー中: ${stats.forReview}\n承認済み: ${stats.approved}\nアーカイブ: ${stats.archived}`,
        short: true
      },
      {
        title: 'プロジェクト状況',
        value: `オープン: ${stats.open}\n進行中: ${stats.inProgress}\n完了: ${stats.closed}\n遅延: ${stats.delayed}`,
        short: true
      }
    ],
    footer: '文書管理システム',
    footer_icon: 'https://cdn-icons-png.flaticon.com/512/2991/2991110.png',
    ts: Math.floor(Date.now() / 1000)
  };
  
  return sendToSlack(message, [attachment]);
}

/**
 * テスト通知を送信
 */
function testSlackNotification() {
  const message = '🔔 Slack連携テスト';
  
  const attachment = {
    fallback: message,
    color: 'good',
    title: 'テスト通知',
    text: 'Slack通知機能が正常に動作しています',
    fields: [
      {
        title: 'テスト実行者',
        value: Session.getActiveUser().getEmail(),
        short: true
      },
      {
        title: '実行時刻',
        value: formatDateTime(new Date()),
        short: true
      }
    ],
    footer: '文書管理システム',
    footer_icon: 'https://cdn-icons-png.flaticon.com/512/2991/2991110.png',
    ts: Math.floor(Date.now() / 1000)
  };
  
  const result = sendToSlack(message, [attachment]);
  
  if (result) {
    SpreadsheetApp.getUi().alert('Slackテスト通知を送信しました');
  } else {
    SpreadsheetApp.getUi().alert('Slack通知の送信に失敗しました。Webhook URLを確認してください。');
  }
}