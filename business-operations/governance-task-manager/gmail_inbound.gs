// Gmail受信処理

// 新着メール処理（メイン関数）
function processNewEmails() {
  const query = getConfig('PROCESSING_QUERY') || 'label:inbox is:unread';
  logActivity('PROCESS_START', `Processing emails with query: ${query}`);
  
  try {
    const threads = GmailApp.search(query);
    
    if (threads.length === 0) {
      logActivity('PROCESS_INFO', 'No new emails found');
      return;
    }
    
    threads.forEach(thread => {
      processThread(thread);
    });
    
    logActivity('PROCESS_END', `Processed ${threads.length} threads`);
  } catch (e) {
    logActivity('PROCESS_ERROR', e.toString());
    throw e;
  }
}

// スレッド処理
function processThread(thread) {
  const messages = thread.getMessages();
  
  messages.forEach(msg => {
    try {
      processMessage(msg, thread);
    } catch (e) {
      logActivity('MESSAGE_ERROR', `Failed to process message: ${e.toString()}`);
    }
  });
}

// メッセージ処理
function processMessage(msg, thread) {
  const messageId = msg.getId();
  
  // 処理済みチェック
  if (isProcessed(messageId)) {
    logActivity('SKIP', `Message ${messageId} already processed`);
    return;
  }
  
  // メール情報抽出
  const from = extractEmail(msg.getFrom());
  const subject = msg.getSubject();
  const htmlBody = msg.getBody();
  const plainBody = msg.getPlainBody() || htmlToText(htmlBody);
  const receivedDate = msg.getDate();
  
  // Inboxにログ記録
  logInbox(messageId, thread.getId(), from, subject, plainBody.substring(0, 200), 'NEW');
  
  try {
    // OpenAI呼び出し
    const orgProfile = getConfig('ORG_PROFILE_JSON') || '{}';
    const result = callOpenAI(plainBody, orgProfile);
    
    // 検証
    validateOpenAIResponse(result);
    
    // データ書き込み（改善されたエンジンを使用）
    writeWorkSpec(result.work_spec);
    
    // 新しいデータ処理エンジンを強制使用
    writeFlowRowsImproved(result.flow_rows);
    
    // ビジュアルフロー生成
    if (typeof generateVisualFlow === 'function') {
      generateVisualFlow();
    }
    
    // 共有設定
    const shareSuccess = handleSharing(from);
    
    // 返信メール送信
    sendNotificationEmail(from, result.work_spec, ss().getUrl());
    
    // 処理済みマーク
    markProcessed(messageId);
    labelThreadProcessed(thread);
    
    logActivity('PROCESS_SUCCESS', `Successfully processed message ${messageId}`);
  } catch (e) {
    logError(messageId, e);
    
    // エラー通知メール送信
    sendErrorNotificationEmail(from, subject, e.toString());
    
    throw e;
  }
}

// 共有設定処理
function handleSharing(senderEmail) {
  let shareSuccess = false;
  
  // ANYONE_WITH_LINKの設定を試行
  if (String(getConfig('SHARE_ANYONE_WITH_LINK')).toUpperCase() === 'TRUE') {
    shareSuccess = shareSheetAnyWithLink();
  }
  
  // 送信者を編集者として追加
  const editorSuccess = addEditor(senderEmail);
  
  return shareSuccess && editorSuccess;
}

// 成功通知メール送信
function sendNotificationEmail(to, workSpec, sheetUrl) {
  const subject = `[WORK-SPEC READY] ${workSpec.title}`;
  const plainBody = buildPlainTextNotification(workSpec, sheetUrl);
  const htmlBody = buildHtmlNotification(workSpec, sheetUrl);
  
  GmailApp.sendEmail(to, subject, plainBody, {
    htmlBody: htmlBody,
    name: 'タスク管理システム'
  });
  
  logActivity('EMAIL_SENT', `Notification sent to ${to}`);
}

// プレーンテキスト通知作成
function buildPlainTextNotification(workSpec, sheetUrl) {
  return `業務記述書の作成が完了しました。

タイトル: ${workSpec.title}
概要: ${workSpec.summary}

スプレッドシートURL: ${sheetUrl}

このスプレッドシートでは以下の内容を確認・編集できます：
- 業務記述書（詳細仕様）
- タスクフロー表
- ビジュアルフロー図

【重要な注意事項】
- 本書面は自動生成されたものです。最終的な判断は専門家にご確認ください。
- 法令・規制に関する記載は参考情報であり、法的助言ではありません。
- スプレッドシートは編集可能です。必要に応じて内容を更新してください。

---
タスク管理システム by Google Apps Script`;
}

// HTML通知作成
function buildHtmlNotification(workSpec, sheetUrl) {
  return `
    <div style="font-family: 'Noto Sans JP', Arial, sans-serif; max-width: 600px; margin: 0 auto;">
      <div style="background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; padding: 20px; text-align: center; border-radius: 10px 10px 0 0;">
        <h1 style="margin: 0; font-size: 24px;">📋 業務記述書が完成しました</h1>
      </div>
      
      <div style="padding: 20px; background-color: #f8f9fa; border: 1px solid #e9ecef; border-top: none;">
        <h2 style="color: #495057; margin-top: 0;">${workSpec.title}</h2>
        <p style="font-size: 16px; line-height: 1.5; color: #6c757d;">${workSpec.summary}</p>
        
        <div style="background-color: white; padding: 15px; border-radius: 8px; margin: 20px 0; border-left: 4px solid #28a745;">
          <h3 style="margin-top: 0; color: #28a745;">📊 スプレッドシート（編集可能）</h3>
          <p style="margin-bottom: 10px;">以下のリンクから業務記述書とタスク管理シートにアクセス・編集できます：</p>
          <a href="${sheetUrl}" style="display: inline-block; background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; padding: 12px 24px; text-decoration: none; border-radius: 4px; font-weight: bold;">📝 スプレッドシートを開く</a>
        </div>
        
        ${workSpec.timeline && workSpec.timeline.length > 0 ? `
        <div style="background-color: white; padding: 15px; border-radius: 8px; margin: 20px 0; border-left: 4px solid #ffc107;">
          <h3 style="margin-top: 0; color: #ff9800;">⏰ 主要マイルストーン</h3>
          <ul style="margin: 0; padding-left: 20px;">
            ${workSpec.timeline.map(phase => `
              <li style="margin-bottom: 8px;">
                <strong>${phase.phase}</strong> (${phase.duration_hint})
                ${phase.milestones && phase.milestones.length > 0 ? 
                  `<ul style="margin-top: 5px;">${phase.milestones.map(milestone => 
                    `<li style="color: #6c757d;">${milestone}</li>`
                  ).join('')}</ul>` 
                  : ''}
              </li>
            `).join('')}
          </ul>
        </div>
        ` : ''}
        
        <div style="background-color: white; padding: 15px; border-radius: 8px; margin: 20px 0; border-left: 4px solid #dc3545;">
          <h3 style="margin-top: 0; color: #dc3545;">⚠️ 重要な注意事項</h3>
          <ul style="margin: 0; padding-left: 20px; color: #6c757d;">
            <li>本書面は自動生成されたものです。最終的な判断は専門家にご確認ください。</li>
            <li>法令・規制に関する記載は参考情報であり、法的助言ではありません。</li>
            <li>スプレッドシートは編集可能です。必要に応じて内容を更新してください。</li>
          </ul>
        </div>
        
        <div style="text-align: center; margin-top: 30px; padding-top: 20px; border-top: 1px solid #dee2e6;">
          <p style="color: #6c757d; font-size: 14px; margin: 0;">
            このメールは自動送信されています。<br>
            タスク管理システム by Google Apps Script
          </p>
        </div>
      </div>
    </div>
  `;
}

// エラー通知メール送信
function sendErrorNotificationEmail(to, originalSubject, errorMessage) {
  const subject = `[WORK-SPEC ERROR] 処理エラー: ${originalSubject}`;
  const body = `業務記述書の作成中にエラーが発生しました。

元の件名: ${originalSubject}
エラー内容: ${errorMessage}

お手数ですが、システム管理者にお問い合わせください。

---
タスク管理システム by Google Apps Script`;
  
  try {
    GmailApp.sendEmail(to, subject, body);
    logActivity('ERROR_EMAIL_SENT', `Error notification sent to ${to}`);
  } catch (e) {
    logActivity('ERROR_EMAIL_FAILED', `Failed to send error notification: ${e.toString()}`);
  }
}

// スレッドに処理済みラベルを付与
function labelThreadProcessed(thread) {
  try {
    // 既存のラベルを取得または作成
    let label = GmailApp.getUserLabelByName('PROCESSED');
    if (!label) {
      label = GmailApp.createLabel('PROCESSED');
    }
    
    thread.addLabel(label);
    thread.markRead();
    
    logActivity('LABEL', `Added PROCESSED label to thread ${thread.getId()}`);
  } catch (e) {
    logActivity('LABEL_ERROR', `Failed to label thread: ${e.toString()}`);
  }
}