// Gmail送信処理（任意の業務メール送信機能）

// サイドバーUI表示
function showEmailComposer() {
  const html = HtmlService.createHtmlOutput(getEmailComposerHtml())
    .setTitle('業務メール作成')
    .setWidth(400);
  SpreadsheetApp.getUi().showSidebar(html);
}

// メール送信（サイドバーから呼び出し）
function sendBusinessEmail(to, subject, body) {
  try {
    // 入力検証
    if (!to || !subject || !body) {
      throw new Error('宛先、件名、本文はすべて必須です。');
    }
    
    // メールアドレス検証
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    if (!emailRegex.test(to)) {
      throw new Error('有効なメールアドレスを入力してください。');
    }
    
    // デフォルトの件名プレフィックスを追加
    const prefixedSubject = subject.startsWith('[WORK-REQ]') ? subject : `[WORK-REQ] ${subject}`;
    
    // HTML形式のメール本文作成
    const htmlBody = createBusinessEmailHtml(body);
    
    // メール送信
    GmailApp.sendEmail(to, prefixedSubject, body, {
      htmlBody: htmlBody,
      name: 'タスク管理システム'
    });
    
    // ログ記録
    logActivity('OUTBOUND_EMAIL', `Sent to: ${to}, Subject: ${prefixedSubject}`);
    
    return {
      success: true,
      message: 'メールを送信しました。'
    };
    
  } catch (e) {
    logActivity('OUTBOUND_ERROR', e.toString());
    return {
      success: false,
      message: `エラー: ${e.toString()}`
    };
  }
}

// ビジネスメールHTML作成
function createBusinessEmailHtml(body) {
  // 改行をHTMLのbrタグに変換
  const htmlBody = body.replace(/\n/g, '<br>');
  
  return `
    <div style="font-family: 'Noto Sans JP', Arial, sans-serif; max-width: 600px; margin: 0 auto;">
      <div style="background-color: #f8f9fa; padding: 20px; border: 1px solid #dee2e6; border-radius: 8px;">
        <div style="background-color: white; padding: 20px; border-radius: 4px;">
          <p style="font-size: 16px; line-height: 1.6; color: #333; margin: 0;">
            ${htmlBody}
          </p>
        </div>
        
        <div style="margin-top: 20px; padding-top: 20px; border-top: 1px solid #dee2e6;">
          <p style="color: #6c757d; font-size: 14px; margin: 0;">
            このメールはタスク管理システムから送信されています。<br>
            業務内容に基づいて自動的に業務記述書とタスクフローが生成されます。
          </p>
        </div>
      </div>
    </div>
  `;
}

// サイドバーHTML取得
function getEmailComposerHtml() {
  const defaultTo = getConfig('DEFAULT_TO_EMAIL') || '';
  
  return `
    <!DOCTYPE html>
    <html>
      <head>
        <base target="_top">
        <style>
          body {
            font-family: 'Noto Sans JP', Arial, sans-serif;
            padding: 15px;
            margin: 0;
          }
          .form-group {
            margin-bottom: 15px;
          }
          label {
            display: block;
            margin-bottom: 5px;
            font-weight: bold;
            color: #333;
          }
          input, textarea {
            width: 100%;
            padding: 8px;
            border: 1px solid #ddd;
            border-radius: 4px;
            font-size: 14px;
            box-sizing: border-box;
          }
          textarea {
            resize: vertical;
            min-height: 150px;
          }
          button {
            width: 100%;
            padding: 10px;
            border: none;
            border-radius: 4px;
            font-size: 16px;
            font-weight: bold;
            cursor: pointer;
            transition: background-color 0.3s;
          }
          .btn-primary {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            margin-bottom: 10px;
          }
          .btn-primary:hover {
            opacity: 0.9;
          }
          .btn-secondary {
            background-color: #6c757d;
            color: white;
          }
          .btn-secondary:hover {
            background-color: #5a6268;
          }
          .loading {
            display: none;
            text-align: center;
            padding: 20px;
          }
          .spinner {
            border: 3px solid #f3f3f3;
            border-top: 3px solid #667eea;
            border-radius: 50%;
            width: 40px;
            height: 40px;
            animation: spin 1s linear infinite;
            margin: 0 auto;
          }
          @keyframes spin {
            0% { transform: rotate(0deg); }
            100% { transform: rotate(360deg); }
          }
          .message {
            padding: 10px;
            border-radius: 4px;
            margin-bottom: 15px;
            display: none;
          }
          .message.success {
            background-color: #d4edda;
            color: #155724;
            border: 1px solid #c3e6cb;
          }
          .message.error {
            background-color: #f8d7da;
            color: #721c24;
            border: 1px solid #f5c6cb;
          }
          .info {
            background-color: #e3f2fd;
            border: 1px solid #90caf9;
            border-radius: 4px;
            padding: 10px;
            margin-bottom: 15px;
            color: #1565c0;
            font-size: 13px;
          }
        </style>
      </head>
      <body>
        <h2 style="color: #333; margin-top: 0;">業務メール作成</h2>
        
        <div class="info">
          ℹ️ このフォームから送信されたメールは自動的に処理され、業務記述書とタスクフローが生成されます。
        </div>
        
        <div id="message" class="message"></div>
        
        <form id="emailForm">
          <div class="form-group">
            <label for="to">宛先メールアドレス *</label>
            <input type="email" id="to" name="to" value="${defaultTo}" required placeholder="example@example.com">
          </div>
          
          <div class="form-group">
            <label for="subject">件名 *</label>
            <input type="text" id="subject" name="subject" required placeholder="業務依頼のタイトル">
            <small style="color: #666; font-size: 12px;">※ [WORK-REQ] プレフィックスは自動付与されます</small>
          </div>
          
          <div class="form-group">
            <label for="body">業務内容 *</label>
            <textarea id="body" name="body" required placeholder="実施したい業務の詳細を記入してください。&#10;&#10;例：&#10;- 業務の目的&#10;- 必要な成果物&#10;- 期限&#10;- 関係者&#10;- その他要件"></textarea>
          </div>
          
          <button type="submit" class="btn-primary">送信</button>
          <button type="button" class="btn-secondary" onclick="clearForm()">クリア</button>
        </form>
        
        <div id="loading" class="loading">
          <div class="spinner"></div>
          <p>送信中...</p>
        </div>
        
        <script>
          document.getElementById('emailForm').addEventListener('submit', function(e) {
            e.preventDefault();
            
            const to = document.getElementById('to').value;
            const subject = document.getElementById('subject').value;
            const body = document.getElementById('body').value;
            
            // フォームを非表示、ローディング表示
            document.getElementById('emailForm').style.display = 'none';
            document.getElementById('loading').style.display = 'block';
            document.getElementById('message').style.display = 'none';
            
            // GASの関数を呼び出し
            google.script.run
              .withSuccessHandler(function(result) {
                document.getElementById('loading').style.display = 'none';
                document.getElementById('emailForm').style.display = 'block';
                
                const messageDiv = document.getElementById('message');
                messageDiv.className = result.success ? 'message success' : 'message error';
                messageDiv.textContent = result.message;
                messageDiv.style.display = 'block';
                
                if (result.success) {
                  // 成功時はフォームをクリア
                  clearForm();
                  
                  // 3秒後にメッセージを非表示
                  setTimeout(function() {
                    messageDiv.style.display = 'none';
                  }, 3000);
                }
              })
              .withFailureHandler(function(error) {
                document.getElementById('loading').style.display = 'none';
                document.getElementById('emailForm').style.display = 'block';
                
                const messageDiv = document.getElementById('message');
                messageDiv.className = 'message error';
                messageDiv.textContent = 'エラーが発生しました: ' + error.toString();
                messageDiv.style.display = 'block';
              })
              .sendBusinessEmail(to, subject, body);
          });
          
          function clearForm() {
            document.getElementById('subject').value = '';
            document.getElementById('body').value = '';
            // 宛先はデフォルト値があれば保持
          }
        </script>
      </body>
    </html>
  `;
}