// Slack Bot作成
// https://api.slack.com/apps

// Botの作成手順
// https://blog.da-vinci-studio.com/entry/2022/09/13/101530

// Bot更新時に必要な処理（デプロイ更新）
// https://ryjkmr.com/gas-web-app-deploy-new-same-url/

// 利用開始までの流れ
// 1. SlackにBotを設定したいチャンネルを作成しchannel_idをコピー
// 2. channel_idをコードに貼り付け
// 3. OpenAI API Keyを発行しプロジェクト設定のスクリプトプロパティに入力
// 4. GASのデプロイでウェブアプリを作成し、自分で実行、全員が利用可能で共有し発行されたウェブアプリURLをコピー
// 5. SlackAppを以下のManifestにウェブアプリURLを貼り付けしSlackBotを作成
// 6. SlackAppのOAuthのBot Tokenをコピー
// 7. Bot Tokenをプロジェクト設定のスクリプトプロパティに入力
// 8. 作成したBotをSlackにインストール
// 9. チャンネルにBotを追加
// 10. Botの動作確認を行う

const slackManifest = `
display_information:
  name: ChatGPT
features:
  bot_user:
    display_name: ChatGPT
    always_online: false
oauth_config:
  scopes:
    bot:
      - app_mentions:read
      - channels:history
      - chat:write
      - chat:write.public
      - groups:read
      - reactions:read
      - users:read
      - channels:read
      - files:read
      - files:write
      - remote_files:read
      - remote_files:share
settings:
  event_subscriptions:
    request_url: {GAS ウェブアプリURL}
    bot_events:
      - message.channels
      - file_shared
      - file_change
      - file_deleted
  org_deploy_enabled: false
  socket_mode_enabled: false
  token_rotation_enabled: false
`