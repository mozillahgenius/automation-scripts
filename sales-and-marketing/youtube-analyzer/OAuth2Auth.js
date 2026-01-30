// OAuth2設定
var CLIENT_ID = 'YOUR_GOOGLE_CLIENT_ID';
var CLIENT_SECRET = 'YOUR_GOOGLE_CLIENT_SECRET';

function getOAuth2Service() {
  return OAuth2.createService('youtube')
      .setAuthorizationBaseUrl('https://accounts.google.com/o/oauth2/auth')
          .setTokenUrl('https://oauth2.googleapis.com/token')
              .setClientId(CLIENT_ID)
                  .setClientSecret(CLIENT_SECRET)
                      .setCallbackFunction('authCallback')
                          .setPropertyStore(PropertiesService.getUserProperties())
                              .setScope([
                                    'https://www.googleapis.com/auth/youtube.readonly',
                                          'https://www.googleapis.com/auth/yt-analytics.readonly'
                                              ].join(' '))
                                                  .setParam('access_type', 'offline')
                                                      .setParam('prompt', 'consent');
                                                      }

                                                      function authCallback(request) {
                                                        var service = getOAuth2Service();
                                                          var authorized = service.handleCallback(request);
                                                            if (authorized) {
                                                                return HtmlService.createHtmlOutput('認証成功！このタブを閉じてください。');
                                                                  } else {
                                                                      return HtmlService.createHtmlOutput('認証失敗。');
                                                                        }
                                                                        }

                                                                        function startAuth() {
                                                                          var service = getOAuth2Service();
                                                                            if (!service.hasAccess()) {
                                                                                var authorizationUrl = service.getAuthorizationUrl();
                                                                                    Logger.log('以下のURLを開いて認証してください:');
                                                                                        Logger.log(authorizationUrl);
                                                                                            
                                                                                                var html = HtmlService.createHtmlOutput(
                                                                                                      '<p>以下のリンクをクリックして認証してください:</p>' +
                                                                                                            '<p><a href="' + authorizationUrl + '" target="_blank">認証ページを開く</a></p>' +
                                                                                                                  '<p>認証後、ブランドアカウントを選択してください。</p>'
                                                                                                                      ).setWidth(400).setHeight(200);
                                                                                                                          
                                                                                                                              SpreadsheetApp.getUi().showModalDialog(html, 'YouTube認証');
                                                                                                                                } else {
                                                                                                                                    Logger.log('既に認証済みです');
                                                                                                                                      }
                                                                                                                                      }

                                                                                                                                      function resetOAuth() {
                                                                                                                                        var service = getOAuth2Service();
                                                                                                                                          service.reset();
                                                                                                                                            Logger.log('OAuth認証をリセットしました。startAuth()を実行して再認証してください。');
                                                                                                                                            }

                                                                                                                                            function testOAuth2Analytics() {
                                                                                                                                              var service = getOAuth2Service();
                                                                                                                                                
                                                                                                                                                  if (!service.hasAccess()) {
                                                                                                                                                      Logger.log('認証が必要です。startAuth()を実行してください。');
                                                                                                                                                          return;
                                                                                                                                                            }
                                                                                                                                                              
                                                                                                                                                                var accessToken = service.getAccessToken();
                                                                                                                                                                  
                                                                                                                                                                    var endDate = new Date();
                                                                                                                                                                      var startDate = new Date();
                                                                                                                                                                        startDate.setDate(startDate.getDate() - 28);
                                                                                                                                                                          
                                                                                                                                                                            var startDateStr = Utilities.formatDate(startDate, 'Asia/Tokyo', 'yyyy-MM-dd');
                                                                                                                                                                              var endDateStr = Utilities.formatDate(endDate, 'Asia/Tokyo', 'yyyy-MM-dd');
                                                                                                                                                                                
                                                                                                                                                                                  // ブランドアカウントのチャンネルID
                                                                                                                                                                                    var brandChannelId = 'YOUTUBE_CHANNEL_ID_PLACEHOLDER';
                                                                                                                                                                                      
                                                                                                                                                                                        var url = 'https://youtubeanalytics.googleapis.com/v2/reports?' +
                                                                                                                                                                                            'ids=channel%3D%3D' + brandChannelId +
                                                                                                                                                                                                '&startDate=' + startDateStr +
                                                                                                                                                                                                    '&endDate=' + endDateStr +
                                                                                                                                                                                                        '&metrics=views,estimatedMinutesWatched' +
                                                                                                                                                                                                            '&dimensions=video' +
                                                                                                                                                                                                                '&maxResults=5' +
                                                                                                                                                                                                                    '&sort=-views';
                                                                                                                                                                                                                      
                                                                                                                                                                                                                        Logger.log('URL: ' + url);
                                                                                                                                                                                                                          
                                                                                                                                                                                                                            try {
                                                                                                                                                                                                                                var response = UrlFetchApp.fetch(url, {
                                                                                                                                                                                                                                      headers: {
                                                                                                                                                                                                                                              'Authorization': 'Bearer ' + accessToken
                                                                                                                                                                                                                                                    },
                                                                                                                                                                                                                                                          muteHttpExceptions: true
                                                                                                                                                                                                                                                              });
                                                                                                                                                                                                                                                                  
                                                                                                                                                                                                                                                                      var result = JSON.parse(response.getContentText());
                                                                                                                                                                                                                                                                          Logger.log('応答: ' + JSON.stringify(result, null, 2));
                                                                                                                                                                                                                                                                              
                                                                                                                                                                                                                                                                                } catch (e) {
                                                                                                                                                                                                                                                                                    Logger.log('エラー: ' + e.message);
                                                                                                                                                                                                                                                                                      }
                                                                                                                                                                                                                                                                                      }