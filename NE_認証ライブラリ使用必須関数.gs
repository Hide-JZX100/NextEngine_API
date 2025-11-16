/**
 * 【前提条件】
 * 認証ライブラリを利用するGASプロジェクトで、以下の設定が必要です。
 * 
 * 1. Webアプリとしてデプロイ
 *    - GASエディタの「デプロイ」>「新しいデプロイ」からWebアプリとしてデプロイし、デプロイURLを取得します。
 * 
 * 2. スクリプトプロパティの設定
 *    - GASエディタの「プロジェクトの設定」(歯車アイコン) >「スクリプト プロパティ」に以下を設定します。
 * 
 *      キー           | 値
 *      -------------|------------------------------------
 *      CLIENT_ID    | ネクストエンジンアプリのクライアントID
 *      CLIENT_SECRET| ネクストエンジンアプリのクライアントシークレット
 *      REDIRECT_URI | 上記1で取得したWebアプリのデプロイURL
 * 
 * 3. ライブラリの追加
 *     - 左メニュー「ライブラリ」の「+」をクリック
 *     - 認証プロジェクトのスクリプトIDを入力
 *     - 認証プロジェクトの「プロジェクトの設定」→「スクリプトID」
 *     - 「検索」をクリック
 *     - 最新バージョンを選択(重要!)
 *     - 識別子: NEAuth と入力
 *     - 「追加」をクリック
 * 
*/

/**
 * 認証URL生成テスト
 * ライブラリの関数を呼び出して認証URLを生成
 */
function testGenerateAuthUrl() {
  console.log('=== 認証URL生成テスト ===');
  
  // 自分のプロジェクトのスクリプトプロパティを取得
  const myProperties = PropertiesService.getScriptProperties();
  
  // ライブラリに渡して認証URLを生成
  const authUrl = NEAuth.generateAuthUrl(myProperties);
  
  console.log('認証URL:', authUrl);
  console.log('');
  console.log('このURLをブラウザで開いて認証を完了してください');
  
  return authUrl;
}

/**
 * doGet関数 - ネクストエンジンからのリダイレクトを受け取る
 * Webアプリとしてデプロイされている場合に自動的に実行される
 */
function doGet(e) {
  const uid = e.parameter.uid;
  const state = e.parameter.state;
  
  console.log('doGet実行: uid=', uid, 'state=', state);
  
  if (uid && state) {
    try {
      // このプロジェクトのスクリプトプロパティを取得
      const myProperties = PropertiesService.getScriptProperties();
      const clientId = myProperties.getProperty('CLIENT_ID');
      const clientSecret = myProperties.getProperty('CLIENT_SECRET');
      
      console.log('CLIENT_ID:', clientId ? '設定済み' : '未設定');
      console.log('CLIENT_SECRET:', clientSecret ? '設定済み' : '未設定');
      
      // プロパティが設定されているか確認
      if (!clientId || !clientSecret) {
        throw new Error('スクリプトプロパティが設定されていません');
      }
      
      // アクセストークン取得のAPIリクエスト
      const url = 'https://api.next-engine.org/api_neauth';
      
      const payload = {
        'uid': uid,
        'state': state,
        'client_id': clientId,
        'client_secret': clientSecret
      };
      
      const options = {
        'method': 'POST',
        'headers': {
          'Content-Type': 'application/x-www-form-urlencoded'
        },
        'payload': Object.keys(payload).map(key => 
          encodeURIComponent(key) + '=' + encodeURIComponent(payload[key])
        ).join('&')
      };
      
      console.log('アクセストークン取得リクエスト送信中...');
      const response = UrlFetchApp.fetch(url, options);
      const responseText = response.getContentText();
      const responseData = JSON.parse(responseText);
      
      console.log('レスポンス:', responseData);
      
      if (responseData.result === 'success') {
        // トークンをこのプロジェクトのスクリプトプロパティに保存
        myProperties.setProperties({
          'ACCESS_TOKEN': responseData.access_token,
          'REFRESH_TOKEN': responseData.refresh_token,
          'TOKEN_OBTAINED_AT': new Date().getTime().toString()
        });
        
        console.log('トークンを保存しました');
        
        // 成功画面を表示
        return HtmlService.createHtmlOutput(`
          <html>
            <head>
              <title>ネクストエンジン認証完了</title>
              <style>
                body { 
                  font-family: 'Helvetica Neue', Arial, sans-serif; 
                  max-width: 600px; 
                  margin: 50px auto; 
                  padding: 20px;
                  background-color: #f5f5f5;
                }
                .container {
                  background: white;
                  padding: 30px;
                  border-radius: 10px;
                  box-shadow: 0 2px 10px rgba(0,0,0,0.1);
                }
                .success { color: #28a745; }
                .info { color: #17a2b8; }
                .code { 
                  background: #f8f9fa; 
                  padding: 10px; 
                  border-radius: 5px; 
                  font-family: monospace;
                  margin: 10px 0;
                }
              </style>
            </head>
            <body>
              <div class="container">
                <h2 class="success">✅ 認証成功!</h2>
                <p>ネクストエンジンAPIの認証が完了しました。</p>
                
                <h3>取得した情報:</h3>
                <div class="code">
                  <strong>UID:</strong> ${uid}<br>
                  <strong>State:</strong> ${state}<br>
                  <strong>Access Token:</strong> ${responseData.access_token.substring(0, 20)}...<br>
                  <strong>Refresh Token:</strong> ${responseData.refresh_token.substring(0, 20)}...
                </div>
                
                <h3 class="info">次のステップ:</h3>
                <p>GASエディタに戻り、以下の関数を実行してAPI接続をテストしてください:</p>
                <ul>
                  <li><code>testApiConnection()</code> - API接続テスト</li>
                  <li><code>testInventoryApi()</code> - 在庫取得テスト</li>
                </ul>
                
                <p><small>このページを閉じて構いません。</small></p>
              </div>
            </body>
          </html>
        `);
      } else {
        throw new Error('認証失敗: ' + JSON.stringify(responseData));
      }
      
    } catch (error) {
      console.error('認証エラー:', error.message);
      
      // エラー画面を表示
      return HtmlService.createHtmlOutput(`
        <html>
          <head>
            <title>認証エラー</title>
            <style>
              body { 
                font-family: Arial, sans-serif; 
                max-width: 600px; 
                margin: 50px auto; 
                padding: 20px;
              }
              .error { color: #dc3545; }
              .code {
                background: #f8f9fa;
                padding: 10px;
                border-radius: 5px;
                font-family: monospace;
                margin: 10px 0;
              }
            </style>
          </head>
          <body>
            <h2 class="error">❌ 認証エラー</h2>
            <p>認証処理中にエラーが発生しました:</p>
            <div class="code">${error.message}</div>
            <p>GASエディタでログを確認してください。</p>
            <p>スクリプトプロパティが正しく設定されているか確認してください:</p>
            <ul>
              <li>CLIENT_ID</li>
              <li>CLIENT_SECRET</li>
              <li>REDIRECT_URI</li>
            </ul>
          </body>
        </html>
      `);
    }
  } else {
    // パラメータエラー画面
    return HtmlService.createHtmlOutput(`
      <html>
        <head>
          <title>パラメータエラー</title>
          <style>
            body { 
              font-family: Arial, sans-serif; 
              max-width: 600px; 
              margin: 50px auto; 
              padding: 20px;
            }
            .error { color: #dc3545; }
          </style>
        </head>
        <body>
          <h2 class="error">❌ パラメータエラー</h2>
          <p>必要なパラメータ（uid、state）が見つかりません。</p>
          <p>認証URLから正しくリダイレクトされていない可能性があります。</p>
          <p>GASエディタで <code>testGenerateAuthUrl()</code> を実行して、正しい認証URLを取得してください。</p>
        </body>
      </html>
    `);
  }
}