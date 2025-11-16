/**
=============================================================================
ネクストエンジンAPI認証ライブラリ v2.0
=============================================================================

* 【目的】
* ネクストエンジンAPIとの認証を確立し、アクセストークンの取得を行う
* ライブラリとして他のプロジェクトから呼び出し可能
* 
* 【機能】
* 1. OAuth2認証フローの実行
* 2. アクセストークンとリフレッシュトークンの取得・保存
* 3. トークンの有効性確認
* 4. 外部プロジェクトのスクリプトプロパティに対応

【v2.0の変更点】
* - 全関数で外部プロパティを受け取れるように改良
* - 後方互換性を維持(引数なしでも動作)
* - ライブラリとして使用する際の利便性向上

【スクリプトプロパティの設定方法】
呼び出し元のGASプロジェクトで以下を設定:
1. GASエディタで「プロジェクトの設定」を開く（歯車のアイコン）
2. 「スクリプトプロパティ」セクションまでスクロール
3. 「スクリプトプロパティの追加」をクリックし、以下のキーと値を設定

   キー                     | 値
   -------------------------|------------------------------------
   CLIENT_ID                | ネクストエンジンのクライアントID
   CLIENT_SECRET            | ネクストエンジンのクライアントシークレット
   REDIRECT_URI             | ネクストエンジンのリダイレクトURI

* 
* 【注意事項】
* - テスト環境での動作確認を前提としています
* - 本番環境での使用前に十分なテストを行ってください
* - アクセストークンは安全に管理してください

【ライブラリとしての使用方法】
呼び出し元プロジェクトで:
```javascript
function myFunction() {
  const props = PropertiesService.getScriptProperties();
  const authUrl = NEAuth.generateAuthUrl(props);
  console.log(authUrl);
}
```

【スタンドアロンでの使用方法】
このプロジェクト内で直接実行する場合:
```javascript
function test() {
  const authUrl = generateAuthUrl();  // 引数なしで自身のプロパティを使用
  console.log(authUrl);
}
```

=============================================================================
*/

// ネクストエンジンAPIのエンドポイント
const NE_BASE_URL = 'https://base.next-engine.org';  // 認証画面用
const NE_API_URL = 'https://api.next-engine.org';    // API呼び出し用

/**
 * スクリプトプロパティから設定値を取得
 * 
 * @param {PropertiesService.Properties} externalProperties - 外部から渡されたプロパティ(オプション)
 * @return {Object} 設定オブジェクト {clientId, clientSecret, redirectUri}
 */
function getScriptProperties(externalProperties) {
  // 外部からプロパティが渡された場合はそれを使用、なければ自身のプロパティを使用
  const properties = externalProperties || PropertiesService.getScriptProperties();
  
  const clientId = properties.getProperty('CLIENT_ID');
  const clientSecret = properties.getProperty('CLIENT_SECRET');
  const redirectUri = properties.getProperty('REDIRECT_URI');
  
  if (!clientId || !clientSecret || !redirectUri) {
    throw new Error('必要なスクリプトプロパティが設定されていません。CLIENT_ID, CLIENT_SECRET, REDIRECT_URIを設定してください。');
  }
  
  return {
    clientId,
    clientSecret,
    redirectUri
  };
}

/**
 * ステップ1: 認証URLを生成してログ出力
 * 手動でブラウザでアクセスして認証を行う
 * 
 * @param {PropertiesService.Properties} externalProperties - 外部から渡されたプロパティ(オプション)
 * @return {string} 認証URL
 */
function generateAuthUrl(externalProperties) {
  try {
    const config = getScriptProperties(externalProperties);
    
    // ネクストエンジンの認証URL
    const authUrl = `${NE_BASE_URL}/users/sign_in?client_id=${config.clientId}&redirect_uri=${encodeURIComponent(config.redirectUri)}`;
    
    console.log('=== ネクストエンジン認証URL ===');
    console.log(authUrl);
    console.log('');
    console.log('上記URLをブラウザでアクセスしてください');
    console.log('認証後、リダイレクトURLに uid と state パラメータが付与されて戻ってきます');
    console.log('例: https://your-redirect-uri?uid=XXXXX&state=YYYYY');
    
    return authUrl;
    
  } catch (error) {
    console.error('認証URL生成エラー:', error.message);
    throw error;
  }
}

/**
 * WebアプリとしてのGETリクエスト処理
 * ネクストエンジンからのリダイレクト時にuidとstateを受け取る
 * 
 * 注意: この関数は呼び出し元のプロジェクトで自動的に実行されるため、
 * 外部プロパティの指定は不要です。
 * 
 * @param {Object} e - イベントオブジェクト
 * @return {HtmlService.HtmlOutput} HTML出力
 */
function doGet(e) {
  const uid = e.parameter.uid;
  const state = e.parameter.state;
  
  if (uid && state) {
    try {
      // 自動的にアクセストークンを取得
      // doGet()は呼び出し元で実行されるため、自動的に呼び出し元のプロパティが使用される
      const result = getAccessToken(uid, state);
      
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
              <h2 class="success">✅ 認証成功！</h2>
              <p>ネクストエンジンAPIの認証が完了しました。</p>
              
              <h3>取得した情報:</h3>
              <div class="code">
                <strong>UID:</strong> ${uid}<br>
                <strong>State:</strong> ${state}<br>
                <strong>Access Token:</strong> ${result.access_token.substring(0, 20)}...<br>
                <strong>Refresh Token:</strong> ${result.refresh_token.substring(0, 20)}...
              </div>
              
              <h3 class="info">次のステップ:</h3>
              <p>GASエディタに戻り、<code>testApiConnection()</code> を実行してAPI接続をテストしてください。</p>
              
              <p><small>このページを閉じて構いません。</small></p>
            </div>
          </body>
        </html>
      `);
      
    } catch (error) {
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
            </style>
          </head>
          <body>
            <h2 class="error">❌ 認証エラー</h2>
            <p>認証処理中にエラーが発生しました:</p>
            <p><strong>${error.message}</strong></p>
            <p>GASエディタでログを確認してください。</p>
          </body>
        </html>
      `);
    }
  } else {
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
          <p>GASエディタで <code>generateAuthUrl()</code> を実行して、正しい認証URLを取得してください。</p>
        </body>
      </html>
    `);
  }
}

/**
 * ステップ2: uidとstateを使用してアクセストークンを取得
 * 
 * @param {string} uid - 認証後に取得されるuid
 * @param {string} state - 認証後に取得されるstate
 * @param {PropertiesService.Properties} externalProperties - 外部から渡されたプロパティ(オプション)
 * @return {Object} トークン情報 {access_token, refresh_token}
 */
function getAccessToken(uid, state, externalProperties) {
  try {
    if (!uid || !state) {
      throw new Error('uid と state が必要です');
    }
    
    const config = getScriptProperties(externalProperties);
    
    // アクセストークン取得のリクエスト
    const url = `${NE_API_URL}/api_neauth`;
    
    const payload = {
      'uid': uid,
      'state': state,
      'client_id': config.clientId,
      'client_secret': config.clientSecret
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
    console.log('使用URL:', url);
    console.log('ペイロード:', payload);
    const response = UrlFetchApp.fetch(url, options);
    const responseText = response.getContentText();
    
    console.log('レスポンス:', responseText);
    
    if (response.getResponseCode() !== 200) {
      throw new Error(`APIリクエストが失敗しました。ステータス: ${response.getResponseCode()}`);
    }
    
    const responseData = JSON.parse(responseText);
    
    if (responseData.result === 'success') {
      // トークンをスクリプトプロパティに保存
      const properties = externalProperties || PropertiesService.getScriptProperties();
      properties.setProperties({
        'ACCESS_TOKEN': responseData.access_token,
        'REFRESH_TOKEN': responseData.refresh_token,
        'TOKEN_OBTAINED_AT': new Date().getTime().toString()
      });
      
      console.log('=== 認証成功 ===');
      console.log('アクセストークン:', responseData.access_token);
      console.log('リフレッシュトークン:', responseData.refresh_token);
      console.log('トークンをスクリプトプロパティに保存しました');
      
      return responseData;
    } else {
      throw new Error(`認証失敗: ${JSON.stringify(responseData)}`);
    }
    
  } catch (error) {
    console.error('アクセストークン取得エラー:', error.message);
    throw error;
  }
}

/**
 * ステップ3: 保存されたトークンでAPI接続をテスト
 * ユーザー情報を取得して認証が正常に動作しているか確認
 * 
 * @param {PropertiesService.Properties} externalProperties - 外部から渡されたプロパティ(オプション)
 * @return {Object} ユーザー情報
 */
function testApiConnection(externalProperties) {
  try {
    const properties = externalProperties || PropertiesService.getScriptProperties();
    const accessToken = properties.getProperty('ACCESS_TOKEN');
    const refreshToken = properties.getProperty('REFRESH_TOKEN');
    
    if (!accessToken || !refreshToken) {
      throw new Error('アクセストークンが見つかりません。先にgetAccessToken()を実行してください。');
    }
    
    // ログインユーザー情報取得APIでテスト
    const url = `${NE_API_URL}/api_v1_login_user/info`;
    
    const payload = {
      'access_token': accessToken,
      'refresh_token': refreshToken
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
    
    console.log('API接続テスト実行中...');
    console.log('使用URL:', url);
    console.log('ペイロード:', payload);
    
    const response = UrlFetchApp.fetch(url, options);
    const responseText = response.getContentText();
    
    console.log('レスポンスコード:', response.getResponseCode());
    console.log('API レスポンス:', responseText);
    
    const responseData = JSON.parse(responseText);
    
    if (responseData.result === 'success') {
      console.log('=== API接続テスト成功 ===');
      console.log('ユーザー情報:', responseData.data);
      
      // トークンが更新された場合は保存
      if (responseData.access_token && responseData.refresh_token) {
        properties.setProperties({
          'ACCESS_TOKEN': responseData.access_token,
          'REFRESH_TOKEN': responseData.refresh_token,
          'TOKEN_UPDATED_AT': new Date().getTime().toString()
        });
        console.log('トークンが更新されました');
      }
      
      return responseData.data;
    } else {
      throw new Error(`API接続テスト失敗: ${JSON.stringify(responseData)}`);
    }
    
  } catch (error) {
    console.error('API接続テストエラー:', error.message);
    throw error;
  }
}

/**
 * 在庫APIのテスト（最終目標に向けて）
 * 商品マスタ情報の取得をテスト
 * 
 * @param {PropertiesService.Properties} externalProperties - 外部から渡されたプロパティ(オプション)
 * @return {Array} 商品情報の配列
 */
function testInventoryApi(externalProperties) {
  try {
    const properties = externalProperties || PropertiesService.getScriptProperties();
    const accessToken = properties.getProperty('ACCESS_TOKEN');
    const refreshToken = properties.getProperty('REFRESH_TOKEN');
    
    if (!accessToken || !refreshToken) {
      throw new Error('アクセストークンが見つかりません。');
    }
    
    // 商品マスタ検索API
    const url = `${NE_API_URL}/api_v1_master_goods/search`;
    
    const payload = {
      'access_token': accessToken,
      'refresh_token': refreshToken,
      'fields': 'goods_id,goods_name,stock_quantity', // 必要な項目のみ
      'limit': '5' // テスト用に5件に制限
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
    
    console.log('在庫API テスト実行中...');
    console.log('使用URL:', url);
    
    const response = UrlFetchApp.fetch(url, options);
    const responseText = response.getContentText();
    
    console.log('レスポンスコード:', response.getResponseCode());
    console.log('在庫API レスポンス:', responseText);
    
    const responseData = JSON.parse(responseText);
    
    if (responseData.result === 'success') {
      console.log('=== 在庫API テスト成功 ===');
      console.log('取得商品数:', responseData.count);
      console.log('商品情報:', responseData.data);
      
      // トークンが更新された場合は保存
      if (responseData.access_token && responseData.refresh_token) {
        properties.setProperties({
          'ACCESS_TOKEN': responseData.access_token,
          'REFRESH_TOKEN': responseData.refresh_token,
          'TOKEN_UPDATED_AT': new Date().getTime().toString()
        });
        console.log('トークンが更新されました');
      }
      
      return responseData.data;
    } else {
      console.log('在庫API エラー:', responseData);
      
      if (responseData.code === '004001') {
        console.log('');
        console.log('【解決方法】');
        console.log('ネクストエンジンの「アプリを作る」→アプリ編集→「API」タブで');
        console.log('「商品マスタ検索」の権限を有効にしてください');
      }
      
      return responseData;
    }
    
  } catch (error) {
    console.error('在庫API テストエラー:', error.message);
    throw error;
  }
}

/**
 * 保存されているトークン情報を表示
 * 
 * @param {PropertiesService.Properties} externalProperties - 外部から渡されたプロパティ(オプション)
 */
function showStoredTokens(externalProperties) {
  const properties = externalProperties || PropertiesService.getScriptProperties();
  
  console.log('=== 保存されているトークン情報 ===');
  console.log('ACCESS_TOKEN:', properties.getProperty('ACCESS_TOKEN') || '未設定');
  console.log('REFRESH_TOKEN:', properties.getProperty('REFRESH_TOKEN') || '未設定');
  console.log('TOKEN_OBTAINED_AT:', properties.getProperty('TOKEN_OBTAINED_AT') || '未設定');
  console.log('TOKEN_UPDATED_AT:', properties.getProperty('TOKEN_UPDATED_AT') || '未設定');
}

/**
 * スクリプトプロパティをクリア（テスト用）
 * 
 * @param {PropertiesService.Properties} externalProperties - 外部から渡されたプロパティ(オプション)
 */
function clearProperties(externalProperties) {
  const properties = externalProperties || PropertiesService.getScriptProperties();
  properties.deleteProperty('ACCESS_TOKEN');
  properties.deleteProperty('REFRESH_TOKEN');
  properties.deleteProperty('TOKEN_OBTAINED_AT');
  properties.deleteProperty('TOKEN_UPDATED_AT');
  
  console.log('トークン情報をクリアしました');
}

/**
 * 現在のデプロイURLとスクリプトプロパティを確認
 * 
 * @param {PropertiesService.Properties} externalProperties - 外部から渡されたプロパティ(オプション)
 */
function checkCurrentDeployment(externalProperties) {
  console.log('=== 現在のデプロイ情報確認 ===');
  
  // スクリプトプロパティの確認
  const properties = externalProperties || PropertiesService.getScriptProperties();
  const clientId = properties.getProperty('CLIENT_ID');
  const redirectUri = properties.getProperty('REDIRECT_URI');
  const clientSecret = properties.getProperty('CLIENT_SECRET');
  
  console.log('CLIENT_ID:', clientId || '未設定');
  console.log('REDIRECT_URI:', redirectUri || '未設定');
  console.log('CLIENT_SECRET:', clientSecret ? '設定済み' : '未設定');
  
  console.log('');
  console.log('=== 確認ポイント ===');
  console.log('1. GASのデプロイメニューで最新のWebアプリURLを確認');
  console.log('2. そのURLがREDIRECT_URIと完全一致しているか');
  console.log('3. ネクストエンジンの本番環境設定でリダイレクトURIが同じか');
  console.log('4. クライアントIDが本番環境用になっているか');
  
  // 認証URL生成（デバッグ用）
  if (clientId && redirectUri) {
    const authUrl = `${NE_BASE_URL}/users/sign_in?client_id=${clientId}&redirect_uri=${encodeURIComponent(redirectUri)}`;
    console.log('');
    console.log('生成される認証URL:');
    console.log(authUrl);
  }
}

/**
 * 認証フロー全体のガイド表示
 */
function showAuthGuide() {
  console.log('=== ネクストエンジンAPI認証ガイド v2.0 ===');
  console.log('');
  console.log('【ライブラリとして使用する場合】');
  console.log('1. 呼び出し元のプロジェクトでスクリプトプロパティを設定:');
  console.log('   - CLIENT_ID');
  console.log('   - CLIENT_SECRET');
  console.log('   - REDIRECT_URI');
  console.log('');
  console.log('2. 呼び出し元のプロジェクトをWebアプリとしてデプロイ');
  console.log('');
  console.log('3. 呼び出し元で以下のように使用:');
  console.log('   const props = PropertiesService.getScriptProperties();');
  console.log('   const authUrl = NEAuth.generateAuthUrl(props);');
  console.log('');
  console.log('【スタンドアロンで使用する場合】');
  console.log('1. このプロジェクトでスクリプトプロパティを設定');
  console.log('2. このプロジェクトをWebアプリとしてデプロイ');
  console.log('3. generateAuthUrl() を実行（引数なし）');
  console.log('');
  console.log('【認証手順】');
  console.log('1. generateAuthUrl() で認証URLを取得');
  console.log('2. ブラウザで認証URLにアクセス');
  console.log('3. ネクストエンジンにログイン');
  console.log('4. 自動的にアクセストークンが取得されます');
  console.log('5. testApiConnection() で動作確認');
}