/**
=============================================================================
ネクストエンジン認証ライブラリ (NEAuth)
=============================================================================
 * バージョン: 4.0
 * 
 * 【概要】
 * Google Apps Script (GAS) からネクストエンジンAPIを利用する際の認証プロセスを補助するライブラリです。
 * 認証URLの生成やトークンの有効性検証など、共通化できる処理を提供します。
 * 
 * このライブラリは認証フローの一部を簡略化しますが、トークンを取得するための
 * `doGet(e)` 関数の実装とWebアプリとしてのデプロイは、各プロジェクト側で行う必要があります。
 * 
 * ---
 * 
 * 【前提条件】
 * このライブラリを利用するGASプロジェクトで、以下の設定が必要です。
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
 * 3. `doGet(e)` 関数の実装
 *    - 認証後のリダイレクトを受け取り、アクセストークンを取得・保存する `doGet(e)` 関数をプロジェクトに実装する必要があります。
 * 
 * ---
 * 
 * 【利用手順】
 * 1. `generateAuthUrl()` を実行して認証URLを生成します。
 *    ```javascript
 *    function getAuthUrl() {
 *      const props = PropertiesService.getScriptProperties();
 *      const authUrl = NEAuth.generateAuthUrl(props);
 *      console.log(authUrl);
 *    }
 *    ```
 * 
 * 2. 生成されたURLにアクセスし、ネクストエンジンで認証を行います。
 * 
 * 3. 認証後、`REDIRECT_URI` にリダイレクトされ、プロジェクトに実装した `doGet(e)` が実行されてトークンが保存されます。
 * 
 * 4. `testApiConnection()` を実行して、保存されたトークンでAPIに接続できるかテストします。
 *    ```javascript
 *    function checkConnection() {
 *      const props = PropertiesService.getScriptProperties();
 *      NEAuth.testApiConnection(props);
 *    }
 *    ```
 * 
 * ---
 * 
 * 【注意事項】
 * - アクセストークンやクライアントシークレットは機密情報です。第三者に漏洩しないよう厳重に管理してください。
 * - このライブラリは認証の補助を目的としており、APIの各エンドポイントの呼び出し機能は含みません。
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
 * 認証URLを生成してログ出力
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
 * 保存されたトークンでAPI接続をテスト
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