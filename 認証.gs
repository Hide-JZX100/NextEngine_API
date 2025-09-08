/**
 * ネクストエンジンAPI認証テスト用スクリプト
 * 
 * 【目的】
 * ネクストエンジンAPIとの認証を確立し、アクセストークンの取得を行う
 * 
 * 【機能】
 * 1. OAuth2認証フローの実行
 * 2. アクセストークンとリフレッシュトークンの取得・保存
 * 3. トークンの有効性確認
 * 
 * 【設定方法】
 * スクリプトプロパティに以下の値を設定してください：
 * - CLIENT_ID: ネクストエンジンのクライアントID
 * - CLIENT_SECRET: ネクストエンジンのクライアントシークレット
 * - REDIRECT_URI: リダイレクトURI（スクリーンショット参照）
 * 
 * 【注意事項】
 * - テスト環境での動作確認を前提としています
 * - 本番環境での使用前に十分なテストを行ってください
 * - アクセストークンは安全に管理してください
 */

// ネクストエンジンAPIのエンドポイント
const NE_BASE_URL = 'https://base.next-engine.org';
const NE_API_URL = 'https://api.next-engine.org';

/**
 * スクリプトプロパティから設定値を取得
 */
function getScriptProperties() {
  const properties = PropertiesService.getScriptProperties();
  
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
 */
function generateAuthUrl() {
  try {
    const config = getScriptProperties();
    
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
 * ステップ2: uidとstateを使用してアクセストークンを取得
 * @param {string} uid - 認証後に取得されるuid
 * @param {string} state - 認証後に取得されるstate
 */
function getAccessToken(uid, state) {
  try {
    if (!uid || !state) {
      throw new Error('uid と state が必要です');
    }
    
    const config = getScriptProperties();
    
    // アクセストークン取得のリクエスト
    const url = `${NE_API_URL}/api_v1_neauth.json`;
    
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
    const response = UrlFetchApp.fetch(url, options);
    const responseText = response.getContentText();
    
    console.log('レスポンス:', responseText);
    
    if (response.getResponseCode() !== 200) {
      throw new Error(`APIリクエストが失敗しました。ステータス: ${response.getResponseCode()}`);
    }
    
    const responseData = JSON.parse(responseText);
    
    if (responseData.result === 'success') {
      // トークンをスクリプトプロパティに保存
      const properties = PropertiesService.getScriptProperties();
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
 */
function testApiConnection() {
  try {
    const properties = PropertiesService.getScriptProperties();
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
    const response = UrlFetchApp.fetch(url, options);
    const responseText = response.getContentText();
    
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
 */
function showStoredTokens() {
  const properties = PropertiesService.getScriptProperties();
  
  console.log('=== 保存されているトークン情報 ===');
  console.log('ACCESS_TOKEN:', properties.getProperty('ACCESS_TOKEN') || '未設定');
  console.log('REFRESH_TOKEN:', properties.getProperty('REFRESH_TOKEN') || '未設定');
  console.log('TOKEN_OBTAINED_AT:', properties.getProperty('TOKEN_OBTAINED_AT') || '未設定');
  console.log('TOKEN_UPDATED_AT:', properties.getProperty('TOKEN_UPDATED_AT') || '未設定');
}

/**
 * スクリプトプロパティをクリア（テスト用）
 */
function clearProperties() {
  const properties = PropertiesService.getScriptProperties();
  properties.deleteProperty('ACCESS_TOKEN');
  properties.deleteProperty('REFRESH_TOKEN');
  properties.deleteProperty('TOKEN_OBTAINED_AT');
  properties.deleteProperty('TOKEN_UPDATED_AT');
  
  console.log('トークン情報をクリアしました');
}

/**
 * 認証フロー全体のガイド表示
 */
function showAuthGuide() {
  console.log('=== ネクストエンジンAPI認証ガイド ===');
  console.log('');
  console.log('【事前準備】');
  console.log('1. スクリプトプロパティに以下を設定:');
  console.log('   - CLIENT_ID: あなたのクライアントID');
  console.log('   - CLIENT_SECRET: あなたのクライアントシークレット');
  console.log('   - REDIRECT_URI: リダイレクトURI');
  console.log('');
  console.log('【認証手順】');
  console.log('1. generateAuthUrl() を実行して認証URLを取得');
  console.log('2. ブラウザで認証URLにアクセス');
  console.log('3. ネクストエンジンにログイン');
  console.log('4. リダイレクト後のURLから uid と state を取得');
  console.log('5. getAccessToken(uid, state) を実行');
  console.log('6. testApiConnection() で動作確認');
  console.log('');
  console.log('【その他の関数】');
  console.log('- showStoredTokens(): 保存済みトークン情報表示');
  console.log('- clearProperties(): トークン情報クリア');
}