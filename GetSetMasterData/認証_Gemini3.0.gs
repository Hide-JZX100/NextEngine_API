/**
 * ネクストエンジンAPI OAuth2認証用スクリプト
 * * 【概要】
 * ネクストエンジンの認証フロー（UID/State方式）を処理し、
 * アクセストークンとリフレッシュトークンを管理します。
 * * 【作成日】2025-10-17
 * 【更新日】2025-11-23 エンドポイント修正 (api_neauth)
 * 【更新日】2025-11-23 フィールド名修正 (pic_name, company_name)
 * 【作成者】 & Gemini
 */

// ==================================================
// 設定・定数
// ==================================================

/**
 * スクリプトプロパティから設定値を取得するヘルパー関数
 * ※事前に[プロジェクトの設定] > [スクリプトプロパティ]に以下のキーで値を設定してください。
 * - NEXT_ENGINE_CLIENT_ID
 * - NEXT_ENGINE_CLIENT_SECRET
 * - NEXT_ENGINE_REDIRECT_URI (GASのウェブアプリURL)
 */
function getProperty(key) {
  return PropertiesService.getScriptProperties().getProperty(key);
}

// ==================================================
// 1. 認証URL生成
// ==================================================

/**
 * ユーザー認証用のURLを生成し、ログに出力する関数
 * この関数を実行し、ログに出たURLにアクセスしてください。
 */
function generateAuthUrl() {
  const clientId = getProperty('NEXT_ENGINE_CLIENT_ID');
  const redirectUri = getProperty('NEXT_ENGINE_REDIRECT_URI'); // デプロイしたGASのURL

  if (!clientId || !redirectUri) {
    Logger.log('エラー: スクリプトプロパティに CLIENT_ID または REDIRECT_URI が設定されていません。');
    return;
  }

  // ネクストエンジンのログイン画面URL
  // ユーザーがログイン承認を行うと、REDIRECT_URIにuidとstate付きで戻ってきます
  const authUrl = `https://base.next-engine.org/users/sign_in/?client_id=${clientId}&redirect_uri=${encodeURIComponent(redirectUri)}`;
  
  Logger.log('以下のURLをブラウザで開いてログインしてください:');
  Logger.log(authUrl);
}

// ==================================================
// 2. リダイレクト受け取り (Webアプリ)
// ==================================================

/**
 * WebアプリへのGETリクエストを処理する関数 (GAS標準関数)
 * ネクストエンジンからのリダイレクトを受け取り、uidとstateを抽出します。
 */
function doGet(e) {
  // パラメータチェック
  if (!e || !e.parameter) {
    return HtmlService.createHtmlOutput('パラメータが見つかりません。');
  }

  const uid = e.parameter.uid;
  const state = e.parameter.state;

  if (uid && state) {
    // uidとstateを使ってトークンを取得しにいく
    try {
      const result = getAccessToken(uid, state);
      if (result) {
        return HtmlService.createHtmlOutput('認証に成功しました！タブを閉じてGASエディタに戻り、testApiConnectionを実行してください。');
      } else {
        return HtmlService.createHtmlOutput('トークンの取得に失敗しました。ログを確認してください。');
      }
    } catch (error) {
      return HtmlService.createHtmlOutput('エラーが発生しました: ' + error.message);
    }
  } else {
    return HtmlService.createHtmlOutput('UIDまたはStateが不足しています。');
  }
}

// ==================================================
// 3. トークン取得・保存
// ==================================================

/**
 * uidとstateを使用してアクセストークンを取得し、保存する関数
 * * @param {string} uid - ネクストエンジンから返されたuid
 * @param {string} state - ネクストエンジンから返されたstate
 * @return {boolean} 成功した場合はtrue
 */
function getAccessToken(uid, state) {
  // 【修正済み】アクセストークン取得の正しいエンドポイント
  const url = 'https://api.next-engine.org/api_neauth';
  
  const payload = {
    'client_id': getProperty('NEXT_ENGINE_CLIENT_ID'),
    'client_secret': getProperty('NEXT_ENGINE_CLIENT_SECRET'),
    'uid': uid,
    'state': state
  };

  const options = {
    'method': 'post',
    'payload': payload
  };

  try {
    Logger.log('トークン取得リクエストを送信します...');
    const response = UrlFetchApp.fetch(url, options);
    const json = JSON.parse(response.getContentText());

    if (json.result === 'success') {
      // トークンを保存
      saveTokens(json.access_token, json.refresh_token);
      Logger.log('トークン取得成功: アクセストークンとリフレッシュトークンを保存しました。');
      
      // 【追加】ご指摘いただいた出力パラメータをここでログ確認
      // api_neauthは認証時にユーザー情報を返してくれます
      Logger.log('--- 認証ユーザー情報 ---');
      Logger.log('企業名: ' + json.company_name);
      Logger.log('担当者名: ' + json.pic_name);
      Logger.log('----------------------');
      
      return true;
    } else {
      Logger.log('トークン取得失敗: ' + JSON.stringify(json));
      return false;
    }
  } catch (e) {
    Logger.log('通信エラー: ' + e.toString());
    throw e;
  }
}

/**
 * トークンをスクリプトプロパティに保存する共通関数
 * * @param {string} accessToken 
 * @param {string} refreshToken 
 */
function saveTokens(accessToken, refreshToken) {
  const props = PropertiesService.getScriptProperties();
  const data = {};
  
  if (accessToken) data['NEXT_ENGINE_ACCESS_TOKEN'] = accessToken;
  if (refreshToken) data['NEXT_ENGINE_REFRESH_TOKEN'] = refreshToken;
  
  props.setProperties(data);
}

// ==================================================
// 4. API接続テスト・汎用リクエスト
// ==================================================

/**
 * API接続テスト用関数
 * ログインユーザー情報を取得して、トークンが有効か確認します。
 */
function testApiConnection() {
  // ログインユーザー情報取得API
  const apiPath = 'api_v1_login_user/info'; 
  
  // 下記の汎用関数を使ってリクエスト
  const response = callNextEngineApi(apiPath, {});
  
  if (response && response.result === 'success') {
    Logger.log('APIテスト成功！');

    // 【修正】ネクストエンジンの仕様に合わせてフィールド名を修正
    // user_name ではなく pic_name (担当者名) が一般的です
    if (response.data && response.data[0]) {
      const user = response.data[0];
      Logger.log('担当者名 (pic_name): ' + user.pic_name);
      Logger.log('担当者カナ (pic_kana): ' + user.pic_kana);
      
      // デバッグ用：念のため返ってきた全データをログに出して構造を確認できるようにします
      Logger.log('▼取得データ全量:');
      Logger.log(JSON.stringify(user));
    } else {
      Logger.log('データ取得成功ですが、data配列が空です: ' + JSON.stringify(response));
    }
  } else {
    Logger.log('APIテスト失敗、またはデータが取得できませんでした。');
  }
}

/**
 * ネクストエンジンAPIを呼び出す汎用関数
 * ★重要: APIレスポンスに新しいトークンが含まれていれば自動更新します。
 * * @param {string} path - APIのエンドポイント (例: 'api_v1_master_goods/search')
 * @param {Object} params - APIに送信するパラメータオブジェクト
 * @return {Object} APIからのレスポンスJSON
 */
function callNextEngineApi(path, params) {
  const accessToken = getProperty('NEXT_ENGINE_ACCESS_TOKEN');
  const refreshToken = getProperty('NEXT_ENGINE_REFRESH_TOKEN');

  if (!accessToken || !refreshToken) {
    Logger.log('エラー: トークンがありません。generateAuthUrlから認証を行ってください。');
    return null;
  }

  const url = `https://api.next-engine.org/${path}`;

  // 必須パラメータであるトークンを付与
  const payload = {
    'access_token': accessToken,
    'refresh_token': refreshToken,
    ...params // 他の検索条件などを結合
  };

  const options = {
    'method': 'post',
    'payload': payload,
    'muteHttpExceptions': true // エラーでもレスポンス内容を確認するため
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const json = JSON.parse(response.getContentText());

    // ★ここが重要: API利用時にトークンが更新されていたら保存し直す
    // ネクストエンジンは有効期限切れが近いや切れた場合に、新しいトークンをレスポンスに含めて返します
    if (json.access_token || json.refresh_token) {
      Logger.log('APIレスポンスによりトークンを更新します。');
      saveTokens(json.access_token, json.refresh_token);
    }

    // API側のエラーハンドリング
    if (json.result !== 'success') {
      Logger.log('APIエラー: ' + json.code + ' - ' + json.message);
      
      // 認証エラー(コード400系の一部)の場合、再認証を促すログを出すなどの分岐も可能
      if (json.code === '002003' || json.code === '002004') {
         Logger.log('認証トークンが無効です。generateAuthUrlから再認証してください。');
      }
    }

    return json;

  } catch (e) {
    Logger.log('API呼び出し中の例外: ' + e.toString());
    return null;
  }
}