// ============================================================
// 02_API通信.gs
// ネクストエンジン Amazon受注データ取得 - API通信モジュール
// ============================================================

/**
 * ネクストエンジンAPIへの単一リクエストを実行する基底関数
 *
 * 【処理フロー】
 * 1. スクリプトプロパティから ACCESS_TOKEN と REFRESH_TOKEN を取得
 * 2. ペイロードに access_token と refresh_token の両方を追加（必須）
 * 3. UrlFetchApp.fetch() でPOSTリクエストを送信
 * 4. レスポンスをJSONパース
 * 5. result === 'success' 以外はエラーとしてthrow
 * 6. レスポンス内のトークン（access_token, refresh_token）をスクリプトプロパティへ保存
 * 7. レスポンスデータを返す
 *
 * 【使用タイミング】
 * - ネクストエンジンの各種API（受注検索など）をコールする際
 *
 * @param {string} endpoint - APIエンドポイント（例: '/api_v1_receiveorder_base/search'）
 * @param {Object} payload - リクエストパラメータ（トークン以外）
 * @returns {Object} APIレスポンスオブジェクト
 */
function callNeApi(endpoint, payload) {
  const props = PropertiesService.getScriptProperties();
  const accessToken = props.getProperty('ACCESS_TOKEN');
  const refreshToken = props.getProperty('REFRESH_TOKEN');

  if (!accessToken || !refreshToken) {
    throw new Error('アクセストークンまたはリフレッシュトークンが設定されていません。00_NE_認証ライブラリ使用必須関数 を用いて再度認証を行ってください。');
  }

  // ペイロードにトークンを追加
  payload.access_token = accessToken;
  payload.refresh_token = refreshToken;

  const url = NE_API_BASE_URL + endpoint;

  const options = {
    method: 'post',
    payload: payload,
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);
  const responseData = JSON.parse(response.getContentText());

  // API側でエラーが返された場合
  if (responseData.result !== 'success') {
    throw new Error('APIリクエストエラー: ' + JSON.stringify(responseData));
  }

  // 新しいトークンをスクリプトプロパティへ上書き保存（ローテーション対応）
  props.setProperties({
    'ACCESS_TOKEN': responseData.access_token,
    'REFRESH_TOKEN': responseData.refresh_token,
    'TOKEN_UPDATED_AT': new Date().getTime().toString()
  });

  return responseData;
}

/**
 * 指定した出荷確定日で出荷確定済み受注を全件取得する（ページネーション対応）
 *
 * 【処理フロー】
 * 1. offset = 0 で開始
 * 2. callNeApi() を呼び出し（フィルタ: 店舗ID・受注状態・出荷確定日・指定フィールド）
 * 3. 取得データを結果配列に追加
 * 4. 取得件数 < API_PAGE_LIMIT であればループ終了
 * 5. offset += API_PAGE_LIMIT してウェイトを入れたのちループ継続
 * 6. 全件を返す
 *
 * 【使用タイミング】
 * - メインのデータ取得・転記処理（Phase 4）の実行時
 * - 動作確認テスト実行時
 *
 * @param {string} targetDateStr - 取得対象日（形式: "YYYY-MM-DD"）
 * @returns {Array} 受注データの配列（全ページ分）
 */
function fetchOrdersByShipDate(targetDateStr) {
  let allOrders = [];
  let offset = 0;

  // スプレッドシートに出力するフィールドのみをAPIから取得するようカンマ区切りに変換
  const fields = CONFIG_FIELDS.map(function (f) { return f.api; }).join(',');

  while (true) {
    const payload = {
      'fields': fields,
      'receive_order_shop_id-eq': getShopId(),
      'receive_order_order_status_id-eq': ORDER_STATUS_SHIPPED,
      'receive_order_send_date-eq': targetDateStr,
      'limit': API_PAGE_LIMIT,
      'offset': offset
    };

    // APIコール実行
    const response = callNeApi(NE_ENDPOINT_ORDER_SEARCH, payload);
    const dataList = response.data || [];

    // 取得したデータを全件保持用の配列に結合
    allOrders = allOrders.concat(dataList);

    // 今回取得したデータ件数がlimitより少なければ最終ページと判断しループを抜ける
    if (dataList.length < API_PAGE_LIMIT) {
      break;
    }

    // 次ページ用のoffsetを加算し、APIのクォータ制限回避のため一定時間待機
    offset += API_PAGE_LIMIT;
    Utilities.sleep(API_PAGE_WAIT_MS);
  }

  return allOrders;
}

/**
 * 【Phase 2 テスト用】動作確認関数
 * 開発モード（楽天市場店）で過去の任意の日付を指定し、API通信テストを行う
 *
 * 【処理フロー】
 * 1. スクリプトプロパティ DEVELOPMENT_MODE を "true" に設定（テスト後元に戻す）
 * 2. testDate で指定した日付のデータを fetchOrdersByShipDate() で取得
 * 3. 取得件数と先頭1件のデータをコンソールに出力
 *
 * 【使用タイミング】
 * - Phase 2実装完了後の動作確認時（手動で実行）
 */
function testPhase2() {
  console.log('=== Phase 2 テスト開始 ===');

  const props = PropertiesService.getScriptProperties();
  const originalMode = props.getProperty('DEVELOPMENT_MODE');

  try {
    // 確実に開発モード（楽天市場店等）でテストを行うため一時的に true をセット
    props.setProperty('DEVELOPMENT_MODE', 'true');
    console.log('テスト対象店舗ID: ' + getShopId());

    // ★ ここを開発環境（テストデータが存在する日）の出荷確定日に変更してテストしてください
    const testDate = '2026-03-19';
    console.log('テスト対象 出荷確定日: ' + testDate);
    console.log('APIリクエスト実行中...（しばらくお待ちください）');

    // データ取得
    const orders = fetchOrdersByShipDate(testDate);

    console.log('取得総件数: ' + orders.length);

    if (orders.length > 0) {
      console.log('先頭1件目のデータ: ' + JSON.stringify(orders[0]));
    } else {
      console.log('指定した日付に該当するデータはありませんでした。testDate（' + testDate + '）をデータが存在する日に変更して再度お試しください。');
    }

  } catch (error) {
    console.error('APIテスト中にエラーが発生しました: ' + error.message);
  } finally {
    // 開発モード設定を元の状態に戻す
    if (originalMode !== null) {
      props.setProperty('DEVELOPMENT_MODE', originalMode);
    } else {
      props.deleteProperty('DEVELOPMENT_MODE');
    }
  }

  console.log('=== Phase 2 テスト完了 ===');
}
