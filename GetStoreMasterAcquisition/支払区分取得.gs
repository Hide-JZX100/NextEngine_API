/**
 * 支払区分情報取得テスト
 * 目的: 支払区分情報(api_v1_system_paymentmethod/info)のAPI動作確認
 * 共通処理: テスト共通関数.gs の testApiCall() を使用
 */
function fetchAndLogPaymentMethodMaster() {
    // /infoエンドポイントはfieldsパラメータが不要なためnullを渡す
    testApiCall('api_v1_system_paymentmethod/info', null, '支払区分情報');
}
