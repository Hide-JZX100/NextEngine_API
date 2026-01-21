/**
 * 受注キャンセル区分情報取得スクリプト
 * 
 * 目的: ネクストエンジンAPIから受注キャンセル区分マスタを取得し、ログ出力する
 * エンドポイント: /api_v1_system_canceltype/info
 * 特徴: fieldsパラメータが不要
 */

function fetchAndLogCancelTypeMaster() {
    // /infoエンドポイントはfieldsパラメータが不要なためnullを渡す
    // 共通関数 testApiCall がよしなに処理してくれる
    testApiCall('api_v1_system_canceltype/info', null, '受注キャンセル区分');
}
