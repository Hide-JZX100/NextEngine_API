/**
 * 仕入先マスタ取得テスト
 * 目的: 仕入先マスタ(api_v1_master_supplier/search)のAPI動作確認
 * 共通処理: テスト共通関数.gs の testApiCall() を使用
 */
function fetchAndLogSupplierMaster() {
    // 取得する主要フィールド（動作確認用）
    // 本番用では全てのフィールドを取得しますが、テストでは主要なものだけでOKです
    var fields = 'supplier_id,supplier_name,supplier_kana,supplier_tel,supplier_mail_address';

    testApiCall('api_v1_master_supplier/search', fields, '仕入先マスタ');
}
