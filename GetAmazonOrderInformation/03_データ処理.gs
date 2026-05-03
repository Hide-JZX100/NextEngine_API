/**
 * @file 03_データ処理.gs
 * @description ネクストエンジン 受注データ整形・加工モジュール
 *
 * 【概要】
 * APIから取得した生のJSONデータ（オブジェクトの配列）を、Googleスプレッドシートへの
 * 出力に適した形式（2次元配列）に変換・加工します。
 * 
 * 【主な機能】
 * - formatOrderData: 取得データを設定に基づき2次元配列化。
 *
 * 【データ加工ルール】
 * - 01_設定ファイル.gs の CONFIG_FIELDS で定義された順序に従い、列を構成します。
 * - 値が null または undefined の場合は、スプレッドシートでの表示を考慮し空文字（''）に置換します。
 * - すべてのデータは String() によって文字列化されます。これにより、電話番号や郵便番号が
 *   数値として扱われ、先頭の0が消える（0落ち）などのスプレッドシート側での意図しない型変換を防ぎます。
 *
 * 【依存関係】
 * - 01_設定ファイル.gs (CONFIG_FIELDS)
 */

/**
 * APIレスポンス（オブジェクトの配列）をシートへ書き込むための2次元配列に変換する
 * 
 * 【処理フロー】
 * 1. 引数の rawOrders（オブジェクト配列）をループ処理
 * 2. 各オブジェクトから、CONFIG_FIELDS の api の順に値を抽出
 * 3. 値が null または undefined の場合は空文字 '' に変換し、それ以外は文字列化する
 * 4. 2次元配列（行 × 列）として返す
 * 
 * 【使用タイミング】
 * - fetchOrdersByShipDate() でデータを取得した後、スプレッドシートへの出力前
 * 
 * @param {Array} rawOrders - fetchOrdersByShipDate() の戻り値
 * @returns {string[][]} 2次元配列（行 × 列）
 */
function formatOrderData(rawOrders) {
  if (!rawOrders || rawOrders.length === 0) {
    return [];
  }
  
  const formattedData = rawOrders.map(function(order) {
    return CONFIG_FIELDS.map(function(fieldConfig) {
      const val = order[fieldConfig.api];
      
      // 値が null または undefined の場合は空文字に変換
      if (val === null || val === undefined) {
        return '';
      }
      
      return String(val);
    });
  });
  
  return formattedData;
}
