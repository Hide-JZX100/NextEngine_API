// ============================================================
// 03_データ処理.gs
// ネクストエンジン Amazon受注データ取得 - データ整形処理
// ============================================================

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
