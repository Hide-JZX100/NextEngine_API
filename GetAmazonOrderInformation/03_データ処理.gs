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
 * APIレスポンス（オブジェクトの配列）を、Googleスプレッドシートへの書き込みに適した2次元配列に変換・整形します。
 *
 * 【詳細仕様】
 * 1. 空データ制御: 引数が空（null, undefined, 空配列）の場合は即座に空配列を返し、後続の書き込み処理でのエラーを防止します。
 * 2. フィールド抽出: `01_設定ファイル.gs` 内の `CONFIG_FIELDS` で定義された順序とAPIキーに基づき、
 *    オブジェクトから値を抽出します。これにより、APIのレスポンス順序に依存せず、定義した通りのシート列構成を維持します。
 * 3. 欠損値補完: APIから返却された値が `null` または `undefined` の場合、スプレッドシート上での
 *    見た目や計算への影響を考慮し、一律で空文字（''）へ変換します。
 * 4. 型の統一（文字列化）: すべての値を `String()` を用いて明示的に文字列へ変換します。
 *    これは、電話番号の市外局番や郵便番号の先頭にある「0」が、スプレッドシートの自動型判定によって
 *    数値として扱われ消えてしまう（いわゆる「0落ち」現象）を防ぐための重要な処理です。
 *
 * 【使用タイミング】
 * - APIから取得したデータをシートへ展開（`setValues`）する直前の変換処理として使用します。
 *
 * @param {Object[]} rawOrders - ネクストエンジンAPIから取得した、受注オブジェクトの配列（JSONデータ）
 * @returns {string[][]} スプレッドシートの `Range.setValues()` メソッドでそのまま使用可能な2次元配列（行 × 列）
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
