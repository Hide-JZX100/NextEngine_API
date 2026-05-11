/**
 * @file 04_スプレッドシート書き込み.gs
 * @description ネクストエンジン 受注データ出力（スプレッドシート）モジュール
 *
 * 【概要】
 * 整形済みの受注データを、指定されたGoogleスプレッドシートのシートへ出力します。
 * 
 * 【主な機能】
 * - writeToSheet: ヘッダーの自動生成（初回のみ）およびデータの出力（追記）。
 * - 出力スタイルの制御: ヘッダー行への背景色、太字、中央揃え等の適用。
 *
 * 【出力仕様】
 * - シートが空（最終行が0）の場合、01_設定ファイル.gs の CONFIG_FIELDS に基づきヘッダー行を生成します。
 * - データは常に最終行の次から追記されます。
 * - 書き込み先はスクリプトプロパティ「OUTPUT_SPREADSHEET_ID」および「OUTPUT_SHEET_NAME」に依存します。
 *
 * 【依存関係】
 * - 01_設定ファイル.gs (各プロパティ取得関数, CONFIG_FIELDS)
 * - 03_データ処理.gs (formatOrderData) ※テスト用関数で利用
 * - Google Apps Script (SpreadsheetApp)
 */

/**
 * 整形済みの受注データをスプレッドシートの指定したシートへ書き込みます。
 *
 * 【詳細仕様】
 * 1. 接続チェック: スクリプトプロパティから取得したIDとシート名に基づき、対象シートへのアクセスを試みます。
 *    設定が不正な場合やシートが存在しない場合は、明確なエラーメッセージをスローして処理を中断します。
 * 2. ヘッダー自動生成: シートが完全に空（`getLastRow() === 0`）の状態である場合のみ、
 *    `CONFIG_FIELDS` で定義されたヘッダー名を1行目に書き込みます。
 *    その際、視認性を高めるために以下のスタイルを適用します：
 *    - 背景色: オレンジ (#e67e22)
 *    - 文字色: 白 (#ffffff)
 *    - フォント: 太字
 *    - 配置: 中央揃え
 * 3. 追記ロジック: データは常に現在の最終行（`getLastRow()`）の直後から書き込まれます。
 *    これにより、複数回の実行結果を一つのシートに蓄積することが可能です。
 * 4. 一括書き込み: GASの実行速度を最適化するため、`setValues()` メソッドを使用して
 *    2次元配列を一度に書き込み、API呼び出し回数を最小限に抑えています。
 *
 * 【制約事項】
 * - すべてのデータ行はヘッダーと同じ列数（`CONFIG_FIELDS` の要素数）であることを前提としています。
 *
 * 【使用タイミング】
 * - APIから取得し、`formatOrderData()` で整形されたデータの最終出力として実行します。
 *
 * @param {string[][]} data - 書き込み対象となる2次元配列データ（行 × 列）
 * @throws {Error} スプレッドシートIDやシート名が不正な場合にスローされます。
 */
function writeToSheet(data) {
  const spreadsheetId = getOutputSpreadsheetId();
  const sheetName = getOutputSheetName();
  
  if (!spreadsheetId) {
    throw new Error('出力先スプレッドシートIDが設定されていません。スクリプトプロパティ「OUTPUT_SPREADSHEET_ID」を確認してください。');
  }
  
  const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  const sheet = spreadsheet.getSheetByName(sheetName);
  
  if (!sheet) {
    throw new Error('指定されたシート名（' + sheetName + '）が見つかりません。シートタブを作成するか、スクリプトプロパティ「OUTPUT_SHEET_NAME」を確認してください。');
  }
  
  // 既存の最終行を取得
  const lastRow = sheet.getLastRow();
  
  // 1. ヘッダーの書き込み（シートが空の場合のみ）
  if (lastRow === 0) {
    const headers = CONFIG_FIELDS.map(function(f) { return f.header; });
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setValues([headers]);
    headerRange.setBackground('#e67e22');
    headerRange.setFontColor('#ffffff');
    headerRange.setFontWeight('bold');
    headerRange.setHorizontalAlignment('center');
  }
  
  // 2. データ行の書き込み（最終行の下から追記）
  if (data && data.length > 0) {
    // 現在の最終行を再度取得し、その次の行（最低でも2行目）を開始行とする
    const startRow = Math.max(sheet.getLastRow() + 1, 2);
    const dataRange = sheet.getRange(startRow, 1, data.length, data[0].length);
    dataRange.setValues(data);
  }
}

/**
 * 【開発・検証用】Phase 3：データ整形およびシート書き込みの動作確認テスト
 *
 * 実際のAPI通信を行わずに、固定のダミーデータを使用して「オブジェクト配列から2次元配列への変換」
 * および「スプレッドシートへの物理的な書き込み」が正しく行われるかを検証します。
 *
 * 【テスト項目】
 * - `null` や `undefined` を含むデータが空文字として適切に処理されるか。
 * - ヘッダーが未存在の場合に正しく生成され、スタイルが適用されるか。
 * - 既存データがある場合に、その下の行から正しく追記されるか。
 *
 * 【使用タイミング】
 * - 書き込み機能の実装直後や、出力レイアウト（`CONFIG_FIELDS`）を変更した際のデバッグ時に実行します。
 */
function testPhase3() {
  console.log('=== Phase 3 テスト開始 ===');
  
  // ダミーデータ
  // （値がnullやundefinedの場合に空文字になるかもテストする）
  const dummyRawOrders = [
    {
      receive_order_id: 'TEST-0001',
      receive_order_date: '2025-01-01 10:00:00',
      receive_order_send_plan_date: '2025-01-02',
      receive_order_purchaser_name: '山田 太郎',
      receive_order_purchaser_zip_code: '100-0001',
      receive_order_purchaser_address1: '東京都千代田区千代田1-1',
      receive_order_purchaser_address2: null, // nullテスト用
      receive_order_purchaser_tel: '03-1234-5678',
      receive_order_purchaser_mail_address: 'taro.yamada@example.com',
      receive_order_consignee_name: '山田 花子',
      receive_order_consignee_zip_code: '100-0001',
      receive_order_consignee_address1: '東京都千代田区千代田1-1',
      receive_order_consignee_address2: undefined, // undefinedテスト用
      receive_order_consignee_tel: '03-1234-5678'
    },
    {
      receive_order_id: 'TEST-0002',
      receive_order_date: '2025-01-01 11:30:00',
      receive_order_send_plan_date: '2025-01-03',
      receive_order_purchaser_name: '佐藤 次郎',
      receive_order_purchaser_zip_code: '530-0001',
      receive_order_purchaser_address1: '大阪府大阪市北区梅田1-1',
      receive_order_purchaser_address2: 'テストビルディング 2F',
      receive_order_purchaser_tel: '06-8765-4321',
      receive_order_purchaser_mail_address: 'jiro.sato@example.com',
      receive_order_consignee_name: '佐藤 次郎',
      receive_order_consignee_zip_code: '530-0001',
      receive_order_consignee_address1: '大阪府大阪市北区梅田1-1',
      receive_order_consignee_address2: 'テストビルディング 2F',
      receive_order_consignee_tel: '06-8765-4321'
    }
  ];
  
  try {
    console.log('1. ダミーデータを2次元配列に整形します...');
    const formattedData = formatOrderData(dummyRawOrders);
    console.log('　-> 整形完了。行数: ' + formattedData.length);
    
    console.log('2. スプレッドシートへの書き込みを実行します...');
    writeToSheet(formattedData);
    
    console.log('　-> 書き込み完了。指定したスプレッドシートのタブ（デフォルト: API）をご確認ください。');
  } catch (error) {
    console.error('Phase 3 テスト中にエラーが発生しました: ' + error.message);
  }
  
  console.log('=== Phase 3 テスト完了 ===');
}
