// ============================================================
// 04_スプレッドシート書き込み.gs
// ネクストエンジン Amazon受注データ取得 - スプレッドシート出力
// ============================================================

/**
 * スプレッドシートの指定タブへヘッダーとデータを書き込む
 * 
 * 【処理フロー】
 * 1. getOutputSpreadsheetId() と getOutputSheetName() でシートを取得
 * 2. シートの既存データをすべてクリア（clearContents()）
 * 3. 1行目にヘッダーを書き込み、指定のスタイル（オレンジ背景・白文字・太字・中央揃え）を適用
 * 4. 2行目以降にデータを一括で書き込む
 * 
 * 【使用タイミング】
 * - データの整形完了後、最終出力処理時
 * 
 * @param {string[][]} data - formatOrderData() の戻り値
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
  
  // 1. シートの既存データをすべてクリア
  sheet.clearContents();
  
  // 2. ヘッダー行のデータを生成
  const headers = CONFIG_FIELDS.map(function(f) { return f.header; });
  
  // 3. ヘッダーの書き込みとスタイル設定
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setBackground('#e67e22');
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');
  headerRange.setHorizontalAlignment('center');
  
  // 4. データ行の書き込み
  if (data && data.length > 0) {
    const dataRange = sheet.getRange(2, 1, data.length, data[0].length);
    dataRange.setValues(data);
  }
}

/**
 * 【Phase 3 テスト用】動作確認関数
 * ダミーのAPIレスポンスデータを用いて、整形処理とシート書き込み処理を確認する
 * 
 * 【処理フロー】
 * 1. ダミーの受注データ配列（オブジェクトの配列）を定義
 * 2. formatOrderData() で2次元配列に変換
 * 3. writeToSheet() で対象のスプレッドシートに出力
 * 4. コンソールに完了ログを出力
 * 
 * 【使用タイミング】
 * - Phase 3実装完了後の動作確認時（手動で実行）
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
