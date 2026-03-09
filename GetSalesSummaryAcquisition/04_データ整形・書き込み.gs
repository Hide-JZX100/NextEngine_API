/**
 * =============================================================================
 * ネクストエンジン受注明細取得スクリプト - Phase 4: データ整形・書き込み
 * =============================================================================
 * 
 * 【概要】
 * APIから取得した受注明細データをスプレッドシート形式に整形し、
 * Googleスプレッドシートに書き込む機能を提供します。
 * 
 * 【Phase 4 実装内容】
 * 1. APIレスポンスからスプレッドシート形式へのデータ変換
 * 2. 店舗名の付与(店舗コード→店舗名)
 * 3. 日付フォーマット変換(YYYY-MM-DD HH:MM:SS → YYYY/MM/DD)
 * 4. スプレッドシートへの書き込み処理
 * 5. ヘッダー保持機能(2行目からクリアして書き込み)
 * 6. テスト関数
 * 
 * 【スプレッドシートのフォーマット】
 * A列: 店舗名
 * B列: 伝票番号
 * C列: 受注日
 * D列: 商品コード(伝票)
 * E列: 商品名(伝票)
 * F列: 受注数
 * G列: 小計
 * H列: 発送代
 * I列: キャンセル区分
 * J列: 店舗コード
 * 
 * @version 1.0
 * @date 2025-11-24
 */

// =============================================================================
// スプレッドシート列定義
// =============================================================================

/**
 * スプレッドシートのヘッダー定義
 * 
 * サンプル.txtに基づいたヘッダー行です。
 */
const SPREADSHEET_HEADERS = [
  '店舗名',
  '伝票番号',
  '受注日',
  '商品コード(伝票)',
  '商品名(伝票)',
  '受注数',
  '小計',
  '発送代',
  'キャンセル区分',
  '店舗コード'
];

// =============================================================================
// データ整形
// =============================================================================

/**
 * 日付を YYYY/MM/DD 形式にフォーマット
 * 
 * ネクストエンジンAPIの日時形式(YYYY-MM-DD HH:MM:SS)を
 * スプレッドシート用の日付形式(YYYY/MM/DD)に変換します。
 * 
 * @param {string} dateTimeStr - 日時文字列(YYYY-MM-DD HH:MM:SS)
 * @return {string} 日付文字列(YYYY/MM/DD)
 */
function formatDateForSpreadsheet(dateTimeStr) {
  if (!dateTimeStr) {
    return '';
  }
  
  // 'YYYY-MM-DD HH:MM:SS' から 'YYYY-MM-DD' 部分を抽出
  const datePart = dateTimeStr.split(' ')[0];
  
  // '-' を '/' に置換
  return datePart.replace(/-/g, '/');
}

/**
 * APIレスポンスデータをスプレッドシート行データに変換
 * 
 * 1件の受注明細データをスプレッドシートの1行分の配列に変換します。
 * 
 * @param {Object} row - 受注明細データ(APIレスポンス)
 * @param {Map<string, string>} shopMap - 店舗名マップ
 * @return {Array} スプレッドシート行データ
 */
function convertToSpreadsheetRow(row, shopMap) {
  // 店舗コードから店舗名を取得
  const shopCode = row.receive_order_shop_id || '';
  const shopName = getShopName(shopMap, shopCode);
  
  // 受注日をフォーマット
  const orderDate = formatDateForSpreadsheet(row.receive_order_date);
  
  // キャンセル区分(0 or 1)
  const cancelFlag = row.receive_order_row_cancel_flag || '0';
  
  return [
    shopName,                                      // A列: 店舗名
    row.receive_order_row_receive_order_id || '', // B列: 伝票番号
    orderDate,                                     // C列: 受注日
    row.receive_order_row_goods_id || '',         // D列: 商品コード
    row.receive_order_row_goods_name || '',       // E列: 商品名
    row.receive_order_row_quantity || 0,          // F列: 受注数
    row.receive_order_row_sub_total_price || 0,   // G列: 小計
    row.receive_order_delivery_fee_amount || 0,   // H列: 発送代
    cancelFlag,                                    // I列: キャンセル区分
    shopCode                                       // J列: 店舗コード
  ];
}

/**
 * APIレスポンスデータ配列をスプレッドシート用2次元配列に変換
 * 
 * 全受注明細データをスプレッドシート形式に一括変換します。
 * 
 * @param {Array} data - 受注明細データ配列
 * @param {Map<string, string>} shopMap - 店舗名マップ
 * @return {Array<Array>} スプレッドシート用2次元配列
 */
function convertAllToSpreadsheetData(data, shopMap) {
  logMessage('=== データ整形処理開始 ===');
  logMessage(`変換前データ件数: ${data.length}件`);
  
  const startTime = new Date();
  
  const spreadsheetData = data.map(row => convertToSpreadsheetRow(row, shopMap));
  
  const endTime = new Date();
  const elapsedTime = (endTime - startTime) / 1000;
  
  logMessage(`変換後データ件数: ${spreadsheetData.length}件`);
  logMessage(`処理時間: ${elapsedTime.toFixed(2)}秒`);
  logMessage('=== データ整形処理完了 ===');
  
  return spreadsheetData;
}

// =============================================================================
// スプレッドシート書き込み
// =============================================================================

/**
 * 対象スプレッドシートを開く
 * 
 * スクリプトプロパティから対象スプレッドシートIDとシート名を取得し、
 * 対象のシートオブジェクトを返します。
 * 
 * @return {GoogleAppsScript.Spreadsheet.Sheet} 対象シート
 * @throws {Error} スプレッドシートまたはシートが見つからない場合
 */
function openTargetSheet() {
  const config = getScriptConfig();
  
  try {
    // スプレッドシートを開く
    const spreadsheet = SpreadsheetApp.openById(config.targetSpreadsheetId);
    
    // シートを取得
    const sheet = spreadsheet.getSheetByName(config.targetSheetName);
    
    if (!sheet) {
      throw new Error(
        `シート "${config.targetSheetName}" が見つかりません。\n` +
        `スプレッドシートID: ${config.targetSpreadsheetId}`
      );
    }
    
    logMessage(`対象シートを開きました: ${spreadsheet.getName()} / ${sheet.getName()}`);
    
    return sheet;
    
  } catch (error) {
    throw new Error(
      `対象スプレッドシートを開けませんでした: ${error.message}\n` +
      `TARGET_SPREADSHEET_ID と TARGET_SHEET_NAME を確認してください。`
    );
  }
}

/**
 * スプレッドシートのデータ範囲をクリア(ヘッダー行は保持)
 * 
 * 2行目以降のデータをクリアします。
 * 1行目のヘッダー行は保持されます。
 * 
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - 対象シート
 */
function clearSheetData(sheet) {
  const lastRow = sheet.getLastRow();
  
  if (lastRow <= 1) {
    logMessage('クリア対象データなし(ヘッダー行のみ)');
    return;
  }
  
  // 2行目から最終行までをクリア
  const dataRange = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn());
  dataRange.clear();
  
  logMessage(`データ範囲をクリアしました: 2行目～${lastRow}行目`);
}

/**
 * スプレッドシートにヘッダーを書き込む
 * 
 * 1行目にヘッダー行を書き込みます。
 * 既にヘッダーが存在する場合は上書きされます。
 * 
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - 対象シート
 */
function writeHeaders(sheet) {
  const headerRange = sheet.getRange(1, 1, 1, SPREADSHEET_HEADERS.length);
  headerRange.setValues([SPREADSHEET_HEADERS]);
  
  // ヘッダー行の書式設定(オプション)
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#f3f3f3');
  
  logMessage(`ヘッダーを書き込みました: ${SPREADSHEET_HEADERS.length}列`);
}

/**
 * スプレッドシートにデータを書き込む
 * 
 * 2行目からデータを書き込みます。
 * 既存データは事前にクリアされます(ヘッダー行は保持)。
 * 
 * @param {Array<Array>} data - スプレッドシート用2次元配列
 */
function writeToSpreadsheet(data) {
  logMessage('=== スプレッドシート書き込み処理開始 ===');
  
  const startTime = new Date();
  
  const sheet = openTargetSheet();
  
  // 既存データをクリア(ヘッダー行は保持)
  clearSheetData(sheet);
  
  // ヘッダーを書き込み
  writeHeaders(sheet);
  
  if (data.length === 0) {
    logMessage('⚠️ 書き込むデータがありません');
    logMessage('=== スプレッドシート書き込み処理完了 ===');
    return;
  }
  
  // データを書き込み(2行目から)
  const dataRange = sheet.getRange(2, 1, data.length, data[0].length);
  dataRange.setValues(data);
  
  const endTime = new Date();
  const elapsedTime = (endTime - startTime) / 1000;
  
  logMessage(`データを書き込みました: ${data.length}行 × ${data[0].length}列`);
  logMessage(`書き込み範囲: 2行目～${data.length + 1}行目`);
  logMessage(`処理時間: ${elapsedTime.toFixed(2)}秒`);
  logMessage('=== スプレッドシート書き込み処理完了 ===');
}

// =============================================================================
// テスト関数
// =============================================================================

/**
 * 日付フォーマットテスト
 * 
 * 日付変換が正しく動作するかテストします。
 */
function testDateFormat() {
  console.log('=== 日付フォーマットテスト ===');
  
  const testCases = [
    '2025-11-24 12:34:56',
    '2025-01-01 00:00:00',
    '2025-12-31 23:59:59',
    null,
    ''
  ];
  
  testCases.forEach(testCase => {
    const result = formatDateForSpreadsheet(testCase);
    console.log(`入力: "${testCase}" → 出力: "${result}"`);
  });
  
  console.log('');
  console.log('✅ 日付フォーマットテスト完了!');
}

/**
 * データ変換テスト
 * 
 * APIレスポンスからスプレッドシート形式への変換をテストします。
 */
function testDataConversion() {
  console.log('=== データ変換テスト ===');
  
  try {
    // 店舗マスタを読み込み
    const shopMap = getShopMapWithCache();
    
    // テストデータ(APIレスポンス形式)
    const testData = [
      {
        receive_order_row_receive_order_id: '1853465',
        receive_order_date: '2025-11-16 10:30:00',
        receive_order_row_goods_id: '0010-d-bb101p-en-gy',
        receive_order_row_goods_name: '◇Ｄ−ＥＮ101Ｐ4Ｖ／ＧＹ',
        receive_order_row_quantity: '1',
        receive_order_row_sub_total_price: '13979',
        receive_order_delivery_fee_amount: '0',
        receive_order_row_cancel_flag: '0',
        receive_order_shop_id: '2'
      },
      {
        receive_order_row_receive_order_id: '1853468',
        receive_order_date: '2025-11-16 14:20:00',
        receive_order_row_goods_id: '0013-en005tf-sxsd-zo-gy',
        receive_order_row_goods_name: '◇Ｓ＋ＳＤ−ＥＮ005ＴＦ／ＧＹ',
        receive_order_row_quantity: '1',
        receive_order_row_sub_total_price: '21998',
        receive_order_delivery_fee_amount: '0',
        receive_order_row_cancel_flag: '0',
        receive_order_shop_id: '3'
      }
    ];
    
    console.log(`テストデータ: ${testData.length}件`);
    console.log('');
    
    // 変換実行
    const converted = convertAllToSpreadsheetData(testData, shopMap);
    
    console.log('✅ 変換成功!');
    console.log('');
    
    console.log('【変換結果】');
    converted.forEach((row, index) => {
      console.log(`[${index + 1}]`, row);
    });
    console.log('');
    
    console.log('【フォーマット確認】');
    console.log('A列(店舗名):', converted[0][0]);
    console.log('C列(受注日):', converted[0][2], '← YYYY/MM/DD形式か確認');
    console.log('F列(受注数):', converted[0][5], '← 数値型か確認');
    console.log('');
    
    console.log('✅ データ変換テスト完了!');
    
    return converted;
    
  } catch (error) {
    console.error('❌ 変換エラー:', error.message);
    throw error;
  }
}

/**
 * スプレッドシート接続テスト
 * 
 * 対象スプレッドシートとシートが正しく開けるかテストします。
 */
function testTargetSheetConnection() {
  console.log('=== 対象シート接続テスト ===');
  
  try {
    const sheet = openTargetSheet();
    
    console.log('✅ シート接続成功!');
    console.log('');
    console.log('【シート情報】');
    console.log('- スプレッドシート名:', sheet.getParent().getName());
    console.log('- シート名:', sheet.getName());
    console.log('- 最終行:', sheet.getLastRow());
    console.log('- 最終列:', sheet.getLastColumn());
    console.log('');
    
    // 現在のヘッダー行を表示
    if (sheet.getLastRow() >= 1) {
      const headerRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());
      const headers = headerRange.getValues()[0];
      console.log('【現在のヘッダー行】');
      headers.forEach((header, index) => {
        console.log(`${String.fromCharCode(65 + index)}列: ${header}`);
      });
      console.log('');
    }
    
    console.log('✅ 対象シート接続テスト完了!');
    
  } catch (error) {
    console.error('❌ 接続エラー:', error.message);
    throw error;
  }
}

/**
 * スプレッドシート書き込みテスト
 * 
 * テストデータをスプレッドシートに書き込みます。
 * ⚠️ 実際にスプレッドシートが更新されます!
 */
function testWriteToSpreadsheet() {
  console.log('=== スプレッドシート書き込みテスト ===');
  console.log('⚠️ 実際にスプレッドシートが更新されます!');
  console.log('');
  
  try {
    // テストデータを作成
    const shopMap = getShopMapWithCache();
    
    const testData = [
      {
        receive_order_row_receive_order_id: '9999991',
        receive_order_date: '2025-11-24 10:00:00',
        receive_order_row_goods_id: 'TEST-001',
        receive_order_row_goods_name: 'テスト商品1',
        receive_order_row_quantity: '1',
        receive_order_row_sub_total_price: '1000',
        receive_order_delivery_fee_amount: '500',
        receive_order_row_cancel_flag: '0',
        receive_order_shop_id: '1'
      },
      {
        receive_order_row_receive_order_id: '9999992',
        receive_order_date: '2025-11-24 11:00:00',
        receive_order_row_goods_id: 'TEST-002',
        receive_order_row_goods_name: 'テスト商品2',
        receive_order_row_quantity: '2',
        receive_order_row_sub_total_price: '2000',
        receive_order_delivery_fee_amount: '500',
        receive_order_row_cancel_flag: '0',
        receive_order_shop_id: '2'
      }
    ];
    
    console.log(`テストデータ: ${testData.length}件`);
    console.log('');
    
    // データ変換
    const converted = convertAllToSpreadsheetData(testData, shopMap);
    
    // スプレッドシートに書き込み
    writeToSpreadsheet(converted);
    
    console.log('');
    console.log('✅ スプレッドシート書き込みテスト完了!');
    console.log('');
    console.log('【確認事項】');
    console.log('- スプレッドシートを開いて、データが正しく書き込まれているか確認してください');
    console.log('- ヘッダー行が保持されているか確認してください');
    console.log('- 2行目からテストデータが書き込まれているか確認してください');
    
  } catch (error) {
    console.error('❌ 書き込みエラー:', error.message);
    throw error;
  }
}

// =============================================================================
// Phase 4 統合テスト
// =============================================================================

/**
 * Phase 4 統合テスト
 * 
 * Phase 4で実装した全機能をテストします。
 * ⚠️ testWriteToSpreadsheet() は実際にスプレッドシートを更新します。
 */
function testPhase4() {
  console.log('╔════════════════════════════════════════════════════════════╗');
  console.log('║  Phase 4: データ整形・書き込み - 統合テスト              ║');
  console.log('╚════════════════════════════════════════════════════════════╝');
  console.log('');
  
  try {
    // 1. 日付フォーマットテスト
    console.log('【1】日付フォーマットテスト');
    testDateFormat();
    console.log('');
    
    // 2. データ変換テスト
    console.log('【2】データ変換テスト');
    testDataConversion();
    console.log('');
    
    // 3. スプレッドシート接続テスト
    console.log('【3】対象シート接続テスト');
    testTargetSheetConnection();
    console.log('');
    
    // 4. スプレッドシート書き込みテスト
    console.log('【4】スプレッドシート書き込みテスト');
    console.log('⚠️ 実際にスプレッドシートが更新されます!');
    console.log('実行しますか? (手動でtestWriteToSpreadsheet()を実行してください)');
    console.log('');
    
    console.log('╔════════════════════════════════════════════════════════════╗');
    console.log('║  ✅ Phase 4 統合テスト: すべて成功!                       ║');
    console.log('╚════════════════════════════════════════════════════════════╝');
    console.log('');
    console.log('【次のステップ】');
    console.log('1. testWriteToSpreadsheet() を実行して書き込みテスト');
    console.log('2. Phase 5: メイン処理とエラーハンドリングの開発に進みます');
    
  } catch (error) {
    console.error('');
    console.error('╔════════════════════════════════════════════════════════════╗');
    console.error('║  ❌ Phase 4 統合テスト: エラー発生                       ║');
    console.error('╚════════════════════════════════════════════════════════════╝');
    console.error('');
    console.error('エラー内容:', error.message);
    console.error('');
    console.error('【確認事項】');
    console.error('- TARGET_SPREADSHEET_ID が正しいか');
    console.error('- TARGET_SHEET_NAME が正しいか');
    console.error('- スプレッドシートへのアクセス権限があるか');
    
    throw error;
  }
}