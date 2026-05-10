/**
 * @file 04_データ整形・書き込み.gs
 * @description ネクストエンジン受注明細取得スクリプト - Phase 4: データ整形・書き込み。
 * APIから取得した受注明細データをスプレッドシート形式に整形し、Googleスプレッドシートに書き込む機能を提供します。
 * 
 * ### 実装内容
 * 1. APIレスポンスからスプレッドシート形式へのデータ変換
 * 2. 店舗名の付与(店舗コード→店舗名)
 * 3. 日付フォーマット変換(YYYY-MM-DD HH:MM:SS → YYYY/MM/DD)
 * 4. スプレッドシートへの書き込み処理
 * 5. ヘッダー保持機能(2行目からクリアして書き込み)
 * 
 * ### スプレッドシートのフォーマット
 * - A: 店舗名, B: 伝票番号, C: 受注日, D: 商品コード, E: 商品名
 * - F: 受注数, G: 小計, H: 発送代, I: キャンセル区分, J: 店舗コード
 * 
 * ### 依存関係
 * - **01_基盤構築.gs**: logMessage(), LOG_LEVEL
 * - **02_店舗マスタ連携.gs**: getShopName(), getShopMapWithCache()
 *
 * ### 推奨テスト実行順序
 * 1. testDateFormat() -> testDataConversion() -> testTargetSheetConnection()
 * 2. testWriteToSpreadsheet() (実際に更新されます)
 * 3. testPhase4() (統合テスト)
 * 
 * ### 注意事項
 * - **データクリア**: `writeToSpreadsheet()` は実行時に2行目以降を全削除します。
 * - **冪等性**: 何度実行しても（同じ期間であれば）同じ結果になります。
 * - **型変換**: APIからの数値文字列はスプレッドシート書き込み時に自動変換されますが、不具合がある場合は明示的な `Number()` 変換を検討してください。
 * - **ヘッダー**: 1行目は実行のたびに `SPREADSHEET_HEADERS` 定数で上書きされます。
 * 
 * @version 1.0
 * @date 2025-11-24
 * @see testPhase4 - 統合テスト
 * @see writeToSpreadsheet - メインの書き込み処理
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
 * @details
 * ネクストエンジンAPIから返される日時文字列（例: "2025-11-24 12:34:56"）を、
 * スプレッドシートで見やすい日付形式（例: "2025/11/24"）に変換します。
 * 
 * 内部処理としては、まず空白で分割して「日付部分」のみを取り出し、
 * 次にハイフン（-）をスラッシュ（/）に置換しています。
 * これにより、スプレッドシート側で日付シリアル値として認識されやすくなります。
 * 
 * @example
 * formatDateForSpreadsheet("2025-11-24 12:34:56") // -> "2025/11/24"
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
 * @details
 * この関数は、本システムの中で最も重要な「構造変換」を担います。
 * 
 * **【N88-BASICの概念との比較】**
 * APIから返ってくるデータ（Object）は、いわば「ラベル付きの変数群」です。
 * 例： `row.receive_order_row_goods_id` という名前を指定してデータを取り出します。
 * 
 * 一方、スプレッドシートへの書き込み（Array）は「メモリ上の連続したアドレス」のようなものです。
 * `[0]`番目がA列、`[1]`番目がB列...というように、「場所（インデックス）」でデータを管理します。
 * 
 * この関数では、バラバラの名前が付いたデータを、スプレッドシートの左（A列）から右（J列）へ
 * 並べたい順番通りに1つの「行（配列）」としてパッキングします。
 * 
 * 1. 店舗コードを元に、マスタから「店舗名」を引きく（VLOOKUPのような処理）
 * 2. APIの日付を整形する
 * 3. 各項目を `[A, B, C, ...]` の順序で配列に格納する
 * 
 * この「名前ベース」から「位置ベース」への変換を行うことで、
 * Google Sheets API（setValues）が受け入れ可能な形式になります。
 * 
 * @param {Object} row - 受注明細データ（APIから返された1レコード分のオブジェクト）
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
 * @details
 * APIから取得した「レコードの集合（配列）」全体に対して、
 * `convertToSpreadsheetRow` を繰り返し適用し、スプレッドシート用の「表」を作成します。
 * 
 * JavaScriptの `.map()` メソッドを使用しています。これはBASICで言うところの
 * `FOR I=1 TO N ... NEXT` ループに相当し、各レコードを1行ずつスプレッドシート形式へ
 * 変換して新しいリストを作成する処理を高速に行います。
 * 
 * 結果として `[[A1, B1, C1...], [A2, B2, C2...], ...]` という
 * 「行の配列の配列（2次元配列）」が返されます。
 * 
 * @param {Array<Object>} data - APIから取得した受注明細データの配列
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
 * @details
 * `getScriptConfig` で読み込んだ設定情報（TARGET_SPREADSHEET_ID など）を使用し、
 * 書き込み先となるスプレッドシートと、その中の特定のシートを取得します。
 * 
 * 実行ユーザーがそのスプレッドシートに対して編集権限を持っていない場合や、
 * IDが間違っている場合は、ここでエラー（例外）が発生します。
 * 
 * @return {GoogleAppsScript.Spreadsheet.Sheet} 書き込み対象のシートオブジェクト
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
 * @details
 * シート内の既存データを削除します。
 * 重要なのは `1行目（ヘッダー）を残す` ことです。
 * `getRange(2, 1, lastRow - 1, ...)` という指定により、
 * 「2行目の1列目から、データが存在する最後の行まで」の範囲を選択してクリアします。
 * 
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - クリア操作を行うシートオブジェクト
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
 * @details
 * 定数 `SPREADSHEET_HEADERS` で定義された列名をシートの1行目に書き込みます。
 * `setValues` は2次元配列を要求するため、`[SPREADSHEET_HEADERS]` のように
 * 配列をさらにもう一つの配列で包んで（1行分の表として）渡しています。
 * 
 * 書き込み後、視認性を高めるためにフォントを太字にし、背景色（薄いグレー）を設定します。
 * 
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - 書き込み先のシートオブジェクト
 * @see SPREADSHEET_HEADERS
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
 * @details
 * 整形済みの2次元配列を、スプレッドシートの2行目以降に一括で書き込みます。
 * 
 * 1. シートを開く
 * 2. 既存データをクリアする
 * 3. ヘッダーを再作成する
 * 4. データを流し込む
 * 
 * 1セルずつ書き込むのではなく、`setValues()` を使って一括で流し込むのがGASの高速化の定石です。
 * 
 * @param {Array<Array<string|number>>} data - スプレッドシート形式に変換済みの2次元配列
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
 * @details
 * `formatDateForSpreadsheet` 関数の単体テストです。
 * 正常な日時文字列、月の変わり目、空データなどのパターンを渡し、
 * すべて "YYYY/MM/DD" 形式、または空文字で返ってくるかをログで確認します。
 * スプレッドシート側で日付フィルターが正しく機能するかどうかの生命線となるテストです。
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
 * @details
 * APIから返ってきたと想定される「擬似データ」を使い、
 * スプレッドシート用の2次元配列に正しく変換されるかを検証します。
 * 特に「店舗コード」がマスタを参照して「店舗名」に置き換わっているか、
 * 列の順番が定義通り（A列=店舗名, B列=伝票番号...）になっているかを重点的に確認します。
 * @return {Array<Array>} 変換後の2次元配列
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
 * @details
 * 書き込み先の `TARGET_SPREADSHEET_ID` が正しいかを確認します。
 * 実際にシートを開き、現在の最終行数や既存のヘッダー内容を表示します。
 * これにより、意図しないシートを上書きしてしまうリスクを防ぎます。
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
 * @details
 * 少量のテスト用ダミーデータを生成し、実際にスプレッドシートへ書き込みを行います。
 * **※この関数を実行すると、対象シートの内容が書き換わります。**
 * 
 * 実行後、スプレッドシートを開いて「2行目から正しく並んでいるか」を目視確認するためのものです。
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
 * @details
 * 「データ整形・書き込み」フェーズの全機能をシーケンシャルに検証します。
 * データの変換ロジックからシートへの接続確認までを行います。
 * ※事故防止のため、実際の書き込みテスト（testWriteToSpreadsheet）は、
 * ログで案内を出した上でユーザーが手動で実行する流れにしています。
 * @throws {Error} 変換ロジックや接続に不備がある場合
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