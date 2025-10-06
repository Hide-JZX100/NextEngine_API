/**
=============================================================================
データ変換テスト（ステップ4-1）
=============================================================================

* 【目的】
* APIレスポンスをスプレッドシート形式（14列の配列）に変換する
* 
* 【機能】
* - 1件のデータ変換テスト
* - 複数件のデータ変換テスト
* - null/undefinedの処理
* 

=============================================================================
1. コアデータ変換関数
convertApiDataToSheetRow(apiRowData)
💡 APIのオブジェクトをスプレッドシートの1行配列に変換
APIから取得した1件分のデータオブジェクトを、スプレッドシートの1行に対応する14列の配列に変換する核となるロジックです。
機能:
ネクストエンジンのフィールド名（例: receive_order_send_plan_date）を、スプレッドシートの列順に従って新しい配列にマッピングします。
|| '' や || 0 という記述により、APIデータに該当するフィールドが存在しない（nullやundefined）場合でも、エラーで停止することなく、空の文字列（''）または数値のゼロ（0）で値を補完します。
日付、伝票番号、文字列型のフィールドは || '' で空文字に。
数量や寸法、重さなど数値型のフィールドは || 0 でゼロに。
配列の要素数は必ず14列になるように構成されています。
エラー処理: 変換中に予期せぬエラーが発生した場合（例: apiRowDataが予期せぬ形式だった場合）、その行データの内容をログに出力し、エラーを再スローします。


2. データ変換テスト関数
testConvertOneRow()
💡 1件のデータでマッピングとNULL処理を検証
データ変換ロジックが期待通りに動作するかを、最小単位で確認するためのテスト関数です。
機能:
fetchAllShippingData関数（未定義、前のステップで実装された全データ取得関数）を呼び出し、最初の1件のAPIデータだけを取得します。
その1件のデータをconvertApiDataToSheetRowに渡し、変換を実行します。
変換後の配列の**長さ（14列であること）**と、各列の値が元のAPIデータに基づいて正しくマッピングされているかを確認できるように、詳細をコンソールに出力します。

testConvertAllRows()
💡 全件のデータ変換と処理の最終確認
取得した全データに対して変換処理を実行し、大量データでもエラーなく、件数を変えずに変換が完了するかを確認するためのテスト関数です。
機能:
fetchAllShippingDataで指定期間の全データ（例では123件を想定）を取得します。
全データをconvertApiDataToSheetRowでループ処理し、進捗状況（10件ごと）をコンソールに表示します。
変換後のデータ件数が取得件数と一致することを確認します。
変換後のデータの最初と最後の行を一部表示し、変換処理がデータ全体で均一かつ正確に行われたことを確認します。
重要な処理: このテストが成功すれば、大量のAPIデータをスプレッドシートに書き込める状態になり、次のステップであるスプレッドシート書き込みに進む準備が完了します。

=============================================================================
*/

/**
 * APIデータを1行分のスプレッドシートデータに変換
 * 
 * @param {Object} apiRowData - APIから取得した1行分のデータ
 * @return {Array} 14列の配列（スプレッドシートの1行分）
 */
function convertApiDataToSheetRow(apiRowData) {
  try {
    // Shipping_piece.csvの14列に対応する配列を作成
    return [
      // 列1: 出荷予定日（日付型）
      apiRowData.receive_order_send_plan_date || '',
      
      // 列2: 伝票番号（整数）
      apiRowData.receive_order_row_receive_order_id || '',
      
      // 列3: 商品コード（文字列）
      apiRowData.receive_order_row_goods_id || '',
      
      // 列4: 商品名（文字列）
      apiRowData.receive_order_row_goods_name || '',
      
      // 列5: 受注数（整数）
      apiRowData.receive_order_row_quantity || 0,
      
      // 列6: 引当数（整数）
      apiRowData.receive_order_row_stock_allocation_quantity || 0,
      
      // 列7: 奥行き(cm)（浮動小数点）
      apiRowData.goods_length || 0,
      
      // 列8: 幅(cm)（浮動小数点）
      apiRowData.goods_width || 0,
      
      // 列9: 高さ(cm)（浮動小数点）
      apiRowData.goods_height || 0,
      
      // 列10: 発送方法コード（整数）
      apiRowData.receive_order_delivery_id || '',
      
      // 列11: 発送方法（文字列）
      apiRowData.receive_order_delivery_name || '',
      
      // 列12: 重さ(g)（浮動小数点）
      apiRowData.goods_weight || 0,
      
      // 列13: 受注状態区分（整数）
      apiRowData.receive_order_order_status_id || '',
      
      // 列14: 送り先住所1（文字列）
      apiRowData.receive_order_consignee_address1 || ''
    ];
    
  } catch (error) {
    console.error('データ変換エラー:', error);
    console.error('エラーが発生した行データ:', JSON.stringify(apiRowData, null, 2));
    throw error;
  }
}

/**
 * テスト: 1件のデータ変換
 * 
 * このテストの目的:
 * 1. 変換ロジックが正しく動作するか確認
 * 2. 各列が期待通りの値になっているか確認
 * 3. null/undefinedが適切に処理されているか確認
 */
function testConvertOneRow() {
  try {
    console.log('=== 1件データ変換テスト開始 ===');
    console.log('');
    
    // fetchAllShippingData()で取得したデータを使用
    // まず1件だけ取得
    const apiData = fetchAllShippingData('2025-10-03', '2025-10-05');
    
    if (!apiData || apiData.length === 0) {
      throw new Error('テスト用データが取得できませんでした');
    }
    
    console.log('--- 元のAPIデータ（1件目） ---');
    console.log(JSON.stringify(apiData[0], null, 2));
    console.log('');
    
    // データ変換実行
    const convertedRow = convertApiDataToSheetRow(apiData[0]);
    
    console.log('--- 変換後のデータ ---');
    console.log('配列の長さ:', convertedRow.length, '（期待値: 14）');
    console.log('');
    
    // 各列の内容を確認
    const columnNames = [
      '出荷予定日',
      '伝票番号',
      '商品コード',
      '商品名',
      '受注数',
      '引当数',
      '奥行き(cm)',
      '幅(cm)',
      '高さ(cm)',
      '発送方法コード',
      '発送方法',
      '重さ(g)',
      '受注状態区分',
      '送り先住所1'
    ];
    
    for (let i = 0; i < convertedRow.length; i++) {
      console.log(`列${i + 1}: ${columnNames[i]} = ${convertedRow[i]}`);
    }
    
    console.log('');
    console.log('=== 変換テスト成功 ===');
    console.log('');
    console.log('=== 次のステップ ===');
    console.log('1. 各列の値が正しいか確認してください');
    console.log('2. 空データ（null/undefined）が適切に処理されているか確認してください');
    console.log('3. 問題なければ、全123件の変換テストに進みます');
    
    return convertedRow;
    
  } catch (error) {
    console.error('1件データ変換テストエラー:', error.message);
    console.error('エラー詳細:', error);
    throw error;
  }
}

/**
 * テスト: 全データの変換
 * 
 * このテストの目的:
 * 1. 全データが正しく変換できるか確認
 * 2. 変換エラーが発生しないか確認
 * 3. 変換後のデータ構造が正しいか確認
 */
function testConvertAllRows() {
  try {
    console.log('=== 全データ変換テスト開始 ===');
    console.log('');
    
    // 全データを取得
    const apiData = fetchAllShippingData('2025-10-03', '2025-10-05');
    
    console.log(`取得データ件数: ${apiData.length}件`);
    console.log('');
    console.log('変換処理開始...');
    
    // 全データを変換
    const convertedData = [];
    for (let i = 0; i < apiData.length; i++) {
      const row = convertApiDataToSheetRow(apiData[i]);
      convertedData.push(row);
      
      // 進捗表示（10件ごと）
      if ((i + 1) % 10 === 0 || i === apiData.length - 1) {
        console.log(`変換完了: ${i + 1}/${apiData.length}件`);
      }
    }
    
    console.log('');
    console.log('=== 全データ変換成功 ===');
    console.log(`変換済みデータ件数: ${convertedData.length}件`);
    console.log('');
    
    // 最初の3行と最後の3行を表示
    console.log('--- 最初の3行 ---');
    for (let i = 0; i < Math.min(3, convertedData.length); i++) {
      console.log(`${i + 1}行目:`);
      console.log(`  出荷予定日: ${convertedData[i][0]}`);
      console.log(`  伝票番号: ${convertedData[i][1]}`);
      console.log(`  商品コード: ${convertedData[i][2]}`);
      console.log(`  商品名: ${convertedData[i][3]}`);
    }
    
    if (convertedData.length > 6) {
      console.log('');
      console.log('--- 最後の3行 ---');
      for (let i = Math.max(0, convertedData.length - 3); i < convertedData.length; i++) {
        console.log(`${i + 1}行目:`);
        console.log(`  出荷予定日: ${convertedData[i][0]}`);
        console.log(`  伝票番号: ${convertedData[i][1]}`);
        console.log(`  商品コード: ${convertedData[i][2]}`);
        console.log(`  商品名: ${convertedData[i][3]}`);
      }
    }
    
    console.log('');
    console.log('=== 次のステップ ===');
    console.log('1. 変換済みデータ件数が期待通りか確認してください（123件のはず）');
    console.log('2. サンプル表示されたデータが正しいか確認してください');
    console.log('3. 問題なければ、スプレッドシート書き込みに進みます');
    
    return convertedData;
    
  } catch (error) {
    console.error('全データ変換テストエラー:', error.message);
    console.error('エラー詳細:', error);
    throw error;
  }
}