/**
=============================================================================
ネクストエンジン出荷明細取得スクリプト（単一API版テスト）
=============================================================================

* 【目的】
* ネクストエンジンAPIから出荷明細データを取得する
* 
* 【機能】
* フェーズ2: 1行取得（最小単位での動作確認）
* - testFetchShippingData(): API 1コールで1件だけ出荷明細を取得してログに表示
* 
* 【前提条件】
* - 認証.gs が同じプロジェクトに存在すること
* - testApiConnection() で認証が通っていること
* - ACCESS_TOKEN と REFRESH_TOKEN がスクリプトプロパティに保存されていること
*
* 【スクリプトプロパティの設定方法】
* 1. GASエディタで「プロジェクトの設定」を開く（歯車のアイコン）
* 2. 「スクリプトプロパティ」セクションまでスクロール
* 3. 「スクリプトプロパティの追加」をクリックし、以下のキーと値を設定
*    キー                     | 値
*    -------------------------|------------------------------------
*    SPREADSHEET_ID          | 出荷情報を更新したいスプレッドシートのID
*    SHEET_NAME              | 出荷情報を更新したいシート名
*    LOG_SHEET_NAME          | 実行時間を記載したいシート名
*
* 【使用方法】
* 1. testFetchShippingData() を実行
* 2. ログを確認して、出荷明細が1件取得できていることを確認
* 
* 【注意事項】
* - このスクリプトはテスト用です
* - まだスプレッドシートへの書き込みは行いません
* - API呼び出し回数を抑えるため、limit=1で1件のみ取得します
*
=============================================================================
*/

/**
 * テスト用: 出荷明細を1件だけ取得してログ表示
 * 
 * このテストの目的:
 * 1. 受注明細検索APIのエンドポイントが正しいか確認
 * 2. 必要なパラメータが正しく設定できているか確認
 * 3. レスポンスの構造を理解する
 * 4. どんなフィールドが取得できるか確認
 */
function testFetchShippingData() {
  try {
    console.log('=== 出荷明細取得テスト開始 ===');
    
    // ネクストエンジンAPIのエンドポイント（スクリプトプロパティから取得、なければデフォルト）
    const NE_API_URL = PropertiesService.getScriptProperties().getProperty('NE_API_URL') || 'https://api.next-engine.org';
    
    // スクリプトプロパティからトークンを取得
    const properties = PropertiesService.getScriptProperties();
    const accessToken = properties.getProperty('ACCESS_TOKEN');
    const refreshToken = properties.getProperty('REFRESH_TOKEN');
    
    if (!accessToken || !refreshToken) {
      throw new Error('アクセストークンが見つかりません。先に認証.gsのtestApiConnection()を実行してください。');
    }
    
    console.log('トークン取得完了');
    
    // 受注明細検索APIのエンドポイント
    // ネクストエンジンでは「受注明細」が出荷に関する情報を持っています
    const url = `${NE_API_URL}/api_v1_receiveorder_row/search`;
    
    // APIリクエストのパラメータ
    const payload = {
      'access_token': accessToken,
      'refresh_token': refreshToken,
      // テスト用に1件のみ取得
      'limit': '1',
      // Shipping_piece.csv の全項目を取得
      // 受注明細、商品マスタ、受注伝票の情報を1回のAPIコールで取得
      'fields': [
        // 受注明細の基本情報
        'receive_order_row_receive_order_id',        // 伝票番号
        'receive_order_row_goods_id',                // 商品コード
        'receive_order_row_goods_name',              // 商品名
        'receive_order_row_quantity',                // 受注数
        'receive_order_row_stock_allocation_quantity', // 引当数
        
        // 商品マスタの寸法・重量情報
        'goods_length',                              // 奥行き（cm）
        'goods_width',                               // 幅（cm）
        'goods_height',                              // 高さ（cm）
        'goods_weight',                              // 重さ（g）
        
        // 受注伝票の情報
        'receive_order_send_plan_date',              // 出荷予定日
        'receive_order_delivery_id',                 // 発送方法コード
        'receive_order_delivery_name',               // 発送方法
        'receive_order_order_status_id',             // 受注状態区分
        'receive_order_consignee_address1'           // 送り先住所1
      ].join(',')
    };
    
    // HTTPリクエストのオプション
    const options = {
      'method': 'POST',
      'headers': {
        'Content-Type': 'application/x-www-form-urlencoded'
      },
      'payload': Object.keys(payload).map(key => 
        encodeURIComponent(key) + '=' + encodeURIComponent(payload[key])
      ).join('&')
    };
    
    console.log('APIリクエスト送信中...');
    console.log('エンドポイント:', url);
    console.log('パラメータ:', payload);
    
    // APIリクエスト実行
    const response = UrlFetchApp.fetch(url, options);
    const responseText = response.getContentText();
    
    console.log('レスポンスコード:', response.getResponseCode());
    
    // レスポンスをJSON形式にパース
    const responseData = JSON.parse(responseText);
    
    console.log('APIレスポンス:');
    console.log(JSON.stringify(responseData, null, 2));
    
    // 結果の確認
    if (responseData.result === 'success') {
      console.log('');
      console.log('=== 出荷明細取得成功 ===');
      console.log('取得件数:', responseData.count);
      console.log('');
      console.log('取得データ:');
      console.log(JSON.stringify(responseData.data, null, 2));
      
      // トークンが更新された場合は保存
      if (responseData.access_token && responseData.refresh_token) {
        properties.setProperties({
          'ACCESS_TOKEN': responseData.access_token,
          'REFRESH_TOKEN': responseData.refresh_token,
          'TOKEN_UPDATED_AT': new Date().getTime().toString()
        });
        console.log('');
        console.log('トークンが更新されました');
      }
      
      console.log('');
      console.log('=== 次のステップ ===');
      console.log('1. 取得できたフィールド（項目）を確認してください');
      console.log('2. Shipping_piece.csvと比較して、必要な項目が揃っているか確認してください');
      console.log('3. 不足している項目があれば、fieldsパラメータに追加します');
      
      return responseData.data;
      
    } else {
      console.log('');
      console.log('=== API エラー ===');
      console.log('エラーコード:', responseData.code);
      console.log('エラーメッセージ:', responseData.message);
      
      // よくあるエラーの対処方法を表示
      if (responseData.code === '004001') {
        console.log('');
        console.log('【解決方法】');
        console.log('ネクストエンジンの「アプリを作る」→アプリ編集→「API」タブで');
        console.log('「受注明細検索」の権限を有効にしてください');
      }
      
      throw new Error(`API エラー: ${responseData.message}`);
    }
    
  } catch (error) {
    console.error('出荷明細取得エラー:', error.message);
    console.error('エラー詳細:', error);
    throw error;
  }
}

/**
 * 取得可能なフィールド一覧を表示（参考用）
 * 
 * ネクストエンジンの受注明細検索APIで取得できる主なフィールド:
 * - receive_order_row_date: 受注日
 * - receive_order_row_shop_cut_form_id: 伝票番号
 * - receive_order_row_goods_code: 商品コード
 * - receive_order_row_goods_name: 商品名
 * - receive_order_row_quantity: 数量
 * - receive_order_row_price: 単価
 * - receive_order_row_goods_width: 幅
 * - receive_order_row_goods_depth: 奥行
 * - receive_order_row_goods_height: 高さ
 * - receive_order_row_goods_weight: 重量
 * 
 * など、多数のフィールドが取得可能です
 */
function showAvailableFields() {
  console.log('=== 受注明細検索APIで取得可能な主なフィールド ===');
  console.log('');
  console.log('基本情報:');
  console.log('  receive_order_row_date: 受注日');
  console.log('  receive_order_row_shop_cut_form_id: 伝票番号');
  console.log('  receive_order_row_goods_code: 商品コード');
  console.log('  receive_order_row_goods_name: 商品名');
  console.log('  receive_order_row_quantity: 数量');
  console.log('  receive_order_row_price: 単価');
  console.log('');
  console.log('商品寸法・重量:');
  console.log('  receive_order_row_goods_width: 幅（cm）');
  console.log('  receive_order_row_goods_depth: 奥行（cm）');
  console.log('  receive_order_row_goods_height: 高さ（cm）');
  console.log('  receive_order_row_goods_weight: 重量（g）');
  console.log('');
  console.log('配送情報:');
  console.log('  receive_order_row_delivery_method_code: 配送方法コード');
  console.log('  receive_order_row_delivery_method_name: 配送方法名');
  console.log('');
  console.log('※ 詳細はネクストエンジンAPIドキュメントを参照してください');
}