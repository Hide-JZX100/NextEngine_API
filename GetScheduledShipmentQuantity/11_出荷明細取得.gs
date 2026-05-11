/**
 * @fileoverview ネクストエンジン出荷明細取得スクリプト（日付指定テスト版）
 * 
 * 【目的】ネクストエンジンAPIから出荷明細データを取得する。
 * 【現在の開発フェーズ】フェーズ3: 日付指定と複数件取得
 *   - ✅ フェーズ2: 1行取得（完了）
 *   - 🚀 フェーズ3-1,2: 日付指定と複数件取得（進行中）
 *   - 📅 フェーズ3-3: ページング処理（1000件ずつ）
 * 
 * 【スクリプトプロパティ設定】
 * SPREADSHEET_ID, SHEET_NAME, LOG_SHEET_NAME, OPERATION_LOG_SHEET_NAME を設定してください。
 * 
 * 【日付指定の仕様】
 * 出荷予定日 (receive_order_send_plan_date) で絞り込みを行います。
 * テスト用固定日付: 2025-10-03 ～ 2025-10-05
*/

/**
 * 💡 最小単位でのAPI接続・データ取得テスト（フェーズ2完了確認）
 * 
 * ネクストエンジンAPIから出荷明細データを1件だけ取得し、認証やエンドポイント、
 * 基本的なフィールドの取得が正しく行えるかを確認します。
 * 
 * 機能:
 * 1. スクリプトプロパティからアクセストークンとリフレッシュトークンを取得。
 * 2. ネクストエンジンの受注明細検索APIエンドポイントにアクセス。
 * 3. パラメータ limit: '1' を設定し、1件のみデータを取得。
 * 4. 14項目のフィールド（伝票番号、商品コード、寸法、出荷予定日など）を取得。
 * 5. APIレスポンスをコンソールに表示し、取得結果やデータ構造を検証。
 * 6. トークンが更新されていた場合、スクリプトプロパティに再保存。
 * 7. エラー発生時、権限不足（004001）などの詳細な解決方法を表示。
 * 
 * @returns {Object[]} 取得した出荷明細データの配列
 *
 * 【テストの目的】
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
      console.log('2. Shipping_piece.csv の全14項目が揃っているか確認してください');
      console.log('   ✓ 出荷予定日 (receive_order_send_plan_date)');
      console.log('   ✓ 伝票番号 (receive_order_row_receive_order_id)');
      console.log('   ✓ 商品コード (receive_order_row_goods_id)');
      console.log('   ✓ 商品名 (receive_order_row_goods_name)');
      console.log('   ✓ 受注数 (receive_order_row_quantity)');
      console.log('   ✓ 引当数 (receive_order_row_stock_allocation_quantity)');
      console.log('   ✓ 奥行き (goods_length)');
      console.log('   ✓ 幅 (goods_width)');
      console.log('   ✓ 高さ (goods_height)');
      console.log('   ✓ 発送方法コード (receive_order_delivery_id)');
      console.log('   ✓ 発送方法 (receive_order_delivery_name)');
      console.log('   ✓ 重さ (goods_weight)');
      console.log('   ✓ 受注状態区分 (receive_order_order_status_id)');
      console.log('   ✓ 送り先住所1 (receive_order_consignee_address1)');
      console.log('3. すべて揃っていれば、次のフェーズ（日付指定）に進みます');

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
 * 💡 APIフィールドの参考情報表示
 * 
 * ネクストエンジンの受注明細検索APIで取得可能な主なフィールド（項目）の一部を、
 * 開発者が確認しやすいように一覧表示します。
 * この関数自体はAPIコールを行いません。
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

/**
 * 💡 ページング処理による全データ取得（高性能ロジック / フェーズ3-ステップ3）
 * 
 * 指定期間内の全データを、ネクストエンジンAPIの制限である1000件ずつ繰り返し取得（ページング）することで、
 * 高速かつ確実に大量のデータを集約します。
 * 
 * 機能:
 * 1. offsetとlimit（1000件）を使用してwhileループを回し、全データを取得。
 * 2. 取得したデータを `allData` 配列に累積。
 * 3. トークンが更新された場合は即座にスクリプトプロパティに保存。
 * 4. APIコールの実行時間を計測し、負荷状況に応じて待機時間（Utilities.sleep）を調整。
 * 5. 最終的に総取得件数と総APIコール回数を表示し、処理効率を検証。
 * 
 * @param {string} [startDate='2025-10-03'] - 開始日（YYYY-MM-DD形式）
 * @param {string} [endDate='2025-10-05'] - 終了日（YYYY-MM-DD形式）
 * @returns {Object[]} 取得した全データの配列
 * 
 * 【テストの目的】
 * - 指定期間の全データを取得（最大3000件程度を想定）
 * - 1回のAPIコールで最大件数を取得し、ネットワーク負荷と処理時間を最適化する
 */
function fetchAllShippingData(startDate = '2025-10-03', endDate = '2025-10-05') {
  // ★★★ 全体の実行時間計測開始 ★★★
  const overallStartTime = new Date();

  try {
    console.log('=== 全データ取得開始 ===');
    console.log(`期間: ${startDate} ～ ${endDate}`);
    console.log('');

    // ネクストエンジンAPIのエンドポイント（スクリプトプロパティから取得、なければデフォルト）
    const NE_API_URL = PropertiesService.getScriptProperties().getProperty('NE_API_URL') || 'https://api.next-engine.org';

    // スクリプトプロパティからトークンを取得
    const properties = PropertiesService.getScriptProperties();
    let accessToken = properties.getProperty('ACCESS_TOKEN');
    let refreshToken = properties.getProperty('REFRESH_TOKEN');

    if (!accessToken || !refreshToken) {
      throw new Error('アクセストークンが見つかりません。先に認証.gsのtestApiConnection()を実行してください。');
    }

    // ★★★ 認証処理の時間計測（トークン取得は一瞬なので計測は参考程度）★★★
    const authEndTime = new Date();
    const authElapsedTime = (authEndTime - overallStartTime) / 1000;

    console.log(`トークン取得完了: ${authElapsedTime.toFixed(2)}秒`);
    console.log('');

    // 全データを格納する配列
    const allData = [];

    // ページング用の変数
    let offset = 0;
    const limit = 1000; // 1回のAPIコールで最大1000件取得
    let hasMoreData = true;
    let callCount = 0;

    // 受注明細検索APIのエンドポイント
    const url = `${NE_API_URL}/api_v1_receiveorder_row/search`;

    // 日付をネクストエンジンAPI用のフォーマットに変換
    const formattedStartDate = `${startDate} 00:00:00`;
    const formattedEndDate = `${endDate} 23:59:59`;

    // ★★★ API呼び出し開始時刻 ★★★
    const apiCallsStartTime = new Date();

    // ページングループ
    while (hasMoreData) {
      callCount++;

      console.log(`--- APIコール ${callCount}回目 ---`);
      console.log(`offset: ${offset}, limit: ${limit}`);

      // ★★★ 個別APIコールの時間計測開始 ★★★
      const apiCallStartTime = new Date();

      // APIリクエストのパラメータ
      const payload = {
        'access_token': accessToken,
        'refresh_token': refreshToken,
        'limit': limit.toString(),
        'offset': offset.toString(),
        // 出荷予定日で絞り込む
        'receive_order_send_plan_date-gte': formattedStartDate,
        'receive_order_send_plan_date-lte': formattedEndDate,
        // キャンセル行を除外
        'receive_order_row_cancel_flag-eq': '0',
        // Shipping_piece.csv の全項目を取得
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
          'receive_order_date',                        // 受注日
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

      // APIリクエスト実行
      const response = UrlFetchApp.fetch(url, options);
      const responseText = response.getContentText();
      const responseData = JSON.parse(responseText);

      // ★★★ 個別APIコールの時間計測終了 ★★★
      const apiCallEndTime = new Date();
      const apiCallElapsedTime = (apiCallEndTime - apiCallStartTime) / 1000;

      // トークンが更新された場合は保存して、次のループで使用
      if (responseData.access_token && responseData.refresh_token) {
        accessToken = responseData.access_token;
        refreshToken = responseData.refresh_token;
        properties.setProperties({
          'ACCESS_TOKEN': accessToken,
          'REFRESH_TOKEN': refreshToken,
          'TOKEN_UPDATED_AT': new Date().getTime().toString()
        });
      }

      // 結果の確認
      if (responseData.result === 'success') {
        const fetchedCount = parseInt(responseData.count);
        console.log(`取得件数: ${fetchedCount}件`);
        console.log(`APIコール時間: ${apiCallElapsedTime.toFixed(2)}秒`); // ★★★ 追加 ★★★

        // データを配列に追加
        if (responseData.data && responseData.data.length > 0) {
          allData.push(...responseData.data);
          console.log(`累積取得件数: ${allData.length}件`);
        }

        // 取得件数がlimit未満なら、これ以上データはない
        if (fetchedCount < limit) {
          hasMoreData = false;
          console.log('全データ取得完了');
        } else {
          // 次のページへ
          offset += limit;

          // APIコールが30秒を超えた場合は5秒待機、それ以外は2秒待機
          const waitTime = apiCallElapsedTime > 30 ? 5000 : 2000;
          console.log(`次ページ取得まで ${waitTime / 1000}秒待機`);
          Utilities.sleep(waitTime);

          console.log('');
        }

      } else {
        console.error('API エラー:', responseData.code, responseData.message);
        throw new Error(`API エラー: ${responseData.message}`);
      }
    }

    // ★★★ 全体の実行時間計測終了 ★★★
    const overallEndTime = new Date();
    const overallElapsedTime = (overallEndTime - overallStartTime) / 1000;
    const apiCallsElapsedTime = (overallEndTime - apiCallsStartTime) / 1000;

    console.log('');
    console.log('=== 全データ取得成功 ===');
    console.log(`総APIコール回数: ${callCount}回`);
    console.log(`総取得件数: ${allData.length}件`);
    console.log(`データ取得総時間: ${overallElapsedTime.toFixed(2)}秒`); // ★★★ 追加 ★★★
    console.log(`  - 認証: ${authElapsedTime.toFixed(2)}秒`); // ★★★ 追加 ★★★
    console.log(`  - API呼び出し: ${apiCallsElapsedTime.toFixed(2)}秒`); // ★★★ 追加 ★★★
    console.log('');

    // 最初の3件と最後の3件を表示
    /**if (allData.length > 0) {
      console.log('--- 最初の3件 ---');
      for (let i = 0; i < Math.min(3, allData.length); i++) {
        console.log(`${i + 1}件目:`);
        console.log(`  伝票番号: ${allData[i].receive_order_row_receive_order_id}`);
        console.log(`  商品コード: ${allData[i].receive_order_row_goods_id}`);
        console.log(`  出荷予定日: ${allData[i].receive_order_send_plan_date}`);
      }
      
      if (allData.length > 6) {
        console.log('');
        console.log('--- 最後の3件 ---');
        for (let i = Math.max(0, allData.length - 3); i < allData.length; i++) {
          console.log(`${i + 1}件目:`);
          console.log(`  伝票番号: ${allData[i].receive_order_row_receive_order_id}`);
          console.log(`  商品コード: ${allData[i].receive_order_row_goods_id}`);
          console.log(`  出荷予定日: ${allData[i].receive_order_send_plan_date}`);
        }
      }
    }*/

    console.log('');
    console.log('=== 次のステップ ===');
    console.log('1. 総取得件数が期待通りか確認してください');
    console.log('2. APIコール回数が適切か確認してください（3000件なら3〜4回）');
    console.log('3. 次はスプレッドシートへの書き込み機能を実装します');

    return allData;

  } catch (error) {
    console.error('全データ取得エラー:', error.message);
    console.error('エラー詳細:', error);
    throw error;
  }
}

/**
 * 💡 日付指定による複数件取得テスト（フェーズ3-ステップ1・2）
 * 
 * 指定した期間の出荷予定データを取得する機能を確認するための関数です。
 * ページング処理はまだ行わず、指定したlimit件数までを取得します。
 * 
 * 機能:
 * 1. 引数で指定された期間のデータを取得（デフォルト: 2025-10-03～10-05）。
 * 2. 出荷予定日フィルタ (`-gte`, `-lte`) を使用して絞り込み。
 * 3. 日付文字列をAPI規定形式 (`YYYY-MM-DD 00:00:00`) に変換。
 * 4. 取得したデータ（最初の3件）と総件数をコンソールに出力。
 * 
 * @param {string} [startDate='2025-10-03'] - 開始日（YYYY-MM-DD形式）
 * @param {string} [endDate='2025-10-05'] - 終了日（YYYY-MM-DD形式）
 * @param {number} [limit=10] - 取得件数制限
 * @returns {Object[]} 取得したデータの配列
 * 
 * 【テストの目的】
 * 1. 日付範囲でフィルタリングできるか確認
 * 2. 出荷予定日での絞り込みが正しく動作するか確認
*/
function fetchShippingDataByDate(startDate = '2025-10-03', endDate = '2025-10-05', limit = 10) {
  try {
    console.log('=== 日付指定出荷明細取得テスト開始 ===');
    console.log(`期間: ${startDate} ～ ${endDate}`);
    console.log(`取得件数制限: ${limit}件`);
    console.log('');

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
    const url = `${NE_API_URL}/api_v1_receiveorder_row/search`;

    // 日付をネクストエンジンAPI用のフォーマットに変換（YYYY-MM-DD → YYYY-MM-DD 00:00:00）
    const formattedStartDate = `${startDate} 00:00:00`;
    const formattedEndDate = `${endDate} 23:59:59`;

    // APIリクエストのパラメータ
    const payload = {
      'access_token': accessToken,
      'refresh_token': refreshToken,
      // 取得件数を指定
      'limit': limit.toString(),
      // 【重要】日付フィルタ
      // 出荷予定日で絞り込む（-gte: 以上、-lte: 以下）
      'receive_order_send_plan_date-gte': formattedStartDate,
      'receive_order_send_plan_date-lte': formattedEndDate,

      // Shipping_piece.csv の全項目を取得
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
        'receive_order_date',                        // 受注日
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
    console.log('日付フィルタ:');
    console.log(`  出荷予定日: ${formattedStartDate} ～ ${formattedEndDate}`);
    console.log('');

    // APIリクエスト実行
    const response = UrlFetchApp.fetch(url, options);
    const responseText = response.getContentText();

    console.log('レスポンスコード:', response.getResponseCode());

    // レスポンスをJSON形式にパース
    const responseData = JSON.parse(responseText);

    // 結果の確認
    if (responseData.result === 'success') {
      console.log('');
      console.log('=== 日付指定出荷明細取得成功 ===');
      console.log('取得件数:', responseData.count);
      console.log('');

      // 最初の3件だけログに表示（全件表示するとログが長くなるため）
      const displayCount = Math.min(3, responseData.data.length);
      console.log(`取得データ（最初の${displayCount}件を表示）:`);
      for (let i = 0; i < displayCount; i++) {
        console.log('');
        console.log(`--- ${i + 1}件目 ---`);
        console.log(JSON.stringify(responseData.data[i], null, 2));
      }

      if (responseData.data.length > 3) {
        console.log('');
        console.log(`... 他 ${responseData.data.length - 3} 件`);
      }

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
      console.log('1. 指定した日付範囲のデータが取得できているか確認してください');
      console.log('2. 取得件数が期待通りか確認してください');
      console.log('3. 次はlimitを増やして、より多くのデータ取得をテストします');

      return responseData.data;

    } else {
      console.log('');
      console.log('=== API エラー ===');
      console.log('エラーコード:', responseData.code);
      console.log('エラーメッセージ:', responseData.message);
      console.log('');
      console.log('レスポンス全体:');
      console.log(JSON.stringify(responseData, null, 2));

      throw new Error(`API エラー: ${responseData.message}`);
    }

  } catch (error) {
    console.error('日付指定出荷明細取得エラー:', error.message);
    console.error('エラー詳細:', error);
    throw error;
  }
}