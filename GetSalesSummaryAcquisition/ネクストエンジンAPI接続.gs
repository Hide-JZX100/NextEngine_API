/**
 * =============================================================================
 * ネクストエンジン受注明細取得スクリプト - Phase 3: ネクストエンジンAPI接続
 * =============================================================================
 * 
 * 【概要】
 * ネクストエンジンAPIを呼び出し、受注明細データを取得する機能を提供します。
 * ページネーション、キャンセル除外、トークン自動更新に対応しています。
 * 
 * 【Phase 3 実装内容】
 * 1. ネクストエンジンAPI呼び出し基本関数
 * 2. 受注明細検索API呼び出し
 * 3. ページネーション処理(1000件超対応)
 * 4. キャンセル除外フィルタリング
 * 5. トークン自動更新処理
 * 6. テスト関数
 * 
 * 【取得するフィールド】
 * サンプル.txtに基づき、以下のフィールドを取得します:
 * - receive_order_row_receive_order_id: 伝票番号
 * - receive_order_date: 受注日
 * - receive_order_row_goods_id: 商品コード
 * - receive_order_row_goods_name: 商品名
 * - receive_order_row_quantity: 受注数
 * - receive_order_row_sub_total_price: 小計
 * - receive_order_delivery_fee_amount: 発送代
 * - receive_order_row_cancel_flag: キャンセルフラグ
 * - receive_order_shop_id: 店舗コード
 * 
 * @version 1.0
 * @date 2025-11-24
 */

// =============================================================================
// ネクストエンジンAPI呼び出し基本関数
// =============================================================================

/**
 * ネクストエンジンAPIを呼び出す
 * 
 * 共通的なAPI呼び出し処理を行います。
 * トークンの自動更新にも対応しています。
 * 
 * @param {string} endpoint - APIエンドポイント(例: '/api_v1_receiveorder_row/search')
 * @param {Object} params - リクエストパラメータ
 * @return {Object} APIレスポンス
 * @throws {Error} API呼び出し失敗時
 */
function callNextEngineApi(endpoint, params = {}) {
  const config = getScriptConfig();
  
  // アクセストークンとリフレッシュトークンを追加
  const requestParams = {
    access_token: config.accessToken,
    refresh_token: config.refreshToken,
    wait_flag: NE_WAIT_FLAG,
    ...params
  };
  
  // URLエンコードされたペイロードを作成
  const payload = Object.keys(requestParams)
    .map(key => `${encodeURIComponent(key)}=${encodeURIComponent(requestParams[key])}`)
    .join('&');
  
  const url = `${NE_API_BASE_URL}${endpoint}`;
  
  const options = {
    method: 'POST',
    headers: {
      'Content-Type': 'application/x-www-form-urlencoded'
    },
    payload: payload,
    muteHttpExceptions: true // HTTPエラーでも例外を投げない
  };
  
  logMessage(`API呼び出し: ${endpoint}`, LOG_LEVEL.SAMPLE);
  
  try {
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();
    
    logMessage(`レスポンスコード: ${responseCode}`, LOG_LEVEL.SAMPLE);
    
    // レスポンスをパース
    let responseData;
    try {
      responseData = JSON.parse(responseText);
    } catch (parseError) {
      throw new Error(`JSONパースエラー: ${responseText.substring(0, 200)}`);
    }
    
    // エラーチェック
    if (responseData.result !== 'success') {
      throw new Error(
        `API呼び出し失敗: ${responseData.result}\n` +
        `メッセージ: ${JSON.stringify(responseData)}`
      );
    }
    
    // トークンが更新された場合はスクリプトプロパティに保存
    if (responseData.access_token && responseData.refresh_token) {
      const props = PropertiesService.getScriptProperties();
      props.setProperties({
        'ACCESS_TOKEN': responseData.access_token,
        'REFRESH_TOKEN': responseData.refresh_token,
        'TOKEN_UPDATED_AT': new Date().getTime().toString()
      });
      
      logMessage('✅ トークンが更新されました', LOG_LEVEL.SAMPLE);
    }
    
    return responseData;
    
  } catch (error) {
    throw new Error(`API呼び出しエラー: ${error.message}`);
  }
}

// =============================================================================
// 受注明細検索API
// =============================================================================

/**
 * 受注明細検索APIのフィールド定義
 * 
 * 取得するフィールドをカンマ区切りで定義します。
 */
const RECEIVEORDER_ROW_FIELDS = [
  'receive_order_row_receive_order_id',    // 伝票番号
  'receive_order_date',                     // 受注日 ※受注伝票のフィールド
  'receive_order_row_goods_id',             // 商品コード
  'receive_order_row_goods_name',           // 商品名
  'receive_order_row_quantity',             // 受注数
  'receive_order_row_sub_total_price',      // 小計
  'receive_order_delivery_fee_amount',      // 発送代 ※受注伝票のフィールド
  'receive_order_row_cancel_flag',          // キャンセルフラグ
  'receive_order_shop_id'                   // 店舗コード ※受注伝票のフィールド
].join(',');

/**
 * 受注明細を検索(1ページ分)
 * 
 * 指定された条件で受注明細を検索し、1ページ分(最大1000件)を取得します。
 * 
 * @param {string} startDate - 検索開始日時(YYYY-MM-DD HH:MM:SS)
 * @param {string} endDate - 検索終了日時(YYYY-MM-DD HH:MM:SS)
 * @param {number} offset - オフセット(開始位置)
 * @param {number} limit - 取得件数(最大1000)
 * @return {Object} {data: Array, count: number, hasMore: boolean}
 */
function searchReceiveOrderRowPage(startDate, endDate, offset = 0, limit = NE_API_LIMIT) {
  logMessage(`受注明細検索: offset=${offset}, limit=${limit}`, LOG_LEVEL.SAMPLE);
  
  // デバッグ: 日付パラメータを確認
  logMessage(`検索期間: ${startDate} ～ ${endDate}`, LOG_LEVEL.SAMPLE);
  
  // 検索パラメータ
  // ※受注明細検索APIで受注伝票のフィールドを検索条件に使う場合、
  //   パラメータ名は「receiveorder_」プレフィックスなしで指定します
  const params = {
    fields: RECEIVEORDER_ROW_FIELDS,
    'receive_order_date-gte': startDate,  // 受注日 >= 開始日時
    'receive_order_date-lte': endDate,    // 受注日 <= 終了日時
    offset: offset,
    limit: limit
  };
  
  // API呼び出し
  const response = callNextEngineApi(NE_RECEIVEORDER_ROW_SEARCH_ENDPOINT, params);
  
  // レスポンスから必要な情報を抽出
  const data = response.data || [];
  const count = response.count || 0;
  const hasMore = (offset + data.length) < count;
  
  logMessage(
    `取得結果: ${data.length}件 (全体: ${count}件, 残り: ${hasMore ? 'あり' : 'なし'})`,
    LOG_LEVEL.SAMPLE
  );
  
  return {
    data,
    count,
    hasMore
  };
}

/**
 * 受注明細を全件取得(ページネーション対応)
 * 
 * 指定された期間の受注明細を全件取得します。
 * 1000件を超える場合は自動的にページネーション処理を行います。
 * 
 * @param {string} startDate - 検索開始日時(YYYY-MM-DD HH:MM:SS)
 * @param {string} endDate - 検索終了日時(YYYY-MM-DD HH:MM:SS)
 * @return {Array} 受注明細データの配列
 */
function searchReceiveOrderRowAll(startDate, endDate) {
  logMessage('=== 受注明細全件取得開始 ===');
  logMessage(`期間: ${startDate} ～ ${endDate}`);
  
  const allData = [];
  let offset = 0;
  let pageCount = 0;
  let totalCount = 0;
  
  const startTime = new Date();
  
  // ページネーション処理
  while (true) {
    pageCount++;
    
    logMessage(`--- ページ ${pageCount} 取得中 (offset: ${offset}) ---`, LOG_LEVEL.SAMPLE);
    
    const result = searchReceiveOrderRowPage(startDate, endDate, offset, NE_API_LIMIT);
    
    // 初回でtotalCountを保存
    if (totalCount === 0) {
      totalCount = result.count;
      logMessage(`全体件数: ${totalCount}件`);
    }
    
    // データを追加
    allData.push(...result.data);
    
    logMessage(`累計: ${allData.length}件 / ${totalCount}件`, LOG_LEVEL.SAMPLE);
    
    // 次のページがない場合は終了
    if (!result.hasMore) {
      break;
    }
    
    // オフセットを更新
    offset += NE_API_LIMIT;
    
    // API負荷軽減のため少し待機(オプション)
    Utilities.sleep(500); // 0.5秒待機
  }
  
  const endTime = new Date();
  const elapsedTime = (endTime - startTime) / 1000;
  
  logMessage('');
  logMessage('=== 受注明細全件取得完了 ===');
  logMessage(`取得件数: ${allData.length}件`);
  logMessage(`ページ数: ${pageCount}ページ`);
  logMessage(`処理時間: ${elapsedTime.toFixed(2)}秒`);
  
  return allData;
}

// =============================================================================
// キャンセル除外フィルタリング
// =============================================================================

/**
 * キャンセル明細を除外
 * 
 * receive_order_row_cancel_flag = "1" の明細行を除外します。
 * 
 * @param {Array} data - 受注明細データ配列
 * @return {Array} キャンセル除外後のデータ配列
 */
function filterCancelledRows(data) {
  logMessage('=== キャンセル明細除外処理 ===');
  logMessage(`除外前: ${data.length}件`);
  
  const filtered = data.filter(row => {
    // cancel_flag が "1" の場合は除外(false を返す)
    return row.receive_order_row_cancel_flag !== '1';
  });
  
  const cancelledCount = data.length - filtered.length;
  
  logMessage(`除外後: ${filtered.length}件`);
  logMessage(`除外数: ${cancelledCount}件`);
  
  return filtered;
}

// =============================================================================
// テスト関数
// =============================================================================

/**
 * ネクストエンジンAPI接続テスト
 * 
 * ネクストエンジンAPIに接続できるかテストします。
 * 認証ライブラリのtestApiConnection()を利用します。
 * 
 * ※関数名を testNEApiConnection() に変更(認証ライブラリとの重複回避)
 */
function testNEApiConnection() {
  console.log('=== ネクストエンジンAPI接続テスト ===');
  
  try {
    const props = PropertiesService.getScriptProperties();
    const result = NEAuth.testApiConnection(props);
    
    console.log('✅ API接続成功!');
    console.log('ユーザー情報:', result);
    console.log('');
    
    return result;
    
  } catch (error) {
    console.error('❌ API接続エラー:', error.message);
    console.error('');
    console.error('【確認事項】');
    console.error('- ACCESS_TOKEN が有効か');
    console.error('- REFRESH_TOKEN が有効か');
    console.error('- 認証ライブラリ(NEAuth)が追加されているか');
    throw error;
  }
}

/**
 * 受注明細検索テスト(1ページのみ)
 * 
 * 受注明細検索APIを呼び出し、1ページ分のデータを取得します。
 */
function testSearchReceiveOrderRowPage() {
  console.log('=== 受注明細検索テスト(1ページ) ===');
  
  try {
    const dateRange = getSearchDateRange();
    
    console.log('検索期間:');
    console.log(`  開始: ${dateRange.startDateStr}`);
    console.log(`  終了: ${dateRange.endDateStr}`);
    console.log('');
    
    const result = searchReceiveOrderRowPage(
      dateRange.startDateStr,
      dateRange.endDateStr,
      0,    // offset
      10    // limit(テストなので10件のみ)
    );
    
    console.log('✅ 検索成功!');
    console.log(`取得件数: ${result.data.length}件`);
    console.log(`全体件数: ${result.count}件`);
    console.log(`次ページあり: ${result.hasMore ? 'はい' : 'いいえ'}`);
    console.log('');
    
    if (result.data.length > 0) {
      console.log('【先頭3件のデータ】');
      logData(result.data.slice(0, 3), '受注明細サンプル');
    }
    
    return result;
    
  } catch (error) {
    console.error('❌ 検索エラー:', error.message);
    throw error;
  }
}

/**
 * 受注明細全件取得テスト
 * 
 * ページネーション処理を含めた全件取得をテストします。
 * ※実際のデータ量によっては時間がかかります
 */
function testSearchReceiveOrderRowAll() {
  console.log('=== 受注明細全件取得テスト ===');
  console.log('⚠️ 実データを取得します。データ量によっては時間がかかります。');
  console.log('');
  
  try {
    const dateRange = getSearchDateRange();
    
    console.log('検索期間:');
    console.log(`  開始: ${dateRange.startDateStr}`);
    console.log(`  終了: ${dateRange.endDateStr}`);
    console.log('');
    
    const data = searchReceiveOrderRowAll(
      dateRange.startDateStr,
      dateRange.endDateStr
    );
    
    console.log('✅ 全件取得成功!');
    console.log(`取得件数: ${data.length}件`);
    console.log('');
    
    if (data.length > 0) {
      console.log('【データサンプル(先頭3件)】');
      logData(data.slice(0, 3), '受注明細');
      
      console.log('【データサンプル(末尾3件)】');
      logData(data.slice(-3), '受注明細');
    }
    
    return data;
    
  } catch (error) {
    console.error('❌ 全件取得エラー:', error.message);
    throw error;
  }
}

/**
 * キャンセル除外テスト
 * 
 * キャンセルフラグによるフィルタリングをテストします。
 */
function testFilterCancelledRows() {
  console.log('=== キャンセル除外テスト ===');
  
  try {
    const dateRange = getSearchDateRange();
    
    // データ取得
    console.log('データ取得中...');
    const allData = searchReceiveOrderRowAll(
      dateRange.startDateStr,
      dateRange.endDateStr
    );
    
    console.log('');
    
    // キャンセル除外
    const filtered = filterCancelledRows(allData);
    
    console.log('');
    console.log('✅ キャンセル除外テスト完了!');
    console.log('');
    
    // キャンセルフラグの内訳を表示
    const cancelCounts = {};
    allData.forEach(row => {
      const flag = row.receive_order_row_cancel_flag || '0';
      cancelCounts[flag] = (cancelCounts[flag] || 0) + 1;
    });
    
    console.log('【キャンセルフラグ内訳】');
    Object.keys(cancelCounts).forEach(flag => {
      const label = flag === '1' ? 'キャンセル' : '有効';
      console.log(`  ${label}(${flag}): ${cancelCounts[flag]}件`);
    });
    
    return filtered;
    
  } catch (error) {
    console.error('❌ キャンセル除外エラー:', error.message);
    throw error;
  }
}

// =============================================================================
// Phase 3 統合テスト
// =============================================================================

/**
 * Phase 3 統合テスト
 * 
 * Phase 3で実装した全機能をテストします。
 * ⚠️ 実際のAPIを呼び出すため、実行には注意してください。
 */
function testPhase3() {
  console.log('╔════════════════════════════════════════════════════════════╗');
  console.log('║  Phase 3: ネクストエンジンAPI接続 - 統合テスト           ║');
  console.log('╚════════════════════════════════════════════════════════════╝');
  console.log('');
  console.log('⚠️ 実際のAPIを呼び出します。本番環境での実行には注意してください。');
  console.log('');
  
  try {
    // 1. API接続テスト
    console.log('【1】ネクストエンジンAPI接続テスト');
    testNEApiConnection();
    console.log('');
    
    // 2. 受注明細検索テスト(1ページ)
    console.log('【2】受注明細検索テスト(1ページ)');
    testSearchReceiveOrderRowPage();
    console.log('');
    
    // 3. キャンセル除外テスト
    console.log('【3】キャンセル除外テスト');
    testFilterCancelledRows();
    console.log('');
    
    console.log('╔════════════════════════════════════════════════════════════╗');
    console.log('║  ✅ Phase 3 統合テスト: すべて成功!                       ║');
    console.log('╚════════════════════════════════════════════════════════════╝');
    console.log('');
    console.log('【次のステップ】');
    console.log('Phase 4: データ整形・書き込みの開発に進みます。');
    console.log('');
    console.log('【オプションテスト】');
    console.log('全件取得テストを実行する場合:');
    console.log('  testSearchReceiveOrderRowAll()');
    
  } catch (error) {
    console.error('');
    console.error('╔════════════════════════════════════════════════════════════╗');
    console.error('║  ❌ Phase 3 統合テスト: エラー発生                       ║');
    console.error('╚════════════════════════════════════════════════════════════╝');
    console.error('');
    console.error('エラー内容:', error.message);
    console.error('');
    console.error('【確認事項】');
    console.error('- ACCESS_TOKEN / REFRESH_TOKEN が有効か');
    console.error('- ネクストエンジンAPIが正常に動作しているか');
    console.error('- 認証ライブラリ(NEAuth)がバージョン5で追加されているか');
    console.error('- ネットワーク接続が正常か');
    
    throw error;
  }
}

function debugDateRange() {
  console.log('=== 日付範囲デバッグ ===');
  
  const dateRange = getSearchDateRange();
  
  console.log('startDate:', dateRange.startDate);
  console.log('endDate:', dateRange.endDate);
  console.log('startDateStr:', dateRange.startDateStr);
  console.log('endDateStr:', dateRange.endDateStr);
  console.log('');
  
  // 型チェック
  console.log('startDateStr type:', typeof dateRange.startDateStr);
  console.log('endDateStr type:', typeof dateRange.endDateStr);
}