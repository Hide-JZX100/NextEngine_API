/**
 * @file 03_ネクストエンジンAPI接続.gs
 * @description ネクストエンジン受注明細取得スクリプト - Phase 3: ネクストエンジンAPI接続。
 * ネクストエンジンAPIを呼び出し、受注明細データを取得する機能を提供します。
 * ページネーション、キャンセル除外、トークン自動更新に対応しています。
 * 
 * ### 実装内容
 * 1. ネクストエンジンAPI呼び出し基本関数
 * 2. 受注明細検索API呼び出し
 * 3. ページネーション処理(1000件超対応)
 * 4. キャンセル除外フィルタリング
 * 5. トークン自動更新処理
 * 
 * ### 取得フィールド
 * - 伝票番号、受注日、商品コード、商品名、受注数、小計、発送代、キャンセルフラグ、店舗コード
 * 
 * ### 依存関係 (01_基盤構築.gs)
 * - getScriptConfig(), logMessage(), LOG_LEVEL
 * - NE_API_BASE_URL, NE_RECEIVEORDER_ROW_SEARCH_ENDPOINT, NE_API_LIMIT
 * - ※NEAuthライブラリ(v5)が追加されている必要があります。
 * 
 * ### 推奨テスト実行順序
 * 1. testNEApiConnection() -> testSearchReceiveOrderRowPage()
 * 2. testFilterCancelledRows() -> testPhase3() (統合テスト)
 * 
 * ### 注意事項
 * - ページネーション終了判定は「返却件数 < limit」で行っています。
 * - 集中時間帯(07:00〜22:00)の実行はエラー頻度が高まる可能性があります。
 * - トークンはAPIレスポンスに含まれる最新値で自動上書きされます。
 * 
 * @version 1.0
 * @date 2025-11-24
 * @see testPhase3 - 統合テスト
 * @see searchReceiveOrderRowAll - ページネーション対応全件取得
 */

// =============================================================================
// ネクストエンジンAPI呼び出し基本関数
// =============================================================================

/**
 * ネクストエンジンAPIを呼び出す
 * 
 * @details
 * ネクストエンジンAPIへのHTTP POSTリクエストをカプセル化した共通関数です。
 * 内部で以下の処理を自動的に行います：
 * 1. スクリプトプロパティからの最新トークンの取得とリクエストへの付与。
 * 2. オブジェクト形式のパラメータを `application/x-www-form-urlencoded` 形式に変換。
 * 3. `UrlFetchApp` を使用した同期通信（HTTPエラーを例外として投げず、レスポンス内容で判定）。
 * 4. APIレスポンスに含まれる新しい `access_token` / `refresh_token` の自動検知とスクリプトプロパティへの保存。
 * 
 * トークンの更新が発生した場合、以降のAPI呼び出しでは自動的に新しいトークンが使用されるため、
 * 呼び出し側はトークンの有効期限を意識する必要がありません。
 * 
 * @param {string} endpoint - APIエンドポイント（例: `/api_v1_receiveorder_row/search`）
 * @param {Object} [params={}] - APIに渡すリクエストパラメータ。トークン類は自動でマージされます。
 * @return {Object} パース済みのJSONレスポンスオブジェクト
 * @note API側でエラー（result != "success"）が返った場合は、メッセージを抽出して例外をスローします。
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
 * @details
 * 指定された日時範囲内の受注明細（receiveorder_row）を検索し、1ページ（最大1000件）を取得します。
 * 
 * 【検索条件の仕様】
 * - `receive_order_date`: 受注日時で絞り込みます。
 * - `receive_order_row_cancel_flag-eq: "0"`: APIレベルで有効なデータのみに絞り込みます。
 * - ページネーション用パラメータ（offset, limit）により取得範囲を制御します。
 * 
 * @param {string} startDate - 検索開始日時（`YYYY-MM-DD HH:mm:ss` 形式）
 * @param {string} endDate - 検索終了日時（`YYYY-MM-DD HH:mm:ss` 形式）
 * @param {number} [offset=0] - 取得開始位置（件数ベース）
 * @param {number} [limit=NE_API_LIMIT] - 取得件数。APIの仕様上、最大1000まで。
 * @return {Object} 検索結果オブジェクト：
 *   - data {Array<Object>}: 明細データの配列
 *   - count {number}: 検索条件に合致する全データの総件数
 *   - hasMore {boolean}: 次のページ（未取得分）が存在するかどうかのフラグ
 * @see RECEIVEORDER_ROW_FIELDS, callNextEngineApi
 */
function searchReceiveOrderRowPage(startDate, endDate, offset = 0, limit = NE_API_LIMIT) {
  logMessage(`受注明細検索: offset=${offset}, limit=${limit}`, LOG_LEVEL.SAMPLE);

  // デバッグ: 日付パラメータを確認
  logMessage(`検索期間: ${startDate} ～ ${endDate}`, LOG_LEVEL.SAMPLE);

  // 検索パラメータ
  // ※受注明細検索APIで受注伝票のフィールドを検索条件に使う場合、
  //   パラメータ名は「receiveorder_」プレフィックスなしで指定します
  // ※キャンセルフラグ = 0 (有効な明細のみ取得)
  const params = {
    fields: RECEIVEORDER_ROW_FIELDS,
    'receive_order_date-gte': startDate,              // 受注日 >= 開始日時
    'receive_order_date-lte': endDate,                // 受注日 <= 終了日時
    'receive_order_row_cancel_flag-eq': '0',          // キャンセルフラグ = 0 (有効データのみ)
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
    `取得結果: ${data.length}件 (このページのcount: ${count}件)`,
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
 * @details
 * 1回のAPI呼び出し制限（1000件）を超える大量のデータを取得するためのラッパー関数です。
 * `searchReceiveOrderRowPage` をループ実行し、全件を一つの配列に統合して返します。
 * 
 * 終了判定は「返却されたデータ件数 < limit」に基づいて行われます。
 * ネクストエンジンAPIの負荷軽減およびAPI制限を考慮し、各リクエストの間に `Utilities.sleep` 
 * による短い待機時間を設けています。
 * 
 * 処理の開始から終了までの経過時間や平均処理速度を計算し、ログに出力します。
 * 
 * @param {string} startDate - 検索開始日時（`YYYY-MM-DD HH:mm:ss` 形式）
 * @param {string} endDate - 検索終了日時（`YYYY-MM-DD HH:mm:ss` 形式）
 * @return {Array<Object>} 全ての受注明細データのフラットな配列
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
      logMessage(`API返却の全体件数: ${totalCount}件`);
    }

    // データがない場合は終了
    if (result.data.length === 0) {
      logMessage('データ取得完了: これ以上データがありません', LOG_LEVEL.SAMPLE);
      break;
    }

    // データを追加
    allData.push(...result.data);

    logMessage(`累計: ${allData.length}件`, LOG_LEVEL.SAMPLE);

    // 取得件数がlimitより少ない場合は最後のページ
    if (result.data.length < NE_API_LIMIT) {
      logMessage('最終ページに到達しました', LOG_LEVEL.SAMPLE);
      break;
    }

    // オフセットを更新
    offset += NE_API_LIMIT;

    // API負荷軽減のため少し待機
    Utilities.sleep(500); // 0.5秒待機
  }

  const endTime = new Date();
  const elapsedTime = (endTime - startTime) / 1000;

  logMessage('');
  logMessage('=== 受注明細全件取得完了 ===');
  logMessage(`実際の取得件数: ${allData.length}件`);
  logMessage(`ページ数: ${pageCount}ページ`);
  logMessage(`処理時間: ${elapsedTime.toFixed(2)}秒`);
  logMessage(`平均速度: ${(allData.length / elapsedTime).toFixed(0)}件/秒`);

  return allData;
}

// =============================================================================
// キャンセル除外フィルタリング
// =============================================================================

/**
 * キャンセル明細を除外
 * 
 * @details
 * APIレスポンス内の各明細行の `receive_order_row_cancel_flag` を確認し、
 * 値が "1" （キャンセル済み）のものを配列から取り除きます。
 * 
 * 通常、APIリクエスト時にフィルタリングを行いますが、データ整合性を担保するための
 * 二重チェック（セーフガード）として機能します。
 * 
 * @param {Array<Object>} data - フィルタリング前の受注明細データ配列
 * @return {Array<Object>} キャンセル行が除外された、有効な明細データのみの配列
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
 * @details
 * 認証用スクリプトプロパティ（トークン）が有効であるか、およびAPIサーバーと通信可能かを検証します。
 * 内部的に `NEAuth` ライブラリを使用し、ログインユーザー情報を取得します。
 * 
 * @return {Object} 接続成功時のユーザー情報（ライブラリからの戻り値）
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
 * @details
 * 基盤構築(01)の日付計算ロジックと、本ファイルの検索ロジックを統合してテストします。
 * テスト実行時間を短縮するため、`limit=10` に制限してリクエストを行い、
 * 戻り値の構造やフィールドの欠落がないかを確認します。
 * @return {Object} APIレスポンスの抜粋（data, count, hasMore）
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
 * @details
 * ページネーション（ループ処理）が正しく機能するかを検証します。
 * 7日間の全受注明細を取得するため、データ量によっては数秒〜数十秒の時間を要します。
 * 取得後のサンプルとして、配列の「先頭」と「末尾」を表示し、データの連続性を確認します。
 * 
 * @return {Array<Object>} 取得された全データ
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
 * デバッグ用: 日付範囲確認
 * 
 * @details
 * APIに渡す直前の日付パラメータの「型」と「フォーマット」をコンソールに出力します。
 */
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
  console.log('');
  console.log('✅ 日付範囲デバッグ完了');
}

/**
 * キャンセル除外テスト
 * 
 * @details
 * 実際にAPIから取得したデータに対して `filterCancelledRows` を適用し、
 * 除外前後の件数の差分を表示します。
 * また、キャンセルフラグごとの内訳（有効 vs キャンセル）を集計して出力します。
 * @return {Array<Object>} フィルタリング後のデータ
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
 * @details
 * 「API接続」フェーズで実装した全ての通信・処理工程をシーケンシャルに検証します。
 * 通信確認、単発検索、フィルタリングロジックの3段階をチェックします。
 * この統合テストが成功すれば、ネクストエンジンからのデータ抽出基盤は完成となります。
 * 
 * @throws {Error} いずれかのAPIリクエストまたはフィルタ処理で失敗した場合
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