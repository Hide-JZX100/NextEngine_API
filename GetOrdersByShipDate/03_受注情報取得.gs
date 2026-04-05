/**
 * 03_受注情報取得.gs
 * Phase 2：受注情報取得
 */

/**
 * ネクストエンジンAPIから受注明細を全件取得する（ページネーション対応）
 * @param {Date} startDate - 出荷予定日の開始日（nullの場合は自動計算）
 * @param {Date} endDate - 出荷予定日の終了日（nullの場合は自動計算）
 * @return {Array} 受注明細データの配列
 */
function fetchOrderRows(startDate, endDate) {
  const dateRange = getDateRange(startDate, endDate);
  const startStr = formatDateForApi(dateRange.start);
  const endStr = formatDateForApi(dateRange.end);

  console.log(`取得対象期間: ${startStr} ～ ${endStr}`);

  const props = PropertiesService.getScriptProperties();
  const token = {
    accessToken: props.getProperty('ACCESS_TOKEN'),
    refreshToken: props.getProperty('REFRESH_TOKEN')
  };

  if (!token.accessToken || !token.refreshToken) {
    throw new Error('ACCESS_TOKEN または REFRESH_TOKEN が取得できません。testGenerateAuthUrl()で再認証してください。');
  }

  const url = NE_API_BASE_URL + NE_ENDPOINT_ORDER_ROW;

  let allRows = [];
  let offset = 0;
  let hasNext = true;
  let pageCount = 1;

  while (hasNext) {
    const payload = {
      access_token: token.accessToken,
      refresh_token: token.refreshToken,
      wait_flag: '1',
      fields: ORDER_FIELDS,
      'receive_order_send_plan_date-gte': startStr,
      'receive_order_send_plan_date-lte': endStr,
      'receive_order_row_cancel_flag-eq': '0',
      'receive_order_cancel_type_id-eq': '0',
      offset: String(offset),
      limit: String(LIMIT)
    };

    const options = {
      method: 'post',
      contentType: 'application/x-www-form-urlencoded',
      payload: payload
    };

    const response = UrlFetchApp.fetch(url, options);
    const json = JSON.parse(response.getContentText());

    if (json.result !== 'success') {
      throw new Error('API取得エラー: ' + json.message + ' (code: ' + json.code + ')');
    }

    // トークンが更新された場合はプロパティを更新
    if (json.access_token && json.access_token !== token.accessToken) {
      props.setProperty('ACCESS_TOKEN', json.access_token);
      props.setProperty('REFRESH_TOKEN', json.refresh_token);
      token.accessToken = json.access_token;
      token.refreshToken = json.refresh_token;
    }

    const data = json.data || [];
    allRows = allRows.concat(data);

    console.log(`${pageCount}ページ目取得完了: ${data.length}件 (合計: ${allRows.length}件)`);

    if (data.length < LIMIT) {
      hasNext = false;
    } else {
      offset += LIMIT;
      pageCount++;
      Utilities.sleep(500); // 連続リクエストの間隔を空ける
    }
  }

  return allRows;
}

/**
 * ウォームアップ関数
 * 本番取得前に同じエンドポイントへ limit=1 で空打ちし、
 * コールドスタートの遅延を事前に吸収する
 * ※ウォームアップ失敗時は警告のみ出力し、本番処理は続行する
 */
function warmUp() {
  console.log('ウォームアップ開始...');

  const props = PropertiesService.getScriptProperties();
  const token = {
    accessToken: props.getProperty('ACCESS_TOKEN'),
    refreshToken: props.getProperty('REFRESH_TOKEN')
  };

  const url = NE_API_BASE_URL + NE_ENDPOINT_ORDER_ROW;
  const payload = {
    access_token: token.accessToken,
    refresh_token: token.refreshToken,
    wait_flag: '1',
    fields: 'receive_order_row_receive_order_id',
    offset: '0',
    limit: '1'
  };

  const options = {
    method: 'post',
    contentType: 'application/x-www-form-urlencoded',
    payload: payload
  };

  try {
    const startTime = new Date();
    const response = UrlFetchApp.fetch(url, options);
    const json = JSON.parse(response.getContentText());
    const elapsed = ((new Date() - startTime) / 1000).toFixed(1);

    // トークンが更新された場合はプロパティを更新
    if (json.access_token && json.access_token !== token.accessToken) {
      props.setProperty('ACCESS_TOKEN', json.access_token);
      props.setProperty('REFRESH_TOKEN', json.refresh_token);
      console.log('ウォームアップ中にトークンを更新しました');
    }

    console.log(`ウォームアップ完了 (${elapsed}秒)`);

    // 10分後に scheduledRun を動的トリガーで予約
    const newTrigger = ScriptApp.newTrigger('scheduledRun')
      .timeBased()
      .after(10 * 60 * 1000)
      .create();
    // 動的トリガーのUIDを保存（scheduledRun 側で削除するために使用）
    PropertiesService.getScriptProperties()
      .setProperty('WARMUP_TRIGGER_UIDS', newTrigger.getUniqueId());
    console.log('scheduledRun を10分後に予約しました (UID: ' + newTrigger.getUniqueId() + ')');

  } catch (e) {
    // ウォームアップ失敗は本番処理に影響させない
    console.warn('ウォームアップ失敗（本番処理は続行します）:', e.message);
  }
}

/**
 * APIレスポンスの1行データをスプレッドシート書き込み用配列に変換する
 * @param {Object} row - APIレスポンスの1件分データ
 * @return {Array} ORDER_HEADERS の順序に対応した値の配列
 */
function convertOrderRowToArray(row) {
  // 住所結合処理
  const address1 = row.receive_order_purchaser_address1 || '';
  const address2 = row.receive_order_purchaser_address2 || '';
  const purchaserAddress = address2 ? `${address1} ${address2}` : address1;

  const c_address1 = row.receive_order_consignee_address1 || '';
  const c_address2 = row.receive_order_consignee_address2 || '';
  const consigneeAddress = c_address2 ? `${c_address1} ${c_address2}` : c_address1;

  // 日付文字列を YYYY/MM/DD に変換する関数
  const toDateStr = (dateString) => {
    if (!dateString || dateString === '0000-00-00 00:00:00') return '';
    // YYYY-MM-DD 等の形式をパース
    const d = new Date(dateString.replace(/-/g, '/'));
    return formatDate(d);
  };

  return [
    toDateStr(row.receive_order_send_plan_date), // 出荷予定日
    '', // ロケーションコード（空文字）
    row.receive_order_row_goods_id || '', // 商品コード
    row.receive_order_row_goods_name || '', // 商品名
    row.receive_order_row_quantity || 0, // 受注数
    row.receive_order_row_sub_total_price || 0, // 小計
    row.receive_order_row_unit_price || 0, // 売単価
    row.receive_order_row_receive_order_id || '', // 伝票番号
    row.receive_order_row_shop_cut_form_id || '', // 受注番号
    toDateStr(row.receive_order_date), // 受注日
    row.receive_order_payment_method_name || '', // 支払方法
    row.receive_order_total_amount || 0, // 総合計
    row.receive_order_purchaser_name || '', // 購入者名
    row.receive_order_purchaser_kana || '', // 購入者カナ
    row.receive_order_purchaser_tel || '', // 購入者電話番号
    row.receive_order_purchaser_zip_code || '', // 購入者郵便番号
    purchaserAddress, // 購入者（住所1+住所2）
    row.receive_order_purchaser_mail_address || '', // 購入者メールアドレス
    row.receive_order_consignee_name || '', // 送り先名
    row.receive_order_consignee_kana || '', // 送り先カナ
    row.receive_order_consignee_tel || '', // 送り先電話番号
    row.receive_order_consignee_zip_code || '', // 送り先郵便番号
    consigneeAddress, // 送り先（住所1+住所2）
    row.receive_order_shop_id || '', // 店舗コード
    row.receive_order_picking_instruct || '', // ピッキング指示
    row.receive_order_delivery_name || '', // 発送方法
    row.receive_order_delivery_cut_form_id || '' // 発送伝票番号
  ];
}

/**
 * 受注明細データをスプレッドシートの受注情報シートに書き込む
 * @param {Array} rows - APIから取得した受注明細データの配列
 * @param {boolean} append - true の場合はクリアせずに最終行以降に追記する
 */
function writeOrdersToSheet(rows, append = false) {
  const config = getConfig();
  const ss = SpreadsheetApp.openById(config.targetSpreadsheetId);
  const sheet = getOrCreateSheet(ss, config.sheetNameOrders);

  if (!append) {
    // 追記モードでなければクリアしてヘッダーを設定
    clearAndSetHeader(sheet, ORDER_HEADERS);
  } else {
    // 追記モードの時、シートが完全に空であればヘッダーを設定
    if (sheet.getLastRow() === 0) {
      clearAndSetHeader(sheet, ORDER_HEADERS);
    }
  }

  if (rows && rows.length > 0) {
    const outputData = rows.map(convertOrderRowToArray);
    // 追記の場合はデータが存在する最終行の次から書き込み開始
    const startRow = append ? (sheet.getLastRow() + 1) : 2;
    sheet.getRange(startRow, 1, outputData.length, ORDER_HEADERS.length).setValues(outputData);
    console.log(`シート「${config.sheetNameOrders}」に ${outputData.length} 件${append ? '追記' : '書き込み'}しました。`);
  } else {
    console.log('取得データが0件のため、処理をスキップしました。');
  }
}

/**
 * 受注情報の取得から書き込みまでを一括実行するメイン関数
 * 手動実行・自動実行の両方から呼び出される
 * @param {Date|null} startDate - 開始日（省略時null：自動計算）
 * @param {Date|null} endDate - 終了日（省略時null：自動計算）
 * @param {boolean} forceAppend - true の場合は日付によらず強制的に追記する
 */
function updateOrders(startDate = null, endDate = null, forceAppend = false) {
  try {
    const range = getDateRange(startDate, endDate);
    console.log(`開始: updateOrders (${formatDate(range.start)} - ${formatDate(range.end)})`);

    // 開始日が1日（月初め）の場合のみ上書き、それ以外（10日・20日スタート）は追記
    const isAppendMode = forceAppend || (range.start.getDate() !== 1);

    const rows = fetchOrderRows(startDate, endDate);
    writeOrdersToSheet(rows, isAppendMode);

    console.log('完了: updateOrders');
  } catch (error) {
    console.error('updateOrders でエラーが発生しました:', error.stack || error.message);
    throw error;
  }
}

/**
 * testFetchOnly
 * 書き込みは行わず、fetchOrderRows(null, null) の結果件数と
 * 先頭3件のデータを console.log で出力するだけのテスト関数
 */
function testFetchOnly() {
  console.log('--- testFetchOnly 開始 ---');
  try {
    // const rows = fetchOrderRows(null, null);

    // 期間を明示的に指定（2025年10月）
    const start = new Date(2025, 9, 1);  // 9 = 10月
    const end = new Date(2025, 9, 31);

    // 変更: start, end を渡す
    const rows = fetchOrderRows(start, end);

    console.log(`取得件数: ${rows.length}件`);
    if (rows.length > 0) {
      const displayCount = Math.min(3, rows.length);
      console.log(`先頭${displayCount}件の生データ:`);
      for (let i = 0; i < displayCount; i++) {
        console.log(JSON.stringify(rows[i], null, 2));
      }
    }
    console.log('--- testFetchOnly 終了 ---');
  } catch (error) {
    console.error('エラー:', error.message);
  }
}

/**
 * updateOrders の動作確認用テスト関数
 * 実行日から自動計算された日付範囲で受注情報を取得・書き込みする
 * ※本番実行前に必ずこの関数でテストを行うこと
 */
function testUpdateOrders() {
  console.log('--- testUpdateOrders 開始 ---');
  // updateOrders(null, null);
  updateOrders(new Date(2025, 9, 1), new Date(2025, 9, 31));
  console.log('--- testUpdateOrders 終了 ---');
}
