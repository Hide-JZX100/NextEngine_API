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
  // 注意：NEAuthライブラリが有効である前提です
  const token = NEAuth.getValidToken(props);
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
 */
function writeOrdersToSheet(rows) {
  const config = getConfig();
  const ss = SpreadsheetApp.openById(config.targetSpreadsheetId);
  const sheet = getOrCreateSheet(ss, config.sheetNameOrders);
  
  clearAndSetHeader(sheet, ORDER_HEADERS);

  if (rows && rows.length > 0) {
    const outputData = rows.map(convertOrderRowToArray);
    sheet.getRange(2, 1, outputData.length, ORDER_HEADERS.length).setValues(outputData);
    console.log(`シート「${config.sheetNameOrders}」に ${outputData.length} 件書き込みました。`);
  } else {
    console.log('取得データが0件のため、ヘッダーのみ設定しました。');
  }
}

/**
 * 受注情報の取得から書き込みまでを一括実行するメイン関数
 * 手動実行・自動実行の両方から呼び出される
 * @param {Date|null} startDate - 開始日（省略時null：自動計算）
 * @param {Date|null} endDate - 終了日（省略時null：自動計算）
 */
function updateOrders(startDate = null, endDate = null) {
  try {
    const range = getDateRange(startDate, endDate);
    console.log(`開始: updateOrders (${formatDate(range.start)} - ${formatDate(range.end)})`);
    
    const rows = fetchOrderRows(startDate, endDate);
    writeOrdersToSheet(rows);
    
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
    const rows = fetchOrderRows(null, null);
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
  updateOrders(null, null);
  console.log('--- testUpdateOrders 終了 ---');
}
