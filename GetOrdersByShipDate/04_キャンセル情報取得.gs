/**
 * 04_キャンセル情報取得.gs
 * Phase 3：キャンセル情報取得
 */

/**
 * ネクストエンジンAPIからキャンセル情報を全件取得する（ページネーション対応）
 *
 * 【検索ロジック】
 * 出荷予定日（receive_order_send_plan_date）が指定期間内で、かつ受注キャンセル区分（receive_order_cancel_type_id）が0以外の受注を取得する。
 * 注文分割された明細を除外するため、明細行キャンセルフラグではなく受注キャンセル区分を条件とする。
 * 出荷後キャンセルも対象とするため、出荷予定日ベースで絞り込む。
 *
 * @param {Date} startDate - 受注日の開始日（nullの場合は自動計算：前月1日）
 * @param {Date} endDate - 受注日の終了日（nullの場合は自動計算：前月末日）
 * @return {Array} キャンセル明細データの配列
 */
function fetchCancelRows(startDate, endDate) {
  let start, end;
  if (startDate && endDate) {
    start = startDate;
    end = endDate;
  } else {
    // 省略時は前月1日〜前月末日
    const today = new Date();
    const year = today.getFullYear();
    const month = today.getMonth(); // 0-based
    start = new Date(year, month - 1, 1);
    end = new Date(year, month, 0); // 前月末日
  }

  const startStr = formatDateForApi(start);
  const endStr = formatDateForApi(end);

  console.log(`キャンセル取得対象期間(出荷予定日ベース): ${startStr} ～ ${endStr}`);

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
      fields: CANCEL_FIELDS,
      'receive_order_send_plan_date-gte': startStr,
      'receive_order_send_plan_date-lte': endStr,
      'receive_order_cancel_type_id-neq': '0',
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
      throw new Error('API取得エラー(キャンセル): ' + json.message + ' (code: ' + json.code + ')');
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

    console.log(`キャンセル情報 ${pageCount}ページ目取得完了: ${data.length}件 (合計: ${allRows.length}件)`);

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
 * APIレスポンスの1行データをキャンセル情報シート書き込み用配列に変換する
 * @param {Object} row - APIレスポンスの1件分データ
 * @param {Map} shopNameMap - 店舗コードと店舗名のマッピング
 * @return {Array} CANCEL_HEADERS の順序に対応した値の配列
 */
function convertCancelRowToArray(row, shopNameMap) {
  const shopCode = String(row.receive_order_shop_id || '');
  const shopName = shopNameMap.has(shopCode) ? shopNameMap.get(shopCode) : `不明(${shopCode})`;

  const toDateStr = (dateString) => {
    if (!dateString || dateString === '0000-00-00 00:00:00') return '';
    const d = new Date(dateString.replace(/-/g, '/'));
    return formatDate(d);
  };

  return [
    shopName, // 店舗名
    row.receive_order_row_receive_order_id || '', // 伝票番号
    toDateStr(row.receive_order_cancel_date), // キャンセル日
    row.receive_order_row_goods_id || '', // 商品コード（伝票）
    row.receive_order_row_goods_name || '', // 商品名（伝票）
    row.receive_order_row_quantity || 0, // 受注数
    row.receive_order_row_unit_price || 0, // 小計
    row.receive_order_delivery_fee_amount || 0, // 発送代
    toDateStr(row.receive_order_date), // 受注日
    shopCode // 店舗コード
  ];
}

/**
 * キャンセルデータをスプレッドシートのキャンセル情報シートに書き込む
 * @param {Array} rows - APIから取得したキャンセルデータの配列
 * @param {Map} shopNameMap - 店舗コードと店舗名のマッピング
 */
function writeCancelsToSheet(rows, shopNameMap) {
  const config = getConfig();
  const ss = SpreadsheetApp.openById(config.targetSpreadsheetId);
  const sheet = getOrCreateSheet(ss, config.sheetNameCancel);

  clearAndSetHeader(sheet, CANCEL_HEADERS);

  if (rows && rows.length > 0) {
    const outputData = rows.map(row => convertCancelRowToArray(row, shopNameMap));
    sheet.getRange(2, 1, outputData.length, CANCEL_HEADERS.length).setValues(outputData);
    console.log(`シート「${config.sheetNameCancel}」にキャンセル情報を ${outputData.length} 件書き込みました。`);
  } else {
    console.log('キャンセルデータが0件のため、ヘッダーのみ設定しました。');
  }
}

/**
 * キャンセル情報の取得から書き込みまでを一括実行するメイン関数
 * 毎月1日のみ scheduledRun から呼び出される。手動再実行時は manualRunCancels から呼び出す。
 * @param {Date} startDate - 受注日の開始日（省略時null：自動計算）
 * @param {Date} endDate - 受注日の終了日（省略時null：自動計算）
 */
function updateCancels(startDate = null, endDate = null) {
  try {
    let start, end;
    if (startDate && endDate) {
      start = startDate;
      end = endDate;
    } else {
      const today = new Date();
      start = new Date(today.getFullYear(), today.getMonth() - 1, 1);
      end = new Date(today.getFullYear(), today.getMonth(), 0);
    }

    console.log(`開始: updateCancels (${formatDate(start)} - ${formatDate(end)})`);

    // 店舗マスタ取得
    const shopNameMap = getShopNameMap();

    const rows = fetchCancelRows(startDate, endDate);
    writeCancelsToSheet(rows, shopNameMap);

    console.log('完了: updateCancels');
  } catch (error) {
    console.error('updateCancels でエラーが発生しました:', error.stack || error.message);
    throw error;
  }
}

/**
 * testFetchCancelOnly
 * 書き込みは行わず、fetchCancelRows() の結果件数と先頭3件を出力
 * 今回はテスト用に 2025/10/1〜2025/10/31 を明示的に指定
 */
function testFetchCancelOnly() {
  console.log('--- testFetchCancelOnly 開始 (2025/10/1〜2025/10/31) ---');
  try {
    const start = new Date(2025, 9, 1);
    const end = new Date(2025, 9, 31);

    const rows = fetchCancelRows(start, end);
    console.log(`キャンセルデータ取得件数: ${rows.length}件`);
    if (rows.length > 0) {
      const displayCount = Math.min(3, rows.length);
      console.log(`先頭${displayCount}件の生データ:`);
      for (let i = 0; i < displayCount; i++) {
        console.log(JSON.stringify(rows[i], null, 2));
      }
    }
  } catch (error) {
    console.error('エラー:', error.message);
  }
  console.log('--- testFetchCancelOnly 終了 ---');
}

/**
 * updateCancels の動作確認用テスト関数
 * 今回はテスト環境のデータがある 2025/10/1〜2025/10/31 を指定して取得・書き込みする
 */
function testUpdateCancels() {
  console.log('--- testUpdateCancels 開始 (2025/10/1〜2025/10/31) ---');
  const start = new Date(2025, 9, 1);
  const end = new Date(2025, 9, 31);
  updateCancels(start, end);
  console.log('--- testUpdateCancels 終了 ---');
}
