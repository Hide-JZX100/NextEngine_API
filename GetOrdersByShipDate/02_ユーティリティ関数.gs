/**
 * @file 02_ユーティリティ関数.gs
 * @description Phase 1：基盤構築。日付操作、店舗名マップ取得、シート操作などの汎用的なユーティリティ関数群を管理します。
 */

/**
 * 期間の開始日と終了日を計算する
 * @param {Date|null} startDate - 開始日
 * @param {Date|null} endDate - 終了日
 * @return {Object} { start: Date, end: Date }
 * 引数が両方nullの場合、実行日から自動計算する:
 *  - 1日の場合：前月20日〜前月末日
 *  - 10日の場合：当月1日〜9日
 *  - 20日の場合：当月10日〜19日
 *  - それ以外の日付の場合：当月1日〜実行前日
 */
function getDateRange(startDate, endDate) {
  if (startDate && endDate) {
    return { start: startDate, end: endDate };
  }

  const today = new Date();
  const date = today.getDate();
  const year = today.getFullYear();
  const month = today.getMonth(); // 0-based

  let start, end;

  if (date === 3) {
    // 実行日が3日の場合：前月20日〜前月末日
    start = new Date(year, month - 1, 20);
    end = new Date(year, month, 0); // 前月末日
  } else if (date === 13) {
    // 実行日が13日の場合：当月1日〜9日（月の上書き起点）
    start = new Date(year, month, 1);
    end = new Date(year, month, 9);
  } else if (date === 23) {
    // 実行日が23日の場合：当月10日〜19日
    start = new Date(year, month, 10);
    end = new Date(year, month, 19);
  } else {
    // それ以外の日付の場合（手動実行）：当月1日〜実行前日
    start = new Date(year, month, 1);
    end = new Date(year, month, date - 1);
  }

  return { start: start, end: end };
}

/**
 * 日付を 'YYYY/MM/DD' 形式の文字列にフォーマットする
 * @param {Date} date - 日付
 * @return {string} 'YYYY/MM/DD' 形式の文字列
 */
function formatDate(date) {
  if (!date || isNaN(date.getTime())) return '';
  const yyyy = date.getFullYear();
  const mm = String(date.getMonth() + 1).padStart(2, '0');
  const dd = String(date.getDate()).padStart(2, '0');
  return `${yyyy}/${mm}/${dd}`;
}

/**
 * 日付をAPI用の 'YYYY-MM-DD' 形式の文字列にフォーマットする
 * @param {Date} date - 日付
 * @return {string} 'YYYY-MM-DD' 形式の文字列
 */
function formatDateForApi(date) {
  if (!date || isNaN(date.getTime())) return '';
  const yyyy = date.getFullYear();
  const mm = String(date.getMonth() + 1).padStart(2, '0');
  const dd = String(date.getDate()).padStart(2, '0');
  return `${yyyy}-${mm}-${dd}`;
}

/**
 * 店舗情報マスタから「店舗コード・店舗名」の対応表を取得する
 * @return {Map} キー：店舗コード文字列、値：店舗名文字列 の Map
 */
function getShopNameMap() {
  const config = getConfig();
  const masterSs = SpreadsheetApp.openById(config.masterSpreadsheetId);
  const sheet = masterSs.getSheetByName(config.sheetNameShop);

  if (!sheet) {
    throw new Error('店舗情報マスタシートが見つかりません: ' + config.sheetNameShop);
  }

  const data = sheet.getDataRange().getValues();
  const shopMap = new Map();

  // 空行は無視
  for (let i = 0; i < data.length; i++) {
    const code = String(data[i][SHOP_CODE_COLUMN - 1]).trim();
    const name = String(data[i][SHOP_NAME_COLUMN - 1]).trim();

    if (code !== '' && name !== '') {
      shopMap.set(code, name);
    }
  }

  return shopMap;
}

/**
 * 指定したシートが存在しなければ作成して返す
 * @param {SpreadsheetApp.Spreadsheet} spreadsheet - スプレッドシートオブジェクト
 * @param {string} sheetName - シート名
 * @return {SpreadsheetApp.Sheet} シートオブジェクト
 */
function getOrCreateSheet(spreadsheet, sheetName) {
  let sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
    console.log('シートを作成しました: ' + sheetName);
  }
  return sheet;
}

/**
 * シートの全内容をクリアし、1行目にヘッダーを書き込む
 * @param {SpreadsheetApp.Sheet} sheet - シートオブジェクト
 * @param {Array} headers - ヘッダー配列
 */
function clearAndSetHeader(sheet, headers) {
  sheet.clear();
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#f3f3f3'); // 薄いグレー
}
