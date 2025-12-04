/**
 * =============================================================================
 * 売れ数転記スクリプト - 設定ファイル
 * =============================================================================
 * 
 * 【概要】
 * メーカー別転記、BtoB集計、タイムスタンプ更新の設定を一元管理します。
 * 
 * 【使い方】
 * - メーカー追加: MAKER_CONFIGS 配列に1行追加
 * - メーカー無効化: active を false に変更
 * - 設定変更: 各定数の値を修正
 * 
 * @version 1.0
 * @date 2025-12-01
 */

// =============================================================================
// 共通設定
// =============================================================================

/**
 * コピー元スプレッドシートID
 * 週間集計データが格納されているスプレッドシート
 */
const SOURCE_SPREADSHEET_ID = '1h0jeG2XOu8rRXlMcUYh2Gi67mx2HADMztpqusNLBc0E';

/**
 * コピーする行の設定
 */
const COPY_ROW_CONFIG = {
  startRow: 2,    // コピー開始行(週間データの開始)
  endRow: 8,      // コピー終了行(週間データの終了)
  rowCount: 7     // コピー行数(2行目～8行目 = 7行)
};

/**
 * コピー先共通シート名
 * すべてのメーカーで共通のシート名
 */
const TARGET_SHEET_NAME = 'SKU別売れ数';

// =============================================================================
// メーカー別転記設定
// =============================================================================

/**
 * メーカー別転記設定
 * 
 * @property {string} sheetName - コピー元シート名(週間集計シート名)
 * @property {string} targetId - コピー先スプレッドシートID
 * @property {string} name - メーカー名(ログ表示用)
 * @property {boolean} active - 有効/無効フラグ(true: 転記する, false: 転記しない)
 */
const MAKER_CONFIGS = [
  // === 有効なメーカー ===
  { sheetName: '週間集計_JCO', targetId: '1199zcuvnGL0BmaurTc1NNDLdEEgr5qzU146cjaaKyd0', name: 'JCO', active: true },
  { sheetName: '価格帯別集計', targetId: '1fL4K9HZJ72Hpe4lNP47deGpyYJI1qyM1WPOYugjTf7I', name: '価格帯別集計', active: true },
  { sheetName: '週間集計_ZHI', targetId: '1ueClrIM0FBV7T4kuwHNkK8fknkIqbGEEcJUuSt-bJgw', name: 'ZHI', active: true },
  { sheetName: '週間集計_DHF', targetId: '1_A2sQpaB5uFUUKkf-CB-Mxv7JUs5tWtigJqPhyeYKtI', name: 'DHF', active: true },
  { sheetName: '週間集計_MB', targetId: '1RFdJbE2y9jZUoWOXy2lJp0bmJ2PElcDRVxG_OIIpyEY', name: 'マットレスバンド', active: true },
  { sheetName: '週間集計_FSF', targetId: '15Q6fWIwCbdRX1D9OA--2xeFqJi7MfpWru-_4ysL9kQQ', name: 'FSF', active: true },
  { sheetName: '週間集計_JTM', targetId: '1Zbomb_q79I50zD0OVAsL1CBPgfmbse8HZfkED6L2nJQ', name: 'JTM', active: true },
  { sheetName: '週間集計_SLM', targetId: '1xoS-OlVhDD5qqvlNvZMGUpF0fF4GgcB792BMJedCP6I', name: 'SLM', active: true },
  { sheetName: '週間集計_CDF', targetId: '1P2_E4sZosuN2B7rGbQWa81obl9MU1Q5ryJQC6HNaedw', name: 'CDF', active: true },
  { sheetName: '週間集計_QAT', targetId: '1EfEQD-tW2UO7exzbRDm0rSKfyDD_VEqJpCoCBZUyIXM', name: 'QAT', active: true },
  { sheetName: '週間集計_ZNZ', targetId: '1iISRrQ-_BEzYXxNA5HlxXhxreItOCpdgeF5JShRS3LA', name: 'ZNZ', active: true },
  { sheetName: '週間集計_COT', targetId: '1gMwhCENoSUrhJjboGBp5I0Wj_O3AijjAqWxKLN_LVwg', name: 'COT', active: true },
  { sheetName: '週間集計_Enjoy', targetId: '1BpSNJbMGoVMSC8_ozKdjgW-3SOm578-LVxxzj3DBl2Q', name: 'Enjoy', active: true },
  { sheetName: '週間集計_万鵬家具', targetId: '1LnaHXiGFfCn7U78hX1gxxKI-Lq_ZEL0kVYwQLIcFUgg', name: '万鵬家具', active: true },
  { sheetName: '週間集計_WAM', targetId: '1Alzq_TqvoeykLPScN0UgM-TlUkBVna9Q8OmQyTRrwhc', name: 'WAM', active: true },
  { sheetName: '週間集計_DIA', targetId: '1-Href6oqzu1VOontNtaIAJHJjtHGhrnWn1C_I-sR-3g', name: 'DIA', active: true },
  { sheetName: '週間集計_MSI', targetId: '1obPmlVq8WCV1fG_EJXSMPeDK6PwaYu9AhZ1MpCjBT58', name: 'MSI', active: true },
  { sheetName: '週間集計_EON', targetId: '1c1OcweBAV-OCwMnBqrQszOXuwtLeJJ4c8lDy2W1Du8Q', name: 'EON', active: true },
  { sheetName: '週間集計_STL', targetId: '1xM0zNiBs26tCs5_2DZP_rKJLIgIP8YgQeVoMZHikDE8', name: 'STL', active: true },
  { sheetName: '週間集計_VIA', targetId: '1sUGiTXymLCZoetpsr1_biBGlf5UuDNuMDXzaYyL_OJM', name: 'VIA', active: true },
  { sheetName: '週間集計_CAN', targetId: '1QBRvzwYa2ngbKvAnyF6jf6c__eAR94jHGaIJqYODw3M', name: 'CAN', active: true },
  { sheetName: '週間集計_PIT', targetId: '1y1xucmACOUaZpOaITphy4WQAza9Aw3FCjk4TkuQxX7M', name: 'PIT', active: true },
  { sheetName: '週間集計_GFT', targetId: '13N3JXvoLERoOTlUz_p7rus1gjt39-H9s_WnZ0qsOCn0', name: 'GFT', active: true },
  { sheetName: 'サービス分集計', targetId: '1LqgzV9egbflf1fxgzhUMUG2zOHCg-Sry5prxNM4e16A', name: 'サービス分集計', active: true },
  { sheetName: '週間集計_FB_MT', targetId: '1Gz0PEHr9Luek_EfqMkjO_itrCOzKkqfsucRreI-i-CA', name: 'フランスベッド', active: true },
  { sheetName: '週間集計_国内マットレス', targetId: '1fMrHuf_LQ4N70JLyMkoDylKc0vQ0Xmpi05pRXiRhbDo', name: '国内マットレス', active: true },
  { sheetName: '週間集計_国内寝装品', targetId: '1L4ZuNvCYFedvTqMxYzdhGh5W4nOHMGtwozs5q2LPAII', name: '国内寝装品', active: true },
  { sheetName: '週間集計_熊井綿業', targetId: '1gAti4hx74-Ap9y4W1jfG9DQO1_QH5NlzLK8pQkwCg9c', name: '熊井綿業', active: true },
  
  // === 無効化されたメーカー ===
  { sheetName: '週間集計_水美家具', targetId: '1gmKncD9TaF4GxK0PutviQN5Le0_TdOcV18ZWObTy3dA', name: '水美家具', active: false },
  { sheetName: '週間集計_麦昆', targetId: '1wVo6vICu6mC5KWzg8QSnOpoacPFR0eKEEL5_XesBS0s', name: '麦昆', active: false },
  { sheetName: '週間集計_Lis', targetId: '1diqrL8SR6vWon2ZOAlWnkipW1hWoaRVOGgsNI8XcQvs', name: 'Lis', active: false },
  { sheetName: '週間集計_大鵬家具', targetId: '1voxbyc0t_YNeR0Gw0ONRUGTNbXIlf0uk8hCldNsNnXQ', name: '大鵬家具', active: false },
  { sheetName: '週間集計_JIN', targetId: '1PfvhEyrRF4asXfefZnAj4lCUmMsMTwPY7V5GUVEhbx8', name: 'JIN', active: false },
  { sheetName: '週間集計_MTF', targetId: '17DRulfss0flYFX_jrJxJajLmJFcGh41KkE7ldDHRZ40', name: 'MTF', active: false },
  { sheetName: '週間集計_音部・ロリアン寝装', targetId: '1YJZ9y1eKLzw61Fah_9dYOrEPHjHMdv0q-AJqubHqOCc', name: '音部・ロリアン寝装', active: false },
  { sheetName: '週間集計_FB_GS3', targetId: '1ftyD8W0EPhe0vtDiP6vsMdBIDOx9MOVAr0XvtFz7QvM', name: 'FB_GS3', active: false },
  { sheetName: '週間集計_FB_その他Sup', targetId: '1QABOENmE-Znt6a00cUGPv_f-AdULd9hj4E_yVMceyRU', name: 'FB_その他Sup', active: false },
  { sheetName: '週間集計_ドリームベッド', targetId: '1tNvDxKCEx5dQBrLeEcVMcoNF7qfNsF1ingqDKXb5gxw', name: 'ドリームベッド', active: false },
  { sheetName: '週間集計_国内その他', targetId: '1yl05SCUuu33lnnMWThmApkCuQk_S_cAFm9LSs9Ce28s', name: '国内その他', active: false }
];

// =============================================================================
// BtoB集計設定
// =============================================================================

/**
 * BtoB集計設定
 * 
 * @property {string} ssId - BtoB集計スプレッドシートID
 * @property {string} workSheetName - work シート名(抽出元)
 * @property {string} btobSheetName - BtoB シート名(蓄積先)
 * @property {Object} keyColumns - キー列のインデックス(0始まり)
 */
const BTOB_CONFIG = {
  ssId: '1ZX3fVnIqupZws-2N599qLZ05auGHDOdqf04E1OGXkTU',
  workSheetName: 'work',
  btobSheetName: 'BtoB',
  keyColumns: {
    denpyo: 0,   // A列: 伝票番号
    shohin: 2    // C列: 商品コード(伝票)
  }
};

// =============================================================================
// タイムスタンプ設定
// =============================================================================

/**
 * タイムスタンプ更新対象
 * 
 * @property {string} ssId - スプレッドシートID
 * @property {string} sheetName - シート名
 * @property {string} range - セル範囲(A1形式)
 * @property {string} label - ラベル(ログ表示用)
 */
const TIMESTAMP_TARGETS = [
  {
    ssId: '1X2ZHsY15-p4j0tsuqiAyJYcxh9SuYY05nLfZyD0weNk',
    sheetName: 'EON',
    range: 'A1',
    label: 'メーカー別発注状況一覧'
  },
  {
    ssId: '1grFbz3UQLMu7v3gmNTZOzW08EYBqsM6iMFZ7fFF9HLs',
    sheetName: '使い方',
    range: 'F1',
    label: 'データインポート用'
  }
];

// =============================================================================
// エラー通知設定
// =============================================================================

/**
 * エラー通知メール設定
 * 
 * @property {string} recipient - 通知先メールアドレス(省略時は実行者のメール)
 * @property {boolean} notifyOnSuccess - 成功時も通知するか(true: 通知する, false: 通知しない)
 */
const NOTIFICATION_CONFIG = {
  recipient: Session.getActiveUser().getEmail(),  // 実行者のメールアドレス
  notifyOnSuccess: false  // 成功時は通知しない
};

// =============================================================================
// ヘルパー関数
// =============================================================================

/**
 * 有効なメーカー設定のみを取得
 * 
 * @return {Array} 有効なメーカー設定の配列
 */
function getActiveMakers() {
  return MAKER_CONFIGS.filter(config => config.active);
}

/**
 * 無効なメーカー設定のみを取得
 * 
 * @return {Array} 無効なメーカー設定の配列
 */
function getInactiveMakers() {
  return MAKER_CONFIGS.filter(config => !config.active);
}

/**
 * 設定情報を表示(デバッグ用)
 */
function showConfig() {
  console.log('=== 売れ数転記スクリプト設定情報 ===');
  console.log('');
  
  console.log('【メーカー設定】');
  console.log(`有効: ${getActiveMakers().length}件`);
  console.log(`無効: ${getInactiveMakers().length}件`);
  console.log(`合計: ${MAKER_CONFIGS.length}件`);
  console.log('');
  
  console.log('【有効なメーカー一覧】');
  getActiveMakers().forEach((config, index) => {
    console.log(`${index + 1}. ${config.name} (${config.sheetName})`);
  });
  console.log('');
  
  console.log('【BtoB集計設定】');
  console.log(`スプレッドシートID: ${BTOB_CONFIG.ssId}`);
  console.log(`workシート: ${BTOB_CONFIG.workSheetName}`);
  console.log(`BtoBシート: ${BTOB_CONFIG.btobSheetName}`);
  console.log('');
  
  console.log('【タイムスタンプ更新対象】');
  TIMESTAMP_TARGETS.forEach((target, index) => {
    console.log(`${index + 1}. ${target.label} (${target.sheetName} ${target.range})`);
  });
  console.log('');
  
  console.log('【通知設定】');
  console.log(`通知先: ${NOTIFICATION_CONFIG.recipient}`);
  console.log(`成功時通知: ${NOTIFICATION_CONFIG.notifyOnSuccess ? 'あり' : 'なし'}`);
}

/**
 * =============================================================================
 * 売れ数転記スクリプト - 共通関数
 * =============================================================================
 * 
 * 【概要】
 * メーカー別転記、BtoB集計で使用する共通関数を提供します。
 * 
 * 【主な機能】
 * - スプレッドシート操作
 * - 日付判定
 * - データコピー
 * - エラーハンドリング
 * 
 * @version 1.0
 * @date 2025-12-01
 */

// =============================================================================
// スプレッドシート操作
// =============================================================================

/**
 * スプレッドシートを開く
 * 
 * @param {string} ssId - スプレッドシートID
 * @return {GoogleAppsScript.Spreadsheet.Spreadsheet} スプレッドシート
 * @throws {Error} スプレッドシートが見つからない場合
 */
function openSpreadsheet(ssId) {
  try {
    return SpreadsheetApp.openById(ssId);
  } catch (error) {
    throw new Error(`スプレッドシートを開けません: ${ssId}\n${error.message}`);
  }
}

/**
 * シートを取得
 * 
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss - スプレッドシート
 * @param {string} sheetName - シート名
 * @return {GoogleAppsScript.Spreadsheet.Sheet} シート
 * @throws {Error} シートが見つからない場合
 */
function getSheet(ss, sheetName) {
  const sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    throw new Error(`シートが見つかりません: ${sheetName}\nスプレッドシート: ${ss.getName()}`);
  }
  
  return sheet;
}

/**
 * シートの最終行を取得(空白を除く)
 * 
 * A列を基準に、値が入っている最終行を返します。
 * 
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - シート
 * @return {number} 最終行番号(1始まり)
 */
function getLastRowWithData(sheet) {
  const lastRow = sheet.getLastRow();
  
  if (lastRow <= 1) {
    return 1; // ヘッダー行のみ
  }
  
  // A列を下から順に確認
  for (let i = lastRow; i >= 1; i--) {
    const value = sheet.getRange(i, 1).getValue();
    if (value !== '') {
      return i;
    }
  }
  
  return 1; // データなし
}

// =============================================================================
// 日付判定
// =============================================================================

/**
 * コピー先の貼付開始行を計算
 * 
 * コピー先の最終行の日付とコピー元の日付を比較し、
 * 貼付開始行を決定します。
 * 
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sourceSheet - コピー元シート
 * @param {GoogleAppsScript.Spreadsheet.Sheet} targetSheet - コピー先シート
 * @return {number} 貼付開始行
 */
function calculateStartRow(sourceSheet, targetSheet) {
  // コピー先の最終行を取得
  const targetLastRow = getLastRowWithData(targetSheet);
  
  // コピー先が空の場合(ヘッダーのみ)
  if (targetLastRow <= 1) {
    return 2; // 2行目から開始
  }
  
  // コピー先の最終行の日付を取得
  const targetLastDate = targetSheet.getRange(targetLastRow, 1).getValue();
  
  // コピー元の日付を取得(2行目のA列)
  const sourceDate = sourceSheet.getRange(COPY_ROW_CONFIG.startRow, 1).getValue();
  
  // 日付を比較
  const targetDateStr = formatDateToString(targetLastDate);
  const sourceDateStr = formatDateToString(sourceDate);
  
  if (targetDateStr === sourceDateStr) {
    // 同じ日付 → 上書き(最終行から開始)
    return targetLastRow - COPY_ROW_CONFIG.rowCount + 1;
  } else {
    // 異なる日付 → 追記(次の行から開始)
    return targetLastRow + 1;
  }
}

/**
 * 日付をYYYY/MM/DD形式の文字列に変換
 * 
 * @param {Date|string} date - 日付
 * @return {string} YYYY/MM/DD形式の文字列
 */
function formatDateToString(date) {
  if (!date) {
    return '';
  }
  
  // 既に文字列の場合
  if (typeof date === 'string') {
    return date;
  }
  
  // Dateオブジェクトの場合
  if (date instanceof Date) {
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const day = String(date.getDate()).padStart(2, '0');
    return `${year}/${month}/${day}`;
  }
  
  return String(date);
}

// =============================================================================
// データコピー
// =============================================================================

/**
 * メーカーデータをコピー
 * 
 * コピー元シートからコピー先シートへ、週間データ(7行分)をコピーします。
 * 
 * @param {Object} config - メーカー設定オブジェクト
 * @return {Object} 実行結果 {success: boolean, message: string}
 */
function copyMakerData(config) {
  try {
    // コピー元を開く
    const sourceSS = openSpreadsheet(SOURCE_SPREADSHEET_ID);
    const sourceSheet = getSheet(sourceSS, config.sheetName);
    
    // コピー先を開く
    const targetSS = openSpreadsheet(config.targetId);
    const targetSheet = getSheet(targetSS, TARGET_SHEET_NAME);
    
    // 貼付開始行を計算
    const startRow = calculateStartRow(sourceSheet, targetSheet);
    
    // コピー元のデータを取得(2行目～8行目、全列)
    const lastColumn = sourceSheet.getLastColumn();
    const sourceRange = sourceSheet.getRange(
      COPY_ROW_CONFIG.startRow,
      1,
      COPY_ROW_CONFIG.rowCount,
      lastColumn
    );
    const sourceData = sourceRange.getValues();
    
    // コピー先に貼り付け
    const targetRange = targetSheet.getRange(
      startRow,
      1,
      COPY_ROW_CONFIG.rowCount,
      lastColumn
    );
    targetRange.setValues(sourceData);
    
    return {
      success: true,
      message: `${config.name}: コピー成功 (${startRow}行目～)`
    };
    
  } catch (error) {
    return {
      success: false,
      message: `${config.name}: コピー失敗 - ${error.message}`
    };
  }
}

// =============================================================================
// BtoB集計
// =============================================================================

/**
 * BtoB新規データを取得
 * 
 * workシートから新規データ(BtoBシートに存在しないデータ)を抽出します。
 * 
 * @return {Array<Array>} 新規データの配列
 */
function getBtoBNewData() {
  try {
    // スプレッドシートを開く
    const ss = openSpreadsheet(BTOB_CONFIG.ssId);
    const workSheet = getSheet(ss, BTOB_CONFIG.workSheetName);
    const btobSheet = getSheet(ss, BTOB_CONFIG.btobSheetName);
    
    // workシートのデータを取得
    const workLastRow = getLastRowWithData(workSheet);
    if (workLastRow <= 1) {
      // データなし
      return [];
    }
    
    const workData = workSheet.getRange(2, 1, workLastRow - 1, workSheet.getLastColumn()).getValues();
    
    // BtoBシートの既存データをSetに格納(高速化)
    const btobLastRow = getLastRowWithData(btobSheet);
    const btobKeys = new Set();
    
    if (btobLastRow > 1) {
      const btobData = btobSheet.getRange(2, 1, btobLastRow - 1, btobSheet.getLastColumn()).getValues();
      
      btobData.forEach(row => {
        const denpyo = String(row[BTOB_CONFIG.keyColumns.denpyo]);
        const shohin = String(row[BTOB_CONFIG.keyColumns.shohin]);
        const key = `${denpyo}_${shohin}`;
        btobKeys.add(key);
      });
    }
    
    // 新規データを抽出
    const newData = [];
    
    workData.forEach(row => {
      const denpyo = String(row[BTOB_CONFIG.keyColumns.denpyo]);
      const shohin = String(row[BTOB_CONFIG.keyColumns.shohin]);
      const key = `${denpyo}_${shohin}`;
      
      // BtoBシートに存在しない場合は新規データ
      if (!btobKeys.has(key)) {
        newData.push(row);
      }
    });
    
    return newData;
    
  } catch (error) {
    throw new Error(`BtoB新規データ取得エラー: ${error.message}`);
  }
}

/**
 * BtoB新規データを追記
 * 
 * 新規データをBtoBシートの最下行に追記します。
 * 
 * @param {Array<Array>} newData - 新規データ配列
 * @return {Object} 実行結果 {success: boolean, count: number, message: string}
 */
function appendBtoBData(newData) {
  try {
    if (newData.length === 0) {
      return {
        success: true,
        count: 0,
        message: 'BtoB集計: 新規データなし'
      };
    }
    
    // スプレッドシートを開く
    const ss = openSpreadsheet(BTOB_CONFIG.ssId);
    const btobSheet = getSheet(ss, BTOB_CONFIG.btobSheetName);
    
    // 最終行の次に追記
    const lastRow = getLastRowWithData(btobSheet);
    const startRow = lastRow + 1;
    
    // データを書き込み
    const range = btobSheet.getRange(startRow, 1, newData.length, newData[0].length);
    range.setValues(newData);
    
    return {
      success: true,
      count: newData.length,
      message: `BtoB集計: ${newData.length}件追加`
    };
    
  } catch (error) {
    return {
      success: false,
      count: 0,
      message: `BtoB集計エラー: ${error.message}`
    };
  }
}

// =============================================================================
// タイムスタンプ更新
// =============================================================================

/**
 * タイムスタンプを更新
 * 
 * 指定されたセルに現在時刻を書き込みます。
 * 
 * @param {Object} target - タイムスタンプ対象オブジェクト
 * @param {Date} timestamp - 記録する時刻
 * @return {Object} 実行結果 {success: boolean, message: string}
 */
function updateTimestamp(target, timestamp) {
  try {
    const ss = openSpreadsheet(target.ssId);
    const sheet = getSheet(ss, target.sheetName);
    
    sheet.getRange(target.range).setValue(timestamp);
    
    return {
      success: true,
      message: `${target.label}: タイムスタンプ更新`
    };
    
  } catch (error) {
    return {
      success: false,
      message: `${target.label}: タイムスタンプ更新失敗 - ${error.message}`
    };
  }
}

// =============================================================================
// エラー通知
// =============================================================================

/**
 * エラー通知メールを送信
 * 
 * @param {string} subject - 件名
 * @param {string} body - 本文
 */
function sendErrorNotification(subject, body) {
  try {
    MailApp.sendEmail(NOTIFICATION_CONFIG.recipient, subject, body);
  } catch (error) {
    console.error('メール送信エラー:', error.message);
  }
}

// =============================================================================
// テスト関数
// =============================================================================

/**
 * 日付判定テスト
 */
function testDateCalculation() {
  console.log('=== 日付判定テスト ===');
  
  try {
    const sourceSS = openSpreadsheet(SOURCE_SPREADSHEET_ID);
    const sourceSheet = getSheet(sourceSS, '週間集計_JCO');
    
    const targetSS = openSpreadsheet('1199zcuvnGL0BmaurTc1NNDLdEEgr5qzU146cjaaKyd0');
    const targetSheet = getSheet(targetSS, TARGET_SHEET_NAME);
    
    const startRow = calculateStartRow(sourceSheet, targetSheet);
    
    console.log('コピー元シート:', sourceSheet.getName());
    console.log('コピー先シート:', targetSheet.getName());
    console.log('貼付開始行:', startRow);
    console.log('');
    console.log('✅ 日付判定テスト完了');
    
  } catch (error) {
    console.error('❌ 日付判定テストエラー:', error.message);
  }
}

/**
 * BtoB集計テスト
 */
function testBtoBSummary() {
  console.log('=== BtoB集計テスト ===');
  
  try {
    const newData = getBtoBNewData();
    
    console.log('新規データ件数:', newData.length);
    
    if (newData.length > 0) {
      console.log('');
      console.log('【先頭3件のデータ】');
      newData.slice(0, 3).forEach((row, index) => {
        console.log(`[${index + 1}]`, row);
      });
    }
    
    console.log('');
    console.log('✅ BtoB集計テスト完了');
    
  } catch (error) {
    console.error('❌ BtoB集計テストエラー:', error.message);
  }
}

/**
 * =============================================================================
 * 売れ数転記スクリプト - メイン処理
 * =============================================================================
 * 
 * 【概要】
 * メーカー別転記、BtoB集計、タイムスタンプ更新を実行します。
 * 
 * 【実行順序】
 * 1. メーカー別転記処理
 * 2. BtoB集計処理
 * 3. タイムスタンプ更新
 * 
 * 【エラーハンドリング】
 * - メーカー転記: 個別にエラーハンドリング(1つ失敗しても続行)
 * - BtoB集計: エラー時は記録のみ
 * - タイムスタンプ: 成功したメーカーがあれば更新
 * 
 * @version 1.0
 * @date 2025-12-01
 */

// =============================================================================
// メイン処理
// =============================================================================

/**
 * 日次更新メイン処理
 * 
 * メーカー別転記、BtoB集計、タイムスタンプ更新を実行します。
 * エラーが発生してもできる限り処理を続行し、最後に結果をまとめて表示します。
 * 
 * @return {Object} 実行結果
 */
function dailyUpdate() {
  console.log('╔════════════════════════════════════════════════════════════╗');
  console.log('║  売れ数転記スクリプト - 日次更新開始                     ║');
  console.log('╚════════════════════════════════════════════════════════════╝');
  console.log('');
  
  const startTime = new Date();
  const results = {
    maker: [],
    btob: null,
    timestamp: [],
    startTime: startTime,
    endTime: null,
    elapsedTime: 0
  };
  
  try {
    // ========================================
    // 1. メーカー別転記処理
    // ========================================
    console.log('【1】メーカー別転記処理');
    results.maker = executeMakerCopy();
    console.log('');
    
    // ========================================
    // 2. BtoB集計処理
    // ========================================
    console.log('【2】BtoB集計処理');
    results.btob = executeBtoBSummary();
    console.log('');
    
    // ========================================
    // 3. タイムスタンプ更新
    // ========================================
    console.log('【3】タイムスタンプ更新');
    
    // 成功したメーカーがある場合のみタイムスタンプ更新
    const successCount = results.maker.filter(r => r.success).length;
    
    if (successCount > 0) {
      const timestamp = new Date();
      results.timestamp = updateAllTimestamps(timestamp);
    } else {
      console.log('⚠️ すべてのメーカー転記が失敗したため、タイムスタンプは更新しません');
      results.timestamp = [];
    }
    console.log('');
    
    // ========================================
    // 4. 実行結果サマリー
    // ========================================
    results.endTime = new Date();
    results.elapsedTime = (results.endTime - startTime) / 1000;
    
    displaySummary(results);
    
    // エラーがあった場合は通知
    const hasError = checkForErrors(results);
    if (hasError) {
      sendErrorReport(results);
    }
    
    return results;
    
  } catch (error) {
    // 予期しないエラー
    console.error('');
    console.error('╔════════════════════════════════════════════════════════════╗');
    console.error('║  ❌ 致命的エラー発生                                     ║');
    console.error('╚════════════════════════════════════════════════════════════╝');
    console.error('');
    console.error('エラー内容:', error.message);
    console.error('スタックトレース:', error.stack);
    
    // エラー通知
    const subject = '❌ [エラー] 売れ数転記スクリプト - 致命的エラー';
    const body = `
売れ数転記スクリプトで致命的エラーが発生しました。

【エラー内容】
${error.message}

【スタックトレース】
${error.stack}

【発生時刻】
${new Date().toLocaleString('ja-JP')}

【対処方法】
GASエディタで実行ログを確認してください。
URL: https://script.google.com/home/projects/${ScriptApp.getScriptId()}/edit
    `;
    
    sendErrorNotification(subject, body);
    
    throw error;
  }
}

// =============================================================================
// 個別処理実行
// =============================================================================

/**
 * メーカー別転記処理を実行
 * 
 * 有効なすべてのメーカーのデータを転記します。
 * 個別にエラーハンドリングを行い、1つ失敗しても続行します。
 * 
 * @return {Array} 実行結果の配列
 */
function executeMakerCopy() {
  const activeMakers = getActiveMakers();
  const results = [];
  
  console.log(`有効なメーカー: ${activeMakers.length}件`);
  console.log('');
  
  activeMakers.forEach((config, index) => {
    const makerStartTime = new Date();
    
    console.log(`[${index + 1}/${activeMakers.length}] ${config.name} 処理中...`);
    
    const result = copyMakerData(config);
    
    const makerEndTime = new Date();
    const makerElapsedTime = ((makerEndTime - makerStartTime) / 1000).toFixed(2);
    
    results.push({
      name: config.name,
      success: result.success,
      message: result.message,
      elapsedTime: makerElapsedTime
    });
    
    if (result.success) {
      console.log(`  ✅ ${result.message} (${makerElapsedTime}秒)`);
    } else {
      console.error(`  ❌ ${result.message} (${makerElapsedTime}秒)`);
    }
  });
  
  // サマリー表示
  const successCount = results.filter(r => r.success).length;
  const failCount = results.filter(r => !r.success).length;
  
  console.log('');
  console.log('【メーカー転記結果】');
  console.log(`成功: ${successCount}件`);
  console.log(`失敗: ${failCount}件`);
  console.log(`合計: ${activeMakers.length}件`);
  
  return results;
}

/**
 * BtoB集計処理を実行
 * 
 * workシートから新規データを取得し、BtoBシートに追記します。
 * 
 * @return {Object} 実行結果
 */
function executeBtoBSummary() {
  try {
    console.log('BtoB新規データ取得中...');
    
    const newData = getBtoBNewData();
    
    console.log(`新規データ: ${newData.length}件`);
    
    if (newData.length === 0) {
      console.log('✅ BtoB集計: 新規データなし');
      return {
        success: true,
        count: 0,
        message: '新規データなし'
      };
    }
    
    console.log('BtoBシートに追記中...');
    
    const result = appendBtoBData(newData);
    
    if (result.success) {
      console.log(`✅ ${result.message}`);
    } else {
      console.error(`❌ ${result.message}`);
    }
    
    return result;
    
  } catch (error) {
    const errorMessage = `BtoB集計エラー: ${error.message}`;
    console.error(`❌ ${errorMessage}`);
    
    return {
      success: false,
      count: 0,
      message: errorMessage
    };
  }
}

/**
 * すべてのタイムスタンプを更新
 * 
 * @param {Date} timestamp - 記録する時刻
 * @return {Array} 実行結果の配列
 */
function updateAllTimestamps(timestamp) {
  const results = [];
  
  TIMESTAMP_TARGETS.forEach(target => {
    const result = updateTimestamp(target, timestamp);
    results.push(result);
    
    if (result.success) {
      console.log(`✅ ${result.message}`);
    } else {
      console.error(`❌ ${result.message}`);
    }
  });
  
  return results;
}

// =============================================================================
// 結果表示
// =============================================================================

/**
 * 実行結果サマリーを表示
 * 
 * @param {Object} results - 実行結果オブジェクト
 */
function displaySummary(results) {
  console.log('╔════════════════════════════════════════════════════════════╗');
  console.log('║  実行結果サマリー                                         ║');
  console.log('╚════════════════════════════════════════════════════════════╝');
  console.log('');
  
  // 実行時間
  console.log('【実行時間】');
  console.log(`開始: ${results.startTime.toLocaleString('ja-JP')}`);
  console.log(`終了: ${results.endTime.toLocaleString('ja-JP')}`);
  console.log(`所要時間: ${results.elapsedTime.toFixed(2)}秒`);
  console.log('');
  
  // メーカー転記
  const makerSuccess = results.maker.filter(r => r.success).length;
  const makerFail = results.maker.filter(r => !r.success).length;
  
  console.log('【メーカー別転記】');
  console.log(`成功: ${makerSuccess}件`);
  console.log(`失敗: ${makerFail}件`);
  
  if (makerFail > 0) {
    console.log('');
    console.log('失敗したメーカー:');
    results.maker
      .filter(r => !r.success)
      .forEach(r => {
        console.log(`  - ${r.name}: ${r.message}`);
      });
  }
  console.log('');
  
  // BtoB集計
  console.log('【BtoB集計】');
  if (results.btob) {
    console.log(`ステータス: ${results.btob.success ? '成功' : '失敗'}`);
    console.log(`追加件数: ${results.btob.count}件`);
    console.log(`メッセージ: ${results.btob.message}`);
  } else {
    console.log('ステータス: 未実行');
  }
  console.log('');
  
  // タイムスタンプ
  const tsSuccess = results.timestamp.filter(r => r.success).length;
  const tsFail = results.timestamp.filter(r => !r.success).length;
  
  console.log('【タイムスタンプ更新】');
  console.log(`成功: ${tsSuccess}件`);
  console.log(`失敗: ${tsFail}件`);
  console.log('');
  
  // 総合判定
  const hasError = makerFail > 0 || (results.btob && !results.btob.success) || tsFail > 0;
  
  if (hasError) {
    console.log('⚠️ 一部の処理でエラーが発生しました');
  } else {
    console.log('✅ すべての処理が正常に完了しました');
  }
}

/**
 * エラーの有無をチェック
 * 
 * @param {Object} results - 実行結果オブジェクト
 * @return {boolean} エラーがある場合 true
 */
function checkForErrors(results) {
  const makerFail = results.maker.filter(r => !r.success).length;
  const btobFail = results.btob && !results.btob.success;
  const tsFail = results.timestamp.filter(r => !r.success).length;
  
  return makerFail > 0 || btobFail || tsFail > 0;
}

/**
 * エラーレポートをメール送信
 * 
 * @param {Object} results - 実行結果オブジェクト
 */
function sendErrorReport(results) {
  const subject = '⚠️ [エラーあり] 売れ数転記スクリプト実行結果';
  
  // メーカー転記エラー
  const makerErrors = results.maker
    .filter(r => !r.success)
    .map(r => `  - ${r.name}: ${r.message}`)
    .join('\n');
  
  // BtoB集計エラー
  const btobError = (results.btob && !results.btob.success) 
    ? `エラー: ${results.btob.message}` 
    : '正常';
  
  // タイムスタンプエラー
  const tsErrors = results.timestamp
    .filter(r => !r.success)
    .map(r => `  - ${r.message}`)
    .join('\n');
  
  const body = `
売れ数転記スクリプトでエラーが発生しました。

【実行時間】
開始: ${results.startTime.toLocaleString('ja-JP')}
終了: ${results.endTime.toLocaleString('ja-JP')}
所要時間: ${results.elapsedTime.toFixed(2)}秒

【メーカー別転記】
成功: ${results.maker.filter(r => r.success).length}件
失敗: ${results.maker.filter(r => !r.success).length}件

失敗したメーカー:
${makerErrors || 'なし'}

【BtoB集計】
${btobError}

【タイムスタンプ更新】
成功: ${results.timestamp.filter(r => r.success).length}件
失敗: ${results.timestamp.filter(r => !r.success).length}件

${tsErrors ? 'エラー:\n' + tsErrors : ''}

【対処方法】
GASエディタで詳細ログを確認してください。
URL: https://script.google.com/home/projects/${ScriptApp.getScriptId()}/edit

このメールは自動送信されています。
  `;
  
  sendErrorNotification(subject, body);
  
  console.log('');
  console.log(`エラー通知メールを送信しました: ${NOTIFICATION_CONFIG.recipient}`);
}

// =============================================================================
// テスト関数
// =============================================================================

/**
 * メイン処理テスト
 * 
 * ⚠️ 実際にデータ転記が実行されます!
 */
function testDailyUpdate() {
  console.log('=== 日次更新テスト ===');
  console.log('⚠️ 実際にデータ転記が実行されます!');
  console.log('');
  
  const results = dailyUpdate();
  
  return results;
}