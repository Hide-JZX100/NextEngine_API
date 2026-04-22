// ============================================================
// 01_設定ファイル.gs
// ネクストエンジン Amazon受注データ取得 - 定数・設定定義
// ============================================================

/** APIベースURL */
const NE_API_BASE_URL = 'https://api.next-engine.org';

/** 受注伝票検索エンドポイント */
const NE_ENDPOINT_ORDER_SEARCH = '/api_v1_receiveorder_base/search';

/** 受注状態区分：出荷確定済み */
const ORDER_STATUS_SHIPPED = '50';

/** 店舗ID */
const SHOP_ID_AMAZON     = '20';  // 本番：Amazon
const SHOP_ID_RAKUTEN    = '2';   // 開発：ベッド＆マットレス楽天市場店

/** 1回のAPIコールで取得する最大件数 */
const API_PAGE_LIMIT = 1000;

/** ページ間ウェイト（ミリ秒） */
const API_PAGE_WAIT_MS = 500;

/**
 * 取得フィールド定義
 * api  : ネクストエンジンAPIのフィールド名
 * header: スプレッドシートのヘッダー表示名
 */
const CONFIG_FIELDS = [
  { api: 'receive_order_id',                  header: '伝票番号'             },
  { api: 'receive_order_date',                header: '受注日'               },
  { api: 'receive_order_send_plan_date',      header: '出荷予定日'           },
  { api: 'receive_order_purchaser_name',      header: '購入者名'             },
  { api: 'receive_order_purchaser_zip_code',  header: '購入者郵便番号'       },
  { api: 'receive_order_purchaser_address1',  header: '購入者住所1'          },
  { api: 'receive_order_purchaser_address2',  header: '購入者住所2'          },
  { api: 'receive_order_purchaser_tel',       header: '購入者電話番号'       },
  { api: 'receive_order_purchaser_mail_address', header: '購入者メールアドレス' },
  { api: 'receive_order_consignee_name',      header: '送り先名'             },
  { api: 'receive_order_consignee_zip_code',  header: '送り先郵便番号'       },
  { api: 'receive_order_consignee_address1',  header: '送り先住所1'          },
  { api: 'receive_order_consignee_address2',  header: '送り先住所2'          },
  { api: 'receive_order_consignee_tel',       header: '送り先電話番号'       },
];

/**
 * 出力先スプレッドシートIDを取得する
 * スクリプトプロパティ OUTPUT_SPREADSHEET_ID から取得
 * 
 * 【使用タイミング】
 * - スプレッドシート書き込み処理時
 * 
 * @returns {string} スプレッドシートID
 */
function getOutputSpreadsheetId() {
  const props = PropertiesService.getScriptProperties();
  return props.getProperty('OUTPUT_SPREADSHEET_ID') || '';
}

/**
 * 出力先シートタブ名を取得する
 * スクリプトプロパティ OUTPUT_SHEET_NAME から取得（未設定時のデフォルトは 'API'）
 * 
 * 【使用タイミング】
 * - スプレッドシート書き込み処理時
 * 
 * @returns {string} シート名
 */
function getOutputSheetName() {
  const props = PropertiesService.getScriptProperties();
  return props.getProperty('OUTPUT_SHEET_NAME') || 'API';
}

/**
 * 開発モード判定
 * スクリプトプロパティ DEVELOPMENT_MODE が "true" の場合に開発モードとみなす
 * 
 * 【処理フロー】
 * 1. スクリプトプロパティを取得
 * 2. プロパティ値が "true" か判定して返す
 * 
 * 【使用タイミング】
 * - 取得先店舗の切り替えなど環境分岐時
 * 
 * @returns {boolean} true: 開発モード / false: 本番モード
 */
function isDevelopmentMode() {
  const props = PropertiesService.getScriptProperties();
  return props.getProperty('DEVELOPMENT_MODE') === 'true';
}

/**
 * 現在の環境に応じた店舗IDを返す
 * 
 * 【処理フロー】
 * 1. 開発モード判定関数を呼び出す
 * 2. 開発モードなら楽天市場店のID、本番モードならAmazonのIDを返す
 * 
 * 【使用タイミング】
 * - API実行時の抽出条件指定時
 * 
 * @returns {string} 店舗ID
 */
function getShopId() {
  return isDevelopmentMode() ? SHOP_ID_RAKUTEN : SHOP_ID_AMAZON;
}

/**
 * 【Phase 1 テスト用】動作確認関数
 * 各種定数・設定値・プロパティが正しく取得できているかログに出力する
 * 
 * 【処理フロー】
 * 1. 定数およびスクリプトプロパティの内容をログに出力する
 * 
 * 【使用タイミング】
 * - Phase 1の実装完了後の動作確認時（手動実行）
 */
function testPhase1() {
  console.log('=== Phase 1 テスト開始 ===');
  console.log('【環境設定】');
  console.log('　・開発モード: ' + isDevelopmentMode());
  console.log('　・対象店舗ID: ' + getShopId());
  
  console.log('【API設定】');
  console.log('　・取得ページリミット: ' + API_PAGE_LIMIT);
  console.log('　・対象フィールド数: ' + CONFIG_FIELDS.length);
  
  const spreadsheetId = getOutputSpreadsheetId();
  const sheetName = getOutputSheetName();
  
  console.log('【スプレッドシート設定】');
  console.log('　・スプレッドシートID: ' + (spreadsheetId ? spreadsheetId : '未設定（スクリプトプロパティ OUTPUT_SPREADSHEET_ID を設定してください）'));
  console.log('　・シートタブ名: ' + (getOutputSheetName() ? sheetName : '未設定（デフォルトの "API" が使用されます）'));
  
  console.log('=== Phase 1 テスト完了 ===');
}
