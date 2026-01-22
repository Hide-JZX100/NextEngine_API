/**
 * ネクストエンジン マスタ情報同期スクリプト (GAS)
 * 目的: 店舗マスタ、モール/カートマスタ、受注キャンセル区分マスタを
 *       ネクストエンジンAPIから取得し、Googleスプレッドシートに自動で書き込む。
 * 実行トリガー: 毎週1回
 * 
 * 変更履歴:
 * - 2026-01-21: リファクタリング - 共通関数を共通ライブラリ.gsに移動、各マスタ同期処理を別ファイルに分離
 * - 認証の仕組みを統合し、APIコール時にトークンを自動更新するロジックを追加
 */

// ====================================================================
// メイン実行関数 (トリガー設定用)
// ====================================================================

/**
 * 週間トリガーに設定するメイン関数
 * 全てのマスタ同期処理を呼び出して実行する
 */
function mainMasterSync() {
  Logger.log('--- マスタ情報同期処理 開始 ---');

  // 1. 設定値の取得（初期）
  let config = getAppConfig();
  let token = config.accessToken;
  let refreshToken = config.refreshToken;


  // 2.【店舗マスタ同期】
  // 直前の処理でトークンが更新されている可能性があるため再読込は不要（初回なので）
  syncShopMaster(config, token, refreshToken);


  // 3.【モールマスタ同期】
  // 店舗マスタ同期でトークンが更新された可能性があるため、最新の値を再取得
  config = getAppConfig();
  token = config.accessToken;
  refreshToken = config.refreshToken;

  syncMallMaster(config, token, refreshToken);


  // 4.【受注キャンセル区分マスタ同期】
  // モールマスタ同期でトークンが更新された可能性があるため、再取得
  config = getAppConfig();
  token = config.accessToken;
  refreshToken = config.refreshToken;

  syncCancelTypeMaster(config, token, refreshToken);


  // 5.【支払区分マスタ同期】
  // キャンセル区分マスタ同期でトークンが更新された可能性があるため、再取得
  config = getAppConfig();
  token = config.accessToken;
  refreshToken = config.refreshToken;

  syncPaymentMethodMaster(config, token, refreshToken);


  Logger.log('--- マスタ情報同期処理 完了 ---');
}