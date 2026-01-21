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

  // 1. 設定値の取得（共通ライブラリ.gsの関数を使用）
  const config = getAppConfig();
  const token = config.accessToken;
  const refreshToken = config.refreshToken;

  // 2. 各マスタの同期処理を実行

  // 店舗マスタ同期（店舗マスタ同期.gs）
  syncShopMaster(config, token, refreshToken);

  // モールマスタ同期（モールマスタ同期.gs）
  syncMallMaster(config, token, refreshToken);

  // 受注キャンセル区分マスタ同期（受注キャンセル区分マスタ同期.gs）
  syncCancelTypeMaster(config, token, refreshToken);

  Logger.log('--- マスタ情報同期処理 完了 ---');
}