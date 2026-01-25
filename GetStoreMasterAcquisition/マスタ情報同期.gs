/**
 * マスタ情報同期.gs
 * 
 * 【ファイルの役割】
 * 本プロジェクトのメインエントリポイント（実行開始地点）となるスクリプトです。
 * Google Apps Scriptのトリガー機能（時間主導型トリガーなど）からは、本ファイル内の
 * `mainMasterSync` 関数を指定して実行してください。
 * 
 * 【処理概要】
 * 以下の順序で各マスタ情報の同期（取得・スプレッドシート書き込み）を一括で行います。
 * 1. 店舗マスタ
 * 2. モール/カートマスタ
 * 3. 受注キャンセル区分マスタ
 * 4. 支払区分マスタ
 * 5. 仕入先マスタ
 * 
 * 【重要な仕様: トークン管理】
 * ネクストエンジンAPIは呼び出し毎にトークンが更新される仕様です。
 * そのため、本スクリプトでは各マスタ同期処理の直前に必ず最新のトークンを
 * スクリプトプロパティから再読み込みする実装となっています。
 * 
 * 作成日: 2026-01-21
 * 最終更新: 2026-01-22 トークン再取得ロジック追加
 */

// ====================================================================
// メイン実行関数 (トリガー設定用)
// ====================================================================

/**
 * マスタ情報同期のメイン実行関数
 * 
 * 全てのマスタ同期処理をシーケンシャル（順次）に実行します。
 * 各同期処理は独立した関数（別ファイル）として実装されており、本関数はそれらのオーケストレーションを行います。
 * 
 * 【認証トークンの取り扱いについて】
 * 各同期関数（syncShopMaster等）を実行すると、内部でAPIコールが行われ、
 * アクセストークンが新しく更新される可能性があります。
 * そのため、次の同期関数を呼ぶ前に `getAppConfig()` を実行し、
 * 必ず最新のトークンとリフレッシュトークンを取得して渡すようにしています。
 * これにより、連続実行時の認証エラー（Code: 002002）を回避しています。
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


  // 6.【仕入先マスタ同期】
  // 支払区分マスタ同期でトークンが更新された可能性があるため、再取得
  config = getAppConfig();
  token = config.accessToken;
  refreshToken = config.refreshToken;

  syncSupplierMaster(config, token, refreshToken);


  Logger.log('--- マスタ情報同期処理 完了 ---');
}