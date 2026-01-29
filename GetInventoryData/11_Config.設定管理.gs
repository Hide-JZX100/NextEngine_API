/**
 * =============================================================================
 * Config.gs - 設定・定数管理
 * =============================================================================

 =============================================================================
 ネクストエンジン在庫情報取得スクリプト（統合版・完成版）
 =============================================================================
 * 【改善内容】
 * 1. ログレベル設定機能（MINIMAL/SUMMARY/DETAILED）
 * 2. バッチ処理ログの最適化（最初3件+最後3件方式）
 * 3. エラー時の詳細ログ出力
 * 4. 処理速度の向上（ログ出力削減により5-9秒短縮）
 * 
 * 【バージョン】
 * v2.0 - ログ最適化版
 * 
 * 【主な変更点】
 * - ログ出力を約99%削減（3000件 → 約20件）
 * - エラー発生時のみ詳細情報を自動出力
 * - 本番運用/デバッグモードの切り替え可能
 * - 実行時間が平均20-25秒（従来版 28-32秒）
 =============================================================================

 * 【目的】
 * 商品コードを配列で渡し、一度のAPIコールで複数商品の在庫情報を効率的に取得

 * 【主な改善点】
 * 1. 商品マスタAPIで複数商品を一度に検索（最大1000件）←商品マスタAPIは使わず
      在庫マスタAPIだけに修正
 * 2. 在庫マスタAPIで複数商品の在庫を一度に取得（最大1000件）
      const MAX_ITEMS_PER_CALL = 1000; で定義
 * 3. バッチ処理による大幅な高速化
 * 4. APIコール数の削減によるレート制限回避
 * 5. 単一API版での性能比較テスト追加

 * 【注意事項】
 * - 認証スクリプトで事前にトークンを取得済みである必要があります
 * - 一度に処理できる商品数は最大1000件です
 * - 大量データの場合は自動的にバッチ分割します

 実行時間を任意のシートのA1セルに記載するようにしました。
 スクリプトプロパティの設定を行ってください。
 【スクリプトプロパティの設定方法】
 1. GASエディタで「プロジェクトの設定」を開く（歯車のアイコン）
 2. 「スクリプトプロパティ」セクションまでスクロール
 3. 「スクリプトプロパティの追加」をクリックし、以下のキーと値を設定
   キー                     | 値
   -------------------------|------------------------------------
    SPREADSHEET_ID          | 在庫情報を更新したいスプレッドシートのID
    SHEET_NAME              | 在庫情報を更新したいシート名
    LOG_SHEET_NAME          | 実行時間を記載したいシート名
 *

 主要な処理を実行する関数

 getSpreadsheetConfig()
 スクリプトプロパティからスプレッドシートのIDとシート名を取得します。
 設定がなければエラーを発生させ、スクリプトの実行に必要な情報が揃っているかを確認します。

 getStoredTokens()
 スクリプトプロパティに保存されているアクセストークンとリフレッシュトークンを取得します。
 API呼び出しのたびに認証情報を取得する手間を省くためのユーティリティ関数です。

*/

// API関連定数
const NE_API_URL = 'https://api.next-engine.org';

// スプレッドシート列定義
const COLUMNS = {
    GOODS_CODE: 0,        // A列: 商品コード(GAS Index: 1)
    GOODS_NAME: 1,        // B列: 商品名(GAS Index: 2)
    STOCK_QTY: 2,        // C列: 在庫数(GAS Index: 3)
    ALLOCATED_QTY: 3,    // D列: 引当数(GAS Index: 4)
    FREE_QTY: 4,         // E列: フリー在庫数(GAS Index: 5)
    RESERVE_QTY: 5,      // F列: 予約在庫数(GAS Index: 6)
    RESERVE_ALLOCATED_QTY: 6,  // G列: 予約引当数(GAS Index: 7)
    RESERVE_FREE_QTY: 7, // H列: 予約フリー在庫数(GAS Index: 8)
    DEFECTIVE_QTY: 8,    // I列: 不良在庫数(GAS Index: 9)
    ORDER_REMAINING_QTY: 9,    // J列: 発注残数(GAS Index: 10)
    SHORTAGE_QTY: 10,    // K列: 欠品数(GAS Index: 11)
    JAN_CODE: 11         // L列: JANコード(GAS Index: 12)
};

// 処理設定値
const MAX_ITEMS_PER_CALL = 1000;
const API_WAIT_TIME = 500;

// ログレベル設定
const LOG_LEVEL = {
    MINIMAL: 1,    // 最小限: 開始/終了/サマリーのみ（本番運用推奨）
    SUMMARY: 2,    // サマリー: バッチ集計 + 最初/最後3件（デフォルト）
    DETAILED: 3    // 詳細: 全商品コード出力（デバッグ用）
};

// リトライ設定
const RETRY_CONFIG = {
    MAX_RETRIES: 3,              // 最大リトライ回数
    ENABLE_RETRY: true,          // リトライ機能の有効/無効
    LOG_RETRY_STATS: true        // リトライ統計のログ出力
};

// ============================================================================
// ユーティリティ関数
// ============================================================================

/**
 * スプレッドシート設定を取得
 */
function getSpreadsheetConfig() {
    const properties = PropertiesService.getScriptProperties();
    const SPREADSHEET_ID = properties.getProperty('SPREADSHEET_ID');
    const SHEET_NAME = properties.getProperty('SHEET_NAME');

    if (!SPREADSHEET_ID || !SHEET_NAME) {
        throw new Error('スプレッドシート設定が不完全です。スクリプトプロパティにSPREADSHEET_IDとSHEET_NAMEを設定してください。');
    }

    return {
        SPREADSHEET_ID,
        SHEET_NAME
    };
}

/**
 * 保存されたトークンを取得
 * (認証.gsで保存されたものを使用)
 */
function getStoredTokens() {
    const properties = PropertiesService.getScriptProperties();
    const accessToken = properties.getProperty('ACCESS_TOKEN');
    const refreshToken = properties.getProperty('REFRESH_TOKEN');

    if (!accessToken || !refreshToken) {
        throw new Error('アクセストークンが見つかりません。先に認証を完了してください。');
    }

    return {
        accessToken,
        refreshToken
    };
}
