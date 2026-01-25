/**
 * ユーティリティ関数ファイル (GetShipmentCompletionSlip)
 * 
 * 【目的】
 * プロジェクト全体で使用される共通の補助機能をまとめたファイルです。
 * 主に日付操作やフォーマット変換など、ビジネスロジックに依存しない汎用的な処理を提供します。
 * 環境設定（config.gs）に基づいた動作の切り替えもここで行われます。
 * 
 * 【機能】
 * - 検索対象日付の算出: 開発モードと本番モードで異なる日付を返却
 *   - 開発モード: `config.gs` で指定された固定日付
 *   - 本番モード: 実行日の前日（昨日）
 * - 日付フォーマット変換: DateオブジェクトをAPI仕様に合わせた文字列形式に変換
 * 
 * 【含まれる関数】
 * - getTargetDate(): 現在のモードに応じた検索対象日を返すメイン関数
 * - formatDate(date): Dateオブジェクトを 'YYYY-MM-DD' 形式に変換するヘルパー関数
 * 
 * 【依存関係】
 * - config.gs: 開発モードの設定(DEV_MODE)を参照
 */

/**
 * 検索対象の日付を取得する
 * CONFIG.IS_DEV_MODEによって挙動が変わる
 * 
 * @return {string} 'YYYY-MM-DD' 形式の日付文字列
 */
function getTargetDate() {
    let targetDate;

    if (CONFIG.DEV_MODE.ENABLED) {
        // 開発モード:固定日付
        targetDate = new Date(CONFIG.DEV_MODE.TARGET_DATE);
        console.log('【開発モード】対象日(固定): ' + CONFIG.DEV_MODE.TARGET_DATE);
    } else {
        // 本番モード：昨日
        const today = new Date();
        targetDate = new Date(today);
        targetDate.setDate(today.getDate() - 1); // 昨日
        console.log('【本番モード】対象日(昨日): ' + formatDate(targetDate));
    }

    return formatDate(targetDate);
}

/**
 * DateオブジェクトをYYYY-MM-DD形式の文字列に変換する
 * 
 * @param {Date} date - 変換するDateオブジェクト
 * @return {string} 'YYYY-MM-DD' 形式の文字列
 */
function formatDate(date) {
    const y = date.getFullYear();
    const m = ('00' + (date.getMonth() + 1)).slice(-2);
    const d = ('00' + date.getDate()).slice(-2);
    return `${y}-${m}-${d}`;
}
