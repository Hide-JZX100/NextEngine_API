/**
 * 日付関連のユーティリティ関数
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
        console.log('【開発モード】対象日(固定): ' CONFIG.DEV_MODE.TARGET_DATE);
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
