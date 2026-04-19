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
 *   - 本番モード: 実行日の前月
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
 * 検索対象の日付範囲を取得する
 * 
 * @param {number} batchNumber - バッチ番号（1, 2, 3）
 * @return {Object} {startDate: 'YYYY-MM-DD', endDate: 'YYYY-MM-DD'}
 */
function getTargetDateRange(batchNumber) {
    if (!batchNumber || batchNumber < 1 || batchNumber > CONFIG.BATCH.RANGES.length) {
        throw new Error(`無効なバッチ番号: ${batchNumber}`);
    }

    const range = CONFIG.BATCH.RANGES[batchNumber - 1];
    let baseDate;

    if (CONFIG.DEV_MODE.ENABLED) {
        // 開発モード: 固定日付の月を使用
        baseDate = new Date(CONFIG.DEV_MODE.TARGET_DATE);
        console.log('【開発モード】基準日: ' + CONFIG.DEV_MODE.TARGET_DATE);
    } else {
        // 本番モード: 前月
        const today = new Date();
        baseDate = new Date(today.getFullYear(), today.getMonth() - 1, 1);
        console.log('【本番モード】対象月: 前月');
    }

    // 前月の年月を取得
    const year = baseDate.getFullYear();
    const month = baseDate.getMonth();

    // 開始日
    const startDate = new Date(year, month, range.start);

    // 終了日（末日の場合は実際の末日を計算）
    let endDay = range.end;
    const lastDayOfMonth = new Date(year, month + 1, 0).getDate();
    if (endDay > lastDayOfMonth) {
        endDay = lastDayOfMonth;
    }
    const endDate = new Date(year, month, endDay);

    const result = {
        startDate: formatDate(startDate),
        endDate: formatDate(endDate)
    };

    console.log(`バッチ${batchNumber}: ${result.startDate} ～ ${result.endDate}`);
    return result;
}

/**
 * 前月の年月を取得（YYYY-MM形式）
 * 
 * @return {string} 'YYYY-MM' 形式の文字列
 */
function getPreviousMonth() {
    const today = new Date();
    const year = today.getFullYear();
    const month = today.getMonth(); // 0-11

    if (month === 0) {
        // 1月の場合は前年の12月
        return `${year - 1}-12`;
    } else {
        const m = ('00' + month).slice(-2);
        return `${year}-${m}`;
    }
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

/**
 * ウォームアップ: API接続を事前に確立してコールドスタートを回避
 * メイン処理と同じエンドポイントに1件だけリクエストを送る
 * 
 * @param {string} startDate - 検索開始日
 */
function warmupApiConnection(startDate) {
    console.log('=== ウォームアップ開始 ===');

    try {
        const props = PropertiesService.getScriptProperties();
        const accessToken = props.getProperty('ACCESS_TOKEN');
        const refreshToken = props.getProperty('REFRESH_TOKEN');

        if (!accessToken || !refreshToken) {
            console.log('トークンなし。ウォームアップをスキップ');
            return;
        }

        // メイン処理と同じエンドポイントに最小限のリクエスト
        const params = {
            'access_token': accessToken,
            'refresh_token': refreshToken,
            'wait_flag': '1',
            'fields': 'receive_order_id',  // 最小限のフィールド
            'receive_order_send_plan_date-gte': startDate,
            'limit': '1',  // 1件のみ
            'offset': '0'
        };

        const url = CONFIG.API.BASE_URL + CONFIG.API.ENDPOINT_SEARCH;
        const options = {
            'method': 'post',
            'payload': params,
            'muteHttpExceptions': true
        };

        const response = UrlFetchApp.fetch(url, options);
        const json = JSON.parse(response.getContentText());

        // トークン更新
        if (json.access_token && json.refresh_token) {
            props.setProperties({
                'ACCESS_TOKEN': json.access_token,
                'REFRESH_TOKEN': json.refresh_token,
                'TOKEN_UPDATED_AT': new Date().getTime().toString()
            });
        }

        console.log('✅ ウォームアップ完了');

    } catch (error) {
        console.log('ウォームアップエラー(処理は続行):', error.message);
    }
}