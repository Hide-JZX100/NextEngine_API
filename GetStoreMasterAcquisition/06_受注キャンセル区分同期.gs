/**
 * 受注キャンセル区分マスタ同期処理
 * 
 * 目的: 受注キャンセル区分マスタをネクストエンジンAPIから取得し、
 *       Googleスプレッドシートに書き込む
 * 
 * 呼び出し元: master_sync_main.gs の mainMasterSync() 関数
 * 作成日: 2026-01-19
 */

/**
 * 受注キャンセル区分マスタの同期処理を実行
 * 
 * @param {Object} config - 設定オブジェクト（getAppConfig()の戻り値）
 * @param {string} token - アクセストークン
 * @param {string} refreshToken - リフレッシュトークン
 */
function syncCancelTypeMaster(config, token, refreshToken) {
    Logger.log('--- 受注キャンセル区分マスタの処理開始 ---');

    // ヘッダーマップ定義（APIフィールド名 → 日本語項目名）
    const cancelTypeHeaderMap = {
        'cancel_type_id': '受注キャンセル区分ID',
        'cancel_type_name': '受注キャンセル名',
        'cancel_type_deleted_flag': '非表示フラグ'
    };

    // API呼び出し
    // ★ /infoエンドポイントはfieldsパラメータが不要なのでnullを渡す
    const cancelTypeData = nextEngineApiSearch(
        'api_v1_system_canceltype/info',
        null,  // fieldsは不要
        token,
        refreshToken
    );

    // データ取得成功時の処理
    if (cancelTypeData) {
        // JSONデータをスプレッドシート用の2次元配列に変換
        const cancelTypeArray = jsonToSheetArray(cancelTypeData, cancelTypeHeaderMap);

        // スプレッドシートに書き込み
        writeToSheet(cancelTypeArray, config.SHEET_NAME_CANCEL, config.SPREADSHEET_ID);

        Logger.log(`受注キャンセル区分マスタ: ${cancelTypeData.length}件の処理完了`);
    } else {
        Logger.log('受注キャンセル区分マスタ: データ取得に失敗しました');
    }

    Logger.log('--- 受注キャンセル区分マスタの処理完了 ---');
}

//単体テスト用関数
function testSyncCancelTypeMaster() {
    const config = getAppConfig();
    syncCancelTypeMaster(config, config.accessToken, config.refreshToken);
}