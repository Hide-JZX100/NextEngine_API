/**
 * クリーンアップ処理
 * 古いトリガーとスプレッドシートデータを削除
 */

/**
 * 古いトリガーと古いデータをクリーンアップ
 * 月初のバッチ1実行前に呼び出される
 */
function cleanupOldTriggersAndData() {
    console.log('=== クリーンアップ開始 ===');

    // 1. 古いトリガーを削除
    deleteAllBatchTriggers();

    // 2. スプレッドシートのデータを削除
    clearSheetData();

    console.log('=== クリーンアップ完了 ===');
}

/**
 * すべてのバッチ関連トリガーを削除
 */
function deleteAllBatchTriggers() {
    const triggers = ScriptApp.getProjectTriggers();
    let deletedCount = 0;

    triggers.forEach(trigger => {
        const handlerFunction = trigger.getHandlerFunction();

        // executeBatch1, executeBatch2, executeBatch3 のトリガーを削除
        if (handlerFunction.startsWith('executeBatch')) {
            ScriptApp.deleteTrigger(trigger);
            deletedCount++;
            console.log(`トリガー削除: ${handlerFunction}`);
        }
    });

    console.log(`削除したトリガー数: ${deletedCount}`);
}

/**
 * 特定の関数名のトリガーを削除
 * 各バッチ実行開始時に自分自身のトリガーを削除するために使用
 * 
 * @param {string} functionName - 削除する関数名
 */
function deleteMyTrigger(functionName) {
    const triggers = ScriptApp.getProjectTriggers();

    triggers.forEach(trigger => {
        if (trigger.getHandlerFunction() === functionName) {
            ScriptApp.deleteTrigger(trigger);
            console.log(`自分のトリガーを削除: ${functionName}`);
        }
    });
}

/**
 * スプレッドシートのデータを削除（ヘッダー行は残す）
 */
function clearSheetData() {
    try {
        const props = PropertiesService.getScriptProperties();
        const sheetId = props.getProperty(CONFIG.SHEET.PROPERTY_KEY_ID);
        const sheetName = props.getProperty(CONFIG.SHEET.PROPERTY_KEY_NAME);

        if (!sheetId || !sheetName) {
            console.log('スプレッドシート設定が見つかりません。データ削除をスキップ');
            return;
        }

        const ss = SpreadsheetApp.openById(sheetId);
        const sheet = ss.getSheetByName(sheetName);

        if (!sheet) {
            console.log(`シート '${sheetName}' が存在しません。データ削除をスキップ`);
            return;
        }

        const lastRow = sheet.getLastRow();

        if (lastRow > 1) {
            // 2行目以降のデータをクリア（行削除ではなく内容クリア）
            const lastCol = sheet.getLastColumn();
            sheet.getRange(2, 1, lastRow - 1, lastCol).clearContent();
            console.log(`${lastRow - 1} 行のデータをクリアしました`);
        } else {
            console.log('削除するデータがありません');
        }

    } catch (error) {
        console.error('データ削除エラー:', error.message);
        // エラーでも処理は続行（致命的ではない）
    }
}

/**
 * 【開発用】クリーンアップのテスト
 */
function testCleanup() {
    cleanupOldTriggersAndData();
}