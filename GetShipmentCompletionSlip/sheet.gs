/**
 * スプレッドシートへの出力処理を行う
 */

/**
 * データをスプレッドシートに書き込む
 * 指定されたタブが存在しない場合は新規作成する
 * 
 * @param {Array<Array>} data - 書き込むデータの2次元配列
 */
function writeToSpreadsheet(data) {
    if (!data || data.length === 0) {
        console.log('書き込むデータがありません。');
        return;
    }

    console.log(`=== スプレッドシート書込開始: ${data.length} 件 ===`);

    try {
        const props = PropertiesService.getScriptProperties();
        const sheetId = props.getProperty(CONFIG.SHEET.PROPERTY_KEY_ID);
        const sheetName = props.getProperty(CONFIG.SHEET.PROPERTY_KEY_NAME);

        if (!sheetId) throw new Error('スクリプトプロパティにスプレッドシートID (SPREADSHEET_ID) が設定されていません。');
        if (!sheetName) throw new Error('スクリプトプロパティにシート名 (SHEET_NAME) が設定されていません。');

        const ss = SpreadsheetApp.openById(sheetId);
        let sheet = ss.getSheetByName(sheetName);

        // シートが存在しない場合は作成
        if (!sheet) {
            console.log(`シート '${sheetName}' が存在しないため作成します。`);
            sheet = ss.insertSheet(sheetName);
            // ヘッダー行を追加 (もし空の場合)
            appendHeader(sheet);
        } else {
            // シートが存在しても、データが空(1行目がない)ならヘッダー追加
            if (sheet.getLastRow() === 0) {
                appendHeader(sheet);
            }
        }

        // データの書き込み (最終行の次に追加)
        const startRow = sheet.getLastRow() + 1;
        const numRows = data.length;
        const numCols = data[0].length;

        sheet.getRange(startRow, 1, numRows, numCols).setValues(data);

        console.log(`書込完了: ${numRows} 行追記しました。`);

    } catch (e) {
        console.error('スプレッドシート書込エラー:', e.message);
        throw e;
    }
}

function appendHeader(sheet) {
    const headers = CONFIG.FIELDS.map(field => field.header);
    const headerRange = sheet.getRange(1, 1, 1, headers.length);

    // ヘッダー値を設定
    headerRange.setValues([headers]);

    // ヘッダーのスタイル設定
    headerRange
        .setBackground('#e67e22')           // オレンジ系(温かみのある色)
        .setFontColor('#ffffff')            // 白文字
        .setFontWeight('bold')              // 太字
        .setHorizontalAlignment('center');  // 中央揃え

    console.log('ヘッダー行を追加しました。');
}

/**
 * 【開発用】書込テスト
 */
function testWriteToSheet() {
    // ダミーデータ
    const dummyData = [
        [
            '9999999', 'TEST-ORDER-001', '2025-10-03', '2025-10-01', 'テスト 太郎',
            'クレジットカード', '10000', '9000', '1000', '0',
            '0', '0', '0', '1234-5678-9012', '1',
            '', '', '', '', '0',
            '0', '', '', '50', '出荷確定済'
        ]
    ];

    try {
        writeToSpreadsheet(dummyData);
    } catch (e) {
        console.error('テスト失敗:', e.message);
    }
}
