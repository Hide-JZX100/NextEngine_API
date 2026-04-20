/**
 * スプレッドシート出力処理ファイル (GetShipmentCompletionSlip)
 * 
 * 【目的】
 * 加工済みの出荷伝票データをGoogleスプレッドシートに書き込む機能を提供するファイルです。
 * シートの存在確認、自動作成、ヘッダー追加、データの追記といった一連の出力処理を担当します。
 * 
 * 【処理内容】
 * 1. スクリプトプロパティから出力先のIDとシート名を取得
 * 2. シートが存在しない場合は新規作成し、ヘッダー行を設定
 * 3. 2次元配列形式のデータを指定シートの末尾に追記
 * 4. ヘッダー行には視認性を高めるための書式設定（背景色、太字など）を適用
 * 
 * 【含まれる関数】
 * - writeToSpreadsheet(data): データをシートに書き込むメイン関数
 * - appendHeader(sheet): ヘッダー行を作成・装飾する内部ヘルパー関数
 * - testWriteToSheet(): 書き込み動作を確認するための開発用テスト関数
 * 
 * 【依存関係】
 * - config.gs: シートID、シート名、フィールド定義(ヘッダー名)を参照
 */

/**
 * データをスプレッドシートに書き込む（Sheets API版）
 * 指定されたタブが存在しない場合は新規作成する
 * 
 * @param {Array<Array>} data - 書き込むデータの2次元配列
 */
function writeToSpreadsheet(data) {
    if (!data || data.length === 0) {
        console.log('書き込むデータがありません。');
        return;
    }

    console.log(`=== スプレッドシート書込開始 (Sheets API): ${data.length} 件 ===`);

    try {
        const props = PropertiesService.getScriptProperties();
        const sheetId = props.getProperty(CONFIG.SHEET.PROPERTY_KEY_ID);
        const sheetName = props.getProperty(CONFIG.SHEET.PROPERTY_KEY_NAME);

        if (!sheetId) throw new Error('スクリプトプロパティにスプレッドシートID (SPREADSHEET_ID) が設定されていません。');
        if (!sheetName) throw new Error('スクリプトプロパティにシート名 (SHEET_NAME) が設定されていません。');

        // シートの存在確認とヘッダー確認（SpreadsheetAppを使用）
        const ss = SpreadsheetApp.openById(sheetId);
        let sheet = ss.getSheetByName(sheetName);

        if (!sheet) {
            console.log(`シート '${sheetName}' が存在しないため作成します。`);
            sheet = ss.insertSheet(sheetName);
            appendHeader(sheet);
        } else if (sheet.getLastRow() === 0) {
            appendHeader(sheet);
        }

        // Sheets API でデータを追記
        const range = `${sheetName}!A:A`; // A列を基準に最終行を自動判定

        const resource = {
            values: data
        };

        const options = {
            valueInputOption: 'USER_ENTERED'  // スプレッドシートが自動判定
        };

        // values.append を使用して最終行の次に追記
        Sheets.Spreadsheets.Values.append(resource, sheetId, range, options);

        console.log(`書込完了: ${data.length} 行追記しました (Sheets API)`);

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
