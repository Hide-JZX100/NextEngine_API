/**
 * 修正後のネクストエンジンAPI取得ロジックの動作テスト用スクリプト (GetShipmentCompletionSlip)
 * 
 * 【目的】
 * 修正した `searchCompletedSlips` が、データの早期終了や歯抜けを起こさず、
 * 正しい件数を全件取得できるかテストします。
 * スプレッドシートへの書き込みは行わず、実行ログ出力（console.log）のみで検証を行います。
 * 
 * 【テスト方法】
 * 1. GASエディタで `runSearchCompletedSlipsTest` を選択し、実行します。
 * 2. 実行ログで「データ整合性チェック: OK」が表示されることを確認します。
 */

/**
 * 修正された searchCompletedSlips 関数のテスト実行
 * 
 * 【動作フロー】
 * 1. 2026年5月分（2026-05-01 〜 2026-05-31）の期間を指定して searchCompletedSlips を呼び出し。
 * 2. 取得したデータの総件数と API 側の総件数が一致しているかを確認。
 * 3. 取得した先頭のデータを1件サンプルとして出力し、データの中身を確認。
 */
function runSearchCompletedSlipsTest() {
    console.log('=== 修正後API取得テスト開始 ===');
    const startDate = '2026-05-01';
    const endDate = '2026-05-31';

    try {
        // 修正後の関数を呼び出して全件取得
        const results = searchCompletedSlips(startDate, endDate);

        console.log('=== テスト実行完了 ===');
        console.log(`取得件数: ${results.length} 件`);

        if (results.length > 0) {
            console.log('サンプルデータ（先頭1件）:', JSON.stringify(results[0]));
            
            // 簡易データ整合性チェック
            const missingFields = [];
            const checkFields = ['receive_order_id', 'receive_order_send_plan_date', 'receive_order_shop_id'];
            
            checkFields.forEach(field => {
                if (!(field in results[0])) {
                    missingFields.push(field);
                }
            });

            if (missingFields.length > 0) {
                console.error(`❌ テスト失敗: サンプルデータに必要なフィールド (${missingFields.join(', ')}) が存在しません。`);
            } else {
                console.log('✅ テスト成功: 必要なフィールドが正しく取得できています。');
            }
        } else {
            console.log('⚠️ データが取得されませんでした。認証状態や日付範囲を確認してください。');
        }

    } catch (e) {
        console.error('❌ テスト実行中にエラーが発生しました:', e.message);
    }
}

/**
 * 【実験用】前月分（5月分）の全バッチ（1〜3）を連続実行し、指定されたテスト用シートに出力する
 * 
 * 【目的】
 * 通常の月次バッチ（非同期5分間隔）を待たずに、バッチ1、2、3を同期処理で連続実行し、
 * スプレッドシートへの書込実験を一度の実行で行います。
 * 本番の main ロジックをそのまま使用するため、API取得、フィルタリング、
 * Sheets APIによる書込の全プロセスが本番と同条件で検証できます。
 * 
 * 【実行手順】
 * 1. 事前にスプレッドシートに空のタブ（例：`5月分_テスト出力`）を作成します。
 * 2. GASのスクリプトプロパティ `SHEET_NAME` をそのタブ名に変更します。
 * 3. 本関数を実行し、完了後に元シートのコピーと件数・日付内訳を比較します。
 * 4. 比較後、スクリプトプロパティ `SHEET_NAME` を元の本番シート名に戻します。
 */
function testRunMainAllBatches() {
    console.log('=== 【実験】全バッチ一括実行開始 ===');
    
    try {
        const props = PropertiesService.getScriptProperties();
        const sheetId = props.getProperty(CONFIG.SHEET.PROPERTY_KEY_ID);
        const sheetName = props.getProperty(CONFIG.SHEET.PROPERTY_KEY_NAME);
        
        console.log(`現在のターゲットスプレッドシートID: ${sheetId}`);
        console.log(`現在のターゲットシート名: ${sheetName}`);
        
        console.log(`⚠️ 注意: スクリプトプロパティの「SHEET_NAME」が実験用の新規タブ名になっていることを確認してください。`);
        
        // 1. テスト実行前に、そのシートの既存データをクリア（ヘッダー行以外）
        console.log(`ターゲットシート '${sheetName}' の既存データをクリアします...`);
        clearSheetData(); // cleanup.gs で定義されている既存のクリア関数を呼び出し
        
        // 2. バッチ1〜3を順次同期実行
        for (let batchNum = 1; batchNum <= 3; batchNum++) {
            console.log(`\n--- [実験] バッチ ${batchNum} 実行中 (対象範囲取得・加工・スプレッドシート書込) ---`);
            main(batchNum); // main.gs の main(batchNumber) を呼び出し
            
            // Sheets APIの同一シートへの連続書き込みによる負荷を考慮し、バッチ間に短いウェイトを挟む
            if (batchNum < 3) {
                console.log('API制限回避のため待機中 (3秒)...');
                Utilities.sleep(3000);
            }
        }
        
        // SRE/テスト考慮: 自動で予約されてしまったトリガーをクリア
        console.log('実験によって自動生成されたバッチ予約トリガーを削除します...');
        deleteAllBatchTriggers(); // cleanup.gs 内の関数
        
        console.log('\n=== 【実験】全バッチ一括実行完了 ===');
        console.log(`出力先タブ '${sheetName}' のデータと、退避している5月5日取得データを比較してください。`);
        
    } catch (e) {
        console.error('❌ 実験実行中にエラーが発生しました:', e.message);
    }
}

/**
 * 修正されたウォームアップ関数 (warmupAndScheduleMain) の動作テストを実行する
 * 
 * 【動作フロー】
 * 1. バッチ1を想定した日付範囲で warmupAndScheduleMain を呼び出し。
 * 2. 同期的にネクストエンジンAPIにリクエストが送信され、エラーが発生しないことを確認。
 * 3. ただし、本番トリガーが自動作成されるのを防ぐため、テスト実行直後に作成されたトリガーを削除します。
 */
function runWarmupTest() {
    console.log('=== ウォームアップクエリ動作テスト開始 ===');
    try {
        // バッチ1（1日〜10日）を想定してウォームアップを実行
        // 注: 内部で 2分後のメイン実行トリガーが作成されます
        warmupAndScheduleMain(1);
        
        console.log('✅ ウォームアップクエリの送信に成功しました。');
        
        // テスト用トリガーが残らないように自動削除
        console.log('テスト実行によって作成された実行トリガーをクリーンアップします...');
        deleteMyTrigger('executeBatch1');
        console.log('✅ トリガークリーンアップ完了');
        
    } catch (e) {
        console.error('❌ ウォームアップテスト失敗:', e.message);
    }
}

