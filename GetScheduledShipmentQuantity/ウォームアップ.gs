/**
 * ネクストエンジンAPI接続のウォームアップ
 * * 【目的】
 * 本番のデータ取得処理の前に、軽量なリクエストを送り
 * ネクストエンジン側のサーバーやキャッシュを活性化（ウォームアップ）させる。
 * * 【動作】
 * - ログインユーザー情報を取得するだけの最も軽量なリクエストを送る
 * - トークンが更新された場合は、既存の仕組みと同様にスクリプトプロパティを更新する
 * - エラーが発生しても、本番処理を阻害しないようログ出力に留める
 */
function warmUpNextEngineConnection() {
    console.log("=== ウォームアップ開始 ===");

    try {
        const properties = PropertiesService.getScriptProperties();
        const accessToken = properties.getProperty('ACCESS_TOKEN');
        const refreshToken = properties.getProperty('REFRESH_TOKEN');

        if (!accessToken || !refreshToken) {
            console.warn("トークンが見つからないため、ウォームアップをスキップします。");
            return;
        }

        // 最も軽量な「ログインユーザー情報取得」を利用
        const url = `${NE_API_URL}/api_v1_login_user/info`;

        const payload = {
            'access_token': accessToken,
            'refresh_token': refreshToken
        };

        const options = {
            'method': 'POST',
            'headers': {
                'Content-Type': 'application/x-www-form-urlencoded'
            },
            'payload': Object.keys(payload).map(key =>
                encodeURIComponent(key) + '=' + encodeURIComponent(payload[key])
            ).join('&'),
            'muteHttpExceptions': true // エラー時も例外を投げずにレスポンスを受け取る
        };

        const startTime = new Date();
        const response = UrlFetchApp.fetch(url, options);
        const endTime = new Date();
        const duration = (endTime - startTime) / 1000;

        console.log(`応答時間: ${duration.toFixed(2)}秒`);

        const responseData = JSON.parse(response.getContentText());

        if (responseData.result === 'success') {
            console.log('ウォームアップ成功: サーバーは正常に応答しています');

            // トークンが更新された場合の保存処理（既存のロジックを踏襲）
            if (responseData.access_token && responseData.refresh_token) {
                properties.setProperties({
                    'ACCESS_TOKEN': responseData.access_token,
                    'REFRESH_TOKEN': responseData.refresh_token,
                    'TOKEN_UPDATED_AT': new Date().getTime().toString()
                });
                console.log('トークンを更新しました');
            }
        } else {
            console.warn('ウォームアップAPIからエラーが返されました:', responseData.message);
        }

    } catch (error) {
        console.error('ウォームアップ中に予期せぬエラーが発生しました:', error.message);
    }

    console.log("=== ウォームアップ終了 ===");
}