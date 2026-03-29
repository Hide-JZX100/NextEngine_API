# GetInventoryData

## 概要
ネクストエンジンの商品マスタ・在庫マスタとGoogleスプレッドシートをAPI接続し、
商品情報・在庫情報の同期を自動化するGoogle Apps Scriptプロジェクトです。
SFTPサーバーや手動ダウンロードを使用しないフルクラウド構成で運用します。

## 主な機能

### 商品マスタ全件同期（1日1回 / 0:10実行）
`updateInventoryDataFromGoodsMaster`
- ネクストエンジン商品マスタAPI（`/api_v1_master_goods/search`）から全件取得
- ロケーションに `xxxxxx` を含む商品を除外し、空欄を含むそれ以外の商品を対象とする
- 商品コード・商品名・JANコード・在庫数など12項目をスプレッドシートに全件書き直し
- ページネーション対応（1,000件 × 4ページ、約3,200件を約5秒で処理）
- 実行完了後に翌日分のトリガーを自動登録（自己スケジューリング方式）

### 在庫情報リアルタイム更新（1日6回）
`updateInventoryDataBatchWithRetry`
- ネクストエンジン在庫マスタAPI（`/api_v1_master_stock/search`）から在庫情報を取得
- スプレッドシートのA列商品コードを元に在庫数・引当数など9項目を更新
- 1,000件バッチ処理による高速化（約3,200件を約18秒で処理）
- エクスポネンシャルバックオフによる自動リトライ機能（最大3回）
- リトライ統計・エラーログをスプレッドシートに自動記録

## 取得項目一覧

| 列 | 項目名 | フィールド名 | 更新元 |
|----|--------|-------------|--------|
| A | 商品コード | goods_id | 商品マスタ |
| B | 商品名 | goods_name | 商品マスタ |
| C | 在庫数 | stock_quantity | 商品マスタ / 在庫マスタ |
| D | 引当数 | stock_allocation_quantity | 商品マスタ / 在庫マスタ |
| E | フリー在庫数 | stock_free_quantity | 商品マスタ / 在庫マスタ |
| F | 予約在庫数 | stock_advance_order_quantity | 商品マスタ / 在庫マスタ |
| G | 予約引当数 | stock_advance_order_allocation_quantity | 商品マスタ / 在庫マスタ |
| H | 予約フリー在庫数 | stock_advance_order_free_quantity | 商品マスタ / 在庫マスタ |
| I | 不良在庫数 | stock_defective_quantity | 商品マスタ / 在庫マスタ |
| J | 発注残数 | stock_remaining_order_quantity | 商品マスタ / 在庫マスタ |
| K | 欠品数 | stock_out_quantity | 商品マスタ / 在庫マスタ |
| L | JANコード | goods_jan_code | 商品マスタ |

## ファイル構成

| ファイル | 役割 |
|---------|------|
| 00_認証.gs | ネクストエンジンOAuth2認証・トークン管理 |
| 10_Main.gs | エントリーポイント・処理オーケストレーション |
| 11_Config.gs | 定数・設定値管理 |
| 12_Logger.gs | ログ出力・リトライ統計管理 |
| 13_NextEngineAPI.gs | NE APIへのHTTPリクエスト・トークン更新 |
| 14_InventoryLogic.gs | データ取得・整形（ビジネスロジック） |
| 15_SpreadsheetRepository.gs | スプレッドシートへの書き込み・ログ記録 |
| トリガー設定.gs | 時間ベーストリガーの登録・削除 |
| 99_Tests.gs | 動作確認・診断ツール |

## スクリプトプロパティ設定

| キー | 値 |
|------|----|
| CLIENT_ID | ネクストエンジン クライアントID |
| CLIENT_SECRET | ネクストエンジン クライアントシークレット |
| REDIRECT_URI | リダイレクトURI |
| SPREADSHEET_ID | 対象スプレッドシートID |
| SHEET_NAME | 在庫データシート名 |
| LOG_SHEET_NAME | 実行タイムスタンプ記録シート名 |
| TRIGGER_FUNCTION_NAME | トリガー実行関数名 |
| TRIGGER_MODE | TODAY または TOMORROW |
| LOG_LEVEL | 1: MINIMAL / 2: SUMMARY / 3: DETAILED |
| TEST_SPREADSHEET_ID | テスト用スプレッドシートID |

## トリガー構成

| 時刻 | 関数 | 目的 |
|------|------|------|
| 0:10 | updateInventoryDataFromGoodsMaster | 商品マスタ全件同期 |
| 8:00 | updateInventoryDataBatchWithRetry | 在庫情報更新 |
| 10:00 | updateInventoryDataBatchWithRetry | 在庫情報更新 |
| 13:30 | updateInventoryDataBatchWithRetry | 在庫情報更新 |
| 16:00 | updateInventoryDataBatchWithRetry | 在庫情報更新 |
| 19:00 | updateInventoryDataBatchWithRetry | 在庫情報更新 |
| 21:00 | updateInventoryDataBatchWithRetry | 在庫情報更新 |