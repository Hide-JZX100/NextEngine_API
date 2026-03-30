# GetScheduledShipmentQuantity (出荷・集荷計画 最適化プロジェクト)

## 概要
ネクストエンジンの受注情報から、当日および翌日以降の出荷予定数を自動集計し、社内外へ共有するための基盤を構築します。

## 開発背景と目的
かつて、出荷予定数の把握に時間がかかり、現場への共有が遅れるという課題がありました。「データが存在する以上、即座に判明するはず」という信念のもと、SFTPサーバーを用いた連携から着手し、現在はネクストエンジンAPIを用いた「フルクラウド連携」へと昇華させました。

## 主な機能・特徴
- **フルクラウド化による即時性**: SFTPサーバー等の物理インフラを排し、GASとAPIによるダイレクトなデータ連携を実現。
- **最適化された更新サイクル**: 月間APIコール数制限（無料枠）を考慮しつつ、実務に最適な「1日5回」の自動更新をトリガー設定。
- **社外連携への貢献**: 集計されたデータは社内だけでなく運送会社へも共有され、集荷車両の配備計画など、物流全体の最適化に寄与しています。

## 実務上のメリット
「人がダウンロード作業を行わない」ことを前提とした設計により、現場の待ち時間をゼロにしました。ITの力で物流のリードタイムを短縮し、取引先（運送会社）との信頼関係構築にも貢献しているプロジェクトです。

---

## ファイル構成

| ファイル名 | 役割 |
|---|---|
| `00_認証.gs` | ネクストエンジンOAuth2認証・トークン取得・保存 |
| `10_本番用_出荷予定データ更新.gs` | メイン処理統括・リトライ・スプレッドシート書き込み・外部ライブラリ呼び出し |
| `11_出荷明細取得.gs` | APIページング処理による全データ取得・ウォームアップ用取得 |
| `12_データ変換.gs` | APIレスポンス → スプレッドシート14列形式への変換 |
| `13_スプレッドシート書き込み.gs` | ヘッダー作成・データ書き込みテスト関数群 |
| `14_トリガー設定スクリプト.gs` | 時間ベーストリガーの作成・削除・スケジュール管理 |
| `15_ウォームアップ.gs` | 本番実行前のAPIキャッシュ暖機処理 |

### 外部ライブラリ（GASライブラリとして参照）
| ライブラリ名 | 説明 |
|---|---|
| `Yamato` | ヤマト運輸向け出荷予定数転記。集計済みデータを指定スプレッドシートへ書き込む。 |
| `Sagawa` | 佐川急便向け出荷予定数転記。集計済みデータを指定スプレッドシートへ書き込む。 |
| `JP`     | 日本郵便向け出荷予定数転記。集計済みデータを指定スプレッドシートへ書き込む。 |

> これらのライブラリはそれぞれ独立したGASプロジェクトとして管理されており、
> 本プロジェクトから `ライブラリID` で参照します。
> 各ライブラリのセットアップはそれぞれのREADMEを参照してください。

---

## 使用APIエンドポイント

| エンドポイント | 用途 |
|---|---|
| `POST /api_neauth` | アクセストークン取得 |
| `POST /api_v1_login_user/info` | 認証テスト・トークン更新確認 |
| `POST /api_v1_receiveorder_row/search` | 受注明細検索（出荷予定データ取得） |

**ホスト**
- 認証: `https://base.next-engine.org`
- API: `https://api.next-engine.org`

---

## 取得フィールド一覧

1件のAPIレスポンスからスプレッドシートの以下14列を生成します。

| 列 | 項目名 | フィールド名 | データ型 |
|---|---|---|---|
| 1 | 出荷予定日 | `receive_order_send_plan_date` | 日時型 |
| 2 | 伝票番号 | `receive_order_row_receive_order_id` | 数値型 |
| 3 | 商品コード | `receive_order_row_goods_id` | 文字列型 |
| 4 | 商品名 | `receive_order_row_goods_name` | 文字列型 |
| 5 | 受注数 | `receive_order_row_quantity` | 数値型 |
| 6 | 引当数 | `receive_order_row_stock_allocation_quantity` | 数値型 |
| 7 | 奥行き（cm） | `goods_length` | 数値型 |
| 8 | 幅（cm） | `goods_width` | 数値型 |
| 9 | 高さ（cm） | `goods_height` | 数値型 |
| 10 | 発送方法コード | `receive_order_delivery_id` | 文字列型 |
| 11 | 発送方法 | `receive_order_delivery_name` | 文字列型 |
| 12 | 重さ（g） | `goods_weight` | 数値型 |
| 13 | 受注状態区分 | `receive_order_order_status_id` | 文字列型 |
| 14 | 送り先住所1 | `receive_order_consignee_address1` | 文字列型 |

> **注意**: キャンセルフラグが立っている明細行（`receive_order_row_cancel_flag != '0'`）はAPIクエリ段階で除外されます。

---

## スクリプトプロパティ設定

GASエディタの「プロジェクトの設定」→「スクリプトプロパティ」に以下を設定してください。

### 認証情報

| キー | 説明 |
|---|---|
| `CLIENT_ID` | ネクストエンジンのクライアントID |
| `CLIENT_SECRET` | ネクストエンジンのクライアントシークレット |
| `REDIRECT_URI` | GAS WebアプリのデプロイURL |

### トークン（認証後に自動保存）

| キー | 説明 |
|---|---|
| `ACCESS_TOKEN` | アクセストークン（認証後に自動設定） |
| `REFRESH_TOKEN` | リフレッシュトークン（認証後に自動設定） |

### スプレッドシート設定

| キー | 説明 |
|---|---|
| `SPREADSHEET_ID` | 出荷予定データを書き込むスプレッドシートのID |
| `SHEET_NAME` | 出荷予定データを書き込むシート名 |
| `LOG_SHEET_NAME` | エラーログを記録するシート名 |
| `OPERATION_LOG_SHEET_NAME` | 実行完了日時を記録するシート名（A1セルに記録） |

### トリガー設定

| キー | 説明 |
|---|---|
| `TRIGGER_FUNCTION_NAME` | トリガーで実行する関数名（例: `updateShippingData`） |
| `TRIGGER_MODE` | `TODAY`（当日設定）または `TOMORROW`（翌日設定） |

---

## トリガースケジュール

| 時刻 | 実行関数 | 用途 |
|---|---|---|
| 7:35 | `warmUpNextEngineConnection` | APIキャッシュ暖機 |
| 7:56 | `updateShippingData` | 出荷予定データ更新 |
| 8:57 | `updateShippingData` | 出荷予定データ更新 |
| 9:57 | `updateShippingData` | 出荷予定データ更新 |
| 12:27 | `updateShippingData` | 出荷予定データ更新 |
| 13:27 | `updateShippingData` | 出荷予定データ更新 |

> トリガーは `14_トリガー設定スクリプト.gs` の `setTrigger()` を手動実行することで再設定されます。

---

## セットアップ手順

1. スクリプトプロパティに認証情報（`CLIENT_ID`, `CLIENT_SECRET`, `REDIRECT_URI`）を設定する
2. GASスクリプトをWebアプリとしてデプロイし、そのURLを `REDIRECT_URI` に設定する
3. `00_認証.gs` の `generateAuthUrl()` を実行し、表示されたURLをブラウザで開いて認証する
4. 認証完了後、`testApiConnection()` でAPI接続を確認する
5. スプレッドシート関連のスクリプトプロパティ（`SPREADSHEET_ID` 等）を設定する
6. `TRIGGER_MODE` を `TOMORROW` に設定し、`setTrigger()` を実行してトリガーを登録する

---

## 処理フロー

```
setTrigger()
    │
    ├─ 7:35  warmUpNextEngineConnection()   # APIキャッシュ暖機
    │
    └─ 7:56〜 updateShippingData()          # メイン処理
            │
            ├─ fetchAllShippingDataWithRetry()   # APIページング取得（最大3リトライ）
            ├─ convertAllDataToSheetRows()       # 14列形式に変換
            ├─ writeDataToSheet()                # スプレッドシート書き込み
            ├─ recordExecutionTimestamp()        # 実行完了日時記録
            └─ Yamato/Sagawa/JP .Shipping()     # 外部ライブラリへ転記
```