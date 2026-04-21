# ネクストエンジン Amazon受注データ取得 GAS開発指示書

## プロジェクト概要

Amazonからの受注データは出荷後30日で個人情報が消去される。
これを自動でGoogleスプレッドシートに転記するGoogle Apps Script（GAS）を開発する。
認証にはNEAuthライブラリ（既存）を使用する。

---

## 前提条件・重要注意事項

### 認証について（必読）

- 認証はGASライブラリ `NEAuth` を使用する（識別子: `NEAuth`）
- **APIコール時は必ず `access_token` と `refresh_token` の両方をペイロードに含めること**
  - `access_token` のみ送信するとAPIエラーになるバグが過去に発生している
- **APIコールのたびにレスポンスから新しいトークンを取得し、スクリプトプロパティへ保存すること**
  - ネクストエンジンはリクエストごとにトークンをローテーションする仕様
  - 保存しないと次回のAPIコールで認証エラーになる
- トークン情報はすべてスクリプトプロパティに保存する（コード内にハードコードしない）

### スクリプトプロパティ一覧

| キー | 説明 |
|---|---|
| `CLIENT_ID` | ネクストエンジンアプリのクライアントID |
| `CLIENT_SECRET` | ネクストエンジンアプリのクライアントシークレット |
| `REDIRECT_URI` | WebアプリのデプロイURL |
| `ACCESS_TOKEN` | APIアクセストークン（自動更新） |
| `REFRESH_TOKEN` | リフレッシュトークン（自動更新） |
| `TOKEN_OBTAINED_AT` | トークン取得日時（ミリ秒） |
| `TOKEN_UPDATED_AT` | トークン最終更新日時（ミリ秒） |
| `DEVELOPMENT_MODE` | `"true"` で開発モード、`"false"` または未設定で本番モード |

### ページネーションについて

- 1回のAPIコールで最大1000件取得可能（`limit: 1000`）
- レスポンスの `count` フィールドはページ内件数を返すのであって総件数ではない
- **ページネーション終了判定：取得件数 < limit のとき終了**（countによる判定は不可）
- ページ間のウェイト：500ms（GASのURLフェッチ帯域クォータ対策）

### 既存ファイルについて

- `00_NE_認証ライブラリ使用必須関数.gs` は既存ファイルのため **変更不要・変更禁止**
- 新規作成するファイルは `01_` 〜 `05_` のプレフィックスを付ける

---

## ファイル構成

```
00_NE_認証ライブラリ使用必須関数.gs  ← 既存。変更しない
01_設定ファイル.gs                  ← Phase 1で作成
02_API通信.gs                      ← Phase 2で作成
03_データ処理.gs                    ← Phase 3で作成
04_スプレッドシート書き込み.gs       ← Phase 3で作成
05_自動実行と通知.gs                ← Phase 4で作成
```

---

## 取得対象フィールド

エンドポイント: `https://api.next-engine.org/api_v1_receiveorder_base/search`

| ヘッダー名（シート表示名） | APIフィールド名 |
|---|---|
| 伝票番号 | `receive_order_id` |
| 受注日 | `receive_order_date` |
| 出荷予定日 | `receive_order_send_plan_date` |
| 購入者名 | `receive_order_purchaser_name` |
| 購入者郵便番号 | `receive_order_purchaser_zip_code` |
| 購入者住所1 | `receive_order_purchaser_address1` |
| 購入者住所2 | `receive_order_purchaser_address2` |
| 購入者電話番号 | `receive_order_purchaser_tel` |
| 購入者メールアドレス | `receive_order_purchaser_mail_address` |
| 送り先名 | `receive_order_consignee_name` |
| 送り先郵便番号 | `receive_order_consignee_zip_code` |
| 送り先住所1 | `receive_order_consignee_address1` |
| 送り先住所2 | `receive_order_consignee_address2` |
| 送り先電話番号 | `receive_order_consignee_tel` |

---

## Phase 1：`01_設定ファイル.gs` の作成

### 実装内容

以下の定数・設定をすべてこのファイルに集約する。
コード内にIDや数値をハードコードしてはならない。

```javascript
// ============================================================
// 01_設定ファイル.gs
// ネクストエンジン Amazon受注データ取得 - 定数・設定定義
// ============================================================

/** APIベースURL */
const NE_API_BASE_URL = 'https://api.next-engine.org';

/** 受注伝票検索エンドポイント */
const NE_ENDPOINT_ORDER_SEARCH = '/api_v1_receiveorder_base/search';

/** 受注状態区分：出荷確定済み */
const ORDER_STATUS_SHIPPED = '50';

/** 店舗ID */
const SHOP_ID_AMAZON     = '20';  // 本番：Amazon
const SHOP_ID_RAKUTEN    = '2';   // 開発：ベッド＆マットレス楽天市場店

/** 1回のAPIコールで取得する最大件数 */
const API_PAGE_LIMIT = 1000;

/** ページ間ウェイト（ミリ秒） */
const API_PAGE_WAIT_MS = 500;

/** 出力先スプレッドシートID */
const OUTPUT_SPREADSHEET_ID = '手動でスクリプトプロパティに保存';

/** 出力先シートタブ名 */
const OUTPUT_SHEET_NAME = 'API';

/**
 * 取得フィールド定義
 * api  : ネクストエンジンAPIのフィールド名
 * header: スプレッドシートのヘッダー表示名
 */
const CONFIG_FIELDS = [
  { api: 'receive_order_id',                  header: '伝票番号'             },
  { api: 'receive_order_date',                header: '受注日'               },
  { api: 'receive_order_send_plan_date',      header: '出荷予定日'           },
  { api: 'receive_order_purchaser_name',      header: '購入者名'             },
  { api: 'receive_order_purchaser_zip_code',  header: '購入者郵便番号'       },
  { api: 'receive_order_purchaser_address1',  header: '購入者住所1'          },
  { api: 'receive_order_purchaser_address2',  header: '購入者住所2'          },
  { api: 'receive_order_purchaser_tel',       header: '購入者電話番号'       },
  { api: 'receive_order_purchaser_mail_address', header: '購入者メールアドレス' },
  { api: 'receive_order_consignee_name',      header: '送り先名'             },
  { api: 'receive_order_consignee_zip_code',  header: '送り先郵便番号'       },
  { api: 'receive_order_consignee_address1',  header: '送り先住所1'          },
  { api: 'receive_order_consignee_address2',  header: '送り先住所2'          },
  { api: 'receive_order_consignee_tel',       header: '送り先電話番号'       },
];

/**
 * 開発モード判定
 * スクリプトプロパティ DEVELOPMENT_MODE が "true" の場合に開発モードとみなす
 * @returns {boolean} true: 開発モード / false: 本番モード
 */
function isDevelopmentMode() {
  const props = PropertiesService.getScriptProperties();
  return props.getProperty('DEVELOPMENT_MODE') === 'true';
}

/**
 * 現在の環境に応じた店舗IDを返す
 * @returns {string} 店舗ID
 */
function getShopId() {
  return isDevelopmentMode() ? SHOP_ID_RAKUTEN : SHOP_ID_AMAZON;
}
```

---

## Phase 2：`02_API通信.gs` の作成

### 実装内容

#### 関数1: `callNeApi(endpoint, payload)`

**役割：** ネクストエンジンAPIへの単一リクエストを実行する基底関数

**処理フロー：**
1. スクリプトプロパティから `ACCESS_TOKEN` と `REFRESH_TOKEN` を取得
2. ペイロードに `access_token` と `refresh_token` の両方を追加（必須）
3. `UrlFetchApp.fetch()` でPOSTリクエストを送信
4. レスポンスをJSONパース
5. `result === 'success'` 以外はエラーとしてthrow
6. **レスポンス内のトークン（`access_token`, `refresh_token`）をスクリプトプロパティへ保存**
7. レスポンスデータを返す

**引数：**
- `endpoint` {string} — APIエンドポイント（例: `/api_v1_receiveorder_base/search`）
- `payload` {Object} — リクエストパラメータ（トークン以外）

**戻り値：** APIレスポンスオブジェクト

---

#### 関数2: `fetchOrdersByShipDate(targetDateStr)`

**役割：** 指定した出荷確定日で出荷確定済み受注を全件取得する（ページネーション対応）

**引数：**
- `targetDateStr` {string} — 取得対象日（形式: `"YYYY-MM-DD"`）

**戻り値：** 受注データの配列（全ページ分）

**処理フロー：**
1. `offset = 0` で開始
2. ループ：
   a. `callNeApi()` を呼び出し（フィルタ: 店舗ID・受注状態・出荷確定日）
   b. 取得データを結果配列に追加
   c. **取得件数 < API_PAGE_LIMIT であればループ終了**
   d. `offset += API_PAGE_LIMIT`
   e. `Utilities.sleep(API_PAGE_WAIT_MS)` でウェイト
3. 全件を返す

**APIリクエストのペイロード例：**

```javascript
{
  'fields': CONFIG_FIELDS.map(f => f.api).join(','),
  'receive_order_shop_id-eq':            getShopId(),
  'receive_order_order_status_id-eq':    ORDER_STATUS_SHIPPED,
  'receive_order_send_date-eq':          targetDateStr,  // 出荷確定日
  'limit':                               API_PAGE_LIMIT,
  'offset':                              offset,
}
```

---

## Phase 3：`03_データ処理.gs` と `04_スプレッドシート書き込み.gs` の作成

### `03_データ処理.gs`

#### 関数: `formatOrderData(rawOrders)`

**役割：** APIレスポンス（オブジェクト配列）をシートへ書き込むための2次元配列に変換する

**引数：**
- `rawOrders` {Array} — `fetchOrdersByShipDate()` の戻り値

**戻り値：** `string[][]` — 2次元配列（行 × 列）

**処理：**
- `CONFIG_FIELDS` の `api` 順でフィールドを並べる
- 値が `null` / `undefined` の場合は空文字 `''` に変換する

---

### `04_スプレッドシート書き込み.gs`

#### 関数: `writeToSheet(data)`

**役割：** スプレッドシートの指定タブへヘッダーとデータを書き込む

**引数：**
- `data` {string[][]} — `formatOrderData()` の戻り値

**処理フロー：**
1. `OUTPUT_SPREADSHEET_ID` と `OUTPUT_SHEET_NAME` でシートを取得
2. シートの既存データをすべてクリア（`clearContents()`）
3. 1行目にヘッダーを書き込む（`CONFIG_FIELDS` の `header` を使用）
4. 2行目以降にデータを書き込む（`setValues()` で一括書き込み）

**ヘッダー行のスタイル：**
- 背景色: `#e67e22`（オレンジ）
- 文字色: `#ffffff`（白）
- 太字: `true`
- 水平位置: 中央揃え

---

## Phase 4：`05_自動実行と通知.gs` の作成

### 実装する関数一覧

#### 関数1: `dailyRun()`

**役割：** 毎日トリガーから自動実行される本番用関数

**処理：**
1. 実行前日の日付を `YYYY-MM-DD` 形式で取得
2. `fetchOrdersByShipDate(昨日の日付)` でデータ取得
3. `formatOrderData()` で整形
4. `writeToSheet()` で書き込み
5. 実行結果（取得件数・日付）をコンソールログに出力

**トリガー設定（手動で設定）：**
- 関数: `dailyRun`
- イベントのソース: 時間主導型
- タイプ: 日付ベースのタイマー
- 推奨時刻: 午前1時〜2時（JST）

---

#### 関数2: `manualRun(dateStr)`

**役割：** 任意の出荷確定日を指定して手動実行する

**引数：**
- `dateStr` {string} — 対象日付（形式: `"YYYY-MM-DD"`）省略時は前日

**使用例：**
```javascript
manualRun('2025-01-15');  // 2025年1月15日分を取得
manualRun();              // 省略時は前日
```

---

#### 関数3: `testRun()`

**役割：** 開発環境（楽天市場店 店舗ID:2）で動作確認を行うテスト用関数

**処理：**
1. スクリプトプロパティ `DEVELOPMENT_MODE` を一時的に `"true"` に設定
2. 任意の出荷予定日（`receive_order_send_plan_date`）でデータ取得
   - テスト対象日付はこの関数内の定数 `TEST_DATE` で指定する
3. 取得件数とデータ先頭1件の内容をコンソールログに出力（書き込みは行わない）
4. `DEVELOPMENT_MODE` を元に戻す（後処理）

**注意：** `testRun()` はシートへの書き込みを行わない（ログ確認のみ）

---

## コーディング規約

### ヘッダーコメントの形式

各関数には必ず以下の形式でJSDocコメントを記述すること。

```javascript
/**
 * 関数の役割を1行で記述
 *
 * 【処理フロー】
 * 1. ステップ1
 * 2. ステップ2
 *
 * 【使用タイミング】
 * - いつ呼び出されるか
 *
 * @param {型} 引数名 - 説明
 * @returns {型} 戻り値の説明
 */
```

### その他の規約

- `const` を優先し `var` は使用しない
- マジックナンバーは使用せず、必ず `01_設定ファイル.gs` の定数を参照する
- スプレッドシートIDやエンドポイントをコード内にハードコードしない
- ページネーションの終了判定は **取得件数 < limit** で行う（countフィールドは使用しない）
- エラーは `throw new Error('メッセージ')` でスローし、呼び出し元でcatchする

---

## 開発の進め方（フェーズ順）

```
Phase 1: 01_設定ファイル.gs を作成・動作確認
    ↓
Phase 2: 02_API通信.gs を作成し、APIコールとページネーションを確認
    ↓
Phase 3: 03_データ処理.gs と 04_スプレッドシート書き込み.gs を作成
         → スプレッドシートへの書き込みを確認
    ↓
Phase 4: 05_自動実行と通知.gs を作成
         → dailyRun / manualRun / testRun の動作確認
    ↓
Phase 5（将来）: エラー通知・実行ログ・リトライ処理
```

各Phaseは動作確認が取れてから次のPhaseに進む。
