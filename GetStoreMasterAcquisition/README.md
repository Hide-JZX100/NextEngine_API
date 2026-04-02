# GetStoreMasterAcquisition (多機能マスタ一括同期基盤)

## 概要
ネクストエンジン内の各種マスタ情報を一括取得し、他の全プロジェクトの基盤となる
参照データ群を最新の状態に維持します。

## 開発背景と目的
当初は店舗マスタとモールマスタの取得を目的に開発されましたが、実装後のパフォーマンスが
非常に良好（24秒以内で完結）であったため、運用効率を最大化すべく、他の重要マスタの
取得機能も統合する形へ拡張しました。

## 主な機能・特徴
- **マルチマスタ同期**: 店舗・モールマスタに加え、受注キャンセル区分、支払い区分、
  仕入先マスタなど、業務に必要な情報を網羅的に取得。
- **超高速処理**: 複雑なマスタ群の取得をわずか約24秒で完了させる最適化されたロジック。
- **無人運用（スケジュール実行）**: 毎週月曜日午前1時台のトリガー設定により、
  週明けの業務開始時には常に最新のマスタが整っている環境を実現。

## 実務上のメリット
「必要になったらすぐに追加できる」という高い拡張性を備えた設計にしています。
今後、新しい分析軸や管理項目が増えた際も、最小限の工数で自動取得のラインナップに
加えることが可能です。

---

## ファイル構成
```
GetStoreMasterAcquisition/
│
├── 00_NextEngineAuth.gs        # 認証スクリプト（初回セットアップ時のみ使用）
├── 01_共通ライブラリ.gs          # 全スクリプト共通の基盤関数群
├── 02_マスタ情報同期.gs          # メインエントリポイント（トリガー設定先）
├── 03_テスト共通関数.gs          # APIテスト用ユーティリティ
│
├── 04_店舗マスタ取得.gs          # 店舗マスタ 動作確認用
├── 04_店舗マスタ同期.gs          # 店舗マスタ 本番同期処理
├── 04_店舗マスタ検索.txt         # ネクストエンジンAPIフィールド定義
│
├── 05_モールマスタ取得.gs         # モールマスタ 動作確認用
├── 05_モールマスタ同期.gs         # モールマスタ 本番同期処理
├── 05_モールマスタ検索.txt        # ネクストエンジンAPIフィールド定義
│
├── 06_受注キャンセル区分取得.gs    # 受注キャンセル区分 動作確認用
├── 06_受注キャンセル区分同期.gs    # 受注キャンセル区分 本番同期処理
├── 06_受注キャンセル区分検索.txt   # ネクストエンジンAPIフィールド定義
│
├── 07_支払区分取得.gs            # 支払区分 動作確認用
├── 07_支払区分同期.gs            # 支払区分 本番同期処理
├── 07_支払区分検索.txt           # ネクストエンジンAPIフィールド定義
│
├── 08_仕入先マスタ取得.gs         # 仕入先マスタ 動作確認用
├── 08_仕入先マスタ同期.gs         # 仕入先マスタ 本番同期処理
├── 08_仕入先マスタ検索.txt        # ネクストエンジンAPIフィールド定義
│
└── README.md
```

### ファイル命名規則
- 数字プレフィックスで処理順・レイヤーを表現
- 「〇〇取得.gs」: 開発・動作確認用（ログ出力のみ、スプレッドシートへの書き込みなし）
- 「〇〇同期.gs」: 本番用（スプレッドシートへの書き込みあり）
- 「〇〇検索.txt」: ネクストエンジンAPIの公式フィールド定義（参照用）

---

## スクリプトプロパティ

本プロジェクトの実行には、GASのスクリプトプロパティへの事前設定が必要です。
値はソースコードに直接記述せず、すべてプロパティ経由で管理します。

### 認証情報（初回認証後に自動保存されるもの）

| キー名 | 説明 |
|---|---|
| `NEXT_ENGINE_CLIENT_ID` | ネクストエンジンアプリのクライアントID |
| `NEXT_ENGINE_CLIENT_SECRET` | ネクストエンジンアプリのクライアントシークレット |
| `NEXT_ENGINE_REDIRECT_URI` | 認証後のリダイレクト先（本GASのデプロイURL） |
| `NEXT_ENGINE_ACCESS_TOKEN` | アクセストークン（認証後に自動保存・自動更新） |
| `NEXT_ENGINE_REFRESH_TOKEN` | リフレッシュトークン（認証後に自動保存・自動更新） |

### スプレッドシート設定（手動で設定するもの）

| キー名 | 説明 |
|---|---|
| `SPREADSHEET_ID` | 書き込み先スプレッドシートのID |
| `SHEET_NAME_SHOP` | 店舗マスタを書き込むシート名 |
| `SHEET_NAME_MALL` | モールマスタを書き込むシート名 |
| `SHEET_NAME_CANCEL` | 受注キャンセル区分を書き込むシート名 |
| `SHEET_NAME_PAYMENT` | 支払区分情報を書き込むシート名 |
| `SHEET_NAME_SUPPLIER` | 仕入先マスタを書き込むシート名 |

---

## 使用APIエンドポイント

| マスタ名 | エンドポイント | タイプ |
|---|---|---|
| 店舗マスタ | `/api_v1_master_shop/search` | search |
| モールマスタ | `/api_v1_master_shop/search` | search ※1 |
| 受注キャンセル区分 | `/api_v1_system_canceltype/info` | info |
| 支払区分 | `/api_v1_system_paymentmethod/info` | info |
| 仕入先マスタ | `/api_v1_master_supplier/search` | search |

> ※1 モールマスタは店舗マスタと同一エンドポイントを使用し、取得フィールドで区別します。
>
> **searchとinfoの違い**:
> - `search`: `fields` パラメータで取得項目を指定する必要があります。
> - `info`: `fields` パラメータ不要。システム定義の固定マスタを全件返します。

---

## 主要関数一覧

### 00_NextEngineAuth.gs（認証 / 初回セットアップ専用）

| 関数名 | 説明 |
|---|---|
| `generateAuthUrl()` | 認証用URLを生成してログに出力する。初回認証時に実行。 |
| `doGet(e)` | 認証リダイレクトを受け取るWebアプリ関数。GAS標準関数。 |
| `getAccessToken(uid, state)` | UIDとStateを使ってトークンを取得・保存する。 |
| `testApiConnection()` | ログインユーザー情報取得でトークンの有効性を確認する。 |
| `callNextEngineApi(path, params)` | 汎用APIリクエスト関数。トークン自動更新機能付き。 |

### 01_共通ライブラリ.gs（全スクリプトの基盤）

| 関数名 | 説明 |
|---|---|
| `getProperty(key)` | スクリプトプロパティから設定値を取得するヘルパー。 |
| `saveTokens(accessToken, refreshToken)` | トークンをスクリプトプロパティに保存する。 |
| `getAppConfig()` | 全設定値を一括取得・バリデーションして返す。 |
| `nextEngineApiSearch(endpoint, fields, accessToken, refreshToken)` | ネクストエンジンAPIを実行する中核関数。トークン自動更新対応。 |
| `jsonToSheetArray(data, headerMap)` | APIレスポンスをスプレッドシート用2次元配列に変換する。 |
| `writeToSheet(sheetArray, sheetName, spreadsheetId, textColumnIndices)` | 指定シートにデータを書き込む。シート未存在時は自動作成。 |

### 02_マスタ情報同期.gs（メインエントリポイント）

| 関数名 | 説明 |
|---|---|
| `mainMasterSync()` | 全マスタ同期をシーケンシャルに実行する。トリガー設定対象。 |

### 03_テスト共通関数.gs（開発・動作確認用）

| 関数名 | 説明 |
|---|---|
| `testApiCall(endpoint, fields, logPrefix)` | APIを呼び出して結果をログ出力する共通テスト関数。 |

### 各マスタ同期関数（04〜08）

| 関数名 | 説明 |
|---|---|
| `fetchAndLogShopMaster()` | 店舗マスタをAPIから取得してログ出力（動作確認用）。 |
| `syncShopMaster(config, token, refreshToken)` | 店舗マスタを取得してスプレッドシートに書き込む。 |
| `fetchAndLogMallMaster()` | モールマスタをAPIから取得してログ出力（動作確認用）。 |
| `syncMallMaster(config, token, refreshToken)` | モールマスタを取得してスプレッドシートに書き込む。 |
| `fetchAndLogCancelTypeMaster()` | 受注キャンセル区分をAPIから取得してログ出力（動作確認用）。 |
| `syncCancelTypeMaster(config, token, refreshToken)` | 受注キャンセル区分を取得してスプレッドシートに書き込む。 |
| `fetchAndLogPaymentMethodMaster()` | 支払区分をAPIから取得してログ出力（動作確認用）。 |
| `syncPaymentMethodMaster(config, token, refreshToken)` | 支払区分を取得してスプレッドシートに書き込む。 |
| `fetchAndLogSupplierMaster()` | 仕入先マスタをAPIから取得してログ出力（動作確認用）。 |
| `syncSupplierMaster(config, token, refreshToken)` | 仕入先マスタを取得してスプレッドシートに書き込む。 |

---

## 初回セットアップ手順

1. GASプロジェクトをWebアプリとしてデプロイし、URLを控える。
2. スクリプトプロパティに `NEXT_ENGINE_CLIENT_ID`、`NEXT_ENGINE_CLIENT_SECRET`、
   `NEXT_ENGINE_REDIRECT_URI`（デプロイURL）を設定する。
3. `generateAuthUrl()` を実行し、ログに表示されたURLをブラウザで開いてログイン認証を行う。
4. 認証後、`NEXT_ENGINE_ACCESS_TOKEN` と `NEXT_ENGINE_REFRESH_TOKEN` が
   スクリプトプロパティに自動保存されることを確認する。
5. スプレッドシート関連のプロパティ（`SPREADSHEET_ID` 等）を設定する。
6. 各「取得.gs」の動作確認関数（`fetchAndLogShopMaster` 等）を個別に実行して、
   ログにデータが出力されることを確認する。
7. `mainMasterSync()` を手動実行し、全マスタがスプレッドシートに書き込まれることを確認する。

---

## トリガー設定

| トリガー対象関数 | 実行タイミング | 備考 |
|---|---|---|
| `mainMasterSync` | 毎週月曜日 午前1時台 | 週次バッチ。週明けの業務開始時に最新マスタを準備する。 |

---

## トークン管理の仕様

ネクストエンジンAPIはリクエスト毎にトークンが更新される仕様です。
本プロジェクトでは以下の方針でトークンを管理しています。

- APIレスポンスに新しいトークンが含まれていた場合、`saveTokens()` で即時上書き保存します。
- `mainMasterSync()` 内では、各マスタ同期関数の呼び出し前に `getAppConfig()` で
  最新トークンを再取得することで、連続実行時の認証エラーを防止しています。