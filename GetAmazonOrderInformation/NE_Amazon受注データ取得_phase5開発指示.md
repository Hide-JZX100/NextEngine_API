# ネクストエンジン Amazon受注データ取得 Phase 5 開発指示書

## Phase 5 概要

Phase 4 までで構築した受注データ取得スクリプトに、SRE的視点での安定化機能を追加する。

### Phase 5 の子フェーズ構成

```
Phase 5-1: ログ基盤の構築
           → 06_ログと通知.gs（新規作成）
           → LOGシートへの書き込み関数を実装
    ↓
Phase 5-2: リトライ処理の実装
           → 02_API通信.gs（修正）
           → callNeApi() に指数バックオフ付きリトライを追加
    ↓
Phase 5-3: 全体統合
           → 06_ログと通知.gs（修正）にエラー通知メール関数を追加
           → 05_自動実行と通知.gs（修正）のdailyRun() / manualRun() に統合
```

各子フェーズは動作確認が取れてから次の子フェーズに進む。

---

## スクリプトプロパティの追加

Phase 5 で以下のスクリプトプロパティを追加する。
既存のプロパティは変更しない。

| キー | 説明 | デフォルト値 |
|---|---|---|
| `LOG_SHEET_NAME` | ログ出力先シートタブ名 | `LOG` |
| `NOTIFY_EMAIL` | エラー通知先メールアドレス | （必須・未設定時はエラー） |

> **注意:** `OUTPUT_SPREADSHEET_ID` は既存のものをログシートでも共用する。
> LOGタブは既存スプレッドシート（`OUTPUT_SPREADSHEET_ID`）内に追加する。

---

## Phase 5-1：ログ基盤の構築

### 新規作成ファイル：`06_ログと通知.gs`

---

### ログシートの仕様

**シート名:** スクリプトプロパティ `LOG_SHEET_NAME`（未設定時は `'LOG'`）

**ヘッダー行のスタイル（既存APIシートと統一）:**
- 背景色: `#e67e22`（オレンジ）
- 文字色: `#ffffff`（白）
- 太字: `true`
- 水平位置: 中央揃え

**カラム定義:**

| 列 | ヘッダー名 | 内容 |
|---|---|---|
| A | 実行日時 | `yyyy-MM-dd HH:mm:ss` 形式 |
| B | 対象日付 | `YYYY-MM-DD` 形式（取得対象の出荷確定日） |
| C | 実行関数名 | `dailyRun` / `manualRun` |
| D | 取得件数 | 数値（0件の場合も記録） |
| E | 実行結果 | `成功` / `0件` / `エラー` |
| F | エラー内容 | エラー時のみ記載。正常時は空文字 |

---

### 関数1: `getLogSheetName()`

**役割:** ログシートのタブ名をスクリプトプロパティから取得する

**処理フロー:**
1. スクリプトプロパティ `LOG_SHEET_NAME` を取得
2. 未設定の場合はデフォルト値 `'LOG'` を返す

**引数:** なし

**戻り値:** `{string}` ログシートのタブ名

---

### 関数2: `initLogSheet()`

**役割:** LOGシートが存在しない場合に新規作成し、ヘッダー行を書き込む

**処理フロー:**
1. `OUTPUT_SPREADSHEET_ID` のスプレッドシートを開く
2. `getLogSheetName()` でシート名を取得
3. 同名のシートが既に存在する場合は何もせず終了（冪等性を保つ）
4. 存在しない場合はシートを新規作成
5. ヘッダー行（実行日時・対象日付・実行関数名・取得件数・実行結果・エラー内容）を1行目に書き込む
6. ヘッダー行にスタイルを適用（オレンジ背景・白文字・太字・中央揃え）

**引数:** なし

**戻り値:** なし

---

### 関数3: `writeLog(params)`

**役割:** LOGシートに実行結果の1行を追記する

**処理フロー:**
1. `OUTPUT_SPREADSHEET_ID` のスプレッドシートを開く
2. `getLogSheetName()` でログシートを取得
3. LOGシートが存在しない場合は `initLogSheet()` を呼んで初期化する
4. 現在日時を `yyyy-MM-dd HH:mm:ss` 形式で取得
5. 引数の内容を1行の配列に組み立てる
6. シートの最終行の次の行に `appendRow()` で追記する

**引数:**

```javascript
/**
 * @param {Object} params - ログパラメータ
 * @param {string} params.targetDate   - 取得対象日付（"YYYY-MM-DD"）
 * @param {string} params.funcName     - 実行関数名（"dailyRun" など）
 * @param {number} params.count        - 取得件数
 * @param {string} params.status       - 実行結果（"成功" / "0件" / "エラー"）
 * @param {string} [params.errorMsg]   - エラー内容（省略可。正常時は空文字）
 */
```

**戻り値:** なし

---

### 関数4: `testPhase5_1()`

**役割:** Phase 5-1 の動作確認用テスト関数

**処理フロー:**
1. `initLogSheet()` を呼び出してLOGシートを初期化（既存なら何もしない）
2. 成功ログのダミーデータで `writeLog()` を呼び出す
3. 0件ログのダミーデータで `writeLog()` を呼び出す
4. エラーログのダミーデータで `writeLog()` を呼び出す
5. コンソールに完了メッセージを出力
6. スプレッドシートのLOGタブを目視確認するよう促す

**ダミーデータ例:**

```javascript
// 成功ログ
writeLog({ targetDate: '2026-04-22', funcName: 'dailyRun', count: 42, status: '成功', errorMsg: '' });

// 0件ログ
writeLog({ targetDate: '2026-04-21', funcName: 'dailyRun', count: 0,  status: '0件', errorMsg: '' });

// エラーログ
writeLog({ targetDate: '2026-04-20', funcName: 'dailyRun', count: 0,  status: 'エラー', errorMsg: 'APIリクエストエラー: {"result":"error","message":"token error"}' });
```

---

## Phase 5-2：リトライ処理の実装

### 修正ファイル：`02_API通信.gs`

---

### 追加する定数（`01_設定ファイル.gs` に追記）

```javascript
/** APIリトライ最大回数 */
const API_RETRY_MAX = 3;

/** APIリトライ初回ウェイト（ミリ秒）。指数バックオフで倍増する */
const API_RETRY_BASE_WAIT_MS = 1000;
```

---

### 修正内容：`callNeApi()` にリトライ処理を追加

**リトライの設計方針:**

- 最大リトライ回数: `API_RETRY_MAX`（定数）= 3回
- ウェイト: 指数バックオフ（1回目: 1秒 → 2回目: 2秒 → 3回目: 4秒）
  - 計算式: `API_RETRY_BASE_WAIT_MS * Math.pow(2, attempt)`
- **リトライ対象:** ネットワークエラー・タイムアウト・一時的なサーバーエラー
  - 具体的には `try-catch` で補足した例外全般（ただし下記を除く）
- **リトライ対象外（即時エラー）:** 認証エラー（`result: 'error'` かつ `message` に `token` を含む場合）
  - 認証エラーはリトライしても解決しないため即座にスローする
- 全リトライ失敗時は最後のエラーをスローする
- リトライ試行のたびにコンソールログに試行回数を出力する

**修正後の処理フロー（`callNeApi` 全体）:**

```
1. スクリプトプロパティから ACCESS_TOKEN と REFRESH_TOKEN を取得
2. attempt = 0 でリトライループ開始
3.   ペイロードにトークンを追加
4.   UrlFetchApp.fetch() でPOSTリクエストを送信
5.   レスポンスをJSONパース
6.   result !== 'success' の場合:
       - 認証エラー（tokenを含む）なら即座にスロー（リトライしない）
       - それ以外は例外をスロー → catch でリトライへ
7.   新しいトークンをスクリプトプロパティへ保存
8.   レスポンスデータを返す（成功）
9.   catch:
       - attempt < API_RETRY_MAX であれば:
           コンソールに「リトライ N/MAX回目」を出力
           指数バックオフでウェイト
           attempt++ してループ継続
       - attempt >= API_RETRY_MAX であれば:
           最後のエラーをスロー（リトライ上限到達）
```

---

### 関数: `testPhase5_2()`

**役割:** Phase 5-2 の動作確認用テスト関数

**処理フロー:**
1. 意図的に無効なエンドポイントや不正なペイロードを送ってエラーを発生させる
2. リトライが `API_RETRY_MAX` 回実行されることをコンソールログで確認する
3. 最終的にエラーがスローされることを確認する

> **注意:** このテストは実際にAPIコールを行うため、
> `ACCESS_TOKEN` と `REFRESH_TOKEN` が有効な状態で実行すること。

---

## Phase 5-3：全体統合

### 修正ファイル1：`06_ログと通知.gs`（追記）

---

### 追加関数: `sendErrorNotification(params)`

**役割:** エラー発生時に通知先メールアドレスへエラー通知メールを送信する

**処理フロー:**
1. スクリプトプロパティ `NOTIFY_EMAIL` から通知先メールアドレスを取得
2. 未設定の場合はコンソールにエラーを出力してメール送信をスキップ（例外はスローしない）
3. 件名・本文を組み立てて `MailApp.sendEmail()` で送信する

**引数:**

```javascript
/**
 * @param {Object} params - 通知パラメータ
 * @param {string} params.targetDate - 取得対象日付
 * @param {string} params.funcName   - エラーが発生した関数名
 * @param {string} params.errorMsg   - エラー内容
 */
```

**メール件名:** `[エラー] ネクストエンジン受注データ取得失敗`

**メール本文（構成）:**

```
受注データの自動取得中にエラーが発生しました。

■ 対象日付: {targetDate}
■ 実行関数: {funcName}
■ 発生時刻: {現在日時}
■ エラー内容:
{errorMsg}

【対処方法】
1. GASエディタのログを確認してください
2. トークンエラーの場合は testGenerateAuthUrl() を実行して再認証してください
3. ネクストエンジン側の一時的なエラーの場合は manualRun('YYYY-MM-DD') で手動再実行してください

このメールは自動送信されています。
```

**戻り値:** なし

---

### 修正ファイル2：`05_自動実行と通知.gs`

---

### 修正内容：`dailyRun()` に統合

**修正後の処理フロー:**

```
1. 実行前日の日付を取得
2. try:
     a. fetchOrdersByShipDate(targetDate) でデータ取得（リトライ処理は callNeApi 内で自動実行）
     b. 取得件数が 0 件の場合:
          - writeLog({ status: '0件', ... }) でログ記録
          - メール通知はしない
          - 処理を終了する（書き込みはしない）
     c. formatOrderData() で整形
     d. writeToSheet() でスプレッドシートへ書き込み
     e. writeLog({ status: '成功', count: 件数, ... }) でログ記録
3. catch(error):
     a. writeLog({ status: 'エラー', errorMsg: error.message, ... }) でログ記録
     b. sendErrorNotification({ errorMsg: error.message, ... }) でメール通知
     c. error を再スロー（GASの実行ログにエラーを残す）
```

---

### 修正内容：`manualRun()` に統合

**修正後の処理フロー（`dailyRun` と同じ構造）:**

```
1. 引数の日付を取得（省略時は前日）
2. try:
     a. fetchOrdersByShipDate(targetDate)
     b. 取得件数 0 件 → writeLog({ status: '0件', ... })
     c. 正常取得 → formatOrderData() → writeToSheet() → writeLog({ status: '成功', ... })
3. catch(error):
     a. writeLog({ status: 'エラー', ... })
     b. sendErrorNotification()
     c. error を再スロー
```

> **注意:** `manualRun()` はエラー時もメール通知する。
> 手動実行中のエラーも確実に把握できるようにするため。

---

### 関数: `testPhase5_3()`

**役割:** Phase 5-3 の動作確認用テスト関数（統合テスト）

**処理フロー:**
1. 開発モードを `true` に設定
2. `manualRun('テスト用日付')` を実行（楽天市場店のデータで統合動作を確認）
3. LOGシートにログが記録されていることを目視確認
4. 開発モードを元の状態に戻す

---

## コーディング規約（Phase 5 追加分）

### 既存規約の継続

- `const` を優先し `var` は使用しない
- 各関数には JSDoc 形式のヘッダーコメントを必ず記述する
- マジックナンバーは使用せず `01_設定ファイル.gs` の定数を参照する

### Phase 5 追加規約

- `writeLog()` はエラーが発生しても例外をスローしない（ログ失敗で本処理が止まらないようにする）
  - `writeLog()` 内部を `try-catch` で囲み、ログ書き込み失敗はコンソールエラーに留める
- `sendErrorNotification()` も同様に例外をスローしない
- リトライのウェイト・回数はすべて `01_設定ファイル.gs` の定数で管理する（コード内にハードコードしない）

---

## ファイル構成（Phase 5 完了後）

```
00_NE_認証ライブラリ使用必須関数.gs  ← 変更なし
01_設定ファイル.gs                  ← API_RETRY_MAX / API_RETRY_BASE_WAIT_MS を追記
02_API通信.gs                      ← callNeApi() にリトライ処理を追加
03_データ処理.gs                    ← 変更なし
04_スプレッドシート書き込み.gs       ← 変更なし
05_自動実行と通知.gs                ← dailyRun() / manualRun() にログ・通知を統合
06_ログと通知.gs                    ← 新規作成（ログ書き込み・エラー通知メール）
```

---

## 開発の進め方（子フェーズ順）

```
Phase 5-1: 06_ログと通知.gs を新規作成
           → testPhase5_1() でLOGシートへの書き込みを確認
    ↓
Phase 5-2: 01_設定ファイル.gs に定数追加
           02_API通信.gs の callNeApi() にリトライを追加
           → testPhase5_2() でリトライ動作をコンソールログで確認
    ↓
Phase 5-3: 06_ログと通知.gs に sendErrorNotification() を追加
           05_自動実行と通知.gs の dailyRun() / manualRun() を修正
           → testPhase5_3() で統合動作を確認
           → LOGシートへの記録とメール通知が正常に動作することを確認
```