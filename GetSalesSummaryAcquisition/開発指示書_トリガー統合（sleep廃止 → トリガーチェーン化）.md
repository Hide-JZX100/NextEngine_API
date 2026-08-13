# 開発指示書：07/08トリガー統合（sleep廃止 → トリガーチェーン化）

- **対象プロジェクト**: `GetSalesSummaryAcquisition`
- **ブランチ名**: `feature/integrate-daily-update-trigger`（`main` から分岐）
- **開発手法**: アジャイル・スモールステップ。本指示書の変更のみで1コミット完結させる。

---

## 1. 背景と目的

現在、`07_トリガー作成スクリプト.gs` と `08_トリガー作成スクリプト2.gs` の2ファイルで、`mainWithRetry()` と `dailyUpdate()` を非同期の別トリガーとして個別に定時実行している。08番の実行時刻は07番から本来「共通の遅延秒数後」を意図していたが、実装上は時間帯ごとに個別の固定時刻（+90分/+65分/+5分）になっており、これは**修正漏れ**であることが判明した。

また `05_メイン処理とエラーハンドリング.gs` には、両者を1関数内で `Utilities.sleep(5分)` を挟んで同期的に呼び出す `integratedDailyUpdate()` がすでに存在するが、GASの6分実行制限に抵触するリスクがあるコメントが付されており、実運用では使われていない。

**今回の目的**：`integratedDailyUpdate()` の sleep 待機を廃止し、`mainWithRetry()` 完了後に `dailyUpdate()` 起動用のワンタイムトリガーを動的作成する方式（トリガーチェーン）に置き換える。これにより

- GASの6分制限を回避（各関数が独立した実行枠を持つ）
- 07/08の2ファイル運用を07（+05番のロジック）に一本化
- 遅延秒数をスクリプトプロパティで柔軟に調整可能にする

---

## 2. 変更対象ファイル

| ファイル | 変更内容 |
| :--- | :--- |
| `05_メイン処理とエラーハンドリング.gs` | `integratedDailyUpdate()` を改修。新規ヘルパー関数2つを追加 |
| `07_トリガー作成スクリプト.gs` | コード変更なし（スクリプトプロパティ変更のみ、下記5章参照） |
| `08_トリガー作成スクリプト2.gs` | **ファイルごとOldフォルダに移動** |
| `README.md` | 運用手順の記載更新、mermaid記法を追記（本指示書の対象外。別途フォローアップ） |

---

## 3. `05_メイン処理とエラーハンドリング.gs` の変更詳細

### 3-1. ファイルヘッダーコメント（1〜38行目付近）の更新

「`integratedDailyUpdate()` は内部に5分待機を含むため、GASの6分制限に注意」という記述を、実態に合わせて修正する。

- 削除: `- **時間制限**: \`integratedDailyUpdate()\` は内部に5分待機を含むため、GASの6分制限に注意してください。`
- 追加: `- **トリガーチェーン**: \`integratedDailyUpdate()\` は mainWithRetry() 完了後、sleepせずに dailyUpdate() 起動用のワンタイムトリガーを作成して終了します。両処理は別々の実行枠で動くため6分制限の影響を受けません。`

`@see integratedDailyUpdate - 統合実行関数` の説明文（37行目）も「NE取得 + 売れ数転記トリガー予約」のように更新する。

### 3-2. `integratedDailyUpdate()` 関数（461〜561行目付近）の置き換え

**変更前の処理フロー**：
1. `mainWithRetry()` 実行
2. `Utilities.sleep(5分)`
3. `dailyUpdate()` を直接呼び出し
4. 両方の結果をまとめて `displayIntegratedSummary()`

**変更後の処理フロー**：
1. `mainWithRetry()` 実行
2. 失敗時：転記処理はスキップし、その旨をログとサマリーに記録して終了（エラー通知は `mainWithRetry` 内の `main()` が既に送信済みのため、ここでの追加送信は不要）
3. 成功時：`getDailyUpdateDelaySeconds()` でスクリプトプロパティから遅延秒数を取得し、`scheduleDailyUpdateTrigger(delaySeconds)` を呼び出してワンタイムトリガーを作成
4. `results.dailyUpdate` には実行結果オブジェクトではなく `'scheduled'`（予約済み）のような状態を格納
5. `displayIntegratedSummary()` を呼び出して終了

関数のJSDocコメント（442〜460行目）も「5分待機の理由」の節を削除し、「トリガーチェーンによる非同期連携」の説明に差し替える。

### 3-3. 新規ヘルパー関数の追加

`integratedDailyUpdate()` の直後に、以下2つの関数を新規追加する。JSDocヘッダーを必ず付けること。

```javascript
/**
 * スクリプトプロパティから dailyUpdate 起動までの遅延秒数を取得
 *
 * @details
 * ハードコーディング・フォールバックデフォルトは禁止のため、
 * スクリプトプロパティ 'DAILY_UPDATE_DELAY_SECONDS' が未設定、
 * または不正な値の場合は例外を送出する。
 *
 * @return {number} 遅延秒数
 * @throws {Error} プロパティ未設定または不正値の場合
 */
function getDailyUpdateDelaySeconds() {
  // 実装はAntigravityに委任
}

/**
 * dailyUpdate() 起動用のワンタイムトリガーを作成
 *
 * @details
 * 重複起動を防ぐため、既存の dailyUpdate 向けトリガーを
 * 07_トリガー作成スクリプト.gs の deleteTriggersForFunction() で
 * 先に削除してから、指定秒数後に発火するワンタイムトリガーを作成する。
 *
 * @param {number} delaySeconds - 発火までの遅延秒数
 */
function scheduleDailyUpdateTrigger(delaySeconds) {
  // 実装はAntigravityに委任
}
```

（※上記は関数シグネチャとJSDocのみの指示。中身のロジックはAntigravityが実装。`deleteTriggersForFunction` は07番ファイルに既存のものをそのまま再利用し、重複定義しないこと。）

### 3-4. `displayIntegratedSummary()`（572〜606行目付近）の文言修正

「【売れ数転記処理】ステータス: 完了(詳細は上記参照)」の分岐を、`results.dailyUpdate === 'scheduled'` かどうかで判定し、「ステータス: 起動予約済み（非同期実行）」のような表記に変更する。転記スキップ時は「ステータス: スキップ（NE取得失敗のため）」と表示する。

---

## 4. `08_トリガー作成スクリプト2.gs` の移動

Oldフォルダへ移動させ、過去の履歴を保存しておく。

---

## 5. デプロイ時のスクリプトプロパティ変更（コード外の手動作業）

Antigravityによるコード実装完了後、GASエディタの「スクリプトプロパティ」で以下を手動変更する。**この章はコード変更ではなく運用設定変更**であることに注意。

| プロパティ | 変更内容 |
| :--- | :--- |
| `TRIGGER_FUNCTION_NAME` | `mainWithRetry` → `integratedDailyUpdate` に変更 |
| `TRIGGER_FUNCTION_NAME2` | 削除（08番ファイル廃止に伴い不要） |
| `DAILY_UPDATE_DELAY_SECONDS` | 新規追加。初期値は `30`（叩き台。運用しながら調整） |

---

## 6. テスト手順（段階的テスト）

1. **個別関数テスト**（テスト用スプレッドシート環境）
   - `getDailyUpdateDelaySeconds()` を単体実行し、正常値取得・未設定時エラーの両方を確認
   - `scheduleDailyUpdateTrigger(30)` を単体実行し、GASの「トリガー」画面に `dailyUpdate` のワンタイムトリガーが1件作成されることを確認
2. **小規模統合テスト**
   - `integratedDailyUpdate()` を手動実行し、`mainWithRetry()` 完了直後にトリガーが作成され、指定秒数後に `dailyUpdate()` が自動実行されることを確認
   - `mainWithRetry()` が失敗するケース（意図的にトークンを無効化するなど）を再現し、トリガーが作成されないことを確認
3. **本番確認**
   - `07_トリガー作成スクリプト.gs` の `setTrigger()` を実行し、`8:00/12:00/17:00` に `integratedDailyUpdate` が起動、その約30秒後に `dailyUpdate` が起動することをログで確認
   - 旧`08_トリガー作成スクリプト2.gs` によるトリガーが残っていないか、GASの「トリガー」画面で確認・手動削除

---

## 7. 注意事項（既存ルール再掲）

- 設定値はすべてスクリプトプロパティ管理。ハードコーディング・フォールバックデフォルト禁止
- 既存関数（`mainWithRetry`, `main` など）は改修しない。ロールバック可能性のため append-only 原則を維持
- 既存の説明文・JSDocコメントは可能な限り残し、内容が変わった箇所のみ修正する
- 実装後、変更点をdiff形式でレビューできるようにコミットを小さく保つ
