# GetShipmentCompletionSlip (同梱伝票管理・経理連携自動化)

## 概要
ネクストエンジンから「出荷完了」ステータスの伝票データを月次で自動取得し、特に複雑な「同梱処理」が施された伝票の紐付け情報を抽出してGoogleスプレッドシートへ記録します。

## 開発背景と目的
EC運営において、運賃コスト削減のための「同梱処理」は不可欠ですが、従来のCSV出力では同梱元・先の紐付け情報が不足していました。そのため、現場では伝票番号を手入力でメモし、経理へ連携するというアナログな高負荷作業が発生していました。

本プロジェクトは、APIを利用することで「手動では取得困難なデータ」を自動抽出し、コスト削減と業務効率化の両立を実現するために立ち上げられました。

---

## 主な機能・特徴

### API限定データの活用
CSVダウンロードでは取得できない以下の情報を直接取得:
- 同梱元伝票番号
- 同梱先伝票番号
- 複数配送親伝票番号
- 分割元伝票番号
- 複写元伝票番号

### フルクラウド自動化
- **完全自動実行**: 毎月5日に前月分データを自動取得
- **バッチ分割処理**: 大量データを3バッチに分割し、API制限・実行時間制限を回避
- **自己スケジューリング**: 各バッチが次のバッチを自動的にスケジュール
- **Sheets API活用**: 高速・安定したデータ書き込みを実現

### 高度な処理機能
- **ウォームアップ機能**: コールドスタート対策で安定した実行速度を維持
- **トークン自動更新**: ネクストエンジンAPIトークンの自動ローテーション対応
- **クリーンアップ機能**: 月初に古いデータと不要トリガーを自動削除

---

## システム構成

### 処理フロー
【毎月5日 8:00 自動実行】
↓
warmupAndScheduleMain()
├─ クリーンアップ(前月データ削除 + 古いトリガー削除)
└─ ウォームアップ(API接続確立)
↓ 2分後
executeBatch1() → main(1)
├─ 前月1日～10日のデータ取得
├─ フィルタリング・加工
├─ Sheets APIで書き込み
└─ 5分後にバッチ2をスケジュール
↓ 5分後
executeBatch2() → main(2)

├─ 前月11日～20日のデータ取得
├─ フィルタリング・加工
├─ Sheets APIで書き込み
└─ 5分後にバッチ3をスケジュール
↓ 5分後
executeBatch3() → main(3)
├─ 前月21日～末日のデータ取得
├─ フィルタリング・加工
└─ Sheets APIで書き込み
↓
【完了】

### 取得データの条件
- **出荷予定日**: 前月1日～末日
- **受注状態**: 出荷確定済み（区分=50）または 分割・統合によりキャンセル（区分=3）

---

## セットアップ

### 1. 必須ライブラリの追加

#### NEAuth認証ライブラリ
1. GASエディタ左メニュー「ライブラリ」の「+」をクリック
2. 認証プロジェクトのスクリプトIDを入力
3. 最新バージョンを選択
4. 識別子: `NEAuth` と入力

#### Google Sheets API
1. GASエディタ左メニュー「サービス」の「+」をクリック
2. 「Google Sheets API」を検索
3. バージョン「v4」を選択して追加

### 2. スクリプトプロパティの設定

GASエディタ「プロジェクトの設定」→「スクリプト プロパティ」

| キー | 値 | 説明 |
|------|-----|------|
| `CLIENT_ID` | (ネクストエンジンアプリのID) | API認証用 |
| `CLIENT_SECRET` | (シークレット) | API認証用 |
| `REDIRECT_URI` | (WebアプリのデプロイURL) | 認証リダイレクト先 |
| `SPREADSHEET_ID` | (スプレッドシートID) | 出力先シートのID |
| `SHEET_NAME` | (シート名) | 出力先タブ名 |
| `IS_DEV_MODE` | `false` | 本番: `false`, 開発: `true` |

### 3. 初回認証

```javascript
// 1. 認証URL生成
testGenerateAuthUrl();
// → 表示されたURLをブラウザで開く

// 2. 認証完了後、接続テスト
testApiConnection();
```

### 4. トリガー設定

GASエディタ「トリガー」→「トリガーを追加」

| 項目 | 設定値 |
|------|--------|
| 実行する関数 | `warmupAndScheduleMain` |
| イベントのソース | 時間主導型 |
| タイプ | 月ベースのタイマー |
| 日 | 5日 |
| 時刻 | 午前8時~9時 |

※パラメータは不要（デフォルトでバッチ1が実行される）

---

## 開発モードでのテスト

### 開発モード設定
```javascript
// config.gs
DEV_MODE: {
    ENABLED: true,  // 開発モード有効化
    TARGET_DATE: '2025/10/03'  // テスト用基準日
}
```

### 手動テスト実行
```javascript
// ウォームアップ + バッチ1
warmupAndScheduleMain(1);
// → 2分後に自動的にバッチ1が実行される

// ★ 10分以上待つ ★

// ウォームアップ + バッチ2
warmupAndScheduleMain(2);

// ★ 10分以上待つ ★

// ウォームアップ + バッチ3
warmupAndScheduleMain(3);
```

**注意**: 手動テスト時は各バッチ間に10分以上の間隔を空けてください（Sheets API制限回避のため）

---

## ファイル構成

| ファイル名 | 役割 |
|-----------|------|
| `config.gs` | 全体設定（API URL、フィールド定義、バッチ設定） |
| `main.gs` | メイン処理・バッチ実行・自己スケジューリング |
| `api.gs` | ネクストエンジンAPI呼び出し・トークン管理 |
| `process.gs` | データフィルタリング・加工処理 |
| `sheet.gs` | Sheets APIによるスプレッドシート書き込み |
| `utils.gs` | 日付計算・ウォームアップ処理 |
| `cleanup.gs` | 古いデータ・トリガーのクリーンアップ |
| `NE_認証ライブラリ使用必須関数.gs` | 認証関連の補助関数 |

---

## 取得フィールド一覧

| ヘッダー名 | APIフィールド名 |
|-----------|----------------|
| 店舗コード | `receive_order_shop_id` |
| 伝票番号 | `receive_order_id` |
| 受注番号 | `receive_order_shop_cut_form_id` |
| 出荷予定日 | `receive_order_send_plan_date` |
| 受注日 | `receive_order_date` |
| 購入者名 | `receive_order_purchaser_name` |
| 支払方法 | `receive_order_payment_method_name` |
| 総合計 | `receive_order_total_amount` |
| 商品計 | `receive_order_goods_amount` |
| 税金 | `receive_order_tax_amount` |
| 発送代 | `receive_order_delivery_fee_amount` |
| 手数料 | `receive_order_charge_amount` |
| 他費用 | `receive_order_other_amount` |
| ポイント数 | `receive_order_point_amount` |
| 発送伝票番号 | `receive_order_delivery_cut_form_id` |
| 支払区分 | `receive_order_payment_method_id` |
| 同梱先伝票番号 | `receive_order_include_to_order_id` |
| 複数配送親伝票番号 | `receive_order_multi_delivery_parent_order_id` |
| 分割元伝票番号 | `receive_order_divide_from_order_id` |
| 複写元伝票番号 | `receive_order_copy_from_order_id` |
| 複数配送親フラグ | `receive_order_multi_delivery_parent_flag` |
| 受注キャンセル区分 | `receive_order_cancel_type_id` |
| 受注キャンセル名 | `receive_order_cancel_type_name` |
| 受注キャンセル日 | `receive_order_cancel_date` |
| 受注状態区分 | `receive_order_order_status_id` |
| 受注状態名 | `receive_order_order_status_name` |

---

## トラブルシューティング

### トークンエラーが発生する
```javascript
// 再認証を実行
testGenerateAuthUrl();
// → URLをブラウザで開いて認証

// 接続確認
testApiConnection();
```

### バッチが途中で止まった
- GAS実行ログで停止箇所を確認
- 手動で続きのバッチを実行: `warmupAndScheduleMain(2)` または `warmupAndScheduleMain(3)`
- 次回の自動実行で自動的にリセットされる

### データが重複している
```javascript
// クリーンアップを手動実行
cleanupOldTriggersAndData();
```

### Sheets API書き込みエラー
- 出力先スプレッドシートのアクセス権限を確認
- スプレッドシートIDが正しく設定されているか確認
- バッチ間隔を10分に延長（`config.gs`の`INTERVAL_MINUTES`）

---

## 実務上のメリット

本システムの導入により、以下の効果が期待されます:

### コスト削減
- **運賃軽減**: 同梱処理による配送費削減効果を正確に把握
- **人件費削減**: 手動メモ作業・突合作業の完全自動化

### 業務品質向上
- **ヒューマンエラー排除**: 手入力ミスをゼロ化
- **処理速度向上**: 月次締め作業の大幅な時間短縮
- **データ精度向上**: API直接取得による確実な紐付け情報

### 組織的メリット
- **全社的生産性向上**: 現場・経理セクション双方の業務効率化
- **データドリブン経営**: 同梱効果の定量的な把握が可能に
- **スケーラビリティ**: 取扱量増加にも自動対応

---

## 開発履歴

- 初版: Gemini 2.0 Proで基本機能を実装
- リファクタリング: Claude Sonnet 4.5でコード可読性向上
- 月次バッチ化: 3バッチ分割・自己スケジューリング実装
- 安定化: Sheets API導入、ウォームアップ機能追加

---

## ライセンス・注意事項

本プロジェクトは社内業務効率化を目的としています。
ネクストエンジンAPIの利用規約を遵守し、取得データの取り扱いには十分注意してください。