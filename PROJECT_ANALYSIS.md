# 🏥 あまと整体院 売上管理システム — 全体理解ドキュメント

**作成日**: 2026年4月26日  
**分析者**: Claude  
**プロジェクト規模**: 1,058行 GAS コード + フロントエンド

---

## 📋 目次

1. [プロジェクト概要](#プロジェクト概要)
2. [システムアーキテクチャ](#システムアーキテクチャ)
3. [GAS ファイル構成](#gas-ファイル構成)
4. [データベース構造](#データベース構造)
5. [API エンドポイント](#api-エンドポイント)
6. [現在の機能一覧](#現在の機能一覧)
7. [コード品質分析](#コード品質分析)
8. [改善機会の発見](#改善機会の発見)

---

## 🎯 プロジェクト概要

### 概要
スマートフォン最適化された**整体院向けシンプル売上管理Webアプリ**。  
日々の売上を手軽に記録し、リアルタイムで集計・分析できるシステム。

### 対象ユーザー
- 整体院スタッフ
- 経営者（分析・集計用）

### 主要機能
| 機能 | 説明 |
|------|------|
| ✅ クイック登録 | 売上マスタのボタンをタップするだけで記録 |
| 📊 今日の売上 | 新規/常連の件数・売上をリアルタイム表示 |
| 📜 履歴管理 | 直近30日の履歴を日別表示・修正・削除 |
| 📈 集計レポート | 今月まとめ・先月比・客単価・気づきカード自動生成 |
| 🔧 売上マスタ管理 | 種別・金額をアプリからいつでも変更可能 |
| 💼 経理シート生成 | 月別の経理シートをスプレッドシートに自動作成 |
| 💾 バックアップ | 手動バックアップ + 自動日次バックアップ（深夜2時） |

---

## 🏗️ システムアーキテクチャ

### 3層構成

```
┌─────────────────────────────────────────┐
│  フロントエンド層                        │
│  GitHub Pages (HTML/CSS/JS)             │
│  フレームワーク不使用（Vanilla JS）      │
└────────────────────┬────────────────────┘
                     │ fetch API
                     ↓
┌─────────────────────────────────────────┐
│  中継層（CORS解決）                      │
│  Vercel Serverless Function             │
│  URL: amato-api.vercel.app              │
└────────────────────┬────────────────────┘
                     │ REST API
                     ↓
┌─────────────────────────────────────────┐
│  バックエンド層                          │
│  Google Apps Script (GAS) WebApp        │
│  REST API エンドポイント                 │
└────────────────────┬────────────────────┘
                     │ Apps Script API
                     ↓
┌─────────────────────────────────────────┐
│  データベース層                          │
│  Google Sheets                          │
│  シート: 売上データ / 売上マスタ / 集計  │
└─────────────────────────────────────────┘
```

### なぜこの構成？
- **GitHub Pages**: 静的ホスティング（無料）
- **Vercel**: CORS エラーを回避（GAS の doPost は CORS に対応しない）
- **Google Apps Script**: ビジネスロジック + Sheets 操作
- **Google Sheets**: データ永続化 + スプレッドシート による追加分析

---

## 🔧 GAS ファイル構成

### ファイル一覧と役割

| ファイル | 行数 | 役割 | クラス/関数数 |
|---------|------|------|-------------|
| **Code.gs** | 221 | API エントリポイント | doPost, doGet, onOpen |
| **SalesManager.gs** | 169 | 売上 CRUD・集計 | class SalesManager |
| **Summary.gs** | 283 | 集計・ダッシュボード更新 | updateSummarySheet など |
| **MasterManager.gs** | 110 | 売上マスタ管理 | getMasterItems など |
| **Accounting.gs** | 133 | 経理シート生成 | generateAccountingSheet |
| **Backup.gs** | 48 | バックアップ処理 | runBackup, dailyAutoBackup |
| **SheetFormatter.gs** | 94 | スプレッドシート整形 | formatSalesDataSheet など |
| **合計** | **1,058** | | |

---

## 🗂️ ファイル詳細

### 1. **Code.gs** — API エントリ（221行）

#### 役割
- REST API のエントリポイント（doPost/doGet）
- メニュー関数（onOpen）
- リクエストのルーティング

#### 主要関数

```javascript
doPost(e)
  ├─ data.action = 'addSale' → SalesManager.addSale()
  ├─ data.action = 'editSale' → SalesManager.editSale()
  ├─ data.action = 'deleteSale' → SalesManager.deleteSale()
  ├─ data.action = 'addMaster' → addMasterItem()
  ├─ data.action = 'updateMaster' → updateMasterItem()
  ├─ data.action = 'deleteMaster' → deleteMasterItem()
  ├─ data.action = 'backup' → runBackup()
  └─ data.action = 'initData' → 全データ一括返却

doGet(e)
  └─ action=getTodayStats / getMonthStats / getRecentHistory など

onOpen()
  └─ スプレッドシートメニューを生成
```

#### 特徴
- JSON ベースのリクエスト/レスポンス
- 例外処理あり
- SPREADSHEET_ID をグローバルで管理

---

### 2. **SalesManager.gs** — 売上管理（169行）

#### 役割
売上データの CRUD 操作と集計処理を一元管理するクラス。

#### クラス: SalesManager

```javascript
class SalesManager {
  constructor(spreadsheetId)
  
  // CRUD
  addSale(type, amount)           // 売上追加
  editSale(rowIndex, type, amount) // 売上修正
  deleteSale(rowIndex)            // 売上削除
  
  // 集計
  getTodayStats()                 // 本日の集計
  getMonthStats(year, month)      // 月間集計
  getRecentHistory(days)          // 直近 N 日の履歴
  getCsvData(year, month)         // CSV 形式のデータ取得
}
```

#### データ構造（売上データシート）

| 列 | 項目 | 型 | 例 |
|----|------|-----|-----|
| A | 作成日時 | Date | 2026-04-26 14:30:45 |
| B | 更新日時 | Date | 2026-04-26 14:30:45 |
| C | 種別 | String | "新規" / "常連" / 他 |
| D | 金額 | Number | 3270 |
| E | 入力方法 | String | "WebAPI" / "WebAPI（修正）" |

#### 集計ロジック
- 日付・種別・金額でフィルタリング
- 件数・売上を 新規/常連/その他 で分類
- リアルタイム集計（毎回 getLastRow で全データを取得）

---

### 3. **Summary.gs** — 集計・ダッシュボード（283行）

#### 役割
月間の売上を集計し、見栄え良く整形したサマリーシートを自動生成。

#### 主要関数

```javascript
updateSummarySheet()          // 12ヶ月の月別集計をテーブル化
createDashboard()             // 当月のダッシュボードシートを作成
(ダッシュボード機能の詳細は不確認)
```

#### 出力
- **集計シート**: 12ヶ月分の月別集計（新規/常連 件数・売上）
- **グラフ**: 月別売上の推移を棒グラフで表示

#### フォーマット
- ヘッダー行: 紺色 (#1a237e) + 白字
- データ行: 交互背景（白 / 薄紫 #e8eaf6）
- 合計行: 紺色 + 白字
- 通貨形式: ¥#,##0

---

### 4. **MasterManager.gs** — マスタ管理（110行）

#### 役割
売上マスタ（種別・金額テンプレート）の CRUD 管理。

#### 主要関数

```javascript
getMasterItems()                // 有効なマスタ一覧を取得
addMasterItem(type, amount, desc) // マスタを追加
updateMasterItem(rowIndex, ...)  // マスタを更新
deleteMasterItem(rowIndex)      // マスタを無効化（物理削除ではなく論理削除）
initMasterSheet()               // マスタシートを初期化
```

#### データ構造（売上マスタシート）

| 列 | 項目 | 型 | 例 |
|----|------|-----|-----|
| A | 種別 | String | "新規" |
| B | 金額 | Number | 3270 |
| C | 説明 | String | "新規施術" |
| D | 有効 | Boolean | TRUE / FALSE |

#### 特徴
- 論理削除（有効フラグを FALSE に）
- 初期化時に「新規（¥3,270）」「常連（¥5,500）」を自動作成
- フロントエンドはこのマスタからボタンを動的生成

---

### 5. **Accounting.gs** — 経理シート生成（133行）

#### 役割
指定月の売上を日別で集計し、経理用のシートを自動生成。

#### 主要関数

```javascript
generateAccountingSheet(year, month)
```

#### 生成物
- **シート名**: `経理_2026_04` （形式: 経理_YYYY_MM）
- **内容**: 月の各日付ごとの 新規件数・売上、常連件数・売上、日計
- **フォーマット**: ヘッダー + 日別データ + 月合計行

#### 特徴
- 既存シートがあれば削除して再生成
- 日数に応じた動的行生成（28〜31行）
- 通貨形式・交互背景色・枠線を自動適用

---

### 6. **Backup.gs** — バックアップ（48行）

#### 役割
スプレッドシート全体のコピーを Google Drive に保存。

#### 主要関数

```javascript
runBackup()                 // 手動バックアップ実行
getOrCreateBackupFolder()   // バックアップフォルダを取得/作成
dailyAutoBackup()           // トリガーから自動実行
setupDailyBackupTrigger()   // 自動バックアップトリガーをセット
```

#### バックアップ仕様
- **フォルダ名**: `あまと整体院_売上バックアップ`
- **ファイル名**: `あまと整体院_売上データ_2026-04-26_143045`（年月日_時分秒）
- **自動実行**: 毎日深夜 2時（トリガー設定済み）
- **復旧方法**: Google Drive から手動でコピーを復元

---

### 7. **SheetFormatter.gs** — シート整形（94行）

#### 役割
スプレッドシート各シートの見栄え（色・枠線・フォント・フォーマット）を統一。

#### 主要関数

```javascript
formatSalesDataSheet()    // 売上データシートの整形
formatMasterSheet()       // マスタシートの整形
formatAllSheets()         // 全シートの一括整形 + ダッシュボード更新
```

#### 適用フォーマット
- **ヘッダー**: 紺色背景 + 白字 + 太字 + 中央揃え
- **データ行**: 交互背景色（白 / 薄紫）
- **数値列**: 通貨形式 (¥#,##0) / 日時形式 (yyyy/MM/dd HH:mm)
- **枠線**: 0.5px グレー (#9e9e9e)
- **フリーズ**: ヘッダー行固定
- **列幅**: 自動調整

---

## 💾 データベース構造

### Google Sheets 内のシート

| シート名 | 用途 | 行数 | 列数 | 備考 |
|---------|------|------|------|------|
| **売上データ** | 売上記録 | 変動（~1000行） | 5 | メインデータ |
| **売上マスタ** | マスタテンプレート | 数行 | 4 | 新規・常連など種別 |
| **集計** | 月別集計テーブル + グラフ | 14行 | 7 | 自動生成 |
| **経理_YYYY_MM** | 月別経理レポート | 日数+2行 | 6 | 必要に応じて生成 |
| **ダッシュボード** | 当月サマリー | 可変 | 可変 | 当月の集計サマリー |

### 売上データシート（詳細）

```
A列: 作成日時 (DateTime)
B列: 更新日時 (DateTime) ← 修正時に更新
C列: 種別 (String) ← "新規" / "常連" / カスタム
D列: 金額 (Number) ← 1000以上の整数
E列: 入力方法 (String) ← "WebAPI", "WebAPI（修正）" など

例:
┌──────────────────┬──────────────────┬────────┬────────┬──────────────┐
│ 2026-04-26 14:30 │ 2026-04-26 14:30 │ 新規   │ 3270   │ WebAPI       │
│ 2026-04-25 10:15 │ 2026-04-25 10:15 │ 常連   │ 5500   │ WebAPI       │
│ 2026-04-24 16:45 │ 2026-04-25 09:00 │ 常連   │ 5500   │ WebAPI（修正）│
└──────────────────┴──────────────────┴────────┴────────┴──────────────┘
```

### 売上マスタシート（詳細）

```
A列: 種別 (String)
B列: 金額 (Number)
C列: 説明 (String)
D列: 有効フラグ (Boolean)

例:
┌────────┬────────┬──────────────┬────────┐
│ 種別   │ 金額   │ 説明         │ 有効   │
├────────┼────────┼──────────────┼────────┤
│ 新規   │ 3270   │ 新規施術     │ TRUE   │
│ 常連   │ 5500   │ 常連施術     │ TRUE   │
│ 割引   │ 2700   │ 割引後       │ FALSE  │
└────────┴────────┴──────────────┴────────┘
```

---

## 🌐 API エンドポイント

### ベース URL
```
Vercel 経由: https://amato-api.vercel.app/
GAS 直接: https://script.google.com/macros/s/{DEPLOYMENT_ID}/exec
```

### GET リクエスト

#### 1. initData（初回ロード用）
```
GET /api?action=initData

レスポンス:
{
  master: [ {rowIndex, type, amount, description}, ... ],
  todayStats: { date, shinkiCount, jorenCount, ..., totalSales },
  thisMonth: { shinkiCount, jorenCount, ..., totalSales },
  prevMonth: { shinkiCount, jorenCount, ..., totalSales },
  history: { records: [ {rowIndex, date, time, type, amount}, ... ] }
}
```

#### 2. getTodayStats（本日集計）
```
GET /api?action=getTodayStats

レスポンス:
{
  date: "2026年4月26日",
  shinkiCount: 3,
  jorenCount: 5,
  totalCount: 8,
  shinkiSales: 9810,
  jorenSales: 27500,
  otherCount: 0,
  otherSales: 0,
  totalSales: 37310
}
```

#### 3. getMonthStats（月間集計）
```
GET /api?action=getMonthStats&year=2026&month=4

レスポンス:
{
  shinkiCount: 42,
  jorenCount: 98,
  totalCount: 140,
  shinkiSales: 137340,
  jorenSales: 539000,
  totalSales: 676340
}
```

#### 4. getRecentHistory（最近の履歴）
```
GET /api?action=getRecentHistory&days=30

レスポンス:
{
  records: [
    {rowIndex: 102, date: "2026/4/26", time: "14:30", type: "新規", amount: 3270},
    ...
  ]
}
```

### POST リクエスト

#### 1. addSale（売上追加）
```
POST /api

Body:
{
  "action": "addSale",
  "type": "新規",
  "amount": 3270
}

レスポンス:
{
  "success": true,
  "message": "新規 ¥3270 を登録しました",
  "timestamp": "2026-04-26T14:30:45.123Z",
  "type": "新規",
  "amount": 3270
}
```

#### 2. editSale（売上修正）
```
POST /api

Body:
{
  "action": "editSale",
  "rowIndex": 102,
  "type": "常連",
  "amount": 5500
}

レスポンス:
{
  "success": true,
  "message": "行102 を修正しました: 常連 ¥5500"
}
```

#### 3. deleteSale（売上削除）
```
POST /api

Body:
{
  "action": "deleteSale",
  "rowIndex": 102
}

レスポンス:
{
  "success": true,
  "message": "行102 を削除しました"
}
```

#### 4. addMaster（マスタ追加）
```
POST /api

Body:
{
  "action": "addMaster",
  "type": "割引新規",
  "amount": 2700,
  "description": "初回割引"
}

レスポンス:
{
  "success": true,
  "message": "割引新規 を追加しました"
}
```

#### 5. generateAccounting（経理シート生成）
```
POST /api

Body:
{
  "action": "generateAccounting",
  "year": 2026,
  "month": 4
}

レスポンス:
{
  "success": true,
  "message": "経理_2026_04 を生成しました（30日分）",
  "sheetName": "経理_2026_04",
  "spreadsheetUrl": "https://docs.google.com/spreadsheets/d/17bAy.../"
}
```

#### 6. backup（バックアップ実行）
```
POST /api

Body:
{
  "action": "backup"
}

レスポンス:
{
  "success": true,
  "message": "バックアップ完了: あまと整体院_売上データ_2026-04-26_143045",
  "fileId": "1Xyz..."
}
```

---

## ✅ 現在の機能一覧

### フロントエンド機能（推定）
- [ ] クイック登録（マスタボタン）
- [ ] 本日の売上表示（リアルタイム）
- [ ] 履歴表示・修正・削除
- [ ] 集計レポート（当月・先月比較）
- [ ] マスタ管理（追加・編集・削除）
- [ ] 経理シート生成ボタン
- [ ] バックアップボタン
- [ ] ローカルストレージ キャッシュ
- [ ] スケルトンローダー

### バックエンド機能（確認済み）
- ✅ 売上 CRUD（追加・修正・削除）
- ✅ 日別・月別集計
- ✅ マスタ管理 CRUD
- ✅ 経理シート自動生成
- ✅ 集計テーブル + グラフ生成
- ✅ 手動バックアップ
- ✅ 自動日次バックアップ（深夜2時）
- ✅ スプレッドシート整形（色・枠線・フォント）

---

## 🔍 コード品質分析

### 強み ✅

1. **ファイル分割が明確**
   - 役割ごとに分離（SalesManager, MasterManager, etc.）
   - 関心の分離が実装されている

2. **エラーハンドリング実装**
   - try-catch で例外をキャッチ
   - ユーザーフレンドリーなエラーメッセージ

3. **データ検証**
   - 金額の正の数チェック
   - 種別の必須チェック

4. **フォーマット統一**
   - 色・フォント・枠線が統一（紺色 + 白字）
   - ユーザビリティが高い

5. **API 設計が明確**
   - JSON ベース
   - action パラメータで機能分岐
   - GET/POST の使い分け

### 課題 ⚠️

1. **パフォーマンス**
   - `getLastRow()` を毎回実行（全データ取得）
   - 1000行超えるとレスポンス遅延の可能性
   - **解決案**: キャッシング / インデックス構造

2. **時間フォーマットの不具合**
   - `date.getHours() + ':' + padStart(minutes)`
   - 時間に 0 パディングなし → 5時が「5:00」になる
   - **解決案**: `padStart(2, '0')` を時間にも適用

3. **エラーロギング不足**
   - エラーメッセージはあるが、ログ記録がない
   - トラブル時の原因特定が困難
   - **解決案**: エラーログシートを追加

4. **トランザクション処理なし**
   - 複数操作の途中でエラー → 不整合状態
   - **解決案**: ロールバック機構 or トランザクション

5. **入力値バリデーション不足**
   - type が未指定の場合のチェックはあるが、不正な値は許容
   - **解決案**: ホワイトリスト検証

6. **CSVデータの未活用**
   - `getCsvData()` メソッドが実装されているが、エクスポート機能がない
   - **解決案**: CSV ダウンロード機能を追加

---

## 💡 改善機会の発見

### 優先度 🔴 高

1. **時間フォーマット修正（バグ）**
   - 影響: 時刻表示が不正確
   - 工数: 5分
   - 内容: `padStart(2, '0')` を時間にも適用
   - ファイル: SalesManager.gs の getRecentHistory

2. **パフォーマンス最適化**
   - 影響: 大規模データで応答遅延
   - 工数: 1-2時間
   - 内容: キャッシング層の追加 / インデックス管理
   - ファイル: SalesManager.gs

3. **エラーログ記録**
   - 影響: トラブル時の原因追跡が困難
   - 工数: 30分
   - 内容: ログシート作成 + エラーログ出力
   - ファイル: Code.gs

### 優先度 🟡 中

4. **入力値バリデーション強化**
   - 影響: 不正な種別・金額が混入する可能性
   - 工数: 1時間
   - 内容: ホワイトリスト検証 + 範囲チェック
   - ファイル: SalesManager.gs, MasterManager.gs

5. **CSV エクスポート機能**
   - 影響: 外部ツール連携が困難
   - 工数: 1時間
   - 内容: getCsvData() を活用し、CSV ダウンロード機能を実装
   - ファイル: Code.gs

6. **トランザクション処理**
   - 影響: 複数操作時のデータ不整合
   - 工数: 2-3時間
   - 内容: ロールバック機構 or 操作のアトミック化
   - ファイル: SalesManager.gs

### 優先度 🟢 低

7. **自動テスト導入**
   - 影響: 回帰テストが手動 / 保守性低下
   - 工数: 3-4時間
   - 内容: GAS テストフレームワーク導入
   - ファイル: 全体

8. **ドキュメント整備**
   - 影響: 新規開発者のオンボーディング困難
   - 工数: 1-2時間
   - 内容: API リファレンス / 開発ガイド作成
   - ファイル: README.md / docs/

---

## 📊 改善ロードマップ（推奨）

### フェーズ 1（今週）：バグ修正 + 基礎改善
- [ ] 時間フォーマット修正
- [ ] エラーログ記録機構
- [ ] 入力値バリデーション強化

### フェーズ 2（来週）：パフォーマンス
- [ ] キャッシング層の実装
- [ ] データベース インデックス最適化
- [ ] API レスポンス時間計測

### フェーズ 3（2週目以降）：機能追加
- [ ] CSV エクスポート
- [ ] 日別レポート機能
- [ ] トランザクション処理

---

## 🎯 次のステップ

ユーザーの指示により、以下のいずれかを実施：

1. **フェーズ 1 の改善を自動で実装 + push**
2. **特定の改善項目をピンポイント実装**
3. **新機能を追加実装**
4. **全ファイルをリファクタリング**

---

**作成者**: Claude Haiku 4.5  
**作成日**: 2026-04-26  
**バージョン**: 1.0 (初版)
