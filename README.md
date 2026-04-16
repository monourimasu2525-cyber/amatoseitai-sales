# あまと整体院 売上管理システム

スマートフォン最適化された整体院向けシンプル売上管理Webアプリ。

## 🌐 アクセスURL

**本番URL**: https://monourimasu2525-cyber.github.io/amatoseitai-sales/

---

## 📱 機能一覧

| 機能 | 説明 |
|------|------|
| 売上登録 | 新規（¥3,270）・常連（¥5,500）をワンタップで記録 |
| 本日集計 | 新規/常連の件数・売上をリアルタイム表示 |
| 前月比較 | 今月と先月の売上を並べて比較 |
| 履歴表示 | 月別に売上履歴を一覧表示 |
| CSV出力 | 月別データをCSVでダウンロード |
| バックアップ | 手動ボタン or 毎日深夜0〜1時に自動バックアップ |

---

## 🏗️ システム構成

```
スマートフォン（ブラウザ）
        ↓
GitHub Pages（フロントエンド: index.html）
        ↓ fetch API
Google Apps Script WebApp（バックエンド: Code.gs）
        ↓
Google Sheets（データベース: 売上データシート）
        ↓
Google Drive（バックアップ: あまと整体院_売上バックアップフォルダ）
```

### 各コンポーネント

| コンポーネント | 詳細 |
|-------------|------|
| フロントエンド | GitHub Pages（HTML/CSS/JS、フレームワーク不使用） |
| バックエンド | Google Apps Script WebApp（REST API） |
| データベース | Google Sheets「売上データ」シート |
| バックアップ先 | Google Drive「あまと整体院_売上バックアップ」フォルダ |

---

## 📊 データ構造

**売上データシート（列構成）**

| 列 | 内容 |
|----|------|
| A | タイムスタンプ（登録日時） |
| B | タイムスタンプ（同上） |
| C | 種別（新規 / 常連） |
| D | 金額 |
| E | 入力元（WebAPI / スプレッドシート直接入力等） |

---

## 🔧 GAS APIエンドポイント

**GAS WebApp URL**:
```
https://script.google.com/macros/s/AKfycbw6VmizWzyTprGGs4rVB7i9L9jlf3ifwu_YjjuVTHyWIwcBB5MdUCz-GrUwKaG29s8r/exec
```

### GET リクエスト

| action | 説明 | パラメータ |
|--------|------|-----------|
| `getTodayStats` | 本日の売上集計（デフォルト） | なし |
| `getMonthStats` | 月間集計 | `year`, `month` |
| `getRecentHistory` | 最近の履歴 | `days`（デフォルト30） |
| `getCsv` | CSV形式でデータ取得 | `year`, `month`（省略=全件） |
| `getSheets` | シート名確認（デバッグ用） | なし |

### POST リクエスト

| action | 説明 | ボディ例 |
|--------|------|---------|
| `addSale` | 売上登録 | `{"action":"addSale","type":"新規","amount":3270}` |
| `backup` | 手動バックアップ | `{"action":"backup"}` |

---

## ⚙️ GAS関数一覧

| 関数名 | 用途 |
|--------|------|
| `doGet(e)` | GET リクエスト処理 |
| `doPost(e)` | POST リクエスト処理 |
| `runBackup()` | バックアップ実行 |
| `dailyAutoBackup()` | 自動バックアップ（トリガーから呼ばれる） |
| `setupDailyBackupTrigger()` | 自動バックアップトリガー登録（初回のみ手動実行） |
| `addNewCustomer()` | スプレッドシートメニューから新規登録 |
| `addRegularCustomer()` | スプレッドシートメニューから常連登録 |
| `manualBackupFromSheet()` | スプレッドシートメニューから手動バックアップ |

---

## 🚀 デプロイ・更新手順

### フロントエンド更新（index.html）

```bash
cd /tmp/amatoseitai-sales
# index.html を編集
git add index.html
git commit -m "変更内容"
git push origin main
# → 数分でGitHub Pagesに反映
```

### バックエンド更新（Code.gs）

1. GASエディタ（https://script.google.com/home/projects/1Qps2uB0qRqDnTZjpZtdCAwrD9MumxpiC2m8uQroq8FuoEiR11gK-4bCb/edit）を開く
2. Code.gsを編集・保存
3. 「デプロイ」→「デプロイを管理」→「編集」→バージョン「新バージョン」選択→「デプロイ」
4. **URLは変わらない**（同じデプロイIDを更新するため）

---

## 🔐 関連リソース

| リソース | URL / ID |
|---------|---------|
| GitHubリポジトリ | https://github.com/monourimasu2525-cyber/amatoseitai-sales |
| GitHub Pages | https://monourimasu2525-cyber.github.io/amatoseitai-sales/ |
| スプレッドシート | https://docs.google.com/spreadsheets/d/17bAyQngDEjoDgqSLLUU5p45HWXomF09bLf_h6FySsjs |
| GASエディタ | https://script.google.com/home/projects/1Qps2uB0qRqDnTZjpZtdCAwrD9MumxpiC2m8uQroq8FuoEiR11gK-4bCb/edit |

---

## 📈 今後の予定

- [ ] UI改善（スマホ操作性向上）
- [ ] 前月比較グラフ表示
- [ ] スプレッドシート集計シート連携
- [ ] 経理シート自動生成

---

## 📝 変更履歴

| 日付 | バージョン | 内容 |
|------|-----------|------|
| 2026-04-16 | v1 | 初回デプロイ |
| 2026-04-16 | v2 | addHeaderバグ修正、CORS対応 |
| 2026-04-16 | v3 | バックアップ機能・履歴・CSV・前月比較追加 |
| 2026-04-16 | v4 | シート名フォールバック対応、デバッグエンドポイント追加 |
