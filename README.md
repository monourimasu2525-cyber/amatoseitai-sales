# あまと整体院 売上管理システム

スマートフォン最適化された整体院向けシンプル売上管理Webアプリ。

## 🌐 アクセスURL

**本番URL**: https://monourimasu2525-cyber.github.io/amatoseitai-sales/

---

## 📱 機能一覧

| 機能 | 説明 |
|------|------|
| クイック登録 | 売上マスタのボタンをタップするだけで記録 |
| 今日の売上 | 新規/常連の件数・売上をリアルタイム表示 |
| 履歴管理 | 直近30日の履歴を日別表示・修正・削除 |
| 集計レポート | 今月まとめ・先月比・客単価・気づきカードを自動生成 |
| 売上マスタ管理 | 種別・金額をアプリからいつでも変更可能 |
| 経理シート生成 | 月別の経理シートをスプレッドシートに自動作成 |
| バックアップ | 手動バックアップ |

---

## 🏗️ システム構成

```
スマートフォン（ブラウザ）
        ↓
GitHub Pages（フロントエンド: index.html）
        ↓ fetch API（CORS解決）
Vercel（中継プロキシ: amato-api.vercel.app）
        ↓
Google Apps Script WebApp（バックエンド）
        ↓
Google Sheets（データベース）
```

### 各コンポーネント

| コンポーネント | 詳細 |
|-------------|------|
| フロントエンド | GitHub Pages（HTML/CSS/JS、フレームワーク不使用） |
| 中継プロキシ | Vercel Serverless Function（CORSエラー回避・キャッシュ） |
| バックエンド | Google Apps Script WebApp（REST API） |
| データベース | Google Sheets「売上データ」シート |

---

## 📊 データ構造

**売上データシート（列構成）**

| 列 | 内容 |
|----|------|
| A | 年 |
| B | 月 |
| C | 日 |
| D | 時間（HH:mm） |
| E | 種別（新規 / 常連） |
| F | 金額（円） |
| G | 入力方法（WebApp 等） |

---

## 🔧 GAS APIエンドポイント

**アクセス先（Vercel経由）**:
```
https://amato-api.vercel.app/
```

**GAS WebApp URL（直接）**:
```
https://script.google.com/macros/s/AKfycby3931P6ityh6mjGOiBVmLdRSR12lXss9TxYYAxd9XHBYKNJCPhSjBw1_b9cjh2e-_B/exec
```

### GET リクエスト

| action | 説明 | パラメータ |
|--------|------|-----------|
| `initData` | 全データ一括取得（master/todayStats/thisMonth/prevMonth/history） | なし |
| `getTodayStats` | 本日の売上集計 | なし |
| `getMonthStats` | 月間集計 | `year`, `month` |
| `getRecentHistory` | 最近の履歴 | `days`（デフォルト30） |
| `getMaster` | 売上マスタ一覧 | なし |

### POST リクエスト

| action | 説明 | ボディ例 |
|--------|------|---------|
| `addSale` | 売上登録 | `{"action":"addSale","type":"新規","amount":3270}` |
| `editSale` | 売上修正 | `{"action":"editSale","rowIndex":2,"type":"常連","amount":5500}` |
| `deleteSale` | 売上削除 | `{"action":"deleteSale","rowIndex":2}` |
| `addMaster` | マスタ追加 | `{"action":"addMaster","type":"種別名","amount":1000}` |
| `updateMaster` | マスタ更新 | `{"action":"updateMaster","rowIndex":1,"type":"新規","amount":3270}` |
| `deleteMaster` | マスタ削除 | `{"action":"deleteMaster","rowIndex":1}` |
| `generateAccounting` | 経理シート生成 | `{"action":"generateAccounting","year":2026,"month":4}` |
| `backup` | 手動バックアップ | `{"action":"backup"}` |

---

## ⚙️ GASファイル構成（src/以下）

| ファイル | 役割 |
|---------|------|
| `Code.gs` | doGet/doPost APIエントリ・onOpenメニュー |
| `SalesManager.gs` | 売上CRUD・集計・履歴取得 |
| `MasterManager.gs` | 売上マスタ管理 |
| `Summary.gs` | 集計・ダッシュボード更新 |
| `Accounting.gs` | 経理シート生成 |
| `Backup.gs` | バックアップ処理 |
| `SheetFormatter.gs` | スプレッドシート整形 |

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

### バックエンド更新（GAS）

```bash
cd /tmp/amatoseitai-sales
# src/ 以下を編集
clasp push --force
clasp deploy --deploymentId AKfycby3931P6ityh6mjGOiBVmLdRSR12lXss9TxYYAxd9XHBYKNJCPhSjBw1_b9cjh2e-_B --description "vXX - 変更内容"
```

※claspのOAuthトークンが切れている場合は再認証が必要

---

## 🔐 関連リソース

| リソース | URL / ID |
|---------|---------|
| GitHubリポジトリ | https://github.com/monourimasu2525-cyber/amatoseitai-sales |
| GitHub Pages | https://monourimasu2525-cyber.github.io/amatoseitai-sales/ |
| Vercel API | https://amato-api.vercel.app |
| スプレッドシート | https://docs.google.com/spreadsheets/d/17bAyQngDEjoDgqSLLUU5p45HWXomF09bLf_h6FySsjs |
| GASエディタ | https://script.google.com/home/projects/1Qps2uB0qRqDnTZjpZtdCAwrD9MumxpiC2m8uQroq8FuoEiR11gK-4bCb/edit |
| GAS WebApp URL | https://script.google.com/macros/s/AKfycby3931P6ityh6mjGOiBVmLdRSR12lXss9TxYYAxd9XHBYKNJCPhSjBw1_b9cjh2e-_B/exec |

---

## 📈 今後の予定

- [ ] PWA対応（ホーム画面にアイコン追加）
- [ ] スプレッドシート集計シートとの完全同期
- [ ] 月次目標設定・達成率表示
- [ ] CSV出力ボタン

---

## 📝 変更履歴

| 日付 | 内容 |
|------|------|
| 2026-04-17 | Vercel中継プロキシ導入（CORS解決） |
| 2026-04-17 | GAS v13: 売上データシート列構成修正（年/月/日/時間/種別/金額/入力方法） |
| 2026-04-17 | initData一発API・LocalStorageキャッシュ・スケルトンローダー追加 |
| 2026-04-17 | UIフルリビルド: navy accent・レポート集計画面 |
| 2026-04-16 | GAS 7ファイル分割・売上マスタ管理・修正削除機能追加 |
| 2026-04-16 | 初回デプロイ |
