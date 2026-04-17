# タイムカードアプリ - プロジェクトコンテキスト

## 🚨 最重要ルール
1. **日本語で回答・コミットメッセージも日本語**
2. **Next.jsやVite等のフレームワークは絶対に導入しない** — このアプリは単一HTMLファイル (`timecard.html`) で完結しており、ビルドステップなし
3. **不要なパッケージを `npm install` しない** — このプロジェクトにpackage.jsonは不要
4. **GitHub Actionsや`.github/workflows/`は不要** — デプロイはNetlify（GitHub連携）で自動
5. **変更前に必ずgitの現状を確認し、既存コードを理解してから編集すること**

## プロジェクト概要

飲食店のスタッフ出退勤管理を行うWebアプリ。
- **運用形態**: 店舗のPCやスタッフ個人のスマートフォンからアクセス
- **デプロイ先**: Vercel（GitHub連携で自動デプロイ）
  - Netlifyのビルドクレジット切れのためVercelに移行済み
- **バックエンド**: Google Apps Script (GAS) — Google スプレッドシートにデータを保存
- **ユーザー**: 飲食店オーナー（管理者）とスタッフ

## 技術スタック

| 項目 | 技術 |
|------|------|
| フロントエンド | **単一HTML** (`timecard.html`) — React 18 + Babel (CDN) + Tailwind CSS (CDN) |
| バックエンド | **Google Apps Script** (`gas_backup.gs`) — スプレッドシート連携 |
| Excel出力 | **ExcelJS** (CDN) — 埋め込みBase64テンプレートからタイムカード帳票を生成 |
| ホスティング | Vercel (旧: Netlify → GitHub Pages) |
| リポジトリ | GitHub: `sbsrd3mission/timecard` |

## ファイル構成

```
timecard/
├── timecard.html          # 🔑 メインアプリ（全機能がこの1ファイルに集約）
├── gas_backup.gs          # 🔑 GASバックエンドスクリプト（スプレッドシートにコピーして使用）
├── vercel.json            # Vercel設定（ルートをtimecard.htmlにリライト）
├── netlify.toml           # 旧Netlify設定（参考用、現在はVercelを使用）
├── .gitignore
├── convert_to_docx.py     # 取扱説明書のMarkdown → docx変換スクリプト
├── convert_staff_docx.py  # スタッフ用説明書の変換スクリプト
├── test_excel.html        # Excel出力のテスト用ページ
├── タイムカード入力空バージョン.xlsx  # Excelテンプレート元ファイル
├── 外注用タイムカード.xlsx           # 外注者用Excelテンプレート
├── GAS設定手順書.md
├── タイムカード設定情報.md
├── ロードの仕方.md
├── 取扱説明書_オーナー・店長用.md/.docx
├── 取扱説明書_スタッフ用.md/.docx
├── コレトちゃん.avif       # マスコットキャラ画像
├── レクトくん.png          # マスコットキャラ画像
├── timecard1.jpg           # 取扱説明書用スクリーンショット
├── timecard2.jpg
└── manual_images/          # 取扱説明書用画像
```

## アプリの主要機能

### スタッフ向け
- 出勤・退勤のワンタップ打刻
- 中抜け（休憩）の記録
- 賄い有無の記録
- 自己修正機能（備考に「本人修正済」タグ付き）
- 有給申請
- 月次履歴の閲覧

### 管理者向け（PINコード認証）
- スタッフの追加・削除
- スタッフ属性の管理（外注フラグ、時給、交通費）
- 全スタッフの打刻データ修正
- 月次Excelタイムカード出力（テンプレートベース）
- 外注者用のExcel出力（給与計算付き）
- GAS接続URL設定
- PINコード変更

### データ同期
- GAS経由でGoogle スプレッドシートに双方向同期
- 3秒ポーリングによるリアルタイム同期
- フォーカス復帰時の即時同期
- 未送信データの自動リトライ
- 競合・重複防止のためのローカル更新タイマー（8秒ロック）
- スタッフリスト・PINコードのクラウド同期

## GAS (gas_backup.gs) の構造

- `doGet(e)`: GET エンドポイント
  - `?action=getAll` — 全打刻データ取得
  - `?action=getSettings` — スタッフリスト・PIN取得
  - `?action=ping` — 接続テスト
- `doPost(e)`: POST エンドポイント
  - `action: 'record'` — 単一打刻データ保存
  - `action: 'sync'` — 複数データ一括同期
  - `action: 'delete'` — 打刻データ削除
  - `action: 'saveSettings'` — 設定保存
- シート命名規則: `スタッフ名_YYYYMM` (例: `草野_202604`)
- 設定シート: `AppSettings` (staffList, adminPin を保存)

## timecard.html の構造

約1500行の単一HTMLファイル。React JSXコードが `<script type="text/babel">` 内に記述されている。

### 主要な変数・関数
- `EXCEL_TEMPLATE_B64` — Base64エンコードされたExcelテンプレート
- `useAppData()` — カスタムフック：全状態管理 (staffList, timeRecords, adminPin, gasUrl, 同期関数群)
- `sendToGAS()` / `deleteFromGAS()` / `syncFromGAS()` — GAS連携
- `saveSettingsToGAS()` / `syncSettingsFromGAS()` — 設定同期
- `EditRecordModal` — 打刻修正モーダル
- `MonthlyTimeEditor` — 月次カレンダー形式の履歴表示
- `AdminDashboard` — 管理者画面
- `handleExportWithTemplate()` — Excelエクスポート

### 外注者機能（最近追加）
- スタッフに `isOutsourced: true` フラグ
- `hourlyRate`（時給）, `dailyTransport`（交通費/日）の属性
- 外注者用Excelテンプレート（通常テンプレートと同一レイアウト、賄いカウント除外）
- 給与計算ロジック内蔵

## デプロイ方法

1. コードを修正
2. `git add . && git commit -m "修正内容" && git push origin main`
3. Vercel が自動デプロイ（GitHub連携済み）

## 既知の課題・注意点

- **Vercelに移行済み**: Netlifyのビルドクレジット切れのため、Vercelに移行済み。vercel.jsonでルート→timecard.htmlのリライト設定を実施
- **GAS URL**: 各デバイスのローカルストレージに保存されている。新しいデバイスでは管理者画面からGAS URLを設定する必要がある
- **GAS認証トークン**: GitHubリポジトリの公開設定変更時に新しいトークンが必要になった場合は、GASの再デプロイが必要
- **同期の競合**: ローカル更新後8秒間はクラウドからの同期をブロックするロジックが入っている。この値を安易に変更しないこと
- **Excelテンプレート**: `EXCEL_TEMPLATE_B64` はタイムカード入力空バージョン.xlsx をBase64化したもの。テンプレート変更時はxlsxを編集してから再エンコードする

## 開発の進め方のルール

1. **小さな変更を積み重ねる** — 一度に大量のコードを書き換えない
2. **既存のコメントを保持する** — 修正理由や経緯がコメントに残っている
3. **変更後は動作確認** — ブラウザでtimecard.htmlを開いて確認
4. **git commitのメッセージは日本語** — 例: `fix: 出勤打刻時に賄いがリセットされるバグを修正`
5. **余計なファイルを追加しない** — package.json, node_modules, .github/workflows/ 等は不要
