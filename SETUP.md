# 今CoCo（いまココ）デプロイ＆環境構築ガイド

構成：GitHub Pages（index.html）＋ GAS（Code.gs）＋ スプレッドシート ＋ Google Drive ＋ Gemini API（すべて無料枠）

---

## 1. スプレッドシート & Driveフォルダの作成

1. Googleドライブで新しいスプレッドシートを作成し、名前を「今CoCo_DB」にする。
2. メニュー「拡張機能」→「Apps Script」を開く（スプレッドシートに紐づいたGASになります）。
3. 最初からある `コード.gs` の中身をすべて消し、`Code.gs` の内容を貼り付けて保存。
4. 関数選択で `setup` を選び「実行」。初回は権限の承認画面が出るので許可する
   （「このアプリはGoogleで確認されていません」→「詳細」→「（安全ではないページ）に移動」）。

`setup()` が自動で行うこと：

| 処理 | 内容 |
|---|---|
| シート作成 | `events` / `users` / `requests` の3シートとヘッダー行 |
| Driveフォルダ作成 | マイドライブ直下に「今CoCo_media」 |
| スクリプトプロパティ | `SPREADSHEET_ID` / `DRIVE_FOLDER_ID` / `PEPPER` を自動保存 |
| トリガー | `expireEvents` を10分おきに実行（期限切れ→`expired`、メディアをゴミ箱へ） |

IDの確認方法（手動で変えたい場合）：
- スプレッドシートID：URL `https://docs.google.com/spreadsheets/d/【ここ】/edit`
- フォルダID：フォルダを開いたURL `https://drive.google.com/drive/folders/【ここ】`

既存フォルダを使いたい場合は、`setup()` 実行前にスクリプトプロパティ `DRIVE_FOLDER_ID` に入れておけばそれを使います。

> `events` シートには指示書の17列に加え、`hashtags`（AI生成タグ）と `media_file_id`（削除用のDriveファイルID）の2列を末尾に追加しています。

---

## 2. Gemini APIキーの取得と保存

1. https://aistudio.google.com/apikey を開き、Googleアカウントでログイン。
2. 「APIキーを作成」→ 表示されたキーをコピー。
3. Apps Scriptエディタ左の「⚙ プロジェクトの設定」→ 下部「スクリプト プロパティ」→「スクリプト プロパティを追加」。

| プロパティ | 値 | 必須 |
|---|---|---|
| `GEMINI_API_KEY` | コピーしたキー | 必須 |
| `GEMINI_MODEL` | 例：`gemini-3.8-flash`（未設定ならこれ） | 任意 |
| `MODERATION_FAIL_OPEN` | `true` にするとGemini障害時もチェックなしで投稿を通す | 任意 |
| `TRASH_MEDIA_ON_EXPIRE` | `false` にすると期限切れでも写真・動画を残す | 任意 |

4. エディタで `testGemini` を実行し、ログに `"safe": true` と絵文字・キャッチコピーが出ればOK。

> 既定モデルは `gemini-3.8-flash` です。無料枠の上限に当たる場合は `gemini-3.5-flash-lite` などに切り替えてください。モデルを変えるときはコードではなく `GEMINI_MODEL` を書き換えるだけで済みます。

---

## 3. GAS をウェブアプリとしてデプロイ

1. 右上「デプロイ」→「新しいデプロイ」。
2. 歯車アイコン →「ウェブアプリ」を選択。
3. 設定：
   - 説明：`今CoCo v1`
   - 次のユーザーとして実行：**自分**
   - アクセスできるユーザー：**全員**
4. 「デプロイ」→ 表示された **ウェブアプリURL**（`https://script.google.com/macros/s/.../exec`）をコピー。
5. ブラウザで `【URL】?action=ping` を開き、`{"ok":true,"app":"今CoCo",...}` が出れば成功。

### ⚠ Code.gs を修正したとき
保存しただけでは本番に反映されません。
「デプロイ」→「デプロイを管理」→ 鉛筆アイコン →「バージョン：**新しいバージョン**」→「デプロイ」。
この手順ならURLは変わらないので、index.html の書き換えは不要です。

---

## 4. GitHub Pages に公開

1. GitHub で新しいリポジトリ `imacoco` を作成（Public）。
2. `index.html` の上の方にある次の行を、手順3のURLに書き換える：
   ```js
   const GAS_API_URL = 'https://script.google.com/macros/s/XXXXXXXXXXXXXXXXXXXXXXXX/exec';
   ```
3. `index.html` をリポジトリ直下にアップロードしてコミット。
4. リポジトリの「Settings」→「Pages」→ Source：`Deploy from a branch`、Branch：`main` / `(root)` →「Save」。
5. 1〜2分後、`https://kenken6291.github.io/imacoco/` で公開されます。

位置情報は HTTPS でしか動かないため、スマホで試すときは GitHub Pages のURLで開いてください（ファイルを直接開くと現在地が取れません）。

---

## 5. 動作確認チェックリスト

- [ ] 地図が表示され、現在地の青い点が出る
- [ ] 右上「ニックネーム登録」で登録 → `users` シートに行が増える
- [ ] 「今ココ！」→ 3ステップで投稿 → 「Geminiが投稿をチェック中...」→ 地図にピンが落ちてくる
- [ ] `events` シートに行が増え、`summary` / `icon` / `hashtags` がAI生成されている
- [ ] 写真付き投稿で「今CoCo_media」フォルダにファイルができ、詳細に写真が表示される
- [ ] 別アカウントで「参加したい」→ 投稿者のマイページに参加リクエストが表示される
- [ ] 「シェア」で `?e=投稿ID` 付きURLが共有され、開くとその投稿の詳細が開く
- [ ] 「今すぐ終了」で地図から消え、`status` が `cancelled` になる
- [ ] 終了時刻を過ぎると（最大10分以内に）`status` が `expired` になる

---

## 6. 仕様メモ

| 項目 | 値 |
|---|---|
| 1人あたりの同時公開数 | 5件（`CONFIG.MAX_ACTIVE_POSTS_PER_USER`） |
| 表示時間の上限 | 72時間（`CONFIG.MAX_EXPIRE_HOURS`） |
| 写真 | 端末側で長辺1280px・JPEGに圧縮して送信、Geminiが画像もチェック |
| 動画 | 15MBまで（端末側制限）。Driveプレーヤーで再生。Geminiチェックは本文のみ |
| ログイン | ニックネーム＋パスワード（SHA-256×100回＋ソルト＋ペッパー）、5回失敗で10分ロック、トークン90日有効 |
| 一覧キャッシュ | 20秒（CacheService）。投稿・削除・参加時に即クリア |
| 自動更新 | 画面表示中は60秒ごと、アプリに戻ったとき |

---

## 7. トラブルシューティング

| 症状 | 原因と対処 |
|---|---|
| 「サーバーの応答を読めません」 | デプロイの「アクセスできるユーザー」が「全員」になっていない。または古いURLを使っている |
| 修正が反映されない | 「新しいバージョン」で再デプロイしていない（手順3の⚠） |
| 「AIチェックが混み合っています」 | Gemini無料枠のレート制限。1分待って再投稿 |
| 「GEMINI_API_KEY が未設定です」 | スクリプトプロパティ名の綴りを確認 |
| 写真が表示されない | Driveの共有が制限されている（Google Workspaceの組織アカウントでは外部共有不可の場合あり）。個人のGmailアカウントで運用する |
| 動画が「処理中」と出る | Drive側の変換待ち。数分後に再生できるようになる |
| iPhoneで写真が読めない | HEIC形式の一部。カメラ設定「互換性優先」にするか撮影ボタンから直接撮る |
| 長押しで投稿できない | 一部ブラウザで長押しがコピーメニューになる。右下「今ココ！」→「地図で選ぶ」を使う |
