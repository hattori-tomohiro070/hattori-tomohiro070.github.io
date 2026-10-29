# 服部智大 個人ウェブサイト

掲載内容はすべて **`website_content.docx`** で管理する。
docx を編集して push すると、GitHub Actions が `data.js` を生成してサイトを更新する。

## 更新の流れ
1. `website_content.docx` を Word で編集して保存（書き方のルールは docx の冒頭に記載）
2. push する
   ```sh
   git add website_content.docx
   git commit -m "業績を更新"
   git push
   ```
3. 数分後に `https://<ユーザー名>.github.io/` に反映される
   （進行状況はリポジトリの Actions タブで確認できる。赤い×の場合はログに原因が日本語で出る）

## 手元で確認したいとき
```sh
pip3 install python-docx        # 初回のみ
python3 scripts/build_data.py   # docx → data.js
open index.html
```

## 初回セットアップ
1. GitHub で `<ユーザー名>.github.io` という名前の public リポジトリを作る
2. このフォルダを push する
   ```sh
   git init && git add . && git commit -m "Initial site"
   git branch -M main
   git remote add origin https://github.com/<ユーザー名>/<ユーザー名>.github.io.git
   git push -u origin main
   ```
3. リポジトリの Settings → Pages → Build and deployment の Source を **GitHub Actions** にする

## ファイル構成
| ファイル | 役割 |
|---|---|
| `website_content.docx` | 掲載情報の原本（ここだけ編集する） |
| `index.html` | ページのデザイン |
| `scripts/build_data.py` | docx → `data.js` の変換 |
| `.github/workflows/deploy.yml` | push 時の自動変換・公開 |
| `data.js` | 自動生成物（git 管理外） |
