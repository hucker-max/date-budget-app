# Date Budget App

デート費用と予算管理の単一ファイル Web アプリです。

## ファイル

- `index.html`

## 使い方

ローカル確認:

```bash
cd "/Users/okadaigo/Documents/New project"
python3 -m http.server 8000
```

ブラウザで `http://localhost:8000/index.html` を開きます。

## GitHub Pages 公開

1. GitHub に新しいリポジトリを作成
2. このフォルダで Git 初期化済みなら、そのまま以下を実行

```bash
git add index.html .gitignore README.md
git commit -m "Add shared date budget app"
git branch -M main
git remote add origin <YOUR_GITHUB_REPO_URL>
git push -u origin main
```

3. GitHub のリポジトリ設定で `Settings > Pages`
4. `Deploy from a branch`
5. `main` / `/ (root)` を選択
6. 公開 URL を果林さんと共有

## 共有の考え方

- 画面公開先: GitHub Pages
- データ保存先: Google スプレッドシート + Apps Script
- 2人とも同じ URL を開いて入力する

## Apps Script

Web アプリ URL を `同期・設定` タブに設定して使います。

