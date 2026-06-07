# ココナラ販売実績スクレイパー

ココナラ コンテンツマーケットの「ツール」検索結果から、販売実績が一定数を超える商品を抽出して Excel に出力するスクレイパー。

## 概要

- **目的**: ココナラ コンテンツマーケットで「ツール」を検索し、各商品のタイトル・販売実績数を取得。販売実績 3 件超のものを Excel に出力。
- **実装**: Python + [DrissionPage](https://github.com/g1879/DrissionPage)（実 Chromium / Edge ベースの stealth ブラウザ自動化）
- **実行環境**: GitHub Actions（ubuntu-latest）で定期/手動実行

## ファイル構成

- `scraper.py` - メインスクレイパー（DrissionPage 版）
- `requirements.txt` - Python 依存関係
- `.github/workflows/scrape.yml` - GitHub Actions ワークフロー
- `scraper_cloakbrowser_backup.py` - 旧 CloakBrowser 版のバックアップ（参考用、ローカル実行専用）
- `requirements_cloakbrowser_backup.txt` - 旧依存関係

## 使い方

### GitHub Actions（推奨）

1. このリポジトリをフォーク
2. Actions タブ → `Scrape Coconala Tools` → `Run workflow`
3. 完了後、Artifacts セクションから `coconala-tools-xlsx` をダウンロード

### ローカル実行

```bash
# 依存関係インストール
pip install -r requirements.txt

# 全フェーズ実行（collect → fetch → excel）
python scraper.py

# フェーズ別実行
python scraper.py --collect
python scraper.py --fetch --workers 3
python scraper.py --excel
```

## 出力

- `output/coconala_tools.xlsx` - 販売実績 3 件超の商品を販売数降順でソート
  - 列: No. / タイトル / 販売実績数 / URL
- `output/urls_cache.json` - 収集した商品 URL リスト
- `output/data_cache.json` - フェッチ結果（{url: {title, sales_count}}）

## 注意事項

- ココナラのサーバーへの負荷軽減のため、リクエスト間隔は 4±2 秒
- 同一プロセス内で複数ページを取得する場合、最大 3 ワーカーで並列化
- 403 Forbidden 発生時は最大 3 回まで自動リトライ
- 並列ワーカー数は `NUM_WORKERS` (デフォルト 3) で調整可能
