# クラウドワークス領収書自動ダウンロードツール

クラウドワークス (プロジェクト形式) の支払い一覧ページから領収書を一括でダウンロードするPythonスクリプトです。手動ログインに対応し、未発行の領収書も自動的に発行・ダウンロードします。

注意!: 一度発行された領収書は内容が変更できません。また、生成AIを用い、手早くコードの作成を行いました。一部余分なコードが混ざっていますが使用に問題はないと思われます。

## 特徴

- 手動ログイン
- 未発行の領収書を自動的に発行し、一括ダウンロード
- 発行済み領収書の一括ダウンロード
- PDFファイル名の自動整理（通し番号付与）
- エラー時の自動リトライと手動介入オプション
- 詳細なログ記録
- ヘッダー要素を自動的に非表示化してクリーンなPDFを生成

## 必要要件

- Python 3.8以上
- Google Chrome
- ChromeDriver（自動的にインストールされいるはず）

## インストール

1. リポジトリをクローン：
```zsh
git clone https://github.com/yamanori99/crowdworks-receipt-downloader.git
cd crowdworks-receipt-downloader
```

2. 必要なパッケージをインストール：
```zsh
pip install -r requirements.txt
```

## 使用方法

1. スクリプトを実行：
```zsh
python3 receipt_download_manual_login.py
```

2. ブラウザが起動したら、手動でログインしてください。

3. ログイン後、自動的に領収書のダウンロードが開始されます。

4. ダウンロードした領収書は `receipts_YYYYMMDD_HHMMSS` フォルマットのフォルダに保存されます。

## 設定オプション

コマンドライン引数で以下のオプションを指定できます：

- `--download-dir`: ダウンロード先ディレクトリを指定
- `--log-file`: ログファイル名を指定
- `--config`: 設定ファイルのパスを指定

例：
```zsh
python3 receipt_download_manual_login.py --download-dir ./my_receipts --log-file my_log.log
```

## エラー処理

- ダウンロード失敗時は自動的にリトライします
- 3回リトライしても失敗した場合は、以下のオプションが表示されます：
  1. 再試行
  2. この領収書をスキップ
  3. 処理を中止

## 注意事項

- このスクリプトは2025年3月10日現在のクラウドワークスのウェブサイト構造に対応しています
- 大量の領収書を一括ダウンロードする際は、サーバーに負荷をかけないよう適度な間隔を設けています
- ログイン情報は保存されません（セキュリティ上の理由により手動ログインを採用）

## ログ

- 詳細なログは `receipt_download_manual.log` に記録されます
- エラー発生時のスクリーンショットは自動的に保存されます

## トラブルシューティング

1. ログインできない場合：
   - ブラウザを手動で操作してログインしてください

2. PDFが保存されない場合：
   - スクリーンショットとして自動的に保存されます
   - 手動での保存オプションが表示されます

3. ページ読み込みエラーの場合：
   - 自動的にリトライされます
   - 必要に応じて手動介入が可能です

