# SkyseaAutoDataUpload

Skyseaのデータアップロードを自動でアップロードするスクリプト

# Requirement

* urllib3 1.26.2
* PyYAML 5.3.1
* pysmb 1.2.6
* pywin32 300
* selenium 3.141.0
* pywinauto 0.6.8

# What is this
SMBサーバーに保存されたcsvファイルをxlsxファイルに変換し、
パスワードをかけて、Redmineのナレッジベースにアップロードするスクリプト
(ナレッジベースはプラグイン)

# Installation

```bash
pip install -r requirements.txt
```

# Usage

1. 下記を実行し必要なモジュールをインストールする
```bash
pip install -r requirements.txt
```
2. Settings.yaml のUserName, Passwordをそれぞれ入力する。

3. 下記を実行する
```bash
python main.py 
```
4. [Windows セキュリティの重要な警告]というポップアップが出るので、アクセスを許可する（初回実行時のみ）

# Note

* ホスト情報やログイン情報が変更された際は、併せて Setting.yaml の内容も変更すること。
"# FileUploadAutomation" 
