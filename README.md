# アクセスレポート自動配信システム

## プロジェクト概要

このプロジェクトは、特定のウェブサイト（FARO Biz）からアクセスレポートを自動的にダウンロードし、PDFに変換し、指定されたメールアドレスに配信するシステムです。また、処理の完了やエラー発生時には通知メールを送信し、一時ファイルを自動的にクリーンアップします。

## 機能

本システムは以下の主要な機能を提供し、それぞれが独立したモジュールとして構成されています。

1.  **FARO Bizへのログイン**: `faro_biz_automation.py` が担当。環境変数または設定ファイルに設定された認証情報を使用してFARO Bizにログインします。
2.  **アクセスレポートのダウンロード**: `faro_biz_automation.py` が担当。指定されたウェブサイト（松和園）のアクセスレポートをPowerPoint形式でダウンロードします。
3.  **PowerPointからPDFへの変換**: `powerpoint_converter.py` が担当。ダウンロードしたPowerPointファイルをPDF形式に変換します。
4.  **PDFレポートのメール送信**: `mailer.py` が担当。変換されたPDFレポートを、`config.py` で設定された宛先にメールで送信します。メールの件名と本文は自動生成され、前月の情報が含まれます。
5.  **処理完了通知**: `mailer.py` が担当。レポートのメール送信が完了したことを、`config.py` で設定されたメールアドレスに通知します。
6.  **エラー通知**: `mailer.py` が担当。処理中にタイムアウトや予期せぬエラーが発生した場合、`config.py` で設定されたメールアドレスにエラー内容を通知します。
7.  **一時ファイルのクリーンアップ**: `main.py` が担当。ダウンロードしたPowerPointファイルと変換されたPDFファイルを自動的に削除します。

## 動作環境

*   Windows OS (PDF変換に `pywin32` が必要)
*   Python 3.x
*   `uv` (パッケージ管理ツール)
*   Google Chrome (Selenium WebDriver用)

## セットアップ

1.  **リポジトリのクローン**:
    ```bash
    git clone https://github.com/your-repo/access-report-delivery.git
    cd access-report-delivery
    ```

2.  **Python環境のセットアップ**:
    `uv` を使用して依存関係をインストールします。
    ```bash
    uv sync
    ```
    `pywin32` がインストールされていない場合は、以下のコマンドでインストールしてください。
    ```bash
    uv pip install pywin32
    ```

3.  **Chrome WebDriverの準備**:
    Google Chromeがインストールされていることを確認してください。Seleniumが自動的に適切なWebDriverをダウンロードして使用します。

4.  **環境変数の設定**:
    以下の環境変数を設定してください。これらの情報は `config.py` で読み込まれ、設定されていない場合はデフォルト値が使用されます。本番環境では、これらの情報を環境変数として設定することを強く推奨します。

    *   `FARO_BIZ_EMAIL`: FARO Bizのログインメールアドレス
    *   `FARO_BIZ_PASSWORD`: FARO Bizのログインパスワード
    *   `EMAIL_SENDER`: 送信元メールアドレス（Gmailの場合、2段階認証を有効にし、アプリパスワードを使用してください）
    *   `EMAIL_PASSWORD`: 送信元メールアドレスのパスワード（Gmailの場合、アプリパスワード）
    *   `REPORT_RECIPIENT`: レポートの送信先メールアドレス
    *   `NOTIFICATION_RECIPIENT`: 通知メールの送信先メールアドレス
    *   `SMTP_SERVER`: SMTPサーバーのアドレス (例: `smtp.gmail.com`)
    *   `SMTP_PORT`: SMTPサーバーのポート番号 (例: `587`)

    **Windowsの場合（コマンドプロンプト）**:
    ```cmd
    set FARO_BIZ_EMAIL=your_faro_biz_email@example.com
    set FARO_BIZ_PASSWORD=your_faro_biz_password
    set EMAIL_SENDER=your_gmail_address@gmail.com
    set EMAIL_PASSWORD=your_gmail_app_password
    set REPORT_RECIPIENT=report_recipient@example.com
    set NOTIFICATION_RECIPIENT=notification_recipient@example.com
    set SMTP_SERVER=smtp.gmail.com
    set SMTP_PORT=587
    ```

## 実行方法

プロジェクトのルートディレクトリで以下のコマンドを実行します。

```bash
uv run python main.py
```

**文字化け対策（Windowsコマンドプロンプト）**:
出力の文字化けが発生する場合は、以下のコマンドを先に実行してからスクリプトを実行してください。

```cmd
chcp 65001
uv run python main.py
```

## 開発者

webLinx 野村 幸志朗