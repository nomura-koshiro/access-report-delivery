import os
from pathlib import Path

# --- 定数定義 ---

# ログイン情報 (環境変数などから取得することを強く推奨します)
LOGIN_EMAIL = os.getenv("FARO_BIZ_EMAIL", "nomura.koshiro@gmail.com")  # ★ご自身のメールアドレスに変更してください
LOGIN_PASSWORD = os.getenv("FARO_BIZ_PASSWORD", "gobbledyGook@1103")  # ★ご自身のパスワードに変更してください

if os.getenv("FARO_BIZ_EMAIL") is None or os.getenv("FARO_BIZ_PASSWORD") is None:
    print("警告: FARO_BIZ_EMAIL または FARO_BIZ_PASSWORD 環境変数が設定されていません。")
    print("      ハードコードされた認証情報を使用します。本番環境では環境変数を使用することを強く推奨します。")

# URL
BASE_URL = "https://app.faro-biz.com"
LOGIN_URL = f"{BASE_URL}/login"

# ディレクトリ設定
# スクリプトの場所に 'downloads' という名前のディレクトリを作成
DOWNLOAD_DIR = Path(__file__).parent / "downloads"

# Selenium関連
WAIT_TIMEOUT = 60  # 秒
DOWNLOAD_TIMEOUT = 300  # 秒 (現在、wait_for_download_completion関数では直接使用されていません)

# レポート作成設定
WEBSITE_NAME = "松和園"

# XPathセレクタ
NEW_REPORT_BUTTON_XPATH = '//a[@href="https://app.faro-biz.com/report/create"]'
CREATE_REPORT_BUTTON_XPATH = '//form//button[contains(., "レポートを作成")]'
# ダウンロードリンクまたはボタンのXPath
DOWNLOAD_LINK_XPATH = (
    '//a[contains(., "ダウンロード") or button[contains(., "ダウンロード")]]'
)

# --- メール設定 ---
SMTP_SERVER = os.getenv("SMTP_SERVER", "smtp.gmail.com")
SMTP_PORT = int(os.getenv("SMTP_PORT", "587"))
EMAIL_SENDER = os.getenv("EMAIL_SENDER", "nomura.koshiro@gmail.com")
EMAIL_PASSWORD = os.getenv("EMAIL_PASSWORD", "xthv uvav mrai froe")  # ★環境変数で設定することを推奨
REPORT_RECIPIENT = os.getenv("REPORT_RECIPIENT", "nagomi@showaen.or.jp")
NOTIFICATION_RECIPIENT = os.getenv("NOTIFICATION_RECIPIENT", "nomura.koshiro@gmail.com")

if os.getenv("EMAIL_SENDER") is None or os.getenv("EMAIL_PASSWORD") is None:
    print("警告: EMAIL_SENDER または EMAIL_PASSWORD 環境変数が設定されていません。")
    print("      ハードコードされたメール認証情報を使用します。本番環境では環境変数を使用することを強く推奨します。")

# PowerPoint PDF保存形式の定数
PPT_SAVE_AS_PDF = 32
