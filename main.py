import sys
from pathlib import Path
from typing import Optional
import os

from selenium import webdriver
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.webdriver import WebDriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select, WebDriverWait

_win32com_available = False
if sys.platform == "win32":
    try:
        import win32com.client
        _win32com_available = True
    except ImportError:
        print("警告: pywin32 ライブラリがインストールされていません。PDF変換機能は利用できません。")
        print("      PDF変換が必要な場合は、'pip install pywin32' または 'uv pip install pywin32' を実行してください。")

import os

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
REPORT_URL = f"{BASE_URL}/report"

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


def setup_driver() -> WebDriver:
    """Chrome WebDriverのセットアップと初期化を行います。"""
    print("WebDriverをセットアップしています...")
    chrome_options = Options()
    # ダウンロードディレクトリを指定
    prefs = {"download.default_directory": str(DOWNLOAD_DIR)}
    chrome_options.add_experimental_option("prefs", prefs)

    # downloadsフォルダがなければ作成
    DOWNLOAD_DIR.mkdir(exist_ok=True)

    return webdriver.Chrome(options=chrome_options)


def login_to_faro_biz(driver: WebDriver, wait: WebDriverWait):
    """FARO Bizにログインします。"""
    print("FARO Bizにアクセスし、ログインしています...")
    driver.get(LOGIN_URL)
    wait.until(EC.presence_of_element_located((By.NAME, "email"))).send_keys(
        LOGIN_EMAIL
    )
    driver.find_element(By.NAME, "password").send_keys(LOGIN_PASSWORD)
    # IDセレクタを使い、クリック可能になるまで待機
    wait.until(EC.element_to_be_clickable((By.ID, "btn-login"))).click()
    # ログイン後のリダイレクトを待機
    wait.until(EC.url_changes(LOGIN_URL))
    print("ログインしました。")


def create_and_download_report(driver: WebDriver, wait: WebDriverWait) -> Path:
    """レポートを作成し、PowerPointファイルをダウンロードします。"""
    print("レポートページにアクセスしています...")

    print("「レポートの新規作成」ボタンをクリックします。")
    wait.until(EC.element_to_be_clickable((By.XPATH, NEW_REPORT_BUTTON_XPATH))).click()

    print("レポート作成ページへの遷移を待ちます。")
    wait.until(EC.url_to_be(f"{BASE_URL}/report/create"))

    print(f"ウェブサイト「{WEBSITE_NAME}」を選択します。")
    # Selectクラスを使ってドロップダウンを操作
    select_element = wait.until(
        EC.presence_of_element_located((By.ID, "website_id"))
    )
    select = Select(select_element)
    select.select_by_visible_text(WEBSITE_NAME)

    print("「レポートを作成」ボタンをクリックします。")
    create_report_button = wait.until(
        EC.element_to_be_clickable((By.XPATH, CREATE_REPORT_BUTTON_XPATH))
    )
    # JavaScriptを使ってクリックを試みる
    driver.execute_script("arguments[0].click();", create_report_button)

    print("ダウンロードボタンが表示されるのを待っています...")
    download_link = wait.until(
        EC.element_to_be_clickable((By.XPATH, DOWNLOAD_LINK_XPATH))
    )

    # ダウンロード前に既存のpptxファイルを削除
    for f in DOWNLOAD_DIR.glob("*.pptx"):
        f.unlink()

    print("ダウンロードボタンをクリックします。")
    download_link.click()

    return wait_for_download_completion(wait)


def wait_for_download_completion(wait: WebDriverWait) -> Path:
    """ファイルのダウンロード完了を待ちます。"""
    print("ファイルのダウンロードを待っています...")

    # .crdownloadファイルがなくなるまで待機
    try:
        WebDriverWait(wait.driver, DOWNLOAD_TIMEOUT).until(lambda driver: not any(DOWNLOAD_DIR.glob("*.crdownload")))
    except TimeoutException:
        raise TimeoutException("ダウンロード中のファイル(.crdownload)の消滅がタイムアウトしました。")

    # .pptxファイルがダウンロードされるまで待機
    try:
        WebDriverWait(wait.driver, DOWNLOAD_TIMEOUT).until(lambda driver: any(DOWNLOAD_DIR.glob("*.pptx")))
    except TimeoutException:
        raise TimeoutException("PowerPointファイルのダウンロードがタイムアウトしました。")

    # 作成時刻が最新のpptxファイルを取得
    downloaded_file = max(
        DOWNLOAD_DIR.glob("*.pptx"), key=lambda f: f.stat().st_ctime
    )
    print(f"ダウンロード完了: {downloaded_file.name}")
    return downloaded_file


def convert_ppt_to_pdf(pptx_path: Path):
    """PowerPointファイルをPDFに変換します。"""
    if not _win32com_available:
        print("PDF変換はWindows環境でのみサポートされており、pywin32ライブラリが必要です。")
        return

    powerpoint = None
    presentation = None
    try:
        print(f"'{pptx_path.name}' をPDFに変換しています...")
        powerpoint = win32com.client.Dispatch("PowerPoint.Application")
        # powerpoint.Visible = 1  # デバッグ時以外は非表示が望ましい

        pdf_path = pptx_path.with_suffix(".pdf")

        presentation = powerpoint.Presentations.Open(str(pptx_path.resolve()))
        presentation.SaveAs(str(pdf_path.resolve()), 32)  # 32はPDF形式(ppSaveAsPDF)
        print(f"PDFに変換完了: {pdf_path.name}")

    except Exception as e:
        print(f"PDF変換中にエラーが発生しました: {e}")
    finally:
        if presentation:
            presentation.Close()
        if powerpoint:
            powerpoint.Quit()


def main():
    """メイン処理"""
    driver: Optional[WebDriver] = None
    try:
        driver = setup_driver()
        wait = WebDriverWait(driver, WAIT_TIMEOUT)

        login_to_faro_biz(driver, wait)
        downloaded_pptx = create_and_download_report(driver, wait)
        convert_ppt_to_pdf(downloaded_pptx)

        print("すべての処理が正常に完了しました。")

    except TimeoutException as e:
        print(f"処理がタイムアウトしました: {e}")
    except Exception as e:
        print(f"予期せぬエラーが発生しました: {e}")
    finally:
        if driver:
            driver.quit()
            print("ブラウザを閉じました。")

if __name__ == "__main__":
    main()
