from pathlib import Path

from selenium import webdriver
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.webdriver import WebDriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select, WebDriverWait

from config import (
    BASE_URL,
    CREATE_REPORT_BUTTON_XPATH,
    DOWNLOAD_DIR,
    DOWNLOAD_LINK_XPATH,
    DOWNLOAD_TIMEOUT,
    LOGIN_EMAIL,
    LOGIN_PASSWORD,
    LOGIN_URL,
    NEW_REPORT_BUTTON_XPATH,
    WEBSITE_NAME,
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
    wait.until(EC.presence_of_element_located((By.NAME, "email"))).send_keys(LOGIN_EMAIL)
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
    wait.until(EC.element_to_be_clickable(
        (By.XPATH, NEW_REPORT_BUTTON_XPATH))).click()

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

    return wait_for_download_completion(driver)


def wait_for_download_completion(driver: WebDriver) -> Path:
    """ファイルのダウンロード完了を待ちます。"""
    print("ファイルのダウンロードを待っています...")

    # .crdownloadファイルがなくなるまで待機
    try:
        WebDriverWait(driver, DOWNLOAD_TIMEOUT).until(
            lambda d: not any(DOWNLOAD_DIR.glob("*.crdownload")))
    except TimeoutException:
        raise TimeoutException(
            "ダウンロード中のファイル(.crdownload)の消滅がタイムアウトしました。")

    # .pptxファイルがダウンロードされるまで待機
    try:
        WebDriverWait(driver, DOWNLOAD_TIMEOUT).until(
            lambda d: any(DOWNLOAD_DIR.glob("*.pptx")))
    except TimeoutException:
        raise TimeoutException(
            "PowerPointファイルのダウンロードがタイムアウトしました。")

    # 作成時刻が最新のpptxファイルを取得
    downloaded_file = max(
        DOWNLOAD_DIR.glob("*.pptx"), key=lambda f: f.stat().st_ctime
    )
    print(f"ダウンロード完了: {downloaded_file.name}")
    return downloaded_file
