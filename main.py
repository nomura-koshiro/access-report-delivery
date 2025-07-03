from pathlib import Path
from typing import Optional

from selenium.webdriver.chrome.webdriver import WebDriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import TimeoutException

from config import WAIT_TIMEOUT
from faro_biz_automation import setup_driver, login_to_faro_biz, create_and_download_report
from powerpoint_converter import convert_ppt_to_pdf
from mailer import send_email_with_attachment, send_notification_email

def cleanup_files(downloaded_pptx: Optional[Path], converted_pdf: Optional[Path]):
    """ダウンロードしたPowerPointファイルと変換したPDFファイルを削除します。"""
    print("一時ファイルを削除しています...")
    if downloaded_pptx and downloaded_pptx.exists():
        downloaded_pptx.unlink()
        print(f"削除完了: {downloaded_pptx.name}")
    if converted_pdf and converted_pdf.exists():
        converted_pdf.unlink()
        print(f"削除完了: {converted_pdf.name}")
    print("一時ファイルの削除が完了しました。")

def _handle_success(converted_pdf: Path):
    """処理成功時のメール送信と通知を行います。"""
    send_email_with_attachment(converted_pdf)
    send_notification_email(
        "アクセスレポート送信完了通知",
        f"アクセスレポート({converted_pdf.name})の送信が完了しました。"
    )

def _handle_failure(exception: Exception, error_type: str):
    """処理失敗時の通知を行います。"""
    send_notification_email(
        f"アクセスレポート処理失敗通知 ({error_type})",
        f"アクセスレポートの処理中に{error_type}が発生しました: {exception}"
    )

def run_automation_flow(
    driver: WebDriver, wait: WebDriverWait
) -> Optional[Path]:
    """メインの自動化処理（ログイン、レポート作成、PDF変換）を実行します。"""
    login_to_faro_biz(driver, wait)
    downloaded_pptx = create_and_download_report(driver, wait)
    converted_pdf = convert_ppt_to_pdf(downloaded_pptx)
    return converted_pdf

def main():
    """アクセスレポート自動化のメイン関数。"""
    driver: Optional[WebDriver] = None
    downloaded_pptx: Optional[Path] = None  # cleanup_filesのために保持
    converted_pdf: Optional[Path] = None  # cleanup_filesのために保持
    try:
        driver = setup_driver()
        wait = WebDriverWait(driver, WAIT_TIMEOUT)

        converted_pdf = run_automation_flow(driver, wait)

        if converted_pdf:
            _handle_success(converted_pdf)
        else:
            _handle_failure(Exception("PDF変換失敗"), "PDF変換失敗")

        print("すべての処理が正常に完了しました。")

    except TimeoutException as e:
        _handle_failure(e, "タイムアウト")
    except Exception as e:
        _handle_failure(e, "予期せぬエラー")
    finally:
        if driver:
            driver.quit()
            print("ブラウザを閉じました。")
        cleanup_files(downloaded_pptx, converted_pdf)

if __name__ == "__main__":
    main()