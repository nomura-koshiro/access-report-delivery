import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from pathlib import Path
from datetime import datetime

from config import (
    EMAIL_SENDER,
    EMAIL_PASSWORD,
    REPORT_RECIPIENT,
    NOTIFICATION_RECIPIENT,
    SMTP_SERVER,
    SMTP_PORT,
)

def _get_report_email_body(target_month: int) -> str:
    """レポート送付メールの本文を生成します。"""
    return f"""
松和園　吉永様

お世話になっております、webLinx野村です。
掲題の件ですが、2025年{target_month}月のアクセスレポートを作成しましたので、送付いたします。

以上です、よろしくお願いいたします。

□■━━━━━━━━━━━━━━━━━━━━━━━━━━━━

株式会社webLinx

    代表取締役  野村  幸志朗

    〒810-0001
        福岡県福岡市中央区天神1丁目1番-1号
        fabbit Global Gateway “ACROS Fukuoka”

    Email : nomura.koshiro@weblinx.jp
    Tel : 080-3855-2456
    URL : https://weblinx.jp/

━━━━━━━━━━━━━━━━━━━━━━━━━━━━■□
"""

def send_email_with_attachment(pdf_path: Path):
    """指定されたPDFファイルをメールで送信します。"""
    print(f"'{pdf_path.name}' をメールで送信しています...")

    if not pdf_path.exists():
        print(f"エラー: PDFファイルが見つかりません: {pdf_path}")
        return
    print(f"PDFファイルが存在します: {pdf_path}")

    msg = MIMEMultipart()
    msg["From"] = EMAIL_SENDER
    msg["To"] = REPORT_RECIPIENT
    msg["Sender"] = "nomura.koshiro@weblinx.jp"
    msg["Bcc"] = EMAIL_SENDER  # 送信者自身にBccでコピーを送る

    # 現在の月-1を計算
    current_month = datetime.now().month
    target_month = current_month - 1 if current_month > 1 else 12

    msg["Subject"] = f"2025年{target_month}月のレポートの送付"

    body = _get_report_email_body(target_month)
    msg.attach(MIMEText(body, "plain"))

    print(f"PDFファイルを開いています: {pdf_path}")
    with open(pdf_path, "rb") as f:
        attach = MIMEApplication(f.read(), _subtype="pdf")
        attach.add_header("Content-Disposition", "attachment",
                           filename=pdf_path.name)
        msg.attach(attach)
    print("PDFファイルをメールに添付しました。")

    try:
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as smtp:
            smtp.starttls()
            smtp.login(EMAIL_SENDER, EMAIL_PASSWORD)
            smtp.send_message(msg)
        print("PDFレポートを正常に送信しました。")
    except Exception as e:
        print(f"PDFレポートの送信中にエラーが発生しました: {e}")

def send_notification_email(subject: str, body: str):
    """処理完了通知メールを送信します。"""
    print(f"通知メールを送信しています: {subject}")
    msg = MIMEMultipart()
    msg["From"] = EMAIL_SENDER
    msg["To"] = NOTIFICATION_RECIPIENT
    msg["Subject"] = subject
    msg.attach(MIMEText(body, "plain"))

    try:
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as smtp:
            smtp.starttls()
            smtp.login(EMAIL_SENDER, EMAIL_PASSWORD)
            smtp.send_message(msg)
        print("通知メールを正常に送信しました。")
    except Exception as e:
        print(f"通知メールの送信中にエラーが発生しました: {e}")
