import sys
from pathlib import Path
from typing import Optional

from config import PPT_SAVE_AS_PDF

_win32com_available = False
if sys.platform == "win32":
    try:
        import win32com.client
        _win32com_available = True
    except ImportError:
        print("警告: pywin32 ライブラリがインストールされていません。"
              "PDF変換機能は利用できません。")
        print("      PDF変換が必要な場合は、'pip install pywin32' または "
              "'uv pip install pywin32' を実行してください。")

def convert_ppt_to_pdf(pptx_path: Path) -> Optional[Path]:
    """PowerPointファイルをPDFに変換します。"""
    if not _win32com_available:
        print("PDF変換はWindows環境でのみサポートされており、pywin32ライブラリが必要です。")
        return None

    powerpoint = None
    presentation = None
    try:
        print(f"'{pptx_path.name}' をPDFに変換しています...")
        powerpoint = win32com.client.Dispatch("PowerPoint.Application")
        # powerpoint.Visible = 1  # デバッグ時以外は非表示が望ましい

        pdf_path = pptx_path.with_suffix(".pdf")

        presentation = powerpoint.Presentations.Open(str(pptx_path.resolve()))
        presentation.SaveAs(str(pdf_path.resolve()), PPT_SAVE_AS_PDF)  # PDF形式
        print(f"PDFに変換完了: {pdf_path.name}")
        return pdf_path

    except Exception as e:
        print(f"PDF変換中にエラーが発生しました: {e}")
        return None
    finally:
        if presentation:
            presentation.Close()
        if powerpoint:
            powerpoint.Quit()
