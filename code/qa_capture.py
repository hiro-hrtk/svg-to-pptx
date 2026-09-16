"""QA層。PowerPoint COM経由でスライド1枚目をPNGにエクスポートする(視覚QA用)。

Windows専用(COM依存)。
使い方:
    uv run python code/qa_capture.py output/<session>/<file>.pptx
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

import win32com.client


def export_slide_png(pptx_path: str, out_path: str | None = None, dpi: int = 150) -> Path:
    pptx_path = Path(pptx_path).resolve()
    if out_path is None:
        out_path = pptx_path.parent / f"{pptx_path.stem}_qa_slide1.png"
    else:
        out_path = Path(out_path).resolve()

    # 既にPowerPointが起動済みならそのインスタンスを使い、新規に起動した場合のみ
    # 後でQuit()する。既存インスタンスをQuit()すると、ユーザーが開いている他の
    # プレゼンテーションまで巻き添えで閉じてしまうため。
    try:
        app = win32com.client.GetActiveObject("PowerPoint.Application")
        created_new = False
    except Exception:
        app = win32com.client.Dispatch("PowerPoint.Application")
        created_new = True

    pres = app.Presentations.Open(str(pptx_path), WithWindow=False)
    try:
        slide = pres.Slides(1)
        width_px = int(pres.PageSetup.SlideWidth / 72 * dpi)
        height_px = int(pres.PageSetup.SlideHeight / 72 * dpi)
        slide.Export(str(out_path), "PNG", width_px, height_px)
    finally:
        pres.Close()
        if created_new:
            app.Quit()

    return out_path


def main():
    parser = argparse.ArgumentParser(description="PPTX 1枚目をPNGにエクスポート(QA用)")
    parser.add_argument("pptx_path")
    parser.add_argument("--out", default=None)
    parser.add_argument("--dpi", type=int, default=150)
    args = parser.parse_args()

    out_path = export_slide_png(args.pptx_path, args.out, args.dpi)
    print(f"QA PNG: {out_path}")


if __name__ == "__main__":
    sys.exit(main())
