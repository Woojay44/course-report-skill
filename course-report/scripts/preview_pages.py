#!/usr/bin/env python3
"""把预览 PDF 逐页渲染成 PNG，供验收时看图（替代 pdftoppm；优先用 pymupdf）。

用法：python scripts/preview_pages.py out/preview/report_preview.pdf [--dpi 60] [--out 目录]
默认输出到 PDF 同目录下的 pages/。
"""

import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from toolchain import render_page_images  # noqa: E402


def main():
    ap = argparse.ArgumentParser(description="预览 PDF → 逐页 PNG")
    ap.add_argument("pdf", type=Path)
    ap.add_argument("--dpi", type=int, default=60)
    ap.add_argument("--out", type=Path)
    opts = ap.parse_args()
    if not opts.pdf.exists():
        sys.exit("找不到 " + str(opts.pdf))
    out = opts.out or opts.pdf.parent / "pages"
    paths = render_page_images(opts.pdf, out, dpi=opts.dpi)
    if not paths:
        sys.exit("无法渲染页面：请 pip install pymupdf，或安装 poppler（pdftoppm）")
    for path in paths:
        print(path)


if __name__ == "__main__":
    main()
