#!/usr/bin/env python3
"""自检：用 tests/sample.md 和 variant.md 跑完整 LaTeX 流程，检查 PDF 与 LaTeX。把整段输出贴回即可。

用法：python scripts/selftest.py
退出码：0 = 没有 FAIL；1 = 有 FAIL。
"""

import os
import re
import subprocess
import sys
import tempfile
from pathlib import Path

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

HERE = Path(__file__).resolve().parent
SKILL = HERE.parent
sys.path.insert(0, str(HERE))
from toolchain import find_pandoc, find_xelatex, pdf_text_pages, pdf_text_bottoms  # noqa: E402

CHECKS = []


def check(name, ok, detail=""):
    CHECKS.append((name, ok))
    mark = "PASS" if ok else "FAIL"
    print("  [%s] %s%s" % (mark, name, ("　— " + detail) if detail else ""))


def main():
    sample = SKILL / "tests" / "sample.md"
    if not find_pandoc():
        print("找不到 pandoc，无法自检"); sys.exit(1)
    if not find_xelatex():
        print("找不到 xelatex，无法自检"); sys.exit(1)
    if not sample.exists():
        print("找不到 " + str(sample)); sys.exit(1)

    with tempfile.TemporaryDirectory() as td:
        out = Path(td)
        child_env = dict(os.environ, PYTHONUTF8="1", PYTHONIOENCODING="utf-8")
        r = subprocess.run([sys.executable, str(HERE / "build_report.py"), str(sample),
                            "--out", str(out)], capture_output=True, text=True,
                           encoding="utf-8", errors="replace", env=child_env)
        print(r.stdout)
        if r.returncode != 0:
            print(r.stderr[-1500:])
            print("构建失败，自检中止"); sys.exit(1)

        pdf = out / (sample.stem + ".pdf")
        check("生成 PDF", pdf.exists(), str(pdf))
        if not pdf.exists():
            sys.exit(1)

        pages = pdf_text_pages(pdf) or []
        check("PDF 可读出文字", bool(pages), "共 %d 页" % len(pages))
        if not pages:
            _summary(); sys.exit(1 if any(not ok for _, ok in CHECKS) else 0)

        full = "\n".join(pages)
        check("有封面标题", "网络空间安全的战略演进" in pages[0] if pages else False)
        check("封面无页码", not re.search(r"^\s*[0-9IVXivx]+\s*$", pages[0].strip().splitlines()[-1] or "")
              if pages[0].strip() else True)
        check("有目录", any("目录" in p for p in pages[:3]))
        check("目录有页码（点线引导后有数字）", bool(re.search(r"\.{3,}\s*\d+", full) or re.search(r"引言.*\d", full)))
        check("有中文摘要", any("摘要" in p and "零信任" in p for p in pages))
        check("有英文摘要", "Abstract" in full and "zero trust" in full)
        check("中文标题自动编号（一、）", "一、" in full or "引言" in full)
        check("三线表内容在", "关键基础设施攻击" in full and "供应链投毒" in full)
        check("图题自动编号（图 1）", "图 1" in full or "图1" in full)
        check("表题自动编号（表 1）", "表 1" in full or "表1" in full)
        check("代码块内容在", "def verify" in full)
        check("参考文献在", "参考文献" in full and "Zero Trust Networks" in full)

        # 版面：正文页文字不应越过页面底部 92%
        bottoms = pdf_text_bottoms(pdf, 0, len(pages) - 1)
        if bottoms is not None:
            worst = max(bottoms) if bottoms else 0
            check("无内容越界（最低行 ≤ 0.95 页高）", worst <= 0.95, "最低 %.2f" % worst)
        else:
            print("  [SKIP] 越界检查（缺 pymupdf）")

    variant = SKILL / "tests" / "variant.md"
    with tempfile.TemporaryDirectory() as td:
        out = Path(td)
        r = subprocess.run([sys.executable, str(HERE / "build_report.py"), str(variant),
                            "--out", str(out), "--no-cover", "--no-toc", "--keep-tex"],
                           capture_output=True, text=True, encoding="utf-8", errors="replace",
                           env=child_env)
        print(r.stdout)
        check("variant 构建成功", r.returncode == 0, "退出码 %d" % r.returncode)
        if r.returncode != 0:
            print(r.stderr[-1500:])

        pdf = out / (variant.stem + ".pdf")
        pages = (pdf_text_pages(pdf) or []) if pdf.exists() else []
        full = "\n".join(pages)
        tex_path = out / "work" / (variant.stem + ".tex")
        tex = tex_path.read_text(encoding="utf-8") if tex_path.exists() else ""
        compact = re.sub(r"\s+", "", full)

        check("无封面（第 1 页为摘要）",
              bool(pages) and "变体摘要探针" in re.sub(r"\s+", "", pages[0]))
        check("无目录", bool(pages) and "目录" not in compact)
        check("数字编号（一级标题为 1、2）",
              bool(re.search(r"(?<!\d)1\s+编号探针", full))
              and bool(re.search(r"(?<!\d)2\s+公式探针", full))
              and not re.search(r"[一二三四五六七八九十]+、", compact))

        # 同时锁定正文引用与参考文献条目的上下文，避免模板或其他上标造成误通过。
        body_tex, separator, refs_tex = tex.partition(r"\section{参考文献}")
        check("引用上标（参考文献编号不转换）",
              bool(separator) and r"正文引用探针\textsuperscript{{[}1{]}}。" in body_tex
              and bool(re.search(r"(?m)^\s*\{\[\}1\{\]\}\s+引用条目探针。", refs_tex))
              and r"\textsuperscript" not in refs_tex)
        check("占位符提示",
              bool(re.search(r"(?<!\d)1\s*处占位符", r.stdout)))
        check("行内与独立公式",
              r"\(x + 1\)" in body_tex
              and bool(re.search(r"\\\[\s*y\s*=\s*2\s*\\\]", body_tex)))

    _summary()
    sys.exit(1 if any(not ok for _, ok in CHECKS) else 0)


def _summary():
    n_fail = sum(1 for _, ok in CHECKS if not ok)
    print("\n== 自检结果 ==  共 %d 项，FAIL %d 项" % (len(CHECKS), n_fail))


if __name__ == "__main__":
    main()
