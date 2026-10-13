#!/usr/bin/env python3
"""探测 LaTeX 后端的运行环境：pandoc、xelatex、TeX 宏包、字体、PDF 读取工具。

用法：python scripts/probe_env.py          # 人类可读摘要
      python scripts/probe_env.py --json   # 另附 JSON，便于贴回给维护者
"""

import importlib.util
import json
import re
import shutil
import subprocess
import sys
from pathlib import Path

SKILL = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(SKILL / "scripts"))
from toolchain import find_pandoc, find_xelatex, installed_font_names  # noqa: E402
TOOLS = ["pandoc", "xelatex", "kpsewhich", "pdftotext", "pdftoppm"]
# LaTeX 后端模板 templates/latex/report.tmpl.tex 实际用到的宏包
TEX_FILES = ["ctexart.cls", "xeCJK.sty", "fontspec.sty", "geometry.sty", "fancyhdr.sty",
             "caption.sty", "booktabs.sty", "longtable.sty", "tabularx.sty", "array.sty",
             "multirow.sty", "makecell.sty", "calc.sty", "graphicx.sty", "float.sty",
             "listings.sty", "enumitem.sty", "xcolor.sty", "hyperref.sty"]
FONTS = {
    "宋体 SimSun": ["simsun", "宋体"], "新宋体 NSimSun": ["nsimsun", "新宋体"],
    "黑体 SimHei": ["simhei", "黑体"], "楷体 KaiTi": ["kaiti", "楷体"],
    "Times New Roman": ["times new roman"],
    "思源宋体 Noto Serif CJK SC": ["noto serif cjk sc"],
    "思源黑体 Noto Sans CJK SC": ["noto sans cjk sc"],
    "AR PL UMing CN": ["ar pl uming cn"], "TeX Gyre Termes": ["tex gyre termes"],
    "Liberation Serif": ["liberation serif"],
    "Songti SC (macOS)": ["songti sc"], "Heiti SC (macOS)": ["heiti sc"],
}


def sh(cmd, timeout=30):
    try:
        r = subprocess.run(cmd, capture_output=True, text=True, timeout=timeout)
        return r.stdout
    except Exception:
        return ""


def main():
    info = {"tools": {}, "tex_files": {}, "fonts": {}, "python": {}}
    for t in TOOLS:
        info["tools"][t] = shutil.which(t)
    info["tools"]["pandoc"] = find_pandoc()  # 含 PANDOC 环境变量、pypandoc_binary 自带的
    info["tools"]["xelatex"] = find_xelatex()  # 含 XELATEX 环境变量
    m = re.search(r"pandoc\s+(\d+)\.(\d+)", sh([info["tools"]["pandoc"], "--version"])) \
        if info["tools"]["pandoc"] else None
    info["pandoc_version"] = "%s.%s" % m.groups() if m else None
    info["python"]["version"] = sys.version.split()[0]
    info["python"]["pymupdf"] = importlib.util.find_spec("pymupdf") is not None
    if shutil.which("kpsewhich"):
        for f in TEX_FILES:
            info["tex_files"][f] = bool(sh(["kpsewhich", f]).strip())
    families = installed_font_names()
    for label, keys in FONTS.items():
        info["fonts"][label] = any(k in families for k in keys)

    pandoc_ok = bool(info["pandoc_version"]) and int(info["pandoc_version"].split(".")[0]) >= 3
    latex_ready = bool(info["tools"]["xelatex"] and pandoc_ok
                       and info["tex_files"] and all(info["tex_files"].values()))

    problems = []
    if not info["tools"]["pandoc"]:
        problems.append("缺 pandoc：正文转 LaTeX 片段依赖它（pip install pypandoc_binary，或装 pandoc）")
    elif not pandoc_ok:
        problems.append("pandoc 版本低于 3.0，图表编号过滤器需要 3.0+")
    if not info["tools"]["xelatex"]:
        problems.append("缺 xelatex：LaTeX 后端出 PDF 依赖它（装 TeX Live / MiKTeX / TinyTeX）")
    if info["tex_files"]:
        miss = [k for k, v in info["tex_files"].items() if not v]
        if miss:
            problems.append("缺 TeX 宏包：%s（TinyTeX 用 tlmgr install 补；多数随 TeX Live/MiKTeX 自带）"
                            % "、".join(miss))
    if not info["python"]["pymupdf"]:
        if not info["tools"]["pdftotext"]:
            problems.append("缺 pymupdf 和 pdftotext：无法读取 PDF 文字做验收（pip install pymupdf）")
        if not info["tools"]["pdftoppm"]:
            problems.append("缺 pymupdf 和 pdftoppm：无法把 PDF 逐页转图片验收（pip install pymupdf）")
    if not info["fonts"]["黑体 SimHei"]:
        problems.append("本环境没有黑体 SimHei：ctexart 标题会退回替代字体")
    if not info["fonts"]["宋体 SimSun"]:
        problems.append("本环境没有宋体 SimSun：ctexart 正文/封面会退回替代字体")

    info["latex_ready"] = latex_ready
    info["recommended_backend"] = "latex" if latex_ready else "none"
    info["problems"] = problems

    print("== 工具 ==")
    for t, path in info["tools"].items():
        print("  %-10s %s" % (t, path or "—"))
    print("  pandoc 版本：", info["pandoc_version"] or "—")
    print("  pymupdf：", "有" if info["python"]["pymupdf"] else "无")
    if info["tex_files"]:
        miss = [k for k, v in info["tex_files"].items() if not v]
        print("== TeX 宏包 ==\n  缺失：", miss or "无")
    print("== 字体 ==")
    for label, ok in info["fonts"].items():
        print("  %-28s %s" % (label, "有" if ok else "—"))
    print("== 结论 ==")
    print("  LaTeX 后端就绪（xelatex + pandoc≥3 + 宏包齐全）：", latex_ready)
    print("  推荐后端：", info["recommended_backend"])
    for p in problems:
        print("  注意：" + p)
    if "--json" in sys.argv:
        print(json.dumps(info, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
