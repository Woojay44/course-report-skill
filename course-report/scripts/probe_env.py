#!/usr/bin/env python3
"""探测运行环境并推荐后端。

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
from toolchain import find_pandoc, find_renderer, installed_font_names, word_available  # noqa: E402
TOOLS = ["pandoc", "soffice", "libreoffice", "xelatex", "latexmk", "tectonic",
         "pdftotext", "pdffonts", "pdftoppm"]
TEX_FILES = ["ctexart.cls", "xeCJK.sty", "fontspec.sty", "titlesec.sty", "tocloft.sty",
             "fancyhdr.sty", "caption.sty", "booktabs.sty", "longtable.sty", "graphicx.sty",
             "float.sty", "hyperref.sty", "amsmath.sty", "enumitem.sty", "array.sty",
             "fvextra.sty", "xcolor.sty", "setspace.sty", "geometry.sty"]
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
    info["tools"]["word"] = "Word（COM）" if word_available() else None
    m = re.search(r"pandoc\s+(\d+)\.(\d+)", sh([info["tools"]["pandoc"], "--version"])) \
        if info["tools"]["pandoc"] else None
    info["pandoc_version"] = "%s.%s" % m.groups() if m else None
    info["python"]["version"] = sys.version.split()[0]
    info["python"]["python-docx"] = importlib.util.find_spec("docx") is not None
    info["python"]["pymupdf"] = importlib.util.find_spec("pymupdf") is not None
    if shutil.which("kpsewhich"):
        for f in TEX_FILES:
            info["tex_files"][f] = bool(sh(["kpsewhich", f]).strip())
    families = installed_font_names()
    for label, keys in FONTS.items():
        info["fonts"][label] = any(k in families for k in keys)

    defaults = (SKILL / "templates" / "defaults.yaml")
    strict = True
    if defaults.exists():
        mm = re.search(r"^font_strict:\s*(true|false)", defaults.read_text(encoding="utf-8"), re.M | re.I)
        strict = not mm or mm.group(1).lower() == "true"

    pandoc_ok = bool(info["pandoc_version"]) and int(info["pandoc_version"].split(".")[0]) >= 3
    docx_ready = pandoc_ok and info["python"]["python-docx"]
    latex_ready = bool(info["tools"]["xelatex"] and info["tools"]["latexmk"] and pandoc_ok
                       and info["tex_files"] and all(info["tex_files"].values()))
    renderer = find_renderer()
    preview = renderer is not None

    problems = []
    if not info["tools"]["pandoc"]:
        problems.append("缺 pandoc：docx 后端和 Markdown 转换都依赖它")
    elif not pandoc_ok:
        problems.append("pandoc 版本低于 3.0，图表编号过滤器需要 3.0+")
    if not info["python"]["python-docx"]:
        problems.append("缺 python-docx：pip install python-docx")
    if not preview:
        problems.append("没有渲染器（LibreOffice，或 Windows 上的 Word）：无法生成预览，"
                        "目录页码不会回填，需在 Word 里更新目录")
    if not info["python"]["pymupdf"]:
        if not info["tools"]["pdftotext"]:
            problems.append("缺 pymupdf 和 pdftotext：无法读取预览 PDF 做页面检查（pip install pymupdf）")
        if not info["tools"]["pdftoppm"]:
            problems.append("缺 pymupdf 和 pdftoppm：无法把预览页转成图片（pip install pymupdf）")
    if not info["fonts"]["黑体 SimHei"]:
        problems.append("本环境没有黑体 SimHei：预览里标题用替代字体，目录页码可能与 Word 相差一两页，"
                        "交付前请在 Word 里更新目录（docx 文件里写的仍是宋体/黑体）")
    if not info["fonts"]["宋体 SimSun"]:
        problems.append("本环境没有宋体 SimSun：预览用替代字体，同上")

    info["docx_ready"] = docx_ready
    info["latex_ready"] = latex_ready
    info["latex_backend_available_in_this_version"] = False
    info["font_strict"] = strict
    info["recommended_backend"] = "docx" if docx_ready else "none"
    info["renderer"] = renderer
    info["problems"] = problems

    print("== 工具 ==")
    for t, path in info["tools"].items():
        print("  %-10s %s" % (t, path or "—"))
    print("  pandoc 版本：", info["pandoc_version"] or "—")
    print("  python-docx：", "有" if info["python"]["python-docx"] else "无")
    print("  pymupdf：", "有" if info["python"]["pymupdf"] else "无")
    if info["tex_files"]:
        miss = [k for k, v in info["tex_files"].items() if not v]
        print("== TeX 宏包 ==\n  缺失：", miss or "无")
    print("== 字体 ==")
    for label, ok in info["fonts"].items():
        print("  %-28s %s" % (label, "有" if ok else "—"))
    print("== 结论 ==")
    print("  docx 后端可用：", docx_ready, "｜ LaTeX 环境就绪：", latex_ready,
          "（本版本尚未提供 LaTeX 后端）")
    print("  课程字体硬性要求（font_strict）：", strict)
    print("  推荐后端：", info["recommended_backend"])
    print("  预览/目录页码渲染器：", {"soffice": "LibreOffice", "word": "Word"}.get(renderer, "无"))
    for p in problems:
        print("  注意：" + p)
    if "--json" in sys.argv:
        print(json.dumps(info, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
