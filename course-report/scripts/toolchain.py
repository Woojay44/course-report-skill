"""外部工具探测与调用：pandoc、PDF 渲染（LibreOffice / Windows 上的 Word）、PDF 文字与页面图像。

其他脚本统一从这里取，保证 probe_env 报告的和实际使用的是同一套工具。
"""

import os
import re
import shutil
import subprocess
import sys
from pathlib import Path

IS_WINDOWS = sys.platform == "win32"


# ---------------------------------------------------------------- pandoc
def find_pandoc():
    """环境变量 PANDOC → PATH → pypandoc_binary 自带的 pandoc。"""
    env = os.environ.get("PANDOC")
    if env and Path(env).exists():
        return env
    exe = shutil.which("pandoc")
    if exe:
        return exe
    try:
        import pypandoc
        path = pypandoc.get_pandoc_path()
    except Exception:
        return None
    # Windows 上返回的路径不带 .exe
    for cand in (path, path + ".exe") if path else ():
        if Path(cand).is_file():
            return cand
    return None


# ---------------------------------------------------------------- 渲染器
def find_xelatex():
    """环境变量 XELATEX → PATH。"""
    env = os.environ.get("XELATEX")
    if env and Path(env).exists():
        return env
    return shutil.which("xelatex")


def run_xelatex(tex_path, outdir, cwd, runs=2, timeout=300):
    """用 xelatex 把 tex 编成 PDF，跑 runs 遍（默认两遍，生成目录页码）。
    cwd 设为 report.md 所在目录，使正文里的 screenshots/... 相对路径能被找到
    （规避含中文的 \\graphicspath 在部分环境下的解析问题）。
    返回 (pdf 路径或 None, 最后一遍的日志)。"""
    xe = find_xelatex()
    if not xe:
        return None, "找不到 xelatex"
    tex_path, outdir, cwd = Path(tex_path).resolve(), Path(outdir).resolve(), Path(cwd).resolve()
    outdir.mkdir(parents=True, exist_ok=True)
    # xelatex 把参数里的反斜杠当控制序列，Windows 下必须转成正斜杠；
    # tex/输出与 cwd 跨盘时 relpath 会抛错，退回绝对路径（正斜杠）。
    try:
        tex_rel = os.path.relpath(str(tex_path), str(cwd)).replace(os.sep, "/")
        out_rel = os.path.relpath(str(outdir), str(cwd)).replace(os.sep, "/")
    except ValueError:
        tex_rel = str(tex_path).replace(os.sep, "/")
        out_rel = str(outdir).replace(os.sep, "/")
    log = ""
    for _ in range(runs):
        r = subprocess.run([xe, "-interaction=nonstopmode", "-halt-on-error",
                            "-output-directory", out_rel, tex_rel],
                           cwd=str(cwd), capture_output=True, text=True,
                           encoding="utf-8", errors="replace", timeout=timeout)
        log = r.stdout or ""
        if r.returncode != 0:
            return None, log + "\n" + (r.stderr or "")
    pdf = outdir / (tex_path.stem + ".pdf")
    return (pdf if pdf.exists() else None), log


def find_soffice():
    return shutil.which("soffice") or shutil.which("libreoffice")


def word_available():
    """Windows 上是否注册了 Word 的 COM 接口。"""
    if not IS_WINDOWS:
        return False
    try:
        import winreg
        winreg.CloseKey(winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r"Word.Application\CLSID"))
        return True
    except OSError:
        return False


def find_renderer():
    """返回 "soffice"、"word" 或 None。"""
    if find_soffice():
        return "soffice"
    if word_available():
        return "word"
    return None


# Word 已经在运行时，New-Object 可能接到同一个实例：只关自己打开的文档，
# 实例里还有用户的文档就不 Quit。
WORD_EXPORT_PS = r"""
$ErrorActionPreference = 'Stop'
$w = New-Object -ComObject Word.Application
$w.Visible = $false
$w.DisplayAlerts = 0
$d = $null
try {
  $d = $w.Documents.Open($env:CR_DOCX, $false, $true, $false)
  if ($env:CR_UPDATE_TOC -eq '1') { foreach ($t in $d.TablesOfContents) { $t.Update() } }
  $d.ExportAsFixedFormat($env:CR_PDF, 17)
} finally {
  if ($d -ne $null) { $d.Close(0) }
  if ($w.Documents.Count -eq 0) { $w.Quit() }
}
"""


def render_pdf(docx_path, outdir, update_toc=False, timeout=300):
    """把 docx 渲染成 PDF，返回 (pdf 路径, 渲染器名)；失败返回 (None, 渲染器名或 None)。
    update_toc 只对 Word 有效：导出前先更新目录域（预览用；回填页码时不要开）。"""
    docx_path, outdir = Path(docx_path).resolve(), Path(outdir).resolve()
    outdir.mkdir(parents=True, exist_ok=True)
    pdf = outdir / (docx_path.stem + ".pdf")
    if pdf.exists():
        pdf.unlink()
    renderer = find_renderer()
    try:
        if renderer == "soffice":
            subprocess.run([find_soffice(), "--headless", "--convert-to", "pdf", "--outdir",
                            str(outdir), str(docx_path)], capture_output=True, timeout=timeout,
                           check=True)
        elif renderer == "word":
            env = dict(os.environ, CR_DOCX=str(docx_path), CR_PDF=str(pdf),
                       CR_UPDATE_TOC="1" if update_toc else "0")
            subprocess.run(["powershell", "-NoProfile", "-NonInteractive", "-Command",
                            WORD_EXPORT_PS], capture_output=True, timeout=timeout, check=True,
                           env=env)
        else:
            return None, None
    except Exception:
        return None, renderer
    return (pdf if pdf.exists() else None), renderer


# ---------------------------------------------------------------- PDF 读取
def pdf_text_pages(pdf):
    """逐页文字（保留行结构）：pdftotext → pymupdf → pypdf。失败返回 None。"""
    if shutil.which("pdftotext"):
        try:
            out = subprocess.run(["pdftotext", "-enc", "UTF-8", "-layout", str(pdf), "-"],
                                 capture_output=True, text=True, encoding="utf-8",
                                 timeout=120).stdout
            pages = out.split("\f")
            if pages and not pages[-1].strip():
                pages = pages[:-1]
            if pages:
                return pages
        except Exception:
            pass
    try:
        import pymupdf
        with pymupdf.open(str(pdf)) as d:
            return [pg.get_text(sort=True) for pg in d]
    except Exception:
        pass
    try:
        from pypdf import PdfReader
        return [pg.extract_text() or "" for pg in PdfReader(str(pdf)).pages]
    except Exception:
        return None


def pdf_text_bottoms(pdf, first, last):
    """第 first..last 页（从 0 计）正文文字的最低位置占页高的比例（排除页脚 10%）。"""
    try:
        import pymupdf
    except ImportError:
        return None
    ratios = []
    with pymupdf.open(str(pdf)) as d:
        for i in range(first, last + 1):
            pg = d[i]
            h = pg.rect.height
            ys = [b[3] for b in pg.get_text("blocks") if b[3] < 0.9 * h]
            ratios.append(max(ys) / h if ys else 0.0)
    return ratios


def render_page_images(pdf, outdir, dpi=60, prefix="page"):
    """把每页渲染成 PNG：pymupdf → pdftoppm。返回图片路径列表（失败为空）。"""
    outdir = Path(outdir)
    outdir.mkdir(parents=True, exist_ok=True)
    try:
        import pymupdf
        paths = []
        with pymupdf.open(str(pdf)) as d:
            for i, pg in enumerate(d):
                path = outdir / ("%s-%02d.png" % (prefix, i + 1))
                pg.get_pixmap(dpi=dpi).save(str(path))
                paths.append(path)
        return paths
    except ImportError:
        pass
    if shutil.which("pdftoppm"):
        subprocess.run(["pdftoppm", "-r", str(dpi), "-png", str(pdf), str(outdir / prefix)],
                       capture_output=True, timeout=300)
        return sorted(outdir.glob(prefix + "-*.png"))
    return []


# ---------------------------------------------------------------- 字体
def installed_font_names():
    """已安装字体名（小写拼成一个字符串）：Windows 读注册表，其他系统用 fc-list。"""
    if IS_WINDOWS:
        import winreg
        names = []
        key_path = r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts"
        for hive in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
            try:
                with winreg.OpenKey(hive, key_path) as key:
                    i = 0
                    while True:
                        try:
                            names.append(winreg.EnumValue(key, i)[0])
                        except OSError:
                            break
                        i += 1
            except OSError:
                continue
        return re.sub(r"\s*\((truetype|opentype)\)", "", "\n".join(names), flags=re.I).lower()
    if shutil.which("fc-list"):
        try:
            return subprocess.run(["fc-list", ":", "family"], capture_output=True, text=True,
                                  timeout=30).stdout.lower()
        except Exception:
            return ""
    return ""
