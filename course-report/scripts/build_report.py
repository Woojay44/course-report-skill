#!/usr/bin/env python3
"""结构化 Markdown → 课程报告 PDF（LaTeX 后端：ctexart + xelatex）。
用法：
python scripts/build_report.py report.md --out out/
可选：
--defaults 文件                      个人默认值（姓名、学号、时区…）；缺省用 templates/defaults.local.yaml，
                                     没有时用 templates/defaults.yaml（示例值）
--numbering cn|decimal|manual        标题编号（默认取 defaults / front matter）
--class-opts "11pt,a4paper"          文档类选项
--no-cover / --no-toc
--keep-tex                           保留中间 .tex（默认也会留在 out/work 下）
流程：读取 front matter → 生成封面/摘要信息 → pandoc 把正文转成 LaTeX 片段（Lua 过滤器加中文编号/图表号）
      → 套 week07 风格模板（titlepage 封面 + \\pagenumbering Roman 目录/摘要 + arabic 正文）
      → xelatex 跑两遍出 PDF。
"""

import argparse
import json
import re
import subprocess
import sys
from datetime import datetime
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
from toolchain import find_pandoc, find_xelatex, run_xelatex  # noqa: E402
SKILL = HERE.parent
FILTER = HERE / "filters" / "report.lua"
TEMPLATE = SKILL / "templates" / "latex" / "report.tmpl.tex"
DEFAULTS_EXAMPLE = SKILL / "templates" / "defaults.yaml"
DEFAULTS_LOCAL = SKILL / "templates" / "defaults.local.yaml"  # 本人默认值，不进 git


def default_defaults():
    return DEFAULTS_LOCAL if DEFAULTS_LOCAL.exists() else DEFAULTS_EXAMPLE


def log(msg):
    print("[build_report] " + msg, flush=True)


def die(msg, code=2):
    print("[build_report] 错误：" + msg, file=sys.stderr, flush=True)
    sys.exit(code)


# ---------------------------------------------------------------- 元数据解析（借 pandoc，不依赖 PyYAML）


def inlines_text(items):
    out = []
    for it in items:
        t, c = it.get("t"), it.get("c")
        if t == "Str":
            out.append(c)
        elif t in ("Space", "SoftBreak"):
            out.append(" ")
        elif t == "LineBreak":
            out.append(" \n ")
        elif t == "Quoted":
            q = ("“", "”") if c[0].get("t") == "DoubleQuote" else ("‘", "’")
            out.append(q[0] + inlines_text(c[1]) + q[1])
        elif t == "Code":
            out.append(c[1])
        elif t in ("Link", "Image", "Span"):
            out.append(inlines_text(c[1]))
        elif t == "Cite":
            out.append(inlines_text(c[1]))
        elif t in ("Emph", "Strong", "Underline", "Strikeout", "Superscript", "Subscript", "SmallCaps"):
            out.append(inlines_text(c))
    return "".join(out)


def meta_value(v):
    t, c = v.get("t"), v.get("c")
    if t in ("MetaString", "MetaBool"):
        return c
    if t == "MetaInlines":
        return inlines_text(c)
    if t == "MetaBlocks":
        return [inlines_text(b["c"]) for b in c if b.get("t") in ("Para", "Plain")]
    if t == "MetaList":
        return [meta_value(x) for x in c]
    if t == "MetaMap":
        return {k: meta_value(x) for k, x in c.items()}
    return None


def pandoc_meta(pandoc, md_path=None, metadata_file=None):
    cmd = [pandoc, "-f", "markdown", "-t", "json"]
    stdin = ""
    if metadata_file:
        cmd += ["--metadata-file", str(metadata_file)]
    if md_path:
        cmd.append(str(md_path))
    r = subprocess.run(cmd, input=stdin if not md_path else None, capture_output=True, text=True, encoding="utf-8")
    if r.returncode != 0:
        die("pandoc 解析元数据失败：" + r.stderr.strip()[:300])
    meta = json.loads(r.stdout).get("meta", {})
    return {k: meta_value(v) for k, v in meta.items()}


def as_text(v):
    if isinstance(v, list):
        return "、".join(str(x) for x in v)
    return "" if v is None else str(v).strip()


def listify(v):
    if v is None or v == "":
        return []
    return v if isinstance(v, list) else [v]


# ---------------------------------------------------------------- LaTeX 转义与片段拼装
FRONT = re.compile(r"\A---[ \t]*\n.*?\n(?:---|\.\.\.)[ \t]*\n", re.S)

_LATEX_ESC = {
    "\\": r"\textbackslash{}", "&": r"\&", "%": r"\%", "$": r"\$", "#": r"\#",
    "_": r"\_", "{": r"\{", "}": r"\}", "~": r"\textasciitilde{}",
    "^": r"\textasciicircum{}",
}


def latex_escape(s):
    """转义注入模板的元数据文本（封面字段、标题、摘要、关键词）。
    正文由 pandoc 自行转义，不走这里。"""
    return "".join(_LATEX_ESC.get(ch, ch) for ch in str(s))


def resolve_date(raw, tz):
    """date 为空或 'auto' 时按时区取当天，格式“2026年10月12日”；否则原样用。"""
    raw = as_text(raw)
    if raw and raw.lower() != "auto":
        return raw
    try:
        from zoneinfo import ZoneInfo
        now = datetime.now(ZoneInfo(tz))
    except Exception:
        now = datetime.now()
    return "%d年%d月%d日" % (now.year, now.month, now.day)


def build_cover(org, course, title, subtitle, fields):
    """titlepage 封面（week07 风格）。fields：[[标签, 值], ...]。"""
    L = [r"\begin{titlepage}", r"  \coversong", r"  \centering", r"  \vspace*{2.5cm}"]
    if org:
        L += [r"  {\zihao{-2} %s}\par" % latex_escape(org), r"  \vspace{1.1cm}"]
    if course:
        L += [r"  {\zihao{3} %s}\par" % latex_escape(course),
              r"  \nointerlineskip\vspace{0.35cm}"]
    L += [r"  \rule{0.5\textwidth}{0.6pt}\par", r"  \vspace{1.0cm}"]
    L.append(r"  {\zihao{1}\bfseries %s}\par" % latex_escape(title))
    if subtitle:
        L += [r"  \vspace{0.6cm}", r"  {\zihao{3} %s}\par" % latex_escape(subtitle)]
    L.append(r"  \vfill")
    L += [r"  {\zihao{3}", r"    \renewcommand{\arraystretch}{2.1}", r"    \begin{tabular}{l}"]
    for lab, val in fields:
        if as_text(val):
            L.append(r"      \coverfield{%s}{%s} \\" % (latex_escape(lab), latex_escape(val)))
    L += [r"    \end{tabular}\par}", r"  \vspace{2.8cm}", r"\end{titlepage}"]
    return "\n".join(L)


def abstract_tex(heading, paras, kw_label, keywords):
    L = [r"\begin{center}{\bfseries\zihao{-3} %s}\end{center}" % latex_escape(heading)]
    for p in paras:
        L += [r"\par", latex_escape(p)]
    if keywords:
        L += [r"\par\vspace{0.6em}",
              r"\noindent\textbf{%s}%s" % (latex_escape(kw_label), latex_escape("、".join(keywords)))]
    return "\n".join(L)


def build_front(want_toc, zh_paras, zh_kw, en_paras, en_kw):
    """前置部分：Roman 页码下的目录与摘要。"""
    L = [r"\pagenumbering{Roman}"]
    if want_toc:
        L += [r"\tableofcontents", r"\clearpage"]
    if zh_paras:
        L.append(abstract_tex("摘要", zh_paras, "关键词：", zh_kw))
    if en_paras:
        if zh_paras:
            L.append(r"\vspace{1.5em}")
        L.append(abstract_tex("Abstract", en_paras, "Keywords: ", en_kw))
    if zh_paras or en_paras:
        L.append(r"\clearpage")
    return "\n".join(L)


def strip_listings_language(tex):
    """listings 不认识 javascript/apache/html 等语言标签，去掉 language=... 选项。"""
    def repl(m):
        opts = m.group(1)
        opts = re.sub(r"language=[^,\]]*", "", opts, flags=re.I)
        opts = re.sub(r",\s*,", ",", opts)
        opts = opts.strip().strip(",").strip()
        return r"\begin{lstlisting}[" + opts + "]"
    return re.sub(r"\\begin\{lstlisting\}\[([^\]]*)\]", repl, tex)


# ---------------------------------------------------------------- 校验


def lint(md_text, md_dir, meta):
    problems = []
    n = md_text.count("【待补充")
    if n:
        problems.append("正文里还有 %d 处占位符，交付前需补全或向用户说明" % n)
    for _alt, src in re.findall(r"!\[([^\]]*)\]\(\s*<?([^)\s>]+)>?(?:\s+\"[^\"]*\")?\s*\)", md_text):
        if not re.match(r"^[a-z]+://", src) and not (md_dir / src).exists():
            problems.append("图片文件不存在：" + src)
    if not re.search(r"^#{1,3}\s+\S", FRONT.sub("", md_text), re.M):
        problems.append("正文里没有任何 # 标题，目录会是空的")
    if not meta.get("keywords") and meta.get("abstract"):
        problems.append("有摘要但没有 keywords")
    return problems


def run_logged(cmd, tag):
    r = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8")
    for line in (r.stdout or "").splitlines():
        print(line, flush=True)
    if r.returncode != 0:
        print((r.stderr or "")[-1500:], file=sys.stderr)
        die("%s 失败（退出码 %d）" % (tag, r.returncode), 4)


# ---------------------------------------------------------------- 主流程


def main():
    ap = argparse.ArgumentParser(description="结构化 Markdown → 课程报告 PDF（LaTeX 后端）")
    ap.add_argument("markdown", type=Path)
    ap.add_argument("--out", type=Path, default=Path("out"))
    ap.add_argument("--defaults", type=Path, default=None)
    ap.add_argument("--backend", choices=["latex"], default="latex",
                    help="仅保留 LaTeX 后端（docx 后端已移除）")
    ap.add_argument("--numbering", choices=["cn", "decimal", "manual"])
    ap.add_argument("--class-opts", default=None)
    ap.add_argument("--no-cover", action="store_true")
    ap.add_argument("--no-toc", action="store_true")
    ap.add_argument("--keep-tex", action="store_true")
    opts = ap.parse_args()

    pandoc = find_pandoc()
    if not pandoc:
        die("找不到 pandoc（可安装 pandoc，或 pip install pypandoc_binary，或用环境变量 PANDOC 指定）", 3)
    if not find_xelatex():
        die("找不到 xelatex（请安装 TeX Live / MiKTeX，或用环境变量 XELATEX 指定）", 3)
    if not TEMPLATE.exists():
        die("找不到模板：" + str(TEMPLATE))

    md = opts.markdown.resolve()
    if not md.exists():
        die("找不到 " + str(md))
    md_text = md.read_text(encoding="utf-8")

    opts.defaults = opts.defaults or default_defaults()
    if opts.defaults == DEFAULTS_EXAMPLE:
        log("提示：没有 templates/defaults.local.yaml，封面署名用的是 defaults.yaml 里的示例值")
    defaults = pandoc_meta(pandoc, metadata_file=opts.defaults) if opts.defaults.exists() else {}
    meta = pandoc_meta(pandoc, md_path=md)

    title = as_text(meta.get("title"))
    if not title:
        die("front matter 必须有 title")
    numbering = opts.numbering or as_text(meta.get("numbering")) or as_text(defaults.get("numbering")) or "cn"
    tz = as_text(meta.get("timezone")) or as_text(defaults.get("timezone")) or "Asia/Taipei"
    class_opts = opts.class_opts or as_text(meta.get("class_opts")) or "11pt,a4paper"
    want_cover = not opts.no_cover and meta.get("cover") is not False
    want_toc = not opts.no_toc and meta.get("toc") is not False

    for problem in lint(md_text, md.parent, meta):
        log("提示：" + problem)

    out = opts.out.resolve()
    work = out / "work"
    work.mkdir(parents=True, exist_ok=True)

    # 1) pandoc：正文（去 front matter）→ LaTeX 片段，Lua 过滤器加中文编号/图表号，listings 排代码
    body_src = work / "body_src.md"
    body_src.write_text(FRONT.sub("", md_text, count=1), encoding="utf-8")
    body_tex_raw = work / "body_raw.tex"
    run_logged([pandoc, str(body_src), "-f", "markdown", "-t", "latex", "--listings",
                "--wrap=preserve", "-o", str(body_tex_raw),
                "--lua-filter", str(FILTER), "-M", "numbering=" + numbering,
                "--resource-path", str(md.parent)], "pandoc")
    body_tex = strip_listings_language(body_tex_raw.read_text(encoding="utf-8"))
    log("pandoc 转换完成（标题编号方案：%s；已清理 listings 语言标签）" % numbering)

    # 2) 封面字段
    name = as_text(meta.get("author")) or as_text(defaults.get("name"))
    org = as_text(meta.get("org")) or as_text(defaults.get("org"))
    course = as_text(meta.get("course")) or as_text(defaults.get("course"))
    subtitle = as_text(meta.get("subtitle"))
    fields = []
    if name:
        fields.append(["姓名", name])
    sid = as_text(meta.get("student_id")) or as_text(defaults.get("student_id"))
    if sid:
        fields.append(["学号", sid])
    for item in listify(meta.get("cover_fields")):
        if isinstance(item, dict):
            fields += [[k, as_text(v)] for k, v in item.items()]
    fields.append(["写作日期", resolve_date(meta.get("date"), tz)])

    cover = build_cover(org, course, title, subtitle, fields) if want_cover else ""

    # 3) 前置（目录 + 摘要），正文切回 arabic
    front = build_front(want_toc,
                        [str(p) for p in listify(meta.get("abstract"))],
                        [str(k) for k in listify(meta.get("keywords"))],
                        [str(p) for p in listify(meta.get("abstract_en"))],
                        [str(k) for k in listify(meta.get("keywords_en"))])
    body = r"\pagenumbering{arabic}" + "\n" + body_tex

    # 4) 套模板
    head_l = course or org or title
    head_r = title if (course or org) else name
    gpath = str(md.parent).replace("\\", "/")
    graphicspath = r"\graphicspath{{%s/}}" % gpath
    tex = TEMPLATE.read_text(encoding="utf-8")
    for token, value in [
        ("@@CLASSOPTS@@", class_opts),
        ("@@GRAPHICSPATH@@", graphicspath),
        ("@@HEADL@@", latex_escape(head_l)),
        ("@@HEADR@@", latex_escape(head_r)),
        ("@@COVER@@", cover),
        ("@@FRONT@@", front),
        ("@@BODY@@", body),
    ]:
        tex = tex.replace(token, value)

    report_tex = work / (md.stem + ".tex")
    report_tex.write_text(tex, encoding="utf-8")
    log("已生成 LaTeX：" + str(report_tex))

    # 5) xelatex 跑两遍（cwd = report.md 所在目录，使 screenshots/... 相对路径可解析）
    pdf, xlog = run_xelatex(report_tex, work, cwd=md.parent, runs=2)
    if not pdf:
        tail = "\n".join([ln for ln in xlog.splitlines() if "!" in ln or "Error" in ln][-20:]) or xlog[-1500:]
        print(tail, file=sys.stderr)
        die("xelatex 编译失败，详见上方日志与 %s" % (work / (md.stem + ".log")), 4)

    final = out / (md.stem + ".pdf")
    if final.exists():
        final.unlink()
    pdf.replace(final)
    log("已生成 PDF：" + str(final))
    log("提示：如目录页码异常，确认 xelatex 已跑两遍（本脚本默认两遍）。")


if __name__ == "__main__":
    main()
