#!/usr/bin/env python3
"""结构化 Markdown → 课程报告（阶段 1：docx 后端）。
用法：
python scripts/build_report.py report.md --out out/
可选：
--defaults 文件                      个人默认值（姓名、学号、时区…）；缺省用 templates/defaults.local.yaml，
                                     没有时用 templates/defaults.yaml（示例值）
--backend docx                       目前只有 docx（LaTeX 后端阶段 2 提供）
--numbering cn|decimal|manual        标题编号（默认取 defaults / front matter）
--page-numbers roman|continuous
--no-cover / --no-toc / --no-preview
流程：读取 front matter → 生成封面信息 → 摘要拼进正文 → pandoc 转 docx（Lua 过滤器编号） → format_docx.py 排版（封面、目录、摘要、页码分节）→ 预览 PDF（LibreOffice，或 Windows 上的 Word）。
"""

import argparse
import json
import re
import subprocess
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
from toolchain import find_pandoc, render_pdf  # noqa: E402
SKILL = HERE.parent
FILTER = HERE / "filters" / "report.lua"
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


# ---------------------------------------------------------------- 正文准备
FRONT = re.compile(r"\A---[ \t]*\n.*?\n(?:---|\.\.\.)[ \t]*\n", re.S)


def md_escape(text):
    text = re.sub(r"([\\`*_{}\[\]<>#|~$^])", r"\\\1", text)
    return re.sub(r"^(\d+)\.", r"\1\.", text)


def abstract_block(title, paras, kw_label, keywords, sep):
    lines = [md_escape(title), ""]
    for p in paras:
        lines += [md_escape(p), ""]
    if keywords:
        lines += [md_escape(kw_label + sep.join(keywords)), ""]
    return "\n".join(lines) + "\n"


def listify(v):
    if v is None or v == "":
        return []
    return v if isinstance(v, list) else [v]


def lint(md_text, md_dir, meta):
    problems = []
    n = md_text.count("【待补充")
    if n:
        problems.append("正文里还有 %d 处占位符，交付前需补全或向用户说明" % n)
    # 兼容 ![图题](a.png)、![图题](<a b.png>)、![图题](a.png "标题")
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
    if r.returncode!= 0:
        print((r.stderr or "")[-1500:], file=sys.stderr)
        die("%s 失败（退出码 %d）" % (tag, r.returncode), 4)
def main():
    ap = argparse.ArgumentParser(description="结构化 Markdown → 课程报告")
    ap.add_argument("markdown", type=Path)
    ap.add_argument("--out", type=Path, default=Path("out"))
    ap.add_argument("--defaults", type=Path, default=None)
    ap.add_argument("--backend", choices=["auto", "docx", "latex"], default="auto")
    ap.add_argument("--numbering", choices=["cn", "decimal", "manual"])
    ap.add_argument("--page-numbers", choices=["roman", "continuous"])
    ap.add_argument("--no-cover", action="store_true")
    ap.add_argument("--no-toc", action="store_true")
    ap.add_argument("--no-preview", action="store_true")
    opts = ap.parse_args()

    pandoc = find_pandoc()
    if not pandoc:
        die("找不到 pandoc（可安装 pandoc，或 pip install pypandoc_binary，或用环境变量 PANDOC 指定）", 3)
    if opts.backend == "latex":
        die("LaTeX 后端将在阶段 2 提供，目前请用 --backend docx 或 auto", 3)
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
    scheme = opts.page_numbers or as_text(meta.get("page_numbers")) or as_text(defaults.get("page_numbers")) or "roman"
    tz = as_text(meta.get("timezone")) or as_text(defaults.get("timezone")) or "Asia/Taipei"
    want_cover = not opts.no_cover and meta.get("cover") is not False
    want_toc = not opts.no_toc and meta.get("toc") is not False

    for problem in lint(md_text, md.parent, meta):
        log("提示：" + problem)

    out = opts.out.resolve()
    work = out / "work"
    work.mkdir(parents=True, exist_ok=True)

    # 正文：去掉 front matter，把摘要拼在最前面（format_docx 按"摘要/Abstract"标题识别）
    body = FRONT.sub("", md_text, count=1)
    head = ""
    zh_paras = listify(meta.get("abstract"))
    if zh_paras:
        head += abstract_block("摘要", zh_paras, "关键词：", [str(k) for k in listify(meta.get("keywords"))], "；")
    en_paras = listify(meta.get("abstract_en"))
    if en_paras:
        head += abstract_block("Abstract", en_paras, "Keywords: ",
            [str(k) for k in listify(meta.get("keywords_en"))], "; ")
    body_md = work / "body.md"
    body_md.write_text(head + body, encoding="utf-8")

    base = work / "base.docx"
    run_logged([pandoc, str(body_md), "-f", "markdown", "-t", "docx", "-o", str(base),
        "--lua-filter", str(FILTER), "-M", "numbering=" + numbering,
        "--resource-path", str(md.parent)], "pandoc")
    log("pandoc 转换完成（标题编号方案：%s）" % numbering)

    cmd = [sys.executable, str(HERE / "format_docx.py"), str(base), str(out / (md.stem + ".docx")),
        "--defaults", "--fix-abstract", "--three-line", "all", "--timezone", tz,
        "--page-numbers", scheme]
    name = as_text(meta.get("author")) or as_text(defaults.get("name"))
    if name:
        cmd += ["--author", name]
    if want_cover:
        sid = as_text(meta.get("student_id")) or as_text(defaults.get("student_id"))
        fields = []
        if name:
            fields.append(["姓名", name])
        if sid:
            fields.append(["学号", sid])
        for item in listify(meta.get("cover_fields")):
            if isinstance(item, dict):
                fields += [[k, as_text(v)] for k, v in item.items()]
        fields.append(["写作日期", as_text(meta.get("date")) or "auto"])
        info = {"org": as_text(meta.get("org")) or as_text(defaults.get("org")),
            "course": as_text(meta.get("course")) or as_text(defaults.get("course")),
            "title": title, "subtitle": as_text(meta.get("subtitle")), "fields": fields}
        cover_json = work / "cover.json"
        cover_json.write_text(json.dumps(info, ensure_ascii=False, indent=2), encoding="utf-8")
        cmd += ["--cover", str(cover_json)]
    if want_toc:
        cmd.append("--toc")
    run_logged(cmd, "format_docx")

    docx_out = out / (md.stem + ".docx")
    log("已生成：" + str(docx_out))

    if not opts.no_preview:
        prev = out / "preview"
        pdf, renderer = render_pdf(docx_out, prev, update_toc=True)
        if pdf:
            final = prev / (md.stem + "_preview.pdf")
            pdf.replace(final)
            note = "Word 渲染，与本机 Word 打开效果一致" if renderer == "word" else \
                "字体为渲染环境的替代字体，仅供检查版面"
            log("预览 PDF：%s（%s；不是交付文件）" % (final, note))
        elif renderer:
            log("警告：用 %s 生成预览 PDF 失败" % renderer)
        else:
            log("提示：没有可用的渲染器（LibreOffice 或 Windows 上的 Word），未生成预览 PDF")
    log("提醒：在 Word 里打开后，右键目录 → 更新域 → 更新整个目录，再导出 PDF")


if __name__ == "__main__":
    main()
