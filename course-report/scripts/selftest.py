#!/usr/bin/env python3
"""自检：用 tests/sample.md 跑一遍完整流程，逐项检查结构。把整段输出贴回即可。

用法：python scripts/selftest.py
"""

import re
import subprocess
import sys
import tempfile
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn

SKILL = Path(__file__).resolve().parent.parent
results = []


def check(name, ok, detail=""):
    results.append(bool(ok))
    print("%s  %s%s" % ("PASS" if ok else "FAIL", name, ("  —  " + str(detail)) if detail else ""))


def text_of(el):
    return "".join(t.text or "" for t in el.iter(qn("w:t")))


def last_line(page_text):
    lines = [l.strip() for l in page_text.splitlines() if l.strip()]
    return lines[-1] if lines else ""


def main():
    out = Path(tempfile.mkdtemp(prefix="cr_selftest_"))
    r = subprocess.run([sys.executable, str(SKILL / "scripts" / "build_report.py"),
                        str(SKILL / "tests" / "sample.md"), "--out", str(out)],
                       capture_output=True, text=True, encoding="utf-8")
    print(r.stdout)
    if r.returncode != 0:
        print(r.stderr[-1500:])
    check("构建成功", r.returncode == 0, "退出码 %s" % r.returncode)
    docx_path = out / "sample.docx"
    if not docx_path.exists():
        print("SELFTEST RESULT: 0/%d（没有生成 docx）" % len(results))
        sys.exit(1)

    doc = Document(docx_path)
    secs = doc.sections
    check("分成 3 节（封面/目录摘要/正文）", len(secs) == 3, len(secs))

    body = list(doc.element.body.iterchildren())
    first_sect_idx = next((i for i, e in enumerate(body)
                           if e.tag == qn("w:p") and e.find(qn("w:pPr")) is not None
                           and e.find(qn("w:pPr")).find(qn("w:sectPr")) is not None), 0)
    cover_text = "".join(text_of(e) for e in body[:first_sect_idx + 1])
    check("封面含姓名", "张三" in cover_text)
    check("封面含学号", "2024000000000" in cover_text)
    check("封面含写作日期（自动）", re.search(r"\d{4}年\d{1,2}月\d{1,2}日", cover_text), re.findall(r"\d{4}年\d{1,2}月\d{1,2}日", cover_text))

    def has_page_field(section):
        return any("PAGE" in (t.text or "") for t in section.footer._element.iter(qn("w:instrText")))

    check("封面无页码、其余两节有页码", not has_page_field(secs[0]) and has_page_field(secs[1]) and has_page_field(secs[2]))
    fmt = []
    for s in secs:
        pg = s._sectPr.find(qn("w:pgNumType"))
        fmt.append((pg.get(qn("w:fmt")), pg.get(qn("w:start"))) if pg is not None else None)
    check("目录摘要节为罗马数字、正文节从 1 开始", fmt[1] == ("lowerRoman", "1") and fmt[2] == ("decimal", "1"), fmt)
    cols = [(s._sectPr.find(qn("w:cols")).get(qn("w:num")) if s._sectPr.find(qn("w:cols")) is not None else "1") for s in secs]
    check("所有分节单栏", all(c == "1" for c in cols), cols)

    titles = [p for p in doc.paragraphs if p.style.name == "Abstract Title"]
    check("摘要标题套摘要样式（中英各一）", len(titles) == 2, [t.text for t in titles])
    abs_shd = 0
    for p in doc.paragraphs:
        if p.style.name in ("Abstract Title", "Abstract Text", "Keywords"):
            abs_shd += sum(1 for _ in p._p.iter(qn("w:shd")))
    check("摘要没有底纹", abs_shd == 0, abs_shd)
    check("摘要没有外框表格", not any(len(t.rows) == 1 and len(t.columns) == 1 for t in doc.tables))

    h1 = [p.text for p in doc.paragraphs if p.style.name == "Heading 1"]
    h2 = [p.text for p in doc.paragraphs if p.style.name == "Heading 2"]
    check("一级标题编号为“一、”", h1 and h1[0].startswith("一、"), h1[:1])
    check("二级标题编号为“（一）”", h2 and h2[0].startswith("（一）"), h2[:1])
    check("参考文献不编号", "参考文献" in h1, h1[-1:])

    toc = [p for p in doc.paragraphs if p.style.name in ("toc 1", "toc 2", "toc 3")]
    heads = [p for p in doc.paragraphs if re.match(r"^Heading [123]$", p.style.name)]
    check("目录条目数与 1–3 级标题数一致", len(toc) == len(heads), "%d vs %d" % (len(toc), len(heads)))
    instr = [t.text for t in doc.element.body.iter(qn("w:instrText"))]
    check("目录是 TOC 域", any("TOC" in (t or "") for t in instr))

    def border_ok(tbl):
        b = tbl._tbl.tblPr.find(qn("w:tblBorders"))
        return b is not None and b.find(qn("w:top")).get(qn("w:val")) == "single" and b.find(qn("w:insideV")).get(qn("w:val")) == "nil"
    check("表格为三线表", doc.tables and all(border_ok(t) for t in doc.tables if len(t.columns) > 2))

    pdf = out / "preview" / "sample_preview.pdf"
    if pdf.exists():
        txt = subprocess.run(["pdftotext", "-layout", str(pdf), "-"], capture_output=True, text=True).stdout
        pages = txt.split("\f")
        if pages and not pages[-1].strip():
            pages = pages[:-1]
        leader = re.compile(r"(\.\s*){4,}")
        check("封面页没有页码", not re.fullmatch(r"[ivx\d]+", last_line(pages[0])), last_line(pages[0]))
        check("目录页页码为 i", last_line(pages[1]) == "i", last_line(pages[1]))
        first_body = next((i for i, p in enumerate(pages) if any(l.strip().startswith("一、") and not leader.search(l) for l in p.splitlines())), None)
        check("正文首页页码为 1", first_body is not None and last_line(pages[first_body]) == "1",
              last_line(pages[first_body]) if first_body is not None else "找不到正文首页")
        nums = re.findall(r"\.{4,}\s*(\d+)\s*$", pages[1], re.M)
        check("目录页码回填为数字", len(nums) == len(toc), "%d 条" % len(nums))
    else:
        print("（没有预览 PDF：soffice 不可用，跳过页面检查）")

    passed = sum(results)
    print("\nSELFTEST RESULT: %d/%d 通过  输出目录：%s" % (passed, len(results), out))
    sys.exit(0 if passed == len(results) else 1)


if __name__ == "__main__":
    main()
