# 更新记录

## v4（2026-10-12 / 2026-10-13）

**后端换成 LaTeX**
- `build_report.py` 改为 pandoc 出 LaTeX 片段 → 套 `templates/latex/report.tmpl.tex`（ctexart）→ xelatex 两遍出 PDF；不再输出 docx。
- 新增 `--class-opts`、`--no-cover`、`--no-toc`、`--keep-tex`，可在 `out/work/*.tex` 上手改后重编。

**清理 DOCX 时期遗留**
- `toolchain.py` 删除 LibreOffice / Word 渲染（`find_soffice`、`word_available`、`find_renderer`、`render_pdf`）。
- `probe_env.py` 不再探测 soffice / Word，改为检查 xelatex 与 TeX 宏包。
- `requirements.txt` 去掉 `python-docx`；README、仓库简介改为 LaTeX / PDF。

## v3（2026-10-11）

基于 v2（2026-10-02）。问题编号见 [issues.md](issues.md)。

**能跑通**
- 修复 `build_report.py` 入口缩进：路线 B 原来什么都不做（1）。
- 修复图片路径检查崩溃，并支持带标题的图片写法（2）；占位符检查改为统计“【待补充”（3）。

**版面**
- “无首行缩进”同时写字符单位和长度单位，表格、题注、图片、封面不再继承正文的 2 字符缩进（4）。
- 封面信息表不套表格样式首行格式，去掉多余横线（5）；列宽按内容加宽，不超版心（6）。
- 正文列表按普通段落排：编号占首行缩进、折行顶格、逐级缩进、编号后接空格；带编号的标题不动（21）。

**自检**
- 构建成功须生成 docx（11）；新增版面检查 10 项、文档属性检查 3 项（12）。
- 结果三态 PASS / FAIL / SKIP，页面检查做不了记 SKIP；`--require-pages` 严格模式（13）。

**Windows**
- 新增 `scripts/toolchain.py`：pandoc 查找（PATH、`PANDOC`、`pypandoc_binary`）、渲染器（LibreOffice，Windows 上回退到 Word）、PDF 文字与页面图像、已安装字体（注册表）（7、9、10、22、23）。
- 新增 `scripts/preview_pages.py` 代替 `pdftoppm`（8）。
- UTC+8 时区缺 `tzdata` 时静默按固定偏移计算（20）。

**文档属性**
- 作者、最后修改者写封面姓名（`--author`），标题写封面题目；不再残留 `numbering` 自定义属性（15）。

**说明文档**
- SKILL.md 明确排版模式 / 写作模式；补充渲染器、预览图片、自检三态、文档属性说明；版面规范补列表、封面列宽、文档属性规则（17）。

**不改**
- 默认署名（`templates/defaults.yaml`）是用户本人的默认值，有意保留（14）。
