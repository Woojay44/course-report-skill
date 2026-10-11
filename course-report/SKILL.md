---
name: course-report
description: 中文课程报告的排版与交付：把结构化 Markdown 或已有 DOCX 排成带封面、自动目录、论文式摘要、页码分节（前置页罗马数字、正文从 1 开始）的规范 Word 文档，并做渲染检查。用户要求排版、加封面、加目录、改摘要格式、统一页码，或让你根据内容写课程报告并排版时触发。
---

优先级：用户明确要求 > 课程模板或格式说明 > 原稿一致格式 > 本 skill 默认格式。

本 skill 目前的交付后端是 **DOCX**：课程要求宋体/黑体，用户在本机 Word 里更新目录后自行导出 PDF。LaTeX 后端尚未提供，不要假装有。

## 先判断走哪条路

| 用户给的是 | 走 |
|---|---|
| 已有 `.docx`，要改排版 | 路线 A：`format_docx.py` |
| 已有 `.md`/纯文本，或让你根据项目材料写报告 | 路线 B：写成结构化 Markdown → `build_report.py` |

开始前先运行一次 `python scripts/probe_env.py`，读一下“注意”行（缺 pandoc、python-docx、soffice 等会影响哪一步）。

## 路线 B：结构化 Markdown → DOCX

1. 按 `references/markdown-spec.md` 写 `report.md`：front matter 写标题、摘要、关键词；正文用 `#`/`##`/`###` 标题，**不要手打编号**（脚本按“一、（一）1.”自动加）。
2. 运行：

   ```
   python scripts/build_report.py report.md --out out/
   ```

3. 姓名、学号、时区取自 `templates/defaults.yaml`；`date: auto`（默认）按该时区取当天。用户本次另有说明时，在 front matter 里用 `author`、`student_id`、`cover_fields` 覆盖，不要改 defaults 文件。
4. 读脚本输出：带“警告”“提示”的行必须处理或在交付时说明。
5. 逐页检查 `out/preview/*_preview.pdf`（`pdftoppm -r 60 -png` 渲染后看图），见下方“验收”。
6. 交付 `out/<名称>.docx`，不交预览 PDF。

### 项目材料起草（基础规则，完整写作模式将单独提供）

- 只写材料能支撑的事实；数据、结果、结论都能对应到用户给的材料。
- 缺信息写 `【待补充：……】`，不编造实验结果、数据、参考文献；交付时列出所有待补充项。
- 图表只用真实数据生成；参考文献只用用户提供的，或有把握核实的。
- 课程是否要求声明 AI 辅助，由用户决定，不自动添加声明。

## 路线 A：已有 DOCX 改排版

1. 先按“识别”列清：封面、摘要、目录、标题、正文、列表、图表、公式、参考文献。
2. 标题套 `Heading 1–3`，题注套 `Caption`，缺样式先补；不要手动拼目录。
3. 封面：原稿已有 → `--existing-cover`；没有 → 写 `cover.json` 后 `--cover cover.json`（值写 `"auto"` 表示当天日期）。
4. 无模板时的完整命令：

   ```
   python scripts/format_docx.py 输入.docx 输出.docx --defaults --fix-abstract \
     --cover cover.json --toc --page-numbers
   ```

   有模板：把 `--defaults` 换成 `--black-only`，去掉 `--page-numbers`。表格要转三线表：加 `--three-line all` 或 `--three-line 1,3`（序号按原稿中表格出现顺序，摘要外框表格也算）。
5. 读日志：“已删除旧封面内容”里的每一项都要已写进 `cover.json`；“警告”行按提示处理。
6. 公式转 Word 原生公式、列表转原生列表，脚本不处理，需手工保留原编号和内容。

## 版面（详见 `references/layout-spec.md`）

封面 → 目录 → 摘要（中文一页，有英文则再一页）→ 正文。封面无页码；目录和摘要罗马数字；正文从 1。目录单栏连续、点线引导、行距 1.75，可多页。摘要无底纹无边框。

## 验收（必做，未渲染的页面不得声称已检查）

逐页看：封面单独一页无页码；目录紧跟封面、单栏连续、翻页处没有异常大片留白；摘要无灰底；正文首页页码为 1；图在图题之上、表题在表之上；含图表页无越界；参考文献页和末页无异常空白。可运行 `python scripts/selftest.py` 做结构自检。

## 交付说明（至多三行）

1. 做了哪些假设；哪些项目留空、待补充或有警告。
2. 固定提醒：打开 Word 后右键目录 → 更新域 → 更新整个目录，再导出 PDF（预览用的是渲染环境字体，页码可能与 Word 差一两页）。
3. 若预览里公式是空白：这是预览环境缺少 LibreOffice Math，docx 里是 Word 原生公式，以 Word 为准。

## 禁止事项

- 排版模式：不改写、扩写、缩写、删减正文；封面、目录只来自已有内容。
- 写作模式：不编造事实、数据、引用。
- 不用空格或连续空段落模拟缩进、居中、分页、封面留白。
- 不把 LaTeX 源码或公式截图当作最终公式。
- 不手拼目录、不手写分页/分栏；出现断页或分栏先看 `references/troubleshooting.md`。
