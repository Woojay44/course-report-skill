---
name: course-report
description: 中文课程报告的排版与交付：把结构化 Markdown 排成带封面、自动目录、论文式摘要、页码分节（前置页罗马数字、正文从 1 开始）的规范 PDF（LaTeX / ctexart + xelatex 排版），并逐页做版面检查。用户要求排版、加封面、加目录、改摘要格式、统一页码，或让你根据内容写课程报告并排版时触发。
---

优先级：用户明确要求 > 课程模板或格式说明 > 原稿一致格式 > 本 skill 默认格式。

两种模式，开始前先判断：

- **排版模式**（默认）：用户给了现成内容（`.md`、纯文本），只改版面，不改内容。
- **写作模式**：用户明确要你根据项目材料（代码、数据、运行结果、实验要求等）写报告。按下方“项目材料起草”的规则写成结构化 Markdown，再排版。

交付后端是 **LaTeX**：用 `ctexart + xelatex` 排版，直接出 **PDF**（week07 风格：titlepage 封面、booktabs 三线表、listings 代码框、fancyhdr 页眉页脚）。不再输出 docx。

## 流程总览

结构化 Markdown → `build_report.py`：读 front matter → 生成封面/摘要 → pandoc 把正文转成 LaTeX 片段（Lua 过滤器加中文编号与图表号，listings 排代码）→ 套 `templates/latex/report.tmpl.tex` 模板 → xelatex 跑两遍出 PDF。

开始前先运行一次 `python scripts/probe_env.py`，读“注意”行（缺 pandoc / xelatex / TeX 宏包 / 字体会影响哪一步）。

- pandoc：系统安装的、`pip install pypandoc_binary` 自带的、或环境变量 `PANDOC` 指定的都可以（需 3.0+，图表编号过滤器依赖它）。
- xelatex：TeX Live / MiKTeX / TinyTeX 均可，或环境变量 `XELATEX` 指定。TinyTeX 是精简版，缺宏包时用 `tlmgr install <包名>` 补（probe_env 会列出缺哪些）。
- 字体：ctexart 默认用宋体 SimSun / 黑体 SimHei；缺失会退回替代字体，probe_env 会提示。

## 排版步骤

1. 按 `references/markdown-spec.md` 写 `report.md`：front matter 写标题、摘要、关键词、封面字段；正文用 `#`/`##`/`###` 标题，**不要手打编号**（脚本按“一、（一）1.”自动加）。正文引用写 `[1]`，多篇连引写 `[4][5]`（不写 `[4,5]`），自动排成右上角上标角标。
2. 运行：

   ```
   python scripts/build_report.py report.md --out out/
   ```

3. 姓名、学号、时区取自 `templates/defaults.local.yaml`（用户本人的默认署名，不进 git；没有时退回 `templates/defaults.yaml` 的示例值，脚本会提示）；`date: auto`（默认）按该时区取当天。用户本次另有说明（例如换人署名）时，在 front matter 里用 `author`、`student_id`、`cover_fields` 覆盖，不要改 defaults 文件。用户要求修改自己的默认署名时，改 `defaults.local.yaml`。
4. 读脚本输出：带“警告”“提示”的行必须处理或在交付时说明（包括“还有 N 处占位符”“图片文件不存在”）。xelatex 编译报错时，脚本会打印日志里的报错行，并指向 `out/work/<名称>.log`。
5. 逐页检查 PDF：`python scripts/preview_pages.py out/<名称>.pdf` 把每页转成 PNG 后看图，见下方“验收”。
6. 交付 `out/<名称>.pdf`。中间文件（`.tex`、`.log` 等）在 `out/work/` 下，便于排查，不必交付。

可选参数：`--numbering cn|decimal|manual`（标题编号）、`--class-opts "11pt,a4paper"`（文档类选项）、`--no-cover`、`--no-toc`、`--keep-tex`。

### 项目材料起草（基础规则，完整写作模式将单独提供）

- 只写材料能支撑的事实；数据、结果、结论都能对应到用户给的材料。
- 缺信息写 `【待补充：……】`，不编造实验结果、数据、参考文献；交付时列出所有待补充项。
- 图表只用真实数据生成；参考文献只用用户提供的，或有把握核实的。
- 课程是否要求声明 AI 辅助，由用户决定，不自动添加声明。

#### 先判断报告类别

- **法学类**（案例评析、法条分析、法律实务等）：
  1. 先查清案情和裁判要旨，判决书原文、法律条文尽量核实。
  2. 表明观点并做辨析：写出对立观点各自的理由，相似案例与本案的区别。
  3. 正文用 `[n]` 标注引用（见上，自动排成上标角标）。
  4. 末尾写“参考文献”，以法律法规、司法解释、裁判文书、审判指南和学术论文/专著为主；每条都要能核实，查不到的信息（如案号）就不写，不编造。
- **技术类**（网安、实验报告等）：参考文献不必卡这么严，按内容需要列出（官方文档、工具、标准、文章等）即可。
- 其他学科按本节基础规则来；拿不准类别或命名规则时先问用户。

## 图片与代码的注意事项（已在模板/脚本中处理，排查时参考）

- 图片路径相对 `report.md`（如 `screenshots/a.png`）。脚本把 xelatex 的工作目录设为 `report.md` 所在目录，相对路径即可找到；并辅以 `\graphicspath` 指向该目录绝对路径。图片用自定义的 `\pandocbounded` 限制为不超过版心宽、0.8 页高，防越界。
- 代码围栏的语言标签（`javascript`、`apache`、`php` 等）listings 不认识，脚本在转换后会自动去掉 `language=...`。代码框开启了自动换行，长行不会越界。
- 模板导言区已自备 pandoc 片段需要的辅助宏：`\tightlist`、`\real`、`\passthrough`、`\pandocbounded`。
- 封面字段、摘要、标题等注入模板的文本都做了 LaTeX 转义；正文由 pandoc 自行转义。
- xelatex 跑两遍才能生成目录页码（脚本默认两遍）。

## 版面（详见 `references/layout-spec.md`）

封面 → 目录 → 摘要（中文，有英文则再接一段）→ 正文。封面无页码；目录和摘要用罗马数字；正文从 1（阿拉伯数字）。目录单栏连续、点线引导，可多页。摘要无底纹无边框。三线表用 booktabs，表题在表上；图题在图下。

## 验收（必做，未渲染的页面不得声称已检查）

`python scripts/preview_pages.py out/<名称>.pdf` 逐页转 PNG 后逐页看：封面单独一页无页码、信息表字段不折行；目录紧跟封面、单栏连续、点线引导、页码正确、翻页处没有异常大片留白；摘要无灰底无边框；正文首页页码为 1；三线表为 booktabs 三线、题注在表上、数字没被挤断；代码框不越界、长行已换行；图在图题之上、含图表页无越界；参考文献页和末页无异常空白。

改过脚本或模板、或怀疑环境有问题时运行 `python scripts/selftest.py`：用 `tests/sample.md` 跑一遍完整流程，检查结构与版面共 15 项，结果为 PASS / FAIL。

## 交付说明（至多三行）

1. 做了哪些假设；哪些项目留空、待补充或有警告。
2. 交付的是 `out/<名称>.pdf`，可直接提交。
3. 若版面有需用户确认的取舍（例如某张图单独占页、某表跨页续排），一句话说明。

## 禁止事项

- 排版模式：不改写、扩写、缩写、删减正文；封面、目录只来自已有内容。
- 写作模式：不编造事实、数据、引用。
- 不用空格或连续空段落模拟缩进、居中、分页、封面留白。
- 不手拼目录、不手写分页/分栏；出现断页或异常留白先看 `references/troubleshooting.md`。
