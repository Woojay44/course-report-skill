# 常见问题（LaTeX 后端）

中间文件与日志在 `out/work/` 下：`<名称>.tex`（套好模板的完整源）、`<名称>.log`（xelatex 日志）、`body_raw.tex`（pandoc 原始片段）。排查从这三者入手。

## xelatex 编译失败

脚本会打印日志里含 `!` / `Error` 的行，并指向 `out/work/<名称>.log`。常见原因：

- **`File 'xxx.sty' not found`**：缺 TeX 宏包。TinyTeX 用 `tlmgr install <包名>` 安装（先 `tlmgr update --self`）。先跑 `python scripts/probe_env.py` 看“缺 TeX 宏包”一行列出的全部缺失项。
- **`Undefined control sequence`**：多为片段里出现了模板未定义的宏。模板导言区已自备 `\tightlist`、`\real`、`\passthrough`、`\pandocbounded`；若 pandoc 片段用到别的宏，在 `templates/latex/report.tmpl.tex` 导言区补 `\providecommand`。
- **路径里的反斜杠被当控制序列**：传给 xelatex 的路径必须用正斜杠（`toolchain.run_xelatex` 已处理；手动调用时注意）。

## 图片找不到 / 图片越界

- 找不到：图片路径相对 `report.md`（如 `screenshots/a.png`）。脚本把 xelatex 工作目录设为 `report.md` 所在目录，并设 `\graphicspath` 指向该目录绝对路径。确认图片确实在该相对位置（脚本的 lint 也会提示“图片文件不存在”）。
- 越界：模板用 `\pandocbounded` 把图片限制为不超过版心宽、0.8 页高。若仍嫌大，可在 Markdown 里改图或后续调小该上限。

## 代码框报错或语言标签问题

listings 不认识 `javascript`、`apache`、`php` 等语言标签，脚本在 pandoc 转换后会自动删除 `language=...`。若新增了别的高亮需求，改 `build_report.py` 的 `strip_listings_language` 或模板的 `\lstset`。长行已开启 `breaklines` 自动换行。

## 目录页码没出来 / 不对

xelatex 需要跑两遍才能生成目录页码，脚本默认两遍。若手动编译，记得编两次。

## 公式不显示

行内 `$...$`、独立 `$$...$$` 由 xelatex 直接排版。用到 `\sum`、`\frac` 等需要 amsmath 的命令时，在模板导言区加 `\usepackage{amsmath}`（基础符号无需）。

## 字体是替代字体

环境里没有 SimSun / SimHei 时，ctexart 会退回替代字体（probe_env 会提示）。装上宋体 / 黑体后重跑即可。

## 摘要 / 关键词没出现

摘要、关键词取自 front matter 的 `abstract` / `keywords`（英文为 `abstract_en` / `keywords_en`）。确认 front matter 里有这些字段；正文里不要再手写“摘要”标题。
