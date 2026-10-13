# course-report-skill

中文课程报告排版 skill（`course-report`）的维护项目：把结构化 Markdown 用 LaTeX（`ctexart` + xelatex）排成带封面、自动目录、论文式摘要、分节页码的规范 PDF。

> v4 起后端从 DOCX 换成 LaTeX，只输出 PDF；v3 及以前的 DOCX 版本见 git 历史。

首个提交是原版 v2（DOCX 后端），后续改进见 git 历史和 [docs/CHANGELOG.md](docs/CHANGELOG.md)。

## 依赖

- Python 3.10+，依赖见 `requirements.txt`（`pip install -r requirements.txt`）
- pandoc ≥ 3.0：系统安装的，或 `pypandoc_binary` 自带的，或用环境变量 `PANDOC` 指定
- xelatex：TeX Live / MiKTeX / TinyTeX 均可，或用环境变量 `XELATEX` 指定；缺宏包时 `probe_env.py` 会列出
- 字体：宋体 SimSun、黑体 SimHei（缺失时 ctexart 退回替代字体）

## 目录结构

```
course-report/          skill 本体（SKILL.md、scripts/、references/、templates/、tests/）
docs/issues.md          v2（DOCX 版）问题清单，留作历史
docs/CHANGELOG.md       版本更新记录
tools/package.py        打包成 dist/course-report-<版本>.zip
```

## 个人默认署名

`course-report/templates/defaults.yaml` 是示例值。复制为同目录下的 `defaults.local.yaml` 再改成自己的姓名、学号（该文件已在 `.gitignore` 中），脚本会优先读取它。

## 运行

```powershell
python course-report/scripts/probe_env.py      # 环境探测
python course-report/scripts/selftest.py       # 自检（LaTeX 全流程 + PDF 结构与版面）
python course-report/scripts/build_report.py report.md --out out/   # → out/report.pdf
python course-report/scripts/preview_pages.py out/report.pdf        # 逐页 PNG
```

## 打包

```powershell
python tools/package.py v4      # → dist/course-report-v4.zip
```

## 许可证

[MIT](LICENSE)
