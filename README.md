# course-report-skill

中文课程报告排版 skill（`course-report`）的维护项目：把结构化 Markdown 或已有 DOCX 排成带封面、自动目录、论文式摘要、分节页码的 Word 文档。

本项目从 `D:\Project-codex\chinese-course-docx-formatting\archives\course-report-v2.zip` 导入，首个提交即原版 v2，后续改进见 git 历史。已知问题见 [docs/issues.md](docs/issues.md)。

## 依赖

- Python 3.10+，依赖见 `requirements.txt`（`pip install -r requirements.txt`）
- pandoc ≥ 3.0：系统安装的，或 `pypandoc_binary` 自带的，或用环境变量 `PANDOC` 指定
- 渲染器（预览、目录页码回填）：LibreOffice；Windows 上没有时自动用本机 Word

## 目录结构

```
course-report/          skill 本体（SKILL.md、scripts/、references/、templates/、tests/）
docs/issues.md          v2 问题清单（✅ 为已修复）
docs/CHANGELOG.md       版本更新记录
tools/package.py        打包成 dist/course-report-<版本>.zip
```

## 个人默认署名

`course-report/templates/defaults.yaml` 是示例值。复制为同目录下的 `defaults.local.yaml` 再改成自己的姓名、学号（该文件已在 `.gitignore` 中），脚本会优先读取它。

## 运行

```powershell
python course-report/scripts/probe_env.py      # 环境探测
python course-report/scripts/selftest.py --require-pages   # 自检（34 项，含页面）
python course-report/scripts/build_report.py report.md --out out/
python course-report/scripts/preview_pages.py out/preview/report_preview.pdf   # 逐页 PNG
```

## 打包

```powershell
python tools/package.py v3      # → dist/course-report-v3.zip
```
