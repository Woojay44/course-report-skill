# course-report-skill

中文课程报告排版 skill（`course-report`）的维护项目：把结构化 Markdown 或已有 DOCX 排成带封面、自动目录、论文式摘要、分节页码的 Word 文档。

本项目从 `D:\Project-codex\chinese-course-docx-formatting\archives\course-report-v2.zip` 导入，首个提交即原版 v2，后续改进见 git 历史。已知问题见 [docs/issues.md](docs/issues.md)。

## 依赖

- Python 3.10+，`python-docx`
- pandoc ≥ 3.0
- 可选：LibreOffice（预览、目录页码回填）

## 目录结构

```
course-report/          skill 本体（SKILL.md、scripts/、references/、templates/、tests/）
docs/issues.md          v2 问题清单
```

## 运行

```powershell
python course-report/scripts/probe_env.py      # 环境探测
python course-report/scripts/selftest.py       # 自检
python course-report/scripts/build_report.py report.md --out out/
```
