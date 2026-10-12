# 结构化 Markdown 规范

## front matter（文件开头的 YAML）

```yaml
---
title: 报告题目                 # 必填
subtitle:                       # 可选
org: 某某大学                   # 可选，缺省取 defaults.yaml
course: 《课程名》课程报告      # 可选，缺省取 defaults.yaml
author: 张三                    # 可选，缺省取 defaults.local.yaml（没有则 defaults.yaml）的 name；用于封面署名与页眉
student_id: "2024000000000"     # 可选，缺省取 defaults.yaml（务必加引号）
cover_fields:                   # 可选：额外封面字段，顺序即显示顺序
  - 学院: 计算机学院
  - 指导教师: 李老师
date: auto                      # auto=按时区取当天；也可写死 2026年9月30日
abstract: |                     # 摘要，纯文本，段落之间空行
  第一段……

  第二段……
keywords: [关键词一, 关键词二]
abstract_en: |                  # 可选，有则另起一页排英文摘要
  English abstract ……
keywords_en: [keyword one, keyword two]
numbering: cn                   # cn | decimal | manual，缺省取 defaults.yaml
cover: true                     # false 则不生成封面
toc: true                       # false 则不生成目录
---
```

封面字段顺序固定：姓名、学号、`cover_fields`、写作日期。没有值的字段不显示。

## 正文允许的语法

- 标题：`#` 一级，`##` 二级，`###` 三级（到三级进目录）。**不要在标题文字里手打编号**，脚本按 `numbering` 自动加。`参考文献`、`致谢`不编号。
- 段落、有序/无序列表、`**加粗**`、`*斜体*`。
- 图：`![图题](figures/a.png)`，路径相对于 md 文件；脚本自动加“图 N　”，图题在图下。
- 表：管道表格，表后一行 `: 表题`；脚本自动加“表 N　”，表题在表上，表格自动转三线表。
- 公式：`$行内$`、`$$独立$$`，由 xelatex 直接排成数学公式。
- 代码块：三个反引号围栏。
- 参考文献：`# 参考文献` 下每条一段，自己写 `[1] ……`（该标题不编号，自动进目录）。

## 不要用

原始 HTML、原始 LaTeX、手打的章节编号、用空行或空格排版、在摘要里用格式标记（摘要按纯文本处理，`* _ #` 会被转义成普通字符）。

## 占位符

缺信息写 `【待补充：……】`。脚本发现残留会在日志里提示，交付前必须补全或向用户说明。
