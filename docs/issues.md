# v2 问题清单

基线：`D:\Project-codex\chinese-course-docx-formatting\archives\course-report-v2.zip`（2026-10-02），即本仓库首个提交。
来源：2026-10-11 用 v2 实际排版 iris-knn-tree-lab 的实验报告，加上对原版的 pyflakes、selftest、路线 A/B 实跑。

## 一、脚本跑不通

1. `build_report.py` 末尾 `if __name__ == "__main__"` 缩进进了 `main()` 里：路线 B 什么都不做，退出码 0。
2. `build_report.py` `lint()`：`re.findall` 有两个分组、返回元组，循环当成字符串用 → 有图片就 TypeError。
3. `build_report.py` `lint()`：`md_text.count("")` 恒非 0 → 每次都误报占位符（应为 `"【待补充"`）。

## 二、版面错误

4. `format_docx.py` `set_first_line(ppr, 0)` 只写 `firstLine="0"`，不覆盖从 Normal 继承的 `firstLineChars="200"`；Word 中字符单位优先 → 表格、封面、题注、图片段落仍缩进 2 字符：表格数字被挤断，封面标签“姓/名”折行。
5. 封面信息表继承 pandoc 默认表格样式 `Table` 的 `firstRow` 下边框 → 首行标签格下多出一段横线。
6. 封面值列固定 7.2 cm，稍长的字段值折行。

## 三、Windows 环境

7. 预览与目录页码回填只支持 soffice；有 Word 时不用 → 目录页码不回填，验收无法进行。
8. 验收依赖 `pdftoppm`，Windows 通常没有，无替代方案。
9. 字体检测只用 `fc-list`，Windows 上恒报缺宋体/黑体。
10. 只在 PATH 找 pandoc，不识别 `pypandoc_binary` 自带的 pandoc。

## 四、自检

11. “构建成功”只看退出码 → 问题 1 时仍 PASS。
12. 只查结构，不查问题 4–6 这类可见版面缺陷（修复 1–3 后 17/17 通过，但版面仍有问题）。
13. 无 soffice 时整段页面检查跳过，结果仍显示全部通过。

## 五、隐私与元数据

14. `templates/defaults.yaml` 与 `selftest.py` 写死真实姓名、学号，随包分发。
15. 输出 docx 残留自定义属性 `numbering=cn`、pandoc 样式名；作者为空，Word 再保存时“最后修改者”变成本机账户名。

## 六、源码管理

16. 原仓库根目录是最旧版本（2026-09-28，仅说明文字），最新 v2 只在 zip 里。
17. 两版定位矛盾：旧版“不负责撰写”，v2 支持“根据项目材料写报告”。
18. 原仓库无 git，三个 zip 命名不统一。
19. pyflakes：`format_docx.py` 未用 `sys`、`Section`；`probe_env.py` 未用 `docx`。
