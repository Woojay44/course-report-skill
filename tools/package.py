#!/usr/bin/env python3
"""把 course-report/ 打包成 dist/course-report-<版本>.zip（zip 内顶层目录为 course-report/）。

用法：python tools/package.py v3
"""

import sys
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SKILL = ROOT / "course-report"
SKIP_DIRS = {"__pycache__", ".pytest_cache"}
SKIP_SUFFIXES = {".pyc", ".pyo"}
SKIP_NAMES = {"defaults.local.yaml"}  # 个人署名，已 gitignore，不随包分发


def main():
    if len(sys.argv) != 2:
        sys.exit("用法：python tools/package.py <版本，如 v3>")
    out = ROOT / "dist" / ("course-report-%s.zip" % sys.argv[1])
    out.parent.mkdir(exist_ok=True)
    files = sorted(p for p in SKILL.rglob("*") if p.is_file()
                   and not SKIP_DIRS & set(p.relative_to(SKILL).parts)
                   and p.suffix not in SKIP_SUFFIXES
                   and p.name not in SKIP_NAMES)
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        for p in files:
            z.write(p, Path("course-report") / p.relative_to(SKILL))
    print("%s（%d 个文件）" % (out, len(files)))


if __name__ == "__main__":
    main()
