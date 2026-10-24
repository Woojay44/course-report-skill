#!/usr/bin/env python3
"""把 course-report/ 打包成 dist/course-report-<版本>.zip（zip 内顶层目录为 course-report/）。

用法：python tools/package.py v3 [--out 输出目录]
"""

import argparse
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SKILL = ROOT / "course-report"
SKIP_DIRS = {"__pycache__", ".pytest_cache"}
SKIP_SUFFIXES = {".pyc", ".pyo"}
SKIP_NAMES = {"defaults.local.yaml"}  # 个人署名，已 gitignore，不随包分发


def _package_files(skill):
    """返回要分发的文件；仅根据相对于 skill 的路径筛选。"""
    return sorted(p for p in skill.rglob("*") if p.is_file()
                   and not SKIP_DIRS & set(p.relative_to(skill).parts)
                   and p.suffix not in SKIP_SUFFIXES
                   and p.name not in SKIP_NAMES)


def _write_package(skill, out):
    files = _package_files(skill)
    out.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        for p in files:
            z.write(p, Path("course-report") / p.relative_to(skill))
    return len(files)


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("version", help="版本，如 v3")
    ap.add_argument("--out", type=Path, default=ROOT / "dist", help="输出目录（默认 dist/）")
    opts = ap.parse_args()
    out = opts.out / ("course-report-%s.zip" % opts.version)
    count = _write_package(SKILL, out)
    print("%s（%d 个文件）" % (out, count))


if __name__ == "__main__":
    main()
