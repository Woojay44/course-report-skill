#!/usr/bin/env python3
"""打包回归测试；源目录只读，所有测试产物写入临时目录。"""

import contextlib
import io
import os
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest
from unittest.mock import patch
import zipfile

# 导入被测脚本时也不在源码目录留下字节码。
sys.dont_write_bytecode = True
import package


class PackageTests(unittest.TestCase):
    def test_fake_skill_filters_and_archive_root(self):
        with tempfile.TemporaryDirectory() as td:
            root = Path(td)
            skill = root / "fake-skill"
            included = {"SKILL.md", "scripts/x.py"}
            excluded = {"templates/defaults.local.yaml", "__pycache__/x.pyc",
                        "a.pyc", "a.pyo", ".pytest_cache/cache.txt"}
            for name in included | excluded:
                path = skill / name
                path.parent.mkdir(parents=True, exist_ok=True)
                path.write_text("test fixture", encoding="utf-8")

            selected = {p.relative_to(skill).as_posix() for p in package._package_files(skill)}
            self.assertEqual(selected, included)
            out = root / "output" / "course-report-test.zip"
            self.assertEqual(package._write_package(skill, out), len(included))
            with zipfile.ZipFile(out) as archive:
                self.assertEqual(set(archive.namelist()),
                                 {"course-report/" + name for name in included})

    def test_real_skill_cli_out(self):
        with tempfile.TemporaryDirectory() as td:
            outdir = Path(td) / "packages"
            result = subprocess.run(
                [sys.executable, "-B", str(package.ROOT / "tools" / "package.py"),
                 "vtest", "--out", str(outdir)],
                cwd=td, capture_output=True, text=True, encoding="utf-8", errors="replace",
                env=dict(os.environ, PYTHONDONTWRITEBYTECODE="1", PYTHONIOENCODING="utf-8"))
            self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
            with zipfile.ZipFile(outdir / "course-report-vtest.zip") as archive:
                names = set(archive.namelist())
            self.assertFalse(any(Path(name).name == "defaults.local.yaml" for name in names))
            self.assertTrue({"course-report/SKILL.md", "course-report/scripts/build_report.py",
                             "course-report/templates/defaults.yaml"} <= names)
            self.assertTrue(all(name.startswith("course-report/") for name in names))

    def test_default_output_path(self):
        with tempfile.TemporaryDirectory() as td:
            root = Path(td)
            skill = root / "fake-skill"
            skill.mkdir()
            (skill / "SKILL.md").write_text("test fixture", encoding="utf-8")
            with patch.object(package, "ROOT", root), patch.object(package, "SKILL", skill), \
                    patch.object(sys, "argv", ["package.py", "vtest"]), \
                    contextlib.redirect_stdout(io.StringIO()):
                package.main()
            with zipfile.ZipFile(root / "dist" / "course-report-vtest.zip") as archive:
                self.assertEqual(archive.namelist(), ["course-report/SKILL.md"])


if __name__ == "__main__":
    unittest.main()
