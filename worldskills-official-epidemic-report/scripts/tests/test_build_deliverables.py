"""build 流程门禁：前置 P0-P2、外部生成器、P4、P5 渲染哈希核验。

仅用 stdlib；builder/renderer 为测试临时脚本（复制预生成交付件 / 计算真实 sha），
不访问网络、不爬虫，也不触碰任何真实用户资料。
"""

from __future__ import annotations

import hashlib
import json
import tempfile
import unittest
from pathlib import Path

from _helpers import PYTHON, paragraph, run_script, write_docx, write_json

import build_deliverables as bd

BUILDER = (
    "import sys, shutil\n"
    "src_docx, dst_docx, src_xlsx, dst_xlsx = sys.argv[1:5]\n"
    "shutil.copy2(src_docx, dst_docx)\n"
    "shutil.copy2(src_xlsx, dst_xlsx)\n"
)
RENDERER = (
    "import sys, hashlib, json\n"
    "docx, rc = sys.argv[1:3]\n"
    "h = hashlib.sha256()\n"
    "with open(docx, 'rb') as f:\n"
    "    for b in iter(lambda: f.read(1 << 20), b''):\n"
    "        h.update(b)\n"
    "json.dump({'docx_sha256': h.hexdigest(), 'checked': True, 'pages': 3, 'issues': []},\n"
    "          open(rc, 'w', encoding='utf-8'))\n"
)
RENDERER_BAD = (
    "import sys, json\n"
    "docx, rc = sys.argv[1:3]\n"
    "json.dump({'docx_sha256': '0' * 64, 'checked': True, 'pages': 3, 'issues': []},\n"
    "          open(rc, 'w', encoding='utf-8'))\n"
)


class TempCase(unittest.TestCase):
    def setUp(self) -> None:
        self._temp = tempfile.TemporaryDirectory()
        self.root = Path(self._temp.name)
        self.addCleanup(self._temp.cleanup)
        from _fixtures import (
            build_complete_env,
            research_payload_complete,
            write_registry,
            write_report,
            write_workbook,
            write_json,
        )

        # 使用完整环境（所有病种均有本地简介、无 null），否则 P0c 会按新规则阻断
        self.env = build_complete_env(self.root / "env")
        # 预生成一份规范的交付件，供 builder 复制
        write_report(self.env)
        write_workbook(self.env)
        write_json(self.env.research_json, research_payload_complete(self.env))
        write_registry(self.env)
        self.builder_py = self.root / "builder.py"
        self.renderer_py = self.root / "renderer.py"
        self.renderer_bad_py = self.root / "renderer_bad.py"
        self.builder_py.write_text(BUILDER, encoding="utf-8")
        self.renderer_py.write_text(RENDERER, encoding="utf-8")
        self.renderer_bad_py.write_text(RENDERER_BAD, encoding="utf-8")

    def _base_args(self, out_docx: Path, out_xlsx: Path, render_check: Path) -> list[str]:
        return [
            "--registry",
            str(self.env.registry_json),
            "--research",
            str(self.env.research_json),
            "--xlsx",
            str(out_xlsx),
            "--docx",
            str(out_docx),
            "--render-check",
            str(render_check),
            "--list-file",
            str(self.env.list_file),
            "--disease-config",
            str(self.env.disease_config),
            "--intro-index",
            str(self.env.intro_index),
            "--intro-dir",
            str(self.env.intro_dir),
            "--local-dir",
            str(self.env.local_dir),
            "--template",
            str(self.env.docx),
        ]

    def _run_build(self, out_docx: Path, out_xlsx: Path, render_check: Path, bad_renderer: bool = False) -> int:
        src_docx, src_xlsx = str(self.env.docx), str(self.env.xlsx)
        renderer = self.renderer_bad_py if bad_renderer else self.renderer_py
        builder_tail = [
            "--builder",
            str(PYTHON),
            str(self.builder_py),
            src_docx,
            str(out_docx),
            src_xlsx,
            str(out_xlsx),
            "--renderer",
            str(PYTHON),
            str(renderer),
            "{docx}",
            "{render_check}",
        ]
        return run_script(
            "build_deliverables.py", *self._base_args(out_docx, out_xlsx, render_check), *builder_tail
        ).returncode


class TestRenderCheckUnit(unittest.TestCase):
    def setUp(self) -> None:
        self._temp = tempfile.TemporaryDirectory()
        self.root = Path(self._temp.name)
        self.addCleanup(self._temp.cleanup)

    def test_missing_render_check_file(self) -> None:
        docx = write_docx(self.root / "d.docx", [paragraph("x")])
        rc = self.root / "missing.json"
        errors = bd.validate_render_check(rc, docx)
        self.assertTrue(any("不存在" in e for e in errors))

    def test_wrong_sha(self) -> None:
        docx = write_docx(self.root / "d.docx", [paragraph("x")])
        rc = write_json(self.root / "rc.json", {"docx_sha256": "0" * 64, "checked": True, "pages": 3, "issues": []})
        errors = bd.validate_render_check(rc, docx)
        self.assertTrue(any("不是同一文件" in e for e in errors))

    def test_pages_not_positive(self) -> None:
        docx = write_docx(self.root / "d.docx", [paragraph("x")])
        rc = write_json(self.root / "rc.json", {"docx_sha256": hashlib.sha256(docx.read_bytes()).hexdigest(), "checked": True, "pages": 0, "issues": []})
        errors = bd.validate_render_check(rc, docx)
        self.assertTrue(any("pages" in e for e in errors))

    def test_issues_nonempty(self) -> None:
        docx = write_docx(self.root / "d.docx", [paragraph("x")])
        rc = write_json(self.root / "rc.json", {"docx_sha256": hashlib.sha256(docx.read_bytes()).hexdigest(), "checked": True, "pages": 3, "issues": ["排版异常"]})
        errors = bd.validate_render_check(rc, docx)
        self.assertTrue(any("问题" in e for e in errors))

    def test_matching_sha_passes(self) -> None:
        docx = write_docx(self.root / "d.docx", [paragraph("x")])
        rc = write_json(self.root / "rc.json", {"docx_sha256": hashlib.sha256(docx.read_bytes()).hexdigest(), "checked": True, "pages": 3, "issues": []})
        self.assertEqual(bd.validate_render_check(rc, docx), [])


class TestBuildGates(TempCase):
    def test_missing_builder_arg_errors(self) -> None:
        out_docx = self.root / "out.docx"
        out_xlsx = self.root / "out.xlsx"
        rc = self.root / "rc.json"
        result = run_script(
            "build_deliverables.py",
            "--registry",
            str(self.env.registry_json),
            "--research",
            str(self.env.research_json),
            "--xlsx",
            str(out_xlsx),
            "--docx",
            str(out_docx),
            "--render-check",
            str(rc),
            "--list-file",
            str(self.env.list_file),
            "--disease-config",
            str(self.env.disease_config),
        )
        self.assertNotEqual(result.returncode, 0)

    def test_end_to_end_success(self) -> None:
        out_docx = self.root / "out.docx"
        out_xlsx = self.root / "out.xlsx"
        rc = self.root / "rc.json"
        self.assertEqual(self._run_build(out_docx, out_xlsx, rc), 0)
        self.assertTrue(out_docx.is_file() and out_xlsx.is_file())
        self.assertTrue(rc.is_file())

    def test_render_hash_mismatch_fails(self) -> None:
        out_docx = self.root / "out.docx"
        out_xlsx = self.root / "out.xlsx"
        rc = self.root / "rc.json"
        self.assertEqual(self._run_build(out_docx, out_xlsx, rc, bad_renderer=True), 1)

    def test_source_template_not_modified(self) -> None:
        before = hashlib.sha256(self.env.docx.read_bytes()).hexdigest()
        out_docx = self.root / "out.docx"
        out_xlsx = self.root / "out.xlsx"
        rc = self.root / "rc.json"
        self._run_build(out_docx, out_xlsx, rc)
        after = hashlib.sha256(self.env.docx.read_bytes()).hexdigest()
        self.assertEqual(before, after)  # 源模板/源交付件未被修改


if __name__ == "__main__":
    unittest.main()
