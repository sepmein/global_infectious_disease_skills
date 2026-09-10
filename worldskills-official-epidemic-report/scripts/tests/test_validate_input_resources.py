"""P0c 输入资源校验：简介索引键集合、简介文件存在性与 sha256、本地资料、模板。"""

from __future__ import annotations

import json
import tempfile
import unittest
from pathlib import Path

from _fixtures import DISEASES, build_complete_env, build_env
from _helpers import paragraph, run_script, write_docx

DISEASE_NAMES = [name for name, _ in DISEASES]


class TempCase(unittest.TestCase):
    def setUp(self) -> None:
        self._temp = tempfile.TemporaryDirectory()
        self.root = Path(self._temp.name)
        self.addCleanup(self._temp.cleanup)
        self.env = build_env(self.root / "env")
        # 生成一个真实存在的文件作为 --template 存在性检查的对象
        self.env.docx = write_docx(self.env.docx, [paragraph("模板占位文档")])

    def _run(self, extra=None) -> int:
        args = [
            "--disease-config",
            str(self.env.disease_config),
            "--intro-index",
            str(self.env.intro_index),
            "--intro-dir",
            str(self.env.intro_dir),
            "--local-dir",
            str(self.env.local_dir),
            "--template",
            str(self.env.docx),  # 任何存在的文件都可作为模板存在性检查的对象
        ]
        if extra:
            args += extra
        return run_script("validate_input_resources.py", *args).returncode

    def _stderr(self, extra=None) -> str:
        args = [
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
        if extra:
            args += extra
        return run_script("validate_input_resources.py", *args).stderr


class TestPositive(TempCase):
    def setUp(self) -> None:
        super().setUp()
        # 完整环境：所有病种都有本地简介文件（无 null），用于happy path
        self.complete = build_complete_env(self.root / "complete")
        self.complete.docx = write_docx(self.complete.docx, [paragraph("模板占位文档")])

    def _base_args(self, extra=None) -> list[str]:
        args = [
            "--disease-config",
            str(self.complete.disease_config),
            "--intro-index",
            str(self.complete.intro_index),
            "--intro-dir",
            str(self.complete.intro_dir),
            "--local-dir",
            str(self.complete.local_dir),
            "--template",
            str(self.complete.docx),
        ]
        if extra:
            args += extra
        return args

    def _run_complete(self, extra=None) -> int:
        return run_script("validate_input_resources.py", *self._base_args(extra)).returncode

    def test_valid_without_requirements_passes(self) -> None:
        self.assertEqual(self._run_complete(), 0)

    def test_require_template_passes_when_present(self) -> None:
        self.assertEqual(self._run_complete(["--require-template"]), 0)


class TestNullIntro(TempCase):
    def test_null_intro_fallback_warns_and_passes(self) -> None:
        """简介缺失（索引为 null）不再阻断：输出 fallback_required=true 并 warning，但退出 0。"""
        # build_env 中麻疹简介为 null
        result = run_script(
            "validate_input_resources.py",
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
        )
        self.assertEqual(result.returncode, 0)
        payload = json.loads(result.stdout)
        self.assertTrue(payload["intro_fallback_required"])
        self.assertIn("麻疹", payload["intro_missing"])
        self.assertIn("简介缺失", result.stderr)

    def test_null_intro_fallback_required_true_require_template(self) -> None:
        """--require-template 下，仅 null 简介也应退出 0（fallback 而非阻断）。"""
        result = run_script(
            "validate_input_resources.py",
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
            "--require-template",
        )
        self.assertEqual(result.returncode, 0)
        self.assertTrue(json.loads(result.stdout)["intro_fallback_required"])


class TestIntroIndex(TempCase):
    def test_missing_intro_file_fallback_warns_and_passes(self) -> None:
        """索引指向的文件不存在：允许以权威网络资料替代（warning + fallback），退出 0。"""
        index = json.loads(self.env.intro_index.read_text(encoding="utf-8"))
        index["登革热"] = "不存在的简介.pdf"
        self.env.intro_index.write_text(json.dumps(index, ensure_ascii=False), encoding="utf-8")
        result = run_script(
            "validate_input_resources.py",
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
        )
        self.assertEqual(result.returncode, 0)
        self.assertTrue(json.loads(result.stdout)["intro_fallback_required"])
        self.assertIn("索引指向的简介文件不存在", result.stderr)

    def test_key_missing_auto_maps_none_and_passes(self) -> None:
        """新增病种键缺失：自动映射为 None 并 warning，不阻断；仍不得伪造修复。"""
        bad = self.root / "bad_index.json"
        bad.write_text(json.dumps({"登革热": "x.pdf"}, ensure_ascii=False), encoding="utf-8")
        args = [
            "--disease-config",
            str(self.env.disease_config),
            "--intro-index",
            str(bad),
            "--intro-dir",
            str(self.env.intro_dir),
            "--local-dir",
            str(self.env.local_dir),
            "--template",
            str(self.env.docx),
        ]
        result = run_script("validate_input_resources.py", *args)
        self.assertEqual(result.returncode, 0)
        self.assertTrue(json.loads(result.stdout)["intro_fallback_required"])
        self.assertIn("自动映射", result.stderr)

    def test_extra_key_rejected(self) -> None:
        """配置外病种（额外键）仍拒绝：不得扩大范围。"""
        bad = self.root / "extra_index.json"
        bad.write_text(
            json.dumps({"登革热": "x.pdf", "麻疹": None, "鼠疫": "y.pdf"}, ensure_ascii=False),
            encoding="utf-8",
        )
        args = [
            "--disease-config",
            str(self.env.disease_config),
            "--intro-index",
            str(bad),
            "--intro-dir",
            str(self.env.intro_dir),
            "--local-dir",
            str(self.env.local_dir),
            "--template",
            str(self.env.docx),
        ]
        result = run_script("validate_input_resources.py", *args)
        self.assertEqual(result.returncode, 1)
        self.assertIn("配置外病种", result.stderr)

    def test_corrupt_index_blocks(self) -> None:
        """索引文件本身损坏（无法解析 JSON）仍阻断。"""
        bad = self.root / "corrupt_index.json"
        bad.write_text("{这不是合法json", encoding="utf-8")
        args = [
            "--disease-config",
            str(self.env.disease_config),
            "--intro-index",
            str(bad),
            "--intro-dir",
            str(self.env.intro_dir),
            "--local-dir",
            str(self.env.local_dir),
            "--template",
            str(self.env.docx),
        ]
        result = run_script("validate_input_resources.py", *args)
        self.assertEqual(result.returncode, 1)
        self.assertIn("JSON", result.stderr)

    def test_missing_intro_dir_all_fallback_and_passes(self) -> None:
        """整个简介目录不存在：所有本地简介视为缺失并 warning，退出 0（fallback）。"""
        missing_dir = self.root / "no_intro_dir"
        result = run_script(
            "validate_input_resources.py",
            "--disease-config",
            str(self.env.disease_config),
            "--intro-index",
            str(self.env.intro_index),
            "--intro-dir",
            str(missing_dir),
            "--local-dir",
            str(self.env.local_dir),
            "--template",
            str(self.env.docx),
        )
        self.assertEqual(result.returncode, 0)
        self.assertTrue(json.loads(result.stdout)["intro_fallback_required"])
        self.assertIn("疾病简介目录不存在", result.stderr)


class TestLocalAndTemplate(TempCase):
    def test_require_local_inputs_empty_dir_fails(self) -> None:
        empty = self.root / "empty_local"
        empty.mkdir()
        args = [
            "--disease-config",
            str(self.env.disease_config),
            "--intro-index",
            str(self.env.intro_index),
            "--intro-dir",
            str(self.env.intro_dir),
            "--local-dir",
            str(empty),
            "--template",
            str(self.env.docx),
            "--require-local-inputs",
        ]
        self.assertEqual(run_script("validate_input_resources.py", *args).returncode, 1)
        self.assertIn("为空", run_script("validate_input_resources.py", *args).stderr)

    def test_require_template_missing_fails(self) -> None:
        missing = self.root / "no_template.docx"
        args = [
            "--disease-config",
            str(self.env.disease_config),
            "--intro-index",
            str(self.env.intro_index),
            "--intro-dir",
            str(self.env.intro_dir),
            "--local-dir",
            str(self.env.local_dir),
            "--template",
            str(missing),
            "--require-template",
        ]
        self.assertEqual(run_script("validate_input_resources.py", *args).returncode, 1)
        self.assertIn("写作模板不存在", run_script("validate_input_resources.py", *args).stderr)


if __name__ == "__main__":
    unittest.main()
