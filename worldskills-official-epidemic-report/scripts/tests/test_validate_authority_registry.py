"""P1 权威机构台账校验：访问失败只要求准确记录尝试，不得伪称可用，覆盖不放宽。"""

from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

import validate_authority_registry as reg

from _fixtures import COUNTRIES, build_env, registry_payload, write_registry
from _helpers import run_script

CODES = {code for code, _, _ in COUNTRIES}


class TempCase(unittest.TestCase):
    def setUp(self) -> None:
        self._temp = tempfile.TemporaryDirectory()
        self.root = Path(self._temp.name)
        self.addCleanup(self._temp.cleanup)
        self.env = build_env(self.root / "env")


class TestValidateFunction(TempCase):
    def test_valid_payload_passes(self) -> None:
        errors, stats = reg.validate(registry_payload()["records"], CODES)
        self.assertEqual(errors, [])
        # 日本为访问受限但记录完整，应计入“访问失败但记录完整”
        self.assertEqual(stats["访问失败但记录完整"], 1)
        self.assertEqual(stats["已验证可检索"], 1)

    def test_missing_required_field(self) -> None:
        payload = registry_payload()
        del payload["records"][0]["机构名称"]
        errors, _ = reg.validate(payload["records"], CODES)
        self.assertTrue(any("机构名称" in e for e in errors))

    def test_out_of_scope_code_rejected(self) -> None:
        payload = registry_payload()
        payload["records"].append(dict(payload["records"][0], 地区代码="ZZ", 中文标准名称="虚构", 英文名称="Zz"))
        errors, _ = reg.validate(payload["records"], CODES)
        self.assertTrue(any("名单外" in e for e in errors))

    def test_under_coverage_rejected(self) -> None:
        payload = registry_payload()
        payload["records"].pop()  # 删除一个，导致覆盖不足
        errors, _ = reg.validate(payload["records"], CODES)
        self.assertTrue(any("缺少动态名单中的" in e for e in errors))

    def test_pseudo_available_without_success_rejected(self) -> None:
        payload = registry_payload()
        rec = payload["records"][0]
        rec["直接访问状态"] = "连接超时（15s）"  # 失败记录，但结论伪称可用
        rec["浏览器回退状态"] = "浏览器加载失败：需要人机验证"
        rec["台账结论"] = "已验证可检索"
        errors, _ = reg.validate(payload["records"], CODES)
        self.assertTrue(any("没有任何成功访问痕迹" in e for e in errors))

    def test_limited_but_recorded_is_acceptable(self) -> None:
        """访问受限但明确记录了尝试，结论为“已验证但访问受限”应可接受。"""
        payload = registry_payload()
        errors, _ = reg.validate(payload["records"], CODES)
        self.assertFalse(any("日本" in e for e in errors))

    def test_limited_conclusion_but_both_success_rejected(self) -> None:
        payload = registry_payload()
        rec = payload["records"][1]
        rec["直接访问状态"] = "HTTP 200"
        rec["浏览器回退状态"] = "不适用：直接访问成功"
        rec["台账结论"] = "已验证但访问受限"
        errors, _ = reg.validate(payload["records"], CODES)
        self.assertTrue(any("均为成功" in e for e in errors))


class TestCLI(TempCase):
    def _run(self) -> int:
        return run_script(
            "validate_authority_registry.py",
            str(self.env.registry_json),
            "--list-file",
            str(self.env.list_file),
        ).returncode

    def test_cli_valid_passes(self) -> None:
        write_registry(self.env)
        self.assertEqual(self._run(), 0)

    def test_cli_bad_payload_fails(self) -> None:
        payload = registry_payload()
        payload["records"][0]["台账结论"] = "已验证可检索"
        payload["records"][0]["直接访问状态"] = "连接超时"
        payload["records"][0]["浏览器回退状态"] = "浏览器加载失败：未知"
        write_registry(self.env, payload)
        self.assertEqual(self._run(), 1)


if __name__ == "__main__":
    unittest.main()
