"""P2 检索与证据总账（contract 2.x）：国家×病种全覆盖、本地痕迹、未读简介、本地不足需网络、表外拒绝。"""

from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from _fixtures import (
    CATEGORY_HEADINGS,
    COUNTRIES,
    MEASLES_URL,
    build_complete_env,
    build_env,
    research_payload,
    research_payload_complete,
    write_research,
)
from _helpers import run_script
import json


class TempCase(unittest.TestCase):
    def setUp(self) -> None:
        self._temp = tempfile.TemporaryDirectory()
        self.root = Path(self._temp.name)
        self.addCleanup(self._temp.cleanup)
        self.env = build_env(self.root / "env")

    def _write(self, payload: dict) -> None:
        self.env.research_json.write_text(
            json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8"
        )

    def _args(self, require_complete: bool = False) -> list[str]:
        args = [
            str(self.env.research_json),
            "--list-file",
            str(self.env.list_file),
            "--disease-config",
            str(self.env.disease_config),
            "--intro-index",
            str(self.env.intro_index),
            "--intro-dir",
            str(self.env.intro_dir),
        ]
        if require_complete:
            args.append("--require-complete")
        return args

    def _run(self, mutate=None, require_complete: bool = False, *, payload=None) -> int:
        if payload is None:
            payload = research_payload(self.env)
        if mutate is not None:
            mutate(payload)
        self._write(payload)
        return run_script("validate_country_research.py", *self._args(require_complete)).returncode

    def _stderr(self, mutate=None, require_complete: bool = False, *, payload=None) -> str:
        if payload is None:
            payload = research_payload(self.env)
        if mutate is not None:
            mutate(payload)
        self._write(payload)
        return run_script("validate_country_research.py", *self._args(require_complete)).stderr


class TestPositive(TempCase):
    def setUp(self) -> None:
        super().setUp()
        self.env = build_complete_env(self.root / "complete")

    def test_valid_complete_payload_passes(self) -> None:
        self.assertEqual(
            self._run(payload=research_payload_complete(self.env), require_complete=True), 0
        )


class TestCoverageAndOutOfScope(TempCase):
    def test_country_disease_coverage_missing(self) -> None:
        def mutate(payload: dict) -> None:
            payload["records"][0]["disease_checks"].pop()  # 少一个病种
        self.assertEqual(self._run(mutate), 1)
        self.assertIn("缺少", self._stderr(mutate))

    def test_out_of_scope_disease_rejected(self) -> None:
        def mutate(payload: dict) -> None:
            payload["records"][0]["disease_checks"].append(
                {
                    "disease": "鼠疫",
                    "category": "新发少见及高致病性传染病",
                    "status": "已完成未纳入",
                    "reason": "表外病种不应出现",
                    "evidence_ids": [],
                    "internet_required": True,
                    "internet_searches": [],
                }
            )
        self.assertEqual(self._run(mutate), 1)
        self.assertIn("配置内", self._stderr(mutate))

    def test_out_of_scope_country_rejected(self) -> None:
        def mutate(payload: dict) -> None:
            payload["records"].append(
                {
                    "country_code": "ZZ",
                    "country_name_zh": "虚构",
                    "research_status": "已完成",
                    "disease_checks": [
                        {
                            "disease": "登革热",
                            "category": "蚊媒及其他虫媒传染病",
                            "status": "已完成未纳入",
                            "reason": "x",
                            "evidence_ids": [],
                            "internet_required": True,
                            "internet_searches": [],
                        }
                    ],
                }
            )
        self.assertEqual(self._run(mutate), 1)
        self.assertIn("动态名单内", self._stderr(mutate))


class TestLocalInputs(TempCase):
    def test_missing_local_path_blocks(self) -> None:
        def mutate(payload: dict) -> None:
            payload["local_inputs"]["paths"].append(str(self.root / "不存在的目录"))
        self.assertEqual(self._run(mutate), 1)
        self.assertIn("阻断", self._stderr(mutate))

    def test_local_only_but_no_verified_local_evidence(self) -> None:
        def mutate(payload: dict) -> None:
            jp = payload["records"][1]
            check = next(c for c in jp["disease_checks"] if c["disease"] == "登革热")
            check["evidence_ids"] = []  # 去掉本地证据，但仍声明本地可完成
            check["internet_required"] = False
        self.assertEqual(self._run(mutate), 1)
        self.assertIn("本地证据", self._stderr(mutate))

    def test_local_only_with_fake_verified_status(self) -> None:
        def mutate(payload: dict) -> None:
            # 本地证据未声明人工核实，却想以本地完成
            payload["sources"][1]["verification"] = "未人工核实"
        self.assertEqual(self._run(mutate), 1)


class TestIntroAndSearches(TempCase):
    def test_unread_available_intro_rejected(self) -> None:
        def mutate(payload: dict) -> None:
            section = next(s for s in payload["disease_sections"] if s["disease"] == "登革热")
            section["intro_read"] = False  # 简介文件存在却未读取
        self.assertEqual(self._run(mutate), 1)
        self.assertIn("intro_read", self._stderr(mutate))

    def test_missing_intro_must_use_web(self) -> None:
        def mutate(payload: dict) -> None:
            # 麻疹简介本就缺失，若伪造本地充分且无网络检索应被拒
            section = next(s for s in payload["disease_sections"] if s["disease"] == "麻疹")
            section["sections"]["overview"]["local_sufficient"] = True
            section["sections"]["overview"]["internet_searches"] = []
            section["sections"]["overview"]["evidence_ids"] = []
        self.assertEqual(self._run(mutate), 1)
        self.assertIn("网络检索", self._stderr(mutate))

    def test_fabricated_intro_file_for_null_rejected(self) -> None:
        def mutate(payload: dict) -> None:
            section = next(s for s in payload["disease_sections"] if s["disease"] == "麻疹")
            section["intro_file"] = "麻疹简介.pdf"  # 索引为 null，不得伪造
            section["intro_read"] = True
        self.assertEqual(self._run(mutate), 1)
        self.assertIn("伪造", self._stderr(mutate))


class TestNoEpidemicClaim(TempCase):
    def test_unsearched_must_not_claim_no_epidemic(self) -> None:
        def mutate(payload: dict) -> None:
            check = payload["records"][0]["disease_checks"][0]
            check["status"] = "检索未完成"
            check["reason"] = "尚未检索，初步判断无疫情"  # 未检索却给无疫情结论
        self.assertEqual(self._run(mutate), 1)
        self.assertIn("无疫情", self._stderr(mutate))

    def test_completed_but_no_evidence_rejected(self) -> None:
        def mutate(payload: dict) -> None:
            check = payload["records"][0]["disease_checks"][0]
            check["status"] = "已完成并纳入"
            check["evidence_ids"] = []
        self.assertEqual(self._run(mutate), 1)
        self.assertIn("没有任何已读取证据", self._stderr(mutate))


class TestRequireCompleteNullIntro(TempCase):
    def test_null_intro_web_fallback_passes_require_complete(self) -> None:
        """简介缺失（索引 null）在 require_complete 时，若有同病种已读取已核实的 official 网络来源 + 回退原因，则可通过。"""
        self.assertEqual(self._run(require_complete=True), 0)

    def test_intro_missing_file_web_fallback_passes(self) -> None:
        """索引指向的文件不存在：同样允许权威网络回退，require_complete 可通过。"""
        # 把索引中麻疹改为指向不存在的文件
        index = json.loads(self.env.intro_index.read_text(encoding="utf-8"))
        index["麻疹"] = "不存在的麻疹简介.pdf"
        self.env.intro_index.write_text(json.dumps(index, ensure_ascii=False), encoding="utf-8")
        self.assertEqual(self._run(require_complete=True), 0)


class TestIntroFallbackRejections(TempCase):
    def test_null_intro_fallback_no_search_rejected(self) -> None:
        """缺网络检索（无 internet_searches）必须拒。"""
        def mutate(payload: dict) -> None:
            section = next(s for s in payload["disease_sections"] if s["disease"] == "麻疹")
            section["sections"]["overview"]["internet_searches"] = []
        self.assertEqual(self._run(mutate, require_complete=True), 1)
        self.assertIn("internet_searches", self._stderr(mutate, require_complete=True))

    def test_null_intro_fallback_unverified_rejected(self) -> None:
        """网络来源未核实（verification 不在已核实集合）必须拒。"""
        def mutate(payload: dict) -> None:
            for source in payload["sources"]:
                if source["id"] == "W3":
                    source["verification"] = "未人工核实"
        self.assertEqual(self._run(mutate, require_complete=True), 1)
        self.assertIn("authority_type='official'", self._stderr(mutate, require_complete=True))

    def test_null_intro_fallback_non_official_rejected(self) -> None:
        """仅媒体/非 official 来源必须拒。"""
        def mutate(payload: dict) -> None:
            for source in payload["sources"]:
                if source["id"] == "W3":
                    source["authority_type"] = "media"
                    source["authority_basis"] = "某媒体报道"
        self.assertEqual(self._run(mutate, require_complete=True), 1)
        self.assertIn("authority_type='official'", self._stderr(mutate, require_complete=True))

    def test_null_intro_fallback_wrong_disease_rejected(self) -> None:
        """回退来源病种不同（非同病种）必须拒。"""
        def mutate(payload: dict) -> None:
            for source in payload["sources"]:
                if source["id"] == "W3":
                    source["disease"] = "登革热"
        self.assertEqual(self._run(mutate, require_complete=True), 1)
        self.assertIn("同病种", self._stderr(mutate, require_complete=True))

    def test_null_intro_fallback_missing_reason_rejected(self) -> None:
        """未填写 intro_fallback_reason 必须拒。"""
        def mutate(payload: dict) -> None:
            section = next(s for s in payload["disease_sections"] if s["disease"] == "麻疹")
            section["intro_fallback_reason"] = ""
        self.assertEqual(self._run(mutate, require_complete=True), 1)
        self.assertIn("intro_fallback_reason", self._stderr(mutate, require_complete=True))

    def test_null_intro_fallback_search_without_evidence_rejected(self) -> None:
        """有搜索但结果未对应到 overview.evidence_ids（有搜索无对应引用）必须拒。"""
        def mutate(payload: dict) -> None:
            section = next(s for s in payload["disease_sections"] if s["disease"] == "麻疹")
            section["sections"]["overview"]["evidence_ids"] = []
        self.assertEqual(self._run(mutate, require_complete=True), 1)
        self.assertIn("有交集", self._stderr(mutate, require_complete=True))


class TestVerificationAndSearches(TempCase):
    def test_verification_核实一致_accepted(self) -> None:
        def mutate(payload: dict) -> None:
            for source in payload["sources"]:
                if source["id"] == "L2":
                    source["verification"] = "核实一致"
        self.assertEqual(self._run(mutate), 0)

    def test_sufficient_requires_same_disease_verified(self) -> None:
        def mutate(payload: dict) -> None:
            for source in payload["sources"]:
                if source["id"] == "W2":
                    source["verification"] = "未人工核实"
        self.assertEqual(self._run(mutate), 1)
        self.assertIn("同病种已核实", self._stderr(mutate))

    def test_search_unfinished_not_counted(self) -> None:
        def mutate(payload: dict) -> None:
            check = payload["records"][0]["disease_checks"][0]
            check["internet_searches"] = [
                {"query": "x", "tool": "web", "search_date": "2026-09-05", "result_source_ids": ["W1"], "status": "未完成"}
            ]
        self.assertEqual(self._run(mutate), 1)
        self.assertIn("已完成的 internet_searches", self._stderr(mutate))

    def test_search_local_impersonation_rejected(self) -> None:
        def mutate(payload: dict) -> None:
            check = payload["records"][0]["disease_checks"][0]
            check["internet_searches"] = [
                {"query": "x", "tool": "web", "search_date": "2026-09-05", "result_source_ids": ["L2"], "status": "已完成"}
            ]
        self.assertEqual(self._run(mutate), 1)
        self.assertIn("冒充", self._stderr(mutate))

    def test_local_source_origin_url_allowed(self) -> None:
        def mutate(payload: dict) -> None:
            for source in payload["sources"]:
                if source["id"] == "L1":
                    source["source_origin_url"] = "https://example.org/dengue-intro"
        self.assertEqual(self._run(mutate), 0)


class TestCountryVerifiedEvidence(TempCase):
    def test_global_verified_cannot_satisfy_country(self) -> None:
        """本国本地完成时，不得借用 GLOBAL 已核实来源而本国未核实通过。"""
        def mutate(payload: dict) -> None:
            jp = payload["records"][1]
            check = next(c for c in jp["disease_checks"] if c["disease"] == "登革热")
            check["internet_required"] = False
            check["evidence_ids"] = ["L1"]  # L1 为 GLOBAL 登革热已核实，非 JP
        self.assertEqual(self._run(mutate), 1)
        self.assertIn("同国家", self._stderr(mutate))

    def test_completed_record_missing_disease_still_errors(self) -> None:
        def mutate(payload: dict) -> None:
            payload["records"][0]["disease_checks"].pop()  # 即使 research_status=已完成 也缺病种
            payload["records"][0]["research_status"] = "已完成"
        self.assertEqual(self._run(mutate), 1)
        self.assertIn("缺少", self._stderr(mutate))


if __name__ == "__main__":
    unittest.main()
