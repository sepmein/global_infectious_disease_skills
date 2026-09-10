"""P4 交付件校验：Word 结构/格式、占位残留、Excel 底表/证据清单、逐国覆盖、交叉校验。"""

from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

import skill_config as sc
from docx_reader import read_docx

import validate_deliverables as vd

from _fixtures import build_env, report_sections, write_report, write_workbook
from _helpers import (
    paragraph,
    run_script,
    valid_report_docx,
    write_docx,
    write_xlsx,
)

COUNTRIES = [("TH", "Thailand", "泰国"), ("JP", "Japan", "日本")]


class TempCase(unittest.TestCase):
    def setUp(self) -> None:
        self._temp = tempfile.TemporaryDirectory()
        self.root = Path(self._temp.name)
        self.addCleanup(self._temp.cleanup)
        self.env = build_env(self.root / "env")
        self.participants = sc.load_participants(self.env.list_file)
        self.diseases = sc.load_diseases(self.env.disease_config)


class TestWordStructurePositive(TempCase):
    def test_valid_report_passes_structure(self) -> None:
        path = write_report(self.env)
        doc = read_docx(path)
        errors = vd.validate_word_structure(doc, self.diseases, self.participants.count)
        self.assertEqual(errors, [])


class TestWordPlaceholders(TempCase):
    def test_placeholder_fragments_rejected(self) -> None:
        path = write_docx(
            self.root / "bad.docx",
            [
                paragraph("标题", font="方正小标宋简体", size="44", title=True),
                paragraph("示例数据 12,500，所有数字和情境均为虚构"),
            ],
        )
        doc = read_docx(path)
        errors = vd.validate_word_placeholders(doc)
        self.assertTrue(errors)

    def test_bare_fill_word_allowed(self) -> None:
        # 用户规范风险语境下的一般“填写”提法（如填报必要性说明）不得被误拒；
        # 仅明确的模板占位/写作指导句（如“本部分回答”）才拦截。
        path = write_docx(
            self.root / "fill.docx",
            [paragraph("标题", font="方正小标宋简体", size="44", title=True), paragraph("请按监测要求如实填写病例数据")],
        )
        doc = read_docx(path)
        errors = vd.validate_word_placeholders(doc)
        self.assertEqual(errors, [])

    def test_writing_guidance_rejected(self) -> None:
        # 明确的模板写作指导句“本部分回答”仍须被拒绝；单独的“填写”一词不再拦截。
        path = write_docx(
            self.root / "guide.docx",
            [paragraph("标题", font="方正小标宋简体", size="44", title=True), paragraph("填写本部分回答")],
        )
        doc = read_docx(path)
        errors = vd.validate_word_placeholders(doc)
        self.assertTrue(any("本部分回答" in e for e in errors))


class TestWordPublicSources(TempCase):
    def _errors(self, text: str) -> list[str]:
        path = write_docx(
            self.root / "source.docx",
            [paragraph("标题", font="方正小标宋简体", size="44", title=True), paragraph(text)],
        )
        return vd.validate_word_public_sources(read_docx(path))

    def test_http_source_allowed(self) -> None:
        self.assertEqual(self._errors("信息来源：https://www.who.int/example，访问日期2026-09-09。"), [])

    def test_windows_local_path_rejected(self) -> None:
        errors = self._errors(r"信息来源：C:\data\weekly_report.pdf，第3页。")
        self.assertTrue(any("本地文件来源" in item for item in errors))

    def test_local_prefix_rejected_even_with_web_url(self) -> None:
        errors = self._errors("信息来源：https://example.org/report；LOCAL:LE1")
        self.assertTrue(any("本地文件来源" in item for item in errors))

    def test_local_filename_rejected_even_with_web_url(self) -> None:
        errors = self._errors("信息来源：https://example.org/report；附件 weekly_report.xlsx 第2表。")
        self.assertTrue(any("本地文件来源" in item for item in errors))

    def test_source_line_without_http_rejected(self) -> None:
        errors = self._errors("信息来源：世界卫生组织疾病主题页。")
        self.assertTrue(any("只允许HTTP(S)网络链接" in item for item in errors))

    def test_non_source_body_local_path_rejected(self) -> None:
        errors = self._errors(r"病例数据摘自 D:\reports\outbreak.csv。")
        self.assertTrue(any("本地文件来源" in item for item in errors))

    def test_local_intro_wording_rejected(self) -> None:
        errors = self._errors("信息来源：https://example.org/disease；本地简介第2页。")
        self.assertTrue(any("本地文件来源" in item for item in errors))


class TestWordFormat(TempCase):
    def test_wrong_font_rejected(self) -> None:
        path = write_docx(
            self.root / "fmt.docx",
            [
                paragraph("标题", font="方正小标宋简体", size="44", title=True),
                paragraph("正文应使用仿宋", font="宋体", size="36"),
            ],
        )
        doc = read_docx(path)
        errors = vd.validate_word_format(doc, self.diseases)
        self.assertTrue(any("字体" in e for e in errors))


class TestWorkbook(TempCase):
    def test_valid_workbook(self) -> None:
        path = write_workbook(self.env)
        errors, events = vd.validate_workbook(path, self.participants, self.diseases)
        self.assertEqual(errors, [])
        self.assertEqual(events, {("TH", "登革热"), ("JP", "麻疹")})

    def test_missing_sheet_rejected(self) -> None:
        path = write_xlsx(
            self.root / "partial.xlsx",
            {"疫情摸底底表": [list(vd.MAIN_HEADERS)]},
        )
        errors, _ = vd.validate_workbook(path, self.participants, self.diseases)
        self.assertTrue(any("缺少工作表" in e for e in errors))

    def test_unconfigured_disease_rejected(self) -> None:
        path = write_workbook(
            self.env,
            main_rows=[
                ["TH - Thailand 泰国", "鼠疫", "异常上升", "新发", "5", "0", "2026-08-31", "https://x.org", "高"]
            ],
        )
        errors, _ = vd.validate_workbook(path, self.participants, self.diseases)
        self.assertTrue(any("配置外病种" in e for e in errors))

    def test_local_evidence_file_missing_rejected(self) -> None:
        path = write_workbook(
            self.env,
            local_rows=[["LE9", "麻疹", "JP - Japan 日本", str(self.root / "无此文件.txt"), "第1页", "2026-08-31", "人工核实一致"]],
        )
        errors, _ = vd.validate_workbook(path, self.participants, self.diseases)
        self.assertTrue(any("文件不存在" in e for e in errors))

    def test_country_log_under_coverage_rejected(self) -> None:
        path = write_workbook(self.env, log_codes=["TH"])
        errors, _ = vd.validate_workbook(path, self.participants, self.diseases)
        self.assertTrue(any("逐国检索记录缺少" in e for e in errors))


class TestCrossCheck(TempCase):
    def setUp(self) -> None:
        super().setUp()
        from _fixtures import write_research

        write_research(self.env)

    def test_cross_check_consistent(self) -> None:
        errors = vd.cross_check_research(self.env.research_json, {("TH", "登革热"), ("JP", "麻疹")})
        self.assertEqual(errors, [])

    def test_cross_check_included_not_in_events(self) -> None:
        # 检索总账标记纳入，但底表无此事件
        errors = vd.cross_check_research(self.env.research_json, {("TH", "登革热")})
        self.assertTrue(any("底表缺少对应事件" in e for e in errors))


class TestCLI(TempCase):
    def test_cli_valid_passes(self) -> None:
        from _fixtures import write_research

        write_research(self.env)
        write_workbook(self.env)
        write_report(self.env)
        code = run_script(
            "validate_deliverables.py",
            "--xlsx",
            str(self.env.xlsx),
            "--docx",
            str(self.env.docx),
            "--research",
            str(self.env.research_json),
            "--list-file",
            str(self.env.list_file),
            "--disease-config",
            str(self.env.disease_config),
        ).returncode
        self.assertEqual(code, 0)

    def test_cli_placeholder_fails(self) -> None:
        # 用含占位的 docx 替换报告，结构/格式会失败，但至少占位检查应触发
        write_workbook(self.env)
        bad = write_docx(
            self.root / "bad.docx",
            [paragraph("标题", font="方正小标宋简体", size="44", title=True), paragraph("示例 12,500")],
        )
        result = run_script(
            "validate_deliverables.py",
            "--xlsx",
            str(self.env.xlsx),
            "--docx",
            str(bad),
            "--list-file",
            str(self.env.list_file),
            "--disease-config",
            str(self.env.disease_config),
        )
        self.assertEqual(result.returncode, 1)
        self.assertIn("占位", result.stderr)

    def test_cli_rejects_word_local_source_leak(self) -> None:
        from _fixtures import write_research

        write_research(self.env)
        write_workbook(self.env)
        sections = report_sections()
        sections[sc.CATEGORY_HEADINGS[0]][3] = paragraph(
            r"信息来源：https://example.org/global/dengue-2026；C:\private\dengue.pdf，第2页。"
        )
        write_report(self.env, sections=sections)
        result = run_script(
            "validate_deliverables.py",
            "--xlsx",
            str(self.env.xlsx),
            "--docx",
            str(self.env.docx),
            "--research",
            str(self.env.research_json),
            "--list-file",
            str(self.env.list_file),
            "--disease-config",
            str(self.env.disease_config),
        )
        self.assertEqual(result.returncode, 1)
        self.assertIn("本地文件来源", result.stderr)

    def test_cli_rejects_unregistered_word_url(self) -> None:
        from _fixtures import write_research

        write_research(self.env)
        write_workbook(self.env)
        sections = report_sections()
        sections[sc.CATEGORY_HEADINGS[0]][3] = paragraph(
            "信息来源：https://unregistered.example.org/dengue，访问日期2026-09-05。"
        )
        write_report(self.env, sections=sections)
        result = run_script(
            "validate_deliverables.py",
            "--xlsx",
            str(self.env.xlsx),
            "--docx",
            str(self.env.docx),
            "--research",
            str(self.env.research_json),
            "--list-file",
            str(self.env.list_file),
            "--disease-config",
            str(self.env.disease_config),
        )
        self.assertEqual(result.returncode, 1)
        self.assertIn("已核实web链接", result.stderr)


if __name__ == "__main__":
    unittest.main()
