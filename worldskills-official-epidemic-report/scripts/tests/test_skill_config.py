"""共享配置模块：双格式读取、多工作表精准选表、分类与去重校验。"""

from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from _helpers import (
    write_disease_csv,
    write_disease_xlsx,
    write_participant_csv,
    write_participant_xlsx,
    write_xlsx,
)

import skill_config as sc

COUNTRIES = [("TH", "Thailand", "泰国"), ("JP", "Japan", "日本"), ("BR", "Brazil", "巴西")]
DISEASES = [("登革热", "蚊媒及其他虫媒传染病"), ("麻疹", "呼吸道传染病"), ("霍乱", "肠道及食源性传染病")]


class TempCase(unittest.TestCase):
    def setUp(self) -> None:
        self._temp = tempfile.TemporaryDirectory()
        self.root = Path(self._temp.name)
        self.addCleanup(self._temp.cleanup)


class TestParticipantList(TempCase):
    def test_real_ooxml_with_csv_suffix_selects_correct_sheet(self) -> None:
        """名单后缀为 .csv 但实际是 xlsx，且含易误读的“名单差异汇总”表。"""
        path = write_participant_xlsx(self.root / "参赛国家和地区名单.csv", COUNTRIES)
        participants = sc.load_participants(path)
        self.assertEqual(participants.sheet_name, "国家地区合集")
        self.assertEqual(participants.count, 3)
        self.assertEqual(participants.codes, {"TH", "JP", "BR"})
        self.assertNotIn("XX", participants.codes)  # 差异汇总表中的代码不得混入

    def test_plain_csv_three_columns(self) -> None:
        path = write_participant_csv(self.root / "list.csv", COUNTRIES)
        participants = sc.load_participants(path)
        self.assertEqual(participants.schema_name, "三列格式")
        self.assertEqual(participants.count, 3)

    def test_legacy_merged_two_columns(self) -> None:
        path = write_participant_csv(self.root / "legacy.csv", COUNTRIES, merged=True)
        participants = sc.load_participants(path)
        self.assertEqual(participants.schema_name, "旧合并两列格式")
        self.assertEqual(participants.by_code["TH"].name_en, "Thailand")

    def test_duplicate_code_rejected(self) -> None:
        path = write_participant_csv(self.root / "dup.csv", COUNTRIES + [("TH", "Thailand", "泰国")])
        with self.assertRaises(sc.ConfigError) as ctx:
            sc.load_participants(path)
        self.assertIn("重复", str(ctx.exception))

    def test_empty_chinese_name_rejected(self) -> None:
        path = write_participant_csv(self.root / "empty.csv", COUNTRIES + [("XX", "Nowhere", "")])
        with self.assertRaises(sc.ConfigError) as ctx:
            sc.load_participants(path)
        self.assertIn("中文名称", str(ctx.exception))

    def test_wrong_header_rejected(self) -> None:
        path = write_xlsx(self.root / "bad.xlsx", {"Sheet1": [["国家", "名称"], ["TH", "泰国"]]})
        with self.assertRaises(sc.ConfigError):
            sc.load_participants(path)

    def test_ambiguous_sheets_require_explicit_choice(self) -> None:
        rows = [["代码", "英文名称", "中文名称"], ["TH", "Thailand", "泰国"]]
        path = write_xlsx(self.root / "two.xlsx", {"名单A": rows, "名单B": rows})
        with self.assertRaises(sc.ConfigError) as ctx:
            sc.load_participants(path)
        self.assertIn("--sheet", str(ctx.exception))
        participants = sc.load_participants(path, sheet="名单B")
        self.assertEqual(participants.sheet_name, "名单B")


class TestDiseaseConfig(TempCase):
    def test_xlsx_and_csv_agree(self) -> None:
        xlsx = write_disease_xlsx(self.root / "d.xlsx", DISEASES)
        csv_path = write_disease_csv(self.root / "d.csv", DISEASES)
        from_xlsx = sc.load_diseases(xlsx)
        from_csv = sc.load_diseases(csv_path)
        self.assertEqual(from_xlsx.names, from_csv.names)
        self.assertEqual(from_xlsx.category_by_disease["麻疹"], "呼吸道传染病")
        self.assertEqual(from_xlsx.by_category["新发少见及高致病性传染病"], ())

    def test_category_must_come_from_fixed_four(self) -> None:
        path = write_disease_csv(self.root / "bad.csv", DISEASES + [("狂犬病", "其他传染病")])
        with self.assertRaises(sc.ConfigError) as ctx:
            sc.load_diseases(path)
        self.assertIn("不在固定四类", str(ctx.exception))

    def test_missing_category_column_rejected(self) -> None:
        path = self.root / "single.csv"
        path.write_text("重点病种\n登革热\n麻疹\n", encoding="utf-8")
        with self.assertRaises(sc.ConfigError) as ctx:
            sc.load_diseases(path)
        self.assertIn("所属分类", str(ctx.exception))

    def test_empty_category_rejected(self) -> None:
        path = write_disease_csv(self.root / "emptycat.csv", DISEASES + [("水痘", "")])
        with self.assertRaises(sc.ConfigError) as ctx:
            sc.load_diseases(path)
        self.assertIn("缺少所属分类", str(ctx.exception))

    def test_duplicate_disease_rejected(self) -> None:
        path = write_disease_csv(self.root / "dup.csv", DISEASES + [("麻疹", "呼吸道传染病")])
        with self.assertRaises(sc.ConfigError) as ctx:
            sc.load_diseases(path)
        self.assertIn("重复", str(ctx.exception))


class TestIntroIndex(TempCase):
    def setUp(self) -> None:
        super().setUp()
        self.diseases = sc.load_diseases(write_disease_csv(self.root / "d.csv", DISEASES))
        self.intro_dir = self.root / "疾病简介"
        self.intro_dir.mkdir()

    def _index(self, payload: dict) -> Path:
        import json

        path = self.root / "疾病简介索引.json"
        path.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")
        return path

    def test_null_entry_reported_as_missing(self) -> None:
        path = self._index({"登革热": "登革热.pdf", "麻疹": "麻疹.pdf", "霍乱": None})
        index = sc.load_intro_index(path, intro_dir=self.intro_dir, diseases=self.diseases)
        self.assertEqual(index.missing_diseases, ("霍乱",))

    def test_key_missing_auto_maps_none_with_warning(self) -> None:
        """新增病种键缺失：自动映射为缺失（None）并告警，不再阻断。"""
        path = self._index({"登革热": "登革热.pdf"})
        index = sc.load_intro_index(path, intro_dir=self.intro_dir, diseases=self.diseases)
        self.assertEqual(index.mapping.get("麻疹"), None)
        self.assertEqual(index.mapping.get("霍乱"), None)
        self.assertTrue(any("自动映射" in warning for warning in index.warnings))

    def test_extra_key_out_of_scope_rejected(self) -> None:
        """配置外病种（额外键）仍拒绝：不得扩大范围。"""
        path = self._index(
            {"登革热": "a.pdf", "麻疹": "b.pdf", "霍乱": None, "鼠疫": "c.pdf"}
        )
        with self.assertRaises(sc.ConfigError) as ctx:
            sc.load_intro_index(path, intro_dir=self.intro_dir, diseases=self.diseases)
        self.assertIn("配置外病种", str(ctx.exception))

    def test_out_of_scope_disease_rejected(self) -> None:
        path = self._index(
            {"登革热": "a.pdf", "麻疹": "b.pdf", "霍乱": None, "鼠疫": "c.pdf"}
        )
        with self.assertRaises(sc.ConfigError) as ctx:
            sc.load_intro_index(path, intro_dir=self.intro_dir, diseases=self.diseases)
        self.assertIn("配置外病种", str(ctx.exception))

    def test_missing_index_file(self) -> None:
        with self.assertRaises(sc.ConfigError) as ctx:
            sc.load_intro_index(self.root / "缺失.json", intro_dir=self.intro_dir, diseases=self.diseases)
        self.assertIn("疾病简介索引", str(ctx.exception))


if __name__ == "__main__":
    unittest.main()
