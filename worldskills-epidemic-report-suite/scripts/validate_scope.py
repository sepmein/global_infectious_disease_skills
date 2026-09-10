#!/usr/bin/env python3
"""Validate dynamic participant and priority-disease configuration files."""

from __future__ import annotations

import argparse
import csv
import json
import re
import sys
from pathlib import Path


sys.stdout.reconfigure(encoding="utf-8")
sys.stderr.reconfigure(encoding="utf-8")


SKILL_DIR = Path(__file__).resolve().parents[1]
DEFAULT_PARTICIPANTS = SKILL_DIR / "references" / "participants.csv"
DEFAULT_DISEASES = SKILL_DIR / "references" / "priority-diseases.xlsx"
PARTICIPANT_COLUMNS = {"地区（代码 - 英文名称）", "中文标准名称"}
DISEASE_COLUMN = "重点病种"
CATEGORY_COLUMN = "所属分类"
ALLOWED_CATEGORIES = {
    "蚊媒及其他虫媒传染病",
    "呼吸道传染病",
    "肠道及食源性传染病",
    "新发少见及高致病性传染病",
}
CODE_PATTERN = re.compile(r"^[A-Z0-9]{2,3}$")


def read_rows(path: Path) -> list[dict[str, str]]:
    suffix = path.suffix.lower()
    if suffix == ".csv":
        with path.open("r", encoding="utf-8-sig", newline="") as handle:
            return [
                {str(key or "").strip(): str(value or "").strip() for key, value in row.items()}
                for row in csv.DictReader(handle)
            ]
    if suffix == ".xlsx":
        from openpyxl import load_workbook

        sheet = load_workbook(path, read_only=True, data_only=True).active
        values = list(sheet.values)
        if not values:
            return []
        headers = [str(value or "").strip() for value in values[0]]
        return [
            dict(zip(headers, [str(value or "").strip() for value in row]))
            for row in values[1:]
        ]
    raise ValueError("仅支持 CSV 或 XLSX")


def validate_participants(path: Path) -> list[dict[str, str]]:
    rows = read_rows(path)
    headers = set(rows[0]) if rows else set()
    missing = PARTICIPANT_COLUMNS - headers
    if missing:
        raise ValueError(f"名单缺少字段：{'、'.join(sorted(missing))}")

    countries: list[dict[str, str]] = []
    seen_codes: set[str] = set()
    seen_names: set[tuple[str, str]] = set()
    errors: list[str] = []
    for row_number, row in enumerate(rows, start=2):
        combined = row.get("地区（代码 - 英文名称）", "").strip()
        chinese = row.get("中文标准名称", "").strip()
        if not combined and not chinese:
            continue
        parts = [part.strip() for part in combined.split(" - ", 1)]
        if len(parts) != 2 or not all(parts):
            errors.append(f"第 {row_number} 行第一列格式应为“代码 - 英文名称”")
            continue
        code, english = parts
        if not CODE_PATTERN.fullmatch(code):
            errors.append(f"第 {row_number} 行地区代码格式无效：{code}")
        if not chinese:
            errors.append(f"第 {row_number} 行缺少中文标准名称")
        if code in seen_codes:
            errors.append(f"第 {row_number} 行地区代码重复：{code}")
        if (chinese, english) in seen_names:
            errors.append(f"第 {row_number} 行名称重复：{chinese} / {english}")
        seen_codes.add(code)
        seen_names.add((chinese, english))
        countries.append(
            {"country_code": code, "country_name_en": english, "country_name_zh": chinese}
        )
    if not countries:
        errors.append("名单中没有有效国家或地区")
    if errors:
        raise ValueError("\n".join(errors))
    return countries


def validate_diseases(path: Path) -> list[dict[str, str]]:
    rows = read_rows(path)
    required = {DISEASE_COLUMN, CATEGORY_COLUMN}
    missing = required - (set(rows[0]) if rows else set())
    if missing:
        raise ValueError(f"重点病种配置缺少字段：{'、'.join(sorted(missing))}")
    diseases = [
        {
            "disease": row.get(DISEASE_COLUMN, "").strip(),
            "category": row.get(CATEGORY_COLUMN, "").strip(),
        }
        for row in rows
    ]
    errors: list[str] = []
    names = [row["disease"] for row in diseases]
    if not diseases or any(not row["disease"] or not row["category"] for row in diseases):
        errors.append("重点病种配置为空，或重点病种/所属分类包含空值")
    if len(names) != len(set(names)):
        errors.append("重点病种配置包含重复名称")
    invalid_categories = sorted({row["category"] for row in diseases if row["category"] not in ALLOWED_CATEGORIES})
    if invalid_categories:
        errors.append(f"重点病种配置包含模板外分类：{'、'.join(invalid_categories)}")
    if errors:
        raise ValueError("\n".join(errors))
    return diseases


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--participants", type=Path, default=DEFAULT_PARTICIPANTS)
    parser.add_argument("--diseases", type=Path, default=DEFAULT_DISEASES)
    args = parser.parse_args()
    try:
        countries = validate_participants(args.participants)
        diseases = validate_diseases(args.diseases)
    except (OSError, ValueError) as exc:
        print(str(exc), file=sys.stderr)
        return 1
    print(
        json.dumps(
            {
                "participants_file": str(args.participants.resolve()),
                "participant_count": len(countries),
                "countries": countries,
                "diseases_file": str(args.diseases.resolve()),
                "priority_disease_count": len(diseases),
                "priority_diseases": [row["disease"] for row in diseases],
                "disease_categories": {
                    row["disease"]: row["category"] for row in diseases
                },
            },
            ensure_ascii=False,
            indent=2,
        )
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
