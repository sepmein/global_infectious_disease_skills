"""校验世界技能大赛参赛国家和地区名单。"""

from __future__ import annotations

import csv
import json
import sys
from pathlib import Path


SKILL_DIR = Path(__file__).resolve().parents[1]
LIST_FILE = SKILL_DIR / "references" / "参赛国家和地区名单.csv"
CODE_AND_ENGLISH_COLUMN = "地区（代码 - 英文名称）"
CHINESE_NAME_COLUMN = "中文标准名称"
REQUIRED_COLUMNS = {CODE_AND_ENGLISH_COLUMN, CHINESE_NAME_COLUMN}


def main() -> int:
    if not LIST_FILE.exists():
        print(f"未找到名单文件：{LIST_FILE}", file=sys.stderr)
        return 1

    with LIST_FILE.open("r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle)
        missing = REQUIRED_COLUMNS - set(reader.fieldnames or [])
        if missing:
            print(f"名单缺少字段：{', '.join(sorted(missing))}", file=sys.stderr)
            return 1
        rows = list(reader)

    participant_rows = [
        row
        for row in rows
        if str(row.get(CODE_AND_ENGLISH_COLUMN, "")).strip()
        or str(row.get(CHINESE_NAME_COLUMN, "")).strip()
    ]
    if not participant_rows:
        print("名单中没有可用的国家或地区，请更新 references/参赛国家和地区名单.csv", file=sys.stderr)
        return 1

    errors: list[str] = []
    countries: list[dict[str, str]] = []
    seen_codes: set[str] = set()
    seen_names: set[tuple[str, str]] = set()
    for index, row in enumerate(participant_rows, start=2):
        code_and_english = str(row.get(CODE_AND_ENGLISH_COLUMN, "")).strip()
        country_zh = str(row.get(CHINESE_NAME_COLUMN, "")).strip()
        parts = [part.strip() for part in code_and_english.split(" - ", maxsplit=1)]
        if len(parts) != 2 or not all(parts):
            errors.append(f"第 {index} 行第一列格式应为“代码 - 英文名称”：{code_and_english}")
            continue

        country_code, country_en = parts
        if not country_zh:
            errors.append(f"第 {index} 行缺少中文标准名称")
        if country_code in seen_codes:
            errors.append(f"第 {index} 行地区代码重复：{country_code}")
        seen_codes.add(country_code)
        name_key = (country_zh, country_en)
        if name_key in seen_names:
            errors.append(f"第 {index} 行国家或地区名称重复：{country_zh} / {country_en}")
        seen_names.add(name_key)
        countries.append(
            {
                "country_code": country_code,
                "country_name_en": country_en,
                "country_name_zh": country_zh,
            }
        )

    if errors:
        print("\n".join(errors), file=sys.stderr)
        return 1

    print(
        json.dumps(
            {
                "list_file": str(LIST_FILE),
                "participant_country_count": len(countries),
                "countries": countries,
            },
            ensure_ascii=False,
            indent=2,
        )
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
