"""Validate and print enabled priority-disease configuration from CSV or XLSX."""

from __future__ import annotations

import argparse
import csv
import json
import sys
from pathlib import Path


REQUIRED = {"重点病种"}


def read_rows(path: Path) -> list[dict[str, str]]:
    if path.suffix.lower() == ".csv":
        with path.open("r", encoding="utf-8-sig", newline="") as handle:
            return list(csv.DictReader(handle))
    if path.suffix.lower() == ".xlsx":
        from openpyxl import load_workbook
        sheet = load_workbook(path, read_only=True, data_only=True).active
        rows = list(sheet.values)
        headers = [str(value or "").strip() for value in rows[0]]
        return [dict(zip(headers, [str(value or "").strip() for value in row])) for row in rows[1:]]
    raise ValueError("配置表仅支持 .csv 或 .xlsx")


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("config", nargs="?", type=Path, default=Path(__file__).resolve().parents[1] / "references" / "重点病种检索配置.csv")
    args = parser.parse_args()
    try:
        rows = read_rows(args.config)
    except Exception as exc:
        print(f"无法读取重点病种配置：{exc}", file=sys.stderr)
        return 1
    if not rows or not REQUIRED.issubset(rows[0]):
        print("重点病种配置缺少必需字段", file=sys.stderr)
        return 1
    names = [row.get("重点病种", "").strip() for row in rows]
    if not names or any(not name for name in names) or len(names) != len(set(names)):
        print("重点病种配置为空、存在空病种名称或重复病种名称", file=sys.stderr)
        return 1
    print(json.dumps({"config_file": str(args.config), "disease_count": len(names), "diseases": names}, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
