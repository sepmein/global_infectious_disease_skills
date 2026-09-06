"""Reject incomplete country research records before report generation."""

from __future__ import annotations

import argparse
import csv
import json
import sys
from pathlib import Path


REQUIRED_PAGE_TYPES = {"monitoring", "outbreak_or_situation", "disease_data"}
ALLOWED_STATUS = {"已完成并纳入", "已完成未纳入", "国家级资料待核实", "检索未完成"}
SKILL_DIR = Path(__file__).resolve().parents[1]
LIST_FILE = SKILL_DIR / "references" / "参赛国家和地区名单.csv"


def participant_codes() -> set[str]:
    with LIST_FILE.open("r", encoding="utf-8-sig", newline="") as handle:
        return {
            row["地区（代码 - 英文名称）"].split(" - ", 1)[0].strip()
            for row in csv.DictReader(handle)
        }


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("research_json", type=Path)
    parser.add_argument(
        "--require-complete",
        action="store_true",
        help="用于正式交付件门禁；拒绝待核实或未完成的逐国记录。",
    )
    args = parser.parse_args()
    records = json.loads(args.research_json.read_text(encoding="utf-8")).get("records", [])
    errors: list[str] = []
    expected_codes = participant_codes()
    seen_codes: set[str] = set()
    for row in records:
        name = row.get("country_name_zh") or row.get("country_zh") or "未命名国家"
        code = str(row.get("country_code", "")).strip()
        if not code:
            errors.append(f"{name}：缺少地区代码")
        elif code in seen_codes:
            errors.append(f"{name}：地区代码重复：{code}")
        else:
            seen_codes.add(code)
        status = row.get("research_status")
        if status not in ALLOWED_STATUS:
            errors.append(f"{name}：缺少有效检索结束状态")
            continue
        pages = row.get("pages", [])
        page_types = {page.get("page_type") for page in pages if page.get("access_status") in {"HTTP 200", "浏览器加载成功"}}
        if status.startswith("已完成") and "monitoring" not in page_types:
            errors.append(f"{name}：标记为完成但未实际访问国家级监测/周报或数据页")
        if status == "已完成并纳入" and not (page_types & {"outbreak_or_situation", "disease_data", "monitoring"}):
            errors.append(f"{name}：已纳入但没有可读的事件或监测证据页")
        if status == "已完成未纳入" and not row.get("not_included_reason"):
            errors.append(f"{name}：已完成未纳入但缺少原因")
        if status == "检索未完成" and row.get("conclusion") in {"无疫情", "零病例"}:
            errors.append(f"{name}：检索未完成却给出无疫情/零病例结论")
        if "regional_only" in page_types and status.startswith("已完成") and "monitoring" not in page_types:
            errors.append(f"{name}：仅有区域来源，不能标记为国家级检索完成")
        if args.require_complete and status in {"国家级资料待核实", "检索未完成"}:
            errors.append(f"{name}：正式交付件前逐国检索尚未完成：{status}")
    if seen_codes != expected_codes:
        errors.append(
            f"逐国检索记录与动态名单不一致：缺少{len(expected_codes - seen_codes)}个，多出{len(seen_codes - expected_codes)}个"
        )
    if errors:
        print("\n".join(errors), file=sys.stderr)
        return 1
    print(f"逐国检索完备性校验通过：{len(records)}个国家/地区。")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
