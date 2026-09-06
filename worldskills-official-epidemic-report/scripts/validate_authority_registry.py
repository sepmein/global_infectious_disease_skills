"""Hard-gate a newly built national-authority registry before epidemic research."""

from __future__ import annotations

import argparse
import csv
import json
import sys
from pathlib import Path


SKILL_DIR = Path(__file__).resolve().parents[1]
LIST_FILE = SKILL_DIR / "references" / "参赛国家和地区名单.csv"
REQUIRED_FIELDS = {
    "地区代码", "中文标准名称", "英文名称", "发现检索词", "候选页面 URL", "机构名称", "机构类型",
    "机构权威性核实依据 URL", "规范入口 URL", "直接访问状态", "直接访问最终 URL", "浏览器回退状态",
    "浏览器最终 URL", "区域交叉机构", "区域交叉 URL", "访问日期", "台账结论",
}
ALLOWED_TYPES = {"国家CDC/法定监测机构", "国家公共卫生机构", "卫生部/卫生主管部门", "未确认"}
ALLOWED_CONCLUSIONS = {"已验证可检索", "已验证但访问受限", "机构未确认"}


def participant_codes() -> set[str]:
    with LIST_FILE.open("r", encoding="utf-8-sig", newline="") as handle:
        return {row["地区（代码 - 英文名称）"].split(" - ", 1)[0].strip() for row in csv.DictReader(handle)}


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("registry_json", type=Path)
    args = parser.parse_args()
    records = json.loads(args.registry_json.read_text(encoding="utf-8")).get("records", [])
    errors: list[str] = []
    seen: set[str] = set()
    for row in records:
        code = str(row.get("地区代码", "")).strip()
        label = code or "未标明代码的记录"
        missing = [field for field in REQUIRED_FIELDS if not str(row.get(field, "")).strip()]
        if missing:
            errors.append(f"{label}：缺少字段：{'、'.join(sorted(missing))}")
        if code in seen:
            errors.append(f"{label}：地区代码重复")
        seen.add(code)
        if row.get("机构类型") not in ALLOWED_TYPES:
            errors.append(f"{label}：机构类型无效")
        if row.get("台账结论") not in ALLOWED_CONCLUSIONS:
            errors.append(f"{label}：缺少有效权威机构台账结论")
        if row.get("台账结论") == "机构未确认":
            continue
        direct = str(row.get("直接访问状态", ""))
        browser = str(row.get("浏览器回退状态", ""))
        if not direct.startswith("HTTP 2") and browser not in {"不适用：直接访问成功", "浏览器加载成功"}:
            errors.append(f"{label}：直连受限后没有成功或明确记录浏览器回退")
    expected = participant_codes()
    if seen != expected:
        errors.append(f"台账代码与动态名单不一致：缺少{len(expected - seen)}个，多出{len(seen - expected)}个")
    if errors:
        print("\n".join(errors), file=sys.stderr)
        return 1
    unresolved = sum(row.get("台账结论") == "机构未确认" for row in records)
    print(f"权威机构台账校验通过：{len(records)}个国家/地区，其中{unresolved}个国家级入口未确认，需在报告审计页保留回退未完成状态。")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
