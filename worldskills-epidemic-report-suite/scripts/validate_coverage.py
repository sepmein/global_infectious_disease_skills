#!/usr/bin/env python3
"""Validate one country-coverage record for every dynamic participant."""

from __future__ import annotations

import argparse
import json
import sys
from collections import Counter
from pathlib import Path

from validate_scope import DEFAULT_PARTICIPANTS, validate_participants


sys.stdout.reconfigure(encoding="utf-8")
sys.stderr.reconfigure(encoding="utf-8")


ALLOWED_STATUS = {
    "VERIFIED_EVENT",
    "CHECKED_NO_SIGNAL",
    "SIGNAL_UNVERIFIED",
    "ACCESS_FAILED",
    "NOT_CHECKED",
}


def read_records(path: Path) -> list[dict]:
    text = path.read_text(encoding="utf-8-sig").strip()
    if not text:
        return []
    try:
        payload = json.loads(text)
    except json.JSONDecodeError:
        payload = None
    if isinstance(payload, list):
        return payload
    if isinstance(payload, dict):
        records = payload.get("records")
        return records if isinstance(records, list) else [payload]
    return [json.loads(line) for line in text.splitlines() if line.strip()]


def list_value(value) -> list[str]:
    if isinstance(value, list):
        return [str(item).strip() for item in value if str(item).strip()]
    return [item.strip() for item in str(value or "").replace("；", ";").split(";") if item.strip()]


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("coverage", type=Path)
    parser.add_argument("--participants", type=Path, default=DEFAULT_PARTICIPANTS)
    parser.add_argument(
        "--require-resolved",
        action="store_true",
        help="Also reject SIGNAL_UNVERIFIED and ACCESS_FAILED records.",
    )
    args = parser.parse_args()
    try:
        countries = validate_participants(args.participants)
        records = read_records(args.coverage)
    except (OSError, ValueError, json.JSONDecodeError) as exc:
        print(str(exc), file=sys.stderr)
        return 1

    expected = {row["country_code"]: row for row in countries}
    seen: set[str] = set()
    errors: list[str] = []
    counts: Counter[str] = Counter()
    for row in records:
        code = str(row.get("country_code", "")).strip()
        label = code or "未标明代码的记录"
        if code not in expected:
            errors.append(f"{label}：不在动态名单中")
            continue
        if code in seen:
            errors.append(f"{label}：覆盖记录重复")
        seen.add(code)
        status = str(row.get("coverage_status", "")).strip()
        if status not in ALLOWED_STATUS:
            errors.append(f"{label}：覆盖状态无效：{status or '空'}")
            continue
        counts[status] += 1
        channels = list_value(row.get("channels_checked"))
        if status != "NOT_CHECKED" and not row.get("checked_at"):
            errors.append(f"{label}：缺少 checked_at")
        if status in {"VERIFIED_EVENT", "CHECKED_NO_SIGNAL", "SIGNAL_UNVERIFIED"} and not channels:
            errors.append(f"{label}：状态为 {status} 但未记录实际检查渠道")
        if status == "VERIFIED_EVENT" and not list_value(row.get("event_ids")):
            errors.append(f"{label}：已核实事件但缺少 event_ids")
        if status in {"SIGNAL_UNVERIFIED", "ACCESS_FAILED"} and not row.get("gap_reason"):
            errors.append(f"{label}：覆盖缺口缺少 gap_reason")
        if status == "NOT_CHECKED":
            errors.append(f"{label}：本轮未检查，禁止正式成稿")
        if args.require_resolved and status in {"SIGNAL_UNVERIFIED", "ACCESS_FAILED"}:
            errors.append(f"{label}：严格发布模式下覆盖状态尚未解决：{status}")

    missing = sorted(set(expected) - seen)
    if missing:
        errors.append(f"缺少 {len(missing)} 个名单条目的覆盖记录：{', '.join(missing)}")
    if errors:
        print("\n".join(errors), file=sys.stderr)
        return 1
    print(
        json.dumps(
            {"participant_count": len(expected), "coverage_count": len(records), "status_counts": counts},
            ensure_ascii=False,
            default=dict,
        )
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
