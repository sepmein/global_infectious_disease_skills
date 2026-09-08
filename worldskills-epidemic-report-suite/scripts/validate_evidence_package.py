#!/usr/bin/env python3
"""Validate Source, Event, and Observation linkage for an evidence package."""

from __future__ import annotations

import argparse
import json
import math
import sys
from pathlib import Path

from validate_coverage import list_value, read_records


sys.stdout.reconfigure(encoding="utf-8")
sys.stderr.reconfigure(encoding="utf-8")


ALLOWED_CHANNELS = {"LOCAL", "WEB"}
ALLOWED_GRADES = {"A", "B", "C", "D"}
INCLUDED_VALUES = {"INCLUDED", "已纳入", "YES", "TRUE"}


def non_negative_number(value: object) -> bool:
    if value in (None, ""):
        return True
    if isinstance(value, bool):
        return False
    try:
        number = float(value)
    except (TypeError, ValueError):
        return False
    return math.isfinite(number) and number >= 0


def unique_by(records: list[dict], field: str, label: str, errors: list[str]) -> dict[str, dict]:
    result: dict[str, dict] = {}
    for row_number, row in enumerate(records, start=1):
        record_id = str(row.get(field, "")).strip()
        if not record_id:
            errors.append(f"{label} 第 {row_number} 条缺少 {field}")
        elif record_id in result:
            errors.append(f"{label} {field} 重复：{record_id}")
        else:
            result[record_id] = row
    return result


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--sources", type=Path, required=True)
    parser.add_argument("--events", type=Path, required=True)
    parser.add_argument("--observations", type=Path, required=True)
    args = parser.parse_args()
    try:
        source_rows = read_records(args.sources)
        event_rows = read_records(args.events)
        observation_rows = read_records(args.observations)
    except (OSError, ValueError, json.JSONDecodeError) as exc:
        print(str(exc), file=sys.stderr)
        return 1

    errors: list[str] = []
    sources = unique_by(source_rows, "source_id", "Source", errors)
    events = unique_by(event_rows, "event_id", "Event", errors)
    observations = unique_by(observation_rows, "observation_id", "Observation", errors)

    for source_id, row in sources.items():
        channel = str(row.get("source_channel", "")).strip()
        grade = str(row.get("evidence_grade", "")).strip().upper()
        if channel not in ALLOWED_CHANNELS:
            errors.append(f"{source_id}：source_channel 无效")
        if grade not in ALLOWED_GRADES:
            errors.append(f"{source_id}：evidence_grade 无效")
        for field in ("publisher", "title", "locator"):
            if not str(row.get(field, "")).strip():
                errors.append(f"{source_id}：缺少 {field}")
        if channel == "LOCAL":
            if not row.get("local_absolute_path") or not row.get("sha256"):
                errors.append(f"{source_id}：本地来源缺少绝对路径或 SHA-256")
        elif not row.get("original_url") and not row.get("retrieval_url"):
            errors.append(f"{source_id}：网络来源缺少 URL")

    for event_id, row in events.items():
        for field in ("country_code", "disease_canonical"):
            if not str(row.get(field, "")).strip():
                errors.append(f"{event_id}：缺少 {field}")
        source_ids = list_value(row.get("source_ids"))
        missing = sorted(set(source_ids) - set(sources))
        if missing:
            errors.append(f"{event_id}：引用不存在的 Source_ID：{', '.join(missing)}")
        included = str(row.get("inclusion_status", "")).strip().upper() in INCLUDED_VALUES
        if included and not source_ids:
            errors.append(f"{event_id}：已纳入但没有来源")
        if included and source_ids and all(
            str(sources[source_id].get("evidence_grade", "")).upper() == "D"
            for source_id in source_ids
            if source_id in sources
        ):
            errors.append(f"{event_id}：已纳入事件仅由 D 级线索支持")

    for observation_id, row in observations.items():
        event_id = str(row.get("event_id", "")).strip()
        if event_id not in events:
            errors.append(f"{observation_id}：引用不存在的 Event_ID：{event_id or '空'}")
        if not row.get("observed_at"):
            errors.append(f"{observation_id}：缺少 observed_at")
        for field in ("case_count", "death_count"):
            if not non_negative_number(row.get(field)):
                errors.append(f"{observation_id}：{field} 必须为空或非负数")
        source_ids = list_value(row.get("source_ids"))
        if not source_ids:
            errors.append(f"{observation_id}：缺少 source_ids")
        missing = sorted(set(source_ids) - set(sources))
        if missing:
            errors.append(f"{observation_id}：引用不存在的 Source_ID：{', '.join(missing)}")

    if errors:
        print("\n".join(errors), file=sys.stderr)
        return 1
    print(
        json.dumps(
            {"sources": len(sources), "events": len(events), "observations": len(observations)},
            ensure_ascii=False,
        )
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
