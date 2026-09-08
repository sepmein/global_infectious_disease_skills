#!/usr/bin/env python3
"""Compare current and previous event observations and emit material changes."""

from __future__ import annotations

import argparse
import json
import sys
from datetime import datetime
from pathlib import Path

from validate_coverage import list_value, read_records


sys.stdout.reconfigure(encoding="utf-8")
sys.stderr.reconfigure(encoding="utf-8")


def parse_number(value: object) -> float | None:
    if value in (None, "") or isinstance(value, bool):
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def sort_key(row: dict) -> tuple[str, str, str]:
    return (
        str(row.get("observed_at", "")),
        str(row.get("statistic_end", "")),
        str(row.get("observation_id", "")),
    )


def latest_by_event(records: list[dict]) -> dict[str, dict]:
    latest: dict[str, dict] = {}
    for row in records:
        event_id = str(row.get("event_id", "")).strip()
        if not event_id:
            continue
        if event_id not in latest or sort_key(row) > sort_key(latest[event_id]):
            latest[event_id] = row
    return latest


def area_set(value: object) -> set[str]:
    if isinstance(value, list):
        return {str(item).strip() for item in value if str(item).strip()}
    normalized = str(value or "").replace("、", ";").replace(",", ";").replace("，", ";")
    return set(list_value(normalized))


def compare(previous: dict | None, current: dict) -> list[dict]:
    event_id = str(current.get("event_id", "")).strip()
    changes: list[dict] = []
    if previous is None:
        return [{"change_type": "NEW_EVENT", "detail": "本期首次出现该事件"}]

    metric_changed = any(
        str(previous.get(field, "")).strip() != str(current.get(field, "")).strip()
        for field in ("case_metric", "case_definition", "death_definition")
    )
    previous_cases, current_cases = parse_number(previous.get("case_count")), parse_number(current.get("case_count"))
    previous_deaths, current_deaths = parse_number(previous.get("death_count")), parse_number(current.get("death_count"))
    if metric_changed:
        changes.append({"change_type": "DATA_REVISION", "detail": "病例或死亡统计口径发生变化"})
    else:
        if previous_cases is not None and current_cases is not None:
            if current_cases > previous_cases:
                changes.append(
                    {"change_type": "CASE_INCREASE", "detail": f"病例数由 {previous_cases:g} 增至 {current_cases:g}", "delta": current_cases - previous_cases}
                )
            elif current_cases < previous_cases:
                changes.append({"change_type": "DATA_REVISION", "detail": f"病例数由 {previous_cases:g} 修订为 {current_cases:g}"})
        if previous_deaths is not None and current_deaths is not None:
            if current_deaths > previous_deaths:
                changes.append(
                    {"change_type": "DEATH_INCREASE", "detail": f"死亡数由 {previous_deaths:g} 增至 {current_deaths:g}", "delta": current_deaths - previous_deaths}
                )
            elif current_deaths < previous_deaths:
                changes.append({"change_type": "DATA_REVISION", "detail": f"死亡数由 {previous_deaths:g} 修订为 {current_deaths:g}"})

    new_areas = sorted(area_set(current.get("affected_areas")) - area_set(previous.get("affected_areas")))
    if new_areas:
        changes.append({"change_type": "AREA_EXPANSION", "detail": f"新增受影响地区：{'、'.join(new_areas)}"})
    for field, change_type, label in (
        ("event_status", "STATUS_CHANGE", "事件状态"),
        ("pathogen_status", "PATHOGEN_CHANGE", "病原体确认状态"),
    ):
        old, new = str(previous.get(field, "")).strip(), str(current.get(field, "")).strip()
        if old != new and new:
            changes.append({"change_type": change_type, "detail": f"{label}由“{old or '未报告'}”变为“{new}”"})
    return changes


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--current", type=Path, required=True)
    parser.add_argument("--previous", type=Path)
    parser.add_argument("--out", type=Path, required=True)
    args = parser.parse_args()
    current = latest_by_event(read_records(args.current))
    previous = latest_by_event(read_records(args.previous)) if args.previous else {}
    emitted: list[dict] = []
    for event_id, current_row in sorted(current.items()):
        for change in compare(previous.get(event_id), current_row):
            emitted.append(
                {
                    "event_id": event_id,
                    "current_observation_id": current_row.get("observation_id", ""),
                    "previous_observation_id": (previous.get(event_id) or {}).get("observation_id", ""),
                    "detected_at": datetime.now().astimezone().isoformat(),
                    **change,
                }
            )
    args.out.parent.mkdir(parents=True, exist_ok=True)
    with args.out.open("w", encoding="utf-8", newline="\n") as handle:
        for row in emitted:
            handle.write(json.dumps(row, ensure_ascii=False) + "\n")
    print(json.dumps({"current_events": len(current), "previous_events": len(previous), "material_changes": len(emitted), "output": str(args.out.resolve())}, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
