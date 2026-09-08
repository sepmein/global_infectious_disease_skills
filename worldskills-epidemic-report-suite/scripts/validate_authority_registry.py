#!/usr/bin/env python3
"""Validate a reusable national public-health authority registry."""

from __future__ import annotations

import argparse
import json
import sys
from datetime import date, datetime, timezone
from pathlib import Path

from validate_coverage import read_records
from validate_scope import DEFAULT_PARTICIPANTS, validate_participants


sys.stdout.reconfigure(encoding="utf-8")
sys.stderr.reconfigure(encoding="utf-8")


ALLOWED_STATUS = {"VERIFIED", "ACCESS_RESTRICTED", "UNRESOLVED"}


def parse_date(value: object) -> date | None:
    try:
        return date.fromisoformat(str(value)[:10])
    except (TypeError, ValueError):
        return None


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("registry", type=Path)
    parser.add_argument("--participants", type=Path, default=DEFAULT_PARTICIPANTS)
    parser.add_argument("--max-age-days", type=int, default=30)
    parser.add_argument("--allow-stale", action="store_true")
    args = parser.parse_args()
    try:
        expected = {row["country_code"] for row in validate_participants(args.participants)}
        records = read_records(args.registry)
    except (OSError, ValueError, json.JSONDecodeError) as exc:
        print(str(exc), file=sys.stderr)
        return 1

    today = datetime.now(timezone.utc).date()
    seen: set[str] = set()
    errors: list[str] = []
    stale: list[str] = []
    for row in records:
        code = str(row.get("country_code", "")).strip()
        label = code or "未标明代码的记录"
        if code not in expected:
            errors.append(f"{label}：不在动态名单中")
            continue
        if code in seen:
            errors.append(f"{label}：权威机构记录重复")
        seen.add(code)
        status = str(row.get("verification_status", "")).strip()
        if status not in ALLOWED_STATUS:
            errors.append(f"{label}：verification_status 无效")
        if status != "UNRESOLVED":
            for field in ("authority_name", "authority_type", "canonical_url", "verified_at"):
                if not str(row.get(field, "")).strip():
                    errors.append(f"{label}：缺少 {field}")
        verified_at = parse_date(row.get("verified_at"))
        if verified_at and (today - verified_at).days > args.max_age_days:
            stale.append(code)
        if status == "UNRESOLVED" and not row.get("notes"):
            errors.append(f"{label}：机构未确认但缺少 notes")
    missing = sorted(expected - seen)
    if missing:
        errors.append(f"缺少 {len(missing)} 个国家/地区的机构台账记录")
    if stale and not args.allow_stale:
        errors.append(f"{len(stale)} 个机构身份超过 {args.max_age_days} 天未重新确认：{', '.join(stale)}")
    if errors:
        print("\n".join(errors), file=sys.stderr)
        return 1
    print(json.dumps({"records": len(records), "stale": stale}, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
