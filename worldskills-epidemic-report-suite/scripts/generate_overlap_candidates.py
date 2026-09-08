#!/usr/bin/env python3
"""Generate conservative overlap candidates from normalized JSONL evidence records."""

from __future__ import annotations

import argparse
import json
import re
import sys
from datetime import date
from pathlib import Path
from urllib.parse import parse_qsl, urlencode, urlsplit, urlunsplit

sys.stdout.reconfigure(encoding="utf-8")
sys.stderr.reconfigure(encoding="utf-8")

TRACKING_KEYS = {"utm_source", "utm_medium", "utm_campaign", "utm_term", "utm_content", "fbclid", "gclid"}


def clean(value) -> str:
    return re.sub(r"\s+", " ", str(value or "")).strip().casefold()


def normalized_url(value) -> str:
    value = str(value or "").strip()
    if not value:
        return ""
    try:
        parts = urlsplit(value)
        query = urlencode(sorted((key, val) for key, val in parse_qsl(parts.query, keep_blank_values=True) if key.casefold() not in TRACKING_KEYS))
        path = parts.path.rstrip("/") or "/"
        return urlunsplit((parts.scheme.casefold(), parts.netloc.casefold(), path, query, ""))
    except ValueError:
        return value.casefold()


def parse_date(value):
    try:
        return date.fromisoformat(str(value)[:10])
    except (TypeError, ValueError):
        return None


def numeric(value):
    if value in (None, ""):
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def same_values(left, right) -> bool:
    observed = False
    for key in ("case_count", "death_count"):
        left_value, right_value = numeric(left.get(key)), numeric(right.get(key))
        if left_value is None or right_value is None:
            continue
        observed = True
        if left_value != right_value:
            return False
    return observed


def classify(left, right):
    left_hash = clean(left.get("content_sha256") or left.get("sha256"))
    right_hash = clean(right.get("content_sha256") or right.get("sha256"))
    left_urls = {normalized_url(left.get("original_url")), normalized_url(left.get("retrieval_url"))} - {""}
    right_urls = {normalized_url(right.get("original_url")), normalized_url(right.get("retrieval_url"))} - {""}
    if (left_hash and left_hash == right_hash) or left_urls.intersection(right_urls):
        return "SAME_SNAPSHOT", "identical content hash or normalized URL"

    country_same = clean(left.get("country_iso") or left.get("country_region")) == clean(right.get("country_iso") or right.get("country_region"))
    disease_same = clean(left.get("disease_canonical") or left.get("disease")) == clean(right.get("disease_canonical") or right.get("disease"))
    if not (country_same and disease_same):
        return None

    metric_same = clean(left.get("case_metric")) == clean(right.get("case_metric"))
    definition_same = clean(left.get("case_definition")) == clean(right.get("case_definition"))
    left_end, right_end = parse_date(left.get("statistic_end")), parse_date(right.get("statistic_end"))
    cutoff_same = bool(left_end and right_end and left_end == right_end)
    near = bool(left_end and right_end and abs((left_end - right_end).days) <= 45)

    if cutoff_same and metric_same and definition_same:
        if same_values(left, right):
            return "CORROBORATED", "same place, disease, cutoff, metric, definition, and reported values"
        return "CONFLICT", "same place, disease, cutoff, metric, and definition but values differ or are incomplete"
    if near or not (left_end and right_end):
        return "PARTIAL_OVERLAP", "same place and disease with different or incomplete period/metric/definition"
    return None


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("input", help="Normalized evidence JSONL")
    parser.add_argument("--out", required=True, help="Candidate JSONL")
    args = parser.parse_args()

    records = []
    with Path(args.input).open("r", encoding="utf-8-sig") as handle:
        for number, line in enumerate(handle, 1):
            if line.strip():
                record = json.loads(line)
                record.setdefault("record_id", f"R{number:05d}")
                records.append(record)

    pairs = []
    for index, left in enumerate(records):
        for right in records[index + 1:]:
            result = classify(left, right)
            if not result:
                continue
            status, reason = result
            pairs.append({
                "left_record_id": left["record_id"],
                "right_record_id": right["record_id"],
                "left_channel": left.get("source_channel", ""),
                "right_channel": right.get("source_channel", ""),
                "candidate_status": status,
                "candidate_reason": reason,
                "review_status": "REVIEW_REQUIRED",
                "review_decision": "",
            })

    out = Path(args.out).expanduser().resolve()
    out.parent.mkdir(parents=True, exist_ok=True)
    with out.open("w", encoding="utf-8", newline="\n") as handle:
        for pair in pairs:
            handle.write(json.dumps(pair, ensure_ascii=False) + "\n")
    print(json.dumps({"records": len(records), "candidate_pairs": len(pairs), "output": str(out)}, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
