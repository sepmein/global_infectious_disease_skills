#!/usr/bin/env python3
"""Create a read-only, hash-based manifest of a local crawler-data folder."""

from __future__ import annotations

import argparse
import csv
import hashlib
import json
import re
import sys
import zipfile
from datetime import datetime, timezone
from pathlib import Path
from urllib.parse import urlsplit, urlunsplit

sys.stdout.reconfigure(encoding="utf-8")
sys.stderr.reconfigure(encoding="utf-8")

SUPPORTED = {".json", ".jsonl", ".csv", ".tsv", ".xlsx", ".xls", ".html", ".htm", ".txt", ".md", ".xml", ".pdf", ".docx", ".zip"}
TEXT_TYPES = {".json", ".jsonl", ".csv", ".tsv", ".html", ".htm", ".txt", ".md", ".xml"}
ZIP_XML_TYPES = {".docx", ".xlsx"}
URL_RE = re.compile(r"https?://[^\s<>\"'\\)\]]+", re.IGNORECASE)


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def normalize_url(value: str) -> str:
    value = value.rstrip(".,;:，。；：")
    try:
        parts = urlsplit(value)
        return urlunsplit((parts.scheme.lower(), parts.netloc.lower(), parts.path, parts.query, ""))
    except ValueError:
        return value


def extract_urls(path: Path, max_bytes: int) -> list[str]:
    chunks: list[str] = []
    if path.suffix.lower() in TEXT_TYPES:
        chunks.append(path.read_bytes()[:max_bytes].decode("utf-8", errors="ignore"))
    elif path.suffix.lower() in ZIP_XML_TYPES:
        with zipfile.ZipFile(path) as archive:
            used = 0
            for name in archive.namelist():
                if not (name.endswith(".xml") or name.endswith(".rels")):
                    continue
                data = archive.read(name)
                remaining = max_bytes - used
                if remaining <= 0:
                    break
                chunks.append(data[:remaining].decode("utf-8", errors="ignore"))
                used += min(len(data), remaining)
    urls = {normalize_url(match) for chunk in chunks for match in URL_RE.findall(chunk)}
    return sorted(urls)


def iter_files(root: Path):
    for path in sorted(root.rglob("*"), key=lambda p: str(p).lower()):
        if path.is_file() and not path.name.startswith("~$"):
            yield path


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("folder", help="Local crawler-data folder")
    parser.add_argument("--out", required=True, help="Output JSONL manifest")
    parser.add_argument("--csv", dest="csv_out", help="Optional CSV copy")
    parser.add_argument("--max-scan-bytes", type=int, default=5_000_000, help="Maximum bytes scanned per file for embedded URLs")
    args = parser.parse_args()

    root = Path(args.folder).expanduser().resolve()
    if not root.exists() or not root.is_dir():
        print(f"ERROR: local folder is missing or is not a directory: {root}", file=sys.stderr)
        return 2

    generated_at = datetime.now(timezone.utc).isoformat()
    records = []
    for path in iter_files(root):
        stat = path.stat()
        record = {
            "file_id": "",
            "local_absolute_path": str(path.resolve()),
            "local_relative_path": str(path.relative_to(root)),
            "extension": path.suffix.lower(),
            "supported": path.suffix.lower() in SUPPORTED,
            "size_bytes": stat.st_size,
            "modified_at": datetime.fromtimestamp(stat.st_mtime, timezone.utc).isoformat(),
            "sha256": "",
            "embedded_urls": [],
            "inventory_generated_at": generated_at,
            "inventory_error": "",
        }
        try:
            record["sha256"] = sha256_file(path)
            record["file_id"] = f"L-{record['sha256'][:16]}"
            if record["supported"]:
                record["embedded_urls"] = extract_urls(path, args.max_scan_bytes)
        except Exception as exc:
            record["inventory_error"] = f"{type(exc).__name__}: {exc}"
        records.append(record)

    out = Path(args.out).expanduser().resolve()
    out.parent.mkdir(parents=True, exist_ok=True)
    with out.open("w", encoding="utf-8", newline="\n") as handle:
        for record in records:
            handle.write(json.dumps(record, ensure_ascii=False) + "\n")

    if args.csv_out:
        csv_out = Path(args.csv_out).expanduser().resolve()
        csv_out.parent.mkdir(parents=True, exist_ok=True)
        fields = list(records[0].keys()) if records else ["file_id", "local_absolute_path", "local_relative_path", "extension", "supported", "size_bytes", "modified_at", "sha256", "embedded_urls", "inventory_generated_at", "inventory_error"]
        with csv_out.open("w", encoding="utf-8-sig", newline="") as handle:
            writer = csv.DictWriter(handle, fieldnames=fields)
            writer.writeheader()
            for record in records:
                row = dict(record)
                row["embedded_urls"] = ";".join(row["embedded_urls"])
                writer.writerow(row)

    print(json.dumps({"folder": str(root), "files": len(records), "supported": sum(bool(r["supported"]) for r in records), "errors": sum(bool(r["inventory_error"]) for r in records), "manifest": str(out)}, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
