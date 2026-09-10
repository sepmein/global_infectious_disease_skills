"""P0：校验并输出世界技能大赛参赛国家和地区动态名单。

名单文件按魔术字节判定真实格式（纯文本 CSV 或真实 OOXML 工作簿），
支持“代码/英文名称/中文名称”三列格式与旧“地区（代码 - 英文名称）+ 中文标准名称”合并格式。
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from skill_config import (
    ConfigError,
    DEFAULT_PARTICIPANT_LIST,
    is_ooxml_zip,
    load_participants,
)


def main() -> int:
    parser = argparse.ArgumentParser(description="校验参赛国家和地区名单并输出结构化结果。")
    parser.add_argument(
        "list_file",
        nargs="?",
        type=Path,
        default=DEFAULT_PARTICIPANT_LIST,
        help=f"名单文件路径，默认 {DEFAULT_PARTICIPANT_LIST}",
    )
    parser.add_argument("--sheet", help="名单为工作簿且存在多个候选表时，明确指定工作表名。")
    parser.add_argument("--quiet", action="store_true", help="只输出统计摘要，不输出完整国家清单。")
    args = parser.parse_args()

    try:
        participants = load_participants(args.list_file, sheet=args.sheet)
    except ConfigError as exc:
        print(f"名单校验未通过：\n{exc}", file=sys.stderr)
        return 1

    payload = {
        "list_file": str(participants.path),
        "real_format": "ooxml" if is_ooxml_zip(participants.path) else "csv",
        "sheet_name": participants.sheet_name,
        "schema": participants.schema_name,
        "participant_country_count": participants.count,
    }
    if not args.quiet:
        payload["countries"] = [
            {
                "country_code": country.code,
                "country_name_en": country.name_en,
                "country_name_zh": country.name_zh,
            }
            for country in participants.countries
        ]
    print(json.dumps(payload, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
