"""P0：校验并输出重点病种检索配置（病种 + 所属分类）。

重点病种检索配置是本轮检索与正文分类的唯一范围主源，必须同时提供“重点病种”和“所属分类”，
分类只能取固定四类；缺少分类、分类越界、病种为空或重复均直接拒绝。
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from skill_config import (
    CATEGORY_HEADINGS,
    CATEGORY_ORDER,
    ConfigError,
    DEFAULT_DISEASE_CONFIG,
    is_ooxml_zip,
    load_diseases,
)


def main() -> int:
    parser = argparse.ArgumentParser(description="校验重点病种检索配置并输出结构化结果。")
    parser.add_argument(
        "config",
        nargs="?",
        type=Path,
        default=DEFAULT_DISEASE_CONFIG,
        help=f"重点病种配置路径，默认 {DEFAULT_DISEASE_CONFIG}",
    )
    parser.add_argument("--sheet", help="配置为工作簿且存在多个候选表时，明确指定工作表名。")
    args = parser.parse_args()

    try:
        config = load_diseases(args.config, sheet=args.sheet)
    except ConfigError as exc:
        print(f"重点病种配置校验未通过：\n{exc}", file=sys.stderr)
        return 1

    by_category = config.by_category
    empty = [category for category, names in by_category.items() if not names]
    payload = {
        "config_file": str(config.path),
        "real_format": "ooxml" if is_ooxml_zip(config.path) else "csv",
        "sheet_name": config.sheet_name,
        "disease_count": config.count,
        "fixed_categories": list(CATEGORY_ORDER),
        "category_headings": list(CATEGORY_HEADINGS),
        "diseases": [
            {"disease": disease.name, "category": disease.category} for disease in config.diseases
        ],
        "diseases_by_category": {category: list(names) for category, names in by_category.items()},
        "categories_without_disease": empty,
    }
    print(json.dumps(payload, ensure_ascii=False, indent=2))
    if empty:
        print(
            "提示：以下固定分类在本轮配置中没有病种，正文仍须保留分类标题："
            + "、".join(empty),
            file=sys.stderr,
        )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
