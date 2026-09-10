"""P0：校验本轮可用的输入资源（疾病简介索引、简介文件、本地资料根目录）。

- 疾病简介索引由主代理提供，格式为 {病种: 相对文件名 或 null}，键集合必须与重点病种配置严格一致。
- 非 null 且确实存在的简介文件必须真实存在，并输出真实 SHA-256 与大小，供检索阶段登记痕迹。
- 索引值为 null（如水痘）、新增病种键缺失、或索引指向的文件不存在：本地简介缺失，允许以权威网络资料替代，
  但必须真实执行检索与访问、登记 official 来源与原因，且绝不伪称已读本地简介；脚本输出 fallback_required 而非阻断。
- 索引文件本身损坏（无法解析为 JSON）仍阻断；路径穿越与绝对路径仍拒绝；配置外病种（额外键）仍拒绝。
- 本地资料根目录缺失或被显式要求但为空时，给出明确阻断与建议路径。

本脚本只读，不创建、不修改任何用户资料。
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from skill_config import (
    ConfigError,
    DEFAULT_DISEASE_CONFIG,
    DEFAULT_INTRO_DIR,
    DEFAULT_INTRO_INDEX,
    DEFAULT_LOCAL_MATERIAL_DIR,
    DEFAULT_TEMPLATE_DOCX,
    load_diseases,
    load_intro_index,
    local_material_hint,
    sha256_file,
)


def main() -> int:
    parser = argparse.ArgumentParser(description="校验疾病简介索引、简介文件与本地资料输入。")
    parser.add_argument("--disease-config", type=Path, default=DEFAULT_DISEASE_CONFIG)
    parser.add_argument("--intro-index", type=Path, default=DEFAULT_INTRO_INDEX)
    parser.add_argument("--intro-dir", type=Path, default=DEFAULT_INTRO_DIR)
    parser.add_argument("--local-dir", type=Path, default=DEFAULT_LOCAL_MATERIAL_DIR)
    parser.add_argument("--template", type=Path, default=DEFAULT_TEMPLATE_DOCX)
    parser.add_argument(
        "--require-local-inputs",
        action="store_true",
        help="用户明确提供了本地疫情资料时使用：要求本地资料目录存在且非空。",
    )
    parser.add_argument(
        "--require-template",
        action="store_true",
        help="要求写作模板存在（build 流程默认开启）。",
    )
    args = parser.parse_args()

    errors: list[str] = []
    try:
        diseases = load_diseases(args.disease_config)
    except ConfigError as exc:
        print(f"重点病种配置读取失败：\n{exc}", file=sys.stderr)
        return 1

    if not args.intro_dir.is_dir():
        print(
            f"疾病简介目录不存在：{args.intro_dir}；将把所有本地简介视为缺失并以权威网络资料替代。\n"
            + local_material_hint(),
            file=sys.stderr,
        )

    try:
        intro_index = load_intro_index(args.intro_index, intro_dir=args.intro_dir, diseases=diseases)
    except ConfigError as exc:
        print(f"疾病简介索引校验未通过：\n{exc}", file=sys.stderr)
        return 1

    intros: list[dict[str, object]] = []
    missing_files: list[str] = []
    fallback_required = False
    for disease in diseases.names:
        filename = intro_index.mapping.get(disease)
        if not filename:
            # 索引为 null，或新增病种键缺失被自动映射为 None：本地简介缺失，允许网络回退
            intros.append(
                {"disease": disease, "intro_file": None, "status": "简介缺失（索引为null或键缺失）", "fallback": True}
            )
            fallback_required = True
            continue
        path = intro_index.resolved_path(disease)
        if path is None or not path.is_file():
            # 索引指向的文件不存在：允许以权威网络资料替代（不得改索引掩盖，但也不阻断）
            missing_files.append(f"{disease} -> {filename}")
            intros.append(
                {"disease": disease, "intro_file": filename, "status": "索引指向的文件不存在（可网络回退）", "fallback": True}
            )
            fallback_required = True
            continue
        intros.append(
            {
                "disease": disease,
                "intro_file": filename,
                "status": "可用",
                "path": str(path),
                "size_bytes": path.stat().st_size,
                "sha256": sha256_file(path),
            }
        )
    if missing_files:
        print(
            "以下病种索引指向的简介文件不存在，将以权威网络资料替代，不得伪称已读本地简介：\n"
            + "\n".join(f"  - {item}" for item in missing_files),
            file=sys.stderr,
        )
    # 简介缺失（索引为 null / 键缺失 / 指向文件不存在）不再阻断正式交付：允许以权威网络资料替代，
    # 但下游 P2 必须真实检索、引用 official 来源并登记 intro_fallback_reason，且不得伪称已读本地简介。
    missing_intros = [item["disease"] for item in intros if str(item["status"]).startswith("简介缺失")]
    if missing_intros:
        print(
            "以下病种简介确实缺失（索引为 null 或键缺失），将以权威网络资料替代，不得伪称已读本地简介或伪造修复："
            + "、".join(str(item) for item in missing_intros),
            file=sys.stderr,
        )
    for warning in intro_index.warnings:
        print(f"提示：{warning}", file=sys.stderr)

    unreferenced = []
    if args.intro_dir.is_dir():
        unreferenced = sorted(
            item.name
            for item in args.intro_dir.iterdir()
            if item.is_file() and item.name not in {value for value in intro_index.mapping.values() if value}
        )

    # 仅在明确要求本地疫情资料时才扫描本地目录；默认不扫描，
    # 以支持只通过检索总账 local_inputs.paths 提供单个本地文件（无需整个目录）。
    local_dir_exists = False
    local_files: list[str] = []
    if args.require_local_inputs:
        local_dir_exists = args.local_dir.is_dir()
        local_files = (
            sorted(str(path) for path in args.local_dir.iterdir() if path.is_file())
            if local_dir_exists
            else []
        )
        if not local_dir_exists:
            errors.append(f"要求提供本地疫情资料，但目录不存在：{args.local_dir}\n" + local_material_hint())
        elif not local_files:
            errors.append(f"要求提供本地疫情资料，但目录为空：{args.local_dir}\n" + local_material_hint())

    template_ok = args.template.is_file()
    if args.require_template and not template_ok:
        errors.append(f"写作模板不存在：{args.template}")

    if errors:
        print("输入资源校验未通过：", file=sys.stderr)
        print("\n".join(f"- {item}" for item in errors), file=sys.stderr)
        return 1

    payload = {
        "disease_config": str(diseases.path),
        "disease_count": diseases.count,
        "intro_index": str(intro_index.path),
        "intro_dir": str(args.intro_dir),
        "intro_available": sum(1 for item in intros if item["status"] == "可用"),
        "intro_missing": [item["disease"] for item in intros if item["status"] in ("简介缺失", "简介缺失（索引为null或键缺失）")],
        "intro_fallback_required": fallback_required,
        "intros": intros,
        "intro_dir_unreferenced_files": unreferenced,
        "local_material_dir": str(args.local_dir),
        "local_material_dir_exists": local_dir_exists,
        "local_material_file_count": len(local_files),
        "template_docx": str(args.template),
        "template_exists": template_ok,
    }
    print(json.dumps(payload, ensure_ascii=False, indent=2))
    if payload["intro_missing"]:
        print(
            "提示：以下病种本地简介缺失，本轮将以权威网络资料替代并登记 intro_fallback_reason 与 official 来源；"
            "绝不伪称已读本地简介，也不得伪造本地简介文件："
            + "、".join(str(item) for item in payload["intro_missing"]),
            file=sys.stderr,
        )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
