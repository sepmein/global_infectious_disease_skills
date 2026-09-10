"""P2：校验检索与证据总账（research JSON），拒绝不完整或无痕迹的检索。

契约版本 2.x，顶层四个必填块：
  sources / local_inputs / disease_sections / records

核心硬性规则：
1. records 必须覆盖动态名单的全部国家和地区，每个国家的 disease_checks 必须覆盖全部配置病种（国家 × 病种全覆盖）。
2. 任何“已完成”的结论都必须有可核对的证据痕迹；未检索不得给出无疫情/零病例结论。
3. 仅用本地资料（internet_required=false）时，本地证据必须已实际读取、有定位、有真实哈希且声明已人工核实，
   否则必须转为需要网络检索并留下检索记录。
4. 本地资料请求路径缺失即阻断，并打印建议路径。
5. 疾病简介必须与 references/疾病简介索引.json 的映射严格一致；本地已有对应简介时必须实际读取，不得借网络回退跳过。
   索引为 null（简介缺失）、新增病种键缺失、或索引指向文件不存在的病种：允许以权威网络资料替代（intro 回退），
   但必须真实检索、引用 authority_type='official'（authority_basis 非空）的已读取已核实同病种来源，
   并登记非空 intro_fallback_reason；绝不伪称已读本地简介，也不得伪造本地简介文件。
6. 全球汇总只允许 country_code=GLOBAL 且病种在配置内，不得据此展开表外国家的逐国事件。

本脚本只读文件、只核对痕迹（含真实 SHA-256），不能替代人工事实核实。
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from skill_config import (
    AUTHORITY_TYPES,
    CATEGORY_ORDER,
    ConfigError,
    DEFAULT_DISEASE_CONFIG,
    DEFAULT_INTRO_DIR,
    DEFAULT_INTRO_INDEX,
    DEFAULT_PARTICIPANT_LIST,
    DISEASE_CHECK_COMPLETED,
    DISEASE_CHECK_STATUSES,
    GLOBAL_SCOPE,
    HUMAN_VERIFICATION_DISCLAIMER,
    PART_KEYS,
    READ_STATUSES,
    RESEARCH_STATUSES,
    SEARCH_STATUSES,
    SOURCE_TYPES,
    VERIFICATION_STATUSES,
    VERIFIED_SET,
    is_valid_date,
    load_diseases,
    load_intro_index,
    load_participants,
    local_material_hint,
    sha256_file,
)

SCHEMA_MAJOR = "2"
NO_EPIDEMIC_CLAIMS: tuple[str, ...] = ("无疫情", "零病例", "无相关疫情", "无病例", "未发生疫情")
DATE_OR_UNPUBLISHED = "未公布"
USABLE_READ_STATUS = "已读取"


class Blocked(Exception):
    """需要立即阻断并给出路径建议的输入问题。"""


def _as_str(value: object) -> str:
    if value is None:
        return ""
    return str(value).strip()


def _is_bool(value: object) -> bool:
    return isinstance(value, bool)


def _date_or_unpublished(value: str) -> bool:
    return value == DATE_OR_UNPUBLISHED or is_valid_date(value)


def _normalize(path_text: str) -> str:
    try:
        return str(Path(path_text).resolve()).lower()
    except OSError:
        return path_text.strip().lower()


# ------------------------------------------------------------------ local_inputs
def _within_roots(path_text: str, roots: list[Path]) -> bool:
    """判断 path_text 是否位于给定根目录之一内部（用于约束文件范围）。"""
    try:
        target = Path(path_text).resolve()
    except OSError:
        return False
    for root in roots:
        try:
            target.relative_to(root.resolve())
            return True
        except ValueError:
            continue
    return False


def validate_local_inputs(payload: dict, errors: list[str], *, intro_dir: Path) -> dict[str, dict]:
    """校验 local_inputs。

    - requested 仅表示“用户提供了本地疫情资料”，不得强迫纯网络检索用户将其置为 true。
    - requested=false 时 files 仅允许位于简介目录（登记必备简介文件），不强制已读取。
    - requested=true 时：paths 非空且须存在；每个目录须非空且目录下全部文件都登记到 files 并已读取；
      所有 files 必须位于声明的 paths 或简介目录内；空目录不可替代上传。
    - 缺失请求路径直接阻断。
    """
    block = payload.get("local_inputs")
    if not isinstance(block, dict):
        errors.append("local_inputs 必须是对象：{requested: bool, paths: [], files: []}")
        return {}

    requested = block.get("requested")
    if not _is_bool(requested):
        errors.append("local_inputs.requested 必须是布尔值")
        requested = bool(requested)

    paths = block.get("paths")
    if not isinstance(paths, list):
        errors.append("local_inputs.paths 必须是数组")
        paths = []
    files = block.get("files")
    if not isinstance(files, list):
        errors.append("local_inputs.files 必须是数组")
        files = []

    intro_root = Path(intro_dir).resolve()

    if requested:
        if not paths:
            raise Blocked("local_inputs.requested 为 true 但 paths 为空：\n" + local_material_hint())
        # 声明路径必须存在；目录须非空且全部文件登记并读取
        allowed_roots: list[Path] = [intro_root]
        for raw in paths:
            text = _as_str(raw)
            if not text or not Path(text).exists():
                raise Blocked(
                    "以下用户指定的本地资料路径不存在，检索与生成一律阻断：\n"
                    + f"  - {text or '（空路径）'}\n"
                    + local_material_hint()
                )
            root = Path(text).resolve()
            allowed_roots.append(root)
            if root.is_dir():
                candidates = root.rglob("*") if block.get("recursive") is True else root.iterdir()
                children = [child for child in candidates if child.is_file() and not child.is_symlink() and child.suffix.lower() in {".pdf", ".doc", ".docx", ".xls", ".xlsx", ".csv", ".tsv", ".json", ".jsonl", ".txt", ".md", ".html", ".htm", ".png", ".jpg", ".jpeg"}]
                if not children:
                    raise Blocked(
                        f"声明的本地资料目录为空，空目录不可替代上传：{text}\n" + local_material_hint()
                    )
                registered = {_normalize(str(item.get("path", ""))) for item in files if isinstance(item, dict)}
                for child in children:
                    if _normalize(str(child)) not in registered:
                        errors.append(
                            f"目录 {text} 下的文件未全部登记到 local_inputs.files：{child}"
                        )
    else:
        # 仅网络检索：files 只能登记简介目录内的必备简介文件
        for index, item in enumerate(files, start=1):
            text = _as_str(item.get("path")) if isinstance(item, dict) else ""
            if text and not _within_roots(text, [intro_root]):
                errors.append(
                    f"local_inputs.files[{index}] 在 requested=false 时只能位于简介目录：{text}"
                )

    if requested:
        registered_paths = {_normalize(str(item.get("path", ""))) for item in files if isinstance(item, dict)}
        for raw in paths:
            path = Path(str(raw))
            if path.is_file() and _normalize(str(path)) not in registered_paths:
                errors.append(f"用户指定文件未读取登记：{path}")

    registry: dict[str, dict] = {}
    for index, item in enumerate(files, start=1):
        label = f"local_inputs.files[{index}]"
        if not isinstance(item, dict):
            errors.append(f"{label} 必须是对象")
            continue
        path_text = _as_str(item.get("path"))
        digest = _as_str(item.get("sha256")).lower()
        read_status = _as_str(item.get("read_status"))
        locator = _as_str(item.get("locator"))
        if not path_text:
            errors.append(f"{label} 缺少 path")
            continue
        file_path = Path(path_text)
        if not file_path.is_file():
            raise Blocked(
                f"{label} 登记的本地文件不存在或不是文件：{path_text}\n" + local_material_hint()
            )
        if read_status not in READ_STATUSES:
            errors.append(f"{label} read_status 无效（{'/'.join(sorted(READ_STATUSES))}）：{read_status or '（空）'}")
        if requested and read_status == USABLE_READ_STATUS and not locator:
            errors.append(f"{label} 标记已读取但缺少 locator 定位")
        if requested and read_status != USABLE_READ_STATUS:
            errors.append(f"{label} requested=true 时登记的本地文件必须已读取，实际：{read_status or '（空）'}")
        if len(digest) != 64 or any(char not in "0123456789abcdef" for char in digest):
            errors.append(f"{label} sha256 必须是 64 位十六进制小写摘要：{digest or '（空）'}")
        else:
            actual = sha256_file(file_path)
            if actual != digest:
                errors.append(
                    f"{label} sha256 与实际文件不一致：记录 {digest}，实际 {actual}（{path_text}）"
                )
        if requested and not _within_roots(path_text, allowed_roots):
            errors.append(
                f"{label} 不在声明的本地资料路径或简介目录内：{path_text}"
            )
        key = _normalize(path_text)
        if key in registry:
            errors.append(f"{label} 本地文件重复登记：{path_text}")
        registry[key] = {
            "path": path_text,
            "read_status": read_status,
            "locator": locator,
            "sha256": digest,
        }
    return registry


# ------------------------------------------------------------------ sources
def validate_sources(
    payload: dict,
    participant_codes: set[str],
    disease_names: set[str],
    local_files: dict[str, dict],
    errors: list[str],
) -> dict[str, dict]:
    sources = payload.get("sources")
    if not isinstance(sources, list) or not sources:
        errors.append("sources 必须是非空数组")
        return {}

    registry: dict[str, dict] = {}
    for index, item in enumerate(sources, start=1):
        label = f"sources[{index}]"
        if not isinstance(item, dict):
            errors.append(f"{label} 必须是对象")
            continue
        source_id = _as_str(item.get("id"))
        if not source_id:
            errors.append(f"{label} 缺少 id")
            continue
        label = f"来源 {source_id}"
        if source_id in registry:
            errors.append(f"{label} id 重复")
            continue

        source_type = _as_str(item.get("type"))
        if source_type not in SOURCE_TYPES:
            errors.append(f"{label} type 必须是 local 或 web，实际：{source_type or '（空）'}")
        path_text = _as_str(item.get("path"))
        url = _as_str(item.get("url"))
        origin_url = _as_str(item.get("source_origin_url"))  # 本地来源的原始出处 URL，允许单独字段
        if source_type == "local":
            if not path_text:
                errors.append(f"{label} 为本地来源但缺少 path")
            else:
                key = _normalize(path_text)
                if not Path(path_text).is_file():
                    errors.append(f"{label} 本地文件不存在：{path_text}")
                elif key not in local_files:
                    errors.append(
                        f"{label} 本地文件未登记在 local_inputs.files（须带真实 sha256 与读取状态）：{path_text}"
                    )
            if url:
                errors.append(f"{label} 为本地来源，不应填写 url（原始出处请用 source_origin_url）")
            if origin_url and not origin_url.startswith(("http://", "https://")):
                errors.append(f"{label} source_origin_url 必须是 http(s) 链接：{origin_url or '（空）'}")
        elif source_type == "web":
            if not url.startswith(("http://", "https://")):
                errors.append(f"{label} 为网络来源，url 必须是 http(s) 链接：{url or '（空）'}")
            if path_text:
                errors.append(f"{label} 为网络来源，不应填写 path")

        locator = _as_str(item.get("locator"))
        if not locator:
            errors.append(f"{label} 缺少 locator（页码、章节、表号或锚点等定位）")

        read_status = _as_str(item.get("read_status"))
        if read_status not in READ_STATUSES:
            errors.append(f"{label} read_status 无效：{read_status or '（空）'}")
        if source_type == "local" and path_text:
            entry = local_files.get(_normalize(path_text))
            if entry and entry["read_status"] != read_status:
                errors.append(
                    f"{label} read_status（{read_status}）与 local_inputs.files 记录（{entry['read_status']}）不一致"
                )

        authority_type = _as_str(item.get("authority_type"))
        authority_basis = _as_str(item.get("authority_basis"))
        if authority_type and authority_type not in AUTHORITY_TYPES:
            errors.append(
                f"{label} authority_type 必须是 {'/'.join(sorted(AUTHORITY_TYPES))} 之一，实际：{authority_type or '（空）'}"
            )

        verification = _as_str(item.get("verification"))
        if verification not in VERIFICATION_STATUSES:
            errors.append(f"{label} verification 无效（{'/'.join(sorted(VERIFICATION_STATUSES))}）：{verification or '（空）'}")

        country_code = _as_str(item.get("country_code"))
        if country_code != GLOBAL_SCOPE and country_code not in participant_codes:
            errors.append(
                f"{label} country_code 必须是动态名单内的代码或 GLOBAL，实际：{country_code or '（空）'}"
            )
        disease = _as_str(item.get("disease"))
        if disease not in disease_names:
            errors.append(f"{label} disease 必须是配置内病种，实际：{disease or '（空）'}")

        for field in ("data_date", "source_date"):
            value = _as_str(item.get(field))
            if not _date_or_unpublished(value):
                errors.append(f"{label} {field} 须为 YYYY-MM-DD 或“未公布”，实际：{value or '（空）'}")
        access_date = _as_str(item.get("access_date"))
        if not is_valid_date(access_date):
            errors.append(f"{label} access_date 须为 YYYY-MM-DD，实际：{access_date or '（空）'}")

        registry[source_id] = {
            "id": source_id,
            "type": source_type,
            "path": path_text,
            "url": url,
            "origin_url": origin_url,
            "locator": locator,
            "read_status": read_status,
            "verification": verification,
            "country_code": country_code,
            "disease": disease,
            "authority_type": authority_type,
            "authority_basis": authority_basis,
        }
    return registry


def _resolve_evidence(
    ids: object,
    sources: dict[str, dict],
    label: str,
    errors: list[str],
) -> list[dict]:
    if not isinstance(ids, list):
        errors.append(f"{label} evidence_ids 必须是数组")
        return []
    resolved: list[dict] = []
    for raw in ids:
        source_id = _as_str(raw)
        source = sources.get(source_id)
        if source is None:
            errors.append(f"{label} 引用了不存在的证据编号：{source_id or '（空）'}")
            continue
        if source["read_status"] != USABLE_READ_STATUS:
            errors.append(
                f"{label} 引用的来源 {source_id} read_status 为“{source['read_status']}”，未读取的来源不能作为证据"
            )
            continue
        resolved.append(source)
    return resolved


def _validate_searches(searches: object, sources: dict[str, dict], label: str, errors: list[str]) -> int:
    """校验网络检索记录。仅“已完成且结果含已读取网络来源”的检索计入有效检索，
    状态为“未完成”或仅引用本地来源冒充的一律不计数。"""
    if not isinstance(searches, list):
        errors.append(f"{label} internet_searches 必须是数组")
        return 0
    effective = 0
    for index, item in enumerate(searches, start=1):
        item_label = f"{label} internet_searches[{index}]"
        if not isinstance(item, dict):
            errors.append(f"{item_label} 必须是对象")
            continue
        if not _as_str(item.get("query")):
            errors.append(f"{item_label} 缺少 query")
        if not _as_str(item.get("tool")):
            errors.append(f"{item_label} 缺少 tool（实际使用的检索工具或入口）")
        search_date = _as_str(item.get("search_date"))
        if not is_valid_date(search_date):
            errors.append(f"{item_label} search_date 须为 YYYY-MM-DD，实际：{search_date or '（空）'}")
        status = _as_str(item.get("status"))
        if status not in SEARCH_STATUSES:
            errors.append(f"{item_label} status 须为“已完成”或“未完成”，实际：{status or '（空）'}")
        result_ids = item.get("result_source_ids")
        if not isinstance(result_ids, list):
            errors.append(f"{item_label} result_source_ids 必须是数组（无结果时给空数组）")
            continue
        web_hit = False
        for raw in result_ids:
            source_id = _as_str(raw)
            if source_id not in sources:
                errors.append(f"{item_label} result_source_ids 引用了不存在的来源：{source_id or '（空）'}")
                continue
            source = sources[source_id]
            if source["type"] == "local":
                errors.append(
                    f"{item_label} 网络检索结果不得引用本地来源冒充：{source_id}（本地资料请在 local_inputs 中登记）"
                )
                continue
            web_hit = True
        if status == "已完成" and web_hit:
            effective += 1
    return effective


# ------------------------------------------------------------------ disease_sections
def validate_disease_sections(
    payload: dict,
    diseases,
    intro_index,
    sources: dict[str, dict],
    require_complete: bool,
    errors: list[str],
) -> int:
    sections = payload.get("disease_sections")
    if not isinstance(sections, list):
        errors.append("disease_sections 必须是数组")
        return 0

    category_by_disease = diseases.category_by_disease
    seen: set[str] = set()
    for index, item in enumerate(sections, start=1):
        label = f"disease_sections[{index}]"
        if not isinstance(item, dict):
            errors.append(f"{label} 必须是对象")
            continue
        disease = _as_str(item.get("disease"))
        if disease not in category_by_disease:
            errors.append(f"{label} disease 不在配置内：{disease or '（空）'}")
            continue
        label = f"病种正文 {disease}"
        if disease in seen:
            errors.append(f"{label} 重复出现")
            continue
        seen.add(disease)

        category = _as_str(item.get("category"))
        if category != category_by_disease[disease]:
            errors.append(
                f"{label} category 与配置不符：记录“{category or '（空）'}”，配置“{category_by_disease[disease]}”"
            )

        blocks = item.get("sections")
        if not isinstance(blocks, dict):
            errors.append(f"{label} sections 必须是对象，键为 {'/'.join(PART_KEYS)}")
            blocks = {}
        missing_keys = [key for key in PART_KEYS if key not in blocks]
        extra_keys = [key for key in blocks if key not in PART_KEYS]
        if missing_keys:
            errors.append(f"{label} sections 缺少：{'、'.join(missing_keys)}")
        if extra_keys:
            errors.append(f"{label} sections 存在未定义键：{'、'.join(extra_keys)}")

        overview_evidence: list[dict] = []
        for key in PART_KEYS:
            block = blocks.get(key)
            if block is None:
                continue
            block_label = f"{label} sections.{key}"
            if not isinstance(block, dict):
                errors.append(f"{block_label} 必须是对象")
                continue
            sufficient = block.get("sufficient")
            local_sufficient = block.get("local_sufficient")
            if not _is_bool(sufficient):
                errors.append(f"{block_label} sufficient 必须是布尔值")
            if not _is_bool(local_sufficient):
                errors.append(f"{block_label} local_sufficient 必须是布尔值")
            evidence = _resolve_evidence(block.get("evidence_ids"), sources, block_label, errors)
            search_count = _validate_searches(block.get("internet_searches"), sources, block_label, errors)
            if key == "overview":
                overview_evidence = evidence

            if sufficient is False and not _as_str(block.get("reason")):
                errors.append(f"{block_label} sufficient 为 false 时必须写明 reason")
            if require_complete and sufficient is not True:
                errors.append(f"{block_label} 尚未达到充分（sufficient≠true），不得进入正式交付")
            if sufficient is True:
                if not evidence:
                    errors.append(f"{block_label} 声明充分但没有任何已读取证据")
                elif not any(
                    source["verification"] in VERIFIED_SET and source["disease"] == disease
                    for source in evidence
                ):
                    errors.append(
                        f"{block_label} 声明充分（sufficient=true）但没有任何同病种已核实（人工核实一致/核实一致）来源"
                    )
            if local_sufficient is True and not any(source["type"] == "local" for source in evidence):
                errors.append(f"{block_label} 声明本地资料充分但未引用任何已读取的本地来源")
            if local_sufficient is False:
                if search_count == 0:
                    errors.append(f"{block_label} 本地资料不充分时必须记录 internet_searches")
                if not any(source["type"] == "web" for source in evidence):
                    errors.append(f"{block_label} 本地资料不充分时必须引用至少一条已读取的网络来源")
            for source in evidence:
                if source["disease"] != disease:
                    errors.append(
                        f"{block_label} 引用的来源 {source['id']} 病种为“{source['disease']}”，与本节病种不一致"
                    )

        # 疾病简介：必须与索引映射严格一致；缺失时允许以权威网络资料替代（绝不伪称已读本地简介）
        expected_file = intro_index.mapping.get(disease) if intro_index else None
        intro_file = _as_str(item.get("intro_file"))
        intro_read = item.get("intro_read")
        intro_locator = _as_str(item.get("intro_locator"))
        intro_fallback_reason = _as_str(item.get("intro_fallback_reason"))
        if intro_index is None:
            continue

        # 判定本地简介是否真实可读：索引非 null 且文件确实存在
        intro_readable = bool(expected_file)
        expected_path = None
        if intro_readable:
            expected_path = intro_index.resolved_path(disease)
            if expected_path is None or not expected_path.is_file():
                intro_readable = False

        if intro_readable:
            # 本地已有对应简介：必须实际读取，不得借网络回退跳过
            if intro_file != expected_file:
                errors.append(
                    f"{label} intro_file 与疾病简介索引不一致：记录“{intro_file or '（空）'}”，索引“{expected_file}”"
                )
            if intro_read is not True:
                errors.append(f"{label} 存在简介文件时 intro_read 必须为 true（须实际读取）")
            if not intro_locator:
                errors.append(f"{label} 缺少 intro_locator（简介中的定位）")
            if expected_path is not None:
                expected_key = _normalize(str(expected_path))
                if not any(
                    source["type"] == "local" and _normalize(source["path"]) == expected_key
                    for source in overview_evidence
                ):
                    errors.append(
                        f"{label} sections.overview 必须引用该简介文件对应的已读取本地来源：{expected_path}"
                    )
        else:
            # 简介缺失（索引为 null / 新增病种键缺失 / 索引指向文件不存在）：允许权威网络回退
            if intro_file:
                errors.append(
                    f"{label} 简介缺失（索引为 null、键缺失或指向文件不存在），不得填写 intro_file“{intro_file}”，"
                    "缺失的简介不能伪造修复"
                )
            if intro_read is True:
                errors.append(f"{label} 简介缺失（索引为 null、键缺失或指向文件不存在），intro_read 不得为 true")
            overview = blocks.get("overview") if isinstance(blocks, dict) else None
            if not isinstance(overview, dict):
                if require_complete:
                    errors.append(
                        f"{label} 简介缺失且采用网络替代时，必须提供 sections.overview 并以网络检索补足"
                    )
            else:
                if overview.get("local_sufficient") is not False:
                    errors.append(
                        f"{label} 简介缺失时 sections.overview.local_sufficient 必须为 false，并以网络检索补足"
                    )
                searches = overview.get("internet_searches")
                if not isinstance(searches, list) or not searches:
                    errors.append(f"{label} 简介缺失时 sections.overview 必须记录 internet_searches")
                if require_complete:
                    if not intro_fallback_reason:
                        errors.append(
                            f"{label} 简介缺失且采用网络替代时，必须填写 intro_fallback_reason"
                            "（说明替代的权威来源、URL、发布日期、访问日期与替代原因）"
                        )
                    # 至少一条“已完成”检索，其 web 结果 id 与 overview.evidence_ids 有交集
                    if isinstance(searches, list) and searches:
                        completed_web_ids: set[str] = set()
                        for search in searches:
                            if not isinstance(search, dict):
                                continue
                            if _as_str(search.get("status")) != "已完成":
                                continue
                            for raw_id in search.get("result_source_ids") or []:
                                sid = _as_str(raw_id)
                                src = sources.get(sid)
                                if src is not None and src["type"] == "web":
                                    completed_web_ids.add(sid)
                        evidence_ids = {_as_str(x) for x in (overview.get("evidence_ids") or [])}
                        if not (completed_web_ids & evidence_ids):
                            errors.append(
                                f"{label} 简介缺失时，至少一条已完成网络检索的 web 结果必须与 overview.evidence_ids 有交集"
                                "（有搜索无对应引用将被拒绝）"
                            )
                    # 至少一条同病种、已读取、已核实、authority_type='official'（authority_basis 非空）的网络来源
                    official_ok = any(
                        src["type"] == "web"
                        and src["disease"] == disease
                        and src["read_status"] == USABLE_READ_STATUS
                        and src["verification"] in VERIFIED_SET
                        and _as_str(src.get("authority_type")) == "official"
                        and bool(_as_str(src.get("authority_basis")))
                        for src in overview_evidence
                    )
                    if not official_ok:
                        errors.append(
                            f"{label} 简介缺失时，sections.overview 必须引用至少一条同病种、已读取、已核实且"
                            " authority_type='official'（authority_basis 非空）的权威网络来源；"
                            "失败/未读/未核实/仅媒体/来源不同病种/来源缺失均不能替代本地简介"
                        )

    expected = set(diseases.names)
    missing = sorted(expected - seen)
    if missing:
        errors.append(f"disease_sections 缺少配置病种：{'、'.join(missing)}")
    return len(seen)


# ------------------------------------------------------------------ records
def validate_records(
    payload: dict,
    participants,
    diseases,
    sources: dict[str, dict],
    require_complete: bool,
    errors: list[str],
) -> dict[str, int]:
    records = payload.get("records")
    stats = {status: 0 for status in DISEASE_CHECK_STATUSES}
    stats["国家已完成"] = 0
    stats["国家检索未完成"] = 0
    if not isinstance(records, list) or not records:
        errors.append("records 必须是非空数组")
        return stats

    category_by_disease = diseases.category_by_disease
    expected_diseases = set(diseases.names)
    seen_codes: set[str] = set()
    for index, record in enumerate(records, start=1):
        label = f"records[{index}]"
        if not isinstance(record, dict):
            errors.append(f"{label} 必须是对象")
            continue
        code = _as_str(record.get("country_code"))
        name = _as_str(record.get("country_name_zh")) or code or "未命名国家"
        label = f"{name}（{code or '缺代码'}）"
        if not code:
            errors.append(f"{label} 缺少 country_code")
        elif code not in participants.codes:
            errors.append(f"{label} country_code 不在动态名单内：{code}")
        elif code in seen_codes:
            errors.append(f"{label} country_code 重复：{code}")
        else:
            seen_codes.add(code)

        research_status = _as_str(record.get("research_status"))
        if research_status not in RESEARCH_STATUSES:
            errors.append(f"{label} research_status 须为“已完成”或“检索未完成”，实际：{research_status or '（空）'}")

        checks = record.get("disease_checks")
        if not isinstance(checks, list):
            errors.append(f"{label} disease_checks 必须是数组")
            continue

        seen_diseases: set[str] = set()
        incomplete = 0
        for order, check in enumerate(checks, start=1):
            check_label = f"{label} disease_checks[{order}]"
            if not isinstance(check, dict):
                errors.append(f"{check_label} 必须是对象")
                continue
            disease = _as_str(check.get("disease"))
            if disease not in category_by_disease:
                errors.append(f"{check_label} disease 不在配置内：{disease or '（空）'}")
                continue
            check_label = f"{label} · {disease}"
            if disease in seen_diseases:
                errors.append(f"{check_label} 病种重复")
                continue
            seen_diseases.add(disease)

            category = _as_str(check.get("category"))
            if category != category_by_disease[disease]:
                errors.append(
                    f"{check_label} category 与配置不符：记录“{category or '（空）'}”，配置“{category_by_disease[disease]}”"
                )

            status = _as_str(check.get("status"))
            if status not in DISEASE_CHECK_STATUSES:
                errors.append(f"{check_label} status 无效（{'/'.join(sorted(DISEASE_CHECK_STATUSES))}）：{status or '（空）'}")
                continue
            stats[status] += 1
            reason = _as_str(check.get("reason"))
            internet_required = check.get("internet_required")
            if not _is_bool(internet_required):
                errors.append(f"{check_label} internet_required 必须是布尔值")
            evidence = _resolve_evidence(check.get("evidence_ids"), sources, check_label, errors)
            search_count = _validate_searches(
                check.get("internet_searches"), sources, check_label, errors
            )

            for source in evidence:
                if source["disease"] != disease:
                    errors.append(
                        f"{check_label} 引用的来源 {source['id']} 病种为“{source['disease']}”，与本条不一致"
                    )
                if source["country_code"] not in {code, GLOBAL_SCOPE}:
                    errors.append(
                        f"{check_label} 引用的来源 {source['id']} 国家为“{source['country_code']}”，"
                        "不属于本国也不是 GLOBAL"
                    )

            local_evidence = [source for source in evidence if source["type"] == "local"]
            web_evidence = [source for source in evidence if source["type"] == "web"]
            country_evidence = [source for source in evidence if source["country_code"] == code]
            country_verified = [
                source
                for source in evidence
                if source["verification"] in VERIFIED_SET and source["country_code"] == code
            ]
            local_verified_same_country = [
                source
                for source in local_evidence
                if source["verification"] in VERIFIED_SET and source["country_code"] == code
            ]

            if status == "检索未完成":
                incomplete += 1
                if not reason:
                    errors.append(f"{check_label} 检索未完成必须写明 reason")
                if any(claim in reason for claim in NO_EPIDEMIC_CLAIMS):
                    errors.append(
                        f"{check_label} 检索未完成却在 reason 中给出无疫情/零病例结论：{reason}"
                    )
                if require_complete:
                    errors.append(f"{check_label} 正式交付前该国家×病种检索尚未完成")
            else:
                if not evidence:
                    errors.append(
                        f"{check_label} 标记为“{status}”但没有任何已读取证据；未检索不得得出结论"
                    )
                if not country_evidence:
                    errors.append(
                        f"{check_label} 缺少本国来源证据；不得仅凭 GLOBAL 全球汇总展开逐国结论"
                    )
                if status == "已完成未纳入" and not reason:
                    errors.append(f"{check_label} 已完成未纳入必须写明 reason（如经核实本期无相关疫情）")
                if status == "已完成并纳入" and not country_verified:
                    errors.append(
                        f"{check_label} 标记纳入但没有任何“同国家”声明已核实（人工核实一致/核实一致）的证据；"
                        "不得借用 GLOBAL 全球汇总已核实而本国未核实通过"
                    )
                if internet_required is False:
                    if web_evidence:
                        errors.append(
                            f"{check_label} internet_required 为 false 却引用了网络来源；请更正为 true"
                        )
                    if not local_evidence:
                        errors.append(
                            f"{check_label} internet_required 为 false 但没有已读取的本地证据"
                        )
                    if not local_verified_same_country:
                        errors.append(
                            f"{check_label} 仅用本地资料完成时，必须有“同国家”声明已核实（人工核实一致/核实一致）的本地证据，"
                            "否则须转为网络检索"
                        )
                    for source in local_evidence:
                        if not source["locator"]:
                            errors.append(f"{check_label} 本地证据 {source['id']} 缺少定位")
                elif internet_required is True:
                    if search_count == 0:
                        errors.append(
                            f"{check_label} internet_required 为 true 但没有已完成的 internet_searches 记录"
                            "（未完成的检索不计入有效检索）"
                        )
                    if not web_evidence:
                        errors.append(
                            f"{check_label} internet_required 为 true 但没有已读取的网络证据"
                        )

        missing_diseases = sorted(expected_diseases - seen_diseases)
        extra = sorted(seen_diseases - expected_diseases)
        if missing_diseases:
            errors.append(
                f"{label} disease_checks 缺少 {len(missing_diseases)} 个配置病种：{'、'.join(missing_diseases)}"
            )
        if extra:
            errors.append(f"{label} disease_checks 存在配置外病种：{'、'.join(extra)}")

        if research_status == "已完成":
            stats["国家已完成"] += 1
            if incomplete:
                errors.append(
                    f"{label} research_status 为“已完成”但仍有 {incomplete} 个病种检索未完成"
                )
        elif research_status == "检索未完成":
            stats["国家检索未完成"] += 1
            if not incomplete:
                errors.append(
                    f"{label} research_status 为“检索未完成”但所有病种均已完成；请更正状态"
                )
            if require_complete:
                errors.append(f"{label} 正式交付前逐国检索尚未完成")

    missing_codes = sorted(participants.codes - seen_codes)
    if missing_codes:
        errors.append(
            f"records 缺少动态名单中的 {len(missing_codes)} 个国家/地区：{'、'.join(missing_codes)}"
        )
    return stats


# ------------------------------------------------------------------ 主流程
def main() -> int:
    parser = argparse.ArgumentParser(description="校验检索与证据总账（contract 2.x）。")
    parser.add_argument("research_json", type=Path)
    parser.add_argument("--list-file", type=Path, default=DEFAULT_PARTICIPANT_LIST)
    parser.add_argument("--disease-config", type=Path, default=DEFAULT_DISEASE_CONFIG)
    parser.add_argument("--intro-index", type=Path, default=DEFAULT_INTRO_INDEX)
    parser.add_argument("--intro-dir", type=Path, default=DEFAULT_INTRO_DIR)
    parser.add_argument(
        "--skip-intro-index",
        action="store_true",
        help="仅在索引尚未提供的中间阶段使用；正式交付门禁不得使用。",
    )
    parser.add_argument(
        "--require-complete",
        action="store_true",
        help="正式交付件门禁：要求国家×病种全部完成、各正文部分均达到充分。",
    )
    args = parser.parse_args()

    try:
        participants = load_participants(args.list_file)
        diseases = load_diseases(args.disease_config)
        intro_index = (
            None
            if args.skip_intro_index
            else load_intro_index(args.intro_index, intro_dir=args.intro_dir, diseases=diseases)
        )
    except ConfigError as exc:
        print(f"配置读取失败：\n{exc}", file=sys.stderr)
        return 1
    if args.skip_intro_index and args.require_complete:
        print("--skip-intro-index 不能与 --require-complete 同时使用。", file=sys.stderr)
        return 1

    try:
        payload = json.loads(args.research_json.read_text(encoding="utf-8"))
    except (OSError, ValueError) as exc:
        print(f"无法读取检索总账 JSON：{args.research_json}（{exc}）", file=sys.stderr)
        return 1
    if not isinstance(payload, dict):
        print("检索总账顶层必须是对象。", file=sys.stderr)
        return 1

    errors: list[str] = []
    version = _as_str(payload.get("schema_version"))
    if not version.startswith(f"{SCHEMA_MAJOR}."):
        errors.append(f"schema_version 必须是 {SCHEMA_MAJOR}.x，实际：{version or '（空）'}")
    period = payload.get("monitoring_period")
    if period is not None:
        if not isinstance(period, dict):
            errors.append("monitoring_period 必须是对象 {start, end}")
        else:
            for key in ("start", "end"):
                value = _as_str(period.get(key))
                if not is_valid_date(value):
                    errors.append(f"monitoring_period.{key} 须为 YYYY-MM-DD，实际：{value or '（空）'}")

    try:
        local_files = validate_local_inputs(payload, errors, intro_dir=args.intro_dir)
    except Blocked as exc:
        print(f"本地资料输入阻断：\n{exc}", file=sys.stderr)
        return 1

    sources = validate_sources(payload, participants.codes, diseases.name_set, local_files, errors)
    section_count = validate_disease_sections(
        payload, diseases, intro_index, sources, args.require_complete, errors
    )
    stats = validate_records(payload, participants, diseases, sources, args.require_complete, errors)

    if errors:
        print("检索总账校验未通过：", file=sys.stderr)
        print("\n".join(f"- {item}" for item in errors), file=sys.stderr)
        print(HUMAN_VERIFICATION_DISCLAIMER, file=sys.stderr)
        return 1

    expected_checks = participants.count * diseases.count
    summary = {
        "research_json": str(args.research_json),
        "schema_version": version,
        "participant_country_count": participants.count,
        "disease_count": diseases.count,
        "expected_country_disease_checks": expected_checks,
        "disease_sections": section_count,
        "source_count": len(sources),
        "local_source_count": sum(1 for source in sources.values() if source["type"] == "local"),
        "web_source_count": sum(1 for source in sources.values() if source["type"] == "web"),
        "declared_verified_sources": sum(
            1 for source in sources.values() if source["verification"] in VERIFIED_SET
        ),
        "unverified_sources": sum(
            1 for source in sources.values() if source["verification"] not in VERIFIED_SET
        ),
        "intro_index": None if intro_index is None else str(intro_index.path),
        "intro_missing_diseases": [] if intro_index is None else list(intro_index.missing_diseases),
        "status_counts": stats,
        "require_complete": args.require_complete,
    }
    print(json.dumps(summary, ensure_ascii=False, indent=2))
    print(HUMAN_VERIFICATION_DISCLAIMER)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
