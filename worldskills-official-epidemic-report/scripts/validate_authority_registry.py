"""P1：校验各国权威机构台账。

与旧版的区别：
- 名单来源改为共享模块（支持真实 OOXML 名单文件与两种表头格式），国家覆盖仍要求与动态名单完全一致，不放宽。
- 访问失败不再等于校验失败：直连或浏览器回退失败时，只要求“准确记录了这次尝试”（状态、最终 URL 或明确不适用原因、访问日期）。
- 反过来严禁伪称可用：结论为“已验证可检索”时必须有成功访问痕迹；两项记录均为成功却结论为访问受限也会被拒绝。

本脚本不访问网络，只核对台账记录自身的完整性与自洽性。
"""

from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path

from skill_config import (
    ConfigError,
    DEFAULT_PARTICIPANT_LIST,
    is_valid_date,
    load_participants,
)

# 任何情况下都必须填写的字段
ALWAYS_REQUIRED_FIELDS: tuple[str, ...] = (
    "地区代码",
    "中文标准名称",
    "英文名称",
    "发现检索词",
    "候选页面 URL",
    "机构名称",
    "机构类型",
    "规范入口 URL",
    "直接访问状态",
    "浏览器回退状态",
    "访问日期",
    "台账结论",
)
# 允许以“无”“不适用”“未确认”等明确说明代替具体值的字段
EXPLAINABLE_FIELDS: tuple[str, ...] = (
    "机构权威性核实依据 URL",
    "直接访问最终 URL",
    "浏览器最终 URL",
    "区域交叉机构",
    "区域交叉 URL",
)
ALL_FIELDS: tuple[str, ...] = ALWAYS_REQUIRED_FIELDS + EXPLAINABLE_FIELDS

ALLOWED_TYPES = {"国家CDC/法定监测机构", "国家公共卫生机构", "卫生部/卫生主管部门", "未确认"}
ALLOWED_CONCLUSIONS = {"已验证可检索", "已验证但访问受限", "机构未确认"}

HTTP_STATUS_PATTERN = re.compile(r"^HTTP\s+\d{3}")
# 直接访问状态可接受的失败/受限记录前缀（必须是具体的尝试结果，不能空泛）
DIRECT_FAILURE_PREFIXES: tuple[str, ...] = (
    "连接超时",
    "读取超时",
    "DNS解析失败",
    "TLS错误",
    "证书错误",
    "连接被拒绝",
    "连接重置",
    "网络不可达",
    "重定向循环",
    "需要人机验证",
    "机器人拦截",
    "地域限制",
    "请求被阻断",
)
BROWSER_SUCCESS_VALUES = {"不适用：直接访问成功", "浏览器加载成功"}
BROWSER_OTHER_PREFIXES: tuple[str, ...] = ("浏览器加载失败", "浏览器加载受限", "未执行")
EXPLANATION_PREFIXES: tuple[str, ...] = ("无", "不适用", "未确认", "未获取", "未检索到")


def _text(record: dict[str, object], field: str) -> str:
    return str(record.get(field, "") or "").strip()


def _direct_ok(value: str) -> bool:
    return bool(HTTP_STATUS_PATTERN.match(value)) and value.split()[1].startswith("2")


def _direct_recorded(value: str) -> bool:
    """直接访问状态是否为一次可核对的具体尝试记录。"""
    if HTTP_STATUS_PATTERN.match(value):
        return True
    return any(value.startswith(prefix) for prefix in DIRECT_FAILURE_PREFIXES)


def _browser_recorded(value: str) -> bool:
    if value in BROWSER_SUCCESS_VALUES:
        return True
    return any(value.startswith(prefix) for prefix in BROWSER_OTHER_PREFIXES)


def _explained(value: str) -> bool:
    return any(value.startswith(prefix) for prefix in EXPLANATION_PREFIXES) or value.startswith("http")


def validate(records: list[dict[str, object]], expected_codes: set[str]) -> tuple[list[str], dict[str, int]]:
    errors: list[str] = []
    seen: set[str] = set()
    stats = {conclusion: 0 for conclusion in ALLOWED_CONCLUSIONS}
    stats["直连成功"] = 0
    stats["浏览器回退成功"] = 0
    stats["访问失败但记录完整"] = 0

    for record in records:
        code = _text(record, "地区代码")
        label = code or "未标明代码的记录"
        missing = [field for field in ALWAYS_REQUIRED_FIELDS if not _text(record, field)]
        if missing:
            errors.append(f"{label}：缺少必填字段：{'、'.join(missing)}")
        for field in EXPLAINABLE_FIELDS:
            value = _text(record, field)
            if not value:
                errors.append(f"{label}：字段“{field}”为空；无此项时须写“无”或“不适用：原因”")
            elif not _explained(value):
                errors.append(f"{label}：字段“{field}”应为 URL 或“无/不适用/未确认”类明确说明，实际为：{value}")

        if code:
            if code in seen:
                errors.append(f"{label}：地区代码重复")
            seen.add(code)

        access_date = _text(record, "访问日期")
        if access_date and not is_valid_date(access_date):
            errors.append(f"{label}：访问日期须为 YYYY-MM-DD，实际为：{access_date}")

        org_type = _text(record, "机构类型")
        if org_type and org_type not in ALLOWED_TYPES:
            errors.append(f"{label}：机构类型无效：{org_type}")

        conclusion = _text(record, "台账结论")
        if conclusion not in ALLOWED_CONCLUSIONS:
            errors.append(f"{label}：台账结论无效或缺失：{conclusion or '（空）'}")
            continue
        stats[conclusion] += 1

        direct = _text(record, "直接访问状态")
        browser = _text(record, "浏览器回退状态")
        if direct and not _direct_recorded(direct):
            errors.append(
                f"{label}：直接访问状态不是可核对的尝试记录（应为“HTTP <状态码>”或明确失败原因）：{direct}"
            )
        if browser and not _browser_recorded(browser):
            errors.append(
                f"{label}：浏览器回退状态不是可核对的尝试记录（应为“不适用：直接访问成功”"
                f"“浏览器加载成功”“浏览器加载失败：原因”或“未执行：原因”）：{browser}"
            )

        direct_success = _direct_ok(direct)
        browser_success = browser == "浏览器加载成功"
        if direct_success:
            stats["直连成功"] += 1
        elif browser_success:
            stats["浏览器回退成功"] += 1
        elif conclusion != "机构未确认":
            stats["访问失败但记录完整"] += 1

        if conclusion == "已验证可检索" and not (direct_success or browser_success):
            errors.append(
                f"{label}：结论为“已验证可检索”但没有任何成功访问痕迹；"
                "不得伪称可用，请改为“已验证但访问受限”或“机构未确认”"
            )
        if conclusion == "已验证但访问受限" and direct_success and browser in BROWSER_SUCCESS_VALUES:
            errors.append(
                f"{label}：结论为“已验证但访问受限”但直连与浏览器记录均为成功；请更正结论或如实记录受限原因"
            )
        if conclusion == "机构未确认" and org_type not in {"未确认", ""}:
            errors.append(f"{label}：结论为“机构未确认”时机构类型应记为“未确认”，实际为：{org_type}")

    missing_codes = sorted(expected_codes - seen)
    extra_codes = sorted(seen - expected_codes)
    if missing_codes:
        errors.append(f"台账缺少动态名单中的 {len(missing_codes)} 个国家/地区：{'、'.join(missing_codes)}")
    if extra_codes:
        errors.append(f"台账存在动态名单外的 {len(extra_codes)} 个代码：{'、'.join(extra_codes)}")
    return errors, stats


def main() -> int:
    parser = argparse.ArgumentParser(description="校验权威机构台账记录的完整性与自洽性（不访问网络）。")
    parser.add_argument("registry_json", type=Path)
    parser.add_argument("--list-file", type=Path, default=DEFAULT_PARTICIPANT_LIST)
    parser.add_argument("--sheet", help="名单工作簿存在多个候选表时指定工作表名。")
    args = parser.parse_args()

    try:
        participants = load_participants(args.list_file, sheet=args.sheet)
    except ConfigError as exc:
        print(f"名单读取失败：\n{exc}", file=sys.stderr)
        return 1
    try:
        payload = json.loads(args.registry_json.read_text(encoding="utf-8"))
    except (OSError, ValueError) as exc:
        print(f"无法读取台账 JSON：{args.registry_json}（{exc}）", file=sys.stderr)
        return 1
    records = payload.get("records")
    if not isinstance(records, list) or not records:
        print("台账 JSON 顶层需要非空的 records 数组。", file=sys.stderr)
        return 1

    errors, stats = validate(records, participants.codes)
    if errors:
        print("权威机构台账校验未通过：", file=sys.stderr)
        print("\n".join(errors), file=sys.stderr)
        return 1

    print(
        "权威机构台账校验通过："
        f"{len(records)} 个国家/地区，覆盖动态名单 {participants.count} 项；"
        f"直连成功 {stats['直连成功']}，浏览器回退成功 {stats['浏览器回退成功']}，"
        f"访问失败但记录完整 {stats['访问失败但记录完整']}，机构未确认 {stats['机构未确认']}。"
    )
    print(
        "注意：本脚本只核对台账记录自身，未验证机构入口当前是否真的可访问；"
        "访问受限与机构未确认项必须在报告审计页如实保留为受限/未完成。"
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
