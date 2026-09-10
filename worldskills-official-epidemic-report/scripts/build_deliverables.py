"""唯一允许的交付件生成流程：先过全部前置门禁，再调用外部生成器，最后过结构与渲染门禁。

门禁顺序：
  P0a 名单校验            validate_participant_list.py
  P0b 重点病种配置校验    validate_priority_disease_config.py（默认 xlsx）
  P0c 输入资源校验        validate_input_resources.py（疾病简介索引/简介文件/模板，可要求本地资料）
  P1  权威机构台账校验    validate_authority_registry.py
  P2  检索总账校验        validate_country_research.py --require-complete
  ——  运行 --builder 指定的外部生成器（必须是本机只读加工，禁止爬虫/联网抓取）
  P4  交付件结构校验      validate_deliverables.py --xlsx --docx --research
  P5  渲染核验            读取 --render-check JSON，要求 docx_sha256 与当前 docx 实际哈希一致

渲染核验契约（--render-check 指向的 JSON）：
  {"docx_sha256": "<64位小写十六进制>", "checked": true, "pages": <正整数>, "issues": []}
必须由真实渲染器产生（可用 --renderer 让本脚本调用），不得手写捏造；缺失或不匹配一律返回非 0 并声明未完成。
"""

from __future__ import annotations

import argparse
import json
import shutil
import subprocess
import sys
from pathlib import Path

from skill_config import (
    DEFAULT_DISEASE_CONFIG,
    DEFAULT_INTRO_DIR,
    DEFAULT_INTRO_INDEX,
    DEFAULT_LOCAL_MATERIAL_DIR,
    DEFAULT_PARTICIPANT_LIST,
    DEFAULT_TEMPLATE_DOCX,
    sha256_file,
)

SCRIPTS_DIR = Path(__file__).resolve().parent
RENDER_CHECK_REQUIRED_KEYS = ("docx_sha256", "checked", "pages", "issues")


def run_gate(label: str, command: list[str]) -> None:
    print(f"==> {label}")
    completed = subprocess.run(command, check=False)
    if completed.returncode:
        raise SystemExit(f"{label}未通过；禁止生成或修改交付件。")


def validate_render_check(path: Path, docx: Path) -> list[str]:
    if not path.is_file():
        return [f"渲染核验结果不存在：{path}；请用 --renderer 生成，或说明渲染核验未完成"]
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, ValueError) as exc:
        return [f"渲染核验结果无法解析为 JSON：{path}（{exc}）"]
    if not isinstance(payload, dict):
        return [f"渲染核验结果顶层必须是对象：{path}"]

    errors: list[str] = []
    missing = [key for key in RENDER_CHECK_REQUIRED_KEYS if key not in payload]
    if missing:
        errors.append(f"渲染核验结果缺少字段：{'、'.join(missing)}")
        return errors
    if payload.get("checked") is not True:
        errors.append("渲染核验结果 checked 必须为 true")
    pages = payload.get("pages")
    if not isinstance(pages, int) or isinstance(pages, bool) or pages <= 0:
        errors.append(f"渲染核验结果 pages 必须是正整数，实际：{pages!r}")
    issues = payload.get("issues")
    if not isinstance(issues, list):
        errors.append("渲染核验结果 issues 必须是数组")
    elif issues:
        errors.append(f"渲染核验报告了 {len(issues)} 个问题：{issues[:10]}")
    recorded = str(payload.get("docx_sha256", "")).strip().lower()
    actual = sha256_file(docx)
    if recorded != actual:
        errors.append(
            f"渲染核验结果与当前 Word 不是同一文件：记录 {recorded or '（空）'}，实际 {actual}"
        )
    return errors


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--registry", type=Path, required=True, help="权威机构台账 JSON。")
    parser.add_argument("--research", type=Path, required=True, help="检索总账 JSON（contract 2.x）。")
    parser.add_argument("--xlsx", type=Path, required=True, help="待生成的 Excel 底表路径（必须尚不存在）。")
    parser.add_argument("--docx", type=Path, required=True, help="待生成的 Word 报告路径（必须尚不存在）。")
    parser.add_argument("--render-check", type=Path, required=True, help="渲染核验结果 JSON 路径。")
    parser.add_argument("--list-file", type=Path, default=DEFAULT_PARTICIPANT_LIST)
    parser.add_argument("--disease-config", type=Path, default=DEFAULT_DISEASE_CONFIG)
    parser.add_argument("--intro-index", type=Path, default=DEFAULT_INTRO_INDEX)
    parser.add_argument("--intro-dir", type=Path, default=DEFAULT_INTRO_DIR)
    parser.add_argument("--local-dir", type=Path, default=DEFAULT_LOCAL_MATERIAL_DIR)
    parser.add_argument(
        "--require-local-inputs",
        action="store_true",
        help="用户提供了本地疫情资料时使用：要求本地资料目录存在且非空。",
    )
    parser.add_argument("--template", type=Path, default=DEFAULT_TEMPLATE_DOCX, help="写作模板（只读）。")
    parser.add_argument(
        "--template-copy",
        type=Path,
        help="把写作模板复制到该路径供生成器使用；目标不得已存在，源模板绝不修改。",
    )
    parser.add_argument(
        "--renderer",
        nargs=argparse.REMAINDER,
        help="渲染器命令；出现 {docx}/{render_check} 占位符时按位替换，否则在命令末尾追加这两个路径。"
        " 必须真实渲染并写出渲染核验 JSON。与 --builder 互斥地放在命令末尾之前。",
    )
    parser.add_argument(
        "--builder",
        nargs=argparse.REMAINDER,
        help="生成器命令（必填，放在命令最后）；只能在全部前置门禁通过后执行，且必须是本机只读加工，禁止爬虫。",
    )
    args = parser.parse_args()

    builder = args.builder or []
    renderer = args.renderer or []
    # argparse 的 REMAINDER 会吞掉后续参数，若两者同时给出需手工切分。
    if "--builder" in renderer:
        index = renderer.index("--builder")
        builder = renderer[index + 1 :]
        renderer = renderer[:index]
    if "--renderer" in builder:
        index = builder.index("--renderer")
        renderer = builder[index + 1 :]
        builder = builder[:index]
    if not builder:
        parser.error("--builder 后必须提供生成器命令")

    if args.xlsx.exists() or args.docx.exists():
        raise SystemExit("目标交付件已存在；为避免绕过门禁编辑现有文件，请使用新的版本文件名。")

    python = sys.executable
    run_gate("P0a 名单校验", [python, str(SCRIPTS_DIR / "validate_participant_list.py"), str(args.list_file), "--quiet"])
    run_gate(
        "P0b 重点病种配置校验",
        [python, str(SCRIPTS_DIR / "validate_priority_disease_config.py"), str(args.disease_config)],
    )
    resources = [
        python,
        str(SCRIPTS_DIR / "validate_input_resources.py"),
        "--disease-config",
        str(args.disease_config),
        "--intro-index",
        str(args.intro_index),
        "--intro-dir",
        str(args.intro_dir),
        "--local-dir",
        str(args.local_dir),
        "--template",
        str(args.template),
        "--require-template",
    ]
    if args.require_local_inputs:
        resources.append("--require-local-inputs")
    run_gate("P0c 输入资源校验", resources)
    run_gate(
        "P1 权威机构台账校验",
        [
            python,
            str(SCRIPTS_DIR / "validate_authority_registry.py"),
            str(args.registry),
            "--list-file",
            str(args.list_file),
        ],
    )
    run_gate(
        "P2 检索总账校验",
        [
            python,
            str(SCRIPTS_DIR / "validate_country_research.py"),
            str(args.research),
            "--list-file",
            str(args.list_file),
            "--disease-config",
            str(args.disease_config),
            "--intro-index",
            str(args.intro_index),
            "--intro-dir",
            str(args.intro_dir),
            "--require-complete",
        ],
    )

    if args.template_copy:
        if not args.template.is_file():
            raise SystemExit(f"写作模板不存在：{args.template}")
        if args.template_copy.exists():
            raise SystemExit(f"模板副本已存在，请换新路径：{args.template_copy}")
        args.template_copy.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(args.template, args.template_copy)
        print(f"==> 已复制写作模板副本（源文件未修改）：{args.template_copy}")

    print("==> 运行外部生成器")
    completed = subprocess.run(builder, check=False)
    if completed.returncode:
        return completed.returncode
    if not args.xlsx.is_file() or not args.docx.is_file():
        raise SystemExit("生成器未创建所声明的 Excel 和 Word 交付件。")

    deliverable_gate = [
        python,
        str(SCRIPTS_DIR / "validate_deliverables.py"),
        "--xlsx",
        str(args.xlsx),
        "--docx",
        str(args.docx),
        "--research",
        str(args.research),
        "--list-file",
        str(args.list_file),
        "--disease-config",
        str(args.disease_config),
    ]
    run_gate("P4 交付件结构校验", deliverable_gate)

    if renderer:
        command = [
            str(args.docx) if token == "{docx}" else str(args.render_check) if token == "{render_check}" else token
            for token in renderer
        ]
        if "{docx}" not in renderer and "{render_check}" not in renderer:
            command = command + [str(args.docx), str(args.render_check)]
        print("==> 运行渲染核验器")
        rendered = subprocess.run(command, check=False)
        if rendered.returncode:
            print("渲染核验器执行失败；不得据此声明渲染已通过。", file=sys.stderr)
            return rendered.returncode

    errors = validate_render_check(args.render_check, args.docx)
    if errors:
        print("P5 渲染核验未通过（交付未完成）：", file=sys.stderr)
        print("\n".join(f"- {item}" for item in errors), file=sys.stderr)
        return 1

    print(
        "P0-P5 门禁全部通过：名单、病种配置、输入资源、权威机构台账、检索总账、交付件结构与渲染核验均已确认。"
    )
    print("提醒：结构与痕迹校验不等于疫情事实已核实，事实核实仍须人工完成。")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
