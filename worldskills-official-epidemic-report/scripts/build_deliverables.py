"""Run the only permitted report-build sequence after all hard gates pass."""

from __future__ import annotations

import argparse
import subprocess
import sys
from pathlib import Path


SKILL_DIR = Path(__file__).resolve().parents[1]
SCRIPTS_DIR = SKILL_DIR / "scripts"
DEFAULT_LIST = SKILL_DIR / "references" / "参赛国家和地区名单.csv"
DEFAULT_DISEASE_CONFIG = SKILL_DIR / "references" / "重点病种检索配置.csv"


def run_gate(label: str, command: list[str]) -> None:
    completed = subprocess.run(command, check=False)
    if completed.returncode:
        raise SystemExit(f"{label}未通过；禁止生成或修改交付件。")


def main() -> int:
    parser = argparse.ArgumentParser(
        description="先执行 P0-P2 门禁，随后运行生成器，再执行 P4 结构校验。"
    )
    parser.add_argument("--registry", type=Path, required=True)
    parser.add_argument("--research", type=Path, required=True)
    parser.add_argument("--xlsx", type=Path, required=True)
    parser.add_argument("--docx", type=Path, required=True)
    parser.add_argument("--disease-config", type=Path, default=DEFAULT_DISEASE_CONFIG)
    parser.add_argument(
        "--builder",
        nargs=argparse.REMAINDER,
        required=True,
        help="在 --builder 后提供生成器命令；该命令只能在所有门禁通过后执行。",
    )
    args = parser.parse_args()
    if not args.builder:
        parser.error("--builder 后必须提供生成器命令")
    if args.xlsx.exists() or args.docx.exists():
        raise SystemExit("目标交付件已存在；为避免绕过门禁编辑现有文件，请使用新的版本文件名。")

    python = sys.executable
    run_gate("P0 名单校验", [python, str(SCRIPTS_DIR / "validate_participant_list.py")])
    run_gate("P0 重点病种配置校验", [python, str(SCRIPTS_DIR / "validate_priority_disease_config.py"), str(args.disease_config)])
    run_gate("P1 权威机构台账校验", [python, str(SCRIPTS_DIR / "validate_authority_registry.py"), str(args.registry)])
    run_gate(
        "P2 逐国检索校验",
        [python, str(SCRIPTS_DIR / "validate_country_research.py"), str(args.research), "--require-complete"],
    )

    completed = subprocess.run(args.builder, check=False)
    if completed.returncode:
        return completed.returncode
    if not args.xlsx.is_file() or not args.docx.is_file():
        raise SystemExit("生成器未创建所声明的 Excel 和 Word 交付件。")
    run_gate(
        "P4 交付件结构校验",
        [python, str(SCRIPTS_DIR / "validate_deliverables.py"), "--xlsx", str(args.xlsx), "--docx", str(args.docx)],
    )
    print("P0-P4 门禁均已通过；Word 渲染检查仍须在本脚本返回后完成。")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
