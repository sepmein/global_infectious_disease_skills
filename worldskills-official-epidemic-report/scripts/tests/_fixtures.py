"""构造一整套自洽的小型夹具：名单、病种、简介索引、本地资料、台账、检索总账、交付件。

所有夹具都写在测试临时目录内，不读取也不修改仓库中的真实资料。
"""

from __future__ import annotations

import copy
import json
from dataclasses import dataclass, field
from pathlib import Path

from _helpers import (
    CATEGORY_HEADINGS,
    disease_section_paragraphs,
    sha256,
    valid_report_docx,
    write_disease_xlsx,
    write_json,
    write_participant_xlsx,
    write_xlsx,
)

COUNTRIES = [("TH", "Thailand", "泰国"), ("JP", "Japan", "日本")]
DISEASES = [("登革热", "蚊媒及其他虫媒传染病"), ("麻疹", "呼吸道传染病")]
MAIN_HEADERS = [
    "国家/地区",
    "重点疾病",
    "近期疫情情况",
    "高发/暴发/异常上升/新发",
    "病例（包括疑似、确诊）",
    "死亡数",
    "最近更新时间",
    "信息来源",
    "建议世赛关注程度",
]
WEB_HEADERS = ["来源类型", "疾病", "国家/地区", "日期", "来源文档", "来源 URL", "访问日期", "核实状态"]
LOCAL_HEADERS = ["证据编号", "疾病", "国家/地区", "文件路径", "定位", "数据截止日期", "核实状态"]
REGISTRY_FIELDS = [
    "地区代码",
    "中文标准名称",
    "英文名称",
    "发现检索词",
    "候选页面 URL",
    "机构名称",
    "机构类型",
    "机构权威性核实依据 URL",
    "规范入口 URL",
    "直接访问状态",
    "直接访问最终 URL",
    "浏览器回退状态",
    "浏览器最终 URL",
    "区域交叉机构",
    "区域交叉 URL",
    "访问日期",
    "台账结论",
]
DENGUE_URL = "https://example.org/th/dengue-week35"
DENGUE_GLOBAL_URL = "https://example.org/global/dengue-2026"
MEASLES_URL = "https://example.org/jp/measles-week35"
MEASLES_GLOBAL_URL = "https://example.org/global/measles-2026"
MEASLES_TH_URL = "https://example.org/th/measles-week35"


@dataclass
class Env:
    root: Path
    list_file: Path
    disease_config: Path
    intro_dir: Path
    intro_index: Path
    local_dir: Path
    intro_dengue: Path
    local_jp_dengue: Path
    local_jp_measles: Path
    registry_json: Path = field(init=False)
    research_json: Path = field(init=False)
    xlsx: Path = field(init=False)
    docx: Path = field(init=False)


def build_env(root: Path) -> Env:
    root.mkdir(parents=True, exist_ok=True)
    references = root / "references"
    intro_dir = references / "疾病简介"
    local_dir = root / "本地疫情资料"
    intro_dir.mkdir(parents=True, exist_ok=True)
    local_dir.mkdir(parents=True, exist_ok=True)

    list_file = write_participant_xlsx(references / "参赛国家和地区名单.csv", COUNTRIES)
    disease_config = write_disease_xlsx(references / "重点病种检索配置.xlsx", DISEASES)

    intro_dengue = intro_dir / "登革热和重症登革热.txt"
    intro_dengue.write_text("登革热简介：病原学、传播途径与既往全球分布。", encoding="utf-8")
    # 麻疹简介故意缺失，模拟真实缺失（不得伪造修复）
    intro_index = write_json(
        references / "疾病简介索引.json",
        {"登革热": intro_dengue.name, "麻疹": None},
    )

    local_jp_dengue = local_dir / "日本登革热国家监测周报.txt"
    local_jp_dengue.write_text("日本登革热监测周报：本期无本地传播病例。", encoding="utf-8")
    local_jp_measles = local_dir / "日本麻疹病例通报.txt"
    local_jp_measles.write_text("日本麻疹病例通报：截至2026年8月31日累计报告病例。", encoding="utf-8")

    env = Env(
        root=root,
        list_file=list_file,
        disease_config=disease_config,
        intro_dir=intro_dir,
        intro_index=intro_index,
        local_dir=local_dir,
        intro_dengue=intro_dengue,
        local_jp_dengue=local_jp_dengue,
        local_jp_measles=local_jp_measles,
    )
    env.registry_json = root / "registry.json"
    env.research_json = root / "research.json"
    env.xlsx = root / "deliverable.xlsx"
    env.docx = root / "deliverable.docx"
    return env


def build_complete_env(root: Path) -> "Env":
    """与 build_env 类似，但所有病种都提供本地简介文件（无 null），用于完整交付门禁的happy path。"""
    env = build_env(root)
    measles_intro = env.intro_dir / "麻疹简介.txt"
    measles_intro.write_text("麻疹简介：病原学、临床特征与全球分布概况。", encoding="utf-8")
    index = json.loads(env.intro_index.read_text(encoding="utf-8"))
    index["麻疹"] = measles_intro.name
    env.intro_index.write_text(json.dumps(index, ensure_ascii=False), encoding="utf-8")
    return env


def research_payload_complete(env: "Env") -> dict:
    """在 research_payload 基础上，将缺失的麻疹简介补齐为本地文件引用，使 require_complete 可整体通过。"""
    payload = research_payload(env)
    measles_intro = env.intro_dir / "麻疹简介.txt"
    payload["sources"].append(
        {
            "id": "L4",
            "type": "local",
            "path": str(measles_intro),
            "source_origin_url": "https://example.org/global/measles-factsheet",
            "locator": "第1节 病原学",
            "read_status": "已读取",
            "verification": "核实一致",
            "country_code": "GLOBAL",
            "disease": "麻疹",
            "data_date": "未公布",
            "source_date": "未公布",
            "access_date": "2026-09-05",
        }
    )
    payload["local_inputs"]["files"].append(
        {
            "path": str(measles_intro),
            "sha256": sha256(measles_intro),
            "read_status": "已读取",
            "locator": "第1节 病原学",
        }
    )
    for section in payload["disease_sections"]:
        if section["disease"] == "麻疹":
            section["intro_file"] = measles_intro.name
            section["intro_read"] = True
            section["intro_locator"] = "第1节 病原学"
            section.pop("intro_fallback_reason", None)
            section["sections"]["overview"]["evidence_ids"] = ["L4"]
            section["sections"]["overview"]["local_sufficient"] = True
            section["sections"]["overview"]["internet_searches"] = []
    return payload


# ------------------------------------------------------------------ 台账
def registry_payload() -> dict:
    records = []
    for index, (code, name_en, name_zh) in enumerate(COUNTRIES):
        limited = index == 1
        records.append(
            {
                "地区代码": code,
                "中文标准名称": name_zh,
                "英文名称": name_en,
                "发现检索词": f"{name_en} ministry of health surveillance weekly",
                "候选页面 URL": f"https://example.org/{code.lower()}/candidate",
                "机构名称": f"{name_en} National Institute of Public Health",
                "机构类型": "国家CDC/法定监测机构",
                "机构权威性核实依据 URL": f"https://example.org/{code.lower()}/about",
                "规范入口 URL": f"https://example.org/{code.lower()}/surveillance",
                "直接访问状态": "连接超时（15s，重试2次）" if limited else "HTTP 200",
                "直接访问最终 URL": "无" if limited else f"https://example.org/{code.lower()}/surveillance",
                "浏览器回退状态": "浏览器加载失败：需要人机验证" if limited else "不适用：直接访问成功",
                "浏览器最终 URL": "无" if limited else "不适用",
                "区域交叉机构": "无",
                "区域交叉 URL": "无",
                "访问日期": "2026-09-05",
                "台账结论": "已验证但访问受限" if limited else "已验证可检索",
            }
        )
    return {"records": records}


def write_registry(env: Env, payload: dict | None = None) -> Path:
    return write_json(env.registry_json, payload or registry_payload())


# ------------------------------------------------------------------ 检索总账
def _search(query: str, source_ids: list[str]) -> dict:
    return {
        "query": query,
        "tool": "websearch",
        "search_date": "2026-09-05",
        "result_source_ids": source_ids,
        "status": "已完成",
    }


def research_payload(env: Env) -> dict:
    sources = [
        {
            "id": "L1",
            "type": "local",
            "path": str(env.intro_dengue),
            "locator": "第1节 病原学",
            "read_status": "已读取",
            "verification": "人工核实一致",
            "country_code": "GLOBAL",
            "disease": "登革热",
            "data_date": "未公布",
            "source_date": "未公布",
            "access_date": "2026-09-05",
        },
        {
            "id": "L2",
            "type": "local",
            "path": str(env.local_jp_dengue),
            "locator": "第2页 表1",
            "read_status": "已读取",
            "verification": "人工核实一致",
            "country_code": "JP",
            "disease": "登革热",
            "data_date": "2026-08-31",
            "source_date": "2026-09-03",
            "access_date": "2026-09-05",
        },
        {
            "id": "L3",
            "type": "local",
            "path": str(env.local_jp_measles),
            "locator": "第1页",
            "read_status": "已读取",
            "verification": "人工核实一致",
            "country_code": "JP",
            "disease": "麻疹",
            "data_date": "2026-08-31",
            "source_date": "2026-09-02",
            "access_date": "2026-09-05",
        },
        {
            "id": "W1",
            "type": "web",
            "url": DENGUE_URL,
            "locator": "第35周表2",
            "read_status": "已读取",
            "verification": "人工核实一致",
            "country_code": "TH",
            "disease": "登革热",
            "data_date": "2026-08-31",
            "source_date": "2026-09-02",
            "access_date": "2026-09-05",
        },
        {
            "id": "W2",
            "type": "web",
            "url": DENGUE_GLOBAL_URL,
            "locator": "全球汇总表",
            "read_status": "已读取",
            "verification": "人工核实一致",
            "country_code": "GLOBAL",
            "disease": "登革热",
            "data_date": "2026-08-31",
            "source_date": "2026-09-01",
            "access_date": "2026-09-05",
        },
        {
            "id": "W3",
            "type": "web",
            "url": MEASLES_GLOBAL_URL,
            "locator": "全球麻疹监测页",
            "read_status": "已读取",
            "verification": "人工核实一致",
            "country_code": "GLOBAL",
            "disease": "麻疹",
            "data_date": "2026-08-31",
            "source_date": "2026-09-01",
            "access_date": "2026-09-05",
            "authority_type": "official",
            "authority_basis": "WHO 麻疹疾病主题页（官方权威机构，用于全球疾病概述）",
        },
        {
            "id": "W4",
            "type": "web",
            "url": MEASLES_URL,
            "locator": "第35周麻疹病例表",
            "read_status": "已读取",
            "verification": "人工核实一致",
            "country_code": "JP",
            "disease": "麻疹",
            "data_date": "2026-08-31",
            "source_date": "2026-09-02",
            "access_date": "2026-09-05",
        },
        {
            "id": "W5",
            "type": "web",
            "url": MEASLES_TH_URL,
            "locator": "月度传染病统计",
            "read_status": "已读取",
            "verification": "人工核实一致",
            "country_code": "TH",
            "disease": "麻疹",
            "data_date": "2026-08-31",
            "source_date": "2026-09-02",
            "access_date": "2026-09-05",
        },
    ]
    local_inputs = {
        "requested": True,
        "paths": [str(env.local_dir), str(env.intro_dir)],
        "files": [
            {
                "path": str(env.intro_dengue),
                "sha256": sha256(env.intro_dengue),
                "read_status": "已读取",
                "locator": "第1节 病原学",
            },
            {
                "path": str(env.local_jp_dengue),
                "sha256": sha256(env.local_jp_dengue),
                "read_status": "已读取",
                "locator": "第2页 表1",
            },
            {
                "path": str(env.local_jp_measles),
                "sha256": sha256(env.local_jp_measles),
                "read_status": "已读取",
                "locator": "第1页",
            },
        ],
    }
    disease_sections = [
        {
            "disease": "登革热",
            "category": "蚊媒及其他虫媒传染病",
            "intro_file": env.intro_dengue.name,
            "intro_read": True,
            "intro_locator": "第1节 病原学",
            "sections": {
                "overview": {
                    "sufficient": True,
                    "local_sufficient": True,
                    "evidence_ids": ["L1"],
                    "internet_searches": [],
                },
                "global": {
                    "sufficient": True,
                    "local_sufficient": False,
                    "evidence_ids": ["W2"],
                    "internet_searches": [_search("dengue global situation 2026", ["W2"])],
                },
                "countries": {
                    "sufficient": True,
                    "local_sufficient": False,
                    "evidence_ids": ["W1", "L2"],
                    "internet_searches": [_search("Thailand dengue weekly report", ["W1"])],
                },
                "risk": {
                    "sufficient": True,
                    "local_sufficient": False,
                    "evidence_ids": ["W2"],
                    "internet_searches": [_search("dengue risk factors 2026", ["W2"])],
                },
            },
        },
        {
            "disease": "麻疹",
            "category": "呼吸道传染病",
            "intro_file": "",
            "intro_read": False,
            "intro_locator": "",
            "intro_fallback_reason": (
                "疾病简介索引为 null，本地简介缺失；已用 WHO 麻疹疾病主题页（authority_type=official）"
                "替代，并真实访问登记；未伪称已读本地简介。"
            ),
            "sections": {
                "overview": {
                    "sufficient": True,
                    "local_sufficient": False,
                    "evidence_ids": ["W3"],
                    "internet_searches": [_search("measles disease overview WHO", ["W3"])],
                },
                "global": {
                    "sufficient": True,
                    "local_sufficient": False,
                    "evidence_ids": ["W3"],
                    "internet_searches": [_search("measles global cases 2026", ["W3"])],
                },
                "countries": {
                    "sufficient": True,
                    "local_sufficient": False,
                    "evidence_ids": ["W4", "W5"],
                    "internet_searches": [_search("Japan measles weekly report", ["W4"])],
                },
                "risk": {
                    "sufficient": True,
                    "local_sufficient": False,
                    "evidence_ids": ["W3"],
                    "internet_searches": [_search("measles immunity gap risk", ["W3"])],
                },
            },
        },
    ]
    records = [
        {
            "country_code": "TH",
            "country_name_zh": "泰国",
            "research_status": "已完成",
            "disease_checks": [
                {
                    "disease": "登革热",
                    "category": "蚊媒及其他虫媒传染病",
                    "status": "已完成并纳入",
                    "reason": "报告病例较上年同期上升",
                    "evidence_ids": ["W1"],
                    "internet_required": True,
                    "internet_searches": [_search("Thailand dengue weekly report", ["W1"])],
                },
                {
                    "disease": "麻疹",
                    "category": "呼吸道传染病",
                    "status": "已完成未纳入",
                    "reason": "经核实本期国家级监测未报告异常上升",
                    "evidence_ids": ["W5"],
                    "internet_required": True,
                    "internet_searches": [_search("Thailand measles surveillance 2026", ["W5"])],
                },
            ],
        },
        {
            "country_code": "JP",
            "country_name_zh": "日本",
            "research_status": "已完成",
            "disease_checks": [
                {
                    "disease": "登革热",
                    "category": "蚊媒及其他虫媒传染病",
                    "status": "已完成未纳入",
                    "reason": "用户提供的国家监测周报显示本期无本地传播病例",
                    "evidence_ids": ["L2"],
                    "internet_required": False,
                    "internet_searches": [],
                },
                {
                    "disease": "麻疹",
                    "category": "呼吸道传染病",
                    "status": "已完成并纳入",
                    "reason": "本年度报告病例高于上年同期",
                    "evidence_ids": ["W4", "L3"],
                    "internet_required": True,
                    "internet_searches": [_search("Japan measles weekly report", ["W4"])],
                },
            ],
        },
    ]
    return {
        "schema_version": "2.0",
        "monitoring_period": {"start": "2026-08-01", "end": "2026-09-04"},
        "sources": sources,
        "local_inputs": local_inputs,
        "disease_sections": disease_sections,
        "records": records,
    }


def write_research(env: Env, mutate=None, *, path: Path | None = None) -> Path:
    payload = research_payload(env)
    if mutate is not None:
        mutate(payload)
    return write_json(path or env.research_json, payload)


def clone_payload(env: Env) -> dict:
    return copy.deepcopy(research_payload(env))


# ------------------------------------------------------------------ 交付件
def workbook_sheets(env: Env, *, main_rows=None, log_codes=None, web_rows=None, local_rows=None) -> dict:
    main = main_rows if main_rows is not None else [
        [
            "TH - Thailand 泰国",
            "登革热",
            "全国报告病例较上年同期上升",
            "异常上升",
            "累计报告 12 例（含疑似）",
            "0",
            "2026-08-31",
            DENGUE_URL,
            "高",
        ],
        [
            "JP - Japan 日本",
            "麻疹",
            "本年度报告病例高于上年同期",
            "异常上升",
            "累计报告 8 例（确诊）",
            "0",
            "2026-08-31",
            f"{MEASLES_URL}\nLOCAL:LE1",
            "中",
        ],
    ]
    web = web_rows if web_rows is not None else [
        ["国家级监测", "登革热", "TH - Thailand 泰国", "2026-09-02", "第35周监测周报", DENGUE_URL, "2026-09-05", "人工核实一致"],
        ["国家级监测", "麻疹", "JP - Japan 日本", "2026-09-02", "第35周麻疹病例表", MEASLES_URL, "2026-09-05", "人工核实一致"],
    ]
    local = local_rows if local_rows is not None else [
        ["LE1", "麻疹", "JP - Japan 日本", str(env.local_jp_measles), "第1页", "2026-08-31", "人工核实一致"],
        ["LE2", "登革热", "JP - Japan 日本", str(env.local_jp_dengue), "第2页 表1", "2026-08-31", "人工核实一致"],
    ]
    codes = log_codes if log_codes is not None else [code for code, _, _ in COUNTRIES]
    log = [["地区代码", "中文标准名称", "检索状态"]]
    zh_by_code = {code: zh for code, _, zh in COUNTRIES}
    for code in codes:
        log.append([code, zh_by_code.get(code, code), "已完成"])
    return {
        "疫情摸底底表": [MAIN_HEADERS, *main],
        "网络证据清单": [WEB_HEADERS, *web],
        "本地证据清单": [LOCAL_HEADERS, *local],
        "逐国检索记录": log,
    }


def write_workbook(env: Env, *, path: Path | None = None, **kwargs) -> Path:
    return write_xlsx(path or env.xlsx, workbook_sheets(env, **kwargs))


def report_sections(*, dengue_cards=None, measles_cards=None) -> dict[str, list[str]]:
    dengue = disease_section_paragraphs(
        "登革热",
        cards=dengue_cards if dengue_cards is not None else [("泰国", DENGUE_URL)],
        source_url=DENGUE_GLOBAL_URL,
    )
    measles = disease_section_paragraphs(
        "麻疹",
        cards=measles_cards if measles_cards is not None else [("日本", MEASLES_URL)],
        source_url=MEASLES_GLOBAL_URL,
    )
    return {CATEGORY_HEADINGS[0]: dengue, CATEGORY_HEADINGS[1]: measles}


def write_report(env: Env, *, path: Path | None = None, sections=None, focus_rows=None) -> Path:
    return valid_report_docx(
        path or env.docx,
        participant_count=len(COUNTRIES),
        focus_rows=focus_rows
        if focus_rows is not None
        else [
            ["TH - Thailand 泰国", "登革热", "全国报告病例较上年同期上升", "异常上升"],
            ["JP - Japan 日本", "麻疹", "本年度报告病例高于上年同期", "异常上升"],
        ],
        sections=sections if sections is not None else report_sections(),
    )
