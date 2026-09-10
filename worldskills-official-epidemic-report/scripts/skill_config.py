"""共享配置与只读表格读取模块（仅标准库）。

职责：
1. 以文件魔术字节判定真实格式，支持纯文本 CSV 与真实 OOXML ZIP（.csv 后缀实为 xlsx 的情况）。
2. 从 references 下的名单文件与重点病种配置文件动态读取参赛国家/地区和病种及其所属分类。
3. 集中提供交付件结构、公文格式、正文占位黑名单等常量，供各校验脚本共享。

本模块只读文件，绝不写入或修改任何用户资料。
"""

from __future__ import annotations

import csv
import datetime
import hashlib
import io
import re
import zipfile
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field
from pathlib import Path

SKILL_DIR = Path(__file__).resolve().parents[1]
REFERENCES_DIR = SKILL_DIR / "references"
ASSETS_DIR = SKILL_DIR / "assets"

# ---------------------------------------------------------------- 默认固定路径
DEFAULT_PARTICIPANT_LIST = REFERENCES_DIR / "参赛国家和地区名单.csv"
DEFAULT_DISEASE_CONFIG = REFERENCES_DIR / "重点病种检索配置.xlsx"
DEFAULT_INTRO_DIR = REFERENCES_DIR / "疾病简介"
DEFAULT_INTRO_INDEX = REFERENCES_DIR / "疾病简介索引.json"
DEFAULT_LOCAL_MATERIAL_DIR = SKILL_DIR / "本地疫情资料"
DEFAULT_TEMPLATE_DOCX = ASSETS_DIR / "世界技能大赛传染病疫情报告_写作模板.docx"

# ---------------------------------------------------------------- 固定分类与标题
CATEGORY_ORDER: tuple[str, ...] = (
    "蚊媒及其他虫媒传染病",
    "呼吸道传染病",
    "肠道及食源性传染病",
    "新发少见及高致病性传染病",
)
CATEGORY_HEADING_PREFIX: tuple[str, ...] = ("（一）", "（二）", "（三）", "（四）")
CATEGORY_HEADINGS: tuple[str, ...] = tuple(
    prefix + name for prefix, name in zip(CATEGORY_HEADING_PREFIX, CATEGORY_ORDER)
)
TOP_HEADINGS: tuple[str, ...] = ("一、重点国家", "二、重点疫情摸底")

# 每个病种下固定四部分标题，顺序与文字必须完全一致
DISEASE_PART_TITLES: tuple[str, ...] = (
    "一、疾病概述",
    "二、全球流行概况",
    "三、重点国家流行情况",
    "四、风险因素",
)
PART_KEYS: tuple[str, ...] = ("overview", "global", "countries", "risk")
PART_KEY_BY_TITLE: dict[str, str] = dict(zip(DISEASE_PART_TITLES, PART_KEYS))
# 需要 2-3 段正文的部分（三、重点国家流行情况按国家卡片单独校验）
NARRATIVE_PART_KEYS: tuple[str, ...] = ("overview", "global", "risk")
NARRATIVE_MIN_PARAGRAPHS = 2
NARRATIVE_MAX_PARAGRAPHS = 3

# 重点国家“小卡片”八项固定字段（按段首匹配）
COUNTRY_CARD_FIELDS: tuple[str, ...] = (
    "1. 疫情规模和截止时间",
    "2. 重症、死亡和医疗负担",
    "3. 主要受影响地区",
    "4. 年龄、职业或其他重点人群",
    "5. 与上年同期及历史基线比较",
    "6. 监测、实验室及病例管理能力",
    "7. 其他重要信息",
    "8. 信息来源",
)

# ---------------------------------------------------------------- 交付件表头
MAIN_SHEET_NAME = "疫情摸底底表"
WEB_EVIDENCE_SHEET_NAME = "网络证据清单"
LOCAL_EVIDENCE_SHEET_NAME = "本地证据清单"
COUNTRY_LOG_SHEET_NAME = "逐国检索记录"

MAIN_HEADERS: tuple[str, ...] = (
    "国家/地区",
    "重点疾病",
    "近期疫情情况",
    "高发/暴发/异常上升/新发",
    "病例（包括疑似、确诊）",
    "死亡数",
    "最近更新时间",
    "信息来源",
    "建议世赛关注程度",
)
WEB_EVIDENCE_HEADERS: tuple[str, ...] = (
    "来源类型",
    "疾病",
    "国家/地区",
    "日期",
    "来源文档",
    "来源 URL",
    "访问日期",
    "核实状态",
)
LOCAL_EVIDENCE_HEADERS: tuple[str, ...] = (
    "证据编号",
    "疾病",
    "国家/地区",
    "文件路径",
    "定位",
    "数据截止日期",
    "核实状态",
)
FOCUS_TABLE_HEADERS: tuple[str, ...] = ("国家/地区", "重点疾病", "近期疫情情况", "疫情性质")
FOCUS_TABLE_MAX_ROWS = 10
LOCAL_SOURCE_PREFIX = "LOCAL:"

# ---------------------------------------------------------------- 公文格式常量
TITLE_FONT = "方正小标宋简体"
TITLE_SIZE_HALF_POINT = "44"  # 22 磅
BODY_FONT = "仿宋_GB2312"
BODY_SIZE_HALF_POINT = "32"  # 16 磅
HEADING_FONTS: frozenset[str] = frozenset({"黑体", "楷体_GB2312", BODY_FONT})
TABLE_FONT_PREFIX = "仿宋"
TABLE_SIZE_HALF_POINTS: frozenset[str] = frozenset({"24", "32"})  # 12 磅（允许），16 磅
BODY_LINE_TWIPS = "540"  # 27 磅固定行距
BODY_LINE_RULE = "exact"
BODY_FIRST_LINE_CHARS = "200"  # 首行缩进 2 字符
BODY_FIRST_LINE_TWIPS = "640"  # 16 磅 * 2 字符
BODY_JUSTIFY = "both"
PAGE_WIDTH_TWIPS = 11906  # A4
PAGE_HEIGHT_TWIPS = 16838
MARGIN_LEFT_RIGHT_TWIPS = 1587  # 28 mm
MARGIN_TOP_BOTTOM_TWIPS = 1440  # 25.4 mm
TWIPS_TOLERANCE = 20

# ---------------------------------------------------------------- 模板占位残留
# 正式交付 Word 中一旦出现下列片段，即视为模板占位/示例/写作指导未清理。
PLACEHOLDER_FRAGMENTS: tuple[str, ...] = (
    "具体的病种名称",
    "××",
    "示例",
    "所有数字和情境均为虚构",
    "本分类下每个病种均独立填写",
    "本部分回答",
    "用两至三段",
    "逐条列出",
    "按实际纳入数量逐一重复",
    "国家/地区名称",
    "每个重点国家和地区采用下列统一",
    "并非真实文献",
    "格式占位",
    "〈",
    "${",
    "占位",
    "同上",
)

# ---------------------------------------------------------------- 枚举
SOURCE_TYPES: frozenset[str] = frozenset({"local", "web"})
# 来源权威性类型；intro 缺失回退时强制要求 official。其余来源可选，不强制。
AUTHORITY_TYPES: frozenset[str] = frozenset(
    {"official", "media", "international", "academic", "other"}
)
READ_STATUSES: frozenset[str] = frozenset({"已读取", "读取失败", "未读取"})
VERIFICATION_STATUSES: frozenset[str] = frozenset({"人工核实一致", "核实一致", "人工核实存疑", "未人工核实"})
# 视为“已核实”的声明集合（智能体实际核对 / 人工核实一致，二者兼容）
VERIFIED_SET: frozenset[str] = frozenset({"人工核实一致", "核实一致"})
GLOBAL_SCOPE = "GLOBAL"

DISEASE_CHECK_STATUSES: frozenset[str] = frozenset({"已完成并纳入", "已完成未纳入", "检索未完成"})
DISEASE_CHECK_COMPLETED: frozenset[str] = frozenset({"已完成并纳入", "已完成未纳入"})
RESEARCH_STATUSES: frozenset[str] = frozenset({"已完成", "检索未完成"})
SEARCH_STATUSES: frozenset[str] = frozenset({"已完成", "未完成"})

DATE_PATTERN = re.compile(r"^\d{4}-\d{2}-\d{2}$")

HUMAN_VERIFICATION_DISCLAIMER = (
    "声明：本脚本只校验结构、覆盖度与文件痕迹（含真实哈希），"
    "不能替代人工事实核实；JSON 中的核实字段仅是记录的声明，不构成已核实的证明。"
)


class ConfigError(Exception):
    """配置或表格读取失败。"""


# ================================================================ 通用工具
def sha256_file(path: Path) -> str:
    """以只读方式计算文件 SHA-256。"""
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def is_ooxml_zip(path: Path) -> bool:
    """按魔术字节判定文件是否为真实 OOXML ZIP（后缀不可信）。"""
    try:
        with path.open("rb") as handle:
            if handle.read(4)[:2] != b"PK":
                return False
    except OSError as exc:  # pragma: no cover - 交由调用方报错
        raise ConfigError(f"无法读取文件：{path}（{exc}）") from exc
    try:
        with zipfile.ZipFile(path) as archive:
            names = set(archive.namelist())
    except zipfile.BadZipFile:
        return False
    return "xl/workbook.xml" in names


def _clean(value: str | None) -> str:
    if value is None:
        return ""
    return value.replace("\u00a0", " ").replace("\u3000", " ").strip()


def _column_index(cell_ref: str) -> int:
    """把 'AB12' 形式的单元格引用转换为 0 基列号。"""
    index = 0
    for char in cell_ref:
        if not char.isalpha():
            break
        index = index * 26 + (ord(char.upper()) - 64)
    return max(index - 1, 0)


# ================================================================ 表格读取
_MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def _shared_strings(archive: zipfile.ZipFile) -> list[str]:
    if "xl/sharedStrings.xml" not in archive.namelist():
        return []
    root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
    values: list[str] = []
    for item in root.findall(f"{{{_MAIN_NS}}}si"):
        values.append("".join(node.text or "" for node in item.iter(f"{{{_MAIN_NS}}}t")))
    return values


def _resolve_part(archive: zipfile.ZipFile, target: str) -> str | None:
    names = archive.namelist()
    for candidate in (target, f"xl/{target.lstrip('/')}", target.lstrip("/")):
        if candidate in names:
            return candidate
    return None


def _sheet_rows(archive: zipfile.ZipFile, part: str, shared: list[str]) -> list[list[str]]:
    root = ET.fromstring(archive.read(part))
    rows: list[list[str]] = []
    for row in root.iter(f"{{{_MAIN_NS}}}row"):
        cells: list[str] = []
        for cell in row.findall(f"{{{_MAIN_NS}}}c"):
            index = _column_index(cell.get("r") or "")
            cell_type = cell.get("t")
            value_node = cell.find(f"{{{_MAIN_NS}}}v")
            inline_node = cell.find(f"{{{_MAIN_NS}}}is")
            if cell_type == "s" and value_node is not None and value_node.text is not None:
                try:
                    text = shared[int(value_node.text)]
                except (ValueError, IndexError):
                    text = ""
            elif cell_type == "inlineStr" and inline_node is not None:
                text = "".join(node.text or "" for node in inline_node.iter(f"{{{_MAIN_NS}}}t"))
            else:
                text = value_node.text if value_node is not None else ""
            while len(cells) <= index:
                cells.append("")
            cells[index] = _clean(text)
        rows.append(cells)
    return rows


def read_workbook_sheets(path: Path) -> dict[str, list[list[str]]]:
    """读取 OOXML 工作簿的全部工作表，保持工作簿内顺序。"""
    with zipfile.ZipFile(path) as archive:
        shared = _shared_strings(archive)
        workbook = ET.fromstring(archive.read("xl/workbook.xml"))
        rels = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
        rel_map = {node.get("Id"): node.get("Target") or "" for node in rels}
        sheets: dict[str, list[list[str]]] = {}
        container = workbook.find(f"{{{_MAIN_NS}}}sheets")
        if container is None:
            raise ConfigError(f"工作簿缺少工作表定义：{path}")
        for sheet in container:
            name = _clean(sheet.get("name"))
            rel_id = sheet.get(f"{{{_REL_NS}}}id")
            target = rel_map.get(rel_id or "")
            part = _resolve_part(archive, target) if target else None
            if part is None:
                raise ConfigError(f"工作簿工作表“{name}”的数据部件缺失：{path}")
            sheets[name] = _sheet_rows(archive, part, shared)
        return sheets


def read_csv_rows(path: Path) -> list[list[str]]:
    """读取纯文本 CSV（UTF-8/带 BOM）。"""
    raw = path.read_bytes()
    for encoding in ("utf-8-sig", "utf-8", "gbk"):
        try:
            text = raw.decode(encoding)
        except UnicodeDecodeError:
            continue
        return [[_clean(cell) for cell in row] for row in csv.reader(io.StringIO(text, newline=""))]
    raise ConfigError(f"无法按 UTF-8/GBK 解码 CSV：{path}")


@dataclass(frozen=True)
class Schema:
    """一种可接受的表头结构。anchor 必须是表头行第一个非空单元格。"""

    name: str
    anchor: str
    required: tuple[str, ...]


@dataclass
class Table:
    """选定的表格及其表头定位结果。"""

    path: Path
    sheet_name: str | None
    schema: Schema
    header: list[str]
    header_row_number: int
    rows: list[list[str]] = field(default_factory=list)

    def column(self, name: str) -> int:
        return self.header.index(name)

    def value(self, row: list[str], name: str) -> str:
        index = self.column(name)
        return row[index] if index < len(row) else ""


# 名称中含下列字样的工作表不作为数据源（避免误读“名单差异汇总”等辅助表）
_EXCLUDED_SHEET_MARKERS: tuple[str, ...] = ("差异", "汇总", "说明", "统计", "变更", "备注", "对照")
_MAX_HEADER_SCAN_ROWS = 5


def _find_header(rows: list[list[str]], schemas: tuple[Schema, ...]) -> tuple[Schema, int] | None:
    for row_index, row in enumerate(rows[:_MAX_HEADER_SCAN_ROWS]):
        cells = [_clean(cell) for cell in row]
        first = next((cell for cell in cells if cell), "")
        for schema in schemas:
            if first != schema.anchor:
                continue
            if all(column in cells for column in schema.required):
                return schema, row_index
    return None


def load_table(
    path: Path,
    schemas: tuple[Schema, ...],
    *,
    preferred_sheets: tuple[str, ...] = (),
    sheet: str | None = None,
) -> Table:
    """按真实格式读取表格，并在多工作表时精准选定正确的配置表。"""
    if not path.exists():
        raise ConfigError(f"未找到文件：{path}")
    if not path.is_file():
        raise ConfigError(f"路径不是文件：{path}")

    expected = "；".join(f"{schema.name}={'/'.join(schema.required)}" for schema in schemas)

    if not is_ooxml_zip(path):
        rows = read_csv_rows(path)
        found = _find_header(rows, schemas)
        if found is None:
            raise ConfigError(f"CSV 表头不符合要求（期望之一：{expected}）：{path}")
        schema, header_index = found
        header = [_clean(cell) for cell in rows[header_index]]
        return Table(
            path=path,
            sheet_name=None,
            schema=schema,
            header=header,
            header_row_number=header_index + 1,
            rows=rows[header_index + 1 :],
        )

    sheets = read_workbook_sheets(path)
    if sheet is not None:
        if sheet not in sheets:
            raise ConfigError(
                f"工作簿中不存在指定工作表“{sheet}”：{path}；可选工作表：{'、'.join(sheets)}"
            )
        candidates = {sheet: sheets[sheet]}
    else:
        candidates = sheets

    matched: list[tuple[str, Schema, int]] = []
    for name, rows in candidates.items():
        found = _find_header(rows, schemas)
        if found is None:
            continue
        if sheet is None and any(marker in name for marker in _EXCLUDED_SHEET_MARKERS):
            continue
        matched.append((name, found[0], found[1]))

    if not matched:
        raise ConfigError(
            f"工作簿中没有符合要求的工作表（期望之一：{expected}）：{path}；"
            f"实际工作表：{'、'.join(sheets)}"
        )

    chosen: tuple[str, Schema, int] | None = None
    if sheet is None and preferred_sheets:
        for preferred in preferred_sheets:
            for item in matched:
                if item[0] == preferred:
                    chosen = item
                    break
            if chosen is not None:
                break
    if chosen is None:
        if len(matched) > 1:
            raise ConfigError(
                f"工作簿中有多个候选工作表：{'、'.join(item[0] for item in matched)}；"
                f"请用 --sheet 明确指定：{path}"
            )
        chosen = matched[0]

    name, schema, header_index = chosen
    rows = sheets[name]
    header = [_clean(cell) for cell in rows[header_index]]
    return Table(
        path=path,
        sheet_name=name,
        schema=schema,
        header=header,
        header_row_number=header_index + 1,
        rows=rows[header_index + 1 :],
    )


# ================================================================ 参赛名单
PARTICIPANT_SCHEMA_SPLIT = Schema(
    name="三列格式",
    anchor="代码",
    required=("代码", "英文名称", "中文名称"),
)
PARTICIPANT_SCHEMA_MERGED = Schema(
    name="旧合并两列格式",
    anchor="地区（代码 - 英文名称）",
    required=("地区（代码 - 英文名称）", "中文标准名称"),
)
PARTICIPANT_SCHEMAS = (PARTICIPANT_SCHEMA_SPLIT, PARTICIPANT_SCHEMA_MERGED)
PARTICIPANT_PREFERRED_SHEETS = ("国家地区合集", "参赛国家和地区名单", "名单")


@dataclass(frozen=True)
class Country:
    code: str
    name_en: str
    name_zh: str


@dataclass
class ParticipantList:
    path: Path
    sheet_name: str | None
    schema_name: str
    countries: tuple[Country, ...]

    @property
    def count(self) -> int:
        return len(self.countries)

    @property
    def codes(self) -> set[str]:
        return {country.code for country in self.countries}

    @property
    def by_code(self) -> dict[str, Country]:
        return {country.code: country for country in self.countries}


def load_participants(
    path: Path | None = None, *, sheet: str | None = None
) -> ParticipantList:
    """动态读取参赛国家和地区名单，禁止空名称与重复代码。"""
    target = Path(path) if path else DEFAULT_PARTICIPANT_LIST
    table = load_table(
        target,
        PARTICIPANT_SCHEMAS,
        preferred_sheets=PARTICIPANT_PREFERRED_SHEETS,
        sheet=sheet,
    )
    errors: list[str] = []
    countries: list[Country] = []
    seen_codes: set[str] = set()
    seen_names: set[tuple[str, str]] = set()
    merged = table.schema is PARTICIPANT_SCHEMA_MERGED

    for offset, row in enumerate(table.rows):
        line = table.header_row_number + offset + 1
        if not any(_clean(cell) for cell in row):
            continue
        if merged:
            merged_value = table.value(row, "地区（代码 - 英文名称）")
            name_zh = table.value(row, "中文标准名称")
            parts = [part.strip() for part in merged_value.split(" - ", maxsplit=1)]
            if len(parts) != 2 or not all(parts):
                errors.append(f"第 {line} 行“地区（代码 - 英文名称）”格式应为“代码 - 英文名称”：{merged_value!r}")
                continue
            code, name_en = parts
        else:
            code = table.value(row, "代码")
            name_en = table.value(row, "英文名称")
            name_zh = table.value(row, "中文名称")

        if not code:
            errors.append(f"第 {line} 行缺少地区代码")
            continue
        if not name_en:
            errors.append(f"第 {line} 行（{code}）缺少英文名称")
        if not name_zh:
            errors.append(f"第 {line} 行（{code}）缺少中文名称")
        if code in seen_codes:
            errors.append(f"第 {line} 行地区代码重复：{code}")
        seen_codes.add(code)
        key = (name_zh, name_en)
        if key in seen_names:
            errors.append(f"第 {line} 行国家或地区名称重复：{name_zh} / {name_en}")
        seen_names.add(key)
        if name_en and name_zh:
            countries.append(Country(code=code, name_en=name_en, name_zh=name_zh))

    if not countries and not errors:
        errors.append(f"名单中没有可用的国家或地区：{target}")
    if errors:
        raise ConfigError("\n".join(errors))
    return ParticipantList(
        path=table.path,
        sheet_name=table.sheet_name,
        schema_name=table.schema.name,
        countries=tuple(countries),
    )


# ================================================================ 重点病种配置
DISEASE_SCHEMA = Schema(name="病种与分类", anchor="重点病种", required=("重点病种", "所属分类"))
DISEASE_SCHEMAS = (DISEASE_SCHEMA,)
DISEASE_PREFERRED_SHEETS = ("重点病种检索配置", "重点病种", "病种配置")


@dataclass(frozen=True)
class Disease:
    name: str
    category: str


@dataclass
class DiseaseConfig:
    path: Path
    sheet_name: str | None
    diseases: tuple[Disease, ...]

    @property
    def names(self) -> tuple[str, ...]:
        return tuple(disease.name for disease in self.diseases)

    @property
    def name_set(self) -> set[str]:
        return set(self.names)

    @property
    def category_by_disease(self) -> dict[str, str]:
        return {disease.name: disease.category for disease in self.diseases}

    @property
    def by_category(self) -> dict[str, tuple[str, ...]]:
        grouped: dict[str, list[str]] = {category: [] for category in CATEGORY_ORDER}
        for disease in self.diseases:
            grouped[disease.category].append(disease.name)
        return {category: tuple(names) for category, names in grouped.items()}

    @property
    def count(self) -> int:
        return len(self.diseases)


def load_diseases(path: Path | None = None, *, sheet: str | None = None) -> DiseaseConfig:
    """动态读取重点病种及其所属分类；分类必须来自固定四类。"""
    target = Path(path) if path else DEFAULT_DISEASE_CONFIG
    table = load_table(
        target,
        DISEASE_SCHEMAS,
        preferred_sheets=DISEASE_PREFERRED_SHEETS,
        sheet=sheet,
    )
    errors: list[str] = []
    diseases: list[Disease] = []
    seen: set[str] = set()
    for offset, row in enumerate(table.rows):
        line = table.header_row_number + offset + 1
        if not any(_clean(cell) for cell in row):
            continue
        name = table.value(row, "重点病种")
        category = table.value(row, "所属分类")
        if not name:
            errors.append(f"第 {line} 行病种名称为空")
            continue
        if name in seen:
            errors.append(f"第 {line} 行病种重复：{name}")
        seen.add(name)
        if not category:
            errors.append(f"第 {line} 行病种“{name}”缺少所属分类")
            continue
        if category not in CATEGORY_ORDER:
            errors.append(
                f"第 {line} 行病种“{name}”的分类“{category}”不在固定四类中：{'、'.join(CATEGORY_ORDER)}"
            )
            continue
        diseases.append(Disease(name=name, category=category))

    if not diseases and not errors:
        errors.append(f"重点病种配置为空：{target}")
    if errors:
        raise ConfigError("\n".join(errors))
    return DiseaseConfig(
        path=table.path,
        sheet_name=table.sheet_name,
        diseases=tuple(diseases),
    )


# ================================================================ 疾病简介索引
@dataclass
class IntroIndex:
    """疾病简介索引：病种 -> 相对文件名或 None（简介缺失）。"""

    path: Path
    intro_dir: Path
    mapping: dict[str, str | None]
    warnings: tuple[str, ...] = ()

    @property
    def missing_diseases(self) -> tuple[str, ...]:
        return tuple(name for name, value in self.mapping.items() if not value)

    def resolved_path(self, disease: str) -> Path | None:
        filename = self.mapping.get(disease)
        if not filename:
            return None
        return self.intro_dir / filename


def load_intro_index(
    path: Path | None = None,
    *,
    intro_dir: Path | None = None,
    diseases: DiseaseConfig | None = None,
) -> IntroIndex:
    """读取疾病简介索引；键集合必须与病种配置严格一致。"""
    import json

    target = Path(path) if path else DEFAULT_INTRO_INDEX
    directory = Path(intro_dir) if intro_dir else DEFAULT_INTRO_DIR
    if not target.is_file():
        raise ConfigError(
            f"未找到疾病简介索引：{target}；请由主代理提供 {{病种: 相对文件名 或 null}} 的 JSON 映射"
        )
    try:
        raw = json.loads(target.read_text(encoding="utf-8"))
    except (OSError, ValueError) as exc:
        raise ConfigError(f"疾病简介索引无法解析为 JSON：{target}（{exc}）") from exc
    if not isinstance(raw, dict):
        raise ConfigError(f"疾病简介索引顶层必须是对象：{target}")

    mapping: dict[str, str | None] = {}
    errors: list[str] = []
    warnings: list[str] = []
    for key, value in raw.items():
        if not isinstance(key, str) or not key.strip():
            errors.append("索引存在空病种键")
            continue
        if value is None:
            mapping[key.strip()] = None
            continue
        if not isinstance(value, str):
            errors.append(f"病种“{key}”的索引值必须是字符串或 null")
            continue
        cleaned = value.strip()
        if not cleaned:
            mapping[key.strip()] = None
            continue
        # 安全约束：只能是简介目录内的相对文件名，禁止绝对路径与 ../ 穿越
        if Path(cleaned).is_absolute():
            errors.append(
                f"病种“{key}”的简介文件名必须是简介目录内的相对文件名，不得为绝对路径：{cleaned}"
            )
            continue
        if ".." in Path(cleaned).parts:
            errors.append(f"病种“{key}”的简介文件名不得包含路径穿越（..）：{cleaned}")
            continue
        if not (directory / cleaned).is_file() and directory.exists():
            # 文件存在性错误在 validate_input_resources / P2 阶段再统一报告，
            # 这里仅做结构安全校验，不阻断尚未提供文件的场景。
            pass
        mapping[key.strip()] = cleaned

    if diseases is not None:
        expected = set(diseases.names)
        actual = set(mapping)
        missing = sorted(expected - actual)
        extra = sorted(actual - expected)
        # 键缺失（新增病种）：自动映射为缺失（None）并告警，允许以权威网络资料替代，
        # 但不得伪称已读本地简介；不阻断 P0/P2，由下游回退分支处理。
        for name in missing:
            mapping[name] = None
            warnings.append(
                f"疾病简介索引缺少病种“{name}”，自动映射为缺失（null）："
                "本地简介缺失，将以权威网络资料替代，不得伪称已读本地简介或伪造修复"
            )
        # 配置外病种（额外键）：不得扩大范围，保留原拒绝。
        if extra:
            errors.append(f"疾病简介索引存在配置外病种：{'、'.join(extra)}")
    if errors:
        raise ConfigError("\n".join(errors))
    return IntroIndex(path=target, intro_dir=directory, mapping=mapping, warnings=tuple(warnings))


# ================================================================ 校验辅助
def is_valid_date(value: str) -> bool:
    """用真实日历验证 YYYY-MM-DD（含闰年、月份天数等）。"""
    if not isinstance(value, str) or not DATE_PATTERN.match(value):
        return False
    year, month, day = (int(part) for part in value.split("-"))
    try:
        datetime.date(year, month, day)
    except ValueError:
        return False
    return 1900 <= year <= 2999


def local_material_hint() -> str:
    """本地资料缺失时给出的固定建议路径提示。"""
    return (
        "请确认本地资料是否放在以下建议路径（脚本只读，不会修改）：\n"
        f"  - 本地疫情资料根目录：{DEFAULT_LOCAL_MATERIAL_DIR}\n"
        f"  - 疾病简介目录：{DEFAULT_INTRO_DIR}"
    )
