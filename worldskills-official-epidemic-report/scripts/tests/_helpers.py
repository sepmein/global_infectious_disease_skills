"""单元测试用的最小 OOXML 夹具生成器（仅标准库，不触碰用户真实资料）。"""

from __future__ import annotations

import hashlib
import json
import subprocess
import sys
import zipfile
from pathlib import Path
from xml.sax.saxutils import escape

SCRIPTS_DIR = Path(__file__).resolve().parents[1]
if str(SCRIPTS_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPTS_DIR))

PYTHON = sys.executable

CATEGORY_HEADINGS = (
    "（一）蚊媒及其他虫媒传染病",
    "（二）呼吸道传染病",
    "（三）肠道及食源性传染病",
    "（四）新发少见及高致病性传染病",
)
CARD_FIELDS = (
    "1. 疫情规模和截止时间",
    "2. 重症、死亡和医疗负担",
    "3. 主要受影响地区",
    "4. 年龄、职业或其他重点人群",
    "5. 与上年同期及历史基线比较",
    "6. 监测、实验室及病例管理能力",
    "7. 其他重要信息",
    "8. 信息来源",
)


# ------------------------------------------------------------------ xlsx
def _column_name(index: int) -> str:
    name = ""
    index += 1
    while index:
        index, remainder = divmod(index - 1, 26)
        name = chr(65 + remainder) + name
    return name


def write_xlsx(path: Path, sheets: dict[str, list[list[str]]]) -> Path:
    """写出最小可读工作簿（inline string 单元格）。"""
    path.parent.mkdir(parents=True, exist_ok=True)
    sheet_names = list(sheets)
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
        overrides = "".join(
            f'<Override PartName="/xl/worksheets/sheet{index}.xml"'
            ' ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
            for index in range(1, len(sheet_names) + 1)
        )
        archive.writestr(
            "[Content_Types].xml",
            '<?xml version="1.0" encoding="UTF-8"?>'
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            '<Override PartName="/xl/workbook.xml"'
            ' ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
            f"{overrides}</Types>",
        )
        archive.writestr(
            "_rels/.rels",
            '<?xml version="1.0" encoding="UTF-8"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Target="xl/workbook.xml"'
            ' Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"/>'
            "</Relationships>",
        )
        sheet_tags = "".join(
            f'<sheet name="{escape(name)}" sheetId="{index}" r:id="rId{index}"/>'
            for index, name in enumerate(sheet_names, start=1)
        )
        archive.writestr(
            "xl/workbook.xml",
            '<?xml version="1.0" encoding="UTF-8"?>'
            '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
            ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            f"<sheets>{sheet_tags}</sheets></workbook>",
        )
        rels = "".join(
            f'<Relationship Id="rId{index}" Target="worksheets/sheet{index}.xml"'
            ' Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"/>'
            for index in range(1, len(sheet_names) + 1)
        )
        archive.writestr(
            "xl/_rels/workbook.xml.rels",
            '<?xml version="1.0" encoding="UTF-8"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            f"{rels}</Relationships>",
        )
        for index, name in enumerate(sheet_names, start=1):
            rows_xml = []
            for row_index, row in enumerate(sheets[name], start=1):
                cells = "".join(
                    f'<c r="{_column_name(column)}{row_index}" t="inlineStr">'
                    f"<is><t>{escape(str(value))}</t></is></c>"
                    for column, value in enumerate(row)
                    if str(value) != ""
                )
                rows_xml.append(f'<row r="{row_index}">{cells}</row>')
            archive.writestr(
                f"xl/worksheets/sheet{index}.xml",
                '<?xml version="1.0" encoding="UTF-8"?>'
                '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
                f"<sheetData>{''.join(rows_xml)}</sheetData></worksheet>",
            )
    return path


def write_participant_xlsx(path: Path, countries: list[tuple[str, str, str]]) -> Path:
    """写出与真实名单同结构的工作簿：正确表 + 易误读的“名单差异汇总”表。"""
    main = [["代码", "英文名称", "中文名称", "所在名单", "备注"]]
    for code, name_en, name_zh in countries:
        main.append([code, name_en, name_zh, "英文名单、中文名单", "两份名单均有"])
    diff = [
        ["项目", "数量", "", "代码", "英文名称", "中文名称", "所在名单"],
        ["英文名单", str(len(countries)), "", "XX", "Nowhere", "无此国", "仅英文名单"],
    ]
    return write_xlsx(path, {"国家地区合集": main, "名单差异汇总": diff})


def write_participant_csv(path: Path, countries: list[tuple[str, str, str]], *, merged: bool = False) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    if merged:
        lines = ["地区（代码 - 英文名称）,中文标准名称"]
        lines += [f"{code} - {name_en},{name_zh}" for code, name_en, name_zh in countries]
    else:
        lines = ["代码,英文名称,中文名称"]
        lines += [f"{code},{name_en},{name_zh}" for code, name_en, name_zh in countries]
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")
    return path


def write_disease_xlsx(path: Path, diseases: list[tuple[str, str]]) -> Path:
    rows = [["重点病种", "所属分类"]] + [[name, category] for name, category in diseases]
    return write_xlsx(path, {"重点病种检索配置": rows})


def write_disease_csv(path: Path, diseases: list[tuple[str, str]]) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    lines = ["重点病种,所属分类"] + [f"{name},{category}" for name, category in diseases]
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")
    return path


# ------------------------------------------------------------------ docx
_BODY_PPR = (
    '<w:pPr><w:widowControl w:val="1"/>'
    '<w:spacing w:before="0" w:after="0" w:line="540" w:lineRule="exact"/>'
    '<w:ind w:left="0" w:right="0" w:firstLineChars="200"/><w:jc w:val="both"/></w:pPr>'
)
_TITLE_PPR = (
    '<w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'
    '<w:ind w:left="0" w:right="0" w:firstLineChars="0"/><w:jc w:val="center"/></w:pPr>'
)


def _run(text: str, font: str, size: str) -> str:
    return (
        f'<w:r><w:rPr><w:rFonts w:ascii="{font}" w:hAnsi="{font}" w:eastAsia="{font}"/>'
        f'<w:sz w:val="{size}"/><w:szCs w:val="{size}"/></w:rPr>'
        f"<w:t xml:space=\"preserve\">{escape(text)}</w:t></w:r>"
    )


def paragraph(text: str, *, font: str = "仿宋_GB2312", size: str = "32", title: bool = False) -> str:
    return f"<w:p>{_TITLE_PPR if title else _BODY_PPR}{_run(text, font, size)}</w:p>"


def table(header: list[str], rows: list[list[str]]) -> str:
    def cell(text: str) -> str:
        inner = (
            '<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="360" w:lineRule="exact"/>'
            '<w:ind w:left="0" w:right="0" w:firstLineChars="0"/><w:jc w:val="center"/></w:pPr>'
            f'{_run(text, "仿宋_GB2312", "24")}</w:p>'
        )
        return f'<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>{inner}</w:tc>'

    parts = ["<w:tbl>"]
    for row in [header, *rows]:
        parts.append("<w:tr>" + "".join(cell(value) for value in row) + "</w:tr>")
    parts.append("</w:tbl>")
    return "".join(parts)


def write_docx(path: Path, body: list[str]) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    section = (
        '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
        '<w:pgMar w:top="1440" w:right="1587" w:bottom="1440" w:left="1587"'
        ' w:header="851" w:footer="992" w:gutter="0"/></w:sectPr>'
    )
    document = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        f"<w:body>{''.join(body)}{section}</w:body></w:document>"
    )
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
        archive.writestr(
            "[Content_Types].xml",
            '<?xml version="1.0" encoding="UTF-8"?>'
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            '<Override PartName="/word/document.xml"'
            ' ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
            "</Types>",
        )
        archive.writestr(
            "_rels/.rels",
            '<?xml version="1.0" encoding="UTF-8"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Target="word/document.xml"'
            ' Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"/>'
            "</Relationships>",
        )
        archive.writestr("word/document.xml", document)
    return path


def disease_section_paragraphs(
    disease: str,
    *,
    cards: list[tuple[str, str]] | None = None,
    no_event_text: str = "经核实本期名单国家和地区无相关疫情，无入表事件。",
    source_url: str = "https://example.org/weekly-report-2026-35",
) -> list[str]:
    """生成一个病种的四部分正文（结构合规、无占位残留）。"""
    blocks: list[str] = [paragraph(disease)]
    blocks.append(paragraph("一、疾病概述"))
    blocks.append(paragraph("该病由已知病原体引起，传播条件与宿主分布在公开资料中有稳定记载。"))
    blocks.append(paragraph("临床表现以发热为主，重症风险集中于既往有基础疾病的人群。"))
    blocks.append(paragraph(f"信息来源：世界卫生组织疾病主题页，{source_url}，访问日期 2026-09-05。"))
    blocks.append(paragraph("二、全球流行概况"))
    blocks.append(paragraph("本年度全球报告病例较上年同期上升，统计口径以各国常规监测为准。"))
    blocks.append(paragraph("地理分布仍集中于传统流行区，并向邻近区域小幅扩展。"))
    blocks.append(paragraph(f"信息来源：世界卫生组织全球监测数据，{source_url}，数据截至 2026-08-31。"))
    blocks.append(paragraph("三、重点国家流行情况"))
    if cards:
        for country, url in cards:
            blocks.append(paragraph(f"{country}｜{disease}"))
            details = (
                "截至2026年8月31日，该国本年度累计报告病例数与上年同期相比上升。",
                "同期报告重症与死亡病例，医疗服务压力集中在首都地区。",
                "病例集中于三个主要行政区，其余地区零星报告。",
                "报告病例以青壮年为主，官方未公布职业分布。",
                "与上年同期相比病例上升，比较采用同一监测系统与病例定义。",
                "全国实行每周报告制度，国家参考实验室可开展病原学检测。",
                "卫生部门已加强媒介控制与医疗机构分诊。",
                f"{url}（机构周报，数据截至 2026-08-31，访问日期 2026-09-05）",
            )
            for field, detail in zip(CARD_FIELDS, details):
                blocks.append(paragraph(f"{field}：{detail}"))
    else:
        blocks.append(paragraph(no_event_text))
    blocks.append(paragraph("四、风险因素"))
    blocks.append(paragraph("人群免疫水平与媒介分布是影响传播的主要因素。"))
    blocks.append(paragraph("跨境人员流动与监测报告能力差异会放大扩散风险。"))
    blocks.append(paragraph(f"信息来源：区域公共卫生机构风险评估，{source_url}，访问日期 2026-09-05。"))
    return blocks


def valid_report_docx(
    path: Path,
    *,
    participant_count: int,
    focus_rows: list[list[str]],
    sections: dict[str, list[str]],
) -> Path:
    body: list[str] = [
        paragraph("世界技能大赛参赛国家和地区重点传染病疫情报告", font="方正小标宋简体", size="44", title=True),
        paragraph("摸底时间：2026年8月1日—9月4日"),
        paragraph(f"摸底范围：世界技能大赛参赛国家和地区（{participant_count}个）"),
        paragraph("一、重点国家", font="黑体"),
        table(["国家/地区", "重点疾病", "近期疫情情况", "疫情性质"], focus_rows),
        paragraph("二、重点疫情摸底", font="黑体"),
    ]
    for heading in CATEGORY_HEADINGS:
        body.append(paragraph(heading, font="楷体_GB2312"))
        body.extend(sections.get(heading, []))
    return write_docx(path, body)


# ------------------------------------------------------------------ 其他
def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    digest.update(path.read_bytes())
    return digest.hexdigest()


def write_json(path: Path, payload: object) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    return path


def run_script(name: str, *arguments: str) -> subprocess.CompletedProcess[str]:
    return subprocess.run(
        [PYTHON, str(SCRIPTS_DIR / name), *arguments],
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
    )
