"""P4：校验交付件结构（Word 正文骨架与公文格式 + Excel 底表与证据清单）。

Word 要求（与旧版 5 张小表流程不同）：
- 只有一张“重点国家”表，最多 10 行数据；
- 二级为固定四个分类标题，各分类下按配置顺序逐病种独立成节；
- 每个病种：病种名称单独成段，随后严格顺序出现四个部分标题；
- 一/二/四部分各 2-3 段正文（来源段单独列出、不计入）；
- 第三部分为重点国家小卡片，每张卡片按段首出现 8 个固定字段；确无入表疫情时允许用一段说明代替，不得编造卡片；
- 正文不得残留模板占位、示例数字、写作指导等文字；
- 公文格式：标题 22 磅方正小标宋简体，正文 16 磅仿宋_GB2312，固定行距 27 磅，首行缩进 2 字符，
  两端对齐，段前段后 0，A4 页面，左右 28mm、上下 25.4mm；表格允许 12 磅仿宋。

Word 的所有可见来源只允许实际核实的 HTTP(S) 链接；禁止展示本地文件名、路径、页码定位、哈希、证据编号或 LOCAL: 标识。

Excel 要求：主表 9 列不变，网络证据清单 8 列，新增本地证据清单 7 列，逐国检索记录须按动态名单逐一覆盖。
主表“信息来源”可填 URL 或 LOCAL:<证据编号>，多条以换行分隔，并须能在对应证据清单中找到同疾病同国家的证据；Excel 的本地证据仅供内部审计。

本脚本只做结构与可追溯性校验：不判断疫情事实是否正确，也不代表 Word 已渲染验证（渲染须由 build 流程的 --render-check 完成）。
"""

from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path

from docx_reader import Paragraph, WordDocument, read_docx
from skill_config import (
    BODY_FIRST_LINE_CHARS,
    BODY_FIRST_LINE_TWIPS,
    BODY_FONT,
    BODY_JUSTIFY,
    BODY_LINE_RULE,
    BODY_LINE_TWIPS,
    BODY_SIZE_HALF_POINT,
    CATEGORY_HEADINGS,
    CATEGORY_ORDER,
    COUNTRY_CARD_FIELDS,
    COUNTRY_LOG_SHEET_NAME,
    ConfigError,
    DEFAULT_DISEASE_CONFIG,
    DEFAULT_PARTICIPANT_LIST,
    DISEASE_PART_TITLES,
    FOCUS_TABLE_HEADERS,
    FOCUS_TABLE_MAX_ROWS,
    HEADING_FONTS,
    LOCAL_EVIDENCE_HEADERS,
    LOCAL_EVIDENCE_SHEET_NAME,
    LOCAL_SOURCE_PREFIX,
    MAIN_HEADERS,
    MAIN_SHEET_NAME,
    MARGIN_LEFT_RIGHT_TWIPS,
    MARGIN_TOP_BOTTOM_TWIPS,
    NARRATIVE_MAX_PARAGRAPHS,
    NARRATIVE_MIN_PARAGRAPHS,
    PAGE_HEIGHT_TWIPS,
    PAGE_WIDTH_TWIPS,
    PLACEHOLDER_FRAGMENTS,
    TABLE_FONT_PREFIX,
    TABLE_SIZE_HALF_POINTS,
    TITLE_FONT,
    TITLE_SIZE_HALF_POINT,
    TWIPS_TOLERANCE,
    VERIFIED_SET,
    WEB_EVIDENCE_HEADERS,
    WEB_EVIDENCE_SHEET_NAME,
    load_diseases,
    load_participants,
    read_workbook_sheets,
)

SOURCE_LINE_PREFIXES = ("信息来源", "检索来源", "来源", "参考来源", "URL", "数据来源")
SOURCE_LINE_PATTERN = re.compile(r"^\[\d+\]")
HTTP_URL_PATTERN = re.compile(r"https?://[^\s，。；;、（）()【】\[\]<>]+", re.IGNORECASE)
LOCAL_SOURCE_LEAK_PATTERN = re.compile(
    r"(?:"
    r"[A-Za-z]:[\\/]|"  # Windows absolute path
    r"(?:^|[\s（(【\[])(?:\\\\|//)[^\s，。；;、）)】\]]+|"  # UNC / network path
    r"(?:^|[\s（(【\[])(?:\.\.?[\\/]|~[\\/])[^\s，。；;、）)】\]]+|"  # relative/home path
    r"LOCAL\s*:|SHA[-_ ]?256|"
    r"(?:本地文件|本地路径|本地证据|本地资料|本地材料|本地简介|上传文件|上传附件|用户提供的文件|附件路径|文件绝对路径|文件相对路径)|"
    r"[^\s，。；;、）)】\]>]+\.(?:pdf|docx?|xlsx?|csv|tsv|json|txt|md|html?|png|jpe?g|gif|bmp|tiff?|zip)(?=$|[\s，。；;、）)】\]>)])"
    r")",
    re.IGNORECASE,
)
NO_EVENT_HINTS = ("无相关疫情", "未纳入", "本期无", "未检索到可核实", "无入表")
MAX_REPORTED = 20


def _is_source_line(text: str) -> bool:
    return text.startswith(SOURCE_LINE_PREFIXES) or bool(SOURCE_LINE_PATTERN.match(text))


def _source_urls(text: str) -> list[str]:
    return HTTP_URL_PATTERN.findall(text)


def _has_source_reference(text: str) -> bool:
    return bool(_source_urls(text))


def _without_http_urls(text: str) -> str:
    return HTTP_URL_PATTERN.sub("", text)


def _cap(errors: list[str], label: str, items: list[str]) -> None:
    if not items:
        return
    shown = items[:MAX_REPORTED]
    suffix = "" if len(items) <= MAX_REPORTED else f"（另有 {len(items) - MAX_REPORTED} 处同类问题）"
    errors.append(f"{label}共 {len(items)} 处：" + "；".join(shown) + suffix)


# ------------------------------------------------------------------ Word 结构
def _segment(paragraphs: list[Paragraph], start: int, stops: set[str]) -> tuple[list[Paragraph], int]:
    body: list[Paragraph] = []
    index = start
    while index < len(paragraphs):
        text = paragraphs[index].text
        if text in stops:
            break
        body.append(paragraphs[index])
        index += 1
    return body, index


def validate_word_structure(document: WordDocument, diseases, participant_count: int) -> list[str]:
    errors: list[str] = []
    paragraphs = [item for item in document.body_paragraphs if item.text]
    texts = [item.text for item in paragraphs]

    if not paragraphs:
        return ["Word 正文没有任何段落"]

    title = paragraphs[0]
    if TITLE_FONT not in title.fonts:
        errors.append(f"Word 标题字体应为{TITLE_FONT}，实际：{'、'.join(sorted(title.fonts)) or '未设置'}")
    if TITLE_SIZE_HALF_POINT not in title.sizes:
        errors.append(f"Word 标题字号应为 22 磅（sz={TITLE_SIZE_HALF_POINT}），实际：{'、'.join(sorted(title.sizes)) or '未设置'}")

    for heading in ("一、重点国家", "二、重点疫情摸底"):
        if heading not in texts:
            errors.append(f"Word 缺少章节标题：{heading}")
    scope = next((text for text in texts if text.startswith("摸底范围")), "")
    if not scope:
        errors.append("Word 缺少“摸底范围”说明段")
    elif f"{participant_count}个" not in scope:
        errors.append(f"Word 摸底范围未使用动态名单数 {participant_count}：{scope}")

    if len(document.tables) != 1:
        errors.append(f"Word 应只有一张重点国家表，实际有 {len(document.tables)} 张表")
    if document.tables:
        table = document.tables[0]
        if table.header != list(FOCUS_TABLE_HEADERS):
            errors.append(f"Word 重点国家表列名或顺序不符：{table.header}")
        data_rows = table.row_count - 1
        if data_rows > FOCUS_TABLE_MAX_ROWS:
            errors.append(f"Word 重点国家表数据行 {data_rows} 行，超过 {FOCUS_TABLE_MAX_ROWS} 个上限")
        if data_rows < 1:
            errors.append("Word 重点国家表没有数据行")

    # 分类标题顺序
    positions: dict[str, int] = {}
    for heading in CATEGORY_HEADINGS:
        if heading not in texts:
            errors.append(f"Word 缺少分类标题：{heading}")
        else:
            positions[heading] = texts.index(heading)
    ordered = [positions[heading] for heading in CATEGORY_HEADINGS if heading in positions]
    if ordered != sorted(ordered):
        errors.append("Word 四个分类标题顺序与固定顺序不一致")

    by_category = diseases.by_category
    disease_names = set(diseases.names)
    stops = set(CATEGORY_HEADINGS) | set(DISEASE_PART_TITLES) | disease_names | {"一、重点国家", "二、重点疫情摸底"}

    for category_index, heading in enumerate(CATEGORY_HEADINGS):
        if heading not in positions:
            continue
        category = CATEGORY_ORDER[category_index]
        start = positions[heading] + 1
        end = len(paragraphs)
        for later in CATEGORY_HEADINGS[category_index + 1 :]:
            if later in positions:
                end = positions[later]
                break
        segment = paragraphs[start:end]
        segment_texts = [item.text for item in segment]
        expected = list(by_category.get(category, ()))
        found = [text for text in segment_texts if text in disease_names]
        if found != expected:
            errors.append(
                f"分类“{category}”下病种及顺序应为 {expected or '（配置内无病种）'}，实际为 {found or '（无）'}"
            )
            continue

        for disease in expected:
            offset = segment_texts.index(disease)
            cursor = offset + 1
            for part_index, part_title in enumerate(DISEASE_PART_TITLES):
                if cursor >= len(segment) or segment_texts[cursor] != part_title:
                    actual = segment_texts[cursor] if cursor < len(segment) else "（文档结束）"
                    errors.append(
                        f"病种“{disease}”第 {part_index + 1} 部分标题应为“{part_title}”，实际为“{actual}”"
                    )
                    break
                body, cursor = _segment(segment, cursor + 1, stops)
                narrative = [item for item in body if not _is_source_line(item.text)]
                source_lines = [item for item in body if _is_source_line(item.text)]
                if part_title == "三、重点国家流行情况":
                    errors.extend(_validate_cards(disease, body))
                    continue
                if not NARRATIVE_MIN_PARAGRAPHS <= len(narrative) <= NARRATIVE_MAX_PARAGRAPHS:
                    errors.append(
                        f"病种“{disease}”“{part_title}”应有 {NARRATIVE_MIN_PARAGRAPHS}-{NARRATIVE_MAX_PARAGRAPHS} 段正文"
                        f"（来源段不计），实际 {len(narrative)} 段"
                    )
                if not source_lines:
                    errors.append(f"病种“{disease}”“{part_title}”缺少信息来源段")
                elif not any(_has_source_reference(item.text) for item in source_lines):
                    errors.append(
                        f"病种“{disease}”“{part_title}”的来源段缺少HTTP(S)网络链接"
                    )
    return errors


def _validate_cards(disease: str, body: list[Paragraph]) -> list[str]:
    errors: list[str] = []
    texts = [item.text for item in body]
    starts = [index for index, text in enumerate(texts) if text.startswith(COUNTRY_CARD_FIELDS[0])]
    if not starts:
        explanation = [text for text in texts if not _is_source_line(text)]
        if len(explanation) != 1:
            errors.append(
                f"病种“{disease}”第三部分没有国家卡片时，应只用一段明确说明本期无入表疫情，实际 {len(explanation)} 段"
            )
        elif not any(hint in explanation[0] for hint in NO_EVENT_HINTS):
            errors.append(
                f"病种“{disease}”第三部分未出现国家卡片，且说明段未明确“本期无相关疫情/未纳入”：{explanation[0][:40]}"
            )
        return errors

    for card_index, start in enumerate(starts, start=1):
        end = starts[card_index] if card_index < len(starts) else len(texts)
        card = texts[start:end]
        cursor = 0
        for field in COUNTRY_CARD_FIELDS:
            matched = -1
            for offset in range(cursor, len(card)):
                if card[offset].startswith(field):
                    matched = offset
                    break
            if matched < 0:
                errors.append(f"病种“{disease}”第 {card_index} 张国家卡片缺少字段或顺序错误：{field}")
                break
            cursor = matched + 1
        else:
            source_block = card[cursor - 1 :]
            if not any(_has_source_reference(text) for text in source_block):
                errors.append(
                    f"病种“{disease}”第 {card_index} 张国家卡片“8. 信息来源”缺少HTTP(S)网络链接"
                )
        previous = texts[start - 1] if start > 0 else ""
        if not previous or previous.startswith(COUNTRY_CARD_FIELDS):
            errors.append(f"病种“{disease}”第 {card_index} 张国家卡片缺少国家/地区与病种的卡片标题段")
    return errors


def validate_word_placeholders(document: WordDocument) -> list[str]:
    hits: list[str] = []
    for paragraph in document.paragraphs:
        for fragment in PLACEHOLDER_FRAGMENTS:
            if fragment in paragraph.text:
                hits.append(f"第 {paragraph.order + 1} 段含“{fragment}”")
                break
    errors: list[str] = []
    _cap(errors, "Word 残留模板占位/示例/写作指导文字", hits)
    return errors


def validate_word_public_sources(
    document: WordDocument,
    verified_web_urls: set[str] | None = None,
) -> list[str]:
    """阻止本地文件来源泄露，并要求可见来源只使用已核实HTTP(S)链接。"""
    leaks: list[str] = []
    invalid_sources: list[str] = []
    unverified_urls: list[str] = []
    for paragraph in document.paragraphs:
        text = paragraph.text.strip()
        if not text:
            continue
        urls = _source_urls(text)
        scrubbed = _without_http_urls(text)
        match = LOCAL_SOURCE_LEAK_PATTERN.search(scrubbed)
        if match:
            token = match.group(0).strip()
            leaks.append(f"第 {paragraph.order + 1} 段含本地来源信息“{token[:80]}”")
        if _is_source_line(text) and not urls:
            invalid_sources.append(f"第 {paragraph.order + 1} 段未列HTTP(S)网络链接")
        if verified_web_urls is not None:
            for url in urls:
                if url not in verified_web_urls:
                    unverified_urls.append(f"第 {paragraph.order + 1} 段链接未在已核实web来源中登记：{url}")

    errors: list[str] = []
    _cap(errors, "Word 禁止展示本地文件来源", leaks)
    _cap(errors, "Word 来源段只允许HTTP(S)网络链接", invalid_sources)
    _cap(errors, "Word 仅可展示research总账中的已核实web链接", unverified_urls)
    return errors


def load_verified_web_urls(research_json: Path) -> tuple[set[str], list[str]]:
    try:
        payload = json.loads(research_json.read_text(encoding="utf-8"))
    except (OSError, ValueError) as exc:
        return set(), [f"无法读取检索总账用于核对Word网络来源：{research_json}（{exc}）"]
    urls: set[str] = set()
    for source in payload.get("sources", []):
        if not isinstance(source, dict):
            continue
        if str(source.get("type", "")).strip() != "web":
            continue
        if str(source.get("read_status", "")).strip() != "已读取":
            continue
        if str(source.get("verification", "")).strip() not in VERIFIED_SET:
            continue
        url = str(source.get("url", "")).strip()
        if url.startswith(("http://", "https://")):
            urls.add(url)
    return urls, []


def validate_word_format(document: WordDocument, diseases) -> list[str]:
    errors: list[str] = []
    headings = (
        set(CATEGORY_HEADINGS)
        | set(DISEASE_PART_TITLES)
        | set(diseases.names)
        | {"一、重点国家", "二、重点疫情摸底"}
    )
    font_issues: list[str] = []
    size_issues: list[str] = []
    layout_issues: list[str] = []
    table_issues: list[str] = []
    body = [item for item in document.body_paragraphs if item.text]

    for paragraph in document.paragraphs:
        if not paragraph.text:
            continue
        label = f"第 {paragraph.order + 1} 段"
        if paragraph.in_table:
            for font in paragraph.fonts:
                if not font.startswith(TABLE_FONT_PREFIX):
                    table_issues.append(f"{label}表格字体 {font or '未设置'}")
            for size in paragraph.sizes:
                if size not in TABLE_SIZE_HALF_POINTS:
                    table_issues.append(f"{label}表格字号 {size or '未设置'}")
            continue
        if body and paragraph is body[0]:
            continue  # 标题单独校验
        for size in paragraph.sizes:
            if size != BODY_SIZE_HALF_POINT:
                size_issues.append(f"{label}字号 {size or '未设置'}")
        allowed = HEADING_FONTS if paragraph.text in headings else {BODY_FONT}
        for font in paragraph.fonts:
            if font not in allowed:
                font_issues.append(f"{label}字体 {font or '未设置'}")
        if paragraph.line != BODY_LINE_TWIPS or paragraph.line_rule != BODY_LINE_RULE:
            layout_issues.append(f"{label}行距 {paragraph.line or '未设置'}/{paragraph.line_rule or '未设置'}")
        if paragraph.space_before not in {None, "0"} or paragraph.space_after not in {None, "0"}:
            layout_issues.append(f"{label}段前/段后 {paragraph.space_before}/{paragraph.space_after}")
        if paragraph.first_line_chars != BODY_FIRST_LINE_CHARS and paragraph.first_line != BODY_FIRST_LINE_TWIPS:
            layout_issues.append(f"{label}首行缩进 {paragraph.first_line_chars or paragraph.first_line or '未设置'}")
        if paragraph.justify != BODY_JUSTIFY:
            layout_issues.append(f"{label}对齐 {paragraph.justify or '未设置'}")

    _cap(errors, f"正文字体应为{BODY_FONT}（标题类允许黑体/楷体_GB2312）", font_issues)
    _cap(errors, "正文字号应为 16 磅（sz=32）", size_issues)
    _cap(errors, "段落格式应为固定行距 27 磅、首行缩进 2 字符、两端对齐、段前段后 0", layout_issues)
    _cap(errors, f"表格应使用{TABLE_FONT_PREFIX}系字体、12 磅或 16 磅", table_issues)

    section = document.section
    if section is None:
        errors.append("Word 缺少页面设置（sectPr）")
        return errors

    def near(actual: int | None, expected: int, name: str) -> None:
        if actual is None or abs(actual - expected) > TWIPS_TOLERANCE:
            errors.append(f"页面{name}应为 {expected} twips，实际 {actual}")

    near(section.page_width, PAGE_WIDTH_TWIPS, "宽度（A4）")
    near(section.page_height, PAGE_HEIGHT_TWIPS, "高度（A4）")
    near(section.margin_left, MARGIN_LEFT_RIGHT_TWIPS, "左边距（28mm）")
    near(section.margin_right, MARGIN_LEFT_RIGHT_TWIPS, "右边距（28mm）")
    near(section.margin_top, MARGIN_TOP_BOTTOM_TWIPS, "上边距（25.4mm）")
    near(section.margin_bottom, MARGIN_TOP_BOTTOM_TWIPS, "下边距（25.4mm）")
    return errors


# ------------------------------------------------------------------ Excel
def _rows(sheet: list[list[str]]) -> tuple[list[str], list[list[str]]]:
    if not sheet:
        return [], []
    return sheet[0], [row for row in sheet[1:] if any(cell.strip() for cell in row)]


def _cell(row: list[str], index: int) -> str:
    return row[index].strip() if index < len(row) else ""


def _match_country(text: str, participants) -> str | None:
    for country in participants.countries:
        if re.search(rf"(?<![A-Za-z]){re.escape(country.code)}(?![A-Za-z])", text):
            return country.code
    lowered = text.lower()
    for country in participants.countries:
        if country.name_zh and country.name_zh in text:
            return country.code
        if country.name_en and country.name_en.lower() in lowered:
            return country.code
    return None


def validate_workbook(
    path: Path,
    participants,
    diseases,
) -> tuple[list[str], set[tuple[str, str]]]:
    errors: list[str] = []
    events: set[tuple[str, str]] = set()
    sheets = read_workbook_sheets(path)
    for name in (MAIN_SHEET_NAME, WEB_EVIDENCE_SHEET_NAME, LOCAL_EVIDENCE_SHEET_NAME, COUNTRY_LOG_SHEET_NAME):
        if name not in sheets:
            errors.append(f"Excel 缺少工作表：{name}")
    if errors:
        return errors, events

    main_header, main_rows = _rows(sheets[MAIN_SHEET_NAME])
    web_header, web_rows = _rows(sheets[WEB_EVIDENCE_SHEET_NAME])
    local_header, local_rows = _rows(sheets[LOCAL_EVIDENCE_SHEET_NAME])
    log_header, log_rows = _rows(sheets[COUNTRY_LOG_SHEET_NAME])

    if main_header[: len(MAIN_HEADERS)] != list(MAIN_HEADERS) or len(main_header) != len(MAIN_HEADERS):
        errors.append(f"主表列名或顺序不符合要求（应为 {len(MAIN_HEADERS)} 列）：{main_header}")
    if web_header[: len(WEB_EVIDENCE_HEADERS)] != list(WEB_EVIDENCE_HEADERS):
        errors.append(f"网络证据清单列名或顺序不符合要求：{web_header}")
    if local_header[: len(LOCAL_EVIDENCE_HEADERS)] != list(LOCAL_EVIDENCE_HEADERS):
        errors.append(f"本地证据清单列名或顺序不符合要求：{local_header}")
    if errors:
        return errors, events

    web_by_url: dict[str, list[tuple[str, str]]] = {}
    for index, row in enumerate(web_rows, start=2):
        url = _cell(row, 5)
        if not url:
            errors.append(f"网络证据清单第 {index} 行缺少来源 URL")
            continue
        if not url.startswith(("http://", "https://")):
            errors.append(f"网络证据清单第 {index} 行来源 URL 不是 http(s) 链接：{url}")
        web_by_url.setdefault(url, []).append((_cell(row, 1), _cell(row, 2)))

    local_by_id: dict[str, tuple[str, str, str]] = {}
    for index, row in enumerate(local_rows, start=2):
        evidence_id = _cell(row, 0)
        file_path = _cell(row, 3)
        if not evidence_id:
            errors.append(f"本地证据清单第 {index} 行缺少证据编号")
            continue
        if evidence_id in local_by_id:
            errors.append(f"本地证据清单第 {index} 行证据编号重复：{evidence_id}")
        if not file_path:
            errors.append(f"本地证据清单第 {index} 行缺少文件路径")
        elif not Path(file_path).is_file():
            errors.append(f"本地证据清单第 {index} 行文件不存在：{file_path}")
        if not _cell(row, 4):
            errors.append(f"本地证据清单第 {index} 行缺少定位")
        local_by_id[evidence_id] = (_cell(row, 1), _cell(row, 2), file_path)

    configured = diseases.name_set
    unconfigured: list[str] = []
    for index, row in enumerate(main_rows, start=2):
        country_cell = _cell(row, 0)
        disease = _cell(row, 1)
        sources_cell = row[7] if len(row) > 7 else ""
        code = _match_country(country_cell, participants)
        if code is None:
            errors.append(f"主表第 {index} 行国家/地区不在动态名单内：{country_cell or '（空）'}")
        if not disease:
            errors.append(f"主表第 {index} 行缺少重点疾病")
        elif disease not in configured:
            unconfigured.append(f"第 {index} 行“{disease}”")
        if code and disease:
            events.add((code, disease))

        items = [item.strip() for item in str(sources_cell).replace("\r", "\n").split("\n") if item.strip()]
        if not items:
            errors.append(f"主表第 {index} 行缺少信息来源（URL 或 {LOCAL_SOURCE_PREFIX}<证据编号>）")
            continue
        matched_evidence = False
        for item in items:
            if item.startswith(LOCAL_SOURCE_PREFIX):
                evidence_id = item[len(LOCAL_SOURCE_PREFIX) :].strip()
                entry = local_by_id.get(evidence_id)
                if entry is None:
                    errors.append(f"主表第 {index} 行引用的本地证据编号不存在：{evidence_id}")
                    continue
                if entry[0] == disease and _match_country(entry[1], participants) == code:
                    matched_evidence = True
            elif item.startswith(("http://", "https://")):
                entries = web_by_url.get(item)
                if entries is None:
                    errors.append(f"主表第 {index} 行来源 URL 未写入网络证据清单：{item}")
                    continue
                if any(
                    entry[0] == disease and _match_country(entry[1], participants) == code
                    for entry in entries
                ):
                    matched_evidence = True
            else:
                errors.append(
                    f"主表第 {index} 行信息来源既不是 http(s) 链接也不是 {LOCAL_SOURCE_PREFIX}<证据编号>：{item}"
                )
        if not matched_evidence:
            errors.append(
                f"主表第 {index} 行（{country_cell}｜{disease}）没有同疾病同国家的证据记录"
            )

    if unconfigured:
        errors.append(
            "主表存在配置外病种："
            + "、".join(unconfigured[:MAX_REPORTED])
            + "；本流程不允许表外病种，须先更新重点病种检索配置"
        )

    code_column = next(
        (index for index, name in enumerate(log_header) if name.strip() in {"地区代码", "国家/地区代码", "代码"}),
        None,
    )
    if code_column is None:
        errors.append(f"逐国检索记录缺少“地区代码”列：{log_header}")
    else:
        logged = {_cell(row, code_column) for row in log_rows if _cell(row, code_column)}
        missing = sorted(participants.codes - logged)
        extra = sorted(logged - participants.codes)
        if missing:
            errors.append(f"逐国检索记录缺少 {len(missing)} 个名单国家/地区：{'、'.join(missing[:MAX_REPORTED])}")
        if extra:
            errors.append(f"逐国检索记录存在名单外代码：{'、'.join(extra[:MAX_REPORTED])}")
        if len(log_rows) != len(logged):
            errors.append(f"逐国检索记录有 {len(log_rows)} 行但只有 {len(logged)} 个唯一代码，存在重复行")
    return errors, events


def cross_check_research(research_json: Path, events: set[tuple[str, str]]) -> list[str]:
    errors: list[str] = []
    try:
        payload = json.loads(research_json.read_text(encoding="utf-8"))
    except (OSError, ValueError) as exc:
        return [f"无法读取检索总账用于交叉校验：{research_json}（{exc}）"]
    included: set[tuple[str, str]] = set()
    for record in payload.get("records", []):
        code = str(record.get("country_code", "")).strip()
        for check in record.get("disease_checks", []):
            if str(check.get("status", "")).strip() == "已完成并纳入":
                included.add((code, str(check.get("disease", "")).strip()))
    missing = sorted(included - events)
    extra = sorted(events - included)
    if missing:
        errors.append(
            "检索总账标记“已完成并纳入”但底表缺少对应事件："
            + "、".join(f"{code}·{disease}" for code, disease in missing[:MAX_REPORTED])
        )
    if extra:
        errors.append(
            "底表存在检索总账未标记纳入的事件："
            + "、".join(f"{code}·{disease}" for code, disease in extra[:MAX_REPORTED])
        )
    return errors


def main() -> int:
    parser = argparse.ArgumentParser(description="校验 Word 与 Excel 交付件结构（不判断疫情事实）。")
    parser.add_argument("--xlsx", type=Path, required=True)
    parser.add_argument("--docx", type=Path, required=True)
    parser.add_argument("--research", type=Path, help="检索总账 JSON，用于底表事件与纳入结论交叉校验。")
    parser.add_argument("--list-file", type=Path, default=DEFAULT_PARTICIPANT_LIST)
    parser.add_argument("--disease-config", type=Path, default=DEFAULT_DISEASE_CONFIG)
    args = parser.parse_args()

    try:
        participants = load_participants(args.list_file)
        diseases = load_diseases(args.disease_config)
    except ConfigError as exc:
        print(f"配置读取失败：\n{exc}", file=sys.stderr)
        return 1
    for path in (args.xlsx, args.docx):
        if not path.is_file():
            print(f"交付件不存在：{path}", file=sys.stderr)
            return 1

    errors: list[str] = []
    try:
        workbook_errors, events = validate_workbook(
            args.xlsx,
            participants,
            diseases,
        )
    except (OSError, ValueError, ConfigError) as exc:
        print(f"无法解析 Excel：{args.xlsx}（{exc}）", file=sys.stderr)
        return 1
    errors.extend(workbook_errors)

    try:
        document = read_docx(args.docx)
    except (OSError, ValueError) as exc:
        print(f"无法解析 Word：{args.docx}（{exc}）", file=sys.stderr)
        return 1
    verified_web_urls: set[str] | None = None
    if args.research:
        verified_web_urls, source_registry_errors = load_verified_web_urls(args.research)
        errors.extend(source_registry_errors)
    errors.extend(validate_word_placeholders(document))
    errors.extend(validate_word_structure(document, diseases, participants.count))
    errors.extend(validate_word_public_sources(document, verified_web_urls))
    errors.extend(validate_word_format(document, diseases))
    if args.research:
        errors.extend(cross_check_research(args.research, events))

    if errors:
        print("交付件结构校验未通过：", file=sys.stderr)
        print("\n".join(f"- {item}" for item in errors), file=sys.stderr)
        return 1

    print(
        "交付件结构校验通过："
        f"动态名单 {participants.count} 个国家/地区、配置病种 {diseases.count} 个；"
        f"底表事件 {len(events)} 条；Word 含 1 张重点国家表与四个分类逐病种四部分结构。"
    )
    print(
        "说明：本结果只代表结构、格式、证据可追溯性及Word未展示本地文件来源的校验通过，不代表疫情事实已核实；"
        "Word 渲染与分页须由 build 流程的 --render-check 结果确认。"
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
