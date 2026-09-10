"""只读解析 .docx（OOXML）为段落、表格与页面设置（仅标准库）。"""

from __future__ import annotations

import zipfile
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field
from pathlib import Path

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _q(tag: str) -> str:
    return f"{{{W}}}{tag}"


def _val(node: ET.Element | None, attr: str = "val") -> str | None:
    if node is None:
        return None
    return node.get(_q(attr))


@dataclass
class Paragraph:
    order: int
    text: str
    fonts: set[str] = field(default_factory=set)
    sizes: set[str] = field(default_factory=set)
    line: str | None = None
    line_rule: str | None = None
    space_before: str | None = None
    space_after: str | None = None
    first_line_chars: str | None = None
    first_line: str | None = None
    justify: str | None = None
    in_table: bool = False


@dataclass
class Table:
    order: int
    header: list[str]
    row_count: int


@dataclass
class Section:
    page_width: int | None
    page_height: int | None
    margin_top: int | None
    margin_bottom: int | None
    margin_left: int | None
    margin_right: int | None


@dataclass
class WordDocument:
    path: Path
    paragraphs: list[Paragraph]
    tables: list[Table]
    section: Section | None

    @property
    def body_paragraphs(self) -> list[Paragraph]:
        return [item for item in self.paragraphs if not item.in_table]

    @property
    def text(self) -> str:
        return "\n".join(item.text for item in self.paragraphs)


def _parse_paragraph(element: ET.Element, order: int, in_table: bool) -> Paragraph:
    text = "".join(node.text or "" for node in element.iter(_q("t")))
    paragraph = Paragraph(order=order, text=text.strip(), in_table=in_table)
    properties = element.find(_q("pPr"))
    if properties is not None:
        spacing = properties.find(_q("spacing"))
        if spacing is not None:
            paragraph.line = spacing.get(_q("line"))
            paragraph.line_rule = spacing.get(_q("lineRule"))
            paragraph.space_before = spacing.get(_q("before"))
            paragraph.space_after = spacing.get(_q("after"))
        indent = properties.find(_q("ind"))
        if indent is not None:
            paragraph.first_line_chars = indent.get(_q("firstLineChars"))
            paragraph.first_line = indent.get(_q("firstLine"))
        paragraph.justify = _val(properties.find(_q("jc")))
    for run in element.findall(_q("r")):
        run_text = "".join(node.text or "" for node in run.iter(_q("t")))
        if not run_text.strip():
            continue
        run_properties = run.find(_q("rPr"))
        if run_properties is None:
            paragraph.fonts.add("")
            paragraph.sizes.add("")
            continue
        fonts = run_properties.find(_q("rFonts"))
        east_asia = fonts.get(_q("eastAsia")) if fonts is not None else None
        size = _val(run_properties.find(_q("sz")))
        paragraph.fonts.add(east_asia or "")
        paragraph.sizes.add(size or "")
    return paragraph


def read_docx(path: Path) -> WordDocument:
    with zipfile.ZipFile(path) as archive:
        root = ET.fromstring(archive.read("word/document.xml"))
    body = root.find(_q("body"))
    if body is None:
        raise ValueError(f"Word 文档缺少 body：{path}")

    paragraphs: list[Paragraph] = []
    tables: list[Table] = []
    section: Section | None = None
    order = 0
    for element in body:
        tag = element.tag
        if tag == _q("p"):
            paragraphs.append(_parse_paragraph(element, order, in_table=False))
            order += 1
        elif tag == _q("tbl"):
            rows = element.findall(_q("tr"))
            header: list[str] = []
            if rows:
                header = [
                    "".join(node.text or "" for node in cell.iter(_q("t"))).strip()
                    for cell in rows[0].findall(_q("tc"))
                ]
            tables.append(Table(order=order, header=header, row_count=len(rows)))
            for row in rows:
                for cell in row.findall(_q("tc")):
                    for child in cell.findall(_q("p")):
                        paragraphs.append(_parse_paragraph(child, order, in_table=True))
            order += 1
        elif tag == _q("sectPr"):
            size = element.find(_q("pgSz"))
            margin = element.find(_q("pgMar"))

            def number(node: ET.Element | None, attr: str) -> int | None:
                if node is None:
                    return None
                raw = node.get(_q(attr))
                try:
                    return int(raw) if raw is not None else None
                except ValueError:
                    return None

            section = Section(
                page_width=number(size, "w"),
                page_height=number(size, "h"),
                margin_top=number(margin, "top"),
                margin_bottom=number(margin, "bottom"),
                margin_left=number(margin, "left"),
                margin_right=number(margin, "right"),
            )
    return WordDocument(path=path, paragraphs=paragraphs, tables=tables, section=section)
