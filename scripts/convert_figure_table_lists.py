from __future__ import annotations

import argparse
import re
from pathlib import Path

from docx import Document
from docx.enum.text import WD_TAB_ALIGNMENT, WD_TAB_LEADER
from docx.oxml import OxmlElement
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.shared import Inches
from docx.table import Table
from docx.text.paragraph import Paragraph


FIG_RE = re.compile(r"^Fig\.\s+(\d+\.\d+):\s*(.+)$")
TABLE_RE = re.compile(r"^Table\s+(\d+\.\d+):\s*(.+)$")


def iter_blocks(doc: Document):
    for child in doc.element.body.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, doc)
        elif isinstance(child, CT_Tbl):
            yield Table(child, doc)


def insert_paragraph_before(table: Table, text: str, style_name: str = "Normal") -> Paragraph:
    element = table._tbl
    new_p = OxmlElement("w:p")
    element.addprevious(new_p)
    paragraph = Paragraph(new_p, table._parent)
    paragraph.style = style_name
    paragraph.add_run(text)
    return paragraph


def format_list_paragraph(paragraph: Paragraph) -> None:
    paragraph.paragraph_format.left_indent = Inches(0)
    paragraph.paragraph_format.first_line_indent = Inches(0)
    paragraph.paragraph_format.space_before = 0
    paragraph.paragraph_format.space_after = 0
    paragraph.paragraph_format.tab_stops.clear_all()
    paragraph.paragraph_format.tab_stops.add_tab_stop(
        Inches(6.2), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS
    )


def collect_captions(doc: Document):
    figures: list[tuple[str, str]] = []
    tables: list[tuple[str, str]] = []
    for p in doc.paragraphs:
        text = " ".join(p.text.split())
        fig_match = FIG_RE.match(text)
        if fig_match:
            figures.append((fig_match.group(1), fig_match.group(2)))
            continue
        table_match = TABLE_RE.match(text)
        if table_match:
            tables.append((table_match.group(1), table_match.group(2)))
    return figures, tables


def find_front_list_tables(doc: Document) -> tuple[Table, Table]:
    tables = list(doc.tables)
    if len(tables) < 2:
        raise RuntimeError("Could not find the front List of Figures and List of Tables tables.")
    return tables[0], tables[1]


def replace_table_with_entries(table: Table, entries: list[tuple[str, str]], pages: dict[str, str]) -> None:
    tbl_element = table._tbl
    parent = tbl_element.getparent()
    style_name = "Normal"
    for number, title in entries:
        page = pages.get(number, "")
        paragraph = insert_paragraph_before(table, f"{number} {title}\t{page}", style_name)
        format_list_paragraph(paragraph)
    parent.remove(tbl_element)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("input", type=Path)
    parser.add_argument("output", type=Path)
    parser.add_argument("--figure-pages", type=Path)
    parser.add_argument("--table-pages", type=Path)
    args = parser.parse_args()

    doc = Document(args.input)
    figures, tables = collect_captions(doc)
    figure_pages = {}
    table_pages = {}
    if args.figure_pages and args.figure_pages.exists():
        for line in args.figure_pages.read_text(encoding="utf-8").splitlines():
            number, page = line.split("\t", 1)
            figure_pages[number.lstrip("\ufeff").strip()] = page.strip()
    if args.table_pages and args.table_pages.exists():
        for line in args.table_pages.read_text(encoding="utf-8").splitlines():
            number, page = line.split("\t", 1)
            table_pages[number.lstrip("\ufeff").strip()] = page.strip()

    lof_table, lot_table = find_front_list_tables(doc)
    replace_table_with_entries(lof_table, figures, figure_pages)
    replace_table_with_entries(lot_table, tables, table_pages)

    doc.save(args.output)
    print(f"Saved: {args.output}")
    print(f"Figure entries: {len(figures)}")
    print(f"Table entries: {len(tables)}")


if __name__ == "__main__":
    main()
