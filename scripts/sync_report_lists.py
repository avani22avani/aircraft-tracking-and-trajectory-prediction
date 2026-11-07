from __future__ import annotations

import argparse
import copy
import re
from pathlib import Path

from docx import Document
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.table import Table
from docx.text.paragraph import Paragraph


FIG_RE = re.compile(r"^(Fig\.)\s+(\d+)\.(\d+):\s*(.+)$", re.I)
TABLE_RE = re.compile(r"^(Table)\s+(\d+)\.(\d+):\s*(.+)$", re.I)


ABBREVIATIONS = {
    "AA": "Aircraft Address field in a Mode-S message",
    "ADRDE": "Aerial Delivery Research and Development Establishment",
    "ADS-B": "Automatic Dependent Surveillance - Broadcast",
    "API": "Application Programming Interface",
    "ASCII": "American Standard Code for Information Interchange",
    "ATC": "Air Traffic Control",
    "CA": "Capability field in a Mode-S message",
    "CPR": "Compact Position Reporting",
    "CSV": "Comma-Separated Values",
    "CSS3": "Cascading Style Sheets Level 3",
    "DF": "Downlink Format",
    "DF17": "Downlink Format 17",
    "DFD": "Data-Flow Diagram",
    "DO-260B": "RTCA Minimum Operational Performance Standard for ADS-B (Rev. B)",
    "DRDO": "Defence Research and Development Organisation",
    "EASA": "European Union Aviation Safety Agency",
    "ES6": "ECMAScript 2015 JavaScript standard",
    "ETA": "Estimated Time of Arrival",
    "GNSS": "Global Navigation Satellite System",
    "GPS": "Global Positioning System",
    "GPU": "Graphics Processing Unit",
    "GRU": "Gated Recurrent Unit",
    "HEX": "Hexadecimal",
    "HTML": "HyperText Markup Language",
    "HTTP": "HyperText Transfer Protocol",
    "Hz": "Hertz",
    "ICAO": "International Civil Aviation Organisation",
    "IEEE": "Institute of Electrical and Electronics Engineers",
    "JSON": "JavaScript Object Notation",
    "LSTM": "Long Short-Term Memory",
    "ME": "Message field (Mode-S Extended Squitter payload)",
    "MHz": "Mega-Hertz",
    "ML": "Machine Learning",
    "OLS": "Ordinary Least Squares",
    "OSI": "Open Systems Interconnection",
    "PI": "Parity/Interrogator Identifier field in a Mode-S message",
    "PRED": "Predicted aircraft marker used in the dashboard",
    "PSR": "Primary Surveillance Radar",
    "RAW": "Raw unprocessed data/log format",
    "REST": "Representational State Transfer",
    "RF": "Radio Frequency",
    "RTCA": "Radio Technical Commission for Aeronautics",
    "SDR": "Software-Defined Radio",
    "SQLite": "Self-contained, serverless SQL database engine",
    "SSR": "Secondary Surveillance Radar",
    "TCP": "Transmission Control Protocol",
    "TIS-B": "Traffic Information Service - Broadcast",
    "UAT": "Universal Access Transceiver (978 MHz ADS-B link)",
    "UTC": "Coordinated Universal Time",
}


def iter_blocks(doc: Document):
    for child in doc.element.body.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, doc)
        elif isinstance(child, CT_Tbl):
            yield Table(child, doc)


def set_paragraph_text_preserve_first_run(paragraph: Paragraph, text: str) -> None:
    if paragraph.text == text:
        return
    if not paragraph.runs:
        paragraph.add_run(text)
        return
    first = paragraph.runs[0]
    for run in paragraph.runs[1:]:
        run.text = ""
    first.text = text


def set_cell_text_preserve_style(cell, text: str) -> None:
    paragraph = cell.paragraphs[0]
    set_paragraph_text_preserve_first_run(paragraph, text)
    for extra in cell.paragraphs[1:]:
        extra._element.getparent().remove(extra._element)


def replace_table_rows(table: Table, rows: list[tuple[str, str]]) -> None:
    tbl = table._tbl
    header = copy.deepcopy(table.rows[0]._tr)
    template = copy.deepcopy(table.rows[1]._tr if len(table.rows) > 1 else table.rows[0]._tr)

    for tr in list(tbl.tr_lst):
        tbl.remove(tr)

    tbl.append(header)
    new_table = Table(tbl, table._parent)
    set_cell_text_preserve_style(new_table.rows[0].cells[0], rows[0][0])
    set_cell_text_preserve_style(new_table.rows[0].cells[1], rows[0][1])

    for left, right in rows[1:]:
        tr = copy.deepcopy(template)
        tbl.append(tr)
        row = Table(tbl, table._parent).rows[-1]
        set_cell_text_preserve_style(row.cells[0], left)
        set_cell_text_preserve_style(row.cells[1], right)


def renumber_captions(doc: Document):
    fig_counts: dict[int, int] = {}
    table_counts: dict[int, int] = {}
    figures: list[tuple[str, str]] = []
    tables: list[tuple[str, str]] = []
    changes: list[str] = []

    for block in iter_blocks(doc):
        if not isinstance(block, Paragraph):
            continue

        text = " ".join(block.text.split())
        fig_match = FIG_RE.match(text)
        table_match = TABLE_RE.match(text)

        if fig_match:
            chapter = int(fig_match.group(2))
            fig_counts[chapter] = fig_counts.get(chapter, 0) + 1
            new_no = f"Fig. {chapter}.{fig_counts[chapter]}"
            new_text = f"{new_no}: {fig_match.group(4)}"
            if text != new_text:
                changes.append(f"{text} -> {new_text}")
                set_paragraph_text_preserve_first_run(block, new_text)
            figures.append((new_no, fig_match.group(4)))

        elif table_match:
            chapter = int(table_match.group(2))
            table_counts[chapter] = table_counts.get(chapter, 0) + 1
            new_no = f"Table {chapter}.{table_counts[chapter]}"
            new_text = f"{new_no}: {table_match.group(4)}"
            if text != new_text:
                changes.append(f"{text} -> {new_text}")
                set_paragraph_text_preserve_first_run(block, new_text)
            tables.append((new_no, table_match.group(4)))

    return figures, tables, changes


def find_front_tables(doc: Document):
    tables = doc.tables
    if len(tables) < 3:
        raise RuntimeError("Expected List of Figures, List of Tables, and List of Abbreviations tables.")
    return tables[0], tables[1], tables[2]


def build_abbreviation_rows() -> list[tuple[str, str]]:
    rows = [("Abbreviation", "Expansion")]
    for key in sorted(ABBREVIATIONS, key=lambda item: item.upper()):
        rows.append((key, ABBREVIATIONS[key]))
    return rows


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("input", type=Path)
    parser.add_argument("output", type=Path)
    args = parser.parse_args()

    doc = Document(args.input)
    figures, tables, caption_changes = renumber_captions(doc)
    lof_table, lot_table, abbr_table = find_front_tables(doc)

    replace_table_rows(lof_table, [("Figure No.", "Title / Caption"), *figures])
    replace_table_rows(lot_table, [("Table No.", "Title / Caption"), *tables])
    replace_table_rows(abbr_table, build_abbreviation_rows())

    doc.save(args.output)
    print(f"Saved: {args.output}")
    print(f"Figures listed: {len(figures)}")
    print(f"Tables listed: {len(tables)}")
    print(f"Abbreviations listed: {len(ABBREVIATIONS)}")
    print(f"Caption number changes: {len(caption_changes)}")
    for change in caption_changes:
        print(change)


if __name__ == "__main__":
    main()
