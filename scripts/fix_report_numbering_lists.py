import re
import sys
from collections import defaultdict
from copy import deepcopy
from pathlib import Path

from docx import Document


SOURCE = Path(sys.argv[1])
OUTPUT = Path(sys.argv[2])

FIG_CAPTION_RE = re.compile(r"^(?:Fig\.|Figure)\s+(\d+(?:\.\d+)*):\s*(.+)$", re.I)
TABLE_CAPTION_RE = re.compile(r"^Table\s+(\d+(?:\.\d+)*):\s*(.+)$", re.I)

EXTRA_ABBREVIATIONS = {
    "CPR": "Compact Position Reporting",
    "DF17": "Downlink Format 17",
    "GPU": "Graphics Processing Unit",
    "GPS": "Global Positioning System",
    "GRU": "Gated Recurrent Unit",
    "LSTM": "Long Short-Term Memory",
    "OSI": "Open Systems Interconnection",
    "PI": "Parity/Interrogator Identifier field in a Mode-S message",
    "PRED": "Predicted aircraft marker used in the dashboard",
    "PSR": "Primary Surveillance Radar",
    "RTL-SDR": "RTL2832U-based Software-Defined Radio receiver",
    "TCP": "Transmission Control Protocol",
}


def set_paragraph_text(paragraph, text):
    for run in paragraph.runs:
        run.text = ""
    if paragraph.runs:
        paragraph.runs[0].text = text
    else:
        paragraph.add_run(text)


def set_cell_text(cell, text):
    if cell.paragraphs:
        set_paragraph_text(cell.paragraphs[0], text)
        for paragraph in cell.paragraphs[1:]:
            paragraph._element.getparent().remove(paragraph._element)
    else:
        cell.text = text


def delete_row(table, row):
    table._tbl.remove(row._tr)


def rebuild_two_column_table(table, header, rows):
    # Keep the first row formatting, and clone the first body row if available.
    while len(table.rows) > 1:
        delete_row(table, table.rows[-1])
    set_cell_text(table.rows[0].cells[0], header[0])
    set_cell_text(table.rows[0].cells[1], header[1])

    template = deepcopy(table.rows[0]._tr)
    for left, right in rows:
        table._tbl.append(deepcopy(template))
        row = table.rows[-1]
        set_cell_text(row.cells[0], left)
        set_cell_text(row.cells[1], right)


def current_chapter_for_index(paragraphs, index):
    chapter = None
    for paragraph in paragraphs[: index + 1]:
        text = paragraph.text.strip()
        match = re.match(r"^CHAPTER\s+(\d+)$", text, re.I)
        if match:
            chapter = int(match.group(1))
    return chapter


def renumber_figure_captions(doc):
    counters = defaultdict(int)
    old_to_new = {}
    figure_rows = []
    paragraphs = doc.paragraphs
    caption_indices = set()

    for index, paragraph in enumerate(paragraphs):
        text = paragraph.text.strip()
        match = FIG_CAPTION_RE.match(text)
        if not match:
            continue
        caption_indices.add(index)
        chapter = current_chapter_for_index(paragraphs, index)
        if chapter is None:
            chapter = 1
        counters[chapter] += 1
        old_number = match.group(1)
        title = match.group(2).strip()
        new_number = f"{chapter}.{counters[chapter]}"
        old_to_new[old_number] = new_number
        new_caption = f"Fig. {new_number}: {title}"
        set_paragraph_text(paragraph, new_caption)
        figure_rows.append((f"Fig. {new_number}", title))

    # Fix textual references such as "Figure 3.31 presents..." or "Fig. 3.2.21".
    for index, paragraph in enumerate(doc.paragraphs):
        if index in caption_indices:
            continue
        text = paragraph.text
        original = text
        for old, new in sorted(old_to_new.items(), key=lambda item: len(item[0]), reverse=True):
            text = re.sub(rf"\bFigure\s+{re.escape(old)}\b", f"Figure {new}", text)
            text = re.sub(rf"\bFig\.\s+{re.escape(old)}\b", f"Fig. {new}", text)
        if text != original:
            set_paragraph_text(paragraph, text)

    return figure_rows


def collect_table_captions(doc):
    rows = []
    for paragraph in doc.paragraphs:
        text = paragraph.text.strip()
        match = TABLE_CAPTION_RE.match(text)
        if match:
            rows.append((f"Table {match.group(1)}", match.group(2).strip()))
    return rows


def update_abbreviation_table(table):
    existing = {}
    for row in table.rows[1:]:
        key = row.cells[0].text.strip()
        val = row.cells[1].text.strip()
        if key:
            existing[key] = val

    for key, val in EXTRA_ABBREVIATIONS.items():
        existing.setdefault(key, val)

    rows = sorted(existing.items(), key=lambda item: item[0].replace("-", "").upper())
    rebuild_two_column_table(table, ("Abbreviation", "Expansion"), rows)


def fix_known_cross_references(doc):
    for paragraph in doc.paragraphs:
        text = paragraph.text
        if "presents the same system from a data-flow perspective" in text:
            text = re.sub(r"\bFigure\s+\d+(?:\.\d+)*", "Figure 3.8", text, count=1)
            text = re.sub(
                r"complementing the structural \(module-and-tier\) view of Figure\s+\d+(?:\.\d+)*",
                "complementing the structural (module-and-tier) view of Figure 3.7",
                text,
            )
            set_paragraph_text(paragraph, text)
        elif "summarizes the operational control flow" in text:
            text = re.sub(r"\bFigure\s+\d+(?:\.\d+)*", "Figure 3.9", text, count=1)
            set_paragraph_text(paragraph, text)


def main():
    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    doc = Document(SOURCE)

    figure_rows = renumber_figure_captions(doc)
    table_rows = collect_table_captions(doc)

    # Expected front matter layout in this report:
    # table 0 = List of Figures, table 1 = List of Tables, table 2 = List of Abbreviations.
    rebuild_two_column_table(doc.tables[0], ("Figure No.", "Title / Caption"), figure_rows)
    rebuild_two_column_table(doc.tables[1], ("Table No.", "Title / Caption"), table_rows)
    update_abbreviation_table(doc.tables[2])
    fix_known_cross_references(doc)

    doc.save(OUTPUT)
    print(OUTPUT)


if __name__ == "__main__":
    main()
