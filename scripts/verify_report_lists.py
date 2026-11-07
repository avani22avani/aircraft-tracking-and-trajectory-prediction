import re
import sys
from pathlib import Path

from docx import Document


doc = Document(Path(sys.argv[1]))

fig_re = re.compile(r"^(?:Fig\.|Figure)\s+(\d+(?:\.\d+)*):\s*(.+)$", re.I)
tab_re = re.compile(r"^Table\s+(\d+(?:\.\d+)*):\s*(.+)$", re.I)

body_figures = []
body_tables = []
for paragraph in doc.paragraphs:
    text = paragraph.text.strip()
    m = fig_re.match(text)
    if m:
        body_figures.append((f"Fig. {m.group(1)}", m.group(2).strip()))
    m = tab_re.match(text)
    if m:
        body_tables.append((f"Table {m.group(1)}", m.group(2).strip()))

list_figures = [
    (row.cells[0].text.strip(), row.cells[1].text.strip())
    for row in doc.tables[0].rows[1:]
]
list_tables = [
    (row.cells[0].text.strip(), row.cells[1].text.strip())
    for row in doc.tables[1].rows[1:]
]

bad = []
for i, paragraph in enumerate(doc.paragraphs):
    if any(token in paragraph.text for token in ["Ã", "Â", "â‚", "â€", "Г", "�"]):
        bad.append((i, paragraph.text[:160]))
for ti, table in enumerate(doc.tables):
    for ri, row in enumerate(table.rows):
        for ci, cell in enumerate(row.cells):
            if any(token in cell.text for token in ["Ã", "Â", "â‚", "â€", "Г", "�"]):
                bad.append((f"T{ti}R{ri}C{ci}", cell.text[:160]))

def labels(rows):
    return [row[0] for row in rows]

print("body_figures", len(body_figures), labels(body_figures))
print("list_figures", len(list_figures), labels(list_figures))
print("figures_match", body_figures == list_figures)
print("body_tables", len(body_tables), labels(body_tables))
print("list_tables", len(list_tables), labels(list_tables))
print("tables_match", body_tables == list_tables)
print("abbrev_rows", len(doc.tables[2].rows) - 1)
print("bad", len(bad))
print("images", len(doc.inline_shapes))
print("format", doc.styles["Normal"].font.name, doc.styles["Normal"].font.size.pt)
print("sections", [(round(s.left_margin.inches, 2), round(s.right_margin.inches, 2), round(s.top_margin.inches, 2), round(s.bottom_margin.inches, 2)) for s in doc.sections])
for item in bad[:10]:
    print(item)
