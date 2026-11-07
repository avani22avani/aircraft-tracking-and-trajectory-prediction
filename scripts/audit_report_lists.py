import re
import sys
from pathlib import Path

from docx import Document


doc_path = Path(sys.argv[1])
doc = Document(doc_path)

print("PARAGRAPHS", len(doc.paragraphs), "TABLES", len(doc.tables), "IMAGES", len(doc.inline_shapes))

print("\nBODY FIGURE CAPTIONS")
for i, paragraph in enumerate(doc.paragraphs):
    text = paragraph.text.strip()
    if re.match(r"^(Fig\.|Figure)\s*\d+(\.\d+)?", text, flags=re.I):
        print(f"{i:03}: {text}")

print("\nBODY TABLE CAPTIONS")
for i, paragraph in enumerate(doc.paragraphs):
    text = paragraph.text.strip()
    if re.match(r"^Table\s+\d+(\.\d+)?", text, flags=re.I):
        print(f"{i:03}: {text}")

print("\nTABLE OVERVIEW")
for ti, table in enumerate(doc.tables):
    first = " | ".join(cell.text.replace("\n", " / ").strip() for cell in table.rows[0].cells) if table.rows else ""
    print(f"TABLE {ti}: rows={len(table.rows)} cols={len(table.columns)} first={first[:160]}")
    if ti <= 2:
        for row in table.rows[:80]:
            print("  " + " | ".join(cell.text.replace("\n", " / ").strip()[:180] for cell in row.cells))
        if len(table.rows) > 80:
            print("  ...")

print("\nABBREVIATION TERMS IN BODY")
terms = set()
for paragraph in doc.paragraphs:
    text = paragraph.text
    for term in re.findall(r"\b[A-Z][A-Z0-9-]{1,}\b", text):
        if len(term) > 1 and not term.isdigit():
            terms.add(term)
print(", ".join(sorted(terms)))
