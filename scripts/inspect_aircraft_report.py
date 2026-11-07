import re
import sys
from pathlib import Path

from docx import Document
from PIL import Image


doc_path = Path(sys.argv[1])
image_paths = [Path(p) for p in sys.argv[2:]]

doc = Document(doc_path)
print("paragraphs", len(doc.paragraphs), "tables", len(doc.tables), "sections", len(doc.sections))

print("\nHEADINGS_AND_NONEMPTY")
for idx, paragraph in enumerate(doc.paragraphs):
    text = paragraph.text.strip()
    if not text:
        continue
    is_heading = (
        text.isupper()
        or re.match(r"^(CHAPTER|[0-9]+(\.[0-9]+)*\s)", text)
        or text in {"REFERENCES", "BIBLIOGRAPHY", "APPENDIX", "ABSTRACT"}
    )
    if is_heading or idx < 80:
        print(f"{idx:03}: {text[:220]}")

print("\nCORRUPTED_CANDIDATES")
for idx, paragraph in enumerate(doc.paragraphs):
    text = paragraph.text
    if any(token in text for token in ("Ã", "Â", "â", "�", "ƒ", "€", "™")):
        print(f"{idx:03}: {text}")

print("\nTABLES")
for t_idx, table in enumerate(doc.tables):
    print("TABLE", t_idx, len(table.rows), len(table.columns))
    for row in table.rows[:12]:
        print(" | ".join(cell.text.replace("\n", " / ")[:120] for cell in row.cells))
    if len(table.rows) > 12:
        print("...")

print("\nIMAGES")
for path in image_paths:
    with Image.open(path) as img:
        print(path.name, img.size)
