from __future__ import annotations

import re
from pathlib import Path

from docx import Document
from docx.enum.section import WD_ORIENT, WD_SECTION
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt
from PIL import Image, ImageDraw, ImageFont


ROOT = Path(__file__).resolve().parents[1]
OUT_DIR = ROOT / "output" / "docx"
ASSET_DIR = ROOT / "tmp" / "docx_report_assets"
OUT_DOCX = OUT_DIR / "Final_Project_Report_Final_Reviewed.docx"

SCREENSHOT_DIR = Path(
    r"C:\Users\avani'\AppData\Local\Packages\5319275A.WhatsAppDesktop_cv1g1gvanyjgm\LocalState\sessions"
    r"\6BCE5C9EBFE4E225093CF9E3A17A015C6141928C\transfers\2026-24"
)
UI_SCREENSHOTS = [
    SCREENSHOT_DIR / "WhatsApp Image 2026-06-15 at 21.57.36.jpeg",
    SCREENSHOT_DIR / "WhatsApp Image 2026-06-15 at 21.58.14.jpeg",
    SCREENSHOT_DIR / "WhatsApp Image 2026-06-15 at 22.01.41.jpeg",
    SCREENSHOT_DIR / "WhatsApp Image 2026-06-15 at 22.04.24.jpeg",
    SCREENSHOT_DIR / "WhatsApp Image 2026-06-15 at 22.03.49.jpeg",
]

REFERENCES = [
    '[1] Federal Aviation Administration, "Automatic Dependent Surveillance-Broadcast (ADS-B)," [Online]. Available: https://www.faa.gov/air_traffic/technology/adsb. [Accessed: 15-Jun-2026].',
    "[2] International Civil Aviation Organization, Technical Provisions for Mode S Services and Extended Squitter, 2nd ed., ICAO Doc 9871, 2012.",
    '[3] Python Software Foundation, "Python 3 Documentation," [Online]. Available: https://docs.python.org/3/. [Accessed: 15-Jun-2026].',
    '[4] Pallets Projects, "Flask Documentation 3.1.x," [Online]. Available: https://flask.palletsprojects.com/en/stable/. [Accessed: 15-Jun-2026].',
    '[5] SQLite Consortium, "About SQLite," [Online]. Available: https://www.sqlite.org/about.html. [Accessed: 15-Jun-2026].',
    '[6] SQLite Consortium, "Write-Ahead Logging," [Online]. Available: https://www.sqlite.org/wal.html. [Accessed: 15-Jun-2026].',
    '[7] Leaflet, "Leaflet API Reference," [Online]. Available: https://leafletjs.com/reference.html. [Accessed: 15-Jun-2026].',
    '[8] Chart.js Contributors, "Chart.js Documentation," [Online]. Available: https://www.chartjs.org/docs/latest/. [Accessed: 15-Jun-2026].',
    '[9] uAvionix, "pingStation ADS-B Receiver," [Online]. Available: https://uavionix.com/products/pingstation/. [Accessed: 15-Jun-2026].',
    '[10] R. E. Kalman, "A New Approach to Linear Filtering and Prediction Problems," Journal of Basic Engineering, vol. 82, no. 1, pp. 35-45, 1960, doi: 10.1115/1.3662552.',
    '[11] L. Breiman, "Random Forests," Machine Learning, vol. 45, pp. 5-32, 2001, doi: 10.1023/A:1010933404324.',
    '[12] S. Hochreiter and J. Schmidhuber, "Long Short-Term Memory," Neural Computation, vol. 9, no. 8, pp. 1735-1780, 1997, doi: 10.1162/neco.1997.9.8.1735.',
    '[13] Y. Liu and M. Hansen, "Predicting Aircraft Trajectories: A Deep Generative Convolutional Recurrent Neural Networks Approach," arXiv:1812.11670, 2018.',
    '[14] H. Sahfienya and A. C. Regan, "4D Flight Trajectory Prediction Using a Hybrid Deep Learning Prediction Method Based on ADS-B Technology," arXiv:2110.07774, 2021.',
    '[15] OpenStreetMap Contributors, "OpenStreetMap," [Online]. Available: https://www.openstreetmap.org/. [Accessed: 15-Jun-2026].',
]


def font(name: str = "arial.ttf", size: int = 30) -> ImageFont.FreeTypeFont:
    candidates = [
        Path(r"C:\Windows\Fonts") / name,
        Path(r"C:\Windows\Fonts\times.ttf"),
        Path(r"C:\Windows\Fonts\arial.ttf"),
    ]
    for candidate in candidates:
        if candidate.exists():
            return ImageFont.truetype(str(candidate), size)
    return ImageFont.load_default()


F_REG = font("times.ttf", 30)
F_BOLD = font("timesbd.ttf", 31)
F_SMALL = font("times.ttf", 24)
F_SMALL_BOLD = font("timesbd.ttf", 24)


def text_center(draw: ImageDraw.ImageDraw, box, text, fnt, fill=(0, 0, 0)):
    x1, y1, x2, y2 = box
    lines = text.split("\n")
    heights = [draw.textbbox((0, 0), line, font=fnt)[3] for line in lines]
    total = sum(heights) + (len(lines) - 1) * 8
    y = y1 + (y2 - y1 - total) / 2
    for line, h in zip(lines, heights):
        w = draw.textbbox((0, 0), line, font=fnt)[2]
        draw.text((x1 + (x2 - x1 - w) / 2, y), line, font=fnt, fill=fill)
        y += h + 8


def arrow(draw: ImageDraw.ImageDraw, start, end, fill=(30, 70, 110), width=4):
    draw.line([start, end], fill=fill, width=width)
    sx, sy = start
    ex, ey = end
    dx, dy = ex - sx, ey - sy
    if abs(dx) >= abs(dy):
        sign = 1 if dx > 0 else -1
        pts = [(ex, ey), (ex - sign * 16, ey - 9), (ex - sign * 16, ey + 9)]
    else:
        sign = 1 if dy > 0 else -1
        pts = [(ex, ey), (ex - 9, ey - sign * 16), (ex + 9, ey - sign * 16)]
    draw.polygon(pts, fill=fill)


def save_diagram(name: str, title: str, boxes: list[tuple[str, tuple[int, int, int, int], str]], arrows: list[tuple[int, int, int, int]]):
    img = Image.new("RGB", (1600, 620), "white")
    d = ImageDraw.Draw(img)
    d.text((45, 28), title, font=F_BOLD, fill=(0, 0, 0))
    for label, rect, color in boxes:
        d.rounded_rectangle(rect, radius=18, fill=color, outline=(40, 75, 95), width=4)
        text_center(d, rect, label, F_SMALL_BOLD if "\n" not in label else F_SMALL)
    for x1, y1, x2, y2 in arrows:
        arrow(d, (x1, y1), (x2, y2))
    path = ASSET_DIR / name
    img.save(path)
    return path


def make_diagrams() -> dict[str, Path]:
    ASSET_DIR.mkdir(parents=True, exist_ok=True)
    diagrams = {}
    diagrams["architecture"] = save_diagram(
        "architecture.png",
        "System Architecture",
        [
            ("PingStation\nReceiver / Replay Log", (45, 170, 310, 330), "#E9F4FF"),
            ("Decoder", (385, 170, 610, 330), "#FFF0C7"),
            ("Preprocessor\nValidation", (690, 170, 930, 330), "#EAF7E9"),
            ("SQLite + CSV\nStorage", (1010, 170, 1245, 330), "#F4E9FF"),
            ("Flask API +\nBrowser Dashboard", (1325, 170, 1555, 330), "#FFE6E6"),
        ],
        [(310, 250, 385, 250), (610, 250, 690, 250), (930, 250, 1010, 250), (1245, 250, 1325, 250)],
    )
    diagrams["dfd"] = save_diagram(
        "dfd.png",
        "Level-1 Data Flow Diagram",
        [
            ("1. Read\nSource", (70, 160, 300, 305), "#E9F4FF"),
            ("2. Decode\nMessages", (370, 160, 600, 305), "#FFF0C7"),
            ("3. Merge and\nValidate State", (670, 160, 930, 305), "#EAF7E9"),
            ("4. Store\nSamples", (1000, 160, 1230, 305), "#F4E9FF"),
            ("5. Render API\nResponse", (1300, 160, 1535, 305), "#FFE6E6"),
            ("Current Aircraft\nStore", (665, 415, 935, 540), "#F2F2F2"),
            ("Trajectory\nDatabase", (1000, 415, 1230, 540), "#F2F2F2"),
        ],
        [(300, 232, 370, 232), (600, 232, 670, 232), (930, 232, 1000, 232), (1230, 232, 1300, 232), (800, 305, 800, 415), (1115, 305, 1115, 415)],
    )
    diagrams["er"] = save_diagram(
        "er_diagram.png",
        "Entity Relationship Diagram",
        [
            ("aircraft\nicao24 PK\ncallsign, model", (70, 160, 350, 330), "#E9F4FF"),
            ("trajectory_sample\nid PK\nicao24 FK\ntime, lat, lon, alt", (495, 135, 835, 355), "#FFF0C7"),
            ("status_log\nid PK\nfeed, counts, time", (975, 160, 1280, 330), "#EAF7E9"),
            ("export_file\nid PK\ntype, generated_at", (1325, 160, 1560, 330), "#F4E9FF"),
        ],
        [(350, 245, 495, 245), (835, 245, 975, 245), (1280, 245, 1325, 245)],
    )
    diagrams["class"] = save_diagram(
        "class_diagram.png",
        "Class Diagram",
        [
            ("PingStationRawDecoder\n+ split_frames()\n+ decode_payload()\n+ get_static_snapshot()", (55, 155, 375, 340), "#E9F4FF"),
            ("PingStationPreprocessor\n+ clean_altitude_m()\n+ clean_heading_deg()\n+ clean_speed_ms()\n+ preprocess()", (465, 115, 795, 370), "#FFF0C7"),
            ("FlightTrajectoryStore\n+ upsert_current()\n+ get_latest_matrix()\n+ build_predicted_rows()\n+ get_trajectory_history()", (890, 115, 1230, 370), "#EAF7E9"),
            ("Flask Application\n+ /api/aircraft\n+ /api/status\n+ /api/logging/status\n+ /api/training/trajectory", (390, 415, 770, 575), "#F4E9FF"),
            ("Dashboard JavaScript\n+ refresh()\n+ refreshStatus()\n+ renderTrail()\n+ select()", (890, 415, 1230, 575), "#FFE6E6"),
        ],
        [
            (375, 247, 465, 247),
            (795, 247, 890, 247),
            (630, 415, 630, 382),
            (795, 500, 890, 500),
        ],
    )
    diagrams["sequence"] = save_diagram(
        "sequence.png",
        "Sequence of Operation",
        [
            ("Source", (70, 115, 300, 210), "#E9F4FF"),
            ("Decoder", (370, 115, 600, 210), "#FFF0C7"),
            ("Store", (670, 115, 900, 210), "#EAF7E9"),
            ("Flask API", (970, 115, 1200, 210), "#F4E9FF"),
            ("Browser", (1270, 115, 1500, 210), "#FFE6E6"),
            ("Read frames", (100, 310, 270, 385), "#FFFFFF"),
            ("Decode aircraft", (395, 310, 575, 385), "#FFFFFF"),
            ("Persist state", (695, 310, 875, 385), "#FFFFFF"),
            ("Return JSON", (995, 310, 1175, 385), "#FFFFFF"),
            ("Update map", (1295, 310, 1475, 385), "#FFFFFF"),
        ],
        [(300, 162, 370, 162), (600, 162, 670, 162), (900, 162, 970, 162), (1200, 162, 1270, 162), (270, 350, 395, 350), (575, 350, 695, 350), (875, 350, 995, 350), (1175, 350, 1295, 350)],
    )
    diagrams["pipeline"] = save_diagram(
        "prediction_pipeline.png",
        "Trajectory Prediction Pipeline",
        [
            ("Recent trajectory\nsamples", (70, 210, 320, 360), "#E9F4FF"),
            ("Validate age\nand count", (395, 210, 645, 360), "#F2F2F2"),
            ("Fit lat/lon/alt\nagainst time", (720, 210, 970, 360), "#FFF0C7"),
            ("Extrapolate to\noutage age", (1045, 210, 1295, 360), "#FFE6E6"),
            ("Return labelled\nprediction", (1370, 210, 1560, 360), "#EAF7E9"),
        ],
        [(320, 285, 395, 285), (645, 285, 720, 285), (970, 285, 1045, 285), (1295, 285, 1370, 285)],
    )
    return diagrams


def annotate_ui_screenshots() -> list[Path]:
    labels = [
        [("Simulation FALSE", (1230, 70), (1035, 80)), ("Live stream age", (1330, 105), (1335, 84)), ("Selected aircraft", (70, 210), (120, 190)), ("Altitude chart", (1280, 595), (1245, 690))],
        [("Flight detail card", (1140, 640), (1030, 690)), ("Live aircraft markers", (985, 280), (960, 330)), ("Route controls", (125, 560), (110, 505))],
        [("Simulation TRUE", (1300, 75), (1235, 72)), ("Simulated aircraft list", (85, 205), (100, 175)), ("Map view for replay", (775, 160), (830, 210))],
        [("Predicted aircraft", (150, 180), (160, 170)), ("Dotted predicted path", (505, 390), (570, 410)), ("Prediction tag", (780, 555), (750, 515)), ("Selected detail panel", (985, 810), (990, 785))],
        [("API requests", (600, 90), (645, 80)), ("Repeated /api/aircraft polling", (680, 560), (635, 560))],
    ]
    result = []
    for idx, src in enumerate(UI_SCREENSHOTS, start=1):
        if not src.exists():
            continue
        img = Image.open(src).convert("RGB")
        w, h = img.size
        d = ImageDraw.Draw(img)
        for text, label_xy, target_xy in labels[idx - 1]:
            lx, ly = label_xy
            tx, ty = target_xy
            lx = min(max(lx, 10), w - 260)
            ly = min(max(ly, 10), h - 58)
            box = (lx, ly, lx + 245, ly + 46)
            d.rounded_rectangle(box, radius=8, fill=(255, 246, 246), outline=(180, 40, 40), width=3)
            d.text((lx + 10, ly + 10), text, font=F_SMALL_BOLD, fill=(120, 0, 0))
            arrow(d, (lx + 8, ly + 23), target_xy, fill=(190, 30, 30), width=4)
        out = ASSET_DIR / f"ui_annotated_{idx}.png"
        img.save(out)
        result.append(out)
    return result


def set_page_number(paragraph):
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = paragraph.add_run()
    fld_char1 = OxmlElement("w:fldChar")
    fld_char1.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = "PAGE"
    fld_char2 = OxmlElement("w:fldChar")
    fld_char2.set(qn("w:fldCharType"), "end")
    run._r.append(fld_char1)
    run._r.append(instr)
    run._r.append(fld_char2)


def configure_section(section):
    section.page_width = Inches(8.27)
    section.page_height = Inches(11.69)
    section.top_margin = Inches(1)
    section.bottom_margin = Inches(1)
    section.left_margin = Inches(1.5)
    section.right_margin = Inches(1.25)
    section.header_distance = Inches(0.3)
    section.footer_distance = Inches(0.35)


def configure_doc(doc: Document):
    configure_section(doc.sections[0])
    styles = doc.styles
    normal = styles["Normal"]
    normal.font.name = "Times New Roman"
    normal._element.rPr.rFonts.set(qn("w:eastAsia"), "Times New Roman")
    normal.font.size = Pt(12)
    normal.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
    normal.paragraph_format.space_after = Pt(6)
    for name in ["Heading 1", "Heading 2", "Heading 3"]:
        style = styles[name]
        style.font.name = "Times New Roman"
        style._element.rPr.rFonts.set(qn("w:eastAsia"), "Times New Roman")
        style.font.size = Pt(12)
        style.font.bold = True
        style.font.color.rgb = None
        style.paragraph_format.space_before = Pt(14 if name == "Heading 1" else 10)
        style.paragraph_format.space_after = Pt(6)
        style.paragraph_format.keep_with_next = True
    for name in ["Caption", "List Bullet"]:
        styles[name].font.name = "Times New Roman"
        styles[name].font.size = Pt(11 if name == "Caption" else 12)


def apply_clean_footers(doc: Document):
    for section in doc.sections:
        section.footer.is_linked_to_previous = False
        p = section.footer.paragraphs[0]
        for child in list(p._p):
            p._p.remove(child)
        set_page_number(p)
        for run in p.runs:
            run.font.name = "Times New Roman"
            run.font.size = Pt(10)


def add_para(doc, text="", bold=False, align=None, style=None):
    p = doc.add_paragraph(style=style)
    if align is not None:
        p.alignment = align
    if not style:
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY if align is None else align
        p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
    run = p.add_run(text)
    run.font.name = "Times New Roman"
    run.font.size = Pt(12)
    run.bold = bold
    return p


def add_heading(doc, text, level=1):
    return doc.add_heading(text, level=level)


def add_bullets(doc, items):
    for item in items:
        p = doc.add_paragraph(style="List Bullet")
        p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
        p.add_run(item).font.name = "Times New Roman"


def add_table(doc, rows, widths=None, font_size=10):
    table = doc.add_table(rows=len(rows), cols=len(rows[0]))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = "Table Grid"
    for r, row in enumerate(rows):
        for c, value in enumerate(row):
            cell = table.cell(r, c)
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            if widths:
                cell.width = Inches(widths[c])
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if r == 0 else WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.space_after = Pt(0)
            run = p.add_run(str(value))
            run.font.name = "Times New Roman"
            run.font.size = Pt(font_size)
            run.bold = r == 0
    doc.add_paragraph()
    return table


def add_figure(doc, path: Path, caption: str, width=5.45):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run().add_picture(str(path), width=Inches(width))
    cap = doc.add_paragraph(caption, style="Caption")
    cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cap.paragraph_format.space_after = Pt(8)


def add_title_page(doc):
    for _ in range(4):
        doc.add_paragraph()
    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = title.add_run("ADS-B LIVE AIRCRAFT TRACKER")
    r.bold = True
    r.font.name = "Times New Roman"
    r.font.size = Pt(18)
    sub = doc.add_paragraph()
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sub.add_run("Final Project Report").bold = True
    doc.add_paragraph()
    add_para(doc, "Submitted in partial fulfilment of the requirements for the award of the degree of Bachelor of Technology", align=WD_ALIGN_PARAGRAPH.CENTER)
    doc.add_paragraph()
    add_para(doc, "Submitted by", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    add_para(doc, "Avani Sharma", align=WD_ALIGN_PARAGRAPH.CENTER)
    doc.add_paragraph()
    add_para(doc, "Under the mentorship of", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    add_para(doc, "Mr. Ajeet Gaur, Scientist 'F'", align=WD_ALIGN_PARAGRAPH.CENTER)
    doc.add_paragraph()
    add_para(doc, "Department of Computer Science and Engineering", align=WD_ALIGN_PARAGRAPH.CENTER)
    add_para(doc, "2026", align=WD_ALIGN_PARAGRAPH.CENTER)
    doc.add_page_break()


def add_front_matter(doc):
    add_heading(doc, "CERTIFICATE", 1)
    add_para(
        doc,
        "This is to certify that the project report entitled ADS-B Live Aircraft Tracker has been prepared by Avani Sharma under the mentorship of Mr. Ajeet Gaur, Scientist 'F'. The work described in this report is an academic implementation of an aircraft monitoring and data-logging system using ADS-B receiver data, local storage and a browser-based dashboard.",
    )
    doc.add_paragraph("\n\n")
    add_para(doc, "Mentor: Mr. Ajeet Gaur", bold=True)
    add_para(doc, "Student: Avani Sharma", bold=True)
    doc.add_page_break()

    add_heading(doc, "ABSTRACT", 1)
    add_para(
        doc,
        "Automatic Dependent Surveillance-Broadcast (ADS-B) enables equipped aircraft to broadcast identity, position and motion information for surveillance and monitoring. This project designs and implements a local ADS-B aircraft-tracking system that accepts structured PingStation receiver output or replayed data, decodes the useful aircraft fields, filters unusable values, maintains current aircraft state, stores trajectory samples and presents the information through a web dashboard.",
    )
    add_para(
        doc,
        "The system uses Python and Flask for the backend, SQLite for embedded storage, and HTML, CSS, JavaScript, Leaflet and Chart.js for the browser interface. It supports live and simulation modes, selected-aircraft details, route traces, altitude graphs, exportable trajectory files and short-gap prediction. Prediction is deliberately limited to transparent sliding-window linear regression with strict technical thresholds, making the implementation explainable for academic evaluation.",
    )
    doc.add_page_break()

    add_heading(doc, "ACKNOWLEDGEMENT", 1)
    add_para(
        doc,
        "I express my sincere gratitude to my project mentor, Mr. Ajeet Gaur, Scientist 'F', for his valuable guidance, technical suggestions and support throughout the development of this project. His direction helped in improving the system design, implementation and documentation.",
    )
    add_para(
        doc,
        "I would also like to thank my colleague, Avani Sharma, for cooperation, thoughtful discussions and assistance during the project work. I am thankful to the faculty members of the department for their encouragement and for providing the academic environment required to complete this report.",
    )
    doc.add_page_break()

    add_heading(doc, "TABLE OF CONTENTS", 1)
    toc = [
        ("1. Objective", "1"),
        ("2. Requirement Analysis (SRS)", "2"),
        ("3. Software Design Specification", "8"),
        ("4. System Design", "11"),
        ("5. Coding and Implementation", "18"),
        ("6. Testing", "30"),
        ("7. Results and User Interface", "13"),
        ("Appendix A: Source Code Listing", "16"),
        ("Appendix B: File Inventory and Execution Notes", "72"),
        ("Appendix G: Algorithm Selection and Comparison", "72"),
        ("References", "76"),
    ]
    add_table(doc, [["Section", "Page"]]+toc, [4.5, 0.8], font_size=11)
    add_heading(doc, "LIST OF FIGURES", 1)
    figures = [
        ["Figure 1", "System Architecture"],
        ["Figure 2", "Level-1 Data Flow Diagram"],
        ["Figure 3", "Use-Case Diagram"],
        ["Figure 4", "Entity Relationship Diagram"],
        ["Figure 5", "Class Diagram"],
        ["Figure 6", "Sequence of Operation"],
        ["Figure 7", "Trajectory Prediction Pipeline"],
        ["Figure 8-Figure 12", "User Interface Screenshots"],
    ]
    add_table(doc, [["Figure", "Title"]]+figures, [1.5, 3.8], font_size=11)
    doc.add_page_break()


def add_srs(doc):
    add_heading(doc, "1. OBJECTIVE", 1)
    add_para(doc, "The objective of this project is to design and implement an ADS-B aircraft-tracking system that converts structured PingStation receiver data into useful aircraft information. Incoming surveillance records may be partial, repeated or stale, so the software decodes, validates, normalizes, stores and displays the data through a browser-based interface.")
    add_bullets(doc, [
        "Decode ICAO address, callsign, latitude, longitude, altitude, heading and speed when available.",
        "Normalize values into consistent units and reject unusable location records.",
        "Display current aircraft, route traces, status information and selected-aircraft details.",
        "Maintain local trajectory history for analysis and short-outage prediction.",
        "Export aircraft and trajectory data to CSV and Excel-compatible files.",
    ])
    add_heading(doc, "2. REQUIREMENT ANALYSIS (SRS)", 1)
    add_heading(doc, "2.1 Introduction", 2)
    add_para(doc, "Automatic Dependent Surveillance-Broadcast (ADS-B) enables equipped aircraft to broadcast identity, position and motion information for surveillance use [1], [2]. This project receives the higher-level stream supplied by a PingStation receiver, or replays a captured stream through the same decoder. It is an academic monitoring and data-logging application, not a replacement for certified air-traffic-control systems.")
    add_heading(doc, "2.2 Scope", 2)
    add_bullets(doc, [
        "Captured PingStation log replay and optional live TCP-stream reception.",
        "Frame separation, aircraft-field extraction and state merging.",
        "Validation, normalization, local storage, visualization and export.",
        "Short-gap trajectory extrapolation using recent observations; long-horizon flight-plan prediction is outside scope.",
    ])
    add_heading(doc, "2.3 Functional Requirements", 2)
    add_table(doc, [
        ["Requirement", "Description"],
        ["FR1", "Read aircraft data from a captured replay file or a live PingStation TCP stream."],
        ["FR2", "Decode identity, position and motion fields when those fields are available."],
        ["FR3", "Filter unusable records, including records with both coordinates equal to zero."],
        ["FR4", "Expose aircraft, status, logging and trajectory data through JSON API endpoints."],
        ["FR5", "Display aircraft on an interactive map with callsign and selected-aircraft details."],
        ["FR6", "Store current aircraft records and meaningful historical trajectory samples locally."],
        ["FR7", "Support CSV and ICAO-wise Excel-compatible exports for later analysis."],
    ], [1.1, 4.1])
    add_heading(doc, "2.4 Non-Functional Requirements", 2)
    add_bullets(doc, [
        "The aircraft display should refresh without a full page reload.",
        "The application should run locally without a separate database server.",
        "Decoder, preprocessing, storage and interface modules should remain independently maintainable.",
        "Repeated observations should not create unnecessary trajectory samples.",
        "Predicted records must be visibly distinguished from received records.",
    ])
    add_heading(doc, "2.5 Hardware and Software Requirements", 2)
    add_table(doc, [
        ["Category", "Requirement"],
        ["Processor", "Intel/AMD-compatible processor capable of running Python and a modern browser."],
        ["Memory", "Minimum 4 GB RAM recommended for the local backend and browser dashboard."],
        ["Operating system", "Windows 10/11 or another operating system supported by Python and Flask."],
        ["Backend", "Python, Flask, SQLite, requests and openpyxl."],
        ["Frontend", "HTML, CSS, JavaScript, Leaflet, Chart.js and OpenStreetMap tiles."],
        ["Data", "Captured PingStation logs, SQLite database, CSV exports and ICAO workbooks."],
    ], [1.4, 3.8])
    add_heading(doc, "2.6 Feasibility Study", 2)
    add_para(doc, "The system is technically feasible because its processing rate and limits are explicitly bounded. The browser requests aircraft data every 1 second and status information every 4 seconds. Live records are removed from the active snapshot after 45 seconds without data; simulation records use a 12-second stale limit. A trajectory sample is stored after 5 seconds or when movement exceeds configured distance, altitude, heading or speed thresholds. Prediction uses 4 to 20 recent samples and is limited to a maximum outage age of 180 seconds. These values are implemented configuration limits rather than claimed aviation-performance guarantees.")
    add_para(doc, "The selected software stack is suitable for a single-host academic deployment. Python and Flask provide the application and API layer [3], [4]; SQLite provides embedded local storage [5] and the project enables WAL mode for efficient local read/write activity [6]. Leaflet and Chart.js provide map and chart rendering [7], [8]. The principal constraints are receiver-data quality, internet access for map tiles and optional enrichment, and reduced prediction accuracy during sharp turns or prolonged outages.")
    add_table(doc, [
        ["Parameter", "Implemented value", "Purpose"],
        ["Aircraft API refresh", "1 s", "Updates markers and selected-aircraft data."],
        ["Status refresh", "4 s", "Updates live/simulation and receiver-health badges."],
        ["Live stale limit", "45 s", "Removes aircraft no longer present in the live snapshot."],
        ["Replay stale limit", "12 s", "Keeps deterministic replay responsive."],
        ["Storage thresholds", "5 s; 30 m; 20 m; 5 deg; 2 m/s", "Stores time or meaningful movement changes."],
        ["Regression window", "4-20 samples", "Provides enough points while emphasizing recent motion."],
        ["Prediction age limit", "180 s", "Prevents long, increasingly uncertain extrapolation."],
        ["Trajectory retention", "30 days", "Bounds local historical storage."],
    ], [1.45, 1.45, 2.3], font_size=9)
    add_heading(doc, "2.7 Product Functions", 2)
    add_bullets(doc, [
        "Select live or replay input, decode and normalize aircraft records.",
        "Display live and predicted aircraft with clear visual distinction.",
        "Store, retrieve and export trajectory history and logging status.",
    ])


def add_design(doc, diagrams):
    add_heading(doc, "3. SOFTWARE DESIGN SPECIFICATION", 1)
    add_para(doc, "The design separates receiver input, decoding, validation, persistence, API routing and browser rendering so that each responsibility can be tested and changed independently. The browser never reads the database directly; it uses JSON endpoints exposed by the Flask application.")
    add_figure(doc, diagrams["architecture"], "Figure 1: System Architecture")
    add_heading(doc, "3.1 Module Description", 2)
    add_table(doc, [
        ["Module", "Responsibility"],
        ["data_source.py", "Selects replay or live source and supplies structured records to the application."],
        ["pingstation_raw_decoder.py", "Separates frames and extracts ICAO, callsign, position and motion fields."],
        ["pingstation_preprocess.py", "Normalizes units, rejects unusable coordinates and prepares clean records."],
        ["flight_data_store.py", "Stores current aircraft, trajectory samples, exports and short-gap predictions."],
        ["app.py", "Defines Flask routes and connects backend state to the browser dashboard."],
        ["templates/index.html", "Renders the interactive map, details panel, charts and route controls."],
    ], [1.7, 3.6])
    add_figure(doc, diagrams["dfd"], "Figure 2: Level-1 Data Flow Diagram")
    add_heading(doc, "3.2 Use-Case Descriptions", 2)
    add_table(doc, [
        ["Use Case", "Primary Actor", "Brief Description"],
        ["Start tracker", "Student / Operator", "Launch the server and initialize decoder, storage and data source."],
        ["Replay captured data", "System", "Read saved PingStation frames and feed them to the decoder."],
        ["Decode aircraft record", "Decoder", "Extract available identity, position and motion fields."],
        ["View live aircraft", "Browser client", "Fetch JSON records and render aircraft markers."],
        ["Select aircraft", "Student / Operator", "Open details, route trace and altitude history."],
        ["Export trajectory", "Student / Operator", "Create CSV or Excel-compatible output for later analysis."],
    ], [1.45, 1.3, 2.55], font_size=9)
    diagrams["usecase"] = make_usecase()
    add_figure(doc, diagrams["usecase"], "Figure 3: Use-Case Diagram")
    add_heading(doc, "4. SYSTEM DESIGN", 1)
    add_para(doc, "System design focuses on data movement, persistent state and the interaction between backend services and the dashboard. The database is intentionally embedded so the project can run on a single laptop without a separate database server.")
    add_figure(doc, diagrams["er"], "Figure 4: Entity Relationship Diagram")
    add_figure(doc, diagrams["class"], "Figure 5: Class Diagram")
    add_para(doc, "The class diagram has been redrawn with separated connectors so that arrows terminate at box boundaries instead of passing through class contents. This makes the relationships readable in both printed report pages and presentation slides.")
    add_figure(doc, diagrams["sequence"], "Figure 6: Sequence of Operation")


def make_usecase() -> Path:
    img = Image.new("RGB", (1600, 620), "white")
    d = ImageDraw.Draw(img)
    d.text((45, 28), "Use-Case Diagram", font=F_BOLD, fill=(0, 0, 0))
    d.ellipse((90, 190, 190, 290), outline=(30, 30, 30), width=5)
    d.line((140, 290, 140, 430), fill=(30, 30, 30), width=5)
    d.line((65, 345, 215, 345), fill=(30, 30, 30), width=5)
    d.line((140, 430, 80, 545), fill=(30, 30, 30), width=5)
    d.line((140, 430, 200, 545), fill=(30, 30, 30), width=5)
    d.text((55, 555), "Student / Operator", font=F_SMALL_BOLD, fill=(0, 0, 0))
    boundary = (365, 120, 1540, 550)
    d.rounded_rectangle(boundary, radius=20, outline=(40, 75, 95), width=5)
    d.text((760, 135), "ADS-B Live Tracker System", font=F_SMALL_BOLD, fill=(0, 0, 0))
    cases = [
        ("Start Tracker", (440, 205, 690, 275)),
        ("Replay Receiver Data", (760, 205, 1060, 275)),
        ("View Live Aircraft", (1130, 205, 1430, 275)),
        ("Search Aircraft", (440, 330, 690, 400)),
        ("Select Aircraft", (760, 330, 1060, 400)),
        ("Export Trajectory", (1130, 330, 1430, 400)),
        ("Store Samples", (610, 455, 860, 525)),
        ("Review Logs", (950, 455, 1200, 525)),
    ]
    for label, rect in cases:
        d.ellipse(rect, fill="#F8FBFF", outline=(40, 75, 95), width=4)
        text_center(d, rect, label, F_SMALL_BOLD)
        arrow(d, (215, 345), (rect[0], (rect[1] + rect[3]) // 2), fill=(100, 100, 100), width=3)
    path = ASSET_DIR / "usecase.png"
    img.save(path)
    return path


def add_implementation_testing(doc, diagrams):
    add_heading(doc, "5. CODING AND IMPLEMENTATION", 1)
    add_para(doc, "Implementation follows a small modular Python application. The backend starts the configured data source, decodes incoming records, keeps an in-memory aircraft snapshot, writes meaningful trajectory samples and serves browser-facing JSON routes. The frontend periodically requests these routes and updates map markers, route traces, selected-aircraft cards and charts.")
    add_table(doc, [
        ["Endpoint / File", "Purpose"],
        ["/api/aircraft", "Returns current aircraft and labelled predicted rows for the dashboard."],
        ["/api/status", "Returns receiver/feed mode, aircraft counts and health information."],
        ["/api/logging/status", "Reports SQLite, CSV and workbook logging state."],
        ["/api/training/trajectory", "Returns stored trajectory points for analysis/export use."],
        ["ADSBTracker.cfg", "Optional configuration file read at startup with environment-variable overrides."],
    ], [1.8, 3.4])
    add_heading(doc, "6. TESTING", 1)
    add_para(doc, "Testing used captured PingStation replay data and local Flask execution. Python source files were compiled with compileall, JavaScript syntax was checked with Node.js, and functional tests verified the data path from replay/live input to API output, storage, prediction and visualization.")
    add_table(doc, [
        ["Test Case", "Expected Result", "Observed Result"],
        ["Python syntax", "All modules compile.", "Compilation completed successfully."],
        ["Dashboard refresh", "Frontend requests API pages.", "No reload required during monitoring."],
        ["Replay captured log", "Known frames are decoded.", "Known aircraft records were produced."],
        ["Aircraft API", "JSON contains normalized records.", "Dashboard consumed the returned records."],
        ["Zero-coordinate filtering", "Invalid records are excluded.", "Rows with both coordinates zero were removed."],
        ["Trajectory storage", "Latest and historical rows are stored.", "SQLite contains current and trajectory tables."],
        ["Prediction labelling", "Prediction rows are visually distinct.", "Predicted records carry a prediction label."],
        ["CSV export", "Exporting data can be opened.", "Exported CSV opens with expected columns."],
    ], [1.55, 1.9, 1.9], font_size=9)
    add_heading(doc, "6.1 Observed Outputs", 2)
    add_bullets(doc, [
        "Current aircraft records are visible in the sidebar and on the map.",
        "Selected aircraft details update without a full page reload.",
        "Route traces and altitude charts are available for selected aircraft.",
        "Prediction rows are marked separately from live receiver observations.",
    ])
    add_figure(doc, diagrams["pipeline"], "Figure 7: Trajectory Prediction Pipeline")


def add_ui_results(doc, ui_images):
    add_heading(doc, "7. RESULTS AND USER INTERFACE", 1)
    add_para(doc, "The user interface displays current aircraft, search controls, feed status, simulation/live indicators, selected-aircraft details, route traces and altitude charts. The following screenshots are resized with clear callouts for report and presentation use.")
    captions = [
        "Figure 8: Live Dashboard with Selected Aircraft",
        "Figure 9: Aircraft Detail Panel and Route History",
        "Figure 10: Simulation Mode Dashboard",
        "Figure 11: Predicted Aircraft and Dotted Path",
        "Figure 12: Local Flask API Activity in Terminal",
    ]
    for path, caption in zip(ui_images, captions):
        add_figure(doc, path, caption, width=5.45)
        add_bullets(doc, [
            "Highlighted labels point to the feature most likely to be discussed during presentation.",
            "The screenshot is kept wide so map, sidebar and status controls remain readable.",
        ])


def add_algorithm_appendix(doc, diagrams):
    add_heading(doc, "APPENDIX G: ALGORITHM SELECTION AND COMPARISON", 1)
    add_heading(doc, "G.1 Algorithm Used in the Project", 2)
    add_para(doc, "The implemented predictor uses sliding-window ordinary least-squares linear regression. Separate straight-line models are fitted to recent latitude, longitude and altitude observations as functions of time. The method is an interpretable short-horizon baseline; it does not claim to reconstruct complete flight plans.")
    add_heading(doc, "Implemented Constraints", 3)
    add_bullets(doc, [
        "Up to 20 recent samples are read; at least 4 valid time-ordered samples are required.",
        "Predictions are produced only after an aircraft disappears from the current session snapshot.",
        "The maximum prediction age is 180 seconds; older records are omitted.",
        "Constant speed, heading and vertical rate provide a deterministic fallback.",
        "Predicted rows are output-only and are never inserted into the real trajectory table.",
    ])
    add_heading(doc, "G.2 Selection Rationale", 2)
    add_para(doc, "Linear regression is computationally light, explainable and independent of a large training dataset. Its limitation is the local straight-motion assumption. Kalman filtering offers stronger state/noise modelling [10], while random forests [11] and LSTM-based models [12]-[14] require representative training data and more extensive validation. The current method is therefore appropriate as a transparent baseline for short receiver gaps.")
    add_figure(doc, diagrams["pipeline"], "Figure 13: Trajectory Prediction Pipeline")
    add_table(doc, [
        ["Criterion", "Project-fit explanation"],
        ["Interpretability", "Each predicted coordinate is produced from recent timestamped observations, so the method can be explained during review."],
        ["Data requirement", "No large labelled training set is required; the algorithm works with the local trajectory samples already stored by the application."],
        ["Runtime cost", "The computation is light enough for a browser-refresh workflow and is bounded by the 4-20 sample window."],
        ["Risk control", "The 180-second maximum prediction age prevents long extrapolations from being presented as reliable outputs."],
    ], [1.45, 3.85], font_size=9)

    section = doc.add_section(WD_SECTION.NEW_PAGE)
    section.orientation = WD_ORIENT.LANDSCAPE
    section.page_width, section.page_height = section.page_height, section.page_width
    section.top_margin = Inches(1)
    section.bottom_margin = Inches(1)
    section.left_margin = Inches(1)
    section.right_margin = Inches(1)
    add_heading(doc, "G.3 Comparison of Candidate Trajectory-Prediction Algorithms", 2)
    rows = [
        ["Algorithm", "Core idea", "History / training need", "Online cost", "Noise and turn behaviour", "Limitation", "Role in this project"],
        ["Last observation carried forward", "Hold the last position.", "One observation; no training.", "Negligible.", "No noise growth, but no motion or turn representation.", "Moving aircraft appears stationary.", "Emergency display fallback only."],
        ["Constant velocity and heading", "Advance position using speed, heading and vertical rate.", "One recent state; no training.", "Very low.", "Sensitive to noisy motion fields and cannot anticipate turns.", "Error grows during manoeuvres.", "Deterministic fallback when regression is unavailable."],
        ["Sliding-window linear regression (USED)", "Fit latitude, longitude and altitude against time.", "At least 4 and up to 20 samples; no offline training.", "Low.", "Averages several observations but assumes locally straight motion.", "Weak during sustained turns or acceleration.", "Transparent baseline for short receiver gaps."],
        ["Kalman filter [10]", "Update a dynamic state using prediction and measurement steps.", "Sequential observations; no labelled training set.", "Low to moderate.", "Explicitly models measurement noise; turn handling depends on state model.", "Needs process and covariance tuning.", "Most practical next real-time upgrade."],
        ["Polynomial regression", "Fit a curved function of position versus time.", "Several recent samples; no offline training.", "Low to moderate.", "Can follow gentle curvature but may amplify noise outside the sample window.", "Extrapolation can become unstable.", "Research comparison with a tightly limited horizon."],
        ["Random forest / boosting [11]", "Learn nonlinear corrections from engineered trajectory features.", "Representative labelled trajectories required.", "Moderate.", "Can learn nonlinear patterns represented in training data.", "Does not naturally model time unless features encode it.", "Offline comparative study after a verified dataset is collected."],
        ["LSTM / GRU [12]-[14]", "Learn temporal dependencies from trajectory sequences.", "Large, diverse sequential dataset required.", "Moderate to high.", "Can represent complex manoeuvres if training coverage is sufficient.", "Data-, compute- and validation-intensive.", "Future large-dataset work, not justified for the current prototype."],
    ]
    add_table(doc, rows, [1.15, 1.45, 1.65, 0.8, 1.8, 1.55, 1.65], font_size=8)
    add_para(doc, "Table 15: Trajectory Prediction Algorithm Comparison", align=WD_ALIGN_PARAGRAPH.CENTER)
    add_para(doc, "Selection conclusion: retain sliding-window linear regression as the explainable prototype baseline and evaluate Kalman filtering next using measured gap-prediction error.", align=WD_ALIGN_PARAGRAPH.CENTER)
    doc.add_page_break()
    add_heading(doc, "G.4 Implementation-Oriented Comparison", 2)
    rows2 = [
        ["Algorithm", "New data or state required", "Parameters / tuning", "Integration change", "Validation needed", "Recommended decision"],
        ["Constant velocity and heading", "Current latitude, longitude, speed, heading and vertical rate.", "Maximum age and physical sanity limits.", "Already available as the regression fallback.", "Replay tests for straight, climbing and descending motion.", "Retain as a safe fallback, not the primary predictor."],
        ["Sliding-window linear regression (USED)", "Recent timestamped latitude, longitude and altitude samples.", "Window size, minimum sample count and maximum prediction age.", "Already implemented in the storage prediction path.", "Measure horizontal and altitude error at 5-180 s gaps.", "Keep as the explainable prototype baseline."],
        ["Kalman filter [10]", "Position and velocity state plus uncertainty estimates.", "Process noise, measurement noise and initial covariance.", "Add a per-aircraft filter state and reset rules.", "Test noisy data, missing updates and manoeuvre transitions.", "Preferred next upgrade after quantitative baseline testing."],
        ["Polynomial regression", "Recent samples with enough time spread for a stable curve fit.", "Polynomial order, window length and strict extrapolation horizon.", "Replace the linear fit while retaining current history queries.", "Check overshoot, oscillation and sensitivity to outliers.", "Use only as a controlled experiment."],
        ["Random forest / boosting [11]", "Engineered features and labelled future-position targets.", "Tree depth, ensemble size and feature design.", "Add offline training, model loading and version control.", "Use held-out flights and prevent route leakage between splits.", "Defer until a representative labelled dataset exists."],
        ["LSTM / GRU [12]-[14]", "Long normalized trajectory sequences and future target sequences.", "Architecture, sequence length, optimizer and regularization.", "Add training pipeline, model runtime and hardware planning.", "Use route-separated validation, uncertainty checks and drift monitoring.", "Future research only; excessive for the current dataset."],
    ]
    add_table(doc, rows2, [1.35, 1.8, 1.65, 1.65, 1.8, 1.65], font_size=8)
    add_para(doc, "Table 16: Implementation-Oriented Algorithm Comparison", align=WD_ALIGN_PARAGRAPH.CENTER)

    section = doc.add_section(WD_SECTION.NEW_PAGE)
    configure_section(section)


def add_code_appendix(doc):
    add_heading(doc, "APPENDIX A: SOURCE CODE LISTING", 1)
    add_para(doc, "The following listing contains the main modules only. Third-party library code and generated cache files are intentionally excluded.")
    files = [
        ROOT / "app.py",
        ROOT / "data_source.py",
        ROOT / "preprocessing" / "pingstation_raw_decoder.py",
        ROOT / "preprocessing" / "pingstation_preprocess.py",
        ROOT / "storage" / "flight_data_store.py",
        ROOT / "templates" / "index.html",
    ]
    for file_path in files:
        if not file_path.exists():
            continue
        doc.add_page_break()
        add_heading(doc, f"A.{files.index(file_path) + 1} {file_path.relative_to(ROOT)}", 2)
        text = file_path.read_text(encoding="utf-8", errors="replace").splitlines()
        for i, line in enumerate(text, start=1):
            p = doc.add_paragraph()
            p.paragraph_format.line_spacing = Pt(10)
            p.paragraph_format.space_after = Pt(0)
            run = p.add_run(f"{i:04d}: {line[:98]}")
            run.font.name = "Courier New"
            run.font.size = Pt(8.2)


def add_references(doc):
    add_heading(doc, "REFERENCES", 1)
    for ref in REFERENCES:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.first_line_indent = Inches(-0.25)
        p.paragraph_format.left_indent = Inches(0.25)
        p.paragraph_format.line_spacing = Pt(12)
        p.paragraph_format.space_after = Pt(2)
        run = p.add_run(ref)
        run.font.name = "Times New Roman"
        run.font.size = Pt(10)


def add_appendices(doc):
    add_heading(doc, "APPENDIX B: FILE INVENTORY AND EXECUTION NOTES", 1)
    add_table(doc, [
        ["Item", "Description"],
        ["Backend entry point", "app.py starts Flask routes and background update flow."],
        ["Data source", "data_source.py reads replay files or live TCP stream settings."],
        ["Preprocessing", "preprocessing package decodes and normalizes PingStation records."],
        ["Storage", "storage/flight_data_store.py manages SQLite, CSV export and prediction."],
        ["Frontend", "templates/index.html contains the dashboard, map logic and charts."],
    ], [1.8, 3.5])
    add_heading(doc, "APPENDIX C: SAMPLE OUTPUTS", 1)
    add_bullets(doc, [
        "Aircraft API output contains ICAO address, callsign, altitude, speed, heading, source and prediction label.",
        "Trajectory export files preserve timestamped latitude, longitude, altitude and selected motion fields.",
        "Status output reports feed mode, visibility count, tracked count and last update age.",
    ])


def build():
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    ASSET_DIR.mkdir(parents=True, exist_ok=True)
    diagrams = make_diagrams()
    ui_images = annotate_ui_screenshots()
    doc = Document()
    configure_doc(doc)
    add_title_page(doc)
    add_front_matter(doc)
    add_srs(doc)
    add_design(doc, diagrams)
    add_implementation_testing(doc, diagrams)
    add_ui_results(doc, ui_images)
    add_code_appendix(doc)
    add_appendices(doc)
    add_algorithm_appendix(doc, diagrams)
    add_references(doc)
    apply_clean_footers(doc)
    doc.save(OUT_DOCX)
    print(OUT_DOCX)


if __name__ == "__main__":
    build()
