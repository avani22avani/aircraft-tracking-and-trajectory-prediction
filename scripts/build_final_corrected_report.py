from __future__ import annotations

import importlib.util
import math
import re
import textwrap
from pathlib import Path

import pdfplumber
import pypdfium2 as pdfium
from pypdf import PdfReader, PdfWriter
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_LEFT
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas
from reportlab.platypus import Paragraph, Table, TableStyle


ROOT = Path(r"C:\Users\avani'\Desktop\adsb_tracker_2")
SOURCE_PDF = ROOT / "output" / "pdf" / "Final_Project_Report_Revised.pdf"
OUTPUT_PDF = ROOT / "output" / "pdf" / "Final_Project_Report_Final_Reviewed.pdf"
TMP = ROOT / "tmp" / "final_corrected_report"
TMP.mkdir(parents=True, exist_ok=True)


def register_fonts() -> None:
    fonts = Path(r"C:\Windows\Fonts")
    for name, file_name in [
        ("TNR", "times.ttf"),
        ("TNR-Bold", "timesbd.ttf"),
        ("TNR-Italic", "timesi.ttf"),
        ("Consolas", "consola.ttf"),
        ("Consolas-Bold", "consolab.ttf"),
    ]:
        pdfmetrics.registerFont(TTFont(name, str(fonts / file_name)))


def style(name: str, size=11, leading=16.5, bold=False, italic=False, align=TA_LEFT, left=0, first=0):
    font_name = "TNR-Bold" if bold else "TNR-Italic" if italic else "TNR"
    return ParagraphStyle(
        name,
        fontName=font_name,
        fontSize=size,
        leading=leading,
        alignment=align,
        textColor=colors.black,
        leftIndent=left,
        firstLineIndent=first,
        spaceAfter=0,
    )


BODY = style("Body", 12, 18, align=TA_JUSTIFY)
BODY_LEFT = style("BodyLeft", 12, 18)
SMALL = style("Small", 9, 12)
TABLE = style("Table", 8.2, 10)
TABLE_HEAD = style("TableHead", 8.2, 10, bold=True, align=TA_CENTER)
REPORT_LEFT = 38.1 * mm
REPORT_RIGHT = 31.75 * mm
REPORT_WIDTH = A4[0] - REPORT_LEFT - REPORT_RIGHT


def page_number(c: canvas.Canvas, value: str, page_size=A4):
    c.setFont("TNR", 11)
    c.drawCentredString(page_size[0] / 2, 15 * mm, value)


def draw_paragraph(c, text, st, x, top, width):
    p = Paragraph(text, st)
    _, h = p.wrap(width, 1000)
    p.drawOn(c, x, top - h)
    return top - h


def draw_heading(c, text, x, y, size=11):
    c.setFont("TNR-Bold", size)
    c.drawString(x, y, text)
    return y - 7 * mm


def draw_bullets(c, items, x, y, width, size=10.5, leading=14):
    st = style(f"bullets-{x}-{y}", size, leading)
    for item in items:
        c.setFont("TNR", size)
        c.drawString(x, y - 1.2 * mm, "•")
        y = draw_paragraph(c, item, st, x + 6 * mm, y + 1 * mm, width - 6 * mm) - 2 * mm
    return y


def table_cell(text, head=False):
    return Paragraph(str(text), TABLE_HEAD if head else TABLE)


def draw_table(c, rows, widths, x, top, header=True, font_size=8.2, padding=4, highlight_rows=None):
    data = []
    for ri, row in enumerate(rows):
        data.append([table_cell(v, head=header and ri == 0) for v in row])
    t = Table(data, colWidths=widths, repeatRows=1 if header else 0)
    commands = [
        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#333333")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), padding),
        ("RIGHTPADDING", (0, 0), (-1, -1), padding),
        ("TOPPADDING", (0, 0), (-1, -1), padding),
        ("BOTTOMPADDING", (0, 0), (-1, -1), padding),
    ]
    if header:
        commands.append(("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#E8EEF5")))
    for row_index in highlight_rows or []:
        commands.append(("BACKGROUND", (0, row_index), (-1, row_index), colors.HexColor("#FFF0B3")))
    t.setStyle(TableStyle(commands))
    _, h = t.wrap(sum(widths), 1000)
    t.drawOn(c, x, top - h)
    return top - h


def render_and_flatten(pdf_path: Path, output_pdf: Path, scale=2.6):
    rendered = pdfium.PdfDocument(str(pdf_path))[0].render(scale=scale).to_pil()
    image_path = TMP / f"{output_pdf.stem}.png"
    rendered.save(image_path, optimize=True)
    c = canvas.Canvas(str(output_pdf), pagesize=A4)
    c.drawImage(str(image_path), 0, 0, A4[0], A4[1], preserveAspectRatio=False)
    c.save()


def flatten_original_with_overlay(page_index: int, overlay: Path, output: Path):
    base = PdfReader(str(SOURCE_PDF))
    page = base.pages[page_index]
    page.merge_page(PdfReader(str(overlay)).pages[0])
    merged = TMP / f"merged-{page_index+1}.pdf"
    w = PdfWriter()
    w.add_page(page)
    with merged.open("wb") as fh:
        w.write(fh)
    render_and_flatten(merged, output)


def make_certificate(path: Path):
    c = canvas.Canvas(str(path), pagesize=A4)
    w, h = A4
    x = 31 * mm
    width = w - 62 * mm
    c.setFont("TNR-Bold", 14)
    c.drawCentredString(w / 2, h - 32 * mm, "CERTIFICATE")
    text = (
        'Certified that <b>"APARNA TIWARI"</b> (Roll No. 2216018, Student ID BTBTS22091) '
        'has carried out the project work titled <b>"AIRCRAFT TRACKING AND TRAJECTORY PREDICTION"</b> '
        'from __________ to __________ for the award of the <b>B.Tech (Computer Science - Artificial Intelligence)</b> '
        'degree from Banasthali Vidyapith under my supervision.'
    )
    y = draw_paragraph(c, text, style("Cert", 12, 20, align=TA_JUSTIFY), x, h - 55 * mm, width) - 7 * mm
    text2 = (
        "The report embodies the results of original work and study carried out by the student. "
        "Its contents have not formed the basis for the award of any other degree to the candidate or to any other person."
    )
    draw_paragraph(c, text2, style("Cert2", 12, 20, align=TA_JUSTIFY), x, y, width)
    c.setFont("TNR", 11)
    c.drawString(x, 76 * mm, "(Signature of the Supervisor)")
    c.setFont("TNR-Bold", 11)
    c.drawString(x, 68 * mm, "Mr. Ajeet Gaur")
    c.setFont("TNR", 11)
    c.drawString(x, 60 * mm, "Designation: Scientist 'F'")
    c.drawString(x, 52 * mm, "Place: Agra, Uttar Pradesh")
    c.drawString(x, 44 * mm, "Date:")
    page_number(c, "i")
    c.save()


def make_acknowledgement(path: Path):
    c = canvas.Canvas(str(path), pagesize=A4)
    x = 38.1 * mm
    width = A4[0] - x - 31.75 * mm
    y = A4[1] - 25.4 * mm
    c.setFont("TNR-Bold", 12)
    c.drawString(x, y, "ACKNOWLEDGEMENT")
    y -= 12 * mm
    paragraphs = [
        (
            "I express my sincere gratitude to my project mentor, <b>Mr. Ajeet Gaur, Scientist 'F', "
            "Aerial Delivery Research and Development Establishment (DRDO), Agra</b>, for his technical "
            "guidance, constructive feedback and continuous encouragement throughout the planning, implementation "
            "and documentation of this project. His suggestions helped me evaluate the system objectively and present "
            "the work in a disciplined and technically accurate manner."
        ),
        (
            "I am thankful to the faculty members of the Department of Computer Science and Artificial Intelligence, "
            "Banasthali Vidyapith, for providing the academic foundation and institutional support required to complete "
            "this project. I also acknowledge the facilities and working environment that enabled the development, "
            "testing and documentation of the aircraft-tracking application."
        ),
        (
            "I would also like to thank my colleague, <b>Avani Sharma</b>, for her cooperation, thoughtful discussions "
            "and support during the project. Her assistance in reviewing the application workflow and presentation of "
            "the results was valuable during the final stages of the work."
        ),
        (
            "Finally, I am deeply grateful to my family and friends for their patience, motivation and encouragement. "
            "Their support gave me the confidence to address the technical challenges encountered during the project "
            "and to complete this report with care and honesty."
        ),
    ]
    for paragraph in paragraphs:
        y = draw_paragraph(c, paragraph, BODY, x, y, width) - 7 * mm
    c.setFont("TNR-Bold", 12)
    c.drawRightString(A4[0] - 31.75 * mm, 46 * mm, "Aparna Tiwari")
    page_number(c, "iv")
    c.save()


def draw_index_line(c, label, page, x, y, width, size=10.5):
    c.setFont("TNR", size)
    page_text = str(page)
    page_width = pdfmetrics.stringWidth(page_text, "TNR", size)
    c.drawString(x, y, label)
    c.drawRightString(x + width, y, page_text)
    label_width = pdfmetrics.stringWidth(label, "TNR", size)
    dot_width = pdfmetrics.stringWidth(".", "TNR", size)
    start = x + label_width + 5
    end = x + width - page_width - 5
    if end > start and dot_width > 0:
        c.drawString(start, y, "." * int((end - start) / dot_width))


def make_contents_continuation(path: Path):
    c = canvas.Canvas(str(path), pagesize=A4)
    x = REPORT_LEFT
    width = REPORT_WIDTH
    y = A4[1] - 26 * mm
    c.setFont("TNR-Bold", 12)
    c.drawCentredString(A4[0] / 2, y, "TABLE OF CONTENTS (continued)")
    y -= 12 * mm
    rows = [
        ("Storage and prediction excerpt", "43"),
        ("Dashboard excerpt", "59"),
        ("Appendix C: Project File Inventory", "62"),
        ("Appendix D: Sample Output Records", "62"),
        ("Appendix E: Execution Steps", "64"),
        ("Appendix F: Sample Raw PingStation Log", "64"),
        ("Raw hex replay sample", "64"),
        ("Appendix G: Algorithm Selection and Comparison", "68"),
        ("Appendix H: User Interface Screenshots", "71"),
        ("REFERENCES", "76-77"),
    ]
    for label, page in rows:
        draw_index_line(c, label, page, x, y, width, 11)
        y -= 8 * mm
    page_number(c, "vii")
    c.save()


def make_list_pages():
    figures_path = TMP / "page-9-list-of-figures.pdf"
    c = canvas.Canvas(str(figures_path), pagesize=A4)
    x = REPORT_LEFT
    width = REPORT_WIDTH
    y = A4[1] - 26 * mm
    c.setFont("TNR-Bold", 12)
    c.drawString(x, y, "LIST OF FIGURES")
    y -= 11 * mm
    figures = [
        ("Figure 1: System Architecture", "7"), ("Figure 2: Level-1 Data Flow", "8"),
        ("Figure 3: Database Design", "9"), ("Figure 4: Use-Case Diagram", "5"),
        ("Figure 5: Entity Relationship Diagram", "9"), ("Figure 6: Class Diagram", "11"),
        ("Figure 7: Activity Diagram", "13"), ("Figure 8: Sequence Diagram", "12"),
        ("Figure 9: Flowchart of Data Processing", "14"), ("Figure 10: Component Diagram", "15"),
        ("Figure 11: Deployment Diagram", "15"), ("Figure 12: Database Schema Diagram", "10"),
        ("Figure 13: API Route Map", "16"), ("Figure 14: Trajectory Prediction Pipeline", "68"),
        ("Figure 15: Live Tracking Dashboard and Key Controls", "71"),
        ("Figure 16: Aircraft Details, Route and Altitude History", "72"),
        ("Figure 17: Simulation Mode Dashboard", "73"),
        ("Figure 18: Backend API Request Stream", "74"),
        ("Figure 19: Predicted Trajectory and Dotted Path", "75"),
    ]
    for label, page in figures:
        draw_index_line(c, label, page, x, y, width, 9.6)
        y -= 7.2 * mm
    page_number(c, "viii")
    c.save()

    tables_path = TMP / "page-10-list-of-tables.pdf"
    c = canvas.Canvas(str(tables_path), pagesize=A4)
    y = A4[1] - 26 * mm
    c.setFont("TNR-Bold", 12)
    c.drawString(x, y, "LIST OF TABLES")
    y -= 11 * mm
    tables = [
        ("Table 1: Hardware and Software Requirements", "2"),
        ("Table 2: Functional Requirements", "2"),
        ("Table 3: Main System Modules", "10"), ("Table 4: Database Tables", "10"),
        ("Table 5: Test Cases", "18"), ("Table 6: Observed Project Outputs", "19"),
        ("Table 7: Actors", "4"), ("Table 8: Use-Case Descriptions", "4"),
        ("Table 9: Data Dictionary", "6"), ("Table 10: API Specification", "6"),
        ("Table 11: Important Configuration Parameters", "17"),
        ("Table 12: Project File Inventory", "62"),
        ("Table 13: Sample Latest Aircraft Records", "62"),
        ("Table 14: Generated Trajectory CSV Files", "64"),
        ("Table 15: Trajectory Prediction Algorithm Comparison", "69"),
        ("Table 16: Implementation-Oriented Algorithm Comparison", "70"),
    ]
    for label, page in tables:
        draw_index_line(c, label, page, x, y, width, 9.8)
        y -= 7.8 * mm
    page_number(c, "ix")
    c.save()
    return figures_path, tables_path


def make_srs_pages():
    # Pages 11-14, numbered 1-4.
    x = REPORT_LEFT
    width = REPORT_WIDTH

    p11 = TMP / "page-11-srs.pdf"
    c = canvas.Canvas(str(p11), pagesize=A4)
    y = A4[1] - 24 * mm
    y = draw_heading(c, "1. OBJECTIVE", x, y, 12)
    objective = (
        "The objective of this project is to design and implement an ADS-B aircraft-tracking system that converts "
        "structured PingStation receiver data into useful aircraft information. Incoming surveillance records may be "
        "partial, repeated or stale, so the software decodes, validates, normalizes, stores and displays the data through "
        "a browser-based interface."
    )
    y = draw_paragraph(c, objective, BODY, x, y, width) - 4 * mm
    y = draw_bullets(c, [
        "Decode ICAO address, callsign, latitude, longitude, altitude, heading and speed when available.",
        "Normalize values into consistent units and reject unusable location records.",
        "Display current aircraft, route traces, status information and selected-aircraft details.",
        "Maintain local trajectory history for analysis and short-outage prediction.",
        "Export aircraft and trajectory data to CSV and Excel-compatible files.",
    ], x, y, width)
    y = draw_heading(c, "2. REQUIREMENT ANALYSIS (SRS)", x, y - 2 * mm, 12)
    y = draw_heading(c, "2.1 Introduction", x, y, 11)
    intro = (
        "Automatic Dependent Surveillance-Broadcast (ADS-B) enables equipped aircraft to broadcast identity, position "
        "and motion information for surveillance use [1], [2]. This project receives the higher-level stream supplied by "
        "a PingStation receiver, or replays a captured stream through the same decoder. It is an academic monitoring and "
        "data-logging application, not a replacement for certified air-traffic-control systems."
    )
    y = draw_paragraph(c, intro, BODY, x, y, width) - 4 * mm
    y = draw_heading(c, "2.2 Scope", x, y, 11)
    draw_bullets(c, [
        "Captured PingStation log replay and optional live TCP-stream reception.",
        "Frame separation, aircraft-field extraction and state merging.",
        "Validation, normalization, local storage, visualization and export.",
        "Short-gap trajectory extrapolation using recent observations; long-horizon flight-plan prediction is outside scope.",
    ], x, y, width)
    page_number(c, "1")
    c.save()

    p12 = TMP / "page-12-srs.pdf"
    c = canvas.Canvas(str(p12), pagesize=A4)
    y = A4[1] - 24 * mm
    y = draw_heading(c, "2.3 Functional Requirements", x, y, 11)
    functional = [
        ["Requirement", "Description"],
        ["FR1", "Read aircraft data from a captured replay file or a live PingStation TCP stream."],
        ["FR2", "Decode identity, position and motion fields when those fields are available."],
        ["FR3", "Filter unusable records, including records with both coordinates equal to zero."],
        ["FR4", "Expose aircraft, status, logging and trajectory data through JSON API endpoints."],
        ["FR5", "Display aircraft on an interactive map with callsign and selected-aircraft details."],
        ["FR6", "Store current aircraft records and meaningful historical trajectory samples locally."],
        ["FR7", "Support CSV and ICAO-wise Excel-compatible exports for later analysis."],
    ]
    y = draw_table(c, functional, [29 * mm, 111 * mm], x, y, padding=4) - 7 * mm
    y = draw_heading(c, "2.4 Non-Functional Requirements", x, y, 11)
    y = draw_bullets(c, [
        "The aircraft display should refresh without a full page reload.",
        "The application should run locally without a separate database server.",
        "Decoder, preprocessing, storage and interface modules should remain independently maintainable.",
        "Repeated observations should not create unnecessary trajectory samples.",
        "Predicted records must be visibly distinguished from received records.",
    ], x, y, width, size=10.2, leading=13.5)
    y = draw_heading(c, "2.5 Hardware and Software Requirements", x, y - 2 * mm, 11)
    hardware = [
        ["Category", "Requirement"],
        ["Processor", "Intel/AMD-compatible processor capable of running Python and a modern browser."],
        ["Memory", "Minimum 4 GB RAM recommended for the local backend and browser dashboard."],
        ["Operating system", "Windows 10/11 or another operating system supported by Python and Flask."],
        ["Backend", "Python, Flask, SQLite, requests and openpyxl."],
        ["Frontend", "HTML, CSS, JavaScript, Leaflet, Chart.js and OpenStreetMap tiles."],
        ["Data", "Captured PingStation logs, SQLite database, CSV exports and ICAO workbooks."],
    ]
    draw_table(c, hardware, [34 * mm, 106 * mm], x, y, padding=3)
    page_number(c, "2")
    c.save()

    p13 = TMP / "page-13-feasibility.pdf"
    c = canvas.Canvas(str(p13), pagesize=A4)
    y = A4[1] - 24 * mm
    y = draw_heading(c, "2.6 Feasibility Study", x, y, 12)
    feasibility = (
        "The system is technically feasible because its processing rate and limits are explicitly bounded. The browser "
        "requests aircraft data every 1 second and status information every 4 seconds. Live records are removed from the "
        "active snapshot after 45 seconds without data; simulation records use a 12-second stale limit. A trajectory sample "
        "is stored after 5 seconds or when movement exceeds configured distance, altitude, heading or speed thresholds. "
        "Prediction uses 4 to 20 recent samples and is limited to a maximum outage age of 180 seconds. These values are "
        "implemented configuration limits rather than claimed aviation-performance guarantees."
    )
    y = draw_paragraph(c, feasibility, BODY, x, y, width) - 5 * mm
    feasibility2 = (
        "The selected software stack is suitable for a single-host academic deployment. Python and Flask provide the "
        "application and API layer [3], [4]; SQLite provides embedded local storage [5] and the project enables WAL mode "
        "for efficient local read/write activity [6]. Leaflet and Chart.js provide map and chart rendering [7], [8]. "
        "The principal constraints are receiver-data quality, internet access for map tiles and optional enrichment, and "
        "reduced prediction accuracy during sharp turns or prolonged outages."
    )
    y = draw_paragraph(c, feasibility2, BODY, x, y, width) - 7 * mm
    c.setFont("TNR-Bold", 10.5)
    c.drawCentredString(A4[0] / 2, y, "Implemented Technical Thresholds")
    y -= 5 * mm
    thresholds = [
        ["Parameter", "Implemented value", "Purpose"],
        ["Aircraft API refresh", "1 s", "Updates markers and selected-aircraft data."],
        ["Status refresh", "4 s", "Updates live/simulation and receiver-health badges."],
        ["Live stale limit", "45 s", "Removes aircraft no longer present in the live snapshot."],
        ["Replay stale limit", "12 s", "Keeps deterministic replay responsive."],
        ["Storage thresholds", "5 s; 30 m; 20 m; 5 deg; 2 m/s", "Stores time or meaningful movement changes."],
        ["Regression window", "4-20 samples", "Provides enough points while emphasizing recent motion."],
        ["Prediction age limit", "180 s", "Prevents long, increasingly uncertain extrapolation."],
        ["Trajectory retention", "30 days", "Bounds local historical storage."],
    ]
    y = draw_table(c, thresholds, [37 * mm, 36 * mm, 67 * mm], x, y, padding=4) - 7 * mm
    y = draw_heading(c, "2.7 Product Functions", x, y, 11)
    draw_bullets(c, [
        "Select live or replay input, decode and normalize aircraft records.",
        "Display live and predicted aircraft with clear visual distinction.",
        "Store, retrieve and export trajectory history and logging status.",
    ], x, y, width, size=10.2, leading=13)
    page_number(c, "3")
    c.save()

    p14 = TMP / "page-14-actors.pdf"
    c = canvas.Canvas(str(p14), pagesize=A4)
    y = A4[1] - 24 * mm
    y = draw_heading(c, "2.8 Actors", x, y, 12)
    actors = [
        ["Actor", "Role in the System"],
        ["Student / Operator", "Runs the application, monitors aircraft and exports data."],
        ["PingStation receiver", "Provides the structured live receiver stream used by the decoder."],
        ["Browser client", "Requests Flask APIs and renders the dashboard."],
        ["Storage module", "Stores current records, trajectory samples and export files."],
        ["Future analysis module", "Uses saved trajectories for comparative prediction research."],
    ]
    y = draw_table(c, actors, [40 * mm, 100 * mm], x, y, padding=5) - 8 * mm
    y = draw_heading(c, "2.9 Use-Case Descriptions", x, y, 12)
    use_cases = [
        ["Use Case", "Primary Actor", "Brief Description"],
        ["Start tracker", "Student / Operator", "Launch the server and initialize the decoder, storage and data source."],
        ["Replay captured data", "System", "Read saved PingStation frames and feed them to the decoder in batches."],
        ["Decode aircraft record", "Decoder", "Extract available identity, position and motion fields."],
        ["View live aircraft", "Browser client", "Fetch JSON records and render aircraft markers."],
        ["Select aircraft", "Student / Operator", "Open details, route trace and altitude history."],
        ["Save trajectory sample", "Storage module", "Write meaningful movement samples to SQLite and CSV."],
        ["Export trajectory", "Student / Operator", "Create a trajectory CSV for a chosen time range."],
        ["Review logging status", "Student / Operator", "Check database, matrix CSV and workbook logging status."],
    ]
    draw_table(c, use_cases, [31 * mm, 32 * mm, 77 * mm], x, y, padding=3)
    page_number(c, "4")
    c.save()
    return [p11, p12, p13, p14]


def make_citation_overlay(page_no: int, replacements: list[tuple[str, str]], output: Path):
    """Replace short prose phrases while preserving the original page geometry."""
    overlay = TMP / f"overlay-{page_no}.pdf"
    c = canvas.Canvas(str(overlay), pagesize=A4)
    with pdfplumber.open(str(SOURCE_PDF)) as pdf:
        page = pdf.pages[page_no - 1]
        words = page.extract_words(use_text_flow=True, keep_blank_chars=False)
        for old, new in replacements:
            hits = [w for w in words if w["text"].strip('.,;:()') == old]
            for hit in hits:
                x0, x1 = hit["x0"], hit["x1"]
                y0 = A4[1] - hit["bottom"]
                y1 = A4[1] - hit["top"]
                c.setFillColor(colors.white)
                c.rect(x0 - 0.8, y0 - 0.8, max(x1 - x0 + 6, pdfmetrics.stringWidth(new, "TNR", 10.5) + 2), y1 - y0 + 1.6, fill=1, stroke=0)
                c.setFillColor(colors.black)
                font_size = max(8.2, min(11.2, (y1 - y0) * 0.83))
                c.setFont("TNR", font_size)
                c.drawString(x0, y0 + 0.6, new)
    c.showPage()
    c.save()
    flatten_original_with_overlay(page_no - 1, overlay, output)


def make_page20_module_fix(path: Path):
    overlay = TMP / "page20-overlay.pdf"
    c = canvas.Canvas(str(overlay), pagesize=A4)
    c.setFillColor(colors.white)
    c.rect(0, 465, A4[0], 300, fill=1, stroke=0)
    c.setFillColor(colors.black)
    c.setFont("TNR-Bold", 10.5)
    c.drawString(101, 728, "Figure 12: Database Schema Diagram")

    def schema_box(x, y, width, title, fields, fill):
        row_h = 12
        height = 36 + row_h * len(fields)
        c.setFillColor(colors.white)
        c.setStrokeColor(colors.HexColor("#40566B"))
        c.roundRect(x, y, width, height, 5, fill=1, stroke=1)
        c.setFillColor(fill)
        c.roundRect(x, y + height - 24, width, 24, 5, fill=1, stroke=1)
        c.setFillColor(colors.black)
        c.setFont("TNR-Bold", 8.5)
        c.drawCentredString(x + width / 2, y + height - 16, title)
        c.setFont("TNR", 6.4)
        line_y = y + height - 36
        for name, kind in fields:
            c.drawString(x + 8, line_y, name)
            c.drawRightString(x + width - 8, line_y, kind)
            line_y -= row_h
        return height

    latest_fields = [
        ("flight_key", "TEXT PK"), ("icao_address", "TEXT"), ("callsign", "TEXT"),
        ("timestamp", "TEXT"), ("lat_e6", "INTEGER"), ("lon_e6", "INTEGER"),
        ("altitude_m", "INTEGER"), ("speed_cms", "INTEGER"),
        ("heading_de2", "INTEGER"), ("vs_cms", "INTEGER"),
        ("squawk", "INTEGER NULL"), ("source_ip", "TEXT NULL"),
        ("sample_count", "INTEGER"),
    ]
    history_fields = [
        ("id", "INTEGER PK"), ("flight_key", "TEXT"), ("timestamp", "TEXT"),
        ("lat_e6", "INTEGER"), ("lon_e6", "INTEGER"), ("altitude_m", "INTEGER"),
        ("speed_cms", "INTEGER"), ("heading_de2", "INTEGER"),
        ("vs_cms", "INTEGER"), ("squawk", "INTEGER NULL"),
        ("source_ip", "TEXT NULL"), ("flight_key, timestamp", "UNIQUE"),
    ]
    box_y = 510
    schema_box(102, box_y, 155, "flight_latest", latest_fields, colors.HexColor("#DCEAF7"))
    schema_box(338, box_y, 155, "trajectory_samples", history_fields, colors.HexColor("#E2F0DF"))
    c.setStrokeColor(colors.black)
    c.line(257, 615, 338, 615)
    c.line(330, 619, 338, 615)
    c.line(330, 611, 338, 615)
    c.setFont("TNR", 6.5)
    c.drawCentredString(297.5, 631, "logical flight_key link")
    c.drawCentredString(297.5, 621, "latest state to history")
    c.setFont("TNR-Italic", 10.5)
    c.drawCentredString(A4[0] / 2, 480, "Figure 12: Database Schema Diagram")

    c.setFillColor(colors.white)
    # Mask the full lower-page band so no fragments of the original split
    # tables remain visible beside the corrected tables.
    c.rect(0, 65, A4[0], 400, fill=1, stroke=0)
    c.setFillColor(colors.black)
    c.setFont("TNR-Italic", 9.5)
    c.drawCentredString(A4[0] / 2, 398, "Table 4: Database Tables")
    db_rows = [["Table", "Purpose"], ["flight_latest", "Stores the latest known state for each flight key."], ["trajectory_samples", "Stores historical position and movement samples over time."]]
    draw_table(c, db_rows, [94, 300], 101, 388, padding=3)
    c.setFont("TNR-Bold", 10.5)
    c.drawString(101, 321, "4.4 Main System Modules")
    c.setFont("TNR-Italic", 9.5)
    c.drawCentredString(A4[0] / 2, 306, "Table 3: Main System Modules")
    modules = [
        ["Module", "File", "Responsibility"],
        ["Flask application", "app.py", "Dashboard routes, JSON APIs, status and exports."],
        ["Data source", "data_source.py", "Live/replay selection, configuration and receiver state."],
        ["Raw decoder", "preprocessing/pingstation_raw_decoder.py", "Frame splitting, field decoding and aircraft-state merging."],
        ["Preprocessor", "preprocessing/pingstation_preprocess.py", "Unit normalization and identity/timestamp cleanup."],
        ["Storage", "storage/flight_data_store.py", "SQLite, trajectory history, prediction and file exports."],
        ["Dashboard", "templates/index.html", "Map, list, details, route trace and altitude chart."],
    ]
    draw_table(c, modules, [88, 150, 156], 101, 296, padding=2.5)
    c.save()
    flatten_original_with_overlay(19, overlay, path)


def make_page18_data_flow_fix(path: Path):
    overlay = TMP / "page18-data-flow-overlay.pdf"
    c = canvas.Canvas(str(overlay), pagesize=A4)
    c.setFillColor(colors.white)
    c.rect(0, 610, A4[0], 180, fill=1, stroke=0)
    x = 38.1 * mm
    width = A4[0] - x - 31.75 * mm
    c.setFillColor(colors.black)
    c.setFont("TNR-Bold", 12)
    c.drawString(x, 760, "4.2 Data Flow Design")
    text = (
        "The data flow begins with receiver frames and ends with browser visualization and export files. "
        "The decoder validates supported frame layouts and merges partial updates into an ICAO-keyed aircraft state. "
        "The preprocessor normalizes available fields, while the storage layer excludes records whose latitude and "
        "longitude are both zero and suppresses unnecessary duplicate trajectory samples. This preserves useful partial "
        "identity or motion information without treating an unusable position as a valid track point."
    )
    draw_paragraph(c, text, BODY, x, 740, width)
    c.showPage()
    c.save()
    flatten_original_with_overlay(17, overlay, path)


def make_page21_sequence_fix(path: Path):
    c = canvas.Canvas(str(path), pagesize=A4)
    x = 38.1 * mm
    width = A4[0] - x - 31.75 * mm
    c.setFont("TNR-Bold", 12)
    c.drawString(x, 770, "Figure 6: Class Diagram")

    def class_box(bx, by, bw, title, lines, header="#DCEAF7"):
        row_h = 13
        bh = 39 + row_h * len(lines)
        c.setFillColor(colors.white)
        c.setStrokeColor(colors.HexColor("#40566B"))
        c.roundRect(bx, by, bw, bh, 4, fill=1, stroke=1)
        c.setFillColor(colors.HexColor(header))
        c.roundRect(bx, by + bh - 27, bw, 27, 4, fill=1, stroke=1)
        c.setFillColor(colors.black)
        c.setFont("TNR-Bold", 7.7)
        c.drawCentredString(bx + bw / 2, by + bh - 18, title)
        c.setFont("TNR", 6.6)
        ly = by + bh - 40
        for line in lines:
            c.drawString(bx + 7, ly, line)
            ly -= row_h
        return bh

    bw = 126
    top_y = 574
    class_box(73, top_y, bw, "PingstationRawDecoder", ["+ split_frames()", "+ decode_payload()", "+ get_state_snapshot()"])
    class_box(235, top_y, bw, "PingstationPreprocessor", ["+ clean_altitude_mm()", "+ clean_heading_de2()", "+ clean_speed_cms()", "+ preprocess()"])
    class_box(397, top_y, bw, "FlightTrajectoryStore", ["+ ingest_many()", "+ get_latest_matrix()", "+ build_predicted_rows()", "+ get_trajectory_history()"])
    class_box(154, 468, 164, "Flask Application", ["+ /api/aircraft", "+ /api/training/matrix", "+ /api/status", "+ /api/logging/status", "+ /api/training/trajectory"], "#E7EDF5")
    class_box(354, 468, 154, "Dashboard JavaScript", ["+ refresh()", "+ refreshStatus()", "+ ensure()", "+ renderList()", "+ select()"], "#E7EDF5")
    c.setStrokeColor(colors.black)
    for x1, y1, x2, y2 in [(199, 625, 235, 625), (361, 625, 397, 625), (298, 574, 245, 548), (318, 515, 354, 515)]:
        c.line(x1, y1, x2, y2)
        c.line(x2, y2, x2 - 5, y2 + 3)
        c.line(x2, y2, x2 - 5, y2 - 3)
    c.setFont("TNR-Italic", 10)
    c.drawCentredString(A4[0] / 2, 448, "Figure 6: Class Diagram")

    c.setFont("TNR-Bold", 12)
    c.drawString(x, 420, "4.5 Sequence of Operation")
    text = (
        "At startup, the data-source module reads ADSBTracker.cfg when it is present, then applies environment-variable "
        "overrides and built-in defaults. Flask creates shared preprocessing and storage objects, while the data-source "
        "module owns the shared decoder used by both replay and live receiver modes. Every second, the browser requests "
        "aircraft data; the backend decodes and normalizes the current state, stores eligible samples, adds age-limited "
        "predictions for temporarily missing aircraft and returns JSON for the dashboard."
    )
    y = draw_paragraph(c, text, BODY, x, 401, width) - 4 * mm
    steps = [
        "Application starts, reads effective configuration and prepares output directories.",
        "Data source selects captured replay or live TCP receiver mode.",
        "Shared decoder reconstructs frames and merges partial aircraft updates.",
        "Preprocessor normalizes identity, position, altitude, heading, speed and time fields.",
        "Storage updates the latest state and writes only threshold-qualified trajectory samples.",
        "Prediction logic may append labelled output-only rows for recently missing aircraft.",
        "Flask returns JSON and the dashboard updates markers, tracks, details and charts.",
    ]
    draw_bullets(c, steps, x, y, width, size=9.7, leading=12.7)
    page_number(c, "11")
    c.save()


def make_page24_design_fix(path: Path):
    overlay = TMP / "page24-overlay.pdf"
    c = canvas.Canvas(str(overlay), pagesize=A4)
    c.setFillColor(colors.white)
    c.rect(99, 89, 400, 150, fill=1, stroke=0)
    c.setFillColor(colors.black)
    c.setFont("TNR-Bold", 10.5)
    c.drawString(105, 224, "4.6 Design Rationale")
    text = (
        "A local Flask application was selected because the project requires compact JSON APIs rather than a large "
        "enterprise framework [4]. SQLite provides embedded file-based storage [5], and WAL mode supports the project's "
        "local read/write pattern [6]. CSV and Excel-compatible exports make captured data easy to inspect with common "
        "analytical tools. Leaflet provides interactive map layers and markers [7], while Chart.js renders the selected "
        "aircraft's altitude history [8]."
    )
    draw_paragraph(c, text, style("DesignFix", 9.2, 13.8, align=TA_JUSTIFY), 105, 207, 386)
    c.save()
    flatten_original_with_overlay(23, overlay, path)


def make_coding_testing_pages():
    x = REPORT_LEFT
    width = REPORT_WIDTH
    outputs = []

    p27 = TMP / "page-27-coding.pdf"
    c = canvas.Canvas(str(p27), pagesize=A4)
    y = A4[1] - 24 * mm
    y = draw_heading(c, "5.2 Decoder and Preprocessing", x, y, 11)
    y = draw_paragraph(c, (
        "The decoder reconstructs complete length-prefixed frames, extracts available aircraft fields and maintains an "
        "ICAO-keyed state snapshot. The preprocessor converts records to a stable schema: latitude/longitude in decimal "
        "degrees, altitude in millimetres, heading in degrees multiplied by 100, and velocities in centimetres per second."
    ), BODY, x, y, width) - 5 * mm
    y = draw_heading(c, "5.3 Storage and Export", x, y, 11)
    y = draw_paragraph(c, (
        "SQLite stores one latest record per flight and a filtered trajectory history. A new trajectory point is saved when "
        "time or configured motion thresholds are exceeded. The storage module also produces daily CSV files, trajectory "
        "exports and ICAO-wise workbooks. SQLite is embedded [5], and WAL mode is enabled for the local workload [6]."
    ), BODY, x, y, width) - 5 * mm
    y = draw_heading(c, "5.4 User Interface", x, y, 11)
    y = draw_paragraph(c, (
        "Leaflet renders map markers, live paths and dotted prediction paths [7]. Chart.js renders altitude history for the "
        "selected aircraft [8]. The interface polls aircraft data every second and status data every four seconds, while "
        "search, selection, route information and live/simulation controls remain on one dashboard."
    ), BODY, x, y, width) - 5 * mm
    y = draw_heading(c, "5.5 Error Handling and Data Quality", x, y, 11)
    y = draw_paragraph(c, (
        "The data source retains incomplete TCP bytes until a complete frame is available. The decoder checks supported "
        "layouts and coordinate ranges, the preprocessor handles missing values, and the storage layer rejects zero-position "
        "samples and suppresses excessive duplicates. Predicted rows are clearly labelled and are never written back as real observations."
    ), BODY, x, y, width) - 5 * mm
    y = draw_heading(c, "5.6 Configuration Parameters", x, y, 11)
    params = [
        ["Parameter", "Default / Purpose"],
        ["simulation_mode", "true; chooses captured replay or live receiver mode."],
        ["replay_batch", "20 frames per browser refresh cycle."],
        ["simulation stale_seconds", "12 s before replay aircraft leave the active snapshot."],
        ["live stale_seconds", "45 s before live aircraft leave the active snapshot."],
        ["FLIGHT_DB_MIN_INTERVAL_SEC", "5 s storage interval threshold."],
        ["FLIGHT_PREDICTION_MAX_AGE_SEC", "180 s maximum short-outage prediction age."],
        ["FLIGHT_DB_HISTORY_RETENTION_DAYS", "30 days of local trajectory history."],
    ]
    draw_table(c, params, [70 * mm, 70 * mm], x, y, padding=4)
    page_number(c, "17")
    c.save()
    outputs.append(p27)

    p28 = TMP / "page-28-testing.pdf"
    c = canvas.Canvas(str(p28), pagesize=A4)
    y = A4[1] - 24 * mm
    y = draw_heading(c, "6. TESTING", x, y, 12)
    y = draw_paragraph(c, (
        "Testing used captured PingStation replay data and local Flask execution. Python source files were compiled with "
        "<font name='Consolas'>compileall</font>, and the dashboard JavaScript passed Node.js syntax checking. Functional "
        "tests then verified the data path from replay/live input to API output, storage, prediction and visualization."
    ), BODY, x, y, width) - 5 * mm
    c.setFont("TNR-Italic", 9.5)
    c.drawCentredString(A4[0] / 2, y, "Table 5: Test Cases")
    y -= 5 * mm
    tests1 = [
        ["Test Case", "Expected Result", "Observed Result"],
        ["Python syntax", "Project Python modules compile.", "compileall completed successfully."],
        ["Dashboard syntax", "Embedded JavaScript parses.", "Node.js --check completed successfully."],
        ["Replay captured log", "Saved frames are decoded.", "Known aircraft records were produced."],
        ["Aircraft API", "JSON contains normalized records.", "Dashboard consumes /api/aircraft output."],
        ["Mode/status API", "Simulation/live state is returned.", "Status and runtime mode endpoints respond."],
        ["Coordinate filtering", "Zero-position rows are excluded.", "Rows with both coordinates zero are removed."],
        ["Trajectory storage", "Latest and historical rows are stored.", "SQLite contains current and trajectory tables."],
    ]
    draw_table(c, tests1, [37 * mm, 50 * mm, 53 * mm], x, y, padding=3)
    page_number(c, "18")
    c.save()
    outputs.append(p28)

    p29 = TMP / "page-29-testing-ui.pdf"
    c = canvas.Canvas(str(p29), pagesize=A4)
    y = A4[1] - 24 * mm
    c.setFont("TNR-Italic", 9.5)
    c.drawCentredString(A4[0] / 2, y, "Table 5: Test Cases (continued)")
    y -= 5 * mm
    tests2 = [
        ["Test Case", "Expected Result", "Observed Result"],
        ["Short-gap prediction", "Missing current aircraft may be extrapolated.", "Predicted rows are labelled PRED and age-limited."],
        ["Prediction recovery", "Real data replaces predictions.", "Prediction path clears when observations return."],
        ["CSV export", "Trajectory data can be exported.", "Export endpoint writes a CSV and reports row count."],
        ["Workbook logging", "ICAO files are created when enabled.", "Per-aircraft and master workbooks are present."],
        ["Dashboard display", "Map and details update.", "Markers, routes, tabs and altitude chart render."],
    ]
    y = draw_table(c, tests2, [37 * mm, 50 * mm, 53 * mm], x, y, padding=3) - 7 * mm
    y = draw_heading(c, "6.1 Observed Outputs", x, y, 11)
    observed = [
        ["Output", "Workspace snapshot (15 June 2026)"],
        ["Latest aircraft records in SQLite", "15"],
        ["Trajectory samples in SQLite", "649"],
        ["Daily trajectory CSV files", "11"],
        ["Trajectory export CSV files", "1"],
        ["ICAO workbooks in latest session", "15"],
    ]
    y = draw_table(c, observed, [85 * mm, 55 * mm], x, y, padding=3) - 5 * mm
    y = draw_paragraph(c, (
        "These counts are a dated workspace snapshot, not aviation-accuracy benchmarks. They may change when the system "
        "is run again, when simulation starts a new session, or when retention and export settings are changed."
    ), style("OutputNote", 9.6, 13.5, italic=True), x, y, width) - 5 * mm
    y = draw_heading(c, "7. USER INTERFACES", x, y, 11)
    y = draw_paragraph(c, (
        "The single-page dashboard contains a status bar, searchable traffic list, central map and resizable detail panel. "
        "Live paths are solid, predicted continuations are red and dotted, and PRED labels prevent extrapolated states "
        "from being mistaken for current receiver observations."
    ), BODY_LEFT, x, y, width)
    page_number(c, "19")
    c.save()
    outputs.append(p29)

    p30 = TMP / "page-30-conclusion.pdf"
    c = canvas.Canvas(str(p30), pagesize=A4)
    y = A4[1] - 24 * mm
    y = draw_heading(c, "7.1 Interface Layout and Controls", x, y, 11)
    y = draw_bullets(c, [
        "Search by callsign or ICAO address and select aircraft from the map or traffic list.",
        "Use map zoom and center controls; inspect flight, aircraft and route tabs in the lower panel.",
        "Read live/simulation mode, connection health, prediction age and altitude history without leaving the dashboard.",
    ], x, y, width) - 4 * mm
    y = draw_heading(c, "8. CONCLUSION AND FUTURE SCOPE", x, y, 12)
    y = draw_heading(c, "8.1 Conclusion", x, y, 11)
    y = draw_paragraph(c, (
        "The project implements a complete local pipeline from PingStation receiver or replay input to decoded aircraft "
        "records, persistent trajectory history and browser visualization. The tested source compiles successfully, the "
        "dashboard JavaScript passes syntax checking, and short receiver gaps are bridged by clearly labelled regression "
        "or kinematic predictions. The system is appropriate for education, software experimentation and trajectory-data "
        "study; it is not certified surveillance equipment."
    ), BODY, x, y, width) - 5 * mm
    y = draw_heading(c, "8.2 Future Scope", x, y, 11)
    y = draw_bullets(c, [
        "Evaluate longer live-receiver sessions and report quantitative error metrics against trusted reference tracks.",
        "Compare the current linear baseline with Kalman filtering [10], tree ensembles [11] and recurrent models [12]-[14].",
        "Add quality dashboards, altitude/time filters, authentication and controlled network deployment.",
        "Investigate route- and weather-aware trajectory models using a larger verified dataset.",
    ], x, y, width) - 4 * mm
    y = draw_heading(c, "APPENDICES", x, y, 11)
    draw_bullets(c, [
        "Appendix A: Configuration Summary",
        "Appendix B: Verified Main Code Listings",
        "Appendix C: Project File Inventory",
        "Appendix D: Sample Output Records",
        "Appendix E: Execution Steps",
        "Appendix F: Sample Raw PingStation Log",
        "Appendix G: Algorithm Selection and Comparison",
        "Appendix H: User Interface Screenshots",
    ], x, y, width, size=10, leading=12.5)
    page_number(c, "20")
    c.save()
    outputs.append(p30)
    return outputs


def visual_wrap(line: str, width: int):
    if len(line) <= width:
        return [line]
    pieces = []
    remaining = line
    first = True
    while len(remaining) > width:
        cut = width
        for marker in [", ", ") ", " = ", " and ", " or ", " + ", " - "]:
            idx = remaining.rfind(marker, 0, width)
            if idx > width * 0.55:
                cut = idx + len(marker)
                break
        pieces.append(remaining[:cut])
        remaining = ("| " + remaining[cut:].lstrip())
        first = False
    pieces.append(remaining)
    return pieces


def code_segments():
    # Page allocation intentionally matches the existing Appendix B page ranges.
    return [
        ("Backend route excerpt", ROOT / "app.py", 4, None),
        ("Data source excerpt", ROOT / "data_source.py", 4, None),
        ("Raw decoder excerpt", ROOT / "preprocessing" / "pingstation_raw_decoder.py", 11, None),
        ("Preprocessing excerpt", ROOT / "preprocessing" / "pingstation_preprocess.py", 3, None),
        ("Storage and prediction excerpt", ROOT / "storage" / "flight_data_store.py", 16, None),
        ("Dashboard excerpt", ROOT / "templates" / "index.html", 3, list(range(99, 130)) + list(range(129, 180)) + list(range(199, 370)) + list(range(403, 408))),
    ]


def make_code_pages():
    outputs = []
    report_page = 21
    for title, file_path, page_count, selected in code_segments():
        source_lines = file_path.read_text(encoding="utf-8", errors="replace").splitlines()
        indices = selected if selected is not None else list(range(len(source_lines)))
        wrapped = []
        for idx in indices:
            line = source_lines[idx].replace("\t", "    ")
            for part_no, part in enumerate(visual_wrap(line, 122)):
                wrapped.append((idx + 1 if part_no == 0 else None, part))
        chunks = []
        base = len(wrapped) // page_count
        extra = len(wrapped) % page_count
        pos = 0
        for pi in range(page_count):
            size = base + (1 if pi < extra else 0)
            chunks.append(wrapped[pos:pos + size])
            pos += size
        for pi, chunk in enumerate(chunks):
            out = TMP / f"code-{report_page:02}.pdf"
            c = canvas.Canvas(str(out), pagesize=A4)
            c.setFont("TNR-Bold", 9.5)
            label = f"Appendix B: {title}" if pi == 0 else f"{title} (continued)"
            c.drawString(18 * mm, A4[1] - 16 * mm, label)
            c.setFont("TNR-Italic", 7.2)
            c.drawRightString(A4[0] - 18 * mm, A4[1] - 16 * mm, str(file_path.relative_to(ROOT)).replace("\\", "/"))
            if report_page == 21:
                c.setFont("TNR", 7.4)
                c.drawString(18 * mm, A4[1] - 22 * mm, "Verified 15 June 2026: Python compileall and dashboard Node.js syntax check passed. Lines beginning '| ' are visual wraps.")
                top = A4[1] - 28 * mm
            else:
                top = A4[1] - 22 * mm
            available_h = top - 23 * mm
            max_lines = max(1, len(chunk))
            leading = min(8.0, available_h / max_lines)
            font_size = max(4.5, min(6.8, leading - 0.7))
            c.setFont("Consolas", font_size)
            y = top
            for line_no, code_line in chunk:
                prefix = f"{line_no:4d} " if line_no is not None else "     "
                safe = code_line.encode("latin-1", "replace").decode("latin-1")
                c.drawString(15 * mm, y, prefix + safe)
                y -= leading
            page_number(c, str(report_page))
            c.save()
            outputs.append(out)
            report_page += 1
    assert report_page == 62, report_page
    return outputs


def make_algorithm_pages():
    p78 = TMP / "page-78-algorithm.pdf"
    c = canvas.Canvas(str(p78), pagesize=A4)
    x = REPORT_LEFT
    width = REPORT_WIDTH
    y = A4[1] - 24 * mm
    y = draw_heading(c, "APPENDIX G: ALGORITHM SELECTION AND COMPARISON", x, y, 12)
    y = draw_heading(c, "G.1 Algorithm Used in the Project", x, y, 11)
    y = draw_paragraph(c, (
        "The implemented predictor uses sliding-window ordinary least-squares linear regression. Separate straight-line "
        "models are fitted to recent latitude, longitude and altitude observations as functions of time. The method is an "
        "interpretable short-horizon baseline; it does not claim to reconstruct complete flight plans."
    ), BODY, x, y, width) - 4 * mm
    y = draw_heading(c, "Implemented constraints", x, y, 10.5)
    y = draw_bullets(c, [
        "Up to 20 recent samples are read; at least 4 valid time-ordered samples are required.",
        "Predictions are produced only after an aircraft disappears from the current session snapshot.",
        "The maximum prediction age is 180 seconds; older records are omitted.",
        "Constant speed, heading and vertical rate provide a deterministic fallback.",
        "Predicted rows are output-only and are never inserted into the real trajectory table.",
    ], x, y, width, size=10.2, leading=13.5) - 4 * mm
    y = draw_heading(c, "G.2 Selection Rationale", x, y, 11)
    y = draw_paragraph(c, (
        "Linear regression is computationally light, explainable and independent of a large training dataset. Its limitation "
        "is the local straight-motion assumption. Kalman filtering offers stronger state/noise modelling [10], while random "
        "forests [11] and LSTM-based models [12]-[14] require representative training data and more extensive validation. "
        "The current method is therefore appropriate as a transparent baseline for short receiver gaps."
    ), BODY, x, y, width) - 8 * mm
    # Reuse a compact pipeline diagram.
    labels = ["Recent trajectory\nsamples", "Validate age\nand count", "Fit lat/lon/alt\nagainst time", "Extrapolate to\noutage age", "Return labelled\nprediction"]
    fills = ["#EAF3FA", "#F4F4F4", "#FFF1C7", "#FCE0E0", "#E4F3E8"]
    box_w, gap, box_h = 25 * mm, 3 * mm, 25 * mm
    start_x = 36.5 * mm
    diagram_y = 62 * mm
    for i, label in enumerate(labels):
        bx = start_x + i * (box_w + gap)
        c.setFillColor(colors.HexColor(fills[i])); c.setStrokeColor(colors.HexColor("#315D7D"))
        c.roundRect(bx, diagram_y, box_w, box_h, 2 * mm, fill=1, stroke=1)
        c.setFillColor(colors.black); c.setFont("TNR-Bold" if i == 2 else "TNR", 7.5)
        for j, line in enumerate(label.split("\n")):
            c.drawCentredString(bx + box_w / 2, diagram_y + (15 - j * 5) * mm, line)
        if i < 4:
            c.line(bx + box_w, diagram_y + box_h / 2, bx + box_w + gap, diagram_y + box_h / 2)
            c.line(bx + box_w + gap, diagram_y + box_h / 2, bx + box_w + gap - 3, diagram_y + box_h / 2 + 2)
            c.line(bx + box_w + gap, diagram_y + box_h / 2, bx + box_w + gap - 3, diagram_y + box_h / 2 - 2)
    c.setFont("TNR-Italic", 10)
    c.drawCentredString(A4[0] / 2, 51 * mm, "Figure 14: Trajectory Prediction Pipeline")
    page_number(c, "68")
    c.save()

    p79 = TMP / "page-79-comparison.pdf"
    size = landscape(A4)
    c = canvas.Canvas(str(p79), pagesize=size)
    x = 18 * mm
    c.setFont("TNR-Bold", 12)
    c.drawString(x, size[1] - 18 * mm, "G.3 Comparison of Candidate Trajectory-Prediction Algorithms")
    c.setFont("TNR", 9.5)
    c.drawString(x, size[1] - 27 * mm, "The highlighted row is implemented. The comparison is qualitative and does not claim measured accuracy equivalence.")
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
    compact = style("ComparisonCompact", 7.2, 8.6)
    compact_head = style("ComparisonCompactHead", 7.2, 8.6, bold=True, align=TA_CENTER)
    t = Table([[Paragraph(str(v), compact_head if ri == 0 else compact) for v in row] for ri, row in enumerate(rows)], colWidths=[31 * mm, 39 * mm, 43 * mm, 22 * mm, 47 * mm, 42 * mm, 45 * mm], repeatRows=1)
    t.setStyle(TableStyle([
        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#333333")),
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#D9E7F2")),
        ("BACKGROUND", (0, 3), (-1, 3), colors.HexColor("#FFF0B3")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 5), ("RIGHTPADDING", (0, 0), (-1, -1), 5),
        ("TOPPADDING", (0, 0), (-1, -1), 5), ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
    ]))
    _, th = t.wrap(size[0] - 36 * mm, size[1] - 55 * mm)
    t.drawOn(c, x, size[1] - 42 * mm - th)
    c.setFont("TNR-Italic", 9.5)
    c.drawCentredString(size[0] / 2, 20 * mm, "Table 15: Trajectory Prediction Algorithm Comparison")
    page_number(c, "69", size)
    c.save()

    p80 = TMP / "page-80-comparison-detail.pdf"
    c = canvas.Canvas(str(p80), pagesize=size)
    c.setFont("TNR-Bold", 12)
    c.drawString(x, size[1] - 18 * mm, "G.4 Implementation-Oriented Comparison")
    c.setFont("TNR", 9.5)
    c.drawString(x, size[1] - 27 * mm, "This table explains what would be required to replace or extend the current predictor in the existing Flask application.")
    rows2 = [
        ["Algorithm", "New data or state required", "Parameters / tuning", "Integration change", "Validation needed", "Recommended decision"],
        ["Constant velocity and heading", "Current latitude, longitude, speed, heading and vertical rate.", "Maximum age and physical sanity limits.", "Already available as the regression fallback.", "Replay tests for straight, climbing and descending motion.", "Retain as a safe fallback, not the primary predictor."],
        ["Sliding-window linear regression (USED)", "Recent timestamped latitude, longitude and altitude samples.", "Window size, minimum sample count and maximum prediction age.", "Already implemented in the storage prediction path.", "Measure horizontal and altitude error at 5-180 s gaps.", "Keep as the explainable prototype baseline."],
        ["Kalman filter [10]", "Position and velocity state plus uncertainty estimates.", "Process noise, measurement noise and initial covariance.", "Add a per-aircraft filter state and reset rules.", "Test noisy data, missing updates and manoeuvre transitions.", "Preferred next upgrade after quantitative baseline testing."],
        ["Polynomial regression", "Recent samples with enough time spread for a stable curve fit.", "Polynomial order, window length and strict extrapolation horizon.", "Replace the linear fit while retaining current history queries.", "Check overshoot, oscillation and sensitivity to outliers.", "Use only as a controlled experiment."],
        ["Random forest / boosting [11]", "Engineered features and labelled future-position targets.", "Tree depth, ensemble size and feature design.", "Add offline training, model loading and version control.", "Use held-out flights and prevent route leakage between splits.", "Defer until a representative labelled dataset exists."],
        ["LSTM / GRU [12]-[14]", "Long normalized trajectory sequences and future target sequences.", "Architecture, sequence length, optimizer and regularization.", "Add training pipeline, model runtime and hardware planning.", "Use route-separated validation, uncertainty checks and drift monitoring.", "Future research only; excessive for the current dataset."],
    ]
    t2 = Table([[Paragraph(str(v), compact_head if ri == 0 else compact) for v in row] for ri, row in enumerate(rows2)], colWidths=[34 * mm, 52 * mm, 47 * mm, 47 * mm, 50 * mm, 43 * mm], repeatRows=1)
    t2.setStyle(TableStyle([
        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#333333")),
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#D9E7F2")),
        ("BACKGROUND", (0, 2), (-1, 2), colors.HexColor("#FFF0B3")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 5), ("RIGHTPADDING", (0, 0), (-1, -1), 5),
        ("TOPPADDING", (0, 0), (-1, -1), 5), ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
    ]))
    _, th2 = t2.wrap(size[0] - 36 * mm, size[1] - 58 * mm)
    t2.drawOn(c, x, size[1] - 42 * mm - th2)
    c.setFont("TNR", 9.3)
    c.drawString(x, 27 * mm, "Selection conclusion: retain linear regression for transparency; evaluate Kalman filtering next using measured gap-prediction error.")
    c.setFont("TNR-Italic", 9.5)
    c.drawCentredString(size[0] / 2, 20 * mm, "Table 16: Implementation-Oriented Algorithm Comparison")
    page_number(c, "70", size)
    c.save()
    return [p78, p79, p80]


REFERENCES = [
    '[1] Federal Aviation Administration, "Automatic Dependent Surveillance-Broadcast (ADS-B)," [Online]. Available: https://www.faa.gov/air_traffic/technology/adsb. [Accessed: 15-Jun-2026].',
    '[2] International Civil Aviation Organization, Technical Provisions for Mode S Services and Extended Squitter, 2nd ed., ICAO Doc 9871, 2012.',
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


def make_reference_pages():
    outputs = []
    chunks = [REFERENCES[:9], REFERENCES[9:]]
    for i, refs in enumerate(chunks):
        out = TMP / f"references-{76+i}.pdf"
        c = canvas.Canvas(str(out), pagesize=A4)
        x = 28 * mm
        width = A4[0] - 56 * mm
        y = A4[1] - 25 * mm
        c.setFont("TNR-Bold", 12)
        c.drawString(x, y, "REFERENCES" if i == 0 else "REFERENCES (continued)")
        y -= 10 * mm
        ref_style = style(f"Ref{i}", 10.2, 15.3, align=TA_LEFT, left=7 * mm, first=-7 * mm)
        for ref in refs:
            y = draw_paragraph(c, ref, ref_style, x, y, width) - 4 * mm
        page_number(c, str(76 + i))
        c.save()
        outputs.append(out)
    return outputs


def make_ui_page_update(page_index: int, new_number: int, output: Path, remove_annotated=False):
    overlay = TMP / f"ui-overlay-{page_index+1}.pdf"
    c = canvas.Canvas(str(overlay), pagesize=A4)
    if remove_annotated:
        c.setFillColor(colors.white)
        c.rect(35, 735, 525, 82, fill=1, stroke=0)
        c.setFillColor(colors.black)
        c.setFont("TNR-Bold", 11)
        c.drawString(52, 778, "APPENDIX H: USER INTERFACE SCREENSHOTS")
    c.setFillColor(colors.white)
    c.rect(0, 25, A4[0], 40, fill=1, stroke=0)
    page_number(c, str(new_number))
    c.showPage()
    c.save()
    flatten_original_with_overlay(page_index, overlay, output)


def build():
    register_fonts()
    replacements: dict[int, Path] = {}
    make_certificate(TMP / "page-2-certificate.pdf")
    replacements[1] = TMP / "page-2-certificate.pdf"
    make_acknowledgement(TMP / "page-5-acknowledgement.pdf")
    replacements[4] = TMP / "page-5-acknowledgement.pdf"

    # Consistent product capitalization on preserved prose pages.
    for page_no in [3, 4, 15, 16, 17]:
        out = TMP / f"page-{page_no}-capitalization.pdf"
        make_citation_overlay(page_no, [("Pingstation", "PingStation")], out)
        replacements[page_no - 1] = out

    make_contents_continuation(TMP / "page-8-contents.pdf")
    replacements[7] = TMP / "page-8-contents.pdf"
    figure_list, table_list = make_list_pages()
    replacements[8] = figure_list
    replacements[9] = table_list

    for offset, page in enumerate(make_srs_pages(), start=10):
        replacements[offset] = page

    make_page20_module_fix(TMP / "page-20-modules.pdf")
    replacements[19] = TMP / "page-20-modules.pdf"
    make_page18_data_flow_fix(TMP / "page-18-data-flow.pdf")
    replacements[17] = TMP / "page-18-data-flow.pdf"
    make_page21_sequence_fix(TMP / "page-21-sequence.pdf")
    replacements[20] = TMP / "page-21-sequence.pdf"
    make_page24_design_fix(TMP / "page-24-design.pdf")
    replacements[23] = TMP / "page-24-design.pdf"

    for offset, page in enumerate(make_coding_testing_pages(), start=26):
        replacements[offset] = page

    for offset, page in enumerate(make_code_pages(), start=30):
        replacements[offset] = page

    algorithm_pages = make_algorithm_pages()
    replacements[77] = algorithm_pages[0]
    replacements[78] = algorithm_pages[1]
    inserted_comparison_page = algorithm_pages[2]

    for page_index, new_number in zip(range(79, 84), range(71, 76)):
        out = TMP / f"page-{page_index+1}-ui-updated.pdf"
        make_ui_page_update(page_index, new_number, out, remove_annotated=(page_index == 79))
        replacements[page_index] = out

    refs = make_reference_pages()
    reader = PdfReader(str(SOURCE_PDF))
    writer = PdfWriter()
    for index, page in enumerate(reader.pages):
        if index == 79:
            writer.add_page(PdfReader(str(inserted_comparison_page)).pages[0])
        if index == 84:
            writer.add_page(PdfReader(str(refs[0])).pages[0])
            writer.add_page(PdfReader(str(refs[1])).pages[0])
            continue
        if index in replacements:
            writer.add_page(PdfReader(str(replacements[index])).pages[0])
        else:
            writer.add_page(page)
    writer.add_metadata({
        "/Title": "Aircraft Tracking and Trajectory Prediction - Final Corrected Report",
        "/Author": "Aparna Tiwari",
        "/Subject": "College final project report with verified code listings and IEEE-style references",
    })
    OUTPUT_PDF.parent.mkdir(parents=True, exist_ok=True)
    with OUTPUT_PDF.open("wb") as fh:
        writer.write(fh)
    print(OUTPUT_PDF)


if __name__ == "__main__":
    build()
