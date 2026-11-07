from __future__ import annotations

import io
import math
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont
import pypdfium2 as pdfium
from pypdf import PdfReader, PdfWriter
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas
from reportlab.platypus import Paragraph, Table, TableStyle


ROOT = Path(r"C:\Users\avani'\Desktop\adsb_tracker_2")
SOURCE_PDF = Path(r"C:\Users\avani'\Desktop\Final_Project_Report.pdf")
OUTPUT_DIR = ROOT / "output" / "pdf"
TMP_DIR = ROOT / "tmp" / "revised_report"
ANNOTATED_DIR = TMP_DIR / "annotated"
OUTPUT_PDF = OUTPUT_DIR / "Final_Project_Report_Revised.pdf"

SCREENSHOTS = [
    Path(r"C:\Users\avani'\AppData\Local\Packages\5319275A.WhatsAppDesktop_cv1g1gvanyjgm\LocalState\sessions\6BCE5C9EBFE4E225093CF9E3A17A015C6141928C\transfers\2026-24\WhatsApp Image 2026-06-15 at 21.57.36.jpeg"),
    Path(r"C:\Users\avani'\AppData\Local\Packages\5319275A.WhatsAppDesktop_cv1g1gvanyjgm\LocalState\sessions\6BCE5C9EBFE4E225093CF9E3A17A015C6141928C\transfers\2026-24\WhatsApp Image 2026-06-15 at 21.58.14.jpeg"),
    Path(r"C:\Users\avani'\AppData\Local\Packages\5319275A.WhatsAppDesktop_cv1g1gvanyjgm\LocalState\sessions\6BCE5C9EBFE4E225093CF9E3A17A015C6141928C\transfers\2026-24\WhatsApp Image 2026-06-15 at 22.01.41.jpeg"),
    Path(r"C:\Users\avani'\AppData\Local\Packages\5319275A.WhatsAppDesktop_cv1g1gvanyjgm\LocalState\sessions\6BCE5C9EBFE4E225093CF9E3A17A015C6141928C\transfers\2026-24\WhatsApp Image 2026-06-15 at 22.03.49.jpeg"),
    Path(r"C:\Users\avani'\AppData\Local\Packages\5319275A.WhatsAppDesktop_cv1g1gvanyjgm\LocalState\sessions\6BCE5C9EBFE4E225093CF9E3A17A015C6141928C\transfers\2026-24\WhatsApp Image 2026-06-15 at 22.04.24.jpeg"),
]


def register_fonts() -> None:
    font_dir = Path(r"C:\Windows\Fonts")
    regular = font_dir / "times.ttf"
    bold = font_dir / "timesbd.ttf"
    italic = font_dir / "timesi.ttf"
    if regular.exists():
        pdfmetrics.registerFont(TTFont("TNR", str(regular)))
        pdfmetrics.registerFont(TTFont("TNR-Bold", str(bold)))
        pdfmetrics.registerFont(TTFont("TNR-Italic", str(italic)))
    else:
        # Built-in Times is metrically close and keeps the PDF portable.
        pdfmetrics.registerFont(TTFont("TNR", str(font_dir / "arial.ttf")))
        pdfmetrics.registerFont(TTFont("TNR-Bold", str(font_dir / "arialbd.ttf")))
        pdfmetrics.registerFont(TTFont("TNR-Italic", str(font_dir / "ariali.ttf")))


def font(size: int, bold: bool = False) -> ImageFont.FreeTypeFont:
    path = Path(r"C:\Windows\Fonts\arialbd.ttf" if bold else r"C:\Windows\Fonts\arial.ttf")
    return ImageFont.truetype(str(path), size)


def arrow_head(draw: ImageDraw.ImageDraw, start: tuple[int, int], end: tuple[int, int], color: str, width: int = 5) -> None:
    draw.line([start, end], fill=color, width=width)
    angle = math.atan2(end[1] - start[1], end[0] - start[0])
    length = 18
    spread = 0.55
    p1 = (end[0] - length * math.cos(angle - spread), end[1] - length * math.sin(angle - spread))
    p2 = (end[0] - length * math.cos(angle + spread), end[1] - length * math.sin(angle + spread))
    draw.polygon([end, p1, p2], fill=color)


def callout(draw: ImageDraw.ImageDraw, label: str, box: tuple[int, int, int, int], target: tuple[int, int], accent: str = "#d62828") -> None:
    x1, y1, x2, y2 = box
    draw.rounded_rectangle(box, radius=10, fill=(255, 255, 255, 235), outline=accent, width=4)
    f = font(max(18, int((y2 - y1) * 0.34)), bold=True)
    max_width = x2 - x1 - 20
    words = label.split()
    lines: list[str] = []
    current = ""
    for word in words:
        trial = f"{current} {word}".strip()
        if draw.textbbox((0, 0), trial, font=f)[2] <= max_width:
            current = trial
        else:
            if current:
                lines.append(current)
            current = word
    if current:
        lines.append(current)
    line_h = f.size + 3
    total_h = line_h * len(lines)
    y = y1 + (y2 - y1 - total_h) // 2
    for line in lines:
        draw.text((x1 + 10, y), line, fill="#111111", font=f)
        y += line_h
    candidates = [
        ((x1 + x2) // 2, y1),
        ((x1 + x2) // 2, y2),
        (x1, (y1 + y2) // 2),
        (x2, (y1 + y2) // 2),
    ]
    start = min(candidates, key=lambda p: (p[0] - target[0]) ** 2 + (p[1] - target[1]) ** 2)
    arrow_head(draw, start, target, accent)


def annotate_screenshots() -> list[Path]:
    ANNOTATED_DIR.mkdir(parents=True, exist_ok=True)
    specs = [
        [
            ("Live mode: simulation is OFF", (850, 118, 1210, 190), (1235, 79)),
            ("Fresh live-feed indicator", (1110, 205, 1448, 275), (1390, 78)),
            ("Selected aircraft telemetry", (470, 555, 850, 625), (640, 652)),
            ("Origin-to-destination route", (470, 655, 850, 725), (790, 711)),
        ],
        [
            ("Aircraft marker and live popup", (610, 105, 1000, 175), (860, 288)),
            ("Aircraft identity/photo lookup", (920, 500, 1280, 570), (1065, 578)),
            ("Altitude history for selection", (1110, 645, 1450, 715), (1310, 610)),
        ],
        [
            ("Simulation mode enabled", (940, 115, 1260, 185), (1450, 72)),
            ("SIM source badge", (1180, 205, 1460, 270), (1560, 72)),
            ("Replayed aircraft records", (20, 375, 350, 445), (120, 260)),
        ],
        [
            ("Aircraft endpoint returns HTTP 200", (790, 90, 1390, 165), (620, 135)),
            ("Status endpoint health checks", (790, 235, 1390, 310), (615, 285)),
            ("Continuous backend request stream", (780, 520, 1390, 595), (590, 560)),
        ],
        [
            ("Predicted aircraft are marked PRED", (15, 105, 410, 180), (135, 160)),
            ("Dotted line is the extrapolated path", (600, 115, 1060, 190), (425, 260)),
            ("Prediction age: 156 seconds", (35, 620, 430, 690), (400, 684)),
            ("Predicted state in details panel", (590, 625, 1020, 695), (710, 684)),
        ],
    ]
    outputs = []
    for idx, (src, annotations) in enumerate(zip(SCREENSHOTS, specs), 1):
        im = Image.open(src).convert("RGB")
        draw = ImageDraw.Draw(im, "RGBA")
        for label, box, target in annotations:
            callout(draw, label, box, target)
        out = ANNOTATED_DIR / f"ui-{idx}.png"
        im.save(out, quality=95)
        outputs.append(out)
    return outputs


def page_number(c: canvas.Canvas, n: str, page_size=A4) -> None:
    w, _ = page_size
    c.setFont("TNR", 11)
    c.drawCentredString(w / 2, 15 * mm, n)


def paragraph_style(name: str, size: float = 11, leading: float = 16.5, bold: bool = False, alignment: int = TA_LEFT) -> ParagraphStyle:
    return ParagraphStyle(
        name,
        fontName="TNR-Bold" if bold else "TNR",
        fontSize=size,
        leading=leading,
        alignment=alignment,
        textColor=colors.black,
        spaceAfter=5,
    )


BODY = paragraph_style("Body")
BODY_SMALL = paragraph_style("BodySmall", 9.7, 13.5)
HEADING = paragraph_style("Heading", 12, 16, True)
CAPTION = paragraph_style("Caption", 10, 13, False, TA_CENTER)
TABLE_TEXT = paragraph_style("TableText", 8.5, 10.5)
TABLE_HEAD = paragraph_style("TableHead", 8.5, 10.5, True, TA_CENTER)


def draw_para(c: canvas.Canvas, text: str, style: ParagraphStyle, x: float, y_top: float, width: float) -> float:
    p = Paragraph(text, style)
    _, h = p.wrap(width, 1000)
    p.drawOn(c, x, y_top - h)
    return y_top - h


def create_overlay(path: Path, painter) -> None:
    c = canvas.Canvas(str(path), pagesize=A4)
    painter(c)
    c.save()


def create_feasibility_overlay(path: Path) -> None:
    def paint(c: canvas.Canvas) -> None:
        # Replace the original generic heading/paragraph, retaining surrounding sections.
        c.setFillColor(colors.white)
        c.rect(101, 445, 405, 190, stroke=0, fill=1)
        c.setFillColor(colors.black)
        c.setFont("TNR-Bold", 11)
        c.drawString(108, 614, "2.6 Feasibility Study")
        text = (
            "Technical feasibility is supported by the implemented operating limits, not only by tool availability. "
            "The dashboard requests aircraft data every <b>1 s</b> and status every <b>4 s</b>. Live aircraft are "
            "removed after <b>45 s</b> without a frame; replay records use a <b>12 s</b> stale limit. A trajectory point "
            "is stored when it is the first sample, at least <b>5 s</b> has elapsed, movement is at least <b>30 m</b>, "
            "altitude changes by <b>20 m</b>, heading changes by <b>5 degrees</b>, or speed changes by <b>2 m/s</b>. "
            "Prediction uses the latest <b>4-20 samples</b> and is limited to feed gaps of at most <b>180 s</b>; if the "
            "linear fit is unavailable, constant-speed/heading kinematics is used. SQLite runs locally in WAL mode, "
            "so no separate database server is required and a 4 GB system is sufficient for the project workload. "
            "The main technical limits are map-tile/network availability, receiver data quality, and reduced accuracy "
            "during sharp turns or long outages. The system is suitable for academic monitoring, not certified air-traffic control."
        )
        style = paragraph_style("Feasibility", 8.25, 10.55)
        draw_para(c, text, style, 108, 594, 392)
    create_overlay(path, paint)


def create_toc_overlay(path: Path) -> None:
    def dotted(c: canvas.Canvas, label: str, page: str, y: float, x: float = 108) -> None:
        c.setFont("TNR", 11)
        c.drawString(x, y, label)
        c.setDash(1, 2)
        c.line(x + pdfmetrics.stringWidth(label, "TNR", 11) + 3, y - 1, 495, y - 1)
        c.setDash()
        c.drawRightString(505, y, page)
    def paint(c: canvas.Canvas) -> None:
        c.setFillColor(colors.white)
        c.rect(100, 510, 420, 145, stroke=0, fill=1)
        c.setFillColor(colors.black)
        dotted(c, "Appendix F: Sample Raw Pingstation Log", "64", 636)
        dotted(c, "Raw hex replay sample", "64", 612, x=132)
        dotted(c, "Appendix G: Algorithm Selection and Comparison", "68", 588)
        dotted(c, "Appendix H: Annotated User Interface Screenshots", "70", 564)
        dotted(c, "REFERENCES", "75", 540)
    create_overlay(path, paint)


def create_list_figures_overlay(path: Path) -> None:
    def paint(c: canvas.Canvas) -> None:
        c.setFont("TNR", 11)
        y = 397
        for line in [
            "Figure 14: Trajectory Prediction Pipeline",
            "Figure 15: Live Tracking Dashboard and Key Controls",
            "Figure 16: Aircraft Details, Route and Altitude History",
            "Figure 17: Simulation Mode Dashboard",
            "Figure 18: Backend API Request Stream",
            "Figure 19: Predicted Trajectory and Dotted Path",
        ]:
            c.drawString(108, y, line)
            y -= 26
    create_overlay(path, paint)


def create_list_tables_overlay(path: Path) -> None:
    def paint(c: canvas.Canvas) -> None:
        c.setFont("TNR", 11)
        c.drawString(108, 369, "Table 15: Trajectory Prediction Algorithm Comparison")
    create_overlay(path, paint)


def create_reference_overlay(path: Path) -> None:
    def paint(c: canvas.Canvas) -> None:
        y = 505
        refs = [
            "Kalman, R. E. (1960), A New Approach to Linear Filtering and Prediction Problems,",
            "Journal of Basic Engineering, 82(1), 35-45.",
            "Breiman, L. (2001), Random Forests, Machine Learning, 45, 5-32.",
            "Hochreiter, S. and Schmidhuber, J. (1997), Long Short-Term Memory,",
            "Neural Computation, 9(8), 1735-1780.",
        ]
        c.setFont("TNR", 11)
        for line in refs:
            c.drawString(108, y, line)
            y -= 23
        c.setFillColor(colors.white)
        c.rect(260, 34, 80, 25, stroke=0, fill=1)
        c.setFillColor(colors.black)
        page_number(c, "75")
    create_overlay(path, paint)


def create_sequence_page(path: Path) -> None:
    c = canvas.Canvas(str(path), pagesize=A4)
    w, h = A4
    c.setFont("TNR-Bold", 11)
    c.drawString(55 * mm, h - 32 * mm, "Figure 8: Sequence Diagram")

    actors = ["Browser", "Flask API", "Data Source", "Decoder /\nPreprocessor", "Storage", "Prediction"]
    xs = [22, 55, 88, 121, 154, 187]
    top = h - 52 * mm
    bottom = 56 * mm
    for x_mm, actor in zip(xs, actors):
        x = x_mm * mm
        c.setFillColor(colors.HexColor("#edf4fb"))
        c.setStrokeColor(colors.HexColor("#365b7d"))
        c.roundRect(x - 13 * mm, top, 26 * mm, 15 * mm, 2 * mm, fill=1, stroke=1)
        c.setFillColor(colors.black)
        lines = actor.split("\n")
        for j, line in enumerate(lines):
            c.setFont("TNR-Bold", 7.5)
            c.drawCentredString(x, top + (9.5 - j * 4.2) * mm, line)
        c.setDash(2, 2)
        c.setStrokeColor(colors.HexColor("#73879a"))
        c.line(x, top, x, bottom)
        c.setDash()

    def msg(y_mm: float, src: int, dst: int, text: str, dashed: bool = False, color="#263238") -> None:
        y = y_mm * mm
        x1, x2 = xs[src] * mm, xs[dst] * mm
        c.setStrokeColor(colors.HexColor(color))
        if dashed:
            c.setDash(4, 3)
        c.line(x1, y, x2, y)
        c.setDash()
        direction = 1 if x2 > x1 else -1
        c.line(x2, y, x2 - direction * 3 * mm, y + 1.5 * mm)
        c.line(x2, y, x2 - direction * 3 * mm, y - 1.5 * mm)
        c.setFont("TNR", 7.4)
        c.setFillColor(colors.HexColor(color))
        c.drawCentredString((x1 + x2) / 2, y + 2.2 * mm, text)

    msg(226, 0, 1, "GET /api/aircraft (every 1 s)")
    msg(212, 1, 2, "get_aircraft_data()")
    msg(198, 2, 3, "decode frames / obtain state")
    msg(184, 3, 1, "normalized live records", dashed=True)
    msg(170, 1, 4, "ingest_many(records)")
    msg(156, 4, 5, "build predicted rows for missing aircraft")
    msg(142, 5, 1, "short-gap predictions", dashed=True, color="#b22222")
    msg(128, 1, 0, "JSON: live + predicted aircraft", dashed=True)
    msg(108, 0, 1, "GET /api/status (every 4 s)")
    msg(94, 1, 2, "mode and receiver health")
    msg(80, 2, 1, "SIM / LIVE status", dashed=True)
    msg(66, 1, 0, "status badge data", dashed=True)

    note = (
        "The browser periodically requests aircraft and health data. The backend decodes and normalizes the current feed, "
        "stores meaningful samples, and adds prediction rows only for aircraft that have temporarily disappeared. "
        "The response is then rendered as map markers, solid live tracks and dotted predicted paths."
    )
    draw_para(c, note, paragraph_style("SequenceNote", 9.2, 12.2), 24 * mm, 57 * mm, w - 48 * mm)
    c.setFont("TNR-Italic", 10)
    c.drawCentredString(w / 2, 29 * mm, "Figure 8: Sequence Diagram")
    page_number(c, "12")
    c.save()


def draw_prediction_pipeline(c: canvas.Canvas, y: float) -> None:
    w, _ = A4
    left = 23 * mm
    gap = 4 * mm
    box_w = 31 * mm
    box_h = 25 * mm
    labels = [
        "Recent stored\ntrajectory points",
        "Validate age\nand sample count",
        "Fit latitude, longitude\nand altitude vs. time",
        "Extrapolate to\ncurrent outage age",
        "Return predicted row\nfor dotted display",
    ]
    fills = ["#eef5fb", "#f5f5f5", "#fff3cd", "#ffe5e5", "#e8f5e9"]
    for i, (label, fill) in enumerate(zip(labels, fills)):
        x = left + i * (box_w + gap)
        c.setFillColor(colors.HexColor(fill))
        c.setStrokeColor(colors.HexColor("#315d7d"))
        c.roundRect(x, y, box_w, box_h, 2 * mm, fill=1, stroke=1)
        c.setFillColor(colors.black)
        for j, line in enumerate(label.split("\n")):
            c.setFont("TNR-Bold" if i == 2 else "TNR", 7.6)
            c.drawCentredString(x + box_w / 2, y + (15 - j * 5) * mm, line)
        if i < len(labels) - 1:
            x1 = x + box_w
            x2 = x + box_w + gap
            c.setStrokeColor(colors.HexColor("#315d7d"))
            c.line(x1 + 1, y + box_h / 2, x2 - 1, y + box_h / 2)
            c.line(x2 - 1, y + box_h / 2, x2 - 4, y + box_h / 2 + 2)
            c.line(x2 - 1, y + box_h / 2, x2 - 4, y + box_h / 2 - 2)
    c.setFont("TNR-Italic", 10)
    c.drawCentredString(w / 2, y - 8 * mm, "Figure 14: Trajectory Prediction Pipeline")


def create_algorithm_overview(path: Path) -> None:
    c = canvas.Canvas(str(path), pagesize=A4)
    w, h = A4
    x = 25 * mm
    width = w - 50 * mm
    y = h - 25 * mm
    c.setFont("TNR-Bold", 12)
    c.drawString(x, y, "APPENDIX G: ALGORITHM SELECTION AND COMPARISON")
    y -= 12 * mm
    c.setFont("TNR-Bold", 11)
    c.drawString(x, y, "G.1 Algorithm Used in the Project")
    y -= 7 * mm
    text = (
        "The implemented predictor uses <b>sliding-window ordinary least-squares linear regression</b>. "
        "For each aircraft, the latest stored samples are ordered by time. Separate straight lines are fitted for "
        "latitude, longitude and altitude: <i>value(t) = a + bt</i>. The fitted lines are evaluated at the last "
        "sample time plus the number of seconds for which the aircraft has been absent from the current feed."
    )
    y = draw_para(c, text, BODY, x, y, width) - 5 * mm

    facts = [
        "Input window: up to the latest 20 stored trajectory points; at least 4 valid points are required.",
        "Prediction condition: the aircraft must have been observed in the current session but be absent from the current live/replay snapshot.",
        "Prediction horizon: 1-180 seconds with the current configuration; older records are not displayed as predictions.",
        "Fallback: constant speed, heading and vertical rate are used when regression cannot be fitted.",
        "Safety against feedback: predicted rows are output-only and are not written back into the training tables.",
    ]
    c.setFont("TNR-Bold", 11)
    c.drawString(x, y, "Implemented constraints")
    y -= 6 * mm
    for item in facts:
        c.setFont("TNR", 10.3)
        c.drawString(x, y, "•")
        y = draw_para(c, item, paragraph_style("Fact", 10.3, 14), x + 6 * mm, y + 2.8 * mm, width - 6 * mm) - 2.2 * mm

    c.setFont("TNR-Bold", 11)
    c.drawString(x, y - 1 * mm, "G.2 Why This Method Was Selected")
    y -= 9 * mm
    rationale = (
        "For a local academic tracker, this method provides a useful balance: it is understandable, deterministic, "
        "fast enough to run during every API refresh, and able to smooth small variations across several observations. "
        "It needs no offline training dataset or GPU. Its main limitation is that a straight-line trend cannot model a "
        "sharp turn or a prolonged outage, which is why the prediction age is deliberately capped and the interface "
        "visually distinguishes predicted data from live data."
    )
    y = draw_para(c, rationale, BODY, x, y, width)
    draw_prediction_pipeline(c, 64 * mm)
    page_number(c, "68")
    c.save()


def create_algorithm_table(path: Path) -> None:
    size = landscape(A4)
    c = canvas.Canvas(str(path), pagesize=size)
    w, h = size
    c.setFont("TNR-Bold", 12)
    c.drawString(18 * mm, h - 18 * mm, "G.3 Comparison of Candidate Trajectory-Prediction Algorithms")
    intro = (
        "The comparison below explains the alternatives in plain language. The highlighted row is the method currently implemented."
    )
    draw_para(c, intro, paragraph_style("Intro", 9.5, 12), 18 * mm, h - 25 * mm, w - 36 * mm)

    data = [[
        Paragraph("Algorithm", TABLE_HEAD),
        Paragraph("How it works", TABLE_HEAD),
        Paragraph("Main advantage", TABLE_HEAD),
        Paragraph("Main limitation", TABLE_HEAD),
        Paragraph("Suitability for this project", TABLE_HEAD),
    ]]
    rows = [
        ("Last observation carried forward", "Keeps the last known position unchanged.", "Almost no computation and works with one record.", "A moving aircraft immediately appears stationary.", "Poor; useful only as an emergency display fallback."),
        ("Constant velocity and heading", "Moves the last position using current speed, heading and vertical rate.", "Very fast and does not require a training dataset.", "Does not smooth noise and cannot adapt well to turns.", "Good fallback; it is used when regression cannot be fitted."),
        ("Sliding-window linear regression (USED)", "Fits straight lines to recent latitude, longitude and altitude values over time.", "Simple, explainable, smooths several observations and is inexpensive.", "Assumes locally straight motion; accuracy drops during turns or long gaps.", "Best present choice for short gaps, local execution and limited data."),
        ("Kalman filter", "Recursively estimates position and velocity while modelling measurement noise.", "Strong noise handling and efficient real-time updates.", "Needs a motion model and careful covariance tuning.", "Strong next upgrade when receiver-noise statistics are available."),
        ("Polynomial regression", "Fits a curved time trend instead of a straight line.", "Can represent gentle curvature over a short window.", "Higher-order extrapolation can become unstable and overfit noise.", "Conditional; requires strict order and horizon controls."),
        ("Random forest / gradient boosting", "Learns nonlinear corrections from engineered flight-state features.", "Can model interactions that simple formulas miss.", "Requires representative labelled data and does not naturally model sequence continuity.", "Research option after a larger verified trajectory dataset is collected."),
        ("LSTM / GRU neural network", "Learns temporal patterns from complete sequences of prior flights.", "Can model complex route and manoeuvre patterns.", "Needs substantial training data, tuning and more computing power; less explainable.", "Future work for large-scale historical datasets, not the current local baseline."),
    ]
    for row in rows:
        data.append([Paragraph(cell, TABLE_TEXT) for cell in row])

    col_widths = [39 * mm, 58 * mm, 52 * mm, 55 * mm, 57 * mm]
    table = Table(data, colWidths=col_widths, repeatRows=1)
    style = [
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#d9e7f2")),
        ("GRID", (0, 0), (-1, -1), 0.55, colors.HexColor("#444444")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 5),
        ("RIGHTPADDING", (0, 0), (-1, -1), 5),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ("BACKGROUND", (0, 3), (-1, 3), colors.HexColor("#fff1b8")),
    ]
    table.setStyle(TableStyle(style))
    tw, th = table.wrap(w - 36 * mm, h - 55 * mm)
    table.drawOn(c, 18 * mm, h - 42 * mm - th)
    c.setFont("TNR-Italic", 9.5)
    c.drawCentredString(w / 2, 20 * mm, "Table 15: Trajectory Prediction Algorithm Comparison")
    page_number(c, "69", page_size=size)
    c.save()


def create_screenshot_page(path: Path, image_path: Path, page_num: int, figure_num: int, title: str, bullets: list[str], show_appendix_heading: bool = False) -> None:
    c = canvas.Canvas(str(path), pagesize=A4)
    w, h = A4
    x = 18 * mm
    usable_w = w - 36 * mm
    y = h - 20 * mm
    if show_appendix_heading:
        c.setFont("TNR-Bold", 12)
        c.drawString(x, y, "APPENDIX H: ANNOTATED USER INTERFACE SCREENSHOTS")
        y -= 10 * mm
    c.setFont("TNR-Bold", 11)
    c.drawString(x, y, title)
    y -= 8 * mm

    with Image.open(image_path) as im:
        iw, ih = im.size
    draw_w = usable_w
    draw_h = draw_w * ih / iw
    max_h = 112 * mm
    if draw_h > max_h:
        draw_h = max_h
        draw_w = draw_h * iw / ih
    image_x = (w - draw_w) / 2
    image_y = y - draw_h
    c.drawImage(str(image_path), image_x, image_y, draw_w, draw_h, preserveAspectRatio=True, mask="auto")
    c.setStrokeColor(colors.HexColor("#444444"))
    c.rect(image_x, image_y, draw_w, draw_h, stroke=1, fill=0)
    caption_y = image_y - 7 * mm
    c.setFont("TNR-Italic", 10)
    c.drawCentredString(w / 2, caption_y, f"Figure {figure_num}: {title}")
    text_y = caption_y - 10 * mm
    for item in bullets:
        c.setFont("TNR", 10.2)
        c.drawString(x, text_y, "•")
        text_y = draw_para(c, item, paragraph_style(f"Bullet{figure_num}", 10.2, 14), x + 6 * mm, text_y + 2.5 * mm, usable_w - 6 * mm) - 2.5 * mm
    page_number(c, str(page_num))
    c.save()


def create_supplement_pages(annotated: list[Path]) -> list[Path]:
    pages: list[Path] = []
    overview = TMP_DIR / "appendix-g-overview.pdf"
    comparison = TMP_DIR / "appendix-g-comparison.pdf"
    create_algorithm_overview(overview)
    create_algorithm_table(comparison)
    pages.extend([overview, comparison])

    screenshot_specs = [
        (70, 15, "Live Tracking Dashboard and Key Controls", [
            "The top badges distinguish live operation from simulation and show how recently receiver data arrived.",
            "The lower panel combines flight measurements, route endpoints and altitude history for the selected aircraft.",
        ]),
        (71, 16, "Aircraft Details, Route and Altitude History", [
            "Selecting a marker opens its callsign, ICAO address, altitude, speed and heading.",
            "Aircraft metadata, route lookup and altitude history are kept together so the operator can interpret the track without changing screens.",
        ]),
        (72, 17, "Simulation Mode Dashboard", [
            "Simulation TRUE means records are replayed from the configured captured log rather than received from the live TCP stream.",
            "The same decoder, storage, prediction and interface pipeline is exercised, which makes repeatable testing possible without receiver hardware.",
        ]),
        (73, 18, "Backend API Request Stream", [
            "HTTP 200 responses confirm that the aircraft and status endpoints are serving data successfully.",
            "The browser requests aircraft data every second and health/mode information every four seconds; repeated log lines are therefore expected during operation.",
        ]),
        (74, 19, "Predicted Trajectory and Dotted Path", [
            "PRED labels and the prediction-age field explicitly separate extrapolated aircraft states from currently received observations.",
            "The red dotted line shows the predicted continuation; solid blue history represents received trajectory samples.",
        ]),
    ]
    for i, (page_num, fig_num, title, bullets) in enumerate(screenshot_specs):
        out = TMP_DIR / f"appendix-h-{page_num}.pdf"
        create_screenshot_page(out, annotated[i], page_num, fig_num, title, bullets, show_appendix_heading=i == 0)
        pages.append(out)
    return pages


def merge_overlay(page, overlay_path: Path) -> None:
    overlay = PdfReader(str(overlay_path)).pages[0]
    page.merge_page(overlay)


def create_flattened_replacement(source_index: int, overlay_path: Path, output_path: Path) -> None:
    """Flatten one original page plus its overlay to avoid cross-reader resource collisions."""
    source_reader = PdfReader(str(SOURCE_PDF))
    page = source_reader.pages[source_index]
    overlay_reader = PdfReader(str(overlay_path))
    page.merge_page(overlay_reader.pages[0])

    composite_pdf = TMP_DIR / f"composite-{source_index + 1}.pdf"
    single = PdfWriter()
    single.add_page(page)
    with composite_pdf.open("wb") as fh:
        single.write(fh)

    rendered = pdfium.PdfDocument(str(composite_pdf))[0].render(scale=2.75).to_pil()
    background = TMP_DIR / f"composite-{source_index + 1}.png"
    rendered.save(background, optimize=True)

    c = canvas.Canvas(str(output_path), pagesize=A4)
    c.drawImage(str(background), 0, 0, A4[0], A4[1], preserveAspectRatio=False)
    c.save()


def build_final_pdf(supplement_pages: list[Path]) -> None:
    reader = PdfReader(str(SOURCE_PDF))
    writer = PdfWriter()

    replacements = {
        7: TMP_DIR / "toc-page.pdf",
        8: TMP_DIR / "figures-page.pdf",
        9: TMP_DIR / "tables-page.pdf",
        12: TMP_DIR / "feasibility-page.pdf",
        21: TMP_DIR / "sequence-page.pdf",
        77: TMP_DIR / "references-page.pdf",
    }
    for index, page in enumerate(reader.pages):
        if index == 77:
            for supplement in supplement_pages:
                writer.add_page(PdfReader(str(supplement)).pages[0])
        if index in replacements:
            writer.add_page(PdfReader(str(replacements[index])).pages[0])
            continue
        writer.add_page(page)

    writer.add_metadata({
        "/Title": "Aircraft Tracking and Trajectory Prediction - Revised Final Project Report",
        "/Author": "Aparna Tiwari",
        "/Subject": "Revised college project report with algorithm comparison and annotated interface figures",
    })
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    with OUTPUT_PDF.open("wb") as fh:
        writer.write(fh)


def main() -> None:
    register_fonts()
    TMP_DIR.mkdir(parents=True, exist_ok=True)
    annotated = annotate_screenshots()
    create_feasibility_overlay(TMP_DIR / "feasibility-overlay.pdf")
    create_toc_overlay(TMP_DIR / "toc-overlay.pdf")
    create_list_figures_overlay(TMP_DIR / "figures-overlay.pdf")
    create_list_tables_overlay(TMP_DIR / "tables-overlay.pdf")
    create_reference_overlay(TMP_DIR / "references-overlay.pdf")
    create_sequence_page(TMP_DIR / "sequence-page.pdf")
    create_flattened_replacement(7, TMP_DIR / "toc-overlay.pdf", TMP_DIR / "toc-page.pdf")
    create_flattened_replacement(8, TMP_DIR / "figures-overlay.pdf", TMP_DIR / "figures-page.pdf")
    create_flattened_replacement(9, TMP_DIR / "tables-overlay.pdf", TMP_DIR / "tables-page.pdf")
    create_flattened_replacement(12, TMP_DIR / "feasibility-overlay.pdf", TMP_DIR / "feasibility-page.pdf")
    create_flattened_replacement(77, TMP_DIR / "references-overlay.pdf", TMP_DIR / "references-page.pdf")
    supplement_pages = create_supplement_pages(annotated)
    build_final_pdf(supplement_pages)
    print(OUTPUT_PDF)


if __name__ == "__main__":
    main()
