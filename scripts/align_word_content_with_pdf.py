from __future__ import annotations

from copy import deepcopy
from pathlib import Path

import pdfplumber
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.shared import Inches, Pt


ROOT = Path(__file__).resolve().parents[1]
DOCX = ROOT / "output" / "docx" / "Final_Project_Report_Final_Reviewed.docx"
PDF = ROOT / "output" / "pdf" / "Final_Project_Report_Final_Reviewed.pdf"
BACKUP = ROOT / "output" / "docx" / "Final_Project_Report_Final_Reviewed_before_content_align.docx"


def copy_run_format(src, dst):
    dst.bold = src.bold
    dst.italic = src.italic
    dst.underline = src.underline
    dst.font.name = src.font.name
    dst.font.size = src.font.size


def set_paragraph_text(paragraph, text):
    fmt_run = paragraph.runs[0] if paragraph.runs else None
    for run in paragraph.runs:
        run.text = ""
    run = paragraph.runs[0] if paragraph.runs else paragraph.add_run()
    run.text = text
    if fmt_run is not None and run is not fmt_run:
        copy_run_format(fmt_run, run)


def find_para(doc: Document, text: str):
    for p in doc.paragraphs:
        if p.text.strip() == text:
            return p
    return None


def insert_para_before(anchor, text: str, style_name: str | None = None, bold: bool = False):
    new_p = deepcopy(anchor._p)
    anchor._p.addprevious(new_p)
    paragraph = anchor._parent.paragraphs[0].__class__(new_p, anchor._parent)
    for child in list(paragraph._p):
        paragraph._p.remove(child)
    if style_name:
        paragraph.style = style_name
    paragraph.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
    paragraph.paragraph_format.space_after = Pt(6)
    run = paragraph.add_run(text)
    run.font.name = "Times New Roman"
    run.font.size = Pt(12)
    run.bold = bold
    return paragraph


def insert_paras_before(anchor, items):
    for kind, text in items:
        if kind == "h1":
            insert_para_before(anchor, text, "Heading 1")
        elif kind == "h2":
            insert_para_before(anchor, text, "Heading 2")
        elif kind == "h3":
            insert_para_before(anchor, text, "Heading 3")
        elif kind == "bullet":
            p = insert_para_before(anchor, text, "List Bullet")
            p.paragraph_format.left_indent = Inches(0.25)
        else:
            insert_para_before(anchor, text)


def insert_table_before(doc: Document, anchor, rows, widths=None, font_size=10):
    table = doc.add_table(rows=len(rows), cols=len(rows[0]))
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
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
    anchor._p.addprevious(table._tbl)
    return table


def update_first_table_after(doc: Document, heading: str, rows):
    seen = False
    for block in doc.element.body:
        if block.tag.endswith("p"):
            text = "".join(t.text or "" for t in block.iter() if t.tag.endswith("t")).strip()
            if text == heading:
                seen = True
        elif seen and block.tag.endswith("tbl"):
            table = None
            for t in doc.tables:
                if t._tbl is block:
                    table = t
                    break
            if table is None:
                return
            while len(table.rows) < len(rows):
                table.add_row()
            for r, row in enumerate(rows):
                for c, value in enumerate(row):
                    if c < len(table.columns):
                        table.cell(r, c).text = str(value)
            return


def extract_raw_hex_lines(limit=70):
    lines = []
    with pdfplumber.open(str(PDF)) as pdf:
        for idx in [75, 76]:
            text = pdf.pages[idx].extract_text() or ""
            for line in text.splitlines():
                s = line.strip()
                if len(s) >= 24 and all(ch in "0123456789abcdefABCDEF" for ch in s):
                    lines.append(s)
                if len(lines) >= limit:
                    return lines
    return lines


def build_alignment_patch():
    doc = Document(str(DOCX))
    if not BACKUP.exists():
        BACKUP.write_bytes(DOCX.read_bytes())

    # Keep the existing formatting and update the visible table contents only.
    update_first_table_after(doc, "TABLE OF CONTENTS", [
        ["Section", "Page"],
        ["1. Objective", "1"],
        ["2. Requirement Analysis (SRS)", "1"],
        ["3. Requirement Specification", "5"],
        ["4. System Design (SDS)", "7"],
        ["5. Coding", "16"],
        ["6. Testing", "18"],
        ["7. User Interfaces", "19"],
        ["8. Conclusion and Future Scope", "20"],
        ["Appendix A: Configuration Summary", "20"],
        ["Appendix B: Verified Main Code Listings", "21"],
        ["Appendix C: Project File Inventory", "62"],
        ["Appendix D: Sample Output Records", "62"],
        ["Appendix E: Execution Steps", "64"],
        ["Appendix F: Sample Raw PingStation Log", "64"],
        ["Appendix G: Algorithm Selection and Comparison", "68"],
        ["Appendix H: User Interface Screenshots", "71"],
        ["References", "76-77"],
    ])

    old_to_new = {
        "3. SOFTWARE DESIGN SPECIFICATION": "4. SYSTEM DESIGN (SDS)",
        "3.1 Module Description": "4.4 Main System Modules",
        "3.2 Use-Case Descriptions": "2.9 Use-Case Descriptions",
        "4. SYSTEM DESIGN": "4.3 Database Design and Class Model",
        "5. CODING AND IMPLEMENTATION": "5. CODING",
        "7. RESULTS AND USER INTERFACE": "7. USER INTERFACES",
        "APPENDIX A: SOURCE CODE LISTING": "Appendix B: Verified Main Code Listings",
        "APPENDIX B: FILE INVENTORY AND EXECUTION NOTES": "Appendix C: Project File Inventory",
        "APPENDIX C: SAMPLE OUTPUTS": "Appendix D: Sample Output Records",
    }
    for p in doc.paragraphs:
        key = p.text.strip()
        if key in old_to_new:
            set_paragraph_text(p, old_to_new[key])

    design_anchor = find_para(doc, "4. SYSTEM DESIGN (SDS)")
    if design_anchor is not None:
        insert_paras_before(design_anchor, [
            ("h2", "2.8 Actors"),
        ])
        insert_table_before(doc, design_anchor, [
            ["Actor", "Role in the System"],
            ["Student / Operator", "Runs the application, monitors aircraft and exports data."],
            ["PingStation receiver", "Provides the structured live receiver stream used by the decoder."],
            ["Browser client", "Requests Flask APIs and renders the dashboard."],
            ["Storage module", "Stores current records, trajectory samples and export files."],
            ["Future analysis module", "Uses saved trajectories for comparative prediction research."],
        ], [1.6, 3.7], 9)
        insert_paras_before(design_anchor, [
            ("h1", "3. REQUIREMENT SPECIFICATION"),
            ("h2", "3.1 System Inputs"),
            ("bullet", "Captured PingStation hexadecimal replay file or live PingStation TCP stream."),
            ("bullet", "Configuration values from ADSBTracker.cfg and environment-variable overrides."),
            ("bullet", "Browser requests for aircraft, status, logging and trajectory endpoints."),
            ("h2", "3.2 System Outputs"),
            ("bullet", "Current aircraft JSON records with ICAO, callsign, position, altitude, heading and speed fields."),
            ("bullet", "Interactive browser map, selected-aircraft panels, route traces and altitude charts."),
            ("bullet", "SQLite records, CSV exports, ICAO-wise Excel files and status logs."),
            ("h2", "3.3 Assumptions and Constraints"),
            ("bullet", "The PingStation receiver supplies already-demodulated structured frames; raw RF demodulation is outside scope."),
            ("bullet", "Predictions are short-gap continuations only and are not certified flight-plan forecasts."),
            ("bullet", "OpenStreetMap tiles and optional aircraft-enrichment services require internet connectivity."),
            ("h2", "3.4 Data Dictionary"),
        ])
        insert_table_before(doc, design_anchor, [
            ["Field", "Meaning"],
            ["icaoAddress", "Unique ICAO24 aircraft address."],
            ["callsign", "Aircraft callsign when available."],
            ["latDD / lonDD", "Latitude and longitude in decimal degrees."],
            ["altitudeMm", "Altitude in millimetres after normalization."],
            ["headingDeg2", "Heading in degrees multiplied by 100."],
            ["speedCms", "Ground speed in centimetres per second."],
            ["isPredicted", "True only for short-gap prediction output rows."],
        ], [1.55, 3.75], 9)
        insert_paras_before(design_anchor, [("h2", "3.5 API Specification")])
        insert_table_before(doc, design_anchor, [
            ["Endpoint", "Purpose"],
            ["/api/aircraft", "Returns current aircraft and labelled predicted rows."],
            ["/api/status", "Returns feed mode, aircraft counts and receiver health."],
            ["/api/logging/status", "Reports database, CSV and workbook logging status."],
            ["/api/training/trajectory", "Returns stored trajectory points for analysis and export."],
        ], [1.8, 3.5], 9)

    coding_anchor = find_para(doc, "5. CODING")
    if coding_anchor is not None:
        # Add PDF-like subsection labels immediately after the main coding heading.
        next_anchor = find_para(doc, "6. TESTING") or coding_anchor
        insert_paras_before(next_anchor, [
            ("h2", "5.1 Backend API"),
            ("body", "The Flask backend exposes compact JSON endpoints for aircraft state, receiver status, logging status and trajectory retrieval. The browser refreshes these endpoints without requiring a full page reload."),
            ("h2", "5.2 Decoder and Preprocessing"),
            ("body", "The decoder reconstructs complete length-prefixed frames, extracts available aircraft fields and maintains an ICAO-keyed state snapshot. The preprocessor converts records to a stable schema and rejects unusable coordinates."),
            ("h2", "5.3 Storage and Export"),
            ("body", "SQLite stores one latest record per flight and a filtered trajectory history. The storage module also produces daily CSV files and ICAO-wise Excel-compatible files for later analysis."),
            ("h2", "5.4 User Interface"),
            ("body", "The dashboard displays aircraft markers, route traces, selected-aircraft details, feed status, simulation/live badges and altitude history charts."),
            ("h2", "5.5 Error Handling and Data Quality"),
            ("body", "Invalid coordinate records, repeated stale updates and missing optional fields are handled without stopping the tracking interface."),
            ("h2", "5.6 Configuration Parameters"),
        ])
        insert_table_before(doc, next_anchor, [
            ["Parameter", "Implemented Value"],
            ["Aircraft API refresh", "1 second"],
            ["Status refresh", "4 seconds"],
            ["Live stale limit", "45 seconds"],
            ["Simulation stale limit", "12 seconds"],
            ["Regression window", "4 to 20 samples"],
            ["Prediction age limit", "180 seconds"],
        ], [2.2, 3.1], 9)

    appendix_anchor = find_para(doc, "Appendix B: Verified Main Code Listings")
    if appendix_anchor is not None:
        insert_paras_before(appendix_anchor, [
            ("h1", "Appendix A: Configuration Summary"),
            ("body", "The report PDF includes a compact configuration appendix before the code listings. The values below summarize the runtime settings most relevant to replay demonstration, logging and prediction."),
        ])
        insert_table_before(doc, appendix_anchor, [
            ["Setting", "Purpose"],
            ["SIMULATION_MODE", "Selects replay demonstration or live receiver mode."],
            ["PINGSTATION_LISTEN_PORT", "Configures the TCP port used for live receiver input."],
            ["FLIGHT_DB_MIN_INTERVAL_SEC", "Controls minimum time between stored trajectory samples."],
            ["FLIGHT_PREDICTION_MAX_AGE_SEC", "Bounds the maximum age of prediction output."],
            ["FLIGHT_ICAO_EXCEL_ENABLED", "Enables ICAO-wise Excel-compatible export files."],
        ], [2.25, 3.1], 9)

    algo_anchor = find_para(doc, "APPENDIX G: ALGORITHM SELECTION AND COMPARISON")
    if algo_anchor is not None:
        insert_paras_before(algo_anchor, [
            ("h1", "Appendix E: Execution Steps"),
            ("body", "1. Open the project folder adsb_tracker_2."),
            ("body", "2. Activate the Python environment or use the configured Python interpreter."),
            ("body", "3. Confirm that SIMULATION_MODE is true for replay-based demonstration."),
            ("body", "4. Run the Flask application using python app.py."),
            ("body", "5. Open http://localhost:5005 in a browser and inspect aircraft markers, route traces and exports."),
            ("h1", "Appendix F: Sample Raw PingStation Log"),
            ("body", "The following excerpt is taken from the bundled PingStation replay file. Each line is one captured payload represented as hexadecimal text and is passed through the same decoder path as live TCP bytes."),
        ])
        for line in extract_raw_hex_lines(45):
            p = insert_para_before(algo_anchor, line)
            p.paragraph_format.line_spacing = Pt(9)
            p.paragraph_format.space_after = Pt(0)
            for run in p.runs:
                run.font.name = "Courier New"
                run.font.size = Pt(7.5)

    refs_anchor = find_para(doc, "REFERENCES")
    if refs_anchor is not None:
        insert_paras_before(refs_anchor, [
            ("h1", "Appendix H: User Interface Screenshots"),
            ("body", "The user-interface screenshots appear in Section 7 with callouts for live/simulation status, selected aircraft details, predicted paths and route/altitude views. They are retained in the Word document for presentation use while keeping the main report format consistent."),
        ])

    ui_heading = find_para(doc, "7. USER INTERFACES")
    if ui_heading is not None:
        conclusion_anchor = appendix_anchor or refs_anchor
        if conclusion_anchor is not None:
            insert_paras_before(conclusion_anchor, [
                ("h1", "8. CONCLUSION AND FUTURE SCOPE"),
                ("h2", "8.1 Conclusion"),
                ("body", "The project demonstrates that a PingStation receiver and a single-host Python application can provide a practical local aircraft-tracking and logging platform. The same decoder supports live and replay modes, and the sliding-window linear-regression predictor provides an explainable baseline for short receiver gaps."),
                ("h2", "8.2 Future Scope"),
                ("bullet", "Quantitatively evaluate prediction error over multiple replay sessions and gap durations."),
                ("bullet", "Compare the current regression baseline with Kalman filtering and later machine-learning models."),
                ("bullet", "Improve aircraft metadata enrichment and offline map support for field demonstrations."),
            ])

    doc.save(str(DOCX))


if __name__ == "__main__":
    build_alignment_patch()
    print(DOCX)
