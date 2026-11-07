from datetime import datetime
from pathlib import Path

from docx import Document
from docx.enum.section import WD_SECTION_START
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt


ROOT = Path(__file__).resolve().parents[1]
REPORTS_DIR = ROOT / "reports"
OUT_PATH = REPORTS_DIR / "adsb_tracker_project_report_guideline_format.docx"


def set_run_font(run, size=12, bold=False, italic=False):
    run.font.name = "Times New Roman"
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "Times New Roman")
    run.font.size = Pt(size)
    run.bold = bold
    run.italic = italic


def set_paragraph_spacing(paragraph, line_spacing=1.5, after=0):
    fmt = paragraph.paragraph_format
    fmt.line_spacing = line_spacing
    fmt.space_after = Pt(after)


def add_text(doc, text="", bold=False, italic=False, size=12, align=None):
    p = doc.add_paragraph()
    if align is not None:
        p.alignment = align
    set_paragraph_spacing(p)
    r = p.add_run(text)
    set_run_font(r, size=size, bold=bold, italic=italic)
    return p


def add_heading(doc, text, level=1):
    p = doc.add_paragraph()
    set_paragraph_spacing(p, after=6)
    r = p.add_run(text)
    size = 14 if level == 1 else 12
    set_run_font(r, size=size, bold=True)
    return p


def add_body(doc, text):
    return add_text(doc, text)


def add_caption(doc, text, above=False):
    p = add_text(doc, text, italic=True)
    if above:
        p.paragraph_format.space_after = Pt(4)
    else:
        p.paragraph_format.space_before = Pt(4)
    return p


def add_page_number(paragraph):
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = paragraph.add_run()
    fld_char_begin = OxmlElement("w:fldChar")
    fld_char_begin.set(qn("w:fldCharType"), "begin")
    instr_text = OxmlElement("w:instrText")
    instr_text.set(qn("xml:space"), "preserve")
    instr_text.text = "PAGE"
    fld_char_end = OxmlElement("w:fldChar")
    fld_char_end.set(qn("w:fldCharType"), "end")
    run._r.append(fld_char_begin)
    run._r.append(instr_text)
    run._r.append(fld_char_end)
    set_run_font(run)


def set_page_number_format(section, fmt, start=None):
    sect_pr = section._sectPr
    pg_num_type = sect_pr.find(qn("w:pgNumType"))
    if pg_num_type is None:
        pg_num_type = OxmlElement("w:pgNumType")
        sect_pr.append(pg_num_type)
    pg_num_type.set(qn("w:fmt"), fmt)
    if start is not None:
        pg_num_type.set(qn("w:start"), str(start))


def configure_section(section, with_page_number=False, fmt=None, start=None):
    section.page_height = Inches(11.69)
    section.page_width = Inches(8.27)
    section.top_margin = Inches(1)
    section.bottom_margin = Inches(1)
    section.right_margin = Inches(1)
    section.left_margin = Inches(1.5)
    section.header_distance = Inches(0.5)
    section.footer_distance = Inches(0.5)

    section.header.is_linked_to_previous = False
    section.footer.is_linked_to_previous = False

    if with_page_number:
        footer = section.footer.paragraphs[0]
        footer.clear()
        add_page_number(footer)
        if fmt:
            set_page_number_format(section, fmt, start=start)
    else:
        section.footer.paragraphs[0].clear()


def add_section_break(doc, fmt=None, start=None, with_page_number=True):
    section = doc.add_section(WD_SECTION_START.NEW_PAGE)
    configure_section(section, with_page_number=with_page_number, fmt=fmt, start=start)
    return section


def add_manual_toc(doc):
    add_heading(doc, "TABLE OF CONTENTS")
    rows = [
        ("List of Figures", "ii"),
        ("List of Tables", "iii"),
        ("Abbreviations", "iv"),
        ("Certificate", "v"),
        ("Acknowledgements", "vi"),
        ("Abstract", "vii"),
        ("CHAPTER 1: INTRODUCTION", "1"),
        ("CHAPTER 2: LITERATURE REVIEW", "2"),
        ("CHAPTER 3: WORK CARRIED OUT", "3"),
        ("CHAPTER 4: EXPERIMENTAL RESULT AND COMPARISION", "5"),
        ("CHAPTER 5: CONCLUSION, SUMMARY & SCOPE", "6"),
        ("References", "7"),
        ("Bibliography", "8"),
        ("Appendix", "9"),
    ]
    table = doc.add_table(rows=0, cols=2)
    table.style = "Table Grid"
    for title, page in rows:
        cells = table.add_row().cells
        cells[0].text = title
        cells[1].text = page
        for cell in cells:
            for paragraph in cell.paragraphs:
                set_paragraph_spacing(paragraph, line_spacing=1.5)
                for run in paragraph.runs:
                    set_run_font(run)


def add_front_matter(doc):
    add_manual_toc(doc)
    doc.add_page_break()

    add_heading(doc, "LIST OF FIGURES")
    add_body(doc, "Figure 3.1 System flow of the ADS-B tracker ........................................ 4")
    doc.add_page_break()

    add_heading(doc, "LIST OF TABLES")
    add_body(doc, "Table 3.1 Main modules of the ADS-B tracker ..................................... 4")
    doc.add_page_break()

    add_heading(doc, "ABBREVIATIONS")
    abbreviations = [
        "ADS-B - Automatic Dependent Surveillance-Broadcast",
        "API - Application Programming Interface",
        "CSV - Comma-Separated Values",
        "GPS - Global Positioning System",
        "ICAO - International Civil Aviation Organization",
        "TCP - Transmission Control Protocol",
    ]
    for item in abbreviations:
        add_body(doc, item)
    doc.add_page_break()

    add_heading(doc, "CERTIFICATE")
    add_body(
        doc,
        "This is to certify that the project report entitled \"ADS-B LIVE TRACKER USING PINGSTATION DATA\" "
        "is a record of work carried out by the student(s) under the guidance of the supervisor.",
    )
    add_body(doc, "Guide Name: ____________________")
    add_body(doc, "Department: ____________________")
    doc.add_page_break()

    add_heading(doc, "ACKNOWLEDGEMENTS")
    add_body(
        doc,
        "We express our sincere gratitude to our guide, department, and institute for their support and guidance during this project work.",
    )
    doc.add_page_break()

    add_heading(doc, "ABSTRACT")
    add_body(
        doc,
        "This project develops an ADS-B tracker that decodes Pingstation aircraft data, processes it, stores trajectory records, and displays live aircraft movement on a web dashboard.",
    )


def add_body_sections(doc):
    add_heading(doc, "CHAPTER 1: INTRODUCTION")
    add_heading(doc, "1.1 General", level=2)
    add_body(
        doc,
        "ADS-B is used for aircraft surveillance by broadcasting identity, position, altitude, and velocity information in real time [1].",
    )
    add_heading(doc, "1.2 Motivation", level=2)
    add_body(
        doc,
        "The project aims to convert raw aircraft broadcast data into a clear and usable tracking system for observation and later analysis.",
    )
    add_heading(doc, "1.3 Objectives", level=2)
    add_body(
        doc,
        "The objectives are to decode Pingstation data, visualize live aircraft movement, and save structured trajectory data for future prediction work.",
    )
    add_heading(doc, "1.4 Scope", level=2)
    add_body(
        doc,
        "The present scope covers data decoding, preprocessing, storage, export, and dashboard display for aircraft tracking.",
    )

    add_heading(doc, "CHAPTER 2: LITERATURE REVIEW")
    add_heading(doc, "2.1 ADS-B Based Tracking", level=2)
    add_body(
        doc,
        "ADS-B provides near real-time surveillance and improves shared situational awareness for aircraft tracking applications [1].",
    )
    add_heading(doc, "2.2 Web-Based Visualization", level=2)
    add_body(
        doc,
        "Web tools such as Flask and Leaflet support lightweight APIs and interactive map-based displays for monitoring systems [2], [3].",
    )
    add_heading(doc, "2.3 Local Data Storage", level=2)
    add_body(
        doc,
        "SQLite is suitable for compact local storage where structured trajectory records are needed without a separate database server [4].",
    )

    add_heading(doc, "CHAPTER 3: WORK CARRIED OUT")
    add_heading(doc, "3.1 Project Problem", level=2)
    add_body(
        doc,
        "Raw Pingstation messages are difficult to use directly because they must be decoded, cleaned, stored, and presented in a readable form.",
    )
    add_heading(doc, "3.2 Solution Methodology", level=2)
    add_body(
        doc,
        "The system decodes raw frames, normalizes aircraft fields, stores results, and serves them to a browser dashboard through Flask APIs.",
    )
    add_caption(doc, "Figure 3.1 System flow of the ADS-B tracker.")
    figure = add_text(
        doc,
        "Pingstation Data -> Decoder -> Preprocessor -> Storage/API -> Dashboard",
        align=WD_ALIGN_PARAGRAPH.CENTER,
    )
    figure.paragraph_format.space_after = Pt(6)

    add_heading(doc, "3.3 Main Modules", level=2)
    add_caption(doc, "Table 3.1 Main modules of the ADS-B tracker.", above=True)
    table = doc.add_table(rows=1, cols=2)
    table.style = "Table Grid"
    hdr = table.rows[0].cells
    hdr[0].text = "Module"
    hdr[1].text = "Purpose"
    rows = [
        ("Decoder", "Reads raw Pingstation frames."),
        ("Preprocessor", "Normalizes aircraft fields."),
        ("Storage", "Saves latest and trajectory data."),
        ("Dashboard", "Shows aircraft on a live map."),
    ]
    for row in rows:
        cells = table.add_row().cells
        cells[0].text = row[0]
        cells[1].text = row[1]
    for row in table.rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                set_paragraph_spacing(paragraph, line_spacing=1.5)
                for run in paragraph.runs:
                    set_run_font(run)

    add_heading(doc, "3.4 Interpretation of Output", level=2)
    add_body(
        doc,
        "The output includes live aircraft markers, latest aircraft records, daily trajectory CSV files, and exportable history for future model development.",
    )

    add_heading(doc, "CHAPTER 4: EXPERIMENTAL RESULT AND COMPARISION")
    add_heading(doc, "4.1 Test Setup", level=2)
    add_body(
        doc,
        "The current system was tested using replayed Pingstation log data and local web execution.",
    )
    add_heading(doc, "4.2 Results", level=2)
    add_body(
        doc,
        "The tracker successfully displayed aircraft positions, stored trajectory samples, and generated CSV outputs from the replay stream.",
    )
    add_heading(doc, "4.3 Comparison", level=2)
    add_body(
        doc,
        "Compared with raw log inspection, the developed system provides structured records, visualization, and easier reuse of aircraft data.",
    )

    add_heading(doc, "CHAPTER 5: CONCLUSION, SUMMARY & SCOPE")
    add_heading(doc, "5.1 Summary", level=2)
    add_body(
        doc,
        "A working ADS-B tracker pipeline has been developed from raw data decoding to web visualization and trajectory storage.",
    )
    add_heading(doc, "5.2 Conclusions", level=2)
    add_body(
        doc,
        "The current implementation confirms that Pingstation data can be converted into usable real-time and historical aircraft information.",
    )
    add_heading(doc, "5.3 Scope for Future Work", level=2)
    add_body(
        doc,
        "Future work may include live receiver testing, prediction models, richer analytics, and deployment improvements.",
    )

    add_heading(doc, "REFERENCES")
    references = [
        "[1] Federal Aviation Administration, \"Automatic Dependent Surveillance-Broadcast (ADS-B),\" Aug. 30, 2023. [Online]. Available: https://www.faa.gov/air_traffic/technology/adsb/",
        "[2] Pallets, \"Flask Documentation.\" [Online]. Available: https://flask.palletsprojects.com/",
        "[3] Leaflet, \"Leaflet API reference.\" [Online]. Available: https://leafletjs.com/reference",
        "[4] SQLite, \"About SQLite.\" [Online]. Available: https://www.sqlite.org/about.html",
    ]
    for ref in references:
        add_body(doc, ref)

    add_heading(doc, "BIBLIOGRAPHY")
    add_body(
        doc,
        "Additional detailed bibliography will be expanded after the literature survey is finalized.",
    )

    add_heading(doc, "APPENDIX")
    add_body(
        doc,
        "Appendix A may later include screenshots, sample datasets, configuration details, and selected source code excerpts.",
    )


def build_document():
    REPORTS_DIR.mkdir(exist_ok=True)
    doc = Document()
    style = doc.styles["Normal"]
    style.font.name = "Times New Roman"
    style._element.rPr.rFonts.set(qn("w:eastAsia"), "Times New Roman")
    style.font.size = Pt(12)

    configure_section(doc.sections[0], with_page_number=False)

    title = add_text(
        doc,
        "PROJECT REPORT",
        bold=True,
        size=15,
        align=WD_ALIGN_PARAGRAPH.CENTER,
    )
    title.paragraph_format.space_after = Pt(6)
    add_text(doc, "ON", bold=True, size=14, align=WD_ALIGN_PARAGRAPH.CENTER)
    add_text(
        doc,
        "ADS-B LIVE TRACKER USING PINGSTATION DATA",
        bold=True,
        size=14,
        align=WD_ALIGN_PARAGRAPH.CENTER,
    )
    add_text(doc, "", align=WD_ALIGN_PARAGRAPH.CENTER)
    add_text(doc, "to be submitted by", italic=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    add_text(doc, "AUTHOR NAME / ROLL NUMBER", bold=True, size=14, align=WD_ALIGN_PARAGRAPH.CENTER)
    add_text(doc, "", align=WD_ALIGN_PARAGRAPH.CENTER)
    add_text(doc, "Under the supervision of", italic=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    add_text(doc, "GUIDE NAME", bold=True, size=14, align=WD_ALIGN_PARAGRAPH.CENTER)
    add_text(doc, "AFFILIATION OF GUIDE", bold=True, size=14, align=WD_ALIGN_PARAGRAPH.CENTER)
    add_text(doc, "", align=WD_ALIGN_PARAGRAPH.CENTER)
    add_text(doc, "For the award of the degree", italic=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    add_text(doc, "of", italic=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    add_text(doc, "<YOUR_CLASS>", bold=True, size=14, align=WD_ALIGN_PARAGRAPH.CENTER)
    add_text(doc, "<YOUR_STREAM>", bold=True, size=14, align=WD_ALIGN_PARAGRAPH.CENTER)
    add_text(doc, "", align=WD_ALIGN_PARAGRAPH.CENTER)
    add_text(doc, "DEPARTMENT OF PHYSICAL SCIENCES", bold=True, size=14, align=WD_ALIGN_PARAGRAPH.CENTER)
    add_text(doc, "BANASTHALI VIDYAPITH", bold=True, size=14, align=WD_ALIGN_PARAGRAPH.CENTER)
    add_text(doc, "BANASTHALI - 304022", bold=True, size=14, align=WD_ALIGN_PARAGRAPH.CENTER)
    add_text(doc, datetime.now().strftime("%B-%Y").upper(), bold=True, size=14, align=WD_ALIGN_PARAGRAPH.CENTER)

    add_section_break(doc, fmt="lowerRoman", start=1, with_page_number=True)
    add_front_matter(doc)

    add_section_break(doc, fmt="decimal", start=1, with_page_number=True)
    add_body_sections(doc)

    doc.save(OUT_PATH)
    return OUT_PATH


if __name__ == "__main__":
    print(build_document())
