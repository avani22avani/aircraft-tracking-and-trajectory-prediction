import re
import shutil
import sys
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.shared import Inches
from docx.text.paragraph import Paragraph
from PIL import Image, ImageDraw, ImageFont


DOC_PATH = Path(sys.argv[1])
OUT_PATH = Path(sys.argv[2])
IMAGE_PATHS = [Path(p) for p in sys.argv[3:]]
ASSET_DIR = OUT_PATH.parent / "annotated_screenshots"


MOJIBAKE_REPLACEMENTS = {
    "Ã¢â‚¬â€": "—",
    "Ã¢â‚¬â€œ": "–",
    "Ã¢â‚¬â€˜": "‘",
    "Ã¢â‚¬Ëœ": "‘",
    "Ã¢â‚¬â„¢": "’",
    "Ã¢â‚¬Å“": "“",
    "Ã¢â‚¬Â": "”",
    "Ã¢â‚¬Â”": "—",
    "Ã¢â‚¬Â–": "–",
    "Ã¢â‚¬Â¢": "•",
    "Ã¢â€°Ë†": "≈",
    "Ã¢â€ â€™": "→",
    "Ã‚Âµ": "µ",
    "Ã‚Â°": "°",
    "Ã‚Â": " ",
    "ÃƒÂ©": "é",
    "ÃƒÂ¤": "ä",
    "Гўв‚¬вЂќ": "—",
    "Гўв‚¬Лњ": "‘",
    "Гўв‚¬в„ў": "’",
    "Гўв‚¬Е“": "“",
    "Гўв‚¬Вќ": "”",
    "Гўв‚¬Вў": "•",
    "Г‚Вµ": "µ",
    "Г‚В": " ",
    "â€¢": "•",
}


BODY_CONTENT = {
    "1.1 General": [
        "Air traffic surveillance means finding and following aircraft in the sky. Earlier systems mainly used radar, but modern aircraft can also send their own position and flight data using ADS-B.",
        "In ADS-B, the aircraft broadcasts details such as ICAO address, callsign, position, altitude, speed and heading. A ground receiver such as PingStation can receive this data and pass it to software for display and storage.",
        "This project builds a simple aircraft tracker that receives PingStation data, decodes useful flight information, stores it and shows aircraft on a live web map.",
    ],
    "1.2  Motivation": [
        "The project was selected because it connects Electronics and Communication topics with real software development. It includes radio reception, digital communication, message decoding, networking, databases and web visualization.",
        "The output is easy to understand because real aircraft can be seen moving on the map. This makes the system useful for demonstration, learning and future research work.",
        "Instead of using a ready-made tracker, the project implements the main parts of the pipeline so that the working of the system can be understood clearly.",
    ],
    "1.3  Objectives of the Project": [
        "The main objective is to develop a live aircraft tracking system using ADS-B/PingStation data.",
        "The system should decode aircraft messages, extract flight parameters, store trajectory samples and show aircraft movement on an interactive map.",
        "The project also aims to support live and simulation modes, route display, aircraft information lookup, predicted path display, Excel/SQLite logging and a packaged Windows application.",
    ],
    "1.4  Scope of the Project": [
        "The project covers the software part from PingStation output to dashboard display. It includes frame decoding, preprocessing, database storage, prediction and user-interface design.",
        "The present work does not design the aircraft transmitter or PingStation hardware. It uses the receiver output and focuses on processing, analysis and visualization of the received data.",
    ],
    "1.5  Organisation of the Report": [
        "Chapter 1 introduces the project. Chapter 2 gives the required background. Chapter 3 explains the implemented system and algorithms.",
        "Chapter 4 presents testing and results. Chapter 5 gives the conclusion, limitations and future scope. The appendices include screenshots and selected code snippets.",
    ],
    "2.1  Evolution of Aircraft Surveillance Technology": [
        "Aircraft surveillance started with primary radar, where a ground antenna detects reflections from aircraft. Later, secondary radar used aircraft transponders to send identity and altitude information.",
        "ADS-B is a newer method. Here the aircraft itself broadcasts its position and movement data. This improves update speed and makes tracking possible with smaller ground receivers.",
    ],
    "2.2  ADS-B Standards and the Regulatory Landscape": [
        "ADS-B messages are commonly transmitted on 1090 MHz using Mode-S Extended Squitter. The message contains fields such as aircraft address, position, altitude and velocity.",
        "Different aviation bodies define the rules for these messages so that aircraft and ground systems can work together safely.",
    ],
    "2.3  Existing Open-Source and Commercial ADS-B Platforms": [
        "Open-source tools such as dump1090 can decode ADS-B messages from radio receivers. Commercial and crowd-sourced services such as FlightAware and FlightRadar24 show aircraft on global maps.",
        "This project is smaller than those systems, but it is useful because the complete local pipeline is visible and can be modified for learning and experiments.",
    ],
    "2.4  Approaches to Flight-Trajectory Prediction": [
        "Trajectory prediction means estimating where an aircraft may move next. Simple methods use current speed and heading, while advanced methods use filters or machine-learning models.",
        "For this project, a sliding-window linear regression method is used because it is simple, fast and easy to explain. A constant-velocity method is also kept as a fallback.",
    ],
    "2.5  Summary and Identification of the Research Gap": [
        "Many existing trackers either focus only on decoding or provide polished results without showing the internal processing. This project fills a learning gap by showing reception, decoding, storage, prediction and display together.",
        "The system is suitable for an Electronics and Communication project because it links RF reception, digital data, computer networks and real-time visualization.",
    ],
    "3.1  System Overview and Requirements Analysis": [
        "The system takes aircraft data from PingStation in either live mode or simulation mode. In live mode, it receives a TCP stream. In simulation mode, it replays previously recorded data.",
        "The main requirements are: decode messages correctly, show aircraft on the map, identify predicted flights, store data for later use, and provide a simple interface for users.",
        "The dashboard also shows whether the system is in simulation mode or live mode, the number of tracked aircraft and the age of the latest data.",
    ],
    "3.2  The 1090 MHz Mode-S Extended-Squitter Physical Layer": [
        "The ADS-B signal is transmitted at 1090 MHz by aircraft transponders. The receiver hardware captures this radio signal and provides decoded or semi-decoded digital data to the software.",
        "In this project, the software starts from the PingStation output. Therefore, the focus is not on building RF hardware, but on understanding the received message stream and using it properly.",
    ],
    "3.2.1  Frequency Allocation and Link Budget Considerations": [
        "ADS-B reception depends on antenna position, aircraft altitude, distance and line-of-sight conditions. High-flying aircraft can be received from a longer distance than low-flying aircraft.",
        "The report uses this idea to explain why range changes with aircraft height and receiver location.",
    ],
    "3.2.2  Pulse-Position Modulation and the Preamble": [
        "ADS-B/Mode-S messages use pulse-based signalling. A known preamble helps the receiver detect where a message starts.",
        "The PingStation receiver handles the radio-level detection. The project software uses the digital message stream provided by the receiver.",
    ],
    "3.2.3  The DF17 Extended-Squitter Frame Structure": [
        "A DF17 message contains a downlink format field, aircraft address, message data and parity. The useful information can include identity, position or velocity.",
        "The decoder reads the available fields and updates the current state of the aircraft.",
    ],
    "3.2.4  Compact Position Reporting (CPR)": [
        "Aircraft position is encoded compactly in ADS-B messages. A receiver or decoder must convert this information back into latitude and longitude.",
        "In this project, the software handles supported coordinate layouts from the PingStation output and checks that the values are realistic.",
    ],
    "3.2.5  Antenna Considerations and Radiation Pattern": [
        "A good antenna improves reception. For aircraft tracking, an antenna with wide coverage around the receiver is useful because aircraft may appear from any direction.",
        "Height and clear surroundings are important because 1090 MHz signals mainly travel by line of sight.",
    ],
    "3.2.6  Layered View of the ADS-B Communication Stack": [
        "The system can be understood in layers: aircraft transmitter, radio channel, ground receiver, TCP/data stream, decoder, database and web dashboard.",
        "This layered view makes it easier to test and improve each part separately.",
    ],
    "3.3  Software Architecture": [
        "The software has three main parts: data acquisition, backend processing and frontend display.",
        "The backend is written in Python and Flask. It receives aircraft data, processes it, stores it and sends it to the browser through API endpoints.",
        "The frontend uses Leaflet for the map and Chart.js for graphs. It updates aircraft positions regularly without reloading the page.",
    ],
    "3.3.1  Module Responsibilities": [
        "Each file has a clear role. data_source.py handles live or replay input. pingstation_raw_decoder.py decodes raw frames. pingstation_preprocess.py cleans values. flight_data_store.py stores and predicts data. app.py provides APIs.",
        "The table in this section lists the main modules and their responsibilities.",
    ],
    "3.3.2  Control Flow: Main Acquisition Loop": [
        "The system repeatedly receives or replays data, decodes it, normalizes it, stores important samples and sends the latest aircraft list to the dashboard.",
        "The browser calls the aircraft API repeatedly, so the map stays updated while the backend keeps writing logs and database records.",
    ],
    "3.4.1  The PingStation Hexadecimal Log Stream": [
        "The PingStation data is saved in session folders. The system stores hexadecimal data, ASCII data and raw bytes so that received messages can be checked later.",
        "These logs are helpful during debugging because they show the exact input that was processed by the decoder.",
    ],
    "3.4.2  Frame Decoding": [
        "The decoder separates complete frames from the byte stream. It uses the length byte, skips impossible lengths and keeps incomplete data until the next chunk arrives.",
        "After separating a frame, the decoder extracts fields such as ICAO address, callsign, altitude, speed, heading, position, vertical speed and squawk when they are present.",
    ],
    "3.4.3  Validation and Normalisation (PingstationPreprocessor)": [
        "The preprocessor makes all values consistent. For example, altitude is converted to millimetres, speed to centimetres per second and heading to degrees multiplied by 100.",
        "It also standardizes timestamps, callsigns and ICAO addresses so that the storage and display layers receive clean data.",
    ],
    "3.5  Database Design (FlightTrajectoryStore)": [
        "The database stores two main types of data. The latest table keeps the current state of each aircraft. The trajectory table stores time-based samples for later analysis.",
        "The store also creates CSV files, per-ICAO Excel files and a master workbook. This makes it easy to inspect data flight by flight.",
    ],
    "3.6.1  Prediction Objective and Activation": [
        "Prediction is used when an aircraft disappears briefly from the live feed. Instead of removing it immediately, the system estimates its short-term position.",
        "Predicted aircraft are clearly marked with a PRED label so the user can separate estimated data from received data.",
    ],
    "3.6.2  Sliding-Window Ordinary Least-Squares Model": [
        "The prediction method looks at recent latitude, longitude and altitude samples. It fits a simple straight-line trend over time and extends that trend for a short duration.",
        "This method is easy to understand and fast enough for a live dashboard.",
    ],
    "3.6.3  Why the Regression Is More Stable Than Two-Point Extrapolation": [
        "Using only the last two points can be noisy because one bad message may change the result. Linear regression uses several recent points, so it smooths small errors.",
        "This makes the displayed predicted path more stable during short data gaps.",
    ],
    "3.6.4  Kinematic Fallback, Output and Recovery": [
        "If there are not enough samples for regression, the system uses speed and heading to estimate movement. This is called the kinematic fallback.",
        "When real data returns, the predicted marker is replaced by the received aircraft position.",
    ],
    "3.6.5  Comparison of Suitable Prediction Algorithms": [
        "Several algorithms can be used for prediction, such as constant velocity, Kalman filter, regression and neural networks.",
        "For this project, linear regression was selected because it needs little data, works quickly and is easy to explain in a student project.",
    ],
    "3.6.6  Justification for Selecting Sliding-Window Linear Regression": [
        "Sliding-window linear regression is suitable for short gaps because aircraft usually move smoothly over a few seconds.",
        "It is not the best method for sharp turns or long prediction times, but it is a good first model for this tracker.",
    ],
    "3.7  Aircraft Classification Logic": [
        "The system can look up aircraft information using the ICAO address. It may show aircraft type, registration, category and photo when the external service has data.",
        "This improves the dashboard because the user sees more than just position and speed.",
    ],
    "3.8  User-Interface Design": [
        "The dashboard is designed like a simple radar display. The left side lists aircraft, the map shows their location, and the lower panel shows details for the selected aircraft.",
        "Important features include search, live/simulation status, map markers, predicted aircraft labels, route information, altitude graph and photo display.",
        "Annotated screenshots in this section show the main parts of the interface.",
    ],
    "3.9  Development Tools, Technologies and Environment": [
        "The project uses Python, Flask, SQLite, HTML, CSS, JavaScript, Leaflet, Chart.js and PyInstaller.",
        "These tools were selected because they are lightweight and suitable for a local real-time tracking application.",
    ],
    "4.1  Test Setup and Data-Collection Methodology": [
        "Testing was done using both live PingStation data and recorded data replay. The browser was opened at localhost:5005 and the backend APIs were called repeatedly.",
        "Screenshots show both simulation mode and live mode. Terminal logs confirm that the browser was requesting /api/aircraft and /api/status successfully.",
    ],
    "4.2  Reception Performance and Range Analysis": [
        "Reception range depends on antenna height, aircraft altitude and line of sight. Higher aircraft were visible from a longer distance than low aircraft.",
        "The observed range values in the table should be treated as practical observations for this setup, not as a universal receiver limit.",
    ],
    "4.3  Message-Decoding Accuracy and Signal-Quality Analysis": [
        "The system was checked by replaying known logs and confirming that aircraft identifiers, positions and motion fields appeared correctly on the dashboard.",
        "Some messages may contain only partial information, so the decoder merges new fields with the previous known state of the same aircraft.",
    ],
    "4.4  Trajectory-Prediction Accuracy": [
        "Prediction was used only for short gaps in reception. The displayed predicted aircraft and dotted path help the user understand that the position is estimated.",
        "The current prediction is useful for demonstration, but a larger labelled dataset is needed for formal accuracy measurement.",
    ],
    "4.5  End-to-End Latency and System Throughput": [
        "The browser polls the backend at regular intervals. This simple method keeps the display updated and is easy to implement.",
        "For future improvement, WebSocket-based push updates can reduce delay and make the display feel more live.",
    ],
    "4.6  Comparison with Existing Systems": [
        "Compared with simple decoders, this project adds storage, prediction, route display, photo/type lookup and a complete dashboard.",
        "Compared with large commercial platforms, this project is smaller and local, but its internal working is open and easy to study.",
    ],
    "4.7  Discussion of Results": [
        "The results show that the system can receive or replay aircraft data, decode it, display aircraft on a map and save useful records.",
        "The main strengths are simple visualization, session logging, Excel/SQLite storage and understandable prediction logic. The main limitations are receiver coverage and the need for more validation data.",
    ],
    "5.1  Summary of Work Carried Out": [
        "A complete ADS-B aircraft tracker was developed. It includes data acquisition, decoding, preprocessing, storage, prediction, lookup and dashboard display.",
        "The system supports simulation mode, live mode, predicted aircraft display and user-interface features that make the output easy to understand.",
    ],
    "5.2  Conclusions": [
        "The project proves that PingStation aircraft data can be converted into a useful live tracking system.",
        "It also shows how Electronics and Communication concepts such as RF surveillance, digital messages and telemetry can be combined with software tools.",
    ],
    "5.3  Limitations of the Present Implementation": [
        "The system depends on receiver coverage, network availability for maps and aircraft lookup, and the quality of the received messages.",
        "The prediction model is meant for short gaps only and should not be treated as a certified navigation or safety tool.",
    ],
    "5.4  Scope for Future Work": [
        "Future work can include better receiver placement, more live testing, offline maps, WebSocket updates and stronger validation of decoded data.",
        "The prediction part can be improved using Kalman filtering, route-based models or machine-learning methods after collecting a larger dataset.",
    ],
}

BODY_CONTENT_NORMALIZED = {
    re.sub(r"\s+", " ", key).strip(): value for key, value in BODY_CONTENT.items()
}


def fix_text(text: str) -> str:
    if not text:
        return text
    previous = None
    fixed = text
    while fixed != previous:
        previous = fixed
        for bad, good in MOJIBAKE_REPLACEMENTS.items():
            fixed = fixed.replace(bad, good)
    fixed = fixed.replace("*real*", "real")
    return re.sub(r"[ \t]{2,}", " ", fixed).strip() if fixed.strip() else fixed


def set_para_text(paragraph, text):
    for run in paragraph.runs:
        run.text = ""
    if paragraph.runs:
        paragraph.runs[0].text = text
    else:
        paragraph.add_run(text)


def force_para_text(paragraph, text):
    p_pr = paragraph._p.pPr
    for child in list(paragraph._p):
        if child is not p_pr:
            paragraph._p.remove(child)
    paragraph.add_run(text)


def has_drawing(paragraph):
    return bool(paragraph._p.xpath(".//w:drawing"))


def is_section_heading(text):
    stripped = text.strip()
    return (
        stripped in BODY_CONTENT
        or re.match(r"^(CHAPTER [1-5]|INTRODUCTION|LITERATURE REVIEW|WORK CARRIED OUT|EXPERIMENTAL RESULTS AND COMPARISON|SUMMARY, CONCLUSION AND SCOPE FOR FUTURE WORK)$", stripped)
        or re.match(r"^\d+(\.\d+)*\s{1,}\S+", stripped)
    )


def simplify_body_paragraphs(doc):
    current_heading = None
    content_index = 0
    in_body = False
    for paragraph in doc.paragraphs:
        text = fix_text(paragraph.text.strip())
        if text == "CHAPTER 1":
            in_body = True
        if text == "REFERENCES":
            in_body = False
        if text in BODY_CONTENT_NORMALIZED:
            current_heading = text
            content_index = 0
            set_para_text(paragraph, text)
            continue
        if not in_body or current_heading is None:
            continue
        if not text:
            continue
        if is_section_heading(text) or text.startswith(("Figure ", "Fig.", "Table ", "Code Snippet")):
            continue
        if has_drawing(paragraph):
            continue
        replacements = BODY_CONTENT_NORMALIZED.get(current_heading, [])
        if content_index < len(replacements):
            set_para_text(paragraph, replacements[content_index])
        else:
            set_para_text(paragraph, "")
        content_index += 1


def insert_paragraph_after(paragraph, text="", style=None):
    new_p = OxmlElement("w:p")
    paragraph._p.addnext(new_p)
    new_para = Paragraph(new_p, paragraph._parent)
    if style is not None:
        new_para.style = style
    if text:
        new_para.add_run(text)
    return new_para


def insert_paragraph_before(paragraph, text="", style=None):
    new_p = OxmlElement("w:p")
    paragraph._p.addprevious(new_p)
    new_para = Paragraph(new_p, paragraph._parent)
    if style is not None:
        new_para.style = style
    if text:
        new_para.add_run(text)
    return new_para


def annotate_image(source: Path, output: Path, labels):
    with Image.open(source).convert("RGB") as img:
        draw = ImageDraw.Draw(img)
        w, h = img.size
        try:
            font = ImageFont.truetype("arial.ttf", max(20, int(w * 0.017)))
            small_font = ImageFont.truetype("arial.ttf", max(17, int(w * 0.014)))
        except OSError:
            font = ImageFont.load_default()
            small_font = font

        for label, box, target in labels:
            x, y = int(box[0] * w), int(box[1] * h)
            tx, ty = int(target[0] * w), int(target[1] * h)
            text_bbox = draw.textbbox((0, 0), label, font=font)
            tw = text_bbox[2] - text_bbox[0]
            th = text_bbox[3] - text_bbox[1]
            pad = 8
            rect = (x, y, x + tw + 2 * pad, y + th + 2 * pad)
            draw.rounded_rectangle(rect, radius=8, fill=(255, 238, 130), outline=(190, 30, 30), width=3)
            draw.text((x + pad, y + pad), label, fill=(15, 25, 35), font=font)
            start = (x + tw // 2 + pad, y + th + 2 * pad)
            draw.line((start[0], start[1], tx, ty), fill=(220, 40, 40), width=5)
            r = max(8, int(w * 0.007))
            draw.ellipse((tx - r, ty - r, tx + r, ty + r), fill=(220, 40, 40), outline=(255, 255, 255), width=3)

        output.parent.mkdir(parents=True, exist_ok=True)
        img.save(output)


def create_annotated_images():
    ASSET_DIR.mkdir(parents=True, exist_ok=True)
    specs = [
        (
            IMAGE_PATHS[0],
            "annotated_live_mode.png",
            [
                ("Simulation OFF", (0.72, 0.03), (0.83, 0.11)),
                ("Live feed age", (0.84, 0.15), (0.91, 0.11)),
                ("Selected aircraft", (0.54, 0.05), (0.58, 0.41)),
                ("Route + distance panel", (0.56, 0.86), (0.73, 0.93)),
                ("Altitude chart", (0.76, 0.70), (0.89, 0.82)),
            ],
        ),
        (
            IMAGE_PATHS[1],
            "annotated_aircraft_details.png",
            [
                ("Search aircraft", (0.17, 0.13), (0.31, 0.10)),
                ("Aircraft popup", (0.51, 0.08), (0.58, 0.32)),
                ("Photo lookup", (0.66, 0.64), (0.72, 0.78)),
                ("Flight details", (0.08, 0.73), (0.25, 0.89)),
                ("Altitude trend", (0.76, 0.70), (0.88, 0.79)),
            ],
        ),
        (
            IMAGE_PATHS[2],
            "annotated_simulation_mode.png",
            [
                ("Simulation TRUE", (0.70, 0.04), (0.91, 0.08)),
                ("Simulated feed", (0.78, 0.13), (0.84, 0.08)),
                ("Live traffic list", (0.02, 0.18), (0.08, 0.27)),
                ("Aircraft markers", (0.50, 0.13), (0.57, 0.26)),
                ("Select aircraft for details", (0.16, 0.76), (0.25, 0.80)),
            ],
        ),
        (
            IMAGE_PATHS[3],
            "annotated_terminal_api.png",
            [
                ("Browser polling aircraft API", (0.46, 0.05), (0.50, 0.18)),
                ("Status API check", (0.50, 0.31), (0.52, 0.31)),
                ("HTTP 200 = success", (0.63, 0.16), (0.56, 0.17)),
                ("Local Flask server", (0.04, 0.82), (0.09, 0.91)),
            ],
        ),
        (
            IMAGE_PATHS[4],
            "annotated_predicted_path.png",
            [
                ("Predicted aircraft", (0.02, 0.10), (0.08, 0.19)),
                ("Dotted predicted path", (0.14, 0.16), (0.34, 0.55)),
                ("Prediction badge", (0.18, 0.73), (0.15, 0.82)),
                ("Selected aircraft data", (0.25, 0.80), (0.44, 0.89)),
                ("Altitude graph", (0.76, 0.73), (0.88, 0.86)),
            ],
        ),
    ]
    outputs = []
    for source, name, labels in specs:
        output = ASSET_DIR / name
        annotate_image(source, output, labels)
        outputs.append(output)
    return outputs


def add_figure_list_rows(doc):
    if not doc.tables:
        return
    table = doc.tables[0]
    existing = "\n".join(cell.text for row in table.rows for cell in row.cells)
    new_rows = [
        ("Fig. 3.12", "Annotated live-mode dashboard showing live feed, selected aircraft, route panel and altitude graph"),
        ("Fig. 3.13", "Annotated selected-aircraft view showing popup, aircraft details, photo lookup and graph"),
        ("Fig. 3.14", "Annotated simulation-mode dashboard showing simulation toggle and simulated aircraft feed"),
        ("Fig. 3.15", "Annotated terminal output showing Flask API polling and successful HTTP responses"),
        ("Fig. 3.16", "Annotated predicted-aircraft view showing PRED labels and dotted predicted path"),
    ]
    for fig_no, caption in new_rows:
        if fig_no in existing:
            continue
        row = table.add_row()
        row.cells[0].text = fig_no
        row.cells[1].text = caption


def insert_screenshots(doc, image_outputs):
    heading = next((p for p in doc.paragraphs if fix_text(p.text.strip()) == "3.8 User-Interface Design"), None)
    next_heading = next((p for p in doc.paragraphs if fix_text(p.text.strip()) == "3.9 Development Tools, Technologies and Environment"), None)
    if next_heading is None and heading is None:
        return

    captions = [
        "Figure 3.12 Annotated live-mode dashboard. It shows live stream status, simulation disabled, selected aircraft details, route/distance panel and altitude chart.",
        "Figure 3.13 Annotated selected-aircraft view. It highlights the aircraft popup, photo lookup, flight details and altitude trend.",
        "Figure 3.14 Annotated simulation-mode dashboard. It shows Simulation: TRUE, simulated feed status, aircraft list and map markers.",
        "Figure 3.15 Annotated Flask terminal output. It shows repeated /api/aircraft and /api/status requests with successful HTTP 200 responses.",
        "Figure 3.16 Annotated predicted-aircraft view. It highlights PRED labels, dotted predicted path, selected aircraft data and altitude graph.",
    ]

    if next_heading is not None:
        intro = insert_paragraph_before(next_heading, "3.8.1  User Interface Screenshots", style=heading.style if heading is not None else None)
    else:
        intro = insert_paragraph_after(heading, "3.8.1  User Interface Screenshots", style=heading.style if heading is not None else None)
    cursor = intro
    note = insert_paragraph_after(cursor, "The following annotated screenshots show the important parts of the implemented user interface. These figures can also be used directly in the project presentation.")
    cursor = note
    for image_path, caption in zip(image_outputs, captions):
        pic_para = insert_paragraph_after(cursor, "")
        pic_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        pic_para.add_run().add_picture(str(image_path), width=Inches(5.7))
        cursor = pic_para
        cap_para = insert_paragraph_after(cursor, caption)
        cap_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        cursor = cap_para


def replace_appendix_b(doc):
    start = None
    end = None
    for idx, paragraph in enumerate(doc.paragraphs):
        text = fix_text(paragraph.text.strip())
        if text == "APPENDIX B — ILLUSTRATIVE CODE EXCERPT":
            start = idx
        elif start is not None and text == "APPENDIX C — GLOSSARY OF SELECTED TERMS":
            end = idx
            break
    if start is None:
        return

    snippets = [
        "Code Snippet B.1 Live aircraft API route",
        "The browser calls this route repeatedly. The route gets aircraft data, preprocesses it, stores it and returns JSON for the map.",
        '@app.route("/api/aircraft")\ndef aircraft():\n    data = get_aircraft_data()\n    raw_list = data.get("aircraft", [])\n    processed = [preprocessor.preprocess(ac) for ac in raw_list]\n    flight_store.ingest_many(processed)\n    return jsonify({"aircraft": processed})',
        "Code Snippet B.2 Frame separation logic",
        "This logic uses the first byte as the frame length. It skips invalid bytes and waits if a full frame has not arrived.",
        "def split_frames(buffer_bytes):\n    frames = []\n    idx = 0\n    while idx < len(buffer_bytes):\n        frame_len = buffer_bytes[idx]\n        if frame_len < 10 or frame_len > 120:\n            idx += 1\n            continue\n        if idx + frame_len > len(buffer_bytes):\n            break\n        frames.append(buffer_bytes[idx:idx + frame_len])\n        idx += frame_len\n    return frames, buffer_bytes[idx:]",
        "Code Snippet B.3 Short-gap prediction idea",
        "The prediction module uses recent samples. If enough points are available, it fits a simple trend; otherwise it falls back to speed and heading.",
        "if enough_recent_samples:\n    predicted_position = linear_regression(recent_lat_lon_alt)\nelse:\n    predicted_position = kinematic_update(last_position, speed, heading, age_sec)",
    ]

    paragraphs = doc.paragraphs
    available = paragraphs[start + 1 : end if end is not None else len(paragraphs)]
    for i, paragraph in enumerate(available):
        if i < len(snippets):
            set_para_text(paragraph, snippets[i])
        else:
            set_para_text(paragraph, "")


def fix_all_text(doc):
    for paragraph in doc.paragraphs:
        fixed = fix_text(paragraph.text)
        if fixed != paragraph.text:
            set_para_text(paragraph, fixed)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    fixed = fix_text(paragraph.text)
                    if fixed != paragraph.text:
                        set_para_text(paragraph, fixed)


def clean_appendix_toc_lines(doc):
    replacements = {
        "APPENDIX A": "APPENDIX A — SAMPLE RAW MESSAGE LOG\t53",
        "APPENDIX B": "APPENDIX B — ILLUSTRATIVE CODE EXCERPT\t54",
        "APPENDIX C": "APPENDIX C — GLOSSARY OF SELECTED TERMS\t55",
    }
    for paragraph in doc.paragraphs:
        text = paragraph.text
        for key, value in replacements.items():
            if key in text and "\t" in text and any(token in text for token in ("Ã", "Â", "â", "Г")):
                force_para_text(paragraph, value)
                break


def main():
    OUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    image_outputs = create_annotated_images()

    shutil.copy2(DOC_PATH, OUT_PATH)
    doc = Document(OUT_PATH)

    fix_all_text(doc)
    simplify_body_paragraphs(doc)
    add_figure_list_rows(doc)
    insert_screenshots(doc, image_outputs)
    replace_appendix_b(doc)
    fix_all_text(doc)
    clean_appendix_toc_lines(doc)

    doc.save(OUT_PATH)
    print(OUT_PATH)


if __name__ == "__main__":
    main()
