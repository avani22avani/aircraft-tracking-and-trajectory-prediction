from copy import deepcopy
from pathlib import Path
import sys
import zipfile

from lxml import etree


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W_NS}
W = f"{{{W_NS}}}"


def paragraph_text(paragraph):
    return "".join(paragraph.xpath(".//w:t/text()", namespaces=NS)).strip()


def set_paragraph_text(paragraph, text):
    paragraph_properties = paragraph.find("w:pPr", NS)
    first_run = paragraph.find("w:r", NS)
    run_properties = None
    if first_run is not None:
        source_properties = first_run.find("w:rPr", NS)
        if source_properties is not None:
            run_properties = deepcopy(source_properties)

    for child in list(paragraph):
        if child is not paragraph_properties:
            paragraph.remove(child)

    run = etree.SubElement(paragraph, W + "r")
    if run_properties is not None:
        run.append(run_properties)
    text_node = etree.SubElement(run, W + "t")
    if text.startswith(" ") or text.endswith(" ") or "  " in text:
        text_node.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    text_node.text = text


def find_paragraph(root, exact_text):
    for paragraph in root.xpath(".//w:body/w:p", namespaces=NS):
        if paragraph_text(paragraph) == exact_text:
            return paragraph
    raise ValueError(f"Paragraph not found: {exact_text}")


def replace_with_paragraphs(root, old_text, new_paragraphs):
    paragraph = find_paragraph(root, old_text)
    set_paragraph_text(paragraph, new_paragraphs[0])
    cursor = paragraph
    for text in new_paragraphs[1:]:
        new_paragraph = deepcopy(paragraph)
        set_paragraph_text(new_paragraph, text)
        cursor.addnext(new_paragraph)
        cursor = new_paragraph


def append_paragraphs_after(root, anchor_text, paragraphs):
    anchor = find_paragraph(root, anchor_text)
    cursor = anchor
    for text in paragraphs:
        new_paragraph = deepcopy(anchor)
        set_paragraph_text(new_paragraph, text)
        cursor.addnext(new_paragraph)
        cursor = new_paragraph


def set_cell_text(cell, text):
    paragraphs = cell.xpath("./w:p", namespaces=NS)
    if paragraphs:
        set_paragraph_text(paragraphs[0], text)
        for paragraph in paragraphs[1:]:
            cell.remove(paragraph)
    else:
        paragraph = etree.SubElement(cell, W + "p")
        set_paragraph_text(paragraph, text)


def update_module_table(root):
    tables = root.xpath(".//w:body/w:tbl", namespaces=NS)
    if len(tables) < 2:
        raise ValueError("Module table not found")

    table = tables[1]
    rows = table.xpath("./w:tr", namespaces=NS)
    module_rows = [
        ("Data source", "Replays captured logs or receives a live Pingstation TCP stream."),
        ("Raw capture", "Stores original binary bytes with corresponding hexadecimal and ASCII logs for communication analysis."),
        ("Decoder", "Separates length-prefixed frames and decodes ICAO, callsign, position, altitude, speed, track, vertical rate and squawk."),
        ("Preprocessor", "Validates identifiers and converts telemetry into consistent engineering units and timestamps."),
        ("Storage", "Maintains latest aircraft state and historical trajectory samples in SQLite, CSV and Excel formats."),
        ("Prediction", "Uses recent trajectory points for short-horizon linear-regression prediction during temporary feed loss."),
        ("Web application", "Provides Flask APIs for aircraft data, history, export, lookup and logging status."),
        ("Dashboard", "Displays searchable aircraft markers, tracks, flight details and altitude trends on an interactive map."),
        ("Desktop package", "Bundles the application as ADSBTracker.exe for execution on Windows."),
    ]

    template_row = rows[-1]
    while len(rows) - 1 < len(module_rows):
        new_row = deepcopy(template_row)
        table.append(new_row)
        rows.append(new_row)

    while len(rows) - 1 > len(module_rows):
        table.remove(rows.pop())

    for row, (module, purpose) in zip(rows[1:], module_rows):
        cells = row.xpath("./w:tc", namespaces=NS)
        set_cell_text(cells[0], module)
        set_cell_text(cells[1], purpose)


def update_document_xml(xml_bytes):
    parser = etree.XMLParser(remove_blank_text=False)
    root = etree.fromstring(xml_bytes, parser)

    introduction_before = [
        paragraph_text(p)
        for p in root.xpath(".//w:body/w:p", namespaces=NS)
        if paragraph_text(p)
        in {
            "CHAPTER 1: INTRODUCTION",
            "1.1 General",
            "1.2 Motivation",
            "1.3 Objectives",
            "1.4 Scope",
        }
        or paragraph_text(p).startswith(
            (
                "Air traffic monitoring is an important",
                "Automatic Dependent Surveillance- Broadcast",
                "This project is based on developing",
                "The project also includes data storage",
                "Raw ADS-B or Pingstation data is not",
                "The project aims to convert raw aircraft",
                "The objectives are to decode Pingstation",
                "The present scope covers data decoding",
            )
        )
    ]

    replace_with_paragraphs(
        root,
        'This project develops an ADS-B tracker that decodes Pingstation aircraft data, processes it, stores trajectory records, and displays live aircraft movement on a web dashboard.',
        [
            "Automatic Dependent Surveillance-Broadcast (ADS-B) is a cooperative surveillance technique in which aircraft broadcast identity, position and motion information for reception by ground stations and other suitably equipped users. This project develops a software-based ADS-B tracking and data-logging system around the digital output obtained from a Pingstation receiver.",
            "The application accepts captured data for simulation or a live TCP stream, separates and decodes binary aircraft frames, normalizes the recovered telemetry and displays aircraft on an interactive web map. The decoded parameters include ICAO address, callsign, latitude, longitude, altitude, ground speed, heading, vertical rate, squawk and source information.",
            "For analysis and future trajectory-prediction work, the system stores latest aircraft states and historical samples in SQLite, CSV and ICAO-wise Excel files. It also provides short-duration position prediction during temporary feed loss, aircraft-information lookup, data-export APIs and a packaged Windows executable.",
            "The project combines concepts from Electronics and Communication Engineering, including radio surveillance, digital communication, binary protocol decoding, telemetry processing, networking and real-time visualization. Testing with the available project data produced 17 latest-aircraft records, 2,142 trajectory samples, multiple session logs and ICAO-specific workbooks.",
        ],
    )

    replace_with_paragraphs(
        root,
        "ADS-B provides near real-time surveillance and improves shared situational awareness for aircraft tracking applications [1].",
        [
            "ADS-B is an important development in cooperative air-surveillance systems. An aircraft determines its own navigation state using onboard sensors and satellite navigation, and its transponder periodically broadcasts identity, position, altitude and velocity information. Unlike primary radar, the ground system does not estimate position from reflected radio energy; it receives data reported by the aircraft.",
            "From an Electronics and Communication Engineering perspective, the complete chain contains an airborne navigation source, digital message generation, a radio-frequency transmitter, propagation through the wireless channel, a ground receiver and a data-processing terminal. Mode S Extended Squitter, commonly called 1090ES, operates at 1090 MHz and is widely used for ADS-B Out [5].",
            "In this project, the Pingstation unit represents the receiving interface between the radio-surveillance environment and the software system. The application works on the binary feed supplied by the receiver and reconstructs useful aircraft states from successive messages. This makes the work relevant to digital communication, telemetry, protocol interpretation and surveillance-system integration.",
        ],
    )

    replace_with_paragraphs(
        root,
        "Web tools such as Flask and Leaflet support lightweight APIs and interactive map-based displays for monitoring systems [2], [3].",
        [
            "A tracking system must transfer rapidly changing telemetry from the processing layer to the operator without repeatedly reloading the complete interface. Flask is used to expose aircraft state, trajectory history, export and lookup functions through HTTP/JSON APIs [2]. The browser polls the aircraft endpoint and updates the displayed state.",
            "Leaflet provides the geographic map, aircraft markers, tooltips, tracks, route lines and map controls [3]. Chart.js is used for the selected aircraft's altitude history [7]. Together, these components convert communication-system output into a visual surveillance display that is easier to interpret than raw hexadecimal data.",
            "The software also uses TCP for the optional live Pingstation connection. TCP provides an ordered byte stream, while the decoder retains incomplete bytes between reads and separates complete length-prefixed messages. This is necessary because network packet boundaries do not always correspond to application-frame boundaries.",
        ],
    )

    replace_with_paragraphs(
        root,
        "SQLite is suitable for compact local storage where structured trajectory records are needed without a separate database server [4].",
        [
            "Aircraft broadcasts are time-dependent, and meaningful analysis requires both the latest state and a history of earlier positions. SQLite was selected because it is a self-contained, serverless database stored in a single local file [4]. The database uses separate tables for the latest aircraft matrix and historical trajectory samples.",
            "The storage layer reduces unnecessary repetition by saving a new trajectory point when sufficient time, distance, altitude, heading or speed change has occurred. Write-ahead logging supports continuous application access, while configurable retention prevents unlimited historical growth.",
            "CSV exports support simple inspection and machine-learning preparation. The project additionally creates one Excel workbook per ICAO address and a master workbook containing an index and aircraft-specific sheets. Raw communication evidence is retained separately as binary, hexadecimal and ASCII session logs.",
        ],
    )

    replace_with_paragraphs(
        root,
        "Raw Pingstation messages are difficult to use directly because they must be decoded, cleaned, stored, and presented in a readable form.",
        [
            "The main problem was to transform a continuous Pingstation byte stream into reliable aircraft records. The receiver output contains length-prefixed binary frames, and individual messages may carry only selected fields. Data can also arrive in incomplete chunks, contain noise bytes or temporarily disappear from the live feed.",
            "A second problem was the difference between communication-oriented values and user-oriented engineering units. Altitude may be represented in feet or millimetres, speed in knots or centimetres per second, heading in tenths or hundredths of a degree, and coordinates in binary floating-point form. These values must be interpreted consistently before visualization or storage.",
            "The system also had to maintain aircraft identity across sparse messages, avoid duplicate display records, preserve useful raw data for later verification, and handle temporary reception gaps. Therefore, the project required coordinated work in protocol decoding, signal-data interpretation, networking, data management and user-interface design.",
        ],
    )

    replace_with_paragraphs(
        root,
        "The system decodes raw frames, normalizes aircraft fields, stores results, and serves them to a browser dashboard through Flask APIs.",
        [
            "The application supports two input modes. In simulation mode it replays previously captured Pingstation hexadecimal data in configurable batches. In live mode it opens a TCP listener, by default on port 30003, receives receiver bytes continuously and writes session-wise raw binary, hexadecimal and ASCII files.",
            "The frame separator reads the first byte as the message length, rejects invalid lengths, preserves incomplete data and resynchronizes after noise. The decoder interprets compressed field flags and recovers ICAO address, callsign, altitude, ground speed, track, latitude, longitude, vertical speed and squawk. When a message contains only a subset of fields, it merges the new values with the last known aircraft state.",
            "The preprocessing stage validates identifiers and normalizes altitude, speed, heading, coordinates and timestamps. Records with unusable positions are removed, and duplicate records are resolved by selecting the row containing the richest valid telemetry.",
            "Processed records are written to a SQLite database and optional CSV and Excel outputs. Sampling thresholds based on elapsed time, travelled distance and changes in altitude, heading or speed reduce redundant trajectory points while preserving significant motion.",
            "Flask APIs supply live aircraft data, the latest training matrix, trajectory history, CSV export, logging status and aircraft information. The dashboard updates the map, searchable aircraft list, selected-flight panel, track line and altitude graph. External lookup results can provide manufacturer, type, registration, category and an available aircraft photograph.",
            "For a short loss of reception, the storage module fits linear regression to recent latitude, longitude and altitude samples. If sufficient samples are unavailable, a kinematic estimate based on speed and heading is used. Predicted points are marked as output-only and are not inserted into the training tables.",
            "The completed application is packaged through PyInstaller as ADSBTracker.exe, allowing the Windows build to run with its included dependencies [6]. Runtime-path handling separates packaged resources from writable logs and database files.",
        ],
    )

    replace_with_paragraphs(
        root,
        "Pingstation Data -> Decoder -> Preprocessor -> Storage/API -> Dashboard",
        [
            "Aircraft ADS-B transmission -> Pingstation receiver -> TCP/captured byte stream -> raw session logger -> frame decoder -> telemetry preprocessor -> SQLite/CSV/Excel storage -> Flask API -> live map and analysis dashboard"
        ],
    )

    replace_with_paragraphs(
        root,
        "The output includes live aircraft markers, latest aircraft records, daily trajectory CSV files, and exportable history for future model development.",
        [
            "The primary visual output is a live map containing aircraft-specific markers, callsign labels, heading orientation, movement tracks and predicted positions when applicable. Selecting an aircraft displays ICAO address, coordinates, altitude, speed, heading, vertical rate, squawk, source and available registration or type information.",
            "The analytical output includes an altitude-time graph, latest-aircraft matrix, trajectory-history API and timestamped CSV exports. The storage workspace contains the SQLite database, session-wise raw captures, separate ICAO Excel files and a master ICAO workbook.",
            "The packaged project contains evidence of repeated testing: 30 session directories, 124 Excel-format log files, a Windows executable, 17 latest-aircraft database rows and 2,142 stored trajectory samples. These values describe the supplied project archive and are not presented as formal performance benchmarks.",
        ],
    )

    replace_with_paragraphs(
        root,
        "The current system was tested using replayed Pingstation log data and local web execution.",
        [
            "Functional testing was carried out on a Windows system using captured Pingstation files and the local Flask server at port 5005. Simulation replay processed 20 entries per refresh by default, while stale simulated records were removed after the configured interval. The interface requested updated aircraft data approximately once per second.",
            "The test path covered raw-file parsing, frame separation, decoding, preprocessing, database insertion, Excel logging, API response and browser visualization. The live-receiver path is also implemented through a TCP listener on port 30003, although the supplied archive mainly provides replay and stored-session evidence.",
            "The database, exported CSV files, ICAO workbooks and generated ADSBTracker.exe were inspected as implementation outputs. Testing focused on end-to-end functional operation rather than laboratory measurement of receiver sensitivity, RF coverage, packet-error rate or prediction accuracy.",
        ],
    )

    replace_with_paragraphs(
        root,
        "The tracker successfully displayed aircraft positions, stored trajectory samples, and generated CSV outputs from the replay stream.",
        [
            "The decoder successfully recovered aircraft identifiers and available motion parameters from the supplied Pingstation messages. Successive sparse messages were merged so that position, altitude, speed and heading could remain available in the displayed aircraft state.",
            "The supplied SQLite database contains 17 latest-aircraft records and 2,142 historical trajectory samples. Example tracked identifiers include 801748, 407A3E, 801609 and 801657. Stored values include coordinates, altitude, horizontal speed, heading, vertical speed, timestamp and sample count.",
            "The logging system produced session-specific folders for raw data and ICAO workbooks. The archive contains 30 session directories and 124 Excel-format files, including aircraft-specific files and master workbooks. A trajectory CSV export and a packaged Windows executable were also generated.",
            "The user interface presents live and predicted aircraft on the map, supports callsign or ICAO search, displays selected-aircraft telemetry and plots altitude history. The application can enrich a valid ICAO address with aircraft type, registration, category and photo information when external lookup services are reachable.",
        ],
    )

    replace_with_paragraphs(
        root,
        "Compared with raw log inspection, the developed system provides structured records, visualization, and easier reuse of aircraft data.",
        [
            "Raw hexadecimal and ASCII logs are useful for communication debugging but require manual interpretation. The developed system automatically separates frames, associates fields with an ICAO identity, applies unit conversion and presents the result as readable telemetry. This substantially reduces the effort required to understand receiver output.",
            "A simple latest-value display would lose historical motion, whereas the implemented SQLite trajectory table and ICAO workbooks preserve time-ordered samples. Threshold-based storage also avoids recording every repeated update, giving a better balance between detail and file growth.",
            "During a brief feed interruption, retaining the last marker would show a stationary aircraft. The implemented regression or kinematic fallback provides a temporary estimated position and clearly marks it as predicted. However, no quantitative comparison with reference radar data has yet been performed, so prediction accuracy remains a subject for future validation.",
            "The desktop executable improves ease of use compared with launching individual Python modules. Its current limitations include dependence on network access for map tiles and aircraft enrichment, possible errors in incomplete receiver messages, and the absence of formal RF-range and packet-loss measurements.",
        ],
    )

    replace_with_paragraphs(
        root,
        "A working ADS-B tracker pipeline has been developed from raw data decoding to web visualization and trajectory storage.",
        [
            "The project has produced an end-to-end ADS-B tracking application that accepts Pingstation data, captures the communication stream, decodes aircraft messages, normalizes telemetry, stores trajectory history and displays aircraft through a browser-based surveillance interface.",
            "The work integrates Electronics and Communication concepts with software implementation. Its communication-related elements include ADS-B surveillance, receiver interfacing, TCP streaming, binary framing, field-flag interpretation, telemetry conversion and handling of intermittent data.",
            "Additional completed features include ICAO-wise Excel logging, a master aircraft workbook, data-export APIs, aircraft-information enrichment, short-horizon trajectory prediction and a packaged Windows executable.",
        ],
    )

    replace_with_paragraphs(
        root,
        "The current implementation confirms that Pingstation data can be converted into usable real-time and historical aircraft information.",
        [
            "The implementation demonstrates that a Pingstation output stream can be transformed into useful real-time and historical aircraft information through systematic frame decoding and state reconstruction.",
            "The combination of raw capture, structured storage and visualization is valuable because it supports three different needs: communication debugging, operator observation and later data analysis. Keeping predicted data separate from measured trajectory samples also preserves the integrity of the stored dataset.",
            "The project objectives of aircraft-data decoding, live visualization, trajectory storage and export have been achieved at a functional prototype level. The available database, session logs, Excel workbooks and executable provide evidence of successful integration.",
        ],
    )

    replace_with_paragraphs(
        root,
        "Future work may include live receiver testing, prediction models, richer analytics, and deployment improvements.",
        [
            "Future work should include controlled testing with a continuously connected Pingstation receiver. Receiver coverage, message rate, packet loss, stale-data behaviour and decoded coordinates should be compared with a trusted surveillance reference.",
            "The prediction module can be improved by training and evaluating suitable trajectory models on larger datasets. Accuracy should be reported using position, altitude and time-horizon errors, with separate evaluation for straight flight, turns, climb and descent.",
            "Further Electronics and Communication work may include RF-link assessment, antenna placement studies, received-signal-strength logging, interference analysis and support for standard ADS-B message formats closer to the receiver front end.",
            "Software improvements may include offline map support, authentication, configurable alerts, multi-receiver fusion, automatic backup, test automation and deployment on an embedded or edge-computing platform.",
        ],
    )

    replace_with_paragraphs(
        root,
        "Additional detailed bibliography will be expanded after the literature survey is finalized.",
        [
            "Federal Aviation Administration publications on ADS-B surveillance and equipment.",
            "International Civil Aviation Organization material on ADS-B implementation and surveillance applications.",
            "Official documentation for Flask, Leaflet, SQLite, Chart.js, openpyxl and PyInstaller.",
            "Project source code, captured Pingstation logs, SQLite database, CSV exports and ICAO Excel workbooks supplied with the project archive.",
        ],
    )

    replace_with_paragraphs(
        root,
        "Appendix A may later include screenshots, sample datasets, configuration details, and selected source code excerpts.",
        [
            "Appendix A: Main software files",
            "app.py contains the Flask application and API routes. data_source.py selects simulation or live input and creates session directories. pingstation_raw_decoder.py performs raw capture, frame separation and aircraft-message decoding. pingstation_preprocess.py normalizes telemetry. flight_data_store.py manages SQLite, CSV, Excel and prediction functions. index.html implements the dashboard.",
            "Appendix B: Important interfaces",
            "The application provides endpoints for live aircraft data, latest matrix data, aircraft information lookup, logging status, trajectory history and CSV export. Default local ports are 5005 for the web application and 30003 for the optional Pingstation TCP listener.",
            "Appendix C: Generated outputs",
            "Generated outputs include flight_training.db, trajectory CSV exports, raw binary/hexadecimal/ASCII session logs, one workbook per ICAO address, MASTER_ICAO.xlsx and dist/ADSBTracker.exe.",
        ],
    )

    append_paragraphs_after(
        root,
        '[4] SQLite, "About SQLite." [Online]. Available: https://www.sqlite.org/about.html',
        [
            '[5] Federal Aviation Administration, "Installation: ADS-B Out," [Online]. Available: https://www.faa.gov/air_traffic/technology/equipadsb/installation',
            '[6] PyInstaller Development Team, "PyInstaller Manual," [Online]. Available: https://pyinstaller.org/en/stable/',
            '[7] Chart.js Contributors, "Chart.js Documentation," [Online]. Available: https://www.chartjs.org/docs/latest/',
            '[8] International Civil Aviation Organization, "Automatic Dependent Surveillance-Broadcast (ADS-B)," [Online]. Available: https://www.icao.int/',
        ],
    )

    update_module_table(root)

    introduction_after = [
        paragraph_text(p)
        for p in root.xpath(".//w:body/w:p", namespaces=NS)
        if paragraph_text(p)
        in {
            "CHAPTER 1: INTRODUCTION",
            "1.1 General",
            "1.2 Motivation",
            "1.3 Objectives",
            "1.4 Scope",
        }
        or paragraph_text(p).startswith(
            (
                "Air traffic monitoring is an important",
                "Automatic Dependent Surveillance- Broadcast",
                "This project is based on developing",
                "The project also includes data storage",
                "Raw ADS-B or Pingstation data is not",
                "The project aims to convert raw aircraft",
                "The objectives are to decode Pingstation",
                "The present scope covers data decoding",
            )
        )
    ]
    if introduction_before != introduction_after:
        raise RuntimeError("Chapter 1 introduction changed unexpectedly")

    return etree.tostring(
        root,
        xml_declaration=True,
        encoding="UTF-8",
        standalone=True,
    )


def main():
    source_path = Path(sys.argv[1])
    output_path = Path(sys.argv[2])
    output_path.parent.mkdir(parents=True, exist_ok=True)

    with zipfile.ZipFile(source_path, "r") as source, zipfile.ZipFile(
        output_path, "w"
    ) as output:
        for item in source.infolist():
            data = source.read(item.filename)
            if item.filename == "word/document.xml":
                data = update_document_xml(data)
            output.writestr(item, data)

    print(output_path)


if __name__ == "__main__":
    main()
