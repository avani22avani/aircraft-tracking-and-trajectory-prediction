from copy import deepcopy
from pathlib import Path
import sys

from docx import Document
from docx.text.paragraph import Paragraph


SOURCE = Path(sys.argv[1])
OUTPUT = Path(sys.argv[2])


CONTENT = {
    "1.1 General": [
        "Air traffic surveillance means continuously observing aircraft so that their identity, position and movement are known. This is important for safe air traffic control, route monitoring, airport operations and aviation research.",
        "Earlier aircraft surveillance mainly depended on radar. Primary radar detects reflections from an aircraft body, while secondary radar receives replies from an aircraft transponder. These systems are useful, but they require large ground infrastructure and may update the aircraft position only after a fixed scan interval.",
        "Automatic Dependent Surveillance-Broadcast, or ADS-B, is a modern method in which the aircraft broadcasts its own information. The aircraft uses onboard navigation sources to determine its position and then transmits details such as ICAO address, callsign, latitude, longitude, altitude, speed and heading.",
        "A ground receiver such as PingStation can receive ADS-B information and provide it to a computer system. The received data is not directly useful to a normal user until it is decoded, cleaned, stored and displayed in an understandable form.",
        "This project builds an aircraft tracking system around PingStation data. The system receives live or recorded aircraft data, decodes useful flight parameters, stores them in database and file formats, and displays aircraft movement on a web-based map.",
        "The project is also connected to Electronics and Communication Engineering because it involves radio reception, digital communication, message framing, telemetry decoding, computer networking and real-time data visualization.",
    ],
    "1.2 Motivation": [
        "The motivation for this project comes from the practical value of ADS-B data. Aircraft in the surrounding airspace continuously broadcast real information, and this data can be used to build a working surveillance and learning system.",
        "Raw PingStation or ADS-B data is difficult to understand directly because it may appear as hexadecimal messages, partial fields or continuous network data. A proper software pipeline is required to convert that raw data into meaningful aircraft information.",
        "The project was also selected because it combines classroom concepts with a real application. It uses communication-system ideas such as signal reception, data framing and decoding, and also uses software concepts such as APIs, databases, prediction and user-interface design.",
        "The output is easy to demonstrate because aircraft can be seen moving on the map with callsign, altitude, speed, route and prediction information. This makes the project useful for presentations, laboratory demonstrations and future academic work.",
        "Instead of simply using an existing aircraft tracking platform, this project implements the important internal blocks. This helps in understanding how the data travels from receiver output to decoded records, stored trajectory samples and final visual display.",
    ],
    "1.3 Objectives of the Project": [
        "The main objective of the project is to develop a live aircraft tracking system using ADS-B/PingStation data.",
        "The system should support both live stream mode and simulation mode. Live mode receives data from the PingStation stream, while simulation mode replays already recorded data for testing and demonstration.",
        "Another objective is to decode aircraft messages and extract important flight parameters such as ICAO address, callsign, position, altitude, speed, heading, vertical speed, squawk and timestamp.",
        "The project aims to preprocess the extracted values so that units and formats remain consistent. This makes the same data usable for display, database storage, Excel logging and future analysis.",
        "The system should store the latest aircraft state as well as historical trajectory samples. SQLite, CSV files and ICAO-wise Excel files are used so that the data can be checked and reused later.",
        "The user interface should show aircraft on an interactive map, provide search and selection options, display aircraft details, show route information, plot altitude trends and mark predicted aircraft clearly.",
        "The project also aims to include short-gap trajectory prediction. When an aircraft temporarily disappears from the feed, the system estimates its position for a short time and marks it as predicted.",
        "Finally, the project aims to package the application as a Windows executable so that it can be opened more easily during demonstrations.",
    ],
    "1.4 Scope of the Project": [
        "The scope of the project starts from the digital data received from the PingStation receiver and ends at the dashboard display and stored output files.",
        "The project includes data acquisition, frame separation, message decoding, preprocessing, latest-state storage, trajectory-history storage, aircraft lookup, route display, prediction and user-interface design.",
        "It also includes diagnostic logging. The system can store raw bytes, hexadecimal data, ASCII data, SQLite records, CSV exports and Excel workbooks for later verification.",
        "The project does not design the aircraft transmitter, ADS-B transponder or PingStation hardware. It uses the available receiver output and focuses on software processing and visualization.",
        "The current system is suitable for educational use, demonstrations and small-scale local tracking. It is not intended to replace certified air traffic control systems or safety-critical aviation equipment.",
        "Future improvements may include better receiver placement, more live testing, stronger prediction models, multi-receiver support and offline map support.",
    ],
    "1.5 Organisation of the Report": [
        "Chapter 1 introduces the project, its motivation, objectives and scope.",
        "Chapter 2 explains the background of aircraft surveillance, ADS-B standards, existing tracking platforms and trajectory-prediction methods.",
        "Chapter 3 describes the work carried out, including system architecture, decoding pipeline, database design, prediction algorithm, aircraft classification and user-interface design.",
        "Chapter 4 presents the testing approach, observed results, comparison with existing systems and discussion of performance.",
        "Chapter 5 gives the summary, conclusions, limitations and future scope. The appendices include sample logs, selected code snippets, glossary terms and user-interface screenshots.",
    ],
    "2.1 Evolution of Aircraft Surveillance Technology": [
        "Aircraft surveillance has developed in several stages. The earliest widely used method was Primary Surveillance Radar (PSR), where a ground radar sends radio pulses and detects the reflected signal from an aircraft.",
        "PSR is useful because it can detect aircraft without depending on cooperation from the aircraft. However, it cannot directly provide aircraft identity or detailed flight data, and its accuracy depends on radar range, antenna rotation and environmental conditions.",
        "Secondary Surveillance Radar (SSR) improved this by using an aircraft transponder. The ground station sends an interrogation, and the aircraft replies with coded information such as identity and altitude. Mode A, Mode C and later Mode S are part of this development.",
        "Mode S introduced selective addressing, where each aircraft has a unique 24-bit ICAO address. This reduced confusion between replies from multiple aircraft and allowed more structured data exchange.",
        "ADS-B is a further development in which the aircraft broadcasts its own position and movement data without waiting for interrogation. This allows faster updates and makes aircraft tracking possible with smaller receivers such as PingStation.",
        "For this project, ADS-B is useful because it provides real aircraft telemetry that can be decoded, stored and displayed using a local software system.",
    ],
    "2.2 ADS-B Standards and the Regulatory Landscape": [
        "ADS-B commonly uses the 1090 MHz Mode-S Extended Squitter link. The word squitter means that the aircraft sends the message automatically, without a direct request from the ground station.",
        "A typical ADS-B message contains the aircraft address and a message field that may carry identity, position, altitude or velocity information. Different messages may contain different parts of the aircraft state, so the software has to combine updates over time.",
        "Standards are important because aircraft, receivers and ground systems may be manufactured by different organizations. Common standards make sure that the transmitted message can be understood by different compatible systems.",
        "International and national aviation authorities define the rules for ADS-B operation, message format, performance and equipment requirements. These rules help maintain safe and interoperable surveillance across different airspaces.",
        "From an Electronics and Communication point of view, ADS-B is an example of a real digital communication system. It includes carrier frequency, modulation, message framing, error checking, addressing and receiver processing.",
        "In this project, the receiver hardware provides the message stream, and the software focuses on interpreting that stream and using it for tracking and analysis.",
    ],
    "2.3 Existing Open-Source and Commercial ADS-B Platforms": [
        "Several existing platforms already use ADS-B data. Open-source decoders such as dump1090 and readsb can receive and decode ADS-B messages when used with suitable radio hardware.",
        "These open-source tools are powerful because they expose decoded aircraft data and can be modified by developers. They are often used by hobbyists, researchers and aviation-data networks.",
        "Commercial and crowd-sourced platforms such as FlightAware and FlightRadar24 collect data from many receivers and show aircraft on large global maps. These platforms provide polished displays and wide coverage, but their internal processing is not fully visible to the user.",
        "The present project is smaller than these global systems, but it has an educational advantage. The complete local pipeline is visible: input stream, decoding, preprocessing, database storage, prediction and web display.",
        "This makes the project suitable for understanding how an aircraft tracking system works internally rather than only viewing the final map output.",
    ],
    "2.4 Approaches to Flight-Trajectory Prediction": [
        "Flight-trajectory prediction means estimating the future position of an aircraft using its recent movement. This is useful when the live feed has a short gap or when the system wants to show the likely direction of travel.",
        "The simplest method is constant-velocity prediction. It uses the latest position, speed and heading to project where the aircraft may be after a short time. This method is fast but can be affected by one noisy sample.",
        "Another approach is linear regression. It uses several recent position samples and fits a simple trend over time. This can reduce the effect of small errors and gives a smoother short-term prediction.",
        "More advanced methods include Kalman filters, Extended Kalman filters and machine-learning models such as LSTM or GRU networks. These can handle noise or complex motion better, but they require more data, tuning and validation.",
        "For this project, sliding-window linear regression was selected as the main method because it is simple, explainable and fast enough for real-time display. A constant-velocity method is kept as a fallback when enough samples are not available.",
        "This choice matches the educational aim of the project: the algorithm should work for short gaps and should also be easy to understand and explain in the report and presentation.",
    ],
    "2.5 Summary and Identification of the Research Gap": [
        "The literature and existing systems show that ADS-B tracking is already well established. However, many systems either focus only on decoding messages or provide a final polished display without explaining the internal processing steps.",
        "For student learning, it is useful to build a system where every stage is visible. This includes live or replay input, raw logging, frame decoding, field normalization, storage, prediction and dashboard display.",
        "The research gap addressed by this project is the need for a compact and understandable local ADS-B tracking system that combines Electronics and Communication concepts with practical software implementation.",
        "The project also adds features that are useful for demonstration, such as simulation/live switching, predicted aircraft labels, route display, altitude graph, aircraft photo/type lookup and Excel/SQLite logging.",
        "Thus, the project is not only a map display. It is a complete learning-oriented pipeline that shows how received aircraft data can be converted into useful, stored and visual information.",
    ],
}


def insert_paragraph_before(paragraph, text, template_xml):
    new_p = deepcopy(template_xml)
    paragraph._p.addprevious(new_p)
    new_para = Paragraph(new_p, paragraph._parent)
    set_paragraph_text(new_para, text)
    return new_para


def set_paragraph_text(paragraph, text):
    for run in paragraph.runs:
        run.text = ""
    if paragraph.runs:
        paragraph.runs[0].text = text
    else:
        paragraph.add_run(text)


def delete_paragraph(paragraph):
    element = paragraph._element
    element.getparent().remove(element)
    paragraph._p = paragraph._element = None


def is_target_heading(text):
    return text.strip() in CONTENT


def is_next_boundary(text):
    stripped = text.strip()
    return (
        stripped in CONTENT
        or stripped in {"CHAPTER 2", "LITERATURE REVIEW", "CHAPTER 3", "WORK CARRIED OUT"}
        or stripped.startswith("3.1 ")
    )


def replace_section(doc, heading_text, paragraphs):
    all_paragraphs = doc.paragraphs
    start = next(i for i, p in enumerate(all_paragraphs) if p.text.strip() == heading_text)
    end = len(all_paragraphs)
    for i in range(start + 1, len(all_paragraphs)):
        if is_next_boundary(all_paragraphs[i].text):
            end = i
            break

    old_body = all_paragraphs[start + 1 : end]
    template = next((p for p in old_body if p.text.strip()), all_paragraphs[start])
    template_xml = deepcopy(template._p)
    anchor = all_paragraphs[end] if end < len(all_paragraphs) else None

    for old in old_body:
        delete_paragraph(old)

    if anchor is not None:
        for text in paragraphs:
            insert_paragraph_before(anchor, text, template_xml)
    else:
        for text in paragraphs:
            doc.add_paragraph(text, style=template.style)


def main():
    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    doc = Document(SOURCE)
    for heading, paragraphs in CONTENT.items():
        replace_section(doc, heading, paragraphs)
    doc.save(OUTPUT)
    print(OUTPUT)


if __name__ == "__main__":
    main()
