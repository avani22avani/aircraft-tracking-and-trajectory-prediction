from copy import deepcopy
from pathlib import Path
import sys

from docx import Document
from docx.text.paragraph import Paragraph


SOURCE = Path(sys.argv[1])
OUTPUT = Path(sys.argv[2])


EXPANSIONS = {
    "1.1 General": [
        "In a normal aircraft tracking workflow, the receiver first captures the aircraft broadcast, the software interprets the received message, and the final information is shown to the user. This project follows the same idea on a smaller scale so that each stage can be understood clearly.",
        "The word surveillance in this project does not mean only drawing an aircraft icon on a map. It also includes identifying the aircraft, storing its movement over time, checking whether the feed is live or simulated, and presenting useful details such as route, speed, heading and altitude history.",
        "The use of ADS-B makes the project practical for a laboratory environment because aircraft data can be received from real traffic without needing access to airport radar systems. This makes the project suitable for Electronics and Communication students who want to connect radio systems with real-time software.",
    ],
    "1.2 Motivation": [
        "Another important motivation was the need for a system that could work even when live data was not available. The simulation mode allows recorded PingStation data to be replayed, so the project can still be demonstrated in a classroom, viva or presentation without depending completely on live aircraft traffic.",
        "The project also gives practical experience with debugging real data. Unlike fixed textbook examples, received aircraft messages may be incomplete, repeated or temporarily missing. Handling such cases makes the project closer to real engineering work.",
        "The dashboard gives immediate visual feedback. If the decoder, database or API is not working properly, the map and aircraft list make the problem visible. This feedback helped in testing and improving the system step by step.",
    ],
    "1.3 Objectives of the Project": [
        "A further objective is to keep the system modular. Each major task, such as data sourcing, decoding, preprocessing, storage, prediction and visualization, is handled by a separate part of the program. This makes the project easier to test and extend.",
        "The project also aims to preserve diagnostic data. Raw logs, hexadecimal logs and ASCII logs are useful for checking what was actually received from the stream. This is important when the displayed result needs to be verified later.",
        "Another objective is to make the project useful for both technical and non-technical viewers. The backend contains the detailed processing, while the dashboard explains the result visually through labels, route panels, graphs and status indicators.",
    ],
    "1.4 Scope of the Project": [
        "The scope includes a local single-receiver tracking system. It is designed to run on a computer connected to the receiver or to replay logs from the same project folder. The system is therefore suitable for demonstration and development work.",
        "The project includes basic prediction only for short data gaps. The predicted positions are shown to maintain continuity on the map, but they are clearly marked so that the user knows the aircraft is not currently being received as a fresh live update.",
        "The scope also includes data export and record keeping. This is useful because a tracking system becomes more valuable when the displayed information can also be saved and studied later.",
    ],
    "1.5 Organisation of the Report": [
        "The report first builds the background needed to understand ADS-B, then moves toward the implementation details. This order helps the reader understand why each software block is needed.",
        "The later chapters focus on results, limitations and future improvement. These sections show what the system can do at present and what can be improved if more time, data or hardware is available.",
    ],
    "2.1 Evolution of Aircraft Surveillance Technology": [
        "Primary radar was a major advancement because it could detect an aircraft by reflected radio energy. However, it gives limited information unless it is combined with other systems. It generally shows that a target exists, but not always who the aircraft is or what flight details it is broadcasting.",
        "Secondary radar improved the situation by using the aircraft transponder. The aircraft replies to the ground station with coded data, so the ground system can identify the aircraft more easily. This shows how communication between aircraft and ground equipment became a key part of surveillance.",
        "ADS-B changes the approach further by allowing the aircraft to broadcast information regularly. This reduces dependence on rotating radar scans and allows many receivers to hear the same message. Such a system is useful for distributed tracking, research networks and low-cost local receivers.",
        "For a student project, this evolution is important because it shows that aircraft tracking is not only a software problem. It is a complete communication chain that includes radio transmission, receiver hardware, digital messages, decoding and display.",
    ],
    "2.2 ADS-B Standards and the Regulatory Landscape": [
        "Mode-S Extended Squitter messages are structured so that receivers can identify the type of information carried in each message. Some messages carry identity, some carry position, and some carry velocity. A tracker must therefore combine multiple messages to build a complete aircraft state.",
        "The 24-bit ICAO address is an important part of this process. It acts like a unique aircraft identifier in the received data, allowing the software to update the correct aircraft record each time a new message arrives.",
        "The standards also define performance expectations for aircraft equipment. This is important because surveillance data is used in aviation operations, where reliability and interoperability matter.",
        "Although this project is not a certified aviation system, understanding the standards helps explain why the received message fields have a fixed structure and why the same receiver can observe aircraft from different airlines and regions.",
    ],
    "2.3 Existing Open-Source and Commercial ADS-B Platforms": [
        "Open-source platforms are useful for learning because their code and behaviour can be inspected. They show how ADS-B messages are received and decoded, and many of them can output aircraft data in JSON or other formats.",
        "Commercial systems usually provide a more polished user experience. They combine data from thousands of receivers, clean the data, add flight information and show global coverage. However, the user normally sees only the final result, not the complete processing pipeline.",
        "The proposed system stands between these two categories. It is not as large as a commercial network, but it includes more than a basic decoder. It adds local storage, prediction, route information, screenshots for presentation and a dashboard designed around the project requirements.",
        "This makes the project valuable academically because the internal logic is visible. A reader can understand how the receiver output becomes a live map, rather than treating the tracker as a black box.",
    ],
    "2.4 Approaches to Flight-Trajectory Prediction": [
        "Prediction is especially useful in a live display because received data may not arrive at perfectly regular intervals. A short delay in the feed can make the marker appear frozen or disappear, even though the aircraft is still moving in reality.",
        "A constant-velocity method is easy to implement because it uses the latest speed and heading. However, if the last heading is noisy or if the aircraft is turning, the predicted point may quickly move away from the real path.",
        "Linear regression reduces this problem by using several recent samples. It does not try to model complex aircraft behaviour, but it gives a smooth trend for a short time. This is appropriate for the present project because the prediction is only used for short outages.",
        "Kalman filters are commonly used in tracking systems because they can combine measurements with a motion model. They are a good future improvement, but they require careful tuning of noise values and state variables.",
        "Machine-learning methods may perform better when large historical datasets are available. At the current stage, the project has limited local data, so a simple and explainable method is more suitable than a large trained model.",
    ],
    "2.5 Summary and Identification of the Research Gap": [
        "The review shows that there is a need for an educational system that is more detailed than a simple map, but still easier to understand than a large commercial aircraft tracking network.",
        "This project addresses that need by keeping every stage available for explanation. The user can see the live or simulated feed, the decoded aircraft, the stored records, the predicted path and the dashboard output.",
        "The project also focuses on practical implementation. It does not stop at theory; it connects the ADS-B communication concept to actual code, logs, database records and user-interface screenshots.",
        "Because of this, the project can be used as a base for future work such as better prediction, multi-receiver comparison, receiver range testing or data-quality analysis.",
    ],
    "3.1 System Overview and Requirements Analysis": [
        "The complete system can be viewed as a pipeline. Data enters the system from a PingStation live stream or from recorded files. The backend processes this data and the frontend displays it to the user.",
        "The first requirement is reliable input handling. In live mode, the system must listen to a stream and process messages as they arrive. In simulation mode, it must replay saved data in a controlled way so testing remains repeatable.",
        "The second requirement is correct decoding. The software must separate frames, reject unusable data and extract fields only when they are available. This is important because not every received message contains every aircraft parameter.",
        "The third requirement is clear presentation. A user should be able to understand which aircraft are live, which are predicted, what route or direction they are following and whether the feed is running in simulation or live mode.",
        "The fourth requirement is storage. A tracker that only displays current aircraft is useful for monitoring, but a tracker that also stores data is useful for analysis, reporting and future model development.",
        "The fifth requirement is usability. The dashboard should not require the user to inspect logs manually. Status chips, aircraft lists, map markers and selected-aircraft panels make the output understandable quickly.",
        "The implemented system satisfies these requirements through separate modules for data input, decoding, preprocessing, database storage, prediction, aircraft lookup and dashboard display.",
    ],
    "3.2 The 1090 MHz Mode-S Extended-Squitter Physical Layer": [
        "The physical layer is the part of the system where aircraft information is carried through radio waves. ADS-B Out on 1090 MHz is widely used because many commercial aircraft already carry Mode-S transponders.",
        "The aircraft transmitter sends short digital messages at regular intervals. The receiver must detect these signals, recover the message bits and provide the decoded or semi-decoded output to the computer.",
        "In this project, the PingStation receiver handles the low-level radio reception. The software therefore starts from the receiver's output stream, but the report still discusses the physical layer because it explains why range, antenna placement and message quality matter.",
    ],
    "3.2.1 Frequency Allocation and Link Budget Considerations": [
        "The 1090 MHz frequency lies in the L-band region. Signals at this frequency usually travel in a line-of-sight manner, which means that obstacles, terrain and antenna height can strongly affect reception.",
        "Aircraft at high altitude are easier to receive from a long distance because the line of sight between the aircraft and receiver is less likely to be blocked. Low-altitude aircraft disappear sooner because the Earth, buildings or terrain can block the path.",
        "A simple link-budget view includes transmitter power, antenna gain, path loss, receiver sensitivity and cable loss. The project does not measure all of these values separately, but it uses the concept to explain differences in observed reception range.",
        "In practical terms, placing the receiver antenna higher and away from obstructions can improve the number of received aircraft and the continuity of the data stream.",
    ],
    "3.2.2 Pulse-Position Modulation and the Preamble": [
        "ADS-B/Mode-S messages use short pulses. The preamble is a known pulse pattern placed at the beginning of a message so that the receiver can detect that a valid message is starting.",
        "After detecting the preamble, the receiver reads the following bits using the expected timing pattern. This timing-based structure is why accurate radio reception and clock recovery are important at the hardware level.",
        "The software in this project does not demodulate raw radio samples. However, understanding the preamble and pulse structure helps explain why the receiver output may sometimes miss or corrupt messages when the signal is weak.",
        "When the received signal is poor, the software may receive fewer valid frames. Therefore, the later software stages must be able to handle missing updates gracefully.",
    ],
    "3.2.3 The DF17 Extended-Squitter Frame Structure": [
        "DF17 is one of the important message formats used for ADS-B. It includes fields that identify the message type, aircraft address, message content and parity information.",
        "The aircraft address allows the system to connect multiple messages to the same aircraft. This is necessary because one message may contain identity while another message may contain position or velocity.",
        "The message content is interpreted according to type codes. For example, one type of message may provide callsign, while another may provide airborne position or velocity.",
        "The project decoder focuses on recovering the aircraft state from the PingStation output. It stores the best available current values and updates them whenever a new message contains better or newer information.",
    ],
    "3.2.4 Compact Position Reporting (CPR)": [
        "Compact Position Reporting is used in ADS-B to send position efficiently. Instead of sending latitude and longitude in a simple text form, the aircraft sends encoded values that must be resolved by the receiver or decoder.",
        "In many practical receiver outputs, some of this work may already be handled by the receiver firmware. Still, the software must check that the decoded latitude and longitude are within valid geographic limits.",
        "The project includes plausibility checks for position values. If a position is zero or outside a valid range, the row is not useful for display and should not be treated as a real aircraft location.",
    ],
    "3.2.5 Antenna Considerations and Radiation Pattern": [
        "The antenna is one of the most important parts of a receiving setup. Even a good decoder cannot recover messages that the receiver never receives properly.",
        "An antenna used for aircraft tracking should ideally have wide coverage around the receiver because aircraft can appear from different directions. A clear view of the sky is also helpful.",
        "Cable quality and connector condition can also affect received signal strength. Losses between the antenna and receiver reduce the signal available to the receiver.",
        "In future work, antenna placement can be studied more carefully by comparing aircraft count, maximum range and message continuity at different locations.",
    ],
    "3.2.6 Layered View of the ADS-B Communication Stack": [
        "A layered view makes the system easier to understand. The lowest layer is the radio transmission from the aircraft. The next layer is receiver hardware, which converts the signal into digital output.",
        "Above that, the software separates frames and decodes fields. The preprocessing layer then converts values into consistent units and formats.",
        "The storage layer keeps both latest states and historical samples. Finally, the application layer provides APIs and the web dashboard used by the user.",
        "If a problem occurs, this layered structure helps locate it. For example, no aircraft on the map may be caused by antenna problems, receiver-stream problems, decoder problems, API problems or browser display problems.",
    ],
    "3.3 Software Architecture": [
        "The software architecture was designed to keep responsibilities separate. This makes the system easier to debug, because each file has a clear purpose.",
        "The data-source module decides whether data comes from a live PingStation stream or a recorded simulation file. This allows the rest of the application to use a common data format without worrying about the input mode.",
        "The decoder module focuses on the structure of incoming messages. It separates frames, extracts aircraft fields and maintains the latest known state for each aircraft.",
        "The preprocessing module standardizes the output. It converts values into consistent units and ensures that fields such as ICAO address, callsign and timestamp are present.",
        "The storage module handles SQLite, CSV and Excel logging. It also contains the prediction logic because prediction depends on recent stored trajectory samples.",
        "The Flask application connects backend data to the browser. It provides API endpoints for live aircraft, status, trajectory history, export and aircraft information lookup.",
        "The dashboard is responsible for user interaction. It shows the map, list of aircraft, flight details, route information, altitude graph and live/simulation status.",
    ],
    "3.3.1 Module Responsibilities": [
        "The module division also supports future improvement. For example, a new prediction method can be added in the storage module without changing the frontend map logic.",
        "Similarly, a different receiver or data source could be added by modifying the data-source layer while keeping the dashboard and database mostly unchanged.",
        "This modular design is useful in engineering projects because it reduces dependency between unrelated parts of the system.",
    ],
    "3.3.2 Control Flow: Main Acquisition Loop": [
        "The main loop starts when the browser requests aircraft data from the Flask server. The server asks the data-source module for the latest aircraft information.",
        "If the system is in simulation mode, a batch of recorded messages is processed. If it is in live mode, the current receiver snapshot is used.",
        "The raw rows are passed through preprocessing and then stored. The latest aircraft list is returned to the browser as JSON.",
        "The browser updates the aircraft markers, aircraft list, detail panel and graph. This cycle repeats so that the interface remains live.",
    ],
    "3.4 Data-Acquisition and Decoding Pipeline": [
        "The acquisition and decoding pipeline is the core of the project. It converts received communication data into aircraft information that a user can understand.",
        "The pipeline must be tolerant of real-world behaviour. Messages may arrive close together, may be incomplete at the end of a network read, or may not contain all expected fields.",
        "To handle this, the decoder keeps state in memory. When a new message updates only one part of the aircraft state, previous valid values are preserved until better values are received.",
        "This stateful approach is important for a smooth dashboard because aircraft messages are not always complete snapshots of every flight parameter.",
    ],
    "3.4.1 The PingStation Hexadecimal Log Stream": [
        "The PingStation log stream is valuable for testing because the same data can be replayed many times. This makes debugging more reliable than depending only on live traffic.",
        "The project stores multiple forms of the received stream. Hexadecimal logs help inspect the actual byte values, ASCII logs help identify readable tokens, and raw logs preserve the original data.",
        "Session-wise storage is useful because each run of the system can be separated. This makes it easier to compare results from different tests or demonstrations.",
    ],
    "3.4.2 Frame Decoding": [
        "Frame decoding begins by identifying complete messages in the input stream. The decoder uses the frame length to decide where one message ends and the next begins.",
        "If a frame is incomplete, it is not decoded immediately. The remaining bytes are kept until the next chunk arrives. This prevents the decoder from treating incomplete data as a full message.",
        "The decoder also skips values that look impossible for a valid frame length. This helps it resynchronise if noise or unrelated bytes appear in the stream.",
        "After a frame is separated, the decoder extracts available fields. These fields may include ICAO address, callsign, altitude, speed, heading, latitude, longitude, vertical speed and squawk.",
    ],
    "3.4.3 Validation and Normalisation (PingstationPreprocessor)": [
        "Validation and normalisation are needed because raw decoded values may not always be in the same unit or format. A consistent format is required before saving or displaying data.",
        "Altitude is converted into a standard internal form. Speed and heading are also converted so that the frontend and database do not need to guess the unit.",
        "Timestamps are normalised to a consistent time format. This is important because trajectory history and prediction both depend on the time order of samples.",
        "The preprocessor also provides safe defaults. If a field is missing, the system can still create a structured row instead of failing completely.",
    ],
    "3.5 Database Design (FlightTrajectoryStore)": [
        "The database design separates current state from history. The current-state table stores one latest row for each aircraft, while the trajectory table stores multiple time-ordered samples.",
        "This design is efficient because the dashboard often needs the latest state quickly, while prediction and analysis need historical samples.",
        "The trajectory table does not need to store every repeated update. The system checks time, distance, altitude change, heading change and speed change before saving a new sample.",
        "This reduces unnecessary storage while still preserving meaningful movement. It also keeps Excel and CSV outputs easier to inspect.",
        "ICAO-wise Excel files are useful for presentation and analysis because all data for one aircraft can be opened separately. The master workbook provides a combined index.",
    ],
    "3.6 Trajectory-Prediction Algorithm": [
        "Trajectory prediction in this project is meant to improve display continuity. It is not designed for certified navigation or air traffic control decisions.",
        "The prediction module uses recent real observations from the trajectory store. It estimates where the aircraft may be after a short period if fresh data is temporarily unavailable.",
        "Predicted rows are marked differently from real received rows. This is important because the user should always know whether a point comes from the receiver or from the prediction model.",
    ],
    "3.6.1 Prediction Objective and Activation": [
        "The prediction activates only when an aircraft was recently seen but is no longer present in the active live set. This prevents old aircraft from being predicted indefinitely.",
        "A maximum prediction age is used. If the data gap becomes too long, the aircraft should no longer be shown as a reliable prediction.",
    ],
    "3.6.2 Sliding-Window Ordinary Least-Squares Model": [
        "The sliding window contains recent samples for the same aircraft. The model fits separate straight-line trends for latitude, longitude and altitude over time.",
        "The method is called ordinary least squares because it chooses the line that minimises the total squared error between the measured samples and the fitted line.",
        "Using separate trends keeps the implementation simple. Latitude, longitude and altitude are predicted independently, and then the result is combined into a predicted aircraft state.",
        "The method is computationally light, so it can run during each dashboard refresh without slowing the system significantly.",
    ],
    "3.6.3 Why the Regression Is More Stable Than Two-Point Extrapolation": [
        "Two-point extrapolation uses only the last two samples. If either of those samples is noisy, the predicted path can jump suddenly.",
        "Regression uses several recent samples, so one small error has less effect on the final trend. This makes the displayed prediction smoother.",
        "The approach is especially suitable for short periods where the aircraft is moving almost steadily. It is less suitable for long-term forecasting or sudden manoeuvres.",
    ],
    "3.6.4 Kinematic Fallback, Output and Recovery": [
        "The kinematic fallback uses the latest position, ground speed and heading. It is useful when there are not enough samples to fit the regression model.",
        "The fallback is simple but still better than removing the aircraft immediately during a short feed gap. It gives the user a reasonable indication of direction.",
        "When fresh received data returns, the system stops using the predicted row and updates the aircraft using the real received state again.",
    ],
    "3.6.5 Comparison of Suitable Prediction Algorithms": [
        "The algorithm comparison shows that each method has a trade-off. Simple methods are fast and easy to understand, while advanced methods may be more accurate but need more tuning.",
        "Kalman filtering is a strong future candidate because it is widely used in tracking problems. However, it needs noise assumptions and careful state design.",
        "Machine-learning models may become useful after more trajectory data is collected. At the current stage, the project benefits more from a transparent baseline than from a complex model.",
    ],
    "3.6.6 Justification for Selecting Sliding-Window Linear Regression": [
        "Sliding-window linear regression was selected because it balances simplicity, speed and usefulness. It can run locally without training a separate model.",
        "The model is easy to explain in a report and presentation. This is important because the project is educational and should be understandable to examiners and classmates.",
        "The method also keeps measured and predicted data separate. This protects the stored training data from being mixed with estimated values.",
    ],
    "3.7 Aircraft Classification Logic": [
        "Aircraft classification improves the user experience by adding context to each aircraft. Instead of showing only a hex address, the system can show type, registration, category or photo when data is available.",
        "The lookup logic uses the ICAO address as the key. If external data is found, the dashboard can display a more informative aircraft card.",
        "Because external services may fail or be unavailable, the lookup is treated as optional. The tracker still works even if photo or registration data cannot be loaded.",
        "The classification also helps choose suitable icons or labels, making the map easier to understand at a glance.",
    ],
    "3.8 User-Interface Design": [
        "The interface was designed to show the most important information without requiring the user to inspect raw logs. The map occupies the main area because position is the most important aircraft-tracking output.",
        "The live traffic list helps the user select an aircraft quickly. The selected-aircraft panel then shows detailed values such as altitude, speed, vertical speed, heading, position, route and distance.",
        "Status chips at the top show whether the feed is live or simulated. This is important during demonstrations because the viewer can immediately understand the current operating mode.",
        "Predicted aircraft are marked clearly with a PRED label. Dotted path lines help show the estimated movement, while normal markers represent currently received aircraft.",
    ],
    "3.8.1 User Interface Screenshots": [
        "The annotated screenshots show the dashboard in different operating conditions. They are included so that a reader can connect the written explanation with the actual implemented interface.",
        "The screenshots also help during presentation because they point directly to important features such as simulation status, live status, predicted path, route display, selected aircraft data and API activity.",
    ],
    "3.9 Development Tools, Technologies and Environment": [
        "Python was used because it provides simple libraries for networking, data processing, databases and web applications. Flask was selected because it is lightweight and suitable for a local API server.",
        "SQLite was used because it does not require a separate database server. This makes the application easier to run on a normal Windows computer.",
        "Leaflet was used for the interactive map, while Chart.js was used for plotting altitude trends. These tools make the browser dashboard responsive and easy to understand.",
        "PyInstaller was used to create a Windows executable. This helps in demonstration because the application can be run without manually starting every Python file.",
    ],
    "4.1 Test Setup and Data-Collection Methodology": [
        "Testing was performed with both recorded data and live operation. Recorded replay was useful because the same messages could be tested repeatedly after code changes.",
        "Live testing was useful because it showed how the system behaves with real receiver updates, changing aircraft count and temporary gaps in the stream.",
        "The backend terminal output was also observed. Successful repeated requests to the API showed that the browser and Flask server were communicating correctly.",
        "During testing, attention was given to whether aircraft appeared on the map, whether the selected panel updated, whether predicted labels appeared correctly and whether data was written to storage files.",
    ],
    "4.2 Reception Performance and Range Analysis": [
        "Reception performance depends strongly on the environment. A receiver placed near obstructions may receive fewer aircraft than one placed higher with a clearer view of the sky.",
        "High-altitude aircraft generally remain visible over a larger range because the line-of-sight path is longer. Low-altitude aircraft may disappear sooner because of terrain, buildings or the curvature of the Earth.",
        "The observed range values should therefore be understood as results for the current setup. They are not fixed values for all PingStation receivers or all antennas.",
        "The range analysis is still useful because it shows the practical relationship between communication theory and real receiver behaviour.",
    ],
    "4.3 Message-Decoding Accuracy and Signal-Quality Analysis": [
        "The decoding process was checked by observing whether expected aircraft identifiers and motion fields appeared in the dashboard and stored records.",
        "Signal quality affects decoding indirectly. When the receiver does not get a clean message, the software may not receive a valid frame or may receive only partial updates.",
        "The state-merging logic helps reduce visible gaps by preserving previous valid values until new valid values arrive. This makes the display smoother while still using received data.",
        "However, the system should not assume that every displayed value is newly received at every refresh. Some fields may be the most recent known values from earlier messages.",
    ],
    "4.4 Trajectory-Prediction Accuracy": [
        "Prediction accuracy depends on how smoothly the aircraft is moving and how long the feed gap lasts. Short gaps during steady flight are easier to predict than long gaps during turns or climbs.",
        "The present prediction feature is mainly a continuity aid for the dashboard. It keeps the aircraft visible for a short time and helps the user understand the likely direction of travel.",
        "A formal accuracy study would require comparing predicted positions with later received positions or an external trusted reference. This can be added in future work.",
        "The current implementation clearly marks predictions so that estimated positions are not confused with live receiver measurements.",
    ],
    "4.5 End-to-End Latency and System Throughput": [
        "End-to-end latency is the time taken for received aircraft data to appear on the dashboard. It includes input handling, decoding, preprocessing, storage, API response and browser rendering.",
        "The current browser polling approach is simple and reliable for a local demonstration. The browser requests updated data at regular intervals and redraws the relevant interface elements.",
        "For a small number of aircraft, this approach is sufficient. If the system is extended to many users or many receivers, a push-based method such as WebSockets may be better.",
        "Throughput was acceptable for the project scale because the backend could process the received messages and return aircraft updates without visible backlog during testing.",
    ],
    "4.6 Comparison with Existing Systems": [
        "Compared with a basic ADS-B decoder, this project provides a more complete local application. It includes database storage, Excel logging, aircraft lookup, prediction and a user-friendly dashboard.",
        "Compared with global commercial platforms, this project has limited coverage because it uses a local receiver. However, it provides better visibility into the internal working of the system.",
        "The project is therefore best understood as an educational and experimental tracker. Its value is not global coverage, but the ability to show how each processing stage works.",
        "This makes it suitable for academic assessment, laboratory demonstration and future extension by students.",
    ],
    "4.7 Discussion of Results": [
        "The results show that the system meets the main project goals. It can receive or replay aircraft data, decode it, store it and display it on a map.",
        "The interface makes the results understandable because live aircraft, predicted aircraft, routes, speed, altitude and graphs are presented visually.",
        "The storage outputs are also important results. They prove that the project is not only a display tool but also a data collection and analysis platform.",
        "The main technical challenges were handling incomplete data, keeping units consistent, preserving useful historical samples and making prediction understandable.",
        "Overall, the results support the conclusion that the implemented system is a strong prototype for ADS-B aircraft tracking and trajectory analysis.",
    ],
    "5.1 Summary of Work Carried Out": [
        "A complete ADS-B aircraft tracking system was developed using PingStation data. The system receives aircraft data either from a live PingStation stream or from previously recorded data in simulation mode.",
        "The received data is processed through different stages such as data acquisition, frame decoding, preprocessing, validation, storage, prediction and visualization. Important aircraft parameters such as ICAO address, callsign, latitude, longitude, altitude, speed, heading, vertical speed and timestamp are extracted and converted into a useful format.",
        "A Flask-based backend was developed to manage the data processing and provide API endpoints for the frontend dashboard. The frontend displays aircraft on an interactive map and provides useful features such as search, aircraft selection, route information, altitude graph, live/simulation status and predicted aircraft labels.",
        "The system also stores aircraft data for future use. SQLite database, CSV files and ICAO-wise Excel logs are used to save latest aircraft records and historical trajectory samples. This makes the collected data useful for later analysis, comparison and prediction-related work.",
        "A short-gap trajectory prediction feature was also added. When an aircraft temporarily disappears from the received feed, the system estimates its position for a short time using recent movement data and marks it clearly as predicted.",
        "The project also includes practical demonstration features such as annotated dashboard screenshots, route display, aircraft photo/type lookup and a packaged Windows executable. These features make the system easier to present and easier to understand.",
    ],
    "5.2 Conclusions": [
        "The project successfully shows that PingStation ADS-B data can be converted into a useful live aircraft tracking system. Raw received data can be decoded, cleaned, stored and displayed in a form that is easy to understand.",
        "The developed system connects Electronics and Communication concepts with practical software implementation. It includes radio-based aircraft surveillance, digital message decoding, telemetry processing, networking, database management and web-based visualization.",
        "The live map interface makes the system useful for demonstration and learning. It allows users to observe aircraft movement, view important flight details and understand how ADS-B data changes over time.",
        "The storage and export features make the project more useful because the data is not limited to live display only. Stored trajectory samples can be used later for analysis, report generation and future prediction model development.",
        "Overall, the project achieved its main objectives of ADS-B data processing, live aircraft visualization, trajectory storage and basic prediction. It provides a strong foundation for further improvements in aircraft tracking and flight data analysis.",
    ],
    "5.3 Limitations of the Present Implementation": [
        "The performance of the system depends on the coverage of the PingStation receiver and antenna placement. Aircraft that are too far away, blocked by buildings or terrain, or flying at low altitude may not be received properly.",
        "The system also depends on the quality of received ADS-B messages. Some messages may be incomplete, delayed or missing certain fields, which can affect the displayed aircraft information.",
        "Map display and aircraft information lookup may require internet access. If network connectivity is unavailable, external map tiles, aircraft photos or registration details may not load properly.",
        "The prediction feature is designed only for short gaps in the received data. It should not be treated as a certified navigation or air traffic control tool. The prediction may become inaccurate during sharp turns, sudden climbs, descents or long data gaps.",
        "The current system is mainly designed for local use and demonstration. It is not yet optimized for large-scale deployment, multiple users, multiple receivers or continuous long-term operation.",
    ],
    "5.4 Scope for Future Work": [
        "Future work can include improving the receiver setup by using better antenna placement, higher mounting height and more controlled live testing. This can help improve reception range and data quality.",
        "The system can be extended to support offline maps so that the dashboard can work even when internet access is not available. This would make the project more reliable for demonstrations and field use.",
        "The current polling-based update method can be improved using WebSockets or Server-Sent Events. This would allow the backend to push aircraft updates to the browser in real time and reduce display delay.",
        "The prediction module can be improved by comparing different algorithms such as Kalman filter, Extended Kalman filter, route-based models and machine-learning methods. These methods can be tested after collecting a larger and cleaner trajectory dataset.",
        "More validation can be added by comparing decoded aircraft positions with trusted reference data. This would help measure decoding accuracy, prediction error and system reliability.",
        "The project can also be expanded to support multiple receivers. Multi-receiver support can improve coverage and may allow future work on multilateration or position verification.",
        "Additional dashboard features can be added, such as alert generation, aircraft filtering by altitude or callsign, route history replay, export controls and better analytics for stored flight data.",
    ],
}


def set_paragraph_text(paragraph, text):
    for run in paragraph.runs:
        run.text = ""
    if paragraph.runs:
        paragraph.runs[0].text = text
    else:
        paragraph.add_run(text)


def insert_after(paragraph, text, template):
    new_xml = deepcopy(template._p)
    paragraph._p.addnext(new_xml)
    new_para = Paragraph(new_xml, paragraph._parent)
    set_paragraph_text(new_para, text)
    return new_para


def main():
    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    doc = Document(SOURCE)

    # Insert additional detail after each target heading. Reuse the first
    # following body paragraph as a style template to preserve the report format.
    for heading, paragraphs in EXPANSIONS.items():
        all_paragraphs = doc.paragraphs
        idx = next(i for i, p in enumerate(all_paragraphs) if p.text.strip() == heading)
        template = next((p for p in all_paragraphs[idx + 1 :] if p.text.strip()), all_paragraphs[idx])
        cursor = all_paragraphs[idx]
        for text in paragraphs:
            cursor = insert_after(cursor, text, template)

    doc.save(OUTPUT)
    print(OUTPUT)


if __name__ == "__main__":
    main()
