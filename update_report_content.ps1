$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem

$sourceDir = Join-Path $PSScriptRoot 'report_original_unpacked'
$workDir = Join-Path $PSScriptRoot 'report_updated_unpacked'
$outputDocx = Join-Path $PSScriptRoot 'Aircraft_Tracker_Project_Report_Updated.docx'

if (Test-Path -LiteralPath $workDir) {
    Remove-Item -LiteralPath $workDir -Recurse -Force
}
Copy-Item -LiteralPath $sourceDir -Destination $workDir -Recurse

$documentPath = Join-Path $workDir 'word\document.xml'
$xml = [xml](Get-Content -LiteralPath $documentPath -Raw)
$ns = [System.Xml.XmlNamespaceManager]::new($xml.NameTable)
$ns.AddNamespace('w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main')
$paragraphs = $xml.SelectNodes('//w:body/w:p | //w:body/w:tbl//w:p', $ns)

$updates = @{
    32 = @'
Automatic Dependent Surveillance-Broadcast (ADS-B) is a cooperative surveillance technology that has become a key component of modern air-traffic management. Suitably equipped aircraft periodically broadcast their International Civil Aviation Organisation (ICAO) 24-bit address, callsign, position, altitude, ground speed, heading and vertical rate. This project, "Design and Implementation of a Real-Time Aircraft Tracking and Trajectory Prediction System", presents an end-to-end platform that receives the structured binary stream supplied by a uAvionix PingStation, separates it into complete frames, decodes and normalises aircraft records, stores continuous per-aircraft trajectory histories, predicts short communication gaps using sliding-window linear regression with a deterministic kinematic fallback, and presents the live air picture through a Flask and Leaflet.js browser interface.
'@
    33 = @'
The work spans both electronics/communication and computer-science aspects. The report examines the 1090 MHz Mode-S Extended-Squitter physical layer, including Pulse-Position Modulation, the preamble, frame structure and radio-propagation considerations, while recognising that the PingStation performs RF demodulation and exposes a higher-level length-prefixed stream to the software. The implemented software is divided into acquisition, decoding, preprocessing, storage/prediction and presentation layers. Live hardware and recorded simulation data pass through the same decoder, while session-specific HEX, ASCII, exact-stream and intermediate decoded logs support repeatable testing and diagnosis.
'@
    34 = @'
The system has been exercised with both live PingStation reception and replay of recorded PingStation data. Functional testing verified dynamic ICAO detection, continuous SQLite and per-aircraft Excel logging, automatic creation of a master workbook, one-second dashboard updates, live/simulation switching, short-outage trajectory prediction and resumption of actual tracking when reception returns. Prediction is deliberately limited to a configurable short horizon (180 seconds by default), because extrapolation uncertainty grows rapidly during turns and speed changes. The current predictor is an interpretable machine-learning baseline rather than a trained deep-neural model; Kalman filtering and recurrent sequence models remain appropriate subjects for future comparative research.
'@
    36 = 'Keywords: ADS-B, Mode-S Extended Squitter, PingStation, TCP Frame Decoding, SQLite, Linear Regression, Trajectory Prediction, Flask, Leaflet.js, Real-Time Tracking System.'
    135 = 'Logical data model of the FlightTrajectoryStore: the flight_latest table, timestamped trajectory_samples table, per-ICAO Excel files, a session master workbook, and non-persistent predicted output records.'
    137 = 'Sliding-window linear-regression trajectory prediction: separate least-squares trends are fitted to recent latitude, longitude and altitude observations and extrapolated only during a short reception outage; constant-velocity kinematics provide a fallback when insufficient samples are available.'
    139 = 'Flowchart of the aircraft-icon classification logic: external ICAO type information is matched against curated helicopter and propeller type-designator sets, with the standard fixed-wing icon used as the default.'
    147 = 'Short-outage trajectory-prediction validation, comparing predicted positions against subsequently received positions and distinguishing the regression estimate from the constant-velocity fallback.'
    272 = ([char]0x2022) + '  To design a compact FlightTrajectoryStore that maintains the latest state and timestamped trajectory samples in SQLite, while automatically creating session-based per-ICAO Excel files, a master workbook, and diagnostic HEX, ASCII, raw and exact-stream logs.'
    273 = ([char]0x2022) + '  To develop an interpretable short-outage trajectory predictor based on sliding-window linear regression over recent real observations, with a constant-velocity kinematic fallback and a configurable maximum prediction age.'
    274 = ([char]0x2022) + '  To build an interactive browser dashboard using Flask, Leaflet.js and Chart.js, including smooth aircraft movement, actual and predicted track overlays, aircraft/route information, a resizable detail panel, live status feedback, and a direct simulation-mode switch.'
    277 = @'
The scope of the present work is deliberately bounded to a single PingStation receiver and the structured stream emitted by that receiver. The software does not demodulate raw I/Q samples; instead, it receives length-prefixed PingStation frames over TCP or replays captured frames through the same decoder. Multilateration and independent verification of an aircraft's self-reported position are outside the current implementation. The trajectory predictor uses recent observations from the current session and is intended to bridge short reception gaps, not to reconstruct complete flight plans or make long-horizon strategic forecasts. Advanced filters and trained sequence models are therefore treated as future comparative extensions.
'@
    296 = @'
Given these considerations, the present project adopts sliding-window ordinary least-squares linear regression as its primary short-outage predictor. Separate functions of time are fitted to the most recent latitude, longitude and altitude samples for each aircraft, reducing the influence of any single noisy observation compared with two-point extrapolation. When too few valid samples are available, the system falls back to constant-velocity kinematics based on the most recent speed, heading and vertical rate. This combination is computationally light, explainable during evaluation, independent of a large pre-trained dataset, and suitable for the project's real-time single-host environment.
'@
    324 = @'
Unlike many communication protocols, Mode-S uses its 24-bit Parity/Interrogator-Identity field for integrity checking rather than appending a separate CRC trailer. Recomputing the polynomial remainder is a standard way to validate a raw Mode-S frame and remains important background theory for understanding the receiver. In the implemented architecture, however, RF demodulation, Mode-S framing and the associated low-level integrity handling occur inside the PingStation firmware. The Python application receives the PingStation's higher-level structured stream and therefore concentrates on TCP frame completeness, supported message layouts, field bounds, state continuity and timestamp validity.
'@
    338 = @'
Translating the conceptual block diagram into software, the system is structured as a layered pipeline. The Hardware/Acquisition layer contains the antenna, PingStation and TCP receiver; the Data layer selects live or replay input, captures session logs, separates complete frames and decodes aircraft fields; the Processing layer normalises records, maintains aircraft state, stores trajectories and builds short-outage predictions; and the Presentation layer exposes Flask REST endpoints to the Leaflet.js client. Both live and simulation modes deliberately share the same decoder so that recorded data tests the exact logic used with hardware.
'@
    350 = 'Selects live PingStation TCP reception or recorded replay data, manages the runtime mode switch and configuration file, creates per-session output directories, and passes both sources through the same decoder.'
    352 = 'Performs cached enrichment look-ups for aircraft type/registration and photographs, and obtains scheduled origin/destination information from a route service. If a route cannot be verified, the interface reports it as unavailable rather than inventing an estimated destination.'
    354 = 'Normalises decoded fields and units, standardises timestamps, supplies safe defaults, and produces a stable schema containing ICAO, callsign, position, altitude, heading and velocity values.'
    356 = 'Owns the SQLite flight_latest and trajectory_samples tables; continuously appends real trajectory observations; creates one Excel file per ICAO and a session master workbook; and produces short-outage prediction records without storing them as real training observations.'
    358 = 'Implements REST endpoints for aircraft data, status and live/simulation mode control, aircraft enrichment, verified route lookup, logging status and trajectory-history export, and serves the browser application.'
    360 = 'Implements the Leaflet.js map, locally bundled frontend libraries and aircraft icons, searchable live-traffic list, smooth marker interpolation, red dotted prediction paths, route and aircraft panels, altitude chart, resizable lower panel, offline-map status and the simulation-mode button.'
    362 = @'
This modular decomposition has a direct pedagogical benefit: each module corresponds to a coherent topic - source abstraction and TCP acquisition, binary frame processing, data cleaning, relational time-series storage, interpretable regression-based prediction, and interactive visualisation. Separating these responsibilities also made it possible to diagnose decoding, logging and prediction errors without rewriting unrelated portions of the project.
'@
    364 = @'
Figure 3.4 summarises the operational control flow. In live mode, a background TCP receiver accepts the PingStation connection, preserves incoming bytes, separates complete frames and updates decoder state. In simulation mode, each browser data request advances a controlled batch of recorded frames through the same decoder. The /api/aircraft request then preprocesses the current snapshot, appends genuine observations to the FlightTrajectoryStore, creates predictions only for aircraft that have temporarily disappeared, and returns the combined display records to the browser.
'@
    369 = @'
The uAvionix PingStation performs RF reception and low-level demodulation in its firmware and connects to the application over TCP. The application listens on a configurable local port (30003 by default) and receives a continuous binary byte stream. TCP does not preserve application-message boundaries, so a single network read may contain part of one frame, several complete frames, or a mixture of both. The software therefore retains incomplete bytes and uses the PingStation frame-length field to reconstruct complete messages before decoding. Simulation mode replays captured hexadecimal payload lines through this identical frame-splitting and decoding path.
'@
    370 = @'
For every application session, the capture logger creates separate HEX and ASCII text files, an exact unfiltered STREAM.bin file, and a RAW.txt JSON-lines file containing decoded records before preprocessing. The session also contains an icao_excel directory and a master workbook. This staged logging distinguishes the original receiver bytes, printable callsign evidence, intermediate decoder output and final normalised flight records, making faults traceable without mixing different processing levels.
'@
    372 = @'
The decoding stage first separates the continuous byte stream into complete length-prefixed frames. Obvious noise lengths are skipped and an incomplete trailing frame is retained until more bytes arrive. Each complete frame is then interpreted using the observed PingStation compressed-message layout and legacy fallbacks. The decoder extracts the ICAO address, callsign, latitude, longitude, altitude, ground speed, heading, vertical speed and squawk when those fields are present. Latitude/longitude candidates are checked against valid geographic ranges, and when more than one byte layout appears numerically valid, continuity with the aircraft's previous position is used to select the plausible result.
'@
    373 = @'
Decoding is incremental because individual PingStation frames may contain only a subset of the complete aircraft state. A dictionary keyed primarily by ICAO address preserves the latest known values and merges only fields present in each new frame. Zero-valued updates are prevented from erasing valid speed, heading or vertical-rate information unless they are meaningful. Each record also carries its timestamp, source address, frame length and an internal last-seen time, allowing stale aircraft to be removed from the active snapshot while retaining their stored history.
'@
    375 = @'
The preprocessor converts the decoder's flexible records into a stable application schema. Latitude and longitude are represented as decimal degrees; altitude is normalised to millimetres internally; heading is stored as degrees multiplied by 100; horizontal and vertical velocities are stored in centimetres per second; callsign and ICAO identifiers are upper-cased; and timestamps are normalised to ISO format. Record-level validation is performed through frame completeness, field presence, numeric conversion, geographic bounds and continuity checks. Low-level RF demodulation and Mode-S integrity handling are responsibilities of the PingStation firmware rather than being recomputed by the current Python preprocessor.
'@
    379 = 'Frame completeness and resynchronisation'
    380 = 'Use the leading frame-length byte to extract complete messages; retain incomplete trailing bytes and skip implausible lengths so that later frames can be recovered correctly.'
    382 = 'Normalise altitude to millimetres, heading to degrees multiplied by 100, and horizontal/vertical speed to centimetres per second for consistent internal processing.'
    384 = 'Convert valid timestamps to ISO form and use a current UTC timestamp when a source timestamp is absent or malformed.'
    386 = 'Reject unusable positions, enforce valid latitude/longitude ranges during decoding, and prefer coordinate interpretations that preserve continuity with the previous aircraft position.'
    390 = 'Decode supported PingStation coordinate layouts and retain the most plausible candidate using geographic bounds and previous-position continuity.'
    392 = @'
Only useful, structured aircraft records are passed to the FlightTrajectoryStore. Keeping capture, decode, normalisation and storage as separate stages makes the behaviour auditable: the exact incoming bytes remain available even if a later interpretation must be corrected, while downstream code receives a predictable schema.
'@
    394 = @'
Persistent storage is provided by SQLite together with session-based Excel and diagnostic files. The flight_latest table contains one current row per aircraft for fast retrieval, while trajectory_samples appends timestamped genuine observations for historical analysis and prediction. Latitude and longitude are stored as scaled integers to six decimal places in SQLite, avoiding repeated floating-point comparison issues. For human-readable analysis, every newly observed ICAO automatically receives its own Excel file containing timestamp, ICAO, callsign, latitude, longitude, altitude, heading, speed and vertical rate; the Excel display preserves latitude and longitude to eight decimal places. A master workbook created for the session maintains an index and individual ICAO sheets.
'@
    396 = 'Fig. 3.5: Current FlightTrajectoryStore model: flight_latest, trajectory_samples, session-based per-ICAO Excel files, a master workbook, and non-persistent prediction output.'
    397 = @'
An index on (flight_key, timestamp) supports efficient retrieval of an aircraft's recent history. To control storage growth, a new trajectory sample is appended when sufficient time, distance, altitude, heading or speed change has occurred, rather than duplicating identical values on every browser poll. Historical retention is configurable, while the latest-state row is updated whenever data is received. Predictions are intentionally excluded from the real-observation tables so that later model training is not contaminated by the system's own estimates.
'@
    399 = @'
The primary trajectory predictor is a sliding-window ordinary least-squares regression model. For an aircraft that has stopped appearing in the current receiver snapshot, the store retrieves up to twenty of its most recent genuine trajectory samples. Provided at least four valid timestamped samples are available, independent straight-line models are fitted for latitude, longitude and altitude as functions of elapsed time.
'@
    400 = 'b = SUM[(t_i - mean(t))(y_i - mean(y))] / SUM[(t_i - mean(t))^2]'
    401 = 'a = mean(y) - b mean(t),    predicted y(t_future) = a + b t_future'
    402 = @'
The calculation is applied separately to latitude, longitude and altitude. The fitted latitude and longitude slopes are also converted into an estimated speed and bearing, while the altitude slope provides vertical speed. Predictions are generated only after real data for that ICAO disappears and only until the configurable maximum age is reached (180 seconds by default). When the aircraft returns, its actual observation immediately replaces prediction output. Because the regression uses several recent samples, a single noisy coordinate has less influence than it would in a two-point or last-state extrapolation.
'@
    404 = 'Fig. 3.6: Regression-based short-outage prediction. Recent latitude, longitude and altitude samples define least-squares trends; a constant-velocity projection is retained as the fallback when regression cannot be fitted.'
    405 = @'
If fewer than four suitable samples exist or a fitted result is not finite, the predictor falls back to deterministic kinematics. It estimates missing motion parameters from the latest two stored samples where necessary, computes distance as speed multiplied by outage duration, projects latitude/longitude along the latest heading, and applies vertical speed to altitude. Output records are marked isPredicted and include predictionAgeSec. They are displayed with a red dotted path but are not written back into training data. This hierarchy provides a transparent ML baseline while retaining predictable behaviour for sparse tracks.
'@
    407 = @'
The interface selects among helicopter, propeller-aircraft and standard fixed-wing marker icons. Aircraft metadata is requested from hexdb.io using the ICAO address. The returned ICAO type designator and descriptive type text are compared with curated helicopter and propeller designator sets maintained in aircraft_lookup.py. A matching rotorcraft type selects the helicopter icon, a matching piston or turboprop type selects the propeller icon, and all remaining or unavailable types use the standard fixed-wing/airliner silhouette. Results are cached so that one-second interface polling does not repeatedly call the external service.
'@
    409 = 'Fig. 3.7: Aircraft-icon classification logic based on cached external ICAO type information and curated helicopter/propeller designator sets, with a standard fixed-wing default.'
    410 = @'
This classification is intentionally coarse and affects presentation rather than surveillance processing. The result depends on the completeness and correctness of the external aircraft database and on the coverage of the curated designator sets. A type not present in those sets may therefore receive the standard fixed-wing icon even when a more specialised classification would be possible. The detailed manufacturer, type and registration returned by the lookup service are shown separately when available, and network failure falls back safely to the default icon without affecting reception or tracking.
'@
    412 = @'
The browser front end was designed for immediacy, drill-down and resilience. The Leaflet map displays every tracked aircraft with a rotated category-specific marker and shadow; the left panel provides callsign/ICAO search and live status; and the resizable bottom panel provides flight, aircraft and route tabs, current telemetry, altitude history and optional photographs. A toolbar button changes simulation mode without editing source code. Scheduled route names are shown only when a route service returns a verified origin and destination; otherwise the interface explicitly reports that the route is unavailable, avoiding misleading heading-based guesses.
'@
    415 = @'
The client requests /api/aircraft approximately once per second and reconciles the response with existing marker layers. New ICAOs create new markers, current ICAOs are smoothly interpolated and rotated over the refresh interval, and departed ICAOs are removed after their real or predicted lifetime ends. Actual paths use a solid light-blue line, while outage predictions use a red dotted line and a PRED label. The status endpoint is polled separately to show simulation, live-wait, live-connected or socket-error states. OpenStreetMap tiles and online enrichment may become unavailable without internet, but the locally bundled interface, PingStation reception, decoding, logging, prediction and aircraft marker updates continue to operate.
'@
    437 = 'hexdb.io / planespotters.net / adsbdb.com APIs'
    438 = 'Optional cached enrichment services used for aircraft type and registration, representative photographs, and verified scheduled origin/destination route information.'
    449 = @'
The system was evaluated with two complementary data sources: live TCP sessions from the PingStation and deterministic replay of the project's captured PingStation hexadecimal stream. Replay is not a separate synthetic aircraft generator; it passes logged payloads through the same frame splitter, decoder, state merger, preprocessor, storage and prediction logic used in live mode. Testing focused on repeatability and functional correctness: decoding known captured aircraft, automatic session-folder creation, per-ICAO Excel generation, continuous trajectory appending, live/simulation switching, disappearance-triggered prediction, actual-data recovery, offline core operation and one-second browser updates.
'@
    473 = @'
Decoder reliability was assessed by replaying captured PingStation payloads and comparing the resulting ICAO, callsign, position and motion fields with known examples observed during live operation. Particular attention was paid to TCP frame boundaries, short control packets, ambiguous coordinate byte layouts and partial messages. Length-based frame separation prevented incomplete or concatenated network reads from shifting field offsets, while state merging preserved fields that were not repeated in every message. Exact stream and intermediate RAW logs made decoding decisions reproducible during correction of field-layout errors.
'@
    476 = @'
Functional replay tests confirmed that the current decoder consistently recovered known aircraft records from the supplied capture, including ICAO-based identity, callsign, Indian-region coordinates, altitude and motion values. The report does not claim a statistically measured software CRC pass rate because the current Python pipeline receives the PingStation's structured output rather than raw RF samples; formal bit-error-rate and CRC benchmarking would require lower-level receiver data and remains outside the present evaluation.
'@
    478 = @'
The current trajectory feature was evaluated as a short communication-gap bridge rather than as a thirty-minute flight-plan forecast. During replay, an aircraft was allowed to accumulate genuine trajectory samples and was then made inactive by advancing the recorded stream to other aircraft. The system correctly generated isPredicted records from the recent trajectory, increased predictionAgeSec while data remained absent, displayed a red dotted continuation, stopped predictions after the configured maximum age, and returned to actual tracking when the ICAO reappeared.
'@
    480 = 'Fig. 4.3: Short-outage prediction validation procedure, comparing regression and kinematic-fallback estimates with positions received after the aircraft returns.'
    481 = @'
The regression method produced smoother short-term continuation than estimates derived solely from the last two positions because it fitted the overall trend of several recent observations. Its intended horizon is deliberately limited to 180 seconds by default. Error is expected to increase during turns, climbs, descents and speed changes, and no unsupported long-horizon accuracy figure is claimed. A formal research comparison should calculate Haversine position error, altitude mean absolute error, heading error and runtime for constant velocity, linear regression, Kalman filtering and a trained sequence model on held-out trajectories.
'@
    503 = 'Sliding-window linear regression with constant-velocity fallback (built-in)'
    507 = 'SQLite + session HEX/ASCII/RAW logs + per-ICAO Excel/master workbook'
    521 = @'
Taken together, the tests show that the current application performs the complete operational pipeline required for the project: it accepts live or replayed PingStation data, reconstructs complete frames, decodes and normalises dynamic aircraft records, logs each session at multiple processing stages, maintains compact historical trajectories, bridges short reception gaps with clearly labelled predictions, and updates an interactive map at approximately one-second intervals. The results also identify clear research limits: the receiver is single-station, online route/type enrichment is best-effort, polling contributes visible latency, and the baseline predictor cannot anticipate future turns or flight-plan constraints.
'@
    526 = @'
This project designed and implemented an end-to-end ADS-B aircraft-tracking platform spanning 1090 MHz reception theory, PingStation TCP acquisition, length-prefixed frame separation, proprietary-message decoding, incremental state merging, unit normalisation, SQLite time-series storage, session-based HEX/ASCII/RAW capture, automatic per-ICAO Excel and master-workbook generation, short-outage trajectory prediction, optional aircraft/route enrichment, and a Leaflet.js/Chart.js/Flask dashboard. The final interface includes smooth marker movement, actual and red dotted predicted paths, route and aircraft tabs, status feedback, a direct simulation switch, a resizable information panel and graceful continuation of the core tracking functions when internet services are unavailable.
'@
    529 = @'
Several conclusions can be drawn. First, a PingStation and a single-host Python application can provide a practical local aircraft-tracking and logging platform without dependence on a commercial aggregation network. Second, using the same decoder for live and replay data makes captured sessions valuable for repeatable debugging and regression testing. Third, a sliding-window linear-regression model provides an interpretable, computationally inexpensive baseline for bridging brief reception gaps and is more resistant to isolated noisy samples than two-point extrapolation, although it cannot anticipate manoeuvres. Finally, the separation of acquisition, decoding, preprocessing, storage, prediction and presentation made substantial iterative corrections possible without destabilising the rest of the system.
'@
    531 = @'
The implementation has important limitations. It relies on one receiver and cannot independently validate an aircraft position through multilateration. The decoder is based on observed PingStation message layouts and should be validated against additional firmware versions and captures. SQLite stores coordinates to six decimal places, while eight-decimal formatting in Excel does not create accuracy beyond that present in the received data. The regression predictor assumes that the recent trend continues and is therefore intended only for short outages; it has no knowledge of airways, weather, air-traffic-control instructions or filed flight plans. Scheduled routes, aircraft metadata, photographs and base-map tiles depend on external services, although their failure does not stop local reception, decoding, logging or marker updates. The client still uses periodic polling and the application is designed for a single-host laboratory deployment.
'@
    536 = ([char]0x2022) + '  Advanced trajectory-model comparison: evaluate the current linear-regression baseline against constant velocity, Kalman and Extended Kalman filters, gradient-boosted models, and LSTM/GRU sequence networks using held-out trajectories and clearly defined position, altitude and runtime metrics.'
    538 = ([char]0x2022) + '  Data-quality and decoder validation: collect labelled captures from additional PingStation sessions and firmware versions, quantify field-decoding accuracy, and introduce automated replay tests for known ICAO, callsign, position and motion values.'
    576 = @'
The following excerpt is taken from the bundled PingStation replay file used by the current data_source.py module. Each line is one captured payload represented as hexadecimal text. Unlike a conventional 28-character raw DF17 example, these lines represent the structured, variable-length stream emitted by the PingStation and are passed through the same length-based frame splitter and decoder as live TCP bytes.
'@
    577 = '0e86db07800dbb48080000c0fe00'
    578 = '0d600b07801748020800765c00'
    579 = '2602df07407a3eff09084157433636364c200088b80c029105338ad74170489c420000020401'
    580 = '26f3ef07407a3eff09084157433636364c200088b80b029205338ad74170489c420000020401'
    581 = '1311dc078017484e0800765cfd016004400200'
    582 = '2669e107407a3eff09084157433636364c200088b80b0292050877d741894c9c420000020401'
    584 = @'
Decoding the first complete aircraft-bearing frame beginning 2602df... yields ICAO address 407A3E, callsign AWC666L, latitude approximately 26.942480 degrees, longitude approximately 78.141479 degrees, altitude 10,668 m, heading 142.5 degrees, horizontal speed approximately 269.57 m/s and squawk 1026. Later frames update only the fields that changed, while the decoder's ICAO-keyed state preserves the remaining values.
'@
    587 = 'The following simplified excerpt illustrates the ordinary least-squares calculation used by the current trajectory-prediction engine:'
    588 = 'def fit_linear_regression_1d(xs, ys):'
    590 = '    n = len(xs)'
    591 = '    if n < 2:'
    592 = '        return None'
    593 = '    x_mean = sum(xs) / n'
    594 = '    y_mean = sum(ys) / n'
    595 = '    sxx = sum((x - x_mean) ** 2 for x in xs)'
    596 = '    if sxx <= 0:'
    597 = '        return None'
    599 = '    sxy = sum((x - x_mean) * (y - y_mean)'
    600 = '              for x, y in zip(xs, ys))'
    601 = '    slope = sxy / sxx'
    602 = '    intercept = y_mean - slope * x_mean'
    603 = '    return intercept, slope'
    604 = 'predicted_y = intercept + slope * future_time'
    606 = @'
The production implementation applies this fit independently to recent latitude, longitude and altitude samples. It requires at least four valid trajectory points, derives speed, heading and vertical speed from the fitted slopes, checks that outputs are finite, and falls back to deterministic kinematics when a reliable fit cannot be produced. Predicted records are marked separately and never inserted into the real trajectory-training table.
'@
    613 = 'Linear regression:A statistical method that fits a straight-line relationship between a dependent value and time by minimising the sum of squared residuals; in this project it is used as an interpretable short-outage trajectory baseline.'
}

foreach ($index in ($updates.Keys | Sort-Object)) {
    if ($index -ge $paragraphs.Count) {
        throw "Paragraph index $index is outside the document."
    }
    $paragraph = $paragraphs[$index]
    $textNodes = @($paragraph.SelectNodes('.//w:t', $ns))
    if ($textNodes.Count -eq 0) {
        throw "Paragraph index $index has no text node."
    }
    $newText = [string]$updates[$index]
    $newText = $newText.Trim("`r", "`n")
    $textNodes[0].InnerText = $newText
    $textNodes[0].SetAttribute('space', 'http://www.w3.org/XML/1998/namespace', 'preserve') | Out-Null
    for ($i = 1; $i -lt $textNodes.Count; $i++) {
        $textNodes[$i].InnerText = ''
    }
}

$settings = [System.Xml.XmlWriterSettings]::new()
$settings.Encoding = [System.Text.UTF8Encoding]::new($false)
$settings.Indent = $false
$writer = [System.Xml.XmlWriter]::Create($documentPath, $settings)
$xml.Save($writer)
$writer.Close()

if (Test-Path -LiteralPath $outputDocx) {
    Remove-Item -LiteralPath $outputDocx -Force
}
[System.IO.Compression.ZipFile]::CreateFromDirectory($workDir, $outputDocx, [System.IO.Compression.CompressionLevel]::Optimal, $false)

Write-Output "Updated paragraphs: $($updates.Count)"
Write-Output "Output: $outputDocx"
