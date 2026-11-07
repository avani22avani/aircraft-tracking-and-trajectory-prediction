import os
import threading
import webbrowser
from datetime import datetime
from pathlib import Path

from flask import Flask, jsonify, render_template, request
from data_source import (
    ICAO_EXCEL_DIR,
    SESSION_DIR,
    SIMULATION_MODE,
    get_aircraft_data,
    get_status,
    set_simulation_mode,
)
from preprocessing.pingstation_preprocess import PingstationPreprocessor
from storage.flight_data_store import FlightTrajectoryStore
from runtime_paths import app_dir, resource_dir
from aircraft_lookup import lookup_aircraft_info, lookup_flight_route

app = Flask(
    __name__,
    template_folder=str(resource_dir() / "templates"),
    static_folder=str(resource_dir() / "static"),
)
preprocessor = PingstationPreprocessor()
flight_store = FlightTrajectoryStore(
    db_path=os.getenv(
        "FLIGHT_DB_PATH",
        str(app_dir() / "logs" / "flight_training.db"),
    ),
    enabled=os.getenv("FLIGHT_DB_ENABLED", "true").lower() == "true",
    min_interval_sec=int(os.getenv("FLIGHT_DB_MIN_INTERVAL_SEC", "5")),
    min_distance_m=float(os.getenv("FLIGHT_DB_MIN_DISTANCE_M", "30")),
    min_altitude_change_m=int(os.getenv("FLIGHT_DB_MIN_ALT_CHANGE_M", "20")),
    min_heading_change_de2=int(os.getenv("FLIGHT_DB_MIN_HEADING_CHANGE_DE2", "500")),
    min_speed_change_cms=int(os.getenv("FLIGHT_DB_MIN_SPEED_CHANGE_CMS", "200")),
    history_retention_days=int(os.getenv("FLIGHT_DB_HISTORY_RETENTION_DAYS", "30")),
    append_csv_enabled=os.getenv("FLIGHT_TRAJECTORY_APPEND_CSV_ENABLED", "true").lower() == "true",
    append_csv_dir=os.getenv(
        "FLIGHT_TRAJECTORY_APPEND_CSV_DIR",
        str(app_dir() / "logs" / "training_csv"),
    ),
    icao_excel_enabled=os.getenv("FLIGHT_ICAO_EXCEL_ENABLED", "true").lower() == "true",
    icao_excel_dir=os.getenv(
        "FLIGHT_ICAO_EXCEL_DIR",
        str(ICAO_EXCEL_DIR),
    ),
    icao_master_enabled=os.getenv("FLIGHT_ICAO_MASTER_ENABLED", "true").lower() == "true",
    icao_master_path=os.getenv(
        "FLIGHT_ICAO_MASTER_PATH",
        str(SESSION_DIR / "MASTER_ICAO.xlsx"),
    ),
    reset_excel_on_start=os.getenv("FLIGHT_ICAO_RESET_ON_START", "false").lower() == "true",
    icao_excel_extension=os.getenv("FLIGHT_ICAO_EXCEL_EXTENSION", "xls"),
)
matrix_csv_enabled = os.getenv("FLIGHT_MATRIX_CSV_ENABLED", "false").lower() == "true"
matrix_csv_path = os.getenv(
    "FLIGHT_MATRIX_CSV_PATH",
    str(app_dir() / "logs" / "flight_matrix_latest.csv"),
)
matrix_csv_limit = int(os.getenv("FLIGHT_MATRIX_CSV_LIMIT", "5000"))
prediction_enabled = os.getenv("FLIGHT_PREDICTION_ENABLED", "true").lower() == "true"
prediction_max_age_sec = int(os.getenv("FLIGHT_PREDICTION_MAX_AGE_SEC", "180"))
prediction_limit = int(os.getenv("FLIGHT_PREDICTION_LIMIT", "1000"))

if SIMULATION_MODE:
    # Simulation timestamps are synthetic and restart from the same base each
    # run; without this, leftover trajectory data from a previous run can mix
    # with the current run and make predicted aircraft "trace back" their route.
    flight_store.reset_for_new_session()

if not matrix_csv_enabled:
    try:
        Path(matrix_csv_path).unlink(missing_ok=True)
    except Exception as exc:
        app.logger.warning("Could not remove matrix CSV file: %s", exc)

if os.getenv("FLIGHT_ICAO_EXCEL_ENABLED", "true").lower() == "true" and not flight_store.icao_excel_enabled:
    msg = "ICAO Excel logging is enabled but unavailable. Install openpyxl in the same Python environment."
    app.logger.warning(msg)
    print(f"WARNING: {msg}")


def _build_storage_rows(raw_rows, normalized_rows):
    """
    Preserve normalized core fields while keeping optional telemetry from raw rows
    (e.g. verVelocityCMS/sourceIp/squawk) for logging.
    """
    merged_rows = []
    total = max(len(raw_rows), len(normalized_rows))
    for idx in range(total):
        raw = raw_rows[idx] if idx < len(raw_rows) and isinstance(raw_rows[idx], dict) else {}
        norm = normalized_rows[idx] if idx < len(normalized_rows) and isinstance(normalized_rows[idx], dict) else {}
        rec = dict(raw)
        rec.update(norm)
        merged_rows.append(rec)
    return merged_rows

@app.route("/")
def home():
    return render_template("index.html")

@app.route("/api/aircraft")
def aircraft():
    # Get data from source
    data = get_aircraft_data()
    
    # If data is a dict with an 'aircraft' key, get that list. 
    # Otherwise, assume data itself is the list.
    raw_list = data.get("aircraft", []) if isinstance(data, dict) else data
    storage_input_list = (
        data.get("aircraft_events", data.get("aircraft_raw", raw_list))
        if isinstance(data, dict)
        else raw_list
    )
    
    # Process the list
    processed_list = [preprocessor.preprocess(ac) for ac in raw_list]

    # Pass through a few extra telemetry fields the UI can use for a richer
    # flight-detail panel (vertical speed, squawk, source feed) without
    # changing the normalized schema used elsewhere.
    for idx, source_row in enumerate(raw_list):
        if idx >= len(processed_list) or not isinstance(source_row, dict):
            continue
        for key in ("verVelocityCMS", "squawk", "sourceIp"):
            value = source_row.get(key)
            if value not in (None, ""):
                processed_list[idx][key] = value

    processed_storage_list = _build_storage_rows(
        storage_input_list,
        [preprocessor.preprocess(ac) for ac in storage_input_list],
    )
    flight_store.ingest_many(processed_storage_list)

    if prediction_enabled:
        active_icao = {
            str(row.get("icaoAddress", "")).strip().upper()
            for row in raw_list
            if isinstance(row, dict) and str(row.get("icaoAddress", "")).strip()
        }
        predicted_rows = flight_store.build_predicted_rows(
            active_icao_set=active_icao,
            max_prediction_age_sec=prediction_max_age_sec,
            limit=prediction_limit,
        )
        if predicted_rows:
            processed_list.extend(predicted_rows)

    if matrix_csv_enabled:
        try:
            flight_store.write_latest_matrix_csv(matrix_csv_path, limit=matrix_csv_limit)
        except Exception as exc:
            app.logger.warning("Matrix CSV export skipped: %s", exc)
    
    return jsonify({"aircraft": processed_list})


@app.route("/api/training/matrix")
def training_matrix():
    limit = int(os.getenv("FLIGHT_DB_MATRIX_LIMIT", "500"))
    rows = flight_store.get_latest_matrix(limit=limit)
    return jsonify({"rows": rows, "count": len(rows)})


@app.route("/api/aircraft/info")
def aircraft_info():
    icao = request.args.get("icao", "")
    info = lookup_aircraft_info(icao)
    if info is None:
        return jsonify({"found": False}), 404
    return jsonify({"found": True, **info})


@app.route("/api/route")
def flight_route():
    callsign = request.args.get("callsign", "")
    route = lookup_flight_route(callsign)
    if route is None:
        return jsonify({"found": False}), 404
    return jsonify({"found": True, **route})


@app.route("/api/training/matrix/csv-info")
def training_matrix_csv_info():
    return jsonify(
        {
            "csvEnabled": matrix_csv_enabled,
            "csvPath": matrix_csv_path,
            "csvLimit": matrix_csv_limit,
        }
    )


@app.route("/api/status")
def app_status():
    """Live health/mode status — polled by the UI to show SIMULATION vs LIVE badge."""
    return jsonify(get_status())


@app.route("/api/mode", methods=["POST"])
def update_mode():
    payload = request.get_json(silent=True) or {}
    value = payload.get("simulation")
    if not isinstance(value, bool):
        return jsonify({"error": "simulation must be true or false"}), 400
    try:
        return jsonify(set_simulation_mode(value))
    except Exception as exc:
        app.logger.exception("Could not switch data source mode")
        return jsonify({"error": str(exc) or "Could not switch mode"}), 500


@app.route("/api/logging/status")
def logging_status():
    source_status = get_status()
    return jsonify(
        {
            "icaoExcelEnabled": bool(getattr(flight_store, "icao_excel_enabled", False)),
            "icaoExcelDir": str(getattr(flight_store, "icao_excel_dir", "")),
            "icaoMasterEnabled": bool(getattr(flight_store, "icao_master_enabled", False)),
            "icaoMasterPath": str(getattr(flight_store, "icao_master_path", "")),
            "matrixCsvEnabled": matrix_csv_enabled,
            "matrixCsvPath": matrix_csv_path,
            "sessionDir": str(SESSION_DIR),
            "hexFile": source_status.get("hexFile"),
            "asciiFile": source_status.get("asciiFile"),
            "rawFile": source_status.get("rawFile"),
        }
    )


@app.route("/api/training/trajectory")
def training_trajectory():
    flight_key = request.args.get("flightKey")
    hours = int(request.args.get("hours", "24"))
    limit = int(request.args.get("limit", "200000"))
    rows = flight_store.get_trajectory_history(flight_key=flight_key, hours=hours, limit=limit)
    return jsonify({"rows": rows, "count": len(rows), "flightKey": flight_key or "ALL", "hours": hours})


@app.route("/api/training/trajectory/export")
def training_trajectory_export():
    flight_key = request.args.get("flightKey")
    hours = int(request.args.get("hours", "24"))
    limit = int(request.args.get("limit", "200000"))
    stamp = datetime.utcnow().strftime("%Y%m%dT%H%M%SZ")
    flight_tag = (flight_key or "ALL").upper()
    out_path = (
        app_dir()
        / "logs"
        / "exports"
        / f"trajectory_{flight_tag}_{hours}h_{stamp}.csv"
    )
    csv_path, row_count = flight_store.write_trajectory_history_csv(
        output_path=out_path,
        flight_key=flight_key,
        hours=hours,
        limit=limit,
    )
    return jsonify({"csvPath": csv_path, "rowCount": row_count, "flightKey": flight_tag, "hours": hours})

if __name__ == "__main__":
    port = int(os.getenv("ADSB_PORT", "5005"))
    if os.getenv("WERKZEUG_RUN_MAIN") != "true":
        threading.Timer(1.5, lambda: webbrowser.open(f"http://localhost:{port}/")).start()
    app.run(host="localhost", port=port, debug=False)
