import configparser
import os
import re
import threading
from datetime import datetime, timedelta, timezone
from pathlib import Path

from preprocessing.pingstation_raw_decoder import (
    PingstationRawCaptureLogger,
    PingstationRawDecoder,
    PingstationStreamReceiver,
)
from runtime_paths import app_dir, resource_dir


# ---------------------------------------------------------------------------
# Config-file support
# A file called  ADSBTracker.cfg  placed next to the exe (or in the project
# root during development) can override any of the env-var defaults below.
# If the file does not exist it is created automatically with full comments.
# ---------------------------------------------------------------------------
_CFG_PATH = app_dir() / "ADSBTracker.cfg"
_CFG_TEMPLATE = """\
; ADS-B Tracker configuration file
; Edit this file to switch between simulation and live mode.
; Restart the application after making changes.
;
; ------------------------------------------------------------
; MODE SELECTION
; ------------------------------------------------------------
[mode]
; Set to false to receive LIVE data from the PingStation device.
; Set to true  to replay the bundled sample log file (no hardware needed).
simulation_mode = true

; ------------------------------------------------------------
; LIVE MODE  (only used when simulation_mode = false)
; ------------------------------------------------------------
[live]
; The PingStation device must be configured (via its web interface) to send
; its RAW data stream to THIS PC's IP address on the port below.
; The tracker opens a TCP server socket and waits for the PingStation to
; connect to it.
listen_host = 0.0.0.0
listen_port = 30003
; Seconds before an aircraft that stops broadcasting is removed from the map.
stale_seconds = 45
; Save a copy of every received byte to  logs/sessions/<session>/
capture_enabled = true

; ------------------------------------------------------------
; SIMULATION MODE  (only used when simulation_mode = true)
; ------------------------------------------------------------
[simulation]
; Path to the hex log file to replay (leave blank to use the bundled sample).
source_file =
; How many log entries to process per refresh cycle.
replay_batch = 20
; Loop the replay file indefinitely (true) or stop at end (false).
replay_loop = false
; Seconds before a replayed aircraft is considered stale and removed.
stale_seconds = 12
"""

def _load_cfg():
    """Return a ConfigParser loaded from ADSBTracker.cfg, creating it if absent."""
    if not _CFG_PATH.exists():
        try:
            _CFG_PATH.write_text(_CFG_TEMPLATE, encoding="utf-8")
            print(f"[Config] Created default configuration file: {_CFG_PATH}")
        except OSError as exc:
            print(f"[Config] Could not write config file: {exc}")
    cfg = configparser.ConfigParser()
    cfg.read(str(_CFG_PATH), encoding="utf-8")
    return cfg

def _cfg_get(cfg, section, key, fallback_env, default):
    """Read from config file, fall back to env var, then to hard-coded default."""
    env_val = os.getenv(fallback_env)
    if env_val is not None:
        return env_val          # env var always wins over file
    try:
        return cfg.get(section, key)
    except (configparser.NoSectionError, configparser.NoOptionError):
        return default

_cfg = _load_cfg()


DEFAULT_SIMULATION_LOG_FILE = str(
    resource_dir() / "pingstation_data_log" / "pingstation_data_log" / "HEX DATA_H.txt"
)

IST = timezone(timedelta(hours=5, minutes=30))
_SESSION_STAMP = datetime.now(IST).strftime("%Y%m%d_%H%M%S")
SESSION_DIR = Path(
    os.getenv(
        "PINGSTATION_SESSION_DIR",
        str(app_dir() / "logs" / "sessions" / f"session_{_SESSION_STAMP}"),
    )
)
RAW_HEX_DIR = SESSION_DIR / "hex_data"
RAW_ASCII_DIR = SESSION_DIR / "ascii_data"
ICAO_EXCEL_DIR = SESSION_DIR / "icao_excel"
RAW_HEX_DIR.mkdir(parents=True, exist_ok=True)
RAW_ASCII_DIR.mkdir(parents=True, exist_ok=True)
ICAO_EXCEL_DIR.mkdir(parents=True, exist_ok=True)

# ---------------------------------------------------------------------------
# Effective settings  (config file â†’ env var â†’ hard-coded default)
# ---------------------------------------------------------------------------
SIMULATION_MODE = _cfg_get(_cfg, "mode", "simulation_mode",
                            "SIMULATION_MODE", "true").strip().lower() == "true"
SIMULATION_REPLAY_BATCH = int(_cfg_get(_cfg, "simulation", "replay_batch",
                                        "SIMULATION_REPLAY_BATCH", "20"))
SIMULATION_REPLAY_LOOP = _cfg_get(_cfg, "simulation", "replay_loop",
                                   "SIMULATION_REPLAY_LOOP", "false").strip().lower() == "true"
SIMULATION_STALE_SECONDS = int(_cfg_get(_cfg, "simulation", "stale_seconds",
                                         "SIMULATION_STALE_SECONDS", "12"))

PINGSTATION_LISTEN_HOST = _cfg_get(_cfg, "live", "listen_host",
                                    "PINGSTATION_LISTEN_HOST", "0.0.0.0")
PINGSTATION_LISTEN_PORT = int(_cfg_get(_cfg, "live", "listen_port",
                                        "PINGSTATION_LISTEN_PORT", "30003"))
PINGSTATION_STALE_SECONDS = int(_cfg_get(_cfg, "live", "stale_seconds",
                                          "PINGSTATION_STALE_SECONDS", "45"))
PINGSTATION_CAPTURE_ENABLED = _cfg_get(_cfg, "live", "capture_enabled",
                                        "PINGSTATION_CAPTURE_ENABLED", "true").strip().lower() == "true"
PINGSTATION_CAPTURE_FILE = os.getenv(
    "PINGSTATION_CAPTURE_FILE",
    str(SESSION_DIR),
)

print(
    f"[Config] Mode: {'SIMULATION' if SIMULATION_MODE else 'LIVE'}"
    + (f" | Listening on port {PINGSTATION_LISTEN_PORT}" if not SIMULATION_MODE else "")
    + f" | Config: {_CFG_PATH}"
)


# Single shared decoder instance/algorithm for both simulation replay and the
# live PingStation feed â€” simulation exists only to verify this decoder works
# correctly, so it must run through the exact same code path.
_decoder = PingstationRawDecoder()
_receiver = None
_capture_logger = (
    PingstationRawCaptureLogger(PINGSTATION_CAPTURE_FILE)
    if PINGSTATION_CAPTURE_ENABLED or SIMULATION_MODE
    else None
)
_replay_entries = []
_replay_index = 0
_replay_file_mtime = None
_replay_file_path = None
_mode_lock = threading.RLock()


def _ensure_capture_logger():
    global _capture_logger
    if _capture_logger is None:
        _capture_logger = PingstationRawCaptureLogger(PINGSTATION_CAPTURE_FILE)
    return _capture_logger


def set_simulation_mode(enabled):
    """Switch data sources immediately and persist the choice beside the app/exe."""
    global SIMULATION_MODE, _receiver, _replay_index

    enabled = bool(enabled)
    with _mode_lock:
        if _receiver is not None:
            _receiver.stop()
            _receiver = None

        SIMULATION_MODE = enabled
        _decoder.reset_state()
        _replay_index = 0
        if enabled or PINGSTATION_CAPTURE_ENABLED:
            _ensure_capture_logger()

        if not _cfg.has_section("mode"):
            _cfg.add_section("mode")
        _cfg.set("mode", "simulation_mode", "true" if enabled else "false")
        try:
            config_text = _CFG_PATH.read_text(encoding="utf-8")
            replacement = f"simulation_mode = {'true' if enabled else 'false'}"
            updated_text, replacements = re.subn(
                r"(?im)^\s*simulation_mode\s*=\s*.*$",
                replacement,
                config_text,
                count=1,
            )
            if replacements == 0:
                updated_text = config_text.rstrip() + f"\n\n[mode]\n{replacement}\n"
            _CFG_PATH.write_text(updated_text, encoding="utf-8")
        except OSError as exc:
            print(f"[Config] Could not persist simulation mode: {exc}")

    return get_status()


def _get_simulation_log_file():
    return os.getenv("SIMULATION_SOURCE_FILE", DEFAULT_SIMULATION_LOG_FILE).strip()


def _is_useful_record(record):
    lat = float(record.get("latDD", 0) or 0)
    lon = float(record.get("lonDD", 0) or 0)
    return not (lat == 0 and lon == 0)


def _clean_records(records):
    filtered = [row for row in records if _is_useful_record(row)]
    by_callsign = {}

    for row in filtered:
        callsign = str(row.get("callsign", "")).strip().upper()
        icao = str(row.get("icaoAddress", "")).strip().upper()
        key = icao if icao and icao != "UNKNOWN" else callsign
        if not key:
            continue

        current = by_callsign.get(key)
        if current is None:
            by_callsign[key] = row
            continue

        def score(rec):
            lat = float(rec.get("latDD", 0) or 0)
            lon = float(rec.get("lonDD", 0) or 0)
            altitude = float(rec.get("altitudeMM", 0) or 0)
            speed = float(rec.get("horVelocityCMS", 0) or 0)
            heading = float(rec.get("headingDE2", 0) or 0)
            squawk = int(rec.get("squawk", 0) or 0)
            frame_len = int(rec.get("frameLength", 0) or 0)
            coord_score = 100 if (lat != 0 and lon != 0) else 0
            altitude_score = 10 if altitude > 0 else 0
            speed_score = 5 if speed > 0 else 0
            heading_score = 5 if heading > 0 else 0
            squawk_score = 2 if squawk > 0 else 0
            return coord_score + altitude_score + speed_score + heading_score + squawk_score + frame_len

        if score(row) >= score(current):
            by_callsign[key] = row

    return list(by_callsign.values())


def _load_replay_entries():
    global _replay_entries, _replay_index, _replay_file_mtime, _replay_file_path

    path = Path(_get_simulation_log_file())
    if not path.exists():
        _replay_entries = []
        _replay_index = 0
        return

    mtime = path.stat().st_mtime
    path_str = str(path.resolve())
    should_reparse = (
        _replay_file_mtime is None
        or _replay_file_mtime != mtime
        or _replay_file_path != path_str
        or not _replay_entries
    )
    if not should_reparse:
        return

    _replay_file_mtime = mtime
    _replay_file_path = path_str
    _replay_entries = _decoder.parse_log_entries(str(path), max_entries=None)
    _replay_index = 0
    _decoder.reset_state()

    # Warm up the state so the very first API response has richer fields
    # (heading/speed/squawk) for flights that start with sparse packets.
    # Reset afterwards so real replay (starting again at index 0) builds each
    # aircraft's timestamp forward from its first appearance â€” otherwise the
    # warmup leaves "timestamp" at each aircraft's *last* appearance, and the
    # very first real batch immediately overwrites it with an earlier
    # timestamp, storing an out-of-order trajectory point that later makes
    # the predictor extrapolate from the wrong segment of the route.
    warmup = len(_replay_entries)
    for i in range(warmup):
        raw_time, source_ip, payload, callsign_hint = _replay_entries[i]
        _decoder.decode_payload(
            payload,
            source_ip=source_ip,
            timestamp=raw_time,
            callsign_hint=callsign_hint,
        )
    _decoder.reset_state()
    _replay_index = 0


def _get_replayed_snapshot():
    global _replay_index

    _load_replay_entries()
    if not _replay_entries:
        return [], [], []

    batch = max(1, SIMULATION_REPLAY_BATCH)
    processed = 0
    decoded_events = []

    while processed < batch:
        if _replay_index >= len(_replay_entries):
            if not SIMULATION_REPLAY_LOOP:
                break
            _replay_index = 0
            _decoder.reset_state()

        raw_time, source_ip, payload, callsign_hint = _replay_entries[_replay_index]
        _replay_index += 1
        if _capture_logger is not None:
            _capture_logger.log_chunk(payload, source_ip, 0)
        decoded = _decoder.decode_payload(
            payload,
            source_ip=source_ip,
            timestamp=raw_time,
            callsign_hint=callsign_hint,
        )
        decoded_events.extend(decoded)
        if _capture_logger is not None:
            _capture_logger.log_decoded_many(decoded)
        processed += 1

    stale_seconds = None if SIMULATION_STALE_SECONDS <= 0 else SIMULATION_STALE_SECONDS
    raw_rows = _decoder.get_state_snapshot(stale_seconds=stale_seconds)
    return _clean_records(raw_rows), raw_rows, decoded_events


def _start_receiver_if_needed():
    global _receiver, _capture_logger
    if _receiver is not None:
        return

    _decoder.reset_state()
    _receiver = PingstationStreamReceiver(
        host=PINGSTATION_LISTEN_HOST,
        port=PINGSTATION_LISTEN_PORT,
        decoder=_decoder,
        stale_seconds=PINGSTATION_STALE_SECONDS,
        raw_logger=_capture_logger,
    )
    _receiver.start()


def get_aircraft_data():
    with _mode_lock:
        simulation_mode = SIMULATION_MODE

    if simulation_mode:
        cleaned_rows, raw_rows, decoded_events = _get_replayed_snapshot()
        return {
            "aircraft": cleaned_rows,
            "aircraft_raw": raw_rows,
            "aircraft_events": decoded_events,
        }

    _start_receiver_if_needed()
    live_rows = _receiver.get_snapshot() if _receiver else []
    live_events = _receiver.drain_records() if _receiver else []
    return {
        "aircraft": _clean_records(live_rows),
        "aircraft_raw": live_rows,
        "aircraft_events": live_events,
    }


def get_status():
    """Return a dict that describes the current data-source state.
    Called by the /api/status Flask route so the UI can show mode & health.
    """
    import time as _time

    with _mode_lock:
        simulation_mode = SIMULATION_MODE

    if simulation_mode:
        log_file = _get_simulation_log_file()
        return {
            "mode": "simulation",
            "simulationFile": log_file,
            "simulationFileExists": Path(log_file).exists(),
            "replayIndex": _replay_index,
            "replayTotal": len(_replay_entries),
            "configFile": str(_CFG_PATH),
            "sessionDir": str(SESSION_DIR),
            "hexFile": str(_capture_logger.hex_file_path) if _capture_logger else None,
            "asciiFile": str(_capture_logger.ascii_file_path) if _capture_logger else None,
            "rawFile": str(_capture_logger.raw_file_path) if _capture_logger else None,
        }

    # --- live mode ---
    _start_receiver_if_needed()
    conn_status = _receiver.get_connection_status() if _receiver else {}
    last_data = conn_status.get("lastDataTime")
    seconds_since_data = (
        round(_time.time() - last_data, 1) if last_data is not None else None
    )
    return {
        "mode": "live",
        "connected": conn_status.get("connected", False),
        "listenHost": conn_status.get("listenHost", PINGSTATION_LISTEN_HOST),
        "listenPort": conn_status.get("listenPort", PINGSTATION_LISTEN_PORT),
        "lastDataSecondsAgo": seconds_since_data,
        "totalFramesDecoded": conn_status.get("totalFramesDecoded", 0),
        "totalConnections": conn_status.get("totalConnections", 0),
        "bindError": conn_status.get("bindError"),
        "configFile": str(_CFG_PATH),
        "sessionDir": str(SESSION_DIR),
        "hexFile": str(_capture_logger.hex_file_path) if _capture_logger else None,
        "asciiFile": str(_capture_logger.ascii_file_path) if _capture_logger else None,
        "rawFile": str(_capture_logger.raw_file_path) if _capture_logger else None,
    }

