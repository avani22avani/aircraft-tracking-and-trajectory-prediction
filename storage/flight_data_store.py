import math
import sqlite3
import threading
import csv
import sys
import importlib
import io
from collections import defaultdict
from datetime import datetime, timedelta, timezone
from pathlib import Path

def _try_import_openpyxl():
    try:
        mod = importlib.import_module("openpyxl")
        return getattr(mod, "Workbook"), getattr(mod, "load_workbook"), True
    except ModuleNotFoundError:
        pass

    # Fallback: project-local virtualenv site-packages
    project_root = Path(__file__).resolve().parents[1]
    venv_site = project_root / "venv" / "Lib" / "site-packages"
    if venv_site.exists():
        venv_site_str = str(venv_site)
        if venv_site_str not in sys.path:
            sys.path.insert(0, venv_site_str)
        try:
            mod = importlib.import_module("openpyxl")
            return getattr(mod, "Workbook"), getattr(mod, "load_workbook"), True
        except ModuleNotFoundError:
            pass

    return None, None, False


Workbook, load_workbook, _OPENPYXL_AVAILABLE = _try_import_openpyxl()


def _now_iso():
    return datetime.now(timezone.utc).isoformat()


def _parse_ts(value):
    if value in (None, ""):
        return _now_iso(), None
    text = str(value).strip()
    if not text:
        return _now_iso(), None
    try:
        dt = datetime.fromisoformat(text.replace("Z", "+00:00"))
        if dt.tzinfo is None:
            dt = dt.replace(tzinfo=timezone.utc)
        return dt.isoformat(), dt
    except ValueError:
        try:
            dt = datetime.strptime(text, "%Y-%m-%d %H:%M:%S").replace(tzinfo=timezone.utc)
            return dt.isoformat(), dt
        except ValueError:
            return _now_iso(), None


def _to_int(value, default=0):
    try:
        return int(round(float(value)))
    except (TypeError, ValueError):
        return default


def _to_float(value, default=0.0):
    try:
        return float(value)
    except (TypeError, ValueError):
        return default


def _haversine_m(lat1, lon1, lat2, lon2):
    r = 6371000.0
    p1 = math.radians(lat1)
    p2 = math.radians(lat2)
    dlat = math.radians(lat2 - lat1)
    dlon = math.radians(lon2 - lon1)
    a = math.sin(dlat / 2) ** 2 + math.cos(p1) * math.cos(p2) * math.sin(dlon / 2) ** 2
    return 2 * r * math.atan2(math.sqrt(a), math.sqrt(1 - a))


def _bearing_deg(lat1, lon1, lat2, lon2):
    p1 = math.radians(lat1)
    p2 = math.radians(lat2)
    dlon = math.radians(lon2 - lon1)
    y = math.sin(dlon) * math.cos(p2)
    x = math.cos(p1) * math.sin(p2) - math.sin(p1) * math.cos(p2) * math.cos(dlon)
    deg = (math.degrees(math.atan2(y, x)) + 360.0) % 360.0
    return deg


class FlightTrajectoryStore:
    """
    Compact training-data store.
    - flight_latest: matrix-like table (one row per flight, fields in columns).
    - trajectory_samples: time-series points for trajectory prediction.
    """

    def __init__(
        self,
        db_path,
        enabled=True,
        min_interval_sec=5,
        min_distance_m=30,
        min_altitude_change_m=20,
        min_heading_change_de2=500,
        min_speed_change_cms=200,
        history_retention_days=30,
        append_csv_enabled=True,
        append_csv_dir=None,
        icao_excel_enabled=True,
        icao_excel_dir=None,
        icao_master_enabled=True,
        icao_master_path=None,
        reset_excel_on_start=True,
        icao_excel_extension="xlsx",
    ):
        self.enabled = bool(enabled)
        self.db_path = Path(db_path)
        self.min_interval_sec = int(min_interval_sec)
        self.min_distance_m = float(min_distance_m)
        self.min_altitude_change_m = int(min_altitude_change_m)
        self.min_heading_change_de2 = int(min_heading_change_de2)
        self.min_speed_change_cms = int(min_speed_change_cms)
        self.history_retention_days = int(history_retention_days)
        self.append_csv_enabled = bool(append_csv_enabled)
        self.append_csv_dir = Path(append_csv_dir) if append_csv_dir else self.db_path.parent / "training_csv"
        self.icao_excel_enabled = bool(icao_excel_enabled) and _OPENPYXL_AVAILABLE
        self.icao_excel_dir = Path(icao_excel_dir) if icao_excel_dir else self.db_path.parent / "icao_excel"
        ext = str(icao_excel_extension or "xlsx").strip().lower().lstrip(".")
        self.icao_excel_extension = ext if ext in ("xls", "xlsx") else "xlsx"
        self.icao_master_enabled = bool(icao_master_enabled) and self.icao_excel_enabled
        self.reset_excel_on_start = bool(reset_excel_on_start)
        self.icao_master_path = Path(icao_master_path) if icao_master_path else None
        self._master_wb = None
        self._master_headers = [
            "timestamp",
            "ICAO",
            "callsign",
            "latitude",
            "longitude",
            "altitude",
            "heading",
            "speed",
            "vertical_rate",
        ]

        self._lock = threading.Lock()
        self._last_saved = {}
        self._last_ingest_wallclock = {}
        self._sample_counter = {}
        self._last_purge_day = None
        self._flight_has_lat_e5 = False
        self._flight_has_lon_e5 = False
        self._traj_has_lat_e5 = False
        self._traj_has_lon_e5 = False

        self.db_path.parent.mkdir(parents=True, exist_ok=True)
        self.append_csv_dir.mkdir(parents=True, exist_ok=True)
        self.icao_excel_dir.mkdir(parents=True, exist_ok=True)
        if self.icao_excel_enabled:
            self._init_excel_logging_workspace()
        self._conn = sqlite3.connect(str(self.db_path), check_same_thread=False)
        self._conn.row_factory = sqlite3.Row
        self._init_db()

    def _init_excel_logging_workspace(self):
        self.icao_excel_dir.mkdir(parents=True, exist_ok=True)

        if self.reset_excel_on_start:
            for pattern in ("*.xlsx", "*.xls"):
                for p in self.icao_excel_dir.glob(pattern):
                    try:
                        p.unlink()
                    except OSError:
                        continue

        if not self.icao_master_enabled:
            return

        if self.icao_master_path is None:
            stamp = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")
            self.icao_master_path = self.icao_excel_dir / f"MASTER_ICAO_{stamp}.xlsx"
        self._load_or_create_master_wb()

    def _load_or_create_master_wb(self):
        if not self.icao_master_enabled or self.icao_master_path is None:
            return
        try:
            if self.icao_master_path.exists():
                self._master_wb = self._load_workbook_path(self.icao_master_path)
            else:
                self._master_wb = Workbook()
                ws = self._master_wb.active
                ws.title = "ICAO_INDEX"
                ws.append(["ICAO", "SheetName", "RecordCount"])
                self._save_workbook_path(self._master_wb, self.icao_master_path)
        except Exception:
            self._master_wb = None
            self.icao_master_enabled = False

    def _load_workbook_path(self, file_path):
        file_path = Path(file_path)
        if file_path.suffix.lower() == ".xls":
            with file_path.open("rb") as f:
                data = f.read()
            return load_workbook(io.BytesIO(data))
        return load_workbook(file_path)

    def _save_workbook_path(self, workbook, file_path):
        file_path = Path(file_path)
        if file_path.suffix.lower() == ".xls":
            buf = io.BytesIO()
            workbook.save(buf)
            file_path.write_bytes(buf.getvalue())
            return
        workbook.save(file_path)

    def _init_db(self):
        cur = self._conn.cursor()
        cur.execute("PRAGMA journal_mode=WAL")
        cur.execute("PRAGMA synchronous=NORMAL")
        cur.execute("PRAGMA temp_store=MEMORY")
        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS flight_latest (
                flight_key TEXT PRIMARY KEY,
                icao_address TEXT NOT NULL,
                callsign TEXT NOT NULL,
                timestamp TEXT NOT NULL,
                lat_e6 INTEGER NOT NULL,
                lon_e6 INTEGER NOT NULL,
                altitude_m INTEGER NOT NULL,
                speed_cms INTEGER NOT NULL,
                heading_de2 INTEGER NOT NULL,
                vs_cms INTEGER NOT NULL,
                squawk INTEGER,
                source_ip TEXT,
                sample_count INTEGER NOT NULL DEFAULT 0
            )
            """
        )
        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS trajectory_samples (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                flight_key TEXT NOT NULL,
                timestamp TEXT NOT NULL,
                lat_e6 INTEGER NOT NULL,
                lon_e6 INTEGER NOT NULL,
                altitude_m INTEGER NOT NULL,
                speed_cms INTEGER NOT NULL,
                heading_de2 INTEGER NOT NULL,
                vs_cms INTEGER NOT NULL,
                squawk INTEGER,
                source_ip TEXT,
                UNIQUE(flight_key, timestamp)
            )
            """
        )
        cur.execute(
            "CREATE INDEX IF NOT EXISTS idx_traj_flight_time ON trajectory_samples(flight_key, timestamp)"
        )
        self._migrate_lat_precision_schema(cur)
        self._refresh_schema_flags(cur)
        self._conn.commit()

    def reset_for_new_session(self):
        """
        Wipe trajectory/prediction tables and in-memory caches.
        Called at startup in simulation mode so stale flight paths from a
        previous run can never be mixed with the current run's timestamps
        (which would make the predictor extrapolate from the wrong segment
        of a route and appear to "trace back").
        """
        if not self.enabled:
            return
        with self._lock:
            cur = self._conn.cursor()
            cur.execute("DELETE FROM trajectory_samples")
            cur.execute("DELETE FROM flight_latest")
            self._conn.commit()
            self._last_saved.clear()
            self._last_ingest_wallclock.clear()
            self._sample_counter.clear()

    def _table_columns(self, cur, table_name):
        cur.execute(f"PRAGMA table_info({table_name})")
        return {row[1] for row in cur.fetchall()}

    def _migrate_lat_precision_schema(self, cur):
        """
        Upgrade existing DBs from lat_e5/lon_e5 to lat_e6/lon_e6 without data loss.
        """
        # flight_latest
        cols = self._table_columns(cur, "flight_latest")
        if "lat_e6" not in cols:
            cur.execute("ALTER TABLE flight_latest ADD COLUMN lat_e6 INTEGER")
        if "lon_e6" not in cols:
            cur.execute("ALTER TABLE flight_latest ADD COLUMN lon_e6 INTEGER")

        cols = self._table_columns(cur, "flight_latest")
        if "lat_e5" in cols and "lon_e5" in cols:
            cur.execute(
                """
                UPDATE flight_latest
                SET
                    lat_e6 = COALESCE(lat_e6, lat_e5 * 10),
                    lon_e6 = COALESCE(lon_e6, lon_e5 * 10)
                """
            )

        # trajectory_samples
        cols = self._table_columns(cur, "trajectory_samples")
        if "lat_e6" not in cols:
            cur.execute("ALTER TABLE trajectory_samples ADD COLUMN lat_e6 INTEGER")
        if "lon_e6" not in cols:
            cur.execute("ALTER TABLE trajectory_samples ADD COLUMN lon_e6 INTEGER")

        cols = self._table_columns(cur, "trajectory_samples")
        if "lat_e5" in cols and "lon_e5" in cols:
            cur.execute(
                """
                UPDATE trajectory_samples
                SET
                    lat_e6 = COALESCE(lat_e6, lat_e5 * 10),
                    lon_e6 = COALESCE(lon_e6, lon_e5 * 10)
                """
            )

    def _refresh_schema_flags(self, cur):
        flight_cols = self._table_columns(cur, "flight_latest")
        traj_cols = self._table_columns(cur, "trajectory_samples")
        self._flight_has_lat_e5 = "lat_e5" in flight_cols
        self._flight_has_lon_e5 = "lon_e5" in flight_cols
        self._traj_has_lat_e5 = "lat_e5" in traj_cols
        self._traj_has_lon_e5 = "lon_e5" in traj_cols

    def _quantize(self, row):
        lat = _to_float(row.get("latDD"), 0.0)
        lon = _to_float(row.get("lonDD"), 0.0)
        alt_m = int(round(_to_float(row.get("altitudeMM"), 0.0) / 1000.0))
        return {
            "lat_e6": int(round(lat * 1000000)),
            "lon_e6": int(round(lon * 1000000)),
            "altitude_m": alt_m,
            "speed_cms": _to_int(row.get("horVelocityCMS"), 0),
            "heading_de2": _to_int(row.get("headingDE2"), 0),
            "vs_cms": _to_int(row.get("verVelocityCMS"), 0),
            "squawk": _to_int(row.get("squawk"), 0),
        }

    def _should_store(self, flight_key, ts_iso, q):
        prev = self._last_saved.get(flight_key)
        if prev is None:
            return True

        prev_dt = prev["dt"]
        cur_iso, cur_dt = _parse_ts(ts_iso)
        if cur_dt is None or prev_dt is None:
            return True

        elapsed = (cur_dt - prev_dt).total_seconds()
        if elapsed >= self.min_interval_sec:
            return True

        lat1 = prev["lat_e6"] / 1000000.0
        lon1 = prev["lon_e6"] / 1000000.0
        lat2 = q["lat_e6"] / 1000000.0
        lon2 = q["lon_e6"] / 1000000.0
        dist_m = _haversine_m(lat1, lon1, lat2, lon2)

        if dist_m >= self.min_distance_m:
            return True
        if abs(q["altitude_m"] - prev["altitude_m"]) >= self.min_altitude_change_m:
            return True
        if abs(q["heading_de2"] - prev["heading_de2"]) >= self.min_heading_change_de2:
            return True
        if abs(q["speed_cms"] - prev["speed_cms"]) >= self.min_speed_change_cms:
            return True
        return False

    def ingest_many(self, rows):
        if not self.enabled or not rows:
            return

        with self._lock:
            cur = self._conn.cursor()
            include_flight_lat_e5 = self._flight_has_lat_e5
            include_flight_lon_e5 = self._flight_has_lon_e5
            include_traj_lat_e5 = self._traj_has_lat_e5
            include_traj_lon_e5 = self._traj_has_lon_e5
            appended_rows = []
            for row in rows:
                ingest_wallclock = datetime.now(timezone.utc)
                icao = str(row.get("icaoAddress", "")).strip().upper()
                callsign = str(row.get("callsign", "")).strip().upper()
                if not icao and not callsign:
                    continue
                flight_key = icao if icao else callsign

                ts_iso, ts_dt = _parse_ts(row.get("timestamp"))
                q = self._quantize(row)
                lat_e5 = int(round(q["lat_e6"] / 10.0))
                lon_e5 = int(round(q["lon_e6"] / 10.0))
                source_ip = str(row.get("sourceIp", "")).strip()

                self._sample_counter[flight_key] = self._sample_counter.get(flight_key, 0)
                sample_count = self._sample_counter[flight_key]

                flight_cols = [
                    "flight_key", "icao_address", "callsign", "timestamp",
                    "lat_e6", "lon_e6"
                ]
                flight_vals = [
                    flight_key, icao or "UNKNOWN", callsign or "UNKNOWN", ts_iso,
                    q["lat_e6"], q["lon_e6"]
                ]
                if include_flight_lat_e5:
                    flight_cols.append("lat_e5")
                    flight_vals.append(lat_e5)
                if include_flight_lon_e5:
                    flight_cols.append("lon_e5")
                    flight_vals.append(lon_e5)
                flight_cols.extend(["altitude_m", "speed_cms", "heading_de2", "vs_cms", "squawk", "source_ip", "sample_count"])
                flight_vals.extend([
                    q["altitude_m"],
                    q["speed_cms"],
                    q["heading_de2"],
                    q["vs_cms"],
                    q["squawk"] if q["squawk"] > 0 else None,
                    source_ip,
                    sample_count,
                ])
                update_cols = [c for c in flight_cols if c != "flight_key"]
                cur.execute(
                    f"""
                    INSERT INTO flight_latest ({", ".join(flight_cols)})
                    VALUES ({", ".join(["?"] * len(flight_cols))})
                    ON CONFLICT(flight_key) DO UPDATE SET
                        {", ".join([f"{c}=excluded.{c}" for c in update_cols])}
                    """,
                    tuple(flight_vals),
                )
                self._last_ingest_wallclock[flight_key] = ingest_wallclock

                if q["lat_e6"] == 0 and q["lon_e6"] == 0:
                    continue
                if not self._should_store(flight_key, ts_iso, q):
                    continue

                traj_cols = ["flight_key", "timestamp", "lat_e6", "lon_e6"]
                traj_vals = [flight_key, ts_iso, q["lat_e6"], q["lon_e6"]]
                if include_traj_lat_e5:
                    traj_cols.append("lat_e5")
                    traj_vals.append(lat_e5)
                if include_traj_lon_e5:
                    traj_cols.append("lon_e5")
                    traj_vals.append(lon_e5)
                traj_cols.extend(["altitude_m", "speed_cms", "heading_de2", "vs_cms", "squawk", "source_ip"])
                traj_vals.extend([
                    q["altitude_m"],
                    q["speed_cms"],
                    q["heading_de2"],
                    q["vs_cms"],
                    q["squawk"] if q["squawk"] > 0 else None,
                    source_ip,
                ])
                cur.execute(
                    f"""
                    INSERT OR IGNORE INTO trajectory_samples ({", ".join(traj_cols)})
                    VALUES ({", ".join(["?"] * len(traj_cols))})
                    """,
                    tuple(traj_vals),
                )
                self._sample_counter[flight_key] = self._sample_counter.get(flight_key, 0) + 1
                self._last_saved[flight_key] = {
                    "dt": ts_dt,
                    "lat_e6": q["lat_e6"],
                    "lon_e6": q["lon_e6"],
                    "altitude_m": q["altitude_m"],
                    "speed_cms": q["speed_cms"],
                    "heading_de2": q["heading_de2"],
                }
                appended_rows.append(
                    {
                        "flightKey": flight_key,
                        "timestamp": ts_iso,
                        "icaoAddress": icao or "UNKNOWN",
                        "callsign": callsign or "UNKNOWN",
                        "latDD": q["lat_e6"] / 1000000.0,
                        "lonDD": q["lon_e6"] / 1000000.0,
                        "altitudeM": q["altitude_m"],
                        "speedCMS": q["speed_cms"],
                        "headingDE2": q["heading_de2"],
                        "verVelocityCMS": q["vs_cms"],
                        "squawk": q["squawk"] if q["squawk"] > 0 else None,
                        "sourceIp": source_ip,
                    }
                )

            self._conn.commit()
            if self.append_csv_enabled and appended_rows:
                self._append_trajectory_csv_rows(appended_rows)
            if self.icao_excel_enabled and rows:
                self._append_icao_excel_rows(rows)
            self._purge_old_history_if_needed()

    def get_latest_matrix(self, limit=1000):
        if not self.enabled:
            return []
        limit = max(1, int(limit))
        cur = self._conn.cursor()
        cur.execute(
            """
            SELECT
                flight_key,
                icao_address,
                callsign,
                timestamp,
                lat_e6,
                lon_e6,
                altitude_m,
                speed_cms,
                heading_de2,
                vs_cms,
                squawk,
                source_ip,
                sample_count
            FROM flight_latest
            ORDER BY timestamp DESC
            LIMIT ?
            """,
            (limit,),
        )
        rows = []
        for r in cur.fetchall():
            rows.append(
                {
                    "flightKey": r["flight_key"],
                    "icaoAddress": r["icao_address"],
                    "callsign": r["callsign"],
                    "timestamp": r["timestamp"],
                    "latDD": r["lat_e6"] / 1000000.0,
                    "lonDD": r["lon_e6"] / 1000000.0,
                    "altitudeM": r["altitude_m"],
                    "speedCMS": r["speed_cms"],
                    "headingDE2": r["heading_de2"],
                    "verVelocityCMS": r["vs_cms"],
                    "squawk": r["squawk"],
                    "sourceIp": r["source_ip"],
                    "sampleCount": r["sample_count"],
                }
            )
        return rows

    def write_latest_matrix_csv(self, output_path, limit=5000):
        if not self.enabled:
            return None

        rows = self.get_latest_matrix(limit=limit)
        out_path = Path(output_path)
        out_path.parent.mkdir(parents=True, exist_ok=True)

        headers = [
            "flightKey",
            "icaoAddress",
            "callsign",
            "timestamp",
            "latDD",
            "lonDD",
            "altitudeM",
            "speedCMS",
            "headingDE2",
            "verVelocityCMS",
            "squawk",
            "sourceIp",
            "sampleCount",
        ]

        try:
            with out_path.open("w", newline="", encoding="utf-8") as f:
                writer = csv.DictWriter(f, fieldnames=headers)
                writer.writeheader()
                for row in rows:
                    writer.writerow(row)
        except PermissionError:
            # Common on Windows when CSV is open in Excel; do not break API/logging.
            return None

        return str(out_path)

    def _append_trajectory_csv_rows(self, rows):
        if not rows:
            return
        day_key = datetime.now(timezone.utc).strftime("%Y%m%d")
        csv_path = self.append_csv_dir / f"trajectory_{day_key}.csv"
        headers = [
            "flightKey",
            "timestamp",
            "icaoAddress",
            "callsign",
            "latDD",
            "lonDD",
            "altitudeM",
            "speedCMS",
            "headingDE2",
            "verVelocityCMS",
            "squawk",
            "sourceIp",
        ]
        write_header = not csv_path.exists() or csv_path.stat().st_size == 0
        with csv_path.open("a", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(f, fieldnames=headers)
            if write_header:
                writer.writeheader()
            for row in rows:
                writer.writerow(row)

    def _purge_old_history_if_needed(self):
        if self.history_retention_days <= 0:
            return

        today = datetime.now(timezone.utc).date()
        if self._last_purge_day == today:
            return

        cutoff = datetime.now(timezone.utc) - timedelta(days=self.history_retention_days)
        cutoff_iso = cutoff.isoformat()
        cur = self._conn.cursor()
        cur.execute("DELETE FROM trajectory_samples WHERE timestamp < ?", (cutoff_iso,))
        self._conn.commit()
        self._last_purge_day = today

    def get_trajectory_history(self, flight_key=None, hours=24, limit=200000):
        if not self.enabled:
            return []

        limit = max(1, int(limit))
        hours = max(1, int(hours))
        cutoff_iso = (datetime.now(timezone.utc) - timedelta(hours=hours)).isoformat()

        cur = self._conn.cursor()
        if flight_key:
            cur.execute(
                """
                SELECT flight_key, timestamp, lat_e6, lon_e6, altitude_m, speed_cms, heading_de2,
                       vs_cms, squawk, source_ip
                FROM trajectory_samples
                WHERE flight_key = ? AND timestamp >= ?
                ORDER BY timestamp ASC
                LIMIT ?
                """,
                (str(flight_key).strip().upper(), cutoff_iso, limit),
            )
        else:
            cur.execute(
                """
                SELECT flight_key, timestamp, lat_e6, lon_e6, altitude_m, speed_cms, heading_de2,
                       vs_cms, squawk, source_ip
                FROM trajectory_samples
                WHERE timestamp >= ?
                ORDER BY timestamp ASC
                LIMIT ?
                """,
                (cutoff_iso, limit),
            )

        out = []
        for r in cur.fetchall():
            out.append(
                {
                    "flightKey": r["flight_key"],
                    "timestamp": r["timestamp"],
                    "latDD": r["lat_e6"] / 1000000.0,
                    "lonDD": r["lon_e6"] / 1000000.0,
                    "altitudeM": r["altitude_m"],
                    "speedCMS": r["speed_cms"],
                    "headingDE2": r["heading_de2"],
                    "verVelocityCMS": r["vs_cms"],
                    "squawk": r["squawk"],
                    "sourceIp": r["source_ip"],
                }
            )
        return out

    def write_trajectory_history_csv(self, output_path, flight_key=None, hours=24, limit=200000):
        rows = self.get_trajectory_history(flight_key=flight_key, hours=hours, limit=limit)
        out_path = Path(output_path)
        out_path.parent.mkdir(parents=True, exist_ok=True)
        headers = [
            "flightKey",
            "timestamp",
            "latDD",
            "lonDD",
            "altitudeM",
            "speedCMS",
            "headingDE2",
            "verVelocityCMS",
            "squawk",
            "sourceIp",
        ]
        with out_path.open("w", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(f, fieldnames=headers)
            writer.writeheader()
            for row in rows:
                writer.writerow(row)
        return str(out_path), len(rows)

    def _estimate_motion_from_recent_samples(self, flight_key):
        cur = self._conn.cursor()
        cur.execute(
            """
            SELECT timestamp, lat_e6, lon_e6, altitude_m
            FROM trajectory_samples
            WHERE flight_key = ?
            ORDER BY timestamp DESC
            LIMIT 2
            """,
            (str(flight_key).strip().upper(),),
        )
        points = cur.fetchall()
        if len(points) < 2:
            return None

        p_new = points[0]
        p_old = points[1]
        _, dt_new = _parse_ts(p_new["timestamp"])
        _, dt_old = _parse_ts(p_old["timestamp"])
        if dt_new is None or dt_old is None:
            return None

        dt = (dt_new - dt_old).total_seconds()
        if dt <= 0:
            return None

        lat1 = p_old["lat_e6"] / 1000000.0
        lon1 = p_old["lon_e6"] / 1000000.0
        lat2 = p_new["lat_e6"] / 1000000.0
        lon2 = p_new["lon_e6"] / 1000000.0
        dist_m = _haversine_m(lat1, lon1, lat2, lon2)
        speed_cms = int(round((dist_m / dt) * 100.0))
        heading_de2 = int(round(_bearing_deg(lat1, lon1, lat2, lon2) * 100.0))
        vs_cms = int(round(((p_new["altitude_m"] - p_old["altitude_m"]) / dt) * 100.0))

        return {
            "speed_cms": max(0, speed_cms),
            "heading_de2": max(0, heading_de2),
            "vs_cms": vs_cms,
        }

    def _fit_linear_regression_1d(self, xs, ys):
        n = len(xs)
        if n < 2:
            return None
        x_mean = sum(xs) / n
        y_mean = sum(ys) / n
        sxx = 0.0
        sxy = 0.0
        for x, y in zip(xs, ys):
            dx = x - x_mean
            sxx += dx * dx
            sxy += dx * (y - y_mean)
        if sxx <= 0.0:
            return None
        b = sxy / sxx
        a = y_mean - b * x_mean
        return a, b

    def _predict_from_linear_regression(self, flight_key, age_sec, fallback_lat, fallback_lon, fallback_alt_m):
        """
        Common ML baseline: sliding-window linear regression over recent trajectory points.
        Fits lat(t), lon(t), alt(t) and extrapolates for outage duration.
        """
        cur = self._conn.cursor()
        cur.execute(
            """
            SELECT timestamp, lat_e6, lon_e6, altitude_m
            FROM trajectory_samples
            WHERE flight_key = ?
            ORDER BY timestamp DESC
            LIMIT 20
            """,
            (str(flight_key).strip().upper(),),
        )
        rows = cur.fetchall()
        if len(rows) < 4:
            return None

        pts = []
        for r in reversed(rows):
            _, dt = _parse_ts(r["timestamp"])
            if dt is None:
                continue
            pts.append(
                (
                    dt,
                    r["lat_e6"] / 1000000.0,
                    r["lon_e6"] / 1000000.0,
                    float(r["altitude_m"] or 0.0),
                )
            )
        if len(pts) < 4:
            return None

        t0 = pts[0][0]
        xs = [(p[0] - t0).total_seconds() for p in pts]
        lat_vals = [p[1] for p in pts]
        lon_vals = [p[2] for p in pts]
        alt_vals = [p[3] for p in pts]

        lat_fit = self._fit_linear_regression_1d(xs, lat_vals)
        lon_fit = self._fit_linear_regression_1d(xs, lon_vals)
        alt_fit = self._fit_linear_regression_1d(xs, alt_vals)
        if lat_fit is None or lon_fit is None or alt_fit is None:
            return None

        t_last = xs[-1]
        t_future = t_last + float(age_sec)
        lat_pred = lat_fit[0] + lat_fit[1] * t_future
        lon_pred = lon_fit[0] + lon_fit[1] * t_future
        alt_pred_m = max(0.0, alt_fit[0] + alt_fit[1] * t_future)

        # Derive heading/speed from regression slopes (deg/s -> meters/s approx).
        lat_rate = lat_fit[1]
        lon_rate = lon_fit[1]
        lat_now = lat_fit[0] + lat_fit[1] * t_last
        lon_now = lon_fit[0] + lon_fit[1] * t_last
        lat_next = lat_now + lat_rate
        lon_next = lon_now + lon_rate
        speed_cms = int(round(_haversine_m(lat_now, lon_now, lat_next, lon_next) * 100.0))
        heading_de2 = int(round(_bearing_deg(lat_now, lon_now, lat_next, lon_next) * 100.0))
        vs_cms = int(round(alt_fit[1] * 100.0))

        # Keep fallback stability if regression outputs become non-finite.
        if not (math.isfinite(lat_pred) and math.isfinite(lon_pred)):
            return None
        if not math.isfinite(alt_pred_m):
            alt_pred_m = float(fallback_alt_m)
        if speed_cms < 0:
            speed_cms = 0
        if heading_de2 < 0:
            heading_de2 = 0

        return {
            "lat": lat_pred if math.isfinite(lat_pred) else fallback_lat,
            "lon": lon_pred if math.isfinite(lon_pred) else fallback_lon,
            "alt_m": alt_pred_m,
            "speed_cms": speed_cms,
            "heading_de2": heading_de2,
            "vs_cms": vs_cms,
        }

    def build_predicted_rows(self, active_icao_set=None, max_prediction_age_sec=120, limit=1000):
        """
        Build short-horizon predicted rows for flights that dropped from live feed.
        Predictions are output-only (not written into training tables).
        """
        if not self.enabled:
            return []

        active_icao_set = {str(x).strip().upper() for x in (active_icao_set or set()) if str(x).strip()}
        max_prediction_age_sec = max(1, int(max_prediction_age_sec))
        limit = max(1, int(limit))

        now_dt = datetime.now(timezone.utc)

        cur = self._conn.cursor()
        cur.execute(
            """
            SELECT
                flight_key,
                icao_address,
                callsign,
                timestamp,
                lat_e6,
                lon_e6,
                altitude_m,
                speed_cms,
                heading_de2,
                vs_cms,
                source_ip
            FROM flight_latest
            ORDER BY timestamp DESC
            LIMIT ?
            """,
            (limit,),
        )

        predicted = []
        for r in cur.fetchall():
            icao = str(r["icao_address"] or "").strip().upper()
            if not icao or icao == "UNKNOWN":
                continue
            if icao in active_icao_set:
                continue

            flight_key = str(r["flight_key"] or "").strip().upper()
            last_seen_wallclock = self._last_ingest_wallclock.get(flight_key)
            if last_seen_wallclock is None:
                # Skip aircraft not seen in the current session — prevents data from
                # a previous simulation run leaking into live mode as "predictions".
                continue
            age_sec = (now_dt - last_seen_wallclock).total_seconds()
            if age_sec <= 0 or age_sec > max_prediction_age_sec:
                continue

            lat = r["lat_e6"] / 1000000.0
            lon = r["lon_e6"] / 1000000.0
            altitude_m = int(r["altitude_m"] or 0)
            speed_cms = int(r["speed_cms"] or 0)
            heading_de2 = int(r["heading_de2"] or 0)
            vs_cms = int(r["vs_cms"] or 0)

            reg_pred = self._predict_from_linear_regression(
                flight_key=r["flight_key"],
                age_sec=age_sec,
                fallback_lat=lat,
                fallback_lon=lon,
                fallback_alt_m=altitude_m,
            )
            if reg_pred is not None:
                pred_lat = reg_pred["lat"]
                pred_lon = reg_pred["lon"]
                pred_alt_m = reg_pred["alt_m"]
                speed_cms = reg_pred["speed_cms"]
                heading_de2 = reg_pred["heading_de2"]
                vs_cms = reg_pred["vs_cms"]
            else:
                # Fallback to deterministic kinematics if regression cannot fit.
                if speed_cms <= 0 or heading_de2 <= 0:
                    motion = self._estimate_motion_from_recent_samples(r["flight_key"])
                    if motion is not None:
                        if speed_cms <= 0:
                            speed_cms = motion["speed_cms"]
                        if heading_de2 <= 0:
                            heading_de2 = motion["heading_de2"]
                        if vs_cms == 0:
                            vs_cms = motion["vs_cms"]

                speed_mps = speed_cms / 100.0
                heading_deg = heading_de2 / 100.0
                dist_m = speed_mps * age_sec
                heading_rad = math.radians(heading_deg)

                dlat = (dist_m / 111320.0) * math.cos(heading_rad)
                scale = max(0.2, math.cos(math.radians(lat)))
                dlon = (dist_m / (111320.0 * scale)) * math.sin(heading_rad)

                pred_lat = lat + dlat
                pred_lon = lon + dlon
                pred_alt_m = max(0.0, altitude_m + (vs_cms / 100.0) * age_sec)

            predicted.append(
                {
                    "icaoAddress": icao,
                    "callsign": str(r["callsign"] or "UNKNOWN").strip().upper() or "UNKNOWN",
                    "latDD": pred_lat,
                    "lonDD": pred_lon,
                    "altitudeMM": int(round(pred_alt_m * 1000.0)),
                    "headingDE2": heading_de2,
                    "horVelocityCMS": speed_cms,
                    "verVelocityCMS": vs_cms,
                    "timestamp": now_dt.isoformat(),
                    "sourceIp": "PREDICTED",
                    "isPredicted": True,
                    "predictionAgeSec": int(round(age_sec)),
                }
            )

        return predicted

    def _append_icao_excel_rows(self, rows):
        """
        Create/update one Excel file per ICAO with continuous append.
        File format: <ICAO>.xlsx
        """
        if not rows:
            return

        headers = list(self._master_headers)

        grouped_rows = defaultdict(list)
        for row in rows:
            record = self._to_icao_excel_record(row)
            if record is None:
                continue
            grouped_rows[record["icao"]].append(record)

        for icao, icao_rows in grouped_rows.items():
            file_path = self.icao_excel_dir / f"{icao}.{self.icao_excel_extension}"
            try:
                if file_path.exists():
                    wb = self._load_workbook_path(file_path)
                    ws = wb.active
                else:
                    wb = Workbook()
                    ws = wb.active
                    ws.title = "flight_data"

                self._ensure_excel_headers(ws, headers)
                for record in icao_rows:
                    ws.append(
                        [
                            record["timestamp"],
                            record["icao"],
                            record["callsign"],
                            record["latitude"],
                            record["longitude"],
                            record["altitude"],
                            record["heading"],
                            record["speed"],
                            record["vertical_rate"],
                        ]
                    )
                    ws.cell(row=ws.max_row, column=4).number_format = "0.00000000"
                    ws.cell(row=ws.max_row, column=5).number_format = "0.00000000"
                self._save_workbook_path(wb, file_path)
            except Exception:
                # Keep logging uninterrupted for other ICAO files.
                continue

        self._append_master_excel_rows(grouped_rows)

    def _append_master_excel_rows(self, grouped_rows):
        if not self.icao_master_enabled or not grouped_rows:
            return
        if self._master_wb is None:
            self._load_or_create_master_wb()
        if self._master_wb is None or self.icao_master_path is None:
            return

        try:
            index_ws = self._master_wb["ICAO_INDEX"]
        except KeyError:
            index_ws = self._master_wb.create_sheet("ICAO_INDEX", 0)
            index_ws.append(["ICAO", "SheetName", "RecordCount"])

        index_row_for_icao = {}
        for row_idx in range(2, index_ws.max_row + 1):
            key = str(index_ws.cell(row=row_idx, column=1).value or "").strip().upper()
            if key:
                index_row_for_icao[key] = row_idx

        for icao, icao_rows in grouped_rows.items():
            sheet_name = icao[:31]
            ws = self._master_wb[sheet_name] if sheet_name in self._master_wb.sheetnames else self._master_wb.create_sheet(sheet_name)
            self._ensure_excel_headers(ws, self._master_headers)
            for record in icao_rows:
                ws.append(
                    [
                        record["timestamp"],
                        record["icao"],
                        record["callsign"],
                        record["latitude"],
                        record["longitude"],
                        record["altitude"],
                        record["heading"],
                        record["speed"],
                        record["vertical_rate"],
                    ]
                )
                ws.cell(row=ws.max_row, column=4).number_format = "0.00000000"
                ws.cell(row=ws.max_row, column=5).number_format = "0.00000000"

            current_count = max(0, ws.max_row - 1)
            if icao in index_row_for_icao:
                r = index_row_for_icao[icao]
                index_ws.cell(row=r, column=2, value=sheet_name)
                index_ws.cell(row=r, column=3, value=current_count)
            else:
                index_ws.append([icao, sheet_name, current_count])
                index_row_for_icao[icao] = index_ws.max_row

        try:
            self._save_workbook_path(self._master_wb, self.icao_master_path)
        except Exception:
            pass

    def _ensure_excel_headers(self, ws, headers):
        current = [cell.value for cell in ws[1]] if ws.max_row >= 1 else []
        if current[: len(headers)] == headers:
            return
        if ws.max_row == 1 and not any(current):
            for idx, value in enumerate(headers, start=1):
                ws.cell(row=1, column=idx, value=value)
            return
        looks_like_header = any(
            str(v).strip().lower() in {"timestamp", "icaoaddress", "icao"}
            for v in current
            if v is not None
        )
        if not looks_like_header:
            ws.insert_rows(1)
        for idx, value in enumerate(headers, start=1):
            ws.cell(row=1, column=idx, value=value)

    def _to_icao_excel_record(self, row):
        if not isinstance(row, dict):
            return None

        icao_raw = (
            row.get("icaoAddress")
            or row.get("icao")
            or row.get("hex")
            or ""
        )
        icao = str(icao_raw).strip().upper()
        callsign = str(row.get("callsign", "")).strip().upper()

        if not icao or icao == "UNKNOWN":
            if not callsign:
                return None
            # No decoded Mode-S ICAO - identify this flight by its callsign
            # so it still gets logged to its own excel file/sheet.
            identifier = "CS_" + callsign
        else:
            identifier = icao

        safe_icao = "".join(ch for ch in identifier if ch.isalnum())
        if not safe_icao:
            return None

        ts_iso, _ = _parse_ts(row.get("timestamp"))
        callsign = callsign or "UNKNOWN"

        lat = round(_to_float(row.get("latDD", row.get("lat", row.get("latitude", 0.0))), 0.0), 8)
        lon = round(_to_float(row.get("lonDD", row.get("lon", row.get("longitude", 0.0))), 0.0), 8)

        altitude_mm = _to_float(row.get("altitudeMM"), None)
        altitude_m = (
            round(altitude_mm / 1000.0, 3)
            if altitude_mm is not None
            else _to_float(row.get("altitude"), 0.0)
        )
        heading_de2 = _to_float(row.get("headingDE2"), None)
        heading_deg = (
            round(heading_de2 / 100.0, 2)
            if heading_de2 is not None
            else _to_float(row.get("heading"), 0.0)
        )
        speed_cms = _to_float(row.get("horVelocityCMS"), None)
        speed_mps = (
            round(speed_cms / 100.0, 3)
            if speed_cms is not None
            else _to_float(row.get("speed"), 0.0)
        )
        vs_cms = _to_float(row.get("verVelocityCMS"), None)
        vertical_rate_mps = (
            round(vs_cms / 100.0, 3)
            if vs_cms is not None
            else _to_float(row.get("verticalRate"), 0.0)
        )

        return {
            "timestamp": ts_iso,
            "icao": safe_icao,
            "callsign": callsign,
            "latitude": lat,
            "longitude": lon,
            "altitude": altitude_m,
            "heading": heading_deg,
            "speed": speed_mps,
            "vertical_rate": vertical_rate_mps,
        }
