import os
import re
import sqlite3
import sys
import tempfile
import zipfile


zip_path = sys.argv[1]
with zipfile.ZipFile(zip_path) as archive:
    for name in (
        "adsb_tracker_2/storage/flight_data_store.py",
        "adsb_tracker_2/preprocessing/pingstation_raw_decoder.py",
        "adsb_tracker_2/templates/index.html",
    ):
        source = archive.read(name).decode("utf-8", "replace")
        print(f"\n### {name}")
        if name.endswith(".py"):
            for match in re.finditer(r"^(class\s+\w+|\s*def\s+\w+)", source, re.MULTILINE):
                print(match.group(1).strip())
        else:
            for pattern in (
                r'data-tab="([^"]+)"',
                r'id="([^"]+)"',
                r"<button[^>]*>([^<]+)",
            ):
                print(pattern, sorted(set(re.findall(pattern, source)))[:100])

    print("\n### SELECTED IMPLEMENTATION")
    for name, symbols in (
        (
            "adsb_tracker_2/storage/flight_data_store.py",
            (
                "build_predicted_rows",
                "_predict_from_linear_regression",
                "_append_icao_excel_rows",
                "_append_master_excel_rows",
            ),
        ),
        (
            "adsb_tracker_2/preprocessing/pingstation_raw_decoder.py",
            ("log_chunk", "_decode_compressed_vrs", "split_frames"),
        ),
    ):
        source = archive.read(name).decode("utf-8", "replace")
        lines = source.splitlines()
        for symbol in symbols:
            start = next(
                (
                    index
                    for index, line in enumerate(lines)
                    if re.match(rf"^\s*def\s+{re.escape(symbol)}\b", line)
                ),
                None,
            )
            if start is None:
                continue
            base_indent = len(lines[start]) - len(lines[start].lstrip())
            end = len(lines)
            for index in range(start + 1, len(lines)):
                stripped = lines[index].lstrip()
                indent = len(lines[index]) - len(stripped)
                if stripped and indent <= base_indent and re.match(
                    r"^(def|class)\s+", stripped
                ):
                    end = index
                    break
            print(f"\n{name}:{symbol}")
            print("\n".join(lines[start : min(end, start + 100)]))

    db_bytes = archive.read("adsb_tracker_2/logs/flight_training.db")
    handle = tempfile.NamedTemporaryFile(delete=False, suffix=".db")
    try:
        handle.write(db_bytes)
        handle.close()
        connection = sqlite3.connect(handle.name)
        cursor = connection.cursor()
        print("\n### DATABASE")
        print(
            "tables",
            cursor.execute(
                "SELECT name FROM sqlite_master WHERE type='table' ORDER BY name"
            ).fetchall(),
        )
        for table in ("flight_latest", "trajectory_samples"):
            print(table, cursor.execute(f"SELECT COUNT(*) FROM {table}").fetchone()[0])
        print(
            "latest",
            cursor.execute(
                """
                SELECT flight_key, callsign, altitude_m, speed_cms, sample_count
                FROM flight_latest
                ORDER BY timestamp DESC
                LIMIT 10
                """
            ).fetchall(),
        )
        connection.close()
    finally:
        try:
            os.unlink(handle.name)
        except OSError:
            pass

    names = archive.namelist()
    sessions = {
        part
        for name in names
        for part in name.split("/")
        if part.startswith("session_")
    }
    excel_files = [
        name
        for name in names
        if name.lower().endswith((".xlsx", ".xls"))
        and "/logs/" in name
    ]
    print("\n### ARTIFACTS")
    print("sessions", len(sessions))
    print("excel_files", len(excel_files))
    print("executable", any(name.endswith("dist/ADSBTracker.exe") for name in names))
