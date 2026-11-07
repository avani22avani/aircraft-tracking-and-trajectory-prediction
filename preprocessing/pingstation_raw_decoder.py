import json
import re
import socket
import struct
import threading
import time
from datetime import datetime, timedelta
from pathlib import Path


def _now_iso():
    return datetime.utcnow().isoformat() + "Z"


class PingstationRawCaptureLogger:
    """
    Logs each session at three processing stages:
    - <date_time>_engi_H.txt (hex lines)
    - <date_time>_engi_A.txt (ascii lines)
    - <date_time>_engi_RAW.txt (decoded records before preprocessing)
    """

    def __init__(self, file_path):
        target = Path(file_path)
        self.capture_dir = target if target.suffix == "" else target.parent
        self.capture_dir.mkdir(parents=True, exist_ok=True)
        self.hex_dir = self.capture_dir / "hex_data"
        self.ascii_dir = self.capture_dir / "ascii_data"
        self.hex_dir.mkdir(parents=True, exist_ok=True)
        self.ascii_dir.mkdir(parents=True, exist_ok=True)
        stamp = datetime.utcnow().strftime("%Y%m%d_%H%M%S")
        self.hex_file_path = self.hex_dir / f"{stamp}_engi_H.txt"
        self.ascii_file_path = self.ascii_dir / f"{stamp}_engi_A.txt"
        self.raw_file_path = self.capture_dir / f"{stamp}_engi_RAW.txt"
        self.stream_file_path = self.capture_dir / f"{stamp}_engi_STREAM.bin"
        self.hex_file_path.touch(exist_ok=True)
        self.ascii_file_path.touch(exist_ok=True)
        self.raw_file_path.touch(exist_ok=True)
        self.stream_file_path.touch(exist_ok=True)
        self._lock = threading.Lock()

    def _extract_ascii_tokens(self, payload):
        text = payload.decode("ascii", errors="ignore")
        return re.findall(r"[A-Za-z0-9]{3,}", text)

    def log_chunk(self, payload, source_ip, source_port):
        if not payload:
            return

        hex_data = payload.hex()
        ascii_tokens = self._extract_ascii_tokens(payload)
        ascii_line = ", ".join(ascii_tokens) if ascii_tokens else "None"

        with self._lock:
            with self.hex_file_path.open("a", encoding="utf-8") as hf:
                hf.write(hex_data + "\n")
            with self.ascii_file_path.open("a", encoding="utf-8") as af:
                af.write(ascii_line + "\n")
            # Preserve the exact receiver stream separately for protocol debugging.
            with self.stream_file_path.open("ab") as rf:
                rf.write(payload)

    def log_decoded(self, record):
        """Append one decoded record before preprocessing as JSON Lines text."""
        if not isinstance(record, dict) or not record:
            return
        clean_record = {
            key: value
            for key, value in record.items()
            if not str(key).startswith("_")
        }
        with self._lock:
            with self.raw_file_path.open("a", encoding="utf-8") as rf:
                rf.write(json.dumps(clean_record, ensure_ascii=True, separators=(",", ":")) + "\n")

    def log_decoded_many(self, records):
        for record in records or []:
            self.log_decoded(record)


class PingstationRawDecoder:
    """
    Decodes Pingstation length-prefixed binary frames into structured aircraft records.
    """

    def __init__(self):
        self._state = {}

    # Compressed VRS field flags.
    _F_CALLSIGN = 0x0001
    _F_ALTITUDE = 0x0002
    _F_GROUNDSPEED = 0x0004
    _F_TRACK = 0x0008
    _F_LAT = 0x0010
    _F_LON = 0x0020
    _F_VSPEED = 0x0040
    _F_SQUAWK = 0x0080
    _F_COMP_FLAGS = 0x0F00

    def reset_state(self):
        self._state.clear()

    def split_frames(self, buffer_bytes):
        """
        Splits a byte buffer into complete frames using the first byte as frame length.
        Returns (frames, remainder).
        """
        frames = []
        idx = 0
        total = len(buffer_bytes)

        while idx < total:
            frame_len = buffer_bytes[idx]

            # Discard obvious noise and re-sync.
            if frame_len < 10 or frame_len > 120:
                idx += 1
                continue

            end_idx = idx + frame_len
            if end_idx > total:
                break

            frames.append(buffer_bytes[idx:end_idx])
            idx = end_idx

        return frames, buffer_bytes[idx:]

    def _decode_callsign(self, frame_bytes):
        if len(frame_bytes) >= 18:
            field = frame_bytes[10:18]
            text = field.decode("ascii", errors="ignore").strip().upper()
            text = re.sub(r"[^A-Z0-9]", "", text)
            if text:
                return text

        # Fallback when fixed field is absent/corrupt.
        tokens = re.findall(rb"[A-Z0-9]{3,8}", frame_bytes)
        for token in tokens:
            candidate = token.decode("ascii", errors="ignore").upper()
            if any(ch.isalpha() for ch in candidate):
                return candidate

        return "UNKNOWN"

    def _read_float_int24(self, frame_bytes, idx):
        if idx + 3 > len(frame_bytes):
            return None, idx
        raw = int.from_bytes(frame_bytes[idx:idx + 3], "big")
        idx += 3
        sign = -1 if (raw & 0x800000) else 1
        value = raw & 0x7FFFFF
        return sign * value, idx

    def _decode_compressed_vrs(self, frame_bytes):
        """
        Decode one Compressed VRS message.
        Returns dict of decoded fields or None if frame doesn't match.
        """
        if len(frame_bytes) < 9:
            return None
        if frame_bytes[0] != len(frame_bytes):
            return None

        icao = frame_bytes[4:7].hex().upper()
        flags = int.from_bytes(frame_bytes[7:9], "little")
        idx = 9
        out = {"icaoAddress": icao}

        if flags & self._F_CALLSIGN:
            if idx >= len(frame_bytes):
                return None
            n = frame_bytes[idx]
            idx += 1
            if idx + n > len(frame_bytes):
                return None
            callsign = frame_bytes[idx:idx + n].decode("ascii", errors="ignore").strip().upper()
            callsign = re.sub(r"[^A-Z0-9]", "", callsign)
            idx += n
            if callsign:
                out["callsign"] = callsign

        if flags & self._F_ALTITUDE:
            altitude_ft, idx = self._read_float_int24(frame_bytes, idx)
            if altitude_ft is None:
                return None
            out["altitudeMM"] = int(round(float(altitude_ft) * 304.8))

        if flags & self._F_GROUNDSPEED:
            if idx + 2 > len(frame_bytes):
                return None
            groundspeed_knots = int.from_bytes(frame_bytes[idx:idx + 2], "little", signed=True)
            idx += 2
            if groundspeed_knots >= 0:
                out["horVelocityCMS"] = int(round(groundspeed_knots * 0.514444 * 100))

        if flags & self._F_TRACK:
            if idx + 2 > len(frame_bytes):
                return None
            track_tenths = int.from_bytes(frame_bytes[idx:idx + 2], "little", signed=True)
            idx += 2
            if track_tenths >= 0:
                heading_deg = (track_tenths / 10.0) % 360.0
                out["headingDE2"] = int(round(heading_deg * 100))

        if flags & self._F_LAT:
            if idx + 4 > len(frame_bytes):
                return None
            out["latDD"] = struct.unpack("<f", frame_bytes[idx:idx + 4])[0]
            idx += 4

        if flags & self._F_LON:
            if idx + 4 > len(frame_bytes):
                return None
            out["lonDD"] = struct.unpack("<f", frame_bytes[idx:idx + 4])[0]
            idx += 4

        if flags & self._F_VSPEED:
            if idx + 2 > len(frame_bytes):
                return None
            vs_fpm = int.from_bytes(frame_bytes[idx:idx + 2], "little", signed=True)
            idx += 2
            out["verVelocityCMS"] = int(round(vs_fpm * 0.00508 * 100))

        if flags & self._F_SQUAWK:
            if idx + 2 > len(frame_bytes):
                return None
            out["squawk"] = int.from_bytes(frame_bytes[idx:idx + 2], "little", signed=False)
            idx += 2

        # Packed compressed flags byte (squawk changed / emergency / ident / on-ground).
        if flags & self._F_COMP_FLAGS:
            if idx + 1 > len(frame_bytes):
                return None
            out["compressedFlags"] = frame_bytes[idx]
            idx += 1

        # Minimum sanity: must at least resolve callsign or position/speed info.
        if len(out) == 1:
            return None
        return out

    def _decode_lat_lon(self, frame_bytes, prev_lat=None, prev_lon=None):
        candidates = []

        # Observed Pingstation layout in sample logs:
        # latitude/longitude as little-endian float32 at offsets 25..28 and 29..32.
        if len(frame_bytes) >= 33:
            try:
                lat_le = struct.unpack("<f", frame_bytes[25:29])[0]
                lon_le = struct.unpack("<f", frame_bytes[29:33])[0]
                if -90.0 <= lat_le <= 90.0 and -180.0 <= lon_le <= 180.0:
                    candidates.append(("le_25", lat_le, lon_le))
            except struct.error:
                pass

        # Legacy fallback used by older parser assumptions.
        if len(frame_bytes) >= 36:
            try:
                lat_be = struct.unpack(">f", frame_bytes[28:32])[0]
                lon_be = struct.unpack(">f", frame_bytes[32:36])[0]
                if -90.0 <= lat_be <= 90.0 and -180.0 <= lon_be <= 180.0:
                    candidates.append(("be_28", lat_be, lon_be))
            except struct.error:
                pass

        if not candidates:
            return None, None

        has_prev = (
            prev_lat is not None
            and prev_lon is not None
            and prev_lat != 0
            and prev_lon != 0
        )
        if has_prev:
            # Keep trajectory continuous when multiple layouts look numerically valid.
            def continuity_score(item):
                _, lat, lon = item
                return abs(lat - prev_lat) + abs(lon - prev_lon)

            _, lat, lon = min(candidates, key=continuity_score)
            return lat, lon

        # Prefer observed Pingstation sample layout for first fix.
        for layout_name in ("le_25", "be_28"):
            for name, lat, lon in candidates:
                if name == layout_name:
                    return lat, lon

        _, lat, lon = candidates[0]
        return lat, lon

    def decode_frame(self, frame_bytes, source_ip=None, timestamp=None, callsign_hint=None):
        # Control/housekeeping packets are usually shorter and not useful for aircraft state.
        if len(frame_bytes) < 23:
            return None

        decoded = self._decode_compressed_vrs(frame_bytes)
        if decoded is None:
            # Legacy fallback when frame is not Compressed VRS.
            icao = frame_bytes[4:7].hex().upper() if len(frame_bytes) >= 7 else "UNKNOWN"
            callsign = self._decode_callsign(frame_bytes)
            if callsign_hint:
                callsign = callsign_hint.strip().upper()

            prev = self._state.get(icao if icao and icao != "000000" else callsign, {})
            prev_lat = prev.get("latDD")
            prev_lon = prev.get("lonDD")
            lat, lon = self._decode_lat_lon(frame_bytes, prev_lat=prev_lat, prev_lon=prev_lon)

            decoded = {
                "icaoAddress": icao,
                "callsign": callsign,
            }
            if lat is not None:
                decoded["latDD"] = lat
                decoded["lonDD"] = lon
            if len(frame_bytes) >= 21:
                altitude_ft = int.from_bytes(frame_bytes[19:21], "big")
                if altitude_ft > 0:
                    decoded["altitudeMM"] = int(round(altitude_ft * 304.8))

        if callsign_hint:
            decoded["callsign"] = callsign_hint.strip().upper()

        icao = decoded.get("icaoAddress", "UNKNOWN")
        callsign = decoded.get("callsign", "UNKNOWN")
        record_key = icao if icao and icao != "000000" else callsign
        if not record_key:
            return None

        now_epoch = time.time()
        iso_time = timestamp or _now_iso()

        prev = self._state.get(record_key, {})
        merged = {
            "icaoAddress": icao or prev.get("icaoAddress", "UNKNOWN"),
            "callsign": callsign or prev.get("callsign", "UNKNOWN"),
            "latDD": prev.get("latDD", 0.0),
            "lonDD": prev.get("lonDD", 0.0),
            "altitudeMM": prev.get("altitudeMM", 0),
            "headingDE2": prev.get("headingDE2", 0),
            "horVelocityCMS": prev.get("horVelocityCMS", 0),
            "timestamp": iso_time,
            "frameLength": len(frame_bytes),
            "sourceIp": source_ip or prev.get("sourceIp", ""),
            "_last_seen": now_epoch,
        }

        for key in (
            "latDD",
            "lonDD",
            "altitudeMM",
            "headingDE2",
            "horVelocityCMS",
            "verVelocityCMS",
            "squawk",
            "compressedFlags",
        ):
            if key in decoded and decoded[key] is not None:
                if key in ("headingDE2", "horVelocityCMS", "verVelocityCMS", "squawk"):
                    new_value = decoded[key]
                    old_value = merged.get(key, 0)
                    if new_value == 0 and old_value not in (None, 0):
                        continue
                merged[key] = decoded[key]

        self._state[record_key] = merged
        return dict(merged)

    def decode_payload(self, payload_bytes, source_ip=None, timestamp=None, callsign_hint=None):
        frames, remainder = self.split_frames(payload_bytes)
        decoded = []
        frame_hint = callsign_hint if len(frames) <= 1 else None

        for frame in frames:
            rec = self.decode_frame(
                frame,
                source_ip=source_ip,
                timestamp=timestamp,
                callsign_hint=frame_hint,
            )
            if rec:
                decoded.append(rec)

        # If the whole payload is exactly one frame but not split for some reason.
        if not decoded and not remainder and payload_bytes and payload_bytes[0] == len(payload_bytes):
            rec = self.decode_frame(
                payload_bytes,
                source_ip=source_ip,
                timestamp=timestamp,
                callsign_hint=callsign_hint,
            )
            if rec:
                decoded.append(rec)

        return decoded

    def decode_log_file(self, file_path, max_entries=300, reset_state=False):
        entries = self.parse_log_entries(file_path, max_entries=max_entries)
        if reset_state:
            self.reset_state()

        for raw_time, source_ip, payload, callsign_hint in entries:
            self.decode_payload(
                payload,
                source_ip=source_ip,
                timestamp=raw_time,
                callsign_hint=callsign_hint,
            )

        return self.get_state_snapshot(stale_seconds=None)

    def parse_log_entries(self, file_path, max_entries=300):
        path = Path(file_path)
        if not path.exists():
            return []

        text = path.read_text(encoding="utf-8", errors="ignore")
        # Allow simulation files that include log-banner lines or mixed sections.
        text = text.replace("PINGSTATION 3 LIVE DATA LOG\n", "")
        full_log_pattern = re.compile(
            r"TIME\s*:\s*(?P<time>[^\n]+)\s*.*?"
            r"SOURCE IP\s*:\s*(?P<ip>[^\n]+)\s*.*?"
            r"HEX DATA:\s*\n(?P<hex>[0-9a-fA-F]+)\s*.*?"
            r"ASCII DATA:\s*\n(?P<ascii>[^\n]+)",
            re.DOTALL,
        )
        full_log_hex_only_pattern = re.compile(
            r"TIME\s*:\s*(?P<time>[^\n]+)\s*.*?"
            r"SOURCE IP\s*:\s*(?P<ip>[^\n]+)\s*.*?"
            r"HEX DATA:\s*\n(?P<hex>[0-9a-fA-F]+)",
            re.DOTALL,
        )
        simple_pattern = re.compile(
            r"HEX DATA:\s*\n(?P<hex>[0-9a-fA-F]+)\s*\n\s*ASCII DATA:\s*\n(?P<ascii>[^\n]+)",
            re.DOTALL,
        )

        rows = []
        for match in full_log_pattern.finditer(text):
            raw_time = match.group("time").strip()
            source_ip = match.group("ip").strip()
            hex_data = match.group("hex").strip()
            ascii_data = match.group("ascii").strip().upper()

            try:
                payload = bytes.fromhex(hex_data)
            except ValueError:
                continue

            tokens = [t.strip() for t in ascii_data.split(",") if t.strip()]
            callsign_hint = tokens[0] if len(tokens) == 1 else None
            rows.append((raw_time, source_ip, payload, callsign_hint))

        # Backward compatibility for logs that do not include ASCII section.
        if not rows:
            for match in full_log_hex_only_pattern.finditer(text):
                raw_time = match.group("time").strip()
                source_ip = match.group("ip").strip()
                hex_data = match.group("hex").strip()

                try:
                    payload = bytes.fromhex(hex_data)
                except ValueError:
                    continue

                rows.append((raw_time, source_ip, payload, None))

        if not rows:
            base_time = datetime.utcnow()
            for idx, match in enumerate(simple_pattern.finditer(text)):
                hex_data = match.group("hex").strip()
                ascii_data = match.group("ascii").strip().upper()

                try:
                    payload = bytes.fromhex(hex_data)
                except ValueError:
                    continue

                # ASCII line may contain multiple tokens; first token is the callsign candidate.
                tokens = [t.strip() for t in ascii_data.split(",") if t.strip()]
                callsign_hint = tokens[0] if tokens else None
                timestamp = (base_time + timedelta(seconds=idx)).isoformat() + "Z"
                rows.append((timestamp, "", payload, callsign_hint))

        # New compact format:
        # <hex_line>
        #
        # <ascii_line>
        if not rows:
            hex_line_re = re.compile(r"^[0-9a-fA-F]{10,}$")
            clean_lines = []
            for line in text.splitlines():
                stripped = line.strip()
                if not stripped:
                    continue
                if stripped.startswith("="):
                    continue
                if stripped.upper() == "PINGSTATION 3 LIVE DATA LOG":
                    continue
                clean_lines.append(stripped)

            base_time = datetime.utcnow()
            idx = 0
            event_idx = 0
            while idx < len(clean_lines):
                hex_candidate = clean_lines[idx]
                if not hex_line_re.match(hex_candidate):
                    idx += 1
                    continue

                try:
                    payload = bytes.fromhex(hex_candidate)
                except ValueError:
                    idx += 1
                    continue

                callsign_hint = None
                if idx + 1 < len(clean_lines):
                    ascii_candidate = clean_lines[idx + 1].upper()
                    if not hex_line_re.match(ascii_candidate):
                        tokens = [t.strip() for t in ascii_candidate.split(",") if t.strip()]
                        if tokens and tokens[0] != "NONE":
                            callsign_hint = tokens[0]
                        idx += 1

                timestamp = (base_time + timedelta(seconds=event_idx)).isoformat() + "Z"
                rows.append((timestamp, "", payload, callsign_hint))
                event_idx += 1
                idx += 1

        # Split-file format support:
        # - Hex file: <date_time>_engi_H.txt (one hex payload per line)
        # - Ascii file: matching <date_time>_engi_A.txt (one ascii line per line)
        if not rows:
            hex_line_re = re.compile(r"^[0-9a-fA-F]{10,}$")
            hex_lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
            if hex_lines and all(hex_line_re.match(ln) for ln in hex_lines):
                ascii_lines = []
                ascii_candidates = []
                stem = path.stem
                if stem.endswith("_engi_H"):
                    ascii_candidates.append(path.with_name(path.name.replace("_engi_H.txt", "_engi_A.txt")))
                    if path.parent.name.lower() == "hex_data":
                        ascii_candidates.append(path.parent.parent / "ascii_data" / path.name.replace("_engi_H.txt", "_engi_A.txt"))
                ascii_candidates.append(path.with_name(f"{stem}_A.txt"))
                ascii_candidates.append(path.with_name(f"{stem}_ascii.txt"))

                for ascii_path in ascii_candidates:
                    if ascii_path.exists():
                        ascii_lines = [
                            ln.strip()
                            for ln in ascii_path.read_text(encoding="utf-8", errors="ignore").splitlines()
                            if ln.strip()
                        ]
                        break

                base_time = datetime.utcnow()
                for idx, hex_data in enumerate(hex_lines):
                    try:
                        payload = bytes.fromhex(hex_data)
                    except ValueError:
                        continue

                    callsign_hint = None
                    if idx < len(ascii_lines):
                        ascii_candidate = ascii_lines[idx].upper()
                        tokens = [t.strip() for t in ascii_candidate.split(",") if t.strip()]
                        if tokens and tokens[0] != "NONE":
                            callsign_hint = tokens[0]

                    timestamp = (base_time + timedelta(seconds=idx)).isoformat() + "Z"
                    rows.append((timestamp, "", payload, callsign_hint))

        if max_entries is not None and max_entries > 0:
            rows = rows[-max_entries:]

        return rows

    def get_state_snapshot(self, stale_seconds=45):
        now = time.time()
        rows = []
        stale_keys = []

        for key, value in self._state.items():
            last_seen = value.get("_last_seen", now)
            if stale_seconds is not None and (now - last_seen) > stale_seconds:
                stale_keys.append(key)
                continue

            clean = {k: v for k, v in value.items() if not k.startswith("_")}
            rows.append(clean)

        for key in stale_keys:
            self._state.pop(key, None)

        return rows


class PingstationStreamReceiver:
    """
    TCP listener that receives live Pingstation bytes and keeps a decoded snapshot.
    The PingStation device must be configured to connect TO this PC on the
    configured port (default 30003).  The PC acts as the TCP server.
    """

    def __init__(self, host, port, decoder=None, stale_seconds=45, raw_logger=None):
        self.host = host
        self.port = port
        self.stale_seconds = stale_seconds
        self.decoder = decoder or PingstationRawDecoder()
        self.raw_logger = raw_logger

        self._lock = threading.Lock()
        self._pending_records = []
        self._stop_event = threading.Event()
        self._thread = None
        self._server_socket = None
        self._client_socket = None
        self._socket_lock = threading.Lock()

        # ---- connection-status tracking (exposed via get_connection_status) ----
        self._status_lock = threading.Lock()
        self._is_connected = False          # True while a client socket is open
        self._last_data_time = None         # epoch float, last time bytes arrived
        self._total_frames_decoded = 0
        self._total_connections = 0
        self._bind_error = None             # last OSError message from bind/listen

    # ------------------------------------------------------------------
    def get_connection_status(self):
        """Return a dict describing the live receiver state."""
        with self._status_lock:
            return {
                "connected": self._is_connected,
                "lastDataTime": self._last_data_time,
                "totalFramesDecoded": self._total_frames_decoded,
                "totalConnections": self._total_connections,
                "bindError": self._bind_error,
                "listenPort": self.port,
                "listenHost": self.host,
            }

    # ------------------------------------------------------------------
    def start(self):
        if self._thread and self._thread.is_alive():
            return
        self._thread = threading.Thread(
            target=self._run, name="pingstation-listener", daemon=True
        )
        self._thread.start()

    def stop(self):
        self._stop_event.set()
        with self._socket_lock:
            sockets = (self._client_socket, self._server_socket)
        for sock in sockets:
            if sock is None:
                continue
            try:
                sock.shutdown(socket.SHUT_RDWR)
            except OSError:
                pass
            try:
                sock.close()
            except OSError:
                pass
        if self._thread and self._thread.is_alive():
            self._thread.join(timeout=3)

    # ------------------------------------------------------------------
    def _make_server_socket(self):
        """Create and bind the TCP server socket (platform-aware)."""
        server = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        # On Windows SO_REUSEADDR is dangerous (lets other processes steal the
        # port). Use SO_EXCLUSIVEADDRUSE instead so only one listener owns it.
        if hasattr(socket, "SO_EXCLUSIVEADDRUSE"):          # Windows
            server.setsockopt(socket.SOL_SOCKET, socket.SO_EXCLUSIVEADDRUSE, 1)
        else:                                                # Linux / macOS
            server.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
        server.bind((self.host, self.port))
        server.listen(5)
        server.settimeout(1.0)
        return server

    # ------------------------------------------------------------------
    def _run(self):
        while not self._stop_event.is_set():
            server = None
            try:
                server = self._make_server_socket()
                with self._socket_lock:
                    self._server_socket = server
                with self._status_lock:
                    self._bind_error = None   # cleared — bind succeeded
                print(
                    f"[PingStation] Listening on {self.host}:{self.port} "
                    f"— waiting for PingStation connection …"
                )

                while not self._stop_event.is_set():
                    # --- accept loop ----------------------------------------
                    try:
                        conn, addr = server.accept()
                    except socket.timeout:
                        continue
                    except OSError:
                        if self._stop_event.is_set():
                            break
                        raise

                    with self._socket_lock:
                        self._client_socket = conn

                    with self._status_lock:
                        self._is_connected = True
                        self._total_connections += 1
                    print(
                        f"[PingStation] Client connected from {addr[0]}:{addr[1]}"
                    )

                    try:
                        self._handle_connection(conn, addr)
                    finally:
                        with self._socket_lock:
                            self._client_socket = None
                        with self._status_lock:
                            self._is_connected = False
                        print(f"[PingStation] Client {addr[0]}:{addr[1]} disconnected")

            except OSError as exc:
                if self._stop_event.is_set():
                    break
                with self._status_lock:
                    self._bind_error = str(exc)
                    self._is_connected = False
                print(
                    f"[PingStation] Socket error on port {self.port}: {exc} — "
                    f"retrying in 3 s …"
                )
                time.sleep(3.0)
            finally:
                with self._socket_lock:
                    self._server_socket = None
                if server is not None:
                    try:
                        server.close()
                    except OSError:
                        pass

    # ------------------------------------------------------------------
    def _handle_connection(self, conn, addr):
        """Process all bytes from a single connected PingStation client."""
        conn.settimeout(1.0)
        buffer_bytes = b""

        with conn:
            while not self._stop_event.is_set():
                try:
                    chunk = conn.recv(16384)
                except socket.timeout:
                    continue
                except OSError:
                    break

                if not chunk:
                    break

                if self.raw_logger is not None:
                    try:
                        self.raw_logger.log_chunk(chunk, addr[0], addr[1])
                    except OSError:
                        pass

                buffer_bytes += chunk
                frames, buffer_bytes = self.decoder.split_frames(buffer_bytes)
                if not frames:
                    continue

                with self._lock:
                    for frame in frames:
                        record = self.decoder.decode_frame(frame, source_ip=addr[0])
                        if record is not None:
                            self._pending_records.append(record)
                            if self.raw_logger is not None:
                                try:
                                    self.raw_logger.log_decoded(record)
                                except OSError:
                                    pass

                with self._status_lock:
                    self._last_data_time = time.time()
                    self._total_frames_decoded += len(frames)

    # ------------------------------------------------------------------
    def get_snapshot(self):
        with self._lock:
            return self.decoder.get_state_snapshot(stale_seconds=self.stale_seconds)

    def drain_records(self):
        """Return each newly decoded record once for persistent storage."""
        with self._lock:
            rows = self._pending_records
            self._pending_records = []
            return rows
