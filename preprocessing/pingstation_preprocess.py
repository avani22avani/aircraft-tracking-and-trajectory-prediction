from datetime import datetime, timezone


class PingstationPreprocessor:
    def _first_present(self, raw_data, *keys, default=None):
        for key in keys:
            if key in raw_data and raw_data[key] not in (None, ""):
                return raw_data[key]
        return default

    def _to_float(self, value, default=0.0):
        try:
            return float(value)
        except (TypeError, ValueError):
            return default

    def clean_altitude_mm(self, altitude):
        """Normalizes altitude to millimeters."""
        if altitude is None:
            return 0
 
        if isinstance(altitude, str):
            text = altitude.strip().lower().replace(",", "")
            if not text:
                return 0

            if "ft" in text:
                feet = self._to_float(text.replace("ft", "").strip())
                meters = feet * 0.3048
                return int(meters * 1000)

            if text.endswith("m"):
                meters = self._to_float(text.replace("m", "").strip())
                return int(meters * 1000)

            numeric = self._to_float(text)
            return int(numeric if numeric > 100000 else numeric * 1000)

        numeric = self._to_float(altitude)

        # Pingstation-style values are often already in mm and usually larger than 100k.
        if numeric > 100000:
            return int(numeric)

        # Otherwise treat as meters.
        return int(numeric * 1000)

    def clean_heading_de2(self, heading):
        """Normalizes heading to degrees*100."""
        if heading is None:
            return 0

        if isinstance(heading, str):
            heading = heading.strip().replace("deg", "").replace("°", "")

        value = self._to_float(heading)

        if value <= 360:
            return int(value * 100)

        return int(value)

    def clean_speed_cms(self, speed):
        """Normalizes horizontal speed to centimeters/second."""
        if speed is None:
            return 0

        if isinstance(speed, str):
            text = speed.strip().lower().replace(",", "")
            if "kt" in text:
                knots = self._to_float(text.replace("kt", "").strip())
                mps = knots * 0.514444
                return int(mps * 100)
            speed = text

        value = self._to_float(speed)

        # If the value looks like m/s, convert to cm/s.
        if value < 1000:
            return int(value * 100)

        return int(value)

    def _normalize_timestamp(self, timestamp):
        if timestamp in (None, ""):
            return datetime.now(timezone.utc).isoformat()

        text = str(timestamp).strip()
        if not text:
            return datetime.now(timezone.utc).isoformat()

        try:
            return datetime.fromisoformat(text.replace("Z", "+00:00")).isoformat()
        except ValueError:
            return datetime.now(timezone.utc).isoformat()

    def extract_time_features(self, timestamp):
        try:
            dt = datetime.fromisoformat(str(timestamp).replace("Z", "+00:00"))
            return {"hour": dt.hour, "minute": dt.minute}
        except (TypeError, ValueError):
            now = datetime.now()
            return {"hour": now.hour, "minute": now.minute}

    def preprocess(self, raw_data):
        if not isinstance(raw_data, dict):
            return {
                "icaoAddress": "UNKNOWN",
                "callsign": "UNKNOWN",
                "latDD": 0.0,
                "lonDD": 0.0,
                "altitudeMM": 0,
                "headingDE2": 0,
                "horVelocityCMS": 0,
                "timestamp": datetime.now(timezone.utc).isoformat(),
                "hour": datetime.now().hour,
                "minute": datetime.now().minute,
            }

        lat = self._to_float(
            self._first_present(raw_data, "latDD", "lat", "latitude", default=0.0)
        )
        lon = self._to_float(
            self._first_present(raw_data, "lonDD", "lon", "longitude", default=0.0)
        )

        if "altitudeMM" in raw_data and raw_data.get("altitudeMM") not in (None, ""):
            altitude = int(self._to_float(raw_data.get("altitudeMM"), default=0))
        else:
            altitude = self.clean_altitude_mm(
                self._first_present(raw_data, "altitudeMM", "altitude", "alt", default=0)
            )

        if "headingDE2" in raw_data and raw_data.get("headingDE2") not in (None, ""):
            heading = int(self._to_float(raw_data.get("headingDE2"), default=0))
        else:
            heading = self.clean_heading_de2(
                self._first_present(raw_data, "headingDE2", "heading", "track", default=0)
            )

        if "horVelocityCMS" in raw_data and raw_data.get("horVelocityCMS") not in (None, ""):
            speed = int(self._to_float(raw_data.get("horVelocityCMS"), default=0))
        else:
            speed = self.clean_speed_cms(
                self._first_present(
                    raw_data,
                    "horVelocityCMS",
                    "speed",
                    "groundspeed",
                    "gs",
                    default=0,
                )
            )

        callsign = self._first_present(
            raw_data, "callsign", "flight_no", "flight", default="UNKNOWN"
        )
        callsign = str(callsign).strip().upper() or "UNKNOWN"

        icao = self._first_present(
            raw_data, "icaoAddress", "icao", "hex", default="UNKNOWN"
        )
        icao = str(icao).strip().upper() or "UNKNOWN"

        timestamp = self._normalize_timestamp(
            self._first_present(raw_data, "timestamp", "time", default=None)
        )

        return {
            "icaoAddress": icao,
            "callsign": callsign,
            "latDD": lat,
            "lonDD": lon,
            "altitudeMM": altitude,
            "headingDE2": heading,
            "horVelocityCMS": speed,
            "timestamp": timestamp,
            **self.extract_time_features(timestamp),
        }
