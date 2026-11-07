"""
Best-effort enrichment of an aircraft's ICAO24 hex code with its type,
photo, and a coarse category (used to pick a map icon: jet / prop / helicopter).

Looks the data up from free public sources (hexdb.io for type/registration,
planespotters.net for a representative photo) and caches results in-memory so
the once-a-second polling from the UI doesn't hammer those APIs. Network
failures are swallowed - callers always get a dict back, possibly with empty
fields, so the UI can fall back to generic icons/placeholders gracefully.
"""

import threading
import time

import requests

# Curated ICAO type-designator sets used to classify an aircraft into a
# coarse category for icon selection. Not exhaustive - anything not matched
# falls back to "jet" (the generic fixed-wing/airliner silhouette).
HELICOPTER_TYPE_CODES = {
    "R22", "R44", "R66", "R30", "EC20", "EC25", "EC30", "EC35", "EC45", "EC55",
    "EC75", "AS32", "AS50", "AS55", "AS65", "A109", "A139", "A169", "B06",
    "B204", "B205", "B212", "B214", "B222", "B230", "B407", "B412", "B429",
    "B430", "H500", "H53", "H60", "H64", "S76", "S92", "GAZL", "LAMA", "MI2",
    "MI8", "MI17", "MI24", "MI26", "W3SZ", "EN28", "KMAX", "H160", "H145",
    "H135", "H125", "AW09", "AW119", "AW139", "AW169", "AW189",
}

PROP_TYPE_CODES = {
    "C172", "C152", "C182", "C206", "C208", "PA28", "PA31", "PA34", "PA46",
    "BE20", "BE9L", "BE35", "BE36", "BE58", "BE99", "B350", "DH8A", "DH8B",
    "DH8C", "DH8D", "AT43", "AT45", "AT72", "AT75", "AT76", "SF34", "SW4",
    "JS31", "JS32", "J328", "EMB1", "C30J", "DHC6", "DHC2", "TBM7", "TBM8",
    "SR22", "SR20", "DA40", "DA42", "DA62", "M20P", "P28A", "P32R", "PC12",
    "PC6T", "C130", "AN26", "AN32", "Y12", "P180", "F406", "C212",
}

POSITIVE_TTL_SEC = 24 * 3600   # remember confirmed lookups for a day
NEGATIVE_TTL_SEC = 30 * 60     # retry "not found" / failed lookups sooner
REQUEST_TIMEOUT_SEC = 4

# planespotters.net rejects generic library user-agent strings; identify
# ourselves so photo lookups don't get blocked.
REQUEST_HEADERS = {
    "User-Agent": "ADSBTracker-DesktopApp/1.0 (+https://github.com/; contact: adsb-tracker@example.com)"
}

_cache = {}
_lock = threading.Lock()

_route_cache = {}
_route_lock = threading.Lock()

# Some IATA codes are duplicated internationally. Prefer the Indian airport
# for Indian domestic airline routes instead of accepting the provider's
# older South Sudan HGI record.
_COUNTRY_AIRPORT_CORRECTIONS = {
    ("IN", "HGI"): {
        "name": "Donyi Polo Airport",
        "icaoCode": "VEHO",
        "iataCode": "HGI",
        "municipality": "Itanagar",
        "countryName": "India",
        "latitude": 26.9700,
        "longitude": 93.6467,
    },
}


def _classify(icao_type_code, type_text):
    code = (icao_type_code or "").strip().upper()
    text = (type_text or "").strip().upper()
    if code in HELICOPTER_TYPE_CODES or "HELICOPTER" in text or "GYROCOPTER" in text:
        return "helicopter"
    if code in PROP_TYPE_CODES:
        return "prop"
    return "jet"


def _empty_info(icao24):
    return {
        "icaoAddress": icao24,
        "manufacturer": "",
        "type": "",
        "icaoTypeCode": "",
        "registration": "",
        "category": "jet",
        "photoUrl": None,
        "photoLink": None,
        "photoCredit": None,
    }


def _fetch_type_info(icao24, info):
    try:
        resp = requests.get(
            f"https://hexdb.io/api/v1/aircraft/{icao24}",
            headers=REQUEST_HEADERS,
            timeout=REQUEST_TIMEOUT_SEC,
        )
    except requests.RequestException:
        return False
    if not resp.ok:
        return False
    try:
        payload = resp.json()
    except ValueError:
        return False
    if not isinstance(payload, dict) or not (payload.get("Registration") or payload.get("ICAOTypeCode")):
        return False

    info["manufacturer"] = str(payload.get("Manufacturer") or "").strip()
    info["type"] = str(payload.get("Type") or "").strip()
    info["icaoTypeCode"] = str(payload.get("ICAOTypeCode") or "").strip().upper()
    info["registration"] = str(payload.get("Registration") or "").strip()
    info["category"] = _classify(info["icaoTypeCode"], info["type"])
    return True


def _fetch_photo(icao24, info):
    try:
        resp = requests.get(
            f"https://api.planespotters.net/pub/photos/hex/{icao24}",
            headers=REQUEST_HEADERS,
            timeout=REQUEST_TIMEOUT_SEC,
        )
    except requests.RequestException:
        return False
    if not resp.ok:
        return False
    try:
        payload = resp.json()
    except ValueError:
        return False

    photos = payload.get("photos") if isinstance(payload, dict) else None
    if not photos:
        return False

    photo = photos[0]
    thumb = photo.get("thumbnail_large") or photo.get("thumbnail") or {}
    src = thumb.get("src")
    if not src:
        return False

    info["photoUrl"] = src
    info["photoLink"] = photo.get("link")
    photographer = (photo.get("photographer") or "").strip()
    info["photoCredit"] = f"© {photographer} / planespotters.net" if photographer else "planespotters.net"
    return True


def _airport_summary(payload):
    if not isinstance(payload, dict):
        return None
    name = str(payload.get("name") or "").strip()
    if not name:
        return None
    return {
        "name": name,
        "icaoCode": str(payload.get("icao_code") or "").strip().upper(),
        "iataCode": str(payload.get("iata_code") or "").strip().upper(),
        "municipality": str(payload.get("municipality") or "").strip(),
        "countryName": str(payload.get("country_name") or "").strip(),
        "latitude": payload.get("latitude"),
        "longitude": payload.get("longitude"),
    }


def _correct_ambiguous_airport(airport, country_iso):
    if not airport:
        return airport
    correction = _COUNTRY_AIRPORT_CORRECTIONS.get(
        (str(country_iso or "").upper(), airport.get("iataCode", ""))
    )
    return dict(correction) if correction else airport


def lookup_flight_route(callsign):
    """
    Returns the scheduled origin/destination airports for a flight number,
    looked up from adsbdb.com's free callsign-route API (no key required).
    Cached in-memory; returns None on failure or when not found.
    """
    callsign = str(callsign or "").strip().upper()
    if not callsign or callsign == "UNKNOWN":
        return None

    now = time.time()
    with _route_lock:
        cached = _route_cache.get(callsign)
    if cached and (now - cached["fetchedAt"]) < cached["ttl"]:
        return cached["data"]

    data = None
    try:
        resp = requests.get(
            f"https://api.adsbdb.com/v0/callsign/{callsign}",
            headers=REQUEST_HEADERS,
            timeout=REQUEST_TIMEOUT_SEC,
        )
        if resp.ok:
            payload = resp.json()
            route = (payload or {}).get("response", {}).get("flightroute")
            if isinstance(route, dict):
                airline = route.get("airline") or {}
                origin = _airport_summary(route.get("origin"))
                destination = _airport_summary(route.get("destination"))
                domestic_country = str(airline.get("country_iso") or "").upper()
                if origin and str((route.get("origin") or {}).get("country_iso_name") or "").upper() == domestic_country:
                    destination = _correct_ambiguous_airport(destination, domestic_country)
                if origin or destination:
                    data = {
                        "callsign": callsign,
                        "callsignIata": str(route.get("callsign_iata") or "").strip().upper(),
                        "airlineName": str(airline.get("name") or "").strip(),
                        "origin": origin,
                        "destination": destination,
                    }
    except (requests.RequestException, ValueError):
        data = None

    with _route_lock:
        _route_cache[callsign] = {
            "data": data,
            "fetchedAt": now,
            "ttl": POSITIVE_TTL_SEC if data else NEGATIVE_TTL_SEC,
        }
    return data


def lookup_aircraft_info(icao24):
    """Returns an enrichment dict for the given ICAO24 hex (cached, best-effort)."""
    icao24 = str(icao24 or "").strip().upper()
    if not icao24 or icao24 == "UNKNOWN":
        return None

    now = time.time()
    with _lock:
        cached = _cache.get(icao24)
    if cached and (now - cached["fetchedAt"]) < cached["ttl"]:
        return cached["data"]

    info = _empty_info(icao24)
    found_type = _fetch_type_info(icao24, info)
    found_photo = _fetch_photo(icao24, info)
    found_anything = found_type or found_photo

    with _lock:
        _cache[icao24] = {
            "data": info,
            "fetchedAt": now,
            "ttl": POSITIVE_TTL_SEC if found_anything else NEGATIVE_TTL_SEC,
        }
    return info
