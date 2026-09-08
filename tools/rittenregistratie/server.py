#!/usr/bin/env python3
import json
import os
import sqlite3
from datetime import datetime, timezone
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.error import HTTPError, URLError
from urllib.parse import urlparse
from urllib.request import Request, urlopen

HOST = os.environ.get("RITTEN_HOST", "127.0.0.1")
PORT = int(os.environ.get("RITTEN_PORT", "8765"))
DB_PATH = os.environ.get("RITTEN_DB", "/var/lib/rittenregistratie/ritten.db")
OSRM_BASE = os.environ.get("RITTEN_OSRM", "https://router.project-osrm.org")
USER_AGENT = "artsjeroen-rittenregistratie/0.3 (+https://artsjeroen.ddns.net/tools/rittenregistratie/)"

SCHEMA = """
CREATE TABLE IF NOT EXISTS rides (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    ride_date TEXT NOT NULL,
    ride_type TEXT NOT NULL CHECK (ride_type IN ('business', 'private')),
    start_odometer INTEGER NOT NULL CHECK (start_odometer >= 0),
    end_odometer INTEGER NOT NULL CHECK (end_odometer >= start_odometer),
    distance INTEGER NOT NULL CHECK (distance = end_odometer - start_odometer),
    departure_address TEXT NOT NULL,
    arrival_address TEXT NOT NULL,
    departure_lat REAL,
    departure_lon REAL,
    arrival_lat REAL,
    arrival_lon REAL,
    notes TEXT NOT NULL DEFAULT '',
    created_at TEXT NOT NULL
);
CREATE INDEX IF NOT EXISTS idx_rides_date ON rides(ride_date, id);

CREATE TABLE IF NOT EXISTS vehicle (
    id INTEGER PRIMARY KEY CHECK (id = 1),
    make TEXT NOT NULL,
    model TEXT NOT NULL,
    plate TEXT NOT NULL,
    use_from TEXT NOT NULL,
    use_to TEXT,
    updated_at TEXT NOT NULL
);
"""


def db_connect():
    os.makedirs(os.path.dirname(DB_PATH), exist_ok=True)
    connection = sqlite3.connect(DB_PATH)
    connection.row_factory = sqlite3.Row
    connection.execute("PRAGMA journal_mode=WAL")
    connection.execute("PRAGMA foreign_keys=ON")
    connection.executescript(SCHEMA)
    return connection


def row_to_dict(row):
    return {
        "id": row["id"],
        "date": row["ride_date"],
        "type": row["ride_type"],
        "startOdometer": row["start_odometer"],
        "endOdometer": row["end_odometer"],
        "distance": row["distance"],
        "departureAddress": row["departure_address"],
        "arrivalAddress": row["arrival_address"],
        "departureCoords": None if row["departure_lat"] is None else {
            "lat": row["departure_lat"],
            "lon": row["departure_lon"],
        },
        "arrivalCoords": None if row["arrival_lat"] is None else {
            "lat": row["arrival_lat"],
            "lon": row["arrival_lon"],
        },
        "notes": row["notes"],
        "createdAt": row["created_at"],
    }


def vehicle_to_dict(row):
    if row is None:
        return None
    return {
        "make": row["make"],
        "model": row["model"],
        "plate": row["plate"],
        "useFrom": row["use_from"],
        "useTo": row["use_to"],
        "updatedAt": row["updated_at"],
    }


def validate_date(value, label, required=True):
    if value in (None, ""):
        if required:
            raise ValueError(f"{label} is verplicht")
        return None
    try:
        datetime.strptime(value, "%Y-%m-%d")
    except (TypeError, ValueError):
        raise ValueError(f"{label} moet YYYY-MM-DD zijn")
    return value


def validate_vehicle(payload):
    make = str(payload.get("make") or "").strip()
    model = str(payload.get("model") or "").strip()
    plate = str(payload.get("plate") or "").strip().upper()
    use_from = validate_date(payload.get("useFrom"), "Begindatum gebruik")
    use_to = validate_date(payload.get("useTo"), "Einddatum gebruik", required=False)

    if not make:
        raise ValueError("Merk is verplicht")
    if not model:
        raise ValueError("Type/model is verplicht")
    if not plate:
        raise ValueError("Kenteken is verplicht")
    if len(make) > 80 or len(model) > 120 or len(plate) > 20:
        raise ValueError("Voertuiggegevens zijn te lang")
    if use_to and use_to < use_from:
        raise ValueError("Einddatum gebruik kan niet vóór de begindatum liggen")

    return make, model, plate, use_from, use_to


def validate_ride(payload):
    required = [
        "date", "type", "startOdometer", "endOdometer",
        "departureAddress", "arrivalAddress"
    ]
    missing = [key for key in required if payload.get(key) in (None, "")]
    if missing:
        raise ValueError("Ontbrekende velden: " + ", ".join(missing))

    ride_type = payload["type"]
    if ride_type not in ("business", "private"):
        raise ValueError("Ongeldig rittype")

    try:
        start = int(payload["startOdometer"])
        end = int(payload["endOdometer"])
    except (TypeError, ValueError):
        raise ValueError("Kilometerstanden moeten gehele getallen zijn")

    if start < 0 or end < start:
        raise ValueError("Ongeldige kilometerstanden")

    validate_date(payload["date"], "Datum")
    return start, end


def validate_coord(value, minimum, maximum, label):
    try:
        number = float(value)
    except (TypeError, ValueError):
        raise ValueError(f"Ongeldige {label}")
    if number < minimum or number > maximum:
        raise ValueError(f"Ongeldige {label}")
    return number


def fetch_route(payload):
    departure = payload.get("departure") or {}
    arrival = payload.get("arrival") or {}
    dep_lat = validate_coord(departure.get("lat"), -90, 90, "vertrek-latitude")
    dep_lon = validate_coord(departure.get("lon"), -180, 180, "vertrek-longitude")
    arr_lat = validate_coord(arrival.get("lat"), -90, 90, "aankomst-latitude")
    arr_lon = validate_coord(arrival.get("lon"), -180, 180, "aankomst-longitude")

    url = (
        f"{OSRM_BASE}/route/v1/driving/"
        f"{dep_lon},{dep_lat};{arr_lon},{arr_lat}"
        "?overview=false&alternatives=false&steps=false"
    )
    request = Request(url, headers={
        "Accept": "application/json",
        "User-Agent": USER_AGENT,
        "Referer": "https://artsjeroen.ddns.net/tools/rittenregistratie/",
    })

    try:
        with urlopen(request, timeout=10) as response:
            result = json.loads(response.read().decode("utf-8"))
    except (HTTPError, URLError, TimeoutError, json.JSONDecodeError) as error:
        raise RuntimeError("Route-service tijdelijk niet beschikbaar") from error

    if result.get("code") != "Ok" or not result.get("routes"):
        raise RuntimeError("Geen autoroute gevonden")

    route = result["routes"][0]
    return {
        "distanceKm": round(float(route["distance"]) / 1000, 1),
        "durationMinutes": round(float(route["duration"]) / 60),
        "provider": "OSRM / OpenStreetMap",
    }


class Handler(BaseHTTPRequestHandler):
    server_version = "RittenregistratieAPI/0.3"

    def log_message(self, fmt, *args):
        print(f"{self.address_string()} - {fmt % args}")

    def send_json(self, status, payload):
        body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(body)

    def read_json(self):
        length = int(self.headers.get("Content-Length", "0"))
        if length <= 0 or length > 65536:
            raise ValueError("Ongeldige requestgrootte")
        raw = self.rfile.read(length)
        try:
            return json.loads(raw.decode("utf-8"))
        except (UnicodeDecodeError, json.JSONDecodeError):
            raise ValueError("Ongeldige JSON")

    def do_GET(self):
        path = urlparse(self.path).path
        if path == "/api/health":
            self.send_json(200, {"ok": True, "database": DB_PATH})
            return

        if path == "/api/rides":
            with db_connect() as db:
                rows = db.execute(
                    "SELECT * FROM rides ORDER BY ride_date ASC, id ASC"
                ).fetchall()
            self.send_json(200, {"rides": [row_to_dict(row) for row in rows]})
            return

        if path == "/api/vehicle":
            with db_connect() as db:
                row = db.execute("SELECT * FROM vehicle WHERE id = 1").fetchone()
            self.send_json(200, {"vehicle": vehicle_to_dict(row)})
            return

        self.send_json(404, {"error": "Niet gevonden"})

    def do_PUT(self):
        path = urlparse(self.path).path
        if path != "/api/vehicle":
            self.send_json(404, {"error": "Niet gevonden"})
            return

        try:
            payload = self.read_json()
            make, model, plate, use_from, use_to = validate_vehicle(payload)
            updated_at = datetime.now(timezone.utc).isoformat()
            with db_connect() as db:
                db.execute(
                    """
                    INSERT INTO vehicle (id, make, model, plate, use_from, use_to, updated_at)
                    VALUES (1, ?, ?, ?, ?, ?, ?)
                    ON CONFLICT(id) DO UPDATE SET
                        make = excluded.make,
                        model = excluded.model,
                        plate = excluded.plate,
                        use_from = excluded.use_from,
                        use_to = excluded.use_to,
                        updated_at = excluded.updated_at
                    """,
                    (make, model, plate, use_from, use_to, updated_at),
                )
                row = db.execute("SELECT * FROM vehicle WHERE id = 1").fetchone()
                db.commit()
            self.send_json(200, {"vehicle": vehicle_to_dict(row)})
        except ValueError as error:
            self.send_json(400, {"error": str(error)})
        except sqlite3.Error:
            self.send_json(500, {"error": "Databasefout"})

    def do_POST(self):
        path = urlparse(self.path).path

        if path == "/api/route":
            try:
                payload = self.read_json()
                self.send_json(200, {"route": fetch_route(payload)})
            except ValueError as error:
                self.send_json(400, {"error": str(error)})
            except RuntimeError as error:
                self.send_json(503, {"error": str(error)})
            return

        if path != "/api/rides":
            self.send_json(404, {"error": "Niet gevonden"})
            return

        try:
            payload = self.read_json()
            start, end = validate_ride(payload)
            departure = payload.get("departureCoords") or {}
            arrival = payload.get("arrivalCoords") or {}
            created_at = datetime.now(timezone.utc).isoformat()

            with db_connect() as db:
                vehicle = db.execute("SELECT id FROM vehicle WHERE id = 1").fetchone()
                if vehicle is None:
                    raise ValueError("Sla eerst de voertuiggegevens op")

                previous = db.execute(
                    "SELECT end_odometer FROM rides ORDER BY id DESC LIMIT 1"
                ).fetchone()
                if previous is not None and start != previous["end_odometer"]:
                    raise ValueError(
                        f"Niet sluitend: vorige eindstand is {previous['end_odometer']} km"
                    )

                cursor = db.execute(
                    """
                    INSERT INTO rides (
                        ride_date, ride_type, start_odometer, end_odometer, distance,
                        departure_address, arrival_address,
                        departure_lat, departure_lon, arrival_lat, arrival_lon,
                        notes, created_at
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        payload["date"], payload["type"], start, end, end - start,
                        str(payload["departureAddress"]).strip(),
                        str(payload["arrivalAddress"]).strip(),
                        departure.get("lat"), departure.get("lon"),
                        arrival.get("lat"), arrival.get("lon"),
                        str(payload.get("notes") or "").strip(), created_at,
                    ),
                )
                row = db.execute(
                    "SELECT * FROM rides WHERE id = ?", (cursor.lastrowid,)
                ).fetchone()
                db.commit()

            self.send_json(201, {"ride": row_to_dict(row)})
        except ValueError as error:
            self.send_json(400, {"error": str(error)})
        except sqlite3.Error:
            self.send_json(500, {"error": "Databasefout"})


if __name__ == "__main__":
    with db_connect():
        pass
    print(f"Rittenregistratie API luistert op http://{HOST}:{PORT}")
    print(f"SQLite: {DB_PATH}")
    ThreadingHTTPServer((HOST, PORT), Handler).serve_forever()
