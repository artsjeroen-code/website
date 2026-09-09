#!/usr/bin/env python3
import json
import os
import sqlite3
from datetime import datetime, timezone
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.error import HTTPError, URLError
from urllib.parse import parse_qs, urlparse
from urllib.request import Request, urlopen

HOST = os.environ.get("RITTEN_HOST", "127.0.0.1")
PORT = int(os.environ.get("RITTEN_PORT", "8765"))
DB_PATH = os.environ.get("RITTEN_DB", "/var/lib/rittenregistratie/ritten.db")
OSRM_BASE = os.environ.get("RITTEN_OSRM", "https://router.project-osrm.org")
USER_AGENT = "artsjeroen-rittenregistratie/0.5 (+https://artsjeroen.ddns.net/tools/rittenregistratie/)"

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

CREATE TABLE IF NOT EXISTS vehicles (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    make TEXT NOT NULL,
    model TEXT NOT NULL,
    plate TEXT NOT NULL UNIQUE,
    use_from TEXT NOT NULL,
    use_to TEXT,
    created_at TEXT NOT NULL,
    updated_at TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS ride_audit (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    ride_id INTEGER NOT NULL,
    corrected_at TEXT NOT NULL,
    reason TEXT NOT NULL,
    old_json TEXT NOT NULL,
    new_json TEXT NOT NULL,
    FOREIGN KEY (ride_id) REFERENCES rides(id)
);
CREATE INDEX IF NOT EXISTS idx_ride_audit_ride ON ride_audit(ride_id, id);
"""


def table_columns(db, table):
    return {row[1] for row in db.execute(f"PRAGMA table_info({table})").fetchall()}


def migrate_schema(db):
    columns = table_columns(db, "rides")
    if "vehicle_id" not in columns:
        db.execute("ALTER TABLE rides ADD COLUMN vehicle_id INTEGER")

    legacy = db.execute("SELECT * FROM vehicle WHERE id = 1").fetchone()
    vehicle_count = db.execute("SELECT COUNT(*) FROM vehicles").fetchone()[0]
    if legacy is not None and vehicle_count == 0:
        created_at = legacy["updated_at"] or datetime.now(timezone.utc).isoformat()
        cursor = db.execute(
            """
            INSERT INTO vehicles (make, model, plate, use_from, use_to, created_at, updated_at)
            VALUES (?, ?, ?, ?, ?, ?, ?)
            """,
            (
                legacy["make"], legacy["model"], legacy["plate"], legacy["use_from"],
                legacy["use_to"], created_at, created_at,
            ),
        )
        db.execute("UPDATE rides SET vehicle_id = ? WHERE vehicle_id IS NULL", (cursor.lastrowid,))
    elif vehicle_count == 1:
        only_vehicle = db.execute("SELECT id FROM vehicles LIMIT 1").fetchone()
        db.execute("UPDATE rides SET vehicle_id = ? WHERE vehicle_id IS NULL", (only_vehicle["id"],))

    db.execute("CREATE INDEX IF NOT EXISTS idx_rides_vehicle_id ON rides(vehicle_id, id)")
    db.commit()


def db_connect():
    os.makedirs(os.path.dirname(DB_PATH), exist_ok=True)
    connection = sqlite3.connect(DB_PATH)
    connection.row_factory = sqlite3.Row
    connection.execute("PRAGMA journal_mode=WAL")
    connection.execute("PRAGMA foreign_keys=ON")
    connection.executescript(SCHEMA)
    migrate_schema(connection)
    return connection


def row_to_dict(row):
    keys = row.keys()
    return {
        "id": row["id"],
        "vehicleId": row["vehicle_id"] if "vehicle_id" in keys else None,
        "vehiclePlate": row["vehicle_plate"] if "vehicle_plate" in keys else None,
        "date": row["ride_date"],
        "type": row["ride_type"],
        "startOdometer": row["start_odometer"],
        "endOdometer": row["end_odometer"],
        "distance": row["distance"],
        "departureAddress": row["departure_address"],
        "arrivalAddress": row["arrival_address"],
        "departureCoords": None if row["departure_lat"] is None else {
            "lat": row["departure_lat"], "lon": row["departure_lon"]
        },
        "arrivalCoords": None if row["arrival_lat"] is None else {
            "lat": row["arrival_lat"], "lon": row["arrival_lon"]
        },
        "notes": row["notes"],
        "createdAt": row["created_at"],
    }


def vehicle_to_dict(row):
    if row is None:
        return None
    return {
        "id": row["id"],
        "make": row["make"],
        "model": row["model"],
        "plate": row["plate"],
        "useFrom": row["use_from"],
        "useTo": row["use_to"],
        "createdAt": row["created_at"],
        "updatedAt": row["updated_at"],
    }


def audit_to_dict(row):
    return {
        "id": row["id"],
        "rideId": row["ride_id"],
        "correctedAt": row["corrected_at"],
        "reason": row["reason"],
        "old": json.loads(row["old_json"]),
        "new": json.loads(row["new_json"]),
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
    required = ["date", "type", "startOdometer", "endOdometer", "departureAddress", "arrivalAddress"]
    missing = [key for key in required if payload.get(key) in (None, "")]
    if missing:
        raise ValueError("Ontbrekende velden: " + ", ".join(missing))
    if payload["type"] not in ("business", "private"):
        raise ValueError("Ongeldig rittype")
    try:
        start = int(payload["startOdometer"])
        end = int(payload["endOdometer"])
    except (TypeError, ValueError):
        raise ValueError("Kilometerstanden moeten gehele getallen zijn")
    if start < 0 or end < start:
        raise ValueError("Ongeldige kilometerstanden")
    validate_date(payload["date"], "Datum")
    departure_address = str(payload["departureAddress"]).strip()
    arrival_address = str(payload["arrivalAddress"]).strip()
    notes = str(payload.get("notes") or "").strip()
    if not departure_address or not arrival_address:
        raise ValueError("Vertrek- en aankomstadres zijn verplicht")
    if len(departure_address) > 300 or len(arrival_address) > 300 or len(notes) > 500:
        raise ValueError("Een tekstveld is te lang")
    return start, end


def validate_vehicle_id(payload):
    try:
        vehicle_id = int(payload.get("vehicleId"))
    except (TypeError, ValueError):
        raise ValueError("Kies een geldig voertuig")
    if vehicle_id <= 0:
        raise ValueError("Kies een geldig voertuig")
    return vehicle_id


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
    url = f"{OSRM_BASE}/route/v1/driving/{dep_lon},{dep_lat};{arr_lon},{arr_lat}?overview=false&alternatives=false&steps=false"
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


def correction_coords(payload, key, existing_lat, existing_lon):
    coords = payload.get(key)
    if coords is None:
        return existing_lat, existing_lon
    if not isinstance(coords, dict):
        raise ValueError("Ongeldige GPS-coördinaten")
    return (
        validate_coord(coords.get("lat"), -90, 90, "latitude"),
        validate_coord(coords.get("lon"), -180, 180, "longitude"),
    )


def ride_select_sql(where=""):
    return f"""
        SELECT r.*, v.plate AS vehicle_plate
        FROM rides r
        LEFT JOIN vehicles v ON v.id = r.vehicle_id
        {where}
    """


class Handler(BaseHTTPRequestHandler):
    server_version = "RittenregistratieAPI/0.5"

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
        try:
            parsed = urlparse(self.path)
            path = parsed.path
            if path == "/api/health":
                self.send_json(200, {"ok": True, "database": DB_PATH})
                return

            if path == "/api/rides":
                with db_connect() as db:
                    rows = db.execute(ride_select_sql() + " ORDER BY r.ride_date ASC, r.id ASC").fetchall()
                self.send_json(200, {"rides": [row_to_dict(row) for row in rows]})
                return

            if path == "/api/vehicles":
                with db_connect() as db:
                    rows = db.execute("SELECT * FROM vehicles ORDER BY use_from ASC, id ASC").fetchall()
                self.send_json(200, {"vehicles": [vehicle_to_dict(row) for row in rows]})
                return

            # Tijdelijke compatibiliteit voor oudere frontend tijdens deployment.
            if path == "/api/vehicle":
                with db_connect() as db:
                    row = db.execute("SELECT * FROM vehicles ORDER BY id DESC LIMIT 1").fetchone()
                self.send_json(200, {"vehicle": vehicle_to_dict(row)})
                return

            if path == "/api/audit":
                ride_id = parse_qs(parsed.query).get("rideId", [None])[0]
                with db_connect() as db:
                    if ride_id:
                        try:
                            ride_id_int = int(ride_id)
                        except ValueError:
                            raise ValueError("Ongeldig ritnummer")
                        rows = db.execute(
                            "SELECT * FROM ride_audit WHERE ride_id = ? ORDER BY id DESC", (ride_id_int,)
                        ).fetchall()
                    else:
                        rows = db.execute("SELECT * FROM ride_audit ORDER BY id DESC LIMIT 200").fetchall()
                self.send_json(200, {"audit": [audit_to_dict(row) for row in rows]})
                return

            self.send_json(404, {"error": "Niet gevonden"})
        except ValueError as error:
            self.send_json(400, {"error": str(error)})
        except sqlite3.Error:
            self.send_json(500, {"error": "Databasefout"})

    def do_POST(self):
        path = urlparse(self.path).path

        if path == "/api/route":
            try:
                self.send_json(200, {"route": fetch_route(self.read_json())})
            except ValueError as error:
                self.send_json(400, {"error": str(error)})
            except RuntimeError as error:
                self.send_json(503, {"error": str(error)})
            return

        if path == "/api/vehicles":
            try:
                payload = self.read_json()
                make, model, plate, use_from, use_to = validate_vehicle(payload)
                now = datetime.now(timezone.utc).isoformat()
                with db_connect() as db:
                    existing = db.execute("SELECT id FROM vehicles WHERE UPPER(plate) = UPPER(?)", (plate,)).fetchone()
                    if existing:
                        raise ValueError("Dit kenteken bestaat al")
                    cursor = db.execute(
                        """
                        INSERT INTO vehicles (make, model, plate, use_from, use_to, created_at, updated_at)
                        VALUES (?, ?, ?, ?, ?, ?, ?)
                        """,
                        (make, model, plate, use_from, use_to, now, now),
                    )
                    row = db.execute("SELECT * FROM vehicles WHERE id = ?", (cursor.lastrowid,)).fetchone()
                    db.commit()
                self.send_json(201, {"vehicle": vehicle_to_dict(row)})
            except ValueError as error:
                self.send_json(400, {"error": str(error)})
            except sqlite3.IntegrityError:
                self.send_json(400, {"error": "Dit kenteken bestaat al"})
            except sqlite3.Error:
                self.send_json(500, {"error": "Databasefout"})
            return

        if path != "/api/rides":
            self.send_json(404, {"error": "Niet gevonden"})
            return

        try:
            payload = self.read_json()
            start, end = validate_ride(payload)
            vehicle_id = validate_vehicle_id(payload)
            departure = payload.get("departureCoords") or {}
            arrival = payload.get("arrivalCoords") or {}
            created_at = datetime.now(timezone.utc).isoformat()

            with db_connect() as db:
                vehicle = db.execute("SELECT * FROM vehicles WHERE id = ?", (vehicle_id,)).fetchone()
                if vehicle is None:
                    raise ValueError("Geselecteerd voertuig bestaat niet")
                if payload["date"] < vehicle["use_from"]:
                    raise ValueError(f"Ritdatum ligt vóór de gebruiksperiode van {vehicle['plate']}")
                if vehicle["use_to"] and payload["date"] > vehicle["use_to"]:
                    raise ValueError(f"Ritdatum ligt ná de gebruiksperiode van {vehicle['plate']}")

                previous = db.execute(
                    "SELECT end_odometer FROM rides WHERE vehicle_id = ? ORDER BY id DESC LIMIT 1",
                    (vehicle_id,),
                ).fetchone()
                if previous is not None and start != previous["end_odometer"]:
                    raise ValueError(
                        f"Niet sluitend voor {vehicle['plate']}: vorige eindstand is {previous['end_odometer']} km"
                    )

                cursor = db.execute(
                    """
                    INSERT INTO rides (
                        vehicle_id, ride_date, ride_type, start_odometer, end_odometer, distance,
                        departure_address, arrival_address, departure_lat, departure_lon,
                        arrival_lat, arrival_lon, notes, created_at
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        vehicle_id, payload["date"], payload["type"], start, end, end - start,
                        str(payload["departureAddress"]).strip(), str(payload["arrivalAddress"]).strip(),
                        departure.get("lat"), departure.get("lon"), arrival.get("lat"), arrival.get("lon"),
                        str(payload.get("notes") or "").strip(), created_at,
                    ),
                )
                row = db.execute(
                    ride_select_sql("WHERE r.id = ?"), (cursor.lastrowid,)
                ).fetchone()
                db.commit()
            self.send_json(201, {"ride": row_to_dict(row)})
        except ValueError as error:
            self.send_json(400, {"error": str(error)})
        except sqlite3.Error:
            self.send_json(500, {"error": "Databasefout"})

    def do_PUT(self):
        path = urlparse(self.path).path
        parts = path.strip("/").split("/")
        if len(parts) != 3 or parts[0] != "api" or parts[1] != "vehicles":
            self.send_json(404, {"error": "Niet gevonden"})
            return
        try:
            vehicle_id = int(parts[2])
            payload = self.read_json()
            make, model, plate, use_from, use_to = validate_vehicle(payload)
            now = datetime.now(timezone.utc).isoformat()
            with db_connect() as db:
                current = db.execute("SELECT id FROM vehicles WHERE id = ?", (vehicle_id,)).fetchone()
                if current is None:
                    raise ValueError("Voertuig niet gevonden")
                duplicate = db.execute(
                    "SELECT id FROM vehicles WHERE UPPER(plate)=UPPER(?) AND id<>?", (plate, vehicle_id)
                ).fetchone()
                if duplicate:
                    raise ValueError("Dit kenteken bestaat al")
                db.execute(
                    "UPDATE vehicles SET make=?, model=?, plate=?, use_from=?, use_to=?, updated_at=? WHERE id=?",
                    (make, model, plate, use_from, use_to, now, vehicle_id),
                )
                row = db.execute("SELECT * FROM vehicles WHERE id = ?", (vehicle_id,)).fetchone()
                db.commit()
            self.send_json(200, {"vehicle": vehicle_to_dict(row)})
        except (TypeError, ValueError) as error:
            self.send_json(400, {"error": str(error)})
        except sqlite3.Error:
            self.send_json(500, {"error": "Databasefout"})

    def do_PATCH(self):
        path = urlparse(self.path).path
        parts = path.strip("/").split("/")
        if len(parts) != 3 or parts[0] != "api" or parts[1] != "rides":
            self.send_json(404, {"error": "Niet gevonden"})
            return
        try:
            ride_id = int(parts[2])
            payload = self.read_json()
            reason = str(payload.get("reason") or "").strip()
            if len(reason) < 5:
                raise ValueError("Geef een duidelijke correctiereden van minimaal 5 tekens")
            if len(reason) > 500:
                raise ValueError("Correctiereden is te lang")
            start, end = validate_ride(payload)

            with db_connect() as db:
                current = db.execute(ride_select_sql("WHERE r.id = ?"), (ride_id,)).fetchone()
                if current is None:
                    raise ValueError("Rit niet gevonden")
                vehicle_id = current["vehicle_id"]
                old_snapshot = row_to_dict(current)

                prev = db.execute(
                    "SELECT end_odometer FROM rides WHERE vehicle_id = ? AND id < ? ORDER BY id DESC LIMIT 1",
                    (vehicle_id, ride_id),
                ).fetchone()
                nxt = db.execute(
                    "SELECT start_odometer FROM rides WHERE vehicle_id = ? AND id > ? ORDER BY id ASC LIMIT 1",
                    (vehicle_id, ride_id),
                ).fetchone()
                plate = current["vehicle_plate"] or "dit voertuig"
                if prev is not None and start != prev["end_odometer"]:
                    raise ValueError(
                        f"Correctie verbreekt de kilometerketen van {plate}: vorige rit eindigt op {prev['end_odometer']} km"
                    )
                if nxt is not None and end != nxt["start_odometer"]:
                    raise ValueError(
                        f"Correctie verbreekt de kilometerketen van {plate}: volgende rit begint op {nxt['start_odometer']} km"
                    )

                dep_lat, dep_lon = correction_coords(
                    payload, "departureCoords", current["departure_lat"], current["departure_lon"]
                )
                arr_lat, arr_lon = correction_coords(
                    payload, "arrivalCoords", current["arrival_lat"], current["arrival_lon"]
                )
                db.execute(
                    """
                    UPDATE rides SET ride_date=?, ride_type=?, start_odometer=?, end_odometer=?, distance=?,
                        departure_address=?, arrival_address=?, departure_lat=?, departure_lon=?,
                        arrival_lat=?, arrival_lon=?, notes=? WHERE id=?
                    """,
                    (
                        payload["date"], payload["type"], start, end, end-start,
                        str(payload["departureAddress"]).strip(), str(payload["arrivalAddress"]).strip(),
                        dep_lat, dep_lon, arr_lat, arr_lon, str(payload.get("notes") or "").strip(), ride_id,
                    ),
                )
                updated = db.execute(ride_select_sql("WHERE r.id = ?"), (ride_id,)).fetchone()
                new_snapshot = row_to_dict(updated)
                corrected_at = datetime.now(timezone.utc).isoformat()
                db.execute(
                    "INSERT INTO ride_audit (ride_id, corrected_at, reason, old_json, new_json) VALUES (?, ?, ?, ?, ?)",
                    (
                        ride_id, corrected_at, reason,
                        json.dumps(old_snapshot, ensure_ascii=False),
                        json.dumps(new_snapshot, ensure_ascii=False),
                    ),
                )
                db.commit()
            self.send_json(200, {"ride": new_snapshot, "auditRecorded": True})
        except (TypeError, ValueError) as error:
            self.send_json(400, {"error": str(error)})
        except sqlite3.Error:
            self.send_json(500, {"error": "Databasefout"})


if __name__ == "__main__":
    with db_connect():
        pass
    print(f"Rittenregistratie API luistert op http://{HOST}:{PORT}")
    print(f"SQLite: {DB_PATH}")
    ThreadingHTTPServer((HOST, PORT), Handler).serve_forever()
