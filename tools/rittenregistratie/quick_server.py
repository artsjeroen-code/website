#!/usr/bin/env python3
import json
import sqlite3
from datetime import datetime, timezone
from urllib.parse import urlparse

import server as base

FULL_REGISTRATION_FROM = "2027-01-01"
PRIVATE_KM_LIMIT = 500

QUICK_SCHEMA = """
CREATE TABLE IF NOT EXISTS quick_rides (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    start_captured_at TEXT NOT NULL,
    start_lat REAL NOT NULL,
    start_lon REAL NOT NULL,
    start_accuracy REAL,
    start_address TEXT NOT NULL DEFAULT '',
    start_odometer INTEGER,
    end_captured_at TEXT,
    end_lat REAL,
    end_lon REAL,
    end_accuracy REAL,
    end_address TEXT NOT NULL DEFAULT '',
    end_odometer INTEGER,
    notes TEXT NOT NULL DEFAULT '',
    created_at TEXT NOT NULL,
    archived_at TEXT
);
CREATE INDEX IF NOT EXISTS idx_quick_rides_open ON quick_rides(archived_at, end_captured_at, id);
"""


def ensure_quick_columns(db):
    columns = {row["name"] for row in db.execute("PRAGMA table_info(quick_rides)").fetchall()}
    if "start_odometer" not in columns:
        db.execute("ALTER TABLE quick_rides ADD COLUMN start_odometer INTEGER")
    if "end_odometer" not in columns:
        db.execute("ALTER TABLE quick_rides ADD COLUMN end_odometer INTEGER")
    if "notes" not in columns:
        db.execute("ALTER TABLE quick_rides ADD COLUMN notes TEXT NOT NULL DEFAULT ''")
    db.commit()


def db_connect():
    db = base.db_connect()
    db.executescript(QUICK_SCHEMA)
    ensure_quick_columns(db)
    return db


def full_registration(date_value):
    return str(date_value or "") >= FULL_REGISTRATION_FROM


def validate_policy_ride(payload):
    start, end = base.validate_ride(payload)
    if not full_registration(payload.get("date")) and payload.get("type") != "business":
        raise ValueError("Tot en met 2026 worden alleen zakelijke ritten geregistreerd")
    return start, end


def validate_optional_odometer(value):
    if value in (None, ""):
        return None
    try:
        number = int(str(value).strip())
    except (TypeError, ValueError):
        raise ValueError("Kilometerstand moet een heel getal zijn")
    if number < 0:
        raise ValueError("Kilometerstand mag niet negatief zijn")
    return number


def validate_notes(value):
    notes = str(value or "").strip()
    if len(notes) > 500:
        raise ValueError("Omschrijving is te lang")
    return notes


def quick_to_dict(row):
    return {
        "id": row["id"],
        "startCapturedAt": row["start_captured_at"],
        "startCoords": {"lat": row["start_lat"], "lon": row["start_lon"], "accuracy": row["start_accuracy"]},
        "startAddress": row["start_address"],
        "startOdometer": row["start_odometer"],
        "endCapturedAt": row["end_captured_at"],
        "endCoords": None if row["end_lat"] is None else {"lat": row["end_lat"], "lon": row["end_lon"], "accuracy": row["end_accuracy"]},
        "endAddress": row["end_address"],
        "endOdometer": row["end_odometer"],
        "notes": row["notes"],
        "complete": row["end_captured_at"] is not None,
    }


def capture_values(payload):
    coords = payload.get("coords") or {}
    lat = base.validate_coord(coords.get("lat"), -90, 90, "latitude")
    lon = base.validate_coord(coords.get("lon"), -180, 180, "longitude")
    accuracy = coords.get("accuracy")
    if accuracy in (None, ""):
        accuracy = None
    else:
        try:
            accuracy = max(0.0, float(accuracy))
        except (TypeError, ValueError):
            raise ValueError("Ongeldige GPS-nauwkeurigheid")
    captured_at = str(payload.get("capturedAt") or "").strip()
    if len(captured_at) < 16 or len(captured_at) > 40:
        raise ValueError("Ongeldig datum/tijdstip")
    address = str(payload.get("address") or "").strip()
    if len(address) > 300:
        raise ValueError("Adres is te lang")
    odometer = validate_optional_odometer(payload.get("odometer"))
    return captured_at, lat, lon, accuracy, address, odometer


class Handler(base.Handler):
    server_version = "RittenregistratieAPI/1.0"

    def do_GET(self):
        path = urlparse(self.path).path
        if path == "/api/policy":
            self.send_json(200, {
                "fullRegistrationFrom": FULL_REGISTRATION_FROM,
                "privateKmLimit": PRIVATE_KM_LIMIT,
            })
            return
        if path != "/api/quick-rides":
            return super().do_GET()
        try:
            with db_connect() as db:
                rows = db.execute(
                    "SELECT * FROM quick_rides WHERE archived_at IS NULL ORDER BY id DESC LIMIT 100"
                ).fetchall()
            self.send_json(200, {"quickRides": [quick_to_dict(row) for row in rows]})
        except sqlite3.Error:
            self.send_json(500, {"error": "Databasefout"})

    def save_policy_ride(self):
        try:
            payload = self.read_json()
            start, end = validate_policy_ride(payload)
            vehicle_id = base.validate_vehicle_id(payload)
            departure = payload.get("departureCoords") or {}
            arrival = payload.get("arrivalCoords") or {}
            departure_time = base.validate_time(payload.get("departureTime"), "Vertrektijd")
            arrival_time = base.validate_time(payload.get("arrivalTime"), "Aankomsttijd")
            created_at = datetime.now(timezone.utc).isoformat()

            with db_connect() as db:
                vehicle = db.execute("SELECT * FROM vehicles WHERE id = ?", (vehicle_id,)).fetchone()
                if vehicle is None:
                    raise ValueError("Geselecteerd voertuig bestaat niet")
                if payload["date"] < vehicle["use_from"]:
                    raise ValueError(f"Ritdatum ligt vóór de gebruiksperiode van {vehicle['plate']}")
                if vehicle["use_to"] and payload["date"] > vehicle["use_to"]:
                    raise ValueError(f"Ritdatum ligt ná de gebruiksperiode van {vehicle['plate']}")

                if full_registration(payload["date"]):
                    previous = db.execute(
                        """
                        SELECT end_odometer, ride_date FROM rides
                        WHERE vehicle_id = ? AND ride_date >= ?
                        ORDER BY ride_date DESC, id DESC LIMIT 1
                        """,
                        (vehicle_id, FULL_REGISTRATION_FROM),
                    ).fetchone()
                    if previous is not None:
                        if payload["date"] < previous["ride_date"]:
                            raise ValueError("Voeg ritten vanaf 2027 in chronologische volgorde toe")
                        if start != previous["end_odometer"]:
                            raise ValueError(
                                f"Niet sluitend voor {vehicle['plate']}: vorige eindstand is {previous['end_odometer']} km"
                            )

                cursor = db.execute(
                    """
                    INSERT INTO rides (
                        vehicle_id, ride_date, ride_type, start_odometer, end_odometer, distance,
                        departure_address, arrival_address, departure_lat, departure_lon,
                        arrival_lat, arrival_lon, departure_time, arrival_time, notes, created_at
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        vehicle_id, payload["date"], payload["type"], start, end, end - start,
                        str(payload["departureAddress"]).strip(), str(payload["arrivalAddress"]).strip(),
                        departure.get("lat"), departure.get("lon"), arrival.get("lat"), arrival.get("lon"),
                        departure_time, arrival_time, str(payload.get("notes") or "").strip(), created_at,
                    ),
                )
                row = db.execute(base.ride_select_sql("WHERE r.id = ?"), (cursor.lastrowid,)).fetchone()
                db.commit()
            self.send_json(201, {"ride": base.row_to_dict(row)})
        except ValueError as error:
            self.send_json(400, {"error": str(error)})
        except sqlite3.Error:
            self.send_json(500, {"error": "Databasefout"})

    def do_POST(self):
        path = urlparse(self.path).path
        if path == "/api/rides":
            self.save_policy_ride()
            return

        if path not in ("/api/quick-rides/start", "/api/quick-rides/end") and not (
            path.startswith("/api/quick-rides/") and path.endswith("/archive")
        ):
            return super().do_POST()

        try:
            if path == "/api/quick-rides/start":
                payload = self.read_json()
                captured_at, lat, lon, accuracy, address, odometer = capture_values(payload)
                notes = validate_notes(payload.get("notes"))
                created_at = datetime.now(timezone.utc).isoformat()
                with db_connect() as db:
                    open_row = db.execute(
                        "SELECT id FROM quick_rides WHERE archived_at IS NULL AND end_captured_at IS NULL ORDER BY id DESC LIMIT 1"
                    ).fetchone()
                    if open_row:
                        raise ValueError("Er staat nog een beginpunt open. Registreer eerst het eindpunt.")
                    cursor = db.execute(
                        """
                        INSERT INTO quick_rides (
                            start_captured_at, start_lat, start_lon, start_accuracy, start_address,
                            start_odometer, notes, created_at
                        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                        """,
                        (captured_at, lat, lon, accuracy, address, odometer, notes, created_at),
                    )
                    row = db.execute("SELECT * FROM quick_rides WHERE id=?", (cursor.lastrowid,)).fetchone()
                    db.commit()
                self.send_json(201, {"quickRide": quick_to_dict(row)})
                return

            if path == "/api/quick-rides/end":
                payload = self.read_json()
                captured_at, lat, lon, accuracy, address, odometer = capture_values(payload)
                with db_connect() as db:
                    row = db.execute(
                        "SELECT * FROM quick_rides WHERE archived_at IS NULL AND end_captured_at IS NULL ORDER BY id DESC LIMIT 1"
                    ).fetchone()
                    if row is None:
                        raise ValueError("Er is geen open beginpunt. Registreer eerst het beginpunt.")
                    if odometer is not None and row["start_odometer"] is not None and odometer < row["start_odometer"]:
                        raise ValueError(f"Eindstand kan niet lager zijn dan beginstand {row['start_odometer']} km")
                    db.execute(
                        """
                        UPDATE quick_rides
                        SET end_captured_at=?, end_lat=?, end_lon=?, end_accuracy=?, end_address=?, end_odometer=?
                        WHERE id=?
                        """,
                        (captured_at, lat, lon, accuracy, address, odometer, row["id"]),
                    )
                    updated = db.execute("SELECT * FROM quick_rides WHERE id=?", (row["id"],)).fetchone()
                    db.commit()
                self.send_json(200, {"quickRide": quick_to_dict(updated)})
                return

            parts = path.strip("/").split("/")
            quick_id = int(parts[2])
            with db_connect() as db:
                row = db.execute("SELECT id FROM quick_rides WHERE id=? AND archived_at IS NULL", (quick_id,)).fetchone()
                if row is None:
                    raise ValueError("Snelle registratie niet gevonden")
                db.execute(
                    "UPDATE quick_rides SET archived_at=? WHERE id=?",
                    (datetime.now(timezone.utc).isoformat(), quick_id),
                )
                db.commit()
            self.send_json(200, {"ok": True})
        except (TypeError, ValueError) as error:
            self.send_json(400, {"error": str(error)})
        except sqlite3.Error:
            self.send_json(500, {"error": "Databasefout"})

    def do_PATCH(self):
        path = urlparse(self.path).path
        parts = path.strip("/").split("/")
        if len(parts) != 3 or parts[0] != "api" or parts[1] != "rides":
            return super().do_PATCH()
        try:
            ride_id = int(parts[2])
            payload = self.read_json()
            reason = str(payload.get("reason") or "").strip()
            if len(reason) < 5:
                raise ValueError("Geef een duidelijke correctiereden van minimaal 5 tekens")
            if len(reason) > 500:
                raise ValueError("Correctiereden is te lang")
            start, end = validate_policy_ride(payload)

            with db_connect() as db:
                current = db.execute(base.ride_select_sql("WHERE r.id = ?"), (ride_id,)).fetchone()
                if current is None:
                    raise ValueError("Rit niet gevonden")
                vehicle_id = current["vehicle_id"]
                old_snapshot = base.row_to_dict(current)
                plate = current["vehicle_plate"] or "dit voertuig"

                if full_registration(payload["date"]):
                    prev = db.execute(
                        """
                        SELECT end_odometer FROM rides
                        WHERE vehicle_id=? AND id<? AND ride_date>=?
                        ORDER BY id DESC LIMIT 1
                        """,
                        (vehicle_id, ride_id, FULL_REGISTRATION_FROM),
                    ).fetchone()
                    nxt = db.execute(
                        """
                        SELECT start_odometer FROM rides
                        WHERE vehicle_id=? AND id>? AND ride_date>=?
                        ORDER BY id ASC LIMIT 1
                        """,
                        (vehicle_id, ride_id, FULL_REGISTRATION_FROM),
                    ).fetchone()
                    if prev is not None and start != prev["end_odometer"]:
                        raise ValueError(
                            f"Correctie verbreekt de kilometerketen van {plate}: vorige rit eindigt op {prev['end_odometer']} km"
                        )
                    if nxt is not None and end != nxt["start_odometer"]:
                        raise ValueError(
                            f"Correctie verbreekt de kilometerketen van {plate}: volgende rit begint op {nxt['start_odometer']} km"
                        )

                dep_lat, dep_lon = base.correction_coords(
                    payload, "departureCoords", current["departure_lat"], current["departure_lon"]
                )
                arr_lat, arr_lon = base.correction_coords(
                    payload, "arrivalCoords", current["arrival_lat"], current["arrival_lon"]
                )
                departure_time = base.validate_time(
                    payload.get("departureTime", current["departure_time"]), "Vertrektijd"
                )
                arrival_time = base.validate_time(
                    payload.get("arrivalTime", current["arrival_time"]), "Aankomsttijd"
                )
                db.execute(
                    """
                    UPDATE rides SET ride_date=?, ride_type=?, start_odometer=?, end_odometer=?, distance=?,
                        departure_address=?, arrival_address=?, departure_lat=?, departure_lon=?,
                        arrival_lat=?, arrival_lon=?, departure_time=?, arrival_time=?, notes=? WHERE id=?
                    """,
                    (
                        payload["date"], payload["type"], start, end, end-start,
                        str(payload["departureAddress"]).strip(), str(payload["arrivalAddress"]).strip(),
                        dep_lat, dep_lon, arr_lat, arr_lon, departure_time, arrival_time,
                        str(payload.get("notes") or "").strip(), ride_id,
                    ),
                )
                updated = db.execute(base.ride_select_sql("WHERE r.id = ?"), (ride_id,)).fetchone()
                new_snapshot = base.row_to_dict(updated)
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
    print(f"Rittenregistratie API + snelle invoer luistert op http://{base.HOST}:{base.PORT}")
    print(f"Registratiebeleid: zakelijk t/m 2026; volledig vanaf {FULL_REGISTRATION_FROM}; privélimiet {PRIVATE_KM_LIMIT} km")
    base.ThreadingHTTPServer((base.HOST, base.PORT), Handler).serve_forever()