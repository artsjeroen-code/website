#!/usr/bin/env python3
import json
import os
import sqlite3
from datetime import datetime, timezone
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import urlparse

HOST = os.environ.get("RITTEN_HOST", "127.0.0.1")
PORT = int(os.environ.get("RITTEN_PORT", "8765"))
DB_PATH = os.environ.get("RITTEN_DB", "/var/lib/rittenregistratie/ritten.db")

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

    try:
        datetime.strptime(payload["date"], "%Y-%m-%d")
    except (TypeError, ValueError):
        raise ValueError("Datum moet YYYY-MM-DD zijn")

    return start, end


class Handler(BaseHTTPRequestHandler):
    server_version = "RittenregistratieAPI/0.1"

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

        self.send_json(404, {"error": "Niet gevonden"})

    def do_POST(self):
        path = urlparse(self.path).path
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
