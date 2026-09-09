#!/usr/bin/env python3
import sqlite3
from datetime import datetime, timezone
from urllib.parse import urlparse

import server as base

QUICK_SCHEMA = """
CREATE TABLE IF NOT EXISTS quick_rides (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    start_captured_at TEXT NOT NULL,
    start_lat REAL NOT NULL,
    start_lon REAL NOT NULL,
    start_accuracy REAL,
    start_address TEXT NOT NULL DEFAULT '',
    end_captured_at TEXT,
    end_lat REAL,
    end_lon REAL,
    end_accuracy REAL,
    end_address TEXT NOT NULL DEFAULT '',
    created_at TEXT NOT NULL,
    archived_at TEXT
);
CREATE INDEX IF NOT EXISTS idx_quick_rides_open ON quick_rides(archived_at, end_captured_at, id);
"""


def db_connect():
    db = base.db_connect()
    db.executescript(QUICK_SCHEMA)
    return db


def quick_to_dict(row):
    return {
        "id": row["id"],
        "startCapturedAt": row["start_captured_at"],
        "startCoords": {"lat": row["start_lat"], "lon": row["start_lon"], "accuracy": row["start_accuracy"]},
        "startAddress": row["start_address"],
        "endCapturedAt": row["end_captured_at"],
        "endCoords": None if row["end_lat"] is None else {"lat": row["end_lat"], "lon": row["end_lon"], "accuracy": row["end_accuracy"]},
        "endAddress": row["end_address"],
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
    return captured_at, lat, lon, accuracy, address


class Handler(base.Handler):
    server_version = "RittenregistratieAPI/0.8"

    def do_GET(self):
        if urlparse(self.path).path != "/api/quick-rides":
            return super().do_GET()
        try:
            with db_connect() as db:
                rows = db.execute(
                    "SELECT * FROM quick_rides WHERE archived_at IS NULL ORDER BY id DESC LIMIT 100"
                ).fetchall()
            self.send_json(200, {"quickRides": [quick_to_dict(row) for row in rows]})
        except sqlite3.Error:
            self.send_json(500, {"error": "Databasefout"})

    def do_POST(self):
        path = urlparse(self.path).path
        if path not in ("/api/quick-rides/start", "/api/quick-rides/end") and not (
            path.startswith("/api/quick-rides/") and path.endswith("/archive")
        ):
            return super().do_POST()

        try:
            if path == "/api/quick-rides/start":
                payload = self.read_json()
                captured_at, lat, lon, accuracy, address = capture_values(payload)
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
                            start_captured_at, start_lat, start_lon, start_accuracy, start_address, created_at
                        ) VALUES (?, ?, ?, ?, ?, ?)
                        """,
                        (captured_at, lat, lon, accuracy, address, created_at),
                    )
                    row = db.execute("SELECT * FROM quick_rides WHERE id=?", (cursor.lastrowid,)).fetchone()
                    db.commit()
                self.send_json(201, {"quickRide": quick_to_dict(row)})
                return

            if path == "/api/quick-rides/end":
                payload = self.read_json()
                captured_at, lat, lon, accuracy, address = capture_values(payload)
                with db_connect() as db:
                    row = db.execute(
                        "SELECT * FROM quick_rides WHERE archived_at IS NULL AND end_captured_at IS NULL ORDER BY id DESC LIMIT 1"
                    ).fetchone()
                    if row is None:
                        raise ValueError("Er is geen open beginpunt. Registreer eerst het beginpunt.")
                    db.execute(
                        """
                        UPDATE quick_rides
                        SET end_captured_at=?, end_lat=?, end_lon=?, end_accuracy=?, end_address=?
                        WHERE id=?
                        """,
                        (captured_at, lat, lon, accuracy, address, row["id"]),
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


if __name__ == "__main__":
    with db_connect():
        pass
    print(f"Rittenregistratie API + snelle invoer luistert op http://{base.HOST}:{base.PORT}")
    base.ThreadingHTTPServer((base.HOST, base.PORT), Handler).serve_forever()
