#!/usr/bin/env python3
import hmac
import json
import os
import sqlite3
from datetime import datetime
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import urlencode, urlparse
from urllib.request import Request, urlopen

import quick_server as quick

HOST = os.environ.get("RITTEN_SHORTCUT_HOST", "127.0.0.1")
PORT = int(os.environ.get("RITTEN_SHORTCUT_PORT", "8767"))
TOKEN = os.environ.get("RITTEN_SHORTCUT_TOKEN", "").strip()


def reverse_geocode(lat, lon):
    params = urlencode({
        "format": "jsonv2",
        "lat": str(lat),
        "lon": str(lon),
        "zoom": "18",
        "addressdetails": "1",
        "accept-language": "nl",
    })
    request = Request(
        f"https://nominatim.openstreetmap.org/reverse?{params}",
        headers={
            "Accept": "application/json",
            "User-Agent": "artsjeroen-rittenregistratie/1.0",
        },
    )
    try:
        with urlopen(request, timeout=6) as response:
            result = json.load(response)
        address = result.get("address") or {}
        road = address.get("road") or address.get("pedestrian") or address.get("residential") or ""
        street = " ".join(part for part in (road, address.get("house_number") or "") if part)
        locality = " ".join(part for part in (
            address.get("postcode") or "",
            address.get("city") or address.get("town") or address.get("village") or address.get("municipality") or "",
        ) if part)
        return ", ".join(part for part in (street, locality) if part) or result.get("display_name") or ""
    except Exception:
        return ""


class Handler(BaseHTTPRequestHandler):
    server_version = "RittenregistratieShortcut/1.3"

    def send_json(self, status, payload):
        body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(body)

    def authorized(self):
        if not TOKEN:
            return False
        header = self.headers.get("Authorization", "")
        prefix = "Bearer "
        if not header.startswith(prefix):
            return False
        return hmac.compare_digest(header[len(prefix):].strip(), TOKEN)

    def read_json(self):
        length = int(self.headers.get("Content-Length", "0") or "0")
        if length <= 0 or length > 8192:
            raise ValueError("Ongeldige aanvraag")
        try:
            return json.loads(self.rfile.read(length).decode("utf-8"))
        except (UnicodeDecodeError, json.JSONDecodeError):
            raise ValueError("Ongeldige JSON")

    def do_POST(self):
        if urlparse(self.path).path != "/api/shortcut":
            self.send_json(404, {"status": "not_found", "message": "Niet gevonden"})
            return
        if not self.authorized():
            self.send_json(401, {"status": "unauthorized", "message": "Shortcut-token is ongeldig"})
            return

        try:
            payload = self.read_json()
            action = str(payload.get("action") or "").strip().lower()
            if action not in ("start", "end"):
                raise ValueError("Actie moet start of end zijn")

            coords = payload.get("coords") or {}
            lat_value = payload.get("lat", coords.get("lat"))
            lon_value = payload.get("lon", coords.get("lon"))
            accuracy_value = payload.get("accuracy", coords.get("accuracy"))

            lat = quick.base.validate_coord(lat_value, -90, 90, "latitude")
            lon = quick.base.validate_coord(lon_value, -180, 180, "longitude")
            accuracy = accuracy_value
            if accuracy in (None, ""):
                accuracy = None
            else:
                accuracy = max(0.0, float(accuracy))

            odometer = quick.validate_optional_odometer(payload.get("odometer"))
            if odometer is None:
                raise ValueError("Kilometerstand ontbreekt")

            notes = quick.validate_notes(payload.get("notes")) if action == "start" else ""
            captured_at = datetime.now().astimezone().strftime("%Y-%m-%dT%H:%M:%S")
            address = str(payload.get("address") or "").strip() or reverse_geocode(lat, lon)
            if not address:
                address = f"GPS {lat:.6f}, {lon:.6f}"
            created_at = datetime.now().astimezone().isoformat()

            with quick.db_connect() as db:
                open_row = db.execute(
                    "SELECT * FROM quick_rides WHERE archived_at IS NULL AND end_captured_at IS NULL ORDER BY id DESC LIMIT 1"
                ).fetchone()

                if action == "start":
                    if open_row is not None:
                        self.send_json(200, {
                            "status": "already_running",
                            "message": "Er loopt al een rit. Registreer eerst het eindpunt.",
                            "quickRide": quick.quick_to_dict(open_row),
                        })
                        return
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
                    self.send_json(201, {
                        "status": "started",
                        "message": f"Rit gestart op {odometer} km.",
                        "quickRide": quick.quick_to_dict(row),
                    })
                    return

                if open_row is None:
                    self.send_json(200, {
                        "status": "no_active_ride",
                        "message": "Er is geen actieve rit. Registreer eerst het beginpunt.",
                    })
                    return

                if open_row["start_odometer"] is not None and odometer < open_row["start_odometer"]:
                    raise ValueError(
                        f"Eindstand {odometer} km kan niet lager zijn dan beginstand {open_row['start_odometer']} km"
                    )

                db.execute(
                    """
                    UPDATE quick_rides
                    SET end_captured_at=?, end_lat=?, end_lon=?, end_accuracy=?, end_address=?, end_odometer=?
                    WHERE id=?
                    """,
                    (captured_at, lat, lon, accuracy, address, odometer, open_row["id"]),
                )
                row = db.execute("SELECT * FROM quick_rides WHERE id=?", (open_row["id"],)).fetchone()
                db.commit()
                distance = None if row["start_odometer"] is None else odometer - row["start_odometer"]
                suffix = "" if distance is None else f" Afstand: {distance} km."
                self.send_json(200, {
                    "status": "ended",
                    "message": f"Rit beëindigd op {odometer} km.{suffix} Concept-rit staat klaar om aan te vullen.",
                    "quickRide": quick.quick_to_dict(row),
                })
        except (TypeError, ValueError) as error:
            self.send_json(400, {"status": "invalid_request", "message": str(error)})
        except sqlite3.Error:
            self.send_json(500, {"status": "database_error", "message": "Databasefout"})
        except Exception:
            self.send_json(500, {"status": "server_error", "message": "Registratie is mislukt"})

    def log_message(self, format, *args):
        return


if __name__ == "__main__":
    if not TOKEN:
        raise SystemExit("RITTEN_SHORTCUT_TOKEN ontbreekt")
    with quick.db_connect():
        pass
    print(f"Rittenregistratie Shortcut API luistert op http://{HOST}:{PORT}")
    ThreadingHTTPServer((HOST, PORT), Handler).serve_forever()
