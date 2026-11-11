
import json
import os
import sqlite3
import threading
import time

from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import urlsplit, parse_qs

from config import (
    HOST,
    PORT,
    DB_PATH,
    PASSWORD_HASH,
    SESSION_HOURS,
    MAX_MESSAGE_BYTES,
    MESSAGE_TTL_SECONDS,
)

from auth import (
    verify_password,
    session_cookie,
    clear_session_cookie,
)

from sessions import (
    create_session,
    valid_session,
    revoke_session,
)

from messages import (
    get_messages,
    save_message,
    delete_message,
    cleanup_expired,
)

# Alleen deze herkomst mag browseracties uitvoeren.
# Voor staging later instellen via een environmentvariabele.
ALLOWED_ORIGIN = os.environ.get(
    "BEAM_ALLOWED_ORIGIN",
    "http://127.0.0.1:8770"
).rstrip("/")

MAX_LOGIN_BYTES = 4096
MAX_FAILURES = 5
LOCK_SECONDS = 60

LOGIN_FAILURES = {}
LOGIN_LOCK = threading.Lock()


SCHEMA = """
CREATE TABLE IF NOT EXISTS messages (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    direction TEXT NOT NULL CHECK (
        direction IN ('incoming', 'outgoing')
    ),
    content TEXT NOT NULL,
    created_at INTEGER NOT NULL,
    expires_at INTEGER NOT NULL
);

CREATE INDEX IF NOT EXISTS idx_messages_expiry
ON messages(expires_at);

CREATE TABLE IF NOT EXISTS sessions (
    token_hash TEXT PRIMARY KEY,
    created_at INTEGER NOT NULL,
    expires_at INTEGER NOT NULL
);

CREATE INDEX IF NOT EXISTS idx_sessions_expiry
ON sessions(expires_at);
"""


def connect_db():
    connection = sqlite3.connect(DB_PATH, timeout=10)
    connection.row_factory = sqlite3.Row
    return connection


def init_db():
    directory = os.path.dirname(DB_PATH)

    if not os.path.isdir(directory):
        raise RuntimeError(
            "Database directory does not exist"
        )

    with connect_db() as db:
        db.executescript(SCHEMA)


def login_blocked(client):
    now = time.monotonic()

    with LOGIN_LOCK:
        attempts, blocked_until = LOGIN_FAILURES.get(
            client, (0, 0)
        )

        if blocked_until > now:
            return True

        if blocked_until:
            LOGIN_FAILURES.pop(client, None)

        return False


def record_login_failure(client):
    now = time.monotonic()

    with LOGIN_LOCK:
        attempts, blocked_until = LOGIN_FAILURES.get(
            client, (0, 0)
        )

        if blocked_until and blocked_until <= now:
            attempts = 0

        attempts += 1

        if attempts >= MAX_FAILURES:
            LOGIN_FAILURES[client] = (
                attempts, now + LOCK_SECONDS
            )
        else:
            LOGIN_FAILURES[client] = (
                attempts, 0
            )


def clear_login_failures(client):
    with LOGIN_LOCK:
        LOGIN_FAILURES.pop(client, None)


CLEANUP_INTERVAL_SECONDS = 60


def cleanup_worker(stop_event):
    """Verwijder periodiek verlopen berichten en sessies."""
    while not stop_event.is_set():
        try:
            with connect_db() as db:
                cleanup_expired(db)

                db.execute(
                    "DELETE FROM sessions WHERE expires_at <= ?",
                    (int(time.time()),)
                )
                db.commit()

        except sqlite3.Error:
            # Meld een databasefout zonder berichtinhoud te loggen.
            print("Beam Me Up: database cleanup failed", flush=True)

        stop_event.wait(CLEANUP_INTERVAL_SECONDS)


class Handler(BaseHTTPRequestHandler):

    def send_json(self, status, payload, cookie=None):
        data = json.dumps(payload).encode("utf-8")

        self.send_response(status)

        self.send_header(
            "Content-Type",
            "application/json; charset=utf-8"
        )
        self.send_header(
            "Content-Length", str(len(data))
        )
        self.send_header("Cache-Control", "no-store")
        self.send_header("Pragma", "no-cache")
        self.send_header(
            "X-Content-Type-Options", "nosniff"
        )
        self.send_header("X-Frame-Options", "DENY")
        self.send_header(
            "Referrer-Policy", "no-referrer"
        )

        if cookie:
            self.send_header("Set-Cookie", cookie)

        self.end_headers()
        self.wfile.write(data)

    def origin_allowed(self):
        return self.headers.get("Origin") == ALLOWED_ORIGIN

    def read_json(self):
        if self.headers.get(
            "Content-Type", ""
        ).split(";")[0].strip().lower() != "application/json":
            raise ValueError("JSON required")

        try:
            length = int(
                self.headers.get("Content-Length", "-1")
            )
        except ValueError:
            raise ValueError("Invalid content length")

        if not 0 < length <= MAX_LOGIN_BYTES:
            raise ValueError("Invalid content length")

        raw = self.rfile.read(length)

        try:
            payload = json.loads(raw.decode("utf-8"))
        except (UnicodeDecodeError, json.JSONDecodeError):
            raise ValueError("Invalid JSON")

        if not isinstance(payload, dict):
            raise ValueError("JSON object required")

        return payload


    def read_message_json(self):
        """Lees een bericht met een begrensde requestgrootte."""
        content_type = self.headers.get(
            "Content-Type", ""
        ).split(";")[0].strip().lower()

        if content_type != "application/json":
            raise ValueError("JSON required")

        try:
            length = int(
                self.headers.get("Content-Length", "-1")
            )
        except ValueError:
            raise ValueError("Invalid content length")

        if not 0 < length <= 20000:
            raise ValueError("Request too large")

        raw = self.rfile.read(length)

        try:
            payload = json.loads(raw.decode("utf-8"))
        except (UnicodeDecodeError, json.JSONDecodeError):
            raise ValueError("Invalid JSON")

        if not isinstance(payload, dict):
            raise ValueError("JSON object required")

        return payload


    def do_GET(self):
        path = urlsplit(self.path).path

        if path == "/health":
            self.send_json(200, {
                "ok": True,
                "service": "beam-me-up"
            })
            return

        if path == "/auth/status":
            with connect_db() as db:
                authenticated = valid_session(
                    db, self.headers
                )

            self.send_json(200, {
                "authenticated": authenticated,
                "passwordEnabled": bool(PASSWORD_HASH)
            })
            return


        if path == "/messages":
            # Alleen ingelogde gebruikers mogen berichten lezen.
            with connect_db() as db:
                if not valid_session(db, self.headers):
                    self.send_json(401, {
                        "error": "Not authenticated"
                    })
                    return

                params = parse_qs(
                    urlsplit(self.path).query
                )

                directions = params.get("direction", [])

                if len(directions) != 1 or directions[0] not in (
                    "incoming", "outgoing"
                ):
                    self.send_json(400, {
                        "error": "Invalid direction"
                    })
                    return

                messages = get_messages(
                    db, directions[0]
                )

            self.send_json(200, {
                "messages": messages
            })
            return

        self.send_json(403, {
            "error": "Access denied"
        })

    def do_POST(self):
        path = urlsplit(self.path).path


        if path == "/messages":
            if not self.origin_allowed():
                self.send_json(403, {
                    "error": "Origin not allowed"
                })
                return

            with connect_db() as db:
                if not valid_session(db, self.headers):
                    self.send_json(401, {
                        "error": "Not authenticated"
                    })
                    return

                try:
                    payload = self.read_message_json()

                    message_id = save_message(
                        db,
                        payload.get("direction"),
                        payload.get("content")
                    )

                except ValueError:
                    self.send_json(400, {
                        "error": "Invalid message"
                    })
                    return

            self.send_json(201, {
                "ok": True,
                "id": message_id
            })
            return


        if path not in (
            "/auth/login",
            "/auth/logout"
        ):
            self.send_json(403, {
                "error": "Access denied"
            })
            return

        if not self.origin_allowed():
            self.send_json(403, {
                "error": "Origin not allowed"
            })
            return

        if path == "/auth/login":
            self.handle_login()
            return

        if path == "/auth/logout":
            self.handle_logout()
            return

    def handle_login(self):
        client = self.client_address[0]

        if not PASSWORD_HASH:
            self.send_json(503, {
                "error": "Authentication not configured"
            })
            return

        if login_blocked(client):
            self.send_json(429, {
                "error": "Too many attempts"
            })
            return

        try:
            payload = self.read_json()
        except ValueError:
            self.send_json(400, {
                "error": "Invalid request"
            })
            return

        password = payload.get("password")

        if not isinstance(password, str):
            self.send_json(400, {
                "error": "Invalid request"
            })
            return

        if not verify_password(
            password, PASSWORD_HASH
        ):
            record_login_failure(client)

            self.send_json(401, {
                "error": "Invalid credentials"
            })
            return

        clear_login_failures(client)

        with connect_db() as db:
            token = create_session(db)

        self.send_json(
            200,
            {"ok": True},
            cookie=session_cookie(
                token,
                max_age=SESSION_HOURS * 3600
            )
        )

    def handle_logout(self):
        with connect_db() as db:
            if not valid_session(db, self.headers):
                self.send_json(401, {
                    "error": "Not authenticated"
                })
                return

            revoke_session(db, self.headers)

        self.send_json(
            200,
            {"ok": True},
            cookie=clear_session_cookie()
        )

   
    def do_DELETE(self):
        path = urlsplit(self.path).path

        if not path.startswith("/messages/"):
            self.send_json(404, {
                "error": "Not found"
            })
            return

        if not self.origin_allowed():
            self.send_json(403, {
                "error": "Origin not allowed"
            })
            return

        with connect_db() as db:
            if not valid_session(db, self.headers):
                self.send_json(401, {
                    "error": "Not authenticated"
                })
                return

            message_id_text = path.removeprefix("/messages/")

            if (
                not message_id_text.isascii()
                or not message_id_text.isdecimal()
            ):
                self.send_json(400, {
                    "error": "Invalid message ID"
                })
                return

            message_id = int(message_id_text)

            try:
                deleted = delete_message(db, message_id)
            except ValueError:
                self.send_json(400, {
                    "error": "Invalid message ID"
                })
                return

        if not deleted:
            self.send_json(404, {
                "error": "Message not found"
            })
            return

        self.send_json(200, {
            "ok": True,
            "deleted": message_id
        })


if __name__ == "__main__":
    init_db()

    stop_event = threading.Event()

    cleanup_thread = threading.Thread(
        target=cleanup_worker,
        args=(stop_event,),
        name="beam-cleanup",
        daemon=True,
    )
    cleanup_thread.start()

    server = ThreadingHTTPServer(
        (HOST, PORT), Handler
    )

    print(
        f"Beam Me Up listening on {HOST}:{PORT}",
        flush=True
    )

    try:
        server.serve_forever()
    finally:
        stop_event.set()
        server.server_close()
        cleanup_thread.join(timeout=2)

