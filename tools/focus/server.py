#!/usr/bin/env python3
import hashlib
import hmac
import json
import os
import secrets
import sqlite3
import threading
import time
from http.cookies import SimpleCookie
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import urlparse

HOST = os.environ.get("FOCUS_HOST", "127.0.0.1")
PORT = int(os.environ.get("FOCUS_PORT", "8768"))
DB_PATH = os.environ.get("FOCUS_DB", "/var/lib/focustimer/focus.db")
PASSWORD_HASH = os.environ.get("FOCUS_PASSWORD_HASH", "").strip()
SESSION_DAYS = int(os.environ.get("FOCUS_SESSION_DAYS", "30"))

COOKIE_NAME = "focus_session"
COOKIE_PATH = "/tools/focus/"
PASSWORD_MAX_FAILURES = 5
PASSWORD_LOCK_SECONDS = 60

PASSWORD_FAILURES = {}
PASSWORD_LOCK = threading.Lock()

SCHEMA = """
CREATE TABLE IF NOT EXISTS tasks (
    id TEXT PRIMARY KEY,
    text TEXT NOT NULL,
    completed INTEGER NOT NULL DEFAULT 0 CHECK (completed IN (0, 1)),
    active INTEGER NOT NULL DEFAULT 0 CHECK (active IN (0, 1)),
    note TEXT NOT NULL DEFAULT '',
    estimated_blocks INTEGER NOT NULL DEFAULT 1 CHECK (estimated_blocks BETWEEN 1 AND 99),
    focus_blocks_done INTEGER NOT NULL DEFAULT 0 CHECK (focus_blocks_done >= 0),
    sort_order INTEGER NOT NULL DEFAULT 0,
    updated_at TEXT NOT NULL,
    deleted_at TEXT
);
CREATE INDEX IF NOT EXISTS idx_tasks_sort_order ON tasks(sort_order, updated_at);

CREATE TABLE IF NOT EXISTS sessions (
    token_hash TEXT PRIMARY KEY,
    created_at INTEGER NOT NULL,
    expires_at INTEGER NOT NULL
);
CREATE INDEX IF NOT EXISTS idx_focus_sessions_expires ON sessions(expires_at);
"""

def db_connect():
    connection = sqlite3.connect(DB_PATH)
    connection.row_factory = sqlite3.Row
    return connection

def init_db():
    os.makedirs(os.path.dirname(DB_PATH), exist_ok=True)
    with db_connect() as db:
        db.executescript(SCHEMA)
        db.commit()

def token_hash(token):
    return hashlib.sha256(token.encode("utf-8")).hexdigest()

def session_token_from_headers(headers):
    raw = headers.get("Cookie", "")
    if not raw:
        return None
    cookie = SimpleCookie()
    try:
        cookie.load(raw)
    except Exception:
        return None
    morsel = cookie.get(COOKIE_NAME)
    return morsel.value if morsel else None

def create_session(db):
    now = int(time.time())
    expires_at = now + SESSION_DAYS * 86400
    token = secrets.token_urlsafe(32)
    db.execute("DELETE FROM sessions WHERE expires_at <= ?", (now,))
    db.execute(
        "INSERT INTO sessions (token_hash, created_at, expires_at) VALUES (?, ?, ?)",
        (token_hash(token), now, expires_at),
    )
    db.commit()
    return token

def valid_session(headers):
    token = session_token_from_headers(headers)
    if not token:
        return False
    now = int(time.time())
    with db_connect() as db:
        row = db.execute(
            "SELECT expires_at FROM sessions WHERE token_hash = ?",
            (token_hash(token),),
        ).fetchone()
        if not row:
            return False
        if row["expires_at"] <= now:
            db.execute("DELETE FROM sessions WHERE token_hash = ?", (token_hash(token),))
            db.commit()
            return False
        return True

def verify_password(password):
    if not PASSWORD_HASH:
        return False
    try:
        algorithm, iterations_text, salt_hex, expected_hex = PASSWORD_HASH.split("$", 3)
        if algorithm != "pbkdf2_sha256":
            return False
        iterations = int(iterations_text)
        salt = bytes.fromhex(salt_hex)
        expected = bytes.fromhex(expected_hex)
        actual = hashlib.pbkdf2_hmac(
            "sha256",
            password.encode("utf-8"),
            salt,
            iterations,
        )
        return hmac.compare_digest(actual, expected)
    except (ValueError, TypeError):
        return False

def client_ip(headers, fallback):
    forwarded = headers.get("X-Real-IP", "").strip()
    return forwarded or fallback

def password_rate_limited(ip):
    now = time.time()
    with PASSWORD_LOCK:
        failures = [
            stamp for stamp in PASSWORD_FAILURES.get(ip, [])
            if now - stamp < PASSWORD_LOCK_SECONDS
        ]
        PASSWORD_FAILURES[ip] = failures
        return len(failures) >= PASSWORD_MAX_FAILURES

def record_password_failure(ip):
    now = time.time()
    with PASSWORD_LOCK:
        failures = [
            stamp for stamp in PASSWORD_FAILURES.get(ip, [])
            if now - stamp < PASSWORD_LOCK_SECONDS
        ]
        failures.append(now)
        PASSWORD_FAILURES[ip] = failures

def clear_password_failures(ip):
    with PASSWORD_LOCK:
        PASSWORD_FAILURES.pop(ip, None)

class Handler(BaseHTTPRequestHandler):
    server_version = "FocusSync/0.2"

    def send_json(self, status, payload, session_token=None, clear_session=False):
        body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        if session_token:
            max_age = SESSION_DAYS * 86400
            self.send_header(
                "Set-Cookie",
                f"{COOKIE_NAME}={session_token}; Path={COOKIE_PATH}; Max-Age={max_age}; HttpOnly; Secure; SameSite=Strict",
            )
        elif clear_session:
            self.send_header(
                "Set-Cookie",
                f"{COOKIE_NAME}=; Path={COOKIE_PATH}; Max-Age=0; HttpOnly; Secure; SameSite=Strict",
            )
        self.end_headers()
        self.wfile.write(body)

    def read_json(self):
        length = int(self.headers.get("Content-Length", "0"))
        if length <= 0 or length > 16384:
            raise ValueError("Ongeldige requestgrootte")
        try:
            return json.loads(self.rfile.read(length).decode("utf-8"))
        except (UnicodeDecodeError, json.JSONDecodeError):
            raise ValueError("Ongeldige JSON")

    def do_GET(self):
        path = urlparse(self.path).path

        if path == "/health":
            try:
                with db_connect() as db:
                    db.execute("SELECT 1").fetchone()
                self.send_json(200, {"status": "ok"})
            except Exception:
                self.send_json(500, {"status": "error"})
            return

        if path == "/auth/status":
            self.send_json(200, {
                "authenticated": valid_session(self.headers),
                "passwordEnabled": bool(PASSWORD_HASH),
            })
            return

        self.send_json(404, {"error": "Not found"})

    def do_POST(self):
        path = urlparse(self.path).path

        try:
            if path == "/auth/login":
                if not PASSWORD_HASH:
                    self.send_json(503, {"error": "Focus-login is nog niet ingesteld"})
                    return

                ip = client_ip(self.headers, self.client_address[0])
                if password_rate_limited(ip):
                    self.send_json(429, {"error": "Te veel mislukte pogingen. Probeer het over een minuut opnieuw."})
                    return

                payload = self.read_json()
                password = str(payload.get("password") or "")
                if not verify_password(password):
                    record_password_failure(ip)
                    self.send_json(401, {"error": "Onjuist wachtwoord"})
                    return

                clear_password_failures(ip)
                with db_connect() as db:
                    token = create_session(db)
                self.send_json(200, {"ok": True}, session_token=token)
                return

            if path == "/auth/logout":
                token = session_token_from_headers(self.headers)
                if token:
                    with db_connect() as db:
                        db.execute(
                            "DELETE FROM sessions WHERE token_hash = ?",
                            (token_hash(token),),
                        )
                        db.commit()
                self.send_json(200, {"ok": True}, clear_session=True)
                return

            self.send_json(404, {"error": "Not found"})
        except ValueError as error:
            self.send_json(400, {"error": str(error)})

    def log_message(self, fmt, *args):
        print("%s - - [%s] %s" % (
            self.client_address[0],
            self.log_date_time_string(),
            fmt % args,
        ))

if __name__ == "__main__":
    init_db()
    server = ThreadingHTTPServer((HOST, PORT), Handler)
    print(f"Focus API luistert op http://{HOST}:{PORT}")
    server.serve_forever()
