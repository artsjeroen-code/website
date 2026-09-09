#!/usr/bin/env python3
import base64
import hashlib
import hmac
import json
import os
import secrets
import sqlite3
import threading
import time
import traceback
from datetime import datetime, timedelta, timezone
from http.cookies import SimpleCookie
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import urlparse

import fido2.features

fido2.features.webauthn_json_mapping.enabled = True

from fido2.server import Fido2Server
from fido2.webauthn import AttestedCredentialData, AuthenticationResponse, RegistrationResponse

HOST = os.environ.get("RITTEN_AUTH_HOST", "127.0.0.1")
PORT = int(os.environ.get("RITTEN_AUTH_PORT", "8766"))
DB_PATH = os.environ.get("RITTEN_AUTH_DB", "/var/lib/rittenregistratie/auth.db")
RP_ID = os.environ.get("RITTEN_RP_ID", "artsjeroen.ddns.net")
RP_NAME = os.environ.get("RITTEN_RP_NAME", "Rittenregistratie")
USER_NAME = os.environ.get("RITTEN_AUTH_USER", "jeroen")
SESSION_DAYS = int(os.environ.get("RITTEN_SESSION_DAYS", "30"))
PASSWORD_HASH = os.environ.get("RITTEN_AUTH_PASSWORD_HASH", "").strip()
COOKIE_NAME = "ritten_session"
COOKIE_PATH = "/tools/rittenregistratie/"
TX_TTL_SECONDS = 300
PASSWORD_MAX_FAILURES = 5
PASSWORD_LOCK_SECONDS = 60

SERVER = Fido2Server({"id": RP_ID, "name": RP_NAME})
TX_LOCK = threading.Lock()
TRANSACTIONS = {}
PASSWORD_LOCK = threading.Lock()
PASSWORD_FAILURES = {}

SCHEMA = """
CREATE TABLE IF NOT EXISTS passkeys (
    credential_id BLOB PRIMARY KEY,
    credential_data BLOB NOT NULL,
    label TEXT NOT NULL DEFAULT '',
    created_at TEXT NOT NULL,
    last_used_at TEXT
);

CREATE TABLE IF NOT EXISTS sessions (
    token_hash TEXT PRIMARY KEY,
    created_at TEXT NOT NULL,
    expires_at TEXT NOT NULL
);
CREATE INDEX IF NOT EXISTS idx_sessions_expires ON sessions(expires_at);
"""


def db_connect():
    os.makedirs(os.path.dirname(DB_PATH), exist_ok=True)
    db = sqlite3.connect(DB_PATH)
    db.row_factory = sqlite3.Row
    db.executescript(SCHEMA)
    return db


def b64url(data):
    return base64.urlsafe_b64encode(bytes(data)).rstrip(b"=").decode("ascii")


def b64decode(value):
    text = str(value or "")
    return base64.urlsafe_b64decode(text + "=" * ((4 - len(text) % 4) % 4))


def jsonable(value):
    if isinstance(value, bytes):
        return b64url(value)
    if isinstance(value, dict):
        return {str(k): jsonable(v) for k, v in value.items()}
    if isinstance(value, (list, tuple)):
        return [jsonable(v) for v in value]
    if hasattr(value, "value"):
        return jsonable(value.value)
    if hasattr(value, "items"):
        return {str(k): jsonable(v) for k, v in value.items()}
    return value


def options_json(options):
    return jsonable(dict(options))


def cleanup_transactions():
    cutoff = time.time() - TX_TTL_SECONDS
    with TX_LOCK:
        for token in [key for key, tx in TRANSACTIONS.items() if tx["created"] < cutoff]:
            TRANSACTIONS.pop(token, None)


def put_transaction(kind, state):
    cleanup_transactions()
    token = secrets.token_urlsafe(24)
    with TX_LOCK:
        TRANSACTIONS[token] = {"kind": kind, "state": state, "created": time.time()}
    return token


def pop_transaction(token, kind):
    cleanup_transactions()
    with TX_LOCK:
        tx = TRANSACTIONS.pop(str(token or ""), None)
    if not tx or tx["kind"] != kind:
        raise ValueError("Aanmeldpoging is verlopen; probeer opnieuw")
    return tx["state"]


def load_credentials(db):
    rows = db.execute("SELECT * FROM passkeys ORDER BY created_at ASC").fetchall()
    return [(row, AttestedCredentialData(bytes(row["credential_data"]))) for row in rows]


def token_hash(token):
    return hashlib.sha256(token.encode("utf-8")).hexdigest()


def create_session(db):
    token = secrets.token_urlsafe(32)
    now = datetime.now(timezone.utc)
    expires = now + timedelta(days=SESSION_DAYS)
    db.execute("DELETE FROM sessions WHERE expires_at <= ?", (now.isoformat(),))
    db.execute(
        "INSERT INTO sessions (token_hash, created_at, expires_at) VALUES (?, ?, ?)",
        (token_hash(token), now.isoformat(), expires.isoformat()),
    )
    db.commit()
    return token


def session_token_from_headers(headers):
    raw = headers.get("Cookie") or ""
    cookie = SimpleCookie()
    try:
        cookie.load(raw)
    except Exception:
        return None
    morsel = cookie.get(COOKIE_NAME)
    return morsel.value if morsel else None


def valid_session(headers):
    token = session_token_from_headers(headers)
    if not token:
        return False
    now = datetime.now(timezone.utc).isoformat()
    with db_connect() as db:
        db.execute("DELETE FROM sessions WHERE expires_at <= ?", (now,))
        row = db.execute(
            "SELECT 1 FROM sessions WHERE token_hash = ? AND expires_at > ?",
            (token_hash(token), now),
        ).fetchone()
        db.commit()
    return row is not None


def user_entity():
    user_id = hashlib.sha256((RP_ID + "\0" + USER_NAME).encode("utf-8")).digest()[:32]
    return {"id": user_id, "name": USER_NAME, "displayName": USER_NAME}


def verify_password(password):
    if not PASSWORD_HASH:
        return False
    try:
        scheme, iterations_text, salt_text, digest_text = PASSWORD_HASH.split("$", 3)
        if scheme != "pbkdf2_sha256":
            return False
        iterations = int(iterations_text)
        if iterations < 100000:
            return False
        salt = b64decode(salt_text)
        expected = b64decode(digest_text)
        actual = hashlib.pbkdf2_hmac("sha256", str(password).encode("utf-8"), salt, iterations)
        return hmac.compare_digest(actual, expected)
    except (TypeError, ValueError, base64.binascii.Error):
        return False


def password_rate_limited(client_ip):
    now = time.time()
    with PASSWORD_LOCK:
        failures = [stamp for stamp in PASSWORD_FAILURES.get(client_ip, []) if now - stamp < PASSWORD_LOCK_SECONDS]
        PASSWORD_FAILURES[client_ip] = failures
        return len(failures) >= PASSWORD_MAX_FAILURES


def record_password_failure(client_ip):
    now = time.time()
    with PASSWORD_LOCK:
        failures = [stamp for stamp in PASSWORD_FAILURES.get(client_ip, []) if now - stamp < PASSWORD_LOCK_SECONDS]
        failures.append(now)
        PASSWORD_FAILURES[client_ip] = failures


def clear_password_failures(client_ip):
    with PASSWORD_LOCK:
        PASSWORD_FAILURES.pop(client_ip, None)


class Handler(BaseHTTPRequestHandler):
    server_version = "RittenAuth/1.3"

    def log_message(self, fmt, *args):
        print(f"{self.address_string()} - {fmt % args}")

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
        if length <= 0 or length > 131072:
            raise ValueError("Ongeldige requestgrootte")
        try:
            return json.loads(self.rfile.read(length).decode("utf-8"))
        except (UnicodeDecodeError, json.JSONDecodeError):
            raise ValueError("Ongeldige JSON")

    def do_GET(self):
        path = urlparse(self.path).path
        if path == "/api/auth/check":
            if valid_session(self.headers):
                self.send_json(200, {"ok": True})
            else:
                self.send_json(401, {"error": "Niet ingelogd"})
            return

        if path == "/api/auth/status":
            with db_connect() as db:
                count = db.execute("SELECT COUNT(*) FROM passkeys").fetchone()[0]
            self.send_json(200, {
                "authenticated": valid_session(self.headers),
                "passkeyCount": count,
                "webauthn": True,
                "passwordEnabled": bool(PASSWORD_HASH),
            })
            return

        self.send_json(404, {"error": "Niet gevonden"})

    def do_POST(self):
        path = urlparse(self.path).path
        try:
            if path == "/api/auth/register/begin":
                with db_connect() as db:
                    credentials = [credential for _, credential in load_credentials(db)]
                options, state = SERVER.register_begin(
                    user_entity(),
                    credentials,
                    resident_key_requirement="preferred",
                    user_verification="required",
                )
                tx = put_transaction("register", state)
                self.send_json(200, {"transaction": tx, "options": options_json(options)})
                return

            if path == "/api/auth/register/complete":
                payload = self.read_json()
                state = pop_transaction(payload.get("transaction"), "register")
                response = RegistrationResponse.from_dict(payload.get("credential") or {})
                auth_data = SERVER.register_complete(state, response)
                credential = auth_data.credential_data
                if credential is None:
                    raise ValueError("Passkey bevat geen bruikbare credential")
                now = datetime.now(timezone.utc).isoformat()
                with db_connect() as db:
                    db.execute(
                        "INSERT OR REPLACE INTO passkeys (credential_id, credential_data, label, created_at, last_used_at) VALUES (?, ?, ?, ?, COALESCE((SELECT last_used_at FROM passkeys WHERE credential_id=?), NULL))",
                        (credential.credential_id, bytes(credential), "Passkey", now, credential.credential_id),
                    )
                    token = create_session(db)
                self.send_json(201, {"ok": True}, session_token=token)
                return

            if path == "/api/auth/login/begin":
                with db_connect() as db:
                    credentials = [credential for _, credential in load_credentials(db)]
                if not credentials:
                    raise ValueError("Er is nog geen passkey geregistreerd")
                options, state = SERVER.authenticate_begin(
                    credentials,
                    user_verification="required",
                )
                tx = put_transaction("login", state)
                self.send_json(200, {"transaction": tx, "options": options_json(options)})
                return

            if path == "/api/auth/login/complete":
                payload = self.read_json()
                state = pop_transaction(payload.get("transaction"), "login")
                response = AuthenticationResponse.from_dict(payload.get("credential") or {})
                with db_connect() as db:
                    stored = load_credentials(db)
                    credentials = [credential for _, credential in stored]
                    credential = SERVER.authenticate_complete(
                        state,
                        credentials,
                        response,
                    )
                    db.execute(
                        "UPDATE passkeys SET last_used_at=? WHERE credential_id=?",
                        (datetime.now(timezone.utc).isoformat(), credential.credential_id),
                    )
                    token = create_session(db)
                self.send_json(200, {"ok": True}, session_token=token)
                return

            if path == "/api/auth/password":
                if not PASSWORD_HASH:
                    self.send_json(404, {"error": "Wachtwoord-login is niet ingesteld"})
                    return
                client_ip = self.client_address[0]
                if password_rate_limited(client_ip):
                    self.send_json(429, {"error": "Te veel mislukte pogingen. Wacht een minuut en probeer opnieuw."})
                    return
                payload = self.read_json()
                password = str(payload.get("password") or "")
                if not verify_password(password):
                    record_password_failure(client_ip)
                    self.send_json(401, {"error": "Onjuist wachtwoord"})
                    return
                clear_password_failures(client_ip)
                with db_connect() as db:
                    token = create_session(db)
                self.send_json(200, {"ok": True}, session_token=token)
                return

            if path == "/api/auth/logout":
                token = session_token_from_headers(self.headers)
                if token:
                    with db_connect() as db:
                        db.execute("DELETE FROM sessions WHERE token_hash=?", (token_hash(token),))
                        db.commit()
                self.send_json(200, {"ok": True}, clear_session=True)
                return

            self.send_json(404, {"error": "Niet gevonden"})
        except ValueError as error:
            self.send_json(400, {"error": str(error)})
        except Exception as error:
            print(f"Authenticatiefout: {type(error).__name__}: {error}")
            traceback.print_exc()
            self.send_json(500, {"error": "Authenticatie is mislukt"})


if __name__ == "__main__":
    with db_connect():
        pass
    print(f"Authenticatieservice luistert op http://{HOST}:{PORT}; RP={RP_ID}; password={'on' if PASSWORD_HASH else 'off'}")
    ThreadingHTTPServer((HOST, PORT), Handler).serve_forever()
