#!/usr/bin/env python3
import json
import os
import sqlite3
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

HOST = os.environ.get("FOCUS_HOST", "127.0.0.1")
PORT = int(os.environ.get("FOCUS_PORT", "8766"))
DB_PATH = os.environ.get("FOCUS_DB", "/var/lib/focustimer/focus.db")

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

class Handler(BaseHTTPRequestHandler):
    def send_json(self, status, payload):
        body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self):
        if self.path == "/health":
            try:
                with db_connect() as db:
                    db.execute("SELECT 1").fetchone()
                self.send_json(200, {"status": "ok"})
            except Exception:
                self.send_json(500, {"status": "error"})
            return

        self.send_json(404, {"error": "Not found"})

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
