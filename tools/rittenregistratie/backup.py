#!/usr/bin/env python3
import os
import sqlite3
from datetime import datetime, timedelta, timezone
from pathlib import Path

DB_PATH = Path(os.environ.get("RITTEN_DB", "/var/lib/rittenregistratie/ritten.db"))
BACKUP_DIR = Path(os.environ.get("RITTEN_BACKUP_DIR", "/var/backups/rittenregistratie"))
RETENTION_DAYS = int(os.environ.get("RITTEN_BACKUP_RETENTION_DAYS", "35"))


def create_backup():
    if not DB_PATH.is_file():
        raise FileNotFoundError(f"Database niet gevonden: {DB_PATH}")

    BACKUP_DIR.mkdir(parents=True, exist_ok=True)
    now = datetime.now(timezone.utc)
    stamp = now.strftime("%Y%m%dT%H%M%SZ")
    final_path = BACKUP_DIR / f"ritten-{stamp}.db"
    temp_path = BACKUP_DIR / f".ritten-{stamp}.tmp"

    if temp_path.exists():
        temp_path.unlink()

    source = sqlite3.connect(f"file:{DB_PATH}?mode=ro", uri=True)
    destination = sqlite3.connect(temp_path)
    try:
        source.backup(destination)
        result = destination.execute("PRAGMA integrity_check").fetchone()
        if not result or result[0] != "ok":
            raise RuntimeError(f"Integriteitscontrole mislukt: {result}")
    finally:
        destination.close()
        source.close()

    os.chmod(temp_path, 0o640)
    temp_path.replace(final_path)
    return final_path


def rotate_backups():
    cutoff = datetime.now(timezone.utc) - timedelta(days=RETENTION_DAYS)
    removed = []
    for path in BACKUP_DIR.glob("ritten-*.db"):
        modified = datetime.fromtimestamp(path.stat().st_mtime, tz=timezone.utc)
        if modified < cutoff:
            path.unlink()
            removed.append(path.name)
    return removed


def main():
    backup = create_backup()
    removed = rotate_backups()
    print(f"Back-up gemaakt: {backup}")
    print(f"Bewaartermijn: {RETENTION_DAYS} dagen")
    if removed:
        print("Verwijderd: " + ", ".join(removed))


if __name__ == "__main__":
    main()
