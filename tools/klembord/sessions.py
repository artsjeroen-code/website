
import time

from auth import (
    generate_session_token,
    get_session_token,
    hash_session_token,
    session_expiry,
)
from config import SESSION_HOURS


def create_session(db):
    """Maak een nieuwe sessie aan."""
    now = int(time.time())
    token = generate_session_token()

    db.execute(
        "DELETE FROM sessions WHERE expires_at <= ?",
        (now,),
    )

    db.execute(
        """
        INSERT INTO sessions
        (token_hash, created_at, expires_at)
        VALUES (?, ?, ?)
        """,
        (
            hash_session_token(token),
            now,
            session_expiry(SESSION_HOURS),
        ),
    )

    db.commit()
    return token


def valid_session(db, headers):
    """Controleer of de sessie bestaat en niet verlopen is."""
    token = get_session_token(headers)

    if not token:
        return False

    now = int(time.time())

    row = db.execute(
        """
        SELECT expires_at FROM sessions
        WHERE token_hash = ?
        """,
        (hash_session_token(token),),
    ).fetchone()

    if not row:
        return False

    if row[0] <= now:
        db.execute(
            "DELETE FROM sessions WHERE token_hash = ?",
            (hash_session_token(token),),
        )
        db.commit()
        return False

    return True


def revoke_session(db, headers):
    """Beëindig de huidige sessie."""
    token = get_session_token(headers)

    if token:
        db.execute(
            "DELETE FROM sessions WHERE token_hash = ?",
            (hash_session_token(token),),
        )
        db.commit()
