
import time

MAX_MESSAGE_BYTES = 16384
MESSAGE_TTL_SECONDS = 1800
VALID_DIRECTIONS = {"incoming", "outgoing"}


def cleanup_expired(db):
    """Verwijder verlopen berichten."""
    db.execute(
        "DELETE FROM messages WHERE expires_at <= ?",
        (int(time.time()),)
    )
    db.commit()


def save_message(db, direction, content):
    """Sla een tijdelijk bericht op."""
    if direction not in VALID_DIRECTIONS:
        raise ValueError("Invalid direction")

    if not isinstance(content, str):
        raise ValueError("Invalid content")

    if not content.strip():
        raise ValueError("Empty message")

    if len(content.encode("utf-8")) > MAX_MESSAGE_BYTES:
        raise ValueError("Message too large")

    now = int(time.time())

    cleanup_expired(db)

    cursor = db.execute(
        """
        INSERT INTO messages
        (direction, content, created_at, expires_at)
        VALUES (?, ?, ?, ?)
        """,
        (
            direction,
            content,
            now,
            now + MESSAGE_TTL_SECONDS
        )
    )

    db.commit()
    return cursor.lastrowid


def get_messages(db, direction):
    """Haal actieve berichten op."""
    if direction is not None and direction not in VALID_DIRECTIONS:
        raise ValueError("Invalid direction")

    cleanup_expired(db)

    if direction is None:
        rows = db.execute(
            """
            SELECT id, direction, content, created_at, expires_at
            FROM messages
            WHERE expires_at > ?
            ORDER BY created_at DESC, id DESC
            """,
            (int(time.time()),)
        ).fetchall()
    else:
        rows = db.execute(
            """
            SELECT id, direction, content, created_at, expires_at
            FROM messages
            WHERE direction = ? AND expires_at > ?
            ORDER BY created_at DESC, id DESC
            """,
            (direction, int(time.time()))
        ).fetchall()

    return [dict(row) for row in rows]


def delete_message(db, message_id):
    """Verwijder een bericht."""
    if not isinstance(message_id, int) or message_id <= 0:
        raise ValueError("Invalid message ID")

    cleanup_expired(db)

    cursor = db.execute(
        "DELETE FROM messages WHERE id = ?",
        (message_id,)
    )

    db.commit()
    return cursor.rowcount > 0
