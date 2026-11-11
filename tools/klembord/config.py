
import os

HOST = os.environ.get("BEAM_HOST", "127.0.0.1")
PORT = int(os.environ.get("BEAM_PORT", "8770"))

DB_PATH = os.environ.get(
    "BEAM_DB",
    "/var/lib/klembord/staging/klembord.db"
)

PASSWORD_HASH = os.environ.get(
    "BEAM_PASSWORD_HASH", ""
).strip()

SESSION_HOURS = 12
MESSAGE_TTL_SECONDS = 1800
MAX_MESSAGE_BYTES = 16384

COOKIE_NAME = "beam_session"
COOKIE_PATH = "/tools/klembord/"
