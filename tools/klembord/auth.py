
import hashlib
import hmac
import secrets
import time

from http.cookies import SimpleCookie

from config import COOKIE_NAME


def verify_password(password, stored_hash):
    """Controleer een wachtwoord tegen een PBKDF2-hash."""
    if not stored_hash:
        return False

    try:
        algorithm, iterations, salt_hex, expected_hex = (
            stored_hash.split("$", 3)
        )

        if algorithm != "pbkdf2_sha256":
            return False

        iterations = int(iterations)

        if not 100000 <= iterations <= 2000000:
            return False

        salt = bytes.fromhex(salt_hex)
        expected = bytes.fromhex(expected_hex)

        actual = hashlib.pbkdf2_hmac(
            "sha256",
            password.encode("utf-8"),
            salt,
            iterations
        )

        return hmac.compare_digest(actual, expected)

    except (ValueError, TypeError, AttributeError):
        return False


def generate_session_token():
    """Genereer een cryptografisch veilig sessietoken."""
    return secrets.token_urlsafe(32)


def hash_session_token(token):
    """Sla uitsluitend een hash van het sessietoken op."""
    return hashlib.sha256(
        token.encode("utf-8")
    ).hexdigest()


def get_session_token(headers):
    """Lees het sessietoken uit de cookie."""
    raw_cookie = headers.get("Cookie", "")

    if not raw_cookie:
        return None

    cookie = SimpleCookie()

    try:
        cookie.load(raw_cookie)
    except Exception:
        return None

    session = cookie.get(COOKIE_NAME)
    return session.value if session else None


def session_expiry(hours=12):
    return int(time.time()) + hours * 3600



def session_cookie(token, max_age=43200):
    """Maak een beveiligde sessiecookie."""
    from config import COOKIE_PATH

    return (
        f"{COOKIE_NAME}={token}; "
        f"Path={COOKIE_PATH}; "
        f"Max-Age={max_age}; "
        "HttpOnly; Secure; SameSite=Strict"
    )


def clear_session_cookie():
    """Verwijder de sessiecookie uit de browser."""
    from config import COOKIE_PATH

    return (
        f"{COOKIE_NAME}=; "
        f"Path={COOKIE_PATH}; "
        "Max-Age=0; "
        "HttpOnly; Secure; SameSite=Strict"
    )
