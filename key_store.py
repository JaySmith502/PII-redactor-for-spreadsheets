import base64
import binascii
from pathlib import Path

try:
    from dotenv import dotenv_values
except ImportError:
    dotenv_values = None


KEY_NAME = "SECRET_KEY"
KEY_BYTES = 64
DEFAULT_ENV_FILE = Path(__file__).resolve().with_name(".env")


class KeyConfigurationError(RuntimeError):
    """The web encryption key is missing or invalid."""


def _decode_urlsafe(value: str) -> bytes:
    padding = "=" * (-len(value) % 4)
    try:
        return base64.urlsafe_b64decode(value + padding)
    except (ValueError, binascii.Error) as exc:
        raise KeyConfigurationError("SECRET_KEY is not valid URL-safe Base64") from exc


def load_web_key(env_file: Path | None = None) -> bytes:
    path = (env_file or DEFAULT_ENV_FILE).resolve()
    if not path.is_file():
        raise KeyConfigurationError(".env is missing")
    if dotenv_values is None:
        raise KeyConfigurationError("python-dotenv is required to load .env")
    values = dotenv_values(path)
    encoded = values.get(KEY_NAME)
    if not encoded or not isinstance(encoded, str):
        raise KeyConfigurationError(".env must contain SECRET_KEY")
    key = _decode_urlsafe(encoded.strip())
    if len(key) != KEY_BYTES:
        raise KeyConfigurationError(f"SECRET_KEY must decode to exactly {KEY_BYTES} bytes")
    return key
