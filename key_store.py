import base64
import binascii
import os
from pathlib import Path

try:
    from dotenv import dotenv_values
except ImportError:
    dotenv_values = None


KEY_NAME = "SECRET_KEY"
KEY_BYTES = 64
DEFAULT_ENV_FILE = Path(__file__).resolve().with_name(".env")
TRUTHY_SETTINGS = {"1", "true", "yes", "on"}


class KeyConfigurationError(RuntimeError):
    """The web encryption key is missing or invalid."""


def _decode_urlsafe(value: str) -> bytes:
    padding = "=" * (-len(value) % 4)
    try:
        return base64.urlsafe_b64decode(value + padding)
    except (ValueError, binascii.Error) as exc:
        raise KeyConfigurationError("SECRET_KEY is not valid URL-safe Base64") from exc


def setting_enabled(name: str, env_file: Path | None = None) -> bool:
    """Return whether boolean *name* is switched on, defaulting to off.

    The process environment is checked first so an operator can override the
    file, then ``.env``. Anything other than an explicit truthy value counts as
    disabled, so an unsafe feature stays off when the setting is absent or
    misspelled.
    """
    value = os.environ.get(name)
    if value is None and dotenv_values is not None:
        path = (env_file or DEFAULT_ENV_FILE).resolve()
        if path.is_file():
            value = dotenv_values(path).get(name)
    return isinstance(value, str) and value.strip().casefold() in TRUTHY_SETTINGS


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
