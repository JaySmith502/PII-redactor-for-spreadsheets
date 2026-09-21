from pathlib import Path

from key_store import DEFAULT_ENV_FILE, KeyConfigurationError, load_web_key


path = DEFAULT_ENV_FILE
print("exists=", path.exists())

try:
    key = load_web_key(path)
except KeyConfigurationError as exc:
    print("valid_64byte=", False)
    print("error=", exc)
else:
    raw = Path(path).read_text(encoding="utf-8")
    secret_line = next(
        (
            line
            for line in raw.splitlines()
            if line.strip() and not line.lstrip().startswith("#") and line.split("=", 1)[0].strip() == "SECRET_KEY"
        ),
        "",
    )
    value = secret_line.split("=", 1)[1].strip().strip("\"'") if "=" in secret_line else ""
    print("secret_len=", len(value))
    print("decoded_len=", len(key))
    print("valid_64byte=", len(key) == 64)
