#!/usr/bin/env bash
set -euo pipefail

write_step() { printf '\n>> %s\n' "$1"; }
write_ok() { printf '   [OK] %s\n' "$1"; }
write_warn() { printf '   [!] %s\n' "$1"; }
write_fail() { printf '   [X] %s\n' "$1"; }

python_cmd=""

find_python_cmd() {
  python_cmd=""
  for candidate in python3 python; do
    if command -v "$candidate" >/dev/null 2>&1; then
      if "$candidate" - <<'PY' >/dev/null 2>&1
import sys
sys.exit(0 if sys.version_info >= (3, 10) else 1)
PY
      then
        python_cmd="$candidate"
        return 0
      fi
    fi
  done
  return 1
}

run_as_root() {
  if [ "$(id -u)" -eq 0 ]; then
    "$@"
  else
    if ! command -v sudo >/dev/null 2>&1; then
      write_fail "sudo is required to install Python automatically on this system."
      exit 1
    fi
    sudo "$@"
  fi
}

install_python() {
  case "$(uname -s)" in
    Darwin)
      if command -v brew >/dev/null 2>&1; then
        write_step "Installing Python with Homebrew..."
        brew install python
      else
        write_fail "Python 3.10+ was not found, and Homebrew is not installed."
        printf '   Install Python from https://www.python.org/downloads/macos/ or install Homebrew, then run ./setup.sh again.\n'
        exit 1
      fi
      ;;
    Linux)
      if command -v apt-get >/dev/null 2>&1; then
        write_step "Installing Python with apt..."
        run_as_root apt-get update
        run_as_root apt-get install -y python3 python3-pip python3-venv
      elif command -v dnf >/dev/null 2>&1; then
        write_step "Installing Python with dnf..."
        run_as_root dnf install -y python3 python3-pip
      elif command -v yum >/dev/null 2>&1; then
        write_step "Installing Python with yum..."
        run_as_root yum install -y python3 python3-pip
      elif command -v pacman >/dev/null 2>&1; then
        write_step "Installing Python with pacman..."
        run_as_root pacman -Sy --noconfirm python python-pip
      elif command -v zypper >/dev/null 2>&1; then
        write_step "Installing Python with zypper..."
        run_as_root zypper install -y python3 python3-pip
      else
        write_fail "Python 3.10+ was not found, and no supported package manager was detected."
        printf '   Install Python 3.10 or newer with your system package manager, then run ./setup.sh again.\n'
        exit 1
      fi
      ;;
    *)
      write_fail "Automatic Python installation is only supported on macOS and common Linux distributions."
      printf '   Install Python 3.10 or newer, then run ./setup.sh again.\n'
      exit 1
      ;;
  esac
}

test_secret_key() {
  local key="${1:-}"
  if [ -z "$key" ]; then
    return 1
  fi
  "$python_cmd" - "$key" <<'PY'
import base64
import sys

value = sys.argv[1].strip()
padding = "=" * (-len(value) % 4)
try:
    decoded = base64.urlsafe_b64decode(value + padding)
except Exception:
    sys.exit(1)
sys.exit(0 if len(decoded) == 64 else 1)
PY
}

copy_to_clipboard() {
  local value="$1"
  if command -v pbcopy >/dev/null 2>&1; then
    printf '%s' "$value" | pbcopy
    return 0
  fi
  if command -v wl-copy >/dev/null 2>&1; then
    printf '%s' "$value" | wl-copy
    return 0
  fi
  if command -v xclip >/dev/null 2>&1; then
    printf '%s' "$value" | xclip -selection clipboard
    return 0
  fi
  if command -v xsel >/dev/null 2>&1; then
    printf '%s' "$value" | xsel --clipboard --input
    return 0
  fi
  return 1
}

key_copy_action() {
  local key="$1"
  printf '\n'
  write_step "Copy Key action"
  printf '   Press C to copy the generated key to the local clipboard.\n'
  printf '   Press S to show the key once in this terminal.\n'
  printf '   Press Enter to skip.\n'
  printf '   Choice: '
  read -r choice

  case "$choice" in
    c|C)
      if copy_to_clipboard "$key"; then
        write_ok "SECRET_KEY copied to the local clipboard"
      else
        write_warn "Clipboard copy failed because no supported clipboard tool was found."
        write_warn "On macOS this usually requires pbcopy. On Linux install wl-clipboard, xclip, or xsel."
      fi
      ;;
    s|S)
      printf '\n'
      write_warn "Generated SECRET_KEY:"
      printf '   %s\n' "$key"
      write_warn "Do not paste this key into chat, email, tickets, or logs."
      ;;
    *)
      write_warn "Skipped key copy/display."
      ;;
  esac
}

printf '\n'
printf '=============================================\n'
printf '  XLSX PII Scrubber - First-time Setup\n'
printf '=============================================\n'

write_step "Checking Python version..."
if ! find_python_cmd; then
  write_warn "Python 3.10+ was not found. Setup will try to install it now."
  install_python
  if ! find_python_cmd; then
    write_fail "Python was installed, but this terminal cannot find Python 3.10+ yet."
    printf '   Open a new terminal in this folder and run ./setup.sh again.\n'
    exit 1
  fi
fi
write_ok "$("$python_cmd" --version) found"

write_step "Installing dependencies from requirements.txt..."
"$python_cmd" -m pip install -r requirements.txt --quiet
write_ok "All packages installed"

write_step "Checking .env..."
if [ -f ".env" ]; then
  write_warn ".env already exists and will not be overwritten."
else
  if [ ! -f ".env.example" ]; then
    write_fail ".env.example was not found. Run this script from the project root folder."
    exit 1
  fi
  cp ".env.example" ".env"
  write_ok ".env created from .env.example"
fi

write_step "Checking SECRET_KEY in .env..."
key_value=""
if grep -Eq '^[[:space:]]*SECRET_KEY[[:space:]]*=' ".env"; then
  key_value="$(grep -E '^[[:space:]]*SECRET_KEY[[:space:]]*=' ".env" | tail -n 1 | sed 's/^[^=]*=//' | sed 's/^[[:space:]]*//' | sed 's/[[:space:]]*$//' | sed 's/^"//' | sed 's/"$//' | sed "s/^'//" | sed "s/'$//")"
fi

placeholder="paste_the_complete_generated_key_here"
need_new_key=0
if [ -z "$key_value" ] || [ "$key_value" = "$placeholder" ]; then
  need_new_key=1
elif ! test_secret_key "$key_value"; then
  need_new_key=1
fi

if [ "$need_new_key" -eq 1 ]; then
  if [ -n "$key_value" ] && [ "$key_value" != "$placeholder" ]; then
    write_warn "Current SECRET_KEY is not valid for this app. Creating a new 64-byte key..."
  else
    write_warn "SECRET_KEY is not set. Creating a new key..."
  fi

  new_key="$("$python_cmd" - <<'PY'
import base64
import secrets

print(base64.urlsafe_b64encode(secrets.token_bytes(64)).decode("ascii").rstrip("="))
PY
)"

  if grep -Eq '^[[:space:]]*SECRET_KEY[[:space:]]*=' ".env"; then
    tmp_file="$(mktemp)"
    sed -E "s|^[[:space:]]*SECRET_KEY[[:space:]]*=.*$|SECRET_KEY=\"$new_key\"|" ".env" > "$tmp_file"
    mv "$tmp_file" ".env"
  else
    printf '\nSECRET_KEY="%s"\n' "$new_key" >> ".env"
  fi

  write_ok "New SECRET_KEY generated and saved to .env"
  write_warn "Keep this key in a safe place. Do not change it while encrypted files still need to be restored."
  key_copy_action "$new_key"
else
  write_ok "SECRET_KEY is already available"
fi

write_step "Validating configuration..."
if [ -f "_check_env_key.py" ]; then
  "$python_cmd" _check_env_key.py
  write_ok "Configuration validation passed"
else
  write_warn "_check_env_key.py was not found, skipping validation."
fi

printf '\n'
printf '=============================================\n'
printf '  Setup complete. The app is ready to run.\n'
printf '=============================================\n'
printf '\n'
printf '  To start the development server:\n'
printf '    %s -m uvicorn app:app --host 127.0.0.1 --port 8000 --reload\n' "$python_cmd"
printf '\n'
printf '  Open this in a browser:\n'
printf '    http://127.0.0.1:8000\n'
printf '\n'
