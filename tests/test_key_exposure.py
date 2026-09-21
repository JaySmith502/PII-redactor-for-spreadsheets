"""The copy-key feature hands the master decryption key to a browser.

The web app has no login, so that endpoint is off unless an operator opts in.
These tests pin the safe default so a client deployment cannot enable it by
accident.
"""

import base64
import os
import secrets
import tempfile
import unittest
from pathlib import Path
from unittest import mock

from app import CSRF_COOKIE, KEY_EXPOSURE_SETTING, create_app

try:
    from fastapi.testclient import TestClient
except RuntimeError:
    TestClient = None


def _write_env(path: Path) -> str:
    key = base64.urlsafe_b64encode(secrets.token_bytes(64)).decode("ascii").rstrip("=")
    path.write_text(f'SECRET_KEY="{key}"\n', encoding="utf-8")
    return key


@unittest.skipIf(TestClient is None, "FastAPI TestClient dependencies are not installed")
class KeyExposureTest(unittest.TestCase):
    def _client(self, root: Path, **env: str):
        app = create_app(env_file=root / ".env", runtime_dir=root / "runtime")
        return TestClient(app)

    def test_copy_key_is_disabled_by_default(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            _write_env(root / ".env")
            with mock.patch.dict(os.environ, {}, clear=False):
                os.environ.pop(KEY_EXPOSURE_SETTING, None)
                with self._client(root) as client:
                    home = client.get("/")
                    self.assertEqual(home.status_code, 200)
                    self.assertNotIn("copy-key-button", home.text)

                    response = client.post(
                        "/secret-key",
                        data={"csrf_token": client.cookies.get(CSRF_COOKIE)},
                    )
                    self.assertEqual(response.status_code, 404)

    def test_copy_key_works_when_operator_opts_in(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            key = _write_env(root / ".env")
            with mock.patch.dict(os.environ, {KEY_EXPOSURE_SETTING: "1"}):
                with self._client(root) as client:
                    home = client.get("/")
                    self.assertIn("copy-key-button", home.text)

                    response = client.post(
                        "/secret-key",
                        data={"csrf_token": client.cookies.get(CSRF_COOKIE)},
                    )
                    self.assertEqual(response.status_code, 200)
                    self.assertEqual(response.json()["key"], key)

    def test_flag_in_env_file_also_enables_copy_key(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            key = _write_env(root / ".env")
            with (root / ".env").open("a", encoding="utf-8") as handle:
                handle.write(f"{KEY_EXPOSURE_SETTING}=true\n")
            with mock.patch.dict(os.environ, {}, clear=False):
                os.environ.pop(KEY_EXPOSURE_SETTING, None)
                with self._client(root) as client:
                    self.assertIn("copy-key-button", client.get("/").text)
                    response = client.post(
                        "/secret-key",
                        data={"csrf_token": client.cookies.get(CSRF_COOKIE)},
                    )
                    self.assertEqual(response.json()["key"], key)

    def test_opt_in_still_requires_csrf(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            _write_env(root / ".env")
            with mock.patch.dict(os.environ, {KEY_EXPOSURE_SETTING: "yes"}):
                with self._client(root) as client:
                    response = client.post("/secret-key", data={"csrf_token": "wrong"})
                    self.assertEqual(response.status_code, 403)


if __name__ == "__main__":
    unittest.main()
