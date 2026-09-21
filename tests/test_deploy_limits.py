"""Regression tests for the upload and job-store limits.

``MAX_PENDING_JOBS`` bounds how many uploaded files may sit in the runtime
directory waiting for a download. It only protects a client deployment if both
endpoints that can park a file honour it, so these tests pin the behaviour for
``/peek`` and ``/process`` and check that refused work leaves no residue behind.
"""

import base64
import json
import secrets
import tempfile
import unittest
import zipfile
from pathlib import Path
from unittest import mock

import app as app_module
from app import CSRF_COOKIE, create_app

try:
    from fastapi.testclient import TestClient
except RuntimeError:
    TestClient = None

MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
ENCRYPT_FIELDS = json.dumps(["0|A"])


def _write_env(path: Path) -> None:
    key = base64.urlsafe_b64encode(secrets.token_bytes(64)).decode("ascii").rstrip("=")
    path.write_text(f'SECRET_KEY="{key}"\n', encoding="utf-8")


def _minimal_xlsx(path: Path) -> bytes:
    """Create a one-sheet workbook with a header row and one data row."""
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        archive.writestr(
            "[Content_Types].xml",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
            '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
            "</Types>",
        )
        archive.writestr(
            "xl/workbook.xml",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<workbook xmlns="{MAIN_NS}" xmlns:r="{REL_NS}">'
            '<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>',
        )
        archive.writestr(
            "xl/_rels/workbook.xml.rels",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            f'<Relationship Id="rId1" Type="{REL_NS}/worksheet" Target="worksheets/sheet1.xml"/>'
            "</Relationships>",
        )
        archive.writestr(
            "xl/worksheets/sheet1.xml",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<worksheet xmlns="{MAIN_NS}"><sheetData>'
            '<row r="1"><c r="A1" t="inlineStr"><is><t>Full Name</t></is></c></row>'
            '<row r="2"><c r="A2" t="inlineStr"><is><t>Ada Lovelace</t></is></c></row>'
            "</sheetData></worksheet>",
        )
    return path.read_bytes()


@unittest.skipIf(TestClient is None, "FastAPI TestClient dependencies are not installed")
class JobStoreLimitTest(unittest.TestCase):
    def _workspace(self, root: Path) -> bytes:
        _write_env(root / ".env")
        return _minimal_xlsx(root / "book.xlsx")

    @staticmethod
    def _upload(payload: bytes) -> dict:
        return {
            "files": {
                "workbook": (
                    "book.xlsx",
                    payload,
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            }
        }

    def test_full_job_store_refuses_both_endpoints_without_leaving_files(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            payload = self._workspace(root)
            runtime = root / "runtime"

            with mock.patch.object(app_module, "MAX_PENDING_JOBS", 0):
                application = create_app(env_file=root / ".env", runtime_dir=runtime)
                with TestClient(application) as client:
                    self.assertEqual(client.get("/").status_code, 200)
                    csrf = client.cookies.get(CSRF_COOKIE)

                    peek = client.post(
                        "/peek", data={"csrf_token": csrf}, **self._upload(payload)
                    )
                    self.assertEqual(peek.status_code, 503)

                    process = client.post(
                        "/process",
                        data={
                            "mode": "encrypt",
                            "csrf_token": csrf,
                            "selected_columns": ENCRYPT_FIELDS,
                            "replacement_rules": "",
                            "peek_token": "",
                        },
                        **self._upload(payload),
                    )
                    self.assertEqual(process.status_code, 503)
                    self.assertEqual(application.state.jobs, {})

            self.assertEqual(list(runtime.iterdir()), [])

    def test_claimed_and_downloaded_jobs_release_their_slot(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            payload = self._workspace(root)
            runtime = root / "runtime"

            with mock.patch.object(app_module, "MAX_PENDING_JOBS", 1):
                application = create_app(env_file=root / ".env", runtime_dir=runtime)
                with TestClient(application) as client:
                    self.assertEqual(client.get("/").status_code, 200)
                    csrf = client.cookies.get(CSRF_COOKIE)

                    first = client.post(
                        "/peek", data={"csrf_token": csrf}, **self._upload(payload)
                    )
                    self.assertEqual(first.status_code, 200)
                    peek_token = first.json()["peek_token"]
                    self.assertEqual(len(application.state.jobs), 1)

                    # The single slot is taken, so both entry points refuse.
                    second = client.post(
                        "/peek", data={"csrf_token": csrf}, **self._upload(payload)
                    )
                    self.assertEqual(second.status_code, 503)
                    refused = client.post(
                        "/process",
                        data={
                            "mode": "encrypt",
                            "csrf_token": csrf,
                            "selected_columns": ENCRYPT_FIELDS,
                            "replacement_rules": "",
                            "peek_token": "",
                        },
                        **self._upload(payload),
                    )
                    self.assertEqual(refused.status_code, 503)
                    self.assertEqual(len(application.state.jobs), 1)

                    # Claiming the parked peek upload reuses its own slot.
                    claimed = client.post(
                        "/process",
                        data={
                            "mode": "encrypt",
                            "csrf_token": csrf,
                            "selected_columns": ENCRYPT_FIELDS,
                            "replacement_rules": "",
                            "peek_token": peek_token,
                        },
                    )
                    self.assertEqual(claimed.status_code, 200)
                    self.assertEqual(len(application.state.jobs), 1)

                    download_token = next(iter(application.state.jobs))
                    downloaded = client.post(
                        "/download",
                        data={"download_token": download_token, "csrf_token": csrf},
                    )
                    self.assertEqual(downloaded.status_code, 200)
                    self.assertEqual(application.state.jobs, {})

                    # The slot is free again, so a fresh upload is accepted.
                    fresh = client.post(
                        "/process",
                        data={
                            "mode": "encrypt",
                            "csrf_token": csrf,
                            "selected_columns": ENCRYPT_FIELDS,
                            "replacement_rules": "",
                            "peek_token": "",
                        },
                        **self._upload(payload),
                    )
                    self.assertEqual(fresh.status_code, 200)
                    self.assertEqual(len(application.state.jobs), 1)


if __name__ == "__main__":
    unittest.main()
