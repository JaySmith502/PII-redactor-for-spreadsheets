import json
import re
import tempfile
import unittest
from pathlib import Path

import document_processor
from app import CSRF_COOKIE, create_app

try:
    from fastapi.testclient import TestClient
except RuntimeError:
    TestClient = None


@unittest.skipIf(document_processor.fitz is None, "PyMuPDF is not installed")
@unittest.skipIf(TestClient is None, "FastAPI TestClient dependencies are not installed")
class ReplaceUploadTest(unittest.TestCase):
    def test_pdf_upload_is_accepted_for_replace_only(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            source = root / "sample.pdf"
            runtime = root / "runtime"

            document = document_processor.fitz.open()
            page = document.new_page()
            page.insert_text((72, 72), "John Smith")
            document.save(source)
            document.close()

            app = create_app(env_file=root / "missing.env", runtime_dir=runtime)
            with TestClient(app) as client:
                home = client.get("/")
                self.assertEqual(home.status_code, 200)
                csrf = client.cookies.get(CSRF_COOKIE)
                self.assertTrue(csrf)

                with source.open("rb") as upload:
                    response = client.post(
                        "/process",
                        data={
                            "mode": "replace",
                            "csrf_token": csrf,
                            "replacement_rules": json.dumps(
                                [{"find": "John Smith", "replace": "CLIENT_001"}]
                            ),
                        },
                        files={"workbook": ("sample.pdf", upload, "application/pdf")},
                    )

                self.assertEqual(response.status_code, 200)
                self.assertIn("sample.replaced.pdf", response.text)
                self.assertIn("1 replacement match applied", response.text)

                token_match = re.search(
                    r'name="download_token" value="([^"]+)"',
                    response.text,
                )
                self.assertIsNotNone(token_match)


if __name__ == "__main__":
    unittest.main()
