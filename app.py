"""FastAPI entry point for the server-rendered PII Scrubber web app."""

import asyncio
import base64
import hmac
import json
import logging
import re
import secrets
import shutil
import time
from contextlib import asynccontextmanager
from dataclasses import dataclass
from pathlib import Path

from fastapi import FastAPI, File, Form, Request, UploadFile, status
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse, Response
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from starlette.background import BackgroundTask
from starlette.concurrency import run_in_threadpool

try:
    from .document_processor import replace_text_document
    from .key_store import KeyConfigurationError, load_web_key
    from .processor import (
        CellCipher,
        ReplacementRule,
        ScrubberError,
        peek_workbook_columns,
        selected_column_parts,
        transform_workbook,
    )
except ImportError:
    from document_processor import replace_text_document
    from key_store import KeyConfigurationError, load_web_key
    from processor import (
        CellCipher,
        ReplacementRule,
        ScrubberError,
        peek_workbook_columns,
        selected_column_parts,
        transform_workbook,
    )


APP_DIR = Path(__file__).resolve().parent
MAX_UPLOAD_BYTES = 25 * 1024 * 1024
MAX_REQUEST_BYTES = 26 * 1024 * 1024
JOB_TTL_SECONDS = 15 * 60
CLEANUP_INTERVAL_SECONDS = 60
CSRF_COOKIE = "pii_scrubber_csrf"
LOGGER = logging.getLogger(__name__)
EXCEL_SUFFIX = ".xlsx"
REPLACE_DOCUMENT_SUFFIXES = {".xlsx", ".docx", ".pdf"}


@dataclass(frozen=True)
class DownloadJob:
    directory: Path
    output_path: Path
    download_name: str
    created_at: float


def _safe_output_name(original_name: str | None, mode: str) -> str:
    supplied = (original_name or "workbook.xlsx").replace("\\", "/").rsplit("/", 1)[-1]
    suffix = Path(supplied).suffix.casefold()
    if suffix not in REPLACE_DOCUMENT_SUFFIXES:
        suffix = ".xlsx"
    stem = Path(supplied).stem
    stem = re.sub(r"[^A-Za-z0-9 ._-]+", "_", stem).strip(" .")[:120]
    if not stem:
        stem = "document" if mode == "replace" else "workbook"
    if mode == "encrypt":
        return f"{stem}.encrypted.xlsx"
    if mode == "replace":
        return f"{stem}.replaced{suffix}"
    stem = re.sub(r"\.encrypted", "", stem, flags=re.IGNORECASE).strip(" .")
    if not stem:
        stem = "workbook"
    return f"{stem}_restored.xlsx"


def _file_suffix(filename: str | None) -> str:
    return Path(filename or "").suffix.casefold()


def _valid_upload_suffix(filename: str | None, mode: str) -> bool:
    suffix = _file_suffix(filename)
    if mode in {"encrypt", "decrypt"}:
        return suffix == EXCEL_SUFFIX
    if mode == "replace":
        return suffix in REPLACE_DOCUMENT_SUFFIXES
    return False


def _download_media_type(path: Path) -> str:
    suffix = path.suffix.casefold()
    if suffix == ".pdf":
        return "application/pdf"
    if suffix == ".docx":
        return "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    return "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def _remove_job_directory(directory: Path, runtime_root: Path) -> None:
    try:
        resolved_directory = directory.resolve()
        resolved_root = runtime_root.resolve()
    except OSError:
        return
    if resolved_directory == resolved_root or resolved_root not in resolved_directory.parents:
        return
    shutil.rmtree(resolved_directory, ignore_errors=True)


def _public_error(error: Exception, mode: str | None = None) -> str:
    message = str(error).casefold()
    if "already contains an encrypted value" in message:
        return "This workbook already contains encrypted values and cannot be encrypted again."
    if "key is wrong" in message or "ciphertext was modified" in message:
        return "The workbook could not be decrypted. The key may be different or the encrypted data was changed."
    if "select at least one column" in message:
        return "Select at least one field to encrypt."
    if "selected columns" in message:
        return "The selected fields could not be read. Upload the file again and retry."
    if "at least one replacement" in message:
        return "Add at least one find and replace rule before using Replace only."
    if "no replacement values were found" in message:
        return "No replacement values were found in the selected file."
    if "no configured values" in message or "no values were found" in message:
        if mode == "decrypt":
            return "No encrypted values were found in this workbook."
        return "No values were found in the selected fields to encrypt."
    if "replacement" in message:
        if "pymupdf" in message:
            return "PDF replacement requires the local PyMuPDF package. Install requirements and retry."
        return "Check the find and replace rules. Each rule needs a sensitive value and a safe label."
    if "docx" in message:
        return "The selected file is not a supported or readable .docx document."
    if "pdf" in message:
        return "The selected file is not a supported or readable PDF document."
    if "document" in message:
        return "The selected document could not be processed safely. No output was created."
    if "too many package members" in message or "safe processing limit" in message:
        return "This workbook is too large or complex to process safely."
    if "password-protected" in message:
        return "Password-protected Excel workbooks are not supported."
    if "xlsx" in message or "workbook" in message or "zip" in message:
        return "The selected file is not a supported or readable .xlsx workbook."
    return "The file could not be processed safely. No output was created."


def _valid_csrf(request: Request, submitted: str) -> bool:
    cookie = request.cookies.get(CSRF_COOKIE, "")
    return bool(cookie and submitted and hmac.compare_digest(cookie, submitted))


async def _write_upload(upload: UploadFile, destination: Path) -> None:
    total = 0
    with destination.open("xb") as output:
        while True:
            chunk = await upload.read(1024 * 1024)
            if not chunk:
                break
            total += len(chunk)
            if total > MAX_UPLOAD_BYTES:
                raise ScrubberError("File exceeds the 25 MiB upload limit")
            output.write(chunk)


def _parse_selected_columns(raw: str, mode: str) -> set[str] | None:
    if mode != "encrypt":
        return None
    if not raw:
        raise ScrubberError("Select at least one column to encrypt")
    try:
        decoded = json.loads(raw)
    except json.JSONDecodeError as exc:
        raise ScrubberError("Selected columns could not be read") from exc
    if not isinstance(decoded, list):
        raise ScrubberError("Selected columns could not be read")
    columns: set[str] = set()
    for item in decoded:
        if not isinstance(item, str) or selected_column_parts(item) is None:
            raise ScrubberError("Selected columns could not be read")
        columns.add(item)
    if not columns:
        raise ScrubberError("Select at least one column to encrypt")
    return columns


def _parse_replacement_rules(raw: str, mode: str) -> list[ReplacementRule]:
    if mode not in {"encrypt", "replace"}:
        return []
    if not raw:
        if mode == "replace":
            raise ScrubberError("At least one replacement rule is required")
        return []
    try:
        decoded = json.loads(raw)
    except json.JSONDecodeError as exc:
        raise ScrubberError("Find and replace rules could not be read") from exc
    if not isinstance(decoded, list) or len(decoded) > 100:
        raise ScrubberError("Find and replace rules could not be read")

    rules: list[ReplacementRule] = []
    for item in decoded:
        if not isinstance(item, dict):
            raise ScrubberError("Find and replace rules could not be read")
        raw_find = item.get("find", "")
        raw_replace = item.get("replace", "")
        if not isinstance(raw_find, str) or not isinstance(raw_replace, str):
            raise ScrubberError("Find and replace rules could not be read")
        find_value = raw_find.strip()
        replace_value = raw_replace.strip()
        if not find_value and not replace_value:
            continue
        if not find_value or len(find_value) > 500 or len(replace_value) > 500:
            raise ScrubberError("Each replacement needs a sensitive value and a safe label")
        if find_value == replace_value:
            continue
        rules.append(ReplacementRule(find=find_value, replace=replace_value))
    if mode == "replace" and not rules:
        raise ScrubberError("At least one replacement rule is required")
    return rules


async def _cleanup_expired_jobs(app: FastAPI) -> None:
    while True:
        await asyncio.sleep(CLEANUP_INTERVAL_SECONDS)
        cutoff = time.time() - JOB_TTL_SECONDS
        expired: list[DownloadJob] = []
        async with app.state.jobs_lock:
            for token, job in list(app.state.jobs.items()):
                if job.created_at < cutoff:
                    expired.append(app.state.jobs.pop(token))
        for job in expired:
            await run_in_threadpool(
                _remove_job_directory, job.directory, app.state.runtime_dir
            )


def create_app(
    env_file: Path | None = None,
    runtime_dir: Path | None = None,
) -> FastAPI:
    resolved_runtime = (runtime_dir or (APP_DIR / ".tmp" / "jobs")).resolve()
    templates = Jinja2Templates(directory=str(APP_DIR / "templates"))
    templates.env.globals["CSRF_COOKIE"] = CSRF_COOKIE

    @asynccontextmanager
    async def lifespan(application: FastAPI):
        resolved_runtime.mkdir(parents=True, exist_ok=True)
        for child in resolved_runtime.iterdir():
            if child.is_dir():
                _remove_job_directory(child, resolved_runtime)
        application.state.runtime_dir = resolved_runtime
        application.state.jobs = {}
        application.state.jobs_lock = asyncio.Lock()
        application.state.processing_slots = asyncio.Semaphore(2)
        application.state.cipher = None
        application.state.key_ready = False
        try:
            application.state.cipher = CellCipher(load_web_key(env_file))
            application.state.key_ready = True
        except KeyConfigurationError as exc:
            LOGGER.error("SECRET_KEY configuration error: %s", exc)
            application.state.key_ready = False
        except ScrubberError as exc:
            LOGGER.error("Encryption service initialization failed: %s", exc)
            application.state.key_ready = False
        cleanup_task = asyncio.create_task(_cleanup_expired_jobs(application))
        try:
            yield
        finally:
            cleanup_task.cancel()
            try:
                await cleanup_task
            except asyncio.CancelledError:
                pass
            for job in list(application.state.jobs.values()):
                _remove_job_directory(job.directory, resolved_runtime)

    application = FastAPI(
        title="PII Scrubber",
        docs_url=None,
        redoc_url=None,
        openapi_url=None,
        lifespan=lifespan,
    )
    application.mount("/static", StaticFiles(directory=str(APP_DIR / "static")), name="static")

    @application.middleware("http")
    async def security_headers(request: Request, call_next):
        response = None
        content_length = request.headers.get("content-length")
        if request.url.path in {"/process", "/peek"} and content_length:
            try:
                if int(content_length) > MAX_REQUEST_BYTES:
                    response = templates.TemplateResponse(
                        request=request,
                        name="error.html",
                        context={"message": "The selected file exceeds the 25 MiB limit."},
                        status_code=status.HTTP_413_REQUEST_ENTITY_TOO_LARGE,
                    )
            except ValueError:
                response = Response(status_code=status.HTTP_400_BAD_REQUEST)
        if response is None:
            response = await call_next(request)
        response.headers["Cache-Control"] = "no-store, max-age=0"
        response.headers["Pragma"] = "no-cache"
        response.headers["X-Content-Type-Options"] = "nosniff"
        response.headers["X-Frame-Options"] = "DENY"
        response.headers["Referrer-Policy"] = "no-referrer"
        response.headers["Content-Security-Policy"] = (
            "default-src 'self'; script-src 'self'; style-src 'self'; img-src 'self' data:; "
            "font-src 'self'; connect-src 'self'; form-action 'self'; frame-ancestors 'none'; "
            "base-uri 'self'"
        )
        return response

    @application.get("/", response_class=HTMLResponse)
    async def index(request: Request):
        csrf_token = request.cookies.get(CSRF_COOKIE) or secrets.token_urlsafe(32)
        response = templates.TemplateResponse(
            request=request,
            name="index.html",
            context={
                "csrf_token": csrf_token,
                "key_ready": request.app.state.key_ready,
                "max_upload_mb": MAX_UPLOAD_BYTES // (1024 * 1024),
            },
        )
        response.set_cookie(
            CSRF_COOKIE,
            csrf_token,
            httponly=True,
            secure=request.url.scheme == "https",
            samesite="strict",
            path="/",
        )
        return response

    @application.post("/peek", response_class=JSONResponse)
    async def peek(
        request: Request,
        workbook: UploadFile = File(...),
        csrf_token: str = Form(...),
    ):
        """Return column info for the uploaded workbook so the user can choose what to encrypt."""
        if not _valid_csrf(request, csrf_token):
            await workbook.close()
            return JSONResponse({"error": "CSRF token invalid"}, status_code=403)
        if not request.app.state.key_ready:
            await workbook.close()
            return JSONResponse({"error": "Service unavailable"}, status_code=503)
        if not workbook.filename or Path(workbook.filename).suffix.casefold() != ".xlsx":
            await workbook.close()
            return JSONResponse({"error": "Only .xlsx files are supported"}, status_code=400)

        token = secrets.token_urlsafe(32)
        peek_dir = request.app.state.runtime_dir / f"peek_{token}"
        peek_dir.mkdir(parents=False, exist_ok=False)
        input_path = peek_dir / "input.xlsx"
        original_filename = workbook.filename or "workbook.xlsx"
        try:
            await _write_upload(workbook, input_path)
            columns = await run_in_threadpool(peek_workbook_columns, input_path)
        except (ScrubberError, OSError, ValueError) as exc:
            _remove_job_directory(peek_dir, request.app.state.runtime_dir)
            return JSONResponse({"error": _public_error(exc)}, status_code=422)
        except Exception:
            _remove_job_directory(peek_dir, request.app.state.runtime_dir)
            return JSONResponse({"error": "Failed to read the file"}, status_code=500)
        finally:
            await workbook.close()

        # Store the uploaded file so /process can reuse it (keyed by peek_token)
        peek_job = DownloadJob(peek_dir, input_path, original_filename, time.time())
        async with request.app.state.jobs_lock:
            request.app.state.jobs[f"peek_{token}"] = peek_job

        return JSONResponse({"peek_token": token, "columns": columns})

    @application.post("/process", response_class=HTMLResponse)
    async def process(
        request: Request,
        workbook: UploadFile = File(None),
        mode: str = Form(...),
        csrf_token: str = Form(...),
        selected_columns: str = Form(default=""),
        replacement_rules: str = Form(default=""),
        peek_token: str = Form(default=""),
    ):
        if not _valid_csrf(request, csrf_token):
            if workbook:
                await workbook.close()
            return templates.TemplateResponse(
                request=request,
                name="error.html",
                context={"message": "The form expired. Return to the upload page and try again."},
                status_code=status.HTTP_403_FORBIDDEN,
            )
        if mode in {"encrypt", "decrypt"} and (
            not request.app.state.key_ready or request.app.state.cipher is None
        ):
            if workbook:
                await workbook.close()
            return templates.TemplateResponse(
                request=request,
                name="error.html",
                context={
                    "message": "The encryption service is temporarily unavailable. "
                    "Please contact the administrator."
                },
                status_code=status.HTTP_503_SERVICE_UNAVAILABLE,
            )
        if mode not in {"encrypt", "decrypt", "replace"}:
            if workbook:
                await workbook.close()
            return templates.TemplateResponse(
                request=request,
                name="error.html",
                context={"message": "Select Encrypt, Decrypt, or Replace only."},
                status_code=status.HTTP_400_BAD_REQUEST,
            )
        if not peek_token and workbook and workbook.filename:
            if not _valid_upload_suffix(workbook.filename, mode):
                await workbook.close()
                message = (
                    "Encrypt and decrypt require a .xlsx workbook."
                    if mode in {"encrypt", "decrypt"}
                    else "Replace only supports .xlsx, .docx, and .pdf files."
                )
                return templates.TemplateResponse(
                    request=request,
                    name="error.html",
                    context={"message": message},
                    status_code=status.HTTP_400_BAD_REQUEST,
                )

        # Resolve input: either reuse peek-cached file or accept fresh upload
        if peek_token:
            async with request.app.state.jobs_lock:
                peek_job = request.app.state.jobs.pop(f"peek_{peek_token}", None)
        else:
            peek_job = None

        if peek_token and peek_job is None and mode == "encrypt":
            if workbook:
                await workbook.close()
            return templates.TemplateResponse(
                request=request,
                name="error.html",
                context={"message": "This upload expired. Upload the file again."},
                status_code=status.HTTP_404_NOT_FOUND,
            )

        if peek_job is not None:
            # Reuse the already-uploaded file from /peek
            job_directory = peek_job.directory
            input_path = peek_job.output_path  # output_path holds the input xlsx in peek jobs
            download_name = _safe_output_name(peek_job.download_name, mode)
            token = secrets.token_urlsafe(32)
            output_path = job_directory / "output.xlsx"
            try:
                col_set = _parse_selected_columns(selected_columns, mode)
                replacements = _parse_replacement_rules(replacement_rules, mode)
                async with request.app.state.processing_slots:
                    stats = await run_in_threadpool(
                        transform_workbook,
                        input_path,
                        output_path,
                        request.app.state.cipher,
                        mode,
                        False,
                        col_set,
                        replacements,
                    )
                replacement_count = stats.replacements
            except (ScrubberError, OSError, ValueError) as exc:
                _remove_job_directory(job_directory, request.app.state.runtime_dir)
                return templates.TemplateResponse(
                    request=request,
                    name="error.html",
                    context={"message": _public_error(exc, mode)},
                    status_code=422,
                )
            except Exception:
                _remove_job_directory(job_directory, request.app.state.runtime_dir)
                return templates.TemplateResponse(
                    request=request,
                    name="error.html",
                    context={"message": "The file could not be processed safely. No output was created."},
                    status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
                )
        else:
            if workbook is None or not workbook.filename:
                return templates.TemplateResponse(
                    request=request,
                    name="error.html",
                    context={"message": "No file was provided."},
                    status_code=status.HTTP_400_BAD_REQUEST,
                )
            if not _valid_upload_suffix(workbook.filename, mode):
                await workbook.close()
                message = (
                    "Encrypt and decrypt require a .xlsx workbook."
                    if mode in {"encrypt", "decrypt"}
                    else "Replace only supports .xlsx, .docx, and .pdf files."
                )
                return templates.TemplateResponse(
                    request=request,
                    name="error.html",
                    context={"message": message},
                    status_code=status.HTTP_400_BAD_REQUEST,
                )
            token = secrets.token_urlsafe(32)
            job_directory = request.app.state.runtime_dir / token
            suffix = _file_suffix(workbook.filename)
            input_path = job_directory / f"input{suffix}"
            output_path = job_directory / f"output{suffix}"
            download_name = _safe_output_name(workbook.filename, mode)
            job_directory.mkdir(parents=False, exist_ok=False)
            try:
                await _write_upload(workbook, input_path)
                col_set = _parse_selected_columns(selected_columns, mode)
                replacements = _parse_replacement_rules(replacement_rules, mode)
                async with request.app.state.processing_slots:
                    if mode == "replace":
                        if suffix == ".xlsx":
                            stats = await run_in_threadpool(
                                transform_workbook,
                                input_path,
                                output_path,
                                None,
                                "replace",
                                False,
                                None,
                                replacements,
                            )
                            replacement_count = stats.replacements
                        else:
                            replacement_count = await run_in_threadpool(
                                replace_text_document,
                                input_path,
                                output_path,
                                replacements,
                                False,
                            )
                            stats = None
                    else:
                        stats = await run_in_threadpool(
                            transform_workbook,
                            input_path,
                            output_path,
                            request.app.state.cipher,
                            mode,
                            False,
                            col_set,
                            replacements,
                        )
                        replacement_count = stats.replacements
                input_path.unlink(missing_ok=True)
            except (ScrubberError, OSError, ValueError) as exc:
                _remove_job_directory(job_directory, request.app.state.runtime_dir)
                return templates.TemplateResponse(
                    request=request,
                    name="error.html",
                    context={"message": _public_error(exc, mode)},
                    status_code=422,
                )
            except Exception:
                _remove_job_directory(job_directory, request.app.state.runtime_dir)
                return templates.TemplateResponse(
                    request=request,
                    name="error.html",
                    context={"message": "The file could not be processed safely. No output was created."},
                    status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
                )
            finally:
                await workbook.close()

        job = DownloadJob(job_directory, output_path, download_name, time.time())
        async with request.app.state.jobs_lock:
            request.app.state.jobs[token] = job

        action = "encrypted" if mode == "encrypt" else "decrypted"
        if mode == "replace":
            action = "updated"
        
        labels_by_sheet: dict[str, list[str]] = {}
        if 'stats' in locals() and stats:
            for sheet_name, col_header in stats.encrypted_columns:
                labels_by_sheet.setdefault(sheet_name, []).append(col_header)
                
        return templates.TemplateResponse(
            request=request,
            name="result.html",
            context={
                "action": action,
                "download_name": download_name,
                "download_token": token,
                "csrf_token": csrf_token,
                "labels_by_sheet": labels_by_sheet,
                "replacement_count": replacement_count if 'replacement_count' in locals() else 0,
            },
        )

    @application.post("/secret-key", response_class=JSONResponse)
    async def secret_key(request: Request, csrf_token: str = Form(...)):
        """Return the configured SECRET_KEY so an operator can copy it from the UI.

        This deliberately exposes the key to the browser. The application has no
        login, so anyone who can reach this endpoint can read the key that
        decrypts every workbook the app has produced.
        """
        if not _valid_csrf(request, csrf_token):
            return JSONResponse(
                {"error": "The page expired. Reload the page and try again."},
                status_code=status.HTTP_403_FORBIDDEN,
            )
        try:
            key = await run_in_threadpool(load_web_key, env_file)
        except KeyConfigurationError:
            return JSONResponse(
                {"error": "SECRET_KEY is not configured on the server."},
                status_code=status.HTTP_503_SERVICE_UNAVAILABLE,
            )
        return JSONResponse(
            {"key": base64.urlsafe_b64encode(key).decode("ascii").rstrip("=")}
        )

    @application.post("/download")
    async def download(
        request: Request,
        download_token: str = Form(...),
        csrf_token: str = Form(...),
    ):
        if not _valid_csrf(request, csrf_token):
            return templates.TemplateResponse(
                request=request,
                name="error.html",
                context={"message": "The download request expired."},
                status_code=status.HTTP_403_FORBIDDEN,
            )
        async with request.app.state.jobs_lock:
            job = request.app.state.jobs.pop(download_token, None)
        if job is None or not job.output_path.is_file():
            return templates.TemplateResponse(
                request=request,
                name="error.html",
                context={"message": "This download is no longer available. Upload the file again."},
                status_code=status.HTTP_404_NOT_FOUND,
            )
        return FileResponse(
            job.output_path,
            media_type=_download_media_type(job.output_path),
            filename=job.download_name,
            background=BackgroundTask(
                _remove_job_directory, job.directory, request.app.state.runtime_dir
            ),
        )

    return application


app = create_app()


if __name__ == "__main__":
    import uvicorn

    uvicorn.run("app:app", host="127.0.0.1", port=8000, reload=False, workers=1)
