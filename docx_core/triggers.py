"""docx_core/triggers.py
Azure Functions v2 Blueprint — all DOCX trigger definitions.

Registered by function_app.py via app.register_functions(bp).
Blueprint pattern keeps triggers self-contained inside docx_core/ so the
file is brand-new on any merge and carries zero conflict risk.

Triggers
--------
POST  /api/generate_docx          HTTP — synchronous JSON → .docx
Blob  {BLOB_DOCX_INPUT_CONTAINER} Blob — async pipeline (upload → generate → upload → SAS)

Environment variables
---------------------
Required (production):
    STORAGE_ACCOUNT_NAME              Storage account (Managed Identity)
    OR AZURE_STORAGE_CONNECTION_STRING  Full connection string (local / CI)

Optional (all have sensible defaults):
    BLOB_DOCX_INPUT_CONTAINER     JSON inputs landing container   [input-docx-json]
    BLOB_DOCX_TEMPLATE_CONTAINER  DOCX template container         [templates-docx]
    BLOB_DOCX_TEMPLATE_NAME       Template blob name              [Preliminary_Minutes_BM20260219.docx]
    BLOB_DOCX_OUTPUT_CONTAINER    Generated DOCX output container [output-docx]
    SAS_EXPIRY_HOURS              SAS URL validity in hours       [24]
    USE_LOCAL_STORAGE             "true" → local filesystem shim  [false]
"""

from __future__ import annotations

import json
import logging
import os
import time
from datetime import datetime, timezone
from pathlib import Path

import azure.functions as func

from docx_core.generator import DocxConfig, generate_docx

# ── Storage backend: local filesystem or Azure Blob ──────────────────────────
if os.getenv("USE_LOCAL_STORAGE", "false").lower() == "true":
    from docx_core import blob_store_local as blob_store
    logging.getLogger(__name__).info("Storage: LOCAL file system")
else:
    from docx_core import blob_store  # type: ignore[no-redef]
    logging.getLogger(__name__).info("Storage: AZURE blob storage")

log = logging.getLogger(__name__)

bp = func.Blueprint()

DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"


# ── Config helpers ────────────────────────────────────────────────────────────

def _env(key: str, default: str) -> str:
    return os.getenv(key, default)


def _env_int(key: str, default: int) -> int:
    try:
        return int(os.getenv(key, str(default)))
    except ValueError:
        log.warning("Invalid value for %s; using default %d", key, default)
        return default


# ── Sidecar helpers ───────────────────────────────────────────────────────────

def _write_metadata_sidecar(container: str, stem: str, metadata: dict) -> None:
    sidecar = f"{stem}_metadata.json"
    try:
        blob_store.upload_blob(
            container=container, blob_name=sidecar,
            data=json.dumps(metadata, indent=2).encode("utf-8"),
            content_type="application/json",
        )
        log.info("Metadata sidecar: %s/%s", container, sidecar)
    except Exception as exc:
        log.warning("Failed to write metadata sidecar: %s", exc)


def _write_error_sidecar(stem: str, message: str) -> None:
    container = _env("BLOB_DOCX_OUTPUT_CONTAINER", "output-docx")
    sidecar = f"{stem}_error.json"
    payload = {
        "status": "error",
        "message": message,
        "timestamp": datetime.now(tz=timezone.utc).isoformat(),
    }
    try:
        blob_store.upload_blob(
            container=container, blob_name=sidecar,
            data=json.dumps(payload, indent=2).encode("utf-8"),
            content_type="application/json",
        )
    except Exception as exc:
        log.warning("Failed to write error sidecar: %s", exc)


# ── HTTP trigger ──────────────────────────────────────────────────────────────

@bp.route(route="generate_docx", methods=["POST"], auth_level=func.AuthLevel.FUNCTION)
def generate_docx_http(req: func.HttpRequest) -> func.HttpResponse:
    """
    POST /api/generate_docx

    Body    : agendadocx_input.json payload — ``{"body": { ... }}``
    Returns : binary .docx file, or JSON error on failure.

    Designed for Power Automate HTTP connector:
        Method  POST
        Headers Content-Type: application/json
                x-functions-key: <key>
        Body    { "body": { "Meeting_Title": "...", ... } }
    """
    start_ts = time.monotonic()

    try:
        content_json: dict = req.get_json()
    except ValueError:
        return func.HttpResponse(
            json.dumps({"error": "ValidationError", "message": "Request body is not valid JSON."}),
            status_code=400,
            mimetype="application/json",
        )

    log.info("[HTTP] generate_docx — %d top-level keys", len(content_json) if content_json else 0)

    try:
        template_container = _env("BLOB_DOCX_TEMPLATE_CONTAINER", "templates-docx")
        template_name = _env("BLOB_DOCX_TEMPLATE_NAME", "Preliminary_Minutes_BM20260219.docx")
        template_bytes = blob_store.download_blob(template_container, template_name)
        log.info("[HTTP] Template: %s/%s (%d bytes)", template_container, template_name, len(template_bytes))

        docx_bytes = generate_docx(
            template_bytes=template_bytes,
            content_json=content_json,
            cfg=DocxConfig(),
        )
        elapsed = round(time.monotonic() - start_ts, 2)
        log.info("[HTTP] DOCX generated: %d bytes in %.2fs", len(docx_bytes), elapsed)

        return func.HttpResponse(
            body=docx_bytes,
            status_code=200,
            mimetype=DOCX_MIME,
            headers={"Content-Disposition": 'attachment; filename="AgendaMinutes_Generated.docx"'},
        )

    except ValueError as exc:
        log.warning("[HTTP] Validation error: %s", exc)
        return func.HttpResponse(
            json.dumps({"error": "ValidationError", "message": str(exc)}),
            status_code=400,
            mimetype="application/json",
        )

    except Exception as exc:
        log.exception("[HTTP] Unexpected error: %s", exc)
        return func.HttpResponse(
            json.dumps({"error": "InternalError", "message": str(exc)}),
            status_code=500,
            mimetype="application/json",
        )


# ── Blob trigger ──────────────────────────────────────────────────────────────

@bp.blob_trigger(
    arg_name="input_blob",
    path="{BLOB_DOCX_INPUT_CONTAINER}/{name}",
    connection="AzureWebJobsStorage",
)
def generate_agenda_docx(input_blob: func.InputStream) -> None:
    """
    Fires on new blobs in {BLOB_DOCX_INPUT_CONTAINER}.
    Skips non-.json blobs silently.

    Pipeline:  parse JSON → download template → generate DOCX →
               upload output → generate SAS URL → write metadata sidecar.
    """
    blob_path: str = input_blob.name
    blob_name: str = blob_path.split("/", 1)[-1]
    stem: str = Path(blob_name).stem

    log.info("▶  [DOCX] Triggered: %s", blob_path)

    if not blob_name.lower().endswith(".json"):
        log.info("[DOCX] Skipping non-JSON blob: %s", blob_name)
        return

    start_ts = time.monotonic()

    try:
        # 1. Parse JSON
        try:
            content_json: dict = json.loads(input_blob.read().decode("utf-8"))
        except (json.JSONDecodeError, UnicodeDecodeError) as exc:
            log.error("[DOCX] JSON parse error '%s': %s", blob_name, exc)
            _write_error_sidecar(stem, str(exc))
            return

        log.info("[DOCX] Parsed JSON: %d top-level keys", len(content_json))

        # 2. Download template
        template_container = _env("BLOB_DOCX_TEMPLATE_CONTAINER", "templates-docx")
        template_name = _env("BLOB_DOCX_TEMPLATE_NAME", "Preliminary_Minutes_BM20260219.docx")
        template_bytes = blob_store.download_blob(template_container, template_name)
        log.info("[DOCX] Template: %s/%s (%d bytes)", template_container, template_name, len(template_bytes))

        # 3. Generate DOCX
        docx_bytes = generate_docx(
            template_bytes=template_bytes,
            content_json=content_json,
            cfg=DocxConfig(),
        )
        log.info("[DOCX] Generated: %d bytes", len(docx_bytes))

        # 4. Upload output
        output_container = _env("BLOB_DOCX_OUTPUT_CONTAINER", "output-docx")
        output_blob = f"{stem}_generated.docx"
        blob_store.upload_blob(
            container=output_container, blob_name=output_blob,
            data=docx_bytes, content_type=DOCX_MIME,
        )

        # 5. SAS URL
        sas_expiry = _env_int("SAS_EXPIRY_HOURS", 24)
        sas_url = blob_store.generate_sas_url(
            container=output_container, blob_name=output_blob, expiry_hours=sas_expiry,
        )

        # 6. Metadata sidecar
        elapsed = round(time.monotonic() - start_ts, 2)
        topic_count = len(content_json.get("body", {}).get("Topics_Discussed", []))
        _write_metadata_sidecar(output_container, stem, {
            "status": "success",
            "source_blob": blob_path,
            "output_blob": f"{output_container}/{output_blob}",
            "sas_url": sas_url,
            "sas_expiry_hours": sas_expiry,
            "input_topic_count": topic_count,
            "generated_at": datetime.now(tz=timezone.utc).isoformat(),
            "processing_seconds": elapsed,
        })
        log.info("✅  [DOCX] Done in %.2fs → %s/%s", elapsed, output_container, output_blob)

    except ValueError as exc:
        log.error("[DOCX] Validation error '%s': %s", blob_name, exc)
        _write_error_sidecar(stem, f"ValidationError: {exc}")

    except Exception as exc:
        log.exception("[DOCX] Unexpected error '%s': %s", blob_name, exc)
        _write_error_sidecar(stem, f"UnexpectedError: {exc}")
        raise
