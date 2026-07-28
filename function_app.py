"""
function_app.py
───────────────
Azure Functions v2 (decorator-based) entry point — Agenda PPT Generator.

Trigger
-------
Blob Storage trigger on container  : {BLOB_INPUT_CONTAINER}  (default: input-json)
Blob name pattern                   : {name}

What it does
------------
1.  Parse and validate the dropped JSON blob.
2.  Download the PPTX template from     {BLOB_TEMPLATE_CONTAINER}/{BLOB_TEMPLATE_NAME}.
3.  Generate the populated PPTX in a    /tmp/<invocation-id>/ working directory.
4.  Upload the output PPTX to           {BLOB_OUTPUT_CONTAINER}/<stem>_generated.pptx.
5.  Write a JSON metadata sidecar to    {BLOB_OUTPUT_CONTAINER}/<stem>_metadata.json.

Environment variables
---------------------
Required in production:

    STORAGE_ACCOUNT_NAME             Storage account name (Managed Identity)
                                     OR
    AZURE_STORAGE_CONNECTION_STRING  Full connection string (local dev / CI)

Optional (sensible defaults shown):

    BLOB_INPUT_CONTAINER        [input-json]
    BLOB_TEMPLATE_CONTAINER     [templates-ppt]
    BLOB_TEMPLATE_NAME          [AgendaTemplate.pptx]
    BLOB_OUTPUT_CONTAINER       [output-agenda_ppt]
    SAS_EXPIRY_HOURS            [24]
    PPTX_SCRIPTS_DIR            [/home/site/wwwroot/scripts]
    MAX_BULLETS_PER_SLIDE       [4]
"""

from __future__ import annotations

import json
import logging
import os
import shutil
import tempfile
import time
from datetime import datetime, timezone
from pathlib import Path

import azure.functions as func

# ── Conditional import: local filesystem or Azure Blob Storage ───────────────
if os.getenv("USE_LOCAL_STORAGE", "false").lower() == "true":
    from ppt_core import blob_store_local as blob_store
else:
    from ppt_core import blob_store

from ppt_core.generator import GeneratorConfig, generate_pptx

# ─────────────────────────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  [%(name)s]  %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger(__name__)


def _env(key: str, default: str) -> str:
    return os.getenv(key, default)


def _env_int(key: str, default: int) -> int:
    try:
        return int(os.getenv(key, str(default)))
    except ValueError:
        log.warning("Invalid value for %s; using default %d", key, default)
        return default


# ─────────────────────────────────────────────────────────────────────────────
app = func.FunctionApp(http_auth_level=func.AuthLevel.FUNCTION)

from docx_core.triggers import bp as docx_bp
app.register_blueprint(docx_bp)

log.info(
    "Agenda PPT Function App initialized - Using %s",
    "LOCAL" if os.getenv("USE_LOCAL_STORAGE", "").lower() == "true" else "AZURE",
)


# ─────────────────────────────────────────────────────────────────────────────
# HTTP trigger
# ─────────────────────────────────────────────────────────────────────────────

@app.route(route="generate_agenda_ppt", methods=["POST"], auth_level=func.AuthLevel.FUNCTION)
def generate_agenda_ppt_http(req: func.HttpRequest) -> func.HttpResponse:
    """
    POST /api/generate_agenda_ppt
    Body : JSON payload (agenda schema)
    Returns: binary .pptx, or JSON error on failure.
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

    log.info("[HTTP] generate_agenda_ppt called — %d top-level keys", len(content_json) if content_json else 0)

    workdir: Path | None = None
    try:
        template_container = _env("BLOB_TEMPLATE_CONTAINER", "templates-ppt")
        template_blob_name = _env("BLOB_TEMPLATE_NAME", "AgendaTemplate.pptx")
        template_bytes = blob_store.download_blob(template_container, template_blob_name)
        log.info("[HTTP] Template loaded: %s/%s (%d bytes)", template_container, template_blob_name, len(template_bytes))

        workdir = Path(tempfile.mkdtemp(prefix="pptx_agenda_http_"))
        cfg = GeneratorConfig(
            scripts_dir=Path(_env("PPTX_SCRIPTS_DIR", "/home/site/wwwroot/scripts")),
            workdir=workdir,
            max_bullets=_env_int("MAX_BULLETS_PER_SLIDE", 4),
        )

        pptx_bytes = generate_pptx(template_bytes=template_bytes, content_json=content_json, cfg=cfg)
        elapsed_s = round(time.monotonic() - start_ts, 2)
        log.info("[HTTP] PPTX generated: %d bytes in %.2fs", len(pptx_bytes), elapsed_s)

        return func.HttpResponse(
            body=pptx_bytes,
            status_code=200,
            mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            headers={"Content-Disposition": 'attachment; filename="Agenda_Generated.pptx"'},
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
    finally:
        if workdir and workdir.exists():
            shutil.rmtree(workdir, ignore_errors=True)


# ─────────────────────────────────────────────────────────────────────────────
# Blob trigger
# ─────────────────────────────────────────────────────────────────────────────

@app.blob_trigger(
    arg_name="input_blob",
    path="{BLOB_INPUT_CONTAINER}/{name}",
    connection="AzureWebJobsStorage",
)
def generate_agenda_ppt(input_blob: func.InputStream) -> None:
    """Fires when a JSON blob lands in the input container."""
    blob_path: str = input_blob.name
    blob_name: str = blob_path.split("/", 1)[-1]
    stem: str = Path(blob_name).stem

    log.info("Triggered by blob: %s", blob_path)

    if not blob_name.lower().endswith(".json"):
        log.info("Skipping non-JSON blob: %s", blob_name)
        return

    start_ts = time.monotonic()
    workdir: Path | None = None

    try:
        raw_bytes = input_blob.read()
        try:
            content_json: dict = json.loads(raw_bytes.decode("utf-8"))
        except (json.JSONDecodeError, UnicodeDecodeError) as exc:
            log.error("Failed to parse JSON blob '%s': %s", blob_name, exc)
            _write_error_sidecar(stem, str(exc))
            return

        log.info("Parsed JSON: %d top-level keys", len(content_json))

        template_container = _env("BLOB_TEMPLATE_CONTAINER", "templates-ppt")
        template_blob_name = _env("BLOB_TEMPLATE_NAME", "AgendaTemplate.pptx")
        template_bytes = blob_store.download_blob(template_container, template_blob_name)
        log.info("Template downloaded: %s/%s (%d bytes)", template_container, template_blob_name, len(template_bytes))

        workdir = Path(tempfile.mkdtemp(prefix="pptx_agenda_fn_"))
        cfg = GeneratorConfig(
            scripts_dir=Path(_env("PPTX_SCRIPTS_DIR", "/home/site/wwwroot/scripts")),
            workdir=workdir,
            max_bullets=_env_int("MAX_BULLETS_PER_SLIDE", 4),
        )

        pptx_bytes = generate_pptx(template_bytes=template_bytes, content_json=content_json, cfg=cfg)
        log.info("PPTX generated: %d bytes", len(pptx_bytes))

        output_container = _env("BLOB_OUTPUT_CONTAINER", "output-agenda_ppt")
        output_blob_name = "Agenda_PPT.pptx"
        blob_store.upload_blob(container=output_container, blob_name=output_blob_name, data=pptx_bytes)

        sas_expiry_hours = _env_int("SAS_EXPIRY_HOURS", 24)
        sas_url = blob_store.generate_sas_url(
            container=output_container, blob_name=output_blob_name, expiry_hours=sas_expiry_hours
        )

        elapsed_s = round(time.monotonic() - start_ts, 2)
        metadata = {
            "status": "success",
            "source_blob": blob_path,
            "output_blob": f"{output_container}/{output_blob_name}",
            "sas_url": sas_url,
            "sas_expiry_hours": sas_expiry_hours,
            "generated_at": datetime.now(tz=timezone.utc).isoformat(),
            "processing_seconds": elapsed_s,
        }
        _write_metadata_sidecar(output_container, stem, metadata)
        log.info("Done in %.2fs  ->  %s/%s", elapsed_s, output_container, output_blob_name)

    except ValueError as exc:
        log.error("Validation error for '%s': %s", blob_name, exc)
        _write_error_sidecar(stem, f"ValidationError: {exc}")
    except Exception as exc:
        log.exception("Unexpected error processing '%s': %s", blob_name, exc)
        _write_error_sidecar(stem, f"UnexpectedError: {exc}")
        raise
    finally:
        if workdir and workdir.exists():
            shutil.rmtree(workdir, ignore_errors=True)


# ─────────────────────────────────────────────────────────────────────────────
# Sidecar helpers
# ─────────────────────────────────────────────────────────────────────────────

def _write_metadata_sidecar(container: str, stem: str, metadata: dict) -> None:
    sidecar_name = f"{stem}_metadata.json"
    try:
        blob_store.upload_blob(
            container=container,
            blob_name=sidecar_name,
            data=json.dumps(metadata, indent=2).encode("utf-8"),
            content_type="application/json",
        )
        log.info("Metadata sidecar written: %s/%s", container, sidecar_name)
    except Exception as exc:
        log.warning("Failed to write metadata sidecar: %s", exc)


def _write_error_sidecar(stem: str, message: str) -> None:
    output_container = _env("BLOB_OUTPUT_CONTAINER", "output-agenda_ppt")
    sidecar_name = f"{stem}_error.json"
    payload = {
        "status": "error",
        "message": message,
        "timestamp": datetime.now(tz=timezone.utc).isoformat(),
    }
    try:
        blob_store.upload_blob(
            container=output_container,
            blob_name=sidecar_name,
            data=json.dumps(payload, indent=2).encode("utf-8"),
            content_type="application/json",
        )
    except Exception as exc:
        log.warning("Failed to write error sidecar: %s", exc)
