"""
function_app.py
───────────────
Azure Functions v2 (decorator-based) entry point.

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
5.  Write a JSON metadata sidecar to    {BLOB_OUTPUT_CONTAINER}/<stem>_metadata.json
    containing the SAS URL, blob name, slide count, and processing duration.

Environment variables
---------------------
Required in production (set via Function App Configuration or Key Vault ref):

    STORAGE_ACCOUNT_NAME        Storage account name (used with Managed Identity)
                                OR
    AZURE_STORAGE_CONNECTION_STRING  Full connection string (local dev / CI)

Optional (all have sensible defaults):

    BLOB_INPUT_CONTAINER        Container that receives JSON files      [input-json]
    BLOB_TEMPLATE_CONTAINER     Container holding the PPTX template     [templates]
    BLOB_TEMPLATE_NAME          Template blob name                      [PowerpointTemplate_BoardMeetingGovernance.pptx]
    BLOB_OUTPUT_CONTAINER       Container for generated PPTX output     [output-pptx]
    SAS_EXPIRY_HOURS            SAS URL validity window in hours        [24]
    PPTX_SCRIPTS_DIR            Absolute path to pptx skill scripts     [/home/site/wwwroot/scripts]
    MAX_BULLETS_PER_SLIDE       Max bullets before slide splits         [4]

Local testing (func start)
--------------------------
    cp local.settings.json.example local.settings.json
    # fill in your storage details
    func start

To drop a test blob:
    az storage blob upload \
        --account-name <acct> \
        --container-name input-json \
        --name test_meeting.json \
        --file Meeting_Minutes_Presentation_Inputs.json
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

from ppt_core.blob_store import download_blob, generate_sas_url, upload_blob
from ppt_core.generator import GeneratorConfig, TemplateGeometry, TemplateSlotMap, generate_pptx

# ─────────────────────────────────────────────────────────────────────────────
# Logging  — Azure Functions runtime surfaces these in Application Insights
# ─────────────────────────────────────────────────────────────────────────────

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  [%(name)s]  %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger(__name__)


# ─────────────────────────────────────────────────────────────────────────────
# Config helpers
# ─────────────────────────────────────────────────────────────────────────────

def _env(key: str, default: str) -> str:
    return os.getenv(key, default)


def _env_int(key: str, default: int) -> int:
    try:
        return int(os.getenv(key, str(default)))
    except ValueError:
        log.warning("Invalid value for %s; using default %d", key, default)
        return default


# ─────────────────────────────────────────────────────────────────────────────
# Function App
# ─────────────────────────────────────────────────────────────────────────────

app = func.FunctionApp(http_auth_level=func.AuthLevel.FUNCTION)


@app.blob_trigger(
    arg_name="input_blob",
    path="{BLOB_INPUT_CONTAINER}/{name}",          # resolved at runtime by the runtime
    connection="AzureWebJobsStorage",
)
def generate_board_meeting_ppt(input_blob: func.InputStream) -> None:
    """
    Blob-triggered function.

    Fires whenever a new blob lands in the input container.
    Only processes blobs with a .json extension; silently skips all others.

    Parameters
    ----------
    input_blob : azure.functions.InputStream
        The trigger blob stream.  Properties used:
            .name   – full blob path (container/blob-name)
            .read() – raw bytes
    """
    blob_path: str = input_blob.name          # e.g. "input-json/meeting_2026.json"
    blob_name: str = blob_path.split("/", 1)[-1]   # e.g. "meeting_2026.json"
    stem: str = Path(blob_name).stem               # e.g. "meeting_2026"

    log.info("▶  Triggered by blob: %s", blob_path)

    # ── Guard: only process .json blobs ───────────────────────────────────────
    if not blob_name.lower().endswith(".json"):
        log.info("Skipping non-JSON blob: %s", blob_name)
        return

    start_ts = time.monotonic()
    workdir: Path | None = None

    try:
        # ── 1. Read & parse JSON ──────────────────────────────────────────────
        raw_bytes = input_blob.read()
        try:
            content_json: dict = json.loads(raw_bytes.decode("utf-8"))
        except (json.JSONDecodeError, UnicodeDecodeError) as exc:
            log.error("Failed to parse JSON blob '%s': %s", blob_name, exc)
            _write_error_sidecar(stem, str(exc))
            return

        log.info("Parsed JSON: %d top-level keys", len(content_json))

        # ── 2. Download template ──────────────────────────────────────────────
        template_container = _env("BLOB_TEMPLATE_CONTAINER", "templates")
        template_blob_name = _env(
            "BLOB_TEMPLATE_NAME",
            "PowerpointTemplate_BoardMeetingGovernance.pptx",
        )
        template_bytes = download_blob(template_container, template_blob_name)
        log.info(
            "Template downloaded: %s/%s (%d bytes)",
            template_container, template_blob_name, len(template_bytes),
        )

        # ── 3. Generate PPTX ──────────────────────────────────────────────────
        workdir = Path(tempfile.mkdtemp(prefix="pptx_fn_"))

        scripts_dir = Path(
            _env("PPTX_SCRIPTS_DIR", "/home/site/wwwroot/scripts")
        )

        cfg = GeneratorConfig(
            scripts_dir=scripts_dir,
            workdir=workdir,
            max_bullets=_env_int("MAX_BULLETS_PER_SLIDE", 4),
            geometry=TemplateGeometry(),
            slot_map=TemplateSlotMap(),
        )

        pptx_bytes = generate_pptx(
            template_bytes=template_bytes,
            content_json=content_json,
            cfg=cfg,
        )
        log.info("PPTX generated: %d bytes", len(pptx_bytes))

        # ── 4. Upload output PPTX ─────────────────────────────────────────────
        output_container = _env("BLOB_OUTPUT_CONTAINER", "output-pptx")
        output_blob_name = f"{stem}_generated.pptx"

        upload_blob(
            container=output_container,
            blob_name=output_blob_name,
            data=pptx_bytes,
        )

        # ── 5. Generate SAS URL ───────────────────────────────────────────────
        sas_expiry_hours = _env_int("SAS_EXPIRY_HOURS", 24)
        sas_url = generate_sas_url(
            container=output_container,
            blob_name=output_blob_name,
            expiry_hours=sas_expiry_hours,
        )

        # ── 6. Write metadata sidecar ─────────────────────────────────────────
        elapsed_s = round(time.monotonic() - start_ts, 2)
        slide_count = len(content_json.get("slides", []))

        metadata = {
            "status": "success",
            "source_blob": blob_path,
            "output_blob": f"{output_container}/{output_blob_name}",
            "sas_url": sas_url,
            "sas_expiry_hours": sas_expiry_hours,
            "input_slide_count": slide_count,
            "generated_at": datetime.now(tz=timezone.utc).isoformat(),
            "processing_seconds": elapsed_s,
        }

        _write_metadata_sidecar(output_container, stem, metadata)
        log.info(
            "✅  Done in %.2fs  →  %s/%s",
            elapsed_s, output_container, output_blob_name,
        )

    except ValueError as exc:
        # Schema / validation errors — these are caller errors, not infra errors
        log.error("Validation error for '%s': %s", blob_name, exc)
        _write_error_sidecar(stem, f"ValidationError: {exc}")

    except Exception as exc:
        # Unexpected errors — log full traceback for App Insights
        log.exception("Unexpected error processing '%s': %s", blob_name, exc)
        _write_error_sidecar(stem, f"UnexpectedError: {exc}")
        raise   # re-raise so Azure retries (if retry policy is configured)

    finally:
        # Always clean up the temp workdir to avoid disk pressure on Premium plan
        if workdir and workdir.exists():
            shutil.rmtree(workdir, ignore_errors=True)
            log.debug("Workdir cleaned: %s", workdir)


# ─────────────────────────────────────────────────────────────────────────────
# Sidecar helpers
# ─────────────────────────────────────────────────────────────────────────────

def _write_metadata_sidecar(container: str, stem: str, metadata: dict) -> None:
    """Upload a JSON metadata file alongside the generated PPTX."""
    sidecar_name = f"{stem}_metadata.json"
    try:
        upload_blob(
            container=container,
            blob_name=sidecar_name,
            data=json.dumps(metadata, indent=2).encode("utf-8"),
            content_type="application/json",
        )
        log.info("Metadata sidecar written: %s/%s", container, sidecar_name)
    except Exception as exc:
        # Non-fatal — don't let sidecar failure mask the main result
        log.warning("Failed to write metadata sidecar: %s", exc)


def _write_error_sidecar(stem: str, message: str) -> None:
    """Write an error JSON sidecar to the output container for observability."""
    output_container = _env("BLOB_OUTPUT_CONTAINER", "output-pptx")
    sidecar_name = f"{stem}_error.json"
    payload = {
        "status": "error",
        "message": message,
        "timestamp": datetime.now(tz=timezone.utc).isoformat(),
    }
    try:
        upload_blob(
            container=output_container,
            blob_name=sidecar_name,
            data=json.dumps(payload, indent=2).encode("utf-8"),
            content_type="application/json",
        )
    except Exception as exc:
        log.warning("Failed to write error sidecar: %s", exc)
