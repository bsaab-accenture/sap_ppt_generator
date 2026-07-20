"""docx_core/blob_store_local.py
Local filesystem shim — mirrors the blob_store interface for offline testing.
Maps container names directly to sub-folders under ./local_storage/.
"""

import logging
from pathlib import Path

log = logging.getLogger(__name__)

LOCAL_STORAGE_DIR = Path("./local_storage")


def download_blob(container: str, blob_name: str) -> bytes:
    path = LOCAL_STORAGE_DIR / container / blob_name
    log.info("Reading local file: %s", path)
    return path.read_bytes()


def upload_blob(container: str, blob_name: str, data: bytes,
                content_type: str = "", overwrite: bool = True) -> None:
    path = LOCAL_STORAGE_DIR / container / blob_name
    path.parent.mkdir(parents=True, exist_ok=True)
    log.info("Writing local file: %s", path)
    path.write_bytes(data)


def generate_sas_url(container: str, blob_name: str, expiry_hours: int = 24) -> str:
    return f"file://{LOCAL_STORAGE_DIR / container / blob_name}"
