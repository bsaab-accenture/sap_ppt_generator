"""Local filesystem implementation of blob_store for development."""
import logging
from pathlib import Path

log = logging.getLogger(__name__)

LOCAL_STORAGE_DIR = Path("./local_storage")

def download_blob(container: str, blob_name: str) -> bytes:
    """Download from local filesystem."""
    file_path = LOCAL_STORAGE_DIR / container / blob_name
    log.info("Reading local file: %s", file_path)
    return file_path.read_bytes()

def upload_blob(container: str, blob_name: str, data: bytes, 
                content_type: str = "", overwrite: bool = True) -> None:
    """Write to local filesystem."""
    file_path = LOCAL_STORAGE_DIR / container / blob_name
    file_path.parent.mkdir(parents=True, exist_ok=True)
    log.info("Writing local file: %s", file_path)
    file_path.write_bytes(data)

def generate_sas_url(container: str, blob_name: str, expiry_hours: int) -> str:
    """Return local file path instead of SAS URL."""
    return f"file://{LOCAL_STORAGE_DIR / container / blob_name}"