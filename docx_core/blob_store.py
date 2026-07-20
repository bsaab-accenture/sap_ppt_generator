"""docx_core/blob_store.py
Thin Azure Blob Storage wrapper — download, upload, and SAS URL generation.

Authentication
--------------
Prefers AZURE_STORAGE_CONNECTION_STRING (local dev / CI).
Falls back to DefaultAzureCredential + STORAGE_ACCOUNT_NAME (Managed Identity).
"""

from __future__ import annotations

import logging
import os
from datetime import datetime, timedelta, timezone
from typing import Optional

from azure.identity import DefaultAzureCredential
from azure.storage.blob import (
    BlobServiceClient,
    BlobSasPermissions,
    generate_blob_sas,
)

log = logging.getLogger(__name__)

_CONN_STR_ENV = "AZURE_STORAGE_CONNECTION_STRING"
_ACCOUNT_NAME_ENV = "STORAGE_ACCOUNT_NAME"

DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"


def _get_service_client() -> BlobServiceClient:
    conn_str = os.getenv(_CONN_STR_ENV)
    if conn_str:
        return BlobServiceClient.from_connection_string(conn_str)
    account_name = os.getenv(_ACCOUNT_NAME_ENV)
    if not account_name:
        raise EnvironmentError(
            f"Set either {_CONN_STR_ENV} or {_ACCOUNT_NAME_ENV}."
        )
    return BlobServiceClient(
        account_url=f"https://{account_name}.blob.core.windows.net",
        credential=DefaultAzureCredential(),
    )


def download_blob(container: str, blob_name: str) -> bytes:
    """Download a blob and return its raw bytes."""
    client = _get_service_client()
    log.info("Downloading blob: %s/%s", container, blob_name)
    return client.get_blob_client(container=container, blob=blob_name).download_blob().readall()


def upload_blob(
    container: str,
    blob_name: str,
    data: bytes,
    content_type: str = DOCX_MIME,
    overwrite: bool = True,
) -> None:
    """Upload bytes to a blob, creating the container if it does not exist."""
    client = _get_service_client()
    cc = client.get_container_client(container)
    if not cc.exists():
        log.info("Container '%s' missing — creating", container)
        cc.create_container()
    log.info("Uploading blob: %s/%s (%d bytes)", container, blob_name, len(data))
    client.get_blob_client(container=container, blob=blob_name).upload_blob(
        data, overwrite=overwrite, content_settings={"content_type": content_type}
    )


def generate_sas_url(container: str, blob_name: str, expiry_hours: int = 24) -> str:
    """Generate a read-only SAS URL.  Caps at 168 h on Managed Identity path."""
    svc = _get_service_client()
    account_name: str = svc.account_name  # type: ignore[assignment]
    now = datetime.now(tz=timezone.utc)

    conn_str = os.getenv(_CONN_STR_ENV)
    if conn_str:
        account_key: Optional[str] = None
        for part in conn_str.split(";"):
            if part.startswith("AccountKey="):
                account_key = part[len("AccountKey="):]
                break
        if account_key is None:
            raise ValueError(f"{_CONN_STR_ENV} has no AccountKey segment.")
        expiry = now + timedelta(hours=expiry_hours)
        token = generate_blob_sas(
            account_name=account_name, container_name=container, blob_name=blob_name,
            account_key=account_key, permission=BlobSasPermissions(read=True), expiry=expiry,
        )
    else:
        expiry = now + timedelta(hours=min(expiry_hours, 168))
        dk = svc.get_user_delegation_key(key_start_time=now, key_expiry_time=expiry)
        token = generate_blob_sas(
            account_name=account_name, container_name=container, blob_name=blob_name,
            user_delegation_key=dk, permission=BlobSasPermissions(read=True), expiry=expiry,
        )

    url = f"https://{account_name}.blob.core.windows.net/{container}/{blob_name}?{token}"
    log.info("SAS URL generated (expires %s): %s...", expiry.isoformat(), url[:80])
    return url
