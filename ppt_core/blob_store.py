"""
ppt_core/blob_store.py
──────────────────────
Thin wrapper around azure-storage-blob for the three operations
the function needs:

    download_blob(container, blob_name) -> bytes
    upload_blob(container, blob_name, data: bytes) -> None
    generate_sas_url(container, blob_name, expiry_hours) -> str

Authentication
--------------
Uses DefaultAzureCredential (Managed Identity in Azure, az login locally).
Falls back to AZURE_STORAGE_CONNECTION_STRING env-var when provided —
useful for local dev / CI where MI is unavailable.

Connection-string mode is detected automatically: if the env-var is set
it takes precedence over MI, so local testing never requires an identity.
"""

from __future__ import annotations

import logging
import os
from datetime import datetime, timedelta, timezone
from typing import Optional

from azure.identity import DefaultAzureCredential, ManagedIdentityCredential
from azure.storage.blob import (
    BlobClient,
    BlobSasPermissions,
    BlobServiceClient,
    generate_blob_sas,
)

log = logging.getLogger(__name__)

# ── Connection strategy ───────────────────────────────────────────────────────

_CONN_STR_ENV = "AZURE_STORAGE_CONNECTION_STRING"
_ACCOUNT_NAME_ENV = "STORAGE_ACCOUNT_NAME"


def _get_service_client() -> BlobServiceClient:
    """
    Build a BlobServiceClient using either:
      1. Connection string  (AZURE_STORAGE_CONNECTION_STRING env-var)
      2. Managed Identity   (DefaultAzureCredential — works in Azure)

    Raises
    ------
    EnvironmentError  if neither is configured.
    """
    conn_str = os.getenv(_CONN_STR_ENV)
    if conn_str:
        log.debug("BlobServiceClient: using connection string")
        return BlobServiceClient.from_connection_string(conn_str)

    account_name = os.getenv(_ACCOUNT_NAME_ENV)
    if not account_name:
        raise EnvironmentError(
            f"Set either {_CONN_STR_ENV} or {_ACCOUNT_NAME_ENV} environment variable."
        )

    log.debug("BlobServiceClient: using DefaultAzureCredential (account=%s)", account_name)
    credential = DefaultAzureCredential()
    return BlobServiceClient(
        account_url=f"https://{account_name}.blob.core.windows.net",
        credential=credential,
    )


# ── Public helpers ────────────────────────────────────────────────────────────

def download_blob(container: str, blob_name: str) -> bytes:
    """
    Download a blob and return its raw bytes.

    Parameters
    ----------
    container : container name
    blob_name : blob path/name within the container

    Raises
    ------
    azure.core.exceptions.ResourceNotFoundError  if blob does not exist
    """
    client = _get_service_client()
    blob_client = client.get_blob_client(container=container, blob=blob_name)
    log.info("Downloading blob: %s/%s", container, blob_name)
    return blob_client.download_blob().readall()


def upload_blob(
    container: str,
    blob_name: str,
    data: bytes,
    content_type: str = "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    overwrite: bool = True,
) -> None:
    """
    Upload bytes to a blob, creating the container if it doesn't exist.

    Parameters
    ----------
    container    : target container name
    blob_name    : target blob path/name
    data         : raw bytes to upload
    content_type : MIME type for the blob (defaults to .pptx MIME)
    overwrite    : if True, overwrite any existing blob with the same name
    """
    client = _get_service_client()
    container_client = client.get_container_client(container)

    if not container_client.exists():
        log.info("Container '%s' does not exist — creating", container)
        container_client.create_container()

    blob_client = client.get_blob_client(container=container, blob=blob_name)
    log.info(
        "Uploading blob: %s/%s (%d bytes)",
        container, blob_name, len(data),
    )
    blob_client.upload_blob(
        data,
        overwrite=overwrite,
        content_settings={"content_type": content_type},
    )


def generate_sas_url(
    container: str,
    blob_name: str,
    expiry_hours: int = 24,
) -> str:
    """
    Generate a read-only SAS URL for a blob.

    Two signing paths depending on available credentials:

    1. Connection string present (AZURE_STORAGE_CONNECTION_STRING):
       Signs with the account key parsed from the connection string.

    2. Managed Identity (STORAGE_ACCOUNT_NAME, no connection string):
       Signs with a user-delegation key obtained via get_user_delegation_key().
       Requires the "Storage Blob Delegator" role on the Managed Identity.
       Azure hard-limits user-delegation keys to 7 days; expiry_hours is
       silently capped at 168.

    Parameters
    ----------
    container    : container name
    blob_name    : blob path/name
    expiry_hours : SAS validity window in hours (capped at 168 h on MI path)

    Returns
    -------
    str  – full HTTPS URL including SAS token query string
    """
    service_client = _get_service_client()
    account_name: str = service_client.account_name  # type: ignore[assignment]

    now = datetime.now(tz=timezone.utc)

    conn_str = os.getenv(_CONN_STR_ENV)
    if conn_str:
        # ── Connection-string path: sign with account key ─────────────────────
        account_key: Optional[str] = None
        for part in conn_str.split(";"):
            if part.startswith("AccountKey="):
                account_key = part[len("AccountKey="):]
                break
        if account_key is None:
            raise ValueError(
                f"{_CONN_STR_ENV} is set but contains no AccountKey segment."
            )
        expiry = now + timedelta(hours=expiry_hours)
        sas_token = generate_blob_sas(
            account_name=account_name,
            container_name=container,
            blob_name=blob_name,
            account_key=account_key,
            permission=BlobSasPermissions(read=True),
            expiry=expiry,
        )
    else:
        # ── Managed Identity path: user-delegation SAS ────────────────────────
        # Azure caps user-delegation key validity at 7 days (168 h).
        capped_hours = min(expiry_hours, 168)
        expiry = now + timedelta(hours=capped_hours)
        delegation_key = service_client.get_user_delegation_key(
            key_start_time=now,
            key_expiry_time=expiry,
        )
        sas_token = generate_blob_sas(
            account_name=account_name,
            container_name=container,
            blob_name=blob_name,
            user_delegation_key=delegation_key,
            permission=BlobSasPermissions(read=True),
            expiry=expiry,
        )

    url = (
        f"https://{account_name}.blob.core.windows.net"
        f"/{container}/{blob_name}?{sas_token}"
    )
    log.info("SAS URL generated (expires %s UTC): %s...", expiry.isoformat(), url[:80])
    return url
