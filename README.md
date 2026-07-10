# Board Meeting PPT Generator — Azure Function

Blob-triggered Azure Function (Python v2 model) that watches an Azure Storage
container for dropped JSON files, generates a fully-populated PowerPoint deck
from a template, and writes the result (plus a SAS URL) back to Storage.

---

## Architecture

```
input-json/          templates/                    output-pptx/
meeting.json  ──►  [Blob Trigger]                      │
                       │                               │
                       ├── download template ──────────┤
                       │   (PowerpointTemplate.pptx)   │
                       │                               │
                       ├── generate_pptx()             │
                       │   • validate JSON             │
                       │   • split >4-bullet slides    │
                       │   • inject title + content    │
                       │   • hide unused circles       │
                       │                               │
                       ├── upload PPTX ────────────────► meeting_generated.pptx
                       └── write metadata  ────────────► meeting_metadata.json
                                                             {
                                                               "sas_url": "...",
                                                               "status":  "success",
                                                               ...
                                                             }
```

---

## Project structure

```
├── function_app.py              Azure Function entry point (v2 decorator model)
├── host.json                    Function App runtime config
├── requirements.txt             Python dependencies
├── local.settings.json.example  Copy → local.settings.json for local dev
│
├── ppt_core/
│   ├── __init__.py
│   ├── generator.py             Pure-Python PPT generation logic (no Azure deps)
│   └── blob_store.py            Azure Blob Storage helper (download/upload/SAS)
│
├── scripts/                     pptx skill scripts (unpack/pack/add_slide/clean)
│   └── office/
│       ├── unpack.py
│       ├── pack.py
│       └── ...
│
└── infra/
    └── main.bicep               Bicep IaC (storage, function app, MI, roles)
```

---

## Deployment

### 1. Prerequisites

```bash
# Tools
az --version          # Azure CLI ≥ 2.60
func --version        # Azure Functions Core Tools ≥ 4.0
python3 --version      # Python ≥ 3.11
```

### 2. Deploy infrastructure

```bash
az group create -n rg-pptgen -l westeurope

az deployment group create \
    -g rg-pptgen \
    -f infra/main.bicep \
    -p appName=pptgen env=prod
```

Note the outputs — you'll need `storageAccountName` and `functionAppName`.

### 3. Upload the PPTX template

```bash
az storage blob upload \
    --account-name <storageAccountName> \
    --container-name templates \
    --name PowerpointTemplate_BoardMeetingGovernance.pptx \
    --file PowerpointTemplate_BoardMeetingGovernance.pptx \
    --auth-mode login
```

### 4. Deploy the Function App

```bash
# Install deps
pip install -r requirements.txt --target .python_packages/lib/site-packages

# Deploy code
func azure functionapp publish <functionAppName> --python
```

### 5. Verify

Drop a test JSON blob:

```bash
az storage blob upload \
    --account-name <storageAccountName> \
    --container-name input-json \
    --name test_meeting.json \
    --file Meeting_Minutes_Presentation_Inputs.json \
    --auth-mode login
```

Check Application Insights (or `func azure functionapp logstream`) — within
seconds you should see the function trigger and a `meeting_generated.pptx`
appear in `output-pptx`.

Read the SAS URL from the metadata sidecar:

```bash
az storage blob download \
    --account-name <storageAccountName> \
    --container-name output-pptx \
    --name test_meeting_metadata.json \
    --file - \
    --auth-mode login | python -m json.tool
```

---

## Local development

```bash
# 1. Copy and fill in settings
cp local.settings.json.example local.settings.json
# edit: set AZURE_STORAGE_CONNECTION_STRING

# 2. Install deps
pip install -r requirements.txt

# 3. Start emulator (or point at real Storage)
func start

# 4. Drop a test blob — this fires the trigger
az storage blob upload \
    --connection-string "UseDevelopmentStorage=true" \
    --container-name input-json \
    --name test.json \
    --file Meeting_Minutes_Presentation_Inputs.json
```

---

## Environment variables

| Variable | Default | Description |
|---|---|---|
| `STORAGE_ACCOUNT_NAME` | *(required)* | Storage account name (used with Managed Identity) |
| `AZURE_STORAGE_CONNECTION_STRING` | *(optional)* | Overrides MI — use for local dev / CI |
| `BLOB_INPUT_CONTAINER` | `input-json` | Container that triggers the function |
| `BLOB_TEMPLATE_CONTAINER` | `templates` | Container holding the PPTX template |
| `BLOB_TEMPLATE_NAME` | `PowerpointTemplate_BoardMeetingGovernance.pptx` | Template blob name |
| `BLOB_OUTPUT_CONTAINER` | `output-pptx` | Container for generated output |
| `SAS_EXPIRY_HOURS` | `24` | SAS URL validity window |
| `MAX_BULLETS_PER_SLIDE` | `4` | Bullets before auto-split |
| `PPTX_SCRIPTS_DIR` | `/home/site/wwwroot/scripts` | Path to pptx skill scripts |

---

## JSON input schema

```json
{
  "slides": [
    {
      "title": "Executive Summary",
      "bullets": [
        "Four main topics reviewed, none fully resolved",
        "Legal constraints created a hard blocker",
        "Operating Model validation remains pending",
        "Commercialization faces revenue uncertainties",
        "CEO pushed for decisive actions next meeting"
      ],
      "speaker_notes": "Core outcomes from the meeting."
    }
  ]
}
```

**Rules:**
- `title` — required, non-empty string
- `bullets` — required, non-empty array of strings; **>4 bullets auto-split** into "Title - Part 1", "Part 2", …
- `speaker_notes` — optional, not rendered in the deck

---

## Output — metadata sidecar

Every successful run writes `<stem>_metadata.json` to `output-pptx`:

```json
{
  "status": "success",
  "source_blob": "input-json/meeting_2026.json",
  "output_blob": "output-pptx/meeting_2026_generated.pptx",
  "sas_url": "https://stpptgenprod.blob.core.windows.net/output-pptx/meeting_2026_generated.pptx?sv=...",
  "sas_expiry_hours": 24,
  "input_slide_count": 6,
  "generated_at": "2026-06-02T10:30:00+00:00",
  "processing_seconds": 4.7
}
```

On failure, `<stem>_error.json` is written instead:

```json
{
  "status": "error",
  "message": "ValidationError: slides[1].bullets must be a non-empty array.",
  "timestamp": "2026-06-02T10:30:01+00:00"
}
```

---

## Retry policy

`host.json` configures exponential backoff (3 retries, 5 s → 2 min).
Unexpected errors (`Exception`) are re-raised so the runtime retries.
Validation errors (`ValueError`) are swallowed — no point retrying a
malformed payload.

---

## Security notes

- **No connection strings in app settings** in production — Managed Identity
  is used for all Storage access.
- **SAS URLs** are read-only and expire after `SAS_EXPIRY_HOURS`.
- **Public blob access** is disabled on the Storage Account.
- TLS 1.2 minimum is enforced on both the Storage Account and Function App.
- For MI-only SAS signing, grant the **Storage Blob Delegator** role and
  switch to the user-delegation SAS code path (see `blob_store.py` NOTE).
