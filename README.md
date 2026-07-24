# Board Meeting PPT Generator — Azure Function

Azure Function (Python v2 model) with two triggers:

| Trigger | Use case |
|---|---|
| **HTTP POST `/api/generate_ppt`** | Primary path for Power Automate — synchronous, returns binary PPTX in the response body |
| **Blob trigger on `input-json/`** | Batch / async path — drop a JSON file into Storage and the PPTX is written to `output-pptx/` with a metadata sidecar |

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

## HTTP trigger — Power Automate integration

### Endpoint

```
POST https://<function-app-name>.azurewebsites.net/api/generate_ppt
```

### Request

```
Content-Type: application/json
x-functions-key: <function-api-key>
```

Body schema:

```json
{
  "slides": [
    {
      "title": "string (required)",
      "bullets": ["string", "..."],
      "speaker_notes": "string (optional)"
    }
  ]
}
```

Rules: `title` non-empty; `bullets` non-empty array; >4 bullets auto-split into "Title - Part 1", "Part 2", …

### Responses

| Status | Body | When |
|---|---|---|
| 200 | Binary `.pptx` (`Content-Type: application/vnd.openxmlformats-officedocument.presentationml.presentation`) | Success |
| 400 | `{"error": "ValidationError", "message": "..."}` | Bad payload |
| 500 | `{"error": "InternalError", "message": "..."}` | Unexpected failure |

### curl test

```bash
curl -X POST https://<func>.azurewebsites.net/api/generate_ppt \
  -H "Content-Type: application/json" \
  -H "x-functions-key: <key>" \
  -d @local_storage/input-json/Meeting_Minutes_Presentation_Inputs.json \
  --output BoardMeeting_Generated.pptx \
  -w "HTTP %{http_code} -- %{size_download} bytes\n"

# Verify output is a valid PPTX
python -c "
import zipfile, re
z = zipfile.ZipFile('BoardMeeting_Generated.pptx')
prs = z.read('ppt/presentation.xml').decode()
slides = re.findall(r'<p:sldId[^>]*r:id=\"(rId\d+)\"', prs)
print(f'Valid PPTX -- {len(slides)} ordered slides')
"
```

### Power Automate integration (3-step)

**Step 1 — HTTP action** (calls the function):

| Setting | Value |
|---|---|
| Method | `POST` |
| URI | `https://<func>.azurewebsites.net/api/generate_ppt` |
| Headers | `Content-Type: application/json`, `x-functions-key: <key>` |
| Body | JSON payload built from Dataverse approved minutes data |

**Step 2 — SharePoint: Create file** (saves the PPTX):

| Setting | Value |
|---|---|
| Site Address | Your SharePoint site |
| Folder Path | `/Board Meeting Minutes/@{formatDateTime(utcNow(), 'yyyy-MM-dd')}_@{triggerBody()?['MeetingName']}` |
| File Name | `Meeting_Minutes_@{formatDateTime(utcNow(), 'yyyyMMdd_HHmmss')}.pptx` |
| File Content | `@{body('HTTP')}` |

**Step 3 — Error handler** (run after HTTP action on failure):

```
Configure run after: has failed / has timed out
Action: Send email / Teams notification
Body: "PPT generation failed: @{body('HTTP')['message']}"
```

Retrieve the function key: Azure Portal > Function App > Functions > `generate_ppt` > Function Keys.

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
| `BLOB_TEMPLATE_CONTAINER` | `templates-ppt` | Container holding the PPTX template |
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
- For MI-only SAS signing, grant the **Storage Blob Delegator** role.
  `blob_store.py` automatically uses user-delegation SAS when no connection
  string is present (expiry capped at 168 h / 7 days per Azure limits).

---

## Open questions

Before going to production, align on the following:

1. **Dataverse schema** — what are the exact field names for approved meeting
   title, date, attendee list, and per-topic bullet arrays in the Dataverse
   entity that feeds the Power Automate flow?

2. **Trigger flow** — which Power Automate flow calls the HTTP endpoint:
   7.3 (topic-level approval), 7.4 (meeting-level approval), or both?

3. **SharePoint destination** — confirm the exact folder path and file naming
   convention for the generated PPTX (e.g.
   `/Board Meeting Minutes/{YYYY-MM-DD}_{MeetingName}/Meeting_Minutes_{timestamp}.pptx`).

4. **Auth key distribution** — how will the function API key be stored
   securely in Power Automate? Recommend: Azure Key Vault reference surfaced
   as a Power Automate environment variable.

5. **Template ownership** — who is responsible for uploading a new PPTX
   template to the `templates-ppt` blob container when the branding changes?
   Is there a change-management process for this?

6. **Error routing** — where should generation failures be surfaced?
   Email alert to board-office? Power Apps error screen? Teams notification?
