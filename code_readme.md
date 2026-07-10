# SAP Board Meeting PPT Generator — Code Walkthrough

## What This Project Does

An **Azure Function** that watches a blob storage container. When you drop a JSON file in, it automatically generates a fully-populated PowerPoint deck from a template and writes the result + a SAS download link back to storage.

---

## Data Flow Diagram (Full)

```
Drop JSON → input-json container
    ↓  [blob trigger]
function_app.py
    ├── Parse JSON
    ├── Download template PPTX (templates-ppt/)
    ├── generate_pptx()
    │       ├── validate_content()        ← schema guard
    │       ├── build_content_slides()    ← split >4 bullets
    │       ├── unpack.py                 ← ZIP → XML files
    │       ├── parse presentation.xml    ← rId map + slide order
    │       ├── add_slide.py (×N)         ← clone extra content slides if needed
    │       ├── inject_title_slide()      ← regex replace on slide1 XML
    │       ├── inject_bullets() ×N       ← regex replace on content slide XMLs
    │       ├── hide_unused_circles()     ← strip unused numbered circles
    │       ├── clean.py                  ← remove temp files
    │       └── pack.py                   ← XML files → ZIP → .pptx bytes
    ├── Upload PPTX → output-pptx/<stem>_generated.pptx
    ├── Generate SAS URL (24h read-only)
    └── Write metadata sidecar → output-pptx/<stem>_metadata.json
              { status, sas_url, slide_count, processing_seconds, … }
```

---

## Layer 1 — Entry Point: `function_app.py`

This is the Azure Functions v2 app. One function: `generate_board_meeting_ppt`.

**Trigger:** Any `.json` blob dropped into the `input-json` container fires it automatically via `@app.blob_trigger`.

### Flow (6 steps)

| Step | What happens |
|------|-------------|
| 1 | Read & parse the JSON blob |
| 2 | Download the `.pptx` template from the `templates-ppt` container |
| 3 | Call `generate_pptx()` — pure Python, no Azure deps |
| 4 | Upload the output PPTX to `output-pptx/<stem>_generated.pptx` |
| 5 | Generate a SAS URL (time-limited download link, default 24h) |
| 6 | Write a `<stem>_metadata.json` sidecar with status, URL, timing |

### Error handling

- `ValueError` (bad input schema) → writes `<stem>_error.json`, does **NOT** retry
- Unexpected `Exception` → writes error sidecar AND re-raises so Azure retries (exponential backoff: 3 retries, 5s → 2min, configured in `host.json`)

### Storage mode switch (`function_app.py` line 66)

```python
if os.getenv("USE_LOCAL_STORAGE", "false").lower() == "true":
    from ppt_core import blob_store_local as blob_store   # plain files
else:
    from ppt_core import blob_store                       # real Azure
```

Set `USE_LOCAL_STORAGE=true` to go fully offline for local development.

---

## Layer 2 — PPT Generation: `ppt_core/generator.py`

The core logic. **Zero Azure dependencies** — fully unit-testable in isolation. A PPTX file is just a ZIP of XML files; the generator surgically edits that XML.

### Key Data Structures

| Class | Purpose |
|-------|---------|
| `TemplateGeometry` | Hardcoded EMU coordinates for the SAP template (positions of the 4 numbered circles, bullet text box sizes). 914400 EMU = 1 inch. |
| `TemplateSlotMap` | Which slide index is the title (1), which are content (4–8), which is the closing (9) |
| `GeneratorConfig` | Bundles geometry + slot_map + `scripts_dir`, `workdir`, `max_bullets` |
| `ContentSlide` | Final data model: `(title, bullets, speaker_notes)` |

### `generate_pptx()` — The Main Pipeline (9 Steps)

| Step | What happens |
|------|-------------|
| 1 | `validate_content()` — strict schema check, raises `ValueError` on bad input |
| 2 | `build_content_slides()` — if a slide has >4 bullets, `split_slide()` chunks it into "Title - Part 1", "Part 2", … |
| 3 | Write template bytes to temp `workdir/template.pptx` |
| 4 | **Unpack** — runs `scripts/office/unpack.py` which `zipfile.ZipFile.extractall()` the PPTX into XML files |
| 5 | Parse `presentation.xml` to get the rId→file mapping and visual slide order |
| 6 | **Clone extra slides** if needed — runs `scripts/add_slide.py` which copies an existing slide XML + updates `[Content_Types].xml` + `presentation.xml.rels` |
| 7a | Inject title slide — regex replaces `{meetingTitle}`, `{presenter}, SAP`, `{date}` placeholders in the slide XML |
| 7b | Inject each content slide — replaces `{sectionTitle}`, strips the `{sectionPoints}` box, inserts individual `<p:sp>` text boxes aligned to each circle row, hides unused circles |
| 8 | **Clean** — runs `scripts/clean.py` (removes `.DS_Store`, `~$*`, `.tmp` files) |
| 9 | **Pack** — runs `scripts/office/pack.py` which re-zips everything back into a `.pptx`, returns the bytes |

### Slide Splitting Logic

```
Input slide with 6 bullets (max=4)
    ↓
split_slide()
    ├── "Title - Part 1"  [bullet 1, 2, 3, 4]  + speaker_notes
    └── "Title - Part 2"  [bullet 5, 6]          (no notes)
```

### XML Injection (how placeholders work)

The template has placeholders split across multiple XML `<a:r>` runs:

```xml
<!-- Template XML (3-run pattern) -->
<a:r><a:rPr.../><a:t>{</a:t></a:r>
<a:r><a:rPr.../><a:t>meetingTitle</a:t></a:r>
<a:r><a:rPr.../><a:t>}</a:t></a:r>

<!-- After inject_title_slide() -->
<a:r><a:rPr lang="en-US" b="1" sz="3600">...</a:rPr><a:t>Actual Title</a:t></a:r>
```

Bullets are injected as standalone `<p:sp>` (text box) elements, each positioned at the exact EMU y-coordinate of its numbered circle row.

---

## Layer 3 — Storage: `ppt_core/blob_store.py` vs `blob_store_local.py`

### `blob_store.py` (Azure — production + real dev)

| Function | What it does |
|----------|-------------|
| `_get_service_client()` | Tries `AZURE_STORAGE_CONNECTION_STRING` first, then falls back to `DefaultAzureCredential` (Managed Identity in Azure, `az login` locally) |
| `download_blob(container, blob_name)` | Returns raw `bytes` of the blob |
| `upload_blob(container, blob_name, data)` | Auto-creates the container if missing, then uploads |
| `generate_sas_url(container, blob_name, expiry_hours)` | Parses account key from connection string, generates a read-only time-limited SAS token |

> **Note:** The Managed Identity path for `generate_sas_url` raises `NotImplementedError`. A commented-out user-delegation SAS alternative exists at the bottom of `blob_store.py` for production MI environments.

### `blob_store_local.py` (local dev — no Azure needed)

- `download_blob` → reads from `./local_storage/<container>/<blob>`
- `upload_blob` → writes to `./local_storage/<container>/<blob>` (creates dirs as needed)
- `generate_sas_url` → returns a `file://` path

---

## Layer 4 — Script Utilities: `scripts/`

| Script | Role |
|--------|------|
| `office/unpack.py` | Extracts PPTX (ZIP archive) → directory of XML files |
| `office/pack.py` | Re-zips directory of XML files → `.pptx` file |
| `add_slide.py` | Clones a slide XML file, adds new `Relationship` entry in `.rels`, adds `Override` in `[Content_Types].xml`, prints the new `rId` to stdout |
| `clean.py` | Removes `.DS_Store`, `~$*`, `*.tmp` files from the unpacked directory |

All scripts are invoked via `subprocess.run()` from `generator.py` using `_run_script()`.

---

## Local Dev Entry Points

| File | Purpose |
|------|---------|
| `run_local.py` | Minimal runner: reads from `local_storage/`, calls `generate_pptx()` directly, writes output |
| `test_local.py` | More complete: auto-discovers JSON files, handles errors, full logging |
| `setup_local_storage.sh` | Shell script to scaffold the `local_storage/` folder structure |

### Running locally

```bash
# 1. Put your template in place
cp /path/to/PowerpointTemplate_BoardMeetingGovernance.pptx local_storage/templates/

# 2. Put a test JSON in place
cp Meeting_Minutes_Presentation_Inputs.json local_storage/input-json/test_meeting.json

# 3. Run
USE_LOCAL_STORAGE=true python test_local.py
# Output → local_storage/output-pptx/test_meeting_generated.pptx
```

---

## JSON Input Schema

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

## Output — Metadata Sidecar

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

## Environment Variables

| Variable | Default | Description |
|----------|---------|-------------|
| `STORAGE_ACCOUNT_NAME` | *(required)* | Storage account name (used with Managed Identity) |
| `AZURE_STORAGE_CONNECTION_STRING` | *(optional)* | Overrides MI — use for local dev / CI |
| `USE_LOCAL_STORAGE` | `false` | Set to `true` to use local filesystem instead of Azure |
| `BLOB_INPUT_CONTAINER` | `input-json` | Container that triggers the function |
| `BLOB_TEMPLATE_CONTAINER` | `templates-ppt` | Container holding the PPTX template |
| `BLOB_TEMPLATE_NAME` | `PowerpointTemplate_BoardMeetingGovernance.pptx` | Template blob name |
| `BLOB_OUTPUT_CONTAINER` | `output-pptx` | Container for generated output |
| `SAS_EXPIRY_HOURS` | `24` | SAS URL validity window |
| `MAX_BULLETS_PER_SLIDE` | `4` | Bullets before auto-split |
| `PPTX_SCRIPTS_DIR` | `/home/site/wwwroot/scripts` | Path to pptx skill scripts |

---

## Key Design Decisions

1. **`generator.py` has zero Azure imports** — fully unit-testable without any cloud credentials
2. **Dual storage mode** — flip `USE_LOCAL_STORAGE=true` to go fully offline for development
3. **Regex-based XML editing** — avoids heavy XML parsers; works well but is fragile if the template structure changes significantly
4. **SAS via account key** — works with connection strings; the MI path for production would need the `Storage Account Key Operator` role or user-delegation SAS (code is commented in `blob_store.py`)
5. **Retry on infra errors only** — `ValueError` (bad JSON input) is swallowed and not retried; unexpected `Exception` is re-raised so Azure retries with exponential backoff (configured in `host.json`)
6. **Scripts called via subprocess** — `unpack.py`, `pack.py`, `add_slide.py`, `clean.py` are invoked as child processes, keeping them independently runnable and testable

---

## `host.json` — Runtime Configuration

```json
{
  "functionTimeout": "00:10:00",
  "retry": {
    "strategy": "exponentialBackoff",
    "maxRetryCount": 3,
    "minimumInterval": "00:00:05",
    "maximumInterval": "00:02:00"
  },
  "extensions": {
    "blobs": { "maxDegreeOfParallelism": 4 }
  }
}
```

- 10-minute timeout per invocation
- Up to 3 retries with backoff between 5s and 2 minutes
- Up to 4 blob triggers processed in parallel

---

## Dependencies (`requirements.txt`)

| Package | Version | Purpose |
|---------|---------|---------|
| `azure-functions` | 1.21.3 | Azure Functions Python worker |
| `azure-storage-blob` | 12.22.0 | Blob download / upload / SAS |
| `azure-identity` | 1.17.1 | Managed Identity + DefaultAzureCredential |
| `azure-core` | 1.30.2 | Azure SDK core (pinned for reproducibility) |

All PPTX manipulation is done with Python's built-in `zipfile` and string/regex operations — no `python-pptx` dependency.
