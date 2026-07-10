# Local Testing Guide

Complete guide for running and testing the Azure Function locally using filesystem storage instead of Azure Blob Storage.

## Prerequisites

- Azure Functions Core Tools v4 installed (`func --version`)
- Python 3.9+ with dependencies installed (`pip install -r requirements.txt`)
- PowerPoint template file
- Test JSON input files

## Quick Start

### 1. Setup Local Storage

```bash
chmod +x setup_local_storage.sh && ./setup_local_storage.sh
```

This creates:
```
local_storage/
├── input-json/          # Put test JSON files here
├── templates/           # Put PPTX template here
└── output-pptx/         # Generated files appear here
```

### 2. Add Required Files

**Template:**
```bash
cp /path/to/your/PowerpointTemplate_BoardMeetingGovernance.pptx \
   local_storage/templates/
```

**Test JSON:**
```bash
cp /path/to/test_meeting.json \
   local_storage/input-json/
```

### 3. Choose Testing Method

## Method A: Standalone Script (Fastest)

No Azure Functions runtime needed. Direct execution.

```bash
python test_local.py
```

Output appears in `local_storage/output-pptx/`.

**Pros:**
- Fastest iteration
- No func runtime overhead
- Simple debugging

**Cons:**
- Doesn't test blob trigger behavior
- Doesn't test Azure Functions runtime integration

## Method B: Azure Functions Runtime (Most Realistic)

Tests the actual function trigger and runtime behavior.

### Step 1: Start the Function

```bash
func start
```

You should see:
```
🚀 Function App initialized - Using LOCAL file storage
Functions:
  generate_board_meeting_ppt: blobTrigger
```

### Step 2: Trigger by Adding File

In another terminal, add a JSON file to input folder:

```bash
cp test_data/example.json local_storage/input-json/
```

Or create a new file:

```bash
cat > local_storage/input-json/test_$(date +%s).json << 'EOF'
{
  "meeting_title": "Board Meeting",
  "meeting_date": "2026-06-03",
  "slides": [
    {"title": "Welcome", "content": "Test slide"}
  ]
}
EOF
```

### Step 3: Watch Logs

The function should trigger automatically and process the file:

```
▶  Triggered by blob: input-json/test_1717401234.json
Parsed JSON: 3 top-level keys
Template downloaded: templates/PowerpointTemplate_BoardMeetingGovernance.pptx (245123 bytes)
PPTX generated: 189456 bytes
✅  Done in 2.34s  →  output-pptx/test_1717401234_generated.pptx
```

## Configuration

Edit `local.settings.json` to customize:

```json
{
  "Values": {
    "USE_LOCAL_STORAGE": "true",           // Switch between local/Azure
    "BLOB_INPUT_CONTAINER": "input-json",
    "BLOB_TEMPLATE_CONTAINER": "templates",
    "BLOB_TEMPLATE_NAME": "PowerpointTemplate_BoardMeetingGovernance.pptx",
    "BLOB_OUTPUT_CONTAINER": "output-pptx",
    "MAX_BULLETS_PER_SLIDE": "4",
    "PPTX_SCRIPTS_DIR": "./scripts"
  }
}
```

## Switching to Azure Storage

To test with real Azure Blob Storage:

1. Set `USE_LOCAL_STORAGE` to `false` in `local.settings.json`
2. Add your Azure storage credentials:

```json
{
  "Values": {
    "USE_LOCAL_STORAGE": "false",
    "AZURE_STORAGE_CONNECTION_STRING": "DefaultEndpointsProtocol=https;AccountName=...",
    // OR use Managed Identity:
    "STORAGE_ACCOUNT_NAME": "your-storage-account"
  }
}
```

3. Restart: `func start`

## Troubleshooting

### "Template not found"

```bash
# Check template exists
ls -lh local_storage/templates/*.pptx

# Verify filename matches config
grep BLOB_TEMPLATE_NAME local.settings.json
```

### "No JSON files found"

```bash
# Check input directory
ls -lh local_storage/input-json/*.json

# Create test file
echo '{"test": true}' > local_storage/input-json/test.json
```

### "Module not found"

```bash
# Install dependencies
pip install -r requirements.txt

# Check Python path
python -c "import ppt_core; print(ppt_core.__file__)"
```

### Function not triggering

```bash
# Check function is running
func start

# Verify local storage mode
grep USE_LOCAL_STORAGE local.settings.json

# Add file AFTER function starts
cp test.json local_storage/input-json/
```

## Verification

After successful run, check:

```bash
# Output file exists
ls -lh local_storage/output-pptx/*_generated.pptx

# Metadata sidecar created
ls -lh local_storage/output-pptx/*_metadata.json

# View metadata
cat local_storage/output-pptx/*_metadata.json | jq .
```

## Performance Testing

Generate multiple test files:

```bash
for i in {1..5}; do
  cp local_storage/input-json/test.json \
     local_storage/input-json/test_$i.json
done
```

Watch processing:

```bash
func start | grep "Done in"
```

## Cleanup

```bash
# Remove generated files
rm -rf local_storage/output-pptx/*

# Reset test files
rm -rf local_storage/input-json/*
```

## Next Steps

- Add your actual PowerPoint template
- Create comprehensive test JSON fixtures
- Test error scenarios (invalid JSON, missing fields)
- Benchmark performance with large presentations
- Deploy to Azure when ready
