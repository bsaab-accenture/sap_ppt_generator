Prerequisites
Install these if you haven't already:

# Azure CLI
brew update && brew install azure-cli

# Azure Functions Core Tools (v4)
brew tap azure/functions
brew install azure-functions-core-tools@4

# Azure Storage Emulator (cross-platform replacement for old emulator)
npm install -g azurite
Also install the Azure Functions VS Code extension.

Step 1 — Create local.settings.json
The file referenced in the docstring (local.settings.json.example) needs to be copied and filled in. Create local.settings.json in the function folder:


{
  "IsEncrypted": false,
  "Values": {
    "AzureFunctionsJobHost__extensionBundle__id": "Microsoft.Azure.Functions.ExtensionBundle",
    "FUNCTIONS_WORKER_RUNTIME": "python",

    "AzureWebJobsStorage": "UseDevelopmentStorage=true",

    "AZURE_STORAGE_CONNECTION_STRING": "<your-real-storage-connection-string>",

    "BLOB_INPUT_CONTAINER": "input-json",
    "BLOB_TEMPLATE_CONTAINER": "templates-ppt",
    "BLOB_TEMPLATE_NAME": "PowerpointTemplate_BoardMeetingGovernance.pptx",
    "BLOB_OUTPUT_CONTAINER": "output-pptx",
    "SAS_EXPIRY_HOURS": "24"
  }
}
Two options for storage:

Option	When to use	Value for AZURE_STORAGE_CONNECTION_STRING
Real Azure storage	Need real data / template blobs	Connection string from Azure portal
Azurite (local emulator)	Isolated local dev	UseDevelopmentStorage=true
Step 2 — Option A: Use Azurite (local emulator)
Best for development — no Azure costs, works offline.


# Start Azurite in a terminal
azurite --location ~/.azurite --debug ~/.azurite/debug.log
Then create the required containers using Azure Storage Explorer or CLI:


# Point CLI at local emulator
export AZURE_STORAGE_CONNECTION_STRING="UseDevelopmentStorage=true"

az storage container create --name input-json --connection-string $AZURE_STORAGE_CONNECTION_STRING
az storage container create --name templates-ppt --connection-string $AZURE_STORAGE_CONNECTION_STRING
az storage container create --name output-pptx --connection-string $AZURE_STORAGE_CONNECTION_STRING

# Upload your PPTX template
az storage blob upload \
  --container-name templates-ppt \
  --name PowerpointTemplate_BoardMeetingGovernance.pptx \
  --file /path/to/your/template.pptx \
  --connection-string $AZURE_STORAGE_CONNECTION_STRING
Set both AzureWebJobsStorage and AZURE_STORAGE_CONNECTION_STRING to UseDevelopmentStorage=true in local.settings.json.

Step 3 — Option B: Use real Azure storage
Get the connection string from the portal:
Storage Account → Security + networking → Access keys → Connection string

Paste it as the value for AZURE_STORAGE_CONNECTION_STRING in local.settings.json. Set AzureWebJobsStorage to the same string.

Step 4 — Start the function

cd sap_board/ppt_generator/board_meeting_ppt_azure_function
func start
Or press F5 in VS Code (the Azure Functions extension handles this automatically if a launch.json exists).

Step 5 — Trigger it by uploading a test blob

az storage blob upload \
  --account-name <acct-or-emulator> \
  --container-name input-json \
  --name test_meeting.json \
  --file Meeting_Minutes_Presentation_Inputs.json \
  --connection-string $AZURE_STORAGE_CONNECTION_STRING
The blob trigger will fire automatically and you'll see logs in the terminal.

Common issues
Problem	Fix
AzureWebJobsStorage missing	Add it to local.settings.json — Functions runtime requires it
Blob trigger doesn't fire	Make sure AzureWebJobsStorage points to same storage as your containers
local.settings.json not found	Must be in the same folder as function_app.py
local.settings.json committed to git	Add it to .gitignore — it contains secrets
Make sure local.settings.json is in your .gitignore since it will contain your storage connection string.