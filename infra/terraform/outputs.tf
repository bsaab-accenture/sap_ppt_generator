################################################################################
# outputs.tf
# ──────────────────────────────────────────────────────────────────────────────
# Values emitted after a successful `terraform apply`.
# Useful for:
#   - CI/CD pipeline steps that upload the template or deploy function code
#   - Smoke-test scripts that drop a test blob
#   - Dashboards and runbooks
################################################################################

output "function_app_name" {
  description = "Name of the deployed Azure Function App."
  value       = azurerm_linux_function_app.main.name
}

output "function_app_url" {
  description = "Default HTTPS hostname of the Function App."
  value       = "https://${azurerm_linux_function_app.main.default_hostname}"
}

output "storage_account_name" {
  description = "Name of the Storage Account (needed for az storage blob upload commands)."
  value       = azurerm_storage_account.main.name
}

output "input_container_name" {
  description = "Blob container that triggers the function (drop JSON files here)."
  value       = azurerm_storage_container.input_json.name
}

output "template_container_name" {
  description = "Blob container holding the PPTX template."
  value       = azurerm_storage_container.templates.name
}

output "output_container_name" {
  description = "Blob container where generated PPTX and metadata sidecars are written."
  value       = azurerm_storage_container.output_pptx.name
}

output "managed_identity_principal_id" {
  description = "Object ID of the Function App's system-assigned Managed Identity."
  value       = azurerm_linux_function_app.main.identity[0].principal_id
}

output "app_insights_connection_string" {
  description = "Application Insights connection string (also set as a Function App setting)."
  value       = azurerm_application_insights.main.connection_string
  sensitive   = true   # contains instrumentation key; suppress from default output
}

output "log_analytics_workspace_id" {
  description = "Resource ID of the Log Analytics workspace backing Application Insights."
  value       = azurerm_log_analytics_workspace.main.id
}

# ── Convenience commands printed after apply ────────────────────────────────────

output "next_steps" {
  description = "Quick-reference commands to complete the deployment."
  value       = <<-EOT
    # 1. Upload the PPTX template:
    az storage blob upload \
        --account-name ${azurerm_storage_account.main.name} \
        --container-name ${azurerm_storage_container.templates.name} \
        --name ${var.template_blob_name} \
        --file PowerpointTemplate_BoardMeetingGovernance.pptx \
        --auth-mode login

    # 2. Deploy function code:
    func azure functionapp publish ${azurerm_linux_function_app.main.name} --python

    # 3. Drop a test JSON to trigger generation:
    az storage blob upload \
        --account-name ${azurerm_storage_account.main.name} \
        --container-name ${azurerm_storage_container.input_json.name} \
        --name test_meeting.json \
        --file Meeting_Minutes_Presentation_Inputs.json \
        --auth-mode login

    # 4. Fetch the metadata sidecar (contains SAS URL):
    az storage blob download \
        --account-name ${azurerm_storage_account.main.name} \
        --container-name ${azurerm_storage_container.output_pptx.name} \
        --name test_meeting_metadata.json \
        --file - --auth-mode login | python -m json.tool
  EOT
}
