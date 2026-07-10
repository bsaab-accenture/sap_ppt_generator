################################################################################
# main.tf
# ──────────────────────────────────────────────────────────────────────────────
# Board Meeting PPT Generator — complete Azure infrastructure.
#
# Resources provisioned
# ─────────────────────
#   • azurerm_storage_account              — StorageV2, Standard LRS, HTTPS-only
#   • azurerm_storage_container ×3         — input-json / templates / output-pptx
#   • azurerm_log_analytics_workspace      — PerGB2018, configurable retention
#   • azurerm_application_insights         — workspace-based, web kind
#   • azurerm_service_plan                 — Premium EP1, Linux, elastic scaling
#   • azurerm_linux_function_app           — Python 3.11, v2 model, system MI
#   • azurerm_role_assignment ×2           — Storage Blob Data Contributor
#                                            Storage Account Key Operator
#
# Naming convention
# ─────────────────
#   All resource names use the local `prefix` = "${app_name}${environment}".
#   Storage account names are additionally lowercased and capped at 24 chars
#   (Azure requirement). A random_id suffix guarantees global uniqueness.
################################################################################

# ── Locals ─────────────────────────────────────────────────────────────────────

locals {
  prefix = "${var.app_name}${var.environment}"

  # Storage account names: lowercase, max 24 chars, globally unique
  # random_id adds 4 hex chars to avoid conflicts across deployments
  storage_account_name = lower(substr("st${local.prefix}${random_id.storage_suffix.hex}", 0, 24))

  resource_names = {
    storage_account  = local.storage_account_name
    function_app     = "func-${local.prefix}"
    app_service_plan = "asp-${local.prefix}"
    app_insights     = "appi-${local.prefix}"
    log_analytics    = "law-${local.prefix}"
  }

  # Well-known Azure built-in role definition IDs (stable across all tenants)
  role_ids = {
    storage_blob_data_contributor  = "ba92f5b4-2d11-453d-a403-e96b0029c9fe"
    storage_account_key_operator   = "81a9662b-bebf-436f-a333-f67b29880f12"
  }

  common_tags = merge(var.tags, {
    environment = var.environment
    app_name    = var.app_name
  })
}

# ── Random suffix for globally-unique storage account name ─────────────────────

resource "random_id" "storage_suffix" {
  byte_length = 2   # 4 hex chars — enough to avoid accidental collisions
}

# ── Log Analytics Workspace ────────────────────────────────────────────────────

resource "azurerm_log_analytics_workspace" "main" {
  name                = local.resource_names.log_analytics
  location            = var.location
  resource_group_name = var.resource_group_name
  sku                 = "PerGB2018"
  retention_in_days   = var.log_retention_days
  tags                = local.common_tags
}

# ── Application Insights ───────────────────────────────────────────────────────

resource "azurerm_application_insights" "main" {
  name                = local.resource_names.app_insights
  location            = var.location
  resource_group_name = var.resource_group_name
  workspace_id        = azurerm_log_analytics_workspace.main.id
  application_type    = "web"
  retention_in_days   = var.log_retention_days
  tags                = local.common_tags
}

# ── Storage Account ────────────────────────────────────────────────────────────

resource "azurerm_storage_account" "main" {
  name                     = local.resource_names.storage_account
  resource_group_name      = var.resource_group_name
  location                 = var.location
  account_tier             = "Standard"
  account_replication_type = "LRS"
  account_kind             = "StorageV2"
  access_tier              = "Hot"

  # Security hardening
  enable_https_traffic_only       = true
  min_tls_version                 = "TLS1_2"
  allow_nested_items_to_be_public = false   # no anonymous blob access

  blob_properties {
    delete_retention_policy {
      days = var.blob_delete_retention_days > 0 ? var.blob_delete_retention_days : null
    }
  }

  # Network: allow all by default; tighten with network_rules in production
  # to restrict to the Function App's outbound IPs or a VNet service endpoint.
  # network_rules {
  #   default_action             = "Deny"
  #   ip_rules                   = ["<your-ci-ip>"]
  #   virtual_network_subnet_ids = [azurerm_subnet.func.id]
  # }

  tags = local.common_tags
}

# ── Blob Containers ────────────────────────────────────────────────────────────

resource "azurerm_storage_container" "input_json" {
  name                  = var.blob_input_container
  storage_account_name  = azurerm_storage_account.main.name
  container_access_type = "private"
}

resource "azurerm_storage_container" "templates" {
  name                  = var.blob_template_container
  storage_account_name  = azurerm_storage_account.main.name
  container_access_type = "private"
}

resource "azurerm_storage_container" "output_pptx" {
  name                  = var.blob_output_container
  storage_account_name  = azurerm_storage_account.main.name
  container_access_type = "private"
}

# ── App Service Plan — Premium EP1 ────────────────────────────────────────────

resource "azurerm_service_plan" "main" {
  name                         = local.resource_names.app_service_plan
  resource_group_name          = var.resource_group_name
  location                     = var.location
  os_type                      = "Linux"
  sku_name                     = "EP1"             # Elastic Premium: pre-warmed, VNet-ready
  maximum_elastic_worker_count = var.max_elastic_worker_count
  tags                         = local.common_tags
}

# ── Linux Function App ─────────────────────────────────────────────────────────

resource "azurerm_linux_function_app" "main" {
  name                       = local.resource_names.function_app
  resource_group_name        = var.resource_group_name
  location                   = var.location
  service_plan_id            = azurerm_service_plan.main.id
  storage_account_name       = azurerm_storage_account.main.name
  storage_uses_managed_identity = true   # no storage connection string in app settings
  https_only                 = true

  # System-assigned Managed Identity — no credentials stored anywhere
  identity {
    type = "SystemAssigned"
  }

  site_config {
    pre_warmed_instance_count = var.pre_warmed_instance_count
    ftps_state                = "Disabled"
    minimum_tls_version       = "1.2"
    http2_enabled             = true

    application_stack {
      python_version = var.python_version
    }

    # Blob trigger: allow up to 4 concurrent blob processing jobs per instance
    # Matches host.json extensions.blobs.maxDegreeOfParallelism = 4
    app_scale_limit = 5
  }

  app_settings = {
    # ── Azure Functions runtime ──────────────────────────────────────────────
    FUNCTIONS_EXTENSION_VERSION = "~4"
    FUNCTIONS_WORKER_RUNTIME    = "python"

    # ── Application Insights ─────────────────────────────────────────────────
    APPLICATIONINSIGHTS_CONNECTION_STRING        = azurerm_application_insights.main.connection_string
    ApplicationInsightsAgent_EXTENSION_VERSION   = "~3"

    # ── Storage (Managed Identity — no key stored) ───────────────────────────
    STORAGE_ACCOUNT_NAME = azurerm_storage_account.main.name

    # ── Container names ───────────────────────────────────────────────────────
    BLOB_INPUT_CONTAINER    = var.blob_input_container
    BLOB_TEMPLATE_CONTAINER = var.blob_template_container
    BLOB_TEMPLATE_NAME      = var.template_blob_name
    BLOB_OUTPUT_CONTAINER   = var.blob_output_container

    # ── Generation parameters ─────────────────────────────────────────────────
    SAS_EXPIRY_HOURS      = tostring(var.sas_expiry_hours)
    MAX_BULLETS_PER_SLIDE = tostring(var.max_bullets_per_slide)
    PPTX_SCRIPTS_DIR      = var.pptx_scripts_dir

    # ── Pre-warm ──────────────────────────────────────────────────────────────
    PRE_WARMED_INSTANCE_COUNT = tostring(var.pre_warmed_instance_count)
  }

  tags = local.common_tags

  # Explicit dependency: role assignments must exist before the function
  # attempts to access storage on first startup.
  depends_on = [
    azurerm_role_assignment.func_blob_contributor,
    azurerm_role_assignment.func_key_operator,
  ]
}

# ── Role assignments — Function App MI → Storage Account ──────────────────────
#
# Storage Blob Data Contributor (ba92f5b4-...)
#   Grants: read, write, delete, list blobs.
#   Required for: download_blob(), upload_blob() in blob_store.py.
#
# Storage Account Key Operator (81a9662b-...)
#   Grants: list + regenerate storage account keys.
#   Required for: generate_sas_url() signing with account key in blob_store.py.
#   Alternative: grant Storage Blob Delegator + use user-delegation SAS
#                (see NOTE in blob_store.py).

resource "azurerm_role_assignment" "func_blob_contributor" {
  scope                = azurerm_storage_account.main.id
  role_definition_id   = "/subscriptions/${data.azurerm_client_config.current.subscription_id}/providers/Microsoft.Authorization/roleDefinitions/${local.role_ids.storage_blob_data_contributor}"
  principal_id         = azurerm_linux_function_app.main.identity[0].principal_id
  principal_type       = "ServicePrincipal"

  # Terraform sometimes races on MI propagation; a small delay is baked into
  # the depends_on chain implicitly, but add a lifecycle guard for safety.
  lifecycle {
    ignore_changes = [scope]   # scope is stable; prevents spurious plan diffs
  }
}

resource "azurerm_role_assignment" "func_key_operator" {
  scope                = azurerm_storage_account.main.id
  role_definition_id   = "/subscriptions/${data.azurerm_client_config.current.subscription_id}/providers/Microsoft.Authorization/roleDefinitions/${local.role_ids.storage_account_key_operator}"
  principal_id         = azurerm_linux_function_app.main.identity[0].principal_id
  principal_type       = "ServicePrincipal"

  lifecycle {
    ignore_changes = [scope]
  }
}

# ── Data sources ───────────────────────────────────────────────────────────────

# Current authenticated client — used for subscription_id in role assignment scope
data "azurerm_client_config" "current" {}
