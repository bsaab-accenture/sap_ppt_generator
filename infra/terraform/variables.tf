################################################################################
# variables.tf
# ──────────────────────────────────────────────────────────────────────────────
# All input variables for the Board Meeting PPT Generator infrastructure.
# Override in terraform.tfvars or via -var flags / CI environment variables.
################################################################################

# ── Naming & placement ─────────────────────────────────────────────────────────

variable "app_name" {
  description = "Short application name used as a prefix for all resource names (e.g. 'pptgen')."
  type        = string

  validation {
    condition     = length(var.app_name) >= 3 && length(var.app_name) <= 12 && can(regex("^[a-z0-9]+$", var.app_name))
    error_message = "app_name must be 3–12 lowercase alphanumeric characters."
  }
}

variable "environment" {
  description = "Deployment environment. Controls resource naming suffix and retention policies."
  type        = string
  default     = "prod"

  validation {
    condition     = contains(["dev", "staging", "prod"], var.environment)
    error_message = "environment must be one of: dev, staging, prod."
  }
}

variable "location" {
  description = "Azure region for all resources."
  type        = string
  default     = "westeurope"
}

variable "resource_group_name" {
  description = "Name of the target resource group. Must already exist."
  type        = string
}

variable "tags" {
  description = "Tags applied to every resource."
  type        = map(string)
  default = {
    project   = "board-meeting-ppt-generator"
    managedBy = "terraform"
  }
}

# ── Storage ────────────────────────────────────────────────────────────────────

variable "blob_input_container" {
  description = "Storage container that receives JSON input files and triggers the function."
  type        = string
  default     = "input-json"
}

variable "blob_template_container" {
  description = "Storage container holding the PPTX template file."
  type        = string
  default     = "templates"
}

variable "blob_output_container" {
  description = "Storage container where generated PPTX files and metadata sidecars are written."
  type        = string
  default     = "output-pptx"
}

variable "blob_delete_retention_days" {
  description = "Soft-delete retention for blobs in days. Set to 0 to disable."
  type        = number
  default     = 7

  validation {
    condition     = var.blob_delete_retention_days >= 0 && var.blob_delete_retention_days <= 365
    error_message = "blob_delete_retention_days must be between 0 and 365."
  }
}

# ── Function App ───────────────────────────────────────────────────────────────

variable "python_version" {
  description = "Python runtime version for the Function App."
  type        = string
  default     = "3.11"

  validation {
    condition     = contains(["3.10", "3.11", "3.12"], var.python_version)
    error_message = "python_version must be one of: 3.10, 3.11, 3.12."
  }
}

variable "pre_warmed_instance_count" {
  description = "Number of pre-warmed instances on the Premium plan. Minimum 1 to avoid cold starts."
  type        = number
  default     = 1

  validation {
    condition     = var.pre_warmed_instance_count >= 1 && var.pre_warmed_instance_count <= 10
    error_message = "pre_warmed_instance_count must be between 1 and 10."
  }
}

variable "max_elastic_worker_count" {
  description = "Maximum number of elastic worker instances for burst scaling."
  type        = number
  default     = 5
}

# ── Application-level settings ─────────────────────────────────────────────────

variable "template_blob_name" {
  description = "Name of the PPTX template blob inside blob_template_container."
  type        = string
  default     = "PowerpointTemplate_BoardMeetingGovernance.pptx"
}

variable "sas_expiry_hours" {
  description = "SAS URL validity window in hours for generated PPTX output blobs."
  type        = number
  default     = 24

  validation {
    condition     = var.sas_expiry_hours >= 1 && var.sas_expiry_hours <= 168
    error_message = "sas_expiry_hours must be between 1 and 168 (one week)."
  }
}

variable "max_bullets_per_slide" {
  description = "Maximum bullet points per slide before auto-splitting into Part 1 / Part 2."
  type        = number
  default     = 4

  validation {
    condition     = var.max_bullets_per_slide >= 1 && var.max_bullets_per_slide <= 10
    error_message = "max_bullets_per_slide must be between 1 and 10."
  }
}

variable "pptx_scripts_dir" {
  description = "Absolute path to the pptx skill scripts directory inside the Function App."
  type        = string
  default     = "/home/site/wwwroot/scripts"
}

# ── Observability ──────────────────────────────────────────────────────────────

variable "log_retention_days" {
  description = "Log Analytics workspace and Application Insights retention in days."
  type        = number
  default     = 30

  validation {
    condition     = contains([30, 60, 90, 120, 180, 270, 365], var.log_retention_days)
    error_message = "log_retention_days must be one of: 30, 60, 90, 120, 180, 270, 365."
  }
}
