################################################################################
# versions.tf
# ──────────────────────────────────────────────────────────────────────────────
# Provider requirements and Terraform backend configuration.
#
# Backend
# -------
# The azurerm backend block is commented out below.
# Uncomment and fill in the values to store state in Azure Blob Storage
# (strongly recommended for team / CI use):
#
#   terraform {
#     backend "azurerm" {
#       resource_group_name  = "rg-tfstate"
#       storage_account_name = "sttfstatepptgen"
#       container_name       = "tfstate"
#       key                  = "pptgen.prod.tfstate"
#     }
#   }
#
# Bootstrap that backend storage account once with:
#   az group create -n rg-tfstate -l westeurope
#   az storage account create -n sttfstatepptgen -g rg-tfstate --sku Standard_LRS
#   az storage container create -n tfstate --account-name sttfstatepptgen
################################################################################

terraform {
  required_version = ">= 1.7.0"

  required_providers {
    azurerm = {
      source  = "hashicorp/azurerm"
      version = "~> 3.110"
    }
    random = {
      source  = "hashicorp/random"
      version = "~> 3.6"
    }
  }

  # Uncomment to enable remote state in Azure Blob Storage:
  # backend "azurerm" {
  #   resource_group_name  = "rg-tfstate"
  #   storage_account_name = "sttfstatepptgen"
  #   container_name       = "tfstate"
  #   key                  = "pptgen.prod.tfstate"
  # }
}

provider "azurerm" {
  features {
    resource_group {
      # Prevent accidental deletion of non-empty resource groups
      prevent_deletion_if_contains_resources = true
    }
    key_vault {
      purge_soft_delete_on_destroy    = false
      recover_soft_deleted_key_vaults = true
    }
  }
}
