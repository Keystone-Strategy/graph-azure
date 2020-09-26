#  TODO: figure out how to retrieve certificates

variable "create_batch_account" {
  type    = number
  default = 0
}

locals {
  batch_account_count = var.create_batch_account == 1 ? 1 : 0
}

resource "azurerm_batch_account" "j1dev" {
  count               = local.batch_account_count
  name                = "j1devbatchaccount"
  resource_group_name = azurerm_resource_group.j1dev.name
  location            = azurerm_resource_group.j1dev.location
  # This can be BatchService | UserSubscription
  pool_allocation_mode = "BatchService"
  storage_account_id   = azurerm_storage_account.j1dev.id
}

resource "azurerm_batch_pool" "j1dev" {
  count               = local.batch_account_count
  name                = "j1devbatchpool"
  resource_group_name = azurerm_resource_group.j1dev.name
  account_name        = azurerm_batch_account.j1dev[count.index].name
  display_name        = "J1 Dev Batch Account Pool"
  vm_size             = "Standard_A1"
  node_agent_sku_id   = "batch.node.ubuntu 16.04"

  storage_image_reference {
    publisher = "microsoft-azure-batch"
    offer     = "ubuntu-server-container"
    sku       = "16-04-lts"
    version   = "latest"
  }

  auto_scale {
    evaluation_interval = "PT15M"

    formula = <<EOF
        startingNumberOfVMs = 1;
        maxNumberofVMs = 2;
        pendingTaskSamplePercent = $PendingTasks.GetSamplePercent(180 * TimeInterval_Second);
        pendingTaskSamples = pendingTaskSamplePercent < 70 ? startingNumberOfVMs : avg($PendingTasks.GetSample(180 *   TimeInterval_Second));
        $TargetDedicatedNodes=min(maxNumberofVMs, pendingTaskSamples);
  EOF
  }

  start_task {
    command_line         = "echo 'Hello j1dev batch account'"
    max_task_retry_count = 1
    wait_for_success     = true

    user_identity {
      # The user identity the task runs under
      auto_user {
        # Admin | NonAdmin, defaults to NonAdmin
        elevation_level = "NonAdmin"
        # Task | Pool defaults to Task
        scope = "Task"
      }
    }
  }
}

resource "azurerm_batch_application" "example" {
  count               = local.batch_account_count
  name                = "j1devbatchapplication"
  resource_group_name = azurerm_resource_group.j1dev.name
  account_name        = azurerm_batch_account.j1dev[count.index].name
}
