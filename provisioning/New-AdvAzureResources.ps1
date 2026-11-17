<#
.SYNOPSIS
    Provisions the AdvocateDocs POC Azure resources in one dedicated resource group.
.DESCRIPTION
    Creates (idempotent) in the subscription/account recorded in .project.json (tenants.dev.azure):
      - resource group      rg-advtest-docs-poc   (tags project/owner/env)
      - storage account     stadvtestdocspoc      (StorageV2, Standard_LRS, TLS1.2, no public blob access)
      - function app        func-advtest-docs-poc (Flex Consumption, dotnet-isolated 8.0; creates its own
                                                   Application Insights component with the same name)
    The subscription is shared: the script only ever touches the one resource group it creates.
    Without -Execute it prints the plan and checks what already exists.
.PARAMETER Location
    Azure region. Must support Flex Consumption.
.PARAMETER Execute
    Create missing resources. Without it nothing is written.
.EXAMPLE
    ./New-AdvAzureResources.ps1            # plan
    ./New-AdvAzureResources.ps1 -Execute   # create
.NOTES
    use-when: create or check the AdvocateDocs POC Azure resource group, storage and function app
    domain: azure
    safety: test-writes
    deps: az CLI 2.60+, signed in as the .project.json azure.account (az login --tenant <tenantId>), pwsh 7.4+
    tenant-aware: yes
#>
[CmdletBinding()]
param(
    [ValidateSet('centralindia', 'southindia', 'uaenorth')][string]$Location = 'centralindia',
    [switch]$Execute
)
$ErrorActionPreference = 'Stop'

$config = Get-Content (Join-Path $PSScriptRoot '../.project.json') -Raw | ConvertFrom-Json
$azure = $config.tenants.dev.azure
$names = @{
    Group    = 'rg-advtest-docs-poc'
    Storage  = 'stadvtestdocspoc'
    Function = 'func-advtest-docs-poc'
}
$tags = @('project=AdvocateDocs-POC', "owner=$($azure.account.Split('@')[0])", 'env=poc')

function Invoke-Az {
    # az writes warnings to stderr; only the exit code decides failure.
    $output = & az @args 2>$null
    if ($LASTEXITCODE -ne 0) { throw "az $($args -join ' ') failed (exit $LASTEXITCODE)." }
    return $output
}

function Test-AzExists([string[]]$azArgs) {
    & az @azArgs --only-show-errors 2>$null | Out-Null
    return $LASTEXITCODE -eq 0
}

try {
    $account = Invoke-Az account show --query '{user:user.name, tenant:tenantId, sub:id}' -o json | ConvertFrom-Json
    if ($account.user -ne $azure.account -or $account.tenant -ne $config.tenants.dev.tenantId) {
        throw "Signed in as $($account.user) (tenant $($account.tenant)). Run: az login --tenant $($config.tenants.dev.tenantId) as $($azure.account)."
    }
    $subscription = if ($azure.subscriptionId) { $azure.subscriptionId } else { $account.sub }
    Invoke-Az account set --subscription $subscription | Out-Null

    $plan = @(
        [pscustomobject]@{ Resource = 'Resource group'; Name = $names.Group; Exists = (Test-AzExists @('group', 'show', '-n', $names.Group)) }
        [pscustomobject]@{ Resource = 'Storage account'; Name = $names.Storage; Exists = (Test-AzExists @('storage', 'account', 'show', '-n', $names.Storage, '-g', $names.Group)) }
        [pscustomobject]@{ Resource = 'Function app (Flex, .NET 8 isolated) + App Insights'; Name = $names.Function; Exists = (Test-AzExists @('functionapp', 'show', '-n', $names.Function, '-g', $names.Group)) }
    )
    Write-Information "Subscription $subscription, region $Location" -InformationAction Continue
    $plan | Format-Table -AutoSize | Out-String | Write-Information -InformationAction Continue
    if (-not $Execute) { Write-Warning 'Plan only — nothing created. Re-run with -Execute.'; return }

    if (-not $plan[0].Exists) {
        Invoke-Az group create -n $names.Group -l $Location --tags @tags -o none | Out-Null
    }
    if (-not $plan[1].Exists) {
        Invoke-Az storage account create -n $names.Storage -g $names.Group -l $Location --sku Standard_LRS --kind StorageV2 `
            --min-tls-version TLS1_2 --allow-blob-public-access false --tags @tags -o none | Out-Null
    }
    if (-not $plan[2].Exists) {
        Invoke-Az functionapp create -n $names.Function -g $names.Group --storage-account $names.Storage `
            --flexconsumption-location $Location --runtime dotnet-isolated --runtime-version 8.0 --tags @tags -o none | Out-Null
    }

    Invoke-Az resource list -g $names.Group --query '[].{type:type, name:name, location:location}' -o table
}
catch {
    Write-Error "Azure provisioning failed: $($_.Exception.Message)"
    throw
}
