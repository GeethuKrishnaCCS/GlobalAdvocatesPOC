<#
.SYNOPSIS
    Builds the Office add-in for production and publishes the web files to an Azure Storage static website.
.DESCRIPTION
    Every code release goes through here and is live at once: the registered manifest points at the hosted URL,
    so only ribbon/permission/id/version changes need Publish-AdvAddinManifest.ps1.
    Steps: ensure storage account + static website (dev resource group from .project.json), build with
    ADV_ADDIN_URL_PROD = the static website origin, upload dist/ with no-cache (bundle names are not hashed),
    then smoke-check taskpane.html, manifest.xml and that the hosted manifest has no localhost URL left.
    Writes the prod manifest to dist/manifest.xml for Publish-AdvAddinManifest.ps1.
.PARAMETER StorageAccount
    Storage account name (globally unique, 3-24 lowercase letters/digits).
.PARAMETER Execute
    Create/upload. Without it the script only shows the plan.
.EXAMPLE
    ./scripts/Deploy-AdvAddinWeb.ps1            # plan
    ./scripts/Deploy-AdvAddinWeb.ps1 -Execute   # dev tenant (az signed in to the ccsdev01 subscription)
.NOTES
    use-when: release add-in code changes; hosted add-in shows old version; first production hosting for the add-in
    domain: azure
    safety: tenant-write-gate
    deps: az CLI signed in to the project subscription, node/npm, Repos/AdvocateDocs-POC/.project.json
    tenant-aware: yes
#>
[CmdletBinding()]
param(
    [ValidatePattern('^[a-z0-9]{3,24}$')][string]$StorageAccount = 'stadvtestaddin01',
    [switch]$Execute
)
$ErrorActionPreference = 'Stop'

$addinRoot = Split-Path $PSScriptRoot -Parent
$projectsRoot = if ($env:PROJECTS_ROOT) { $env:PROJECTS_ROOT } else { 'D:\Projects' }
$azure = (Get-Content (Join-Path $projectsRoot 'Repos/AdvocateDocs-POC/.project.json') -Raw | ConvertFrom-Json).tenants.dev.azure

function Invoke-Az {
    $out = az @args 2>&1
    if ($LASTEXITCODE -ne 0) { throw "az $($args -join ' ') failed: $out" }
    $out
}

Write-Information "Subscription $($azure.subscriptionId), RG $($azure.resourceGroup), account $StorageAccount ($($azure.location))" -InformationAction Continue
Write-Information '1. Ensure StorageV2 account + static website (index taskpane.html)' -InformationAction Continue
Write-Information '2. npm run build with ADV_ADDIN_URL_PROD = static website origin' -InformationAction Continue
Write-Information '3. Upload dist/ to $web (overwrite, Cache-Control no-cache)' -InformationAction Continue
Write-Information '4. Smoke: taskpane.html + manifest.xml 200, no localhost in hosted manifest' -InformationAction Continue
if (-not $Execute) { Write-Warning 'Plan only - nothing created or uploaded. Re-run with -Execute.'; return }

try {
    Invoke-Az account set --subscription $azure.subscriptionId | Out-Null
    $exists = az storage account show -n $StorageAccount -g $azure.resourceGroup --query name -o tsv 2>$null
    if (-not $exists) {
        Invoke-Az storage account create -n $StorageAccount -g $azure.resourceGroup -l $azure.location --sku Standard_LRS --kind StorageV2 --min-tls-version TLS1_2 | Out-Null
        Write-Information "Created storage account $StorageAccount" -InformationAction Continue
    }
    $key = Invoke-Az storage account keys list -n $StorageAccount -g $azure.resourceGroup --query '[0].value' -o tsv
    Invoke-Az storage blob service-properties update --account-name $StorageAccount --account-key $key --static-website --index-document taskpane.html | Out-Null
    $origin = (Invoke-Az storage account show -n $StorageAccount -g $azure.resourceGroup --query primaryEndpoints.web -o tsv).Trim()

    Push-Location $addinRoot
    try {
        $env:ADV_ADDIN_URL_PROD = $origin
        npm run build
        if ($LASTEXITCODE -ne 0) { throw 'npm run build failed' }
    }
    finally { Remove-Item Env:ADV_ADDIN_URL_PROD -ErrorAction SilentlyContinue; Pop-Location }

    Invoke-Az storage blob upload-batch --account-name $StorageAccount --account-key $key -d '$web' -s (Join-Path $addinRoot 'dist') `
        --overwrite --content-cache-control 'no-cache' --only-show-errors | Out-Null

    $manifest = Invoke-WebRequest "$($origin)manifest.xml" -UseBasicParsing
    $pane = Invoke-WebRequest "$($origin)taskpane.html" -UseBasicParsing
    if ($pane.StatusCode -ne 200 -or $manifest.StatusCode -ne 200) { throw 'Smoke check failed: hosted files not reachable' }
    if ($manifest.Content -match 'localhost') { throw 'Hosted manifest still references localhost' }
    Write-Information "Live: $($origin)taskpane.html. Prod manifest: $(Join-Path $addinRoot 'dist/manifest.xml')" -InformationAction Continue
}
catch {
    Write-Error "Deploy failed: $($_.Exception.Message)"
}
