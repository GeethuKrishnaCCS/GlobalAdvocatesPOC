<#
.SYNOPSIS
    Deploys advocate-docs.sppkg to a test matter site's app catalog and syncs the command set properties.
.DESCRIPTION
    SharePoint applies elements.xml ClientSideComponentProperties only when the feature is first activated;
    an app upgrade does not update them. This script deploys/upgrades the app, then sets the DocActions web
    custom action properties to the current values so new properties (e.g. list names) take effect.
.PARAMETER Alias
    Matter site alias (/sites/<Alias>). Must already have a site collection app catalog.
.PARAMETER Execute
    Deploy and write. Without it the script only shows current vs wanted properties.
.EXAMPLE
    ./Deploy-AdvApp.ps1 -Execute
.NOTES
    use-when: deploy a new advocate-docs build to the test matter site; command set properties out of date after upgrade
    domain: tenant-sp
    safety: test-writes
    deps: PnP.PowerShell 3.x, pwsh 7.4+, built sharepoint/solution/advocate-docs.sppkg
    tenant-aware: yes
#>
[CmdletBinding()]
param(
    [ValidatePattern('^ADVTest[A-Za-z0-9]+$')][string]$Alias = 'ADVTestMatter01',
    [switch]$Execute
)
$ErrorActionPreference = 'Stop'

$componentId = '474974a6-76ee-4efa-9e8b-80bc77bbb7ec'
[string]$properties = [ordered]@{
    documentsLibrary   = 'ADVTestDocuments'
    bundlesLibrary     = 'ADVTestBundles'
    bundleRegistryList = 'ADVTestBundleRegistry'
    bundleItemsList    = 'ADVTestBundleItems'
    bundleManifestsLibrary = 'ADVTestBundleManifests'
    conversionLibrary  = 'ADVTestConversion'
    functionBaseUrl    = 'https://func-advtest-docs-poc.azurewebsites.net'
    apiResource        = 'api://daf35d1e-0756-4216-a61d-827d5eca434f'
} | ConvertTo-Json -Compress
$package = Join-Path $PSScriptRoot '../advocate-docs/sharepoint/solution/advocate-docs.sppkg'

try {
    $config = Get-Content (Join-Path $PSScriptRoot '../.project.json') -Raw | ConvertFrom-Json
    $siteUrl = "$($config.tenants.dev.url)/sites/$Alias"
    $conn = Connect-PnPOnline -Url $siteUrl -ClientId $config.entraClientId -Thumbprint $config.entraCertThumbprint `
        -Tenant $config.tenants.dev.tenantId -ReturnConnection

    $actions = (Invoke-PnPSPRestMethod -Url "/_api/web/UserCustomActions?`$select=Id,ClientSideComponentId,ClientSideComponentProperties" -Connection $conn).value |
        Where-Object { $_.ClientSideComponentId -eq $componentId }
    $actions | ForEach-Object { Write-Information "Custom action $($_.Id): $($_.ClientSideComponentProperties) -> $properties" -InformationAction Continue }
    if (-not $Execute) { Write-Warning 'Plan only — nothing deployed. Re-run with -Execute.'; return }

    $app = Add-PnPApp -Path $package -Scope Site -Publish -Overwrite -Connection $conn
    try { Update-PnPApp -Identity $app.Id -Scope Site -Connection $conn } catch { Install-PnPApp -Identity $app.Id -Scope Site -Connection $conn }
    $actions = (Invoke-PnPSPRestMethod -Url "/_api/web/UserCustomActions?`$select=Id,ClientSideComponentId" -Connection $conn).value |
        Where-Object { $_.ClientSideComponentId -eq $componentId }
    foreach ($a in $actions) {
        Invoke-PnPSPRestMethod -Method Merge -Url "/_api/web/UserCustomActions('$($a.Id)')" `
            -Content @{ ClientSideComponentProperties = $properties } -ContentType 'application/json;odata=nometadata' -Connection $conn | Out-Null
    }
    Write-Information "Deployed $($app.Title) and synced properties on $(@($actions).Count) custom action(s)." -InformationAction Continue
}
catch {
    Write-Error "Deploy failed: $($_.Exception.Message)"
    throw
}


