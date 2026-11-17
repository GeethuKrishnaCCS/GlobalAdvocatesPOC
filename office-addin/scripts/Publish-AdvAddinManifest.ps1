<#
.SYNOPSIS
    Registers or updates the add-in's XML manifest through Centralized Deployment and assigns it to users.
.DESCRIPTION
    Only needed when the manifest changes (ribbon, permissions, id, version); code releases go through
    Deploy-AdvAddinWeb.ps1. Uses the O365CentralizedAddInDeployment module (XML manifests only; the unified
    manifest needs the admin center Integrated apps page). Users see changes after Microsoft's propagation,
    up to 24 hours.
    An admin (Global or Exchange admin) runs it and signs in interactively. On a customer tenant, only with
    the customer's approval.
.PARAMETER ManifestPath
    Production manifest (dist/manifest.xml from Deploy-AdvAddinWeb.ps1; must not reference localhost).
.PARAMETER ProductId
    Existing add-in ProductId (from Get-OrganizationAddIn) to update. Omit for the first registration. An update
    needs a higher <Version> in manifest.xml than the registered one, or the upload fails.
.PARAMETER Members
    Users/groups (UPN or group email) to assign. Omit to assign to everyone.
.PARAMETER AssignOnly
    With -ProductId: only add the -Members assignments; the manifest is not re-uploaded (no version bump needed).
.PARAMETER Execute
    Register/update. Without it the script only shows the plan.
.EXAMPLE
    ./scripts/Publish-AdvAddinManifest.ps1 -Members dev06@ccsdev01.onmicrosoft.com            # plan
    ./scripts/Publish-AdvAddinManifest.ps1 -Members dev06@ccsdev01.onmicrosoft.com -Execute   # register
.NOTES
    use-when: add-in ribbon or manifest changed for production; first production deployment of the add-in; add-in missing for users
    domain: tenant-sp
    safety: tenant-write-gate
    deps: Windows PowerShell 5.1 or pwsh 7 with O365CentralizedAddInDeployment (Install-Module -Scope CurrentUser), tenant admin account
    tenant-aware: yes
#>
[CmdletBinding()]
param(
    [string]$ManifestPath = (Join-Path (Split-Path $PSScriptRoot -Parent) 'dist/manifest.xml'),
    [ValidatePattern('^[0-9a-fA-F-]{36}$')][string]$ProductId,
    [string[]]$Members,
    [switch]$AssignOnly,
    [switch]$Execute
)
$ErrorActionPreference = 'Stop'
$locale = 'en-US'

if (-not (Test-Path -LiteralPath $ManifestPath)) { throw "No manifest at $ManifestPath - run Deploy-AdvAddinWeb.ps1 -Execute first." }
# Absolute from here on: the O365CentralizedAddInDeployment cmdlets open the file with .NET, which resolves relative
# paths against the process start folder, not the PowerShell location (cd), so a relative path failed there (2026-11-09).
$ManifestPath = (Resolve-Path -LiteralPath $ManifestPath).ProviderPath
[xml]$xml = Get-Content $ManifestPath -Raw
if ($xml.OuterXml -match 'localhost') { throw "$ManifestPath references localhost - use the production build from Deploy-AdvAddinWeb.ps1." }

if ($AssignOnly -and -not $ProductId) { throw '-AssignOnly needs -ProductId of the registered add-in.' }
$operation = if ($AssignOnly) { "Assign only, ProductId $ProductId" } elseif ($ProductId) { "Update ProductId $ProductId" } else { 'Register new add-in' }
$assignment = if ($Members) { $Members -join ', ' } else { 'everyone' }
Write-Information "$operation from $ManifestPath (Id $($xml.OfficeApp.Id), version $($xml.OfficeApp.Version), hosts $($xml.OfficeApp.Hosts.Host.Name -join ', '))" -InformationAction Continue
Write-Information "Assign to: $assignment" -InformationAction Continue
if (-not $Execute) { Write-Warning 'Plan only - nothing registered. Re-run with -Execute (tenant admin).'; return }

try {
    Import-Module O365CentralizedAddInDeployment -ErrorAction Stop
    Connect-OrganizationAddInService
    if ($AssignOnly) {
        Write-Information 'AssignOnly: manifest not re-uploaded.' -InformationAction Continue
    }
    elseif ($ProductId) {
        # Centralized Deployment rejects an update whose <Version> is not higher than the registered one.
        Set-OrganizationAddIn -ProductId $ProductId -ManifestPath $ManifestPath -Locale $locale
    }
    else {
        $ProductId = (New-OrganizationAddIn -ManifestPath $ManifestPath -Locale $locale).ProductId
        Write-Information "Registered ProductId $ProductId - pass it as -ProductId for later updates." -InformationAction Continue
    }
    if ($Members) {
        Set-OrganizationAddInAssignments -ProductId $ProductId -Add -Members $Members
    }
    else {
        Set-OrganizationAddInAssignments -ProductId $ProductId -AssignToEveryone $true
    }
    Get-OrganizationAddIn -ProductId $ProductId
}
catch {
    Write-Error "Publish failed: $($_.Exception.Message)"
}
