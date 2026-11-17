<#
.SYNOPSIS
    Grants the AdvocateDocs Function's managed identity Graph Sites.Selected and write access to one matter site.
.DESCRIPTION
    Two tenant-level grants, run once per environment by an admin (not by automation):
      1. App roles Sites.Selected on Microsoft Graph and on SharePoint Online for the Function's managed identity
         (signed-in az admin). The SharePoint role is for REST column updates that add no version.
      2. Site permission 'write' for that identity on the matter site only (PnP, project certificate app).
    Sites.Selected gives no access anywhere until a site is granted explicitly, so the Function can only touch
    the sites listed here.
.PARAMETER Alias
    Matter site alias (/sites/<Alias>).
.PARAMETER Execute
    Perform the grants. Without it the script only shows what it would grant.
.EXAMPLE
    ./Grant-AdvFunctionAccess.ps1            # plan
    ./Grant-AdvFunctionAccess.ps1 -Execute   # grant (as a Global / Privileged Role admin signed in to az)
.NOTES
    use-when: bundle build fails with 403 from Graph; new matter site needs the Function's access
    domain: azure
    safety: tenant-write-gate
    deps: az CLI signed in as a tenant admin, PnP.PowerShell 3.x, Repos/AdvocateDocs-POC/.project.json
    tenant-aware: yes
#>
[CmdletBinding()]
param(
    [ValidatePattern('^ADVTest[A-Za-z0-9]+$')][string]$Alias = 'ADVTestMatter01',
    [switch]$Execute
)
$ErrorActionPreference = 'Stop'
$graphAppId = '00000003-0000-0000-c000-000000000000'
$sharePointAppId = '00000003-0000-0ff1-ce00-000000000000'

# az rest goes through cmd.exe, which mangles Graph URLs with (), $ and quotes; call Graph directly with az's token.
$script:GraphToken = $null
function Invoke-Graph([string]$Method, [string]$Url, $Body) {
    if (-not $script:GraphToken) {
        $script:GraphToken = az account get-access-token --resource-type ms-graph --query accessToken -o tsv 2>$null
        if (-not $script:GraphToken) { throw 'No Graph token from az. Sign in: az login --tenant <tenantId>.' }
    }
    $request = @{ Method = $Method; Uri = $Url; Headers = @{ Authorization = "Bearer $script:GraphToken" } }
    if ($Body) { $request.Body = ($Body | ConvertTo-Json -Depth 6); $request.ContentType = 'application/json' }
    return Invoke-RestMethod @request
}

try {
    $config = Get-Content (Join-Path $PSScriptRoot '../.project.json') -Raw | ConvertFrom-Json
    $azure = $config.tenants.dev.azure
    $principalId = (az functionapp identity show -g $azure.resourceGroup -n $azure.functionApp --query principalId -o tsv 2>$null)
    if (-not $principalId) { throw "Function $($azure.functionApp) has no managed identity." }

    $mi = Invoke-Graph GET "https://graph.microsoft.com/v1.0/servicePrincipals/$principalId`?`$select=id,appId,displayName"
    $assigned = (Invoke-Graph GET "https://graph.microsoft.com/v1.0/servicePrincipals/$principalId/appRoleAssignments").value
    # Sites.Selected on both APIs: Graph for files/lists, SharePoint REST for column updates without a new version.
    $roles = foreach ($resource in @(@{ Name = 'Microsoft Graph'; AppId = $graphAppId }, @{ Name = 'SharePoint Online'; AppId = $sharePointAppId })) {
        $sp = Invoke-Graph GET "https://graph.microsoft.com/v1.0/servicePrincipals(appId='$($resource.AppId)')?`$select=id,appRoles"
        $roleId = ($sp.appRoles | Where-Object value -eq 'Sites.Selected').id
        [pscustomobject]@{ Name = $resource.Name; ResourceId = $sp.id; RoleId = $roleId
            Assigned = [bool]($assigned | Where-Object { $_.resourceId -eq $sp.id -and $_.appRoleId -eq $roleId }) }
    }
    $siteUrl = "$($config.tenants.dev.url)/sites/$Alias"

    Write-Information "Managed identity: $($mi.displayName) (appId $($mi.appId))" -InformationAction Continue
    foreach ($role in $roles) {
        Write-Information "1. $($role.Name) Sites.Selected: $(if ($role.Assigned) { 'already assigned' } else { 'to assign' })" -InformationAction Continue
    }
    Write-Information "2. Site write on $siteUrl" -InformationAction Continue
    if (-not $Execute) { Write-Warning 'Plan only — nothing granted. Re-run with -Execute.'; return }

    foreach ($role in $roles | Where-Object { -not $_.Assigned }) {
        Invoke-Graph POST "https://graph.microsoft.com/v1.0/servicePrincipals/$principalId/appRoleAssignments" @{
            principalId = $principalId; resourceId = $role.ResourceId; appRoleId = $role.RoleId
        } | Out-Null
        Write-Information "Assigned $($role.Name) Sites.Selected." -InformationAction Continue
    }
    # Site grants need Graph Sites.FullControl.All, which neither az nor the automation app holds (by design):
    # an admin completes step 2 in Graph Explorer.
    $conn = Connect-PnPOnline -Url $siteUrl -ClientId $config.entraClientId -Thumbprint $config.entraCertThumbprint `
        -Tenant $config.tenants.dev.tenantId -ReturnConnection
    $graphSiteId = "$(([uri]$siteUrl).Host),$((Get-PnPSite -Includes Id -Connection $conn).Id),$((Get-PnPWeb -Includes Id -Connection $conn).Id)"
    $body = @{ roles = @('write'); grantedToIdentities = @(@{ application = @{ id = $mi.appId; displayName = $mi.displayName } }) } | ConvertTo-Json -Depth 5
    Write-Information @"
Step 2 (manual): Graph Explorer as a tenant admin, consent Sites.FullControl.All, then
  POST https://graph.microsoft.com/v1.0/sites/$graphSiteId/permissions
$body
"@ -InformationAction Continue
}
catch {
    Write-Error "Grant failed: $($_.Exception.Message)"
    throw
}
