<#
.SYNOPSIS
    Preflight for the Office add-in dev loop: everything an E2E run needs, each with PASS/FAIL and the exact fix.
.DESCRIPTION
    Read-only. Checks, in order:
      1. Dev server answers https://localhost:3000 and serves manifest.xml.
      2. manifest.xml validates (office-addin-manifest).
      3. Entra app (docsApiApp in .project.json): SPA redirect URIs for NAA, admin consent for the add-in's Graph scopes.
         Read with the project certificate app (app-only); needs Application.Read.All on it, else reported as manual.
      4. Playwright session (.auth/office-web.json) exists and its Microsoft sign-in cookie has not expired.
      5. e2e/.env has a test document URL per host.
    Exit code 1 when any check fails, so callers can gate an E2E run on it.
.PARAMETER SkipEntra
    Skip check 3 (no tenant call).
.EXAMPLE
    ./scripts/Test-AddinDevReady.ps1
.NOTES
    use-when: before an add-in E2E run; add-in pane won't sign in; ribbon tab missing in Office on the web; Playwright session expired
    domain: tenant-sp
    safety: read-only
    deps: pwsh 7+, PnP.PowerShell 3.x, node/npm (office-addin-manifest), Repos/AdvocateDocs-POC/.project.json
    tenant-aware: yes
#>
[CmdletBinding()]
param([switch]$SkipEntra)
$ErrorActionPreference = 'Stop'

$addinRoot = Split-Path $PSScriptRoot -Parent
# Main checkout's config: worktree branches can predate ids added later (e.g. docsApiApp).
$projectsRoot = if ($env:PROJECTS_ROOT) { $env:PROJECTS_ROOT } else { 'D:\Projects' }
$config = Get-Content (Join-Path $projectsRoot 'Repos/AdvocateDocs-POC/.project.json') -Raw | ConvertFrom-Json
$devOrigin = 'https://localhost:3000'
$requiredRedirects = @("$devOrigin/taskpane.html", 'brk-multihub://localhost:3000')
$requiredScopes = @('Files.ReadWrite.All', 'Sites.Read.All', 'Mail.ReadWrite') # mirrors GRAPH_SCOPES in src/shared/config.ts
$docEnvKeys = @('ADV_E2E_DOC_WORD', 'ADV_E2E_DOC_EXCEL', 'ADV_E2E_DOC_POWERPOINT')
$results = [System.Collections.Generic.List[object]]::new()

function Add-Result([string]$Check, [bool]$Pass, [string]$Detail, [string]$Fix = '') {
    $results.Add([pscustomobject]@{ Check = $Check; Result = $(if ($Pass) { 'PASS' } else { 'FAIL' }); Detail = $Detail; Fix = $(if ($Pass) { '' } else { $Fix }) })
}

# 1. Dev server
try {
    $manifest = Invoke-WebRequest "$devOrigin/manifest.xml" -SkipCertificateCheck -TimeoutSec 10
    Add-Result 'Dev server' ($manifest.StatusCode -eq 200) "$devOrigin/manifest.xml $($manifest.StatusCode)"
}
catch {
    Add-Result 'Dev server' $false $_.Exception.Message 'npm run dev-server (Playwright starts it too: npm run e2e)'
}

# 2. Manifest
Push-Location $addinRoot
try {
    $out = npx --no-install office-addin-manifest validate manifest.xml 2>&1 | Out-String
    Add-Result 'Manifest valid' ($out -match 'The manifest is valid') 'manifest.xml' ($out.Trim() -split "`n" | Select-Object -Last 5 | Out-String)
}
finally { Pop-Location }

# 3. Entra app registration + consent
if (-not $SkipEntra) {
    $app = $config.tenants.dev.docsApiApp
    if (-not $app.objectId) { throw 'tenants.dev.docsApiApp missing in .project.json' }
    try {
        Connect-PnPOnline -Url $config.tenants.dev.solutionSite -ClientId $config.entraClientId -Thumbprint $config.entraCertThumbprint `
            -Tenant $config.tenants.dev.tenantId -ErrorAction Stop
        $headers = @{ Authorization = "Bearer $(Get-PnPAccessToken -ResourceTypeName Graph)" }
        $graph = 'https://graph.microsoft.com/v1.0'

        $reg = Invoke-RestMethod "$graph/applications/$($app.objectId)?`$select=spa" -Headers $headers
        $missing = @($requiredRedirects | Where-Object { $_ -notin $reg.spa.redirectUris })
        Add-Result 'Entra SPA redirect URIs' ($missing.Count -eq 0) "registered: $($reg.spa.redirectUris -join ', ')" `
            "Entra > App registrations > $($app.displayName) ($($app.clientId)) > Authentication > Single-page application: add $($missing -join ', ')"

        $graphSp = Invoke-RestMethod "$graph/servicePrincipals(appId='00000003-0000-0000-c000-000000000000')?`$select=id" -Headers $headers
        $grants = (Invoke-RestMethod "$graph/oauth2PermissionGrants?`$filter=clientId eq '$($app.servicePrincipalId)' and resourceId eq '$($graphSp.id)' and consentType eq 'AllPrincipals'" -Headers $headers).value
        $granted = @($grants.scope -split ' ' | Where-Object { $_ })
        $notConsented = @($requiredScopes | Where-Object { $_ -notin $granted })
        Add-Result 'Graph admin consent' ($notConsented.Count -eq 0) "consented: $($granted -join ' ')" `
            "Entra > App registrations > $($app.displayName) > API permissions: add delegated Graph $($notConsented -join ', '), then Grant admin consent for $($config.tenants.dev.name)"
    }
    catch {
        Add-Result 'Entra registration' $false $_.Exception.Message `
            "Grant Application.Read.All (application) + admin consent to cert app $($config.entraClientId), or check redirect URIs/consent by hand"
    }
}

# 4. Playwright session
$statePath = Join-Path $addinRoot '.auth/office-web.json'
if (Test-Path $statePath) {
    $now = [DateTimeOffset]::UtcNow.ToUnixTimeSeconds()
    $auth = (Get-Content $statePath -Raw | ConvertFrom-Json).cookies |
        Where-Object { $_.domain -match 'login\.microsoftonline\.com' -and $_.name -like 'ESTSAUTH*' }
    $live = @($auth | Where-Object { $_.expires -eq -1 -or $_.expires -gt $now })
    $expiry = ($live | Where-Object expires -gt 0 | Measure-Object expires -Maximum).Maximum
    Add-Result 'Playwright session' ($live.Count -gt 0) $(if ($expiry) { "sign-in cookie until $([DateTimeOffset]::FromUnixTimeSeconds($expiry).LocalDateTime)" } else { 'no persistent sign-in cookie' }) `
        'npm run e2e:login (sign in as dev06 in the opened Edge window)'
}
else {
    Add-Result 'Playwright session' $false 'no .auth/office-web.json' 'npm run e2e:login (sign in as dev06 in the opened Edge window)'
}

# 5. Test documents
$envPath = Join-Path $addinRoot 'e2e/.env'
$envText = if (Test-Path $envPath) { Get-Content $envPath -Raw } else { '' }
foreach ($key in $docEnvKeys) {
    $set = $envText -match "(?m)^$key=https://\S+"
    Add-Result "Test doc $key" $set $(if ($set) { 'set' } else { 'missing' }) 'Open the ADVTest file in Office on the web, copy the address bar into e2e/.env (see e2e/.env.example)'
}

foreach ($r in $results) {
    Write-Host ("{0}  {1} - {2}" -f $r.Result, $r.Check, $r.Detail) -ForegroundColor $(if ($r.Result -eq 'PASS') { 'Green' } else { 'Red' })
    if ($r.Fix) { Write-Host "      fix: $($r.Fix.Trim())" }
}
if ($results.Result -contains 'FAIL') { exit 1 }
