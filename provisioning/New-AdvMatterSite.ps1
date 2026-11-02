<#
.SYNOPSIS
    Provisions a test matter site: Documents / Bundles libraries, BundleItems list, BundleManifests library, columns, indexes.
.DESCRIPTION
    One matter (case) = one SharePoint site. Creates the site and the four containers the DocActions
    Command Set and the bundle Function expect. Idempotent: existing site / lists / fields are skipped.
    All names carry the project testMarker (ADVTest) — SpoInternal refuses writes otherwise.
    Production matter sites will use a site template; this script is the POC equivalent.
.PARAMETER Alias
    Site URL alias (/sites/<Alias>). Must start with the testMarker.
.PARAMETER Owner
    Site owner UPN (required for app-only site creation).
.PARAMETER Execute
    Write. Without it the script only prints the plan.
.EXAMPLE
    ./New-AdvMatterSite.ps1 -Alias ADVTestMatter01 -Owner ccsdev@ccsdev01.onmicrosoft.com            # plan
    ./New-AdvMatterSite.ps1 -Alias ADVTestMatter01 -Owner ccsdev@ccsdev01.onmicrosoft.com -Execute   # write
.NOTES
    use-when: create a test matter site for the AdvocateDocs POC; reset bundle libraries/columns on a matter site
    domain: tenant-sp
    safety: test-writes
    deps: PnP.PowerShell 3.x, pwsh 7.4+, Scripts/SpoInternal.psm1, Repos/AdvocateDocs-POC/.project.json
    tenant-aware: yes
#>
[CmdletBinding()]
param(
    [ValidatePattern('^ADVTest[A-Za-z0-9]+$')][string]$Alias = 'ADVTestMatter01',
    [Parameter(Mandatory)][string]$Owner,
    [switch]$Execute
)
$ErrorActionPreference = 'Stop'
$projectsRoot = if ($env:PROJECTS_ROOT) { $env:PROJECTS_ROOT } else { 'D:\Projects' }
Import-Module (Join-Path $projectsRoot 'Scripts/SpoInternal.psm1') -Force

$project = 'AdvocateDocs-POC'
$tenantUrl = 'https://ccsdev01.sharepoint.com'
$siteUrl = "$tenantUrl/sites/$Alias"

# Container internal names (URL segments). The Command Set reads these from its properties.
$lists = @(
    @{ Name = 'ADVTestDocuments';       Template = 'DocumentLibrary'; Label = 'Documents (test)' }
    @{ Name = 'ADVTestBundles';         Template = 'DocumentLibrary'; Label = 'Bundles (test)' }
    @{ Name = 'ADVTestBundleItems';     Template = 'GenericList';     Label = 'Bundle items (test)' }
    @{ Name = 'ADVTestBundleManifests'; Template = 'DocumentLibrary'; Label = 'Bundle manifests (test)' }
    # One row per bundle (header); members live in ADVTestBundleItems keyed by BundleId.
    @{ Name = 'ADVTestBundleRegistry';  Template = 'GenericList';     Label = 'Bundle registry (test)' }
    # Scratch copies for "send previous version as PDF" (Graph converts only current files). Hidden, unversioned.
    @{ Name = 'ADVTestConversion';      Template = 'DocumentLibrary'; Label = 'Conversion scratch (test)'; Hidden = $true }
)

$fields = @(
    @{ List = 'ADVTestDocuments'; Name = 'DocType';  Type = 'Text';     Label = 'Document type'; Indexed = $true }
    @{ List = 'ADVTestDocuments'; Name = 'DocDate';  Type = 'DateTime'; Label = 'Document date'; Indexed = $true }
    @{ List = 'ADVTestDocuments'; Name = 'IsFinal';  Type = 'Boolean';  Label = 'Final';         Indexed = $true }

    @{ List = 'ADVTestBundles'; Name = 'BundleId';        Type = 'Text';   Label = 'Bundle ID';        Indexed = $true }
    @{ List = 'ADVTestBundles'; Name = 'ManifestVersion'; Type = 'Number'; Label = 'Manifest version' }
    @{ List = 'ADVTestBundles'; Name = 'BuildStatus';     Type = 'Choice'; Label = 'Build status'; Choices = @('Queued', 'Building', 'Ready', 'Failed') }
    @{ List = 'ADVTestBundles'; Name = 'PageCount';       Type = 'Number'; Label = 'Pages' }
    @{ List = 'ADVTestBundles'; Name = 'VolumeNo';        Type = 'Number'; Label = 'Volume' }

    @{ List = 'ADVTestBundleItems'; Name = 'BundleId';           Type = 'Text';     Label = 'Bundle ID'; Indexed = $true }
    @{ List = 'ADVTestBundleItems'; Name = 'SourceUniqueId';     Type = 'Text';     Label = 'Source document ID' }
    @{ List = 'ADVTestBundleItems'; Name = 'SourceVersionLabel'; Type = 'Text';     Label = 'Source version' }
    # Documents library item id — lets the build Function find the file through Graph.
    @{ List = 'ADVTestBundleItems'; Name = 'SourceItemId';       Type = 'Number';   Label = 'Source item ID' }
    @{ List = 'ADVTestBundleItems'; Name = 'SortOrder';          Type = 'Number';   Label = 'Order' }
    @{ List = 'ADVTestBundleItems'; Name = 'IndexTitle';         Type = 'Text';     Label = 'Index title' }
    @{ List = 'ADVTestBundleItems'; Name = 'IndexDate';          Type = 'DateTime'; Label = 'Index date' }

    @{ List = 'ADVTestBundleRegistry'; Name = 'BundleId';     Type = 'Text';   Label = 'Bundle ID'; Indexed = $true }
    @{ List = 'ADVTestBundleRegistry'; Name = 'BundleStatus'; Type = 'Choice'; Label = 'Status'; Choices = @('Draft', 'Built') }
    @{ List = 'ADVTestBundleRegistry'; Name = 'MemberCount';  Type = 'Number'; Label = 'Documents' }
    @{ List = 'ADVTestBundleRegistry'; Name = 'LatestPdfUrl'; Type = 'Text';   Label = 'Latest PDF' }

    @{ List = 'ADVTestBundleManifests'; Name = 'BundleId';     Type = 'Text';   Label = 'Bundle ID'; Indexed = $true }
    @{ List = 'ADVTestBundleManifests'; Name = 'BuildVersion'; Type = 'Number'; Label = 'Build version' }
)

$ledger = [System.Collections.Generic.List[object]]::new()

try {
    # Site — needs a tenant-level connection.
    Connect-SpoInternal -ProjectName $project -SiteUrl $tenantUrl
    $siteExists = $true
    try { [void](Get-SpoSite -Alias $Alias) } catch { $siteExists = $false }
    if ($siteExists) {
        Write-Information "Site $siteUrl exists — skipped." -InformationAction Continue
    } else {
        $ledger.Add((New-SpoSite -Alias $Alias -Type CommunicationSite -Label "Matter $Alias (test)" -Owner $Owner -Execute:$Execute))
        if (-not $Execute) { $ledger; Write-Warning 'Plan only. Lists/fields are planned after the site exists.'; return }
    }

    Connect-SpoInternal -ProjectName $project -SiteUrl $siteUrl
    $existingLists = @(Get-SpoList -IncludeHidden | ForEach-Object Name)
    foreach ($l in $lists) {
        if ($existingLists -contains $l.Name) { continue }
        # Manifests and scratch are system data — kept out of the quick launch.
        $onQuickLaunch = $l.Name -ne 'ADVTestBundleManifests' -and -not $l.Hidden
        $ledger.Add((New-SpoList -Name $l.Name -Template $l.Template -Label $l.Label -OnQuickLaunch:$onQuickLaunch -Execute:$Execute))
    }
    if (-not $Execute) { $ledger; return }

    foreach ($f in $fields) {
        $exists = $true
        try { [void](Resolve-SpoField -List $f.List -Name $f.Name) } catch { $exists = $false }
        if ($exists) { continue }
        $fieldArgs = @{ List = $f.List; Name = $f.Name; Type = $f.Type; Label = $f.Label; AddToDefaultView = $true; Execute = $true }
        if ($f.Indexed) { $fieldArgs.Indexed = $true }
        if ($f.Choices) { $fieldArgs.Choices = $f.Choices }
        $ledger.Add((New-SpoField @fieldArgs))
    }

    # Versioning: major versions on every library (bundle + document version history depend on it).
    # SpoInternal keeps its connection private, so open one for the list settings it doesn't wrap.
    $config = Get-Content (Join-Path $projectsRoot "Repos/$project/.project.json") -Raw | ConvertFrom-Json
    $conn = Connect-PnPOnline -Url $siteUrl -ClientId $config.entraClientId -Thumbprint $config.entraCertThumbprint `
        -Tenant $config.tenants.dev.tenantId -ReturnConnection
    foreach ($l in $lists | Where-Object { $_.Template -eq 'DocumentLibrary' -and -not $_.Hidden }) {
        $listObj = Resolve-SpoList -Name $l.Name
        Set-PnPList -Identity $listObj.Id -EnableVersioning $true -MajorVersions 500 -Connection $conn | Out-Null
        $ledger.Add([pscustomobject]@{ Operation = 'Set versioning (major, 500)'; Object = $l.Name; Mode = 'EXECUTED' })
    }
    foreach ($l in $lists | Where-Object Hidden) {
        $listObj = Resolve-SpoList -Name $l.Name
        Set-PnPList -Identity $listObj.Id -Hidden $true -EnableVersioning $false -Connection $conn | Out-Null
        $ledger.Add([pscustomobject]@{ Operation = 'Hide, versioning off'; Object = $l.Name; Mode = 'EXECUTED' })
    }
}
catch {
    Write-Error "Provisioning failed: $($_.Exception.Message)"
    throw
}
finally {
    $ledger | Format-Table Operation, Object, Mode -AutoSize | Out-String | Write-Information -InformationAction Continue
}
