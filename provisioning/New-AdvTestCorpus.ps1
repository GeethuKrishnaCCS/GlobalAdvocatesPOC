<#
.SYNOPSIS
    Generates and uploads a mixed test corpus (docx, xlsx, pdf, eml, txt) to a test matter site's Documents library.
.DESCRIPTION
    Files are built locally without Office: minimal OpenXML packages for docx/xlsx, a hand-written
    multi-page PDF, plain MIME for eml. Every file name starts with the testMarker (ADVTest_).
    A share of the files is uploaded a second time so they carry version history (v1.0 -> v2.0).
    pptx is not generated (no Office automation); add real samples by hand when needed.
.PARAMETER Alias
    Matter site alias (/sites/<Alias>).
.PARAMETER Count
    Number of files to generate.
.PARAMETER VersionedShare
    Fraction of files uploaded twice to create a second version.
.PARAMETER Execute
    Upload. Without it files are only generated locally.
.EXAMPLE
    ./New-AdvTestCorpus.ps1 -Count 300 -Execute
.NOTES
    use-when: need a realistic mixed-type document set with versions on the AdvocateDocs test matter site
    domain: tenant-sp
    safety: test-writes
    deps: PnP.PowerShell 3.x, pwsh 7.4+, Repos/AdvocateDocs-POC/.project.json
    tenant-aware: yes
#>
[CmdletBinding()]
param(
    [ValidatePattern('^ADVTest[A-Za-z0-9]+$')][string]$Alias = 'ADVTestMatter01',
    [ValidateRange(1, 5000)][int]$Count = 300,
    [ValidateRange(0.0, 1.0)][double]$VersionedShare = 0.1,
    [string]$OutDir = (Join-Path ([IO.Path]::GetTempPath()) 'advtest-corpus'),
    # Re-upload only these extensions (e.g. after fixing a generator); default = all.
    [ValidateSet('docx', 'xlsx', 'pdf', 'eml', 'txt')][string[]]$UploadTypes = @('docx', 'xlsx', 'pdf', 'eml', 'txt'),
    [switch]$Execute
)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression, System.IO.Compression.FileSystem

$library = 'ADVTestDocuments'
$marker = 'ADVTest'
$types = @('docx', 'xlsx', 'pdf', 'eml', 'txt')
$docTypes = @('Pleading', 'Affidavit', 'Correspondence', 'Evidence', 'Order', 'Contract')
$words = -split 'court matter petition hearing evidence witness counsel judgment appeal order claim party respondent applicant affidavit exhibit section clause tribunal bench filing notice relief interim decree statement record schedule annexure'
$rng = [Random]::new(42)

function Get-Text([int]$wordCount) {
    $sb = [Text.StringBuilder]::new()
    for ($i = 0; $i -lt $wordCount; $i++) {
        [void]$sb.Append($words[$rng.Next($words.Count)]).Append($(if ($i % 14 -eq 13) { ".`n" } else { ' ' }))
    }
    $sb.ToString()
}

function New-Zip([string]$path, [hashtable]$entries) {
    if (Test-Path $path) { Remove-Item $path }
    $zip = [IO.Compression.ZipFile]::Open($path, 'Create')
    try {
        foreach ($name in $entries.Keys) {
            $w = [IO.StreamWriter]::new($zip.CreateEntry($name).Open(), [Text.UTF8Encoding]::new($false))
            try { $w.Write($entries[$name]) } finally { $w.Dispose() }
        }
    } finally { $zip.Dispose() }
}

function Esc([string]$s) { [Security.SecurityElement]::Escape($s) }

function New-Docx([string]$path, [string]$title, [string]$body) {
    $paras = ($body -split "`n" | ForEach-Object { "<w:p><w:r><w:t xml:space=`"preserve`">$(Esc $_)</w:t></w:r></w:p>" }) -join ''
    New-Zip $path @{
        '[Content_Types].xml' = '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'
        '_rels/.rels'         = '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'
        'word/document.xml'   = "<?xml version=`"1.0`" encoding=`"UTF-8`"?><w:document xmlns:w=`"http://schemas.openxmlformats.org/wordprocessingml/2006/main`"><w:body><w:p><w:r><w:rPr><w:b/><w:sz w:val=`"32`"/></w:rPr><w:t>$(Esc $title)</w:t></w:r></w:p>$paras</w:body></w:document>"
    }
}

function New-Xlsx([string]$path, [string]$title, [int]$rows) {
    $sb = [Text.StringBuilder]::new()
    [void]$sb.Append("<row r=`"1`"><c r=`"A1`" t=`"inlineStr`"><is><t>$(Esc $title)</t></is></c></row>")
    for ($r = 2; $r -le $rows; $r++) {
        [void]$sb.Append("<row r=`"$r`"><c r=`"A$r`" t=`"inlineStr`"><is><t>$($words[$rng.Next($words.Count)])</t></is></c><c r=`"B$r`"><v>$($rng.Next(1, 100000))</v></c><c r=`"C$r`"><v>$($rng.NextDouble())</v></c></row>")
    }
    New-Zip $path @{
        '[Content_Types].xml'        = '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>'
        '_rels/.rels'                = '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'
        'xl/workbook.xml'            = '<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Schedule" sheetId="1" r:id="rId1"/></sheets></workbook>'
        'xl/_rels/workbook.xml.rels' = '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>'
        'xl/worksheets/sheet1.xml'   = "<?xml version=`"1.0`" encoding=`"UTF-8`"?><worksheet xmlns=`"http://schemas.openxmlformats.org/spreadsheetml/2006/main`"><sheetData>$($sb.ToString())</sheetData></worksheet>"
    }
}

function New-Pdf([string]$path, [string]$title, [int]$pages) {
    # Minimal valid PDF: one Helvetica font, one content stream per page, xref offsets computed.
    $objs = [Collections.Generic.List[string]]::new()
    $kids = (1..$pages | ForEach-Object { "$(3 + 2 * $_) 0 R" }) -join ' '
    $objs.Add('<< /Type /Catalog /Pages 2 0 R >>')
    $objs.Add("<< /Type /Pages /Kids [$kids] /Count $pages >>")
    $objs.Add('<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>')
    for ($p = 1; $p -le $pages; $p++) {
        $lines = (Get-Text 300) -split "`n" | Select-Object -First 40
        $text = "BT /F1 16 Tf 50 790 Td ($title - page $p) Tj /F1 10 Tf 0 -24 Td 14 TL " + (($lines | ForEach-Object { "($($_ -replace '[()\\]', '')) '" }) -join ' ') + ' ET'
        $objs.Add("<< /Length $($text.Length) >>`nstream`n$text`nendstream")
        $objs.Add("<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595 842] /Resources << /Font << /F1 3 0 R >> >> /Contents $(2 + 2 * $p) 0 R >>")
    }
    $sb = [Text.StringBuilder]::new("%PDF-1.4`n")
    $offsets = @()
    for ($i = 0; $i -lt $objs.Count; $i++) {
        $offsets += $sb.Length
        [void]$sb.Append("$($i + 1) 0 obj`n$($objs[$i])`nendobj`n")
    }
    $xref = $sb.Length
    [void]$sb.Append("xref`n0 $($objs.Count + 1)`n0000000000 65535 f `n")
    foreach ($o in $offsets) { [void]$sb.Append(("{0:D10} 00000 n `n" -f $o)) }
    [void]$sb.Append("trailer`n<< /Size $($objs.Count + 1) /Root 1 0 R >>`nstartxref`n$xref`n%%EOF")
    [IO.File]::WriteAllText($path, $sb.ToString(), [Text.Encoding]::ASCII)
}

function New-Eml([string]$path, [string]$title, [string]$body) {
    $date = (Get-Date).AddDays(-$rng.Next(1, 900)).ToString('r')
    @"
From: counsel@advtest.example
To: client@advtest.example
Subject: $title
Date: $date
MIME-Version: 1.0
Content-Type: text/plain; charset=utf-8

$body
"@ | Set-Content -Path $path -Encoding utf8
}

# --- Generate ---
New-Item -ItemType Directory -Force $OutDir | Out-Null
$files = for ($i = 1; $i -le $Count; $i++) {
    $ext = $types[($i - 1) % $types.Count]
    $docType = $docTypes[$rng.Next($docTypes.Count)]
    $title = '{0}_{1}_{2:D4}' -f $marker, $docType, $i
    $path = Join-Path $OutDir "$title.$ext"
    switch ($ext) {
        'docx' { New-Docx $path $title (Get-Text $rng.Next(15000, 60000)) }
        'xlsx' { New-Xlsx $path $title $rng.Next(2000, 8000) }
        'pdf'  { New-Pdf $path $title $rng.Next(3, 25) }
        'eml'  { New-Eml $path $title (Get-Text $rng.Next(400, 3000)) }
        'txt'  { Set-Content -Path $path -Value (Get-Text $rng.Next(15000, 60000)) -Encoding utf8 }
    }
    [pscustomobject]@{ Path = $path; DocType = $docType; DocDate = (Get-Date).AddDays(-$rng.Next(1, 900)).Date }
}
$totalMb = [math]::Round((($files | ForEach-Object { (Get-Item $_.Path).Length }) | Measure-Object -Sum).Sum / 1MB, 1)
Write-Information "Generated $($files.Count) files ($totalMb MB) in $OutDir" -InformationAction Continue
if (-not $Execute) { Write-Warning 'Generated locally only. Re-run with -Execute to upload.'; return }

# --- Upload ---
$config = Get-Content (Join-Path $PSScriptRoot '../.project.json') -Raw | ConvertFrom-Json
$siteUrl = "$($config.tenants.dev.url)/sites/$Alias"
$conn = Connect-PnPOnline -Url $siteUrl -ClientId $config.entraClientId -Thumbprint $config.entraCertThumbprint `
    -Tenant $config.tenants.dev.tenantId -ReturnConnection
$versioned = [int][math]::Floor($Count * $VersionedShare)
$isPartial = $UploadTypes.Count -lt $types.Count
$files = @($files | Where-Object { $UploadTypes -contains [IO.Path]::GetExtension($_.Path).TrimStart('.') })
if ($isPartial) { $versioned = 0 }
$n = 0
foreach ($f in $files) {
    $n++
    Write-Progress -Activity 'Uploading corpus' -Status (Split-Path $f.Path -Leaf) -PercentComplete ($n / $files.Count * 100)
    $values = @{ DocType = $f.DocType; DocDate = $f.DocDate }
    [void](Add-PnPFile -Path $f.Path -Folder $library -Values $values -Connection $conn)
}
# Second version for the first N files (same content + a change marker line is enough for history).
foreach ($f in $files | Select-Object -First $versioned) {
    if ($f.Path -match '\.(txt|eml)$') { Add-Content -Path $f.Path -Value "`nRevised $(Get-Date -Format s)" }
    [void](Add-PnPFile -Path $f.Path -Folder $library -Values @{ DocType = $f.DocType } -Connection $conn)
}
Write-Progress -Activity 'Uploading corpus' -Completed
Write-Information "Uploaded $($files.Count) files to $siteUrl/$library; $versioned with a second version." -InformationAction Continue
