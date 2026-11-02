<#
.SYNOPSIS
    Generates one sample per extra file type (pptx, ODF, rtf, html, md, csv, json, xml, log, images, zip, damaged pdf)
    and uploads them to a test matter site's Documents library.
.DESCRIPTION
    Complements New-AdvTestCorpus.ps1 (docx/xlsx/pdf/eml/txt) so bundle builds and sends can be tested on every
    kind of file a matter holds. Built without Office: minimal OpenXML for pptx, minimal ODF for odt/ods/odp,
    System.Drawing for images. Some types are expected to become placeholder pages in a bundle (zip, svg,
    damaged pdf); that is part of the test. Every name starts with ADVTest_Sample_. The pptx is uploaded twice
    so it has a previous version for "send previous version as PDF".
    Not generated (need Office or proprietary binary formats): doc, xls, ppt, msg, vsdx.
.PARAMETER Alias
    Matter site alias (/sites/<Alias>).
.PARAMETER Execute
    Upload. Without it files are only generated locally.
.EXAMPLE
    ./New-AdvTestTypeSamples.ps1 -Execute
.NOTES
    use-when: need one file of every supported/unsupported type on the AdvocateDocs test matter site
    domain: tenant-sp
    safety: test-writes
    deps: PnP.PowerShell 3.x, pwsh 7.4+ on Windows (System.Drawing), Repos/AdvocateDocs-POC/.project.json
    tenant-aware: yes
#>
[CmdletBinding()]
param(
    [ValidatePattern('^ADVTest[A-Za-z0-9]+$')][string]$Alias = 'ADVTestMatter01',
    [string]$OutDir = (Join-Path ([IO.Path]::GetTempPath()) 'advtest-type-samples'),
    # Upload only these extensions or file-name wildcards (e.g. eml, *EML_PDF*); default = all.
    # Re-uploading adds a version to existing files.
    [string[]]$Only,
    [switch]$Execute
)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression, System.IO.Compression.FileSystem, System.Drawing

$library = 'ADVTestDocuments'
$prefix = 'ADVTest_Sample'
$para = 'The applicant filed the petition before the tribunal. Counsel for the respondent sought an adjournment to file a reply affidavit with annexures. The bench directed both parties to complete pleadings and list the matter for hearing.'

function Esc([string]$s) { [Security.SecurityElement]::Escape($s) }

# Ordered entries; with -OdfMimetype the first entry is stored uncompressed, as ODF requires.
function New-Package([string]$path, [System.Collections.Specialized.OrderedDictionary]$entries, [switch]$OdfMimetype) {
    if (Test-Path $path) { Remove-Item $path }
    $zip = [IO.Compression.ZipFile]::Open($path, 'Create')
    try {
        $first = $true
        foreach ($name in $entries.Keys) {
            $level = if ($first -and $OdfMimetype) { 'NoCompression' } else { 'Optimal' }
            $first = $false
            $w = [IO.StreamWriter]::new($zip.CreateEntry($name, $level).Open(), [Text.UTF8Encoding]::new($false))
            try { $w.Write($entries[$name]) } finally { $w.Dispose() }
        }
    } finally { $zip.Dispose() }
}

function New-Pptx([string]$path, [string]$title, [int]$slides) {
    $ns = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"'
    $rel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    $ct = 'application/vnd.openxmlformats-officedocument.presentationml'
    $emptyTree = '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>'
    function Shape([int]$id, [string]$text, [int]$y, [int]$h, [int]$size) {
        "<p:sp><p:nvSpPr><p:cNvPr id=`"$id`" name=`"Text $id`"/><p:cNvSpPr txBox=`"1`"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x=`"457200`" y=`"$y`"/><a:ext cx=`"8229600`" cy=`"$h`"/></a:xfrm><a:prstGeom prst=`"rect`"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr wrap=`"square`"/><a:lstStyle/><a:p><a:r><a:rPr lang=`"en-GB`" sz=`"$size`"/><a:t>$(Esc $text)</a:t></a:r></a:p></p:txBody></p:sp>"
    }
    $theme = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="ADVTest"><a:themeElements><a:clrScheme name="ADVTest"><a:dk1><a:srgbClr val="000000"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="1F3864"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2><a:accent1><a:srgbClr val="4472C4"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2><a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4><a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6><a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme><a:fontScheme name="ADVTest"><a:majorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme><a:fmtScheme name="ADVTest"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst><a:lnStyleLst><a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements></a:theme>'
    $entries = [ordered]@{}
    $slideOverrides = (1..$slides | ForEach-Object { "<Override PartName=`"/ppt/slides/slide$_.xml`" ContentType=`"$ct.slide+xml`"/>" }) -join ''
    $entries['[Content_Types].xml'] = "<?xml version=`"1.0`" encoding=`"UTF-8`" standalone=`"yes`"?><Types xmlns=`"http://schemas.openxmlformats.org/package/2006/content-types`"><Default Extension=`"rels`" ContentType=`"application/vnd.openxmlformats-package.relationships+xml`"/><Default Extension=`"xml`" ContentType=`"application/xml`"/><Override PartName=`"/ppt/presentation.xml`" ContentType=`"$ct.presentation.main+xml`"/><Override PartName=`"/ppt/slideMasters/slideMaster1.xml`" ContentType=`"$ct.slideMaster+xml`"/><Override PartName=`"/ppt/slideLayouts/slideLayout1.xml`" ContentType=`"$ct.slideLayout+xml`"/><Override PartName=`"/ppt/theme/theme1.xml`" ContentType=`"application/vnd.openxmlformats-officedocument.theme+xml`"/>$slideOverrides</Types>"
    $entries['_rels/.rels'] = "<?xml version=`"1.0`" encoding=`"UTF-8`" standalone=`"yes`"?><Relationships xmlns=`"http://schemas.openxmlformats.org/package/2006/relationships`"><Relationship Id=`"rId1`" Type=`"$rel/officeDocument`" Target=`"ppt/presentation.xml`"/></Relationships>"
    $sldIds = (1..$slides | ForEach-Object { "<p:sldId id=`"$(255 + $_)`" r:id=`"rId$(1 + $_)`"/>" }) -join ''
    $entries['ppt/presentation.xml'] = "<?xml version=`"1.0`" encoding=`"UTF-8`" standalone=`"yes`"?><p:presentation $ns><p:sldMasterIdLst><p:sldMasterId id=`"2147483648`" r:id=`"rId1`"/></p:sldMasterIdLst><p:sldIdLst>$sldIds</p:sldIdLst><p:sldSz cx=`"9144000`" cy=`"6858000`" type=`"screen4x3`"/><p:notesSz cx=`"6858000`" cy=`"9144000`"/></p:presentation>"
    $presRels = "<Relationship Id=`"rId1`" Type=`"$rel/slideMaster`" Target=`"slideMasters/slideMaster1.xml`"/>" +
        ((1..$slides | ForEach-Object { "<Relationship Id=`"rId$(1 + $_)`" Type=`"$rel/slide`" Target=`"slides/slide$_.xml`"/>" }) -join '') +
        "<Relationship Id=`"rId$(2 + $slides)`" Type=`"$rel/theme`" Target=`"theme/theme1.xml`"/>"
    $entries['ppt/_rels/presentation.xml.rels'] = "<?xml version=`"1.0`" encoding=`"UTF-8`" standalone=`"yes`"?><Relationships xmlns=`"http://schemas.openxmlformats.org/package/2006/relationships`">$presRels</Relationships>"
    $entries['ppt/slideMasters/slideMaster1.xml'] = "<?xml version=`"1.0`" encoding=`"UTF-8`" standalone=`"yes`"?><p:sldMaster $ns><p:cSld><p:spTree>$emptyTree</p:spTree></p:cSld><p:clrMap bg1=`"lt1`" tx1=`"dk1`" bg2=`"lt2`" tx2=`"dk2`" accent1=`"accent1`" accent2=`"accent2`" accent3=`"accent3`" accent4=`"accent4`" accent5=`"accent5`" accent6=`"accent6`" hlink=`"hlink`" folHlink=`"folHlink`"/><p:sldLayoutIdLst><p:sldLayoutId id=`"2147483649`" r:id=`"rId1`"/></p:sldLayoutIdLst></p:sldMaster>"
    $entries['ppt/slideMasters/_rels/slideMaster1.xml.rels'] = "<?xml version=`"1.0`" encoding=`"UTF-8`" standalone=`"yes`"?><Relationships xmlns=`"http://schemas.openxmlformats.org/package/2006/relationships`"><Relationship Id=`"rId1`" Type=`"$rel/slideLayout`" Target=`"../slideLayouts/slideLayout1.xml`"/><Relationship Id=`"rId2`" Type=`"$rel/theme`" Target=`"../theme/theme1.xml`"/></Relationships>"
    $entries['ppt/slideLayouts/slideLayout1.xml'] = "<?xml version=`"1.0`" encoding=`"UTF-8`" standalone=`"yes`"?><p:sldLayout $ns type=`"blank`"><p:cSld name=`"Blank`"><p:spTree>$emptyTree</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>"
    $entries['ppt/slideLayouts/_rels/slideLayout1.xml.rels'] = "<?xml version=`"1.0`" encoding=`"UTF-8`" standalone=`"yes`"?><Relationships xmlns=`"http://schemas.openxmlformats.org/package/2006/relationships`"><Relationship Id=`"rId1`" Type=`"$rel/slideMaster`" Target=`"../slideMasters/slideMaster1.xml`"/></Relationships>"
    $entries['ppt/theme/theme1.xml'] = $theme
    for ($s = 1; $s -le $slides; $s++) {
        $tree = $emptyTree + (Shape 2 "$title - slide $s of $slides" 457200 914400 3200) + (Shape 3 $para 1600200 4114800 2000)
        $entries["ppt/slides/slide$s.xml"] = "<?xml version=`"1.0`" encoding=`"UTF-8`" standalone=`"yes`"?><p:sld $ns><p:cSld><p:spTree>$tree</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>"
        $entries["ppt/slides/_rels/slide$s.xml.rels"] = "<?xml version=`"1.0`" encoding=`"UTF-8`" standalone=`"yes`"?><Relationships xmlns=`"http://schemas.openxmlformats.org/package/2006/relationships`"><Relationship Id=`"rId1`" Type=`"$rel/slideLayout`" Target=`"../slideLayouts/slideLayout1.xml`"/></Relationships>"
    }
    New-Package $path $entries
}

function New-Odf([string]$path, [string]$kind, [string]$title) {
    $mime = @{ text = 'application/vnd.oasis.opendocument.text'; spreadsheet = 'application/vnd.oasis.opendocument.spreadsheet'; presentation = 'application/vnd.oasis.opendocument.presentation' }[$kind]
    $ns = 'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" office:version="1.2"'
    $body = switch ($kind) {
        'text' { "<office:text><text:h text:outline-level=`"1`">$(Esc $title)</text:h>" + ((1..6 | ForEach-Object { "<text:p>$(Esc $para)</text:p>" }) -join '') + '</office:text>' }
        'spreadsheet' {
            $rows = (1..40 | ForEach-Object { "<table:table-row><table:table-cell office:value-type=`"string`"><text:p>Exhibit $_</text:p></table:table-cell><table:table-cell office:value-type=`"float`" office:value=`"$($_ * 125)`"><text:p>$($_ * 125)</text:p></table:table-cell></table:table-row>" }) -join ''
            "<office:spreadsheet><table:table table:name=`"Schedule`"><table:table-row><table:table-cell office:value-type=`"string`"><text:p>$(Esc $title)</text:p></table:table-cell></table:table-row>$rows</table:table></office:spreadsheet>"
        }
        'presentation' {
            '<office:presentation>' + ((1..3 | ForEach-Object { "<draw:page draw:name=`"Slide $_`"><draw:frame svg:x=`"1cm`" svg:y=`"1cm`" svg:width=`"24cm`" svg:height=`"12cm`"><draw:text-box><text:p>$(Esc "$title - slide $_")</text:p><text:p>$(Esc $para)</text:p></draw:text-box></draw:frame></draw:page>" }) -join '') + '</office:presentation>'
        }
    }
    $entries = [ordered]@{}
    $entries['mimetype'] = $mime
    $entries['META-INF/manifest.xml'] = "<?xml version=`"1.0`" encoding=`"UTF-8`"?><manifest:manifest xmlns:manifest=`"urn:oasis:names:tc:opendocument:xmlns:manifest:1.0`" manifest:version=`"1.2`"><manifest:file-entry manifest:full-path=`"/`" manifest:media-type=`"$mime`"/><manifest:file-entry manifest:full-path=`"content.xml`" manifest:media-type=`"text/xml`"/></manifest:manifest>"
    $entries['content.xml'] = "<?xml version=`"1.0`" encoding=`"UTF-8`"?><office:document-content $ns><office:body>$body</office:body></office:document-content>"
    New-Package $path $entries -OdfMimetype
}

function New-Image([string]$path, [string]$title, [Drawing.Imaging.ImageFormat]$format) {
    $bmp = [Drawing.Bitmap]::new(1200, 800)
    $g = [Drawing.Graphics]::FromImage($bmp)
    try {
        $g.Clear([Drawing.Color]::White)
        $g.FillRectangle([Drawing.Brushes]::SteelBlue, 0, 0, 1200, 120)
        $g.DrawString($title, [Drawing.Font]::new('Arial', 36, [Drawing.FontStyle]::Bold), [Drawing.Brushes]::White, 40, 30)
        $g.DrawString($para, [Drawing.Font]::new('Arial', 20), [Drawing.Brushes]::Black, [Drawing.RectangleF]::new(40, 180, 1120, 400))
        $g.DrawEllipse([Drawing.Pen]::new([Drawing.Color]::DarkOrange, 8), 850, 520, 260, 220)
        $bmp.Save($path, $format)
    } finally { $g.Dispose(); $bmp.Dispose() }
}

# Minimal valid PDF (Helvetica, one text block per page) for email attachments.
function New-SimplePdf([string]$path, [string]$title, [int]$pages) {
    $objs = [Collections.Generic.List[string]]::new()
    $objs.Add('<< /Type /Catalog /Pages 2 0 R >>')
    $objs.Add("<< /Type /Pages /Kids [$((1..$pages | ForEach-Object { "$(3 + 2 * $_) 0 R" }) -join ' ')] /Count $pages >>")
    $objs.Add('<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>')
    for ($p = 1; $p -le $pages; $p++) {
        $text = "BT /F1 18 Tf 50 780 Td ($title - page $p of $pages) Tj /F1 11 Tf 0 -30 Td ($($para.Substring(0, 90))) Tj ET"
        $objs.Add("<< /Length $($text.Length) >>`nstream`n$text`nendstream")
        $objs.Add("<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595 842] /Resources << /Font << /F1 3 0 R >> >> /Contents $(2 + 2 * $p) 0 R >>")
    }
    $sb = [Text.StringBuilder]::new("%PDF-1.4`n")
    $offsets = for ($i = 0; $i -lt $objs.Count; $i++) { $sb.Length; [void]$sb.Append("$($i + 1) 0 obj`n$($objs[$i])`nendobj`n") }
    $xref = $sb.Length
    [void]$sb.Append("xref`n0 $($objs.Count + 1)`n0000000000 65535 f `n")
    foreach ($o in $offsets) { [void]$sb.Append(("{0:D10} 00000 n `n" -f $o)) }
    [void]$sb.Append("trailer`n<< /Size $($objs.Count + 1) /Root 1 0 R >>`nstartxref`n$xref`n%%EOF")
    [IO.File]::WriteAllText($path, $sb.ToString(), [Text.Encoding]::ASCII)
}

# Small signature image, as Outlook embeds it (image00N.png).
function New-Logo([string]$path, [int]$width, [int]$height, [string]$text) {
    $bmp = [Drawing.Bitmap]::new($width, $height)
    $g = [Drawing.Graphics]::FromImage($bmp)
    try {
        $g.Clear([Drawing.Color]::FromArgb(17, 45, 78))
        $g.DrawString($text, [Drawing.Font]::new('Georgia', [Math]::Max(8, $height / 3), [Drawing.FontStyle]::Bold), [Drawing.Brushes]::Goldenrod, 6, $height / 5)
        $bmp.Save($path, [Drawing.Imaging.ImageFormat]::Png)
    } finally { $g.Dispose(); $bmp.Dispose() }
}

# MIME email with base64 attachments. -Html gives an HTML body; -Inline maps Content-ID -> image file it references.
# -InlinePng is shorthand for one inline photo with a default HTML body.
function New-EmlWithAttachments([string]$path, [string]$subject, [string[]]$attachments, [string]$InlinePng,
    [string]$Html, [System.Collections.Specialized.OrderedDictionary]$Inline) {
    $nl = "`r`n"
    function B64([string]$file) { [Convert]::ToBase64String([IO.File]::ReadAllBytes($file), 'InsertLineBreaks') }
    function Part([string]$file, [string]$disposition, [string]$cid) {
        $name = Split-Path $file -Leaf
        $type = @{ '.png' = 'image/png'; '.jpg' = 'image/jpeg'; '.pdf' = 'application/pdf' }[[IO.Path]::GetExtension($file)]
        $cidHeader = if ($cid) { "Content-ID: <$cid>$nl" } else { '' }
        "Content-Type: $type; name=`"$name`"${nl}Content-Transfer-Encoding: base64${nl}Content-Disposition: $disposition; filename=`"$name`"$nl$cidHeader$nl$(B64 $file)$nl"
    }
    if ($InlinePng) {
        $Html = "<html><body><p>Dear client,</p><p>$(Esc $para)</p><p>Site photograph below:</p><img src=`"cid:advtest-inline-1`" width=`"600`"/><p>Regards,<br/>Counsel</p></body></html>"
        $Inline = [ordered]@{ 'advtest-inline-1' = $InlinePng }
    }
    $mixed = 'ADVTest-mixed-boundary'
    $text = "Dear client,$nl$nl$para$nl${nl}Attached: $(($attachments | ForEach-Object { Split-Path $_ -Leaf }) -join ', ').$nl${nl}Regards,${nl}Counsel"
    $body = if ($Html) {
        $related = 'ADVTest-related-boundary'
        $inlineParts = ($Inline.Keys | ForEach-Object { "--$related$nl$(Part $Inline[$_] 'inline' $_)" }) -join ''
        "Content-Type: multipart/related; boundary=`"$related`"$nl$nl--$related${nl}Content-Type: text/html; charset=utf-8$nl$nl$Html$nl$inlineParts--$related--$nl"
    } else {
        "Content-Type: text/plain; charset=utf-8$nl$nl$text$nl"
    }
    $parts = (@($body) + ($attachments | ForEach-Object { Part $_ 'attachment' $null })) | ForEach-Object { "--$mixed$nl$_" }
    $eml = "From: counsel@advtest.example${nl}To: client@advtest.example${nl}Cc: paralegal@advtest.example${nl}Subject: $subject${nl}Date: $((Get-Date).AddDays(-3).ToString('r'))${nl}MIME-Version: 1.0${nl}Content-Type: multipart/mixed; boundary=`"$mixed`"$nl$nl" +
        ($parts -join '') + "--$mixed--$nl"
    [IO.File]::WriteAllText($path, $eml, [Text.Encoding]::ASCII)
}

# --- Generate ---
New-Item -ItemType Directory -Force $OutDir | Out-Null
function P([string]$ext) { Join-Path $OutDir "${prefix}_$($ext.ToUpperInvariant()).$ext" }
function T([string]$ext) { "${prefix}_$($ext.ToUpperInvariant())" }

New-Pptx (P 'pptx') (T 'pptx') 4
New-Odf (P 'odt') 'text' (T 'odt')
New-Odf (P 'ods') 'spreadsheet' (T 'ods')
New-Odf (P 'odp') 'presentation' (T 'odp')
Set-Content (P 'rtf') -Encoding ascii -Value "{\rtf1\ansi\deff0{\fonttbl{\f0 Times New Roman;}}\f0\fs32\b $(T 'rtf')\b0\par\fs24 $para\par $para\par}"
Set-Content (P 'html') -Encoding utf8 -Value "<!doctype html><html><head><meta charset=`"utf-8`"><title>$(T 'html')</title></head><body><h1>$(T 'html')</h1><p>$para</p><table border=`"1`"><tr><th>Exhibit</th><th>Pages</th></tr><tr><td>A</td><td>12</td></tr><tr><td>B</td><td>4</td></tr></table></body></html>"
Set-Content (P 'md') -Encoding utf8 -Value "# $(T 'md')`n`n$para`n`n- Exhibit A`n- Exhibit B`n`n| Exhibit | Pages |`n|---|---|`n| A | 12 |`n| B | 4 |"
Set-Content (P 'csv') -Encoding utf8 -Value (@('Exhibit,Description,Pages') + (1..60 | ForEach-Object { "E$_,Annexure $_ to the affidavit,$($_ % 9 + 1)" }))
Set-Content (P 'json') -Encoding utf8 -Value (@{ title = (T 'json'); matter = 'ADVTestMatter01'; exhibits = @(1..10 | ForEach-Object { @{ id = "E$_"; pages = $_ } }) } | ConvertTo-Json -Depth 4)
$exhibits = (1..10 | ForEach-Object { '<exhibit id="E{0}" pages="{0}"/>' -f $_ }) -join ''
Set-Content (P 'xml') -Encoding utf8 -Value ('<?xml version="1.0" encoding="utf-8"?><matter id="ADVTestMatter01"><title>{0}</title>{1}</matter>' -f (T 'xml'), $exhibits)
Set-Content (P 'log') -Encoding utf8 -Value (1..80 | ForEach-Object { "2026-10-$('{0:D2}' -f ($_ % 28 + 1))T10:$('{0:D2}' -f ($_ % 60)):00 INFO filing $_ accepted by registry" })
Set-Content (P 'svg') -Encoding utf8 -Value "<svg xmlns=`"http://www.w3.org/2000/svg`" width=`"400`" height=`"200`"><rect width=`"400`" height=`"200`" fill=`"steelblue`"/><text x=`"20`" y=`"110`" font-size=`"28`" fill=`"white`">$(T 'svg')</text></svg>"
New-Image (P 'png') (T 'png') ([Drawing.Imaging.ImageFormat]::Png)
New-Image (P 'jpg') (T 'jpg') ([Drawing.Imaging.ImageFormat]::Jpeg)
New-Image (P 'gif') (T 'gif') ([Drawing.Imaging.ImageFormat]::Gif)
New-Image (P 'bmp') (T 'bmp') ([Drawing.Imaging.ImageFormat]::Bmp)
New-Image (P 'tiff') (T 'tiff') ([Drawing.Imaging.ImageFormat]::Tiff)
# Emails with image attachments (generated after the images they carry).
New-EmlWithAttachments (Join-Path $OutDir "${prefix}_EML_PNG.eml") 'ADVTest_Sample_EML_PNG - exhibit photo attached' @((P 'png'))
New-EmlWithAttachments (Join-Path $OutDir "${prefix}_EML_JPG_PNG.eml") 'ADVTest_Sample_EML_JPG_PNG - two images attached' @((P 'jpg'), (P 'png'))
New-EmlWithAttachments (Join-Path $OutDir "${prefix}_EML_INLINE.eml") 'ADVTest_Sample_EML_INLINE - inline photo plus attachment' @((P 'png')) -InlinePng (P 'jpg')

# Attachment files live in a subfolder (realistic names, not uploaded on their own).
$att = Join-Path $OutDir 'attachments'
New-Item -ItemType Directory -Force $att | Out-Null
New-SimplePdf (Join-Path $att 'Statement_of_Claim.pdf') 'Statement of Claim' 3
New-SimplePdf (Join-Path $att 'Invoice_INV-2041.pdf') 'Invoice INV-2041' 1
Copy-Item (P 'jpg') (Join-Path $att 'Site_Photo_1.jpg') -Force
Copy-Item (P 'png') (Join-Path $att 'Floor_Plan.png') -Force
New-Logo (Join-Path $att 'image001.png') 220 60 'ADVTest Legal LLP'
New-Logo (Join-Path $att 'image002.png') 24 24 'in'

New-EmlWithAttachments (Join-Path $OutDir "${prefix}_EML_PDF.eml") 'ADVTest_Sample_EML_PDF - statement of claim attached' @((Join-Path $att 'Statement_of_Claim.pdf'))
New-EmlWithAttachments (Join-Path $OutDir "${prefix}_EML_PDF_IMAGES.eml") 'ADVTest_Sample_EML_PDF_IMAGES - claim, invoice and site photos' `
    @((Join-Path $att 'Statement_of_Claim.pdf'), (Join-Path $att 'Invoice_INV-2041.pdf'), (Join-Path $att 'Site_Photo_1.jpg'), (Join-Path $att 'Floor_Plan.png'))
# Outlook-style signature: logo and social icon are inline (cid) images, not exhibits; one real PDF attachment.
$signatureHtml = "<html><body><p>Dear client,</p><p>$(Esc $para)</p><p>Please find our invoice attached.</p>" +
    '<p>Kind regards,</p><table><tr><td><img src="cid:image001.png@01DB2F1A.5C3E8B10" width="220" height="60" alt="ADVTest Legal LLP"/></td></tr>' +
    '<tr><td><b>Priya Menon</b> | Senior Associate<br/>ADVTest Legal LLP | +91 80 0000 0000<br/>' +
    '<img src="cid:image002.png@01DB2F1A.5C3E8B10" width="24" height="24" alt="LinkedIn"/></td></tr></table></body></html>'
New-EmlWithAttachments (Join-Path $OutDir "${prefix}_EML_SIGNATURE.eml") 'ADVTest_Sample_EML_SIGNATURE - invoice (signature logo inline)' `
    @((Join-Path $att 'Invoice_INV-2041.pdf')) -Html $signatureHtml `
    -Inline ([ordered]@{ 'image001.png@01DB2F1A.5C3E8B10' = (Join-Path $att 'image001.png'); 'image002.png@01DB2F1A.5C3E8B10' = (Join-Path $att 'image002.png') })
$zipEntries = [ordered]@{ 'readme.txt' = $para }
New-Package (P 'zip') $zipEntries
# Truncated PDF: header only, no objects — must become a placeholder page, not fail the build.
[IO.File]::WriteAllText((Join-Path $OutDir "${prefix}_DAMAGED.pdf"), "%PDF-1.4`n1 0 obj`n<< /Type /Catalog", [Text.Encoding]::ASCII)

$files = Get-ChildItem $OutDir -File | Where-Object Name -like "${prefix}_*"
Write-Information "Generated $($files.Count) samples in ${OutDir}: $(($files.Extension | Sort-Object -Unique) -join ' ')" -InformationAction Continue
if (-not $Execute) { Write-Warning 'Generated locally only. Re-run with -Execute to upload.'; return }

# --- Upload ---
$config = Get-Content (Join-Path $PSScriptRoot '../.project.json') -Raw | ConvertFrom-Json
$siteUrl = "$($config.tenants.dev.url)/sites/$Alias"
$conn = Connect-PnPOnline -Url $siteUrl -ClientId $config.entraClientId -Thumbprint $config.entraCertThumbprint `
    -Tenant $config.tenants.dev.tenantId -ReturnConnection
if ($Only) { $files = @($files | Where-Object { $f = $_; $Only | Where-Object { $f.Extension.TrimStart('.') -eq $_ -or $f.Name -like $_ } }) }
foreach ($f in $files) {
    [void](Add-PnPFile -Path $f.FullName -Folder $library -Values @{ DocType = 'Evidence'; DocDate = (Get-Date).Date } -Connection $conn)
}
# Second version of the pptx (more slides), for "send previous version" tests.
if (-not $Only -or $Only -contains 'pptx') {
    New-Pptx (P 'pptx') (T 'pptx') 6
    [void](Add-PnPFile -Path (P 'pptx') -Folder $library -Values @{ DocType = 'Evidence' } -Connection $conn)
}
Write-Information "Uploaded $($files.Count) samples to $siteUrl/${library}: $(($files.Name) -join ', ')" -InformationAction Continue
