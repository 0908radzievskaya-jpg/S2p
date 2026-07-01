param(
    [string]$RootPath = "C:\Users\User\Documents\Создание КП на входящих",
    [string]$ArchivePath = "C:\Users\User\Documents\Создание КП на входящих\archive",
    [string]$MemoryPath = "C:\Users\User\.codex\automations\automation\memory.md",
    [string]$SortedRootPath = "C:\Users\User\Documents\Создание КП на входящих\sorted_mail",
    [string]$UnprocessedRootPath = "C:\Users\User\Documents\Создание КП на входящих\unprocessed",
    [int]$UnprocessedRetentionDays = 7
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

if (-not (Test-Path -LiteralPath $RootPath)) {
    $RootPath = Split-Path -Parent $PSCommandPath
}
if (-not (Test-Path -LiteralPath $ArchivePath)) {
    $ArchivePath = Join-Path $RootPath "archive"
}
if (-not [System.IO.Path]::IsPathRooted($SortedRootPath)) {
    $SortedRootPath = Join-Path $RootPath $SortedRootPath
}
if (-not [System.IO.Path]::IsPathRooted($UnprocessedRootPath)) {
    $UnprocessedRootPath = Join-Path $RootPath $UnprocessedRootPath
}

$headers = @(
    "Processed At",
    "Mailbox",
    "Message UID",
    "Subject",
    "Object",
    "Customer/Sender",
    "Contacts",
    "Address/Location",
    "Work Type",
    "Stage",
    "Deadline",
    "Documentation Scope",
    "Sections",
    "Source Files",
    "Links",
    "Key Technical Parameters",
    "Volumes",
    "Requirements",
    "Constraints",
    "Missing Information",
    "Risks",
    "Next Action Recommendation",
    "Data Sufficiency Status",
    "Materials Folder"
)

function Fix-Text {
    param([AllowNull()][string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) {
        return ""
    }
    try {
        $bytes = [System.Text.Encoding]::GetEncoding(1251).GetBytes($Text)
        $fixed = [System.Text.Encoding]::UTF8.GetString($bytes)
        if ($fixed -match '[A-Za-zА-Яа-я]') {
            return $fixed
        }
    } catch {
    }
    return $Text
}

function Get-LastProcessedUtc {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) {
        return $null
    }
    function Parse-UtcValue {
        param([string]$Value)
        $dto = [datetimeoffset]::MinValue
        if ([datetimeoffset]::TryParse($Value, [cultureinfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::RoundtripKind, [ref]$dto)) {
            return $dto.UtcDateTime
        }
        return $null
    }

    $raw = Get-Content -LiteralPath $Path -Raw -Encoding UTF8
    $match = [regex]::Match($raw, 'LastProcessedFileTimeUtc:\s*(.+)')
    if ($match.Success) {
        $value = $match.Groups[1].Value.Trim()
        $parsed = Parse-UtcValue -Value $value
        if ($parsed) {
            return $parsed
        }
    }

    $lastRunMatch = [regex]::Match($raw, 'Last run:\s*(.+)')
    if ($lastRunMatch.Success) {
        $parsed = Parse-UtcValue -Value $lastRunMatch.Groups[1].Value.Trim()
        if ($parsed) {
            return $parsed
        }
    }

    return $null
}

function Update-MemoryFile {
    param(
        [string]$Path,
        [datetime]$LastProcessedUtc,
        [datetime]$RunStarted,
        [int]$ProcessedCount,
        [string]$ReportPath
    )

    $memoryDir = Split-Path -Parent $Path
    if ($memoryDir) {
        New-Item -ItemType Directory -Path $memoryDir -Force | Out-Null
    }

    $lines = @(
        "Last run: $($RunStarted.ToUniversalTime().ToString('o'))",
        "Run time: local archive processing",
        "LastProcessedFileTimeUtc: $($LastProcessedUtc.ToString('o'))",
        "",
        "Summary:",
        "- Processed folders: $ProcessedCount",
        "- Report: $ReportPath"
    )

    Set-Content -LiteralPath $Path -Value ($lines -join "`r`n") -Encoding UTF8
}

function ConvertTo-CellRef {
    param([int]$Column, [int]$Row)
    $letters = ""
    $n = $Column
    while ($n -gt 0) {
        $rem = ($n - 1) % 26
        $letters = [char](65 + $rem) + $letters
        $n = [math]::Floor(($n - 1) / 26)
    }
    return "$letters$Row"
}

function Escape-XmlText {
    param([AllowNull()][string]$Value)
    if ([string]::IsNullOrEmpty($Value)) {
        return ""
    }
    return [System.Security.SecurityElement]::Escape($Value)
}

function Get-DirectoryMaxWriteTimeUtc {
    param([string]$DirectoryPath)
    $files = Get-ChildItem -LiteralPath $DirectoryPath -Recurse -File
    if (-not $files) {
        return [datetime]::MinValue
    }
    return ($files | Measure-Object -Property LastWriteTimeUtc -Maximum).Maximum
}

function Get-TenderObjectName {
    param(
        [string]$Subject,
        [string]$MessageText
    )
    $pattern = '\d+\s+"([^"]+)"'
    $match = [regex]::Match($MessageText, $pattern)
    if ($match.Success) {
        return $match.Groups[1].Value.Trim()
    }
    return $Subject
}

function Get-DeadlineFromText {
    param([string]$Text)
    $match = [regex]::Match($Text, '(\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2})')
    if ($match.Success) {
        return $match.Groups[1].Value + " MSK"
    }
    return ""
}

function Get-SufficiencyMeta {
    param(
        [string]$Sender,
        [string]$Subject,
        [string]$MissingInfo,
        [string]$Deadline,
        [string[]]$DownloadedFiles,
        [string[]]$Links
    )

    $senderLower = $Sender.ToLowerInvariant()
    $subjectLower = $Subject.ToLowerInvariant()
    $linkBlob = ($Links -join "`n").ToLowerInvariant()

    if ($senderLower -match 'metro|meatinfo' -or $subjectLower -match 'metro|meatinfo') {
        return @{
            Status = "insufficient"
            Risks = "Message does not contain design input data; high risk of wasting effort on a non-project lead."
            Recommendation = "Exclude from cost estimation workflow and keep only as informational correspondence."
        }
    }

    if ($linkBlob -match 'partner\.samolet\.ru') {
        $riskParts = @("Tender invitation contains only a high-level notice and lacks a full design package.")
        if ($MissingInfo) {
            $riskParts += "Missing: $MissingInfo."
        }
        if ($Deadline) {
            $riskParts += "Participation response deadline: $Deadline."
        }
        if (-not $DownloadedFiles -or $DownloadedFiles.Count -eq 0) {
            $riskParts += "No linked files were saved locally."
        }
        return @{
            Status = "partially sufficient"
            Risks = ($riskParts -join " ")
            Recommendation = "Review tender relevance and request TOR, source files, address, stage, quantities, and commercial submission rules."
        }
    }

    if ([string]::IsNullOrWhiteSpace($MissingInfo)) {
        return @{
            Status = "sufficient"
            Risks = ""
            Recommendation = "Inputs look adequate for preparing a draft estimate/commercial offer."
        }
    }

    return @{
        Status = "partially sufficient"
        Risks = "Input data is incomplete: $MissingInfo."
        Recommendation = "Clarify missing details before preparing the estimate/commercial offer."
    }
}

function Get-StatusFolderName {
    param([string]$Status)
    switch ($Status.ToLowerInvariant()) {
        "sufficient" { return "sufficient" }
        "partially sufficient" { return "partially_sufficient" }
        "insufficient" { return "insufficient" }
        default { return "unclassified" }
    }
}

function Sync-SortedMessageFolder {
    param(
        [string]$SourceDir,
        [string]$ArchiveRoot,
        [string]$SortedRoot,
        [string]$Status
    )

    $archiveRootResolved = [System.IO.Path]::GetFullPath($ArchiveRoot)
    $sourceResolved = [System.IO.Path]::GetFullPath($SourceDir)
    $sortedRootResolved = [System.IO.Path]::GetFullPath($SortedRoot)

    if (-not $sourceResolved.StartsWith($archiveRootResolved, [System.StringComparison]::OrdinalIgnoreCase)) {
        return ""
    }

    $relativePath = $sourceResolved.Substring($archiveRootResolved.Length).TrimStart('\')
    foreach ($statusFolder in @("sufficient", "partially_sufficient", "insufficient", "unclassified")) {
        $existingTarget = Join-Path (Join-Path $sortedRootResolved $statusFolder) $relativePath
        if (Test-Path -LiteralPath $existingTarget) {
            $existingResolved = [System.IO.Path]::GetFullPath($existingTarget)
            if ($existingResolved.StartsWith($sortedRootResolved, [System.StringComparison]::OrdinalIgnoreCase)) {
                Remove-Item -LiteralPath $existingResolved -Recurse -Force
            }
        }
    }

    $statusFolderName = Get-StatusFolderName -Status $Status
    $targetDir = Join-Path (Join-Path $sortedRootResolved $statusFolderName) $relativePath
    $targetParent = Split-Path -Parent $targetDir
    try {
        if (-not (Test-Path -LiteralPath $targetParent)) {
            New-Item -ItemType Directory -Path $targetParent -Force | Out-Null
        }
        if (Test-Path -LiteralPath $targetDir) {
            Remove-Item -LiteralPath $targetDir -Recurse -Force
        }
        Copy-Item -LiteralPath $sourceResolved -Destination $targetDir -Recurse -Force
        return $targetDir
    } catch {
        return ""
    }
}

function Remove-StaleUnprocessedDownloads {
    param(
        [string]$RootPath,
        [int]$RetentionDays
    )

    $removed = New-Object System.Collections.Generic.List[string]
    if (-not (Test-Path -LiteralPath $RootPath)) {
        return $removed
    }

    $cutoffUtc = (Get-Date).ToUniversalTime().AddDays(-1 * $RetentionDays)
    $downloadDirs = Get-ChildItem -LiteralPath $RootPath -Recurse -Directory |
        Where-Object { $_.Name -in @("downloads", "downloads_refreshed") }

    foreach ($dir in $downloadDirs) {
        $files = Get-ChildItem -LiteralPath $dir.FullName -Recurse -File
        foreach ($file in $files) {
            if ($file.LastWriteTimeUtc -le $cutoffUtc) {
                Remove-Item -LiteralPath $file.FullName -Force
                $removed.Add($file.FullName)
            }
        }

        Get-ChildItem -LiteralPath $dir.FullName -Recurse -Directory |
            Sort-Object FullName -Descending |
            ForEach-Object {
                if (-not (Get-ChildItem -LiteralPath $_.FullName -Force | Select-Object -First 1)) {
                    Remove-Item -LiteralPath $_.FullName -Force
                }
            }

        if (-not (Get-ChildItem -LiteralPath $dir.FullName -Force | Select-Object -First 1)) {
            Remove-Item -LiteralPath $dir.FullName -Force
        }
    }

    return $removed
}

function New-XlsxReport {
    param(
        [string]$OutputPath,
        [object[]]$Rows
    )

    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem

    $sheetRows = New-Object System.Collections.Generic.List[string]

    $headerCells = for ($i = 0; $i -lt $headers.Count; $i++) {
        $ref = ConvertTo-CellRef -Column ($i + 1) -Row 1
        "<c r=`"$ref`" s=`"1`" t=`"inlineStr`"><is><t>$(Escape-XmlText $headers[$i])</t></is></c>"
    }
    $headerXml = $headerCells -join ""
    $sheetRows.Add("<row r=`"1`">$headerXml</row>")

    $rowIndex = 2
    foreach ($row in $Rows) {
        $cells = for ($i = 0; $i -lt $headers.Count; $i++) {
            $ref = ConvertTo-CellRef -Column ($i + 1) -Row $rowIndex
            $value = [string]$row[$headers[$i]]
            "<c r=`"$ref`" s=`"2`" t=`"inlineStr`"><is><t xml:space=`"preserve`">$(Escape-XmlText $value)</t></is></c>"
        }
        $cellXml = $cells -join ""
        $sheetRows.Add("<row r=`"$rowIndex`">$cellXml</row>")
        $rowIndex++
    }

    $lastColumnRef = ConvertTo-CellRef -Column $headers.Count -Row ([math]::Max($Rows.Count + 1, 1))
    $dimension = "A1:$lastColumnRef"
    $sheetData = $sheetRows -join ""

    $contentTypes = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>
'@

    $rels = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>
'@

    $workbook = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Report" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>
'@

    $workbookRels = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>
'@

    $styles = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="2">
    <font><sz val="11"/><name val="Calibri"/></font>
    <font><b/><sz val="11"/><name val="Calibri"/></font>
  </fonts>
  <fills count="2">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFD9EAF7"/><bgColor indexed="64"/></patternFill></fill>
  </fills>
  <borders count="1">
    <border><left/><right/><top/><bottom/><diagonal/></border>
  </borders>
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="3">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="1" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment wrapText="1" vertical="top"/></xf>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment wrapText="1" vertical="top"/></xf>
  </cellXfs>
  <cellStyles count="1">
    <cellStyle name="Normal" xfId="0" builtinId="0"/>
  </cellStyles>
</styleSheet>
'@

    $sheet = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="$dimension"/>
  <sheetViews><sheetView workbookViewId="0"/></sheetViews>
  <sheetFormatPr defaultRowHeight="18"/>
  <cols>
    <col min="1" max="4" width="18" customWidth="1"/>
    <col min="5" max="5" width="28" customWidth="1"/>
    <col min="6" max="8" width="24" customWidth="1"/>
    <col min="9" max="12" width="20" customWidth="1"/>
    <col min="13" max="15" width="30" customWidth="1"/>
    <col min="16" max="24" width="35" customWidth="1"/>
  </cols>
  <sheetData>$sheetData</sheetData>
</worksheet>
"@

    $created = (Get-Date).ToUniversalTime().ToString("s") + "Z"
    $core = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:creator>Codex</dc:creator>
  <cp:lastModifiedBy>Codex</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">$created</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">$created</dcterms:modified>
  <dc:title>Incoming KP report</dc:title>
</cp:coreProperties>
"@

    $app = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>Codex</Application>
</Properties>
'@

    if (Test-Path -LiteralPath $OutputPath) {
        Remove-Item -LiteralPath $OutputPath -Force
    }

    $zip = [System.IO.Compression.ZipFile]::Open($OutputPath, [System.IO.Compression.ZipArchiveMode]::Create)
    try {
        $entries = @{
            "[Content_Types].xml" = $contentTypes
            "_rels/.rels" = $rels
            "xl/workbook.xml" = $workbook
            "xl/_rels/workbook.xml.rels" = $workbookRels
            "xl/worksheets/sheet1.xml" = $sheet
            "xl/styles.xml" = $styles
            "docProps/core.xml" = $core
            "docProps/app.xml" = $app
        }

        foreach ($entryName in $entries.Keys) {
            $entry = $zip.CreateEntry($entryName)
            $writer = New-Object System.IO.StreamWriter($entry.Open(), [System.Text.UTF8Encoding]::new($false))
            try {
                $writer.Write($entries[$entryName])
            } finally {
                $writer.Dispose()
            }
        }
    } finally {
        $zip.Dispose()
    }
}

$runStarted = Get-Date
$lastProcessedUtc = Get-LastProcessedUtc -Path $MemoryPath
$removedUnprocessedDownloads = Remove-StaleUnprocessedDownloads -RootPath $UnprocessedRootPath -RetentionDays $UnprocessedRetentionDays
$analysisFiles = Get-ChildItem -LiteralPath $ArchivePath -Recurse -Filter analysis.json -File | Sort-Object FullName
$latestDateFolderPath = $null

if (-not $lastProcessedUtc) {
    $latestDateFolder = Get-ChildItem -LiteralPath $ArchivePath -Directory |
        Where-Object { $_.Name -match '^\d{4}-\d{2}-\d{2}$' } |
        Sort-Object Name -Descending |
        Select-Object -First 1
    if ($latestDateFolder) {
        $latestDateFolderPath = $latestDateFolder.FullName
    }
}

$selected = foreach ($file in $analysisFiles) {
    $messageDir = $file.Directory.FullName
    $maxUtc = Get-DirectoryMaxWriteTimeUtc -DirectoryPath $messageDir
    if ($lastProcessedUtc) {
        if ($maxUtc -gt $lastProcessedUtc) {
            [pscustomobject]@{ File = $file; MaxUtc = $maxUtc }
        }
    } elseif ($latestDateFolderPath -and $file.FullName.StartsWith($latestDateFolderPath, [System.StringComparison]::OrdinalIgnoreCase)) {
        [pscustomobject]@{ File = $file; MaxUtc = $maxUtc }
    }
}

$reportRows = New-Object System.Collections.Generic.List[hashtable]
$processedFolders = New-Object System.Collections.Generic.List[string]
$downloadFailures = New-Object System.Collections.Generic.List[string]
$sortedFolders = New-Object System.Collections.Generic.List[string]

foreach ($entry in $selected) {
    $analysisPath = $entry.File.FullName
    $messageDir = Split-Path -Parent $analysisPath
    $analysis = Get-Content -LiteralPath $analysisPath -Raw -Encoding UTF8 | ConvertFrom-Json
    $messagePath = Join-Path $messageDir "message.txt"
    $messageTextRaw = if (Test-Path -LiteralPath $messagePath) { Get-Content -LiteralPath $messagePath -Raw -Encoding UTF8 } else { "" }
    $messageText = Fix-Text $messageTextRaw

    $fields = $analysis.fields
    $subject = Fix-Text ([string]$analysis.subject)
    $sender = Fix-Text ([string]$analysis.sender)
    $customer = Fix-Text ([string]$fields.customer)
    $contacts = Fix-Text ([string]$fields.contacts)
    $address = Fix-Text ([string]$fields.address)
    $workType = Fix-Text ([string]$fields.work_type)
    $stage = Fix-Text ([string]$fields.stage)
    $requirements = Fix-Text ([string]$fields.requirements)
    $constraints = Fix-Text ([string]$fields.constraints)
    $volumes = Fix-Text ([string]$fields.volumes)
    $techParams = Fix-Text ([string]$fields.tech_params)
    $missingInfo = Fix-Text ([string]$fields.missing_info)

    $deadline = Fix-Text ([string]$fields.deadline)
    if ([string]::IsNullOrWhiteSpace($deadline)) {
        $deadline = Get-DeadlineFromText -Text ($requirements + "`n" + $messageText)
    }

    $downloadedFiles = @()
    if ($analysis.downloaded_files) {
        $downloadedFiles = @($analysis.downloaded_files | ForEach-Object { [string]$_ })
    }

    $attachments = @()
    if ($analysis.attachments) {
        $attachments = @($analysis.attachments | ForEach-Object { [string]$_ })
    }

    $links = @()
    if ($analysis.links) {
        $links = @($analysis.links | ForEach-Object { [string]$_ })
    }

    $objectName = $subject
    if (($links -join "`n") -match 'partner\.samolet\.ru') {
        $objectName = Get-TenderObjectName -Subject $subject -MessageText $messageText
    } elseif (-not [string]::IsNullOrWhiteSpace([string]$fields.object)) {
        $objectName = Fix-Text ([string]$fields.object)
    }

    $sourceFiles = New-Object System.Collections.Generic.List[string]
    foreach ($path in @(
        (Join-Path $messageDir "message.eml"),
        (Join-Path $messageDir "message.txt"),
        $analysisPath
    ) + $attachments + $downloadedFiles) {
        if ($path -and (Test-Path -LiteralPath $path)) {
            $sourceFiles.Add($path)
        }
    }

    if ($links.Count -gt 0 -and $downloadedFiles.Count -eq 0) {
        $downloadFailures.Add("$subject :: linked materials were not saved locally")
    }

    $sections = @()
    if ($analysis.sections) {
        $sections = @($analysis.sections | ForEach-Object { Fix-Text ([string]$_) })
    }

    $meta = Get-SufficiencyMeta -Sender $sender -Subject $subject -MissingInfo $missingInfo -Deadline $deadline -DownloadedFiles $downloadedFiles -Links $links
    $isRelevant = $false
    if ($null -ne $analysis.is_relevant) {
        $isRelevant = [System.Convert]::ToBoolean($analysis.is_relevant)
    } elseif ($fields -and $null -ne $fields.is_relevant) {
        $isRelevant = ([string]$fields.is_relevant).ToLowerInvariant() -eq "yes"
    }
    if ($isRelevant) {
        $sortedPath = Sync-SortedMessageFolder -SourceDir $messageDir -ArchiveRoot $ArchivePath -SortedRoot $SortedRootPath -Status ([string]$meta.Status)
        if ($sortedPath) {
            $sortedFolders.Add($sortedPath)
        }
    }

    $docScopeParts = @()
    if ($sections.Count -gt 0) {
        $docScopeParts += $sections
    }
    if ($attachments.Count -gt 0) {
        $docScopeParts += "Attachments: " + (($attachments | ForEach-Object { Split-Path -Leaf $_ }) -join ", ")
    }
    if ($downloadedFiles.Count -gt 0) {
        $docScopeParts += "Downloaded: " + (($downloadedFiles | ForEach-Object { Split-Path -Leaf $_ }) -join ", ")
    }

    $reportRows.Add(@{
        "Processed At" = $runStarted.ToString("yyyy-MM-dd HH:mm:ss")
        "Mailbox" = Fix-Text ([string]$analysis.mailbox)
        "Message UID" = [string]$analysis.uid
        "Subject" = $subject
        "Object" = $objectName
        "Customer/Sender" = if ($customer) { $customer } else { $sender }
        "Contacts" = $contacts
        "Address/Location" = $address
        "Work Type" = $workType
        "Stage" = $stage
        "Deadline" = $deadline
        "Documentation Scope" = ($docScopeParts -join "; ")
        "Sections" = ($sections -join "; ")
        "Source Files" = ($sourceFiles -join "`n")
        "Links" = ($links -join "`n")
        "Key Technical Parameters" = $techParams
        "Volumes" = $volumes
        "Requirements" = $requirements
        "Constraints" = $constraints
        "Missing Information" = $missingInfo
        "Risks" = [string]$meta.Risks
        "Next Action Recommendation" = [string]$meta.Recommendation
        "Data Sufficiency Status" = [string]$meta.Status
        "Materials Folder" = $messageDir
    })
    $processedFolders.Add($messageDir)
}

$reportPath = Join-Path $RootPath ("run-report-" + $runStarted.ToString("yyyy-MM-dd_HH-mm-ss") + ".xlsx")
New-XlsxReport -OutputPath $reportPath -Rows $reportRows

$maxProcessedUtc = if ($selected) {
    ($selected | Measure-Object -Property MaxUtc -Maximum).Maximum
} else {
    $lastProcessedUtc
}

[pscustomobject]@{
    report_path = $reportPath
    processed_count = $reportRows.Count
    processed_folders = @($processedFolders)
    sorted_folders = @($sortedFolders)
    failed_downloads = @($downloadFailures)
    removed_unprocessed_downloads = @($removedUnprocessedDownloads)
    last_processed_file_time_utc = if ($maxProcessedUtc) { $maxProcessedUtc.ToString("o") } else { "" }
} | ConvertTo-Json -Depth 6

if ($maxProcessedUtc) {
    Update-MemoryFile -Path $MemoryPath -LastProcessedUtc $maxProcessedUtc -RunStarted $runStarted -ProcessedCount $reportRows.Count -ReportPath $reportPath
}
