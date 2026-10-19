#requires -Version 7.4.6
<#
.SYNOPSIS
Plans or uploads only locally remediated documents to the priority workbook's destinations.
.DESCRIPTION
Default is offline: no SharePoint connection or remote writes. -Preflight connects and backs
up existing destinations without uploading. -Apply performs the same complete preflight,
then overwrites files and restores Author, Editor, Created and Modified.
Existing target files are required. No folders, permissions, publication or retention
settings are changed. Use a new OutputFolder each run. A portable -PlanCsv lets the
server use only this script, the plan and -DocumentsFolder, without the workbook
or remediation history.
Metadata writes use explicitly UTC CSOM DateTimes to avoid server-local timezone
conversion. After an interrupted upload, -RestoreMetadataSnapshot performs a
separate metadata-only recovery from the original backup's .metadata.json file.
Do not combine recovery with -Apply or -Preflight.
Exact SHA-256 verification is the default. -AllowOfficeMetadataChanges permits
recognized metadata-only DOCX/XLSX/PPTX package differences and writes a part audit;
all other package parts and all PDF bytes must still match.
#>
[CmdletBinding()]
param(
    [string]$RepositoryRoot = (Split-Path -Parent $PSScriptRoot),
    [string]$DocumentsFolder,
    [string]$PriorityWorkbook,
    [string]$RemediationRunsFolder,
    [string]$PlanCsv,
    [string]$OutputFolder,
    [string]$ClientId,
    [string]$RestoreMetadataSnapshot,
    [switch]$AllowOfficeMetadataChanges,
    [switch]$Preflight,
    [switch]$Apply
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Read-PriorityWorkbook {
    param([string]$Path)
    $workbookStream = [IO.File]::Open($Path, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::ReadWrite)
    $zip = [IO.Compression.ZipArchive]::new($workbookStream, [IO.Compression.ZipArchiveMode]::Read)
    try {
        function Read-ZipXml([string]$Name) {
            $entry = $zip.GetEntry($Name)
            if ($null -eq $entry) { throw "Workbook part missing: $Name" }
            $stream = $entry.Open()
            try {
                $settings = [System.Xml.XmlReaderSettings]::new()
                $settings.DtdProcessing = [System.Xml.DtdProcessing]::Prohibit
                $settings.XmlResolver = $null
                $reader = [System.Xml.XmlReader]::Create($stream, $settings)
                try {
                    $xml = [System.Xml.XmlDocument]::new()
                    $xml.XmlResolver = $null
                    $xml.Load($reader)
                    return ,$xml
                } finally { $reader.Dispose() }
            } finally { $stream.Dispose() }
        }
        $strings = [System.Collections.Generic.List[string]]::new()
        if ($zip.GetEntry('xl/sharedStrings.xml')) {
            $xml = Read-ZipXml 'xl/sharedStrings.xml'
            foreach ($item in $xml.SelectNodes("//*[local-name()='si']")) {
                $strings.Add(($item.SelectNodes(".//*[local-name()='t']") |
                    ForEach-Object InnerText) -join '')
            }
        }
        $records = [System.Collections.Generic.List[object]]::new()
        foreach ($entry in $zip.Entries | Where-Object FullName -Match '^xl/worksheets/sheet\d+\.xml$') {
            $xml = Read-ZipXml $entry.FullName
            $headers = @{}
            foreach ($row in $xml.SelectNodes("//*[local-name()='sheetData']/*[local-name()='row']")) {
                $cells = @{}
                foreach ($cell in $row.SelectNodes("*[local-name()='c']")) {
                    $column = $cell.GetAttribute('r') -replace '\d', ''
                    $valueNode = $cell.SelectSingleNode("*[local-name()='v']")
                    $value = if ($null -ne $valueNode) { $valueNode.InnerText } else { '' }
                    switch ($cell.GetAttribute('t')) {
                        's' { $value = $strings[[int]$value] }
                        'inlineStr' {
                            $value = ($cell.SelectNodes(".//*[local-name()='t']") |
                                ForEach-Object InnerText) -join ''
                        }
                    }
                    $cells[$column] = $value
                }
                if ($headers.Count -eq 0) {
                    if ($cells.Values -contains 'DocumentFriendlyName' -and
                        $cells.Values -contains 'DocumentContainerURL') {
                        foreach ($key in $cells.Keys) { $headers[$cells[$key]] = $key }
                    } elseif (@($cells.Values | Where-Object { $_ }).Count) {
                        throw "Unexpected worksheet header in $($entry.FullName)."
                    }
                    continue
                }
                $name = [string]$cells[$headers['DocumentFriendlyName']]
                $container = [string]$cells[$headers['DocumentContainerURL']]
                if (-not $name -and -not $container) { continue }
                if (-not $name -or -not $container) {
                    throw "Incomplete priority row in $($entry.FullName), row $($row.GetAttribute('r'))."
                }
                $records.Add([pscustomobject]@{ Name = $name; Container = $container })
            }
        }
        if ($records.Count -eq 0) { throw 'No priority rows found in workbook.' }
        return $records.ToArray()
    } finally { $zip.Dispose(); $workbookStream.Dispose() }
}

function Resolve-UploadDestination {
    param([string]$Container, [string]$Name)
    if (-not $Name -or $Name.IndexOfAny([char[]]'\/') -ge 0 -or $Name -in '.', '..') {
        throw "Invalid document filename: $Name"
    }
    $uri = [uri]$Container
    if (-not $uri.IsAbsoluteUri -or $uri.Scheme -ne 'https' -or
        $uri.Authority -ne 'whallc.sharepoint.com' -or $uri.UserInfo -or $uri.Fragment) {
        throw "Untrusted destination URL: $Container"
    }
    $path = [uri]::UnescapeDataString($uri.AbsolutePath).TrimEnd('/')
    $parameters = @{}
    foreach ($pair in $uri.Query.TrimStart('?') -split '&') {
        if (-not $pair) { continue }
        $parts = $pair -split '=', 2
        if ($parts.Count -ne 2) { throw "Invalid query in container URL: $Container" }
        $key = [uri]::UnescapeDataString($parts[0])
        if ($parameters.ContainsKey($key)) { throw "Duplicate query key: $Container" }
        $parameters[$key] = [uri]::UnescapeDataString($parts[1].Replace('+', ' '))
    }
    if ($path -match '(?i)^(.*)/Forms/AllItems\.aspx$') {
        $library = $Matches[1]
        $path = $library
        $folderValues = @($parameters.Keys | Where-Object { $_ -in 'id', 'RootFolder' } |
            ForEach-Object { [string]$parameters[$_] } | Select-Object -Unique)
        if ($folderValues.Count -gt 1) { throw "Conflicting folder selectors: $Container" }
        if ($folderValues.Count) {
            $path = $folderValues[0].TrimEnd('/')
            if ($path -ne $library -and -not $path.StartsWith($library + '/', [StringComparison]::OrdinalIgnoreCase)) {
                throw "Folder selector leaves its library: $Container"
            }
            if ([IO.Path]::GetExtension($path)) { throw "Container selects a file, not a folder: $Container" }
        }
    } elseif ($path -match '(?i)\.aspx$' -or $parameters.Count) {
        throw "Unsupported container/view URL: $Container"
    }
    if ($path -match '(^|/)\.\.?(/|$)' -or $path.Contains('\') -or $path.Contains('//') -or
        $path -notmatch '^/sites/[^/]+/.+') {
        throw "Invalid destination folder: $path"
    }
    $webPath = [regex]::Match($path, '^/sites/[^/]+').Value
    return [pscustomobject]@{
        WebUrl = 'https://whallc.sharepoint.com' + $webPath
        Folder = $path
        FileRef = $path + '/' + $Name
    }
}

function New-FixedDocumentPlan {
    param([string]$Documents, [string]$Workbook, [string]$Runs)
    $priorities = @(Read-PriorityWorkbook $Workbook)
    $changed = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $latest = @{}
    foreach ($directory in Get-ChildItem -LiteralPath $Runs -Directory) {
        $manifestPath = Join-Path $directory.FullName 'manifest.json'
        $filesPath = Join-Path $directory.FullName 'files.csv'
        if (-not (Test-Path -LiteralPath $manifestPath) -or -not (Test-Path -LiteralPath $filesPath)) { continue }
        $manifest = Get-Content -LiteralPath $manifestPath -Raw | ConvertFrom-Json
        $stamp = [datetime]::Parse($manifest.Timestamp, [cultureinfo]::InvariantCulture)
        foreach ($row in Import-Csv -LiteralPath $filesPath) {
            if ($row.Status -eq 'updated' -and $manifest.Applied) { [void]$changed.Add($row.File) }
            if ($row.SHA256After -match '^[a-fA-F0-9]{64}$') {
                if (-not $latest.ContainsKey($row.File) -or $stamp -gt $latest[$row.File].Timestamp) {
                    $latest[$row.File] = @{ Timestamp = $stamp; Hash = $row.SHA256After }
                } elseif ($stamp -eq $latest[$row.File].Timestamp -and $row.SHA256After -ne $latest[$row.File].Hash) {
                    throw "Conflicting manifest hashes for $($row.File)."
                }
            }
        }
    }
    if ($changed.Count -eq 0) { throw 'No applied remediation records found; refusing to select all files.' }
    $targets = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $plan = [System.Collections.Generic.List[object]]::new()
    $base = [IO.Path]::GetFullPath($Documents).TrimEnd('\') + '\'
    foreach ($relative in $changed | Sort-Object) {
        $local = [IO.Path]::GetFullPath((Join-Path $Documents $relative))
        if (-not $local.StartsWith($base, [StringComparison]::OrdinalIgnoreCase) -or
            -not (Test-Path -LiteralPath $local -PathType Leaf)) {
            throw "Missing or unsafe local document: $relative"
        }
        $hash = (Get-FileHash -LiteralPath $local -Algorithm SHA256).Hash
        if (-not $latest.ContainsKey($relative) -or $hash -ne $latest[$relative].Hash) {
            throw "Document differs from its latest verified remediation hash: $relative"
        }
        $name = [IO.Path]::GetFileName($local)
        $matches = @($priorities | Where-Object Name -EQ $name)
        if ($matches.Count -ne 1) { throw "Expected exactly one priority row for $relative; found $($matches.Count)." }
        $destination = Resolve-UploadDestination $matches[0].Container $name
        if (-not $targets.Add($destination.FileRef)) { throw "Duplicate remote target: $($destination.FileRef)" }
        $plan.Add([pscustomobject]@{
            RelativeFile = $relative; LocalPath = $local; SHA256 = $hash
            WebUrl = $destination.WebUrl; Folder = $destination.Folder; FileRef = $destination.FileRef
        })
    }
    return $plan.ToArray()
}

function Import-FixedDocumentPlan {
    param([string]$Path, [string]$Documents)
    $rows = @(Import-Csv -LiteralPath $Path)
    if ($rows.Count -eq 0) { throw 'Upload plan is empty.' }
    $base = [IO.Path]::GetFullPath($Documents).TrimEnd('\') + '\'
    $targets = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($row in $rows) {
        $local = [IO.Path]::GetFullPath((Join-Path $Documents $row.RelativeFile))
        if (-not $local.StartsWith($base, [StringComparison]::OrdinalIgnoreCase) -or
            [IO.Path]::IsPathRooted($row.RelativeFile) -or
            -not (Test-Path -LiteralPath $local -PathType Leaf)) {
            throw "Missing or unsafe local document in portable plan: $($row.RelativeFile)"
        }
        if ($row.SHA256 -notmatch '^[a-fA-F0-9]{64}$' -or
            (Get-FileHash -LiteralPath $local -Algorithm SHA256).Hash -ne $row.SHA256) {
            throw "Local file differs from portable plan hash: $($row.RelativeFile)"
        }
        $folderUrl = 'https://whallc.sharepoint.com' + (
            ($row.Folder -split '/' | ForEach-Object { [uri]::EscapeDataString($_) }) -join '/'
        )
        $destination = Resolve-UploadDestination $folderUrl ([IO.Path]::GetFileName($local))
        if ($destination.WebUrl -cne $row.WebUrl -or $destination.Folder -cne $row.Folder -or
            $destination.FileRef -cne $row.FileRef -or -not $targets.Add($destination.FileRef)) {
            throw "Inconsistent or duplicate destination in portable plan: $($row.RelativeFile)"
        }
        [pscustomobject]@{
            RelativeFile = $row.RelativeFile; LocalPath = $local; SHA256 = $row.SHA256
            WebUrl = $destination.WebUrl; Folder = $destination.Folder; FileRef = $destination.FileRef
        }
    }
}

function ConvertTo-SharePointTimestamp {
    param([datetime]$Value)
    if ($Value.Kind -eq [DateTimeKind]::Unspecified) {
        $Value = [datetime]::SpecifyKind($Value, [DateTimeKind]::Utc)
    }
    return $Value.ToUniversalTime().ToString('o')
}

function Get-RemoteDocumentState {
    param($Plan, $Connection)
    $file = Get-PnPFile -Url $Plan.FileRef -AsFileObject -Connection $Connection -ErrorAction Stop
    if ($null -eq $file) { throw "Destination file does not exist: $($Plan.FileRef)" }
    Get-PnPProperty -ClientObject $file -Property UniqueId,ETag,CheckOutType,LockedByUser,ListItemAllFields -Connection $Connection | Out-Null
    $item = $file.ListItemAllFields
    Get-PnPProperty -ClientObject $item -Property ParentList -Connection $Connection | Out-Null
    $list = $item.ParentList
    Get-PnPProperty -ClientObject $list -Property Id,ForceCheckout,EnableModeration -Connection $Connection | Out-Null
    $system = Get-RemoteSystemMetadata ([string]$list.Id) ([int]$item.Id) $Connection
    return [pscustomobject]@{
        File = $file; Item = $item; List = $list
        UniqueId = [string]$file.UniqueId; ETag = [string]$file.ETag
        ItemId = [int]$item.Id; ListId = [string]$list.Id
        Author = $system.Author; Editor = $system.Editor
        Created = $system.Created; Modified = $system.Modified
    }
}

function Get-RemoteSystemMetadata {
    param([string]$ListId, [int]$ItemId, $Connection)
    $endpoint = "/_api/web/lists(guid'$ListId')/items($ItemId)?`$select=Id,Created,Modified,AuthorId,EditorId"
    $fields = Invoke-PnPSPRestMethod -Url $endpoint -Method Get -Connection $Connection -ErrorAction Stop
    if ([int]$fields.Id -ne $ItemId -or [int]$fields.AuthorId -le 0 -or [int]$fields.EditorId -le 0) {
        throw 'REST metadata returned missing or inconsistent item/user IDs.'
    }
    $dates = @{}
    foreach ($key in 'Created','Modified') {
        $value = $fields.$key
        if ($value -is [datetime]) {
            if ($value.Kind -eq [DateTimeKind]::Unspecified) {
                throw "REST $key has no timezone; refusing to infer UTC."
            }
            $dates[$key] = $value.ToUniversalTime().ToString('o')
        } else {
            if ([string]$value -notmatch '(Z|[+-]\d{2}:\d{2})$') {
                throw "REST $key has no explicit timezone; refusing to infer UTC."
            }
            $date = [datetimeoffset]::Parse([string]$value, [cultureinfo]::InvariantCulture)
            $dates[$key] = $date.UtcDateTime.ToString('o')
        }
    }
    return [pscustomobject]@{
        Author = [int]$fields.AuthorId; Editor = [int]$fields.EditorId
        Created = $dates.Created; Modified = $dates.Modified
    }
}

function Assert-RemoteDocumentWritable {
    param($State)
    if ([string]$State.File.CheckOutType -ne 'None') { throw 'Target is checked out; check it in before uploading.' }
    if ($null -ne $State.File.LockedByUser -and $State.File.LockedByUser.ServerObjectIsNull -ne $true) {
        throw 'Target is locked by a user; close the editing session first.'
    }
    if ($State.List.ForceCheckout) { throw 'Library requires check-out; use a separately approved check-in workflow.' }
    if ($State.List.EnableModeration) { throw 'Library uses content approval; use a separately approved publishing workflow.' }
    $fields = $State.Item.FieldValues
    if ($fields.ContainsKey('_ComplianceTag') -and [string]$fields['_ComplianceTag']) {
        throw 'Target has a retention label; have its compliance requirements reviewed before overwriting.'
    }
    if ($fields.ContainsKey('_ModerationStatus') -and [int]$fields['_ModerationStatus'] -eq 3) {
        throw 'Target is a draft; publishing state requires an explicit workflow.'
    }
    if ($fields.ContainsKey('_UIVersionString') -and [string]$fields['_UIVersionString'] -match '\.[1-9]\d*$') {
        throw 'Target is a minor/draft version; publishing state requires an explicit workflow.'
    }
    if ($State.Author -le 0 -or $State.Editor -le 0) { throw 'Invalid Author/Editor user IDs.' }
}

function Get-RemoteDocumentBytes {
    param($Plan, $Connection)
    $stream = Get-PnPFile -Url $Plan.FileRef -AsMemoryStream -Connection $Connection -ErrorAction Stop
    if ($null -eq $stream) { throw "Could not read remote file: $($Plan.FileRef)" }
    try { return ,$stream.ToArray() } finally { $stream.Dispose() }
}

function Get-BytesSHA256 {
    param([byte[]]$Bytes)
    return [Convert]::ToHexString([Security.Cryptography.SHA256]::HashData($Bytes))
}

function Read-OfficePartXml {
    param($Entry)
    $stream = $Entry.Open()
    try {
        $settings = [Xml.XmlReaderSettings]::new()
        $settings.DtdProcessing = [Xml.DtdProcessing]::Prohibit
        $settings.XmlResolver = $null
        $reader = [Xml.XmlReader]::Create($stream, $settings)
        try {
            $xml = [Xml.XmlDocument]::new()
            $xml.XmlResolver = $null
            $xml.PreserveWhitespace = $true
            $xml.Load($reader)
            return ,$xml
        } finally { $reader.Dispose() }
    } finally { $stream.Dispose() }
}

function Compare-OfficeDocumentPackage {
    param([string]$LocalPath, [byte[]]$RemoteBytes)
    $local = [IO.Compression.ZipFile]::OpenRead($LocalPath)
    $memory = [IO.MemoryStream]::new($RemoteBytes, $false)
    $remote = $null
    try {
        $remote = [IO.Compression.ZipArchive]::new($memory, [IO.Compression.ZipArchiveMode]::Read)
        $maps = @()
        foreach ($zip in @($local, $remote)) {
            $map = [Collections.Generic.Dictionary[string,object]]::new([StringComparer]::Ordinal)
            foreach ($entry in $zip.Entries) {
                if (-not $map.TryAdd($entry.FullName, $entry)) { throw "Duplicate Office package part: $($entry.FullName)" }
            }
            $maps += ,$map
        }
        if (@($maps[0].Keys | Where-Object { $_.StartsWith('_xmlsignatures/', [StringComparison]::OrdinalIgnoreCase) }).Count -or
            @($maps[1].Keys | Where-Object { $_.StartsWith('_xmlsignatures/', [StringComparison]::OrdinalIgnoreCase) }).Count) {
            throw 'Signed Office packages require exact-byte verification.'
        }
        $allowed = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
        foreach ($name in $maps[0].Keys) {
            if (-not $maps[1].ContainsKey($name)) { continue }
            $expectedRoot = switch -Regex ($name) {
                '^docProps/core\.xml$' { @('coreProperties','http://schemas.openxmlformats.org/package/2006/metadata/core-properties'); break }
                '^docProps/custom\.xml$' { @('Properties','http://schemas.openxmlformats.org/officeDocument/2006/custom-properties'); break }
                '^customXml/item\d+\.xml$' { @(); break }
                default { continue }
            }
            if ($name -notmatch '^docProps/(core|custom)\.xml$|^customXml/item\d+\.xml$') { continue }
            $x = Read-OfficePartXml $maps[0][$name]
            $y = Read-OfficePartXml $maps[1][$name]
            if ($name.StartsWith('customXml/')) {
                $roots = @{
                    properties = 'http://schemas.microsoft.com/office/2006/metadata/properties'
                    contentTypeSchema = 'http://schemas.microsoft.com/office/2006/metadata/contentType'
                }
                $rootName = $x.DocumentElement.LocalName
                if (-not $roots.ContainsKey($rootName)) { continue }
                $expectedRoot = @($rootName, $roots[$rootName])
            }
            if ($x.DocumentElement.LocalName -cne $expectedRoot[0] -or
                $y.DocumentElement.LocalName -cne $expectedRoot[0] -or
                $x.DocumentElement.NamespaceURI -cne $expectedRoot[1] -or
                $y.DocumentElement.NamespaceURI -cne $expectedRoot[1]) { continue }
            [void]$allowed.Add($name)
            if ($name -notmatch '^customXml/(item\d+)\.xml$') { continue }
            $relsName = 'customXml/_rels/' + $Matches[1] + '.xml.rels'
            if (-not $maps[0].ContainsKey($relsName) -or -not $maps[1].ContainsKey($relsName)) { continue }
            $rels = Read-OfficePartXml $maps[0][$relsName]
            foreach ($rel in $rels.DocumentElement.ChildNodes) {
                if ($rel -isnot [Xml.XmlElement] -or
                    $rel.GetAttribute('Type') -cne 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps' -or
                    $rel.GetAttribute('TargetMode') -eq 'External' -or
                    $rel.GetAttribute('Target') -notmatch '^itemProps\d+\.xml$') { continue }
                $propsName = 'customXml/' + $rel.GetAttribute('Target')
                if (-not $maps[0].ContainsKey($propsName) -or -not $maps[1].ContainsKey($propsName)) { continue }
                $px = Read-OfficePartXml $maps[0][$propsName]
                $py = Read-OfficePartXml $maps[1][$propsName]
                $ns = 'http://schemas.openxmlformats.org/officeDocument/2006/customXml'
                if ($px.DocumentElement.LocalName -cne 'datastoreItem' -or
                    $py.DocumentElement.LocalName -cne 'datastoreItem' -or
                    $px.DocumentElement.NamespaceURI -cne $ns -or
                    $py.DocumentElement.NamespaceURI -cne $ns) { continue }
                # SharePoint can drop schemaRefs when rebuilding a library metadata datastore.
                $px.DocumentElement.RemoveAttribute('itemID', $ns)
                $py.DocumentElement.RemoveAttribute('itemID', $ns)
                $oldRefs = @($px.DocumentElement.ChildNodes | Where-Object {
                    $_ -is [Xml.XmlElement] -and $_.LocalName -ceq 'schemaRefs' -and $_.NamespaceURI -ceq $ns
                })
                $newRefs = @($py.DocumentElement.ChildNodes | Where-Object {
                    $_ -is [Xml.XmlElement] -and $_.LocalName -ceq 'schemaRefs' -and $_.NamespaceURI -ceq $ns
                })
                if ($oldRefs.Count -eq 1 -and $newRefs.Count -eq 0) {
                    $validRefs = $oldRefs[0].Attributes.Count -eq 0
                    foreach ($ref in $oldRefs[0].ChildNodes) {
                        if ($ref -is [Xml.XmlWhitespace]) { continue }
                        if ($ref -isnot [Xml.XmlElement] -or $ref.LocalName -cne 'schemaRef' -or
                            $ref.NamespaceURI -cne $ns -or $ref.Attributes.Count -ne 1 -or
                            -not $ref.HasAttribute('uri', $ns) -or $ref.HasChildNodes) {
                            $validRefs = $false
                        }
                    }
                    if ($validRefs) { [void]$px.DocumentElement.RemoveChild($oldRefs[0]) }
                }
                foreach ($document in @($px, $py)) {
                    if (-not $document.DocumentElement.HasChildNodes) { $document.DocumentElement.IsEmpty = $true }
                }
                if ($px.OuterXml -ceq $py.OuterXml) { [void]$allowed.Add($propsName) }
            }
        }
        $differences = [Collections.Generic.List[object]]::new()
        $names = [Collections.Generic.HashSet[string]]::new($maps[0].Keys, [StringComparer]::Ordinal)
        $names.UnionWith($maps[1].Keys)
        foreach ($name in $names | Sort-Object) {
            $a = if ($maps[0].ContainsKey($name)) { $maps[0][$name] } else { $null }
            $b = if ($maps[1].ContainsKey($name)) { $maps[1][$name] } else { $null }
            if ($null -eq $a -or $null -eq $b) {
                if ($null -eq $a -and $name -cmatch '^\[trash\]/\d+\.dat$') {
                    $differences.Add([pscustomobject]@{ Part = $name; Change = 'AddedSharePointTrash' })
                    continue
                }
                throw "Office package part added or removed outside the metadata allowance: $name"
            }
            $as = $a.Open(); $bs = $b.Open()
            try {
                $ah = [Convert]::ToHexString([Security.Cryptography.SHA256]::HashData($as))
                $bh = [Convert]::ToHexString([Security.Cryptography.SHA256]::HashData($bs))
            } finally { $as.Dispose(); $bs.Dispose() }
            if ($ah -ceq $bh) { continue }
            if (-not $allowed.Contains($name)) {
                throw "Office content or unapproved package part differs: $name"
            }
            $differences.Add([pscustomobject]@{ Part = $name; Change = 'RecognizedOfficeMetadata' })
        }
        return ,$differences.ToArray()
    } finally {
        if ($null -ne $remote) { $remote.Dispose() }
        $memory.Dispose()
        $local.Dispose()
    }
}

function Assert-PlannedDocumentContent {
    param($Plan, [byte[]]$Bytes)
    $remoteHash = Get-BytesSHA256 $Bytes
    if ($remoteHash -eq $Plan.SHA256) { return 'ExactSHA256' }
    if (-not $AllowOfficeMetadataChanges -or [IO.Path]::GetExtension($Plan.LocalPath) -notin '.docx','.xlsx','.pptx') {
        throw 'Remote bytes differ from the planned upload; refusing to restore potentially stale metadata.'
    }
    if ((Get-FileHash -LiteralPath $Plan.LocalPath -Algorithm SHA256).Hash -ne $Plan.SHA256) {
        throw 'Local document changed; refusing Office package comparison.'
    }
    $differences = Compare-OfficeDocumentPackage $Plan.LocalPath $Bytes
    if (-not $Plan.PSObject.Properties['VerificationAuditPath'] -or -not $Plan.VerificationAuditPath) {
        throw 'Office metadata acceptance requires a persistent verification audit path.'
    }
    [ordered]@{
        Timestamp = [datetime]::UtcNow.ToString('o')
        Target = $Plan.FileRef; LocalSHA256 = $Plan.SHA256; RemoteSHA256 = $remoteHash
        Differences = @($differences)
    } | ConvertTo-Json -Depth 5 -Compress |
        Add-Content -LiteralPath $Plan.VerificationAuditPath -Encoding utf8
    Write-Warning "Accepted Office metadata-only differences for $($Plan.FileRef); package content and relationships match. See $($Plan.VerificationAuditPath)"
    return 'OfficePackageContent'
}

function Read-StableDocumentBackup {
    param($Plan, $Connection)
    for ($attempt = 1; $attempt -le 3; $attempt++) {
        $before = Get-RemoteDocumentState $Plan $Connection
        Assert-RemoteDocumentWritable $before
        $bytes = Get-RemoteDocumentBytes $Plan $Connection
        $after = Get-RemoteDocumentState $Plan $Connection
        Assert-RemoteDocumentWritable $after
        foreach ($key in 'UniqueId','ItemId','ListId') {
            if ($before.$key -cne $after.$key) {
                throw "Destination identity changed while backing up $($Plan.FileRef): $key [$($before.$key)] -> [$($after.$key)]."
            }
        }
        $differences = @(
            foreach ($key in 'ETag','Author','Editor','Created','Modified') {
                if ($before.$key -cne $after.$key) { "$key [$($before.$key)] -> [$($after.$key)]" }
            }
        )
        if ($differences.Count -eq 0) {
            return [pscustomobject]@{ State = $after; Bytes = $bytes }
        }
        $detail = "Destination changed while backing up $($Plan.FileRef) (attempt $attempt/3): $($differences -join '; ')"
        if ($attempt -eq 3) { throw "$detail. No stable backup obtained; nothing uploaded." }
        Write-Warning "$detail. Discarding this read and retrying the complete snapshot in 2 seconds."
        Start-Sleep -Seconds 2
    }
}

function Restore-DocumentMetadata {
    param($Plan, $Before, $Connection, [string]$ExpectedETag)
    $current = Get-RemoteDocumentState $Plan $Connection
    if ($current.UniqueId -ne $Before.UniqueId -or $current.ItemId -ne $Before.ItemId -or $current.ListId -ne $Before.ListId) {
        throw 'Remote file identity changed; refusing to restore metadata onto another item.'
    }
    if (-not $ExpectedETag -or $current.ETag -cne $ExpectedETag) {
        throw 'Remote ETag changed after upload; refusing to restore potentially stale metadata.'
    }
    Assert-RemoteDocumentWritable $current
    Assert-PlannedDocumentContent $Plan (Get-RemoteDocumentBytes $Plan $Connection) | Out-Null
    $changedFields = @('Author','Editor','Created','Modified' | Where-Object { $current.$_ -cne $Before.$_ })
    if ($changedFields.Count -eq 0) { return }
    Set-DocumentMetadataValues $current @{
        Author = [string]$Before.Author; Editor = [string]$Before.Editor
        Created = [datetime]::Parse($Before.Created, [cultureinfo]::InvariantCulture).ToUniversalTime()
        Modified = [datetime]::Parse($Before.Modified, [cultureinfo]::InvariantCulture).ToUniversalTime()
    }
    $after = Get-RemoteDocumentState $Plan $Connection
    foreach ($key in 'Author','Editor','Created','Modified','UniqueId','ItemId','ListId') {
        if ($after.$key -cne $Before.$key) {
            throw "Post-upload verification failed for $key. Expected [$($Before.$key)], actual [$($after.$key)]."
        }
    }
}

function Set-DocumentMetadataValues {
    param($Current, [hashtable]$Values)
    # Send UTC values directly; verify via REST rather than timezone-less CSOM field values.
    foreach ($key in 'Author','Editor') {
        $user = [Microsoft.SharePoint.Client.FieldUserValue]::new()
        $user.LookupId = [int]$Values[$key]
        $Current.Item[$key] = $user
    }
    foreach ($key in 'Created','Modified') {
        $date = [datetime]$Values[$key]
        if ($date.Kind -ne [DateTimeKind]::Utc) { throw "$key must be an explicitly UTC DateTime." }
        $Current.Item[$key] = $date
    }
    $Current.Item.UpdateOverwriteVersion()
    $Current.Item.Context.ExecuteQuery()
}

function Restore-SnapshotDocumentMetadata {
    param([string]$Path, [object[]]$Plan, [string]$ApplicationId)
    $json = [System.Text.Json.JsonDocument]::Parse([IO.File]::ReadAllText($Path))
    try {
        $root = $json.RootElement
        $before = [pscustomobject]@{
            UniqueId = $root.GetProperty('UniqueId').GetString()
            ItemId = $root.GetProperty('ItemId').GetInt32()
            ListId = $root.GetProperty('ListId').GetString()
            Author = $root.GetProperty('Author').GetInt32()
            Editor = $root.GetProperty('Editor').GetInt32()
            Created = $root.GetProperty('Created').GetString()
            Modified = $root.GetProperty('Modified').GetString()
        }
        $target = $root.GetProperty('Target').GetString()
        $web = $root.GetProperty('WebUrl').GetString()
        $backupHash = $root.GetProperty('SHA256').GetString()
    } finally { $json.Dispose() }
    $entries = @($Plan | Where-Object { $_.FileRef -ceq $target -and $_.WebUrl -ceq $web })
    if ($entries.Count -ne 1) { throw 'Snapshot must match exactly one planned destination.' }
    if (-not $Path.EndsWith('.metadata.json', [StringComparison]::OrdinalIgnoreCase)) {
        throw 'Expected a .metadata.json snapshot adjacent to its original backup.'
    }
    $backup = $Path.Substring(0, $Path.Length - '.metadata.json'.Length)
    if ((Get-FileHash -LiteralPath $backup -Algorithm SHA256).Hash -ne $backupHash) {
        throw 'Original backup hash does not match its metadata snapshot.'
    }
    foreach ($key in 'Created','Modified') {
        $date = [datetime]::Parse($before.$key, [cultureinfo]::InvariantCulture, [Globalization.DateTimeStyles]::RoundtripKind)
        if ($date.Kind -ne [DateTimeKind]::Utc) { throw "Snapshot $key must specify UTC." }
    }
    $entry = $entries[0]
    $connection = Connect-DocumentUploadSite -Url $web -ApplicationId $ApplicationId
    $current = Get-RemoteDocumentState $entry $connection
    Restore-DocumentMetadata $entry $before $connection $current.ETag
    Assert-PlannedDocumentContent $entry (Get-RemoteDocumentBytes $entry $connection) | Out-Null
    return [pscustomobject]@{
        Target = $target; Status = 'MetadataRestoredAndVerified'; Snapshot = $Path
        Author = $before.Author; Editor = $before.Editor
        Created = $before.Created; Modified = $before.Modified
    }
}

function Invoke-FixedDocumentUpload {
    param($Plan, $Before, [string]$RemoteHash, $Connection)
    if ((Get-FileHash -LiteralPath $Plan.LocalPath -Algorithm SHA256).Hash -ne $Plan.SHA256) {
        throw 'Local file changed after planning.'
    }
    $current = Get-RemoteDocumentState $Plan $Connection
    Assert-RemoteDocumentWritable $current
    foreach ($key in 'UniqueId','ItemId','ListId','ETag','Author','Editor','Created','Modified') {
        if ($current.$key -cne $Before.$key) { throw "Destination changed after preflight ($key); nothing uploaded." }
    }
    $hash = Get-BytesSHA256 (Get-RemoteDocumentBytes $Plan $Connection)
    if ($hash -ne $RemoteHash) { throw 'Remote content changed after preflight; nothing uploaded.' }
    if ($hash -eq $Plan.SHA256) { return 'AlreadyIdentical' }
    try {
        Send-ConditionalDocumentContent $Plan $current
    } catch {
        $uploadError = $_.Exception.Message
        # A rejected ETag and an accepted upload with a lost response can look alike.
        throw "Upload failed: $uploadError. Remote outcome is unverified. No metadata recovery attempted, to avoid overwriting a concurrent edit. Inspect remote bytes/metadata and the saved backups before retrying."
    }
    $written = Get-RemoteDocumentState $Plan $Connection
    Restore-DocumentMetadata $Plan $Before $Connection $written.ETag
    $verification = Assert-PlannedDocumentContent $Plan (Get-RemoteDocumentBytes $Plan $Connection)
    if ($verification -eq 'OfficePackageContent') { return 'UploadedAndVerifiedOfficeContent' }
    return 'UploadedAndVerified'
}

function Send-ConditionalDocumentContent {
    param($Plan, $Current)
    $stream = [IO.File]::OpenRead($Plan.LocalPath)
    try {
        $information = [Microsoft.SharePoint.Client.FileSaveBinaryInformation]::new()
        $information.ContentStream = $stream
        $information.ETag = $Current.ETag
        $information.CheckRequiredFields = $true
        $Current.File.SaveBinary($information)
        # Do not retry a content upload automatically after an ambiguous network failure.
        $Current.File.Context.ExecuteQuery()
    } finally { $stream.Dispose() }
}

function Connect-DocumentUploadSite {
    param([string]$Url, [string]$ApplicationId)
    $connectParams = @{
        Url = $Url
        Interactive = $true
        PersistLogin = $true
        ReturnConnection = $true
        ErrorAction = 'Stop'
    }
    if ($ApplicationId) {
        $parsedId = [guid]::Empty
        if (-not [guid]::TryParse($ApplicationId, [ref]$parsedId) -or $parsedId -eq [guid]::Empty) {
            throw 'ClientId must be an application client ID GUID, not a placeholder or tenant name. Omit it to use the existing PnP configuration.'
        }
        $connectParams.ClientId = $ApplicationId
    }
    return Connect-PnPOnline @connectParams
}

function Invoke-PublishFixedDocuments {
    $root = [IO.Path]::GetFullPath($RepositoryRoot)
    $documents = if ($DocumentsFolder) { [IO.Path]::GetFullPath($DocumentsFolder) } else { Join-Path $root 'UpdatedEmbeddedLinkCopies' }
    $workbook = if ($PriorityWorkbook) { [IO.Path]::GetFullPath($PriorityWorkbook) } else { Join-Path $root 'EmbeddedLinksPriorityList-Copy.xlsx' }
    $runs = if ($RemediationRunsFolder) { [IO.Path]::GetFullPath($RemediationRunsFolder) } else { Join-Path $root 'LinkRemediationRuns' }
    $output = if ($OutputFolder) { [IO.Path]::GetFullPath($OutputFolder) } else {
        Join-Path $root ('UploadRuns\' + (Get-Date -Format 'yyyyMMdd-HHmmss'))
    }
    if ($Apply -and $Preflight) { throw 'Choose -Apply or -Preflight, not both.' }
    if ($RestoreMetadataSnapshot -and ($Apply -or $Preflight)) {
        throw 'Metadata recovery is a separate write operation; do not combine it with -Apply or -Preflight.'
    }
    if (Test-Path -LiteralPath $output) { throw 'OutputFolder already exists; use a new folder for every run.' }
    $plan = if ($PlanCsv) { @(Import-FixedDocumentPlan ([IO.Path]::GetFullPath($PlanCsv)) $documents) } else {
        @(New-FixedDocumentPlan $documents $workbook $runs)
    }
    New-Item -ItemType Directory -Path $output -Force | Out-Null
    if ($AllowOfficeMetadataChanges) {
        foreach ($entry in $plan) {
            $entry | Add-Member -NotePropertyName VerificationAuditPath -NotePropertyValue (Join-Path $output 'office-verification.jsonl')
        }
    }
    $plan | Select-Object RelativeFile,SHA256,WebUrl,Folder,FileRef |
        Export-Csv -LiteralPath (Join-Path $output 'upload-plan.csv') -NoTypeInformation -Encoding utf8BOM
    Write-Host "Planned $($plan.Count) corrected documents. Review: $(Join-Path $output 'upload-plan.csv')"
    if ($RestoreMetadataSnapshot) {
        Import-Module PnP.PowerShell -MinimumVersion 3.3.0 -ErrorAction Stop
        $result = Restore-SnapshotDocumentMetadata ([IO.Path]::GetFullPath($RestoreMetadataSnapshot)) $plan $ClientId
        $result | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $output 'metadata-recovery.json') -Encoding utf8
        Write-Host "MetadataRestoredAndVerified: $($result.Target). No document content uploaded."
        return
    }
    if (-not $Apply -and -not $Preflight) {
        Write-Host 'Offline plan only. No SharePoint connection or remote changes.'
        return
    }
    Import-Module PnP.PowerShell -MinimumVersion 3.3.0 -ErrorAction Stop
    $connections = @{}
    $baselines = @{}
    $results = [System.Collections.Generic.List[object]]::new()
    foreach ($entry in $plan) {
        $results.Add([pscustomobject]@{ File = $entry.RelativeFile; Target = $entry.FileRef; Status = 'Pending'; Error = '' })
    }
    $resultPath = Join-Path $output 'upload-results.csv'
    try {
        for ($i = 0; $i -lt $plan.Count; $i++) {
            $entry = $plan[$i]
            if (-not $connections.ContainsKey($entry.WebUrl)) {
                $connections[$entry.WebUrl] = Connect-DocumentUploadSite -Url $entry.WebUrl -ApplicationId $ClientId
            }
            $connection = $connections[$entry.WebUrl]
            $snapshot = Read-StableDocumentBackup $entry $connection
            $state = $snapshot.State
            $bytes = $snapshot.Bytes
            $backup = Join-Path (Join-Path $output 'remote-backups') $entry.RelativeFile
            New-Item -ItemType Directory -Path (Split-Path -Parent $backup) -Force | Out-Null
            [IO.File]::WriteAllBytes($backup, $bytes)
            $hash = Get-BytesSHA256 $bytes
            if ((Get-FileHash -LiteralPath $backup -Algorithm SHA256).Hash -ne $hash) { throw 'Remote backup verification failed.' }
            $metadata = [ordered]@{ Target = $entry.FileRef; WebUrl = $entry.WebUrl; SHA256 = $hash; Backup = $backup }
            foreach ($key in 'UniqueId','ItemId','ListId','ETag','Author','Editor','Created','Modified') { $metadata[$key] = $state.$key }
            $metadata | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath ($backup + '.metadata.json') -Encoding utf8
            $baselines[$entry.RelativeFile] = @{ State = $state; Hash = $hash }
            $results[$i].Status = 'PreflightPassed'
            $results | Export-Csv -LiteralPath $resultPath -NoTypeInformation -Encoding utf8BOM
        }
        if (-not $Apply) { Write-Host 'Preflight passed. Remote originals backed up; nothing uploaded.'; return }
        for ($i = 0; $i -lt $plan.Count; $i++) {
            $entry = $plan[$i]
            $results[$i].Status = 'UploadStarted'
            $results | Export-Csv -LiteralPath $resultPath -NoTypeInformation -Encoding utf8BOM
            $baseline = $baselines[$entry.RelativeFile]
            $results[$i].Status = Invoke-FixedDocumentUpload $entry $baseline.State $baseline.Hash $connections[$entry.WebUrl]
            $results | Export-Csv -LiteralPath $resultPath -NoTypeInformation -Encoding utf8BOM
            Write-Host "$($results[$i].Status): $($entry.RelativeFile)"
        }
    } catch {
        $results[$i].Error = $_.Exception.Message
        $results[$i].Status = if ($results[$i].Status -eq 'UploadStarted') { 'UploadFailedNeedsReview' } else { 'PreflightOrSafetyCheckFailed' }
        $results | Export-Csv -LiteralPath $resultPath -NoTypeInformation -Encoding utf8BOM
        throw "Stopped; no further uploads attempted. $($_.Exception.Message) Results and originals: $output"
    }
}

if ($MyInvocation.InvocationName -ne '.') { Invoke-PublishFixedDocuments }
