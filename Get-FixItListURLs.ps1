<#
.SYNOPSIS
    Tenant-wide link-risk audit for site migration planning.

.DESCRIPTION
    This script:
      1) Verifies required modules and installs missing modules.
      2) Verifies connection/access to the SharePoint admin endpoint every run.
      3) Uses cached PnP sign-in (PersistLogin) after first successful auth.
      4) Enumerates SharePoint site collections (no OneDrive personal sites).
      5) Lets the operator pick one site collection/Teams-connected site scope.
      6) Prompts for the old URL base and new URL base.
      7) Scans documents tenant-wide for embedded URLs that will break.
      8) Optionally scans additional list/page content into a separate CSV.

    Output is written to timestamped CSV files:
      - Documents folder: one row per breaking embedded URL in documents.
      - Other folder: one row per breaking embedded URL in non-document content.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$RootUrl = "https://whallc.sharepoint.com",

    [Parameter(Mandatory = $false)]
    [string]$AdminUrl = "https://whallc-admin.sharepoint.com",

    [Parameter(Mandatory = $false)]
    [string]$ClientId,

    [Parameter(Mandatory = $false)]
    [switch]$EnableDocumentContentScan,

    [Parameter(Mandatory = $false)]
    [string]$DocumentsOutputFolder = "P:\IT\Private\zzBPosey\whaitapp09\SharePoint-Revamp\FixItList-Audits\Documents",

    [Parameter(Mandatory = $false)]
    [string]$OtherOutputFolder = "P:\IT\Private\zzBPosey\whaitapp09\SharePoint-Revamp\FixItList-Audits\Other"
)

$ErrorActionPreference = "Stop"

$SupportedExtensions = @('.docx', '.xlsx', '.pptx', '.pdf')
$LegacyExtensions = @('.doc', '.xls', '.ppt')
$SkipListTitles = @('Form Templates', 'Preservation Hold Library')

$DocLinkPrefixByExt = @{
    '.docx' = 'w'; '.doc' = 'w'
    '.xlsx' = 'x'; '.xls' = 'x'
    '.pptx' = 'p'; '.ppt' = 'p'
    '.pdf'  = 'b'
}

$PdfNoiseHosts = @(
    'ns.adobe.com', 'www.w3.org', 'w3.org', 'purl.org',
    'www.iec.ch', 'iec.ch', 'www.color.org', 'www.npes.org', 'www.aiim.org'
)

function Install-RequiredModule {
    param([Parameter(Mandatory = $true)][string]$Name)

    if (-not (Get-Module -ListAvailable -Name $Name)) {
        Write-Host "Module $Name was not found. Installing for current user..."
        Install-Module -Name $Name -Scope CurrentUser -Force -AllowClobber
    }

    Import-Module $Name -ErrorAction Stop
}

function Connect-ValidatedPnP {
    param(
        [Parameter(Mandatory = $true)][string]$Url,
        [Parameter(Mandatory = $false)][string]$ClientId,
        [Parameter(Mandatory = $false)][switch]$AdminConnection
    )

    for ($attempt = 1; $attempt -le 3; $attempt++) {
        try {
            Write-Host "Connecting to $Url (attempt $attempt of 3)..."

            $connectParams = @{
                Url         = $Url
                Interactive = $true
                PersistLogin = $true
                ReturnConnection = $true
            }
            if ($ClientId) {
                $connectParams.ClientId = $ClientId
            }

            $conn = Connect-PnPOnline @connectParams

            if ($AdminConnection) {
                Get-PnPTenantSite -Connection $conn -Detailed | Select-Object -First 1 | Out-Null
            }
            else {
                Get-PnPWeb -Connection $conn -Includes Url | Out-Null
            }

            return $conn
        }
        catch {
            Write-Warning ("Connection validation failed for {0}: {1}" -f $Url, $_.Exception.Message)
            if ($attempt -eq 3) {
                throw "Unable to establish a validated connection to $Url after 3 attempts."
            }
        }
    }
}

function New-DirectoryIfMissing {
    param([Parameter(Mandatory = $true)][string]$Path)

    if (-not (Test-Path -Path $Path)) {
        New-Item -ItemType Directory -Path $Path -Force | Out-Null
    }
}

function Get-TimestampedReportPath {
    param(
        [Parameter(Mandatory = $true)][string]$Folder,
        [Parameter(Mandatory = $true)][string]$Prefix
    )

    $stamp = Get-Date -Format "yyyyMMdd-HHmmss"
    return Join-Path -Path $Folder -ChildPath "$Prefix-$stamp.csv"
}

function Get-EncodedServerRelativePath {
    param([Parameter(Mandatory = $true)][string]$Path)

    return (($Path -split '/') | ForEach-Object { [Uri]::EscapeDataString($_) }) -join '/'
}

function Get-DocumentBrowserUrl {
    param(
        [Parameter(Mandatory = $true)][string]$TenantBase,
        [Parameter(Mandatory = $true)][string]$FileRef,
        [Parameter(Mandatory = $true)][string]$Extension,
        [Parameter(Mandatory = $false)]$UniqueId
    )

    $encPath = Get-EncodedServerRelativePath -Path $FileRef
    $prefix = $DocLinkPrefixByExt[$Extension]
    if (-not $prefix -or -not $UniqueId) {
        return $TenantBase + $encPath
    }

    $guid = ([Guid]$UniqueId).ToString('N')
    return "$TenantBase/:${prefix}:/r$encPath" + "?d=$prefix$guid&csf=1&web=1"
}

function Get-FolderBrowserUrl {
    param(
        [Parameter(Mandatory = $true)][string]$TenantBase,
        [Parameter(Mandatory = $true)][string]$LibraryRootPath,
        [Parameter(Mandatory = $true)][Guid]$ListId,
        [Parameter(Mandatory = $true)][string]$FileDirRef
    )

    $cid = $ListId.ToString().ToLower()
    return ("{0}{1}?CID={2}&RootFolder={3}" -f
        $TenantBase,
        (Get-EncodedServerRelativePath -Path "$LibraryRootPath/Forms/AllItems.aspx"),
        $cid,
        [Uri]::EscapeDataString($FileDirRef))
}

function Get-CleanedValue {
    param([Parameter(Mandatory = $true)][string]$Value)

    $decoded = [System.Net.WebUtility]::HtmlDecode($Value)
    return $decoded.Trim().TrimEnd("'", '"', ':', ',', ')', ']', '}', '.', ';')
}

function Get-OfficeDocReferences {
    param([Parameter(Mandatory = $true)][System.IO.Stream]$Stream)

    $found = New-Object System.Collections.Generic.List[Object]
    $archive = [System.IO.Compression.ZipArchive]::new($Stream, [System.IO.Compression.ZipArchiveMode]::Read)
    try {
        foreach ($entry in $archive.Entries) {
            if ($entry.FullName -like "*_rels/*.rels") {
                $s = $entry.Open()
                $r = New-Object System.IO.StreamReader($s)
                $xmlText = $r.ReadToEnd()
                $r.Close()
                $s.Close()

                try {
                    [xml]$xml = $xmlText
                    foreach ($rel in $xml.Relationships.Relationship) {
                        if ($rel.TargetMode -eq "External" -and $rel.Target) {
                            $found.Add([PSCustomObject]@{ Value = $rel.Target; Confidence = "High"; Kind = "hyperlink" })
                        }
                    }
                }
                catch {
                }
            }
            elseif ($entry.FullName -match '\.xml$' -and
                ($entry.FullName -match 'document\.xml$|sheet\d*\.xml$|slide\d*\.xml$|header\d*\.xml$|footer\d*\.xml$|notesSlide\d*\.xml$')) {

                $s = $entry.Open()
                $r = New-Object System.IO.StreamReader($s)
                $text = $r.ReadToEnd()
                $r.Close()
                $s.Close()

                $plain = [regex]::Replace($text, '<[^>]+>', ' ')
                foreach ($m in [regex]::Matches($plain, 'https?://[^\s"<>]+')) {
                    $found.Add([PSCustomObject]@{ Value = $m.Value; Confidence = "Medium"; Kind = "plaintext" })
                }
            }
        }
    }
    finally {
        $archive.Dispose()
    }

    return $found | Sort-Object Value, Kind -Unique
}

function Get-PdfReferences {
    param([Parameter(Mandatory = $true)][byte[]]$Bytes)

    $found = New-Object System.Collections.Generic.List[Object]
    $text = [System.Text.Encoding]::Latin1.GetString($Bytes)

    foreach ($m in [regex]::Matches($text, 'https?://[^\s\)"<>\\]+')) {
        $hostMatch = [regex]::Match($m.Value, '^https?://([^/:?#]+)')
        if ($hostMatch.Success -and ($PdfNoiseHosts -contains $hostMatch.Groups[1].Value.ToLower())) {
            continue
        }
        $found.Add([PSCustomObject]@{ Value = $m.Value; Confidence = "Low"; Kind = "plaintext-pdf" })
    }

    return $found | Sort-Object Value, Kind -Unique
}

function Test-IsBreakingUrl {
    param(
        [Parameter(Mandatory = $true)][string]$Value,
        [Parameter(Mandatory = $true)][string]$OldBaseUrl,
        [Parameter(Mandatory = $true)][string]$TenantBase,
        [Parameter(Mandatory = $true)][string]$OldBasePath
    )

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return $false
    }

    $clean = Get-CleanedValue -Value $Value

    if ($clean.StartsWith($OldBaseUrl, [System.StringComparison]::OrdinalIgnoreCase)) {
        return $true
    }

    $viewerPattern = '^(?i)' + [regex]::Escape($TenantBase) + '/:[a-z]:/[rtu]' + [regex]::Escape($OldBasePath) + '(?=$|[/?#])'
    if ([regex]::IsMatch($clean, $viewerPattern)) {
        return $true
    }

    return $false
}

function Get-RoleAssignmentsWithOwnerLevel {
    param(
        [Parameter(Mandatory = $true)]$Securable,
        [Parameter(Mandatory = $true)][string]$CacheKey,
        [Parameter(Mandatory = $true)][hashtable]$Cache
    )

    if ($Cache.ContainsKey($CacheKey)) {
        return $Cache[$CacheKey]
    }

    $names = New-Object System.Collections.Generic.List[string]
    $assignments = Get-PnPProperty -ClientObject $Securable -Property RoleAssignments

    foreach ($ra in $assignments) {
        $member = Get-PnPProperty -ClientObject $ra -Property Member
        $bindings = Get-PnPProperty -ClientObject $ra -Property RoleDefinitionBindings
        $roles = @($bindings | Select-Object -ExpandProperty Name)

        if ($roles.Count -eq 0) {
            continue
        }

        $ownerRole = $roles | Where-Object {
            $_ -match '^(?i)full control$' -or $_ -match 'owner'
        }

        if ($ownerRole.Count -gt 0) {
            $names.Add($member.Title)
        }
    }

    $result = ($names | Sort-Object -Unique) -join "`n"
    $Cache[$CacheKey] = $result
    return $result
}

function Get-EffectiveOwnerPermissions {
    param(
        [Parameter(Mandatory = $true)]$Item,
        [Parameter(Mandatory = $true)]$List,
        [Parameter(Mandatory = $true)]$Web,
        [Parameter(Mandatory = $true)][hashtable]$Cache
    )

    $itemUnique = $false
    $listUnique = $false

    try {
        $itemUnique = [bool](Get-PnPProperty -ClientObject $Item -Property HasUniqueRoleAssignments)
    }
    catch {
    }

    try {
        $listUnique = [bool](Get-PnPProperty -ClientObject $List -Property HasUniqueRoleAssignments)
    }
    catch {
    }

    if ($itemUnique) {
        return Get-RoleAssignmentsWithOwnerLevel -Securable $Item -CacheKey ("Item:" + $Item.Id) -Cache $Cache
    }

    if ($listUnique) {
        return Get-RoleAssignmentsWithOwnerLevel -Securable $List -CacheKey ("List:" + $List.Id) -Cache $Cache
    }

    return Get-RoleAssignmentsWithOwnerLevel -Securable $Web -CacheKey ("Web:" + $Web.Id) -Cache $Cache
}

function Get-FolderDisplayName {
    param(
        [Parameter(Mandatory = $true)]$List,
        [Parameter(Mandatory = $true)][string]$FolderPath,
        [Parameter(Mandatory = $true)]$Connection,
        [Parameter(Mandatory = $true)][hashtable]$Cache
    )

    $cacheKey = "$($List.Id)|$FolderPath"
    if ($Cache.ContainsKey($cacheKey)) {
        return $Cache[$cacheKey]
    }

    $fallback = ($FolderPath.Trim('/') -split '/')[-1]

    $query = @"
<View Scope='RecursiveAll'>
  <Query>
    <Where>
      <And>
        <Eq><FieldRef Name='FSObjType'/><Value Type='Integer'>1</Value></Eq>
        <Eq><FieldRef Name='FileRef'/><Value Type='Text'>$FolderPath</Value></Eq>
      </And>
    </Where>
  </Query>
  <RowLimit>1</RowLimit>
</View>
"@

    try {
        $folderItem = Get-PnPListItem -List $List.Title -Connection $Connection -Query $query -Fields "Title", "FileLeafRef"
        if ($folderItem -and $folderItem.Count -gt 0) {
            $title = [string]$folderItem[0].FieldValues["Title"]
            $leaf = [string]$folderItem[0].FieldValues["FileLeafRef"]
            if (-not [string]::IsNullOrWhiteSpace($title)) {
                $fallback = $title
            }
            elseif (-not [string]::IsNullOrWhiteSpace($leaf)) {
                $fallback = $leaf
            }
        }
    }
    catch {
    }

    $Cache[$cacheKey] = $fallback
    return $fallback
}

function Select-SiteForScope {
    param([Parameter(Mandatory = $true)][array]$Sites)

    $teamsSites = @($Sites | Where-Object {
            ($_.GroupId -and $_.GroupId -ne [Guid]::Empty) -or $_.Template -like 'GROUP*'
        } | Sort-Object Url)

    $sharePointSites = @($Sites | Where-Object {
            -not ((($_.GroupId -and $_.GroupId -ne [Guid]::Empty) -or $_.Template -like 'GROUP*'))
        } | Sort-Object Url)

    Write-Host ""
    Write-Host "Choose which category to select from:"
    Write-Host "  1. SharePoint site collections"
    Write-Host "  2. Teams-connected sites"

    $firstChoice = $null
    while (-not $firstChoice) {
        $in = Read-Host "Enter 1 or 2"
        if ($in -in @('1', '2')) {
            $firstChoice = $in
        }
        else {
            Write-Host "  Invalid choice. Enter 1 or 2."
        }
    }

    $categorySites = @()
    if ($firstChoice -eq '1') {
        $categorySites = $sharePointSites
    }
    else {
        $categorySites = $teamsSites
    }

    if ($categorySites.Count -eq 0) {
        throw "No sites were found in the selected category."
    }

    Write-Host ""
    for ($i = 0; $i -lt $categorySites.Count; $i++) {
        $kind = if ((($categorySites[$i].GroupId -and $categorySites[$i].GroupId -ne [Guid]::Empty) -or $categorySites[$i].Template -like 'GROUP*')) {
            'Teams-connected'
        }
        else {
            'SharePoint'
        }
        Write-Host ("  {0,4}. [{1}] {2}" -f ($i + 1), $kind, $categorySites[$i].Url)
    }

    $selected = $null
    while (-not $selected) {
        $pick = Read-Host "`nPick exactly one site by number"
        if ($pick -match '^\d+$') {
            $idx = [int]$pick
            if ($idx -ge 1 -and $idx -le $categorySites.Count) {
                $selected = $categorySites[$idx - 1]
            }
        }
        if (-not $selected) {
            Write-Host "  Invalid selection."
        }
    }

    return $selected
}

function Get-BooleanPrompt {
    param(
        [Parameter(Mandatory = $true)][string]$Message,
        [Parameter(Mandatory = $false)][bool]$DefaultYes = $false
    )

    $suffix = if ($DefaultYes) { "[Y/n]" } else { "[y/N]" }
    while ($true) {
        $value = Read-Host "$Message $suffix"
        if ([string]::IsNullOrWhiteSpace($value)) {
            return $DefaultYes
        }
        switch ($value.Trim().ToLower()) {
            'y' { return $true }
            'yes' { return $true }
            'n' { return $false }
            'no' { return $false }
            default { Write-Host "  Please enter y or n." }
        }
    }
}

function Find-UrlsInText {
    param([Parameter(Mandatory = $true)][string]$Text)

    $hits = New-Object System.Collections.Generic.List[string]
    foreach ($m in [regex]::Matches($Text, 'https?://[^\s"''<>\)\]]+')) {
        $hits.Add($m.Value)
    }

    return $hits | Sort-Object -Unique
}

function Write-DocumentFinding {
    param(
        [Parameter(Mandatory = $true)][string]$OutputPath,
        [Parameter(Mandatory = $true)][pscustomobject]$Row
    )

    $Row | Export-Csv -Path $OutputPath -NoTypeInformation -Append
}

function Write-OtherFinding {
    param(
        [Parameter(Mandatory = $true)][string]$OutputPath,
        [Parameter(Mandatory = $true)][pscustomobject]$Row
    )

    $Row | Export-Csv -Path $OutputPath -NoTypeInformation -Append
}

Install-RequiredModule -Name "PnP.PowerShell"

New-DirectoryIfMissing -Path $DocumentsOutputFolder
New-DirectoryIfMissing -Path $OtherOutputFolder

$RootUrl = $RootUrl.TrimEnd('/')
$AdminUrl = $AdminUrl.TrimEnd('/')

$adminConn = Connect-ValidatedPnP -Url $AdminUrl -ClientId $ClientId -AdminConnection
$rootConn = Connect-ValidatedPnP -Url $RootUrl -ClientId $ClientId

Write-Host "`nCollecting tenant sites (excluding OneDrive personal sites)..."
$tenantSites = Get-PnPTenantSite -Connection $adminConn -Detailed
$tenantSites = $tenantSites |
Where-Object {
    $_.Url -and
    $_.Url -notmatch '-my\.sharepoint\.com' -and
    $_.Template -notlike 'SPSPERS*'
} |
Sort-Object Url -Unique

if ($tenantSites.Count -eq 0) {
    throw "No eligible SharePoint site collections found (excluding OneDrive)."
}

$selectedScope = Select-SiteForScope -Sites $tenantSites
Write-Host "`nSelected scope site: $($selectedScope.Url)"

$showRootInventory = Get-BooleanPrompt -Message "Show root subsite inventory under $RootUrl?" -DefaultYes $false
if ($showRootInventory) {
    Write-Host ("`nRoot subsite inventory under {0}:" -f $RootUrl)
    try {
        $rootSubsites = Get-PnPSubWeb -Connection $rootConn -Recurse -Includes Url, Title
        if ($rootSubsites.Count -eq 0) {
            Write-Host "  (none found)"
        }
        else {
            foreach ($web in $rootSubsites | Sort-Object Url) {
                $title = if ($web.Title) { $web.Title } else { '(no title)' }
                Write-Host "  $($web.Url)  |  $title"
            }
        }
    }
    catch {
        Write-Warning "Could not enumerate root subsites: $($_.Exception.Message)"
    }
}
else {
    Write-Host "Root subsite inventory skipped."
}

$oldBaseInput = Read-Host "Old base URL to find (press Enter to use selected scope URL) [$($selectedScope.Url)]"
$OldBaseUrl = if ([string]::IsNullOrWhiteSpace($oldBaseInput)) {
    $selectedScope.Url.TrimEnd('/')
}
else {
    $oldBaseInput.TrimEnd('/')
}

if (-not $OldBaseUrl.StartsWith($selectedScope.Url.TrimEnd('/'), [System.StringComparison]::OrdinalIgnoreCase)) {
    throw "Old base URL must stay within the selected scope site collection for this run."
}

$NewBaseUrl = Read-Host "New base URL (required)"
if ([string]::IsNullOrWhiteSpace($NewBaseUrl)) {
    throw "New base URL is required."
}
$NewBaseUrl = $NewBaseUrl.TrimEnd('/')

$scanOther = Get-BooleanPrompt -Message "After documents, scan additional content (site pages/list text fields) to a separate CSV?" -DefaultYes $false

$scanDocumentContent = $false
if ($EnableDocumentContentScan) {
    $scanDocumentContent = Get-BooleanPrompt -Message "Document embedded-link scanning reads/downloads file content and may update access analytics/last-access signals. Continue document content scan?" -DefaultYes $false
}
else {
    Write-Warning "Document content scan is disabled by default. Use -EnableDocumentContentScan to allow this step."
}

$runStamp = Get-Date -Format "yyyyMMdd-HHmmss"
$docReportPath = Join-Path -Path $DocumentsOutputFolder -ChildPath ("Documents-BreakingUrls-" + $runStamp + ".csv")
$otherReportPath = Join-Path -Path $OtherOutputFolder -ChildPath ("OtherContent-BreakingUrls-" + $runStamp + ".csv")

$oldUri = [Uri]$OldBaseUrl
$TenantBase = $oldUri.GetLeftPart([UriPartial]::Authority)
$OldBasePath = $oldUri.AbsolutePath.TrimEnd('/')

Write-Host "`nAudit configuration:"
Write-Host "  Selected scope site:  $($selectedScope.Url)"
Write-Host "  Old base URL:         $OldBaseUrl"
Write-Host "  New base URL:         $NewBaseUrl"
Write-Host "  Documents report:     $docReportPath"
Write-Host "  Document content scan: $scanDocumentContent"
if ($scanOther) {
    Write-Host "  Other report:         $otherReportPath"
}

$stats = [PSCustomObject]@{
    SitesVisited = 0
    WebsVisited = 0
    LibrariesVisited = 0
    DocumentsRead = 0
    DocumentLinkRows = 0
    OtherLinkRows = 0
    DocumentsSkippedLegacy = 0
    DocumentsSkippedUnsupported = 0
}

$permissionCache = @{}
$folderNameCache = @{}

$scanSiteCollections = @($tenantSites | Sort-Object Url)

Write-Host "`nStarting tenant-wide document scan (SharePoint only, no OneDrive)..."
if ($scanDocumentContent) {
foreach ($site in $scanSiteCollections) {
    $siteUrl = $site.Url.TrimEnd('/')
    $stats.SitesVisited++
    Write-Host "`n[$($stats.SitesVisited)/$($scanSiteCollections.Count)] Site collection: $siteUrl"

    $siteConn = $null
    try {
        $siteConn = Connect-ValidatedPnP -Url $siteUrl -ClientId $ClientId
    }
    catch {
        Write-Warning "Skipping site collection due to connection/auth failure: $siteUrl"
        continue
    }

    $webUrls = New-Object System.Collections.Generic.List[string]
    try {
        $rootWeb = Get-PnPWeb -Connection $siteConn -Includes Url
        $webUrls.Add($rootWeb.Url)
        $subs = Get-PnPSubWeb -Connection $siteConn -Recurse -Includes Url
        foreach ($s in $subs) { $webUrls.Add($s.Url) }
    }
    catch {
        Write-Warning ("Could not enumerate all webs in {0}: {1}" -f $siteUrl, $_.Exception.Message)
        continue
    }

    foreach ($webUrl in $webUrls | Sort-Object -Unique) {
        $stats.WebsVisited++
        Write-Host "  Web: $webUrl"

        $webConn = $null
        try {
            $webConn = Connect-ValidatedPnP -Url $webUrl -ClientId $ClientId
        }
        catch {
            Write-Warning "  Skipping web due to connection/auth failure: $webUrl"
            continue
        }

        $web = $null
        try {
            $web = Get-PnPWeb -Connection $webConn -Includes Id, Url
        }
        catch {
            Write-Warning "  Could not read web metadata for $webUrl"
            continue
        }

        $lists = $null
        try {
            $lists = Get-PnPList -Connection $webConn -Includes BaseTemplate, Hidden, Title, ItemCount, Id, HasUniqueRoleAssignments, RootFolder.ServerRelativeUrl
        }
        catch {
            Write-Warning "  Could not read lists for $webUrl"
            continue
        }

        foreach ($list in $lists) {
            if ($list.Hidden) { continue }
            if ($SkipListTitles -contains $list.Title) { continue }
            if ($list.ItemCount -eq 0) { continue }
            if ($list.BaseTemplate -ne 101) { continue }

            $stats.LibrariesVisited++
            Write-Host "    Library: $($list.Title)"

            $libraryRootPath = $list.RootFolder.ServerRelativeUrl.TrimEnd('/')

            $fields = @('FileRef', 'FileLeafRef', 'FileDirRef', 'UniqueId', 'FSObjType', 'Author', 'Editor', 'ID')
            $items = $null

            try {
                $items = Get-PnPListItem -List $list.Title -Connection $webConn -PageSize 500 -Fields $fields
            }
            catch {
                Write-Warning "    Could not read items for library $($list.Title): $($_.Exception.Message)"
                continue
            }

            foreach ($item in $items) {
                if ($item.FieldValues['FSObjType'] -eq 1) {
                    continue
                }

                $fileRef = [string]$item.FieldValues['FileRef']
                $fileLeaf = [string]$item.FieldValues['FileLeafRef']
                if ([string]::IsNullOrWhiteSpace($fileRef) -or [string]::IsNullOrWhiteSpace($fileLeaf)) {
                    continue
                }

                $stats.DocumentsRead++

                $ext = [System.IO.Path]::GetExtension($fileLeaf).ToLower()
                if ($LegacyExtensions -contains $ext) {
                    $stats.DocumentsSkippedLegacy++
                    continue
                }
                if (-not ($SupportedExtensions -contains $ext)) {
                    $stats.DocumentsSkippedUnsupported++
                    continue
                }

                $docUrl = Get-DocumentBrowserUrl -TenantBase $TenantBase -FileRef $fileRef -Extension $ext -UniqueId $item.FieldValues['UniqueId']
                $folderUrl = Get-FolderBrowserUrl -TenantBase $TenantBase -LibraryRootPath $libraryRootPath -ListId $list.Id -FileDirRef $item.FieldValues['FileDirRef']
                $folderName = Get-FolderDisplayName -List $list -FolderPath $item.FieldValues['FileDirRef'] -Connection $webConn -Cache $folderNameCache

                $authorValue = $item.FieldValues['Author']
                $createdBy = if ($authorValue -and $authorValue.LookupValue) { $authorValue.LookupValue } else { '' }

                $editorValue = $item.FieldValues['Editor']
                $lastModifiedBy = if ($editorValue -and $editorValue.LookupValue) { $editorValue.LookupValue } else { '' }

                $ownerPermissions = Get-EffectiveOwnerPermissions -Item $item -List $list -Web $web -Cache $permissionCache

                $refs = @()
                try {
                    $ms = Get-PnPFile -Url $fileRef -AsMemoryStream -Connection $webConn
                    if ($ext -eq '.pdf') {
                        $refs = Get-PdfReferences -Bytes $ms.ToArray()
                    }
                    else {
                        $refs = Get-OfficeDocReferences -Stream $ms
                    }
                }
                catch {
                    continue
                }

                foreach ($ref in $refs) {
                    if (-not (Test-IsBreakingUrl -Value $ref.Value -OldBaseUrl $OldBaseUrl -TenantBase $TenantBase -OldBasePath $OldBasePath)) {
                        continue
                    }

                    $row = [PSCustomObject]@{
                        'Document URL' = $docUrl
                        'Document Folder URL' = $folderUrl
                        'Document Folder Name' = $folderName
                        'Document File Name' = $fileLeaf
                        'Created By' = $createdBy
                        'Last Modified' = $lastModifiedBy
                        'Permissions' = $ownerPermissions
                        'Every Embedded URL that will break' = $ref.Value
                        'Old Base URL' = $OldBaseUrl
                        'New Base URL' = $NewBaseUrl
                        'Site Collection Containing Document' = $siteUrl
                        'Web Containing Document' = $webUrl
                    }

                    Write-DocumentFinding -OutputPath $docReportPath -Row $row
                    $stats.DocumentLinkRows++
                }
            }
        }
    }
}
}
else {
    Write-Warning "Document content scan was skipped to preserve last-access signal integrity. Documents CSV may be empty/not created for this run."
}

if ($scanOther) {
    Write-Host "`nStarting optional scan for additional content..."

    foreach ($site in $scanSiteCollections) {
        $siteUrl = $site.Url.TrimEnd('/')
        Write-Host "`nOther-content scan site collection: $siteUrl"

        $siteConn = $null
        try {
            $siteConn = Connect-ValidatedPnP -Url $siteUrl -ClientId $ClientId
        }
        catch {
            Write-Warning "Skipping other-content scan for $siteUrl due to auth/connection failure."
            continue
        }

        $webUrls = New-Object System.Collections.Generic.List[string]
        try {
            $rootWeb = Get-PnPWeb -Connection $siteConn -Includes Url
            $webUrls.Add($rootWeb.Url)
            $subs = Get-PnPSubWeb -Connection $siteConn -Recurse -Includes Url
            foreach ($s in $subs) { $webUrls.Add($s.Url) }
        }
        catch {
            Write-Warning "Could not enumerate webs for other-content scan in $siteUrl"
            continue
        }

        foreach ($webUrl in $webUrls | Sort-Object -Unique) {
            $webConn = $null
            try {
                $webConn = Connect-ValidatedPnP -Url $webUrl -ClientId $ClientId
            }
            catch {
                continue
            }

            $lists = $null
            try {
                $lists = Get-PnPList -Connection $webConn -Includes BaseTemplate, Hidden, Title, ItemCount, DefaultDisplayFormUrl, RootFolder.ServerRelativeUrl
            }
            catch {
                continue
            }

            foreach ($list in $lists) {
                if ($list.Hidden) { continue }
                if ($list.ItemCount -eq 0) { continue }
                if ($list.BaseTemplate -eq 101) { continue }

                $fields = $null
                try {
                    $fields = Get-PnPField -List $list.Title -Connection $webConn
                }
                catch {
                    continue
                }

                $textFields = @($fields | Where-Object {
                        -not $_.Hidden -and -not $_.ReadOnlyField -and $_.TypeAsString -in @('Text', 'Note', 'URL')
                    })

                if ($textFields.Count -eq 0) {
                    continue
                }

                $items = $null
                try {
                    $items = Get-PnPListItem -List $list.Title -Connection $webConn -PageSize 200
                }
                catch {
                    continue
                }

                foreach ($item in $items) {
                    $itemId = $item.Id
                    $itemTitle = [string]$item.FieldValues['Title']

                    $itemUrl = ''
                    if ($item.FieldValues.ContainsKey('FileRef') -and $item.FieldValues['FileRef']) {
                        $itemUrl = $TenantBase + [string]$item.FieldValues['FileRef']
                    }
                    elseif ($list.DefaultDisplayFormUrl) {
                        $itemUrl = $TenantBase + $list.DefaultDisplayFormUrl + "?ID=$itemId"
                    }

                    foreach ($field in $textFields) {
                        $raw = $item.FieldValues[$field.InternalName]
                        if (-not $raw) {
                            continue
                        }

                        $textValues = New-Object System.Collections.Generic.List[string]

                        if ($raw -is [string]) {
                            $textValues.Add($raw)
                        }
                        elseif ($raw -is [Microsoft.SharePoint.Client.FieldUrlValue]) {
                            if ($raw.Url) { $textValues.Add($raw.Url) }
                            if ($raw.Description) { $textValues.Add($raw.Description) }
                        }
                        else {
                            $textValues.Add([string]$raw)
                        }

                        foreach ($text in $textValues) {
                            foreach ($url in (Find-UrlsInText -Text $text)) {
                                if (-not (Test-IsBreakingUrl -Value $url -OldBaseUrl $OldBaseUrl -TenantBase $TenantBase -OldBasePath $OldBasePath)) {
                                    continue
                                }

                                $row = [PSCustomObject]@{
                                    'Content Type' = 'ListOrPageItem'
                                    'Site Collection' = $siteUrl
                                    'Web URL' = $webUrl
                                    'List Title' = $list.Title
                                    'Item ID' = $itemId
                                    'Item Title' = $itemTitle
                                    'Item URL' = $itemUrl
                                    'Field Name' = $field.InternalName
                                    'Every Embedded URL that will break' = $url
                                    'Old Base URL' = $OldBaseUrl
                                    'New Base URL' = $NewBaseUrl
                                }

                                Write-OtherFinding -OutputPath $otherReportPath -Row $row
                                $stats.OtherLinkRows++
                            }
                        }
                    }
                }
            }
        }
    }
}

Write-Host "`nAudit complete."
Write-Host "  Site collections visited:       $($stats.SitesVisited)"
Write-Host "  Webs visited:                   $($stats.WebsVisited)"
Write-Host "  Libraries visited:              $($stats.LibrariesVisited)"
Write-Host "  Documents inspected:            $($stats.DocumentsRead)"
Write-Host "  Document rows written:          $($stats.DocumentLinkRows)"
Write-Host "  Legacy docs skipped:            $($stats.DocumentsSkippedLegacy)"
Write-Host "  Unsupported docs skipped:       $($stats.DocumentsSkippedUnsupported)"
if ($scanOther) {
    Write-Host "  Other-content rows written:     $($stats.OtherLinkRows)"
}
Write-Host "`nDocuments CSV: $docReportPath"
if ($scanOther) {
    Write-Host "Other CSV:     $otherReportPath"
}

