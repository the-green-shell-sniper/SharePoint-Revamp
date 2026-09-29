<#
.SYNOPSIS
    Builds a per-site old-URL -> new-URL lookup for documents, for use by the
    SharePoint URL Fixer tool (and anyone else) when a bookmark can't be
    fixed by pattern-matching alone.

.DESCRIPTION
    The URL Fixer HTML tool can rewrite a document's URL by pattern-matching
    (renaming site path segments, inserting /sites, etc.), but it cannot
    invent a migrated document's new item GUID - that GUID is assigned fresh
    when the document is copied into the new site and has no mathematical
    relationship to the old GUID. Links that carry the old GUID as their only
    way to identify the file (SharePoint's own "resolved" Doc.aspx address-bar
    form) can only be fixed by actually looking up the new copy.

    This script does that lookup, live, for every document in a chosen
    classic site/subsite tree:
      1) Picks a site collection and one or more subsites, the same way
         Get-FixItListURLsv5.ps1 does (menu selection + cached scannability
         probing so empty/system webs are skipped automatically).
      2) Reads the Old Site URL / New Site URL CSV (the same format as
         OldSiteURL-NewSiteURL-Lookup.csv in the repo root) to know where
         each chosen site's content moved to.
      3) Enumerates every document library and file in the old site tree.
      4) For each file, builds the old "Copy Link" URL and (for Office/PDF
         types) the old "resolved" URL, using the file's real UniqueId.
      5) Computes the expected new document path and queries the NEW site
         live to confirm the file exists there and to read its real (new)
         UniqueId, then builds the matching new "Copy Link" and "resolved"
         URLs from that real GUID.
      6) Writes one CSV per top-level site collection scanned, so a single
         run never mixes multiple sites' documents into one report.

    Rows where the new document can't be found (folder renamed during
    migration, file not yet migrated, etc.) are still written, with
    NewFound = No and a Notes explanation, so they can be chased down
    manually - this script does not guess at folder-level renames beyond
    what the Old/New Site URL CSV says about the site itself.

.PARAMETER RootUrl
    Tenant root URL. Default https://whallc.sharepoint.com

.PARAMETER AdminUrl
    SharePoint admin center URL, used to enumerate site collections.

.PARAMETER ClientId
    Entra ID App Registration Client ID for interactive sign-in.

.PARAMETER UrlMappingCsvPath
    Path to the Old Site URL / New Site URL CSV (headers "Old Site URL" and
    "New Site URL"). Defaults to the CSV in the server's Support-Files folder.

.PARAMETER OutputFolder
    Parent folder for document lookup output files.

.PARAMETER MasterFolder
    Folder containing the per-site [sitename]-master.csv lookup inventories.
    If omitted, defaults to the server's DocumentUrlLookup-OutputFiles\Masters folder.

.PARAMETER TempMasterFolder
    Folder where timestamped temporary master scan files are written.
    If omitted, defaults to the server's DocumentUrlLookup-OutputFiles\Masters\Temp folder.

.PARAMETER FixerOutputFolder
    Folder where filtered [sitename]-fixer.csv files are written. If omitted,
    defaults to the server's DocumentUrlLookup-OutputFiles\FixerFiles folder.

.PARAMETER EmbeddedLinksFolder
    Folder containing per-site embedded-link reports and temporary scan files.
    If omitted, defaults to the server's DocumentUrlLookup-OutputFiles\EmbeddedLinks folder.

.PARAMETER SitesMasterFolder
    Master output folder used for webs whose URL path starts with /sites.

.PARAMETER SitesFixerOutputFolder
    Fixer output folder used for webs whose URL path starts with /sites.

.PARAMETER SitesEmbeddedLinksFolder
    Embedded-links output folder used for webs whose URL path starts with /sites.

.PARAMETER StateRootFolder
    Folder used to cache the subsite scannability inventory between runs.

.PARAMETER UsePersistedInventoryCache
    Reuse the cached scannability results from a prior run instead of
    re-probing every subsite. Default true.

.PARAMETER SkipAlreadyConfirmedDocuments
    Skip the live lookup against the new site for a document already confirmed
    found there in a prior run (same old-side Modified date). Set to $false to
    force re-verifying every document, e.g. if you suspect files moved again.
    Default true.

.EXAMPLE
    .\Get-DocumentUrlLookup.ps1 -ClientId "..."
#>

[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost', '', Justification = 'Interactive script intentionally writes status and prompts to the host.')]
[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$RootUrl = "https://whallc.sharepoint.com",

    [Parameter(Mandatory = $false)]
    [string]$AdminUrl = "https://whallc-admin.sharepoint.com",

    [Parameter(Mandatory = $false)]
    [string]$ClientId,

    # Defaults to the relocated server CSV; set blank to select or enter a path at runtime.
    [Parameter(Mandatory = $false)]
    [string]$UrlMappingCsvPath = "",

    [Parameter(Mandatory = $false)]
    [string]$OutputFolder = "C:\Users\!bposey\OneDrive - Womens Healthcare Associates LLC\Gary Henderson's files - Project Files_bposey\DocumentUrlLookup-OutputFiles",

    [Parameter(Mandatory = $false)]
    [string]$MasterFolder = "C:\Users\!bposey\OneDrive - Womens Healthcare Associates LLC\Gary Henderson's files - Project Files_bposey\DocumentUrlLookup-OutputFiles\Masters",

    [Parameter(Mandatory = $false)]
    [string]$TempMasterFolder = "C:\Users\!bposey\OneDrive - Womens Healthcare Associates LLC\Gary Henderson's files - Project Files_bposey\DocumentUrlLookup-OutputFiles\Masters\Temp",

    [Parameter(Mandatory = $false)]
    [string]$FixerOutputFolder = "C:\Users\!bposey\OneDrive - Womens Healthcare Associates LLC\Gary Henderson's files - Project Files_bposey\DocumentUrlLookup-OutputFiles\FixerFiles",

    [Parameter(Mandatory = $false)]
    [string]$EmbeddedLinksFolder = "C:\Users\!bposey\OneDrive - Womens Healthcare Associates LLC\Gary Henderson's files - Project Files_bposey\DocumentUrlLookup-OutputFiles\EmbeddedLinks",

    [Parameter(Mandatory = $false)]
    [string]$SitesMasterFolder = "C:\Users\!bposey\OneDrive - Womens Healthcare Associates LLC\Gary Henderson's files - Project Files_bposey\DocumentUrlLookup-OutputFiles\Sites-Masters",

    [Parameter(Mandatory = $false)]
    [string]$SitesTempMasterFolder = "C:\Users\!bposey\OneDrive - Womens Healthcare Associates LLC\Gary Henderson's files - Project Files_bposey\DocumentUrlLookup-OutputFiles\Sites-Masters\Temp",

    [Parameter(Mandatory = $false)]
    [string]$SitesFixerOutputFolder = "C:\Users\!bposey\OneDrive - Womens Healthcare Associates LLC\Gary Henderson's files - Project Files_bposey\DocumentUrlLookup-OutputFiles\Sites-FixerFiles",

    [Parameter(Mandatory = $false)]
    [string]$SitesEmbeddedLinksFolder = "C:\Users\!bposey\OneDrive - Womens Healthcare Associates LLC\Gary Henderson's files - Project Files_bposey\DocumentUrlLookup-OutputFiles\Sites-EmbeddedLinks",

    [Parameter(Mandatory = $false)]
    [string]$StateRootFolder = "C:\Users\!bposey\OneDrive - Womens Healthcare Associates LLC\Gary Henderson's files - Project Files_bposey\Checkpoints-and-Caches",

    [Parameter(Mandatory = $false)]
    [bool]$UsePersistedInventoryCache = $true,

    # Skips the live lookup against the new site for a document that was already
    # confirmed found there in a prior run and whose old-side Modified date hasn't
    # changed since - this is the actual expensive part of a re-run (network calls),
    # not the CSV write. Set to $false to force re-verifying every document.
    [Parameter(Mandatory = $false)]
    [bool]$SkipAlreadyConfirmedDocuments = $true,

    # Scan Office/PDF/page content and human-authored list fields for embedded links.
    # File content is downloaded, so this is off by default and prompted for.
    [Parameter(Mandatory = $false)]
    [switch]$EnableDocumentContentScan
)

$ErrorActionPreference = "Stop"

$ScriptSchemaVersion = 'DocumentUrlLookup-v1'
$InventoryCacheSchemaVersion = 'v2'
$PdfPigAssemblyFolder = 'C:\Users\!bposey\Desktop\PdfPig'
$SkipListTitles = @('Form Templates', 'Preservation Hold Library')
$FixerRequiredExtensions = @('.docx', '.doc', '.pdf', '.pptx', '.xls', '.xlsx', '.ppt', '.vsdx', '.vsd', '.png', '.aspx', '.one')
$FileLookupListBaseTemplates = @(101, 119, 850)

# Office web-viewer prefix per extension: /:w:/r/... Word, :x: Excel, :p:
# PowerPoint, :b: PDF. Same letter is reused in the ?d=<letter><guid> form.
$DocLinkPrefixByExt = @{
    '.docx' = 'w'; '.doc' = 'w'
    '.xlsx' = 'x'; '.xls' = 'x'
    '.pptx' = 'p'; '.ppt' = 'p'
    '.pdf'  = 'b'
}

# --- Shared scaffolding, ported from Get-FixItListURLsv5.ps1 -----------------

function Set-ScriptWindowTitle {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param([Parameter(Mandatory = $true)][string]$ScriptName)

    try {
        $startedAt = Get-Date
        $pacificTz = [System.TimeZoneInfo]::FindSystemTimeZoneById('Pacific Standard Time')
        $startedAtPacific = if ($startedAt.Kind -eq [System.DateTimeKind]::Utc) {
            [System.TimeZoneInfo]::ConvertTimeFromUtc($startedAt, $pacificTz)
        }
        else {
            [System.TimeZoneInfo]::ConvertTime($startedAt, $pacificTz)
        }

        if ($PSCmdlet.ShouldProcess($ScriptName, 'Set console window title')) {
            $Host.UI.RawUI.WindowTitle = ('{0} | Started {1} PT' -f $ScriptName, $startedAtPacific.ToString('yyyy-MM-dd HH:mm:ss'))
        }
    }
    catch {
        Write-Verbose 'Could not set the console window title.'
    }
}

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
                Url               = $Url
                Interactive       = $true
                PersistLogin      = $true
                ReturnConnection  = $true
            }
            if ($ClientId) {
                $connectParams.ClientId = $ClientId
            }

            $conn = Connect-PnPOnline @connectParams

            try {
                $ctx = Get-PnPContext -Connection $conn
                if ($ctx) {
                    $ctx.RequestTimeout = -1
                }
            }
            catch {
                Write-Verbose "Could not update CSOM request timeout on the current connection."
            }

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

function Update-PnPConnection {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [Parameter(Mandatory = $true)][string]$Url,
        [Parameter(Mandatory = $false)][string]$ClientId,
        [Parameter(Mandatory = $false)][switch]$AdminConnection
    )

    Write-Warning "Refreshing connection for $Url after a timeout or transient failure."

    if ($PSCmdlet.ShouldProcess($Url, 'Refresh PnP connection')) {
        return Connect-ValidatedPnP -Url $Url -ClientId $ClientId -AdminConnection:$AdminConnection
    }
}

function Test-IsTimeoutFailure {
    param([Parameter(Mandatory = $true)]$ErrorRecord)

    $msg = [string]$ErrorRecord.Exception.Message
    return ($msg -match '(?i)timed out|timeout|request was canceled|HttpClient\.Timeout|The operation has timed out')
}

function Invoke-WithRetry {
    param(
        [Parameter(Mandatory = $true)][scriptblock]$Operation,
        [Parameter(Mandatory = $false)][scriptblock]$OnTimeout,
        [Parameter(Mandatory = $false)][int]$MaxAttempts = 3,
        [Parameter(Mandatory = $false)][string]$Label = "operation"
    )

    for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
        try {
            return & $Operation
        }
        catch {
            $isTimeout = Test-IsTimeoutFailure -ErrorRecord $_
            if ((-not $isTimeout) -or ($attempt -eq $MaxAttempts)) {
                throw
            }

            Write-Warning ("{0} timed out (attempt {1}/{2}); refreshing and retrying..." -f $Label, $attempt, $MaxAttempts)
            if ($OnTimeout) {
                & $OnTimeout
            }
        }
    }
}

function Invoke-PnPConnectionOperation {
    param(
        [Parameter(Mandatory = $true)][string]$Url,
        [Parameter(Mandatory = $false)][string]$ClientId,
        [Parameter(Mandatory = $true)]$Connection,
        [Parameter(Mandatory = $true)][scriptblock]$Operation,
        [Parameter(Mandatory = $false)][string]$Label = 'PnP operation',
        [Parameter(Mandatory = $false)][int]$MaxAttempts = 4,
        [Parameter(Mandatory = $false)][switch]$AdminConnection
    )

    $targetUrl = $Url
    $targetClientId = $ClientId
    $operationToRun = $Operation
    $useAdminConnection = $AdminConnection
    $state = @{ Connection = $Connection }
    $result = Invoke-WithRetry -Label $Label -MaxAttempts $MaxAttempts -OnTimeout {
        $state.Connection = Update-PnPConnection -Url $targetUrl -ClientId $targetClientId -AdminConnection:$useAdminConnection
    } -Operation {
        & $operationToRun $state.Connection
    }

    return [PSCustomObject]@{
        Connection = $state.Connection
        Result     = $result
    }
}

function New-DirectoryIfMissing {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param([Parameter(Mandatory = $true)][string]$Path)

    if (-not (Test-Path -Path $Path)) {
        if ($PSCmdlet.ShouldProcess($Path, 'Create directory')) {
            New-Item -ItemType Directory -Path $Path -Force | Out-Null
        }
    }
}

function Get-SlugValue {
    param([Parameter(Mandatory = $true)][AllowEmptyString()][string]$Text)

    $slug = [regex]::Replace($Text, '[^A-Za-z0-9]', '')
    if ([string]::IsNullOrWhiteSpace($slug)) {
        return ''
    }

    return $slug
}

function Get-SafeFileSlug {
    param(
        [Parameter(Mandatory = $true)][AllowEmptyString()][string]$Text,
        [Parameter(Mandatory = $false)][int]$MaxLength = 72,
        [Parameter(Mandatory = $false)][string]$Fallback = 'SelectedSite'
    )

    $slug = Get-SlugValue -Text $Text
    if ([string]::IsNullOrWhiteSpace($slug)) {
        $slug = $Fallback
    }

    if ($slug.Length -gt $MaxLength) {
        return $slug.Substring(0, $MaxLength)
    }

    return $slug
}

function Get-Sha256Hex {
    param([Parameter(Mandatory = $true)][string]$Text)

    $sha = [System.Security.Cryptography.SHA256]::Create()
    try {
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($Text)
        $hash = $sha.ComputeHash($bytes)
        return ([System.BitConverter]::ToString($hash) -replace '-', '').ToLowerInvariant()
    }
    finally {
        $sha.Dispose()
    }
}

function Get-InventorySubsiteCachePath {
    param(
        [Parameter(Mandatory = $true)][string]$StateRoot,
        [Parameter(Mandatory = $true)][string]$ScopeUrl,
        [Parameter(Mandatory = $true)][string]$SchemaVersion
    )

    $fingerprint = (Get-Sha256Hex -Text ($ScopeUrl.TrimEnd('/').ToLowerInvariant() + '|' + $SchemaVersion + '|' + $InventoryCacheSchemaVersion)).Substring(0, 16)
    return Join-Path -Path $StateRoot -ChildPath ("Get-DocumentUrlLookup_InventoryCache_{0}.json" -f $fingerprint)
}

function Get-InventorySubsiteCache {
    param([Parameter(Mandatory = $true)][string]$Path)

    if (-not (Test-Path -Path $Path)) {
        return @{}
    }

    try {
        $raw = Get-Content -Path $Path -Raw -Encoding UTF8
        if ([string]::IsNullOrWhiteSpace($raw)) {
            return @{}
        }

        $parsed = $raw | ConvertFrom-Json
        $cache = @{}
        if ($parsed -and $parsed.PSObject.Properties['Entries']) {
            foreach ($entry in @($parsed.Entries)) {
                if (-not $entry -or -not $entry.Url) {
                    continue
                }

                $key = ([string]$entry.Url).TrimEnd('/').ToLowerInvariant()
                $cache[$key] = [PSCustomObject]@{
                    Url         = ([string]$entry.Url).TrimEnd('/')
                    Title       = [string]$entry.Title
                    IsScannable = [bool]$entry.IsScannable
                    Reason      = [string]$entry.Reason
                }
            }
        }

        return $cache
    }
    catch {
        Write-Verbose "Could not load inventory cache from ${Path}: $($_.Exception.Message)"
        return @{}
    }
}

function Save-InventorySubsiteCache {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][array]$Entries,
        [Parameter(Mandatory = $true)][string]$ScopeUrl,
        [Parameter(Mandatory = $true)][string]$SchemaVersion
    )

    try {
        [PSCustomObject][ordered]@{
            SchemaVersion          = $SchemaVersion
            InventorySchemaVersion = $InventoryCacheSchemaVersion
            ScopeUrl               = $ScopeUrl
            SavedAt                = (Get-Date).ToString('o')
            Entries                = $Entries
        } | ConvertTo-Json -Depth 6 | Set-Content -Path $Path -Encoding UTF8 -Force
    }
    catch {
        Write-Verbose "Could not save inventory cache to ${Path}: $($_.Exception.Message)"
    }
}

function Test-IsRecycleBinContainer {
    param(
        [Parameter(Mandatory = $false)][string]$ListTitle,
        [Parameter(Mandatory = $false)][string]$RootFolderPath
    )

    $title = if ($ListTitle) { $ListTitle.Trim() } else { '' }
    $path = if ($RootFolderPath) { $RootFolderPath.ToLowerInvariant() } else { '' }

    if ($title -match '(?i)\brecycle\s*bin\b|\bsecond[-\s]*stage\b') {
        return $true
    }

    if ($path -match '(?i)/recycle(?:%20|\s)*bin(?:/|$)' -or
        $path -match '(?i)/site(?:%20|\s)*collection(?:%20|\s)*recycle(?:%20|\s)*bin(?:/|$)' -or
        $path -match '(?i)/second(?:%20|\s|-)*stage(?:/|$)' -or
        $path -match '(?i)/_layouts/15/recyclebin\.aspx') {
        return $true
    }

    return $false
}

function Test-IsSiteContentsInventoryList {
    param([Parameter(Mandatory = $true)]$List)

    if (-not $List) { return $false }
    if ($List.Hidden) { return $false }

    $rootFolderPath = if ($List.RootFolder -and $List.RootFolder.ServerRelativeUrl) { [string]$List.RootFolder.ServerRelativeUrl } else { '' }
    if (Test-IsRecycleBinContainer -ListTitle $List.Title -RootFolderPath $rootFolderPath) { return $false }
    if ($SkipListTitles -contains $List.Title) { return $false }
    if ($List.ItemCount -eq 0) { return $false }

    return $true
}

function Test-IsFileLookupList {
    param([Parameter(Mandatory = $true)]$List)

    if (-not (Test-IsSiteContentsInventoryList -List $List)) { return $false }
    return ($FileLookupListBaseTemplates -contains [int]$List.BaseTemplate)
}

function Test-IsInventorySubsiteScannable {
    param(
        [Parameter(Mandatory = $true)][string]$WebUrl,
        [Parameter(Mandatory = $true)][string]$RootHost,
        [Parameter(Mandatory = $false)][string]$ClientId
    )

    try {
        $webHost = ([Uri]$WebUrl).Host.ToLowerInvariant()
        if (-not $webHost.Equals($RootHost, [System.StringComparison]::OrdinalIgnoreCase)) {
            return [PSCustomObject]@{ IsScannable = $false; Reason = 'Different host than tenant root (likely app/deprecated web).' }
        }
    }
    catch {
        return [PSCustomObject]@{ IsScannable = $false; Reason = 'Invalid URL format.' }
    }

    $webConn = $null
    try {
        $webConn = Connect-ValidatedPnP -Url $WebUrl -ClientId $ClientId
    }
    catch {
        return [PSCustomObject]@{ IsScannable = $false; Reason = 'Connection/auth failed for this web.' }
    }

    try {
        $webResult = Invoke-PnPConnectionOperation -Url $WebUrl -ClientId $ClientId -Connection $webConn -Label ('Inventory scannability web metadata: ' + $WebUrl) -MaxAttempts 3 -Operation {
            param($activeConnection)
            Get-PnPWeb -Connection $activeConnection -Includes WebTemplate, Configuration
        }
        $webConn = $webResult.Connection
        $web = $webResult.Result

        $webTemplateName = if ($web -and $web.WebTemplate) { [string]$web.WebTemplate } else { '' }
        if (-not [string]::IsNullOrWhiteSpace($webTemplateName) -and ($webTemplateName -match '(?i)^APP')) {
            return [PSCustomObject]@{ IsScannable = $false; Reason = 'App web template is not scanned by this script.' }
        }

        $listResult = Invoke-PnPConnectionOperation -Url $WebUrl -ClientId $ClientId -Connection $webConn -Label ('Inventory scannability lists: ' + $WebUrl) -MaxAttempts 3 -Operation {
            param($activeConnection)
            Get-PnPList -Connection $activeConnection -Includes BaseTemplate, Hidden, Title, ItemCount, RootFolder.ServerRelativeUrl
        }
        $lists = @($listResult.Result)

        foreach ($list in $lists) {
            if (-not (Test-IsSiteContentsInventoryList -List $list)) { continue }

            return [PSCustomObject]@{ IsScannable = $true; Reason = '' }
        }

        return [PSCustomObject]@{ IsScannable = $false; Reason = 'No non-empty visible Site Contents lists/libraries found.' }
    }
    catch {
        return [PSCustomObject]@{ IsScannable = $false; Reason = ('Scannability probe failed: ' + $_.Exception.Message) }
    }
}

function Select-SiteForScope {
    param([Parameter(Mandatory = $true)][array]$Sites)

    $teamsSites = @($Sites | Where-Object {
            ($_.GroupId -and $_.GroupId -ne [Guid]::Empty) -or $_.Template -like 'GROUP*' -or $_.Template -like 'TEAMCHANNEL*'
        } | Sort-Object Url)

    $sharePointSites = @($Sites | Where-Object {
            -not ((($_.GroupId -and $_.GroupId -ne [Guid]::Empty) -or $_.Template -like 'GROUP*' -or $_.Template -like 'TEAMCHANNEL*'))
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

    $categorySites = if ($firstChoice -eq '1') { $sharePointSites } else { $teamsSites }

    if ($categorySites.Count -eq 0) {
        throw "No sites were found in the selected category."
    }

    Write-Host ""
    for ($i = 0; $i -lt $categorySites.Count; $i++) {
        $kind = if ((($categorySites[$i].GroupId -and $categorySites[$i].GroupId -ne [Guid]::Empty) -or $categorySites[$i].Template -like 'GROUP*' -or $categorySites[$i].Template -like 'TEAMCHANNEL*')) {
            'Teams-connected'
        }
        else {
            'SharePoint'
        }
        $title = if (-not [string]::IsNullOrWhiteSpace([string]$categorySites[$i].Title)) { [string]$categorySites[$i].Title } else { '(no title)' }
        Write-Host ("  {0,4}. [{1}] {2}  |  {3}" -f ($i + 1), $kind, $categorySites[$i].Url, $title)
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

# --- New helpers specific to this script -------------------------------------

function Get-SiteUrlMapFromCsv {
    param([Parameter(Mandatory = $true)][string]$CsvPath)

    if (-not (Test-Path -Path $CsvPath)) {
        throw "Old/New site URL lookup CSV was not found: $CsvPath"
    }

    $rows = @(Import-Csv -Path $CsvPath)
    if ($rows.Count -eq 0) {
        throw "Old/New site URL lookup CSV has no data rows: $CsvPath"
    }

    $columnNames = @($rows[0].PSObject.Properties.Name)
    $oldColumn = @($columnNames | Where-Object { $_ -ieq 'Old Site URL' } | Select-Object -First 1)
    $newColumn = @($columnNames | Where-Object { $_ -ieq 'New Site URL' } | Select-Object -First 1)
    $linkBaseColumn = @($columnNames | Where-Object { $_ -ieq 'New Link Base URL' } | Select-Object -First 1)
    if ($oldColumn.Count -eq 0 -or $newColumn.Count -eq 0) {
        throw "CSV must include headers named 'Old Site URL' and 'New Site URL': $CsvPath"
    }

    $map = New-Object System.Collections.Generic.List[Object]
    foreach ($row in $rows) {
        $oldBase = ([string]$row.($oldColumn[0])).TrimEnd('/')
        $newBase = ([string]$row.($newColumn[0])).TrimEnd('/')
        if ([string]::IsNullOrWhiteSpace($oldBase) -or [string]::IsNullOrWhiteSpace($newBase)) {
            continue
        }

        # Usually the same as NewBase - only differs when a site is being
        # scanned from a staging location but should still resolve to a
        # different public base URL once links are built (e.g. the tenant
        # root's content currently lives at a staging site, but future links
        # should still point at the root).
        $linkBase = if ($linkBaseColumn.Count -gt 0) { ([string]$row.($linkBaseColumn[0])).TrimEnd('/') } else { '' }
        if ([string]::IsNullOrWhiteSpace($linkBase)) {
            $linkBase = $newBase
        }

        $map.Add([PSCustomObject]@{ OldBase = $oldBase; NewBase = $newBase; NewLinkBase = $linkBase })
    }

    # Longest OldBase first, so a nested subsite matches its own row before
    # falling back to a shorter parent-site prefix.
    return @($map | Sort-Object { $_.OldBase.Length } -Descending)
}

function Get-NewSiteBaseForUrl {
    param(
        [Parameter(Mandatory = $true)][string]$Url,
        [Parameter(Mandatory = $true)][array]$SiteUrlMap
    )

    $normalized = $Url.TrimEnd('/')
    foreach ($entry in $SiteUrlMap) {
        $oldBase = $entry.OldBase
        if ($normalized.Equals($oldBase, [System.StringComparison]::OrdinalIgnoreCase) -or
            $normalized.StartsWith(($oldBase + '/'), [System.StringComparison]::OrdinalIgnoreCase)) {
            $suffix = $normalized.Substring($oldBase.Length)
            return [PSCustomObject]@{ OldBase = $oldBase; NewBase = $entry.NewBase; NewLinkBase = $entry.NewLinkBase; Suffix = $suffix }
        }
    }

    return $null
}

function Get-ExactNewSiteBaseForUrl {
    param(
        [Parameter(Mandatory = $true)][string]$Url,
        [Parameter(Mandatory = $true)][array]$SiteUrlMap
    )

    $normalized = $Url.TrimEnd('/')
    foreach ($entry in $SiteUrlMap) {
        if ($normalized.Equals(([string]$entry.OldBase).TrimEnd('/'), [System.StringComparison]::OrdinalIgnoreCase)) {
            return [PSCustomObject]@{
                OldBase     = $entry.OldBase
                NewBase     = $entry.NewBase
                NewLinkBase = $entry.NewLinkBase
                Suffix      = ''
            }
        }
    }

    return $null
}

function Get-EncodedServerRelativePath {
    param([Parameter(Mandatory = $true)][string]$Path)

    return (($Path -split '/') | ForEach-Object { [Uri]::EscapeDataString($_) }) -join '/'
}

function Get-DocumentBrowserUrl {
    param(
        [Parameter(Mandatory = $true)][string]$TenantBase,
        [Parameter(Mandatory = $true)][string]$FileRef,
        [Parameter(Mandatory = $true)][AllowEmptyString()][string]$Extension,
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

# SharePoint's own id=/parent= query values additionally escape "." as %2E,
# which [Uri]::EscapeDataString alone does not do.
function Get-SharePointIdParamValue {
    param([Parameter(Mandatory = $true)][string]$Path)

    return ([Uri]::EscapeDataString($Path) -replace '\.', '%2E')
}

# Word/Excel/PowerPoint resolve through Doc.aspx (site path + sourcedoc GUID);
# PDFs resolve through the containing folder's own AllItems.aspx?id=&parent=.
function Get-ResolvedDocumentUrl {
    param(
        [Parameter(Mandatory = $true)][string]$TenantBase,
        [Parameter(Mandatory = $true)][AllowEmptyString()][string]$SiteServerRelativeUrl,
        [Parameter(Mandatory = $true)][string]$FileRef,
        [Parameter(Mandatory = $true)][string]$FolderRef,
        [Parameter(Mandatory = $true)][string]$FileName,
        [Parameter(Mandatory = $true)][AllowEmptyString()][string]$Extension,
        [Parameter(Mandatory = $false)]$UniqueId
    )

    $prefix = if ([string]::IsNullOrWhiteSpace($Extension)) { $null } else { $DocLinkPrefixByExt[$Extension] }
    if (-not $prefix) {
        return ''
    }

    if ($Extension -eq '.pdf') {
        $idValue = Get-SharePointIdParamValue -Path $FileRef
        $parentValue = Get-SharePointIdParamValue -Path $FolderRef
        $folderAllItems = Get-EncodedServerRelativePath -Path "$FolderRef/Forms/AllItems.aspx"
        return "$TenantBase$folderAllItems" + "?id=$idValue&parent=$parentValue"
    }

    if (-not $UniqueId) {
        return ''
    }

    $guid = ([Guid]$UniqueId).ToString().ToUpperInvariant()
    $encodedFile = [Uri]::EscapeDataString($FileName)
    return "$TenantBase/:${prefix}:/r$SiteServerRelativeUrl/_layouts/15/Doc.aspx?sourcedoc=%7B$guid%7D&file=$encodedFile&action=default&mobileredirect=true"
}

function Get-NormalizedExtension {
    param([Parameter(Mandatory = $true)][string]$FileName)

    $ext = [System.IO.Path]::GetExtension($FileName)
    if ([string]::IsNullOrWhiteSpace($ext)) {
        return ''
    }

    return $ext.Trim().ToLowerInvariant()
}

function Initialize-PdfPig {
    $requiredAssemblies = @(
        'UglyToad.PdfPig.Core.dll',
        'UglyToad.PdfPig.Tokens.dll',
        'UglyToad.PdfPig.Fonts.dll',
        'UglyToad.PdfPig.Tokenization.dll',
        'UglyToad.PdfPig.Package.dll',
        'UglyToad.PdfPig.dll',
        'UglyToad.PdfPig.DocumentLayoutAnalysis.dll'
    )

    $missingAssemblies = @($requiredAssemblies | Where-Object {
            -not (Test-Path -LiteralPath (Join-Path -Path $PdfPigAssemblyFolder -ChildPath $_))
        })
    if ($missingAssemblies.Count -gt 0) {
        throw "PdfPig is not configured. Copy all seven PdfPig .NET Standard 2.0 DLLs to '$PdfPigAssemblyFolder'. Missing: $($missingAssemblies -join ', ')"
    }

    $script:PdfPigAssemblyFolder = $PdfPigAssemblyFolder
    if (-not $script:PdfPigAssemblyResolveHandler) {
        $script:PdfPigAssemblyResolveHandler = [System.ResolveEventHandler]{
            param($resolveSender, $resolveEventArgs)

            $dependencyName = ([System.Reflection.AssemblyName]::new($resolveEventArgs.Name)).Name + '.dll'
            $dependencyPath = Join-Path -Path $script:PdfPigAssemblyFolder -ChildPath $dependencyName
            if (Test-Path -LiteralPath $dependencyPath) {
                return [System.Reflection.Assembly]::LoadFrom($dependencyPath)
            }
            return $null
        }
        [AppDomain]::CurrentDomain.add_AssemblyResolve($script:PdfPigAssemblyResolveHandler)
    }

    foreach ($assemblyName in $requiredAssemblies) {
        $assemblyPath = Join-Path -Path $PdfPigAssemblyFolder -ChildPath $assemblyName
        $alreadyLoaded = [AppDomain]::CurrentDomain.GetAssemblies() | Where-Object {
            try { [string]$_.Location -eq $assemblyPath } catch { $false }
        } | Select-Object -First 1
        if (-not $alreadyLoaded) {
            [System.Reflection.Assembly]::LoadFrom($assemblyPath) | Out-Null
        }
    }
}

# Only the human-visible content parts of the OOXML package (the document/sheet/slide
# bodies) are read - not core.xml/app.xml metadata, not comments - since we only care
# about links a reader would actually see and could click.
function Get-OfficeDocumentHyperlinks {
    param([Parameter(Mandatory = $true)][System.IO.Stream]$Stream)

    $found = New-Object System.Collections.Generic.List[Object]
    $relationsByXml = @{}
    $archive = [System.IO.Compression.ZipArchive]::new($Stream, [System.IO.Compression.ZipArchiveMode]::Read)
    try {
        foreach ($entry in $archive.Entries) {
            if ($entry.FullName -notlike "*_rels/*.rels") { continue }

            $relationshipStream = $entry.Open()
            $relationshipReader = New-Object System.IO.StreamReader($relationshipStream)
            try {
                [xml]$relationshipXml = $relationshipReader.ReadToEnd()
            }
            finally {
                $relationshipReader.Close()
                $relationshipStream.Close()
            }

            try {
                $xmlPath = $entry.FullName -replace '/_rels/([^/]+)\.rels$', '/$1'
                $relationMap = @{}
                foreach ($rel in $relationshipXml.Relationships.Relationship) {
                    if ($rel.TargetMode -eq 'External' -and $rel.Target) {
                        $relationMap[$rel.Id] = [string]$rel.Target
                    }
                }
                if ($relationMap.Count -gt 0) { $relationsByXml[$xmlPath] = $relationMap }
            }
            catch {
                Write-Verbose "Could not parse relationship XML in $($entry.FullName)."
            }
        }

        $sharedStrings = @()
        $sharedStringsEntry = $archive.GetEntry('xl/sharedStrings.xml')
        if ($sharedStringsEntry) {
            $sharedStringsStream = $sharedStringsEntry.Open()
            $sharedStringsReader = New-Object System.IO.StreamReader($sharedStringsStream)
            try {
                [xml]$sharedStringsXml = $sharedStringsReader.ReadToEnd()
                $sharedStringsNs = New-Object System.Xml.XmlNamespaceManager($sharedStringsXml.NameTable)
                $sharedStringsNs.AddNamespace('x', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main')
                foreach ($stringItem in $sharedStringsXml.SelectNodes('//x:si', $sharedStringsNs)) {
                    $sharedStrings += [string]$stringItem.InnerText
                }
            }
            finally {
                $sharedStringsReader.Close()
                $sharedStringsStream.Close()
            }
        }

        foreach ($entry in $archive.Entries) {
            if ($entry.FullName -notmatch '(?:^|/)word/document\.xml$|(?:^|/)xl/worksheets/sheet\d*\.xml$|(?:^|/)ppt/slides/slide\d*\.xml$') { continue }
            if (-not $relationsByXml.ContainsKey($entry.FullName)) { continue }

            $contentStream = $entry.Open()
            $contentReader = New-Object System.IO.StreamReader($contentStream)
            try {
                [xml]$contentXml = $contentReader.ReadToEnd()
            }
            finally {
                $contentReader.Close()
                $contentStream.Close()
            }

            try {
                $ns = New-Object System.Xml.XmlNamespaceManager($contentXml.NameTable)
                $ns.AddNamespace('r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
                $ns.AddNamespace('p', 'http://schemas.openxmlformats.org/presentationml/2006/main')
                $ns.AddNamespace('x', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main')
                foreach ($node in $contentXml.SelectNodes('//*[@r:id]', $ns)) {
                    $relId = $node.GetAttribute('id', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
                    if (-not $relationsByXml[$entry.FullName].ContainsKey($relId)) { continue }

                    $displayText = [string]$node.InnerText
                    if ([string]::IsNullOrWhiteSpace($displayText) -and $node.Attributes['display']) {
                        $displayText = [string]$node.Attributes['display'].Value
                    }
                    if ([string]::IsNullOrWhiteSpace($displayText) -and $entry.FullName -match 'xl/worksheets/' -and $node.LocalName -eq 'hyperlink' -and $node.Attributes['ref']) {
                        $cellRef = $node.Attributes['ref'].Value
                        $cell = $contentXml.SelectSingleNode("//x:c[@r='$cellRef']", $ns)
                        if ($cell) {
                            $valueNode = $cell.SelectSingleNode('x:v', $ns)
                            if ($valueNode -and $cell.Attributes['t'] -and $cell.Attributes['t'].Value -eq 's') {
                                $stringIndex = 0
                                if ([int]::TryParse($valueNode.InnerText, [ref]$stringIndex) -and $stringIndex -lt $sharedStrings.Count) {
                                    $displayText = $sharedStrings[$stringIndex]
                                }
                            }
                            elseif ($cell.SelectSingleNode('x:is', $ns)) {
                                $displayText = [string]$cell.SelectSingleNode('x:is', $ns).InnerText
                            }
                            elseif ($valueNode) {
                                $displayText = [string]$valueNode.InnerText
                            }
                        }
                    }
                    if ([string]::IsNullOrWhiteSpace($displayText) -and $entry.FullName -match 'ppt/slides/' -and $node.LocalName -eq 'hlinkClick') {
                        $shape = $node.SelectSingleNode('ancestor::p:sp[1]', $ns)
                        if ($shape) { $displayText = [string]$shape.InnerText }
                    }
                    $found.Add([PSCustomObject]@{ Text = $displayText.Trim(); Url = [string]$relationsByXml[$entry.FullName][$relId] })
                }
            }
            catch {
                Write-Verbose "Could not parse hyperlinks in $($entry.FullName)."
            }
        }
    }
    finally {
        $archive.Dispose()
    }

    return $found.ToArray()
}

# PDF URI link annotations are the actual clickable links a reader sees/clicks - not a
# raw-text scan, which would also catch incidental URL-shaped text that isn't a link.
function Get-PdfHyperlinks {
    param([Parameter(Mandatory = $true)][AllowEmptyCollection()][byte[]]$Bytes)

    if ($Bytes.Length -eq 0) { return @() }

    $found = New-Object System.Collections.Generic.List[Object]
    $document = [UglyToad.PdfPig.PdfDocument]::Open($Bytes)
    try {
        foreach ($page in $document.GetPages()) {
            foreach ($hyperlink in $page.GetHyperlinks()) {
                $displayText = [string]$hyperlink.Text
                if ([string]::IsNullOrWhiteSpace($displayText) -and $hyperlink.Letters.Count -gt 0) {
                    $displayText = (@($hyperlink.Letters | ForEach-Object { [string]$_.Value }) -join '')
                }
                $found.Add([PSCustomObject]@{ Text = $displayText.Trim(); Url = [string]$hyperlink.Uri })
            }
        }
    }
    finally {
        $document.Dispose()
    }

    return $found.ToArray()
}

function Get-HtmlHyperlinks {
    param([Parameter(Mandatory = $true)][AllowEmptyString()][string]$Html)

    $found = New-Object System.Collections.Generic.List[Object]
    foreach ($anchor in [regex]::Matches($Html, '(?is)<a\b(?<attrs>[^>]*)>(?<body>.*?)</a\s*>')) {
        $href = [regex]::Match($anchor.Groups['attrs'].Value, '(?is)\bhref\s*=\s*(?:"(?<double>[^"]*)"|''(?<single>[^'']*)''|(?<bare>[^\s>]+))')
        if (-not $href.Success) { continue }
        $url = if ($href.Groups['double'].Success) { $href.Groups['double'].Value } elseif ($href.Groups['single'].Success) { $href.Groups['single'].Value } else { $href.Groups['bare'].Value }
        $displayText = [regex]::Replace($anchor.Groups['body'].Value, '(?is)<[^>]+>', ' ')
        $displayText = [System.Net.WebUtility]::HtmlDecode(($displayText -replace '\s+', ' ').Trim())
        $url = [System.Net.WebUtility]::HtmlDecode($url.Trim())
        $found.Add([PSCustomObject]@{ Text = $displayText; Url = $url })
    }
    return $found.ToArray()
}

function Get-CanvasContentHyperlinks {
    param([Parameter(Mandatory = $true)][AllowEmptyString()][string]$CanvasContent)

    $found = New-Object System.Collections.Generic.List[Object]
    if ([string]::IsNullOrWhiteSpace($CanvasContent)) { return @() }

    try {
        $canvas = ConvertFrom-Json -InputObject $CanvasContent -ErrorAction Stop
    }
    catch {
        return @(Get-HtmlHyperlinks -Html $CanvasContent)
    }

    $pending = New-Object System.Collections.Stack
    $pending.Push($canvas)
    while ($pending.Count -gt 0) {
        $current = $pending.Pop()
        if ($null -eq $current -or $current -is [string]) { continue }
        if ($current -is [System.Array]) {
            foreach ($entry in $current) { $pending.Push($entry) }
            continue
        }

        $innerHtml = $current.PSObject.Properties['innerHTML']
        if ($innerHtml -and $innerHtml.Value -is [string]) {
            foreach ($link in @(Get-HtmlHyperlinks -Html ([string]$innerHtml.Value))) { $found.Add($link) }
        }

        foreach ($urlPropertyName in @('linkUrl', 'href')) {
            $urlProperty = $current.PSObject.Properties[$urlPropertyName]
            if (-not $urlProperty -or $urlProperty.Value -isnot [string] -or [string]::IsNullOrWhiteSpace($urlProperty.Value)) { continue }
            $displayText = if ($current.PSObject.Properties['title']) { [string]$current.title } elseif ($current.PSObject.Properties['displayText']) { [string]$current.displayText } else { '' }
            $found.Add([PSCustomObject]@{ Text = $displayText; Url = [string]$urlProperty.Value })
        }

        $sourceItemProperty = $current.PSObject.Properties['sourceItem']
        if ($sourceItemProperty -and $sourceItemProperty.Value) {
            $sourceItem = $sourceItemProperty.Value
            $targetUrl = if ($sourceItem.PSObject.Properties['url']) { [string]$sourceItem.url } else { '' }
            $displayText = if ($current.PSObject.Properties['title']) { [string]$current.title } elseif ($sourceItem.PSObject.Properties['title']) { [string]$sourceItem.title } else { '' }
            if (-not [string]::IsNullOrWhiteSpace($targetUrl)) {
                $found.Add([PSCustomObject]@{ Text = $displayText; Url = $targetUrl })
            }
        }

        foreach ($property in $current.PSObject.Properties) {
            if ($property.Name -eq 'innerHTML' -or $property.Name -eq 'sourceItem') { continue }
            if ($null -ne $property.Value -and $property.Value -isnot [string]) { $pending.Push($property.Value) }
        }
    }

    return $found.ToArray()
}

function Get-FieldValueHyperlinks {
    param(
        [Parameter(Mandatory = $false)]$Value,
        [Parameter(Mandatory = $false)][switch]$CanvasContent
    )

    $found = New-Object System.Collections.Generic.List[Object]
    if ($null -eq $Value) { return @() }

    if ($Value.PSObject.Properties['Url']) {
        $url = [string]$Value.Url
        if (-not [string]::IsNullOrWhiteSpace($url)) {
            $displayText = if ($Value.PSObject.Properties['Description']) { [string]$Value.Description } else { '' }
            $found.Add([PSCustomObject]@{ Text = $displayText; Url = $url })
        }
        return $found.ToArray()
    }

    foreach ($textValue in @($Value)) {
        if ($textValue -isnot [string]) { continue }
        $hyperlinks = if ($CanvasContent) { Get-CanvasContentHyperlinks -CanvasContent $textValue } else { Get-HtmlHyperlinks -Html $textValue }
        foreach ($link in @($hyperlinks)) {
            $found.Add($link)
        }
    }
    return $found.ToArray()
}

function Test-IsInternalTenantUrl {
    param(
        [Parameter(Mandatory = $true)][string]$Url,
        [Parameter(Mandatory = $true)][string]$RootUrl
    )

    try {
        return ([Uri]$Url).Host.Equals(([Uri]$RootUrl).Host, [System.StringComparison]::OrdinalIgnoreCase)
    }
    catch {
        return $false
    }
}

function Test-IsUrlWithinOldSiteBase {
    param(
        [Parameter(Mandatory = $true)][string]$Url,
        [Parameter(Mandatory = $true)][array]$SiteUrlMap
    )

    try {
        $uri = [Uri]$Url
        $urlPath = [Uri]::UnescapeDataString($uri.AbsolutePath).TrimEnd('/')
        $urlPath = $urlPath -replace '(?i)^/:[a-z]:/(?:r|s)(?=/)', ''
        foreach ($entry in $SiteUrlMap) {
            $oldBaseUri = [Uri]$entry.OldBase
            if (-not $uri.Host.Equals($oldBaseUri.Host, [System.StringComparison]::OrdinalIgnoreCase)) { continue }

            $oldPath = $oldBaseUri.AbsolutePath.TrimEnd('/')
            if ([string]::IsNullOrWhiteSpace($oldPath) -and $urlPath -match '(?i)^/sites(?:/|$)') { continue }
            if ($urlPath.Equals($oldPath, [System.StringComparison]::OrdinalIgnoreCase) -or
                $urlPath.StartsWith(($oldPath + '/'), [System.StringComparison]::OrdinalIgnoreCase)) {
                return $true
            }
        }
    }
    catch {
        return $false
    }

    return $false
}

function Get-EmbeddedFileServerRelativePath {
    param([Parameter(Mandatory = $true)][string]$Url)

    try {
        $uri = [Uri]$Url
        $path = [Uri]::UnescapeDataString($uri.AbsolutePath)
        $viewerPathMatch = [regex]::Match($path, '(?i)^/:[a-z]:/r(/.*)$')
        if ($viewerPathMatch.Success) {
            $path = $viewerPathMatch.Groups[1].Value
        }

        if ($path -match '(?i)/Forms/(?:AllItems|DispForm|EditForm)\.aspx$') {
            $idMatch = [regex]::Match($uri.Query, '(?i)(?:^|[?&])id=([^&]+)')
            if ($idMatch.Success) {
                $idPath = [Uri]::UnescapeDataString($idMatch.Groups[1].Value.Replace('+', ' '))
                if ($idPath.StartsWith('/')) { return $idPath.TrimEnd('/') }
            }
            return $null
        }

        if ($path -match '(?i)/_layouts/15/Doc\.aspx$') {
            return $null
        }

        return $path.TrimEnd('/')
    }
    catch {
        return $null
    }
}

function Resolve-EmbeddedTargetUrl {
    param(
        [Parameter(Mandatory = $true)][string]$Url,
        [Parameter(Mandatory = $true)][string]$RootUrl,
        [Parameter(Mandatory = $true)][array]$SiteUrlMap,
        [Parameter(Mandatory = $false)][string]$ClientId,
        [Parameter(Mandatory = $true)][System.Collections.IDictionary]$ConnectionCache
    )

    $mappedUrl = ''
    $mappedSubsite = ''
    $mappedMatch = Get-NewSiteBaseForUrl -Url $Url -SiteUrlMap $SiteUrlMap
    if ($mappedMatch) {
        $mappedUrl = $mappedMatch.NewLinkBase + $mappedMatch.Suffix
        $mappedSubsite = $mappedMatch.NewLinkBase
    }

    $oldFilePath = Get-EmbeddedFileServerRelativePath -Url $Url
    $pathMatch = $null
    if ($oldFilePath) {
        try {
            $embeddedUri = [Uri]$Url
            $pathOnlyUrl = '{0}://{1}{2}' -f $embeddedUri.Scheme, $embeddedUri.Host, $oldFilePath
            $pathMatch = Get-NewSiteBaseForUrl -Url $pathOnlyUrl -SiteUrlMap $SiteUrlMap
        }
        catch {
            $pathMatch = $null
        }
    }

    if ([string]::IsNullOrWhiteSpace($oldFilePath) -or -not $pathMatch) {
        return [PSCustomObject]@{
            Found      = if ($mappedMatch) { 'Mapped' } else { 'No' }
            NewSubsite = $mappedSubsite
            NewUrl     = $mappedUrl
            Resolution = if ($mappedMatch) { 'Mapped base URL only; embedded link does not expose a file path.' } else { 'No matching Old Site URL row.' }
        }
    }

    $locationOverride = Get-NewDocumentLocationOverride -FileRef $oldFilePath
    if ($locationOverride) {
        $relativeWithinOverrideLibrary = $oldFilePath.Substring($locationOverride.OldLibraryPath.Length)
        $newFilePath = $locationOverride.NewLibraryPath + $relativeWithinOverrideLibrary
        $newLinkFilePath = $newFilePath
        $connectionBase = $locationOverride.ConnectionBase.TrimEnd('/')
        $mappedSubsite = $connectionBase
        $mappedUrl = $RootUrl + (Get-EncodedServerRelativePath -Path $newLinkFilePath)
    }
    else {
        $newFilePath = ([Uri]$pathMatch.NewBase).AbsolutePath.TrimEnd('/') + $pathMatch.Suffix
        $newLinkFilePath = ([Uri]$pathMatch.NewLinkBase).AbsolutePath.TrimEnd('/') + $pathMatch.Suffix
        $connectionBase = $pathMatch.NewBase.TrimEnd('/')
    }

    if (-not $ConnectionCache.ContainsKey($connectionBase)) {
        try {
            $ConnectionCache[$connectionBase] = Connect-ValidatedPnP -Url $connectionBase -ClientId $ClientId
        }
        catch {
            $ConnectionCache[$connectionBase] = $null
        }
    }

    $connection = $ConnectionCache[$connectionBase]
    if (-not $connection) {
        return [PSCustomObject]@{
            Found      = 'No'
            NewSubsite = $mappedSubsite
            NewUrl     = $mappedUrl
            Resolution = 'Mapped target, but could not connect to its new site.'
        }
    }

    try {
        $itemResult = Invoke-PnPConnectionOperation -Url $connectionBase -ClientId $ClientId -Connection $connection -Label ('Lookup embedded target: ' + $newFilePath) -MaxAttempts 2 -Operation {
            param($activeConnection)
            Get-PnPFile -Url $newFilePath -AsListItem -Connection $activeConnection
        }
        $ConnectionCache[$connectionBase] = $itemResult.Connection
        $newItem = $itemResult.Result
        $newUniqueId = $newItem.FieldValues['UniqueId']
        $extension = Get-NormalizedExtension -FileName ([System.IO.Path]::GetFileName($oldFilePath))

        return [PSCustomObject]@{
            Found      = 'Yes'
            NewSubsite = $mappedSubsite
            NewUrl     = Get-DocumentBrowserUrl -TenantBase $RootUrl -FileRef $newLinkFilePath -Extension $extension -UniqueId $newUniqueId
            Resolution = if ($locationOverride) { 'Live lookup by document-library location override.' } else { 'Live lookup by mapped embedded file path.' }
        }
    }
    catch {
        return [PSCustomObject]@{
            Found      = 'No'
            NewSubsite = $mappedSubsite
            NewUrl     = $mappedUrl
            Resolution = 'Mapped target path was not found in the new site.'
        }
    }
}

function Get-NavigationNodeTreeItems {
    param([Parameter(Mandatory = $false)][array]$Nodes = @())

    $items = New-Object System.Collections.Generic.List[Object]
    $pending = New-Object System.Collections.Stack
    foreach ($node in $Nodes) { if ($node) { $pending.Push($node) } }
    while ($pending.Count -gt 0) {
        $node = $pending.Pop()
        $items.Add($node)
        $childrenProperty = $node.PSObject.Properties['Children']
        if ($childrenProperty) {
            foreach ($child in @($childrenProperty.Value)) { if ($child) { $pending.Push($child) } }
        }
    }
    return $items.ToArray()
}

function Get-NewDocumentLocationOverride {
    param([Parameter(Mandatory = $true)][string]$FileRef)

    $overrides = @(
        [PSCustomObject]@{
            OldLibraryPath = '/2021 E&M Coding Changes'
            ConnectionBase = 'https://whallc.sharepoint.com/sites/PatientFinancialServices'
            NewLibraryPath = '/sites/PatientFinancialServices/2021 EM Coding Changes'
        }
        [PSCustomObject]@{
            OldLibraryPath = '/2021 EM Coding Changes'
            ConnectionBase = 'https://whallc.sharepoint.com/sites/PatientFinancialServices'
            NewLibraryPath = '/sites/PatientFinancialServices/2021 EM Coding Changes'
        }
        [PSCustomObject]@{
            OldLibraryPath = '/COVID19 Resources'
            ConnectionBase = 'https://whallc.sharepoint.com/sites/ClinicalandPatientServices'
            NewLibraryPath = '/sites/ClinicalandPatientServices/COVID19 Resources'
        }
        [PSCustomObject]@{
            OldLibraryPath = '/COVID19 Workflows'
            ConnectionBase = 'https://whallc.sharepoint.com/sites/ClinicalandPatientServices'
            NewLibraryPath = '/sites/ClinicalandPatientServices/COVID19 Workflows'
        }
    )

    foreach ($override in $overrides) {
        if ($FileRef.Equals($override.OldLibraryPath, [System.StringComparison]::OrdinalIgnoreCase) -or
            $FileRef.StartsWith(($override.OldLibraryPath + '/'), [System.StringComparison]::OrdinalIgnoreCase)) {
            return $override
        }
    }

    return $null
}

function Get-ExistingCsvSelectionPath {
    param([Parameter(Mandatory = $false)][string]$DefaultDirectory = '')

    try {
        Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop

        $openDialog = New-Object System.Windows.Forms.OpenFileDialog
        $openDialog.Title = 'Select Old/New Site URL Lookup CSV File'
        $openDialog.Filter = 'CSV files (*.csv)|*.csv|All files (*.*)|*.*'
        $openDialog.DefaultExt = 'csv'
        $openDialog.CheckFileExists = $true
        $openDialog.Multiselect = $false
        if (-not [string]::IsNullOrWhiteSpace($DefaultDirectory) -and (Test-Path -Path $DefaultDirectory)) {
            $openDialog.InitialDirectory = $DefaultDirectory
        }

        $dialogResult = $openDialog.ShowDialog()
        if ($dialogResult -eq [System.Windows.Forms.DialogResult]::OK -and -not [string]::IsNullOrWhiteSpace($openDialog.FileName)) {
            return [string]$openDialog.FileName
        }
    }
    catch {
        Write-Host '  Could not open CSV selection dialog. You can still enter a path manually.'
    }

    $manualPath = ''
    while ([string]::IsNullOrWhiteSpace($manualPath) -or -not (Test-Path -Path $manualPath)) {
        $manualPath = Read-Host 'Enter full path to the Old/New Site URL Lookup CSV'
        if (-not [string]::IsNullOrWhiteSpace($manualPath) -and -not (Test-Path -Path $manualPath)) {
            Write-Host "  File not found: $manualPath"
        }
    }

    return $manualPath
}

function Get-BooleanPrompt {
    param(
        [Parameter(Mandatory = $true)][string]$Message,
        [Parameter(Mandatory = $false)][bool]$DefaultYes = $false
    )

    $suffix = '[Y/N]'
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

function Get-UrlMappingCsvPreferencePath {
    param([Parameter(Mandatory = $true)][string]$StateRootFolder)

    return (Join-Path -Path $StateRootFolder -ChildPath 'Get-DocumentUrlLookup_UrlMappingCsvPreference.json')
}

function Get-UrlMappingCsvPreference {
    param([Parameter(Mandatory = $true)][string]$Path)

    if (-not (Test-Path -Path $Path)) {
        return [PSCustomObject]@{ LastCompletedCsvPath = '' }
    }

    try {
        $raw = Get-Content -Path $Path -Raw -Encoding UTF8 | ConvertFrom-Json
        return [PSCustomObject]@{
            LastCompletedCsvPath = if ($raw.PSObject.Properties['LastCompletedCsvPath']) { [string]$raw.LastCompletedCsvPath } else { '' }
        }
    }
    catch {
        Write-Verbose "Could not load URL mapping CSV preference file from ${Path}: $($_.Exception.Message)"
        return [PSCustomObject]@{ LastCompletedCsvPath = '' }
    }
}

function Save-UrlMappingCsvPreference {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$CompletedCsvPath
    )

    try {
        [ordered]@{
            LastCompletedCsvPath = $CompletedCsvPath
            UpdatedAt            = (Get-Date).ToString('o')
        } | ConvertTo-Json -Depth 4 | Set-Content -Path $Path -Encoding UTF8 -Force
    }
    catch {
        Write-Verbose "Could not save URL mapping CSV preference file to ${Path}: $($_.Exception.Message)"
    }
}

# Each run writes a fresh, short, timestamped file rather than mutating a prior one in
# place - this just reads the most recent prior file for the same site (if any) so
# -SkipAlreadyConfirmedDocuments has something to check a document's old status against.
function Import-ExistingLookupTable {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$KeyProperty
    )

    $table = [ordered]@{}
    if (-not (Test-Path -Path $Path)) {
        return $table
    }

    try {
        foreach ($row in @(Import-Csv -Path $Path)) {
            $key = ([string]$row.$KeyProperty).Trim().ToLowerInvariant()
            if ([string]::IsNullOrWhiteSpace($key)) { continue }
            $table[$key] = $row
        }
    }
    catch {
        Write-Warning "Could not read existing lookup table at ${Path}: $($_.Exception.Message)"
    }

    return $table
}

function Export-CsvTableWithRetry {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][array]$Rows,
        [Parameter(Mandatory = $false)][int]$MaxAttempts = 5,
        [Parameter(Mandatory = $false)][int]$DelayMilliseconds = 500
    )

    for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
        try {
            $Rows | Export-Csv -Path $Path -NoTypeInformation -Force
            return
        }
        catch {
            $isLocked = ([string]$_.Exception.Message -match '(?i)being used by another process|cannot access the file')
            if ((-not $isLocked) -or ($attempt -eq $MaxAttempts)) {
                throw
            }
            Start-Sleep -Milliseconds ($DelayMilliseconds * $attempt)
        }
    }
}

# Cheap path for a row whose key never existed in the table before - appending is O(1)
# instead of rewriting the whole (potentially large) table for every new file found.
function Add-CsvRowWithRetry {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][pscustomobject]$Row,
        [Parameter(Mandatory = $false)][int]$MaxAttempts = 5,
        [Parameter(Mandatory = $false)][int]$DelayMilliseconds = 500
    )

    for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
        try {
            $Row | Export-Csv -Path $Path -NoTypeInformation -Append
            return
        }
        catch {
            $isLocked = ([string]$_.Exception.Message -match '(?i)being used by another process|cannot access the file')
            if ((-not $isLocked) -or ($attempt -eq $MaxAttempts)) {
                throw
            }
            Start-Sleep -Milliseconds ($DelayMilliseconds * $attempt)
        }
    }
}

function Get-MasterLookupTablePath {
    param(
        [Parameter(Mandatory = $true)][string]$Folder,
        [Parameter(Mandatory = $true)][string]$Identifier
    )

    return Join-Path -Path $Folder -ChildPath ("{0}-master.csv" -f $Identifier)
}

function Get-TemporaryLookupTablePath {
    param(
        [Parameter(Mandatory = $true)][string]$Folder,
        [Parameter(Mandatory = $true)][string]$Identifier,
        [Parameter(Mandatory = $true)][string]$RunStamp
    )

    return Join-Path -Path $Folder -ChildPath ("{0}-{1}-temp.csv" -f $RunStamp, $Identifier)
}

function Get-FixerLookupTablePath {
    param(
        [Parameter(Mandatory = $true)][string]$Folder,
        [Parameter(Mandatory = $true)][string]$Identifier
    )

    return Join-Path -Path $Folder -ChildPath ("{0}-fixer.csv" -f $Identifier)
}

function Copy-CsvFileWithRetry {
    param(
        [Parameter(Mandatory = $true)][string]$SourcePath,
        [Parameter(Mandatory = $true)][string]$DestinationPath,
        [Parameter(Mandatory = $false)][int]$MaxAttempts = 5,
        [Parameter(Mandatory = $false)][int]$DelayMilliseconds = 500
    )

    for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
        try {
            Copy-Item -Path $SourcePath -Destination $DestinationPath -Force
            return
        }
        catch {
            $isLocked = ([string]$_.Exception.Message -match '(?i)being used by another process|cannot access the file')
            if ((-not $isLocked) -or ($attempt -eq $MaxAttempts)) {
                throw
            }
            Start-Sleep -Milliseconds ($DelayMilliseconds * $attempt)
        }
    }
}

function Remove-FileWithRetry {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $false)][int]$MaxAttempts = 5,
        [Parameter(Mandatory = $false)][int]$DelayMilliseconds = 500
    )

    if (-not (Test-Path -Path $Path)) {
        return
    }

    for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
        try {
            Remove-Item -Path $Path -Force
            return
        }
        catch {
            $isLocked = ([string]$_.Exception.Message -match '(?i)being used by another process|cannot access the file')
            if ((-not $isLocked) -or ($attempt -eq $MaxAttempts)) {
                throw
            }
            Start-Sleep -Milliseconds ($DelayMilliseconds * $attempt)
        }
    }
}

function Publish-MasterLookupTable {
    param(
        [Parameter(Mandatory = $true)][string]$TempPath,
        [Parameter(Mandatory = $true)][string]$MasterPath
    )

    Copy-CsvFileWithRetry -SourcePath $TempPath -DestinationPath $MasterPath
    Remove-FileWithRetry -Path $TempPath
}

function Export-FixerLookupTableFromMaster {
    param(
        [Parameter(Mandatory = $true)][string]$MasterPath,
        [Parameter(Mandatory = $true)][string]$FixerPath,
        [Parameter(Mandatory = $true)][array]$RequiredExtensions
    )

    if (-not (Test-Path -Path $MasterPath)) {
        throw "Master lookup table was not found: $MasterPath"
    }

    $masterRows = @(Import-Csv -Path $MasterPath)
    $filteredRows = @($masterRows | Where-Object {
            $extension = ([string]$_.Extension).Trim().ToLowerInvariant()
            ($RequiredExtensions -contains $extension) -and
            (-not [string]::IsNullOrWhiteSpace([string]$_.OldResolvedUrl))
        })

    Remove-FileWithRetry -Path $FixerPath
    if ($filteredRows.Count -gt 0) {
        Export-CsvTableWithRetry -Path $FixerPath -Rows $filteredRows
    }
    elseif ($masterRows.Count -gt 0) {
        $headerLine = (@($masterRows[0].PSObject.Properties.Name) | ForEach-Object { '"' + ($_ -replace '"', '""') + '"' }) -join ','
        Set-Content -Path $FixerPath -Value $headerLine -Encoding UTF8 -Force
    }
    else {
        Set-Content -Path $FixerPath -Value '' -Encoding UTF8 -Force
    }

    return $filteredRows.Count
}

function Publish-EmbeddedLinksLookupTable {
    param(
        [Parameter(Mandatory = $true)][string]$Folder,
        [Parameter(Mandatory = $true)][string]$Identifier,
        [Parameter(Mandatory = $true)][string]$RunStamp,
        [Parameter(Mandatory = $true)][System.Collections.IDictionary]$RowsByKey,
        [Parameter(Mandatory = $true)][System.Collections.IDictionary]$ScannedSourceUrls
    )

    $columns = @('Site', 'Subsite', 'LibraryorListName', 'FileName', 'Extension', 'DocumentorItemURL', 'EmbeddedURLText', 'EmbeddedURL', 'NewFound', 'NewSubsite', 'NewDocumentorItemURL', 'EmbeddedNewFound', 'EmbeddedNewSubsite', 'EmbeddedNewURL', 'EmbeddedResolution')
    $reportPath = Join-Path -Path $Folder -ChildPath ("$Identifier-EmbeddedLinks.csv")
    $scanTempPath = Join-Path -Path $Folder -ChildPath ("$Identifier-EmbeddedLinks-$RunStamp-scan-temp.csv")
    $mergedTempPath = Join-Path -Path $Folder -ChildPath ("$Identifier-EmbeddedLinks-$RunStamp-merged-temp.csv")

    $currentRows = @($RowsByKey.Values)
    if ($currentRows.Count -gt 0) {
        Export-CsvTableWithRetry -Path $scanTempPath -Rows $currentRows
    }
    else {
        $headerLine = (@($columns | ForEach-Object { '"' + $_ + '"' }) -join ',')
        Set-Content -Path $scanTempPath -Value $headerLine -Encoding UTF8 -Force
    }

    $mergedRowsByKey = [ordered]@{}
    if (Test-Path -Path $reportPath) {
        foreach ($row in @(Import-Csv -Path $reportPath)) {
            $sourceUrlKey = ([string]$row.DocumentorItemURL).Trim().ToLowerInvariant()
            $sourceKey = $sourceUrlKey + "`n" + [string]$row.LibraryorListName
            if ($ScannedSourceUrls.ContainsKey($sourceKey)) { continue }
            $rowKey = Get-Sha256Hex -Text ($sourceKey + "`n" + [string]$row.EmbeddedURLText + "`n" + [string]$row.EmbeddedURL)
            $mergedRowsByKey[$rowKey] = $row
        }
    }

    foreach ($row in @(Import-Csv -Path $scanTempPath)) {
        $sourceUrlKey = ([string]$row.DocumentorItemURL).Trim().ToLowerInvariant()
        $sourceKey = $sourceUrlKey + "`n" + [string]$row.LibraryorListName
        $rowKey = Get-Sha256Hex -Text ($sourceKey + "`n" + [string]$row.EmbeddedURLText + "`n" + [string]$row.EmbeddedURL)
        $mergedRowsByKey[$rowKey] = $row
    }

    $mergedRows = New-Object System.Collections.Generic.List[Object]
    foreach ($row in $mergedRowsByKey.Values) { $mergedRows.Add($row) }
    if ($mergedRows.Count -gt 0) {
        Export-CsvTableWithRetry -Path $mergedTempPath -Rows $mergedRows.ToArray()
    }
    else {
        $headerLine = (@($columns | ForEach-Object { '"' + $_ + '"' }) -join ',')
        Set-Content -Path $mergedTempPath -Value $headerLine -Encoding UTF8 -Force
    }

    Copy-CsvFileWithRetry -SourcePath $mergedTempPath -DestinationPath $reportPath
    Remove-FileWithRetry -Path $scanTempPath
    Remove-FileWithRetry -Path $mergedTempPath

    $staleTimestampedReports = Get-ChildItem -Path $Folder -Filter ("*-{0}-EmbeddedLinks.csv" -f $Identifier) -File -ErrorAction SilentlyContinue |
        Where-Object { $_.FullName -ne $reportPath }
    foreach ($staleReport in $staleTimestampedReports) {
        Remove-FileWithRetry -Path $staleReport.FullName
    }

    return $mergedRows.Count
}

function Get-CounterCheckpointPath {
    param(
        [Parameter(Mandatory = $true)][string]$StateRoot,
        [Parameter(Mandatory = $true)][string]$ScopeUrl,
        [Parameter(Mandatory = $true)][string]$SchemaVersion
    )

    $fingerprint = (Get-Sha256Hex -Text ($ScopeUrl.TrimEnd('/').ToLowerInvariant() + '|' + $SchemaVersion)).Substring(0, 16)
    return Join-Path -Path $StateRoot -ChildPath ("Get-DocumentUrlLookup_CounterCheckpoint_{0}.json" -f $fingerprint)
}

# Snapshot of running totals, saved after every library, so a crash still
# leaves a readable record of how far the scan got (the CSV table itself is
# already rewritten up through the last completed library via Export-CsvTableWithRetry).
function Save-CounterCheckpoint {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$OutputPath,
        [Parameter(Mandatory = $true)][int]$FilesScanned,
        [Parameter(Mandatory = $true)][int]$NewFoundCount,
        [Parameter(Mandatory = $true)][int]$NewNotFoundCount,
        [Parameter(Mandatory = $true)][string]$LastLibraryCompleted
    )

    try {
        [PSCustomObject][ordered]@{
            SavedAt              = (Get-Date).ToString('o')
            OutputPath           = $OutputPath
            FilesScanned         = $FilesScanned
            NewFoundCount        = $NewFoundCount
            NewNotFoundCount     = $NewNotFoundCount
            LastLibraryCompleted = $LastLibraryCompleted
        } | ConvertTo-Json -Depth 4 | Set-Content -Path $Path -Encoding UTF8 -Force
    }
    catch {
        Write-Verbose "Could not save counter checkpoint to ${Path}: $($_.Exception.Message)"
    }
}

# --- Setup --------------------------------------------------------------

Set-ScriptWindowTitle -ScriptName 'Get-DocumentUrlLookup.ps1'
Install-RequiredModule -Name "PnP.PowerShell"

if ([string]::IsNullOrWhiteSpace($MasterFolder)) {
    $MasterFolder = Join-Path -Path $OutputFolder -ChildPath 'Masters'
}
if ([string]::IsNullOrWhiteSpace($FixerOutputFolder)) {
    $FixerOutputFolder = Join-Path -Path $OutputFolder -ChildPath 'FixerFiles'
}
if ([string]::IsNullOrWhiteSpace($SitesMasterFolder)) {
    $SitesMasterFolder = Join-Path -Path $OutputFolder -ChildPath 'Sites-Masters'
}
if ([string]::IsNullOrWhiteSpace($SitesTempMasterFolder)) {
    $SitesTempMasterFolder = Join-Path -Path $SitesMasterFolder -ChildPath 'Temp'
}
if ([string]::IsNullOrWhiteSpace($SitesFixerOutputFolder)) {
    $SitesFixerOutputFolder = Join-Path -Path $OutputFolder -ChildPath 'Sites-FixerFiles'
}
if ([string]::IsNullOrWhiteSpace($SitesEmbeddedLinksFolder)) {
    $SitesEmbeddedLinksFolder = Join-Path -Path $OutputFolder -ChildPath 'Sites-EmbeddedLinks'
}

New-DirectoryIfMissing -Path $OutputFolder
New-DirectoryIfMissing -Path $MasterFolder
New-DirectoryIfMissing -Path $TempMasterFolder
New-DirectoryIfMissing -Path $FixerOutputFolder
New-DirectoryIfMissing -Path $EmbeddedLinksFolder
New-DirectoryIfMissing -Path $SitesMasterFolder
New-DirectoryIfMissing -Path $SitesTempMasterFolder
New-DirectoryIfMissing -Path $SitesFixerOutputFolder
New-DirectoryIfMissing -Path $SitesEmbeddedLinksFolder
New-DirectoryIfMissing -Path $StateRootFolder

$RootUrl = $RootUrl.TrimEnd('/')
$AdminUrl = $AdminUrl.TrimEnd('/')

$adminConn = Connect-ValidatedPnP -Url $AdminUrl -ClientId $ClientId -AdminConnection

Write-Host "`nCollecting tenant sites (excluding OneDrive personal sites)..."
$tenantSiteResult = Invoke-PnPConnectionOperation -Url $AdminUrl -ClientId $ClientId -Connection $adminConn -AdminConnection -Label 'Get tenant sites' -MaxAttempts 4 -Operation {
    param($activeConnection)
    Get-PnPTenantSite -Connection $activeConnection -Detailed
}
$adminConn = $tenantSiteResult.Connection
$tenantSites = $tenantSiteResult.Result |
Where-Object { $_.Url -and $_.Url -notmatch '-my\.sharepoint\.com' -and $_.Template -notlike 'SPSPERS*' } |
Sort-Object Url -Unique

if ($tenantSites.Count -eq 0) {
    throw "No eligible SharePoint site collections found (excluding OneDrive)."
}

$selectedScope = Select-SiteForScope -Sites $tenantSites
$selectedScopeUrl = $selectedScope.Url.TrimEnd('/')
Write-Host "`nSelected inventory scope: $selectedScopeUrl"

$scopeConn = Connect-ValidatedPnP -Url $selectedScopeUrl -ClientId $ClientId

Write-Host ("`nSubsite inventory under selected scope {0}:" -f $selectedScopeUrl)
$scopeRootResult = Invoke-PnPConnectionOperation -Url $selectedScopeUrl -ClientId $ClientId -Connection $scopeConn -Label 'Get selected scope root web' -MaxAttempts 4 -Operation {
    param($activeConnection)
    Get-PnPWeb -Connection $activeConnection -Includes Url, Title
}
$scopeConn = $scopeRootResult.Connection
$scopeRootWeb = $scopeRootResult.Result

$scopeSubsites = New-Object System.Collections.Generic.List[Object]
$scopeSubsites.Add([PSCustomObject]@{ Url = $scopeRootWeb.Url; Title = $scopeRootWeb.Title })

$scopeChildResult = Invoke-PnPConnectionOperation -Url $selectedScopeUrl -ClientId $ClientId -Connection $scopeConn -Label 'Get selected scope child webs' -MaxAttempts 4 -Operation {
    param($activeConnection)
    Get-PnPSubWeb -Connection $activeConnection -Includes Url, Title
}
$scopeConn = $scopeChildResult.Connection
foreach ($child in @($scopeChildResult.Result | Sort-Object Url)) {
    $scopeSubsites.Add([PSCustomObject]@{ Url = $child.Url; Title = $child.Title })
}

$rootHostForInventory = ([Uri]$RootUrl).Host.ToLowerInvariant()
$scannableScopeSubsites = New-Object System.Collections.Generic.List[Object]
$excludedScopeSubsites = New-Object System.Collections.Generic.List[Object]

$inventoryCachePath = ''
$cachedInventoryByKey = @{}
$freshInventoryByKey = @{}
$inventoryEntriesForSave = New-Object System.Collections.Generic.List[Object]
$inventoryCacheHits = 0
$inventoryDeltaProbes = 0

if ($UsePersistedInventoryCache) {
    $inventoryCachePath = Get-InventorySubsiteCachePath -StateRoot $StateRootFolder -ScopeUrl $selectedScopeUrl -SchemaVersion $ScriptSchemaVersion
    $cachedInventoryByKey = Get-InventorySubsiteCache -Path $inventoryCachePath
}

foreach ($scopeSub in @($scopeSubsites | Sort-Object Url -Unique)) {
    $scopeSubUrl = ([string]$scopeSub.Url).TrimEnd('/')
    $scopeSubKey = $scopeSubUrl.ToLowerInvariant()
    $freshInventoryByKey[$scopeSubKey] = $true

    $isScannable = $false
    $reason = ''

    if ($UsePersistedInventoryCache -and $cachedInventoryByKey.ContainsKey($scopeSubKey)) {
        $cached = $cachedInventoryByKey[$scopeSubKey]
        $isScannable = [bool]$cached.IsScannable
        $reason = [string]$cached.Reason
        $inventoryCacheHits++
    }
    else {
        $inventoryDeltaProbes++
        $probe = Test-IsInventorySubsiteScannable -WebUrl $scopeSubUrl -RootHost $rootHostForInventory -ClientId $ClientId
        $isScannable = [bool]$probe.IsScannable
        $reason = [string]$probe.Reason
    }

    $inventoryEntriesForSave.Add([PSCustomObject]@{ Url = $scopeSubUrl; Title = [string]$scopeSub.Title; IsScannable = $isScannable; Reason = $reason })

    if ($isScannable) {
        $scannableScopeSubsites.Add($scopeSub)
    }
    else {
        $excludedScopeSubsites.Add([PSCustomObject]@{ Url = $scopeSubUrl; Reason = $reason })
    }
}

if ($UsePersistedInventoryCache) {
    Save-InventorySubsiteCache -Path $inventoryCachePath -Entries $inventoryEntriesForSave.ToArray() -ScopeUrl $selectedScopeUrl -SchemaVersion $ScriptSchemaVersion
    Write-Host ("Inventory cache summary: reused {0}, delta-probed {1}." -f $inventoryCacheHits, $inventoryDeltaProbes)
}

if ($excludedScopeSubsites.Count -gt 0) {
    Write-Host ("Filtered out {0} site/web(s) that are not scannable by this script:" -f $excludedScopeSubsites.Count)
    foreach ($excluded in $excludedScopeSubsites) {
        Write-Host ("  - {0}  ({1})" -f $excluded.Url, $excluded.Reason)
    }
}

$scopeSubsites = $scannableScopeSubsites
if (-not $scopeSubsites -or $scopeSubsites.Count -eq 0) {
    throw "No scannable sites/webs were found under $selectedScopeUrl after inventory filtering."
}

for ($i = 0; $i -lt $scopeSubsites.Count; $i++) {
    $title = if ($scopeSubsites[$i].Title) { $scopeSubsites[$i].Title } else { '(no title)' }
    Write-Host ("  {0,4}. {1}  |  {2}" -f ($i + 1), $scopeSubsites[$i].Url, $title)
}
Write-Host "     0. ALL listed sites/webs"

$selectedIndices = $null
while ($null -eq $selectedIndices) {
    $pickRaw = Read-Host "`nPick one or more sites/webs by number (comma-separated, e.g. 1,4,7). Use 0 for ALL listed sites/webs"
    if ([string]::IsNullOrWhiteSpace($pickRaw)) {
        Write-Host "  Selection is required."
        continue
    }

    $parts = @($pickRaw -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    if ($parts -contains '0') {
        $selectedIndices = @(0..($scopeSubsites.Count - 1))
        continue
    }

    $indices = New-Object System.Collections.Generic.List[int]
    $allValid = $true
    foreach ($part in $parts) {
        if ($part -match '^\d+$' -and [int]$part -ge 1 -and [int]$part -le $scopeSubsites.Count) {
            $indices.Add([int]$part - 1)
        }
        else {
            $allValid = $false
        }
    }

    if ($allValid -and $indices.Count -gt 0) {
        $selectedIndices = @($indices | Sort-Object -Unique)
    }
    else {
        Write-Host "  Invalid selection. Use numbers from the list, comma-separated, or 0 for all."
    }
}

$selectedSubsites = @($selectedIndices | ForEach-Object { $scopeSubsites[$_] })

# --- URL mapping CSV, validated against the exact subsites just chosen ----
# (loaded here, after selection, so coverage gaps surface before any
# scanning starts instead of crashing partway through a run)

$urlMappingCsvPreferencePath = Get-UrlMappingCsvPreferencePath -StateRootFolder $StateRootFolder
$urlMappingCsvPreference = Get-UrlMappingCsvPreference -Path $urlMappingCsvPreferencePath
$firstAttemptPath = $UrlMappingCsvPath
$defaultUrlMappingCsvPath = "C:\Users\!bposey\OneDrive - Womens Healthcare Associates LLC\Gary Henderson's files - Project Files_bposey\Support-Files\OldSiteURL-NewSiteURL-Lookup.csv"

$siteUrlMap = $null
while (-not $siteUrlMap) {
    $csvPath = ''
    if (-not [string]::IsNullOrWhiteSpace($firstAttemptPath) -and (Test-Path -Path $firstAttemptPath)) {
        $csvPath = $firstAttemptPath
    }
    $firstAttemptPath = ''

    if ([string]::IsNullOrWhiteSpace($csvPath)) {
        $cachedCsvPath = [string]$urlMappingCsvPreference.LastCompletedCsvPath
        if (-not [string]::IsNullOrWhiteSpace($cachedCsvPath) -and (Test-Path -Path $cachedCsvPath)) {
            $reuse = Get-BooleanPrompt -Message ("Use last completed Old/New Site URL CSV? {0}" -f $cachedCsvPath) -DefaultYes $true
            if ($reuse) { $csvPath = $cachedCsvPath }
        }
    }

    if ([string]::IsNullOrWhiteSpace($csvPath)) {
        $csvPath = Get-ExistingCsvSelectionPath -DefaultDirectory (Split-Path -Path $defaultUrlMappingCsvPath -Parent)
    }

    if ([string]::IsNullOrWhiteSpace($csvPath)) {
        $csvPath = Read-Host 'Enter full path to the Old/New Site URL Lookup CSV'
    }

    if ([string]::IsNullOrWhiteSpace($csvPath) -or -not (Test-Path -Path $csvPath)) {
        Write-Host "  File not found: $csvPath"
        continue
    }

    try {
        $candidateMap = Get-SiteUrlMapFromCsv -CsvPath $csvPath
    }
    catch {
        Write-Warning ("Could not load CSV: {0}" -f $_.Exception.Message)
        continue
    }

    $missing = New-Object System.Collections.Generic.List[string]
    foreach ($sub in $selectedSubsites) {
        $subUrl = ([string]$sub.Url).TrimEnd('/')
        $isSitesWeb = ([Uri]$subUrl).AbsolutePath -match '(?i)^/sites(?:/|$)'
        if ((-not $isSitesWeb) -and -not (Get-ExactNewSiteBaseForUrl -Url $subUrl -SiteUrlMap $candidateMap)) {
            $missing.Add($subUrl)
        }
    }

    if ($missing.Count -gt 0) {
        Write-Warning ("This CSV has no row covering {0} of the {1} selected subsite(s):" -f $missing.Count, $selectedSubsites.Count)
        foreach ($m in $missing) { Write-Host "  - $m" }
        Write-Host "  (Old URLs will still be recorded for these; new URLs will be blank unless you pick a different CSV.)"
        $proceedAnyway = Get-BooleanPrompt -Message 'Continue with this CSV anyway?' -DefaultYes $false
        if (-not $proceedAnyway) {
            continue
        }
    }

    $siteUrlMap = $candidateMap
    Save-UrlMappingCsvPreference -Path $urlMappingCsvPreferencePath -CompletedCsvPath $csvPath
    Write-Host ("Loaded {0} site row(s) from {1}" -f $siteUrlMap.Count, $csvPath)
}

$automaticEmbeddedScanSites = @{}
foreach ($selectedSubsite in $selectedSubsites) {
    $selectedSubsiteUrl = ([string]$selectedSubsite.Url).TrimEnd('/')
    $isSitesWeb = ([Uri]$selectedSubsiteUrl).AbsolutePath -match '(?i)^/sites(?:/|$)'
    if ($isSitesWeb -or -not (Get-ExactNewSiteBaseForUrl -Url $selectedSubsiteUrl -SiteUrlMap $siteUrlMap)) {
        $automaticEmbeddedScanSites[$selectedSubsiteUrl.ToLowerInvariant()] = $true
    }
}

$scanDocumentContent = $false
if ($PSBoundParameters.ContainsKey('EnableDocumentContentScan')) {
    $scanDocumentContent = [bool]$EnableDocumentContentScan
}
else {
    $scanDocumentContent = Get-BooleanPrompt -Message "Also scan document, PDF, page, and list text/URL fields for embedded whallc.sharepoint.com links (downloads supported file content and writes per-site EmbeddedLinks reports)?" -DefaultYes $false
}
if ($automaticEmbeddedScanSites.Count -gt 0) {
    Write-Host ("Automatically enabling legacy embedded-link scanning for {0} selected site(s); only links under CSV Old Site URL bases will be reported." -f $automaticEmbeddedScanSites.Count)
}
if ($scanDocumentContent -or $automaticEmbeddedScanSites.Count -gt 0) {
    Initialize-PdfPig
}

# --- Enumeration ----------------------------------------------------------
# One master CSV per NEW/destination site (not per selected scope) - a classic
# scope like the tenant root can contain dozens of subsites that are each being
# promoted to their own separate new site collection. Each run writes a temp
# file, then publishes it over the matching [site]-master.csv after the scan.

function Get-ShortSiteIdentifier {
    param(
        [Parameter(Mandatory = $true)]$Subsite,
        [Parameter(Mandatory = $false)]$SiteMatch
    )

    $scannedUrl = ([string]$Subsite.Url).TrimEnd('/')
    $scannedPath = ([Uri]$scannedUrl).AbsolutePath.TrimEnd('/')
    $lastSegment = ($scannedPath -split '/')[-1]
    if ([string]::IsNullOrWhiteSpace($lastSegment)) {
        $lastSegment = [string]$Subsite.Title
    }

    $identifier = Get-SafeFileSlug -Text $lastSegment -MaxLength 30 -Fallback 'Site'
    return $identifier
}

function Repair-SiteCsvOutput {
    param(
        [Parameter(Mandatory = $true)][string]$CanonicalPath,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][string[]]$LegacyPaths
    )

    $existingFiles = @(@($CanonicalPath) + $LegacyPaths | Where-Object { Test-Path -LiteralPath $_ } | ForEach-Object { Get-Item -LiteralPath $_ })
    if ($existingFiles.Count -eq 0) { return }

    $newestFile = @($existingFiles | Sort-Object LastWriteTimeUtc -Descending)[0]
    if ($newestFile.FullName -ne $CanonicalPath) {
        Write-Host ("  Replacing canonical output with newer legacy file: {0} -> {1}" -f $newestFile.FullName, $CanonicalPath)
        Copy-CsvFileWithRetry -SourcePath $newestFile.FullName -DestinationPath $CanonicalPath
    }

    foreach ($legacyPath in $LegacyPaths) {
        Remove-FileWithRetry -Path $legacyPath
    }
}

function Repair-SiteOutputInventory {
    param(
        [Parameter(Mandatory = $true)][string]$Identifier,
        [Parameter(Mandatory = $true)][string]$MasterFolder,
        [Parameter(Mandatory = $true)][string]$FixerFolder,
        [Parameter(Mandatory = $true)][string]$EmbeddedFolder,
        [Parameter(Mandatory = $true)][string]$TempFolder,
        [Parameter(Mandatory = $false)][bool]$IncludeSitesPrefixLegacy = $false,
        [Parameter(Mandatory = $false)][string]$LegacyMasterFolder = '',
        [Parameter(Mandatory = $false)][string]$LegacyFixerFolder = '',
        [Parameter(Mandatory = $false)][string]$LegacyEmbeddedFolder = ''
    )

    $legacyMasterPaths = @()
    $legacyFixerPaths = @()
    $legacyEmbeddedPaths = @()
    if ($IncludeSitesPrefixLegacy) {
        $legacyMasterPaths += Join-Path $MasterFolder "sites-$Identifier-master.csv"
        $legacyFixerPaths += Join-Path $FixerFolder "sites-$Identifier-fixer.csv"
        $legacyEmbeddedPaths += Join-Path $EmbeddedFolder "sites-$Identifier-EmbeddedLinks.csv"
        if (-not [string]::IsNullOrWhiteSpace($LegacyMasterFolder)) { $legacyMasterPaths += Join-Path $LegacyMasterFolder "sites-$Identifier-master.csv" }
        if (-not [string]::IsNullOrWhiteSpace($LegacyFixerFolder)) { $legacyFixerPaths += Join-Path $LegacyFixerFolder "sites-$Identifier-fixer.csv" }
        if (-not [string]::IsNullOrWhiteSpace($LegacyEmbeddedFolder)) { $legacyEmbeddedPaths += Join-Path $LegacyEmbeddedFolder "sites-$Identifier-EmbeddedLinks.csv" }
    }

    Repair-SiteCsvOutput -CanonicalPath (Join-Path $MasterFolder "$Identifier-master.csv") -LegacyPaths $legacyMasterPaths
    Repair-SiteCsvOutput -CanonicalPath (Join-Path $FixerFolder "$Identifier-fixer.csv") -LegacyPaths $legacyFixerPaths
    Repair-SiteCsvOutput -CanonicalPath (Join-Path $EmbeddedFolder "$Identifier-EmbeddedLinks.csv") -LegacyPaths $legacyEmbeddedPaths

    foreach ($tempFile in @(Get-ChildItem -Path $TempFolder -File -ErrorAction SilentlyContinue | Where-Object {
                $_.Name -match ("(?i)^\d{8}-\d{6}-(?:sites-)?{0}-temp\.csv$" -f [regex]::Escape($Identifier))
            })) {
        Remove-FileWithRetry -Path $tempFile.FullName
    }
    foreach ($tempFile in @(Get-ChildItem -Path $EmbeddedFolder -File -ErrorAction SilentlyContinue | Where-Object {
                $_.Name -match ("(?i)^(?:sites-)?{0}-EmbeddedLinks-\d{{8}}-\d{{6}}-(?:scan|merged)-temp\.csv$" -f [regex]::Escape($Identifier))
            })) {
        Remove-FileWithRetry -Path $tempFile.FullName
    }
}

function Assert-SiteOutputInventory {
    param(
        [Parameter(Mandatory = $true)][string]$Identifier,
        [Parameter(Mandatory = $true)][string]$MasterFolder,
        [Parameter(Mandatory = $true)][string]$FixerFolder,
        [Parameter(Mandatory = $true)][string]$EmbeddedFolder,
        [Parameter(Mandatory = $false)][bool]$CheckSitesPrefixLegacy = $false,
        [Parameter(Mandatory = $false)][string]$LegacyMasterFolder = '',
        [Parameter(Mandatory = $false)][string]$LegacyFixerFolder = '',
        [Parameter(Mandatory = $false)][string]$LegacyEmbeddedFolder = ''
    )

    $expectedPaths = @(
        (Join-Path $MasterFolder "$Identifier-master.csv"),
        (Join-Path $FixerFolder "$Identifier-fixer.csv"),
        (Join-Path $EmbeddedFolder "$Identifier-EmbeddedLinks.csv")
    )
    $legacyPaths = @()
    if ($CheckSitesPrefixLegacy) {
        $legacyPaths += Join-Path $MasterFolder "sites-$Identifier-master.csv"
        $legacyPaths += Join-Path $FixerFolder "sites-$Identifier-fixer.csv"
        $legacyPaths += Join-Path $EmbeddedFolder "sites-$Identifier-EmbeddedLinks.csv"
        if (-not [string]::IsNullOrWhiteSpace($LegacyMasterFolder)) { $legacyPaths += Join-Path $LegacyMasterFolder "sites-$Identifier-master.csv" }
        if (-not [string]::IsNullOrWhiteSpace($LegacyFixerFolder)) { $legacyPaths += Join-Path $LegacyFixerFolder "sites-$Identifier-fixer.csv" }
        if (-not [string]::IsNullOrWhiteSpace($LegacyEmbeddedFolder)) { $legacyPaths += Join-Path $LegacyEmbeddedFolder "sites-$Identifier-EmbeddedLinks.csv" }
    }
    $missingPaths = @($expectedPaths | Where-Object { -not (Test-Path -LiteralPath $_) })
    $remainingLegacyPaths = @($legacyPaths | Where-Object { Test-Path -LiteralPath $_ })
    if ($missingPaths.Count -gt 0 -or $remainingLegacyPaths.Count -gt 0) {
        throw ("Output inventory check failed for {0}. Missing: [{1}]. Legacy duplicates: [{2}]." -f $Identifier, ($missingPaths -join '; '), ($remainingLegacyPaths -join '; '))
    }

    Write-Host ("  Output inventory verified: {0}-master.csv, {0}-fixer.csv, {0}-EmbeddedLinks.csv" -f $Identifier)
}

$runStamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$newConnectionsByBase = @{}
$embeddedTargetResolutionCache = @{}
$filesScannedTotal = 0
$newFoundCountTotal = 0
$newFoundFromCacheCountTotal = 0
$newNotFoundCountTotal = 0
$extensionCountsTotal = @{}
$publishedSiteIdentifiers = [ordered]@{}

# Embedded-link findings are keyed by source item and link details; scanned-source keys
# allow the publisher to refresh scanned items while retaining cached unscanned items.
$embeddedLinksBySite = @{}
$scannedEmbeddedSourcesBySite = @{}
$ContentScanExtensions = @('.docx', '.xlsx', '.pptx', '.pdf', '.htm', '.html')

foreach ($subsite in $selectedSubsites) {
    $subsiteUrl = ([string]$subsite.Url).TrimEnd('/')
    Write-Host "`nScanning $subsiteUrl ..."

    $siteMatch = Get-ExactNewSiteBaseForUrl -Url $subsiteUrl -SiteUrlMap $siteUrlMap
    $isSitesWeb = ([Uri]$subsiteUrl).AbsolutePath -match '(?i)^/sites(?:/|$)'
    $scanThisSiteContent = $scanDocumentContent -or $automaticEmbeddedScanSites.ContainsKey($subsiteUrl.ToLowerInvariant())
    $scanOldBaseOnly = $isSitesWeb -or ((-not $siteMatch) -and (-not $scanDocumentContent))
    if ((-not $siteMatch) -and (-not $isSitesWeb)) {
        Write-Warning "  No matching row in the Old/New Site URL CSV for $subsiteUrl - old URLs will still be recorded, but new URLs cannot be looked up."
    }
    elseif ($isSitesWeb) {
        Write-Host "  Modern /sites web: scanning content, items, and navigation for links under mapped Old Site URL bases."
    }

    $siteIdentifier = Get-ShortSiteIdentifier -Subsite $subsite -SiteMatch $siteMatch
    $currentMasterFolder = if ($isSitesWeb) { $SitesMasterFolder } else { $MasterFolder }
    $currentTempMasterFolder = if ($isSitesWeb) { $SitesTempMasterFolder } else { $TempMasterFolder }
    $currentFixerFolder = if ($isSitesWeb) { $SitesFixerOutputFolder } else { $FixerOutputFolder }
    $currentEmbeddedFolder = if ($isSitesWeb) { $SitesEmbeddedLinksFolder } else { $EmbeddedLinksFolder }
    $outputNamespace = if ($isSitesWeb) { 'sites' } else { 'classic' }
    $siteOutputKey = "$outputNamespace|$siteIdentifier"
    if ($isSitesWeb) {
        Repair-SiteOutputInventory -Identifier $siteIdentifier -MasterFolder $currentMasterFolder -FixerFolder $currentFixerFolder -EmbeddedFolder $currentEmbeddedFolder -TempFolder $currentTempMasterFolder -IncludeSitesPrefixLegacy $true -LegacyMasterFolder $MasterFolder -LegacyFixerFolder $FixerOutputFolder -LegacyEmbeddedFolder $EmbeddedLinksFolder
    }
    else {
        Repair-SiteOutputInventory -Identifier $siteIdentifier -MasterFolder $currentMasterFolder -FixerFolder $currentFixerFolder -EmbeddedFolder $currentEmbeddedFolder -TempFolder $currentTempMasterFolder
    }
    $masterPath = Get-MasterLookupTablePath -Folder $currentMasterFolder -Identifier $siteIdentifier
    $outputPath = Get-TemporaryLookupTablePath -Folder $currentTempMasterFolder -Identifier $siteIdentifier -RunStamp $runStamp
    $fixerPath = Get-FixerLookupTablePath -Folder $currentFixerFolder -Identifier $siteIdentifier
    $mergedRowsByKey = if (Test-Path -Path $masterPath) { Import-ExistingLookupTable -Path $masterPath -KeyProperty 'OldFileRef' } else { [ordered]@{} }
    $counterCheckpointPath = Get-CounterCheckpointPath -StateRoot $StateRootFolder -ScopeUrl $subsiteUrl -SchemaVersion $ScriptSchemaVersion
    if (Test-Path -Path $masterPath) {
        Write-Host ("  Writing temp scan {0} (using {1} existing row(s) from master {2} to skip already-confirmed lookups)" -f $outputPath, $mergedRowsByKey.Count, $masterPath)
    }
    else {
        Write-Host ("  Writing temp scan {0} (no master file found yet for this site; one will be created at {1})" -f $outputPath, $masterPath)
    }

    if (-not $embeddedLinksBySite.ContainsKey($siteOutputKey)) { $embeddedLinksBySite[$siteOutputKey] = [ordered]@{} }
    if (-not $scannedEmbeddedSourcesBySite.ContainsKey($siteOutputKey)) { $scannedEmbeddedSourcesBySite[$siteOutputKey] = @{} }

    $filesScanned = 0
    $newFoundCount = 0
    $newFoundFromCacheCount = 0
    $newNotFoundCount = 0
    $notFoundCountThisSite = 0
    $extensionCounts = @{}

    $oldSiteServerRelative = ([Uri]$subsiteUrl).AbsolutePath.TrimEnd('/')
    $newSiteScanBaseAbsolute = if ($siteMatch) { ($siteMatch.NewBase + $siteMatch.Suffix).TrimEnd('/') } else { '' }
    $newSiteScanServerRelative = if ($newSiteScanBaseAbsolute) { ([Uri]$newSiteScanBaseAbsolute).AbsolutePath.TrimEnd('/') } else { '' }
    $newSiteLinkBaseAbsolute = if ($siteMatch) { ($siteMatch.NewLinkBase + $siteMatch.Suffix).TrimEnd('/') } else { '' }
    $newSiteLinkServerRelative = if ($newSiteLinkBaseAbsolute) { ([Uri]$newSiteLinkBaseAbsolute).AbsolutePath.TrimEnd('/') } else { '' }

    $subConn = $null
    try {
        $subConn = Connect-ValidatedPnP -Url $subsiteUrl -ClientId $ClientId
    }
    catch {
        Write-Warning "  Could not connect to $subsiteUrl - skipping: $($_.Exception.Message)"
        continue
    }

    $listResult = Invoke-PnPConnectionOperation -Url $subsiteUrl -ClientId $ClientId -Connection $subConn -Label ('Get lists: ' + $subsiteUrl) -MaxAttempts 3 -Operation {
        param($activeConnection)
        Get-PnPList -Connection $activeConnection -Includes BaseTemplate, Hidden, Title, ItemCount, RootFolder.ServerRelativeUrl
    }
    $subConn = $listResult.Connection
    $libraries = @($listResult.Result | Where-Object {
            Test-IsSiteContentsInventoryList -List $_
        })

    $masterColumns = @('Site', 'Subsite', 'Library', 'FileName', 'Extension', 'OldFileRef', 'OldCopyLinkUrl', 'OldResolvedUrl', 'NewFileRef', 'NewFound', 'NewCopyLinkUrl', 'NewResolvedUrl', 'Modified', 'Notes')
    $masterHeaderLine = (@($masterColumns | ForEach-Object { '"' + $_ + '"' }) -join ',')
    Set-Content -Path $outputPath -Value $masterHeaderLine -Encoding UTF8 -Force

    foreach ($library in $libraries) {
        $librarySupportsFileLookup = Test-IsFileLookupList -List $library
        $libraryKind = if ($librarySupportsFileLookup) { 'file/page library' } else { 'site contents list' }
        Write-Host ("  {0}: {1} ({2} items)" -f $libraryKind, $library.Title, $library.ItemCount)

        $contentFieldNames = @()
        $contentFieldScanSucceeded = $false
        if ($scanThisSiteContent) {
            try {
                $fieldResult = Invoke-PnPConnectionOperation -Url $subsiteUrl -ClientId $ClientId -Connection $subConn -Label ('Get content fields: ' + $library.Title) -MaxAttempts 2 -Operation {
                    param($activeConnection)
                    Get-PnPField -List $library.Title -Connection $activeConnection
                }
                $subConn = $fieldResult.Connection
                $contentFieldNames = @($fieldResult.Result | Where-Object {
                    $pageContentField = ($library.Title -eq 'Site Pages' -and $_.InternalName -in @('CanvasContent1', 'WikiField', 'PublishingPageContent', 'LayoutWebpartsContent'))
                    (($pageContentField) -or ($_.TypeAsString -in @('Text', 'Note', 'URL') -and -not $_.Hidden -and -not $_.ReadOnlyField))
                    } | ForEach-Object { [string]$_.InternalName } | Where-Object { $_ })
                $contentFieldScanSucceeded = $true
            }
            catch {
                Write-Verbose "Could not read text/URL fields from $($library.Title): $($_.Exception.Message)"
            }
        }

        $fields = @(@('FileRef', 'FileLeafRef', 'FileDirRef', 'UniqueId', 'FSObjType', 'Modified', 'Title') + $contentFieldNames | Select-Object -Unique)
        $itemsResult = Invoke-PnPConnectionOperation -Url $subsiteUrl -ClientId $ClientId -Connection $subConn -Label ('Get items: ' + $library.Title) -MaxAttempts 3 -Operation {
            param($activeConnection)
            Get-PnPListItem -List $library.Title -Connection $activeConnection -PageSize 500 -Fields $fields
        }
        $subConn = $itemsResult.Connection
        $items = @($itemsResult.Result)

        $itemIndex = 0
        foreach ($item in $items) {
            $itemIndex++
            if ($itemIndex % 50 -eq 0) { Write-Host "    ...$itemIndex / $($items.Count)" }

            if ($item.FieldValues['FSObjType'] -eq 1) { continue }

            $fileRef = [string]$item.FieldValues['FileRef']
            $fileLeaf = [string]$item.FieldValues['FileLeafRef']
            $fileDirRef = [string]$item.FieldValues['FileDirRef']
            $uniqueId = $item.FieldValues['UniqueId']
            if ([string]::IsNullOrWhiteSpace($fileRef)) { continue }
            if ($fileRef.LastIndexOf('/') -lt 0) { continue }
            if ([string]::IsNullOrWhiteSpace($fileLeaf)) {
                $fileLeaf = ($fileRef -split '/')[-1]
            }
            if ([string]::IsNullOrWhiteSpace($fileLeaf)) { continue }
            if ([string]::IsNullOrWhiteSpace($fileDirRef)) {
                $fileDirRef = $fileRef.Substring(0, $fileRef.LastIndexOf('/'))
                if (-not $fileDirRef) { $fileDirRef = '/' }
            }

            $ext = Get-NormalizedExtension -FileName $fileLeaf
            $oldCopyLinkUrl = Get-DocumentBrowserUrl -TenantBase $RootUrl -FileRef $fileRef -Extension $ext -UniqueId $uniqueId
            $oldResolvedUrl = Get-ResolvedDocumentUrl -TenantBase $RootUrl -SiteServerRelativeUrl $oldSiteServerRelative -FileRef $fileRef -FolderRef $fileDirRef -FileName $fileLeaf -Extension $ext -UniqueId $uniqueId
            $modifiedValue = $item.FieldValues['Modified']

            $newFileRef = ''
            $newFound = 'No'
            $newCopyLinkUrl = ''
            $newResolvedUrl = ''
            $notes = ''

            $rowKey = $fileRef.TrimEnd('/').ToLowerInvariant()
            $existingRow = $mergedRowsByKey[$rowKey]
            $locationOverride = Get-NewDocumentLocationOverride -FileRef $fileRef
            $newSubsiteForItem = if ($siteMatch) { $newSiteLinkBaseAbsolute } else { '' }
            if ($locationOverride) {
                $newSubsiteForItem = $RootUrl + ([Uri]$locationOverride.ConnectionBase).AbsolutePath.TrimEnd('/')
            }
            $canSkipLiveLookup = $librarySupportsFileLookup -and $SkipAlreadyConfirmedDocuments -and $existingRow -and
            (-not $locationOverride) -and
            ($existingRow.NewFound -eq 'Yes') -and
            ([string]$existingRow.Modified -eq [string]$modifiedValue)

            if ($canSkipLiveLookup) {
                $newFileRef = [string]$existingRow.NewFileRef
                $newFound = 'Yes'
                $newCopyLinkUrl = [string]$existingRow.NewCopyLinkUrl
                $newResolvedUrl = [string]$existingRow.NewResolvedUrl
                $notes = [string]$existingRow.Notes
                $newFoundCount++
                $newFoundFromCacheCount++
            }
            elseif (-not $librarySupportsFileLookup) {
                if ($siteMatch) {
                    $relativeWithinSite = $fileRef.Substring($oldSiteServerRelative.Length)
                    $newFileRef = $newSiteLinkServerRelative + $relativeWithinSite
                    $newCopyLinkUrl = $RootUrl + (Get-EncodedServerRelativePath -Path $newFileRef)
                    $notes = 'Site Contents item inventoried only; live new-file lookup is only run for document/page libraries.'
                }
                else {
                    $notes = 'Site Contents item inventoried only; site not found in Old/New Site URL CSV.'
                }
            }
            elseif (-not $siteMatch) {
                $notes = 'Site not found in Old/New Site URL CSV.'
            }
            else {
                $relativeWithinSite = $fileRef.Substring($oldSiteServerRelative.Length)
                $relativeDirWithinSite = $fileDirRef.Substring($oldSiteServerRelative.Length)
                $lookupScanBaseAbsolute = $newSiteScanBaseAbsolute
                $lookupScanServerRelative = $newSiteScanServerRelative
                $lookupLinkServerRelative = $newSiteLinkServerRelative
                $lookupResolvedSiteServerRelative = $newSiteLinkServerRelative
                if ($locationOverride) {
                    $relativeWithinOverrideLibrary = $fileRef.Substring($locationOverride.OldLibraryPath.Length)
                    $relativeDirWithinOverrideLibrary = $fileDirRef.Substring($locationOverride.OldLibraryPath.Length)
                    $lookupScanBaseAbsolute = $locationOverride.ConnectionBase
                    $lookupScanServerRelative = $locationOverride.NewLibraryPath
                    $lookupLinkServerRelative = $locationOverride.NewLibraryPath
                    $lookupResolvedSiteServerRelative = ([Uri]$locationOverride.ConnectionBase).AbsolutePath.TrimEnd('/')
                    $relativeWithinSite = $relativeWithinOverrideLibrary
                    $relativeDirWithinSite = $relativeDirWithinOverrideLibrary
                }

                $newFileRefForQuery = $lookupScanServerRelative + $relativeWithinSite
                $newFileRef = $lookupLinkServerRelative + $relativeWithinSite
                $newFileDirRefForLink = $lookupLinkServerRelative + $relativeDirWithinSite

                if (-not $newConnectionsByBase.ContainsKey($lookupScanBaseAbsolute)) {
                    try {
                        $newConnectionsByBase[$lookupScanBaseAbsolute] = Connect-ValidatedPnP -Url $lookupScanBaseAbsolute -ClientId $ClientId
                    }
                    catch {
                        $newConnectionsByBase[$lookupScanBaseAbsolute] = $null
                        Write-Warning "  Could not connect to new site $lookupScanBaseAbsolute : $($_.Exception.Message)"
                    }
                }

                $newConn = $newConnectionsByBase[$lookupScanBaseAbsolute]
                if (-not $newConn) {
                    $notes = 'Could not connect to new site.'
                }
                else {
                    try {
                        $newItemResult = Invoke-PnPConnectionOperation -Url $lookupScanBaseAbsolute -ClientId $ClientId -Connection $newConn -Label ('Lookup new item: ' + $newFileRefForQuery) -MaxAttempts 2 -Operation {
                            param($activeConnection)
                            Get-PnPFile -Url $newFileRefForQuery -AsListItem -Connection $activeConnection
                        }
                        $newConnectionsByBase[$lookupScanBaseAbsolute] = $newItemResult.Connection
                        $newItem = $newItemResult.Result
                        $newUniqueId = $newItem.FieldValues['UniqueId']

                        $newFound = 'Yes'
                        $newCopyLinkUrl = Get-DocumentBrowserUrl -TenantBase $RootUrl -FileRef $newFileRef -Extension $ext -UniqueId $newUniqueId
                        $newResolvedUrl = Get-ResolvedDocumentUrl -TenantBase $RootUrl -SiteServerRelativeUrl $lookupResolvedSiteServerRelative -FileRef $newFileRef -FolderRef $newFileDirRefForLink -FileName $fileLeaf -Extension $ext -UniqueId $newUniqueId
                        $newFoundCount++
                    }
                    catch {
                        $notes = 'Not found at the expected new path - folder may have been renamed during migration. Verify manually.'
                        $notFoundCountThisSite++
                        $newNotFoundCount++
                    }
                }
            }

            $filesScanned++
            $extKey = if ([string]::IsNullOrWhiteSpace($ext)) { '(none)' } else { $ext }
            if (-not $extensionCounts.ContainsKey($extKey)) { $extensionCounts[$extKey] = 0 }
            $extensionCounts[$extKey]++

            if ($scanThisSiteContent) {
                $sourceHyperlinks = New-Object System.Collections.Generic.List[Object]
                $sourceContentScanned = $false
                $fileContentScanSucceeded = $false

                $pageContentFields = @('CanvasContent1', 'WikiField', 'PublishingPageContent', 'LayoutWebpartsContent')
                if ($librarySupportsFileLookup -and ($ContentScanExtensions -contains $ext)) {
                    $contentStream = $null
                    try {
                        $contentResult = Invoke-PnPConnectionOperation -Url $subsiteUrl -ClientId $ClientId -Connection $subConn -Label ('Download content: ' + $fileRef) -MaxAttempts 2 -Operation {
                            param($activeConnection)
                            Get-PnPFile -Url $fileRef -AsMemoryStream -Connection $activeConnection
                        }
                        $subConn = $contentResult.Connection
                        $contentStream = $contentResult.Result
                        if ($ext -eq '.pdf') {
                            foreach ($link in @(Get-PdfHyperlinks -Bytes $contentStream.ToArray())) { $sourceHyperlinks.Add($link) }
                        }
                        elseif ($ext -in @('.htm', '.html')) {
                            $htmlContent = [System.Text.Encoding]::UTF8.GetString($contentStream.ToArray())
                            foreach ($link in @(Get-HtmlHyperlinks -Html $htmlContent)) { $sourceHyperlinks.Add($link) }
                        }
                        else {
                            foreach ($link in @(Get-OfficeDocumentHyperlinks -Stream $contentStream)) { $sourceHyperlinks.Add($link) }
                        }
                        $fileContentScanSucceeded = $true
                    }
                    catch {
                        Write-Verbose "Could not scan content of $fileRef : $($_.Exception.Message)"
                    }
                    finally {
                        if ($contentStream) { $contentStream.Dispose() }
                    }
                }

                foreach ($fieldName in $contentFieldNames) {
                    $isPageCanvasField = ($library.Title -eq 'Site Pages' -and $pageContentFields -contains $fieldName)
                    $fieldLinks = if ($isPageCanvasField) {
                        Get-CanvasContentHyperlinks -CanvasContent ([string]$item.FieldValues[$fieldName])
                    }
                    else {
                        Get-FieldValueHyperlinks -Value $item.FieldValues[$fieldName]
                    }
                    foreach ($link in @($fieldLinks)) {
                        $sourceHyperlinks.Add($link)
                    }
                }
                $requiresFileContentScan = $librarySupportsFileLookup -and ($ContentScanExtensions -contains $ext)
                if ($contentFieldScanSucceeded -and (-not $requiresFileContentScan -or $fileContentScanSucceeded)) {
                    $sourceContentScanned = $true
                }

                if ($sourceContentScanned) {
                    $sourceDocumentKey = $oldCopyLinkUrl.Trim().ToLowerInvariant()
                    $scannedEmbeddedSourcesBySite[$siteOutputKey][($sourceDocumentKey + "`n" + $library.Title)] = $true
                }

                $embeddedSourceNewFound = $newFound
                $embeddedSourceNewDocumentUrl = if ($newFound -eq 'Yes') { $newCopyLinkUrl } else { '' }
                if ($sourceContentScanned -and $sourceHyperlinks.Count -gt 0 -and -not $librarySupportsFileLookup -and $siteMatch) {
                    if (-not $newConnectionsByBase.ContainsKey($newSiteScanBaseAbsolute)) {
                        try {
                            $newConnectionsByBase[$newSiteScanBaseAbsolute] = Connect-ValidatedPnP -Url $newSiteScanBaseAbsolute -ClientId $ClientId
                        }
                        catch {
                            $newConnectionsByBase[$newSiteScanBaseAbsolute] = $null
                        }
                    }
                    $listLookupConnection = $newConnectionsByBase[$newSiteScanBaseAbsolute]
                    if ($listLookupConnection) {
                        try {
                            $newListItemResult = Invoke-PnPConnectionOperation -Url $newSiteScanBaseAbsolute -ClientId $ClientId -Connection $listLookupConnection -Label ('Lookup new list item: ' + $library.Title + ' #' + $item.Id) -MaxAttempts 2 -Operation {
                                param($activeConnection)
                                Get-PnPListItem -List $library.Title -Id $item.Id -Fields @('FileRef') -Connection $activeConnection
                            }
                            $newConnectionsByBase[$newSiteScanBaseAbsolute] = $newListItemResult.Connection
                            $newListItemFileRef = [string]$newListItemResult.Result.FieldValues['FileRef']
                            if ([string]::IsNullOrWhiteSpace($newListItemFileRef)) {
                                $relativeWithinListSite = $fileRef.Substring($oldSiteServerRelative.Length)
                                $newListItemFileRef = $newSiteLinkServerRelative + $relativeWithinListSite
                            }
                            $embeddedSourceNewFound = 'Yes'
                            $embeddedSourceNewDocumentUrl = $RootUrl + (Get-EncodedServerRelativePath -Path $newListItemFileRef)
                        }
                        catch {
                            $embeddedSourceNewFound = 'No'
                            $embeddedSourceNewDocumentUrl = ''
                        }
                    }
                }

                foreach ($link in $sourceHyperlinks) {
                    $embeddedUrl = [string]$link.Url
                    try {
                        $embeddedUri = [Uri]$embeddedUrl
                        if (-not $embeddedUri.IsAbsoluteUri) {
                            $sourceFolderUrl = $RootUrl + $fileDirRef.TrimEnd('/') + '/'
                            $embeddedUri = [Uri]::new([Uri]$sourceFolderUrl, $embeddedUri)
                        }
                        $embeddedUrl = $embeddedUri.AbsoluteUri
                    }
                    catch {
                        continue
                    }
                    if (-not (Test-IsInternalTenantUrl -Url $embeddedUrl -RootUrl $RootUrl)) { continue }
                    if ($scanOldBaseOnly -and -not (Test-IsUrlWithinOldSiteBase -Url $embeddedUrl -SiteUrlMap $siteUrlMap)) { continue }

                    $embeddedTargetKey = $embeddedUrl.Trim().ToLowerInvariant()
                    if (-not $embeddedTargetResolutionCache.ContainsKey($embeddedTargetKey)) {
                        $embeddedTargetResolutionCache[$embeddedTargetKey] = Resolve-EmbeddedTargetUrl -Url $embeddedUrl -RootUrl $RootUrl -SiteUrlMap $siteUrlMap -ClientId $ClientId -ConnectionCache $newConnectionsByBase
                    }
                    $embeddedTarget = $embeddedTargetResolutionCache[$embeddedTargetKey]

                    $embeddedRow = [PSCustomObject][ordered]@{
                        Site                 = $selectedScopeUrl
                        Subsite              = $subsiteUrl
                        LibraryorListName    = $library.Title
                        FileName             = if ($librarySupportsFileLookup) { $fileLeaf } elseif ($item.FieldValues['Title']) { [string]$item.FieldValues['Title'] } else { $fileLeaf }
                        Extension            = if ($librarySupportsFileLookup) { $ext } else { '' }
                        DocumentorItemURL    = $oldCopyLinkUrl
                        EmbeddedURLText      = [string]$link.Text
                        EmbeddedURL          = $embeddedUrl
                        NewFound             = $embeddedSourceNewFound
                        NewSubsite           = $newSubsiteForItem
                        NewDocumentorItemURL = $embeddedSourceNewDocumentUrl
                        EmbeddedNewFound     = $embeddedTarget.Found
                        EmbeddedNewSubsite   = $embeddedTarget.NewSubsite
                        EmbeddedNewURL       = $embeddedTarget.NewUrl
                        EmbeddedResolution   = $embeddedTarget.Resolution
                    }
                    $embeddedKey = Get-Sha256Hex -Text ($sourceDocumentKey + "`n" + $library.Title + "`n" + $embeddedRow.EmbeddedURLText + "`n" + $embeddedUrl)
                    $embeddedLinksBySite[$siteOutputKey][$embeddedKey] = $embeddedRow
                }
            }

            $row = [PSCustomObject][ordered]@{
                Site           = $selectedScopeUrl
                Subsite        = $subsiteUrl
                Library        = $library.Title
                FileName       = $fileLeaf
                Extension      = $ext
                OldFileRef     = $fileRef
                OldCopyLinkUrl = $oldCopyLinkUrl
                OldResolvedUrl = $oldResolvedUrl
                NewFileRef     = $newFileRef
                NewFound       = $newFound
                NewCopyLinkUrl = $newCopyLinkUrl
                NewResolvedUrl = $newResolvedUrl
                Modified       = $modifiedValue
                Notes          = $notes
            }

            # Appended immediately (the temp file for this run starts empty), so a crash
            # never loses more than the row in flight - $mergedRowsByKey is only read from
            # the current master snapshot, never written back to until the scan publishes.
            Add-CsvRowWithRetry -Path $outputPath -Row $row
        }

        Save-CounterCheckpoint -Path $counterCheckpointPath -OutputPath $outputPath -FilesScanned $filesScanned -NewFoundCount $newFoundCount -NewNotFoundCount $newNotFoundCount -LastLibraryCompleted ($subsiteUrl + ' | ' + $library.Title)
    }

    if ($scanThisSiteContent) {
        foreach ($navigationLocation in @('TopNavigationBar', 'QuickLaunch', 'Footer')) {
            $navigationContainerName = "Navigation - $navigationLocation"
            try {
                $navigationResult = Invoke-PnPConnectionOperation -Url $subsiteUrl -ClientId $ClientId -Connection $subConn -Label ('Get navigation: ' + $navigationLocation) -MaxAttempts 2 -Operation {
                    param($activeConnection)
                    Get-PnPNavigationNode -Location $navigationLocation -Tree -Connection $activeConnection
                }
                $subConn = $navigationResult.Connection
                $navigationSourceKey = $subsiteUrl.TrimEnd('/').ToLowerInvariant() + "`n" + $navigationContainerName
                $scannedEmbeddedSourcesBySite[$siteOutputKey][$navigationSourceKey] = $true
            }
            catch {
                Write-Verbose "Could not read $navigationLocation navigation for $subsiteUrl : $($_.Exception.Message)"
                continue
            }

            $sourceNavigationNodes = @(Get-NavigationNodeTreeItems -Nodes @($navigationResult.Result))
            $destinationNodesByTitle = @{}
            if ($siteMatch) {
                if (-not $newConnectionsByBase.ContainsKey($newSiteScanBaseAbsolute)) {
                    try {
                        $newConnectionsByBase[$newSiteScanBaseAbsolute] = Connect-ValidatedPnP -Url $newSiteScanBaseAbsolute -ClientId $ClientId
                    }
                    catch {
                        $newConnectionsByBase[$newSiteScanBaseAbsolute] = $null
                    }
                }
                $destinationConnection = $newConnectionsByBase[$newSiteScanBaseAbsolute]
                if ($destinationConnection) {
                    try {
                        $destinationNavigationResult = Invoke-PnPConnectionOperation -Url $newSiteScanBaseAbsolute -ClientId $ClientId -Connection $destinationConnection -Label ('Get destination navigation: ' + $navigationLocation) -MaxAttempts 2 -Operation {
                            param($activeConnection)
                            Get-PnPNavigationNode -Location $navigationLocation -Tree -Connection $activeConnection
                        }
                        $newConnectionsByBase[$newSiteScanBaseAbsolute] = $destinationNavigationResult.Connection
                        foreach ($destinationNode in @(Get-NavigationNodeTreeItems -Nodes @($destinationNavigationResult.Result))) {
                            $destinationTitle = [string]$destinationNode.Title
                            if (-not [string]::IsNullOrWhiteSpace($destinationTitle)) { $destinationNodesByTitle[$destinationTitle] = $destinationNode }
                        }
                    }
                    catch {
                        Write-Verbose "Could not read destination $navigationLocation navigation at $newSiteScanBaseAbsolute : $($_.Exception.Message)"
                    }
                }
            }

            foreach ($navigationNode in $sourceNavigationNodes) {
                $embeddedUrl = [string]$navigationNode.Url
                if ([string]::IsNullOrWhiteSpace($embeddedUrl)) { continue }
                try {
                    $navigationUri = [Uri]$embeddedUrl
                    if (-not $navigationUri.IsAbsoluteUri) {
                        $navigationUri = [Uri]::new([Uri]($subsiteUrl.TrimEnd('/') + '/'), $navigationUri)
                    }
                    $embeddedUrl = $navigationUri.AbsoluteUri
                }
                catch {
                    continue
                }
                if (-not (Test-IsInternalTenantUrl -Url $embeddedUrl -RootUrl $RootUrl)) { continue }
                if ($scanOldBaseOnly -and -not (Test-IsUrlWithinOldSiteBase -Url $embeddedUrl -SiteUrlMap $siteUrlMap)) { continue }

                $embeddedTargetKey = $embeddedUrl.Trim().ToLowerInvariant()
                if (-not $embeddedTargetResolutionCache.ContainsKey($embeddedTargetKey)) {
                    $embeddedTargetResolutionCache[$embeddedTargetKey] = Resolve-EmbeddedTargetUrl -Url $embeddedUrl -RootUrl $RootUrl -SiteUrlMap $siteUrlMap -ClientId $ClientId -ConnectionCache $newConnectionsByBase
                }
                $embeddedTarget = $embeddedTargetResolutionCache[$embeddedTargetKey]

                $nodeTitle = [string]$navigationNode.Title
                $destinationNode = if ($destinationNodesByTitle.ContainsKey($nodeTitle)) { $destinationNodesByTitle[$nodeTitle] } else { $null }
                $embeddedRow = [PSCustomObject][ordered]@{
                    Site                 = $selectedScopeUrl
                    Subsite              = $subsiteUrl
                    LibraryorListName    = $navigationContainerName
                    FileName             = $nodeTitle
                    Extension            = ''
                    DocumentorItemURL    = $subsiteUrl
                    EmbeddedURLText      = $nodeTitle
                    EmbeddedURL          = $embeddedUrl
                    NewFound             = if ($destinationNode) { 'Yes' } else { 'No' }
                    NewSubsite           = if ($siteMatch) { $newSiteLinkBaseAbsolute } else { '' }
                    NewDocumentorItemURL = if ($destinationNode) { $newSiteLinkBaseAbsolute } else { '' }
                    EmbeddedNewFound     = $embeddedTarget.Found
                    EmbeddedNewSubsite   = $embeddedTarget.NewSubsite
                    EmbeddedNewURL       = $embeddedTarget.NewUrl
                    EmbeddedResolution   = $embeddedTarget.Resolution
                }
                $rowKey = Get-Sha256Hex -Text ($subsiteUrl.ToLowerInvariant() + "`n" + $navigationContainerName + "`n" + $nodeTitle.ToLowerInvariant() + "`n" + $embeddedUrl.ToLowerInvariant())
                $embeddedLinksBySite[$siteOutputKey][$rowKey] = $embeddedRow
            }
        }
    }

    if ($filesScanned -eq 0) {
        Write-Warning "  No documents were found in $subsiteUrl."
    }

    Publish-MasterLookupTable -TempPath $outputPath -MasterPath $masterPath
    $fixerRowsWritten = Export-FixerLookupTableFromMaster -MasterPath $masterPath -FixerPath $fixerPath -RequiredExtensions $FixerRequiredExtensions
    $publishedSiteIdentifiers[$siteOutputKey] = [PSCustomObject]@{
        Identifier       = $siteIdentifier
        IsSitesWeb       = $isSitesWeb
        MasterFolder    = $currentMasterFolder
        FixerFolder     = $currentFixerFolder
        EmbeddedFolder  = $currentEmbeddedFolder
    }

    Write-Host ("  Wrote $filesScanned row(s) to master $masterPath")
    Write-Host ("  Wrote {0} filtered Fixer row(s) to {1}" -f $fixerRowsWritten, $fixerPath)
    Write-Host ("  New document found: {0} ({1} via cached prior confirmation)   Not found: {2}" -f $newFoundCount, $newFoundFromCacheCount, $newNotFoundCount)
    if ($notFoundCountThisSite -gt 0) {
        Write-Host ("  {0} document(s) could not be located at their expected new path - see Notes column." -f $notFoundCountThisSite)
    }

    $filesScannedTotal += $filesScanned
    $newFoundCountTotal += $newFoundCount
    $newFoundFromCacheCountTotal += $newFoundFromCacheCount
    $newNotFoundCountTotal += $newNotFoundCount
    foreach ($extKey in $extensionCounts.Keys) {
        if (-not $extensionCountsTotal.ContainsKey($extKey)) { $extensionCountsTotal[$extKey] = 0 }
        $extensionCountsTotal[$extKey] += $extensionCounts[$extKey]
    }
}

if ($filesScannedTotal -eq 0) {
    Write-Warning "`nNo documents were found in the selected subsites. Canonical header-only CSVs were written for each successfully scanned site."
}
else {
    Write-Host ("`nAll sites combined: refreshed {0} master row(s) across {1} site CSV(s)." -f $filesScannedTotal, $selectedSubsites.Count)

    Write-Host "`nExtension breakdown:"
    foreach ($extKey in @($extensionCountsTotal.Keys | Sort-Object)) {
        Write-Host ("  {0,-10} {1}" -f $extKey, $extensionCountsTotal[$extKey])
    }

    Write-Host ("`nNew document found: {0} ({1} via cached prior confirmation, no live lookup)   Not found: {2}" -f $newFoundCountTotal, $newFoundFromCacheCountTotal, $newNotFoundCountTotal)
}

foreach ($siteOutputKey in @($publishedSiteIdentifiers.Keys | Sort-Object)) {
    $outputPlan = $publishedSiteIdentifiers[$siteOutputKey]
    $siteIdentifier = $outputPlan.Identifier
    $embeddedOutputPath = Join-Path -Path $outputPlan.EmbeddedFolder -ChildPath ("$siteIdentifier-EmbeddedLinks.csv")
    $embeddedRowsWritten = Publish-EmbeddedLinksLookupTable -Folder $outputPlan.EmbeddedFolder -Identifier $siteIdentifier -RunStamp $runStamp -RowsByKey $embeddedLinksBySite[$siteOutputKey] -ScannedSourceUrls $scannedEmbeddedSourcesBySite[$siteOutputKey]
    Write-Host "`nWrote $embeddedRowsWritten embedded-link row(s): $embeddedOutputPath"
    if ($outputPlan.IsSitesWeb) {
        Assert-SiteOutputInventory -Identifier $siteIdentifier -MasterFolder $outputPlan.MasterFolder -FixerFolder $outputPlan.FixerFolder -EmbeddedFolder $outputPlan.EmbeddedFolder -CheckSitesPrefixLegacy $true -LegacyMasterFolder $MasterFolder -LegacyFixerFolder $FixerOutputFolder -LegacyEmbeddedFolder $EmbeddedLinksFolder
    }
    else {
        Assert-SiteOutputInventory -Identifier $siteIdentifier -MasterFolder $outputPlan.MasterFolder -FixerFolder $outputPlan.FixerFolder -EmbeddedFolder $outputPlan.EmbeddedFolder
    }
}
