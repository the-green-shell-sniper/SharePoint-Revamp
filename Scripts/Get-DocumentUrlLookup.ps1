<#
.SYNOPSIS
    Builds a per-site old-URL -> new-URL lookup for documents, for use by the
    Central Park Link Finder (and anyone else) when a bookmark can't be
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

.PARAMETER SitesEmbeddedLinksFolder
    Folder for per-site embedded-link reports scanned from /sites webs.

.PARAMETER SitesMasterFolder
    Folder for [site]-master.csv files scanned from /sites webs (temp scans go in its Temp subfolder).

.PARAMETER SitesFixerOutputFolder
    Folder for [site]-fixer.csv files scanned from /sites webs.

.PARAMETER GhostscriptPath
    Optional path to gswin64c.exe for Deep PDF OCR. If blank, the script
    searches the standard C:\Program Files\gs installation folders.

.PARAMETER TesseractPath
    Path to tesseract.exe for Deep PDF OCR. English language data must be installed.

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

.PARAMETER EnableDeepPdfOcr
    OCR rendered PDF pages for URLs visible only in scanned/image content.
    If omitted during an embedded-link scan, the script prompts with No as the default.

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
    [string]$SitesEmbeddedLinksFolder = "C:\Users\!bposey\OneDrive - Womens Healthcare Associates LLC\Gary Henderson's files - Project Files_bposey\DocumentUrlLookup-OutputFiles\Sites-EmbeddedLinks",

    [Parameter(Mandatory = $false)]
    [string]$SitesMasterFolder = "C:\Users\!bposey\OneDrive - Womens Healthcare Associates LLC\Gary Henderson's files - Project Files_bposey\DocumentUrlLookup-OutputFiles\Sites-Masters",

    [Parameter(Mandatory = $false)]
    [string]$SitesFixerOutputFolder = "C:\Users\!bposey\OneDrive - Womens Healthcare Associates LLC\Gary Henderson's files - Project Files_bposey\DocumentUrlLookup-OutputFiles\Sites-FixerFiles",

    [Parameter(Mandatory = $false)]
    [string]$GhostscriptPath = '',

    [Parameter(Mandatory = $false)]
    [string]$TesseractPath = 'C:\Program Files\Tesseract-OCR\tesseract.exe',

    [Parameter(Mandatory = $false)]
    [string]$JavaPath = 'java.exe',

    [Parameter(Mandatory = $false)]
    [string]$TikaAppPath = 'C:\Tools\DocumentUrlLookup\Tika\tika-app-4.1.0.jar',

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
    [switch]$EnableDocumentContentScan,

    # OCR rendered PDF pages to find URLs visible only in scanned/image content.
    # If omitted, the operator is prompted with a No default.
    [Parameter(Mandatory = $false)]
    [switch]$EnableDeepPdfOcr
)
. (Join-Path $PSScriptRoot 'SPOutputLimits.Common.ps1')

$ErrorActionPreference = "Stop"

$ScriptSchemaVersion = 'DocumentUrlLookup-v1'
$InventoryCacheSchemaVersion = 'v2'
$PdfPigAssemblyFolder = 'C:\Users\!bposey\Desktop\PdfPig'
$SkipListTitles = @('Form Templates', 'Preservation Hold Library')
$FixerRequiredExtensions = @('.docx', '.doc', '.pdf', '.pptx', '.xls', '.xlsx', '.ppt', '.vsdx', '.vsd', '.png', '.aspx', '.one')
$FileLookupListBaseTemplates = @(101, 119, 850)
$OpenXmlContentExtensions = @(
    '.docx', '.docm', '.dotx', '.dotm',
    '.xlsx', '.xlsm', '.xltx', '.xltm',
    '.pptx', '.pptm', '.potx', '.potm', '.ppsx', '.ppsm'
)
$ContentScanExtensions = @(
    $OpenXmlContentExtensions + @(
        '.doc', '.xls', '.ppt', '.rtf', '.odt', '.ods', '.odp', '.vsd', '.vsdx', '.one', '.msg',
        '.pdf', '.txt', '.csv', '.htm', '.html', '.aspx'
    )
)

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
        throw "PdfPig is not configured. Copy all seven UglyToad.PdfPig*.dll files from the PdfPig NuGet package's lib\net8.0 folder to '$PdfPigAssemblyFolder'. Missing: $($missingAssemblies -join ', ')"
    }

    $script:PdfPigAssemblyFolder = $PdfPigAssemblyFolder
    if (-not $script:PdfPigAssemblyResolveHandler) {
        $script:PdfPigAssemblyResolveHandler = [System.ResolveEventHandler]{
            param($resolveSender, $resolveEventArgs)

            $null = $resolveSender
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
            $alreadyLoaded = [System.Reflection.Assembly]::LoadFrom($assemblyPath)
        }

        # A missing dependency would otherwise only surface when each PDF is opened.
        foreach ($reference in $alreadyLoaded.GetReferencedAssemblies()) {
            try {
                [System.Reflection.Assembly]::Load($reference) | Out-Null
            }
            catch {
                throw "PdfPig dependency '$($reference.FullName)' required by '$assemblyName' could not be loaded. Replace the DLLs in '$PdfPigAssemblyFolder' with the seven UglyToad.PdfPig*.dll files from the PdfPig NuGet package's lib\net8.0 folder (no external dependencies), then start a new PowerShell session. Detail: $($_.Exception.Message)"
            }
        }
    }
}

# Only the human-visible content parts of the OOXML package (the document/sheet/slide
# bodies) are read - not core.xml/app.xml metadata, not comments - since we only care
# about links a reader would actually see and could click.
function Get-OfficeDocumentHyperlink {
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
                    if ($rel.TargetMode -eq 'External' -and $rel.Target -and [string]$rel.Type -match '(?i)(?:^|/)hyperlink$') {
                        $relationMap[$rel.Id] = [string]$rel.Target
                    }
                }
                if ($relationMap.Count -gt 0) { $relationsByXml[$xmlPath] = $relationMap }
            }
            catch {
                throw "Could not parse relationship XML in $($entry.FullName): $($_.Exception.Message)"
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
            if ($entry.FullName -notmatch '(?i)(?:^|/)word/(?:document|header\d+|footer\d+|comments\d*|footnotes|endnotes)\.xml$|(?:^|/)xl/worksheets/sheet\d*\.xml$|(?:^|/)ppt/(?:slides/slide\d*|notesSlides/notesSlide\d*)\.xml$') { continue }
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
                $ns.AddNamespace('a', 'http://schemas.openxmlformats.org/drawingml/2006/main')
                $lastLink = $null
                $lastRun = $null
                $lastRelId = ''
                foreach ($node in $contentXml.SelectNodes('//*[@r:id]', $ns)) {
                    $relId = $node.GetAttribute('id', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
                    if (-not $relationsByXml[$entry.FullName].ContainsKey($relId)) { continue }
                    # End-of-paragraph formatting repeats the link but is not clickable text.
                    if ($node.ParentNode -and $node.ParentNode.LocalName -eq 'endParaRPr') { continue }

                    $textRun = if ($node.LocalName -eq 'hlinkClick') { $node.SelectSingleNode('ancestor::a:r[1]', $ns) } else { $null }
                    $displayText = if ($textRun) { [string]$textRun.InnerText } else { [string]$node.InnerText }
                    # One visible link is often split across adjacent runs that share the same relationship.
                    if ($textRun -and $lastRun -and $relId -eq $lastRelId -and [object]::ReferenceEquals($textRun.PreviousSibling, $lastRun)) {
                        $lastLink.Text += $displayText
                        $lastRun = $textRun
                        continue
                    }
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
                    if ([string]::IsNullOrWhiteSpace($displayText) -and $entry.FullName -match 'ppt/' -and $node.LocalName -eq 'hlinkClick') {
                        $shape = $node.SelectSingleNode('ancestor::p:sp[1]', $ns)
                        if ($shape) { $displayText = [string]$shape.InnerText }
                    }
                    $location = $entry.FullName
                    if ($node.Attributes['ref']) { $location += ' | cell ' + $node.Attributes['ref'].Value }
                    $lastLink = [PSCustomObject]@{ Text = $displayText; Url = [string]$relationsByXml[$entry.FullName][$relId]; Location = $location }
                    $lastRun = $textRun
                    $lastRelId = $relId
                    $found.Add($lastLink)
                }
            }
            catch {
                throw "Could not parse hyperlinks in $($entry.FullName): $($_.Exception.Message)"
            }
        }
    }
    finally {
        $archive.Dispose()
    }

    foreach ($link in $found) { $link.Text = ([string]$link.Text).Trim() }
    return $found.ToArray()
}

function Get-InternalUrlOccurrencesFromText {
    param([Parameter(Mandatory = $true)][AllowEmptyString()][string]$Text)

    $found = New-Object System.Collections.Generic.List[Object]
    foreach ($match in [regex]::Matches($Text, '(?i)(?:https?:)?//whallc\.sharepoint\.com[^\s<>"''\]\)]+')) {
        $url = $match.Value.TrimEnd('.', ',', ';', ':', '!', '?', ')', ']', '}')
        if (-not [string]::IsNullOrWhiteSpace($url)) {
            $found.Add([PSCustomObject]@{ Text = $url; Url = $url })
        }
    }
    return $found.ToArray()
}

function Get-PlainTextHyperlink {
    param([Parameter(Mandatory = $true)][System.IO.Stream]$Stream)

    $reader = [System.IO.StreamReader]::new($Stream, [System.Text.Encoding]::UTF8, $true, 1024, $true)
    try {
        return @(Get-InternalUrlOccurrencesFromText -Text $reader.ReadToEnd())
    }
    finally { $reader.Dispose() }
}

function Invoke-PdfExtractionProcess {
    param(
        [Parameter(Mandatory = $true)][string]$FilePath,
        [Parameter(Mandatory = $true)][string[]]$ArgumentList,
        [Parameter(Mandatory = $false)][int]$TimeoutSeconds = 600
    )

    $startInfo = [System.Diagnostics.ProcessStartInfo]::new()
    $startInfo.FileName = $FilePath
    $startInfo.UseShellExecute = $false
    $startInfo.CreateNoWindow = $true
    $startInfo.RedirectStandardOutput = $true
    $startInfo.RedirectStandardError = $true
    foreach ($argument in $ArgumentList) { $startInfo.ArgumentList.Add($argument) }

    $process = [System.Diagnostics.Process]::new()
    $process.StartInfo = $startInfo
    try {
        if (-not $process.Start()) { throw "Could not start PDF extraction tool: $FilePath" }
        $stdoutTask = $process.StandardOutput.ReadToEndAsync()
        $stderrTask = $process.StandardError.ReadToEndAsync()
        if (-not $process.WaitForExit($TimeoutSeconds * 1000)) {
            try { $process.Kill($true) } catch { Write-Verbose "Could not terminate timed-out PDF process: $($_.Exception.Message)" }
            throw "PDF extraction tool timed out after $TimeoutSeconds second(s): $FilePath"
        }
        $process.WaitForExit()
        $stdout = $stdoutTask.GetAwaiter().GetResult()
        $stderr = $stderrTask.GetAwaiter().GetResult()
        if ($process.ExitCode -ne 0) {
            throw "PDF extraction tool failed with exit code $($process.ExitCode): $stderr"
        }
        return [PSCustomObject]@{ StandardOutput = $stdout; StandardError = $stderr; ExitCode = $process.ExitCode }
    }
    finally {
        $process.Dispose()
    }
}

function Get-PdfOcrUrlOccurrence {
    param([Parameter(Mandatory = $true)][AllowEmptyCollection()][byte[]]$Bytes)

    if (-not (Test-Path -LiteralPath $TesseractPath -PathType Leaf)) {
        throw "Tesseract was not found: $TesseractPath"
    }
    if (-not $script:ResolvedGhostscriptPath) {
        $ghostscript = if (-not [string]::IsNullOrWhiteSpace($GhostscriptPath)) {
            Get-Item -LiteralPath $GhostscriptPath -ErrorAction SilentlyContinue
        }
        else {
            Get-ChildItem -Path 'C:\Program Files\gs\gs*\bin\gswin64c.exe' -File -ErrorAction SilentlyContinue |
                Sort-Object LastWriteTimeUtc -Descending | Select-Object -First 1
        }
        if (-not $ghostscript) { throw 'Ghostscript gswin64c.exe was not found. Install it or pass -GhostscriptPath.' }
        $script:ResolvedGhostscriptPath = $ghostscript.FullName
    }

    $found = New-Object System.Collections.Generic.List[Object]
    $workFolder = Join-Path ([System.IO.Path]::GetTempPath()) ('Get-DocumentUrlLookup-PdfOcr-' + [Guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Path $workFolder -Force | Out-Null
    try {
        $pdfPath = Join-Path $workFolder 'source.pdf'
        $pagePattern = Join-Path $workFolder 'page-%06d.png'
        Assert-SPOutputSize -Path $pdfPath -Bytes $Bytes.LongLength
        [void](Assert-SPOutputPath $pagePattern)
        [System.IO.File]::WriteAllBytes($pdfPath, $Bytes)
        $null = Invoke-PdfExtractionProcess -FilePath $script:ResolvedGhostscriptPath -ArgumentList @(
            '-dSAFER', '-dBATCH', '-dNOPAUSE', '-dUseCropBox', '-dFIXEDMEDIA',
            '-dDEVICEWIDTHPOINTS=612', '-dDEVICEHEIGHTPOINTS=792', '-dPDFFitPage',
            '-sDEVICE=pnggray', '-r300',
            '-dTextAlphaBits=4', '-dGraphicsAlphaBits=4', "-sOutputFile=$pagePattern", $pdfPath
        ) -TimeoutSeconds 1800
        foreach ($page in Get-ChildItem -LiteralPath $workFolder -File -Filter 'page-*.png') {
            Assert-SPOutputSize -Path $page.FullName -Bytes $page.Length
        }

        $pageImages = @(Get-ChildItem -Path $workFolder -Filter 'page-*.png' -File | Sort-Object Name)
        if ($pageImages.Count -eq 0) { throw 'Ghostscript did not render any PDF pages.' }
        foreach ($pageImage in $pageImages) {
            $ocrResult = Invoke-PdfExtractionProcess -FilePath $TesseractPath -ArgumentList @(
                $pageImage.FullName, 'stdout', '-l', 'eng', '--psm', '11'
            ) -TimeoutSeconds 300
            foreach ($link in @(Get-InternalUrlOccurrencesFromText -Text $ocrResult.StandardOutput)) {
                $pageNumber = [int]([System.IO.Path]::GetFileNameWithoutExtension($pageImage.Name) -replace '^page-', '')
                $link | Add-Member -NotePropertyName Location -NotePropertyValue "PDF page $pageNumber | OCR visible URL"
                $found.Add($link)
            }
        }
    }
    finally {
        Remove-Item -LiteralPath $workFolder -Recurse -Force -ErrorAction SilentlyContinue
    }
    return $found.ToArray()
}

# PdfPig reads clickable URI annotations and extractable page text, not document metadata.
function Get-PdfHyperlink {
    param([Parameter(Mandatory = $true)][AllowEmptyCollection()][byte[]]$Bytes)

    if ($Bytes.Length -eq 0) { return [PSCustomObject]@{ Links = @(); Errors = @('PDF file is empty.') } }

    $found = New-Object System.Collections.Generic.List[Object]
    $errors = New-Object System.Collections.Generic.List[string]
    $document = $null
    try {
        $document = [UglyToad.PdfPig.PdfDocument]::Open($Bytes)
        foreach ($page in $document.GetPages()) {
            $annotationTextUrlCounts = @{}
            foreach ($hyperlink in $page.GetHyperlinks()) {
                $displayText = [string]$hyperlink.Text
                if ([string]::IsNullOrWhiteSpace($displayText) -and $hyperlink.Letters.Count -gt 0) {
                    $displayText = (@($hyperlink.Letters | ForEach-Object { [string]$_.Value }) -join '')
                }
                $annotationUrl = [string]$hyperlink.Uri
                $found.Add([PSCustomObject]@{ Text = $displayText.Trim(); Url = $annotationUrl; Location = "PDF page $($page.Number) | hyperlink annotation" })
                if (-not [string]::IsNullOrWhiteSpace($displayText) -and
                    $displayText.Trim().TrimEnd('.', ',', ';', ':') -ieq $annotationUrl.Trim().TrimEnd('.', ',', ';', ':')) {
                    $annotationKey = $annotationUrl.Trim().ToLowerInvariant()
                    if (-not $annotationTextUrlCounts.ContainsKey($annotationKey)) { $annotationTextUrlCounts[$annotationKey] = 0 }
                    $annotationTextUrlCounts[$annotationKey]++
                }
            }

            foreach ($textLink in @(Get-InternalUrlOccurrencesFromText -Text ([string]$page.Text))) {
                $textKey = ([string]$textLink.Url).Trim().ToLowerInvariant()
                if ($annotationTextUrlCounts.ContainsKey($textKey) -and $annotationTextUrlCounts[$textKey] -gt 0) {
                    $annotationTextUrlCounts[$textKey]--
                    continue
                }
                $textLink | Add-Member -NotePropertyName Location -NotePropertyValue "PDF page $($page.Number) | visible URL"
                $found.Add($textLink)
            }
        }
    }
    catch {
        $errors.Add("PdfPig could not extract PDF annotations/text: $($_.Exception.Message)")
    }
    finally {
        if ($document) { $document.Dispose() }
    }

    if ($script:DeepPdfOcr) {
        try {
            $visibleUrlOccurrences = @{}
            foreach ($existingLink in $found) {
                if ([string]::IsNullOrWhiteSpace([string]$existingLink.Url) -or
                    -not ([string]$existingLink.Text).Trim().Equals(([string]$existingLink.Url).Trim(), [System.StringComparison]::OrdinalIgnoreCase)) {
                    continue
                }
                $existingUrlKey = ([string]$existingLink.Url).Trim().ToLowerInvariant()
                if (-not $visibleUrlOccurrences.ContainsKey($existingUrlKey)) { $visibleUrlOccurrences[$existingUrlKey] = 0 }
                $visibleUrlOccurrences[$existingUrlKey]++
            }

            foreach ($ocrLink in @(Get-PdfOcrUrlOccurrence -Bytes $Bytes)) {
                $ocrUrlKey = ([string]$ocrLink.Url).Trim().ToLowerInvariant()
                if ($visibleUrlOccurrences.ContainsKey($ocrUrlKey) -and $visibleUrlOccurrences[$ocrUrlKey] -gt 0) {
                    $visibleUrlOccurrences[$ocrUrlKey]--
                    continue
                }
                $found.Add($ocrLink)
            }
        }
        catch {
            $errors.Add("PDF image OCR failed: $($_.Exception.Message)")
        }
    }

    return [PSCustomObject]@{ Links = $found.ToArray(); Errors = $errors.ToArray() }
}

function Get-HtmlHyperlink {
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

function Get-TikaDocumentHyperlink {
    param(
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][byte[]]$Bytes,
        [Parameter(Mandatory = $true)][string]$Extension
    )

    if (-not $script:ResolvedJavaPath) {
        $javaCommand = Get-Command $JavaPath -ErrorAction SilentlyContinue
        if (-not $javaCommand) { throw "Java was not found: $JavaPath" }
        $script:ResolvedJavaPath = if ($javaCommand.Source) { [string]$javaCommand.Source } else { [string]$javaCommand.Path }
    }
    if (-not (Test-Path -LiteralPath $TikaAppPath -PathType Leaf)) {
        throw "Apache Tika was not found: $TikaAppPath"
    }

    $tikaFolder = Split-Path -Path $TikaAppPath -Parent
    if (-not (Test-Path -LiteralPath (Join-Path $tikaFolder 'lib') -PathType Container)) {
        throw "Apache Tika 4.x requires its adjacent lib directory: $(Join-Path $tikaFolder 'lib')"
    }

    $workFolder = Join-Path ([System.IO.Path]::GetTempPath()) ('Get-DocumentUrlLookup-' + [Guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Path $workFolder -Force | Out-Null
    try {
        $sourcePath = Join-Path $workFolder ("source{0}" -f $Extension)
        Assert-SPOutputSize -Path $sourcePath -Bytes $Bytes.LongLength
        [System.IO.File]::WriteAllBytes($sourcePath, $Bytes)

        $startInfo = [System.Diagnostics.ProcessStartInfo]::new()
        $startInfo.FileName = $script:ResolvedJavaPath
        $startInfo.WorkingDirectory = $tikaFolder
        $startInfo.UseShellExecute = $false
        $startInfo.CreateNoWindow = $true
        $startInfo.RedirectStandardOutput = $true
        $startInfo.RedirectStandardError = $true
        foreach ($argument in @('-jar', $TikaAppPath, '-x', $sourcePath)) {
            $startInfo.ArgumentList.Add($argument)
        }

        $process = [System.Diagnostics.Process]::new()
        $process.StartInfo = $startInfo
        try {
            if (-not $process.Start()) { throw 'Could not start Apache Tika.' }
            $stdoutTask = $process.StandardOutput.ReadToEndAsync()
            $stderrTask = $process.StandardError.ReadToEndAsync()
            if (-not $process.WaitForExit(900000)) {
                try { $process.Kill($true) }
                catch { Write-Verbose "Could not terminate timed-out Tika process: $($_.Exception.Message)" }
                throw "Apache Tika timed out while parsing $Extension content."
            }
            $process.WaitForExit()
            $xhtml = $stdoutTask.GetAwaiter().GetResult()
            $stderr = $stderrTask.GetAwaiter().GetResult()
            if ($process.ExitCode -ne 0) {
                throw "Apache Tika failed with exit code $($process.ExitCode): $stderr"
            }
            return @(Get-HtmlHyperlink -Html $xhtml)
        }
        finally {
            $process.Dispose()
        }
    }
    finally {
        Remove-Item -LiteralPath $workFolder -Recurse -Force -ErrorAction SilentlyContinue
    }
}

function Get-CanvasContentHyperlink {
    param([Parameter(Mandatory = $true)][AllowEmptyString()][string]$CanvasContent)

    $found = New-Object System.Collections.Generic.List[Object]
    if ([string]::IsNullOrWhiteSpace($CanvasContent)) { return @() }

    try {
        $canvas = ConvertFrom-Json -InputObject $CanvasContent -ErrorAction Stop
    }
    catch {
        return @(Get-HtmlHyperlink -Html $CanvasContent)
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
            foreach ($link in @(Get-HtmlHyperlink -Html ([string]$innerHtml.Value))) { $found.Add($link) }
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

        $serverProcessedContent = $current.PSObject.Properties['serverProcessedContent']
        $linkCollection = if ($serverProcessedContent -and $serverProcessedContent.Value) {
            $serverProcessedContent.Value.PSObject.Properties['links']
        }
        else { $null }
        if ($linkCollection -and $linkCollection.Value) {
            $pendingLinks = New-Object System.Collections.Stack
            $pendingLinks.Push($linkCollection.Value)
            while ($pendingLinks.Count -gt 0) {
                $linkValue = $pendingLinks.Pop()
                if ($linkValue -is [string]) {
                    $candidateUrl = $linkValue.Trim()
                    if ($candidateUrl -match '(?i)^(?:https?:)?//|^/|^~site(?:collection)?/') {
                        $found.Add([PSCustomObject]@{ Text = ''; Url = $candidateUrl })
                    }
                    continue
                }
                if ($linkValue -is [System.Array]) {
                    foreach ($entry in $linkValue) { if ($null -ne $entry) { $pendingLinks.Push($entry) } }
                    continue
                }

                foreach ($linkProperty in $linkValue.PSObject.Properties) {
                    if ($linkProperty.Value -is [string] -and $linkProperty.Name -match '(?i)^(url|href|linkurl|targeturl)$') {
                        $found.Add([PSCustomObject]@{ Text = ''; Url = [string]$linkProperty.Value })
                    }
                    elseif ($linkProperty.Value -is [string] -and ([string]$linkProperty.Value).Trim() -match '(?i)^(?:https?:)?//|^/|^~site(?:collection)?/') {
                        $found.Add([PSCustomObject]@{ Text = ''; Url = ([string]$linkProperty.Value).Trim() })
                    }
                    elseif ($null -ne $linkProperty.Value -and $linkProperty.Value -isnot [string]) {
                        $pendingLinks.Push($linkProperty.Value)
                    }
                }
            }
        }

        foreach ($property in $current.PSObject.Properties) {
            if ($property.Name -eq 'innerHTML' -or $property.Name -eq 'sourceItem') { continue }
            if ($null -ne $property.Value -and $property.Value -isnot [string]) { $pending.Push($property.Value) }
        }
    }

    return $found.ToArray()
}

function Get-FieldValueHyperlink {
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
        $hyperlinks = if ($CanvasContent) { Get-CanvasContentHyperlink -CanvasContent $textValue } else { Get-HtmlHyperlink -Html $textValue }
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

function Get-UrlQueryValue {
    param(
        [Parameter(Mandatory = $true)][AllowEmptyString()][string]$Query,
        [Parameter(Mandatory = $true)][string]$Name
    )

    foreach ($pair in ($Query.TrimStart('?') -split '&')) {
        $separator = $pair.IndexOf('=')
        if ($separator -lt 1) { continue }
        if ($pair.Substring(0, $separator) -ieq $Name) {
            return [Uri]::UnescapeDataString(($pair.Substring($separator + 1) -replace '\+', ' '))
        }
    }
    return ''
}

# Old document GUID from Doc.aspx?sourcedoc=, copy-link ?d=<letter><guid>, or ?UniqueId=.
function Get-DocumentGuidFromUrl {
    param([Parameter(Mandatory = $true)][AllowEmptyString()][string]$Url)

    $guidPattern = '([0-9a-f]{8}-?[0-9a-f]{4}-?[0-9a-f]{4}-?[0-9a-f]{4}-?[0-9a-f]{12})'
    $match = [regex]::Match($Url, '(?i)[?&]sourcedoc=(?:%7B|\{)?' + $guidPattern)
    if (-not $match.Success) { $match = [regex]::Match($Url, '(?i)[?&]d=[a-z]([0-9a-f]{32})(?:&|$)') }
    if (-not $match.Success) { $match = [regex]::Match($Url, '(?i)[?&]UniqueId=(?:%7B|\{)?' + $guidPattern) }
    if (-not $match.Success) { return '' }
    return ([Guid]$match.Groups[1].Value).ToString('N')
}

function Get-EmbeddedLinkTarget {
    param([Parameter(Mandatory = $true)][string]$Url)

    $uri = [Uri]$Url
    $path = ([Uri]::UnescapeDataString($uri.AbsolutePath) -replace '(?i)^/:[a-z]:/[rs](?=/)', '').TrimEnd('/')
    $idPath = Get-UrlQueryValue -Query $uri.Query -Name 'id'
    $filePath = ''
    if ($idPath.StartsWith('/')) {
        $filePath = $idPath.TrimEnd('/')
    }
    elseif ($path -notmatch '(?i)/_layouts/|/_vti_bin/|/Forms/[^/]+\.aspx$') {
        $filePath = $path
    }

    return [PSCustomObject]@{
        Origin           = '{0}://{1}' -f $uri.Scheme, $uri.Host
        Path             = $path
        FilePath         = $filePath
        WebPath          = if ($path -match '(?i)^(.*?)/_layouts/') { $Matches[1] } else { '' }
        Query            = $uri.Query
        DocumentGuid     = Get-DocumentGuidFromUrl -Url $Url
        UsesResolvedForm = ($path -match '(?i)/_layouts/15/Doc\.aspx$') -or -not [string]::IsNullOrWhiteSpace($idPath)
    }
}

# Indexes every classic [site]-master.csv by old document GUID and old server-relative path.
function Import-MasterLinkIndex {
    param([Parameter(Mandatory = $true)][string]$Folder)

    $index = [PSCustomObject]@{ ByGuid = @{}; ByPath = @{}; RowCount = 0; FileCount = 0 }
    foreach ($masterFile in @(Get-ChildItem -LiteralPath $Folder -Filter '*-master.csv' -File -ErrorAction SilentlyContinue)) {
        Add-MasterFileToLinkIndex -Index $index -Path $masterFile.FullName
    }
    return $index
}

function Add-MasterFileToLinkIndex {
    param(
        [Parameter(Mandatory = $true)]$Index,
        [Parameter(Mandatory = $true)][string]$Path
    )

    $Index.FileCount++
    $masterName = [System.IO.Path]::GetFileName($Path)
    foreach ($row in @(Import-Csv -LiteralPath $Path)) {
        $entry = [PSCustomObject]@{ Source = $masterName; Row = $row }
        $pathKey = ([string]$row.OldFileRef).Trim().TrimEnd('/').ToLowerInvariant()
        if ($pathKey) { $Index.ByPath[$pathKey] = $entry }
        foreach ($documentUrl in @([string]$row.OldCopyLinkUrl, [string]$row.OldResolvedUrl)) {
            $documentGuid = Get-DocumentGuidFromUrl -Url $documentUrl
            if ($documentGuid) { $Index.ByGuid[$documentGuid] = $entry }
        }
        $Index.RowCount++
    }
}

function ConvertTo-EmbeddedTargetResult {
    param(
        [Parameter(Mandatory = $true)][string]$Found,
        [Parameter(Mandatory = $false)][string]$NewSubsite = '',
        [Parameter(Mandatory = $false)][string]$NewUrl = '',
        [Parameter(Mandatory = $true)][string]$Resolution
    )

    return [PSCustomObject]@{ Found = $Found; NewSubsite = $NewSubsite; NewUrl = $NewUrl; Resolution = $Resolution }
}

function Get-PnPConnectionFromCache {
    param(
        [Parameter(Mandatory = $true)][string]$Url,
        [Parameter(Mandatory = $false)][string]$ClientId,
        [Parameter(Mandatory = $true)][System.Collections.IDictionary]$ConnectionCache
    )

    $cacheKey = $Url.TrimEnd('/')
    if (-not $ConnectionCache.ContainsKey($cacheKey)) {
        try {
            $ConnectionCache[$cacheKey] = Connect-ValidatedPnP -Url $cacheKey -ClientId $ClientId
        }
        catch {
            $ConnectionCache[$cacheKey] = $null
        }
    }
    return $ConnectionCache[$cacheKey]
}

# Used when a link only carries the old GUID (Doc.aspx) and no master row has it yet.
function Get-OldFilePathById {
    param(
        [Parameter(Mandatory = $true)][string]$WebUrl,
        [Parameter(Mandatory = $true)][string]$DocumentGuid,
        [Parameter(Mandatory = $false)][string]$ClientId,
        [Parameter(Mandatory = $true)][System.Collections.IDictionary]$ConnectionCache
    )

    $oldWebUrl = $WebUrl.TrimEnd('/')
    $oldWebConnection = Get-PnPConnectionFromCache -Url $oldWebUrl -ClientId $ClientId -ConnectionCache $ConnectionCache
    if (-not $oldWebConnection) { return '' }

    $restUrl = "$oldWebUrl/_api/web/GetFileById('$(([Guid]$DocumentGuid).ToString('D'))')?`$select=ServerRelativeUrl"
    try {
        $restResult = Invoke-PnPConnectionOperation -Url $oldWebUrl -ClientId $ClientId -Connection $oldWebConnection -Label ('Lookup old document by ID: ' + $DocumentGuid) -MaxAttempts 2 -Operation {
            param($activeConnection)
            Invoke-PnPSPRestMethod -Method Get -Url $restUrl -Connection $activeConnection
        }
        $ConnectionCache[$oldWebUrl] = $restResult.Connection
        $fileInfo = $restResult.Result
        if ($fileInfo.PSObject.Properties['ServerRelativeUrl']) { return [string]$fileInfo.ServerRelativeUrl }
        if ($fileInfo.PSObject.Properties['d']) { return [string]$fileInfo.d.ServerRelativeUrl }
    }
    catch {
        Write-Verbose "Could not resolve document $DocumentGuid in $oldWebUrl : $($_.Exception.Message)"
    }
    return ''
}

function Resolve-EmbeddedTargetUrl {
    param(
        [Parameter(Mandatory = $true)][string]$Url,
        [Parameter(Mandatory = $true)][string]$RootUrl,
        [Parameter(Mandatory = $true)][array]$SiteUrlMap,
        [Parameter(Mandatory = $false)][string]$ClientId,
        [Parameter(Mandatory = $true)][System.Collections.IDictionary]$ConnectionCache,
        [Parameter(Mandatory = $true)]$MasterIndex
    )

    $target = Get-EmbeddedLinkTarget -Url $Url
    $mappingPath = if ($target.FilePath) { $target.FilePath } elseif ($target.WebPath) { $target.WebPath } else { $target.Path }
    $mappedMatch = Get-NewSiteBaseForUrl -Url ($target.Origin + $mappingPath) -SiteUrlMap $SiteUrlMap
    $mappedSubsite = if ($mappedMatch) { $mappedMatch.NewLinkBase } else { '' }
    $mappedUrl = if (-not $mappedMatch) { '' } elseif ($mappedMatch.Suffix) { $mappedMatch.NewLinkBase + (Get-EncodedServerRelativePath -Path $mappedMatch.Suffix) } else { $mappedMatch.NewLinkBase }

    $filePath = $target.FilePath
    $masterEntry = $null
    $matchedBy = ''
    if ($target.DocumentGuid -and $MasterIndex.ByGuid.ContainsKey($target.DocumentGuid)) {
        $masterEntry = $MasterIndex.ByGuid[$target.DocumentGuid]
        $matchedBy = 'document ID'
    }
    if (-not $masterEntry -and $filePath -and $MasterIndex.ByPath.ContainsKey($filePath.ToLowerInvariant())) {
        $masterEntry = $MasterIndex.ByPath[$filePath.ToLowerInvariant()]
        $matchedBy = 'file path'
    }
    if (-not $masterEntry -and -not $filePath -and $target.DocumentGuid) {
        $oldWebPath = if ($target.WebPath) { $target.WebPath } elseif ($mappedMatch) { ([Uri]$mappedMatch.OldBase).AbsolutePath.TrimEnd('/') } else { '' }
        $filePath = Get-OldFilePathById -WebUrl ($target.Origin + $oldWebPath) -DocumentGuid $target.DocumentGuid -ClientId $ClientId -ConnectionCache $ConnectionCache
        if ($filePath -and $MasterIndex.ByPath.ContainsKey($filePath.ToLowerInvariant())) {
            $masterEntry = $MasterIndex.ByPath[$filePath.ToLowerInvariant()]
            $matchedBy = 'document ID (resolved in the old site) and file path'
        }
    }

    $masterNote = ''
    if ($masterEntry) {
        $masterRow = $masterEntry.Row
        $masterNewFileRef = [string]$masterRow.NewFileRef
        $masterMatch = Get-NewSiteBaseForUrl -Url ($RootUrl + [string]$masterRow.OldFileRef) -SiteUrlMap $SiteUrlMap
        $masterSubsite = if ($masterMatch -and $masterNewFileRef.StartsWith(([Uri]$masterMatch.NewLinkBase).AbsolutePath.TrimEnd('/') + '/', [System.StringComparison]::OrdinalIgnoreCase)) {
            $masterMatch.NewLinkBase
        }
        elseif ($masterNewFileRef -match '^(/sites/[^/]+)') {
            $RootUrl + $Matches[1]
        }
        else {
            $mappedSubsite
        }

        if ([string]$masterRow.NewFound -eq 'Yes') {
            $masterNewUrl = if ($target.UsesResolvedForm -and $masterRow.NewResolvedUrl) { [string]$masterRow.NewResolvedUrl } elseif ($masterRow.NewCopyLinkUrl) { [string]$masterRow.NewCopyLinkUrl } else { [string]$masterRow.NewResolvedUrl }
            return ConvertTo-EmbeddedTargetResult -Found 'Yes' -NewSubsite $masterSubsite -NewUrl $masterNewUrl -Resolution "Matched $($masterEntry.Source) by $matchedBy."
        }

        # The master can be older than the migration, so re-check its old path live.
        $filePath = ([string]$masterRow.OldFileRef).TrimEnd('/')
        $masterNote = " $($masterEntry.Source) (matched by $matchedBy) lists it as not found at the new location."
        if ($masterRow.Notes) { $masterNote += " Master notes: $($masterRow.Notes)" }
    }

    if (-not $filePath) {
        if ($target.DocumentGuid) {
            return ConvertTo-EmbeddedTargetResult -Found 'No' -NewSubsite $mappedSubsite -NewUrl $mappedUrl -Resolution "Document ID $($target.DocumentGuid) was not found in any master lookup or in the old site (it may have been deleted)."
        }
        if ($mappedMatch) {
            return ConvertTo-EmbeddedTargetResult -Found 'Mapped' -NewSubsite $mappedSubsite -NewUrl $mappedUrl -Resolution 'Mapped by site base only; the link targets a site page or view, not a specific file.'
        }
        return ConvertTo-EmbeddedTargetResult -Found 'No' -Resolution 'No matching Old Site URL row.'
    }

    $pathMatch = Get-NewSiteBaseForUrl -Url ($target.Origin + $filePath) -SiteUrlMap $SiteUrlMap
    if (-not $pathMatch) {
        return ConvertTo-EmbeddedTargetResult -Found 'No' -Resolution ('No matching Old Site URL row.' + $masterNote)
    }

    $lookupBase = $pathMatch.NewBase
    $lookupPath = ([Uri]$pathMatch.NewBase).AbsolutePath.TrimEnd('/') + $pathMatch.Suffix
    $linkPath = ([Uri]$pathMatch.NewLinkBase).AbsolutePath.TrimEnd('/') + $pathMatch.Suffix
    $linkSubsite = $pathMatch.NewLinkBase
    $locationOverride = Get-NewDocumentLocationOverride -FileRef $filePath
    if ($locationOverride) {
        $lookupBase = $locationOverride.ConnectionBase
        $lookupPath = $locationOverride.NewLibraryPath + $filePath.Substring($locationOverride.OldLibraryPath.Length)
        $linkPath = $lookupPath
        $linkSubsite = $RootUrl + ([Uri]$locationOverride.ConnectionBase).AbsolutePath.TrimEnd('/')
    }
    $expectedUrl = $RootUrl + (Get-EncodedServerRelativePath -Path $linkPath)

    $connection = Get-PnPConnectionFromCache -Url $lookupBase -ClientId $ClientId -ConnectionCache $ConnectionCache
    if (-not $connection) {
        return ConvertTo-EmbeddedTargetResult -Found 'No' -NewSubsite $linkSubsite -NewUrl $expectedUrl -Resolution ("Could not connect to $lookupBase to verify the new location." + $masterNote)
    }

    try {
        $itemResult = Invoke-PnPConnectionOperation -Url $lookupBase -ClientId $ClientId -Connection $connection -Label ('Lookup embedded target: ' + $lookupPath) -MaxAttempts 2 -Operation {
            param($activeConnection)
            Get-PnPFile -Url $lookupPath -AsListItem -Connection $activeConnection
        }
        $ConnectionCache[$lookupBase.TrimEnd('/')] = $itemResult.Connection
        $newUniqueId = $itemResult.Result.FieldValues['UniqueId']
        $extension = Get-NormalizedExtension -FileName ([System.IO.Path]::GetFileName($linkPath))
        return ConvertTo-EmbeddedTargetResult -Found 'Yes' -NewSubsite $linkSubsite -NewUrl (Get-DocumentBrowserUrl -TenantBase $RootUrl -FileRef $linkPath -Extension $extension -UniqueId $newUniqueId) -Resolution ('Live lookup found the file at its mapped new path.' + $masterNote)
    }
    catch {
        Write-Verbose "No file at $lookupPath : $($_.Exception.Message)"
    }

    try {
        $folderResult = Invoke-PnPConnectionOperation -Url $lookupBase -ClientId $ClientId -Connection $connection -Label ('Lookup embedded folder: ' + $lookupPath) -MaxAttempts 2 -Operation {
            param($activeConnection)
            Get-PnPFolder -Url $lookupPath -Connection $activeConnection
        }
        $ConnectionCache[$lookupBase.TrimEnd('/')] = $folderResult.Connection
        return ConvertTo-EmbeddedTargetResult -Found 'Yes' -NewSubsite $linkSubsite -NewUrl $expectedUrl -Resolution 'Live lookup found the folder at its mapped new path.'
    }
    catch {
        Write-Verbose "No folder at $lookupPath : $($_.Exception.Message)"
    }

    return ConvertTo-EmbeddedTargetResult -Found 'No' -NewSubsite $linkSubsite -NewUrl $expectedUrl -Resolution ("Not found at the mapped new path $lookupPath; it may have been renamed, moved, deleted, or not migrated yet." + $masterNote)
}

function Get-NavigationNodeTreeItem {
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
    [CmdletBinding(SupportsShouldProcess = $true)]
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
            if ($PSCmdlet.ShouldProcess($Path, 'Remove file')) {
                Remove-Item -Path $Path -Force
                return
            }
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

    $mergedRowsByKey = Import-ExistingLookupTable -Path $MasterPath -KeyProperty 'OldFileRef'
    foreach ($row in @(Import-Csv -Path $TempPath)) {
        $key = ([string]$row.OldFileRef).Trim().ToLowerInvariant()
        if ([string]::IsNullOrWhiteSpace($key)) { continue }
        $mergedRowsByKey[$key] = $row
    }

    $mergedRows = New-Object System.Collections.Generic.List[Object]
    foreach ($row in $mergedRowsByKey.Values) {
        $mergedRows.Add($row)
    }

    Export-CsvTableWithRetry -Path $MasterPath -Rows $mergedRows.ToArray()
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
        [Parameter(Mandatory = $true)][System.Collections.IDictionary]$ScannedSourceUrls,
        [Parameter(Mandatory = $false)][AllowEmptyCollection()][array]$ScanErrors = @()
    )

    $columns = @('Site', 'Subsite', 'LibraryorListName', 'FolderPath', 'FileName', 'Extension', 'DocumentorItemURL', 'EmbeddedOccurrence', 'EmbeddedLocation', 'EmbeddedURLText', 'EmbeddedURL', 'NewFound', 'NewSubsite', 'NewDocumentorItemURL', 'EmbeddedNewFound', 'EmbeddedNewSubsite', 'EmbeddedNewURL', 'EmbeddedResolution')
    $reportPath = Join-Path -Path $Folder -ChildPath ("$Identifier-EmbeddedLinks.csv")
    $scanTempPath = Join-Path -Path $Folder -ChildPath ("$Identifier-EmbeddedLinks-$RunStamp-scan-temp.csv")
    $mergedTempPath = Join-Path -Path $Folder -ChildPath ("$Identifier-EmbeddedLinks-$RunStamp-merged-temp.csv")

    $currentRows = @($RowsByKey.Values | Where-Object { Test-IsEmbeddedLinkReportRow -Row $_ })
    if ($currentRows.Count -gt 0) {
        Export-CsvTableWithRetry -Path $scanTempPath -Rows $currentRows
    }
    else {
        $headerLine = (@($columns | ForEach-Object { '"' + $_ + '"' }) -join ',')
        Set-Content -Path $scanTempPath -Value $headerLine -Encoding UTF8 -Force
    }

    $mergedRowsByKey = [ordered]@{}
    $retainedSourceKeys = @{}
    if (Test-Path -Path $reportPath) {
        $legacyRowIndex = 0
        foreach ($row in @(Import-Csv -Path $reportPath)) {
            $legacyRowIndex++
            if (-not (Test-IsEmbeddedLinkReportRow -Row $row)) { continue }
            $sourceUrlKey = ([string]$row.DocumentorItemURL).Trim().ToLowerInvariant()
            $sourceKey = $sourceUrlKey + "`n" + [string]$row.LibraryorListName
            if ($ScannedSourceUrls.ContainsKey($sourceKey)) { continue }
            $retainedSourceKeys[$sourceKey] = $true
            $occurrence = [string]$row.EmbeddedOccurrence
            if ([string]::IsNullOrWhiteSpace($occurrence)) { $occurrence = "legacy-$legacyRowIndex" }
            $row | Add-Member -MemberType NoteProperty -Name EmbeddedOccurrence -Value $occurrence -Force
            $rowKey = Get-Sha256Hex -Text ($sourceKey + "`n" + $occurrence + "`n" + [string]$row.EmbeddedURLText + "`n" + [string]$row.EmbeddedURL)
            $mergedRowsByKey[$rowKey] = $row
        }
    }

    foreach ($row in @(Import-Csv -Path $scanTempPath)) {
        $sourceUrlKey = ([string]$row.DocumentorItemURL).Trim().ToLowerInvariant()
        $sourceKey = $sourceUrlKey + "`n" + [string]$row.LibraryorListName
        # A partial scan must not mix new instance numbers with a prior source snapshot.
        if ($retainedSourceKeys.ContainsKey($sourceKey)) { continue }
        $rowKey = Get-Sha256Hex -Text ($sourceKey + "`n" + [string]$row.EmbeddedOccurrence + "`n" + [string]$row.EmbeddedURLText + "`n" + [string]$row.EmbeddedURL)
        $mergedRowsByKey[$rowKey] = $row
    }

    $mergedRows = New-Object System.Collections.Generic.List[Object]
    foreach ($row in $mergedRowsByKey.Values) { $mergedRows.Add(($row | Select-Object -Property $columns)) }
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

    $errorPath = Join-Path $Folder ("$Identifier-EmbeddedLinks-ScanErrors.csv")
    $errorColumns = @('Site', 'Subsite', 'LibraryorListName', 'FolderPath', 'FileName', 'Extension', 'DocumentorItemURL', 'ScanError')
    if ($ScanErrors.Count -gt 0) {
        Export-CsvTableWithRetry -Path $errorPath -Rows @($ScanErrors | Select-Object -Property $errorColumns)
        Write-Warning "$($ScanErrors.Count) embedded-content scan failure(s); see $errorPath. Findings may be incomplete, and prior findings for failed sources are retained."
    }
    else {
        $errorHeader = ($errorColumns | ForEach-Object { '"' + $_ + '"' }) -join ','
        Set-Content -Path $errorPath -Value $errorHeader -Encoding UTF8 -Force
    }

    $staleTimestampedReports = Get-ChildItem -Path $Folder -Filter ("*-{0}-EmbeddedLinks.csv" -f $Identifier) -File -ErrorAction SilentlyContinue |
        Where-Object { $_.FullName -ne $reportPath }
    foreach ($staleReport in $staleTimestampedReports) {
        Remove-FileWithRetry -Path $staleReport.FullName
    }

    return $mergedRows.Count
}

function Test-IsEmbeddedLinkReportRow {
    param([Parameter(Mandatory = $true)]$Row)

    if ([string]::IsNullOrWhiteSpace([string]$Row.EmbeddedURL)) { return $false }
    $extension = ([string]$Row.Extension).Trim().ToLowerInvariant()
    return [string]::IsNullOrWhiteSpace($extension) -or ($ContentScanExtensions -contains $extension)
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
if ([string]::IsNullOrWhiteSpace($SitesEmbeddedLinksFolder)) {
    $SitesEmbeddedLinksFolder = Join-Path -Path $OutputFolder -ChildPath 'Sites-EmbeddedLinks'
}
if ([string]::IsNullOrWhiteSpace($SitesMasterFolder)) {
    $SitesMasterFolder = Join-Path -Path $OutputFolder -ChildPath 'Sites-Masters'
}
if ([string]::IsNullOrWhiteSpace($SitesFixerOutputFolder)) {
    $SitesFixerOutputFolder = Join-Path -Path $OutputFolder -ChildPath 'Sites-FixerFiles'
}
$SitesTempMasterFolder = Join-Path -Path $SitesMasterFolder -ChildPath 'Temp'

New-DirectoryIfMissing -Path $OutputFolder
New-DirectoryIfMissing -Path $MasterFolder
New-DirectoryIfMissing -Path $TempMasterFolder
New-DirectoryIfMissing -Path $FixerOutputFolder
New-DirectoryIfMissing -Path $EmbeddedLinksFolder
New-DirectoryIfMissing -Path $SitesEmbeddedLinksFolder
New-DirectoryIfMissing -Path $SitesMasterFolder
New-DirectoryIfMissing -Path $SitesTempMasterFolder
New-DirectoryIfMissing -Path $SitesFixerOutputFolder
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
    Get-PnPSubWeb -Connection $activeConnection -Recurse -Includes Url, Title
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
        if (-not (Get-ExactNewSiteBaseForUrl -Url $subUrl -SiteUrlMap $candidateMap)) {
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
    if (-not (Get-ExactNewSiteBaseForUrl -Url $selectedSubsiteUrl -SiteUrlMap $siteUrlMap)) {
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
$script:DeepPdfOcr = $false
if ($scanDocumentContent -or $automaticEmbeddedScanSites.Count -gt 0) {
    $script:DeepPdfOcr = if ($PSBoundParameters.ContainsKey('EnableDeepPdfOcr')) {
        [bool]$EnableDeepPdfOcr
    }
    else {
        Get-BooleanPrompt -Message 'Also use Ghostscript and Tesseract OCR on every PDF page to find links in scanned/image content? This can significantly increase run time.' -DefaultYes $false
    }
    if (-not $script:DeepPdfOcr) {
        Write-Warning 'Deep PDF OCR is off: URLs that appear only as images in scanned PDFs will not be found.'
    }
}
if ($automaticEmbeddedScanSites.Count -gt 0) {
    Write-Host ("Automatically enabling embedded-link scanning for {0} selected site(s) missing from the Old/New Site URL CSV; only links under CSV Old Site URL bases will be reported." -f $automaticEmbeddedScanSites.Count)
}
$masterLinkIndex = $null
if ($scanDocumentContent -or $automaticEmbeddedScanSites.Count -gt 0) {
    Initialize-PdfPig
    $masterLinkIndex = Import-MasterLinkIndex -Folder $MasterFolder
    Write-Host ("Embedded-link targets will be resolved first from {0} master row(s) in {1} classic master file(s) in {2}, then by live lookup." -f $masterLinkIndex.RowCount, $masterLinkIndex.FileCount, $MasterFolder)
    if ($masterLinkIndex.FileCount -eq 0) {
        Write-Warning 'No classic master files were found. Run the master scans of the old (root) sites first for the most accurate new URLs.'
    }
}

# --- Enumeration ----------------------------------------------------------
# One master CSV per NEW/destination site (not per selected scope) - a classic
# scope like the tenant root can contain dozens of subsites that are each being
# promoted to their own separate new site collection. Each run writes a temp
# file, then publishes it over the matching [site]-master.csv after the scan.

function Get-ShortSiteIdentifier {
    param(
        [Parameter(Mandatory = $true)]$Subsite,
        [Parameter(Mandatory = $false)][switch]$Legacy
    )

    $scannedUrl = ([string]$Subsite.Url).TrimEnd('/')
    $scannedPath = ([Uri]$scannedUrl).AbsolutePath.TrimEnd('/')
    $pathSegments = @($scannedPath -split '/' | Where-Object { $_ })
    $isSitesCollection = $pathSegments.Count -gt 0 -and $pathSegments[0].Equals('sites', [System.StringComparison]::OrdinalIgnoreCase)
    $identifierSegments = @(
        if ($isSitesCollection) { $pathSegments | Select-Object -Skip 1 }
        else { $pathSegments }
    )

    if ($identifierSegments.Count -eq 0) {
        return Get-SafeFileSlug -Text ([string]$Subsite.Title) -MaxLength 30 -Fallback 'Site'
    }
    if ($identifierSegments.Count -eq 1) {
        return Get-SafeFileSlug -Text $identifierSegments[0] -MaxLength 30 -Fallback 'Site'
    }

    $fingerprint = (Get-Sha256Hex -Text $scannedUrl.ToLowerInvariant()).Substring(0, 8)
    if ($Legacy) {
        $legacyHierarchy = @($identifierSegments | ForEach-Object {
                Get-SafeFileSlug -Text ([string]$_) -MaxLength 14 -Fallback 'Web'
            }) -join '-'
        $legacyHierarchy = Get-SafeFileSlug -Text $legacyHierarchy -MaxLength 38 -Fallback 'Site'
        return "$legacyHierarchy-$fingerprint"
    }

    $hierarchy = @($identifierSegments | ForEach-Object {
            Get-SafeFileSlug -Text ([string]$_) -MaxLength 64 -Fallback 'Web'
        }) -join '-'
    if ($hierarchy.Length -le 100) {
        return $hierarchy
    }
    return $hierarchy.Substring(0, 91).TrimEnd('-') + '-' + $fingerprint
}

# Renames an output file written under an earlier identifier scheme so history carries over.
function Move-LegacySiteOutputFile {
    param(
        [Parameter(Mandatory = $true)][string]$LegacyPath,
        [Parameter(Mandatory = $true)][string]$CurrentPath
    )

    if ($LegacyPath -eq $CurrentPath -or -not (Test-Path -LiteralPath $LegacyPath)) { return }
    if (Test-Path -LiteralPath $CurrentPath) {
        Write-Warning "  Both $LegacyPath and $CurrentPath exist; keeping $CurrentPath. Remove the older file once you have confirmed it is no longer needed."
        return
    }
    Move-Item -LiteralPath $LegacyPath -Destination $CurrentPath
    Write-Host "  Renamed $LegacyPath -> $CurrentPath"
}

$runStamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$newConnectionsByBase = @{}
$embeddedTargetResolutionCache = @{}
$filesScannedTotal = 0
$newFoundCountTotal = 0
$newFoundFromCacheCountTotal = 0
$newNotFoundCountTotal = 0
$extensionCountsTotal = @{}

# Embedded-link findings are keyed by source item and link details; scanned-source keys
# allow the publisher to refresh scanned items while retaining cached unscanned items.
$embeddedLinksBySite = @{}
$embeddedLinksOutputFolderBySite = @{}
$scannedEmbeddedSourcesBySite = @{}
$embeddedScanErrorsBySite = @{}

foreach ($subsite in $selectedSubsites) {
    $subsiteUrl = ([string]$subsite.Url).TrimEnd('/')
    Write-Host "`nScanning $subsiteUrl ..."

    $siteMatch = Get-ExactNewSiteBaseForUrl -Url $subsiteUrl -SiteUrlMap $siteUrlMap
    $isNewStructureWeb = ([Uri]$subsiteUrl).AbsolutePath -match '(?i)^/sites(?:/|$)'
    $scanThisSiteContent = $scanDocumentContent -or $automaticEmbeddedScanSites.ContainsKey($subsiteUrl.ToLowerInvariant())
    $scanOldBaseOnly = $isNewStructureWeb -or ((-not $siteMatch) -and (-not $scanDocumentContent))
    if ((-not $siteMatch) -and (-not $isNewStructureWeb)) {
        Write-Warning "  No matching row in the Old/New Site URL CSV for $subsiteUrl - old URLs will still be recorded, but new URLs cannot be looked up."
    }
    elseif ($isNewStructureWeb) {
        Write-Host '  New-structure web: reporting only links matched to the Old Site URL bases in the selected CSV.'
    }

    $siteIdentifier = Get-ShortSiteIdentifier -Subsite $subsite
    $legacySiteIdentifier = Get-ShortSiteIdentifier -Subsite $subsite -Legacy
    $siteMasterFolder = if ($isNewStructureWeb) { $SitesMasterFolder } else { $MasterFolder }
    $siteTempMasterFolder = if ($isNewStructureWeb) { $SitesTempMasterFolder } else { $TempMasterFolder }
    $siteFixerFolder = if ($isNewStructureWeb) { $SitesFixerOutputFolder } else { $FixerOutputFolder }
    $embeddedOutputFolder = if ($isNewStructureWeb) { $SitesEmbeddedLinksFolder } else { $EmbeddedLinksFolder }
    $embeddedReportKey = $embeddedOutputFolder + '|' + $siteIdentifier
    $masterPath = Get-MasterLookupTablePath -Folder $siteMasterFolder -Identifier $siteIdentifier
    $outputPath = Get-TemporaryLookupTablePath -Folder $siteTempMasterFolder -Identifier $siteIdentifier -RunStamp $runStamp
    $fixerPath = Get-FixerLookupTablePath -Folder $siteFixerFolder -Identifier $siteIdentifier
    if ($legacySiteIdentifier -ne $siteIdentifier) {
        Move-LegacySiteOutputFile -LegacyPath (Get-MasterLookupTablePath -Folder $MasterFolder -Identifier $legacySiteIdentifier) -CurrentPath $masterPath
        Move-LegacySiteOutputFile -LegacyPath (Get-FixerLookupTablePath -Folder $FixerOutputFolder -Identifier $legacySiteIdentifier) -CurrentPath $fixerPath
        Move-LegacySiteOutputFile -LegacyPath (Join-Path $embeddedOutputFolder ("{0}-EmbeddedLinks.csv" -f $legacySiteIdentifier)) -CurrentPath (Join-Path $embeddedOutputFolder ("{0}-EmbeddedLinks.csv" -f $siteIdentifier))
    }
    $mergedRowsByKey = if (Test-Path -Path $masterPath) { Import-ExistingLookupTable -Path $masterPath -KeyProperty 'OldFileRef' } else { [ordered]@{} }
    $counterCheckpointPath = Get-CounterCheckpointPath -StateRoot $StateRootFolder -ScopeUrl $subsiteUrl -SchemaVersion $ScriptSchemaVersion
    if (Test-Path -Path $masterPath) {
        Write-Host ("  Writing temp scan {0} (using {1} existing row(s) from master {2} to skip already-confirmed lookups)" -f $outputPath, $mergedRowsByKey.Count, $masterPath)
    }
    else {
        Write-Host ("  Writing temp scan {0} (no master file found yet for this site; one will be created at {1})" -f $outputPath, $masterPath)
    }

    if ($scanThisSiteContent) {
        if (-not $embeddedLinksBySite.ContainsKey($embeddedReportKey)) { $embeddedLinksBySite[$embeddedReportKey] = [ordered]@{} }
        if (-not $scannedEmbeddedSourcesBySite.ContainsKey($embeddedReportKey)) { $scannedEmbeddedSourcesBySite[$embeddedReportKey] = @{} }
        if (-not $embeddedScanErrorsBySite.ContainsKey($embeddedReportKey)) { $embeddedScanErrorsBySite[$embeddedReportKey] = New-Object System.Collections.Generic.List[Object] }
        $embeddedLinksOutputFolderBySite[$embeddedReportKey] = [PSCustomObject]@{ Folder = $embeddedOutputFolder; Identifier = $siteIdentifier }

        if ($isNewStructureWeb) {
            $legacyReportPath = Join-Path $EmbeddedLinksFolder ("sites-{0}-EmbeddedLinks.csv" -f $siteIdentifier)
            $canonicalReportPath = Join-Path $embeddedOutputFolder ("{0}-EmbeddedLinks.csv" -f $siteIdentifier)
            if (Test-Path -LiteralPath $legacyReportPath) {
                if (-not (Test-Path -LiteralPath $canonicalReportPath) -or
                    (Get-Item -LiteralPath $legacyReportPath).LastWriteTimeUtc -gt (Get-Item -LiteralPath $canonicalReportPath).LastWriteTimeUtc) {
                    Copy-CsvFileWithRetry -SourcePath $legacyReportPath -DestinationPath $canonicalReportPath
                }
                Remove-FileWithRetry -Path $legacyReportPath
            }
        }
    }

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

    foreach ($library in $libraries) {
        $librarySupportsFileLookup = Test-IsFileLookupList -List $library
        $libraryKind = if ($librarySupportsFileLookup) { 'file/page library' } else { 'site contents list' }
        Write-Host ("  {0}: {1} ({2} items)" -f $libraryKind, $library.Title, $library.ItemCount)

        $contentFieldNames = @()
        $contentFieldScanSucceeded = $false
        $contentFieldScanError = ''
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
                $contentFieldScanError = "Could not read text/URL fields from $($library.Title): $($_.Exception.Message)"
                Write-Warning "  $contentFieldScanError"
            }
        }

        $fields = @(@('FileRef', 'FileLeafRef', 'FileDirRef', 'UniqueId', 'FSObjType', 'Modified', 'Title', 'Attachments') + $contentFieldNames | Select-Object -Unique)
        $viewFieldsXml = @($fields | ForEach-Object {
                "<FieldRef Name='{0}'/>" -f [System.Security.SecurityElement]::Escape([string]$_)
            }) -join ''
        $recursiveItemsQuery = "<View Scope='RecursiveAll'><ViewFields>$viewFieldsXml</ViewFields><RowLimit Paged='TRUE'>500</RowLimit></View>"
        $itemsResult = Invoke-PnPConnectionOperation -Url $subsiteUrl -ClientId $ClientId -Connection $subConn -Label ('Get items: ' + $library.Title) -MaxAttempts 3 -Operation {
            param($activeConnection)
            Get-PnPListItem -List $library.Title -Connection $activeConnection -Query $recursiveItemsQuery -PageSize 500
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

            if ($scanThisSiteContent -and (-not $librarySupportsFileLookup -or $ContentScanExtensions -contains $ext)) {
                $sourceHyperlinks = New-Object System.Collections.Generic.List[Object]
                $sourceScanErrors = New-Object System.Collections.Generic.List[string]
                $sourceContentScanned = $false
                $fileContentScanSucceeded = $false
                $attachmentContentScanSucceeded = $true

                if (-not $contentFieldScanSucceeded) {
                    $sourceScanErrors.Add($contentFieldScanError)
                }

                $pageContentFields = @('CanvasContent1', 'WikiField', 'PublishingPageContent', 'LayoutWebpartsContent')
                $documentContentLinkCount = 0
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
                            $pdfExtraction = Get-PdfHyperlink -Bytes $contentStream.ToArray()
                            foreach ($link in @($pdfExtraction.Links)) { $sourceHyperlinks.Add($link) }
                            foreach ($pdfError in @($pdfExtraction.Errors)) { $sourceScanErrors.Add([string]$pdfError) }
                            $fileContentScanSucceeded = ($pdfExtraction.Errors.Count -eq 0)
                        }
                        elseif ($ext -in @('.htm', '.html', '.aspx')) {
                            $htmlContent = [System.Text.Encoding]::UTF8.GetString($contentStream.ToArray())
                            foreach ($link in @(Get-HtmlHyperlink -Html $htmlContent)) { $sourceHyperlinks.Add($link) }
                        }
                        elseif ($OpenXmlContentExtensions -contains $ext) {
                            foreach ($link in @(Get-OfficeDocumentHyperlink -Stream $contentStream)) { $sourceHyperlinks.Add($link) }
                        }
                        elseif ($ext -in @('.txt', '.csv')) {
                            foreach ($link in @(Get-PlainTextHyperlink -Stream $contentStream)) { $sourceHyperlinks.Add($link) }
                        }
                        else {
                            foreach ($link in @(Get-TikaDocumentHyperlink -Bytes $contentStream.ToArray() -Extension $ext)) { $sourceHyperlinks.Add($link) }
                        }
                        if ($ext -ne '.pdf') { $fileContentScanSucceeded = $true }
                        if ($ext -notin @('.htm', '.html', '.aspx')) { $documentContentLinkCount = $sourceHyperlinks.Count }
                    }
                    catch {
                        $sourceScanErrors.Add("Could not scan file content: $($_.Exception.Message)")
                        Write-Warning "  Could not scan content of $fileRef : $($_.Exception.Message)"
                    }
                    finally {
                        if ($contentStream) { $contentStream.Dispose() }
                    }
                }

                $hasListAttachments = (-not $librarySupportsFileLookup) -and [bool]$item.FieldValues['Attachments']
                if ($hasListAttachments) {
                    try {
                        $attachmentListTitle = $library.Title
                        $attachmentItemId = $item.Id
                        $attachmentResult = Invoke-PnPConnectionOperation -Url $subsiteUrl -ClientId $ClientId -Connection $subConn -Label ('Get attachments: ' + $library.Title + ' #' + $item.Id) -MaxAttempts 2 -Operation {
                            param($activeConnection)
                            $activeItem = Get-PnPListItem -List $attachmentListTitle -Id $attachmentItemId -Fields @('Attachments') -Connection $activeConnection
                            Get-PnPProperty -ClientObject $activeItem -Property AttachmentFiles -Connection $activeConnection
                        }
                        $subConn = $attachmentResult.Connection
                        foreach ($attachment in @($attachmentResult.Result)) {
                            $attachmentUrl = [string]$attachment.ServerRelativeUrl
                            $attachmentName = if ($attachment.FileName) { [string]$attachment.FileName } else { ($attachmentUrl -split '/')[-1] }
                            $attachmentExtension = Get-NormalizedExtension -FileName $attachmentName
                            if ($ContentScanExtensions -notcontains $attachmentExtension) {
                                continue
                            }

                            $attachmentStream = $null
                            try {
                                $attachmentContentResult = Invoke-PnPConnectionOperation -Url $subsiteUrl -ClientId $ClientId -Connection $subConn -Label ('Download attachment: ' + $attachmentUrl) -MaxAttempts 2 -Operation {
                                    param($activeConnection)
                                    Get-PnPFile -Url $attachmentUrl -AsMemoryStream -Connection $activeConnection
                                }
                                $subConn = $attachmentContentResult.Connection
                                $attachmentStream = $attachmentContentResult.Result
                                $attachmentFolder = $attachmentUrl.Substring(0, $attachmentUrl.LastIndexOf('/'))
                                if ($attachmentExtension -eq '.pdf') {
                                    $pdfAttachmentExtraction = Get-PdfHyperlink -Bytes $attachmentStream.ToArray()
                                    $attachmentLinks = @($pdfAttachmentExtraction.Links)
                                    foreach ($pdfError in @($pdfAttachmentExtraction.Errors)) {
                                        $attachmentContentScanSucceeded = $false
                                        $sourceScanErrors.Add("Attachment '$attachmentName': $pdfError")
                                    }
                                }
                                elseif ($attachmentExtension -in @('.htm', '.html', '.aspx')) {
                                    $attachmentHtml = [System.Text.Encoding]::UTF8.GetString($attachmentStream.ToArray())
                                    $attachmentLinks = @(Get-HtmlHyperlink -Html $attachmentHtml)
                                }
                                elseif ($OpenXmlContentExtensions -contains $attachmentExtension) {
                                    $attachmentLinks = @(Get-OfficeDocumentHyperlink -Stream $attachmentStream)
                                }
                                elseif ($attachmentExtension -in @('.txt', '.csv')) {
                                    $attachmentLinks = @(Get-PlainTextHyperlink -Stream $attachmentStream)
                                }
                                else {
                                    $attachmentLinks = @(Get-TikaDocumentHyperlink -Bytes $attachmentStream.ToArray() -Extension $attachmentExtension)
                                }

                                foreach ($attachmentLink in $attachmentLinks) {
                                    $attachmentLinkUrl = [string]$attachmentLink.Url
                                    try {
                                        $attachmentLinkUri = [Uri]$attachmentLinkUrl
                                        if (-not $attachmentLinkUri.IsAbsoluteUri) {
                                            $attachmentLinkUri = [Uri]::new([Uri]($RootUrl + $attachmentFolder.TrimEnd('/') + '/'), $attachmentLinkUri)
                                        }
                                        $attachmentLinkUrl = $attachmentLinkUri.AbsoluteUri
                                    }
                                    catch {
                                        $attachmentContentScanSucceeded = $false
                                        $sourceScanErrors.Add("Attachment '$attachmentName' contains an invalid link target '$attachmentLinkUrl'.")
                                        continue
                                    }
                                    $attachmentLocation = "Attachment: $attachmentName"
                                    if ($attachmentLink.PSObject.Properties['Location']) { $attachmentLocation += ' | ' + [string]$attachmentLink.Location }
                                    $sourceHyperlinks.Add([PSCustomObject]@{ Text = [string]$attachmentLink.Text; Url = $attachmentLinkUrl; Location = $attachmentLocation })
                                }
                            }
                            catch {
                                $attachmentContentScanSucceeded = $false
                                $sourceScanErrors.Add("Could not scan list attachment '$attachmentName': $($_.Exception.Message)")
                                Write-Warning "  Could not scan attachment $attachmentName : $($_.Exception.Message)"
                            }
                            finally {
                                if ($attachmentStream) { $attachmentStream.Dispose() }
                            }
                        }
                    }
                    catch {
                        $attachmentContentScanSucceeded = $false
                        $sourceScanErrors.Add("Could not enumerate list attachments: $($_.Exception.Message)")
                        Write-Warning "  Could not enumerate attachments for $($library.Title) item $($item.Id): $($_.Exception.Message)"
                    }
                }

                foreach ($fieldName in $contentFieldNames) {
                    $isPageCanvasField = ($library.Title -eq 'Site Pages' -and $pageContentFields -contains $fieldName)
                    $fieldLinks = if ($isPageCanvasField) {
                        Get-CanvasContentHyperlink -CanvasContent ([string]$item.FieldValues[$fieldName])
                    }
                    else {
                        Get-FieldValueHyperlink -Value $item.FieldValues[$fieldName]
                    }
                    foreach ($link in @($fieldLinks)) {
                        $link | Add-Member -NotePropertyName Location -NotePropertyValue ("Field: $fieldName") -Force
                        $sourceHyperlinks.Add($link)
                    }
                }
                $requiresFileContentScan = $librarySupportsFileLookup
                if ($contentFieldScanSucceeded -and (-not $requiresFileContentScan -or $fileContentScanSucceeded) -and $attachmentContentScanSucceeded) {
                    $sourceContentScanned = $true
                }

                $sourceDocumentKey = $oldCopyLinkUrl.Trim().ToLowerInvariant()

                $embeddedSourceNewFound = $newFound
                $embeddedSourceNewDocumentUrl = if ($newFound -eq 'Yes') { $newCopyLinkUrl } else { '' }
                $embeddedSourceNewSubsite = $newSubsiteForItem
                if ($isNewStructureWeb) {
                    $embeddedSourceNewFound = 'N/A (already in new site)'
                    $embeddedSourceNewDocumentUrl = $oldCopyLinkUrl
                    $embeddedSourceNewSubsite = $subsiteUrl
                }
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

                $sourceOccurrence = 0
                $sourceLinkIndex = 0
                foreach ($link in $sourceHyperlinks) {
                    $sourceLinkIndex++
                    $embeddedUrl = [string]$link.Url
                    if ([string]::IsNullOrWhiteSpace($embeddedUrl)) { continue }
                    $isRootRelativeDocumentLink = ($sourceLinkIndex -le $documentContentLinkCount) -and $embeddedUrl.StartsWith('/') -and -not $embeddedUrl.StartsWith('//')
                    try {
                        $embeddedUri = [Uri]$embeddedUrl
                        if (-not $embeddedUri.IsAbsoluteUri) {
                            $sourceFolderUrl = $RootUrl + $fileDirRef.TrimEnd('/') + '/'
                            $embeddedUri = [Uri]::new([Uri]$sourceFolderUrl, $embeddedUri)
                        }
                        $embeddedUrl = $embeddedUri.AbsoluteUri
                    }
                    catch {
                        $sourceScanErrors.Add("Embedded link target could not be normalized: $($_.Exception.Message)")
                        continue
                    }
                    if (-not (Test-IsInternalTenantUrl -Url $embeddedUrl -RootUrl $RootUrl)) { continue }
                    if ($scanOldBaseOnly -and -not (Test-IsUrlWithinOldSiteBase -Url $embeddedUrl -SiteUrlMap $siteUrlMap)) { continue }
                    # Host-less "/x" links inside files only hit the tenant-root row by accident.
                    if ($isRootRelativeDocumentLink) {
                        $rootRelativeMatch = Get-NewSiteBaseForUrl -Url ($RootUrl + [Uri]::UnescapeDataString(([Uri]$embeddedUrl).AbsolutePath)) -SiteUrlMap $siteUrlMap
                        if (-not $rootRelativeMatch -or $rootRelativeMatch.OldBase.TrimEnd('/') -ieq $RootUrl) { continue }
                    }

                    $embeddedTargetKey = $embeddedUrl.Trim().ToLowerInvariant()
                    if (-not $embeddedTargetResolutionCache.ContainsKey($embeddedTargetKey)) {
                        $embeddedTargetResolutionCache[$embeddedTargetKey] = Resolve-EmbeddedTargetUrl -Url $embeddedUrl -RootUrl $RootUrl -SiteUrlMap $siteUrlMap -ClientId $ClientId -ConnectionCache $newConnectionsByBase -MasterIndex $masterLinkIndex
                    }
                    $embeddedTarget = $embeddedTargetResolutionCache[$embeddedTargetKey]
                    $sourceOccurrence++

                    $embeddedRow = [PSCustomObject][ordered]@{
                        Site                 = $selectedScopeUrl
                        Subsite              = $subsiteUrl
                        LibraryorListName    = $library.Title
                        FolderPath           = $fileDirRef
                        FileName             = if ($librarySupportsFileLookup) { $fileLeaf } elseif ($item.FieldValues['Title']) { [string]$item.FieldValues['Title'] } else { $fileLeaf }
                        Extension            = if ($librarySupportsFileLookup) { $ext } else { '' }
                        DocumentorItemURL    = $oldCopyLinkUrl
                        EmbeddedOccurrence   = $sourceOccurrence
                        EmbeddedLocation     = if ($link.PSObject.Properties['Location']) { [string]$link.Location } else { 'File content' }
                        EmbeddedURLText      = [string]$link.Text
                        EmbeddedURL          = $embeddedUrl
                        NewFound             = $embeddedSourceNewFound
                        NewSubsite           = $embeddedSourceNewSubsite
                        NewDocumentorItemURL = $embeddedSourceNewDocumentUrl
                        EmbeddedNewFound     = $embeddedTarget.Found
                        EmbeddedNewSubsite   = $embeddedTarget.NewSubsite
                        EmbeddedNewURL       = $embeddedTarget.NewUrl
                        EmbeddedResolution   = $embeddedTarget.Resolution
                    }
                    $embeddedKey = Get-Sha256Hex -Text ($sourceDocumentKey + "`n" + $library.Title + "`n" + $sourceOccurrence + "`n" + $embeddedRow.EmbeddedURLText + "`n" + $embeddedUrl)
                    $embeddedLinksBySite[$embeddedReportKey][$embeddedKey] = $embeddedRow
                }

                if ($sourceScanErrors.Count -gt 0) {
                    $scanErrorRow = [PSCustomObject][ordered]@{
                        Site                 = $selectedScopeUrl
                        Subsite              = $subsiteUrl
                        LibraryorListName    = $library.Title
                        FolderPath           = $fileDirRef
                        FileName             = if ($librarySupportsFileLookup) { $fileLeaf } elseif ($item.FieldValues['Title']) { [string]$item.FieldValues['Title'] } else { $fileLeaf }
                        Extension            = if ($librarySupportsFileLookup) { $ext } else { '' }
                        DocumentorItemURL    = $oldCopyLinkUrl
                        ScanError            = @($sourceScanErrors | Select-Object -Unique) -join ' '
                    }
                    $embeddedScanErrorsBySite[$embeddedReportKey].Add($scanErrorRow)
                }
                elseif ($sourceContentScanned) {
                    $scannedEmbeddedSourcesBySite[$embeddedReportKey][($sourceDocumentKey + "`n" + $library.Title)] = $true
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
                $scannedEmbeddedSourcesBySite[$embeddedReportKey][$navigationSourceKey] = $true
            }
            catch {
                $navigationError = "Could not read $navigationLocation navigation for $subsiteUrl`: $($_.Exception.Message)"
                Write-Warning "  $navigationError"
                $navigationErrorRow = [PSCustomObject][ordered]@{
                    Site                 = $selectedScopeUrl
                    Subsite              = $subsiteUrl
                    LibraryorListName    = $navigationContainerName
                    FolderPath           = ''
                    FileName             = ''
                    Extension            = ''
                    DocumentorItemURL    = $subsiteUrl
                    ScanError            = $navigationError
                }
                $embeddedScanErrorsBySite[$embeddedReportKey].Add($navigationErrorRow)
                continue
            }

            $sourceNavigationNodes = @(Get-NavigationNodeTreeItem -Nodes @($navigationResult.Result))
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
                        foreach ($destinationNode in @(Get-NavigationNodeTreeItem -Nodes @($destinationNavigationResult.Result))) {
                            $destinationTitle = [string]$destinationNode.Title
                            if (-not [string]::IsNullOrWhiteSpace($destinationTitle)) { $destinationNodesByTitle[$destinationTitle] = $destinationNode }
                        }
                    }
                    catch {
                        Write-Verbose "Could not read destination $navigationLocation navigation at $newSiteScanBaseAbsolute : $($_.Exception.Message)"
                    }
                }
            }

            $navigationOccurrence = 0
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
                    $scannedEmbeddedSourcesBySite[$embeddedReportKey].Remove($navigationSourceKey)
                    $embeddedScanErrorsBySite[$embeddedReportKey].Add([PSCustomObject]@{
                        Site = $selectedScopeUrl; Subsite = $subsiteUrl; LibraryorListName = $navigationContainerName
                        DocumentorItemURL = $subsiteUrl; ScanError = "Invalid navigation target '$embeddedUrl': $($_.Exception.Message)"
                    })
                    continue
                }
                if (-not (Test-IsInternalTenantUrl -Url $embeddedUrl -RootUrl $RootUrl)) { continue }
                if ($scanOldBaseOnly -and -not (Test-IsUrlWithinOldSiteBase -Url $embeddedUrl -SiteUrlMap $siteUrlMap)) { continue }

                $embeddedTargetKey = $embeddedUrl.Trim().ToLowerInvariant()
                if (-not $embeddedTargetResolutionCache.ContainsKey($embeddedTargetKey)) {
                    $embeddedTargetResolutionCache[$embeddedTargetKey] = Resolve-EmbeddedTargetUrl -Url $embeddedUrl -RootUrl $RootUrl -SiteUrlMap $siteUrlMap -ClientId $ClientId -ConnectionCache $newConnectionsByBase -MasterIndex $masterLinkIndex
                }
                $embeddedTarget = $embeddedTargetResolutionCache[$embeddedTargetKey]
                $navigationOccurrence++

                $nodeTitle = [string]$navigationNode.Title
                $destinationNode = if ($destinationNodesByTitle.ContainsKey($nodeTitle)) { $destinationNodesByTitle[$nodeTitle] } else { $null }
                $embeddedRow = [PSCustomObject][ordered]@{
                    Site                 = $selectedScopeUrl
                    Subsite              = $subsiteUrl
                    LibraryorListName    = $navigationContainerName
                    FolderPath           = ''
                    FileName             = $nodeTitle
                    Extension            = ''
                    DocumentorItemURL    = $subsiteUrl
                    EmbeddedOccurrence   = $navigationOccurrence
                    EmbeddedLocation     = "$navigationLocation | node $($navigationNode.Id)"
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
                $rowKey = Get-Sha256Hex -Text ($subsiteUrl.ToLowerInvariant() + "`n" + $navigationContainerName + "`n" + $navigationOccurrence + "`n" + $nodeTitle.ToLowerInvariant() + "`n" + $embeddedUrl.ToLowerInvariant())
                $embeddedLinksBySite[$embeddedReportKey][$rowKey] = $embeddedRow
            }
        }
    }

    if ($filesScanned -eq 0) {
        Write-Warning "  No documents were found in $subsiteUrl."
    }
    else {
        Publish-MasterLookupTable -TempPath $outputPath -MasterPath $masterPath
        $fixerRowsWritten = Export-FixerLookupTableFromMaster -MasterPath $masterPath -FixerPath $fixerPath -RequiredExtensions $FixerRequiredExtensions
        if ($masterLinkIndex -and -not $isNewStructureWeb) {
            Add-MasterFileToLinkIndex -Index $masterLinkIndex -Path $masterPath
            $embeddedTargetResolutionCache.Clear()
        }

        Write-Host ("  Wrote $filesScanned row(s) to master $masterPath")
        Write-Host ("  Wrote {0} filtered Fixer row(s) to {1}" -f $fixerRowsWritten, $fixerPath)
        Write-Host ("  New document found: {0} ({1} via cached prior confirmation)   Not found: {2}" -f $newFoundCount, $newFoundFromCacheCount, $newNotFoundCount)
        if ($notFoundCountThisSite -gt 0) {
            Write-Host ("  {0} document(s) could not be located at their expected new path - see Notes column." -f $notFoundCountThisSite)
        }
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
    Write-Warning "`nNo documents were found in the selected subsites. No CSVs were written."
}
else {
    Write-Host ("`nAll sites combined: refreshed {0} master row(s) across {1} site CSV(s)." -f $filesScannedTotal, $selectedSubsites.Count)

    Write-Host "`nExtension breakdown:"
    foreach ($extKey in @($extensionCountsTotal.Keys | Sort-Object)) {
        Write-Host ("  {0,-10} {1}" -f $extKey, $extensionCountsTotal[$extKey])
    }

    Write-Host ("`nNew document found: {0} ({1} via cached prior confirmation, no live lookup)   Not found: {2}" -f $newFoundCountTotal, $newFoundFromCacheCountTotal, $newNotFoundCountTotal)
}

if ($scanDocumentContent -or $automaticEmbeddedScanSites.Count -gt 0) {
    foreach ($embeddedReportKey in @($embeddedLinksBySite.Keys | Sort-Object)) {
        $embeddedReport = $embeddedLinksOutputFolderBySite[$embeddedReportKey]
        $embeddedOutputPath = Join-Path -Path $embeddedReport.Folder -ChildPath ("$($embeddedReport.Identifier)-EmbeddedLinks.csv")
        $embeddedRowsWritten = Publish-EmbeddedLinksLookupTable -Folder $embeddedReport.Folder -Identifier $embeddedReport.Identifier -RunStamp $runStamp -RowsByKey $embeddedLinksBySite[$embeddedReportKey] -ScannedSourceUrls $scannedEmbeddedSourcesBySite[$embeddedReportKey] -ScanErrors $embeddedScanErrorsBySite[$embeddedReportKey].ToArray()
        Write-Host "`nWrote $embeddedRowsWritten embedded-link row(s): $embeddedOutputPath"
    }
}
