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

[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost', '', Justification = 'Interactive script intentionally writes status and prompts to the host.')]
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
    [string]$DocumentsOutputFolder = "C:\Users\!bposey\OneDrive - Womens Healthcare Associates LLC\Gary Henderson's files - Project Files_bposey\SharePoint-Migration-Scripts\FixItList-OutputFiles\Documents",

    [Parameter(Mandatory = $false)]
    [string]$OtherOutputFolder = "C:\Users\!bposey\OneDrive - Womens Healthcare Associates LLC\Gary Henderson's files - Project Files_bposey\SharePoint-Migration-Scripts\FixItList-OutputFiles\NonDocuments",

    [Parameter(Mandatory = $false)]
    [string]$LogsOutputFolder = "C:\Users\!bposey\OneDrive - Womens Healthcare Associates LLC\Gary Henderson's files - Project Files_bposey\SharePoint-Migration-Scripts\FixItList-OutputFiles\Logs"
,
    [Parameter(Mandatory = $false)]
    [string]$StateRootFolder = "C:\Users\!bposey\OneDrive - Womens Healthcare Associates LLC\Gary Henderson's files - Project Files_bposey\SharePoint-Migration-Scripts\Checkpoints-and-Caches"
,
    [Parameter(Mandatory = $false)]
    [switch]$UsePersistedExternalUrlCache
,
    [Parameter(Mandatory = $false)]
    [bool]$UsePersistedInventoryCache = $true
)

$ErrorActionPreference = "Stop"

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

Set-ScriptWindowTitle -ScriptName 'Get-FixItListURLsv5.ps1'

$DocumentScannableExtensions = @('.docx', '.xlsx', '.pptx', '.pdf', '.doc', '.xls', '.ppt', '.txt', '.csv', '.rtf')
$OtherFileScannableExtensions = @('.htm', '.aspx')
$OpenXmlExtensions = @('.docx', '.xlsx', '.pptx')
$AlwaysSkipExtensions = @('.jpg', '.png', '.vsdx', '.one')
$SkipListTitles = @('Form Templates', 'Preservation Hold Library')

$SkipSiteUrls = @(
    'https://whallc.sharepoint.com/sites/zzBackup',
    'https://whallc.sharepoint.com/sites/zzCentralParkBackup',
    'https://whallc.sharepoint.com/sites/zzRootBackup',
    'https://whallc.sharepoint.com/sites/zzzTestGroup',
    'https://whallc.sharepoint.com/sites/zzzTestGroup2'
)

$DocLinkPrefixByExt = @{
    '.docx' = 'w'; '.doc' = 'w'
    '.xlsx' = 'x'; '.xls' = 'x'
    '.pptx' = 'p'; '.ppt' = 'p'
    '.pdf'  = 'b'
}

$PdfNoiseHosts = @(
    'ns.adobe.com', 'www.w3.org', 'w3.org', 'purl.org',
    'www.iec.ch', 'iec.ch', 'www.color.org', 'www.npes.org', 'www.aiim.org',
    'www.activepdf.com', 'activepdf.com'
)

$ScriptSchemaVersion = 'v5.0-v3scaffold'
$ExternalUrlStatusCacheTtlDays = 30
$InventoryCacheSchemaVersion = 'v1'

function Install-RequiredModule {
    param([Parameter(Mandatory = $true)][string]$Name)

    if (-not (Get-Module -ListAvailable -Name $Name)) {
        Write-Host "Module $Name was not found. Installing for current user..."
        Install-Module -Name $Name -Scope CurrentUser -Force -AllowClobber
    }

    Import-Module $Name -ErrorAction Stop
}

function Test-ConnectedAccountPermissionSet {
    param(
        [Parameter(Mandatory = $true)]$Connection,
        [Parameter(Mandatory = $true)][string]$Url,
        [Parameter(Mandatory = $false)][switch]$AdminConnection
    )

    $issues = New-Object System.Collections.Generic.List[string]
    $requiredScopes = @('Sites.Read.All', 'User.Read', 'Directory.Read.All')

    try {
        if ($AdminConnection) {
            Get-PnPTenantSite -Connection $Connection -Detailed | Select-Object -First 1 | Out-Null
        }
        else {
            Get-PnPWeb -Connection $Connection -Includes Url | Out-Null
        }
    }
    catch {
        $issues.Add(("SharePoint access validation failed for {0}: {1}" -f $Url, $_.Exception.Message))
    }

    $graphValidated = $false
    $graphMessage = ''
    try {
        if (-not (Get-Module -ListAvailable -Name 'Microsoft.Graph.Authentication')) {
            Write-Host 'Installing Microsoft.Graph.Authentication to validate Graph permissions for the signed-in account...'
            Install-Module -Name 'Microsoft.Graph.Authentication' -Scope CurrentUser -Force -AllowClobber
        }

        Import-Module Microsoft.Graph.Authentication -ErrorAction Stop

        if (-not (Get-Command -Name 'Invoke-MgGraphRequest' -ErrorAction SilentlyContinue)) {
            throw 'Microsoft Graph PowerShell module is not available after import.'
        }

        if (-not (Get-MgContext -ErrorAction SilentlyContinue)) {
            Connect-MgGraph -Scopes $requiredScopes -NoWelcome -ErrorAction Stop
        }

        $null = Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/organization?$select=id,displayName'
        $graphValidated = $true
    }
    catch {
        $graphMessage = $_.Exception.Message
        $issues.Add("Graph access validation failed: $graphMessage")
    }

    if ($issues.Count -gt 0) {
        $permissionException = [System.InvalidOperationException]::new(
            "Connected account does not have the required scope/permissions for this script. Missing or insufficient permissions: $($issues -join '; ')"
        )
        $permissionException.Data['RequiredScopes'] = $requiredScopes
        $permissionException.Data['Issues'] = @($issues)
        throw $permissionException
    }

    if (-not $graphValidated) {
        Write-Warning "Graph permission check could not be completed for the signed-in account. Continuing with SharePoint validation only. $graphMessage"
    }

    return [PSCustomObject]@{
        SharePointValidated = $true
        GraphValidated = $graphValidated
        GraphMessage = $graphMessage
        RequiredScopes = $requiredScopes
        Issues = @()
    }
}

function Show-PermissionReauthPrompt {
    param(
        [Parameter(Mandatory = $true)][string]$AccountName,
        [Parameter(Mandatory = $true)][string]$Message,
        [Parameter(Mandatory = $true)][string[]]$RequiredScopes
    )

    Write-Host ""
    Write-Host "The current account '$AccountName' does not have the required permissions for this script."
    Write-Host $Message
    Write-Host "Required scopes/permissions:"
    foreach ($scope in $RequiredScopes) {
        Write-Host "  - $scope"
    }

    while ($true) {
        $choice = Read-Host "Request these permissions for the current account (R) or sign in with a different account (S)? [R/S]"
        if ([string]::IsNullOrWhiteSpace($choice)) {
            $choice = 'R'
        }

        switch ($choice.Trim().ToLowerInvariant()) {
            'r' {
                Write-Host "Please request the required permissions for account '$AccountName' through your admin or consent process."
                return 'request'
            }
            's' {
                return 'switch'
            }
            default {
                Write-Host '  Please enter R or S.'
            }
        }
    }
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

            # Increase CSOM request timeout to reduce long-running permission/read failures.
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

function Reset-PnPAuthenticationConnection {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [Parameter(Mandatory = $true)][string]$Url,
        [Parameter(Mandatory = $false)][string]$ClientId,
        [Parameter(Mandatory = $false)][switch]$AdminConnection
    )

    Write-Warning "Authentication issue detected for $Url; retrying the interactive PnP connection."

    if ($PSCmdlet.ShouldProcess($Url, 'Reset PnP authentication connection')) {
        return Connect-ValidatedPnP -Url $Url -ClientId $ClientId -AdminConnection:$AdminConnection
    }
}

function Test-IsTimeoutFailure {
    param([Parameter(Mandatory = $true)]$ErrorRecord)

    $msg = [string]$ErrorRecord.Exception.Message
    return ($msg -match '(?i)timed out|timeout|request was canceled|HttpClient\.Timeout|The operation has timed out')
}

function Test-IsAuthFailureMessage {
    param([Parameter(Mandatory = $true)][string]$Message)

    if ([string]::IsNullOrWhiteSpace($Message)) {
        return $false
    }

    return ($Message -match '(?i)(not signed in|use connect-pnponline|sign in|signin|token|unauthorized|authentication failed|authentication required|401|expired token)')
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
        Result = $result
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

function Get-TimestampedReportPath {
    param(
        [Parameter(Mandatory = $true)][string]$Folder,
        [Parameter(Mandatory = $true)][string]$Prefix
    )

    $stamp = Get-Date -Format "yyyyMMdd-HHmmss"
    return Join-Path -Path $Folder -ChildPath "$Prefix-$stamp.csv"
}

function Get-SlugValue {
    param([Parameter(Mandatory = $true)][string]$Text)

    $slug = [regex]::Replace($Text, '[^A-Za-z0-9]', '')
    if ([string]::IsNullOrWhiteSpace($slug)) {
        return ''
    }

    return $slug
}

function Get-SafeFileSlug {
    param(
        [Parameter(Mandatory = $true)][string]$Text,
        [Parameter(Mandatory = $false)][int]$MaxLength = 72,
        [Parameter(Mandatory = $false)][string]$Fallback = 'SelectedSites'
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

function Get-ExternalUrlStatusCachePath {
    param(
        [Parameter(Mandatory = $true)][string]$StateRoot,
        [Parameter(Mandatory = $true)][string]$ScopeUrl,
        [Parameter(Mandatory = $true)][string]$SchemaVersion
    )

    $fingerprint = (Get-Sha256Hex -Text ($ScopeUrl.TrimEnd('/').ToLowerInvariant() + '|' + $SchemaVersion)).Substring(0, 16)
    return Join-Path -Path $StateRoot -ChildPath ("Get-FixItListURLsv5_ExternalUrlCache_{0}.json" -f $fingerprint)
}

function Get-InventorySubsiteCachePath {
    param(
        [Parameter(Mandatory = $true)][string]$StateRoot,
        [Parameter(Mandatory = $true)][string]$ScopeUrl,
        [Parameter(Mandatory = $true)][string]$SchemaVersion
    )

    $fingerprint = (Get-Sha256Hex -Text ($ScopeUrl.TrimEnd('/').ToLowerInvariant() + '|' + $SchemaVersion + '|' + $InventoryCacheSchemaVersion)).Substring(0, 16)
    return Join-Path -Path $StateRoot -ChildPath ("Get-FixItListURLsv5_InventoryCache_{0}.json" -f $fingerprint)
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
                    Url = ([string]$entry.Url).TrimEnd('/')
                    Title = [string]$entry.Title
                    IsScannable = [bool]$entry.IsScannable
                    Reason = [string]$entry.Reason
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
            SchemaVersion = $SchemaVersion
            InventorySchemaVersion = $InventoryCacheSchemaVersion
            ScopeUrl = $ScopeUrl
            SavedAt = (Get-Date).ToString('o')
            Entries = $Entries
        } | ConvertTo-Json -Depth 6 | Set-Content -Path $Path -Encoding UTF8 -Force
    }
    catch {
        Write-Verbose "Could not save inventory cache to ${Path}: $($_.Exception.Message)"
    }
}

function Get-ExternalUrlStatusCache {
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
            foreach ($p in $parsed.Entries.PSObject.Properties) {
                $cache[[string]$p.Name] = $p.Value
            }
        }

        return $cache
    }
    catch {
        Write-Verbose "Could not load external URL cache from ${Path}: $($_.Exception.Message)"
        return @{}
    }
}

function Save-ExternalUrlStatusCache {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][hashtable]$Cache,
        [Parameter(Mandatory = $true)][string]$SchemaVersion
    )

    try {
        [PSCustomObject][ordered]@{
            SchemaVersion = $SchemaVersion
            SavedAt = (Get-Date).ToString('o')
            Entries = $Cache
        } | ConvertTo-Json -Depth 8 | Set-Content -Path $Path -Encoding UTF8 -Force
    }
    catch {
        Write-Verbose "Could not save external URL cache to ${Path}: $($_.Exception.Message)"
    }
}

function Get-UniqueSubsiteScanRootSet {
    param([Parameter(Mandatory = $true)][array]$Subsites)

    $ordered = @($Subsites | Where-Object { $_ -and $_.Url } | Sort-Object { $_.Url.TrimEnd('/').Length }, Url)
    $roots = New-Object System.Collections.Generic.List[object]

    foreach ($sub in $ordered) {
        $subUrl = [string]$sub.Url
        if ([string]::IsNullOrWhiteSpace($subUrl)) {
            continue
        }

        $normalized = $subUrl.TrimEnd('/')
        $isCovered = $false
        foreach ($root in $roots) {
            $rootUrl = ([string]$root.Url).TrimEnd('/')
            if ($normalized.Equals($rootUrl, [System.StringComparison]::OrdinalIgnoreCase) -or
                $normalized.StartsWith(($rootUrl + '/'), [System.StringComparison]::OrdinalIgnoreCase)) {
                $isCovered = $true
                break
            }
        }

        if (-not $isCovered) {
            $roots.Add([PSCustomObject]@{ Url = $normalized; Title = [string]$sub.Title })
        }
    }

    return @($roots | Sort-Object Url -Unique)
}

function Get-NormalizedExtension {
    param([Parameter(Mandatory = $true)][string]$FileName)

    $normalizedName = $FileName.Trim().TrimEnd('.')
    if ([string]::IsNullOrWhiteSpace($normalizedName)) {
        return ''
    }

    $ext = [System.IO.Path]::GetExtension($normalizedName)
    if ([string]::IsNullOrWhiteSpace($ext)) {
        return ''
    }

    $normalized = $ext.Trim().ToLowerInvariant()
    if ($normalized.EndsWith('.')) {
        $normalized = $normalized.TrimEnd('.')
    }

    return $normalized
}

function Add-ExtensionCount {
    param(
        [Parameter(Mandatory = $true)][hashtable]$Bucket,
        [Parameter(Mandatory = $true)][AllowEmptyString()][string]$Extension,
        [Parameter(Mandatory = $false)][int]$Increment = 1
    )

    $key = if ([string]::IsNullOrWhiteSpace($Extension)) { '(none)' } else { $Extension.ToLowerInvariant() }
    if (-not $Bucket.ContainsKey($key)) {
        $Bucket[$key] = 0
    }

    $Bucket[$key] += $Increment
}

function Get-ExtensionCountTotal {
    param([Parameter(Mandatory = $true)][hashtable]$Bucket)

    if ($Bucket.Count -eq 0) {
        return 0
    }

    return (@($Bucket.Values) | Measure-Object -Sum).Sum
}

function Write-ExtensionCountBlock {
    param(
        [Parameter(Mandatory = $true)][string]$Title,
        [Parameter(Mandatory = $true)][hashtable]$Bucket
    )

    Write-Host $Title
    if ($Bucket.Count -eq 0) {
        Write-Host '    (none)'
        return
    }

    foreach ($ext in @($Bucket.Keys | Sort-Object)) {
        Write-Host ("    {0,-10} {1}" -f $ext, $Bucket[$ext])
    }
}

function Add-ExtensionDetailRows {
    param(
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][System.Collections.Generic.List[object]]$Rows,
        [Parameter(Mandatory = $true)][string]$Category,
        [Parameter(Mandatory = $true)][hashtable]$Bucket
    )

    if ($Bucket.Count -eq 0) {
        $Rows.Add([PSCustomObject][ordered]@{
                Category  = $Category
                Extension = '(none)'
                Count     = 0
            })
        return
    }

    foreach ($ext in @($Bucket.Keys | Sort-Object)) {
        $Rows.Add([PSCustomObject][ordered]@{
                Category  = $Category
                Extension = $ext
                Count     = [int]$Bucket[$ext]
            })
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

function Test-IsInventorySubsiteScannable {
    param(
        [Parameter(Mandatory = $true)][string]$WebUrl,
        [Parameter(Mandatory = $true)][string]$RootHost,
        [Parameter(Mandatory = $false)][string]$ClientId
    )

    try {
        $webHost = ([Uri]$WebUrl).Host.ToLowerInvariant()
        if (-not $webHost.Equals($RootHost, [System.StringComparison]::OrdinalIgnoreCase)) {
            return [PSCustomObject]@{
                IsScannable = $false
                Reason = 'Different host than tenant root (likely app/deprecated web).'
            }
        }
    }
    catch {
        return [PSCustomObject]@{
            IsScannable = $false
            Reason = 'Invalid URL format.'
        }
    }

    $webConn = $null
    try {
        $webConn = Connect-ValidatedPnP -Url $WebUrl -ClientId $ClientId
    }
    catch {
        return [PSCustomObject]@{
            IsScannable = $false
            Reason = 'Connection/auth failed for this web.'
        }
    }

    try {
        $webResult = Invoke-PnPConnectionOperation -Url $WebUrl -ClientId $ClientId -Connection $webConn -Label ('Inventory scannability web metadata: ' + $WebUrl) -MaxAttempts 3 -Operation {
            param($activeConnection)
            Get-PnPWeb -Connection $activeConnection -Includes WebTemplate, Configuration
        }
        $webConn = $webResult.Connection
        $web = $webResult.Result

        $webTemplateName = if ($web -and $web.WebTemplate) { [string]$web.WebTemplate } else { '' }
        $webTemplateConfig = if ($web -and $web.PSObject.Properties['Configuration']) { [string]$web.Configuration } else { '' }
        $webTemplateFull = if (-not [string]::IsNullOrWhiteSpace($webTemplateName) -and -not [string]::IsNullOrWhiteSpace($webTemplateConfig)) {
            ($webTemplateName + '#' + $webTemplateConfig)
        }
        else {
            $webTemplateName
        }

        if (-not [string]::IsNullOrWhiteSpace($webTemplateName) -and ($webTemplateName -match '(?i)^APP')) {
            return [PSCustomObject]@{
                IsScannable = $false
                Reason = 'App web template is not scanned by this script.'
            }
        }

        if (-not [string]::IsNullOrWhiteSpace($webTemplateName) -and (
                $webTemplateName.Equals('MPS#0', [System.StringComparison]::OrdinalIgnoreCase) -or
                $webTemplateName.Equals('MPS', [System.StringComparison]::OrdinalIgnoreCase) -or
                $webTemplateFull.Equals('MPS#0', [System.StringComparison]::OrdinalIgnoreCase)
            )) {
            return [PSCustomObject]@{
                IsScannable = $false
                Reason = ("MPS#0 site template is excluded from this script inventory. Detected template: {0}" -f $webTemplateFull)
            }
        }

        $listResult = Invoke-PnPConnectionOperation -Url $WebUrl -ClientId $ClientId -Connection $webConn -Label ('Inventory scannability lists: ' + $WebUrl) -MaxAttempts 3 -Operation {
            param($activeConnection)
            Get-PnPList -Connection $activeConnection -Includes BaseTemplate, Hidden, Title, ItemCount, RootFolder.ServerRelativeUrl
        }
        $lists = @($listResult.Result)

        foreach ($list in $lists) {
            if ($list.Hidden) { continue }
            if (Test-IsRecycleBinContainer -ListTitle $list.Title -RootFolderPath $list.RootFolder.ServerRelativeUrl) { continue }
            if ($SkipListTitles -contains $list.Title) { continue }
            if ($list.ItemCount -eq 0) { continue }

            return [PSCustomObject]@{
                IsScannable = $true
                Reason = ''
            }
        }

        return [PSCustomObject]@{
            IsScannable = $false
            Reason = 'No non-empty scannable lists/libraries found.'
        }
    }
    catch {
        return [PSCustomObject]@{
            IsScannable = $false
            Reason = ('Scannability probe failed: ' + $_.Exception.Message)
        }
    }
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

function Get-NormalizedUrlForMatching {
    param([Parameter(Mandatory = $true)][string]$Value)

    $clean = Get-CleanedValue -Value $Value
    if ([string]::IsNullOrWhiteSpace($clean)) {
        return ''
    }

    $centralParkFix = Get-CentralParkCorrection -Url $clean
    if ($centralParkFix) {
        return $centralParkFix.NewUrl
    }

    return $clean
}

function Test-IsSourceUrlOnSelectedBase {
    param(
        [Parameter(Mandatory = $true)][string]$SourceUrl,
        [Parameter(Mandatory = $true)][array]$Mappings
    )

    if ([string]::IsNullOrWhiteSpace($SourceUrl)) {
        return $false
    }

    $clean = Get-CleanedValue -Value $SourceUrl
    foreach ($m in $Mappings) {
        if ($clean.Equals($m.OldBaseUrl, [System.StringComparison]::OrdinalIgnoreCase) -or
            $clean.StartsWith(($m.OldBaseUrl + '/'), [System.StringComparison]::OrdinalIgnoreCase)) {
            return $true
        }
    }

    return $false
}

function Test-IsNonActionableUrl {
    param([Parameter(Mandatory = $true)][string]$Url)

    $clean = Get-CleanedValue -Value $Url
    if ([string]::IsNullOrWhiteSpace($clean)) {
        return $true
    }

    # Legacy on-prem SharePoint library/view links can appear in Word field-code internals.
    if ($clean -match '(?i)^https?://(?:centralpark|whaiis01)(?=$|[/:?#])' -and
        $clean -match '(?i)/forms/(?:default|allitems)\.aspx(?:$|[?#])') {
        return $true
    }

    # XML / PDF / Office metadata namespaces and schema anchors are not user-authored content.
    if ($clean -match '(?i)^https?://(?:www\.)?w3\.org(?:/|$)' -or
        $clean -match '(?i)^https?://(?:www\.)?ns\.adobe\.com(?:/|$)' -or
        $clean -match '(?i)^https?://(?:www\.)?activepdf\.com(?:/|$)' -or
        $clean -match '(?i)^https?://(?:www\.)?purl\.org(?:/|$)' -or
        $clean -match '(?i)^https?://(?:www\.)?dublincore\.org(?:/|$)' -or
        $clean -match '(?i)^https?://(?:www\.)?iptc\.org(?:/|$)' -or
        $clean -match '(?i)^https?://schemas\.openxmlformats\.org(?:/|$)' -or
        $clean -match '(?i)^https?://schemas\.microsoft\.com(?:/|$)' -or
        $clean -match '(?i)^https?://schemas\.') {
        return $true
    }

    # Service plumbing URLs and clearly incomplete hosts are often system-generated content.
    if ($clean -match '(?i)^https?://[^/]*pushfp\.svc\.ms(?:/|$)' -or
        $clean -match '(?i)^https?://forms\.office\.com/images/icon-form-logo\.png(?:$|[?#])' -or
        $clean -match '(?i)^https?://www(?:$|[/?#])') {
        return $true
    }

    # SharePoint auto-generated form/view/list URLs and doc metadata redirects are not human-placed content.
    if ($clean -match '(?i)/_layouts/(?:15/)?(?:DocIdRedir|ListForm|ViewForm|EditForm|NewForm|DisplayForm|FormServer|Redirector|Dialog|ErrorPage)\.aspx(?:\?|$)' -or
        $clean -match '(?i)/_layouts/(?:15/)?(?:docidredir|listform|viewform|editform|newform|displayform|formserver|redirector|dialog|errorpage)\.aspx(?:\?|$)' -or
        $clean -match '(?i)(?:^|/)_layouts/(?:15/)?[^/?#]+\.aspx\?(?:.*(?:PageType=|Source=|ListId=|ID=|CID=|RootFolder=))' -or
        $clean -match '(?i)(?:^|/)_layouts/(?:15/)?[^/?#]+\.aspx(?:$|[?#])') {
        return $true
    }

    # DTD file references embedded in document/PDF binary metadata; not human-placed.
    if ($clean -match '(?i)/DTDs/[^/?#]+\.dtd(?:$|[?#])') {
        return $true
    }

    # Office relationship/metadata package URLs are internal document scaffolding, not authored content.
    if ($clean -match '(?i)(?:^|/)(?:_rels|docProps|package|relationships)(?:/|$)' -or
        $clean -match '(?i)(?:^|/)(?:document|core|extended-properties|custom-properties)\.xml(?:$|[?#])') {
        return $true
    }

    # Discard URLs whose host portion is too short to be a real hostname (binary scan artifact).
    $uri = $null
    if ([Uri]::TryCreate($clean, [UriKind]::Absolute, [ref]$uri) -and $uri.Host.Length -le 1) {
        return $true
    }

    return $false
}

function Test-IsReferenceSectionText {
    param([Parameter(Mandatory = $false)][string]$Text)

    if ([string]::IsNullOrWhiteSpace($Text)) {
        return $false
    }

    $norm = (($Text -replace '\s+', ' ').Trim()).ToLowerInvariant()
    return ($norm -match '\b(reference|references|bibliography|works cited|citations|footnote|footnotes|endnote|endnotes)\b')
}

function Test-IsSuppressedDocumentReference {
    param([Parameter(Mandatory = $true)]$Reference)

    if ($Reference.PSObject.Properties['Kind']) {
        $kind = [string]$Reference.Kind
        if ($kind -match '(?i)header|footer|notes') {
            return $true
        }
    }

    $textParts = @()
    if ($Reference.PSObject.Properties['Text']) {
        $textParts += [string]$Reference.Text
    }
    if ($Reference.PSObject.Properties['Context']) {
        $textParts += [string]$Reference.Context
    }

    foreach ($part in $textParts) {
        if (Test-IsReferenceSectionText -Text $part) {
            return $true
        }
    }

    return $false
}

function Test-IsExactRowDuplicate {
    param(
        [Parameter(Mandatory = $true)][pscustomobject]$Row,
        [Parameter(Mandatory = $true)]$SeenRows
    )

    $fingerprint = $Row | ConvertTo-Json -Compress -Depth 6
    if ($SeenRows.Contains($fingerprint)) {
        return $true
    }

    $null = $SeenRows.Add($fingerprint)
    return $false
}

function Format-ReportDatePacific {
    param([Parameter(Mandatory = $false)]$Value)

    if (-not $Value) {
        return ''
    }

    $dt = $null
    if ($Value -is [datetime]) {
        $dt = [datetime]$Value
    }
    else {
        $parsed = $null
        if (-not [datetime]::TryParse([string]$Value, [ref]$parsed)) {
            return [string]$Value
        }
        $dt = $parsed
    }

    try {
        $pacificTz = [System.TimeZoneInfo]::FindSystemTimeZoneById('Pacific Standard Time')
        $utcDate = if ($dt.Kind -eq [System.DateTimeKind]::Utc) { $dt } else { $dt.ToUniversalTime() }
        $pacificDate = [System.TimeZoneInfo]::ConvertTimeFromUtc($utcDate, $pacificTz)
        return $pacificDate.ToString('yyyy-MM-dd HH:mm:ss')
    }
    catch {
        return $dt.ToString('yyyy-MM-dd HH:mm:ss')
    }
}

function Get-UrlDisplayText {
    param(
        [Parameter(Mandatory = $true)][string]$SourceText,
        [Parameter(Mandatory = $true)][string]$Url,
        [Parameter(Mandatory = $false)][int]$MaxLength = 180
    )

    $normalized = ([string]$SourceText -replace '\s+', ' ').Trim()
    if ([string]::IsNullOrWhiteSpace($normalized)) {
        return $Url
    }

    $idx = $normalized.IndexOf($Url, [System.StringComparison]::OrdinalIgnoreCase)
    if ($idx -lt 0) {
        if ($normalized.Length -le $MaxLength) {
            return $normalized
        }
        return $normalized.Substring(0, $MaxLength)
    }

    $start = [Math]::Max(0, $idx - 60)
    $len = [Math]::Min($MaxLength, $normalized.Length - $start)
    $snippet = $normalized.Substring($start, $len).Trim()
    if ($start -gt 0) { $snippet = '...' + $snippet }
    if (($start + $len) -lt $normalized.Length) { $snippet += '...' }
    return $snippet
}

function Get-OfficeDocReference {
    param([Parameter(Mandatory = $true)][System.IO.Stream]$Stream)

    $found = New-Object System.Collections.Generic.List[Object]
    $relationsByXml = @{}
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
                    $xmlPath = $entry.FullName -replace '/_rels/([^/]+)\.rels$', '/$1'
                    $relationMap = @{}
                    foreach ($rel in $xml.Relationships.Relationship) {
                        if ($rel.TargetMode -eq "External" -and $rel.Target) {
                            $relationMap[$rel.Id] = $rel.Target
                        }
                    }

                    if ($relationMap.Count -gt 0) {
                        $relationsByXml[$xmlPath] = $relationMap
                    }
                }
                catch {
                    Write-Verbose "Could not parse relationship XML in $($entry.FullName)."
                }
            }
            elseif ($entry.FullName -match '\.xml$' -and
                ($entry.FullName -match 'document\.xml$|sheet\d*\.xml$|slide\d*\.xml$')) {

                $s = $entry.Open()
                $r = New-Object System.IO.StreamReader($s)
                $text = $r.ReadToEnd()
                $r.Close()
                $s.Close()

                $plain = [regex]::Replace($text, '<[^>]+>', ' ')

                if ($relationsByXml.ContainsKey($entry.FullName)) {
                    try {
                        [xml]$xml = $text
                        $ns = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
                        $ns.AddNamespace('w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main')
                        $ns.AddNamespace('r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')

                        foreach ($hyperlink in $xml.SelectNodes('//w:hyperlink[@r:id]', $ns)) {
                            $relId = $hyperlink.GetAttribute('id', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
                            if ($relationsByXml[$entry.FullName].ContainsKey($relId)) {
                                $displayText = ([regex]::Replace($hyperlink.OuterXml, '<[^>]+>', ' ')).Trim()
                                if ([string]::IsNullOrWhiteSpace($displayText)) {
                                    $displayText = $relationsByXml[$entry.FullName][$relId]
                                }

                                $context = Get-UrlDisplayText -SourceText $plain -Url $relationsByXml[$entry.FullName][$relId]

                                $found.Add([PSCustomObject]@{ Value = $relationsByXml[$entry.FullName][$relId]; Text = $displayText; Context = $context; Confidence = "High"; Kind = "hyperlink" })
                            }
                        }
                    }
                    catch {
                        Write-Verbose "Could not parse hyperlink display text in $($entry.FullName)."
                    }
                }

            }
        }
    }
    finally {
        $archive.Dispose()
    }

    return $found | Sort-Object Value, Kind, Text -Unique
}

function Get-PdfReference {
    param([Parameter(Mandatory = $true)][byte[]]$Bytes)

    $found = New-Object System.Collections.Generic.List[Object]
    $text = [System.Text.Encoding]::Latin1.GetString($Bytes)

    # Prefer explicit PDF URI actions. These are far more likely to be user-authored links.
    foreach ($m in [regex]::Matches($text, '(?is)/URI\s*\((https?://[^)\s<>"]+)\)')) {
        $candidate = [string]$m.Groups[1].Value
        $hostMatch = [regex]::Match($candidate, '^https?://([^/:?#\s\x00-\x1f\u0080-\uffff]+)')
        if ($hostMatch.Success -and ($PdfNoiseHosts -contains $hostMatch.Groups[1].Value.ToLowerInvariant())) {
            continue
        }

        if (Test-IsNonActionableUrl -Url $candidate) {
            continue
        }

        $context = Get-UrlDisplayText -SourceText $text -Url $candidate
        $found.Add([PSCustomObject]@{ Value = $candidate; Text = $candidate; Context = $context; Confidence = "High"; Kind = "pdf-uri-action" })
    }

    if ($found.Count -eq 0) {
        foreach ($m in [regex]::Matches($text, 'https?://[^\s\x00-\x1f\u0080-\uffff\)"<>\\]+')) {
            $hostMatch = [regex]::Match($m.Value, '^https?://([^/:?#\s\x00-\x1f\u0080-\uffff]+)')
            if ($hostMatch.Success -and ($PdfNoiseHosts -contains $hostMatch.Groups[1].Value.ToLowerInvariant())) {
                continue
            }

            if (Test-IsNonActionableUrl -Url $m.Value) {
                continue
            }

            $context = Get-UrlDisplayText -SourceText $text -Url $m.Value
            $found.Add([PSCustomObject]@{ Value = $m.Value; Text = $m.Value; Context = $context; Confidence = "Low"; Kind = "plaintext-pdf-fallback" })
        }
    }

    return $found | Sort-Object Value, Kind, Text -Unique
}

function Get-TextReferenceFromByteContent {
    param(
        [Parameter(Mandatory = $true)][byte[]]$Bytes,
        [Parameter(Mandatory = $false)][string]$Kind = 'plaintext'
    )

    $found = New-Object System.Collections.Generic.List[Object]
    $decodedTexts = New-Object System.Collections.Generic.List[string]

    try {
        $decodedTexts.Add([System.Text.Encoding]::UTF8.GetString($Bytes))
    }
    catch {
        Write-Verbose 'Could not decode byte content as UTF-8.'
    }

    try {
        $decodedTexts.Add([System.Text.Encoding]::Unicode.GetString($Bytes))
    }
    catch {
        Write-Verbose 'Could not decode byte content as UTF-16.'
    }

    $decodedTexts.Add([System.Text.Encoding]::Latin1.GetString($Bytes))

    foreach ($text in $decodedTexts) {
        foreach ($m in [regex]::Matches($text, 'https?://[^\s\x00-\x1f\u0080-\uffff\)"<>\\]+')) {
            $context = Get-UrlDisplayText -SourceText $text -Url $m.Value
            $found.Add([PSCustomObject]@{ Value = $m.Value; Text = $m.Value; Context = $context; Confidence = 'Low'; Kind = $Kind })
        }
    }

    return $found | Sort-Object Value, Kind, Text -Unique
}

function Get-UrlMappingMatch {
    param(
        [Parameter(Mandatory = $true)][string]$Value,
        [Parameter(Mandatory = $true)][array]$Mappings
    )

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return $null
    }

    $clean = Get-NormalizedUrlForMatching -Value $Value

    foreach ($m in $Mappings) {
        $idx = $clean.IndexOf($m.OldBaseUrl, [System.StringComparison]::OrdinalIgnoreCase)
        if ($idx -ge 0) {
            $suffix = $clean.Substring($idx + $m.OldBaseUrl.Length)
            $oldUrl = $clean.Substring($idx)
            $newUrl = $m.NewBaseUrl + $suffix
            return [PSCustomObject]@{
                OldBaseUrl = $m.OldBaseUrl
                NewBaseUrl = $m.NewBaseUrl
                OldUrl     = $oldUrl
                NewUrl     = $newUrl
            }
        }

        $viewerPattern = '^(?i)' + [regex]::Escape($m.TenantBase) + '/:[a-z]:/[rtu]' + [regex]::Escape($m.OldBasePath) + '(?=$|[/?#])'
        $vm = [regex]::Match($clean, $viewerPattern)
        if ($vm.Success) {
            $suffix = $clean.Substring($vm.Length)
            $oldUrl = $m.OldBaseUrl + $suffix
            $newUrl = $m.NewBaseUrl + $suffix
            return [PSCustomObject]@{
                OldBaseUrl = $m.OldBaseUrl
                NewBaseUrl = $m.NewBaseUrl
                OldUrl     = $oldUrl
                NewUrl     = $newUrl
            }
        }

        if ($clean.StartsWith($m.OldBasePath, [System.StringComparison]::OrdinalIgnoreCase)) {
            $suffix = $clean.Substring($m.OldBasePath.Length)
            $oldUrl = $m.OldBaseUrl + $suffix
            $newUrl = $m.NewBaseUrl + $suffix
            return [PSCustomObject]@{
                OldBaseUrl = $m.OldBaseUrl
                NewBaseUrl = $m.NewBaseUrl
                OldUrl     = $oldUrl
                NewUrl     = $newUrl
            }
        }
    }

    return $null
}

function Get-SiteKeyForSourceUrl {
    param(
        [Parameter(Mandatory = $true)][string]$SourceUrl,
        [Parameter(Mandatory = $true)][array]$Mappings
    )

    if ([string]::IsNullOrWhiteSpace($SourceUrl)) {
        return ''
    }

    $clean = Get-CleanedValue -Value $SourceUrl
    if ([string]::IsNullOrWhiteSpace($clean)) {
        return ''
    }

    foreach ($m in $Mappings) {
        $base = [string]$m.OldBaseUrl
        if ([string]::IsNullOrWhiteSpace($base)) {
            continue
        }

        if ($clean.Equals($base, [System.StringComparison]::OrdinalIgnoreCase) -or
            $clean.StartsWith(($base + '/'), [System.StringComparison]::OrdinalIgnoreCase)) {
            return $base
        }
    }

    return ''
}

function Get-CentralParkCorrection {
    param([Parameter(Mandatory = $true)][string]$Url)

    $clean = Get-CleanedValue -Value $Url
    if ([string]::IsNullOrWhiteSpace($clean)) {
        return $null
    }

    # centralpark was the on-prem server name for the SharePoint Online tenant root.
    $prefixPattern = '^(?i)https?://centralpark(?=$|[/:?#])'
    if ($clean -match $prefixPattern) {
        $newUrl = [regex]::Replace($clean, $prefixPattern, 'https://whallc.sharepoint.com')
        return [PSCustomObject]@{ OldUrl = $clean; NewUrl = $newUrl }
    }

    # whaiis01 was the on-prem server name for the /whaiis01 site collection.
    $whaiis01Pattern = '^(?i)https?://whaiis01(?=$|[/:?#])'
    if ($clean -match $whaiis01Pattern) {
        $newUrl = [regex]::Replace($clean, $whaiis01Pattern, 'https://whallc.sharepoint.com/whaiis01')
        return [PSCustomObject]@{ OldUrl = $clean; NewUrl = $newUrl }
    }

    return $null
}

function Get-LinkTypeLabel {
    param(
        [Parameter(Mandatory = $true)][string]$Url,
        [Parameter(Mandatory = $false)][string[]]$TenantHosts = @()
    )

    $clean = Get-CleanedValue -Value $Url
    if ([string]::IsNullOrWhiteSpace($clean)) {
        return 'External'
    }

    if (Get-CentralParkCorrection -Url $clean) {
        return 'Internal'
    }

    $uri = $null
    if (-not [Uri]::TryCreate($clean, [UriKind]::Absolute, [ref]$uri)) {
        return 'External'
    }

    $tenantHostValue = $uri.Host.ToLowerInvariant()
    if ($tenantHostValue -eq 'whallc.sharepoint.com' -or $tenantHostValue.EndsWith('.whallc.sharepoint.com') -or ($TenantHosts -contains $tenantHostValue)) {
        return 'Internal'
    }

    return 'External'
}

function Get-RoleAssignmentTopLevel {
    param(
        [Parameter(Mandatory = $true)]$Securable,
        [Parameter(Mandatory = $true)]$Connection,
        [Parameter(Mandatory = $true)][string]$ConnectionUrl,
        [Parameter(Mandatory = $false)][string]$ClientId,
        [Parameter(Mandatory = $true)][string]$CacheKey,
        [Parameter(Mandatory = $true)][hashtable]$Cache
    )

    if ($Cache.ContainsKey($CacheKey)) {
        return $Cache[$CacheKey]
    }

    $entries = New-Object System.Collections.Generic.List[string]
    $assignmentExpansionFailures = 0
    $assignments = $null
    $null = $Securable
    $null = $Connection
    try {
        $assignmentResult = Invoke-PnPConnectionOperation -Url $ConnectionUrl -ClientId $ClientId -Connection $Connection -Label 'Get role assignments' -MaxAttempts 4 -Operation {
            param($activeConnection)
            Get-PnPProperty -ClientObject $Securable -Property RoleAssignments -Connection $activeConnection
        }
        $Connection = $assignmentResult.Connection
        $assignments = $assignmentResult.Result
    }
    catch {
        $msg = $_.Exception.Message
        $result = "(could not read permissions: $msg)"
        if (-not (Test-IsAuthFailureMessage -Message $msg)) {
            $Cache[$CacheKey] = $result
        }
        return $result
    }

    foreach ($ra in $assignments) {
        $member = $null
        $bindings = $null
        try {
            $memberResult = Invoke-PnPConnectionOperation -Url $ConnectionUrl -ClientId $ClientId -Connection $Connection -Label 'Get role assignment member' -MaxAttempts 4 -Operation {
                param($activeConnection)
                Get-PnPProperty -ClientObject $ra -Property Member -Connection $activeConnection
            }
            $Connection = $memberResult.Connection
            $member = $memberResult.Result

            $bindingResult = Invoke-PnPConnectionOperation -Url $ConnectionUrl -ClientId $ClientId -Connection $Connection -Label 'Get role assignment bindings' -MaxAttempts 4 -Operation {
                param($activeConnection)
                Get-PnPProperty -ClientObject $ra -Property RoleDefinitionBindings -Connection $activeConnection
            }
            $Connection = $bindingResult.Connection
            $bindings = $bindingResult.Result
        }
        catch {
            $assignmentExpansionFailures++
            continue
        }
        $roles = @($bindings | Where-Object { $_.Name -ne 'Limited Access' } | Select-Object -ExpandProperty Name)

        if ($roles.Count -eq 0) {
            continue
        }

        $roleText = $roles -join '/'
        if ($member.PrincipalType -eq 'SharePointGroup' -and $member.LoginName -like 'SharingLinks.*') {
            $kind = ($member.LoginName -split '\.')[2]
            if ([string]::IsNullOrWhiteSpace($kind)) { $kind = 'sharing' }
            $entries.Add("Link: $kind [$roleText]")
        }
        elseif ($member.PrincipalType -eq 'SharePointGroup') {
            $entries.Add("Group: $($member.Title) [$roleText]")
        }
        elseif ($member.PrincipalType -eq 'SecurityGroup') {
            $entries.Add("Security Group: $($member.Title) [$roleText]")
        }
        elseif ($member.PrincipalType -eq 'User') {
            if ($member.Email) {
                $entries.Add("User: $($member.Title) <$($member.Email)> [$roleText]")
            }
            else {
                $entries.Add("User: $($member.Title) [$roleText]")
            }
        }
        else {
            $entries.Add("$($member.Title) [$roleText]")
        }
    }

    if ($entries.Count -eq 0 -and $assignmentExpansionFailures -gt 0) {
        $result = '(could not read permissions: failed to expand one or more role assignments)'
        $Cache[$CacheKey] = $result
        return $result
    }

    $result = ($entries | Sort-Object -Unique) -join "`n"
    $Cache[$CacheKey] = $result
    return $result
}

function Get-EffectiveTopLevelPermission {
    param(
        [Parameter(Mandatory = $true)]$Item,
        [Parameter(Mandatory = $true)]$List,
        [Parameter(Mandatory = $true)]$Web,
        [Parameter(Mandatory = $true)]$Connection,
        [Parameter(Mandatory = $true)][string]$ConnectionUrl,
        [Parameter(Mandatory = $false)][string]$ClientId,
        [Parameter(Mandatory = $true)][hashtable]$Cache
    )

    $itemUnique = $false
    $listUnique = $false

    try {
        $itemUniqueResult = Invoke-PnPConnectionOperation -Url $ConnectionUrl -ClientId $ClientId -Connection $Connection -Label 'Get item unique permission flag' -MaxAttempts 4 -Operation {
            param($activeConnection)
            Get-PnPProperty -ClientObject $Item -Property HasUniqueRoleAssignments -Connection $activeConnection
        }
        $Connection = $itemUniqueResult.Connection
        $itemUnique = [bool]$itemUniqueResult.Result
    }
    catch {
        Write-Verbose "Could not determine item-level unique permissions for item $($Item.Id)."
    }

    try {
        $listUniqueResult = Invoke-PnPConnectionOperation -Url $ConnectionUrl -ClientId $ClientId -Connection $Connection -Label 'Get list unique permission flag' -MaxAttempts 4 -Operation {
            param($activeConnection)
            Get-PnPProperty -ClientObject $List -Property HasUniqueRoleAssignments -Connection $activeConnection
        }
        $Connection = $listUniqueResult.Connection
        $listUnique = [bool]$listUniqueResult.Result
    }
    catch {
        Write-Verbose "Could not determine list-level unique permissions for list $($List.Title)."
    }

    if ($itemUnique) {
        return Get-RoleAssignmentTopLevel -Securable $Item -Connection $Connection -ConnectionUrl $ConnectionUrl -ClientId $ClientId -CacheKey ("Item:" + $List.Id + ":" + $Item.Id) -Cache $Cache
    }

    if ($listUnique) {
        return Get-RoleAssignmentTopLevel -Securable $List -Connection $Connection -ConnectionUrl $ConnectionUrl -ClientId $ClientId -CacheKey ("List:" + $Web.Id + ":" + $List.Id) -Cache $Cache
    }

    return Get-RoleAssignmentTopLevel -Securable $Web -Connection $Connection -ConnectionUrl $ConnectionUrl -ClientId $ClientId -CacheKey ("Web:" + $Web.Id) -Cache $Cache
}

function Test-IsExternalBrokenUrl {
    param(
        [Parameter(Mandatory = $true)][string]$Url,
        [Parameter(Mandatory = $true)][string[]]$TenantHosts,
        [Parameter(Mandatory = $true)][hashtable]$Cache,
        [Parameter(Mandatory = $true)]$HttpClient
    )

    $clean = Get-CleanedValue -Value $Url
    if ([string]::IsNullOrWhiteSpace($clean)) {
        return $false
    }

    # Treat duplicated protocols and similar malformed values as broken/actionable.
    if ($clean -match '(?i)https?://https?://') {
        $Cache[$clean] = [ordered]@{ Status = $true; CheckedAt = (Get-Date).ToString('o') }
        return $true
    }

    if ($Cache.ContainsKey($clean)) {
        $cachedEntry = $Cache[$clean]
        if ($cachedEntry -is [bool]) {
            return [bool]$cachedEntry
        }

        $cachedStatus = $null
        $checkedAt = $null
        if ($cachedEntry -is [System.Collections.IDictionary]) {
            $cachedStatus = $cachedEntry['Status']
            $checkedAt = $cachedEntry['CheckedAt']
        }
        elseif ($cachedEntry.PSObject -and $cachedEntry.PSObject.Properties['Status']) {
            $cachedStatus = $cachedEntry.Status
            $checkedAt = $cachedEntry.CheckedAt
        }

        if ($null -ne $checkedAt) {
            try {
                $checkedAtValue = [datetime]::Parse([string]$checkedAt)
                if ((Get-Date) - $checkedAtValue -lt [TimeSpan]::FromDays($ExternalUrlStatusCacheTtlDays)) {
                    return [bool]$cachedStatus
                }
            }
            catch {
                Write-Verbose "Could not parse cached external URL status timestamp; refreshing cache entry."
            }
        }
    }

    $uri = $null
    if (-not [Uri]::TryCreate($clean, [UriKind]::Absolute, [ref]$uri)) {
        $Cache[$clean] = [ordered]@{ Status = $true; CheckedAt = (Get-Date).ToString('o') }
        return $true
    }

    if ($uri.Scheme -notin @('http', 'https')) {
        $Cache[$clean] = [ordered]@{ Status = $false; CheckedAt = (Get-Date).ToString('o') }
        return $false
    }

    $urlHost = $uri.Host.ToLower()
    $isSharePointHost = $urlHost -eq 'sharepoint.com' -or $urlHost.EndsWith('.sharepoint.com')
    if ($isSharePointHost -or ($TenantHosts -contains $urlHost)) {
        $Cache[$clean] = [ordered]@{ Status = $false; CheckedAt = (Get-Date).ToString('o') }
        return $false
    }

    $broken = $false
    try {
        $headReq = New-Object System.Net.Http.HttpRequestMessage([System.Net.Http.HttpMethod]::Head, $uri)
        $headResp = $HttpClient.Send($headReq, [System.Net.Http.HttpCompletionOption]::ResponseHeadersRead)
        $headStatus = [int]$headResp.StatusCode
        if ($headStatus -ge 400) {
            if ($headStatus -in 405, 501) {
                $getReq = New-Object System.Net.Http.HttpRequestMessage([System.Net.Http.HttpMethod]::Get, $uri)
                $getResp = $HttpClient.Send($getReq, [System.Net.Http.HttpCompletionOption]::ResponseHeadersRead)
                $getStatus = [int]$getResp.StatusCode
                $broken = ($getStatus -in @(400, 404, 410, 451))
            }
            else {
                if ($headStatus -in @(401, 403, 407, 408, 409, 423, 425, 429) -or $headStatus -ge 500) {
                    $broken = $false
                }
                else {
                    $broken = ($headStatus -in @(400, 404, 410, 451))
                }
            }
        }
    }
    catch {
        try {
            $getReq = New-Object System.Net.Http.HttpRequestMessage([System.Net.Http.HttpMethod]::Get, $uri)
            $getResp = $HttpClient.Send($getReq, [System.Net.Http.HttpCompletionOption]::ResponseHeadersRead)
            $getStatus = [int]$getResp.StatusCode
            if ($getStatus -in @(401, 403, 407, 408, 409, 423, 425, 429) -or $getStatus -ge 500) {
                $broken = $false
            }
            else {
                $broken = ($getStatus -in @(400, 404, 410, 451))
            }
        }
        catch {
            $broken = ([string]$_.Exception.Message -match '(?i)No such host is known|Name or service not known|DNS|getaddrinfo|Could not resolve host|The remote name could not be resolved')
        }
    }

    $Cache[$clean] = [ordered]@{ Status = $broken; CheckedAt = (Get-Date).ToString('o') }
    return $broken
}

function Get-FolderDisplayName {
    param(
        [Parameter(Mandatory = $true)]$List,
        [Parameter(Mandatory = $true)][string]$FolderPath,
        [Parameter(Mandatory = $true)]$Connection,
        [Parameter(Mandatory = $true)][string]$ConnectionUrl,
        [Parameter(Mandatory = $false)][string]$ClientId,
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
        $folderResult = Invoke-PnPConnectionOperation -Url $ConnectionUrl -ClientId $ClientId -Connection $Connection -Label ('Get folder display name: ' + $FolderPath) -MaxAttempts 4 -Operation {
            param($activeConnection)
            Get-PnPListItem -List $List.Title -Connection $activeConnection -Query $query -Fields "Title", "FileLeafRef"
        }
        $folderItem = $folderResult.Result
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
        Write-Verbose "Could not resolve folder display name for $FolderPath in list $($List.Title)."
    }

    $Cache[$cacheKey] = $fallback
    return $fallback
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

function Get-NewBaseUrlMapFromCsv {
    param(
        [Parameter(Mandatory = $true)][string]$CsvPath,
        [Parameter(Mandatory = $true)][array]$PromptSubsites
    )

    if (-not (Test-Path -Path $CsvPath)) {
        throw "CSV file was not found: $CsvPath"
    }

    $rows = @(Import-Csv -Path $CsvPath)
    if ($rows.Count -eq 0) {
        throw 'CSV file has no data rows. Include URL and New base URL values.'
    }

    $firstRow = $rows[0]
    $columnNames = @($firstRow.PSObject.Properties.Name)
    $urlColumn = @($columnNames | Where-Object { $_ -ieq 'URL' } | Select-Object -First 1)
    $newBaseColumn = @($columnNames | Where-Object { $_ -ieq 'New base URL' } | Select-Object -First 1)

    if ($urlColumn.Count -eq 0 -or $newBaseColumn.Count -eq 0) {
        throw 'CSV must include headers named URL and New base URL.'
    }

    $urlKey = [string]$urlColumn[0]
    $newBaseKey = [string]$newBaseColumn[0]

    $map = @{}
    foreach ($row in $rows) {
        $oldRaw = if ($row.PSObject.Properties[$urlKey]) { [string]$row.$urlKey } else { '' }
        $newRaw = if ($row.PSObject.Properties[$newBaseKey]) { [string]$row.$newBaseKey } else { '' }

        if ([string]::IsNullOrWhiteSpace($oldRaw) -and [string]::IsNullOrWhiteSpace($newRaw)) {
            continue
        }

        $oldBase = $oldRaw.TrimEnd('/')
        $newBase = $newRaw.TrimEnd('/')

        if ([string]::IsNullOrWhiteSpace($oldBase)) {
            throw 'CSV contains a row with a blank URL value.'
        }
        if ([string]::IsNullOrWhiteSpace($newBase)) {
            throw ("CSV row for URL '{0}' has a blank New base URL value." -f $oldBase)
        }
        if ($oldBase -notmatch '^https?://') {
            throw ("CSV URL value is not a valid absolute URL: {0}" -f $oldBase)
        }
        if ($newBase -notmatch '^https?://') {
            throw ("CSV New base URL value is not a valid absolute URL: {0}" -f $newBase)
        }

        if ($map.ContainsKey($oldBase) -and -not ([string]$map[$oldBase]).Equals($newBase, [System.StringComparison]::OrdinalIgnoreCase)) {
            throw ("CSV contains conflicting New base URL values for URL '{0}'." -f $oldBase)
        }

        $map[$oldBase] = $newBase
    }

    $missing = New-Object System.Collections.Generic.List[string]
    foreach ($sub in $PromptSubsites) {
        $oldBase = ([string]$sub.Url).TrimEnd('/')
        if (-not $map.ContainsKey($oldBase)) {
            $missing.Add($oldBase)
        }
    }

    if ($missing.Count -gt 0) {
        $preview = @($missing | Select-Object -First 8)
        $suffix = if ($missing.Count -gt 8) { ' ...' } else { '' }
        throw ("CSV is missing {0} required URL row(s). Missing examples: {1}{2}" -f $missing.Count, ($preview -join '; '), $suffix)
    }

    return $map
}

function Export-NewBaseUrlCsvTemplate {
    param(
        [Parameter(Mandatory = $true)][string]$CsvPath,
        [Parameter(Mandatory = $true)][array]$PromptSubsites
    )

    $rows = foreach ($sub in @($PromptSubsites | Sort-Object Url)) {
        [PSCustomObject][ordered]@{
            'URL' = ([string]$sub.Url).TrimEnd('/')
            'New base URL' = ''
        }
    }

    @($rows) | Export-Csv -Path $CsvPath -NoTypeInformation
}

function Get-CsvTemplateExportPath {
    param(
        [Parameter(Mandatory = $false)][string]$DefaultDirectory = '',
        [Parameter(Mandatory = $false)][string]$DefaultFileName = 'Get-FixItListURLsv5-UrlMappingTemplate.csv'
    )

    $selectedPath = ''

    try {
        Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop

        $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
        $saveDialog.Title = 'Save URL Mapping Template CSV'
        $saveDialog.Filter = 'CSV files (*.csv)|*.csv|All files (*.*)|*.*'
        $saveDialog.DefaultExt = 'csv'
        $saveDialog.AddExtension = $true
        $saveDialog.OverwritePrompt = $true
        $saveDialog.CheckPathExists = $true
        if (-not [string]::IsNullOrWhiteSpace($DefaultFileName)) {
            $saveDialog.FileName = $DefaultFileName
        }
        if (-not [string]::IsNullOrWhiteSpace($DefaultDirectory) -and (Test-Path -Path $DefaultDirectory)) {
            $saveDialog.InitialDirectory = $DefaultDirectory
        }

        $dialogResult = $saveDialog.ShowDialog()
        if ($dialogResult -eq [System.Windows.Forms.DialogResult]::OK -and -not [string]::IsNullOrWhiteSpace($saveDialog.FileName)) {
            $selectedPath = [string]$saveDialog.FileName
        }
    }
    catch {
        Write-Host "  Could not open file picker dialog. You can still enter a path manually."
    }

    while ([string]::IsNullOrWhiteSpace($selectedPath)) {
        $manualPath = Read-Host 'Enter full path where the CSV template should be exported'
        if ([string]::IsNullOrWhiteSpace($manualPath)) {
            Write-Host '  Template export path is required.'
            continue
        }

        $selectedPath = $manualPath
    }

    return $selectedPath
}

function Get-ExistingCsvSelectionPath {
    param(
        [Parameter(Mandatory = $false)][string]$DefaultDirectory = ''
    )

    try {
        Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop

        $openDialog = New-Object System.Windows.Forms.OpenFileDialog
        $openDialog.Title = 'Select URL Mapping CSV File'
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

    return ''
}

function Get-UrlMappingCsvPreferencePath {
    param([Parameter(Mandatory = $true)][string]$StateRootFolder)

    return (Join-Path -Path $StateRootFolder -ChildPath 'Get-FixItListURLsv5_UrlMappingCsvPreference.json')
}

function Get-UrlMappingCsvPreference {
    param([Parameter(Mandatory = $true)][string]$Path)

    if (-not (Test-Path -Path $Path)) {
        return [PSCustomObject]@{
            LastTemplatePath = ''
            LastCompletedCsvPath = ''
            UpdatedAt = ''
        }
    }

    try {
        $raw = Get-Content -Path $Path -Raw -Encoding UTF8 | ConvertFrom-Json
        return [PSCustomObject]@{
            LastTemplatePath = if ($raw.PSObject.Properties['LastTemplatePath']) { [string]$raw.LastTemplatePath } else { '' }
            LastCompletedCsvPath = if ($raw.PSObject.Properties['LastCompletedCsvPath']) { [string]$raw.LastCompletedCsvPath } else { '' }
            UpdatedAt = if ($raw.PSObject.Properties['UpdatedAt']) { [string]$raw.UpdatedAt } else { '' }
        }
    }
    catch {
        Write-Verbose "Could not load URL mapping CSV preference file from ${Path}: $($_.Exception.Message)"
        return [PSCustomObject]@{
            LastTemplatePath = ''
            LastCompletedCsvPath = ''
            UpdatedAt = ''
        }
    }
}

function Save-UrlMappingCsvPreference {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $false)][string]$TemplatePath = '',
        [Parameter(Mandatory = $false)][string]$CompletedCsvPath = ''
    )

    $current = Get-UrlMappingCsvPreference -Path $Path
    $lastTemplatePath = if ([string]::IsNullOrWhiteSpace($TemplatePath)) { [string]$current.LastTemplatePath } else { $TemplatePath }
    $lastCompletedCsvPath = if ([string]::IsNullOrWhiteSpace($CompletedCsvPath)) { [string]$current.LastCompletedCsvPath } else { $CompletedCsvPath }

    try {
        [ordered]@{
            LastTemplatePath = $lastTemplatePath
            LastCompletedCsvPath = $lastCompletedCsvPath
            UpdatedAt = (Get-Date).ToString('o')
        } | ConvertTo-Json -Depth 4 | Set-Content -Path $Path -Encoding UTF8 -Force
    }
    catch {
        Write-Verbose "Could not save URL mapping CSV preference file to ${Path}: $($_.Exception.Message)"
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

function Get-NoFindingsDocumentRow {
    return [PSCustomObject][ordered]@{
        'Link Type' = ''
        'Site Title' = ''
        'Document Title' = ''
        'Document URL' = ''
        'Document Folder' = ''
        'Document Folder URL' = ''
        'Broken URL Text' = ''
        'Broken Embedded URL' = 'No findings to report'
        'New URL' = ''
        'Created By' = ''
        'Last Modified By' = ''
        'Last Modified Date' = ''
        'Permissions' = ''
    }
}

function Get-NoFindingsOtherRow {
    return [PSCustomObject][ordered]@{
        'Link Type' = ''
        'Site Title' = ''
        'Content Type' = ''
        'Item ID' = ''
        'Item Title' = ''
        'Display Text' = ''
        'Item URL' = 'No findings to report'
        'Site Contents Folder' = ''
        'Site Contents Folder URL' = ''
        'New Item URL' = ''
        'Created By' = ''
        'Last Modified By' = ''
        'Last Modified Date' = ''
        'Permissions' = ''
    }
}

function Export-CsvRowWithRetry {
    param(
        [Parameter(Mandatory = $true)][string]$OutputPath,
        [Parameter(Mandatory = $true)][pscustomobject]$Row,
        [Parameter(Mandatory = $false)][int]$MaxAttempts = 5,
        [Parameter(Mandatory = $false)][int]$DelayMilliseconds = 500
    )

    for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
        try {
            $Row | Export-Csv -Path $OutputPath -NoTypeInformation -Append
            return
        }
        catch {
            # OneDrive/AV can transiently hold a lock right after a write; retry rather than fail the scan.
            $isLocked = ([string]$_.Exception.Message -match '(?i)being used by another process|cannot access the file')
            if ((-not $isLocked) -or ($attempt -eq $MaxAttempts)) {
                throw
            }
            Start-Sleep -Milliseconds ($DelayMilliseconds * $attempt)
        }
    }
}

function Write-DocumentFinding {
    param(
        [Parameter(Mandatory = $true)][string]$OutputPath,
        [Parameter(Mandatory = $true)][pscustomobject]$Row
    )

    Export-CsvRowWithRetry -OutputPath $OutputPath -Row $Row
}

function Write-OtherFinding {
    param(
        [Parameter(Mandatory = $true)][string]$OutputPath,
        [Parameter(Mandatory = $true)][pscustomobject]$Row
    )

    Export-CsvRowWithRetry -OutputPath $OutputPath -Row $Row
}

function Get-OverviewOnlyReportPath {
    param([Parameter(Mandatory = $true)][string]$SourceReportPath)

    $folder = Split-Path -Path $SourceReportPath -Parent
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($SourceReportPath)
    $ext = [System.IO.Path]::GetExtension($SourceReportPath)
    return (Join-Path -Path $folder -ChildPath ($baseName + '_OverviewOnly' + $ext))
}

function Get-NoFindingsDocumentOverviewRow {
    return [PSCustomObject][ordered]@{
        'Site Title' = ''
        'Document Title' = ''
        'Document URL' = 'No findings to report'
        'Document Folder' = ''
        'Document Folder URL' = ''
        'Broken Link Count' = 0
    }
}

function Get-NoFindingsOtherOverviewRow {
    return [PSCustomObject][ordered]@{
        'Site Title' = ''
        'Content Type' = ''
        'Item ID' = ''
        'Item Title' = ''
        'Site Contents Folder' = ''
        'Site Contents Folder URL' = 'No findings to report'
        'Broken Link Count' = 0
    }
}

function Export-OverviewOnlyReports {
    param(
        [Parameter(Mandatory = $true)][string]$DocumentReportPath,
        [Parameter(Mandatory = $false)][string]$OtherReportPath,
        [Parameter(Mandatory = $true)][bool]$IncludeOther
    )

    $documentOverviewPath = Get-OverviewOnlyReportPath -SourceReportPath $DocumentReportPath
    $otherOverviewPath = ''

    $docRows = @()
    if (Test-Path -Path $DocumentReportPath) {
        $docRows = @(Import-Csv -Path $DocumentReportPath)
    }

    $docFindingRows = @($docRows | Where-Object { $_.'Broken Embedded URL' -and $_.'Broken Embedded URL' -ne 'No findings to report' })
    if ($docFindingRows.Count -eq 0) {
        Get-NoFindingsDocumentOverviewRow | Export-Csv -Path $documentOverviewPath -NoTypeInformation
    }
    else {
        $docOverviewRows = foreach ($grp in ($docFindingRows | Group-Object 'Document URL')) {
            $first = $grp.Group[0]
            $brokenUrls = @($grp.Group | ForEach-Object { $_.'Broken Embedded URL' } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Sort-Object -Unique)
            [PSCustomObject][ordered]@{
                'Site Title' = $first.'Site Title'
                'Document Title' = $first.'Document Title'
                'Document URL' = $first.'Document URL'
                'Document Folder' = $first.'Document Folder'
                'Document Folder URL' = $first.'Document Folder URL'
                'Broken Link Count' = $brokenUrls.Count
            }
        }

        @($docOverviewRows | Sort-Object 'Site Title', 'Document Title', 'Document URL') | Export-Csv -Path $documentOverviewPath -NoTypeInformation
    }

    if ($IncludeOther -and -not [string]::IsNullOrWhiteSpace($OtherReportPath)) {
        $otherOverviewPath = Get-OverviewOnlyReportPath -SourceReportPath $OtherReportPath

        $otherRows = @()
        if (Test-Path -Path $OtherReportPath) {
            $otherRows = @(Import-Csv -Path $OtherReportPath)
        }

        $otherFindingRows = @($otherRows | Where-Object { $_.'Item URL' -and $_.'Item URL' -ne 'No findings to report' })
        if ($otherFindingRows.Count -eq 0) {
            Get-NoFindingsOtherOverviewRow | Export-Csv -Path $otherOverviewPath -NoTypeInformation
        }
        else {
            $otherOverviewRows = foreach ($grp in ($otherFindingRows | Group-Object 'Site Title', 'Content Type', 'Item ID', 'Item Title', 'Site Contents Folder', 'Site Contents Folder URL')) {
                $first = $grp.Group[0]
                $brokenUrls = @($grp.Group | ForEach-Object { $_.'Item URL' } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Sort-Object -Unique)
                [PSCustomObject][ordered]@{
                    'Site Title' = $first.'Site Title'
                    'Content Type' = $first.'Content Type'
                    'Item ID' = $first.'Item ID'
                    'Item Title' = $first.'Item Title'
                    'Site Contents Folder' = $first.'Site Contents Folder'
                    'Site Contents Folder URL' = $first.'Site Contents Folder URL'
                    'Broken Link Count' = $brokenUrls.Count
                }
            }

            @($otherOverviewRows | Sort-Object 'Site Title', 'Content Type', 'Item ID', 'Item Title') | Export-Csv -Path $otherOverviewPath -NoTypeInformation
        }
    }

    return [PSCustomObject]@{
        DocumentOverviewPath = $documentOverviewPath
        OtherOverviewPath = $otherOverviewPath
    }
}

function Save-CounterCheckpoint {
    # Persists in-memory extension counters to JSON so a mid-scan crash does not lose all per-site data.
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][hashtable]$DocScanned,
        [Parameter(Mandatory = $true)][hashtable]$DocSkippedForced,
        [Parameter(Mandatory = $true)][hashtable]$DocSkippedUnsupported,
        [Parameter(Mandatory = $true)][hashtable]$OtherScanned,
        [Parameter(Mandatory = $true)][hashtable]$OtherSkippedForced,
        [Parameter(Mandatory = $true)][hashtable]$OtherSkippedUnsupported
    )

    try {
        [PSCustomObject]@{
            SavedAt               = (Get-Date -Format 'o')
            DocScanned            = $DocScanned
            DocSkippedForced      = $DocSkippedForced
            DocSkippedUnsupported = $DocSkippedUnsupported
            OtherScanned          = $OtherScanned
            OtherSkippedForced    = $OtherSkippedForced
            OtherSkippedUnsupported = $OtherSkippedUnsupported
        } | ConvertTo-Json -Depth 5 | Set-Content -Path $Path -Encoding UTF8 -Force
    }
    catch {
        Write-Verbose "Could not save counter checkpoint to ${Path}: $($_.Exception.Message)"
    }
}

Install-RequiredModule -Name "PnP.PowerShell"

New-DirectoryIfMissing -Path $DocumentsOutputFolder
New-DirectoryIfMissing -Path $OtherOutputFolder
New-DirectoryIfMissing -Path $LogsOutputFolder
if ($UsePersistedExternalUrlCache) {
    New-DirectoryIfMissing -Path $StateRootFolder
}
if ($UsePersistedInventoryCache) {
    New-DirectoryIfMissing -Path $StateRootFolder
}

$RootUrl = $RootUrl.TrimEnd('/')
$AdminUrl = $AdminUrl.TrimEnd('/')

$adminConn = Connect-ValidatedPnP -Url $AdminUrl -ClientId $ClientId -AdminConnection
$permissionCheck = $null
$accountName = $null
try {
    $permissionCheck = Test-ConnectedAccountPermissionSet -Connection $adminConn -Url $AdminUrl -AdminConnection
}
catch {
    $permissionFailure = $_.Exception
    $permissionFailureMessage = if ($permissionFailure -and -not [string]::IsNullOrWhiteSpace($permissionFailure.Message)) {
        $permissionFailure.Message
    }
    elseif ($_.TargetObject -and $_.TargetObject.PSObject.Properties['Message']) {
        [string]$_.TargetObject.Message
    }
    else {
        $_.ToString()
    }

    Write-Warning "Initial permission/scope validation failed: $permissionFailureMessage"

    $requiredScopes = @('Sites.Read.All', 'User.Read', 'Directory.Read.All')
    if ($permissionFailure -and $permissionFailure.Data -is [System.Collections.IDictionary] -and $permissionFailure.Data.Contains('RequiredScopes') -and $permissionFailure.Data['RequiredScopes']) {
        $requiredScopes = @($permissionFailure.Data['RequiredScopes'])
    }
    elseif ($_.TargetObject -and $_.TargetObject.PSObject.Properties['RequiredScopes'] -and $_.TargetObject.RequiredScopes) {
        $requiredScopes = @($_.TargetObject.RequiredScopes)
    }

    $accountName = 'current account'
    $action = Show-PermissionReauthPrompt -AccountName $accountName -Message $permissionFailureMessage -RequiredScopes $requiredScopes
    if ($action -eq 'switch') {
        $adminConn = Connect-ValidatedPnP -Url $AdminUrl -ClientId $ClientId -AdminConnection
        $permissionCheck = Test-ConnectedAccountPermissionSet -Connection $adminConn -Url $AdminUrl -AdminConnection
    }
    else {
        throw $permissionFailureMessage
    }
}

if (-not $permissionCheck.SharePointValidated) {
    throw 'The connected account did not pass the required SharePoint permission validation.'
}

Write-Host "`nCollecting tenant sites (excluding OneDrive personal sites)..."
$tenantSiteResult = Invoke-PnPConnectionOperation -Url $AdminUrl -ClientId $ClientId -Connection $adminConn -AdminConnection -Label 'Get tenant sites' -MaxAttempts 4 -Operation {
    param($activeConnection)
    Get-PnPTenantSite -Connection $activeConnection -Detailed
}
$adminConn = $tenantSiteResult.Connection
$tenantSites = $tenantSiteResult.Result
$tenantSites = $tenantSites |
Where-Object {
    $_.Url -and
    $_.Url -notmatch '-my\.sharepoint\.com' -and
    $_.Template -notlike 'SPSPERS*' -and
    ($SkipSiteUrls -notcontains $_.Url.TrimEnd('/'))
} |
Sort-Object Url -Unique

if ($tenantSites.Count -eq 0) {
    throw "No eligible SharePoint site collections found (excluding OneDrive)."
}

$selectedScope = Select-SiteForScope -Sites $tenantSites
$selectedScopeUrl = $selectedScope.Url.TrimEnd('/')
Write-Host "`nSelected inventory scope: $selectedScopeUrl"

$scopeConn = Connect-ValidatedPnP -Url $selectedScopeUrl -ClientId $ClientId

Write-Host ("`nSubsite inventory under selected scope {0}:" -f $selectedScopeUrl)
$scopeSubsites = $null
try {
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
    $children = @($scopeChildResult.Result | Sort-Object Url)
    foreach ($child in $children) {
        $scopeSubsites.Add([PSCustomObject]@{ Url = $child.Url; Title = $child.Title })
    }
}
catch {
    throw "Could not enumerate subsites under ${selectedScopeUrl}: $($_.Exception.Message)"
}

if (-not $scopeSubsites -or $scopeSubsites.Count -eq 0) {
    throw "No sites were found under $selectedScopeUrl."
}

$rootHostForInventory = ([Uri]$RootUrl).Host.ToLowerInvariant()
$scannableScopeSubsites = New-Object System.Collections.Generic.List[Object]
$excludedScopeSubsites = New-Object System.Collections.Generic.List[Object]

$inventoryCacheHits = 0
$inventoryDeltaProbes = 0
$inventoryRemovedCount = 0
$inventoryCachePath = ''
$cachedInventoryByKey = @{}
$freshInventoryByKey = @{}
$inventoryEntriesForSave = New-Object System.Collections.Generic.List[Object]

if ($UsePersistedInventoryCache) {
    $inventoryCachePath = Get-InventorySubsiteCachePath -StateRoot $StateRootFolder -ScopeUrl $selectedScopeUrl -SchemaVersion $ScriptSchemaVersion
    $cachedInventoryByKey = Get-InventorySubsiteCache -Path $inventoryCachePath
}

foreach ($scopeSub in @($scopeSubsites | Sort-Object Url -Unique)) {
    $scopeSubUrl = ([string]$scopeSub.Url).TrimEnd('/')
    $scopeSubTitle = [string]$scopeSub.Title
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

    $inventoryEntriesForSave.Add([PSCustomObject]@{
            Url = $scopeSubUrl
            Title = $scopeSubTitle
            IsScannable = $isScannable
            Reason = $reason
        })

    if ($isScannable) {
        $scannableScopeSubsites.Add($scopeSub)
    }
    else {
        $excludedScopeSubsites.Add([PSCustomObject]@{
                Url = $scopeSubUrl
                Reason = $reason
            })
    }
}

if ($UsePersistedInventoryCache) {
    foreach ($cachedKey in $cachedInventoryByKey.Keys) {
        if (-not $freshInventoryByKey.ContainsKey([string]$cachedKey)) {
            $inventoryRemovedCount++
        }
    }

    Save-InventorySubsiteCache -Path $inventoryCachePath -Entries $inventoryEntriesForSave.ToArray() -ScopeUrl $selectedScopeUrl -SchemaVersion $ScriptSchemaVersion
    Write-Host ("Inventory cache summary: reused {0}, delta-probed {1}, removed from cache {2}." -f $inventoryCacheHits, $inventoryDeltaProbes, $inventoryRemovedCount)
}

if ($excludedScopeSubsites.Count -gt 0) {
    Write-Host ("Filtered out {0} subsite(s) that are not scannable by this script:" -f $excludedScopeSubsites.Count)
    foreach ($excluded in $excludedScopeSubsites) {
        Write-Host ("  - {0}  ({1})" -f $excluded.Url, $excluded.Reason)
    }
}

$scopeSubsites = $scannableScopeSubsites

if (-not $scopeSubsites -or $scopeSubsites.Count -eq 0) {
    throw "No scannable subsites were found under $selectedScopeUrl after inventory filtering."
}

for ($i = 0; $i -lt $scopeSubsites.Count; $i++) {
    $title = if ($scopeSubsites[$i].Title) { $scopeSubsites[$i].Title } else { '(no title)' }
    Write-Host ("  {0,4}. {1}  |  {2}" -f ($i + 1), $scopeSubsites[$i].Url, $title)
}

Write-Host "     0. ALL listed subsites"

$selectedIndices = $null
$selectedAllSubsites = $false
while (-not $selectedIndices) {
    $pickRaw = Read-Host "`nPick one or more subsites by number (comma-separated, e.g. 1,4,7). Use 0 for ALL subsites"
    if ([string]::IsNullOrWhiteSpace($pickRaw)) {
        Write-Host "  Selection is required."
        continue
    }

    $parts = @($pickRaw -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    $valid = $true
    $numbers = New-Object System.Collections.Generic.List[int]

    if ($parts.Count -eq 1 -and $parts[0] -eq '0') {
        $selectedAllSubsites = $true
        $selectedIndices = @(1..$scopeSubsites.Count)
        break
    }

    foreach ($p in $parts) {
        if ($p -notmatch '^\d+$') {
            $valid = $false
            break
        }
        $n = [int]$p
        if ($n -eq 0) {
            $valid = $false
            break
        }
        if ($n -lt 1 -or $n -gt $scopeSubsites.Count) {
            $valid = $false
            break
        }
        $numbers.Add($n)
    }

    if (-not $valid -or $numbers.Count -eq 0) {
        Write-Host "  Invalid selection. Use comma-separated numbers from the list."
        continue
    }

    $selectedIndices = @($numbers | Sort-Object -Unique)
}

$selectedSubsites = @()
foreach ($idx in $selectedIndices) {
    $selectedSubsites += $scopeSubsites[$idx - 1]
}

$urlMappings = New-Object System.Collections.Generic.List[object]
$mappingPromptSubsites = @($selectedSubsites | Sort-Object Url -Unique)
$newBaseByOldBase = @{}
$urlMappingCsvPreferencePath = Get-UrlMappingCsvPreferencePath -StateRootFolder $StateRootFolder
$urlMappingCsvPreference = Get-UrlMappingCsvPreference -Path $urlMappingCsvPreferencePath

if ($mappingPromptSubsites.Count -gt 0) {
    $useCsvMappings = Get-BooleanPrompt -Message 'Use a CSV file for URL mappings (recommended when mapping many subsites)?' -DefaultYes $true

    if ($useCsvMappings) {
        while ($true) {
            Write-Host ''
            Write-Host 'CSV mapping options:'
            Write-Host '  1. Export CSV template, fill it in, then import it'
            Write-Host '  2. Import an existing completed CSV now'
            Write-Host '  3. Switch to manual prompts instead'

            $choice = Read-Host 'Choose 1, 2, or 3'
            if ([string]::IsNullOrWhiteSpace($choice)) {
                $choice = '1'
            }

            if ($choice -eq '3') {
                $useCsvMappings = $false
                break
            }

            if ($choice -eq '1') {
                $templatePath = Get-CsvTemplateExportPath -DefaultDirectory $LogsOutputFolder -DefaultFileName 'Get-FixItListURLsv5-UrlMappingTemplate.csv'
                try {
                    Export-NewBaseUrlCsvTemplate -CsvPath $templatePath -PromptSubsites $mappingPromptSubsites
                    Write-Host ("  Template exported: {0}" -f $templatePath)
                    Save-UrlMappingCsvPreference -Path $urlMappingCsvPreferencePath -TemplatePath $templatePath
                    $urlMappingCsvPreference = Get-UrlMappingCsvPreference -Path $urlMappingCsvPreferencePath
                }
                catch {
                    Write-Warning ("Could not export CSV template: {0}" -f $_.Exception.Message)
                    continue
                }
            }
            elseif ($choice -ne '2') {
                Write-Host '  Please choose 1, 2, or 3.'
                continue
            }

            $csvPath = ''
            if ($choice -eq '1') {
                $cachedCsvPath = [string]$urlMappingCsvPreference.LastCompletedCsvPath
                if (-not [string]::IsNullOrWhiteSpace($cachedCsvPath) -and (Test-Path -Path $cachedCsvPath)) {
                    $reuse = Get-BooleanPrompt -Message ("Use last completed CSV file? {0}" -f $cachedCsvPath) -DefaultYes $true
                    if ($reuse) {
                        $csvPath = $cachedCsvPath
                    }
                }
            }

            if ([string]::IsNullOrWhiteSpace($csvPath)) {
                $dialogPath = Get-ExistingCsvSelectionPath -DefaultDirectory $LogsOutputFolder
                if (-not [string]::IsNullOrWhiteSpace($dialogPath)) {
                    $csvPath = $dialogPath
                }
            }

            if ([string]::IsNullOrWhiteSpace($csvPath)) {
                $cachedCsvPath = [string]$urlMappingCsvPreference.LastCompletedCsvPath
                if (-not [string]::IsNullOrWhiteSpace($cachedCsvPath) -and (Test-Path -Path $cachedCsvPath)) {
                    $useCached = Get-BooleanPrompt -Message ("Use last completed CSV file? {0}" -f $cachedCsvPath) -DefaultYes $true
                    if ($useCached) {
                        $csvPath = $cachedCsvPath
                    }
                }
            }

            if ([string]::IsNullOrWhiteSpace($csvPath)) {
                $csvPath = Read-Host 'Enter full path to completed URL mapping CSV'
            }

            if ([string]::IsNullOrWhiteSpace($csvPath)) {
                Write-Host '  CSV path is required.'
                continue
            }

            try {
                $newBaseByOldBase = Get-NewBaseUrlMapFromCsv -CsvPath $csvPath -PromptSubsites $mappingPromptSubsites
                Save-UrlMappingCsvPreference -Path $urlMappingCsvPreferencePath -CompletedCsvPath $csvPath
                $urlMappingCsvPreference = Get-UrlMappingCsvPreference -Path $urlMappingCsvPreferencePath
                break
            }
            catch {
                Write-Warning ("CSV mapping validation failed: {0}" -f $_.Exception.Message)
            }
        }
    }

    if (-not $useCsvMappings) {
        foreach ($sub in $mappingPromptSubsites) {
            $oldBase = $sub.Url.TrimEnd('/')
            $newBase = $null

            while (-not $newBase) {
                $inputNew = Read-Host ("New base URL for {0}" -f $oldBase)
                if ([string]::IsNullOrWhiteSpace($inputNew)) {
                    Write-Host "  New base URL is required."
                    continue
                }
                $candidate = $inputNew.TrimEnd('/')
                if ($candidate -notmatch '^https?://') {
                    Write-Host "  Enter a valid absolute URL (https://...)."
                    continue
                }
                $newBase = $candidate
            }

            $newBaseByOldBase[$oldBase] = $newBase
        }
    }
}

foreach ($sub in $mappingPromptSubsites) {
    $oldBase = $sub.Url.TrimEnd('/')
    if (-not $newBaseByOldBase.ContainsKey($oldBase)) {
        throw ("No mapping was provided for selected URL: {0}" -f $oldBase)
    }

    $newBase = [string]$newBaseByOldBase[$oldBase]
    $oldUri = [Uri]$oldBase
    $urlMappings.Add([PSCustomObject]@{
            OldBaseUrl = $oldBase
            NewBaseUrl = $newBase
            TenantBase = $oldUri.GetLeftPart([UriPartial]::Authority)
            OldBasePath = $oldUri.AbsolutePath.TrimEnd('/')
        })
}

$urlMappingsForMatch = @($urlMappings | Sort-Object { $_.OldBaseUrl.Length } -Descending)

$scanOther = Get-BooleanPrompt -Message "After documents, scan additional content (site pages/list text fields) to a separate CSV?" -DefaultYes $false

$scanDocumentContent = $false
if ($PSBoundParameters.ContainsKey('EnableDocumentContentScan')) {
    $scanDocumentContent = [bool]$EnableDocumentContentScan
}
else {
    $scanDocumentContent = Get-BooleanPrompt -Message "Document embedded-link scanning reads/downloads file content and may update access analytics/last-access signals. Continue document content scan?" -DefaultYes $false
}

$selectedSiteSlugParts = New-Object System.Collections.Generic.List[string]
foreach ($sub in $selectedSubsites) {
    $rawTitle = if (-not [string]::IsNullOrWhiteSpace($sub.Title)) { [string]$sub.Title } else { ([Uri]$sub.Url).AbsolutePath }
    $titleSlug = Get-SlugValue -Text $rawTitle
    if (-not [string]::IsNullOrWhiteSpace($titleSlug)) {
        $selectedSiteSlugParts.Add($titleSlug)
    }
}

$selectedSiteSlugRaw = (@($selectedSiteSlugParts | Select-Object -Unique) -join '-')
$selectedSiteSlug = Get-SafeFileSlug -Text $selectedSiteSlugRaw -MaxLength 72 -Fallback 'SelectedSites'

$scriptTitleSegment = 'Get-FixItListURLs'
$reportPrefixBase = "$selectedSiteSlug-$scriptTitleSegment"
$runStamp = Get-Date -Format "yyyyMMdd-HHmmss"
$logStamp = Get-Date -Format "yyyy-MM-dd_HHmmss"
$docReportPath = Join-Path -Path $DocumentsOutputFolder -ChildPath ("$reportPrefixBase-Documents-" + $runStamp + ".csv")
$otherReportPath = Join-Path -Path $OtherOutputFolder -ChildPath ("$reportPrefixBase-OtherContent-" + $runStamp + ".csv")
$extensionDetailsPath = Join-Path -Path $LogsOutputFolder -ChildPath ($logStamp + '_Get-FixItListURLsv5_ExtensionDetails.csv')
$counterCheckpointPath = Join-Path -Path $LogsOutputFolder -ChildPath ($logStamp + '_Get-FixItListURLsv5_CounterCheckpoint.json')
$externalUrlStatusCachePath = ''
if ($UsePersistedExternalUrlCache) {
    $externalUrlStatusCachePath = Get-ExternalUrlStatusCachePath -StateRoot $StateRootFolder -ScopeUrl $selectedScopeUrl -SchemaVersion $ScriptSchemaVersion
}

$siteOutputPlans = @{}
$siteOutputStats = @{}
if ($selectedAllSubsites) {
    $slugUsage = @{}
    foreach ($sub in @($mappingPromptSubsites | Sort-Object Url)) {
        $siteKey = ([string]$sub.Url).TrimEnd('/')
        $rawLabel = if (-not [string]::IsNullOrWhiteSpace([string]$sub.Title)) { [string]$sub.Title } else { ([Uri]$siteKey).AbsolutePath }
        $baseSlug = Get-SafeFileSlug -Text $rawLabel -MaxLength 56 -Fallback 'Site'
        $uniqueSlug = $baseSlug
        $suffix = 2
        while ($slugUsage.ContainsKey($uniqueSlug)) {
            $uniqueSlug = "${baseSlug}${suffix}"
            $suffix++
        }
        $slugUsage[$uniqueSlug] = $true

        $siteDocPath = Join-Path -Path $DocumentsOutputFolder -ChildPath ("{0}-Get-FixItListURLsv5-Documents-{1}.csv" -f $uniqueSlug, $runStamp)
        $siteOtherPath = Join-Path -Path $OtherOutputFolder -ChildPath ("{0}-Get-FixItListURLsv5-OtherContent-{1}.csv" -f $uniqueSlug, $runStamp)
        $siteExtPath = Join-Path -Path $LogsOutputFolder -ChildPath ("{0}_Get-FixItListURLsv5_{1}_ExtensionDetails.csv" -f $logStamp, $uniqueSlug)
        $siteCheckpointPath = Join-Path -Path $LogsOutputFolder -ChildPath ("{0}_Get-FixItListURLsv5_{1}_CounterCheckpoint.json" -f $logStamp, $uniqueSlug)

        $siteOutputPlans[$siteKey] = [PSCustomObject]@{
            SiteKey = $siteKey
            SiteTitle = [string]$sub.Title
            SiteSlug = $uniqueSlug
            DocumentReportPath = $siteDocPath
            OtherReportPath = $siteOtherPath
            ExtensionDetailsPath = $siteExtPath
            CounterCheckpointPath = $siteCheckpointPath
        }

        $siteOutputStats[$siteKey] = [PSCustomObject]@{
            DocumentRows = 0
            OtherRows = 0
        }
    }
}
else {
    $singleKey = '__ALL__'
    $siteOutputPlans[$singleKey] = [PSCustomObject]@{
        SiteKey = $singleKey
        SiteTitle = $selectedSiteSlug
        SiteSlug = $selectedSiteSlug
        DocumentReportPath = $docReportPath
        OtherReportPath = $otherReportPath
        ExtensionDetailsPath = $extensionDetailsPath
        CounterCheckpointPath = $counterCheckpointPath
    }
    $siteOutputStats[$singleKey] = [PSCustomObject]@{
        DocumentRows = 0
        OtherRows = 0
    }
}

Write-Host "`nAudit configuration:"
Write-Host "  Selected subsite mappings:"
foreach ($m in $urlMappings) {
    Write-Host ("    OLD: {0}" -f $m.OldBaseUrl)
    Write-Host ("    NEW: {0}" -f $m.NewBaseUrl)
}
Write-Host "  Documents report:     $docReportPath"
Write-Host "  Document content scan: $scanDocumentContent"
Write-Host "  Counter checkpoint:    $counterCheckpointPath"
Write-Host "  Persisted external URL cache: $UsePersistedExternalUrlCache"
if ($scanOther) {
    Write-Host "  Other report:         $otherReportPath"
}
if ($selectedAllSubsites) {
    Write-Host "  Per-site output splitting: enabled (All Sites mode)"
}

$stats = [PSCustomObject]@{
    SitesVisited = 0
    WebsVisited = 0
    LibrariesVisited = 0
    DocumentsRead = 0
    DocumentLinkRows = 0
    OtherLinkRows = 0
    DocumentsSkippedForced = 0
    DocumentsSkippedUnsupported = 0
    OtherFilesScanned = 0
    OtherFilesSkippedForced = 0
    OtherFilesSkippedUnsupported = 0
}

$docScannedByExt = @{}
$docSkippedForcedByExt = @{}
$docSkippedUnsupportedByExt = @{}
$otherScannedByExt = @{}
$otherSkippedForcedByExt = @{}
$otherSkippedUnsupportedByExt = @{}

$permissionCache = @{}
$folderNameCache = @{}
$externalUrlStatusCache = @{}
if ($UsePersistedExternalUrlCache -and -not [string]::IsNullOrWhiteSpace($externalUrlStatusCachePath)) {
    $externalUrlStatusCache = Get-ExternalUrlStatusCache -Path $externalUrlStatusCachePath
}
$docSeenRows = [System.Collections.Generic.HashSet[string]]::new()
$otherSeenRows = [System.Collections.Generic.HashSet[string]]::new()
$tenantHosts = @($tenantSites | ForEach-Object { try { ([Uri]$_.Url).Host.ToLower() } catch { $null } } | Where-Object { $_ } | Sort-Object -Unique)
$httpHandler = [System.Net.Http.HttpClientHandler]::new()
$httpHandler.AllowAutoRedirect = $true
$httpHandler.MaxAutomaticRedirections = 10
$httpClient = [System.Net.Http.HttpClient]::new($httpHandler)
$httpClient.Timeout = [TimeSpan]::FromSeconds(8)

$scanRootSubsites = Get-UniqueSubsiteScanRootSet -Subsites $selectedSubsites
$scanSiteCollections = @($scanRootSubsites | Sort-Object Url)

Write-Host "`nStarting selected-scope document scan (SharePoint only, no OneDrive)..."
if ($scanRootSubsites.Count -lt $selectedSubsites.Count) {
    Write-Host ("Collapsed overlapping selected subsites into {0} unique scan root(s)." -f $scanRootSubsites.Count)
}
if ($scanDocumentContent) {
foreach ($site in $scanSiteCollections) {
    $siteUrl = $site.Url.TrimEnd('/')
    $siteTenantBase = ([Uri]$siteUrl).GetLeftPart([UriPartial]::Authority)
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
        $siteRootResult = Invoke-PnPConnectionOperation -Url $siteUrl -ClientId $ClientId -Connection $siteConn -Label ('Get site root web: ' + $siteUrl) -MaxAttempts 4 -Operation {
            param($activeConnection)
            Get-PnPWeb -Connection $activeConnection -Includes Url
        }
        $siteConn = $siteRootResult.Connection
        $rootWeb = $siteRootResult.Result
        $webUrls.Add($rootWeb.Url)
        $siteSubwebResult = Invoke-PnPConnectionOperation -Url $siteUrl -ClientId $ClientId -Connection $siteConn -Label ('Get child webs: ' + $siteUrl) -MaxAttempts 4 -Operation {
            param($activeConnection)
            Get-PnPSubWeb -Connection $activeConnection -Recurse -Includes Url
        }
        $siteConn = $siteSubwebResult.Connection
        $subs = $siteSubwebResult.Result
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
            $webResult = Invoke-PnPConnectionOperation -Url $webUrl -ClientId $ClientId -Connection $webConn -Label ('Get web metadata: ' + $webUrl) -MaxAttempts 4 -Operation {
                param($activeConnection)
                Get-PnPWeb -Connection $activeConnection -Includes Id, Url, Title
            }
            $webConn = $webResult.Connection
            $web = $webResult.Result
        }
        catch {
            Write-Warning "  Could not read web metadata for $webUrl"
            continue
        }

        $lists = $null
        try {
            $listResult = Invoke-PnPConnectionOperation -Url $webUrl -ClientId $ClientId -Connection $webConn -Label ('Get lists: ' + $webUrl) -MaxAttempts 4 -Operation {
                param($activeConnection)
                Get-PnPList -Connection $activeConnection -Includes BaseTemplate, Hidden, Title, ItemCount, Id, HasUniqueRoleAssignments, RootFolder.ServerRelativeUrl
            }
            $webConn = $listResult.Connection
            $lists = $listResult.Result
        }
        catch {
            Write-Warning "  Could not read lists for $webUrl"
            continue
        }

        $libraryCounterInWeb = 0
        foreach ($list in $lists) {
            if ($list.Hidden) { continue }
            if (Test-IsRecycleBinContainer -ListTitle $list.Title -RootFolderPath $list.RootFolder.ServerRelativeUrl) { continue }
            if ($SkipListTitles -contains $list.Title) { continue }
            if ($list.ItemCount -eq 0) { continue }
            if ($list.BaseTemplate -ne 101) { continue }

            $libraryCounterInWeb++
            if ($libraryCounterInWeb % 5 -eq 0) {
                try {
                    $webConn = Connect-ValidatedPnP -Url $webUrl -ClientId $ClientId
                    $webRefreshResult = Invoke-PnPConnectionOperation -Url $webUrl -ClientId $ClientId -Connection $webConn -Label ('Refresh web metadata: ' + $webUrl) -MaxAttempts 4 -Operation {
                        param($activeConnection)
                        Get-PnPWeb -Connection $activeConnection -Includes Id, Url, Title
                    }
                    $webConn = $webRefreshResult.Connection
                    $web = $webRefreshResult.Result
                }
                catch {
                    Write-Warning "    Could not refresh web connection; continuing with current connection."
                }
            }

            $stats.LibrariesVisited++
            Write-Host "    Library: $($list.Title)"

            $libraryRootPath = $list.RootFolder.ServerRelativeUrl.TrimEnd('/')

            $fields = @('FileRef', 'FileLeafRef', 'FileDirRef', 'UniqueId', 'FSObjType', 'Author', 'Editor', 'Modified', 'ID')
            $items = $null
            $retryState = @{
                WebConn = $webConn
                Web = $web
            }

            try {
                $items = Invoke-WithRetry -Label ("Get list items: " + $list.Title) -MaxAttempts 4 -OnTimeout {
                    $retryState.WebConn = Update-PnPConnection -Url $webUrl -ClientId $ClientId
                } -Operation {
                    Get-PnPListItem -List $list.Title -Connection $retryState.WebConn -PageSize 500 -Fields $fields
                }

                $webConn = $retryState.WebConn
                $web = $retryState.Web
            }
            catch {
                $message = [string]$_.Exception.Message
                if (Test-IsAuthFailureMessage -Message $message) {
                    Write-Warning "    Authentication issue while reading list items for $($list.Title); prompting for a fresh sign-in."
                    try {
                        $retryState.WebConn = Connect-ValidatedPnP -Url $webUrl -ClientId $ClientId
                        $retryWebResult = Invoke-PnPConnectionOperation -Url $webUrl -ClientId $ClientId -Connection $retryState.WebConn -Label ('Re-auth web metadata: ' + $webUrl) -MaxAttempts 4 -Operation {
                            param($activeConnection)
                            Get-PnPWeb -Connection $activeConnection -Includes Id, Url, Title
                        }
                        $retryState.WebConn = $retryWebResult.Connection
                        $retryState.Web = $retryWebResult.Result
                        $retryItemResult = Invoke-PnPConnectionOperation -Url $webUrl -ClientId $ClientId -Connection $retryState.WebConn -Label ('Re-auth list items: ' + $list.Title) -MaxAttempts 4 -Operation {
                            param($activeConnection)
                            Get-PnPListItem -List $list.Title -Connection $activeConnection -PageSize 500 -Fields $fields
                        }
                        $retryState.WebConn = $retryItemResult.Connection
                        $items = $retryItemResult.Result
                        $webConn = $retryState.WebConn
                        $web = $retryState.Web
                    }
                    catch {
                        Write-Warning "    Could not read items for library $($list.Title) after re-authentication: $($_.Exception.Message)"
                        continue
                    }
                }
                else {
                    Write-Warning "    Could not read items for library $($list.Title): $($_.Exception.Message)"
                    continue
                }
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

                $ext = Get-NormalizedExtension -FileName $fileLeaf
                if ($AlwaysSkipExtensions -contains $ext) {
                    $stats.DocumentsSkippedForced++
                    Add-ExtensionCount -Bucket $docSkippedForcedByExt -Extension $ext
                    continue
                }
                if (-not ($DocumentScannableExtensions -contains $ext)) {
                    $stats.DocumentsSkippedUnsupported++
                    Add-ExtensionCount -Bucket $docSkippedUnsupportedByExt -Extension $ext
                    continue
                }
                Add-ExtensionCount -Bucket $docScannedByExt -Extension $ext

                $docUrl = Get-DocumentBrowserUrl -TenantBase $siteTenantBase -FileRef $fileRef -Extension $ext -UniqueId $item.FieldValues['UniqueId']
                $folderUrl = Get-FolderBrowserUrl -TenantBase $siteTenantBase -LibraryRootPath $libraryRootPath -ListId $list.Id -FileDirRef $item.FieldValues['FileDirRef']
                $folderName = Get-FolderDisplayName -List $list -FolderPath $item.FieldValues['FileDirRef'] -Connection $webConn -ConnectionUrl $webUrl -ClientId $ClientId -Cache $folderNameCache

                $authorValue = $item.FieldValues['Author']
                $createdBy = if ($authorValue -and $authorValue.LookupValue) { $authorValue.LookupValue } else { '' }

                $editorValue = $item.FieldValues['Editor']
                $lastModifiedBy = if ($editorValue -and $editorValue.LookupValue) { $editorValue.LookupValue } else { '' }

                $topLevelPermissions = ''
                try {
                    $topLevelPermissions = Get-EffectiveTopLevelPermission -Item $item -List $list -Web $web -Connection $webConn -ConnectionUrl $webUrl -ClientId $ClientId -Cache $permissionCache
                }
                catch {
                    $topLevelPermissions = "(could not read permissions: $($_.Exception.Message))"
                }

                $refs = @()
                try {
                    $fileRetryState = @{ WebConn = $webConn; Web = $web }
                    $ms = Invoke-WithRetry -Label ("Get file stream: " + $fileRef) -MaxAttempts 4 -OnTimeout {
                        $fileRetryState.WebConn = Update-PnPConnection -Url $webUrl -ClientId $ClientId
                    } -Operation {
                        Get-PnPFile -Url $fileRef -AsMemoryStream -Connection $fileRetryState.WebConn
                    }

                    $webConn = $fileRetryState.WebConn
                    $web = $fileRetryState.Web

                    if ($ext -eq '.pdf') {
                        $refs = Get-PdfReference -Bytes $ms.ToArray()
                    }
                    elseif ($OpenXmlExtensions -contains $ext) {
                        $refs = Get-OfficeDocReference -Stream $ms
                    }
                    else {
                        $refs = Get-TextReferenceFromByteContent -Bytes $ms.ToArray() -Kind 'plaintext-doc'
                    }
                }
                catch {
                    $message = [string]$_.Exception.Message
                    if (Test-IsAuthFailureMessage -Message $message) {
                        Write-Warning "    Authentication issue while reading file content for $fileRef; prompting for a fresh sign-in."
                        try {
                            $fileRetryState.WebConn = Connect-ValidatedPnP -Url $webUrl -ClientId $ClientId
                            $fileRetryWebResult = Invoke-PnPConnectionOperation -Url $webUrl -ClientId $ClientId -Connection $fileRetryState.WebConn -Label ('Re-auth web metadata: ' + $webUrl) -MaxAttempts 4 -Operation {
                                param($activeConnection)
                                Get-PnPWeb -Connection $activeConnection -Includes Id, Url, Title
                            }
                            $fileRetryState.WebConn = $fileRetryWebResult.Connection
                            $fileRetryState.Web = $fileRetryWebResult.Result
                            $fileStreamResult = Invoke-PnPConnectionOperation -Url $webUrl -ClientId $ClientId -Connection $fileRetryState.WebConn -Label ('Re-auth file stream: ' + $fileRef) -MaxAttempts 4 -Operation {
                                param($activeConnection)
                                Get-PnPFile -Url $fileRef -AsMemoryStream -Connection $activeConnection
                            }
                            $fileRetryState.WebConn = $fileStreamResult.Connection
                            $ms = $fileStreamResult.Result
                            $webConn = $fileRetryState.WebConn
                            $web = $fileRetryState.Web

                            if ($ext -eq '.pdf') {
                                $refs = Get-PdfReference -Bytes $ms.ToArray()
                            }
                            elseif ($OpenXmlExtensions -contains $ext) {
                                $refs = Get-OfficeDocReference -Stream $ms
                            }
                            else {
                                $refs = Get-TextReferenceFromByteContent -Bytes $ms.ToArray() -Kind 'plaintext-doc'
                            }
                        }
                        catch {
                            continue
                        }
                    }
                    else {
                        continue
                    }
                }

                foreach ($ref in $refs) {
                    if (Test-IsSuppressedDocumentReference -Reference $ref) {
                        continue
                    }

                    $cleanEmbedded = Get-CleanedValue -Value $ref.Value
                    $matched = Get-UrlMappingMatch -Value $ref.Value -Mappings $urlMappingsForMatch
                    $sourceDocUrl = $siteTenantBase + $fileRef
                    $sourceOnSelectedBase = Test-IsSourceUrlOnSelectedBase -SourceUrl $sourceDocUrl -Mappings $urlMappingsForMatch

                    if ((-not $matched) -and (Test-IsNonActionableUrl -Url $cleanEmbedded)) {
                        continue
                    }

                    $externalBroken = $false
                    if (-not $matched -and $sourceOnSelectedBase) {
                        $externalBroken = Test-IsExternalBrokenUrl -Url $ref.Value -TenantHosts $tenantHosts -Cache $externalUrlStatusCache -HttpClient $httpClient
                    }

                    if ((-not $matched) -and (-not $externalBroken)) {
                        continue
                    }

                    $siteTitle = if ([string]::IsNullOrWhiteSpace($web.Title)) { $webUrl } else { $web.Title }
                    $modifiedValue = $item.FieldValues['Modified']
                    $modifiedText = Format-ReportDatePacific -Value $modifiedValue
                    $brokenEmbeddedUrl = if ($matched) { $matched.OldUrl } else { $cleanEmbedded }
                    $brokenUrlText = if (-not [string]::IsNullOrWhiteSpace([string]$ref.Text)) { [string]$ref.Text } else { $cleanEmbedded }
                    $newUrlValue = if ($matched) { $matched.NewUrl } else { '' }
                    $linkType = Get-LinkTypeLabel -Url $ref.Value -TenantHosts $tenantHosts

                    $targetPlanKey = '__ALL__'
                    if ($selectedAllSubsites) {
                        $targetPlanKey = Get-SiteKeyForSourceUrl -SourceUrl $sourceDocUrl -Mappings $urlMappingsForMatch
                        if ([string]::IsNullOrWhiteSpace($targetPlanKey) -or -not $siteOutputPlans.ContainsKey($targetPlanKey)) {
                            continue
                        }
                    }

                    $row = [PSCustomObject][ordered]@{
                        'Link Type' = $linkType
                        'Site Title' = $siteTitle
                        'Document Title' = $fileLeaf
                        'Document URL' = $docUrl
                        'Document Folder' = $folderName
                        'Document Folder URL' = $folderUrl
                        'Broken URL Text' = $brokenUrlText
                        'Broken Embedded URL' = $brokenEmbeddedUrl
                        'New URL' = $newUrlValue
                        'Created By' = $createdBy
                        'Last Modified By' = $lastModifiedBy
                        'Last Modified Date' = $modifiedText
                        'Permissions' = $topLevelPermissions
                    }

                    if (-not (Test-IsExactRowDuplicate -Row $row -SeenRows $docSeenRows)) {
                        Write-DocumentFinding -OutputPath $siteOutputPlans[$targetPlanKey].DocumentReportPath -Row $row
                        $siteOutputStats[$targetPlanKey].DocumentRows++
                        $stats.DocumentLinkRows++
                    }
                }
            }
        }
    }
    if ($selectedAllSubsites) {
        foreach ($plan in $siteOutputPlans.Values) {
            Save-CounterCheckpoint -Path $plan.CounterCheckpointPath -DocScanned $docScannedByExt -DocSkippedForced $docSkippedForcedByExt -DocSkippedUnsupported $docSkippedUnsupportedByExt -OtherScanned $otherScannedByExt -OtherSkippedForced $otherSkippedForcedByExt -OtherSkippedUnsupported $otherSkippedUnsupportedByExt
        }
    }
    else {
        Save-CounterCheckpoint -Path $counterCheckpointPath -DocScanned $docScannedByExt -DocSkippedForced $docSkippedForcedByExt -DocSkippedUnsupported $docSkippedUnsupportedByExt -OtherScanned $otherScannedByExt -OtherSkippedForced $otherSkippedForcedByExt -OtherSkippedUnsupported $otherSkippedUnsupportedByExt
    }
}
}
else {
    Write-Warning "Document content scan was skipped to preserve last-access signal integrity."
}

if ($scanOther) {
    Write-Host "`nStarting optional scan for additional content..."

    foreach ($site in $scanSiteCollections) {
        $siteUrl = $site.Url.TrimEnd('/')
        $siteTenantBase = ([Uri]$siteUrl).GetLeftPart([UriPartial]::Authority)
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
            $otherRootResult = Invoke-PnPConnectionOperation -Url $siteUrl -ClientId $ClientId -Connection $siteConn -Label ('Get site root web: ' + $siteUrl) -MaxAttempts 4 -Operation {
                param($activeConnection)
                Get-PnPWeb -Connection $activeConnection -Includes Url, Title
            }
            $siteConn = $otherRootResult.Connection
            $rootWeb = $otherRootResult.Result
            $webUrls.Add($rootWeb.Url)
            $otherSubwebResult = Invoke-PnPConnectionOperation -Url $siteUrl -ClientId $ClientId -Connection $siteConn -Label ('Get child webs: ' + $siteUrl) -MaxAttempts 4 -Operation {
                param($activeConnection)
                Get-PnPSubWeb -Connection $activeConnection -Recurse -Includes Url
            }
            $siteConn = $otherSubwebResult.Connection
            $subs = $otherSubwebResult.Result
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

            $web = $null
            try {
                $otherWebResult = Invoke-PnPConnectionOperation -Url $webUrl -ClientId $ClientId -Connection $webConn -Label ('Get web metadata: ' + $webUrl) -MaxAttempts 4 -Operation {
                    param($activeConnection)
                    Get-PnPWeb -Connection $activeConnection -Includes Id, Url, Title
                }
                $webConn = $otherWebResult.Connection
                $web = $otherWebResult.Result
            }
            catch {
                continue
            }

            $lists = $null
            try {
                $otherListResult = Invoke-PnPConnectionOperation -Url $webUrl -ClientId $ClientId -Connection $webConn -Label ('Get lists: ' + $webUrl) -MaxAttempts 4 -Operation {
                    param($activeConnection)
                    Get-PnPList -Connection $activeConnection -Includes BaseTemplate, Hidden, Title, ItemCount, DefaultDisplayFormUrl, RootFolder.ServerRelativeUrl
                }
                $webConn = $otherListResult.Connection
                $lists = $otherListResult.Result
            }
            catch {
                continue
            }

            $otherLibraryCounterInWeb = 0
            foreach ($list in $lists) {
                if ($list.Hidden) { continue }
                if (Test-IsRecycleBinContainer -ListTitle $list.Title -RootFolderPath $list.RootFolder.ServerRelativeUrl) { continue }
                if ($list.ItemCount -eq 0) { continue }
                if ($list.BaseTemplate -eq 101) { continue }

                $otherLibraryCounterInWeb++
                if ($otherLibraryCounterInWeb % 5 -eq 0) {
                    try {
                        $webConn = Connect-ValidatedPnP -Url $webUrl -ClientId $ClientId
                        $webRefreshResult = Invoke-PnPConnectionOperation -Url $webUrl -ClientId $ClientId -Connection $webConn -Label ('Refresh web metadata: ' + $webUrl) -MaxAttempts 4 -Operation {
                            param($activeConnection)
                            Get-PnPWeb -Connection $activeConnection -Includes Id, Url, Title
                        }
                        $webConn = $webRefreshResult.Connection
                    }
                    catch {
                        Write-Verbose "Could not refresh non-document scan web connection for $webUrl; continuing with current connection."
                    }
                }

                $fields = $null
                try {
                    $fieldResult = Invoke-PnPConnectionOperation -Url $webUrl -ClientId $ClientId -Connection $webConn -Label ('Get fields: ' + $list.Title) -MaxAttempts 4 -Operation {
                        param($activeConnection)
                        Get-PnPField -List $list.Title -Connection $activeConnection
                    }
                    $webConn = $fieldResult.Connection
                    $fields = $fieldResult.Result
                }
                catch {
                    continue
                }

                $textFields = @($fields | Where-Object {
                        -not $_.Hidden -and -not $_.ReadOnlyField -and $_.TypeAsString -in @('Text', 'Note', 'URL')
                    })

                $items = $null
                try {
                    $otherRetryState = @{ WebConn = $webConn }
                    $items = Invoke-WithRetry -Label ("Get list items: " + $list.Title) -MaxAttempts 4 -OnTimeout {
                        $otherRetryState.WebConn = Update-PnPConnection -Url $webUrl -ClientId $ClientId
                    } -Operation {
                        Get-PnPListItem -List $list.Title -Connection $otherRetryState.WebConn -PageSize 200
                    }

                    $webConn = $otherRetryState.WebConn
                }
                catch {
                    $message = [string]$_.Exception.Message
                    if (Test-IsAuthFailureMessage -Message $message) {
                        Write-Warning "    Authentication issue while reading other-content list items for $($list.Title); prompting for a fresh sign-in."
                        try {
                            $otherRetryState.WebConn = Connect-ValidatedPnP -Url $webUrl -ClientId $ClientId
                            $otherItemResult = Invoke-PnPConnectionOperation -Url $webUrl -ClientId $ClientId -Connection $otherRetryState.WebConn -Label ('Re-auth list items: ' + $list.Title) -MaxAttempts 4 -Operation {
                                param($activeConnection)
                                Get-PnPListItem -List $list.Title -Connection $activeConnection -PageSize 200
                            }
                            $otherRetryState.WebConn = $otherItemResult.Connection
                            $items = $otherItemResult.Result
                            $webConn = $otherRetryState.WebConn
                        }
                        catch {
                            continue
                        }
                    }
                    else {
                        continue
                    }
                }

                foreach ($item in $items) {
                    $itemId = $item.Id
                    $itemTitle = [string]$item.FieldValues['Title']
                    $itemFileLeaf = if ($item.FieldValues.ContainsKey('FileLeafRef')) { [string]$item.FieldValues['FileLeafRef'] } else { '' }
                    $itemFileRef = if ($item.FieldValues.ContainsKey('FileRef')) { [string]$item.FieldValues['FileRef'] } else { '' }
                    $itemExt = if (-not [string]::IsNullOrWhiteSpace($itemFileLeaf)) { Get-NormalizedExtension -FileName $itemFileLeaf } else { '' }

                    if (-not [string]::IsNullOrWhiteSpace($itemExt) -and ($AlwaysSkipExtensions -contains $itemExt)) {
                        $stats.OtherFilesSkippedForced++
                        Add-ExtensionCount -Bucket $otherSkippedForcedByExt -Extension $itemExt
                        continue
                    }

                    $siteContentsFolderPath = if ($item.FieldValues.ContainsKey('FileDirRef') -and $item.FieldValues['FileDirRef']) { [string]$item.FieldValues['FileDirRef'] } else { [string]$list.RootFolder.ServerRelativeUrl }

                    $siteContentsFolder = $list.Title
                    if (-not [string]::IsNullOrWhiteSpace($siteContentsFolderPath)) {
                        try {
                            $siteContentsFolder = Get-FolderDisplayName -List $list -FolderPath $siteContentsFolderPath -Connection $webConn -ConnectionUrl $webUrl -ClientId $ClientId -Cache $folderNameCache
                        }
                        catch {
                            $siteContentsFolder = $list.Title
                        }
                    }

                    $siteContentsFolderUrl = ''
                    try {
                        $siteContentsFolderUrl = Get-FolderBrowserUrl -TenantBase $siteTenantBase -LibraryRootPath $list.RootFolder.ServerRelativeUrl.TrimEnd('/') -ListId $list.Id -FileDirRef $siteContentsFolderPath
                    }
                    catch {
                        $siteContentsFolderUrl = ''
                    }

                    $contentType = [string]$item.FieldValues['ContentType']
                    if ([string]::IsNullOrWhiteSpace($contentType)) {
                        $contentType = if ($list.BaseTemplate -eq 119) { 'Page' } else { 'List Item' }
                    }

                    $siteTitle = if ([string]::IsNullOrWhiteSpace($web.Title)) { $webUrl } else { $web.Title }
                    $authorValue = $item.FieldValues['Author']
                    $createdBy = if ($authorValue -and $authorValue.LookupValue) { $authorValue.LookupValue } else { '' }
                    $editorValue = $item.FieldValues['Editor']
                    $lastModifiedBy = if ($editorValue -and $editorValue.LookupValue) { $editorValue.LookupValue } else { '' }
                    $modifiedValue = $item.FieldValues['Modified']
                    $modifiedText = Format-ReportDatePacific -Value $modifiedValue
                    $otherTopLevelPermissions = ''

                    try {
                        $otherTopLevelPermissions = Get-EffectiveTopLevelPermission -Item $item -List $list -Web $web -Connection $webConn -ConnectionUrl $webUrl -ClientId $ClientId -Cache $permissionCache
                    }
                    catch {
                        $otherTopLevelPermissions = "(could not read permissions: $($_.Exception.Message))"
                    }

                    if (-not [string]::IsNullOrWhiteSpace($itemExt)) {
                        if ($OtherFileScannableExtensions -contains $itemExt -and -not [string]::IsNullOrWhiteSpace($itemFileRef)) {
                            $stats.OtherFilesScanned++
                            Add-ExtensionCount -Bucket $otherScannedByExt -Extension $itemExt

                            try {
                                $otherFileRetryState = @{ WebConn = $webConn }
                                $ms = Invoke-WithRetry -Label ("Get non-document file stream: " + $itemFileRef) -MaxAttempts 4 -OnTimeout {
                                    $otherFileRetryState.WebConn = Update-PnPConnection -Url $webUrl -ClientId $ClientId
                                } -Operation {
                                    Get-PnPFile -Url $itemFileRef -AsMemoryStream -Connection $otherFileRetryState.WebConn
                                }

                                $webConn = $otherFileRetryState.WebConn

                                $fileRefs = Get-TextReferenceFromByteContent -Bytes $ms.ToArray() -Kind 'plaintext-other'
                                foreach ($fileRef in $fileRefs) {
                                    $url = $fileRef.Value
                                    $matched = Get-UrlMappingMatch -Value $url -Mappings $urlMappingsForMatch
                                    $sourceOnSelectedBase = Test-IsSourceUrlOnSelectedBase -SourceUrl ($siteTenantBase + $itemFileRef) -Mappings $urlMappingsForMatch

                                    if ((-not $matched) -and (Test-IsNonActionableUrl -Url $url)) {
                                        continue
                                    }

                                    $externalBroken = $false
                                    if (-not $matched -and $sourceOnSelectedBase) {
                                        $externalBroken = Test-IsExternalBrokenUrl -Url $url -TenantHosts $tenantHosts -Cache $externalUrlStatusCache -HttpClient $httpClient
                                    }

                                    if ((-not $matched) -and (-not $externalBroken)) {
                                        continue
                                    }

                                    $newUrlValue = if ($matched) { $matched.NewUrl } else { '' }
                                    $displayText = $siteContentsFolder
                                    $linkType = Get-LinkTypeLabel -Url $url -TenantHosts $tenantHosts

                                    $targetPlanKey = '__ALL__'
                                    if ($selectedAllSubsites) {
                                        $targetPlanKey = Get-SiteKeyForSourceUrl -SourceUrl ($siteTenantBase + $itemFileRef) -Mappings $urlMappingsForMatch
                                        if ([string]::IsNullOrWhiteSpace($targetPlanKey) -or -not $siteOutputPlans.ContainsKey($targetPlanKey)) {
                                            continue
                                        }
                                    }

                                    $row = [PSCustomObject][ordered]@{
                                        'Link Type' = $linkType
                                        'Site Title' = $siteTitle
                                        'Content Type' = $contentType
                                        'Item ID' = $itemId
                                        'Item Title' = $itemTitle
                                        'Display Text' = $displayText
                                        'Item URL' = $url
                                        'Site Contents Folder' = $siteContentsFolder
                                        'Site Contents Folder URL' = $siteContentsFolderUrl
                                        'New Item URL' = $newUrlValue
                                        'Created By' = $createdBy
                                        'Last Modified By' = $lastModifiedBy
                                        'Last Modified Date' = $modifiedText
                                        'Permissions' = $otherTopLevelPermissions
                                    }

                                    if (-not (Test-IsExactRowDuplicate -Row $row -SeenRows $otherSeenRows)) {
                                        Write-OtherFinding -OutputPath $siteOutputPlans[$targetPlanKey].OtherReportPath -Row $row
                                        $siteOutputStats[$targetPlanKey].OtherRows++
                                        $stats.OtherLinkRows++
                                    }
                                }
                            }
                            catch {
                                $message = [string]$_.Exception.Message
                                if (Test-IsAuthFailureMessage -Message $message) {
                                    try {
                                        $otherFileRetryState.WebConn = Connect-ValidatedPnP -Url $webUrl -ClientId $ClientId
                                        $otherFileResult = Invoke-PnPConnectionOperation -Url $webUrl -ClientId $ClientId -Connection $otherFileRetryState.WebConn -Label ('Re-auth file stream: ' + $itemFileRef) -MaxAttempts 4 -Operation {
                                            param($activeConnection)
                                            Get-PnPFile -Url $itemFileRef -AsMemoryStream -Connection $activeConnection
                                        }

                                        $otherFileRetryState.WebConn = $otherFileResult.Connection
                                        $ms = $otherFileResult.Result

                                        $webConn = $otherFileRetryState.WebConn
                                        $fileRefs = Get-TextReferenceFromByteContent -Bytes $ms.ToArray() -Kind 'plaintext-other'
                                        foreach ($fileRef in $fileRefs) {
                                            $url = $fileRef.Value
                                            $matched = Get-UrlMappingMatch -Value $url -Mappings $urlMappingsForMatch
                                            $sourceOnSelectedBase = Test-IsSourceUrlOnSelectedBase -SourceUrl ($siteTenantBase + $itemFileRef) -Mappings $urlMappingsForMatch

                                            if ((-not $matched) -and (Test-IsNonActionableUrl -Url $url)) {
                                                continue
                                            }

                                            $externalBroken = $false
                                            if (-not $matched -and $sourceOnSelectedBase) {
                                                $externalBroken = Test-IsExternalBrokenUrl -Url $url -TenantHosts $tenantHosts -Cache $externalUrlStatusCache -HttpClient $httpClient
                                            }

                                            if ((-not $matched) -and (-not $externalBroken)) {
                                                continue
                                            }

                                            $newUrlValue = if ($matched) { $matched.NewUrl } else { '' }
                                            $displayText = $siteContentsFolder
                                            $linkType = Get-LinkTypeLabel -Url $url -TenantHosts $tenantHosts

                                            $targetPlanKey = '__ALL__'
                                            if ($selectedAllSubsites) {
                                                $targetPlanKey = Get-SiteKeyForSourceUrl -SourceUrl ($siteTenantBase + $itemFileRef) -Mappings $urlMappingsForMatch
                                                if ([string]::IsNullOrWhiteSpace($targetPlanKey) -or -not $siteOutputPlans.ContainsKey($targetPlanKey)) {
                                                    continue
                                                }
                                            }

                                            $row = [PSCustomObject][ordered]@{
                                                'Link Type' = $linkType
                                                'Site Title' = $siteTitle
                                                'Content Type' = $contentType
                                                'Item ID' = $itemId
                                                'Item Title' = $itemTitle
                                                'Display Text' = $displayText
                                                'Item URL' = $url
                                                'Site Contents Folder' = $siteContentsFolder
                                                'Site Contents Folder URL' = $siteContentsFolderUrl
                                                'New Item URL' = $newUrlValue
                                                'Created By' = $createdBy
                                                'Last Modified By' = $lastModifiedBy
                                                'Last Modified Date' = $modifiedText
                                                'Permissions' = $otherTopLevelPermissions
                                            }

                                            if (-not (Test-IsExactRowDuplicate -Row $row -SeenRows $otherSeenRows)) {
                                                Write-OtherFinding -OutputPath $siteOutputPlans[$targetPlanKey].OtherReportPath -Row $row
                                                $siteOutputStats[$targetPlanKey].OtherRows++
                                                $stats.OtherLinkRows++
                                            }
                                        }
                                    }
                                    catch {
                                        Write-Verbose "Could not parse file content for non-document file $itemFileRef"
                                    }
                                }
                                else {
                                    Write-Verbose "Could not parse file content for non-document file $itemFileRef"
                                }
                            }
                        }
                        elseif (-not ($OtherFileScannableExtensions -contains $itemExt)) {
                            $stats.OtherFilesSkippedUnsupported++
                            Add-ExtensionCount -Bucket $otherSkippedUnsupportedByExt -Extension $itemExt
                        }
                    }

                    if ($textFields.Count -eq 0) {
                        continue
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
                                $matched = Get-UrlMappingMatch -Value $url -Mappings $urlMappingsForMatch
                                $sourceUrlForItem = if (-not [string]::IsNullOrWhiteSpace($itemFileRef)) { $siteTenantBase + $itemFileRef } else { $webUrl }
                                $sourceOnSelectedBase = Test-IsSourceUrlOnSelectedBase -SourceUrl $sourceUrlForItem -Mappings $urlMappingsForMatch

                                if ((-not $matched) -and (Test-IsNonActionableUrl -Url $url)) {
                                    continue
                                }

                                $externalBroken = $false
                                if (-not $matched -and $sourceOnSelectedBase) {
                                    $externalBroken = Test-IsExternalBrokenUrl -Url $url -TenantHosts $tenantHosts -Cache $externalUrlStatusCache -HttpClient $httpClient
                                }

                                if ((-not $matched) -and (-not $externalBroken)) {
                                    continue
                                }

                                $newUrlValue = if ($matched) { $matched.NewUrl } else { '' }
                                $displayText = $siteContentsFolder
                                $linkType = Get-LinkTypeLabel -Url $url -TenantHosts $tenantHosts

                                $targetPlanKey = '__ALL__'
                                if ($selectedAllSubsites) {
                                    $targetPlanKey = Get-SiteKeyForSourceUrl -SourceUrl $sourceUrlForItem -Mappings $urlMappingsForMatch
                                    if ([string]::IsNullOrWhiteSpace($targetPlanKey) -or -not $siteOutputPlans.ContainsKey($targetPlanKey)) {
                                        continue
                                    }
                                }

                                $row = [PSCustomObject][ordered]@{
                                    'Link Type' = $linkType
                                    'Site Title' = $siteTitle
                                    'Content Type' = $contentType
                                    'Item ID' = $itemId
                                    'Item Title' = $itemTitle
                                    'Display Text' = $displayText
                                    'Item URL' = $url
                                    'Site Contents Folder' = $siteContentsFolder
                                    'Site Contents Folder URL' = $siteContentsFolderUrl
                                    'New Item URL' = $newUrlValue
                                    'Created By' = $createdBy
                                    'Last Modified By' = $lastModifiedBy
                                    'Last Modified Date' = $modifiedText
                                    'Permissions' = $otherTopLevelPermissions
                                }

                                if (-not (Test-IsExactRowDuplicate -Row $row -SeenRows $otherSeenRows)) {
                                    Write-OtherFinding -OutputPath $siteOutputPlans[$targetPlanKey].OtherReportPath -Row $row
                                    $siteOutputStats[$targetPlanKey].OtherRows++
                                    $stats.OtherLinkRows++
                                }
                            }
                        }
                    }
                }
            }
        }
        if ($selectedAllSubsites) {
            foreach ($plan in $siteOutputPlans.Values) {
                Save-CounterCheckpoint -Path $plan.CounterCheckpointPath -DocScanned $docScannedByExt -DocSkippedForced $docSkippedForcedByExt -DocSkippedUnsupported $docSkippedUnsupportedByExt -OtherScanned $otherScannedByExt -OtherSkippedForced $otherSkippedForcedByExt -OtherSkippedUnsupported $otherSkippedUnsupportedByExt
            }
        }
        else {
            Save-CounterCheckpoint -Path $counterCheckpointPath -DocScanned $docScannedByExt -DocSkippedForced $docSkippedForcedByExt -DocSkippedUnsupported $docSkippedUnsupportedByExt -OtherScanned $otherScannedByExt -OtherSkippedForced $otherSkippedForcedByExt -OtherSkippedUnsupported $otherSkippedUnsupportedByExt
        }
    }
}

$overviewReportPathsBySite = @{}

$extensionDetailsRows = New-Object 'System.Collections.Generic.List[object]'
Add-ExtensionDetailRows -Rows $extensionDetailsRows -Category 'Documents Scanned' -Bucket $docScannedByExt
Add-ExtensionDetailRows -Rows $extensionDetailsRows -Category 'Documents Skipped Forced Set' -Bucket $docSkippedForcedByExt
Add-ExtensionDetailRows -Rows $extensionDetailsRows -Category 'Documents Skipped Unsupported' -Bucket $docSkippedUnsupportedByExt

if ($scanOther) {
    Add-ExtensionDetailRows -Rows $extensionDetailsRows -Category 'Other Files Scanned' -Bucket $otherScannedByExt
    Add-ExtensionDetailRows -Rows $extensionDetailsRows -Category 'Other Files Skipped Forced Set' -Bucket $otherSkippedForcedByExt
    Add-ExtensionDetailRows -Rows $extensionDetailsRows -Category 'Other Files Skipped Unsupported' -Bucket $otherSkippedUnsupportedByExt
}

foreach ($planKey in $siteOutputPlans.Keys) {
    $plan = $siteOutputPlans[$planKey]
    $planStats = $siteOutputStats[$planKey]

    if ($planStats.DocumentRows -eq 0) {
        Get-NoFindingsDocumentRow | Export-Csv -Path $plan.DocumentReportPath -NoTypeInformation
    }

    if ($scanOther -and $planStats.OtherRows -eq 0) {
        Get-NoFindingsOtherRow | Export-Csv -Path $plan.OtherReportPath -NoTypeInformation
    }

    $overviewReportPathsBySite[$planKey] = Export-OverviewOnlyReports -DocumentReportPath $plan.DocumentReportPath -OtherReportPath $plan.OtherReportPath -IncludeOther:$scanOther

    $extensionDetailsRows | Export-Csv -Path $plan.ExtensionDetailsPath -NoTypeInformation
    Save-CounterCheckpoint -Path $plan.CounterCheckpointPath -DocScanned $docScannedByExt -DocSkippedForced $docSkippedForcedByExt -DocSkippedUnsupported $docSkippedUnsupportedByExt -OtherScanned $otherScannedByExt -OtherSkippedForced $otherSkippedForcedByExt -OtherSkippedUnsupported $otherSkippedUnsupportedByExt
}

$overviewReportPaths = $overviewReportPathsBySite['__ALL__']

if ($UsePersistedExternalUrlCache -and -not [string]::IsNullOrWhiteSpace($externalUrlStatusCachePath)) {
    Save-ExternalUrlStatusCache -Path $externalUrlStatusCachePath -Cache $externalUrlStatusCache -SchemaVersion $ScriptSchemaVersion
}

Write-Host "`nAudit complete."
Write-Host "  Site collections visited:       $($stats.SitesVisited)"
Write-Host "  Webs visited:                   $($stats.WebsVisited)"
Write-Host "  Libraries visited:              $($stats.LibrariesVisited)"
Write-Host "  Documents inspected:            $($stats.DocumentsRead)"
Write-Host "  Document rows written:          $($stats.DocumentLinkRows)"
Write-Host "  Documents skipped (always skip set): $($stats.DocumentsSkippedForced)"
Write-Host "  Unsupported docs skipped:       $($stats.DocumentsSkippedUnsupported)"
if ($scanOther) {
    Write-Host "  Other-content rows written:     $($stats.OtherLinkRows)"
    Write-Host "  Other files scanned (.htm/.aspx): $($stats.OtherFilesScanned)"
    Write-Host "  Other files skipped (always skip set): $($stats.OtherFilesSkippedForced)"
    Write-Host "  Other files skipped unsupported: $($stats.OtherFilesSkippedUnsupported)"
}
if ($selectedAllSubsites) {
    Write-Host "  Extension details CSV:         Per-site output files"
}
else {
    Write-Host "  Extension details CSV:         $extensionDetailsPath"
}
if ($UsePersistedExternalUrlCache -and -not [string]::IsNullOrWhiteSpace($externalUrlStatusCachePath)) {
    Write-Host "  External URL cache:           $externalUrlStatusCachePath"
}
if ($selectedAllSubsites) {
    Write-Host "`nPer-site output files:"
    foreach ($plan in @($siteOutputPlans.Values | Sort-Object SiteSlug)) {
        $planOverview = $overviewReportPathsBySite[$plan.SiteKey]
        Write-Host ("  Site: {0}" -f $plan.SiteKey)
        Write-Host ("    Documents CSV: {0}" -f $plan.DocumentReportPath)
        Write-Host ("    Documents Overview CSV: {0}" -f $planOverview.DocumentOverviewPath)
        if ($scanOther) {
            Write-Host ("    Other CSV: {0}" -f $plan.OtherReportPath)
            Write-Host ("    Other Overview CSV: {0}" -f $planOverview.OtherOverviewPath)
        }
        Write-Host ("    Extension details CSV: {0}" -f $plan.ExtensionDetailsPath)
        Write-Host ("    Counter checkpoint JSON: {0}" -f $plan.CounterCheckpointPath)
    }
}
else {
    Write-Host "`nDocuments CSV: $docReportPath"
    Write-Host "Documents Overview CSV: $($overviewReportPaths.DocumentOverviewPath)"
    if ($scanOther) {
        Write-Host "Other CSV:     $otherReportPath"
        Write-Host "Other Overview CSV: $($overviewReportPaths.OtherOverviewPath)"
    }
}

