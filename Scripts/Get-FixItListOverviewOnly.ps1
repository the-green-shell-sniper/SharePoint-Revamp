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
            - Documents folder: one row per affected document with broken URL occurrence counts.
            - Other folder: one row per affected non-document item with broken URL occurrence counts.
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
    [string]$OtherOutputFolder = "C:\Users\!bposey\OneDrive - Womens Healthcare Associates LLC\Gary Henderson's files - Project Files_bposey\SharePoint-Migration-Scripts\FixItList-OutputFiles\NonDocuments"
)
. (Join-Path $PSScriptRoot 'SPOutputLimits.Common.ps1')

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
                ($entry.FullName -match 'document\.xml$|sheet\d*\.xml$|slide\d*\.xml$|header\d*\.xml$|footer\d*\.xml$|notesSlide\d*\.xml$')) {

                $s = $entry.Open()
                $r = New-Object System.IO.StreamReader($s)
                $text = $r.ReadToEnd()
                $r.Close()
                $s.Close()

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

                                $found.Add([PSCustomObject]@{ Value = $relationsByXml[$entry.FullName][$relId]; Text = $displayText; Confidence = "High"; Kind = "hyperlink" })
                            }
                        }
                    }
                    catch {
                        Write-Verbose "Could not parse hyperlink display text in $($entry.FullName)."
                    }
                }

                $plain = [regex]::Replace($text, '<[^>]+>', ' ')
                foreach ($m in [regex]::Matches($plain, 'https?://[^\s"<>]+')) {
                    $found.Add([PSCustomObject]@{ Value = $m.Value; Text = $m.Value; Confidence = "Medium"; Kind = "plaintext" })
                }
            }
        }
    }
    finally {
        $archive.Dispose()
    }

    return $found
}

function Get-PdfReference {
    param([Parameter(Mandatory = $true)][byte[]]$Bytes)

    $found = New-Object System.Collections.Generic.List[Object]
    $text = [System.Text.Encoding]::Latin1.GetString($Bytes)

    foreach ($m in [regex]::Matches($text, 'https?://[^\s\)"<>\\]+')) {
        $hostMatch = [regex]::Match($m.Value, '^https?://([^/:?#]+)')
        if ($hostMatch.Success -and ($PdfNoiseHosts -contains $hostMatch.Groups[1].Value.ToLower())) {
            continue
        }
        $found.Add([PSCustomObject]@{ Value = $m.Value; Text = $m.Value; Confidence = "Low"; Kind = "plaintext-pdf" })
    }

    return $found
}

function Get-UrlMappingMatch {
    param(
        [Parameter(Mandatory = $true)][string]$Value,
        [Parameter(Mandatory = $true)][array]$Mappings
    )

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return $null
    }

    $clean = Get-CleanedValue -Value $Value

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

function Get-CentralParkCorrection {
    param([Parameter(Mandatory = $true)][string]$Url)

    $clean = Get-CleanedValue -Value $Url
    if ([string]::IsNullOrWhiteSpace($clean)) {
        return $null
    }

    $prefixPattern = '^(?i)https?://centralpark(?=$|[/:?#])'
    if ($clean -notmatch $prefixPattern) {
        return $null
    }

    $newUrl = [regex]::Replace($clean, $prefixPattern, 'https://whallc.sharepoint.com')
    return [PSCustomObject]@{
        OldUrl = $clean
        NewUrl = $newUrl
    }
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
        $assignments = Invoke-WithRetry -Label "Get role assignments" -MaxAttempts 4 -Operation {
            Get-PnPProperty -ClientObject $Securable -Property RoleAssignments -Connection $Connection
        }
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
            $member = Invoke-WithRetry -Label "Get role assignment member" -MaxAttempts 4 -Operation {
                Get-PnPProperty -ClientObject $ra -Property Member -Connection $Connection
            }
            $bindings = Invoke-WithRetry -Label "Get role assignment bindings" -MaxAttempts 4 -Operation {
                Get-PnPProperty -ClientObject $ra -Property RoleDefinitionBindings -Connection $Connection
            }
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
        [Parameter(Mandatory = $true)][hashtable]$Cache
    )

    $itemUnique = $false
    $listUnique = $false

    try {
        $itemUnique = [bool](Get-PnPProperty -ClientObject $Item -Property HasUniqueRoleAssignments -Connection $Connection)
    }
    catch {
        Write-Verbose "Could not determine item-level unique permissions for item $($Item.Id)."
    }

    try {
        $listUnique = [bool](Get-PnPProperty -ClientObject $List -Property HasUniqueRoleAssignments -Connection $Connection)
    }
    catch {
        Write-Verbose "Could not determine list-level unique permissions for list $($List.Title)."
    }

    if ($itemUnique) {
        return Get-RoleAssignmentTopLevel -Securable $Item -Connection $Connection -CacheKey ("Item:" + $List.Id + ":" + $Item.Id) -Cache $Cache
    }

    if ($listUnique) {
        return Get-RoleAssignmentTopLevel -Securable $List -Connection $Connection -CacheKey ("List:" + $Web.Id + ":" + $List.Id) -Cache $Cache
    }

    return Get-RoleAssignmentTopLevel -Securable $Web -Connection $Connection -CacheKey ("Web:" + $Web.Id) -Cache $Cache
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

    if ($Cache.ContainsKey($clean)) {
        return [bool]$Cache[$clean]
    }

    $uri = $null
    if (-not [Uri]::TryCreate($clean, [UriKind]::Absolute, [ref]$uri)) {
        $Cache[$clean] = $false
        return $false
    }

    if ($uri.Scheme -notin @('http', 'https')) {
        $Cache[$clean] = $false
        return $false
    }

    $urlHost = $uri.Host.ToLower()
    $isSharePointHost = $urlHost -eq 'sharepoint.com' -or $urlHost.EndsWith('.sharepoint.com')
    if ($isSharePointHost -or ($TenantHosts -contains $urlHost)) {
        $Cache[$clean] = $false
        return $false
    }

    $broken = $false
    try {
        $headReq = New-Object System.Net.Http.HttpRequestMessage([System.Net.Http.HttpMethod]::Head, $uri)
        $headResp = $HttpClient.Send($headReq)
        if ([int]$headResp.StatusCode -ge 400) {
            if ([int]$headResp.StatusCode -in 405, 501) {
                $getReq = New-Object System.Net.Http.HttpRequestMessage([System.Net.Http.HttpMethod]::Get, $uri)
                $getResp = $HttpClient.Send($getReq, [System.Net.Http.HttpCompletionOption]::ResponseHeadersRead)
                $broken = [int]$getResp.StatusCode -ge 400
            }
            else {
                $broken = $true
            }
        }
    }
    catch {
        $broken = $true
    }

    $Cache[$clean] = $broken
    return $broken
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

    return $hits
}

function Get-NoFindingsDocumentRow {
    return [PSCustomObject][ordered]@{
        'Link Type' = ''
        'Site Title' = ''
        'Site URL' = ''
        'Library' = ''
        'Document Title' = ''
        'Document URL' = 'No findings to report'
        'Document Folder' = ''
        'Document Folder URL' = ''
        'Document Path' = ''
        'File Extension' = ''
        'Broken URL Occurrences' = ''
        'Matched Old Base URL Count' = ''
        'Matched Old Base URLs' = ''
        'Matched New Base URLs' = ''
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
        'Site URL' = ''
        'Content Type' = ''
        'Item ID' = ''
        'Item Title' = ''
        'Item URL' = 'No findings to report'
        'Site Contents Folder' = ''
        'Site Contents Folder URL' = ''
        'Created By' = ''
        'Last Modified By' = ''
        'Last Modified Date' = ''
        'Permissions' = ''
        'Broken URL Occurrences' = ''
        'Field Count With Broken URLs' = ''
        'Fields With Broken URLs' = ''
        'Matched Old Base URL Count' = ''
        'Matched Old Base URLs' = ''
        'Matched New Base URLs' = ''
    }
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
$selectedScopeUrl = $selectedScope.Url.TrimEnd('/')
Write-Host "`nSelected inventory scope: $selectedScopeUrl"

$scopeConn = Connect-ValidatedPnP -Url $selectedScopeUrl -ClientId $ClientId

Write-Host ("`nSubsite inventory under selected scope {0}:" -f $selectedScopeUrl)
$scopeSubsites = $null
try {
    $scopeRootWeb = Get-PnPWeb -Connection $scopeConn -Includes Url, Title
    $scopeSubsites = New-Object System.Collections.Generic.List[Object]
    $scopeSubsites.Add([PSCustomObject]@{ Url = $scopeRootWeb.Url; Title = $scopeRootWeb.Title })

    $children = @(Get-PnPSubWeb -Connection $scopeConn -Recurse -Includes Url, Title | Sort-Object Url)
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

for ($i = 0; $i -lt $scopeSubsites.Count; $i++) {
    $title = if ($scopeSubsites[$i].Title) { $scopeSubsites[$i].Title } else { '(no title)' }
    Write-Host ("  {0,4}. {1}  |  {2}" -f ($i + 1), $scopeSubsites[$i].Url, $title)
}

$selectedIndices = $null
while (-not $selectedIndices) {
    $pickRaw = Read-Host "`nPick one or more subsites by number (comma-separated, e.g. 1,4,7)"
    if ([string]::IsNullOrWhiteSpace($pickRaw)) {
        Write-Host "  Selection is required."
        continue
    }

    $parts = @($pickRaw -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    $valid = $true
    $numbers = New-Object System.Collections.Generic.List[int]

    foreach ($p in $parts) {
        if ($p -notmatch '^\d+$') {
            $valid = $false
            break
        }
        $n = [int]$p
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
foreach ($sub in $selectedSubsites) {
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

$selectedSiteSlug = (@($selectedSiteSlugParts | Select-Object -Unique) -join '-')
if ([string]::IsNullOrWhiteSpace($selectedSiteSlug)) {
    $selectedSiteSlug = 'SelectedSites'
}

$scriptTitleSegment = 'Get-FixItListOverviewOnly'
$reportPrefixBase = "$selectedSiteSlug-$scriptTitleSegment"
$runStamp = Get-Date -Format "yyyyMMdd-HHmmss"
$docReportPath = Join-Path -Path $DocumentsOutputFolder -ChildPath ("$reportPrefixBase-Documents-" + $runStamp + ".csv")
$otherReportPath = Join-Path -Path $OtherOutputFolder -ChildPath ("$reportPrefixBase-OtherContent-" + $runStamp + ".csv")

Write-Host "`nAudit configuration:"
Write-Host "  Selected subsite mappings:"
foreach ($m in $urlMappings) {
    Write-Host ("    OLD: {0}" -f $m.OldBaseUrl)
    Write-Host ("    NEW: {0}" -f $m.NewBaseUrl)
}
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
    DocumentBrokenUrlOccurrences = 0
    OtherLinkRows = 0
    OtherBrokenUrlOccurrences = 0
    DocumentsSkippedLegacy = 0
    DocumentsSkippedUnsupported = 0
}

$permissionCache = @{}
$folderNameCache = @{}
$tenantHosts = @($tenantSites | ForEach-Object { try { ([Uri]$_.Url).Host.ToLower() } catch { $null } } | Where-Object { $_ } | Sort-Object -Unique)
$httpClient = [System.Net.Http.HttpClient]::new()
$httpClient.Timeout = [TimeSpan]::FromSeconds(8)

$scanSiteCollections = @($tenantSites | Sort-Object Url)

Write-Host "`nStarting tenant-wide document scan (SharePoint only, no OneDrive)..."
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
        $siteTenantBase = ([Uri]$webUrl).GetLeftPart([UriPartial]::Authority)
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
            $web = Get-PnPWeb -Connection $webConn -Includes Id, Url, Title
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

        $libraryCounterInWeb = 0
        foreach ($list in $lists) {
            if ($list.Hidden) { continue }
            if ($SkipListTitles -contains $list.Title) { continue }
            if ($list.ItemCount -eq 0) { continue }
            if ($list.BaseTemplate -ne 101) { continue }

            $libraryCounterInWeb++
            if ($libraryCounterInWeb % 5 -eq 0) {
                try {
                    $webConn = Update-PnPConnection -Url $webUrl -ClientId $ClientId
                    $web = Get-PnPWeb -Connection $webConn -Includes Id, Url, Title
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
                    $retryState.Web = Get-PnPWeb -Connection $retryState.WebConn -Includes Id, Url, Title
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
                        $retryState.WebConn = Reset-PnPAuthenticationConnection -Url $webUrl -ClientId $ClientId
                        $retryState.Web = Get-PnPWeb -Connection $retryState.WebConn -Includes Id, Url, Title
                        $items = Get-PnPListItem -List $list.Title -Connection $retryState.WebConn -PageSize 500 -Fields $fields
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

                $ext = [System.IO.Path]::GetExtension($fileLeaf).ToLower()
                if ($LegacyExtensions -contains $ext) {
                    $stats.DocumentsSkippedLegacy++
                    continue
                }
                if (-not ($SupportedExtensions -contains $ext)) {
                    $stats.DocumentsSkippedUnsupported++
                    continue
                }

                $docUrl = Get-DocumentBrowserUrl -TenantBase $siteTenantBase -FileRef $fileRef -Extension $ext -UniqueId $item.FieldValues['UniqueId']
                $folderUrl = Get-FolderBrowserUrl -TenantBase $siteTenantBase -LibraryRootPath $libraryRootPath -ListId $list.Id -FileDirRef $item.FieldValues['FileDirRef']
                $folderName = Get-FolderDisplayName -List $list -FolderPath $item.FieldValues['FileDirRef'] -Connection $webConn -Cache $folderNameCache

                $authorValue = $item.FieldValues['Author']
                $createdBy = if ($authorValue -and $authorValue.LookupValue) { $authorValue.LookupValue } else { '' }

                $editorValue = $item.FieldValues['Editor']
                $lastModifiedBy = if ($editorValue -and $editorValue.LookupValue) { $editorValue.LookupValue } else { '' }

                $topLevelPermissions = ''
                try {
                    $topLevelPermissions = Get-EffectiveTopLevelPermission -Item $item -List $list -Web $web -Connection $webConn -Cache $permissionCache
                }
                catch {
                    $topLevelPermissions = "(could not read permissions: $($_.Exception.Message))"
                }

                $refs = @()
                try {
                    $fileRetryState = @{ WebConn = $webConn; Web = $web }
                    $ms = Invoke-WithRetry -Label ("Get file stream: " + $fileRef) -MaxAttempts 4 -OnTimeout {
                        $fileRetryState.WebConn = Update-PnPConnection -Url $webUrl -ClientId $ClientId
                        $fileRetryState.Web = Get-PnPWeb -Connection $fileRetryState.WebConn -Includes Id, Url, Title
                    } -Operation {
                        Get-PnPFile -Url $fileRef -AsMemoryStream -Connection $fileRetryState.WebConn
                    }

                    $webConn = $fileRetryState.WebConn
                    $web = $fileRetryState.Web

                    if ($ext -eq '.pdf') {
                        $refs = Get-PdfReference -Bytes $ms.ToArray()
                    }
                    else {
                        $refs = Get-OfficeDocReference -Stream $ms
                    }
                }
                catch {
                    $message = [string]$_.Exception.Message
                    if (Test-IsAuthFailureMessage -Message $message) {
                        Write-Warning "    Authentication issue while reading file content for $fileRef; prompting for a fresh sign-in."
                        try {
                            $fileRetryState.WebConn = Reset-PnPAuthenticationConnection -Url $webUrl -ClientId $ClientId
                            $fileRetryState.Web = Get-PnPWeb -Connection $fileRetryState.WebConn -Includes Id, Url, Title
                            $ms = Get-PnPFile -Url $fileRef -AsMemoryStream -Connection $fileRetryState.WebConn
                            $webConn = $fileRetryState.WebConn
                            $web = $fileRetryState.Web

                            if ($ext -eq '.pdf') {
                                $refs = Get-PdfReference -Bytes $ms.ToArray()
                            }
                            else {
                                $refs = Get-OfficeDocReference -Stream $ms
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

                $documentBrokenCount = 0
                $matchedOldBases = New-Object System.Collections.Generic.List[string]
                $matchedNewBases = New-Object System.Collections.Generic.List[string]
                $matchedLinkTypes = New-Object System.Collections.Generic.List[string]
                foreach ($ref in $refs) {
                    $matched = Get-UrlMappingMatch -Value $ref.Value -Mappings $urlMappingsForMatch
                    if (-not $matched) {
                        continue
                    }

                    $documentBrokenCount++
                    $matchedOldBases.Add([string]$matched.OldBaseUrl)
                    $matchedNewBases.Add([string]$matched.NewBaseUrl)
                    $matchedLinkTypes.Add((Get-LinkTypeLabel -Url $ref.Value -TenantHosts $tenantHosts))
                }

                if ($documentBrokenCount -le 0) {
                    continue
                }

                $uniqueOldBases = @($matchedOldBases | Sort-Object -Unique)
                $uniqueNewBases = @($matchedNewBases | Sort-Object -Unique)
                $uniqueLinkTypes = @($matchedLinkTypes | Sort-Object -Unique)
                $linkType = if ($uniqueLinkTypes.Count -eq 1) { [string]$uniqueLinkTypes[0] } else { 'Both' }

                $siteTitle = if ([string]::IsNullOrWhiteSpace($web.Title)) { $webUrl } else { $web.Title }
                $modifiedValue = $item.FieldValues['Modified']
                $modifiedText = Format-ReportDatePacific -Value $modifiedValue

                $row = [PSCustomObject][ordered]@{
                    'Link Type' = $linkType
                    'Site Title' = $siteTitle
                    'Site URL' = $webUrl
                    'Library' = $list.Title
                    'Document Title' = $fileLeaf
                    'Document URL' = $docUrl
                    'Document Folder' = $folderName
                    'Document Folder URL' = $folderUrl
                    'Document Path' = $fileRef
                    'File Extension' = $ext
                    'Broken URL Occurrences' = $documentBrokenCount
                    'Matched Old Base URL Count' = $uniqueOldBases.Count
                    'Matched Old Base URLs' = ($uniqueOldBases -join '; ')
                    'Matched New Base URLs' = ($uniqueNewBases -join '; ')
                    'Created By' = $createdBy
                    'Last Modified By' = $lastModifiedBy
                    'Last Modified Date' = $modifiedText
                    'Permissions' = $topLevelPermissions
                }

                Write-DocumentFinding -OutputPath $docReportPath -Row $row
                $stats.DocumentLinkRows++
                $stats.DocumentBrokenUrlOccurrences += $documentBrokenCount
            }
        }
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
            $rootWeb = Get-PnPWeb -Connection $siteConn -Includes Url, Title
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

            $web = $null
            try {
                $web = Get-PnPWeb -Connection $webConn -Includes Id, Url, Title
            }
            catch {
                continue
            }

            $lists = $null
            try {
                $lists = Get-PnPList -Connection $webConn -Includes BaseTemplate, Hidden, Title, ItemCount, Id, HasUniqueRoleAssignments, DefaultDisplayFormUrl, RootFolder.ServerRelativeUrl
            }
            catch {
                continue
            }

            $otherLibraryCounterInWeb = 0
            foreach ($list in $lists) {
                if ($list.Hidden) { continue }
                if ($list.ItemCount -eq 0) { continue }
                if ($list.BaseTemplate -eq 101) { continue }

                $otherLibraryCounterInWeb++
                if ($otherLibraryCounterInWeb % 5 -eq 0) {
                    try {
                        $webConn = Update-PnPConnection -Url $webUrl -ClientId $ClientId
                        $web = Get-PnPWeb -Connection $webConn -Includes Id, Url, Title
                    }
                    catch {
                        Write-Verbose "Could not refresh non-document scan web connection for $webUrl; continuing with current connection."
                    }
                }

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
                    $otherRetryState = @{ WebConn = $webConn; Web = $web }
                    $items = Invoke-WithRetry -Label ("Get list items: " + $list.Title) -MaxAttempts 4 -OnTimeout {
                        $otherRetryState.WebConn = Update-PnPConnection -Url $webUrl -ClientId $ClientId
                        $otherRetryState.Web = Get-PnPWeb -Connection $otherRetryState.WebConn -Includes Id, Url, Title
                    } -Operation {
                        Get-PnPListItem -List $list.Title -Connection $otherRetryState.WebConn -PageSize 200
                    }

                    $webConn = $otherRetryState.WebConn
                    $web = $otherRetryState.Web
                }
                catch {
                    $message = [string]$_.Exception.Message
                    if (Test-IsAuthFailureMessage -Message $message) {
                        Write-Warning "    Authentication issue while reading other-content list items for $($list.Title); prompting for a fresh sign-in."
                        try {
                            $otherRetryState.WebConn = Reset-PnPAuthenticationConnection -Url $webUrl -ClientId $ClientId
                            $otherRetryState.Web = Get-PnPWeb -Connection $otherRetryState.WebConn -Includes Id, Url, Title
                            $items = Get-PnPListItem -List $list.Title -Connection $otherRetryState.WebConn -PageSize 200
                            $webConn = $otherRetryState.WebConn
                            $web = $otherRetryState.Web
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
                    $itemBrokenCount = 0
                    $matchedOldBases = New-Object System.Collections.Generic.List[string]
                    $matchedNewBases = New-Object System.Collections.Generic.List[string]
                    $matchedFields = New-Object System.Collections.Generic.List[string]
                    $matchedLinkTypes = New-Object System.Collections.Generic.List[string]

                    $location = $list.Title
                    if ($item.FieldValues.ContainsKey('FileDirRef') -and $item.FieldValues['FileDirRef']) {
                        try {
                            $location = Get-FolderDisplayName -List $list -FolderPath $item.FieldValues['FileDirRef'] -Connection $webConn -Cache $folderNameCache
                        }
                        catch {
                            $location = $list.Title
                        }
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
                    $siteContentsFolderPath = if ($item.FieldValues.ContainsKey('FileDirRef') -and $item.FieldValues['FileDirRef']) { [string]$item.FieldValues['FileDirRef'] } else { [string]$list.RootFolder.ServerRelativeUrl }
                    $siteContentsFolderUrl = ''
                    try {
                        $siteContentsFolderUrl = Get-FolderBrowserUrl -TenantBase $siteTenantBase -LibraryRootPath $list.RootFolder.ServerRelativeUrl.TrimEnd('/') -ListId $list.Id -FileDirRef $siteContentsFolderPath
                    }
                    catch {
                        $siteContentsFolderUrl = ''
                    }

                    $otherTopLevelPermissions = ''
                    try {
                        $otherTopLevelPermissions = Get-EffectiveTopLevelPermission -Item $item -List $list -Web $web -Connection $webConn -Cache $permissionCache
                    }
                    catch {
                        $otherTopLevelPermissions = "(could not read permissions: $($_.Exception.Message))"
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
                                if (-not $matched) {
                                    continue
                                }

                                $itemBrokenCount++
                                $matchedOldBases.Add([string]$matched.OldBaseUrl)
                                $matchedNewBases.Add([string]$matched.NewBaseUrl)
                                $matchedFields.Add([string]$field.InternalName)
                                $matchedLinkTypes.Add((Get-LinkTypeLabel -Url $url -TenantHosts $tenantHosts))
                            }
                        }
                    }

                    if ($itemBrokenCount -gt 0) {
                        $uniqueOldBases = @($matchedOldBases | Sort-Object -Unique)
                        $uniqueNewBases = @($matchedNewBases | Sort-Object -Unique)
                        $uniqueFields = @($matchedFields | Sort-Object -Unique)
                        $uniqueLinkTypes = @($matchedLinkTypes | Sort-Object -Unique)
                        $linkType = if ($uniqueLinkTypes.Count -eq 1) { [string]$uniqueLinkTypes[0] } else { 'Both' }
                        $itemUrl = ''
                        if ($list.DefaultDisplayFormUrl) {
                            $itemTenantBase = ([Uri]$webUrl).GetLeftPart([UriPartial]::Authority)
                            $displayFormUrl = [string]$list.DefaultDisplayFormUrl
                            if ($displayFormUrl -match '^https?://') {
                                $itemUrl = "$displayFormUrl?ID=$itemId"
                            }
                            else {
                                $itemUrl = "$itemTenantBase$displayFormUrl?ID=$itemId"
                            }
                        }

                        # one row per item aggregation
                        $row = [PSCustomObject][ordered]@{
                            'Link Type' = $linkType
                            'Site Title' = $siteTitle
                            'Site URL' = $webUrl
                            'Content Type' = $contentType
                            'Item ID' = $itemId
                            'Item Title' = $itemTitle
                            'Item URL' = $itemUrl
                            'Site Contents Folder' = $location
                            'Site Contents Folder URL' = $siteContentsFolderUrl
                            'Created By' = $createdBy
                            'Last Modified By' = $lastModifiedBy
                            'Last Modified Date' = $modifiedText
                            'Permissions' = $otherTopLevelPermissions
                            'Broken URL Occurrences' = $itemBrokenCount
                            'Field Count With Broken URLs' = $uniqueFields.Count
                            'Fields With Broken URLs' = ($uniqueFields -join '; ')
                            'Matched Old Base URL Count' = $uniqueOldBases.Count
                            'Matched Old Base URLs' = ($uniqueOldBases -join '; ')
                            'Matched New Base URLs' = ($uniqueNewBases -join '; ')
                        }

                        Write-OtherFinding -OutputPath $otherReportPath -Row $row
                        $stats.OtherLinkRows++
                        $stats.OtherBrokenUrlOccurrences += $itemBrokenCount
                    }
                }
            }
        }
    }
}

if ($stats.DocumentLinkRows -eq 0) {
    Get-NoFindingsDocumentRow | Export-Csv -Path $docReportPath -NoTypeInformation
}

if ($scanOther -and $stats.OtherLinkRows -eq 0) {
    Get-NoFindingsOtherRow | Export-Csv -Path $otherReportPath -NoTypeInformation
}

Write-Host "`nAudit complete."
Write-Host "  Site collections visited:       $($stats.SitesVisited)"
Write-Host "  Webs visited:                   $($stats.WebsVisited)"
Write-Host "  Libraries visited:              $($stats.LibrariesVisited)"
Write-Host "  Documents inspected:            $($stats.DocumentsRead)"
Write-Host "  Document rows written:          $($stats.DocumentLinkRows)"
Write-Host "  Document broken URL occurrences:$($stats.DocumentBrokenUrlOccurrences)"
Write-Host "  Legacy docs skipped:            $($stats.DocumentsSkippedLegacy)"
Write-Host "  Unsupported docs skipped:       $($stats.DocumentsSkippedUnsupported)"
if ($scanOther) {
    Write-Host "  Other-content rows written:     $($stats.OtherLinkRows)"
    Write-Host "  Other broken URL occurrences:   $($stats.OtherBrokenUrlOccurrences)"
}
Write-Host "`nDocuments CSV: $docReportPath"
if ($scanOther) {
    Write-Host "Other CSV:     $otherReportPath"
}

