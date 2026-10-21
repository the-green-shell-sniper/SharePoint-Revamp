#requires -Version 7.4.6
<#
.SYNOPSIS
    Preview old !bposey edits across new checklist sites and all their subsites.
.DESCRIPTION
    READ ONLY: no Apply parameter or SharePoint write commands.
    Roots come exclusively from New Site URL in the checklist. Recurses actual
    subsites, not other site collections, and deduplicates overlapping roots.
    Scans all lists, including hidden lists, with RecursiveAll paged queries.
    Matches current Editor by exact login/email, never by display name, and
    current Modified strictly before April 1, 2024 midnight Pacific (07:00 UTC).
    Does not download contents or inspect/rewrite historical versions.
    Lists without Modified/Editor cannot qualify and are reported explicitly.
    Web/list containers and library root folders without item-level metadata
    are not editable items; their children are scanned.
    Count differences are informational notes in Coverage.csv, not scan failures.
    Access failures and unavailable native accounts are reported and cause a terminating
    error after the remaining scope is attempted; a failed run is incomplete.
    Reports contain private paths and user metadata. Do not commit them.
    Console output is one summary per checklist root plus updating progress.
    Full errors remain in Errors.csv; use -Verbose for live error details.
.EXAMPLE
    .\Get-SystemEditorRemediationPlan.ps1
#>
[CmdletBinding()]
param(
    [string]$ChecklistPath = (Join-Path $PSScriptRoot '..\ReferenceFiles\NewSiteChecklist.csv'),
    [string]$OutputFolder = (Join-Path $PSScriptRoot '..\Reports\SystemEditorPlans'),
    [string]$ClientId = '2e20d82f-56ba-418e-b465-9718493f0df8'
)
. (Join-Path $PSScriptRoot 'SPOutputLimits.Common.ps1')

function ConvertTo-SystemEditorWebUrl {
    param([string]$Url)
    $uri = $null
    if (-not [uri]::TryCreate($Url.Trim(), [UriKind]::Absolute, [ref]$uri) -or
        $uri.Scheme -ne 'https' -or $uri.Host -ine 'whallc.sharepoint.com' -or
        -not $uri.IsDefaultPort -or $uri.UserInfo -or $uri.Query -or $uri.Fragment) {
        throw "Invalid new-site URL: $Url"
    }
    $path = [uri]::UnescapeDataString($uri.AbsolutePath)
    if ($path -match '(?i)(^|/)archive[^/]*(/|$)') {
        throw "Archive URL is forbidden: $Url"
    }
    return $uri.AbsoluteUri.TrimEnd('/')
}

function Get-SystemEditorChecklistRoots {
    param([string]$Path)
    $rows = @(Import-Csv -LiteralPath $Path -ErrorAction Stop)
    if (-not $rows.Count -or $rows[0].PSObject.Properties.Name -notcontains 'New Site URL') {
        throw 'Checklist must contain rows and the New Site URL column.'
    }
    $roots = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($row in $rows) {
        if ([string]::IsNullOrWhiteSpace($row.'New Site URL')) {
            throw 'Checklist contains a blank New Site URL.'
        }
        [void]$roots.Add((ConvertTo-SystemEditorWebUrl $row.'New Site URL'))
    }
    return @($roots | Sort-Object)
}

function ConvertTo-SystemEditorUtc {
    param([datetime]$Value)
    # SharePoint dates without a Kind still represent UTC, not machine-local time.
    if ($Value.Kind -eq [DateTimeKind]::Unspecified) {
        return [datetime]::SpecifyKind($Value, [DateTimeKind]::Utc)
    }
    return $Value.ToUniversalTime()
}

function Test-SystemEditorSourceUser {
    param($User)
    return ($User.LoginName -ieq '!bposey@whallc.com' -or
        $User.LoginName -ieq 'i:0#.f|membership|!bposey@whallc.com' -or
        $User.Email -ieq '!bposey@whallc.com')
}

function Invoke-SystemEditorRemediationPlan {
    [CmdletBinding()]
    param(
        [string]$ChecklistPath,
        [string]$OutputFolder,
        [string]$ClientId
    )
    $ErrorActionPreference = 'Stop'
    $roots = @(Get-SystemEditorChecklistRoots $ChecklistPath)
    Import-Module PnP.PowerShell -MinimumVersion 3.3.0 -ErrorAction Stop
    $cutoff = [datetime]::new(2024, 4, 1, 7, 0, 0, [DateTimeKind]::Utc)
    $run = Join-Path $OutputFolder ((Get-Date -Format 'yyyyMMdd-HHmmss-fff') + '-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
    [void](Assert-SPOutputPath $run)
    [void][IO.Directory]::CreateDirectory($run)
    $run = (Resolve-Path -LiteralPath $run).Path
    $files = @{
        Plan = Join-Path $run 'Plan.csv'
        Coverage = Join-Path $run 'Coverage.csv'
        Webs = Join-Path $run 'Webs.csv'
        Errors = Join-Path $run 'Errors.csv'
    }
    'WebUrl,ListId,ListTitle,ItemId,UniqueId,Path,Kind,ModifiedUtc,PreviousEditorId,PreviousEditorLogin,PreviousEditorEmail,SystemEditorId,SystemEditorLogin,CreatedUtc,AuthorId,Version,Status' |
        Set-Content -LiteralPath $files.Plan -Encoding utf8
    'WebUrl,ListId,ListTitle,Hidden,ReportedBefore,ReportedAfter,Scanned,Matches,Status,Note' |
        Set-Content -LiteralPath $files.Coverage -Encoding utf8
    'WebUrl,SystemEditorId,SystemEditorLogin,Status' |
        Set-Content -LiteralPath $files.Webs -Encoding utf8
    'WebUrl,ListId,Stage,Error' | Set-Content -LiteralPath $files.Errors -Encoding utf8
    $summary = [ordered]@{
        Mode = 'PreviewOnly'; Status = 'Running'; CutoffUtc = $cutoff.ToString('o')
        SourceEditor = '!bposey@whallc.com'; ChecklistRoots = $roots
        WebsScanned = 0; ListsScanned = 0; ItemsScanned = 0; Matches = 0; Errors = 0
        ReportFolder = $run
    }
    $visited = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    Write-Host "Preview only. No changes will be made."
    Write-Host "Reports: $run"
    $rootIndex = 0
    try {
        foreach ($root in $roots) {
            $rootIndex++
            $errorsBefore = $summary.Errors
            $matchesBefore = $summary.Matches
            $websBefore = $summary.WebsScanned
            $siteName = ([uri]$root).AbsolutePath.Trim('/')
            if (-not $siteName) { $siteName = 'Tenant home' }
            Write-Progress -Id 1 -Activity "Scanning $rootIndex of $($roots.Count): $siteName" -Status 'Finding subsites'
            # Even visited roots get discovery attempted again if earlier discovery failed.
            $stage = 'Connect/discover subsites'
            try {
                $connection = Connect-PnPOnline -Url $root -ClientId $ClientId -Interactive -ReturnConnection -ErrorAction Stop
                $rootWeb = Get-PnPWeb -Connection $connection -Includes Url -ErrorAction Stop
                if ((ConvertTo-SystemEditorWebUrl $rootWeb.Url) -ine $root) {
                    throw "Connected web does not match requested root $root"
                }
                $webUrls = [Collections.Generic.List[string]]::new()
                $webUrls.Add($root)
                try {
                    Get-PnPSubWeb -Recurse -Connection $connection -Includes Url -ErrorAction Stop | ForEach-Object {
                        $url = ConvertTo-SystemEditorWebUrl $_.Url
                        if (-not $url.StartsWith($root + '/', [StringComparison]::OrdinalIgnoreCase)) {
                            throw "Discovered web is outside root $root : $url"
                        }
                        $webUrls.Add($url)
                    }
                } catch {
                    $summary.Errors++
                    [pscustomobject]@{ WebUrl = $root; ListId = ''; Stage = $stage; Error = $_.Exception.Message } |
                        Export-Csv -LiteralPath $files.Errors -Append -NoTypeInformation
                    Write-Verbose "Subsite discovery failed at $root : $($_.Exception.Message)"
                }
                foreach ($url in $webUrls) {
                    if (-not $visited.Add($url)) { continue }
                    $webRecord = [pscustomobject]@{ WebUrl = $url; SystemEditorId = ''; SystemEditorLogin = ''; Status = 'Incomplete' }
                    $stage = 'Read web/users/lists'
                    try {
                        Write-Progress -Id 1 -Activity "Scanning $rootIndex of $($roots.Count): $siteName" -Status ([uri]$url).AbsolutePath
                        $webConnection = if ($url -ieq $root) { $connection } else {
                            Connect-PnPOnline -Url $url -ClientId $ClientId -Interactive -ReturnConnection -ErrorAction Stop
                        }
                        $web = Get-PnPWeb -Connection $webConnection -Includes Url -ErrorAction Stop
                        if ((ConvertTo-SystemEditorWebUrl $web.Url) -ine $url) { throw "Unexpected connected web for $url" }
                        $users = @(Get-PnPUser -Connection $webConnection -ErrorAction Stop)
                        $sourceUsers = @{}
                        foreach ($user in $users) {
                            if (Test-SystemEditorSourceUser $user) { $sourceUsers[[int]$user.Id] = $user }
                        }
                        $systems = @($users | Where-Object { $_.LoginName -ieq 'SHAREPOINT\system' -and $_.Id -gt 0 })
                        if ($systems.Count -gt 1) { throw 'Multiple native System Account identities returned.' }
                        $system = if ($systems.Count -eq 1) { $systems[0] } else { $null }
                        if ($system) {
                            $webRecord.SystemEditorId = $system.Id
                            $webRecord.SystemEditorLogin = $system.LoginName
                        }
                        $lists = @(Get-PnPList -Connection $webConnection -Includes Id,Title,Hidden,ItemCount,BaseType -ErrorAction Stop)
                        foreach ($list in $lists) {
                            Write-Progress -Id 1 -Activity "Scanning $rootIndex of $($roots.Count): $siteName" -Status "$(([uri]$url).AbsolutePath) - $($list.Title)"
                            $record = [pscustomobject]@{
                                WebUrl = $url; ListId = [string]$list.Id; ListTitle = $list.Title; Hidden = $list.Hidden
                                ReportedBefore = $list.ItemCount; ReportedAfter = ''; Scanned = 0; Matches = 0; Status = 'Incomplete'; Note = ''
                            }
                            try {
                                $fields = @(Get-PnPField -List $list.Id -Connection $webConnection -ErrorAction Stop)
                                $names = @($fields.InternalName)
                                if ($names -notcontains 'Modified' -or $names -notcontains 'Editor') {
                                    $record.Status = 'No Modified/Editor fields'
                                    continue
                                }
                                $selected = @('ID', 'Modified', 'Editor', 'FileRef', 'FSObjType', 'UniqueId', 'Created', 'Author', '_UIVersionString' |
                                    Where-Object { $_ -in $names })
                                $viewFields = ($selected | ForEach-Object { "<FieldRef Name='$_'/>" }) -join ''
                                # Scan by indexed ID, not a date/editor filter that can hit the list threshold.
                                $query = "<View Scope='RecursiveAll'><Query><OrderBy><FieldRef Name='ID'/></OrderBy></Query><ViewFields>$viewFields</ViewFields><RowLimit Paged='TRUE'>500</RowLimit></View>"
                                $ids = [Collections.Generic.HashSet[int]]::new()
                                Get-PnPListItem -List $list.Id -Query $query -PageSize 500 -Connection $webConnection -ErrorAction Stop |
                                    ForEach-Object {
                                        $item = $_
                                        if (-not $ids.Add([int]$item.Id)) { throw "Duplicate item ID $($item.Id)" }
                                        $record.Scanned++
                                        $summary.ItemsScanned++
                                        if (-not $item['Modified'] -or -not $item['Editor']) {
                                            throw "Missing Modified/Editor on item $($item.Id)"
                                        }
                                        $modified = ConvertTo-SystemEditorUtc $item['Modified']
                                        if ($modified -ge $cutoff -or -not $sourceUsers.ContainsKey([int]$item['Editor'].LookupId)) { return }
                                        $editor = $sourceUsers[[int]$item['Editor'].LookupId]
                                        $record.Matches++
                                        $summary.Matches++
                                        $kind = if ($item['FSObjType'] -eq 1) { 'Folder' } elseif ($list.BaseType -eq 'DocumentLibrary') { 'File' } else { 'ListItem' }
                                        [pscustomobject]@{
                                            WebUrl = $url; ListId = [string]$list.Id; ListTitle = $list.Title; ItemId = $item.Id
                                            UniqueId = $item['UniqueId']; Path = $item['FileRef']; Kind = $kind
                                            ModifiedUtc = $modified.ToString('o'); PreviousEditorId = $item['Editor'].LookupId
                                            PreviousEditorLogin = $editor.LoginName; PreviousEditorEmail = $editor.Email
                                            SystemEditorId = $webRecord.SystemEditorId; SystemEditorLogin = $webRecord.SystemEditorLogin
                                            CreatedUtc = $(if ($item['Created']) { (ConvertTo-SystemEditorUtc $item['Created']).ToString('o') } else { '' })
                                            AuthorId = $(if ($item['Author']) { $item['Author'].LookupId } else { '' })
                                            Version = $item['_UIVersionString']
                                            Status = $(if ($system) { 'Preview' } else { 'Blocked: native System Account unavailable' })
                                        } | Export-Csv -LiteralPath $files.Plan -Append -NoTypeInformation
                                    }
                                $latest = Get-PnPList -Identity $list.Id -Connection $webConnection -Includes ItemCount -ErrorAction Stop
                                $record.ReportedAfter = $latest.ItemCount
                                if ($record.Scanned -ne $record.ReportedBefore -or $record.Scanned -ne $record.ReportedAfter) {
                                    $record.Note = ("Reported counts: {0} before, {1} after; retrieved: {2}. Count difference alone does not establish missing content or complete coverage." -f
                                        $record.ReportedBefore, $record.ReportedAfter, $record.Scanned)
                                }
                                if ($record.Matches -gt 0 -and -not $system) {
                                    throw 'Qualifying items found, but native SHAREPOINT\system is unavailable. No substitute permitted.'
                                }
                                $record.Status = 'Scanned'
                                $summary.ListsScanned++
                            } catch {
                                $summary.Errors++
                                [pscustomobject]@{ WebUrl = $url; ListId = [string]$list.Id; Stage = 'Scan list'; Error = $_.Exception.Message } |
                                    Export-Csv -LiteralPath $files.Errors -Append -NoTypeInformation
                                Write-Verbose "List '$($list.Title)' at $url : $($_.Exception.Message)"
                            } finally {
                                $record | Export-Csv -LiteralPath $files.Coverage -Append -NoTypeInformation
                            }
                        }
                        $summary.WebsScanned++
                        $webRecord.Status = 'Enumerated (see list coverage)'
                    } catch {
                        $summary.Errors++
                        [pscustomobject]@{ WebUrl = $url; ListId = ''; Stage = $stage; Error = $_.Exception.Message } |
                            Export-Csv -LiteralPath $files.Errors -Append -NoTypeInformation
                        Write-Verbose "Web $url : $($_.Exception.Message)"
                    } finally {
                        $webRecord | Export-Csv -LiteralPath $files.Webs -Append -NoTypeInformation
                    }
                }
            } catch {
                $summary.Errors++
                [pscustomobject]@{ WebUrl = $root; ListId = ''; Stage = $stage; Error = $_.Exception.Message } |
                    Export-Csv -LiteralPath $files.Errors -Append -NoTypeInformation
                Write-Verbose "Root $root : $($_.Exception.Message)"
            } finally {
                Write-Progress -Id 1 -Activity 'Scanning' -Completed
                $siteErrors = $summary.Errors - $errorsBefore
                $siteMatches = $summary.Matches - $matchesBefore
                $siteWebs = $summary.WebsScanned - $websBefore
                if ($siteErrors) {
                    Write-Warning "${siteName}: $siteMatches matches; $siteErrors issues. See Errors.csv."
                } elseif ($siteWebs) {
                    Write-Host "${siteName}: $siteMatches matches."
                }
            }
        }
        $summary.Status = if ($summary.Errors) { 'Incomplete' } else { 'PreviewComplete' }
    } finally {
        Write-Progress -Id 1 -Activity 'Scanning' -Completed
        if ($summary.Status -eq 'Running') { $summary.Status = 'Incomplete' }
        $summary | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $run 'Summary.json') -Encoding utf8
        Write-Host "Preview $($summary.Status): $($summary.Matches) matches; $($summary.Errors) issues. No changes made."
    }
    if ($summary.Errors) { throw "Preview incomplete. Review Errors.csv in the report folder." }
    return [pscustomobject]$summary
}

if ($MyInvocation.InvocationName -ne '.') {
    Invoke-SystemEditorRemediationPlan -ChecklistPath $ChecklistPath -OutputFolder $OutputFolder -ClientId $ClientId | Out-Null
}
