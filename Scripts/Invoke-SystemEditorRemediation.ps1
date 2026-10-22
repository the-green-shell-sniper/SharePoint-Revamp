#requires -Version 7.4.6
<#
.SYNOPSIS
    Revalidate and optionally apply a completed System Editor preview.
.DESCRIPTION
    Read only unless both -Apply and -MaintenanceWindowConfirmed are supplied.
    MaintenanceWindowConfirmed acknowledges that relevant automation is paused
    and concurrent edits are prevented. UpdateOverwriteVersion may trigger flows.
    Only reviewed Plan.csv candidates in live checklist roots/descendant webs
    are eligible. Changed candidates are skipped. No old sites are consulted.
    Resolves native SHAREPOINT\system live, never trusts saved user lookup IDs.
    Saves all returned field values before each write and verifies stored editable
    fields plus core identity, date, author and version fields afterwards.
    Server-managed counters/computed fields are not claimed to remain unchanged.
    No content upload, permissions changes, new version or old-version rewrites.
    Stops on the first read/write/verification error. Never retries or rolls back
    an ambiguous write. Inspect the journal and live state before rerunning.
    Reports use bounded CSV parts and a small Summary.json. Keep them private.
.EXAMPLE
    .\Invoke-SystemEditorRemediation.ps1 -PreviewFolder .\Reports\SystemEditorPlans\<run>
.EXAMPLE
    .\Invoke-SystemEditorRemediation.ps1 -PreviewFolder .\Reports\SystemEditorPlans\<run> -Apply -MaintenanceWindowConfirmed
#>
[CmdletBinding()]
param(
    [string]$PreviewFolder,
    [string]$ChecklistPath = (Join-Path $PSScriptRoot '..\ReferenceFiles\NewSiteChecklist.csv'),
    [string]$OutputFolder = (Join-Path $PSScriptRoot '..\Reports\SystemEditorApply'),
    [string]$ClientId = '2e20d82f-56ba-418e-b465-9718493f0df8',
    [switch]$Apply,
    [switch]$MaintenanceWindowConfirmed
)
. (Join-Path $PSScriptRoot 'Get-SystemEditorRemediationPlan.ps1') `
    -ChecklistPath $ChecklistPath -OutputFolder $OutputFolder -ClientId $ClientId

function Read-SystemEditorApprovedPlan {
    param([string]$Folder, [string]$Checklist)
    $ErrorActionPreference = 'Stop'
    $Folder = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Folder)
    $Checklist = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Checklist)
    $roots = @(Get-SystemEditorChecklistRoots $Checklist)
    $summary = Get-Content -LiteralPath (Join-Path $Folder 'Summary.json') -Raw | ConvertFrom-Json
    $cutoff = [datetimeoffset]'2024-04-01T07:00:00Z'
    if ($summary.Status -cne 'PreviewComplete' -or $summary.Mode -cne 'PreviewOnly' -or
        $summary.Errors -ne 0 -or $summary.SourceEditor -ine '!bposey@whallc.com' -or
        [datetimeoffset]$summary.CutoffUtc -ne $cutoff) {
        throw 'A completed, error-free preview with the expected account and cutoff is required.'
    }
    $savedRoots = @($summary.ChecklistRoots | ForEach-Object { ConvertTo-SystemEditorWebUrl $_ } | Sort-Object -Unique)
    if (@(Compare-Object $roots $savedRoots).Count) { throw 'Checklist roots differ from the reviewed preview.' }
    if (@(Import-Csv -LiteralPath (Join-Path $Folder 'Errors.csv')).Count) { throw 'Preview contains errors.' }
    $webs = @(Import-Csv -LiteralPath (Join-Path $Folder 'Webs.csv'))
    $coverage = @(Import-Csv -LiteralPath (Join-Path $Folder 'Coverage.csv'))
    if ($webs.Count -ne $summary.WebsScanned -or @($webs | Where-Object Status -notlike 'Enumerated*').Count -or
        @($coverage | Where-Object { $_.Status -notin 'Scanned', 'No Modified/Editor fields' }).Count) {
        throw 'Preview web/list coverage is incomplete or inconsistent.'
    }
    $covered = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($row in $coverage | Where-Object Status -eq 'Scanned') {
        [void]$covered.Add("$($row.WebUrl.TrimEnd('/'))|$([guid]$row.ListId)")
    }
    $rows = @(Import-Csv -LiteralPath (Join-Path $Folder 'Plan.csv'))
    if ($rows.Count -ne $summary.Matches -or
        ($coverage | Measure-Object Matches -Sum).Sum -ne $rows.Count) {
        throw 'Plan row count does not match the completed preview.'
    }
    $keys = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($row in $rows) {
        $url = ConvertTo-SystemEditorWebUrl $row.WebUrl
        $listId = [guid]$row.ListId
        if ($url -notin $webs.WebUrl -or -not $covered.Contains("$url|$listId") -or
            $row.Status -cne 'Preview' -or -not (Test-SystemEditorSourceUser ([pscustomobject]@{
                LoginName = $row.PreviousEditorLogin; Email = $row.PreviousEditorEmail
            })) -or [datetimeoffset]$row.ModifiedUtc -ge $cutoff -or
            [guid]$row.UniqueId -eq [guid]::Empty -or $listId -eq [guid]::Empty -or
            [int]$row.ItemId -le 0 -or [int]$row.AuthorId -le 0 -or
            [string]::IsNullOrWhiteSpace($row.Version) -or -not $row.CreatedUtc -or
            -not $keys.Add("$url|$listId|$($row.ItemId)")) {
            throw 'Plan contains an invalid, duplicate or unreviewed candidate.'
        }
        [void][datetimeoffset]::Parse($row.CreatedUtc, [cultureinfo]::InvariantCulture)
    }
    return [pscustomobject]@{ Roots = $roots; Rows = $rows }
}

function Get-SystemEditorFieldSnapshot {
    param($Item)
    $values = @{}
    foreach ($key in $Item.FieldValues.Keys) {
        $values[$key] = [Management.Automation.PSSerializer]::Serialize($Item[$key], 20)
    }
    return $values
}

function Get-SystemEditorCandidateStatus {
    param($Item, $Row, $Editor, [int]$SystemId)
    if ([guid]$Item['UniqueId'] -ne [guid]$Row.UniqueId) { return 'Skipped: identity changed' }
    if ($Item['Editor'].LookupId -eq $SystemId) { return 'Already System Account' }
    if (-not (Test-SystemEditorSourceUser $Editor) -or
        (ConvertTo-SystemEditorUtc $Item['Modified']) -ge [datetime]::new(2024, 4, 1, 7, 0, 0, [DateTimeKind]::Utc)) {
        return 'Skipped: no longer qualifies'
    }
    if ((ConvertTo-SystemEditorUtc $Item['Modified']) -ne ([datetimeoffset]$Row.ModifiedUtc).UtcDateTime -or
        (ConvertTo-SystemEditorUtc $Item['Created']) -ne ([datetimeoffset]$Row.CreatedUtc).UtcDateTime -or
        $Item['Author'].LookupId -ne [int]$Row.AuthorId -or
        $Item['_UIVersionString'] -cne $Row.Version -or $Item['FileRef'] -cne $Row.Path) {
        return 'Skipped: changed since preview'
    }
    if ($Item.FieldValues.ContainsKey('CheckoutUser') -and $Item['CheckoutUser']) { return 'Skipped: checked out' }
    return 'Eligible'
}

function Set-SystemEditorItemValues {
    param($Item, [int]$SystemId, [datetime]$Modified)
    $value = [Microsoft.SharePoint.Client.FieldUserValue]::new()
    $value.LookupId = $SystemId
    $Item['Editor'] = $value
    $Item['Modified'] = $Modified
    $Item.UpdateOverwriteVersion()
    # Execute once: retrying an ambiguous response could overwrite a later edit.
    $Item.Context.ExecuteQuery()
}

function Invoke-SystemEditorApply {
    [CmdletBinding()]
    param(
        [string]$PreviewFolder, [string]$ChecklistPath, [string]$OutputFolder,
        [string]$ClientId, [switch]$Apply, [switch]$MaintenanceWindowConfirmed
    )
    $ErrorActionPreference = 'Stop'
    if ($Apply -and -not $MaintenanceWindowConfirmed) {
        throw 'Before Apply, pause relevant automation and prevent concurrent edits; then pass -MaintenanceWindowConfirmed.'
    }
    $PreviewFolder = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($PreviewFolder)
    $ChecklistPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ChecklistPath)
    $OutputFolder = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputFolder)
    $approved = Read-SystemEditorApprovedPlan $PreviewFolder $ChecklistPath
    Import-Module PnP.PowerShell -MinimumVersion 3.3.0 -ErrorAction Stop
    $run = Assert-SPOutputPath (Join-Path $OutputFolder ((Get-Date -Format 'yyMMdd-HHmmss') + '-' + [guid]::NewGuid().ToString('N').Substring(0, 8)))
    [void][IO.Directory]::CreateDirectory($run)
    $journal = Join-Path $run 'Journal.csv'
    $backup = Join-Path $run 'Before.csv'
    'WebUrl,ListId,ItemId,Status,Detail' | Set-Content -LiteralPath $journal
    'WebUrl,ListId,ItemId,ObjectVersion,Field,ValueXml' | Set-Content -LiteralPath $backup
    $summary = [ordered]@{
        Mode = $(if ($Apply) { 'Apply' } else { 'RecheckOnly' }); Status = 'Running'
        PreviewFolder = $PreviewFolder
        Candidates = $approved.Rows.Count; Verified = 0; Eligible = 0; Skipped = 0; AlreadySystem = 0
        Errors = 0; ReportFolder = $run
    }
    Write-Host "$($summary.Mode): $($summary.Candidates) candidates."
    Write-Host "Reports: $run"
    $active = [pscustomobject]@{ WebUrl = ''; ListId = ''; ItemId = ''; Status = ''; Detail = '' }
    try {
        # Prove descendant membership live; a URL prefix under the tenant home is insufficient.
        $allowed = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        $connections = @{}
        foreach ($root in $approved.Roots) {
            Write-Progress -Id 2 -Activity 'Checking approved sites' -Status $root
            $active.WebUrl = $root
            $connection = Connect-PnPOnline -Url $root -ClientId $ClientId -Interactive -ReturnConnection -ErrorAction Stop
            $web = Get-PnPWeb -Connection $connection -Includes Url -ErrorAction Stop
            if ((ConvertTo-SystemEditorWebUrl $web.Url) -ine $root) { throw 'Connected root differs from checklist.' }
            $connections[$root] = $connection
            [void]$allowed.Add($root)
            Get-PnPSubWeb -Recurse -Connection $connection -Includes Url -ErrorAction Stop | ForEach-Object {
                $child = ConvertTo-SystemEditorWebUrl $_.Url
                if (-not $child.StartsWith($root + '/', [StringComparison]::OrdinalIgnoreCase)) {
                    throw "Subsite outside approved root: $child"
                }
                [void]$allowed.Add($child)
            }
        }
        foreach ($row in $approved.Rows) {
            if (-not $allowed.Contains($row.WebUrl.TrimEnd('/'))) { throw "Candidate web is no longer in approved structure: $($row.WebUrl)" }
        }
        foreach ($group in $approved.Rows | Group-Object WebUrl) {
            $url = $group.Name.TrimEnd('/')
            $active = [pscustomobject]@{ WebUrl = $url; ListId = ''; ItemId = ''; Status = ''; Detail = '' }
            $startVerified = $summary.Verified
            $startEligible = $summary.Eligible
            $startSkipped = $summary.Skipped + $summary.AlreadySystem
            $connection = if ($connections.ContainsKey($url)) { $connections[$url] } else {
                Connect-PnPOnline -Url $url -ClientId $ClientId -Interactive -ReturnConnection -ErrorAction Stop
            }
            $web = Get-PnPWeb -Connection $connection -Includes Url -ErrorAction Stop
            if ((ConvertTo-SystemEditorWebUrl $web.Url) -ine $url) { throw 'Connected candidate web differs from plan.' }
            $system = Get-PnPUser -Identity 'SHAREPOINT\system' -Connection $connection -ErrorAction Stop
            if (-not $system -or $system.LoginName -ine 'SHAREPOINT\system' -or $system.Id -le 0) {
                throw 'Native System Account unavailable. No substitute permitted.'
            }
            foreach ($listGroup in $group.Group | Group-Object ListId) {
                $listId = [guid]$listGroup.Name
                $fields = @(Get-PnPField -List $listId -Connection $connection -ErrorAction Stop)
                $verifyNames = @(@($fields | Where-Object {
                    -not $_.ReadOnlyField -and $_.TypeAsString -notin 'Computed', 'Counter' -and
                    $_.InternalName -notin 'Editor', 'owshiddenversion'
                } | ForEach-Object InternalName) + @(
                    'Modified', 'Created', 'Author', 'UniqueId', '_UIVersionString', '_UIVersion',
                    'ContentTypeId', 'FSObjType', 'FileRef', 'FileLeafRef', 'File_x0020_Size', '_ModerationStatus'
                ) | Sort-Object -Unique)
                foreach ($row in $listGroup.Group) {
                    $active = [pscustomobject]@{ WebUrl = $url; ListId = [string]$listId; ItemId = $row.ItemId; Status = 'Reading'; Detail = '' }
                    Write-Progress -Id 2 -Activity "$($summary.Mode): $url" -Status "Item $($row.ItemId) - $($summary.Verified) verified; $($summary.Skipped) skipped"
                    $item = Get-PnPListItem -List $listId -Id ([int]$row.ItemId) -Connection $connection -ErrorAction Stop
                    foreach ($key in 'UniqueId', 'Modified', 'Created', 'Author', 'Editor', '_UIVersionString') {
                        if (-not $item[$key]) { throw "Required field missing: $key" }
                    }
                    $editor = Get-PnPUser -Identity $item['Editor'].LookupId -Connection $connection -ErrorAction Stop
                    $status = Get-SystemEditorCandidateStatus -Item $item -Row $row -Editor $editor -SystemId $system.Id
                    if ($status -ne 'Eligible') {
                        if ($status -eq 'Already System Account') { $summary.AlreadySystem++ } else { $summary.Skipped++ }
                        $active.Status = $status
                        $active | Export-Csv -LiteralPath $journal -Append -NoTypeInformation
                        continue
                    }
                    $before = Get-SystemEditorFieldSnapshot $item
                    foreach ($name in $verifyNames) {
                        if ($name -in $fields.InternalName -and -not $before.ContainsKey($name)) {
                            throw "Field not loaded for preservation check: $name"
                        }
                    }
                    if (-not $Apply) {
                        $summary.Eligible++
                        $active.Status = 'Eligible (no write)'
                        $active | Export-Csv -LiteralPath $journal -Append -NoTypeInformation
                        continue
                    }
                    $before.Keys | Sort-Object | ForEach-Object {
                        $name = $_
                        [pscustomobject]@{
                            WebUrl = $url; ListId = [string]$listId; ItemId = $row.ItemId
                            ObjectVersion = $item.ObjectVersion; Field = $name; ValueXml = $before[$name]
                        }
                    } | Export-Csv -LiteralPath $backup -Append -NoTypeInformation
                    # Re-read after backup I/O; never write through the re-read object.
                    $latest = Get-PnPListItem -List $listId -Id ([int]$row.ItemId) -Connection $connection -ErrorAction Stop
                    $latestState = Get-SystemEditorFieldSnapshot $latest
                    $different = @($before.Keys | Where-Object { $before[$_] -cne $latestState[$_] })
                    if ($item.ObjectVersion -cne $latest.ObjectVersion -or $before.Count -ne $latestState.Count -or $different.Count) {
                        $summary.Skipped++
                        $active.Status = 'Skipped: changed during recheck'
                        $active | Export-Csv -LiteralPath $journal -Append -NoTypeInformation
                        continue
                    }
                    $active.Status = 'Write started'
                    $active | Export-Csv -LiteralPath $journal -Append -NoTypeInformation
                    Set-SystemEditorItemValues -Item $item -SystemId $system.Id -Modified (ConvertTo-SystemEditorUtc $item['Modified'])
                    $active.Status = 'Verifying'
                    $after = Get-PnPListItem -List $listId -Id ([int]$row.ItemId) -Connection $connection -ErrorAction Stop
                    if ($after['Editor'].LookupId -ne $system.Id) { throw 'Editor verification failed.' }
                    $afterState = Get-SystemEditorFieldSnapshot $after
                    $changedFields = @($verifyNames | Where-Object {
                        $before.ContainsKey($_) -and $before[$_] -cne $afterState[$_]
                    })
                    if ($changedFields.Count) {
                        $afterState.Keys | Sort-Object | ForEach-Object {
                            [pscustomobject]@{
                                WebUrl = $url; ListId = [string]$listId; ItemId = $row.ItemId
                                Field = $_; ValueXml = $afterState[$_]
                            }
                        } | Export-Csv -LiteralPath (Join-Path $run 'FailedAfter.csv') -Append -NoTypeInformation
                    }
                    foreach ($name in $verifyNames | Where-Object { $before.ContainsKey($_) }) {
                        if ($before[$name] -cne $afterState[$name]) { throw "Preservation verification failed: $name" }
                    }
                    $summary.Verified++
                    $active.Status = 'Updated and verified'
                    $active | Export-Csv -LiteralPath $journal -Append -NoTypeInformation
                }
            }
            $count = if ($Apply) { $summary.Verified - $startVerified } else { $summary.Eligible - $startEligible }
            $skipped = $summary.Skipped + $summary.AlreadySystem - $startSkipped
            Write-Host "$(([uri]$url).AbsolutePath): $count $(if ($Apply) { 'verified' } else { 'eligible' }); $skipped skipped/already System."
        }
        $summary.Status = if ($summary.Skipped) { 'CompletedWithSkips' } else { 'Completed' }
    } catch {
        $summary.Status = 'Stopped'
        $summary.Errors++
        $active.Detail = $_.Exception.Message
        $active.Status = "Failed ($($active.Status))"
        $active | Export-Csv -LiteralPath $journal -Append -NoTypeInformation
        throw "Stopped. Inspect Journal.csv and live state before retrying. $($_.Exception.Message)"
    } finally {
        Write-Progress -Id 2 -Activity 'System Editor remediation' -Completed
        if ($summary.Status -eq 'Running') { $summary.Status = 'Interrupted' }
        Write-SPBoundedJson -InputObject $summary -Path (Join-Path $run 'Summary.json') -Depth 5
        Write-Host "$($summary.Status): $($summary.Verified) verified; $($summary.Eligible) eligible (no write); $($summary.Skipped) skipped; $($summary.AlreadySystem) already System."
    }
    return [pscustomobject]$summary
}

if ($MyInvocation.InvocationName -ne '.') {
    if ([string]::IsNullOrWhiteSpace($PreviewFolder)) { throw '-PreviewFolder is required.' }
    Invoke-SystemEditorApply -PreviewFolder $PreviewFolder -ChecklistPath $ChecklistPath -OutputFolder $OutputFolder `
        -ClientId $ClientId -Apply:$Apply -MaintenanceWindowConfirmed:$MaintenanceWindowConfirmed | Out-Null
}
