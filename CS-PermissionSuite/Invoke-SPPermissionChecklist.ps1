#requires -Version 7.4.6
<#
.SYNOPSIS
    Process checklist sites and matched descendant webs sequentially, never the tenant.
.DESCRIPTION
    Match generates content and subweb proposals. DryRun is the default; Apply is explicit.
    Auto explicitly authorizes per-web Match -> approve Score >= 0.80 -> DryRun ->
    Apply -> measured verification. Lower scores stay unapproved. Score is similarity,
    not a probability of identity. Auto can approve ambiguous high-scoring proposals.
    Each web is completely inventoried, planned and reported (and applied in Apply mode)
    before the next web starts. Explicit checklist mappings override descendant matching.
    Exact child-web paths are automatic; renamed child webs require Approved=Yes in
    that parent's WebReview.csv. No counterpart or no approval means skip, never create.
    Descendants of unmatched webs are not searched across unrelated web trees.

    Keep the two permission scripts and SPPermission.Common.ps1 beside this runner.
    Copy the checklist to the server. ReviewRoot contains stable per-web MatchReview.csv
    and WebReview.csv files; OutputFolder contains timestamped execution reports.
    Existing review files are NEVER overwritten. Match writes fresh candidates to each
    run folder; explicitly replace/merge review files if refreshing proposals.
.EXAMPLE
    .\Invoke-SPPermissionChecklist.ps1 -ChecklistFile .\NewSiteChecklist.csv -Mode Match -OutputFolder .\PermissionRuns -ReviewRoot .\PermissionReviews
.EXAMPLE
    .\Invoke-SPPermissionChecklist.ps1 -ChecklistFile .\NewSiteChecklist.csv -Mode DryRun -OutputFolder .\PermissionRuns -ReviewRoot .\PermissionReviews
.EXAMPLE
    .\Invoke-SPPermissionChecklist.ps1 -ChecklistFile .\NewSiteChecklist.csv -Mode Apply -OutputFolder .\PermissionRuns -ReviewRoot .\PermissionReviews
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$ChecklistFile,
    [ValidateSet('Match', 'DryRun', 'Apply', 'Auto')][string]$Mode = 'DryRun',
    [string]$OutputFolder = (Join-Path $PSScriptRoot 'PermissionRuns'),
    [string]$ReviewRoot = (Join-Path $PSScriptRoot 'PermissionReviews'),
    [string[]]$TargetSites,
    [string]$ClientId = '2e20d82f-56ba-418e-b465-9718493f0df8',
    [string]$Tenant,
    [string]$SourceGroupToMap = 'Team Site Members',
    [string[]]$ExcludePaths = @('Access Requests'),
    [ValidateRange(1, 5000)][int]$PageSize = 500,
    [ValidateRange(1, 1024)][int]$MaxContentMB = 25,
    [switch]$NoContent,
    [switch]$SkipPersistLoginSetup
)
. (Join-Path $PSScriptRoot 'SPPermission.Common.ps1')

function Set-SPChecklistAutoApprovals {
    param([string]$Path, [switch]$Webs)
    $rows = @(Import-Csv -LiteralPath $Path -ErrorAction Stop)
    foreach ($row in $rows) {
        if ($null -eq $row.PSObject.Properties['ApprovalPolicy']) {
            $row | Add-Member -NotePropertyName ApprovalPolicy -NotePropertyValue 'Manual review'
        }
        if ($row.Approved -match '^(?i:Yes|True|1|Y)$') { continue }
        $score = 0.0
        if (-not [double]::TryParse($row.Score, [Globalization.NumberStyles]::Float,
            [Globalization.CultureInfo]::InvariantCulture, [ref]$score) -or $score -lt 0 -or $score -gt 1) {
            throw "Invalid match score in '$Path': $($row.Score)"
        }
        if ($score -ge 0.80) {
            $row.Approved = 'Yes'
            $row.ApprovalPolicy = 'Automatic Score >= 0.80'
        }
    }
    if (-not $Webs) {
        # A high-scoring child cannot authorize an unapproved renamed container.
        foreach ($row in $rows | Sort-Object { $_.SourcePath.Length }) {
            $parents = @($rows | Where-Object {
                $_.Kind -in @('Library', 'List', 'Folder') -and $_.SourcePath -ine $row.SourcePath -and
                (Test-SPPermissionPathWithin $row.SourcePath $_.SourcePath)
            })
            if (@($parents | Where-Object { $_.Approved -notmatch '^(?i:Yes|True|1|Y)$' }).Count) {
                $row.Approved = 'No'
                $row.ApprovalPolicy = 'Manual review - renamed parent not approved'
            }
        }
    }
    if ($rows.Count) { $rows | Export-Csv -LiteralPath $Path -NoTypeInformation -Encoding utf8 }
}

function Assert-SPChecklistDryRun {
    param([string]$Folder)
    $plans = @(Get-ChildItem -LiteralPath $Folder -Filter 'PermSync_Plan_DRYRUN_*.json')
    $principals = @(Get-ChildItem -LiteralPath $Folder -Filter 'PermSync_UnavailablePrincipals_DRYRUN_*.csv')
    if ($plans.Count -ne 1 -or $principals.Count -ne 1) { throw "Incomplete automatic preflight in '$Folder'." }
    $plan = Get-Content -LiteralPath $plans[0].FullName -Raw | ConvertFrom-Json -Depth 100
    if (@($plan.Notes | Where-Object Result -eq 'Blocked').Count -or
        @(Import-Csv -LiteralPath $principals[0].FullName).Count) {
        throw "Automatic Apply stopped: blocked scopes or unavailable principals in '$Folder'."
    }
}

function Invoke-SPChecklistReport {
    param([hashtable]$Options)
    & (Join-Path $PSScriptRoot 'Export-SPPermissionSiteReport.ps1') @Options
}

function Get-SPChecklistPairs {
    param([string]$Path)
    $sources = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $targets = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $rows = @(Import-Csv -LiteralPath $Path -ErrorAction Stop)
    if (-not $rows.Count) { throw 'The site checklist is empty.' }
    foreach ($row in $rows) {
        if (-not $row.'Old Site URL' -or -not $row.'New Site URL') { throw 'Checklist needs Old Site URL and New Site URL on every row.' }
        $sourceUrl = $row.'Old Site URL'.Trim().TrimEnd('/')
        $targetUrl = $row.'New Site URL'.Trim().TrimEnd('/')
        Assert-SPPermissionUrls $sourceUrl $targetUrl
        if (-not $sources.Add($sourceUrl) -or -not $targets.Add($targetUrl)) { throw 'Checklist site mappings must be one-to-one.' }
        [pscustomobject]@{ SourceUrl = $sourceUrl; TargetUrl = $targetUrl; Method = 'Checklist' }
    }
}

function Get-SPChecklistSiteKey {
    param([string]$Url)
    $bytes = [Text.Encoding]::UTF8.GetBytes($Url.TrimEnd('/').ToLowerInvariant())
    $hash = [Convert]::ToHexString([Security.Cryptography.SHA256]::HashData($bytes)).Substring(0, 12)
    $leaf = (([uri]$Url).AbsolutePath.Trim('/') -split '/')[-1] -replace '[^a-zA-Z0-9_-]', '_'
    if (-not $leaf) { $leaf = 'Root' }
    "$leaf-$hash"
}

function Get-SPChecklistChildren {
    param($Connection, [string]$Url)
    $web = Invoke-SPPermissionRest $Connection ($Url + '/_api/web?$select=Url')
    if (-not $web.Url -or $web.Url.TrimEnd('/') -ine $Url) { throw "Unexpected selected web: $Url" }
    foreach ($child in @(Get-SPPermissionCollection $Connection (
        $Url + '/_api/web/webs?$select=Id,Title,Url,ServerRelativeUrl,HasUniqueRoleAssignments'
    ))) {
        if (-not $child.Url -or $null -eq $child.HasUniqueRoleAssignments -or
            -not (Test-SPPermissionPathWithin $child.Url.TrimEnd('/') $Url) -or $child.Url.TrimEnd('/') -ieq $Url) {
            throw "Invalid child-web metadata under '$Url'."
        }
        [pscustomobject]@{
            Url = $child.Url.TrimEnd('/'); Title = [string]$child.Title
            RelativePath = $child.Url.TrimEnd('/').Substring($Url.Length).TrimStart('/')
            Unique = [bool]$child.HasUniqueRoleAssignments
        }
    }
}

function Select-SPChecklistChildren {
    param([object[]]$SourceChildren, [object[]]$TargetChildren, [object[]]$Checklist,
        [object[]]$ApprovedRows, [string]$SourceParent, [string]$TargetParent)
    $reservedSources = [Collections.Generic.HashSet[string]]::new([string[]]@($Checklist.SourceUrl), [StringComparer]::OrdinalIgnoreCase)
    $reservedTargets = [Collections.Generic.HashSet[string]]::new([string[]]@($Checklist.TargetUrl), [StringComparer]::OrdinalIgnoreCase)
    $sourcePool = @($SourceChildren | Where-Object { -not $reservedSources.Contains($_.Url) })
    $targetPool = @($TargetChildren | Where-Object { -not $reservedTargets.Contains($_.Url) })
    $pairs = [Collections.Generic.List[object]]::new()
    $usedSource = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $usedTarget = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($sourceChild in $sourcePool | Sort-Object Url) {
        $exact = @($targetPool | Where-Object RelativePath -IEQ $sourceChild.RelativePath)
        if ($exact.Count -gt 1) { throw 'Duplicate child-web path.' }
        if ($exact.Count) {
            $pairs.Add([pscustomobject]@{ SourceUrl = $sourceChild.Url; TargetUrl = $exact[0].Url; Method = 'Exact subweb path' })
            [void]$usedSource.Add($sourceChild.Url); [void]$usedTarget.Add($exact[0].Url)
        }
    }
    foreach ($row in $ApprovedRows) {
        foreach ($field in @('Approved', 'SourceUrl', 'TargetUrl', 'SourceParent', 'TargetParent')) {
            if ($null -eq $row.PSObject.Properties[$field]) { throw "WebReview.csv is missing $field." }
        }
        if ($row.SourceParent -ine $SourceParent -or $row.TargetParent -ine $TargetParent) { throw 'WebReview.csv belongs to another web pair.' }
        if ($row.Approved.Trim() -notmatch '^(?i:Yes|No|True|False|1|0|Y|N)?$') { throw "Invalid web approval: $($row.Approved)" }
        if ($row.Approved.Trim() -notmatch '^(?i:Yes|True|1|Y)$') { continue }
        $sourceChild = @($sourcePool | Where-Object Url -IEQ $row.SourceUrl)
        $targetChild = @($targetPool | Where-Object Url -IEQ $row.TargetUrl)
        if ($sourceChild.Count -ne 1 -or $targetChild.Count -ne 1) { throw 'Approved subweb missing or reserved by the checklist.' }
        if (-not $usedSource.Add($row.SourceUrl) -or -not $usedTarget.Add($row.TargetUrl)) { throw 'Approved subweb conflicts with an exact match or another approval.' }
        $pairs.Add([pscustomobject]@{ SourceUrl = $row.SourceUrl; TargetUrl = $row.TargetUrl; Method = 'Approved fuzzy subweb' })
    }
    $proposals = [Collections.Generic.List[object]]::new()
    foreach ($sourceChild in $sourcePool | Where-Object { -not $usedSource.Contains($_.Url) } | Sort-Object Url) {
        $candidates = @($targetPool | Where-Object { -not $usedTarget.Contains($_.Url) } | ForEach-Object {
            $score = [math]::Max((Get-SPNameSimilarity $sourceChild.Title $_.Title),
                (Get-SPNameSimilarity $sourceChild.RelativePath $_.RelativePath))
            [pscustomobject]@{ Target = $_; Score = $score }
        } | Sort-Object @{ Expression = 'Score'; Descending = $true }, @{ Expression = { $_.Target.Url } })
        if (-not $candidates.Count -or $candidates[0].Score -lt 0.55) { continue }
        $best = $candidates[0]
        $proposals.Add([pscustomobject]@{
            Approved = 'No'; SourceParent = $SourceParent; TargetParent = $TargetParent
            SourceUrl = $sourceChild.Url; TargetUrl = $best.Target.Url; Score = [math]::Round($best.Score, 3)
            Evidence = 'Child web title/URL name similarity; manually confirm identity'
            RunnerUp = $(if ($candidates.Count -gt 1) { $candidates[1].Target.Url } else { '' })
        })
        # Reserve proposals only for other proposals; never authorize them implicitly.
        [void]$usedTarget.Add($best.Target.Url)
    }
    [pscustomobject]@{ Pairs = $pairs.ToArray(); Proposals = $proposals.ToArray() }
}

function Invoke-SPChecklistWeb {
    param($Pair, [string]$RunFolder, [string]$SiteReviewFolder, [hashtable]$Options, [string]$RunMode)
    $arguments = $Options.Clone()
    $arguments.SourceUrl = $Pair.SourceUrl; $arguments.TargetUrl = $Pair.TargetUrl
    $arguments.OutputFolder = $RunFolder
    if ($RunMode -eq 'Auto') {
        foreach ($stage in @('Match', 'DryRun', 'Apply')) {
            $stageFolder = Join-Path $RunFolder $stage
            [void][IO.Directory]::CreateDirectory($stageFolder)
            $stageOptions = $Options.Clone()
            if ($stage -eq 'Apply') { $stageOptions.RequireCompletePlan = $true }
            Invoke-SPChecklistWeb $Pair $stageFolder $SiteReviewFolder $stageOptions $stage
            if ($stage -eq 'Match') {
                Set-SPChecklistAutoApprovals (Join-Path $SiteReviewFolder 'MatchReview.csv')
            } elseif ($stage -eq 'DryRun') {
                Assert-SPChecklistDryRun $stageFolder
            } elseif ($stage -eq 'Apply') {
                $verifyFolder = Join-Path $RunFolder 'Verify'
                [void][IO.Directory]::CreateDirectory($verifyFolder)
                Invoke-SPChecklistWeb $Pair $verifyFolder $SiteReviewFolder $Options 'DryRun'
                Assert-SPChecklistDryRun $verifyFolder
                $verifyPlan = Get-ChildItem -LiteralPath $verifyFolder -Filter 'PermSync_Plan_DRYRUN_*.json'
                $remaining = Get-Content -LiteralPath $verifyPlan.FullName -Raw | ConvertFrom-Json -Depth 100
                if (@($remaining.Steps).Count) { throw "Post-Apply verification found remaining actions in '$verifyFolder'." }
            }
        }
        return
    }
    if ($RunMode -eq 'Match') {
        [void]$arguments.Remove('SourceGroupToMap')
        & (Join-Path $PSScriptRoot 'New-SPPermissionMatchMap.ps1') @arguments
        $csv = @(Get-ChildItem -LiteralPath $RunFolder -Filter 'PermMatch_Review_*.csv')
        if ($csv.Count -ne 1) { throw 'Expected exactly one generated content review CSV.' }
        $reviewPath = Join-Path $SiteReviewFolder 'MatchReview.csv'
        if (-not (Test-Path -LiteralPath $reviewPath)) { Copy-Item -LiteralPath $csv[0].FullName -Destination $reviewPath -ErrorAction Stop }
        else { Write-Warning "Existing $reviewPath retained. Fresh candidates: $($csv[0].FullName)" }
    } else {
        [void]$arguments.Remove('NoContent'); [void]$arguments.Remove('MaxContentMB')
        $reviewPath = Join-Path $SiteReviewFolder 'MatchReview.csv'
        if (Test-Path -LiteralPath $reviewPath) { $arguments.MatchFile = $reviewPath }
        else { Write-Warning "No content review at $reviewPath; exact content paths only." }
        $arguments.Apply = $RunMode -eq 'Apply'
        & (Join-Path $PSScriptRoot 'Sync-SPPermissionsFromSource.ps1') @arguments
    }
}

function Invoke-SPPermissionChecklist {
    $ErrorActionPreference = 'Stop'
    $OutputFolder = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputFolder)
    $ReviewRoot = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ReviewRoot)
    $checklist = @(Get-SPChecklistPairs $ChecklistFile)
    $jobs = @($checklist | Sort-Object @{ Expression = { ([uri]$_.TargetUrl).AbsolutePath.Split('/').Count } }, TargetUrl)
    if ($TargetSites.Count) {
        foreach ($url in $TargetSites) {
            if ($url.TrimEnd('/') -inotin $checklist.TargetUrl) { throw "TargetSites must be checklist URLs: $url" }
        }
        $selected = @($TargetSites | ForEach-Object { $_.TrimEnd('/') })
        $jobs = @($jobs | Where-Object { $_.TargetUrl -iin $selected })
    }
    Import-Module PnP.PowerShell -MinimumVersion 3.3.0 -ErrorAction Stop
    $stamp = Get-Date -Format 'yyyyMMdd-HHmmss-fff'
    $batchFolder = if ($Mode -eq 'Auto') { $OutputFolder } else {
        Join-Path $OutputFolder ("b-" + (Get-Date -Format 'yyMMddHHmmssfff'))
    }
    [void][IO.Directory]::CreateDirectory($batchFolder)
    $summary = [Collections.Generic.List[object]]::new()
    $visitedSources = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $visitedTargets = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $queue = [Collections.Generic.List[object]]::new()
    foreach ($job in $jobs) { $queue.Add($job) }
    $options = @{
        ClientId = $ClientId; Tenant = $Tenant; SourceGroupToMap = $SourceGroupToMap
        ExcludePaths = $ExcludePaths; PageSize = $PageSize; MaxContentMB = $MaxContentMB
        NoContent = $NoContent; SkipPersistLoginSetup = $SkipPersistLoginSetup
    }
    $applicationError = $null
    try {
        for ($index = 0; $index -lt $queue.Count; $index++) {
            $pair = $queue[$index]
            if (-not $visitedSources.Add($pair.SourceUrl) -or -not $visitedTargets.Add($pair.TargetUrl)) { throw 'Repeated web mapping refused.' }
            $siteKey = Get-SPChecklistSiteKey $pair.TargetUrl
            $runFolder = Join-Path $batchFolder ("w-" + $siteKey.Split('-')[-1])
            $siteReviewFolder = Join-Path $ReviewRoot $siteKey
            [void][IO.Directory]::CreateDirectory($runFolder)
            [void][IO.Directory]::CreateDirectory($siteReviewFolder)
            Write-Host "[$Mode] $($pair.SourceUrl) -> $($pair.TargetUrl)"
            $webOptions = $options.Clone()
            if ($pair.TargetUrl -ieq 'https://whallc.sharepoint.com/sites/ClinicalandPatientServices') {
                $webOptions.ExcludePaths = @($ExcludePaths) + @('COVID19 Resources', 'COVID19 Workflows')
            }
            $sourceConnection = Connect-SPPermissionSite $pair.SourceUrl $ClientId $Tenant -SkipPersistLoginSetup:$SkipPersistLoginSetup
            $targetConnection = Connect-SPPermissionSite $pair.TargetUrl $ClientId $Tenant -SkipPersistLoginSetup:$SkipPersistLoginSetup
            try {
                $sourceChildren = @(Get-SPChecklistChildren $sourceConnection $pair.SourceUrl)
                $targetChildren = @(Get-SPChecklistChildren $targetConnection $pair.TargetUrl)
            } catch {
                if ($Mode -eq 'Auto') { throw }
                $status = $_.Exception.Response.StatusCode
                if ([string]$status -notin @('404', 'NotFound') -and $_.Exception.Message -notmatch '\b404\b') { throw }
                $summary.Add([pscustomobject]@{
                    SourceUrl = $pair.SourceUrl; TargetUrl = $pair.TargetUrl; Status = 'Skipped - web not found'
                    Detail = $_.Exception.Message; ReportFolder = $runFolder
                })
                Write-Warning "Web pair missing; skipping: $($pair.SourceUrl) -> $($pair.TargetUrl)"
                continue
            }
            $webReview = Join-Path $siteReviewFolder 'WebReview.csv'
            if ($Mode -eq 'Auto' -and (Test-Path -LiteralPath $webReview)) {
                Set-SPChecklistAutoApprovals $webReview -Webs
            }
            $approvals = if (Test-Path -LiteralPath $webReview) { @(Import-Csv -LiteralPath $webReview -ErrorAction Stop) } else { @() }
            $selection = Select-SPChecklistChildren $sourceChildren $targetChildren $checklist $approvals $pair.SourceUrl $pair.TargetUrl
            $columns = @('Approved','SourceParent','TargetParent','SourceUrl','TargetUrl','Score','Evidence','RunnerUp')
            $webCandidates = Join-Path $runFolder 'WebCandidates.csv'
            Export-SPPermissionCsv $selection.Proposals $columns $webCandidates
            if ($Mode -in @('Match', 'Auto') -and -not (Test-Path -LiteralPath $webReview)) {
                Copy-Item -LiteralPath $webCandidates -Destination $webReview -ErrorAction Stop
                if ($Mode -eq 'Auto') {
                    Set-SPChecklistAutoApprovals $webReview -Webs
                    $approvals = @(Import-Csv -LiteralPath $webReview)
                    $selection = Select-SPChecklistChildren $sourceChildren $targetChildren $checklist $approvals $pair.SourceUrl $pair.TargetUrl
                }
            }
            $webInventory = @(
                foreach ($side in @('Source', 'Target')) {
                    $children = if ($side -eq 'Source') { $sourceChildren } else { $targetChildren }
                    $explicit = if ($side -eq 'Source') { $checklist.SourceUrl } else { $checklist.TargetUrl }
                    $paired = if ($side -eq 'Source') { $selection.Pairs.SourceUrl } else { $selection.Pairs.TargetUrl }
                    foreach ($child in $children) {
                        $state = if ($child.Url -iin $explicit) { 'Separate checklist entry' } elseif ($child.Url -iin $paired) { 'Matched - queued' } else { 'Unmatched/unapproved - skipped' }
                        [pscustomobject]@{ Side = $side; Url = $child.Url; Unique = $child.Unique; Status = $state }
                    }
                }
            )
            Export-SPPermissionCsv $webInventory @('Side','Url','Unique','Status') (Join-Path $runFolder 'WebInventory.csv')
            Invoke-SPChecklistWeb $pair $runFolder $siteReviewFolder $webOptions $Mode
            $summary.Add([pscustomobject]@{
                SourceUrl = $pair.SourceUrl; TargetUrl = $pair.TargetUrl; Status = "$Mode completed - review site reports"
                Detail = "Matched descendant webs queued: $($selection.Pairs.Count)"; ReportFolder = $runFolder
            })
            # Depth-first, but never begin a child until its parent's reports/writes finish.
            $offset = 1
            foreach ($childPair in $selection.Pairs | Sort-Object TargetUrl) {
                $queue.Insert($index + $offset, $childPair)
                $offset++
            }
        }
    } catch {
        $applicationError = $_
        $summary.Add([pscustomobject]@{
            SourceUrl = $pair.SourceUrl; TargetUrl = $pair.TargetUrl; Status = 'Stopped - error'
            Detail = $_.Exception.Message; ReportFolder = $runFolder
        })
        throw
    } finally {
        $summaryPath = Join-Path $batchFolder 'ChecklistSummary.csv'
        Export-SPPermissionCsv $summary.ToArray() @('SourceUrl','TargetUrl','Status','Detail','ReportFolder') $summaryPath
        Write-Host "Checklist summary: $summaryPath"
        if ($Mode -in @('Apply', 'Auto')) {
            $reportFailures = [Collections.Generic.List[string]]::new()
            foreach ($job in $jobs) {
                $reportFolder = Join-Path $batchFolder ("r-" + (Get-SPChecklistSiteKey $job.TargetUrl).Split('-')[-1])
                [void][IO.Directory]::CreateDirectory($reportFolder)
                $reportOptions = @{
                    TargetUrl = $job.TargetUrl; OutputFolder = $reportFolder; ClientId = $ClientId
                    Tenant = $Tenant; PageSize = $PageSize; SkipPersistLoginSetup = $SkipPersistLoginSetup
                }
                $changeRows = @(
                    foreach ($entry in $summary | Where-Object { Test-SPPermissionPathWithin $_.TargetUrl $job.TargetUrl }) {
                        foreach ($file in Get-ChildItem -LiteralPath $entry.ReportFolder -Recurse -Filter 'PermSync_Log_APPLY_*.csv') {
                            foreach ($row in Import-Csv -LiteralPath $file.FullName) {
                                [pscustomobject]@{
                                    WebUrl = $entry.TargetUrl; Path = $row.Path; Action = $row.Action
                                    Result = $row.Result; Detail = $row.Detail; EvidenceFile = $file.FullName
                                }
                            }
                        }
                    }
                )
                Export-SPPermissionCsv $changeRows @('WebUrl','Path','Action','Result','Detail','EvidenceFile') (Join-Path $reportFolder 'Changes.csv')
                try { Invoke-SPChecklistReport $reportOptions }
                catch {
                    $reportFailures.Add("$($job.TargetUrl): $($_.Exception.Message)")
                    Write-Warning "Post-Apply report FAILED for '$($job.TargetUrl)': $($_.Exception.Message)"
                }
            }
            if ($reportFailures.Count -and -not $applicationError) {
                throw "Post-Apply reporting incomplete: $($reportFailures -join '; '). Inspect ChecklistSummary.csv for the application outcome."
            }
        }
    }
}

if ($MyInvocation.InvocationName -ne '.') { Invoke-SPPermissionChecklist }
