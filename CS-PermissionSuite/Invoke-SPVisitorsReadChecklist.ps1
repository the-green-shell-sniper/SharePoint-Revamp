#requires -Version 7.4.6
<#
.SYNOPSIS
    Sequential additive Match -> DryRun -> verified Apply for remaining checklist sites.
.DESCRIPTION
    No Scout or full permission mirror. Default mode is read-only. -Apply authorizes
    the complete pipeline. Failures stop that site's pipeline, not unrelated sites.
    Previously completed CS, Accounting and PFS are excluded by default.
#>
[CmdletBinding()]
param(
    [string]$ChecklistFile = (Join-Path $PSScriptRoot '..\ReferenceFiles\NewSiteChecklist.csv'),
    [string]$OutputFolder = (Join-Path $PSScriptRoot '..\Reports\VisitorsRuns'),
    [string[]]$TargetSites,
    [string[]]$CompletedTargetSites = @(
        'https://whallc.sharepoint.com/sites/ClinicalandPatientServices',
        'https://whallc.sharepoint.com/sites/Accounting',
        'https://whallc.sharepoint.com/sites/PatientFinancialServices'
    ),
    [string]$ClientId = '2e20d82f-56ba-418e-b465-9718493f0df8',
    [string]$Tenant,
    [switch]$SkipPersistLoginSetup,
    [switch]$Apply
)

. (Join-Path $PSScriptRoot 'SPPermission.Common.ps1')

function Get-SPVisitorsChecklistPairs {
    param([string]$Path)
    & {
        param($ScriptFolder, $File)
        . (Join-Path $ScriptFolder 'Invoke-SPPermissionChecklist.ps1') -ChecklistFile $File
        Get-SPChecklistPairs $File
    } $PSScriptRoot $Path
}

function Assert-SPVisitorsStage {
    param([string]$Folder, [string]$Stage)
    $rows = @(Import-Csv -LiteralPath (Join-Path $Folder 'Summary.csv') -ErrorAction Stop)
    if (-not $rows.Count -or @($rows | Where-Object {
        $_.Mode -ne $Stage -or $_.Status -ne 'Completed - inspect coverage' -or $_.Error
    }).Count) { throw "Incomplete or failed $Stage summary in '$Folder'." }
    foreach ($row in $rows) {
        if ($Stage -eq 'Match') {
            if (@(Get-ChildItem -LiteralPath $row.EvidenceFolder -Filter 'PermMatch_Review_*.csv').Count -ne 1) {
                throw "Missing or ambiguous Match review at '$($row.EvidenceFolder)'."
            }
            continue
        }
        $runs = @(Get-ChildItem -LiteralPath $row.EvidenceFolder -Directory -Filter 'rd-*')
        if ($runs.Count -ne 1) { throw "Missing or ambiguous $Stage evidence at '$($row.EvidenceFolder)'." }
        $status = Get-Content -LiteralPath (Join-Path $runs[0].FullName 'Status.json') -Raw |
            ConvertFrom-Json -ErrorAction Stop
        $expected = if ($Stage -eq 'Apply') {
            'Applied - matched additions verified; inspect coverage for skipped content'
        } else { 'DryRun complete - no writes' }
        if ($status.Mode -ne $Stage -or $status.Status -cne $expected -or $status.Error) {
            throw "$Stage did not verify successfully at '$($runs[0].FullName)'."
        }
        $plan = Get-Content -LiteralPath (Join-Path $runs[0].FullName 'Plan.json') -Raw |
            ConvertFrom-Json -Depth 100 -ErrorAction Stop
        if (@($plan.Steps).Count -ne $status.PlannedOperations -or
            @($plan.Coverage).Count -ne $status.SourceReadScopes -or
            @($plan.Notes | Where-Object Result -eq 'Blocked').Count) {
            throw "Inconsistent or blocked $Stage plan at '$($runs[0].FullName)'."
        }
        foreach ($step in $plan.Steps) {
            if ($step.Action -notin @('Add','Break','Freeze') -or
                ($step.Action -in @('Break','Freeze') -and $step.BeforeSignature -cne $step.AfterSignature) -or
                ($step.Action -eq 'Add' -and ($step.Operation.Type -ne 'Add' -or
                    $step.Operation.Role -ne 'Read' -or -not $step.Operation.MemberId -or
                    (Test-SPMigrationPrincipal $step.Operation.Login $step.Operation.Login)))) {
                throw "Unexpected non-additive operation at '$($step.Key)'."
            }
        }
    }
    if ($Stage -eq 'Apply') {
        $report = Get-Content -LiteralPath (Join-Path $Folder 'Report\ReportStatus.json') -Raw |
            ConvertFrom-Json -ErrorAction Stop
        if ($report.Status -ne 'Complete' -or $report.Error) { throw "Final report incomplete in '$Folder'." }
    }
}

function Invoke-SPVisitorsChecklistStage {
    param([hashtable]$Options)
    & (Join-Path $PSScriptRoot 'Invoke-PFSVisitorsRead.ps1') @Options
}

function Invoke-SPVisitorsReadChecklist {
    $ErrorActionPreference = 'Stop'
    $pairs = @(Get-SPVisitorsChecklistPairs $ChecklistFile)
    $targets = @($TargetSites | Where-Object { $_ } | ForEach-Object { $_.TrimEnd('/') })
    $completed = @($CompletedTargetSites | Where-Object { $_ } | ForEach-Object { $_.TrimEnd('/') })
    foreach ($target in $targets) {
        if ($target -inotin $pairs.TargetUrl) { throw "Target is not in the checklist: $target" }
    }
    $root = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputFolder)
    $folder = Join-Path $root ('v-' + (Get-Date -Format 'yyMMddHHmmssfff'))
    [void][IO.Directory]::CreateDirectory($folder)
    $summary = [Collections.Generic.List[object]]::new()
    $issues = [Collections.Generic.List[object]]::new()
    foreach ($pair in $pairs) {
        if (($targets.Count -and $pair.TargetUrl -inotin $targets) -or $pair.TargetUrl -iin $completed) {
            $summary.Add([pscustomobject]@{
                SourceUrl = $pair.SourceUrl; TargetUrl = $pair.TargetUrl; Stage = ''
                Status = 'Skipped - completed or not selected'; Error = ''; EvidenceFolder = ''
            })
            continue
        }
        $key = [Convert]::ToHexString([Security.Cryptography.SHA256]::HashData(
            [Text.Encoding]::UTF8.GetBytes($pair.TargetUrl.ToLowerInvariant()))).Substring(0, 12)
        $siteFolder = Join-Path $folder $key
        $stage = 'Match'
        try {
            foreach ($stage in @('Match','DryRun','Apply')) {
                if ($stage -eq 'Apply' -and -not $Apply) { break }
                $stageFolder = Join-Path $siteFolder $stage
                $excluded = @('Access Requests')
                if ($pair.TargetUrl -ieq 'https://whallc.sharepoint.com/sites/ClinicalandPatientServices') {
                    $excluded += @('COVID19 Resources','COVID19 Workflows')
                }
                if ($pair.TargetUrl -ieq 'https://whallc.sharepoint.com/sites/PatientFinancialServices') {
                    $excluded += '2021 EM Coding Changes'
                }
                $options = @{
                    SourceSiteUrl = $pair.SourceUrl; TargetSiteUrl = $pair.TargetUrl
                    SiteLabel = $key; Mode = $stage; OutputFolder = $stageFolder
                    ReviewRoot = (Join-Path $siteFolder 'Reviews'); DiscoverVisitorsTitle = $true
                    ReservedSourceUrls = @($pairs.SourceUrl); ReservedTargetUrls = @($pairs.TargetUrl)
                    ExcludedPaths = $excluded; ClientId = $ClientId; Tenant = $Tenant
                    SkipPersistLoginSetup = $SkipPersistLoginSetup
                }
                Write-Host "[$stage checklist] $($pair.TargetUrl)"
                Invoke-SPVisitorsChecklistStage $options
                $runs = @(Get-ChildItem -LiteralPath $stageFolder -Directory -Filter 'p-*')
                if ($runs.Count -ne 1) { throw "Expected one $stage run in '$stageFolder'." }
                Assert-SPVisitorsStage $runs[0].FullName $stage
                foreach ($file in @(Get-ChildItem -LiteralPath $runs[0].FullName -Recurse -File -Filter 'Coverage.csv')) {
                    foreach ($row in @(Import-Csv -LiteralPath $file.FullName | Where-Object Status -ne 'Eligible')) {
                        $issues.Add([pscustomobject]@{
                            TargetUrl = $pair.TargetUrl; Stage = $stage; Kind = 'Source Read coverage'
                            Path = $row.SourcePath; Detail = $row.Status; Evidence = $file.FullName
                        })
                    }
                }
                foreach ($row in @(Import-Csv -LiteralPath (Join-Path $runs[0].FullName 'Webs.csv') |
                    Where-Object Status -ne 'Exact path - queued')) {
                    $issues.Add([pscustomobject]@{
                        TargetUrl = $pair.TargetUrl; Stage = $stage; Kind = 'Web coverage'
                        Path = $row.SourceUrl; Detail = "$($row.Status): $($row.TargetUrl)"
                        Evidence = (Join-Path $runs[0].FullName 'Webs.csv')
                    })
                }
                if ($stage -eq 'Match') {
                    foreach ($file in @(Get-ChildItem -LiteralPath (Join-Path $siteFolder 'Reviews') -Recurse -Filter 'MatchReview.csv')) {
                        foreach ($row in @(Import-Csv -LiteralPath $file.FullName |
                            Where-Object Approved -NotMatch '^(?i:Yes|True|1|Y)$')) {
                            $issues.Add([pscustomobject]@{
                                TargetUrl = $pair.TargetUrl; Stage = $stage; Kind = 'Unapproved match'
                                Path = $row.SourcePath; Detail = "Score $($row.Score); target $($row.TargetPath)"
                                Evidence = $file.FullName
                            })
                        }
                    }
                }
            }
            $summary.Add([pscustomobject]@{
                SourceUrl = $pair.SourceUrl; TargetUrl = $pair.TargetUrl; Stage = $(if ($Apply) { 'Apply' } else { 'DryRun' })
                Status = $(if ($Apply) { 'Applied and verified - inspect stragglers' } else { 'DryRun verified - no writes' })
                Error = ''; EvidenceFolder = $siteFolder
            })
        } catch {
            Write-Warning "Stopped '$($pair.TargetUrl)' at ${stage}: $($_.Exception.Message). Continuing with the next checklist site."
            $summary.Add([pscustomobject]@{
                SourceUrl = $pair.SourceUrl; TargetUrl = $pair.TargetUrl; Stage = $stage
                Status = 'Failed - inspect evidence before retry'; Error = $_.Exception.Message; EvidenceFolder = $siteFolder
            })
            $issues.Add([pscustomobject]@{
                TargetUrl = $pair.TargetUrl; Stage = $stage; Kind = 'Failure'
                Path = ''; Detail = $_.Exception.Message; Evidence = $siteFolder
            })
        } finally {
            Export-SPPermissionCsv $summary.ToArray() @('SourceUrl','TargetUrl','Stage','Status','Error','EvidenceFolder') (Join-Path $folder 'Summary.csv')
            Export-SPPermissionCsv $issues.ToArray() @('TargetUrl','Stage','Kind','Path','Detail','Evidence') (Join-Path $folder 'Stragglers.csv')
        }
    }
    Export-SPPermissionCsv $summary.ToArray() @('SourceUrl','TargetUrl','Stage','Status','Error','EvidenceFolder') (Join-Path $folder 'Summary.csv')
    Export-SPPermissionCsv $issues.ToArray() @('TargetUrl','Stage','Kind','Path','Detail','Evidence') (Join-Path $folder 'Stragglers.csv')
    Write-Host "Checklist finished. Failures: $(@($summary | Where-Object Status -Like 'Failed*').Count). Evidence: $folder"
    if (@($summary | Where-Object Status -Like 'Failed*').Count) {
        throw "Some checklist sites failed; see Summary.csv and Stragglers.csv in '$folder'. Successful sites were not rolled back."
    }
}

if ($MyInvocation.InvocationName -ne '.') { Invoke-SPVisitorsReadChecklist }
