#requires -Version 7.4.6
<#
.SYNOPSIS
    Read-only sequential scouts for explicit site pairs in NewSiteChecklist.csv.
.DESCRIPTION
    Reuses Get-CSPermissionScout.ps1, not a second inventory implementation.
    By default skips CS, which has its own scout. -IncludeCS includes it.
    -TargetSites selects exact checklist target URLs and overrides the default CS exclusion.
    Completes each pair before starting the next. Read failures are recorded in the
    batch summary and remaining pairs continue; the batch throws at the end if any failed.
    A count discrepancy is Needs review, not success or proof of missing content.
    No tenant enumeration, automatic subsite recursion or permission writes.
    Listed subsites are separate pairs. Unlisted subsites are outside this scout's scope.
    Keep Get-CSPermissionScout.ps1 and SPPermission.Common.ps1 beside this script.
.EXAMPLE
    .\Invoke-SPPermissionScouts.ps1 -ChecklistFile .\NewSiteChecklist.csv
.EXAMPLE
    .\Invoke-SPPermissionScouts.ps1 -ChecklistFile .\NewSiteChecklist.csv -TargetSites 'https://whallc.sharepoint.com/sites/Accounting'
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$ChecklistFile,
    [string]$OutputFolder = (Join-Path $PSScriptRoot 'Site-PermissionScouts'),
    [string[]]$TargetSites,
    [switch]$IncludeCS,
    [string]$ClientId = '2e20d82f-56ba-418e-b465-9718493f0df8',
    [switch]$SkipPersistLoginSetup
)
. (Join-Path $PSScriptRoot 'SPPermission.Common.ps1')

function Invoke-SPPermissionScouts {
    $ErrorActionPreference = 'Stop'
    $scoutPath = Join-Path $PSScriptRoot 'Get-CSPermissionScout.ps1'
    if (-not (Test-Path -LiteralPath $scoutPath)) { throw 'Missing companion Get-CSPermissionScout.ps1.' }
    $rows = @(Import-Csv -LiteralPath $ChecklistFile -ErrorAction Stop)
    if (-not $rows.Count) { throw 'Checklist is empty.' }
    $sources = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $targets = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $pairs = @(
        foreach ($row in $rows) {
            if (-not $row.'Old Site URL' -or -not $row.'New Site URL') { throw 'Every checklist row needs Old Site URL and New Site URL.' }
            $sourceUrl = $row.'Old Site URL'.Trim().TrimEnd('/')
            $targetUrl = $row.'New Site URL'.Trim().TrimEnd('/')
            Assert-SPPermissionUrls $sourceUrl $targetUrl
            if (-not $sources.Add($sourceUrl) -or -not $targets.Add($targetUrl)) { throw 'Checklist pairs must be one-to-one.' }
            [pscustomobject]@{ SourceUrl = $sourceUrl; TargetUrl = $targetUrl }
        }
    )
    if ($TargetSites.Count) {
        $selected = @($TargetSites | ForEach-Object { $_.Trim().TrimEnd('/') })
        foreach ($url in $selected) {
            if (-not $targets.Contains($url)) { throw "Requested target is not in the checklist: $url" }
        }
        $pairs = @($pairs | Where-Object { $_.TargetUrl -iin $selected })
    } elseif (-not $IncludeCS) {
        $pairs = @($pairs | Where-Object TargetUrl -INE 'https://whallc.sharepoint.com/sites/ClinicalandPatientServices')
    }
    if (-not $pairs.Count) { throw 'No site pairs selected.' }
    $batchFolder = Join-Path $OutputFolder ('Scouts_' + (Get-Date -Format 'yyyyMMdd-HHmmss-fff'))
    [void][IO.Directory]::CreateDirectory($batchFolder)
    $summaryPath = Join-Path $batchFolder 'ScoutSummary.csv'
    $summary = [Collections.Generic.List[object]]::new()
    $index = 0
    foreach ($pair in $pairs) {
        $index++
        $leaf = (([uri]$pair.TargetUrl).AbsolutePath.Trim('/') -split '/')[-1] -replace '[^a-zA-Z0-9_-]', '_'
        if (-not $leaf) { $leaf = 'Root' }
        $pairFolder = Join-Path $batchFolder ('{0:D2}-{1}' -f $index, $leaf)
        $record = [pscustomobject]@{
            SourceUrl = $pair.SourceUrl; TargetUrl = $pair.TargetUrl; Status = 'Failed'
            ListsNeedingReview = $null; Detail = ''; ReportFolder = $pairFolder
        }
        Write-Host "[$index/$($pairs.Count)] Scout $($pair.SourceUrl) -> $($pair.TargetUrl)"
        try {
            & $scoutPath -SourceUrl $pair.SourceUrl -TargetUrl $pair.TargetUrl -OutputFolder $pairFolder `
                -ClientId $ClientId -SkipPersistLoginSetup:$SkipPersistLoginSetup
            $listReports = @(Get-ChildItem -LiteralPath $pairFolder -Recurse -Filter 'Lists.csv' -ErrorAction Stop)
            $errorReports = @(Get-ChildItem -LiteralPath $pairFolder -Recurse -Filter 'Errors.csv' -ErrorAction Stop)
            if ($listReports.Count -ne 1 -or $errorReports.Count -ne 1) { throw 'Scout did not produce the expected reports.' }
            if (@(Import-Csv -LiteralPath $errorReports[0].FullName).Count) { throw 'Scout recorded read errors.' }
            $record.ListsNeedingReview = @(Import-Csv -LiteralPath $listReports[0].FullName |
                Where-Object Status -ne 'Counts and IDs agree').Count
            $record.Status = if ($record.ListsNeedingReview) { 'Needs review' } else { 'Counts and IDs agree' }
            $record.ReportFolder = $listReports[0].DirectoryName
        } catch {
            $record.Detail = $_.Exception.Message
            Write-Warning "Scout failed for '$($pair.TargetUrl)'; reports retained. Continuing to the next pair: $($record.Detail)"
        } finally {
            $summary.Add($record)
            Export-SPPermissionCsv $summary.ToArray() @('SourceUrl','TargetUrl','Status','ListsNeedingReview','Detail','ReportFolder') $summaryPath
        }
    }
    Write-Host "Scout summary: $summaryPath"
    $failures = @($summary | Where-Object Status -eq 'Failed').Count
    $reviews = @($summary | Where-Object Status -eq 'Needs review').Count
    Write-Host "Pairs: $($summary.Count); failed: $failures; needing review: $reviews. No permission changes made."
    if ($failures) { throw "$failures site scout(s) failed. Remaining pairs were attempted; inspect ScoutSummary.csv." }
}

if ($MyInvocation.InvocationName -ne '.') { Invoke-SPPermissionScouts }
