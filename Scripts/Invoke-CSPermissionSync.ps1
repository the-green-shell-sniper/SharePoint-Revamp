#requires -Version 7.4.6
<#
.SYNOPSIS
    CS-only permission workflow. Never selects the other checklist sites.
.DESCRIPTION
    Match -> review CSVs -> DryRun -> review plans -> Apply -> Verify.
    Uses cached PnP interactive/persistent login; no tenant argument is required.
    Includes exact/approved matching descendant webs and all their enumerated content.
    Unmatched content is skipped; inherited target-only scopes can be frozen to
    preserve access. Fuzzy proposals never authorize changes without approval.
    Keep the other four permission scripts beside this launcher.
.EXAMPLE
    .\Invoke-CSPermissionSync.ps1 -Mode Match
.EXAMPLE
    .\Invoke-CSPermissionSync.ps1 -Mode DryRun
.EXAMPLE
    .\Invoke-CSPermissionSync.ps1 -Mode Apply
.EXAMPLE
    .\Invoke-CSPermissionSync.ps1 -Mode Verify
#>
[CmdletBinding()]
param(
    [ValidateSet('Match', 'DryRun', 'Apply', 'Verify')][string]$Mode = 'DryRun',
    [string]$OutputFolder = (Join-Path $PSScriptRoot 'CS-PermissionRuns'),
    [string]$ReviewRoot = (Join-Path $PSScriptRoot 'CS-PermissionReviews'),
    [string]$ClientId = '2e20d82f-56ba-418e-b465-9718493f0df8',
    [string]$Tenant,
    [switch]$NoContent,
    [switch]$SkipPersistLoginSetup
)

function Invoke-CSPermissionSync {
    $ErrorActionPreference = 'Stop'
    $OutputFolder = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputFolder)
    $ReviewRoot = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ReviewRoot)
    foreach ($name in @('SPPermission.Common.ps1', 'New-SPPermissionMatchMap.ps1', 'Sync-SPPermissionsFromSource.ps1', 'Invoke-SPPermissionChecklist.ps1')) {
        if (-not (Test-Path -LiteralPath (Join-Path $PSScriptRoot $name))) { throw "Missing companion script: $name" }
    }
    [void][IO.Directory]::CreateDirectory($OutputFolder)
    $runFolder = Join-Path $OutputFolder ("CS_${Mode}_" + (Get-Date -Format 'yyyyMMdd-HHmmss-fff'))
    [void][IO.Directory]::CreateDirectory($runFolder)
    $checklistPath = Join-Path $runFolder 'CS-SitePair.csv'
    [pscustomobject]@{
        'Old Site URL' = 'https://whallc.sharepoint.com/sites/archive-2026-10-18T115436Z/CS'
        'New Site URL' = 'https://whallc.sharepoint.com/sites/ClinicalandPatientServices'
    } | Export-Csv -LiteralPath $checklistPath -NoTypeInformation -Encoding utf8
    $options = @{
        ChecklistFile = $checklistPath; Mode = $(if ($Mode -eq 'Verify') { 'DryRun' } else { $Mode })
        OutputFolder = $runFolder; ReviewRoot = $ReviewRoot; ClientId = $ClientId; Tenant = $Tenant
        NoContent = $NoContent; SkipPersistLoginSetup = $SkipPersistLoginSetup
    }
    & (Join-Path $PSScriptRoot 'Invoke-SPPermissionChecklist.ps1') @options
    Write-Host "CS reports: $runFolder"
    Write-Host "CS review files: $ReviewRoot"
    if ($Mode -eq 'Match') {
        Write-Host 'Review MatchReview.csv (content) and WebReview.csv (subsites) under CS-PermissionReviews.'
        Write-Host 'All fuzzy rows start Approved=No. Approve only confirmed identities. Run Match again after approving a renamed subsite to inventory its content.'
    }
    if ($Mode -eq 'Verify') {
        $plans = @(Get-ChildItem -LiteralPath $runFolder -Recurse -Filter 'PermSync_Plan_DRYRUN_*.json')
        $summaries = @(Get-ChildItem -LiteralPath $runFolder -Recurse -Filter 'ChecklistSummary.csv')
        if (-not $plans.Count -or $summaries.Count -ne 1) { throw 'CS verification incomplete: expected live dry-run plans and a checklist summary.' }
        $summary = @(Import-Csv -LiteralPath $summaries[0].FullName)
        if (@($summary | Where-Object { $_.Status -notlike 'DryRun completed*' }).Count) {
            throw 'CS verification incomplete: one or more web pairs were skipped or failed. Inspect ChecklistSummary.csv.'
        }
        $remaining = 0; $blocked = 0; $unavailable = 0
        foreach ($path in $plans) {
            $plan = Get-Content -LiteralPath $path.FullName -Raw | ConvertFrom-Json -Depth 100 -AsHashtable
            $remaining += @($plan.Steps).Count
            $blocked += @($plan.Notes | Where-Object Result -eq 'Blocked').Count
        }
        foreach ($path in @(Get-ChildItem -LiteralPath $runFolder -Recurse -Filter 'PermSync_UnavailablePrincipals_DRYRUN_*.csv')) {
            $unavailable += @(Import-Csv -LiteralPath $path.FullName).Count
        }
        if ($remaining -or $blocked -or $unavailable) {
            throw "CS verification needs review: $remaining remaining actions, $blocked blocked scopes, $unavailable unavailable principals. No writes were made by Verify."
        }
        Write-Host 'CS matched scopes verified: zero planned actions, zero blocked scopes and zero unavailable principals. Unmatched/unapproved content remains skipped; test access with your testing account.'
    }
}

if ($MyInvocation.InvocationName -ne '.') { Invoke-CSPermissionSync }
