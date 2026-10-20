#requires -Version 7.4.6
<#
.SYNOPSIS
    PFS-only Scout, Match, DryRun, Apply and recursive reporting.
.DESCRIPTION
    Sequential root-first processing of exact-path descendant web pairs. No full
    mirror, removals, Review-level creation or source user resolution. A web with
    zero source Team Site Members Read scopes makes no permission or EEEU changes.
    Target-only webs/content are reported, not matched by guesswork.
#>
[CmdletBinding()]
param(
    [ValidateSet('Scout', 'Match', 'DryRun', 'Apply', 'Report')][string]$Mode = 'DryRun',
    [string]$OutputFolder = (Join-Path $PSScriptRoot '..\Reports\PFS'),
    [string]$ReviewRoot = (Join-Path $PSScriptRoot '..\Reports\PFSReview'),
    [string]$ClientId = '2e20d82f-56ba-418e-b465-9718493f0df8',
    [string]$Tenant,
    [switch]$SkipPersistLoginSetup
)

function Invoke-PFSVisitorsRead {
    $ErrorActionPreference = 'Stop'
    $sourceRoot = 'https://whallc.sharepoint.com/sites/archive-2026-10-18T115436Z/PFS'
    $targetRoot = 'https://whallc.sharepoint.com/sites/PatientFinancialServices'
    $root = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputFolder)
    $reviews = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ReviewRoot)
    . (Join-Path $PSScriptRoot 'SPPermission.Common.ps1')
    Import-Module PnP.PowerShell -MinimumVersion 3.3.0 -ErrorAction Stop
    $folder = Join-Path $root ('p-' + (Get-Date -Format 'yyMMddHHmmssfff'))
    [void][IO.Directory]::CreateDirectory($folder)
    $auth = @{ ClientId = $ClientId; Tenant = $Tenant; SkipPersistLoginSetup = $SkipPersistLoginSetup }
    if ($Mode -eq 'Report') {
        & (Join-Path $PSScriptRoot 'Export-SPPermissionSiteReport.ps1') -TargetUrl $targetRoot -OutputFolder $folder @auth
        return
    }
    $queue = [Collections.Generic.List[object]]::new()
    $queue.Add([pscustomobject]@{ SourceUrl = $sourceRoot; TargetUrl = $targetRoot })
    $summary = [Collections.Generic.List[object]]::new()
    $webRows = [Collections.Generic.List[object]]::new()
    $seen = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    try {
        for ($index = 0; $index -lt $queue.Count; $index++) {
            $pair = $queue[$index]
            if (-not $seen.Add($pair.TargetUrl)) { throw "Repeated PFS web: $($pair.TargetUrl)" }
            $key = [Convert]::ToHexString([Security.Cryptography.SHA256]::HashData(
                [Text.Encoding]::UTF8.GetBytes($pair.TargetUrl.ToLowerInvariant()))).Substring(0, 12)
            $webFolder = Join-Path $folder $key
            $reviewFolder = Join-Path $reviews $key
            [void][IO.Directory]::CreateDirectory($webFolder)
            [void][IO.Directory]::CreateDirectory($reviewFolder)
            Write-Host "[$Mode PFS] $($pair.SourceUrl) -> $($pair.TargetUrl)"
            $sourceConnection = Connect-SPPermissionSite -Url $pair.SourceUrl @auth
            $targetConnection = Connect-SPPermissionSite -Url $pair.TargetUrl @auth
            $sourceChildren = @(Get-SPPermissionCollection $sourceConnection ($pair.SourceUrl + '/_api/web/webs?$select=Url,Title'))
            $targetChildren = @(Get-SPPermissionCollection $targetConnection ($pair.TargetUrl + '/_api/web/webs?$select=Url,Title'))
            $offset = 1
            foreach ($child in $sourceChildren | Sort-Object Url) {
                if (-not $child.Url -or -not (Test-SPPermissionPathWithin $child.Url $pair.SourceUrl) -or
                    $child.Url.TrimEnd('/') -ieq $pair.SourceUrl) { throw 'Invalid PFS source child web URL.' }
                $relative = $child.Url.TrimEnd('/').Substring($pair.SourceUrl.Length)
                $counterpart = @($targetChildren | Where-Object { $_.Url.TrimEnd('/') -ieq ($pair.TargetUrl + $relative) })
                if ($counterpart.Count -gt 1) { throw 'Duplicate target child URL.' }
                $webRows.Add([pscustomobject]@{
                    SourceUrl = $child.Url; TargetUrl = ($pair.TargetUrl + $relative)
                    Status = $(if ($counterpart.Count) { 'Exact path - queued' } else { 'Unmatched source web - skipped' })
                })
                if ($counterpart.Count) {
                    $queue.Insert($index + $offset, [pscustomobject]@{
                        SourceUrl = $child.Url.TrimEnd('/'); TargetUrl = $counterpart[0].Url.TrimEnd('/')
                    })
                    $offset++
                }
            }
            foreach ($child in $targetChildren) {
                if (-not $child.Url -or -not (Test-SPPermissionPathWithin $child.Url $pair.TargetUrl) -or
                    $child.Url.TrimEnd('/') -ieq $pair.TargetUrl) { throw 'Invalid PFS target child web URL.' }
                if ($child.Url.TrimEnd('/') -inotin $queue.TargetUrl) {
                    $webRows.Add([pscustomobject]@{ SourceUrl = ''; TargetUrl = $child.Url; Status = 'Target-only web - report only' })
                }
            }
            $options = @{ SourceUrl = $pair.SourceUrl; TargetUrl = $pair.TargetUrl; OutputFolder = $webFolder }
            $reviewFile = Join-Path $reviewFolder 'MatchReview.csv'
            switch ($Mode) {
                'Scout' {
                    $scoutAuth = $auth.Clone()
                    [void]$scoutAuth.Remove('Tenant')
                    & (Join-Path $PSScriptRoot 'Get-CSPermissionScout.ps1') @options @scoutAuth
                }
                'Match' {
                    & (Join-Path $PSScriptRoot 'New-SPPermissionMatchMap.ps1') @options @auth -ExcludePaths @('Access Requests', '2021 EM Coding Changes')
                    $generated = @(Get-ChildItem -LiteralPath $webFolder -Filter 'PermMatch_Review_*.csv')
                    if ($generated.Count -ne 1) { throw 'Expected one PFS match review.' }
                    if (-not (Test-Path -LiteralPath $reviewFile)) {
                        Copy-Item -LiteralPath $generated[0].FullName -Destination $reviewFile
                    } else { Write-Warning "Existing approvals retained: $reviewFile. Review fresh candidates in $webFolder." }
                    & {
                        param($ScriptFolder, $Path)
                        . (Join-Path $ScriptFolder 'Invoke-SPPermissionChecklist.ps1') -ChecklistFile 'unused'
                        Set-SPChecklistAutoApprovals $Path
                    } $PSScriptRoot $reviewFile
                }
                default {
                    if (Test-Path -LiteralPath $reviewFile) { $options.MatchFile = $reviewFile }
                    & (Join-Path $PSScriptRoot 'Add-CSVisitorsRead.ps1') @options @auth `
                        -ExpectedVisitorsTitle 'Patient Financial Services Visitors' `
                        -ExcludePaths @('Access Requests', '2021 EM Coding Changes') `
                        -NoMembershipWithoutSourceRead -SkipFinalReport -Apply:($Mode -eq 'Apply')
                }
            }
            $summary.Add([pscustomobject]@{
                SourceUrl = $pair.SourceUrl; TargetUrl = $pair.TargetUrl; Mode = $Mode
                Status = 'Completed - inspect coverage'; EvidenceFolder = $webFolder; Error = ''
            })
        }
    } catch {
        $summary.Add([pscustomobject]@{
            SourceUrl = $pair.SourceUrl; TargetUrl = $pair.TargetUrl; Mode = $Mode
            Status = 'Failed'; EvidenceFolder = $webFolder; Error = $_.Exception.Message
        })
        throw
    } finally {
        Export-SPPermissionCsv $summary.ToArray() @('SourceUrl','TargetUrl','Mode','Status','EvidenceFolder','Error') (Join-Path $folder 'Summary.csv')
        Export-SPPermissionCsv $webRows.ToArray() @('SourceUrl','TargetUrl','Status') (Join-Path $folder 'Webs.csv')
        Write-Host "PFS evidence: $folder"
    }
    if ($Mode -eq 'Apply') {
        & (Join-Path $PSScriptRoot 'Export-SPPermissionSiteReport.ps1') -TargetUrl $targetRoot `
            -OutputFolder (Join-Path $folder 'Report') @auth
    }
}

if ($MyInvocation.InvocationName -ne '.') { Invoke-PFSVisitorsRead }
