#requires -Version 7.4.6
<#
.SYNOPSIS
    Accounting-only additive Visitors Read workflow, root then matched subsites.
.DESCRIPTION
    Reuses the sequential PFS workflow with explicit Accounting URLs/group/paths.
    No full mirror, removals, user resolution or permission-level creation.
    Zero source Team Site Members Read scopes means no EEEU membership changes.
#>
[CmdletBinding()]
param(
    [ValidateSet('Scout', 'Match', 'DryRun', 'Apply', 'Report')][string]$Mode = 'DryRun',
    [string]$OutputFolder = (Join-Path $PSScriptRoot '..\Reports\Accounting'),
    [string]$ReviewRoot = (Join-Path $PSScriptRoot '..\Reports\AccountingReview'),
    [string]$ClientId = '2e20d82f-56ba-418e-b465-9718493f0df8',
    [string]$Tenant,
    [switch]$SkipPersistLoginSetup
)
$ErrorActionPreference = 'Stop'
& (Join-Path $PSScriptRoot 'Invoke-PFSVisitorsRead.ps1') `
    -SourceSiteUrl 'https://whallc.sharepoint.com/sites/archive-2026-10-18T115436Z/Accounting' `
    -TargetSiteUrl 'https://whallc.sharepoint.com/sites/Accounting' `
    -VisitorsTitle 'Accounting Visitors' -ExcludedPaths @('Access Requests') -SiteLabel 'Accounting' `
    -Mode $Mode -OutputFolder $OutputFolder -ReviewRoot $ReviewRoot `
    -ClientId $ClientId -Tenant $Tenant -SkipPersistLoginSetup:$SkipPersistLoginSetup
