#requires -Version 7.4.6
<#
.SYNOPSIS
    Corrective Visitors Read pass for CS, Accounting and PFS only.
.DESCRIPTION
    Reuses the checklist's Match -> checked DryRun -> verified Apply pipeline,
    including matching descendants and site-specific exclusions. Read-only unless
    -Apply is supplied. Temporary catch-up for the three previously skipped sites;
    use Invoke-SPVisitorsReadChecklist.ps1 for future all-site runs.
#>
[CmdletBinding()]
param(
    [string]$ChecklistFile = (Join-Path $PSScriptRoot '..\ReferenceFiles\NewSiteChecklist.csv'),
    [string]$OutputFolder = (Join-Path $PSScriptRoot '..\Reports\VisitorsSecondPassRuns'),
    [string]$ClientId = '2e20d82f-56ba-418e-b465-9718493f0df8',
    [string]$Tenant,
    [switch]$SkipPersistLoginSetup,
    [switch]$Apply
)
. (Join-Path $PSScriptRoot 'SPOutputLimits.Common.ps1')
$ErrorActionPreference = 'Stop'
$options = @{
    ChecklistFile = $ChecklistFile
    OutputFolder = $OutputFolder
    TargetSites = @(
        'https://whallc.sharepoint.com/sites/ClinicalandPatientServices'
        'https://whallc.sharepoint.com/sites/Accounting'
        'https://whallc.sharepoint.com/sites/PatientFinancialServices'
    )
    ClientId = $ClientId
    Tenant = $Tenant
    SkipPersistLoginSetup = $SkipPersistLoginSetup
    Apply = $Apply
}
& (Join-Path $PSScriptRoot 'Invoke-SPVisitorsReadChecklist.ps1') @options
