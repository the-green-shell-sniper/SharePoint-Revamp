#requires -Version 7.4.6
<#
.SYNOPSIS
    Match current Modified/Modified By for the two Documents and Forms PDFs.
.DESCRIPTION
    Preview only unless -Apply is supplied. Reads the archived files and updates
    only Modified and Editor on the targets using UpdateOverwriteVersion.
    No file upload, new version, or rewrite of older versions is performed.
    The current version's metadata is overwritten; Created and Author are preserved.
    SharePoint automation (including Power Automate flows) may be triggered.
.EXAMPLE
    .\Sync-CurrentDocumentMetadata.ps1
.EXAMPLE
    .\Sync-CurrentDocumentMetadata.ps1 -Apply
#>
[CmdletBinding()]
param(
    [string]$ClientId = '2e20d82f-56ba-418e-b465-9718493f0df8',
    [switch]$Apply
)

function Copy-CurrentDocumentMetadata {
    [CmdletBinding()]
    param(
        [string]$SourcePath,
        [string]$TargetPath,
        $SourceConnection,
        $TargetConnection,
        [switch]$Apply
    )
    $ErrorActionPreference = 'Stop'
    $source = Get-PnPFile -Url $SourcePath -AsListItem -Connection $SourceConnection
    $target = Get-PnPFile -Url $TargetPath -AsListItem -Connection $TargetConnection
    if (-not $source['Modified'] -or -not $source['Editor']) {
        throw "Source is missing Modified or Modified By: $SourcePath"
    }
    $editor = Get-PnPUser -Identity $source['Editor'].LookupId -Connection $SourceConnection
    if ([string]::IsNullOrWhiteSpace($editor.LoginName)) {
        throw "Source Modified By has no resolvable login: $SourcePath"
    }
    $modified = ([datetime]$source['Modified']).ToUniversalTime()
    $currentEditor = Get-PnPUser -Identity $target['Editor'].LookupId -Connection $TargetConnection
    $result = [pscustomobject]@{
        Target = $TargetPath
        PreviousModifiedUtc = ([datetime]$target['Modified']).ToUniversalTime().ToString('o')
        PreviousModifiedBy = $currentEditor.LoginName
        SourceModifiedUtc = $modified.ToString('o')
        SourceModifiedBy = $editor.LoginName
        Status = 'Preview'
    }
    if (([datetime]$target['Modified']).ToUniversalTime() -eq $modified -and
        $currentEditor.LoginName -ieq $editor.LoginName) {
        $result.Status = 'Already matches'
        return $result
    }
    if (-not $Apply) { return $result }

    $before = @{}
    foreach ($field in 'Created', 'UniqueId', '_UIVersionString') {
        $before[$field] = $target[$field]
    }
    $before.AuthorId = $target['Author'].LookupId
    $list = Get-PnPProperty -ClientObject $target -Property ParentList -Connection $TargetConnection
    # User lookup IDs are site-collection-specific; resolve the source login in the target.
    $targetEditor = New-PnPUser -LoginName $editor.LoginName -Connection $TargetConnection
    Set-PnPListItem -List $list.Id -Identity $target.Id -Values @{
        Modified = $modified
        Editor = [string]$targetEditor.Id
    } -UpdateType UpdateOverwriteVersion -Connection $TargetConnection | Out-Null

    $after = Get-PnPFile -Url $TargetPath -AsListItem -Connection $TargetConnection
    if (([datetime]$after['Modified']).ToUniversalTime() -ne $modified -or
        $after['Editor'].LookupId -ne $targetEditor.Id) {
        throw "Verification failed for Modified/Modified By on $TargetPath. Inspect live metadata before retrying."
    }
    foreach ($field in 'Created', 'UniqueId', '_UIVersionString') {
        if ($after[$field] -ne $before[$field]) {
            throw "Verification failed: $field changed on $TargetPath. Inspect live metadata before retrying."
        }
    }
    if ($after['Author'].LookupId -ne $before.AuthorId) {
        throw "Verification failed: Author changed on $TargetPath. Inspect live metadata before retrying."
    }
    $result.Status = 'Updated and verified'
    return $result
}

function Invoke-CurrentDocumentMetadataSync {
    $ErrorActionPreference = 'Stop'
    Import-Module PnP.PowerShell -MinimumVersion 3.3.0 -ErrorAction Stop
    $sourceConnection = Connect-PnPOnline -Url 'https://whallc.sharepoint.com/sites/archive-2026-10-18T115436Z' `
        -ClientId $ClientId -Interactive -ReturnConnection -ErrorAction Stop
    $targetConnection = Connect-PnPOnline -Url 'https://whallc.sharepoint.com/sites/DocumentsandForms' `
        -ClientId $ClientId -Interactive -ReturnConnection -ErrorAction Stop

    foreach ($name in 'WHA Clinician Org Chart 2026.11.09.pdf', 'WHA Courier Schedule (SEP 2026).pdf') {
        Copy-CurrentDocumentMetadata `
            -SourcePath "/sites/archive-2026-10-18T115436Z/About Us/$name" `
            -TargetPath "/sites/DocumentsandForms/Shared Documents/$name" `
            -SourceConnection $sourceConnection -TargetConnection $targetConnection -Apply:$Apply |
            Format-List
    }
}

if ($MyInvocation.InvocationName -ne '.') {
    Invoke-CurrentDocumentMetadataSync
}
