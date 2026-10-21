#requires -Version 7.4.6
<#
.SYNOPSIS
    Set the Telephone Directory's current Modified By to native System Account.
.DESCRIPTION
    Preview by default; pass -Apply to write. This single-document test does not
    apply the proposed bulk date/editor filters or enumerate checklist sites.
    Resolves SHAREPOINT\system in the target site collection, never a similarly
    named user. Stops if the native account is unavailable or SharePoint rejects it.
    Preserves Modified, Created, Author and the current version number using
    UpdateOverwriteVersion. Older versions and file contents are not rewritten.
    SharePoint automation, including Power Automate flows, may be triggered.
.EXAMPLE
    .\Set-TelephoneDirectorySystemEditor.ps1
.EXAMPLE
    .\Set-TelephoneDirectorySystemEditor.ps1 -Apply
#>
[CmdletBinding()]
param(
    [string]$ClientId = '2e20d82f-56ba-418e-b465-9718493f0df8',
    [switch]$Apply
)

function Set-TelephoneDirectorySystemEditor {
    [CmdletBinding()]
    param($Connection, [switch]$Apply)

    $ErrorActionPreference = 'Stop'
    $path = '/sites/DocumentsandForms/Shared Documents/WHA Telephone Directory.docx'
    $expectedId = [guid]'9572052f-9179-4e8a-91ee-5dabdabb3600'
    $item = Get-PnPFile -Url $path -AsListItem -Connection $Connection
    if ([guid]$item['UniqueId'] -ne $expectedId) {
        throw 'Telephone Directory identity differs from the supplied link. No changes made.'
    }
    $system = Get-PnPUser -Identity 'SHAREPOINT\system' -Connection $Connection
    if (-not $system -or $system.Id -le 0 -or $system.LoginName -ine 'SHAREPOINT\system') {
        throw 'Native SHAREPOINT\system account is unavailable. No substitute will be used; no changes made.'
    }
    if (-not $item['Modified'] -or -not $item['Created'] -or
        -not $item['Author'] -or -not $item['Editor'] -or -not $item['_UIVersionString']) {
        throw 'Required current metadata is missing. No changes made.'
    }
    $previousEditor = Get-PnPUser -Identity $item['Editor'].LookupId -Connection $Connection
    $modified = ([datetime]$item['Modified']).ToUniversalTime()
    $result = [pscustomobject]@{
        Target = $path
        ModifiedUtc = $modified.ToString('o')
        PreviousModifiedBy = $previousEditor.LoginName
        NewModifiedBy = $system.LoginName
        NewModifiedByTitle = $system.Title
        Status = 'Preview'
    }
    if ($item['Editor'].LookupId -eq $system.Id) {
        $result.Status = 'Already System Account'
        return $result
    }
    if (-not $Apply) { return $result }

    $before = @{}
    foreach ($field in 'Created', 'UniqueId', '_UIVersionString') {
        $before[$field] = $item[$field]
    }
    $before.AuthorId = $item['Author'].LookupId
    $list = Get-PnPProperty -ClientObject $item -Property ParentList -Connection $Connection
    # Explicitly retain Modified: a normal item update would stamp the current time.
    Set-PnPListItem -List $list.Id -Identity $item.Id -Values @{
        Editor = [string]$system.Id
        Modified = $modified
    } -UpdateType UpdateOverwriteVersion -Connection $Connection | Out-Null

    $after = Get-PnPFile -Url $path -AsListItem -Connection $Connection
    if ($after['Editor'].LookupId -ne $system.Id -or
        ([datetime]$after['Modified']).ToUniversalTime() -ne $modified) {
        throw 'Verification failed for Editor/Modified. Inspect live metadata before retrying.'
    }
    foreach ($field in 'Created', 'UniqueId', '_UIVersionString') {
        if ($after[$field] -ne $before[$field]) {
            throw "Verification failed: $field changed. Inspect live metadata before retrying."
        }
    }
    if ($after['Author'].LookupId -ne $before.AuthorId) {
        throw 'Verification failed: Author changed. Inspect live metadata before retrying.'
    }
    $result.Status = 'Updated and verified'
    return $result
}

if ($MyInvocation.InvocationName -ne '.') {
    $ErrorActionPreference = 'Stop'
    Import-Module PnP.PowerShell -MinimumVersion 3.3.0 -ErrorAction Stop
    $connection = Connect-PnPOnline -Url 'https://whallc.sharepoint.com/sites/DocumentsandForms' `
        -ClientId $ClientId -Interactive -ReturnConnection -ErrorAction Stop
    Set-TelephoneDirectorySystemEditor -Connection $connection -Apply:$Apply | Format-List
}
