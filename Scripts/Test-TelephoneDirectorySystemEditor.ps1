#requires -Version 7.4.6
[CmdletBinding()]
param()
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'Set-TelephoneDirectorySystemEditor.ps1')

function Assert-Equal($Actual, $Expected) {
    if ($Actual -cne $Expected) { throw "Expected [$Expected], received [$Actual]." }
}
function Assert-Throws([scriptblock]$Action, [string]$Pattern) {
    try { & $Action | Out-Null } catch {
        if ($_.Exception.Message -notmatch $Pattern) { throw }
        return
    }
    throw "Expected an exception matching: $Pattern"
}
function Reset-TestState {
    $script:item = @{
        Id = 42; UniqueId = [guid]'9572052f-9179-4e8a-91ee-5dabdabb3600'
        Modified = [datetime]'2023-03-31T12:34:56Z'
        Created = [datetime]'2020-01-01T00:00:00Z'
        Editor = @{ LookupId = 11 }; Author = @{ LookupId = 3 }
        _UIVersionString = '4.0'
    }
    $script:system = [pscustomobject]@{ Id = 99; LoginName = 'SHAREPOINT\system'; Title = 'System Account' }
    $script:writes = 0
    $script:corruptField = ''
    $script:failWrite = $false
}
function Get-PnPFile($Url, [switch]$AsListItem, $Connection) {
    Assert-Equal $Url '/sites/DocumentsandForms/Shared Documents/WHA Telephone Directory.docx'
    Assert-Equal $Connection 'target'
    Assert-Equal $AsListItem.IsPresent $true
    return $script:item.Clone()
}
function Get-PnPUser($Identity, $Connection) {
    Assert-Equal $Connection 'target'
    if ($Identity -eq 'SHAREPOINT\system') { return $script:system }
    [pscustomobject]@{ LoginName = 'i:0#.f|membership|!bposey@whallc.com' }
}
function Get-PnPProperty($ClientObject, $Property, $Connection) {
    Assert-Equal $Property 'ParentList'
    Assert-Equal $Connection 'target'
    [pscustomobject]@{ Id = 'target-list' }
}
function Set-PnPListItem($List, $Identity, $Values, $UpdateType, $Connection) {
    Assert-Equal $Connection 'target'
    Assert-Equal $List 'target-list'
    Assert-Equal $Identity 42
    Assert-Equal $UpdateType 'UpdateOverwriteVersion'
    Assert-Equal (($Values.Keys | Sort-Object) -join ',') 'Editor,Modified'
    Assert-Equal $Values.Editor '99'
    Assert-Equal $Values.Modified $script:item.Modified.ToUniversalTime()
    Assert-Equal $Values.Modified.Kind ([DateTimeKind]::Utc)
    if ($script:failWrite) { throw 'Simulated write rejection' }
    $script:writes++
    $script:item.Editor = @{ LookupId = 99 }
    switch ($script:corruptField) {
        'Modified' { $script:item.Modified = [datetime]'2026-10-20T00:00:00Z' }
        'Editor' { $script:item.Editor = @{ LookupId = 11 } }
        'Author' { $script:item.Author = @{ LookupId = 99 } }
        '' { }
        default { $script:item[$script:corruptField] = 'unexpected' }
    }
}

Reset-TestState
$result = Set-TelephoneDirectorySystemEditor -Connection 'target'
Assert-Equal $result.Status 'Preview'
Assert-Equal $script:writes 0
Write-Host 'PASS preview does not write'

Reset-TestState
$result = Set-TelephoneDirectorySystemEditor -Connection 'target' -Apply
Assert-Equal $result.Status 'Updated and verified'
Assert-Equal $script:writes 1
$result = Set-TelephoneDirectorySystemEditor -Connection 'target' -Apply
Assert-Equal $result.Status 'Already System Account'
Assert-Equal $script:writes 1
Write-Host 'PASS native account write preserves metadata and skips already-correct item'

Reset-TestState
$script:item.UniqueId = [guid]::NewGuid()
Assert-Throws { Set-TelephoneDirectorySystemEditor -Connection 'target' -Apply } 'identity differs'
Assert-Equal $script:writes 0
Write-Host 'PASS wrong document identity fails before writing'

foreach ($login in '', 'i:0#.f|membership|system@example.test') {
    Reset-TestState
    $script:system.LoginName = $login
    Assert-Throws { Set-TelephoneDirectorySystemEditor -Connection 'target' -Apply } 'Native.*unavailable'
    Assert-Equal $script:writes 0
}
Reset-TestState
$script:system = $null
Assert-Throws { Set-TelephoneDirectorySystemEditor -Connection 'target' -Apply } 'Native.*unavailable'
Assert-Equal $script:writes 0
Write-Host 'PASS missing native account or same-title substitute rejected'

foreach ($field in 'Created', 'UniqueId', '_UIVersionString', 'Author', 'Modified', 'Editor') {
    Reset-TestState
    $script:corruptField = $field
    Assert-Throws { Set-TelephoneDirectorySystemEditor -Connection 'target' -Apply } 'Verification failed'
}
Write-Host 'PASS verification detects metadata or version changes'

Reset-TestState
$script:failWrite = $true
Assert-Throws { Set-TelephoneDirectorySystemEditor -Connection 'target' -Apply } 'Simulated write rejection'
Assert-Equal $script:writes 0
Write-Host 'PASS SharePoint write rejection propagates without retry or success'
