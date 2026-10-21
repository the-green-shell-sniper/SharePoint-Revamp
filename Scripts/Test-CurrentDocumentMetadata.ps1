#requires -Version 7.4.6
[CmdletBinding()]
param()
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'Sync-CurrentDocumentMetadata.ps1')

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
    $script:source = @{
        Modified = [datetime]'2026-09-10T12:34:56Z'
        Editor = @{ LookupId = 7 }
    }
    $script:target = @{
        Id = 42; Modified = [datetime]'2026-10-19T10:00:00Z'
        Editor = @{ LookupId = 11 }; Author = @{ LookupId = 3 }
        Created = [datetime]'2025-01-01T00:00:00Z'
        UniqueId = [guid]::NewGuid(); _UIVersionString = '4.0'
    }
    $script:writes = 0
    $script:ensures = 0
    $script:corruptField = ''
    $script:failWrite = $false
}
function Get-PnPFile($Url, [switch]$AsListItem, $Connection) {
    if ($Connection -eq 'source') { return $script:source.Clone() }
    return $script:target.Clone()
}
function Get-PnPUser($Identity, $Connection) {
    $login = if ($Connection -eq 'source' -or $Identity -eq 99) { 'i:0#.f|membership|original@example.test' }
        else { 'i:0#.f|membership|current@example.test' }
    [pscustomobject]@{ LoginName = $login }
}
function Get-PnPProperty($ClientObject, $Property, $Connection) {
    Assert-Equal $Connection 'target'
    Assert-Equal $Property 'ParentList'
    [pscustomobject]@{ Id = 'target-list' }
}
function New-PnPUser($LoginName, $Connection) {
    Assert-Equal $Connection 'target'
    Assert-Equal $LoginName 'i:0#.f|membership|original@example.test'
    $script:ensures++
    [pscustomobject]@{ Id = 99 }
}
function Set-PnPListItem($List, $Identity, $Values, $UpdateType, $Connection) {
    Assert-Equal $Connection 'target'
    Assert-Equal $List 'target-list'
    Assert-Equal $Identity 42
    Assert-Equal $UpdateType 'UpdateOverwriteVersion'
    Assert-Equal (($Values.Keys | Sort-Object) -join ',') 'Editor,Modified'
    Assert-Equal $Values.Editor '99'
    Assert-Equal $Values.Modified.Kind ([DateTimeKind]::Utc)
    if ($script:failWrite) { throw 'Simulated write failure' }
    $script:writes++
    $script:target.Modified = $Values.Modified
    $script:target.Editor = @{ LookupId = 99 }
    switch ($script:corruptField) {
        'Modified' { $script:target.Modified = [datetime]'2026-10-20T00:00:00Z' }
        'Editor' { $script:target.Editor = @{ LookupId = 11 } }
        'Author' { $script:target.Author = @{ LookupId = 99 } }
        '' { }
        default { $script:target[$script:corruptField] = 'unexpected' }
    }
}
$options = @{
    SourcePath = '/archive/file.pdf'; TargetPath = '/target/file.pdf'
    SourceConnection = 'source'; TargetConnection = 'target'
}

Reset-TestState
$result = Copy-CurrentDocumentMetadata @options
Assert-Equal $result.Status 'Preview'
Assert-Equal $script:writes 0
Assert-Equal $script:ensures 0
Write-Host 'PASS preview does not write or ensure users'

Reset-TestState
$result = Copy-CurrentDocumentMetadata @options -Apply
Assert-Equal $result.Status 'Updated and verified'
Assert-Equal $script:writes 1
Assert-Equal $script:ensures 1
Write-Host 'PASS apply changes exactly two fields using target user ID and no new version'
$result = Copy-CurrentDocumentMetadata @options -Apply
Assert-Equal $result.Status 'Already matches'
Assert-Equal $script:writes 1
Write-Host 'PASS matching metadata skips writes'

foreach ($field in 'Modified', 'Editor') {
    Reset-TestState
    $script:source.Remove($field)
    Assert-Throws { Copy-CurrentDocumentMetadata @options -Apply } 'missing Modified'
    Assert-Equal $script:writes 0
}
Write-Host 'PASS missing source metadata fails before writing'

foreach ($field in 'Created', 'UniqueId', '_UIVersionString') {
    Reset-TestState
    $script:corruptField = $field
    Assert-Throws { Copy-CurrentDocumentMetadata @options -Apply } "Verification failed: $field changed"
}
Write-Host 'PASS verification detects changed creation time, identity or version'

foreach ($field in 'Modified', 'Editor', 'Author') {
    Reset-TestState
    $script:corruptField = $field
    Assert-Throws { Copy-CurrentDocumentMetadata @options -Apply } 'Verification failed'
}
Write-Host 'PASS verification detects incorrect Modified/Editor and changed Author'

Reset-TestState
$script:failWrite = $true
Assert-Throws { Copy-CurrentDocumentMetadata @options -Apply } 'Simulated write failure'
Write-Host 'PASS write failures propagate without a success result'
