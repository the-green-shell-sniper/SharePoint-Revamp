#requires -Version 7.4.6
[CmdletBinding()]
param()
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'Invoke-SystemEditorRemediation.ps1')
Microsoft.PowerShell.Core\Import-Module PnP.PowerShell -MinimumVersion 3.3.0

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

$script:nativeWriter = ${function:Set-SystemEditorItemValues}
$script:writes = 0
$script:executes = 0
$nativeItem = @{ Editor = $null; Modified = $null }
$nativeItem | Add-Member ScriptMethod UpdateOverwriteVersion { $script:writes++ }
$nativeContext = [pscustomobject]@{}
$nativeContext | Add-Member ScriptMethod ExecuteQuery { $script:executes++ }
$nativeItem | Add-Member NoteProperty Context $nativeContext
$date = [datetime]::new(2023, 1, 1, 0, 0, 0, [DateTimeKind]::Utc)
& $script:nativeWriter $nativeItem 99 $date
Assert-Equal $nativeItem.Editor.LookupId 99
Assert-Equal $nativeItem.Modified $date
Assert-Equal $script:writes 1
Assert-Equal $script:executes 1
Write-Host 'PASS native writer assigns only Editor/retained Modified and executes overwrite once'

function Import-Module($Name, $MinimumVersion) {
    Assert-Equal $Name 'PnP.PowerShell'
}
function Connect-PnPOnline($Url, $ClientId, [switch]$Interactive, [switch]$ReturnConnection) {
    $script:connections++
    if ($Url -notin 'https://whallc.sharepoint.com', 'https://whallc.sharepoint.com/sites/Docs',
        'https://whallc.sharepoint.com/sites/Docs/Child') { throw "Unexpected connection $Url" }
    return $Url
}
function Get-PnPWeb($Connection, $Includes) { [pscustomobject]@{ Url = $Connection } }
function Get-PnPSubWeb([switch]$Recurse, $Connection, $Includes) {
    Assert-Equal $Recurse.IsPresent $true
    if ($Connection -eq 'https://whallc.sharepoint.com/sites/Docs') {
        [pscustomobject]@{ Url = "$Connection/Child" }
    }
}
function Get-PnPUser($Identity, $Connection) {
    if ($Identity -eq 'SHAREPOINT\system') {
        if ($script:scenario -eq 'missing-system') { return $null }
        return [pscustomobject]@{ Id = 99; LoginName = 'SHAREPOINT\system'; Email = '' }
    }
    if ($Identity -eq 99) {
        return [pscustomobject]@{ Id = 99; LoginName = 'SHAREPOINT\system'; Email = '' }
    }
    [pscustomobject]@{ Id = 7; LoginName = 'i:0#.f|membership|!bposey@whallc.com'; Email = '!bposey@whallc.com' }
}
function Get-PnPField($List, $Connection) {
    foreach ($name in $script:live.Keys) {
        [pscustomobject]@{
            InternalName = $name; ReadOnlyField = ($name -in 'UniqueId', 'owshiddenversion', '_UIVersionString', 'FileRef')
            TypeAsString = 'Text'
        }
    }
}
function Get-PnPListItem($List, $Id, $Connection) {
    $script:reads++
    Assert-Equal $Id 1
    Assert-Equal ([string]$List) $script:listId
    if ($script:scenario -eq 'read-error') { throw 'Simulated access denied' }
    if ($script:scenario -eq 'race' -and $script:reads -eq 2) { $script:live.Title = 'Concurrent edit' }
    $values = $script:live.Clone()
    $values | Add-Member NoteProperty FieldValues $values
    $values | Add-Member NoteProperty ObjectVersion '1'
    return $values
}
function Set-SystemEditorItemValues($Item, $SystemId, $Modified) {
    Assert-Equal $SystemId 99
    Assert-Equal $Modified.Kind ([DateTimeKind]::Utc)
    Assert-Equal $Modified $script:live.Modified
    $saved = @(Import-Csv -LiteralPath (Join-Path $script:lastRun 'Before.csv'))
    if (-not @($saved | Where-Object Field -eq 'Editor').Count) { throw 'No before-write backup' }
    $journal = @(Import-Csv -LiteralPath (Join-Path $script:lastRun 'Journal.csv'))
    Assert-Equal $journal[-1].Status 'Write started'
    $script:writes++
    if ($script:scenario -eq 'write-error') { throw 'Simulated ambiguous response' }
    $script:live.Editor = [pscustomobject]@{ LookupId = 99; LookupValue = 'System Account' }
    $script:live.owshiddenversion++
    switch ($script:scenario) {
        'bad-title' { $script:live.Title = 'Unexpected' }
        'bad-date' { $script:live.Modified = $script:live.Modified.AddDays(1) }
        'bad-created' { $script:live.Created = $script:live.Created.AddDays(1) }
        'bad-author' { $script:live.Author = [pscustomobject]@{ LookupId = 11 } }
        'bad-version' { $script:live._UIVersionString = '2.0' }
        'bad-editor' { $script:live.Editor = [pscustomobject]@{ LookupId = 7 } }
    }
}

$temp = Join-Path ([IO.Path]::GetTempPath()) ('wha-editor-apply-' + [guid]::NewGuid().ToString('N'))
[void][IO.Directory]::CreateDirectory($temp)
$script:listId = '126393dc-e80b-4e4d-81c2-42049ed75cf2'
$script:uid = '9572052f-9179-4e8a-91ee-5dabdabb3600'
$csv = Join-Path $temp 'Checklist.csv'
@(
    [pscustomobject]@{ 'New Site URL' = 'https://whallc.sharepoint.com' }
    [pscustomobject]@{ 'New Site URL' = 'https://whallc.sharepoint.com/sites/Docs' }
) | Export-Csv -LiteralPath $csv -NoTypeInformation
$preview = Join-Path $temp 'preview'
[void][IO.Directory]::CreateDirectory($preview)
$output = Join-Path $temp 'apply'
$options = @{ PreviewFolder = $preview; ChecklistPath = $csv; OutputFolder = $output; ClientId = 'test' }

function Reset-Fixture([string]$Scenario = '') {
    $script:scenario = $Scenario
    $script:writes = 0
    $script:reads = 0
    $script:connections = 0
    $script:live = @{
        UniqueId = [guid]$script:uid; Modified = [datetime]::new(2023, 1, 1, 0, 0, 0, [DateTimeKind]::Utc)
        Created = [datetime]::new(2020, 1, 1, 0, 0, 0, [DateTimeKind]::Utc)
        Author = [pscustomobject]@{ LookupId = 3 }
        Editor = [pscustomobject]@{ LookupId = 7; LookupValue = 'Brittany !. Posey' }
        _UIVersionString = '1.0'; FileRef = '/sites/Docs/Child/Docs/item'
        Title = 'Preserve me'; CustomValues = @('A', 'B'); owshiddenversion = 1
    }
    $script:row = [pscustomobject]@{
        WebUrl = 'https://whallc.sharepoint.com/sites/Docs/Child'; ListId = $script:listId
        ItemId = '1'; UniqueId = $script:uid; ModifiedUtc = $script:live.Modified.ToString('o')
        CreatedUtc = $script:live.Created.ToString('o'); AuthorId = '3'; Version = '1.0'; Path = $script:live.FileRef
        PreviousEditorLogin = 'i:0#.f|membership|!bposey@whallc.com'; PreviousEditorEmail = ''
        Status = 'Preview'
    }
    $script:row | Export-Csv -LiteralPath (Join-Path $preview 'Plan.csv') -NoTypeInformation
    [pscustomobject]@{ WebUrl = $script:row.WebUrl; ListId = $script:listId; Matches = '1'; Status = 'Scanned' } |
        Export-Csv -LiteralPath (Join-Path $preview 'Coverage.csv') -NoTypeInformation
    [pscustomobject]@{ WebUrl = $script:row.WebUrl; Status = 'Enumerated (see list coverage)' } |
        Export-Csv -LiteralPath (Join-Path $preview 'Webs.csv') -NoTypeInformation
    'WebUrl,ListId,Stage,Error' | Set-Content -LiteralPath (Join-Path $preview 'Errors.csv')
    @{
        Mode = 'PreviewOnly'; Status = 'PreviewComplete'; Errors = 0; SourceEditor = '!bposey@whallc.com'
        CutoffUtc = '2024-04-01T07:00:00Z'; Matches = 1; WebsScanned = 1
        ChecklistRoots = @('https://whallc.sharepoint.com', 'https://whallc.sharepoint.com/sites/Docs')
    } | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $preview 'Summary.json')
}

# Capture the generated run directory without changing production logging behavior.
$script:realSnapshot = ${function:Get-SystemEditorFieldSnapshot}
function Get-SystemEditorFieldSnapshot($Item) {
    $script:lastRun = @(Get-ChildItem -LiteralPath $output -Directory | Sort-Object CreationTime -Descending)[0].FullName
    & $script:realSnapshot $Item
}
try {
    Reset-Fixture
    Assert-Throws { Invoke-SystemEditorApply @options -Apply } 'MaintenanceWindowConfirmed'
    Assert-Equal $script:connections 0
    $result = Invoke-SystemEditorApply @options
    Assert-Equal $result.Eligible 1
    Assert-Equal $script:writes 0
    Write-Host 'PASS maintenance gate and default read-only recheck'

    Reset-Fixture
    $result = Invoke-SystemEditorApply @options -Apply -MaintenanceWindowConfirmed
    Assert-Equal $result.Verified 1
    Assert-Equal $result.Status 'Completed'
    Assert-Equal $script:writes 1
    Assert-Equal $script:live.Title 'Preserve me'
    Assert-Equal $script:live.owshiddenversion 2
    Write-Host 'PASS backup/journal precede write; stored metadata verified; server counter tolerated'

    Reset-Fixture
    $script:live.Modified = [datetime]::new(2024, 4, 1, 6, 59, 59, [DateTimeKind]::Utc)
    $script:row.ModifiedUtc = $script:live.Modified.ToString('o')
    $editor = Get-PnPUser -Identity 7
    Assert-Equal (Get-SystemEditorCandidateStatus (Get-PnPListItem -List $script:listId -Id 1) $script:row $editor 99) 'Eligible'
    $script:live.Modified = $script:live.Modified.AddSeconds(1)
    Assert-Equal (Get-SystemEditorCandidateStatus (Get-PnPListItem -List $script:listId -Id 1) $script:row $editor 99) 'Skipped: no longer qualifies'
    Write-Host 'PASS Pacific cutoff uses an explicit UTC boundary, including the final second before cutoff'

    foreach ($scenario in 'changed', 'ineligible', 'identity', 'checkout', 'race', 'already') {
        Reset-Fixture $scenario
        switch ($scenario) {
            'changed' { $script:live._UIVersionString = '2.0' }
            'ineligible' { $script:live.Modified = [datetime]::new(2024, 4, 1, 7, 0, 0, [DateTimeKind]::Utc) }
            'identity' { $script:live.UniqueId = [guid]::NewGuid() }
            'checkout' { $script:live.CheckoutUser = [pscustomobject]@{ LookupId = 7 } }
            'already' { $script:live.Editor = [pscustomobject]@{ LookupId = 99 } }
        }
        $result = Invoke-SystemEditorApply @options -Apply -MaintenanceWindowConfirmed
        Assert-Equal $script:writes 0
        if ($scenario -eq 'already') { Assert-Equal $result.AlreadySystem 1 } else { Assert-Equal $result.Skipped 1 }
    }
    Write-Host 'PASS changed/ineligible/replaced/checked-out/concurrent/already-correct items are not written'

    foreach ($scenario in 'bad-title', 'bad-date', 'bad-created', 'bad-author', 'bad-version', 'bad-editor',
        'read-error', 'write-error', 'missing-system') {
        Reset-Fixture $scenario
        Assert-Throws { Invoke-SystemEditorApply @options -Apply -MaintenanceWindowConfirmed } '^Stopped'
        if ($script:writes -gt 1) { throw 'Write retried' }
        $run = @(Get-ChildItem -LiteralPath $output -Directory | Sort-Object CreationTime -Descending)[0].FullName
        $summary = Get-Content -LiteralPath (Join-Path $run 'Summary.json') -Raw | ConvertFrom-Json
        Assert-Equal $summary.Status 'Stopped'
        Assert-Equal $summary.Verified 0
    }
    Write-Host 'PASS read/write/account/metadata verification failures stop with persistent evidence and no retry'

    Reset-Fixture
    $script:row.WebUrl = 'https://whallc.sharepoint.com/sites/Other'
    $script:row | Export-Csv -LiteralPath (Join-Path $preview 'Plan.csv') -NoTypeInformation
    [pscustomobject]@{ WebUrl = $script:row.WebUrl; ListId = $script:listId; Matches = '1'; Status = 'Scanned' } |
        Export-Csv -LiteralPath (Join-Path $preview 'Coverage.csv') -NoTypeInformation
    [pscustomobject]@{ WebUrl = $script:row.WebUrl; Status = 'Enumerated (see list coverage)' } |
        Export-Csv -LiteralPath (Join-Path $preview 'Webs.csv') -NoTypeInformation
    Assert-Throws { Invoke-SystemEditorApply @options -Apply -MaintenanceWindowConfirmed } 'no longer in approved structure'
    Assert-Equal $script:writes 0
    Write-Host 'PASS tenant-home URL cannot authorize an unrelated site collection'

    Reset-Fixture
    $script:row.WebUrl = 'https://whallc.sharepoint.com/sites/Other'
    $script:row | Export-Csv -LiteralPath (Join-Path $preview 'Plan.csv') -NoTypeInformation
    Assert-Throws { Read-SystemEditorApprovedPlan $preview $csv } 'invalid, duplicate or unreviewed'
    Reset-Fixture
    @($script:row, $script:row) | Export-Csv -LiteralPath (Join-Path $preview 'Plan.csv') -NoTypeInformation
    Assert-Throws { Read-SystemEditorApprovedPlan $preview $csv } 'row count'
    Reset-Fixture
    $script:row.ModifiedUtc = '2026-01-01T00:00:00Z'
    $script:row | Export-Csv -LiteralPath (Join-Path $preview 'Plan.csv') -NoTypeInformation
    Assert-Throws { Read-SystemEditorApprovedPlan $preview $csv } 'invalid, duplicate or unreviewed'
    Write-Host 'PASS altered scope, candidate counts and dates rejected before connecting'

    Reset-Fixture
    $savedLimit = $script:SPOutputPartLimit
    try {
        $script:SPOutputPartLimit = 2KB
        $result = Invoke-SystemEditorApply @options -Apply -MaintenanceWindowConfirmed
        Assert-Equal $result.Verified 1
        $parts = @(Get-SPCsvSiblingParts (Join-Path $result.ReportFolder 'Before.csv'))
        if (-not $parts.Count) { throw 'Backup CSV did not split in the bounded-output test' }
        foreach ($file in Get-ChildItem -LiteralPath $result.ReportFolder -File) {
            if ($file.Length -ge 2KB) { throw 'Output limit exceeded' }
        }
        $backupRows = @(Import-Csv -LiteralPath (Join-Path $result.ReportFolder 'Before.csv'))
        Assert-Equal $backupRows.Count $script:live.Count
    } finally { $script:SPOutputPartLimit = $savedLimit }
    Write-Host 'PASS bounded backup CSV parts round-trip without missing fields or oversize files'
} finally {
    Remove-Item -LiteralPath $temp -Recurse -Force
}
