#requires -Version 7.4.6
[CmdletBinding()]
param()
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'Get-SystemEditorRemediationPlan.ps1')

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
function Import-Module($Name, $MinimumVersion) {
    Assert-Equal $Name 'PnP.PowerShell'
}
function Connect-PnPOnline($Url, $ClientId, [switch]$Interactive, [switch]$ReturnConnection) {
    Assert-Equal $Interactive.IsPresent $true
    Assert-Equal $ReturnConnection.IsPresent $true
    if ($Url -match 'archive|outside') { throw "Forbidden connection: $Url" }
    return $Url
}
function Get-PnPWeb($Connection, $Includes) {
    [pscustomobject]@{ Url = $Connection }
}
function Get-PnPSubWeb([switch]$Recurse, $Connection, $Includes) {
    Assert-Equal $Recurse.IsPresent $true
    if ($script:failure -eq 'discovery' -and $Connection -eq 'https://whallc.sharepoint.com') {
        throw 'Simulated discovery failure'
    }
    if ($Connection -eq 'https://whallc.sharepoint.com') {
        [pscustomobject]@{ Url = "$Connection/PortalChild" }
    }
    if ($Connection -eq 'https://whallc.sharepoint.com/sites/IT') {
        [pscustomobject]@{ Url = "$Connection/IT-Internal" }
        [pscustomobject]@{ Url = "$Connection/IT-Internal/Deep" }
    }
    if ($Connection -eq 'https://whallc.sharepoint.com/sites/IT/IT-Internal') {
        [pscustomobject]@{ Url = "$Connection/Deep" }
    }
    if ($script:failure -eq 'scope') {
        [pscustomobject]@{ Url = 'https://outside.sharepoint.com/sites/Unexpected' }
    }
}
function Get-PnPUser($Connection) {
    [pscustomobject]@{ Id = 7; LoginName = 'i:0#.f|membership|!bposey@whallc.com'; Email = '!bposey@whallc.com' }
    [pscustomobject]@{ Id = 8; LoginName = 'i:0#.f|membership|other@example.test'; Email = 'other@example.test'; Title = '!bposey@whallc.com' }
    if ($script:failure -ne 'system') {
        [pscustomobject]@{ Id = 99; LoginName = 'SHAREPOINT\system'; Email = '' }
    }
}
function Get-PnPList($Identity, $Connection, $Includes) {
    if ($Identity) {
        return [pscustomobject]@{ ItemCount = $(if ($script:failure -eq 'count') { 8 } else { 7 }) }
    }
    $script:webScans.Add($Connection)
    [pscustomobject]@{ Id = 'docs'; Title = 'Documents'; Hidden = $true; ItemCount = 7; BaseType = 'DocumentLibrary' }
    [pscustomobject]@{ Id = 'tasks'; Title = 'Tasks'; Hidden = $false; ItemCount = 7; BaseType = 'GenericList' }
    [pscustomobject]@{ Id = 'system'; Title = 'No editor metadata'; Hidden = $true; ItemCount = 0; BaseType = 'GenericList' }
}
function Get-PnPField($List, $Connection) {
    $names = if ($List -eq 'system') { @('ID', 'Title') } else {
        @('ID', 'Modified', 'Editor', 'FileRef', 'FSObjType', 'UniqueId', 'Created', 'Author', '_UIVersionString')
    }
    foreach ($name in $names) { [pscustomobject]@{ InternalName = $name } }
}
function Get-PnPListItem($List, $Query, $PageSize, $Connection) {
    Assert-Equal $PageSize 500
    if ($Query -notmatch "Scope='RecursiveAll'" -or $Query -notmatch "Paged='TRUE'") { throw 'Non-recursive/unpaged query' }
    if ($script:failure -eq 'list' -and $List -eq 'docs') { throw 'Simulated access denied' }
    foreach ($id in 1..7) {
        $date = switch ($id) {
            2 { '2024-04-01T07:00:00Z' }
            4 { '2024-04-01T07:00:01Z' }
            7 { '2024-04-01T06:59:59Z' }
            default { '2023-03-31T12:34:56Z' }
        }
        $item = @{
            Id = $id
            Modified = [datetime]::Parse($date, [cultureinfo]::InvariantCulture, [Globalization.DateTimeStyles]::RoundtripKind)
            Editor = @{ LookupId = $(if ($id -eq 3) { 8 } else { 7 }) }
            FileRef = "$Connection/$List/item$id"; FSObjType = $(if ($id -eq 5) { 1 } else { 0 })
            UniqueId = [guid]::NewGuid(); Created = [datetime]::new(2020, 1, 1, 0, 0, 0, [DateTimeKind]::Utc)
            Author = @{ LookupId = 7 }; _UIVersionString = '4.0'
        }
        $item
    }
}

$temp = Join-Path ([IO.Path]::GetTempPath()) ('wha-system-preview-tests-' + [guid]::NewGuid().ToString('N'))
[void][IO.Directory]::CreateDirectory($temp)
try {
    $csv = Join-Path $temp 'Checklist.csv'
    @(
        [pscustomobject]@{ 'Old Site URL' = 'https://whallc.sharepoint.com/sites/archive-old'; 'New Site URL' = 'https://whallc.sharepoint.com' }
        [pscustomobject]@{ 'New Site URL' = 'https://whallc.sharepoint.com/sites/DocumentsandForms' }
        [pscustomobject]@{ 'New Site URL' = 'https://whallc.sharepoint.com/sites/IT' }
        [pscustomobject]@{ 'New Site URL' = 'https://whallc.sharepoint.com/sites/IT/IT-Internal' }
        [pscustomobject]@{ 'New Site URL' = 'https://whallc.sharepoint.com/sites/IT/' }
    ) | Export-Csv -LiteralPath $csv -NoTypeInformation
    Assert-Equal (@(Get-SystemEditorChecklistRoots $csv).Count) 4
    Assert-Throws { ConvertTo-SystemEditorWebUrl 'https://whallc.sharepoint.com/sites/archive-2026/CS' } 'Archive'
    Assert-Throws { ConvertTo-SystemEditorWebUrl 'https://outside.sharepoint.com/sites/IT' } 'Invalid'
    Assert-Throws { ConvertTo-SystemEditorWebUrl 'https://whallc.sharepoint.com/sites/IT?x=1' } 'Invalid'
    Assert-Equal (Test-SystemEditorSourceUser ([pscustomobject]@{ LoginName = 'OTHER'; Email = ''; Title = '!bposey@whallc.com' })) $false
    Assert-Equal (Test-SystemEditorSourceUser ([pscustomobject]@{ LoginName = '!BPOSEY@whallc.com'; Email = '' })) $true
    Assert-Equal (ConvertTo-SystemEditorUtc ([datetime]::new(2024, 4, 1, 7, 0, 0))).Kind ([DateTimeKind]::Utc)
    Write-Host 'PASS checklist-only roots, archive/foreign URL guards, identity matching and UTC handling'

    $script:failure = ''
    $script:webScans = [Collections.Generic.List[string]]::new()
    $result = Invoke-SystemEditorRemediationPlan -ChecklistPath $csv -OutputFolder $temp -ClientId 'test'
    Assert-Equal $result.Status 'PreviewComplete'
    Assert-Equal $result.WebsScanned 6
    Assert-Equal $result.ItemsScanned 84
    Assert-Equal $result.Matches 48
    Assert-Equal $script:webScans.Count 6
    Assert-Equal (@($script:webScans | Sort-Object -Unique).Count) 6
    $plan = @(Import-Csv -LiteralPath (Join-Path $result.ReportFolder 'Plan.csv'))
    Assert-Equal $plan.Count 48
    Assert-Equal (($plan.ItemId | Sort-Object -Unique) -join ',') '1,5,6,7'
    Assert-Equal (($plan.Kind | Sort-Object -Unique) -join ',') 'File,Folder,ListItem'
    Assert-Equal (@($plan | Where-Object { $_.SystemEditorId -ne '99' -or $_.Status -ne 'Preview' }).Count) 0
    $coverage = @(Import-Csv -LiteralPath (Join-Path $result.ReportFolder 'Coverage.csv'))
    Assert-Equal (@($coverage | Where-Object Status -eq 'No Modified/Editor fields').Count) 6
    Write-Host 'PASS recursive subsites/overlaps, hidden lists, folders/files/items, strict Pacific cutoff and preview reports'

    $script:failure = 'count'
    $script:webScans = [Collections.Generic.List[string]]::new()
    $countWarnings = @()
    $result = Invoke-SystemEditorRemediationPlan -ChecklistPath $csv -OutputFolder $temp -ClientId 'test' `
        -WarningVariable countWarnings
    Assert-Equal $result.Status 'PreviewComplete'
    Assert-Equal $result.Errors 0
    Assert-Equal $result.Matches 48
    Assert-Equal $countWarnings.Count 0
    $coverage = @(Import-Csv -LiteralPath (Join-Path $result.ReportFolder 'Coverage.csv'))
    Assert-Equal (@($coverage | Where-Object { $_.Note -like 'Reported counts:*' }).Count) 12
    Assert-Equal (@(Import-Csv -LiteralPath (Join-Path $result.ReportFolder 'Errors.csv')).Count) 0
    Write-Host 'PASS count differences are report-only notes, with no warning, lost matches or failed preview'

    foreach ($failure in 'list', 'discovery', 'system', 'scope') {
        $script:failure = $failure
        $script:webScans = [Collections.Generic.List[string]]::new()
        $output = Join-Path $temp $failure
        $script:operatorWarnings = @()
        Assert-Throws {
            Invoke-SystemEditorRemediationPlan -ChecklistPath $csv -OutputFolder $output -ClientId 'test' `
                -WarningVariable script:operatorWarnings -WarningAction SilentlyContinue
        } 'Preview incomplete'
        if ($script:operatorWarnings.Count -eq 0 -or $script:operatorWarnings.Count -gt 4) {
            throw 'Expected at most one warning per checklist root with issues.'
        }
        foreach ($warning in $script:operatorWarnings) {
            if ("$warning" -notmatch '^\S.*: \d+ matches; \d+ issues\. See Errors\.csv\.$') {
                throw "Unexpected operator warning: $warning"
            }
        }
        $run = @(Get-ChildItem -LiteralPath $output -Directory)[0].FullName
        $saved = Get-Content -LiteralPath (Join-Path $run 'Summary.json') -Raw | ConvertFrom-Json
        Assert-Equal $saved.Status 'Incomplete'
        if ($saved.Errors -le 0) { throw 'Failure was not recorded' }
        if (-not @(Import-Csv -LiteralPath (Join-Path $run 'Errors.csv')).Count) { throw 'Missing error evidence' }
    }
    Write-Host 'PASS failures stay explicit in reports with concise warnings grouped by checklist root'

    $tokens = $null
    $parseErrors = $null
    $ast = [Management.Automation.Language.Parser]::ParseFile(
        (Join-Path $PSScriptRoot 'Get-SystemEditorRemediationPlan.ps1'), [ref]$tokens, [ref]$parseErrors)
    Assert-Equal $parseErrors.Count 0
    $writes = @($ast.FindAll({
        param($node)
        $node -is [Management.Automation.Language.CommandAst] -and
        $node.GetCommandName() -match '^(Set|New|Add|Remove|Invoke|Ensure|Submit)-PnP'
    }, $true))
    Assert-Equal $writes.Count 0
    if ($ast.ParamBlock.Parameters.Name.VariablePath.UserPath -contains 'Apply') { throw 'Unexpected Apply parameter' }
    Write-Host 'PASS no SharePoint write commands or Apply parameter'
} finally {
    Remove-Item -LiteralPath $temp -Recurse -Force
}
