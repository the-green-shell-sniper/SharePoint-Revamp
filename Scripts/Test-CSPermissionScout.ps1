#requires -Version 7.4.6
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'Get-CSPermissionScout.ps1')
function Assert-Equal($Actual, $Expected) {
    if ($Actual -cne $Expected) { throw "Expected [$Expected], received [$Actual]." }
}
$temp = Join-Path ([IO.Path]::GetTempPath()) ('cs-scout-tests-' + [guid]::NewGuid().ToString('N'))
try {
    function Import-Module {}
    function Connect-SPPermissionSite { param($Url) $Url }
    function Get-PnPWeb { param($Connection, $Includes, $ErrorAction) [pscustomobject]@{ Url = $Connection } }
    function Get-PnPList {
        param($Connection, $Includes, $Identity, $ErrorAction)
        [pscustomobject]@{
            Id = 'pages'; Title = 'Site Pages'; ItemCount = 5; Hidden = $false
            BaseTemplate = 119; HasUniqueRoleAssignments = $false
            RootFolder = [pscustomobject]@{ ServerRelativeUrl = '/SitePages' }
        }
    }
    function Get-PnPField {
        param($List, $Connection, $ErrorAction)
        foreach ($name in @('ID', 'FileRef', 'FileLeafRef', 'FSObjType')) {
            [pscustomobject]@{ InternalName = $name; TypeAsString = 'Text'; Hidden = $false; ReadOnlyField = $false }
        }
    }
    function Get-PnPListItem {
        param($List, $Query, $PageSize, $Connection, $ErrorAction)
        Assert-Equal $List 'pages'
        Assert-Equal $PageSize 500
        Assert-Equal ($Query -match "Scope='RecursiveAll'") $true
        Assert-Equal ($Query -match "Name='Title'") $false
        foreach ($id in 1..4) {
            [pscustomobject]@{ Id = $id; FieldValues = @{ FileRef = "/SitePages/$id.aspx"; FileLeafRef = "$id.aspx"; FSObjType = 0 } }
        }
    }
    function Get-PnPProperty {
        param($ClientObject, $Property, $Connection, $ErrorAction)
        if ($failInheritance -and $ClientObject.Id -eq 2) { throw 'MOCK inheritance denied' }
        return $false
    }
    function Invoke-PnPSPRestMethod {
        param($Connection, $Url, $Method, [switch]$Raw, $ErrorAction)
        Assert-Equal $Method 'Get'
        '{"value":[{"Id":1,"ID":1},{"Id":2,"ID":2},{"Id":3,"ID":3},{"Id":4,"ID":4}]}'
    }
    $OutputFolder = $temp
    $failInheritance = $false
    Invoke-CSPermissionScout
    $report = Get-ChildItem -LiteralPath $temp -Recurse -Filter 'Lists.csv'
    $rows = @(Import-Csv -LiteralPath $report.FullName)
    Assert-Equal $rows.Count 2
    Assert-Equal $rows[0].PnPCount '4'
    Assert-Equal $rows[0].RestCount '4'
    Assert-Equal $rows[0].Status 'Enumerators agree; reported count differs - review'
    Assert-Equal (@(Get-ChildItem -LiteralPath $temp -Recurse -Filter 'Items.csv' | ForEach-Object { Import-Csv -LiteralPath $_.FullName }).Count) 8
    Write-Host 'PASS recursive PnP and REST scouting records count discrepancies without pretending they are resolved.'
    $OutputFolder = Join-Path $temp 'failed'
    $failInheritance = $true
    $caught = $false
    try { Invoke-CSPermissionScout } catch {
        if ($_.Exception.Message -notmatch 'Scout had read failures') { throw }
        $caught = $true
    }
    Assert-Equal $caught $true
    $errorsPath = Get-ChildItem -LiteralPath $OutputFolder -Recurse -Filter 'Errors.csv'
    Assert-Equal (@(Import-Csv -LiteralPath $errorsPath.FullName).Count) 2
    Write-Host 'PASS read failures persist in Errors.csv and cause explicit failure.'
    $checklist = Join-Path $temp 'sites.csv'
    @(
        [pscustomobject]@{ 'Old Site URL' = 'https://source.test/a'; 'New Site URL' = 'https://target.test/a' }
        [pscustomobject]@{ 'Old Site URL' = 'https://source.test/CS'; 'New Site URL' = 'https://whallc.sharepoint.com/sites/ClinicalandPatientServices' }
        [pscustomobject]@{ 'Old Site URL' = 'https://source.test/b'; 'New Site URL' = 'https://target.test/b' }
    ) | Export-Csv -LiteralPath $checklist -NoTypeInformation
    $state = [pscustomobject]@{ Calls = [Collections.Generic.List[string]]::new(); FailUrl = '' }
    function Connect-PnPOnline {
        param($Url, $ClientId, [switch]$Interactive, [switch]$PersistLogin, [switch]$ReturnConnection, $ErrorAction)
        Assert-Equal $Interactive.IsPresent $true
        Assert-Equal $PersistLogin.IsPresent $true
        Assert-Equal $ReturnConnection.IsPresent $true
        $state.Calls.Add($Url)
        $Url
    }
    function Get-PnPWeb {
        param($Connection, $Includes, $ErrorAction)
        if ($Connection -eq $state.FailUrl) { throw 'MOCK inaccessible web' }
        [pscustomobject]@{ Url = $Connection }
    }
    $failInheritance = $false
    $runner = Join-Path $PSScriptRoot 'Invoke-SPPermissionScouts.ps1'
    $batchOutput = Join-Path $temp 'batch'
    & $runner -ChecklistFile $checklist -OutputFolder $batchOutput
    Assert-Equal ($state.Calls -join ',') 'https://source.test/a,https://target.test/a,https://source.test/b,https://target.test/b'
    $batchSummary = Get-ChildItem -LiteralPath $batchOutput -Recurse -Filter 'ScoutSummary.csv'
    $batchRows = @(Import-Csv -LiteralPath $batchSummary.FullName)
    Assert-Equal $batchRows.Count 2
    Assert-Equal $batchRows[0].Status 'Needs review'
    Assert-Equal $batchRows[0].ListsNeedingReview '2'
    Write-Host 'PASS sequential checklist runner reuses the actual scout, excludes CS by default, and flags count differences.'
    $state.Calls.Clear()
    & $runner -ChecklistFile $checklist -OutputFolder (Join-Path $temp 'selected') -TargetSites 'https://target.test/b/'
    Assert-Equal ($state.Calls -join ',') 'https://source.test/b,https://target.test/b'
    Write-Host 'PASS single-pair selection scouts exactly the selected URLs.'
    $state.Calls.Clear()
    & $runner -ChecklistFile $checklist -OutputFolder (Join-Path $temp 'include-cs') -IncludeCS
    Assert-Equal $state.Calls.Count 6
    Assert-Equal $state.Calls.Contains('https://whallc.sharepoint.com/sites/ClinicalandPatientServices') $true
    Write-Host 'PASS IncludeCS opts the CS pair back in.'
    $state.Calls.Clear()
    $state.FailUrl = 'https://source.test/a'
    $failedOutput = Join-Path $temp 'batch-failure'
    $caught = $false
    try { & $runner -ChecklistFile $checklist -OutputFolder $failedOutput } catch {
        if ($_.Exception.Message -notmatch 'site scout\(s\) failed') { throw }
        $caught = $true
    }
    Assert-Equal $caught $true
    Assert-Equal ($state.Calls -join ',') 'https://source.test/a,https://source.test/b,https://target.test/b'
    $batchSummary = Get-ChildItem -LiteralPath $failedOutput -Recurse -Filter 'ScoutSummary.csv'
    $batchRows = @(Import-Csv -LiteralPath $batchSummary.FullName)
    Assert-Equal $batchRows[0].Status 'Failed'
    Assert-Equal $batchRows[1].Status 'Needs review'
    Write-Host 'PASS failed pair is reported explicitly, the next pair completes, and batch fails at the end.'
    $caught = $false
    $state.Calls.Clear()
    try { & $runner -ChecklistFile $checklist -OutputFolder $batchOutput -TargetSites 'https://unlisted.test' } catch {
        if ($_.Exception.Message -notmatch 'not in the checklist') { throw }
        $caught = $true
    }
    Assert-Equal $caught $true
    Assert-Equal $state.Calls.Count 0
    Write-Host 'PASS unlisted targets are rejected before authentication.'
    Write-Host 'All scout tests passed using offline mocks; no PnP module or live connection used.'
} finally {
    if (Test-Path -LiteralPath $temp) { Remove-Item -LiteralPath $temp -Recurse -Force }
}
