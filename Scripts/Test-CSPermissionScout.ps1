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
    Write-Host 'All scout tests passed using offline mocks; no PnP module or live connection used.'
} finally {
    if (Test-Path -LiteralPath $temp) { Remove-Item -LiteralPath $temp -Recurse -Force }
}
