#requires -Version 7.4.6
<#
.SYNOPSIS
    Read-only CS inventory diagnostics, using the repository's recursive PnP pattern.
.DESCRIPTION
    Defaults to archived CS and Clinical & Patient Services. SourceUrl and TargetUrl
    can select another explicit pair. Does not recurse into subsites.
    Does not resolve/create users or change content, permissions or authentication settings
    except enabling the existing PnP persistent login option.
    Enumerates all lists, including hidden/empty lists; compares recursive CAML item IDs
    with paged REST item IDs and records each reported count before/after enumeration.
    Failures are recorded explicitly; a failed scout is never evidence of complete coverage.
    Copy SPPermission.Common.ps1 beside this script. Reports contain sensitive paths;
    keep them private and do not commit them. No file content is downloaded.
.EXAMPLE
    .\Get-CSPermissionScout.ps1
#>
[CmdletBinding()]
param(
    [string]$SourceUrl = 'https://whallc.sharepoint.com/sites/archive-2026-10-18T115436Z/CS',
    [string]$TargetUrl = 'https://whallc.sharepoint.com/sites/ClinicalandPatientServices',
    [string]$OutputFolder = (Join-Path $PSScriptRoot '..\Reports\PermissionsReports\CS-PermissionScout'),
    [string]$ClientId = '2e20d82f-56ba-418e-b465-9718493f0df8',
    [switch]$SkipPersistLoginSetup
)
. (Join-Path $PSScriptRoot 'SPPermission.Common.ps1')

function Get-CSScoutQuery {
    param([string[]]$FieldNames)
    $names = @(@('ID') + $FieldNames | Sort-Object -Unique)
    $viewFields = ($names | ForEach-Object {
        "<FieldRef Name='{0}'/>" -f [Security.SecurityElement]::Escape($_)
    }) -join ''
    "<View Scope='RecursiveAll'><ViewFields>$viewFields</ViewFields><RowLimit Paged='TRUE'>500</RowLimit></View>"
}

function Invoke-CSPermissionScout {
    $ErrorActionPreference = 'Stop'
    Assert-SPPermissionUrls $SourceUrl $TargetUrl
    $SourceUrl = $SourceUrl.TrimEnd('/')
    $TargetUrl = $TargetUrl.TrimEnd('/')
    Import-Module PnP.PowerShell -MinimumVersion 3.3.0 -ErrorAction Stop
    $folder = Join-Path $OutputFolder (Get-Date -Format 'yyyyMMdd-HHmmss-fff')
    [void][IO.Directory]::CreateDirectory($folder)
    $summary = [Collections.Generic.List[object]]::new()
    $fieldsReport = [Collections.Generic.List[object]]::new()
    $itemsReport = [Collections.Generic.List[object]]::new()
    $errorsReport = [Collections.Generic.List[object]]::new()
    $sites = @(
        @{ Side = 'Source'; Url = $SourceUrl }
        @{ Side = 'Target'; Url = $TargetUrl }
    )
    try {
        foreach ($site in $sites) {
            $connection = Connect-SPPermissionSite $site.Url $ClientId '' -SkipPersistLoginSetup:$SkipPersistLoginSetup
            $web = Get-PnPWeb -Connection $connection -Includes Url,Id,HasUniqueRoleAssignments -ErrorAction Stop
            if ($web.Url.TrimEnd('/') -ine $site.Url) { throw "Unexpected web: $($web.Url)" }
            $lists = @(Get-PnPList -Connection $connection -Includes Id,Title,Hidden,BaseType,BaseTemplate,ItemCount,RootFolder.ServerRelativeUrl,HasUniqueRoleAssignments -ErrorAction Stop)
            foreach ($list in $lists) {
                Write-Host "[$($site.Side)] Scouting '$($list.Title)' (reported count $($list.ItemCount))"
                $record = [pscustomobject]@{
                    Side = $site.Side; WebUrl = $site.Url; ListId = [string]$list.Id
                    Title = $list.Title; RootPath = $list.RootFolder.ServerRelativeUrl; Hidden = $list.Hidden
                    BaseTemplate = $list.BaseTemplate; Unique = $list.HasUniqueRoleAssignments
                    ReportedBefore = $list.ItemCount; ReportedAfter = $null
                    PnPCount = $null; RestCount = $null; OnlyPnPIds = ''; OnlyRestIds = ''
                    InheritanceReadFailures = 0; Status = 'Incomplete'
                }
                $summary.Add($record)
                try {
                    $fields = @(Get-PnPField -List $list.Id -Connection $connection -ErrorAction Stop)
                    foreach ($field in $fields) {
                        $fieldsReport.Add([pscustomobject]@{
                            Side = $site.Side; ListId = [string]$list.Id; ListTitle = $list.Title
                            InternalName = $field.InternalName; TypeAsString = $field.TypeAsString
                            Hidden = $field.Hidden; ReadOnly = $field.ReadOnlyField
                        })
                    }
                    $selected = @($fields.InternalName | Where-Object { $_ -in @('FileRef','FileLeafRef','FSObjType','Title','Created','Author') })
                    $query = Get-CSScoutQuery $selected
                    $items = @(Get-PnPListItem -List $list.Id -Query $query -PageSize 500 -Connection $connection -ErrorAction Stop)
                    $record.PnPCount = $items.Count
                    $pnpIds = [Collections.Generic.HashSet[int]]::new()
                    foreach ($item in $items) {
                        if (-not $pnpIds.Add([int]$item.Id)) { throw "Duplicate PnP item ID: $($item.Id)" }
                        $unique = $null
                        try {
                            $unique = Get-PnPProperty -ClientObject $item -Property HasUniqueRoleAssignments -Connection $connection -ErrorAction Stop
                            if ($null -eq $unique) { throw 'Inheritance state was not returned.' }
                        } catch {
                            $record.InheritanceReadFailures++
                            $errorsReport.Add([pscustomobject]@{
                                Side = $site.Side; ListId = [string]$list.Id; ItemId = $item.Id
                                Stage = 'Read inheritance'; Error = $_.Exception.Message
                            })
                        }
                        $itemsReport.Add([pscustomobject]@{
                            Side = $site.Side; ListId = [string]$list.Id; ListTitle = $list.Title; ItemId = $item.Id
                            Path = $item.FieldValues['FileRef']; Leaf = $item.FieldValues['FileLeafRef']
                            FSObjType = $item.FieldValues['FSObjType']; Unique = $unique
                        })
                    }
                    $restIds = [Collections.Generic.HashSet[int]]::new()
                    $last = 0
                    do {
                        $url = $site.Url + "/_api/web/lists(guid'$($list.Id)')/items?`$select=Id&`$filter=Id gt $last&`$orderby=Id&`$top=500"
                        $batch = @(Get-SPPermissionCollection $connection $url)
                        foreach ($item in $batch) {
                            if ([int]$item.Id -le $last -or -not $restIds.Add([int]$item.Id)) { throw 'Invalid or repeated REST item ID.' }
                        }
                        if ($batch.Count) { $last = [int]$batch[-1].Id }
                    } while ($batch.Count -ge 500)
                    $record.RestCount = $restIds.Count
                    $record.OnlyPnPIds = (@($pnpIds | Where-Object { -not $restIds.Contains($_) } | Sort-Object) -join ';')
                    $record.OnlyRestIds = (@($restIds | Where-Object { -not $pnpIds.Contains($_) } | Sort-Object) -join ';')
                    $fresh = Get-PnPList -Identity $list.Id -Connection $connection -Includes ItemCount -ErrorAction Stop
                    $record.ReportedAfter = $fresh.ItemCount
                    $record.Status = if ($record.InheritanceReadFailures) { 'Inheritance failures - review' }
                        elseif ($record.OnlyPnPIds -or $record.OnlyRestIds) { 'PnP/REST disagreement - review' }
                        elseif ($record.ReportedBefore -ne $record.ReportedAfter) { 'Reported count changed - review' }
                        elseif ($record.PnPCount -ne $record.ReportedAfter) { 'Enumerators agree; reported count differs - review' }
                        else { 'Counts and IDs agree' }
                } catch {
                    $record.Status = 'List scout failed - review'
                    $errorsReport.Add([pscustomobject]@{
                        Side = $site.Side; ListId = [string]$list.Id; ItemId = ''
                        Stage = 'Inventory/compare'; Error = $_.Exception.Message
                    })
                    Write-Warning "Scout failed for '$($list.Title)': $($_.Exception.Message)"
                }
            }
        }
    } catch {
        $errorsReport.Add([pscustomobject]@{ Side = $site.Side; ListId = ''; ItemId = ''; Stage = 'Web/connection'; Error = $_.Exception.Message })
        throw
    } finally {
        Export-SPPermissionCsv $summary.ToArray() @('Side','WebUrl','ListId','Title','RootPath','Hidden','BaseTemplate','Unique','ReportedBefore','ReportedAfter','PnPCount','RestCount','OnlyPnPIds','OnlyRestIds','InheritanceReadFailures','Status') (Join-Path $folder 'Lists.csv')
        Export-SPPermissionCsv $fieldsReport.ToArray() @('Side','ListId','ListTitle','InternalName','TypeAsString','Hidden','ReadOnly') (Join-Path $folder 'Fields.csv')
        Export-SPPermissionCsv $itemsReport.ToArray() @('Side','ListId','ListTitle','ItemId','Path','Leaf','FSObjType','Unique') (Join-Path $folder 'Items.csv')
        Export-SPPermissionCsv $errorsReport.ToArray() @('Side','ListId','ItemId','Stage','Error') (Join-Path $folder 'Errors.csv')
        Write-Host "Scout reports: $folder"
    }
    $reviewCount = @($summary | Where-Object Status -ne 'Counts and IDs agree').Count
    Write-Host "Read-only scout finished. Lists needing review: $reviewCount; recorded errors: $($errorsReport.Count). No permission changes made."
    if ($errorsReport.Count) { throw 'Scout had read failures. Reports were saved; do not interpret this run as complete.' }
}

if ($MyInvocation.InvocationName -ne '.') { Invoke-CSPermissionScout }
