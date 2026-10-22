#requires -Version 7.4.6
<#
.SYNOPSIS
    Read-only comparison of failed System Editor writes with their saved metadata.
.DESCRIPTION
    Reads Failed (Verifying) journal entries and Before.csv, then saves current
    fields and exact differences. Makes no item updates. Keep reports private.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$RunFolder,
    [string]$ClientId = '2e20d82f-56ba-418e-b465-9718493f0df8'
)
. (Join-Path $PSScriptRoot 'Invoke-SystemEditorRemediation.ps1')
$ErrorActionPreference = 'Stop'
$RunFolder = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($RunFolder)
$failed = @(Import-Csv -LiteralPath (Join-Path $RunFolder 'Journal.csv') |
    Where-Object Status -eq 'Failed (Verifying)')
if (-not $failed.Count) { throw 'No failed verification entries found.' }
$beforeRows = @(Import-Csv -LiteralPath (Join-Path $RunFolder 'Before.csv'))
$output = Assert-SPOutputPath (Join-Path $RunFolder ('check-' + (Get-Date -Format 'yyMMdd-HHmmss')))
[void][IO.Directory]::CreateDirectory($output)
$afterPath = Join-Path $output 'After.csv'
$diffPath = Join-Path $output 'Differences.csv'
'WebUrl,ListId,ItemId,Field,ValueXml' | Set-Content -LiteralPath $afterPath
'WebUrl,ListId,ItemId,Field,Before,After' | Set-Content -LiteralPath $diffPath
Import-Module PnP.PowerShell -MinimumVersion 3.3.0 -ErrorAction Stop
foreach ($entry in $failed) {
    $url = ConvertTo-SystemEditorWebUrl $entry.WebUrl
    $saved = @($beforeRows | Where-Object {
        $_.WebUrl -ieq $url -and $_.ListId -ieq $entry.ListId -and $_.ItemId -eq $entry.ItemId
    })
    if (-not $saved.Count) { throw 'Missing pre-write metadata for failed item.' }
    $connection = Connect-PnPOnline -Url $url -ClientId $ClientId -Interactive -ReturnConnection -ErrorAction Stop
    $item = Get-PnPListItem -List ([guid]$entry.ListId) -Id ([int]$entry.ItemId) -Connection $connection -ErrorAction Stop
    $identity = @($saved | Where-Object Field -eq 'UniqueId')
    if ($identity.Count -ne 1 -or
        [guid]$item['UniqueId'] -ne [guid][Management.Automation.PSSerializer]::Deserialize($identity[0].ValueXml)) {
        throw 'Item identity differs from backup. Stopping.'
    }
    $after = Get-SystemEditorFieldSnapshot $item
    $after.Keys | Sort-Object | ForEach-Object {
        [pscustomobject]@{
            WebUrl = $url; ListId = $entry.ListId; ItemId = $entry.ItemId
            Field = $_; ValueXml = $after[$_]
        }
    } | Export-Csv -LiteralPath $afterPath -Append -NoTypeInformation
    $before = @{}
    foreach ($row in $saved) { $before[$row.Field] = $row.ValueXml }
    foreach ($field in @(@($before.Keys) + @($after.Keys) | Sort-Object -Unique)) {
        if ($before[$field] -ceq $after[$field]) { continue }
        $old = if ($before.ContainsKey($field)) { [Management.Automation.PSSerializer]::Deserialize($before[$field]) } else { $null }
        $new = if ($after.ContainsKey($field)) { [Management.Automation.PSSerializer]::Deserialize($after[$field]) } else { $null }
        if ($field -eq 'MetaInfo') {
            $oldMap = @{}
            $newMap = @{}
            foreach ($side in @(@{ Text = $old; Map = $oldMap }, @{ Text = $new; Map = $newMap })) {
                foreach ($line in ([string]$side.Text -split '\r?\n' | Where-Object { $_ -ne '' })) {
                    if ($line -notmatch '^(.*):([A-Z]{2})\|(.*)$') { throw 'Unrecognized MetaInfo entry; full fields saved in After.csv.' }
                    $key = $Matches[1]
                    if ($side.Map.ContainsKey($key)) { throw "Duplicate MetaInfo key: $key" }
                    $side.Map[$key] = "$($Matches[2])|$($Matches[3])"
                }
            }
            foreach ($key in @(@($oldMap.Keys) + @($newMap.Keys) | Sort-Object -Unique)) {
                if ($oldMap[$key] -ceq $newMap[$key]) { continue }
                [pscustomobject]@{
                    WebUrl = $url; ListId = $entry.ListId; ItemId = $entry.ItemId
                    Field = "MetaInfo/$key"; Before = $oldMap[$key]; After = $newMap[$key]
                } | Export-Csv -LiteralPath $diffPath -Append -NoTypeInformation
            }
        } else {
            [pscustomobject]@{
                WebUrl = $url; ListId = $entry.ListId; ItemId = $entry.ItemId
                Field = $field; Before = $before[$field]; After = $after[$field]
            } | Export-Csv -LiteralPath $diffPath -Append -NoTypeInformation
        }
    }
}
Write-Host "Read-only comparison saved: $output"
