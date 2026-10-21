#requires -Version 7.4.6
<#
.SYNOPSIS
    Preview or relocate completed permission-run reports to short archive paths.
.DESCRIPTION
    Never run against active workers or review folders. Defaults to preview only.
    Apply writes a durable path/hash manifest before moving any files. Contents are
    unchanged; embedded old paths are resolved through the manifest, not rewritten.
    No directories are deleted. A partial move can be resumed with the same options.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$SourceFolder,
    [string]$ArchiveRoot = (Join-Path $PSScriptRoot '..\Reports\PermArchive'),
    [switch]$Apply
)
. (Join-Path $PSScriptRoot 'SPOutputLimits.Common.ps1')
$ErrorActionPreference = 'Stop'
$source = (Resolve-Path -LiteralPath $SourceFolder -ErrorAction Stop).ProviderPath.TrimEnd('\')
$archive = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ArchiveRoot).TrimEnd('\')
$repo = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..')).TrimEnd('\')
if ($source -ieq $repo -or -not $source.StartsWith("$repo\Reports\", [StringComparison]::OrdinalIgnoreCase)) {
    throw 'SourceFolder must be a specific completed report folder beneath this repository Reports folder.'
}
if ($archive -ieq $source -or $archive.StartsWith("$source\", [StringComparison]::OrdinalIgnoreCase) -or
    $source.StartsWith("$archive\", [StringComparison]::OrdinalIgnoreCase)) {
    throw 'Source and archive must not overlap.'
}
$sourceHash = [Convert]::ToHexString([Security.Cryptography.SHA256]::HashData(
    [Text.Encoding]::UTF8.GetBytes($source.ToLowerInvariant()))).Substring(0, 16)
$destination = Join-Path $archive $sourceHash
$manifestPath = Join-Path $destination 'Paths.csv'
$rows = [Collections.Generic.List[object]]::new()
if (Test-Path -LiteralPath $manifestPath) {
    foreach ($row in Import-Csv -LiteralPath $manifestPath) { $rows.Add($row) }
}
foreach ($file in Get-ChildItem -LiteralPath $source -Recurse -File) {
    if (@($rows | Where-Object OriginalPath -IEQ $file.FullName).Count) { continue }
    $relative = [IO.Path]::GetRelativePath($source, $file.FullName)
    $id = [Convert]::ToHexString([Security.Cryptography.SHA256]::HashData(
        [Text.Encoding]::UTF8.GetBytes($relative))).Substring(0, 24)
    $target = Join-Path $destination ($id + $file.Extension)
    if ($target.Length -gt 240) { throw "Archive path still exceeds 240 characters: $target. Choose a shorter ArchiveRoot." }
    $rows.Add([pscustomobject]@{
        OriginalPath = $file.FullName; RelativePath = $relative; ArchivePath = $target
        SHA256 = (Get-FileHash -LiteralPath $file.FullName -Algorithm SHA256).Hash
    })
}
if (-not $Apply) {
    $rows | Select-Object RelativePath, ArchivePath
    Write-Host "Preview: $($rows.Count) files. Use -Apply only after all workers using this report folder have stopped."
    return
}
[void](Assert-SPOutputPath $destination)
[void][IO.Directory]::CreateDirectory($destination)
$rows | Export-Csv -LiteralPath "$manifestPath.tmp" -NoTypeInformation -Encoding utf8
Move-Item -LiteralPath "$manifestPath.tmp" -Destination $manifestPath -Force
foreach ($row in $rows) {
    if (Test-Path -LiteralPath $row.ArchivePath) {
        if ((Get-FileHash -LiteralPath $row.ArchivePath).Hash -cne $row.SHA256) { throw "Archive hash mismatch: $($row.ArchivePath)" }
        if (Test-Path -LiteralPath $row.OriginalPath) { throw "Both source and archive exist; refusing to remove either: $($row.OriginalPath)" }
        continue
    }
    if ((Get-FileHash -LiteralPath $row.OriginalPath).Hash -cne $row.SHA256) {
        throw "Source changed since manifest creation: $($row.OriginalPath)"
    }
    Move-Item -LiteralPath $row.OriginalPath -Destination $row.ArchivePath -ErrorAction Stop
    if ((Get-FileHash -LiteralPath $row.ArchivePath).Hash -cne $row.SHA256) { throw "Post-move hash mismatch: $($row.ArchivePath)" }
}
Write-Host "Archived $($rows.Count) files. Original path lookup and hashes: $manifestPath"
