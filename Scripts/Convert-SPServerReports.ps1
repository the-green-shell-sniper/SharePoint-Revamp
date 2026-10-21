#requires -Version 7.4.6
<#
.SYNOPSIS
    Previews or converts completed CSV/JSON reports at or above 99 MiB.
.DESCRIPTION
    CSVs retain their original path plus numbered, header-repeating siblings.
    JSON becomes readable token-record CSVs and a small index at its original
    path. Updated Scripts readers load indexed evidence automatically in memory.
    Verifies records and sizes before replacing originals. No backups, Git
    operations, or SharePoint writes. Stop all report writers before applying.
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [string]$Folder = (Join-Path $PSScriptRoot '..\Reports'),
    [switch]$Apply
)
. (Join-Path $PSScriptRoot 'SPOutputLimits.Common.ps1')

function Convert-SPServerReports {
    [CmdletBinding(SupportsShouldProcess)]
    param([Parameter(Mandatory)][string]$Folder, [switch]$Apply)
    $ErrorActionPreference = 'Stop'
    $root = Get-Item -LiteralPath $Folder -ErrorAction Stop
    if (-not $root.PSIsContainer) { throw "Not a report directory: $Folder" }
    $threshold = 99MB
    $files = @(Get-ChildItem -LiteralPath $root.FullName -Recurse -File -ErrorAction Stop |
        Where-Object Length -ge $threshold)
    foreach ($file in $files) {
        if ($file.Extension -inotmatch '^\.(csv|json)$') {
            Write-Warning "Unsupported oversized file; retained unchanged: $($file.FullName)"
            continue
        }
        if (-not $Apply) {
            [pscustomobject]@{
                File = $file.FullName; MiB = [math]::Round($file.Length / 1MB, 2)
                Action = if ($file.Extension -ieq '.json') { 'Convert JSON to indexed CSV records' } else { 'Split CSV' }
            }
            continue
        }
        if (-not $PSCmdlet.ShouldProcess($file.FullName, 'Convert, verify, and replace original without a backup')) { continue }
        [void](Assert-SPOutputPath $file.FullName)
        $id = [Convert]::ToHexString([Security.Cryptography.SHA256]::HashData(
            [Text.Encoding]::UTF8.GetBytes($file.FullName)
        )).Substring(0, 12).ToLowerInvariant()
        $csvPath = Assert-SPOutputPath (Join-Path $file.DirectoryName "j-$id.csv")
        $indexPath = Assert-SPOutputPath (Join-Path $file.DirectoryName "j-$id.tmp")
        if ((Test-Path -LiteralPath $csvPath) -or @(Get-SPCsvSiblingParts $csvPath).Count -or
            (Test-Path -LiteralPath $indexPath)) {
            throw "Conversion outputs already exist for '$($file.FullName)'. Inspect '$csvPath' before retrying. Original retained."
        }
        if (@(Get-SPCsvSiblingParts $file.FullName).Count) {
            throw "Existing CSV siblings must be inspected before converting '$($file.FullName)'. Original retained."
        }
        $sourceHash = [Security.Cryptography.IncrementalHash]::CreateHash([Security.Cryptography.HashAlgorithmName]::SHA256)
        $verifyHash = [Security.Cryptography.IncrementalHash]::CreateHash([Security.Cryptography.HashAlgorithmName]::SHA256)
        $source = [IO.File]::Open($file.FullName, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::Read)
        $reader = [IO.StreamReader]::new($source, [Text.UTF8Encoding]::new($false, $true), $true)
        $count = 0L; $verified = 0L
        try {
            if ($file.Extension -ieq '.json') {
                Get-SPJsonTokenRecords $reader | ForEach-Object {
                    Add-SPReportRecordHash $sourceHash $_
                    $count++
                    $_
                } | Export-Csv -LiteralPath $csvPath -NoTypeInformation -Encoding utf8
                $validator = [Newtonsoft.Json.JsonTextWriter]::new([IO.TextWriter]::Null)
                try {
                    Import-Csv -LiteralPath $csvPath -ErrorAction Stop | ForEach-Object {
                        Add-SPReportRecordHash $verifyHash $_
                        Write-SPJsonTokenRecord $validator $_
                        $verified++
                    }
                    if ($validator.WriteState -ne [Newtonsoft.Json.WriteState]::Start) {
                        throw 'Incomplete JSON container in converted records.'
                    }
                } finally { $validator.Close() }
            } else {
                # Import-Csv handles quoted multiline records; no physical-line splitting.
                Microsoft.PowerShell.Utility\Import-Csv -LiteralPath $file.FullName -Encoding utf8 -ErrorAction Stop |
                    ForEach-Object {
                        $record = ConvertTo-Json -InputObject $_ -Compress -Depth 5
                        $sourceHash.AppendData([Text.Encoding]::UTF8.GetBytes($record + "`n"))
                        $count++
                        $_
                    } | Export-Csv -LiteralPath $csvPath -NoTypeInformation -Encoding utf8
                Import-Csv -LiteralPath $csvPath -ErrorAction Stop | ForEach-Object {
                    $record = ConvertTo-Json -InputObject $_ -Compress -Depth 5
                    $verifyHash.AppendData([Text.Encoding]::UTF8.GetBytes($record + "`n"))
                    $verified++
                }
            }
            $digest = [Convert]::ToHexString($sourceHash.GetHashAndReset())
            if (-not $count -or $count -ne $verified -or
                $digest -cne [Convert]::ToHexString($verifyHash.GetHashAndReset())) {
                throw 'Record verification failed; original retained.'
            }
            $outputs = @((Get-Item -LiteralPath $csvPath)) + @(Get-SPCsvSiblingParts $csvPath)
            foreach ($output in $outputs) { Assert-SPOutputSize $output.FullName $output.Length }
            if ($file.Extension -ieq '.json') {
                Write-SPBoundedJson -InputObject ([ordered]@{
                    Format = 'SPReportRecordsV1'
                    Records = [IO.Path]::GetFileName($csvPath)
                    RecordCount = $count
                    SHA256 = $digest
                }) -Path $indexPath -Depth 5
            } else {
                $destinations = @($file.FullName) + @(
                    for ($i = 1; $i -lt $outputs.Count; $i++) { "$($file.FullName).part{0:D4}.csv" -f ($i + 1) }
                )
                foreach ($destination in $destinations) { [void](Assert-SPOutputPath $destination) }
            }
        } catch {
            throw "Conversion failed for '$($file.FullName)'; original retained. Inspect partial output '$csvPath'. $($_.Exception.Message)"
        } finally {
            $reader.Dispose()
            $source.Dispose()
            $sourceHash.Dispose()
            $verifyHash.Dispose()
        }
        $current = Get-Item -LiteralPath $file.FullName
        if ($current.Length -ne $file.Length -or $current.LastWriteTimeUtc -ne $file.LastWriteTimeUtc) {
            throw "Source changed during conversion; original retained: $($file.FullName)"
        }
        if ($file.Extension -ieq '.json') {
            [IO.File]::Move($indexPath, $file.FullName, $true)
        } else {
            # Publish later parts first; replace the original only after all later parts are present.
            for ($i = 1; $i -lt $outputs.Count; $i++) {
                [IO.File]::Move($outputs[$i].FullName, $destinations[$i], $false)
            }
            [IO.File]::Move($csvPath, $file.FullName, $true)
        }
        [pscustomobject]@{
            File = $file.FullName; Records = $count; Parts = $outputs.Count
            Status = 'Converted and verified; no backup'
        }
    }
    $remaining = @(Get-ChildItem -LiteralPath $root.FullName -Recurse -File -ErrorAction Stop |
        Where-Object Length -ge $threshold)
    if ($Apply -and -not $WhatIfPreference -and $remaining.Count) {
        foreach ($file in $remaining) { Write-Warning "Still at or above 99 MiB: $($file.FullName)" }
        throw "$($remaining.Count) oversized files remain. Do not commit them."
    }
    if (-not $files.Count) { Write-Host "No files at or above 99 MiB in '$($root.FullName)'." }
}

if ($MyInvocation.InvocationName -ne '.') {
    $PSBoundParameters['Folder'] = $Folder
    Convert-SPServerReports @PSBoundParameters
}
