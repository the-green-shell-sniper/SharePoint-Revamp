#requires -Version 7.4.6
<#
.SYNOPSIS
    Splits completed oversized CSV reports into readable parts below 90 MiB.
.DESCRIPTION
    Previews by default. With -Apply, streams CSV records (including multiline
    fields), repeats the header, verifies part sizes, then deletes the original.
    No backups, Git changes, JSON changes, or SharePoint writes. Run only after
    permission pipelines finish: their evidence paths reference original files.
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [string]$Folder = (Join-Path $PSScriptRoot '..\Reports\VisitorsRuns'),
    [ValidateRange(1, 90)][int]$MaxPartMiB = 90,
    [switch]$Apply
)
. (Join-Path $PSScriptRoot 'SPOutputLimits.Common.ps1')

function Split-SPVisitorReports {
    $ErrorActionPreference = 'Stop'
    $root = Get-Item -LiteralPath $Folder
    if (-not $root.PSIsContainer) { throw "Not a report directory: $Folder" }
    $limit = [long]$MaxPartMiB * 1MB
    $files = @(Get-ChildItem -LiteralPath $root.FullName -Recurse -File -Filter '*.csv' |
        Where-Object Length -ge $limit)
    if (-not $files.Count) {
        Write-Host "No CSV reports at or above $MaxPartMiB MiB in '$($root.FullName)'."
    }
    foreach ($file in $files) {
        if (-not $Apply) {
            [pscustomobject]@{ File = $file.FullName; Bytes = $file.Length; Status = 'Would split' }
            continue
        }
        if (-not $PSCmdlet.ShouldProcess($file.FullName, 'Split CSV and delete original after verification (no backup)')) {
            continue
        }
        $partsFolder = $file.FullName + '.parts'
        [void](Assert-SPOutputPath $partsFolder)
        if (Test-Path -LiteralPath $partsFolder) {
            throw "Output already exists: '$partsFolder'. Inspect/remove it before retrying; original retained."
        }
        [void][IO.Directory]::CreateDirectory($partsFolder)
        $encoding = [Text.UTF8Encoding]::new($false, $true)
        $writer = $null
        $header = $null
        $headerBytes = 0L
        $partBytes = 0L
        $rows = 0L
        $parts = [Collections.Generic.List[string]]::new()
        try {
            Microsoft.PowerShell.Utility\Import-Csv -LiteralPath $file.FullName -Encoding utf8 | ForEach-Object {
                # Convert one logical record, not one physical line: fields can contain newlines.
                $csv = @(ConvertTo-Csv -InputObject $_ -NoTypeInformation)
                if ($null -eq $header) {
                    $header = $csv[0]
                    $headerBytes = $encoding.GetByteCount($header + "`r`n")
                }
                $record = $csv[1]
                $recordBytes = $encoding.GetByteCount($record + "`r`n")
                if ($headerBytes + $recordBytes -ge $limit) {
                    throw "A single CSV record plus its header cannot fit below $MaxPartMiB MiB."
                }
                if ($null -eq $writer -or $partBytes + $recordBytes -ge $limit) {
                    if ($null -ne $writer) { $writer.Dispose(); $writer = $null }
                    $path = Join-Path $partsFolder ('{0}.part{1:D4}.csv' -f $file.BaseName, ($parts.Count + 1))
                    [void](Assert-SPOutputPath $path)
                    $writer = [IO.StreamWriter]::new($path, $false, $encoding)
                    $writer.NewLine = "`r`n"
                    $parts.Add($path)
                    $writer.WriteLine($header)
                    $partBytes = $headerBytes
                }
                $writer.WriteLine($record)
                $partBytes += $recordBytes
                $rows++
            }
            if ($null -ne $writer) { $writer.Dispose(); $writer = $null }
            if (-not $rows) { throw 'No CSV data records found; cannot safely split this file.' }
            foreach ($path in $parts) {
                if ((Get-Item -LiteralPath $path).Length -ge $limit) {
                    throw "Part is not below the size limit: '$path'."
                }
            }
            $current = Get-Item -LiteralPath $file.FullName
            if ($current.Length -ne $file.Length -or $current.LastWriteTimeUtc -ne $file.LastWriteTimeUtc) {
                throw 'Source changed while splitting. Stop the report writer before retrying.'
            }
        } catch {
            throw "Split failed for '$($file.FullName)'. Original retained; inspect partial output at '$partsFolder'. $($_.Exception.Message)"
        } finally {
            if ($null -ne $writer) { $writer.Dispose() }
        }
        Remove-Item -LiteralPath $file.FullName -ErrorAction Stop
        [pscustomobject]@{
            File = $file.FullName; PartsFolder = $partsFolder
            Rows = $rows; Parts = $parts.Count; Status = 'Split; original deleted'
        }
    }
}

if ($MyInvocation.InvocationName -ne '.') { Split-SPVisitorReports }
