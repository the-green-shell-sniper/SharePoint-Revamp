# Shared output safeguards. Dot-source from each repository script after its param block.
$script:SPOutputPartLimit = 90MB
. (Join-Path $PSScriptRoot 'SPReportRecords.Common.ps1')

function Assert-SPOutputPath {
    param([Parameter(Mandatory)][string]$Path)
    $fullPath = [IO.Path]::GetFullPath($Path)
    if ($fullPath.Length -ge 256) {
        throw "Output path is $($fullPath.Length) characters; the repository limit is fewer than 256: $fullPath"
    }
    return $fullPath
}

function Assert-SPOutputSize {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][long]$Bytes
    )
    [void](Assert-SPOutputPath $Path)
    if ($Bytes -ge $script:SPOutputPartLimit) {
        throw "Output is $Bytes bytes; each generated file must remain below 90 MiB: $Path"
    }
}

function Get-SPOutputEncoding {
    param([string]$Name = 'utf8')
    switch -Regex ($Name) {
        '^(?i:utf8|utf8NoBOM)$' { return [Text.UTF8Encoding]::new($false) }
        '^(?i:utf8BOM)$' { return [Text.UTF8Encoding]::new($true) }
        '^(?i:unicode|utf16)$' { return [Text.Encoding]::Unicode }
        '^(?i:bigendianunicode)$' { return [Text.Encoding]::BigEndianUnicode }
        '^(?i:utf32)$' { return [Text.Encoding]::UTF32 }
        '^(?i:ascii)$' { return [Text.Encoding]::ASCII }
        '^(?i:default)$' { return [Text.Encoding]::Default }
        default { return [Text.Encoding]::GetEncoding($Name) }
    }
}

function Get-SPCsvSiblingParts {
    param([Parameter(Mandatory)][string]$Path)
    $directory = [IO.Path]::GetDirectoryName($Path)
    $name = [IO.Path]::GetFileName($Path)
    $pattern = "$name.part*.csv"
    $namePattern = '^' + [regex]::Escape($name) + '\.part(\d+)\.csv$'
    @(
        Get-ChildItem -LiteralPath $directory -File -Filter $pattern -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -match $namePattern } |
            Sort-Object { [int]([regex]::Match($_.Name, $namePattern).Groups[1].Value) }
    )
}

function Get-SPTextSiblingParts {
    param([Parameter(Mandatory)][string]$Path)
    $directory = [IO.Path]::GetDirectoryName($Path)
    $name = [IO.Path]::GetFileName($Path)
    $extension = [IO.Path]::GetExtension($Path)
    $pattern = "$name.part*$extension"
    $namePattern = '^' + [regex]::Escape($name) + '\.part(\d+)' + [regex]::Escape($extension) + '$'
    @(
        Get-ChildItem -LiteralPath $directory -File -Filter $pattern -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -match $namePattern } |
            Sort-Object { [int]([regex]::Match($_.Name, $namePattern).Groups[1].Value) }
    )
}

function Export-Csv {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ValueFromPipeline)][object]$InputObject,
        [Parameter(Position = 0)][string]$Path,
        [string]$LiteralPath,
        [switch]$Append,
        [switch]$NoTypeInformation,
        [switch]$UseCulture,
        [string]$Delimiter = ',',
        [string]$Encoding = 'utf8',
        [switch]$Force
    )
    begin {
        $target = if ($LiteralPath) { $LiteralPath } else { $Path }
        if (-not $target) { throw 'Export-Csv requires Path or LiteralPath.' }
        $target = Assert-SPOutputPath $target
        $codec = Get-SPOutputEncoding $Encoding
        $newline = "`r`n"
        $newlineBytes = $codec.GetByteCount($newline)
        $cultureDelimiter = if ($UseCulture) {
            [Globalization.CultureInfo]::CurrentCulture.TextInfo.ListSeparator
        } else { $Delimiter }
        $siblings = @(Get-SPCsvSiblingParts $target)
        if (-not $Append) {
            if (Test-Path -LiteralPath $target -PathType Leaf) { Remove-Item -LiteralPath $target -Force:$Force -ErrorAction Stop }
            foreach ($part in $siblings) { Remove-Item -LiteralPath $part.FullName -Force -ErrorAction Stop }
            $siblings = @()
        }
        $currentPath = $null
        $currentBytes = 0L
        $header = $null
        $headerBytes = 0L
        $writer = $null
        $appendExisting = $false
        if ($Append) {
            $existing = @()
            if (Test-Path -LiteralPath $target -PathType Leaf) { $existing += Get-Item -LiteralPath $target }
            $existing += $siblings
            if ($existing.Count) {
                $last = $existing[-1]
                if ($last.Length -ge $script:SPOutputPartLimit) {
                    $partNumber = if ($last.FullName -eq $target) { 2 } else {
                        [int]([regex]::Match($last.Name, '\.part(\d+)\.csv$').Groups[1].Value) + 1
                    }
                    $currentPath = "$target.part{0:D4}.csv" -f $partNumber
                    $currentPath = Assert-SPOutputPath $currentPath
                    $currentBytes = 0L
                    $appendExisting = $false
                } else {
                    $currentPath = $last.FullName
                    $currentBytes = $last.Length
                    $appendExisting = $true
                    $first = Microsoft.PowerShell.Utility\Import-Csv -LiteralPath $last.FullName |
                        Select-Object -First 1
                    if ($null -ne $first) { $header = @($first.PSObject.Properties.Name) }
                }
            }
        } else {
            $currentPath = $target
        }
    }
    process {
        $row = $InputObject
        if ($header) {
            $ordered = [ordered]@{}
            foreach ($column in $header) {
                $property = $row.PSObject.Properties[$column]
                $ordered[$column] = if ($property) { $property.Value } else { $null }
            }
            $row = [pscustomobject]$ordered
        }
        $convert = @{
            InputObject = $row
            Delimiter = $cultureDelimiter
            NoTypeInformation = $true
        }
        $csvLines = @(Microsoft.PowerShell.Utility\ConvertTo-Csv @convert)
        if ($csvLines.Count -ge 2) {
            $record = $csvLines[-1]
            if ($null -eq $header) {
                $header = @($row.PSObject.Properties.Name)
                $headerBytes = 0L
                foreach ($line in $csvLines[0..($csvLines.Count - 2)]) {
                    $headerBytes += $codec.GetByteCount($line + $newline)
                }
            }
            $recordBytes = [long]$codec.GetByteCount($record) + $newlineBytes
            if ($recordBytes + $headerBytes -ge $script:SPOutputPartLimit) {
                throw "A single CSV row plus its header cannot fit below 90 MiB: $target"
            }
            if ($null -eq $currentPath) { $currentPath = $target }
            $headerPending = -not $appendExisting -and $currentBytes -eq 0
            $preambleBytes = if ($currentBytes -eq 0) { $codec.GetPreamble().Length } else { 0 }
            $projected = $currentBytes + $preambleBytes + $(if ($headerPending) { $headerBytes } else { 0 }) + $recordBytes
            if ($projected -ge $script:SPOutputPartLimit) {
                if ($writer) { $writer.Dispose(); $writer = $null }
                $csvParts = @(Get-SPCsvSiblingParts $target)
                $partNumber = if ($csvParts.Count) {
                    [int]([regex]::Match($csvParts[-1].Name, '\.part(\d+)\.csv$').Groups[1].Value) + 1
                } else { 2 }
                $currentPath = "$target.part{0:D4}.csv" -f $partNumber
                $currentPath = Assert-SPOutputPath $currentPath
                $currentBytes = 0L
                $appendExisting = $false
                $headerPending = $true
                $preambleBytes = $codec.GetPreamble().Length
            }
            if ($null -eq $writer) {
                $directory = [IO.Path]::GetDirectoryName($currentPath)
                if (-not [IO.Directory]::Exists($directory)) { [void][IO.Directory]::CreateDirectory($directory) }
                $mode = if ($appendExisting) { [IO.FileMode]::Append } else { [IO.FileMode]::CreateNew }
                $stream = [IO.File]::Open($currentPath, $mode, [IO.FileAccess]::Write, [IO.FileShare]::Read)
                $writer = [IO.StreamWriter]::new($stream, $codec)
                $writer.NewLine = $newline
            }
            if ($headerPending) {
                foreach ($line in $csvLines[0..($csvLines.Count - 2)]) { $writer.WriteLine($line) }
                $currentBytes += $preambleBytes + $headerBytes
                $appendExisting = $true
            }
            $writer.WriteLine($record)
            $currentBytes += $recordBytes
            if ($currentBytes -ge $script:SPOutputPartLimit) {
                throw "CSV part reached the 90 MiB safety threshold unexpectedly: $currentPath"
            }
        }
    }
    end {
        if ($writer) { $writer.Dispose() }
        $outputs = @()
        if (Test-Path -LiteralPath $target -PathType Leaf) { $outputs += Get-Item -LiteralPath $target }
        $outputs += Get-SPCsvSiblingParts $target
        foreach ($output in $outputs) {
            if ($output.Length -ge $script:SPOutputPartLimit) {
                throw "CSV output reached the 90 MiB safety threshold: $($output.FullName)"
            }
            [void](Assert-SPOutputPath $output.FullName)
        }
    }
    clean {
        if ($writer) { $writer.Dispose() }
    }
}

function Import-Csv {
    [CmdletBinding()]
    param(
        [Parameter(Position = 0, ValueFromPipeline, ValueFromPipelineByPropertyName)][string[]]$Path,
        [string[]]$LiteralPath,
        [string[]]$Header,
        [string]$Delimiter = ',',
        [string]$Encoding = 'utf8',
        [switch]$UseCulture,
        [switch]$IncludeTypeInformation
    )
    process {
        $paths = if ($LiteralPath) { $LiteralPath } else { $Path }
        foreach ($inputPath in $paths) {
            if ([Management.Automation.WildcardPattern]::ContainsWildcardCharacters($inputPath)) {
                $nativeOptions = @{ Encoding = $Encoding }
                if ($UseCulture) { $nativeOptions.UseCulture = $true }
                else { $nativeOptions.Delimiter = $Delimiter }
                if ($Header) { $nativeOptions.Header = $Header }
                foreach ($file in Get-ChildItem -Path $inputPath -File -ErrorAction Stop) {
                    Microsoft.PowerShell.Utility\Import-Csv -LiteralPath $file.FullName @nativeOptions
                }
                continue
            }
            $target = [IO.Path]::GetFullPath($inputPath)
            $parts = @(Get-SPCsvSiblingParts $target)
            if ($parts.Count -and $Header) {
                throw "A custom Header cannot be used when importing split CSV output: $target"
            }
            $files = @()
            if (Test-Path -LiteralPath $target -PathType Leaf) { $files += Get-Item -LiteralPath $target }
            $files += $parts
            $nativeOptions = @{
                Encoding = $Encoding
            }
            if ($UseCulture) { $nativeOptions.UseCulture = $true }
            else { $nativeOptions.Delimiter = $Delimiter }
            if ($Header) { $nativeOptions.Header = $Header }
            if (-not $files.Count) {
                Microsoft.PowerShell.Utility\Import-Csv -LiteralPath $target @nativeOptions
                continue
            }
            foreach ($file in $files) {
                Microsoft.PowerShell.Utility\Import-Csv -LiteralPath $file.FullName @nativeOptions
            }
        }
    }
}

function Set-Content {
    [CmdletBinding(DefaultParameterSetName = 'Path')]
    param(
        [Parameter(Mandatory, Position = 1, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [Alias('Value')][object]$InputObject,
        [Parameter(Position = 0)][string[]]$Path,
        [string[]]$LiteralPath,
        [string]$Encoding = 'utf8',
        [switch]$NoNewline,
        [switch]$Force,
        [switch]$PassThru
    )
    begin { $values = [Collections.Generic.List[object]]::new() }
    process { $values.Add($InputObject) }
    end {
        $targets = if ($LiteralPath) { $LiteralPath } else { $Path }
        if (-not $targets.Count) { throw 'Set-Content requires Path or LiteralPath.' }
        $codec = Get-SPOutputEncoding $Encoding
        $text = (@($values | ForEach-Object { [string]$_ }) -join [Environment]::NewLine)
        if (-not $NoNewline) { $text += [Environment]::NewLine }
        foreach ($target in $targets) {
            $fullPath = Assert-SPOutputPath $target
            $bytes = $codec.GetPreamble() + $codec.GetBytes($text)
            if ($bytes.LongLength -ge $script:SPOutputPartLimit) {
                $format = if ([IO.Path]::GetExtension($fullPath) -ieq '.json') {
                    'Write large datasets as record-based CSV parts; opaque JSON chunks are not supported.'
                } else { 'Partition this output into bounded records before writing.' }
                throw "Output would reach 90 MiB. $format Path: $fullPath"
            }
            Assert-SPOutputSize -Path $fullPath -Bytes $bytes.LongLength
            $parameters = @{
                LiteralPath = $fullPath; Encoding = $Encoding; NoNewline = $NoNewline
                Force = $Force; PassThru = $PassThru
            }
            $values | Microsoft.PowerShell.Management\Set-Content @parameters
        }
    }
}

function Write-SPBoundedJson {
    param(
        [Parameter(Mandatory)]$InputObject,
        [Parameter(Mandatory)][string]$Path,
        [ValidateRange(2, 100)][int]$Depth = 100
    )
    $target = Assert-SPOutputPath $Path
    $json = ConvertTo-Json -InputObject $InputObject -Depth $Depth
    $bytes = [Text.UTF8Encoding]::new($false).GetBytes($json + [Environment]::NewLine)
    if ($bytes.LongLength -ge $script:SPOutputPartLimit) {
        throw "JSON document would reach 90 MiB. Write its records as bounded CSV parts instead: $target"
    }
    Assert-SPOutputSize -Path $target -Bytes $bytes.LongLength
    [IO.File]::WriteAllBytes($target, $bytes)
}

function Get-Content {
    [CmdletBinding(DefaultParameterSetName = 'Path')]
    param(
        [Parameter(Mandatory, Position = 0, ValueFromPipeline, ValueFromPipelineByPropertyName, ParameterSetName = 'Path')]
        [string[]]$Path,
        [Parameter(Mandatory, ParameterSetName = 'LiteralPath')][string[]]$LiteralPath,
        [switch]$Raw,
        [string]$Encoding = 'utf8',
        [long]$TotalCount,
        [long]$Tail,
        [long]$ReadCount = 1,
        [string]$Delimiter = "`n",
        [switch]$AsByteStream,
        [switch]$Wait,
        [string]$Stream
    )
    process {
        $paths = if ($LiteralPath) { $LiteralPath } else { $Path }
        foreach ($inputPath in $paths) {
            if ([Management.Automation.WildcardPattern]::ContainsWildcardCharacters($inputPath)) {
                $native = @{ Path = $inputPath; Encoding = $Encoding }
                if ($Raw) { $native.Raw = $true }
                if ($PSBoundParameters.ContainsKey('TotalCount')) { $native.TotalCount = $TotalCount }
                if ($PSBoundParameters.ContainsKey('Tail')) { $native.Tail = $Tail }
                if ($ReadCount -ne 1) { $native.ReadCount = $ReadCount }
                if ($Delimiter -cne "`n") { $native.Delimiter = $Delimiter }
                if ($AsByteStream) { $native.AsByteStream = $true }
                if ($Wait) { $native.Wait = $true }
                if ($Stream) { $native.Stream = $Stream }
                Microsoft.PowerShell.Management\Get-Content @native
                continue
            }
            $target = [IO.Path]::GetFullPath($inputPath)
            $manifest = Get-SPJsonRecordManifest $target
            if ($manifest) {
                if ($PSBoundParameters.ContainsKey('TotalCount') -or $PSBoundParameters.ContainsKey('Tail') -or
                    $ReadCount -ne 1 -or $Delimiter -cne "`n" -or $AsByteStream -or $Wait -or $Stream) {
                    throw "Use default text or -Raw to read converted JSON report evidence: $target"
                }
                $jsonText = Read-SPJsonRecordReport $target $manifest
                if ($Raw) { $jsonText }
                else {
                    $reader = [IO.StringReader]::new($jsonText)
                    try { while ($null -ne ($line = $reader.ReadLine())) { $line } }
                    finally { $reader.Dispose() }
                }
                continue
            }
            $textParts = @(Get-SPTextSiblingParts $target)
            if ((Test-Path -LiteralPath $target -PathType Leaf) -or $textParts.Count) {
                $files = @()
                if (Test-Path -LiteralPath $target -PathType Leaf) { $files += Get-Item -LiteralPath $target }
                $files += $textParts
                if (($textParts.Count -or ($files.Count -and $files[0].FullName -ne $target)) -and
                    ($PSBoundParameters.ContainsKey('TotalCount') -or $PSBoundParameters.ContainsKey('Tail') -or
                        $ReadCount -ne 1 -or $Delimiter -cne "`n" -or $AsByteStream -or $Stream)) {
                    throw "Use default text or -Raw when reading a rotating output set: $target"
                }
                if ($files.Count -gt 1) {
                    if ($Wait) { throw "Cannot wait on a rotating output file set: $target" }
                    if ($Raw) {
                        $content = [Text.StringBuilder]::new()
                        foreach ($file in $files) {
                            [void]$content.Append((Microsoft.PowerShell.Management\Get-Content `
                                -LiteralPath $file.FullName -Raw -Encoding $Encoding))
                        }
                        $content.ToString()
                    } else {
                        foreach ($file in $files) {
                            Microsoft.PowerShell.Management\Get-Content -LiteralPath $file.FullName `
                                -Encoding $Encoding
                        }
                    }
                    continue
                }
                if ($files.Count -eq 1 -and $files[0].FullName -ne $target) {
                    Microsoft.PowerShell.Management\Get-Content -LiteralPath $files[0].FullName `
                        -Raw:$Raw -Encoding $Encoding
                    continue
                }
            }
            $native = @{ LiteralPath = $target; Encoding = $Encoding }
            if ($Raw) { $native.Raw = $true }
            if ($PSBoundParameters.ContainsKey('TotalCount')) { $native.TotalCount = $TotalCount }
            if ($PSBoundParameters.ContainsKey('Tail')) { $native.Tail = $Tail }
            if ($ReadCount -ne 1) { $native.ReadCount = $ReadCount }
            if ($Delimiter -cne "`n") { $native.Delimiter = $Delimiter }
            if ($AsByteStream) { $native.AsByteStream = $true }
            if ($Wait) { $native.Wait = $true }
            if ($Stream) { $native.Stream = $Stream }
            Microsoft.PowerShell.Management\Get-Content @native
            continue
        }
    }
}

function Add-Content {
    [CmdletBinding(DefaultParameterSetName = 'Path')]
    param(
        [Parameter(Mandatory, Position = 1, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [Alias('Value')][object]$InputObject,
        [Parameter(Position = 0)][string[]]$Path,
        [string[]]$LiteralPath,
        [string]$Encoding = 'utf8',
        [switch]$NoNewline,
        [switch]$PassThru
    )
    begin {
        $targets = if ($LiteralPath) { $LiteralPath } else { $Path }
        if ($targets.Count -ne 1) { throw 'Add-Content accepts one output path under repository limits.' }
        $target = Assert-SPOutputPath ([string]$targets)
        $codec = Get-SPOutputEncoding $Encoding
        $newline = if ($NoNewline) { '' } else { [Environment]::NewLine }
        $existing = @()
        if (Test-Path -LiteralPath $target -PathType Leaf) { $existing += Get-Item -LiteralPath $target }
        $existing += Get-SPTextSiblingParts $target
        if ($existing.Count) {
            $last = $existing[-1]
            $currentPath = $last.FullName
            $currentBytes = $last.Length
            if ($currentBytes -ge $script:SPOutputPartLimit) {
                $number = if ($currentPath -eq $target) { 2 } else {
                    [int]([regex]::Match($last.Name, '\.part(\d+)(?:\.|$)').Groups[1].Value) + 1
                }
                $currentPath = "$target.part{0:D4}{1}" -f $number, [IO.Path]::GetExtension($target)
                $currentPath = Assert-SPOutputPath $currentPath
                $currentBytes = 0L
            }
        } else {
            $currentPath = $target
            $currentBytes = 0L
        }
    }
    process {
        $value = [string]$InputObject
        $additionBytes = $codec.GetByteCount($value + $newline)
        if ($additionBytes -ge $script:SPOutputPartLimit) {
            throw "A single text record cannot fit below 90 MiB: $currentPath"
        }
        $preamble = if ($currentBytes -eq 0) { $codec.GetPreamble().Length } else { 0 }
        if ($currentBytes + $preamble + $additionBytes -ge $script:SPOutputPartLimit) {
            $textParts = @(Get-SPTextSiblingParts $target)
            $number = if ($textParts.Count) {
                [int]([regex]::Match($textParts[-1].Name, '\.part(\d+)(?:\.|$)').Groups[1].Value) + 1
            } else { 2 }
            $currentPath = "$target.part{0:D4}{1}" -f $number, [IO.Path]::GetExtension($target)
            $currentPath = Assert-SPOutputPath $currentPath
            $currentBytes = 0L
        }
        if (-not (Test-Path -LiteralPath $currentPath -PathType Leaf)) {
            $directory = [IO.Path]::GetDirectoryName($currentPath)
            if (-not [IO.Directory]::Exists($directory)) { [void][IO.Directory]::CreateDirectory($directory) }
        }
        $native = @{ LiteralPath = $currentPath; Encoding = $Encoding; NoNewline = $NoNewline; PassThru = $PassThru }
        Microsoft.PowerShell.Management\Add-Content -Value $InputObject @native
        $currentBytes += $preamble + $additionBytes
        if ($currentBytes -ge $script:SPOutputPartLimit -or (Get-Item -LiteralPath $currentPath).Length -ge $script:SPOutputPartLimit) {
            throw "Text output reached the 90 MiB safety threshold unexpectedly: $currentPath"
        }
    }
}

function Copy-Item {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, Position = 0)][Alias('LiteralPath')][string[]]$Path,
        [Parameter(Mandatory, Position = 1)][string]$Destination,
        [switch]$Force,
        [switch]$Recurse,
        [switch]$Container,
        [switch]$PassThru
    )
    $destinationPath = Assert-SPOutputPath $Destination
    foreach ($source in $Path) {
        $items = if ([Management.Automation.WildcardPattern]::ContainsWildcardCharacters($source)) {
            @(Get-Item -Path $source -ErrorAction Stop)
        } else { @(Get-Item -LiteralPath $source -ErrorAction Stop) }
        foreach ($item in $items) {
            $finalPath = if (Test-Path -LiteralPath $destinationPath -PathType Container) {
                Join-Path $destinationPath $item.Name
            } else { $destinationPath }
            [void](Assert-SPOutputPath $finalPath)
            if ($item.PSIsContainer -and $Recurse) {
                foreach ($child in Get-ChildItem -LiteralPath $item.FullName -Recurse -File) {
                    $relative = [IO.Path]::GetRelativePath($item.FullName, $child.FullName)
                    [void](Assert-SPOutputPath (Join-Path $finalPath $relative))
                    if ($child.Length -ge $script:SPOutputPartLimit) {
                        throw "Refusing to copy a file at or above 90 MiB; split or omit it: $($child.FullName)"
                    }
                }
            } elseif (-not $item.PSIsContainer -and $item.Length -ge $script:SPOutputPartLimit) {
                throw "Refusing to copy a file at or above 90 MiB; split or omit it: $($item.FullName)"
            }
        }
    }
    Microsoft.PowerShell.Management\Copy-Item -Path $Path -Destination $destinationPath `
        -Force:$Force -Recurse:$Recurse -Container:$Container -PassThru:$PassThru
}

function Move-Item {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, Position = 0)][Alias('LiteralPath')][string[]]$Path,
        [Parameter(Mandatory, Position = 1)][string]$Destination,
        [switch]$Force,
        [switch]$PassThru
    )
    $destinationPath = Assert-SPOutputPath $Destination
    foreach ($source in $Path) {
        $items = if ([Management.Automation.WildcardPattern]::ContainsWildcardCharacters($source)) {
            @(Get-Item -Path $source -ErrorAction Stop)
        } else { @(Get-Item -LiteralPath $source -ErrorAction Stop) }
        foreach ($item in $items) {
            $finalPath = if (Test-Path -LiteralPath $destinationPath -PathType Container) {
                Join-Path $destinationPath $item.Name
            } else { $destinationPath }
            [void](Assert-SPOutputPath $finalPath)
            if ($item.PSIsContainer) {
                foreach ($child in Get-ChildItem -LiteralPath $item.FullName -Recurse -File) {
                    $relative = [IO.Path]::GetRelativePath($item.FullName, $child.FullName)
                    [void](Assert-SPOutputPath (Join-Path $finalPath $relative))
                    if ($child.Length -ge $script:SPOutputPartLimit) {
                        throw "Refusing to move a file at or above 90 MiB into repository outputs: $($child.FullName)"
                    }
                }
            } elseif ($item.Length -ge $script:SPOutputPartLimit) {
                throw "Refusing to move a file at or above 90 MiB into repository outputs: $($item.FullName)"
            }
        }
    }
    Microsoft.PowerShell.Management\Move-Item -Path $Path -Destination $destinationPath `
        -Force:$Force -PassThru:$PassThru
}

function New-Item {
    [CmdletBinding()]
    param(
        [Parameter(Position = 0)][string[]]$Path,
        [Parameter(Position = 1)][string]$Name,
        [string]$ItemType,
        [object]$Value,
        [switch]$Force
    )
    foreach ($pathItem in $Path) {
        [void](Assert-SPOutputPath $pathItem)
        if ($Name) { [void](Assert-SPOutputPath (Join-Path $pathItem $Name)) }
    }
    Microsoft.PowerShell.Management\New-Item -Path $Path -Name $Name -ItemType $ItemType `
        -Value $Value -Force:$Force
}

function Export-Clixml {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ValueFromPipeline)][object]$InputObject,
        [Parameter(Mandatory, Position = 0)][Alias('LiteralPath')][string]$Path,
        [ValidateRange(1, 100)][int]$Depth = 2,
        [switch]$Force
    )
    begin { $items = [Collections.Generic.List[object]]::new() }
    process { $items.Add($InputObject) }
    end {
        $target = Assert-SPOutputPath $Path
        $object = if ($items.Count -eq 1) { $items[0] } else { $items.ToArray() }
        $xml = [Management.Automation.PSSerializer]::Serialize($object, $Depth)
        $bytes = [Text.UTF8Encoding]::new($true).GetBytes($xml)
        Assert-SPOutputSize -Path $target -Bytes $bytes.LongLength
        if ((Test-Path -LiteralPath $target) -and -not $Force) {
            throw "Output already exists; specify -Force to replace it: $target"
        }
        [IO.File]::WriteAllBytes($target, $bytes)
    }
}
