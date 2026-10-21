#requires -Version 7.4.6
[CmdletBinding()]
param()
$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent $PSScriptRoot
. (Join-Path $repo 'Scripts\Convert-SPServerReports.ps1')

function Assert-Test {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw $Message }
}

function New-TestFolder {
    param([string]$Name)
    $path = Join-Path $temp $Name
    [void][IO.Directory]::CreateDirectory($path)
    return $path
}

$temp = Join-Path ([IO.Path]::GetTempPath()) ('sp-record-tests-' + [guid]::NewGuid().ToString('N'))
[void][IO.Directory]::CreateDirectory($temp)
try {
    # Actual selection boundary, without allocating a 99 MiB JSON object in memory.
    $boundary = New-TestFolder 'boundary'
    $jsonPath = Join-Path $boundary 'Plan.json'
    $stream = [IO.File]::Create($jsonPath)
    $chunk = [Text.Encoding]::UTF8.GetBytes(' ' * 1MB)
    try {
        $stream.Write([Text.Encoding]::UTF8.GetBytes('{"Steps":[],"Notes":[]}'))
        while ($stream.Length -lt 99MB) {
            $stream.Write($chunk, 0, [int][math]::Min($chunk.Length, 99MB - $stream.Length))
        }
    } finally { $stream.Dispose() }
    $small = Join-Path $boundary 'small.json'
    [IO.File]::WriteAllText($small, '{"keep":true}')
    $preview = @(Convert-SPServerReports -Folder $boundary)
    Assert-Test ($preview.Count -eq 1 -and (Get-Item $jsonPath).Length -eq 99MB) 'Preview modified source or missed exact 99 MiB boundary.'
    Convert-SPServerReports -Folder $boundary -Apply -WhatIf | Out-Null
    Assert-Test ((Get-Item $jsonPath).Length -eq 99MB) 'WhatIf modified the original.'
    Convert-SPServerReports -Folder $boundary -Apply | Out-Null
    $plan = Get-Content -LiteralPath $jsonPath -Raw | ConvertFrom-Json
    Assert-Test ($plan.Steps.Count -eq 0 -and $plan.Notes.Count -eq 0) 'Actual-threshold JSON round trip failed.'
    Assert-Test ([IO.File]::ReadAllText($small) -ceq '{"keep":true}') 'Small JSON was modified.'
    $csvPath = Join-Path $boundary 'Large.csv'
    $writer = [IO.StreamWriter]::new($csvPath, $false, [Text.UTF8Encoding]::new($false))
    $largeField = 'x' * 1MB
    try {
        $writer.WriteLine('"Id","Value"')
        for ($i = 0; $i -lt 100; $i++) { $writer.WriteLine(('"{0}","{1}"' -f $i, $largeField)) }
    } finally { $writer.Dispose() }
    Convert-SPServerReports -Folder $boundary -Apply | Out-Null
    $largeParts = @((Get-Item -LiteralPath $csvPath)) + @(Get-SPCsvSiblingParts $csvPath)
    Assert-Test ($largeParts.Count -eq 2) 'Actual >100 MiB CSV did not split into two parts.'
    foreach ($file in $largeParts) { Assert-Test ($file.Length -lt 90MB) 'Actual CSV part reached the 90 MiB safety threshold.' }
    $largeCount = 0
    Import-Csv -LiteralPath $csvPath | ForEach-Object {
        Assert-Test ($_.Id -ceq [string]$largeCount -and $_.Value.Length -eq 1MB) 'Actual-size CSV content/order changed.'
        $largeCount++
    }
    Assert-Test ($largeCount -eq 100) 'Actual-size CSV record count changed.'

    # Small injected limits exercise rollover without generating hundreds of MiB per test.
    $definition = (Get-Command Convert-SPServerReports).Definition.Replace('$threshold = 99MB', '$threshold = 2KB')
    Set-Item -Path Function:\Convert-SPServerReports -Value ([scriptblock]::Create($definition))
    $script:SPOutputPartLimit = 1KB
    $jsonFolder = New-TestFolder 'json'
    $jsonPath = Join-Path $jsonFolder 'Plan.json'
    $original = [ordered]@{
        Steps = @(1..40 | ForEach-Object {
            [ordered]@{ Key = "scope-$_"; Enabled = $true; Amount = [decimal]'123.456789'; LargeId = [long]::MaxValue }
        })
        Notes = @(); Object = @{}; Missing = $null
        Text = "line1`nline2`t`"quoted`" " + [char]0x00E9
    }
    [IO.File]::WriteAllText($jsonPath, (ConvertTo-Json -InputObject $original -Depth 20))
    Convert-SPServerReports -Folder $jsonFolder -Apply | Out-Null
    $manifest = Get-SPJsonRecordManifest $jsonPath
    $recordsPath = Join-Path $jsonFolder $manifest.Records
    Assert-Test (@(Get-SPCsvSiblingParts $recordsPath).Count -gt 1) 'JSON did not create multiple CSV parts.'
    $restored = Get-Content -LiteralPath $jsonPath -Raw | ConvertFrom-Json -Depth 100
    Assert-Test ($restored.Steps.Count -eq 40 -and $restored.Steps[-1].Key -ceq 'scope-40' -and
        $restored.Steps[0].LargeId -eq [long]::MaxValue -and $restored.Text -ceq $original.Text -and
        $restored.Notes.Count -eq 0 -and $null -eq $restored.Missing) 'JSON values/order/types did not survive conversion.'
    foreach ($file in Get-ChildItem -LiteralPath $jsonFolder -File) {
        Assert-Test ($file.Length -lt 1KB) "Injected file-size cap exceeded: $($file.FullName)"
    }
    $part = @(Get-SPCsvSiblingParts $recordsPath)[0]
    Remove-Item -LiteralPath $part.FullName
    $failed = $false
    try { Get-Content -LiteralPath $jsonPath -Raw | Out-Null } catch { $failed = $true }
    Assert-Test $failed 'Missing JSON record part was silently accepted.'

    $csvFolder = New-TestFolder 'csv'
    $csvPath = Join-Path $csvFolder 'EffectivePermissions.csv'
    $rows = @(1..40 | ForEach-Object { [pscustomobject]@{ Id = "$_"; Text = "line1`nline2,`"quoted`" " + ('x' * 70) } })
    $rows | Microsoft.PowerShell.Utility\Export-Csv -LiteralPath $csvPath -NoTypeInformation -Encoding utf8
    Convert-SPServerReports -Folder $csvFolder -Apply | Out-Null
    $actual = @(Import-Csv -LiteralPath $csvPath)
    Assert-Test ($actual.Count -eq 40 -and $actual[0].Text -ceq $rows[0].Text -and
        $actual[-1].Id -ceq '40') 'CSV multiline fields/order/count changed.'
    Assert-Test (@(Get-SPCsvSiblingParts $csvPath).Count -gt 0) 'CSV did not roll over.'
    foreach ($file in Get-ChildItem -LiteralPath $csvFolder -File) {
        Assert-Test ($file.Length -lt 1KB) 'CSV part exceeded injected size cap.'
    }
    $precisionFolder = New-TestFolder 'precision'
    $precisionPath = Join-Path $precisionFolder 'numbers.json'
    $numbers = '{"fraction":0.123456789012345678901234567890123456789,"huge":1e400,"tiny":1e-400}'
    [IO.File]::WriteAllText($precisionPath, $numbers + (' ' * 2500))
    Convert-SPServerReports -Folder $precisionFolder -Apply | Out-Null
    $precisionText = Get-Content -LiteralPath $precisionPath -Raw
    Assert-Test ($precisionText -ceq $numbers) 'Numeric precision was changed by JSON conversion.'

    foreach ($case in @('malformed', 'large-scalar', 'collision', 'unsupported')) {
        $folder = New-TestFolder $case
        $path = Join-Path $folder $(if ($case -eq 'unsupported') { 'file.bin' } else { 'file.json' })
        $content = switch ($case) {
            'malformed' { '{"broken":[' + (' ' * 2500) }
            'large-scalar' { '{"value":"' + ('x' * 2500) + '"}' }
            'collision' { '{"value":true}' + (' ' * 2500) }
            'unsupported' { 'x' * 2500 }
        }
        [IO.File]::WriteAllText($path, $content)
        if ($case -eq 'collision') {
            $id = [Convert]::ToHexString([Security.Cryptography.SHA256]::HashData(
                [Text.Encoding]::UTF8.GetBytes($path)
            )).Substring(0, 12).ToLowerInvariant()
            [IO.File]::WriteAllText((Join-Path $folder "j-$id.csv"), 'existing')
        }
        $failed = $false
        try { Convert-SPServerReports -Folder $folder -Apply -WarningAction SilentlyContinue | Out-Null }
        catch { $failed = $true }
        Assert-Test $failed "Expected failure was not reported for $case."
        Assert-Test ([IO.File]::ReadAllText($path) -ceq $content) "Original modified on $case failure."
    }

    # Legacy permission preflight still sees blocked notes through the indexed JSON reader.
    . (Join-Path $repo 'Scripts\Invoke-SPPermissionChecklist.ps1') -ChecklistFile (Join-Path $temp 'unused.csv')
    $script:SPOutputPartLimit = 1KB
    $legacy = New-TestFolder 'legacy'
    $path = Join-Path $legacy 'PermSync_Plan_DRYRUN_test.json'
    [IO.File]::WriteAllText($path, '{"Steps":[],"Notes":[{"Result":"Blocked"}]}' + (' ' * 2500))
    [IO.File]::WriteAllText((Join-Path $legacy 'PermSync_UnavailablePrincipals_DRYRUN_test.csv'), '"Principal"')
    Convert-SPServerReports -Folder $legacy -Apply | Out-Null
    $failed = $false
    try { Assert-SPChecklistDryRun $legacy } catch { $failed = $_.Exception.Message -like '*blocked scopes*' }
    Assert-Test $failed 'Legacy preflight did not reject blocked notes after JSON conversion.'
    $tooLong = Join-Path $temp ('x' * 260)
    $failed = $false
    try { Assert-SPOutputPath $tooLong | Out-Null } catch { $failed = $true }
    Assert-Test $failed 'Overlong output path was accepted.'
    'PASS: selection boundary, preview/WhatIf, indexed JSON, CSV rollover, failure preservation, and legacy blocked-note checks.'
} finally {
    Remove-Item -LiteralPath $temp -Recurse -Force
}
