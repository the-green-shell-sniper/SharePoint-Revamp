#requires -Version 7.4.6
<#
.SYNOPSIS
    Read-only proposals for renamed lists, folders, files and list items.
.DESCRIPTION
    Every proposal starts with Approved=No, including High confidence proposals.
    Review each row and approve parent renames AND their proposed children explicitly.
    Exact site-relative paths are reserved before fuzzy matching. Files are compared
    only within a matched library and with the same extension. No SharePoint writes
    are performed. Content downloads remain in memory; errors appear in evidence.
    Copy SPPermission.Common.ps1 beside this script when running on another machine.
.EXAMPLE
    .\New-SPPermissionMatchMap.ps1 -Tenant '<tenant-guid-or-domain>' -OutputFolder .\PermissionRuns
#>
[CmdletBinding()]
param(
    [string]$SourceUrl = 'https://whallc.sharepoint.com/sites/archive-2026-10-18T115436Z/CS',
    [string]$TargetUrl = 'https://whallc.sharepoint.com/sites/ClinicalandPatientServices',
    [string]$ClientId = '2e20d82f-56ba-418e-b465-9718493f0df8',
    [string]$Tenant,
    [string]$OutputFolder = (Join-Path $PSScriptRoot '..\Outputs\PermissionRuns'),
    [ValidateRange(0, 1)][double]$HighThreshold = 0.85,
    [ValidateRange(0, 1)][double]$ReviewThreshold = 0.55,
    [ValidateRange(0, 1)][double]$MinMargin = 0.10,
    [ValidateRange(1, 1024)][int]$MaxContentMB = 25,
    [ValidateRange(1, 5000)][int]$PageSize = 500,
    [string[]]$ExcludePaths = @('Access Requests'),
    [switch]$NoContent,
    [switch]$SkipPersistLoginSetup
)
. (Join-Path $PSScriptRoot 'SPPermission.Common.ps1')

function Get-SPNameRatio {
    param([string]$Left, [string]$Right)
    if ($Left -ceq $Right) { return 1.0 }
    $length = [math]::Max($Left.Length, $Right.Length)
    if (-not $Left -or -not $Right) { return 0.0 }
    $previous = [int[]](0..$Right.Length)
    $current = [int[]]::new($Right.Length + 1)
    for ($i = 1; $i -le $Left.Length; $i++) {
        $current[0] = $i
        for ($j = 1; $j -le $Right.Length; $j++) {
            $cost = if ($Left[$i - 1] -ceq $Right[$j - 1]) { 0 } else { 1 }
            $current[$j] = [math]::Min([math]::Min($current[$j - 1] + 1, $previous[$j] + 1), $previous[$j - 1] + $cost)
        }
        $swap = $previous; $previous = $current; $current = $swap
    }
    1.0 - $previous[$Right.Length] / $length
}

function Get-SPNormalizedName {
    param([string]$Name)
    $nameOnly = [IO.Path]::GetFileNameWithoutExtension($Name).ToLowerInvariant()
    $nameOnly = $nameOnly -replace '\(\d+\)', ' ' -replace '\b(copy|final|rev(ised)?|v\d+(\.\d+)*|version\s*\d+)\b', ' '
    (($nameOnly -replace '[^\p{L}\p{Nd}]+', ' ') -replace '\s+', ' ').Trim()
}

function Get-SPNameSimilarity {
    param([string]$Left, [string]$Right)
    $leftName = Get-SPNormalizedName $Left; $rightName = Get-SPNormalizedName $Right
    if (-not $leftName -or -not $rightName) { return 0.0 }
    $leftSorted = (($leftName -split ' ') | Sort-Object) -join ' '
    $rightSorted = (($rightName -split ' ') | Sort-Object) -join ' '
    [math]::Max((Get-SPNameRatio $leftName $rightName), (Get-SPNameRatio $leftSorted $rightSorted))
}

function Get-SPContentFingerprint {
    param($Connection, $Item, [int]$MaxMB)
    $stream = $null; $zip = $null; $zipStream = $null
    try {
        if ($null -eq $Item.Size) { throw 'Size unknown; content download skipped.' }
        if ($Item.Size -gt $MaxMB * 1MB) { throw "Larger than $MaxMB MB; content download skipped." }
        $stream = Get-PnPFile -Url $Item.FileRef -AsMemoryStream -Connection $Connection -ErrorAction Stop
        if ($stream.Length -gt $MaxMB * 1MB) { throw 'Downloaded content exceeds size limit.' }
        $bytes = $stream.ToArray()
        $hash = [Convert]::ToHexString([Security.Cryptography.SHA256]::HashData($bytes))
        $text = ''
        if ($Item.Ext -in @('.docx', '.xlsx', '.pptx')) {
            $pattern = switch ($Item.Ext) {
                '.docx' { '^word/(document|header\d*|footer\d*)\.xml$' }
                '.xlsx' { '^xl/(sharedStrings|worksheets/sheet\d+)\.xml$' }
                '.pptx' { '^ppt/slides/slide\d+\.xml$' }
            }
            $zipStream = [IO.MemoryStream]::new($bytes)
            $zip = [IO.Compression.ZipArchive]::new($zipStream, [IO.Compression.ZipArchiveMode]::Read)
            $builder = [Text.StringBuilder]::new()
            $expanded = 0L
            foreach ($entry in $zip.Entries | Where-Object FullName -match $pattern) {
                $expanded += $entry.Length
                if ($expanded -gt 50MB) { throw 'Expanded Office text exceeds 50 MB limit.' }
                $reader = [IO.StreamReader]::new($entry.Open())
                try { [void]$builder.Append(' ').Append($reader.ReadToEnd()) } finally { $reader.Dispose() }
            }
            $text = $builder.ToString()
        } elseif ($Item.Ext -in @('.txt', '.csv', '.htm', '.html', '.md')) {
            $text = [Text.Encoding]::UTF8.GetString($bytes)
        }
        $words = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        $text = [System.Net.WebUtility]::HtmlDecode(($text -replace '<[^>]+>', ' '))
        foreach ($word in ($text.ToLowerInvariant() -split '[^\p{L}\p{Nd}]+')) {
            if ($word.Length -ge 3) { [void]$words.Add($word) }
        }
        [pscustomobject]@{ Hash = $hash; Words = $words; Error = '' }
    } catch {
        Write-Warning "Content comparison unavailable for '$($Item.Key)': $($_.Exception.Message)"
        [pscustomobject]@{ Hash = ''; Words = $null; Error = $_.Exception.Message }
    } finally {
        if ($zip) { $zip.Dispose() }
        if ($zipStream) { $zipStream.Dispose() }
        if ($stream) { $stream.Dispose() }
    }
}

function Get-SPMatchScore {
    param($SourceItem, $TargetItem, $SourceFingerprint, $TargetFingerprint)
    $nameA = if ($SourceItem.Kind -eq 'Item' -and $SourceItem.Title) { $SourceItem.Title } else { $SourceItem.Leaf }
    $nameB = if ($TargetItem.Kind -eq 'Item' -and $TargetItem.Title) { $TargetItem.Title } else { $TargetItem.Leaf }
    $nameScore = Get-SPNameSimilarity $nameA $nameB
    $sameCreated = ($SourceItem.Created -and $SourceItem.Created -eq $TargetItem.Created -and
        $SourceItem.Author -and $SourceItem.Author -ieq $TargetItem.Author)
    $parentA = ($SourceItem.Rel -replace '[^/]+$', '').ToLowerInvariant()
    $parentB = ($TargetItem.Rel -replace '[^/]+$', '').ToLowerInvariant()
    $weighted = 0.35 * $nameScore + 0.20 * [int][bool]$sameCreated + 0.10 * (Get-SPNameRatio $parentA $parentB)
    $weight = 0.65
    $evidence = [Collections.Generic.List[string]]::new()
    $evidence.Add("Name=$([math]::Round($nameScore, 3)); Created+Author=$([bool]$sameCreated)")
    if ($null -ne $SourceItem.Size -and $null -ne $TargetItem.Size) {
        $max = [math]::Max($SourceItem.Size, $TargetItem.Size)
        $sizeScore = if ($max) { 1 - [math]::Abs($SourceItem.Size - $TargetItem.Size) / $max } else { 1 }
        $weighted += 0.05 * $sizeScore; $weight += 0.05
    }
    $identical = $false
    if ($SourceFingerprint -and $TargetFingerprint -and $SourceFingerprint.Hash -and $TargetFingerprint.Hash) {
        $identical = $SourceFingerprint.Hash -eq $TargetFingerprint.Hash
        if ($identical) { $evidence.Add('Identical bytes'); $weighted += 0.30; $weight += 0.30 }
        elseif ($SourceFingerprint.Words.Count -and $TargetFingerprint.Words.Count) {
            $intersection = 0
            foreach ($word in $SourceFingerprint.Words) { if ($TargetFingerprint.Words.Contains($word)) { $intersection++ } }
            $similarity = $intersection / ($SourceFingerprint.Words.Count + $TargetFingerprint.Words.Count - $intersection)
            $weighted += 0.30 * $similarity; $weight += 0.30
            $evidence.Add("Text Jaccard=$([math]::Round($similarity, 3))")
        } else { $evidence.Add('Different bytes; no comparable text') }
    } elseif ($SourceItem.Kind -eq 'File') {
        $evidence.Add("Content not compared: $($SourceFingerprint.Error) $($TargetFingerprint.Error)")
    }
    $score = $weighted / $weight
    if ($identical) { $score = [math]::Max(0.98, $score) }
    [pscustomobject]@{ Score = [math]::Round($score, 3); Identical = $identical; Evidence = $evidence -join ' | ' }
}

function Invoke-NewSPPermissionMatchMap {
    $ErrorActionPreference = 'Stop'
    Assert-SPPermissionUrls $SourceUrl $TargetUrl
    $SourceUrl = $SourceUrl.TrimEnd('/')
    $TargetUrl = $TargetUrl.TrimEnd('/')
    if ($ReviewThreshold -gt $HighThreshold) { throw 'ReviewThreshold must not exceed HighThreshold.' }
    Import-Module PnP.PowerShell -MinimumVersion 3.3.0 -ErrorAction Stop
    [void][IO.Directory]::CreateDirectory($OutputFolder)
    $stamp = Get-Date -Format 'yyyyMMdd-HHmmss-fff'
    $sourceConnection = Connect-SPPermissionSite $SourceUrl $ClientId $Tenant -SkipPersistLoginSetup:$SkipPersistLoginSetup
    $targetConnection = Connect-SPPermissionSite $TargetUrl $ClientId $Tenant -SkipPersistLoginSetup:$SkipPersistLoginSetup
    $sourceInventory = Get-SPPermissionInventory $sourceConnection $SourceUrl $PageSize
    $targetInventory = Get-SPPermissionInventory $targetConnection $TargetUrl $PageSize
    $exact = Get-SPPermissionMatches $sourceInventory $targetInventory '' $SourceUrl $TargetUrl
    $usedTargets = [Collections.Generic.HashSet[string]]::new([string[]]@($exact.Values), [StringComparer]::OrdinalIgnoreCase)
    $pairedSources = [Collections.Generic.HashSet[string]]::new([string[]]@($exact.Keys), [StringComparer]::OrdinalIgnoreCase)
    $prefixes = [Collections.Generic.Dictionary[string, string]]::new([StringComparer]::OrdinalIgnoreCase)
    $proposals = [Collections.Generic.List[object]]::new()
    $unmatched = [Collections.Generic.List[object]]::new()
    $fingerprints = @{}
    $listPairs = [Collections.Generic.List[object]]::new()
    foreach ($sourceList in @($sourceInventory.Values | Where-Object { $_.Kind -in @('List', 'Library') } | Sort-Object Key)) {
        if (Test-SPPermissionExcluded $sourceList.Key $ExcludePaths) { continue }
        if ($exact.ContainsKey($sourceList.Key)) {
            $listPairs.Add(@($sourceList, $targetInventory[$sourceList.Key]))
            continue
        }
        $candidates = @($targetInventory.Values | Where-Object {
            $_.Kind -eq $sourceList.Kind -and $_.BaseTemplate -eq $sourceList.BaseTemplate -and
            -not $usedTargets.Contains($_.Key) -and -not (Test-SPPermissionExcluded $_.Key $ExcludePaths)
        } | ForEach-Object {
            $name = [math]::Max((Get-SPNameSimilarity $sourceList.Title $_.Title), (Get-SPNameSimilarity $sourceList.Leaf $_.Leaf))
            $max = [math]::Max($sourceList.ItemCount, $_.ItemCount)
            $count = if ($max) { 1 - [math]::Abs($sourceList.ItemCount - $_.ItemCount) / $max } else { 1 }
            [pscustomobject]@{ Target = $_; Score = [math]::Round(0.8 * $name + 0.2 * $count, 3) }
        } | Sort-Object @{ Expression = 'Score'; Descending = $true }, @{ Expression = { $_.Target.Key } })
        if (-not $candidates.Count -or $candidates[0].Score -lt $ReviewThreshold) { continue }
        $best = $candidates[0]
        $margin = if ($candidates.Count -gt 1) { $best.Score - $candidates[1].Score } else { 1 }
        $tier = if ($best.Score -ge $HighThreshold -and $margin -ge $MinMargin) { 'High' } else { 'Medium' }
        $proposals.Add([pscustomobject]@{
            Approved = 'No'; Confidence = $tier; Score = $best.Score; Kind = $sourceList.Kind
            PermissionRelevant = 'Yes'; SourceUnique = $sourceList.Unique; TargetUnique = $best.Target.Unique
            SourcePath = $sourceList.Key; TargetPath = $best.Target.Key; Method = 'Fuzzy list'
            Evidence = 'Title/URL similarity and item count; review list identity'
            RunnerUp = $(if ($candidates.Count -gt 1) { "$($candidates[1].Target.Key) ($($candidates[1].Score))" } else { '' })
            SourceUrl = $SourceUrl; TargetUrl = $TargetUrl
        })
        [void]$usedTargets.Add($best.Target.Key); [void]$pairedSources.Add($sourceList.Key)
        $prefixes[$sourceList.Key] = $best.Target.Key
        $listPairs.Add(@($sourceList, $best.Target))
    }
    foreach ($listPair in $listPairs) {
        $sourceList = $listPair[0]; $targetList = $listPair[1]
        $sourceItems = @($sourceInventory.Values | Where-Object {
            $_.Kind -in @('Folder', 'File', 'Item') -and $_.ListKey -ieq $sourceList.Key -and
            -not $pairedSources.Contains($_.Key) -and -not (Test-SPPermissionExcluded $_.Key $ExcludePaths)
        })
        foreach ($depth in @($sourceItems.Depth | Sort-Object -Unique)) {
            $scored = [Collections.Generic.List[object]]::new()
            foreach ($sourceItem in @($sourceItems | Where-Object Depth -eq $depth | Sort-Object Key)) {
                $translated = $sourceItem.Key
                $parentPrefix = @($prefixes.Keys | Where-Object {
                    Test-SPPermissionPathWithin $sourceItem.Key $_
                } | Sort-Object Length -Descending | Select-Object -First 1)
                if ($parentPrefix.Count) { $translated = $prefixes[$parentPrefix[0]] + $sourceItem.Key.Substring($parentPrefix[0].Length) }
                foreach ($targetItem in @($targetInventory.Values | Where-Object {
                    $_.ListKey -ieq $targetList.Key -and $_.Kind -eq $sourceItem.Kind -and
                    -not $usedTargets.Contains($_.Key) -and -not (Test-SPPermissionExcluded $_.Key $ExcludePaths) -and
                    ($sourceItem.Kind -ne 'File' -or $_.Ext -ieq $sourceItem.Ext)
                })) {
                    $method = 'Fuzzy'
                    if ($translated -ieq $targetItem.Key -and $translated -ine $sourceItem.Key) {
                        $score = [pscustomobject]@{ Score = 1.0; Identical = $false; Evidence = "Exact path under proposed parent; parent must also be approved: $($parentPrefix[0])" }
                        $method = 'Translated parent'
                    } else {
                        $nameA = if ($sourceItem.Kind -eq 'Item' -and $sourceItem.Title) { $sourceItem.Title } else { $sourceItem.Leaf }
                        $nameB = if ($targetItem.Kind -eq 'Item' -and $targetItem.Title) { $targetItem.Title } else { $targetItem.Leaf }
                        $sameCreated = ($sourceItem.Created -and $sourceItem.Created -eq $targetItem.Created -and
                            $sourceItem.Author -and $sourceItem.Author -ieq $targetItem.Author)
                        $sameSize = ($null -ne $sourceItem.Size -and $null -ne $targetItem.Size -and $sourceItem.Size -eq $targetItem.Size)
                        if ((Get-SPNameSimilarity $nameA $nameB) -lt 0.4 -and -not $sameCreated -and -not $sameSize) { continue }
                        $sourceFingerprint = $null; $targetFingerprint = $null
                        if ($sourceItem.Kind -eq 'File' -and -not $NoContent) {
                            foreach ($side in @('Source', 'Target')) {
                                $item = if ($side -eq 'Source') { $sourceItem } else { $targetItem }
                                $connection = if ($side -eq 'Source') { $sourceConnection } else { $targetConnection }
                                $cacheKey = "$side|$($item.Key)"
                                if (-not $fingerprints.ContainsKey($cacheKey)) {
                                    $fingerprints[$cacheKey] = Get-SPContentFingerprint $connection $item $MaxContentMB
                                }
                            }
                            $sourceFingerprint = $fingerprints["Source|$($sourceItem.Key)"]
                            $targetFingerprint = $fingerprints["Target|$($targetItem.Key)"]
                        }
                        $score = Get-SPMatchScore $sourceItem $targetItem $sourceFingerprint $targetFingerprint
                    }
                    $scored.Add([pscustomobject]@{
                        Source = $sourceItem; Target = $targetItem; Score = $score.Score
                        Identical = $score.Identical; Evidence = $score.Evidence; Method = $method
                    })
                }
            }
            foreach ($candidate in @($scored | Sort-Object @{ Expression = 'Score'; Descending = $true },
                @{ Expression = { $_.Source.Key } }, @{ Expression = { $_.Target.Key } })) {
                if ($candidate.Score -lt $ReviewThreshold -or $pairedSources.Contains($candidate.Source.Key) -or $usedTargets.Contains($candidate.Target.Key)) { continue }
                $others = @($scored | Where-Object {
                    ($_.Source.Key -ieq $candidate.Source.Key -and $_.Target.Key -ine $candidate.Target.Key) -or
                    ($_.Target.Key -ieq $candidate.Target.Key -and $_.Source.Key -ine $candidate.Source.Key)
                } | Sort-Object Score -Descending)
                $margin = if ($others.Count) { $candidate.Score - $others[0].Score } else { 1 }
                $tier = if ($candidate.Score -ge $HighThreshold -and $margin -ge $MinMargin) { 'High' } else { 'Medium' }
                $proposals.Add([pscustomobject]@{
                    Approved = 'No'; Confidence = $tier; Score = $candidate.Score; Kind = $candidate.Source.Kind
                    PermissionRelevant = $(if ($candidate.Source.Unique -or $candidate.Target.Unique -or $candidate.Source.Kind -eq 'Folder') { 'Yes' } else { 'No' })
                    SourceUnique = $candidate.Source.Unique; TargetUnique = $candidate.Target.Unique
                    SourcePath = $candidate.Source.Key; TargetPath = $candidate.Target.Key; Method = $candidate.Method
                    Evidence = $candidate.Evidence
                    RunnerUp = $(if ($others.Count) { "$($others[0].Source.Key) -> $($others[0].Target.Key) ($($others[0].Score))" } else { '' })
                    SourceUrl = $SourceUrl; TargetUrl = $TargetUrl
                })
                [void]$pairedSources.Add($candidate.Source.Key); [void]$usedTargets.Add($candidate.Target.Key)
                if ($candidate.Source.Kind -eq 'Folder') { $prefixes[$candidate.Source.Key] = $candidate.Target.Key }
            }
        }
    }
    foreach ($side in @('Source', 'Target')) {
        $inventory = if ($side -eq 'Source') { $sourceInventory } else { $targetInventory }
        $paired = if ($side -eq 'Source') { $pairedSources } else { $usedTargets }
        foreach ($item in $inventory.Values | Sort-Object Key) {
            if ($item.Key -and -not $paired.Contains($item.Key) -and -not (Test-SPPermissionExcluded $item.Key $ExcludePaths)) {
                $unmatched.Add([pscustomobject]@{ Side = $side; Kind = $item.Kind; Path = $item.Key; Unique = $item.Unique })
            }
        }
    }
    $reviewPath = Join-Path $OutputFolder "PermMatch_Review_$stamp.csv"
    Export-SPPermissionCsv $proposals.ToArray() @('Approved','Confidence','Score','Kind','PermissionRelevant','SourceUnique','TargetUnique','SourcePath','TargetPath','Method','Evidence','RunnerUp','SourceUrl','TargetUrl') $reviewPath
    Export-SPPermissionCsv $unmatched.ToArray() @('Side','Kind','Path','Unique') (Join-Path $OutputFolder "PermMatch_Unmatched_$stamp.csv")
    Write-Host "Proposed pairs: $($proposals.Count); unmatched objects: $($unmatched.Count). ALL proposals require approval."
    Write-Host "Review: $reviewPath"
}

if ($MyInvocation.InvocationName -ne '.') { Invoke-NewSPPermissionMatchMap }
