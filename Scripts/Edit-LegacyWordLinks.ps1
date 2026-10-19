<#
.SYNOPSIS
Inspects or edits a local legacy DOC using Word, without converting its format.
.DESCRIPTION
Called by Update-EmbeddedLinks.py - normally do not invoke directly.
Requires desktop Word and an interactive Windows user profile (not a service account).
Macros and automatic link updates are disabled. Only a temporary copy is edited.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$InputPath,
    [Parameter(Mandatory)][string]$OutputJson,
    [string]$OutputPath,
    [string]$ReplacementsJson
)
$ErrorActionPreference = 'Stop'

function Get-LinkSnapshot {
    param($Document)
    $records = [System.Collections.Generic.List[object]]::new()
    $visible = [System.Text.StringBuilder]::new()
    foreach ($story in $Document.StoryRanges) {
        $range = $story
        $chain = 0
        while ($null -ne $range) {
            $chain++
            [void]$visible.Append("$($range.StoryType):$chain`n")
            [void]$visible.Append([string]$range.Text)
            for ($i = 1; $i -le $range.Hyperlinks.Count; $i++) {
                $link = $range.Hyperlinks.Item($i)
                $records.Add([pscustomobject]@{
                    Location = "Word story $($range.StoryType), chain $chain, hyperlink $i"
                    Address = [string]$link.Address
                    SubAddress = [string]$link.SubAddress
                    Text = [string]$link.Range.Text
                })
            }
            $range = $range.NextStoryRange
        }
    }
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($visible.ToString())
    $hash = [Convert]::ToHexString([System.Security.Cryptography.SHA256]::HashData($bytes))
    return [pscustomobject]@{ TextSHA256 = $hash; Links = @($records.ToArray()) }
}

$word = $null
$document = $null
$oldUpdateLinks = $null
$oldSecurity = $null
try {
    $source = (Resolve-Path -LiteralPath $InputPath).Path
    if ([System.IO.Path]::GetExtension($source) -ine '.doc') {
        throw 'This helper only handles legacy .doc files.'
    }
    $working = $source
    if ($OutputPath) {
        if (-not $ReplacementsJson) { throw 'ReplacementsJson is required when editing.' }
        if (Test-Path -LiteralPath $OutputPath) { throw 'OutputPath must not already exist.' }
        Copy-Item -LiteralPath $source -Destination $OutputPath
        $working = [System.IO.Path]::GetFullPath($OutputPath)
    }
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $oldSecurity = $word.AutomationSecurity
    $word.AutomationSecurity = 3
    $oldUpdateLinks = $word.Options.UpdateLinksAtOpen
    $word.Options.UpdateLinksAtOpen = $false
    $document = $word.Documents.Open($working, $false, [bool](-not $OutputPath), $false)
    $before = Get-LinkSnapshot -Document $document
    if ($OutputPath) {
        if ($document.ProtectionType -ne -1) { throw 'Protected Word document; no edit attempted.' }
        if ($document.Signatures.Count -gt 0) { throw 'Signed Word document; no edit attempted.' }
        $replacements = @(Get-Content -LiteralPath $ReplacementsJson -Raw | ConvertFrom-Json)
        $byLocation = @{}
        foreach ($replacement in $replacements) {
            if ($byLocation.ContainsKey($replacement.Location)) { throw 'Duplicate replacement location.' }
            $byLocation[$replacement.Location] = $replacement
        }
        $changed = 0
        foreach ($story in $document.StoryRanges) {
            $range = $story
            $chain = 0
            while ($null -ne $range) {
                $chain++
                for ($i = 1; $i -le $range.Hyperlinks.Count; $i++) {
                    $location = "Word story $($range.StoryType), chain $chain, hyperlink $i"
                    if ($byLocation.ContainsKey($location)) {
                        $replacement = $byLocation[$location]
                        $link = $range.Hyperlinks.Item($i)
                        if ([string]$link.Address -cne $replacement.OldURL) {
                            throw "Hyperlink address changed before editing: $location"
                        }
                        $link.Address = $replacement.NewURL
                        $changed++
                    }
                }
                $range = $range.NextStoryRange
            }
        }
        if ($changed -ne $replacements.Count) { throw 'Not all replacements were applied.' }
        $after = Get-LinkSnapshot -Document $document
        if ($before.TextSHA256 -cne $after.TextSHA256) { throw 'Visible Word text changed.' }
        $document.Save()
        $document.Close(0)
        $document = $word.Documents.Open($working, $false, $true, $false)
        $after = Get-LinkSnapshot -Document $document
        if ($before.TextSHA256 -cne $after.TextSHA256) { throw 'Visible Word text changed after saving.' }
        if ($before.Links.Count -ne $after.Links.Count) { throw 'Word hyperlink count changed.' }
        for ($i = 0; $i -lt $before.Links.Count; $i++) {
            $a, $b = $before.Links[$i], $after.Links[$i]
            $expected = if ($byLocation.ContainsKey($a.Location)) {
                $byLocation[$a.Location].NewURL
            } else { $a.Address }
            if ($b.Address -cne $expected -or $a.SubAddress -cne $b.SubAddress -or $a.Text -cne $b.Text) {
                throw "Saved Word hyperlink verification failed: $($a.Location)"
            }
        }
        $before = $after
    }
    $before | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $OutputJson -Encoding utf8
}
finally {
    if ($null -ne $document) { $document.Close(0) }
    if ($null -ne $word) {
        if ($null -ne $oldUpdateLinks) { $word.Options.UpdateLinksAtOpen = $oldUpdateLinks }
        if ($null -ne $oldSecurity) { $word.AutomationSecurity = $oldSecurity }
        $word.Quit()
        [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($word)
    }
}
