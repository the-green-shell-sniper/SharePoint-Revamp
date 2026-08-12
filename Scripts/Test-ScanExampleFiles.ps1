# ─── Test harness for URLsv3 scan logic ──────────────────────────────────────
# Loads function definitions directly from the actual Get-FixItListURLsv3.ps1
# so any change to that script is reflected here automatically.
# Only thing to configure: the URL mapping that matches the intended real run.

$baseDir   = "c:\Users\bposey\OneDrive - Womens Healthcare Associates LLC\Gary Henderson's files - Project Files_bposey\GitHub\SharePoint-Revamp"
$scanDir   = Join-Path $baseDir "ExampleFilesForScanning"
$sourceCsv = Join-Path $baseDir "Reports\PatientFinancialServices-Get-FixItListURLs-Documents-20260810-181633.csv"
$outputCsv = Join-Path $baseDir "Reports\ExampleFilesForScanning-TestOutput-$(Get-Date -Format 'yyyyMMdd-HHmmss').csv"

# ─── URL mapping — must match what the real run would use ────────────────────
$testMappings = @(
    [PSCustomObject]@{
        OldBaseUrl = 'https://whallc.sharepoint.com/PFS'
        NewBaseUrl = 'https://whallc.sharepoint.com/sites/PatientFinancialServices'
        TenantBase = 'https://whallc.sharepoint.com'
        OldBasePath = '/PFS'
    }
    [PSCustomObject]@{
        OldBaseUrl = 'http://centralpark/PFS'
        NewBaseUrl = 'https://whallc.sharepoint.com/sites/PatientFinancialServices'
        TenantBase = 'http://centralpark'
        OldBasePath = '/PFS'
    }
    [PSCustomObject]@{
        OldBaseUrl = 'https://whallc.sharepoint.com/whaiis01'
        NewBaseUrl = 'https://whallc.sharepoint.com/sites/whaiis01'
        TenantBase = 'https://whallc.sharepoint.com'
        OldBasePath = '/whaiis01'
    }
) | Sort-Object { $_.OldBaseUrl.Length } -Descending

# ─── Load all function definitions from the real URLsv3 script ───────────────
# Extracts from $ErrorActionPreference down to (not including) Install-RequiredModule,
# which contains all constants and helper functions without any SharePoint calls.
$v3Path     = Join-Path $baseDir "Scripts\Get-FixItListURLsv3.ps1"
$v3Lines    = Get-Content $v3Path
$startIdx   = (($v3Lines | Select-String '^\$ErrorActionPreference\s*=' | Select-Object -First 1).LineNumber) - 1
$stopIdx    = (($v3Lines | Select-String '^Install-RequiredModule\s+-Name' | Select-Object -First 1).LineNumber) - 2
$tmpScript  = [System.IO.Path]::ChangeExtension([System.IO.Path]::GetTempFileName(), '.ps1')
$v3Lines[$startIdx..$stopIdx] | Set-Content $tmpScript -Encoding UTF8
. $tmpScript
Remove-Item $tmpScript -ErrorAction SilentlyContinue

Add-Type -AssemblyName System.IO.Compression.FileSystem

# ─── Regression checks for artifact filtering ───────────────────────────────
$nonActionableShouldBeBlocked = @(
    'https://whallc.sharepoint.com/_layouts/15/DocIdRedir.aspx?ID=ABC123',
    'https://whallc.sharepoint.com/_layouts/15/listform.aspx?PageType=8&ListId=abc',
    'https://whallc.sharepoint.com/_layouts/15/Redirector.aspx?Source=library',
    'https://schemas.openxmlformats.org/package/2006/relationships',
    'https://schemas.microsoft.com/office/2006/relationships'
)
$nonActionableFailures = @()
foreach ($u in $nonActionableShouldBeBlocked) {
    if (-not (Test-IsNonActionableUrl -Url $u)) {
        $nonActionableFailures += $u
    }
}
if ($nonActionableFailures.Count -gt 0) {
    throw "Non-actionable URL regression: $($nonActionableFailures -join '; ')"
}
Write-Host "Artifact filter regression checks passed: $($nonActionableShouldBeBlocked.Count) cases"

# ─── Build document metadata lookup from the real 20260810 report ────────────
$docMeta = @{}
foreach ($row in (Import-Csv $sourceCsv)) {
    $t = $row.'Document Title'
    if (-not $docMeta.ContainsKey($t)) { $docMeta[$t] = $row }
}

# ─── Scan files ───────────────────────────────────────────────────────────────
$rows   = New-Object System.Collections.Generic.List[PSCustomObject]
$seenFP = [System.Collections.Generic.HashSet[string]]::new()

foreach ($f in (Get-ChildItem $scanDir | Sort-Object Name)) {
    $ext  = $f.Extension.ToLowerInvariant()
    $title = $f.Name
    $meta  = $docMeta[$title]

    # Use the same dispatch logic as URLsv3
    $bytes = [System.IO.File]::ReadAllBytes($f.FullName)
    $refs  = @()
    if ($ext -eq '.pdf') {
        $refs = Get-PdfReference -Bytes $bytes
    } elseif ($OpenXmlExtensions -contains $ext) {
        $ms = New-Object System.IO.MemoryStream(,$bytes)
        $refs = Get-OfficeDocReference -Stream $ms
    } else {
        $refs = Get-TextReferenceFromByteContent -Bytes $bytes -Kind 'plaintext-doc'
    }

    foreach ($ref in $refs) {
        if (Test-IsSuppressedDocumentReference -Reference $ref) { continue }

        $cleanEmbedded = Get-CleanedValue -Value $ref.Value
        if (Test-IsNonActionableUrl -Url $cleanEmbedded) { continue }

        $matched   = Get-UrlMappingMatch -Value $ref.Value -Mappings $testMappings
        $brokenUrl = if ($matched) { $matched.OldUrl } else { $cleanEmbedded }
        $newUrl    = if ($matched) { $matched.NewUrl } else { '' }
        $linkType  = Get-LinkTypeLabel -Url $ref.Value -TenantHosts @()
        $brokenText = if (-not [string]::IsNullOrWhiteSpace([string]$ref.Text)) { [string]$ref.Text } else { $cleanEmbedded }

        $fp = "$title|$brokenUrl"
        if ($seenFP.Contains($fp)) { continue }
        $null = $seenFP.Add($fp)

        $rows.Add([PSCustomObject][ordered]@{
            'Link Type'           = $linkType
            'Site Title'          = if ($meta) { $meta.'Site Title'          } else { '(not in source CSV)' }
            'Document Title'      = $title
            'Document URL'        = if ($meta) { $meta.'Document URL'        } else { '' }
            'Document Folder'     = if ($meta) { $meta.'Document Folder'     } else { '' }
            'Document Folder URL' = if ($meta) { $meta.'Document Folder URL' } else { '' }
            'Broken URL Text'     = $brokenText
            'Broken Embedded URL' = $brokenUrl
            'New URL'             = $newUrl
            'Created By'          = if ($meta) { $meta.'Created By'          } else { '' }
            'Last Modified By'    = if ($meta) { $meta.'Last Modified By'    } else { '' }
            'Last Modified Date'  = if ($meta) { $meta.'Last Modified Date'  } else { '' }
            'Permissions'         = if ($meta) { $meta.'Permissions'         } else { '' }
        })
    }
}

if ($rows.Count -eq 0) {
    [PSCustomObject][ordered]@{
        'Link Type'=''; 'Site Title'=''; 'Document Title'=''; 'Document URL'='';
        'Document Folder'=''; 'Document Folder URL'=''; 'Broken URL Text'='';
        'Broken Embedded URL'='No findings'; 'New URL'='';
        'Created By'=''; 'Last Modified By'=''; 'Last Modified Date'=''; 'Permissions'=''
    } | Export-Csv $outputCsv -NoTypeInformation
} else {
    $rows | Export-Csv $outputCsv -NoTypeInformation
}

Write-Host "Written $($rows.Count) rows to $outputCsv"
$rows | Select-Object 'Document Title','Link Type','Broken Embedded URL','New URL' | Format-Table -AutoSize | Out-String -Width 220 | Write-Host
