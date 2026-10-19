#requires -Version 7.4.6
[CmdletBinding()]
param()
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'Publish-FixedEmbeddedDocuments.ps1')

$script:passed = 0
function Assert-Equal($Actual, $Expected) {
    if ($Actual -cne $Expected) { throw "Expected [$Expected], received [$Actual]." }
}
function Assert-Throws([scriptblock]$Action, [string]$Pattern) {
    try { & $Action | Out-Null } catch {
        if ($_.Exception.Message -notmatch $Pattern) { throw }
        return
    }
    throw "Expected an exception matching: $Pattern"
}
function Test-Case([string]$Name, [scriptblock]$Action) {
    & $Action
    $script:passed++
    Write-Host "PASS $Name"
}

$temp = Join-Path ([IO.Path]::GetTempPath()) ('wha-upload-tests-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $temp | Out-Null
try {
    Test-Case 'Library view resolves without Forms/AllItems.aspx' {
        $d = Resolve-UploadDestination 'https://whallc.sharepoint.com/sites/Epic/Tip%20Sheets/Forms/AllItems.aspx' 'Guide.pdf'
        Assert-Equal $d.FileRef '/sites/Epic/Tip Sheets/Guide.pdf'
        Assert-Equal $d.WebUrl 'https://whallc.sharepoint.com/sites/Epic'
    }
    Test-Case 'Direct subfolder and doubled spaces are retained' {
        $d = Resolve-UploadDestination 'https://whallc.sharepoint.com/sites/whasites/Policies%20%20Procedures/Subfolder' 'Guide.docx'
        Assert-Equal $d.FileRef '/sites/whasites/Policies  Procedures/Subfolder/Guide.docx'
    }
    Test-Case 'Folder view honors RootFolder and encoded ampersands' {
        $d = Resolve-UploadDestination 'https://whallc.sharepoint.com/sites/Epic/Tip%20Sheets/Forms/AllItems.aspx?RootFolder=%2Fsites%2FEpic%2FTip%20Sheets%2FLab%20%26%20RN' 'Guide.pdf'
        Assert-Equal $d.FileRef '/sites/Epic/Tip Sheets/Lab & RN/Guide.pdf'
    }
    Test-Case 'Untrusted host and cross-library folder are rejected' {
        Assert-Throws { Resolve-UploadDestination 'https://whallc.sharepoint.com.evil.test/sites/Epic/Tip' 'Guide.pdf' } 'Untrusted'
        Assert-Throws { Resolve-UploadDestination 'https://whallc.sharepoint.com/sites/Epic/Tip/Forms/AllItems.aspx?id=%2Fsites%2FHR%2FDocs' 'Guide.pdf' } 'leaves its library'
        Assert-Throws { Resolve-UploadDestination 'https://whallc.sharepoint.com/sites/Epic/Tip' '..\Guide.pdf' } 'Invalid document'
    }
    Test-Case 'File-selection and unsupported ASPX views are rejected' {
        Assert-Throws { Resolve-UploadDestination 'https://whallc.sharepoint.com/sites/Epic/Tip/Forms/AllItems.aspx?id=%2Fsites%2FEpic%2FTip%2FGuide.pdf' 'Guide.pdf' } 'selects a file'
        Assert-Throws { Resolve-UploadDestination 'https://whallc.sharepoint.com/sites/Epic/Tip/Forms/Other.aspx' 'Guide.pdf' } 'Unsupported'
    }
    Test-Case 'CSOM unspecified dates are interpreted as UTC, not server local time' {
        Assert-Equal (ConvertTo-SharePointTimestamp ([datetime]::new(2025,1,2,3,4,5))) '2025-01-02T03:04:05.0000000Z'
    }
    $local = Join-Path $temp 'Guide.pdf'
    [IO.File]::WriteAllText($local, 'fixed-content')
    $entry = [pscustomobject]@{
        RelativeFile = 'Guide.pdf'; LocalPath = 'C:\old-machine\Guide.pdf'
        SHA256 = (Get-FileHash -LiteralPath $local).Hash
        WebUrl = 'https://whallc.sharepoint.com/sites/Epic'
        Folder = '/sites/Epic/Tip Sheets'; FileRef = '/sites/Epic/Tip Sheets/Guide.pdf'
    }
    $csv = Join-Path $temp 'plan.csv'
    $entry | Export-Csv -LiteralPath $csv -NoTypeInformation
    Test-Case 'Portable plan rebases local paths and checks bytes' {
        $plan = @(Import-FixedDocumentPlan $csv $temp)
        Assert-Equal $plan.Count 1
        Assert-Equal $plan[0].LocalPath $local
    }
    Test-Case 'Portable plan rejects local tampering' {
        [IO.File]::WriteAllText($local, 'tampered')
        Assert-Throws { Import-FixedDocumentPlan $csv $temp } 'differs from portable plan'
        [IO.File]::WriteAllText($local, 'fixed-content')
    }
    Test-Case 'Portable plan rejects traversal and inconsistent destinations' {
        $bad = $entry.PSObject.Copy()
        $bad.RelativeFile = '..\Guide.pdf'
        $bad | Export-Csv -LiteralPath $csv -NoTypeInformation
        Assert-Throws { Import-FixedDocumentPlan $csv $temp } 'unsafe'
        $bad = $entry.PSObject.Copy()
        $bad.FileRef = '/sites/Epic/Other/Guide.pdf'
        $bad | Export-Csv -LiteralPath $csv -NoTypeInformation
        Assert-Throws { Import-FixedDocumentPlan $csv $temp } 'Inconsistent'
        $entry | Export-Csv -LiteralPath $csv -NoTypeInformation
    }
    Test-Case 'Portable plan rejects duplicate destinations' {
        @($entry, $entry) | Export-Csv -LiteralPath $csv -NoTypeInformation
        Assert-Throws { Import-FixedDocumentPlan $csv $temp } 'duplicate destination'
        $entry | Export-Csv -LiteralPath $csv -NoTypeInformation
    }
    Test-Case 'Default main execution only writes an offline plan' {
        function Connect-PnPOnline { throw 'Offline planning must not authenticate.' }
        $PlanCsv = $csv
        $DocumentsFolder = $temp
        $OutputFolder = Join-Path $temp 'offline-output'
        Invoke-PublishFixedDocuments
        Assert-Equal @(Import-Csv -LiteralPath (Join-Path $OutputFolder 'upload-plan.csv')).Count 1
        Remove-Item -LiteralPath (Join-Path $OutputFolder 'upload-plan.csv')
        Remove-Item -LiteralPath $OutputFolder
    }
    Test-Case 'Conditional CSOM write submits the original ETag and disables retries' {
        Import-Module PnP.PowerShell -MinimumVersion 3.3.0 -ErrorAction Stop
        $script:information = $null
        $script:queryCalls = 0
        $fakeContext = [pscustomobject]@{}
        $fakeFile = [pscustomobject]@{ Context = $fakeContext }
        $fakeFile | Add-Member -MemberType ScriptMethod -Name SaveBinary -Value {
            param($Information)
            $script:information = $Information
        }
        $fakeContext | Add-Member -MemberType ScriptMethod -Name ExecuteQuery -Value {
            Assert-Equal $script:information.ETag '"preflight-etag"'
            Assert-Equal $script:information.CheckRequiredFields $true
            $reader = [IO.StreamReader]::new($script:information.ContentStream, [Text.Encoding]::UTF8, $false, 1024, $true)
            try { Assert-Equal $reader.ReadToEnd() 'fixed-content' } finally { $reader.Dispose() }
            $script:queryCalls++
        }
        $rebased = @(Import-FixedDocumentPlan $csv $temp)[0]
        Send-ConditionalDocumentContent $rebased ([pscustomobject]@{ File = $fakeFile; ETag = '"preflight-etag"' })
        Assert-Equal $script:queryCalls 1
    }
    $plan = @(Import-FixedDocumentPlan $csv $temp)[0]
    function New-TestState {
        return [pscustomobject]@{
            UniqueId = 'existing-file'; ItemId = 17; ListId = 'existing-list'; ETag = '"1"'
            Author = 8; Editor = 9; Created = '2024-01-02T03:04:05.0000000Z'
            Modified = '2025-02-03T04:05:06.0000000Z'
            File = [pscustomobject]@{ CheckOutType = 'None'; LockedByUser = $null }
            Item = [pscustomobject]@{ FieldValues = @{ _UIVersionString = '4.0' } }
            List = [pscustomobject]@{ ForceCheckout = $false; EnableModeration = $false }
        }
    }
    Test-Case 'Checkout, locks, retention, approval and drafts block upload' {
        $s = New-TestState; $s.File.CheckOutType = 'Online'
        Assert-Throws { Assert-RemoteDocumentWritable $s } 'checked out'
        $s = New-TestState; $s.File.LockedByUser = [pscustomobject]@{ ServerObjectIsNull = $false }
        Assert-Throws { Assert-RemoteDocumentWritable $s } 'locked'
        $s = New-TestState; $s.Item.FieldValues['_ComplianceTag'] = 'Record'
        Assert-Throws { Assert-RemoteDocumentWritable $s } 'retention'
        $s = New-TestState; $s.List.EnableModeration = $true
        Assert-Throws { Assert-RemoteDocumentWritable $s } 'approval'
        $s = New-TestState; $s.List.ForceCheckout = $true
        Assert-Throws { Assert-RemoteDocumentWritable $s } 'check-out'
        $s = New-TestState; $s.Item.FieldValues['_UIVersionString'] = '4.1'
        Assert-Throws { Assert-RemoteDocumentWritable $s } 'minor/draft'
    }
    $script:before = New-TestState
    $script:remote = New-TestState
    $script:remoteBytes = [Text.Encoding]::UTF8.GetBytes('old-content')
    $script:uploads = 0
    $script:setCalls = 0
    $script:uploadFailure = $false
    $script:metadataFailure = $false
    function Get-RemoteDocumentState { param($Plan, $Connection) return $script:remote }
    function Get-RemoteDocumentBytes { param($Plan, $Connection) return ,$script:remoteBytes }
    function Send-ConditionalDocumentContent {
        param($Plan, $Current)
        $script:uploads++
        $script:remoteBytes = [IO.File]::ReadAllBytes($Plan.LocalPath)
        $script:remote.Editor = 999
        $script:remote.Modified = '2026-10-19T00:00:00.0000000Z'
        if ($script:uploadFailure) { throw 'simulated ambiguous network failure' }
    }
    function Set-PnPListItem {
        param($List, $Identity, $UpdateType, $Values, $Connection, $ErrorAction)
        $script:setCalls++
        Assert-Equal $List 'existing-list'
        Assert-Equal $Identity 17
        Assert-Equal $UpdateType 'UpdateOverwriteVersion'
        if ($script:metadataFailure) { throw 'simulated metadata denial' }
        $script:remote.Author = [int]$Values.Author
        $script:remote.Editor = [int]$Values.Editor
        $script:remote.Created = ConvertTo-SharePointTimestamp $Values.Created
        $script:remote.Modified = ConvertTo-SharePointTimestamp $Values.Modified
    }
    Test-Case 'Upload preserves and verifies original metadata and content' {
        $result = Invoke-FixedDocumentUpload $plan $script:before (Get-BytesSHA256 $script:remoteBytes) 'connection'
        Assert-Equal $result 'UploadedAndVerified'
        Assert-Equal $script:uploads 1
        Assert-Equal $script:remote.Editor 9
        Assert-Equal $script:remote.Modified $script:before.Modified
    }
    Test-Case 'Identical remote bytes skip the write' {
        Assert-Equal (Invoke-FixedDocumentUpload $plan $script:before $plan.SHA256 'connection') 'AlreadyIdentical'
        Assert-Equal $script:uploads 1
    }
    Test-Case 'Concurrent ETag change prevents any upload' {
        $script:remote.ETag = '"changed"'
        Assert-Throws { Invoke-FixedDocumentUpload $plan $script:before $plan.SHA256 'connection' } 'Destination changed'
        Assert-Equal $script:uploads 1
        $script:remote = New-TestState
    }
    Test-Case 'Concurrent remote byte change prevents any upload' {
        Assert-Throws { Invoke-FixedDocumentUpload $plan $script:before 'incorrect-hash' 'connection' } 'content changed'
        Assert-Equal $script:uploads 1
    }
    Test-Case 'Ambiguous upload error does not write potentially stale metadata' {
        $script:uploadFailure = $true
        $script:remoteBytes = [Text.Encoding]::UTF8.GetBytes('old-content')
        $beforeSetCalls = $script:setCalls
        Assert-Throws { Invoke-FixedDocumentUpload $plan $script:before (Get-BytesSHA256 $script:remoteBytes) 'connection' } 'No metadata recovery attempted'
        Assert-Equal $script:setCalls $beforeSetCalls
        Assert-Equal $script:remote.Editor 999
        $script:uploadFailure = $false
    }
    Test-Case 'Metadata failure is surfaced, not successful upload status' {
        $script:metadataFailure = $true
        $script:remote = New-TestState
        $script:remoteBytes = [Text.Encoding]::UTF8.GetBytes('old-content')
        Assert-Throws { Invoke-FixedDocumentUpload $plan $script:before (Get-BytesSHA256 $script:remoteBytes) 'connection' } 'metadata denial'
        $script:metadataFailure = $false
    }
    Test-Case 'Identity change prevents restoring metadata on a replacement file' {
        $script:remote = New-TestState
        $script:remote.UniqueId = 'different-file'
        Assert-Throws { Restore-DocumentMetadata $plan $script:before 'connection' } 'identity changed'
    }
    Test-Case 'Metadata restoration rejects concurrent ETag and content changes' {
        $script:remote = New-TestState
        Assert-Throws { Restore-DocumentMetadata $plan $script:before 'connection' '"different-etag"' } 'ETag changed'
        $script:remoteBytes = [Text.Encoding]::UTF8.GetBytes('concurrent-editor-content')
        Assert-Throws { Restore-DocumentMetadata $plan $script:before 'connection' $script:remote.ETag } 'Remote bytes differ'
    }
    Write-Host "$script:passed tests passed. No SharePoint connections or writes were made."
} finally {
    foreach ($file in Get-ChildItem -LiteralPath $temp -File) { Remove-Item -LiteralPath $file.FullName }
    Remove-Item -LiteralPath $temp
}
