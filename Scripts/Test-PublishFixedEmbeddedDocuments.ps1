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
    function Write-TestOfficePackage {
        param([string]$Path, [hashtable]$Parts)
        $zip = [IO.Compression.ZipFile]::Open($Path, [IO.Compression.ZipArchiveMode]::Create)
        try {
            foreach ($name in $Parts.Keys) {
                $writer = [IO.StreamWriter]::new($zip.CreateEntry($name).Open())
                try { $writer.Write($Parts[$name]) } finally { $writer.Dispose() }
            }
        } finally { $zip.Dispose() }
    }
    $officeLocal = Join-Path $temp 'local.docx'
    $officeRemote = Join-Path $temp 'remote.docx'
    $parts = @{
        '[Content_Types].xml' = '<Types/>'
        'word/document.xml' = '<document>Visible label</document>'
        'word/_rels/document.xml.rels' = '<Relationships><Relationship Target="https://whallc.sharepoint.com/sites/Epic/Tip/Fixed.docx"/></Relationships>'
        'customXml/item1.xml' = '<p:properties xmlns:p="http://schemas.microsoft.com/office/2006/metadata/properties"><documentManagement><Audience>Lab</Audience></documentManagement></p:properties>'
        'customXml/item9.xml' = '<businessData>Do not ignore me</businessData>'
        'docProps/custom.xml' = '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"><property>old-library</property></Properties>'
    }
    Write-TestOfficePackage $officeLocal $parts
    $changed = $parts.Clone()
    $changed['customXml/item1.xml'] = $parts['customXml/item1.xml'].Replace('Lab', 'Destination Lab')
    $changed['docProps/custom.xml'] = $parts['docProps/custom.xml'].Replace('old-library','new-library')
    $changed['[trash]/0001.dat'] = 'discarded metadata'
    Write-TestOfficePackage $officeRemote $changed
    Test-Case 'Office package comparison accepts only recognized metadata changes' {
        $result = Compare-OfficeDocumentPackage $officeLocal ([IO.File]::ReadAllBytes($officeRemote))
        Assert-Equal $result.Count 3
    }
    Test-Case 'Office package comparison rejects visible text, hyperlink and arbitrary custom XML edits' {
        foreach ($part in 'word/document.xml','word/_rels/document.xml.rels','customXml/item9.xml','[Content_Types].xml') {
            $bad = $changed.Clone()
            $bad[$part] = $bad[$part].Replace('Visible label','Changed label').Replace('Fixed.docx','Wrong.docx').
                Replace('Do not ignore me','Changed data').Replace('<Types/>','<Types><Changed/></Types>')
            $path = Join-Path $temp ([guid]::NewGuid().ToString('N') + '.docx')
            Write-TestOfficePackage $path $bad
            Assert-Throws { Compare-OfficeDocumentPackage $officeLocal ([IO.File]::ReadAllBytes($path)) } 'differs'
        }
    }
    Test-Case 'Office package comparison rejects missing parts and signatures' {
        $bad = $changed.Clone()
        $bad.Remove('word/document.xml')
        $path = Join-Path $temp 'missing-part.docx'
        Write-TestOfficePackage $path $bad
        Assert-Throws { Compare-OfficeDocumentPackage $officeLocal ([IO.File]::ReadAllBytes($path)) } 'added or removed'
        $bad = $changed.Clone()
        $bad['_xmlsignatures/sig1.xml'] = '<signature/>'
        $path = Join-Path $temp 'signed.docx'
        Write-TestOfficePackage $path $bad
        Assert-Throws { Compare-OfficeDocumentPackage $officeLocal ([IO.File]::ReadAllBytes($path)) } 'Signed Office'
    }
    Test-Case 'Office metadata datastore allowance rejects unrelated property edits' {
        $a = $parts.Clone()
        $b = $parts.Clone()
        $rels = '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps" Target="itemProps1.xml"/></Relationships>'
        $props = '<ds:datastoreItem xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml" ds:itemID="{old}"/>'
        $a['customXml/_rels/item1.xml.rels'] = $rels
        $b['customXml/_rels/item1.xml.rels'] = $rels
        $a['customXml/itemProps1.xml'] = $props
        $b['customXml/itemProps1.xml'] = $props.Replace('{old}','{new}')
        $ap = Join-Path $temp 'datastore-a.docx'
        $bp = Join-Path $temp 'datastore-b.docx'
        Write-TestOfficePackage $ap $a
        Write-TestOfficePackage $bp $b
        Assert-Equal (Compare-OfficeDocumentPackage $ap ([IO.File]::ReadAllBytes($bp))).Count 1
        $b['customXml/itemProps1.xml'] = $props.Replace('/>','><ds:schemaRefs/></ds:datastoreItem>')
        $cp = Join-Path $temp 'datastore-bad.docx'
        Write-TestOfficePackage $cp $b
        Assert-Throws { Compare-OfficeDocumentPackage $ap ([IO.File]::ReadAllBytes($cp)) } 'unapproved'
    }
    Test-Case 'Observed C-17 schema-reference removal is accepted without ignoring content' {
        $repo = Split-Path -Parent $PSScriptRoot
        $remotePath = Join-Path $repo 'UploadRuns\diagnostic-C17-20261019-062325\C-17.remote.docx'
        if (Test-Path -LiteralPath $remotePath) {
            $sourcePath = Join-Path $repo 'UpdatedEmbeddedLinkCopies\HR\C-17 Blood Specimen Collection & Competency Checklist 2026.04.15.docx'
            $diffs = Compare-OfficeDocumentPackage $sourcePath ([IO.File]::ReadAllBytes($remotePath))
            Assert-Equal $diffs.Count 9
        }
    }
    Test-Case 'Metadata datastore permits schema list removal but rejects replacement' {
        $a = $parts.Clone()
        $b = $parts.Clone()
        $rels = '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps" Target="itemProps1.xml"/></Relationships>'
        $head = '<ds:datastoreItem xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml" ds:itemID="old">'
        $refs = '<ds:schemaRefs><ds:schemaRef ds:uri="original"/></ds:schemaRefs>'
        $a['customXml/_rels/item1.xml.rels'] = $rels
        $b['customXml/_rels/item1.xml.rels'] = $rels
        $a['customXml/itemProps1.xml'] = $head + $refs + '</ds:datastoreItem>'
        $b['customXml/itemProps1.xml'] = $head + '</ds:datastoreItem>'
        $ap = Join-Path $temp 'schema-original.docx'
        $bp = Join-Path $temp 'schema-removed.docx'
        Write-TestOfficePackage $ap $a
        Write-TestOfficePackage $bp $b
        Assert-Equal (Compare-OfficeDocumentPackage $ap ([IO.File]::ReadAllBytes($bp))).Count 1
        $b['customXml/itemProps1.xml'] = $head + $refs.Replace('original','replacement') + '</ds:datastoreItem>'
        $cp = Join-Path $temp 'schema-replaced.docx'
        Write-TestOfficePackage $cp $b
        Assert-Throws { Compare-OfficeDocumentPackage $ap ([IO.File]::ReadAllBytes($cp)) } 'unapproved'
    }
    Test-Case 'Office comparison is opt-in, audited and does not weaken PDFs' {
        $p = [pscustomobject]@{
            LocalPath = $officeLocal; SHA256 = (Get-FileHash -LiteralPath $officeLocal).Hash
            FileRef = '/sites/Epic/Tip/local.docx'; VerificationAuditPath = (Join-Path $temp 'office-audit.jsonl')
        }
        $bytes = [IO.File]::ReadAllBytes($officeRemote)
        Assert-Throws { Assert-PlannedDocumentContent $p $bytes } 'Remote bytes differ'
        $AllowOfficeMetadataChanges = $true
        Assert-Equal (Assert-PlannedDocumentContent $p $bytes) 'OfficePackageContent'
        $audit = Get-Content -LiteralPath $p.VerificationAuditPath -Raw | ConvertFrom-Json
        Assert-Equal @($audit.Differences).Count 3
        $p.LocalPath = Join-Path $temp 'local.pdf'
        Assert-Throws { Assert-PlannedDocumentContent $p $bytes } 'Remote bytes differ'
    }
    Test-Case 'Observed SharePoint Word rewrite preserves all non-metadata parts' {
        $repo = Split-Path -Parent $PSScriptRoot
        $remotePath = Join-Path $repo 'UploadRuns\diagnostic-20261019-061217\Documenting an Adverse Reaction.remote.docx'
        if (Test-Path -LiteralPath $remotePath) {
            $sourcePath = Join-Path $repo 'UpdatedEmbeddedLinkCopies\Epic\Documenting an Adverse Reaction.docx'
            $diffs = Compare-OfficeDocumentPackage $sourcePath ([IO.File]::ReadAllBytes($remotePath))
            Assert-Equal $diffs.Count 8
        }
    }
    Test-Case 'REST metadata preserves explicit winter and summer UTC dates' {
        function Invoke-PnPSPRestMethod {
            param($Url, $Method, $Connection, $ErrorAction)
            Assert-Equal $Method 'Get'
            return [pscustomobject]@{
                Id = 22; AuthorId = 100; EditorId = 249
                Created = '2023-12-05T10:27:56Z'; Modified = '2024-07-08T15:12:21Z'
            }
        }
        $s = Get-RemoteSystemMetadata '9d3531ea-1546-4827-ba72-ad0a5b77d11a' 22 'connection'
        Assert-Equal $s.Created '2023-12-05T10:27:56.0000000Z'
        Assert-Equal $s.Modified '2024-07-08T15:12:21.0000000Z'
        Assert-Equal $s.Author 100
    }
    Test-Case 'REST metadata refuses timezone-less timestamps' {
        function Invoke-PnPSPRestMethod {
            param($Url, $Method, $Connection, $ErrorAction)
            return [pscustomobject]@{
                Id = 22; AuthorId = 100; EditorId = 249
                Created = '2023-12-05T02:27:56'; Modified = '2024-07-08T08:12:21'
            }
        }
        Assert-Throws { Get-RemoteSystemMetadata 'list' 22 'connection' } 'no explicit timezone'
    }
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
        Set-Variable -Name PlanCsv -Value $csv -Scope Script
        Set-Variable -Name DocumentsFolder -Value $temp -Scope Script
        $OutputFolder = Join-Path $temp 'offline-output'
        Invoke-PublishFixedDocuments
        Assert-Equal @(Import-Csv -LiteralPath (Join-Path $OutputFolder 'upload-plan.csv')).Count 1
        Remove-Item -LiteralPath (Join-Path $OutputFolder 'upload-plan.csv')
        Remove-Item -LiteralPath $OutputFolder
    }
    Test-Case 'PnP authentication reuses persistent login without forcing ClientId' {
        function Connect-PnPOnline {
            param($Url, [switch]$Interactive, [switch]$PersistLogin, [switch]$ReturnConnection, $ClientId, $ErrorAction)
            Assert-Equal $Url 'https://whallc.sharepoint.com/sites/Epic'
            Assert-Equal $Interactive.IsPresent $true
            Assert-Equal $PersistLogin.IsPresent $true
            Assert-Equal $ReturnConnection.IsPresent $true
            Assert-Equal $PSBoundParameters.ContainsKey('ClientId') $false
            return 'cached-connection'
        }
        Assert-Equal (Connect-DocumentUploadSite 'https://whallc.sharepoint.com/sites/Epic') 'cached-connection'
    }
    Test-Case 'Explicit PnP application ID is passed through; placeholders are rejected' {
        function Connect-PnPOnline {
            param($Url, [switch]$Interactive, [switch]$PersistLogin, [switch]$ReturnConnection, $ClientId, $ErrorAction)
            Assert-Equal $ClientId '12345678-1234-1234-1234-123456789abc'
            return 'explicit-connection'
        }
        Assert-Equal (Connect-DocumentUploadSite 'https://whallc.sharepoint.com/sites/Epic' '12345678-1234-1234-1234-123456789abc') 'explicit-connection'
        Assert-Throws { Connect-DocumentUploadSite 'https://whallc.sharepoint.com/sites/Epic' '<your-existing-PnP-application-client-id>' } 'GUID'
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
    Test-Case 'Metadata CSOM write preserves UTC kinds and original user IDs' {
        $script:metadataQueries = 0
        $script:metadataUpdates = 0
        $context = [pscustomobject]@{}
        $context | Add-Member ScriptMethod ExecuteQuery { $script:metadataQueries++ }
        $item = @{ Context = $context }
        $item | Add-Member ScriptMethod UpdateOverwriteVersion { $script:metadataUpdates++ }
        $values = @{
            Author = '100'; Editor = '249'
            Created = [datetime]::Parse('2023-12-05T10:27:56.0000000Z', [cultureinfo]::InvariantCulture, [Globalization.DateTimeStyles]::RoundtripKind)
            Modified = [datetime]::Parse('2024-07-08T15:12:21.0000000Z', [cultureinfo]::InvariantCulture, [Globalization.DateTimeStyles]::RoundtripKind)
        }
        Set-DocumentMetadataValues ([pscustomobject]@{ Item = $item }) $values
        Assert-Equal $item.Created.Kind ([DateTimeKind]::Utc)
        Assert-Equal $item.Created.ToString('o') '2023-12-05T10:27:56.0000000Z'
        Assert-Equal $item.Modified.ToString('o') '2024-07-08T15:12:21.0000000Z'
        Assert-Equal $item.Author.LookupId 100
        Assert-Equal $item.Editor.LookupId 249
        Assert-Equal $script:metadataUpdates 1
        Assert-Equal $script:metadataQueries 1
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
    Test-Case 'Stable backup retries a changed snapshot and returns only stable bytes' {
        $script:reads = 0
        $script:downloads = 0
        function Start-Sleep { param($Seconds) Assert-Equal $Seconds 2 }
        function Get-RemoteDocumentState {
            param($Plan, $Connection)
            $script:reads++
            $s = New-TestState
            if ($script:reads -ge 2) { $s.ETag = '"2"' }
            return $s
        }
        function Get-RemoteDocumentBytes {
            param($Plan, $Connection)
            $script:downloads++
            return ,[Text.Encoding]::UTF8.GetBytes("download-$script:downloads")
        }
        $snapshot = Read-StableDocumentBackup $plan 'connection'
        Assert-Equal $script:reads 4
        Assert-Equal ([Text.Encoding]::UTF8.GetString($snapshot.Bytes)) 'download-2'
        Assert-Equal $snapshot.State.ETag '"2"'
    }
    Test-Case 'Persistently changing backup fails after three reads with diagnostics' {
        $script:reads = 0
        function Start-Sleep { param($Seconds) }
        function Get-RemoteDocumentState {
            param($Plan, $Connection)
            $script:reads++
            $s = New-TestState
            $s.ETag = [string]$script:reads
            return $s
        }
        Assert-Throws { Read-StableDocumentBackup $plan 'connection' } 'attempt 3/3.*ETag.*No stable backup'
        Assert-Equal $script:reads 6
    }
    Test-Case 'Backup rejects identity changes without retrying' {
        $script:reads = 0
        function Get-RemoteDocumentState {
            param($Plan, $Connection)
            $script:reads++
            $s = New-TestState
            $s.UniqueId = [string]$script:reads
            return $s
        }
        Assert-Throws { Read-StableDocumentBackup $plan 'connection' } 'identity changed'
        Assert-Equal $script:reads 2
    }
    function Send-ConditionalDocumentContent {
        param($Plan, $Current)
        $script:uploads++
        $script:remoteBytes = [IO.File]::ReadAllBytes($Plan.LocalPath)
        $script:remote.Editor = 999
        $script:remote.Modified = '2026-10-19T00:00:00.0000000Z'
        if ($script:uploadFailure) { throw 'simulated ambiguous network failure' }
    }
    function Set-DocumentMetadataValues {
        param($Current, $Values)
        $script:setCalls++
        Assert-Equal $Current.ListId 'existing-list'
        Assert-Equal $Current.ItemId 17
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
    $backup = Join-Path $temp 'original-backup.pdf'
    [IO.File]::WriteAllText($backup, 'original-content')
    $snapshotPath = $backup + '.metadata.json'
    $snapshotData = [ordered]@{
        Target = $plan.FileRef; WebUrl = $plan.WebUrl
        SHA256 = (Get-FileHash -LiteralPath $backup).Hash
        UniqueId = 'existing-file'; ItemId = 17; ListId = 'existing-list'
        Author = 8; Editor = 9
        Created = '2024-01-02T03:04:05.0000000Z'; Modified = '2025-02-03T04:05:06.0000000Z'
    }
    $snapshotData | ConvertTo-Json | Set-Content -LiteralPath $snapshotPath
    function Connect-DocumentUploadSite { param($Url, $ApplicationId) return 'connection' }
    Test-Case 'Snapshot recovery preserves original UTC strings and never uploads content' {
        $script:remote = New-TestState
        $script:remote.Created = '2024-01-01T19:04:05.0000000Z'
        $script:remoteBytes = [IO.File]::ReadAllBytes($plan.LocalPath)
        $uploadsBefore = $script:uploads
        $result = Restore-SnapshotDocumentMetadata $snapshotPath @($plan)
        Assert-Equal $result.Status 'MetadataRestoredAndVerified'
        Assert-Equal $script:remote.Created $snapshotData.Created
        Assert-Equal $script:uploads $uploadsBefore
    }
    Test-Case 'Snapshot recovery accepts audited metadata-only Office variants without content upload' {
        $AllowOfficeMetadataChanges = $true
        $officePlan = [pscustomobject]@{
            LocalPath = $officeLocal; SHA256 = (Get-FileHash -LiteralPath $officeLocal).Hash
            FileRef = $plan.FileRef; WebUrl = $plan.WebUrl
            VerificationAuditPath = (Join-Path $temp 'recovery-office-audit.jsonl')
        }
        $script:remote = New-TestState
        $script:remote.Editor = 999
        $script:remoteBytes = [IO.File]::ReadAllBytes($officeRemote)
        $uploadsBefore = $script:uploads
        $result = Restore-SnapshotDocumentMetadata $snapshotPath @($officePlan)
        Assert-Equal $result.Status 'MetadataRestoredAndVerified'
        Assert-Equal $script:remote.Editor 9
        Assert-Equal $script:uploads $uploadsBefore
        Assert-Equal @(Get-Content -LiteralPath $officePlan.VerificationAuditPath).Count 2
    }
    Test-Case 'Snapshot recovery rejects mismatched destination and damaged backup' {
        $snapshotData.Target = '/sites/Epic/Wrong/Guide.pdf'
        $snapshotData | ConvertTo-Json | Set-Content -LiteralPath $snapshotPath
        Assert-Throws { Restore-SnapshotDocumentMetadata $snapshotPath @($plan) } 'exactly one'
        $snapshotData.Target = $plan.FileRef
        $snapshotData | ConvertTo-Json | Set-Content -LiteralPath $snapshotPath
        [IO.File]::WriteAllText($backup, 'damaged')
        Assert-Throws { Restore-SnapshotDocumentMetadata $snapshotPath @($plan) } 'backup hash'
    }
    Write-Host "$script:passed tests passed. No SharePoint connections or writes were made."
} finally {
    foreach ($file in Get-ChildItem -LiteralPath $temp -File) { Remove-Item -LiteralPath $file.FullName }
    Remove-Item -LiteralPath $temp
}
