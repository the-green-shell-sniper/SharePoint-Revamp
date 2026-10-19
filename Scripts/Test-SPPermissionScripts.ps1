#requires -Version 7.4.6
[CmdletBinding()]
param()
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'New-SPPermissionMatchMap.ps1')
. (Join-Path $PSScriptRoot 'Sync-SPPermissionsFromSource.ps1')

$script:passed = 0
function Assert-Equal($Actual, $Expected) {
    if ($Actual -cne $Expected) { throw "Expected [$Expected], received [$Actual]." }
}
function Assert-Throws([scriptblock]$Action, [string]$Pattern) {
    try { & $Action | Out-Null } catch {
        if ($_.Exception.Message -notmatch $Pattern) { throw }
        return
    }
    throw "Expected exception matching '$Pattern'."
}
function Test-Case([string]$Name, [scriptblock]$Action) {
    & $Action
    $script:passed++
    Write-Host "PASS $Name"
}
function New-TestMap {
    ,([Collections.Generic.Dictionary[string, object]]::new([StringComparer]::OrdinalIgnoreCase))
}
function New-TestAcl([string]$Login, [string[]]$Roles = @('Read'), [int]$Id = 12, [int]$PType = 1) {
    [pscustomobject]@{ LoginName = $Login; Title = $Login; MemberId = $Id; PrincipalType = $PType; Roles = $Roles; Bindings = @(); Skip = $false }
}
function New-TestScope([string]$Key, [bool]$Unique = $true, $Acl = @(), [string]$Kind = 'File') {
    if (-not $Key) { $Kind = 'Web' }
    [pscustomobject]@{
        Key = $Key; Kind = $Kind; Unique = $Unique; Acl = @($Acl); Api = "https://example.test/_api/$Key"
        Id = $Key; BaseTemplate = 101; Ext = '.docx'; ListKey = 'Docs'; Leaf = ($Key -split '/')[-1]
        Title = ($Key -split '/')[-1]; Created = '2020-01-01'; Author = 'author@example.test'; Size = 123
        Rel = ($Key -replace '^Docs/', ''); Depth = ($Key -split '/').Count; ItemCount = 1; FileRef = "/$Key"
    }
}
function New-TestDesired([string]$Login, [string[]]$Roles = @('Read')) {
    $map = New-TestMap
    $map[$Login] = [pscustomobject]@{
        Login = $Login; Title = $Login; PType = 1; Id = 12
        Roles = [Collections.Generic.HashSet[string]]::new($Roles, [StringComparer]::OrdinalIgnoreCase)
    }
    return ,$map
}
function New-TestPlanFixture {
    $fixture = [pscustomobject]@{ Source = New-TestMap; Target = New-TestMap; Matches = New-TestMap; Desired = New-TestMap }
    foreach ($key in @('', 'Docs', 'Docs/a.docx')) {
        $kind = if ($key -eq 'Docs') { 'Library' } else { 'File' }
        $fixture.Source[$key] = New-TestScope $key ($key -eq '') @(New-TestAcl 'visitors') $kind
        $fixture.Target[$key] = New-TestScope $key ($key -eq '') @(New-TestAcl 'old-user') $kind
        $fixture.Matches[$key] = $key
        $fixture.Desired[$key] = New-TestDesired 'visitors'
    }
    $fixture
}
$roles = [Collections.Generic.Dictionary[string, int]]::new([StringComparer]::OrdinalIgnoreCase)
$roles['Read'] = 1073741826; $roles['Design'] = 1073741828; $roles['Contribute'] = 1073741827
$temp = Join-Path ([IO.Path]::GetTempPath()) ('sp-permission-tests-' + [guid]::NewGuid().ToString('N'))
[void][IO.Directory]::CreateDirectory($temp)
try {
    Test-Case 'Interactive login uses explicit tenant, persistent cache and returned connection' {
        function Connect-PnPOnline {
            param($Url, $ClientId, $Tenant, [switch]$Interactive, [switch]$PersistLogin, [switch]$ReturnConnection, $ErrorAction)
            Assert-Equal $Tenant 'tenant-guid'
            Assert-Equal $ClientId 'app-guid'
            Assert-Equal $PersistLogin.IsPresent $true
            Assert-Equal $ReturnConnection.IsPresent $true
            Assert-Equal $Interactive.IsPresent $true
            'connection'
        }
        Assert-Equal (Connect-SPPermissionSite 'https://example.test' 'app-guid' 'tenant-guid') 'connection'
    }
    Test-Case 'Cache setup escape hatch omits PersistLogin without clearing sessions' {
        function Connect-PnPOnline {
            param($Url, [switch]$Interactive, [switch]$ReturnConnection, $ErrorAction, [switch]$PersistLogin)
            Assert-Equal $PSBoundParameters.ContainsKey('PersistLogin') $false
            'connection'
        }
        Assert-Equal (Connect-SPPermissionSite 'https://example.test' '' '' -SkipPersistLoginSetup) 'connection'
    }
    Test-Case 'REST writes are explicit-connection and never retried after ambiguous failure' {
        $script:writeAttempts = 0
        function Invoke-PnPSPRestMethod {
            param($Connection, $Url, $Method, [switch]$Raw, $ErrorAction)
            Assert-Equal $Connection 'target-only'
            $script:writeAttempts++
            throw '503 lost response'
        }
        Assert-Throws { Invoke-SPPermissionRest 'target-only' 'https://example.test/_api' -Method Post } '503'
        Assert-Equal $script:writeAttempts 1
    }
    Test-Case 'REST collection pagination follows continuation and accepts an empty final page' {
        function Invoke-SPPermissionRest {
            param($Connection, $Url)
            if ($Url -match 'page2') { return [pscustomobject]@{ value = @() } }
            [pscustomobject]@{ value = @([pscustomobject]@{ Id = 1 }); 'odata.nextLink' = 'https://example.test/page2' }
        }
        $result = @(Get-SPPermissionCollection 'source' 'https://example.test/page1')
        Assert-Equal $result.Count 1
        Assert-Equal $result[0].Id 1
    }
    Test-Case 'Malformed collection cannot become an empty inventory' {
        function Invoke-SPPermissionRest { [pscustomobject]@{ error = 'denied' } }
        Assert-Throws { Get-SPPermissionCollection 'source' 'https://example.test' } 'Expected REST collection'
    }
    Test-Case 'Literal folder names containing wildcard syntax remain literal' {
        Assert-Equal (Test-SPPermissionPathWithin 'Docs/[Private]/a.docx' 'Docs/[Private]') $true
        Assert-Equal (Test-SPPermissionPathWithin 'Docs/P/a.docx' 'Docs/[Private]') $false
        Assert-Equal (Test-SPPermissionPathWithin 'Docs2/a.docx' 'Docs') $false
    }
    Test-Case 'Inventory errors propagate instead of returning a partial map' {
        function Invoke-SPPermissionRest {
            [pscustomobject]@{ Id = 'web'; Url = 'https://example.test'; ServerRelativeUrl = '/sites/source'; HasUniqueRoleAssignments = $true }
        }
        function Get-SPPermissionCollection {
            param($Connection, $Url)
            if ($Url -match '/webs') { return }
            if ($Url -match '/lists\?') {
                return [pscustomobject]@{
                    Id = 'list'; Title = 'Docs'; BaseType = 1; BaseTemplate = 101; ItemCount = 1; HasUniqueRoleAssignments = $false
                    RootFolder = [pscustomobject]@{ ServerRelativeUrl = '/sites/source/Docs' }
                }
            }
            throw 'Cannot enumerate items'
        }
        Assert-Throws { Get-SPPermissionInventory 'source' 'https://example.test' } 'Cannot enumerate items'
    }
    Test-Case 'SharePoint group mapping uses title but retains target login and ID' {
        $groups = New-TestMap
        $groups['Visitors display name'] = [pscustomobject]@{ LoginName = 'Visitors-login'; Id = 55 }
        $visitor = [pscustomobject]@{ Title = 'Visitors display name' }
        $sourceAcl = @(New-TestAcl 'Team Site Members' @('Read') 1 8)
        $desired = Get-SPDesiredAcl $sourceAcl $groups $visitor 'Team Site Members'
        Assert-Equal $desired['Visitors-login'].Id 55
        Assert-Equal $desired['Visitors-login'].Roles.Contains('Read') $true
    }
    Test-Case 'Security-trimmed or changing item counts cannot masquerade as complete enumeration' {
        function Invoke-SPPermissionRest {
            [pscustomobject]@{ Id = 'web'; Url = 'https://example.test'; ServerRelativeUrl = '/sites/source'; HasUniqueRoleAssignments = $true }
        }
        function Get-SPPermissionCollection {
            param($Connection, $Url)
            if ($Url -match '/webs') { return }
            if ($Url -match '/lists\?') {
                return [pscustomobject]@{
                    Id = 'list'; Title = 'Docs'; BaseType = 1; BaseTemplate = 101; ItemCount = 1; HasUniqueRoleAssignments = $false
                    RootFolder = [pscustomobject]@{ ServerRelativeUrl = '/sites/source/Docs' }
                }
            }
            return
        }
        Assert-Throws { Get-SPPermissionInventory 'source' 'https://example.test' } 'Incomplete or changing inventory'
    }
    Test-Case 'Missing principal blocks all removals but allows available additions' {
        $desired = New-TestDesired 'missing'
        $desired['available'] = (New-TestDesired 'available')['available']
        $unavailable = New-TestMap; $unavailable['missing'] = 'deleted'
        $delta = Get-SPPermissionDelta @(New-TestAcl 'existing') $desired $unavailable @() $roles
        Assert-Equal $delta.RemovalsBlocked $true
        Assert-Equal $delta.Operations.Count 1
        Assert-Equal $delta.Operations[0].Type 'Add'
        Assert-Equal $delta.Operations[0].Login 'available'
    }
    Test-Case 'Unknown permission level is fatal before removals can be planned' {
        Assert-Throws {
            Get-SPPermissionDelta @(New-TestAcl 'existing') (New-TestDesired 'new' @('Unknown')) (New-TestMap) @() $roles
        } 'Permission level missing'
    }
    Test-Case 'Root Owners and Members are retained; source Designer role corrections remain exact' {
        $acl = @(New-TestAcl 'owners'; New-TestAcl 'members'; New-TestAcl 'designer' @('Design', 'Contribute'))
        $delta = Get-SPPermissionDelta $acl (New-TestDesired 'designer' @('Design')) (New-TestMap) @('owners','members') $roles
        Assert-Equal $delta.Operations.Count 1
        Assert-Equal $delta.Operations[0].Role 'Contribute'
        Assert-Equal $delta.Operations[0].Type 'Remove'
    }
    Test-Case 'Target-only and excluded descendants freeze before their ancestor changes' {
        $fixture = New-TestPlanFixture
        $fixture.Target['Docs/new.docx'] = New-TestScope 'Docs/new.docx' $false
        $fixture.Source['Docs/excluded.docx'] = New-TestScope 'Docs/excluded.docx' $false
        $fixture.Target['Docs/excluded.docx'] = New-TestScope 'Docs/excluded.docx' $false
        $fixture.Matches['Docs/excluded.docx'] = 'Docs/excluded.docx'
        $plan = New-SPPermissionPlan $fixture.Source $fixture.Target $fixture.Matches $fixture.Desired (New-TestMap) $roles @() @('Docs/excluded.docx')
        Assert-Equal (@($plan.Steps | Where-Object Action -eq 'Freeze').Count) 2
        Assert-Equal $plan.Steps[0].Action 'Freeze'
        Assert-Equal ((Get-SPPermissionEffectiveScope $fixture.Target 'Docs/new.docx').Acl[0].LoginName) 'old-user'
    }
    Test-Case 'NoFreeze is explicit and reports affected content' {
        $fixture = New-TestPlanFixture
        $fixture.Target['Docs/new.docx'] = New-TestScope 'Docs/new.docx' $false
        $plan = New-SPPermissionPlan $fixture.Source $fixture.Target $fixture.Matches $fixture.Desired (New-TestMap) $roles @() @() -NoFreeze
        Assert-Equal (@($plan.Steps | Where-Object Action -eq 'Freeze').Count) 0
        Assert-Equal (@($plan.Notes | Where-Object Result -eq 'Review').Count) 1
    }
    Test-Case 'Dry-run projection uses changed parent ACL rather than stale snapshot' {
        $fixture = New-TestPlanFixture
        $fixture.Source['Docs/a.docx'].Unique = $true
        $plan = New-SPPermissionPlan $fixture.Source $fixture.Target $fixture.Matches $fixture.Desired (New-TestMap) $roles @() @()
        Assert-Equal (@($plan.Steps | Where-Object { $_.Key -eq 'Docs/a.docx' -and $_.Action -in @('Add','Remove') }).Count) 0
        Assert-Equal (@($plan.Steps | Where-Object { $_.Key -eq 'Docs/a.docx' -and $_.Action -eq 'Break' }).Count) 1
        $second = New-SPPermissionPlan $fixture.Source $fixture.Target $fixture.Matches $fixture.Desired (New-TestMap) $roles @() @()
        Assert-Equal $second.Steps.Count 0
    }
    Test-Case 'Inherited source root uses effective ACL and target root becomes unique' {
        $fixture = New-TestPlanFixture
        $fixture.Source[''].Unique = $false
        $fixture.Target[''].Unique = $false
        $plan = New-SPPermissionPlan $fixture.Source $fixture.Target $fixture.Matches $fixture.Desired (New-TestMap) $roles @() @()
        Assert-Equal $plan.Steps[0].Action 'Break'
        Assert-Equal $plan.Steps[0].Key ''
        Assert-Equal $fixture.Target[''].Unique $true
    }
    Test-Case 'Inheritance reset is blocked if source effective grants include a missing principal' {
        $fixture = New-TestPlanFixture
        $fixture.Target['Docs/a.docx'].Unique = $true
        $missing = New-TestMap; $missing['visitors'] = 'missing'
        $plan = New-SPPermissionPlan $fixture.Source $fixture.Target $fixture.Matches $fixture.Desired $missing $roles @() @()
        Assert-Equal (@($plan.Steps | Where-Object Action -eq 'Reset').Count) 0
        Assert-Equal $fixture.Target['Docs/a.docx'].Unique $true
    }
    Test-Case 'Inherited moved object with an incompatible parent is rejected' {
        $fixture = New-TestPlanFixture
        $fixture.Target['Other'] = New-TestScope 'Other' $true @(New-TestAcl 'other') 'Library'
        $fixture.Target['Other/a.docx'] = New-TestScope 'Other/a.docx' $false
        $fixture.Matches['Docs/a.docx'] = 'Other/a.docx'
        Assert-Throws {
            New-SPPermissionPlan $fixture.Source $fixture.Target $fixture.Matches $fixture.Desired (New-TestMap) $roles @() @()
        } 'incompatible/unapproved parent'
    }
    Test-Case 'Failed addition halts apply before removal' {
        $fixture = New-TestPlanFixture
        $plan = New-SPPermissionPlan $fixture.Source $fixture.Target $fixture.Matches $fixture.Desired (New-TestMap) $roles @() @()
        $script:writeUrls = [Collections.Generic.List[string]]::new()
        function Assert-SPPermissionState {}
        function Invoke-SPPermissionRest {
            param($Connection, $Url, $Method)
            $script:writeUrls.Add($Url)
            throw 'MOCK addition failed'
        }
        $log = [Collections.Generic.List[object]]::new()
        Assert-Throws { Invoke-SPPermissionPlan 'target' $plan $log } 'MOCK addition failed'
        Assert-Equal $script:writeUrls.Count 1
        Assert-Equal ($script:writeUrls[0] -match '/addroleassignment') $true
        Assert-Equal $log[0].Result 'Failed - inspect live state'
    }
    Test-Case 'Failed dependent freeze prevents parent writes' {
        $fixture = New-TestPlanFixture
        $fixture.Target['Docs/new.docx'] = New-TestScope 'Docs/new.docx' $false
        $plan = New-SPPermissionPlan $fixture.Source $fixture.Target $fixture.Matches $fixture.Desired (New-TestMap) $roles @() @()
        $script:writes = 0
        function Assert-SPPermissionState {}
        function Invoke-SPPermissionRest { $script:writes++; throw 'MOCK freeze failed' }
        Assert-Throws { Invoke-SPPermissionPlan 'target' $plan ([Collections.Generic.List[object]]::new()) } 'freeze failed'
        Assert-Equal $script:writes 1
    }
    Test-Case 'Concurrent ACL changes stop apply before a write' {
        function Invoke-SPPermissionRest { [pscustomobject]@{ HasUniqueRoleAssignments = $true } }
        function Get-SPPermissionAcl { New-TestAcl 'concurrent-user' }
        $fixture = New-TestPlanFixture
        $plan = New-SPPermissionPlan $fixture.Source $fixture.Target $fixture.Matches $fixture.Desired (New-TestMap) $roles @() @()
        Assert-Throws { Assert-SPPermissionState 'target' $plan.Steps[0] $false } 'differs from the plan'
    }
    Test-Case 'Service and authorization errors are never classified as deleted principals' {
        Assert-Equal (Test-SPPrincipalMissingError 'The specified user old@example.test could not be found.') $true
        Assert-Equal (Test-SPPrincipalMissingError '403 access denied: specified user old@example.test could not be found') $false
        Assert-Equal (Test-SPPrincipalMissingError '503 service unavailable') $false
    }
    Test-Case 'Match CSV rejects duplicate targets, wrong kinds, conflicting exact paths and wrong webs' {
        $source = New-TestMap; $target = New-TestMap
        $source['Docs/a.docx'] = New-TestScope 'Docs/a.docx'
        $source['Docs/b.docx'] = New-TestScope 'Docs/b.docx'
        $target['Docs/new.docx'] = New-TestScope 'Docs/new.docx'
        $path = Join-Path $temp 'bad.csv'
        $row = [pscustomobject]@{
            Approved = 'Yes'; SourcePath = 'Docs/a.docx'; TargetPath = 'Docs/new.docx'
            SourceUrl = 'https://source.test'; TargetUrl = 'https://target.test'
        }
        @($row, ($row | Select-Object *)) | Export-Csv $path -NoTypeInformation
        Assert-Throws { Get-SPPermissionMatches $source $target $path 'https://source.test' 'https://target.test' } 'one-to-one'
        $row | Export-Csv $path -NoTypeInformation
        $target['Docs/new.docx'].Kind = 'Folder'
        Assert-Throws { Get-SPPermissionMatches $source $target $path 'https://source.test' 'https://target.test' } 'kind mismatch'
        $target['Docs/new.docx'].Kind = 'File'
        $target['Docs/a.docx'] = New-TestScope 'Docs/a.docx'
        Assert-Throws { Get-SPPermissionMatches $source $target $path 'https://source.test' 'https://target.test' } 'conflicts with an exact'
        Assert-Throws { Get-SPPermissionMatches $source $target $path 'https://wrong.test' 'https://target.test' } 'different source/target'
    }
    Test-Case 'CSV header is written even when there are no proposals' {
        $path = Join-Path $temp 'empty.csv'
        Export-SPPermissionCsv @() @('Approved','SourcePath','TargetPath') $path
        Assert-Equal (Get-Content -LiteralPath $path -Raw).Trim() '"Approved","SourcePath","TargetPath"'
    }
    Test-Case 'Approved aliases coexist with exact paths and never collide with the automatic regex Matches variable' {
        $source = New-TestMap; $target = New-TestMap
        $source[''] = New-TestScope ''; $target[''] = New-TestScope ''
        $source['Docs/a.docx'] = New-TestScope 'Docs/a.docx'
        $source['Docs/b.docx'] = New-TestScope 'Docs/b.docx'
        $target['Docs/new-a.docx'] = New-TestScope 'Docs/new-a.docx'
        $target['Docs/new-b.docx'] = New-TestScope 'Docs/new-b.docx'
        $path = Join-Path $temp 'approved.csv'
        @(
            [pscustomobject]@{ Approved = 'Yes'; SourcePath = 'Docs/a.docx'; TargetPath = 'Docs/new-a.docx'; SourceUrl = 'https://source.test'; TargetUrl = 'https://target.test' }
            [pscustomobject]@{ Approved = 'Yes'; SourcePath = 'Docs/b.docx'; TargetPath = 'Docs/new-b.docx'; SourceUrl = 'https://source.test'; TargetUrl = 'https://target.test' }
        ) | Export-Csv -LiteralPath $path -NoTypeInformation
        $result = Get-SPPermissionMatches $source $target $path 'https://source.test' 'https://target.test'
        Assert-Equal $result.Count 3
        Assert-Equal $result['Docs/a.docx'] 'Docs/new-a.docx'
        Assert-Equal $result['Docs/b.docx'] 'Docs/new-b.docx'
        Assert-Equal $result.ContainsKey('') $true
    }
    Test-Case 'Matcher names and word-order normalization are deterministic' {
        Assert-Equal (Get-SPNameSimilarity 'Patient Guide v2.docx' 'Guide Patient.docx') 1
        Assert-Equal (Get-SPNameRatio 'kitten' 'sitting') (1.0 - 3.0 / 7.0)
        Assert-Equal (Get-SPNameSimilarity 'copy.docx' 'final.docx') 0
    }
    Test-Case 'ACL reads preserve system bindings in snapshots but never synchronize them' {
        function Get-SPPermissionCollection {
            @(
                [pscustomobject]@{
                    Member = [pscustomobject]@{ Id = 1; LoginName = 'normal'; Title = 'normal'; PrincipalType = 1 }
                    RoleDefinitionBindings = @(
                        [pscustomobject]@{ Name = 'Read'; Id = 1073741826; Hidden = $false; RoleTypeKind = 2 }
                        [pscustomobject]@{ Name = 'Web-Only Limited Access'; Id = 4; Hidden = $false; RoleTypeKind = 0 }
                    )
                }
                [pscustomobject]@{
                    Member = [pscustomobject]@{ Id = 2; LoginName = 'SharingLinks.test'; Title = 'SharingLinks.test'; PrincipalType = 8 }
                    RoleDefinitionBindings = @([pscustomobject]@{ Name = 'Read'; Id = 1073741826; Hidden = $false; RoleTypeKind = 2 })
                }
            )
        }
        $acl = @(Get-SPPermissionAcl 'source' 'https://source.test/_api/web')
        Assert-Equal $acl.Count 2
        Assert-Equal $acl[0].Bindings.Count 2
        Assert-Equal ($acl[0].Roles -join ',') 'Read'
        Assert-Equal $acl[1].Skip $true
    }
    Test-Case 'Content fingerprints compute real SHA-256 and comparable word sets without disk downloads' {
        $bytes = [Text.Encoding]::UTF8.GetBytes('Patient guide for clinical staff')
        function Get-PnPFile {
            param($Url, [switch]$AsMemoryStream, $Connection, $ErrorAction)
            Assert-Equal $Connection 'source'
            [IO.MemoryStream]::new($bytes)
        }
        $item = New-TestScope 'Docs/guide.txt'
        $item.Ext = '.txt'; $item.Size = $bytes.Length
        $fingerprint = Get-SPContentFingerprint 'source' $item 1
        Assert-Equal $fingerprint.Hash ([Convert]::ToHexString([Security.Cryptography.SHA256]::HashData($bytes)))
        Assert-Equal $fingerprint.Words.Contains('clinical') $true
        $item.Size = 2MB
        $large = Get-SPContentFingerprint 'source' $item 1
        Assert-Equal ([bool]$large.Hash) $false
        Assert-Equal ($large.Error -match 'Larger than') $true
    }
    Test-Case 'Matcher end-to-end keeps inventories intact and all proposals unapproved' {
        $sourceFixture = New-TestMap; $targetFixture = New-TestMap
        foreach ($map in @($sourceFixture, $targetFixture)) {
            $map[''] = New-TestScope ''
            $map['Docs'] = New-TestScope 'Docs' $false @() 'Library'
        }
        $sourceFixture['Docs/Patient Guide.docx'] = New-TestScope 'Docs/Patient Guide.docx'
        $sourceFixture['Docs/Second Guide.docx'] = New-TestScope 'Docs/Second Guide.docx'
        $targetFixture['Docs/Patient Guide revised.docx'] = New-TestScope 'Docs/Patient Guide revised.docx'
        $targetFixture['Docs/Second Guide revised.docx'] = New-TestScope 'Docs/Second Guide revised.docx'
        function Import-Module {}
        function Connect-SPPermissionSite { param($Url) $Url }
        function Get-SPPermissionInventory {
            param($Connection)
            if ($Connection -eq 'https://source.test') { return ,$sourceFixture }
            return ,$targetFixture
        }
        $SourceUrl = 'https://source.test'; $TargetUrl = 'https://target.test'
        $OutputFolder = Join-Path $temp 'matcher'; $NoContent = $true
        Invoke-NewSPPermissionMatchMap
        $csv = Get-ChildItem -LiteralPath $OutputFolder -Filter 'PermMatch_Review_*.csv'
        $rows = @(Import-Csv -LiteralPath $csv.FullName)
        Assert-Equal $rows.Count 2
        Assert-Equal (@($rows | Where-Object Approved -ne 'No').Count) 0
        Assert-Equal $sourceFixture.Count 4
        Assert-Equal $targetFixture.Count 4
    }
    Test-Case 'Renamed library and translated descendants remain explicit unapproved CSV rows' {
        $sourceFixture = New-TestMap; $targetFixture = New-TestMap
        $sourceFixture[''] = New-TestScope ''; $targetFixture[''] = New-TestScope ''
        $sourceFixture['OldDocs'] = New-TestScope 'OldDocs' $false @() 'Library'
        $targetFixture['NewDocs'] = New-TestScope 'NewDocs' $false @() 'Library'
        $sourceFixture['OldDocs'].Title = 'Clinical guides'
        $targetFixture['NewDocs'].Title = 'Clinical guides'
        $sourceFixture['OldDocs/guide.docx'] = New-TestScope 'OldDocs/guide.docx'
        $sourceFixture['OldDocs/guide.docx'].ListKey = 'OldDocs'
        $targetFixture['NewDocs/guide.docx'] = New-TestScope 'NewDocs/guide.docx'
        $targetFixture['NewDocs/guide.docx'].ListKey = 'NewDocs'
        function Import-Module {}
        function Connect-SPPermissionSite { param($Url) $Url }
        function Get-SPPermissionInventory {
            param($Connection)
            if ($Connection -eq 'https://source.test') { return ,$sourceFixture }
            return ,$targetFixture
        }
        $SourceUrl = 'https://source.test'; $TargetUrl = 'https://target.test'
        $OutputFolder = Join-Path $temp 'matcher-renames'; $NoContent = $true
        Invoke-NewSPPermissionMatchMap
        $csv = Get-ChildItem -LiteralPath $OutputFolder -Filter 'PermMatch_Review_*.csv'
        $rows = @(Import-Csv -LiteralPath $csv.FullName)
        Assert-Equal $rows.Count 2
        Assert-Equal (@($rows | Where-Object Method -eq 'Translated parent').Count) 1
        Assert-Equal (@($rows | Where-Object Approved -ne 'No').Count) 0
    }
    Test-Case 'Failed source ACL read stops the whole sync without any principal or permission writes' {
        function Import-Module {}
        function Connect-SPPermissionSite { param($Url) $Url }
        function Get-SPPermissionInventory { return ,(New-TestPlanFixture).Source }
        function Get-SPPermissionAcl { throw 'MOCK source ACL denied' }
        function New-PnPUser { throw 'MUST NOT ensure any users' }
        function Invoke-SPPermissionPlan { throw 'MUST NOT apply' }
        $SourceUrl = 'https://source.test'; $TargetUrl = 'https://target.test'
        $OutputFolder = Join-Path $temp 'sync-failure'; $Apply = $true; $MatchFile = ''
        Assert-Throws { Invoke-SyncSPPermissionsFromSource } 'MOCK source ACL denied'
        $csv = Get-ChildItem -LiteralPath $OutputFolder -Filter 'PermSync_Log_*.csv'
        Assert-Equal (@(Import-Csv -LiteralPath $csv.FullName | Where-Object Result -eq 'Error').Count) 1
    }
    Test-Case 'Successful apply writes and verifies additions before removals on the target connection only' {
        $fixture = New-TestPlanFixture
        $plan = New-SPPermissionPlan $fixture.Source $fixture.Target $fixture.Matches $fixture.Desired (New-TestMap) $roles @() @()
        $script:remoteAcl = @(New-TestAcl 'old-user')
        $script:appliedMethods = [Collections.Generic.List[string]]::new()
        function Invoke-SPPermissionRest {
            param($Connection, $Url, $Method)
            Assert-Equal $Connection 'target-connection'
            if ($Method -eq 'Post') {
                if ($Url -match '/addroleassignment') {
                    $script:appliedMethods.Add('Add')
                    $script:remoteAcl += New-TestAcl 'visitors'
                } elseif ($Url -match '/removeroleassignment') {
                    $script:appliedMethods.Add('Remove')
                    $script:remoteAcl = @($script:remoteAcl | Where-Object LoginName -ne 'old-user')
                } else { throw "Unexpected write: $Url" }
                return
            }
            [pscustomobject]@{ HasUniqueRoleAssignments = $true }
        }
        function Get-SPPermissionAcl { $script:remoteAcl }
        $log = [Collections.Generic.List[object]]::new()
        Invoke-SPPermissionPlan 'target-connection' $plan $log
        Assert-Equal ($script:appliedMethods -join ',') 'Add,Remove'
        Assert-Equal (@($log | Where-Object Result -eq 'Verified').Count) 2
    }
    Test-Case 'Full sync dry run persists plan and snapshots without EnsureUser or permission writes' {
        $fixture = New-TestPlanFixture
        $fixture.Source[''].Acl = @(New-TestAcl 'visitors' @('Read') 12 8)
        function Import-Module {}
        function Connect-SPPermissionSite { param($Url) $Url }
        function Get-SPPermissionInventory {
            param($Connection)
            if ($Connection -eq 'https://source.test') { return ,$fixture.Source }
            return ,$fixture.Target
        }
        function Get-SPPermissionAcl {
            param($Connection)
            if ($Connection -eq 'https://source.test') { New-TestAcl 'visitors' @('Read') 12 8 }
            else { New-TestAcl 'old-user' }
        }
        function Get-SPPermissionCollection {
            param($Connection, $Url)
            if ($Url -match '/sitegroups') { return [pscustomobject]@{ Id = 12; Title = 'visitors'; LoginName = 'visitors' } }
            if ($Url -match '/roledefinitions') {
                return [pscustomobject]@{ Id = 1073741826; Name = 'Read'; BasePermissions = [pscustomobject]@{ High = '0'; Low = '1' } }
            }
            if ($Url -match '/siteusers') { return }
            throw "Unexpected collection: $Url"
        }
        function Invoke-SPPermissionRest {
            param($Connection, $Url)
            if ($Url -match 'AssociatedVisitorGroup') { return [pscustomobject]@{ Id = 12; Title = 'visitors'; LoginName = 'visitors' } }
            if ($Url -match 'Associated(Owner|Member)Group') { return [pscustomobject]@{ LoginName = '' } }
            throw "Unexpected REST request: $Url"
        }
        function New-PnPUser { throw 'Dry run must not EnsureUser.' }
        function Invoke-SPPermissionPlan { throw 'Dry run must not write permissions.' }
        $SourceUrl = 'https://source.test'; $TargetUrl = 'https://target.test'
        $OutputFolder = Join-Path $temp 'sync-success'; $Apply = $false; $MatchFile = ''
        Invoke-SyncSPPermissionsFromSource
        $jsonFiles = @(Get-ChildItem -LiteralPath $OutputFolder -Filter '*.json')
        Assert-Equal $jsonFiles.Count 3
        $planFile = Get-ChildItem -LiteralPath $OutputFolder -Filter 'PermSync_Plan_*.json'
        $savedPlan = Get-Content -LiteralPath $planFile.FullName -Raw | ConvertFrom-Json
        Assert-Equal $savedPlan.Steps.Count 2
        Assert-Equal $savedPlan.Steps[0].Action 'Add'
        Assert-Equal $savedPlan.Steps[1].Action 'Remove'
        $beforeFile = Get-ChildItem -LiteralPath $OutputFolder -Filter 'PermSync_TargetSnapshot_*.json'
        $before = Get-Content -LiteralPath $beforeFile.FullName -Raw | ConvertFrom-Json
        Assert-Equal (@($before.Scopes | Where-Object Key -eq '')[0].Acl[0].LoginName) 'old-user'
    }
    Write-Host "All $script:passed offline tests passed. No PnP authentication or live SharePoint requests were made."
} finally {
    Remove-Item -LiteralPath $temp -Recurse -Force
}
