#requires -Version 7.4.6
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'Invoke-SPPermissionQueue.ps1')
. (Join-Path $PSScriptRoot 'Export-SPPermissionSiteReport.ps1') -TargetUrl 'https://target.test/sites/a' -OutputFolder 'unused-offline'
$temp = Join-Path ([IO.Path]::GetTempPath()) ('SPQueueTest-' + [guid]::NewGuid())
[void][IO.Directory]::CreateDirectory($temp)
$passed = 0
function Assert-Queue($Condition, [string]$Message) {
    if (-not $Condition) { throw $Message }
}
function Test-Queue([string]$Name, [scriptblock]$Action) {
    & $Action
    $script:passed++
    Write-Host "PASS $Name"
}
function Import-Module {}
function Connect-PnPOnline { throw 'Offline test must never authenticate.' }
try {
    Test-Queue '80 percent boundary, lower-score review and renamed-parent dependencies' {
        $path = Join-Path $temp 'MatchReview.csv'
        @(
            [pscustomobject]@{ Approved = 'No'; Score = '.799'; Kind = 'Library'; SourcePath = 'Old'; TargetPath = 'New' }
            [pscustomobject]@{ Approved = 'No'; Score = '1'; Kind = 'File'; SourcePath = 'Old/a'; TargetPath = 'New/a' }
            [pscustomobject]@{ Approved = 'No'; Score = '.80'; Kind = 'Item'; SourcePath = 'Docs/a'; TargetPath = 'Docs/b' }
            [pscustomobject]@{ Approved = 'Yes'; Score = '.6'; Kind = 'Item'; SourcePath = 'Docs/c'; TargetPath = 'Docs/d' }
        ) | Export-Csv -LiteralPath $path -NoTypeInformation
        Set-SPChecklistAutoApprovals $path
        $rows = @(Import-Csv -LiteralPath $path)
        Assert-Queue ($rows[0].Approved -eq 'No') 'Below threshold auto-approved.'
        Assert-Queue ($rows[1].Approved -eq 'No') 'Child bypassed unapproved parent.'
        Assert-Queue ($rows[2].Approved -eq 'Yes') 'Exact threshold not approved.'
        Assert-Queue ($rows[3].Approved -eq 'Yes') 'Manual approval lost.'
    }
    Test-Queue 'Dry-run gate refuses unavailable principals or blocked plans' {
        $folder = Join-Path $temp 'preflight'
        [void][IO.Directory]::CreateDirectory($folder)
        $planPath = Join-Path $folder 'PermSync_Plan_DRYRUN_test.json'
        @{ Steps = @(); Notes = @() } | ConvertTo-Json | Set-Content -LiteralPath $planPath
        $csv = Join-Path $folder 'PermSync_UnavailablePrincipals_DRYRUN_test.csv'
        '"Principal"' | Set-Content -LiteralPath $csv
        Assert-SPChecklistDryRun $folder
        @{ Steps = @(); Notes = @(@{ Result = 'Blocked' }) } | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $planPath
        $threw = $false
        try { Assert-SPChecklistDryRun $folder } catch { $threw = $_.Exception.Message -match 'blocked' }
        Assert-Queue $threw 'Blocked plan allowed Apply.'
        @{ Steps = @(); Notes = @() } | ConvertTo-Json | Set-Content -LiteralPath $planPath
        '"missing"' | Add-Content -LiteralPath $csv
        $threw = $false
        try { Assert-SPChecklistDryRun $folder } catch { $threw = $_.Exception.Message -match 'unavailable' }
        Assert-Queue $threw 'Unavailable principal allowed Apply.'
    }
    Test-Queue 'Related target collections serialize without serializing all archived sources' {
        $active = @([pscustomobject]@{ SourceUrl = 'https://source.test/sites/archive/A'; TargetUrl = 'https://target.test/sites/A' })
        $candidate = [pscustomobject]@{ SourceUrl = 'https://source.test/sites/archive/B'; TargetUrl = 'https://target.test/sites/B' }
        Assert-Queue (-not (Test-SPQueueConflict $candidate $active)) 'Independent target sites were serialized.'
        $candidate.TargetUrl = 'https://target.test/sites/A/child'
        Assert-Queue (Test-SPQueueConflict $candidate $active) 'Related targets allowed concurrent writes.'
    }
    Test-Queue 'CS intentional libraries are explicitly excluded from permission changes' {
        $ChecklistFile = Join-Path $temp 'cs-exclusions.csv'
        [pscustomobject]@{
            'Old Site URL' = 'https://source.test/sites/archive/CS'
            'New Site URL' = 'https://whallc.sharepoint.com/sites/ClinicalandPatientServices'
        } | Export-Csv -LiteralPath $ChecklistFile -NoTypeInformation
        function Connect-SPPermissionSite {}
        function Get-SPChecklistChildren {}
        function Invoke-SPChecklistWeb {
            param($Pair, $RunFolder, $SiteReviewFolder, $Options, $RunMode)
            Assert-Queue ('COVID19 Resources' -in $Options.ExcludePaths) 'Resources library not excluded.'
            Assert-Queue ('COVID19 Workflows' -in $Options.ExcludePaths) 'Workflows library not excluded.'
        }
        $Mode = 'DryRun'; $TargetSites = @()
        $OutputFolder = Join-Path $temp 'cs-exclusions-output'; $ReviewRoot = Join-Path $temp 'cs-exclusions-reviews'
        Invoke-SPPermissionChecklist
    }
    Test-Queue 'Recursive report includes inherited, unique and target-only content, root before child' {
        $script:reportOrder = [Collections.Generic.List[string]]::new()
        function Connect-SPPermissionSite { param($Url) $script:reportOrder.Add($Url); $Url }
        function Get-SPPermissionCollection {
            [pscustomobject]@{ Id = 2; Name = 'Read'; BasePermissions = @{ High = '0'; Low = '1' } }
            [pscustomobject]@{ Id = 3; Name = 'Limited Access'; BasePermissions = @{ High = '0'; Low = '0' } }
        }
        function Get-SPMigrationProtection {
            [pscustomobject]@{
                Logins = [Collections.Generic.HashSet[string]]::new([string[]]@('visitors'))
                Memberships = @([pscustomobject]@{ GroupLogin = 'visitors'; MigrationPrincipal = 'SharePoint.Migration.Team.Visitors' })
            }
        }
        function Get-SPPermissionInventory {
            param($Connection)
            $inventory = [Collections.Generic.Dictionary[string, object]]::new()
            foreach ($entry in @(@('', $true), @('Docs', $true), @('Docs/inherited.docx', $false), @('TargetOnly', $false))) {
                $key = $entry[0]
                $inventory[$key] = [pscustomobject]@{
                    Key = $key; Unique = $entry[1]; Kind = $(if ($key) { 'File' } else { 'Web' })
                    Api = "$Connection/_api/$key"; Acl = $null; Depth = ($key -split '/').Count
                }
            }
            if ($Connection -notlike '*/child') {
                $inventory['child'] = [pscustomobject]@{ Key = 'child'; Kind = 'Subweb'; Url = "$Connection/child" }
            }
            return ,$inventory
        }
        function Get-SPPermissionAcl {
            param($Connection, $Api)
            [pscustomobject]@{
                LoginName = 'visitors'; Title = 'Visitors'; PrincipalType = 8; MemberId = 1; Skip = $false
                Bindings = @(
                    @{ Name = 'Read'; Id = 2; Hidden = $false; RoleTypeKind = 2; BasePermissions = @{ High = '0'; Low = '1' } }
                    @{ Name = 'Limited Access'; Id = 3; Hidden = $true; RoleTypeKind = 1 }
                )
            }
        }
        $counts = @{ Webs = 0; Objects = 0; Rows = 0 }
        $rows = @(Get-SPPermissionReportRows 'https://target.test/sites/A' @{} 500 $counts)
        Assert-Queue ($counts.Webs -eq 2 -and $counts.Objects -eq 8 -and $rows.Count -eq 16) 'Incomplete report coverage.'
        Assert-Queue (($script:reportOrder -join ',') -eq 'https://target.test/sites/A,https://target.test/sites/A/child') 'Traversal order incorrect.'
        $inherited = @($rows | Where-Object { $_.Path -eq 'Docs/inherited.docx' -and $_.WebUrl -eq 'https://target.test/sites/A' })
        Assert-Queue ($inherited[0].InheritedFrom -eq 'https://target.test/sites/A/Docs') 'Wrong inheritance origin.'
        Assert-Queue (@($rows | Where-Object Path -eq 'TargetOnly').Count -eq 4) 'Target-only scope omitted.'
        Assert-Queue (@($rows | Where-Object Role -eq 'Limited Access').Count -eq 8) 'Hidden role bindings omitted.'
        Assert-Queue ($rows[0].MigrationPolicy -like 'Intentionally preserved*') 'Migration policy missing.'
        Assert-Queue ($rows[0].MigrationMembers -eq 'SharePoint.Migration.Team.Visitors') 'Containing group membership missing.'
    }
    Test-Queue 'Report failure persists Failed manifest and never success-shaped output' {
        function Get-SPPermissionReportRows { throw 'MOCK unreadable ACL' }
        $TargetUrl = 'https://target.test/sites/A'; $OutputFolder = Join-Path $temp 'failed-report'
        $failed = $false
        try { Invoke-SPPermissionSiteReport } catch { $failed = $_.Exception.Message -match 'unreadable ACL' }
        $manifest = Get-Content -LiteralPath (Join-Path $OutputFolder 'ReportStatus.json') -Raw | ConvertFrom-Json
        Assert-Queue ($failed -and $manifest.Status -eq 'Failed') 'Failed report claimed success.'
    }
    Test-Queue 'Scheduler never exceeds four active workers and retains completed windows' {
        $ChecklistFile = Join-Path $temp 'queue.csv'
        1..7 | ForEach-Object {
            [pscustomobject]@{ 'Old Site URL' = "https://source.test/sites/archive/S$_"; 'New Site URL' = "https://target.test/sites/S$_" }
        } | Export-Csv -LiteralPath $ChecklistFile -NoTypeInformation
        $script:runningFolders = [Collections.Generic.List[string]]::new()
        $script:peak = 0; $script:launched = 0
        function Start-Process {
            param($FilePath, $Verb, $ArgumentList, [switch]$PassThru, $ErrorAction)
            Assert-Queue ($Verb -eq 'RunAs' -and '-NoExit' -in $ArgumentList) 'Worker not elevated/retained.'
            $command = [Text.Encoding]::Unicode.GetString([Convert]::FromBase64String($ArgumentList[-1]))
            $jobPath = [regex]::Match($command, "-WorkerJobFile '([^']+)'").Groups[1].Value
            $job = Import-Clixml -LiteralPath $jobPath
            $script:runningFolders.Add($job.OutputFolder)
            $script:peak = [math]::Max($script:peak, $script:runningFolders.Count)
            $script:launched++
            [pscustomobject]@{ HasExited = $false }
        }
        function Start-Sleep {
            foreach ($folder in $script:runningFolders) {
                @{ TargetUrl = $folder; Status = 'Completed'; Error = '' } | ConvertTo-Json |
                    Set-Content -LiteralPath (Join-Path $folder 'WorkerStatus.json')
            }
            $script:runningFolders.Clear()
        }
        $Apply = $true; $MaxConcurrent = 4; $TargetSites = @()
        $OutputFolder = Join-Path $temp 'queue-output'; $ReviewRoot = Join-Path $temp 'queue-reviews'
        Invoke-SPPermissionQueue
        Assert-Queue ($script:peak -eq 4 -and $script:launched -eq 7) 'Queue concurrency/completion incorrect.'
        $csv = Get-ChildItem -LiteralPath $OutputFolder -Recurse -Filter QueueStatus.csv
        Assert-Queue (@(Import-Csv -LiteralPath $csv.FullName | Where-Object Status -eq 'Completed').Count -eq 7) 'Missing worker outcomes.'
    }
    Test-Queue 'Default queue report paths stay below 240 characters at both repository locations' {
        $roots = @(
            'C:\Users\!bposey\Desktop\GitHub\SharePoint-Revamp'
            [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
        )
        foreach ($root in $roots) {
            $relative = 'Reports\PermRuns\q-261019143547347\E86E6768733F\w-E86E6768733F\DryRun\PermSync_UnavailablePrincipals_DRYRUN_20261019-143547-347.csv'
            Assert-Queue ((Join-Path $root $relative).Length -le 240) "Report path exceeds 240: $root"
        }
    }
    Test-Queue 'Archive preview is read-only; Apply preserves hashes and supports resuming' {
        $repoRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
        $source = Join-Path $repoRoot ('Reports\path-test-' + [guid]::NewGuid().ToString('N'))
        $archive = Join-Path $temp 'archive'
        [void][IO.Directory]::CreateDirectory($source)
        try {
            $file = Join-Path $source 'example.csv'
            'test,evidence' | Set-Content -LiteralPath $file
            $hash = (Get-FileHash -LiteralPath $file).Hash
            $scriptPath = Join-Path $PSScriptRoot 'Move-SPPermissionReportsToShortPaths.ps1'
            & $scriptPath -SourceFolder $source -ArchiveRoot $archive | Out-Null
            Assert-Queue (Test-Path -LiteralPath $file) 'Preview moved source.'
            Assert-Queue (-not (Test-Path -LiteralPath $archive)) 'Preview wrote archive.'
            & $scriptPath -SourceFolder $source -ArchiveRoot $archive -Apply
            $manifest = Get-ChildItem -LiteralPath $archive -Recurse -Filter Paths.csv
            $row = Import-Csv -LiteralPath $manifest.FullName
            Assert-Queue ((Get-FileHash -LiteralPath $row.ArchivePath).Hash -eq $hash) 'Archive lost contents.'
            Assert-Queue ($row.OriginalPath -eq $file) 'Original mapping lost.'
            Assert-Queue (-not (Test-Path -LiteralPath $file)) 'Source not relocated.'
            & $scriptPath -SourceFolder $source -ArchiveRoot $archive -Apply
        } finally {
            Remove-Item -LiteralPath $source -Recurse -Force
        }
    }
    Write-Host "All $passed queue/report offline tests passed. No windows or live connections were opened."
} finally {
    Remove-Item -LiteralPath $temp -Recurse -Force
}
