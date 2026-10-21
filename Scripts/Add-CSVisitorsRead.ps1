#requires -Version 7.4.6
<#
.SYNOPSIS
    Additive-only source-group Read -> associated target Visitors Read; CS defaults.
.DESCRIPTION
    DryRun by default. Never removes roles, resets inheritance, creates permission
    levels or resolves old source users. Only exact paths and approved MatchFile
    aliases are used. Matched inherited source access is included, so unique target
    children do not miss intended Read access. Unmatched/excluded target content
    is protected with break-and-copy before parent changes. EEEU membership is
    planned, or ensured with -Apply, only in the associated target Visitors group.
#>
[CmdletBinding()]
param(
    [string]$SourceUrl = 'https://whallc.sharepoint.com/sites/archive-2026-10-18T115436Z/CS',
    [string]$TargetUrl = 'https://whallc.sharepoint.com/sites/ClinicalandPatientServices',
    [string]$ExpectedVisitorsTitle = 'Clinical & Patient Services Visitors',
    [string]$SourceReadGroup = 'Team Site Members',
    [string[]]$ExcludePaths = @('Access Requests', 'COVID19 Resources', 'COVID19 Workflows'),
    [switch]$NoMembershipWithoutSourceRead,
    [switch]$SkipFinalReport,
    [string]$OutputFolder = (Join-Path $PSScriptRoot '..\Reports\PermRead'),
    [string]$MatchFile,
    [string]$ClientId = '2e20d82f-56ba-418e-b465-9718493f0df8',
    [string]$Tenant,
    [switch]$SkipPersistLoginSetup,
    [switch]$Apply
)
. (Join-Path $PSScriptRoot 'SPOutputLimits.Common.ps1')
. (Join-Path $PSScriptRoot 'Sync-SPPermissionsFromSource.ps1') `
    -SourceUrl $SourceUrl -TargetUrl $TargetUrl -ExcludePaths $ExcludePaths `
    -OutputFolder $OutputFolder -MatchFile $MatchFile -ClientId $ClientId -Tenant $Tenant `
    -SkipPersistLoginSetup:$SkipPersistLoginSetup -Apply:$Apply

function New-CSVisitorsReadPlan {
    param($Source, $Target, $Pairs, $Visitors, [int]$ReadId,
        [string]$ReadGroup = 'Team Site Members',
        [string[]]$Excluded = @('Access Requests', 'COVID19 Resources', 'COVID19 Workflows'))
    $eligible = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $coverage = [Collections.Generic.List[object]]::new()
    $steps = [Collections.Generic.List[object]]::new()
    $notes = [Collections.Generic.List[object]]::new()
    $exclusions = $Excluded
    foreach ($key in $Source.Keys) {
        if ($Source[$key].Kind -eq 'Subweb') { continue }
        $hasRead = @((Get-SPPermissionEffectiveScope $Source $key).Acl | Where-Object {
            $_.PrincipalType -eq 8 -and $_.Title -ieq $ReadGroup -and $_.Roles -icontains 'Read'
        }).Count -gt 0
        if (-not $hasRead) { continue }
        $state = 'Unmatched - skipped'
        $targetKey = ''
        if ($Pairs.ContainsKey($key)) {
            $targetKey = $Pairs[$key]
            if ((Test-SPPermissionExcluded $key $exclusions) -or (Test-SPPermissionExcluded $targetKey $exclusions)) {
                $state = 'Intentionally excluded'
            } else {
                [void]$eligible.Add($targetKey)
                $state = 'Eligible'
            }
        }
        $coverage.Add([pscustomobject]@{ SourcePath = $key; SourceUnique = $Source[$key].Unique; TargetPath = $targetKey; Status = $state })
    }
    foreach ($key in @($eligible | Sort-Object { ($_ -split '/').Count }, { $_ })) {
        $scope = $Target[$key]
        $current = @((Get-SPPermissionEffectiveScope $Target $key).Acl)
        if (@($current | Where-Object { $_.MemberId -eq $Visitors.Id -and $_.Roles -icontains 'Read' }).Count) { continue }
        if (-not $scope.Unique) {
            $steps.Add((New-SPPermissionStep $scope 'Break' $current $true $current $null))
            $scope.Unique = $true; $scope.Acl = $current
        }
        Add-SPDependentProtection $Target $key $eligible $steps $notes
        $desired = [Collections.Generic.Dictionary[string, object]]::new([StringComparer]::OrdinalIgnoreCase)
        $desired[$Visitors.LoginName] = [pscustomobject]@{ Id = $Visitors.Id; Title = $Visitors.Title; PType = 8 }
        $operation = [pscustomobject]@{
            Type = 'Add'; Login = $Visitors.LoginName; Role = 'Read'; RoleId = $ReadId; MemberId = $Visitors.Id
        }
        $after = Get-SPProjectedAcl $current $operation $desired
        $steps.Add((New-SPPermissionStep $scope 'Add' $current $true $after $operation))
        $scope.Acl = $after
    }
    [pscustomobject]@{ Steps = $steps.ToArray(); Notes = $notes.ToArray(); Coverage = $coverage.ToArray() }
}

function Invoke-CSVisitorsReadReport {
    param([hashtable]$Options)
    & (Join-Path $PSScriptRoot 'Export-SPPermissionSiteReport.ps1') @Options
}

function Assert-CSVisitorsGroup {
    param($Visitors, [string]$ExpectedTitle = 'Clinical & Patient Services Visitors')
    if (-not $Visitors.Id -or -not $Visitors.LoginName -or
        $Visitors.Title -ine $ExpectedTitle) {
        throw "Expected associated group '$ExpectedTitle'; received '$($Visitors.Title)' (ID $($Visitors.Id)). No changes permitted."
    }
    if (Test-SPMigrationPrincipal $Visitors.LoginName $Visitors.Title) {
        throw 'The selected Visitors principal is a protected migration group. No changes permitted.'
    }
}

function Invoke-CSVisitorsRead {
    $ErrorActionPreference = 'Stop'
    Import-Module PnP.PowerShell -MinimumVersion 3.3.0 -ErrorAction Stop
    Assert-SPPermissionUrls $SourceUrl $TargetUrl
    $sourceUrl = $SourceUrl.TrimEnd('/')
    $targetUrl = $TargetUrl.TrimEnd('/')
    $root = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputFolder)
    $folder = Join-Path $root ('rd-' + (Get-Date -Format 'yyMMddHHmmssfff'))
    [void](Assert-SPOutputPath $folder)
    [void][IO.Directory]::CreateDirectory($folder)
    $log = [Collections.Generic.List[object]]::new()
    $logPath = Join-Path $folder 'Changes.csv'
    $status = [ordered]@{ Mode = $(if ($Apply) { 'Apply' } else { 'DryRun' }); Status = 'Running'; Error = '' }
    try {
        $sourceConnection = Connect-SPPermissionSite $sourceUrl $ClientId $Tenant -SkipPersistLoginSetup:$SkipPersistLoginSetup
        $targetConnection = Connect-SPPermissionSite $targetUrl $ClientId $Tenant -SkipPersistLoginSetup:$SkipPersistLoginSetup
        Write-Host "[Inventory source] $sourceUrl"
        $source = Get-SPPermissionInventory $sourceConnection $sourceUrl
        Write-Host "[Inventory target] $targetUrl"
        $target = Get-SPPermissionInventory $targetConnection $targetUrl
        foreach ($side in @('Source', 'Target')) {
            $inventory = if ($side -eq 'Source') { $source } else { $target }
            $connection = if ($side -eq 'Source') { $sourceConnection } else { $targetConnection }
            Write-Host "[Snapshot $side] Reading unique-scope role assignments."
            foreach ($scope in $inventory.Values | Where-Object { $_.Unique -or -not $_.Key }) {
                $api = $scope.Api
                if (-not $scope.Unique) { $api += '/FirstUniqueAncestorSecurableObject' }
                $scope.Acl = @(Get-SPPermissionAcl $connection $api)
            }
            $snapshotRows = @(
                foreach ($scope in $inventory.Values | Sort-Object Key) {
                    [pscustomobject]@{
                        WebUrl = if ($side -eq 'Source') { $sourceUrl } else { $targetUrl }
                        CapturedUtc = [datetime]::UtcNow.ToString('o')
                        ScopePath = $scope.Key
                        Kind = $scope.Kind
                        Unique = $scope.Unique
                        ScopeJson = ConvertTo-Json -InputObject $scope -Depth 50 -Compress
                    }
                }
            )
            Export-SPPermissionCsv $snapshotRows @('WebUrl','CapturedUtc','ScopePath','Kind','Unique','ScopeJson') (
                Join-Path $folder "${side}Scopes.csv"
            )
        }
        $visitors = Invoke-SPPermissionRest $targetConnection ($targetUrl + '/_api/web/AssociatedVisitorGroup?$select=Id,Title,LoginName')
        Assert-CSVisitorsGroup $visitors $ExpectedVisitorsTitle
        $protection = Get-SPMigrationProtection $targetConnection $targetUrl
        foreach ($membership in $protection.Memberships | Where-Object { $_.GroupId -eq $visitors.Id }) {
            $log.Add([pscustomobject]@{
                Path = ''; Action = 'Authorized site Visitors container'; Result = 'By design'
                Detail = "$($visitors.Title) contains $($membership.MigrationPrincipal). Added container access is authorized; direct migration-group assignments/memberships remain unchanged."
            })
        }
        $read = @(Get-SPPermissionCollection $targetConnection ($targetUrl + '/_api/web/roledefinitions?$select=Id,Name,BasePermissions') |
            Where-Object Name -IEQ 'Read')
        $sourceRead = @(Get-SPPermissionCollection $sourceConnection ($sourceUrl + '/_api/web/roledefinitions?$select=Id,Name,BasePermissions') |
            Where-Object Name -IEQ 'Read')
        if ($read.Count -ne 1 -or $sourceRead.Count -ne 1 -or $read[0].Id -le 0 -or
            $null -eq $read[0].BasePermissions.High -or $null -eq $read[0].BasePermissions.Low -or
            [string]$read[0].BasePermissions.High -cne [string]$sourceRead[0].BasePermissions.High -or
            [string]$read[0].BasePermissions.Low -cne [string]$sourceRead[0].BasePermissions.Low) {
            throw 'Read permission levels are missing or have different/unknown masks.'
        }
        $pairs = Get-SPPermissionMatches $source $target $MatchFile $sourceUrl $targetUrl
        $plan = New-CSVisitorsReadPlan $source $target $pairs $visitors $read[0].Id $SourceReadGroup $ExcludePaths
        $stepRows = @(
            for ($index = 0; $index -lt $plan.Steps.Count; $index++) {
                $step = $plan.Steps[$index]
                [pscustomobject]@{
                    Sequence = $index + 1; Key = $step.Key; Action = $step.Action
                    OperationLogin = $step.Operation.Login; OperationRole = $step.Operation.Role
                    StepJson = ConvertTo-Json -InputObject $step -Depth 40 -Compress
                }
            }
        )
        $noteRows = @(
            for ($index = 0; $index -lt $plan.Notes.Count; $index++) {
                $note = $plan.Notes[$index]
                [pscustomobject]@{
                    Sequence = $index + 1; Result = $note.Result
                    NoteJson = ConvertTo-Json -InputObject $note -Depth 40 -Compress
                }
            }
        )
        Export-SPPermissionCsv $stepRows @('Sequence','Key','Action','OperationLogin','OperationRole','StepJson') (
            Join-Path $folder 'PlanSteps.csv'
        )
        Export-SPPermissionCsv $noteRows @('Sequence','Result','NoteJson') (Join-Path $folder 'PlanNotes.csv')
        Export-SPPermissionCsv $plan.Coverage @('SourcePath','SourceUnique','TargetPath','Status') (Join-Path $folder 'Coverage.csv')
        $tenantId = [guid](Get-PnPTenantId -Connection $targetConnection -ErrorAction Stop)
        if ($tenantId -eq [guid]::Empty) { throw 'Tenant ID unavailable for EEEU claim.' }
        $eeeu = "c:0-.f|rolemanager|spo-grid-all-users/$tenantId"
        $membershipUrl = "$targetUrl/_api/web/sitegroups/getbyid($($visitors.Id))/users?`$select=Id,Title,LoginName"
        $members = @(Get-SPPermissionCollection $targetConnection $membershipUrl)
        $hasEEEU = @($members | Where-Object LoginName -IEQ $eeeu).Count -gt 0
        $membershipAllowed = -not $NoMembershipWithoutSourceRead -or $plan.Coverage.Count -gt 0
        Write-Host "[Plan] Source Read scopes: $($plan.Coverage.Count); permission operations: $($plan.Steps.Count); EEEU membership changes allowed: $membershipAllowed"
        $status['SourceReadScopes'] = $plan.Coverage.Count
        $status['PlannedOperations'] = $plan.Steps.Count
        $log.Add([pscustomobject]@{
            Path = ''; Action = 'EEEU in Visitors'
            Result = $(if ($hasEEEU) { 'Already present' } elseif (-not $membershipAllowed) { 'Report only - zero source Read scopes' } else { 'Planned' })
            Detail = $eeeu
        })
        Export-SPPermissionCsv $log.ToArray() @('Path','Action','Result','Detail') $logPath
        if ($Apply) {
            if (-not $hasEEEU -and $membershipAllowed) {
                $log.Add([pscustomobject]@{ Path = ''; Action = 'Add EEEU to Visitors'; Result = 'Write starting - outcome not yet verified'; Detail = $eeeu })
                Export-SPPermissionCsv $log.ToArray() @('Path','Action','Result','Detail') $logPath
                $null = Invoke-SPPermissionRest $targetConnection "$targetUrl/_api/web/sitegroups/getbyid($($visitors.Id))/users" -Method Post `
                    -Content @{ LoginName = $eeeu }
                if (-not @(Get-SPPermissionCollection $targetConnection $membershipUrl | Where-Object LoginName -IEQ $eeeu).Count) {
                    throw 'EEEU membership write could not be verified; inspect live group membership.'
                }
                $log.Add([pscustomobject]@{ Path = ''; Action = 'Add EEEU to Visitors'; Result = 'Verified'; Detail = $eeeu })
            }
            Invoke-SPPermissionPlan $targetConnection $plan $log $logPath
        }
        $status.Status = if ($Apply) { 'Applied - matched additions verified; inspect coverage for skipped content' } else { 'DryRun complete - no writes' }
        Write-Host "$($status.Status). Planned permission operations: $($plan.Steps.Count)."
    } catch {
        $status.Status = 'Failed'; $status.Error = $_.Exception.Message
        throw
    } finally {
        Export-SPPermissionCsv $log.ToArray() @('Path','Action','Result','Detail') $logPath
        $status | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $folder 'Status.json') -Encoding utf8
        Write-Host "Additive Read evidence: $folder"
    }
    if ($Apply -and -not $SkipFinalReport) {
        Invoke-CSVisitorsReadReport @{
            TargetUrl = $targetUrl; OutputFolder = (Join-Path $folder 'Report')
            ClientId = $ClientId; Tenant = $Tenant; SkipPersistLoginSetup = $SkipPersistLoginSetup
        }
    }
}

if ($MyInvocation.InvocationName -ne '.') { Invoke-CSVisitorsRead }
