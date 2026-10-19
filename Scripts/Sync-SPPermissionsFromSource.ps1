#requires -Version 7.4.6
<#
.SYNOPSIS
    Plan, or explicitly apply, source-to-target SharePoint permission synchronization.
.DESCRIPTION
    Dry run by default, but both modes READ the live source and target. The source is
    never written. Explicit PnP connections are used throughout. Copy the adjacent
    SPPermission.Common.ps1 with this script to your execution server.

    Reads complete inventories and ACLs and saves snapshots before planning changes.
    Exact paths are matched first; renamed objects require an approved match CSV.
    Source Team Site Members maps to target associated Visitors; other SharePoint
    groups map by title, and users/Entra groups map by login claim.

    The effective ACL of an inherited SOURCE ROOT is read from its first unique
    ancestor and mirrored onto the target root as unique permissions. Only at that
    root, target associated Owners/Members grants are preserved. Use the checklist
    runner to process matched subsites individually. Group membership synchronization, permission-level creation and sharing links
    are NOT supported. Limited Access and sharing-link assignments are not edited.

    Target-only and excluded content that inherits a changing scope is frozen first,
    copying current grants. -NoFreeze explicitly permits it to follow parent changes.
    Approved inherited objects must have matching structural parents; otherwise the
    complete plan is rejected rather than silently giving moved objects different access.

    Missing principals are reported once. Available grants can still be added, but
    removals/resets on a scope depending on an unavailable principal are blocked.
    A picker miss in dry run is only a prediction; Apply resolves via EnsureUser.
    Authentication/service/authorization errors are fatal, not "deleted principal".

    Apply pre-resolves principals (which may add entries to target Site Users),
    then rechecks and verifies every permission write. A failed/ambiguous write stops
    the run BEFORE any later removals. Earlier writes remain: this is not a transaction.
    JSON snapshots are rollback REFERENCES, not an automatic restore script.
.EXAMPLE
    .\Sync-SPPermissionsFromSource.ps1 -Tenant '<tenant-guid-or-domain>' -MatchFile .\review.csv
.EXAMPLE
    .\Sync-SPPermissionsFromSource.ps1 -Tenant '<tenant-guid-or-domain>' -MatchFile .\review.csv -Apply
#>
[CmdletBinding()]
param(
    [string]$SourceUrl = 'https://whallc.sharepoint.com/sites/archive-2026-10-18T115436Z/CS',
    [string]$TargetUrl = 'https://whallc.sharepoint.com/sites/ClinicalandPatientServices',
    [string]$SourceGroupToMap = 'Team Site Members',
    [string]$ClientId = '2e20d82f-56ba-418e-b465-9718493f0df8',
    [string]$Tenant,
    [string]$OutputFolder = (Join-Path $PSScriptRoot '..\Outputs\PermissionRuns'),
    [string[]]$ExcludePaths = @('Access Requests'),
    [string]$MatchFile,
    [ValidateRange(1, 5000)][int]$PageSize = 500,
    [switch]$NoFreeze,
    [switch]$SkipPersistLoginSetup,
    [switch]$Apply
)
. (Join-Path $PSScriptRoot 'SPPermission.Common.ps1')

function Get-SPDesiredAcl {
    param($Acl, $GroupsByTitle, $Visitors, [string]$GroupToMap)
    $desired = [Collections.Generic.Dictionary[string, object]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($assignment in @($Acl) | Where-Object { -not $_.Skip }) {
        $login = $assignment.LoginName
        $title = $assignment.Title
        $id = 0
        if ($assignment.PrincipalType -eq 8) {
            $title = if ($title -ieq $GroupToMap) { $Visitors.Title } else { $title }
            if ($GroupsByTitle.ContainsKey($title)) {
                $login = $GroupsByTitle[$title].LoginName; $id = [int]$GroupsByTitle[$title].Id
            } else { $login = "SPGROUP:$title" }
        }
        if (-not $desired.ContainsKey($login)) {
            $desired[$login] = [pscustomobject]@{
                Login = $login; Title = $title; PType = $assignment.PrincipalType; Id = $id
                Roles = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
            }
        }
        foreach ($role in $assignment.Roles) { [void]$desired[$login].Roles.Add($role) }
    }
    return ,$desired
}

function Get-SPPermissionDelta {
    param($CurrentAcl, $Desired, $Unavailable, [string[]]$ProtectedLogins, $RoleIds)
    $current = [Collections.Generic.Dictionary[string, object]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($assignment in @($CurrentAcl) | Where-Object { -not $_.Skip }) { $current[$assignment.LoginName] = $assignment }
    $operations = [Collections.Generic.List[object]]::new()
    $blocked = @($Desired.Keys | Where-Object { $Unavailable.ContainsKey($_) }).Count -gt 0
    foreach ($login in @($Desired.Keys | Sort-Object)) {
        if ($Unavailable.ContainsKey($login)) { continue }
        foreach ($role in @($Desired[$login].Roles | Sort-Object)) {
            if (-not $RoleIds.ContainsKey($role)) { throw "Permission level missing on target: $role" }
            if ($current.ContainsKey($login) -and $current[$login].Roles -icontains $role) { continue }
            $operations.Add([pscustomobject]@{ Type = 'Add'; Login = $login; Role = $role; RoleId = $RoleIds[$role]; MemberId = $Desired[$login].Id })
        }
    }
    if (-not $blocked) {
        foreach ($login in @($current.Keys | Sort-Object)) {
            if ($ProtectedLogins -icontains $login -or $Unavailable.ContainsKey($login)) { continue }
            foreach ($role in @($current[$login].Roles | Sort-Object)) {
                if ($Desired.ContainsKey($login) -and $Desired[$login].Roles.Contains($role)) { continue }
                if (-not $RoleIds.ContainsKey($role)) { throw "Current target permission level not found: $role" }
                $operations.Add([pscustomobject]@{
                    Type = 'Remove'; Login = $login; Role = $role; RoleId = $RoleIds[$role]; MemberId = $current[$login].MemberId
                })
            }
        }
    }
    [pscustomobject]@{ Operations = $operations.ToArray(); RemovalsBlocked = $blocked }
}

function Get-SPProjectedAcl {
    param($CurrentAcl, $Operation, $Desired)
    $rows = [Collections.Generic.List[object]]::new()
    $found = $false
    foreach ($assignment in @($CurrentAcl)) {
        if ($assignment.Skip -or $assignment.LoginName -ine $Operation.Login) { $rows.Add($assignment); continue }
        $found = $true
        $roles = @($assignment.Roles)
        if ($Operation.Type -eq 'Add') { $roles += $Operation.Role }
        else { $roles = @($roles | Where-Object { $_ -ine $Operation.Role }) }
        if ($roles.Count) {
            $copy = $assignment | Select-Object *
            $copy.Roles = @($roles | Sort-Object -Unique)
            $rows.Add($copy)
        }
    }
    if (-not $found -and $Operation.Type -eq 'Add') {
        $principal = $Desired[$Operation.Login]
        $rows.Add([pscustomobject]@{
            LoginName = $Operation.Login; MemberId = $principal.Id; Title = $principal.Title
            PrincipalType = $principal.PType; Roles = @($Operation.Role); Bindings = @(); Skip = $false
        })
    }
    return ,$rows.ToArray()
}

function Test-SPPrincipalMissingError {
    param([string]$Message)
    # Do not reinterpret access denied, throttling or authentication as a missing user.
    if ($Message -match '401|403|429|503|access denied|unauthorized|forbidden|token|throttl|timeout|timed out') { return $false }
    return $Message -match 'specified user .+ could not be found|user .+ does not exist|principal .+ (could not be found|does not exist)'
}

function Test-SPPickerPrincipal {
    param($Connection, [string]$WebUrl, [string]$Login, [string]$Title)
    $query = if ($Login -match '\|membership\|(.+)$') { $Matches[1] } else { $Title }
    $body = @{ queryParams = @{
        AllowEmailAddresses = $true; AllowMultipleEntities = $false; AllUrlZones = $false
        MaximumEntitySuggestions = 50; PrincipalSource = 15; PrincipalType = 15; QueryString = $query
    } }
    $response = Invoke-SPPermissionRest $Connection (
        $WebUrl + '/_api/SP.UI.ApplicationPages.ClientPeoplePickerWebServiceInterface.clientPeoplePickerSearchUser'
    ) -Method Post -Content $body
    if (-not (Test-SPPermissionResponseKey $response 'value')) { throw 'Unexpected people-picker response; availability is unknown.' }
    $hits = @($response.value | ConvertFrom-Json -Depth 30 -AsHashtable -ErrorAction Stop)
    return @($hits | Where-Object { $_.Key -ieq $Login -and $_.IsResolved -ne $false }).Count -gt 0
}

function New-SPPermissionStep {
    param($Scope, [string]$Action, $BeforeAcl, [bool]$AfterUnique, $AfterAcl, $Operation)
    [pscustomobject]@{
        Key = $Scope.Key; Api = $Scope.Api; Kind = $Scope.Kind; Action = $Action
        BeforeUnique = $Scope.Unique; BeforeSignature = Get-SPPermissionAclSignature $BeforeAcl
        AfterUnique = $AfterUnique; AfterSignature = Get-SPPermissionAclSignature $AfterAcl
        Operation = $Operation
    }
}

function Add-SPDependentProtection {
    param($Target, [string]$ScopeKey, $EligibleTargets, $Steps, $Log, [switch]$NoFreeze)
    $candidates = @($Target.Values | Where-Object {
        $_.Key -and -not $_.Unique -and -not $EligibleTargets.Contains($_.Key) -and
        (Get-SPPermissionEffectiveScope $Target $_.Key).Key -ieq $ScopeKey
    } | Sort-Object @{ Expression = { ($_.Key -split '/').Count } }, Key)
    foreach ($dependent in $candidates) {
        # Re-evaluate after each freeze so descendants do not get redundant scopes.
        $ancestor = Get-SPPermissionEffectiveScope $Target $dependent.Key
        if ($ancestor.Key -ine $ScopeKey) { continue }
        if ($NoFreeze) {
            $Log.Add([pscustomobject]@{ Path = $dependent.Key; Action = 'Dependent follows changed parent (-NoFreeze)'; Result = 'Review'; Detail = $ScopeKey })
            continue
        }
        $Steps.Add((New-SPPermissionStep $dependent 'Freeze' $ancestor.Acl $true $ancestor.Acl $null))
        $dependent.Unique = $true; $dependent.Acl = @($ancestor.Acl)
    }
}

function New-SPPermissionPlan {
    param($Source, $Target, $ScopeMatches, $DesiredBySource, $Unavailable, $RoleIds, [string[]]$ProtectedAtRoot,
        [string[]]$ExcludePaths, [switch]$NoFreeze)
    $steps = [Collections.Generic.List[object]]::new()
    $log = [Collections.Generic.List[object]]::new()
    $eligible = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($sourceKey in $ScopeMatches.Keys) {
        $targetKey = $ScopeMatches[$sourceKey]
        if (-not (Test-SPPermissionExcluded $sourceKey $ExcludePaths) -and -not (Test-SPPermissionExcluded $targetKey $ExcludePaths)) {
            [void]$eligible.Add($targetKey)
        }
    }
    do {
        $removed = $false
        foreach ($sourceKey in $ScopeMatches.Keys) {
            $targetKey = $ScopeMatches[$sourceKey]
            if (-not $sourceKey -or $Source[$sourceKey].Unique -or -not $eligible.Contains($targetKey)) { continue }
            $sourceParent = Get-SPPermissionParent $Source $sourceKey
            $targetParent = Get-SPPermissionParent $Target $targetKey
            if (-not $ScopeMatches.ContainsKey($sourceParent) -or $ScopeMatches[$sourceParent] -ine $targetParent -or -not $eligible.Contains($targetParent)) {
                [void]$eligible.Remove($targetKey)
                $removed = $true
                $log.Add([pscustomobject]@{
                    Path = $targetKey; Action = 'Incompatible or unmatched parent'; Result = 'Skipped'
                    Detail = "Source: $sourceKey. Effective access cannot be mirrored by inheritance at this destination."
                })
            }
        }
    } while ($removed)
    foreach ($scope in $Source.Values) {
        if (-not $ScopeMatches.ContainsKey($scope.Key)) {
            $log.Add([pscustomobject]@{ Path = $scope.Key; Action = 'Source-only object'; Result = 'Skipped'; Detail = $scope.Kind })
        }
    }
    foreach ($scope in $Target.Values) {
        if (-not $eligible.Contains($scope.Key)) {
            $log.Add([pscustomobject]@{ Path = $scope.Key; Action = 'Target-only or excluded object'; Result = 'Untouched'; Detail = $scope.Kind })
        }
    }
    $ordered = @($ScopeMatches.Keys | Sort-Object @{ Expression = { ($ScopeMatches[$_] -split '/').Count } }, @{ Expression = { $ScopeMatches[$_] } })
    foreach ($sourceKey in $ordered) {
        $targetKey = $ScopeMatches[$sourceKey]
        if (-not $eligible.Contains($targetKey)) { continue }
        $sourceScope = $Source[$sourceKey]; $targetScope = $Target[$targetKey]
        $desired = $DesiredBySource[$sourceKey]
        if ($sourceKey -and -not $sourceScope.Unique) {
            $sourceParent = Get-SPPermissionParent $Source $sourceKey
            $targetParent = Get-SPPermissionParent $Target $targetKey
            if ($targetScope.Unique) {
                $blocked = @($desired.Keys | Where-Object { $Unavailable.ContainsKey($_) }).Count -gt 0
                if ($blocked) {
                    $log.Add([pscustomobject]@{ Path = $targetKey; Action = 'Reset inheritance blocked'; Result = 'Blocked'; Detail = 'Effective source ACL includes unavailable principals; existing grants retained.' })
                    continue
                }
                Add-SPDependentProtection $Target $targetKey $eligible $steps $log -NoFreeze:$NoFreeze
                $before = @($targetScope.Acl)
                $after = @((Get-SPPermissionEffectiveScope $Target $targetParent).Acl)
                $steps.Add((New-SPPermissionStep $targetScope 'Reset' $before $false $after $null))
                $targetScope.Unique = $false; $targetScope.Acl = $null
            }
            continue
        }
        $current = @((Get-SPPermissionEffectiveScope $Target $targetKey).Acl)
        if (-not $targetScope.Unique) {
            $steps.Add((New-SPPermissionStep $targetScope 'Break' $current $true $current $null))
            $targetScope.Unique = $true; $targetScope.Acl = $current
        }
        $protected = if ($sourceKey -eq '') { $ProtectedAtRoot } else { @() }
        $delta = Get-SPPermissionDelta $current $desired $Unavailable $protected $RoleIds
        if ($delta.RemovalsBlocked) {
            $log.Add([pscustomobject]@{ Path = $targetKey; Action = 'Removals blocked'; Result = 'Blocked'; Detail = 'Unavailable desired principals; available additions only. Existing grants retained.' })
        }
        if ($delta.Operations.Count) {
            Add-SPDependentProtection $Target $targetKey $eligible $steps $log -NoFreeze:$NoFreeze
        }
        foreach ($operation in $delta.Operations) {
            $after = Get-SPProjectedAcl $current $operation $desired
            $steps.Add((New-SPPermissionStep $targetScope $operation.Type $current $true $after $operation))
            $current = $after; $targetScope.Acl = $after
        }
    }
    [pscustomobject]@{ Steps = $steps.ToArray(); Notes = $log.ToArray() }
}

function Assert-SPPermissionState {
    param($Connection, $Step, [bool]$After)
    $scope = Invoke-SPPermissionRest $Connection ($Step.Api + '?$select=HasUniqueRoleAssignments')
    if ($null -eq $scope.HasUniqueRoleAssignments) { throw "Missing inheritance state: $($Step.Key)" }
    $aclApi = $Step.Api
    if (-not $scope.HasUniqueRoleAssignments) { $aclApi += '/FirstUniqueAncestorSecurableObject' }
    $signature = Get-SPPermissionAclSignature @(Get-SPPermissionAcl $Connection $aclApi)
    $expectedUnique = if ($After) { $Step.AfterUnique } else { $Step.BeforeUnique }
    $expectedSignature = if ($After) { $Step.AfterSignature } else { $Step.BeforeSignature }
    if ([bool]$scope.HasUniqueRoleAssignments -ne $expectedUnique -or $signature -cne $expectedSignature) {
        throw "ACL/inheritance differs from the plan at '$($Step.Key)' ($(if ($After) { 'after write' } else { 'before write' })). Stop and review; no further writes will run."
    }
}

function Invoke-SPPermissionPlan {
    param($Connection, $Plan, $Log)
    foreach ($step in $Plan.Steps) {
        try {
            Assert-SPPermissionState $Connection $step $false
            $operation = $step.Operation
            switch ($step.Action) {
                { $_ -in @('Freeze', 'Break') } {
                    $null = Invoke-SPPermissionRest $Connection ($step.Api + '/breakroleinheritance(copyRoleAssignments=true,clearSubscopes=false)') -Method Post
                }
                'Reset' { $null = Invoke-SPPermissionRest $Connection ($step.Api + '/resetroleinheritance') -Method Post }
                { $_ -in @('Add', 'Remove') } {
                    if ($operation.MemberId -le 0) { throw "Unresolved target principal: $($operation.Login)" }
                    $method = if ($step.Action -eq 'Add') { 'addroleassignment' } else { 'removeroleassignment' }
                    $null = Invoke-SPPermissionRest $Connection (
                        $step.Api + "/roleassignments/$method(principalid=$($operation.MemberId),roledefid=$($operation.RoleId))"
                    ) -Method Post
                }
                default { throw "Unknown plan action: $($step.Action)" }
            }
            Assert-SPPermissionState $Connection $step $true
            $Log.Add([pscustomobject]@{
                Path = $step.Key; Action = $step.Action; Result = 'Verified'
                Detail = "$($operation.Login) $($operation.Role)"
            })
        } catch {
            $Log.Add([pscustomobject]@{
                Path = $step.Key; Action = $step.Action; Result = 'Failed - inspect live state'
                Detail = $_.Exception.Message
            })
            throw
        }
    }
    # Check final states again, not just the immediate response to each individual write.
    foreach ($group in @($Plan.Steps | Group-Object Key)) {
        Assert-SPPermissionState $Connection $group.Group[-1] $true
    }
}

function Invoke-SyncSPPermissionsFromSource {
    $ErrorActionPreference = 'Stop'
    $OutputFolder = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputFolder)
    Assert-SPPermissionUrls $SourceUrl $TargetUrl
    $SourceUrl = $SourceUrl.TrimEnd('/')
    $TargetUrl = $TargetUrl.TrimEnd('/')
    Import-Module PnP.PowerShell -MinimumVersion 3.3.0 -ErrorAction Stop
    [void][IO.Directory]::CreateDirectory($OutputFolder)
    $stamp = Get-Date -Format 'yyyyMMdd-HHmmss-fff'
    $mode = if ($Apply) { 'APPLY' } else { 'DRYRUN' }
    $logPath = Join-Path $OutputFolder "PermSync_Log_${mode}_$stamp.csv"
    $unavailablePath = Join-Path $OutputFolder "PermSync_UnavailablePrincipals_${mode}_$stamp.csv"
    $log = [Collections.Generic.List[object]]::new()
    $unavailable = [Collections.Generic.Dictionary[string, string]]::new([StringComparer]::OrdinalIgnoreCase)
    $principalUse = [Collections.Generic.Dictionary[string, object]]::new([StringComparer]::OrdinalIgnoreCase)
    try {
        $sourceConnection = Connect-SPPermissionSite $SourceUrl $ClientId $Tenant -SkipPersistLoginSetup:$SkipPersistLoginSetup
        $targetConnection = Connect-SPPermissionSite $TargetUrl $ClientId $Tenant -SkipPersistLoginSetup:$SkipPersistLoginSetup
        $source = Get-SPPermissionInventory $sourceConnection $SourceUrl $PageSize
        $target = Get-SPPermissionInventory $targetConnection $TargetUrl $PageSize
        foreach ($side in @('Source', 'Target')) {
            $inventory = if ($side -eq 'Source') { $source } else { $target }
            foreach ($scope in $inventory.Values | Where-Object { $_.InventoryCountReview }) {
                $log.Add([pscustomobject]@{
                    Path = $scope.Key; Action = 'Inventory count discrepancy'; Result = 'Review'
                    Detail = "$side - $($scope.InventoryCountReview)"
                })
            }
        }
        $scopeMatches = Get-SPPermissionMatches $source $target $MatchFile $SourceUrl $TargetUrl
        $inventoryRows = @(
            foreach ($side in @('Source', 'Target')) {
                $inventory = if ($side -eq 'Source') { $source } else { $target }
                foreach ($scope in $inventory.Values | Sort-Object Key) {
                    $otherPath = ''
                    $matched = $false
                    if ($side -eq 'Source' -and $scopeMatches.ContainsKey($scope.Key)) {
                        $matched = $true; $otherPath = $scopeMatches[$scope.Key]
                    } elseif ($side -eq 'Target') {
                        $sourceKeys = @($scopeMatches.Keys | Where-Object { $scopeMatches[$_] -ieq $scope.Key })
                        if ($sourceKeys.Count) { $matched = $true; $otherPath = $sourceKeys[0] }
                    }
                    $state = if ($scope.Kind -eq 'Subweb') { 'Subweb boundary - processed separately by checklist runner' }
                        elseif (Test-SPPermissionExcluded $scope.Key $ExcludePaths) { 'Excluded' }
                        elseif ($matched) { 'Matched' } else { 'Unmatched - skipped' }
                    [pscustomobject]@{
                        Side = $side; Path = $scope.Key; Kind = $scope.Kind; Unique = $scope.Unique
                        CounterpartPath = $otherPath; Status = $state
                    }
                }
            }
        )
        Export-SPPermissionCsv $inventoryRows @('Side','Path','Kind','Unique','CounterpartPath','Status') (
            Join-Path $OutputFolder "PermSync_InventoryMatches_$stamp.csv"
        )
        foreach ($side in @('Source', 'Target')) {
            $inventory = if ($side -eq 'Source') { $source } else { $target }
            $connection = if ($side -eq 'Source') { $sourceConnection } else { $targetConnection }
            foreach ($scope in $inventory.Values) {
                if ($scope.Unique -or $scope.Key -eq '') {
                    $api = $scope.Api
                    if (-not $scope.Unique) { $api += '/FirstUniqueAncestorSecurableObject' }
                    $scope.Acl = @(Get-SPPermissionAcl $connection $api)
                }
            }
            # Include every scope's inheritance state, not only the unique ACLs.
            [pscustomobject]@{
                WebUrl = $(if ($side -eq 'Source') { $SourceUrl } else { $TargetUrl })
                CapturedUtc = [datetime]::UtcNow.ToString('o')
                Scopes = @($inventory.Values | Sort-Object Key)
            } | ConvertTo-Json -Depth 50 | Set-Content -LiteralPath (
                Join-Path $OutputFolder "PermSync_${side}Snapshot_$stamp.json"
            ) -Encoding utf8
        }
        $groups = [Collections.Generic.Dictionary[string, object]]::new([StringComparer]::OrdinalIgnoreCase)
        foreach ($group in @(Get-SPPermissionCollection $targetConnection ($TargetUrl + '/_api/web/sitegroups?$select=Id,Title,LoginName'))) {
            if ($groups.ContainsKey($group.Title)) { throw "Duplicate target SharePoint group title: $($group.Title)" }
            $groups[$group.Title] = $group
        }
        $visitors = Invoke-SPPermissionRest $targetConnection ($TargetUrl + '/_api/web/AssociatedVisitorGroup?$select=Id,Title,LoginName')
        if (-not $visitors.Id -or -not $groups.ContainsKey($visitors.Title)) { throw 'Target has no valid associated Visitors group.' }
        $protectedAtRoot = @(
            (Invoke-SPPermissionRest $targetConnection ($TargetUrl + '/_api/web/AssociatedOwnerGroup?$select=LoginName')).LoginName
            (Invoke-SPPermissionRest $targetConnection ($TargetUrl + '/_api/web/AssociatedMemberGroup?$select=LoginName')).LoginName
        ) | Where-Object { $_ }
        $sourceRoles = [Collections.Generic.Dictionary[string, object]]::new([StringComparer]::OrdinalIgnoreCase)
        $targetRoles = [Collections.Generic.Dictionary[string, object]]::new([StringComparer]::OrdinalIgnoreCase)
        $roleIds = [Collections.Generic.Dictionary[string, int]]::new([StringComparer]::OrdinalIgnoreCase)
        foreach ($role in @(Get-SPPermissionCollection $sourceConnection ($SourceUrl + '/_api/web/roledefinitions?$select=Id,Name,BasePermissions'))) { $sourceRoles[$role.Name] = $role }
        foreach ($role in @(Get-SPPermissionCollection $targetConnection ($TargetUrl + '/_api/web/roledefinitions?$select=Id,Name,BasePermissions'))) {
            $targetRoles[$role.Name] = $role; $roleIds[$role.Name] = [int]$role.Id
        }
        $desiredBySource = [Collections.Generic.Dictionary[string, object]]::new([StringComparer]::OrdinalIgnoreCase)
        foreach ($sourceKey in $scopeMatches.Keys) {
            if ((Test-SPPermissionExcluded $sourceKey $ExcludePaths) -or (Test-SPPermissionExcluded $scopeMatches[$sourceKey] $ExcludePaths)) { continue }
            $effective = Get-SPPermissionEffectiveScope $source $sourceKey
            $desired = Get-SPDesiredAcl $effective.Acl $groups $visitors $SourceGroupToMap
            $desiredBySource[$sourceKey] = $desired
            foreach ($principal in $desired.Values) {
                foreach ($role in $principal.Roles) {
                    if (-not $sourceRoles.ContainsKey($role) -or -not $targetRoles.ContainsKey($role)) {
                        throw "Permission level '$role' must exist on BOTH webs. No permission writes performed."
                    }
                    $sourceBits = $sourceRoles[$role].BasePermissions; $targetBits = $targetRoles[$role].BasePermissions
                    if ($null -eq $sourceBits -or $null -eq $targetBits -or
                        [string]$sourceBits.High -cne [string]$targetBits.High -or [string]$sourceBits.Low -cne [string]$targetBits.Low) {
                        throw "Permission level '$role' has different/unknown permission bits across webs."
                    }
                }
                if (-not $principalUse.ContainsKey($principal.Login)) {
                    $principalUse[$principal.Login] = [pscustomobject]@{
                        Login = $principal.Login; Title = $principal.Title; PType = $principal.PType; Id = $principal.Id
                        Roles = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
                        Objects = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
                    }
                }
                [void]$principalUse[$principal.Login].Objects.Add($sourceKey)
                foreach ($role in $principal.Roles) { [void]$principalUse[$principal.Login].Roles.Add($role) }
            }
        }
        $siteUsers = [Collections.Generic.Dictionary[string, int]]::new([StringComparer]::OrdinalIgnoreCase)
        foreach ($user in @(Get-SPPermissionCollection $targetConnection ($TargetUrl + '/_api/web/siteusers?$select=Id,LoginName'))) {
            $siteUsers[$user.LoginName] = [int]$user.Id
        }
        foreach ($principal in $principalUse.Values) {
            if ($principal.Login.StartsWith('SPGROUP:', [StringComparison]::OrdinalIgnoreCase)) {
                $unavailable[$principal.Login] = 'SharePoint group absent on target; no group is created.'
            } elseif ($principal.PType -ne 8) {
                if ($Apply) {
                    try {
                        $user = New-PnPUser -LoginName $principal.Login -Connection $targetConnection -ErrorAction Stop
                        if (-not $user.Id) { throw "EnsureUser returned no ID for '$($principal.Login)'." }
                        $principal.Id = [int]$user.Id
                    } catch {
                        if (-not (Test-SPPrincipalMissingError $_.Exception.Message)) { throw }
                        $unavailable[$principal.Login] = $_.Exception.Message
                    }
                } elseif ($siteUsers.ContainsKey($principal.Login)) { $principal.Id = $siteUsers[$principal.Login] }
                elseif (-not (Test-SPPickerPrincipal $targetConnection $TargetUrl $principal.Login $principal.Title)) {
                    $unavailable[$principal.Login] = 'Exact login not found by people picker (prediction only; Apply uses EnsureUser).'
                }
            }
        }
        foreach ($desired in $desiredBySource.Values) {
            foreach ($principal in $desired.Values) { $principal.Id = $principalUse[$principal.Login].Id }
        }
        # Planning mutates the in-memory target; the saved snapshots retain the read state.
        $plan = New-SPPermissionPlan $source $target $scopeMatches $desiredBySource $unavailable $roleIds $protectedAtRoot $ExcludePaths -NoFreeze:$NoFreeze
        foreach ($note in $plan.Notes) { $log.Add($note) }
        $planPath = Join-Path $OutputFolder "PermSync_Plan_${mode}_$stamp.json"
        [pscustomobject]@{
            SourceUrl = $SourceUrl; TargetUrl = $TargetUrl; Mode = $mode
            CreatedUtc = [datetime]::UtcNow.ToString('o'); Steps = @($plan.Steps); Notes = @($plan.Notes)
        } | ConvertTo-Json -Depth 30 | Set-Content -LiteralPath $planPath -Encoding utf8
        foreach ($step in $plan.Steps) {
            $log.Add([pscustomobject]@{
                Path = $step.Key; Action = $step.Action; Result = 'Planned'
                Detail = "$($step.Operation.Login) $($step.Operation.Role)"
            })
        }
        # Persist the complete plan before the first permission write.
        Export-SPPermissionCsv $log.ToArray() @('Path','Action','Result','Detail') $logPath
        if ($Apply) { Invoke-SPPermissionPlan $targetConnection $plan $log }
        $blockedCount = @($plan.Notes | Where-Object Result -eq 'Blocked').Count
        Write-Host "[$mode] Permission actions: $($plan.Steps.Count); blocked scopes: $blockedCount; unavailable principals: $($unavailable.Count)"
        Write-Host "Plan: $planPath"
        if ($blockedCount -or $unavailable.Count) { Write-Warning 'This run is NOT an exact mirror. Review the blocked scopes and unavailable-principals report.' }
    } catch {
        $log.Add([pscustomobject]@{ Path = ''; Action = 'Run stopped'; Result = 'Error'; Detail = $_.Exception.Message })
        throw
    } finally {
        Export-SPPermissionCsv $log.ToArray() @('Path','Action','Result','Detail') $logPath
        $rows = @(
            foreach ($login in $unavailable.Keys) {
                $use = $principalUse[$login]
                [pscustomobject]@{
                    Principal = $login; SourceTitle = $use.Title; PrincipalType = $use.PType
                    Reason = $unavailable[$login]; Roles = (@($use.Roles | Sort-Object) -join '; ')
                    SourceObjects = (@($use.Objects | Sort-Object) -join '; '); ObjectCount = $use.Objects.Count
                }
            }
        )
        Export-SPPermissionCsv $rows @('Principal','SourceTitle','PrincipalType','Reason','Roles','SourceObjects','ObjectCount') $unavailablePath
        Write-Host "Log: $logPath"
        Write-Host "Unavailable principals: $unavailablePath"
    }
}

if ($MyInvocation.InvocationName -ne '.') { Invoke-SyncSPPermissionsFromSource }
