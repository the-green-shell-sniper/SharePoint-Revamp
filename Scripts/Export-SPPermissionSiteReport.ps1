#requires -Version 7.4.6
<#
.SYNOPSIS
    Read-only recursive target permission reference, including unmatched content.
.DESCRIPTION
    Streams one combined CSV for a web and all descendants, root first. Enumerates
    every list and item visible to the account, including hidden and excluded lists.
    Reports inherited role assignments, not expanded group membership or a user
    access simulation. Sharing-link/limited-access bindings are included as observed;
    site admins, app permissions, link recipients and tenant policies are not expanded.
    A failed read leaves a partial CSV and a Failed manifest, never a complete result.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$TargetUrl,
    [Parameter(Mandatory)][string]$OutputFolder,
    [string]$ClientId = '2e20d82f-56ba-418e-b465-9718493f0df8',
    [string]$Tenant,
    [ValidateRange(1, 5000)][int]$PageSize = 500,
    [switch]$SkipPersistLoginSetup
)
. (Join-Path $PSScriptRoot 'SPOutputLimits.Common.ps1')
. (Join-Path $PSScriptRoot 'SPPermission.Common.ps1')

function Get-SPPermissionReportRows {
    param([string]$RootUrl, [hashtable]$ConnectionOptions, [int]$Size, $Counters)
    $queue = [Collections.Generic.List[string]]::new()
    $queue.Add($RootUrl.TrimEnd('/'))
    $seen = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    for ($index = 0; $index -lt $queue.Count; $index++) {
        $url = $queue[$index]
        if (-not $seen.Add($url)) { throw "Repeated report web: $url" }
        Write-Host "[Report] $url"
        $connection = Connect-SPPermissionSite -Url $url @ConnectionOptions
        $definitions = @{}
        foreach ($role in @(Get-SPPermissionCollection $connection (
            $url + '/_api/web/roledefinitions?$select=Id,Name,Description,BasePermissions'
        ))) {
            $definitions[[string]$role.Id] = $role
            if ($Counters.RolePath) {
                [pscustomobject]@{
                    WebUrl = $url; RoleId = $role.Id; Role = $role.Name; Description = $role.Description
                    High = $role.BasePermissions.High; Low = $role.BasePermissions.Low
                } | Export-Csv -LiteralPath $Counters.RolePath -Append -NoTypeInformation -Encoding utf8
            }
        }
        $protection = Get-SPMigrationProtection $connection $url
        if ($Counters.MembershipPath) {
            foreach ($membership in $protection.Memberships) {
                $membership | Export-Csv -LiteralPath $Counters.MembershipPath -Append -NoTypeInformation -Encoding utf8
            }
        }
        $inventory = Get-SPPermissionInventory $connection $url $Size
        $captured = [datetime]::UtcNow.ToString('o')
        $inventoryReview = (@($inventory.Values | Where-Object { $_.InventoryCountReview -or $_.InventoryReadReview } |
            ForEach-Object { "$($_.Key): $($_.InventoryCountReview) $($_.InventoryReadReview)" }) -join '; ')
        $rootOrigin = $url
        if (-not $inventory[''].Unique) {
            $ancestor = Invoke-SPPermissionRest $connection ($url + '/_api/web/FirstUniqueAncestorSecurableObject?$select=Url')
            if (-not $ancestor.Url) { throw "Missing inherited root ACL origin for '$url'." }
            $rootOrigin = $ancestor.Url
        }
        foreach ($scope in $inventory.Values | Where-Object { $_.Kind -ne 'Subweb' -and ($_.Unique -or $_.Key -eq '') }) {
            $api = $scope.Api
            if (-not $scope.Unique) { $api += '/FirstUniqueAncestorSecurableObject' }
            $scope.Acl = @(Get-SPPermissionAcl $connection $api)
        }
        foreach ($scope in $inventory.Values | Where-Object Kind -ne 'Subweb' | Sort-Object Depth, Key) {
            $effective = Get-SPPermissionEffectiveScope $inventory $scope.Key
            $origin = if (-not $effective.Key) { $rootOrigin } else { "$url/$($effective.Key)" }
            $assignments = @($effective.Acl)
            if (-not $assignments.Count) { $assignments = @($null) }
            foreach ($assignment in $assignments) {
                $preserved = $assignment -and ((Test-SPMigrationPrincipal $assignment.LoginName $assignment.Title) -or
                    $protection.Logins.Contains($assignment.LoginName))
                $bindings = if ($assignment) { @($assignment.Bindings) } else { @() }
                if (-not $bindings.Count) { $bindings = @($null) }
                foreach ($binding in $bindings) {
                    $definition = if ($binding) { $definitions[[string]$binding.Id] } else { $null }
                    [pscustomobject]@{
                        SiteUrl = $RootUrl; WebUrl = $url; Path = $scope.Key; Kind = $scope.Kind
                        ObjectTitle = $scope.Title; ObjectId = $scope.Id; ListPath = $scope.ListKey
                        FileRef = $scope.FileRef
                        ObjectApi = $scope.Api; Unique = $scope.Unique; InheritedFrom = $origin
                        Principal = $assignment.LoginName; PrincipalTitle = $assignment.Title
                        PrincipalType = $assignment.PrincipalType; MemberId = $assignment.MemberId
                        Role = $binding.Name; RoleId = $binding.Id; HiddenRole = $binding.Hidden
                        RoleTypeKind = $binding.RoleTypeKind; PermissionHigh = $definition.BasePermissions.High
                        PermissionLow = $definition.BasePermissions.Low
                        MaskStatus = $(if ($definition -and $null -ne $definition.BasePermissions.High -and
                            $null -ne $definition.BasePermissions.Low) { 'Observed' } else { 'Unavailable - inspect role definition' })
                        ExcludedFromSync = ($preserved -or $assignment.Skip -or $binding.Hidden -or $binding.RoleTypeKind -eq 1 -or
                            $binding.Id -eq 1073741825 -or $binding.Name -in @('Limited Access', 'Web-Only Limited Access'))
                        Observation = $(if ($null -eq $assignment) { 'No role assignments returned' } else { 'Observed role binding' })
                        MigrationPolicy = $(if ($preserved) { 'Intentionally preserved - migration group or containing SharePoint group' } else { '' })
                        MigrationMembers = (@($protection.Memberships | Where-Object { $_.GroupLogin -and $_.GroupLogin -ieq $assignment.LoginName } |
                            ForEach-Object MigrationPrincipal) -join '; ')
                        InventoryReview = $inventoryReview
                        CapturedUtc = $captured
                    }
                    $Counters.Rows++
                }
            }
            $Counters.Objects++
        }
        $Counters.Webs++
        $offset = 1
        foreach ($child in $inventory.Values | Where-Object Kind -eq 'Subweb' | Sort-Object Url) {
            if (-not (Test-SPPermissionPathWithin $child.Url $RootUrl)) { throw "Report child outside site: $($child.Url)" }
            $queue.Insert($index + $offset, $child.Url); $offset++
        }
        $inventory = $null
    }
}

function Invoke-SPPermissionSiteReport {
    $ErrorActionPreference = 'Stop'
    Import-Module PnP.PowerShell -MinimumVersion 3.3.0 -ErrorAction Stop
    $folder = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputFolder)
    [void](Assert-SPOutputPath $folder)
    [void][IO.Directory]::CreateDirectory($folder)
    $manifest = [ordered]@{
        SiteUrl = $TargetUrl; StartedUtc = [datetime]::UtcNow.ToString('o')
        Status = 'Running'; Webs = 0; Objects = 0; Rows = 0; Error = ''
        Coverage = 'All enumerated target webs/lists/folders/items, including unmatched/excluded content; signed-in account visibility only.'
        Limitations = 'Role-assignment reference, not expanded individual access. Only the four protected migration principals are checked for direct SharePoint-group membership; nested Entra membership, site admins, app grants, sharing-link recipients and tenant policies are not expanded. Not a transactional snapshot.'
    }
    $manifestPath = Join-Path $folder 'ReportStatus.json'
    $manifest | ConvertTo-Json | Set-Content -LiteralPath $manifestPath -Encoding utf8
    try {
        $manifest['MembershipPath'] = Join-Path $folder 'MigrationGroupMemberships.csv'
        Export-SPPermissionCsv @() @('WebUrl','MigrationPrincipal','MigrationLogin','SharePointGroup','GroupLogin','GroupId','Policy') $manifest.MembershipPath
        $manifest['RolePath'] = Join-Path $folder 'PermissionLevels.csv'
        Export-SPPermissionCsv @() @('WebUrl','RoleId','Role','Description','High','Low') $manifest.RolePath
        $options = @{ ClientId = $ClientId; Tenant = $Tenant; SkipPersistLoginSetup = $SkipPersistLoginSetup }
        Get-SPPermissionReportRows $TargetUrl $options $PageSize $manifest |
            Export-Csv -LiteralPath (Join-Path $folder 'EffectivePermissions.csv') -NoTypeInformation -Encoding utf8
        $manifest.Status = 'Complete'
    } catch {
        $manifest.Status = 'Failed'; $manifest.Error = $_.Exception.Message
        throw
    } finally {
        $manifest['FinishedUtc'] = [datetime]::UtcNow.ToString('o')
        $manifest | ConvertTo-Json | Set-Content -LiteralPath $manifestPath -Encoding utf8
        Write-Host "Site permission report: $folder ($($manifest.Status))"
    }
}

if ($MyInvocation.InvocationName -ne '.') { Invoke-SPPermissionSiteReport }
