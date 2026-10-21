#requires -Version 7.4.6
<#
.SYNOPSIS
    TEMPORARY read-only diagnostic for the three failed Visitors permission sites.
.DESCRIPTION
    Documents and Forms, The Patient Experience and Training only. Does not add
    users, change group membership or write SharePoint permissions. Reads the
    server's saved run evidence (including ignored JSON) plus live state, and writes
    a compact JSON file to SupportFiles for pushing back. Delete this script and its
    output once the recovery is complete.
#>
[CmdletBinding()]
param(
    [string]$RunFolder = (Join-Path $PSScriptRoot '..\Reports\VisitorsRuns\v-261020150543067'),
    [string]$OutputFile = (Join-Path $PSScriptRoot (
        '..\SupportFiles\VisitorsRecoveryEvidence-' + (Get-Date -Format 'yyMMddHHmmss') + '.json')),
    [string]$ClientId = '2e20d82f-56ba-418e-b465-9718493f0df8',
    [string]$Tenant,
    [switch]$SkipPersistLoginSetup
)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'SPPermission.Common.ps1')
Import-Module PnP.PowerShell -MinimumVersion 3.3.0 -ErrorAction Stop
if (-not (Test-Path -LiteralPath $RunFolder -PathType Container)) { throw "Run folder not found: $RunFolder" }
if (Test-Path -LiteralPath $OutputFile) { throw "Output already exists: $OutputFile" }

function Read-RecoveryProbe {
    param([scriptblock]$Action)
    try {
        [pscustomobject]@{ Status = 'Read succeeded'; Data = (& $Action); Error = ''; Stack = '' }
    } catch {
        Write-Warning $_.Exception.Message
        [pscustomobject]@{
            Status = 'Read failed'; Data = $null
            Error = $_.Exception.ToString(); Stack = $_.ScriptStackTrace
        }
    }
}

$sites = @(
    @{ Key = '0099553DA6F5'; Name = 'DocumentsandForms'; Path = '_catalogs/users/10_.000' }
    @{ Key = 'C29814F9B73B'; Name = 'ThePatientExperience'; Path = '' }
    @{ Key = '9352C5262E61'; Name = 'Training'; Path = 'SiteAssets/SitePages/Annual-Development' }
)
$report = [ordered]@{
    CreatedUtc = [datetime]::UtcNow.ToString('o')
    Mode = 'Read-only; no SharePoint writes'
    RunFolder = $RunFolder
    PnPVersion = [string](Get-Module PnP.PowerShell | Select-Object -First 1).Version
    Sites = [Collections.Generic.List[object]]::new()
}
try {
    foreach ($site in $sites) {
        $url = 'https://whallc.sharepoint.com/sites/' + $site.Name
        Write-Host "Reading recovery evidence: $url"
        $entry = [ordered]@{ TargetUrl = $url; FailedPath = $site.Path }
        $report.Sites.Add($entry)
        $entry.SavedEvidence = Read-RecoveryProbe {
            $applyFolder = Join-Path $RunFolder "$($site.Key)\Apply"
            $runs = @(Get-ChildItem -LiteralPath $applyFolder -Recurse -Directory -Filter 'rd-*')
            if ($runs.Count -ne 1) { throw "Expected exactly one Apply evidence folder in '$applyFolder'." }
            $evidence = $runs[0].FullName
            # The failed scope plus each ancestor determines its effective ACL.
            $keys = @('')
            $path = ''
            foreach ($segment in @($site.Path -split '/' | Where-Object { $_ })) {
                $path = if ($path) { "$path/$segment" } else { $segment }
                $keys += $path
            }
            $saved = [ordered]@{ Folder = $evidence; RelevantKeys = $keys }
            $saved.Status = Read-RecoveryProbe {
                Get-Content -LiteralPath (Join-Path $evidence 'Status.json') -Raw |
                    ConvertFrom-Json -Depth 100 -AsHashtable
            }
            $saved.Plan = Read-RecoveryProbe {
                $plan = Get-Content -LiteralPath (Join-Path $evidence 'Plan.json') -Raw |
                    ConvertFrom-Json -Depth 100 -AsHashtable
                [pscustomobject]@{
                    Steps = @($plan.Steps | Where-Object { $_.Key -in $keys })
                    Notes = @($plan.Notes | Where-Object { $_.Path -in $keys })
                }
            }
            $saved.Target = Read-RecoveryProbe {
                $snapshot = Get-Content -LiteralPath (Join-Path $evidence 'Target.json') -Raw |
                    ConvertFrom-Json -Depth 100 -AsHashtable
                @($snapshot.Scopes | Where-Object { $_.Key -in $keys })
            }
            $saved.Changes = @(Import-Csv -LiteralPath (Join-Path $evidence 'Changes.csv') |
                Where-Object { $_.Path -in $keys -or $_.Result -like 'Failed*' -or $_.Action -like '*EEEU*' })
            [pscustomobject]$saved
        }
        $entry.Live = Read-RecoveryProbe {
            $connection = Connect-SPPermissionSite -Url $url -ClientId $ClientId -Tenant $Tenant `
                -SkipPersistLoginSetup:$SkipPersistLoginSetup
            $live = [ordered]@{}
            $live.Visitors = Read-RecoveryProbe {
                $group = Invoke-SPPermissionRest $connection (
                    $url + '/_api/web/AssociatedVisitorGroup?$select=Id,Title,LoginName')
                if (-not $group.Id) { throw "Associated Visitors group unavailable: $url" }
                $tenantId = [guid](Get-PnPTenantId -Connection $connection -ErrorAction Stop)
                if ($tenantId -eq [guid]::Empty) { throw 'Tenant ID unavailable.' }
                $claim = "c:0-.f|rolemanager|spo-grid-all-users/$tenantId"
                $members = @(Get-SPPermissionCollection $connection (
                    "$url/_api/web/sitegroups/getbyid($($group.Id))/users?`$select=Id,Title,LoginName"))
                [pscustomobject]@{
                    Group = $group; EEEUClaim = $claim
                    EEEUPresent = @($members | Where-Object LoginName -IEQ $claim).Count -gt 0
                    MemberCount = $members.Count
                }
            }
            if ($site.Name -eq 'DocumentsandForms') {
                $listId = 'a97bb88d-9ce9-4647-9f70-90bb37f38d93'
                $live.UserInformationItemPnP = Read-RecoveryProbe {
                    $item = Get-PnPListItem -List $listId -Id 10 -Connection $connection -ErrorAction Stop
                    $unique = Get-PnPProperty -ClientObject $item -Property HasUniqueRoleAssignments `
                        -Connection $connection -ErrorAction Stop
                    if ($null -eq $unique) { throw 'PnP returned no inheritance state for item 10.' }
                    [pscustomobject]@{ Id = $item.Id; HasUniqueRoleAssignments = $unique }
                }
                $live.UserInformationItemRest = Read-RecoveryProbe {
                    Invoke-SPPermissionRest $connection (
                        "$url/_api/web/lists(guid'$listId')/items(10)?`$select=Id,HasUniqueRoleAssignments")
                }
            }
            $live.SavedScopeStates = @(
                foreach ($scope in @($entry.SavedEvidence.Data.Target.Data)) {
                    if (-not $scope.Api) { continue }
                    [pscustomobject]@{
                        Key = $scope.Key
                        State = Read-RecoveryProbe {
                            $state = Invoke-SPPermissionRest $connection (
                                $scope.Api + '?$select=HasUniqueRoleAssignments')
                            if ($null -eq $state.HasUniqueRoleAssignments) { throw "Missing inheritance state: $($scope.Key)" }
                            $api = $scope.Api
                            if (-not $state.HasUniqueRoleAssignments) { $api += '/FirstUniqueAncestorSecurableObject' }
                            $acl = @(Get-SPPermissionAcl $connection $api)
                            [pscustomobject]@{
                                Unique = $state.HasUniqueRoleAssignments
                                Signature = Get-SPPermissionAclSignature $acl
                                Acl = $acl
                            }
                        }
                    }
                }
            )
            [pscustomobject]$live
        }
    }
} finally {
    $report | ConvertTo-Json -Depth 100 |
        Set-Content -LiteralPath $OutputFile -Encoding utf8 -ErrorAction Stop
    Write-Host "Recovery evidence saved: $OutputFile"
    Write-Host 'Push this file back. Read failures are recorded explicitly; no SharePoint writes were attempted.'
}
