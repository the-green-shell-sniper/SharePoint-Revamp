#requires -Version 7.4.6

. (Join-Path $PSScriptRoot 'SPOutputLimits.Common.ps1')

function Connect-SPPermissionSite {
    param([string]$Url, [string]$ClientId, [string]$Tenant, [switch]$SkipPersistLoginSetup)
    $options = @{ Url = $Url; Interactive = $true; ReturnConnection = $true; ErrorAction = 'Stop' }
    if ($ClientId) { $options.ClientId = $ClientId }
    if ($Tenant) { $options.Tenant = $Tenant }
    if (-not $SkipPersistLoginSetup) { $options.PersistLogin = $true }
    Connect-PnPOnline @options
}

function Test-SPMigrationPrincipal {
    param([string]$Login, [string]$Title)
    $names = @('SharePoint.Migration.Team.Admins', 'SharePoint.Migration.Team.Members',
        'SharePoint.Migration.Team.Visitors', 'SharePoint.Migration.Team.MSLs')
    return ($Title -iin $names -or $Login -iin $names -or ($Login -split '[|\\]')[-1] -iin $names)
}

function Get-SPMigrationProtection {
    param($Connection, [string]$WebUrl)
    $logins = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $memberships = [Collections.Generic.List[object]]::new()
    foreach ($group in @(Get-SPPermissionCollection $Connection (
        $WebUrl + '/_api/web/sitegroups?$select=Id,Title,LoginName'
    ))) {
        foreach ($member in @(Get-SPPermissionCollection $Connection (
            "$WebUrl/_api/web/sitegroups/getbyid($($group.Id))/users?`$select=Id,Title,LoginName,PrincipalType"
        ))) {
            if (-not (Test-SPMigrationPrincipal $member.LoginName $member.Title)) { continue }
            [void]$logins.Add($member.LoginName)
            [void]$logins.Add($group.LoginName)
            $memberships.Add([pscustomobject]@{
                WebUrl = $WebUrl; MigrationPrincipal = $member.Title; MigrationLogin = $member.LoginName
                SharePointGroup = $group.Title; GroupLogin = $group.LoginName; GroupId = $group.Id
                Policy = 'Intentionally preserved - migration group and containing SharePoint group'
            })
        }
    }
    foreach ($name in @('SharePoint.Migration.Team.Admins', 'SharePoint.Migration.Team.Members',
        'SharePoint.Migration.Team.Visitors', 'SharePoint.Migration.Team.MSLs')) {
        if (@($memberships | Where-Object {
            $_.MigrationPrincipal -ieq $name -or ($_.MigrationLogin -split '[|\\]')[-1] -ieq $name
        }).Count) { continue }
        $memberships.Add([pscustomobject]@{
            WebUrl = $WebUrl; MigrationPrincipal = $name; MigrationLogin = ''
            SharePointGroup = ''; GroupLogin = ''; GroupId = ''
            Policy = 'Not observed in SharePoint-group membership; direct grants reported separately. Intentionally preserved if present.'
        })
    }
    [pscustomobject]@{ Logins = $logins; Memberships = $memberships.ToArray() }
}

function Assert-SPPermissionUrls {
    param([string]$SourceUrl, [string]$TargetUrl)
    foreach ($url in @($SourceUrl, $TargetUrl)) {
        $uri = $null
        if (-not [uri]::TryCreate($url, [UriKind]::Absolute, [ref]$uri) -or
            $uri.Scheme -ne 'https' -or $uri.Query -or $uri.Fragment -or $uri.UserInfo) {
            throw "Use an absolute HTTPS web URL without credentials, query or fragment: $url"
        }
    }
    if ($SourceUrl.TrimEnd('/') -ieq $TargetUrl.TrimEnd('/')) { throw 'Source and target must be different webs.' }
}

function Invoke-SPPermissionRest {
    param($Connection, [string]$Url, [ValidateSet('Get', 'Post')][string]$Method = 'Get', $Content)
    $options = @{ Connection = $Connection; Url = $Url; Method = $Method; Raw = $true; ErrorAction = 'Stop' }
    if ($PSBoundParameters.ContainsKey('Content')) { $options.Content = $Content }
    elseif ($Method -eq 'Post') {
        $options.Content = '{}'
    }
    for ($attempt = 0; ; $attempt++) {
        try {
            $raw = Invoke-PnPSPRestMethod @options
            if ([string]::IsNullOrWhiteSpace($raw)) {
                if ($Method -eq 'Get') { throw "Empty SharePoint response: $Url" }
                return
            }
            # SharePoint item responses can contain both Id and ID.
            return ($raw | ConvertFrom-Json -Depth 100 -AsHashtable)
        } catch {
            # Writes are not retried: a lost response can mean the write already succeeded.
            if ($Method -eq 'Get' -and $attempt -lt 5 -and
                $_.Exception.Message -match '429|503|Too Many Requests|Server Unavailable|throttl') {
                Write-Warning "Read throttled; retrying $Url"
                Start-Sleep -Seconds ([math]::Min(120, 5 * [math]::Pow(2, $attempt)))
                continue
            }
            throw [InvalidOperationException]::new("SharePoint $Method failed at '$Url': $($_.Exception.Message)", $_.Exception)
        }
    }
}

function Test-SPPermissionResponseKey {
    param($Response, [string]$Key)
    if ($null -eq $Response) { return $false }
    if ($Response -is [Collections.IDictionary]) { return $Response.Contains($Key) }
    return $null -ne $Response.PSObject.Properties[$Key]
}

function Get-SPPermissionCollection {
    param($Connection, [string]$Url)
    $seen = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    while ($Url) {
        if (-not $seen.Add($Url)) { throw "Repeated REST continuation: $Url" }
        $response = Invoke-SPPermissionRest $Connection $Url
        if (-not (Test-SPPermissionResponseKey $response 'value')) { throw "Expected REST collection: $Url" }
        foreach ($entry in $response.value) { $entry }
        $next = $response.'odata.nextLink'
        if (-not $next) { $next = $response.'@odata.nextLink' }
        if ($next) {
            $nextUri = [uri]::new([uri]$Url, [string]$next)
            if ($nextUri.Authority -ine ([uri]$Url).Authority) { throw 'Cross-host REST continuation refused.' }
            $Url = $nextUri.AbsoluteUri
        } else { $Url = $null }
    }
}

function Test-SPPermissionPathWithin {
    param([string]$Path, [string]$Parent)
    return ($Path -ieq $Parent -or $Path.StartsWith($Parent.TrimEnd('/') + '/', [StringComparison]::OrdinalIgnoreCase))
}

function Test-SPPermissionExcluded {
    param([string]$Key, [string[]]$ExcludePaths)
    foreach ($path in $ExcludePaths) {
        if (Test-SPPermissionPathWithin $Key $path) { return $true }
    }
    return $false
}

function Get-SPPermissionRelativePath {
    param([string]$FileRef, [string]$Prefix)
    if (-not $FileRef -or -not (Test-SPPermissionPathWithin $FileRef $Prefix)) {
        throw "Object path '$FileRef' is outside web '$Prefix'."
    }
    $FileRef.Substring($Prefix.Length).TrimStart('/')
}

function Get-SPPermissionPnPItems {
    param($Connection, $List, $Fields, [int]$PageSize)
    $names = @('ID') + @($Fields.Keys | Where-Object {
        $_ -in @('FileRef', 'FileLeafRef', 'FSObjType', 'Title', 'Created', 'Author', 'File_x0020_Size')
    })
    $viewFields = (@($names | Sort-Object -Unique | ForEach-Object {
        "<FieldRef Name='{0}'/>" -f [Security.SecurityElement]::Escape($_)
    }) -join '')
    $query = "<View Scope='RecursiveAll'><ViewFields>$viewFields</ViewFields><RowLimit Paged='TRUE'>$PageSize</RowLimit></View>"
    foreach ($item in @(Get-PnPListItem -List $List.Id -Query $query -PageSize $PageSize -Connection $Connection -ErrorAction Stop)) {
        $unique = Get-PnPProperty -ClientObject $item -Property HasUniqueRoleAssignments -Connection $Connection -ErrorAction Stop
        if ($null -eq $unique) { throw "PnP inheritance state missing for list '$($List.Title)', item $($item.Id)." }
        $values = $item.FieldValues
        $result = @{
            Id = $item.Id; FileRef = $values['FileRef']; FileLeafRef = $values['FileLeafRef']
            FileSystemObjectType = $values['FSObjType']; HasUniqueRoleAssignments = [bool]$unique
            File = $null
        }
        if ($null -eq $result.FileSystemObjectType) { throw "PnP object kind missing for '$($List.Title)', item $($item.Id)." }
        foreach ($name in @('Title', 'Created', 'Author')) {
            if (-not $Fields.ContainsKey($name)) { continue }
            $value = $values[$name]
            if ($name -eq 'Author') { $value = @{ EMail = $(if ($value) { $value.Email } else { '' }) } }
            $result[[string]$Fields[$name].EntityPropertyName] = $value
        }
        if ($Fields.ContainsKey('File_x0020_Size') -and $null -ne $values['File_x0020_Size']) {
            $result.File = @{ Length = $values['File_x0020_Size'] }
        }
        $result
    }
}

function Get-SPPermissionInventory {
    param($Connection, [string]$WebUrl, [ValidateRange(1, 5000)][int]$PageSize = 500)
    for ($attempt = 1; $attempt -le 3; $attempt++) {
        try {
            $inventory = Get-SPPermissionInventoryRead $Connection $WebUrl $PageSize
            if ($attempt -gt 1) {
                $inventory[''] | Add-Member -NotePropertyName InventoryReadReview -NotePropertyValue (
                    "Fresh web inventory succeeded on attempt $attempt after item-ID disagreement; earlier inventories discarded. Account visibility remains limited."
                )
            }
            return ,$inventory
        } catch {
            if ($attempt -eq 3 -or $_.Exception.Message -notmatch '^Inventory disagreement for ') { throw }
            Write-Warning "Inventory changed or disagreed on '$WebUrl': $($_.Exception.Message) Discarding inventory and starting fresh attempt $($attempt + 1) of 3."
        }
    }
}

function Get-SPPermissionInventoryRead {
    param($Connection, [string]$WebUrl, [ValidateRange(1, 5000)][int]$PageSize = 500)
    $webApi = "$WebUrl/_api/web"
    $web = Invoke-SPPermissionRest $Connection ($webApi + '?$select=Id,Url,ServerRelativeUrl,HasUniqueRoleAssignments')
    if ($null -eq $web.HasUniqueRoleAssignments -or -not $web.ServerRelativeUrl) { throw "Incomplete web metadata: $WebUrl" }
    if ($web.Url.TrimEnd('/') -ine $WebUrl.TrimEnd('/')) { throw "Connected to unexpected web '$($web.Url)'." }
    $prefix = $web.ServerRelativeUrl.TrimEnd('/')
    $scopes = [Collections.Generic.Dictionary[string, object]]::new([StringComparer]::OrdinalIgnoreCase)
    $scopes[''] = [pscustomobject]@{
        Key = ''; Kind = 'Web'; Api = $webApi; Unique = [bool]$web.HasUniqueRoleAssignments
        Acl = $null; Id = [string]$web.Id; Title = ''; ListKey = ''; Depth = 0
    }
    $subwebs = @(Get-SPPermissionCollection $Connection ($webApi + '/webs?$select=Id,Title,Url,ServerRelativeUrl,HasUniqueRoleAssignments'))
    foreach ($subweb in $subwebs) {
        $key = Get-SPPermissionRelativePath $subweb.ServerRelativeUrl $prefix
        if (-not $key -or $null -eq $subweb.HasUniqueRoleAssignments -or -not $subweb.Url -or $scopes.ContainsKey($key)) {
            throw "Incomplete or duplicate subweb metadata under '$WebUrl'."
        }
        # A boundary protects this web and all descendants while another web is processed.
        $scopes[$key] = [pscustomobject]@{
            Key = $key; Kind = 'Subweb'; Api = "$($subweb.Url.TrimEnd('/'))/_api/web"
            Unique = [bool]$subweb.HasUniqueRoleAssignments; Acl = $null; Id = [string]$subweb.Id
            Title = [string]$subweb.Title; ListKey = ''; Depth = 0; Url = $subweb.Url.TrimEnd('/')
        }
    }
    $lists = @(Get-SPPermissionCollection $Connection ($webApi + '/lists?$select=Id,Title,BaseType,BaseTemplate,ItemCount,HasUniqueRoleAssignments,RootFolder/ServerRelativeUrl&$expand=RootFolder'))
    foreach ($list in $lists) {
        $listKey = Get-SPPermissionRelativePath $list.RootFolder.ServerRelativeUrl $prefix
        $listApi = "$webApi/lists(guid'$($list.Id)')"
        if (-not $listKey -or $scopes.ContainsKey($listKey) -or $null -eq $list.HasUniqueRoleAssignments -or $null -eq $list.ItemCount) {
            throw "Invalid or duplicate list metadata: $($list.Title)"
        }
        $scopes[$listKey] = [pscustomobject]@{
            Key = $listKey; Kind = $(if ($list.BaseType -eq 1) { 'Library' } else { 'List' })
            Api = $listApi; Unique = [bool]$list.HasUniqueRoleAssignments; Acl = $null
            Id = [string]$list.Id; Title = [string]$list.Title; ListKey = $listKey; Depth = 0
            BaseTemplate = [int]$list.BaseTemplate; ItemCount = [int]$list.ItemCount
            Leaf = ($listKey -split '/')[-1]
        }
        $fields = [Collections.Generic.Dictionary[string, object]]::new([StringComparer]::OrdinalIgnoreCase)
        foreach ($field in @(Get-SPPermissionCollection $Connection ($listApi + '/fields?$select=InternalName,EntityPropertyName,TypeAsString'))) {
            if (-not $field.InternalName -or -not $field.EntityPropertyName) { throw "Incomplete field schema in '$listKey'." }
            $fields[$field.InternalName] = $field
        }
        $select = 'Id,FileRef,FileLeafRef,FileSystemObjectType,HasUniqueRoleAssignments'
        $expansions = [Collections.Generic.List[string]]::new()
        $optionalProperties = @{}
        foreach ($name in @('Title', 'Created', 'Author')) {
            if (-not $fields.ContainsKey($name)) {
                Write-Warning "Optional matching metadata '$name' is absent in '$listKey' on '$WebUrl'; that signal will not be used."
                continue
            }
            $property = [string]$fields[$name].EntityPropertyName
            if ($name -eq 'Author') {
                if ($fields[$name].TypeAsString -ne 'User') {
                    Write-Warning "Author is not a single-user field in '$listKey'; author matching will not be used."
                    continue
                }
                $select += ",$property/EMail"
                $expansions.Add($property)
            } else { $select += ",$property" }
            $optionalProperties[$name] = $property
        }
        if ($list.BaseType -eq 1) { $select += ',File/Length'; $expansions.Add('File') }
        $expandQuery = if ($expansions.Count) { '&$expand=' + ($expansions -join ',') } else { '' }
        $last = 0
        $itemCount = 0
        $restIds = [Collections.Generic.HashSet[int]]::new()
        $usedPnP = $false
        do {
            $url = "$listApi/items?`$select=$select$expandQuery&`$filter=Id gt $last&`$orderby=Id&`$top=$PageSize"
            try {
                $batch = @(Get-SPPermissionCollection $Connection $url)
            } catch {
                if ($last -ne 0 -or $_.Exception.Message -notmatch "Column 'ContentTypeId' does not exist") { throw }
                Write-Warning "REST item expansion failed for '$listKey' on '$WebUrl': $($_.Exception.Message). Using recursive PnP field-value enumeration and explicit inheritance reads."
                $batch = @(Get-SPPermissionPnPItems $Connection $list $fields $PageSize | Sort-Object { [int]$_.Id })
                $usedPnP = $true
                $scopes[$listKey] | Add-Member -NotePropertyName InventoryReadReview -NotePropertyValue 'REST ContentTypeId failure; enumerated using recursive PnP and explicit inheritance reads. File sizes are available only where the size field exists.'
            }
            foreach ($item in $batch) {
                $key = Get-SPPermissionRelativePath $item.FileRef $prefix
                if (-not (Test-SPPermissionPathWithin $key $listKey) -or $key -ieq $listKey -or
                    $scopes.ContainsKey($key) -or $null -eq $item.HasUniqueRoleAssignments -or [int]$item.Id -le $last) {
                    throw "Invalid or duplicate item metadata in '$listKey': $($item.Id)"
                }
                $rel = $key.Substring($listKey.Length).TrimStart('/')
                $kind = if ($item.FileSystemObjectType -eq 1) { 'Folder' } elseif ($list.BaseType -eq 1) { 'File' } else { 'Item' }
                $scopes[$key] = [pscustomobject]@{
                    Key = $key; Kind = $kind; Api = "$listApi/items($($item.Id))"
                    Unique = [bool]$item.HasUniqueRoleAssignments; Acl = $null; Id = [string]$item.Id
                    Title = $(if ($optionalProperties.ContainsKey('Title')) { [string]$item.($optionalProperties.Title) } else { '' })
                    ListKey = $listKey; Rel = $rel; Depth = ($rel -split '/').Count
                    Leaf = [string]$item.FileLeafRef; FileRef = [string]$item.FileRef
                    Ext = [IO.Path]::GetExtension([string]$item.FileLeafRef).ToLowerInvariant()
                    Created = $(if ($optionalProperties.ContainsKey('Created')) { [string]$item.($optionalProperties.Created) } else { '' })
                    Author = $(if ($optionalProperties.ContainsKey('Author')) { [string]$item.($optionalProperties.Author).EMail } else { '' })
                    Size = $(if ($kind -eq 'File' -and $null -ne $item.File -and $null -ne $item.File.Length) { [double]$item.File.Length } else { $null })
                }
                $itemCount++
                if (-not $restIds.Add([int]$item.Id)) { throw "Duplicate item ID in '$listKey'." }
            }
            if ($batch.Count) {
                $newLast = [int]$batch[-1].Id
                if ($newLast -le $last) { throw "Item pagination did not advance in '$listKey'." }
                $last = $newLast
            }
        } while (-not $usedPnP -and $batch.Count -ge $PageSize)
        if ($itemCount -ne [int]$list.ItemCount) {
            $query = "<View Scope='RecursiveAll'><ViewFields><FieldRef Name='ID'/></ViewFields><RowLimit Paged='TRUE'>$PageSize</RowLimit></View>"
            $pnpIds = [Collections.Generic.HashSet[int]]::new()
            foreach ($item in @(Get-PnPListItem -List $list.Id -Query $query -PageSize $PageSize -Connection $Connection -ErrorAction Stop)) {
                if ([int]$item.Id -le 0 -or -not $pnpIds.Add([int]$item.Id)) { throw "Invalid or duplicate PnP item ID in '$listKey'." }
            }
            if (-not $restIds.SetEquals($pnpIds)) {
                $method = if ($usedPnP) { 'initial recursive PnP' } else { 'REST' }
                throw "Inventory disagreement for '$listKey': $method read $itemCount items, verification recursive PnP read $($pnpIds.Count); list reports $($list.ItemCount). No permission changes may use this inventory."
            }
            $readMethod = if ($usedPnP) { 'Recursive PnP reads' } else { 'REST and recursive PnP' }
            $detail = "Reported count $($list.ItemCount); $readMethod agree on all $itemCount item IDs. Visibility is limited to the signed-in account."
            $scopes[$listKey] | Add-Member -NotePropertyName InventoryCountReview -NotePropertyValue $detail
            Write-Warning "Count discrepancy in '$listKey' on '$WebUrl': $detail"
        }
    }
    return ,$scopes
}

function Test-SPPermissionItemRestRejected {
    param($ErrorRecord)
    # Some items (for example in the User Information List) reject every REST items(N) URL while PnP/CSOM can read them.
    $ErrorRecord.Exception.Message -match '/items\(\d+\)[^'']*'': .*Value does not fall within the expected range'
}

function Get-SPPermissionPnPSecurable {
    param($Connection, [string]$Api)
    if ($Api -notmatch "/lists\(guid'([0-9a-fA-F-]{36})'\)/items\((\d+)\)(/FirstUniqueAncestorSecurableObject)?$") { return $null }
    $listId = $Matches[1]; $itemId = [int]$Matches[2]; $ancestor = [bool]$Matches[3]
    $item = Get-PnPListItem -List $listId -Id $itemId -Connection $Connection -ErrorAction Stop
    if ($null -eq $item -or [int]$item.Id -ne $itemId) { throw "PnP could not read item $itemId in list $listId." }
    $securable = if ($ancestor) {
        Get-PnPProperty -ClientObject $item -Property FirstUniqueAncestorSecurableObject -Connection $Connection -ErrorAction Stop
    } else { $item }
    if ($null -eq $securable) { throw "PnP returned no securable object for '$Api'." }
    [pscustomobject]@{ Item = $item; Securable = $securable }
}

function Get-SPPermissionPnPRoleAssignments {
    param($Connection, $Securable)
    $assignments = Get-PnPProperty -ClientObject $Securable -Property RoleAssignments -Connection $Connection -ErrorAction Stop
    foreach ($assignment in @($assignments)) {
        $null = Get-PnPProperty -ClientObject $assignment -Property Member, RoleDefinitionBindings -Connection $Connection -ErrorAction Stop
        [pscustomobject]@{
            Member = @{
                Id = $assignment.Member.Id; LoginName = $assignment.Member.LoginName
                Title = $assignment.Member.Title; PrincipalType = [int]$assignment.Member.PrincipalType
            }
            RoleDefinitionBindings = @($assignment.RoleDefinitionBindings | ForEach-Object {
                [pscustomobject]@{ Id = $_.Id; Name = $_.Name; Hidden = $_.Hidden; RoleTypeKind = [int]$_.RoleTypeKind }
            })
        }
    }
}

function Get-SPPermissionAcl {
    param($Connection, [string]$Api)
    try {
        $rawAssignments = @(Get-SPPermissionCollection $Connection ($Api + '/roleassignments?$expand=Member,RoleDefinitionBindings'))
    } catch {
        if (-not (Test-SPPermissionItemRestRejected $_)) { throw }
        $pnp = Get-SPPermissionPnPSecurable $Connection $Api
        if ($null -eq $pnp) { throw }
        Write-Warning "REST rejected '$Api'; reading role assignments through PnP."
        $rawAssignments = @(Get-SPPermissionPnPRoleAssignments $Connection $pnp.Securable)
    }
    foreach ($assignment in $rawAssignments) {
        if (-not $assignment.Member.LoginName -or $null -eq $assignment.RoleDefinitionBindings) {
            throw "Incomplete role assignment: $Api"
        }
        $visible = @($assignment.RoleDefinitionBindings | Where-Object {
            -not $_.Hidden -and $_.RoleTypeKind -ne 1 -and $_.Id -ne 1073741825 -and
            $_.Name -notin @('Limited Access', 'Web-Only Limited Access')
        })
        $skip = ($visible.Count -eq 0 -or $assignment.Member.LoginName -like 'SharingLinks.*' -or
            $assignment.Member.Title -like 'SharingLinks.*' -or $assignment.Member.Title -like 'Limited Access System Group*')
        [pscustomobject]@{
            MemberId = [int]$assignment.Member.Id; LoginName = [string]$assignment.Member.LoginName
            Title = [string]$assignment.Member.Title; PrincipalType = [int]$assignment.Member.PrincipalType
            Roles = @($visible | ForEach-Object Name)
            Bindings = @($assignment.RoleDefinitionBindings); Skip = $skip
        }
    }
}

function Get-SPPermissionParent {
    param($Scopes, [string]$Key)
    while ($Key) {
        $index = $Key.LastIndexOf('/')
        $Key = if ($index -lt 0) { '' } else { $Key.Substring(0, $index) }
        if ($Scopes.ContainsKey($Key)) { return $Key }
    }
    return ''
}

function Get-SPPermissionEffectiveScope {
    param($Scopes, [string]$Key)
    while ($Key -and -not $Scopes[$Key].Unique) { $Key = Get-SPPermissionParent $Scopes $Key }
    $Scopes[$Key]
}

function Get-SPPermissionAclSignature {
    param($Acl)
    $rows = foreach ($assignment in @($Acl) | Where-Object { -not $_.Skip }) {
        foreach ($role in $assignment.Roles) {
            "$($assignment.LoginName.ToLowerInvariant())`t$($role.ToLowerInvariant())"
        }
    }
    (@($rows | Sort-Object -Unique) -join "`n")
}

function Get-SPPermissionMatches {
    param($Source, $Target, [string]$MatchFile, [string]$SourceUrl, [string]$TargetUrl)
    $scopeMatches = [Collections.Generic.Dictionary[string, string]]::new([StringComparer]::OrdinalIgnoreCase)
    $used = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    if ($MatchFile) {
        $rows = @(Import-Csv -LiteralPath $MatchFile -ErrorAction Stop)
        foreach ($row in $rows) {
            foreach ($column in @('Approved', 'SourcePath', 'TargetPath', 'SourceUrl', 'TargetUrl')) {
                if ($null -eq $row.PSObject.Properties[$column]) { throw "Match CSV is missing '$column'." }
            }
            if ($row.SourceUrl.TrimEnd('/') -ine $SourceUrl.TrimEnd('/') -or $row.TargetUrl.TrimEnd('/') -ine $TargetUrl.TrimEnd('/')) {
                throw 'Match CSV belongs to different source/target webs.'
            }
            if ($row.Approved.Trim() -notmatch '^(?i:Yes|No|True|False|1|0|Y|N)?$') { throw "Invalid Approved value '$($row.Approved)'." }
            if ($row.Approved.Trim() -notmatch '^(?i:Yes|True|1|Y)$') { continue }
            $sourceKey = $row.SourcePath; $targetKey = $row.TargetPath
            if (-not $sourceKey -or -not $targetKey -or -not $Source.ContainsKey($sourceKey) -or -not $Target.ContainsKey($targetKey)) {
                throw "Approved match path is missing or invalid: '$sourceKey' -> '$targetKey'."
            }
            if ($Source[$sourceKey].Kind -eq 'Subweb' -or $Target[$targetKey].Kind -eq 'Subweb') {
                throw 'Subwebs must be paired using the checklist runner WebReview.csv, not a content match CSV.'
            }
            if ($scopeMatches.ContainsKey($sourceKey) -or -not $used.Add($targetKey)) { throw 'Approved matches must be one-to-one.' }
            if ($Source[$sourceKey].Kind -ne $Target[$targetKey].Kind) { throw "Approved match kind mismatch: $sourceKey -> $targetKey" }
            if ($Source[$sourceKey].Kind -in @('List', 'Library') -and
                $Source[$sourceKey].BaseTemplate -ne $Target[$targetKey].BaseTemplate) {
                Write-Warning "Allowing approved list template mismatch: '$SourceUrl' [$sourceKey] (BaseTemplate $($Source[$sourceKey].BaseTemplate)) -> '$TargetUrl' [$targetKey] (BaseTemplate $($Target[$targetKey].BaseTemplate)). Only permissions are matched; list schemas are not converted."
            }
            if ($Source[$sourceKey].Kind -eq 'File' -and $Source[$sourceKey].Ext -ine $Target[$targetKey].Ext) {
                throw "File extension mismatch: $sourceKey"
            }
            if ($sourceKey -ine $targetKey -and ($Target.ContainsKey($sourceKey) -or $Source.ContainsKey($targetKey))) {
                throw "Approved alias conflicts with an exact path: $sourceKey -> $targetKey"
            }
            $scopeMatches[$sourceKey] = $targetKey
        }
    }
    foreach ($key in $Source.Keys) {
        if ($Source[$key].Kind -eq 'Subweb' -or ($Target.ContainsKey($key) -and $Target[$key].Kind -eq 'Subweb')) { continue }
        if ($scopeMatches.ContainsKey($key) -or -not $Target.ContainsKey($key) -or $used.Contains($key)) { continue }
        if ($Source[$key].Kind -ne $Target[$key].Kind) { throw "Exact-path kind mismatch: $key" }
        if ($Source[$key].Kind -in @('List', 'Library') -and $Source[$key].BaseTemplate -ne $Target[$key].BaseTemplate) {
            Write-Warning "Allowing exact-path list template mismatch: '$SourceUrl' [$key] (BaseTemplate $($Source[$key].BaseTemplate)) -> '$TargetUrl' [$key] (BaseTemplate $($Target[$key].BaseTemplate)). Only permissions are matched; list schemas are not converted."
        }
        $scopeMatches[$key] = $key
        [void]$used.Add($key)
    }
    return ,$scopeMatches
}

function Export-SPPermissionCsv {
    param([object[]]$Rows, [string[]]$Columns, [string]$Path)
    if ($Rows.Count) { $Rows | Select-Object $Columns | Export-Csv -LiteralPath $Path -NoTypeInformation -Encoding utf8 }
    else { (($Columns | ForEach-Object { '"' + $_.Replace('"', '""') + '"' }) -join ',') | Set-Content -LiteralPath $Path -Encoding utf8 }
}

function Get-SPNameRatio {
    param([string]$Left, [string]$Right)
    if ($Left -ceq $Right) { return 1.0 }
    $length = [math]::Max($Left.Length, $Right.Length)
    if (-not $Left -or -not $Right) { return 0.0 }
    $previous = [int[]](0..$Right.Length)
    $current = [int[]]::new($Right.Length + 1)
    for ($i = 1; $i -le $Left.Length; $i++) {
        $current[0] = $i
        for ($j = 1; $j -le $Right.Length; $j++) {
            $cost = if ($Left[$i - 1] -ceq $Right[$j - 1]) { 0 } else { 1 }
            $current[$j] = [math]::Min([math]::Min($current[$j - 1] + 1, $previous[$j] + 1), $previous[$j - 1] + $cost)
        }
        $swap = $previous; $previous = $current; $current = $swap
    }
    1.0 - $previous[$Right.Length] / $length
}

function Get-SPNormalizedName {
    param([string]$Name)
    $nameOnly = [IO.Path]::GetFileNameWithoutExtension($Name).ToLowerInvariant()
    $nameOnly = $nameOnly -replace '\(\d+\)', ' ' -replace '\b(copy|final|rev(ised)?|v\d+(\.\d+)*|version\s*\d+)\b', ' '
    (($nameOnly -replace '[^\p{L}\p{Nd}]+', ' ') -replace '\s+', ' ').Trim()
}

function Get-SPNameSimilarity {
    param([string]$Left, [string]$Right)
    $leftName = Get-SPNormalizedName $Left; $rightName = Get-SPNormalizedName $Right
    if (-not $leftName -or -not $rightName) { return 0.0 }
    $leftSorted = (($leftName -split ' ') | Sort-Object) -join ' '
    $rightSorted = (($rightName -split ' ') | Sort-Object) -join ' '
    [math]::Max((Get-SPNameRatio $leftName $rightName), (Get-SPNameRatio $leftSorted $rightSorted))
}
