<#
.SYNOPSIS
    Per-document pre-migration report for a single site being promoted out of
    the classic subsite tree: every document's current URL, its future URL
    after promotion, owner (Created By), last modified date/person, which AD
    groups and SharePoint groups have access, who it has been directly shared
    with, and the exact embedded URLs that need fixing.

.DESCRIPTION
    Run with no arguments to pick the site interactively from a numbered menu
    of the 37 subsites in plan/WHA_Subsite_Migration_Worksheet_v2.xlsx (a
    static list - no root-site enumeration, no extra sign-in; paste a full URL
    at the prompt for anything not listed). The promotion target defaults to
    /sites/<site name> (press Enter to accept, or type a different path), and
    the output CSVs are named after the site (e.g. hr-document-access.csv).
    All of this can also be supplied explicitly via -SiteUrl / -TargetSitePath
    / -*OutputPath for scripted runs.

    Scans the chosen site and its subwebs. Produces two CSVs joined by
    DocumentUrl:

      1. Access report (one row per document, every file type):
         Site, Library, DocumentUrl (canonical), DocumentBrowserUrl (the
         address-bar form with /:w:/r/ viewer prefix and ?d=<guid> redirect
         parameter), FutureDocumentUrl, FolderName (the library's DISPLAY
         title - which can differ from the URL segment - plus subfolder path),
         FolderUrl (classic Forms/AllItems.aspx?CID=<list guid> view URL),
         FileName, Owner/Created/Modified/LastModifiedBy, PermissionMode
         (Inherited-from-site / Inherited-from-library / Unique),
         ADGroupsWithAccess, SharePointGroupsWithAccess (expanded one level so
         AD groups nested inside SP groups are surfaced), DirectUserAccess,
         SharingLinks (link kind + named recipients), SharedWithUsers,
         ScannedForLinks, BrokenLinkCount.

      2. Links report (one row per EMBEDDED URL per document - every URL
         found, verbatim, even benign external ones):
         DocumentUrl, FutureDocumentUrl, ReferenceKind, Confidence, RiskType
         (Absolute-Internal / Server-Relative / Modern-Site-Reference, plus
         already-broken-today categories Local-File-Link and
         Legacy-Intranet-Host; "None" for links that won't break), EmbeddedUrl (exactly as stored
         in the document), CleanedUrl (HTML-decoded/trimmed, used for
         classification), SuggestedReplacement - filled in automatically when
         a risky link points at content inside this same site (rewrite to the
         /sites/ path); left blank when it points at another classic site
         whose target address isn't decided yet.

    Scope decisions (agreed 2026-07-30):
      - Documents only. Pages, navigation, and list columns are covered by the
        tenant-wide Find-DocumentLinkRisks.ps1 scan and are not repeated here.
      - AD groups are reported by NAME only; membership expansion is done
        on-prem with Get-ADGroupMember, not via Graph.
      - Permissions: most documents inherit from the library/site, so their
        access columns repeat the site-level answer with PermissionMode noting
        the inheritance. Documents with broken inheritance get their own
        item-level breakdown - that is where individual sharing lives (modern
        sharing links appear as SharingLinks.* role assignments whose group
        membership holds the named recipients).

    Same caveats as the original scanner: legacy .doc/.xls/.ppt are not parsed
    for links (flagged ScannedForLinks=No-legacy); PDFs only surface absolute
    http(s) URLs, not server-relative paths. Results append to the CSVs as
    found, so a crash partway through keeps prior progress.

.PARAMETER RootUrl
    Root site collection URL, prefixed onto menu picks when -SiteUrl is not
    supplied. Default https://whallc.sharepoint.com

.PARAMETER SiteUrl
    The classic site being promoted. Omit to pick from an interactive menu.

.PARAMETER TargetSitePath
    Server-relative path the site will live at after promotion. Omit to be
    prompted, defaulting to /sites/<site name>.

.PARAMETER ClientId
    Entra ID App Registration Client ID for interactive sign-in.

.PARAMETER AccessOutputPath
    Per-document access CSV. Default .\<site name>-document-access.csv

.PARAMETER LinksOutputPath
    Per-link fix CSV. Default .\<site name>-links-to-fix.csv

.EXAMPLE
    # Interactive: pick the site from a menu, accept the derived target path
    .\Get-DocumentAccessReport.ps1 -ClientId "..."

.EXAMPLE
    # Scripted: everything explicit, no prompts
    .\Get-DocumentAccessReport.ps1 -SiteUrl "https://whallc.sharepoint.com/HR" -TargetSitePath "/sites/HR" -ClientId "..."
#>

param(
    [Parameter(Mandatory = $false)]
    [string]$RootUrl = "https://whallc.sharepoint.com",

    [Parameter(Mandatory = $false)]
    [string]$SiteUrl,

    [Parameter(Mandatory = $false)]
    [string]$TargetSitePath,

    [Parameter(Mandatory = $false)]
    [string]$ClientId,

    [Parameter(Mandatory = $false)]
    [string]$AccessOutputPath,

    [Parameter(Mandatory = $false)]
    [string]$LinksOutputPath
)

$ErrorActionPreference = "Stop"
$RootUrl = $RootUrl.TrimEnd('/')

$SupportedExtensions = @('.docx', '.xlsx', '.pptx', '.pdf')
$LegacyExtensions = @('.doc', '.xls', '.ppt')
$SkipListTitles = @('Form Templates', 'Preservation Hold Library')

if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
    Write-Host "PnP.PowerShell module not found - installing for current user..."
    Install-Module -Name PnP.PowerShell -Scope CurrentUser -Force -AllowClobber
}
Import-Module PnP.PowerShell

# --- URL construction helpers ------------------------------------------------

# Office web-viewer prefix per extension: /:w:/r/... for Word, :x: Excel,
# :p: PowerPoint, :b: PDF. Same letter is reused in the ?d=<letter><guid>
# redirect parameter SharePoint shows in the address bar.
$DocLinkPrefixByExt = @{
    '.docx' = 'w'; '.doc' = 'w'
    '.xlsx' = 'x'; '.xls' = 'x'
    '.pptx' = 'p'; '.ppt' = 'p'
    '.pdf'  = 'b'
}

function Get-EncodedServerRelativePath {
    param([string]$Path)
    (($Path -split '/') | ForEach-Object { [Uri]::EscapeDataString($_) }) -join '/'
}

function Get-BrowserUrl {
    # The URL as it appears when a user navigates to the document:
    # https://host/:w:/r/<path>.docx?d=w<uniqueid>&csf=1&web=1
    # (the per-share "e=" token is deliberately omitted - it's minted per
    # sharing action, not derivable, and the link resolves without it).
    # Non-Office/PDF types have no viewer form; they get the canonical URL.
    param([string]$FileRef, [string]$Ext, $UniqueId)
    $encPath = Get-EncodedServerRelativePath -Path $FileRef
    $prefix = $DocLinkPrefixByExt[$Ext]
    if (-not $prefix -or -not $UniqueId) { return $TenantBase + $encPath }
    $guid = ([Guid]$UniqueId).ToString('N')
    return "$TenantBase/:${prefix}:/r$encPath" + "?d=$prefix$guid&csf=1&web=1"
}

function Get-FolderInfo {
    # Folder NAME is the library's display Title (which can diverge from the
    # URL segment), plus any subfolder path. Folder URL is the classic library
    # view: <libraryRoot>/Forms/AllItems.aspx?CID=<list guid>, plus RootFolder
    # when the document sits in a subfolder.
    param([string]$FileDirRef, $List, [string]$LibraryRootPath)

    $sub = ""
    if ($FileDirRef.Length -gt $LibraryRootPath.Length) {
        $sub = $FileDirRef.Substring($LibraryRootPath.Length).Trim('/')
    }
    $name = if ($sub) { "$($List.Title) / $($sub -replace '/', ' / ')" } else { $List.Title }

    $url = $TenantBase + (Get-EncodedServerRelativePath -Path "$LibraryRootPath/Forms/AllItems.aspx") + "?CID=$($List.Id.ToString().ToLower())"
    if ($sub) { $url += "&RootFolder=" + [Uri]::EscapeDataString($FileDirRef) }

    return [PSCustomObject]@{ Name = $name; Url = $url }
}

# --- URL rewrite helpers -----------------------------------------------------

function Get-FutureUrl {
    param([string]$ServerRelativeFileRef)
    # /clinicians/Shared Documents/x.docx -> https://host/sites/clinicians/Shared Documents/x.docx
    if ($ServerRelativeFileRef -and $ServerRelativeFileRef.ToLower().StartsWith($SourceSitePath.ToLower())) {
        return $TenantBase + $TargetSitePath + $ServerRelativeFileRef.Substring($SourceSitePath.Length)
    }
    return ""
}

# Old on-prem SharePoint hostnames whose path structure matches the tenant
# (http://centralpark/Clinicians/... = whallc.sharepoint.com/Clinicians/...).
# Links to these hosts are already dead; treat their paths as tenant paths so
# ones pointing into the scanned site get a rewrite suggestion.
$LegacyHostnames = @('centralpark')

function Get-SuggestedReplacement {
    param([string]$Value)
    # Only links pointing INTO this site get an automatic rewrite; links to
    # other classic sites stay blank until their target names are decided.
    foreach ($h in $LegacyHostnames) {
        if ($Value -match "^(?i)https?://$([regex]::Escape($h))(?=[/:?#]|$)") {
            $Value = $TenantBase + ($Value -replace "^(?i)https?://$([regex]::Escape($h))", '')
            break
        }
    }
    # Matches plain URLs AND the Office viewer form https://host/:w:/r/<path>
    # (everything after /r is an ordinary server-relative path). Viewer links
    # that locate the file by GUID (Doc.aspx?sourcedoc={...}, ?d=...) keep
    # their GUID in the rewrite - those only stay valid if the migration
    # preserves document UniqueIds (ShareGate does by default for documents).
    $pattern = '^(?i)' + [regex]::Escape($TenantBase) + '(/:[a-z]:/[rtu])?' + [regex]::Escape($SourceSitePath) + '(?=$|[/?#])'
    $m = [regex]::Match($Value, $pattern)
    if ($m.Success) {
        return $TenantBase + $m.Groups[1].Value + $TargetSitePath + $Value.Substring($m.Length)
    }
    if ($Value.ToLower().StartsWith($SourceSitePath.ToLower())) {
        $rest = $Value.Substring($SourceSitePath.Length)
        if ($rest -eq "" -or $rest[0] -in '/', '?', '#') {
            return $TargetSitePath + $rest
        }
    }
    return ""
}

# --- link-risk helpers (same classification as Find-DocumentLinkRisks.ps1) ---

function Test-IsModernManagedPath {
    param([string]$Value)
    $path = $Value
    if ($Value -match [regex]::Escape($TenantHostname)) {
        try { $path = ([Uri]$Value).AbsolutePath } catch { }
    }
    return $path -match '^/(sites|teams)/'
}

function Test-IsDeprecatedPortalPath {
    param([string]$Value)
    $path = $Value
    if ($Value -match [regex]::Escape($TenantHostname)) {
        try { $path = ([Uri]$Value).AbsolutePath } catch { }
    }
    return $path -match '^/portals/'
}

function Test-IsSystemAsset {
    param([string]$Value)
    if ($Value -notmatch '_layouts/') { return $false }
    $pathOnly = ($Value -split '\?')[0]
    $ext = [System.IO.Path]::GetExtension($pathOnly).ToLower()
    if (@('.gif', '.png', '.jpg', '.jpeg', '.css', '.js', '.ico', '.woff', '.woff2', '.svg') -contains $ext) {
        return $true
    }
    if ($ext -eq '.aspx' -and ($Value -notmatch '\?')) { return $true }
    return $false
}

function Get-ReferenceRisk {
    param([string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return $null }
    if (Test-IsSystemAsset -Value $Value) { return $null }
    if (Test-IsDeprecatedPortalPath -Value $Value) { return $null }

    # Links that are broken TODAY, before any migration - flagged for owner
    # cleanup (confirmed in scope 7/30):
    # file:// and UNC links point at local drives or old file servers and are
    # dead for anyone but the original author's machine.
    if ($Value -match '^(?i)file:' -or $Value -match '^\\\\') { return "Local-File-Link" }
    # http(s) to a single-label hostname (no dot, e.g. http://centralpark) is
    # an old intranet server name - unreachable outside the LAN if it
    # resolves at all.
    if ($Value -match '^(?i)https?://[^./:?#]+([:/?#]|$)') { return "Legacy-Intranet-Host" }

    $isInternal = $Value -match [regex]::Escape($TenantHostname)
    $isServerRelative = $Value -match '^/[^/]'
    if (-not $isInternal -and -not $isServerRelative) { return $null }

    if (Test-IsModernManagedPath -Value $Value) { return "Modern-Site-Reference" }
    if ($isInternal) { return "Absolute-Internal" }
    return "Server-Relative"
}

function Get-CleanedValue {
    param([string]$Value)
    $decoded = [System.Net.WebUtility]::HtmlDecode($Value)
    return $decoded.TrimEnd("'", '"', ':', ',', ')', ']', '}', '.', ';')
}

function Get-OfficeDocReferences {
    param([System.IO.Stream]$Stream)

    $found = New-Object System.Collections.Generic.List[Object]
    $archive = [System.IO.Compression.ZipArchive]::new($Stream, [System.IO.Compression.ZipArchiveMode]::Read)
    try {
        foreach ($entry in $archive.Entries) {
            if ($entry.FullName -like "*_rels/*.rels") {
                $s = $entry.Open()
                $r = New-Object System.IO.StreamReader($s)
                $xmlText = $r.ReadToEnd()
                $r.Close(); $s.Close()
                try {
                    [xml]$xml = $xmlText
                    foreach ($rel in $xml.Relationships.Relationship) {
                        if ($rel.TargetMode -eq "External" -and $rel.Target) {
                            $found.Add([PSCustomObject]@{ Value = $rel.Target; Confidence = "High"; Kind = "hyperlink" })
                        }
                    }
                } catch { }
            }
            elseif ($entry.FullName -match '\.xml$' -and
                    ($entry.FullName -match 'document\.xml$|sheet\d*\.xml$|slide\d*\.xml$|header\d*\.xml$|footer\d*\.xml$|notesSlide\d*\.xml$')) {
                $s = $entry.Open()
                $r = New-Object System.IO.StreamReader($s)
                $text = $r.ReadToEnd()
                $r.Close(); $s.Close()
                $plain = [regex]::Replace($text, '<[^>]+>', ' ')
                foreach ($m in [regex]::Matches($plain, 'https?://[^\s"<>]+')) {
                    $found.Add([PSCustomObject]@{ Value = $m.Value; Confidence = "Medium"; Kind = "plaintext" })
                }
            }
        }
    } finally {
        $archive.Dispose()
    }
    return $found | Sort-Object Value, Kind -Unique
}

# Hosts that appear inside every PDF's internal metadata (XMP namespaces,
# RDF/Dublin Core schemas, ICC/PDF-X/PDF-A identifiers) - format plumbing,
# not links anyone embedded. In the clinicians pilot these were 93% of all
# extracted PDF URLs, so they're dropped at extraction.
$PdfNoiseHosts = @('ns.adobe.com', 'www.w3.org', 'w3.org', 'purl.org',
                   'www.iec.ch', 'iec.ch', 'www.color.org', 'www.npes.org', 'www.aiim.org')

function Get-PdfReferences {
    param([byte[]]$Bytes)

    $found = New-Object System.Collections.Generic.List[Object]
    $text = [System.Text.Encoding]::Latin1.GetString($Bytes)
    foreach ($m in [regex]::Matches($text, 'https?://[^\s\)"<>\\]+')) {
        $hostMatch = [regex]::Match($m.Value, '^https?://([^/:?#]+)')
        if ($hostMatch.Success -and ($PdfNoiseHosts -contains $hostMatch.Groups[1].Value.ToLower())) { continue }
        $found.Add([PSCustomObject]@{ Value = $m.Value; Confidence = "Low"; Kind = "plaintext-pdf" })
    }
    return $found | Sort-Object Value, Kind -Unique
}

# --- permission helpers ------------------------------------------------------

# Cache of expanded SP-group membership keyed by group Id, so the same group
# (e.g. "Clinicians Members") is only expanded once no matter how many
# uniquely-permissioned items reference it.
$GroupMemberCache = @{}

function Get-GroupMemberSummary {
    param($Group)
    if ($GroupMemberCache.ContainsKey($Group.Id)) { return $GroupMemberCache[$Group.Id] }

    $adGroups = New-Object System.Collections.Generic.List[string]
    $users = New-Object System.Collections.Generic.List[string]
    try {
        $members = Get-PnPProperty -ClientObject $Group -Property Users
        foreach ($u in $members) {
            if ($u.PrincipalType -eq 'SecurityGroup') {
                $adGroups.Add($u.Title)
            } elseif ($u.PrincipalType -eq 'User') {
                $label = if ($u.Email) { "$($u.Title) <$($u.Email)>" } else { $u.Title }
                $users.Add($label)
            }
        }
    } catch {
        $users.Add("(could not expand: $($_.Exception.Message))")
    }
    $summary = [PSCustomObject]@{ ADGroups = $adGroups; Users = $users }
    $GroupMemberCache[$Group.Id] = $summary
    return $summary
}

function Get-PermissionBreakdown {
    <#
        Given a securable object (web, list, or list item), enumerate its role
        assignments and return the access picture:
          ADGroups      - AD security groups granted access (directly, or nested
                          one level inside a SharePoint group), with role
          SPGroups      - SharePoint groups with role and expanded membership
          DirectUsers   - individual users granted access directly, with role
          SharingLinks  - modern sharing links (kind + named recipients)
        "Limited Access" assignments are skipped - they're plumbing, not access.
    #>
    param($Securable)

    $adGroups = New-Object System.Collections.Generic.List[string]
    $spGroups = New-Object System.Collections.Generic.List[string]
    $directUsers = New-Object System.Collections.Generic.List[string]
    $sharingLinks = New-Object System.Collections.Generic.List[string]

    $roleAssignments = Get-PnPProperty -ClientObject $Securable -Property RoleAssignments

    foreach ($ra in $roleAssignments) {
        $member = Get-PnPProperty -ClientObject $ra -Property Member
        $bindings = Get-PnPProperty -ClientObject $ra -Property RoleDefinitionBindings

        $roles = @($bindings | Where-Object { $_.Name -ne 'Limited Access' } | Select-Object -ExpandProperty Name)
        if ($roles.Count -eq 0) { continue }
        $roleText = $roles -join '/'

        if ($member.PrincipalType -eq 'SecurityGroup') {
            $adGroups.Add("$($member.Title) [$roleText]")
        }
        elseif ($member.PrincipalType -eq 'User') {
            $label = if ($member.Email) { "$($member.Title) <$($member.Email)>" } else { $member.Title }
            $directUsers.Add("$label [$roleText]")
        }
        elseif ($member.PrincipalType -eq 'SharePointGroup') {
            if ($member.LoginName -like 'SharingLinks.*') {
                # SharingLinks.<itemGuid>.<kind>.<linkGuid> - the group's members
                # are the people the link was issued to (specific-people links)
                # or who have redeemed it (org-wide links).
                $kind = ($member.LoginName -split '\.')[2]
                $m = Get-GroupMemberSummary -Group $member
                $who = @($m.Users + $m.ADGroups)
                $whoText = if ($who.Count -gt 0) { $who -join ', ' } else { '(no named recipients yet)' }
                $sharingLinks.Add("$kind link [$roleText] -> $whoText")
            } else {
                $m = Get-GroupMemberSummary -Group $member
                foreach ($g in $m.ADGroups) { $adGroups.Add("$g [via SP group '$($member.Title)', $roleText]") }
                $memberText = if ($m.Users.Count -gt 0) { " {members: $($m.Users -join ', ')}" } else { "" }
                $spGroups.Add("$($member.Title) [$roleText]$memberText")
            }
        }
    }

    return [PSCustomObject]@{
        ADGroups     = ($adGroups | Sort-Object -Unique) -join '; '
        SPGroups     = ($spGroups | Sort-Object -Unique) -join '; '
        DirectUsers  = ($directUsers | Sort-Object -Unique) -join '; '
        SharingLinks = ($sharingLinks | Sort-Object -Unique) -join '; '
    }
}

# --- output helpers ----------------------------------------------------------

function Write-AccessRow {
    param($Row)
    $Row | Export-Csv -Path $AccessOutputPath -NoTypeInformation -Append
}

function Write-LinkRow {
    param(
        [string]$DocumentUrl,
        [string]$FutureDocumentUrl,
        [string]$Kind,
        [string]$Confidence,
        [string]$RiskType,
        [string]$EmbeddedUrl,
        [string]$CleanedUrl,
        [string]$SuggestedReplacement
    )
    [PSCustomObject]@{
        DocumentUrl          = $DocumentUrl
        FutureDocumentUrl    = $FutureDocumentUrl
        ReferenceKind        = $Kind
        Confidence           = $Confidence
        RiskType             = $RiskType
        EmbeddedUrl          = $EmbeddedUrl
        CleanedUrl           = $CleanedUrl
        SuggestedReplacement = $SuggestedReplacement
    } | Export-Csv -Path $LinksOutputPath -NoTypeInformation -Append
}

# --- site selection ----------------------------------------------------------

if (-not $SiteUrl) {
    # Static menu from plan/WHA_Subsite_Migration_Worksheet_v2.xlsx (the 37
    # subsites with URL paths in the migration worksheet). No live enumeration:
    # the root-level subweb walk was slow and kept hitting access-denied.
    # Update this list here if the worksheet changes; sites not listed can
    # still be scanned by pasting their full URL at the prompt (or -SiteUrl).
    $PromotionSites = @(
        @{ Path = '/IT&HIM';                    Name = 'IT & HIM Department' }
        @{ Path = '/train';                     Name = 'Training' }
        @{ Path = '/Clinicians';                Name = 'Clinicians' }
        @{ Path = '/HR';                        Name = 'Human Resources' }
        @{ Path = '/Safety';                    Name = 'Safety' }
        @{ Path = '/whasites';                  Name = 'WHA Sites' }
        @{ Path = '/PFS';                       Name = 'Patient Financial Services' }
        @{ Path = '/CS';                        Name = 'Clinical Services' }
        @{ Path = '/Epic';                      Name = 'Epic' }
        @{ Path = '/Accounting';                Name = 'Accounting' }
        @{ Path = '/Forms';                     Name = 'Forms' }
        @{ Path = '/QI';                        Name = 'Quality Improvement' }
        @{ Path = '/Committees';                Name = 'Committees' }
        @{ Path = '/HelpDeskSelfServicePortal'; Name = 'Help Desk Self-Service Portal' }
        @{ Path = '/LLCM';                      Name = 'LLC Members' }
        @{ Path = '/Int';                       Name = 'Interiors Committee' }
        @{ Path = '/TPE';                       Name = 'The Patient Experience' }
        @{ Path = '/LegacySharing';             Name = 'Legacy Sharing' }
        @{ Path = '/hipaa';                     Name = 'HIPAA' }
        @{ Path = '/himcommittee';              Name = 'HIM Committee' }
        @{ Path = '/CRL';                       Name = 'Clinician Resource Library' }
        @{ Path = '/Comp';                      Name = 'Compensation Committee' }
        @{ Path = '/IT-Internal';               Name = 'IT Internal' }
        @{ Path = '/Ops';                       Name = 'Operations Committee' }
        @{ Path = '/Private';                   Name = 'Private' }
        @{ Path = '/Fin';                       Name = 'Finance Committee' }
        @{ Path = '/emr';                       Name = 'EMR' }
        @{ Path = '/PFTC';                      Name = 'Patient Focus Team Care' }
        @{ Path = '/SP';                        Name = 'Strategic Plan' }
        @{ Path = '/ClinicianResources';        Name = 'Clinician Resources' }
        @{ Path = '/IS';                        Name = 'Information Security' }
        @{ Path = '/EMRWeeklyMeeting';          Name = 'Weekly EMR Meeting (deprecated workspace)' }
        @{ Path = '/HIMCommittee';              Name = 'HIM Committee (Meeting Workspace, deprecated)' }
        @{ Path = '/RecruitmentCommittee';      Name = 'Recruitment Committee' }
        @{ Path = '/CtrlAltDelight_Newsletter'; Name = 'Ctrl + Alt + Delight Newsletter' }
        @{ Path = '/HIM';                       Name = 'HIM (found in 7/17 scan)' }
        @{ Path = '/whaiis01';                  Name = 'whaiis01 (found in 7/17 scan)' }
    )

    Write-Host ""
    for ($i = 0; $i -lt $PromotionSites.Count; $i++) {
        Write-Host ("  {0,3}. {1,-28} {2}" -f ($i + 1), $PromotionSites[$i].Path, $PromotionSites[$i].Name)
    }

    while (-not $SiteUrl) {
        $choice = Read-Host "`nSite to scan (number from the list, or paste a full site URL)"
        if ($choice -match '^\d+$' -and [int]$choice -ge 1 -and [int]$choice -le $PromotionSites.Count) {
            $SiteUrl = $RootUrl + $PromotionSites[[int]$choice - 1].Path
        } elseif ($choice -match '^https?://') {
            $SiteUrl = $choice
        } elseif ($choice) {
            Write-Host "  not a valid number or URL - try again."
        }
    }
}

$SiteUrl = $SiteUrl.TrimEnd('/')
$siteUri = [Uri]$SiteUrl
$TenantBase = $siteUri.GetLeftPart([UriPartial]::Authority)
$TenantHostname = $siteUri.Host
$SourceSitePath = $siteUri.AbsolutePath.TrimEnd('/')
$siteLeaf = ($SourceSitePath.Trim('/') -split '/')[-1]

if (-not $TargetSitePath) {
    $defaultTarget = "/sites/$siteLeaf"
    $entered = Read-Host "Target path after promotion [$defaultTarget]"
    $TargetSitePath = if ([string]::IsNullOrWhiteSpace($entered)) { $defaultTarget } else { $entered }
}
$TargetSitePath = "/" + $TargetSitePath.Trim('/')

if (-not $AccessOutputPath) { $AccessOutputPath = ".\" + $siteLeaf.ToLower() + "-document-access.csv" }
if (-not $LinksOutputPath) { $LinksOutputPath = ".\" + $siteLeaf.ToLower() + "-links-to-fix.csv" }

Write-Host "`nSite:    $SiteUrl"
Write-Host "Target:  $TargetSitePath"
Write-Host "Reports: $AccessOutputPath / $LinksOutputPath"

# --- connect and enumerate webs ----------------------------------------------

Write-Host "`nConnecting to $SiteUrl (interactive sign-in)..."
if ($ClientId) {
    Connect-PnPOnline -Url $SiteUrl -Interactive -ClientId $ClientId
} else {
    Connect-PnPOnline -Url $SiteUrl -Interactive
}

$rootWeb = Get-PnPWeb -Includes Url
$allWebs = New-Object System.Collections.Generic.List[string]
$allWebs.Add($rootWeb.Url)
try {
    $subWebs = Get-PnPSubWeb -Recurse -Includes Url
    foreach ($w in $subWebs) { $allWebs.Add($w.Url) }
} catch {
    Write-Warning "could not enumerate subwebs: $($_.Exception.Message)"
}
Write-Host "$($allWebs.Count) web(s) in scope under $SiteUrl."

foreach ($path in @($AccessOutputPath, $LinksOutputPath)) {
    if (Test-Path $path) { Remove-Item -Path $path -Force }
}

$stats = [PSCustomObject]@{
    DocsRecorded    = 0
    UniquePermDocs  = 0
    LinksFlagged    = 0
    ReplacementsSet = 0
    FilesScanned    = 0
    FoldersSkipped  = 0
    LibrariesRead   = 0
}

# --- main loop ---------------------------------------------------------------

foreach ($webUrl in $allWebs) {
    Write-Host "`nWeb: $webUrl"
    try {
        if ($ClientId) {
            $webConn = Connect-PnPOnline -Url $webUrl -Interactive -ClientId $ClientId -ReturnConnection
        } else {
            $webConn = Connect-PnPOnline -Url $webUrl -Interactive -ReturnConnection
        }
    } catch {
        Write-Warning "  could not connect to $webUrl : $($_.Exception.Message)"
        continue
    }

    # Site-level permission picture, computed once - inheriting documents
    # reference this instead of re-enumerating per item.
    Write-Host "  reading site-level permissions..."
    $web = Get-PnPWeb -Connection $webConn
    $webPerms = Get-PermissionBreakdown -Securable $web

    $lists = $null
    try {
        $lists = Get-PnPList -Connection $webConn -Includes BaseTemplate, Hidden, Title, ItemCount, HasUniqueRoleAssignments, Id, RootFolder.ServerRelativeUrl
    } catch {
        Write-Warning "  could not read lists for $webUrl : $($_.Exception.Message)"
        continue
    }

    foreach ($list in $lists) {
        if ($list.Hidden) { continue }
        if ($SkipListTitles -contains $list.Title) { continue }
        if ($list.ItemCount -eq 0) { continue }
        if ($list.BaseTemplate -ne 101) { continue }   # document libraries only

        Write-Host "  library: $($list.Title) ($($list.ItemCount) items)"
        $stats.LibrariesRead++

        $libraryRootPath = $list.RootFolder.ServerRelativeUrl.TrimEnd('/')

        $listPerms = $null
        $inheritedSource = "Inherited-from-site"
        $inheritedPerms = $webPerms
        if ($list.HasUniqueRoleAssignments) {
            Write-Host "    (library has unique permissions - reading them)"
            $listPerms = Get-PermissionBreakdown -Securable $list
            $inheritedSource = "Inherited-from-library"
            $inheritedPerms = $listPerms
        }

        $fields = @("FileRef", "FileLeafRef", "FileDirRef", "UniqueId", "FSObjType", "Modified", "Created", "Author", "Editor", "SharedWithUsers")
        try {
            $items = Get-PnPListItem -List $list.Title -Connection $webConn -PageSize 500 -Fields $fields
        } catch {
            # SharedWithUsers isn't guaranteed to exist on every library
            $fields = $fields | Where-Object { $_ -ne "SharedWithUsers" }
            $items = Get-PnPListItem -List $list.Title -Connection $webConn -PageSize 500 -Fields $fields
        }

        $itemIndex = 0
        foreach ($item in $items) {
            $itemIndex++
            if ($itemIndex % 50 -eq 0) { Write-Host "    ...$itemIndex / $($items.Count)" }
            if ($item.FieldValues["FSObjType"] -eq 1) { $stats.FoldersSkipped++; continue }

            $fileRef = $item.FieldValues["FileRef"]
            $fileLeaf = $item.FieldValues["FileLeafRef"]
            if (-not $fileRef) { continue }

            $ext = [System.IO.Path]::GetExtension($fileLeaf).ToLower()
            $docUrl = $TenantBase + $fileRef
            $futureUrl = Get-FutureUrl -ServerRelativeFileRef $fileRef
            $browserUrl = Get-BrowserUrl -FileRef $fileRef -Ext $ext -UniqueId $item.FieldValues["UniqueId"]
            $folder = Get-FolderInfo -FileDirRef $item.FieldValues["FileDirRef"] -List $list -LibraryRootPath $libraryRootPath

            $authorValue = $item.FieldValues["Author"]
            $owner = if ($authorValue -and $authorValue.LookupValue) { $authorValue.LookupValue } else { "" }
            $editorValue = $item.FieldValues["Editor"]
            $editor = if ($editorValue -and $editorValue.LookupValue) { $editorValue.LookupValue } else { "" }

            $sharedWith = ""
            $swValue = $item.FieldValues["SharedWithUsers"]
            if ($swValue) {
                $sharedWith = (@($swValue) | ForEach-Object {
                    if ($_.Email) { "$($_.LookupValue) <$($_.Email)>" } else { $_.LookupValue }
                }) -join '; '
            }

            # Permission mode: one extra round trip per item to learn whether
            # inheritance is broken; full enumeration only when it is.
            $permMode = $inheritedSource
            $perms = $inheritedPerms
            try {
                $hasUnique = Get-PnPProperty -ClientObject $item -Property HasUniqueRoleAssignments
                if ($hasUnique) {
                    $permMode = "Unique"
                    $perms = Get-PermissionBreakdown -Securable $item
                    $stats.UniquePermDocs++
                }
            } catch {
                $permMode = "Unknown ($($_.Exception.Message))"
            }

            # Link scan (documents only, same file-type support as before).
            # EVERY extracted URL is written verbatim (EmbeddedUrl), even
            # benign external ones - RiskType classifies instead of filtering.
            $scanned = "No-othertype"
            $brokenCount = ""
            if ($LegacyExtensions -contains $ext) {
                $scanned = "No-legacy"
            }
            elseif ($SupportedExtensions -contains $ext) {
                try {
                    $ms = Get-PnPFile -Url $fileRef -AsMemoryStream -Connection $webConn
                    $stats.FilesScanned++
                    $refs = if ($ext -eq '.pdf') {
                        Get-PdfReferences -Bytes $ms.ToArray()
                    } else {
                        Get-OfficeDocReferences -Stream $ms
                    }
                    $scanned = "Yes"
                    $brokenCount = 0
                    foreach ($ref in $refs) {
                        $cleanValue = Get-CleanedValue -Value $ref.Value
                        $risk = Get-ReferenceRisk -Value $cleanValue
                        $replacement = ""
                        if ($risk -and $risk -ne "Modern-Site-Reference") {
                            $brokenCount++
                            $replacement = Get-SuggestedReplacement -Value $cleanValue
                            if ($replacement) { $stats.ReplacementsSet++ }
                        }
                        if (-not $risk) { $risk = "None" }
                        Write-LinkRow -DocumentUrl $docUrl -FutureDocumentUrl $futureUrl `
                            -Kind $ref.Kind -Confidence $ref.Confidence -RiskType $risk `
                            -EmbeddedUrl $ref.Value -CleanedUrl $cleanValue -SuggestedReplacement $replacement
                        $stats.LinksFlagged++
                    }
                } catch {
                    $scanned = "Error ($($_.Exception.Message))"
                }
            }

            Write-AccessRow -Row ([PSCustomObject]@{
                Site                      = $webUrl
                Library                   = $list.Title
                DocumentUrl               = $docUrl
                DocumentBrowserUrl        = $browserUrl
                FutureDocumentUrl         = $futureUrl
                FolderName                = $folder.Name
                FolderUrl                 = $folder.Url
                FileName                  = $fileLeaf
                Owner                     = $owner
                Created                   = $item.FieldValues["Created"]
                Modified                  = $item.FieldValues["Modified"]
                LastModifiedBy            = $editor
                PermissionMode            = $permMode
                ADGroupsWithAccess        = $perms.ADGroups
                SharePointGroupsWithAccess = $perms.SPGroups
                DirectUserAccess          = $perms.DirectUsers
                SharingLinks              = $perms.SharingLinks
                SharedWithUsers           = $sharedWith
                ScannedForLinks           = $scanned
                BrokenLinkCount           = $brokenCount
            })
            $stats.DocsRecorded++
        }
    }
}

Write-Host "`nDone."
Write-Host "Documents recorded: $($stats.DocsRecorded)  (folders skipped: $($stats.FoldersSkipped), libraries: $($stats.LibrariesRead))"
Write-Host "Documents with unique permissions: $($stats.UniquePermDocs)"
Write-Host "Files link-scanned: $($stats.FilesScanned)"
Write-Host "Link findings: $($stats.LinksFlagged)  (auto-replacement suggested for $($stats.ReplacementsSet))"
Write-Host "-> $AccessOutputPath"
Write-Host "-> $LinksOutputPath"

Disconnect-PnPOnline
