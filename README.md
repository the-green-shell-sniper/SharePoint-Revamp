# SharePoint-Revamp

This repository contains PowerShell scripts used for SharePoint migration analysis and link remediation planning.

## Scripts

- [Get-DocumentAccessReport.ps1](Scripts/Get-DocumentAccessReport.ps1)
- [Get-DocumentUrlLookup.ps1](Scripts/Get-DocumentUrlLookup.ps1)
- [Get-FixItListOverviewOnly.ps1](Scripts/Get-FixItListOverviewOnly.ps1)
- [Get-FixItListURLsv2.ps1](Scripts/Get-FixItListURLsv2.ps1)
- [Search-PhrasesInSpreadsheets.ps1](Scripts/Search-PhrasesInSpreadsheets.ps1) - searches every CSV/Excel file under a folder tree for each phrase in a phrases CSV (exact match, then fuzzy for phrases with no exact hit) and writes the file path + row of each match to a CSV.
- [Update-EmbeddedLinks.py](Scripts/Update-EmbeddedLinks.py) - offline, backup-first editing of authored Office/PDF hyperlink targets using the supplied migration lookups.
- [Publish-FixedEmbeddedDocuments.ps1](Scripts/Publish-FixedEmbeddedDocuments.ps1) - plans and uploads corrected files to the priority workbook's destinations, preserving existing target authors/editors and timestamps.
- [New-SPPermissionMatchMap.ps1](Scripts/New-SPPermissionMatchMap.ps1) - read-only proposals for renamed permission scopes; every proposal requires manual approval.
- [Sync-SPPermissionsFromSource.ps1](Scripts/Sync-SPPermissionsFromSource.ps1) - dry-run-first permission synchronization, with explicit source/target connections, snapshots, dependent-content protection and verified writes.
- [Invoke-CSPermissionSync.ps1](Scripts/Invoke-CSPermissionSync.ps1) - CS-only launcher with Match, DryRun, Apply and measured Verify steps; does not select the other checklist sites.
- [Invoke-SPPermissionChecklist.ps1](Scripts/Invoke-SPPermissionChecklist.ps1) - sequential processing of checklist pairs and their exact/approved matching descendant webs.

Deprecated scripts are archived in [Deprecated-Scripts/](Deprecated-Scripts/).

## Synchronizing migrated site permissions

### Remaining checklist sites: unattended additive pipeline

[Invoke-SPVisitorsReadChecklist.ps1](Scripts/Invoke-SPVisitorsReadChecklist.ps1)
runs **one checklist site at a time** through Match, checked DryRun, and (with
`-Apply`) verified Apply plus a complete recursive target report before moving to
the next checklist site. It does not run Scout or the full permission mirror.
It reuses the successful additive CS/PFS/Accounting engine.

```powershell
# Match and checked DryRun only; no permission or membership writes.
& 'C:\Users\!bposey\Desktop\GitHub\SharePoint-Revamp\CS-PermissionSuite\Invoke-SPVisitorsReadChecklist.ps1'

# Unattended remaining sites: Match -> checked DryRun -> verified Apply -> report.
# This command also runs Match and DryRun; the preview command is not a prerequisite.
& 'C:\Users\!bposey\Desktop\GitHub\SharePoint-Revamp\CS-PermissionSuite\Invoke-SPVisitorsReadChecklist.ps1' -Apply
```

- Only mappings in `ReferenceFiles\NewSiteChecklist.csv` are processed. CS,
  Accounting and PFS are skipped by default because their individual Applies
  completed. `-TargetSites` can narrow the checklist; unknown targets are rejected.
  To deliberately revisit a completed site, also pass `-CompletedTargetSites @()`.
- Each site's root and exact-path matching descendants are processed sequentially.
  Explicit checklist source/target mappings are reserved, so a parent does not
  process them a second time or guess a different pairing. Other site collections
  are never discovered via tenant enumeration. Unmatched/target-only child webs
  are reported rather than guessed or created.
- Uses each web's **associated Visitors group**, discovered live (not a group name
  guessed from the URL). A missing group or a protected migration principal stops
  that site. Individual PFS/Accounting runners retain their exact-title defaults.
- Fresh run-specific reviews auto-approve scores >= 0.80 under the existing policy.
  Lower scores and children of unapproved renamed parents remain manual. Exact
  paths are matched directly. Similarity scores are not identity probabilities.
- DryRun must produce complete successful per-web status and a consistent
  additive-only plan before Apply starts. Apply verifies its writes and final
  modified scopes, and its final report must finish successfully. Apply re-reads
  live state and replans; it does not replay the saved DryRun plan verbatim.
- Adds Visitors Read only where corresponding source Team Site Members has
  effective Read. EEEU membership is ensured only for webs with qualifying source
  scopes. Adding EEEU gives those users **all existing permissions held by that
  Visitors group**, not just this run's new Read grants.
- Existing grants and direct migration-group assignments/memberships remain
  unchanged. Unmatched/excluded inherited target objects may have inheritance
  broken with permissions copied to prevent new ancestor access spreading there.
  Access Requests is excluded everywhere; CS/PFS-specific exclusions apply only
  when those sites are deliberately selected again.
- A failure stops the current site's pipeline and continues the next checklist
  site. A failed write is never blindly retried or rolled back. Inspect live state
  and evidence before retrying a failed/partially applied site. Authentication may
  still require interaction if the persistent login cache is unavailable/expired.

Compact evidence is under `Reports\VisitorsRuns\v-*\<site-hash>`. The batch writes
`Summary.csv` and `Stragglers.csv` after each site, including failed stages/errors,
unapproved matches, unmatched source Read coverage and unmatched/target-only webs.
Each stage retains per-web evidence, source/target snapshots, plans, coverage and
write logs. Final summaries distinguish failure from verified matched additions;
successful Apply does not mean skipped objects were completed. Any failed sites
also cause a final terminating error **after** other checklist sites are attempted.
Keep the PowerShell window open and do not run overlapping permission jobs.

### Accounting individual additive runner

[Invoke-AccountingVisitorsRead.ps1](Scripts/Invoke-AccountingVisitorsRead.ps1)
uses the same sequential workflow as PFS, with explicit Accounting URLs and the
exact associated group `Accounting Visitors`. It processes the root followed by
exact-path subsites (the matrices contain `Private` on both sides). It adds Read
only where matching live source scopes have `Team Site Members` Read; unmatched
content is never guessed. No Accounting-specific business-library exclusion is
assumed; Access Requests remains excluded. EEEU membership is report-only on a web
with zero source Read scopes. Direct migration-group assignments/memberships remain
unchanged; normal Visitors container additions are authorized as for CS.

The existing Accounting scout is already in
`Reports\PermissionsReports\Scouts_20261019-131307-438\02-Accounting\20261019-140402-894`:
18 source lists, 30 target lists, no error records. The target User Information
List reports 71 items but both methods agree on 70 IDs. These are historical
observations, not unrestricted visibility guarantees. The batch scout writes its
summary after **each** site, not only at batch completion.

```powershell
# Existing scout can be reviewed first; Scout refreshes it if needed.
& 'C:\Users\!bposey\Desktop\GitHub\SharePoint-Revamp\CS-PermissionSuite\Invoke-AccountingVisitorsRead.ps1' -Mode Scout
& 'C:\Users\!bposey\Desktop\GitHub\SharePoint-Revamp\CS-PermissionSuite\Invoke-AccountingVisitorsRead.ps1' -Mode Match
& 'C:\Users\!bposey\Desktop\GitHub\SharePoint-Revamp\CS-PermissionSuite\Invoke-AccountingVisitorsRead.ps1' -Mode DryRun
# After inspecting coverage, plan and membership:
& 'C:\Users\!bposey\Desktop\GitHub\SharePoint-Revamp\CS-PermissionSuite\Invoke-AccountingVisitorsRead.ps1' -Mode Apply
& 'C:\Users\!bposey\Desktop\GitHub\SharePoint-Revamp\CS-PermissionSuite\Invoke-AccountingVisitorsRead.ps1' -Mode Report
```

Outputs: `Reports\Accounting`; stable approvals: `Reports\AccountingReview`.
The runner reuses `Invoke-PFSVisitorsRead.ps1`, so keep both and their existing
companions together. PFS and CS defaults are unchanged. Modern team-site subsite
to communication-site migration does not imply matching associated groups,
inheritance or role definitions: live validation remains necessary. No Microsoft
365 group membership API is used.

### PFS start-to-finish additive workflow

[Invoke-PFSVisitorsRead.ps1](Scripts/Invoke-PFSVisitorsRead.ps1) is scoped to
archived `/PFS` -> `/sites/PatientFinancialServices`, using source `Team Site Members`
and the exact associated target group `Patient Financial Services Visitors`.
It reuses the tested additive engine (CS defaults remain unchanged), processing
the root web/content first, then exact-path subsites sequentially in the same
PowerShell session. The supplied matrices include a `Coding` subsite on both sides.
Unmatched webs are listed in `Webs.csv`, never automatically paired by title.

The supplied old/new PFS matrices have no `Team Site Members` rows. Matrix evidence
is a baseline, not a substitute for live reads. **Zero live source Read scopes means
zero permission changes and no EEEU membership changes**, even with `-Mode Apply`.
Membership is still reported. There is no fallback to Everyone, a similarly named
group or a broad Visitors grant. The full mirror's unavailable-user/Review behavior
is not involved. Existing migration-group membership entries are not moved/removed;
the normal Visitors container may receive the explicitly authorized additions.

The target-only business library found in the matrices is `2021 E&M Coding Changes`,
site-relative URL path `2021 EM Coding Changes`. It is explicitly excluded.
Other target-only matrix paths are `_catalogs/appdata`, `_catalogs/appfiles`,
`_catalogs/hubsite`, `_catalogs/lt`, `_catalogs/solutions`, `_catalogs/theme`,
`_catalogs/wp`, `FormServerTemplates`, `IWConvertedForms`, `Lists/Events`,
`Lists/TaxonomyHiddenList`, and `Style Library`. These remain unmatched unless
an explicit counterpart/approval exists; do not infer intended migration identity
from similar titles. Access Requests is excluded. CS-specific COVID library
exclusions are not transferred to PFS.

After syncing scripts to the server, run each step and inspect its reports before
continuing:

```powershell
# Read-only schema and inventory comparison, root then subsites.
& 'C:\Users\!bposey\Desktop\GitHub\SharePoint-Revamp\CS-PermissionSuite\Invoke-PFSVisitorsRead.ps1' -Mode Scout
# Read-only matching. Scores >= 0.80 are approved in the separate PFS review files;
# lower scores remain pending. Nothing is applied by Match.
& 'C:\Users\!bposey\Desktop\GitHub\SharePoint-Revamp\CS-PermissionSuite\Invoke-PFSVisitorsRead.ps1' -Mode Match
# Read-only additive plan and EEEU membership observation.
& 'C:\Users\!bposey\Desktop\GitHub\SharePoint-Revamp\CS-PermissionSuite\Invoke-PFSVisitorsRead.ps1' -Mode DryRun
# Explicit production additions only if source Read coverage requires them.
& 'C:\Users\!bposey\Desktop\GitHub\SharePoint-Revamp\CS-PermissionSuite\Invoke-PFSVisitorsRead.ps1' -Mode Apply
# Standalone recursive target reference; no permission/membership changes.
& 'C:\Users\!bposey\Desktop\GitHub\SharePoint-Revamp\CS-PermissionSuite\Invoke-PFSVisitorsRead.ps1' -Mode Report
```

Evidence is under compact `Reports\PFS\p-*\<web-hash>` folders and stable approvals
under `Reports\PFSReview\<web-hash>`. Every mode records `Summary.csv` and `Webs.csv`.
Per-web additive evidence includes `Coverage.csv`, `Plan.json`, `Changes.csv`,
`Status.json`, and before snapshots. Inspect `SourceReadScopes` and
`PlannedOperations` for the expected zero-additions case. Apply performs the existing
per-write/final state checks and then generates one combined recursive target report.
Report mode can be used without Apply; reports include unmatched target-only content.
If a web fails, later webs stop, and earlier writes are not rolled back. Do not run
this alongside another apply process on the same site collection.

### Urgent CS additive-only Visitors Read pass

[Add-CSVisitorsRead.ps1](Scripts/Add-CSVisitorsRead.ps1) is separate from the full
mirror/queue. It only adds target associated Visitors `Read` at exact/approved
matched scopes where source `Team Site Members` has effective `Read`, including
inherited source access at unique target children. It never removes existing
Contribute/Edit/etc., resets inheritance, creates Review levels or attempts to
resolve obsolete source users. The current CS root has no discovered subsites;
this permission pass processes that web only, while its post-Apply report recurses.

```powershell
$repo = 'C:\Users\!bposey\Desktop\GitHub\SharePoint-Revamp'
# Read-only plan:
& "$repo\CS-PermissionSuite\Add-CSVisitorsRead.ps1"
# Explicit production additions, after inspecting the plan:
& "$repo\CS-PermissionSuite\Add-CSVisitorsRead.ps1" -Apply
```

Optionally supply `-MatchFile` with approved aliases; without it only exact paths
are paired, and unmatched source Read locations are explicit in `Coverage.csv`.
The two new COVID19 libraries and Access Requests are excluded. Unmatched/excluded
inherited targets can be frozen with copy to retain their grants before a parent
addition; no grants are removed. Four protected migration groups are not edited.
The associated Visitors group must have the exact title
`Clinical & Patient Services Visitors`; a different associated group stops the pass.
It must not itself be a migration principal. By explicit authorization, this named
site group may receive Read/EEEU even if it contains a protected migration group.
That indirect access is logged as `By design`; direct migration-group grants and
their membership entries remain unchanged. This exception is limited to this
additive CS pass; the full mirror retains its containing-group protection.

The target tenant's EEEU claim is constructed from its discovered tenant ID.
DryRun reports its Visitors membership without writes. Apply adds it only if
absent and rechecks membership. This intentionally makes **all existing Visitors
grants** available through EEEU, not just new Read assignments; review existing
Visitors access before applying. No other group membership is changed.

Evidence uses compact `Reports\PermRead\rd-*` folders: before snapshots, `Plan.json`,
`Coverage.csv`, `Changes.csv`, `Status.json` and a post-Apply recursive `Report`.
Verified additions are not a full permission mirror. A later report failure does
not roll back additions; inspect application status/log separately. Old-user
unavailability does not block this narrow pass because those users are not processed.
The full mirror continues to block unresolved users; this does not silently change
its policy or treat picker misses as proof of deletion.
Parameter-only REST POST operations send explicit empty JSON (`{}`), rather than
leaving the PnP request body implicit. Failed writes are never automatically retried.
After a failed Apply, inspect `Changes.csv` for verified earlier operations and
unknown outcomes before rerunning; a new run rereads live state and plans only
missing additions.

### Queued automatic application and recursive reporting

#### Review level and intentionally preserved migration groups

When source `Review` and target `Review` have different masks, the sync explicitly
maps source grants to a dedicated target level named
`Migrated Review H<High> L<Low>` (for CS: `Migrated Review H176 L138629217`).
DryRun plans this level without creating it, using a placeholder role ID that must
never be used for writes. Apply creates the definition if needed and reads it back
to require an exact full-mask match before assigning it. Existing target `Review`
is never modified. A dedicated definition can remain if later preflight fails;
creation is not a role assignment and is not rolled back. Other mismatched levels
still stop the run. Existing dedicated levels with wrong masks are not overwritten.

These four principals are intentionally preserved at every scope:

- `SharePoint.Migration.Team.Admins`
- `SharePoint.Migration.Team.Members`
- `SharePoint.Migration.Team.Visitors`
- `SharePoint.Migration.Team.MSLs`

No grants are added, removed or remapped for them. The scripts discover their
direct membership in SharePoint groups; **containing SharePoint groups are also
preserved**, because changing a container's role would change the migration group's
access (and may preserve other members' access too). Inheritance resets that would
change protected role bindings are blocked. This is an intentional exception to
the source mirror and can stop automatic preflight. Reports retain these grants
with `MigrationPolicy = Intentionally preserved...` and `MigrationMembers`.
`MigrationGroupMemberships.csv` reports direct SharePoint-group membership, with
explicit not-observed rows where none was returned; it does not expand nested
Entra groups or prove a principal has no other access. Sync logs label preserved
assignments `Intentionally skipped`; per-web membership evidence is also saved.
Each site report also includes `PermissionLevels.csv` with the per-web role
definitions and raw masks; effective rows resolve masks from that definition list.
Protected domain-group identity is recognized by exact title/name or the name at
the end of a login claim; discovered membership also resolves its opaque login claim.
No directory group membership is changed.

An application/preflight error is retained even if its follow-up permission report
also fails. Both errors are reported rather than replacing the original blocker
with a report failure. Inventory disagreement remains fatal; do not bypass it.
An ID-set disagreement triggers at most three fresh whole-web inventory attempts,
discarding each failed inventory. Persistent disagreement still fails closed;
successful retry is warned and recorded as inventory review evidence.

[Invoke-SPPermissionQueue.ps1](Scripts/Invoke-SPPermissionQueue.ps1) previews the
checklist queue by default. `-Apply` explicitly authorizes automated changes:

```powershell
$repo = 'C:\Users\!bposey\Desktop\GitHub\SharePoint-Revamp'

# Preview only (no connections, windows or writes).
& "$repo\Scripts\Invoke-SPPermissionQueue.ps1" `
    -TargetSites 'https://whallc.sharepoint.com/sites/ClinicalandPatientServices'

# Authorizes changes for CS only. Omit TargetSites to select the entire checklist.
& "$repo\Scripts\Invoke-SPPermissionQueue.ps1" -Apply -MaxConcurrent 4 `
    -TargetSites 'https://whallc.sharepoint.com/sites/ClinicalandPatientServices'
```

There are at most four active site workers per queue (configurable downward).
Each uses its own elevated PowerShell **7** window, retained with `-NoExit`.
The site root and its content finish Match, DryRun, Apply and measured verification
before the first matched subsite starts; subsites run sequentially, depth-first,
in the same window. Completed windows release the active slot without closing.
Related **target site collections** are serialized, even across checklist rows;
shared read-only source archive collections do not serialize otherwise independent
targets. There is no per-list/item parallelism.
Worker-held named mutexes also cap actual work at four and serialize related
collections if the scheduler is restarted while earlier worker windows still run.

Auto mode approves exact paths through the existing matcher and fuzzy proposals
with **Score >= 0.80**. This is a similarity score, **not an 80% identity guarantee**.
It intentionally does not require the matcher's High tier or runner-up margin.
Ambiguous high-scoring proposals can therefore be approved; inspect score evidence.
Renamed children cannot bypass unapproved containers. Existing approvals are kept;
previously unapproved high-score rows are automatically approved, with an
`ApprovalPolicy` column documenting this decision. Existing review files are retained,
not replaced by new proposals. Review/merge fresh candidates when content changes.
Lower scores remain `Approved=No` for manual review and are not synchronized. After
approving those rows, rerun the site explicitly. A completed worker with pending
review rows is **not** a claim that everything was synchronized.

The CS `COVID19 Resources` and `COVID19 Workflows` libraries are explicitly excluded
from synchronization by the checklist runner, regardless of score, but still appear
in the permission report. As with other unmatched content, inherited access may be
protected by breaking inheritance **with copy**, retaining grants before ancestor
changes. No source is written. Source `Team Site Members` still maps to target
Visitors; root target Owners/Members remain protected.

Each site's Auto workflow refuses Apply if DryRun has blocked scopes or unavailable
principals, rechecks the live Apply plan, and performs a fresh post-Apply DryRun
requiring zero remaining actions. Failures stop that site, not the remaining queue.
Earlier writes are not rolled back. Apply can ensure target site users before its
final permission-plan gate; these directory entries are not role grants.

Default outputs are under `Reports\PermRuns\q-*`. Queue site folders use a
12-character URL hash, per-web folders use `w-<hash>`, and combined reports use
`r-<hash>`. Auto no longer nests a second timestamped checklist folder.
Checklist summaries/status records retain the full URLs. Review locations and
existing approval files are unchanged.

- `QueueStatus.csv`: queued/running/completed/failed outcomes.
- Per-worker `WorkerStatus.json` and `WorkerTranscript.txt`: retained failure details.
- Per-site `r-*` folder (one combined report bundle for root and descendants):
  - `Changes.csv`: web, path, action, principal/role detail, result and evidence log.
    **Verified** means a write completed and was checked; **Planned** is not applied.
    Write-start entries without a verified outcome require live inspection.
  - `EffectivePermissions.csv`: every enumerated target web, list/library, folder,
    file and item, including hidden lists, unmatched and excluded content; one row per
    principal/role binding with inheritance origin, observation time and inventory caveats.
    Target-only subsites are reported even though they are not permission-synced.
  - `ReportStatus.json`: Complete or Failed, object/web/row counts and limitations.
    A partial CSV with a Failed manifest is **not** a complete reference.
- Per-web snapshots, plans, inventory matches and logs remain available as evidence.

The recursive report is generated after application (also attempted after failure);
`Changes.csv` is saved before scanning so a reporting failure cannot hide earlier
changes. The report streams CSV rows, retaining one web inventory/ACL set at a time.
This bounds memory by web size, not by a fixed byte limit. It is not a transactional
snapshot and can reflect concurrent edits. Enumeration is limited to the signed-in
account; agreement between methods does not prove unrestricted visibility.
Role assignments include observed limited-access and sharing-link bindings but do
not expand SharePoint/Entra groups, link recipients, site admins, app permissions or
tenant policies into individual effective access. Treat these reports as permission
references, not an exhaustive authorization simulation.

Use [Export-SPPermissionSiteReport.ps1](Scripts/Export-SPPermissionSiteReport.ps1)
standalone to refresh a read-only recursive reference. The checklist `Apply` mode
also produces reports without enabling Auto. The CS launcher's existing
Match/DryRun/Apply/Verify workflow stays manual.

Run the queue in an elevated PowerShell 7 console using the same Windows account
as the cached PnP login. UAC/browser authentication may still be needed. Do not launch
another queue or manual apply/scout on overlapping sites while workers remain active.
Keep the scheduler open until workers publish their final status. Closing it does
not terminate elevated workers; inspect retained windows before restarting.
Four workers reduce pressure but do not guarantee Microsoft will not throttle.
The existing bounded read retries remain in effect; ambiguous writes are not retried.
Reports contain sensitive identities/paths; restrict access to the report repository.

#### Shortening existing completed reports

The new output layout is tested below 240 characters at both the server and
development repository locations. Custom longer output roots can still exceed
Windows limits. Old reports are not automatically renamed while workers run.
Use the preview-first relocation script on a **completed** run, never a review
folder or active run:

```powershell
$repo = 'C:\Users\!bposey\Desktop\GitHub\SharePoint-Revamp'
$oldRun = "$repo\Reports\PermissionsReports\PermissionQueue\Queue-20261019-135949-254"
& "$repo\Scripts\Move-SPPermissionReportsToShortPaths.ps1" -SourceFolder $oldRun
# After inspecting the preview and confirming all workers have finished:
& "$repo\Scripts\Move-SPPermissionReportsToShortPaths.ps1" -SourceFolder $oldRun -Apply
```

This relocates files into `Reports\PermArchive\<run-hash>\<file-hash>.<extension>`.
`Paths.csv` preserves original absolute/relative names and SHA-256 hashes before
any move; every move is checked and existing destinations are never overwritten.
Reusing the same arguments resumes a partial move. Embedded paths are not rewritten:
use the manifest to locate their archived evidence. Empty old directories are left
alone. Include both relocated files and old-path deletions in the next commit.
No Git long-path setting is required for the compact layout.

Offline validation:

```powershell
.\Scripts\Test-SPPermissionScripts.ps1
.\Scripts\Test-SPPermissionQueue.ps1
```

### CS only: do this first

If inventory counts/fields fail on the server, run the read-only
[Get-CSPermissionScout.ps1](Scripts/Get-CSPermissionScout.ps1) beside
[SPPermission.Common.ps1](Scripts/SPPermission.Common.ps1) before retrying Apply:

```powershell
.\Get-CSPermissionScout.ps1
```

This uses the existing document-lookup scanner's recursive paged PnP/CAML pattern,
discovers fields, reads item inheritance explicitly, and compares item IDs against
REST on both CS webs. It includes hidden/empty lists but does not recurse into
subsites or change permissions. Timestamped `Reports\PermissionsReports\CS-PermissionScout` reports include
`Lists.csv`, `Fields.csv`, `Items.csv`, and `Errors.csv`. Count discrepancies are
flagged for review, not explained away or used to authorize permission changes.
Read failures are saved and cause an explicit failed result. The sync's current
inventory now cross-checks live recursive PnP item IDs when REST differs from the
reported count. ID disagreement or a failed verification read stops the run.
If the initial REST item query specifically fails with `Column 'ContentTypeId'
does not exist`, that list uses the scout's recursive PnP field-value enumeration
and explicit inheritance reads instead. This fallback is warned and recorded in
snapshots and sync log Review rows, not silently skipped. It does not handle
authorization errors or resume a partially read REST list. Without a size field,
file content comparison is unavailable for that list. REST errors include the
failed URL to identify the exact list/query.
Offline tests: `.\Scripts\Test-CSPermissionScout.ps1`.

#### Scouting the remaining checklist sites

[Invoke-SPPermissionScouts.ps1](Scripts/Invoke-SPPermissionScouts.ps1) reuses that
same scout **sequentially**, completing the source and target of each checklist row
before starting the next. By default it excludes CS, which can continue independently.
Copy the runner, the updated [Get-CSPermissionScout.ps1](Scripts/Get-CSPermissionScout.ps1),
[SPPermission.Common.ps1](Scripts/SPPermission.Common.ps1), and
[NewSiteChecklist.csv](ReferenceFiles/NewSiteChecklist.csv) into the same server folder.
No permission-sync/apply scripts are called.

```powershell
# All remaining checklist pairs, one pair at a time, using persistent login.
.\Invoke-SPPermissionScouts.ps1 -ChecklistFile .\NewSiteChecklist.csv

# Or run just one pair first / retry a failed pair.
.\Invoke-SPPermissionScouts.ps1 -ChecklistFile .\NewSiteChecklist.csv `
    -TargetSites 'https://whallc.sharepoint.com/sites/Accounting'
```

Reports default to `Reports\PermissionsReports\Scouts_<timestamp>` under the repo root.
The server checkout is `C:\Users\!bposey\Desktop\GitHub\SharePoint-Revamp`, so the
shared report root there is
`C:\Users\!bposey\Desktop\GitHub\SharePoint-Revamp\Reports\PermissionsReports`.
Both the `Scripts` CS scout and the copy in `CS-PermissionSuite` default to
`Reports\PermissionsReports\CS-PermissionScout` under that same repo root.
Paths are relative to the scripts, not hard-coded to a Windows account.
Explicit relative output/review arguments are resolved against PowerShell's current
filesystem location before .NET directory creation; they do not use the process's
potentially different working directory (such as `C:\Windows\System32`).
Explicit `-OutputFolder` arguments still override these defaults; already completed
or running scouts retain their original output paths. Existing reports are not moved.
Each pair has its own subfolder with the existing four scout reports.
`ScoutSummary.csv` is saved after every pair and identifies failures, count/ID
discrepancies needing review, and each report location. Read failures do not stop
the other read-only scouts; the batch throws at the end if any pair failed.
`-IncludeCS` includes CS in a full run. An explicit `-TargetSites` selection can
also select CS. Unlisted targets and duplicate mappings are rejected before reads.
Only the explicit checklist webs are scouted: listed subsites run as separate
rows, without automatic recursion into unlisted subsites or tenant enumeration.
No content or permissions are changed. Keep all output private.

You do **not** need to process every site to fix Clinical & Patient Services.
Copy these five scripts together to `P:\IT\Private\zzBPosey\whaitapp09`:

- [Invoke-CSPermissionSync.ps1](Scripts/Invoke-CSPermissionSync.ps1)
- [Invoke-SPPermissionChecklist.ps1](Scripts/Invoke-SPPermissionChecklist.ps1)
- [New-SPPermissionMatchMap.ps1](Scripts/New-SPPermissionMatchMap.ps1)
- [Sync-SPPermissionsFromSource.ps1](Scripts/Sync-SPPermissionsFromSource.ps1)
- [SPPermission.Common.ps1](Scripts/SPPermission.Common.ps1)

The CS launcher pins archived CS to Clinical & Patient Services; it never loads the
full checklist. It uses persistent login by default, with no tenant prompt.
It includes matching descendant webs, processing one web's inventory, plan, writes
and reports before the next. It does not discover unrelated tenant sites.

```powershell
Set-Location 'P:\IT\Private\zzBPosey\whaitapp09'

# 1. Read-only inventories and matching proposals.
.\Invoke-CSPermissionSync.ps1 -Mode Match

# 2. Review MatchReview.csv and WebReview.csv under CS-PermissionReviews.
#    Approve confirmed fuzzy matches only. Empty/header-only CSVs need no edits.
#    If you approve a renamed subsite, run Match again to scan that subsite's content.

# 3. Read-only plan; inspect logs, removals, resets, freezes and blocked scopes.
.\Invoke-CSPermissionSync.ps1 -Mode DryRun

# 4. Only after reviewing the dry-run plan: changes CS and matched descendants.
.\Invoke-CSPermissionSync.ps1 -Mode Apply

# 5. Read-only verification. Throws if actions, blocked scopes or unavailable
#    principals remain, or if a selected web could not be read.
.\Invoke-CSPermissionSync.ps1 -Mode Verify
```

Reports are saved beneath `CS-PermissionRuns`; stable approval files live beneath
`CS-PermissionReviews`. Both default to folders beside the launcher, so no path edits
are needed. Match never overwrites existing approval files; when refreshing,
explicitly merge/replace them using the fresh timestamped candidates in the reports.
Verify covers matched scopes only: unmatched/unapproved objects remain skipped.
EEEU membership and testing-account access still need separate confirmation.

### Shared requirements and single-pair commands

These scripts run against SharePoint **on your execution server**. Local regression
tests use mocks only; they do not establish that your tenant, application consent,
account permissions or server authentication cache are configured correctly.
Use PowerShell **7.4.6+** and `PnP.PowerShell` **3.3.0+**. The signed-in user/app
must be able to enumerate all scopes and read source permissions, and must have
permission to manage target permissions. A trimmed/inaccessible inventory is not
sufficient. No elevated role is assigned automatically.

Copy these three files together into the same server folder:

- [New-SPPermissionMatchMap.ps1](Scripts/New-SPPermissionMatchMap.ps1)
- [Sync-SPPermissionsFromSource.ps1](Scripts/Sync-SPPermissionsFromSource.ps1)
- [SPPermission.Common.ps1](Scripts/SPPermission.Common.ps1) (required shared helper)

The defaults select archived CS at
`https://whallc.sharepoint.com/sites/archive-2026-10-18T115436Z/CS` as source and
`https://whallc.sharepoint.com/sites/ClinicalandPatientServices` as target, with the
application ClientId from the original scripts. Override `-SourceUrl`, `-TargetUrl`
and `-ClientId` for other migrations. Use `-ClientId ''` to defer to existing PnP
application configuration. **A tenant ID is not an application ClientId.**
`-Tenant` optionally pins the tenant by GUID or verified domain.
Both scripts use `-Interactive -PersistLogin -ReturnConnection`; sign-in can still
be required. `-SkipPersistLoginSetup` omits cache enablement for the settings-lock
problem described in the document-upload section; it does not clear cached tokens.
REST JSON is parsed as dictionaries to preserve SharePoint's case-distinct `Id`
and `ID` fields, including nested item responses.
Library file sizes are read from expanded `File/Length`, not the optional
`File_x0020_Size` list field. Folders and ordinary list items have no file size.
Before enumerating items, each list's field schema determines whether Title,
Created and single-user Author metadata can be requested. Absent optional fields
are reported as warnings and excluded from matching signals, without skipping
the list's contents or permission scopes. Field-schema read failures still stop.

Example server workflow, from the folder containing all three scripts:

```powershell
$options = @{
    OutputFolder = (Join-Path $PWD 'PermissionRuns')
}

# Read-only: complete inventories, then rename/content-match proposals.
.\New-SPPermissionMatchMap.ps1 @options

# Open the generated PermMatch_Review CSV. ALL rows start Approved=No.
# Approve only confirmed pairs. Approve parent renames and child rows explicitly.
$review = Read-Host 'Full path to the reviewed PermMatch_Review CSV'

# Live read-only dry run. No EnsureUser or permission writes.
.\Sync-SPPermissionsFromSource.ps1 @options -MatchFile $review

# Review the CSV log and JSON plan, especially Remove, Reset, Freeze and Blocked.
# Only after review, explicitly apply; this re-reads and rebuilds the live plan.
.\Sync-SPPermissionsFromSource.ps1 @options -MatchFile $review -Apply

# Confirm no remaining actions, blocked scopes or unavailable-principal warnings.
.\Sync-SPPermissionsFromSource.ps1 @options -MatchFile $review
```

Omit `-MatchFile` for exact site-relative paths only. The apply script accepts only
approved one-to-one rows for the selected webs and rejects missing paths, conflicting
exact matches, incompatible kinds/extensions, and malformed approvals.
Different list templates (`BaseTemplate`) do not block otherwise matching lists or
libraries in any site, including subsites. Exact-path and approved matches emit a
warning with both template numbers and web URLs, but continue through Match,
DryRun and Apply. Renamed-list/library proposals may also span different templates;
their review CSV evidence includes both template numbers and still requires explicit
approval. List-to-library matches remain rejected. This is permission matching only:
the scripts do not convert list templates, schemas or content.
Approval does not expand a parent alias implicitly: each renamed child needs its
own approved row. Inherited objects must retain compatible mapped parents; a move
to an unrelated/unapproved parent is logged and skipped, with its existing access
protected before ancestor changes, rather than silently changing effective access.
Exact paths are assumed to identify the same migrated object, including list items;
the scripts do not prove content identity for exact matches. Verify migration
identity before using them on sites where paths or list-item IDs were reused.

### Permission policy and safety limits

For later multi-site work, the checklist runner reads only `Old Site URL` and
`New Site URL` from [NewSiteChecklist.csv](ReferenceFiles/NewSiteChecklist.csv).
Explicit rows take precedence over recursive child-path matching, including relocated
subsites such as IT-Internal. `-TargetSites` can select exact checklist target URLs.
Exact child paths recurse automatically; renamed child webs require approved rows
in their parent's `WebReview.csv`. Unmatched webs and their contents are skipped,
not created or searched across unrelated trees. Existing inherited target subwebs
can receive a protective inheritance break before their parent's access changes;
their effective access is preserved, then matched webs are processed separately.
An inherited source web's effective permissions are materialized uniquely at its
mapped target root; cross-site migrations cannot recreate source parent inheritance.

- `Team Site Members` maps to the target's **associated Visitors** group. Other
  SharePoint groups must exist by title; users/Entra groups use their login claims.
  Group memberships, site collection administrators and app permissions are not
  synchronized. Adding EEEU to Visitors is a separate, intentional operation.
- Inheritance follows the source. If the **source root inherits**, its effective
  ancestor ACL is copied onto a unique target root; the target is not made to
  inherit from an unrelated site. Target associated Owners/Members grants are
  preserved **only at the root**.
- Limited Access and sharing-link assignments are not edited. Existing sharing
  links, group membership differences and protected root grants mean this is not
  a guarantee of identical effective access for every person.
- Unmatched and excluded target content is not directly synchronized. If it inherits
  a changing scope, the script freezes its current grants first. `-NoFreeze`
  deliberately lets that content follow the changed parent instead. Excluded paths
  can therefore still receive a protective inheritance break; they are not absolute
  no-write scopes. The default exclusion is `Access Requests`.
- A missing principal is reported once. Available additions may proceed, but
  removals and inheritance resets depending on missing grants are **blocked**, not
  treated as success. Dry-run picker misses are predictions; Apply uses EnsureUser
  and may legitimately produce a different plan. Apply may add target Site Users
  entries during preflight. Authorization, service and unknown resolution errors stop.
- Complete inventory/ACL reads and matching permission-level **names and bitmasks**
  are required. A reported count mismatch triggers a fresh recursive PnP ID read.
  Only exact agreement with the REST ID set allows continued planning; disagreements
  and failed reads stop. Count discrepancies remain warnings and sync log Review rows.
  Both enumerators can be security-trimmed: agreement is not proof of unrestricted
  visibility. Run as an account authorized to see/manage every selected scope.
  Permission levels are not created or silently substituted. Immediate subwebs are
  recorded as permission boundaries; the parent pair does not synchronize their
  roots/content directly. The CS/checklist runner processes matched webs separately.
- Adds precede removals on each scope. Each apply step checks its expected ACL and
  inheritance before writing and verifies them afterward; final modified states are
  checked again. Failed freezes, failed additions and ambiguous write responses stop
  further writes. Writes are not automatically retried or rolled back.
- Use a quiet maintenance window. Snapshots/plans are point-in-time reads, not locks;
  concurrent content additions, group-membership changes and permission changes are
  not transactionally isolated. Earlier successful writes remain after an error.
  Review the log and live state before retrying.

Each run saves source/target JSON snapshots (all enumerated inheritance states and
unique ACLs, including system-managed bindings), a complete inventory/match CSV
before permission changes, a JSON permission plan, a CSV log,
and an unavailable-principals CSV with headers even when empty. These are **manual
rollback references**, not an executable restore operation. Store them securely:
they contain identities, paths and access information. Files use timestamped names
under `Outputs\PermissionRuns` by default; specify a server-local `-OutputFolder`
when copying the scripts elsewhere. Do not commit generated permission reports.

The matcher compares only objects in matched list/library pairs; cross-library
moves are not searched. Signals include names, created/author, size, identical-byte
SHA-256 and word-set similarity for Office/text files. Candidates are cheaply
filtered before downloading; PDFs/images have byte comparison only. Downloads are
in-memory and limited by `-MaxContentMB` (25 by default); expanded Office text is
limited to 50 MB. `-NoContent` disables downloads. Comparison errors are warnings
and evidence, not proof of a match. Scores are heuristics, never approval.

The supplied [old CS](SupportFiles/PermissionsMatrixReports/oldCS_Permissions%20matrix%20report.xlsx)
and [new CS](SupportFiles/PermissionsMatrixReports/newCS_Permissions%20matrix%20report.xlsx)
reports informed these defaults: the source snapshot contains 108 custom scopes
(including the web), 91 explicit Team Site Members Read assignments, and the target
has 81 explicit extra Contribute assignments for CentralPark.CS.Designer. These
are historical report counts, **not hard-coded expected live changes**. New target
libraries such as COVID19 Resources/Workflows are not inferred to be old CS content.
The scripts always inventory the selected live webs when run on your server.

Run the mock-only regression suite locally or on the server, without PnP installed
or a SharePoint connection:

```powershell
.\Scripts\Test-SPPermissionScripts.ps1
```

## Uploading corrected documents

[Publish-FixedEmbeddedDocuments.ps1](Scripts/Publish-FixedEmbeddedDocuments.ps1) selects
**only documents with an applied `updated` record** in the remediation-run manifests,
matches their exact filenames to [EmbeddedLinksPriorityList-Copy.xlsx](EmbeddedLinksPriorityList-Copy.xlsx),
and checks current local SHA-256 hashes against the latest remediation snapshots.
It reads XLSX directly; Excel and ImportExcel are not required.
No filename fuzzy matching is used for upload destinations. Duplicate/missing workbook
rows or targets, changed local files, and unsafe URLs stop planning.

Default execution is **offline plan only**:

```powershell
# On the repository workstation: no SharePoint connection or upload.
.\Scripts\Publish-FixedEmbeddedDocuments.ps1 `
  -OutputFolder .\UploadRuns\server-plan-unique-name
```

Copy the generated `upload-plan.csv` and the script to your server working folder,
for example `P:\IT\Private\zzBPosey\whaitapp09`. The plan contains relative document
paths, hashes and exact remote targets. Keep the corrected document subfolders
intact under `P:\IT\Private\zzBPosey\whaitapp09\UpdatedEmbeddedLinkCopies`.
With `-PlanCsv`, you do **not** need to copy the workbook, Python/PDF tooling, or
the remediation history/backups to the server. Edits made after planning require
a fresh verified plan; do not change the hashes to bypass validation.

Use **PowerShell 7.4.6+**, not Windows PowerShell 5.1, and `PnP.PowerShell` 3.3.0+.
If the module is missing, install it in your server user profile:

```powershell
Install-Module PnP.PowerShell -Scope CurrentUser
```

Run the following from PowerShell 7 on the server, under the same Windows user
as your working PnP scripts. Like `Get-DocumentUrlLookup.ps1`, the uploader uses
`-Interactive -PersistLogin` and leaves `ClientId` optional so PnP can use its
existing application configuration and persisted sign-in. If needed, supply
`-ClientId` with your actual Entra application client ID GUID, not a placeholder.
A cached tenant ID or Microsoft Graph session alone is not a PnP application
configuration or PnP token cache; sign-in may still be required.
If PnP fails in `Settings.Save()` / `EnableCaching()` because its local
`.m365pnppowershell/settings.json` is locked, use `-SkipPersistLoginSetup`.
This omits `-PersistLogin`, avoiding that cache-enablement settings write without
deleting settings or clearing authentication. Existing cache entries can still
be used by PnP; this is not a general cache-disable switch and does not bypass
all possible settings-file reads/locks. Authentication failures still stop the run.
The signed-in user/app
must have permission to update document content and the system metadata fields.

```powershell
$work = 'P:\IT\Private\zzBPosey\whaitapp09'
$options = @{
    DocumentsFolder = "$work\UpdatedEmbeddedLinkCopies"
    PlanCsv = "$work\upload-plan.csv"
}

# Live read-only SharePoint preflight; backs up remote bytes/metadata locally.
& "$work\Publish-FixedEmbeddedDocuments.ps1" @options `
  -OutputFolder "$work\UploadRuns\preflight-unique-name" -Preflight

# Only after reviewing the plan/preflight and confirming replacement is intended:
# use a NEW folder. Apply repeats complete live preflight before writing anything.
& "$work\Publish-FixedEmbeddedDocuments.ps1" @options `
  -OutputFolder "$work\UploadRuns\apply-unique-name" -Apply
```

### Metadata preservation and safety

The script captures the **existing destination file's** `Author` (Created By),
`Editor` (Modified By), `Created`, and `Modified` immediately before upload.
It overwrites the existing file, then assigns typed user lookup values and explicitly
UTC DateTimes directly through CSOM `ListItem.UpdateOverwriteVersion()` to restore
those fields and verifies all four, plus file/item
identity and the uploaded byte hash (or the explicitly opted-in Office package
comparison described below). User IDs are reused within the same site;
no user-account creation or cross-site identity guessing is needed.
This preserves the current target metadata, **not old-site metadata from before migration**,
and does not change the document's internal Office author properties.

Every target must already exist. The script never creates folders or files at
missing destinations. Complete preflight for all files must pass before any upload.
Each backup must have matching identity, ETag, author/editor and timestamps before
and after downloading. A changing ETag or metadata causes the entire read to be
discarded and retried, at most three times with two-second pauses. Warnings show
the differing values; persistent changes fail preflight. Identity changes and
locks fail immediately. These retries perform no remote writes and do not retry uploads.
Checked-out/locked targets, required check-out, content approval, minor/draft versions,
and retention-labelled files are refused rather than bypassing their workflows.
Subfolder URLs and library `Forms/AllItems.aspx` URLs are resolved from the workbook;
`id`/`RootFolder` folder selectors must remain inside the specified library.

Immediately before each upload, remote identity, metadata, ETag and content hash
are rechecked. The CSOM binary write supplies the current ETag to reject concurrent
changes and disables automatic query retries. Exact-byte identical content is skipped.
Uploads are not a tenant-wide transaction: an error stops further files, but earlier
verified uploads remain. No automatic content rollback is attempted. A failed or
ambiguous upload may have changed remote bytes; its result is explicitly marked
for review. Metadata is not written after an upload exception, because the error
could represent another person's concurrent edit. Inspect the backups and remote
state before retrying. After an acknowledged upload, metadata restoration also
checks the uploaded bytes, identity and post-upload ETag before writing.

Each run saves `upload-plan.csv`; live runs also save `upload-results.csv`,
`remote-backups/` with original SharePoint bytes, and adjacent `.metadata.json`
snapshots containing user IDs, timestamps, identity, ETag and original SHA-256.
These are backups of the remote files **before uploading**, separate from the
earlier local hyperlink-editing backups. Retain them until validation is complete.

**Important:** preserving the current item fields does not erase audit records,
retain every historical version's attribution, or prevent a new content version.
SharePoint audit/version history can still identify the uploading user. Uploads
and `UpdateOverwriteVersion` can trigger automation/Power Automate flows; a flow
may change metadata after verification. Coordinate with the site owner before
applying, and spot-check the destination fields afterward. Insufficient permissions,
retention policies, and tenant automation can prevent preservation; failures are
reported, never silently treated as success.

### Recovering metadata after an interrupted upload

Do not blindly rerun `-Apply` after a metadata verification failure: a fresh run
would capture the already-changed destination metadata. Keep the original apply
run's backups. `-RestoreMetadataSnapshot` is an explicit **metadata-only write**
mode using the original adjacent `.metadata.json` and backup. It verifies the
backup hash, plan destination, remote file/item identity, writability, current
ETag and corrected content hash before restoring, then verifies metadata and
content. It never uploads document bytes. It cannot be combined with `-Preflight`
or `-Apply`. Use a new output directory; success writes `metadata-recovery.json`.

```powershell
& "$work\Publish-FixedEmbeddedDocuments.ps1" @options `
  -RestoreMetadataSnapshot '<full-path-to-original-backup.metadata.json>' `
  -OutputFolder "$work\UploadRuns\metadata-recovery-unique-name"
```

The initial verifier incorrectly treated timezone-less CSOM field dates as UTC;
site-local representations could therefore report a false timestamp mismatch.
The current script reads Author/Editor IDs and explicit UTC timestamps through
SharePoint REST for both capture and verification, and sends UTC DateTime values
directly through CSOM for writes. Recovery always uses the saved baseline rather
than guessing an offset or adding hours. After recovery succeeds, a
fresh `-Apply` run can safely skip already-identical content and process the rest.

### Office metadata rewriting during upload

SharePoint can rewrite Office package properties when a file enters its destination
library. A verified example changed the library content-type schema, custom document
properties, metadata datastore ID, and added `[trash]` entries while leaving every
Word content, relationship and media part byte-identical. Whole-file SHA-256 therefore
changed even though the authored content and corrected hyperlinks did not.

Exact-byte verification remains the default. Explicit `-AllowOfficeMetadataChanges`
enables a restricted DOCX/XLSX/PPTX package comparison during upload verification
and metadata recovery:

- All package part names and uncompressed bytes must match, except recognized
  `docProps/core.xml`, `docProps/custom.xml`, and custom XML rooted in the Office
  document-property or content-type-schema namespaces.
- Datastore properties linked to those metadata parts may change their item ID.
  SharePoint may also remove the complete schema-reference list from such a datastore
  when rebuilding library metadata (observed on C-17). Only a list consisting entirely
  of empty `schemaRef` elements with a single schema-URI attribute may be removed.
  Added or replaced references, unrelated attributes/children, and changes in
  non-metadata datastores are rejected. Relationships themselves must match exactly.
- Added `[trash]/<digits>.dat` entries are permitted. Other additions, deletions,
  content/hyperlink changes, or arbitrary custom XML changes fail verification.
- Digitally signed Office packages and all other formats, including PDFs, still
  require exact bytes. No metadata allowance applies to them.
- Every accepted comparison appends local/remote hashes and differing part names to
  `office-verification.jsonl`. Success after such a comparison is reported as
  `UploadedAndVerifiedOfficeContent`, not as a whole-file hash match.

Add the switch to both metadata recovery and subsequent `-Apply` commands if this
behavior is approved. The script does not suppress content mismatches or loosen
lock/identity/concurrency checks. Metadata-only differences can be applied again on
a subsequent upload; only exact-byte matches are automatically skipped. Keep the
original failed apply snapshot for recovery before starting a fresh apply run.

Offline tests (no SharePoint connection):

```powershell
.\Scripts\Test-PublishFixedEmbeddedDocuments.ps1
```

## Offline embedded hyperlink updates

[Update-EmbeddedLinks.py](Scripts/Update-EmbeddedLinks.py) recursively processes a local
document folder without connecting to SharePoint. It changes only hyperlink targets
on `whallc.sharepoint.com` (including bare-host URLs), not visible labels, visible URL
text, external URLs, images, or document relationships such as templates.
Legacy `http://centralpark/...` links are outside this hostname scope and remain unchanged.

Requirements: Python 3.10+ and the pinned `pypdf` dependency below. Acrobat, OCR,
Java/Tika, and PDF rendering utilities are not needed to edit existing PDF links.
Optional legacy `.doc` support uses installed desktop Word via PowerShell 7 and
[Edit-LegacyWordLinks.ps1](Scripts/Edit-LegacyWordLinks.ps1). Run that option in a
Windows desktop user session, not an unattended server/service account. Word is
not needed for DOCX/XLSX/PPTX editing.

From the repository root in PowerShell:

```powershell
python -m pip install -r .\Scripts\embedded-link-requirements.txt

# Preview: reads documents, produces audit reports, changes no documents.
python .\Scripts\Update-EmbeddedLinks.py `
  --input .\UpdatedEmbeddedLinkCopies `
  --run-dir .\LinkRemediationRuns\preview-unique-name `
  --legacy-word

# Apply in place. Use a NEW run directory on every invocation.
python .\Scripts\Update-EmbeddedLinks.py `
  --input .\UpdatedEmbeddedLinkCopies `
  --run-dir .\LinkRemediationRuns\apply-unique-name `
  --legacy-word --apply

python -m unittest discover -s .\Scripts -p test_update_embedded_links.py -v
```

Omit `--legacy-word` if Word is unavailable; legacy DOC files will be explicitly
reported as unsupported. Copy the repository lookup folders and scripts with the
documents if running on another machine. `--root` can specify a different lookup root.
These commands only edit local files: they neither download nor upload documents.

Resolution uses `Masters`, `FixerFiles`, and `EmbeddedLinks` CSVs, matching exact
URLs, document GUIDs, and document paths. Opaque sharing tokens are case-sensitive.
Only destinations recorded as **found** are used. `FirstPassDocURLLookupTables`
can fill gaps but never override a later mapping or a later "not found".
The site-base map alone is not proof that a file exists; the script deliberately
does not blindly replace site prefixes or guess renamed files. Conflicting and
unresolved URLs stay untouched, even if a report contains an unconfirmed proposed URL.
Mappings that omit or change an ASPX page's list-item/filter/view query context are
also rejected rather than silently changing what the link displays.
With explicit approval, add `--allow-view-filter-drop` to accept a confirmed
destination without its old view/filter/sort parameters. Item IDs and folder
parameters are still protected. When an old link selects a specific file and the
tables confirm that file's new path, that path takes precedence over a less-specific
view in the very same library; conflicting file or cross-library destinations
remain unresolved.
Explicit user confirmations and authenticated live checks can be supplied with
`--confirmed-mappings .\ReferenceFiles\ConfirmedEmbeddedLinkMappings.csv`.
This CSV requires `OldURL`, `NewURL`, `Evidence`, and `DropFragment` (`Yes`/`No`).
Confirmed mappings override older scan evidence by URL/document identity; conflicting
confirmations or invalid/external destinations abort the run. Fragments are preserved
unless explicitly approved for removal, such as an obsolete PDF search fragment
after a confirmed conversion to DOCX. The confirmation CSV is hashed in the run manifest.
Fragments/bookmarks are preserved. Confirmed file destinations use encoded,
direct new file paths rather than old GUIDs or obsolete sharing tokens.

Each run produces:

- `files.csv`: per-file outcome, SHA-256 hashes, errors, and backup path.
- `links.csv`: each tenant hyperlink target/location, old/new URL, status, and lookup evidence.
- `review.csv`: unresolved/conflicting URLs grouped for follow-up.
- `manifest.json`: totals and hashes of the lookup sources used.
- `backups/`: exact original bytes of every file actually replaced.

Office edits retain all ZIP members and only change external hyperlink
relationships, Word HYPERLINK field instructions (including split runs), or literal
first arguments of Excel HYPERLINK formulas. Visible text and cached formula values
are verified unchanged. PDF edits append incremental URI-annotation updates,
preserving every original byte, page drawing/text streams, extracted text, dimensions,
annotation rectangles, and external URI targets. Plain visible PDF URLs without an
existing clickable annotation are not converted into links. Signed/encrypted files
are rejected rather than invalidating signatures or removing protection.

Every changed file is verified in memory, backed up and hash-checked, then replaced
atomically. A nonzero exit code and explicit error row indicate any failed file.
No-change and dry-run files are never rewritten. Keep the backups until the corrected
copies have been reviewed and uploaded. To roll back, copy the selected file from
the run's `backups` tree over the corresponding input file.

Lookup evidence is offline confirmation from the supplied reports, **not** a fresh
availability/permissions check. Unresolved sharing tokens, absent GUID mappings, and
conflicting destinations need fresh authoritative SharePoint data or human review;
installing Acrobat/OCR does not resolve those missing mappings.

## Get-FixItListOverviewOnly.ps1

Tenant-wide remediation finder. Returns one row per affected document/item with broken URL occurrence counts.

Workflow:
1. Connect to SharePoint admin; list eligible site collections (excluding OneDrive personal sites).
2. Operator picks site collection/subsite targets being migrated.
3. Prompt for old base URL → new base URL pairs for selected targets.
4. Scan tenant-wide for impacted references (selection drives mapping only, not scan scope).
5. Write timestamped CSVs — one row per affected document or non-document item.

### Output File Naming
`SiteSlug-Get-FixItListOverviewOnly-ReportType-yyyyMMdd-HHmmss.csv`

### Documents Report Columns
Site Title | Site URL | Library | Document Path | Document URL | Folder URL | Folder |
File Name | File Extension | Created By | Last Modified By | Last Modified Date | Permissions |
Broken URL Occurrences | Matched Old Base URL Count | Matched Old Base URLs | Matched New Base URLs

### Non-Document Report Columns
Site Title | Site URL | List | Content Type | Location | Item Title | Item ID | Item URL |
Broken URL Occurrences | Field Count With Broken URLs | Fields With Broken URLs |
Matched Old Base URL Count | Matched Old Base URLs | Matched New Base URLs


## Get-FixItListURLsv2.ps1

Tenant-wide link-risk audit. Returns one row per individual broken URL found, with full URL-level detail and link classification.

### Output File Naming
`SiteSlug-Get-FixItListURLs-ReportType-yyyyMMdd-HHmmss.csv`

### Documents Report Columns
Link Type | Site Title | Document Title | Document URL | Document Folder | Document Folder URL |
Broken URL Text | Broken Embedded URL | New URL |
Created By | Last Modified By | Last Modified Date | Permissions


## Get-DocumentUrlLookup.ps1

Per-site old-URL -> new-URL lookup for documents, for cases the URL Fixer HTML tool can't pattern-match (links keyed only by the old document's GUID, with no folder path to fall back on). Reuses the site/subsite selection and scannability-inventory caching from Get-FixItListURLsv5.ps1, reads the Old Site URL / New Site URL CSV, then queries the new site live to confirm each document's new path and read its real new GUID.

### Embedded-link reports

Embedded scans inspect human-authored document content, SharePoint page/list content fields,
supported list attachments, and navigation. Only links to the SharePoint tenant are reported;
scans of `/sites` webs retain the existing additional restriction to mapped old-site URLs.
The master document inventory still includes all file types.

The embedded-content allowlist is independent of the Fixer export list:

- Office documents/templates: `.docx`, `.docm`, `.dotx`, `.dotm`, `.xlsx`, `.xlsm`,
  `.xltx`, `.xltm`, `.pptx`, `.pptm`, `.potx`, `.potm`, `.ppsx`, `.ppsm`.
- Other authored documents: `.doc`, `.xls`, `.ppt`, `.rtf`, `.odt`, `.ods`, `.odp`,
  `.vsd`, `.vsdx`, `.one`, `.msg`, `.pdf`.
- Text and web content: `.txt`, `.csv`, `.htm`, `.html`, `.aspx`.

Unsupported files/attachments (including images, video/audio, `.DS_Store`, `.url`,
`.clipchamp`, and `.onetoc2`) are intentionally skipped, not treated as scan failures.
Plain-text/CSV scans report each tenant URL occurrence; PDF scans also inspect visible
URLs, with optional OCR. Legacy Office/OpenDocument/Visio/OneNote/email parsing requires
the configured Java and Apache Tika installation.

`[site]-EmbeddedLinks.csv` contains **one row per reported link instance**, including
repeated URLs with identical display text. `EmbeddedOccurrence` is a sequential instance
number within the source document/item or navigation collection, **not a URL occurrence
count**. `EmbeddedURLText` preserves each instance's extracted display text;
`EmbeddedLocation` identifies the document part/cell, PDF page, list field, attachment,
or navigation node when available. Split formatting runs representing one visible
PowerPoint link remain one instance.

`[site]-EmbeddedLinks-ScanErrors.csv` records actual scan failures from the latest run,
separately from link findings; it has headers only when there are no failures. Warnings
identify incomplete scans. Successfully scanned sources replace prior findings (including
when no matching links remain); failed/unscanned sources retain prior valid findings.
Partial findings do not get mixed into an existing source snapshot, which would make
instance counts misleading; a first-ever partial scan can still publish its known links.
Publishing removes old blank-link diagnostic rows and unsupported-file rows from cached
reports too. Rerun a scan to regenerate instance numbers and location details; prior
unscanned findings cannot acquire details that were never extracted.

### Output File Naming
`SiteSlug-DocumentUrlLookup-yyyyMMdd-HHmmss.csv`

### Columns
Site | Subsite | Library | File Name | Extension | Old File Ref | Old Copy Link URL | Old Resolved URL |
New File Ref | New Found | New Copy Link URL | New Resolved URL | Modified | Notes


## Functional Split

| Need | Script |
|------|--------|
| Per-document/item summary with occurrence counts | Get-FixItListOverviewOnly.ps1 |
| Full per-URL detail with link type classification | Get-FixItListURLsv2.ps1 |
| Deep per-document permissions/access context for a specific site | Get-DocumentAccessReport.ps1 |
| Old/new document URL lookup for bookmark links that can't be pattern-matched | Get-DocumentUrlLookup.ps1 |
