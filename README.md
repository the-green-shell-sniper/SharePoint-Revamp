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

### CS only: do this first

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
exact matches, incompatible kinds/templates/extensions, and malformed approvals.
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
- Complete inventory/ACL reads, item counts matching each list's reported count,
  and matching permission-level **names and bitmasks** are required. Count mismatches
  stop the run rather than accepting security-trimmed/changing inventories.
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
