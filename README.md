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

Deprecated scripts are archived in [Deprecated-Scripts/](Deprecated-Scripts/).

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
identity and the uploaded byte hash. User IDs are reused within the same site;
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
changes and disables automatic query retries. Identical content is skipped.
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
