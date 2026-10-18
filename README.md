# SharePoint-Revamp

This repository contains PowerShell scripts used for SharePoint migration analysis and link remediation planning.

## Scripts

- [Get-DocumentAccessReport.ps1](Scripts/Get-DocumentAccessReport.ps1)
- [Get-DocumentUrlLookup.ps1](Scripts/Get-DocumentUrlLookup.ps1)
- [Get-FixItListOverviewOnly.ps1](Scripts/Get-FixItListOverviewOnly.ps1)
- [Get-FixItListURLsv2.ps1](Scripts/Get-FixItListURLsv2.ps1)
- [Search-PhrasesInSpreadsheets.ps1](Scripts/Search-PhrasesInSpreadsheets.ps1) - searches every CSV/Excel file under a folder tree for each phrase in a phrases CSV (exact match, then fuzzy for phrases with no exact hit) and writes the file path + row of each match to a CSV.

Deprecated scripts are archived in [Deprecated-Scripts/](Deprecated-Scripts/).

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
