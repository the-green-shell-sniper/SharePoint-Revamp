# SharePoint-Revamp

This repository contains PowerShell scripts used for SharePoint migration analysis and link remediation planning.

## Scripts

- [Get-DocumentAccessReport.ps1](Get-DocumentAccessReport.ps1)
- [Get-FixItListDocsOnly.ps1](Get-FixItListDocsOnly.ps1)
- [Get-FixItListURLs.ps1](Get-FixItListURLs.ps1)

## Get-FixItListDocsOnly.ps1 Overview

[Get-FixItListDocsOnly.ps1](Get-FixItListDocsOnly.ps1) is a tenant-wide remediation finder that uses operator-selected migration mappings.

Its workflow is:
1. Connect to SharePoint admin and list eligible site collections (excluding OneDrive personal sites).
2. Let the operator pick site collection/subsite targets that are being migrated.
3. Prompt for old base URL -> new base URL pairs for those selected targets.
4. Run tenant-wide scanning to find impacted references anywhere in SharePoint content.
5. Write action-oriented CSVs with one row per affected entity and occurrence counts.

Important behavior:
- Site/subsite selection drives mapping definitions only.
- The actual detection pass runs tenant-wide (excluding OneDrive personal sites).
- The document report writes one row per affected document.
- The optional non-document report writes one row per affected item.
- Repeated occurrences of the same matched URL inside one document/item increase the occurrence count.

## Get-FixItListDocsOnly.ps1 Output

Output files are timestamped and written to:
- Documents output folder: Documents-BreakingUrls-yyyymmdd-hhmmss.csv
- NonDocuments output folder: OtherContent-BreakingUrls-yyyymmdd-hhmmss.csv (only when optional scan is enabled)

### Documents Report Columns

- Site Title
- Site URL
- Library
- Document Path
- Document URL
- Folder URL
- Folder
- File Name
- File Extension
- Created By
- Last Modified By
- Last Modified Date
- Permissions
- Broken URL Occurrences
- Matched Old Base URL Count
- Matched Old Base URLs
- Matched New Base URLs

Each row represents one document with at least one URL match against the configured migration mappings.

### Non-Document Report Columns

- Site Title
- Site URL
- List
- Content Type
- Location
- Item Title
- Item ID
- Item URL
- Broken URL Occurrences
- Field Count With Broken URLs
- Fields With Broken URLs
- Matched Old Base URL Count
- Matched Old Base URLs
- Matched New Base URLs

Each row represents one non-document list/page item with at least one URL match against the configured migration mappings.

## Functional Split

- Use [Get-DocumentAccessReport.ps1](Get-DocumentAccessReport.ps1) when you need deep per-document permissions/access context for a specific site migration package.
- Use [Get-FixItListDocsOnly.ps1](Get-FixItListDocsOnly.ps1) when you need tenant-wide remediation targeting based on selected old/new migration URL mappings.
- [Get-FixItListURLs.ps1](Get-FixItListURLs.ps1) remains in the repository for legacy/alternate fix-list workflows.
