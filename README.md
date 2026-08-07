# SharePoint-Revamp

This repository contains PowerShell scripts used for SharePoint migration analysis and link remediation planning.

## Scripts

- [Get-DocumentAccessReport.ps1](Scripts/Get-DocumentAccessReport.ps1)
- [Get-FixItListOverviewOnly.ps1](Scripts/Get-FixItListOverviewOnly.ps1)
- [Get-FixItListURLsv2.ps1](Scripts/Get-FixItListURLsv2.ps1)

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


## Functional Split

| Need | Script |
|------|--------|
| Per-document/item summary with occurrence counts | Get-FixItListOverviewOnly.ps1 |
| Full per-URL detail with link type classification | Get-FixItListURLsv2.ps1 |
| Deep per-document permissions/access context for a specific site | Get-DocumentAccessReport.ps1 |
