# SharePoint-Revamp

This repository contains two PowerShell audit scripts used for SharePoint migration planning.

Primary program goal: identify broken links introduced or exposed by each migration phase, and surface only the links staff actually need to fix.

Important context:
- The links output from [Get-DocumentAccessReport.ps1](Get-DocumentAccessReport.ps1) is intentionally broad and can include non-actionable PDF-extracted URLs (for example, Adobe/versioning-related references).
- [Get-FixItListURLs.ps1](Get-FixItListURLs.ps1) is intended to reduce that noise by writing only URLs that match migration mappings or are externally broken.

## Scripts

- [Get-DocumentAccessReport.ps1](Get-DocumentAccessReport.ps1)
- [Get-FixItListURLs.ps1](Get-FixItListURLs.ps1)

## Functional Differences

| Area | Get-DocumentAccessReport.ps1 | Get-FixItListURLs.ps1 |
|---|---|---|
| Primary purpose | Deep per-document migration report for one selected site and its subwebs, including access and full link extraction context. | Action-oriented fix-list generator for migration-related or broken URLs, with optional non-document scanning. |
| Scope selection | Single site URL (interactive static site menu or explicit -SiteUrl). | Starts from tenant site collections, lets you pick scope/mappings, then scans site collections/webs tenant-wide for matches and broken links. |
| Main output design | Two CSVs: one access-centric document inventory + one per-embedded-link dataset (broad capture). | Two CSVs: one document "fix list" (staff-actionable links) + optional non-document fix list. |
| Document access details | Includes ownership, created/modified timestamps, inheritance mode, AD/SP groups, direct users, sharing links, shared-with users. | Includes streamlined document metadata plus permissions summary and URL-fix columns. |
| Link extraction behavior (documents) | Writes every extracted URL from supported docs and classifies RiskType; benign links are still recorded with RiskType=None. This can include PDF noise such as Adobe-related references. | Writes only links that match migration mappings or are externally broken, which is intended to avoid most non-actionable link noise. |
| Suggested replacement logic | Computes FutureDocumentUrl and SuggestedReplacement using source site path to target site path promotion logic. | Computes New URL by old->new base URL mappings you enter for selected subsites. |
| Non-document coverage | Not included (document-focused script). | Optional scan of non-document list/page text fields (Text/Note/URL fields) into a separate CSV. |
| Site title handling in output | Uses site/web URL context in access report columns. | Uses Site Title in both document and non-document outputs (web title fallback to URL). |
| Default output location | Defaults to local per-site files in current folder unless output paths are provided. | Defaults to fixed Documents and NonDocuments folders under SharePoint-Migration-Scripts/FixItList-OutputFiles. |
| Typical use case | Preparing one site promotion package with detailed access and permission context. | Building migration fix lists across a broader tenant footprint, including non-document content. |

## Output Summary

### Get-DocumentAccessReport.ps1

Produces:
- Access report CSV (one row per document)
- Links report CSV (one row per embedded URL)

Strengths:
- Rich permission and sharing detail
- Good for owner reviews and access validation
- Good for one-site migration planning

Watch-outs:
- Links CSV is broad by design; for PDFs, it may include embedded references not useful for staff remediation.

### Get-FixItListURLs.ps1

Produces:
- Documents fix CSV (rows for risky/broken embedded links)
- NonDocuments fix CSV (optional; risky/broken links in list/page fields)

Strengths:
- Good for broad fix-list generation
- Includes optional non-document content scanning
- Designed for mapping legacy URL bases to new destinations

Watch-outs:
- Although filtered for actionability, review output rows for edge-case false positives in complex page/list content.

## Which Script Should I Use?

Use [Get-DocumentAccessReport.ps1](Get-DocumentAccessReport.ps1) when you need:
- Full per-document access and sharing visibility
- One-site migration readiness and permission review
- Full forensic link extraction (including links that may not require staff action)

Use [Get-FixItListURLs.ps1](Get-FixItListURLs.ps1) when you need:
- A staff-facing fix list of broken/risky links to remediate after migration phases
- Optional non-document URL cleanup targets
- Old-base to new-base URL mapping workflow
