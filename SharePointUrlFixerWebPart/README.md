# share-point-url-fixer

## Summary

SharePoint Framework web part that finds migrated WHA SharePoint URLs using
pattern matching and SharePoint lookup tables.

## Using the URL fixer

- Copy the full URL of a WHA SharePoint page or file from the browser's address bar.
- Paste the URL into the tool and press **Enter**, or select **Find Updated Link**.
- Use the updated result to open the migrated page or copy its URL. The tool processes one URL at a time.
- External links are rejected with a red error message before any network lookup.
- Use **Copy Updated Link** to copy the result, or **Open Updated Link** to open it in a new tab.
- To save without opening the destination, drag the titled button in **Save as a bookmark or shortcut**
  onto Edge's bookmarks bar to create a bookmark or your desktop to create a shortcut. It uses the SharePoint page/site title, list/library title, or
  document title/filename rather than generic list-view page names.
  If the title request fails or times out, a notice explains that the name is derived from the URL;
  you can rename the bookmark in Edge. If the bar is hidden, press **Ctrl+Shift+B** to show it.
  The displayed result URL is also draggable.
- Keyboard alternative: open the updated link, press **Ctrl+D** in the new tab,
  and confirm the favorite in Edge.
- Manually delete any favorites, bookmarks, or shortcuts that have the old link.
  Dragging creates a favorite;
  it does not find or replace an existing favorite.

The web part cannot access browser favorites or trigger Edge's native favorite
dialog. No extension is required for the drag-to-favorites workflow.

## Lookup-table troubleshooting

The `/CS` to `/sites/ClinicalandPatientServices` rename rule remains enabled.
Old `Doc.aspx?sourcedoc=...` links need the CSV fallback because the migrated
document has a different GUID; renaming the site alone cannot fix them.
The root-site `/2021 EM Coding Changes` and `/2021 E&M Coding Changes` library
paths are pattern-mapped to `/sites/PatientFinancialServices/2021 EM Coding Changes`;
GUID-only links for those files still use the root-site CSV lookup table.

The fallback discovers CSV files in
`/sites/SharePointURLFixer/URLLookupTables`, reads `URLFixerIndex.csv`, and selects
the matching site's fixer CSV. `Lookup table discovery failed (...)` means the
file-list request failed before any CSV rows were searched, not that the URL
was absent from the table.

Discovery uses the OData 4 `application/json;odata.metadata=none` Accept header
to match `SPHttpClient`'s default protocol. SharePoint error details are displayed
when available, and failed table loads can be retried without reloading the page.
Missing or inaccessible CSV files are reported as table errors rather than
as missing migrated documents. After deploying an updated package, verify a
known GUID-based link and confirm the current user can read the lookup library.

Lookup matching prioritizes the exact saved URL, then the document path (including
PDF `id` paths), then the old document GUID. Equally ranked matches with different
destinations or migration statuses produce a conflict error rather than selecting
the first CSV row. Identical duplicate mappings remain usable.

CSV downloads are validated before caching: required columns, field counts,
WHA document URLs, and Yes/No migration statuses must be valid. Quoted multiline
fields are supported. HTML sign-in responses and malformed tables produce explicit
errors and can be retried after correction.

When a table mapping exists but the destination check fails, the mapped link
remains available with a warning explaining the failure, rather than claiming
no mapping exists. Access-denied/sign-in redirects are not treated as success.
Availability checks cannot guarantee that a generic HTTP 200 page represents
the intended document.

## Used SharePoint Framework Version

![version](https://img.shields.io/badge/version-1.23.2-green.svg)

## Applies to

- [SharePoint Framework](https://aka.ms/spfx)
- [Microsoft 365 tenant](https://docs.microsoft.com/sharepoint/dev/spfx/set-up-your-developer-tenant)

> Get your own free development tenant by subscribing to [Microsoft 365 developer program](http://aka.ms/o365devprogram)

## Prerequisites

> Any special pre-requisites?

## Solution

| Solution    | Author(s)                                               |
| ----------- | ------------------------------------------------------- |
| folder name | Author details (name, company, twitter alias with link) |

## Version history

| Version | Date             | Comments        |
| ------- | ---------------- | --------------- |
| 1.1     | March 10, 2021   | Update comment  |
| 1.0     | January 29, 2021 | Initial release |

## Disclaimer

**THIS CODE IS PROVIDED _AS IS_ WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

---

## Minimal Path to Awesome

- Clone this repository
- Ensure that you are at the solution folder
- in the command-line run:
  - `npm install -g @rushstack/heft`
  - `npm install`
  - `heft start`

> Include any additional steps as needed.

Other build commands can be listed using `heft --help`.

## Features

Description of the extension that expands upon high-level summary above.

This extension illustrates the following concepts:

- topic 1
- topic 2
- topic 3

> Notice that better pictures and documentation will increase the sample usage and the value you are providing for others. Thanks for your submissions advance.

> Share your web part with others through Microsoft 365 Patterns and Practices program to get visibility and exposure. More details on the community, open-source projects and other activities from http://aka.ms/m365pnp.

## References

- [Getting started with SharePoint Framework](https://docs.microsoft.com/sharepoint/dev/spfx/set-up-your-developer-tenant)
- [Building for Microsoft teams](https://docs.microsoft.com/sharepoint/dev/spfx/build-for-teams-overview)
- [Use Microsoft Graph in your solution](https://docs.microsoft.com/sharepoint/dev/spfx/web-parts/get-started/using-microsoft-graph-apis)
- [Publish SharePoint Framework applications to the Marketplace](https://docs.microsoft.com/sharepoint/dev/spfx/publish-to-marketplace-overview)
- [Microsoft 365 Patterns and Practices](https://aka.ms/m365pnp) - Guidance, tooling, samples and open-source controls for your Microsoft 365 development
- [Heft Documentation](https://heft.rushstack.io/)