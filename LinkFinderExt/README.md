# Link Finder Extension

> This is a separate copy of the Central Park Link Finder web part (`../SharePointUrlFixerWebPart`)
> with its own solution, feature, and web part IDs, so both can be installed side by side.
> It is an additional way to use the original Link Finder, not a replacement.
> It adds a **Find New Link** Favorites bar button that sends the page a user is on to this tool.

## Quick Find (Favorites bar button)

### Full-page layout

The manifest supports SharePoint full-page apps (`SharePointFullPage`). The extension
provides its own heading, a plain-language start-to-finish guide, a full-width responsive layout,
and one-time Favorites bar setup instructions (including how to show the bar and drag the button).
The guide covers opening an old link, using Find New Link, confirming the new page,
saving it with Ctrl+D and Done, and removing the old favorite. It links back to the original manual tool.
Instructions use short numbered steps; optional cleanup and help are collapsed under expandable help.
The redirect reminder shows just three actions to perform after the destination opens.
It has no URL input or manual lookup form. The setup heading is hidden during lookup and redirect.
On wide screens, setup and usage instructions sit side by side beneath the introduction.
On smaller screens they stack in reading order. The app has no fixed page-width cap;
individual introductory paragraphs retain a readable line length.
Successful lookups show a **Ctrl+D** reminder and automatically open the destination in
the same tab after ten seconds. **Open now**, **Pause**, and **Resume** controls let users
manage the countdown. Navigation replaces the intermediary history entry to avoid a Back loop.

1. Build and upload `sharepoint/solution/link-finder-quick-find.sppkg` to the App Catalog,
   then add the **Link Finder Extension** web part to the published page
   `https://whallc.sharepoint.com/sites/CentralParkLinkFinder/SitePages/Link-Finder-Extension.aspx`.
2. The web part shows a **Find New Link** button under **First, add the button. You only do this once.**
   Users drag it to their Edge Favorites bar once. Clicking it on the
   Quick Find page itself only shows a reminder to drag it.
3. When an old favorite no longer works, the user opens it and then selects **Find New Link**
   on the Favorites bar. The button (a JavaScript favorite, or "bookmarklet") opens the Quick Find page
   with the current address in the `url` parameter, for example
   `Link-Finder-Extension.aspx?url=https%3A%2F%2Fwhallc.sharepoint.com%2FCS%2F...`.
4. The web part looks up that link automatically, reminds the user to press **Ctrl+D once
   the destination opens**, then redirects. Already-current links return to the same destination.
   Errors and unverified mappings stay on the extension page with Help Desk/manual-tool guidance;
   unverified destinations require an explicit click. Opening with a `?url=` link from email or
   Teams works the same way. Only Central Park links on `whallc.sharepoint.com` are supported.

The button targets the published `Link-Finder-Extension.aspx` page, not a temporary
hosting or preview address. If that page moves, update `QUICK_FIND_PAGE_URL` in
`src/webparts/sharePointUrlFixer/components/quickFind.ts` and redeploy.
Existing favorites retain their saved destination: remove and re-add the **Find New Link**
favorite after updating the destination.

Limitations to test before rollout:

- Organizations can block JavaScript favorites through Edge policy, and Edge may block a JavaScript favorite on
  some pages. Test the button on the old SharePoint error pages users will land on.
- If an old link redirects to a different address before the page loads, the button sends
  the address shown in the address bar, not the original favorite's address.
- The button does not run on `edge://` pages, the new tab page, or PDF viewers.

# Central Park Link Finder

## Summary

SharePoint Framework web part that finds migrated WHA SharePoint URLs using
pattern matching and SharePoint lookup tables.

## Saving the destination

After confirming the destination, press **Ctrl+D** to save it as an Edge favorite.
The web part cannot display a hover prompt after leaving the extension page or trigger
Edge's native favorite dialog. Users must remove obsolete favorites themselves.
Use the separate original Link Finder for manual copy/paste lookups.

## Submission log

The extension logs completed automatic lookup results, including lookup errors and warnings.
Viewing setup instructions does not create a row. Invalid launch parameters and unsupported
source hosts are rejected before lookup. Unexpected request failures are displayed to users.

When a matching CSV row is used, `ResultMessage` appends its source filename,
for example `Lookup table: CS-fixer.csv.` This applies to successful mappings,
mapped-link verification warnings, and matched rows marked not migrated.
Pattern-only results and searches with no matching row do not claim a source
table. The filename is included in the log only, not added to app messages.

`BookmarkButtonName` is blank: the extension no longer fetches a destination title or
generates a per-result favorite button. Logging failures are reported to the browser
console and do not block redirecting.

The destination is the **URLFixerSubmissionLog** list on
<https://whallc.sharepoint.com/sites/CentralParkLinkFinder>, even when the web part
is hosted on another site. Independent item creates avoid the concurrent-write
and file-locking problems of appending to a shared CSV or workbook.
The standalone HTML tool does not use this SPFx logging integration.

### One-time list setup (site owner)

The list has been created at
<https://whallc.sharepoint.com/sites/CentralParkLinkFinder/Lists/URLFixerSubmissionLog/AllItems.aspx>.
The web part uses the internal names verified on that list, including
`Display_x0020_Name`, `ProvidedURL`, and `ReturnedURL`.
Renaming a column's display label does not change its internal name.

For a replacement list, create a **Blank list** on that site named exactly
`URLFixerSubmissionLog`. Keep the default **Title** column. Create the columns
using the initial names below, so SharePoint generates the required internal names.
After creating **Display Name**, you may rename its display label to **DisplayName**.

| Initial column name | Internal name | SharePoint type | Settings |
| --- | --- | --- | --- |
| Display Name | Display_x0020_Name | Single line of text | Signed-in display name |
| Username | Username | Single line of text | Signed-in login/claims name |
| ProvidedURL | ProvidedURL | Multiple lines of text | **Plain text**; optional; **Append changes: No** |
| ReturnedURL | ReturnedURL | Multiple lines of text | **Plain text**; optional; **Append changes: No** |
| Outcome | Outcome | Choice | `Success`, `Warning`, `Error`; single selection |
| ResultMessage | ResultMessage | Multiple lines of text | **Plain text**; **Append changes: No** |
| BookmarkButtonName | BookmarkButtonName | Single line of text | Blank for extension lookups |
| SubmittedAt | SubmittedAt | Date and time | Include time |

**Multiple lines of text alone is not sufficient.** SharePoint may default to
rich text. For **ProvidedURL**, **ReturnedURL**, and **ResultMessage**, open
**Settings > List settings > Columns**, select each column, and set
**Specify the type of text to allow** to **Plain text** (not Rich text or Enhanced
rich text). Set **Append Changes to Existing Text** to **No**, then save.
Plain text and append changes off were confirmed on the live list.
The Email column has been removed and is no longer sent by the web part.
The default All Items view displays DisplayName, Username, ProvidedURL,
ReturnedURL, Outcome, ResultMessage, BookmarkButtonName, and SubmittedAt.
Title is excluded from this view but retained as a column; the app still writes
`URL submission` to it. To hide Title on a replacement list, edit the default
view and deselect Title (including its linked variants), rather than deleting
the column. Column display order does not affect REST writes, which use internal names.

Do not use Hyperlink or single-line fields for URLs: URLs may exceed their
255-character limit. Input text is preserved as submitted; URL processing still
uses the existing whitespace/newline normalization. `ReturnedURL` is blank for
errors, because the lookup's echoed input is not a usable returned link.
Warning results retain their usable destination.

`SubmittedAt` is captured at submission time in UTC using the browser clock.
The built-in **Created** field is SharePoint's authoritative server save time,
and **Created By** is the authenticated writer. The web part supplies the
username and display name from SPFx's signed-in page context.

### Permissions and privacy

Writes use the current user's SharePoint permissions, not an elevated account
or an embedded credential. Users need permission to **add items** on this list;
read-only lookup-table access is not sufficient. Site owners should configure
list-specific permissions rather than granting Edit access to the entire site.
For ordinary users, prefer a custom permission level allowing viewing/adding
items without editing/deleting them, with **Advanced settings > Item-level
Permissions > Read items that were created by the user**. Administrators with
Manage Lists permission can review all entries. Verify with a non-owner account;
do not leave broad inherited Member/Edit access if the log should be restricted.

The app does not display logging notices or logging-failure warnings.
Full URLs (including query strings), names,
and login identifiers can contain sensitive information. Restrict list/export
access and apply your organization's retention policy; this is a usage log, not
a tamper-proof compliance audit.

### Excel export and deployment verification

Open the list and use **Export > Export to Excel** to analyze entries. Include
DisplayName, Username, ProvidedURL, ReturnedURL, Outcome, ResultMessage,
BookmarkButtonName, SubmittedAt, Created, and Created By in the export view.

Build/package the web part and deploy the updated `.sppkg` through the existing
app catalog process. With a regular user account, launch the favorite from a valid old URL
and a Central Park URL with no migrated match, then confirm two rows with the correct
identity, full input, and returned link (blank for the error). Also verify two simultaneous users produce
separate rows, and that ordinary users cannot see other users' entries.

If a write fails (missing list/columns, denied permission, network failure, or
throttling), the failure is recorded using SPFx `Log.error` under
`QuickFindSubmissionLog` for developer diagnostics, not displayed in the app.
Once the bookmark name is resolved, the log write runs independently and does
not delay displaying the URL result.
Failed/unconfirmed writes are not automatically
retried, since an interrupted response may already have created a row.
Logging is best effort: closing/navigating away from the browser before a request
finishes can prevent a save. The web part does not provision the list or change
permissions automatically.

## Lookup-table troubleshooting

The `/CS` to `/sites/ClinicalandPatientServices` rename rule remains enabled.
Old `Doc.aspx?sourcedoc=...` links need the CSV fallback because the migrated
document has a different GUID; renaming the site alone cannot fix them.
The root-site `/2021 EM Coding Changes` and `/2021 E&M Coding Changes` library
paths are pattern-mapped to `/sites/PatientFinancialServices/2021 EM Coding Changes`;
GUID-only links for those files still use the root-site CSV lookup table.

The fallback discovers CSV files in
`/sites/CentralParkLinkFinder/URLLookupTables`, reads `URLFixerIndex.csv`, and selects
the matching site's fixer CSV. `Lookup table discovery failed (...)` means the
file-list request failed before any CSV rows were searched, not that the URL
was absent from the table.

Discovery targets `/_api/web` on `/sites/CentralParkLinkFinder`, the web that owns
the library, rather than the web hosting the web part. The REST web and folder
must agree; otherwise SharePoint reports
`Server relative urls must start with SPWeb.ServerRelativeUrl`.

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