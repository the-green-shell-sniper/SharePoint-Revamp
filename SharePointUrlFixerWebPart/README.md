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
  actual document filename (without its extension) rather than generic list-view page names
  or document Title metadata, which can retain a template's name.
  For document links, the current SharePoint filename takes precedence over the URL's `file` parameter.
  Sites/subsites use their own web title; direct library/list URLs and their views
  use that library/list's title. Folder links (including `id`/`RootFolder` view
  links) use the folder name, not the containing library. List item display/edit
  forms with an item ID use the item's Title, not the list title. Untitled items
  fall back to the list name plus item ID with a lookup notice.
  Subsite documents/pages are queried through their owning web. Metadata for
  direct site/library/folder paths must identify that exact path before its name
  is used; an ancestor container is not accepted as the destination.
  If the title request fails or times out, a notice explains that the name is derived from the URL;
  you can rename the bookmark in Edge. If the bar is hidden, press **Ctrl+Shift+B** to show it.
  The displayed result URL is also draggable.
- Keyboard alternative: open the updated link, press **Ctrl+D** in the new tab,
  and confirm the favorite in Edge.
  The save-help box lists the two bookmark methods first (dragging, then
  keyboard), followed by the desktop shortcut method, as three bullets.
- Manually delete any favorites, bookmarks, or shortcuts that have the old link.
  Dragging creates a favorite;
  it does not find or replace an existing favorite.

The web part cannot access browser favorites or trigger Edge's native favorite
dialog. No extension is required for the drag-to-favorites workflow.

## Submission log

The SPFx web part adds a separate list item for each **Find Updated Link** or
Enter submission, including empty input, rejected URLs, lookup failures, and
warning results. Merely viewing the page, copying/opening a result, or clearing
the input does not create another row. Repeated submissions create separate rows.
Clearing an in-progress lookup does not cancel its log write.

When a matching CSV row is used, `ResultMessage` appends its source filename,
for example `Lookup table: CS-fixer.csv.` This applies to successful mappings,
mapped-link verification warnings, and matched rows marked not migrated.
Pattern-only results and searches with no matching row do not claim a source
table. The filename is included in the log only, not added to app messages.

`BookmarkButtonName` stores the final drag/drop button label, including a
URL-derived fallback name when metadata is unavailable. It is blank for failed
URL submissions without a usable destination. Log creation starts after naming
finishes so it records the same label as the button; clearing a pending submission
does not prevent its log entry from being saved.

The destination is the **URLFixerSubmissionLog** list on
<https://whallc.sharepoint.com/sites/SharePointLinkFixer>, even when the web part
is hosted on another site. Independent item creates avoid the concurrent-write
and file-locking problems of appending to a shared CSV or workbook.
The standalone HTML tool does not use this SPFx logging integration.

### One-time list setup (site owner)

The list has been created at
<https://whallc.sharepoint.com/sites/SharePointLinkFixer/Lists/URLFixerSubmissionLog/AllItems.aspx>.
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
| BookmarkButtonName | BookmarkButtonName | Single line of text | Optional; final button label |
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

The app keeps its pre-logging instructions and does not display logging notices
or logging-failure warnings. Full URLs (including query strings), names,
and login identifiers can contain sensitive information. Restrict list/export
access and apply your organization's retention policy; this is a usage log, not
a tamper-proof compliance audit.

### Excel export and deployment verification

Open the list and use **Export > Export to Excel** to analyze entries. Include
DisplayName, Username, ProvidedURL, ReturnedURL, Outcome, ResultMessage,
BookmarkButtonName, SubmittedAt, Created, and Created By in the export view.

Build/package the web part and deploy the updated `.sppkg` through the existing
app catalog process. With a regular user account, submit a valid old URL and an
invalid URL, then confirm two rows with the correct identity, full input, and
returned link (blank for the error). Also verify two simultaneous users produce
separate rows, and that ordinary users cannot see other users' entries.

If a write fails (missing list/columns, denied permission, network failure, or
throttling), the failure is recorded using SPFx `Log.error` under
`URLFixerSubmissionLog` for developer diagnostics, not displayed in the app.
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
`/sites/SharePointLinkFixer/URLLookupTables`, reads `URLFixerIndex.csv`, and selects
the matching site's fixer CSV. `Lookup table discovery failed (...)` means the
file-list request failed before any CSV rows were searched, not that the URL
was absent from the table.

Discovery targets `/_api/web` on `/sites/SharePointLinkFixer`, the web that owns
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