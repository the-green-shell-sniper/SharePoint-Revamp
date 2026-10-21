# Repository-wide instructions

## Local development and server execution

- VS Code and the assistant work in the operator's local clone. The server has
  a separate clone at `C:\Users\!bposey\Desktop\GitHub\SharePoint-Revamp`.
  The operator handles transferring changes through Git and running scripts
  on that server. Do not confuse the local OneDrive workspace with that clone.
- This environment has no access to the server or Microsoft tenant services.
  An exception is an explicitly requested built-in browser session where the
  operator signs in interactively. That session does not imply general server,
  PowerShell, or API access. Never capture or persist credentials in the repo
  or local development environment.
- Provide copy/paste operator commands for the server clone, with an explicit
  `Set-Location` when needed. Do not use the local workspace path in server
  commands or hardcode either clone path into portable script defaults.
- Base script-relative defaults on `$PSScriptRoot`. Resolve operator-supplied
  relative paths with
  `$ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath()`
  before using .NET path/file APIs or shared report helpers. PowerShell's
  location may differ from `[Environment]::CurrentDirectory`, especially in an
  elevated shell whose process directory is `C:\Windows\system32`. Validate
  that case locally when changing path handling.
- Server paths embedded in copied reports describe where the server wrote
  them; they are provenance, not local paths to open. Inspect the corresponding
  files in the local clone. Read report evidence before asserting completion.

## Authentication and operational safety

- Microsoft Graph, Exchange, SharePoint, and other service authentication is
  cached on the server, never stored in this repo. Prefer existing supported
  server authentication; if it is unavailable or expired, prompt for an
  interactive login. Never add embedded credentials, tokens, or client secrets.
- Active Directory operations use the currently logged-on Windows user's
  identity, not an interactive AD credential prompt.
- Limit local validation to offline syntax, schema, cmdlet/API availability,
  and focused logic checks that establish whether the script works as written.
  Do not validate against a live environment. The operator runs scripts on the
  server and reports runtime errors; do not demand repeated live scans when
  existing valid evidence suffices.
- Remove temporary diagnostic/test scripts after use. Do not retain or add
  test files, tooling probes, or helpers the operator and production scripts
  do not need. Use session/temp storage for disposable validation and clean up
  the specific temporary files created.
- Never run commit, push, or build commands. The operator decides when changes
  are ready to commit, push, or build; do not perform these after each fix.
- Keep operator output succinct: updating progress, concise results, and
  actionable warnings. Put detailed diagnostics in reports, with optional
  verbose output for troubleshooting. Do not repeatedly state that no tenant
  updates were performed merely because this environment has no tenant access.
- Preserve the requested scope and all unrelated data/metadata. State actual
  limitations clearly; do not confuse offline validation with live verification.

## Generated files and paths

- Any script or tool that creates repository output files must keep **every
  individual file strictly below 99 MiB** (`99 * 1024 * 1024` bytes). Treat this
  as a hard limit, not an estimate. Prefer a lower target (for example 90 MiB)
  to leave room for encoding, headers, and later writes. Enforce the limit while
  streaming/growing output and verify the final on-disk byte size before
  reporting success. Fail explicitly if a single record cannot fit; never leave
  a success-shaped oversized output.
- Large reports must be split into multiple files at record/object boundaries
  whenever possible. CSV parts must each have a header; never split a CSV row.
  Never split JSON into opaque byte chunks or require a user to reassemble it.
  Store large structured datasets as readable, record-based CSV parts with
  headers and documented columns, and update every consumer to read the parts.
  Keep JSON for compact metadata/status only; if an oversized JSON document
  cannot be redesigned as records, fail clearly before writing it.
  Existing legacy JSON can be migrated by `Convert-SPServerReports.ps1` to
  readable token/path/value CSV records with a small JSON index at the original
  path. Updated readers must validate and interpret them automatically in
  memory, never recreate an oversized file on disk. This compatibility format
  is not a substitute for domain-specific rows in new reports.
- JSON is appropriate for compact structured evidence such as a permission
  snapshot, proposed plan, status, or manifest: later stages consume these
  objects to verify results or recover context. It is not exempt from the file
  limit. Keep small status/manifests as JSON; partition large snapshots/plans
  into bounded, valid JSON chunks or another suitable record-oriented format
  such as CSV/NDJSON, and update every producer and consumer together.
- Keep every generated full path, including the repository's absolute path,
  filename, extension, and temporary-file suffix, **under 256 characters**.
  Prefer paths under 240 characters for margin. Keep output directory nesting
  shallow and names concise; use stable short identifiers when needed. Check
  paths before creating directories/files and fail clearly if a path is too long.
- Apply these rules to all generated outputs, including logs, transcripts,
  temporary/atomic-write files, manifests, snapshots, plans, CSVs, and archives.
  Do not rely on Git ignore rules, post-run cleanup, or GitHub rejection as
  enforcement.
- PowerShell scripts under `Scripts` must dot-source
  `Scripts/SPOutputLimits.Common.ps1` after their `param` block. It caps CSV
  outputs at 90 MiB with numbered CSV parts, transparently imports those parts
  through `Import-Csv`, rotates append-only text files to numbered siblings,
  rejects oversized JSON and other `Set-Content` output, and provides
  `Assert-SPOutputPath`, `Assert-SPOutputSize`, and `Write-SPBoundedJson` for
  direct .NET or structured writes. Convert large collections to CSV records
  rather than attempting JSON chunking.
- For report-writer changes, search for all consumers and update them in the
  same change. Test multi-part outputs at the size boundary, verify every part
  is below the cap, and confirm concatenated/combined records preserve the
  original data and order.
