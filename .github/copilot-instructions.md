# Repository-wide instructions

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
