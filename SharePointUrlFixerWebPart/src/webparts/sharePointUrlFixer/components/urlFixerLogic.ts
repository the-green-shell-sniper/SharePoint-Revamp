import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

// Path-segment rename rules, shared by the main URL path and by any embedded
// document-path query parameters (id=, sourcedoc=, sourcedocument=).
export const renameRules: Array<[string, string]> = [
  ['whaiis01', 'DocumentsandForms'],
  ['CS', 'ClinicalandPatientServices'],
  ['CRL', 'ClinicianResourceLibrary'],
  ['CtrlAltDelight_Newsletter', 'CtrlAltDelightNewsletter'],
  ['IT&HIM', 'IT'],
  ['LLCM', 'LLCMembers'],
  ['PFS', 'PatientFinancialServices'],
  ['PFTC', 'PatientFocusTeamCare'],
  ['Private', 'Staging'],
  ['SP', 'StrategicPlan'],
  ['TPE', 'ThePatientExperience'],
  ['train', 'Training'],
  ['COVID19%20Resources', 'ClinicalandPatientServices/COVID19%20Resources'],
  ['COVID19%20Workflows', 'ClinicalandPatientServices/COVID19%20Workflows']
];

const OFFICE_EXTENSION_CODES: { [ext: string]: string } = {
  doc: 'w', docx: 'w', docm: 'w',
  xls: 'x', xlsx: 'x', xlsm: 'x',
  ppt: 'p', pptx: 'p', pptm: 'p',
  pdf: 'b'
};

export interface IFixResult {
  original: string;
  fixed: string;
  note: string; // "<level>:<message>", level is ok/warn/err
}

export interface ILookupRow {
  [column: string]: string;
}

function escapeRegex(str: string): string {
  return str.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function detectOfficeCode(pathAndQuery: string): string | undefined {
  const lastSegment = pathAndQuery.split('?')[0].split('/').pop() || '';
  const extMatch = lastSegment.match(/\.([a-z0-9]+)$/i);
  return extMatch ? OFFICE_EXTENSION_CODES[extMatch[1].toLowerCase()] || undefined : undefined;
}

function fixPathQueryValue(rawValue: string): { value: string; changed: boolean } {
  const usesEncodedSlash = /%2F/i.test(rawValue);
  const usesLiteralSlash = rawValue.indexOf('/') !== -1;
  if (!usesEncodedSlash && !usesLiteralSlash) {
    return { value: rawValue, changed: false };
  }

  let normalized = usesEncodedSlash ? rawValue.replace(/%2F/gi, '/') : rawValue;
  const hadLeadingSlash = normalized.startsWith('/');
  if (!hadLeadingSlash) normalized = '/' + normalized;

  const sitesAlready = /^\/sites(\/|$)/i.test(normalized);
  let result = normalized;
  let addedSites = false;
  if (!sitesAlready) {
    result = '/sites' + result;
    addedSites = true;
  }

  let renamed = false;
  renameRules.forEach(([from, to]) => {
    // eslint-disable-next-line @rushstack/security/no-unsafe-regexp -- built from our own hardcoded, pre-escaped renameRules constant, never from user input
    const segmentPattern = new RegExp('\\/' + escapeRegex(from) + '(?=\\/|$)', 'g');
    if (segmentPattern.test(result)) {
      result = result.replace(segmentPattern, '/' + to);
      renamed = true;
    }
  });

  let value = hadLeadingSlash ? result : result.slice(1);
  if (usesEncodedSlash) value = value.replace(/\//g, '%2F');

  return { value, changed: addedSites || renamed };
}

function fixDocumentIdParams(url: string): { url: string; changed: boolean } {
  let changed = false;
  const updatedUrl = url.replace(/([?&](?:id|sourcedoc|sourcedocument|parent)=)([^&]*)/gi, (match, paramPrefix, paramValue) => {
    const fixedParam = fixPathQueryValue(paramValue);
    if (fixedParam.changed) changed = true;
    return paramPrefix + fixedParam.value;
  });
  return { url: updatedUrl, changed };
}

function stripLegacyDocIdParams(url: string): { url: string; removed: boolean } {
  if (!/[?&]d=/i.test(url)) return { url, removed: false };
  let stripped = url.replace(/[?&](?:d|csf|e)=[^&]*/gi, '');
  stripped = stripped.replace(/[?&]{2,}/, (m) => m[0]).replace(/[?&]$/, '');
  if (!stripped.includes('?')) stripped = stripped.replace(/&/, '?');
  return { url: stripped, removed: true };
}

function isUnresolvableDocAspxLink(url: string): boolean {
  return /_layouts\/15\/doc\.aspx/i.test(url) && /[?&]sourcedoc=/i.test(url);
}

// Pattern-matches an old whallc.sharepoint.com URL to its migrated equivalent. Cannot
// resolve links keyed only by the old document's GUID with no folder path to fall
// back on - those are reported as unresolvable so the CSV lookup can be tried instead.
export function fixUrl(raw: string): IFixResult {
  const url = raw.trim();

  const hostMatch = url.match(/^(https?:\/\/)?(whallc\.sharepoint\.com)/i);
  if (!hostMatch) {
    return { original: url, fixed: url, note: "err:Doesn't look like a whallc.sharepoint.com link — left unchanged." };
  }
  const protocol = hostMatch[1] || 'https://';
  const host = hostMatch[2];
  const afterHost = url.slice(hostMatch[0].length);

  const prefixMatch = afterHost.match(/^\/:[a-zA-Z]:(?:\/[a-zA-Z]+)?\//);
  let prefix = prefixMatch ? prefixMatch[0] : '';
  let afterPrefix = prefixMatch ? afterHost.slice(prefix.length) : afterHost;

  let addedWebPrefix = false;
  let webPrefixCode: string | null = null;
  if (!prefix && !/_layouts\/15\/doc\.aspx/i.test(afterHost)) {
    const officeCode = detectOfficeCode(afterHost);
    if (officeCode) {
      prefix = '/:' + officeCode + ':/r/';
      afterPrefix = afterHost.replace(/^\//, '');
      addedWebPrefix = true;
      webPrefixCode = officeCode;
    }
  }

  const alreadyHasSites = prefix ? /^sites\//i.test(afterPrefix) : /^\/sites\//i.test(afterPrefix);

  let fixed: string;
  let addedSites: boolean;
  if (alreadyHasSites) {
    fixed = protocol + host + prefix + afterPrefix;
    addedSites = false;
  } else {
    const insertedPart = prefix ? 'sites/' + afterPrefix : '/sites' + afterPrefix;
    fixed = protocol + host + prefix + insertedPart;
    addedSites = true;
  }

  const renamedSegments: string[] = [];
  renameRules.forEach(([from, to]) => {
    // eslint-disable-next-line @rushstack/security/no-unsafe-regexp -- built from our own hardcoded, pre-escaped renameRules constant, never from user input
    const segmentPattern = new RegExp('\\/' + escapeRegex(from) + '(?=\\/|$)', 'g');
    if (segmentPattern.test(fixed)) {
      fixed = fixed.replace(segmentPattern, '/' + to);
      renamedSegments.push('/' + from + ' → /' + to);
    }
  });
  const csRenamed = renamedSegments.length > 0;

  const idFix = fixDocumentIdParams(fixed);
  fixed = idFix.url;
  const idParamsChanged = idFix.changed;

  if (isUnresolvableDocAspxLink(fixed)) {
    return {
      original: url,
      fixed,
      note: "err:This is a resolved Doc.aspx link keyed only by the old document's GUID (sourcedoc=), with no folder path to fall back on. The migrated copy has a new GUID that can't be derived from this link — open the document's folder in the new site and copy a fresh link instead."
    };
  }

  const docIdStrip = stripLegacyDocIdParams(fixed);
  fixed = docIdStrip.url;

  const parts: string[] = [];
  if (addedWebPrefix) parts.push('added :' + webPrefixCode + ':/r/ so it opens in the web app instead of a desktop app');
  if (addedSites) parts.push('Added /sites');
  if (csRenamed) parts.push('renamed ' + renamedSegments.join(', '));
  if (idParamsChanged) parts.push('updated the id/sourcedoc/parent path(s) in the query string to match');
  if (docIdStrip.removed) parts.push("removed the old document's id/csf/e tracking params (the fixed folder path is enough for SharePoint to resolve the new copy)");

  const note = parts.length > 0 ? 'ok:' + parts.join(', ') + '.' : 'warn:Already has /sites — left unchanged.';
  return { original: url, fixed, note };
}

// --- CSV lookup-table discovery, kept as a fallback when pattern matching can't
// resolve a link (e.g. a GUID-only Doc.aspx link) or the pattern-fixed URL 404s. ---

const LOOKUP_FOLDER = '/sites/SharePointURLFixer/URLLookupTables';
const INDEX_FILE_NAME = 'URLFixerIndex.csv';
const HELP_DESK_MESSAGE = 'No migrated URL found. Please contact the Help Desk for assistance by calling 503-601-3646 or emailing servicedesk@whallc.com.';

function parseCsvLine(line: string): string[] {
  const values: string[] = [];
  let value = '';
  let quoted = false;
  for (let i = 0; i < line.length; i++) {
    const char = line[i];
    if (char === '"') {
      if (quoted && line[i + 1] === '"') { value += '"'; i++; }
      else { quoted = !quoted; }
    } else if (char === ',' && !quoted) {
      values.push(value); value = '';
    } else {
      value += char;
    }
  }
  values.push(value);
  return values;
}

function parseLookupCsv(text: string): ILookupRow[] {
  const lines = text.replace(/^\uFEFF/, '').split(/\r?\n/).filter((line) => line.trim());
  if (lines.length < 2) return [];
  const headers = parseCsvLine(lines[0]);
  return lines.slice(1).map((line) => {
    const values = parseCsvLine(line);
    const row: ILookupRow = {};
    headers.forEach((header, index) => { row[header] = values[index] || ''; });
    return row;
  });
}

function normalizeUrlBase(value: string): string {
  try {
    const parsed = new URL(value);
    const normalizedPath = parsed.pathname
      .replace(/^\/:[a-z]:(?:\/[a-z]+)?\//i, '/')
      .replace(/\/$/, '');
    return `${parsed.origin}${normalizedPath}`.toLowerCase();
  } catch {
    return value.trim().replace(/\/$/, '').toLowerCase();
  }
}

function findIndexRow(url: string, indexRows: ILookupRow[]): ILookupRow | undefined {
  const normalizedInput = normalizeUrlBase(url);
  return indexRows
    .filter((row) => row.OldBaseURL && normalizedInput.startsWith(normalizeUrlBase(row.OldBaseURL)))
    .sort((a, b) => normalizeUrlBase(b.OldBaseURL).length - normalizeUrlBase(a.OldBaseURL).length)[0];
}

function decodeUrlPart(value: string): string {
  try { return decodeURIComponent(value).toLowerCase(); }
  catch { return value.toLowerCase(); }
}

function normalizeDocumentGuid(value: string): string {
  return decodeUrlPart(value).replace(/[{}-]/g, '').replace(/^[a-z](?=[0-9a-f]{32}$)/i, '');
}

function getUrlSignature(url: string): string {
  try {
    const parsed = new URL(url);
    const path = decodeUrlPart(parsed.pathname).replace(/\/$/, '');
    const params = parsed.searchParams;
    if (/_layouts\/15\/doc\.aspx$/i.test(path)) {
      return `guid|${normalizeDocumentGuid(params.get('sourcedoc') || '')}`;
    }
    if (params.has('id') || params.has('parent')) {
      return `pdf|${path}|${decodeUrlPart(params.get('id') || '')}|${decodeUrlPart(params.get('parent') || '')}`;
    }
    if (params.has('d')) {
      return `guid|${normalizeDocumentGuid(params.get('d') || '')}`;
    }
    return `path|${path.replace(/^\/:\w:\/r/i, '')}`;
  } catch {
    return `raw|${url.trim().toLowerCase()}`;
  }
}

function getOldFileRefSignature(url: string): string {
  try {
    const parsed = new URL(url);
    let path = decodeUrlPart(parsed.pathname).replace(/\/$/, '');
    path = path.replace(/^\/:\w:\/r/i, '');
    return path;
  } catch {
    return '';
  }
}

export interface ILookupFile {
  Name: string;
  ServerRelativeUrl: string;
}

function getAbsoluteServerRelativeUrl(siteAbsoluteUrl: string, serverRelativeUrl: string): string {
  return new URL(siteAbsoluteUrl).origin + serverRelativeUrl;
}

async function readJsonResponse<T>(response: SPHttpClientResponse, context: string): Promise<T> {
  const text = await response.text();
  if (!response.ok) throw new Error(`${context} failed (${response.status}).`);
  try {
    return JSON.parse(text) as T;
  } catch {
    throw new Error(`${context} returned HTML instead of JSON. Confirm the URL Lookup Tables library path and user permissions.`);
  }
}

export class LookupTableClient {
  private lookupFilesPromise: Promise<ILookupFile[]> | undefined;
  private lookupIndexPromise: Promise<ILookupRow[]> | undefined;
  private readonly lookupRowsByFile = new Map<string, ILookupRow[]>();

  constructor(private readonly spHttpClient: SPHttpClient, private readonly siteAbsoluteUrl: string) {}

  private async getLookupFiles(): Promise<ILookupFile[]> {
    if (!this.lookupFilesPromise) {
      const endpoint = `${this.siteAbsoluteUrl}/_api/web/GetFolderByServerRelativePath(decodedurl='${LOOKUP_FOLDER}')/Files?$select=Name,ServerRelativeUrl&$top=5000`;
      this.lookupFilesPromise = this.spHttpClient
        .get(endpoint, SPHttpClient.configurations.v1, { headers: { Accept: 'application/json;odata=nometadata' } })
        .then((response: SPHttpClientResponse) => readJsonResponse<{ value?: ILookupFile[] }>(response, 'Lookup table discovery'))
        .then((data: { value?: ILookupFile[] }) => (data.value || []).filter((file) => /\.csv$/i.test(file.Name)));
    }
    return this.lookupFilesPromise;
  }

  private async getLookupIndex(): Promise<ILookupRow[]> {
    if (!this.lookupIndexPromise) {
      const files = await this.getLookupFiles();
      const indexFile = files.find((file) => file.Name.toLowerCase() === INDEX_FILE_NAME.toLowerCase());
      if (!indexFile) throw new Error(`Could not find ${INDEX_FILE_NAME} in the URL Lookup Tables library.`);

      this.lookupIndexPromise = this.spHttpClient
        .get(getAbsoluteServerRelativeUrl(this.siteAbsoluteUrl, indexFile.ServerRelativeUrl), SPHttpClient.configurations.v1)
        .then((response: SPHttpClientResponse) => {
          if (!response.ok) throw new Error(`Could not load ${INDEX_FILE_NAME} (${response.status}).`);
          return response.text();
        })
        .then(parseLookupCsv);
    }
    return this.lookupIndexPromise;
  }

  public async findLookupRow(url: string): Promise<ILookupRow | undefined> {
    const indexRows = await this.getLookupIndex();
    const indexRow = findIndexRow(url, indexRows);
    if (!indexRow) return undefined;

    const files = await this.getLookupFiles();
    const fixerFile = files.find((file) => file.Name.toLowerCase() === indexRow.FixerCsvFileName.toLowerCase());
    if (!fixerFile) return undefined;
    const inputSignature = getUrlSignature(url);
    const inputFileRef = getOldFileRefSignature(url);

    let rows = this.lookupRowsByFile.get(fixerFile.Name);
    if (!rows) {
      const response = await this.spHttpClient.get(
        getAbsoluteServerRelativeUrl(this.siteAbsoluteUrl, fixerFile.ServerRelativeUrl),
        SPHttpClient.configurations.v1
      );
      if (!response.ok) return undefined;
      rows = parseLookupCsv(await response.text());
      this.lookupRowsByFile.set(fixerFile.Name, rows);
    }

    const row = rows.find((candidate) => {
      const exactSignatureMatch = [candidate.OldCopyLinkUrl, candidate.OldResolvedUrl]
        .some((value) => value && getUrlSignature(value) === inputSignature);
      const fileRefMatch = candidate.OldFileRef && getOldFileRefSignature(candidate.OldCopyLinkUrl) === inputFileRef;
      return exactSignatureMatch || fileRefMatch;
    });

    if (row && row.NewFound && row.NewFound.trim().toLowerCase() === 'no') {
      return { ...row, Notes: HELP_DESK_MESSAGE };
    }

    return row;
  }
}

async function isReachableSharePointUrl(url: string): Promise<boolean> {
  try {
    const response = await fetch(url, {
      method: 'HEAD',
      credentials: 'same-origin',
      redirect: 'follow',
      cache: 'no-store'
    });
    return response.ok || (response.status >= 300 && response.status < 400);
  } catch {
    return false;
  }
}

// Tries the fast local pattern fix first, validates it's actually reachable, and
// only falls back to the (slower, network-backed) CSV lookup table when the
// pattern fix is unresolvable or the resulting URL doesn't check out.
export async function lookupOrFixUrl(raw: string, lookupClient: LookupTableClient): Promise<IFixResult> {
  const input = raw.trim();

  const patternResult = fixUrl(input);
  const patternLevel = patternResult.note.slice(0, patternResult.note.indexOf(':'));
  if (patternLevel !== 'err' && (await isReachableSharePointUrl(patternResult.fixed))) {
    return patternResult;
  }

  const lookupRow = await lookupClient.findLookupRow(input);
  if (!lookupRow) return patternResult;

  if (lookupRow.Notes && !lookupRow.NewCopyLinkUrl && !lookupRow.NewResolvedUrl) {
    return {
      original: input,
      fixed: input,
      note: `err:${lookupRow.Notes}`
    };
  }

  const isResolved = /_layouts\/15\/doc\.aspx|[?&](?:id|parent)=/i.test(input);
  return {
    original: input,
    fixed: isResolved
      ? (lookupRow.NewResolvedUrl || lookupRow.NewCopyLinkUrl)
      : (lookupRow.NewCopyLinkUrl || lookupRow.NewResolvedUrl),
    note: 'ok:Pattern match was unavailable or did not resolve; found the migrated document in the SharePoint lookup table.'
  };
}
