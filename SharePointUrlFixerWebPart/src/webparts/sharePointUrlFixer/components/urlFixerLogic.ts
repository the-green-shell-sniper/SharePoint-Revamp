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
  ['2021%20EM%20Coding%20Changes', 'PatientFinancialServices/2021%20EM%20Coding%20Changes'],
  ['2021%20E%26M%20Coding%20Changes', 'PatientFinancialServices/2021%20EM%20Coding%20Changes'],
  ['COVID19%20Resources', 'ClinicalandPatientServices/COVID19%20Resources'],
  ['COVID19%20Workflows', 'ClinicalandPatientServices/COVID19%20Workflows']
];

const OFFICE_EXTENSION_CODES: { [ext: string]: string } = {
  doc: 'w', docx: 'w', docm: 'w',
  xls: 'x', xlsx: 'x', xlsm: 'x',
  ppt: 'p', pptx: 'p', pptm: 'p',
  pdf: 'b'
};

const WHA_SHAREPOINT_HOST_PATTERN = /^(https?:\/\/)?(whallc\.sharepoint\.com)(?=\/|[?#]|$)/i;

export interface IFixResult {
  original: string;
  fixed: string;
  note: string; // "<level>:<message>", level is ok/warn/err
  lookupTableFileName?: string;
}

export interface ILookupRow {
  [column: string]: string;
}

export interface ILookupMatch extends ILookupRow {
  lookupTableFileName: string;
}

function findLegacySiteRule(segment: string): [string, string] | undefined {
  const decoded = decodeUrlPart(segment);
  return renameRules.find(([from]) => decodeUrlPart(from) === decoded);
}

// Renames only the site segment right after the first /sites/, never folders deeper in the path.
function renameSiteSegment(value: string): { value: string; from?: string; to?: string } {
  const match = value.match(/\/sites\/([^/?#]*)/i);
  if (!match || match.index === undefined) return { value };
  const rule = findLegacySiteRule(match[1]);
  if (!rule) return { value };
  const start = match.index + '/sites/'.length;
  return { value: value.slice(0, start) + rule[1] + value.slice(start + match[1].length), from: rule[0], to: rule[1] };
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

  const siteRename = renameSiteSegment(result);
  result = siteRename.value;
  const renamed = siteRename.from !== undefined;

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

  const hostMatch = url.match(WHA_SHAREPOINT_HOST_PATTERN);
  if (!hostMatch) {
    return {
      original: url,
      fixed: url,
      note: "err:You've provided an external link. This tool is only for WHA SharePoint links. Please provide a link that starts with 'https://whallc.sharepoint.com' to continue."
    };
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

  // Every old site lived at the tenant root, so /sites/<non-legacy name> is already a new-structure link.
  if (alreadyHasSites) {
    const siteSegment = afterPrefix.replace(/^\/?sites\//i, '').split(/[/?#]/)[0];
    if (!findLegacySiteRule(siteSegment)) {
      return { original: url, fixed: url, note: 'warn:This link already points to the new SharePoint site — no changes needed.' };
    }
  }

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

  const siteRename = renameSiteSegment(fixed);
  fixed = siteRename.value;
  const csRenamed = siteRename.from !== undefined;
  const renamedSegments = csRenamed ? ['/' + siteRename.from + ' → /' + siteRename.to] : [];

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

const LOOKUP_WEB_PATH = '/sites/SharePointLinkFixer';
const LOOKUP_FOLDER = `${LOOKUP_WEB_PATH}/URLLookupTables`;
const INDEX_FILE_NAME = 'URLFixerIndex.csv';
const HELP_DESK_MESSAGE = 'No migrated URL found. Please contact the Help Desk for assistance by calling 503-601-3646 or emailing servicedesk@whallc.com.';

function parseLookupCsv(text: string, fileName: string): ILookupRow[] {
  const records: string[][] = [];
  let values: string[] = [];
  let value = '';
  let quoted = false;
  let closedQuote = false;
  const csv = text.replace(/^\uFEFF/, '');
  const invalid = (detail: string): Error => new Error(`Invalid lookup CSV ${fileName}: ${detail}`);
  for (let i = 0; i < csv.length; i++) {
    const char = csv[i];
    if (char === '"') {
      if (quoted && csv[i + 1] === '"') { value += '"'; i++; }
      else if (quoted) { quoted = false; closedQuote = true; }
      else if (!value && !closedQuote) { quoted = true; }
      else { throw invalid('unexpected quote.'); }
    } else if (char === ',' && !quoted) {
      values.push(value); value = ''; closedQuote = false;
    } else if ((char === '\r' || char === '\n') && !quoted) {
      values.push(value);
      if (values.some((cell) => cell.trim())) records.push(values);
      values = []; value = ''; closedQuote = false;
      if (char === '\r' && csv[i + 1] === '\n') i++;
    } else {
      if (closedQuote) throw invalid('unexpected text after a quoted field.');
      value += char;
    }
  }
  if (quoted) throw invalid('unterminated quoted field.');
  values.push(value);
  if (values.some((cell) => cell.trim())) records.push(values);
  const headers = records.shift()?.map((header) => header.trim()) || [];
  const isIndex = fileName.toLowerCase() === INDEX_FILE_NAME.toLowerCase();
  const required = isIndex
    ? ['OldBaseURL', 'FixerCsvFileName']
    : ['NewFound', 'NewCopyLinkUrl', 'NewResolvedUrl'];
  if (headers.some((header) => !header) || new Set(headers).size !== headers.length) {
    throw invalid('empty or duplicate column names.');
  }
  if (required.some((header) => headers.indexOf(header) === -1) ||
      (!isIndex && headers.indexOf('OldCopyLinkUrl') === -1 && headers.indexOf('OldResolvedUrl') === -1)) {
    throw invalid('missing required columns. Confirm the download is a CSV, not an HTML sign-in page.');
  }
  return records.map((cells, rowIndex) => {
    if (cells.length !== headers.length) throw invalid(`record ${rowIndex + 2} has the wrong number of fields.`);
    const row: ILookupRow = {};
    headers.forEach((header, index) => { row[header] = cells[index].trim(); });
    const validUrl = (url: string): boolean => {
      try {
        const parsed = new URL(url);
        return /^https?:$/.test(parsed.protocol) && parsed.hostname.toLowerCase() === 'whallc.sharepoint.com';
      } catch { return false; }
    };
    if (isIndex) {
      if (!validUrl(row.OldBaseURL) || !/^[^/\\]+\.csv$/i.test(row.FixerCsvFileName)) {
        throw invalid(`record ${rowIndex + 2} has an invalid site URL or CSV filename.`);
      }
    } else {
      const oldUrls = [row.OldCopyLinkUrl, row.OldResolvedUrl].filter(Boolean);
      const newUrls = [row.NewCopyLinkUrl, row.NewResolvedUrl].filter(Boolean);
      const found = row.NewFound.toLowerCase();
      if (!oldUrls.length || oldUrls.some((url) => !validUrl(url)) ||
          newUrls.some((url) => !validUrl(url)) || (found !== 'yes' && found !== 'no') ||
          (found === 'yes' && !newUrls.length)) {
        throw invalid(`record ${rowIndex + 2} has invalid URLs or migration status.`);
      }
    }
    return row;
  });
}

function normalizeUrlBase(value: string): string {
  try {
    const parsed = new URL(value);
    const normalizedPath = decodeUrlPart(parsed.pathname)
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
    .filter((row) => {
      if (!row.OldBaseURL) return false;
      const base = normalizeUrlBase(row.OldBaseURL);
      return normalizedInput === base || normalizedInput.startsWith(`${base}/`);
    })
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
    if (/_layouts\/15\/doc\.aspx$/i.test(path) && params.get('sourcedoc')) {
      const guid = normalizeDocumentGuid(params.get('sourcedoc') || '');
      if (/^[0-9a-f]{32}$/.test(guid)) return `guid|${guid}`;
    }
    if (params.has('id') || params.has('parent')) {
      return `pdf|${path}|${decodeUrlPart(params.get('id') || '')}|${decodeUrlPart(params.get('parent') || '')}`;
    }
    if (params.has('d')) {
      const guid = normalizeDocumentGuid(params.get('d') || '');
      if (/^[0-9a-f]{32}$/.test(guid)) return `guid|${guid}`;
    }
    return `path|${path.replace(/^\/:\w:\/r/i, '')}`;
  } catch {
    return `raw|${url.trim().toLowerCase()}`;
  }
}

function getOldFileRefSignature(url: string): string {
  try {
    const parsed = new URL(url);
    let path = decodeUrlPart(parsed.searchParams.get('id') || parsed.pathname).replace(/\/$/, '');
    path = path.replace(/^\/:\w:\/r/i, '');
    if (/_layouts\/15\/doc\.aspx$/i.test(path) || /\/forms\/allitems\.aspx$/i.test(path)) return '';
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
  if (!response.ok) {
    const failure = `${context} failed (${response.status}).`;
    let errorBody: { error?: { message?: string | { value?: string } } } | null;
    try {
      errorBody = JSON.parse(text);
    } catch {
      throw new Error(`${failure} ${text.slice(0, 500)}`.trim());
    }
    const message = errorBody?.error?.message;
    const detail = typeof message === 'string' ? message : message?.value;
    throw new Error(detail ? `${failure} ${detail}` : failure);
  }
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
      const lookupWebUrl = getAbsoluteServerRelativeUrl(this.siteAbsoluteUrl, LOOKUP_WEB_PATH);
      const endpoint = `${lookupWebUrl}/_api/web/GetFolderByServerRelativePath(decodedurl='${LOOKUP_FOLDER}')/Files?$select=Name,ServerRelativeUrl&$top=5000`;
      this.lookupFilesPromise = this.spHttpClient
        .get(endpoint, SPHttpClient.configurations.v1, { headers: { Accept: 'application/json;odata.metadata=none' } })
        .then((response: SPHttpClientResponse) => readJsonResponse<{ value?: ILookupFile[] }>(response, 'Lookup table discovery'))
        .then((data: { value?: ILookupFile[] }) => {
          if (!Array.isArray(data?.value)) throw new Error('Lookup table discovery returned an invalid file list.');
          if (data.value.some((file) => !file || typeof file.Name !== 'string' ||
              typeof file.ServerRelativeUrl !== 'string' || !file.ServerRelativeUrl.startsWith(`${LOOKUP_FOLDER}/`))) {
            throw new Error('Lookup table discovery returned an invalid file entry.');
          }
          return data.value.filter((file) => /\.csv$/i.test(file.Name));
        })
        .catch((error: Error) => {
          this.lookupFilesPromise = undefined;
          throw error;
        });
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
        .then((text: string) => parseLookupCsv(text, INDEX_FILE_NAME))
        .catch((error: Error) => {
          this.lookupIndexPromise = undefined;
          throw error;
        });
    }
    return this.lookupIndexPromise;
  }

  public async findLookupRow(url: string): Promise<ILookupMatch | undefined> {
    const indexRows = await this.getLookupIndex();
    const indexRow = findIndexRow(url, indexRows);
    if (!indexRow) return undefined;

    const files = await this.getLookupFiles();
    const fixerFile = files.find((file) => file.Name.toLowerCase() === indexRow.FixerCsvFileName.toLowerCase());
    if (!fixerFile) throw new Error(`Could not find ${indexRow.FixerCsvFileName} in the URL Lookup Tables library.`);
    const inputSignature = getUrlSignature(url);
    const inputFileRef = getOldFileRefSignature(url);

    let rows = this.lookupRowsByFile.get(fixerFile.Name);
    if (!rows) {
      const response = await this.spHttpClient.get(
        getAbsoluteServerRelativeUrl(this.siteAbsoluteUrl, fixerFile.ServerRelativeUrl),
        SPHttpClient.configurations.v1
      );
      if (!response.ok) throw new Error(`Could not load ${fixerFile.Name} (${response.status}).`);
      rows = parseLookupCsv(await response.text(), fixerFile.Name);
      this.lookupRowsByFile.set(fixerFile.Name, rows);
    }

    const oldUrls = (candidate: ILookupRow): string[] => [candidate.OldCopyLinkUrl, candidate.OldResolvedUrl].filter(Boolean);
    let matches = rows.filter((candidate) => oldUrls(candidate).some((value) => value === url.trim()));
    if (!matches.length && inputFileRef) {
      matches = rows.filter((candidate) => {
        const fileRef = candidate.OldFileRef ? decodeUrlPart(candidate.OldFileRef).replace(/\/$/, '') : '';
        return fileRef === inputFileRef || oldUrls(candidate).some((value) => getOldFileRefSignature(value) === inputFileRef);
      });
    }
    if (!matches.length) {
      matches = rows.filter((candidate) => oldUrls(candidate).some((value) => getUrlSignature(value) === inputSignature));
    }
    const outcomes = new Set(matches.map((candidate) => JSON.stringify([
      candidate.NewFound.toLowerCase(), candidate.NewCopyLinkUrl || '', candidate.NewResolvedUrl || ''
    ])));
    if (outcomes.size > 1) {
      throw new Error(`Conflicting mappings in ${fixerFile.Name}. Multiple rows match this link but disagree on the migrated destination or status. Please contact the Help Desk to review the lookup table.`);
    }
    const row = matches[0];

    if (row && row.NewFound && row.NewFound.trim().toLowerCase() === 'no') {
      return { ...row, Notes: HELP_DESK_MESSAGE, lookupTableFileName: fixerFile.Name };
    }

    return row ? { ...row, lookupTableFileName: fixerFile.Name } : undefined;
  }
}

interface IDestinationCheck {
  reachable: boolean;
  reason?: string;
}

async function checkSharePointUrl(url: string): Promise<IDestinationCheck> {
  try {
    const parsed = new URL(url);
    const path = parsed.pathname.replace(/\/$/, '');
    const failure = (response: Response): IDestinationCheck | undefined => {
      if (response.url) {
        const finalUrl = new URL(response.url);
        if (finalUrl.hostname.toLowerCase() !== parsed.hostname.toLowerCase() ||
            /\/(?:accessdenied|error|login|signin)\.aspx$/i.test(finalUrl.pathname)) {
          return { reachable: false, reason: 'SharePoint redirected to a sign-in, access-denied, or error page.' };
        }
      }
      if (!response.ok) {
        return { reachable: false, reason: response.status === 401 || response.status === 403
          ? `SharePoint denied access (${response.status}).`
          : `SharePoint could not verify the destination (HTTP ${response.status}).` };
      }
      return undefined;
    };

    // A SharePoint site route can return a generic shell page even when the site
    // does not exist. Ask that site directly through REST before accepting it.
    if (/^\/sites\/[^/]+$/i.test(path)) {
      const siteResponse = await fetch(`${parsed.origin}${path}/_api/web?$select=Url`, {
        credentials: 'same-origin',
        headers: { Accept: 'application/json;odata=nometadata' },
        cache: 'no-store'
      });
      const siteFailure = failure(siteResponse);
      if (siteFailure) return siteFailure;
      const siteText = await siteResponse.text();
      try {
        const siteData: { Url?: string } | null = JSON.parse(siteText);
        if (typeof siteData?.Url !== 'string' || normalizeUrlBase(siteData.Url) !== normalizeUrlBase(url)) {
          return { reachable: false, reason: 'SharePoint did not return the expected site URL.' };
        }
        return { reachable: true };
      } catch {
        return { reachable: false, reason: 'SharePoint returned an invalid site response instead of site data.' };
      }
    }

    const response = await fetch(url, {
      method: 'HEAD',
      credentials: 'same-origin',
      redirect: 'follow',
      cache: 'no-store'
    });
    return failure(response) || { reachable: true };
  } catch (error) {
    return { reachable: false, reason: `The destination check failed: ${error instanceof Error ? error.message : 'Unknown network error.'}` };
  }
}

function helpDeskResult(input: string): IFixResult {
  return { original: input, fixed: input, note: `err:${HELP_DESK_MESSAGE}` };
}

// Tries the fast local pattern fix first, validates it's actually reachable, and
// only falls back to the (slower, network-backed) CSV lookup table when the
// pattern fix is unresolvable or the resulting URL doesn't check out.
export async function lookupOrFixUrl(raw: string, lookupClient: LookupTableClient): Promise<IFixResult> {
  const input = raw.trim();

  const patternResult = fixUrl(input);
  const patternLevel = patternResult.note.slice(0, patternResult.note.indexOf(':'));
  if (!WHA_SHAREPOINT_HOST_PATTERN.test(input)) {
    return patternResult;
  }
  if (patternLevel !== 'err') {
    const patternCheck = await checkSharePointUrl(patternResult.fixed);
    if (patternCheck.reachable) return patternResult;
  }

  const lookupRow = await lookupClient.findLookupRow(input);
  if (!lookupRow) {
    return helpDeskResult(input);
  }

  if (lookupRow.NewFound?.toLowerCase() === 'no' ||
      (lookupRow.Notes && !lookupRow.NewCopyLinkUrl && !lookupRow.NewResolvedUrl)) {
    return {
      original: input,
      fixed: input,
      lookupTableFileName: lookupRow.lookupTableFileName,
      note: `err:${lookupRow.Notes || HELP_DESK_MESSAGE}`
    };
  }

  const isResolved = /_layouts\/15\/doc\.aspx|[?&](?:id|parent)=/i.test(input);
  const lookupUrl = isResolved
    ? (lookupRow.NewResolvedUrl || lookupRow.NewCopyLinkUrl)
    : (lookupRow.NewCopyLinkUrl || lookupRow.NewResolvedUrl);

  if (!lookupUrl) {
    throw new Error('The matching lookup row has no migrated destination URL.');
  }
  const lookupCheck = await checkSharePointUrl(lookupUrl);
  if (!lookupCheck.reachable) {
    return {
      original: input,
      fixed: lookupUrl,
      lookupTableFileName: lookupRow.lookupTableFileName,
      note: `warn:A migrated URL was found in the lookup table, but its availability could not be verified. ${lookupCheck.reason} You can try opening the mapped link; contact the Help Desk if you cannot access it.`
    };
  }

  return {
    original: input,
    fixed: lookupUrl,
    lookupTableFileName: lookupRow.lookupTableFileName,
    note: 'ok:Pattern match was unavailable or did not resolve; found the migrated document in the SharePoint lookup table.'
  };
}
