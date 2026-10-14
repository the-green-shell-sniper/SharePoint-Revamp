export interface IFavoriteTitle {
  title: string;
  notice?: string;
}

interface ISharePointListPath {
  listPath: string;
  webPaths: string[];
}

const GENERIC_VIEW_PAGES = /^(?:All ?Items|Standard|DispForm|EditForm|NewForm)\.aspx$/i;

function destinationPath(url: URL): string {
  const embeddedPath = url.searchParams.get('id') || url.searchParams.get('sourcedoc');
  if (embeddedPath?.startsWith('/')) return embeddedPath;
  return url.pathname.replace(/%[0-9a-f]{2}(?:%[0-9a-f]{2})*/gi, (part) => {
    try { return decodeURIComponent(part); }
    catch { return part; }
  }).replace(/^\/:[a-z]:(?:\/[a-z]+)?\//i, '/');
}

function sharePointListPath(path: string): ISharePointListPath | undefined {
  const segments = path.split('/').filter(Boolean);
  if (!segments.length || !GENERIC_VIEW_PAGES.test(segments[segments.length - 1])) return undefined;

  const listsIndex = segments.findIndex((segment) => segment.toLowerCase() === 'lists');
  if (listsIndex >= 0 && listsIndex + 1 < segments.length - 1) {
    return {
      listPath: '/' + segments.slice(0, listsIndex + 2).join('/'),
      webPaths: ['/' + segments.slice(0, listsIndex).join('/')]
    };
  }

  const formsIndex = segments.length - 2;
  if (formsIndex <= 0 || segments[formsIndex].toLowerCase() !== 'forms') return undefined;

  const listSegments = segments.slice(0, formsIndex);
  const siteDepth = /^(?:sites|teams)$/i.test(listSegments[0]) ? 2 : 0;
  if (listSegments.length <= siteDepth) return undefined;

  const webPaths: string[] = [];
  for (let depth = listSegments.length - 1; depth >= siteDepth; depth--) {
    const webPath = '/' + listSegments.slice(0, depth).join('/');
    if (webPaths.indexOf(webPath) < 0) webPaths.push(webPath);
  }

  return {
    listPath: '/' + listSegments.join('/'),
    webPaths
  };
}

function readableName(path: string): string {
  const segments = path.split('/').filter(Boolean);
  const pageName = segments[segments.length - 1] || '';
  if (GENERIC_VIEW_PAGES.test(pageName)) {
    const listPath = sharePointListPath(path)?.listPath;
    if (listPath) return listPath.split('/').filter(Boolean).pop()!.replace(/[-_]/g, ' ');
    if (segments.length > 1 && segments[segments.length - 2].toLowerCase() === 'forms') {
      return segments[segments.length - 3] || 'WHA SharePoint';
    }
  }

  const name = pageName || 'WHA SharePoint';
  if (/\.aspx$/i.test(name)) return name.replace(/\.aspx$/i, '').replace(/[-_]/g, ' ');
  return name;
}

function encodeODataPath(path: string): string {
  return encodeURIComponent(path.replace(/'/g, "''")).replace(/%2F/gi, '/');
}

export async function getFavoriteTitle(value: string): Promise<IFavoriteTitle> {
  const url = new URL(/^https?:\/\//i.test(value) ? value : `https://${value}`);
  const path = destinationPath(url);
  const fallback = readableName(url.searchParams.get('file') || path);
  const listPath = sharePointListPath(path);
  const sitePath = path.match(/^\/(?:sites|teams)\/[^/]+/i)?.[0] || '';
  const siteUrl = `${url.origin}${sitePath}`;
  const guid = url.searchParams.get('sourcedoc')?.replace(/[{}]/g, '');
  const isGuidLink = /\/_layouts\/15\/doc\.aspx$/i.test(path) && !!guid && /^[0-9a-f-]{36}$/i.test(guid);
  const isSite = path.replace(/\/$/, '') === sitePath;
  const escapedPath = encodeURIComponent(path.replace(/'/g, "''")).replace(/'/g, '%27');
  const endpoint = isSite
    ? `${siteUrl}/_api/web?$select=Title`
    : isGuidLink
      ? `${siteUrl}/_api/web/GetFileById('${guid}')?$select=Name,ListItemAllFields/Title&$expand=ListItemAllFields`
      : `${siteUrl}/_api/web/GetFileByServerRelativePath(decodedurl='${escapedPath}')?$select=Name,ListItemAllFields/Title&$expand=ListItemAllFields`;
  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), 8000);

  try {
    if (url.hostname.toLowerCase() !== 'whallc.sharepoint.com' || !/^https?:$/.test(url.protocol)) {
      throw new Error('Not a WHA SharePoint URL.');
    }
    const requestOptions: RequestInit = {
      credentials: 'same-origin',
      headers: { Accept: 'application/json;odata=nometadata' },
      signal: controller.signal
    };
    let response: Response;
    if (listPath) {
      let lastStatus = 404;
      let title: string | undefined;
      for (const webPath of listPath.webPaths) {
        const apiPath = webPath === '/' ? '/_api/web' : `${webPath}/_api/web`;
        const listEndpoint = `${url.origin}${apiPath}/GetList('${encodeODataPath(listPath.listPath)}')?$select=Title`;
        response = await fetch(listEndpoint, requestOptions);
        if (!response.ok) {
          lastStatus = response.status;
          if ([400, 403, 404].indexOf(response.status) >= 0) continue;
          throw new Error(`SharePoint list title request failed (${response.status}).`);
        }
        const data: { Title?: unknown } = await response.json();
        if (typeof data.Title === 'string' && data.Title.trim()) {
          title = data.Title.trim();
          break;
        }
      }
      if (title) return { title };
      throw new Error(`SharePoint did not return a list or library title (${lastStatus}).`);
    }

    response = await fetch(endpoint, requestOptions);
    if (!response.ok) throw new Error(`SharePoint title request failed (${response.status}).`);
    const data: { Title?: unknown; Name?: unknown; ListItemAllFields?: { Title?: unknown } } = await response.json();
    const title = isSite ? data.Title : data.ListItemAllFields?.Title;
    if (typeof title === 'string' && title.trim()) return { title: title.trim() };
    if (typeof data.Name === 'string' && data.Name.trim()) return { title: readableName(data.Name.trim()) };
    throw new Error('SharePoint did not return a title or filename.');
  } catch (error) {
    const reason = error instanceof Error ? error.message : 'Unknown title lookup error.';
    console.warn(`Could not retrieve the SharePoint favorite title: ${reason}`);
    return { title: fallback, notice: 'Could not retrieve the SharePoint title. The favorite name is based on the URL; you can rename it in Edge.' };
  } finally {
    clearTimeout(timeout);
  }
}
