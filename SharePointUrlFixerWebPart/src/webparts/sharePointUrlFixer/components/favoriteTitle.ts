export interface IFavoriteTitle {
  title: string;
  notice?: string;
}

function destinationPath(url: URL): string {
  const embeddedPath = url.searchParams.get('id') || url.searchParams.get('sourcedoc');
  if (embeddedPath?.startsWith('/')) return embeddedPath;
  return url.pathname.replace(/%[0-9a-f]{2}(?:%[0-9a-f]{2})*/gi, (part) => {
    try { return decodeURIComponent(part); }
    catch { return part; }
  }).replace(/^\/:[a-z]:(?:\/[a-z]+)?\//i, '/');
}

function readableName(path: string): string {
  const name = path.split('/').filter(Boolean).pop() || 'WHA SharePoint';
  if (/\.aspx$/i.test(name)) return name.replace(/\.aspx$/i, '').replace(/[-_]/g, ' ');
  return name;
}

export async function getFavoriteTitle(value: string): Promise<IFavoriteTitle> {
  const url = new URL(/^https?:\/\//i.test(value) ? value : `https://${value}`);
  const path = destinationPath(url);
  const fallback = readableName(url.searchParams.get('file') || path);
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
    const response = await fetch(endpoint, {
      credentials: 'same-origin',
      headers: { Accept: 'application/json;odata=nometadata' },
      signal: controller.signal
    });
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
