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
  const embeddedPath = url.searchParams.get('id') || url.searchParams.get('sourcedoc') || url.searchParams.get('RootFolder');
  if (embeddedPath?.startsWith('/')) return embeddedPath;
  return url.pathname.replace(/%[0-9a-f]{2}(?:%[0-9a-f]{2})*/gi, (part) => {
    try { return decodeURIComponent(part); }
    catch { return part; }
  }).replace(/^\/:[a-z]:(?:\/[a-z]+)?\//i, '/');
}

function candidateWebPaths(path: string): string[] {
  const segments = path.split('/').filter(Boolean);
  const siteDepth = /^(?:sites|teams)$/i.test(segments[0] || '') ? 2 : 0;
  const paths: string[] = [];
  for (let depth = segments.length - 1; depth >= siteDepth; depth--) {
    paths.push('/' + segments.slice(0, depth).join('/'));
  }
  return paths;
}

function sharePointListPath(path: string): ISharePointListPath | undefined {
  const segments = path.split('/').filter(Boolean);
  if (!segments.length || !/\.aspx$/i.test(segments[segments.length - 1])) return undefined;

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

  return {
    listPath: '/' + listSegments.join('/'),
    webPaths: candidateWebPaths('/' + listSegments.join('/'))
  };
}

function readableName(path: string): string {
  const segments = path.split('/').filter(Boolean);
  const pageName = segments[segments.length - 1] || '';
  const listPath = sharePointListPath(path)?.listPath;
  if (listPath) return listPath.split('/').filter(Boolean).pop()!.replace(/[-_]/g, ' ');
  if (GENERIC_VIEW_PAGES.test(pageName)) {
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
  const itemId = url.searchParams.get('ID') || url.searchParams.get('Id') || url.searchParams.get('id');
  const isListItem = !!listPath && /\/(?:DispForm|EditForm)\.aspx$/i.test(path) && !!itemId && /^[1-9]\d*$/.test(itemId);
  const sitePath = path.match(/^\/(?:sites|teams)\/[^/]+/i)?.[0] || '';
  const guidWebPath = path.split(/\/_layouts\/15\/doc\.aspx/i)[0];
  const guid = url.searchParams.get('sourcedoc')?.replace(/[{}]/g, '');
  const isGuidLink = /\/_layouts\/15\/doc\.aspx$/i.test(path) && !!guid && /^[0-9a-f-]{36}$/i.test(guid);
  const isSite = path.replace(/\/$/, '') === sitePath;
  const isPage = !isGuidLink && /\.aspx$/i.test(path);
  const siteUrl = `${url.origin}${isGuidLink ? guidWebPath : sitePath}`;
  const escapedPath = encodeURIComponent(path.replace(/'/g, "''")).replace(/'/g, '%27');
  const endpoint = isSite
    ? `${siteUrl}/_api/web?$select=Title`
    : isGuidLink
      ? `${siteUrl}/_api/web/GetFileById('${guid}')?$select=Name`
      : `${siteUrl}/_api/web/GetFileByServerRelativePath(decodedurl='${escapedPath}')?$select=${isPage ? 'Name,ListItemAllFields/Title&$expand=ListItemAllFields' : 'Name'}`;
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
    async function readMetadata(endpointUrl: string): Promise<{
      Title?: unknown; Name?: unknown; ServerRelativeUrl?: unknown;
      RootFolder?: { ServerRelativeUrl?: unknown }; ListItemAllFields?: { Title?: unknown };
    } | undefined> {
      const metadataResponse = await fetch(endpointUrl, requestOptions);
      if (!metadataResponse.ok) {
        if ([400, 404].indexOf(metadataResponse.status) >= 0) return undefined;
        throw new Error(`SharePoint title request failed (${metadataResponse.status}).`);
      }
      return metadataResponse.json();
    }
    const apiUrl = (webPath: string): string => `${url.origin}${webPath === '/' ? '' : webPath}/_api/web`;
    let response: Response;
    if (listPath) {
      let lastStatus = 404;
      let title: string | undefined;
      for (const webPath of listPath.webPaths) {
        const apiPath = webPath === '/' ? '/_api/web' : `${webPath}/_api/web`;
        const listEndpoint = `${url.origin}${apiPath}/GetList('${encodeODataPath(listPath.listPath)}')${isListItem ? `/items(${itemId})?$select=Title,File/Name,Folder/Name&$expand=File,Folder` : '?$select=Title'}`;
        response = await fetch(listEndpoint, requestOptions);
        if (!response.ok) {
          lastStatus = response.status;
          if ([400, 404].indexOf(response.status) >= 0) continue;
          throw new Error(`SharePoint list title request failed (${response.status}).`);
        }
        const data: { Title?: unknown; File?: { Name?: unknown }; Folder?: { Name?: unknown } } = await response.json();
        const objectName = data.File?.Name || data.Folder?.Name;
        if (isListItem && typeof objectName === 'string' && objectName.trim()) {
          return { title: objectName.trim() };
        }
        if (typeof data.Title === 'string' && data.Title.trim()) {
          title = data.Title.trim();
          break;
        }
      }
      if (title) return { title };
      throw new Error(`SharePoint did not return a ${isListItem ? 'list item' : 'list or library'} title (${lastStatus}).`);
    }

    if (!isSite && !isGuidLink && !isPage) {
      const file = await readMetadata(endpoint);
      if (file) {
        if (typeof file.Name === 'string' && file.Name.trim()) return { title: file.Name.trim() };
        throw new Error('SharePoint did not return a filename.');
      }
      const objectPath = path.replace(/\/$/, '');
      const web = await readMetadata(`${apiUrl(encodeODataPath(objectPath))}?$select=Title,ServerRelativeUrl`);
      if (web?.ServerRelativeUrl === objectPath && typeof web.Title === 'string' && web.Title.trim()) {
        return { title: web.Title.trim() };
      }
      for (const webPath of candidateWebPaths(objectPath)) {
        if (webPath !== sitePath && webPath !== '/') {
          const nestedFile = await readMetadata(`${apiUrl(encodeODataPath(webPath))}/GetFileByServerRelativePath(decodedurl='${escapedPath}')?$select=Name`);
          if (typeof nestedFile?.Name === 'string' && nestedFile.Name.trim()) return { title: nestedFile.Name.trim() };
        }
        const list = await readMetadata(`${apiUrl(encodeODataPath(webPath))}/GetList('${encodeODataPath(objectPath)}')?$select=Title,RootFolder/ServerRelativeUrl&$expand=RootFolder`);
        if (list?.RootFolder?.ServerRelativeUrl === objectPath && typeof list.Title === 'string' && list.Title.trim()) {
          return { title: list.Title.trim() };
        }
        const folder = await readMetadata(`${apiUrl(encodeODataPath(webPath))}/GetFolderByServerRelativePath(decodedurl='${encodeODataPath(objectPath)}')?$select=Name,ServerRelativeUrl`);
        if (folder?.ServerRelativeUrl === objectPath && typeof folder.Name === 'string' && folder.Name.trim()) {
          return { title: folder.Name.trim() };
        }
      }
      throw new Error('SharePoint did not return a name for the linked site, library, list, or folder.');
    }

    let data = await readMetadata(endpoint);
    if (!data && isPage) {
      for (const webPath of candidateWebPaths(path)) {
        if (webPath === sitePath || webPath === '/') continue;
        data = await readMetadata(`${apiUrl(encodeODataPath(webPath))}/GetFileByServerRelativePath(decodedurl='${escapedPath}')?$select=Name,ListItemAllFields/Title&$expand=ListItemAllFields`);
        if (data) break;
      }
    }
    if (!data) throw new Error('SharePoint could not find the linked page or file.');
    const title = isSite ? data.Title : (isPage ? data.ListItemAllFields?.Title : undefined);
    if (typeof title === 'string' && title.trim()) return { title: title.trim() };
    if (typeof data.Name === 'string' && data.Name.trim()) {
      return { title: isPage ? readableName(data.Name.trim()) : data.Name.trim() };
    }
    throw new Error('SharePoint did not return a title or filename.');
  } catch (error) {
    const reason = error instanceof Error ? error.message : 'Unknown title lookup error.';
    console.warn(`Could not retrieve the SharePoint favorite title: ${reason}`);
    return {
      title: isListItem ? `${fallback} - item ${itemId}` : fallback,
      notice: 'Could not retrieve the SharePoint title. The favorite name is based on the URL; you can rename it in Edge.'
    };
  } finally {
    clearTimeout(timeout);
  }
}
