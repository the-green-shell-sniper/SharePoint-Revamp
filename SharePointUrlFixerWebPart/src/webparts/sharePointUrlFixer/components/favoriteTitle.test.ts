import { getFavoriteTitle } from './favoriteTitle';

describe('SharePoint favorite names', () => {
  const originalFetch = globalThis.fetch;
  const fetchMock = jest.fn();
  const fileUrl = 'https://whallc.sharepoint.com/:w:/r/sites/HR/Shared%20Documents/Guide.docx';

  beforeEach(() => {
    globalThis.fetch = fetchMock;
    fetchMock.mockReset();
    jest.spyOn(console, 'warn').mockImplementation(() => undefined);
  });

  afterEach(() => {
    globalThis.fetch = originalFetch;
    jest.restoreAllMocks();
    jest.useRealTimers();
  });

  function respond(data: object): void {
    fetchMock.mockResolvedValue({ ok: true, json: async () => data });
  }

  function respondAt(endpointPart: string, data: object): void {
    fetchMock.mockImplementation(async (endpoint: string) => endpoint.includes(endpointPart)
      ? { ok: true, json: async () => data }
      : { ok: false, status: 404 });
  }

  it('retrieves the exact subsite title, not the parent site title', async () => {
    respondAt('/sites/HR/MSresource/_api/web?$select=Title,ServerRelativeUrl', {
      Title: 'Medical Staff Resources', ServerRelativeUrl: '/sites/HR/MSresource'
    });
    expect(await getFavoriteTitle('https://whallc.sharepoint.com/sites/HR/MSresource/'))
      .toEqual({ title: 'Medical Staff Resources' });
  });

  it.each([
    ['/sites/HR/Shared Documents', 'Staff Documents'],
    ['/sites/HR/MSresource/Lists/Announcements', 'Staff Announcements']
  ])('names a direct library or list URL %s from that object', async (path, title) => {
    respondAt(`/GetList('${path.replace(/ /g, '%20')}')`, {
      Title: title, RootFolder: { ServerRelativeUrl: path }
    });
    expect(await getFavoriteTitle(`https://whallc.sharepoint.com${path.replace(/ /g, '%20')}`))
      .toEqual({ title });
  });

  it.each(['', '?RootFolder=%2Fsites%2FHR%2FDocuments%2F2026.Notes', '?id=%2Fsites%2FHR%2FDocuments%2F2026.Notes'])(
    'uses the folder name rather than its containing library for %s', async (query) => {
      const path = '/sites/HR/Documents/2026.Notes';
      fetchMock.mockImplementation(async (endpoint: string) => {
        if (endpoint.includes('/GetList(')) return {
          ok: true, json: async () => ({ Title: 'Staff Documents', RootFolder: { ServerRelativeUrl: '/sites/HR/Documents' } })
        };
        if (endpoint.includes('/GetFolderByServerRelativePath(')) return {
          ok: true, json: async () => ({ Name: '2026.Notes', ServerRelativeUrl: path })
        };
        return { ok: false, status: 404 };
      });
      const url = query ? `https://whallc.sharepoint.com/sites/HR/Documents/Forms/AllItems.aspx${query}` : `https://whallc.sharepoint.com${path}`;
      expect(await getFavoriteTitle(url)).toEqual({ title: '2026.Notes' });
    }
  );

  it('names a custom library view from its library, not its .aspx view filename', async () => {
    respond({ Title: 'Staff Documents' });
    expect(await getFavoriteTitle('https://whallc.sharepoint.com/sites/HR/Documents/Forms/ByDepartment.aspx'))
      .toEqual({ title: 'Staff Documents' });
  });

  it.each(['DispForm', 'EditForm'])('names a list item opened with %s from its own Title', async (form) => {
    respond({ Title: 'Annual staff meeting' });
    expect(await getFavoriteTitle(`https://whallc.sharepoint.com/sites/HR/Lists/Announcements/${form}.aspx?ID=42`))
      .toEqual({ title: 'Annual staff meeting' });
    expect(fetchMock.mock.calls[0][0]).toContain("/GetList('/sites/HR/Lists/Announcements')/items(42)?$select=Title");
  });

  it('identifies an untitled list item in its fallback rather than naming only the list', async () => {
    respond({});
    const result = await getFavoriteTitle('https://whallc.sharepoint.com/sites/HR/Lists/Announcements/DispForm.aspx?ID=42');
    expect(result).toEqual({ title: 'Announcements - item 42', notice: expect.any(String) });
  });

  it('uses the filename for a library document opened through its item form', async () => {
    respond({ Title: 'WHA Document Template', File: { Name: 'Release of Information.docx' } });
    expect(await getFavoriteTitle('https://whallc.sharepoint.com/sites/HR/Documents/Forms/DispForm.aspx?ID=42'))
      .toEqual({ title: 'Release of Information.docx' });
  });

  it('resolves documents in subsites through their owning web', async () => {
    respondAt("/sites/HR/MSresource/_api/web/GetFileByServerRelativePath(", { Name: 'Staff Guide.docx' });
    expect(await getFavoriteTitle('https://whallc.sharepoint.com/sites/HR/MSresource/Documents/Staff%20Guide.docx'))
      .toEqual({ title: 'Staff Guide.docx' });
  });

  it('resolves GUID links using the subsite identified by the URL', async () => {
    respond({ Name: 'Staff Guide.docx' });
    expect(await getFavoriteTitle('https://whallc.sharepoint.com/sites/HR/MSresource/_layouts/15/Doc.aspx?sourcedoc=%7B12345678-1234-1234-1234-123456789abc%7D'))
      .toEqual({ title: 'Staff Guide.docx' });
    expect(fetchMock.mock.calls[0][0]).toContain('/sites/HR/MSresource/_api/web/GetFileById(');
  });

  it('resolves page titles in subsites without naming the parent site', async () => {
    respondAt('/sites/HR/MSresource/_api/web/GetFileByServerRelativePath(', {
      Name: 'Staff.aspx', ListItemAllFields: { Title: 'Medical Staff Contacts' }
    });
    expect(await getFavoriteTitle('https://whallc.sharepoint.com/sites/HR/MSresource/SitePages/Staff.aspx'))
      .toEqual({ title: 'Medical Staff Contacts' });
  });

  it('does not treat an ancestor web response as the linked object', async () => {
    fetchMock.mockImplementation(async (endpoint: string) => endpoint.includes('?$select=Title,ServerRelativeUrl')
      ? { ok: true, json: async () => ({ Title: 'Human Resources', ServerRelativeUrl: '/sites/HR' }) }
      : { ok: false, status: 404 });
    expect(await getFavoriteTitle('https://whallc.sharepoint.com/sites/HR/MissingFolder'))
      .toEqual({ title: 'MissingFolder', notice: expect.any(String) });
  });

  it('uses the actual document filename rather than its Title metadata', async () => {
    respond({ Name: 'Guide.docx', ListItemAllFields: { Title: 'Staff Guide' } });
    expect(await getFavoriteTitle(fileUrl)).toEqual({ title: 'Guide.docx' });
    expect(fetchMock).toHaveBeenCalledWith(
      "https://whallc.sharepoint.com/sites/HR/_api/web/GetFileByServerRelativePath(decodedurl='%2Fsites%2FHR%2FShared%20Documents%2FGuide.docx')?$select=Name",
      expect.objectContaining({ credentials: 'same-origin', signal: expect.anything() })
    );
  });

  it('uses the filename when a document has no Title field', async () => {
    respond({ Name: 'Guide.docx', ListItemAllFields: { Title: null } });
    expect(await getFavoriteTitle(fileUrl)).toEqual({ title: 'Guide.docx' });
  });

  it('retrieves a modern page title rather than its .aspx filename', async () => {
    respond({ Name: 'Employee-Benefits.aspx', ListItemAllFields: { Title: 'Employee Benefits at WHA' } });
    expect(await getFavoriteTitle('https://whallc.sharepoint.com/sites/HR/SitePages/Employee-Benefits.aspx'))
      .toEqual({ title: 'Employee Benefits at WHA' });
  });

  it('uses the SharePoint list title instead of a generic Standard.aspx page name', async () => {
    respond({ Title: 'Status Notification Form' });
    const url = 'https://whallc.sharepoint.com/sites/HR/MSresource/Lists/Status%20Notification%20Form/Standard.aspx?view=compact';
    expect(await getFavoriteTitle(url)).toEqual({ title: 'Status Notification Form' });
    expect(fetchMock.mock.calls[0][0]).toBe(
      "https://whallc.sharepoint.com/sites/HR/MSresource/_api/web/GetList('/sites/HR/MSresource/Lists/Status%20Notification%20Form')?$select=Title"
    );
  });

  it('uses the SharePoint library title instead of a generic AllItems.aspx page name', async () => {
    respond({ Title: 'Forms' });
    const url = 'https://whallc.sharepoint.com/sites/Epic/FormsandConsents/Forms/AllItems.aspx?viewid=123';
    expect(await getFavoriteTitle(url)).toEqual({ title: 'Forms' });
    expect(fetchMock.mock.calls[0][0]).toBe(
      "https://whallc.sharepoint.com/sites/Epic/_api/web/GetList('/sites/Epic/FormsandConsents')?$select=Title"
    );
  });

  it('uses the list name instead of a generic page name if the list title lookup fails', async () => {
    fetchMock.mockResolvedValue({ ok: false, status: 404 });
    const result = await getFavoriteTitle(
      'https://whallc.sharepoint.com/sites/HR/MSresource/Lists/Status%20Notification%20Form/Standard.aspx'
    );
    expect(result.title).toBe('Status Notification Form');
    expect(result.notice).toContain('Could not retrieve');
  });

  it.each(['https://whallc.sharepoint.com/sites/HR/', 'whallc.sharepoint.com/sites/HR', 'https://whallc.sharepoint.com/'])(
    'retrieves a site title for %s', async (url) => {
      respond({ Title: 'Human Resources' });
      expect(await getFavoriteTitle(url)).toEqual({ title: 'Human Resources' });
      expect(fetchMock.mock.calls[0][0]).toMatch(/\/_api\/web\?\$select=Title$/);
    }
  );

  it('uses an embedded file path for library-view links', async () => {
    respond({ Name: 'Guide.pdf' });
    await getFavoriteTitle('https://whallc.sharepoint.com/sites/HR/Documents/Forms/AllItems.aspx?id=%2Fsites%2FHR%2FDocuments%2FGuide.pdf');
    expect(fetchMock.mock.calls[0][0]).toContain("decodedurl='%2Fsites%2FHR%2FDocuments%2FGuide.pdf'");
  });

  it('looks up GUID-based document links by file ID', async () => {
    respond({ Name: 'Guide.docx' });
    const id = '12345678-1234-1234-1234-123456789abc';
    await getFavoriteTitle(`https://whallc.sharepoint.com/sites/HR/_layouts/15/Doc.aspx?sourcedoc=%7B${id}%7D`);
    expect(fetchMock.mock.calls[0][0]).toContain(`GetFileById('${id}')`);
  });

  it('names the reported Release of Information GUID link from its filename, not template metadata', async () => {
    respond({ Name: 'Release of Information.docx', ListItemAllFields: { Title: 'WHA Document Template' } });
    const url = 'https://whallc.sharepoint.com/:w:/r/sites/ClinicalandPatientServices/_layouts/15/Doc.aspx?sourcedoc=%7B342C713C-06A6-4844-8132-99503089E0A6%7D&file=Release%20of%20Information.docx&action=default&mobileredirect=true';
    expect(await getFavoriteTitle(url)).toEqual({ title: 'Release of Information.docx' });
    expect(fetchMock.mock.calls[0][0]).toBe(
      "https://whallc.sharepoint.com/sites/ClinicalandPatientServices/_api/web/GetFileById('342C713C-06A6-4844-8132-99503089E0A6')?$select=Name"
    );
  });

  it('prefers the current filename over a stale file query parameter', async () => {
    respond({ Name: 'Renamed Release of Information.docx' });
    expect(await getFavoriteTitle(
      'https://whallc.sharepoint.com/sites/HR/_layouts/15/Doc.aspx?sourcedoc=%7B12345678-1234-1234-1234-123456789abc%7D&file=OldName.docx'
    )).toEqual({ title: 'Renamed Release of Information.docx' });
  });

  it('ignores template metadata for embedded document paths', async () => {
    respond({ Name: 'Release of Information.pdf', ListItemAllFields: { Title: 'WHA Document Template' } });
    expect(await getFavoriteTitle(
      'https://whallc.sharepoint.com/sites/HR/Documents/Forms/AllItems.aspx?id=%2Fsites%2FHR%2FDocuments%2FRelease%20of%20Information.pdf'
    )).toEqual({ title: 'Release of Information.pdf' });
  });

  it('falls back to the file parameter instead of Doc.aspx when a GUID lookup fails', async () => {
    fetchMock.mockRejectedValue(new Error('Network unavailable'));
    expect(await getFavoriteTitle(
      'https://whallc.sharepoint.com/sites/HR/_layouts/15/Doc.aspx?sourcedoc=%7B12345678-1234-1234-1234-123456789abc%7D&file=Release%20of%20Information.docx'
    )).toEqual({ title: 'Release of Information.docx', notice: expect.any(String) });
  });

  it('escapes apostrophes, percent signs and hashes in file paths', async () => {
    respond({ Name: "Staff's #1 100%.docx" });
    await getFavoriteTitle("https://whallc.sharepoint.com/sites/HR/Documents/Staff%27s%20%231%20100%25.docx");
    expect(fetchMock.mock.calls[0][0]).toContain('Staff%27%27s%20%231%20100%25.docx');
  });

  it.each([403, 404, 500])('keeps a readable page name and explains an HTTP %s failure', async (status) => {
    fetchMock.mockResolvedValue({ ok: false, status });
    const result = await getFavoriteTitle('https://whallc.sharepoint.com/sites/HR/SitePages/Employee-Benefits.aspx');
    expect(result.title).toBe('Employee Benefits');
    expect(result.notice).toContain('Could not retrieve');
    expect(console.warn).toHaveBeenCalled();
  });

  it('reports network failures without losing the filename', async () => {
    fetchMock.mockRejectedValue(new Error('Network unavailable'));
    expect(await getFavoriteTitle(fileUrl)).toEqual({ title: 'Guide.docx', notice: expect.any(String) });
  });

  it.each([{}, { Title: 123 }, { ListItemAllFields: { Title: ' ' } }])('handles missing or invalid metadata: %s', async (data) => {
    respond(data);
    expect((await getFavoriteTitle(fileUrl)).notice).toBeDefined();
  });

  it('aborts a title request after eight seconds and uses the URL name', async () => {
    jest.useFakeTimers();
    fetchMock.mockImplementation((_url: string, options: { signal: AbortSignal }) => new Promise((_resolve, reject) => {
      options.signal.addEventListener('abort', () => reject(new Error('Title request timed out')));
    }));
    const pending = getFavoriteTitle(fileUrl);
    jest.advanceTimersByTime(8000);
    expect(await pending).toEqual({ title: 'Guide.docx', notice: expect.any(String) });
  });

  it('does not send requests outside WHA SharePoint', async () => {
    expect((await getFavoriteTitle('https://example.com/Guide.docx')).notice).toBeDefined();
    expect(fetchMock).not.toHaveBeenCalled();
  });
});
