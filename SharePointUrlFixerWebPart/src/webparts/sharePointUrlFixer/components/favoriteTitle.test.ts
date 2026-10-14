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

  it('retrieves a document title using its decoded server-relative path', async () => {
    respond({ Name: 'Guide.docx', ListItemAllFields: { Title: 'Staff Guide' } });
    expect(await getFavoriteTitle(fileUrl)).toEqual({ title: 'Staff Guide' });
    expect(fetchMock).toHaveBeenCalledWith(
      "https://whallc.sharepoint.com/sites/HR/_api/web/GetFileByServerRelativePath(decodedurl='%2Fsites%2FHR%2FShared%20Documents%2FGuide.docx')?$select=Name,ListItemAllFields/Title&$expand=ListItemAllFields",
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
