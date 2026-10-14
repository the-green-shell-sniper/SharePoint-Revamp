jest.mock('@microsoft/sp-http', () => ({ SPHttpClient: Object.assign(jest.fn(), { configurations: { v1: {} } }) }));
jest.mock('@microsoft/sp-core-library', () => ({ ServiceScope: { startNewRoot: jest.fn() } }));

import { ServiceScope } from '@microsoft/sp-core-library';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { fixUrl, LookupTableClient, lookupOrFixUrl } from './urlFixerLogic';

describe('WHA URL validation', () => {
  const externalMessage = "err:You've provided an external link. This tool is only for WHA SharePoint links. Please provide a link that starts with 'https://whallc.sharepoint.com' to continue.";

  it.each([
    'https://example.com/CS/document.docx',
    'https://another.sharepoint.com/CS/document.docx',
    'https://whallc.sharepoint.com.example.com/CS/document.docx',
    'https://whallc.sharepoint.com@evil.example/CS/document.docx',
    'https://whallc.sharepoint.comevil/CS/document.docx',
    'not a URL'
  ])('rejects %s without accessing lookup tables', async (url) => {
    const client = new LookupTableClient(new SPHttpClient(ServiceScope.startNewRoot()), 'https://whallc.sharepoint.com');
    const lookup = jest.spyOn(client, 'findLookupRow').mockRejectedValue(new Error('Must not perform a lookup'));
    expect(fixUrl(url)).toEqual({ original: url, fixed: url, note: externalMessage });
    expect(await lookupOrFixUrl(url, client)).toEqual({ original: url, fixed: url, note: externalMessage });
    expect(lookup).not.toHaveBeenCalled();
  });

  describe('SharePoint lookup tables', () => {
    const siteUrl = 'https://whallc.sharepoint.com/sites/CentralParkLinkFinder';
    const folder = '/sites/CentralParkLinkFinder/URLLookupTables';
    const original = 'https://whallc.sharepoint.com/:w:/r/CS/_layouts/15/Doc.aspx?sourcedoc=%7BB5AA7687-49DC-4681-A326-B80253394172%7D&file=Release%20of%20Information.docx&action=default&mobileredirect=true';
    const resolved = 'https://whallc.sharepoint.com/:w:/r/sites/ClinicalandPatientServices/_layouts/15/Doc.aspx?sourcedoc=%7B342C713C-06A6-4844-8132-99503089E0A6%7D&file=Release%20of%20Information.docx&action=default&mobileredirect=true';
    const copyLink = 'https://whallc.sharepoint.com/:w:/r/sites/ClinicalandPatientServices/Patient%20Services%20formslettersconsents/Release%20of%20Information.docx?d=w342c713c06a64844813299503089e0a6&csf=1&web=1';
    const indexCsv = 'OldBaseURL,FixerCsvFileName,NewBaseURL\r\nhttps://whallc.sharepoint.com/CS,CS-fixer.csv,https://whallc.sharepoint.com/sites/ClinicalandPatientServices';
    const fixerCsv = [
      '"OldResolvedUrl","NewFound","NewCopyLinkUrl","NewResolvedUrl","Notes"',
      `"${original}","Yes","${copyLink}","${resolved}",""`
    ].join('\r\n');
    const legacyEmResolved = 'https://whallc.sharepoint.com/:w:/r/_layouts/15/Doc.aspx?sourcedoc=%7BEED3D553-4CE2-414B-92E3-0D646880F931%7D&file=FAQ_2021%20E%26M%20Coding.docx&action=default&mobileredirect=true';
    const migratedEmResolved = 'https://whallc.sharepoint.com/:w:/r/sites/PatientFinancialServices/_layouts/15/Doc.aspx?sourcedoc=%7B4EED5EFC-08B8-4D4D-94D5-E59FC28EF423%7D&file=FAQ_2021%20E%26M%20Coding.docx&action=default&mobileredirect=true';
    const files = ['URLFixerIndex.csv', 'CS-fixer.csv'].map((Name) => ({ Name, ServerRelativeUrl: `${folder}/${Name}` }));
    const originalFetch = globalThis.fetch;
    const fetchMock = jest.fn();

    let httpClient: SPHttpClient;
    let get: jest.SpyInstance<Promise<SPHttpClientResponse>>;
    let client: LookupTableClient;

    function response(text: string, status = 200): SPHttpClientResponse {
      return { ok: status >= 200 && status < 300, status, text: async () => text } as SPHttpClientResponse;
    }

    function queueTables(): void {
      get.mockResolvedValueOnce(response(JSON.stringify({ value: files })))
        .mockResolvedValueOnce(response(indexCsv))
        .mockResolvedValueOnce(response(fixerCsv));
    }

    function queueFixer(csv: string): void {
      get.mockResolvedValueOnce(response(JSON.stringify({ value: files })))
        .mockResolvedValueOnce(response(indexCsv))
        .mockResolvedValueOnce(response(csv));
    }

    beforeEach(() => {
      globalThis.fetch = fetchMock;
      fetchMock.mockReset();
      httpClient = new SPHttpClient(ServiceScope.startNewRoot());
      httpClient.get = jest.fn();
      get = jest.spyOn(httpClient, 'get');
      client = new LookupTableClient(httpClient, siteUrl);
    });

    afterEach(() => {
      globalThis.fetch = originalFetch;
      jest.restoreAllMocks();
    });

    it('uses an OData 4 Accept header and matches the exact CS Doc.aspx URL', async () => {
      queueTables();
      expect(await client.findLookupRow(original)).toMatchObject({
        OldResolvedUrl: original,
        NewResolvedUrl: resolved,
        NewCopyLinkUrl: copyLink,
        lookupTableFileName: 'CS-fixer.csv'
      });
      expect(get).toHaveBeenNthCalledWith(1,
        `${siteUrl}/_api/web/GetFolderByServerRelativePath(decodedurl='${folder}')/Files?$select=Name,ServerRelativeUrl&$top=5000`,
        SPHttpClient.configurations.v1,
        { headers: { Accept: 'application/json;odata.metadata=none' } }
      );
      expect(get).toHaveBeenNthCalledWith(3, `https://whallc.sharepoint.com${folder}/CS-fixer.csv`, SPHttpClient.configurations.v1);
    });

    it.each([
      'https://whallc.sharepoint.com',
      'https://whallc.sharepoint.com/sites/ClinicalandPatientServices',
      'https://whallc.sharepoint.com/sites/CentralParkLinkFinder/Subweb/'
    ])('discovers tables through their owning web when hosted at %s', async (hostingWeb) => {
      const hostedClient = new LookupTableClient(httpClient, hostingWeb);
      queueTables();
      expect(await hostedClient.findLookupRow(original)).toMatchObject({ NewResolvedUrl: resolved });
      expect(get).toHaveBeenNthCalledWith(1,
        `${siteUrl}/_api/web/GetFolderByServerRelativePath(decodedurl='${folder}')/Files?$select=Name,ServerRelativeUrl&$top=5000`,
        SPHttpClient.configurations.v1,
        { headers: { Accept: 'application/json;odata.metadata=none' } }
      );
      expect(get).toHaveBeenNthCalledWith(2,
        `https://whallc.sharepoint.com${folder}/URLFixerIndex.csv`, SPHttpClient.configurations.v1
      );
      expect(get).toHaveBeenNthCalledWith(3,
        `https://whallc.sharepoint.com${folder}/CS-fixer.csv`, SPHttpClient.configurations.v1
      );
    });

    it('falls back to the table for an old GUID and returns the new resolved GUID', async () => {
      queueTables();
      fetchMock.mockResolvedValue({ ok: true, status: 200 });
      expect(fixUrl(original).note).toMatch(/^err:/);
      expect(fixUrl(original).fixed).toContain('/sites/ClinicalandPatientServices/');
      expect(await lookupOrFixUrl(original, client)).toMatchObject({
        original,
        fixed: resolved,
        lookupTableFileName: 'CS-fixer.csv',
        note: expect.stringMatching(/^ok:/)
      });
      expect(fetchMock).toHaveBeenCalledTimes(1);
      expect(fetchMock).toHaveBeenCalledWith(resolved, expect.objectContaining({ method: 'HEAD' }));
    });

    it('uses the root-site lookup table to resolve an old E&M library Doc.aspx GUID', async () => {
      const rootIndex = 'OldBaseURL,FixerCsvFileName,NewBaseURL\r\nhttps://whallc.sharepoint.com/,WHASharepointCentralPark-fixer.csv,https://whallc.sharepoint.com/';
      const rootFixer = [
        '"OldResolvedUrl","NewFound","NewCopyLinkUrl","NewResolvedUrl","Notes"',
        `"${legacyEmResolved}","Yes","https://whallc.sharepoint.com/:w:/r/sites/PatientFinancialServices/2021%20EM%20Coding%20Changes/FAQ_2021%20E%26M%20Coding.docx?d=w4eed5efc08b84d4d94d5e59fc28ef423&csf=1&web=1","${migratedEmResolved}",""`
      ].join('\r\n');
      const rootFiles = ['URLFixerIndex.csv', 'WHASharepointCentralPark-fixer.csv']
        .map((Name) => ({ Name, ServerRelativeUrl: `${folder}/${Name}` }));
      get.mockResolvedValueOnce(response(JSON.stringify({ value: rootFiles })))
        .mockResolvedValueOnce(response(rootIndex))
        .mockResolvedValueOnce(response(rootFixer));
      fetchMock.mockResolvedValue({ ok: true, status: 200 });

      expect(await lookupOrFixUrl(legacyEmResolved, client)).toMatchObject({
        original: legacyEmResolved,
        fixed: migratedEmResolved,
        lookupTableFileName: 'WHASharepointCentralPark-fixer.csv',
        note: expect.stringMatching(/^ok:/)
      });
      expect(get).toHaveBeenNthCalledWith(3,
        `https://whallc.sharepoint.com${folder}/WHASharepointCentralPark-fixer.csv`,
        SPHttpClient.configurations.v1
      );
    });

    it.each([
      { message: 'Invalid Accept header.' },
      { message: { value: 'Invalid Accept header.' } }
    ])('reports SharePoint discovery details and retries after failure: %j', async (error) => {
      get.mockResolvedValueOnce(response(JSON.stringify({ error }), 400));
      await expect(client.findLookupRow(original)).rejects.toThrow('Lookup table discovery failed (400). Invalid Accept header.');
      queueTables();
      expect(await client.findLookupRow(original)).toMatchObject({ NewResolvedUrl: resolved });
      expect(get).toHaveBeenCalledTimes(4);
    });

    it.each(['{}', 'null'])('rejects an unexpected discovery response: %s', async (body) => {
      get.mockResolvedValueOnce(response(body));
      await expect(client.findLookupRow(original)).rejects.toThrow('Lookup table discovery returned an invalid file list.');
    });

    it.each(['null', '<html>Bad request</html>'])('preserves the HTTP error when the response has no structured error: %s', async (body) => {
      get.mockResolvedValueOnce(response(body, 400));
      await expect(client.findLookupRow(original)).rejects.toThrow('Lookup table discovery failed (400).');
    });

    it('retries an index download after a transient failure', async () => {
      get.mockResolvedValueOnce(response(JSON.stringify({ value: files })))
        .mockResolvedValueOnce(response('Unavailable', 503));
      await expect(client.findLookupRow(original)).rejects.toThrow('Could not load URLFixerIndex.csv (503).');
      get.mockResolvedValueOnce(response(indexCsv)).mockResolvedValueOnce(response(fixerCsv));
      expect(await client.findLookupRow(original)).toMatchObject({ NewResolvedUrl: resolved });
      expect(get).toHaveBeenCalledTimes(4);
    });

    it('reports a missing fixer CSV as a table error', async () => {
      get.mockResolvedValueOnce(response(JSON.stringify({ value: [files[0]] })))
        .mockResolvedValueOnce(response(indexCsv));
      await expect(client.findLookupRow(original)).rejects.toThrow('Could not find CS-fixer.csv in the URL Lookup Tables library.');
    });

    it('reports a fixer CSV download failure instead of claiming the document is missing', async () => {
      get.mockResolvedValueOnce(response(JSON.stringify({ value: files })))
        .mockResolvedValueOnce(response(indexCsv))
        .mockResolvedValueOnce(response('Forbidden', 403));
      await expect(client.findLookupRow(original)).rejects.toThrow('Could not load CS-fixer.csv (403).');
      get.mockResolvedValueOnce(response(fixerCsv));
      expect(await client.findLookupRow(original)).toMatchObject({ NewResolvedUrl: resolved });
    });

    it('prioritizes a verbatim URL over an earlier row with the same GUID', async () => {
      const other = original.replace('file=Release', 'file=Other');
      queueFixer([
        fixerCsv.split('\r\n')[0],
        `"${other}","Yes","${copyLink}","${resolved}",""`,
        `"${original}","No","","","Not migrated"`
      ].join('\r\n'));
      expect(await client.findLookupRow(original)).toMatchObject({ OldResolvedUrl: original, NewFound: 'No' });
      expect(await lookupOrFixUrl(original, client)).toMatchObject({
        note: expect.stringMatching(/^err:/),
        lookupTableFileName: 'CS-fixer.csv'
      });
      expect(fetchMock).not.toHaveBeenCalled();
    });

    it('reports conflicting GUID matches instead of selecting the first row', async () => {
      queueFixer(`${fixerCsv}\r\n"${original.replace('file=Release', 'file=Other')}","No","","","Not migrated"`);
      await expect(client.findLookupRow(`${original}&web=1`)).rejects.toThrow('Conflicting mappings in CS-fixer.csv');
    });

    it('reports conflicting exact matches but permits identical duplicates', async () => {
      queueFixer(`${fixerCsv}\r\n"${original}","No","","","Not migrated"`);
      await expect(client.findLookupRow(original)).rejects.toThrow('Conflicting mappings');
      const duplicateClient = new LookupTableClient(httpClient, siteUrl);
      queueFixer(`${fixerCsv}\r\n${fixerCsv.split('\r\n')[1]}`);
      expect(await duplicateClient.findLookupRow(original)).toMatchObject({ NewResolvedUrl: resolved });
    });

    it('prioritizes document paths when GUIDs are shared across different paths', async () => {
      const path = 'https://whallc.sharepoint.com/:w:/r/CS/Folder/Release.docx';
      const first = `${path}?d=wb5aa768749dc4681a326b80253394172`;
      const second = first.replace('/Folder/', '/OtherFolder/');
      queueFixer([
        '"OldCopyLinkUrl","OldFileRef","NewFound","NewCopyLinkUrl","NewResolvedUrl"',
        `"${first}","/CS/Folder/Release.docx","Yes","${copyLink}","${resolved}"`,
        `"${second}","/CS/OtherFolder/Release.docx","No","",""`
      ].join('\r\n'));
      expect(await client.findLookupRow(`${second}&web=1`)).toMatchObject({ OldCopyLinkUrl: second, NewFound: 'No' });
    });

    it('matches PDF id paths even when optional parent parameters differ', async () => {
      const pdf = 'https://whallc.sharepoint.com/CS/Documents/Forms/AllItems.aspx?id=%2FCS%2FDocuments%2FGuide.pdf&parent=%2FCS%2FDocuments';
      queueFixer(fixerCsv.replace(original, pdf));
      expect(await client.findLookupRow(pdf.split('&parent=')[0])).toMatchObject({ NewResolvedUrl: resolved });
    });

    it.each([
      '<html>Sign in</html>',
      'Wrong,Headers\r\nvalue,value',
      'OldBaseURL,FixerCsvFileName\r\nhttps://whallc.sharepoint.com/CS',
      'OldBaseURL,FixerCsvFileName\r\n"unterminated,CS-fixer.csv'
    ])('rejects invalid index CSV and permits a corrected retry: %s', async (body) => {
      get.mockResolvedValueOnce(response(JSON.stringify({ value: files }))).mockResolvedValueOnce(response(body));
      await expect(client.findLookupRow(original)).rejects.toThrow('Invalid lookup CSV URLFixerIndex.csv');
      get.mockResolvedValueOnce(response(indexCsv)).mockResolvedValueOnce(response(fixerCsv));
      expect(await client.findLookupRow(original)).toMatchObject({ NewResolvedUrl: resolved });
    });

    it.each([
      '<html>Sign in</html>',
      'OldResolvedUrl,NewFound,NewCopyLinkUrl,NewResolvedUrl\r\ninvalid,Yes,,',
      fixerCsv.replace('"Yes"', '"Maybe"'),
      fixerCsv.replace(resolved, 'https://example.com/file'),
      `${fixerCsv},extra`,
      fixerCsv.replace('"Notes"', '"NewFound"')
    ])('rejects invalid fixer CSV without caching it: %s', async (body) => {
      queueFixer(body);
      await expect(client.findLookupRow(original)).rejects.toThrow('Invalid lookup CSV CS-fixer.csv');
      get.mockResolvedValueOnce(response(fixerCsv));
      expect(await client.findLookupRow(original)).toMatchObject({ NewResolvedUrl: resolved });
    });

    it('supports BOM, quoted commas, escaped quotes, and multiline fields', async () => {
      queueFixer(`\uFEFF${fixerCsv.replace(/,""$/, ',"First line,\r\nsecond ""quoted"" line"')}`);
      expect(await client.findLookupRow(original)).toMatchObject({ Notes: 'First line,\r\nsecond "quoted" line' });
    });

    it('accepts a valid header-only table as empty', async () => {
      queueFixer(fixerCsv.split('\r\n')[0]);
      expect(await client.findLookupRow(original)).toBeUndefined();
    });

    it.each([null, { Name: 'CS-fixer.csv' }, { Name: 'CS-fixer.csv', ServerRelativeUrl: '//example.com/file' }])(
      'rejects malformed discovery file entries: %j', async (entry) => {
        get.mockResolvedValueOnce(response(JSON.stringify({ value: [entry] })));
        await expect(client.findLookupRow(original)).rejects.toThrow('Lookup table discovery returned an invalid file entry.');
      }
    );

    it.each([401, 403, 404, 405, 503])('preserves mapped links with a diagnostic warning for HTTP %s', async (status) => {
      queueTables();
      fetchMock.mockResolvedValue({ ok: false, status });
      const result = await lookupOrFixUrl(original, client);
      expect(result.fixed).toBe(resolved);
      expect(result.lookupTableFileName).toBe('CS-fixer.csv');
      expect(result.note).toMatch(/^warn:A migrated URL was found/);
      expect(result.note).toContain(`${status}`);
      expect(result.note).not.toContain('No migrated URL found');
    });

    it('preserves mapped links when a network check throws', async () => {
      queueTables();
      fetchMock.mockRejectedValue(new Error('Network unavailable'));
      expect(await lookupOrFixUrl(original, client)).toMatchObject({
        fixed: resolved, note: expect.stringContaining('Network unavailable')
      });
    });

    it.each([
      'https://whallc.sharepoint.com/_layouts/15/AccessDenied.aspx',
      'https://login.microsoftonline.com/signin'
    ])('does not treat a successful error/sign-in redirect as a verified document: %s', async (url) => {
      queueTables();
      fetchMock.mockResolvedValue({ ok: true, status: 200, url });
      expect(await lookupOrFixUrl(original, client)).toMatchObject({
        fixed: resolved, note: expect.stringMatching(/^warn:.*redirected/)
      });
    });

    it('uses the CSV fallback when the pattern check redirects to an error page', async () => {
      queueFixer(fixerCsv.replace(original, 'https://whallc.sharepoint.com/CS/Guide.docx'));
      fetchMock.mockResolvedValueOnce({ ok: true, status: 200, url: 'https://whallc.sharepoint.com/_layouts/15/error.aspx' })
        .mockResolvedValueOnce({ ok: true, status: 200 });
      expect(await lookupOrFixUrl('https://whallc.sharepoint.com/CS/Guide.docx', client)).toMatchObject({
        fixed: copyLink, lookupTableFileName: 'CS-fixer.csv'
      });
      expect(fetchMock).toHaveBeenCalledTimes(2);
    });

    it('keeps the no-mapping outcome separate from a mapped-link verification warning', async () => {
      queueFixer(fixerCsv.split('\r\n')[0]);
      const result = await lookupOrFixUrl(original, client);
      expect(result.note).toMatch(/^err:No migrated URL found/);
      expect(result.lookupTableFileName).toBeUndefined();
      expect(fetchMock).not.toHaveBeenCalled();
    });

    it('does not attach a table filename to a reachable pattern-only result', async () => {
      fetchMock.mockResolvedValue({ ok: true, status: 200 });
      const result = await lookupOrFixUrl('https://whallc.sharepoint.com/CS/Guide.docx', client);
      expect(result.lookupTableFileName).toBeUndefined();
      expect(get).not.toHaveBeenCalled();
    });

    it('routes encoded site names to their index entry', async () => {
      const encoded = original.replace('/CS/', '/IT%26HIM/');
      get.mockResolvedValueOnce(response(JSON.stringify({ value: files })))
        .mockResolvedValueOnce(response(indexCsv.replace('/CS,', '/IT&HIM,')))
        .mockResolvedValueOnce(response(fixerCsv.replace(original, encoded)));
      expect(await client.findLookupRow(encoded)).toMatchObject({ NewResolvedUrl: resolved });
    });

    it('does not select a site index entry from a partial segment match', async () => {
      get.mockResolvedValueOnce(response(JSON.stringify({ value: files })))
        .mockResolvedValueOnce(response(indexCsv));
      expect(await client.findLookupRow(original.replace('/CS/', '/CSArchive/'))).toBeUndefined();
      expect(get).toHaveBeenCalledTimes(2);
    });
  });

  it('preserves legacy site and document URL conversion', () => {
    const result = fixUrl('https://whallc.sharepoint.com/CS/Ultrasound%20Department/Guide.docx');
    expect(result.fixed).toBe('https://whallc.sharepoint.com/:w:/r/sites/ClinicalandPatientServices/Ultrasound%20Department/Guide.docx');
    expect(result.note).toMatch(/^ok:/);
  });

  it.each([
    'https://whallc.sharepoint.com/2021%20EM%20Coding%20Changes/FAQ.docx',
    'https://whallc.sharepoint.com/:w:/r/2021%20E%26M%20Coding%20Changes/FAQ.docx?d=w123&csf=1&web=1'
  ])('pattern-fixes the legacy E&M library path: %s', (url) => {
    const result = fixUrl(url);
    expect(result.fixed).toContain('/sites/PatientFinancialServices/2021%20EM%20Coding%20Changes/FAQ.docx');
    expect(result.note).toMatch(/^ok:/);
  });

  it('pattern-fixes the legacy E&M path in embedded document parameters', () => {
    const result = fixUrl('https://whallc.sharepoint.com/2021%20EM%20Coding%20Changes/Forms/AllItems.aspx?id=%2F2021%20EM%20Coding%20Changes%2FFAQ.pdf&parent=%2F2021%20EM%20Coding%20Changes');
    expect(result.fixed).toContain('/sites/PatientFinancialServices/2021%20EM%20Coding%20Changes/Forms/AllItems.aspx');
    expect(result.fixed).toContain('id=%2Fsites%2FPatientFinancialServices%2F2021%20EM%20Coding%20Changes%2FFAQ.pdf');
    expect(result.fixed).toContain('parent=%2Fsites%2FPatientFinancialServices%2F2021%20EM%20Coding%20Changes');
  });

  it.each(['https://WHALLC.SHAREPOINT.COM', 'whallc.sharepoint.com', 'http://whallc.sharepoint.com'])(
    'preserves existing accepted host formats: %s', (host) => {
      expect(fixUrl(`${host}/CS/Folder`).note).toMatch(/^ok:/);
    }
  );

  it('preserves already-current URLs', () => {
    const url = 'https://whallc.sharepoint.com/sites/ClinicalandPatientServices/Folder';
    expect(fixUrl(url).fixed).toBe(url);
    expect(fixUrl(url).note).toMatch(/^warn:/);
  });
});
