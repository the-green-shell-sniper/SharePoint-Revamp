jest.mock('@microsoft/sp-http', () => ({ SPHttpClient: jest.fn() }));
jest.mock('@microsoft/sp-core-library', () => ({ ServiceScope: { startNewRoot: jest.fn() } }));

import { ServiceScope } from '@microsoft/sp-core-library';
import { SPHttpClient } from '@microsoft/sp-http';
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

  it('preserves legacy site and document URL conversion', () => {
    const result = fixUrl('https://whallc.sharepoint.com/CS/Ultrasound%20Department/Guide.docx');
    expect(result.fixed).toBe('https://whallc.sharepoint.com/:w:/r/sites/ClinicalandPatientServices/Ultrasound%20Department/Guide.docx');
    expect(result.note).toMatch(/^ok:/);
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
