import { buildQuickFindBookmarklet, getLaunchUrl, getQuickFindPageUrl } from './quickFind';

describe('Quick Find helpers', () => {
  const pageUrl = 'https://whallc.sharepoint.com/sites/CentralParkLinkFinder/SitePages/Quick%20Find.aspx';
  const oldLink = 'https://whallc.sharepoint.com/CS/Ultrasound%20Department/Guide.docx';

  it('reads an encoded old link from the url parameter', () => {
    expect(getLaunchUrl(`?url=${encodeURIComponent(oldLink)}`)).toBe(oldLink);
  });

  it('ignores missing, blank, and non-web url parameters', () => {
    expect(getLaunchUrl('')).toBeUndefined();
    expect(getLaunchUrl('?url=')).toBeUndefined();
    expect(getLaunchUrl('?url=not-a-link')).toBeUndefined();
    // eslint-disable-next-line no-script-url -- the Favorites bar button is intentionally a JavaScript link
    expect(getLaunchUrl(`?url=${encodeURIComponent(`javascript:alert(1)`)}`)).toBeUndefined();
  });

  it('builds the page address without any query string or hash', () => {
    expect(getQuickFindPageUrl({
      origin: 'https://whallc.sharepoint.com',
      pathname: '/sites/CentralParkLinkFinder/SitePages/Quick%20Find.aspx'
    })).toBe(pageUrl);
  });

  it('builds a Favorites bar button that sends the current page to Quick Find', () => {
    const href = buildQuickFindBookmarklet(pageUrl);
    // eslint-disable-next-line no-script-url -- the Favorites bar button is intentionally a JavaScript link
    const prefix = `javascript:`;
    expect(href.startsWith(prefix)).toBe(true);

    // Browsers percent-decode the link before running it.
    const script = decodeURIComponent(href.slice(prefix.length));
    expect(script).toBe(
      `(function(){window.location.href=${JSON.stringify(`${pageUrl}?url=`)}+encodeURIComponent(window.location.href);})();`
    );
  });
});
