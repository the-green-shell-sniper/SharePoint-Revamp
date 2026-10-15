const URL_PARAMETER = 'url';
export function navigateToDestination(url: string): void {
  window.location.replace(url);
}
export const QUICK_FIND_PAGE_URL = 'https://whallc.sharepoint.com/sites/CentralParkLinkFinder/SitePages/Link-Finder-Extension.aspx';

/** Reads the old link passed in by the Favorites bar button, e.g. `?url=https%3A%2F%2F...`. */
export function getLaunchUrl(search: string): string | undefined {
  const value = new URLSearchParams(search).get(URL_PARAMETER)?.trim();
  if (!value) return undefined;
  try {
    const url = new URL(value);
    return url.protocol === 'http:' || url.protocol === 'https:' ? value : undefined;
  } catch {
    return undefined;
  }
}

/**
 * Builds the Favorites bar button (bookmarklet). When clicked on any page, it opens this
 * Quick Find page with that page's address in the `url` parameter.
 */
export function buildQuickFindBookmarklet(pageUrl: string): string {
  const target = `${pageUrl}?${URL_PARAMETER}=`;
  const script = `(function(){window.location.href=${JSON.stringify(target)}+encodeURIComponent(window.location.href);})();`;
  // Browsers percent-decode javascript: links before running them, so keep literal % signs intact.
  return `javascript:${script.replace(/%/g, '%25')}`;
}
