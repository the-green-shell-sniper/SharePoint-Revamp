/* eslint-disable @rushstack/pair-react-dom-render-unmount -- render helper is paired with afterEach cleanup */
jest.mock('@microsoft/sp-http', () => ({ SPHttpClient: jest.fn() }));
jest.mock('@microsoft/sp-core-library', () => ({
  ServiceScope: { startNewRoot: jest.fn() }, Log: { error: jest.fn() }
}));
jest.mock('./urlFixerLogic', () => ({ LookupTableClient: jest.fn(), lookupOrFixUrl: jest.fn() }));
jest.mock('./submissionLog', () => ({ logSubmission: jest.fn() }));
jest.mock('./quickFind', () => ({
  ...jest.requireActual('./quickFind'), navigateToDestination: jest.fn()
}));

import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { act, Simulate } from 'react-dom/test-utils';
import { Log, ServiceScope } from '@microsoft/sp-core-library';
import { SPHttpClient } from '@microsoft/sp-http';
import SharePointUrlFixer from './SharePointUrlFixer';
import { IFixResult, lookupOrFixUrl } from './urlFixerLogic';
import { logSubmission } from './submissionLog';
import { navigateToDestination, QUICK_FIND_PAGE_URL } from './quickFind';

describe('Quick Find setup and automatic navigation', () => {
  let container: HTMLDivElement;
  const oldUrl = 'https://whallc.sharepoint.com/CS/Guide.aspx';
  const newUrl = 'https://whallc.sharepoint.com/sites/ClinicalandPatientServices/Guide.aspx';
  const lookup = jest.mocked(lookupOrFixUrl);
  const navigate = jest.mocked(navigateToDestination);
  const log = jest.mocked(logSubmission);

  async function render(launchUrl?: string, launchError?: string): Promise<void> {
    await act(async () => {
      ReactDOM.render(
        <SharePointUrlFixer spHttpClient={new SPHttpClient(ServiceScope.startNewRoot())}
          siteAbsoluteUrl="https://whallc.sharepoint.com"
          currentUser={{ loginName: 'user@example.test', displayName: 'Test User' }}
          pageUrl={QUICK_FIND_PAGE_URL} launchUrl={launchUrl} launchError={launchError} />,
        container
      );
    });
  }

  function click(label: string): void {
    const button = Array.from(container.querySelectorAll('button')).find((item) => item.textContent === label);
    if (!button) throw new Error(`Missing button: ${label}`);
    act(() => { Simulate.click(button); });
  }

  function advance(seconds: number): void {
    for (let second = 0; second < seconds; second++) {
      act(() => { jest.advanceTimersByTime(1000); });
    }
  }

  beforeEach(() => {
    jest.useFakeTimers();
    jest.clearAllMocks();
    navigate.mockReset();
    lookup.mockResolvedValue({ original: oldUrl, fixed: newUrl, note: 'ok:Found migrated link.' });
    log.mockResolvedValue(undefined);
    container = document.createElement('div');
    document.body.appendChild(container);
  });

  afterEach(() => {
    act(() => { ReactDOM.unmountComponentAtNode(container); });
    container.remove();
    jest.useRealTimers();
  });

  it('shows a standalone setup guide with no input or manual lookup button', async () => {
    await render();
    expect(container.querySelector('textarea, input')).toBeNull();
    expect(container.querySelector('button')).toBeNull();
    expect(container.querySelector('h1')?.textContent).toBe('Link Finder Extension');
    expect(container.querySelector('header')?.textContent).toContain('For Central Park links only.');
    expect(container.querySelector('header')?.textContent).toContain('It will not fix links to other websites.');
    expect(container.textContent).toContain('extra option, not a replacement');
    expect(container.textContent).toContain('For Central Park links only.');
    expect(container.textContent).toContain('Done');
    expect(container.textContent).toContain('Delete');
    expect(container.textContent).toContain('left mouse button');
    expect(container.querySelector('details')?.open).toBe(false);
    expect(container.textContent).toContain('3. Save the new page for next time');
    expect(container.querySelector('#quickFindUse')?.textContent).toBe('2. Use it when a Central Park link stops working');
    expect(container.textContent).toContain('Look just below the box at the top of Edge that shows the website address.');
    expect(lookup).not.toHaveBeenCalled();
    const anchor = container.querySelector<HTMLAnchorElement>('.dragLink');
    expect(anchor?.draggable).toBe(true);
    expect(anchor?.getAttribute('href')).toContain(`${QUICK_FIND_PAGE_URL}?url=`);
    if (!anchor) throw new Error('Missing bookmarklet');
    act(() => { Simulate.click(anchor); });
    expect(container.textContent).toContain('instead of clicking it here');
  });

  it('automatically looks up and opens a new link after ten seconds without the setup heading', async () => {
    await render(oldUrl);
    expect(lookup).toHaveBeenCalledWith(oldUrl, expect.anything());
    expect(container.textContent).toContain('Your new link is ready');
    expect(container.textContent).toContain('Ctrl+D');
    expect(container.querySelector('h1')).toBeNull();
    expect(container.textContent).not.toContain('Find New Link for Central Park');
    expect(container.textContent).toContain('10 seconds');
    advance(9);
    expect(navigate).not.toHaveBeenCalled();
    advance(1);
    expect(navigate).toHaveBeenCalledTimes(1);
    expect(navigate).toHaveBeenCalledWith(newUrl);
    advance(10);
    expect(navigate).toHaveBeenCalledTimes(1);
    expect(log).toHaveBeenCalledWith(expect.anything(), expect.anything(),
      expect.objectContaining({ providedUrl: oldUrl, returnedUrl: newUrl, outcome: 'Success' }));
  });

  it('pauses and resumes the countdown', async () => {
    await render(oldUrl);
    advance(2);
    click('Pause');
    advance(10);
    expect(navigate).not.toHaveBeenCalled();
    click('Resume');
    advance(8);
    expect(navigate).toHaveBeenCalledWith(newUrl);
  });

  it('opens immediately and cancels the countdown', async () => {
    await render(oldUrl);
    click('Open now');
    advance(10);
    expect(navigate).toHaveBeenCalledTimes(1);
  });

  it('cancels navigation when unmounted', async () => {
    await render(oldUrl);
    act(() => { ReactDOM.unmountComponentAtNode(container); });
    advance(10);
    expect(navigate).not.toHaveBeenCalled();
  });

  it('does not start a countdown if unmounted during lookup', async () => {
    let resolveLookup: ((result: IFixResult) => void) | undefined;
    lookup.mockReturnValueOnce(new Promise((resolve) => { resolveLookup = resolve; }));
    await render(oldUrl);
    act(() => { ReactDOM.unmountComponentAtNode(container); });
    await act(async () => {
      if (resolveLookup) resolveLookup({ original: oldUrl, fixed: newUrl, note: 'ok:Found.' });
    });
    advance(10);
    expect(navigate).not.toHaveBeenCalled();
  });

  it('returns to an already-current link without a redirect loop', async () => {
    lookup.mockResolvedValueOnce({ original: newUrl, fixed: newUrl, note: 'warn:Already has /sites — left unchanged.' });
    await render(newUrl);
    expect(container.textContent).toContain('already up to date');
    advance(10);
    expect(navigate).toHaveBeenCalledWith(newUrl);
  });

  it('keeps unverified results visible and requires explicit navigation', async () => {
    lookup.mockResolvedValueOnce({ original: oldUrl, fixed: newUrl, note: 'warn:Availability could not be verified.' });
    await render(oldUrl);
    advance(10);
    expect(navigate).not.toHaveBeenCalled();
    expect(container.textContent).toContain('Availability could not be verified');
    expect(container.querySelector<HTMLAnchorElement>('.resultLink')?.href).toBe(newUrl);
  });

  it('shows lookup errors and links to the manual tool', async () => {
    lookup.mockResolvedValueOnce({ original: oldUrl, fixed: '', note: 'err:No matching link found.' });
    await render(oldUrl);
    advance(10);
    expect(container.querySelector('[role="alert"]')?.textContent).toContain('No matching link found');
    expect(container.textContent).toContain('manual Link Finder');
    expect(navigate).not.toHaveBeenCalled();
  });

  it('reports failed requests', async () => {
    lookup.mockRejectedValueOnce(new Error('Lookup service unavailable.'));
    await render(oldUrl);
    expect(container.querySelector('[role="alert"]')?.textContent).toBe('Lookup service unavailable.');
    advance(10);
    expect(navigate).not.toHaveBeenCalled();
  });

  it('rejects external source URLs before calling lookup', async () => {
    await render('https://example.com/old');
    expect(lookup).not.toHaveBeenCalled();
    expect(container.textContent).toContain('only works with Central Park links');
  });

  it('reports malformed launch parameters instead of showing setup', async () => {
    await render(undefined, 'Invalid incoming link.');
    expect(container.querySelector('[role="alert"]')?.textContent).toBe('Invalid incoming link.');
    expect(lookup).not.toHaveBeenCalled();
  });

  it('rejects unexpected external destinations', async () => {
    lookup.mockResolvedValueOnce({ original: oldUrl, fixed: 'https://example.com/', note: 'ok:Found.' });
    await render(oldUrl);
    advance(10);
    expect(navigate).not.toHaveBeenCalled();
    expect(container.textContent).toContain('trusted Central Park destination');
  });

  it('does not redirect back to the extension itself', async () => {
    lookup.mockResolvedValueOnce({ original: oldUrl, fixed: QUICK_FIND_PAGE_URL, note: 'ok:Found.' });
    await render(oldUrl);
    advance(10);
    expect(navigate).not.toHaveBeenCalled();
    expect(container.textContent).toContain('already on the Link Finder extension page');
  });

  it('logs submission save failures without blocking navigation', async () => {
    log.mockRejectedValueOnce(new Error('Log unavailable.'));
    await render(oldUrl);
    expect(Log.error).toHaveBeenCalledWith('QuickFindSubmissionLog', expect.any(Error));
    advance(10);
    expect(navigate).toHaveBeenCalledWith(newUrl);
  });
});
