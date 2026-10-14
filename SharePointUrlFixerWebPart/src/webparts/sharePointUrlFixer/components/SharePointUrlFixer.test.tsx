jest.mock('@microsoft/sp-http', () => ({ SPHttpClient: jest.fn() }));
jest.mock('@microsoft/sp-core-library', () => ({
  ServiceScope: { startNewRoot: jest.fn() },
  Log: { error: jest.fn() }
}));
jest.mock('./urlFixerLogic', () => ({
  LookupTableClient: jest.fn(),
  lookupOrFixUrl: jest.fn()
}));
jest.mock('./favoriteTitle', () => ({ getFavoriteTitle: jest.fn() }));
jest.mock('./submissionLog', () => ({ logSubmission: jest.fn() }));

import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { act, Simulate } from 'react-dom/test-utils';
import { Log, ServiceScope } from '@microsoft/sp-core-library';
import { SPHttpClient } from '@microsoft/sp-http';
import SharePointUrlFixer from './SharePointUrlFixer';
import { IFixResult, lookupOrFixUrl } from './urlFixerLogic';
import { getFavoriteTitle } from './favoriteTitle';
import { logSubmission } from './submissionLog';

describe('URL fixer interactions', () => {
  let container: HTMLDivElement;
  const lookup = jest.mocked(lookupOrFixUrl);
  const titleLookup = jest.mocked(getFavoriteTitle);
  const log = jest.mocked(logSubmission);
  const currentUser = { loginName: 'i:0#.f|membership|user@wha.example', displayName: 'Test User' };
  const original = 'https://whallc.sharepoint.com/CS/Ultrasound%20Department/Guide.docx';
  const fixed = 'https://whallc.sharepoint.com/sites/ClinicalandPatientServices/Ultrasound%20Department/Guide.docx';

  function textarea(): HTMLTextAreaElement {
    const element = container.querySelector('textarea');
    if (!element) throw new Error('Missing URL input');
    return element;
  }

  function button(label: string): HTMLButtonElement {
    const element = Array.from(container.querySelectorAll('button')).find((candidate) => candidate.textContent === label);
    if (!element) throw new Error(`Missing button: ${label}`);
    return element;
  }

  function enterUrl(url: string): void {
    act(() => {
      textarea().value = url;
      Simulate.change(textarea());
    });
  }

  async function submit(): Promise<void> {
    await act(async () => { Simulate.click(button('Find Updated Link')); });
  }

  beforeEach(() => {
    log.mockReset();
    jest.mocked(Log.error).mockClear();
    log.mockResolvedValue(undefined);
    lookup.mockResolvedValue({ original, fixed, note: 'ok:Updated link ready.' });
    titleLookup.mockResolvedValue({ title: 'Ultrasound Guide' });
    container = document.createElement('div');
    document.body.appendChild(container);
    act(() => {
      ReactDOM.render(
        <SharePointUrlFixer spHttpClient={new SPHttpClient(ServiceScope.startNewRoot())} siteAbsoluteUrl="https://whallc.sharepoint.com" currentUser={currentUser} />,
        container
      );
    });
  });

  afterEach(() => {
    act(() => { ReactDOM.unmountComponentAtNode(container); });
    container.remove();
    jest.restoreAllMocks();
  });

  it('shows the requested placeholder', () => {
    expect(textarea().placeholder).toBe('https://whallc.sharepoint.com/CS/Ultrasound%20Department/...');
  });

  it('explains how to provide a link and find its migrated destination', () => {
    expect(container.querySelector('label[for="urlFixerInput"]')?.textContent).toBe('SharePoint page or file URL');
    expect(textarea().getAttribute('aria-describedby')).toBe('urlFixerInstructions');
    expect(container.textContent).toContain('Copy the full URL from the browser\'s address bar, paste it here');
    expect(container.textContent).toContain('Use the updated result to open or copy the migrated page.');
    expect(container.textContent).toContain('then select Find Updated Link or press Enter.');
  });

  it('submits with Enter instead of inserting a line break', async () => {
    enterUrl(original);
    const preventDefault = jest.fn();
    await act(async () => {
      Simulate.keyDown(textarea(), { key: 'Enter', preventDefault });
    });
    expect(preventDefault).toHaveBeenCalled();
    expect(lookup).toHaveBeenCalledTimes(1);
    expect(lookup).toHaveBeenCalledWith(original, expect.anything());
    expect(log).toHaveBeenCalledTimes(1);
    expect(textarea().value).toBe(original);
  });

  it('does not submit during input method composition', () => {
    enterUrl(original);
    act(() => {
      Simulate.keyDown(textarea(), { key: 'Enter', nativeEvent: new KeyboardEvent('keydown', { isComposing: true }) });
    });
    expect(lookup).not.toHaveBeenCalled();
    expect(log).not.toHaveBeenCalled();
  });

  it('does not start a second lookup while one is pending', async () => {
    let resolveLookup: ((result: IFixResult) => void) | undefined;
    lookup.mockReturnValueOnce(new Promise<IFixResult>((resolve) => { resolveLookup = resolve; }));
    enterUrl(original);
    act(() => { Simulate.click(button('Find Updated Link')); });
    expect(button('Finding link...').disabled).toBe(true);
    act(() => { Simulate.keyDown(textarea(), { key: 'Enter' }); });
    expect(lookup).toHaveBeenCalledTimes(1);
    await act(async () => {
      if (!resolveLookup) throw new Error('Missing pending lookup');
      resolveLookup({ original, fixed, note: 'ok:Updated link ready.' });
    });
  });

  it('offers native draggable destination links and keyboard instructions', async () => {
    enterUrl(original);
    await submit();
    const links = Array.from(container.querySelectorAll('a[draggable="true"]'));
    expect(links).toHaveLength(2);
    links.forEach((link) => { expect(link.getAttribute('href')).toBe(fixed); });
    expect(links[1].textContent).toBe('Ultrasound Guide');
    expect(links[1].getAttribute('title')).toBe('Ultrasound Guide');
    expect(titleLookup).toHaveBeenCalledWith(fixed);
    expect(container.textContent).toContain('Save as a bookmark or shortcut');
    expect(container.textContent).toContain('onto Edge\'s bookmarks bar');
    expect(container.textContent).toContain('onto your desktop to create a shortcut');
    expect(container.textContent).toContain('Ctrl+Shift+B');
    expect(container.textContent).toContain('Ctrl+D');
    const instructions = container.querySelector('ul#urlFixerDragInstructions');
    expect(instructions).not.toBeNull();
    const methods = Array.from(instructions?.querySelectorAll('li') || []).map((item) => item.textContent);
    expect(methods).toEqual([
      "Bookmark by dragging: Drag the button below onto Edge's bookmarks bar. Press Ctrl+Shift+B to show the bar.",
      'Bookmark with the keyboard: Open the updated link, then press Ctrl+D in the new tab to save it as a bookmark.',
      'Desktop shortcut: Drag the button below onto your desktop to create a shortcut.'
    ]);
    expect(instructions?.previousElementSibling?.textContent).toBe('Save as a bookmark or shortcut');
    expect(container.textContent).toContain('You will need to manually delete any favorites, bookmarks, or shortcuts that have the old link.');
    expect(button('Open Updated Link').title).toBe(
      'Opens in a new tab; after opening, press ctrl+d on your keyboard to create a favorite of this link.'
    );
  });

  it('keeps opening and copying the returned URL', async () => {
    const open = jest.spyOn(window, 'open').mockReturnValue(null);
    const writeText = jest.fn().mockResolvedValue(undefined);
    Object.defineProperty(navigator, 'clipboard', { configurable: true, value: { writeText } });
    enterUrl(original);
    await submit();
    act(() => { Simulate.click(button('Open Updated Link')); });
    expect(open).toHaveBeenCalledWith(fixed, '_blank', 'noopener,noreferrer');
    await act(async () => { Simulate.click(button('Copy Updated Link')); });
    expect(writeText).toHaveBeenCalledWith(fixed);
    expect(container.textContent).toContain('Copied to clipboard!');
  });

  it('shows errors as alerts without exposing a draggable result', async () => {
    const message = "You've provided an external link. This tool is only for WHA SharePoint links.";
    lookup.mockResolvedValueOnce({ original: 'https://example.com', fixed: 'https://example.com', note: `err:${message}` });
    enterUrl('https://example.com');
    await submit();
    expect(container.querySelector('[role="alert"]')?.textContent).toBe(message);
    expect(container.querySelector('a[draggable="true"]')).toBeNull();
    expect(titleLookup).not.toHaveBeenCalled();
  });

  it('shows mapped-link verification warnings without calling the link already current', async () => {
    const message = 'A migrated URL was found in the lookup table, but its availability could not be verified. SharePoint denied access (403).';
    lookup.mockResolvedValueOnce({ original, fixed, note: `warn:${message}` });
    enterUrl(original);
    await submit();
    expect(container.querySelector('[role="status"]')?.textContent).toBe(message);
    expect(container.textContent).not.toContain('This link is already current.');
    expect(container.querySelector('a')?.getAttribute('href')).toBe(fixed);
    expect(button('Open Updated Link')).toBeDefined();
    expect(button('Copy Updated Link')).toBeDefined();
  });

  it('shows a URL-based name and notice when the SharePoint title is unavailable', async () => {
    titleLookup.mockResolvedValueOnce({ title: 'Guide.docx', notice: 'Title unavailable; using the URL name.' });
    enterUrl(original);
    await submit();
    expect(container.querySelector('a[aria-describedby="urlFixerDragInstructions"]')?.textContent).toBe('Guide.docx');
    expect(container.textContent).toContain('Title unavailable; using the URL name.');
    expect(button('Open Updated Link')).toBeDefined();
  });

  it('does not restore cleared results when a title request finishes', async () => {
    let resolveTitle: ((value: { title: string }) => void) | undefined;
    titleLookup.mockReturnValueOnce(new Promise((resolve) => { resolveTitle = resolve; }));
    enterUrl(original);
    await submit();
    act(() => { Simulate.click(button('Clear')); });
    await act(async () => {
      if (!resolveTitle) throw new Error('Missing pending title request');
      resolveTitle({ title: 'Ultrasound Guide' });
    });
    expect(textarea().value).toBe('');
    expect(container.querySelector('a')).toBeNull();
    expect(button('Find Updated Link').disabled).toBe(false);
  });

  it('clears the input and result', async () => {
    enterUrl(original);
    await submit();
    act(() => { Simulate.click(button('Clear')); });
    expect(textarea().value).toBe('');
    expect(container.querySelector('a')).toBeNull();
  });

  it('logs the signed-in identity, exact input, usable result and UTC submission time', async () => {
    const providedUrl = `  ${original}\r\n`;
    enterUrl(providedUrl);
    const capturedInput = textarea().value;
    const before = Date.now();
    await submit();
    expect(log).toHaveBeenCalledTimes(1);
    expect(log).toHaveBeenCalledWith(expect.anything(), currentUser, {
      providedUrl: capturedInput,
      returnedUrl: fixed,
      outcome: 'Success',
      message: 'Updated link ready.',
      submittedAt: expect.any(String)
    });
    const loggedTime = log.mock.calls[0][2].submittedAt;
    expect(loggedTime).toMatch(/Z$/);
    expect(Date.parse(loggedTime)).toBeGreaterThanOrEqual(before);
    expect(Date.parse(loggedTime)).toBeLessThanOrEqual(Date.now());
    expect(lookup).toHaveBeenCalledWith(original, expect.anything());
  });

  it('logs warning results with their returned URL', async () => {
    lookup.mockResolvedValueOnce({ original, fixed, note: 'warn:Unable to verify destination.' });
    enterUrl(original);
    await submit();
    expect(log.mock.calls[0][2]).toMatchObject({
      returnedUrl: fixed, outcome: 'Warning', message: 'Unable to verify destination.'
    });
  });

  it.each(['ok', 'warn', 'err'])('includes the matched table only in the logged message for %s results', async (level) => {
    const message = 'Lookup outcome.';
    lookup.mockResolvedValueOnce({ original, fixed, note: `${level}:${message}`, lookupTableFileName: 'CS-fixer.csv' });
    enterUrl(original);
    await submit();
    expect(log.mock.calls[0][2]).toMatchObject({
      message: 'Lookup outcome. Lookup table: CS-fixer.csv.',
      returnedUrl: level === 'err' ? '' : fixed
    });
    expect(container.textContent).not.toContain('CS-fixer.csv');
  });

  it('logs rejected inputs without claiming the input is a returned URL', async () => {
    lookup.mockResolvedValueOnce({ original: 'invalid', fixed: 'invalid', note: 'err:Invalid URL.' });
    enterUrl('invalid');
    await submit();
    expect(log.mock.calls[0][2]).toMatchObject({
      providedUrl: 'invalid', returnedUrl: '', outcome: 'Error', message: 'Invalid URL.'
    });
  });

  it('logs empty submissions without looking up a URL', async () => {
    await submit();
    expect(lookup).not.toHaveBeenCalled();
    expect(log).toHaveBeenCalledTimes(1);
    expect(log.mock.calls[0][2]).toMatchObject({
      providedUrl: '', returnedUrl: '', outcome: 'Error', message: 'Paste a saved SharePoint link first.'
    });
    expect(container.querySelector('[role="alert"]')?.textContent).toBe('Paste a saved SharePoint link first.');
  });

  it('logs thrown lookup errors and keeps the error visible', async () => {
    lookup.mockRejectedValueOnce(new Error('Lookup request failed.'));
    enterUrl(original);
    await submit();
    expect(log.mock.calls[0][2]).toMatchObject({
      providedUrl: original, returnedUrl: '', outcome: 'Error', message: 'Lookup request failed.'
    });
    expect(container.querySelector('[role="alert"]')?.textContent).toBe('Lookup request failed.');
  });

  it('reports log failures only to diagnostics without changing the result or instructions', async () => {
    log.mockRejectedValueOnce(new Error('SharePoint denied access (403).'));
    enterUrl(original);
    await submit();
    expect(Log.error).toHaveBeenCalledWith('URLFixerSubmissionLog', expect.any(Error));
    expect(jest.mocked(Log.error).mock.calls[0][1].message).toContain('SharePoint denied access (403).');
    expect(container.querySelector('[role="alert"]')).toBeNull();
    expect(container.textContent).not.toMatch(/log|records|signed-in|403/i);
    expect(container.querySelector('#urlFixerInstructions')?.textContent).toContain('Copy the full URL');
    expect(container.querySelector('a')?.getAttribute('href')).toBe(fixed);
    expect(button('Copy Updated Link')).toBeDefined();
    await submit();
    expect(log).toHaveBeenCalledTimes(2);
    expect(container.querySelector('[role="alert"]')).toBeNull();
  });

  it('still logs a cleared lookup without restoring its result', async () => {
    let resolveLookup: ((result: IFixResult) => void) | undefined;
    lookup.mockReturnValueOnce(new Promise((resolve) => { resolveLookup = resolve; }));
    enterUrl(original);
    act(() => { Simulate.click(button('Find Updated Link')); });
    act(() => { Simulate.click(button('Clear')); });
    await act(async () => {
      if (!resolveLookup) throw new Error('Missing pending lookup');
      resolveLookup({ original, fixed, note: 'ok:Updated link ready.' });
    });
    expect(log).toHaveBeenCalledTimes(1);
    expect(log.mock.calls[0][2]).toMatchObject({ providedUrl: original, returnedUrl: fixed });
    expect(textarea().value).toBe('');
    expect(container.querySelector('a')).toBeNull();
  });

  it('does not delay a usable result and reports late logging failures only to diagnostics', async () => {
    let rejectLog: ((error: Error) => void) | undefined;
    log.mockReturnValueOnce(new Promise((resolve, reject) => { rejectLog = reject; }));
    enterUrl(original);
    await submit();
    expect(container.querySelector('a')?.getAttribute('href')).toBe(fixed);
    expect(button('Find Updated Link').disabled).toBe(false);
    await act(async () => {
      if (!rejectLog) throw new Error('Missing pending log write');
      rejectLog(new Error('Network unavailable.'));
    });
    expect(Log.error).toHaveBeenCalledWith('URLFixerSubmissionLog', expect.any(Error));
    expect(jest.mocked(Log.error).mock.calls[0][1].message).toContain('Network unavailable.');
    expect(container.querySelector('[role="alert"]')).toBeNull();
    expect(container.textContent).not.toMatch(/log|records|signed-in|Network unavailable/i);
    expect(container.querySelector('a')?.getAttribute('href')).toBe(fixed);
  });

  it('does not log copying, opening or clearing as new submissions', async () => {
    jest.spyOn(window, 'open').mockReturnValue(null);
    Object.defineProperty(navigator, 'clipboard', {
      configurable: true, value: { writeText: jest.fn().mockResolvedValue(undefined) }
    });
    enterUrl(original);
    await submit();
    act(() => { Simulate.click(button('Open Updated Link')); });
    await act(async () => { Simulate.click(button('Copy Updated Link')); });
    act(() => { Simulate.click(button('Clear')); });
    expect(log).toHaveBeenCalledTimes(1);
  });

  it('has no logging disclosure before or after a successful submission', async () => {
    expect(container.textContent).not.toMatch(/log|records|signed-in/i);
    enterUrl(original);
    await submit();
    expect(log).toHaveBeenCalledTimes(1);
    expect(container.textContent).not.toMatch(/log|records|signed-in/i);
    expect(Log.error).not.toHaveBeenCalled();
  });
});
