jest.mock('@microsoft/sp-http', () => ({ SPHttpClient: jest.fn() }));
jest.mock('@microsoft/sp-core-library', () => ({ ServiceScope: { startNewRoot: jest.fn() } }));
jest.mock('./urlFixerLogic', () => ({
  LookupTableClient: jest.fn(),
  lookupOrFixUrl: jest.fn()
}));

import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { act, Simulate } from 'react-dom/test-utils';
import { ServiceScope } from '@microsoft/sp-core-library';
import { SPHttpClient } from '@microsoft/sp-http';
import SharePointUrlFixer from './SharePointUrlFixer';
import { IFixResult, lookupOrFixUrl } from './urlFixerLogic';

describe('URL fixer interactions', () => {
  let container: HTMLDivElement;
  const lookup = jest.mocked(lookupOrFixUrl);
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
    lookup.mockResolvedValue({ original, fixed, note: 'ok:Updated link ready.' });
    container = document.createElement('div');
    document.body.appendChild(container);
    act(() => {
      ReactDOM.render(
        <SharePointUrlFixer spHttpClient={new SPHttpClient(ServiceScope.startNewRoot())} siteAbsoluteUrl="https://whallc.sharepoint.com" />,
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

  it('submits with Enter instead of inserting a line break', async () => {
    enterUrl(original);
    const preventDefault = jest.fn();
    await act(async () => {
      Simulate.keyDown(textarea(), { key: 'Enter', preventDefault });
    });
    expect(preventDefault).toHaveBeenCalled();
    expect(lookup).toHaveBeenCalledTimes(1);
    expect(lookup).toHaveBeenCalledWith(original, expect.anything());
    expect(textarea().value).toBe(original);
  });

  it('does not submit during input method composition', () => {
    enterUrl(original);
    act(() => {
      Simulate.keyDown(textarea(), { key: 'Enter', nativeEvent: new KeyboardEvent('keydown', { isComposing: true }) });
    });
    expect(lookup).not.toHaveBeenCalled();
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
    expect(links[1].textContent).toBe('Drag Updated Link to Favorites Bar');
    expect(container.textContent).toContain('Ctrl+Shift+B');
    expect(container.textContent).toContain('Ctrl+D');
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
  });

  it('clears the input and result', async () => {
    enterUrl(original);
    await submit();
    act(() => { Simulate.click(button('Clear')); });
    expect(textarea().value).toBe('');
    expect(container.querySelector('a')).toBeNull();
  });
});
