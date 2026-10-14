import * as React from 'react';
import { Log } from '@microsoft/sp-core-library';
import styles from './SharePointUrlFixer.module.scss';
import { ISharePointUrlFixerProps } from './ISharePointUrlFixerProps';
import { IFixResult, LookupTableClient, lookupOrFixUrl } from './urlFixerLogic';
import { getFavoriteTitle, IFavoriteTitle } from './favoriteTitle';
import { logSubmission } from './submissionLog';

interface IResultRow extends IFixResult {
  level: 'ok' | 'warn' | 'err';
  note: string;
  favorite?: IFavoriteTitle;
}

interface ISharePointUrlFixerState {
  input: string;
  results: IResultRow[];
  status: { level: 'ok' | 'warn' | 'err'; message: string } | undefined;
  isChecking: boolean;
}

export default class SharePointUrlFixer extends React.Component<ISharePointUrlFixerProps, ISharePointUrlFixerState> {
  private readonly lookupClient: LookupTableClient;
  private requestId = 0;

  public constructor(props: ISharePointUrlFixerProps) {
    super(props);
    this.lookupClient = new LookupTableClient(props.spHttpClient, props.siteAbsoluteUrl);
    this.state = { input: '', results: [], status: undefined, isChecking: false };
  }

  private readonly onInputChange = (event: React.ChangeEvent<HTMLTextAreaElement>): void => {
    this.setState({ input: event.target.value });
  };

  private readonly onInputKeyDown = async (event: React.KeyboardEvent<HTMLTextAreaElement>): Promise<void> => {
    if (event.key === 'Enter' && !event.nativeEvent.isComposing) {
      event.preventDefault();
      await this.onConvert();
    }
  };

  private readonly onConvert = async (): Promise<void> => {
    if (this.state.isChecking) return;
    const requestId = ++this.requestId;
    const providedUrl = this.state.input;
    const submittedAt = new Date().toISOString();
    const input = providedUrl.trim().replace(/[\r\n]+/g, '');

    this.setState({
      results: [], isChecking: true,
      status: { level: 'warn', message: 'Finding the updated link...' }
    });

    let result: IFixResult;
    try {
      result = input
        ? await lookupOrFixUrl(input, this.lookupClient)
        : { original: input, fixed: '', note: 'err:Paste a saved SharePoint link first.' };
    } catch (error) {
      result = {
        original: input, fixed: '',
        note: `err:${error instanceof Error ? error.message : 'Could not load the SharePoint lookup tables.'}`
      };
    }

    const sepIndex = result.note.indexOf(':');
    const level = result.note.slice(0, sepIndex) as 'ok' | 'warn' | 'err';
    const note = result.note.slice(sepIndex + 1);
    const row: IResultRow = { ...result, level, note };
    try {
      if (level !== 'err') {
        row.favorite = await getFavoriteTitle(result.fixed);
      }
      if (requestId !== this.requestId) return;

      this.setState({
        results: input ? [row] : [],
        isChecking: false,
        status: { level, message: level === 'ok' ? 'Updated link ready.' : note }
      });
    } catch (error) {
      if (requestId !== this.requestId) return;
      this.setState({
        isChecking: false,
        status: { level: 'err', message: (error as Error).message || 'Could not load the SharePoint lookup tables.' }
      });
    } finally {
      logSubmission(this.props.spHttpClient, this.props.currentUser, {
        providedUrl,
        returnedUrl: level === 'err' ? '' : result.fixed,
        bookmarkButtonName: row.favorite?.title || '',
        outcome: level === 'ok' ? 'Success' : (level === 'warn' ? 'Warning' : 'Error'),
        message: result.lookupTableFileName ? `${note} Lookup table: ${result.lookupTableFileName}.` : note,
        submittedAt
      }).catch((error: unknown) => {
        const detail = error instanceof Error ? error.message : 'Unknown logging error.';
        Log.error('URLFixerSubmissionLog', new Error(`Submission log save was not confirmed. ${detail}`));
      });
    }
  };

  private readonly onClear = (): void => {
    this.requestId++;
    this.setState({ input: '', results: [], status: undefined, isChecking: false });
  };

  public componentWillUnmount(): void {
    this.requestId++;
  }

  private readonly onCopyResults = (): void => {
    const links = this.state.results.filter((r) => r.level !== 'err').map((r) => r.fixed);
    if (links.length === 0) return;
    navigator.clipboard.writeText(links.join('\n')).then(() => {
      this.setState({ status: { level: 'ok', message: 'Copied to clipboard!' } });
    }).catch(() => { /* clipboard permission denied - nothing actionable to do here */ });
  };

  private readonly statusClassName = (level: 'ok' | 'warn' | 'err'): string => {
    if (level === 'ok') return styles.statusOk;
    if (level === 'warn') return styles.statusWarn;
    return styles.statusErr;
  };

  public render(): React.ReactElement<ISharePointUrlFixerProps> {
    const { input, results, status, isChecking } = this.state;
    const usableResult = results.find((result) => result.level !== 'err');

    return (
      <div className={styles.sharePointUrlFixer}>
        <label className={styles.label} htmlFor="urlFixerInput">SharePoint page or file URL</label>
        <p id="urlFixerInstructions" className={styles.inputInstructions}>
          Copy the full URL from the browser&apos;s address bar, paste it here, then select <strong>Find Updated Link</strong> or press <kbd>Enter</kbd>.
          Use the updated result to open or copy the migrated page.
        </p>
        <textarea
          id="urlFixerInput"
          className={styles.textarea}
          aria-describedby="urlFixerInstructions"
          placeholder="https://whallc.sharepoint.com/CS/Ultrasound%20Department/..."
          value={input}
          onChange={this.onInputChange}
          onKeyDown={this.onInputKeyDown}
        />

        <div className={styles.primaryActions}>
          <button className={styles.button} disabled={isChecking} onClick={this.onConvert}>
            {isChecking ? 'Finding link...' : 'Find Updated Link'}
          </button>
          <button className={styles.secondaryButton} onClick={this.onClear}>Clear</button>
        </div>

        {status && (
          <div className={this.statusClassName(status.level)} role={status.level === 'err' ? 'alert' : 'status'}>{status.message}</div>
        )}

        {usableResult && (
          <div className={styles.resultBlock}>
            <div className={styles.resultBox}>
              <a className={styles.resultLink} href={usableResult.fixed} draggable={true} target="_blank" rel="noreferrer">{usableResult.fixed}</a>
            </div>
            <div className={styles.resultActions}>
              <button
                className={styles.openButton}
                title="Opens in a new tab; after opening, press ctrl+d on your keyboard to create a favorite of this link."
                onClick={() => window.open(usableResult.fixed, '_blank', 'noopener,noreferrer')}
              >
                Open Updated Link
              </button>
              <button className={styles.secondaryButton} onClick={this.onCopyResults}>Copy Updated Link</button>
            </div>
            <div className={styles.favoritesHelp}>
              <div className={styles.favoritesHeading}>Save as a bookmark or shortcut</div>
              <ul id="urlFixerDragInstructions" className={styles.favoritesInstructions}>
                <li><strong>Bookmark by dragging:</strong> Drag the button below onto Edge&apos;s bookmarks bar. Press <kbd>Ctrl+Shift+B</kbd> to show the bar.</li>
                <li><strong>Bookmark with the keyboard:</strong> Open the updated link, then press <kbd>Ctrl+D</kbd> in the new tab to save it as a bookmark.</li>
                <li><strong>Desktop shortcut:</strong> Drag the button below onto your desktop to create a shortcut.</li>
              </ul>
              <a
                className={styles.dragLink}
                href={usableResult.fixed}
                draggable={true}
                target="_blank"
                rel="noreferrer"
                title={usableResult.favorite?.title}
                aria-describedby="urlFixerDragInstructions"
              >
                {usableResult.favorite?.title}
              </a>
              {usableResult.favorite?.notice && (
                <p className={styles.favoritesInstructions} role="status">{usableResult.favorite.notice}</p>
              )}
            </div>
            <div className={styles.saveReminder}>You will need to manually delete any favorites, bookmarks, or shortcuts that have the old link.</div>
          </div>
        )}
      </div>
    );
  }
}
