import * as React from 'react';
import styles from './SharePointUrlFixer.module.scss';
import { ISharePointUrlFixerProps } from './ISharePointUrlFixerProps';
import { IFixResult, LookupTableClient, lookupOrFixUrl } from './urlFixerLogic';

interface IResultRow extends IFixResult {
  level: 'ok' | 'warn' | 'err';
  note: string;
}

interface ISharePointUrlFixerState {
  input: string;
  results: IResultRow[];
  status: { level: 'ok' | 'warn' | 'err'; message: string } | undefined;
  isChecking: boolean;
}

export default class SharePointUrlFixer extends React.Component<ISharePointUrlFixerProps, ISharePointUrlFixerState> {
  private readonly lookupClient: LookupTableClient;

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
    const input = this.state.input.trim().replace(/[\r\n]+/g, '');

    if (!input) {
      this.setState({ results: [], status: { level: 'err', message: 'Paste a saved SharePoint link first.' } });
      return;
    }

    this.setState({ results: [], isChecking: true, status: { level: 'warn', message: 'Finding the updated link...' } });

    try {
      const rawResults = [await lookupOrFixUrl(input, this.lookupClient)];
      let okCount = 0, warnCount = 0, errCount = 0;
      const results: IResultRow[] = rawResults.map((result) => {
        const sepIndex = result.note.indexOf(':');
        const level = result.note.slice(0, sepIndex) as 'ok' | 'warn' | 'err';
        const note = result.note.slice(sepIndex + 1);
        if (level === 'ok') okCount++;
        else if (level === 'warn') warnCount++;
        else errCount++;
        return { ...result, level, note };
      });

      const summary: string[] = [];
      const firstResult = results[0];
      if (errCount) summary.push(firstResult.note);
      else if (warnCount) summary.push('This link is already current.');
      else if (okCount) summary.push('Updated link ready.');

      this.setState({
        results,
        isChecking: false,
        status: { level: errCount ? 'err' : (warnCount ? 'warn' : 'ok'), message: summary.join('') }
      });
    } catch (error) {
      this.setState({
        isChecking: false,
        status: { level: 'err', message: (error as Error).message || 'Could not load the SharePoint lookup tables.' }
      });
    }
  };

  private readonly onClear = (): void => {
    this.setState({ input: '', results: [], status: undefined });
  };

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
              <div className={styles.favoritesHeading}>Save to Edge favorites</div>
              <p id="urlFixerDragInstructions" className={styles.favoritesInstructions}>
                Drag the link below to Edge&apos;s favorites bar (<kbd>Ctrl+Shift+B</kbd> to show it).
                Or open the updated link and press <kbd>Ctrl+D</kbd> to save it.
              </p>
              <a
                className={styles.dragLink}
                href={usableResult.fixed}
                draggable={true}
                target="_blank"
                rel="noreferrer"
                title="Updated SharePoint Link"
                aria-describedby="urlFixerDragInstructions"
              >
                Drag Updated Link to Favorites Bar
              </a>
            </div>
            <div className={styles.saveReminder}>You will need to manually delete any favorites, bookmarks, or shortcuts that have the old link.</div>
          </div>
        )}
      </div>
    );
  }
}
