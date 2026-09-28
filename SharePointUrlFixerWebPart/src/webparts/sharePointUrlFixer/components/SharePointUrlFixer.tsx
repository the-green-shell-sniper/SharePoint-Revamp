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

  private readonly onConvert = async (): Promise<void> => {
    const input = this.state.input.trim().replace(/[\r\n]+/g, '');

    if (!input) {
      this.setState({ results: [], status: { level: 'err', message: 'Paste a saved SharePoint link first.' } });
      return;
    }

    this.setState({ isChecking: true, status: { level: 'warn', message: 'Finding the updated link...' } });

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
      if (okCount) summary.push('Updated link ready');
      if (warnCount) summary.push('This link is already current');
      if (errCount) summary.push('This link needs help');

      this.setState({
        results,
        isChecking: false,
        status: { level: errCount ? 'warn' : 'ok', message: summary.join(', ') + '.' }
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

  private readonly friendlyNote = (result: IResultRow): string => {
    if (result.level === 'err') return result.note;
    if (result.note.toLowerCase().includes('lookup table')) return 'Updated link ready.';
    if (result.level === 'warn') return 'This link already looks current.';
    return 'Updated link ready.';
  };

  public render(): React.ReactElement<ISharePointUrlFixerProps> {
    const { input, results, status, isChecking } = this.state;

    return (
      <div className={styles.sharePointUrlFixer}>
        <h2 className={styles.title}>SharePoint URL Fixer</h2>
        <p className={styles.subtitle}>Paste a saved SharePoint link below and we’ll find the current link for you.</p>

        <label className={styles.label} htmlFor="urlFixerInput">Saved SharePoint link</label>
        <textarea
          id="urlFixerInput"
          className={styles.textarea}
          placeholder="Paste your saved SharePoint link here"
          value={input}
          onChange={this.onInputChange}
        />

        <div>
          <button className={styles.button} disabled={isChecking} onClick={this.onConvert}>
            {isChecking ? 'Finding link...' : 'Find updated link'}
          </button>
          <button className={styles.secondaryButton} onClick={this.onClear}>Clear</button>
        </div>

        {status && (
          <div className={this.statusClassName(status.level)}>{status.message}</div>
        )}

        {results.length > 0 && (
          <div className={styles.resultBlock}>
            <label className={styles.label}>Updated link</label>
            {results.map((result, index) => (
              <div className={styles.resultBox} key={index}>
                {result.level === 'err' ? (
                  <>
                    <div className={styles.errorMessage}>{result.note}</div>
                    <div className={styles.originalLink}>We couldn’t create an updated link for:</div>
                    <div className={styles.originalLink}>{result.original}</div>
                  </>
                ) : (
                  <>
                    <div className={styles.resultNote}>{this.friendlyNote(result)}</div>
                    <a className={styles.resultLink} href={result.fixed} target="_blank" rel="noreferrer">{result.fixed}</a>
                    <div className={styles.resultActions}>
                      <button className={styles.openButton} onClick={() => window.open(result.fixed, '_blank')}>Open updated link</button>
                      <span className={styles.ctrlDHint}>Save it to your favorites after it opens.</span>
                    </div>
                  </>
                )}
              </div>
            ))}
            <button className={styles.secondaryButton} style={{ marginLeft: 0, marginTop: 10 }} onClick={this.onCopyResults}>Copy updated link</button>
          </div>
        )}
      </div>
    );
  }
}
