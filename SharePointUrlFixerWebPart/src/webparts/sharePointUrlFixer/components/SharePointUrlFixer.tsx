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
    const lines = this.state.input.split('\n').map((line) => line.trim()).filter((line) => line.length > 0);

    if (lines.length === 0) {
      this.setState({ results: [], status: { level: 'err', message: 'Please paste at least one URL.' } });
      return;
    }

    this.setState({ isChecking: true, status: { level: 'warn', message: 'Checking SharePoint lookup tables...' } });

    try {
      const rawResults = await Promise.all(lines.map((line) => lookupOrFixUrl(line, this.lookupClient)));
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
      if (okCount) summary.push(okCount + ' fixed');
      if (warnCount) summary.push(warnCount + ' already correct');
      if (errCount) summary.push(errCount + ' not recognized');

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

  public render(): React.ReactElement<ISharePointUrlFixerProps> {
    const { input, results, status, isChecking } = this.state;

    return (
      <div className={styles.sharePointUrlFixer}>
        <p className={styles.subtitle}>Paste one or more whallc.sharepoint.com links below.</p>

        <label className={styles.label} htmlFor="urlFixerInput">URL(s) — one per line</label>
        <textarea
          id="urlFixerInput"
          className={styles.textarea}
          placeholder="https://whallc.sharepoint.com/CS/Ultrasound%20Department/..."
          value={input}
          onChange={this.onInputChange}
        />

        <div>
          <button className={styles.button} disabled={isChecking} onClick={this.onConvert}>
            {isChecking ? 'Checking...' : 'Fix URL(s)'}
          </button>
          <button className={styles.secondaryButton} onClick={this.onClear}>Clear</button>
        </div>

        {status && (
          <div className={this.statusClassName(status.level)}>{status.message}</div>
        )}

        {results.length > 0 && (
          <div className={styles.resultBlock}>
            <label className={styles.label}>Result</label>
            {results.map((result, index) => (
              <div className={styles.resultBox} key={index}>
                {result.level === 'err' ? (
                  <>
                    <div style={{ color: '#b3261e', fontWeight: 600, fontSize: 11 }}>{result.note}</div>
                    <div style={{ color: '#888', marginTop: 4 }}>Original (unchanged): {result.original}</div>
                  </>
                ) : (
                  <>
                    <div style={{ color: '#888', fontSize: 11, marginBottom: 4 }}>{result.note}</div>
                    <a className={styles.resultLink} href={result.fixed} target="_blank" rel="noreferrer">{result.fixed}</a>
                    <div className={styles.resultActions}>
                      <button className={styles.openButton} onClick={() => window.open(result.fixed, '_blank')}>Open in new tab</button>
                      <span className={styles.ctrlDHint}>Once it opens, press <kbd>Ctrl+D</kbd> to save it as an Edge favorite.</span>
                    </div>
                  </>
                )}
              </div>
            ))}
            <button className={styles.secondaryButton} style={{ marginLeft: 0, marginTop: 10 }} onClick={this.onCopyResults}>Copy all</button>
          </div>
        )}
      </div>
    );
  }
}
