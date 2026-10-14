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
  isDragOver: boolean;
  results: IResultRow[];
  status: { level: 'ok' | 'warn' | 'err'; message: string } | undefined;
  isChecking: boolean;
}

export default class SharePointUrlFixer extends React.Component<ISharePointUrlFixerProps, ISharePointUrlFixerState> {
  private readonly lookupClient: LookupTableClient;
  private readonly inputRef = React.createRef<HTMLTextAreaElement>();
  private requestId = 0;

  public constructor(props: ISharePointUrlFixerProps) {
    super(props);
    this.lookupClient = new LookupTableClient(props.spHttpClient, props.siteAbsoluteUrl);
    this.state = { input: '', isDragOver: false, results: [], status: undefined, isChecking: false };
  }

  private readonly onInputDragOver = (event: React.DragEvent<HTMLTextAreaElement>): void => {
    event.preventDefault();
    const { effectAllowed } = event.dataTransfer;
    if (effectAllowed === 'copy' || effectAllowed === 'copyMove') {
      event.dataTransfer.dropEffect = 'copy';
    } else if (effectAllowed === 'move') {
      event.dataTransfer.dropEffect = 'move';
    } else if (effectAllowed === 'none') {
      event.dataTransfer.dropEffect = 'none';
    } else {
      event.dataTransfer.dropEffect = 'link';
    }
    if (!this.state.isDragOver) this.setState({ isDragOver: true });
  };

  private readonly onInputDragLeave = (): void => {
    this.setState({ isDragOver: false });
  };

  private readonly findWebUrl = (value: string): string | undefined => {
    for (const line of value.split(/\r?\n/)) {
      const candidate = line.replace(/^URL=/i, '').trim();
      try {
        const url = new URL(candidate);
        if (url.protocol === 'http:' || url.protocol === 'https:') return candidate;
      } catch {
        // Ignore non-URL drag data and try the next line.
      }
    }
    // Custom formats (for example browser favorites) may embed the URL in JSON or other text.
    const match = value.match(/https?:\/\/[^\s"'<>\\]+/i);
    return match ? match[0] : undefined;
  };

  private readonly getDroppedUrl = (dataTransfer: Pick<DataTransfer, 'getData'> & Partial<Pick<DataTransfer, 'types'>>): string | undefined => {
    for (const type of ['text/uri-list', 'text/x-moz-url', 'text/plain']) {
      const url = this.findWebUrl(dataTransfer.getData(type));
      if (url) return url;
    }

    const html = dataTransfer.getData('text/html');
    if (html) {
      const anchor = new DOMParser().parseFromString(html, 'text/html').querySelector<HTMLAnchorElement>('a[href]');
      if (anchor) {
        try {
          const url = new URL(anchor.href);
          if (url.protocol === 'http:' || url.protocol === 'https:') return anchor.href;
        } catch {
          // Ignore malformed link markup.
        }
      }
    }

    for (const type of Array.from(dataTransfer.types ?? [])) {
      if (type === 'Files' || type === 'text/html') continue;
      const url = this.findWebUrl(dataTransfer.getData(type));
      if (url) return url;
    }
    return undefined;
  };

  private readonly getDroppedFiles = (dataTransfer: DataTransfer): File[] => {
    const files = Array.from(dataTransfer.files ?? []);
    for (const item of Array.from(dataTransfer.items ?? [])) {
      if (item.kind !== 'file') continue;
      const file = item.getAsFile();
      if (file && files.indexOf(file) === -1 && !files.some((existing) => existing.name === file.name && existing.size === file.size)) {
        files.push(file);
      }
    }
    return files;
  };

  private readonly onInputDrop = (event: React.DragEvent<HTMLTextAreaElement>): void => {
    event.preventDefault();
    this.setState({ isDragOver: false });

    const droppedUrl = this.getDroppedUrl(event.dataTransfer);
    if (droppedUrl) {
      this.setState({ input: droppedUrl });
      return;
    }

    const files = this.getDroppedFiles(event.dataTransfer);
    const shortcutFile = files.find((file) => /\.url$/i.test(file.name)) ?? files.find((file) => file.size > 0 && file.size < 65536);
    if (shortcutFile) {
      shortcutFile.text().then((content) => {
        const shortcutUrl = this.findWebUrl(content);
        if (shortcutUrl) {
          this.setState({ input: shortcutUrl });
        } else {
          this.setState({ status: { level: 'err', message: 'The dropped shortcut does not contain a valid web link.' } });
        }
      }).catch((error: unknown) => {
        const detail = error instanceof Error ? error.message : 'Unknown file read error.';
        this.setState({ status: { level: 'err', message: `Could not read the dropped shortcut. ${detail}` } });
      });
      return;
    }

    const types = Array.from(event.dataTransfer.types ?? []).join(', ') || 'none';
    this.setState({
      status: {
        level: 'err',
        message: `Could not read a link from the dropped item. Try copying and pasting the link instead. (Drop formats: ${types})`
      }
    });
  };

  private readonly onInputChange = (event: React.ChangeEvent<HTMLTextAreaElement>): void => {
    this.setState({ input: event.target.value });
  };

  private readonly resizeInput = (): void => {
    const input = this.inputRef.current;
    if (!input) return;
    input.style.height = 'auto';
    input.style.height = `${input.scrollHeight}px`;
  };

  public componentDidMount(): void {
    this.resizeInput();
  }

  public componentDidUpdate(): void {
    this.resizeInput();
  }

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
        status: { level, message: level === 'ok' ? 'Your new link is ready' : note }
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
    this.setState({ input: '', isDragOver: false, results: [], status: undefined, isChecking: false });
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
    const { input, isDragOver, results, status, isChecking } = this.state;
    const usableResult = results.find((result) => result.level !== 'err');
    const showFollowUpSteps = usableResult?.level === 'ok';

    return (
      <div className={styles.sharePointUrlFixer}>
        <textarea
          ref={this.inputRef}
          id="urlFixerInput"
          className={isDragOver ? `${styles.textarea} ${styles.dragOver}` : styles.textarea}
          aria-label="SharePoint page or file URL"
          title="Paste a URL or drag a favorite or shortcut here"
          placeholder="https://whallc.sharepoint.com/CS/Ultrasound%20Department/..."
          value={input}
          onChange={this.onInputChange}
          onKeyDown={this.onInputKeyDown}
          onDragOver={this.onInputDragOver}
          onDragLeave={this.onInputDragLeave}
          onDrop={this.onInputDrop}
        />

        <div className={styles.primaryActions}>
          <button className={styles.button} disabled={isChecking} onClick={this.onConvert}>
            {isChecking ? 'Finding link...' : 'Find New Link'}
          </button>
          <button className={styles.secondaryButton} onClick={this.onClear}>Clear</button>
        </div>
        <br />

        {status && (
          <div className={this.statusClassName(status.level)} role={status.level === 'err' ? 'alert' : 'status'}>{status.message}</div>
        )}

        {usableResult && (
          <div className={styles.resultBlock}>
            {showFollowUpSteps && (
              <>
                <div className={styles.resultStep}>2. Open and confirm your new link</div>
                <div className={styles.resultInstruction}>Select &quot;Open New Link&quot; and check that it takes you to the page or file you need.</div>
              </>
            )}
            <div className={styles.resultBox}>
              <a className={styles.resultLink} href={usableResult.fixed} draggable={true} target="_blank" rel="noreferrer">{usableResult.fixed}</a>
            </div>
            <div className={styles.resultActions}>
              <button
                className={styles.openButton}
                title="Opens in a new tab"
                onClick={() => window.open(usableResult.fixed, '_blank', 'noopener,noreferrer')}
              >
                Open New Link
              </button>
              <button className={styles.secondaryButton} onClick={this.onCopyResults}>Copy New Link</button>
            </div>
            {showFollowUpSteps && (
              <>
                <br />
                <div className={styles.favoritesHelp}>
                  <div className={styles.favoritesHeading}>3. Save your new link as a favorite or shortcut</div>
                  <div id="urlFixerDragInstructions">
                    <ul className={`${styles.favoritesInstructions} ${styles.favoritesList}`}>
                      <li><strong>Drag &amp; Drop favorite or shortcut:</strong> Drag the button below to your desktop to create a shortcut and/or to your Edge favorites bar to create a favorite (press <kbd>Ctrl+Shift+B</kbd> on your keyboard if you do not see the bar).</li>
                      <li><strong>Create a favorite with the keyboard:</strong> While on the new link tab, press <kbd>Ctrl+D</kbd> on your keyboard to save it as a favorite in Edge.</li>
                    </ul>
                  </div>
                  <div className={styles.dragLinkRow}>
                    <span className={styles.dragPrompt}>Drag &amp; Drop Me <span aria-hidden="true">→</span></span>
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
                  </div>
                  {usableResult.favorite?.notice && (
                    <p className={styles.favoritesInstructions} role="status">{usableResult.favorite.notice}</p>
                  )}
                </div>
                <div className={styles.saveReminder}><strong>This tool cannot remove or change any of your existing saved favorites or shortcuts. You will need to manually remove any favorites or shortcuts to that no longer work.</strong></div>
              </>
            )}
          </div>
        )}
      </div>
    );
  }
}
