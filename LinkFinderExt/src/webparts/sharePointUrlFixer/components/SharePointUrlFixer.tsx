import * as React from 'react';
import { Log } from '@microsoft/sp-core-library';
import styles from './SharePointUrlFixer.module.scss';
import { ISharePointUrlFixerProps } from './ISharePointUrlFixerProps';
import { LookupTableClient, lookupOrFixUrl } from './urlFixerLogic';
import { logSubmission } from './submissionLog';
import { buildQuickFindBookmarklet, navigateToDestination, QUICK_FIND_PAGE_URL } from './quickFind';

interface IQuickFindState {
  phase: 'setup' | 'checking' | 'ready' | 'warning' | 'error';
  message: string;
  destination?: string;
  seconds: number;
  paused: boolean;
}

const REDIRECT_SECONDS = 10;

export default class SharePointUrlFixer extends React.Component<ISharePointUrlFixerProps, IQuickFindState> {
  private timer: ReturnType<typeof setInterval> | undefined;
  private disposed = false;
  private readonly lookupClient: LookupTableClient;

  public constructor(props: ISharePointUrlFixerProps) {
    super(props);
    this.lookupClient = new LookupTableClient(props.spHttpClient, props.siteAbsoluteUrl);
    this.state = { phase: 'setup', message: '', seconds: REDIRECT_SECONDS, paused: false };
  }

  public componentDidMount(): void {
    if (this.props.launchError) {
      this.setState({ phase: 'error', message: this.props.launchError });
    } else if (this.props.launchUrl) {
      this.findDestination().catch((error: unknown) => {
        if (!this.disposed) {
          this.setState({ phase: 'error', message: error instanceof Error ? error.message : 'Could not find the new link.' });
        }
      });
    }
  }

  public componentWillUnmount(): void {
    this.disposed = true;
    this.stopTimer();
  }

  private stopTimer(): void {
    if (this.timer !== undefined) clearInterval(this.timer);
    this.timer = undefined;
  }

  private async findDestination(): Promise<void> {
    const providedUrl = this.props.launchUrl || '';
    const submittedAt = new Date().toISOString();
    this.setState({ phase: 'checking', message: 'Finding your new Central Park link...' });
    const source = new URL(providedUrl);
    if (source.hostname.toLowerCase() !== 'whallc.sharepoint.com' ||
        (source.protocol !== 'https:' && source.protocol !== 'http:')) {
      throw new Error('This tool only works with Central Park links on whallc.sharepoint.com.');
    }
    const result = await lookupOrFixUrl(providedUrl, this.lookupClient);
    if (this.disposed) return;
    const separator = result.note.indexOf(':');
    const level = result.note.slice(0, separator);
    const message = result.note.slice(separator + 1);
    logSubmission(this.props.spHttpClient, this.props.currentUser, {
      providedUrl, returnedUrl: level === 'err' ? '' : result.fixed, bookmarkButtonName: '',
      outcome: level === 'ok' ? 'Success' : level === 'warn' ? 'Warning' : 'Error',
      message: result.lookupTableFileName ? `${message} Lookup table: ${result.lookupTableFileName}.` : message,
      submittedAt
    }).catch((error: unknown) => {
      Log.error('QuickFindSubmissionLog', new Error(
        `Submission log save was not confirmed. ${error instanceof Error ? error.message : 'Unknown logging error.'}`
      ));
    });
    if (level === 'err') {
      this.setState({ phase: 'error', message });
      return;
    }
    const destination = new URL(result.fixed);
    const toolPage = new URL(this.props.pageUrl || QUICK_FIND_PAGE_URL);
    if (destination.protocol !== 'https:' || destination.hostname.toLowerCase() !== 'whallc.sharepoint.com') {
      throw new Error('The lookup did not return a trusted Central Park destination. Please contact the Help Desk.');
    }
    if (destination.origin === toolPage.origin &&
        destination.pathname.toLowerCase() === toolPage.pathname.toLowerCase()) {
      throw new Error('You are already on the Link Finder extension page. Open an old Central Park favorite before using Find New Link.');
    }
    const alreadyCurrent = level === 'warn' && /^\/sites\//i.test(source.pathname) &&
      destination.href === source.href && /^Already has \/sites|^This link already points/.test(message);
    if (level !== 'ok' && !alreadyCurrent) {
      this.setState({ phase: 'warning', message, destination: destination.href });
      return;
    }
    this.setState({
      phase: 'ready', destination: destination.href,
      message: alreadyCurrent ? 'This link is already up to date. Returning to your page.' : 'Your new link is ready.',
      seconds: REDIRECT_SECONDS, paused: false
    }, () => {
      this.startTimer();
    });
  }

  private startTimer(): void {
    this.stopTimer();
    this.timer = setInterval(() => {
      if (this.state.seconds <= 1) this.openDestination();
      else this.setState((state) => ({ seconds: state.seconds - 1 }));
    }, 1000);
  }

  private readonly openDestination = (): void => {
    this.stopTimer();
    if (this.state.destination) {
      try {
        navigateToDestination(this.state.destination);
      } catch (error: unknown) {
        this.setState({ phase: 'error', message: error instanceof Error ? error.message : 'Could not open the new link.' });
      }
    }
  };

  private readonly togglePause = (): void => {
    if (this.state.paused) {
      this.setState({ paused: false });
      this.startTimer();
    } else {
      this.stopTimer();
      this.setState({ paused: true });
    }
  };

  private readonly setButtonRef = (anchor: HTMLAnchorElement | null): void => {
    if (anchor) anchor.setAttribute('href', buildQuickFindBookmarklet(this.props.pageUrl || QUICK_FIND_PAGE_URL));
  };

  private readonly explainDrag = (event: React.MouseEvent<HTMLAnchorElement>): void => {
    event.preventDefault();
    this.setState({ message: 'Drag the Find New Link button to your Favorites bar instead of clicking it here.' });
  };

  public render(): React.ReactElement<ISharePointUrlFixerProps> {
    const { phase, message, destination, seconds, paused } = this.state;
    return (
      <main className={styles.sharePointUrlFixer}>
        {phase === 'setup' ? (
          <>
            <header className={styles.pageHeader}>
              <h1 className={styles.title}>Link Finder Extension</h1>
              <p className={styles.subtitle}>Has your usual Central Park link stopped working? Add a button to help find the new page.</p>
              <p className={styles.saveHint}><strong>For Central Park links only.</strong> It will not fix links to other websites.</p>
            </header>
            <section aria-labelledby="quickFindHowTo">
              <h2 className={styles.favoritesHeading} id="quickFindHowTo">Another way to use Link Finder</h2>
              <p className={styles.inputInstructions}>No copying or typing needed. This is an extra option, not a replacement for the
                original <a href="https://whallc.sharepoint.com/sites/CentralParkLinkFinder">Link Finder</a>.</p>
            </section>
            <section className={`${styles.favoritesHelp} ${styles.quickFindSetup}`} aria-labelledby="quickFindSetup">
              <h2 className={styles.favoritesHeading} id="quickFindSetup">1. Add the button once</h2>
              <ol className={styles.favoritesInstructions} id="quickFindInstructions">
                <li>Open this page in <strong>Microsoft Edge</strong>.</li>
                <li>Look just below the box at the top of Edge that shows the website address.
                  The row of saved links there is the <strong>Favorites bar</strong>.
                  If it is missing, hold <kbd>Ctrl</kbd> and <kbd>Shift</kbd>, press <kbd>B</kbd>, then let go.</li>
                <li>On the button below, hold your <strong>left mouse button</strong> down.
                  Move it up to that row, then let go. <strong>Drag it; do not just click it here.</strong></li>
              </ol>
              <div className={styles.dragLinkRow}>
                <span className={styles.dragPrompt}>Drag to your Favorites bar <span aria-hidden="true">&rarr;</span></span>
                <a ref={this.setButtonRef} className={styles.dragLink} draggable={true} onClick={this.explainDrag}
                  aria-describedby="quickFindInstructions" title="Find New Link">Find New Link</a>
              </div>
              <p className={styles.favoritesInstructions}><strong>Ready?</strong> You should see <strong>Find New Link</strong> in the row at the top of Edge.</p>
              {message && <p role="status">{message}</p>}
            </section>
            <section className={styles.favoritesHelp} aria-labelledby="quickFindUse">
              <h2 className={styles.favoritesHeading} id="quickFindUse">2. Use it when a Central Park link stops working</h2>
              <ol className={styles.favoritesInstructions}>
                <li>Open Central Park the way you normally do. When a Central Park link does not work, <strong>leave that page open</strong>.
                  An error page is OK.</li>
                <li>Click <strong>Find New Link</strong> in the row at the top of Edge.</li>
                <li>Wait. If we find the new page, it will open automatically.</li>
              </ol>
              <h2 className={styles.favoritesHeading}>3. Save the new page for next time</h2>
              <ol className={styles.favoritesInstructions}>
                <li>Check that the new page or file is the one you need.</li>
                <li>On that page, hold <kbd>Ctrl</kbd>, press <kbd>D</kbd>, then let go.</li>
                <li>Click <strong>Done</strong> in the box that appears. This saves a <strong>favorite</strong>: a link you can click next time.</li>
              </ol>
              <details className={styles.favoritesInstructions}>
                <summary>Need help, or want to remove an old saved link?</summary>
                <p>If no new page is found, use the original <a href="https://whallc.sharepoint.com/sites/CentralParkLinkFinder">Link Finder</a> or contact the Help Desk.</p>
                <p>Removing an old favorite is optional. First save and test the new one.
                  Then open Edge&apos;s favorites list (the star-with-lines button near the top-right),
                  right-click the old favorite, and choose <strong>Delete</strong>. If unsure, ask the Help Desk.</p>
                <p>This tool cannot change or remove your saved favorites or desktop shortcuts.</p>
              </details>
            </section>
          </>
        ) : (
          <section aria-label="Link lookup">
            <p className={phase === 'error' ? styles.statusErr : phase === 'warning' ? styles.statusWarn :
              phase === 'checking' ? styles.status : styles.statusOk}
              role={phase === 'error' ? 'alert' : 'status'}>{message}</p>
            {phase === 'ready' && (
              <>
                <div className={styles.saveHint}>
                  <strong>After the new page opens:</strong>
                  <ol className={styles.favoritesInstructions}>
                    <li>Check that it is the page or file you need.</li>
                    <li>Hold <kbd>Ctrl</kbd>, press <kbd>D</kbd>, then let go (<kbd>Ctrl+D</kbd>).</li>
                    <li>Click <strong>Done</strong> to save it for next time.</li>
                  </ol>
                  <strong>Wait for the new page before pressing Ctrl+D.</strong>
                </div>
                <p>{paused ? 'Automatic opening is paused.' : `Opening your page in ${seconds} seconds...`}</p>
                <div className={styles.primaryActions}>
                  <button className={styles.button} onClick={this.openDestination}>Open now</button>
                  <button className={styles.secondaryButton} onClick={this.togglePause}>{paused ? 'Resume' : 'Pause'}</button>
                </div>
              </>
            )}
            {phase === 'warning' && destination && (
              <p><a className={styles.resultLink} href={destination}>Open the unverified destination</a></p>
            )}
            {(phase === 'error' || phase === 'warning') && (
              <p>For help, use the <a href="https://whallc.sharepoint.com/sites/CentralParkLinkFinder">manual Link Finder</a> or contact the Help Desk.</p>
            )}
          </section>
        )}
      </main>
    );
  }
}
