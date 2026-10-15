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

export default class SharePointUrlFixer extends React.Component<ISharePointUrlFixerProps, IQuickFindState> {
  private timer: ReturnType<typeof setInterval> | undefined;
  private disposed = false;
  private readonly lookupClient: LookupTableClient;

  public constructor(props: ISharePointUrlFixerProps) {
    super(props);
    this.lookupClient = new LookupTableClient(props.spHttpClient, props.siteAbsoluteUrl);
    this.state = { phase: 'setup', message: '', seconds: 5, paused: false };
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
      seconds: 5, paused: false
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
        <header className={styles.pageHeader}>
          <h1 className={styles.title}>Find New Link for Central Park</h1>
          <p className={styles.subtitle}>Keep your saved links working after the Central Park migration.</p>
        </header>
        {phase === 'setup' ? (
          <>
            <section aria-labelledby="quickFindHowTo">
              <h2 className={styles.favoritesHeading} id="quickFindHowTo">How it works</h2>
              <ol className={styles.favoritesInstructions}>
                <li>Open the old Central Park favorite or shortcut you want to update.</li>
                <li>Select <strong>Find New Link</strong> on your Edge Favorites bar. We will find and open the new page automatically.</li>
                <li>Confirm the new page, then press <kbd>Ctrl+D</kbd> to save it as a favorite.</li>
              </ol>
              <p className={styles.inputInstructions}>Works only with Central Park links on whallc.sharepoint.com. Other websites are not supported.</p>
            </section>
            <section className={`${styles.favoritesHelp} ${styles.quickFindSetup}`} aria-labelledby="quickFindSetup">
              <h2 className={styles.favoritesHeading} id="quickFindSetup">One-time setup: save Find New Link</h2>
              <p className={styles.favoritesInstructions} id="quickFindInstructions">
                Drag the button below onto your Edge Favorites bar. If the bar is hidden, press <kbd>Ctrl+Shift+B</kbd> to show it.
                Keep this button for whenever an old Central Park link stops working.
              </p>
              <div className={styles.dragLinkRow}>
                <span className={styles.dragPrompt}>Drag to your Favorites bar <span aria-hidden="true">&rarr;</span></span>
                <a ref={this.setButtonRef} className={styles.dragLink} draggable={true} onClick={this.explainDrag}
                  aria-describedby="quickFindInstructions" title="Find New Link">Find New Link</a>
              </div>
              {message && <p role="status">{message}</p>}
            </section>
            <p className={styles.inputInstructions}>This tool cannot replace or delete existing favorites. Remove the old favorite yourself after saving the new one.</p>
          </>
        ) : (
          <section aria-label="Link lookup">
            <p className={phase === 'error' ? styles.statusErr : phase === 'warning' ? styles.statusWarn :
              phase === 'checking' ? styles.status : styles.statusOk}
              role={phase === 'error' ? 'alert' : 'status'}>{message}</p>
            {phase === 'ready' && (
              <>
                <p className={styles.saveHint}><strong>Save your new page:</strong> Once it opens, press <kbd>Ctrl+D</kbd> on your keyboard to save it as an Edge favorite.
                  Save the destination page, not this Link Finder page.</p>
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
