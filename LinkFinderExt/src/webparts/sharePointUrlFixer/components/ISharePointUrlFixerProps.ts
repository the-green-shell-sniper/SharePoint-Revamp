import { SPHttpClient } from '@microsoft/sp-http';
import { ISubmissionUser } from './submissionLog';

export interface ISharePointUrlFixerProps {
  spHttpClient: SPHttpClient;
  siteAbsoluteUrl: string;
  currentUser: ISubmissionUser;
  /** Old link passed in from the Favorites bar button; looked up automatically on load. */
  launchUrl?: string;
  launchError?: string;
  /** Published destination page used to build the Favorites bar button. */
  pageUrl?: string;
}
