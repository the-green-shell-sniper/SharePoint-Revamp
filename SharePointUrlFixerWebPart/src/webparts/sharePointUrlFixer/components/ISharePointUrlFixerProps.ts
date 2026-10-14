import { SPHttpClient } from '@microsoft/sp-http';
import { ISubmissionUser } from './submissionLog';

export interface ISharePointUrlFixerProps {
  spHttpClient: SPHttpClient;
  siteAbsoluteUrl: string;
  currentUser: ISubmissionUser;
}
