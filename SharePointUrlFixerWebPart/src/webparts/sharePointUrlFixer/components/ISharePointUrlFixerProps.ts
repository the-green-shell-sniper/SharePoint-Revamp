import { SPHttpClient } from '@microsoft/sp-http';

export interface ISharePointUrlFixerProps {
  spHttpClient: SPHttpClient;
  siteAbsoluteUrl: string;
}
