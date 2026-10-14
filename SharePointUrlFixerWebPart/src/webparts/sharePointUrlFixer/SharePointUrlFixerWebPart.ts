import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { type IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import SharePointUrlFixer from './components/SharePointUrlFixer';
import { ISharePointUrlFixerProps } from './components/ISharePointUrlFixerProps';

export interface ISharePointUrlFixerWebPartProps {
  description: string;
}

export default class SharePointUrlFixerWebPart extends BaseClientSideWebPart<ISharePointUrlFixerWebPartProps> {

  public render(): void {
    const element: React.ReactElement<ISharePointUrlFixerProps> = React.createElement(
      SharePointUrlFixer,
      {
        spHttpClient: this.context.spHttpClient,
        siteAbsoluteUrl: this.context.pageContext.web.absoluteUrl,
        currentUser: {
          loginName: this.context.pageContext.user.loginName,
          displayName: this.context.pageContext.user.displayName,
          email: this.context.pageContext.user.email
        }
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: 'SharePoint URL Fixer'
          },
          groups: [
            {
              groupName: 'Settings',
              groupFields: []
            }
          ]
        }
      ]
    };
  }
}
