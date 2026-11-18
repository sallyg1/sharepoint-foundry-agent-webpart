import * as React from 'react';
import { createRoot, type Root } from 'react-dom/client';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'DallasPermitChatWebPartStrings';
import DallasPermitChat from './components/DallasPermitChat';
import { IDallasPermitChatProps } from './components/IDallasPermitChatProps';

export interface IDallasPermitChatWebPartProps {
  apiAppIdUri: string;
  functionEndpoint: string;
}

export default class DallasPermitChatWebPart extends BaseClientSideWebPart<IDallasPermitChatWebPartProps> {

  private _root: Root | undefined;

  public render(): void {
    const element: React.ReactElement<IDallasPermitChatProps> = React.createElement(
      DallasPermitChat,
      {
        httpClientFactory: this.context.aadHttpClientFactory,
        apiAppIdUri: this.properties.apiAppIdUri || 'api://<YOUR_ENTRA_API_APP_CLIENT_ID>',
        functionEndpoint: this.properties.functionEndpoint || 'https://<YOUR_FUNCTION_APP_NAME>.azurewebsites.net/api/chat',
        userDisplayName: this.context.pageContext.user.displayName
      }
    );

    if (!this._root) {
      this._root = createRoot(this.domElement);
    }

    this._root.render(element);
  }

  protected onDispose(): void {
    this._root?.unmount();
    this._root = undefined;
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('apiAppIdUri', {
                  label: 'Entra API App ID URI'
                }),
                PropertyPaneTextField('functionEndpoint', {
                  label: 'Azure Function Endpoint',
                  onGetErrorMessage: (value: string): string => {
                    const trimmed = (value || '').trim();
                    if (!trimmed) {
                      return '';
                    }
                    try {
                      const url = new URL(trimmed);
                      return url.protocol === 'https:' ? '' : 'Endpoint must start with https://';
                    } catch {
                      return 'Enter a valid https:// URL';
                    }
                  }
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
