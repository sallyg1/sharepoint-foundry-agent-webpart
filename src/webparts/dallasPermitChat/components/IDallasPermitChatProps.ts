import { AadHttpClientFactory } from '@microsoft/sp-http';

export interface IDallasPermitChatProps {
  httpClientFactory: AadHttpClientFactory;
  apiAppIdUri: string;
  functionEndpoint: string;
  userDisplayName: string;
}
