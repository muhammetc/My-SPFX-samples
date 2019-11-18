
import { SPHttpClient } from '@microsoft/sp-http';

export interface ISpCrudProps {
  listName: string;
  spHttpClient: SPHttpClient;
  siteUrl: string;


}
