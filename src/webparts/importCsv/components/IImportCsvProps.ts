import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPHttpClient } from '@microsoft/sp-http';

export interface IImportCsvProps {
  description: string;
  context: WebPartContext;
  spHttpClient: SPHttpClient;
  siteUrl: string;  
}
