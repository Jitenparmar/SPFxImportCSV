import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPHttpClient } from '@microsoft/sp-http';
import { SPService } from "../../../Services/SPService";

export interface IImportCsvProps {
  SPServiceInstance:SPService;
  description: string;
  context: WebPartContext;
  spHttpClient: SPHttpClient;
  siteUrl: string;  
}
