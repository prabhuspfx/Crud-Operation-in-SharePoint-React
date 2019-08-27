import { SPHttpClient } from "@microsoft/sp-http";

export interface ICrudOperationProps {
  listName: string;
  spHttpClient: SPHttpClient;  
  siteUrl: string;  
}
