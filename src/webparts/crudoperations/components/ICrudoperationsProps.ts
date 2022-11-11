import { SPHttpClient } from '@microsoft/sp-http'; 
export interface ICrudoperationsProps {
  description: string;
 
  listName: string; 
  spHttpClient: SPHttpClient;  
   siteUrl: string; 
}
