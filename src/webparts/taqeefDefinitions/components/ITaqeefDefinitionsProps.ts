import { WebPartContext } from "@microsoft/sp-webpart-base";
export interface ITaqeefDefinitionsProps {
  description: string;
  siteurl: string;
  context: WebPartContext;
  graphClient: any;
  absoluteURL:string;
}
