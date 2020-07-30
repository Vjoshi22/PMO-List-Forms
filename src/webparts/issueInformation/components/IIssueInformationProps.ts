import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IIssueInformationProps {
  description: string;
  currentContext: WebPartContext;
  customGridRequired:string;
}
