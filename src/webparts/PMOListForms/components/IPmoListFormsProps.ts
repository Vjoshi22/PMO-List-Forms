import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IPmoListFormsProps {
  description: string;
  currentContext: WebPartContext;
  customGridRequired: string;
  listGUID:string;
  exceptionLogGUID:string;
}
