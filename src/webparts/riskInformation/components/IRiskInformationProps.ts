import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IRiskInformationProps {
  description: string;
  currentContext: WebPartContext;
  customGridRequired: string;
  listGUID:string;
  ProjectMasterGUID:string;
  exceptionLogGUID:string;
}
