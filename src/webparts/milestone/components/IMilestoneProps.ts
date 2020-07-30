import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IMilestoneProps {
  description: string;
  currentContext: WebPartContext;
  customGridRequired: string;
}
