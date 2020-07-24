import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IHomePageProps {
  description: string;
  currentContext: WebPartContext;
  tileName: string;
}
