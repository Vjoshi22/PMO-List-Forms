import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface ITileNavigaitonPmoProps {
  description: string;
  currentContext:  WebPartContext;
  lists: string | string[];
  tileName: string;
}
