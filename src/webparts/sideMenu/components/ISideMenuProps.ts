import { WebPartContext } from "@microsoft/sp-webpart-base";
export interface ISideMenuProps {
  description: string;
  context: WebPartContext;
  // TeamList: string;
  adminGroup: string;
  hrGroup: string;
}
