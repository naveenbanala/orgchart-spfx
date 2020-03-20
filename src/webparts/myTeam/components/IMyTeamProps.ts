import { WebPartContext } from "@microsoft/sp-webpart-base";


export interface IMyTeamProps {
  description: string;
  context: WebPartContext;
  checkboxPeers: boolean;
  checkboxManagers: boolean;
  checkboxDirectReports: Boolean;
}
