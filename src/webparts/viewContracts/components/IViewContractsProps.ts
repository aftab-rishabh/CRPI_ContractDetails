import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IViewContractsProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  userEmailId: string;
  context: WebPartContext;
  webURL: string;
}

export interface ISPList {
  OpportunityNo: string;
  ProjectName: string;
  ClientOrganization: string;
  ProjectType: string;
  Stage: string;
  EstimatedHours: string;
  ID: string;
  CreatedBY: string;
  WebUrl: string;
  useremaildata: string;
  IndexNo: number;
  Created: Date;
  CreatedDisplay: string;
}
