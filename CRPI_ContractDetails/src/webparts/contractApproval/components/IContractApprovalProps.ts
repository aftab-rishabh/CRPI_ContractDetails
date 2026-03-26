import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IContractApprovalProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  context: WebPartContext;
  webURL:string;
  defaultCollapsed?: boolean;
  className?: string;
}
