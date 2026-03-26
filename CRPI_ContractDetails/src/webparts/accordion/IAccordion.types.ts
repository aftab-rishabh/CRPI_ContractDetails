import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IAccordionProps {
  defaultCollapsed?: boolean;
  title: string;
  className?: string;
  webURL: string;
  context: WebPartContext;
}

export interface IAccordionState {
  expanded: boolean;
  Items: any;
  HTML: any;
  OppID: any;
  files: any;
}
