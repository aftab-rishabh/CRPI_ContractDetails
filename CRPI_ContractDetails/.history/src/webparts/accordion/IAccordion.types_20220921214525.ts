
export interface IAccordionProps {
  defaultCollapsed?: boolean;
  title: string;
  className?: string;
  webURL:string;
}

export interface IAccordionState {
  expanded: boolean;
  Items: any; 
  HTML: any;
}
