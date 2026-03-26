import { ISiteUser } from "@pnp/sp/site-users";
import { IMyActiveTasksItems } from "./IMyActiveTasksItems";

export interface IMyActiveTasksState {
  ProjectListResult: IMyActiveTasksItems[];
  CurrentTab: string;
  openDialog: boolean;
  auditItems: any[];
}
