import { IChatboxItems } from "./IChatboxItems";

export interface IChatboxState {
  listResult: IChatboxItems[];
  Status: string;
  EditedNote: string;
  EditedId: number;
  ProjectId: number;
  isSalesPerson: boolean;
  isManagementPerson: boolean;
  hideDialog: boolean;
  DeletedId: number;
}
