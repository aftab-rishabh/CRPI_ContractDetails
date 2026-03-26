import {ISPList} from '../components/IViewContractsProps';

export interface IViewContractsState {

    ListData:ISPList[];
    ListAllData:ISPList[];
    CurrentTab: string;
    openDialog: boolean;
    auditItems: any[];
    hideDialog: boolean;
    itemID:string;
  }