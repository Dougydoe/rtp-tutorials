import { IDropDownOption } from "./IRtpFormState";

export interface IAgressoData {
    subProjectCodeDescription:string;
    costCentre:string; 
    projectCode:string; 
    purchasingDirectorate:string; 
    primaryFinancialApprover:string; 
    possibleSecondaryFinancialApprovers:IDropDownOption[];
    primaryBusinessManager:string; 
  }