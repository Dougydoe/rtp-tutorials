import { PageContext } from "@microsoft/sp-page-context";
import { IWebPartContext } from '@microsoft/sp-webpart-base';
import { IFormContext } from "./IFormContext";

export interface IRtpFormProps {
    description: string;
    pageContext: PageContext;    
    context:IWebPartContext;    
    listName: string;
    formContext: IFormContext; 
}


