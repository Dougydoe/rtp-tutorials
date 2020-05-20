import { IWebPartContext } from '@microsoft/sp-webpart-base';
import { IFormContext } from './index';

export interface IFormSectionProps {
    context?:IWebPartContext;
    formData: any;
    onUpdate: (field:string, value:any, isAttachment?:boolean) => void;
    validation?: any;
    listName?:string;
    formContext?:IFormContext;
    dropDownOptions:any;
    onError: (field:string) => void;
    disabled?:boolean;
}

