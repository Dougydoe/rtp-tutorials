import { IWebPartContext } from '@microsoft/sp-webpart-base';
import { IFormContext } from './index';

export interface IAttachmentsProps {
    context:IWebPartContext;
    listName?: string;
    formContext: IFormContext;
    onUpdate: (field:string, value:any, isAttachment?:boolean) => void;
    onError: (field:string, valid?:boolean) => void;
    formData:any;
}