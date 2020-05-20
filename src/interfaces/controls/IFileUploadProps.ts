import { IWebPartContext } from '@microsoft/sp-webpart-base';
import { IFormContext } from '../../interfaces';

export interface IFileUploadProps {
  context:IWebPartContext;  
  listName?:string;
  formContext:IFormContext;
  onUpdate: (field:string, value:any, isAttachment?:boolean) => void;
  onError: (field:string, valid?:boolean) => void;
  formData:any;
  label?:string;
  required?:boolean;
  validation?:any;
  field?:string;
}
