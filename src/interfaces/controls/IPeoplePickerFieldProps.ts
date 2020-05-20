import {IWebPartContext} from '@microsoft/sp-webpart-base';
import { IBaseFieldProps } from './IBaseFieldProps';
  
  export interface IPeoplePickerFieldProps extends IBaseFieldProps {
    context: IWebPartContext;
    itemLimit: number;
    placeholder?: string;
  }