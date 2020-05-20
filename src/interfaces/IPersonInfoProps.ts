import { IWebPartContext } from '@microsoft/sp-webpart-base';
import { IBaseFieldProps } from './controls/IBaseFieldProps';

export interface IPersonInfoProps extends IBaseFieldProps {
    context:IWebPartContext;
    itemLimit:number;
    dropDownOptions:any;
}