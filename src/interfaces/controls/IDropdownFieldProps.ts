import {IDropdownOption} from 'office-ui-fabric-react/lib';
import { IBaseFieldProps } from './IBaseFieldProps';

export interface IDropdownFieldProps extends IBaseFieldProps {
    placeHolder?:string;
    dropDownOptions: any;
}