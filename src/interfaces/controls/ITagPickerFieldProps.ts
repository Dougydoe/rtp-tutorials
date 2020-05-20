import { ITag } from 'office-ui-fabric-react/lib';
import { IBaseFieldProps } from './IBaseFieldProps';

export interface ITagPickerFieldProps extends IBaseFieldProps {
    itemLimit?:number;
    suggestionsHeaderText?:string;
    noResultsFoundText?:string;
}