import { IBaseFieldProps } from "./IBaseFieldProps";

export interface ITextBoxFieldProps extends IBaseFieldProps {
    multiline?: boolean;
    maxlength?:number;
    prefix?: string;
}

