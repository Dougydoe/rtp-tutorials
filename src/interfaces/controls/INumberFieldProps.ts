import { IBaseFieldProps } from "./IBaseFieldProps";

export interface INumberFieldProps extends IBaseFieldProps {
    min?: number;
    max?: number;
}