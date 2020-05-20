export interface IBaseFieldProps {
    required?:boolean;
    label?: string;
    disabled?:boolean;
    hidden?:boolean;
    field:string;
    formData: any;
    validation: any;
    onUpdate: (field:string, value:any, isAttachment?:boolean) => void;
    onError: (field:string, valid?:boolean) => void;
    tooltip?:string;
    calloutBody?:string;
}