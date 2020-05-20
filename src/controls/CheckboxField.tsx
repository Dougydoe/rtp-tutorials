import * as React from 'react';
import { ICheckboxFieldProps } from '../interfaces';
import { Checkbox, TooltipHost } from 'office-ui-fabric-react';

export class CheckboxField extends React.Component<ICheckboxFieldProps, {}> {
    
    public static defaultProps:Partial<ICheckboxFieldProps> = {
        disabled: false,
        hidden: false,
    };

    private checkboxChanged = (ev: React.FormEvent<HTMLElement>, checked: boolean): void => {
        this.props.onUpdate(this.props.field, checked);
    }

    private checkIfHidden = ():boolean => {

        // check if section should be hidden
        if (this.props.validation && this.props.validation['hideSection']) {
            const hideSection = this.props.validation['hideSection'];
            if (hideSection(this.props.formData)) {
                return true;
            } else if (this.props.validation[this.props.field]) {
                // check if field should be hidden
                const val:any = this.props.validation[this.props.field];
                if (val.hideWhen == null || val.hideWhen(this.props.formData, this.props.field)) {
                    return val.hidden;
                }
            }
        } else if (this.props.validation && this.props.validation[this.props.field]) {
            const val:any = this.props.validation[this.props.field];
            if (val.hideWhen == null || val.hideWhen(this.props.formData, this.props.field)) {
                return val.hidden;
            }
        }
        return this.props.hidden;
    }

    /**
     * @description if a field needs to be set as disabled on a per field level this can be used
     * @deprecated fields only need to be disabled on a per section level
     */
    /* private fieldDisabled = ():boolean => {
        if (this.props.validation && this.props.validation[this.props.field]) {
            const val:any = this.props.validation[this.props.field];
            if (val['disabledWhen'] && val.disabledWhen(this.props.formData)) {                                             
                return true;
            }
        }
        return this.props.disabled;
    } */

    public render(): React.ReactElement<ICheckboxFieldProps> {
        let tooltip:string = "";
        if (this.props.tooltip && this.props.tooltip[this.props.field]) tooltip = this.props.tooltip[this.props.field];
        const fieldHidden:boolean = this.checkIfHidden();
        let fieldToDisplay:any = null;
        if (!fieldHidden) {
            fieldToDisplay = 
            <div>
                <TooltipHost content={tooltip}>
                <Checkbox 
                    label={this.props.label}
                    checked={this.props.formData[this.props.field]}                
                    onChange={this.checkboxChanged} 
                    disabled={this.props.disabled}
                />
                </TooltipHost>
            </div>;
        } 
        return fieldToDisplay;
        
    }

 
}