import * as React from 'react';
import { IToggleFieldProps} from '../interfaces';
import { Toggle, Label, TooltipHost } from 'office-ui-fabric-react';
// import styles from './FieldStyles.module.scss';

export class ToggleField extends React.Component<IToggleFieldProps, {}> {     
    
    // pushes field value to state
    private handleChange = (checked:boolean):void => {        
        // push value to state.formData
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

    public render(): React.ReactElement<IToggleFieldProps> {                   
        let tooltip:string = "";
        if (this.props.tooltip && this.props.tooltip[this.props.field]) tooltip = this.props.tooltip[this.props.field];
        const fieldHidden = this.checkIfHidden();
        let fieldToDisplay:any = null;
        if (!fieldHidden) {
            fieldToDisplay = 
            <div>
            <TooltipHost content={tooltip}>
            <Label>{this.props.label}</Label> 
            <Toggle 
                defaultChecked={this.props.formData[this.props.field]}
                onText="Compliant"
                offText="Non compliant"
                onChanged={this.handleChange}  
            />
            </TooltipHost>
            </div>;
        }
        return fieldToDisplay;
    }
}