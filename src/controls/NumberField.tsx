import * as React from 'react';
import { INumberFieldProps } from '../interfaces';
import { TextField, Label, TooltipHost } from 'office-ui-fabric-react';
import styles from './FieldStyles.module.scss';

export class NumberField extends React.Component<INumberFieldProps, {}> {

    /**
     * @description does initial validation 
     * @fires when component mounts for this first time, after parent component's state.loading = false 
     */
    public componentDidMount() {
        // console.log('componentDidMount field: ' + this.props.field);
        let error:string = this.getFooterText();
        if (error) {
            this.props.onError(this.props.field);
        } else {
            this.props.onError(this.props.field, true);
        }
    }

    /**
     * @description this component should only update when this.state.formData changes
     * ! double check that this.state.formData is the only thing that should cause a re-render of this component 
     * ? what about dropdown options for choice fields?
     * @param nextProps to be received by the component
     * @returns true if component should update
     */
    public shouldComponentUpdate(nextProps):boolean {
        // console.log('shouldComponentUpdate field: ' + this.props.field);
        // if formData has changed, component should update
        if (this.props.formData != nextProps.formData) {
            return true;
        } else if (this.props.disabled != nextProps.disabled) {
            return true;
        }
        return false;
    }

    /**
     * @description validates the new input using the value stored in formData.field
     * @fires whenever this.shouldComponentUpdate() returns true
     * * This normally fires whenever this.setState() is called on the parent component
     */
    public componentDidUpdate() {
        // console.log('componentDidUpdate field: ' + this.props.field);        
        // do validation 
        let error:string = this.getFooterText();
        if (error) {
            this.props.onError(this.props.field);
        } else {
            this.props.onError(this.props.field, true);
        }
    }

    private fieldRequired = ():boolean => {
        if (this.props.validation && this.props.validation[this.props.field]) {
            const val:any = this.props.validation[this.props.field];
            if (val.validateWhen == null || val.validateWhen(this.props.formData, this.props.field)) {
                return val.required;
            }
        }
        return this.props.required;
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
     * @description checks to see if the field is required and empty
     * @returns an error message if invalid, else returns empty string if valid
     */
    private getFooterText = ():string => {
        let value = null;
        let footer_text:string = "";
        let fieldRequired = this.fieldRequired();

        //get field value
        if (this.props.formData[this.props.field]) {
            value = this.props.formData[this.props.field];

            if (fieldRequired && (value.length == 0)) {
                footer_text = "This field is required";                        
            }
            if (value.length > 0) {
                if (isNaN(Number(value))) {
                    footer_text = "Must be a valid number";                        
                }
            }
        } else if (fieldRequired) {
            footer_text = "This field is required";                        
        }
        return footer_text;
    }

    private handleBlur = (e:React.FocusEvent<HTMLInputElement>):void => {
        let value:string = e.currentTarget.value;
        if (value == "") value = undefined;
        this.props.onUpdate(this.props.field, value);
    }

    public render(): React.ReactElement<INumberFieldProps> {
        let tooltip:string = "";
        if (this.props.tooltip && this.props.tooltip[this.props.field]) tooltip = this.props.tooltip[this.props.field];
        const fieldRequired:boolean = this.fieldRequired(); 
        const footer_text:string = this.getFooterText();
        const fieldHidden:boolean = this.checkIfHidden();
        let fieldToDisplay:any = null;
        if (!fieldHidden) {
            fieldToDisplay = 
            <div>
                <TooltipHost content={tooltip}>
                <Label required={fieldRequired}>{this.props.label}</Label> 
                <TextField 
                    value={this.props.formData[this.props.field]}                
                    onBlur={this.handleBlur}  
                    disabled={this.props.disabled}
                />
                <span className={styles.dsErrorLabel}>{footer_text}</span>
                </TooltipHost>
            </div>;
        }
        return fieldToDisplay;
    }
}

