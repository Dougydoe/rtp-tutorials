import * as React from 'react';
import styles from './FieldStyles.module.scss';
import { IChoiceFieldProps } from '../interfaces/index';
import { ChoiceGroup, IChoiceGroupOption, Label } from 'office-ui-fabric-react';
import { TooltipHost } from 'office-ui-fabric-react/lib';

export class ChoiceField extends React.Component<IChoiceFieldProps, {}> {
    
    /**
     * @description does initial validation 
     * @fires when component mounts for this first time, after parent component's state.loading = false 
     */
    public componentDidMount() {
        // console.log('componentDidMount field: ' + this.props.field);
        let error:string = this.getFooterText();
        if (error) {
            this.props.onError(this.props.field);
        } else if (!error) {
            this.props.onError(this.props.field, true);
        }
    }

    /**
     * @description this component should only update when this.state.formData changes
     * ! double check that this.state.formData and this.state.dropDownOptions are the only things that should cause a re-render of this component 
     * @param nextProps to be received by the component
     * @returns true if component should update
     */
    public shouldComponentUpdate(nextProps):boolean {
        // console.log('shouldComponentUpdate field: ' + this.props.field);
        // if formData OR dropDownOptions have changed, component should update
        if (this.props.formData != nextProps.formData || this.props.dropDownOptions != nextProps.dropDownOptions) {
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
        } else if (!error) {
            this.props.onError(this.props.field, true);
        }
    }
    
    private fieldRequired = ():boolean => {
        if (this.props.validation && this.props.validation[this.props.field]) {
            const val:any = this.props.validation[this.props.field];
            if (val.validateWhen == null || val.validateWhen(this.props.formData, this.props.field)) {
                const currentValue = this.props.formData[this.props.field];
                if(currentValue && currentValue !== undefined) {
                    return false;
                }
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
        }
        // check if field is required
        if (fieldRequired && !value) {
            footer_text = "This field is required";                        
        }
        return footer_text;
    }

    /**
     * @description if a field needs to be set as disabled on a per field level this can be used
     * @deprecated - fields only need to be disabled on a per section level
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

    private handleChange = (e:React.FormEvent<HTMLInputElement>, option:IChoiceGroupOption):void => {
        const value = option.key;
        this.props.onUpdate(this.props.field, value);
    }

    public render(): React.ReactElement<IChoiceFieldProps> {
        let tooltip:string = "";
        if (this.props.tooltip && this.props.tooltip[this.props.field]) tooltip = this.props.tooltip[this.props.field];
        const footer_text = this.getFooterText();
        const fieldHidden:boolean = this.checkIfHidden();
        let fieldToDisplay:any = null;
        if (!fieldHidden) {
            fieldToDisplay = 
            <div>
                <TooltipHost content={tooltip}>
                <Label required={this.fieldRequired()}>{this.props.label}</Label>
                <ChoiceGroup 
                    selectedKey={this.props.formData[this.props.field]}
                    options={this.props.dropDownOptions[this.props.field]}
                    onChange={this.handleChange}
                    disabled={this.props.disabled}
                />
                <span className={styles.dsErrorLabel}>{footer_text}</span>
                </TooltipHost>
            </div>;
        }
        return fieldToDisplay;
    }
}