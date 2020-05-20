import * as React from 'react';
import { IDateFieldProps } from '../interfaces';
import { DatePicker, TooltipHost } from 'office-ui-fabric-react';
import styles from './FieldStyles.module.scss';

export class DateField extends React.Component<IDateFieldProps, {}> {
    
    public static defaultProps:Partial<IDateFieldProps> = {
        required: false,
        disabled: false
    };

    /**
     * @description does initial validation 
     * @fires when component mounts for this first time, after parent component's state.loading = false 
     */
    public componentDidMount() {
        // console.log('componentDidMount field: ' + this.props.field);
        // do initial validation 
        let dateField = this.prepareDateField();
        let error:string = dateField.footerText;
        if (error) {
            this.props.onError(this.props.field);
        } else if (!error) {
            this.props.onError(this.props.field, true);
        }
    }

    /**
     * @description this component should only update when this.state.formData changes
     * ! double check that this.state.formData is the only thing that should cause a re-render of this component 
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
        // do ongoing validation 
        let dateField = this.prepareDateField();
        let error:string = dateField.footerText;
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
     * @returns value of date field and footer text
     */
    private prepareDateField = ():{footerText: string, value: Date} => {
        let value = null;
        let footer_text:string = "";
        let fieldRequired = this.fieldRequired();

        //get field value
        if (this.props.formData[this.props.field]) {
            value = new Date(this.props.formData[this.props.field]);
        }
        // check if field is required and empty
        if (fieldRequired && !value) {
            footer_text = "This field is required";                        
        }
        return {
            footerText: footer_text,
            value: value
        };
    }

    // push field value to state
    private dateChange = (date:Date):void => {
        this.props.onUpdate(this.props.field, date.toDateString());
    }

    public render(): React.ReactElement<IDateFieldProps> {
        let tooltip:string = "";
        if (this.props.tooltip && this.props.tooltip[this.props.field]) tooltip = this.props.tooltip[this.props.field];
        const dateField = this.prepareDateField();
        const value = dateField.value;
        const footerText = dateField.footerText;
        const fieldHidden:boolean = this.checkIfHidden();
        let fieldToDisplay:any = null;
        if (!fieldHidden) {
            fieldToDisplay = 
            <div>
                <TooltipHost content={tooltip}>
                <DatePicker 
                    label={this.props.label}
                    value={value} 
                    onSelectDate={this.dateChange} 
                    disabled={this.props.disabled}
                    isRequired = {this.fieldRequired()}
                />
                <span className={styles.dsErrorLabel}>{footerText}</span>
                </TooltipHost>
            </div>;
        }
        return fieldToDisplay;
    }
}