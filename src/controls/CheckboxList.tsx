import * as React from 'react';
import {CheckboxField} from './index';
import { ICheckboxListProps} from '../interfaces';
import { Label } from 'office-ui-fabric-react';
import error_styles from './FieldStyles.module.scss';
import styles from '../webparts/rtpForm/components/RtpForm.module.scss';

export class CheckboxList extends React.Component<ICheckboxListProps, {value:boolean}> {
    
    private updateFormData = (field:string, value:any):void => {
        // Shallow copy the value object, if there is one
            let newValue = {};
            const currentValue = this.props.formData[this.props.field];
            if (currentValue)  {
                for (var attr in currentValue) {
                    if (currentValue.hasOwnProperty(attr)) newValue[attr] = currentValue[attr];
                }
            }
            newValue[field] = value;
            this.props.onUpdate(this.props.field, newValue);
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

    private getErrorMessage = ():string => {
        if (!this.fieldRequired()) return "";
        const currentValue = this.props.formData[this.props.field];
        for (var attr in currentValue) {
            if (currentValue.hasOwnProperty(attr) && currentValue[attr])
            {
                return "";
            }
        }
        return "This field is required."; 
    }

    public render(): React.ReactElement<ICheckboxListProps> {

        let currentValue:any = this.props.formData[this.props.field];
        
        if (!currentValue) currentValue = {};
        
        let checked:boolean = false;
        let renderItems:any[] = [];

        for (let index = 0; index < this.props.options.length; index++) {
            const element = this.props.options[index];            
            checked = false;
            renderItems.push(              
                <div>
                    <div key={element.subfield} >
                    <div>
                        <br/>
                        <CheckboxField          
                            field={element.subfield} 
                            label={element.label} 
                            formData={currentValue} 
                            onUpdate={this.updateFormData}
                            validation={this.props.validation} 
                            onError={this.props.onError}
                        />
                    </div>
                    </div>
                </div>                               
            );
        }

        let footer_text = this.getErrorMessage();
        let fieldRequired:boolean = this.fieldRequired();

        return (
            <div className={styles.row}>
                <div className={styles.column}>
                    <Label required={fieldRequired}>{this.props.label}</Label>
                    {renderItems}
                    <span className={error_styles.dsErrorLabel}>{footer_text}</span>
                </div>                
            </div>
            
        );
    }
}