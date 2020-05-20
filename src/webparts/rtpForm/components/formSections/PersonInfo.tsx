import * as React from 'react';
import { TextField } from 'office-ui-fabric-react/';
import styles from '../RtpForm.module.scss';
import { PeoplePickerField } from '../../../../controls/PeoplePickerField';
import { IPersonInfoProps } from '../../../../interfaces';
import { TextBoxField, DropdownField } from '../../../../controls';


export class PersonInfo extends React.Component<IPersonInfoProps, {}> {

  public static defaultProps:Partial<IPersonInfoProps> = {
    required: false,
  };

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
    if (this.props.validation && this.props.validation[this.props.field]) {
        const val:any = this.props.validation[this.props.field];
        if (val.hideWhen == null || val.hideWhen(this.props.formData, this.props.field)) {
            return val.hidden;
        }
    }
    return this.props.hidden;
}

  public render(): React.ReactElement<{}> {

    
    let value = null;
    let footer_text:string = "";
    const fieldRequired = this.fieldRequired();

    if (this.props.formData[this.props.field]) {
        value = this.props.formData[this.props.field];
    }
    if (fieldRequired && (!value || value.length === 0)) {
        footer_text = "This field is required";
    }
    const fieldHidden:boolean = this.checkIfHidden();
    let fieldToDisplay:any = null;

    if (!fieldHidden) {
      fieldToDisplay = 
      <div>        
        <div className={styles.container}>
          <div className={styles.row}>
            <div className={styles.column6}>
              <PeoplePickerField 
                label={this.props.label}
                field={this.props.field}
                context={this.props.context}
                formData={this.props.formData}
                validation={this.props.validation}
                onUpdate={this.props.onUpdate}
                itemLimit={1}
                onError={this.props.onError}
                disabled={this.props.disabled}
              />              
            </div>
            <div className={styles.column6}>
              <TextField 
                label="Job Title"
                value={(this.props.formData[this.props.field] && this.props.formData[this.props.field][0]) ? this.props.formData[this.props.field][0].secondaryText : ""}                
                disabled={true}
              />
            </div>
          </div>
          <div className={styles.row}>
            <div className={styles.column6}>
              <TextField 
                label="Department"
                value={(this.props.formData[this.props.field] && this.props.formData[this.props.field][0]) ? this.props.formData[this.props.field][0].tertiaryText : ""}                
                disabled={true}
              />
            </div>
            <div className={styles.column6}>
              <DropdownField 
                label="Purchaser's Directorate"
                formData={this.props.formData}
                field={this.props.field + 'Directorate'}
                onUpdate={this.props.onUpdate}
                validation={this.props.validation}  
                onError={this.props.onError}
                dropDownOptions={this.props.dropDownOptions}
                disabled={this.props.disabled}
              />
            </div>
          </div>
        </div>
      </div>;
    }

    return fieldToDisplay;
  }


}