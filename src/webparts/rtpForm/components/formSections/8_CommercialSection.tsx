import * as React from 'react';
import styles from '../RtpForm.module.scss';
import {IFormSectionProps, IFormData} from '../../../../interfaces/index';
import {
  TextBoxField, 
  ToggleField,
  DropdownField,
  PeoplePickerField,
  NumberField
} from '../../../../controls/index';
import { Label } from 'office-ui-fabric-react/lib/index';
// import { Attachments } from './Attachments';

export class CommercialSection extends React.Component<IFormSectionProps, {}> {
  
  private readonly validation:any = {
    "StaCompliant": {
      hidden: true,
      hideWhen: (formData:IFormData):boolean => {
        if (formData['ProcurementRoute'] == 3) {
          return false;
        }
        return true;
      }
    },    
    "StaCompliantNotes": {
      hidden: true,
      hideWhen: (formData:IFormData):boolean => {
        if (formData['ProcurementRoute'] == 3) {
          return false;
        }
        return true;
      }
    },    
  };

  public render(): JSX.Element {

    const formData = this.props.formData;
    const onUpdate = this.props.onUpdate; 
    const validation = this.validation;
    const onError = this.props.onError;
    const disabled = this.props.disabled;
    const dropDownOptions = this.props.dropDownOptions;

    return (
      <div>
        <div className={styles.row}>  
          <div className={styles.column}>              
            <Label className={styles.title}>Commercial</Label>
          </div>          
        </div>                                    
        <div>
        <div className={styles.row}>
          <div className={styles.column4}>                            
            <ToggleField  
              label="STA Compliant?"
              formData={formData}
              onUpdate={onUpdate}              
              field="StaCompliant"
              validation={validation}
              onError={onError}
              disabled={disabled}
            />    
          </div> 
          </div> 
          <div className={styles.row}>
            <div className={styles.column10}>                            
              <TextBoxField  
                label="STA Compliant Notes"
                formData={formData}
                onUpdate={onUpdate}              
                field="StaCompliantNotes"
                validation={validation}
                multiline={true}
                onError={onError}
                disabled={disabled}
              />    
          </div>
        </div>                                 
        </div>                                            
        <div className={styles.row}>  
          <div className={styles.column6}>                            
          <PeoplePickerField
              context={this.props.context}
              itemLimit={1}
              formData={formData}
              onUpdate={onUpdate}
              validation={validation}
              label="Commercial Manager"                
              field="CommercialManager"
              onError={onError}
              disabled={disabled}
              // placeholder="my placeholder"                  
            />  
          </div> 
          <div className={styles.column6}>                            
          <NumberField  
            label="Contract Id"
            formData={formData}
            onUpdate={onUpdate}              
            field="ContractNumber"
            validation={validation}
            onError={onError}
            disabled={disabled}
          />                             
          </div> 
        </div>
        <div className={styles.row}>  
          <div className={styles.column6}>                            
          <TextBoxField  
            label="Procurement Strategy"
            formData={formData}
            onUpdate={onUpdate}              
            field="ProcurementStrategy"
            validation={validation}
            onError={onError}     
            disabled={disabled}           
          />  
          </div> 
          <div className={styles.column6}>                            
          <DropdownField
            label="Procurement Type"
            formData={formData}
            field="ProcurementType"
            onUpdate={onUpdate}
            validation={validation}                
            placeHolder=""
            dropDownOptions={dropDownOptions}
            onError={onError}
            disabled={disabled}
          /> 
          </div>         
        </div>
        <div className={styles.row}>  
          <div className={styles.column}>                            
          <TextBoxField  
            label="Procurement Notes"
            formData={formData}
            onUpdate={onUpdate}              
            field="ProcurementNotes"
            multiline={true} 
            validation={validation}
            onError={onError}
            disabled={disabled}
          /> 
          </div>          
        </div>
        <div className={styles.row}>  
          <div className={styles.column}>                            
          <TextBoxField  
            label="Reason for Rejection"
            formData={formData}
            onUpdate={onUpdate}              
            field="ReasonForRejection"
            multiline={true} 
            validation={validation}
            onError={onError}
            disabled={disabled}
          /> 
          </div>          
        </div>
        {/* <Attachments 
          context={this.props.context}
          listName={this.props.listName}
          formContext={this.props.formContext}
          onUpdate={onUpdate}
          formData={formData}
        /> */}
      </div>
    );
 }
}





 





