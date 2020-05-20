import * as React from 'react';
import styles from '../RtpForm.module.scss';
import { IFormSectionProps } from '../../../../interfaces';
import {
  TextBoxField,
  NumberField,
  DropdownField, 
  ComboField
} from '../../../../controls';
import { Label } from 'office-ui-fabric-react/lib';
import { TextField } from 'office-ui-fabric-react';

export class BudgetCheckSection extends React.Component<IFormSectionProps, {}> {
  
  private readonly validation:any = {
    "SubProjectCode": {required: true}, 
    "SecondaryBusinessManager": {required: true},
    "SecondaryFinancialApprover": {required: true},
  };

  public render(): JSX.Element {

    const formData = this.props.formData;
    const onUpdate = this.props.onUpdate; 
    const validation = this.validation;
    const onError = this.props.onError;
    const dropDownOptions = this.props.dropDownOptions;
    const disabled = this.props.disabled;

    return (
      <div>
        <div className={styles.row}>
          <div className={styles.column}>
          <Label className={styles.title}>Budget Check</Label>   
          </div>
        </div>   
        <div className={styles.row}>
          <div className={styles.column}>              
          <Label>Please enter your sub project code, this will auto-populate the cost centre, primary business manager, directorate and primary financial approver. When this form is submitted, the business manager will receive an email requesting approval. Once approved, you will be notified. Please also choose the a secondary business manager (in the event that the primary business manager is unavailable to approve your request).</Label>           
          </div>
        </div>
        <div className={styles.row}>
          <div className={styles.column6}>   
            <ComboField
              label="Sub Project Code"
              field="SubProjectCode"
              formData={formData}
              onUpdate={onUpdate}              
              validation={validation}
              onError={onError}
              dropDownOptions={dropDownOptions}
              disabled={disabled}
            />                       
          </div>
          <div className={styles.column6}>  
            <TextBoxField
              label="Purchasing Directorate"
              formData={formData}
              field="DirectorateOfPurchase"
              onUpdate={onUpdate}
              validation={validation}                  
              disabled={true}
              onError={onError}
            />                                
          </div>
        </div>
        <div className={styles.row}>
          <div className={styles.column6}>   
            <TextBoxField  
              label="Sub Project Code Description"
              field="SubProjectCodeDescription"
              formData={formData}
              onUpdate={onUpdate}              
              validation={validation}
              disabled={true}
              onError={onError}
            />                         
          </div>
          <div className={styles.column3}>   
            <TextBoxField  
              label="Project Code"
              field="ProjectCode"
              formData={formData}
              onUpdate={onUpdate}              
              validation={validation}
              disabled={true}
              onError={onError}
            />                         
          </div>
          <div className={styles.column3}>              
            <NumberField  
              label="Cost Centre"
              field="CostCentre"
              formData={formData}
              onUpdate={onUpdate}              
              validation={validation}
              disabled={true}
              onError={onError}
            />       
          </div>
        </div>                      
        <div className={styles.row}>
          <div className={styles.column6}>   
          <TextBoxField
            label="Primary Business Manager"
            formData={formData}
            field="PrimaryBusinessManager"
            onUpdate={onUpdate}
            validation={validation}  
            disabled={true} 
            onError={onError}
          />              
          </div>
          <div className={styles.column6}>
            <DropdownField
              label="Secondary Business Manager"
              formData={formData}
              field="SecondaryBusinessManager"
              onUpdate={onUpdate}
              validation={validation}  
              dropDownOptions={dropDownOptions}   
              onError={onError}     
              disabled={disabled}                        
            />
          </div>
        </div>             
        <div className={styles.row}>
          <div className={styles.column}>              
          <Label className={styles.title}>Financial Approval</Label>  
          </div>
        </div>  
        <div className={styles.row}>
          <div className={styles.column}>              
          <Label>Choose a secondary financial approver from the list below. When this form is submitted, the secondary financial approver will also receive an email requesting approval. Once approved, you will be notified.</Label>                               
          </div>
        </div>   
        <div className={styles.row}>
          <div className={styles.column6}>              
          <TextBoxField
            label="Primary Financial Approver"
            formData={formData}
            field="PrimaryFinancialApprover"
            onUpdate={onUpdate}
            validation={validation}                
            disabled={true}
            onError={onError}
          /> 
          </div>
          <div className={styles.column6}>
          <DropdownField
            label="Secondary Financial Approver"
            formData={formData}
            field="SecondaryFinancialApprover"
            onUpdate={onUpdate}
            validation={validation}    
            dropDownOptions={dropDownOptions}    
            onError={onError}     
            disabled={disabled}                   
          /> 
          </div>
        </div>                
      </div>
    );
 }
}





 





