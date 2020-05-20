import * as React from 'react';
import styles from '../RtpForm.module.scss';
import {
  TextBoxField, 
  NumberField, 
  DateField, 
  TextBoxFieldCallout
} from '../../../../controls/index';
import {IFormSectionProps, IFormData} from '../../../../interfaces/index';
import { Label } from 'office-ui-fabric-react/lib';

export class ContractDetailsSection extends React.Component<IFormSectionProps, {}> {
  
  private readonly validation:any = {
    "ProcurementStartDate": {required: true}, 
    "ProcurementEndDate": {required: true}, 
    "ContractTitle": {required: true}, 
    "ContractValue": {required: true}, 
    "ReasonsForPurchase": {
      required: true,
      validateWhen: (formData:IFormData):boolean => {
        if (formData['ProcurementRoute'] == 3) {
          return true;
        }
        return false;
      },
      hidden: true,
      hideWhen: (formData:IFormData):boolean => {
        if (formData['ProcurementRoute'] == 3) {
          return false;
        } else if (formData['ProcurementRoute'] == 2) {
          return false;
        }
        return true;
      }
    },
  };

  /* private readonly tooltips:any = {
    "ContractTitle": "Please enter the contract title here",
  }; */

  public render(): JSX.Element {

    const formData:IFormData = this.props.formData;
    const onUpdate = this.props.onUpdate; 
    const validation = this.validation;
    const onError = this.props.onError;
    const disabled = this.props.disabled;

    return (
      <div>
        <div className={styles.row}>
            <div className={styles.column10}>
              <Label className={styles.title}>The Requirement</Label>
            </div>                            
          </div>    
          <div className={styles.row}>              
            <div className={styles.column6}>
              <DateField
                label="Contract Start Date"
                field="ProcurementStartDate"
                formData={formData}
                onUpdate={onUpdate}
                validation={validation}
                onError={onError}
                disabled={disabled}
              />                
            </div> 
            <div className={styles.column6}>
              <DateField
                label="Contract End Date"
                field="ProcurementEndDate"
                formData={formData}
                onUpdate={onUpdate}
                validation={validation}
                onError={onError}
                disabled={disabled}
              />                
            </div>                 
          </div>  
          { !(formData['ProcurementRoute'] == 2) ? 
          <div className={styles.row}>
            <div className={styles.column6}>
              <Label className={styles.description}>The Start Date must be at least 6 weeks from the date this form is submitted</Label>
            </div>  
          </div> : null } 
          <div className={styles.row}>
            <div className={styles.column8}>
              <TextBoxField 
                label="Contract Title"
                field="ContractTitle"
                formData={formData}
                onUpdate={onUpdate} 
                validation={validation}   
                onError={onError}  
                disabled={disabled}  
              />                
            </div>              
          </div> 
          <div className={styles.row}>
            <div className={styles.column3}>
              <NumberField 
                label="Previous Contract(s) Value"
                field="PreviousContractValue"
                formData={formData}
                onUpdate={onUpdate}
                validation={validation}
                onError={onError}
                disabled={true}
                // prefix="£"                                   
              />                
            </div>
            <div className={styles.column3}>
              <NumberField 
                label="Contract Value"
                field="ContractValue"
                formData={formData}
                onUpdate={onUpdate}
                validation={validation}
                onError={onError}
                disabled={disabled}
                // prefix="£"                                   
              />                
            </div>
            <div className={styles.column3}>
              <NumberField 
                label="Total Contract Value"
                field="TotalContractValue"
                formData={formData}
                onUpdate={onUpdate}
                validation={validation}
                onError={onError}
                disabled={true}
                // prefix="£"                                   
              />
            </div>
          </div>   
          <div className={styles.row}>
            <div className={styles.column10}>
              <TextBoxField 
                label="Reasons for the purchase (description and reasons for the requirement including consideration of alternatives and value for money - VFM)"
                field="ReasonsForPurchase"
                formData={formData}
                onUpdate={onUpdate}
                multiline={true} 
                validation={validation}  
                onError={onError}  
                disabled={disabled}                                           
              />                
            </div>                  
          </div>               
      </div>
    );
 }
}


