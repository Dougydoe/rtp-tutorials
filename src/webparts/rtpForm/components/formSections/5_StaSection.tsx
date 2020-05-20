import * as React from 'react';
import styles from '../RtpForm.module.scss';
import {Label} from 'office-ui-fabric-react/lib/index';
import {
  TextBoxField, 
  DropdownField, 
  ChoiceField,
  NumberField
} from '../../../../controls/index';
import {IFormSectionProps, IFormData} from '../../../../interfaces/index';

export class StaSection extends React.Component<IFormSectionProps, {}> {
  
  private readonly validation:any = {
    "CompanyName": {
      required: true,      
      validateWhen: (formData:IFormData):boolean => {
        if (formData['ProcurementRoute'] == 3) {
          return true;
        }
        return false;
      }
    },  
    "ContactName": {
      required: true,
      validateWhen: (formData:IFormData):boolean => {
        if (formData['ProcurementRoute'] == 3) {
          return true;
        }
        return false;
      }
    },  
    "TendererContactEmail": {
      required: true,
      validateWhen: (formData:IFormData):boolean => {
        if (formData['ProcurementRoute'] == 3) {
          return true;
        }
        return false;
      }
    },  
    "TendererContactTelephone": {
      required: true,
      validateWhen: (formData:IFormData):boolean => {
        if (formData['ProcurementRoute'] == 3) {
          return true;
        }
        return false;
      }
    },
    "StaType": {
      required: true,
      validateWhen: (formData:IFormData):boolean => {
        if (formData['ProcurementRoute'] == 3) {
          return true;
        }
        return false;
      }
    },
    "UnableProcureCompetitively": {
      required: true,
      validateWhen: (formData:IFormData):boolean => {
        if (formData['ProcurementRoute'] == 3) {
          return true;
        }
        return false;
      }
    },
    "UnableProcureCompetitivelyReason": {
      hidden: true,
      hideWhen: (formData:IFormData):boolean => {
        if (formData['UnableProcureCompetitively'] == "Other (free text box is displayed on this option)") {
          return false;
        }
        return true;
      }
    },
    "RedraftingNotPossible": {
      hidden: true,
      hideWhen: (formData:IFormData):boolean => {
        if (formData['UnableProcureCompetitively'] == "Tight time constraints caused by circumstances outside CMA’s control preclude a full tender and CMA can demonstrate that the proposed supplier represents good value for money, AND, the specification/requirement cannot be redrafted/changed to allow for competition (please state why redrafting is not possible).") {
          return false;
        }
        return true;
      }
    },
    hideSection: (formData:IFormData):boolean => {
      if (formData['ProcurementRoute'] == 3) {
        return false;
      }
      return true;
    },
  };

  public render(): JSX.Element {

    const formData:IFormData = this.props.formData;
    const onUpdate = this.props.onUpdate; 
    const validation = this.validation;
    const onError = this.props.onError;
    const dropDownOptions = this.props.dropDownOptions;
    const disabled = this.props.disabled;

    return (
      <div>
        <div className={styles.row}>
          <div className={styles.column8}>
            <DropdownField
              label="STA Type"
              formData={formData}
              field="StaType"
              onUpdate={onUpdate}
              validation={validation}                
              placeHolder="Please select the type"
              dropDownOptions={dropDownOptions}
              onError={onError}
              disabled={disabled}
            />                
          </div>
        </div> 
        <div className={styles.row}>
          <div className={styles.column10}>
            <ChoiceField
              label="Why are you unable to procure this competitively?"
              field="UnableProcureCompetitively"
              formData={formData}
              onUpdate={onUpdate}
              validation={validation}  
              onError={onError} 
              dropDownOptions={dropDownOptions}  
              disabled={disabled}             
            />                
          </div>    
        </div>  
        <div className={styles.row}>
          <div className={styles.column10}>
            <TextBoxField 
              label="Please state why redrafting is not possible?"
              field="RedraftingNotPossible"
              formData={formData}
              onUpdate={onUpdate}
              validation={validation}
              onError={onError}
              multiline={true}
            />               
          </div>    
        </div>  
        <div className={styles.row}>
          <div className={styles.column10}>
            <TextBoxField
              label="Please detail the other reason you were unable to procure this competitively"
              formData={formData}
              onUpdate={onUpdate}              
              field="UnableProcureCompetitivelyReason"
              validation={validation}
              multiline={true}
              onError={onError}
              disabled={disabled}
            />
          </div> 
        </div> 
        { (formData['ProcurementRoute'] == 3) ?                 
        <div className={styles.row}>
          <div className={styles.column}>
            <Label className={styles.title}>Suggested Tenderer</Label>              
          </div>                  
        </div> : null }
        <div className={styles.row}>
          <div className={styles.column}>
          <ChoiceField
            label="Is the suggested tenderer: ?"
            field="SuggestedTendererStatus"
            formData={formData}
            onUpdate={onUpdate}
            validation={validation}  
            onError={onError} 
            dropDownOptions={dropDownOptions}    
            disabled={disabled}           
            />
          </div>                  
        </div> 
        <div className={styles.row}>
          <div className={styles.column6}>
            <TextBoxField  
              label="Company Name"
              formData={formData}
              onUpdate={onUpdate}              
              field="CompanyName"
              validation={validation}  
              onError={onError}      
              disabled={disabled}        
            />                            
          </div> 
          <div className={styles.column6}>
            <TextBoxField  
              label="Contact Name"
              formData={formData}
              onUpdate={onUpdate}              
              field="ContactName"
              validation={validation} 
              onError={onError}     
              disabled={disabled}          
            />                                       
          </div>                 
        </div>  
        <div className={styles.row}>
          <div className={styles.column6}>
            <TextBoxField  
              label="Email"
              formData={formData}
              onUpdate={onUpdate}              
              field="TendererContactEmail"
              validation={validation}   
              onError={onError}      
              disabled={disabled}       
            />                
          </div> 
          <div className={styles.column6}>
            <NumberField  
              label="Telephone"
              formData={formData}
              onUpdate={onUpdate}              
              field="TendererContactTelephone"
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



