import * as React from 'react';
import styles from '../RtpForm.module.scss';

import { Label } from 'office-ui-fabric-react/lib/index';
import { 
  ChoiceField, 
  CheckboxField,
  TagPickerField,   
  CheckboxFieldCallout,
  CheckboxFieldCalloutBusinessCaseAttached
} from '../../../../controls/index';
import { IFormSectionProps, IFormData } from '../../../../interfaces/index';
import { Attachments } from './Attachments';

export class RequirementSection extends React.Component<IFormSectionProps, {}> { 

  private readonly validation:any = {
    "PurchaseCategory": {
      required: true,
      validateWhen: (formData:IFormData):boolean => {
        if (formData['ProcurementRoute'] == 3) {
          return false; 
        }
        return true;
      },
      hidden: true,
      hideWhen: (formData:IFormData):boolean => {
        if (formData['ProcurementRoute'] == 3) {
          return true;
        }
        return false;
      }
    },
    "BusinessCaseAttached": {
      hidden: true,
      hideWhen: (formData:IFormData):boolean => {
        if (formData['ProcurementRoute'] == 2) {
          return true;
        }
        return false;
      }
    },   
    "BusinessCaseRequirement": {
      hidden: true,
      hideWhen: (formData:IFormData):boolean => {
        if (!formData['BusinessCaseAttached']) {
          return true;
        } else if (formData['ProcurementRoute'] == 2) {
          return true;
        }
        return false;
      }
    }, 
    /* hideSection: (formData:IFormData):boolean => {
      if (formData['ProcurementRoute'] == 3) {
        return true;
      }
      return false;
    }, */  
  };

  private readonly tooltips:any = {
    "BusinessCaseAttached": "",
  };
  
  public render(): JSX.Element {

    const { formData, onUpdate, onError, disabled, dropDownOptions } = this.props;
    const validation = this.validation;

    return (          
      <div>
        <CheckboxFieldCalloutBusinessCaseAttached
          label="Business Case attached?"
          formData={formData}
          field="BusinessCaseAttached"
          onUpdate={onUpdate}
          validation={validation}
          onError={onError}
          disabled={disabled}
        />
        <div className={styles.row}>  
          <div className={styles.column}>                   
            <ChoiceField
              label="I have attached a business case because the requirement:"
              field="BusinessCaseRequirement"
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
        <div className={styles.column}>
          <Label className={styles.subTitle}>Attachments</Label>
        </div>
      </div>                                 
        <Attachments 
          context={this.props.context}
          listName={this.props.listName}
          formContext={this.props.formContext}
          onUpdate={onUpdate}
          formData={formData}   
          onError={onError}           
        />                         
        <div className={styles.row}>  
          <div className={styles.column6}>         
            <TagPickerField
              label="Purchase Category"
              formData={formData}
              field="PurchaseCategory"
              onUpdate={onUpdate}                  
              validation={validation}                  
              suggestionsHeaderText="Categories"
              noResultsFoundText="No matching categories"
              itemLimit={1}
              onError={onError}
              disabled={disabled}
            />
          </div>
        </div>  
      </div>
    );
  }
}

