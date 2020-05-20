import * as React from 'react';
import styles from '../RtpForm.module.scss';
import { IFormSectionProps, IFormData } from '../../../../interfaces/index';
import { Label } from 'office-ui-fabric-react/lib/index';
import { 
  TextBoxField,  
  DropdownField,
} from '../../../../controls/index';

export class FTDPSection extends React.Component<IFormSectionProps, {}> {
  
  private readonly validation:any = {
    "ApprovedScope": {
      required: true,
      validateWhen: (formData:IFormData):boolean => {
        if (formData['ProcurementRoute'] == 2) {
          return true;
        }
        return false;
      },
      hidden: true,
      hideWhen: (formData:IFormData):boolean => {
        if (formData['ProcurementRoute'] == 2) {
          return false;
        }
        return true;
      }
    },
    "RouteToMarket": {
      required: true,
      validateWhen: (formData:IFormData):boolean => {
        if (formData['ProcurementRoute'] == 2) {
          return true;
        }
        return false;
      },
      hidden: true,
      hideWhen: (formData:IFormData):boolean => {
        if (formData['ProcurementRoute'] == 2) {
          return false;
        }
        return true;
      }
    },
    "OtherFrameworkName": {
      hidden: true,
      hideWhen: (formData:IFormData):boolean => {
        if ((formData['ProcurementRoute'] == 2) && formData['RouteToMarket'] == 4) {
          return false;
        }
        return true;
      }
    },
    hideSection: (formData:IFormData):boolean => {
      if (formData['ProcurementRoute'] == 3) {
        return true;
      }
      return false;
    },
  };

  public render(): JSX.Element {

    const formData:IFormData = this.props.formData;
    const onUpdate = this.props.onUpdate; 
    const validation = this.validation;  
    const onError = this.props.onError; 
    const disabled = this.props.disabled; 

    return (
      <div>
        { formData['ProcurementRoute'] == 2 ?
        <div className={styles.row}> 
          <div className={styles.column}>   
            <Label className={styles.title}>Fast Track Delegated Procurement (FTDP)</Label>  
          </div>
        </div> : null }
        { formData['ProcurementRoute'] == 2 ?            
        <div className={styles.row}> 
          <div className={styles.column10}>   
            <Label>Please note that for any FTDP requirements, this RTP is to be completed by staff who are “Registered Purchasers” (RP).</Label>
          </div>
        </div>  : null }            
        <div className={styles.row}> 
            <div className={styles.column10}>   
              <DropdownField
                  label="Approved Scope (indicate which FTDP work strand this work falls under):"            
                  formData={formData}
                  field="ApprovedScope"
                  onUpdate={onUpdate}
                  validation={validation}                      
                  placeHolder=""
                  dropDownOptions={this.props.dropDownOptions}
                  onError={onError}
                  disabled={disabled}
                />
            </div>
          </div>                    
          <div className={styles.row}> 
            <div className={styles.column10}>   
              <DropdownField
                  label="Route to Market"              
                  formData={formData}
                  field="RouteToMarket"
                  onUpdate={onUpdate}
                  validation={validation}                        
                  placeHolder=""
                  dropDownOptions={this.props.dropDownOptions}
                  onError={onError}
                  disabled={disabled}
                />
            </div>
          </div> 
          <div className={styles.row}> 
            <div className={styles.column10}>
              <TextBoxField
                label="Please provide other framework name"
                formData={formData}
                onUpdate={onUpdate}              
                field="OtherFrameworkName"
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





 





