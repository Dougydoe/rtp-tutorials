import * as React from 'react';
import styles from '../RtpForm.module.scss';
import { Label, Icon } from 'office-ui-fabric-react';
import { ComboField, CheckboxFieldCallout, ChoiceFieldCalloutProcurementRoute } from '../../../../controls';
import { IFormSectionProps, IFormData } from '../../../../interfaces';
import { PersonInfo } from './index';
import './HelpIcon.css';

export class PurchaserSection extends React.Component<IFormSectionProps, {}> {

  private readonly validation:any = {
    "PurchaserName": {required: true},    
    "PurchaserNameDirectorate": {required: true},    
    "ContractManager": {required: true},
    "ContractManagerDirectorate": {required: true},
    "Variation": {
      hidden: true,
      hideWhen: (formData:IFormData):boolean => {
        if (!formData['VariationCheck']) {
          return true;
        }
        return false;
      }
    },
    "ProcurementRoute": {required: true},
  };

  public render(): JSX.Element {

    const validation = this.validation;
    const { formData, onUpdate, onError, disabled, dropDownOptions } = this.props;

    return (
      <div>
        <div className={styles.row}>
          <div className={styles.column5}>
            <Label className={styles.subTitle}>For information click on the icon:</Label>
          </div>
          <div className={styles.column2}>
            <Icon iconName="Info" className="icon" />
          </div>
        </div> 
        <div className={styles.row}>
          <div className={styles.column}>
            <Label className={styles.title}>Status: {this.props.formData['ProcurementStatus']}</Label>
          </div>
        </div> 
        { this.props.formData['ProcurementStatus'] != "Draft" && 
          <div className={styles.row}>
            <div className={styles.column}>
              <Label className={styles.title}>Proc Reference: {this.props.formData['ProcReference']}</Label>
            </div>
          </div> 
        }                      
        <div className={styles.row}>
          <div className={styles.column}>
            <Label className={styles.title}>Contract Title: {this.props.formData['ContractTitle']}</Label>
          </div>
        </div>                       
        <div className={styles.row}>
          <div className={styles.column}>
            <Label className={styles.title}>Purchaser</Label>
            </div>
          </div>
          <CheckboxFieldCallout
            label="This is a variation of a previous procurement"
            field="VariationCheck"
            formData={formData}
            onUpdate={onUpdate}
            validation={validation}
            onError={onError}
            disabled={disabled}
            calloutBody="This is a variation to an existing contract"
          />  
          <div className={styles.row}>
            <div className={styles.column6}>
              <ComboField 
                field='Variation'
                formData={formData}
                label='Variation'
                onUpdate={onUpdate}
                validation={validation}
                onError={onError}
                dropDownOptions={dropDownOptions}
                disabled={disabled}
              />
            </div>
          </div>                     
          <PersonInfo
            context={this.props.context}
            itemLimit={1}
            label="Purchaser Name"
            field="PurchaserName"
            formData={formData}
            onUpdate={onUpdate}
            validation={validation}
            onError={onError}
            disabled={disabled}
            dropDownOptions={dropDownOptions}
          />
          <PersonInfo
            context={this.props.context}
            itemLimit={1}
            label="Contract Manager"
            field="ContractManager"
            formData={formData}
            onUpdate={onUpdate}
            validation={validation}
            onError={onError}
            disabled={disabled}
            dropDownOptions={dropDownOptions}
          />                                                                                                                 
          <CheckboxFieldCallout 
            label="Have you spoken with a member of the Commercial team about this procurement request?"
            field="SpokenToCommercial"
            formData={formData}
            onUpdate={onUpdate}
            validation={validation}
            onError={onError}
            disabled={disabled}
            calloutBody="It is advisable to discuss purchases in advance with the Commercial Team prior to filling out this form"
          />
          <ChoiceFieldCalloutProcurementRoute 
            label="What route will this procurement take?"
            field="ProcurementRoute"
            formData={formData}
            onUpdate={onUpdate}
            validation={validation} 
            onError={onError}   
            dropDownOptions={dropDownOptions}   
            disabled={disabled}
            calloutBody="Hello"
          />
      </div>
    );
  }
}


