import * as React from 'react';
import styles from '../RtpForm.module.scss';
import {IFormSectionProps} from '../../../../interfaces/index';
import {
  ChoiceField 
} from '../../../../controls/index';
import {Label} from 'office-ui-fabric-react/lib/index';

export class IASection extends React.Component<IFormSectionProps, {}> {
  
  private readonly validation:any = {
    "DpoOptions": {required: true}, 
    "IsmOptions": {required: true},   
  };

  public render(): JSX.Element {

    const formData = this.props.formData;
    const onUpdate = this.props.onUpdate; 
    const validation = this.validation;
    const onError = this.props.onError;
    const disabled = this.props.disabled;

    return (
      <div>
        {/* <div className={styles.rtpForm}> */}
          {/* <div className={styles.container}> */}
            <div className={styles.row}>
              <div className={styles.column}>
                <Label className={styles.title}>Information Assurance</Label>
              </div>
            </div>
            <div className={styles.row}>
              <div className={styles.column10}>
              <Label>As purchasers on behalf of the CMA, it is your responsibility to ensure that data handling and Information Security is addressed within your requirements. 
                The General Data Protection Regulations (EU) 2016/679
                Determines how information must be treated by the supplier.  These aspects should be reflected within your requirement and any subsequent contract.  The Information Security Manager can advise further on these aspects.
              </Label>                
              </div>
            </div>
            <div className={styles.row}>
              <div className={styles.column10}>
              <ChoiceField
                label="I confirm that I have discussed Data Protection with the Data Protection Officer (DPO) and I confirm that:"
                field="DpoOptions"
                formData={formData}
                onUpdate={onUpdate}
                validation={validation} 
                onError={onError}   
                dropDownOptions={this.props.dropDownOptions}   
                disabled={disabled}         
              />                
              </div>
            </div>
            <div className={styles.row}>
              <div className={styles.column10}>
              <ChoiceField
                label="I confirm that I have liaised with the Information Security Manager and I confirm that:"
                field="IsmOptions"
                formData={formData}
                onUpdate={onUpdate}
                validation={validation}    
                onError={onError}  
                dropDownOptions={this.props.dropDownOptions}      
                disabled={disabled}      
              />                                                      
              </div>
            </div>            
          </div>
        // </div>
      // </div>
    );
 }
}


