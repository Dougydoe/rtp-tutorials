import * as React from 'react';
import { FileUpload } from '../../../../controls';
import { IAttachmentsProps, IFormData } from '../../../../interfaces/index';

export class Attachments extends React.Component<IAttachmentsProps, {}> {

  private validation:any = {
    "Attachments": {
      required: true,
      validateWhen: (formData:IFormData):boolean => {
        if (formData['ProcurementRoute'] == 1) {
          return true; 
        } else if ((formData['ProcurementRoute'] == 3) && (formData['TotalContractValue'] > 50000)) {
          return true;
        }
        return false;
      },
    },
  };

  public render(): JSX.Element {

    return (          
      <div>
        <FileUpload 
          context={this.props.context}
          listName={this.props.listName}
          formContext={this.props.formContext}
          onUpdate={this.props.onUpdate}
          formData={this.props.formData}
          label="Attach specification and any other relevant documentation if available"
          field="Attachments"
          validation={this.validation}
          onError={this.props.onError}
        /> 
      </div>     
    );
  }
}

