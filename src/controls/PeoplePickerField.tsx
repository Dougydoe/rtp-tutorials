import * as React from 'react';
import { IPeoplePickerFieldProps } from '../interfaces';
import { Utility } from '../utils/Utility';
import { IPersonaProps } from 'office-ui-fabric-react/lib/Persona';
import { NormalPeoplePicker, ValidationState } from 'office-ui-fabric-react/lib/Pickers';
import { Label } from 'office-ui-fabric-react/lib/Label';
import styles from './FieldStyles.module.scss';
import { TooltipHost } from 'office-ui-fabric-react/lib';


export class PeoplePickerField extends React.Component<IPeoplePickerFieldProps, {}> {
  
  constructor(props: IPeoplePickerFieldProps) {
    super(props);

    this._onFilterChanged = this._onFilterChanged.bind(this);
  }
  
  public static defaultProps:Partial<IPeoplePickerFieldProps> = {
    required: false,
  };

  /**
     * @description does initial validation 
     * @fires when component mounts for this first time, after parent component's state.loading = false 
     */
    public componentDidMount() {
      // console.log('componentDidMount field: ' + this.props.field);
      let error:string = this.getFooterText();
      if (error) {
          this.props.onError(this.props.field);
      } else if (!error) {
          this.props.onError(this.props.field, true);
      }
  }

  /**
   * @description this component should only update when this.state.formData changes
   * ! double check that this.state.formData are the only things that should cause a re-render of this component 
   * @param nextProps to be received by the component
   * @returns true if component should update
   */
  public shouldComponentUpdate(nextProps):boolean {
      // console.log('shouldComponentUpdate field: ' + this.props.field);
      // if formData has changed, component should update
      if (this.props.formData != nextProps.formData) {
          return true;
      } else if (this.props.disabled != nextProps.disabled) {
        return true;
    }
      return false;
  }

  /**
   * @description validates the new input using the value stored in formData.field
   * @fires whenever this.shouldComponentUpdate() returns true
   * * This normally fires whenever this.setState() is called on the parent component
   */
  public componentDidUpdate() {
      // console.log('componentDidUpdate field: ' + this.props.field);        
      // do validation 
      let error:string = this.getFooterText();
      if (error) {
          this.props.onError(this.props.field);
      } else if (!error) {
          this.props.onError(this.props.field, true);
      }
  }

  private fieldRequired = ():boolean => {
    if (this.props.validation && this.props.validation[this.props.field]) {
        const val:any = this.props.validation[this.props.field];
        if (val.validateWhen == null || val.validateWhen(this.props.formData, this.props.field)) {
            return val.required;
        }
    }
    return this.props.required;
  }

    /**
     * @description checks to see if the field is required and empty
     * @returns an error message if invalid, else returns empty string if valid
     */
    private getFooterText = ():string => {
      let value = null;
      let footer_text:string = "";
      let fieldRequired = this.fieldRequired();

      //get field value
      if (this.props.formData[this.props.field]) {
          value = this.props.formData[this.props.field];
      }
      // check if field is required
      if (fieldRequired && (!value || value.length === 0)) {
          footer_text = "This field is required";                        
      }
      return footer_text;
  }

  private checkIfHidden = ():boolean => {
    // check if section should be hidden
    if (this.props.validation && this.props.validation['hideSection']) {
      const hideSection = this.props.validation['hideSection'];
      if (hideSection(this.props.formData)) {
        return true;
      } else if (this.props.validation[this.props.field]) {
        // check if field should be hidden
        const val:any = this.props.validation[this.props.field];
        if (val.hideWhen == null || val.hideWhen(this.props.formData, this.props.field)) {
            return val.hidden;
        }
      }
    } else if (this.props.validation && this.props.validation[this.props.field]) {
        const val:any = this.props.validation[this.props.field];
        if (val.hideWhen == null || val.hideWhen(this.props.formData, this.props.field)) {
            return val.hidden;
        }
    }
    return this.props.hidden;
}

  private _onItemsChange = (items: any[]): void => {    
    if (this.props.onUpdate) this.props.onUpdate(this.props.field, items);    
  }

  // ! When this renders as disabled it is still possible to remove the currently selected user 
  public render(): React.ReactElement<IPeoplePickerFieldProps> {
    let tooltip:string = "";
    if (this.props.tooltip && this.props.tooltip[this.props.field]) tooltip = this.props.tooltip[this.props.field];
    const fieldRequired = this.fieldRequired();
    const footer_text = this.getFooterText();
    const fieldHidden = this.checkIfHidden();
    let fieldToDisplay:any = null;
    if (!fieldHidden) {
      fieldToDisplay = 
      <div>
        <TooltipHost content={tooltip}>
        <Label required={fieldRequired}>{this.props.label}</Label>
          <NormalPeoplePicker         
            onResolveSuggestions={this._onFilterChanged}   
            onValidateInput={this._validateInput}         
            onChange={this._onItemsChange}      
            selectedItems={this.props.formData[this.props.field]}  
            itemLimit={this.props.itemLimit}   
            disabled={this.props.disabled}      
          />
          <span className={styles.dsErrorLabel}>{footer_text}</span>
          </TooltipHost>
      </div>;
    }
    return fieldToDisplay;
  }

  private async _onFilterChanged(filterText: string, currentPersonas: IPersonaProps[], limitResults?: number) {
    
    if (filterText.length > 2) {

        let filteredPersonas: IPersonaProps[] = new Array<IPersonaProps>();
        let siteUrl:string = this.props.context.pageContext.site.absoluteUrl;    

        return Utility.getPersonaForPeoplePickerField(filterText, siteUrl)
        .then(response => {
          filteredPersonas = response;
          filteredPersonas = this._removeDuplicates(filteredPersonas, currentPersonas);
          filteredPersonas = limitResults ? filteredPersonas.splice(0, limitResults) : filteredPersonas;
          return filteredPersonas;
        });

    } else {
      return [];
    }
  }

  private _removeDuplicates(personas: IPersonaProps[], possibleDupes: IPersonaProps[]) {
    return personas.filter(persona => !this._listContainsPersona(persona, possibleDupes));
  }

  private _listContainsPersona(persona: IPersonaProps, personas: IPersonaProps[]) {
    if (!personas || !personas.length || personas.length === 0) {
      return false;
    }
    return personas.filter(item => item.text === persona.text).length > 0;
  }

  /* private _doesTextStartWith(text: string, filterText: string): boolean {
    return text.toLowerCase().indexOf(filterText.toLowerCase()) === 0;
  } */

  private _validateInput = (input: string): ValidationState => {
    if (input.indexOf('@') !== -1) {
      return ValidationState.valid;
    } else if (input.length > 1) {
      return ValidationState.warning;
    } else {
      return ValidationState.invalid;
    }
  }

}


