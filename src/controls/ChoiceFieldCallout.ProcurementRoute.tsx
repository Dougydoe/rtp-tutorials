import * as React from 'react';
import styles from './FieldStyles.module.scss';
import { IChoiceFieldProps } from '../interfaces';
import { ChoiceGroup, IChoiceGroupOption, Label, Icon, Callout, DirectionalHint } from 'office-ui-fabric-react';
import "./Callout.css";


export interface IChoiceFieldCalloutProcurementRouteState {
  isCalloutVisible:boolean;
}

export class ChoiceFieldCalloutProcurementRoute extends React.Component<IChoiceFieldProps, IChoiceFieldCalloutProcurementRouteState> {
    
    constructor(props:IChoiceFieldProps) {
      super(props);

      this.state = {
        isCalloutVisible: false,
      };
    }

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
     * ! double check that this.state.formData and this.state.dropDownOptions are the only things that should cause a re-render of this component 
     * @param nextProps to be received by the component
     * @returns true if component should update
     */
    public shouldComponentUpdate(nextProps, nextState):boolean {
        // console.log('shouldComponentUpdate field: ' + this.props.field);
        // if formData OR dropDownOptions have changed, component should update
        if (this.props.formData != nextProps.formData || this.props.dropDownOptions != nextProps.dropDownOptions) {
            return true;
        } else if (this.props.disabled != nextProps.disabled) {
            return true;
        } else if (this.state.isCalloutVisible != nextState.isCalloutVisible) {
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
                const currentValue = this.props.formData[this.props.field];
                if(currentValue && currentValue !== undefined) {
                    return false;
                }
                return val.required;
            }
        }
        return this.props.required;
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
        if (fieldRequired && !value) {
            footer_text = "This field is required";                        
        }
        return footer_text;
    }

    /**
     * @description if a field needs to be set as disabled on a per field level this can be used
     * @deprecated - fields only need to be disabled on a per section level
     */
    /* private fieldDisabled = ():boolean => {
        if (this.props.validation && this.props.validation[this.props.field]) {
            const val:any = this.props.validation[this.props.field];
            if (val['disabledWhen'] && val.disabledWhen(this.props.formData)) {                                             
                return true;
            }
        }
        return this.props.disabled;
    } */

    private handleChange = (e:React.FormEvent<HTMLInputElement>, option:IChoiceGroupOption):void => {
        const value = option.key;
        this.props.onUpdate(this.props.field, value);
    }

    private _target = React.createRef<HTMLDivElement>();

    private _onClickIcon = ():void => {
      this.setState({
          isCalloutVisible: !this.state.isCalloutVisible
      });
    } 

    private _dismiss = ():void => {
      this.setState({
          isCalloutVisible: false
      });
    }

    public render(): React.ReactElement<IChoiceFieldProps> {
        const footer_text = this.getFooterText();
        const fieldHidden:boolean = this.checkIfHidden();
        const calloutBody:string = this.props.calloutBody;

        let fieldToDisplay:any = null;
        if (!fieldHidden) {
          fieldToDisplay = 
            <div>
              <div className={styles.row}>
                <div className={styles.column6}>
                  <Label required={this.fieldRequired()}>{this.props.label}</Label>
                  <ChoiceGroup 
                    selectedKey={this.props.formData[this.props.field]}
                    options={this.props.dropDownOptions[this.props.field]}
                    onChange={this.handleChange}
                    disabled={this.props.disabled}
                  />
                  <span className={styles.dsErrorLabel}>{footer_text}</span>
                  <Callout
                    hidden={!this.state.isCalloutVisible}
                    onDismiss={this._dismiss}
                    directionalHint={DirectionalHint.rightCenter}
                    target={this._target.current}
  									calloutMaxWidth={500}
                  >
                    <div className="calloutExampleSubText">
                      <p>This form must be completed to record all single tender action (STA) with a value of 10K and over which are for:</p>
                      <ul>
                        <li>a new contract for a new requirement</li>
                        <li>a direct award under a call- off from Framework Contracts (i.e. award of contract to one supplier without undertaking a competition)</li>
                        <li>a change to an existing contract (extensions/variations to scope) which was not considered or included in the original approval (Business Case) and not covered in the original contract.</li>
                      </ul>
									  </div>
                  </Callout>
                </div>
                <div className={styles.column1} ref={this._target}>
                  <Icon 
                    iconName="Info" 
                    className="iconTextbox"   
                    onClick={this._onClickIcon} 
                  /> 
                </div>
              </div>
            </div>;
        }
        return fieldToDisplay;
    }
}