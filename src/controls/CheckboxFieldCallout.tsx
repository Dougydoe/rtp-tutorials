import * as React from 'react';
import { ICheckboxFieldProps } from '../interfaces';
import { Checkbox, Icon, Callout, DirectionalHint } from 'office-ui-fabric-react';
import "./Callout.css";
import styles from './FieldStyles.module.scss';

export interface ICheckboxFieldCalloutState {
	isCalloutVisible:boolean;
}

export class CheckboxFieldCallout extends React.Component<ICheckboxFieldProps, ICheckboxFieldCalloutState> {
    
    public static defaultProps:Partial<ICheckboxFieldProps> = {
			disabled: false,
			hidden: false,
			width: "small",
    };

    public state = {
			isCalloutVisible: false,
    };

    private checkboxChanged = (ev: React.FormEvent<HTMLElement>, checked: boolean): void => {
			this.props.onUpdate(this.props.field, checked);
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
     * @description if a field needs to be set as disabled on a per field level this can be used
     * @deprecated fields only need to be disabled on a per section level
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

    private _target = React.createRef<HTMLDivElement>();

    private _onClickCalloutIcon = ():void => {
			this.setState({
				isCalloutVisible: !this.state.isCalloutVisible
			});
    } 

    private _dismissCallout = ():void => {
			this.setState({
				isCalloutVisible: false
			});
    }

    public render(): React.ReactElement<ICheckboxFieldProps> {
			const fieldHidden:boolean = this.checkIfHidden();
			const calloutBody:string = this.props.calloutBody;

			let fieldToDisplay:any = null;
			if (!fieldHidden) {
				fieldToDisplay = 
					<div>
						<div className={styles.row}>
							<div className={this.props.width == "small" ? styles.column6 : styles.column8}>
								<Checkbox 
									label={this.props.label}
									checked={this.props.formData[this.props.field]}                
									onChange={this.checkboxChanged} 
									disabled={this.props.disabled}
								/>
								<Callout
									hidden={!this.state.isCalloutVisible}
									onDismiss={this._dismissCallout}
									directionalHint={DirectionalHint.rightCenter}
									target={this._target.current}
									calloutMaxWidth={500}
								>
									<p className="calloutExampleSubText">{calloutBody}</p>
								</Callout>
							</div>
							<div className={styles.column1} ref={this._target}>
								<Icon 
									iconName="Info" 
									className="iconCheckbox"   
									onClick={this._onClickCalloutIcon} 
								/> 
							</div>
						</div>
					</div>;
        } 
        return fieldToDisplay;
        
    }

 
}