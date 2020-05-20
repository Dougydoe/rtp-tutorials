import * as React from 'react';
import { ITagPickerFieldProps} from '../interfaces';
import { Label, TagPicker, ITag, TooltipHost } from 'office-ui-fabric-react';
import styles from './FieldStyles.module.scss';
import { Utility } from '../utils/Utility';
import { Environment, EnvironmentType } from '@microsoft/sp-core-library';
import { MockTagPickerFieldOptions } from '../mock/MockDropdowns';

export class TagPickerField extends React.Component<ITagPickerFieldProps, {}> {

    public static defaultProps:Partial<ITagPickerFieldProps> = {
        required: false,
        disabled: false,
        hidden: false,
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
        // if formData have changed, component should update
        // console.log('formData: ' + JSON.stringify(this.props.formData));
        // console.log('nextProps.formData: ' + JSON.stringify(nextProps.formData));
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
        // ! component does not update if it is not rendered   
        // do validation 
        let error:string = this.getFooterText();
        if (error) {
            // the field is empty
            this.props.onError(this.props.field);
        } else if (!error) {
            // the field is NOT empty
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
        if (fieldRequired && (!value || value.length === 0)) {
            footer_text = "This field is required";                        
        }
        return footer_text;
    }

    private onChange = (items: { key: string; name: string }[]):void => {
        this.props.onUpdate(this.props.field, items);
    }

    public render(): React.ReactElement<ITagPickerFieldProps> {
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
                <TagPicker 
                    selectedItems = {this.props.formData[this.props.field]}
                    onEmptyInputFocus={this._onTpEmptyInputFocus}
                    onResolveSuggestions={this._onTpFilterChanged}
                    onChange={this.onChange}
                    itemLimit = {this.props.itemLimit}
                    pickerSuggestionsProps={{
                        suggestionsHeaderText: this.props.suggestionsHeaderText,
                        noResultsFoundText: this.props.noResultsFoundText
                    }}
                    disabled={this.props.disabled}                
                />
                <span className={styles.dsErrorLabel}>{footer_text}</span>
                </TooltipHost>
            </div>;
        }
        return fieldToDisplay;
    }

    private _onTpFilterChanged = (filterText: string, tagList:ITag[]):ITag[] | PromiseLike<ITag[]> => {
        return this.getTagsFilterChanged(filterText, tagList);
    }

    private async getTagsFilterChanged(filterText: string, tagList:ITag[]):Promise<ITag[]> {
        if (Environment.type === EnvironmentType.Local) {
            return MockTagPickerFieldOptions;
        } else if (Environment.type === EnvironmentType.SharePoint) {
            let lookupTags:ITag[] = await Utility.getTagOptions(this.props.field);          
            if (filterText) {
                // exclude tags that don't match the filterText
                // exlude tags that have already been selected (in the tagList array)
                return lookupTags
                    .filter(tag => tag.name.toLowerCase().indexOf(filterText.toLowerCase()) === 0)
                    .filter(tag => !this._listContainsDocument(tag, tagList));
            } else {
                return [];
            }
        }
    }
        
    private _listContainsDocument(tag: ITag, tagList: ITag[]) {
        if (!tagList || !tagList.length || tagList.length === 0) {
            return false;
        }
        return tagList.filter(compareTag => compareTag.key === tag.key).length > 0;
    }  
    
    private _onTpEmptyInputFocus = (tagList: ITag[]): ITag[] | PromiseLike<ITag[]> => {
        return this.getTagsEmptyInputFocus(tagList);
    }

    private async getTagsEmptyInputFocus(tagList:ITag[]):Promise<ITag[]> {
        let lookupTags:ITag[] = await Utility.getTagOptions(this.props.field);                          
        // exlude tags that have already been selected (in the tagList array)
        return lookupTags.filter(tag => !this._listContainsDocument(tag, tagList));        
    }
}
