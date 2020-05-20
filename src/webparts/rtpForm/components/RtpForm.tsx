import * as React from 'react';
import {
  IRtpFormProps,
  IRtpFormState,
  IFormData,
} from '../../../interfaces';
import styles from './RtpForm.module.scss';
import {
  PurchaserSection,
  RequirementSection,
  IASection,
  FTDPSection,
  BudgetCheckSection,
  StaSection,
  CommercialSection,
  ContractDetailsSection
} from './formSections';
import {
  CommandBar,
  Spinner,
  IContextualMenuItem,
} from 'office-ui-fabric-react/lib/index';
import { Utility, AgressoDataLookup, Form, FileApi } from '../../../utils/';
import { IDropDownOptions } from '../../../interfaces/IRtpFormState';
import { Environment, EnvironmentType } from '@microsoft/sp-core-library';


export default class RtpForm extends React.Component<IRtpFormProps, IRtpFormState> {

  constructor(props: IRtpFormProps) {
    super(props);

    this.state = {
      formData: {
        ProcurementStatus: "Draft"
      },
      errors: [],
      userInfo: {
        userRole: "",
        isCommercialMember: false,
      },
      isSubmitting: false,
      loading: true,
    };
    this.handleSave = this.handleSave.bind(this);
    this.handleSubmit = this.handleSubmit.bind(this);
    this.handleDiscard = this.handleDiscard.bind(this);
    this.updateFormData = this.updateFormData.bind(this);

  }

  /**
   * @description details the workflow when the component mounts for the first time
   * @fires when component mounts for the first time
   */
  public async componentDidMount() {
    const dataToLoad = await Form.load(this.props.formContext, this.props.pageContext, this.props.listName);
    const allDropDownOptions: IDropDownOptions = { ...dataToLoad.dropDownOptions[0], ...dataToLoad.dropDownOptions[1], ...dataToLoad.dropDownOptions[2] };
    if (dataToLoad.formData == null) {
      this.setState({
        userInfo: dataToLoad.userInfo,
        dropDownOptions: allDropDownOptions,
        loading: false,
      });
    } else {
      this.setState({
        formData: dataToLoad.formData,
        userInfo: dataToLoad.userInfo,
        dropDownOptions: allDropDownOptions,
        loading: false,
      });
    }

    // auto-complete the Purchaser field
    if (this.props.formContext.mode == "new") {
      // get current user and add to formData 
      if (Environment.type === EnvironmentType.Local) {
        const currentUserDisplayName = "Annie Lindqvist";
        const currentUserPersona = await Utility.getPersonaForPeoplePickerField(currentUserDisplayName, this.props.pageContext.web.absoluteUrl);
        this.updateFormData('PurchaserName', currentUserPersona);

      } else if (Environment.type === EnvironmentType.SharePoint) {
        const currentUserDisplayName = this.props.pageContext.user.displayName;
        const currentUserPersona = await Utility.getPersonaForPeoplePickerField(currentUserDisplayName, this.props.pageContext.web.absoluteUrl);
        this.updateFormData('PurchaserName', currentUserPersona);
      }
    }

  }

  /**
   * @description a callback function passed down as props to components enabling them to update this component's state 
   * @param field
   */
  private onError = (field: string, valid?: boolean): void => {

    const checkExists: { exists: boolean, index: number } = Utility.checkIfErrorExists(field, this.state.errors);
    const exists = checkExists.exists;
    // check if field is valid and if field already exists in array
    if (valid && exists) {
      // remove field from array
      let errors = this.state.errors;
      const i = checkExists.index;
      errors.splice(i, 1);
      this.setState({ errors: errors });
    } else if (!valid && !exists) {
      // field does not exist in errors array AND field is not valid
      // add field to errors array
      let errors = this.state.errors;
      errors.push(field);
      this.setState({ errors: errors });
    }
  }

  /**
   * @description callback function passed down as props to components enabling them to update this component's state
   * @param field under which to store the value
   * @param value input by the user
   * @param isAttachment if true the value is a file
   */
  private async updateFormData(field: string, value: any, isAttachment?: boolean): Promise<void> {
    // check if value is a file
    if (isAttachment) {
      let formData = { ...this.state.formData };
      let attachments: any[] = [];
      if (formData['Attachments']) {
        attachments = formData['Attachments'];
      }
      // if the file has been removed
      if (!value) {
        for (let i = 0; attachments.length; i++) {
          if (attachments[i]['name'] == field) {
            attachments.splice(i, 1);
            break;
          }
        }
        // if there is a file
      } else {
        attachments.push({
          name: field,
          file: value
        });
      }
      formData['Attachments'] = attachments;
      this.setState({ formData: formData });

      // if the value is NOT a file and is different from the most recently entered value
    } else if (this.state.formData[field] !== value) {
      let formData = { ...this.state.formData };
      formData[field] = value;
      await this.setState({ formData: formData });

      // then, run autocomplete rules
      this.runtimeAutocompleteRules(field);

    }
  }

  private runtimeAutocompleteRules = (field: string): void => {
    switch (field) {
      case 'ProcurementRoute':
        if (this.state.formData[field] == 2) {
          // * Procurement Route = FTDP
          this.updateFormData('ProcurementStrategy', 'Fast track');
          this.updateFormData('StaProcurement', false);
        } else if (this.state.formData[field] == 3) {
          // * Procurement Route = STA
          this.updateFormData('ProcurementStrategy', '');
          this.updateFormData('StaProcurement', true);
        } else if (this.state.formData[field] == 1) {
          // * Procurement Route = RTP
          this.updateFormData('ProcurementStrategy', '');
          this.updateFormData('StaProcurement', false);
        }
        break;
      case 'SubProjectCode':
        if (this.state.formData[field] && this.state.formData['TotalContractValue']) {
          this.updateBudgetCheckAndFinancialApproverSections(this.state.formData[field], this.state.formData['TotalContractValue']);
        }
        break;
      case 'ContractValue':
        this.runtimeUpdateTotalContractValueForVariation();
        break;
      case 'Variation':
        if (this.state.formData[field]) {
          // load data for variation 
          this.runtimeLoadVariationData();
        } else if (this.state.formData['PreviousContractValue']) {
          // empty PreviousContractValue
          this.runtimeUnloadVariation();
        }
        break;
      case 'VariationCheck':
        if (!this.state.formData[field]) {
          this.runtimeUnloadVariation();
        }
        break;
      case 'TotalContractValue':
        if (this.state.formData[field] != '0' && this.state.formData['SubProjectCode']) {
          this.updateBudgetCheckAndFinancialApproverSections(this.state.formData['SubProjectCode'], this.state.formData[field]);
        } else if (this.state.formData['SubProjectCode']) {
          this.emptyBudgetCheckAndFinancialApproverSections();
        }
        break;
    }

  }

  private runtimeUnloadVariation() {
    let formData: IFormData = { ...this.state.formData };
    formData['PreviousContractValue'] = 0;
    formData['Variation'] = undefined;
    this.setState({ formData: formData });
    this.runtimeUpdateTotalContractValueForVariation();
  }

  /**
   * @description loads the form data for the form selected
   */
  private async runtimeLoadVariationData() {
    // console.log('1 - runtimeLoadVariationData called');
    const purchaser = this.state.formData['PurchaserName'];
    const variation = this.state.formData['Variation'];
    const formId = Utility.getFormIdFromProcReference(this.state.formData['Variation']);
    let dataToLoad = await Form.loadVariation(this.props.listName, formId, this.props.pageContext);
    // console.log('10 - variation data to load retrieved: ' + JSON.stringify(dataToLoad));
    let formData = dataToLoad.formData;
    formData['PreviousContractValue'] = formData['TotalContractValue'];
    formData['ContractValue'] = 0;
    formData['PurchaserName'] = purchaser;
    formData['Variation'] = variation;
    formData['ProcurementStatus'] = 'Draft';
    formData['VariationCheck'] = true;
    formData['ProcurementRoute'] = 3;
    const allDropDownOptions: IDropDownOptions = { ...this.state.dropDownOptions, ...dataToLoad.dropDownOptions[0] };
    this.setState({
      formData: formData,
      dropDownOptions: allDropDownOptions
    });
  }

  /**
   * @description adds the previous contract value to the current contract value
   */
  private runtimeUpdateTotalContractValueForVariation() {
    let totalContractValue;
    if (this.state.formData['Variation']) {
      if (this.state.formData['ContractValue']) {
        totalContractValue = parseInt(this.state.formData['PreviousContractValue']) + parseInt(this.state.formData['ContractValue']);
      } else {
        totalContractValue = parseInt(this.state.formData['PreviousContractValue']);
      }
    } else {
      if (this.state.formData['ContractValue'] == undefined) {
        totalContractValue = "0";
      } else {
        totalContractValue = this.state.formData['ContractValue'];
      }
    }
    this.updateFormData('TotalContractValue', totalContractValue);
  }

  /**
   * @description updates state with looked up values from Agresso extract in SharePoint lists to auto-populate form fields
   * @fires from this.runAutocompleteRules()
   * @param subProjectCode 
   * @param contractValue 
   */
  private async updateBudgetCheckAndFinancialApproverSections(subProjectCode, totalContractValue): Promise<void> {
    let values = await AgressoDataLookup.getValuesForBudgetCheckAndFinancialApproverFields(subProjectCode, totalContractValue);
    if (values) {
      let formData = { ...this.state.formData };
      let dropDownOptions = { ...this.state.dropDownOptions };
      formData['SubProjectCodeDescription'] = values.subProjectCodeDescription;
      formData['CostCentre'] = values.costCentre;
      formData['ProjectCode'] = values.projectCode;
      formData['PrimaryBusinessManager'] = values.primaryBusinessManager;
      formData['PrimaryFinancialApprover'] = values.primaryFinancialApprover;
      formData['DirectorateOfPurchase'] = values.purchasingDirectorate;
      dropDownOptions['SecondaryFinancialApprover'] = values.possibleSecondaryFinancialApprovers;
      // update state
      this.setState({
        formData: formData,
        dropDownOptions: dropDownOptions,
      });
    } else {
      this.emptyBudgetCheckAndFinancialApproverSections();
    }
  }

  private emptyBudgetCheckAndFinancialApproverSections = (): void => {
    let formData = { ...this.state.formData };
    let dropDownOptions = { ...this.state.dropDownOptions };

    formData['SubProjectCodeDescription'] = "";
    formData['CostCentre'] = "0";
    formData['ProjectCode'] = "0";
    formData['PrimaryBusinessManager'] = "";
    formData['PrimaryFinancialApprover'] = "";
    formData['DirectorateOfPurchase'] = "";
    formData['SecondaryFinancialApprover'] = "";
    dropDownOptions['SecondaryFinancialApprover'] = [];
    // update state
    this.setState({
      formData: formData,
      dropDownOptions: dropDownOptions,
    });
  }

  /**
   * @description generates buttons on the command bar
   * @returns array of button objects
   */
  private getButtons = (): IContextualMenuItem[] => {

    const save = {
      key: 'save',
      name: 'Save',
      iconProps: {
        iconName: 'Save'
      },
      onClick: this.handleSave,
    };

    const submit = {
      key: 'submit',
      name: 'Submit',
      iconProps: {
        iconName: 'Send'
      },
      onClick: this.handleSubmit,
    };

    const discard = {
      key: 'discard',
      name: 'Discard',
      iconProps: {
        iconName: 'Delete'
      },
      onClick: this.handleDiscard,
    };

    const errors = {
      key: 'validation',
      name: 'Validation Errors',
      iconProps: {
        iconName: 'ErrorBadge'
      },
    };

    let buttons: any[] = [];

    if (this.props.formContext.mode == "new") {
      if (this.state.userInfo.userRole == "editor") {
        if (this.state.formData['PurchaserName'] && this.state.formData['PurchaserName'][0]) {
          buttons.push(save);
        }
        if (this.state.errors.length > 0) {
          buttons.push(errors);
        } else {
          buttons.push(submit);
        }
      }
    } else {
      // form is edit mode
      if (this.state.formData['ProcurementStatus'] == "Draft" && this.state.userInfo.userRole == "editor") {
        buttons.push(save, discard);
        if (this.state.errors.length > 0) {
          buttons.push(errors);
        } else {
          buttons.push(submit);
        }
      } else if (this.state.formData['ProcurementStatus'] == "Returned" && this.state.userInfo.userRole == "editor") {
        if (this.state.errors.length > 0) {
          buttons.push(errors);
        } else {
          buttons.push(save);
        }
      } else if (this.state.formData['ProcurementStatus'] == "Submitted" && this.state.userInfo.isCommercialMember) {
        buttons.push(save);
      } else if (this.state.formData['ProcurementStatus'] == "Approved" && this.state.userInfo.isCommercialMember) {
        buttons.push(save);
      }
    }

    return buttons;
  }

  /**
   * @description runs when the save button is clicked  
   */
  private async handleSave() {
    this.setState({ isSubmitting: true });
    await Form.save(this.props.formContext, this.props.pageContext, this.state.formData, this.props.listName);
    this.setState({ isSubmitting: false });
    if (Environment.type === EnvironmentType.SharePoint) {
      window.location.href = this.props.pageContext.web.serverRelativeUrl;
    }
  }

  private async handleDiscard() {
    let confirmation: boolean = confirm("this will also delete the folder and any documents, do you want to continue?");
    if (confirmation) {
      this.setState({ isSubmitting: true });
      await Form.discard(this.props.formContext, this.props.listName, this.props.pageContext, this.state.formData);
      this.setState({ isSubmitting: false });
      if (Environment.type === EnvironmentType.SharePoint) {
        window.location.href = this.props.pageContext.web.serverRelativeUrl;
      }
    }
  }

  /**
   * @description the FileApi returns an array of added attachments, but the function does not wait to hear back from the FileApi   
   * 
  */
  private async handleSubmit() {
    this.setState({ isSubmitting: true });
    await Form.submit(this.props.formContext, this.props.pageContext, this.state.formData, this.props.listName);
    this.setState({ isSubmitting: false });
    if (Environment.type === EnvironmentType.SharePoint) {
      window.location.href = this.props.pageContext.web.serverRelativeUrl;
    }
  }

  public render(): React.ReactElement<IRtpFormProps> {

    let { formData, dropDownOptions, userInfo, loading, isSubmitting } = this.state;

    let formToDisplay: any;

    if (userInfo.userRole == "none") {
      formToDisplay =
        <div>
          <span>You do not have permissions to view the form.</span>
        </div>;
    }

    if (userInfo.userRole != "none" && loading) {
      formToDisplay =
        <div>
          <Spinner label={"Please wait, loading form..."} />
        </div>;
    }

    if (isSubmitting) {
      formToDisplay =
        <div>
          <Spinner label={"Saving, please do not close the tab! This page will redirect you after the form has been saved..."} />
        </div>;
    }

    if (userInfo.userRole != "none" && !loading && !isSubmitting) {
      formToDisplay =
        <div>
          <CommandBar
            items={this.getButtons()}
          />
          <div className={styles.rtpForm}>
            <div className={styles.container}>
              <PurchaserSection
                formData={formData}
                onUpdate={this.updateFormData}
                context={this.props.context}
                dropDownOptions={dropDownOptions}
                onError={this.onError}
                disabled={formData['ProcurementStatus'] == "Submitted" || formData['ProcurementStatus'] == "Approved"}
              />
              <FTDPSection
                formData={formData}
                onUpdate={this.updateFormData}
                dropDownOptions={dropDownOptions}
                onError={this.onError}
                disabled={formData['ProcurementStatus'] == "Submitted" || formData['ProcurementStatus'] == "Approved"}
              />
              <ContractDetailsSection
                formData={formData}
                onUpdate={this.updateFormData}
                context={this.props.context}
                listName={this.props.listName}
                formContext={this.props.formContext}
                dropDownOptions={dropDownOptions}
                onError={this.onError}
                disabled={formData['ProcurementStatus'] == "Submitted" || formData['ProcurementStatus'] == "Approved"}
              />
              <RequirementSection
                formData={formData}
                onUpdate={this.updateFormData}
                context={this.props.context}
                listName={this.props.listName}
                formContext={this.props.formContext}
                dropDownOptions={dropDownOptions}
                onError={this.onError}
                disabled={formData['ProcurementStatus'] == "Submitted" || formData['ProcurementStatus'] == "Approved"}
              />
              <StaSection
                formData={formData}
                onUpdate={this.updateFormData}
                context={this.props.context}
                listName={this.props.listName}
                formContext={this.props.formContext}
                dropDownOptions={dropDownOptions}
                onError={this.onError}
                disabled={formData['ProcurementStatus'] == "Submitted" || formData['ProcurementStatus'] == "Approved"}
              />
              <IASection
                formData={formData}
                onUpdate={this.updateFormData}
                dropDownOptions={dropDownOptions}
                onError={this.onError}
                disabled={formData['ProcurementStatus'] == "Submitted" || formData['ProcurementStatus'] == "Approved"}
              />
              <BudgetCheckSection
                formData={formData}
                onUpdate={this.updateFormData}
                dropDownOptions={dropDownOptions}
                onError={this.onError}
                disabled={formData['ProcurementStatus'] == "Submitted" || formData['ProcurementStatus'] == "Approved"}
              />
              <CommercialSection
                formData={formData}
                onUpdate={this.updateFormData}
                context={this.props.context}
                dropDownOptions={dropDownOptions}
                onError={this.onError}
                disabled={!this.state.userInfo.isCommercialMember}
                formContext={this.props.formContext}
              />
            </div>
          </div>
        </div>;
    }
    return formToDisplay;
  }

}





