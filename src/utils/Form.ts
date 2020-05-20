import { IFormContext, IFormData } from '../interfaces';
import { PageContext } from '@microsoft/sp-page-context';
import { Utility, FileApi, ListItemApi } from '.';
import { IPersona, ITag } from 'office-ui-fabric-react/lib';
import { sp, PagedItemCollection, PermissionKind, BasePermissions, SharePointQueryable, SharePointQueryableSecurable } from '@pnp/sp';
import { IListItem } from '../interfaces/IListItem';
import { IDropDownOption, IDropdownFieldsOptions, IComboFieldsOptions, ISubProjectCodeOptions, IVariationOptions } from '../interfaces/IRtpFormState';
import { AgressoDataLookup } from './AgressoDataLookup';
import { Environment, EnvironmentType } from '@microsoft/sp-core-library';
import { MockDropdownFieldsOptions, MockComboFieldsOptions } from '../mock/MockDropdowns';
import { MockFormData } from '../mock/MockFormData';
import { MockListItem } from '../mock/MockListItem';

interface SecondaryFinancialApproverOptions {
  SecondaryFinancialApprover: IDropDownOption[];
}

interface DropDownOptions extends IDropdownFieldsOptions, IComboFieldsOptions {
}


/**
 * @description details the workflows for the form
 */
export class Form {

  private static choiceFields = [
    'ProcurementRoute',
    'StaType',
    'ApprovedScope',
    'RouteToMarket',
    'ProcurementType',
    'SecondaryBusinessManager',
    'DpoOptions',
    'IsmOptions',
    'BusinessCaseRequirement',
    'UnableProcureCompetitively',
    'SuggestedTendererStatus',
    'BusinessManagerDirectorates'
  ];

  private static comboFields = [
    'SubProjectCode',
  ];

  // columns to get from ListItemApi.get()
  public static listColumns: string[] = [
    "SpokenToCommercial",
    "StaProcurement",
    "ProcurementRoute/ID",
    "ProcurementRoute/Title",
    "ProcurementStatus",
    "ProcReference",
    "PurchaserName/ID",
    "PurchaserName/Title",
    "PurchaserName/JobTitle",
    "PurchaserName/UserName",
    "PurchaserName/Department",
    "PurchaserName/EMail",
    "ContractManager/ID",
    "ContractManager/Title",
    "ContractManager/JobTitle",
    "ContractManager/UserName",
    "ContractManager/Department",
    "PurchaseCategory/ID",
    "PurchaseCategory/Title",
    "BusinessCaseAttached",
    "BusinessCaseRequirement/ID",
    "BusinessCaseRequirement/Title",
    "ProcurementStartDate",
    "ProcurementEndDate",
    "ContractTitle",
    "ContractValue",
    "TotalContractValue",
    "PreviousContractValue",
    "ReasonsForPurchase",
    "ApprovedScope/ID",
    "ApprovedScope/Title",
    "RouteToMarket/ID",
    "RouteToMarket/Title",
    "OtherFrameworkName",
    "ContractManagerDirectorate/ID",
    "ContractManagerDirectorate/Title",
    "PurchaserNameDirectorate/ID",
    "PurchaserNameDirectorate/Title",
    "StaType/ID",
    "StaType/Title",
    "UnableProcureCompetitively",
    "UnableProcureCompetitivelyReason",
    "SupportingDetails",
    "StaAlternatives",
    "SuggestedTendererStatus/ID",
    "SuggestedTendererStatus/Title",
    "CompanyName",
    "ContactName",
    "TendererContactEmail",
    "TendererContactTelephone",
    "DpoOptions/ID",
    "DpoOptions/Title",
    "IsmOptions/ID",
    "IsmOptions/Title",
    "SubProjectCode",
    "SubProjectCodeDescription",
    "ProjectCode",
    "CostCentre",
    "DirectorateOfPurchase",
    "PrimaryBusinessManager/Title",
    "PrimaryBusinessManager/ID",
    "SecondaryBusinessManager/Title",
    "SecondaryBusinessManager/ID",
    "PrimaryFinancialApprover/Title",
    "PrimaryFinancialApprover/ID",
    "SecondaryFinancialApprover/Title",
    "SecondaryFinancialApprover/ID",
    "StaCompliant",
    "CommercialManager/ID",
    "CommercialManager/Title",
    "CommercialManager/JobTitle",
    "CommercialManager/UserName",
    /* "ContractNumber", 
    "ProcurementStrategy", 
    "ProcurementType/ID", 
    "ProcurementType/Title", 
    "ProcurementNotes", 
    "ReasonForRejection",
    "VariationCheck",
    "Variation", */
  ];

  public static listColumns2: string[] = [
    "ContractNumber",
    "ProcurementStrategy",
    "ProcurementType/ID",
    "ProcurementType/Title",
    "ProcurementNotes",
    "ReasonForRejection",
    "VariationCheck",
    "Variation",
    "RedraftingNotPossible"
  ];

  // columns to expand in ListItemApi.get()
  public static expandListColumns: string[] = [
    "PurchaserNameDirectorate",
    "ContractManagerDirectorate",
    "PurchaserName",
    "ContractManager",
    "ProcurementRoute",
    "PurchaseCategory",
    "BusinessCaseRequirement",
    "ApprovedScope",
    "RouteToMarket",
    "StaType",
    "SuggestedTendererStatus",
    "DpoOptions",
    "IsmOptions",
    // "ProcurementType", 
    "PrimaryBusinessManager",
    "SecondaryBusinessManager",
    "PrimaryFinancialApprover",
    "SecondaryFinancialApprover",
    "CommercialManager"
  ];

  public static expandListColumns2: string[] = [
    "ProcurementType",
  ];

  /**
   * @description details the workflow for loading the form
   */
  public static async load(formContext: IFormContext, pageContext: PageContext, listName: string): Promise<{ userInfo: { userRole: string, isCommercialMember: boolean }, dropDownOptions: any[], formData: any }> {

    // if new form 
    if (formContext.mode == "new") {
      // let userIsEditor = await this.checkUserRoleIsEditorForNewForm(listName);
      /* if (!userIsEditor) {
        // user is not editor
        console.log("form is in new mode");
        console.log("user is not editor");
        return {
          userInfo: {
            userRole: "none",
            isCommercialMember: false,
          },
          dropDownOptions: [],
          formData: null
        };
      } else {
        // user is edtor 
      } */
      console.log("form is in new mode");
      console.log("user is editor");
      // check is user belongs to Commercial team
      const isCommercialMember: Promise<boolean> = this.checkIfUserIsMemberOfCommercial();
      // get drop down options for fields
      const dropdownOptions: Promise<DropDownOptions[]> = this.overseeGetDropdownOptions();
      const formData = null;
      return Promise.all([isCommercialMember, dropdownOptions])
        .then(response => {
          return {
            userInfo: {
              userRole: "editor",
              isCommercialMember: response[0],
            },
            dropDownOptions: response[1],
            formData: formData
          };
        });
    } else {
      // form is edit mode 
      console.log("form is in edit mode");
      let userRole: string;
      // get user role on list item
      try {
        const listItemUserRole: string = await this.checkUserRoleListItemForEditForm(listName, formContext.formId);
        if (listItemUserRole == "none") {
          console.log("user role is none");
          return {
            userInfo: {
              userRole: "none",
              isCommercialMember: false,
            },
            dropDownOptions: [],
            formData: null
          };
        } else {
          // user role is either editor or viewer
          // get formData
          const formData: IFormData = await this.getFormData(listName, formContext, pageContext);
          console.log("retrieved formData");
          // check form status 
          let documentsUserRole;
          if (formData['ProcurementStatus'] == "Draft") {
            console.log("form status is draft, use form guid to retrieve folder");
            documentsUserRole = await this.checkUserRoleDocumentsForEditForm(formContext.formGuid, pageContext.web.serverRelativeUrl);
          } else {
            console.log("form status is not draft, use proc reference to retrieve folder");
            // check user role for parent folder 
            documentsUserRole = await this.checkUserRoleDocumentsForEditForm(formData['ProcReference'], pageContext.web.serverRelativeUrl);
          }

          if (documentsUserRole == "editor" && listItemUserRole == "editor") {
            console.log("user role is editor");
            userRole = "editor";
            // check is user belongs to Commercial team
            const isCommercialMemberPm: Promise<boolean> = this.checkIfUserIsMemberOfCommercial();
            // get drop down options for fields
            const dropdownOptionsPm: Promise<DropDownOptions[]> = this.overseeGetDropdownOptions();
            const secondaryFinancialApproverOptionsPm: Promise<SecondaryFinancialApproverOptions[]> = this.getSecondaryFinancialApproverOptionsOnFormLoad(formData);
            return Promise.all([isCommercialMemberPm, dropdownOptionsPm, secondaryFinancialApproverOptionsPm])
              .then(response => {
                return {
                  userInfo: {
                    userRole: userRole,
                    isCommercialMember: response[0],
                  },
                  dropDownOptions: [...response[1], ...response[2]],
                  formData: formData
                };
              });
          } else if (documentsUserRole == "viewer" && listItemUserRole == "viewer") {
            console.log("user role is viewer");
            userRole = "viewer";
            // get drop down options for fields
            const dropdownOptionsPm: Promise<[IDropdownFieldsOptions, IComboFieldsOptions]> = this.overseeGetDropdownOptions();
            const secondaryFinancialApproverOptionsPm: Promise<SecondaryFinancialApproverOptions[]> = this.getSecondaryFinancialApproverOptionsOnFormLoad(formData);
            return Promise.all([dropdownOptionsPm, secondaryFinancialApproverOptionsPm])
              .then(response => {
                return {
                  userInfo: {
                    userRole: userRole,
                    isCommercialMember: false,
                  },
                  dropDownOptions: [...response[0], ...response[1]],
                  formData: formData
                };
              });
          } else if (documentsUserRole == "editor" && listItemUserRole == "viewer") {
            console.log("user role is viewer but user can add new attachments");
            userRole = "viewer";
            // get drop down options for fields
            const dropdownOptionsPm: Promise<DropDownOptions[]> = this.overseeGetDropdownOptions();
            const secondaryFinancialApproverOptionsPm: Promise<SecondaryFinancialApproverOptions[]> = this.getSecondaryFinancialApproverOptionsOnFormLoad(formData);
            return Promise.all([dropdownOptionsPm, secondaryFinancialApproverOptionsPm])
              .then(response => {
                return {
                  userInfo: {
                    userRole: userRole,
                    isCommercialMember: false,
                  },
                  dropDownOptions: [...response[0], ...response[1]],
                  formData: formData
                };
              });
          } else {
            console.log('user does not have permissions to the parent folder');
          }
        }
      } catch (error) {
        console.log('an error occured while loading an existing form, error: ' + error);
      }
    }
  }

  /**
   * @description checks which options should be included 
   * @fires this.load() when form is in edit mode only
   * @param formData 
   * @returns an array, containing an array of option(s)
   */
  private static async getSecondaryFinancialApproverOptionsOnFormLoad(formData: IFormData): Promise<SecondaryFinancialApproverOptions[]> {
    console.log('8 - getSecondaryFinancialApproverOptionsOnFormLoad called');
    let secondaryFinancialApproverOptions: SecondaryFinancialApproverOptions[] = [{
      SecondaryFinancialApprover: []
    }];

    if (formData['SecondaryFinancialApprover']) {
      // a secondary financial approver has already been selected
      secondaryFinancialApproverOptions[0].SecondaryFinancialApprover.push({
        key: formData['SecondaryFinancialApprover'],
        text: formData['SecondaryFinancialApprover']
      });
    } else if (formData['SubProjectCode'] && formData['ContractValue']) {
      // a sub project code and contract value have been selected but the user ...
      // ...has not yet selected a secondary financial approver
      // so retrieve possible secondary financial approvers 
      const approvalLevel: string = await AgressoDataLookup.getApprovalLevel(formData['ContractValue'].toString());
      const options: IDropDownOption[] = await AgressoDataLookup.getPossibleSecondaryFinancialApproversInApprovalLevel(approvalLevel, formData['DirectorateOfPurchase']);
      secondaryFinancialApproverOptions[0].SecondaryFinancialApprover = options;
    }
    return secondaryFinancialApproverOptions;
  }

  public static async getFormData(listName: string, formContext: IFormContext, pageContext: PageContext): Promise<IFormData> {
    if (Environment.type === EnvironmentType.Local) {
      return MockFormData;
    } else if (Environment.type === EnvironmentType.SharePoint) {
      console.log('getFormData called');
      const listItem1: any = await ListItemApi.get(listName, formContext.formId, this.listColumns, this.expandListColumns);
      console.log('list item 1: ' + JSON.stringify(listItem1));
      const listItem2: any = await ListItemApi.get(listName, formContext.formId, this.listColumns2, this.expandListColumns2);
      console.log('list item 2: ' + JSON.stringify(listItem2));
      const listItem: IListItem = { ...listItem1, ...listItem2 };
      console.log('list item: ' + JSON.stringify(listItem));
      return await this.deserialise(listItem, pageContext.web.absoluteUrl);
    }
  }

  public static async loadVariation(listName: string, formId: number, pageContext: PageContext): Promise<{ dropDownOptions: any[], formData: any }> {
    console.log('2 - loadVariation called');
    const formData: IFormData = await this.getVariationFormData(listName, formId, pageContext);
    console.log('7 - retrieved formData: ' + JSON.stringify(formData));
    const secondaryFinancialApproverOptionsPm: Promise<SecondaryFinancialApproverOptions[]> = this.getSecondaryFinancialApproverOptionsOnFormLoad(formData);
    return await Promise.all([secondaryFinancialApproverOptionsPm])
      .then(response => {
        console.log('9 - retrieved secondaryFinancialApproverOptions: ' + JSON.stringify(response[0]));
        return {
          dropDownOptions: [...response[0]],
          formData: formData,
        };
      });
  }

  private static async getVariationFormData(listName: string, formId: number, pageContext: PageContext): Promise<IFormData> {
    if (Environment.type === EnvironmentType.Local) {
      return MockFormData;
    } else if (Environment.type === EnvironmentType.SharePoint) {
      console.log('3 - get variation form data called');
      const listItem1: any = await ListItemApi.get(listName, formId, this.listColumns, this.expandListColumns);
      const listItem2: any = await ListItemApi.get(listName, formId, this.listColumns2, this.expandListColumns2);
      const listItem: IListItem = { ...listItem1, ...listItem2 };
      console.log('4 - item fetched');
      // console.log('PurchaserName: ' + JSON.stringify(listItem['PurchaserName']));
      return await this.deserialise(listItem, pageContext.web.absoluteUrl);
    }
  }


  /**
   * @description details the workflow for saving a form (new or existing)
   * @param formContext 
   * @param pageContext 
   * @param formData 
   * @param listName 
   */
  public static async save(formContext: IFormContext, pageContext: PageContext, formData: IFormData, listName: string): Promise<void> {
    const listItem = await this.serialise(formData, formContext);
    if (formContext.mode == "new") {
      // add item and get list item id as form id
      const formId: number = await ListItemApi.add(pageContext.web.serverRelativeUrl, listName, formContext.formGuid, listItem);
      if (formId) {
        // await set permissions for purchaser and Commercial to edit the list item
        const listItemPermissionsPm = ListItemApi.setPermissionsOnSave(listName, formId, formData);
        // handle attachments 
        const attachmentsPm = FileApi.handleAttachmentsOnSaveForNewForm(formContext.formGuid, pageContext.web.serverRelativeUrl, formData);
        await Promise.all([listItemPermissionsPm, attachmentsPm]);
      }
      // handle attachments  
    } else if (formContext.mode == "edit") {
      // update list item
      const listItemUpdatePm = ListItemApi.update(listName, formContext.formId, listItem);
      // handle attachments    
      const attachmentsPm = FileApi.handleAttachmentsOnSaveForEditForm(formContext.formGuid, pageContext.web.serverRelativeUrl, formData);
      await Promise.all([listItemUpdatePm, attachmentsPm]);
    }
  }

  /**
   * @description details the workflow for discarding a form (new or existing)
   * @fires from discard button on Command Bar
   * @param formContext
   * @param listName
   * @param formGuid
   * @param serverRelativeUrl
   */
  public static async discard(formContext: IFormContext, listName: string, pageContext: PageContext, formData: any): Promise<void> {
    // delete the existing list item & delete the existing parent folder 
    const deleteListItemPm = ListItemApi.delete(listName, formContext.formId);
    const removeFolderPm = FileApi.removeParentFolder(pageContext.web.serverRelativeUrl, formContext.formGuid, formData);
    await Promise.all([deleteListItemPm, removeFolderPm]);
  }

  /**
     * @description details the workflow for submitting a form (new or existing)
     * @param formContext 
     * @param pageContext 
     * @param formData 
     * @param listName 
     */
  public static async submit(formContext: IFormContext, pageContext: PageContext, formData: IFormData, listName: string): Promise<any> {
    // const formType = this.getFormType(formData);
    const listItem = await this.serialise(formData, formContext);
    if (formContext.mode == "new") {
      // add item and get list item id as form id
      const formId: number = await ListItemApi.add(pageContext.web.serverRelativeUrl, listName, formContext.formGuid, listItem);
      // generate proc reference  
      const procRef = this.generateProcReference(formId, formData);
      // await update list item with proc ref and procurement status = submitted
      await ListItemApi.updateAsSubmitted(listName, formId, procRef);
      // await set permissions for purchaser and Commercial to edit the list item
      await ListItemApi.setPermissionsOnSave(listName, formId, formData);
      // then, update permissions to allow bus mgrs and fin approvers to read the list item
      const listItemPermsPm = ListItemApi.updatePermissionsOnSubmit(listName, formId, formData);
      // handle attachments
      const attachmentsPm = FileApi.handleAttachmentsOnSubmitForNewForm(pageContext.web.serverRelativeUrl, formData, procRef);
      return await Promise.all([listItemPermsPm, attachmentsPm]);

    } else if (formContext.mode == "edit") {
      // generate proc reference
      const procRef = this.generateProcReference(formContext.formId, formData);
      // await update list item with new data
      await ListItemApi.update(listName, formContext.formId, listItem);
      // then, await update list item with proc ref and procurement status = submitted
      await ListItemApi.updateAsSubmitted(listName, formContext.formId, procRef);
      // then, update permissions to allow bus mgrs and fin approvers to read the list item
      const listItemPermsPm = ListItemApi.updatePermissionsOnSubmit(listName, formContext.formId, formData);
      // handle attachments
      // ! error uncaught in promise 
      const attachmentsPm = FileApi.handleAttachmentsOnSubmitForEditForm(formContext.formGuid, pageContext.web.serverRelativeUrl, formData, procRef);
      return await Promise.all([listItemPermsPm, attachmentsPm]);
    }
  }

  private static async serialise(formData: IFormData, formContext: IFormContext): Promise<IListItem> {
    if (Environment.type === EnvironmentType.Local) {
      return MockListItem;
    }
    else if (Environment.type === EnvironmentType.SharePoint) {
      let PurchaserName = (formData['PurchaserName'] && formData['PurchaserName'][0]) ? await Utility.getSharePointUserId(formData['PurchaserName'][0]['optionalText']) : null;
      let ContractManager = (formData['ContractManager'] && formData['ContractManager'][0]) ? await Utility.getSharePointUserId(formData['ContractManager'][0]['optionalText']) : null;
      let CommercialManager = (formData['CommercialManager'] && formData['CommercialManager'][0]) ? await Utility.getSharePointUserId(formData['CommercialManager'][0]['optionalText']) : null;
      let PrimaryBusinessManager = formData['PrimaryBusinessManager'] ? await Utility.getSharePointUserId('i:0#.f|membership|' + formData['PrimaryBusinessManager']) : null;
      let SecondaryBusinessManager = formData['SecondaryBusinessManager'] ? await Utility.getSharePointUserId('i:0#.f|membership|' + formData['SecondaryBusinessManager']) : null;
      let PrimaryFinancialApprover = formData['PrimaryFinancialApprover'] ? await Utility.getSharePointUserId('i:0#.f|membership|' + formData['PrimaryFinancialApprover']) : null;
      let SecondaryFinancialApprover = formData['SecondaryFinancialApprover'] ? await Utility.getSharePointUserId('i:0#.f|membership|' + formData['SecondaryFinancialApprover']) : null;
      let PurchaseCategory = await this.getSelectedTagFromFormData(formData, 'PurchaseCategory');

      return {
        PurchaserNameDirectorateId: formData['PurchaserNameDirectorate'],
        ContractManagerDirectorateId: formData['ContractManagerDirectorate'],
        SpokenToCommercial: formData['SpokenToCommercial'],
        FormGuid: formContext.formGuid,
        ProcReference: formData['ProcReference'],
        StaProcurement: formData['StaProcurement'],
        ProcurementRouteId: formData['ProcurementRoute'],
        ProcurementStatus: formData['ProcurementStatus'],
        PurchaserNameId: PurchaserName,
        ContractManagerId: ContractManager,
        PurchaseCategoryId: PurchaseCategory,
        BusinessCaseAttached: formData['BusinessCaseAttached'],
        BusinessCaseRequirementId: formData['BusinessCaseRequirement'],
        ProcurementStartDate: formData['ProcurementStartDate'],
        ProcurementEndDate: formData['ProcurementEndDate'],
        ContractTitle: formData['ContractTitle'],
        PreviousContractValue: formData['PreviousContractValue'],
        ContractValue: formData['ContractValue'],
        TotalContractValue: formData['TotalContractValue'],
        ReasonsForPurchase: formData['ReasonsForPurchase'],
        ApprovedScopeId: formData['ApprovedScope'],
        RouteToMarketId: formData['RouteToMarket'],
        OtherFrameworkName: formData['OtherFrameworkName'],
        StaTypeId: formData['StaType'],
        UnableProcureCompetitively: formData['UnableProcureCompetitively'],
        UnableProcureCompetitivelyReason: formData['UnableProcureCompetitivelyReason'],
        SuggestedTendererStatusId: formData['SuggestedTendererStatus'],
        CompanyName: formData['CompanyName'],
        ContactName: formData['ContactName'],
        TendererContactEmail: formData['TendererContactEmail'],
        TendererContactTelephone: formData['TendererContactTelephone'],
        DpoOptionsId: formData['DpoOptions'],
        IsmOptionsId: formData['IsmOptions'],
        SubProjectCode: formData['SubProjectCode'],
        SubProjectCodeDescription: formData['SubProjectCodeDescription'],
        ProjectCode: formData['ProjectCode'],
        CostCentre: formData['CostCentre'],
        DirectorateOfPurchase: formData['DirectorateOfPurchase'],
        PrimaryBusinessManagerId: PrimaryBusinessManager,
        SecondaryBusinessManagerId: SecondaryBusinessManager,
        PrimaryFinancialApproverId: PrimaryFinancialApprover,
        SecondaryFinancialApproverId: SecondaryFinancialApprover,
        StaCompliant: formData['StaCompliant'],
        CommercialManagerId: CommercialManager,
        ContractNumber: formData['ContractNumber'],
        ProcurementStrategy: formData['ProcurementStrategy'],
        ProcurementTypeId: formData['ProcurementType'],
        ProcurementNotes: formData['ProcurementNotes'],
        ReasonForRejection: formData['ReasonForRejection'],
        Variation: formData['Variation'],
        VariationCheck: formData['VariationCheck'],
        RedraftingNotPossible: formData['RedraftingNotPossible'],
      };
    }
  }

  private static async deserialise(item: IListItem, absoluteUrl: string): Promise<IFormData> {


    console.log('5 - deserialise called');

    let PurchaserNamePm: Promise<IPersona[]>;
    let ContractManagerPm: Promise<IPersona[]>;
    let PrimaryBusinessManagerPm: Promise<string>;
    let SecondaryBusinessManagerPm: Promise<string>;
    let PrimaryFinancialApproverPm: Promise<string>;
    let SecondaryFinancialApproverPm: Promise<string>;
    let CommercialManagerPm: Promise<IPersona[]>;

    if (item['PurchaserName']) {
      PurchaserNamePm = Utility.getPersonaForPeoplePickerField(item['PurchaserName']['Title'], absoluteUrl);
    }
    if (item['ContractManager']) {
      ContractManagerPm = Utility.getPersonaForPeoplePickerField(item['ContractManager']['Title'], absoluteUrl);
    }
    if (item['PrimaryBusinessManager']) {
      PrimaryBusinessManagerPm = Utility.getUserEmailFromDisplayName(item['PrimaryBusinessManager']['Title']);
    }
    if (item['SecondaryBusinessManager']) {
      SecondaryBusinessManagerPm = Utility.getUserEmailFromDisplayName(item['SecondaryBusinessManager']['Title']);
    }
    if (item['PrimaryFinancialApprover']) {
      PrimaryFinancialApproverPm = Utility.getUserEmailFromDisplayName(item['PrimaryFinancialApprover']['Title']);
    }
    if (item['SecondaryFinancialApprover']) {
      SecondaryFinancialApproverPm = Utility.getUserEmailFromDisplayName(item['SecondaryFinancialApprover']['Title']);
    }
    if (item['CommercialManager']) {
      CommercialManagerPm = Utility.getPersonaForPeoplePickerField(item['CommercialManager']['Title'], absoluteUrl);
    }
    let PurchaseCategoryTag = this.getTagForTagPickerFieldFromListItem(item, "PurchaseCategory");
    let BusinessCaseRequirementId = this.getIdForDropdownFieldFromListItem(item, "BusinessCaseRequirement");
    let ApprovedScopeId = this.getIdForDropdownFieldFromListItem(item, "ApprovedScope");
    let RouteToMarketId = this.getIdForDropdownFieldFromListItem(item, "RouteToMarket");
    let StaTypeId = this.getIdForDropdownFieldFromListItem(item, "StaType");
    let ProcurementRouteId = this.getIdForDropdownFieldFromListItem(item, "ProcurementRoute");
    let DpoOptionsId = this.getIdForDropdownFieldFromListItem(item, "DpoOptions");
    let IsmOptionsId = this.getIdForDropdownFieldFromListItem(item, "IsmOptions");
    let ProcurementTypeId = this.getIdForDropdownFieldFromListItem(item, "ProcurementType");
    let SuggestedTendererStatusId = this.getIdForDropdownFieldFromListItem(item, "SuggestedTendererStatus");
    let ContractManagerDirectorateId = this.getIdForDropdownFieldFromListItem(item, "ContractManagerDirectorate");
    let PurchaserNameDirectorateId = this.getIdForDropdownFieldFromListItem(item, "PurchaserNameDirectorate");

    let PurchaserName: IPersona[];
    let ContractManager: IPersona[];
    let PrimaryBusinessManager: string;
    let SecondaryBusinessManager: string;
    let PrimaryFinancialApprover: string;
    let SecondaryFinancialApprover: string;
    let CommercialManager: IPersona[];

    await Promise.all([PurchaserNamePm, ContractManagerPm, PrimaryBusinessManagerPm, SecondaryBusinessManagerPm, PrimaryFinancialApproverPm, SecondaryFinancialApproverPm, CommercialManagerPm])
      .then(response => {
        console.log('6 - response for getting users resolved during deserialisation: ' + JSON.stringify(response));
        PurchaserName = response[0];
        ContractManager = response[1];
        PrimaryBusinessManager = response[2] ? response[2] : null;
        SecondaryBusinessManager = response[3] ? response[3] : null;
        PrimaryFinancialApprover = response[4] ? response[4] : null;
        SecondaryFinancialApprover = response[5] ? response[5] : null;
        CommercialManager = response[6];
      });

    return {
      PurchaserNameDirectorate: PurchaserNameDirectorateId,
      ContractManagerDirectorate: ContractManagerDirectorateId,
      SpokenToCommercial: item.SpokenToCommercial,
      ProcurementRoute: ProcurementRouteId,
      StaProcurement: item.StaProcurement,
      ProcurementStatus: item.ProcurementStatus,
      ProcReference: item.ProcReference,
      PurchaserName: PurchaserName,
      ContractManager: ContractManager,
      PurchaseCategory: PurchaseCategoryTag,
      BusinessCaseAttached: item.BusinessCaseAttached,
      BusinessCaseRequirement: BusinessCaseRequirementId,
      ProcurementStartDate: item.ProcurementStartDate,
      ProcurementEndDate: item.ProcurementEndDate,
      ContractTitle: item.ContractTitle,
      ContractValue: item.ContractValue,
      PreviousContractValue: item.PreviousContractValue,
      TotalContractValue: item.TotalContractValue,
      ReasonsForPurchase: item.ReasonsForPurchase,
      ApprovedScope: ApprovedScopeId,
      RouteToMarket: RouteToMarketId,
      OtherFrameworkName: item.OtherFrameworkName,
      StaType: StaTypeId,
      UnableProcureCompetitively: item.UnableProcureCompetitively,
      UnableProcureCompetitivelyReason: item.UnableProcureCompetitivelyReason,
      CompanyName: item.CompanyName,
      ContactName: item.ContactName,
      TendererContactEmail: item.TendererContactEmail,
      TendererContactTelephone: item.TendererContactTelephone,
      DpoOptions: DpoOptionsId,
      IsmOptions: IsmOptionsId,
      SubProjectCode: item.SubProjectCode,
      SubProjectCodeDescription: item.SubProjectCodeDescription,
      ProjectCode: item.ProjectCode,
      CostCentre: item.CostCentre,
      DirectorateOfPurchase: item.DirectorateOfPurchase,
      PrimaryBusinessManager: PrimaryBusinessManager,
      SecondaryBusinessManager: SecondaryBusinessManager,
      PrimaryFinancialApprover: PrimaryFinancialApprover,
      SecondaryFinancialApprover: SecondaryFinancialApprover,
      StaCompliant: item.StaCompliant,
      CommercialManager: CommercialManager,
      ContractNumber: item.ContractNumber,
      ProcurementStrategy: item.ProcurementStrategy,
      ProcurementType: ProcurementTypeId,
      ProcurementNotes: item.ProcurementNotes,
      ReasonForRejection: item.ReasonForRejection,
      SuggestedTendererStatus: SuggestedTendererStatusId,
      Variation: item.Variation,
      VariationCheck: item.VariationCheck,
      RedraftingNotPossible: item.RedraftingNotPossible,
    };
  }

  /**
   * @description generates the PROC reference upon the form being submitted by the Purchaser 
   * @fires from this.submit()
   * @param formId the id of the SharePoint list item storing the form data
   * @param formType i.e. STA, RTP. FTDP
   */
  private static generateProcReference(formId: number, formData: IFormData): string {
    if (formData['Variation']) {
      const procRefAndId = formData['Variation'];
      //  * split on comma to remove id 
      const procRef = procRefAndId.split(',')[0];
      const procRefArr: string[] = procRef.split('-', 4);
      if (procRefArr.length == 3) {
        // * this is the first variation
        return `${procRef}-1`;
      } else if (procRefArr.length == 4) {
        // * this is not the first variation
        let variationNumStr: string = procRefArr[3];
        // get the length of the variatonNum
        let numLength: number = variationNumStr.length;
        let variationNum: number = parseInt(procRefArr[3]);
        // remove the variationNum from procRef and replace it with nextVariationNum
        let baseProcRef: string = procRef.slice(0, -numLength);
        // ! assumes that the user has selected the most recent variation 
        const nextVariationNum: number = variationNum += 1;
        const nextVariationNumStr: string = nextVariationNum.toString();
        // add nextVariationNum to baseProcRef
        return `${baseProcRef}${nextVariationNumStr}`;
      }

    } else {
      const date: Date = new Date();
      const year: number = date.getFullYear();
      const procRef = 'PROC-' + formId + '-' + year;
      return procRef;
    }
  }

  /**
   * @description loops through array of choice fields gets the selectable options for each 
   * @returns an object, a property per choice field, containing the selectable options for that field
   */
  private static async overseeGetDropdownOptions(): Promise<any> {
    let dropdownsPm: Promise<IDropdownFieldsOptions> = this.getOptionsForDropDownFields(this.choiceFields);
    let subProjectCodesPm: Promise<ISubProjectCodeOptions> = this.getOptionsForComboFields(this.comboFields);
    let variationsPm: Promise<IVariationOptions> = Utility.getPROCReferences();
    let comboFieldOptions = await Promise.all([subProjectCodesPm, variationsPm]);
    let allComboFieldOptions: IComboFieldsOptions = { ...comboFieldOptions[0], ...comboFieldOptions[1] };
    return Promise.all([dropdownsPm, allComboFieldOptions]);
  }

  /**
   * @description makes a call for each dropdown field to retrieve lookup values
   * @param fields the dropdown fields on the form
   * @dev returns mock data 
   */
  private static async getOptionsForDropDownFields(fields: string[]): Promise<any> {
    if (Environment.type === EnvironmentType.Local) {
      return MockDropdownFieldsOptions;
    } else if (Environment.type === EnvironmentType.SharePoint) {
      let dropdownFieldOptions = {};
      for (let i = 0; i < fields.length; i++) {
        let options: IDropDownOption[] = await this.getOptionsForDropDownField(fields[i]);
        dropdownFieldOptions[fields[i]] = options;
        if (i == (fields.length + - 1)) {
          return dropdownFieldOptions;
        }
      }
    }
  }

  /**
   * @description makes a call for each combo field to retrieve lookup values
   * @param fields the combo fields on the form
   * @dev returns mock data
   */
  private static async getOptionsForComboFields(fields: string[]): Promise<any> {
    if (Environment.type === EnvironmentType.Local) {
      return MockComboFieldsOptions;
    } else if (Environment.type === EnvironmentType.SharePoint) {
      let comboFieldOptions = {};
      for (let i = 0; i < fields.length; i++) {
        let options: { key: string | number, text: string }[] = await this.getOptionsForAComboField(fields[i]);
        comboFieldOptions[fields[i]] = options;
        if (i == (fields.length + - 1)) {
          return comboFieldOptions;
        }
      }
    }
  }

  /**
   * @description gets dropdown options 
   * @param field the dropdown field to retrieve the options for
   * @returns an array of dropdownOptions for the specified field
   */
  private static getOptionsForDropDownField(field: string): PromiseLike<IDropDownOption[]> {

    let options: IDropDownOption[] = new Array;
    let option: IDropDownOption;
    let list: string;
    let name: string = "Title";
    let id: string = "Id";

    switch (field) {
      case 'SecondaryBusinessManager':
        list = 'AllBusinessManagers';
        id = "Title"
        break;
      case 'ProcurementRoute':
        list = 'procurement route';
        break;
      case 'ProcurementType':
        list = 'procurement type';
        break;
      case 'ApprovedScope':
        list = 'approved scope';
        break;
      case 'RouteToMarket':
        list = 'route to market';
        break;
      case 'StaType':
        list = 'sta type';
        break;
      case 'DpoOptions':
        list = 'dpo options';
        break;
      case 'IsmOptions':
        list = 'ism options';
        break;
      case 'BusinessCaseRequirement':
        list = 'business case requirement';
        break;
      case 'SuggestedTendererStatus':
        list = 'suggested tenderer status';
        break;
      case 'BusinessManagerDirectorates':
        list = 'BusinessManagerDirectorates';
        break;
      case 'UnableProcureCompetitively':
        list = 'unable procure competitively';
        name = "Option";
        id = "Option";
        break;
    }

    try {
      return sp.web.lists
        .getByTitle(list)
        .items
        .select(name, "ID")
        .filter("Archive eq 'No'")
        .get()
        .then((response: any[]) => {
          if (field == "SecondaryBusinessManager") {
            response.map(item => {
              option = {
                key: item[id].toLowerCase(),
                text: item[name].toLowerCase()
              };
              options.push(option);
            });
          } else {
            response.map(item => {
              option = {
                key: item[id],
                text: item[name]
              };
              options.push(option);
            });
          }
          return options;
        });
    } catch (error) {
      console.log('an error occured while retrieving options for the dropdown field: ' + field + ' error: ' + error);
    }
  }

  /**
   * @description gets choice options 
   * @param field the choice field to retrieve the options for
   * @returns an array of choiceGroupOptions for the specified field
   */
  private static async getOptionsForAComboField(field): Promise<IDropDownOption[]> {

    let options: IDropDownOption[] = new Array;
    let option: IDropDownOption;
    let list: string;

    switch (field) {
      case 'SubProjectCode':
        list = 'CurrentSubProjects';
        break;
    }

    try {
      let allItems: PagedItemCollection<any[]>[] = [];
      let items = await sp.web.lists
        .getByTitle(list)
        .items
        .select("Title", "Description")
        .getPaged<any>();

      if (items.results.length > 0) {
        allItems.push(...items.results);
        while (items.hasNext) {
          items = await items.getNext();
          allItems.push(...items.results);
          // console.log("total items retrieved for field: " + field + " so far:" + allItems.length);
        }
        allItems.map(item => {
          option = {
            key: item['Title'],
            text: item['Title'] + ' : ' + item['Description']
          };
          options.push(option);
        });
        return options;

      }

    } catch (error) {
      console.log('an error occured while retrieving options for the combo field: ' + field + ' error: ' + error);
    }

  }

  /**
   * @description checks if the user is a member of the SharePoint Group 'Commercial'
   * @dev returns true
   */
  private static async checkIfUserIsMemberOfCommercial(): Promise<boolean> {
    if (Environment.type === EnvironmentType.Local) {
      return true;
    } else if (Environment.type === EnvironmentType.SharePoint) {
      const userGroups: any[] = await Utility.getUserGroupMembership();
      for (let i = 0; i < userGroups.length; i++) {
        if (userGroups[i].Title == "Commercial") {
          return true;
        }
        return false;
      }
    }
  }

  private static getTagForTagPickerFieldFromListItem = (item: any, field: string): ITag[] => {
    return (item[field]) ? [
      {
        key: item[field].ID,
        name: item[field].Title
      }
    ] : null;
  }

  private static getIdForDropdownFieldFromListItem(item: any, field: string): number {
    return (item[field]) ? item[field].ID : null;
  }

  private static getSelectedTagFromFormData = (formData: any, field: string): number => {
    if (formData[field] && formData[field][0]) {
      return formData[field][0].key;
    }
    else return null;
  }

  /**
   * @description checks if the user can add new list items and add new documents 
   * @fires Form.load()
   * @param listName - where the formData is stored
   * @dev returns true
   * @returns true if the user is an editor, false if not 
   * @deprecated everyone should be able to add and manage permissions to do the procurement list and the documents library
   * TODO causes error in IE - user not being passed into querym try wait 100ms before proceeding
   */
  /* private static async checkUserRoleIsEditorForNewForm(listName: string):Promise<boolean> {
    // local environment
    if (Environment.type === EnvironmentType.Local) {
      // user is editor 
      return true;
    } else if (Environment.type === EnvironmentType.SharePoint) {
      const list:SharePointQueryableSecurable = sp.web.lists.getByTitle(listName);
      const docLib:SharePointQueryableSecurable = sp.web.lists.getByTitle('Documents');

      try {
        // * use current user id to check permissions instead of getCurrentUserEffectivePermissions()
        const currentUserLogin:string = await sp.web.currentUser.select('LoginName').get();
        console.log('currentUserLogin: ' + currentUserLogin['LoginName']);
        // * I've taken these calls outside the try catch block to see if error is uncaught in promise now
        // const listPerms:Promise<BasePermissions> = list.getCurrentUserEffectivePermissions();
        const listPerms:BasePermissions = await list.getUserEffectivePermissions(currentUserLogin['LoginName']);
        console.log('listPerms: ' + JSON.stringify(listPerms));
        // const docLibPerms:Promise<BasePermissions> = docLib.getCurrentUserEffectivePermissions();
        const docLibPerms:Promise<BasePermissions> = docLib.getUserEffectivePermissions(currentUserLogin['LoginName']);
        
        let perms = await Promise.all([listPerms, docLibPerms])
        .then(response => {

          return {
            userListPerms : response[0],
            userDocLibPerms : response[1]
          };
        });
        
        const userCanAddListItems:boolean = list.hasPermissions(perms.userListPerms, PermissionKind.AddListItems);
        // ! This means the permission level must include enumerate permissions AND manage permissions
        const userCanAddManageListItemmsPermissions:boolean = list.hasPermissions(perms.userListPerms, PermissionKind.ManagePermissions);
        const userCanAddDocuments:boolean = docLib.hasPermissions(perms.userListPerms, PermissionKind.AddListItems);
        // ! This means the permission level must include enumerate permissions AND manage permissions
        const userCanManageDocumentsPermissions:boolean = docLib.hasPermissions(perms.userListPerms, PermissionKind.ManagePermissions);
    
        let userRole:{addListItems:boolean, permissionListItems:boolean, addDocuments:boolean, permissionDocuments:boolean} = await Promise.all([userCanAddListItems, userCanAddManageListItemmsPermissions, userCanAddDocuments, userCanManageDocumentsPermissions])
        .then(response=> {
          return {
            addListItems : response[0],
            permissionListItems : response[1],
            addDocuments: response[2],
            permissionDocuments: response[3],
          };
        });
    
        if (userRole.addListItems && userRole.permissionListItems && userRole.addDocuments && userRole.permissionDocuments) {
          return true;
        } else {
          return false;
        }
        
      } catch (error) {
        alert('an error occured while checking user permissions to the list and document library, error: ' + error);
      }
    }
    

  } */

  /**
   * @description checks if the user role for the list item when form mode = edit
   * @fires Form.load()
   * @param listName - that stores the formData
   * @param formId - the id of the list item storing the formData
   * @returns the userRole, "editor", "viewer", "none"
   */
  private static checkUserRoleListItemForEditForm(listName: string, formId: number): Promise<string> {
    const listItem: SharePointQueryableSecurable = sp.web.lists.getByTitle(listName).items.getById(formId);
    return listItem.getCurrentUserEffectivePermissions()
      .then(perms => {
        if (listItem.hasPermissions(perms, PermissionKind.EditListItems) && listItem.hasPermissions(perms, PermissionKind.ManagePermissions)) {
          return "editor";
        } else if (listItem.hasPermissions(perms, PermissionKind.ViewListItems)) {
          return "viewer";
        } else {
          return "none";
        }
      });
  }

  /**
   * @description checks whether the user can add or view documents in the parent folder
   * @fires Form.load()
   * @param folderName - this is either the form guid if the form is unsubmitted, or the proc ref if the form is submitted
   * @param serverRelativeUrl 
   * @returns - the user role "editor", "viewer", "none"
   */
  private static async checkUserRoleDocumentsForEditForm(folderName: string, serverRelativeUrl: string): Promise<string> {
    try {
      const parentFolder = await sp.web.getFolderByServerRelativeUrl(serverRelativeUrl + '/Shared Documents/' + folderName);
      const parentFolderBasePerms = await parentFolder.select('*').expand('ListItemAllFields/EffectiveBasePermissions').get();
      const perms = parentFolderBasePerms.ListItemAllFields.EffectiveBasePermissions;
      if (sp.web.hasPermissions(perms, PermissionKind.AddListItems) && sp.web.hasPermissions(perms, PermissionKind.ManagePermissions)) {
        return "editor";
      } else if (sp.web.hasPermissions(perms, PermissionKind.ViewListItems)) {
        return "viewer";
      } else {
        return "none";
      }
      /* return parentFolder.select('*').expand('ListItemAllFields/EffectiveBasePermissions').get()
      .then(folder => {
        const perms = folder.ListItemAllFields.EffectiveBasePermissions;
        if (sp.web.hasPermissions(perms, PermissionKind.AddListItems) && sp.web.hasPermissions(perms, PermissionKind.ManagePermissions)) {
          return "editor";
        } else if (sp.web.hasPermissions(perms, PermissionKind.ViewListItems)) {
          return "viewer";
        } else {
          return "none";
        }
      }); */

    } catch (error) {
      console.log('an error occured while checking user permissions to the procurement request folder, error: ' + error);
    }
  }


}

