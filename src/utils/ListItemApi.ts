import { sp, ItemAddResult } from '@pnp/sp';
import { Utility } from '.';
import { IListItem } from '../interfaces/IListItem';
import { IFormData } from '../interfaces';
import { ITaskAssigneeQuery } from '../mock/MockStaTaskAssignees';
import { Environment, EnvironmentType } from '@microsoft/sp-core-library';


/**
 * @description details methods for interacting with the SharePoint list storing the form data
 */
export class ListItemApi {
  /**
   * @description adds a new SharePoint list item based on the formData
   * @param formData 
   * @param listName 
   * @param formGuid 
   * @returns formId generated from the SharePoint list where the formData is saved
  */
  public static async add(serverRelativeUrl: string, listName: string, formGuid: string, listItem: any): Promise<number> {

    if (Environment.type === EnvironmentType.Local) {
      return 10;
    } else if (Environment.type === EnvironmentType.SharePoint) {
      try {
        // convert formData to listItem 
        // let listItem = await Utility.convertFormData(formData);
        let itemId: number;
        // create new list item 
        let item = await sp.web.lists
          .getByTitle(listName)
          .items
          .add(listItem)
          .then((iar: ItemAddResult) => {
            itemId = iar.data.ID;
            return iar;
          });

        // update list item with hyperlink based on list item id 
        await sp.web.lists
          .getByTitle(listName)
          .items
          .getById(item.data.ID)
          .update({
            Link: {
              "__metadata": { type: "SP.FieldUrlValue" },
              Description: item.data.ContractTitle,
              Url: serverRelativeUrl + `/SitePages/Procurement-Request-Form.aspx?form_id=${item.data.ID}&form_guid=${formGuid}`,
            },
          });
        return itemId;

      } catch (error) {
        alert('form failed to submit: ' + error);
      }
    }
  }

  /**
   * 
   * @param formData 
   * @param listName 
   * @param formId 
   */
  public static async update(listName: string, formId: number, listItem: any): Promise<void> {
    if (Environment.type === EnvironmentType.SharePoint) {
      try {
        return sp.web.lists
          .getByTitle(listName)
          .items
          .getById(formId)
          .update(listItem)
          .then(_ => {
          });
      } catch (error) {
        alert('form failed to update list item, error: ' + error);
      }
    }
  }

  /**
   * 
   * @param listName 
   * @param formId 
   */
  public static delete(listName: string, formId: number): Promise<string> {
    if (Environment.type === EnvironmentType.SharePoint) {
      try {
        return sp.web.lists
          .getByTitle(listName)
          .items
          .getById(formId)
          .recycle();
      } catch (error) {
        alert('form failed to delete:' + error);
      }
    }
  }

  /**
   * 
   * @param listName 
   * @param formId 
   * @param absoluteUrl 
   */
  public static async get(listName: string, formId: number, listColumns: string[], expandListColumns: string[]): Promise<any> {
    if (Environment.type === EnvironmentType.SharePoint) {
      console.log(`3a - get called, fetching list item`);
      console.log(`list columns ${JSON.stringify(listColumns)}`);
      console.log(`expand list columns ${JSON.stringify(expandListColumns)}`);
      try {
        return await sp.web.lists
          .getByTitle(listName)
          .items
          .getById(formId)
          .select(...listColumns)
          .expand(...expandListColumns)
          .get();
      } catch (error) {
        console.log('an error occured while trying to get a SharePoint list item: ' + error);
      }
    }
  }

  /* public static getVariation(listName:string, listColumns:string[], expandListColumns:string[]):Promise<any[]> {
    if (Environment.type === EnvironmentType.SharePoint) {
      try {
        return sp.web.lists
        .getByTitle(listName)
        .items
        .select(...listColumns)
        .expand(...expandListColumns)
        .getAll()
        .then(data => {
          return data[0];
        });
      } catch (error) {
        alert('an error occured while trying to get a SharePoint list item: ' + error);
      }
    }
  } */

  /**
   * @description when a new form is submitted this updates the list item with the newly created formId
   * @fires from Submit.submitForm()
   * @param listName 
   * @param formId 
   * @param procRef 
   */
  public static async updateAsSubmitted(listName: string, formId: number, procRef: string): Promise<void> {
    if (Environment.type === EnvironmentType.SharePoint) {
      try {
        await sp.web.lists
          .getByTitle(listName)
          .items
          .getById(formId)
          .update({
            ProcReference: procRef,
            ProcurementStatus: "Submitted"
          });
      } catch (error) {
        alert('form failed to update proc reference, error: ' + error);
      }
    }
  }

  /**
   * @description sets permissions on the list item for the purchaser and Commercial group to edit the list item
   * @fires from Submit.submitForm()
   * @param listName
   * @param formId 
   */
  public static async setPermissionsOnSave(listName: string, formId: number, formData: IFormData): Promise<void> {
    if (Environment.type === EnvironmentType.SharePoint) {
      try {
        const item = sp.web.lists
          .getByTitle(listName)
          .items
          .getById(formId);
        await item.breakRoleInheritance(false);

        // Get user/group proncipal Id
        // const { Id: currentUserId } = await sp.web.currentUser.select('Id').get();
        const purchaserId: number = await Utility.getSharePointUserId(formData['PurchaserName'][0]['optionalText']);

        const { Id: commercialGroupId } = await sp.web.siteGroups.getByName('Commercial').select('Id').get();
        // Get role definition Id
        const { Id: contributeWithManagePermissionsId } = await sp.web.roleDefinitions.getByName('Contribute with Manage Permissions').get();
        // Assigning permissions
        await Promise.all([item.roleAssignments.add(purchaserId, contributeWithManagePermissionsId), item.roleAssignments.add(commercialGroupId, contributeWithManagePermissionsId)]);
      } catch (error) {
        alert('an error occured while setting permissions on the list item: ' + error);
      }
    }
  }

  /**
   * @description updates the permissions on the list item to allow the business managers and financial approvers to read the list item
   * @fires from Submit.submitForm()
   * @param listName
   * @param formId
   * @param formData
   */
  public static async updatePermissionsOnSubmit(listName: string, formId: number, formData: IFormData): Promise<any> {
    if (Environment.type === EnvironmentType.SharePoint) {
      try {
        const item = sp.web.lists
          .getByTitle(listName)
          .items
          .getById(formId);
        await item.breakRoleInheritance(false);

        // get logon names
        const primaryBusinessManagerLogonName: string = "i:0#.f|membership|" + formData['PrimaryBusinessManager'];
        const secondaryBusinessManagerLogonName: string = "i:0#.f|membership|" + formData['SecondaryBusinessManager'];
        const primaryFinancialApproverLogonName: string = "i:0#.f|membership|" + formData['PrimaryFinancialApprover'];
        const secondaryFinancialApproverLogonName: string = "i:0#.f|membership|" + formData['SecondaryFinancialApprover'];

        // Get user Ids
        const primaryBusinessManagerUserId: number = await Utility.getSharePointUserId(primaryBusinessManagerLogonName);
        const secondaryBusinessManagerUserId: number = await Utility.getSharePointUserId(secondaryBusinessManagerLogonName);
        const primaryFinancialApproverUserId: number = await Utility.getSharePointUserId(primaryFinancialApproverLogonName);
        const secondaryFinancialApproverUserId: number = await Utility.getSharePointUserId(secondaryFinancialApproverLogonName);

        // Get role definition Id
        const { Id: readId } = await sp.web.roleDefinitions.getByName('Read with Enumerate Permissions').get();
        // Assigning permissions
        const primaryBMPermUpdatePm = item.roleAssignments.add(primaryBusinessManagerUserId, readId);
        const secondaryBMPermUpdatePm = item.roleAssignments.add(secondaryBusinessManagerUserId, readId);
        const primaryFAPermUpdatePm = item.roleAssignments.add(primaryFinancialApproverUserId, readId);
        const secondaryFAPermUpdatePm = item.roleAssignments.add(secondaryFinancialApproverUserId, readId);

        if (formData['ProcurementRoute'] == 3) {

          const staTaskAssignees: ITaskAssigneeQuery[] = await Utility.getTaskAssigneesUserIds();

          const ceoUserId: number = staTaskAssignees[0]['TaskAssignee']['ID'];
          const corpHeadUserId: number = staTaskAssignees[1]['TaskAssignee']['ID'];
          const commHeadUserId: number = staTaskAssignees[2]['TaskAssignee']['ID'];
          const finDirUserId: number = staTaskAssignees[3]['TaskAssignee']['ID'];
          const snrDirPeopleUserId: number = staTaskAssignees[4]['TaskAssignee']['ID'];

          const ceoPermUpdatePm = item.roleAssignments.add(ceoUserId, readId);
          const corpHeadPermUpdatePm = item.roleAssignments.add(corpHeadUserId, readId);
          const commHeadPermUpdatePm = item.roleAssignments.add(commHeadUserId, readId);
          const finDirPermUpdatePm = item.roleAssignments.add(finDirUserId, readId);
          const snrDirPeoplePermUpdatePm = item.roleAssignments.add(snrDirPeopleUserId, readId);

          return await Promise.all([primaryBMPermUpdatePm, secondaryBMPermUpdatePm, primaryFAPermUpdatePm, secondaryFAPermUpdatePm, ceoPermUpdatePm, corpHeadPermUpdatePm, commHeadPermUpdatePm, finDirPermUpdatePm, snrDirPeoplePermUpdatePm]);

        } else {
          return await Promise.all([primaryBMPermUpdatePm, secondaryBMPermUpdatePm, primaryFAPermUpdatePm, secondaryFAPermUpdatePm]);
        }


      } catch (error) {
        alert('an error occured while updating permissions on the list item for the bus mgrs and financial approvers: ' + error);
      }
    }
  }

}