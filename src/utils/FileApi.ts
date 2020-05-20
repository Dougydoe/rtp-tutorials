import { sp } from "@pnp/sp";
import { IFormData } from "../interfaces";
import { Utility } from "./Utility";
import { ITaskAssigneeQuery } from "../mock/MockStaTaskAssignees";
import { Environment, EnvironmentType } from "@microsoft/sp-core-library";

/**
 * @description details methods for interacting with the SharePoint document library
 */
export class FileApi {
  private static subFolders = [
    "Approvals",
    "Attachments",
    "Award Evaluation",
    "Clarification of Proposals",
    "Invitation to tender",
    "Notifications and De-briefing",
    "Planning",
    "Pre-Qualifying Stage",
    "Selection Evaluation",
    "Supplier Clarification",
    "Supplier Proposal"
  ];

  /**
   * @fires on the submit button being clicked when the form mode equals "new", fires after SharePointListApi.addItem() executes
   * @description creates folder structure, adds attachments, sets permissions for all users, sets proc ref and renames folder
   * @param serverRelativeUrl
   * @param formData
   * @param formId collected from SharePointListApi.addItem()
   */
  public static async handleAttachmentsOnSubmitForNewForm(serverRelativeUrl: string, formData: any, procRef: string): Promise<void> {

    const parentFolder: {
      Exists: boolean;
      ServerRelativeUrl: string;
    } = await this.provisionParentFolder(procRef);

    if (parentFolder["Exists"]) {
      // parent folder provisioned
      // set permissions on parent folder
      await this.setParentFolderPermissionsOnSubmit(
        formData,
        parentFolder["ServerRelativeUrl"]
      );
      // provision sub folders
      const subFolders = await this.provisionSubFolders(procRef);
      if (subFolders[0]["data"]["Exists"] && formData["Attachments"]) {
        // sub folders provisioned and there are attachments
        await this.addAttachmentsForNewSubmittedForm(
          serverRelativeUrl,
          procRef,
          formData["Attachments"]
        );
      }
    }

  }

  /**
   * @description checks for Attachments folder, adds attachments, updates permissions for bus mgrs and FAs, and renames folder
   * @param formGuid
   * @param serverRelativeUrl
   * @param formData
   * @param procRef
   */
  public static async handleAttachmentsOnSubmitForEditForm(
    formGuid: string,
    serverRelativeUrl: string,
    formData: any,
    procRef: string
  ): Promise<void> {
    if (Environment.type === EnvironmentType.SharePoint) {
      const folderExists: {
        Exists: boolean;
        ServerRelativeUrl: string;
      } = await this.checkAttachmentsFolderExists(
        formGuid,
        serverRelativeUrl,
        formData
      );
      if (folderExists["Exists"]) {
        // the attachments folder exists
        if (formData["Attachments"]) {
          await this.addAttachments(
            serverRelativeUrl,
            formGuid,
            formData["Attachments"],
            formData
          );
        }
      } else {
        // the attachments folder does not exist
        throw new Error(
          "There is no folder for these attachment(s) to be saved to."
        );
      }
      // update parent folder permissions for submit
      await this.updateParentFolderPermissionsOnSubmit(
        formData,
        folderExists["ServerRelativeUrl"]
      );
      // rename parent folder after Proc Reference
      await this.renameParentFolder(formGuid, serverRelativeUrl, procRef);
    }
  }

  /**
   * @description creates folder structure, sets permissions for purchaser & Commercial, adds attachments
   * @param formGuid
   * @param serverRelativeUrl
   * @param formData
   */
  public static async handleAttachmentsOnSaveForNewForm(
    formGuid: string,
    serverRelativeUrl: string,
    formData: IFormData
  ): Promise<any[]> {
    const parentFolder: {
      Exists: boolean;
      ServerRelativeUrl: string;
    } = await this.provisionParentFolder(formGuid);
    if (parentFolder["Exists"]) {
      // parent folder was provisioned
      // set permissions on parent folder for save
      await this.setParentFolderPermissionsOnSave(
        parentFolder["ServerRelativeUrl"],
        formData
      );
      // provision sub folders
      const subFolders = await this.provisionSubFolders(formGuid);
      if (subFolders[0]["data"]["Exists"]) {
        // sub folders were provisioned
        if (formData["Attachments"]) {
          // the form has attachments, add the attachments
          return await this.addAttachments(
            serverRelativeUrl,
            formGuid,
            formData["Attachments"],
            formData
          );
        }
      } else {
        // the sub folders were not provisioned
        throw new Error(
          "sub folders failed to be provisioned, any attachments were not added"
        );
      }
    } else {
      throw new Error(
        "the parent folder failed to be provisioned, any attachments were not added"
      );
    }
  }

  /**
   * @description checks Attachments folder exists, adds attachments
   * @param formGuid
   * @param serverRelativeUrl
   * @param formData
   * ? Should I throw an error, or should I alert()
   */
  public static async handleAttachmentsOnSaveForEditForm(
    formGuid: string,
    serverRelativeUrl: string,
    formData: any
  ): Promise<any[]> {
    if (formData["Attachments"]) {
      // there are attachments
      const folderExists: {
        Exists: boolean;
        ServerRelativeUrl: string;
      } = await this.checkAttachmentsFolderExists(
        formGuid,
        serverRelativeUrl,
        formData
      );
      if (folderExists["Exists"]) {
        // attachments folder exists
        return await this.addAttachments(
          serverRelativeUrl,
          formGuid,
          formData["Attachments"],
          formData
        );
      } else {
        // attachments folder does not exist
        throw new Error(
          "There is no attachments folder for these attachment(s) to be saved to."
        );
      }
    }
  }

  /**
   * @description renames the parent folder from the formGuid to the Proc Reference
   * @fires from this.handleAttachmentsOnSubmitForNewForm() or this.handleAttachmentsOnSubmitForEditForm()
   * @param formGuid
   * @param serverRelativeUrl
   * @param procRef
   */
  private static renameParentFolder(
    formGuid: string,
    serverRelativeUrl: string,
    procRef: string
  ): Promise<void> {

    console.log('renaming parent folder with proc:', procRef);
    return sp.web
      .getFolderByServerRelativeUrl(
        serverRelativeUrl + "/Shared Documents/" + formGuid
      )
      .getItem()
      .then(item => {
        item.update({
          FileLeafRef: procRef
        });
      });
  }

  /**
   * @description checks if the Attachments folder in the parent folder exists
   * @param formGuid
   * @param serverRelativeUrl
   */
  private static checkAttachmentsFolderExists(
    formGuid: string,
    serverRelativeUrl: string,
    formData: any
  ): Promise<{ Exists: boolean; ServerRelativeUrl: string }> {
    let folderUrl: string;
    if (
      formData["ProcurementStatus"] &&
      formData["ProcurementStatus"] == "Draft"
    ) {
      folderUrl =
        serverRelativeUrl + "/Shared Documents/" + formGuid + "/Attachments";
    } else {
      folderUrl =
        serverRelativeUrl +
        "/Shared Documents/" +
        formData["ProcReference"] +
        "/Attachments";
    }
    try {
      return sp.web
        .getFolderByServerRelativeUrl(folderUrl)
        .get()
        .then(response => {
          return {
            Exists: response.Exists,
            ServerRelativeUrl: response.ServerRelativeUrl
          };
        });
    } catch (error) {
      console.log(
        "an error occured while checking if the parent folder exists: " + error
      );
    }
  }

  /**
   * @fires when handleAttachmentsOnSave() is called
   * @description creates a folder for the procurement request form
   * @returns true or false depending on whether the folder was successfully created
   */
  private static provisionParentFolder(
    name: string
  ): Promise<{ Exists: boolean; ServerRelativeUrl: string }> {
    try {
      return sp.web.folders
        .getByName("Shared Documents")
        .folders.add(name)
        .then(response => {
          return {
            Exists: response.data.Exists,
            ServerRelativeUrl: response.data.ServerRelativeUrl
          };
        });
    } catch (error) {
      console.log(
        "an error occured while provisioning folder for procurement request: " +
        error
      );
    }
  }

  /**
   * @description removes the parent folder for an unsubmitted form
   * @fires from FormButtons.discardForm()
   * @param serverRelativeUrl
   * @param formGuid
   * * when ProcurementStatus != Draft the Discard button is hidden, so the Proc Ref is not needed to identify the procurement request folder
   */
  public static removeParentFolder(
    serverRelativeUrl: string,
    formGuid: string,
    formData: any
  ): Promise<string> {
    let folderUrl: string;
    if (
      formData["ProcurementStatus"] &&
      formData["ProcurementStatus"] == "Draft"
    ) {
      // form is in draft mode, use form guid
      folderUrl = serverRelativeUrl + "/Shared Documents/" + formGuid;
    } else {
      // form is no longer in draft, use Proc Ref
      folderUrl =
        serverRelativeUrl + "/Shared Documents/" + formData["ProcReference"];
    }
    try {
      return sp.web.getFolderByServerRelativeUrl(folderUrl).recycle();
    } catch (error) {
      console.log(
        "an error occured while removing the attachments folder, error: " +
        error
      );
    }
  }

  /**
   * @description sets permissions for the purchaser (current user) and Commercial group to Edit the parent folder
   * @fires after provisionProcurementRequestFolder() is sucessfully provisioned
   * @param serverRelativeUrl passed in from response body of provisionProcurementRequestFolder()
   */
  private static async setParentFolderPermissionsOnSave(
    serverRelativeUrl: string,
    formData: IFormData
  ): Promise<any> {
    try {
      const parentFolder = sp.web.getFolderByServerRelativeUrl(
        serverRelativeUrl
      );
      const parentFolderItem = await parentFolder.getItem();
      await parentFolderItem.breakRoleInheritance(false);
      // Get user/group proncipal Id
      // const { Id: currentUserId } = await sp.web.currentUser.select('Id').get();
      const purchaserId: number = await Utility.getSharePointUserId(
        formData["PurchaserName"][0]["optionalText"]
      );

      const { Id: commercialGroupId } = await sp.web.siteGroups
        .getByName("Commercial")
        .select("Id")
        .get();
      // Get role definition Id
      const {
        Id: contributeWithManagePermissionsId
      } = await sp.web.roleDefinitions
        .getByName("Contribute with Manage Permissions")
        .get();
      // Assigning permissions
      const purchaserPermPm = parentFolderItem.roleAssignments.add(
        purchaserId,
        contributeWithManagePermissionsId
      );
      const commercialPermPm = parentFolderItem.roleAssignments.add(
        commercialGroupId,
        contributeWithManagePermissionsId
      );
      return await Promise.all([purchaserPermPm, commercialPermPm]);
    } catch (error) {
      console.log(
        "an error occured while setting permissions on the attachments folder: " +
        error
      );
    }
  }

  /**
   * @fires from this.handleAttachmentsOnSubmitForNewForm()
   * @param formData
   * @param serverRelativeUrl
   */
  private static async setParentFolderPermissionsOnSubmit(
    formData: IFormData,
    serverRelativeUrl: string
  ): Promise<any> {
    try {
      const parentFolder = sp.web.getFolderByServerRelativeUrl(
        serverRelativeUrl
      );
      const parentFolderItem = await parentFolder.getItem();
      await parentFolderItem.breakRoleInheritance(false);
      // get logon names
      const primaryBusinessManagerLogonName: string =
        "i:0#.f|membership|" + formData["PrimaryBusinessManager"];
      const secondaryBusinessManagerLogonName: string =
        "i:0#.f|membership|" + formData["SecondaryBusinessManager"];
      const primaryFinancialApproverLogonName: string =
        "i:0#.f|membership|" + formData["PrimaryFinancialApprover"];
      const secondaryFinancialApproverLogonName: string =
        "i:0#.f|membership|" + formData["SecondaryFinancialApprover"];
      // Get user proncipal Id
      // const { Id: currentUserId } = await sp.web.currentUser.select('Id').get();
      const purchaserId: number = await Utility.getSharePointUserId(
        formData["PurchaserName"][0]["optionalText"]
      );

      const primaryBusinessManagerUserId = await Utility.getSharePointUserId(
        primaryBusinessManagerLogonName
      );
      const secondaryBusinessManagerUserId = await Utility.getSharePointUserId(
        secondaryBusinessManagerLogonName
      );
      const primaryFinancialApproverUserId = await Utility.getSharePointUserId(
        primaryFinancialApproverLogonName
      );
      const secondaryFinancialApproverUserId = await Utility.getSharePointUserId(
        secondaryFinancialApproverLogonName
      );
      // get group principal Id
      const { Id: commercialGroupId } = await sp.web.siteGroups
        .getByName("Commercial")
        .select("Id")
        .get();
      // Get role definition Ids
      const {
        Id: contributeWithManagePermissionsId
      } = await sp.web.roleDefinitions
        .getByName("Contribute with Manage Permissions")
        .get();
      const { Id: readId } = await sp.web.roleDefinitions
        .getByName("Read with Enumerate Permissions")
        .get();
      // Assigning permissions
      const purchaserPermsPm = parentFolderItem.roleAssignments.add(
        purchaserId,
        contributeWithManagePermissionsId
      );
      const commPermsPm = parentFolderItem.roleAssignments.add(
        commercialGroupId,
        contributeWithManagePermissionsId
      );
      const primBMPm = parentFolderItem.roleAssignments.add(
        primaryBusinessManagerUserId,
        readId
      );
      const secBMPm = parentFolderItem.roleAssignments.add(
        secondaryBusinessManagerUserId,
        readId
      );
      const primFAPm = parentFolderItem.roleAssignments.add(
        primaryFinancialApproverUserId,
        readId
      );
      const secFAPm = parentFolderItem.roleAssignments.add(
        secondaryFinancialApproverUserId,
        readId
      );

      if (formData["ProcurementRoute"] == 3) {
        const staTaskAssignees: ITaskAssigneeQuery[] = await Utility.getTaskAssigneesUserIds();
        console.log(staTaskAssignees);
        const ceoUserId: number = staTaskAssignees[0]["TaskAssignee"]["ID"];
        const corpHeadUserId: number =
          staTaskAssignees[1]["TaskAssignee"]["ID"];
        const commHeadUserId: number =
          staTaskAssignees[2]["TaskAssignee"]["ID"];
        const finDirUserId: number = staTaskAssignees[3]["TaskAssignee"]["ID"];
        const snrDirOfPeopleId: number = staTaskAssignees[4].TaskAssignee.ID;

        const ceoPermUpdatePm = parentFolderItem.roleAssignments.add(
          ceoUserId,
          readId
        );
        const corpHeadPermUpdatePm = parentFolderItem.roleAssignments.add(
          corpHeadUserId,
          readId
        );
        const commHeadPermUpdatePm = parentFolderItem.roleAssignments.add(
          commHeadUserId,
          readId
        );
        const finDirPermUpdatePm = parentFolderItem.roleAssignments.add(
          finDirUserId,
          readId
        );
        const snrDirPeoplePermUpdatePm = parentFolderItem.roleAssignments.add(
          snrDirOfPeopleId,
          readId
        );

        return await Promise.all([
          purchaserPermsPm,
          commPermsPm,
          primBMPm,
          secBMPm,
          primFAPm,
          secFAPm,
          ceoPermUpdatePm,
          corpHeadPermUpdatePm,
          commHeadPermUpdatePm,
          finDirPermUpdatePm,
          snrDirPeoplePermUpdatePm
        ]);
      } else {
        return await Promise.all([
          purchaserPermsPm,
          commPermsPm,
          primBMPm,
          secBMPm,
          primFAPm,
          secFAPm
        ]);
      }
    } catch (error) {
      console.log(
        "an error occured while setting permissions on the attachments folder for a submitted status: " +
        error
      );
    }
  }

  /**
   * @description gives read permission to bus mgrs and fin approvers to parent folder
   * @param formData
   * @param serverRelativeUrl
   */
  private static async updateParentFolderPermissionsOnSubmit(formData: any, serverRelativeUrl: string): Promise<any> {
    try {
      const parentFolderServerRelativeUrl = serverRelativeUrl.slice(0, -11);
      const parentFolder = sp.web.getFolderByServerRelativeUrl(parentFolderServerRelativeUrl);
      const parentFolderItem = await parentFolder.getItem();
      // get logon names
      const primaryBusinessManagerLogonName: string =
        "i:0#.f|membership|" + formData["PrimaryBusinessManager"];
      const secondaryBusinessManagerLogonName: string =
        "i:0#.f|membership|" + formData["SecondaryBusinessManager"];
      const primaryFinancialApproverLogonName: string =
        "i:0#.f|membership|" + formData["PrimaryFinancialApprover"];
      const secondaryFinancialApproverLogonName: string =
        "i:0#.f|membership|" + formData["SecondaryFinancialApprover"];

      // Get user Ids
      const primaryBusinessManagerUserId = await Utility.getSharePointUserId(
        primaryBusinessManagerLogonName
      );
      const secondaryBusinessManagerUserId = await Utility.getSharePointUserId(
        secondaryBusinessManagerLogonName
      );
      const primaryFinancialApproverUserId = await Utility.getSharePointUserId(
        primaryFinancialApproverLogonName
      );
      const secondaryFinancialApproverUserId = await Utility.getSharePointUserId(
        secondaryFinancialApproverLogonName
      );

      // Get role definition Id
      const { Id: readId } = await sp.web.roleDefinitions
        .getByName("Read with Enumerate Permissions")
        .get();
      // Assigning permissions
      const primBMPm = parentFolderItem.roleAssignments.add(
        primaryBusinessManagerUserId,
        readId
      );
      const secBMPm = parentFolderItem.roleAssignments.add(
        secondaryBusinessManagerUserId,
        readId
      );
      const primFAPm = parentFolderItem.roleAssignments.add(
        primaryFinancialApproverUserId,
        readId
      );
      const secFAPm = parentFolderItem.roleAssignments.add(
        secondaryFinancialApproverUserId,
        readId
      );

      if (formData["ProcurementRoute"] == 3) {
        const staTaskAssignees: ITaskAssigneeQuery[] = await Utility.getTaskAssigneesUserIds();
        const ceoUserId: number = staTaskAssignees[0]["TaskAssignee"]["ID"];
        const corpHeadUserId: number =
          staTaskAssignees[1]["TaskAssignee"]["ID"];
        const commHeadUserId: number =
          staTaskAssignees[2]["TaskAssignee"]["ID"];
        const finDirUserId: number = staTaskAssignees[3]["TaskAssignee"]["ID"];
        const snrDirOfPeopleId: number = staTaskAssignees[4].TaskAssignee.ID;

        const ceoPermUpdatePm = parentFolderItem.roleAssignments.add(
          ceoUserId,
          readId
        );
        const corpHeadPermUpdatePm = parentFolderItem.roleAssignments.add(
          corpHeadUserId,
          readId
        );
        const commHeadPermUpdatePm = parentFolderItem.roleAssignments.add(
          commHeadUserId,
          readId
        );
        const finDirPermUpdatePm = parentFolderItem.roleAssignments.add(
          finDirUserId,
          readId
        );
        const snrDirPeoplePermUpdatePm = parentFolderItem.roleAssignments.add(
          snrDirOfPeopleId,
          readId
        );

        return await Promise.all([
          primBMPm,
          secBMPm,
          primFAPm,
          secFAPm,
          ceoPermUpdatePm,
          corpHeadPermUpdatePm,
          commHeadPermUpdatePm,
          finDirPermUpdatePm,
          snrDirPeoplePermUpdatePm
        ]);
      } else {
        return await Promise.all([primBMPm, secBMPm, primFAPm, secFAPm]);
        return true;
      }
    } catch (error) {
      console.log(
        "an error occured while updating permissions on the attachments folder for a submitted status: " +
        error
      );
    }
  }

  /**
   * @description creates sub folders in the parent procurement request folder
   * @param formGuid
   */
  private static provisionSubFolders(parentFolderName: string): Promise<any[]> {
    const subFolderPms: Promise<any>[] = this.subFolders.map(folder => {
      try {
        return sp.web.folders.add(
          "Shared Documents/" + parentFolderName + "/" + folder
        );
      } catch (error) {
        console.log(
          "an error occured while creating the " + folder + " folder: " + error
        );
      }
    });
    return Promise.all(subFolderPms);
  }

  /**
   * @description adds the file to a folder
   * @param serverRelativeUrl
   * @param formGuid
   * @param attachments array of attachments
   */
  private static addAttachments(
    serverRelativeUrl: string,
    formGuid: string,
    attachments: any[],
    formData: any
  ): Promise<any[]> {
    let folderUrl: string;
    if (formData["ProcurementStatus"] && formData["ProcurementStatus"] == "Draft") {
      folderUrl = serverRelativeUrl + "/Shared Documents/" + formGuid + "/Attachments";
    } else {
      // ProcurementStatus must be either Approved, Returned or Submitted
      folderUrl = serverRelativeUrl + "/Shared Documents/" + formData["ProcReference"] + "/Attachments";
    }
    let attachmentsPms: Promise<any>[] = attachments.map(attachment => {
      try {
        return sp.web
          .getFolderByServerRelativeUrl(folderUrl)
          .files.add(attachment.name, attachment.file, true);
      } catch (error) {
        console.log(
          "an error occured while adding an attachment " +
          attachment.name +
          "error: " +
          error
        );
      }
    });
    return Promise.all(attachmentsPms);
  }

  private static addAttachmentsForNewSubmittedForm(
    serverRelativeUrl: string,
    procRef: string,
    attachments: any[],
  ): Promise<any[]> {
    const folderUrl = serverRelativeUrl + "/Shared Documents/" + procRef + "/Attachments";
    let attachmentsPms: Promise<any>[] = attachments.map(attachment => {
      try {
        return sp.web
          .getFolderByServerRelativeUrl(folderUrl)
          .files.add(attachment.name, attachment.file, true);
      } catch (error) {
        console.log("an error occured while adding an attachment " + attachment.name + "error: " + error);
      }
    });
    return Promise.all(attachmentsPms);
  }

  /**
   *
   * @description investigating how to add a 'link to a document' in a modern document library
   * @param serverRelativeUrl
   * @param formData
   * @deprecated - unable to create link in modern document library, users will have to do it manually
   */
  /* public static async addFormLinkToParentFolder(serverRelativeUrl?:string, formData?:IFormData) {
    // const folderUrl = serverRelativeUrl + '/Shared Documents/';
    // let contentTypes = await sp.web.lists.getByTitle("TestLinks").contentTypes.get();
    let contentTypes = await sp.web.lists.getByTitle("TestLinks")
      .items
      .add(
        {
          Title: "Test 1",
          ContentTypeId: "0x01030058FD86C279252341AB303852303E4DAF"
        })
    // console.log('content types: ' + JSON.stringify(contentTypes));
    console.log('content types: ' + JSON.stringify(contentTypes));
    
  } */

  /**
   * @description provides a list of existing attachments to display on the form in edit mode
   * @fires from FileUpload.tsx componentDidMount()
   * @param context
   * @param formGuid
   * @param formData - used to check the Procurement Status
   * @returns an array of files
   */
  public static async getAttachments(
    serverRelativeUrl: string,
    formGuid: string,
    formData: any
  ): Promise<any[]> {
    // get list of all files in folder
    let folderUrl: string;
    if (
      formData["ProcurementStatus"] &&
      formData["ProcurementStatus"] == "Draft"
    ) {
      folderUrl =
        serverRelativeUrl + "/Shared Documents/" + formGuid + "/Attachments";
    } else if (formData["ProcurementStatus"]) {
      // ProcurementStatus must be either Approved, Returned or Submitted
      folderUrl =
        serverRelativeUrl +
        "/Shared Documents/" +
        formData["ProcReference"] +
        "/Attachments";
    }
    return sp.web
      .getFolderByServerRelativeUrl(folderUrl)
      .expand("Files")
      .get()
      .then(response => {
        return response.Files;
      });
  }
}
