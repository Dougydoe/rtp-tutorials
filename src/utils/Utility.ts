import { sp, ClientPeoplePickerQueryParameters } from '@pnp/sp';
import { IPersona, PersonaPresence, ITag } from 'office-ui-fabric-react';
import { Environment, EnvironmentType } from '@microsoft/sp-core-library';
import { People, MockVariationOptions } from '../mock/MockDropdowns';
import { ITaskAssigneeQuery } from '../mock/MockStaTaskAssignees';
import { IDropDownOption, IVariationOptions } from '../interfaces/IRtpFormState';

/**
 * @description misc methods 
 */
export class Utility {

  /**
   * @fires by this.convertFormData() and FileApi.updateParentFolderPermissionsOnSubmit()
   * @description check if a user has been selected and gets the user id
   * @param logonName eg "i:0#.f|membership|user
   * @returns the sharepoint user profile id of the user
  */  
  public static getSharePointUserId = (logonName:string):Promise<number> => {     
    return sp.web.ensureUser(logonName)
    .then(response => {        
      return response.data.Id;
    });    
  }

  private static _doesTextStartWith(text: string, filterText: string): boolean {
    return text.toLowerCase().indexOf(filterText.toLowerCase()) === 0;
  }

  /**
   * @fires by Form.converListItemToFormData()
   * @param displayName 
   * @param absoluteUrl 
   * @returns an array of IPersona to display a user in the a people picker field on the form
   */
  public static async getPersonaForPeoplePickerField(displayName:string, absoluteUrl:string):Promise<IPersona[]> {
    if (Environment.type === EnvironmentType.Local) {
      // ? How do I filter an array where the text property of each element contains displayName?
      return People.filter(item => this._doesTextStartWith(item['text'] as string, displayName));
    } else if (Environment.type === EnvironmentType.SharePoint) {
      console.log(`5a - getPersonaForPeoplePickerField called for ${displayName}`);
      let queryParams:ClientPeoplePickerQueryParameters = {
       QueryString: displayName,
       MaximumEntitySuggestions: 10,
       AllowEmailAddresses: true,
       AllowOnlyEmailAddresses: false,
       PrincipalType: 1,
       PrincipalSource: 15,
       SharePointGroupID: 0,
     };
     let tenantBaseUrl:string = absoluteUrl.substring(0, absoluteUrl.indexOf("sharepoint.com") + 14);
     let imageBaseUrl =  tenantBaseUrl + "/_layouts/15/userphoto.aspx?size=S&accountname=";  
 
     return await sp.profiles.clientPeoplePickerSearchUser(queryParams)
     .then(response => {
       console.log(`5b - response from clientPeoplePickerSearchUser ${JSON.stringify(response)}`);
       return [{
         text: response[0].DisplayText,
         secondaryText: response[0].EntityData.Title,
         tertiaryText: response[0].EntityData.Department,                           
         optionalText: response[0].Key,
         imageInitials: Utility.getInitials(response[0].DisplayText), 
         presence: PersonaPresence.none,
         imageUrl: imageBaseUrl + response[0].Description,
       }];
     });
    }
  }
  /**
   * @description used to retrieve the email to be displayed in the business manager and financial approver fields 
   * @param displayName 
   * 
   */
  public static async getUserEmailFromDisplayName(displayName:string):Promise<string> {

    let queryParams:ClientPeoplePickerQueryParameters = {
      QueryString: displayName,
      MaximumEntitySuggestions: 10,
      AllowEmailAddresses: true,
      AllowOnlyEmailAddresses: false,
      PrincipalType: 1,
      PrincipalSource: 15,
      SharePointGroupID: 0,
    };

    return sp.profiles.clientPeoplePickerSearchUser(queryParams)
    .then(response => {
      return response[0]['EntityData']['Email'].toLowerCase();
    });
                
  }

  // this is used by the SPPeoplePicker and PersonInfo controls
  public static getInitials(fullname: string): string {

    if (!fullname) {
        return (null);
    }
    
    var parts = fullname.split(' ');
    
    var initials = "";
    parts.forEach(p => {
        if (p.length > 0)
        {
            initials = initials.concat(p.substring(0, 1).toUpperCase());
        }
    });

    return (initials);
  }

  // get lookup data for different types of choice/dropdown fields
  public static getTagOptions(field): PromiseLike<ITag[]> {

    let lookupItems:ITag[] = new Array; 
    let lookupItem:ITag;
    let list:string; 
    let name:string = "Title";
    let id:string= "Id";

    // console.log('field to retrieve lookup data for: ' + field);
  
    switch(field) {
      case 'PurchaseCategory' : 
        list = "purchase category";      
        break;
    }
  
    return sp.web.lists
        .getByTitle(list)
        .items
        .select(name, "ID")
        .filter("Archive eq 'No'")
        .get()
        .then((response: any[]) => {        
            response.forEach(item => {
                lookupItem = {
                    key: item[id],
                    name: item[name],
                };
                lookupItems.push(lookupItem);
            });          
            return lookupItems;
        });                          
        
  }
  
  // returns the SharePoint groups that the current user is a member of
  public static getUserGroupMembership():Promise<any[]> {
    return sp.web.currentUser.groups.select('Title').get().then(response => {
      return response;
    });

    /* const url = context.pageContext.web.absoluteUrl + '/_api/web/currentuser/groups?$select=Title';
    return context.spHttpClient.get(url, SPHttpClient.configurations.v1)
    .then((response:SPHttpClientResponse) => {
      return response.json();
    })
    .then(response => {
      // console.log('getUserGroupMembership response: ' + JSON.stringify(response.value));
      return response.value;
    }); */
  }

  /**
   * @description checks if field exists in this.state.errors array
   * @param field 
   * @returns true if field exists, and the index of the field in the errors array   
   */
  public static checkIfErrorExists = (field:string, errors:string[]):{exists:boolean, index:number} => {
    let exists:boolean = false;
    let index:number;
    for (let i = 0; i<errors.length; i++) {
      if (errors[i] == field) {
        exists = true;
        index = i;
        break;
      }
    }
    return {
      exists: exists,
      index: index
    };
  }

  /**
   * @description gets all the PROC references from the Procurement list
   * @returns
   */
  public static async getPROCReferences():Promise<IVariationOptions> {
    if (Environment.type === EnvironmentType.Local) {
      return MockVariationOptions;
    } else if (Environment.type === EnvironmentType.SharePoint) {
      try {
        let option:IDropDownOption;
        let options:IDropDownOption[] = new Array;
  
        await sp.web.lists
          .getByTitle('Procurement')
          .items
          .select('ProcReference', 'ID')
          .filter("ProcReference ne null")
          .getAll()
          .then(data => {
            data.map(item => {
              option = {
                key: item['ProcReference']+ ',' + ' Id:' + item['ID'],
                text: item['ProcReference']
              };
              options.push(option);
            });
          });
        return {
          Variation: options
        };
      } catch (error) {
        console.log(`unable to retrieve Proc references: ${error}`);
      }
    }
  }

  /**
   * @description get STA task assignees 
   * @returns array of user ids for the task assignees
   */
  public static getTaskAssigneesUserIds():Promise<ITaskAssigneeQuery[]> {
    try {
      return sp.web.lists
      .getByTitle('TaskAssignees')
      .items
      .select('TaskAssignee/ID', 'Title')
      .expand('TaskAssignee')
      .getAll();
    } catch (error) {
      console.log(`unable to retrieve sta task assignees when setting list item permissions on submit: ${error}`);
    }
  }

  /**
   * 
   * @param procReference 
   * @deprecated
   */
  public static getFormIdFromProcReference(procReference:string):number {
    // const procRefArr = procReference.split('-', 2);
    const procRefArr = procReference.split(':');
    const formId = parseInt(procRefArr[1]);
    return formId;
  }

}




