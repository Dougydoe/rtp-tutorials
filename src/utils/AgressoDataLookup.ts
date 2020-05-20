import { sp } from '@pnp/sp';
import { IDropDownOption } from '../interfaces/IRtpFormState';
import { Environment, EnvironmentType } from '@microsoft/sp-core-library';
import { AgressoData } from '../mock/MockAgressoData';
import { IAgressoData } from '../interfaces';



/**
 * @description details methods used to auto-populate the budget check fields with data extracted from Agresso
 */
export class AgressoDataLookup {
  
    /**
     * Description: sychronously calls a number of functions to retrieve data to autopopulate the form
     * @fires when state.formData.SubProjectCode being updated when ContractValue is not empty
     * @param subProjectCode : Entered into the form 
     * @param contractValue : Entered into the form
     * @returns data to display in fields 
     */
    public static async getValuesForBudgetCheckAndFinancialApproverFields(subProjectCode:string, totalContractValue:string):Promise<IAgressoData> {   
      if (Environment.type === EnvironmentType.Local) {
        return AgressoData;
      } else if (Environment.type === EnvironmentType.SharePoint) {
        let subProjectCodeDescription:string;
        let costCentre:string;
        let projectCode:string; 
        let approvalLevel:string;
        let primaryFinancialApprover:string;
        let purchasingDirectorate:string;
        let possibleSecondaryFinancialApprovers:IDropDownOption[];
        let primaryBusinessManager:string;
    
        try {
          let costCentreAndProjectCodePm:Promise<{costCentre:string, projectCode:string, subProjectCodeDescription:string}> = this.getCostCentreAndProjectCodeAndDescription(subProjectCode);
          let approvalLevelPm:Promise<string> = this.getApprovalLevel(totalContractValue);      
      
          await Promise.all([costCentreAndProjectCodePm, approvalLevelPm]).then(response => {
            costCentre = response[0].costCentre;
            projectCode = response[0].projectCode;
            subProjectCodeDescription = response[0].subProjectCodeDescription;
            approvalLevel = response[1];
          });
      
          let primaryFinancialApproverPm:Promise<string> = this.getPrimaryFinancialApprover(projectCode, approvalLevel);    
          let purchasingDirectoratePm:Promise<string> = this.getPurchasingDirectorate(costCentre);
      
          await Promise.all([primaryFinancialApproverPm, purchasingDirectoratePm]).then(response => {
            primaryFinancialApprover = response[0];
            purchasingDirectorate = response[1];
          });
      
          let possibleSecondaryFinancialApproversPm = this.getPossibleSecondaryFinancialApproversInApprovalLevel(approvalLevel, purchasingDirectorate);
          let primaryBusinessManagerPm = this.getPrimaryBusinessManager(purchasingDirectorate);
      
          await Promise.all([possibleSecondaryFinancialApproversPm, primaryBusinessManagerPm]).then(response => {
            possibleSecondaryFinancialApprovers = response[0];
            primaryBusinessManager = response[1];
          });
  
          if (subProjectCodeDescription && costCentre && projectCode && purchasingDirectorate && primaryFinancialApprover && primaryBusinessManager && possibleSecondaryFinancialApprovers) {
            return {
              subProjectCodeDescription: subProjectCodeDescription,
              costCentre: costCentre,
              projectCode: projectCode,
              purchasingDirectorate: purchasingDirectorate,
              primaryFinancialApprover: primaryFinancialApprover,
              possibleSecondaryFinancialApprovers: possibleSecondaryFinancialApprovers,
              primaryBusinessManager: primaryBusinessManager
            };
          } else {
            throw new Error("Not all of the lookup data was retrieved from Agresso");
          }
      
        } catch (e) {
          console.log('Error retrieving Agresso data: ' + e);
          return null;
        }
      }
    }
  
    /**
     * @description: queries the CurrentSubProjects SP list
     * @fires from Utility.getValuesForBudgetCheckAndFinancialApproverFields()
     * @param subProjectCode : as entered on the form by the Purchaser
     * @returns the project code and the cost centre associated with the sub project code 
     * ? use getPaged() to return more than 100 results
     */
    private static getCostCentreAndProjectCodeAndDescription(subProjectCode:string): Promise<{costCentre:string, projectCode:string, subProjectCodeDescription:string}> {
      
      // filter out based on date to column to ensure that old sub project codes aren't used mistakenly by the Purchaser 
      if (subProjectCode) {
        return sp.web.lists
        .getByTitle('CurrentSubProjects')
        .items
        .select('ProjectCode', 'CostCentre', 'Description')
        .filter("Title eq " + "'" + subProjectCode + "'")
        .get()
        .then(response => {      
          if (response[0]) {
            return {
              costCentre: response[0]['CostCentre'],
              projectCode: response[0]['ProjectCode'],
              subProjectCodeDescription: response[0]['Description']
            };
          } else {
            throw new Error(`No cost centre and project code found in CurrentSubProjects list for ${subProjectCode}.`);
          }
        });    
      } else {
        throw new Error(`Unable to retrieve cost centre and project code as sub project code ${subProjectCode} was not provided.`);
      }
    }
  
    /**
     * Description: queries the DistributionRules SP list to retrieve the approval level which is then used to get the appropriate primary financial approver
     * Invoked: Utility.getValuesForBudgetCheckAndFinancialApproverFields()
     * @param contractValue : the contract value entered on the form
     * Returns: the approval level 
     */
    public static async getApprovalLevel(contractValue:string):Promise<string> {
      
      if (contractValue) {
        const value = parseFloat(contractValue);
        const rules = await sp.web.lists
        .getByTitle('DistributionRules')
        .items
        .select('Receipient', 'Title', 'UpperRange', 'LowerRange')    
        .get();

        console.log(JSON.stringify(rules));

        if (rules) {
          for(let i=0; i < rules.length; i++) {
            let lowerRange = parseFloat(rules[i].LowerRange);
            let upperRange = parseFloat(rules[i].UpperRange);
            if (rules[i].Title == "between" && value >= lowerRange && value <= upperRange) {
              return rules[i].Receipient;                
            } else if (rules[i].Title == "greater than" && value >= lowerRange) {
              return rules[i].Receipient;          
            }
          }    
        } else {
          throw new Error(`No rules found in DistributionRules list.`);
        }
      } else {
        throw new Error(`Unable to retrieve approval level as contract value ${contractValue} was not provided.`);
      }
      
    }
  
    /**
     * Description: queries the CurrentProjectLevelxApprovers SP list (dependent on the approval level) to auto-populate the primary financial approver field on the form
     * Invoked: Utility.getValuesForBudgetCheckAndFinancialApproverFields()
     * @param projectCode : the project code related to the sub project code entered into the form 
     * @param approvalLevel : the necessary approval level for the contract value
     * Returns: the first financial approver's email address 
     */
    private static getPrimaryFinancialApprover(projectCode:string, approvalLevel:string):Promise<string> {
      let projectAppproversList:string;
      switch (approvalLevel) {
        case 'Approval Level 2' : 
        projectAppproversList = 'CurrentProjectLevel2Approvers';
          break;
        case 'Approval Level 3' :
        projectAppproversList = "CurrentProjectLevel3Approvers";
          break;
        case 'Approval Level 4' :
        projectAppproversList = "CurrentProjectLevel4Approvers";
          break;
        case 'Approval Level 5' :
        projectAppproversList = "CurrentProjectLevel5Approvers";
          break;
      }   

      if (projectCode && projectAppproversList) {
        return sp.web.lists
        .getByTitle(projectAppproversList) 
        .items      
        .filter("Title eq " + "'" + projectCode + "'")
        .get()
        .then(response => {
          // check if response is empty
          if (response[0]) {
            // check if first result contains Approver
            if(response[0].Approver) {
              return response[0].Approver;
            } else if (response[1].Approver) {
              return response[1].Approver;
            } else if (response[2].Approver) {
              return response[2].Approver;
            } else {
              throw new Error('No primary financial approver found from the first 3 responses from the list: ' + projectAppproversList);
            }               
          } else {
            throw new Error(`No primary financial approver found in ${projectAppproversList} list for project code ${projectCode}.`);
          }
        });
      } else {
        throw new Error(`Unable to retrieve primary financial approver as either project code ${projectCode}, or approval level ${approvalLevel} were not provided.`);
      }
    }
  
    /**
     * Description: queries the OrganisationalStructure SP list to retrieve the purchasing directorate to display on the form 
     * Invoked: Utility.getValuesForBudgetCheckAndFinancialApproverFields()
     * @param costCentre : responsible for the purchase
     * Returns: the purchasing directorate
     */
    private static getPurchasingDirectorate(costCentre:string):Promise<string> {
      if (costCentre) {
        return sp.web.lists
        .getByTitle('OrganisationalStructure')
        .items
        .select('Directorate')
        .filter("Title eq " + "'" + costCentre + "'")
        .get()
        .then(response => {
          if (response[0]) {
            return response[0].Directorate;
          } else {
            throw new Error(`No directorate found in OrganisationalStructure list for cost centre ${costCentre}.`);
          }
        });
      } else {
        throw new Error(`Unable to retrieve purchasing directorate as cost centre ${costCentre} was not provided.`);
      }
    }
  
    /**
     * Description: queries a list (dependent on approval level) to provide selecteable options on a dropdown field on the form allowing the purchaser to select a secondary financial approver to send the request to 
     * Invoked: Utility.getValuesForBudgetCheckAndFinancialApproverFields()
     * @param approvalLevel : needed for the procurement request's contract value
     * @param directorate : purchasing directorate
     * Returns: list of financial approvers that can be selected from a dropdown field
     */
    public static async getPossibleSecondaryFinancialApproversInApprovalLevel(approvalLevel:string, directorate:string): Promise<IDropDownOption[]> {
  
      let approversList:string;
      switch (approvalLevel) {
        case 'Approval Level 2' : 
        approversList = 'Level2Approvers';
          break;
        case 'Approval Level 3' :
        approversList = "Level3Approvers";
          break;
        case 'Approval Level 4' :
        approversList = "Level4Approvers";
          break;
        case 'Approval Level 5' :
        approversList = "Level5Approvers";
          break;
      }
      
      if (approversList && directorate) {
        const approvers = await sp.web.lists
        .getByTitle(approversList)
        .items
        .select('Approver')
        .filter("Title eq " + "'" + directorate + "'")
        .get();
        
        if (approvers[0]) {
          let possibleSecondaryFinancialApprovers:IDropDownOption[] = [];
          approvers.map(item => {
            if (item.Approver) {
              possibleSecondaryFinancialApprovers.push({key: item.Approver.toLowerCase(), text: item.Approver.toLowerCase()});
            }
          });
          return possibleSecondaryFinancialApprovers;
        } else {
          throw new Error(`No possible secondary financial approvers found in ${approversList} list for directorate ${directorate}.`);
        }
            
      } else {
        throw new Error(`Unable to retrieve possible secondary financial approvers as either the directorate ${directorate} or approval level ${approvalLevel} were not provided.`);
      }
    }
  
    /**   
     * Description: populates the primary business manager field on the form 
     * Invoked: Utility.getValuesForBudgetCheckAndFinancialApproverFields()
     * @param directorate : purchasing directorate
     * Returns: the business manager for the directorate
     */
    private static getPrimaryBusinessManager(directorate:string):Promise<string> {
      
      if (directorate) {
        return sp.web.lists
        .getByTitle('BusinessManagerDirectorates')
        .items
        .select('BusinessManager')
        .filter("Title eq " + "'" + directorate + "'")
        .get()
        .then(response => {
          if (response[0]) {
            return response[0].BusinessManager;
          } else {
            throw new Error(`No business manager found in BusinessManagerDirectorates list for directorate ${directorate}.`);
          }
        });
      } else {
        throw new Error(`Unable to retrieve primary business manager as the directorate ${directorate} was not provided.`);
      }
      
    }
  }