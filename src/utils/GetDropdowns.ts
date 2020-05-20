import { sp } from '@pnp/sp';
import { Environment, EnvironmentType } from '@microsoft/sp-core-library';



export interface IGetOptions {
  readonly choiceFieldsLookup:string[];
  readonly choiceFieldsStatic:string[];
  readonly comboFields:string[];
}

interface IDropdownOption {
  key:string | number;
  text:string;
}

interface IAllDropdownOptions {
  ProcurementRoute: IDropdownOption[];
  StaType:IDropdownOption[];
  ApprovedScope:IDropdownOption[];
  RouteToMarket:IDropdownOption[];
  ProcurementType:IDropdownOption[];
  DpoOptions:IDropdownOption[];
  IsmOptions:IDropdownOption[];
  BusinessCaseRequirement:IDropdownOption[];
  UnableProcureCompetitively:IDropdownOption[];
  SuggestedTendererStatus:IDropdownOption[];
  BusinessManagerDirectorates:IDropdownOption[];
  SecondaryBusinessManager:IDropdownOption[];
}

const MOCKALLDROPDOWNOPTIONS: IAllDropdownOptions = {
  ProcurementRoute: [
    {key:'Request to Procure (RTP)', text:"Request to Procure (RTP)"}, 
    {key:'Single Tender Action (STA)', text:"Single Tender Action (STA)"}, 
    {key:'Fast Track Delegated Procurement (FTDP)', text:"Fast Track Delegated Procurement (FTDP)"}
  ],
  StaType: [{key:1, text:"foo"}, {key:2, text:"bar"}],
  ApprovedScope: [{key:1, text:"foo"}, {key:2, text:"bar"}],
  RouteToMarket: [{key:1, text:"foo"}, {key:2, text:"bar"}],
  ProcurementType: [{key:1, text:"foo"}, {key:2, text:"bar"}],
  DpoOptions: [{key:1, text:"foo"}, {key:2, text:"bar"}],
  IsmOptions: [{key:1, text:"foo"}, {key:2, text:"bar"}],
  BusinessCaseRequirement: [{key:1, text:"foo"}, {key:2, text:"bar"}],
  UnableProcureCompetitively: [{key:1, text:"foo"}, {key:2, text:"bar"}],
  SuggestedTendererStatus: [{key:1, text:"foo"}, {key:2, text:"bar"}],
  BusinessManagerDirectorates: [{key:1, text:"foo"}, {key:2, text:"bar"}],
  SecondaryBusinessManager: [{key:1, text:"foo"}, {key:2, text:"bar"}],
} 


export class GetOptions implements IGetOptions {
    
  /**
   * @description these fields do not include fields rendered as combo fields, they maintain referential integrity
   */
  readonly choiceFieldsLookup = [
    'ProcurementRoute',
    'StaType', 
    'ApprovedScope', 
    'RouteToMarket', 
    'ProcurementType', 
    'DpoOptions', 
    'IsmOptions', 
    'BusinessCaseRequirement', 
    'UnableProcureCompetitively', 
    'SuggestedTendererStatus',
    'BusinessManagerDirectorates',
  ];

  /**
   * @description these fields do not maintain referential integrity
   */
  readonly choiceFieldsStatic = [
    'SecondaryBusinessManager',
  ];

  /**
   * @description these fields are rendered as combo fields, they maintain referential integrity
   */
  readonly comboFields = [
    'SubProjectCode'
  ];

  /**
   * @description async calls to get promises for all choice fields 
   */
  private async getAllChoiceFieldsLookup() {  
    
    let allOptions = {};
    const fields = this.choiceFieldsLookup;
    for (let i=0; i<fields.length; i++) {
      let optionsPm = this.getChoiceFieldLookupOptions(fields[i]);
      allOptions[fields[i]] = optionsPm; 
    }

    
  }

  /**
   * 
   * @param field the choice field to get the lookup options for 
   * @description retrieves the lookup options for the given field from the SP list
   */
  private getChoiceFieldLookupOptions(field:string):Promise<IDropdownOption[]> {
    let list:string;
    let name:string = "Title";
    let id:string = "Id";
    let option:IDropdownOption;
    let options:IDropdownOption[] = new Array;

    switch(field) {
      case 'SecondaryBusinessManager' :
        list = 'AllBusinessManagers';
        id = 'Title';
        break;
      case 'ProcurementRoute' :
        list = 'procurement route';
        break;
      case 'ProcurementType' :
        list = 'procurement type';
        break;
      case 'ApprovedScope' :
        list = 'approved scope';
        break;
      case 'RouteToMarket' :
        list = 'route to market';
        break;
      case 'StaType' :
        list = 'sta type';
        break;  
      case 'DpoOptions' :
        list = 'dpo options';
        break;       
      case 'IsmOptions' :
        list = 'ism options';
        break;
      case 'BusinessCaseRequirement' :
        list = 'business case requirement';
        break;  
      case 'SuggestedTendererStatus' :
        list = 'suggested tenderer status';
        break;  
      case 'BusinessManagerDirectorates' :
        list = 'BusinessManagerDirectorates';
        break;  
      case 'UnableProcureCompetitively' :
        list = 'unable procure competitively';
        name = "Option";
        id = "Option";
        break;
    }

    return sp.web.lists
      .getByTitle(list)
      .items
      .select(name, "ID")
      .filter("Archive eq 'No'")
      .get()
      .then((response:any[]) => {
        response.map(item => {
          option = {
            key: item[id],
            text: item[name]
          };
          options.push(option);
        });
        return options;
      });
  }

  /**
   * 
   * @param field the choice field to get the static options for 
   * @description retrieves the static options for the given field from the SP list
   */
  private getChoiceFieldStaticOptions(field:string):Promise<IDropdownOption[]> {
    let list:string;
    let option:IDropdownOption;
    let options:IDropdownOption[] = new Array;
    let name:string = "Title";
    let id:string= "Id";

    switch(field) {
      case 'SecondaryBusinessManager' :
        list = 'AllBusinessManagers';
        break;
    }

    return sp.web.lists
      .getByTitle(list)
      .items
      .select("Title")
      .get()
      .then((response:any[]) => {
        response.map(item => {
          option = {
            key: item[id].toLowerCase(),
            text: item[name].toLowerCase()
          };
          options.push(option);
        });
        return options;
      });
  }

}