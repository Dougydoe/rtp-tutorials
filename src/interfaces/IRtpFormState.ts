export interface IRtpFormState {
    formData: any;  
    errors: any;  
    userInfo:IUserInfo;
    dropDownOptions?:IDropDownOptions; 
    isSubmitting:boolean;
    loading:boolean;
}

export interface IUserInfo {
    userRole:string;
    isCommercialMember:boolean;
}

export interface IDropDownOption {
    key:string | number;
    text:string;
}

export interface IDropdownFieldsOptions {
    ProcurementRoute: IDropDownOption[];
    StaType:IDropDownOption[];
    ApprovedScope:IDropDownOption[];
    RouteToMarket:IDropDownOption[];
    ProcurementType:IDropDownOption[];
    SecondaryBusinessManager:IDropDownOption[];
    DpoOptions:IDropDownOption[];
    IsmOptions:IDropDownOption[];
    BusinessCaseRequirement:IDropDownOption[];
    UnableProcureCompetitively:IDropDownOption[];
    SuggestedTendererStatus:IDropDownOption[];
    BusinessManagerDirectorates:IDropDownOption[];
}

export interface ISubProjectCodeOptions {
    SubProjectCode:IDropDownOption[];
}

export interface IVariationOptions {
    Variation:IDropDownOption[];
}

export interface IComboFieldsOptions extends ISubProjectCodeOptions, IVariationOptions {
}

export interface IDropDownOptions extends IDropdownFieldsOptions, IComboFieldsOptions {
    SecondaryFinancialApprover?:IDropDownOption[];
} 


