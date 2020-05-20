import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import { IPropertyPaneConfiguration, PropertyPaneTextField } from "@microsoft/sp-property-pane";
import * as strings from 'RtpFormWebPartStrings';
import RtpForm from './components/RtpForm';
import { IRtpFormProps, IFormContext } from '../../interfaces/index';
import { getGUID } from '@pnp/common';
import "@pnp/polyfill-ie11";
import { sp } from "@pnp/sp";
import { initializeIcons } from "office-ui-fabric-react";

export interface IRtpFormWebPartProps {
  description: string;  
}

require("./filepicker.css");
require("./dropzone.css");

export default class RtpFormWebPart extends BaseClientSideWebPart<IRtpFormProps> {

  public render(): void {    
    const element: React.ReactElement<IRtpFormProps> = React.createElement(
      RtpForm,
      {
        description: this.properties.description,
        pageContext: this.context.pageContext,        
        context:this.context,
        listName: "Procurement",
        formContext: this.getFormContext(),       
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  public async onInit():Promise<void> {

    const params = new URLSearchParams(document.location.search.substring(1)); 
    const version = params.get("v");
    const wplocation = window.location.href;

    if (!version) {
      // * page should re-load
      const form_id = params.get("form_id");
      if (!form_id) {
        // * there is a new form 
        const web = params.get("web");
        if (!web) {
          // * there are no existing parameters
          window.location.href = wplocation + '?v=1';
        } else {
          // * there is an existing paramter
          window.location.href = wplocation + '&v=1';
        }
      } else {
        // * there are existing params 
        window.location.href = wplocation + '&v=1';
      }

    }

    initializeIcons();

    return super.onInit().then(_ => {
      sp.setup({
        spfxContext: this.context
      });
    });
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })                
              ]
            }
          ]
        }
      ]
    };
  }

  private getFormContext = (): IFormContext => {

    const params = new URLSearchParams(document.location.search.substring(1));     
    const form_id_string: string = params.get("form_id");    
    const form_id: number = parseInt(form_id_string);
    const mode: string = (form_id_string ? "edit" : "new");   
    let form_guid: string;

    if (mode == "new") {
      form_guid = getGUID();      
    } else if (mode == "edit") {      
      form_guid = params.get("form_guid");            
    }
    
    return {
      mode: mode,
      formId: form_id,
      formGuid: form_guid,
    };

  }

}
